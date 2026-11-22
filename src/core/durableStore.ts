/*
 * 跨版本用户档案。localStorage 会随 WebView 缓存或分区变化丢失；
 * 这里同时写到 OfficeRuntime.storage（跟插件 ID 走）和本机
 * %LOCALAPPDATA%\ExcelAIAgent\userdata\（跟 Windows 用户走）。
 * 新版本启动时若本地为空，自动从这两处继承。
 */
import { AppSettings } from "./types";

export const MEMORY_STORAGE_KEY = "excel-ai-agent-memory";
export const SKILLS_STORAGE_KEY = "excel-ai-agent-skills";
export const SETTINGS_STORAGE_KEY = "excel-ai-agent-settings";
export const TEMPLATES_STORAGE_KEY = "excel-ai-agent-prompt-templates";

const OFFICE_KEY = "excel-ai-agent-profile";
const DISK_URL = "/api/user-data";
const SCHEMA_VERSION = 1;
const HYDRATE_MS = 1500;
const PERSIST_MS = 5000;

export interface UserProfile {
  schemaVersion: number;
  updatedAt: number;
  skills?: unknown[];
  memories?: unknown[];
  settings?: Partial<AppSettings> | null;
  templates?: unknown[];
}

let hydrated = false;
let hydratePromise: Promise<void> | null = null;
let persistTimer: number | null = null;
let persistChain: Promise<void> = Promise.resolve();
let officePersistChain: Promise<void> = Promise.resolve();

function officeStorage(): {
  getItem(key: string): Promise<string | null>;
  setItem(key: string, value: string): Promise<void>;
} | null {
  const rt = (globalThis as unknown as { OfficeRuntime?: { storage?: unknown } }).OfficeRuntime;
  const storage = rt && (rt as { storage?: { getItem?: unknown; setItem?: unknown } }).storage;
  if (!storage || typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
  return storage as { getItem(key: string): Promise<string | null>; setItem(key: string, value: string): Promise<void> };
}

function withTimeout<T>(promise: Promise<T>, ms: number, fallback: T): Promise<T> {
  return new Promise((resolve) => {
    const timer = window.setTimeout(() => resolve(fallback), ms);
    promise.then(
      (value) => {
        window.clearTimeout(timer);
        resolve(value);
      },
      () => {
        window.clearTimeout(timer);
        resolve(fallback);
      }
    );
  });
}

function parseProfile(raw: unknown): UserProfile | null {
  if (typeof raw !== "string" || !raw) return null;
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return null;
    return {
      schemaVersion: typeof parsed.schemaVersion === "number" ? parsed.schemaVersion : SCHEMA_VERSION,
      updatedAt: typeof parsed.updatedAt === "number" ? parsed.updatedAt : 0,
      skills: Array.isArray(parsed.skills) ? parsed.skills : undefined,
      memories: Array.isArray(parsed.memories) ? parsed.memories : undefined,
      settings: parsed.settings && typeof parsed.settings === "object" && !Array.isArray(parsed.settings) ? parsed.settings : undefined,
      templates: Array.isArray(parsed.templates) ? parsed.templates : undefined,
    };
  } catch {
    return null;
  }
}

function readLocalJson<T>(key: string): T | null {
  try {
    const raw = localStorage.getItem(key);
    if (raw == null) return null;
    return JSON.parse(raw) as T;
  } catch {
    return null;
  }
}

function snapshotLocal(): UserProfile {
  return {
    schemaVersion: SCHEMA_VERSION,
    updatedAt: Date.now(),
    skills: readLocalJson<unknown[]>(SKILLS_STORAGE_KEY) || [],
    memories: readLocalJson<unknown[]>(MEMORY_STORAGE_KEY) || [],
    settings: readLocalJson<Partial<AppSettings>>(SETTINGS_STORAGE_KEY),
    templates: readLocalJson<unknown[]>(TEMPLATES_STORAGE_KEY) || [],
  };
}

function writeLocal(profile: UserProfile): void {
  try {
    if (profile.skills) localStorage.setItem(SKILLS_STORAGE_KEY, JSON.stringify(profile.skills));
    if (profile.memories) localStorage.setItem(MEMORY_STORAGE_KEY, JSON.stringify(profile.memories));
    if (profile.settings) localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(profile.settings));
    if (profile.templates) localStorage.setItem(TEMPLATES_STORAGE_KEY, JSON.stringify(profile.templates));
  } catch {
    /* quota */
  }
}

/**
 * 本机磁盘档案不保存 API Key：它由常驻 HTTP 服务读写，文件还可能被同步盘、备份软件带走。
 * Key 只留在 localStorage 与 OfficeRuntime.storage（都在 WebView 自己的存储里）。
 */
export function withoutSecrets(profile: UserProfile): UserProfile {
  if (!profile.settings || !("apiKey" in profile.settings)) return profile;
  const { apiKey: _omit, ...settings } = profile.settings;
  return { ...profile, settings };
}

function inherit(local: UserProfile, remote: UserProfile | null): UserProfile {
  if (!remote) return local;
  // An explicit [] is a saved deletion, not a missing profile.
  const missingArray = (key: string) => !Array.isArray(readLocalJson(key));
  const localHasSettings = readLocalJson(SETTINGS_STORAGE_KEY) != null;
  return {
    schemaVersion: SCHEMA_VERSION,
    updatedAt: Date.now(),
    skills: missingArray(SKILLS_STORAGE_KEY) && remote.skills ? remote.skills : local.skills,
    memories: missingArray(MEMORY_STORAGE_KEY) && remote.memories ? remote.memories : local.memories,
    settings: !localHasSettings && remote.settings ? remote.settings : local.settings,
    templates: missingArray(TEMPLATES_STORAGE_KEY) && remote.templates ? remote.templates : local.templates,
  };
}

async function readOffice(): Promise<UserProfile | null> {
  const storage = officeStorage();
  if (!storage) return null;
  const raw = await storage.getItem(OFFICE_KEY);
  return parseProfile(raw);
}

async function writeOffice(profile: UserProfile): Promise<void> {
  const storage = officeStorage();
  if (!storage) return;
  await storage.setItem(OFFICE_KEY, JSON.stringify(profile));
}

async function readDisk(): Promise<UserProfile | null> {
  const controller = new AbortController();
  const timer = window.setTimeout(() => controller.abort(), HYDRATE_MS);
  try {
    const resp = await fetch(DISK_URL, { method: "GET", cache: "no-store", signal: controller.signal });
    if (resp.status === 404) return null;
    if (!resp.ok) throw new Error(String(resp.status));
    return parseProfile(JSON.stringify(await resp.json()));
  } finally {
    window.clearTimeout(timer);
  }
}

async function writeDisk(profile: UserProfile): Promise<void> {
  const controller = new AbortController();
  const timer = window.setTimeout(() => controller.abort(), PERSIST_MS);
  try {
    const resp = await fetch(DISK_URL, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(profile),
      signal: controller.signal,
    });
    if (!resp.ok) throw new Error(String(resp.status));
  } finally {
    window.clearTimeout(timer);
  }
}

async function writeDurable(profile: UserProfile): Promise<void> {
  // Office storage cannot be aborted. Keep actual writes ordered even after our wait times out.
  officePersistChain = officePersistChain.catch(() => undefined).then(() => writeOffice(profile));
  await Promise.all([
    withTimeout(officePersistChain, PERSIST_MS, undefined),
    withTimeout(writeDisk(withoutSecrets(profile)), PERSIST_MS, undefined),
  ]);
}

function newer(a: UserProfile | null, b: UserProfile | null): UserProfile | null {
  if (!a) return b;
  if (!b) return a;
  return (b.updatedAt || 0) > (a.updatedAt || 0) ? b : a;
}

/** 启动时调用一次：本地为空则从 Office / 本机档案继承，并把当前数据写回备份。 */
export function hydrateUserData(): Promise<void> {
  hydratePromise ||= (async () => {
    const [office, disk] = await Promise.all([
      withTimeout(readOffice().catch(() => null), HYDRATE_MS, null),
      withTimeout(readDisk().catch(() => null), HYDRATE_MS, null),
    ]);
    let remote = newer(office, disk);
    // 磁盘档案较新但不含 Key 时，从 Office 档案补回，避免换版本后要重新填 Key。
    if (remote && remote === disk && remote.settings && !remote.settings.apiKey && office?.settings?.apiKey) {
      remote = { ...remote, settings: { ...remote.settings, apiKey: office.settings.apiKey } };
    }
    const local = snapshotLocal();
    const hasLocal = [SKILLS_STORAGE_KEY, MEMORY_STORAGE_KEY, SETTINGS_STORAGE_KEY, TEMPLATES_STORAGE_KEY]
      .some((key) => readLocalJson(key) != null);
    if (remote || hasLocal) writeLocal(inherit(local, remote));
    hydrated = true;
    // Backup availability must never block rendering the task pane.
    if (remote || hasLocal) void flushUserData();
  })();
  return hydratePromise;
}

export function flushUserData(): Promise<void> {
  if (!hydrated || ![SKILLS_STORAGE_KEY, MEMORY_STORAGE_KEY, SETTINGS_STORAGE_KEY, TEMPLATES_STORAGE_KEY]
    .some((key) => readLocalJson(key) != null)) return Promise.resolve();
  if (persistTimer != null) {
    window.clearTimeout(persistTimer);
    persistTimer = null;
  }
  const profile = snapshotLocal();
  persistChain = persistChain.then(() => writeDurable(profile)).catch(() => undefined);
  return persistChain;
}

/** 记忆/技能/设置变更后调用；短时间内多次写入会合并成一次。 */
export function schedulePersistUserData(): void {
  if (!hydrated) return;
  if (persistTimer != null) window.clearTimeout(persistTimer);
  persistTimer = window.setTimeout(() => {
    persistTimer = null;
    void flushUserData();
  }, 250);
}

if (typeof window !== "undefined") {
  window.addEventListener("pagehide", () => {
    void flushUserData();
  });
}
