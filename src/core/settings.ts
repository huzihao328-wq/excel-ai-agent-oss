import { SETTINGS_STORAGE_KEY, schedulePersistUserData } from "./durableStore";
import { AppSettings } from "./types";

export const DEFAULT_SETTINGS: AppSettings = {
  baseUrl: "https://api.deepseek.com/v1",
  apiKey: "",
  model: "deepseek-chat",
  temperature: 0.2,
  maxToolRounds: 20,
  maxOutputTokens: 4096,
  confirmDestructive: true,
  sendSelectionContext: true,
  askDirection: true,
  useMemory: true,
  autoLearnSkills: true,
  autoCheckpoint: true,
  contextWindowTokens: 0,
};

function asString(value: unknown, fallback: string): string {
  return typeof value === "string" ? value : fallback;
}

function asBool(value: unknown, fallback: boolean): boolean {
  return typeof value === "boolean" ? value : fallback;
}

function asNumber(value: unknown, fallback: number, min: number, max: number): number {
  const n = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(max, Math.max(min, n));
}

export function loadSettings(): AppSettings {
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (!raw) return { ...DEFAULT_SETTINGS };
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return { ...DEFAULT_SETTINGS };
    return {
      baseUrl: asString(parsed.baseUrl, DEFAULT_SETTINGS.baseUrl) || DEFAULT_SETTINGS.baseUrl,
      apiKey: asString(parsed.apiKey, ""),
      model: asString(parsed.model, DEFAULT_SETTINGS.model),
      temperature: asNumber(parsed.temperature, DEFAULT_SETTINGS.temperature, 0, 2),
      maxToolRounds: Math.round(asNumber(parsed.maxToolRounds, DEFAULT_SETTINGS.maxToolRounds, 3, 50)),
      maxOutputTokens: Math.round(asNumber(parsed.maxOutputTokens, DEFAULT_SETTINGS.maxOutputTokens, 0, 128_000)),
      confirmDestructive: asBool(parsed.confirmDestructive, DEFAULT_SETTINGS.confirmDestructive),
      sendSelectionContext: asBool(parsed.sendSelectionContext, DEFAULT_SETTINGS.sendSelectionContext),
      askDirection: asBool(parsed.askDirection, DEFAULT_SETTINGS.askDirection),
      useMemory: asBool(parsed.useMemory, DEFAULT_SETTINGS.useMemory),
      autoLearnSkills: asBool(parsed.autoLearnSkills, DEFAULT_SETTINGS.autoLearnSkills),
      autoCheckpoint: asBool(parsed.autoCheckpoint, DEFAULT_SETTINGS.autoCheckpoint),
      contextWindowTokens: Math.round(asNumber(parsed.contextWindowTokens, 0, 0, 16_000_000)),
    };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export function saveSettings(s: AppSettings): void {
  localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(s));
  schedulePersistUserData();
}

export function endpointOrigin(baseUrl: string): string {
  try {
    return new URL(baseUrl).origin.toLowerCase();
  } catch {
    return "";
  }
}

/** Loopback model servers such as Ollama and LM Studio commonly do not require a key. */
export function isLoopbackEndpoint(baseUrl: string): boolean {
  try {
    const host = new URL(baseUrl).hostname.toLowerCase();
    return host === "localhost" || host === "::1" || host === "[::1]" || /^127(?:\.\d{1,3}){3}$/.test(host);
  } catch {
    return false;
  }
}

export function modelEndpointError(baseUrl: string): string | null {
  let url: URL;
  try {
    url = new URL(baseUrl);
  } catch {
    return "API 地址不是合法 URL。";
  }
  if (url.username || url.password) return "API 地址不能包含用户名或密码。";
  if (url.protocol === "https:") return null;
  if (url.protocol === "http:" && isLoopbackEndpoint(baseUrl)) return null;
  return "远程模型 API 必须使用 HTTPS；只有 localhost / 127.0.0.1 / ::1 可使用 HTTP。";
}

/** 常用服务商预设，方便一键切换（均为 OpenAI 兼容接口） */
export const PROVIDER_PRESETS: { label: string; baseUrl: string; model: string }[] = [
  { label: "OpenCode Go", baseUrl: "https://opencode.ai/zen/go/v1", model: "deepseek-v4-flash" },
  { label: "DeepSeek", baseUrl: "https://api.deepseek.com/v1", model: "deepseek-chat" },
  { label: "Kimi", baseUrl: "https://api.moonshot.cn/v1", model: "moonshot-v1-32k" },
  { label: "通义千问", baseUrl: "https://dashscope.aliyuncs.com/compatible-mode/v1", model: "qwen-plus" },
  { label: "智谱 GLM", baseUrl: "https://open.bigmodel.cn/api/paas/v4", model: "glm-4-air" },
  { label: "豆包·方舟", baseUrl: "https://ark.cn-beijing.volces.com/api/v3", model: "doubao-pro-32k" },
  { label: "硅基流动", baseUrl: "https://api.siliconflow.cn/v1", model: "Qwen/Qwen2.5-72B-Instruct" },
  { label: "OpenAI", baseUrl: "https://api.openai.com/v1", model: "gpt-4o" },
  { label: "OpenRouter", baseUrl: "https://openrouter.ai/api/v1", model: "anthropic/claude-3.5-sonnet" },
  { label: "Groq", baseUrl: "https://api.groq.com/openai/v1", model: "llama-3.3-70b-versatile" },
  { label: "xAI Grok", baseUrl: "https://api.x.ai/v1", model: "grok-2-latest" },
  { label: "Gemini", baseUrl: "https://generativelanguage.googleapis.com/v1beta/openai", model: "gemini-1.5-pro" },
  { label: "Ollama 本地", baseUrl: "http://localhost:11434/v1", model: "qwen2.5:14b" },
  { label: "LM Studio 本地", baseUrl: "http://localhost:1234/v1", model: "" },
];
