/**
 * 任务检查点的整本文件快照，存在任务窗格自己的 IndexedDB 里，而不是用户工作簿里。
 * 以前的做法是把每张表复制成 VeryHidden 备份表：用户删掉的数据会随文件保存、发给别人，
 * 文件体积也翻倍。现在工作簿里只留一个检查点 id（document settings）。
 *
 * IndexedDB 按插件来源共享，多本工作簿的快照都在这里，用 id 区分；过期快照自动清理。
 */

export interface CheckpointFormulasRecord {
  sheet: string;
  range: string;
  formulas: string[][];
}

export interface CheckpointFileRecord {
  id: string;
  createdAt: number;
  bytes: ArrayBuffer;
  formulas: CheckpointFormulasRecord[];
}

// 独立的库名：与附件库分开，互不影响各自的版本升级
const DB_NAME = "excel-ai-agent-checkpoints";
const DB_VERSION = 1;
const STORE = "files";
/** 超过这个时间没被当前检查点引用的快照视为孤儿 */
const MAX_AGE_MS = 14 * 24 * 3600 * 1000;
/** 所有工作簿合计最多保留的快照数 */
const MAX_RECORDS = 10;

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === "undefined") {
      reject(new Error("IndexedDB 不可用"));
      return;
    }
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: "id" });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error("打开检查点库失败"));
  });
}

function reqToPromise<T>(req: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error("检查点库操作失败"));
  });
}

function txDone(tx: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("检查点库写入失败"));
    tx.onabort = () => reject(tx.error || new Error("检查点库写入被中止（可能空间不足）"));
  });
}

export async function putCheckpointFile(record: CheckpointFileRecord): Promise<void> {
  const db = await openDb();
  try {
    const tx = db.transaction(STORE, "readwrite");
    tx.objectStore(STORE).put(record);
    await txDone(tx);
  } finally {
    db.close();
  }
}

export async function getCheckpointFile(id: string): Promise<CheckpointFileRecord | null> {
  const db = await openDb();
  try {
    const tx = db.transaction(STORE, "readonly");
    const item = await reqToPromise(tx.objectStore(STORE).get(id) as IDBRequest<CheckpointFileRecord | undefined>);
    return item || null;
  } finally {
    db.close();
  }
}

export async function hasCheckpointFile(id: string): Promise<boolean> {
  try {
    const db = await openDb();
    try {
      const tx = db.transaction(STORE, "readonly");
      const count = await reqToPromise(tx.objectStore(STORE).count(id));
      return count > 0;
    } finally {
      db.close();
    }
  } catch {
    return false;
  }
}

export async function deleteCheckpointFile(id: string): Promise<void> {
  try {
    const db = await openDb();
    try {
      const tx = db.transaction(STORE, "readwrite");
      tx.objectStore(STORE).delete(id);
      await txDone(tx);
    } finally {
      db.close();
    }
  } catch {
    /* 删不掉的快照会被过期清理带走 */
  }
}

/** 清理过期或超量的快照；keepId（当前工作簿正在用的）永不删除。 */
export async function pruneCheckpointFiles(keepId?: string, now = Date.now()): Promise<void> {
  try {
    const db = await openDb();
    try {
      const tx = db.transaction(STORE, "readwrite");
      const store = tx.objectStore(STORE);
      const req = store.openCursor();
      const meta: { id: string; createdAt: number }[] = [];
      await new Promise<void>((resolve, reject) => {
        req.onsuccess = () => {
          const cursor = req.result;
          if (!cursor) return resolve();
          const value = cursor.value as CheckpointFileRecord;
          meta.push({ id: value.id, createdAt: value.createdAt || 0 });
          cursor.continue();
        };
        req.onerror = () => reject(req.error);
      });
      const others = meta.filter((m) => m.id !== keepId).sort((a, b) => b.createdAt - a.createdAt);
      const budget = MAX_RECORDS - (keepId && meta.some((m) => m.id === keepId) ? 1 : 0);
      others.forEach((m, i) => {
        if (i >= budget || now - m.createdAt > MAX_AGE_MS) store.delete(m.id);
      });
      await txDone(tx);
    } finally {
      db.close();
    }
  } catch {
    /* 清理失败不影响任务 */
  }
}
