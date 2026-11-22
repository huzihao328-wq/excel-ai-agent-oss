/** 附加文件的 IndexedDB 持久化：任务窗格关掉后仍能恢复。 */

export interface StoredAttachment {
  name: string;
  loadedAt: number;
  order: number;
  kind: "csv" | "xlsx";
  bytes: ArrayBuffer;
}

const DB_NAME = "excel-ai-agent";
const DB_VERSION = 1;
const STORE = "attachments";

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === "undefined") {
      reject(new Error("IndexedDB 不可用"));
      return;
    }
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) {
        db.createObjectStore(STORE, { keyPath: "name" });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error("打开附件库失败"));
  });
}

function reqToPromise<T>(req: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error("附件库操作失败"));
  });
}

export async function listStoredAttachments(): Promise<StoredAttachment[]> {
  try {
    const db = await openDb();
    const tx = db.transaction(STORE, "readonly");
    const items = await reqToPromise(tx.objectStore(STORE).getAll() as IDBRequest<StoredAttachment[]>);
    db.close();
    return (items || []).sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  } catch {
    return [];
  }
}

export async function putStoredAttachment(item: StoredAttachment): Promise<void> {
  const db = await openDb();
  const tx = db.transaction(STORE, "readwrite");
  tx.objectStore(STORE).put(item);
  await new Promise<void>((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("保存附件失败"));
  });
  db.close();
}

export async function deleteStoredAttachment(name: string): Promise<void> {
  try {
    const db = await openDb();
    const tx = db.transaction(STORE, "readwrite");
    tx.objectStore(STORE).delete(name);
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error("删除附件失败"));
    });
    db.close();
  } catch {
    /* 内存删除仍有效 */
  }
}

export async function getStoredAttachment(name: string): Promise<StoredAttachment | null> {
  try {
    const db = await openDb();
    const tx = db.transaction(STORE, "readonly");
    const item = await reqToPromise(tx.objectStore(STORE).get(name) as IDBRequest<StoredAttachment | undefined>);
    db.close();
    return item || null;
  } catch {
    return null;
  }
}
