const DB_NAME = "sheetwork-local-v1";
const STORE = "workspace";

function database(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function transaction<T>(mode: IDBTransactionMode, operation: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await database();
  try {
    return await new Promise<T>((resolve, reject) => {
      const tx = db.transaction(STORE, mode);
      const request = operation(tx.objectStore(STORE));
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
      tx.onerror = () => reject(tx.error);
    });
  } finally {
    db.close();
  }
}

export const loadLocal = <T>(key: string) => transaction<T | undefined>("readonly", (store) => store.get(key));
export const saveLocal = (key: string, value: unknown) => transaction<IDBValidKey>("readwrite", (store) => store.put(value, key));
export const removeLocal = (key: string) => transaction<undefined>("readwrite", (store) => store.delete(key));
