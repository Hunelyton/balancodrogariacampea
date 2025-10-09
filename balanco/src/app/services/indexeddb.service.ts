export class IndexedDBService {
  private dbName: string;
  private storeName: string;
  private isFallback = false;
  private fallbackStore = new Map<string, any>();

  constructor(dbName: string, storeName: string) {
    this.dbName = dbName;
    this.storeName = storeName;
  }

  private getIndexedDB(): IDBFactory | undefined {
    if (typeof globalThis !== 'undefined' && 'indexedDB' in globalThis) {
      return (globalThis as unknown as { indexedDB?: IDBFactory }).indexedDB;
    }
    return undefined;
  }

  private openDB(): Promise<IDBDatabase | null> {
    return new Promise((resolve, reject) => {
      const indexedDBRef = this.getIndexedDB();
      if (!indexedDBRef) {
        if (!this.isFallback) {
          this.isFallback = true;
          console.warn(
            'IndexedDB not available; using in-memory storage. Dados não serão persistidos entre sessões.'
          );
        }
        resolve(null);
        return;
      }

      let request: IDBOpenDBRequest;
      try {
        request = indexedDBRef.open(this.dbName);
      } catch (err) {
        if (!this.isFallback) {
          this.isFallback = true;
          console.warn(
            'IndexedDB open() falhou; usando armazenamento em memória. Dados não serão persistidos entre sessões.',
            err
          );
        }
        resolve(null);
        return;
      }

      request.onupgradeneeded = (event: IDBVersionChangeEvent) => {
        const db = (event.target as IDBOpenDBRequest).result;
        if (!db.objectStoreNames.contains(this.storeName)) {
          db.createObjectStore(this.storeName, { keyPath: 'id' });
        }
      };

      request.onsuccess = () => resolve(request.result);
      request.onerror = () => {
        if (!this.isFallback) {
          this.isFallback = true;
          console.warn(
            'IndexedDB request failed; usando armazenamento em memória. Dados não serão persistidos entre sessões.',
            request.error
          );
        }
        resolve(null);
      };
    });
  }

  async setItem(key: string, value: any): Promise<void> {
    const db = await this.openDB();
    if (!db) {
      this.fallbackStore.set(key, value);
      return;
    }
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(this.storeName, 'readwrite');
      const store = transaction.objectStore(this.storeName);
      store.put({ id: key, value });

      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error);
    });
  }

  async getItem(key: string): Promise<any> {
    const db = await this.openDB();
    if (!db) {
      return this.fallbackStore.get(key);
    }
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(this.storeName, 'readonly');
      const store = transaction.objectStore(this.storeName);
      const request = store.get(key);

      request.onsuccess = () => resolve(request.result?.value);
      request.onerror = () => reject(request.error);
    });
  }

  async removeItem(key: string): Promise<void> {
    const db = await this.openDB();
    if (!db) {
      this.fallbackStore.delete(key);
      return;
    }
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(this.storeName, 'readwrite');
      const store = transaction.objectStore(this.storeName);
      store.delete(key);

      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error);
    });
  }

  async clear(): Promise<void> {
    const db = await this.openDB();
    if (!db) {
      this.fallbackStore.clear();
      return;
    }
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(this.storeName, 'readwrite');
      const store = transaction.objectStore(this.storeName);
      store.clear();

      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error);
    });
  }
}
