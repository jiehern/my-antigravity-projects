/**
 * db.js - Localised Database Engine for Class Logger
 * Powered by IndexedDB with seamless fallback to localStorage
 */

const DB_NAME = 'ClassLoggerDB';
const DB_VERSION = 1;
const STORE_CLASSES = 'classes';
const STORE_CATEGORIES = 'categories';

export const DEFAULT_CATEGORIES = [
  { id: 'cat_coach_allen', name: 'Coach Allen', color: '#ff9f0a', icon: '🏊‍♂️', isDefault: true },
  { id: 'cat_aerosplash', name: 'Aerosplash', color: '#0a84ff', icon: '🌊', isDefault: true },
  { id: 'cat_charles', name: 'Charles', color: '#30d158', icon: '🥊', isDefault: true }
];

class LocalDatabase {
  constructor() {
    this.db = null;
    this.isIndexedDBAvailable = typeof window !== 'undefined' && 'indexedDB' in window;
    this.initPromise = null;
  }

  async init() {
    if (this.initPromise) return this.initPromise;

    this.initPromise = new Promise((resolve) => {
      if (!this.isIndexedDBAvailable) {
        console.warn('IndexedDB unavailable, using LocalStorage');
        this.initLocalStorageDefaults();
        return resolve();
      }

      try {
        const request = indexedDB.open(DB_NAME, DB_VERSION);

        request.onupgradeneeded = (event) => {
          const db = event.target.result;
          if (!db.objectStoreNames.contains(STORE_CLASSES)) {
            const classStore = db.createObjectStore(STORE_CLASSES, { keyPath: 'id' });
            classStore.createIndex('timestamp', 'timestamp', { unique: false });
            classStore.createIndex('category', 'category', { unique: false });
            classStore.createIndex('date', 'date', { unique: false });
          }
          if (!db.objectStoreNames.contains(STORE_CATEGORIES)) {
            const catStore = db.createObjectStore(STORE_CATEGORIES, { keyPath: 'id' });
            DEFAULT_CATEGORIES.forEach(cat => catStore.add(cat));
            try { localStorage.setItem('class_logger_categories_seeded', 'true'); } catch (e) {}
          }
        };

        request.onsuccess = (event) => {
          this.db = event.target.result;
          resolve();
        };

        request.onerror = (event) => {
          console.error('IndexedDB open error, falling back to LocalStorage:', event.target.error);
          this.isIndexedDBAvailable = false;
          this.initLocalStorageDefaults();
          resolve();
        };
      } catch (err) {
        console.error('IndexedDB exception, falling back to LocalStorage:', err);
        this.isIndexedDBAvailable = false;
        this.initLocalStorageDefaults();
        resolve();
      }
    });

    return this.initPromise;
  }

  initLocalStorageDefaults() {
    const existingCats = localStorage.getItem('class_logger_categories');
    if (existingCats === null) {
      localStorage.setItem('class_logger_categories', JSON.stringify(DEFAULT_CATEGORIES));
      localStorage.setItem('class_logger_categories_seeded', 'true');
    }
    const existingClasses = localStorage.getItem('class_logger_classes');
    if (existingClasses === null) {
      localStorage.setItem('class_logger_classes', JSON.stringify([]));
    }
  }

  // --- CLASSES CRUD ---

  async getAllClasses() {
    await this.init();
    if (!this.isIndexedDBAvailable || !this.db) {
      const data = localStorage.getItem('class_logger_classes');
      const list = data ? JSON.parse(data) : [];
      return list.sort((a, b) => b.timestamp - a.timestamp);
    }

    return new Promise((resolve) => {
      try {
        const tx = this.db.transaction(STORE_CLASSES, 'readonly');
        const store = tx.objectStore(STORE_CLASSES);
        const request = store.getAll();

        request.onsuccess = () => {
          const results = request.result || [];
          results.sort((a, b) => b.timestamp - a.timestamp);
          resolve(results);
        };
        request.onerror = () => {
          console.error('Error fetching classes:', request.error);
          resolve([]);
        };
      } catch (err) {
        console.error('Transaction error:', err);
        resolve([]);
      }
    });
  }

  async addClass(classItem) {
    await this.init();
    const id = classItem.id || 'cls_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7);
    const item = {
      ...classItem,
      id,
      createdAt: classItem.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    if (!this.isIndexedDBAvailable || !this.db) {
      const list = await this.getAllClasses();
      list.unshift(item);
      localStorage.setItem('class_logger_classes', JSON.stringify(list));
      return item;
    }

    return new Promise((resolve, reject) => {
      try {
        const tx = this.db.transaction(STORE_CLASSES, 'readwrite');
        const store = tx.objectStore(STORE_CLASSES);
        const request = store.add(item);

        request.onsuccess = () => resolve(item);
        request.onerror = () => reject(request.error);
      } catch (err) {
        reject(err);
      }
    });
  }

  async updateClass(id, updatedFields) {
    await this.init();
    if (!this.isIndexedDBAvailable || !this.db) {
      const list = await this.getAllClasses();
      const index = list.findIndex(c => c.id === id);
      if (index === -1) throw new Error('Class not found');
      const updated = {
        ...list[index],
        ...updatedFields,
        updatedAt: new Date().toISOString()
      };
      list[index] = updated;
      localStorage.setItem('class_logger_classes', JSON.stringify(list));
      return updated;
    }

    return new Promise((resolve, reject) => {
      try {
        const tx = this.db.transaction(STORE_CLASSES, 'readwrite');
        const store = tx.objectStore(STORE_CLASSES);
        const getReq = store.get(id);

        getReq.onsuccess = () => {
          if (!getReq.result) {
            return reject(new Error('Class not found'));
          }
          const updated = {
            ...getReq.result,
            ...updatedFields,
            updatedAt: new Date().toISOString()
          };
          const putReq = store.put(updated);
          putReq.onsuccess = () => resolve(updated);
          putReq.onerror = () => reject(putReq.error);
        };
        getReq.onerror = () => reject(getReq.error);
      } catch (err) {
        reject(err);
      }
    });
  }

  async deleteClass(id) {
    await this.init();
    if (!this.isIndexedDBAvailable || !this.db) {
      let list = await this.getAllClasses();
      list = list.filter(c => c.id !== id);
      localStorage.setItem('class_logger_classes', JSON.stringify(list));
      return true;
    }

    return new Promise((resolve, reject) => {
      try {
        const tx = this.db.transaction(STORE_CLASSES, 'readwrite');
        const store = tx.objectStore(STORE_CLASSES);
        const request = store.delete(id);

        request.onsuccess = () => resolve(true);
        request.onerror = () => reject(request.error);
      } catch (err) {
        reject(err);
      }
    });
  }

  async clearAllClasses() {
    await this.init();
    if (!this.isIndexedDBAvailable || !this.db) {
      localStorage.setItem('class_logger_classes', JSON.stringify([]));
      return true;
    }

    return new Promise((resolve, reject) => {
      try {
        const tx = this.db.transaction(STORE_CLASSES, 'readwrite');
        const store = tx.objectStore(STORE_CLASSES);
        const request = store.clear();
        request.onsuccess = () => resolve(true);
        request.onerror = () => reject(request.error);
      } catch (err) {
        reject(err);
      }
    });
  }

  // --- CATEGORIES CRUD ---

  async getAllCategories() {
    await this.init();
    if (!this.isIndexedDBAvailable || !this.db) {
      const data = localStorage.getItem('class_logger_categories');
      return data !== null ? JSON.parse(data) : DEFAULT_CATEGORIES;
    }

    return new Promise((resolve) => {
      try {
        const tx = this.db.transaction(STORE_CATEGORIES, 'readonly');
        const store = tx.objectStore(STORE_CATEGORIES);
        const request = store.getAll();

        request.onsuccess = () => {
          const results = request.result || [];
          const seeded = localStorage.getItem('class_logger_categories_seeded');
          if (results.length === 0 && !seeded) {
            try { localStorage.setItem('class_logger_categories_seeded', 'true'); } catch (e) {}
            this.seedDefaultCategories().then(() => resolve(DEFAULT_CATEGORIES));
          } else {
            try { localStorage.setItem('class_logger_categories_seeded', 'true'); } catch (e) {}
            resolve(results);
          }
        };
        request.onerror = () => resolve([]);
      } catch (err) {
        resolve([]);
      }
    });
  }

  async seedDefaultCategories() {
    if (!this.isIndexedDBAvailable || !this.db) return;
    try {
      const tx = this.db.transaction(STORE_CATEGORIES, 'readwrite');
      const store = tx.objectStore(STORE_CATEGORIES);
      DEFAULT_CATEGORIES.forEach(cat => store.put(cat));
    } catch (e) {}
  }

  async addCategory(category) {
    await this.init();
    const id = category.id || 'cat_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6);
    const item = { ...category, id };

    if (!this.isIndexedDBAvailable || !this.db) {
      const cats = await this.getAllCategories();
      cats.push(item);
      localStorage.setItem('class_logger_categories', JSON.stringify(cats));
      return item;
    }

    return new Promise((resolve, reject) => {
      try {
        const tx = this.db.transaction(STORE_CATEGORIES, 'readwrite');
        const store = tx.objectStore(STORE_CATEGORIES);
        const request = store.put(item);

        request.onsuccess = () => resolve(item);
        request.onerror = () => reject(request.error);
      } catch (err) {
        reject(err);
      }
    });
  }

  async deleteCategory(id) {
    await this.init();
    try { localStorage.setItem('class_logger_categories_seeded', 'true'); } catch (e) {}
    if (!this.isIndexedDBAvailable || !this.db) {
      let cats = await this.getAllCategories();
      cats = cats.filter(c => c.id !== id);
      localStorage.setItem('class_logger_categories', JSON.stringify(cats));
      return true;
    }

    return new Promise((resolve, reject) => {
      try {
        const tx = this.db.transaction(STORE_CATEGORIES, 'readwrite');
        const store = tx.objectStore(STORE_CATEGORIES);
        const request = store.delete(id);

        request.onsuccess = () => resolve(true);
        request.onerror = () => reject(request.error);
      } catch (err) {
        reject(err);
      }
    });
  }

  // --- EXPORT & IMPORT ---

  async getFilteredClasses(filter = {}) {
    let list = await this.getAllClasses();
    const { category, startDate, endDate } = filter;

    if (category && category !== 'all') {
      const catLower = category.toLowerCase();
      list = list.filter(c => c.category && c.category.toLowerCase() === catLower);
    }

    if (startDate) {
      list = list.filter(c => c.date >= startDate);
    }

    if (endDate) {
      list = list.filter(c => c.date <= endDate);
    }

    return list;
  }

  async exportJSON(filter = {}) {
    const classes = await this.getFilteredClasses(filter);
    const categories = await this.getAllCategories();
    return JSON.stringify({
      version: 1,
      appName: 'ClassLogger',
      exportedAt: new Date().toISOString(),
      filter: {
        category: filter.category || 'all',
        startDate: filter.startDate || null,
        endDate: filter.endDate || null
      },
      totalCount: classes.length,
      classes,
      categories
    }, null, 2);
  }

  async exportCSV(filter = {}) {
    const classes = await this.getFilteredClasses(filter);
    const headers = ['Date', 'Day', 'Time', 'Category', 'Duration (mins)', 'Duration (hrs)', 'Notes / Remarks', 'Session ID'];
    const rows = classes.map(c => {
      const d = new Date(c.date + 'T00:00:00');
      const dayOfWeek = isNaN(d.getTime()) ? '' : d.toLocaleDateString('en-US', { weekday: 'short' });
      const hours = c.duration ? (c.duration / 60).toFixed(2) : '1.00';
      return [
        `"${c.date}"`,
        `"${dayOfWeek}"`,
        `"${c.displayTime || c.time}"`,
        `"${(c.category || '').replace(/"/g, '""')}"`,
        c.duration || 60,
        hours,
        `"${(c.note || '').replace(/"/g, '""')}"`,
        `"${c.id}"`
      ];
    });

    return [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
  }

  async importJSON(jsonStr) {
    try {
      const data = typeof jsonStr === 'string' ? JSON.parse(jsonStr) : jsonStr;
      if (!data || !Array.isArray(data.classes)) {
        throw new Error('Invalid JSON format: missing classes list');
      }

      for (const item of data.classes) {
        if (!item.id) {
          item.id = 'cls_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7);
        }
        await this.addClass(item);
      }

      if (Array.isArray(data.categories)) {
        for (const cat of data.categories) {
          await this.addCategory(cat);
        }
      }

      return { success: true, count: data.classes.length };
    } catch (e) {
      console.error('Import failed:', e);
      throw e;
    }
  }

  async seedSampleData() {
    const existing = await this.getAllClasses();
    if (existing.length > 0) return false;

    const now = new Date();
    const samples = [
      {
        date: new Date(now.getTime() - 2 * 3600 * 1000).toISOString().split('T')[0],
        time: '09:00',
        displayTime: '09:00 AM',
        timestamp: now.getTime() - 2 * 3600 * 1000,
        category: 'Coach Allen',
        duration: 60,
        note: 'High-intensity swim interval drills. Completed 24 laps, improved sprint pacing by 1.4s.'
      },
      {
        date: new Date(now.getTime() - 26 * 3600 * 1000).toISOString().split('T')[0],
        time: '17:30',
        displayTime: '05:30 PM',
        timestamp: now.getTime() - 26 * 3600 * 1000,
        category: 'Aerosplash',
        duration: 45,
        note: 'Aerobic water resistance & core stability workout. Great energy in group session.'
      },
      {
        date: new Date(now.getTime() - 50 * 3600 * 1000).toISOString().split('T')[0],
        time: '18:00',
        displayTime: '06:00 PM',
        timestamp: now.getTime() - 50 * 3600 * 1000,
        category: 'Charles',
        duration: 75,
        note: 'Boxing footwork combinations, heavy bag rounds, and defensive duck-and-weave drills.'
      }
    ];

    for (const sample of samples) {
      await this.addClass(sample);
    }
    return true;
  }
}

export const db = new LocalDatabase();
