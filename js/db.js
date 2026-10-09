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

  async upsertClass(classItem) {
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
      const existingIdx = list.findIndex(c => c.id === id);
      if (existingIdx !== -1) {
        list[existingIdx] = item;
      } else {
        list.unshift(item);
      }
      localStorage.setItem('class_logger_classes', JSON.stringify(list));
      return item;
    }

    return new Promise((resolve, reject) => {
      try {
        const tx = this.db.transaction(STORE_CLASSES, 'readwrite');
        const store = tx.objectStore(STORE_CLASSES);
        const request = store.put(item);

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
    const headers = ['Date', 'Day', 'Time', 'Class Type', 'Category', 'Duration (mins)', 'Duration (hrs)', 'Notes / Remarks', 'Session ID'];
    const rows = classes.map(c => {
      const d = new Date(c.date + 'T00:00:00');
      const dayOfWeek = isNaN(d.getTime()) ? '' : d.toLocaleDateString('en-US', { weekday: 'short' });
      const hours = c.duration ? (c.duration / 60).toFixed(2) : '1.00';
      const classType = c.classType || (
        (c.note && c.note.toLowerCase().includes('baby')) || c.duration === 30 ? 'Baby Class' :
        (c.note && (c.note.toLowerCase().includes('pre comp') || c.note.toLowerCase().includes('squad'))) || c.duration === 90 ? 'Squad (Pre Comp)' :
        'LTS'
      );
      return [
        `"${c.date}"`,
        `"${dayOfWeek}"`,
        `"${c.displayTime || c.time}"`,
        `"${classType}"`,
        `"${(c.category || 'AeroSplash').replace(/"/g, '""')}"`,
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
      if (!data) {
        throw new Error('Invalid JSON format: empty payload');
      }

      // Extract array of classes from common structures:
      // 1. data.classes (standard ClassLogger export format)
      // 2. data directly is an Array
      // 3. data.sessions or data.data
      let classesList = null;
      if (Array.isArray(data.classes)) {
        classesList = data.classes;
      } else if (Array.isArray(data)) {
        classesList = data;
      } else if (Array.isArray(data.sessions)) {
        classesList = data.sessions;
      } else if (Array.isArray(data.data)) {
        classesList = data.data;
      } else {
        throw new Error('Invalid JSON format: missing "classes" array');
      }

      let importedCount = 0;
      for (const raw of classesList) {
        if (!raw || typeof raw !== 'object') continue;

        // Auto-normalize fields
        const date = String(raw.date || '').trim() || new Date().toISOString().split('T')[0];
        const time = String(raw.time || '12:00').trim();
        let displayTime = raw.displayTime;
        if (!displayTime) {
          const parts = time.split(':');
          let h = parseInt(parts[0], 10) || 12;
          const m = parseInt(parts[1], 10) || 0;
          const period = h >= 12 ? 'PM' : 'AM';
          let h12 = h % 12;
          if (h12 === 0) h12 = 12;
          displayTime = `${String(h12).padStart(2, '0')}:${String(m).padStart(2, '0')} ${period}`;
        }

        let timestamp = raw.timestamp;
        if (!timestamp) {
          const [y, m, d] = date.split('-').map(Number);
          const [hh, mm] = time.split(':').map(Number);
          timestamp = new Date(y, (m || 1) - 1, d || 1, hh || 12, mm || 0).getTime();
        }

        const category = 'AeroSplash';
        const duration = Number(raw.duration) || 60;
        const note = String(raw.note || '').trim();
        const noteLower = note.toLowerCase();
        const rawType = String(raw.classType || '').trim();
        let classType = 'LTS';
        if (rawType) {
          if (rawType.toLowerCase().includes('baby')) classType = 'Baby Class';
          else if (rawType.toLowerCase().includes('squad') || rawType.toLowerCase().includes('pre comp') || rawType.toLowerCase().includes('precomp')) classType = 'Squad (Pre Comp)';
          else if (rawType.toLowerCase().includes('lts')) classType = 'LTS';
          else classType = rawType;
        } else {
          if (noteLower.includes('baby') || duration === 30) classType = 'Baby Class';
          else if (noteLower.includes('squad') || noteLower.includes('pre comp') || noteLower.includes('precomp') || duration === 90) classType = 'Squad (Pre Comp)';
          else classType = 'LTS';
        }

        const normalizedItem = {
          ...raw,
          id: raw.id || `cls_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
          date,
          time,
          displayTime,
          timestamp,
          category,
          classType,
          duration,
          note,
          createdAt: raw.createdAt || new Date().toISOString(),
          updatedAt: new Date().toISOString()
        };

        await this.upsertClass(normalizedItem);
        importedCount++;
      }

      return { success: true, count: importedCount };
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
        category: 'AeroSplash',
        classType: 'LTS',
        duration: 50,
        note: 'Swim interval drills and stroke mechanics.'
      },
      {
        date: new Date(now.getTime() - 26 * 3600 * 1000).toISOString().split('T')[0],
        time: '16:00',
        displayTime: '04:00 PM',
        timestamp: now.getTime() - 26 * 3600 * 1000,
        category: 'AeroSplash',
        classType: 'Baby Class',
        duration: 30,
        note: 'Water familiarity and bubble blowing exercises.'
      },
      {
        date: new Date(now.getTime() - 50 * 3600 * 1000).toISOString().split('T')[0],
        time: '17:30',
        displayTime: '05:30 PM',
        timestamp: now.getTime() - 50 * 3600 * 1000,
        category: 'AeroSplash',
        classType: 'Squad (Pre Comp)',
        duration: 90,
        note: 'Endurance sets, flip turns, and timed 100m pacing.'
      }
    ];

    for (const sample of samples) {
      await this.addClass(sample);
    }
    return true;
  }
}

export const db = new LocalDatabase();
