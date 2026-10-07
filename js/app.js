/**
 * app.js - Main Application Orchestrator for ClassLogger
 * Integrates Apple Rotary Dial Picker, Localized IndexedDB Storage, Categories, Audio & Analytics
 */

import { db, DEFAULT_CATEGORIES } from './db.js';
import { RotaryDialPicker } from './rotary-picker.js';
import { sound } from './audio.js';
import { ExcelTimesheetImporter } from './excel-importer.js';

class ClassLoggerApp {
  constructor() {
    this.classes = [];
    this.categories = [...DEFAULT_CATEGORIES];
    this.selectedCategoryId = DEFAULT_CATEGORIES[0]?.id || 'cat_coach_allen';
    this.selectedDuration = 60;
    this.editingClassId = null;
    this.activeFilter = 'all';
    this.searchQuery = '';

    // Initialize DOM hooks
    this.initElements();
  }

  async init() {
    console.log('ClassLoggerApp initializing...');

    // 1. Render default categories immediately so cards are visible and clickable
    this.renderCategoryGrid();
    this.renderFilterTabs();

    // 2. Attach UI event listeners immediately so ALL buttons work right away
    this.attachEventListeners();
    this.updateSoundButtonUI();

    // 3. Initialize Rotary Dial Picker
    const mountEl = document.getElementById('rotary-picker-mount');
    if (mountEl) {
      this.rotaryPicker = new RotaryDialPicker(mountEl, {
        initialDate: new Date(),
        onChange: (val) => this.handleRotaryChange(val)
      });
      this.handleRotaryChange(this.rotaryPicker.getValue());
    }

    // 4. Connect to database and load records
    try {
      await db.init();
      const statusText = document.getElementById('db-status-text');
      if (statusText) {
        statusText.textContent = db.isIndexedDBAvailable ? 'Local DB: IndexedDB' : 'Local DB: LocalStorage';
      }

      await this.loadCategories();
      await this.loadClasses();

      // If empty, seed initial sample classes for instant demonstration
      if (this.classes.length === 0) {
        await db.seedSampleData();
        await this.loadClasses();
      }
    } catch (err) {
      console.error('Error during DB init/data load:', err);
    }

    console.log('ClassLoggerApp ready!');
  }

  initElements() {
    // Rotary elements
    this.rotaryDateDisplay = document.getElementById('rotary-date-display');
    this.rotaryTimeDisplay = document.getElementById('rotary-time-display');

    // Form elements
    this.classForm = document.getElementById('class-form');
    this.categoryGrid = document.getElementById('category-picker-grid');
    this.durationContainer = document.getElementById('duration-pills-container');
    this.durationLabel = document.getElementById('selected-duration-label');
    this.notesInput = document.getElementById('class-notes-input');
    this.quickTagsContainer = document.getElementById('quick-note-tags');
    this.btnSubmit = document.getElementById('btn-submit-class');
    this.btnSubmitText = document.getElementById('btn-submit-text');
    this.btnCancelEdit = document.getElementById('btn-cancel-edit');
    this.formHeading = document.getElementById('form-heading');
    this.modeTag = document.getElementById('mode-tag');

    // Quick pills
    this.btnSetNow = document.getElementById('btn-set-now');
    this.btnMinusHour = document.getElementById('btn-minus-hour');
    this.btnPlusHour = document.getElementById('btn-plus-hour');
    this.btnCalendarJump = document.getElementById('btn-calendar-jump');

    // History & list elements
    this.classesList = document.getElementById('classes-list');
    this.searchInput = document.getElementById('history-search-input');
    this.filterTabs = document.getElementById('history-filter-tabs');
    this.logCountTag = document.getElementById('log-count-tag');

    // Header buttons
    this.soundToggleBtn = document.getElementById('sound-toggle-btn');
    this.soundIcon = document.getElementById('sound-icon');
    this.dataMenuBtn = document.getElementById('data-menu-btn');

    // Modals
    this.modalCategory = document.getElementById('modal-category');
    this.categoryForm = document.getElementById('category-form');
    this.btnCloseCatModal = document.getElementById('btn-close-cat-modal');
    this.btnManageCategories = document.getElementById('btn-manage-categories');
    this.colorPaletteOptions = document.getElementById('color-palette-options');

    this.modalData = document.getElementById('modal-data');
    this.btnCloseDataModal = document.getElementById('btn-close-data-modal');
    this.btnExportCsv = document.getElementById('btn-export-csv');
    this.btnExportJson = document.getElementById('btn-export-json');
    this.importJsonInput = document.getElementById('import-json-input');
    this.btnSeedSample = document.getElementById('btn-seed-sample');
    this.btnClearDb = document.getElementById('btn-clear-db');

    // Excel import elements
    this.importExcelBtn = document.getElementById('import-excel-btn');
    this.importExcelModalInput = document.getElementById('import-excel-modal-input');
    this.btnImportSeptember = document.getElementById('btn-import-september');

    this.modalDateJump = document.getElementById('modal-date-jump');
    this.btnCloseDateModal = document.getElementById('btn-close-date-modal');
    this.dateJumpInput = document.getElementById('date-jump-input');
    this.btnConfirmDateJump = document.getElementById('btn-confirm-date-jump');

    this.toastContainer = document.getElementById('toast-container');
  }

  // --- ROTARY PICKER LISTENER ---

  handleRotaryChange(val) {
    if (!val) return;
    if (this.rotaryDateDisplay) {
      this.rotaryDateDisplay.textContent = val.dateLabel || val.date;
    }
    if (this.rotaryTimeDisplay) {
      this.rotaryTimeDisplay.textContent = val.displayTime;
    }
  }

  // --- DATA LOADING & RENDERING ---

  async loadCategories() {
    try {
      const cats = await db.getAllCategories();
      if (cats && cats.length > 0) {
        this.categories = cats;
      }
    } catch (e) {
      console.warn('Failed to load categories, using defaults:', e);
    }

    if (!this.selectedCategoryId && this.categories.length > 0) {
      const defaultCat = this.categories.find(c => c.name === 'Coach Allen') || this.categories[0];
      this.selectedCategoryId = defaultCat.id;
    }
    this.renderCategoryGrid();
    this.renderFilterTabs();
  }

  renderCategoryGrid() {
    if (!this.categoryGrid) return;
    this.categoryGrid.innerHTML = '';

    this.categories.forEach(cat => {
      const card = document.createElement('div');
      card.className = `category-card ${cat.id === this.selectedCategoryId ? 'is-selected' : ''}`;
      card.dataset.id = cat.id;
      card.style.setProperty('--category-color', cat.color || '#ff9f0a');
      card.style.setProperty('--category-glow', `${cat.color || '#ff9f0a'}40`);

      card.innerHTML = `
        <div class="category-icon">${cat.icon || '📌'}</div>
        <div class="category-name" title="${cat.name}">${cat.name}</div>
        <div class="category-check">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
            <polyline points="20 6 9 17 4 12"></polyline>
          </svg>
        </div>
      `;

      card.addEventListener('click', () => {
        this.selectedCategoryId = cat.id;
        sound.playTick(1.2);
        this.renderCategoryGrid();
      });

      this.categoryGrid.appendChild(card);
    });

    // Add "+ Custom Type" button
    const addCard = document.createElement('div');
    addCard.className = 'category-card add-category-btn';
    addCard.innerHTML = `
      <div class="category-icon">➕</div>
      <div class="category-name">New Type</div>
    `;
    addCard.addEventListener('click', () => {
      this.openModal(this.modalCategory);
    });
    this.categoryGrid.appendChild(addCard);
  }

  renderFilterTabs() {
    if (!this.filterTabs) return;
    this.filterTabs.innerHTML = '';

    // "All Classes" Tab
    const allTab = document.createElement('button');
    allTab.type = 'button';
    allTab.className = `filter-tab ${this.activeFilter === 'all' ? 'is-active' : ''}`;
    allTab.dataset.cat = 'all';
    allTab.innerHTML = `All Classes <span class="filter-count">${this.classes.length}</span>`;
    allTab.addEventListener('click', () => {
      this.activeFilter = 'all';
      this.renderFilterTabs();
      this.renderClassesList();
    });
    this.filterTabs.appendChild(allTab);

    // Individual Category Tabs
    this.categories.forEach(cat => {
      const count = this.classes.filter(c => c.category === cat.name).length;
      const tab = document.createElement('button');
      tab.type = 'button';
      tab.className = `filter-tab ${this.activeFilter === cat.id ? 'is-active' : ''}`;
      tab.dataset.cat = cat.id;
      tab.innerHTML = `${cat.icon || ''} ${cat.name} <span class="filter-count">${count}</span>`;
      tab.addEventListener('click', () => {
        this.activeFilter = cat.id;
        this.renderFilterTabs();
        this.renderClassesList();
      });
      this.filterTabs.appendChild(tab);
    });
  }

  async loadClasses() {
    try {
      this.classes = await db.getAllClasses();
    } catch (e) {
      console.warn('Error loading classes:', e);
      this.classes = [];
    }
    this.renderClassesList();
    this.renderFilterTabs();
  }

  renderClassesList() {
    if (!this.classesList) return;

    let filtered = [...this.classes];

    // Filter by active category tab
    if (this.activeFilter !== 'all') {
      const targetCat = this.categories.find(c => c.id === this.activeFilter);
      if (targetCat) {
        filtered = filtered.filter(c => c.category === targetCat.name);
      }
    }

    // Filter by search query
    if (this.searchQuery.trim()) {
      const q = this.searchQuery.toLowerCase();
      filtered = filtered.filter(c => 
        (c.note && c.note.toLowerCase().includes(q)) ||
        (c.category && c.category.toLowerCase().includes(q)) ||
        (c.date && c.date.includes(q)) ||
        (c.displayTime && c.displayTime.toLowerCase().includes(q))
      );
    }

    if (this.logCountTag) {
      this.logCountTag.textContent = `${filtered.length} session${filtered.length === 1 ? '' : 's'}`;
    }

    if (filtered.length === 0) {
      this.classesList.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-icon">⏱️</div>
          <h3 class="empty-state-title">No Classes Found</h3>
          <p class="empty-state-desc">
            ${this.searchQuery ? 'No sessions matched your search query.' : 'Select the date and time using the rotary dials and log your first class!'}
          </p>
          ${!this.searchQuery ? `
            <button type="button" class="quick-pill-btn primary" id="btn-empty-seed" style="margin: 0 auto;">
              🧪 Insert Demo Classes
            </button>
          ` : ''}
        </div>
      `;

      const emptySeedBtn = document.getElementById('btn-empty-seed');
      if (emptySeedBtn) {
        emptySeedBtn.addEventListener('click', async () => {
          await db.seedSampleData();
          await this.loadClasses();
          this.showToast('Sample class sessions inserted!', 'success');
        });
      }
      return;
    }

    this.classesList.innerHTML = '';

    filtered.forEach(item => {
      const catObj = this.categories.find(c => c.name === item.category) || {
        color: '#ff9f0a',
        icon: '📌'
      };

      // Actual formatted calendar date
      const itemDate = new Date(item.date + 'T00:00:00');
      const formattedDate = itemDate.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });
      const today = new Date();
      today.setHours(0,0,0,0);
      const isToday = (today.getTime() === itemDate.getTime());

      const card = document.createElement('article');
      card.className = 'class-log-card';
      card.style.setProperty('--card-accent', catObj.color);

      card.innerHTML = `
        <div class="card-top-row">
          <div class="card-category-badge">
            <span>${catObj.icon}</span>
            <span>${item.category}</span>
          </div>
          <div class="card-date-group">
            ${isToday ? '<span class="card-relative-badge is-today">Today</span>' : ''}
            <span>${formattedDate}</span>
          </div>
        </div>

        <div class="card-main-row">
          <div class="card-time-display">
            <span class="card-time-large">${item.displayTime || item.time}</span>
            <span class="card-duration-badge">${item.duration || 60}m</span>
          </div>

          <div class="card-actions">
            <button type="button" class="card-action-btn edit" title="Edit this entry" data-id="${item.id}">
              ✏️
            </button>
            <button type="button" class="card-action-btn duplicate" title="Duplicate entry" data-id="${item.id}">
              📋
            </button>
            <button type="button" class="card-action-btn delete" title="Delete entry" data-id="${item.id}">
              🗑️
            </button>
          </div>
        </div>

        ${item.note ? `
          <div class="card-note-box">${this.escapeHTML(item.note)}</div>
        ` : ''}
      `;

      // Card action buttons
      card.querySelector('.edit').addEventListener('click', () => this.startEditClass(item));
      card.querySelector('.duplicate').addEventListener('click', () => this.duplicateClass(item));
      card.querySelector('.delete').addEventListener('click', () => this.deleteClass(item.id));

      this.classesList.appendChild(card);
    });
  }

  // --- ATTACH EVENT LISTENERS ---

  attachEventListeners() {
    // 1. Duration Pills
    if (this.durationContainer) {
      this.durationContainer.querySelectorAll('.duration-pill').forEach(pill => {
        pill.addEventListener('click', () => {
          this.durationContainer.querySelectorAll('.duration-pill').forEach(p => p.classList.remove('is-active'));
          pill.classList.add('is-active');
          this.selectedDuration = parseInt(pill.dataset.mins, 10);
          if (this.durationLabel) {
            this.durationLabel.textContent = `${this.selectedDuration} mins`;
          }
          sound.playTick(1.1);
        });
      });
    }

    // 2. Quick Note Tags
    if (this.quickTagsContainer) {
      this.quickTagsContainer.querySelectorAll('.note-tag-chip').forEach(chip => {
        chip.addEventListener('click', () => {
          const tag = chip.dataset.tag;
          const current = this.notesInput.value.trim();
          if (current) {
            this.notesInput.value = `${current}\n• ${tag}`;
          } else {
            this.notesInput.value = `• ${tag}`;
          }
          this.notesInput.focus();
          sound.playTick(1.3);
        });
      });
    }

    // 3. Quick Rotary Controls
    if (this.btnSetNow) {
      this.btnSetNow.addEventListener('click', () => {
        if (this.rotaryPicker) this.rotaryPicker.setNow(true);
        sound.playTick(1.0);
        this.showToast('Set to current time', 'info');
      });
    }

    if (this.btnMinusHour) {
      this.btnMinusHour.addEventListener('click', () => {
        if (this.rotaryPicker) this.rotaryPicker.adjustHours(-1, true);
        sound.playTick(0.9);
      });
    }

    if (this.btnPlusHour) {
      this.btnPlusHour.addEventListener('click', () => {
        if (this.rotaryPicker) this.rotaryPicker.adjustHours(1, true);
        sound.playTick(1.1);
      });
    }

    if (this.btnCalendarJump) {
      this.btnCalendarJump.addEventListener('click', () => {
        if (this.rotaryPicker) {
          const currentVal = this.rotaryPicker.getValue();
          if (this.dateJumpInput) this.dateJumpInput.value = currentVal.date;
        }
        this.openModal(this.modalDateJump);
      });
    }

    if (this.btnConfirmDateJump) {
      this.btnConfirmDateJump.addEventListener('click', () => {
        const selected = this.dateJumpInput ? this.dateJumpInput.value : '';
        if (selected && this.rotaryPicker) {
          const [y, m, d] = selected.split('-').map(Number);
          const cur = this.rotaryPicker.getValue().fullDate;
          cur.setFullYear(y, m - 1, d);
          this.rotaryPicker.setValue(cur, true);
          this.closeModal(this.modalDateJump);
          sound.playTick(1.2);
          this.showToast(`Jumped to ${selected}`, 'info');
        }
      });
    }

    // 4. Form Submission
    if (this.btnSubmit) {
      this.btnSubmit.addEventListener('click', (e) => {
        e.preventDefault();
        this.handleSubmitClass();
      });
    }

    if (this.classForm) {
      this.classForm.addEventListener('submit', (e) => {
        e.preventDefault();
        this.handleSubmitClass();
      });
    }

    if (this.btnCancelEdit) {
      this.btnCancelEdit.addEventListener('click', () => {
        this.cancelEdit();
      });
    }

    // 5. Search filtering
    if (this.searchInput) {
      this.searchInput.addEventListener('input', (e) => {
        this.searchQuery = e.target.value;
        this.renderClassesList();
      });
    }

    // 6. Sound Toggle
    if (this.soundToggleBtn) {
      this.soundToggleBtn.addEventListener('click', () => {
        const isMuted = sound.toggleMute();
        this.updateSoundButtonUI();
        if (!isMuted) sound.playTick(1.0);
        this.showToast(isMuted ? 'Muted mechanical sound' : 'Sound feedback enabled', 'info');
      });
    }

    // 7. Modal Open & Close Triggers
    if (this.btnManageCategories) {
      this.btnManageCategories.addEventListener('click', () => this.openModal(this.modalCategory));
    }
    if (this.btnCloseCatModal) {
      this.btnCloseCatModal.addEventListener('click', () => this.closeModal(this.modalCategory));
    }

    if (this.dataMenuBtn) {
      this.dataMenuBtn.addEventListener('click', () => this.openModal(this.modalData));
    }
    if (this.btnCloseDataModal) {
      this.btnCloseDataModal.addEventListener('click', () => this.closeModal(this.modalData));
    }
    if (this.btnCloseDateModal) {
      this.btnCloseDateModal.addEventListener('click', () => this.closeModal(this.modalDateJump));
    }

    // Close modals on clicking overlay background
    [this.modalCategory, this.modalData, this.modalDateJump].forEach(modal => {
      if (modal) {
        modal.addEventListener('click', (e) => {
          if (e.target === modal) this.closeModal(modal);
        });
      }
    });

    // 8. Custom Category Creation
    if (this.colorPaletteOptions) {
      this.colorPaletteOptions.querySelectorAll('.color-option').forEach(opt => {
        opt.addEventListener('click', () => {
          this.colorPaletteOptions.querySelectorAll('.color-option').forEach(o => o.classList.remove('is-selected'));
          opt.classList.add('is-selected');
        });
      });
    }

    if (this.categoryForm) {
      this.categoryForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        await this.handleSaveCustomCategory();
      });
    }

    // 9. Data Management & Backups
    if (this.btnExportCsv) {
      this.btnExportCsv.addEventListener('click', async () => {
        const csv = await db.exportCSV();
        this.downloadFile(csv, `class-logs-${new Date().toISOString().split('T')[0]}.csv`, 'text/csv');
        this.showToast('Exported classes as CSV', 'success');
      });
    }

    if (this.btnExportJson) {
      this.btnExportJson.addEventListener('click', async () => {
        const json = await db.exportJSON();
        this.downloadFile(json, `class-logger-backup-${new Date().toISOString().split('T')[0]}.json`, 'application/json');
        this.showToast('Backup JSON downloaded', 'success');
      });
    }

    if (this.importJsonInput) {
      this.importJsonInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        try {
          const text = await file.text();
          const res = await db.importJSON(text);
          await this.loadCategories();
          await this.loadClasses();
          this.closeModal(this.modalData);
          this.showToast(`Restored ${res.count} sessions from backup!`, 'success');
        } catch (err) {
          this.showToast('Invalid backup file', 'danger');
        }
      });
    }

    if (this.btnSeedSample) {
      this.btnSeedSample.addEventListener('click', async () => {
        await db.seedSampleData();
        await this.loadClasses();
        this.closeModal(this.modalData);
        this.showToast('Demo sessions added!', 'success');
      });
    }

    if (this.btnClearDb) {
      this.btnClearDb.addEventListener('click', async () => {
        if (confirm('Are you sure you want to clear all logged class records? This cannot be undone.')) {
          await db.clearAllClasses();
          await this.loadClasses();
          sound.playTrash();
          this.closeModal(this.modalData);
          this.showToast('Local database cleared', 'danger');
        }
      });
    }

    // 10. Excel Timesheet Import
    const excelChangeHandler = (e) => {
      const file = e.target.files[0];
      if (file) this.handleExcelImport(file);
      e.target.value = ''; // reset so same file can be re-selected
    };

    if (this.importExcelBtn) {
      this.importExcelBtn.addEventListener('change', excelChangeHandler);
    }
    if (this.importExcelModalInput) {
      this.importExcelModalInput.addEventListener('change', excelChangeHandler);
    }
    if (this.btnImportSeptember) {
      this.btnImportSeptember.addEventListener('click', () => this.handleOneClickSeptemberImport());
    }
  }

  updateSoundButtonUI() {
    if (!this.soundIcon || !this.soundToggleBtn) return;
    const muted = sound.isMuted();
    this.soundIcon.textContent = muted ? '🔇' : '🔊';
    this.soundToggleBtn.title = muted ? 'Unmute Sound' : 'Mute Sound';
  }

  // --- CRUD ACTIONS ---

  async handleSubmitClass() {
    if (!this.rotaryPicker) return;

    const rotaryVal = this.rotaryPicker.getValue();
    const selectedCategory = this.categories.find(c => c.id === this.selectedCategoryId);

    if (!selectedCategory) {
      this.showToast('Please select a class category', 'danger');
      return;
    }

    const note = this.notesInput ? this.notesInput.value.trim() : '';

    const recordData = {
      date: rotaryVal.date,
      time: rotaryVal.time24,
      displayTime: rotaryVal.displayTime,
      timestamp: rotaryVal.timestamp,
      category: selectedCategory.name,
      duration: this.selectedDuration,
      note: note
    };

    if (this.editingClassId) {
      // Update existing record
      await db.updateClass(this.editingClassId, recordData);
      this.showToast(`Updated session for ${selectedCategory.name}`, 'success');
      sound.playSuccess();
      this.cancelEdit();
    } else {
      // Create new record
      await db.addClass(recordData);
      this.showToast(`Logged session for ${selectedCategory.name} at ${rotaryVal.displayTime}`, 'success');
      sound.playSuccess();

      // Reset note field
      if (this.notesInput) this.notesInput.value = '';
    }

    await this.loadClasses();
  }

  startEditClass(item) {
    this.editingClassId = item.id;

    // Load date into rotary picker
    if (this.rotaryPicker) {
      const [y, m, d] = item.date.split('-').map(Number);
      const [hh, mm] = (item.time || '12:00').split(':').map(Number);
      const dateObj = new Date(y, m - 1, d, hh, mm);
      this.rotaryPicker.setValue(dateObj, true);
    }

    // Set category
    const cat = this.categories.find(c => c.name === item.category);
    if (cat) {
      this.selectedCategoryId = cat.id;
      this.renderCategoryGrid();
    }

    // Set duration
    this.selectedDuration = item.duration || 60;
    if (this.durationLabel) {
      this.durationLabel.textContent = `${this.selectedDuration} mins`;
    }
    if (this.durationContainer) {
      this.durationContainer.querySelectorAll('.duration-pill').forEach(pill => {
        pill.classList.toggle('is-active', parseInt(pill.dataset.mins, 10) === this.selectedDuration);
      });
    }

    // Set note
    if (this.notesInput) {
      this.notesInput.value = item.note || '';
    }

    // Update UI for editing
    if (this.btnSubmit) this.btnSubmit.classList.add('is-editing');
    if (this.btnSubmitText) this.btnSubmitText.textContent = 'Save Changes to Database';
    if (this.btnCancelEdit) this.btnCancelEdit.style.display = 'block';
    if (this.formHeading) this.formHeading.innerHTML = `<span>✏️</span> Edit Class Session`;
    if (this.modeTag) this.modeTag.textContent = 'Editing';

    // Scroll smoothly to form
    if (this.classForm) {
      this.classForm.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
    sound.playTick(1.2);
  }

  cancelEdit() {
    this.editingClassId = null;
    if (this.btnSubmit) this.btnSubmit.classList.remove('is-editing');
    if (this.btnSubmitText) this.btnSubmitText.textContent = 'Enter Class Into Database';
    if (this.btnCancelEdit) this.btnCancelEdit.style.display = 'none';
    if (this.formHeading) this.formHeading.innerHTML = `<span>⏱️</span> Log Class Session`;
    if (this.modeTag) this.modeTag.textContent = 'New Entry';
    if (this.notesInput) this.notesInput.value = '';
  }

  async duplicateClass(item) {
    const clone = {
      ...item,
      id: undefined,
      note: item.note ? `${item.note} (Copy)` : ''
    };
    await db.addClass(clone);
    await this.loadClasses();
    sound.playSuccess();
    this.showToast(`Duplicated session for ${item.category}`, 'success');
  }

  async deleteClass(id) {
    if (confirm('Delete this class session?')) {
      await db.deleteClass(id);
      sound.playTrash();
      await this.loadClasses();
      this.showToast('Class entry deleted', 'info');
    }
  }

  async handleSaveCustomCategory() {
    const nameInput = document.getElementById('new-cat-name');
    const iconInput = document.getElementById('new-cat-icon');
    const name = nameInput ? nameInput.value.trim() : '';
    const icon = (iconInput ? iconInput.value.trim() : '') || '⚡';

    if (!name) return;

    const selectedColorEl = this.colorPaletteOptions ? this.colorPaletteOptions.querySelector('.color-option.is-selected') : null;
    const color = selectedColorEl ? selectedColorEl.dataset.color : '#ff9f0a';

    const newCat = await db.addCategory({
      name,
      icon,
      color
    });

    this.selectedCategoryId = newCat.id;
    await this.loadCategories();
    this.closeModal(this.modalCategory);
    if (nameInput) nameInput.value = '';
    sound.playSuccess();
    this.showToast(`Category "${name}" created!`, 'success');
  }

  // --- MODALS & UTILS ---

  openModal(modalEl) {
    if (!modalEl) return;
    modalEl.classList.add('is-open');
    sound.playTick(1.0);
  }

  closeModal(modalEl) {
    if (!modalEl) return;
    modalEl.classList.remove('is-open');
  }

  showToast(message, type = 'info') {
    if (!this.toastContainer) return;
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    let icon = 'ℹ️';
    if (type === 'success') icon = '✓';
    if (type === 'danger') icon = '⚠️';

    toast.innerHTML = `<span>${icon}</span> <span>${this.escapeHTML(message)}</span>`;
    this.toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.animation = 'toast-in 0.3s ease reverse forwards';
      setTimeout(() => toast.remove(), 300);
    }, 3200);
  }

  downloadFile(content, fileName, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  escapeHTML(str) {
    if (!str) return '';
    return str.replace(/[&<>'"]/g, 
      tag => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[tag] || tag)
    );
  }

  // --- EXCEL IMPORT ---

  async handleExcelImport(file) {
    this.showToast(`Parsing "${file.name}"...`, 'info');
    try {
      const arrayBuffer = await file.arrayBuffer();
      const result = ExcelTimesheetImporter.parseWorkbook(arrayBuffer);

      if (!result.classes.length) {
        this.showToast('No valid class sessions found in spreadsheet', 'danger');
        return;
      }

      // Ensure categories from the spreadsheet exist in the database
      for (const catName of result.categories) {
        const existing = this.categories.find(c => c.name.toLowerCase() === catName.toLowerCase());
        if (!existing) {
          const newCat = {
            id: 'cat_import_' + catName.toLowerCase().replace(/[^a-z0-9]/g, '_'),
            name: catName,
            color: `hsl(${Math.floor(Math.random() * 360)}, 60%, 55%)`,
            icon: '📋'
          };
          await db.addCategory(newCat);
        }
      }

      // Reload categories so the grid reflects new ones
      await this.loadCategories();

      // Add each class to the database
      let imported = 0;
      for (const cls of result.classes) {
        // Map category name to category id
        const matchedCat = this.categories.find(c => c.name.toLowerCase() === cls.category.toLowerCase());
        if (matchedCat) cls.categoryId = matchedCat.id;
        await db.addClass(cls);
        imported++;
      }

      await this.loadClasses();
      this.closeModal(this.modalData);
      sound.playTick(1.5);
      this.showToast(`Imported ${imported} class sessions from "${result.sheetName}" sheet!`, 'success');
    } catch (err) {
      console.error('Excel import error:', err);
      this.showToast(`Import error: ${err.message}`, 'danger');
    }
  }

  async handleOneClickSeptemberImport() {
    this.showToast('Fetching September 2026 timesheet...', 'info');
    try {
      const resp = await fetch('/JieHern_CoachTimesheet September2026.xlsx');
      if (!resp.ok) {
        this.showToast('Could not find the September 2026 xlsx file in the project root', 'danger');
        return;
      }
      const arrayBuffer = await resp.arrayBuffer();
      const blob = new Blob([arrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const file = new File([blob], 'JieHern_CoachTimesheet September2026.xlsx');
      await this.handleExcelImport(file);
    } catch (err) {
      console.error('One-click import error:', err);
      this.showToast(`Failed to auto-import: ${err.message}`, 'danger');
    }
  }
}

// Robust launch pattern
function startApp() {
  const app = new ClassLoggerApp();
  app.init().catch(err => console.error('Error starting ClassLogger:', err));
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', startApp);
} else {
  startApp();
}
