/**
 * app.js - Main Application Orchestrator for ClassLogger
 * Integrates Apple Rotary Dial Picker, Localized IndexedDB Storage, Categories, Audio & Analytics
 */

import { db, DEFAULT_CATEGORIES } from './db.js';
import { RotaryDialPicker } from './rotary-picker.js';
import { sound } from './audio.js';
import { ExcelTimesheetImporter, ExcelTimesheetExporter } from './excel-importer.js';

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

      // Clean up / migrate any legacy sessions previously imported under subcategories like LTS, PreComp, Swim Clinic
      await this.migrateLegacyImportCategories();
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
    this.customDurationWrapper = document.getElementById('custom-duration-wrapper');
    this.customDurationInput = document.getElementById('custom-duration-input');
    this.btnApplyCustomDuration = document.getElementById('btn-apply-custom-duration');
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

    // Filtered export elements
    this.btnHeaderExport = document.getElementById('btn-header-export');
    this.btnOpenExportModal = document.getElementById('btn-open-export-modal');
    this.modalExport = document.getElementById('modal-export');
    this.btnCloseExportModal = document.getElementById('btn-close-export-modal');
    this.exportCategorySelect = document.getElementById('export-category-select');
    this.exportDateFrom = document.getElementById('export-date-from');
    this.exportDateTo = document.getElementById('export-date-to');
    this.exportPresetPills = document.getElementById('export-preset-pills');
    this.exportPreviewCount = document.getElementById('export-preview-count');
    this.exportPreviewDuration = document.getElementById('export-preview-duration');
    this.exportPreviewRange = document.getElementById('export-preview-range');
    this.exportCoachName = document.getElementById('export-coach-name');
    this.exportRatePerClass = document.getElementById('export-rate-per-class');
    this.exportFeesLabel = document.getElementById('export-fees-label');
    this.exportCalcTotalClasses = document.getElementById('export-calc-total-classes');
    this.exportCalcRate = document.getElementById('export-calc-rate');
    this.exportCalcTotalFees = document.getElementById('export-calc-total-fees');

    this.btnExportFilteredXlsx = document.getElementById('btn-export-filtered-xlsx');
    this.btnExportFilteredCsv = document.getElementById('btn-export-filtered-csv');
    this.btnExportFilteredJson = document.getElementById('btn-export-filtered-json');

    this.toastContainer = document.getElementById('toast-container');

    // Mobile Navigation & Viewport elements
    this.mobileNavTabs = document.getElementById('mobile-nav-tabs');
    this.mobileTabLog = document.getElementById('mobile-tab-log');
    this.mobileTabHistory = document.getElementById('mobile-tab-history');
    this.mobileHistoryBadge = document.getElementById('mobile-history-badge');
    this.mainGrid = document.querySelector('.main-grid');
    this.activeMobileTab = 'log';
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
    if (this.mobileHistoryBadge) {
      this.mobileHistoryBadge.textContent = this.classes.length;
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

          if (pill.dataset.mins === 'custom') {
            if (this.customDurationWrapper) this.customDurationWrapper.style.display = 'flex';
            if (this.customDurationInput) {
              this.customDurationInput.focus();
              const existingVal = parseInt(this.customDurationInput.value, 10);
              if (existingVal && existingVal > 0) {
                this.selectedDuration = existingVal;
                if (this.durationLabel) this.durationLabel.textContent = `${this.selectedDuration} mins`;
              }
            }
          } else {
            if (this.customDurationWrapper) this.customDurationWrapper.style.display = 'none';
            this.selectedDuration = parseInt(pill.dataset.mins, 10);
            if (this.durationLabel) {
              this.durationLabel.textContent = `${this.selectedDuration} mins`;
            }
          }
          sound.playTick(1.1);
        });
      });
    }

    if (this.customDurationInput) {
      const handleCustomInput = () => {
        const val = parseInt(this.customDurationInput.value, 10);
        if (val && val > 0) {
          this.selectedDuration = val;
          if (this.durationLabel) {
            this.durationLabel.textContent = `${this.selectedDuration} mins`;
          }
        }
      };
      this.customDurationInput.addEventListener('input', handleCustomInput);
      this.customDurationInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          handleCustomInput();
          sound.playTick(1.2);
          if (this.notesInput) this.notesInput.focus();
        }
      });
    }

    if (this.btnApplyCustomDuration) {
      this.btnApplyCustomDuration.addEventListener('click', () => {
        const val = parseInt(this.customDurationInput?.value, 10);
        if (val && val > 0) {
          this.selectedDuration = val;
          if (this.durationLabel) {
            this.durationLabel.textContent = `${this.selectedDuration} mins`;
          }
          sound.playTick(1.2);
          this.showToast(`Custom duration set to ${val} mins`, 'info');
        } else {
          this.showToast('Please enter a valid number of minutes', 'danger');
        }
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

    // Filtered Export triggers
    if (this.btnHeaderExport) {
      this.btnHeaderExport.addEventListener('click', () => this.openExportModal());
    }
    if (this.btnOpenExportModal) {
      this.btnOpenExportModal.addEventListener('click', () => {
        this.closeModal(this.modalData);
        this.openExportModal();
      });
    }
    if (this.btnCloseExportModal) {
      this.btnCloseExportModal.addEventListener('click', () => this.closeModal(this.modalExport));
    }

    // Close modals on clicking overlay background
    [this.modalCategory, this.modalData, this.modalDateJump, this.modalExport].forEach(modal => {
      if (modal) {
        modal.addEventListener('click', (e) => {
          if (e.target === modal) this.closeModal(modal);
        });
      }
    });

    // Export modal filters & buttons
    if (this.exportCategorySelect) {
      this.exportCategorySelect.addEventListener('change', () => this.updateExportPreview());
    }
    if (this.exportDateFrom) {
      this.exportDateFrom.addEventListener('change', () => {
        this.clearActivePresetPill();
        this.updateExportPreview();
      });
    }
    if (this.exportDateTo) {
      this.exportDateTo.addEventListener('change', () => {
        this.clearActivePresetPill();
        this.updateExportPreview();
      });
    }
    if (this.exportPresetPills) {
      this.exportPresetPills.querySelectorAll('.quick-pill-btn').forEach(pill => {
        pill.addEventListener('click', () => {
          this.applyExportDatePreset(pill.dataset.preset);
        });
      });
    }
    if (this.exportRatePerClass) {
      this.exportRatePerClass.addEventListener('input', () => this.updateExportPreview());
    }
    if (this.exportFeesLabel) {
      this.exportFeesLabel.addEventListener('input', () => {
        this.exportFeesLabel.dataset.userEdited = 'true';
        this.updateExportPreview();
      });
    }
    if (this.btnExportFilteredXlsx) {
      this.btnExportFilteredXlsx.addEventListener('click', () => this.handleExportXlsx());
    }
    if (this.btnExportFilteredCsv) {
      this.btnExportFilteredCsv.addEventListener('click', () => this.handleExportCsv());
    }
    if (this.btnExportFilteredJson) {
      this.btnExportFilteredJson.addEventListener('click', () => this.handleExportJson());
    }

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

    // 10. Mobile Segmented Switcher tabs
    if (this.mobileTabLog) {
      this.mobileTabLog.addEventListener('click', () => this.switchMobileTab('log'));
    }
    if (this.mobileTabHistory) {
      this.mobileTabHistory.addEventListener('click', () => this.switchMobileTab('history'));
    }
  }

  switchMobileTab(tab) {
    this.activeMobileTab = tab;
    if (this.mobileTabLog && this.mobileTabHistory) {
      this.mobileTabLog.classList.toggle('is-active', tab === 'log');
      this.mobileTabHistory.classList.toggle('is-active', tab === 'history');
    }
    if (this.mainGrid) {
      this.mainGrid.classList.remove('mobile-view-log', 'mobile-view-history');
      this.mainGrid.classList.add(tab === 'log' ? 'mobile-view-log' : 'mobile-view-history');
    }
    window.scrollTo({ top: 0, behavior: 'smooth' });
    sound.playTick(1.2);
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
    this.switchMobileTab('log');

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
      let matched = false;
      this.durationContainer.querySelectorAll('.duration-pill').forEach(pill => {
        const isMatch = parseInt(pill.dataset.mins, 10) === this.selectedDuration;
        pill.classList.toggle('is-active', isMatch);
        if (isMatch) matched = true;
      });

      const customPill = this.durationContainer.querySelector('[data-mins="custom"]');
      if (!matched) {
        if (customPill) customPill.classList.add('is-active');
        if (this.customDurationWrapper) this.customDurationWrapper.style.display = 'flex';
        if (this.customDurationInput) this.customDurationInput.value = this.selectedDuration;
      } else {
        if (customPill) customPill.classList.remove('is-active');
        if (this.customDurationWrapper) this.customDurationWrapper.style.display = 'none';
      }
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
    if (this.customDurationWrapper) this.customDurationWrapper.style.display = 'none';
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

      // Ensure Aerosplash category exists in the database
      let aerosplashCat = this.categories.find(c => c.name.toLowerCase() === 'aerosplash');
      if (!aerosplashCat) {
        aerosplashCat = await db.addCategory({
          id: 'cat_aerosplash',
          name: 'Aerosplash',
          color: '#0a84ff',
          icon: '🌊',
          isDefault: true
        });
        await this.loadCategories();
      }

      // Add each class strictly under the Aerosplash category
      let imported = 0;
      for (const cls of result.classes) {
        cls.category = 'Aerosplash';
        cls.categoryId = aerosplashCat ? aerosplashCat.id : 'cat_aerosplash';
        await db.addClass(cls);
        imported++;
      }

      await this.loadClasses();
      this.closeModal(this.modalData);
      sound.playTick(1.5);
      this.showToast(`Imported ${imported} class sessions strictly under Aerosplash!`, 'success');
    } catch (err) {
      console.error('Excel import error:', err);
      this.showToast(`Import error: ${err.message}`, 'danger');
    }
  }

  async handleOneClickSeptemberImport() {
    this.showToast('Fetching September 2026 timesheet...', 'info');
    try {
      let resp = await fetch('/JieHern_CoachTimesheet September2026.xlsx');
      if (!resp.ok) {
        resp = await fetch('/JieHern_CoachTimesheet%20September2026.xlsx');
      }
      if (!resp.ok) {
        this.showToast('Could not find the September 2026 xlsx file', 'danger');
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

  // --- FILTERED EXPORT FUNCTION ---

  openExportModal() {
    this.populateExportCategorySelect();
    this.updateExportPreview();
    this.openModal(this.modalExport);
  }

  populateExportCategorySelect() {
    if (!this.exportCategorySelect) return;
    const currentVal = this.exportCategorySelect.value || 'all';
    this.exportCategorySelect.innerHTML = '';

    const allOpt = document.createElement('option');
    allOpt.value = 'all';
    allOpt.textContent = `🌟 All Categories (${this.classes.length} sessions)`;
    this.exportCategorySelect.appendChild(allOpt);

    this.categories.forEach(cat => {
      const count = this.classes.filter(c => c.category && c.category.toLowerCase() === cat.name.toLowerCase()).length;
      const opt = document.createElement('option');
      opt.value = cat.name;
      opt.textContent = `${cat.icon || '📌'} ${cat.name} (${count} session${count === 1 ? '' : 's'})`;
      this.exportCategorySelect.appendChild(opt);
    });

    if ([...this.exportCategorySelect.options].some(o => o.value.toLowerCase() === currentVal.toLowerCase())) {
      this.exportCategorySelect.value = currentVal;
    } else {
      this.exportCategorySelect.value = 'all';
    }
  }

  applyExportDatePreset(preset) {
    if (!this.exportPresetPills) return;
    this.exportPresetPills.querySelectorAll('.quick-pill-btn').forEach(p => {
      p.classList.toggle('is-active', p.dataset.preset === preset);
    });

    const now = new Date();
    const y = now.getFullYear();
    const m = now.getMonth();

    const formatDateStr = (d) => {
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const dd = String(d.getDate()).padStart(2, '0');
      return `${yyyy}-${mm}-${dd}`;
    };

    if (preset === 'all') {
      if (this.exportDateFrom) this.exportDateFrom.value = '';
      if (this.exportDateTo) this.exportDateTo.value = '';
    } else if (preset === 'this-month') {
      const start = new Date(y, m, 1);
      const end = new Date(y, m + 1, 0);
      if (this.exportDateFrom) this.exportDateFrom.value = formatDateStr(start);
      if (this.exportDateTo) this.exportDateTo.value = formatDateStr(end);
    } else if (preset === 'last-month') {
      const start = new Date(y, m - 1, 1);
      const end = new Date(y, m, 0);
      if (this.exportDateFrom) this.exportDateFrom.value = formatDateStr(start);
      if (this.exportDateTo) this.exportDateTo.value = formatDateStr(end);
    } else if (preset === 'last-30') {
      const end = now;
      const start = new Date(now.getTime() - 30 * 24 * 3600 * 1000);
      if (this.exportDateFrom) this.exportDateFrom.value = formatDateStr(start);
      if (this.exportDateTo) this.exportDateTo.value = formatDateStr(end);
    } else if (preset === 'this-year') {
      if (this.exportDateFrom) this.exportDateFrom.value = `${y}-01-01`;
      if (this.exportDateTo) this.exportDateTo.value = `${y}-12-31`;
    }

    this.updateExportPreview();
    sound.playTick(1.1);
  }

  clearActivePresetPill() {
    if (!this.exportPresetPills) return;
    this.exportPresetPills.querySelectorAll('.quick-pill-btn').forEach(p => p.classList.remove('is-active'));
  }

  getFilteredExportData() {
    const category = this.exportCategorySelect ? this.exportCategorySelect.value : 'all';
    const dateFrom = this.exportDateFrom ? this.exportDateFrom.value.trim() : '';
    const dateTo = this.exportDateTo ? this.exportDateTo.value.trim() : '';

    const filtered = this.classes.filter(c => {
      if (category && category !== 'all') {
        if (!c.category || c.category.toLowerCase() !== category.toLowerCase()) return false;
      }
      if (dateFrom && c.date < dateFrom) return false;
      if (dateTo && c.date > dateTo) return false;
      return true;
    });

    return {
      classes: filtered,
      category,
      dateFrom,
      dateTo
    };
  }

  updateExportPreview() {
    const { classes, category, dateFrom, dateTo } = this.getFilteredExportData();
    const count = classes.length;

    if (this.exportPreviewCount) {
      this.exportPreviewCount.textContent = `${count} session${count === 1 ? '' : 's'} found`;
      this.exportPreviewCount.classList.toggle('is-empty', count === 0);
    }

    if (this.exportPreviewDuration) {
      const totalMins = classes.reduce((sum, c) => sum + (Number(c.duration) || 60), 0);
      const hours = Math.floor(totalMins / 60);
      const mins = totalMins % 60;
      this.exportPreviewDuration.textContent = `${hours}h ${mins}m (${(totalMins / 60).toFixed(1)} hrs)`;
    }

    if (this.exportPreviewRange) {
      const catLabel = category === 'all' ? 'All categories' : category;
      let dateLabel = 'All recorded dates';
      if (dateFrom && dateTo) {
        dateLabel = `${dateFrom} to ${dateTo}`;
      } else if (dateFrom) {
        dateLabel = `From ${dateFrom} onwards`;
      } else if (dateTo) {
        dateLabel = `Up to ${dateTo}`;
      }
      this.exportPreviewRange.textContent = `${catLabel} • ${dateLabel}`;
    }

    // Live Billing & Fee Calculations
    const rateVal = this.exportRatePerClass ? (parseFloat(this.exportRatePerClass.value) || 0) : 40.0;
    const totalFees = count * rateVal;

    if (this.exportCalcTotalClasses) {
      this.exportCalcTotalClasses.textContent = `${count} class${count === 1 ? '' : 'es'}`;
    }
    if (this.exportCalcRate) {
      this.exportCalcRate.textContent = `RM ${rateVal.toFixed(1)}`;
    }
    if (this.exportCalcTotalFees) {
      this.exportCalcTotalFees.textContent = `RM ${totalFees.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    }

    // Auto-suggest month and year for fees label placeholder if not manually edited
    if (classes.length && this.exportFeesLabel && !this.exportFeesLabel.dataset.userEdited) {
      const firstDate = classes[0].date || '';
      if (firstDate) {
        const [y, m] = firstDate.split('-').map(Number);
        const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
        const mName = monthNames[m - 1] || 'September';
        this.exportFeesLabel.placeholder = `${mName} ${y || 2026} fees :`;
      }
    }

    const isDisabled = count === 0;
    if (this.btnExportFilteredXlsx) this.btnExportFilteredXlsx.disabled = isDisabled;
    if (this.btnExportFilteredCsv) this.btnExportFilteredCsv.disabled = isDisabled;
    if (this.btnExportFilteredJson) this.btnExportFilteredJson.disabled = isDisabled;
  }

  generateExportFilename(ext, category, dateFrom, dateTo) {
    let catPart = (category && category !== 'all') ? category.replace(/[^a-zA-Z0-9_-]/g, '_') : 'AllCategories';
    let datePart = '';
    if (dateFrom && dateTo) {
      datePart = `_${dateFrom}_to_${dateTo}`;
    } else if (dateFrom) {
      datePart = `_from_${dateFrom}`;
    } else if (dateTo) {
      datePart = `_until_${dateTo}`;
    } else {
      datePart = `_${new Date().toISOString().split('T')[0]}`;
    }
    return `ClassLogger_${catPart}${datePart}.${ext}`;
  }

  handleExportXlsx() {
    const { classes, category, dateFrom, dateTo } = this.getFilteredExportData();
    if (!classes.length) {
      this.showToast('No sessions found for this category/date filter', 'danger');
      return;
    }

    const coachName = (this.exportCoachName && this.exportCoachName.value.trim()) || 'Chew Jie Hern';
    const ratePerClass = this.exportRatePerClass ? (parseFloat(this.exportRatePerClass.value) || 40.0) : 40.0;
    const feesLabel = (this.exportFeesLabel && this.exportFeesLabel.value.trim()) || undefined;

    const blob = ExcelTimesheetExporter.toBlob(classes, {
      coachName,
      ratePerClass,
      feesLabel
    });
    const fileName = this.generateExportFilename('xlsx', category, dateFrom, dateTo);
    this.downloadBlob(blob, fileName);
    sound.playSuccess();
    this.showToast(`Exported ${classes.length} sessions to authentic timesheet Excel (.xlsx)!`, 'success');
  }

  handleExportCsv() {
    const { classes, category, dateFrom, dateTo } = this.getFilteredExportData();
    if (!classes.length) {
      this.showToast('No sessions found for this category/date filter', 'danger');
      return;
    }

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

    const csvContent = [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const fileName = this.generateExportFilename('csv', category, dateFrom, dateTo);
    this.downloadFile(csvContent, fileName, 'text/csv');
    sound.playSuccess();
    this.showToast(`Exported ${classes.length} sessions to CSV!`, 'success');
  }

  handleExportJson() {
    const { classes, category, dateFrom, dateTo } = this.getFilteredExportData();
    if (!classes.length) {
      this.showToast('No sessions found for this category/date filter', 'danger');
      return;
    }

    const payload = {
      appName: 'ClassLogger',
      version: 1,
      exportedAt: new Date().toISOString(),
      filter: { category, dateFrom: dateFrom || null, dateTo: dateTo || null },
      totalCount: classes.length,
      classes
    };

    const fileName = this.generateExportFilename('json', category, dateFrom, dateTo);
    this.downloadFile(JSON.stringify(payload, null, 2), fileName, 'application/json');
    sound.playSuccess();
    this.showToast(`Exported ${classes.length} sessions to JSON!`, 'success');
  }

  downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  async migrateLegacyImportCategories() {
    const legacyNames = ['lts', 'precomp', 'swim clinic'];
    let changed = false;
    for (const c of this.classes) {
      if (c.category && legacyNames.includes(c.category.toLowerCase())) {
        const oldCat = c.category;
        const newNote = c.note ? `${oldCat} - ${c.note}` : oldCat;
        await db.updateClass(c.id, {
          category: 'Aerosplash',
          categoryId: 'cat_aerosplash',
          note: newNote
        });
        changed = true;
      }
    }

    for (const cat of this.categories) {
      if (legacyNames.includes(cat.name.toLowerCase())) {
        await db.deleteCategory(cat.id);
        changed = true;
      }
    }

    if (changed) {
      await this.loadCategories();
      await this.loadClasses();
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
