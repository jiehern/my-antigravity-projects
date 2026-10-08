/**
 * app.js - Main Application Orchestrator for AeroSplash Class Logger
 * Integrates Apple Rotary Dial Picker, Localized IndexedDB Storage, Audio & Analytics
 */

import { db } from './db.js';
import { RotaryDialPicker } from './rotary-picker.js';
import { sound } from './audio.js';
import { ExcelTimesheetImporter, ExcelTimesheetExporter, getClassCreditUnits, calculateManagerBonus } from './excel-importer.js';

class ClassLoggerApp {
  constructor() {
    this.classes = [];
    this.selectedDuration = 60;
    this.editingClassId = null;
    this.searchQuery = '';

    // Initialize DOM hooks
    this.initElements();
  }

  async init() {
    console.log('AeroSplash Class Logger initializing...');

    // 1. Initialize Theme and Accent Color immediately
    this.initTheme();
    this.initAccentColor();

    // 2. Attach UI event listeners immediately so all buttons work right away
    this.attachEventListeners();
    this.updateSoundButtonUI();
    this.renderAssistNamesList();
    this.updateNoteChipsActiveState();
    this.updateBaseRateUI();

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

      await this.loadClasses();

      // Clean up / migrate any legacy sessions previously imported under subcategories
      await this.migrateLegacyImportCategories();
    } catch (err) {
      console.error('Error during DB init/data load:', err);
    }

    console.log('AeroSplash Class Logger ready!');
  }

  initElements() {
    // Rotary elements
    this.rotaryDateDisplay = document.getElementById('rotary-date-display');
    this.rotaryTimeDisplay = document.getElementById('rotary-time-display');

    // Form elements
    this.classForm = document.getElementById('class-form');
    this.durationContainer = document.getElementById('duration-pills-container');
    this.durationLabel = document.getElementById('selected-duration-label');
    this.customDurationWrapper = document.getElementById('custom-duration-wrapper');
    this.customDurationInput = document.getElementById('custom-duration-input');
    this.btnApplyCustomDuration = document.getElementById('btn-apply-custom-duration');
    this.notesInput = document.getElementById('class-notes-input');
    this.quickTagsContainer = document.getElementById('quick-note-tags');
    this.assistDropdownContainer = document.getElementById('assist-dropdown-container');
    this.btnAssistDropdownToggle = document.getElementById('btn-assist-dropdown-toggle');
    this.assistDropdownMenu = document.getElementById('assist-dropdown-menu');
    this.assistNamesList = document.getElementById('assist-names-list');
    this.inputNewAssistName = document.getElementById('input-new-assist-name');
    this.btnAddAssistName = document.getElementById('btn-add-assist-name');
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
    this.logCountTag = document.getElementById('log-count-tag');

    // Header buttons
    this.soundToggleBtn = document.getElementById('sound-toggle-btn');
    this.soundIcon = document.getElementById('sound-icon');
    this.themeToggleBtn = document.getElementById('theme-toggle-btn');
    this.themeIcon = document.getElementById('theme-icon');
    this.dataMenuBtn = document.getElementById('data-menu-btn');

    // Modals
    this.modalData = document.getElementById('modal-data');
    this.btnCloseDataModal = document.getElementById('btn-close-data-modal');
    this.btnExportCsv = document.getElementById('btn-export-csv');
    this.btnExportJson = document.getElementById('btn-export-json');
    this.importJsonInput = document.getElementById('import-json-input');
    this.importJsonHeaderBtn = document.getElementById('import-json-header-btn');
    this.btnSeedSample = document.getElementById('btn-seed-sample');
    this.btnClearDb = document.getElementById('btn-clear-db');

    // Appearance & Accent settings elements
    this.themeSwitchPills = document.getElementById('theme-switch-pills');
    this.accentPaletteGrid = document.getElementById('accent-palette-grid');
    this.accentCurrentPreview = document.getElementById('accent-current-preview');
    this.inputCustomAccentColor = document.getElementById('input-custom-accent-color');

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
    this.exportCalcBaseFees = document.getElementById('export-calc-base-fees');
    this.exportCalcBonus = document.getElementById('export-calc-bonus');
    this.exportCalcTotalFees = document.getElementById('export-calc-total-fees');

    this.btnExportFilteredXlsx = document.getElementById('btn-export-filtered-xlsx');
    this.btnExportFilteredCsv = document.getElementById('btn-export-filtered-csv');
    this.btnExportFilteredJson = document.getElementById('btn-export-filtered-json');

    this.toastContainer = document.getElementById('toast-container');
    this.monthlyBonusCard = document.getElementById('monthly-bonus-card');

    // Base rate setting elements
    this.inputCoachBaseRate = document.getElementById('input-coach-base-rate');
    this.btnSaveBaseRate = document.getElementById('btn-save-base-rate');
    this.settingsRateBadge = document.getElementById('settings-current-rate-badge');

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

  // --- THEME & ACCENT COLOR SYSTEM ---

  initTheme() {
    let savedTheme = 'dark';
    try {
      savedTheme = localStorage.getItem('aerosplash_theme') || 'dark';
    } catch (e) {
      console.warn('Could not read theme from localStorage:', e);
    }
    this.setTheme(savedTheme, false);

    // Watch for OS preference changes if system theme is selected
    if (window.matchMedia) {
      window.matchMedia('(prefers-color-scheme: light)').addEventListener('change', () => {
        let currentTheme = 'dark';
        try {
          currentTheme = localStorage.getItem('aerosplash_theme') || 'dark';
        } catch (e) {}
        if (currentTheme === 'system') {
          this.setTheme('system', false);
        }
      });
    }
  }

  getEffectiveTheme(theme) {
    if (theme === 'system') {
      return (window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches) ? 'light' : 'dark';
    }
    return theme === 'light' ? 'light' : 'dark';
  }

  setTheme(theme, save = true) {
    if (save) {
      try {
        localStorage.setItem('aerosplash_theme', theme);
      } catch (e) {
        console.warn('Could not save theme to localStorage:', e);
      }
    }

    const effectiveTheme = this.getEffectiveTheme(theme);
    document.documentElement.setAttribute('data-theme', effectiveTheme);

    // Update meta theme colors for Safari / Chrome mobile address bar
    const metaThemeColor = document.getElementById('meta-theme-color');
    if (metaThemeColor) {
      metaThemeColor.setAttribute('content', effectiveTheme === 'light' ? '#f2f2f7' : '#0a0a0c');
    }
    const metaColorScheme = document.getElementById('meta-color-scheme');
    if (metaColorScheme) {
      metaColorScheme.setAttribute('content', effectiveTheme);
    }

    // Update Header Theme Icon
    if (this.themeIcon) {
      this.themeIcon.textContent = effectiveTheme === 'light' ? '🌙' : '☀️';
    }
    if (this.themeToggleBtn) {
      this.themeToggleBtn.title = effectiveTheme === 'light' ? 'Switch to Dark Mode' : 'Switch to Light Mode';
    }

    // Update Settings Theme Switch Pills
    if (this.themeSwitchPills) {
      this.themeSwitchPills.querySelectorAll('.theme-pill').forEach(pill => {
        pill.classList.toggle('is-active', pill.dataset.themeVal === theme);
      });
    }
  }

  toggleTheme() {
    let currentTheme = 'dark';
    try {
      currentTheme = localStorage.getItem('aerosplash_theme') || 'dark';
    } catch (e) {}
    const effectiveTheme = this.getEffectiveTheme(currentTheme);
    const newTheme = effectiveTheme === 'dark' ? 'light' : 'dark';
    this.setTheme(newTheme, true);
    sound.playTick(1.2);
    this.showToast(`Switched to ${newTheme === 'dark' ? 'Dark' : 'Light'} Mode`, 'info');
  }

  initAccentColor() {
    let savedAccent = '#ff9f0a';
    try {
      savedAccent = localStorage.getItem('aerosplash_accent_color') || '#ff9f0a';
    } catch (e) {
      console.warn('Could not read accent color from localStorage:', e);
    }
    this.setAccentColor(savedAccent, false);
  }

  hexToRgb(hex) {
    let cleanHex = hex.replace('#', '').trim();
    if (cleanHex.length === 3) {
      cleanHex = cleanHex.split('').map(c => c + c).join('');
    }
    if (cleanHex.length !== 6) {
      return { r: 255, g: 159, b: 10 }; // Fallback orange
    }
    const num = parseInt(cleanHex, 16);
    return {
      r: (num >> 16) & 255,
      g: (num >> 8) & 255,
      b: num & 255
    };
  }

  setAccentColor(hex, save = true) {
    if (!hex || typeof hex !== 'string') return;
    let formattedHex = hex.trim();
    if (!formattedHex.startsWith('#')) formattedHex = '#' + formattedHex;

    const rgb = this.hexToRgb(formattedHex);
    // Perceived luminance formula (ITU-R BT.601)
    const luminance = (rgb.r * 299 + rgb.g * 587 + rgb.b * 114) / 1000;
    const contrastText = luminance > 140 ? '#000000' : '#ffffff';

    const rootStyle = document.documentElement.style;
    rootStyle.setProperty('--accent-color', formattedHex);
    rootStyle.setProperty('--accent-color-rgb', `${rgb.r}, ${rgb.g}, ${rgb.b}`);
    rootStyle.setProperty('--accent-glow', `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, 0.35)`);
    rootStyle.setProperty('--accent-contrast', contrastText);
    rootStyle.setProperty('--apple-orange', formattedHex);
    rootStyle.setProperty('--apple-orange-glow', `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, 0.35)`);
    rootStyle.setProperty('--border-focus', formattedHex);
    rootStyle.setProperty('--card-accent', formattedHex);

    if (save) {
      try {
        localStorage.setItem('aerosplash_accent_color', formattedHex);
      } catch (e) {
        console.warn('Could not save accent color to localStorage:', e);
      }
    }

    // Update UI elements
    if (this.accentCurrentPreview) {
      this.accentCurrentPreview.textContent = formattedHex.toUpperCase();
      this.accentCurrentPreview.style.color = formattedHex;
    }
    if (this.inputCustomAccentColor) {
      this.inputCustomAccentColor.value = formattedHex;
    }
    if (this.accentPaletteGrid) {
      this.accentPaletteGrid.querySelectorAll('.accent-color-btn').forEach(btn => {
        btn.classList.toggle('is-active', btn.dataset.accent.toLowerCase() === formattedHex.toLowerCase());
      });
    }
  }

  // --- DATA LOADING & RENDERING ---

  async loadClasses() {
    try {
      this.classes = await db.getAllClasses();
    } catch (e) {
      console.warn('Error loading classes:', e);
      this.classes = [];
    }
    this.renderClassesList();
  }

  renderClassesList() {
    if (!this.classesList) return;

    let filtered = [...this.classes];

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

    // Update the live Monthly Manager Bonus tracker card
    this.renderMonthlyBonusCard(filtered);

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
      // Actual formatted calendar date
      const itemDate = new Date(item.date + 'T00:00:00');
      const formattedDate = itemDate.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });
      const today = new Date();
      today.setHours(0,0,0,0);
      const isToday = (today.getTime() === itemDate.getTime());

      // Calculate units and display fee (LTS 50m = 1.0 unit; Pre Comp 90m = 1.5 units)
      const units = getClassCreditUnits(item);
      const isPreComp = (units === 1.5);
      const isLts = (units === 1.0 && (item.duration === 50 || (item.note && item.note.toLowerCase().includes('lts'))));
      const currentBaseRate = this.getBaseRate();
      const feeCalculated = (units * currentBaseRate).toFixed(0);
      const unitBadgeLabel = isPreComp ? `1.5 class • RM${feeCalculated}` : isLts ? `1.0 class • RM${feeCalculated}` : `${units.toFixed(1)} class • RM${feeCalculated}`;
      const unitBadgeClass = isPreComp ? 'is-precomp' : isLts ? 'is-lts' : '';

      const card = document.createElement('article');
      card.className = 'class-log-card';
      card.style.setProperty('--card-accent', 'var(--accent-color)');

      card.innerHTML = `
        <div class="card-top-row">
          <div class="card-category-badge">
            <span class="category-badge-icon">🌊</span>
            <span class="category-badge-text">AeroSplash</span>
          </div>
          <span class="card-units-badge ${unitBadgeClass}">
            ${unitBadgeLabel}
          </span>
        </div>

        <div class="card-main-row">
          <div class="card-time-block">
            <div class="card-time-display">
              <span class="card-time-large">${item.displayTime || item.time}</span>
              <span class="card-duration-badge">${item.duration || 60}m</span>
            </div>
            <div class="card-date-line">
              ${isToday ? '<span class="card-relative-badge is-today">Today</span>' : ''}
              <span class="card-date-text">${formattedDate}</span>
            </div>
          </div>

          <div class="card-actions">
            <button type="button" class="card-action-btn edit" title="Edit this entry" data-id="${item.id}" aria-label="Edit class">
              ✏️
            </button>
            <button type="button" class="card-action-btn duplicate" title="Duplicate entry" data-id="${item.id}" aria-label="Duplicate class">
              📋
            </button>
            <button type="button" class="card-action-btn delete" title="Delete entry" data-id="${item.id}" aria-label="Delete class">
              🗑️
            </button>
          </div>
        </div>

        ${item.note ? `
          <div class="card-note-box">
            <span class="card-note-icon">📝</span>
            <span class="card-note-text">${this.escapeHTML(item.note)}</span>
          </div>
        ` : ''}
      `;

      // Card action buttons
      card.querySelector('.edit').addEventListener('click', () => this.startEditClass(item));
      card.querySelector('.duplicate').addEventListener('click', () => this.duplicateClass(item));
      card.querySelector('.delete').addEventListener('click', () => this.deleteClass(item.id));

      this.classesList.appendChild(card);
    });
  }

  // --- MONTHLY MANAGER BONUS CARD ---

  renderMonthlyBonusCard(classesToSummarize = this.classes) {
    if (!this.monthlyBonusCard) return;

    const list = Array.isArray(classesToSummarize) ? classesToSummarize : this.classes;
    const totalSessions = list.length;
    const totalUnits = list.reduce((sum, c) => sum + getClassCreditUnits(c), 0);
    const ratePerClass = this.getBaseRate();
    const baseEarnings = totalUnits * ratePerClass;
    const bonusInfo = calculateManagerBonus(totalUnits);
    const totalPayout = baseEarnings + bonusInfo.bonus;

    let periodLabel = 'This Month';
    if (this.activeFilter === 'all' && !this.searchQuery) {
      if (list.length > 0 && list[0].date) {
        const [y, m] = list[0].date.split('-').map(Number);
        const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
        periodLabel = `${monthNames[m - 1] || 'Current Month'} ${y || 2026}`;
      } else {
        const now = new Date();
        periodLabel = now.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
      }
    } else if (this.searchQuery) {
      periodLabel = 'Search Results';
    }

    // Milestones progress calculation (scale 0 to 50 max)
    const progressPct = Math.min(100, Math.max(0, (totalUnits / 50) * 100));

    // Next goal description
    let nextGoalText = '';
    if (totalUnits >= 50) {
      nextGoalText = '🎉 Maximum Tier 4 Reached! (RM350 Bonus)';
    } else if (totalUnits >= 40) {
      nextGoalText = `🎯 ${bonusInfo.needed} more hrs/units to reach Tier 4 (RM350 Bonus)`;
    } else if (totalUnits >= 30) {
      nextGoalText = `🎯 ${bonusInfo.needed} more hrs/units to reach Tier 3 (RM250 Bonus)`;
    } else if (totalUnits >= 24) {
      nextGoalText = `🎯 ${bonusInfo.needed} more hrs/units to reach Tier 2 (RM150 Bonus)`;
    } else {
      nextGoalText = `🎯 ${bonusInfo.needed} more hrs/units to unlock Tier 1 (RM100 Bonus)`;
    }

    let tierBadgeHtml = '';
    if (bonusInfo.tier > 0) {
      tierBadgeHtml = `<span class="bonus-tier-badge tier-${bonusInfo.tier}">⭐ ${bonusInfo.tierName}: +RM${bonusInfo.bonus}</span>`;
    } else {
      tierBadgeHtml = `<span class="bonus-tier-badge no-tier">Below 24 hrs</span>`;
    }

    this.monthlyBonusCard.innerHTML = `
      <div class="bonus-card-top">
        <div class="bonus-card-title-group">
          <span class="bonus-trophy-icon">🏆</span>
          <div>
            <div class="bonus-card-title">Monthly Earnings & Manager Bonus</div>
            <div class="bonus-card-period">${this.escapeHTML(periodLabel)} • ${totalSessions} class${totalSessions === 1 ? '' : 'es'}</div>
          </div>
        </div>
        ${tierBadgeHtml}
      </div>

      <div class="bonus-stats-grid">
        <div class="bonus-stat-box">
          <span class="stat-box-label">Class Units</span>
          <strong class="stat-box-val stat-val-units">${totalUnits.toFixed(1)} <span class="stat-unit-sub">hrs</span></strong>
        </div>
        <div class="bonus-stat-box">
          <span class="stat-box-label">Base Rate (RM ${ratePerClass.toFixed(0)})</span>
          <strong class="stat-box-val stat-val-base">RM ${baseEarnings.toFixed(0)}</strong>
        </div>
        <div class="bonus-stat-box">
          <span class="stat-box-label">Manager Bonus</span>
          <strong class="stat-box-val stat-val-bonus">${bonusInfo.bonus > 0 ? '+RM ' + bonusInfo.bonus : 'RM 0'}</strong>
        </div>
        <div class="bonus-stat-box stat-box-total">
          <span class="stat-box-label">Total Payout</span>
          <strong class="stat-box-val stat-val-payout">RM ${totalPayout.toFixed(0)}</strong>
        </div>
      </div>

      <div class="bonus-progress-wrap">
        <div class="bonus-progress-track">
          <div class="bonus-progress-fill" style="width: ${progressPct}%;"></div>
          <!-- Milestone markers with concise hour badges -->
          <div class="bonus-milestone mark-24 ${totalUnits >= 24 ? 'is-active' : ''}" style="left: 48%;" title="24 hrs: RM100">
            <span class="milestone-tag">24h</span>
            <span class="milestone-tick"></span>
          </div>
          <div class="bonus-milestone mark-30 ${totalUnits >= 30 ? 'is-active' : ''}" style="left: 60%;" title="30 hrs: RM150">
            <span class="milestone-tag">30h</span>
            <span class="milestone-tick"></span>
          </div>
          <div class="bonus-milestone mark-40 ${totalUnits >= 40 ? 'is-active' : ''}" style="left: 80%;" title="40 hrs: RM250">
            <span class="milestone-tag">40h</span>
            <span class="milestone-tick"></span>
          </div>
          <div class="bonus-milestone mark-50 ${totalUnits >= 50 ? 'is-active' : ''}" style="left: 100%;" title="50 hrs: RM350">
            <span class="milestone-tag">50h</span>
            <span class="milestone-tick"></span>
          </div>
        </div>

        <!-- 4-Tier Milestone Cards (Spacious, beautifully aligned on mobile & desktop) -->
        <div class="bonus-tiers-grid">
          <div class="bonus-tier-pill ${totalUnits >= 24 ? 'is-achieved' : (totalUnits < 24 ? 'is-next' : '')}">
            <div class="tier-pill-hours-row">
              <span class="tier-pill-hours">24h</span>
              ${totalUnits >= 24 ? '<span class="tier-pill-check">✓</span>' : ''}
            </div>
            <span class="tier-pill-reward">+RM100</span>
          </div>
          <div class="bonus-tier-pill ${totalUnits >= 30 ? 'is-achieved' : (totalUnits >= 24 && totalUnits < 30 ? 'is-next' : '')}">
            <div class="tier-pill-hours-row">
              <span class="tier-pill-hours">30h</span>
              ${totalUnits >= 30 ? '<span class="tier-pill-check">✓</span>' : ''}
            </div>
            <span class="tier-pill-reward">+RM150</span>
          </div>
          <div class="bonus-tier-pill ${totalUnits >= 40 ? 'is-achieved' : (totalUnits >= 30 && totalUnits < 40 ? 'is-next' : '')}">
            <div class="tier-pill-hours-row">
              <span class="tier-pill-hours">40h</span>
              ${totalUnits >= 40 ? '<span class="tier-pill-check">✓</span>' : ''}
            </div>
            <span class="tier-pill-reward">+RM250</span>
          </div>
          <div class="bonus-tier-pill ${totalUnits >= 50 ? 'is-achieved' : (totalUnits >= 40 && totalUnits < 50 ? 'is-next' : '')}">
            <div class="tier-pill-hours-row">
              <span class="tier-pill-hours">50h</span>
              ${totalUnits >= 50 ? '<span class="tier-pill-check">✓</span>' : ''}
            </div>
            <span class="tier-pill-reward">+RM350</span>
          </div>
        </div>

        <div class="bonus-footer-row">
          <span class="bonus-next-goal">${nextGoalText}</span>
          <span class="bonus-unit-formula">LTS (50m) = 1.0 unit (RM${ratePerClass.toFixed(0)}) • Pre Comp (90m) = 1.5 units (RM${(ratePerClass * 1.5).toFixed(0)})</span>
        </div>
      </div>
    `;
  }

  // --- DURATION HELPER ---

  setDuration(mins) {
    const parsed = parseInt(mins, 10);
    if (!parsed || parsed <= 0) return;
    this.selectedDuration = parsed;

    if (this.durationLabel) {
      this.durationLabel.textContent = `${parsed} mins`;
    }

    if (this.durationContainer) {
      let matched = false;
      this.durationContainer.querySelectorAll('.duration-pill').forEach(pill => {
        const isMatch = parseInt(pill.dataset.mins, 10) === parsed;
        pill.classList.toggle('is-active', isMatch);
        if (isMatch) matched = true;
      });

      const customPill = this.durationContainer.querySelector('[data-mins="custom"]');
      if (!matched) {
        if (customPill) customPill.classList.add('is-active');
        if (this.customDurationWrapper) this.customDurationWrapper.style.display = 'flex';
        if (this.customDurationInput) this.customDurationInput.value = parsed;
      } else {
        if (customPill) customPill.classList.remove('is-active');
        if (this.customDurationWrapper) this.customDurationWrapper.style.display = 'none';
      }
    }
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

    // 2. Note Options (LTS, Pre Comp, Assist dropdown)
    if (this.quickTagsContainer) {
      // Direct buttons for LTS and Pre Comp
      this.quickTagsContainer.querySelectorAll('.note-tag-chip[data-note-val]').forEach(chip => {
        chip.addEventListener('click', () => {
          const val = chip.dataset.noteVal;
          this.setSessionNote(val);
          sound.playTick(1.2);
        });
      });
    }

    // Assist dropdown toggle
    if (this.btnAssistDropdownToggle) {
      this.btnAssistDropdownToggle.addEventListener('click', (e) => {
        e.stopPropagation();
        this.toggleAssistDropdown();
        sound.playTick(1.1);
      });
    }

    // Assist dropdown list clicks (Select / Delete)
    if (this.assistNamesList) {
      this.assistNamesList.addEventListener('click', (e) => {
        const delBtn = e.target.closest('.assist-name-delete-btn');
        if (delBtn) {
          e.stopPropagation();
          const idx = parseInt(delBtn.dataset.deleteIdx, 10);
          this.handleDeleteAssistName(idx);
          return;
        }

        const selBtn = e.target.closest('.assist-name-select-btn');
        if (selBtn) {
          e.stopPropagation();
          const val = selBtn.dataset.assistVal;
          this.setSessionNote(val);
          this.closeAssistDropdown();
          sound.playTick(1.2);
        }
      });
    }

    // Add new assist name
    if (this.btnAddAssistName) {
      this.btnAddAssistName.addEventListener('click', (e) => {
        e.stopPropagation();
        this.handleAddAssistName();
      });
    }

    if (this.inputNewAssistName) {
      this.inputNewAssistName.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          e.stopPropagation();
          this.handleAddAssistName();
        }
      });
      this.inputNewAssistName.addEventListener('click', (e) => e.stopPropagation());
    }

    // Close assist dropdown on outside click
    document.addEventListener('click', (e) => {
      if (this.assistDropdownContainer && !this.assistDropdownContainer.contains(e.target)) {
        this.closeAssistDropdown();
      }
    });

    // Close assist dropdown on Escape
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        this.closeAssistDropdown();
      }
    });

    // Track user edits in note textarea to dynamically update chip active styling
    if (this.notesInput) {
      this.notesInput.addEventListener('input', () => {
        this.updateNoteChipsActiveState();
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

    // 6. Sound & Theme Controls
    if (this.soundToggleBtn) {
      this.soundToggleBtn.addEventListener('click', () => {
        const isMuted = sound.toggleMute();
        this.updateSoundButtonUI();
        if (!isMuted) sound.playTick(1.0);
        this.showToast(isMuted ? 'Muted mechanical sound' : 'Sound feedback enabled', 'info');
      });
    }

    if (this.themeToggleBtn) {
      this.themeToggleBtn.addEventListener('click', () => {
        this.toggleTheme();
      });
    }

    // Appearance & Accent Settings Listeners
    if (this.themeSwitchPills) {
      this.themeSwitchPills.querySelectorAll('.theme-pill').forEach(pill => {
        pill.addEventListener('click', () => {
          this.setTheme(pill.dataset.themeVal);
        });
      });
    }

    if (this.accentPaletteGrid) {
      this.accentPaletteGrid.querySelectorAll('.accent-color-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          this.setAccentColor(btn.dataset.accent);
          sound.playTick(1.1);
        });
      });
    }

    if (this.inputCustomAccentColor) {
      this.inputCustomAccentColor.addEventListener('input', (e) => {
        this.setAccentColor(e.target.value);
      });
    }

    // 7. Modal Open & Close Triggers
    if (this.dataMenuBtn) {
      this.dataMenuBtn.addEventListener('click', () => this.openModal(this.modalData));
    }
    if (this.btnCloseDataModal) {
      this.btnCloseDataModal.addEventListener('click', () => this.closeModal(this.modalData));
    }
    if (this.btnCloseDateModal) {
      this.btnCloseDateModal.addEventListener('click', () => this.closeModal(this.modalDateJump));
    }

    // Coach Base Rate Settings Triggers
    if (this.btnSaveBaseRate) {
      this.btnSaveBaseRate.addEventListener('click', () => {
        if (this.inputCoachBaseRate) {
          this.setBaseRate(this.inputCoachBaseRate.value);
        }
      });
    }
    if (this.inputCoachBaseRate) {
      this.inputCoachBaseRate.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          this.setBaseRate(this.inputCoachBaseRate.value);
        }
      });
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
    [this.modalData, this.modalDateJump, this.modalExport].forEach(modal => {
      if (modal) {
        modal.addEventListener('click', (e) => {
          if (e.target === modal) this.closeModal(modal);
        });
      }
    });

    // Export modal date filters & buttons
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
      this.exportRatePerClass.addEventListener('input', () => {
        this.exportRatePerClass.dataset.userEdited = 'true';
        this.updateExportPreview();
      });
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

    // JSON Import Listeners
    const jsonChangeHandler = (e) => {
      const file = e.target.files[0];
      if (file) this.handleJsonImport(file);
      e.target.value = '';
    };

    if (this.importJsonInput) {
      this.importJsonInput.addEventListener('change', jsonChangeHandler);
    }
    if (this.importJsonHeaderBtn) {
      this.importJsonHeaderBtn.addEventListener('change', jsonChangeHandler);
    }

    // Drag & Drop File Import (JSON & XLSX)
    window.addEventListener('dragover', (e) => {
      e.preventDefault();
      e.stopPropagation();
    });

    window.addEventListener('drop', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const files = e.dataTransfer?.files;
      if (!files || files.length === 0) return;
      const file = files[0];
      const lower = file.name.toLowerCase();
      if (lower.endsWith('.json')) {
        this.handleJsonImport(file);
      } else if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
        this.handleExcelImport(file);
      }
    });

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

  // --- BASE RATE SETTINGS ---

  getBaseRate() {
    try {
      const stored = localStorage.getItem('class_logger_base_rate');
      if (stored !== null && !isNaN(parseFloat(stored))) {
        const val = parseFloat(stored);
        if (val >= 0) return val;
      }
    } catch (e) {
      console.warn('Error reading base rate from localStorage:', e);
    }
    return 40.0;
  }

  setBaseRate(newRate) {
    const val = parseFloat(newRate);
    if (isNaN(val) || val < 0) {
      this.showToast('Please enter a valid base rate (RM >= 0)', 'danger');
      return false;
    }
    try {
      localStorage.setItem('class_logger_base_rate', val.toString());
    } catch (e) {
      console.warn('Error saving base rate to localStorage:', e);
    }

    this.updateBaseRateUI();
    this.renderClassesList();
    if (this.exportRatePerClass) {
      this.exportRatePerClass.value = val;
      delete this.exportRatePerClass.dataset.userEdited;
    }
    this.updateExportPreview();
    sound.playSuccess();
    this.showToast(`Coach base rate updated to RM ${val.toFixed(1)} / class!`, 'success');
    return true;
  }

  updateBaseRateUI() {
    const rate = this.getBaseRate();
    if (this.inputCoachBaseRate) {
      this.inputCoachBaseRate.value = rate;
    }
    if (this.settingsRateBadge) {
      this.settingsRateBadge.textContent = `RM ${rate.toFixed(1)} / class`;
    }
    if (this.exportRatePerClass && !this.exportRatePerClass.dataset.userEdited) {
      this.exportRatePerClass.value = rate;
    }
  }

  // --- CRUD ACTIONS ---

  async handleSubmitClass() {
    if (!this.rotaryPicker) return;

    const rotaryVal = this.rotaryPicker.getValue();
    const note = this.notesInput ? this.notesInput.value.trim() : '';

    const recordData = {
      date: rotaryVal.date,
      time: rotaryVal.time24,
      displayTime: rotaryVal.displayTime,
      timestamp: rotaryVal.timestamp,
      category: 'AeroSplash',
      duration: this.selectedDuration,
      note: note
    };

    if (this.editingClassId) {
      // Update existing record
      await db.updateClass(this.editingClassId, recordData);
      this.showToast('Updated AeroSplash session', 'success');
      sound.playSuccess();
      this.cancelEdit();
    } else {
      // Create new record
      await db.addClass(recordData);
      this.showToast(`Logged AeroSplash session at ${rotaryVal.displayTime}`, 'success');
      sound.playSuccess();

      // Reset note field
      if (this.notesInput) {
        this.notesInput.value = '';
        this.updateNoteChipsActiveState();
      }
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
      this.updateNoteChipsActiveState();
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
    if (this.notesInput) {
      this.notesInput.value = '';
      this.updateNoteChipsActiveState();
    }
    if (this.customDurationWrapper) this.customDurationWrapper.style.display = 'none';
  }

  async duplicateClass(item) {
    const clone = {
      ...item,
      id: undefined,
      category: 'AeroSplash',
      note: item.note ? `${item.note} (Copy)` : ''
    };
    await db.addClass(clone);
    await this.loadClasses();
    sound.playSuccess();
    this.showToast('Duplicated AeroSplash session', 'success');
  }

  async deleteClass(id) {
    if (confirm('Delete this class session?')) {
      await db.deleteClass(id);
      sound.playTrash();
      await this.loadClasses();
      this.showToast('Class entry deleted', 'info');
    }
  }

  // --- MODALS & UTILS ---

  openModal(modalEl) {
    if (!modalEl) return;
    modalEl.classList.add('is-open');
    if (modalEl === this.modalData) {
      this.updateBaseRateUI();
    }
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

  // --- SESSION NOTE & ASSIST DROPDOWN ---

  getAssistNames() {
    try {
      const saved = localStorage.getItem('class_logger_assist_names');
      if (saved !== null) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch (e) {
      console.warn('Failed to parse assist names', e);
    }
    return ['Coach Allen', 'Charles'];
  }

  saveAssistNames(names) {
    try {
      localStorage.setItem('class_logger_assist_names', JSON.stringify(names));
    } catch (e) {
      console.warn('Failed to save assist names', e);
    }
  }

  renderAssistNamesList() {
    if (!this.assistNamesList) return;
    const names = this.getAssistNames();
    const currentVal = this.notesInput ? this.notesInput.value.trim() : '';

    const isGenericActive = currentVal === 'Assist' || currentVal.startsWith('Assist -');

    let html = `
      <div class="assist-name-row generic-assist-row ${isGenericActive ? 'is-selected' : ''}">
        <button type="button" class="assist-name-select-btn" data-assist-val="Assist">
          <span class="assist-name-icon">👤</span>
          <span class="assist-name-label">Assist (No Name)</span>
          ${isGenericActive ? '<span class="assist-name-check">✓</span>' : ''}
        </button>
      </div>
    `;

    if (names.length === 0) {
      html += `<div class="assist-no-names">No custom names saved yet. Add a name below.</div>`;
    } else {
      names.forEach((name, idx) => {
        const isNameActive = currentVal === `Assist ${name}` || currentVal.startsWith(`Assist ${name} -`) || currentVal.startsWith(`Assist ${name} `);
        html += `
          <div class="assist-name-row ${isNameActive ? 'is-selected' : ''}">
            <button type="button" class="assist-name-select-btn" data-assist-val="Assist ${this.escapeHTML(name)}">
              <span class="assist-name-icon">🏊</span>
              <span class="assist-name-label">${this.escapeHTML(name)}</span>
              ${isNameActive ? '<span class="assist-name-check">✓</span>' : ''}
            </button>
            <button type="button" class="assist-name-delete-btn" data-delete-idx="${idx}" title="Delete ${this.escapeHTML(name)}" aria-label="Delete ${this.escapeHTML(name)}">
              ✕
            </button>
          </div>
        `;
      });
    }

    this.assistNamesList.innerHTML = html;
  }

  toggleAssistDropdown(forceState) {
    if (!this.assistDropdownMenu || !this.assistDropdownContainer) return;
    const isCurrentlyOpen = this.assistDropdownContainer.classList.contains('is-open');
    const shouldOpen = forceState !== undefined ? forceState : !isCurrentlyOpen;

    if (shouldOpen) {
      this.renderAssistNamesList();
      this.assistDropdownContainer.classList.add('is-open');
      this.assistDropdownMenu.style.display = 'flex';
      if (this.btnAssistDropdownToggle) {
        this.btnAssistDropdownToggle.setAttribute('aria-expanded', 'true');
      }
      if (window.innerWidth > 600 && this.inputNewAssistName) {
        setTimeout(() => this.inputNewAssistName.focus(), 50);
      }
    } else {
      this.closeAssistDropdown();
    }
  }

  closeAssistDropdown() {
    if (!this.assistDropdownContainer || !this.assistDropdownMenu) return;
    this.assistDropdownContainer.classList.remove('is-open');
    this.assistDropdownMenu.style.display = 'none';
    if (this.btnAssistDropdownToggle) {
      this.btnAssistDropdownToggle.setAttribute('aria-expanded', 'false');
    }
  }

  setSessionNote(newType) {
    if (!this.notesInput) return;
    const current = this.notesInput.value.trim();

    if (!current) {
      this.notesInput.value = newType;
    } else {
      const isKnownType = current === 'LTS' || current.startsWith('LTS ') || current.startsWith('LTS -') ||
                          current === 'Pre Comp' || current.startsWith('Pre Comp ') || current.startsWith('Pre Comp -') ||
                          current === 'Assist' || current.startsWith('Assist ') || current.startsWith('Assist -');

      if (isKnownType) {
        if (current.includes(' - ')) {
          const remarks = current.split(' - ').slice(1).join(' - ').trim();
          this.notesInput.value = remarks ? `${newType} - ${remarks}` : newType;
        } else {
          this.notesInput.value = newType;
        }
      } else {
        this.notesInput.value = `${newType} - ${current}`;
      }
    }

    // Auto-align duration based on class note: LTS -> 50m; Pre Comp -> 90m (1h30m)
    if (newType === 'LTS' || newType.startsWith('LTS ') || newType.startsWith('LTS -')) {
      this.setDuration(50);
    } else if (newType === 'Pre Comp' || newType.startsWith('Pre Comp ') || newType.startsWith('Pre Comp -')) {
      this.setDuration(90);
    }

    this.updateNoteChipsActiveState();
    this.renderAssistNamesList();
    this.notesInput.focus();
  }

  updateNoteChipsActiveState() {
    if (!this.notesInput) return;
    const val = (this.notesInput.value || '').trim();

    const ltsBtn = document.querySelector('.note-tag-chip[data-note-val="LTS"]');
    const preCompBtn = document.querySelector('.note-tag-chip[data-note-val="Pre Comp"]');
    const assistBtn = document.getElementById('btn-assist-dropdown-toggle');

    if (ltsBtn) {
      const isLTS = val === 'LTS' || val.startsWith('LTS ') || val.startsWith('LTS -');
      ltsBtn.classList.toggle('is-active', isLTS);
    }

    if (preCompBtn) {
      const isPreComp = val === 'Pre Comp' || val.startsWith('Pre Comp ') || val.startsWith('Pre Comp -');
      preCompBtn.classList.toggle('is-active', isPreComp);
    }

    if (assistBtn) {
      const isAssist = val === 'Assist' || val.startsWith('Assist ') || val.startsWith('Assist -');
      assistBtn.classList.toggle('is-active', isAssist);
    }
  }

  handleAddAssistName() {
    if (!this.inputNewAssistName) return;
    let rawName = this.inputNewAssistName.value.trim();
    if (!rawName) return;

    if (rawName.toLowerCase().startsWith('assist ')) {
      rawName = rawName.slice(7).trim();
    }

    if (!rawName) return;

    const names = this.getAssistNames();
    const exists = names.some(n => n.toLowerCase() === rawName.toLowerCase());
    if (!exists) {
      names.push(rawName);
      this.saveAssistNames(names);
    }

    this.setSessionNote(`Assist ${rawName}`);
    this.inputNewAssistName.value = '';
    this.closeAssistDropdown();
    sound.playTick(1.3);
    this.showToast(`Selected Assist: ${rawName}`, 'success');
  }

  handleDeleteAssistName(idx) {
    const names = this.getAssistNames();
    if (idx >= 0 && idx < names.length) {
      const deleted = names.splice(idx, 1)[0];
      this.saveAssistNames(names);
      this.renderAssistNamesList();
      this.showToast(`Removed "${deleted}" from saved list`, 'info');
      sound.playTick(0.9);
    }
  }

  // --- JSON IMPORT ---

  async handleJsonImport(file) {
    if (!file) return;
    this.showToast(`Reading "${file.name}"...`, 'info');
    try {
      const text = await file.text();
      const res = await db.importJSON(text);
      await this.loadCategories();
      await this.loadClasses();
      this.closeModal(this.modalData);
      sound.playSuccess();
      this.showToast(`Successfully imported ${res.count} class session${res.count === 1 ? '' : 's'} from JSON!`, 'success');
    } catch (err) {
      console.error('JSON import error:', err);
      sound.playTrash();
      this.showToast(`Import failed: ${err.message || 'Invalid JSON format'}`, 'danger');
    }
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

      // Add each class as an AeroSplash session
      let imported = 0;
      for (const cls of result.classes) {
        cls.category = 'AeroSplash';
        await db.addClass(cls);
        imported++;
      }

      await this.loadClasses();
      this.closeModal(this.modalData);
      sound.playTick(1.5);
      this.showToast(`Imported ${imported} class sessions as AeroSplash!`, 'success');
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
    if (this.exportRatePerClass && !this.exportRatePerClass.dataset.userEdited) {
      this.exportRatePerClass.value = this.getBaseRate();
    }
    this.updateExportPreview();
    this.openModal(this.modalExport);
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
    const dateFrom = this.exportDateFrom ? this.exportDateFrom.value.trim() : '';
    const dateTo = this.exportDateTo ? this.exportDateTo.value.trim() : '';

    const filtered = this.classes.filter(c => {
      if (dateFrom && c.date < dateFrom) return false;
      if (dateTo && c.date > dateTo) return false;
      return true;
    });

    return {
      classes: filtered,
      dateFrom,
      dateTo
    };
  }

  updateExportPreview() {
    const { classes, dateFrom, dateTo } = this.getFilteredExportData();
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
      let dateLabel = 'All recorded dates';
      if (dateFrom && dateTo) {
        dateLabel = `${dateFrom} to ${dateTo}`;
      } else if (dateFrom) {
        dateLabel = `From ${dateFrom} onwards`;
      } else if (dateTo) {
        dateLabel = `Up to ${dateTo}`;
      }
      this.exportPreviewRange.textContent = dateLabel;
    }

    // Live Billing & Fee Calculations with Units and Manager Bonus
    const totalUnits = classes.reduce((sum, c) => sum + getClassCreditUnits(c), 0);
    const baseRateFallback = this.getBaseRate();
    const rateVal = this.exportRatePerClass ? (parseFloat(this.exportRatePerClass.value) || baseRateFallback) : baseRateFallback;
    const baseFees = totalUnits * rateVal;
    const bonusInfo = calculateManagerBonus(totalUnits);
    const grandTotal = baseFees + bonusInfo.bonus;

    if (this.exportCalcTotalClasses) {
      this.exportCalcTotalClasses.textContent = `${totalUnits.toFixed(1)} units (${count} class${count === 1 ? '' : 'es'})`;
    }
    if (this.exportCalcRate) {
      this.exportCalcRate.textContent = `RM ${rateVal.toFixed(1)}`;
    }
    if (this.exportCalcBaseFees) {
      this.exportCalcBaseFees.textContent = `RM ${baseFees.toFixed(2)}`;
    }
    if (this.exportCalcBonus) {
      this.exportCalcBonus.textContent = bonusInfo.bonus > 0 ? `+RM ${bonusInfo.bonus.toFixed(2)} (${bonusInfo.tierName})` : `RM 0.00`;
    }
    if (this.exportCalcTotalFees) {
      this.exportCalcTotalFees.textContent = `RM ${grandTotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
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

  generateExportFilename(ext, dateFrom, dateTo) {
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
    return `AeroSplash_Classes${datePart}.${ext}`;
  }

  handleExportXlsx() {
    const { classes, dateFrom, dateTo } = this.getFilteredExportData();
    if (!classes.length) {
      this.showToast('No sessions found for this date filter', 'danger');
      return;
    }

    const coachName = (this.exportCoachName && this.exportCoachName.value.trim()) || 'Chew Jie Hern';
    const ratePerClass = this.exportRatePerClass ? (parseFloat(this.exportRatePerClass.value) || this.getBaseRate()) : this.getBaseRate();
    const feesLabel = (this.exportFeesLabel && this.exportFeesLabel.value.trim()) || undefined;

    const blob = ExcelTimesheetExporter.toBlob(classes, {
      coachName,
      ratePerClass,
      feesLabel
    });
    const fileName = this.generateExportFilename('xlsx', dateFrom, dateTo);
    this.downloadBlob(blob, fileName);
    sound.playSuccess();
    this.showToast(`Exported ${classes.length} sessions to authentic timesheet Excel (.xlsx)!`, 'success');
  }

  handleExportCsv() {
    const { classes, dateFrom, dateTo } = this.getFilteredExportData();
    if (!classes.length) {
      this.showToast('No sessions found for this date filter', 'danger');
      return;
    }

    const headers = ['Date', 'Day', 'Time', 'Duration (mins)', 'Duration (hrs)', 'Credit Units', 'Fee (RM)', 'Notes / Remarks', 'Session ID'];
    const rateVal = this.exportRatePerClass ? (parseFloat(this.exportRatePerClass.value) || this.getBaseRate()) : this.getBaseRate();
    const rows = classes.map(c => {
      const d = new Date(c.date + 'T00:00:00');
      const dayOfWeek = isNaN(d.getTime()) ? '' : d.toLocaleDateString('en-US', { weekday: 'short' });
      const hours = c.duration ? (c.duration / 60).toFixed(2) : '1.00';
      const units = getClassCreditUnits(c);
      const fee = (units * rateVal).toFixed(2);
      return [
        `"${c.date}"`,
        `"${dayOfWeek}"`,
        `"${c.displayTime || c.time}"`,
        c.duration || 60,
        hours,
        units.toFixed(1),
        fee,
        `"${(c.note || '').replace(/"/g, '""')}"`,
        `"${c.id}"`
      ];
    });

    const csvContent = [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const fileName = this.generateExportFilename('csv', dateFrom, dateTo);
    this.downloadFile(csvContent, fileName, 'text/csv');
    sound.playSuccess();
    this.showToast(`Exported ${classes.length} sessions to CSV!`, 'success');
  }

  handleExportJson() {
    const { classes, dateFrom, dateTo } = this.getFilteredExportData();
    if (!classes.length) {
      this.showToast('No sessions found for this date filter', 'danger');
      return;
    }

    const payload = {
      appName: 'AeroSplash Class Logger',
      version: 1,
      exportedAt: new Date().toISOString(),
      filter: { dateFrom: dateFrom || null, dateTo: dateTo || null },
      totalCount: classes.length,
      classes
    };

    const fileName = this.generateExportFilename('json', dateFrom, dateTo);
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
          category: 'AeroSplash',
          note: newNote
        });
        changed = true;
      }
    }

    if (changed) {
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
