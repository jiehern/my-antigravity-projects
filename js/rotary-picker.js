/**
 * rotary-picker.js - Apple Clock Rotary Dial / Drum Wheel Picker
 * Replicates the authentic iOS 3D cylindrical drum picker with physics, momentum,
 * 3D perspective curvature, selection lens, and mechanical audio feedback.
 */

import { sound } from './audio.js';

export class RotaryDialPicker {
  constructor(containerEl, options = {}) {
    this.container = containerEl;
    this.options = {
      itemHeight: 44, // Apple HIG standard row height
      visibleItems: 5,
      perspective: 1000,
      onChange: options.onChange || (() => {}),
      initialDate: options.initialDate || new Date(),
      ...options
    };

    this.columns = {};
    this.currentValue = {
      date: '',
      hour: 12,
      minute: 0,
      period: 'AM'
    };

    this.init();
  }

  init() {
    this.container.classList.add('apple-rotary-picker');
    this.container.innerHTML = `
      <div class="rotary-dial-lens" aria-hidden="true"></div>
      <div class="rotary-dial-vignette" aria-hidden="true"></div>
      <div class="rotary-columns-wrapper"></div>
    `;

    this.wrapper = this.container.querySelector('.rotary-columns-wrapper');

    // Build the 4 columns
    this.setupDateColumn();
    this.setupHourColumn();
    this.setupMinuteColumn();
    this.setupPeriodColumn();

    // Set initial values to requested date
    this.setValue(this.options.initialDate, false);
  }

  // --- COLUMN CREATION HELPERS ---

  setupDateColumn() {
    // Generate dates: 60 days in past, 14 days in future
    const dates = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    for (let i = -60; i <= 14; i++) {
      const d = new Date(today);
      d.setDate(today.getDate() + i);

      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const dd = String(d.getDate()).padStart(2, '0');
      const iso = `${yyyy}-${mm}-${dd}`;

      const weekday = d.toLocaleDateString('en-US', { weekday: 'short' });
      const month = d.toLocaleDateString('en-US', { month: 'short' });
      const dayNum = d.getDate();
      const formattedDate = `${weekday}, ${month} ${dayNum}`;

      // Today shows the actual date with a clear marker, yesterday and tomorrow show actual dates
      let label = formattedDate;
      if (i === 0) {
        label = `Today, ${month} ${dayNum}`;
      }

      dates.push({
        value: iso,
        label: label,
        dateObj: d,
        isToday: i === 0
      });
    }

    this.columns.date = this.createDrumColumn('date', dates, false);
  }

  setupHourColumn() {
    // 1 to 12 repeated for smooth circular feeling
    const rawHours = [];
    for (let h = 1; h <= 12; h++) {
      rawHours.push({
        value: h,
        label: String(h).padStart(2, '0')
      });
    }
    // Repeat 5 times for virtually infinite rotary scroll
    const items = [];
    for (let r = 0; r < 5; r++) {
      rawHours.forEach(h => items.push({ ...h }));
    }

    this.columns.hour = this.createDrumColumn('hour', items, true, 12);
  }

  setupMinuteColumn() {
    // 00 to 59
    const rawMins = [];
    for (let m = 0; m < 60; m++) {
      rawMins.push({
        value: m,
        label: String(m).padStart(2, '0')
      });
    }
    // Repeat 3 times for virtually infinite rotary scroll
    const items = [];
    for (let r = 0; r < 3; r++) {
      rawMins.forEach(m => items.push({ ...m }));
    }

    this.columns.minute = this.createDrumColumn('minute', items, true, 60);
  }

  setupPeriodColumn() {
    const items = [
      { value: 'AM', label: 'AM' },
      { value: 'PM', label: 'PM' }
    ];
    this.columns.period = this.createDrumColumn('period', items, false);
  }

  // --- CORE DRUM WHEEL COLUMN ---

  createDrumColumn(columnId, items, isLooping = false, baseCount = 0) {
    const colEl = document.createElement('div');
    colEl.className = `rotary-column rotary-column-${columnId}`;
    colEl.setAttribute('tabindex', '0');
    colEl.setAttribute('role', 'spinbutton');
    colEl.setAttribute('aria-label', `Select ${columnId}`);

    const wheelEl = document.createElement('div');
    wheelEl.className = 'rotary-wheel';

    items.forEach((item, index) => {
      const itemEl = document.createElement('div');
      itemEl.className = 'rotary-item';
      itemEl.dataset.index = index;
      itemEl.textContent = item.label;
      if (item.isToday) itemEl.classList.add('is-today');
      wheelEl.appendChild(itemEl);
    });

    colEl.appendChild(wheelEl);
    this.wrapper.appendChild(colEl);

    const state = {
      id: columnId,
      colEl,
      wheelEl,
      items,
      itemElements: Array.from(wheelEl.children),
      isLooping,
      baseCount,
      currentIndex: 0,
      targetY: 0,
      currentY: 0,
      itemHeight: this.options.itemHeight,
      isDragging: false,
      startY: 0,
      startScrollY: 0,
      lastDragY: 0,
      lastDragTime: 0,
      velocity: 0,
      animationFrame: null
    };

    this.attachColumnEvents(state);
    return state;
  }

  attachColumnEvents(state) {
    const { colEl, wheelEl } = state;

    const onPointerDown = (clientY) => {
      if (state.animationFrame) cancelAnimationFrame(state.animationFrame);
      state.isDragging = true;
      state.startY = clientY;
      state.startScrollY = state.currentY;
      state.lastDragY = clientY;
      state.lastDragTime = Date.now();
      state.velocity = 0;
      colEl.classList.add('is-dragging');
    };

    const onPointerMove = (clientY) => {
      if (!state.isDragging) return;
      const now = Date.now();
      const deltaY = clientY - state.startY;
      const moveDelta = clientY - state.lastDragY;
      const timeDelta = now - state.lastDragTime;

      if (timeDelta > 0) {
        state.velocity = moveDelta / timeDelta;
      }
      state.lastDragY = clientY;
      state.lastDragTime = now;

      state.currentY = state.startScrollY - deltaY;
      this.renderColumnTransform(state);
    };

    const onPointerUp = () => {
      if (!state.isDragging) return;
      state.isDragging = false;
      colEl.classList.remove('is-dragging');

      // Apply momentum flick physics
      let momentumDist = state.velocity * 180;
      let projectedY = state.currentY - momentumDist;
      let targetIndex = Math.round(projectedY / state.itemHeight);

      this.snapToIndex(state, targetIndex, true);
    };

    // Touch handlers
    colEl.addEventListener('touchstart', (e) => {
      if (e.touches.length === 1) onPointerDown(e.touches[0].clientY);
    }, { passive: true });

    colEl.addEventListener('touchmove', (e) => {
      if (e.touches.length === 1) {
        if (e.cancelable) e.preventDefault();
        onPointerMove(e.touches[0].clientY);
      }
    }, { passive: false });

    colEl.addEventListener('touchend', () => onPointerUp(), { passive: true });
    colEl.addEventListener('touchcancel', () => onPointerUp(), { passive: true });

    // Mouse handlers
    colEl.addEventListener('mousedown', (e) => {
      if (e.button === 0) {
        e.preventDefault();
        onPointerDown(e.clientY);
        const onDocMouseMove = (ev) => onPointerMove(ev.clientY);
        const onDocMouseUp = () => {
          document.removeEventListener('mousemove', onDocMouseMove);
          document.removeEventListener('mouseup', onDocMouseUp);
          onPointerUp();
        };
        document.addEventListener('mousemove', onDocMouseMove);
        document.addEventListener('mouseup', onDocMouseUp);
      }
    });

    // Mouse Wheel / Trackpad
    let wheelTimeout = null;
    colEl.addEventListener('wheel', (e) => {
      e.preventDefault();
      if (state.animationFrame) cancelAnimationFrame(state.animationFrame);
      state.currentY += e.deltaY * 0.45;
      this.renderColumnTransform(state);

      clearTimeout(wheelTimeout);
      wheelTimeout = setTimeout(() => {
        const targetIndex = Math.round(state.currentY / state.itemHeight);
        this.snapToIndex(state, targetIndex, true);
      }, 70);
    }, { passive: false });

    // Item Click to Spin
    colEl.addEventListener('click', (e) => {
      const itemEl = e.target.closest('.rotary-item');
      if (itemEl && !state.isDragging) {
        const clickedIdx = parseInt(itemEl.dataset.index, 10);
        this.snapToIndex(state, clickedIdx, true);
      }
    });

    // Keyboard Arrow Keys
    colEl.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        this.snapToIndex(state, state.currentIndex - 1, true);
      } else if (e.key === 'ArrowDown') {
        e.preventDefault();
        this.snapToIndex(state, state.currentIndex + 1, true);
      }
    });
  }

  // --- RENDERING 3D CYLINDER PHYSICS ---

  renderColumnTransform(state) {
    const rawIndex = state.currentY / state.itemHeight;
    const roundedIndex = Math.round(rawIndex);

    // Play tick sound when passing a slot threshold
    if (roundedIndex !== state.lastSoundIndex) {
      sound.playTick(state.id === 'minute' ? 1.05 : 1.0);
      state.lastSoundIndex = roundedIndex;
    }

    const halfVisible = Math.floor(this.options.visibleItems / 2);

    state.itemElements.forEach((el, index) => {
      const offset = index - rawIndex;

      // Only calculate 3D transform for items near the active window
      if (Math.abs(offset) <= halfVisible + 1.5) {
        el.style.display = 'flex';
        // 3D Cylinder geometry
        const angle = offset * 21; // degrees curved around cylinder
        const rad = (angle * Math.PI) / 180;
        const radius = 100;
        const translateY = Math.sin(rad) * radius;
        const translateZ = (Math.cos(rad) - 1) * radius;
        const rotateX = -angle;

        // Visual curve & opacity falloff
        const absOffset = Math.abs(offset);
        const opacity = Math.max(0.12, Math.cos(rad) ** 2.2);
        const scale = Math.max(0.78, 1 - absOffset * 0.04);

        el.style.transform = `translate3d(0, ${translateY}px, ${translateZ}px) rotateX(${rotateX}deg) scale(${scale})`;
        el.style.opacity = opacity.toFixed(3);

        if (Math.abs(offset) < 0.5) {
          el.classList.add('is-active');
        } else {
          el.classList.remove('is-active');
        }
      } else {
        el.style.display = 'none';
        el.classList.remove('is-active');
      }
    });
  }

  snapToIndex(state, index, animate = true) {
    if (state.animationFrame) cancelAnimationFrame(state.animationFrame);

    let targetIndex = index;
    const maxIdx = state.items.length - 1;

    if (!state.isLooping) {
      targetIndex = Math.max(0, Math.min(targetIndex, maxIdx));
    } else {
      // Loop recentering if getting close to ends of repeated buffer
      const base = state.baseCount;
      if (targetIndex < base) {
        targetIndex += base * 2;
        state.currentY += base * 2 * state.itemHeight;
      } else if (targetIndex >= maxIdx - base) {
        targetIndex -= base * 2;
        state.currentY -= base * 2 * state.itemHeight;
      }
    }

    const startY = state.currentY;
    const endY = targetIndex * state.itemHeight;
    state.currentIndex = targetIndex;

    if (!animate) {
      state.currentY = endY;
      this.renderColumnTransform(state);
      this.updateStateValue();
      return;
    }

    const duration = 280; // ms snappy Apple animation
    const startTime = performance.now();

    const step = (now) => {
      const elapsed = now - startTime;
      const progress = Math.min(1, elapsed / duration);
      // Apple-style cubic ease-out
      const ease = 1 - Math.pow(1 - progress, 3);

      state.currentY = startY + (endY - startY) * ease;
      this.renderColumnTransform(state);

      if (progress < 1) {
        state.animationFrame = requestAnimationFrame(step);
      } else {
        state.currentY = endY;
        this.renderColumnTransform(state);
        this.updateStateValue();
      }
    };

    state.animationFrame = requestAnimationFrame(step);
  }

  updateStateValue() {
    // Read current snapped values
    const dateItem = this.columns.date.items[this.columns.date.currentIndex];
    const hourItem = this.columns.hour.items[this.columns.hour.currentIndex];
    const minuteItem = this.columns.minute.items[this.columns.minute.currentIndex];
    const periodItem = this.columns.period.items[this.columns.period.currentIndex];

    if (!dateItem || !hourItem || !minuteItem || !periodItem) return;

    this.currentValue = {
      date: dateItem.value,
      dateLabel: dateItem.label,
      dateObj: dateItem.dateObj,
      hour: hourItem.value,
      minute: minuteItem.value,
      period: periodItem.value
    };

    if (this.options.onChange) {
      this.options.onChange(this.getValue());
    }
  }

  getValue() {
    const { date, hour, minute, period, dateLabel } = this.currentValue;

    // Convert 12h to 24h
    let h24 = hour;
    if (period === 'PM' && hour < 12) h24 = hour + 12;
    if (period === 'AM' && hour === 12) h24 = 0;

    const time24 = `${String(h24).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
    const displayTime = `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')} ${period}`;

    // Create a synthesized Date object
    let fullDate = new Date();
    if (date) {
      const [y, m, d] = date.split('-').map(Number);
      fullDate = new Date(y, m - 1, d, h24, minute, 0, 0);
    }

    return {
      date,
      dateLabel: dateLabel || date,
      hour,
      minute,
      period,
      time24,
      displayTime,
      fullDate,
      timestamp: fullDate.getTime()
    };
  }

  setValue(targetDate, animate = true) {
    const d = targetDate instanceof Date ? targetDate : new Date(targetDate);
    if (isNaN(d.getTime())) return;

    // 1. Date column index
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    const iso = `${yyyy}-${mm}-${dd}`;

    let dateIdx = this.columns.date.items.findIndex(item => item.value === iso);
    if (dateIdx === -1) {
      // Find closest or default to today
      dateIdx = this.columns.date.items.findIndex(item => item.isToday);
      if (dateIdx === -1) dateIdx = 0;
    }

    // 2. Hour (1-12) & Period (AM/PM)
    let rawH = d.getHours();
    const period = rawH >= 12 ? 'PM' : 'AM';
    let h12 = rawH % 12;
    if (h12 === 0) h12 = 12;

    // Find hour index in middle buffer of repeated list
    const baseH = this.columns.hour.baseCount;
    const hourOffset = h12 - 1;
    const hourIdx = baseH * 2 + hourOffset;

    // 3. Minute (0-59)
    const m = d.getMinutes();
    const baseM = this.columns.minute.baseCount;
    const minuteIdx = baseM + m;

    // 4. Period
    const periodIdx = period === 'AM' ? 0 : 1;

    // Snap all columns
    this.snapToIndex(this.columns.date, dateIdx, animate);
    this.snapToIndex(this.columns.hour, hourIdx, animate);
    this.snapToIndex(this.columns.minute, minuteIdx, animate);
    this.snapToIndex(this.columns.period, periodIdx, animate);
  }

  setNow(animate = true) {
    this.setValue(new Date(), animate);
  }

  adjustHours(delta, animate = true) {
    const current = this.getValue().fullDate;
    const next = new Date(current.getTime() + delta * 3600 * 1000);
    this.setValue(next, animate);
  }
}
