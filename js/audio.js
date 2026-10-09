/**
 * audio.js - Apple Clock Haptic & Mechanical Sound Synthesizer
 * Uses Web Audio API to create authentic rotary dial clicks without external audio files
 */

class SoundEngine {
  constructor() {
    this.ctx = null;
    this.muted = false;
    this.lastTickTime = 0;
    this.minTickInterval = 35; // Debounce rapid ticks for smooth feel

    // Restore user mute preference
    const savedMute = localStorage.getItem('class_logger_sound_muted');
    if (savedMute !== null) {
      this.muted = savedMute === 'true';
    }
  }

  initContext() {
    if (!this.ctx && typeof window !== 'undefined') {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (AudioCtx) {
        this.ctx = new AudioCtx();
      }
    }
    if (this.ctx && this.ctx.state === 'suspended') {
      this.ctx.resume().catch(() => {});
    }
  }

  setMuted(muted) {
    this.muted = muted;
    localStorage.setItem('class_logger_sound_muted', String(muted));
  }

  toggleMute() {
    this.setMuted(!this.muted);
    return this.muted;
  }

  isMuted() {
    return this.muted;
  }

  /**
   * Rotary Dial Mechanical Click / Tick (Apple Clock alarm feel)
   */
  playTick(pitchModifier = 1.0) {
    if (this.muted) return;

    const now = Date.now();
    if (now - this.lastTickTime < this.minTickInterval) return;
    this.lastTickTime = now;

    // Haptic vibration feedback on supported mobile devices
    if (typeof navigator !== 'undefined' && typeof navigator.vibrate === 'function') {
      try { navigator.vibrate(6); } catch (_) {}
    }

    this.initContext();
    if (!this.ctx) return;

    try {
      const ctx = this.ctx;
      const t = ctx.currentTime;

      // 1. High frequency transient pop (mechanical click)
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();

      osc.type = 'triangle';
      osc.frequency.setValueAtTime(1400 * pitchModifier, t);
      osc.frequency.exponentialRampToValueAtTime(280 * pitchModifier, t + 0.014);

      gain.gain.setValueAtTime(0.22, t);
      gain.gain.exponentialRampToValueAtTime(0.001, t + 0.015);

      osc.connect(gain);
      gain.connect(ctx.destination);

      osc.start(t);
      osc.stop(t + 0.016);

      // 2. Subtle low-end body resonance
      const osc2 = ctx.createOscillator();
      const gain2 = ctx.createGain();

      osc2.type = 'sine';
      osc2.frequency.setValueAtTime(450 * pitchModifier, t);
      osc2.frequency.exponentialRampToValueAtTime(80, t + 0.018);

      gain2.gain.setValueAtTime(0.12, t);
      gain2.gain.exponentialRampToValueAtTime(0.001, t + 0.02);

      osc2.connect(gain2);
      gain2.connect(ctx.destination);

      osc2.start(t);
      osc2.stop(t + 0.02);

      // Light haptic vibration on supporting mobile devices
      if (typeof navigator !== 'undefined' && navigator.vibrate) {
        navigator.vibrate(6);
      }
    } catch (e) {
      // Audio autoplay policy or silent fail
    }
  }

  /**
   * Rewarding chime when a class is logged
   */
  playSuccess() {
    if (this.muted) return;
    this.initContext();
    if (!this.ctx) return;

    try {
      const ctx = this.ctx;
      const t = ctx.currentTime;

      const notes = [523.25, 659.25, 783.99, 1046.50]; // C5, E5, G5, C6
      notes.forEach((freq, idx) => {
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();

        osc.type = 'sine';
        osc.frequency.setValueAtTime(freq, t + idx * 0.06);

        gain.gain.setValueAtTime(0, t + idx * 0.06);
        gain.gain.linearRampToValueAtTime(0.15, t + idx * 0.06 + 0.02);
        gain.gain.exponentialRampToValueAtTime(0.0001, t + idx * 0.06 + 0.35);

        osc.connect(gain);
        gain.connect(ctx.destination);

        osc.start(t + idx * 0.06);
        osc.stop(t + idx * 0.06 + 0.36);
      });

      if (typeof navigator !== 'undefined' && navigator.vibrate) {
        navigator.vibrate([10, 30, 20]);
      }
    } catch (e) {}
  }

  /**
   * Deletion pop
   */
  playTrash() {
    if (this.muted) return;
    this.initContext();
    if (!this.ctx) return;

    try {
      const ctx = this.ctx;
      const t = ctx.currentTime;
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();

      osc.type = 'sawtooth';
      osc.frequency.setValueAtTime(260, t);
      osc.frequency.exponentialRampToValueAtTime(70, t + 0.08);

      gain.gain.setValueAtTime(0.15, t);
      gain.gain.exponentialRampToValueAtTime(0.001, t + 0.08);

      osc.connect(gain);
      gain.connect(ctx.destination);

      osc.start(t);
      osc.stop(t + 0.09);
    } catch (e) {}
  }

  /**
   * Alert / validation error tone
   */
  playError() {
    if (this.muted) return;
    this.initContext();
    if (!this.ctx) return;

    try {
      const ctx = this.ctx;
      const t = ctx.currentTime;
      [220, 180].forEach((freq, idx) => {
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'sawtooth';
        osc.frequency.setValueAtTime(freq, t + idx * 0.08);
        gain.gain.setValueAtTime(0.12, t + idx * 0.08);
        gain.gain.exponentialRampToValueAtTime(0.001, t + idx * 0.08 + 0.09);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start(t + idx * 0.08);
        osc.stop(t + idx * 0.08 + 0.1);
      });
      if (typeof navigator !== 'undefined' && navigator.vibrate) {
        navigator.vibrate([25, 30, 25]);
      }
    } catch (e) {}
  }
}

export const sound = new SoundEngine();
