const QUALITY_ORDER = Object.freeze(["high", "medium", "low", "fallback"]);

export const QUALITY_PROFILES = Object.freeze({
  high: Object.freeze({
    dprMax: 1.75,
    radialSegments: 12,
    curveSegments: 220,
    shadowMapSize: 1024,
    shadows: true,
    underwaterParticles: 72,
    photoLayers: 4,
    waterTransmission: 0.2,
  }),
  medium: Object.freeze({
    dprMax: 1.35,
    radialSegments: 8,
    curveSegments: 160,
    shadowMapSize: 512,
    shadows: false,
    underwaterParticles: 30,
    photoLayers: 3,
    waterTransmission: 0.1,
  }),
  low: Object.freeze({
    dprMax: 1,
    radialSegments: 6,
    curveSegments: 100,
    shadowMapSize: 0,
    shadows: false,
    underwaterParticles: 0,
    photoLayers: 2,
    waterTransmission: 0,
  }),
  fallback: Object.freeze({
    dprMax: 1,
    radialSegments: 0,
    curveSegments: 0,
    shadowMapSize: 0,
    shadows: false,
    underwaterParticles: 0,
    photoLayers: 0,
    waterTransmission: 0,
  }),
});

export class QualityManager {
  constructor({ canvas, windowRef = globalThis.window, documentRef = globalThis.document } = {}) {
    this.canvas = canvas || null;
    this.window = windowRef || null;
    this.document = documentRef || null;
    this.context = null;
    this.samples = [];
    this.startedAt = now(this.window);
    this.lastEvaluationAt = this.startedAt;
    this.lastChangeAt = -Infinity;
    this.severeWindows = 0;
    this.disposed = false;
    this.quality = this.#selectInitialQuality();

    if (this.document?.documentElement && this.quality === "fallback") {
      this.document.documentElement.dataset.webgl = "fallback";
    }
  }

  get profile() {
    return QUALITY_PROFILES[this.quality];
  }

  get pixelRatio() {
    const dpr = Number(this.window?.devicePixelRatio) || 1;
    return Math.min(dpr, this.profile.dprMax);
  }

  get isFallback() {
    return this.quality === "fallback";
  }

  observeFrame(deltaMs, timestamp = now(this.window)) {
    if (this.disposed || this.isFallback || !Number.isFinite(deltaMs) || deltaMs <= 0 || deltaMs > 1000) {
      return this.quality;
    }

    this.samples.push(deltaMs);
    if (this.samples.length > 360) this.samples.shift();

    const warm = timestamp - this.startedAt >= 6000;
    const evaluationDue = timestamp - this.lastEvaluationAt >= 2500;
    const stableAfterChange = timestamp - this.lastChangeAt >= 15000;
    if (!warm || !evaluationDue || !stableAfterChange || this.samples.length < 120) return this.quality;

    this.lastEvaluationAt = timestamp;
    const sorted = [...this.samples].sort((a, b) => a - b);
    const medianMs = percentile(sorted, 0.5);
    const p90Ms = percentile(sorted, 0.9);
    const medianFps = 1000 / Math.max(1, medianMs);
    const poorThreshold = this.quality === "high" ? 46 : this.quality === "medium" ? 34 : 24;
    const poorP90 = this.quality === "high" ? 31 : this.quality === "medium" ? 44 : 58;
    const poor = medianFps < poorThreshold || p90Ms > poorP90;
    const severe = medianFps < 18 || p90Ms > 95;
    this.severeWindows = severe ? this.severeWindows + 1 : 0;

    if (this.severeWindows >= 3) {
      this.setQuality("fallback", "frame-rate");
    } else if (poor && this.quality !== "low") {
      this.downgrade("frame-rate");
    }

    this.samples.length = 0;
    return this.quality;
  }

  downgrade(reason = "frame-rate") {
    const index = QUALITY_ORDER.indexOf(this.quality);
    if (index < 0 || index >= QUALITY_ORDER.length - 1) return this.quality;
    return this.setQuality(QUALITY_ORDER[index + 1], reason);
  }

  setQuality(next, reason = "manual") {
    if (!QUALITY_PROFILES[next] || next === this.quality) return this.quality;
    const previous = this.quality;
    this.quality = next;
    this.lastChangeAt = now(this.window);
    this.samples.length = 0;
    if (this.document?.documentElement) {
      if (next === "fallback") this.document.documentElement.dataset.webgl = "fallback";
      else if (this.document.documentElement.dataset.webgl !== "pending") {
        this.document.documentElement.dataset.webgl = "available";
      }
    }
    dispatchDocumentEvent(this.document, "landing:qualitychange", { previous, current: next, reason });
    return this.quality;
  }

  dispose() {
    this.disposed = true;
    this.samples.length = 0;
    this.context = null;
  }

  #selectInitialQuality() {
    const forced = forcedQuality(this.window);
    if (forced) {
      if (forced === "fallback") return "fallback";
      const context = this.#createContext();
      return context ? forced : "fallback";
    }

    const context = this.#createContext();
    if (!context) return "fallback";

    const nav = this.window?.navigator || globalThis.navigator || {};
    const width = Math.min(Number(this.window?.innerWidth) || 1280, Number(this.window?.innerHeight) || 720);
    const cores = Number(nav.hardwareConcurrency) || 4;
    const memory = Number(nav.deviceMemory) || 4;
    const dpr = Number(this.window?.devicePixelRatio) || 1;
    const mobile = /Android|iPhone|iPad|iPod|Mobile/i.test(nav.userAgent || "");

    if (mobile || width < 520 || cores <= 4 || memory <= 3) return "low";
    if (width < 820 || cores <= 8 || memory <= 6 || dpr > 2.25) return "medium";
    return "high";
  }

  #createContext() {
    if (!this.canvas?.getContext) return null;
    try {
      this.context = this.canvas.getContext("webgl2", {
        alpha: true,
        antialias: false,
        depth: true,
        powerPreference: "high-performance",
        premultipliedAlpha: true,
        preserveDrawingBuffer: false,
      });
      return this.context;
    } catch (error) {
      console.warn("[landing] WebGL2 capability check failed.", error);
      this.context = null;
      return null;
    }
  }
}

function forcedQuality(windowRef) {
  try {
    const params = new URLSearchParams(windowRef?.location?.search || "");
    const value = params.get("landingQuality");
    if (QUALITY_PROFILES[value]) return value;
    if (params.get("landingFallback") === "1") return "fallback";
  } catch {
    // An unavailable or opaque location should not prevent feature detection.
  }
  return null;
}

function percentile(sorted, ratio) {
  if (!sorted.length) return 0;
  return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * ratio))];
}

function now(windowRef) {
  return Number(windowRef?.performance?.now?.()) || Date.now();
}

function dispatchDocumentEvent(documentRef, name, detail) {
  if (!documentRef?.dispatchEvent) return;
  const EventCtor = documentRef.defaultView?.CustomEvent || globalThis.CustomEvent;
  if (typeof EventCtor === "function") documentRef.dispatchEvent(new EventCtor(name, { detail }));
}
