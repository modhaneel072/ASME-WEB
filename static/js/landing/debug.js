export class LandingDebug {
  constructor({ documentRef = globalThis.document, windowRef = globalThis.window } = {}) {
    this.document = documentRef || null;
    this.window = windowRef || null;
    this.enabled = isDebugEnabled(this.window);
    this.frames = 0;
    this.fps = 0;
    this.sampleStartedAt = timestamp(this.window);
    this.lastPaintAt = -Infinity;
    this.element = null;
  }

  init() {
    if (!this.enabled || !this.document || this.element) return this;
    this.element = this.document.createElement("pre");
    this.element.id = "landing-debug";
    this.element.setAttribute("aria-hidden", "true");
    Object.assign(this.element.style, {
      position: "fixed",
      left: "12px",
      bottom: "12px",
      zIndex: "2147483647",
      margin: "0",
      padding: "10px 12px",
      color: "#ffcd00",
      background: "rgba(5, 6, 7, .88)",
      border: "1px solid rgba(255, 205, 0, .4)",
      font: "11px/1.45 ui-monospace, SFMono-Regular, Consolas, monospace",
      pointerEvents: "none",
      whiteSpace: "pre-wrap",
    });
    this.document.body?.append?.(this.element);
    return this;
  }

  update({ progress, scene, quality, assets, camera, target, chapter, drawCalls, triangles } = {}) {
    if (!this.enabled || !this.element) return;
    const current = timestamp(this.window);
    this.frames += 1;
    const elapsed = current - this.sampleStartedAt;
    if (elapsed >= 500) {
      this.fps = (this.frames * 1000) / elapsed;
      this.frames = 0;
      this.sampleStartedAt = current;
    }
    if (current - this.lastPaintAt < 180) return;
    this.lastPaintAt = current;

    const loaded = typeof assets === "number" ? assets : assets?.loaded ?? 0;
    const registered = assets?.registered ?? "?";
    this.element.textContent = [
      "LANDING DEBUG",
      `progress  ${number(progress, 4)}`,
      `scene     ${scene || "unknown"}`,
      `chapter   ${chapter || scene || "unknown"}`,
      `fps       ${number(this.fps, 1)}`,
      `quality   ${quality || "unknown"}`,
      `assets    ${loaded}/${registered}`,
      `draws     ${integer(drawCalls)}`,
      `triangles ${integer(triangles)}`,
      `camera    ${vector(camera)}`,
      `target    ${vector(target)}`,
    ].join("\n");
  }

  dispose() {
    this.element?.remove?.();
    this.element = null;
  }
}

export function isDebugEnabled(windowRef = globalThis.window) {
  try {
    return new URLSearchParams(windowRef?.location?.search || "").get("landingDebug") === "1";
  } catch {
    return false;
  }
}

function timestamp(windowRef) {
  return Number(windowRef?.performance?.now?.()) || Date.now();
}

function number(value, digits) {
  return Number.isFinite(value) ? value.toFixed(digits) : "0";
}

function integer(value) {
  return Number.isFinite(value) ? String(Math.round(value)) : "0";
}

function vector(value) {
  if (!value) return "0, 0, 0";
  const values = value.isVector3 ? [value.x, value.y, value.z] : value;
  return Array.from(values || [0, 0, 0]).slice(0, 3).map((item) => number(Number(item), 2)).join(", ");
}
