import test from "node:test";
import assert from "node:assert/strict";

import { QualityManager, QUALITY_PROFILES } from "../../static/js/landing/quality-manager.js";

class MockDocument extends EventTarget {
  constructor() {
    super();
    this.documentElement = { dataset: {} };
    this.defaultView = { CustomEvent };
  }
}

function createWindow({ search = "", width = 1440, height = 900, dpr = 1.5, now = () => 1000 } = {}) {
  return {
    innerWidth: width,
    innerHeight: height,
    devicePixelRatio: dpr,
    location: { search },
    performance: { now },
    navigator: {
      hardwareConcurrency: 16,
      deviceMemory: 16,
      userAgent: "Desktop test browser",
    },
  };
}

test("quality profile honors a pinned forced mode and DPR cap", () => {
  const document = new MockDocument();
  document.documentElement.dataset.webgl = "pending";
  const canvas = { getContext: (name) => name === "webgl2" ? {} : null };
  const window = createWindow({ search: "?landingQuality=medium", dpr: 2.5 });
  const manager = new QualityManager({ canvas, documentRef: document, windowRef: window });
  assert.equal(manager.quality, "medium");
  assert.equal(manager.pixelRatio, QUALITY_PROFILES.medium.dprMax);
  assert.equal(document.documentElement.dataset.webgl, "pending");
});

test("missing WebGL2 produces the polished fallback mode", () => {
  const document = new MockDocument();
  const manager = new QualityManager({
    canvas: { getContext: () => null },
    documentRef: document,
    windowRef: createWindow(),
  });
  assert.equal(manager.quality, "fallback");
  assert.equal(manager.isFallback, true);
  assert.equal(document.documentElement.dataset.webgl, "fallback");
});

test("sustained poor frame pacing downgrades once and emits the contract event", () => {
  let clock = 1000;
  const document = new MockDocument();
  const changes = [];
  document.addEventListener("landing:qualitychange", (event) => changes.push(event.detail));
  const manager = new QualityManager({
    canvas: { getContext: () => ({}) },
    documentRef: document,
    windowRef: createWindow({ now: () => clock }),
  });
  assert.equal(manager.quality, "high");

  for (let index = 0; index < 150; index += 1) {
    clock += 50;
    manager.observeFrame(50, clock);
  }
  assert.equal(manager.quality, "medium");
  assert.deepEqual(changes[0], { previous: "high", current: "medium", reason: "frame-rate" });
});

test("low mode retains a stable 30 FPS mobile cadence", () => {
  let clock = 1000;
  const document = new MockDocument();
  const changes = [];
  document.addEventListener("landing:qualitychange", (event) => changes.push(event.detail));
  const manager = new QualityManager({
    canvas: { getContext: () => ({}) },
    documentRef: document,
    windowRef: createWindow({ search: "?landingQuality=low", now: () => clock }),
  });
  for (let index = 0; index < 400; index += 1) {
    clock += 1000 / 30;
    manager.observeFrame(1000 / 30, clock);
  }
  assert.equal(manager.quality, "low");
  assert.deepEqual(changes, []);
});

test("low mode reaches fallback only after three sustained severe windows", () => {
  let clock = 1000;
  const document = new MockDocument();
  const changes = [];
  document.addEventListener("landing:qualitychange", (event) => changes.push(event.detail));
  const manager = new QualityManager({
    canvas: { getContext: () => ({}) },
    documentRef: document,
    windowRef: createWindow({ search: "?landingQuality=low", now: () => clock }),
  });
  for (let windowIndex = 0; windowIndex < 3; windowIndex += 1) {
    for (let sample = 0; sample < 120; sample += 1) {
      clock += 100;
      manager.observeFrame(100, clock);
    }
    if (windowIndex < 2) assert.equal(manager.quality, "low");
  }
  assert.equal(manager.quality, "fallback");
  assert.deepEqual(changes.at(-1), { previous: "low", current: "fallback", reason: "frame-rate" });
});
