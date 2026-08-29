import test from "node:test";
import assert from "node:assert/strict";

import { InputController } from "../../static/js/landing/input-controller.js";
import { ScrollController } from "../../static/js/landing/scroll-controller.js";
import { TeamInteractions } from "../../static/js/landing/team-interactions.js";
import { SCENE_RANGES } from "../../static/js/landing/camera-path.js";

class MockStyle {
  constructor() { this.values = new Map(); }
  setProperty(name, value) { this.values.set(name, value); }
}

class MockClassList {
  constructor() { this.values = new Set(); }
  add(...values) { values.forEach((value) => this.values.add(value)); }
  remove(...values) { values.forEach((value) => this.values.delete(value)); }
}

class MockElement extends EventTarget {
  constructor(id = "") {
    super();
    this.id = id;
    this.hidden = false;
    this.dataset = {};
    this.attributes = new Map();
    this.style = new MockStyle();
    this.classList = new MockClassList();
    this.children = [];
    this.tabIndex = 0;
  }
  setAttribute(name, value) { this.attributes.set(name, String(value)); }
  getAttribute(name) { return this.attributes.get(name); }
  removeAttribute(name) { this.attributes.delete(name); }
  hasAttribute(name) { return this.attributes.has(name); }
  append(...nodes) { this.children.push(...nodes); }
  focus() { this.focused = true; }
  scrollIntoView(options) { this.scrolled = options; }
  querySelector() { return this.focusTarget || null; }
}

class MockWindow extends EventTarget {
  constructor() {
    super();
    this.scrollY = 0;
    this.innerHeight = 100;
    this.innerWidth = 200;
    this.location = { hash: "", search: "" };
    this.mediaQuery = Object.assign(new EventTarget(), { matches: false, addListener() {}, removeListener() {} });
    this.MutationObserver = class {
      constructor(callback) { this.callback = callback; }
      observe() {}
      disconnect() {}
    };
  }
  matchMedia() { return this.mediaQuery; }
  scrollTo(options) { this.lastScroll = options; this.scrollY = options.top; }
  requestAnimationFrame(callback) { callback(); return 1; }
  setTimeout(callback) { callback(); return 1; }
}

class MockDocument extends EventTarget {
  constructor(window) {
    super();
    this.defaultView = Object.assign(window, { CustomEvent });
    this.documentElement = { dataset: {} };
    this.hidden = false;
    this.elements = new Map();
    this.activeElement = null;
  }
  add(element) { this.elements.set(element.id, element); return element; }
  getElementById(id) { return this.elements.get(id) || null; }
  createElement() { return new MockElement(); }
}

test("input controller keeps scroll native while skip, sound, panel, and Escape work", async () => {
  const window = new MockWindow();
  const document = new MockDocument(window);
  const canvas = document.add(new MockElement("landing-canvas"));
  const skip = document.add(new MockElement("skip-experience"));
  const sound = document.add(new MockElement("sound-toggle"));
  const panel = document.add(new MockElement("team-panel"));
  const close = document.add(new MockElement("team-panel-close"));
  const finalCta = document.add(new MockElement("landing-final-cta"));
  finalCta.focusTarget = new MockElement("join-link");
  panel.hidden = true;

  let soundEnabled = false;
  const pauseStates = [];
  const pageHideEvents = [];
  const controller = new InputController({
    documentRef: document,
    windowRef: window,
    onSoundChange: async (enabled) => { soundEnabled = enabled; },
    onPauseChange: (paused) => pauseStates.push(paused),
    onPageHide: (event) => pageHideEvents.push(event.persisted),
  }).init();

  assert.equal(canvas.style.pointerEvents, "none");
  skip.dispatchEvent(new Event("click", { cancelable: true }));
  assert.equal(finalCta.scrolled.block, "start");

  sound.dispatchEvent(new Event("click"));
  await Promise.resolve();
  assert.equal(soundEnabled, true);
  assert.equal(sound.getAttribute("aria-pressed"), "true");

  document.dispatchEvent(new CustomEvent("landing:teampanel", { detail: { team: "iam3d" } }));
  assert.equal(panel.hidden, false);
  assert.equal(panel.dataset.team, "iam3d");
  assert.equal(close.focused, true);

  const escape = new Event("keydown", { cancelable: true });
  Object.defineProperty(escape, "key", { value: "Escape" });
  document.dispatchEvent(escape);
  assert.equal(panel.hidden, true);
  assert.equal(controller.paused, false);
  assert.equal(document.documentElement.dataset.landingAudio, "available");

  document.dispatchEvent(new CustomEvent("landing:teampanel", { detail: { team: "iam3d" } }));
  document.documentElement.dataset.teamPanel = "closed";
  controller.panelObserver.callback([]);
  assert.equal(controller.paused, false);
  assert.equal(pauseStates.at(-1), false);
  panel.hidden = true;
  panel.hidden = false;
  document.documentElement.dataset.teamPanel = "open";
  controller.panelObserver.callback([]);
  assert.equal(controller.paused, true);
  assert.equal(panel.getAttribute("aria-hidden"), "false");
  panel.hidden = true;
  document.documentElement.dataset.teamPanel = "closed";
  controller.panelObserver.callback([]);
  document.hidden = true;
  document.dispatchEvent(new Event("visibilitychange"));
  document.hidden = false;
  document.dispatchEvent(new Event("visibilitychange"));
  assert.deepEqual(pauseStates.slice(-2), [true, false]);
  const pageHide = new Event("pagehide");
  Object.defineProperty(pageHide, "persisted", { value: true });
  window.dispatchEvent(pageHide);
  assert.equal(pauseStates.at(-1), true);
  assert.deepEqual(pageHideEvents, [true]);
  window.dispatchEvent(new Event("pageshow"));
  assert.equal(pauseStates.at(-1), false);
  controller.dispose();
});

test("native fallback scroll produces normalized progress and scene events", () => {
  const window = new MockWindow();
  window.scrollY = 450;
  const document = new MockDocument(window);
  const progress = document.add(new MockElement("landing-progress"));
  const label = document.add(new MockElement("landing-progress-label"));
  const chapters = Object.fromEntries(Object.entries(SCENE_RANGES).map(([scene, range]) => [
    scene,
    { dataset: { start: String(range.start), end: String(range.end) } },
  ]));
  const root = new MockElement("landing-root");
  root.scrollHeight = 1000;
  root.offsetHeight = 1000;
  root.getBoundingClientRect = () => ({ top: -window.scrollY, height: 1000 });
  root.querySelector = (selector) => chapters[selector.match(/data-scene="([^"]+)"/)?.[1]] || null;

  const changes = [];
  document.addEventListener("landing:scenechange", (event) => changes.push(event.detail));
  const controller = new ScrollController({ root, documentRef: document, windowRef: window }).init();
  assert.equal(controller.progress, 0.5);
  assert.equal(controller.scene, "rov");
  assert.equal(document.documentElement.dataset.landingScene, "rov");
  assert.equal(progress.getAttribute("aria-valuenow"), "50");
  assert.equal(label.textContent, "rov");

  window.scrollY = 900;
  window.dispatchEvent(new Event("scroll"));
  assert.equal(controller.progress, 1);
  assert.equal(controller.scene, "join");
  assert.equal(changes.at(-1).current, "join");
  controller.dispose();
});

test("team hotspot event exposes only approved source-of-truth team keys", () => {
  const window = new MockWindow();
  const document = new MockDocument(window);
  const events = [];
  document.addEventListener("landing:teampanel", (event) => events.push(event.detail));
  const interactions = new TeamInteractions({ documentRef: document });
  assert.equal(interactions.open("robotic-dog"), true);
  assert.equal(interactions.open("outdated-placeholder"), false);
  assert.deepEqual(events, [{ team: "robotic-dog" }]);
});

test("chapter-height mapping keeps non-proportional DOM chapters on declared scene ranges", () => {
  const window = new MockWindow();
  window.scrollY = 700;
  window.innerHeight = 100;
  const document = new MockDocument(window);
  document.documentElement.scrollHeight = 1800;
  document.body = { scrollHeight: 1800 };
  const heights = {
    intro: 100,
    design: 300,
    iam3d: 100,
    plane: 400,
    rov: 100,
    "student-design": 150,
    "robotic-dog": 200,
    people: 150,
    join: 300,
  };
  let top = 0;
  const chapters = {};
  for (const [scene, range] of Object.entries(SCENE_RANGES)) {
    const absoluteTop = top;
    const height = heights[scene];
    chapters[scene] = {
      dataset: { start: String(range.start), end: String(range.end) },
      offsetHeight: height,
      getBoundingClientRect: () => ({ top: absoluteTop - window.scrollY, height }),
    };
    top += height;
  }
  const root = new MockElement("landing-root");
  root.scrollHeight = top;
  root.offsetHeight = top;
  root.getBoundingClientRect = () => ({ top: -window.scrollY, height: top });
  root.querySelector = (selector) => chapters[selector.match(/data-scene="([^"]+)"/)?.[1]] || null;

  const controller = new ScrollController({ root, documentRef: document, windowRef: window }).init();
  assert.ok(Math.abs(controller.progress - 0.37) < 0.0001);
  assert.equal(controller.scene, "plane");
  controller.scrollToProgress(0.5, { behavior: "auto" });
  assert.ok(Math.abs(window.lastScroll.top - 953.8461538) < 0.001);
  controller.dispose();
});
