import test from "node:test";
import assert from "node:assert/strict";

import { LandingAssetLoader, SAFE_CONTENT } from "../../static/js/landing/asset-loader.js";

test("missing content data degrades to the five correct team labels", async () => {
  const originalFetch = globalThis.fetch;
  const originalWarn = console.warn;
  globalThis.fetch = async () => ({ ok: false, status: 404 });
  console.warn = () => {};
  try {
    const loader = new LandingAssetLoader({ documentRef: null, timeoutMs: 100 });
    const content = await loader.loadContent("/missing-content.json");
    assert.equal(content, SAFE_CONTENT);
    assert.deepEqual(Object.keys(content.teams).sort(), [
      "autonomous-plane", "iam3d", "mate-rov", "robotic-dog", "student-design",
    ]);
    assert.equal(content.teams["robotic-dog"].status, "Selective team — interview required");
    assert.equal(loader.contentSource, "fallback");
  } finally {
    globalThis.fetch = originalFetch;
    console.warn = originalWarn;
  }
});

test("valid repository content remains the primary copy source", async () => {
  const originalFetch = globalThis.fetch;
  const payload = structuredClone(SAFE_CONTENT);
  globalThis.fetch = async () => ({ ok: true, json: async () => payload });
  try {
    const loader = new LandingAssetLoader({ documentRef: null, timeoutMs: 100 });
    const content = await loader.loadContent();
    assert.equal(content, payload);
    assert.equal(loader.contentSource, "repository");
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("Claude's array-shaped team data is normalized to event contract keys", async () => {
  const originalFetch = globalThis.fetch;
  const payload = {
    teams: [
      { id: "iam3d", name: "IAM3D" },
      { id: "plane", name: "Autonomous Plane" },
      { id: "rov", name: "MATE ROV" },
      { id: "student-design", name: "Student Design Competition" },
      {
        id: "robotic-dog",
        name: "Robotic Dog",
        requires_interview: true,
        interview_note: "SELECTIVE TEAM — INTERVIEW REQUIRED",
      },
    ],
  };
  globalThis.fetch = async () => ({ ok: true, json: async () => payload });
  try {
    const loader = new LandingAssetLoader({ documentRef: null, timeoutMs: 100 });
    const content = await loader.loadContent();
    assert.deepEqual(Object.keys(content.teams), [
      "iam3d", "autonomous-plane", "mate-rov", "student-design", "robotic-dog",
    ]);
    assert.equal(content.teams["autonomous-plane"].id, "autonomous-plane");
    assert.equal(content.teams["mate-rov"].id, "mate-rov");
    assert.equal(content.teams["robotic-dog"].status, "Selective team — interview required");
    assert.equal(loader.contentSource, "repository");
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("a missing optional chapter asset is recorded without rejecting the scene", async () => {
  const originalFetch = globalThis.fetch;
  const originalWarn = console.warn;
  globalThis.fetch = async () => ({ ok: false, status: 404 });
  console.warn = () => {};
  try {
    const loader = new LandingAssetLoader({ documentRef: null, timeoutMs: 100 });
    loader.register([{ id: "optional-photo", type: "json", url: "/missing.json", chapter: "people" }]);
    const asset = await loader.load("optional-photo");
    assert.equal(asset, null);
    assert.equal(loader.errors.length, 1);
    assert.equal(loader.errors[0].required, false);
  } finally {
    globalThis.fetch = originalFetch;
    console.warn = originalWarn;
  }
});

test("an inflight optional asset cannot repopulate a disposed loader", async () => {
  const originalFetch = globalThis.fetch;
  let release;
  globalThis.fetch = () => new Promise((resolve) => { release = resolve; });
  try {
    const loader = new LandingAssetLoader({ documentRef: null, timeoutMs: 1000 });
    loader.register([{ id: "slow-data", type: "json", url: "/slow.json", chapter: "people" }]);
    const pending = loader.load("slow-data");
    loader.dispose();
    release({ ok: true, json: async () => ({ value: 1 }) });
    assert.equal(await pending, null);
    assert.equal(loader.getSummary().loaded, 0);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("a quality downgrade removes unneeded lazy chapter entries before preload", () => {
  const loader = new LandingAssetLoader({ documentRef: null });
  loader.register([0, 1, 2, 3].map((index) => ({
    id: `photo-${index}`,
    type: "texture",
    url: `/photo-${index}.jpg`,
    chapter: "people",
  })));
  assert.deepEqual(loader.limitChapterEntries("people", 2), ["photo-0", "photo-1"]);
  assert.equal(loader.getSummary().registered, 2);
});

test("background chapter preloads do not reopen the global landing loader", async () => {
  const originalFetch = globalThis.fetch;
  const document = new EventTarget();
  document.defaultView = { CustomEvent };
  let loadingEvents = 0;
  document.addEventListener("landing:loading", () => { loadingEvents += 1; });
  globalThis.fetch = async () => ({ ok: true, json: async () => ({ value: 1 }) });
  try {
    const loader = new LandingAssetLoader({ documentRef: document, timeoutMs: 100 });
    loader.register([{ id: "lazy-data", type: "json", url: "/lazy.json", chapter: "people" }]);
    await loader.preloadChapter("people");
    assert.equal(loadingEvents, 0);
  } finally {
    globalThis.fetch = originalFetch;
  }
});
