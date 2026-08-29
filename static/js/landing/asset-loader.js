import {
  LoadingManager,
  NoColorSpace,
  SRGBColorSpace,
  TextureLoader,
} from "../../vendor/three-0.185.1/three.module.min.js";

const TEAM_KEYS = Object.freeze([
  "iam3d",
  "student-design",
  "autonomous-plane",
  "mate-rov",
  "robotic-dog",
]);

const SAFE_CONTENT = Object.freeze({
  teams: Object.freeze({
    iam3d: Object.freeze({ name: "IAM3D" }),
    "student-design": Object.freeze({ name: "Student Design Competition" }),
    "autonomous-plane": Object.freeze({ name: "Autonomous Plane" }),
    "mate-rov": Object.freeze({ name: "MATE ROV" }),
    "robotic-dog": Object.freeze({
      name: "Robotic Dog",
      status: "Selective team — interview required",
    }),
  }),
});

export class LandingAssetLoader {
  constructor({ documentRef = globalThis.document, timeoutMs = 12000 } = {}) {
    this.document = documentRef || null;
    this.timeoutMs = timeoutMs;
    this.entries = new Map();
    this.assets = new Map();
    this.inflight = new Map();
    this.chapterState = new Map();
    this.errors = [];
    this.completedLoads = 0;
    this.totalLoads = 0;
    this.content = SAFE_CONTENT;
    this.contentSource = "fallback";
    this.disposed = false;
    this.dracoLoader = null;
    this.ktx2Loader = null;
    this.gltfLoaderPromise = null;

    this.manager = new LoadingManager();
    this.manager.onError = (url) => {
      console.warn(`[landing] Three.js could not load asset: ${url}`);
    };
    this.textureLoader = new TextureLoader(this.manager);
  }

  register(entries) {
    for (const entry of entries || []) {
      if (!entry?.id || !entry?.url || !entry?.type) continue;
      this.entries.set(entry.id, Object.freeze({
        required: false,
        chapter: "initial",
        timeoutMs: this.timeoutMs,
        ...entry,
      }));
    }
    return this;
  }

  configureModelDecoders({ dracoLoader = null, ktx2Loader = null } = {}) {
    this.dracoLoader = dracoLoader;
    this.ktx2Loader = ktx2Loader;
  }

  limitChapterEntries(chapter, limit) {
    const keep = Math.max(0, Math.floor(Number(limit) || 0));
    const entries = [...this.entries.values()].filter((entry) => entry.chapter === chapter);
    for (const entry of entries.slice(keep)) {
      if (!this.assets.has(entry.id) && !this.inflight.has(entry.id)) this.entries.delete(entry.id);
    }
    return entries.slice(0, keep).map((entry) => entry.id);
  }

  async loadContent(url = "/static/data/landing-content.json", { emitProgress = true } = {}) {
    this.#assertActive();
    if (emitProgress) this.#dispatchLoading(0, "Loading landing content");
    try {
      const response = await fetchWithTimeout(url, this.timeoutMs);
      if (!response.ok) throw new Error(`Content request returned ${response.status}`);
      const data = normalizeContent(await response.json());
      if (this.disposed) return this.content;
      validateContent(data);
      this.content = data;
      this.contentSource = "repository";
      if (emitProgress) this.#dispatchLoading(this.#progress(), "Landing content ready");
    } catch (error) {
      if (this.disposed) return this.content;
      this.content = SAFE_CONTENT;
      this.contentSource = "fallback";
      this.errors.push({ id: "landing-content", url, error, required: false });
      console.warn("[landing] Content data unavailable; using minimal team labels.", error);
    }
    return this.content;
  }

  async loadInitial() {
    const initialEntries = [...this.entries.values()].filter((entry) => entry.chapter === "initial");
    return this.#loadEntries(initialEntries, "Loading core scene", { emitProgress: true });
  }

  async preloadChapter(chapter) {
    if (!chapter || this.chapterState.get(chapter) === "ready") return this.getChapterAssets(chapter);
    if (this.chapterState.get(chapter) === "loading") {
      const pending = [...this.inflight.values()];
      await Promise.allSettled(pending);
      return this.getChapterAssets(chapter);
    }

    this.chapterState.set(chapter, "loading");
    const entries = [...this.entries.values()].filter((entry) => entry.chapter === chapter);
    const result = await this.#loadEntries(entries, `Loading ${readableChapter(chapter)} scene`, { emitProgress: false });
    this.chapterState.set(chapter, "ready");
    return result;
  }

  async load(id, { emitProgress = true } = {}) {
    this.#assertActive();
    if (this.assets.has(id)) return this.assets.get(id);
    if (this.inflight.has(id)) return this.inflight.get(id);
    const entry = this.entries.get(id);
    if (!entry) throw new Error(`Unknown landing asset: ${id}`);

    this.totalLoads += 1;
    const promise = withTimeout(this.#loadEntry(entry), entry.timeoutMs, entry.url)
      .then((asset) => {
        if (this.disposed) {
          disposeAsset(asset);
          return null;
        }
        this.assets.set(id, asset);
        return asset;
      })
      .catch((error) => {
        if (this.disposed) return null;
        this.errors.push({ id, url: entry.url, error, required: entry.required });
        if (entry.required) throw new Error(`Required asset "${id}" failed: ${readableError(error)}`);
        console.warn(`[landing] Optional asset "${id}" was skipped.`, error);
        return null;
      })
      .finally(() => {
        this.completedLoads += 1;
        this.inflight.delete(id);
        if (!this.disposed && emitProgress) {
          this.#dispatchLoading(this.#progress(), `Loaded ${this.completedLoads} of ${this.totalLoads} scene assets`);
        }
      });

    this.inflight.set(id, promise);
    return promise;
  }

  get(id) {
    return this.assets.get(id) ?? null;
  }

  getChapterAssets(chapter) {
    const result = {};
    for (const entry of this.entries.values()) {
      if (entry.chapter === chapter && this.assets.has(entry.id)) result[entry.id] = this.assets.get(entry.id);
    }
    return result;
  }

  cloneModel(id) {
    const gltf = this.get(id);
    return gltf?.scene?.clone?.(true) || null;
  }

  getSummary() {
    return {
      loaded: this.assets.size,
      registered: this.entries.size,
      errors: this.errors.length,
      contentSource: this.contentSource,
      chapters: Object.fromEntries(this.chapterState),
    };
  }

  dispose() {
    if (this.disposed) return;
    this.disposed = true;
    for (const asset of this.assets.values()) disposeAsset(asset);
    this.assets.clear();
    this.entries.clear();
    this.inflight.clear();
    this.chapterState.clear();
    this.dracoLoader?.dispose?.();
    this.ktx2Loader?.dispose?.();
  }

  async #loadEntries(entries, status, { emitProgress = true } = {}) {
    if (!entries.length) return {};
    if (emitProgress) this.#dispatchLoading(this.#progress(), status);
    const settled = await Promise.allSettled(entries.map((entry) => this.load(entry.id, { emitProgress })));
    const requiredFailure = settled.find((item, index) => item.status === "rejected" && entries[index].required);
    if (requiredFailure) throw requiredFailure.reason;
    return Object.fromEntries(entries.map((entry) => [entry.id, this.get(entry.id)]));
  }

  async #loadEntry(entry) {
    switch (entry.type) {
      case "texture": {
        const texture = await this.textureLoader.loadAsync(entry.url);
        texture.colorSpace = entry.colorSpace === "linear" ? NoColorSpace : SRGBColorSpace;
        texture.flipY = entry.flipY ?? true;
        texture.anisotropy = Math.max(1, Math.min(4, Number(entry.anisotropy) || 1));
        texture.needsUpdate = true;
        return texture;
      }
      case "gltf": {
        const loader = await this.#getGltfLoader();
        return loader.loadAsync(entry.url);
      }
      case "json": {
        const response = await fetchWithTimeout(entry.url, entry.timeoutMs);
        if (!response.ok) throw new Error(`Asset request returned ${response.status}`);
        return response.json();
      }
      case "blob":
      case "audio": {
        const response = await fetchWithTimeout(entry.url, entry.timeoutMs);
        if (!response.ok) throw new Error(`Asset request returned ${response.status}`);
        return response.blob();
      }
      default:
        throw new Error(`Unsupported landing asset type: ${entry.type}`);
    }
  }

  async #getGltfLoader() {
    if (!this.gltfLoaderPromise) {
      this.gltfLoaderPromise = import("../../vendor/three-0.185.1/addons/loaders/GLTFLoader.js")
        .then(({ GLTFLoader }) => {
          const loader = new GLTFLoader(this.manager);
          if (this.dracoLoader) loader.setDRACOLoader(this.dracoLoader);
          if (this.ktx2Loader) loader.setKTX2Loader(this.ktx2Loader);
          return loader;
        });
    }
    return this.gltfLoaderPromise;
  }

  #progress() {
    if (!this.totalLoads) return this.contentSource === "repository" ? 0.12 : 0;
    return Math.min(0.96, 0.12 + (this.completedLoads / this.totalLoads) * 0.84);
  }

  #dispatchLoading(progress, status) {
    dispatchDocumentEvent(this.document, "landing:loading", { progress, status });
  }

  #assertActive() {
    if (this.disposed) throw new Error("Landing asset loader has been disposed");
  }
}

function validateContent(data) {
  if (!data || typeof data !== "object" || !data.teams || typeof data.teams !== "object") {
    throw new Error("landing-content.json must contain a teams object");
  }
  for (const key of TEAM_KEYS) {
    if (!data.teams[key]?.name) throw new Error(`landing-content.json is missing team "${key}"`);
  }
  const status = data.teams["robotic-dog"].status;
  if (status !== "Selective team — interview required") {
    throw new Error("Robotic Dog must use the required selective-team status");
  }
}

function normalizeContent(data) {
  if (!data || typeof data !== "object" || !data.teams || typeof data.teams !== "object") {
    throw new Error("landing-content.json must contain team data");
  }

  const aliases = Object.freeze({
    plane: "autonomous-plane",
    rov: "mate-rov",
  });
  const sourceIsArray = Array.isArray(data.teams);
  const sourceEntries = sourceIsArray
    ? data.teams.map((team) => [team?.id, team])
    : Object.entries(data.teams);
  let changed = sourceIsArray;
  const teams = {};

  for (const [sourceKey, sourceTeam] of sourceEntries) {
    if (!sourceKey || !sourceTeam || typeof sourceTeam !== "object") continue;
    const canonicalKey = aliases[sourceKey] || sourceKey;
    let team = sourceTeam;

    if (canonicalKey !== sourceKey || sourceTeam.id === sourceKey && sourceTeam.id !== canonicalKey) {
      team = { ...team, id: canonicalKey };
      changed = true;
    }

    if (canonicalKey === "robotic-dog"
      && team.status !== "Selective team — interview required"
      && (team.requires_interview === true || /selective|interview/i.test(team.interview_note || ""))) {
      team = { ...team, status: "Selective team — interview required" };
      changed = true;
    }

    teams[canonicalKey] = team;
  }

  if (!changed) return data;
  return { ...data, teams };
}

async function fetchWithTimeout(url, timeoutMs) {
  const controller = typeof AbortController === "function" ? new AbortController() : null;
  const timeout = setTimeout(() => controller?.abort(), timeoutMs);
  try {
    return await fetch(url, controller ? { signal: controller.signal, credentials: "same-origin" } : { credentials: "same-origin" });
  } finally {
    clearTimeout(timeout);
  }
}

function withTimeout(promise, timeoutMs, url) {
  let timeout;
  const timeoutPromise = new Promise((_, reject) => {
    timeout = setTimeout(() => reject(new Error(`Timed out while loading ${url}`)), timeoutMs);
  });
  return Promise.race([promise, timeoutPromise]).finally(() => clearTimeout(timeout));
}

function dispatchDocumentEvent(documentRef, name, detail) {
  if (!documentRef?.dispatchEvent) return;
  const EventCtor = documentRef.defaultView?.CustomEvent || globalThis.CustomEvent;
  if (typeof EventCtor === "function") documentRef.dispatchEvent(new EventCtor(name, { detail }));
}

function disposeAsset(asset) {
  if (!asset) return;
  if (asset.isTexture) {
    asset.dispose();
    return;
  }
  const root = asset.scene || asset;
  root.traverse?.((node) => {
    node.geometry?.dispose?.();
    const materials = Array.isArray(node.material) ? node.material : [node.material];
    for (const material of materials) {
      if (!material) continue;
      for (const value of Object.values(material)) value?.isTexture && value.dispose?.();
      material.dispose?.();
    }
  });
}

function readableChapter(chapter) {
  return String(chapter).replaceAll("-", " ");
}

function readableError(error) {
  return error instanceof Error ? error.message : String(error);
}

export { SAFE_CONTENT, TEAM_KEYS, normalizeContent };
