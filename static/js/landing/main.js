import { LandingAssetLoader } from "./asset-loader.js";
import { LandingDebug } from "./debug.js";
import { InputController } from "./input-controller.js";
import { QualityManager } from "./quality-manager.js";
import { SceneController } from "./scene-controller.js";
import { ScrollController } from "./scroll-controller.js";
import { TeamInteractions } from "./team-interactions.js";

const THREE_VERSION = "0.185.1";
const GSAP_VERSION = "3.15.0";
const TEAM_SCENES = new Set(["iam3d", "plane", "rov", "student-design", "robotic-dog"]);

let activeLanding = null;

export async function bootstrapLanding({ documentRef = globalThis.document, windowRef = globalThis.window } = {}) {
  if (activeLanding) return activeLanding;
  const document = documentRef;
  const window = windowRef;
  const root = document?.getElementById?.("landing-root");
  const canvas = document?.getElementById?.("landing-canvas");

  if (!document || !window || !root || !canvas) {
    const missing = [!root && "#landing-root", !canvas && "#landing-canvas"].filter(Boolean).join(" and ");
    const error = new Error(`Landing DOM contract is incomplete: missing ${missing || "document/window"}`);
    reportStandaloneError(document, error);
    return null;
  }

  document.documentElement.dataset.landingReady = "false";
  document.documentElement.dataset.landingScene = "intro";
  document.documentElement.dataset.webgl = "pending";
  root.classList.remove("is-ready", "is-fallback");
  dispatch(document, "landing:loading", { progress: 0, status: "Loading scene" });

  const loaderUi = bindLoaderUi(document);
  const state = {
    document,
    window,
    root,
    canvas,
    disposed: false,
    readyDispatched: false,
    errorReported: false,
    preloadPromise: null,
  };
  activeLanding = state;

  const quality = new QualityManager({ canvas, windowRef: window, documentRef: document });
  state.quality = quality;
  const assets = new LandingAssetLoader({ documentRef: document });
  state.assets = assets;
  const photoManifest = [
    { id: "people-electronics", type: "texture", url: "/static/images/gallery/electronics_lab.jpg", chapter: "people" },
    { id: "people-presentation", type: "texture", url: "/static/images/gallery/team_presenting.jpg", chapter: "people" },
    { id: "people-workshop", type: "texture", url: "/static/images/makeathon/IMG_3507.jpg", chapter: "people" },
    { id: "people-prototype", type: "texture", url: "/static/images/makeathon/IMG_5970.jpg", chapter: "people" },
  ].slice(0, quality.profile.photoLayers);
  assets.register(photoManifest);

  const fallback = (error, { ready = true, disposeAssets = true } = {}) => {
    state.fallbackActive = true;
    const message = normalizeError(error);
    if (!state.errorReported && error) {
      state.errorReported = true;
      console.error("[landing] Activating the accessible fallback.", error);
      dispatch(document, "landing:error", { message });
    }
    state.scene?.dispose?.();
    state.scene = null;
    if (disposeAssets) state.assets?.dispose?.();
    state.teamInteractions?.hideAll?.();
    canvas.hidden = true;
    canvas.style.display = "none";
    const webglRoot = document.getElementById("landing-webgl");
    if (webglRoot) {
      webglRoot.hidden = true;
      webglRoot.setAttribute("aria-hidden", "true");
    }
    const fallbackElement = document.getElementById("landing-fallback");
    if (fallbackElement) {
      fallbackElement.hidden = false;
      fallbackElement.setAttribute("aria-hidden", "false");
    }
    root.classList.add("is-fallback");
    root.classList.remove("is-ready");
    document.documentElement.dataset.webgl = "fallback";
    document.documentElement.dataset.landingReady = "true";
    loaderUi.finish();
    if (ready && !state.readyDispatched) {
      state.readyDispatched = true;
      dispatch(document, "landing:ready", { quality: "fallback" });
    }
  };
  state.fallback = fallback;

  if (quality.isFallback) {
    const contentPromise = assets.loadContent(undefined, { emitProgress: false });
    state.content = assets.content;
    createNonWebglControllers(state, state.content);
    // There are no registered WebGL assets in fallback mode, so keep the
    // content request alive while revealing the usable DOM immediately.
    fallback(new Error("WebGL2 is unavailable; showing the accessible landing experience."), { disposeAssets: false });
    contentPromise.then((content) => {
      if (state.disposed) return;
      state.content = content;
      state.teamInteractions?.setContent?.(content);
    });
    return state;
  }

  try {
    const [content] = await Promise.all([
      assets.loadContent(),
      assets.loadInitial(),
      loadGsapLibraries(document, window),
    ]);
    state.content = content;

    const debug = new LandingDebug({ documentRef: document, windowRef: window }).init();
    state.debug = debug;
    const teamInteractions = new TeamInteractions({ root, documentRef: document, content }).init();
    state.teamInteractions = teamInteractions;

    const scene = new SceneController({
      canvas,
      root,
      qualityManager: quality,
      documentRef: document,
      windowRef: window,
      onError: (error) => fallback(error),
      onFrame: (frame) => onFrame(state, frame),
    }).init();
    state.scene = scene;

    const input = new InputController({
      documentRef: document,
      windowRef: window,
      onPauseChange: (paused) => scene.setPaused(paused),
      onReducedMotionChange: (reduced) => {
        scene.setReducedMotion(reduced);
        state.scroll?.setReducedMotion(reduced);
      },
      onSoundChange: async (enabled) => {
        // Audio assets are intentionally absent. This interface is ready to
        // lazy-load registered audio only after an explicit user gesture.
        state.audioEnabled = enabled;
      },
      onPageHide: (event) => {
        if (!event?.persisted) disposeLanding(state);
      },
    }).init();
    state.input = input;

    const scroll = new ScrollController({
      root,
      gsap: window.gsap,
      ScrollTrigger: window.ScrollTrigger,
      documentRef: document,
      windowRef: window,
      reducedMotion: input.reducedMotion,
      onProgress: (progress) => scene.setProgress(progress),
      onSceneChange: (_previous, current) => preloadAhead(state, current),
    }).init();
    state.scroll = scroll;
    scene.setReducedMotion(input.reducedMotion);
    scene.setProgress(scroll.progress);

    const onQualityChange = (event) => {
      const next = event.detail?.current;
      if (next === "fallback") fallback(new Error("Rendering performance remained below the safe threshold"));
      else {
        scene.setQuality(next);
        assets.limitChapterEntries("people", quality.profile.photoLayers);
      }
    };
    state.onQualityChange = onQualityChange;
    document.addEventListener("landing:qualitychange", onQualityChange);

    canvas.hidden = false;
    canvas.style.removeProperty("display");
    const webglRoot = document.getElementById("landing-webgl");
    if (webglRoot) {
      webglRoot.hidden = false;
      webglRoot.setAttribute("aria-hidden", "true");
    }
    const fallbackElement = document.getElementById("landing-fallback");
    if (fallbackElement) {
      fallbackElement.hidden = true;
      fallbackElement.setAttribute("aria-hidden", "true");
    }
    root.classList.add("is-ready");
    document.documentElement.dataset.webgl = "available";
    document.documentElement.dataset.landingReady = "true";
    loaderUi.finish();
    scene.start();
    state.readyDispatched = true;
    dispatch(document, "landing:ready", { quality: quality.quality });
    preloadAhead(state, scroll.scene);
    return state;
  } catch (error) {
    fallback(error);
    createNonWebglControllers(state, state.content || assets.content);
    return state;
  }
}

function createNonWebglControllers(state, content) {
  if (!state.root || !state.document || !state.window) return;
  state.input ||= new InputController({
    documentRef: state.document,
    windowRef: state.window,
    onPageHide: (event) => {
      if (!event?.persisted) disposeLanding(state);
    },
  }).init();
  state.scroll ||= new ScrollController({
    root: state.root,
    documentRef: state.document,
    windowRef: state.window,
    reducedMotion: state.input.reducedMotion,
  }).init();
  state.teamInteractions ||= new TeamInteractions({
    root: state.root,
    documentRef: state.document,
    content,
  }).init();
  state.teamInteractions.hideAll();
}

function onFrame(state, frame) {
  if (state.disposed || !state.scene) return;
  const before = state.quality.quality;
  const after = state.quality.observeFrame(frame.deltaMs);
  if (before !== after || after === "fallback") return;

  const sceneKey = state.scroll?.scene || frame.scene;
  state.teamInteractions?.update?.({
    scene: sceneKey,
    camera: state.scene.camera,
    anchor: state.scene.getHotspotAnchor(sceneKey),
    width: frame.viewport?.width,
    height: frame.viewport?.height,
    enabled: TEAM_SCENES.has(sceneKey) && !state.input?.paused,
  });
  state.debug?.update?.({
    progress: state.scroll?.progress ?? frame.progress,
    scene: sceneKey,
    chapter: sceneKey,
    quality: state.quality.quality,
    assets: state.assets.getSummary(),
    drawCalls: frame.drawCalls,
    triangles: frame.triangles,
    camera: frame.camera,
    target: frame.target,
  });
}

function preloadAhead(state, currentScene) {
  if (state.disposed || state.fallbackActive || !state.scene || state.preloadPromise || currentScene === "join") return;
  if (currentScene !== "robotic-dog" && currentScene !== "people") return;
  state.preloadPromise = state.assets.preloadChapter("people")
    .then((chapterAssets) => state.scene?.setChapterAssets?.("people", chapterAssets))
    .catch((error) => console.warn("[landing] Workshop photos were skipped; the scene will use its editorial frame fallback.", error))
    .finally(() => { state.preloadPromise = null; });
}

function bindLoaderUi(document) {
  const loader = document.getElementById("landing-loader");
  const progress = document.getElementById("landing-loader-progress");
  const status = document.getElementById("landing-loader-status");
  const onLoading = (event) => {
    const value = Math.max(0, Math.min(1, Number(event.detail?.progress) || 0));
    if (progress) {
      if ("value" in progress) {
        progress.max = 1;
        progress.value = value;
      }
      progress.style.setProperty("--landing-loader-progress", value.toFixed(4));
      progress.style.setProperty("--loader-progress", value.toFixed(4));
      progress.setAttribute("aria-valuenow", String(Math.round(value * 100)));
    }
    loader?.style?.setProperty?.("--landing-loader-progress", value.toFixed(4));
    loader?.style?.setProperty?.("--loader-progress", value.toFixed(4));
    if (status && event.detail?.status) status.textContent = event.detail.status;
  };
  document.addEventListener("landing:loading", onLoading);
  return {
    finish() {
      document.removeEventListener("landing:loading", onLoading);
      if (loader) {
        loader.hidden = true;
        loader.setAttribute("aria-hidden", "true");
      }
    },
  };
}

async function loadGsapLibraries(document, window) {
  if (window.gsap?.version !== GSAP_VERSION) {
    await loadClassicScript(document, `/static/vendor/gsap-${GSAP_VERSION}/gsap.min.js`, `gsap-${GSAP_VERSION}`);
  }
  if (window.ScrollTrigger?.version !== GSAP_VERSION) {
    await loadClassicScript(document, `/static/vendor/gsap-${GSAP_VERSION}/ScrollTrigger.min.js`, `scroll-trigger-${GSAP_VERSION}`);
  }
  if (!window.gsap || !window.ScrollTrigger) throw new Error("Pinned GSAP and ScrollTrigger libraries did not initialize");
  window.gsap.registerPlugin(window.ScrollTrigger);
}

function loadClassicScript(document, src, id) {
  const existing = document.querySelector(`script[data-landing-vendor="${id}"]`);
  if (existing?.dataset.loaded === "true") return Promise.resolve();
  return new Promise((resolve, reject) => {
    const script = existing || document.createElement("script");
    script.src = src;
    script.async = false;
    script.dataset.landingVendor = id;
    const cleanup = () => {
      script.removeEventListener("load", onLoad);
      script.removeEventListener("error", onError);
    };
    const onLoad = () => {
      script.dataset.loaded = "true";
      cleanup();
      resolve();
    };
    const onError = () => {
      cleanup();
      reject(new Error(`Could not load ${src}`));
    };
    script.addEventListener("load", onLoad);
    script.addEventListener("error", onError);
    if (!existing) document.head.append(script);
  });
}

export function disposeLanding(state = activeLanding) {
  if (!state || state.disposed) return;
  state.disposed = true;
  state.document?.removeEventListener?.("landing:qualitychange", state.onQualityChange);
  state.debug?.dispose?.();
  state.teamInteractions?.dispose?.();
  state.scroll?.dispose?.();
  state.input?.dispose?.();
  state.scene?.dispose?.();
  state.assets?.dispose?.();
  state.quality?.dispose?.();
  if (activeLanding === state) activeLanding = null;
}

function reportStandaloneError(document, error) {
  console.error("[landing]", error);
  if (document?.documentElement) {
    document.documentElement.dataset.webgl = "fallback";
    document.documentElement.dataset.landingReady = "true";
  }
  const fallback = document?.getElementById?.("landing-fallback");
  if (fallback) fallback.hidden = false;
  dispatch(document, "landing:error", { message: normalizeError(error) });
  dispatch(document, "landing:ready", { quality: "fallback" });
}

function dispatch(document, name, detail) {
  if (!document?.dispatchEvent) return;
  const EventCtor = document.defaultView?.CustomEvent || globalThis.CustomEvent;
  if (typeof EventCtor === "function") document.dispatchEvent(new EventCtor(name, { detail }));
}

function normalizeError(error) {
  if (!error) return "WebGL is unavailable; showing the accessible landing experience.";
  if (error instanceof Error && error.message) return error.message;
  return String(error);
}

if (typeof document !== "undefined" && typeof window !== "undefined") {
  const start = () => bootstrapLanding().catch((error) => reportStandaloneError(document, error));
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start, { once: true });
  else start();
}

export const LANDING_LIBRARY_VERSIONS = Object.freeze({ three: THREE_VERSION, gsap: GSAP_VERSION });
