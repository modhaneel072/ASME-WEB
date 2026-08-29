import {
  SCENE_KEYS,
  SCENE_RANGES,
  clamp01,
  getSceneAtProgress,
} from "./camera-path.js";

export class ScrollController {
  constructor({
    root,
    gsap = globalThis.gsap,
    ScrollTrigger = globalThis.ScrollTrigger,
    documentRef = globalThis.document,
    windowRef = globalThis.window,
    reducedMotion = false,
    onProgress = () => {},
    onSceneChange = () => {},
  } = {}) {
    this.root = root || null;
    this.gsap = gsap || null;
    this.ScrollTrigger = ScrollTrigger || null;
    this.document = documentRef || null;
    this.window = windowRef || null;
    this.reducedMotion = reducedMotion;
    this.onProgress = onProgress;
    this.onSceneChange = onSceneChange;
    this.progressElement = this.document?.getElementById?.("landing-progress") || null;
    this.progressLabel = this.document?.getElementById?.("landing-progress-label") || null;
    this.state = { progress: 0 };
    this.progress = 0;
    this.targetProgress = 0;
    this.scene = "intro";
    this.hasCommitted = false;
    this.driver = null;
    this.progressSetter = null;
    this.chapterNodes = [];
    this.chapterMetrics = [];
    this.initialized = false;
    this.usingFallback = false;
    this.abortController = typeof AbortController === "function" ? new AbortController() : null;
  }

  init() {
    if (this.initialized || !this.root || !this.document || !this.window) return this;
    this.initialized = true;
    this.#validateChapterContract();
    this.#measureChapters();
    this.#createDriver();
    this.targetProgress = this.#nativeProgress();
    this.#commit(this.targetProgress, true);
    return this;
  }

  setReducedMotion(reduced) {
    const next = Boolean(reduced);
    if (next === this.reducedMotion) return;
    this.reducedMotion = next;
    if (this.initialized) {
      this.#destroyDriver();
      this.#createDriver();
      this.targetProgress = this.#nativeProgress();
      this.#commit(this.targetProgress, true);
    }
  }

  scrollToProgress(progress, { behavior } = {}) {
    if (!this.window || !this.root) return false;
    const value = clamp01(progress);
    this.targetProgress = value;
    this.#measureChapters();
    this.window.scrollTo({
      top: this.#scrollYForProgress(value),
      behavior: behavior || (this.reducedMotion ? "auto" : "smooth"),
    });
    return true;
  }

  scrollToScene(scene, options) {
    const range = SCENE_RANGES[scene];
    if (!range) return false;
    return this.scrollToProgress(range.start, options);
  }

  refresh() {
    this.#measureChapters();
    if (this.driver?.refresh) this.driver.refresh();
    else {
      this.targetProgress = this.#nativeProgress();
      this.#commit(this.targetProgress, true);
    }
  }

  dispose() {
    if (!this.initialized) return;
    this.#destroyDriver();
    this.abortController?.abort();
    if (!this.abortController) this.window?.removeEventListener?.("scroll", this.#onNativeScroll);
    this.initialized = false;
  }

  #createDriver() {
    if (this.gsap?.quickTo && this.ScrollTrigger?.create) {
      this.gsap.registerPlugin?.(this.ScrollTrigger);
      this.state.progress = this.#nativeProgress();
      this.targetProgress = this.state.progress;
      if (!this.reducedMotion) {
        this.progressSetter = this.gsap.quickTo(this.state, "progress", {
          duration: 0.18,
          ease: "power2.out",
          onUpdate: () => this.#commit(this.state.progress),
        });
      }
      this.driver = this.ScrollTrigger.create({
        id: "landing-master-progress",
        trigger: this.root,
        start: "top top",
        end: "bottom bottom",
        invalidateOnRefresh: true,
        fastScrollEnd: false,
        onUpdate: () => {
          const next = this.#nativeProgress();
          this.targetProgress = next;
          if (this.reducedMotion) {
            this.state.progress = next;
            this.#commit(next);
          } else {
            this.progressSetter(next);
          }
        },
        onRefresh: () => {
          const preserve = this.hasCommitted ? this.targetProgress : null;
          this.#measureChapters();
          const next = preserve ?? this.#nativeProgress();
          if (preserve !== null) {
            this.window.scrollTo({ top: this.#scrollYForProgress(preserve), behavior: "auto" });
          }
          this.state.progress = next;
          this.#commit(next, true);
        },
      });
      this.usingFallback = false;
      return;
    }

    this.usingFallback = true;
    const signal = this.abortController?.signal;
    this.window.addEventListener("scroll", this.#onNativeScroll, signal ? { passive: true, signal } : { passive: true });
    this.window.addEventListener("resize", this.#onResize, signal ? { passive: true, signal } : { passive: true });
  }

  #destroyDriver() {
    this.driver?.kill?.();
    this.progressSetter?.tween?.kill?.();
    this.driver = null;
    this.progressSetter = null;
    if (this.usingFallback) {
      this.window?.removeEventListener?.("scroll", this.#onNativeScroll);
      this.window?.removeEventListener?.("resize", this.#onResize);
    }
    this.usingFallback = false;
  }

  #onNativeScroll = () => {
    this.targetProgress = this.#nativeProgress();
    this.#commit(this.targetProgress);
  };

  #onResize = () => {
    const preserve = this.targetProgress;
    this.#measureChapters();
    this.window.scrollTo({ top: this.#scrollYForProgress(preserve), behavior: "auto" });
    this.#commit(preserve, true);
  };

  #nativeProgress() {
    if (!this.chapterMetrics.length) {
      const { start, distance } = this.#scrollMetrics();
      return clamp01(((Number(this.window?.scrollY) || 0) - start) / distance);
    }

    const y = Number(this.window?.scrollY) || 0;
    const first = this.chapterMetrics[0];
    if (y <= first.top) return first.start;
    for (const chapter of this.chapterMetrics) {
      if (y < chapter.endScroll || chapter === this.chapterMetrics.at(-1)) {
        const local = clamp01((y - chapter.top) / Math.max(1, chapter.endScroll - chapter.top));
        return chapter.start + local * (chapter.end - chapter.start);
      }
    }
    return this.chapterMetrics.at(-1)?.end ?? 1;
  }

  #scrollYForProgress(progress) {
    if (!this.chapterMetrics.length) {
      const { start, distance } = this.#scrollMetrics();
      return start + distance * progress;
    }
    const scene = getSceneAtProgress(progress);
    const chapter = this.chapterMetrics.find((item) => item.scene === scene) || this.chapterMetrics.at(-1);
    const span = Math.max(Number.EPSILON, chapter.end - chapter.start);
    const local = clamp01((progress - chapter.start) / span);
    return chapter.top + (chapter.endScroll - chapter.top) * local;
  }

  #measureChapters() {
    if (!this.chapterNodes.length || !this.window) {
      this.chapterMetrics = [];
      return;
    }
    const scrollY = Number(this.window.scrollY) || 0;
    const measured = this.chapterNodes.map(({ node, scene, start, end }) => {
      const rect = node?.getBoundingClientRect?.();
      if (!rect) return null;
      return {
        scene,
        start,
        end,
        top: scrollY + (Number(rect.top) || 0),
        height: Math.max(1, Number(node.offsetHeight) || Number(rect.height) || 1),
      };
    }).filter(Boolean).sort((a, b) => a.top - b.top);

    if (!measured.length) {
      this.chapterMetrics = [];
      return;
    }
    const documentHeight = Math.max(
      Number(this.document?.documentElement?.scrollHeight) || 0,
      Number(this.document?.body?.scrollHeight) || 0,
      this.#scrollMetrics().start + (Number(this.root?.scrollHeight) || 0),
    );
    const maxScroll = Math.max(measured.at(-1).top + 1, documentHeight - (Number(this.window.innerHeight) || 1));
    this.chapterMetrics = measured.map((chapter, index) => ({
      ...chapter,
      endScroll: index < measured.length - 1 ? measured[index + 1].top : maxScroll,
    }));
  }

  #scrollMetrics() {
    if (!this.root || !this.window) return { start: 0, distance: 1 };
    const rect = this.root.getBoundingClientRect?.() || { top: 0, height: 0 };
    const start = (Number(this.window.scrollY) || 0) + (Number(rect.top) || 0);
    const height = Math.max(Number(this.root.scrollHeight) || 0, Number(this.root.offsetHeight) || 0, Number(rect.height) || 0);
    const distance = Math.max(1, height - (Number(this.window.innerHeight) || 1));
    return { start, distance };
  }

  #commit(progress, force = false) {
    const value = clamp01(progress);
    const scene = getSceneAtProgress(value);
    if (!force && Math.abs(value - this.progress) < 0.00005 && scene === this.scene) return;

    const previousScene = this.scene;
    this.progress = value;
    this.scene = scene;
    this.hasCommitted = true;
    this.#updateDom(value, scene);
    this.onProgress(value, scene);
    dispatchDocumentEvent(this.document, "landing:progress", { progress: value, scene });

    if (scene !== previousScene) {
      this.onSceneChange(previousScene, scene, value);
      dispatchDocumentEvent(this.document, "landing:scenechange", {
        previous: previousScene,
        current: scene,
        progress: value,
      });
    }
  }

  #updateDom(progress, scene) {
    if (this.document?.documentElement) this.document.documentElement.dataset.landingScene = scene;
    if (this.progressElement) {
      this.progressElement.style.setProperty("--landing-progress", progress.toFixed(5));
      this.progressElement.setAttribute("aria-valuenow", String(Math.round(progress * 100)));
    }
    if (this.progressLabel) this.progressLabel.textContent = scene.replaceAll("-", " ");
  }

  #validateChapterContract() {
    const mismatches = [];
    this.chapterNodes = [];
    for (const scene of SCENE_KEYS) {
      const node = this.root.querySelector?.(
        `.landing-chapter[data-scene="${scene}"], section[data-scene="${scene}"]`,
      );
      const expected = SCENE_RANGES[scene];
      if (!node) {
        mismatches.push(`${scene}: missing chapter`);
        continue;
      }
      const start = Number(node.dataset.start);
      const end = Number(node.dataset.end);
      this.chapterNodes.push({
        node,
        scene,
        start: Number.isFinite(start) ? start : expected.start,
        end: Number.isFinite(end) ? end : expected.end,
      });
      if (!Number.isFinite(start) || !Number.isFinite(end)
        || Math.abs(start - expected.start) > 0.0001
        || Math.abs(end - expected.end) > 0.0001) {
        mismatches.push(`${scene}: expected ${expected.start}–${expected.end}`);
      }
    }
    if (mismatches.length) console.warn("[landing] Chapter contract mismatch:", mismatches.join("; "));
  }
}

function dispatchDocumentEvent(documentRef, name, detail) {
  if (!documentRef?.dispatchEvent) return;
  const EventCtor = documentRef.defaultView?.CustomEvent || globalThis.CustomEvent;
  if (typeof EventCtor === "function") documentRef.dispatchEvent(new EventCtor(name, { detail }));
}
