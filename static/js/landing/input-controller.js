export class InputController {
  constructor({
    documentRef = globalThis.document,
    windowRef = globalThis.window,
    onPauseChange = () => {},
    onReducedMotionChange = () => {},
    onSoundChange = async () => {},
    onPageHide = () => {},
  } = {}) {
    this.document = documentRef || null;
    this.window = windowRef || null;
    this.onPauseChange = onPauseChange;
    this.onReducedMotionChange = onReducedMotionChange;
    this.onSoundChange = onSoundChange;
    this.onPageHide = onPageHide;
    this.abortController = typeof AbortController === "function" ? new AbortController() : null;
    this.mediaQuery = this.window?.matchMedia?.("(prefers-reduced-motion: reduce)") || null;
    this.reducedMotion = Boolean(this.mediaQuery?.matches);
    this.panelOpen = false;
    this.pageHidden = Boolean(this.document?.hidden);
    this.soundEnabled = false;
    this.previousFocus = null;
    this.initialized = false;
  }

  get paused() {
    return this.panelOpen || this.pageHidden;
  }

  init() {
    if (this.initialized || !this.document) return this;
    this.initialized = true;
    const signal = this.abortController?.signal;
    const listenerOptions = signal ? { signal } : undefined;

    this.canvas = this.document.getElementById("landing-canvas");
    this.skipButton = this.document.getElementById("skip-experience");
    this.soundButton = this.document.getElementById("sound-toggle");
    this.teamPanel = this.document.getElementById("team-panel");
    this.teamPanelClose = this.document.getElementById("team-panel-close");

    const MutationObserverCtor = this.window?.MutationObserver || globalThis.MutationObserver;
    if (this.teamPanel && MutationObserverCtor) {
      this.panelObserver = new MutationObserverCtor(this.#onPanelMutation);
      this.panelObserver.observe(this.teamPanel, {
        attributes: true,
        attributeFilter: ["hidden", "aria-hidden"],
      });
      if (this.document.documentElement) {
        this.panelObserver.observe(this.document.documentElement, {
          attributes: true,
          attributeFilter: ["data-team-panel"],
        });
      }
    }

    if (this.canvas) {
      this.canvas.style.pointerEvents = "none";
      this.canvas.setAttribute("aria-hidden", "true");
      this.canvas.tabIndex = -1;
    }

    this.#setSoundButton(false);
    this.skipButton?.addEventListener("click", this.#onSkip, listenerOptions);
    this.soundButton?.addEventListener("click", this.#onSoundToggle, listenerOptions);
    this.teamPanelClose?.addEventListener("click", this.#onPanelClose, listenerOptions);
    this.document.addEventListener("landing:teampanel", this.#onTeamPanel, listenerOptions);
    this.document.addEventListener("keydown", this.#onKeyDown, listenerOptions);
    this.document.addEventListener("visibilitychange", this.#onVisibilityChange, listenerOptions);
    this.window?.addEventListener?.("pagehide", this.#onPageHide, listenerOptions);
    this.window?.addEventListener?.("pageshow", this.#onPageShow, listenerOptions);
    this.window?.addEventListener?.("hashchange", this.#onHashChange, listenerOptions);

    if (this.mediaQuery?.addEventListener) {
      this.mediaQuery.addEventListener("change", this.#onMotionPreference, listenerOptions);
    } else {
      this.mediaQuery?.addListener?.(this.#onMotionPreference);
    }

    this.onReducedMotionChange(this.reducedMotion);
    if (this.document.documentElement) {
      this.document.documentElement.dataset.landingMotion = this.reducedMotion ? "reduced" : "full";
      this.document.documentElement.dataset.landingAudio = "available";
    }
    this.#syncHashRoute();
    return this;
  }

  skipToFinal({ focus = true } = {}) {
    const destination = this.document?.getElementById("landing-final-cta")
      || this.document?.querySelector?.('[data-scene="join"]');
    if (!destination) return false;
    destination.scrollIntoView({ behavior: this.reducedMotion ? "auto" : "smooth", block: "start" });
    if (focus) {
      const focusTarget = destination.querySelector?.("a, button, [tabindex]:not([tabindex='-1'])") || destination;
      if (!focusTarget.hasAttribute?.("tabindex") && focusTarget === destination) focusTarget.tabIndex = -1;
      this.window?.setTimeout?.(() => focusTarget.focus?.({ preventScroll: true }), this.reducedMotion ? 0 : 450);
    }
    return true;
  }

  closeTeamPanel({ restoreFocus = true } = {}) {
    if (!this.panelOpen) return;
    this.panelOpen = false;
    if (this.teamPanel) {
      this.teamPanel.hidden = true;
      this.teamPanel.setAttribute("aria-hidden", "true");
      delete this.teamPanel.dataset.team;
    }
    this.onPauseChange(this.paused);
    if (restoreFocus) this.previousFocus?.focus?.({ preventScroll: true });
    this.previousFocus = null;
  }

  dispose() {
    if (!this.initialized) return;
    this.abortController?.abort();
    this.panelObserver?.disconnect?.();
    this.panelObserver = null;
    if (!this.abortController) {
      this.skipButton?.removeEventListener("click", this.#onSkip);
      this.soundButton?.removeEventListener("click", this.#onSoundToggle);
      this.teamPanelClose?.removeEventListener("click", this.#onPanelClose);
      this.document?.removeEventListener("landing:teampanel", this.#onTeamPanel);
      this.document?.removeEventListener("keydown", this.#onKeyDown);
      this.document?.removeEventListener("visibilitychange", this.#onVisibilityChange);
      this.window?.removeEventListener?.("pagehide", this.#onPageHide);
      this.window?.removeEventListener?.("pageshow", this.#onPageShow);
      this.window?.removeEventListener?.("hashchange", this.#onHashChange);
      this.mediaQuery?.removeEventListener?.("change", this.#onMotionPreference);
      this.mediaQuery?.removeListener?.(this.#onMotionPreference);
    }
    this.initialized = false;
  }

  #onSkip = (event) => {
    event?.preventDefault?.();
    this.skipToFinal();
  };

  #onSoundToggle = async () => {
    const requested = !this.soundEnabled;
    this.soundButton?.setAttribute("aria-busy", "true");
    try {
      await this.onSoundChange(requested);
      this.soundEnabled = requested;
      this.#setSoundButton(requested);
    } catch (error) {
      console.warn("[landing] Audio could not be enabled.", error);
      this.soundEnabled = false;
      this.#setSoundButton(false);
    } finally {
      this.soundButton?.removeAttribute("aria-busy");
    }
  };

  #onTeamPanel = (event) => {
    const team = event?.detail?.team;
    if (!team || !this.teamPanel) return;
    this.previousFocus = this.document.activeElement;
    this.panelOpen = true;
    this.teamPanel.dataset.team = team;
    this.teamPanel.hidden = false;
    this.teamPanel.setAttribute("aria-hidden", "false");
    this.onPauseChange(this.paused);
    this.teamPanelClose?.focus?.({ preventScroll: true });
  };

  #onPanelMutation = () => {
    const shellState = this.document?.documentElement?.dataset?.teamPanel;
    const visible = Boolean(this.teamPanel && !this.teamPanel.hidden && shellState !== "closed");
    const ariaHidden = String(!visible);
    if (this.teamPanel?.getAttribute?.("aria-hidden") !== ariaHidden) {
      this.teamPanel.setAttribute("aria-hidden", ariaHidden);
    }
    if (visible === this.panelOpen) return;
    this.panelOpen = visible;
    if (visible) this.previousFocus ||= this.document?.activeElement || null;
    else this.previousFocus = null;
    this.onPauseChange(this.paused);
  };

  #onPanelClose = () => this.closeTeamPanel();

  #onKeyDown = (event) => {
    if (event.key === "Escape" && this.panelOpen) {
      event.preventDefault();
      this.closeTeamPanel();
    }
    // Wheel, touch, arrows, Page Up/Down, Home, End, and Space deliberately
    // remain browser-native. The landing never replaces the real scroller.
  };

  #onVisibilityChange = () => {
    this.pageHidden = Boolean(this.document?.hidden);
    this.onPauseChange(this.paused);
  };

  #onMotionPreference = (event) => {
    this.reducedMotion = Boolean(event.matches);
    if (this.document?.documentElement) {
      this.document.documentElement.dataset.landingMotion = this.reducedMotion ? "reduced" : "full";
    }
    this.onReducedMotionChange(this.reducedMotion);
  };

  #onPageHide = (event) => {
    this.pageHidden = true;
    this.onPauseChange(this.paused);
    this.onPageHide(event);
  };

  #onPageShow = () => {
    this.pageHidden = Boolean(this.document?.hidden);
    this.onPauseChange(this.paused);
  };

  #onHashChange = () => this.#syncHashRoute();

  #syncHashRoute() {
    const hash = this.window?.location?.hash;
    if (hash === "#landing-final-cta" || hash === "#join") {
      this.window?.requestAnimationFrame?.(() => this.skipToFinal({ focus: false }));
    }
  }

  #setSoundButton(enabled) {
    if (!this.soundButton) return;
    this.soundButton.setAttribute("aria-pressed", String(enabled));
    this.soundButton.dataset.sound = enabled ? "on" : "off";
    const label = enabled ? "Mute landing audio" : "Enable landing audio";
    this.soundButton.setAttribute("aria-label", label);
  }
}
