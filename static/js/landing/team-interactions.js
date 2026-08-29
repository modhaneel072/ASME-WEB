import { Vector3 } from "../../vendor/three-0.185.1/three.module.min.js";

export const CHAPTER_TEAM_MAP = Object.freeze({
  iam3d: "iam3d",
  plane: "autonomous-plane",
  rov: "mate-rov",
  "student-design": "student-design",
  "robotic-dog": "robotic-dog",
});

export class TeamInteractions {
  constructor({ root, documentRef = globalThis.document, content = null } = {}) {
    this.document = documentRef || null;
    this.root = root || this.document?.getElementById?.("landing-root") || null;
    this.content = content || { teams: {} };
    this.buttons = new Map();
    this.anchor = new Vector3();
    this.createdLayer = false;
    this.initialized = false;
  }

  init() {
    if (this.initialized || !this.root || !this.document) return this;
    this.initialized = true;
    this.layer = this.root.querySelector?.("[data-landing-hotspots]");
    if (!this.layer) {
      this.layer = this.document.createElement("div");
      this.layer.className = "landing-hotspots";
      this.layer.dataset.landingHotspots = "";
      this.layer.setAttribute("aria-live", "off");
      Object.assign(this.layer.style, {
        position: "fixed",
        inset: "0",
        zIndex: "45",
        pointerEvents: "none",
      });
      this.root.append(this.layer);
      this.createdLayer = true;
    }

    for (const [chapter, team] of Object.entries(CHAPTER_TEAM_MAP)) {
      const button = this.document.createElement("button");
      button.type = "button";
      button.className = "landing-hotspot";
      button.dataset.team = team;
      button.dataset.hotspotScene = chapter;
      button.hidden = true;
      button.setAttribute("aria-haspopup", "dialog");
      button.setAttribute("aria-controls", "team-panel");
      Object.assign(button.style, {
        position: "absolute",
        left: "0",
        top: "0",
        pointerEvents: "auto",
        border: "1px solid rgba(255, 205, 0, .72)",
        background: "rgba(7, 8, 9, .88)",
        color: "#f5f4ef",
        padding: ".55rem .75rem",
        cursor: "pointer",
      });

      const marker = this.document.createElement("span");
      marker.className = "landing-hotspot-marker";
      marker.setAttribute("aria-hidden", "true");
      Object.assign(marker.style, {
        display: "inline-block",
        width: ".5rem",
        height: ".5rem",
        marginRight: ".45rem",
        borderRadius: "50%",
        background: "#ffcd00",
      });
      const label = this.document.createElement("span");
      label.className = "landing-hotspot-label";
      label.textContent = this.#teamName(team);
      button.append(marker, label);
      button.addEventListener("click", () => this.open(team));
      this.layer.append(button);
      this.buttons.set(chapter, button);
    }
    return this;
  }

  setContent(content) {
    this.content = content || { teams: {} };
    for (const [chapter, button] of this.buttons) {
      button.querySelector(".landing-hotspot-label").textContent = this.#teamName(CHAPTER_TEAM_MAP[chapter]);
    }
  }

  update({ scene, camera, anchor, width, height, enabled = true } = {}) {
    const activeButton = this.buttons.get(scene);
    for (const button of this.buttons.values()) button.hidden = button !== activeButton;
    if (!activeButton || !camera || !anchor || !enabled) {
      if (activeButton) activeButton.hidden = true;
      return;
    }

    if (anchor.isObject3D) anchor.getWorldPosition(this.anchor);
    else if (anchor.isVector3) this.anchor.copy(anchor);
    else if (Array.isArray(anchor)) this.anchor.fromArray(anchor);
    else {
      activeButton.hidden = true;
      return;
    }

    this.anchor.project(camera);
    const viewportWidth = Math.max(1, width || globalThis.innerWidth || 1);
    const viewportHeight = Math.max(1, height || globalThis.innerHeight || 1);
    const onScreen = this.anchor.z > -1 && this.anchor.z < 1
      && this.anchor.x > -1.12 && this.anchor.x < 1.12
      && this.anchor.y > -1.12 && this.anchor.y < 1.12;
    activeButton.hidden = !onScreen;
    if (!onScreen) return;

    const left = (this.anchor.x * 0.5 + 0.5) * viewportWidth;
    const top = (-this.anchor.y * 0.5 + 0.5) * viewportHeight;
    activeButton.style.transform = `translate3d(${left.toFixed(1)}px, ${top.toFixed(1)}px, 0)`;
  }

  open(team) {
    if (!team || !Object.values(CHAPTER_TEAM_MAP).includes(team)) return false;
    dispatchDocumentEvent(this.document, "landing:teampanel", { team });
    return true;
  }

  hideAll() {
    for (const button of this.buttons.values()) button.hidden = true;
  }

  dispose() {
    this.hideAll();
    this.buttons.clear();
    if (this.createdLayer) this.layer?.remove?.();
    this.layer = null;
    this.initialized = false;
  }

  #teamName(team) {
    return this.content?.teams?.[team]?.name || team.replaceAll("-", " ");
  }
}

function dispatchDocumentEvent(documentRef, name, detail) {
  if (!documentRef?.dispatchEvent) return;
  const EventCtor = documentRef.defaultView?.CustomEvent || globalThis.CustomEvent;
  if (typeof EventCtor === "function") documentRef.dispatchEvent(new EventCtor(name, { detail }));
}
