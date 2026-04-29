/* =====================================================================
   PRECISION PULSE — motion system for the ASME flora-preview
   One cohesive language: scan, draw, magnetic, traced, spotlight.
   ---------------------------------------------------------------------
   Tokens (kept in JS so the easing values match the brief):
     ENTRANCE  cubic-bezier(0.22, 1, 0.36, 1)
     SNAP      cubic-bezier(0.34, 1.56, 0.64, 1)
     LINEAR    none
   ===================================================================== */
gsap.registerPlugin(ScrollTrigger);

const reduced = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

const EASE = {
  entrance: "cubic-bezier(0.22, 1, 0.36, 1)",
  snap:     "cubic-bezier(0.34, 1.56, 0.64, 1)",
};

/* ─────────── Helpers ─────────── */
const $  = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
const lerp = (a, b, t) => a + (b - a) * t;

/* ─────────── Split section H2s into clip-line segments ───────────
   So heading reveals match the hero's split-line treatment instead of
   another fade-up. We do this once at boot — pure DOM wrapping. */
function splitLines(el) {
  if (!el || el.dataset.split === "1") return;
  const text = el.textContent.trim();
  // We only split on word groups for safety; Flora's H2s are one line each.
  el.innerHTML = `<span class="clip-line"><span class="line-inner">${text}</span></span>`;
  el.dataset.split = "1";
}
$$(".section-heading h2").forEach(splitLines);
$$(".cta-card h2").forEach(splitLines);

/* ─────────── Reduced motion: show final state and exit ─────────── */
if (reduced) {
  $$(".reveal-up, .reveal-card, .reveal-left, .wipe-reveal, .clip-line > span")
    .forEach((el) => { el.style.opacity = "1"; el.style.transform = "none"; });
} else {

/* ─────────── Initial states (one place) ─────────── */
gsap.set(".reveal-up",        { y: 28, opacity: 0 });
gsap.set(".reveal-card",      { y: 38, opacity: 0, rotateX: 6, transformPerspective: 1100 });
gsap.set(".reveal-left",      { x: -36, opacity: 0 });
gsap.set(".wipe-reveal",      { y: 22, opacity: 0 });
gsap.set(".clip-line > span", { yPercent: 110, opacity: 1 });
gsap.set(".timeline-line",    { scaleY: 0, transformOrigin: "top center" });
gsap.set(".timeline-dot",     { scale: 0, opacity: 0 });
gsap.set(".bp-corner",        { scale: 0, opacity: 0, transformOrigin: "center center" });
gsap.set(".node",             { scale: 0, opacity: 0 });
gsap.set(".tech-diagram .line", { scaleX: 0, transformOrigin: "left center" });
gsap.set(".sponsor-track",    { x: 0 });

/* =====================================================================
   1 · HERO — interface comes online
   ===================================================================== */
const heroTl = gsap.timeline({ defaults: { ease: EASE.entrance } });

heroTl
  // Faint blueprint grid fades in
  .fromTo(".hero-grid", { opacity: 0 }, { opacity: 0.55, duration: 0.9 })
  // Scan beam sweeps top → bottom, revealing the hero
  .fromTo(".scan-line",
    { yPercent: -100, opacity: 0.0 },
    { yPercent: 220, opacity: 0.9, duration: 1.4, ease: "power1.inOut" },
    "<")
  // Split-line title (each line slides up out of its clip)
  .to(".clip-line > span", {
    yPercent: 0, duration: 0.85, stagger: 0.10, ease: EASE.entrance,
  }, "-=1.10")
  // Supporting copy after the title
  .to(".reveal-up", { y: 0, opacity: 1, duration: 0.7, stagger: 0.10 }, "-=0.55")
  // Hero panel snaps into place with subtle tilt resolve
  .fromTo(".hero-visual",
    { y: 36, opacity: 0, rotateY: -6 },
    { y: 0, opacity: 1, rotateY: 0, duration: 0.95 },
    "-=0.55")
  // Blueprint corner ticks pop after the panel arrives (mech-snap)
  .to(".bp-corner", {
    scale: 1, opacity: 1, duration: 0.5, stagger: 0.05, ease: EASE.snap,
  }, "-=0.40")
  // Diagram nodes pulse in, then connecting lines trace
  .to(".node", { scale: 1, opacity: 1, duration: 0.4, stagger: 0.08, ease: EASE.snap }, "-=0.55")
  .to(".tech-diagram .line", { scaleX: 1, duration: 0.55, stagger: 0.08, ease: EASE.entrance }, "-=0.40");

/* Idle hero ambience — subdued so it never competes with content */
gsap.to(".scan-line", {
  yPercent: 220, duration: 4.4, ease: "power1.inOut", repeat: -1, repeatDelay: 1.6,
  delay: heroTl.duration() + 0.6,
});
gsap.to(".beam",   { y: 280, duration: 5.0, ease: "sine.inOut", repeat: -1, yoyo: true });
gsap.to(".ring-1", { rotate:  360, duration: 24, ease: "none", repeat: -1 });
gsap.to(".ring-2", { rotate: -360, duration: 18, ease: "none", repeat: -1 });
gsap.to(".node",   { scale: 1.10, duration: 1.6, ease: "sine.inOut", repeat: -1, yoyo: true, stagger: 0.18, delay: 1.2 });
gsap.to(".tech-diagram .line", { opacity: 0.55, duration: 1.6, repeat: -1, yoyo: true, ease: "sine.inOut" });

/* =====================================================================
   2 · SECTION TRANSITIONS — wipe / weld instead of fade
   - Headings: split-line reveal (matches hero language)
   - Section kicker: traced underline draws in
   ===================================================================== */
$$(".section-heading").forEach((heading) => {
  const tl = gsap.timeline({
    scrollTrigger: { trigger: heading, start: "top 82%", once: true },
    defaults: { ease: EASE.entrance },
  });
  const lineSpan = heading.querySelector(".clip-line > span");
  const kicker   = heading.querySelector(".section-kicker");
  const body     = heading.querySelector("p:not(.section-kicker)");
  if (kicker)   tl.fromTo(kicker, { opacity: 0, x: -16 }, { opacity: 1, x: 0, duration: 0.5 });
  if (lineSpan) tl.to(lineSpan, { yPercent: 0, duration: 0.85 }, "-=0.20");
  if (body)     tl.fromTo(body,   { opacity: 0, y: 16 }, { opacity: 1, y: 0, duration: 0.6 }, "-=0.45");
});

/* CTA wipe-reveal cards: directional slide + clip wash */
$$(".wipe-reveal").forEach((el) => {
  if (el.classList.contains("section-heading")) return; // already handled
  gsap.to(el, {
    y: 0, opacity: 1, duration: 0.85, ease: EASE.entrance,
    scrollTrigger: { trigger: el, start: "top 84%", once: true },
  });
});

/* =====================================================================
   3 · CARDS — measured staggered waves with subtle 3D resolve
   ===================================================================== */
$$(".feature-grid, .project-grid, .team-grid").forEach((grid) => {
  const cards = $$(".reveal-card", grid);
  if (!cards.length) return;
  gsap.to(cards, {
    y: 0, opacity: 1, rotateX: 0,
    duration: 0.75, ease: EASE.entrance,
    stagger: { each: 0.08, from: "start" },
    scrollTrigger: { trigger: grid, start: "top 84%", once: true },
  });
});

/* Lone reveal-cards not inside a grid (edge case) */
$$(".reveal-card").forEach((el) => {
  if (el.closest(".feature-grid, .project-grid, .team-grid")) return;
  gsap.to(el, {
    y: 0, opacity: 1, rotateX: 0, duration: 0.7, ease: EASE.entrance,
    scrollTrigger: { trigger: el, start: "top 88%", once: true },
  });
});

/* =====================================================================
   4 · TIMELINE — spine draws on scroll, dots pulse in, cards stagger
   ===================================================================== */
const timelineEl = $(".timeline");
if (timelineEl) {
  // Spine drawn from top to bottom, scrubbed to scroll
  gsap.to(".timeline-line", {
    scaleY: 1, ease: "none",
    scrollTrigger: {
      trigger: timelineEl,
      start: "top 70%",
      end: "bottom 80%",
      scrub: 0.8,
    },
  });

  // Dots pulse in as their card crosses the trigger line
  $$(".timeline-item").forEach((item) => {
    const dot = item.querySelector(".timeline-dot");
    const tl = gsap.timeline({
      scrollTrigger: { trigger: item, start: "top 86%", once: true },
      defaults: { ease: EASE.snap },
    });
    if (dot) tl.to(dot, { scale: 1, opacity: 1, duration: 0.45 });
    tl.to(item, { x: 0, opacity: 1, duration: 0.65, ease: EASE.entrance }, "-=0.20");
  });
}

/* =====================================================================
   5 · SPONSOR MARQUEE — continuous, paused on hover
   ===================================================================== */
const track = $(".sponsor-track");
if (track) {
  // Track contains 8 pills (4 + duplicate). We move by half its width to loop seamlessly.
  // Compute on layout; reset on resize.
  let marqueeTween = null;
  function startMarquee() {
    if (marqueeTween) marqueeTween.kill();
    gsap.set(track, { x: 0 });
    const total = track.scrollWidth / 2;
    marqueeTween = gsap.to(track, {
      x: -total,
      duration: total / 60,           // ~60px / sec — calm, premium pace
      ease: "none",
      repeat: -1,
    });
  }
  startMarquee();
  let r = 0;
  window.addEventListener("resize", () => {
    clearTimeout(r);
    r = setTimeout(startMarquee, 200);
  });

  const row = $(".sponsor-row");
  if (row) {
    row.addEventListener("mouseenter", () => marqueeTween && marqueeTween.timeScale(0.15));
    row.addEventListener("mouseleave", () => marqueeTween && marqueeTween.timeScale(1));
  }
}

/* =====================================================================
   6 · MAGNETIC BUTTONS / BRAND
   - Subtle attraction toward cursor (lerp, not 1:1)
   - Spring-back on leave
   ===================================================================== */
$$(".magnetic").forEach((el) => {
  let raf = 0, tx = 0, ty = 0, cx = 0, cy = 0;
  const STRENGTH = 0.30;          // px per px of cursor offset within bounds
  const RANGE = 60;               // clamp so we never fly off
  function tick() {
    cx = lerp(cx, tx, 0.18);
    cy = lerp(cy, ty, 0.18);
    el.style.transform = `translate3d(${cx.toFixed(2)}px, ${cy.toFixed(2)}px, 0)`;
    if (Math.abs(cx - tx) > 0.05 || Math.abs(cy - ty) > 0.05) {
      raf = requestAnimationFrame(tick);
    } else { raf = 0; }
  }
  el.addEventListener("mousemove", (e) => {
    const r = el.getBoundingClientRect();
    const dx = (e.clientX - (r.left + r.width / 2)) * STRENGTH;
    const dy = (e.clientY - (r.top  + r.height / 2)) * STRENGTH;
    tx = Math.max(-RANGE, Math.min(RANGE, dx));
    ty = Math.max(-RANGE, Math.min(RANGE, dy));
    if (!raf) raf = requestAnimationFrame(tick);
  });
  el.addEventListener("mouseleave", () => {
    tx = 0; ty = 0;
    // Snap-spring on release using the snap easing
    gsap.to(el, { x: 0, y: 0, duration: 0.55, ease: EASE.snap, clearProps: "transform" });
    if (raf) { cancelAnimationFrame(raf); raf = 0; }
  });
});

/* =====================================================================
   7 · TILT CARDS — restrained 3D resolve on hover
   - Max 6° tilt, eased; subtle thermal glow tracks cursor via CSS vars
   ===================================================================== */
$$(".tilt-card").forEach((card) => {
  const MAX = 6;
  let rx = 0, ry = 0, gx = 50, gy = 50;
  let tg_rx = 0, tg_ry = 0, raf = 0;
  function tick() {
    rx = lerp(rx, tg_rx, 0.15);
    ry = lerp(ry, tg_ry, 0.15);
    card.style.setProperty("--rx", rx.toFixed(2) + "deg");
    card.style.setProperty("--ry", ry.toFixed(2) + "deg");
    card.style.setProperty("--gx", gx.toFixed(1) + "%");
    card.style.setProperty("--gy", gy.toFixed(1) + "%");
    if (Math.abs(rx - tg_rx) > 0.05 || Math.abs(ry - tg_ry) > 0.05) {
      raf = requestAnimationFrame(tick);
    } else raf = 0;
  }
  card.addEventListener("mousemove", (e) => {
    const r = card.getBoundingClientRect();
    const px = (e.clientX - r.left) / r.width;     // 0..1
    const py = (e.clientY - r.top)  / r.height;
    tg_ry =  (px - 0.5) * 2 * MAX;
    tg_rx = -(py - 0.5) * 2 * MAX;
    gx = px * 100; gy = py * 100;
    card.classList.add("is-tilting");
    if (!raf) raf = requestAnimationFrame(tick);
  });
  card.addEventListener("mouseleave", () => {
    tg_rx = 0; tg_ry = 0;
    card.classList.remove("is-tilting");
    gsap.to(card, {
      "--rx": "0deg", "--ry": "0deg",
      duration: 0.55, ease: EASE.snap,
    });
  });
});

/* =====================================================================
   8 · CURSOR GLOW — fixed thermal follower
   ===================================================================== */
const glow = $(".cursor-glow");
if (glow) {
  let mx = window.innerWidth / 2, my = window.innerHeight / 2;
  let cx = mx, cy = my, raf = 0;
  function tick() {
    cx = lerp(cx, mx, 0.18);
    cy = lerp(cy, my, 0.18);
    glow.style.left = cx + "px";
    glow.style.top  = cy + "px";
    raf = requestAnimationFrame(tick);
  }
  window.addEventListener("mousemove", (e) => { mx = e.clientX; my = e.clientY; });
  raf = requestAnimationFrame(tick);
}

/* =====================================================================
   9 · PARALLAX LAYERS — hero panel responds to mouse
   ===================================================================== */
$$(".parallax-layer").forEach((layer) => {
  const depth = parseFloat(layer.dataset.depth || "0.10");
  let tx = 0, ty = 0, cx = 0, cy = 0, raf = 0;
  function tick() {
    cx = lerp(cx, tx, 0.10);
    cy = lerp(cy, ty, 0.10);
    layer.style.transform =
      `translate3d(${cx.toFixed(2)}px, ${cy.toFixed(2)}px, 0)`;
    if (Math.abs(cx - tx) > 0.04 || Math.abs(cy - ty) > 0.04) {
      raf = requestAnimationFrame(tick);
    } else raf = 0;
  }
  window.addEventListener("mousemove", (e) => {
    const x = (e.clientX / window.innerWidth  - 0.5) * 2;
    const y = (e.clientY / window.innerHeight - 0.5) * 2;
    tx = -x * 30 * depth; ty = -y * 22 * depth;
    if (!raf) raf = requestAnimationFrame(tick);
  });
});

/* Subtle scroll-tied parallax on the hero blueprint grid */
gsap.to(".hero-grid", {
  backgroundPosition: "120px 80px",
  ease: "none",
  scrollTrigger: { trigger: "#hero", scrub: 1.4, start: "top top", end: "bottom top" },
});

/* =====================================================================
   10 · TRACED LINKS — underline draws on hover (CSS handles the visual,
        JS just toggles class for accessibility/touch parity)
   ===================================================================== */
$$(".text-link, .nav-links a, .site-footer a").forEach((a) => {
  if (!a.querySelector(".tl-bar")) {
    const bar = document.createElement("span");
    bar.className = "tl-bar";
    a.appendChild(bar);
  }
});

}  /* end !reduced */
