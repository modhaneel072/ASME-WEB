(() => {
  const nav = document.querySelector("[data-public-nav]");
  const toggle = document.querySelector("[data-nav-toggle]");
  const root = document.documentElement;

  if (!nav) {
    return;
  }

  const syncHeaderOffset = () => {
    root.style.setProperty("--header-offset", `${nav.offsetHeight}px`);
  };

  syncHeaderOffset();

  if (!toggle) {
    window.addEventListener("resize", syncHeaderOffset);
    return;
  }

  const setOpen = (open) => {
    nav.dataset.open = open ? "true" : "false";
    toggle.setAttribute("aria-expanded", open ? "true" : "false");
    document.body.classList.toggle("nav-open", open);
    syncHeaderOffset();
  };

  toggle.addEventListener("click", () => {
    setOpen(nav.dataset.open !== "true");
  });

  nav.querySelectorAll("a").forEach((link) => {
    link.addEventListener("click", () => {
      setOpen(false);
    });
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      setOpen(false);
    }
  });

  window.addEventListener("resize", () => {
    syncHeaderOffset();
    if (window.innerWidth > 980) {
      setOpen(false);
    }
  });
})();
