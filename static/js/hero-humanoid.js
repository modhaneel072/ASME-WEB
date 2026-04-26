(() => {
  const shell = document.getElementById("heroHumanoidShell");
  const viewer = document.getElementById("heroHumanoid");
  const fallback = document.getElementById("heroHumanoidFallback");

  if (!shell || !viewer) {
    return;
  }

  const baseOrbit = { theta: -28, phi: 78, radius: 3.55 };
  const orbitRange = { theta: 4.2, phi: 3.2, radius: 0.14 };
  const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)");

  let pointerX = 0;
  let pointerY = 0;
  let frameId = 0;

  const orbitString = ({ theta, phi, radius }) =>
    `${theta.toFixed(2)}deg ${phi.toFixed(2)}deg ${radius.toFixed(2)}m`;

  const applyBaseOrbit = () => {
    viewer.cameraOrbit = orbitString(baseOrbit);
  };

  const hasWebGL = () => {
    const canvas = document.createElement("canvas");
    return Boolean(
      window.WebGLRenderingContext &&
      (canvas.getContext("webgl2") || canvas.getContext("webgl"))
    );
  };

  const setReady = () => {
    shell.classList.add("is-ready");
    shell.classList.remove("is-error");
    if (fallback) {
      fallback.hidden = true;
    }
  };

  const setError = () => {
    shell.classList.remove("is-ready");
    shell.classList.add("is-error");
    if (fallback) {
      fallback.hidden = false;
    }
  };

  const updateOrbit = () => {
    frameId = 0;

    if (reducedMotion.matches || shell.classList.contains("is-error")) {
      applyBaseOrbit();
      return;
    }

    viewer.cameraOrbit = orbitString({
      theta: baseOrbit.theta + pointerX * orbitRange.theta,
      phi: baseOrbit.phi - pointerY * orbitRange.phi,
      radius: baseOrbit.radius + pointerY * orbitRange.radius,
    });
  };

  const queueOrbitUpdate = () => {
    if (!frameId) {
      frameId = window.requestAnimationFrame(updateOrbit);
    }
  };

  const onPointerMove = (event) => {
    if (reducedMotion.matches || shell.classList.contains("is-error")) {
      return;
    }

    const bounds = shell.getBoundingClientRect();
    pointerX = ((event.clientX - bounds.left) / bounds.width - 0.5) * 2;
    pointerY = ((event.clientY - bounds.top) / bounds.height - 0.5) * 2;
    queueOrbitUpdate();
  };

  const onPointerLeave = () => {
    pointerX = 0;
    pointerY = 0;
    queueOrbitUpdate();
  };

  const onReducedMotionChanged = () => {
    pointerX = 0;
    pointerY = 0;
    applyBaseOrbit();
  };

  const init = async () => {
    if (!hasWebGL()) {
      setError();
      return;
    }

    try {
      await Promise.race([
        customElements.whenDefined("model-viewer"),
        new Promise((_, reject) => {
          window.setTimeout(() => reject(new Error("model-viewer timed out")), 4000);
        }),
      ]);
    } catch (error) {
      setError();
      return;
    }

    applyBaseOrbit();

    viewer.addEventListener("load", () => {
      setReady();
      applyBaseOrbit();
      if (typeof viewer.play === "function") {
        viewer.play();
      }
    });

    viewer.addEventListener("error", () => {
      setError();
    });

    shell.addEventListener("pointermove", onPointerMove);
    shell.addEventListener("pointerleave", onPointerLeave);

    if (typeof reducedMotion.addEventListener === "function") {
      reducedMotion.addEventListener("change", onReducedMotionChanged);
    } else if (typeof reducedMotion.addListener === "function") {
      reducedMotion.addListener(onReducedMotionChanged);
    }

    if (viewer.loaded) {
      setReady();
    }
  };

  init();
})();
