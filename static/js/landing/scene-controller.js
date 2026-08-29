import * as THREE from "../../vendor/three-0.185.1/three.module.min.js";
import {
  SCENE_KEYS,
  clamp01,
  getSceneAtProgress,
  getSceneLocalProgress,
  getSceneWeights,
  sampleCameraPath,
} from "./camera-path.js";
import {
  LANDING_COLORS,
  createMaterialLibrary,
  disposeMaterialLibrary,
} from "./materials.js";
import {
  buildLineFragmentShader,
  buildLineVertexShader,
} from "../../shaders/landing/build-line-shader.js";

const TWO_PI = Math.PI * 2;

/**
 * Real-time scene renderer for “Built at Iowa — The Build Line”.
 *
 * All temporary vehicle/mechanism geometry below is deliberately procedural
 * and generic. It communicates engineering assembly without claiming to be a
 * digital twin of any current competition machine.
 */
export class SceneController {
  constructor({
    canvas,
    root,
    qualityManager,
    documentRef = globalThis.document,
    windowRef = globalThis.window,
    onError = () => {},
    onFrame = () => {},
  } = {}) {
    this.canvas = canvas || null;
    this.root = root || canvas?.parentElement || null;
    this.qualityManager = qualityManager || null;
    this.document = documentRef || null;
    this.window = windowRef || null;
    this.onError = onError;
    this.onFrame = onFrame;
    this.profile = qualityManager?.profile || { dprMax: 1, radialSegments: 6, curveSegments: 100, shadows: false, underwaterParticles: 0, photoLayers: 2 };
    this.quality = qualityManager?.quality || "low";

    this.progress = 0;
    this.targetProgress = 0;
    this.sceneKey = "intro";
    this.reducedMotion = false;
    this.paused = false;
    this.inViewport = true;
    this.running = false;
    this.disposed = false;
    this.raf = 0;
    this.lastFrameAt = 0;
    this.cameraSample = {};
    this.cameraTarget = new THREE.Vector3();
    this.groups = {};
    this.hotspotAnchors = {};
    this.photoMeshes = [];
    this.photoAssetsApplied = false;
    this.initialized = false;
    this.shaderCompileError = null;
    this.shaderFailureReported = false;
    this.ownedMaterials = [];
    this.ownedTextures = new Set();
  }

  init() {
    if (this.disposed) throw new Error("Landing scene controller has been disposed");
    if (!this.canvas || !this.qualityManager?.context) throw new Error("WebGL2 is unavailable for the landing canvas");

    this.renderer = new THREE.WebGLRenderer({
      canvas: this.canvas,
      context: this.qualityManager.context,
      alpha: true,
      antialias: false,
      powerPreference: "high-performance",
      premultipliedAlpha: true,
      preserveDrawingBuffer: false,
    });
    this.renderer.outputColorSpace = THREE.SRGBColorSpace;
    this.renderer.toneMapping = THREE.ACESFilmicToneMapping;
    this.renderer.toneMappingExposure = 0.84;
    this.renderer.shadowMap.enabled = Boolean(this.profile.shadows);
    this.renderer.shadowMap.type = THREE.PCFSoftShadowMap;
    if (this.renderer.debug) {
      this.renderer.debug.checkShaderErrors = true;
      this.renderer.debug.onShaderError = (...details) => {
        const error = new Error("The WebGL scene could not compile on this device");
        this.shaderCompileError = error;
        console.error("[landing] Shader compilation failed.", ...details);
        if (this.initialized && !this.shaderFailureReported) {
          this.shaderFailureReported = true;
          const schedule = this.window?.queueMicrotask || globalThis.queueMicrotask || ((callback) => setTimeout(callback, 0));
          schedule(() => {
            if (!this.disposed) this.onError(error);
          });
        }
      };
    }

    this.scene = new THREE.Scene();
    this.scene.background = new THREE.Color(0x050607);
    this.scene.fog = new THREE.FogExp2(0x050607, 0.018);
    this.camera = new THREE.PerspectiveCamera(46, 1, 0.08, 180);
    this.materials = createMaterialLibrary({ quality: this.quality });

    this._createLighting();
    this._createBuildLine();
    this._createIntro();
    this._createDesignAssembly();
    this._createIam3dAssembly();
    this._createPlaneAssembly();
    this._createRovAssembly();
    this._createStudentDesignMechanism();
    this._createRoboticDog();
    this._createPeopleWorkshop();
    this._createJoinFrame();
    this._setupObservers();
    this.resize();
    this._applyProgress(0);

    try {
      this.renderer.compile(this.scene, this.camera);
      if (this.shaderCompileError) throw this.shaderCompileError;
      // Paint a valid frame before main.js marks data-webgl="available".
      this.renderer.render(this.scene, this.camera);
      if (this.shaderCompileError) throw this.shaderCompileError;
    } catch (error) {
      this.dispose();
      throw new Error(`Landing scene compilation failed: ${error.message || error}`);
    }
    this.initialized = true;
    return this;
  }

  start() {
    if (this.running || this.disposed || !this.renderer) return;
    this.running = true;
    this.lastFrameAt = timestamp(this.window);
    this.raf = this.window.requestAnimationFrame(this._tick);
  }

  setProgress(progress) {
    this.targetProgress = clamp01(progress);
    if (this.reducedMotion) this.progress = this.targetProgress;
  }

  setReducedMotion(reduced) {
    this.reducedMotion = Boolean(reduced);
    if (this.reducedMotion) this.progress = this.targetProgress;
    if (this.rovParticles) {
      this.rovParticles.visible = !this.reducedMotion && this.profile.underwaterParticles > 0;
    }
  }

  setPaused(paused) {
    this.paused = Boolean(paused);
    if (!this.paused) this.lastFrameAt = timestamp(this.window);
  }

  setQuality(quality) {
    this.quality = quality;
    this.profile = this.qualityManager?.profile || this.profile;
    if (!this.renderer || quality === "fallback") return;
    this.renderer.setPixelRatio(this.qualityManager?.pixelRatio || 1);
    this.renderer.shadowMap.enabled = Boolean(this.profile.shadows);
    if (this.rovParticles) {
      this.rovParticles.visible = !this.reducedMotion && this.profile.underwaterParticles > 0;
      this.rovParticles.geometry.setDrawRange(0, Math.min(
        this.rovParticles.geometry.getAttribute("position")?.count || 0,
        this.profile.underwaterParticles,
      ));
    }
    if (this.waterMaterial) {
      this.waterMaterial.transmission = this.profile.waterTransmission || 0;
      this.waterMaterial.needsUpdate = true;
    }
    this._rebuildBuildLineGeometry();
    this.photoMeshes.forEach((mesh, index) => {
      mesh.visible = index < (this.profile.photoLayers || 2);
    });
    this.resize();
  }

  setChapterAssets(chapter, assets = {}) {
    if (chapter !== "people" || this.photoAssetsApplied) return;
    const textures = [
      assets["people-electronics"],
      assets["people-presentation"],
      assets["people-workshop"],
      assets["people-prototype"],
    ].filter(Boolean);
    if (!textures.length) return;
    this._installPhotoLayers(textures.slice(0, this.profile.photoLayers || 2));
    this.photoAssetsApplied = true;
  }

  getHotspotAnchor(scene = this.sceneKey) {
    return this.hotspotAnchors[scene] || null;
  }

  getMetrics() {
    return {
      progress: this.progress,
      scene: this.sceneKey,
      camera: this.camera?.position || null,
      target: this.cameraTarget,
      drawCalls: this.renderer?.info?.render?.calls || 0,
      triangles: this.renderer?.info?.render?.triangles || 0,
      geometries: this.renderer?.info?.memory?.geometries || 0,
      textures: this.renderer?.info?.memory?.textures || 0,
    };
  }

  resize() {
    if (!this.renderer || !this.camera) return;
    const width = Math.max(1, this.canvas.clientWidth || this.root?.clientWidth || this.window?.innerWidth || 1);
    const height = Math.max(1, this.canvas.clientHeight || this.root?.clientHeight || this.window?.innerHeight || 1);
    this.renderer.setPixelRatio(this.qualityManager?.pixelRatio || 1);
    this.renderer.setSize(width, height, false);
    this.camera.aspect = width / height;
    this.camera.updateProjectionMatrix();
    this.viewport = { width, height };
  }

  dispose() {
    if (this.disposed) return;
    this.disposed = true;
    this.running = false;
    this.window?.cancelAnimationFrame?.(this.raf);
    this.resizeObserver?.disconnect?.();
    this.intersectionObserver?.disconnect?.();
    this.canvas?.removeEventListener?.("webglcontextlost", this._onContextLost);
    this.window?.removeEventListener?.("resize", this.resize);

    this.scene?.traverse?.((node) => {
      node.geometry?.dispose?.();
      const materials = Array.isArray(node.material) ? node.material : [node.material];
      for (const material of materials) {
        if (!material || material.userData?.landingSharedMaterial) continue;
        material.dispose?.();
      }
    });
    for (const material of this.ownedMaterials) material?.dispose?.();
    disposeMaterialLibrary(this.materials);
    this.renderer?.dispose?.();
    this.groups = {};
    this.hotspotAnchors = {};
    this.photoMeshes.length = 0;
  }

  _tick = (time) => {
    if (!this.running || this.disposed) return;
    this.raf = this.window.requestAnimationFrame(this._tick);
    const deltaMs = Math.min(100, Math.max(1, time - this.lastFrameAt));
    this.lastFrameAt = time;
    if (this.paused || !this.inViewport || this.document?.hidden) return;

    try {
      if (this.reducedMotion) {
        this.progress = this.targetProgress;
      } else {
        const follow = 1 - Math.exp(-(deltaMs / 1000) * 8.5);
        this.progress += (this.targetProgress - this.progress) * follow;
        if (Math.abs(this.targetProgress - this.progress) < 0.00002) this.progress = this.targetProgress;
      }
      this._applyProgress(this.progress);
      this.renderer.render(this.scene, this.camera);
      this.onFrame({ deltaMs, ...this.getMetrics(), viewport: this.viewport });
    } catch (error) {
      this.running = false;
      console.error("[landing] WebGL render loop stopped.", error);
      this.onError(error);
    }
  };

  _applyProgress(progress) {
    const value = clamp01(progress);
    const sceneKey = getSceneAtProgress(value);
    const weights = getSceneWeights(value);
    sampleCameraPath(value, this.cameraSample);
    this.sceneKey = sceneKey;

    this.camera.position.fromArray(this.cameraSample.position);
    this.cameraTarget.fromArray(this.cameraSample.target);
    this.camera.up.set(0, 1, 0);
    this.camera.lookAt(this.cameraTarget);
    this.camera.rotateZ(this.reducedMotion ? 0 : this.cameraSample.roll);
    if (Math.abs(this.camera.fov - this.cameraSample.fov) > 0.001) {
      this.camera.fov = this.cameraSample.fov;
      this.camera.updateProjectionMatrix();
    }

    for (const key of SCENE_KEYS) {
      if (this.groups[key]) this.groups[key].visible = weights[key] > 0.002;
    }
    this._animateIntro(getSceneLocalProgress("intro", value));
    this._animateDesign(getSceneLocalProgress("design", value));
    this._animateIam3d(getSceneLocalProgress("iam3d", value));
    this._animatePlane(getSceneLocalProgress("plane", value));
    this._animateRov(getSceneLocalProgress("rov", value));
    this._animateStudentDesign(getSceneLocalProgress("student-design", value));
    this._animateRoboticDog(getSceneLocalProgress("robotic-dog", value));
    this._animatePeople(getSceneLocalProgress("people", value));
    this._animateJoin(getSceneLocalProgress("join", value));
    this._updateAtmosphere(value, weights);

    if (this.buildLineMaterial) {
      const introLocal = getSceneLocalProgress("intro", value);
      const minimumReveal = smoothstep(0.06, 0.72, introLocal) * 0.046;
      this.buildLineMaterial.uniforms.uProgress.value = Math.max(minimumReveal, Math.min(1, value * 1.035));
      this.buildLineMaterial.uniforms.uOpacity.value = sceneKey === "join" ? 0.78 : 0.92;
    }
  }

  _createLighting() {
    this.hemisphereLight = new THREE.HemisphereLight(0xb8c0c7, 0x090a0b, 0.72);
    this.scene.add(this.hemisphereLight);
    this.keyLight = new THREE.DirectionalLight(0xfff3d2, 2.15);
    this.keyLight.position.set(-8, 12, 8);
    this.keyLight.castShadow = Boolean(this.profile.shadows);
    if (this.keyLight.castShadow) {
      this.keyLight.shadow.mapSize.set(this.profile.shadowMapSize, this.profile.shadowMapSize);
      this.keyLight.shadow.camera.near = 1;
      this.keyLight.shadow.camera.far = 70;
      this.keyLight.shadow.camera.left = -12;
      this.keyLight.shadow.camera.right = 12;
      this.keyLight.shadow.camera.top = 12;
      this.keyLight.shadow.camera.bottom = -12;
      this.keyLight.shadow.bias = -0.0004;
    }
    this.scene.add(this.keyLight);
    this.rimLight = new THREE.DirectionalLight(LANDING_COLORS.iowaGold, 0.42);
    this.rimLight.position.set(8, 2, -9);
    this.scene.add(this.rimLight);
    this.cameraFill = new THREE.PointLight(0xd9dde0, 150, 20, 2);
    this.cameraFill.position.set(0, 0.6, 0.8);
    this.camera.add(this.cameraFill);
    this.scene.add(this.camera);
  }

  _createBuildLine() {
    const points = [
      [0, 0, -6.3], [0, 0, -3.0], [0, 0, 0], [4.2, 0, 0],
      [9.8, 0, 0], [14.8, 0.3, 0], [20.2, -0.45, -0.8], [25.2, 0.15, 0],
      [31.0, 0.2, 0], [37.0, 1.0, 0], [44.0, 0.9, 0], [52.0, 1.2, 0],
      [59.0, 0.4, 0], [64.0, -1.15, 0], [71.0, -0.45, 0], [81.0, 0.0, 0],
      [89.0, 0.2, 0], [99.4, -0.4, 0], [108.5, 0.2, 0], [118.0, 0.35, 0],
      [126.0, 0.1, 0], [132.0, 0.0, 0], [136.0, -1.65, -3.2], [136.0, -1.65, 3.2],
    ].map((point) => new THREE.Vector3(...point));
    this.buildLineCurve = new THREE.CatmullRomCurve3(points, false, "centripetal", 0.28);
    const geometry = new THREE.TubeGeometry(
      this.buildLineCurve,
      this.profile.curveSegments,
      this.quality === "low" ? 0.028 : 0.035,
      this.profile.radialSegments,
      false,
    );
    this.buildLineMaterial = new THREE.ShaderMaterial({
      name: "BuildLineDatumShader",
      uniforms: {
        uColor: { value: new THREE.Color(LANDING_COLORS.iowaGold) },
        uProgress: { value: 0 },
        uOpacity: { value: 0.92 },
      },
      vertexShader: buildLineVertexShader,
      fragmentShader: buildLineFragmentShader,
      transparent: true,
      depthWrite: true,
      side: THREE.DoubleSide,
    });
    this.ownedMaterials.push(this.buildLineMaterial);
    this.buildLine = new THREE.Mesh(geometry, this.buildLineMaterial);
    this.buildLine.name = "Iowa gold Build Line";
    this.buildLine.renderOrder = 9;
    this.scene.add(this.buildLine);
  }

  _rebuildBuildLineGeometry() {
    if (!this.buildLine || !this.buildLineCurve || !this.profile.curveSegments || !this.profile.radialSegments) return;
    const previous = this.buildLine.geometry;
    this.buildLine.geometry = new THREE.TubeGeometry(
      this.buildLineCurve,
      this.profile.curveSegments,
      this.quality === "low" ? 0.028 : 0.035,
      this.profile.radialSegments,
      false,
    );
    previous?.dispose?.();
  }

  _createIntro() {
    const group = this._sceneGroup("intro", 0, 0, 0);
    const gridPoints = [];
    for (let y = -2.4; y <= 2.4; y += 0.6) {
      gridPoints.push(0.08, y, -6.7, 0.08, y, 1.1);
    }
    for (let z = -6.6; z <= 1.2; z += 0.65) {
      gridPoints.push(0.08, -2.4, z, 0.08, 2.4, z);
    }
    const gridGeometry = new THREE.BufferGeometry();
    gridGeometry.setAttribute("position", new THREE.Float32BufferAttribute(gridPoints, 3));
    this.introGrid = new THREE.LineSegments(gridGeometry, this.materials.draftingLine);
    group.add(this.introGrid);

    const tickPoints = [];
    for (let z = -6; z <= 0.1; z += 0.5) {
      const length = Math.abs(Math.round(z * 2)) % 4 === 0 ? 0.24 : 0.12;
      tickPoints.push(0.02, -length, z, 0.02, length, z);
    }
    const ticksGeometry = new THREE.BufferGeometry();
    ticksGeometry.setAttribute("position", new THREE.Float32BufferAttribute(tickPoints, 3));
    this.introTicks = new THREE.LineSegments(ticksGeometry, this.materials.buildLineWire);
    group.add(this.introTicks);
  }

  _createDesignAssembly() {
    const group = this._sceneGroup("design", 10, 0, 0);
    const radial = Math.max(12, this.profile.radialSegments * 2);
    this.designSolids = [];
    this.designWire = new THREE.Group();
    group.add(this.designWire);

    [-1.25, -0.42, 0.35, 1.05].forEach((x, index) => {
      const radius = [1.55, 1.08, 0.82, 1.28][index];
      const points = [];
      for (let step = 0; step <= 64; step += 1) {
        const angle = (step / 64) * TWO_PI;
        points.push(new THREE.Vector3(x, Math.cos(angle) * radius, Math.sin(angle) * radius));
      }
      const line = new THREE.Line(new THREE.BufferGeometry().setFromPoints(points), this.materials.draftingLine);
      this.designWire.add(line);
    });

    const axisGeometry = new THREE.BufferGeometry().setFromPoints([
      new THREE.Vector3(-3.2, 0, 0), new THREE.Vector3(3.2, 0, 0),
    ]);
    this.designAxis = new THREE.Line(axisGeometry, this.materials.buildLineWire);
    group.add(this.designAxis);

    this._designPart(group, new THREE.CylinderGeometry(1.56, 1.56, 0.24, radial, 1, false), this.materials.machinedAluminum, -1.1, -2.5);
    this._designPart(group, new THREE.CylinderGeometry(1.08, 1.08, 0.52, radial, 1, false), this.materials.darkSteel, -0.42, -1.35);
    this._designPart(group, new THREE.CylinderGeometry(0.84, 0.84, 0.86, radial, 1, false), this.materials.machinedAluminum, 0.36, 1.38);
    this._designPart(group, new THREE.CylinderGeometry(1.30, 1.30, 0.22, radial, 1, false), this.materials.darkSteel, 0.98, 2.45);
    this._designPart(group, new THREE.CylinderGeometry(0.44, 0.44, 2.25, radial, 1, false), this.materials.darkSteel, 0.0, 0.0);

    this.designFasteners = new THREE.Group();
    this.designFasteners.userData.assembledX = -1.32;
    this.designFasteners.userData.explodedX = -3.0;
    const boltGeometry = new THREE.CylinderGeometry(0.095, 0.095, 0.42, 10);
    for (let index = 0; index < 8; index += 1) {
      const angle = (index / 8) * TWO_PI;
      const bolt = new THREE.Mesh(boltGeometry, this.materials.darkSteel);
      bolt.rotation.z = Math.PI / 2;
      bolt.position.set(0, Math.cos(angle) * 1.16, Math.sin(angle) * 1.16);
      this.designFasteners.add(bolt);
    }
    this.designFasteners.position.x = this.designFasteners.userData.explodedX;
    group.add(this.designFasteners);
    this.hotspotAnchors.design = this.designAxis;
  }

  _designPart(group, geometry, material, assembledX, explodedX) {
    const mesh = new THREE.Mesh(geometry, material);
    mesh.rotation.z = Math.PI / 2;
    mesh.position.x = explodedX;
    mesh.userData.assembledX = assembledX;
    mesh.userData.explodedX = explodedX;
    mesh.castShadow = Boolean(this.profile.shadows);
    mesh.receiveShadow = Boolean(this.profile.shadows);
    group.add(mesh);
    this.designSolids.push(mesh);
    return mesh;
  }

  _createIam3dAssembly() {
    // Temporary implementation geometry: an engineering rover archetype, not
    // a representation of Iowa's current IAM3D competition rover.
    const group = this._sceneGroup("iam3d", 24.5, 0, 0);
    this.roverChassis = new THREE.Group();
    group.add(this.roverChassis);
    this.roverWheels = [];

    for (const z of [-1.05, 1.05]) {
      const rail = box(5.4, 0.26, 0.28, this.materials.machinedAluminum, [0, 0.55, z]);
      this.roverChassis.add(rail);
    }
    for (const x of [-2.25, 0, 2.25]) {
      this.roverChassis.add(box(0.28, 0.24, 2.25, this.materials.darkSteel, [x, 0.52, 0]));
      for (const z of [-1.48, 1.48]) {
        const wheel = new THREE.Mesh(
          new THREE.CylinderGeometry(0.92, 0.92, 0.44, Math.max(12, this.profile.radialSegments * 2), 1, false),
          this.materials.rubber,
        );
        wheel.rotation.x = Math.PI / 2;
        wheel.position.set(x, -0.22, z);
        wheel.castShadow = Boolean(this.profile.shadows);
        group.add(wheel);
        this.roverWheels.push(wheel);

        const hub = new THREE.Mesh(new THREE.CylinderGeometry(0.30, 0.30, 0.49, 12), this.materials.machinedAluminum);
        hub.rotation.x = Math.PI / 2;
        hub.position.copy(wheel.position);
        group.add(hub);
        if (x === -2.25 && z === 1.48) this.hotspotAnchors.iam3d = hub;
      }
    }

    this.printedBracket = box(1.3, 0.72, 1.65, this.materials.printedPolymer, [-0.35, 1.0, 0]);
    this.roverChassis.add(this.printedBracket);
    const deck = box(3.5, 0.12, 1.6, this.materials.mattePolymer, [0.55, 1.36, 0]);
    this.roverChassis.add(deck);

    this.iamLayerLines = new THREE.Group();
    for (let index = 0; index < 11; index += 1) {
      const y = 0.67 + index * 0.065;
      const geometry = new THREE.BufferGeometry().setFromPoints([
        new THREE.Vector3(-1.0, y, -0.84), new THREE.Vector3(0.30, y, -0.84),
        new THREE.Vector3(0.30, y, 0.84), new THREE.Vector3(-1.0, y, 0.84),
      ]);
      this.iamLayerLines.add(new THREE.Line(geometry, this.materials.draftingLine));
    }
    this.roverChassis.add(this.iamLayerLines);
  }

  _createPlaneAssembly() {
    const group = this._sceneGroup("plane", 44, 1.0, 0);
    this.planeRoot = new THREE.Group();
    group.add(this.planeRoot);
    this.planeParts = [];
    this.planeRibs = [];

    const skyMaterial = new THREE.ShaderMaterial({
      name: "AnalyticalSkyGradient",
      uniforms: {},
      vertexShader: "varying vec2 vUv; void main(){vUv=uv;gl_Position=projectionMatrix*modelViewMatrix*vec4(position,1.0);}",
      fragmentShader: "varying vec2 vUv; void main(){vec3 low=vec3(0.20,0.25,0.28);vec3 high=vec3(0.50,0.60,0.65);gl_FragColor=vec4(mix(low,high,smoothstep(0.0,1.0,vUv.y)),1.0);}",
      side: THREE.DoubleSide,
      depthWrite: false,
    });
    this.ownedMaterials.push(skyMaterial);
    this.planeSky = new THREE.Mesh(new THREE.PlaneGeometry(26, 16), skyMaterial);
    this.planeSky.rotation.y = Math.PI / 2;
    this.planeSky.position.x = 7.0;
    this.planeSky.renderOrder = -10;
    group.add(this.planeSky);

    const fuselage = new THREE.Mesh(new THREE.CylinderGeometry(0.32, 0.48, 5.7, 12), this.materials.mattePolymer);
    fuselage.rotation.z = Math.PI / 2;
    fuselage.position.x = 0.2;
    this.planeRoot.add(fuselage);
    this.planeParts.push({ mesh: fuselage, threshold: 0.38 });

    const mainWing = box(3.4, 0.13, 8.0, this.materials.printedPolymer, [0.1, 0.05, 0]);
    this.planeRoot.add(mainWing);
    this.planeParts.push({ mesh: mainWing, threshold: 0.28 });
    const tailWing = box(1.25, 0.1, 3.1, this.materials.printedPolymer, [2.35, 0.05, 0]);
    this.planeRoot.add(tailWing);
    this.planeParts.push({ mesh: tailWing, threshold: 0.50 });
    const verticalTail = box(1.0, 1.0, 0.10, this.materials.printedPolymer, [2.35, 0.55, 0]);
    this.planeRoot.add(verticalTail);
    this.planeParts.push({ mesh: verticalTail, threshold: 0.54 });

    const spar = box(0.18, 0.17, 7.5, this.materials.machinedAluminum, [0.0, 0.06, 0]);
    this.planeRoot.add(spar);
    this.planeParts.push({ mesh: spar, threshold: 0.16 });
    for (let index = -6; index <= 6; index += 1) {
      const rib = box(2.7, 0.18, 0.055, this.materials.darkSteel, [0.1, 0.08, index * 0.56]);
      this.planeRoot.add(rib);
      this.planeRibs.push({ mesh: rib, threshold: 0.10 + ((index + 6) / 12) * 0.18 });
    }

    this.planePropeller = new THREE.Group();
    this.planePropeller.position.x = -2.9;
    const propHub = new THREE.Mesh(new THREE.CylinderGeometry(0.32, 0.32, 0.42, 12), this.materials.machinedAluminum);
    propHub.rotation.z = Math.PI / 2;
    this.planePropeller.add(propHub);
    const bladeA = box(0.10, 2.75, 0.20, this.materials.darkSteel);
    const bladeB = box(0.10, 0.20, 2.75, this.materials.darkSteel);
    this.planePropeller.add(bladeA, bladeB);
    this.planeRoot.add(this.planePropeller);
    this.planeParts.push({ mesh: this.planePropeller, threshold: 0.58 });
    this.hotspotAnchors.plane = propHub;

    const electronics = box(1.25, 0.48, 0.78, this.materials.glassHousing, [0.65, 0.36, 0]);
    this.planeRoot.add(electronics);
    this.planeParts.push({ mesh: electronics, threshold: 0.46 });
  }

  _createRovAssembly() {
    const group = this._sceneGroup("rov", 64, -1.0, 0);
    this.rovRoot = new THREE.Group();
    group.add(this.rovRoot);

    const waterMaterial = new THREE.MeshPhysicalMaterial({
      color: 0x31545c,
      roughness: 0.34,
      metalness: 0,
      transparent: true,
      opacity: 0.34,
      transmission: this.profile.waterTransmission || 0,
      depthWrite: false,
      side: THREE.DoubleSide,
    });
    this.ownedMaterials.push(waterMaterial);
    this.waterMaterial = waterMaterial;
    this.waterSurface = new THREE.Mesh(new THREE.PlaneGeometry(17, 14, 1, 1), waterMaterial);
    this.waterSurface.rotation.x = -Math.PI / 2;
    this.waterSurface.position.y = 2.1;
    this.rovRoot.add(this.waterSurface);

    const poolFloor = box(14, 0.08, 11, this.materials.darkSteel, [0, -2.25, 0]);
    poolFloor.receiveShadow = false;
    this.rovRoot.add(poolFloor);

    const frame = new THREE.Group();
    for (const y of [-0.85, 0.95]) {
      for (const z of [-1.35, 1.35]) frame.add(box(3.6, 0.14, 0.14, this.materials.machinedAluminum, [0, y, z]));
      for (const x of [-1.8, 1.8]) frame.add(box(0.14, 0.14, 2.7, this.materials.machinedAluminum, [x, y, 0]));
    }
    for (const x of [-1.8, 1.8]) {
      for (const z of [-1.35, 1.35]) frame.add(box(0.14, 1.8, 0.14, this.materials.machinedAluminum, [x, 0.05, z]));
    }
    this.rovRoot.add(frame);

    this.rovThrusters = [];
    for (const [x, y, z, rotation] of [
      [-1.45, 0.9, -1.35, [Math.PI / 2, 0, 0]],
      [1.45, 0.9, 1.35, [Math.PI / 2, 0, 0]],
      [-1.3, -0.55, 0, [0, 0, Math.PI / 2]],
      [1.3, -0.55, 0, [0, 0, Math.PI / 2]],
    ]) {
      const thruster = new THREE.Mesh(new THREE.CylinderGeometry(0.42, 0.42, 0.78, 12, 1, true), this.materials.darkSteel);
      thruster.position.set(x, y, z);
      thruster.rotation.set(...rotation);
      this.rovRoot.add(thruster);
      this.rovThrusters.push(thruster);
    }

    const cameraHousing = new THREE.Mesh(new THREE.CylinderGeometry(0.42, 0.42, 0.75, 16), this.materials.glassHousing);
    cameraHousing.rotation.z = Math.PI / 2;
    cameraHousing.position.set(-2.05, 0.2, 0);
    this.rovRoot.add(cameraHousing);
    this.hotspotAnchors.rov = cameraHousing;

    const tetherPoints = [
      new THREE.Vector3(1.8, 0.7, 1.2), new THREE.Vector3(3.2, 1.5, 1.8),
      new THREE.Vector3(5.0, 2.0, 2.4), new THREE.Vector3(6.2, 3.8, 3.0),
    ];
    const tetherCurve = new THREE.CatmullRomCurve3(tetherPoints);
    const tether = new THREE.Mesh(
      new THREE.TubeGeometry(tetherCurve, 36, 0.035, 6, false),
      this.materials.rubber,
    );
    this.rovRoot.add(tether);

    this.rovArm = new THREE.Group();
    this.rovArm.position.set(-0.65, -0.72, -1.25);
    this.rovArmPivot = new THREE.Group();
    this.rovArm.add(this.rovArmPivot);
    const upper = box(0.24, 1.45, 0.24, this.materials.buildLineGold, [0, -0.72, 0]);
    this.rovArmPivot.add(upper);
    this.rovElbow = new THREE.Group();
    this.rovElbow.position.y = -1.45;
    this.rovArmPivot.add(this.rovElbow);
    const forearm = box(0.22, 1.2, 0.22, this.materials.machinedAluminum, [0, -0.58, 0]);
    this.rovElbow.add(forearm);
    this.rovClaw = new THREE.Group();
    this.rovClaw.position.y = -1.2;
    this.rovElbow.add(this.rovClaw);
    this.rovClaw.add(box(0.18, 0.58, 0.12, this.materials.darkSteel, [0, -0.28, -0.18], [0.25, 0, 0]));
    this.rovClaw.add(box(0.18, 0.58, 0.12, this.materials.darkSteel, [0, -0.28, 0.18], [-0.25, 0, 0]));
    this.rovRoot.add(this.rovArm);

    this.rovTask = new THREE.Mesh(new THREE.TorusGeometry(0.72, 0.13, 8, 28), this.materials.printedPolymer);
    this.rovTask.rotation.y = Math.PI / 2;
    this.rovTask.position.set(-0.75, -1.65, -3.1);
    this.rovRoot.add(this.rovTask);

    const particleCount = this.profile.underwaterParticles || 0;
    if (particleCount) {
      const positions = new Float32Array(particleCount * 3);
      for (let index = 0; index < particleCount; index += 1) {
        const seed = pseudoRandom(index + 1);
        positions[index * 3] = (seed - 0.5) * 10;
        positions[index * 3 + 1] = pseudoRandom(index + 31) * 4 - 2;
        positions[index * 3 + 2] = (pseudoRandom(index + 67) - 0.5) * 8;
      }
      const geometry = new THREE.BufferGeometry();
      geometry.setAttribute("position", new THREE.BufferAttribute(positions, 3));
      const material = new THREE.PointsMaterial({ color: 0x94a9aa, size: 0.025, transparent: true, opacity: 0.22, depthWrite: false });
      this.ownedMaterials.push(material);
      this.rovParticles = new THREE.Points(geometry, material);
      this.rovRoot.add(this.rovParticles);
    }
  }

  _createStudentDesignMechanism() {
    const group = this._sceneGroup("student-design", 81.5, 0, 0);
    this.mechanismRoot = new THREE.Group();
    group.add(this.mechanismRoot);
    this.mechanismBase = box(0.42, 4.7, 7.2, this.materials.darkSteel, [0.55, 0, 0]);
    this.mechanismRoot.add(this.mechanismBase);

    this.crankPivot = new THREE.Vector3(-0.05, 0, -2.1);
    this.crank = box(0.24, 1.0, 0.18, this.materials.buildLineGold);
    this.mechanismRoot.add(this.crank);
    this.coupler = box(0.23, 1.0, 0.16, this.materials.machinedAluminum);
    this.mechanismRoot.add(this.coupler);
    this.sliderGuide = box(0.18, 0.38, 3.2, this.materials.machinedAluminum, [-0.05, 0.0, 1.65]);
    this.mechanismRoot.add(this.sliderGuide);
    this.slider = box(0.55, 0.82, 0.72, this.materials.printedPolymer, [-0.26, 0, 1.15]);
    this.mechanismRoot.add(this.slider);

    for (const [y, z] of [[0, -2.1], [0, 2.6]]) {
      const bearing = new THREE.Mesh(new THREE.CylinderGeometry(0.38, 0.38, 0.54, 16), this.materials.machinedAluminum);
      bearing.rotation.z = Math.PI / 2;
      bearing.position.set(-0.32, y, z);
      this.mechanismRoot.add(bearing);
      if (z < 0) this.hotspotAnchors["student-design"] = bearing;
    }

    this.stressLink = box(0.27, 1.0, 0.20, this.materials.stressOverlay);
    this.mechanismRoot.add(this.stressLink);
  }

  _createRoboticDog() {
    // Temporary implementation geometry: actuator-like joints and exposed
    // structure, not a claim about the team's eventual quadruped design.
    const group = this._sceneGroup("robotic-dog", 100, 0.4, 0);
    this.dogRoot = new THREE.Group();
    group.add(this.dogRoot);
    this.dogBuildParts = [];
    this.dogLegs = [];

    const bodyFrame = new THREE.Group();
    bodyFrame.add(box(3.2, 0.20, 1.55, this.materials.machinedAluminum, [0, 0.55, 0]));
    bodyFrame.add(box(3.0, 0.24, 0.22, this.materials.darkSteel, [0, 0.1, -0.62]));
    bodyFrame.add(box(3.0, 0.24, 0.22, this.materials.darkSteel, [0, 0.1, 0.62]));
    this.dogRoot.add(bodyFrame);
    this._tagBuild(bodyFrame, 5);

    const bodyStructure = box(2.55, 0.65, 1.30, this.materials.mattePolymer, [0, 0.78, 0]);
    this.dogRoot.add(bodyStructure);
    this._tagBuild(bodyStructure, 6);
    const electronicsTray = box(2.1, 0.15, 1.0, this.materials.printedPolymer, [0, 1.22, 0]);
    this.dogRoot.add(electronicsTray);
    this._tagBuild(electronicsTray, 7);

    for (const x of [-1.18, 1.18]) {
      for (const z of [-0.73, 0.73]) {
        const leg = this._createDogLeg(x, z);
        this.dogRoot.add(leg.root);
        this.dogLegs.push(leg);
      }
    }

    const wiringMaterial = this.materials.buildLineGold;
    const wirePoints = [
      new THREE.Vector3(-1.3, 1.1, -0.5), new THREE.Vector3(0, 1.32, -0.56), new THREE.Vector3(1.3, 1.1, -0.5),
    ];
    const wiring = new THREE.Mesh(new THREE.TubeGeometry(new THREE.CatmullRomCurve3(wirePoints), 28, 0.025, 5), wiringMaterial);
    this.dogRoot.add(wiring);
    this._tagBuild(wiring, 8);

    const sensorPost = box(0.38, 0.62, 0.50, this.materials.machinedAluminum, [-1.52, 1.15, 0]);
    const sensorHousing = box(0.45, 0.30, 0.72, this.materials.darkSteel, [-1.72, 1.42, 0]);
    this.dogRoot.add(sensorPost, sensorHousing);
    this._tagBuild(sensorPost, 9);
    this._tagBuild(sensorHousing, 9);

    const ledMaterial = new THREE.MeshBasicMaterial({ color: 0x6f8f65, toneMapped: false });
    this.ownedMaterials.push(ledMaterial);
    this.dogStatusLed = new THREE.Mesh(new THREE.BoxGeometry(0.04, 0.08, 0.22), ledMaterial);
    this.dogStatusLed.position.set(-1.96, 1.42, 0);
    this.dogRoot.add(this.dogStatusLed);
    this._tagBuild(this.dogStatusLed, 10);
  }

  _createDogLeg(x, z) {
    const root = new THREE.Group();
    root.position.set(x, 0.18, z);
    const hip = new THREE.Group();
    root.add(hip);
    const hipActuator = new THREE.Mesh(new THREE.CylinderGeometry(0.30, 0.30, 0.48, 12), this.materials.darkSteel);
    hipActuator.rotation.x = Math.PI / 2;
    hip.add(hipActuator);
    this._tagBuild(hipActuator, 4);
    const hipFastener = new THREE.Mesh(new THREE.CylinderGeometry(0.085, 0.085, 0.52, 8), this.materials.machinedAluminum);
    hipFastener.rotation.x = Math.PI / 2;
    hip.add(hipFastener);
    this._tagBuild(hipFastener, 4);

    const upper = box(0.30, 1.18, 0.30, this.materials.machinedAluminum, [0, -0.62, 0]);
    hip.add(upper);
    this._tagBuild(upper, 3);
    const knee = new THREE.Group();
    knee.position.y = -1.22;
    hip.add(knee);
    const kneeJoint = new THREE.Mesh(new THREE.CylinderGeometry(0.27, 0.27, 0.42, 12), this.materials.darkSteel);
    kneeJoint.rotation.x = Math.PI / 2;
    knee.add(kneeJoint);
    this._tagBuild(kneeJoint, 1);
    const kneeFastener = new THREE.Mesh(new THREE.CylinderGeometry(0.08, 0.08, 0.46, 8), this.materials.machinedAluminum);
    kneeFastener.rotation.x = Math.PI / 2;
    knee.add(kneeFastener);
    this._tagBuild(kneeFastener, 1);
    const lower = box(0.26, 1.12, 0.26, this.materials.printedPolymer, [0, -0.58, 0]);
    knee.add(lower);
    this._tagBuild(lower, 2);
    const foot = box(0.46, 0.18, 0.34, this.materials.rubber, [0.12, -1.12, 0]);
    knee.add(foot);
    this._tagBuild(foot, 10);
    if (x < 0 && z > 0) this.hotspotAnchors["robotic-dog"] = kneeJoint;
    return { root, hip, knee, side: Math.sign(z), fore: Math.sign(x) };
  }

  _tagBuild(object, step) {
    object.userData.buildStep = step;
    this.dogBuildParts.push(object);
  }

  _createPeopleWorkshop() {
    const group = this._sceneGroup("people", 118, 0.2, 0);
    this.peopleRoot = new THREE.Group();
    group.add(this.peopleRoot);

    const door = new THREE.Group();
    door.position.x = -5.0;
    door.add(box(0.3, 5.2, 0.35, this.materials.darkSteel, [0, 0.5, -3.0]));
    door.add(box(0.3, 5.2, 0.35, this.materials.darkSteel, [0, 0.5, 3.0]));
    door.add(box(0.3, 0.35, 6.3, this.materials.darkSteel, [0, 3.0, 0]));
    const doorwayMaterial = new THREE.MeshBasicMaterial({ color: 0xe4d6ae, transparent: true, opacity: 0.38, side: THREE.DoubleSide, depthWrite: false });
    this.ownedMaterials.push(doorwayMaterial);
    const doorway = new THREE.Mesh(new THREE.PlaneGeometry(6, 5), doorwayMaterial);
    doorway.rotation.y = Math.PI / 2;
    doorway.position.set(0.18, 0.5, 0);
    door.add(doorway);
    this.peopleRoot.add(door);

    this.photoPlaceholderFrames = new THREE.Group();
    const frameMaterial = this.materials.darkSteel;
    const frameSpecs = [
      [0.0, 0.8, -2.6, 4.4, 3.1],
      [1.2, -0.2, 1.6, 3.2, 4.4],
      [2.1, 1.0, 3.5, 3.7, 2.6],
    ];
    for (const [x, y, z, width, height] of frameSpecs) {
      const frame = new THREE.Group();
      frame.position.set(x, y, z);
      frame.add(box(0.10, height + 0.18, 0.12, frameMaterial, [0, 0, -width / 2]));
      frame.add(box(0.10, height + 0.18, 0.12, frameMaterial, [0, 0, width / 2]));
      frame.add(box(0.10, 0.12, width, frameMaterial, [0, height / 2, 0]));
      frame.add(box(0.10, 0.12, width, frameMaterial, [0, -height / 2, 0]));
      this.photoPlaceholderFrames.add(frame);
    }
    this.peopleRoot.add(this.photoPlaceholderFrames);
  }

  _installPhotoLayers(textures) {
    const specs = [
      { position: [0.08, 0.8, -2.6], size: [4.25, 2.85], fit: "landscape" },
      { position: [1.28, -0.2, 1.6], size: [3.0, 4.18], fit: "portrait" },
      { position: [2.18, 1.0, 3.5], size: [3.52, 2.38], fit: "landscape" },
      { position: [2.9, -1.05, -0.45], size: [3.45, 2.0], fit: "landscape" },
    ];
    textures.forEach((texture, index) => {
      const spec = specs[index];
      if (!spec) return;
      const material = new THREE.MeshBasicMaterial({
        map: texture,
        color: 0xd8d8d6,
        transparent: true,
        opacity: 0.94,
        side: THREE.DoubleSide,
        toneMapped: true,
      });
      fitTextureCover(texture, spec.size[0] / spec.size[1]);
      material.userData.landingPhotoMaterial = true;
      this.ownedMaterials.push(material);
      const mesh = new THREE.Mesh(new THREE.PlaneGeometry(spec.size[0], spec.size[1]), material);
      mesh.rotation.y = Math.PI / 2;
      mesh.position.fromArray(spec.position);
      mesh.userData.basePosition = mesh.position.clone();
      mesh.userData.layerIndex = index;
      mesh.renderOrder = 2 + index;
      this.peopleRoot.add(mesh);
      this.photoMeshes.push(mesh);
    });
  }

  _createJoinFrame() {
    const group = this._sceneGroup("join", 136, 0, 0);
    const points = [
      new THREE.Vector3(0, -1.65, -3.2),
      new THREE.Vector3(0, -1.65, 3.2),
      new THREE.Vector3(0, 1.65, 3.2),
      new THREE.Vector3(0, 1.65, -3.2),
      new THREE.Vector3(0, -1.65, -3.2),
    ];
    const geometry = new THREE.BufferGeometry().setFromPoints(points);
    geometry.setDrawRange(0, 0);
    this.joinFrame = new THREE.Line(geometry, this.materials.buildLineWire);
    this.joinFrame.renderOrder = 10;
    group.add(this.joinFrame);
    const quietBackdropMaterial = new THREE.MeshBasicMaterial({ color: 0x070809, transparent: true, opacity: 0.86, side: THREE.DoubleSide, depthWrite: false });
    this.ownedMaterials.push(quietBackdropMaterial);
    this.joinBackdrop = new THREE.Mesh(new THREE.PlaneGeometry(9.5, 6.6), quietBackdropMaterial);
    this.joinBackdrop.rotation.y = Math.PI / 2;
    this.joinBackdrop.position.x = 0.2;
    this.joinBackdrop.renderOrder = -1;
    group.add(this.joinBackdrop);
  }

  _sceneGroup(key, x, y, z) {
    const group = new THREE.Group();
    group.name = `landing-scene-${key}`;
    group.position.set(x, y, z);
    this.groups[key] = group;
    this.scene.add(group);
    return group;
  }

  _animateIntro(local) {
    if (!this.introGrid) return;
    this.introGrid.material.opacity = 0.12 + smoothstep(0.05, 0.45, local) * 0.32;
    this.introTicks.visible = local > 0.12;
  }

  _animateDesign(local) {
    if (!this.designSolids) return;
    const instant = this.reducedMotion && local > 0;
    const sketch = instant ? 1 : smoothstep(0.04, 0.28, local);
    const solid = instant ? 1 : smoothstep(0.24, 0.48, local);
    const assemble = instant ? 1 : mechanicalEase(clamp01((local - 0.42) / 0.45));
    this.designWire.visible = !instant && sketch > 0.01 && local < 0.82;
    this.designAxis.visible = local > 0.06;
    for (const part of this.designSolids) {
      part.visible = solid > 0.01;
      part.position.x = THREE.MathUtils.lerp(part.userData.explodedX, part.userData.assembledX, assemble);
      part.scale.y = Math.max(0.015, solid);
    }
    this.designFasteners.visible = this.quality !== "low" && local > 0.34;
    this.designFasteners.position.x = THREE.MathUtils.lerp(
      this.designFasteners.userData.explodedX,
      this.designFasteners.userData.assembledX,
      assemble,
    );
  }

  _animateIam3d(local) {
    if (!this.roverChassis) return;
    const travel = mechanicalEase(local);
    const rotation = this.reducedMotion ? 0 : travel * TWO_PI * 2.2;
    this.roverWheels.forEach((wheel, index) => {
      wheel.rotation.y = this.reducedMotion ? 0 : rotation * (index % 2 ? 1 : -1);
    });
    const suspension = this.reducedMotion ? 0 : Math.sin(travel * Math.PI * 2) * 0.055 * Math.sin(local * Math.PI);
    this.roverChassis.position.y = suspension;
    this.iamLayerLines.visible = this.quality !== "low" && !this.reducedMotion && local > 0.18 && local < 0.56;
  }

  _animatePlane(local) {
    if (!this.planeRoot) return;
    const instant = this.reducedMotion && local > 0;
    this.planeRibs.forEach(({ mesh, threshold }, index) => {
      const detailEnabled = this.quality !== "low" || index % 2 === 0;
      mesh.visible = detailEnabled && (instant || local >= threshold);
      mesh.scale.z = instant ? 1 : smoothstep(threshold, threshold + 0.10, local);
    });
    this.planeParts.forEach(({ mesh, threshold }) => {
      mesh.visible = instant || local >= threshold;
      const build = instant ? 1 : smoothstep(threshold, threshold + 0.14, local);
      mesh.scale.setScalar(Math.max(0.02, build));
    });
    this.planePropeller.rotation.x = this.reducedMotion ? 0 : local * TWO_PI * 4.5;
    this.planeRoot.rotation.x = this.reducedMotion ? 0 : Math.sin(local * Math.PI) * 0.11;
    this.planeRoot.position.y = this.reducedMotion ? 0 : Math.sin(local * Math.PI) * 0.18;
  }

  _animateRov(local) {
    if (!this.rovRoot) return;
    const motionLocal = this.reducedMotion ? (local > 0 ? 0.68 : 0) : local;
    const enter = this.reducedMotion && local > 0 ? 1 : smoothstep(0.0, 0.25, motionLocal);
    this.rovRoot.position.y = THREE.MathUtils.lerp(1.2, 0, enter);
    this.rovArmPivot.rotation.x = -0.35 + mechanicalEase(motionLocal) * 0.85;
    this.rovElbow.rotation.x = 0.55 - mechanicalEase(motionLocal) * 1.2;
    const grip = smoothstep(0.48, 0.68, motionLocal) - smoothstep(0.78, 0.94, motionLocal);
    this.rovClaw.children[0].rotation.x = 0.25 + grip * 0.28;
    this.rovClaw.children[1].rotation.x = -0.25 - grip * 0.28;
    this.rovTask.rotation.z = this.reducedMotion ? 0 : local * Math.PI * 0.42;
    if (this.rovParticles && !this.reducedMotion) this.rovParticles.position.y = (local * 0.7) % 0.35;
  }

  _animateStudentDesign(local) {
    if (!this.mechanismRoot) return;
    const cycle = this.reducedMotion ? (local > 0 ? 0.86 : 0) : mechanicalEase(local);
    const angle = cycle * TWO_PI;
    const revision = this.reducedMotion ? (local > 0 ? 1 : 0) : smoothstep(0.66, 0.82, local);
    const crankLength = THREE.MathUtils.lerp(1.22, 0.98, revision);
    const crankEnd = {
      y: Math.cos(angle) * crankLength,
      z: this.crankPivot.z + Math.sin(angle) * crankLength,
    };
    const sliderY = THREE.MathUtils.lerp(crankEnd.y * 0.74, crankEnd.y * 0.54, revision);
    const sliderZ = 1.8 + Math.sin(angle) * THREE.MathUtils.lerp(0.55, 0.34, revision);
    setLink(this.crank, { y: 0, z: this.crankPivot.z }, crankEnd);
    setLink(this.coupler, crankEnd, { y: sliderY, z: sliderZ });
    this.slider.position.y = sliderY;
    this.slider.position.z = sliderZ;

    const overload = this.reducedMotion ? 0 : smoothPulse(local, 0.43, 0.50, 0.59);
    setLink(this.stressLink, crankEnd, { y: sliderY, z: sliderZ });
    this.stressLink.material.opacity = overload * 0.72;
    this.stressLink.visible = overload > 0.005;
    const failureOffset = overload * 0.16;
    this.coupler.position.x = -failureOffset;
  }

  _animateRoboticDog(local) {
    if (!this.dogRoot) return;
    const active = local > 0;
    const buildStep = this.reducedMotion
      ? (active ? 10 : 0)
      : Math.floor(clamp01(local / 0.56) * 10.999);
    this.dogBuildParts.forEach((part) => {
      part.visible = buildStep >= (part.userData.buildStep || 10);
    });
    const powered = this.reducedMotion ? active : local >= 0.58;
    this.dogStatusLed.visible = powered && buildStep >= 10;
    const walk = this.reducedMotion ? 0 : clamp01((local - 0.70) / 0.30);
    const settle = this.reducedMotion ? (active ? 1 : 0) : smoothstep(0.56, 0.70, local);
    this.dogRoot.position.x = walk * 2.2;
    this.dogRoot.position.y = -settle * 0.12 + (this.reducedMotion ? 0 : Math.sin(walk * TWO_PI * 2) * 0.025 * walk);
    this.dogLegs.forEach((leg, index) => {
      const opposing = (index === 0 || index === 3) ? 0 : Math.PI;
      const phase = walk * TWO_PI * 1.4 + opposing;
      const stride = this.reducedMotion ? 0 : Math.sin(phase) * 0.32 * walk;
      const lift = this.reducedMotion ? 0 : Math.max(0, Math.cos(phase)) * 0.42 * walk;
      leg.hip.rotation.z = -0.08 + stride;
      leg.knee.rotation.z = 0.38 + lift - stride * 0.35;
    });
  }

  _animatePeople(local) {
    if (!this.peopleRoot) return;
    const reveal = smoothstep(0.08, 0.38, local);
    this.photoPlaceholderFrames.visible = !this.photoMeshes.length;
    this.photoMeshes.forEach((mesh, index) => {
      const base = mesh.userData.basePosition;
      const depth = (index + 1) * 0.11;
      mesh.position.x = base.x + (this.reducedMotion ? 0 : (local - 0.5) * depth);
      mesh.position.y = base.y + (this.reducedMotion ? 0 : Math.sin(local * Math.PI) * depth * 0.45);
      mesh.material.opacity = reveal * 0.94;
    });
  }

  _animateJoin(local) {
    if (!this.joinFrame) return;
    const reveal = smoothstep(0.12, 0.58, local);
    this.joinFrame.geometry.setDrawRange(0, Math.max(0, Math.ceil(reveal * 5)));
    this.joinFrame.visible = reveal > 0;
    this.joinBackdrop.material.opacity = THREE.MathUtils.lerp(0.2, 0.9, smoothstep(0.0, 0.42, local));
  }

  _updateAtmosphere(progress, weights) {
    const planeWeight = weights.plane || 0;
    const rovWeight = weights.rov || 0;
    const peopleWeight = weights.people || 0;
    const joinWeight = weights.join || 0;
    const base = new THREE.Color(0x050607);
    if (planeWeight > 0) base.lerp(new THREE.Color(0x303c42), planeWeight * 0.74);
    if (rovWeight > 0) base.lerp(new THREE.Color(0x071d22), rovWeight * 0.88);
    if (peopleWeight > 0) base.lerp(new THREE.Color(0x15130f), peopleWeight * 0.45);
    if (joinWeight > 0) base.lerp(new THREE.Color(0x060708), joinWeight * 0.92);
    this.scene.background.copy(base);
    this.scene.fog.color.copy(base);
    this.scene.fog.density = 0.015 + rovWeight * 0.032 + joinWeight * 0.006;
    this.hemisphereLight.intensity = 0.86 + planeWeight * 0.72 - rovWeight * 0.18 + peopleWeight * 0.34;
    this.keyLight.intensity = 2.25 + planeWeight * 1.05 - rovWeight * 0.62 - joinWeight * 0.78;
    this.rimLight.intensity = 0.52 + (1 - joinWeight) * 0.28;
    this.cameraFill.intensity = 150 + peopleWeight * 22 - rovWeight * 42 - joinWeight * 72;
    this.renderer.toneMappingExposure = 0.9 + planeWeight * 0.16 - rovWeight * 0.06;
  }

  _setupObservers() {
    this._onContextLost = (event) => {
      event.preventDefault();
      console.error("[landing] WebGL context was lost.");
      this.onError(new Error("The graphics context was lost"));
    };
    this.canvas.addEventListener("webglcontextlost", this._onContextLost, { passive: false });

    const ResizeObserverCtor = this.window?.ResizeObserver || globalThis.ResizeObserver;
    if (ResizeObserverCtor) {
      this.resizeObserver = new ResizeObserverCtor(() => this.resize());
      this.resizeObserver.observe(this.root || this.canvas);
    } else {
      this.resize = this.resize.bind(this);
      this.window?.addEventListener?.("resize", this.resize, { passive: true });
    }

    const IntersectionObserverCtor = this.window?.IntersectionObserver || globalThis.IntersectionObserver;
    if (IntersectionObserverCtor && this.root) {
      this.intersectionObserver = new IntersectionObserverCtor((entries) => {
        this.inViewport = Boolean(entries[0]?.isIntersecting);
        if (this.inViewport) this.lastFrameAt = timestamp(this.window);
      }, { rootMargin: "120px 0px", threshold: 0 });
      this.intersectionObserver.observe(this.root);
    }
  }
}

function box(width, height, depth, material, position = [0, 0, 0], rotation = [0, 0, 0]) {
  const mesh = new THREE.Mesh(new THREE.BoxGeometry(width, height, depth), material);
  mesh.position.set(...position);
  mesh.rotation.set(...rotation);
  return mesh;
}

function setLink(mesh, from, to) {
  const deltaY = to.y - from.y;
  const deltaZ = to.z - from.z;
  const length = Math.max(0.001, Math.hypot(deltaY, deltaZ));
  mesh.position.y = (from.y + to.y) / 2;
  mesh.position.z = (from.z + to.z) / 2;
  mesh.rotation.x = Math.atan2(deltaZ, deltaY);
  mesh.scale.y = length;
}

function smoothstep(edge0, edge1, value) {
  const span = Math.max(Number.EPSILON, edge1 - edge0);
  const t = clamp01((value - edge0) / span);
  return t * t * (3 - 2 * t);
}

function smoothPulse(value, start, peak, end) {
  return smoothstep(start, peak, value) * (1 - smoothstep(peak, end, value));
}

function mechanicalEase(value) {
  const t = clamp01(value);
  return t * t * (3 - 2 * t);
}

function pseudoRandom(seed) {
  const value = Math.sin(seed * 12.9898) * 43758.5453;
  return value - Math.floor(value);
}

function fitTextureCover(texture, targetAspect) {
  const image = texture?.image;
  const width = Number(image?.naturalWidth || image?.videoWidth || image?.width);
  const height = Number(image?.naturalHeight || image?.videoHeight || image?.height);
  if (!width || !height || !Number.isFinite(targetAspect) || targetAspect <= 0) return;

  const imageAspect = width / height;
  texture.offset.set(0, 0);
  texture.repeat.set(1, 1);
  if (imageAspect > targetAspect) {
    const visibleWidth = targetAspect / imageAspect;
    texture.repeat.x = visibleWidth;
    texture.offset.x = (1 - visibleWidth) / 2;
  } else if (imageAspect < targetAspect) {
    const visibleHeight = imageAspect / targetAspect;
    texture.repeat.y = visibleHeight;
    texture.offset.y = (1 - visibleHeight) / 2;
  }
  texture.needsUpdate = true;
}

function timestamp(windowRef) {
  return Number(windowRef?.performance?.now?.()) || Date.now();
}
