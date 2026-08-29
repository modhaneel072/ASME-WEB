/**
 * The Build Line camera path is intentionally data-first. Every camera move is
 * inspectable here instead of being scattered through scene animation code.
 */

export const SCENE_RANGES = Object.freeze({
  intro: Object.freeze({ start: 0.0, end: 0.08 }),
  design: Object.freeze({ start: 0.08, end: 0.18 }),
  iam3d: Object.freeze({ start: 0.18, end: 0.31 }),
  plane: Object.freeze({ start: 0.31, end: 0.43 }),
  rov: Object.freeze({ start: 0.43, end: 0.56 }),
  "student-design": Object.freeze({ start: 0.56, end: 0.68 }),
  "robotic-dog": Object.freeze({ start: 0.68, end: 0.82 }),
  people: Object.freeze({ start: 0.82, end: 0.92 }),
  join: Object.freeze({ start: 0.92, end: 1.0 }),
});

export const SCENE_KEYS = Object.freeze(Object.keys(SCENE_RANGES));

const EASING = Object.freeze({
  linear: (value) => value,
  mechanical: (value) => {
    const t = clamp01(value);
    return t * t * (3 - 2 * t);
  },
  settle: (value) => {
    const t = clamp01(value);
    return 1 - Math.pow(1 - t, 3);
  },
});

export const CAMERA_KEYFRAMES = Object.freeze([
  frame(0.0, [-7.5, 2.4, 8.8], [0.2, 0.15, 0], 0, 46, "intro", "settle"),
  frame(0.045, [-4.4, 1.7, 6.9], [2.5, 0.05, 0], 0, 44, "intro", "mechanical"),
  frame(0.08, [3.0, 1.55, 7.5], [9.2, 0.05, 0], -0.01, 43, "design", "mechanical"),
  frame(0.13, [5.8, 1.0, 5.6], [10.0, 0.0, 0], 0.015, 40, "design", "mechanical"),
  frame(0.18, [15.4, 1.25, 6.4], [21.4, 0.0, 0], 0, 45, "iam3d", "mechanical"),
  frame(0.265, [20.0, 0.65, 4.7], [24.1, -0.15, 0], -0.018, 43, "iam3d", "mechanical"),
  frame(0.31, [28.5, 0.0, 2.2], [34.4, 0.15, 0], 0, 48, "plane", "settle"),
  frame(0.37, [37.1, -1.7, 5.4], [43.0, 1.1, 0], 0.055, 46, "plane", "mechanical"),
  frame(0.43, [52.1, 3.8, 7.2], [59.0, 0.0, 0], -0.018, 47, "rov", "mechanical"),
  frame(0.485, [57.0, 0.45, 5.7], [63.0, -1.1, 0], 0, 44, "rov", "mechanical"),
  frame(0.56, [72.5, 1.15, 7.0], [78.8, 0.0, 0], 0, 44, "student-design", "mechanical"),
  frame(0.62, [76.3, 0.75, 5.4], [81.0, 0.0, 0], 0.012, 41, "student-design", "mechanical"),
  frame(0.68, [89.1, 1.4, 7.6], [95.5, 0.2, 0], 0, 45, "robotic-dog", "settle"),
  frame(0.755, [94.0, 0.9, 5.9], [99.4, 0.1, 0], -0.012, 43, "robotic-dog", "mechanical"),
  frame(0.82, [107.0, 1.7, 8.2], [114.0, 0.35, 0], 0, 47, "people", "settle"),
  frame(0.875, [112.2, 1.2, 6.8], [118.0, 0.25, 0], 0.01, 45, "people", "mechanical"),
  frame(0.92, [124.8, 1.0, 8.8], [132.0, 0.0, 0], 0, 46, "join", "settle"),
  frame(1.0, [129.5, 0.65, 9.8], [136.0, 0.05, 0], 0, 46, "join", "settle"),
]);

function frame(progress, position, target, roll, fov, scene, ease) {
  return Object.freeze({
    progress,
    position: Object.freeze(position),
    target: Object.freeze(target),
    roll,
    fov,
    scene,
    ease,
  });
}

export function clamp01(value) {
  return Math.min(1, Math.max(0, Number.isFinite(value) ? value : 0));
}

export function getSceneAtProgress(progress) {
  const value = clamp01(progress);
  for (let index = 0; index < SCENE_KEYS.length; index += 1) {
    const key = SCENE_KEYS[index];
    const range = SCENE_RANGES[key];
    if (value < range.end || index === SCENE_KEYS.length - 1) return key;
  }
  return "join";
}

export function getSceneLocalProgress(scene, progress) {
  const range = SCENE_RANGES[scene];
  if (!range) return 0;
  const duration = Math.max(Number.EPSILON, range.end - range.start);
  return clamp01((clamp01(progress) - range.start) / duration);
}

export function getSceneWeights(progress, blendRatio = 0.16) {
  const scene = getSceneAtProgress(progress);
  const index = SCENE_KEYS.indexOf(scene);
  const local = getSceneLocalProgress(scene, progress);
  const blend = Math.min(0.35, Math.max(0.02, blendRatio));
  const weights = Object.fromEntries(SCENE_KEYS.map((key) => [key, 0]));
  weights[scene] = 1;

  if (local < blend && index > 0) {
    const mix = local / blend;
    weights[scene] = mix;
    weights[SCENE_KEYS[index - 1]] = 1 - mix;
  } else if (local > 1 - blend && index < SCENE_KEYS.length - 1) {
    const mix = (local - (1 - blend)) / blend;
    weights[scene] = 1 - mix;
    weights[SCENE_KEYS[index + 1]] = mix;
  }

  return weights;
}

export function sampleCameraPath(progress, output = {}) {
  const value = clamp01(progress);
  let left = CAMERA_KEYFRAMES[0];
  let right = CAMERA_KEYFRAMES[CAMERA_KEYFRAMES.length - 1];

  for (let index = 1; index < CAMERA_KEYFRAMES.length; index += 1) {
    right = CAMERA_KEYFRAMES[index];
    if (right.progress >= value) {
      left = CAMERA_KEYFRAMES[index - 1];
      break;
    }
  }

  const span = Math.max(Number.EPSILON, right.progress - left.progress);
  const raw = clamp01((value - left.progress) / span);
  const ease = EASING[left.ease] || EASING.mechanical;
  const mix = ease(raw);

  output.position = interpolateVector(left.position, right.position, mix, output.position);
  output.target = interpolateVector(left.target, right.target, mix, output.target);
  output.roll = interpolateNumber(left.roll, right.roll, mix);
  output.fov = interpolateNumber(left.fov, right.fov, mix);
  output.scene = getSceneAtProgress(value);
  output.sceneWeight = getSceneWeights(value);
  output.ease = left.ease;
  return output;
}

function interpolateVector(from, to, mix, output = [0, 0, 0]) {
  output[0] = interpolateNumber(from[0], to[0], mix);
  output[1] = interpolateNumber(from[1], to[1], mix);
  output[2] = interpolateNumber(from[2], to[2], mix);
  return output;
}

function interpolateNumber(from, to, mix) {
  return from + (to - from) * mix;
}
