import test from "node:test";
import assert from "node:assert/strict";

import {
  CAMERA_KEYFRAMES,
  SCENE_KEYS,
  SCENE_RANGES,
  getSceneAtProgress,
  getSceneLocalProgress,
  getSceneWeights,
  sampleCameraPath,
} from "../../static/js/landing/camera-path.js";

test("master scene ranges are continuous and preserve the required boundaries", () => {
  assert.deepEqual(SCENE_KEYS, [
    "intro", "design", "iam3d", "plane", "rov", "student-design", "robotic-dog", "people", "join",
  ]);
  assert.equal(SCENE_RANGES.intro.start, 0);
  assert.equal(SCENE_RANGES.join.end, 1);

  for (let index = 1; index < SCENE_KEYS.length; index += 1) {
    assert.equal(
      SCENE_RANGES[SCENE_KEYS[index - 1]].end,
      SCENE_RANGES[SCENE_KEYS[index]].start,
      `${SCENE_KEYS[index - 1]} and ${SCENE_KEYS[index]} must touch without a gap`,
    );
  }
});

test("scene selection changes exactly at the normalized range boundaries", () => {
  const expected = [
    [0, "intro"],
    [0.08, "design"],
    [0.18, "iam3d"],
    [0.31, "plane"],
    [0.43, "rov"],
    [0.56, "student-design"],
    [0.68, "robotic-dog"],
    [0.82, "people"],
    [0.92, "join"],
    [1, "join"],
  ];
  for (const [progress, scene] of expected) assert.equal(getSceneAtProgress(progress), scene);
  assert.ok(Math.abs(getSceneLocalProgress("rov", 0.495) - 0.5) < 1e-12);
});

test("camera path remains ordered, inspectable, and bounded", () => {
  for (let index = 1; index < CAMERA_KEYFRAMES.length; index += 1) {
    assert.ok(CAMERA_KEYFRAMES[index].progress > CAMERA_KEYFRAMES[index - 1].progress);
  }
  const start = sampleCameraPath(0);
  const end = sampleCameraPath(1);
  assert.deepEqual(start.position, [-7.5, 2.4, 8.8]);
  assert.deepEqual(end.position, [129.5, 0.65, 9.8]);
  assert.equal(start.scene, "intro");
  assert.equal(end.scene, "join");
  assert.ok(start.fov >= 40 && start.fov <= 50);
  assert.ok(end.fov >= 40 && end.fov <= 50);
});

test("scene transition weights always sum to one", () => {
  for (let step = 0; step <= 1000; step += 1) {
    const weights = getSceneWeights(step / 1000);
    const total = Object.values(weights).reduce((sum, value) => sum + value, 0);
    assert.ok(Math.abs(total - 1) < 1e-9, `weights totaled ${total} at ${step / 1000}`);
    assert.ok(Object.values(weights).every((value) => value >= 0 && value <= 1));
  }
});
