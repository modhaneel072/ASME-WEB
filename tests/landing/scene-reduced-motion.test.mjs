import test from "node:test";
import assert from "node:assert/strict";

import { SceneController } from "../../static/js/landing/scene-controller.js";

test("reduced motion freezes rapid rotations and presents direct assembled states", () => {
  const scene = Object.create(SceneController.prototype);
  scene.reducedMotion = true;

  scene.roverChassis = { position: { y: 1 } };
  scene.roverWheels = [
    { rotation: { y: 10 } },
    { rotation: { y: -10 } },
  ];
  scene.iamLayerLines = { visible: true };
  scene._animateIam3d(0.5);
  assert.deepEqual(scene.roverWheels.map((wheel) => wheel.rotation.y), [0, 0]);
  assert.equal(scene.roverChassis.position.y, 0);
  assert.equal(scene.iamLayerLines.visible, false);

  const rib = { visible: false, scale: { z: 0 } };
  const part = { visible: false, scale: { value: 0, setScalar(value) { this.value = value; } } };
  scene.planeRoot = { rotation: { x: 1 }, position: { y: 1 } };
  scene.planeRibs = [{ mesh: rib, threshold: 0.8 }];
  scene.planeParts = [{ mesh: part, threshold: 0.8 }];
  scene.planePropeller = { rotation: { x: 12 } };
  scene._animatePlane(0.1);
  assert.equal(rib.visible, true);
  assert.equal(rib.scale.z, 1);
  assert.equal(part.visible, true);
  assert.equal(part.scale.value, 1);
  assert.equal(scene.planePropeller.rotation.x, 0);
  assert.equal(scene.planeRoot.rotation.x, 0);
  assert.equal(scene.planeRoot.position.y, 0);

  scene.dogRoot = { position: { x: 1, y: 1 } };
  scene.dogBuildParts = [
    { visible: false, userData: { buildStep: 1 } },
    { visible: false, userData: { buildStep: 10 } },
  ];
  scene.dogStatusLed = { visible: false };
  scene.dogLegs = [{ hip: { rotation: { z: 1 } }, knee: { rotation: { z: 1 } } }];
  scene._animateRoboticDog(0.1);
  assert.equal(scene.dogBuildParts.every((partEntry) => partEntry.visible), true);
  assert.equal(scene.dogStatusLed.visible, true);
  assert.equal(scene.dogRoot.position.x, 0);
  assert.equal(scene.dogLegs[0].hip.rotation.z, -0.08);
  assert.equal(scene.dogLegs[0].knee.rotation.z, 0.38);
});
