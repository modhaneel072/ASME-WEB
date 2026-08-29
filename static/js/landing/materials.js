import {
  AdditiveBlending,
  Color,
  DoubleSide,
  LineBasicMaterial,
  MeshPhysicalMaterial,
  MeshStandardMaterial,
  NormalBlending,
} from "../../vendor/three-0.185.1/three.module.min.js";

export const LANDING_COLORS = Object.freeze({
  iowaGold: 0xffcd00,
  warmGold: 0xd6a900,
  aluminum: 0x9b9da0,
  steel: 0x24272b,
  polymer: 0x17191c,
  printed: 0x3c3f42,
  rubber: 0x090a0b,
  drafting: 0x59616a,
  overload: 0xd8463b,
});

export function createMaterialLibrary({ quality = "high" } = {}) {
  const low = quality === "low";
  const common = { envMapIntensity: low ? 0.18 : 0.32 };

  const library = {
    machinedAluminum: new MeshStandardMaterial({
      ...common,
      name: "machinedAluminum",
      color: LANDING_COLORS.aluminum,
      metalness: 0.72,
      roughness: 0.38,
    }),
    darkSteel: new MeshStandardMaterial({
      ...common,
      name: "darkSteel",
      color: LANDING_COLORS.steel,
      metalness: 0.62,
      roughness: 0.48,
    }),
    mattePolymer: new MeshStandardMaterial({
      ...common,
      name: "mattePolymer",
      color: LANDING_COLORS.polymer,
      metalness: 0.02,
      roughness: 0.84,
    }),
    printedPolymer: new MeshStandardMaterial({
      ...common,
      name: "printedPolymer",
      color: LANDING_COLORS.printed,
      metalness: 0.04,
      roughness: 0.78,
    }),
    rubber: new MeshStandardMaterial({
      ...common,
      name: "rubber",
      color: LANDING_COLORS.rubber,
      metalness: 0,
      roughness: 0.96,
    }),
    glassHousing: new MeshPhysicalMaterial({
      name: "glassHousing",
      color: 0x9aa8ad,
      metalness: 0,
      roughness: 0.18,
      transmission: low ? 0 : 0.34,
      transparent: true,
      opacity: low ? 0.42 : 0.58,
      thickness: 0.12,
      side: DoubleSide,
      depthWrite: false,
    }),
    buildLineGold: new MeshStandardMaterial({
      ...common,
      name: "buildLineGold",
      color: LANDING_COLORS.iowaGold,
      emissive: new Color(LANDING_COLORS.warmGold),
      emissiveIntensity: low ? 0.06 : 0.12,
      metalness: 0.26,
      roughness: 0.46,
    }),
    buildLineWire: new LineBasicMaterial({
      name: "buildLineWire",
      color: LANDING_COLORS.iowaGold,
      transparent: true,
      opacity: 0.9,
      depthWrite: false,
    }),
    draftingLine: new LineBasicMaterial({
      name: "draftingLine",
      color: LANDING_COLORS.drafting,
      transparent: true,
      opacity: 0.46,
      blending: NormalBlending,
      depthWrite: false,
    }),
    stressOverlay: new MeshStandardMaterial({
      name: "stressOverlay",
      color: LANDING_COLORS.overload,
      emissive: new Color(LANDING_COLORS.overload),
      emissiveIntensity: 0.18,
      metalness: 0.08,
      roughness: 0.62,
      transparent: true,
      opacity: 0,
      blending: quality === "high" ? AdditiveBlending : NormalBlending,
      depthWrite: false,
    }),
  };

  Object.values(library).forEach((material) => {
    material.userData.landingSharedMaterial = true;
  });
  return library;
}

export function disposeMaterialLibrary(library) {
  if (!library) return;
  for (const material of Object.values(library)) material?.dispose?.();
}
