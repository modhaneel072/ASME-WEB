export const buildLineVertexShader = /* glsl */ `
  varying float vLineProgress;
  varying vec3 vViewNormal;

  void main() {
    vLineProgress = uv.x;
    vViewNormal = normalize(normalMatrix * normal);
    gl_Position = projectionMatrix * modelViewMatrix * vec4(position, 1.0);
  }
`;
export const buildLineFragmentShader = /* glsl */ `
  uniform vec3 uColor;
  uniform float uProgress;
  uniform float uOpacity;

  varying float vLineProgress;
  varying vec3 vViewNormal;

  void main() {
    float edge = max(fwidth(vLineProgress) * 2.0, 0.0025);
    float reveal = 1.0 - smoothstep(uProgress - edge, uProgress + edge, vLineProgress);
    if (reveal <= 0.001) discard;

    float restrainedLight = 0.72 + 0.28 * abs(vViewNormal.z);
    gl_FragColor = vec4(uColor * restrainedLight, reveal * uOpacity);
  }
`;
