import type { SvgNode } from "./svg-elements.js";

/**
 * Spreadsheet icons and buttons as element descriptions (built with svgElement, never parsed from markup).
 * Internal module.
 */

/** Conditional-format icon sets drawn as small SVGs (shapes and colors follow Excel's icon families). */
export function iconSvg(set: string, index: number): SvgNode {
  const count = Number(set[0]) || 3, level = count - 1 - index; // level 0 = best
  const green = "#3e9a6a", yellow = "#f2c14e", red = "#d9534f", gray = "#8c8c8c", black = "#404040";
  const traffic = [green, yellow, red, black];
  const svg = (...content: SvgNode[]): SvgNode => ["svg", { viewBox: "0 0 16 16", width: "100%", height: "100%", "aria-hidden": "true" }, content];
  const arrow = (angle: number, color: string) => svg(["g", { transform: `rotate(${angle} 8 8)` }, [["path", { d: "M8 2 L13 8 H10 V14 H6 V8 H3 Z", fill: color }]]]);
  const arrowAngles = count === 3 ? [0, 90, 180] : count === 4 ? [0, 45, 135, 180] : [0, 45, 90, 135, 180];
  const colorFor = (position: number) => count === 3 ? [green, yellow, red][position]! : count === 4 ? [green, yellow, "#f08c3a", red][position]! : [green, "#9bc84f", yellow, "#f08c3a", red][position]!;
  if (/Arrows/.test(set)) return arrow(arrowAngles[level] ?? 90, /Gray/.test(set) ? gray : colorFor(level));
  if (/Triangles/.test(set)) return svg(level === 0 ? ["path", { d: "M8 3 L14 12 H2 Z", fill: green }] : level === count - 1 ? ["path", { d: "M8 13 L14 4 H2 Z", fill: red }] : ["rect", { x: 3, y: 7, width: 10, height: 3, fill: yellow }]);
  if (/Stars/.test(set)) {
    const fill = level === 0 ? "#f4b400" : level === 1 ? "url(#h)" : "none";
    return svg(["defs", {}, [["linearGradient", { id: "h" }, [["stop", { offset: "50%", "stop-color": "#f4b400" }], ["stop", { offset: "50%", "stop-color": "#fff" }]]]]], ["path", { d: "M8 1.5l1.9 4.2 4.6.4-3.5 3 1.1 4.5L8 11.2 3.9 13.6 5 9.1 1.5 6.1l4.6-.4z", fill, stroke: "#f4b400" }]);
  }
  if (/Rating|Boxes/.test(set)) { const filled = count - 1 - level; const bars: SvgNode[] = []; for (let bar = 0; bar < 4; bar++) bars.push(["rect", { x: 2 + bar * 3.2, y: 11 - bar * 3, width: 2.4, height: 3 + bar * 3, fill: bar < filled ? "#3f6fb5" : "#d0d7e5" }]); return svg(...bars); }
  if (/Quarters/.test(set)) {
    const portion = (count - 1 - level) / 4; const angle = portion * 2 * Math.PI, x = 8 + 6 * Math.sin(angle), y = 8 - 6 * Math.cos(angle);
    const outline: SvgNode = ["circle", { cx: 8, cy: 8, r: 6, fill: "#fff", stroke: black }];
    return portion >= 1 ? svg(outline, ["circle", { cx: 8, cy: 8, r: 6, fill: black }]) : portion > 0 ? svg(outline, ["path", { d: `M8 8 V2 A6 6 0 ${portion > 0.5 ? 1 : 0} 1 ${x} ${y} Z`, fill: black }]) : svg(outline);
  }
  if (/Symbols/.test(set)) {
    const color = [green, yellow, red][level] ?? red;
    // 3Symbols2 draws the mark alone, in the symbol's color; the other sets draw a white mark on a colored circle.
    const ink = set === "3Symbols2" ? color : "#fff";
    const mark: SvgNode[] = level === 0 ? [["path", { d: "M4.5 8.2l2.3 2.3 4.7-4.8", stroke: ink, "stroke-width": 1.8, fill: "none" }]]
      : level === 1 ? [["rect", { x: 7.1, y: 3.8, width: 1.8, height: 5.2, fill: ink }], ["rect", { x: 7.1, y: 10.4, width: 1.8, height: 1.8, fill: ink }]]
      : [["path", { d: "M5.2 5.2l5.6 5.6M10.8 5.2l-5.6 5.6", stroke: ink, "stroke-width": 1.8 }]];
    return set === "3Symbols2" ? svg(...mark) : svg(["circle", { cx: 8, cy: 8, r: 6.5, fill: color }], ...mark);
  }
  if (/Flags/.test(set)) return svg(["path", { d: "M4 2v12", stroke: black, "stroke-width": 1.2 }], ["path", { d: "M4.6 2.5h8l-2 3 2 3h-8z", fill: [green, yellow, red][level] ?? red }]);
  if (/Signs/.test(set)) return svg(level === 0 ? ["circle", { cx: 8, cy: 8, r: 6, fill: green }] : level === 1 ? ["path", { d: "M8 2 L14 13 H2 Z", fill: yellow }] : ["path", { d: "M8 2 L14 8 L8 14 L2 8 Z", fill: red }]);
  const color = traffic[Math.min(level, traffic.length - 1)]!;
  const light: SvgNode = ["circle", { cx: 8, cy: 8, r: /TrafficLights2/.test(set) ? 5.2 : 6.5, fill: count === 4 && level === 3 ? black : color }];
  return /TrafficLights2/.test(set) ? svg(["circle", { cx: 8, cy: 8, r: 7, fill: black }], light) : svg(light);
}

/** The filter button of a table or AutoFilter header cell. */
export const filterButtonSvg: SvgNode = ["svg", { viewBox: "0 0 16 16", width: "100%", height: "100%", "aria-hidden": "true" }, [
  ["rect", { x: 0.5, y: 0.5, width: 15, height: 15, fill: "#fff", stroke: "#a6a6a6" }],
  ["path", { d: "M4.5 6.5h7L8 10.5z", fill: "#444" }],
]];

/** A slicer's multi-select and clear-filter buttons; the clear button's cross turns red while the slicer filters. */
export function slicerIconsSvg(filtered: boolean): SvgNode {
  return ["svg", { viewBox: "0 0 48 16", width: "100%", height: "100%" }, [
    ["path", { d: "M1.5 3.5l1.5 1.5 3-3M1.5 8.5l1.5 1.5 3-3M1.5 13.5l1.5 1.5 3-3", fill: "none", stroke: "#595959" }],
    ["path", { d: "M8 4.5h8M8 9.5h8M8 14.5h8", stroke: "#595959" }],
    ["path", { d: "M32.5 2.5h11l-4.5 5v5l-2-1.2v-3.8z", fill: "none", stroke: "#7f7f7f" }],
    ["path", { d: "M41.5 10.5l4 4M45.5 10.5l-4 4", stroke: filtered ? "#e81123" : "#a6a6a6", "stroke-width": 1.2 }],
  ]];
}
