/** SVG outlines for DrawingML preset geometries (ECMA-376 presetShapeDefinitions), in shape-local units. */
export type PresetPath = { d: string; filled: boolean };

const presetAliases: Record<string, string> = {
  flowChartProcess: "rect", flowChartAlternateProcess: "roundRect", flowChartDecision: "diamond", flowChartConnector: "ellipse",
  flowChartTerminator: "pill", flowChartInputOutput: "parallelogram", flowChartManualOperation: "flowChartManualOperation",
  straightConnector1: "line", round1Rect: "round1Rect", wedgeRectCallout: "rect", textBox: "rect",
};

/** Returns undefined for presets that are not modelled, so callers can fall back to a rectangle. */
export function presetGeometryPath(name: string | undefined, w: number, h: number, adjustments: Record<string, number> = {}): PresetPath[] | undefined {
  const preset = presetAliases[name ?? "rect"] ?? name ?? "rect";
  const ss = Math.max(0, Math.min(w, h)), hc = w / 2, vc = h / 2;
  const adj = (key: string, fallback: number, min = 0, max = 100_000) => Math.max(min, Math.min(max, adjustments[key] ?? (key === "adj1" ? adjustments["adj"] : undefined) ?? fallback));
  const poly = (...points: Array<[number, number]>) => [{ d: `M ${points.map(([x, y]) => `${x} ${y}`).join(" L ")} Z`, filled: true }];
  const arc = (r: number, x: number, y: number) => r > 0 ? `A ${r} ${r} 0 0 1 ${x} ${y}` : `L ${x} ${y}`;
  const corners = (tl: number, tr: number, br: number, bl: number) => [{
    d: `M ${tl} 0 L ${w - tr} 0 ${arc(tr, w, tr)} L ${w} ${h - br} ${arc(br, w - br, h)} L ${bl} ${h} ${arc(bl, 0, h - bl)} L 0 ${tl} ${arc(tl, tl, 0)} Z`,
    filled: true,
  }];
  const snips = (tl: number, tr: number, br: number, bl: number) => poly([tl, 0], [w - tr, 0], [w, tr], [w, h - br], [w - br, h], [bl, h], [0, h - bl], [0, tl]);
  switch (preset) {
    case "rect": return poly([0, 0], [w, 0], [w, h], [0, h]);
    case "line": return [{ d: `M 0 0 L ${w} ${h}`, filled: false }];
    case "roundRect": { const r = ss * adj("adj", 16_667, 0, 50_000) / 100_000; return corners(r, r, r, r); }
    case "pill": { const r = ss / 2; return corners(r, r, r, r); }
    case "round1Rect": { const r = ss * adj("adj", 16_667, 0, 50_000) / 100_000; return corners(0, r, 0, 0); }
    case "round2SameRect": { const t = ss * adj("adj1", 16_667, 0, 50_000) / 100_000, b = ss * adj("adj2", 0, 0, 50_000) / 100_000; return corners(t, t, b, b); }
    case "round2DiagRect": { const a = ss * adj("adj1", 16_667, 0, 50_000) / 100_000, b = ss * adj("adj2", 0, 0, 50_000) / 100_000; return corners(a, b, a, b); }
    case "snip1Rect": { const s = ss * adj("adj", 16_667, 0, 50_000) / 100_000; return snips(0, s, 0, 0); }
    case "snip2SameRect": { const t = ss * adj("adj1", 16_667, 0, 50_000) / 100_000, b = ss * adj("adj2", 0, 0, 50_000) / 100_000; return snips(t, t, b, b); }
    case "snip2DiagRect": { const a = ss * adj("adj1", 0, 0, 50_000) / 100_000, b = ss * adj("adj2", 16_667, 0, 50_000) / 100_000; return snips(a, b, a, b); }
    case "snipRoundRect": { const r = ss * adj("adj1", 16_667, 0, 50_000) / 100_000, s = ss * adj("adj2", 16_667, 0, 50_000) / 100_000; return [{ d: `M ${r} 0 L ${w - s} 0 L ${w} ${s} L ${w} ${h} L 0 ${h} L 0 ${r} ${arc(r, r, 0)} Z`, filled: true }]; }
    case "ellipse": case "arc": case "donut": case "blockArc": case "chord": case "pie":
      return [{ d: `M 0 ${vc} A ${hc} ${vc} 0 1 1 ${w} ${vc} A ${hc} ${vc} 0 1 1 0 ${vc} Z`, filled: true }];
    case "teardrop": {
      // Round on three quarters, with the point towards the top-right corner (adj 100000 puts it on the corner).
      const a = adj("adj", 100_000, 0, 200_000) / 100_000, x1 = hc + hc * a, y1 = vc - vc * a, x2 = (hc + x1) / 2, y2 = (vc + y1) / 2;
      return [{ d: `M 0 ${vc} A ${hc} ${vc} 0 0 1 ${hc} 0 Q ${x2} 0 ${x1} ${y1} Q ${w} ${y2} ${w} ${vc} A ${hc} ${vc} 0 0 1 ${hc} ${h} A ${hc} ${vc} 0 0 1 0 ${vc} Z`, filled: true }];
    }
    case "triangle": { const x = w * adj("adj", 50_000) / 100_000; return poly([x, 0], [w, h], [0, h]); }
    case "rtTriangle": return poly([0, 0], [w, h], [0, h]);
    case "diamond": return poly([hc, 0], [w, vc], [hc, h], [0, vc]);
    case "parallelogram": { const x = ss * adj("adj", 25_000, 0, 100_000 * w / Math.max(1, ss)) / 100_000; return poly([x, 0], [w, 0], [w - x, h], [0, h]); }
    case "trapezoid": { const x = ss * adj("adj", 25_000, 0, 50_000 * w / Math.max(1, ss)) / 100_000; return poly([0, h], [x, 0], [w - x, 0], [w, h]); }
    case "flowChartManualOperation": return poly([0, 0], [w, 0], [w * 0.8, h], [w * 0.2, h]);
    case "hexagon": { const x = ss * adj("adj", 25_000, 0, 50_000 * w / Math.max(1, ss)) / 100_000; return poly([0, vc], [x, 0], [w - x, 0], [w, vc], [w - x, h], [x, h]); }
    case "octagon": { const x = ss * adj("adj", 29_289, 0, 50_000) / 100_000; return poly([0, x], [x, 0], [w - x, 0], [w, x], [w, h - x], [w - x, h], [x, h], [0, h - x]); }
    case "pentagon": return poly([hc, 0], [w, h * 0.3633], [w * 0.8090, h], [w * 0.1910, h], [0, h * 0.3633]);
    case "homePlate": { const x = w - ss * adj("adj", 50_000, 0, 100_000 * w / Math.max(1, ss)) / 100_000; return poly([0, 0], [x, 0], [w, vc], [x, h], [0, h]); }
    case "chevron": { const d = ss * adj("adj", 50_000, 0, 100_000 * w / Math.max(1, ss)) / 100_000; return poly([0, 0], [w - d, 0], [w, vc], [w - d, h], [0, h], [d, vc]); }
    case "plus": { const x = ss * adj("adj", 25_000, 0, 50_000) / 100_000; return poly([x, 0], [w - x, 0], [w - x, x], [w, x], [w, h - x], [w - x, h - x], [w - x, h], [x, h], [x, h - x], [0, h - x], [0, x], [x, x]); }
    case "rightArrow": case "leftArrow": {
      const half = h * adj("adj1", 50_000) / 200_000, head = ss * adj("adj2", 50_000, 0, 100_000 * w / Math.max(1, ss)) / 100_000;
      const points: Array<[number, number]> = [[0, vc - half], [w - head, vc - half], [w - head, 0], [w, vc], [w - head, h], [w - head, vc + half], [0, vc + half]];
      return poly(...(preset === "leftArrow" ? points.map(([x, y]) => [w - x, y] as [number, number]) : points));
    }
    case "downArrow": case "upArrow": {
      const half = w * adj("adj1", 50_000) / 200_000, head = ss * adj("adj2", 50_000, 0, 100_000 * h / Math.max(1, ss)) / 100_000;
      const points: Array<[number, number]> = [[hc - half, 0], [hc - half, h - head], [0, h - head], [hc, h], [w, h - head], [hc + half, h - head], [hc + half, 0]];
      return poly(...(preset === "upArrow" ? points.map(([x, y]) => [x, h - y] as [number, number]) : points));
    }
    case "leftRightArrow": {
      const half = h * adj("adj1", 50_000) / 200_000, head = ss * adj("adj2", 50_000, 0, 50_000 * w / Math.max(1, ss)) / 100_000;
      return poly([0, vc], [head, 0], [head, vc - half], [w - head, vc - half], [w - head, 0], [w, vc], [w - head, h], [w - head, vc + half], [head, vc + half], [head, h]);
    }
    case "star4": case "star5": case "star6": case "star8": case "star10": case "star12": {
      const count = Number(preset.slice(4)), inner = adj("adj", ({ 4: 12_500, 5: 19_098, 6: 28_868, 8: 38_250, 10: 42_533, 12: 37_500 } as Record<number, number>)[count] ?? 19_098, 0, 50_000) / 50_000;
      const points: Array<[number, number]> = [];
      for (let index = 0; index < count * 2; index++) {
        const angle = -Math.PI / 2 + index * Math.PI / count, radius = index % 2 ? inner : 1;
        points.push([hc + Math.cos(angle) * hc * radius, vc + Math.sin(angle) * vc * radius]);
      }
      return poly(...points);
    }
    case "frame": {
      const t = ss * adj("adj1", 12_500, 0, 50_000) / 100_000;
      return [{ d: `M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z M ${t} ${t} L ${t} ${h - t} L ${w - t} ${h - t} L ${w - t} ${t} Z`, filled: true }];
    }
    case "bentConnector2": return [{ d: `M 0 0 L ${w} 0 L ${w} ${h}`, filled: false }];
    case "bentConnector3": { const x = w * (adjustments["adj1"] ?? 50_000) / 100_000; return [{ d: `M 0 0 L ${x} 0 L ${x} ${h} L ${w} ${h}`, filled: false }]; }
    case "bentConnector4": { const x = w * (adjustments["adj1"] ?? 50_000) / 100_000, y = h * (adjustments["adj2"] ?? 50_000) / 100_000; return [{ d: `M 0 0 L ${x} 0 L ${x} ${y} L ${w} ${y} L ${w} ${h}`, filled: false }]; }
    case "curvedConnector2": return [{ d: `M 0 0 C ${w / 2} 0 ${w} ${h / 2} ${w} ${h}`, filled: false }];
    case "curvedConnector3": { const x = w * (adjustments["adj1"] ?? 50_000) / 100_000; return [{ d: `M 0 0 C ${x / 2} 0 ${x} ${h / 4} ${x} ${vc} C ${x} ${h * 3 / 4} ${(x + w) / 2} ${h} ${w} ${h}`, filled: false }]; }
    default: return undefined;
  }
}
