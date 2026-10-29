import type { Paragraph, Presentation, PptxCamera, PptxElement, PptxFill, PptxImageElement, PptxLine, PptxLineElement, PptxReflection, PptxShapeElement, PptxTextElement } from "../model.js";
import { fontStack } from "./fonts.js";
import { presetGeometryPath } from "./ppt-geometry.js";

/** `runFontLines`: paragraph line boxes take the first run's font, so line spacing follows that font as in Office (spreadsheet drawings). */
export type DrawingRenderOptions = { runFontLines?: boolean };

/** Renders a positioned DrawingML element (slide content, spreadsheet drawings) inside a container sized to its frame; lengths use cqw of that container. */
export function renderPptxElement(element: PptxElement, slideData: Presentation["slides"][number], imageUrls: string[], options: DrawingRenderOptions = {}): HTMLElement {
  const width = Math.max(1, slideData.width ?? 9_144_000), height = Math.max(1, slideData.height ?? 5_143_500);
  const position = (node: HTMLElement) => {
    node.style.position = "absolute";
    node.style.left = `${element.left / width * 100}%`; node.style.top = `${element.top / height * 100}%`;
    node.style.width = `${element.width / width * 100}%`; node.style.height = `${element.height / height * 100}%`;
    const flipH = "flipH" in element && element.flipH, flipV = "flipV" in element && element.flipV;
    const camera = element.kind === "text" ? pptCameraTransform(element.camera, width) : "";
    const transform = [element.rotation ? `rotate(${element.rotation}deg)` : "", camera, flipH || flipV ? `scale(${flipH ? -1 : 1}, ${flipV ? -1 : 1})` : ""].filter(Boolean).join(" ");
    if (transform) node.style.transform = transform;
    node.style.transformOrigin = "center center";
  };
  if (element.kind === "image") {
    const frame = document.createElement("div"); frame.className = "ufp-pptx-image"; position(frame); frame.style.overflow = "hidden";
    const image = document.createElement("img"), bytes = new Uint8Array(element.image.bytes.byteLength); bytes.set(element.image.bytes);
    const url = URL.createObjectURL(new Blob([bytes.buffer], { type: element.image.mimeType })); imageUrls.push(url);
    image.src = url; image.alt = element.image.alt ?? ""; image.draggable = false;
    image.style.position = "absolute"; image.style.maxWidth = "none"; image.style.maxHeight = "none";
    if (element.duotone) {
      const namespace = "http://www.w3.org/2000/svg", filterId = `ufp-pptx-duotone-${Math.random().toString(36).slice(2)}`;
      const svg = document.createElementNS(namespace, "svg"), defs = document.createElementNS(namespace, "defs"), filter = document.createElementNS(namespace, "filter");
      svg.setAttribute("aria-hidden", "true"); svg.style.position = "absolute"; svg.style.width = "0"; svg.style.height = "0"; svg.style.overflow = "hidden";
      filter.id = filterId; filter.setAttribute("color-interpolation-filters", "sRGB");
      const grayscale = document.createElementNS(namespace, "feColorMatrix"); grayscale.setAttribute("type", "saturate"); grayscale.setAttribute("values", "0"); grayscale.setAttribute("result", "ufp-gray"); filter.append(grayscale);
      const transfer = document.createElementNS(namespace, "feComponentTransfer"); transfer.setAttribute("in", "ufp-gray");
      const dark = pptColorChannels(element.duotone.darkColor), light = pptColorChannels(element.duotone.lightColor);
      for (const [channel, index] of [["R", 0], ["G", 1], ["B", 2]] as const) {
        const fn = document.createElementNS(namespace, `feFunc${channel}`);
        fn.setAttribute("type", "table"); fn.setAttribute("tableValues", `${(dark[index] / 255).toFixed(4)} ${(light[index] / 255).toFixed(4)}`); transfer.append(fn);
      }
      const alpha = document.createElementNS(namespace, "feFuncA"); alpha.setAttribute("type", "identity"); transfer.append(alpha); filter.append(transfer); defs.append(filter); svg.append(defs); frame.append(svg);
      image.style.filter = `url(#${filterId})`;
    }
    if (element.crop) {
      const cropWidth = Math.max(0.01, 1 - element.crop.left - element.crop.right), cropHeight = Math.max(0.01, 1 - element.crop.top - element.crop.bottom);
      image.style.width = `${100 / cropWidth}%`; image.style.height = `${100 / cropHeight}%`;
      image.style.left = `${-element.crop.left / cropWidth * 100}%`; image.style.top = `${-element.crop.top / cropHeight * 100}%`;
    } else { image.style.inset = "0"; image.style.width = "100%"; image.style.height = "100%"; }
    if (element.geometry) clipPptxImageToGeometry(frame, element);
    frame.append(image); return frame;
  }
  if (element.kind === "table") {
    const table = document.createElement("div"); table.className = "ufp-pptx-table"; position(table);
    const columns = element.columns.length ? element.columns : element.rows[0]?.cells.map(() => 1) ?? [1];
    const columnTotal = columns.reduce((sum, value) => sum + value, 0) || columns.length;
    table.style.display = "grid"; table.style.gridTemplateColumns = columns.map(value => `${Math.max(0.1, value) / columnTotal * 100}%`).join(" ");
    const rowTotal = element.rows.reduce((sum, row) => sum + (row.height || 1), 0) || element.rows.length;
    table.style.gridTemplateRows = element.rows.map(row => `${(row.height || 1) / rowTotal * 100}%`).join(" ");
    for (const row of element.rows) for (const cellData of row.cells) {
      const cell = document.createElement("div"); cell.className = "ufp-pptx-cell";
      if (cellData.fill) cell.style.background = pptFillCss(cellData.fill);
      cell.style.display = "flex"; cell.style.flexDirection = "column"; cell.style.justifyContent = cellData.verticalAlign === "center" ? "center" : cellData.verticalAlign === "bottom" ? "flex-end" : "flex-start";
      cell.style.padding = `${cellData.margins.top / width * 100}cqw ${cellData.margins.right / width * 100}cqw ${cellData.margins.bottom / width * 100}cqw ${cellData.margins.left / width * 100}cqw`;
      for (const [side, line] of Object.entries(cellData.borders ?? {})) {
        const key = `border${side[0]!.toUpperCase()}${side.slice(1)}` as "borderTop" | "borderRight" | "borderBottom" | "borderLeft";
        cell.style[key] = pptBorderCss(line, width);
      }
      renderPptxParagraphs(cell, cellData.paragraphs, width, cellData.alignment);
      table.append(cell);
    }
    return table;
  }
  if (element.kind === "line" || element.kind === "shape") return renderPptxGeometry(element, slideData, imageUrls);
  const textBox = document.createElement("div"); textBox.className = "ufp-pptx-text"; position(textBox);
  // PowerPoint never clips text to its box; overflowing text stays visible.
  textBox.style.overflow = "visible";
  if (element.warp && appendPptWarpedText(textBox, element)) return textBox;
  textBox.style.whiteSpace = element.wrap ? "pre-wrap" : "pre";
  textBox.style.display = "flex"; textBox.style.flexDirection = "column";
  textBox.style.justifyContent = element.verticalAlign === "center" ? "center" : element.verticalAlign === "bottom" ? "flex-end" : "flex-start";
  textBox.style.padding = `${element.margins.top / width * 100}cqw ${element.margins.right / width * 100}cqw ${element.margins.bottom / width * 100}cqw ${element.margins.left / width * 100}cqw`;
  renderPptxParagraphs(textBox, element.paragraphs, width, undefined, options);
  fitPptxSingleTokenRuns(textBox, element, width);
  if (element.reflection) appendPptReflection(textBox, element.reflection, width);
  return textBox;
}

/** CSS for a DrawingML 3D camera: the face turns in its own frame (before the shape's 2D rotation) and is seen from the camera distance. */
function pptCameraTransform(camera: PptxCamera | undefined, frameWidth: number): string {
  if (!camera) return "";
  const turns = [camera.revolution % 360 ? `rotateZ(${-camera.revolution}deg)` : "", camera.longitude % 360 ? `rotateY(${-camera.longitude}deg)` : "", camera.latitude % 360 ? `rotateX(${-camera.latitude}deg)` : ""].filter(Boolean);
  if (!turns.length) return "";
  return [camera.distance ? `perspective(${camera.distance / frameWidth * 100}cqw)` : "", ...turns].filter(Boolean).join(" ");
}

/**
 * Office lights a turned face towards white: about 44% at a 50° tilt and 30% at 60° at its left edge, fading along its width
 * (calibrated against Excel's perspectiveRelaxed scenes with the three-point rig).
 */
function pptSceneLight(element: PptxShapeElement): { start: number; end: number } | undefined {
  const camera = element.camera;
  if (!camera?.lit) return undefined;
  const radians = (degrees: number) => degrees * Math.PI / 180;
  const tilt = Math.acos(Math.min(1, Math.abs(Math.cos(radians(camera.latitude)) * Math.cos(radians(camera.longitude))))) * 180 / Math.PI;
  const start = Math.max(0.1, Math.min(0.5, 0.3 + 0.014 * (60.4 - tilt)));
  return { start, end: Math.max(0.08, start - 0.22 * element.width / 9_525_000) };
}

function pptLighten(color: string, amount: number): string {
  return `#${pptColorChannels(color).map(channel => Math.round(channel + (255 - channel) * amount).toString(16).padStart(2, "0")).join("")}`;
}

/** Text on a preset warp path (WordArt "Arch Up"/"Arch Down"); returns false for warps that are drawn as plain text. */
function appendPptWarpedText(textBox: HTMLElement, element: PptxTextElement): boolean {
  if (element.warp !== "textArchUp" && element.warp !== "textArchDown") return false;
  const runs = element.paragraphs.flatMap((paragraph, index) => [...(index ? [{ ...paragraph.defaultRun, text: " " }] : []), ...paragraph.runs.filter(run => run.text && run.text !== "\n")]);
  if (!runs.some(run => run.text.trim())) return false;
  // The warp path is the upper or lower half of the ellipse inscribed in the shape box (insets do not apply); glyphs sit with their
  // x-height middle on it, so advances are measured there (fitted against Excel's rendering).
  // Coordinates are in points: browsers cap font sizes far below EMU-sized glyphs.
  const w = Math.max(0.1, element.width / 12_700), h = Math.max(0.1, element.height / 12_700);
  const up = element.warp === "textArchUp", alignment = element.paragraphs[0]?.alignment ?? "left";
  const svg = document.createElementNS(svgNamespace, "svg"), path = document.createElementNS(svgNamespace, "path"), id = pptUniqueId("warp");
  svg.setAttribute("viewBox", `0 0 ${w} ${h}`); svg.setAttribute("preserveAspectRatio", "none");
  svg.style.position = "absolute"; svg.style.inset = "0"; svg.style.width = "100%"; svg.style.height = "100%"; svg.style.overflow = "visible";
  path.id = id; path.setAttribute("fill", "none");
  path.setAttribute("d", `M 0 ${h / 2} A ${w / 2} ${h / 2} 0 0 ${up ? 1 : 0} ${w} ${h / 2}`);
  const text = document.createElementNS(svgNamespace, "text"), textPath = document.createElementNS(svgNamespace, "textPath");
  textPath.setAttribute("href", `#${id}`);
  textPath.setAttribute("startOffset", alignment === "center" ? "50%" : alignment === "right" ? "100%" : "0%");
  text.setAttribute("text-anchor", alignment === "center" ? "middle" : alignment === "right" ? "end" : "start");
  text.setAttribute("dy", String((runs[0]?.fontSize ?? 18) * 0.24));
  for (const run of runs) {
    const span = document.createElementNS(svgNamespace, "tspan"); span.textContent = run.text.replace(/\n/g, " ");
    span.setAttribute("font-size", String(run.fontSize ?? 18));
    span.setAttribute("font-family", fontStack(run.fontFamily ?? "Arial"));
    span.setAttribute("fill", run.color ?? "#000000");
    if (run.bold) span.setAttribute("font-weight", "700");
    if (run.italic) span.setAttribute("font-style", "italic");
    textPath.append(span);
  }
  text.append(textPath); svg.append(path, text); textBox.append(svg);
  textBox.setAttribute("aria-label", runs.map(run => run.text).join(""));
  return true;
}

/** Mirror copy below the box (effectLst/reflection), fading from the start to the end opacity. */
function appendPptReflection(textBox: HTMLElement, reflection: PptxReflection, slideWidth: number): void {
  const mirror = textBox.cloneNode(true) as HTMLElement;
  mirror.classList.add("ufp-pptx-reflection"); mirror.setAttribute("aria-hidden", "true");
  Object.assign(mirror.style, { left: "0", top: `calc(100% + ${reflection.distance / slideWidth * 100}cqw)`, width: "100%", height: "100%", transform: "scaleY(-1)", pointerEvents: "none" });
  const end = reflection.endPosition * 100;
  // The mask applies before the flip, so it fades from the mirrored edge (the copy's local bottom) upwards.
  const mask = `linear-gradient(to top, rgba(0,0,0,${reflection.startOpacity}) 0%, rgba(0,0,0,${reflection.endOpacity}) ${end}%, transparent ${end}%)`;
  mirror.style.setProperty("mask-image", mask); mirror.style.setProperty("-webkit-mask-image", mask);
  if (reflection.blur > 0) mirror.style.filter = `blur(${reflection.blur / 2 / slideWidth * 100}cqw)`;
  textBox.append(mirror);
}

function appendPptSvgGradient(svg: SVGSVGElement, id: string, value: string): boolean {
  const namespace = "http://www.w3.org/2000/svg";
  const linear = /^linear-gradient\(\s*([-\d.]+)deg\s*,([\s\S]*)\)$/i.exec(value);
  const radial = /^radial-gradient\(\s*(?:circle|ellipse)[^,]*?at\s+([\d.]+)%\s+([\d.]+)%\s*,([\s\S]*)\)$/i.exec(value);
  const body = linear?.[2] ?? radial?.[3];
  if (!body) return false;
  const parts: string[] = []; let depth = 0, current = "";
  for (const char of body) {
    if (char === "(" ) depth++;
    else if (char === ")") depth--;
    if (char === "," && depth === 0) { parts.push(current.trim()); current = ""; } else current += char;
  }
  if (current.trim()) parts.push(current.trim());
  let gradient: SVGElement;
  if (linear) {
    // CSS angles run clockwise from "to top"; convert to a direction in SVG's y-down bounding box.
    const angle = (Number(linear[1]) - 90) * Math.PI / 180, dx = Math.cos(angle) * 0.5, dy = Math.sin(angle) * 0.5;
    gradient = document.createElementNS(namespace, "linearGradient");
    gradient.setAttribute("x1", String(0.5 - dx)); gradient.setAttribute("y1", String(0.5 - dy));
    gradient.setAttribute("x2", String(0.5 + dx)); gradient.setAttribute("y2", String(0.5 + dy));
  } else {
    const cx = Number(radial![1]) / 100, cy = Number(radial![2]) / 100;
    gradient = document.createElementNS(namespace, "radialGradient");
    gradient.setAttribute("cx", String(cx)); gradient.setAttribute("cy", String(cy)); gradient.setAttribute("fx", String(cx)); gradient.setAttribute("fy", String(cy));
    gradient.setAttribute("r", String(Math.hypot(Math.max(cx, 1 - cx), Math.max(cy, 1 - cy))));
  }
  gradient.id = id;
  const stops = parts.map((part, index) => {
    const stopMatch = /^(.*?)(?:\s+([\d.]+)%)?$/.exec(part), rawColor = stopMatch?.[1]?.trim();
    if (!rawColor) return undefined;
    const mixed = /^color-mix\(\s*in\s+srgb\s*,\s*(#[\da-f]{6})\s+([\d.]+)%\s*,\s*transparent\s*\)$/i.exec(rawColor);
    return { index, offset: Math.max(0, Math.min(100, Number(stopMatch?.[2] ?? 0))), color: mixed?.[1] ?? rawColor, opacity: mixed ? Number(mixed[2]) / 100 : undefined };
  }).filter((stop): stop is NonNullable<typeof stop> => Boolean(stop)).sort((a, b) => a.offset - b.offset || a.index - b.index);
  for (const stopData of stops) {
    const stop = document.createElementNS(namespace, "stop");
    stop.setAttribute("offset", `${stopData.offset}%`);
    stop.setAttribute("stop-color", stopData.color);
    if (stopData.opacity !== undefined) stop.setAttribute("stop-opacity", String(stopData.opacity));
    gradient.append(stop);
  }
  if (!gradient.childElementCount) return false;
  const defs = document.createElementNS(namespace, "defs"); defs.append(gradient); svg.append(defs);
  return true;
}

const svgNamespace = "http://www.w3.org/2000/svg";
const pptDashPatterns: Record<string, number[]> = { dash: [4, 3], sysDash: [3, 1], lgDash: [8, 3], dashDot: [4, 3, 1, 3], sysDashDot: [3, 1, 1, 1], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDotDot: [3, 1, 1, 1, 1, 1], dot: [1, 3], sysDot: [1, 1] };
const pptUniqueId = (prefix: string) => `ufp-pptx-${prefix}-${Math.random().toString(36).slice(2)}`;

/** Scales absolute M/L/C/Q/A/Z path data, as produced for DrawingML custom geometry, into the shape box. */
function scalePptPathData(d: string, scaleX: number, scaleY: number): string {
  const tokens = d.match(/[MLCQAZ]|-?\d*\.?\d+(?:e[-+]?\d+)?/gi) ?? [];
  const output: string[] = []; let index = 0, command = "";
  for (const token of tokens) {
    if (/^[A-Z]$/i.test(token)) { command = token.toUpperCase(); index = 0; output.push(command); continue; }
    const value = Number(token), slot = index % 7;
    if (command === "A") output.push(String(slot === 0 || slot === 5 ? value * scaleX : slot === 1 || slot === 6 ? value * scaleY : value));
    else output.push(String(index % 2 === 0 ? value * scaleX : value * scaleY));
    index++;
  }
  return output.join(" ");
}

function appendPptDuotoneFilter(defs: SVGElement, duotone: NonNullable<PptxImageElement["duotone"]>): string {
  const filterId = pptUniqueId("duotone"), filter = document.createElementNS(svgNamespace, "filter");
  filter.id = filterId; filter.setAttribute("color-interpolation-filters", "sRGB");
  const grayscale = document.createElementNS(svgNamespace, "feColorMatrix"); grayscale.setAttribute("type", "saturate"); grayscale.setAttribute("values", "0"); grayscale.setAttribute("result", "ufp-gray"); filter.append(grayscale);
  const transfer = document.createElementNS(svgNamespace, "feComponentTransfer"); transfer.setAttribute("in", "ufp-gray");
  const dark = pptColorChannels(duotone.darkColor), light = pptColorChannels(duotone.lightColor);
  for (const [channel, index] of [["R", 0], ["G", 1], ["B", 2]] as const) {
    const fn = document.createElementNS(svgNamespace, `feFunc${channel}`);
    fn.setAttribute("type", "table"); fn.setAttribute("tableValues", `${(dark[index] / 255).toFixed(4)} ${(light[index] / 255).toFixed(4)}`); transfer.append(fn);
  }
  const alpha = document.createElementNS(svgNamespace, "feFuncA"); alpha.setAttribute("type", "identity"); transfer.append(alpha); filter.append(transfer); defs.append(filter);
  return filterId;
}

function appendPptLineEnd(defs: SVGElement, end: NonNullable<PptxLine["headEnd"]>, color: string, opacity: number | undefined): string | undefined {
  const scale = { sm: 2, med: 3, lg: 5 } as const, marker = document.createElementNS(svgNamespace, "marker"), id = pptUniqueId("end");
  marker.id = id; marker.setAttribute("viewBox", "0 0 10 10"); marker.setAttribute("markerUnits", "strokeWidth"); marker.setAttribute("preserveAspectRatio", "none");
  marker.setAttribute("markerWidth", String(scale[end.length ?? "med"])); marker.setAttribute("markerHeight", String(scale[end.width ?? "med"]));
  marker.setAttribute("orient", "auto-start-reverse"); marker.setAttribute("refY", "5");
  const shape = document.createElementNS(svgNamespace, end.type === "oval" ? "ellipse" : "path");
  if (end.type === "oval") { shape.setAttribute("cx", "5"); shape.setAttribute("cy", "5"); shape.setAttribute("rx", "5"); shape.setAttribute("ry", "5"); marker.setAttribute("refX", "5"); }
  else if (end.type === "diamond") { shape.setAttribute("d", "M 0 5 L 5 0 L 10 5 L 5 10 Z"); marker.setAttribute("refX", "5"); }
  else if (end.type === "stealth") { shape.setAttribute("d", "M 0 0 L 10 5 L 0 10 L 3 5 Z"); marker.setAttribute("refX", "8"); }
  else if (end.type === "triangle") { shape.setAttribute("d", "M 0 0 L 10 5 L 0 10 Z"); marker.setAttribute("refX", "8"); }
  else if (end.type === "arrow") { shape.setAttribute("d", "M 0 0 L 10 5 L 0 10"); shape.setAttribute("stroke", color); shape.setAttribute("stroke-width", "1.5"); marker.setAttribute("refX", "9"); }
  else return undefined;
  shape.setAttribute("fill", end.type === "arrow" ? "none" : color);
  if (opacity !== undefined && opacity < 1) shape.setAttribute("opacity", String(opacity));
  marker.append(shape); defs.append(marker);
  return id;
}

/** Draws shapes and lines as SVG in shape-local EMU so geometry, outlines, flips, rotation and shadows match PowerPoint. */
function renderPptxGeometry(element: PptxShapeElement | PptxLineElement, slideData: Presentation["slides"][number], imageUrls: string[]): HTMLElement {
  const slideWidth = Math.max(1, slideData.width ?? 9_144_000), slideHeight = Math.max(1, slideData.height ?? 5_143_500);
  const node = document.createElement("div"); node.className = element.kind === "line" ? "ufp-pptx-line" : "ufp-pptx-shape";
  node.style.position = "absolute"; node.style.overflow = "visible"; node.style.pointerEvents = "none";
  node.style.left = `${element.left / slideWidth * 100}%`; node.style.top = `${element.top / slideHeight * 100}%`;
  node.style.width = `${element.width / slideWidth * 100}%`; node.style.height = `${element.height / slideHeight * 100}%`;
  const w = Math.max(0, element.width), h = Math.max(0, element.height), line = element.line;
  // Zero-width hairlines still render about one device pixel wide in PowerPoint.
  const strokeWidth = line ? Math.max(line.width, slideWidth / 2400) : 0;
  const pad = strokeWidth * 3 + 12_700, toCqw = (value: number) => `${value / slideWidth * 100}cqw`;
  const svg = document.createElementNS(svgNamespace, "svg");
  svg.setAttribute("aria-hidden", "true"); svg.setAttribute("focusable", "false"); svg.setAttribute("preserveAspectRatio", "none");
  svg.setAttribute("viewBox", `${-pad} ${-pad} ${w + pad * 2} ${h + pad * 2}`);
  svg.style.position = "absolute"; svg.style.left = svg.style.top = `-${toCqw(pad)}`; svg.style.overflow = "visible";
  svg.style.width = `calc(100% + ${toCqw(pad * 2)})`; svg.style.height = `calc(100% + ${toCqw(pad * 2)})`;
  const defs = document.createElementNS(svgNamespace, "defs"); svg.append(defs);

  const paths = element.kind === "line" ? [{ d: `M 0 0 L ${w} ${h}`, filled: false, stroked: true }]
    : element.customPaths?.length ? element.customPaths.map(path => ({ d: scalePptPathData(path.d, path.width > 0 ? w / path.width : 1, path.height > 0 ? h / path.height : 1), filled: path.fill !== false, stroked: path.stroke !== false }))
      : (presetGeometryPath(element.geometry, w, h, element.adjustments) ?? presetGeometryPath("rect", w, h) ?? []).map(path => ({ ...path, stroked: true }));
  // A 3D camera turns the face in its own frame, so the 2D rotation moves to CSS around the projected face.
  const camera = element.kind === "shape" ? pptCameraTransform(element.camera, slideWidth) : "";
  if (camera) node.style.transform = [element.rotation ? `rotate(${element.rotation}deg)` : "", camera].filter(Boolean).join(" ");
  const transforms = [
    element.rotation && !camera ? `rotate(${element.rotation} ${w / 2} ${h / 2})` : "",
    element.flipH || element.flipV ? `translate(${element.flipH ? w : 0} ${element.flipV ? h : 0}) scale(${element.flipH ? -1 : 1} ${element.flipV ? -1 : 1})` : "",
  ].filter(Boolean).join(" ");

  // The shadow filter sits outside the rotation so its offset stays in slide space, as with rotWithShape="0".
  const outer = document.createElementNS(svgNamespace, "g"), shapeGroup = document.createElementNS(svgNamespace, "g");
  if (transforms) shapeGroup.setAttribute("transform", transforms);
  const shadow = element.shadow;
  if (shadow && w > 0 && h > 0) {
    const filter = document.createElementNS(svgNamespace, "filter"), drop = document.createElementNS(svgNamespace, "feDropShadow"), id = pptUniqueId("shadow");
    filter.id = id; filter.setAttribute("x", "-50%"); filter.setAttribute("y", "-50%"); filter.setAttribute("width", "200%"); filter.setAttribute("height", "200%");
    drop.setAttribute("dx", String(shadow.x)); drop.setAttribute("dy", String(shadow.y)); drop.setAttribute("stdDeviation", String(shadow.blur / 2));
    drop.setAttribute("flood-color", shadow.color); drop.setAttribute("flood-opacity", String(shadow.opacity ?? 1));
    filter.append(drop); defs.append(filter); outer.setAttribute("filter", `url(#${id})`);
  }

  // Shapes with useBgFill show the part of the slide background picture that lies beneath them.
  const backgroundImage = element.kind === "shape" && element.useBackgroundFill ? slideData.elements?.find((candidate): candidate is PptxImageElement => candidate.kind === "image" && Boolean(candidate.isBackground)) : undefined;
  if (backgroundImage) {
    const clip = document.createElementNS(svgNamespace, "clipPath"), clipId = pptUniqueId("clip");
    clip.id = clipId;
    for (const pathData of paths.filter(path => path.filled)) {
      const clipShape = document.createElementNS(svgNamespace, "path"); clipShape.setAttribute("d", pathData.d);
      if (transforms) clipShape.setAttribute("transform", transforms);
      clip.append(clipShape);
    }
    defs.append(clip);
    const bytes = new Uint8Array(backgroundImage.image.bytes.byteLength); bytes.set(backgroundImage.image.bytes);
    const url = URL.createObjectURL(new Blob([bytes.buffer], { type: backgroundImage.image.mimeType })); imageUrls.push(url);
    const image = document.createElementNS(svgNamespace, "image");
    image.setAttribute("href", url); image.setAttribute("preserveAspectRatio", "none");
    image.setAttribute("x", String(-element.left)); image.setAttribute("y", String(-element.top));
    image.setAttribute("width", String(slideWidth)); image.setAttribute("height", String(slideHeight));
    if (backgroundImage.duotone) image.setAttribute("filter", `url(#${appendPptDuotoneFilter(defs, backgroundImage.duotone)})`);
    const clipped = document.createElementNS(svgNamespace, "g"); clipped.setAttribute("clip-path", `url(#${clipId})`); clipped.append(image);
    outer.append(clipped);
  }

  const fill = element.kind === "shape" && !backgroundImage ? element.fill : undefined;
  const light = element.kind === "shape" ? pptSceneLight(element) : undefined;
  let fillPaint = "none", fillOpacity: number | undefined;
  if (fill && "gradient" in fill) {
    const id = pptUniqueId("fill"), gradient = light ? fill.gradient.replace(/#[\da-f]{6}\b/gi, color => pptLighten(color, (light.start + light.end) / 2)) : fill.gradient;
    if (appendPptSvgGradient(svg, id, gradient)) fillPaint = `url(#${id})`;
  } else if (fill) {
    fillPaint = fill.color; fillOpacity = fill.opacity;
    const id = pptUniqueId("light");
    if (light && appendPptSvgGradient(svg, id, `linear-gradient(90deg, ${pptLighten(fill.color, light.start)} 0%, ${pptLighten(fill.color, light.end)} 100%)`)) fillPaint = `url(#${id})`;
  }
  const startMarker = line?.headEnd ? appendPptLineEnd(defs, line.headEnd, line.color, line.opacity) : undefined;
  const endMarker = line?.tailEnd ? appendPptLineEnd(defs, line.tailEnd, line.color, line.opacity) : undefined;
  for (const pathData of paths) {
    const path = document.createElementNS(svgNamespace, "path");
    path.setAttribute("d", pathData.d); path.setAttribute("fill-rule", "evenodd");
    path.setAttribute("fill", pathData.filled ? fillPaint : "none");
    if (pathData.filled && fillOpacity !== undefined && fillOpacity < 1) path.setAttribute("fill-opacity", String(fillOpacity));
    if (line && pathData.stroked) {
      path.setAttribute("stroke", line.color); path.setAttribute("stroke-width", String(strokeWidth)); path.setAttribute("stroke-linejoin", "round");
      path.setAttribute("stroke-linecap", line.cap === "round" ? "round" : line.cap === "square" ? "square" : "butt");
      if (line.opacity !== undefined && line.opacity < 1) path.setAttribute("stroke-opacity", String(line.opacity));
      const dash = pptDashPatterns[line.style ?? "solid"];
      if (dash) path.setAttribute("stroke-dasharray", dash.map(value => value * strokeWidth).join(" "));
      if (!pathData.filled) {
        if (startMarker) path.setAttribute("marker-start", `url(#${startMarker})`);
        if (endMarker) path.setAttribute("marker-end", `url(#${endMarker})`);
      }
    }
    shapeGroup.append(path);
  }
  outer.append(shapeGroup); svg.append(outer); node.append(svg);
  return node;
}

/** Clips a picture frame to its preset geometry (ellipse, rounded rectangle, ...). */
function clipPptxImageToGeometry(frame: HTMLElement, element: PptxImageElement): void {
  const w = Math.max(1, element.width), h = Math.max(1, element.height), paths = presetGeometryPath(element.geometry, w, h, element.adjustments);
  if (!paths) return;
  const svg = document.createElementNS(svgNamespace, "svg"), clip = document.createElementNS(svgNamespace, "clipPath"), id = pptUniqueId("picture-clip");
  svg.setAttribute("aria-hidden", "true"); svg.style.position = "absolute"; svg.style.width = "0"; svg.style.height = "0";
  clip.id = id; clip.setAttribute("clipPathUnits", "objectBoundingBox");
  for (const pathData of paths) {
    const path = document.createElementNS(svgNamespace, "path"); path.setAttribute("d", pathData.d); path.setAttribute("transform", `scale(${1 / w} ${1 / h})`); clip.append(path);
  }
  svg.append(clip); frame.append(svg); frame.style.clipPath = `url(#${id})`;
}

function renderPptxParagraphs(container: HTMLElement, paragraphs: Paragraph[], slideWidth: number, alignment?: string, options: DrawingRenderOptions = {}): void {
  let end = paragraphs.length;
  while (end > 0 && !paragraphs[end - 1]!.list && !paragraphs[end - 1]!.runs.some(run => run.text.length > 0)) end--;
  for (let index = 0; index < end; index++) {
    const node = renderPptxParagraph(paragraphs[index]!, slideWidth, alignment);
    const family = options.runFontLines ? paragraphs[index]!.runs.find(run => run.text.length > 0)?.fontFamily ?? paragraphs[index]!.defaultRun?.fontFamily : undefined;
    if (family) node.style.fontFamily = fontStack(family);
    // PowerPoint ignores space before the first paragraph of a text body.
    if (index === 0) node.style.marginTop = "0";
    container.append(node);
  }
}

// Some fonts are unavailable in the browser, so slightly shrink a single word if fallback metrics would wrap it into a clipped line.
function fitPptxSingleTokenRuns(textBox: HTMLElement, element: PptxTextElement, slideWidth: number): void {
  const canvas = document.createElement("canvas"), context = canvas.getContext("2d");
  if (!context) return;
  let end = element.paragraphs.length;
  while (end > 0 && !element.paragraphs[end - 1]!.list && !element.paragraphs[end - 1]!.runs.some(run => run.text.length > 0)) end--;
  for (let index = 0; index < end; index++) {
    const paragraph = element.paragraphs[index]!, run = paragraph.runs.length === 1 ? paragraph.runs[0] : undefined;
    if (!run?.text || /\s/.test(run.text) || paragraph.list) continue;
    const fontSize = run.fontSize ?? paragraph.defaultRun?.fontSize;
    if (!fontSize) continue;
    const fontSizeCqw = fontSize / (slideWidth / 12700) * 100;
    context.font = `${run.italic ? "italic " : ""}${run.bold ? "700 " : "400 "}${fontSizeCqw}px ${fontStack(run.fontFamily ?? paragraph.defaultRun?.fontFamily ?? "Arial")}`;
    const textWidth = context.measureText(run.text).width;
    const availableWidth = (element.width - element.margins.left - element.margins.right - (paragraph.indentLeft ?? 0) - (paragraph.indentRight ?? 0)) / slideWidth * 100;
    const scale = textWidth > availableWidth ? Math.max(0.75, availableWidth / textWidth * 0.995) : 1;
    if (scale === 1) continue;
    const node = textBox.children[index] as HTMLParagraphElement | undefined;
    if (!node) continue;
    node.style.fontSize = `${fontSizeCqw * scale}cqw`;
    const span = node.querySelector<HTMLElement>("span:not(.ufp-pptx-bullet)");
    if (span) span.style.fontSize = `${fontSizeCqw * scale}cqw`;
  }
}

function renderPptxParagraph(paragraph: Paragraph, slideWidth: number, alignment?: string): HTMLParagraphElement {
  const node = document.createElement("p"); node.className = "ufp-pptx-paragraph";
  node.style.margin = `${(paragraph.spacingBefore ?? 0) / (slideWidth / 12700) * 100}cqw 0 ${(paragraph.spacingAfter ?? 0) / (slideWidth / 12700) * 100}cqw`;
  node.style.textAlign = alignment ?? paragraph.alignment ?? "left";
  const firstRunFontSize = paragraph.runs.find(run => run.text.length > 0)?.fontSize ?? paragraph.defaultRun?.fontSize;
  if (firstRunFontSize) node.style.fontSize = `${firstRunFontSize / (slideWidth / 12700) * 100}cqw`;
  if (paragraph.indentLeft) node.style.marginLeft = `${paragraph.indentLeft / (slideWidth / 12700) * 100}cqw`;
  if (paragraph.indentRight) node.style.marginRight = `${paragraph.indentRight / (slideWidth / 12700) * 100}cqw`;
  // PowerPoint line spacing percentages are relative to single spacing (about 1.2em), not to the font size.
  if (paragraph.lineHeight) node.style.lineHeight = paragraph.lineHeight < 4 ? String(paragraph.lineHeight * 1.2) :`${paragraph.lineHeight / (slideWidth / 12700) * 100}cqw`;
  // PowerPoint shows no bullet on an empty paragraph, and wraps bulleted text at the text indent (hanging indent).
  if (paragraph.list && paragraph.runs.some(run => run.text.length > 0 && run.text !== "\n")) {
    const label = document.createElement("span"); label.className = "ufp-pptx-bullet"; label.textContent = `${paragraph.list.label} `;
    const hanging = paragraph.list.hanging ?? 0;
    if (hanging > 0) {
      node.style.textIndent = `${-hanging / (slideWidth / 12700) * 100}cqw`;
      label.style.display = "inline-block"; label.style.minWidth = `${hanging / (slideWidth / 12700) * 100}cqw`; label.style.textIndent = "0";
    }
    node.append(label);
  } else if (paragraph.indentFirstLine && !paragraph.list) node.style.textIndent = `${paragraph.indentFirstLine / (slideWidth / 12700) * 100}cqw`;
  for (const run of paragraph.runs) {
    if (run.text === "\n") { node.append(document.createElement("br")); continue; }
    const span = document.createElement("span"); span.textContent = run.text;
    if (run.fontSize) span.style.fontSize = `${run.fontSize / (slideWidth / 12700) * 100}cqw`;
    if (run.fontFamily) span.style.fontFamily = fontStack(run.fontFamily);
    if (run.bold !== undefined) span.style.fontWeight = run.bold ? "700" : "400";
    if (run.italic !== undefined) span.style.fontStyle = run.italic ? "italic" : "normal";
    if (run.underline || run.strike) span.style.textDecoration = [run.underline ? "underline" : "", run.strike ? "line-through" : ""].filter(Boolean).join(" ");
    if (run.caps) span.style.textTransform = "uppercase";
    if (run.smallCaps) span.style.fontVariant = "small-caps";
    if (run.color) span.style.color = run.color;
    if (run.letterSpacing) span.style.letterSpacing = `${run.letterSpacing / (slideWidth / 12700) * 100}cqw`;
    if (run.shadow) {
      const points = (value: number) => `${value / (slideWidth / 12700) * 100}cqw`, { x, y, blur, color, opacity } = run.shadow;
      span.style.textShadow = `${points(x)} ${points(y)} ${points(blur)} ${pptFillCss(opacity !== undefined ? { color, opacity } : { color })}`;
    }
    node.append(span);
  }
  return node;
}

export function pptFillCss(fill: PptxFill): string {
  if ("gradient" in fill) return fill.gradient;
  if (fill.opacity === undefined || fill.opacity >= 1) return fill.color;
  const match = /^#([\da-f]{2})([\da-f]{2})([\da-f]{2})$/i.exec(fill.color);
  return match ? `rgba(${Number.parseInt(match[1]!, 16)},${Number.parseInt(match[2]!, 16)},${Number.parseInt(match[3]!, 16)},${Math.max(0, Math.min(1, fill.opacity))})` : fill.color;
}

function pptColorChannels(color: string): [number, number, number] {
  const hex = color.replace(/^#/, "");
  return [Number.parseInt(hex.slice(0, 2), 16) || 0, Number.parseInt(hex.slice(2, 4), 16) || 0, Number.parseInt(hex.slice(4, 6), 16) || 0];
}
function pptBorderCss(line: PptxLine, slideWidth: number): string {
  const dash: Record<string, string> = { dash: "dashed", dashDot: "dashed", lgDash: "dashed", sysDash: "dashed", dot: "dotted", sysDot: "dotted", solid: "solid" };
  const color = line.opacity !== undefined && line.opacity < 1 ? pptFillCss({ color: line.color, opacity: line.opacity }) : line.color;
  return `${Math.max(0.025, line.width / slideWidth * 100)}cqw ${dash[line.style ?? "solid"] ?? "solid"} ${color}`;
}
