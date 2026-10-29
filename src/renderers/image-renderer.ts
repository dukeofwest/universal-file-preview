import { PreviewError } from "../errors.js";
import type { FileFormat } from "../detection/detect-format.js";
import type { FileRenderer, RendererContext, RendererInstance } from "./types.js";

export class ImageRenderer implements FileRenderer {
  readonly formats: readonly FileFormat[] = ["gif", "png", "jpeg", "webp", "svg"];
  render({ bytes, format, file }: RendererContext): RendererInstance {
    const root = document.createElement("div"); root.className = "ufp-image";
    const mime = format === "jpeg" ? "image/jpeg" : format === "svg" ? "image/svg+xml" : `image/${format}`;
    const content = format === "svg" ? sanitizeSvg(bytes) : bytes;
    const blobBytes = new Uint8Array(content.byteLength); blobBytes.set(content);
    const url = URL.createObjectURL(new Blob([blobBytes.buffer], { type: mime }));
    const image = document.createElement("img");
    image.alt = file.name; image.decoding = "async"; image.src = url;
    image.onerror = () => {
      const error = document.createElement("div"); error.setAttribute("role", "alert"); error.textContent = "The browser could not decode this image.";
      root.replaceChildren(error);
    };
    root.append(image);
    return {
      element: root,
      setZoom(zoom) { image.style.width = `${Math.max(0.1, Math.min(4, zoom)) * 100}%`; },
      fitToWidth() { image.style.width = "100%"; },
      destroy() { image.onerror = null; image.removeAttribute("src"); URL.revokeObjectURL(url); },
    };
  }
}

const svgElements = new Set(["svg", "g", "path", "rect", "circle", "ellipse", "line", "polyline", "polygon", "text", "tspan", "defs", "linearGradient", "radialGradient", "stop", "clipPath", "mask", "title", "desc"]);
const svgAttributes = new Set(["xmlns", "viewBox", "width", "height", "x", "y", "x1", "x2", "y1", "y2", "cx", "cy", "r", "rx", "ry", "d", "points", "fill", "fill-rule", "fill-opacity", "stroke", "stroke-width", "stroke-linecap", "stroke-linejoin", "stroke-dasharray", "stroke-opacity", "opacity", "transform", "font-family", "font-size", "font-weight", "text-anchor", "stop-color", "stop-opacity", "offset", "clip-path", "mask", "preserveAspectRatio", "id"]);
function sanitizeSvg(bytes: Uint8Array): Uint8Array {
  const source = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  const parsed = new DOMParser().parseFromString(source, "image/svg+xml");
  const root = parsed.documentElement;
  if (root.localName !== "svg" || parsed.querySelector("parsererror")) throw new PreviewError("Invalid SVG image.", "invalid-file");
  const clean = (node: Element) => {
    for (const child of [...node.children]) {
      if (!svgElements.has(child.localName)) { child.remove(); continue; }
      clean(child);
    }
    for (const attribute of [...node.attributes]) {
      if (!svgAttributes.has(attribute.name) || /^on/i.test(attribute.name)) { node.removeAttribute(attribute.name); continue; }
      if (["clip-path", "mask"].includes(attribute.name) && !/^url\(#[A-Za-z0-9_.:-]+\)$/.test(attribute.value)) node.removeAttribute(attribute.name);
      if (/url\(|javascript:|data:|https?:/i.test(attribute.value) && !["clip-path", "mask"].includes(attribute.name)) node.removeAttribute(attribute.name);
    }
  };
  clean(root);
  return new TextEncoder().encode(new XMLSerializer().serializeToString(root));
}
