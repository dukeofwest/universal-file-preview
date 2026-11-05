import { PreviewError } from "../errors.js";
import type { FileFormat } from "../detection/detect-format.js";
import { createImageUrl } from "./image-url.js";
import type { FileRenderer, RendererContext, RendererInstance } from "./types.js";

export class ImageRenderer implements FileRenderer {
  readonly formats: readonly FileFormat[] = ["gif", "png", "jpeg", "webp", "svg"];
  render({ bytes, format, file }: RendererContext): RendererInstance {
    const root = document.createElement("div"); root.className = "ufp-image";
    const mime = format === "jpeg" ? "image/jpeg" : format === "svg" ? "image/svg+xml" : `image/${format}`;
    // SVG is sanitized and raster bytes are checked against their signature (image-url.ts).
    const url = createImageUrl(bytes, mime);
    if (!url && format === "svg") throw new PreviewError("Invalid SVG image.", "invalid-file");
    const image = document.createElement("img");
    const showError = () => {
      const error = document.createElement("div"); error.setAttribute("role", "alert"); error.textContent = "The browser could not decode this image.";
      root.replaceChildren(error);
    };
    image.alt = file.name; image.decoding = "async";
    if (url) { image.src = url; image.onerror = showError; root.append(image); } else showError();
    return {
      element: root,
      setZoom(zoom) { image.style.width = `${Math.max(0.1, Math.min(4, zoom)) * 100}%`; },
      fitToWidth() { image.style.width = "100%"; },
      destroy() { image.onerror = null; image.removeAttribute("src"); if (url) URL.revokeObjectURL(url); },
    };
  }
}
