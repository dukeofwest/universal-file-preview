/**
 * The viewer's stylesheet and how it reaches the page. Internal module.
 *
 * Where constructable stylesheets exist, one CSSStyleSheet per document or shadow root holds the viewer CSS and
 * is adopted by that root (adoptedStyleSheets); viewers in the same root share it, and it is removed with the
 * last of them. This needs no 'unsafe-inline' in style-src and no markup sink. Elsewhere the viewer adds a
 * <style> element to its container, as before, carrying `styleNonce` when the application provides one; the
 * library cannot reliably detect a fallback stylesheet that a Content Security Policy blocked.
 */

export interface ViewerStyles {
  /** The fallback <style> element, to place in the container; undefined when the stylesheet is adopted. */
  readonly element?: HTMLStyleElement;
  /** Releases this viewer's use of the stylesheet. */
  release(): void;
}

type StyleRoot = (Document | ShadowRoot) & { adoptedStyleSheets: CSSStyleSheet[] };
const shared = new WeakMap<Document | ShadowRoot, { sheet: CSSStyleSheet; users: number }>();

/** The document or shadow root whose styles apply to the container (its owner document while it is detached). */
function styleRoot(container: HTMLElement): Document | ShadowRoot {
  const root = container.getRootNode();
  // nodeType and "host" instead of instanceof: the container may come from another window (an iframe).
  if (root.nodeType === 11 && "host" in root) return root as ShadowRoot;
  return container.ownerDocument;
}

function constructable(root: Document | ShadowRoot): typeof CSSStyleSheet | undefined {
  const view = (root.nodeType === 9 ? root as Document : root.ownerDocument)?.defaultView;
  const Sheet = view?.CSSStyleSheet;
  if (typeof Sheet !== "function" || typeof Sheet.prototype.replaceSync !== "function" || !("adoptedStyleSheets" in root)) return undefined;
  return Sheet;
}

export function attachViewerStyles(container: HTMLElement, nonce: string | undefined): ViewerStyles {
  const root = styleRoot(container), Sheet = constructable(root);
  if (Sheet) {
    try {
      const target = root as StyleRoot;
      let entry = shared.get(root);
      if (!entry) {
        const sheet = new Sheet();
        sheet.replaceSync(VIEWER_CSS);
        entry = { sheet, users: 0 };
        shared.set(root, entry);
      }
      // Adopted again if the application has replaced the root's adoptedStyleSheets in the meantime.
      if (!target.adoptedStyleSheets.includes(entry.sheet)) target.adoptedStyleSheets = [...target.adoptedStyleSheets, entry.sheet];
      entry.users++;
      const used = entry;
      let released = false;
      return {
        release() {
          if (released) return;
          released = true;
          if (--used.users > 0) return;
          shared.delete(root);
          target.adoptedStyleSheets = target.adoptedStyleSheets.filter(sheet => sheet !== used.sheet);
        },
      };
    } catch {
      // A browser that exposes the API but refuses it: fall back to a <style> element.
    }
  }
  const element = container.ownerDocument.createElement("style");
  if (nonce !== undefined) element.setAttribute("nonce", nonce);
  element.textContent = VIEWER_CSS;
  return { element, release: () => element.remove() };
}

// .ufp-root is contained (layout, paint, style; own stacking context): document content, even position:fixed,
// cannot be laid out or painted outside the viewer (defence in depth).
const styles = `
@font-face{font-family:"ufp Franklin Gothic Book";src:local("Segoe UI"),local("SegoeUI");size-adjust:97.6%;ascent-override:93.9%;descent-override:22.3%;line-gap-override:0%}@font-face{font-family:"ufp Franklin Gothic Book";font-style:italic;src:local("Segoe UI Italic"),local("SegoeUI-Italic");size-adjust:100%;ascent-override:91.7%;descent-override:21.7%;line-gap-override:0%}@font-face{font-family:"ufp Franklin Gothic Book Alt";src:local("Arial"),local("ArialMT"),local("Helvetica"),local("Liberation Sans"),local("Arimo");size-adjust:96%;ascent-override:95.5%;descent-override:22.6%;line-gap-override:0%}@font-face{font-family:"ufp Gill Sans MT";src:local("Corbel");size-adjust:99.2%;ascent-override:93.6%;descent-override:23.2%;line-gap-override:0%;font-feature-settings:"lnum" 1,"tnum" 1}@font-face{font-family:"ufp Gill Sans MT";font-weight:bold;src:local("Corbel Bold"),local("Corbel-Bold");size-adjust:110.4%;ascent-override:84.1%;descent-override:20.8%;line-gap-override:0%;font-feature-settings:"lnum" 1,"tnum" 1}@font-face{font-family:"ufp Gill Sans MT";font-style:italic;src:local("Corbel Italic"),local("Corbel-Italic");size-adjust:94.2%;ascent-override:98.6%;descent-override:24.4%;line-gap-override:0%;font-feature-settings:"lnum" 1,"tnum" 1}@font-face{font-family:"ufp Gill Sans MT Alt";src:local("Arial"),local("ArialMT"),local("Helvetica"),local("Liberation Sans"),local("Arimo");size-adjust:91.6%;ascent-override:101.4%;descent-override:25.1%;line-gap-override:0%}@font-face{font-family:"ufp Gill Sans MT Alt";font-weight:bold;src:local("Arial Bold"),local("Arial-BoldMT"),local("Helvetica Bold"),local("Liberation Sans Bold"),local("Arimo Bold");size-adjust:100.5%;ascent-override:92.4%;descent-override:22.9%;line-gap-override:0%}@font-face{font-family:"ufp Gill Sans MT Alt";font-style:italic;src:local("Arial Italic"),local("Arial-ItalicMT"),local("Helvetica Oblique"),local("Liberation Sans Italic"),local("Arimo Italic");size-adjust:83.7%;ascent-override:111.0%;descent-override:27.5%;line-gap-override:0%}
.ufp-root{font:14px/1.45 system-ui,sans-serif;color:#20242a;background:#f3f5f7;border:1px solid #d5d9de;border-radius:8px;display:flex;flex-direction:column;min-width:0;min-height:240px;height:100%;overflow:hidden;contain:layout paint style;isolation:isolate}
.ufp-root *{box-sizing:border-box}.ufp-toolbar{display:flex;align-items:center;gap:8px;padding:8px 10px;background:#fff;border-bottom:1px solid #d5d9de;min-height:46px}.ufp-toolbar[hidden]{display:none}.ufp-file-title{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin-right:auto;font-weight:600}.ufp-button{font:inherit;color:inherit;background:#fff;border:1px solid #c8cdd3;border-radius:5px;padding:5px 9px;cursor:pointer}.ufp-button:hover{background:#edf1f5}.ufp-button:focus-visible,.ufp-sheet-tab:focus-visible{outline:2px solid #2864dc;outline-offset:2px}.ufp-zoom-value{min-width:42px;text-align:center;font-variant-numeric:tabular-nums}.ufp-status{padding:9px 12px;color:#555;background:#fafbfc}.ufp-status-ready{font-size:12px;color:#56606b;border-bottom:1px solid #e1e4e8}.ufp-status-error{color:#8e1c1c;background:#fff1f0}.ufp-content{flex:1;min-height:0;overflow:auto;padding:16px}.ufp-image{min-height:100%;display:flex;justify-content:center;align-items:flex-start}.ufp-image img{display:block;max-width:none;height:auto;object-fit:contain}.ufp-word-document{display:flex;flex-direction:column;align-items:center;gap:18px}.ufp-word-page{background:white;color:#111;max-width:100%;min-height:700px;margin:0 auto;padding:68px 76px;box-shadow:0 2px 10px #0001;overflow-wrap:anywhere}.ufp-word-header,.ufp-word-footer{flex:none}.ufp-word-body{min-height:1em}.ufp-word-page p{margin:0 0 8pt;white-space:pre-wrap}.ufp-word-page table{border-collapse:collapse;table-layout:fixed;width:100%;margin:8pt 0}.ufp-word-page td{padding:5px;vertical-align:top;overflow-wrap:normal;word-break:normal}.ufp-word-page .ufp-word-image{max-width:100%;height:auto;vertical-align:middle}.ufp-word-list-label{white-space:pre}.ufp-sheet-view{height:100%;display:flex;flex-direction:column;min-height:0}.ufp-sheet-notice{flex:none;margin:0 0 8px;padding:8px 12px;border:1px solid #e3c46b;border-radius:6px;background:#fff8e1;color:#4d3800;font-size:13px;line-height:1.4}.ufp-sheet-notice p{margin:0;font-weight:600}.ufp-sheet-notice ul{margin:4px 0 0;padding-left:20px}.ufp-sheet-tabs{display:flex;overflow-x:auto;border-bottom:1px solid #cbd0d6;background:#fff}.ufp-sheet-tab{border:0;border-bottom:2px solid transparent;background:transparent;padding:8px 12px;font:inherit;white-space:nowrap;cursor:pointer}.ufp-sheet-tab[aria-selected=true]{border-bottom-color:#2365c7;color:#174d96}.ufp-grid-viewport{flex:1;min-height:160px;overflow:auto;background:white}.ufp-grid-scroll{position:relative;width:100%;height:100%;overflow:auto;background:#fff}.ufp-grid-canvas{position:relative}.ufp-grid-scroll *{box-sizing:border-box}.ufp-grid-pane-top{position:sticky;top:0;z-index:3;background:#fff}.ufp-grid-pane-corner{position:sticky;left:0;z-index:2;background:#fff}.ufp-grid-pane-left{position:sticky;left:0;z-index:2;background:#fff}.ufp-grid-body{position:absolute;inset:0;z-index:1}.ufp-cell-header{position:absolute;overflow:hidden;white-space:nowrap;background:#f5f5f5;color:#444;text-align:center;border-right:1px solid #d4d4d4;border-bottom:1px solid #d4d4d4}.ufp-cell-corner{background:#f5f5f5}.ufp-grid-cell{position:absolute;display:flex;flex-direction:column;overflow:hidden;pointer-events:none}.ufp-grid-cell a{pointer-events:auto}.ufp-cell-text{min-width:0}.ufp-gridline,.ufp-cell-fill,.ufp-cell-border,.ufp-data-bar,.ufp-cell-icon,.ufp-filter-button,.ufp-cell-comment{position:absolute;pointer-events:none}.ufp-cell-comment{pointer-events:auto}.ufp-grid-selection{position:absolute;border:2px solid #107c41;pointer-events:none;z-index:2}.ufp-input-message{position:absolute;z-index:3;width:200px;padding:3px 6px;background:#ffffe1;border:1px solid #767676;color:#000;font:12px/1.25 "Segoe UI",Tahoma,sans-serif;white-space:pre-wrap;overflow-wrap:break-word;pointer-events:none}.ufp-sheet-drawing{position:absolute;z-index:1;overflow:visible;pointer-events:none;container-type:inline-size}.ufp-sheet-slicer{position:absolute;z-index:1;overflow:hidden;pointer-events:none}.ufp-slicer-caption,.ufp-slicer-item{overflow:hidden;white-space:nowrap;text-overflow:ellipsis}.ufp-presentation{max-width:1100px;margin:0 auto}.ufp-slide-nav{display:flex;justify-content:center;align-items:center;gap:12px;margin:0 0 12px}.ufp-slide-nav button{font:inherit;padding:5px 9px}.ufp-slide{width:min(100%,960px);min-height:540px;margin:auto;padding:54px;background:white;box-shadow:0 2px 12px #0002;transform-origin:top center}.ufp-slide h1{font-size:2em}.ufp-slide p{white-space:pre-wrap}
.ufp-slide{position:relative;width:min(100%,960px);height:auto;min-height:0;padding:0;overflow:hidden;container-type:inline-size;background:#fff;isolation:isolate}.ufp-slide h1{font-size:2em}.ufp-slide>p{white-space:pre-wrap}.ufp-pptx-image,.ufp-pptx-shape,.ufp-pptx-text,.ufp-pptx-table,.ufp-pptx-line{box-sizing:border-box}.ufp-pptx-image img{display:block}.ufp-pptx-text{line-height:normal;color:#000}.ufp-pptx-paragraph{flex:none;min-height:0;white-space:inherit;overflow-wrap:anywhere}.ufp-pptx-table{overflow:hidden}.ufp-pptx-cell{min-width:0;min-height:0;overflow:hidden;box-sizing:border-box;overflow-wrap:anywhere}.ufp-pptx-cell p{margin:0}.ufp-pptx-bullet{white-space:pre}
`;

/** The whole viewer stylesheet (the same text the viewer placed in its <style> element before 0.3.0). */
export const VIEWER_CSS: string = `${styles}
.ufp-word-page{position:relative}.ufp-word-page p,.ufp-word-page h2{margin:0;font-size:inherit;font-weight:inherit;white-space:pre-wrap}.ufp-word-page table{margin:0}.ufp-word-page td{padding:0}
`;
