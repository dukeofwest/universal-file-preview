import { useEffect, useRef, useState, type ChangeEvent, type DragEvent } from "react";
import { FilePreviewer } from "@dukeofwest/universal-file-preview";

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : "The selected file could not be previewed.";
}

export default function App() {
  const host = useRef<HTMLDivElement>(null);
  const previewer = useRef<FilePreviewer | null>(null);
  const [fileName, setFileName] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [dragging, setDragging] = useState(false);

  useEffect(() => () => {
    previewer.current?.destroy();
    previewer.current = null;
  }, []);

  async function previewFile(file: File): Promise<void> {
    const container = host.current;
    if (!container) return;
    previewer.current?.destroy();
    previewer.current = null;
    container.replaceChildren();
    setFileName(file.name);
    setError("");
    setBusy(true);
    const instance = new FilePreviewer({
      container,
      file: { name: file.name, source: file, mimeType: file.type },
    });
    previewer.current = instance;
    try {
      await instance.open();
    } catch (reason: unknown) {
      if (previewer.current === instance) setError(errorText(reason));
    } finally {
      if (previewer.current === instance) setBusy(false);
    }
  }

  function onFileChange(event: ChangeEvent<HTMLInputElement>): void {
    const file = event.currentTarget.files?.[0];
    if (file) void previewFile(file);
    event.currentTarget.value = "";
  }

  function onDrop(event: DragEvent<HTMLElement>): void {
    event.preventDefault();
    setDragging(false);
    const file = event.dataTransfer.files[0];
    if (file) void previewFile(file);
  }

  function clearFile(): void {
    previewer.current?.destroy();
    previewer.current = null;
    host.current?.replaceChildren();
    setFileName("");
    setError("");
    setBusy(false);
  }

  return (
    <main className="example-shell">
      <header className="example-header">
      <a className="back-link" href="../">← All framework examples</a>
        <p className="eyebrow">Framework integration</p>
        <h1 className="typography-eyebrow-super">React file preview</h1>
        <p className="intro">Preview supported images and Office Open XML files in your React app. Files are read locally in your browser and are never uploaded by this demo.</p>
      </header>

      <section
        className={`upload-card${dragging ? " is-dragging" : ""}`}
        onDragEnter={(event) => { event.preventDefault(); setDragging(true); }}
        onDragOver={(event) => event.preventDefault()}
        onDragLeave={(event) => { if (!(event.relatedTarget instanceof Node) || !event.currentTarget.contains(event.relatedTarget)) setDragging(false); }}
        onDrop={onDrop}
      >
        <div className="upload-copy">
          <strong>{fileName || "Choose a file to preview"}</strong>
          <span>{busy ? "Preparing local preview…" : "Drop a file here, or browse your device."}</span>
        </div>
        <div className="upload-actions">
          <label className="button button-primary">
            Choose file
            <input type="file" accept="image/*,.svg,.docx,.xlsx,.pptx" onChange={onFileChange} />
          </label>
          {fileName && <button className="button button-secondary" type="button" onClick={clearFile}>Clear</button>}
        </div>
      </section>

      {error && <p className="error-message" role="alert">{error}</p>}
      <section className="preview-card" aria-label="File preview">
        {busy && <p className="loading-message" role="status">Loading preview…</p>}
        {!fileName && <div className="empty-state"><span className="file-mark" aria-hidden="true">＋</span><strong>Your preview will appear here</strong><span>Try PNG, JPEG, GIF, WebP, SVG, DOCX, XLSX or PPTX.</span></div>}
        <div className="preview-host" ref={host} />
      </section>
      <footer className="example-footer">
        <span>The preview toolbar includes zoom and original-file download controls.</span>
        <a href="https://www.npmjs.com/package/@dukeofwest/universal-file-preview">npm package</a>
      </footer>
    </main>
  );
}
