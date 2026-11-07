import { CommonModule } from "@angular/common";
import { Component, ElementRef, OnDestroy, ViewChild, ViewEncapsulation } from "@angular/core";
import { FilePreviewer } from "@dukeofwest/universal-file-preview";

@Component({
  selector: "ufp-root",
  standalone: true,
  imports: [CommonModule],
  templateUrl: "./app.component.html",
  styleUrl: "./app.component.css",
  encapsulation: ViewEncapsulation.None,
})
export class AppComponent implements OnDestroy {
  @ViewChild("previewHost", { static: true }) private previewHost!: ElementRef<HTMLDivElement>;

  previewer: FilePreviewer | null = null;
  fileName = "";
  busy = false;
  error = "";
  dragging = false;

  async previewFile(file: File): Promise<void> {
    const container = this.previewHost.nativeElement;
    this.previewer?.destroy();
    this.previewer = null;
    container.replaceChildren();
    this.fileName = file.name;
    this.error = "";
    this.busy = true;
    const instance = new FilePreviewer({
      container,
      file: { name: file.name, source: file, mimeType: file.type },
      // Parse in a Web Worker so the page stays responsive; fall back to the main thread if it cannot start.
      worker: { create: () => new Worker(new URL("./parser.worker", import.meta.url), { type: "module" }), fallback: "main-thread" },
    });
    this.previewer = instance;
    try {
      await instance.open();
    } catch (reason: unknown) {
      if (this.previewer === instance) this.error = this.errorText(reason);
    } finally {
      if (this.previewer === instance) this.busy = false;
    }
  }

  onFileChange(event: Event): void {
    const target = event.currentTarget;
    if (!(target instanceof HTMLInputElement)) return;
    const file = target.files?.[0];
    if (file) void this.previewFile(file);
    target.value = "";
  }

  onDrop(event: DragEvent): void {
    event.preventDefault();
    this.dragging = false;
    const file = event.dataTransfer?.files[0];
    if (file) void this.previewFile(file);
  }

  onDragEnter(event: DragEvent): void {
    event.preventDefault();
    this.dragging = true;
  }

  onDragOver(event: DragEvent): void {
    event.preventDefault();
  }

  onDragLeave(event: DragEvent): void {
    const related = event.relatedTarget;
    const current = event.currentTarget;
    if (!(related instanceof Node) || !(current instanceof Node) || !current.contains(related)) this.dragging = false;
  }

  clearFile(): void {
    this.previewer?.destroy();
    this.previewer = null;
    this.previewHost.nativeElement.replaceChildren();
    this.fileName = "";
    this.error = "";
    this.busy = false;
  }

  ngOnDestroy(): void {
    this.previewer?.destroy();
    this.previewer = null;
  }

  private errorText(reason: unknown): string {
    return reason instanceof Error ? reason.message : "The selected file could not be previewed.";
  }
}
