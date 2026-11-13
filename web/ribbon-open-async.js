// ribbon.openAsync(bytes, name, { onProgress, windowSize }): the ribbon's open() (docxodus dist/ribbon.js, 16.0.0)
// with DocxEditor.openAsync in place of DocxEditor.open, so a long document mounts window by window and the page
// keeps painting (the loader can show progress) instead of freezing for the whole render.
// Docxodus's ribbon has no such method yet: this is the shape proposed for it (RibbonEditor.openAsync); until then
// it is added to the instance. It copies open()'s body, so check it against ribbon.js when the package changes.
import { DocxEditor, TrackedChangeMode } from "./node_modules/docxodus/dist/embed.bundle.js";

export function addOpenAsync(ribbon) {
  ribbon.openAsync = async function (bytes, name, { onProgress, windowSize } = {}) {
    if (!this.exports)
      throw new Error("Docxodus ribbon: WASM exports are not set yet");
    this.history?.beforeOpen();
    const paginated = this.require("paginated").checked;
    const candidateSurface = this.surface.cloneNode(false);
    candidateSurface.dataset.view = paginated ? "paginated" : "continuous";
    const started = performance.now();
    const tracked = this.require("trackchanges").checked
      ? TrackedChangeMode.RenderInline
      : (this.options.trackedChanges ?? TrackedChangeMode.Accept);
    const candidate = await DocxEditor.openAsync(candidateSurface, bytes, this.exports, {
      cssPrefix: this.options.cssPrefix,
      fabricateClasses: this.options.fabricateClasses,
      editable: this.options.editable,
      scale: this.options.scale,
      columnWidth: this.options.columnWidth,
      fitToWidth: this.options.fitToWidth,
      onEdit: (info) => {
        this.history?.edited();
        this.options.onEdit?.(info);
        this.scheduleStats();
      },
      onMove: (info) => { this.history?.edited(); this.options.onMove?.(info); },
      onStoryChange: (which) => this.onStoryChange(which),
      onCommentsChange: (info) => this.onCommentsChange(info),
      paginated,
      headerFooter: this.headerFooter,
      blockDrag: this.options.blockDrag ?? true,
      trackedChanges: tracked,
      showRevisions: this.showRevisions,
      revisionAuthor: this.author,
      revisionIdentities: this.options.revisionIdentities,
      comments: this.options.comments,
      commentAuthor: this.author,
      windowSize,
      onProgress,
    });
    this.live?.close();
    candidate.adoptContainer(this.surface);
    this.surface.dataset.view = paginated ? "paginated" : "continuous";
    this.live = candidate;
    if (name)
      this.documentName = name;
    this.require("docname").textContent = this.documentName;
    this.closeFindBar();
    this.history?.documentOpened(this.documentName);
    const saveButton = this.control("save");
    if (saveButton)
      saveButton.disabled = false;
    this.require("ribbon").setAttribute("aria-disabled", "false");
    this.setState("ready");
    this.setStatus(`Rendered in ${Math.round(performance.now() - started)} ms`);
    this.populateStyles();
    this.populateFonts(this.documentFonts());
    this.syncPageNumbering();
    this.syncLayoutState();
    this.syncReviewState();
    this.syncZoom();
    this.refreshRailCounts();
    this.refreshRailAnchor();
    this.scheduleStats();
    this.options.onOpen?.(this.live);
    return this.live;
  };
}
