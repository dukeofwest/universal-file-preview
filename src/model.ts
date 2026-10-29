export type TextRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontSize?: number;
  fontFamily?: string;
  letterSpacing?: number;
  position?: number;
  caps?: boolean;
  smallCaps?: boolean;
  color?: string;
  highlight?: string;
  underlineStyle?: "solid" | "double" | "dotted" | "dashed" | "wavy";
  underlineColor?: string;
  verticalAlign?: "baseline" | "superscript" | "subscript";
  complexScriptFontSize?: number;
  pageNumber?: boolean;
  drawing?: WordDrawing;
  pageBreak?: boolean;
  image?: { bytes: Uint8Array; mimeType: string; alt?: string; width?: number; height?: number };
  /** Outer text shadow; offsets and blur in points. */
  shadow?: { x: number; y: number; blur: number; color: string; opacity?: number };
};
export type WordDrawingPath = { d: string; transform: string; fill?: string; stroke?: string; strokeWidth?: number };
export type WordDrawing = { left: number; top: number; width: number; height: number; behindText: boolean; paths: WordDrawingPath[] };
export type Paragraph = {
  runs: TextRun[];
  defaultRun?: Partial<TextRun>;
  style?: string;
  alignment?: "left" | "center" | "right" | "justify";
  kind?: "paragraph" | "heading";
  spacingBefore?: number;
  spacingAfter?: number;
  lineHeight?: number;
  lineHeightRule?: "auto" | "exact" | "atLeast";
  background?: string;
  borders?: Partial<Record<"top" | "left" | "bottom" | "right" | "between", TableBorder>>;
  borderSpacing?: Partial<Record<"top" | "left" | "bottom" | "right" | "between", number>>;
  indentLeft?: number;
  indentRight?: number;
  indentFirstLine?: number;
  list?: { label: string; level: number; kind: "ordered" | "bullet"; indentLeft?: number; hanging?: number; tabPosition?: number; labelStyle?: Partial<TextRun> };
  pageBreakBefore?: boolean;
};
export type TableBorder = { color: string; width: number; style: string };
export type TableCellMargins = { top?: number; right?: number; bottom?: number; left?: number };
export type TableCell = { paragraphs: Paragraph[]; width?: number; background?: string; verticalAlign?: "top" | "center" | "bottom"; column?: number; colSpan?: number; rowSpan?: number; borders?: Partial<Record<"top" | "left" | "bottom" | "right" | "insideH" | "insideV", TableBorder>>; margins?: TableCellMargins };
export type Table = {
  rows: Array<{ cells: TableCell[]; height?: number }>;
  width?: number;
  widthPercent?: number;
  alignment?: "left" | "center" | "right";
  borders?: Partial<Record<"top" | "left" | "bottom" | "right" | "insideH" | "insideV", TableBorder>>;
  cellMargins?: TableCellMargins;
};
export type WordBlock = { type: "paragraph"; value: Paragraph } | { type: "table"; value: Table } | { type: "pageBreak" };
export type WordPageLayout = { width: number; height: number; marginTop: number; marginRight: number; marginBottom: number; marginLeft: number; headerDistance?: number; footerDistance?: number };
export type WordDocument = { kind: "word"; blocks: WordBlock[]; title: string; page?: WordPageLayout; defaultRun?: Partial<TextRun>; defaultParagraph?: Partial<Paragraph>; headers?: WordBlock[]; footers?: WordBlock[]; firstPageHeaders?: WordBlock[]; firstPageFooters?: WordBlock[]; differentFirstPageHeaderFooter?: boolean };
export type SpreadsheetFont = { family?: string; size?: number; bold?: boolean; italic?: boolean; underline?: "single" | "double" | "singleAccounting" | "doubleAccounting"; strike?: boolean; color?: string; vertAlign?: "superscript" | "subscript" };
export type SpreadsheetFill =
  | { type: "pattern"; pattern: string; foreground?: string; background?: string }
  | { type: "gradient"; gradient: "linear" | "path"; degree: number; stops: Array<{ position: number; color: string }>; left?: number; right?: number; top?: number; bottom?: number };
export type SpreadsheetBorderEdge = { style: string; color: string };
export type SpreadsheetBorders = Partial<Record<"left" | "right" | "top" | "bottom" | "diagonal", SpreadsheetBorderEdge>> & { diagonalUp?: boolean; diagonalDown?: boolean };
export type SpreadsheetAlignment = { horizontal?: "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed"; vertical?: "top" | "center" | "bottom" | "justify" | "distributed"; wrapText?: boolean; indent?: number; textRotation?: number; shrinkToFit?: boolean };
/** A resolved cell format (cellXfs entry, table style region, or conditional-format overlay). Colors are CSS hex values. */
export type SpreadsheetStyle = { font?: SpreadsheetFont; fill?: SpreadsheetFill; borders?: SpreadsheetBorders; alignment?: SpreadsheetAlignment; numberFormat?: string };
export type SpreadsheetRun = { text: string; font?: SpreadsheetFont };
/** `value` is the raw stored value; `display` is the value formatted with the cell's number format. `repeat` marks a `*x` fill position in `display`. */
export type SpreadsheetCell = { ref: string; value: string; formula?: string; type?: string; style?: number; display?: string; displayColor?: string; repeat?: { index: number; char: string }; runs?: SpreadsheetRun[] };
export type SpreadsheetRange = { top: number; left: number; bottom: number; right: number };
/** Column widths are in Excel character units (as stored); row heights in points. */
export type SpreadsheetColumn = { min: number; max: number; width?: number; hidden?: boolean; style?: number };
export type SpreadsheetRow = { height?: number; hidden?: boolean; style?: number };
export type SpreadsheetConditional = {
  style?: SpreadsheetStyle;
  dataBar?: { color: string; start: number; end: number; gradient: boolean; border?: string; negative?: boolean; axis?: number };
  icon?: { set: string; index: number };
  hideValue?: boolean;
  display?: string;
};
export type SpreadsheetTable = { range: SpreadsheetRange; headerRows: number; totalsRows: number; elements: Partial<Record<string, SpreadsheetStyle & { size?: number }>>; showRowStripes: boolean; showColumnStripes: boolean; showFirstColumn: boolean; showLastColumn: boolean };
export type SpreadsheetAnchor = { from: { column: number; row: number; columnOffset: number; rowOffset: number }; to?: { column: number; row: number; columnOffset: number; rowOffset: number }; width?: number; height?: number; x?: number; y?: number };
/** `frame` is the anchored object's own position and size in EMU; `elements` (shapes, text, pictures of the object and its groups) are positioned in the same space. */
export type SpreadsheetDrawing =
  | { kind: "image"; anchor: SpreadsheetAnchor; bytes: Uint8Array; mimeType: string; name?: string }
  | { kind: "shapes"; anchor: SpreadsheetAnchor; frame: { left: number; top: number; width: number; height: number }; elements: PptxElement[]; name?: string }
  | SpreadsheetSlicer;
/** Slicer style parts as differential formats (font color, fill, borders); `header` is the caption row. */
export type SpreadsheetSlicerStyle = Partial<Record<"wholeTable" | "header" | "selectedItemWithData" | "selectedItemWithNoData" | "unselectedItemWithData" | "unselectedItemWithNoData", SpreadsheetStyle>>;
/** A slicer's buttons in display order; `itemHeight` is in EMU. `filtered` is true when only some items are selected. */
export type SpreadsheetSlicer = { kind: "slicer"; anchor: SpreadsheetAnchor; name: string; caption?: string; columnCount: number; itemHeight: number; items: Array<{ text: string; selected: boolean; hasData: boolean }>; filtered: boolean; style: SpreadsheetSlicerStyle };
/** An in-cell chart: `values` in data order (null for empty cells); `min`/`max` fix the axis for group or custom scaling; colors are CSS hex values and `lineWeight` is in points. */
export type SpreadsheetSparkline = {
  ref: string; type: "line" | "column" | "stacked"; values: Array<number | null>; color: string; lineWeight: number; emptyCells: "gap" | "zero" | "span";
  min?: number; max?: number; negativeColor?: string; axisColor?: string; markersColor?: string; highColor?: string; lowColor?: string; firstColor?: string; lastColor?: string; rightToLeft?: boolean;
};
export type Worksheet = {
  name: string; cells: SpreadsheetCell[]; maxRow: number; maxColumn: number;
  hidden?: boolean; tabColor?: string;
  columns?: SpreadsheetColumn[]; rows?: Record<number, SpreadsheetRow>; defaultColumnWidth?: number; defaultRowHeight?: number;
  merges?: SpreadsheetRange[]; frozen?: { rows: number; columns: number }; topLeftCell?: string;
  showGridLines?: boolean; showHeaders?: boolean; zoom?: number;
  conditional?: Record<string, SpreadsheetConditional>; tables?: SpreadsheetTable[]; filterButtons?: string[];
  hyperlinks?: Record<string, string>; comments?: Record<string, string>; drawings?: SpreadsheetDrawing[]; sparklines?: SpreadsheetSparkline[];
  /** Data validation input messages, shown while a cell of their ranges is selected. */
  inputMessages?: Array<{ ranges: SpreadsheetRange[]; title?: string; text: string }>;
};
export type Workbook = { kind: "workbook"; sheets: Worksheet[]; title: string; styles?: SpreadsheetStyle[]; defaultFont?: SpreadsheetFont; activeSheet?: number; date1904?: boolean };
export type PptxFill = { color: string; opacity?: number } | { gradient: string };
export type PptxLine = { color: string; opacity?: number; width: number; style?: string; cap?: "round" | "square" | "flat"; headEnd?: PptxLineEnd; tailEnd?: PptxLineEnd };
export type PptxLineEnd = { type: string; width?: "sm" | "med" | "lg"; length?: "sm" | "med" | "lg" };
export type PptxShadow = { x: number; y: number; blur: number; color: string; opacity?: number };
/** DrawingML 3D scene camera: rotation angles in degrees, viewing distance in EMU for perspective cameras, and whether a light rig shades the face. */
export type PptxCamera = { latitude: number; longitude: number; revolution: number; distance?: number; lit?: boolean };
/** Mirror image below a box: opacity fades from start to end over `endPosition` (0-1) of the box height; distance and blur in EMU. */
export type PptxReflection = { startOpacity: number; endOpacity: number; endPosition: number; distance: number; blur: number };
export type PptxShapeElement = {
  kind: "shape";
  left: number;
  top: number;
  width: number;
  height: number;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  geometry?: string;
  /** Preset geometry adjust values (avLst), keyed by guide name. */
  adjustments?: Record<string, number>;
  /** Shape is filled with the slide background (p:sp useBgFill="1"). */
  useBackgroundFill?: boolean;
  customPaths?: Array<{ d: string; width: number; height: number; fill?: boolean; stroke?: boolean }>;
  fill?: PptxFill;
  line?: PptxLine;
  shadow?: PptxShadow;
  camera?: PptxCamera;
};
export type PptxLineElement = { kind: "line"; left: number; top: number; width: number; height: number; rotation?: number; flipH?: boolean; flipV?: boolean; line?: PptxLine; shadow?: PptxShadow };
export type PptxTextElement = {
  kind: "text";
  left: number;
  top: number;
  width: number;
  height: number;
  rotation?: number;
  verticalAlign?: "top" | "center" | "bottom";
  wrap?: boolean;
  margins: { top: number; right: number; bottom: number; left: number };
  paragraphs: Paragraph[];
  /** Preset text warp (bodyPr/prstTxWarp), e.g. "textArchUp". */
  warp?: string;
  reflection?: PptxReflection;
  camera?: PptxCamera;
};
export type PptxImageElement = {
  kind: "image";
  left: number;
  top: number;
  width: number;
  height: number;
  rotation?: number;
  image: { bytes: Uint8Array; mimeType: string; alt?: string };
  crop?: { left: number; top: number; right: number; bottom: number };
  flipH?: boolean;
  flipV?: boolean;
  geometry?: string;
  adjustments?: Record<string, number>;
  /** The slide background picture, also used by shapes that fill with the background. */
  isBackground?: boolean;
  duotone?: { darkColor: string; lightColor: string };
};
export type PptxTableElement = {
  kind: "table";
  left: number;
  top: number;
  width: number;
  height: number;
  rotation?: number;
  columns: number[];
  rows: Array<{ height: number; cells: Array<{ paragraphs: Paragraph[]; fill?: PptxFill; verticalAlign?: "top" | "center" | "bottom"; alignment?: "left" | "center" | "right"; margins: { top: number; right: number; bottom: number; left: number }; borders?: Partial<Record<"top" | "right" | "bottom" | "left", PptxLine>> }> }>;
};
export type PptxElement = PptxShapeElement | PptxLineElement | PptxTextElement | PptxImageElement | PptxTableElement;
export type Presentation = { kind: "presentation"; slides: Array<{ title: string; paragraphs: Paragraph[]; width?: number; height?: number; background?: PptxFill; elements?: PptxElement[] }>; title: string };
