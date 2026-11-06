import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync, readdirSync } from "node:fs";
import { fontStack } from "../dist/renderers/fonts.js";
import { imageBlobType, sniffRasterType } from "../dist/renderers/image-url.js";
import { parseDocx, parsePptx } from "../dist/parsers/ooxml.js";
import { parseXlsx } from "../dist/parsers/xlsx.js";
import { parseOds } from "../dist/parsers/odf/ods.js";
import { parseFods } from "../dist/parsers/odf/flat.js";
import { securityFixtures } from "./helpers/security-fixtures.mjs";

// Security regressions: embedded SVG, font-name CSS injection and style construction. Rendering needs a browser;
// scripts/browser-security.mjs checks the same fixtures in headless Edge or Chrome.

const ATTACKER = "https://attacker.invalid";
const fixtures = securityFixtures(ATTACKER);
const parse = (name, { kind, bytes }) => {
  const data = new Uint8Array(bytes);
  return kind === "pptx" ? parsePptx(data, name) : kind === "xlsx" ? parseXlsx(data, name) : kind === "docx" ? parseDocx(data, name) : kind === "ods" ? parseOds(data, name) : parseFods(data, name);
};
/** Every { bytes, mimeType } image and every font family string anywhere in a parsed model. */
function collect(model) {
  const images = [], families = [], seen = new Set();
  const visit = (value) => {
    if (!value || typeof value !== "object" || seen.has(value) || ArrayBuffer.isView(value)) return;
    seen.add(value);
    if (value.bytes instanceof Uint8Array && typeof value.mimeType === "string") images.push(value);
    for (const [key, item] of value instanceof Map ? value.entries() : Object.entries(value)) {
      if ((key === "family" || key === "fontFamily") && typeof item === "string") families.push(item);
      visit(item);
    }
  };
  visit(model);
  return { images, families };
}

// ------------------------------------------------------------------------------------------------
// fontStack: font names from documents
// ------------------------------------------------------------------------------------------------

test("fontStack keeps ordinary Office font names and their fallbacks", () => {
  assert.equal(fontStack("Calibri"), '"Calibri", Carlito, sans-serif');
  assert.equal(fontStack("Times New Roman"), '"Times New Roman", "Liberation Serif", Tinos, serif');
  assert.equal(fontStack("  Courier New "), '"Courier New", "Liberation Mono", Cousine, monospace');
  assert.equal(fontStack("Calibri Light"), '"Calibri Light", "Segoe UI Light", "Helvetica Neue Light", sans-serif');
  for (const name of ["MS Gothic", "@MS Mincho", "ＭＳ 明朝", "맑은 고딕", "Noto Sans CJK JP-Regular", "Segoe UI Symbol", "Wingdings 2", "Arial (Body)", "Foo_Bar 1.0", "A&B", "+mn-lt", "Tahoma+", "Çalışkan Sans", "a".repeat(128)]) {
    assert.equal(fontStack(name).startsWith(`"${name}", `), true, name);
  }
  assert.equal(fontStack("Arial   Black"), '"Arial Black", sans-serif', "runs of spaces collapse");
});

test("fontStack rejects names that could leave the CSS string and uses the generic family", () => {
  const unsafe = [
    "Calibri\n;background:url(https://attacker.invalid/x)", "Calibri\r;x:y", "Calibri\f;x:y", "Calibri\t;x:y", "Cal\vibri", "Cal\0ibri",
    "Calibri;color:red", 'Calibri";color:red', "Calibri';color:red", "Calibri\\", "Calibri\\a;x:y", "Calibri}", "Calibri{", "Calibri<", "Calibri>",
    "url(https://attacker.invalid/x)", "Calibri/*", "Calibri:", "Calibri\u0085", "Calibri\u2028x", "Cal\u00a0ibri", "Cal\u200bibri", "Calibri,Arial",
    "", "   ", "a".repeat(129),
  ];
  for (const name of unsafe) assert.equal(fontStack(name), "sans-serif", JSON.stringify(name));
});

test("fontStack output never contains characters that end a declaration or a string", () => {
  const samples = [...Object.values(fixtures).flatMap(fixture => collect(parse("x", fixture)).families), "Calibri", "Cambria", "x\ny", "x;y", 'x"y'];
  for (const family of samples) {
    const stack = fontStack(family);
    assert.doesNotMatch(stack, /[;{}\\<>\u0000-\u001f\u007f-\u009f\u2028\u2029]/u, JSON.stringify(family));
    assert.equal((stack.match(/"/g) ?? []).length % 2, 0, "quotes stay balanced");
    assert.doesNotMatch(stack, /url\(/i);
  }
});

test("malicious font names in XLSX, DOCX, theme typefaces, PPTX themes, ODS and FODS reach the model raw but are refused by fontStack", () => {
  for (const name of ["xlsx-font-names.xlsx", "xlsx-theme-font.xlsx", "docx-font-names.docx", "pptx-theme-font.pptx", "ods-font-names.ods", "fods-font-names.fods"]) {
    const { families } = collect(parse(name, fixtures[name]));
    const malicious = families.filter(family => family.includes(ATTACKER));
    assert.ok(malicious.length > 0, `${name}: the fixture reaches the model`);
    for (const family of malicious) assert.equal(fontStack(family), "sans-serif", `${name}: ${JSON.stringify(family)}`);
  }
  for (const [label, name] of [["XLSX", "xlsx-font-names.xlsx"], ["DOCX", "docx-font-names.docx"]]) {
    const families = collect(parse(name, fixtures[name])).families;
    for (const character of ["\n", "\r", "\f", '"', "\\", "}"]) assert.ok(families.some(family => family.includes(character)), `${label} fixture covers ${JSON.stringify(character)}`);
  }
});

// ------------------------------------------------------------------------------------------------
// Image types and signatures
// ------------------------------------------------------------------------------------------------

const signatures = {
  png: Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13]),
  jpeg: Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 0, 16]),
  gif87: new TextEncoder().encode("GIF87a\x01\x00"), gif89: new TextEncoder().encode("GIF89a\x01\x00"),
  webp: new TextEncoder().encode("RIFF\x00\x00\x00\x00WEBPVP8 "),
  bmp: Uint8Array.from([0x42, 0x4d, ...new Array(24).fill(0)]),
};

test("sniffRasterType recognises PNG, JPEG, GIF, WebP and BMP signatures only", () => {
  assert.equal(sniffRasterType(signatures.png), "image/png");
  assert.equal(sniffRasterType(signatures.jpeg), "image/jpeg");
  assert.equal(sniffRasterType(signatures.gif87), "image/gif");
  assert.equal(sniffRasterType(signatures.gif89), "image/gif");
  assert.equal(sniffRasterType(signatures.webp), "image/webp");
  assert.equal(sniffRasterType(signatures.bmp), "image/bmp");
  for (const bytes of [new Uint8Array(), signatures.png.subarray(0, 7), new TextEncoder().encode("GIF88a"), new TextEncoder().encode("RIFF\0\0\0\0WAVE"), Uint8Array.from([0x42, 0x4d, 0]),
    new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>'), new TextEncoder().encode("<html><script>x</script>"), Uint8Array.from([0x49, 0x49, 42, 0])]) {
    assert.equal(sniffRasterType(bytes), undefined, String(bytes.subarray(0, 8)));
  }
});

test("imageBlobType takes the type from the content, never from the declared type alone", () => {
  const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"><script>x</script></svg>');
  // Correctly labelled raster images keep working.
  for (const [type, bytes] of [["image/png", signatures.png], ["image/jpeg", signatures.jpeg], ["image/gif", signatures.gif89], ["image/webp", signatures.webp], ["image/bmp", signatures.bmp]]) {
    assert.equal(imageBlobType(bytes, type), type);
    assert.equal(imageBlobType(bytes, `${type.toUpperCase()}; charset=x`), type, "case and parameters are ignored");
  }
  // A raster file with the wrong raster extension is shown with its real type.
  assert.equal(imageBlobType(signatures.jpeg, "image/png"), "image/jpeg");
  // SVG is only ever returned for declared SVG, so the caller sanitizes it.
  assert.equal(imageBlobType(svg, "image/svg+xml"), "image/svg+xml");
  assert.equal(imageBlobType(signatures.png, "image/svg+xml"), "image/png");
  for (const type of ["image/png", "image/jpeg", "image/gif", "image/webp", "image/bmp"]) assert.equal(imageBlobType(svg, type), undefined, `SVG labelled ${type}`);
  // Types outside the allowlist are refused even with a valid signature.
  for (const type of ["text/html", "application/xhtml+xml", "text/xml", "application/octet-stream", "image/x-emf", "image/tiff", "", "image/svg"]) {
    assert.equal(imageBlobType(signatures.png, type), undefined, type);
    assert.equal(imageBlobType(svg, type), undefined, type);
  }
});

test("every embedded image in the malicious packages is either refused or sent to the SVG sanitizer", () => {
  const routes = { svg: 0, refused: 0 };
  for (const [name, fixture] of Object.entries(fixtures)) {
    for (const image of collect(parse(name, fixture)).images) {
      const type = imageBlobType(image.bytes, image.mimeType);
      assert.ok(type === undefined || type === "image/svg+xml", `${name}: ${image.mimeType} -> ${type}`);
      routes[type ? "svg" : "refused"]++;
    }
  }
  // Pictures, shape fills, backgrounds and spreadsheet drawings all carry the SVG; mislabelled copies are refused.
  assert.ok(routes.svg >= 4, JSON.stringify(routes));
  assert.ok(routes.refused >= 2, JSON.stringify(routes));
  for (const name of ["pptx-svg-picture.pptx", "pptx-svg-shape.pptx", "pptx-svg-background.pptx", "xlsx-svg-drawing.xlsx", "pptx-svg-as-png.pptx"]) {
    assert.ok(collect(parse(name, fixtures[name])).images.length >= 1, `${name} produces an image in the model`);
  }
});

// ------------------------------------------------------------------------------------------------
// Source guards and viewer containment
// ------------------------------------------------------------------------------------------------

const rendererSources = readdirSync(new URL("../src/renderers/", import.meta.url)).filter(name => name.endsWith(".ts"))
  .map(name => [name, readFileSync(new URL(`../src/renderers/${name}`, import.meta.url), "utf8")]);

test("renderers never assign concatenated cssText", () => {
  for (const [name, source] of rendererSources) {
    const code = source.replace(/\/\*[\s\S]*?\*\/|\/\/.*$/gm, "");
    assert.doesNotMatch(code, /\bcssText\b/, `${name} sets styles through cssText`);
  }
});

test("image blob URLs are created and released only in image-url.ts", () => {
  for (const [name, source] of rendererSources) {
    if (name === "image-url.ts") continue;
    // Releasing through revokeImageUrls keeps the per-document pixel accounting correct.
    assert.doesNotMatch(source, /createObjectURL|revokeObjectURL|new Blob\(/, `${name} creates or revokes a blob outside image-url.ts`);
  }
});

test("the viewer root contains layout, paint and style, isolates stacking and hides overflow", async () => {
  const injected = readFileSync(new URL("../dist/api/file-previewer.js", import.meta.url), "utf8");
  const sheet = readFileSync(new URL("../styles.css", import.meta.url), "utf8");
  for (const [label, css] of [["injected styles", injected], ["styles.css", sheet]]) {
    const rule = /\.ufp-root\s*\{([^}]*)\}/.exec(css)?.[1]?.replace(/\s+/g, "");
    assert.ok(rule, `${label} has a .ufp-root rule`);
    for (const declaration of ["contain:layoutpaintstyle", "isolation:isolate", "overflow:hidden"]) assert.ok(rule.split(";").includes(declaration), `${label}: ${declaration}`);
  }
});
