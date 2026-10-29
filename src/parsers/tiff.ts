import { inflateRaw } from "../archive/inflate.js";

const MAX_PIXELS = 40_000_000;

/**
 * Converts the first image of a baseline TIFF (strips or tiles; uncompressed, PackBits, LZW or Deflate; bilevel, grayscale,
 * palette or RGB(A) with 8 bits per sample) to a PNG, since browsers cannot display TIFF. Returns undefined for anything else.
 */
export function tiffToPng(bytes: Uint8Array): Uint8Array | undefined {
  try {
    const image = decodeTiff(bytes);
    return image ? encodePng(image.width, image.height, image.rgba, image.profile) : undefined;
  } catch {
    return undefined;
  }
}

function decodeTiff(bytes: Uint8Array): { width: number; height: number; rgba: Uint8Array; profile?: Uint8Array } | undefined {
  if (bytes.length < 8) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const little = bytes[0] === 0x49 && bytes[1] === 0x49;
  if (!little && !(bytes[0] === 0x4d && bytes[1] === 0x4d)) return undefined;
  const u16 = (offset: number) => view.getUint16(offset, little), u32 = (offset: number) => view.getUint32(offset, little);
  if (u16(2) !== 42) return undefined;
  const ifd = u32(4), count = u16(ifd), tags = new Map<number, number[]>();
  let profile: Uint8Array | undefined;
  const typeSizes: Record<number, number> = { 1: 1, 3: 2, 4: 4, 16: 8 };
  for (let index = 0; index < count; index++) {
    const entry = ifd + 2 + index * 12, tag = u16(entry), type = u16(entry + 2), values = u32(entry + 4), size = typeSizes[type];
    // Embedded ICC color profile: Office color-manages the picture with it, so it is carried into the PNG.
    if (tag === 34675 && type === 7 && values > 128 && values <= 1 << 20) { const start = u32(entry + 8); if (start + values <= bytes.length) profile = bytes.slice(start, start + values); continue; }
    if (!size || values > 1 << 24) continue;
    const start = size * values <= 4 ? entry + 8 : u32(entry + 8), list: number[] = [];
    for (let item = 0; item < values; item++) {
      const offset = start + item * size;
      list.push(size === 1 ? bytes[offset]! : size === 2 ? u16(offset) : size === 4 ? u32(offset) : Number(view.getBigUint64(offset, little)));
    }
    tags.set(tag, list);
  }
  const one = (tag: number, fallback: number) => tags.get(tag)?.[0] ?? fallback;
  const width = one(256, 0), height = one(257, 0), compression = one(259, 1), photometric = one(262, 2), samples = one(277, 1), predictor = one(317, 1);
  const bits = tags.get(258) ?? [1], depth = bits[0]!;
  if (!width || !height || width * height > MAX_PIXELS || one(284, 1) !== 1 || bits.some(value => value !== depth)) return undefined;
  if (!(depth === 8 || depth === 1 && samples === 1 && (photometric === 0 || photometric === 1)) || ![1, 5, 8, 32946, 32773].includes(compression)) return undefined;
  const tiled = tags.has(322);
  const blockWidth = tiled ? one(322, width) : width, blockHeight = tiled ? one(323, height) : Math.min(height, one(278, height));
  const offsets = tags.get(tiled ? 324 : 273) ?? [], counts = tags.get(tiled ? 325 : 279) ?? [];
  const rowBytes = Math.ceil(blockWidth * samples * depth / 8), blockBytes = rowBytes * blockHeight;
  const across = Math.ceil(width / blockWidth), down = Math.ceil(height / blockHeight);
  if (offsets.length < across * down) return undefined;
  const colorMap = tags.get(320), alpha = samples === 4 || samples === 2 ? (tags.get(338)?.[0] ?? 2) : 0;
  const rgba = new Uint8Array(width * height * 4);
  for (let block = 0; block < across * down; block++) {
    const offset = offsets[block]!, length = counts[block] ?? bytes.length - offset;
    const data = decompress(bytes.subarray(offset, Math.min(bytes.length, offset + length)), compression, blockBytes);
    if (predictor === 2 && depth === 8) for (let row = 0; row < blockHeight; row++) for (let x = samples; x < rowBytes; x++) data[row * rowBytes + x] = (data[row * rowBytes + x]! + data[row * rowBytes + x - samples]!) & 255;
    const left = (block % across) * blockWidth, top = Math.floor(block / across) * blockHeight;
    for (let row = 0; row < blockHeight && top + row < height; row++) for (let column = 0; column < blockWidth && left + column < width; column++) {
      const target = ((top + row) * width + left + column) * 4, source = row * rowBytes + column * samples;
      if (depth === 1) {
        const bit = (data[row * rowBytes + (column >> 3)]! >> (7 - (column & 7))) & 1, value = (photometric === 0 ? 1 - bit : bit) * 255;
        rgba.set([value, value, value, 255], target);
        continue;
      }
      let r: number, g: number, b: number, a = 255;
      if (photometric === 2 && samples >= 3) { r = data[source]!; g = data[source + 1]!; b = data[source + 2]!; if (alpha && samples === 4) a = data[source + 3]!; }
      else if (photometric === 3 && colorMap) { const index = data[source]!, size = colorMap.length / 3; r = colorMap[index]! >> 8; g = colorMap[size + index]! >> 8; b = colorMap[2 * size + index]! >> 8; }
      else if (photometric === 0 || photometric === 1) { const value = photometric === 0 ? 255 - data[source]! : data[source]!; r = g = b = value; if (alpha && samples === 2) a = data[source + 1]!; }
      else return undefined;
      // Associated (premultiplied) alpha is stored multiplied in; PNG wants straight alpha.
      if (alpha === 1 && a > 0 && a < 255) { r = Math.min(255, Math.round(r * 255 / a)); g = Math.min(255, Math.round(g * 255 / a)); b = Math.min(255, Math.round(b * 255 / a)); }
      rgba[target] = r; rgba[target + 1] = g; rgba[target + 2] = b; rgba[target + 3] = a;
    }
  }
  return { width, height, rgba, ...(profile ? { profile } : {}) };
}

function decompress(input: Uint8Array, compression: number, size: number): Uint8Array {
  if (compression === 1) { const output = new Uint8Array(size); output.set(input.subarray(0, size)); return output; }
  if (compression === 32773) return unpackBits(input, size);
  if (compression === 5) return lzw(input, size);
  // Deflate strips are zlib streams: skip the two-byte header.
  return inflateRaw(input.subarray(2), size);
}

function unpackBits(input: Uint8Array, size: number): Uint8Array {
  const output = new Uint8Array(size);
  let read = 0, write = 0;
  while (read < input.length && write < size) {
    const header = (input[read++]! << 24) >> 24;
    if (header >= 0) { for (let index = 0; index <= header && write < size; index++) output[write++] = input[read++] ?? 0; }
    else if (header !== -128) { const value = input[read++] ?? 0; for (let index = 0; index < 1 - header && write < size; index++) output[write++] = value; }
  }
  return output;
}

/** TIFF LZW: MSB-first codes, 256 = clear, 257 = end, code width grows one code early. */
function lzw(input: Uint8Array, size: number): Uint8Array {
  const output = new Uint8Array(size);
  const prefixes = new Int32Array(4096), suffixes = new Uint8Array(4096), lengths = new Uint16Array(4096);
  for (let code = 0; code < 256; code++) { suffixes[code] = code; lengths[code] = 1; prefixes[code] = -1; }
  let bitPosition = 0, width = 9, next = 258, previous = -1, write = 0;
  const readCode = () => {
    let value = 0;
    for (let index = 0; index < width; index++, bitPosition++) value = (value << 1) | ((input[bitPosition >> 3]! >> (7 - (bitPosition & 7))) & 1);
    return value;
  };
  const firstByte = (code: number) => { while (prefixes[code]! >= 0) code = prefixes[code]!; return suffixes[code]!; };
  const emit = (code: number) => {
    let position = write + lengths[code]! - 1;
    for (let current = code; current >= 0; current = prefixes[current]!, position--) if (position < size) output[position] = suffixes[current]!;
    write += lengths[code]!;
  };
  const add = (prefix: number, suffix: number) => {
    if (next >= 4096) return;
    prefixes[next] = prefix; suffixes[next] = suffix; lengths[next] = lengths[prefix]! + 1; next++;
    if (next + 1 >= 1 << width && width < 12) width++;
  };
  while (bitPosition + width <= input.length * 8 && write < size) {
    const code = readCode();
    if (code === 257) break;
    if (code === 256) { width = 9; next = 258; previous = -1; continue; }
    if (previous < 0) { if (code > 255) break; emit(code); previous = code; continue; }
    if (code < next) { add(previous, firstByte(code)); emit(code); }
    else if (code === next) { add(previous, firstByte(previous)); emit(code); }
    else break;
    previous = code;
  }
  return output;
}

/** Writes an RGBA PNG (with an optional ICC profile) using stored (uncompressed) deflate blocks. */
function encodePng(width: number, height: number, rgba: Uint8Array, profile?: Uint8Array): Uint8Array {
  const stride = width * 4 + 1, raw = new Uint8Array(stride * height);
  for (let row = 0; row < height; row++) raw.set(rgba.subarray(row * width * 4, (row + 1) * width * 4), row * stride + 1);
  const header = new Uint8Array(13), headerView = new DataView(header.buffer);
  headerView.setUint32(0, width); headerView.setUint32(4, height); header[8] = 8; header[9] = 6;
  const chunks = [chunk("IHDR", header)];
  if (profile) { const name = new TextEncoder().encode("ICC Profile"), compressed = zlibStored(profile), data = new Uint8Array(name.length + 2 + compressed.length); data.set(name); data.set(compressed, name.length + 2); chunks.push(chunk("iCCP", data)); }
  chunks.push(chunk("IDAT", zlibStored(raw)), chunk("IEND", new Uint8Array(0)));
  const png = new Uint8Array(8 + chunks.reduce((sum, part) => sum + part.length, 0));
  png.set([137, 80, 78, 71, 13, 10, 26, 10]);
  let offset = 8;
  for (const part of chunks) { png.set(part, offset); offset += part.length; }
  return png;
}

/** A zlib stream of stored deflate blocks. */
function zlibStored(raw: Uint8Array): Uint8Array {
  const blocks = Math.max(1, Math.ceil(raw.length / 65_535)), zlib = new Uint8Array(2 + raw.length + blocks * 5 + 4);
  zlib[0] = 0x78; zlib[1] = 0x01;
  let position = 2;
  for (let block = 0; block < blocks; block++) {
    const start = block * 65_535, length = Math.min(65_535, raw.length - start);
    zlib[position++] = block === blocks - 1 ? 1 : 0;
    zlib[position++] = length & 255; zlib[position++] = length >> 8; zlib[position++] = ~length & 255; zlib[position++] = (~length >> 8) & 255;
    zlib.set(raw.subarray(start, start + length), position); position += length;
  }
  let a = 1, b = 0;
  for (const byte of raw) { a = (a + byte) % 65_521; b = (b + a) % 65_521; }
  new DataView(zlib.buffer).setUint32(position, ((b << 16) | a) >>> 0);
  return zlib;
}

let crcTable: Uint32Array | undefined;
function chunk(type: string, data: Uint8Array): Uint8Array {
  crcTable ??= Uint32Array.from({ length: 256 }, (_, index) => { let value = index; for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1; return value >>> 0; });
  const output = new Uint8Array(12 + data.length), view = new DataView(output.buffer);
  view.setUint32(0, data.length);
  for (let index = 0; index < 4; index++) output[4 + index] = type.charCodeAt(index);
  output.set(data, 8);
  let crc = 0xffffffff;
  for (let index = 4; index < 8 + data.length; index++) crc = crcTable[(crc ^ output[index]!) & 255]! ^ (crc >>> 8);
  view.setUint32(8 + data.length, (crc ^ 0xffffffff) >>> 0);
  return output;
}
