import { deflateRawSync } from "node:zlib";

export function makeZip(entries, { deflate = false } = {}) {
  const localParts = [], centralParts = [];
  let localOffset = 0;
  for (const [name, rawValue] of Object.entries(entries)) {
    const nameBytes = Buffer.from(name), raw = Buffer.from(rawValue), compressed = deflate ? deflateRawSync(raw) : raw;
    const method = deflate ? 8 : 0, crc = crc32(raw);
    const local = Buffer.alloc(30 + nameBytes.length + compressed.length);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x800, 6);
    local.writeUInt16LE(method, 8); local.writeUInt32LE(crc, 14); local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(raw.length, 22); local.writeUInt16LE(nameBytes.length, 26); nameBytes.copy(local, 30); compressed.copy(local, 30 + nameBytes.length);
    localParts.push(local);
    const central = Buffer.alloc(46 + nameBytes.length);
    central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6); central.writeUInt16LE(0x800, 8);
    central.writeUInt16LE(method, 10); central.writeUInt32LE(crc, 16); central.writeUInt32LE(compressed.length, 20);
    central.writeUInt32LE(raw.length, 24); central.writeUInt16LE(nameBytes.length, 28); central.writeUInt32LE(localOffset, 42); nameBytes.copy(central, 46);
    centralParts.push(central); localOffset += local.length;
  }
  const central = Buffer.concat(centralParts), body = Buffer.concat(localParts), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(Object.keys(entries).length, 8); end.writeUInt16LE(Object.keys(entries).length, 10);
  end.writeUInt32LE(central.length, 12); end.writeUInt32LE(body.length, 16);
  return Buffer.concat([body, central, end]);
}

export function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) { crc ^= byte; for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0); }
  return (crc ^ 0xffffffff) >>> 0;
}
