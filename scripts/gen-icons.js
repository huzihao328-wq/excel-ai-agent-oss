/*
 * 生成插件所需的 PNG 图标（无第三方依赖）。
 * 图案：Excel 绿色渐变圆角方块 + 白色四角星（AI 火花）。
 */
const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

const CRC_TABLE = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

function makePng(size) {
  const raw = Buffer.alloc(size * (size * 4 + 1));
  const r = size * 0.18; // 圆角半径
  const cx = (size - 1) / 2;
  const cy = (size - 1) / 2;
  const starR = size * 0.32;

  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0; // filter: none
    for (let x = 0; x < size; x++) {
      const o = y * (size * 4 + 1) + 1 + x * 4;
      // 圆角方块判定
      const dx = Math.max(r - x, x - (size - 1 - r), 0);
      const dy = Math.max(r - y, y - (size - 1 - r), 0);
      const inside = dx * dx + dy * dy <= r * r;
      if (!inside) {
        raw.writeUInt32BE(0x00000000, o);
        continue;
      }
      // 垂直渐变：#1a8c4e -> #0e6b3a
      const t = y / (size - 1);
      let R = Math.round(26 + (14 - 26) * t);
      let G = Math.round(140 + (107 - 140) * t);
      let B = Math.round(78 + (58 - 78) * t);
      // 四角星：|x|^0.6 + |y|^0.6 <= r^0.6
      const sx = Math.abs(x - cx) / starR;
      const sy = Math.abs(y - cy) / starR;
      const v = Math.pow(sx, 0.6) + Math.pow(sy, 0.6);
      if (v <= 1) {
        const edge = Math.min(1, (1 - v) * 6); // 简单抗锯齿
        R = Math.round(R + (255 - R) * edge);
        G = Math.round(G + (255 - G) * edge);
        B = Math.round(B + (255 - B) * edge);
      }
      raw[o] = R;
      raw[o + 1] = G;
      raw[o + 2] = B;
      raw[o + 3] = 255;
    }
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

const outDir = path.join(__dirname, "..", "assets");
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
for (const size of [16, 32, 64, 80, 128]) {
  fs.writeFileSync(path.join(outDir, `icon-${size}.png`), makePng(size));
}
console.log("icons generated in ./assets");
