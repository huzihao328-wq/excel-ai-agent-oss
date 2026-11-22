export interface AttachedImage {
  id: string;
  name: string;
  mime: string;
  dataUrl: string;
  width: number;
  height: number;
}

const IMAGE_EXT = /\.(png|jpe?g|gif|webp|bmp)$/i;
const IMAGE_MIME = /^image\/(png|jpeg|jpg|gif|webp|bmp)$/i;
const MAX_IMAGE_BYTES = 8 * 1024 * 1024;
const MAX_KEEP_ORIGINAL_BYTES = 1.2 * 1024 * 1024;
const MAX_DIMENSION = 1280;
const JPEG_QUALITY = 0.82;

export const MAX_IMAGES_PER_MESSAGE = 6;

let imageId = 0;
const nextId = () => `img_${Date.now()}_${imageId++}`;

export function isImageFile(file: File): boolean {
  return IMAGE_MIME.test(file.type) || IMAGE_EXT.test(file.name);
}

function guessMime(name: string): string {
  const ext = name.split(".").pop()?.toLowerCase();
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "gif") return "image/gif";
  if (ext === "webp") return "image/webp";
  if (ext === "bmp") return "image/bmp";
  return "image/png";
}

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error(`读取图片失败: ${file.name}`));
    reader.readAsDataURL(file);
  });
}

function loadHtmlImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("图片解码失败"));
    img.src = src;
  });
}

/** 读取并压缩图片，避免把过大的原图发给模型和写入本地历史 */
export async function loadAttachedImage(file: File): Promise<AttachedImage> {
  if (file.size > MAX_IMAGE_BYTES) {
    throw new Error(`${file.name} 超过 8MB，请压缩后再发送`);
  }
  const original = await readFileAsDataUrl(file);
  const img = await loadHtmlImage(original);
  const width = img.naturalWidth || img.width;
  const height = img.naturalHeight || img.height;
  const scale = Math.min(1, MAX_DIMENSION / Math.max(width, height, 1));
  const w = Math.max(1, Math.round(width * scale));
  const h = Math.max(1, Math.round(height * scale));
  const mime = file.type && IMAGE_MIME.test(file.type) ? file.type : guessMime(file.name);

  if (scale === 1 && file.size <= MAX_KEEP_ORIGINAL_BYTES) {
    return { id: nextId(), name: file.name || "image.png", mime, dataUrl: original, width, height };
  }

  const canvas = document.createElement("canvas");
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("无法处理图片");
  ctx.drawImage(img, 0, 0, w, h);
  const outMime = mime === "image/png" ? "image/png" : "image/jpeg";
  const dataUrl =
    outMime === "image/png" ? canvas.toDataURL("image/png") : canvas.toDataURL("image/jpeg", JPEG_QUALITY);
  return { id: nextId(), name: file.name || "image.png", mime: outMime, dataUrl, width: w, height: h };
}

export function collectClipboardImages(dt: DataTransfer): File[] {
  const files: File[] = [];
  for (const item of Array.from(dt.items || [])) {
    if (item.kind === "file" && item.type.startsWith("image/")) {
      const file = item.getAsFile();
      if (file) files.push(file);
    }
  }
  if (files.length === 0) {
    for (const file of Array.from(dt.files || [])) {
      if (isImageFile(file)) files.push(file);
    }
  }
  return files;
}
