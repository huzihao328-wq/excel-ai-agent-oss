/** 把表格导出为可下载文件。不依赖 Excel，便于单测。 */

export function sanitizeDownloadName(name: string, fallback: string): string {
  const cleaned = String(name || "")
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim();
  const base = cleaned || fallback;
  return base.length > 120 ? base.slice(0, 120) : base;
}

export function ensureExtension(name: string, ext: "xlsx" | "csv"): string {
  const lower = name.toLowerCase();
  if (lower.endsWith(`.${ext}`)) return name;
  return `${name.replace(/\.(xlsx|xlsm|xls|csv)$/i, "")}.${ext}`;
}

function csvCell(value: unknown): string {
  if (value == null || value === "") return "";
  const s = typeof value === "boolean" ? (value ? "TRUE" : "FALSE") : String(value);
  if (/[",\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

export function gridToCsv(rows: unknown[][]): string {
  return rows.map((row) => (Array.isArray(row) ? row : [row]).map(csvCell).join(",")).join("\r\n");
}

export function csvBlob(rows: unknown[][]): Blob {
  return new Blob(["\uFEFF" + gridToCsv(rows)], { type: "text/csv;charset=utf-8" });
}

export async function xlsxBlob(sheets: { name: string; rows: unknown[][] }[]): Promise<Blob> {
  const XLSX = await import("xlsx");
  const wb = XLSX.utils.book_new();
  const used = new Set<string>();
  for (const sheet of sheets) {
    let name = sanitizeDownloadName(sheet.name, "Sheet").slice(0, 31) || "Sheet";
    if (used.has(name)) {
      let i = 2;
      while (used.has(`${name.slice(0, 28)}(${i})`)) i++;
      name = `${name.slice(0, 28)}(${i})`;
    }
    used.add(name);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheet.rows), name);
  }
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const bytes = out instanceof ArrayBuffer ? new Uint8Array(out) : new Uint8Array(out);
  return new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export function triggerDownload(filename: string, blob: Blob): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.rel = "noopener";
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  a.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 4000);
}
