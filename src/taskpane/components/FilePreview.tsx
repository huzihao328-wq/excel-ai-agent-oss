import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  colIndexToLetter,
  getExternalFiles,
  parseA1Ranges,
  setExternalSelection,
} from "../../core/tools/externalFiles";
import { ToolResult } from "../../core/types";
import MotionTabs from "./MotionTabs";
import { IconClose, IconFile } from "./Icons";

interface Props {
  fileName: string;
  busy: boolean;
  onImport: (args: { file: string; sheet: string; range?: string }) => Promise<ToolResult>;
  onClose: () => void;
  onNotice: (text: string, isError?: boolean) => void;
}

interface SelRect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

type HitKind = "cell" | "col" | "row" | "all";

interface GridHit {
  kind: HitKind;
  r: number;
  c: number;
}

const ROWS_PER_PAGE = 100;
const MAX_COLS = 30;
const CELL_TEXT_LIMIT = 20;
const SCROLL_EDGE = 32;
const SCROLL_STEP = 18;

function readHit(target: EventTarget | null, root: HTMLElement | null): GridHit | null {
  if (!(target instanceof Element) || !root) return null;
  const node = target.closest("[data-hit]") as HTMLElement | null;
  if (!node || !root.contains(node)) return null;
  const kind = node.dataset.hit as HitKind | undefined;
  if (!kind) return null;
  const r = node.dataset.r !== undefined ? Number(node.dataset.r) : NaN;
  const c = node.dataset.c !== undefined ? Number(node.dataset.c) : NaN;
  return { kind, r, c };
}

function hitFromPoint(clientX: number, clientY: number, root: HTMLElement | null): GridHit | null {
  const direct = readHit(document.elementFromPoint(clientX, clientY), root);
  if (direct) return direct;
  return hitFromGeometry(clientX, clientY, root);
}

function formatRect(s: SelRect): string {
  const a = `${colIndexToLetter(s.c1)}${s.r1 + 1}`;
  const b = `${colIndexToLetter(s.c2)}${s.r2 + 1}`;
  return a === b ? a : `${a}:${b}`;
}

function inAnySel(sels: SelRect[], r: number, c: number): boolean {
  return sels.some((s) => r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2);
}

function colInSel(sels: SelRect[], c: number): boolean {
  return sels.some((s) => c >= s.c1 && c <= s.c2);
}

function rowInSel(sels: SelRect[], r: number): boolean {
  return sels.some((s) => r >= s.r1 && r <= s.r2);
}

function toggleColumn(sels: SelRect[], col: number): SelRect[] {
  const next: SelRect[] = [];
  let removed = false;
  for (const s of sels) {
    if (col < s.c1 || col > s.c2) {
      next.push(s);
      continue;
    }
    removed = true;
    if (s.c1 < col) next.push({ ...s, c2: col - 1 });
    if (s.c2 > col) next.push({ ...s, c1: col + 1 });
  }
  return removed ? next : sels;
}

function makeRect(anchor: GridHit, current: GridHit, maxR: number, maxC: number): SelRect {
  if (anchor.kind === "all") return { r1: 0, c1: 0, r2: maxR, c2: maxC };
  if (anchor.kind === "col") {
    return {
      r1: 0,
      c1: Math.min(anchor.c, current.c),
      r2: maxR,
      c2: Math.max(anchor.c, current.c),
    };
  }
  if (anchor.kind === "row") {
    return {
      r1: Math.min(anchor.r, current.r),
      c1: 0,
      r2: Math.max(anchor.r, current.r),
      c2: maxC,
    };
  }
  return {
    r1: Math.min(anchor.r, current.r),
    c1: Math.min(anchor.c, current.c),
    r2: Math.max(anchor.r, current.r),
    c2: Math.max(anchor.c, current.c),
  };
}

/** 指针在格子间隙、表外或命中失败时，按位置夹到最近的行列 */
function hitFromGeometry(clientX: number, clientY: number, root: HTMLElement | null): GridHit | null {
  if (!root) return null;
  const table = root.querySelector("table.preview-grid");
  if (!table) return null;
  const heads = table.querySelectorAll("thead th");
  const bodyRows = table.querySelectorAll("tbody tr");
  if (heads.length < 2 || bodyRows.length === 0) return null;

  const cornerRect = heads[0].getBoundingClientRect();
  const firstColRect = heads[1].getBoundingClientRect();
  const lastColRect = heads[heads.length - 1].getBoundingClientRect();
  const firstRowRect = bodyRows[0].getBoundingClientRect();
  const lastRowRect = bodyRows[bodyRows.length - 1].getBoundingClientRect();

  let col = 0;
  for (let i = 1; i < heads.length; i++) {
    if (clientX >= heads[i].getBoundingClientRect().left) col = i - 1;
  }
  if (clientX < firstColRect.left) col = 0;
  if (clientX > lastColRect.right) col = heads.length - 2;

  let rowEl = bodyRows[0].querySelector("[data-r]") as HTMLElement | null;
  let row = rowEl?.dataset.r !== undefined ? Number(rowEl.dataset.r) : 0;
  for (let i = 0; i < bodyRows.length; i++) {
    if (clientY >= bodyRows[i].getBoundingClientRect().top) {
      rowEl = bodyRows[i].querySelector("[data-r]") as HTMLElement | null;
      if (rowEl?.dataset.r !== undefined) row = Number(rowEl.dataset.r);
    }
  }
  if (clientY < firstRowRect.top) {
    rowEl = bodyRows[0].querySelector("[data-r]") as HTMLElement | null;
    row = rowEl?.dataset.r !== undefined ? Number(rowEl.dataset.r) : 0;
  }
  if (clientY > lastRowRect.bottom) {
    rowEl = bodyRows[bodyRows.length - 1].querySelector("[data-r]") as HTMLElement | null;
    row = rowEl?.dataset.r !== undefined ? Number(rowEl.dataset.r) : row;
  }

  const overHeader = clientY < firstRowRect.top;
  const overRowNo = clientX < cornerRect.right;
  if (overHeader && overRowNo) return { kind: "all", r: NaN, c: NaN };
  if (overHeader) return { kind: "col", r: NaN, c: col };
  if (overRowNo) return { kind: "row", r: row, c: NaN };
  return { kind: "cell", r: row, c: col };
}

export default function FilePreview({ fileName, busy, onImport, onClose, onNotice }: Props) {
  const file = getExternalFiles().find((f) => f.name === fileName);
  const [sheetIdx, setSheetIdx] = useState(0);
  const [page, setPage] = useState(0);
  const [sels, setSels] = useState<SelRect[]>([]);
  const [manualRange, setManualRange] = useState("");
  const [manualError, setManualError] = useState("");
  const [importing, setImporting] = useState(false);
  const importingRef = useRef(false);
  const gridWrapRef = useRef<HTMLDivElement>(null);
  const anchorRef = useRef<GridHit | null>(null);
  const lastHitRef = useRef<GridHit | null>(null);
  const selectingRef = useRef(false);
  const selectionFrameRef = useRef<number>();
  const pointerRef = useRef<{ x: number; y: number }>();
  const touchStartRef = useRef<{ x: number; y: number }>();
  const selsRef = useRef<SelRect[]>([]);
  const baseSelsRef = useRef<SelRect[]>([]);
  const boundsRef = useRef({ totalRows: 0, lastCol: 0 });
  selsRef.current = sels;

  useEffect(() => {
    const stop = () => {
      selectingRef.current = false;
      if (selectionFrameRef.current !== undefined) cancelAnimationFrame(selectionFrameRef.current);
      selectionFrameRef.current = undefined;
    };
    window.addEventListener("pointerup", stop);
    window.addEventListener("mouseup", stop);
    return () => {
      stop();
      window.removeEventListener("pointerup", stop);
      window.removeEventListener("mouseup", stop);
    };
  }, []);

  const sheet = file?.sheets[Math.min(sheetIdx, (file?.sheets.length ?? 1) - 1)];
  const totalRows = sheet?.rows.length ?? 0;
  const totalCols = useMemo(
    () => (sheet ? sheet.rows.reduce((max, row) => Math.max(max, row.length), 0) : 0),
    [sheet]
  );
  const dispCols = Math.min(totalCols, MAX_COLS);
  const lastCol = Math.max(0, dispCols - 1);
  boundsRef.current = { totalRows, lastCol };
  const pageCount = Math.max(1, Math.ceil(totalRows / ROWS_PER_PAGE));
  const rowStart = page * ROWS_PER_PAGE;
  const pageRows = useMemo(() => sheet?.rows.slice(rowStart, rowStart + ROWS_PER_PAGE) ?? [], [sheet, rowStart]);

  const switchSheet = (idx: number) => {
    setSheetIdx(idx);
    setPage(0);
    setSels([]);
    setManualRange("");
    setManualError("");
  };

  const addr = useMemo(() => (sels.length ? sels.map(formatRect).join(",") : ""), [sels]);

  const resolveHit = useCallback((hit: GridHit, fallback: GridHit | null): GridHit => {
    return {
      kind: hit.kind,
      r: Number.isFinite(hit.r) ? hit.r : fallback?.r ?? 0,
      c: Number.isFinite(hit.c) ? hit.c : fallback?.c ?? 0,
    };
  }, []);

  const applySelection = useCallback((anchor: GridHit, current: GridHit) => {
    const { lastCol: maxC, totalRows } = boundsRef.current;
    const maxR = Math.max(0, totalRows - 1);
    if (anchor.kind === "all") {
      setSels([{ r1: 0, c1: 0, r2: maxR, c2: maxC }]);
      return;
    }
    const rect = makeRect(anchor, current, maxR, maxC);
    setSels([...baseSelsRef.current, rect]);
  }, []);

  const autoScroll = useCallback((clientX: number, clientY: number) => {
    const wrap = gridWrapRef.current;
    if (!wrap) return;
    const rect = wrap.getBoundingClientRect();
    let dx = 0;
    let dy = 0;
    if (clientX > rect.right - SCROLL_EDGE) dx = SCROLL_STEP;
    else if (clientX < rect.left + SCROLL_EDGE) dx = -SCROLL_STEP;
    if (clientY > rect.bottom - SCROLL_EDGE) dy = SCROLL_STEP;
    else if (clientY < rect.top + SCROLL_EDGE) dy = -SCROLL_STEP;
    if (dx || dy) wrap.scrollBy(dx, dy);
  }, []);

  const beginSelect = useCallback(
    (hit: GridHit, additive: boolean) => {
      const resolved = resolveHit(hit, hit);
      setManualError("");

      if (additive && hit.kind === "col" && Number.isFinite(hit.c) && colInSel(selsRef.current, hit.c)) {
        setSels(toggleColumn(selsRef.current, hit.c));
        selectingRef.current = false;
        return;
      }

      baseSelsRef.current = hit.kind === "all" ? [] : additive ? selsRef.current : [];
      anchorRef.current = resolved;
      lastHitRef.current = resolved;
      selectingRef.current = true;
      applySelection(resolved, resolved);
    },
    [applySelection, resolveHit]
  );

  const extendSelect = useCallback(
    (clientX: number, clientY: number) => {
      if (!selectingRef.current || !anchorRef.current) return;
      autoScroll(clientX, clientY);
      const raw = hitFromPoint(clientX, clientY, gridWrapRef.current);
      if (!raw) return;
      const resolved = resolveHit(raw, lastHitRef.current);
      if (resolved.r === lastHitRef.current?.r && resolved.c === lastHitRef.current?.c) return;
      lastHitRef.current = resolved;
      applySelection(anchorRef.current, resolved);
    },
    [applySelection, autoScroll, resolveHit]
  );

  const onGridPointerDown = useCallback(
    (e: React.PointerEvent<HTMLDivElement>) => {
      if (e.button !== 0) return;
      if (e.pointerType === "touch") {
        touchStartRef.current = { x: e.clientX, y: e.clientY };
        return; // 触屏拖动交给原生滚动；轻点单元格仍可选中。
      }
      const hit = readHit(e.target, gridWrapRef.current);
      if (!hit) return;
      e.preventDefault();
      try {
        e.currentTarget.setPointerCapture(e.pointerId);
      } catch {
        /* 部分 WebView 不支持 pointer capture，仍可用格子内的 pointermove */
      }
      beginSelect(hit, e.ctrlKey || e.metaKey);
      pointerRef.current = { x: e.clientX, y: e.clientY };
    },
    [beginSelect]
  );

  const onGridPointerMove = useCallback(
    (e: React.PointerEvent<HTMLDivElement>) => {
      if (!selectingRef.current) return;
      pointerRef.current = { x: e.clientX, y: e.clientY };
      if (selectionFrameRef.current !== undefined) return;
      const tick = () => {
        if (!selectingRef.current || !pointerRef.current) { selectionFrameRef.current = undefined; return; }
        extendSelect(pointerRef.current.x, pointerRef.current.y);
        selectionFrameRef.current = requestAnimationFrame(tick);
      };
      selectionFrameRef.current = requestAnimationFrame(tick);
    },
    [extendSelect]
  );

  const onGridPointerUp = useCallback((e: React.PointerEvent<HTMLDivElement>) => {
    if (e.type !== "pointercancel") {
      if (e.pointerType === "touch" && touchStartRef.current && Math.hypot(e.clientX - touchStartRef.current.x, e.clientY - touchStartRef.current.y) < 8) {
        const hit = readHit(e.target, gridWrapRef.current);
        if (hit) beginSelect(hit, false);
      } else if (selectingRef.current) extendSelect(e.clientX, e.clientY);
    }
    touchStartRef.current = undefined;
    selectingRef.current = false;
    if (selectionFrameRef.current !== undefined) cancelAnimationFrame(selectionFrameRef.current);
    selectionFrameRef.current = undefined;
    try {
      if (e.currentTarget.hasPointerCapture?.(e.pointerId)) {
        e.currentTarget.releasePointerCapture(e.pointerId);
      }
    } catch {
      /* ignore */
    }
  }, [beginSelect, extendSelect]);

  const applyManual = () => {
    if (!manualRange.trim()) return;
    try {
      const rects = parseA1Ranges(manualRange.trim()).map((p) => ({
        r1: Math.min(p.r1, Math.max(0, totalRows - 1)),
        c1: Math.min(p.c1, Math.max(0, totalCols - 1)),
        r2: Math.min(p.r2, Math.max(0, totalRows - 1)),
        c2: Math.min(p.c2, Math.max(0, totalCols - 1)),
      }));
      setSels(rects);
      if (rects[0]) setPage(Math.floor(rects[0].r1 / ROWS_PER_PAGE));
      setManualError("");
    } catch (e: any) {
      setManualError(e?.message || "区域格式不正确");
    }
  };

  if (!file || !sheet) {
    return null;
  }

  const refer = (range: string) => {
    setExternalSelection({ file: file.name, sheet: sheet.name, range });
    onNotice(`⬚ 已引用 ${file.caption} / ${sheet.name} 的 ${range}，直接告诉我要对这块数据做什么`);
    onClose();
  };

  const importRange = async (range?: string) => {
    if (busy || importingRef.current) return;
    importingRef.current = true;
    setImporting(true);
    try {
      const result = await onImport({ file: file.name, sheet: sheet.name, range });
      if (result.ok) {
        onNotice(result.summary || "导入完成");
        onClose();
      } else {
        onNotice(result.error || "导入失败", true);
      }
    } catch (e: any) {
      onNotice(`导入失败: ${e?.message || String(e)}`, true);
    } finally {
      importingRef.current = false;
      setImporting(false);
    }
  };

  const fullRange = totalRows > 0 ? `A1:${colIndexToLetter(Math.max(0, totalCols - 1))}${totalRows}` : "";

  return (
    <div className="preview-overlay" onClick={onClose}>
      <div className="preview-panel" onClick={(e) => e.stopPropagation()}>
        <div className="preview-header">
            <div className="preview-title" title={file.caption}><IconFile size={15} /> <span>{file.caption}</span></div>
          <button className="icon-btn" onClick={onClose} title="关闭" aria-label="关闭"><IconClose size={15} /></button>
        </div>

        {file.sheets.length > 1 && (
          <MotionTabs className="preview-tabs" buttonClassName="preview-tab" label="文件工作表" panelId="preview-sheet-grid"
            options={file.sheets.map((s, i) => ({ value: String(i), label: s.name }))} value={String(sheetIdx)} onChange={value => switchSheet(Number(value))} />
        )}

        <div className="preview-toolbar">
          <span className="preview-hint">
            {totalRows}×{totalCols}
            {totalCols > MAX_COLS ? `（仅显示前 ${MAX_COLS} 列）` : ""}
            ，拖动框选；按住 Ctrl 再框选可追加更多列
          </span>
          <div className="preview-range-input">
            <input
              value={manualRange}
              placeholder="或输入区域 如 A2:A50,C2:C50"
              onChange={(e) => setManualRange(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && applyManual()}
              spellCheck={false}
            />
            <button className="inline-btn" onClick={applyManual}>定位</button>
          </div>
        </div>
        {manualError && <div className="preview-error">{manualError}</div>}

        <div
          className="preview-grid-wrap"
          id="preview-sheet-grid" role="tabpanel" aria-label={sheet.name} tabIndex={0}
          ref={gridWrapRef}
          onPointerDown={onGridPointerDown}
          onPointerMove={onGridPointerMove}
          onPointerUp={onGridPointerUp}
          onPointerCancel={onGridPointerUp}
          onDragStart={(e) => e.preventDefault()}
        >
          <table className="preview-grid">
            <thead>
              <tr>
                <th className="rowno-cell corner" data-hit="all" />
                {Array.from({ length: dispCols }, (_, c) => (
                  <th
                    key={c}
                    data-hit="col"
                    data-c={c}
                    className={colInSel(sels, c) ? "col-head hl" : "col-head"}
                  >
                    {colIndexToLetter(c)}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {pageRows.map((row, i) => {
                const r = rowStart + i;
                const rowHl = rowInSel(sels, r);
                return (
                  <tr key={r}>
                    <td
                      data-hit="row"
                      data-r={r}
                      className={rowHl ? "rowno-cell hl" : "rowno-cell"}
                    >
                      {r + 1}
                    </td>
                    {Array.from({ length: dispCols }, (_, c) => {
                      const inSel = inAnySel(sels, r, c);
                      const v = row[c];
                      const text =
                        v === undefined || v === null || v === ""
                          ? ""
                          : String(v).slice(0, CELL_TEXT_LIMIT);
                      return (
                        <td
                          key={c}
                          data-hit="cell"
                          data-r={r}
                          data-c={c}
                          draggable={false}
                          className={inSel ? "data-cell sel" : "data-cell"}
                          title={v !== undefined && v !== null ? String(v) : ""}
                        >
                          {text}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {pageCount > 1 && (
          <div className="preview-pager">
            <button className="btn tiny" disabled={page === 0} onClick={() => setPage((p) => p - 1)}>
              上一页
            </button>
            <span>
              第 {rowStart + 1}~{Math.min(rowStart + ROWS_PER_PAGE, totalRows)} 行 / 共 {totalRows} 行
            </span>
            <button className="btn tiny" disabled={page >= pageCount - 1} onClick={() => setPage((p) => p + 1)}>
              下一页
            </button>
          </div>
        )}

        <div className="preview-footer">
          <span className="preview-sel-addr" title={addr || undefined}>
            {addr ? `已选: ${addr}` : "未选择区域"}
          </span>
          <div className="preview-actions">
            <button className="btn secondary" disabled={!fullRange} onClick={() => refer(fullRange)}>
              引用整表
            </button>
            <button className="btn secondary" disabled={!addr || importing || busy} onClick={() => importRange(addr)}>
              {importing ? "导入中…" : "导入选区"}
            </button>
            <button className="btn primary" disabled={!addr} onClick={() => refer(addr)}>
              引用选区
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
