import { useEffect, useRef, useState } from "react";
import Collapse from "./Collapse";
import { EmptyState, PanelHeader } from "./Panel";
import { IconCheck, IconChevronRight, IconClose, IconFile, IconLayers, IconPlus } from "./Icons";
import { BatchFileState, runBatch } from "../../core/batch";
import { ExternalFile, getExternalFiles, loadExternalFile, onExternalFilesChange } from "../../core/tools/externalFiles";
import { ToolExecContext } from "../../core/tools/excelTools";
import { AppSettings } from "../../core/types";
import { isLoopbackEndpoint } from "../../core/settings";

interface Props {
  settings: AppSettings;
  mainRunning: boolean;
  /** 构建工具执行上下文（确认卡会出现在主聊天流中） */
  buildExecCtx: () => ToolExecContext;
  onClose: () => void;
  /** 批量任务结束时收到一条汇总文案，由主界面以 notice 形式展示 */
  onFinish: (summary: string) => void;
  onRunningChange: (running: boolean) => void;
  /** 中止时解开主聊天里挂起的确认/选择卡片 */
  onAbortCleanup: () => void;
  registerAbort: (fn: (() => void) | null) => void;
}

const STATUS_ICONS: Record<BatchFileState["status"], JSX.Element> = {
  pending: <span className="batch-dot" />,
  running: <span className="spinner" />,
  done: <IconCheck size={13} />,
  incomplete: <span className="batch-bang">!</span>,
  failed: <IconClose size={12} />,
  aborted: <span className="batch-dot hollow" />,
};

export default function BatchPanel({
  settings,
  mainRunning,
  buildExecCtx,
  onClose,
  onFinish,
  onRunningChange,
  onAbortCleanup,
  registerAbort,
}: Props) {
  const [extFiles, setExtFiles] = useState<ExternalFile[]>(getExternalFiles);
  const [selected, setSelected] = useState<Set<string>>(() => new Set(getExternalFiles().map((f) => f.name)));
  const [instruction, setInstruction] = useState("");
  const [phase, setPhase] = useState<"idle" | "running" | "done">("idle");
  const [states, setStates] = useState<BatchFileState[]>([]);
  const [expanded, setExpanded] = useState<string | null>(null);
  const [error, setError] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const abortRef = useRef<AbortController | null>(null);
  const startingRef = useRef(false);
  const knownFilesRef = useRef(new Set(extFiles.map((file) => file.name)));

  useEffect(() => onExternalFilesChange(() => {
    const files = getExternalFiles();
    const previous = knownFilesRef.current;
    knownFilesRef.current = new Set(files.map((file) => file.name));
    setExtFiles(files);
    setSelected((selected) => new Set(files.filter((file) => selected.has(file.name) || !previous.has(file.name)).map((file) => file.name)));
  }), []);

  useEffect(() => {
    registerAbort(() => abortRef.current?.abort());
    return () => {
      const wasRunning = !!abortRef.current;
      abortRef.current?.abort();
      if (wasRunning) onAbortCleanup();
      registerAbort(null);
    };
  }, [onAbortCleanup, registerAbort]);

  const busy = phase === "running";
  const selectedNames = extFiles.filter((f) => selected.has(f.name)).map((f) => f.name);

  const toggleFile = (name: string) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name);
      else next.add(name);
      return next;
    });
  };

  const addFiles = async (fileList: FileList | null) => {
    if (!fileList) return;
    for (const file of Array.from(fileList)) {
      try {
        await loadExternalFile(file);
        setSelected((prev) => new Set(prev).add(file.name));
      } catch (e: any) {
        setError(`附加文件失败: ${e?.message || String(e)}`);
      }
    }
  };

  const start = async () => {
    if (!instruction.trim() || selectedNames.length === 0 || startingRef.current || busy || mainRunning) return;
    if (!settings.apiKey.trim() && !isLoopbackEndpoint(settings.baseUrl)) {
      setError("请先在设置里填写 API Key");
      return;
    }
    if (!settings.model.trim()) {
      setError("请先在模型设置里选择或填写模型名");
      return;
    }
    startingRef.current = true;
    setError("");
    setPhase("running");
    setStates(selectedNames.map((fileName) => ({ fileName, status: "pending", toolTrace: [] })));
    onRunningChange(true);
    const abort = new AbortController();
    abortRef.current = abort;
    try {
      const result = await runBatch(selectedNames, instruction, settings, buildExecCtx(), (patch) => {
        setStates((prev) => prev.map((s) => (s.fileName === patch.fileName ? { ...s, ...patch } : s)));
      }, abort.signal);
      setPhase("done");
      const ok = result.states.filter((s) => s.status === "done").length;
      const failed = result.states.filter((s) => s.status === "failed").length;
      const incomplete = result.states.filter((s) => s.status === "incomplete").length;
      const aborted = result.states.filter((s) => s.status === "aborted").length;
      onFinish(
        result.aborted
          ? `🗂 批量任务已中止：完成 ${ok} 个、需检查 ${incomplete} 个、失败 ${failed} 个、中止 ${aborted} 个。`
          : `🗂 批量任务执行结束：${ok} 个完成${incomplete ? `、${incomplete} 个需检查` : ""}${failed ? `、${failed} 个失败` : ""}。各文件结果和执行详情见批量面板。`
      );
    } catch (e: any) {
      setError(e?.message || String(e));
      setPhase("done");
    } finally {
      startingRef.current = false;
      abortRef.current = null;
      onRunningChange(false);
    }
  };

  const reset = () => {
    if (busy) return;
    setPhase("idle");
    setStates([]);
    setExpanded(null);
  };

  // 显示/隐藏由外层 Presence 控制（保留挂载以维持任务状态），这里不再自行 display:none
  return (
    <div className={`history-overlay${busy ? " overlay-clickthrough" : ""}`} onClick={onClose}>
      <div className="history-panel batch-panel" onClick={(e) => e.stopPropagation()}>
        <PanelHeader icon={<IconLayers size={16} />} title="批量任务" onClose={onClose}
          closeTitle={busy ? "收起，任务继续执行" : "关闭"} closeLabel={busy ? "收起" : undefined} />

        <p className="memory-hint">
          对多个附加文件执行同一条指令（如「统一格式后导入」「按地区汇总」）。每个文件独立跑一遍 AI，结果写入当前工作簿。
        </p>

        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls,.xlsm,.csv"
          multiple
          style={{ display: "none" }}
          onChange={(e) => {
            void addFiles(e.target.files);
            e.target.value = "";
          }}
        />

        <div className="batch-files">
          <div className="batch-files-head">
            <span>附加文件 · {extFiles.length}</span>
            <button className="btn tiny secondary" disabled={busy} onClick={() => fileInputRef.current?.click()}>
              <IconPlus size={12} /> 添加文件
            </button>
          </div>
          {extFiles.length === 0 && (
            <EmptyState icon={<IconFile size={20} />} title="还没有附加文件" hint="点「添加文件」，或直接把 Excel / CSV 拖进对话窗口。" />
          )}
          {extFiles.map((f) => (
            <label key={f.name} className={`batch-file ${selected.has(f.name) ? "checked" : ""}`}>
              <input
                type="checkbox"
                checked={selected.has(f.name)}
                disabled={busy}
                onChange={() => toggleFile(f.name)}
              />
              <span className="file-chip-alias" title={f.caption}>{f.alias}</span>
              <span className="batch-file-name" title={f.sheets.map((s) => `${s.name} (${s.rows.length}行)`).join("\n")}>
                <IconFile size={12} /> {f.name}
              </span>
              <span className="batch-file-meta">{f.sheets.length} 表 · {f.sheets.reduce((s, x) => s + x.rows.length, 0)} 行</span>
            </label>
          ))}
        </div>

        <div className="batch-instruction">
          <textarea
            value={instruction}
            rows={3}
            disabled={busy}
            placeholder="对每个文件执行的指令，例如：清洗数据（去重、去空格、统一日期格式），把结果导入为新 Sheet"
            onChange={(e) => setInstruction(e.target.value)}
          />
        </div>

        {error && <div className="memory-error">{error}</div>}

        {phase === "idle" && (
          <div className="batch-start-row">
            <button
              className="btn primary"
              disabled={mainRunning || busy || selectedNames.length === 0 || !instruction.trim()}
              onClick={() => void start()}
              title={mainRunning ? "主对话正在执行任务" : undefined}
            >
              开始批量处理{selectedNames.length > 0 ? `（${selectedNames.length} 个文件）` : ""}
            </button>
          </div>
        )}

        {phase !== "idle" && (
          <div className="batch-run">
            <div className="batch-run-head">
              <span>{busy ? "执行中…需要确认时会自动展开主对话" : "执行结束"}</span>
              {busy ? (
                <button
                  className="btn tiny danger"
                  onClick={() => {
                    abortRef.current?.abort();
                    onAbortCleanup();
                  }}
                >
                  停止
                </button>
              ) : (
                <button className="btn tiny secondary" onClick={reset}>
                  再来一批
                </button>
              )}
            </div>
            {states.map((s) => (
              <div key={s.fileName} className={`batch-run-item ${s.status}`}>
                <button type="button"
                  className="batch-run-head-row disclosure-button" aria-expanded={expanded === s.fileName} aria-controls={`batch-trace-${s.fileName}`}
                  onClick={() => setExpanded(expanded === s.fileName ? null : s.fileName)}
                >
                  <span className={`batch-status ${s.status}`}>{STATUS_ICONS[s.status]}</span>
                  <span className="batch-file-name">
                    {extFiles.find((f) => f.name === s.fileName)?.caption ?? s.fileName}
                  </span>
                  {s.toolTrace.length > 0 && (
                    <span className={`batch-trace-toggle${expanded === s.fileName ? " open" : ""}`}><IconChevronRight size={11} /> {s.toolTrace.length} 步</span>
                  )}
                </button>
                {s.summary && <div className="batch-run-summary">{s.summary}</div>}
                {s.error && <div className="batch-run-error">{s.error}</div>}
                <Collapse open={expanded === s.fileName && s.toolTrace.length > 0} id={`batch-trace-${s.fileName}`}>
                  <div className="batch-trace">
                    {s.toolTrace.map((t, i) => (
                      <div key={i} className={`batch-trace-step ${t.status}`}>
                        <span className="batch-trace-icon">{t.status === "running" ? <span className="spinner" /> : t.status === "ok" ? <IconCheck size={12} /> : <IconClose size={11} />}</span>
                        <span className="batch-trace-label">{t.label}</span>
                        <span className="batch-trace-args" title={t.argsSummary}>{t.argsSummary}</span>
                      </div>
                    ))}
                  </div>
                </Collapse>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
