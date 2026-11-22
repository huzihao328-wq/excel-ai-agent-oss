import { useEffect, useRef, useState } from "react";
import Collapse from "./Collapse";
import { IconClose, IconServer, IconShield, IconSliders } from "./Icons";
import { AppSettings } from "../../core/types";
import { endpointOrigin, isLoopbackEndpoint, modelEndpointError, PROVIDER_PRESETS } from "../../core/settings";
import { listModels, probeContextWindowInfo, testConnection } from "../../core/llm";
import { mergeModelLists } from "../../core/modelList";
import { explainRunError } from "../../core/errorExplain";
import {
  formatContextWindow,
  getCachedContextWindow,
  parseContextWindowInput,
  ContextWindowInfo,
  getCachedContextInfo,
  describeContextWindow,
} from "../../core/contextWindow";

interface Props {
  settings: AppSettings;
  onSave: (s: AppSettings) => void;
  onClose: () => void;
}

type Status = { kind: "idle" | "busy" | "ok" | "fail"; message: string };

export default function SettingsPanel({ settings, onSave, onClose }: Props) {
  const [draft, setDraft] = useState<AppSettings>({ ...settings });
  const [showKey, setShowKey] = useState(false);
  const [models, setModels] = useState<string[]>([]);
  const [modelFilter, setModelFilter] = useState("");
  const [modelListOpen, setModelListOpen] = useState(false);
  const [status, setStatus] = useState<Status>({ kind: "idle", message: "" });
  const [apiKeyOrigin, setApiKeyOrigin] = useState(settings.apiKey.trim() ? endpointOrigin(settings.baseUrl) : "");
  const [windowText, setWindowText] = useState(settings.contextWindowTokens ? formatContextWindow(settings.contextWindowTokens) : "");
  const [contextInfo, setContextInfo] = useState<ContextWindowInfo | undefined>(() => getCachedContextInfo(settings.baseUrl, settings.model, settings.apiKey));
  const windowFocusRef = useRef(false);
  const requestRef = useRef(0);
  const requestBusyRef = useRef(false);

  useEffect(() => () => { requestRef.current++; }, []);

  const invalidateRequest = () => {
    requestRef.current++;
    requestBusyRef.current = false;
    setStatus({ kind: "idle", message: "" });
  };

  useEffect(() => {
    if (windowFocusRef.current) return;
    setWindowText(draft.contextWindowTokens ? formatContextWindow(draft.contextWindowTokens) : "");
  }, [draft.contextWindowTokens]);

  const set = <K extends keyof AppSettings>(key: K, value: AppSettings[K]) => {
    if (key === "model" || key === "apiKey" || key === "contextWindowTokens") invalidateRequest();
    if (key === "model" || key === "apiKey") setContextInfo(undefined);
    setDraft((d) => ({
      ...d, [key]: value,
      ...(key === "model" || key === "apiKey" ? { contextWindowTokens: 0 } : {}),
    }));
  };

  const applyPreset = (label: string) => {
    const p = PROVIDER_PRESETS.find((x) => x.label === label);
    if (!p) return;
    invalidateRequest();
    setContextInfo(undefined);
    const nextOrigin = endpointOrigin(p.baseUrl);
    const clearKey = !!draft.apiKey.trim() && !!apiKeyOrigin && !!nextOrigin && nextOrigin !== apiKeyOrigin;
    setDraft((d) => ({
      ...d,
      baseUrl: p.baseUrl,
      model: p.model,
      contextWindowTokens: 0,
      ...(clearKey ? { apiKey: "" } : {}),
    }));
    if (clearKey) setApiKeyOrigin("");
    setModels([]);
    setModelListOpen(false);
    setStatus(clearKey
      ? { kind: "ok", message: "API 地址已切换，旧地址的 API Key 已清除，请按需重新填写。" }
      : { kind: "idle", message: "" });
  };

  const changeBaseUrl = (value: string) => {
    invalidateRequest();
    setContextInfo(undefined);
    const nextOrigin = endpointOrigin(value);
    const clearKey = !!draft.apiKey.trim() && !!apiKeyOrigin && !!nextOrigin && nextOrigin !== apiKeyOrigin;
    setDraft((d) => ({ ...d, baseUrl: value, contextWindowTokens: 0, ...(clearKey ? { apiKey: "" } : {}) }));
    if (clearKey) {
      setApiKeyOrigin("");
      setStatus({ kind: "ok", message: "API 主机已改变，旧主机的 API Key 已自动清除。" });
    }
    setModels([]);
    setModelListOpen(false);
  };

  /** 连接类报错说成人话；服务商原话放在括号里备查 */
  const failureText = (e: unknown, listing = false): string => {
    const raw = String((e as Error)?.message || e);
    if (listing && /HTTP 404/.test(raw)) return "这个接口不提供模型列表，直接手动填写模型名即可。";
    const explained = explainRunError(e, { loopback: isLoopbackEndpoint(draft.baseUrl) });
    const prefix = listing && !/^获取模型列表/.test(explained.text) ? "获取模型列表失败：" : "";
    return `${prefix}${explained.text}${explained.detail ? `（${explained.detail}）` : ""}`;
  };

  const detectModels = async () => {
    if (requestBusyRef.current) return;
    requestBusyRef.current = true;
    const request = ++requestRef.current;
    setStatus({ kind: "busy", message: "正在获取模型列表…" });
    try {
      const available = await listModels(draft);
      if (request !== requestRef.current) return;
      const list = mergeModelLists(available, [draft.model]);
      setModels(list);
      setModelFilter("");
      setModelListOpen(true);
      const ctx = await probeContextWindowInfo(draft).catch(() => null);
      if (request !== requestRef.current) return;
      setContextInfo(ctx || undefined);
      const windowHint = ctx ? `，当前模型上下文 ${describeContextWindow(ctx)}` : "";
      setStatus({
        kind: "ok",
        message: available.length
          ? `探测到 ${available.length} 个可用模型${windowHint}，点击选择`
          : `接口正常，但未返回模型列表${windowHint}`,
      });
    } catch (e: any) {
      if (request !== requestRef.current) return;
      setModels(mergeModelLists([draft.model]));
      setStatus({ kind: "fail", message: failureText(e, true) });
    } finally {
      if (request === requestRef.current) requestBusyRef.current = false;
    }
  };

  const runTest = async () => {
    if (requestBusyRef.current) return;
    requestBusyRef.current = true;
    const request = ++requestRef.current;
    setStatus({ kind: "busy", message: "正在测试对话、工具调用往返和上下文窗口…" });
    try {
      const result = await testConnection(draft);
      if (request !== requestRef.current) return;
      setContextInfo(result.contextInfo);
      setStatus({ kind: result.toolCalling?.status === "unverified" ? "fail" : "ok", message: result.message });
    } catch (e: any) {
      if (request !== requestRef.current) return;
      setStatus({ kind: "fail", message: failureText(e) });
    } finally {
      if (request === requestRef.current) requestBusyRef.current = false;
    }
  };

  const probeWindow = async () => {
    if (requestBusyRef.current) return;
    requestBusyRef.current = true;
    const request = ++requestRef.current;
    setStatus({ kind: "busy", message: "正在探测上下文窗口…" });
    try {
      const ctx = await probeContextWindowInfo(draft, { force: true });
      if (request !== requestRef.current) return;
      if (!ctx) {
        setContextInfo(undefined);
        setStatus({ kind: "fail", message: "接口未返回上下文窗口，可按模型名手动填写，如 128k" });
        return;
      }
      setContextInfo(ctx);
      setStatus({ kind: "ok", message: `上下文窗口 ${describeContextWindow(ctx)}。${draft.contextWindowTokens ? "当前仍优先使用手动值；清空输入可恢复自动。" : "已用于自动预算。"}` });
    } catch (e: any) {
      if (request !== requestRef.current) return;
      setStatus({ kind: "fail", message: e?.message || String(e) });
    } finally {
      if (request === requestRef.current) requestBusyRef.current = false;
    }
  };

  const filteredModels = modelFilter
    ? models.filter((m) => m.toLowerCase().includes(modelFilter.toLowerCase()))
    : models;
  const loopback = isLoopbackEndpoint(draft.baseUrl);
  const endpointError = draft.baseUrl ? modelEndpointError(draft.baseUrl) : "请填写 API 地址。";
  const normalized = { ...draft, baseUrl: draft.baseUrl.trim(), apiKey: draft.apiKey.trim(), model: draft.model.trim() };
  // 测试连接成功后直接关掉面板是常见误操作：明确提示还没保存
  const dirty = JSON.stringify(normalized) !== JSON.stringify(settings);

  const toggles: { key: "confirmDestructive" | "sendSelectionContext" | "askDirection" | "useMemory" | "autoLearnSkills" | "autoCheckpoint"; title: string; desc: string; checked: boolean }[] = [
    { key: "confirmDestructive", title: "改动前先确认", desc: "覆盖、删除等破坏性操作先问我（推荐）", checked: draft.confirmDestructive },
    { key: "autoCheckpoint", title: "自动备份工作簿", desc: "动手改表前备份整个工作簿，可一键恢复到任务开始", checked: draft.autoCheckpoint !== false },
    { key: "askDirection", title: "先给方向再动手", desc: "任务有多种做法时，先列出选项让我选（推荐）", checked: draft.askDirection },
    { key: "sendSelectionContext", title: "附带当前选区", desc: "把选区内容作为上下文发给模型（≤200 单元格）", checked: draft.sendSelectionContext },
    { key: "useMemory", title: "长期记忆", desc: "跨对话记住习惯和本工作簿的约定", checked: draft.useMemory !== false },
    { key: "autoLearnSkills", title: "自我进化", desc: "任务完成后复盘沉淀技能，相似任务自动调用（少量额外 token）", checked: draft.autoLearnSkills !== false },
  ];
  const activePreset = PROVIDER_PRESETS.find((p) => p.baseUrl === draft.baseUrl);
  const save = () => {
    if (endpointError) {
      setStatus({ kind: "fail", message: endpointError });
      return;
    }
    try {
      onSave({ ...draft, baseUrl: draft.baseUrl.trim(), apiKey: draft.apiKey.trim(), model: draft.model.trim() });
      onClose();
    } catch {
      setStatus({ kind: "fail", message: "设置未保存，本机存储不可用或空间不足。请释放空间后重试。" });
    }
  };
  const footerMessage = status.message
    ? `${status.message}${status.kind === "ok" && dirty ? "（尚未保存，请点「保存」生效）" : ""}`
    : endpointError || (!draft.model.trim() ? "请填写模型名，或点击「获取模型」选择。" : dirty ? "有未保存的修改" : "");

  return (
    <div className="settings-overlay" onClick={onClose}>
      <div className="settings-panel settings-sheet" role="dialog" aria-modal="true" aria-label="模型设置" onClick={(e) => e.stopPropagation()}>
        <div className="settings-header">
          <h2>设置</h2>
          <button className="icon-btn" onClick={onClose} title="关闭" aria-label="关闭"><IconClose size={15} /></button>
        </div>

        <div className="settings-body">
          <section className="settings-group" aria-labelledby="settings-connection">
            <h3 className="settings-group-title" id="settings-connection"><IconServer size={13} /> 模型连接</h3>
            <div className="settings-card">
              <div className="settings-row-label">服务商{activePreset ? <span className="settings-row-value">{activePreset.label}</span> : <span className="settings-row-value muted">自定义</span>}</div>
              <div className="provider-grid" role="group" aria-label="服务商">
                {PROVIDER_PRESETS.map((p) => (
                  <button
                    key={p.label}
                    className={`provider-btn ${draft.baseUrl === p.baseUrl ? "active" : ""}`}
                    aria-pressed={draft.baseUrl === p.baseUrl}
                    onClick={() => applyPreset(p.label)}
                  >
                    {p.label}
                  </button>
                ))}
              </div>

              <label className="field">
                <span>API 地址<small>OpenAI 兼容</small></span>
                <input
                  value={draft.baseUrl}
                  onChange={(e) => changeBaseUrl(e.target.value)}
                  placeholder="https://api.deepseek.com/v1"
                  spellCheck={false}
                />
              </label>

              <label className="field">
                <span>API Key{loopback ? <small>本地模型可留空</small> : null}</span>
                <div className="input-with-btn">
                  <input
                    type={showKey ? "text" : "password"}
                    value={draft.apiKey}
                    onChange={(e) => {
                      const value = e.target.value;
                      set("apiKey", value);
                      setModels([]);
                      setModelListOpen(false);
                      setApiKeyOrigin(value.trim() ? endpointOrigin(draft.baseUrl) : "");
                    }}
                    placeholder={loopback ? "本地模型通常无需填写" : "sk-…"}
                    spellCheck={false}
                  />
                  <button className="inline-btn" onClick={() => setShowKey((v) => !v)} title={showKey ? "隐藏" : "显示"}>
                    {showKey ? "隐藏" : "显示"}
                  </button>
                </div>
              </label>

              <label className="field">
                <span>模型</span>
                <div className="input-with-btn">
                  <input
                    value={draft.model}
                    onChange={(e) => { set("model", e.target.value); setModelFilter(e.target.value); }}
                    onFocus={() => models.length && setModelListOpen(true)}
                    placeholder="deepseek-chat"
                    spellCheck={false}
                  />
                  <button
                    className="inline-btn accent"
                    onClick={detectModels}
                    disabled={status.kind === "busy" || !!endpointError}
                    title="从 API 获取可用模型列表"
                  >
                    {status.kind === "busy" ? "…" : "获取模型"}
                  </button>
                </div>
              </label>

              <Collapse open={modelListOpen && models.length > 0}>
                <div className="model-list">
                  <div className="model-list-head">
                    <span>{filteredModels.length} / {models.length} 个模型（输入可筛选）</span>
                    <button className="icon-btn" onClick={() => setModelListOpen(false)}>收起</button>
                  </div>
                  <div className="model-list-body">
                    {filteredModels.slice(0, 200).map((m) => (
                      <button
                        key={m}
                        className={`model-item ${draft.model === m ? "active" : ""}`}
                        onClick={() => {
                          set("model", m);
                          setModelListOpen(false);
                          setModelFilter("");
                        }}
                      >
                        <span>{m}</span>
                        {!!getCachedContextWindow(draft.baseUrl, m, draft.apiKey) && (
                          <span className="model-ctx">{formatContextWindow(getCachedContextWindow(draft.baseUrl, m, draft.apiKey))}</span>
                        )}
                      </button>
                    ))}
                    {filteredModels.length === 0 && <div className="model-empty">没有匹配的模型</div>}
                  </div>
                </div>
              </Collapse>

              <label className="field">
                <span>上下文窗口{draft.contextWindowTokens ? <small>{draft.contextWindowTokens.toLocaleString()} tokens</small> : null}</span>
                <div className="input-with-btn">
                  <input
                    value={windowText}
                    onFocus={() => {
                      windowFocusRef.current = true;
                    }}
                    onBlur={() => {
                      windowFocusRef.current = false;
                      setWindowText(draft.contextWindowTokens ? formatContextWindow(draft.contextWindowTokens) : "");
                    }}
                    onChange={(e) => {
                      setWindowText(e.target.value);
                      set("contextWindowTokens", parseContextWindowInput(e.target.value));
                    }}
                    placeholder="接入后自动探测，也可填 128k"
                    spellCheck={false}
                  />
                  <button
                    className="inline-btn accent"
                    onClick={() => void probeWindow()}
                    disabled={status.kind === "busy" || !!endpointError || !draft.model.trim()}
                    title="从当前 API 探测该模型的上下文窗口"
                  >
                    {status.kind === "busy" ? "…" : "探测"}
                  </button>
                </div>
              </label>
              <p className="settings-note">
                {draft.contextWindowTokens ? "优先使用手动窗口。" : "窗口留空时自动探测。"}
                {contextInfo ? `自动值：${describeContextWindow(contextInfo)}。` : "接口未提供时只能保守估算，不能保证实际可用上限。"}
                测试连接会额外执行虚拟工具往返请求，不读取工作簿。
              </p>
              <p className="settings-note">
                {loopback
                  ? "本机回环地址：API Key 可留空，请求只发往本机模型服务。"
                  : "API Key 只保存在本机；切换 API 主机时会自动清除，防止误发给其他服务。"}
              </p>
            </div>
          </section>

          <section className="settings-group" aria-labelledby="settings-generation">
            <h3 className="settings-group-title" id="settings-generation"><IconSliders size={13} /> 生成</h3>
            <div className="settings-card list">
              <label className="settings-row">
                <span className="settings-row-text">
                  <span className="settings-row-title">温度</span>
                  <span className="settings-row-desc">越低越稳定，处理表格建议 0–0.3</span>
                </span>
                <span className="range-control">
                  <input
                    type="range" min={0} max={1} step={0.1}
                    value={draft.temperature}
                    style={{ "--range-fill": `${draft.temperature * 100}%` } as React.CSSProperties}
                    onChange={(e) => set("temperature", Number(e.target.value))}
                  />
                  <output>{draft.temperature.toFixed(1)}</output>
                </span>
              </label>
              <label className="settings-row">
                <span className="settings-row-text">
                  <span className="settings-row-title">最大工具轮数</span>
                  <span className="settings-row-desc">单个任务最多调用工具的轮数</span>
                </span>
                <input
                  className="number-input"
                  type="number" min={3} max={50}
                  value={draft.maxToolRounds}
                  onChange={(e) => set("maxToolRounds", Math.round(Math.min(50, Math.max(3, Number(e.target.value) || 20))))}
                />
              </label>
              <label className="settings-row" title="单次回复的 max_tokens。填 0 则不传，由服务端决定；模型经常长篇大论时调小">
                <span className="settings-row-text">
                  <span className="settings-row-title">单次回复上限</span>
                  <span className="settings-row-desc">max_tokens，0 表示由服务端决定</span>
                </span>
                <input
                  className="number-input"
                  type="number" min={0} max={128000} step={256}
                  value={draft.maxOutputTokens ?? 4096}
                  onChange={(e) => set("maxOutputTokens", Math.min(128000, Math.max(0, Math.round(Number(e.target.value) || 0))))}
                />
              </label>
            </div>
          </section>

          <section className="settings-group" aria-labelledby="settings-behavior">
            <h3 className="settings-group-title" id="settings-behavior"><IconShield size={13} /> 安全与习惯</h3>
            <div className="settings-card list">
              {toggles.map((toggle) => (
                <label key={toggle.key} className="settings-row">
                  <span className="settings-row-text">
                    <span className="settings-row-title">{toggle.title}</span>
                    <span className="settings-row-desc">{toggle.desc}</span>
                  </span>
                  <input
                    type="checkbox"
                    role="switch"
                    className="switch"
                    checked={toggle.checked}
                    onChange={(e) => set(toggle.key, e.target.checked)}
                  />
                </label>
              ))}
            </div>
          </section>
        </div>

        <div className="settings-footer">
          <div className={`test-result ${footerMessage && status.kind === "idle" ? "hint" : status.kind}`} role="status">{footerMessage}</div>
          <div className="settings-actions">
            <button
              className="btn secondary"
              onClick={runTest}
              disabled={status.kind === "busy" || !!endpointError || !draft.model.trim() || (!draft.apiKey.trim() && !loopback)}
            >
              测试连接
            </button>
            <button className="btn primary" onClick={save} disabled={!!endpointError || !draft.model.trim()}>
              保存
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
