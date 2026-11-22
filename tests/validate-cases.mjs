/**
 * 核对 cases.json 里的标准数字是否仍与 ground-truth.json 一致。
 * 不启动 Excel，只做 JSON 对照。
 *
 *   node tests/validate-cases.mjs
 */
import { readFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";
import ts from "typescript";

const ROOT = dirname(fileURLToPath(import.meta.url));
const cases = JSON.parse(readFileSync(join(ROOT, "cases.json"), "utf8"));
const truth = JSON.parse(readFileSync(join(ROOT, "fixtures", "ground-truth.json"), "utf8"));

const errors = [];
const checked = [];

// 安全不变量：模型只能调用结构化、可审计的专用工具，不能恢复任意代码执行通道。
const sourceRoot = join(ROOT, "..", "src");
const definitionsSource = readFileSync(join(sourceRoot, "core", "tools", "definitions.ts"), "utf8");
const executorSource = readFileSync(join(sourceRoot, "core", "tools", "excelTools.ts"), "utf8");
const promptSource = readFileSync(join(sourceRoot, "core", "prompts.ts"), "utf8");
const securityCorpus = `${definitionsSource}\n${executorSource}\n${promptSource}`;
if (/run_excel_js|\bnew\s+Function\s*\(/.test(securityCorpus)) {
  errors.push("安全不变量失败：源码重新暴露了任意 JavaScript 执行通道");
} else {
  checked.push("security: no arbitrary JavaScript execution tool");
}
if (!/不要把大段原始数据/.test(promptSource) || !/不要输出思考过程/.test(promptSource)) {
  errors.push("提示词失败：回复风格没有限制长篇输出或思考链");
} else {
  checked.push("prompt: concise replies, hide chain-of-thought");
}

const safeValuesSource = readFileSync(join(sourceRoot, "core", "tools", "safeValues.ts"), "utf8");
const safeValuesJs = ts.transpileModule(safeValuesSource, {
  compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ES2020 },
}).outputText;
const safeValuesModule = await import(`data:text/javascript;base64,${Buffer.from(safeValuesJs).toString("base64")}`);
const protectedGrid = safeValuesModule.protectFormulaLikeTextGrid([
  ["=1+1", "+SUM(A1:A2)", "-2+3", "@SUM(A1:A2)"],
  ["normal", 42, true, ""],
]);
const expectedProtectedGrid = [
  ["'=1+1", "'+SUM(A1:A2)", "'-2+3", "'@SUM(A1:A2)"],
  ["normal", 42, true, ""],
];
if (JSON.stringify(protectedGrid) !== JSON.stringify(expectedProtectedGrid)) {
  errors.push("安全不变量失败：公式样式文本没有被完整转义");
} else {
  checked.push("security: formula-like text is preserved as text");
}
const restorableGrid = safeValuesModule.buildRestorableGrid(
  [[2, "=literal", "=formula result"]],
  [["=1+1", "=literal", '="=formula result"']]
);
const expectedRestorableGrid = [["=1+1", "'=literal", '="=formula result"']];
if (JSON.stringify(restorableGrid) !== JSON.stringify(expectedRestorableGrid)) {
  errors.push("撤销不变量失败：公式与公式样式文本不能安全区分");
} else {
  checked.push("undo: formulas and formula-like text restore safely");
}
const formulaRecalculationIsSafe = safeValuesModule.cellStatesEqual(
  [[999, "=literal"]],
  [["=SUM(A1:A2)", "=literal"]],
  [[3, "=literal"]],
  [["=SUM(A1:A2)", "=literal"]]
);
const literalEditIsConflict = !safeValuesModule.cellStatesEqual(
  [[3, "changed"]],
  [["=SUM(A1:A2)", "changed"]],
  [[3, "=literal"]],
  [["=SUM(A1:A2)", "=literal"]]
);
if (!formulaRecalculationIsSafe || !literalEditIsConflict) {
  errors.push("撤销不变量失败：公式重算与人工编辑冲突不能正确区分");
} else {
  checked.push("undo: recalculation allowed and later edits detected");
}
const emptyCellsEqual = safeValuesModule.cellStatesEqual(
  [[null, ""]],
  [["", ""]],
  [["", null]],
  [["", ""]]
);
const uncalculatedFormulaIsSafe = safeValuesModule.cellStatesEqual(
  [[42]],
  [["=A1+1"]],
  [["=A1+1"]],
  [["=A1+1"]]
);
const generalFormatsEqual = safeValuesModule.numberFormatGridsEqual([["General"]], [["G/通用格式"]]);
if (!emptyCellsEqual || !uncalculatedFormulaIsSafe || !generalFormatsEqual) {
  errors.push("撤销不变量失败：空单元格、未完成重算或通用格式比较过严");
} else {
  checked.push("undo: empty cells, uncalculated formulas, general formats");
}
if (!/async function applyUndoTo/.test(executorSource) || !/slice\(idx\)[\s\S]{0,80}\.reverse\(\)/.test(executorSource)) {
  errors.push("撤销不变量失败：撤回到节点必须按从新到旧顺序恢复");
} else {
  checked.push("undo: rollback-to-node undoes newest first");
}

const transpileOpts = { compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ES2020 } };
const toDataUrl = (js) => `data:text/javascript;base64,${Buffer.from(js).toString("base64")}`;
const contextWindowJs = ts.transpileModule(
  readFileSync(join(sourceRoot, "core", "contextWindow.ts"), "utf8"),
  transpileOpts
).outputText;
const contextWindowUrl = toDataUrl(contextWindowJs);
const modelListSource = readFileSync(join(sourceRoot, "core", "modelList.ts"), "utf8").replace(
  /from ["']\.\/contextWindow["']/,
  `from "${contextWindowUrl}"`
);
const modelListJs = ts.transpileModule(modelListSource, transpileOpts).outputText;
const modelListModule = await import(toDataUrl(modelListJs));
const contextWindowModule = await import(contextWindowUrl);
const extracted = modelListModule.extractModelIds({
  object: "list",
  data: [{ id: "deepseek-chat" }, { name: "moonshot-v1-32k" }],
  models: [{ model: "qwen-plus" }],
});
const nested = modelListModule.extractModelIds({ data: { data: ["glm-4-air", { id: "glm-4-air" }] } });
const merged = modelListModule.mergeModelLists(["b"], ["a", "b"], ["current"]);
const cataloged = modelListModule.extractModelCatalog({
  data: [{ id: "anthropic/claude-3.5-sonnet", context_length: 200000 }],
});
const guessed = contextWindowModule.guessContextWindow("deepseek-chat");
const formatted = contextWindowModule.formatContextWindow(131072);
const modelListOk =
  extracted.join(",") === "deepseek-chat,moonshot-v1-32k,qwen-plus" &&
  nested.join(",") === "glm-4-air" &&
  merged.join(",") === "b,a,current" &&
  cataloged[0]?.contextWindow === 200000 &&
  guessed === 131072 &&
  formatted === "128K";
if (!modelListOk) {
  errors.push("模型列表解析失败：未能识别常见 /models 返回格式");
} else {
  checked.push("models: extract nested OpenAI-compatible ids");
  checked.push("models: detect context window from API payload and model name");
}

const streamTextSource = readFileSync(join(sourceRoot, "core", "streamText.ts"), "utf8");
const streamTextJs = ts.transpileModule(streamTextSource, {
  compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ES2020 },
}).outputText;
const streamTextModule = await import(`data:text/javascript;base64,${Buffer.from(streamTextJs).toString("base64")}`);
const incremental = ["已", "把", "数据", "写入。"].reduce((acc, piece) => streamTextModule.appendStreamText(acc, piece), "");
const snapshot = ["已完成。", "已完成。请查看 Sheet1。"].reduce((acc, piece) => streamTextModule.appendStreamText(acc, piece), "");
const replayed = ["已把汇总写入新表。", "已把汇总写入新表。"].reduce((acc, piece) => streamTextModule.appendStreamText(acc, piece), "");
const secondShu = streamTextModule.appendStreamText("数据处理完成，", "数");
const doubledSentence = streamTextModule.collapseDuplicatedText("请稍候。请稍候。");
const prefixCharKept = secondShu === "数据处理完成，数";
const streamTextOk =
  incremental === "已把数据写入。" &&
  snapshot === "已完成。请查看 Sheet1。" &&
  replayed === "已把汇总写入新表。" &&
  prefixCharKept &&
  doubledSentence === "请稍候。";
if (!streamTextOk) {
  errors.push("流式去重失败：增量/快照/重复句子处理不正确");
} else {
  checked.push("stream: deltas, snapshots, and repeated sentences");
}
const snapshotName = ["write", "write_range", "write_range"].reduce(
  (acc, piece) => streamTextModule.appendStreamFragment(acc, piece),
  ""
);
const incrementalArgs = ['{"sheet":"', "订单", '"}'].reduce(
  (acc, piece) => streamTextModule.appendStreamFragment(acc, piece),
  ""
);
const clipped = streamTextModule.clipJsonToBudget({ ok: true, data: "x".repeat(4000) }, 200);
let clippedJson = null;
try {
  clippedJson = JSON.parse(clipped);
} catch {
  clippedJson = null;
}
const streamFragOk =
  snapshotName === "write_range" &&
  incrementalArgs === '{"sheet":"订单"}' &&
  clippedJson &&
  clippedJson.truncated === true &&
  clippedJson.ok === true &&
  clipped.length <= 200;
if (!streamFragOk) {
  errors.push("流式工具参数合并或过长结果截断不正确");
} else {
  checked.push("stream: tool-call snapshots and valid truncated JSON");
}
const thinkHidden = streamTextModule.stripReasoningText("<think>先看表头再决定</think>\n已写入 Sheet1。") === "已写入 Sheet1。";
const thinkOpenHidden = streamTextModule.stripReasoningText("<think>\n还在想") === "";
const thinkKept = streamTextModule.stripReasoningText("I think we should use XLOOKUP.") === "I think we should use XLOOKUP.";
let thinkStreamRaw = "";
for (const piece of ["<think>\n", "plan\n", "</think>\n", "已写入 Sheet1。"]) {
  thinkStreamRaw = streamTextModule.appendStreamText(thinkStreamRaw, piece);
}
const thinkStreamVisible = streamTextModule.stripReasoningText(thinkStreamRaw) === "已写入 Sheet1。";
const skipReasoningField =
  streamTextModule.streamPieceFromChoice({
    delta: { reasoning_content: "secret", content: "答案" },
  }) === "答案";
const skipThinkingPart =
  streamTextModule.textFromContent([
    { type: "thinking", text: "nope" },
    { type: "text", text: "可见" },
  ]) === "可见";
if (!thinkHidden || !thinkOpenHidden || !thinkKept || !thinkStreamVisible || !skipReasoningField || !skipThinkingPart) {
  errors.push("思考链过滤失败：<think> / reasoning 仍可能进入用户可见回复");
} else {
  checked.push("stream: reasoning and <think> stripped from replies");
}

const settingsSource = readFileSync(join(sourceRoot, "core", "settings.ts"), "utf8").replace(
  /import \{[^}]+\} from ["']\.\/durableStore["'];\r?\n/,
  "const SETTINGS_STORAGE_KEY = \"excel-ai-agent-settings\";\nfunction schedulePersistUserData() {}\n"
);
const settingsJs = ts.transpileModule(settingsSource, {
  compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ES2020 },
}).outputText;
const settingsModule = await import(`data:text/javascript;base64,${Buffer.from(settingsJs).toString("base64")}`);
const endpointPolicyOk =
  settingsModule.modelEndpointError("https://api.example.com/v1") === null &&
  settingsModule.modelEndpointError("http://localhost:11434/v1") === null &&
  settingsModule.modelEndpointError("http://127.0.0.1:1234/v1") === null &&
  !!settingsModule.modelEndpointError("http://api.example.com/v1") &&
  !!settingsModule.modelEndpointError("https://user:pass@example.com/v1");
if (!endpointPolicyOk) {
  errors.push("网络不变量失败：远程 HTTP 或 URL 内嵌凭据未被阻止");
} else {
  checked.push("network: remote HTTPS enforced and loopback HTTP allowed");
}

const rangeReadSource = readFileSync(join(sourceRoot, "core", "tools", "rangeRead.ts"), "utf8");
const rangeReadJs = ts.transpileModule(rangeReadSource, {
  compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ES2020 },
}).outputText;
const rangeReadModule = await import(`data:text/javascript;base64,${Buffer.from(rangeReadJs).toString("base64")}`);
const fullPlan = rangeReadModule.planRangeRead(20, 8);
const pagePlan = rangeReadModule.planRangeRead(2000, 10, { start_row: 31, max_rows: 40 });
const samplePlan = rangeReadModule.planRangeRead(8000, 12);
const rangeReadOk =
  fullPlan.mode === "full" &&
  fullPlan.rowCount === 20 &&
  pagePlan.mode === "page" &&
  pagePlan.startOffset === 30 &&
  pagePlan.rowCount === 40 &&
  samplePlan.mode === "sample" &&
  samplePlan.headRows === 30 &&
  samplePlan.tailRows === 5 &&
  samplePlan.totalRows === 8000;
if (!rangeReadOk) {
  errors.push("读取计划失败：小表整读 / 分页 / 大表抽样规则不正确");
} else {
  checked.push("read: small full, paging, large sample");
}

const definitionsJs = ts.transpileModule(definitionsSource, {
  compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ES2020 },
}).outputText;
const definitionsModule = await import(`data:text/javascript;base64,${Buffer.from(definitionsJs).toString("base64")}`);
const toolNames = new Set(definitionsModule.TOOL_DEFINITIONS.map((t) => t.function.name));
const labelNames = new Set(Object.keys(definitionsModule.TOOL_LABELS));
const missingLabels = [...toolNames].filter((n) => !labelNames.has(n));
const extraLabels = [...labelNames].filter((n) => !toolNames.has(n));
const requiredPowerTools = [
  "copy_range",
  "fill_down",
  "find_replace",
  "remove_duplicates",
  "rename_sheet",
  "copy_sheet",
  "column_stats",
  "apply_conditional_format",
  "export_file",
  "add_lookup_column",
  "split_column",
  "combine_columns",
  "freeze_panes",
  "clean_column",
  "convert_to_values",
  "polish_range",
  "delete_blank_rows",
  "group_summary",
  "reconcile_tables",
  "stack_sheets",
  "split_by_column",
  "unpivot_columns",
  "fill_blank_cells",
  "extract_text",
  "compare_sheets",
  "set_page_layout",
  "create_sheet_index",
];
const missingPower = requiredPowerTools.filter((n) => !toolNames.has(n));
if (missingLabels.length || extraLabels.length) {
  errors.push(`工具标签不一致：缺 ${missingLabels.join(",") || "无"} / 多 ${extraLabels.join(",") || "无"}`);
} else {
  checked.push(`tools: ${toolNames.size} labeled definitions`);
}
if (missingPower.length) {
  errors.push(`能力缺口：缺少工具 ${missingPower.join(", ")}`);
} else {
  checked.push("tools: native copy/fill/replace/dedupe/stats/lookup/split present");
}

const exportFileSource = readFileSync(join(sourceRoot, "core", "exportFile.ts"), "utf8");
const exportFileJs = ts.transpileModule(exportFileSource, {
  compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ES2020 },
}).outputText;
const exportFileModule = await import(`data:text/javascript;base64,${Buffer.from(exportFileJs).toString("base64")}`);
const csv = exportFileModule.gridToCsv([
  ["名称", "备注"],
  ["苹果", "含,逗号"],
  ["a\"b", "换\n行"],
]);
const csvOk =
  csv.includes("名称,备注") &&
  csv.includes('"含,逗号"') &&
  csv.includes('"a""b"') &&
  exportFileModule.ensureExtension("汇总", "xlsx") === "汇总.xlsx" &&
  exportFileModule.sanitizeDownloadName("a:b/c.xlsx", "x") === "a_b_c.xlsx";
if (!csvOk) {
  errors.push("导出不变量失败：CSV 转义或文件名清理不正确");
} else {
  checked.push("export: csv escaping and download names");
}

const formulaRefsSource = readFileSync(join(sourceRoot, "core", "tools", "formulaRefs.ts"), "utf8");
const formulaRefsJs = ts.transpileModule(formulaRefsSource, {
  compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ES2020 },
}).outputText;
const formulaRefsModule = await import(`data:text/javascript;base64,${Buffer.from(formulaRefsJs).toString("base64")}`);
const lookupGrid = formulaRefsModule.buildLookupFormulaGrid({
  destKeyLetter: "C",
  destFirstDataRow: 2,
  destLastDataRow: 2,
  lookupSheet: "客户",
  lookupKeyLetter: "A",
  lookupFirstDataRow: 2,
  lookupLastDataRow: 6,
  returns: [
    { letter: "B", header: "客户名称" },
    { letter: "D", header: "地区" },
  ],
  engine: "xlookup",
});
let vlookupLeftFailed = false;
try {
  formulaRefsModule.buildLookupFormulaGrid({
    destKeyLetter: "C",
    destFirstDataRow: 2,
    destLastDataRow: 2,
    lookupSheet: "客户",
    lookupKeyLetter: "C",
    lookupFirstDataRow: 2,
    lookupLastDataRow: 6,
    returns: [{ letter: "A", header: "编号" }],
    engine: "vlookup",
  });
} catch {
  vlookupLeftFailed = true;
}
const formulaRefsOk =
  formulaRefsModule.quoteSheetName("订单") === "'订单'" &&
  formulaRefsModule.quoteSheetName("Sheet1") === "Sheet1" &&
  formulaRefsModule.quoteSheetName("A1") === "'A1'" &&
  formulaRefsModule.quoteSheetName("Q1") === "'Q1'" &&
  formulaRefsModule.quoteSheetName("R1C1") === "'R1C1'" &&
  formulaRefsModule.quoteSheetName("O'Brien") === "'O''Brien'" &&
  formulaRefsModule.columnA1Range("客户", "A", 2, 21) === "'客户'!$A$2:$A$21" &&
  JSON.stringify(formulaRefsModule.splitCellText("zhang@example.com", "@")) === JSON.stringify(["zhang", "example.com"]) &&
  formulaRefsModule.resolveDelimiter("at") === "@" &&
  JSON.stringify(formulaRefsModule.padSplitParts(["a"], 3)) === JSON.stringify(["a", "", ""]) &&
  lookupGrid[0][0] === "客户名称" &&
  lookupGrid[1][0] === `=XLOOKUP(C2,'客户'!$A$2:$A$6,'客户'!$B$2:$B$6,"")` &&
  lookupGrid[1][1] === `=XLOOKUP(C2,'客户'!$A$2:$A$6,'客户'!$D$2:$D$6,"")` &&
  formulaRefsModule.textJoinFormula("-", ["A2", "B2"]) === `=TEXTJOIN("-",TRUE,A2,B2)` &&
  vlookupLeftFailed;
if (!formulaRefsOk) {
  errors.push("查找/分列公式生成不正确");
} else {
  checked.push("formulas: xlookup, sheet quoting, split/join");
}

const cleanValuesSource = readFileSync(join(sourceRoot, "core", "tools", "cleanValues.ts"), "utf8");
const cleanValuesJs = ts.transpileModule(cleanValuesSource, {
  compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ES2020 },
}).outputText;
const cleanValuesModule = await import(`data:text/javascript;base64,${Buffer.from(cleanValuesJs).toString("base64")}`);
const dateYear = cleanValuesModule.inferDefaultYear(["2024-03-05", "2024/3/8", "3月12日"], 2099);
const cleanedName = cleanValuesModule.applyCleanOps("  王强  ", ["trim"]);
const cleanedFee = cleanValuesModule.applyCleanOps("¥1,200", ["trim", "to_number"]);
const cleanedDate = cleanValuesModule.applyCleanOps("2024/3/8", ["to_date"]);
const monthDay = cleanValuesModule.parseLooseDate("3月12日", { defaultYear: 2024 });
const cleanedSerial = cleanValuesModule.applyCleanOps(44927, ["to_date"]);
const cleanOk =
  cleanValuesModule.parseLooseNumber("800元") === 800 &&
  cleanValuesModule.parseLooseNumber("1,200") === 1200 &&
  cleanedFee.value === 1200 &&
  cleanedName.value === "王强" &&
  cleanedDate.value === "2024-03-08" &&
  monthDay === "2024-03-15".replace("15", "12") &&
  dateYear === 2024 &&
  cleanedSerial.value === 44927 &&
  !cleanedSerial.changed &&
  cleanValuesModule.parseLooseDate(44927) === null &&
  cleanValuesModule.guessColumnNumberFormat("金额") === "#,##0.00" &&
  cleanValuesModule.isRowBlank(["", null, undefined]) &&
  !cleanValuesModule.isRowBlank(["", "x"]);
if (!cleanOk) {
  errors.push("列清洗解析不正确（trim / 金额 / 日期）");
} else {
  checked.push("clean: trim, loose numbers, loose dates, blank rows");
}

const auditSource = readFileSync(join(sourceRoot, "core", "audit.ts"), "utf8").replace(
  /import \{[^}]+\} from ["']\.\/types["'];\r?\n/,
  ""
);
const auditJs = ts.transpileModule(auditSource, transpileOpts).outputText;
const auditModule = await import(toDataUrl(auditJs));
const auditAbs = auditModule.auditData(
  ["日期", "金额"],
  [
    { row: 10, cells: ["", 100] },
    { row: 11, cells: ["x", 200] },
  ],
  { colStart: 2 }
);
const missing = auditAbs.issues.find((i) => i.type === "missing");
const auditOk =
  missing &&
  missing.detail?.column === "C" &&
  String(missing.sample).includes("C10");
if (!auditOk) {
  errors.push("数据体检列字母没有按工作表绝对列输出");
} else {
  checked.push("audit: column letters are sheet-absolute");
}

const conversationsSource = readFileSync(join(sourceRoot, "core", "conversations.ts"), "utf8");
const agentSource = readFileSync(join(sourceRoot, "core", "agent.ts"), "utf8");
if (/index\.currentId\s*=\s*sanitized\.id/.test(conversationsSource)) {
  errors.push("会话落盘不应把 currentId 改成正在保存的对话");
} else {
  checked.push("persist: saveConversation does not steal currentId");
}
if (!/skipUserPush/.test(agentSource) || !/clipJsonToBudget/.test(agentSource)) {
  errors.push("Agent 循环缺少续跑保护或合法 JSON 截断");
} else {
  checked.push("agent: resume without duplicating user turn; tool JSON stays valid");
}

const llmRetrySource = readFileSync(join(sourceRoot, "core", "llmRetry.ts"), "utf8");
const llmRetryJs = ts.transpileModule(llmRetrySource, {
  compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.ES2020 },
}).outputText;
const llmRetryModule = await import(`data:text/javascript;base64,${Buffer.from(llmRetryJs).toString("base64")}`);
const retryableOk =
  llmRetryModule.isRetryableLlmError(new Error("模型请求失败 (HTTP 429): rate")) &&
  llmRetryModule.isRetryableLlmError(new Error("Failed to fetch")) &&
  !llmRetryModule.isRetryableLlmError(new Error("模型请求失败 (HTTP 401): no")) &&
  !llmRetryModule.isRetryableLlmError(new Error("当前模型可能不支持图片。请在设置里换成支持视觉的模型")) &&
  !llmRetryModule.isRetryableLlmError({ name: "AbortError", message: "aborted" }) &&
  llmRetryModule.isAbortLike({ name: "AbortError", message: "aborted" }) &&
  llmRetryModule.isAbortLike(new Error("The user aborted a request.")) &&
  !llmRetryModule.isAbortLike(new Error("模型请求失败 (HTTP 502): upstream request aborted")) &&
  llmRetryModule.isRetryableLlmError(new Error("模型请求失败 (HTTP 502): upstream request aborted")) &&
  llmRetryModule.isRetryableLlmError(new Error("模型请求失败 (HTTP 408): timeout")) &&
  llmRetryModule.llmRetryDelayMs(1) === 800 &&
  llmRetryModule.llmRetryDelayMs(2) === 2000;
if (!retryableOk) {
  errors.push("模型重试判定不正确");
} else {
  checked.push("llm: retry 429/5xx/network, skip 401/abort/vision");
}
let retryCount = 0;
const retryAc = new AbortController();
const retried = await llmRetryModule.withLlmRetries(
  async () => {
    retryCount += 1;
    if (retryCount < 3) throw new Error("模型请求失败 (HTTP 503): busy");
    return "ok";
  },
  { signal: retryAc.signal, delaysMs: [0, 0] }
);
const startedAc = new AbortController();
let startedCalls = 0;
let startedBlocked = false;
try {
  await llmRetryModule.withLlmRetries(
    async () => {
      startedCalls += 1;
      throw new Error("模型请求失败 (HTTP 503): busy");
    },
    { signal: startedAc.signal, delaysMs: [0], alreadyStarted: () => true }
  );
} catch {
  startedBlocked = startedCalls === 1;
}
if (retried !== "ok" || retryCount !== 3 || !startedBlocked) {
  errors.push("模型重试循环不正确（应重试未开始的 503，且已开始输出后不再重试）");
} else {
  checked.push("llm: retries then succeeds; no retry after stream started");
}

// ---- 文本形式工具调用兜底 / 口头计划识别 ----
const toolCallTextJs = ts.transpileModule(
  readFileSync(join(sourceRoot, "core", "toolCallText.ts"), "utf8"),
  transpileOpts
).outputText;
const toolCallTextModule = await import(toDataUrl(toolCallTextJs));
const knownForText = ["get_workbook_overview", "read_range", "write_range", "sort_range"];
const textToolCases = [
  ['好的。\n<tool_call>\n{"name": "get_workbook_overview", "arguments": {}}\n</tool_call>', "get_workbook_overview", "好的。"],
  ['```json\n{"name": "read_range", "arguments": {"sheet": "Sheet1", "range": "A1:B2"}}\n```', "read_range", ""],
  ['{"name":"sort_range","sheet":"S","key_column":"金额"}', "sort_range", ""],
  ['<function=read_range>{"sheet":"A"}</function>', "read_range", ""],
  ["先看结构。\nget_workbook_overview()\n", "get_workbook_overview", "先看结构。"],
  ['{"tool_calls":[{"id":"x","type":"function","function":{"name":"read_range","arguments":"{\\"sheet\\":\\"A\\"}"}}]}', "read_range", ""],
  ['<tool_call>\n{"name": "read_range", "arguments": {"sheet": "S"}}', "read_range", ""],
  ["我已经把结果写到了 Sheet2。", null, "我已经把结果写到了 Sheet2。"],
  ['{"name":"不是工具","arguments":{}}', null, '{"name":"不是工具","arguments":{}}'],
];
const textToolFailures = textToolCases.filter(([input, name, rest]) => {
  const r = toolCallTextModule.extractTextToolCalls(input, knownForText);
  if (name === null) return r.toolCalls.length !== 0 || r.content !== rest;
  return r.toolCalls.length !== 1 || r.toolCalls[0].function.name !== name || r.content !== rest;
});
const bareArgs = toolCallTextModule.extractTextToolCalls(textToolCases[2][0], knownForText).toolCalls[0]?.function.arguments;
const nestedArgs = toolCallTextModule.extractTextToolCalls(textToolCases[5][0], knownForText).toolCalls[0]?.function.arguments;
if (textToolFailures.length || bareArgs !== '{"sheet":"S","key_column":"金额"}' || nestedArgs !== '{"sheet":"A"}') {
  errors.push(`文本工具调用解析失败: ${textToolFailures.map(([i]) => JSON.stringify(i.slice(0, 30))).join(", ") || "参数归一化不对"}`);
} else {
  checked.push("llm: tool calls written as <tool_call>/```json/bare JSON/name() are recovered");
}
const suggestOk =
  toolCallTextModule.suggestToolName("read_ranges", knownForText) === "read_range" &&
  toolCallTextModule.suggestToolName("workbook_overview", knownForText) === "get_workbook_overview" &&
  toolCallTextModule.suggestToolName("zzz", knownForText) === null;
if (!suggestOk) {
  errors.push("未知工具名纠错建议不正确");
} else {
  checked.push("agent: unknown tool names get a closest-match hint");
}
const planCases = [
  ["我将先调用 get_workbook_overview 查看工作簿结构，然后再进行排序。", true],
  ["好的，接下来我会使用工具读取数据并进行分析。\n1. 查看结构\n2. 排序", true],
  ["已完成：Sheet2 的 B 列按金额降序排列，共 120 行。", false],
  ["请问您希望按哪一列排序？", false],
  ["这个公式的意思是把 A 列求和。", false],
  ["我需要先了解表格结构。请您确认是否可以读取？", false],
];
const planFailures = planCases.filter(([t, exp]) => toolCallTextModule.looksLikePlanOnly(t, knownForText) !== exp);
if (planFailures.length) {
  errors.push(`口头计划识别不正确: ${planFailures.map(([t]) => JSON.stringify(t.slice(0, 20))).join(", ")}`);
} else {
  checked.push("agent: plan-only replies detected, questions/conclusions left alone");
}
if (!/looksLikePlanOnly/.test(agentSource) || !/interim:\s*true/.test(agentSource) || !/compactHistoryForRequest/.test(agentSource)) {
  errors.push("Agent 循环缺少计划纠偏、过程叙述标记或历史压缩");
} else {
  checked.push("agent: nudges plan-only replies, marks pre-tool text as interim, compacts history");
}
if (!/\[用户消息\]/.test(agentSource) || /history\.push\(buildUserMessage\([^)]*contextBlock/.test(agentSource)) {
  errors.push("实时上下文应只拼进当前请求，不应写入历史");
} else {
  checked.push("agent: live context is per-request, not persisted into history");
}

// ---- 请求前历史压缩 ----
const longHistory = [];
for (let i = 0; i < 6; i++) {
  longHistory.push({ role: "user", content: `任务 ${i} ` + "字".repeat(100) });
  longHistory.push({
    role: "assistant",
    content: null,
    tool_calls: [{ id: `c${i}`, type: "function", function: { name: "read_range", arguments: "{}" } }],
  });
  longHistory.push({ role: "tool", tool_call_id: `c${i}`, content: JSON.stringify({ ok: true, data: "x".repeat(4000) }) });
  longHistory.push({ role: "assistant", content: "done " + i });
}
const totalTokens = contextWindowModule.estimateMessagesTokens(longHistory);
const untouched = contextWindowModule.compactHistoryForRequest(longHistory, { budgetTokens: totalTokens + 10 });
const stubbed = contextWindowModule.compactHistoryForRequest(longHistory, { budgetTokens: Math.floor(totalTokens * 0.5) });
const dropped = contextWindowModule.compactHistoryForRequest(longHistory, { budgetTokens: 2500 });
const droppedCallIds = new Set(
  dropped.filter((m) => m.role === "assistant").flatMap((m) => (m.tool_calls || []).map((t) => t.id))
);
const droppedToolIds = dropped.filter((m) => m.role === "tool").map((m) => m.tool_call_id);
const pairsIntact =
  droppedToolIds.every((id) => droppedCallIds.has(id)) && [...droppedCallIds].every((id) => droppedToolIds.includes(id));
const compactOk =
  untouched === longHistory &&
  stubbed !== longHistory &&
  contextWindowModule.estimateMessagesTokens(stubbed) <= Math.floor(totalTokens * 0.5) &&
  stubbed.some((m) => m.role === "tool" && m.content.includes('"omitted":true')) &&
  stubbed[stubbed.length - 2].content.length > 4000 &&
  dropped.length < longHistory.length &&
  contextWindowModule.estimateMessagesTokens(dropped) <= 2500 &&
  /已省略/.test(dropped[0].content) &&
  dropped[dropped.length - 1].content === "done 5" &&
  pairsIntact &&
  longHistory.length === 24 &&
  longHistory[2].content.length > 4000;
if (!compactOk) {
  errors.push("历史压缩不正确：应先折叠旧工具结果、再整轮丢弃，且不改原数组、tool_calls 成对");
} else {
  checked.push("context: history compaction stubs old tool results, drops oldest turns, keeps pairs");
}
// ---- 严格服务商（Gemini）的工具 schema 适配 ----
const toolSchemaJs = ts.transpileModule(readFileSync(join(sourceRoot, "core", "toolSchema.ts"), "utf8"), transpileOpts).outputText;
const toolSchemaModule = await import(toDataUrl(toolSchemaJs));
const looseTools = [
  {
    type: "function",
    function: {
      name: "t",
      description: "d",
      parameters: {
        type: "object",
        properties: {
          key: { description: "表头名或列号" },
          values: { type: "array", items: { type: "array", items: {} } },
          cols: { type: "array", items: {} },
          mode: { type: "string", enum: ["a", "b"] },
        },
        required: ["key"],
      },
    },
  },
];
const untouchedTools = toolSchemaModule.adaptToolsForEndpoint(looseTools, "https://api.deepseek.com/v1");
const strictTools = toolSchemaModule.adaptToolsForEndpoint(looseTools, "https://generativelanguage.googleapis.com/v1beta/openai");
const strictProps = strictTools[0].function.parameters.properties;
const hasType = (node) => node && (node.type || node.anyOf || node.enum);
const strictOk =
  untouchedTools === looseTools &&
  looseTools[0].function.parameters.properties.key.type === undefined &&
  hasType(strictProps.key) &&
  hasType(strictProps.values.items.items) &&
  hasType(strictProps.cols.items) &&
  strictProps.mode.enum.length === 2 &&
  strictTools[0].function.parameters.required[0] === "key";
if (!strictOk) {
  errors.push("工具 schema 适配不正确：Gemini 端点应补齐 type，其余端点保持原样");
} else {
  checked.push("llm: tool schemas get explicit types only for strict endpoints");
}

if (!/max_tokens/.test(readFileSync(join(sourceRoot, "core", "llm.ts"), "utf8").split("function buildChatRequestBody")[1] || "")) {
  errors.push("聊天请求没有带 max_tokens");
} else {
  checked.push("llm: chat requests carry max_tokens when configured");
}

const caseToolNames = new Set();
function collectTools(node) {
  if (!node || typeof node !== "object") return;
  for (const key of ["must_use_tools", "should_use_tools", "must_not"]) {
    if (Array.isArray(node[key])) node[key].forEach((n) => caseToolNames.add(n));
  }
  if (Array.isArray(node.turns)) node.turns.forEach(collectTools);
}
for (const suite of cases.suites) {
  for (const cas of suite.cases) collectTools(cas);
}
const unknownCaseTools = [...caseToolNames].filter((n) => !toolNames.has(n));
if (unknownCaseTools.length) {
  errors.push(`用例引用了未定义工具: ${unknownCaseTools.join(", ")}`);
} else {
  checked.push("cases: all referenced tools exist");
}

function fail(id, msg) {
  errors.push(`${id}: ${msg}`);
}

function workbookFile(suite, cas) {
  const rel = cas.workbook || suite.workbook;
  return rel.replace(/^fixtures\//, "");
}

function factsOf(file) {
  const wb = truth.workbooks[file];
  if (!wb) return null;
  return wb.facts;
}

function catalog(file, facts) {
  if (file === "01-sales-ops.xlsx") {
    const cat = {
      订单行数: facts.order_count,
      金额为空的行数: facts.empty_amount_count,
      已填金额合计: facts.filled_amount_total,
      合计: facts.complete_amount_total,
      ...facts.by_region,
      ...facts.by_level,
      VIP: facts.vip_amount,
      普通: facts.normal_amount,
      vip_amount: facts.vip_amount,
      "ORD-1001 原金额": facts.first_data_amount,
    };
    for (const o of facts.empty_amount_orders) cat[`${o.订单号} 金额`] = o.应填金额;
    for (const [name, qty] of Object.entries(facts.by_product_qty)) cat[`${name} 数量`] = qty;
    return cat;
  }
  if (file === "02-dirty-data.xlsx") {
    return {
      cleaned_row_count: facts.unique_people,
      unique_people: facts.unique_people,
      empty_row_index_1based: facts.empty_row_index_1based,
    };
  }
  if (file === "03-hr-attendance.xlsx") {
    const cat = { ...facts.dept_hours };
    for (const p of facts.people) {
      cat[`${p.姓名} 实际工时`] = p.实际工时;
      cat[`${p.姓名} 加班`] = p.加班;
      cat[`${p.姓名} 缺勤`] = p.缺勤;
    }
    return cat;
  }
  if (file === "04-monthly-report.xlsx") {
    return {
      ...facts.region_totals,
      grand_total: facts.grand_total,
      合计: facts.grand_total,
    };
  }
  if (file === "05-edge-cases.xlsx") {
    return { hidden_cost_p003: facts.hidden_cost_p003 };
  }
  return {};
}

function checkNumeric(id, file, name, value) {
  const facts = factsOf(file);
  if (!facts) {
    fail(id, `找不到工作簿 ${file} 的 ground-truth`);
    return;
  }
  const cat = catalog(file, facts);
  if (!(name in cat)) {
    fail(id, `标准答案里没有「${name}」（${file}）`);
    return;
  }
  if (cat[name] !== value) {
    fail(id, `「${name}」cases=${value}，ground-truth=${cat[name]}`);
    return;
  }
  checked.push(`${id} ${name}=${value}`);
}

function walkNumericChecks(id, file, node) {
  if (!node || typeof node !== "object") return;
  if (Array.isArray(node.numeric_checks)) {
    for (const item of node.numeric_checks) {
      checkNumeric(id, file, item.name, item.value);
    }
  }
  if (Array.isArray(node.turns)) {
    for (const turn of node.turns) walkNumericChecks(id, file, turn.expect || turn);
  }
  if (node.expect) walkNumericChecks(id, file, node.expect);
}

function checkCase(suite, cas) {
  const id = cas.id;
  const file = workbookFile(suite, cas);
  const facts = factsOf(file);
  if (!facts) {
    fail(id, `suite 指向的 ${file} 不在 ground-truth 中`);
    return;
  }

  walkNumericChecks(id, file, cas);

  if (id === "S04" && cas.expect?.row_count !== facts.order_count) {
    fail(id, `row_count=${cas.expect.row_count}，应为 ${facts.order_count}`);
  } else if (id === "S04") {
    checked.push(`${id} row_count=${facts.order_count}`);
  }

  if (id === "S06") {
    const w = cas.expect?.winner;
    const margin = facts.unit_margin["27寸显示器"];
    if (w?.产品 !== "27寸显示器" || w?.单件毛利 !== margin) {
      fail(id, `winner 单件毛利 ${w?.单件毛利}，ground-truth ${margin}`);
    } else {
      checked.push(`${id} 单件毛利=${margin}`);
    }
  }

  if (id === "S07") {
    if (cas.expect?.first_amount !== facts.max_order.金额) {
      fail(id, `first_amount=${cas.expect.first_amount}，应为 ${facts.max_order.金额}`);
    } else {
      checked.push(`${id} first_amount=${facts.max_order.金额}`);
    }
  }

  if (id === "S08" && cas.expect?.min_matches !== facts.p003_match_count) {
    fail(id, `min_matches=${cas.expect.min_matches}，P003 出现 ${facts.p003_match_count} 次`);
  } else if (id === "S08") {
    checked.push(`${id} min_matches=${facts.p003_match_count}`);
  }

  if (id === "S11" && cas.expect?.old_value !== facts.first_data_amount) {
    fail(id, `old_value=${cas.expect.old_value}，G2 应为 ${facts.first_data_amount}`);
  } else if (id === "S11") {
    checked.push(`${id} old_value=${facts.first_data_amount}`);
  }

  if (id === "D02" && cas.expect?.cleaned_row_count !== facts.unique_people) {
    fail(id, `cleaned_row_count=${cas.expect.cleaned_row_count}，应为 ${facts.unique_people}`);
  } else if (id === "D02") {
    checked.push(`${id} cleaned_row_count=${facts.unique_people}`);
  }

  if ((id === "B07" || id === "D10") && cas.expect?.empty_row_index_1based !== facts.empty_row_index_1based) {
    fail(id, `空行行号 ${cas.expect.empty_row_index_1based}，应为 ${facts.empty_row_index_1based}`);
  } else if (id === "B07" || id === "D10") {
    checked.push(`${id} empty_row=${facts.empty_row_index_1based}`);
  }

  if (id === "H01") {
    const got = cas.expect?.people || [];
    for (const p of facts.people) {
      const row = got.find((x) => x.姓名 === p.姓名);
      if (!row) {
        fail(id, `缺少 ${p.姓名}`);
        continue;
      }
      for (const k of ["实际工时", "加班", "缺勤"]) {
        if (row[k] !== p[k]) fail(id, `${p.姓名}.${k} cases=${row[k]} truth=${p[k]}`);
        else checked.push(`${id} ${p.姓名}.${k}=${p[k]}`);
      }
    }
  }

  if (id === "M02") {
    const totals = cas.expect?.region_totals || {};
    for (const [k, v] of Object.entries(facts.region_totals)) {
      if (totals[k] !== v) fail(id, `region_totals.${k} cases=${totals[k]} truth=${v}`);
      else checked.push(`${id} ${k}=${v}`);
    }
    if (cas.expect?.grand_total !== facts.grand_total) {
      fail(id, `grand_total=${cas.expect.grand_total}，应为 ${facts.grand_total}`);
    } else {
      checked.push(`${id} grand_total=${facts.grand_total}`);
    }
  }

  if (id === "M04") {
    const e = cas.expect || {};
    if (e.best_cell?.金额 !== facts.best_cell.金额) {
      fail(id, `best_cell 金额 ${e.best_cell?.金额}，应为 ${facts.best_cell.金额}`);
    } else checked.push(`${id} best_cell=${facts.best_cell.金额}`);
    if (e.best_month_total?.合计 !== facts.best_month.合计) {
      fail(id, `best_month ${e.best_month_total?.合计}，应为 ${facts.best_month.合计}`);
    } else checked.push(`${id} best_month=${facts.best_month.合计}`);
    if (e.worst_month_total?.合计 !== facts.worst_month.合计) {
      fail(id, `worst_month ${e.worst_month_total?.合计}，应为 ${facts.worst_month.合计}`);
    } else checked.push(`${id} worst_month=${facts.worst_month.合计}`);
  }

  if (id === "B01") {
    const t0 = cas.turns?.[0]?.expect?.vip_amount;
    const t1 = cas.turns?.[1]?.expect || {};
    if (t0 !== facts.vip_amount) fail(id, `第一轮 VIP ${t0}，应为 ${facts.vip_amount}`);
    else checked.push(`${id} vip_amount=${facts.vip_amount}`);
    if (t1.VIP !== facts.vip_amount || t1.普通 !== facts.normal_amount) {
      fail(id, `第二轮 VIP/普通 与 ground-truth 不一致`);
    } else {
      checked.push(`${id} 普通=${facts.normal_amount}`);
    }
  }

  if (id === "B06") {
    const qty = cas.expect?.qty || {};
    for (const [name, n] of Object.entries(facts.by_product_qty)) {
      if (qty[name] !== n) fail(id, `qty.${name} cases=${qty[name]} truth=${n}`);
      else checked.push(`${id} ${name} 数量=${n}`);
    }
  }

  if (id === "E04" && cas.expect?.value !== facts.hidden_cost_p003) {
    fail(id, `P003 成本 ${cas.expect.value}，应为 ${facts.hidden_cost_p003}`);
  } else if (id === "E04") {
    checked.push(`${id} hidden_cost=${facts.hidden_cost_p003}`);
  }
}

let caseCount = 0;
for (const suite of cases.suites) {
  for (const cas of suite.cases) {
    caseCount += 1;
    checkCase(suite, cas);
  }
}

if (caseCount !== 56) {
  errors.push(`用例总数 ${caseCount}，期望 56`);
}

console.log(`核对 ${caseCount} 条用例，${checked.length} 个数字锚点`);
if (errors.length) {
  console.error(`失败 ${errors.length} 项：`);
  for (const e of errors) console.error("  - " + e);
  process.exit(1);
}
console.log("cases.json 与 ground-truth.json 数字一致");
