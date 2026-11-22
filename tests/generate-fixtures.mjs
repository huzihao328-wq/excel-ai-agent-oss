/**
 * 生成评测用 xlsx 工作簿 + 标准答案。
 * 不依赖第三方库，直接写 Office Open XML。
 *
 *   node tests/generate-fixtures.mjs
 */
import { deflateRawSync } from "zlib";
import { mkdirSync, writeFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const ROOT = dirname(fileURLToPath(import.meta.url));
const OUT = join(ROOT, "fixtures");

// ---------------------------------------------------------------------------
// CRC32 + ZIP
// ---------------------------------------------------------------------------

const CRC_TABLE = new Uint32Array(256);
for (let i = 0; i < 256; i++) {
  let c = i;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  CRC_TABLE[i] = c >>> 0;
}

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function u16(n) {
  const b = Buffer.alloc(2);
  b.writeUInt16LE(n);
  return b;
}
function u32(n) {
  const b = Buffer.alloc(4);
  b.writeUInt32LE(n);
  return b;
}

function zipStore(files) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const { name, data } of files) {
    const raw = Buffer.isBuffer(data) ? data : Buffer.from(data, "utf8");
    const compressed = deflateRawSync(raw);
    const crc = crc32(raw);
    const nameBuf = Buffer.from(name, "utf8");
    const local = Buffer.concat([
      Buffer.from("PK\u0003\u0004", "binary"),
      u16(20),
      u16(0),
      u16(8),
      u16(0),
      u16(0),
      u32(crc),
      u32(compressed.length),
      u32(raw.length),
      u16(nameBuf.length),
      u16(0),
      nameBuf,
      compressed,
    ]);
    const central = Buffer.concat([
      Buffer.from("PK\u0001\u0002", "binary"),
      u16(20),
      u16(20),
      u16(0),
      u16(8),
      u16(0),
      u16(0),
      u32(crc),
      u32(compressed.length),
      u32(raw.length),
      u16(nameBuf.length),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(0),
      u32(offset),
      nameBuf,
    ]);
    locals.push(local);
    centrals.push(central);
    offset += local.length;
  }
  const centralBuf = Buffer.concat(centrals);
  const end = Buffer.concat([
    Buffer.from("PK\u0005\u0006", "binary"),
    u16(0),
    u16(0),
    u16(files.length),
    u16(files.length),
    u32(centralBuf.length),
    u32(offset),
    u16(0),
  ]);
  return Buffer.concat([...locals, centralBuf, end]);
}

// ---------------------------------------------------------------------------
// 极简 xlsx 写入（inlineStr + 数字）
// ---------------------------------------------------------------------------

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function colLetter(index) {
  let s = "";
  let n = index + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function cellXml(r, c, value) {
  const ref = `${colLetter(c)}${r + 1}`;
  if (value === null || value === undefined || value === "") {
    return `<c r="${ref}"/>`;
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return `<c r="${ref}" t="n"><v>${value}</v></c>`;
  }
  if (typeof value === "boolean") {
    return `<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`;
  }
  return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
}

function sheetXml(rows) {
  const maxCols = rows.reduce((m, r) => Math.max(m, r.length), 0);
  const dim = `A1:${colLetter(Math.max(maxCols - 1, 0))}${Math.max(rows.length, 1)}`;
  const rowXml = rows
    .map((row, r) => {
      const cells = row.map((v, c) => cellXml(r, c, v)).join("");
      return `<row r="${r + 1}">${cells}</row>`;
    })
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<dimension ref="${dim}"/>` +
    `<sheetData>${rowXml}</sheetData>` +
    `</worksheet>`
  );
}

function workbookXml(sheets) {
  const sheetEls = sheets
    .map(
      (s, i) =>
        `<sheet name="${xmlEscape(s.name)}" sheetId="${i + 1}"` +
        `${s.hidden ? ' state="hidden"' : ""} r:id="rId${i + 1}"/>`
    )
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<sheets>${sheetEls}</sheets>` +
    `</workbook>`
  );
}

function workbookRels(sheets) {
  const rels = sheets
    .map(
      (_, i) =>
        `<Relationship Id="rId${i + 1}" ` +
        `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ` +
        `Target="worksheets/sheet${i + 1}.xml"/>`
    )
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    rels +
    `</Relationships>`
  );
}

function contentTypes(sheets) {
  const overrides = sheets
    .map(
      (_, i) =>
        `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ` +
        `ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    )
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
    overrides +
    `</Types>`
  );
}

const ROOT_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
  `</Relationships>`;

const STYLES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
  `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill>` +
  `<fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
  `<cellStyleXfs count="1"><xf/></cellStyleXfs>` +
  `<cellXfs count="1"><xf xfId="0"/></cellXfs>` +
  `</styleSheet>`;

function writeXlsx(filePath, sheets) {
  const files = [
    { name: "[Content_Types].xml", data: contentTypes(sheets) },
    { name: "_rels/.rels", data: ROOT_RELS },
    { name: "xl/workbook.xml", data: workbookXml(sheets) },
    { name: "xl/_rels/workbook.xml.rels", data: workbookRels(sheets) },
    { name: "xl/styles.xml", data: STYLES_XML },
  ];
  sheets.forEach((s, i) => {
    files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(s.rows) });
  });
  mkdirSync(dirname(filePath), { recursive: true });
  writeFileSync(filePath, zipStore(files));
}

// ---------------------------------------------------------------------------
// 工作簿数据
// ---------------------------------------------------------------------------

const CUSTOMERS = [
  ["客户编号", "客户名称", "城市", "地区", "等级"],
  ["C001", "华东贸易", "上海", "华东", "VIP"],
  ["C002", "南方科技", "深圳", "华南", "普通"],
  ["C003", "北方机械", "北京", "华北", "VIP"],
  ["C004", "西部能源", "成都", "西南", "普通"],
  ["C005", "齐鲁商贸", "济南", "华东", "普通"],
];

const PRODUCTS = [
  ["产品编号", "产品名称", "品类", "单价", "成本"],
  ["P001", "无线鼠标", "配件", 45, 28],
  ["P002", "机械键盘", "配件", 289, 160],
  ["P003", "27寸显示器", "显示", 1299, 780],
  ["P004", "笔记本支架", "配件", 79, 32],
  ["P005", "USB-C扩展坞", "配件", 199, 95],
];

/** 金额故意留空 3 行，用于「补公式」用例 */
const ORDER_ROWS = [
  ["ORD-1001", "2024-01-08", "C001", "P001", 10, 45, 450],
  ["ORD-1002", "2024-01-12", "C002", "P003", 2, 1299, 2598],
  ["ORD-1003", "2024-01-15", "C001", "P002", 5, 289, 1445],
  ["ORD-1004", "2024-02-03", "C003", "P005", 8, 199, 1592],
  ["ORD-1005", "2024-02-10", "C004", "P001", 20, 45, 900],
  ["ORD-1006", "2024-02-18", "C002", "P004", 15, 79, 1185],
  ["ORD-1007", "2024-03-02", "C005", "P003", 1, 1299, 1299],
  ["ORD-1008", "2024-03-11", "C001", "P005", 4, 199, 796],
  ["ORD-1009", "2024-03-20", "C003", "P002", 3, 289, 867],
  ["ORD-1010", "2024-04-05", "C004", "P003", 3, 1299, 3897],
  ["ORD-1011", "2024-04-14", "C002", "P001", 30, 45, 1350],
  ["ORD-1012", "2024-04-22", "C005", "P004", 8, 79, 632],
  ["ORD-1013", "2024-05-06", "C001", "P003", 2, 1299, 2598],
  ["ORD-1014", "2024-05-19", "C003", "P001", 12, 45, 540],
  ["ORD-1015", "2024-06-01", "C004", "P002", 6, 289, 1734],
  ["ORD-1016", "2024-06-15", "C002", "P005", 10, 199, 1990],
  ["ORD-1017", "2024-07-03", "C005", "P001", 25, 45, 1125],
  ["ORD-1018", "2024-07-18", "C001", "P004", 20, 79, ""],
  ["ORD-1019", "2024-08-02", "C003", "P003", 1, 1299, ""],
  ["ORD-1020", "2024-08-10", "C004", "P005", 6, 199, ""],
];

const ORDERS = [["订单号", "日期", "客户编号", "产品编号", "数量", "单价", "金额"], ...ORDER_ROWS];

const SIGNUPS = [
  ["姓名", "手机", "邮箱", "城市", "报名日期", "年龄", "状态", "缴费"],
  ["张伟", "13800001111", "zhang@example.com", "北京", "2024-03-05", 28, "已确认", 1200],
  ["李娜", "139-0000-2222", "li@example.com", "上海市", "2024/3/8", 32, "确认", "¥1500"],
  ["  王强  ", "13700003333", "", "北京", "3月12日", 45, "pending", "800元"],
  ["张伟", "13800001111", "zhang@example.com", "北京", "2024-03-05", 28, "已确认", 1200],
  ["", "", "", "", "", "", "", ""],
  ["赵敏", "+86 13600004444", "zhao@example.com", "BJ", "2024.03.15", 29, "待确认", "1,200"],
  ["陈晨", "13500005555", "chen@example.com", "上海", "2024-03-18", 199, "已确认", 1200],
  ["刘洋", "13400006666", "liu@example.com", "北京市", "2024-03-20", 31, "已确认", 1200],
  ["周杰", "13300007777", "", "广州", "2024-03-22", 26, "待确认", ""],
  ["吴芳", "13200008888", "wu@example.com", "深圳", "2024-03-25", 35, "已确认", 2000],
];

const ATTENDANCE = [
  ["姓名", "部门", "周一", "周二", "周三", "周四", "周五", "周六", "周日", "应出勤"],
  ["周婷", "销售", 8, 8, 8, 8, 8, 0, 0, 40],
  ["韩磊", "销售", 9, 8, 10, 8, 8, 4, 0, 40],
  ["沈佳", "研发", 8, 8, 8, 8, 8, 0, 0, 40],
  ["马超", "研发", 8, 0, 8, 8, 12, 8, 0, 40],
  ["林雪", "运营", 7.5, 8, 8, 8, 7.5, 0, 0, 40],
  ["何伟", "运营", 8, 8, 8, 8, 8, 0, 0, 40],
];

const MONTHLY = [
  ["月份", "华东", "华南", "华北", "西南"],
  ["2024-01", 6869, 2598, 0, 0],
  ["2024-02", 0, 1185, 1592, 900],
  ["2024-03", 2095, 0, 867, 0],
  ["2024-04", 632, 1350, 0, 3897],
  ["2024-05", 2598, 0, 540, 0],
  ["2024-06", 0, 1990, 0, 1734],
];

const EDGE_HEADERS_ONLY = [["编号", "名称", "数量", "单价", "备注"]];

const EDGE_WEIRD = [
  ["说明", "值"],
  ["普通中文", "你好世界"],
  ["表情", "测试😀"],
  ["尖括号", "<script>alert(1)</script>"],
  ["看起来像公式", "=1+1"],
  ["超长文本", "这是一段用来测试截断和预览的超长说明文字".repeat(8)],
];

const HIDDEN_COST = [
  ["产品编号", "内部成本", "供应商"],
  ["P001", 28, "苏州A厂"],
  ["P002", 160, "东莞B厂"],
  ["P003", 780, "合肥C厂"],
  ["P004", 32, "宁波D厂"],
  ["P005", 95, "深圳E厂"],
];

// ---------------------------------------------------------------------------
// 标准答案
// ---------------------------------------------------------------------------

function customerMap() {
  const m = {};
  for (const r of CUSTOMERS.slice(1)) m[r[0]] = { id: r[0], name: r[1], city: r[2], region: r[3], level: r[4] };
  return m;
}
function productMap() {
  const m = {};
  for (const r of PRODUCTS.slice(1)) m[r[0]] = { id: r[0], name: r[1], category: r[2], price: r[3], cost: r[4] };
  return m;
}

function filledAmount(row) {
  return row[6] === "" ? row[4] * row[5] : row[6];
}

function salesFacts() {
  const cust = customerMap();
  const prod = productMap();
  const filled = ORDER_ROWS.filter((r) => r[6] !== "");
  const all = ORDER_ROWS.map((r) => ({ row: r, amount: filledAmount(r) }));

  const byRegion = {};
  const byProduct = {};
  const byProductQty = {};
  const byLevel = {};
  let filledTotal = 0;
  let completeTotal = 0;
  for (const { row, amount } of all) {
    completeTotal += amount;
    const region = cust[row[2]].region;
    const level = cust[row[2]].level;
    const pname = prod[row[3]].name;
    byRegion[region] = (byRegion[region] || 0) + amount;
    byProduct[pname] = (byProduct[pname] || 0) + amount;
    byProductQty[pname] = (byProductQty[pname] || 0) + row[4];
    byLevel[level] = (byLevel[level] || 0) + amount;
  }
  for (const r of filled) filledTotal += r[6];

  const unitMargin = {};
  for (const p of Object.values(prod)) {
    unitMargin[p.name] = p.price - p.cost;
  }

  const p003InOrders = ORDER_ROWS.filter((r) => r[3] === "P003").length;
  const p003InProducts = PRODUCTS.slice(1).filter((r) => r[0] === "P003").length;

  const emptyAmountOrders = ORDER_ROWS.filter((r) => r[6] === "").map((r) => ({
    订单号: r[0],
    应填金额: r[4] * r[5],
  }));

  return {
    order_count: ORDER_ROWS.length,
    filled_amount_count: filled.length,
    empty_amount_count: emptyAmountOrders.length,
    empty_amount_orders: emptyAmountOrders,
    filled_amount_total: filledTotal,
    complete_amount_total: completeTotal,
    by_region: byRegion,
    by_product: byProduct,
    by_product_qty: byProductQty,
    by_level: byLevel,
    vip_amount: byLevel.VIP,
    normal_amount: byLevel.普通,
    unique_customers: 5,
    unique_products: 5,
    max_order: { 订单号: "ORD-1010", 金额: 3897 },
    first_data_amount: ORDER_ROWS[0][6],
    unit_margin: unitMargin,
    p003_match_count: p003InOrders + p003InProducts,
    sheets: ["订单", "客户", "产品"],
  };
}

function dirtyFacts() {
  return {
    raw_rows_including_header: SIGNUPS.length,
    data_rows: SIGNUPS.length - 1,
    unique_people: 8,
    duplicate_name: "张伟",
    duplicate_phone: "13800001111",
    empty_row_index_1based: 6,
    missing_email: ["王强", "周杰"],
    missing_payment: ["周杰"],
    age_outlier: { 姓名: "陈晨", 年龄: 199 },
    city_aliases: { 北京: ["北京", "北京市", "BJ"], 上海: ["上海", "上海市"] },
    status_aliases: { 已确认: ["已确认", "确认"], 待确认: ["待确认", "pending"] },
    issues: [
      "完全重复行：张伟 / 13800001111",
      "空行",
      "姓名前后空格：王强",
      "邮箱缺失：王强、周杰",
      "缴费缺失：周杰",
      "年龄异常：陈晨 199",
      "日期格式不统一",
      "手机号格式不统一",
      "城市写法不统一（北京/北京市/BJ，上海/上海市）",
      "状态写法不统一（已确认/确认/pending/待确认）",
      "缴费格式不统一（数字 / ¥ / 元 / 千分位）",
    ],
  };
}

function hrFacts() {
  const people = ATTENDANCE.slice(1).map((r) => {
    const hours = r.slice(2, 9);
    const actual = hours.reduce((a, b) => a + b, 0);
    const required = r[9];
    return {
      姓名: r[0],
      部门: r[1],
      实际工时: actual,
      应出勤: required,
      加班: Math.max(0, actual - required),
      缺勤: Math.max(0, required - actual),
    };
  });
  return {
    people,
    overtime_people: people.filter((p) => p.加班 > 0).map((p) => p.姓名),
    under_people: people.filter((p) => p.缺勤 > 0).map((p) => p.姓名),
    dept_hours: people.reduce((acc, p) => {
      acc[p.部门] = (acc[p.部门] || 0) + p.实际工时;
      return acc;
    }, {}),
  };
}

function monthlyFacts() {
  const rows = MONTHLY.slice(1);
  const regions = MONTHLY[0].slice(1);
  const totals = {};
  for (const name of regions) totals[name] = 0;
  let grand = 0;
  const month_totals = {};
  let bestCell = { 月份: "", 地区: "", 金额: -Infinity };
  for (const r of rows) {
    let monthSum = 0;
    regions.forEach((name, i) => {
      const v = r[i + 1];
      totals[name] += v;
      grand += v;
      monthSum += v;
      if (v > bestCell.金额) bestCell = { 月份: r[0], 地区: name, 金额: v };
    });
    month_totals[r[0]] = monthSum;
  }
  const monthEntries = Object.entries(month_totals);
  const bestMonth = monthEntries.reduce((a, b) => (b[1] > a[1] ? b : a));
  const worstMonth = monthEntries.reduce((a, b) => (b[1] < a[1] ? b : a));
  return {
    region_totals: totals,
    grand_total: grand,
    months: rows.length,
    month_totals,
    best_cell: bestCell,
    best_month: { 月份: bestMonth[0], 合计: bestMonth[1] },
    worst_month: { 月份: worstMonth[0], 合计: worstMonth[1] },
  };
}

function buildGroundTruth() {
  return {
    generated_at: new Date().toISOString(),
    workbooks: {
      "01-sales-ops.xlsx": {
        description: "销售运营：订单 / 客户 / 产品三表，金额列末 3 行为空",
        sheets: ["订单", "客户", "产品"],
        facts: salesFacts(),
      },
      "02-dirty-data.xlsx": {
        description: "报名名单脏数据：重复、空行、格式混乱、异常值",
        sheets: ["报名名单"],
        facts: dirtyFacts(),
      },
      "03-hr-attendance.xlsx": {
        description: "一周考勤，用于公式、加班/缺勤、按部门汇总",
        sheets: ["考勤"],
        facts: hrFacts(),
      },
      "04-monthly-report.xlsx": {
        description: "地区月度销售，适合图表和格式化",
        sheets: ["月度销售"],
        facts: monthlyFacts(),
      },
      "05-edge-cases.xlsx": {
        description: "空表、仅表头、特殊字符、隐藏成本表",
        sheets: ["空表", "仅表头", "特殊字符", "内部成本"],
        hidden_sheets: ["内部成本"],
        facts: {
          visible_sheets: ["空表", "仅表头", "特殊字符"],
          hidden_sheet: "内部成本",
          headers_only_columns: ["编号", "名称", "数量", "单价", "备注"],
          hidden_cost_p003: 780,
        },
      },
    },
  };
}

function main() {
  mkdirSync(OUT, { recursive: true });

  writeXlsx(join(OUT, "01-sales-ops.xlsx"), [
    { name: "订单", rows: ORDERS },
    { name: "客户", rows: CUSTOMERS },
    { name: "产品", rows: PRODUCTS },
  ]);

  writeXlsx(join(OUT, "02-dirty-data.xlsx"), [{ name: "报名名单", rows: SIGNUPS }]);

  writeXlsx(join(OUT, "03-hr-attendance.xlsx"), [{ name: "考勤", rows: ATTENDANCE }]);

  writeXlsx(join(OUT, "04-monthly-report.xlsx"), [{ name: "月度销售", rows: MONTHLY }]);

  writeXlsx(join(OUT, "05-edge-cases.xlsx"), [
    { name: "空表", rows: [] },
    { name: "仅表头", rows: EDGE_HEADERS_ONLY },
    { name: "特殊字符", rows: EDGE_WEIRD },
    { name: "内部成本", rows: HIDDEN_COST, hidden: true },
  ]);

  const truth = buildGroundTruth();
  writeFileSync(join(OUT, "ground-truth.json"), JSON.stringify(truth, null, 2), "utf8");

  const names = Object.keys(truth.workbooks);
  console.log(`已生成 ${names.length} 个工作簿 -> ${OUT}`);
  names.forEach((n) => console.log("  -", n));
}

main();
