import { ToolDefinition } from "../types";

/** 展示给用户看的工具中文名 */
export const TOOL_LABELS: Record<string, string> = {
  get_workbook_overview: "查看工作簿结构",
  read_range: "读取区域",
  write_range: "写入数据",
  set_formulas: "写入公式",
  create_sheet: "新建 Sheet",
  delete_sheet: "删除 Sheet",
  clear_range: "清空区域",
  insert_rows: "插入行",
  delete_rows: "删除行",
  insert_columns: "插入列",
  delete_columns: "删除列",
  format_range: "设置格式",
  sort_range: "排序",
  find_cells: "查找",
  find_replace: "查找替换",
  create_chart: "创建图表",
  copy_range: "复制区域",
  fill_down: "向下填充",
  remove_duplicates: "删除重复项",
  rename_sheet: "重命名 Sheet",
  copy_sheet: "复制 Sheet",
  column_stats: "列统计",
  apply_conditional_format: "条件格式",
  add_lookup_column: "查找匹配",
  split_column: "拆分列",
  combine_columns: "合并列",
  freeze_panes: "冻结窗格",
  clean_column: "清洗列",
  convert_to_values: "公式转成值",
  polish_range: "整理格式",
  delete_blank_rows: "删除空行",
  list_external_files: "查看附加文件",
  read_external_sheet: "读取附加文件",
  import_external_sheet: "导入附加文件",
  export_file: "导出文件",
  ask_user_choice: "请你选一个方向",
  inspect_range: "查看选区公式",
  list_tables: "查看 Excel 表",
  create_table: "转为 Excel 表",
  apply_autofilter: "筛选数据",
  clear_autofilter: "清除筛选",
  create_pivot_table: "创建透视表",
  group_summary: "分组汇总",
  reconcile_tables: "两表对账",
  stack_sheets: "多表合并",
  split_by_column: "按列拆表",
  suggest_followups: "建议下一步",
  remember: "写入记忆",
  forget: "忘掉一条",
  list_memories: "查看记忆",
  save_skill: "沉淀技能",
  list_skills: "查看技能库",
  audit_data: "数据体检",
  diagnose_errors: "诊断全表错误",
  run_script: "运行计算脚本",
  set_data_validation: "数据验证",
  set_visibility: "隐藏/显示",
  unpivot_columns: "宽表转长表",
  fill_blank_cells: "填充空白/拆合并",
  extract_text: "提取信息",
  compare_sheets: "版本比对",
  set_page_layout: "打印设置",
  create_sheet_index: "生成目录",
};

export const TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    type: "function",
    function: {
      name: "get_workbook_overview",
      description:
        "获取整个工作簿的结构概览：所有 Sheet 的名称、数据范围、表头和前几行样本数据。开始任何任务前先调用它。",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
  {
    type: "function",
    function: {
      name: "read_range",
      description:
        "读取指定 Sheet 的一个区域的值。未指定分页时，超过 5000 单元格会自动返回前 30 行+末 5 行抽样（data.sampled=true），不要误以为表只有这么多行。大表统计请用 column_stats，复制请用 copy_range。可用 start_row/max_rows 分页。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "Sheet 名称" },
          range: { type: "string", description: "A1 表示法的区域，如 'A1:D20'；省略则读取整个已用区域" },
          include_formulas: { type: "boolean", description: "是否同时返回公式，默认 false" },
          start_row: { type: "number", description: "区域内的起始行（从 1 开始），用于分页" },
          max_rows: { type: "number", description: "本次最多返回多少行" },
        },
        required: ["sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "write_range",
      description:
        "把二维数组写入指定位置。values 是行优先的二维数组，起点是 start_cell。数字请用数字类型；以 =、+、-、@ 开头的字符串会按普通文本保留，公式必须改用 set_formulas。会自动记录快照供撤销。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "Sheet 名称" },
          start_cell: { type: "string", description: "起始单元格，如 'A1'" },
          values: {
            type: "array",
            items: { type: "array", items: {} },
            description: "二维数组，例如 [[\"名称\",\"数量\"],[\"苹果\",3]]",
          },
        },
        required: ["sheet", "start_cell", "values"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "set_formulas",
      description:
        "把公式二维数组写入指定位置（以 = 开头）。适合批量填充公式列。普通值也可以混在其中。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          start_cell: { type: "string", description: "起始单元格，如 'E2'" },
          formulas: {
            type: "array",
            items: { type: "array", items: {} },
            description: "二维数组，如 [[\"=SUM(A2:D2)\"],[\"=SUM(A3:D3)\"]]",
          },
        },
        required: ["sheet", "start_cell", "formulas"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "create_sheet",
      description: "创建一个新的 Sheet。若同名 Sheet 已存在会报错。",
      parameters: {
        type: "object",
        properties: {
          name: { type: "string", description: "新 Sheet 名称" },
          activate: { type: "boolean", description: "创建后是否切换过去，默认 true" },
        },
        required: ["name"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "delete_sheet",
      description: "删除一个 Sheet。单元格级撤销无法还原；任务检查点可以整段回退。会请求用户确认。",
      parameters: {
        type: "object",
        properties: { name: { type: "string" } },
        required: ["name"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "clear_range",
      description: "清空区域的内容和格式（可撤销内容，格式不可恢复）。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "A1 表示法" },
          what: { type: "string", enum: ["contents", "formats", "all"], description: "默认 contents" },
        },
        required: ["sheet", "range"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_rows",
      description: "在指定行号前插入若干空行（行号从 1 开始）。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          before_row: { type: "number", description: "在该行之前插入" },
          count: { type: "number", description: "插入行数，默认 1" },
        },
        required: ["sheet", "before_row"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "delete_rows",
      description: "删除若干行。单元格级撤销无法还原；任务检查点可以整段回退。会请求用户确认。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          start_row: { type: "number", description: "起始行号（从 1 开始）" },
          count: { type: "number", description: "删除行数，默认 1" },
        },
        required: ["sheet", "start_row"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_columns",
      description: "在指定列前插入若干空列。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          before_column: { type: "string", description: "列字母，如 'C'，在其前插入" },
          count: { type: "number", description: "默认 1" },
        },
        required: ["sheet", "before_column"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "delete_columns",
      description: "删除若干列。单元格级撤销无法还原；任务检查点可以整段回退。会请求用户确认。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          start_column: { type: "string", description: "起始列字母，如 'C'" },
          count: { type: "number", description: "默认 1" },
        },
        required: ["sheet", "start_column"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "format_range",
      description: "设置区域格式：字体、颜色、数字格式、对齐、边框、自动换行、合并单元格、列宽行高等。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string" },
          bold: { type: "boolean" },
          italic: { type: "boolean" },
          font_size: { type: "number" },
          font_color: { type: "string", description: "如 '#FFFFFF'" },
          fill_color: { type: "string", description: "背景色，如 '#1A8C4E'" },
          number_format: { type: "string", description: "如 '0.00'、'0.0%'、'yyyy-mm-dd'、'#,##0'" },
          horizontal_alignment: { type: "string", enum: ["Left", "Center", "Right"] },
          border: { type: "boolean", description: "是否加全边框" },
          autofit_columns: { type: "boolean", description: "是否自适应列宽" },
          autofit_rows: { type: "boolean", description: "是否自适应行高（配合 wrap_text）" },
          vertical_alignment: { type: "string", enum: ["Top", "Center", "Bottom"] },
          wrap_text: { type: "boolean", description: "自动换行" },
          merge: { type: "boolean", description: "true 合并区域为一个单元格（只保留左上角的值，常用于标题行）；false 取消合并" },
          column_width: { type: "number", description: "列宽，按 Excel 字符数，如 12" },
          row_height: { type: "number", description: "行高，单位磅，如 24" },
        },
        required: ["sheet", "range"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "sort_range",
      description:
        "对区域按一列或多列排序。key_column 可以是表头名、列字母或区域内 0 基列号（0=第一列）。省略 range 则用已用区域。多列排序用 keys。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "包含表头的完整区域，如 'A1:E100'；省略则用已用区域" },
          key_column: { description: "排序键：表头名、列字母或区域内 0 基列号" },
          column: { description: "key_column 的别名" },
          ascending: { type: "boolean", description: "默认 true" },
          has_headers: { type: "boolean", description: "区域第一行是否为表头，默认 true" },
          keys: {
            type: "array",
            items: {
              type: "object",
              properties: {
                column: { description: "表头名、列字母或 0 基列号" },
                ascending: { type: "boolean" },
              },
            },
            description: "多列排序，按数组顺序依次比较。提供时优先于 key_column",
          },
        },
        required: ["sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "find_cells",
      description: "在指定 Sheet（或全部 Sheet）中查找文本，返回匹配单元格地址列表。",
      parameters: {
        type: "object",
        properties: {
          text: { type: "string", description: "要查找的文本" },
          sheet: { type: "string", description: "省略则搜索所有 Sheet" },
          match_case: { type: "boolean", description: "默认 false" },
          complete_match: { type: "boolean", description: "是否要求整格匹配，默认 false" },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "find_replace",
      description:
        "在指定区域把查找文本全部替换为新文本。适合改状态词、统一写法。会先统计匹配数并请求确认。大范围替换优先用它，不要逐格 write_range。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "A1 区域；省略则用已用区域" },
          find: { type: "string", description: "要查找的文本" },
          replace: { type: "string", description: "替换成的文本，可为空字符串" },
          match_case: { type: "boolean", description: "默认 false" },
          complete_match: { type: "boolean", description: "是否要求整格匹配，默认 false" },
        },
        required: ["sheet", "find", "replace"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "copy_range",
      description:
        "用 Excel 原生复制把一个区域拷到目标位置，不经过模型读写单元格。大表复制、转置、把结果拷到新表时必须用它，禁止 read_range + write_range。源与目标不能重叠。",
      parameters: {
        type: "object",
        properties: {
          source_sheet: { type: "string" },
          source_range: { type: "string", description: "源区域，如 'A1:D100'" },
          dest_sheet: { type: "string", description: "目标 Sheet，省略则与源相同" },
          dest_start_cell: { type: "string", description: "目标左上角，如 'A1' 或 'E1'" },
          copy_type: { type: "string", enum: ["all", "values", "formulas", "formats"], description: "默认 all" },
          transpose: { type: "boolean", description: "是否转置行列，默认 false" },
        },
        required: ["source_sheet", "source_range", "dest_start_cell"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "fill_down",
      description:
        "把源区域的公式/值按 Excel 填充柄逻辑向下（或向右）填充。写好第一行公式后用它铺到整列，不要对每一行 set_formulas。destination_range 必须从源区域左上角开始并包含源区域。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          source_range: { type: "string", description: "已写好的公式/值，如 'E2' 或 'E2:G2'" },
          destination_range: { type: "string", description: "填充目标，如 'E2:E100'，必须包含 source_range" },
          end_row: { type: "number", description: "若未给 destination_range，则填充到该行号（含）" },
        },
        required: ["sheet", "source_range"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "add_lookup_column",
      description:
        "按关联键把另一张表的列用 XLOOKUP（可改 VLOOKUP）公式写到目标表。跨表合并、配客户名称/地区/产品名时必须用它，不要 read_range 后手填，也不要逐行 set_formulas。默认追加到目标表右侧；要保留原表时先 copy_sheet 再对副本调用。",
      parameters: {
        type: "object",
        properties: {
          dest_sheet: { type: "string", description: "写入公式的目标表" },
          dest_key: { description: "目标表匹配列：表头名、列字母或 0 基列号" },
          lookup_sheet: { type: "string", description: "查找表" },
          lookup_key: { description: "查找表匹配列：表头名、列字母或 0 基列号" },
          return_columns: {
            type: "array",
            items: {},
            description: "要从查找表取回的列（表头名/列字母/0 基列号）。可多列",
          },
          return_column: { description: "只取回一列时可用，等同 return_columns 单元素" },
          dest_start_column: { description: "写入起始列（字母或表头）。省略则追加到已用区域右侧" },
          dest_headers: { type: "array", items: { type: "string" }, description: "新列表头，省略则用查找表原表头" },
          dest_range: { type: "string", description: "目标表数据区域，省略则用已用区域" },
          lookup_range: { type: "string", description: "查找表数据区域，省略则用已用区域" },
          if_not_found: { description: "未匹配时的值，默认空字符串" },
          engine: { type: "string", enum: ["xlookup", "vlookup"], description: "默认 xlookup。返回列在匹配列左侧时不要用 vlookup" },
          dest_has_headers: { type: "boolean", description: "默认 true" },
          lookup_has_headers: { type: "boolean", description: "默认 true" },
        },
        required: ["dest_sheet", "dest_key", "lookup_sheet", "lookup_key"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "split_column",
      description:
        "按分隔符把一列拆成多列写入目标位置（默认表右侧）。拆邮箱、拆逗号分隔字段时用它，不要逐行 read + write。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          column: { description: "要拆的列：表头名、列字母或 0 基列号" },
          delimiter: {
            type: "string",
            description: "分隔符。可用 ',' '@' ';' '，'，或别名 comma / space / tab / whitespace / at",
          },
          max_parts: { type: "number", description: "最多拆成几列；省略则按数据自动检测，上限 8" },
          dest_start_column: { description: "写入起始列。省略则追加到已用区域右侧" },
          dest_headers: { type: "array", items: { type: "string" }, description: "新列表头" },
          has_headers: { type: "boolean", description: "默认 true" },
          trim: { type: "boolean", description: "是否去掉拆出片段首尾空格，默认 true" },
          range: { type: "string", description: "数据区域，省略则用已用区域" },
        },
        required: ["sheet", "column", "delimiter"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "combine_columns",
      description:
        "用 TEXTJOIN 公式把多列拼成一列（可指定分隔符，默认跳过空单元格）。合并姓名/地址等时用它，不要把值读出来再写死。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          columns: { type: "array", items: {}, description: "至少两列：表头名、列字母或 0 基列号" },
          separator: { type: "string", description: "连接符，默认空字符串" },
          dest_header: { type: "string", description: "新列表头" },
          dest_start_column: { description: "写入列。省略则追加到已用区域右侧" },
          dest_column: { description: "dest_start_column 的别名" },
          has_headers: { type: "boolean", description: "默认 true" },
          ignore_empty: { type: "boolean", description: "跳过空单元格，默认 true" },
          range: { type: "string" },
        },
        required: ["sheet", "columns"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "freeze_panes",
      description: "冻结窗格以便滚动时仍看到表头。默认冻结首行。unfreeze=true 取消冻结。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          rows: { type: "number", description: "冻结的行数；未指定行列时默认 1，只指定 columns 时默认 0" },
          columns: { type: "number", description: "冻结的列数，默认 0" },
          cell: { type: "string", description: "从该单元格起冻结其上方和左侧，如 'B2'" },
          unfreeze: { type: "boolean", description: "为 true 时取消冻结" },
        },
        required: ["sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "clean_column",
      description:
        "清洗一列：去首尾/多余空格、把 ¥1500/800元/1,200 转成数字、把 2024/3/8、3月12日、2024.03.15 转成 yyyy-mm-dd、全角转半角、去掉所有空格（「张 三」→「张三」）、英文大小写、金额转人民币大写（壹仟贰佰元整）。脏数据、日期归一、金额文本转数字时用它，不要逐行 read + write。默认原地覆盖；给 dest_start_column 则写到新列。rmb_upper 默认写到右侧新列，不覆盖原金额。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          column: { description: "要清洗的列：表头名、列字母或 0 基列号" },
          operations: {
            type: "array",
            items: { type: "string", enum: ["trim", "to_number", "to_date", "to_halfwidth", "remove_spaces", "upper", "lower", "rmb_upper"] },
            description: "按顺序执行。省略则只 trim。可组合，如 [\"to_halfwidth\",\"trim\",\"to_number\"]。to_halfwidth=全角字母数字符号转半角；remove_spaces=删除所有空白含中间的；rmb_upper=数字金额转人民币大写",
          },
          operation: { type: "string", description: "只做一种操作时可用" },
          dest_start_column: { description: "写入新列的起始列。省略且 inplace 不为 false 则原地覆盖" },
          dest_header: { type: "string", description: "新列表头；原地清洗默认保持原表头" },
          inplace: { type: "boolean", description: "默认 true" },
          has_headers: { type: "boolean", description: "默认 true" },
          default_year: { type: "number", description: "无年份日期（如 3月12日）的默认年；省略则用同列其他日期的年份" },
          range: { type: "string" },
        },
        required: ["sheet", "column"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "convert_to_values",
      description:
        "把区域内的公式替换成当前显示值（粘贴为值）。查找匹配、填好公式后用户说「转成值 / 不要公式」时用它。可撤销。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "A1 区域；省略则用已用区域" },
        },
        required: ["sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "polish_range",
      description:
        "一键把区域整理成可读的表：表头加粗配色、边框、列宽自适应、冻结首行，并按列名给金额/数量/日期套数字格式。用户说「整理格式 / 做得好看一点 / 冻结表头并加边框」时用它，不要连调多次 format_range。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "省略则用已用区域" },
          header: { type: "boolean", description: "是否格式化首行为表头，默认 true" },
          freeze_header: { type: "boolean", description: "是否冻结表头行，默认 true" },
          border: { type: "boolean", description: "默认 true" },
          autofit: { type: "boolean", description: "默认 true" },
          number_formats: { type: "boolean", description: "按列表头自动套数字格式，默认 true" },
          header_fill: { type: "string", description: "表头背景色，默认 '#107C41'" },
          header_font_color: { type: "string", description: "表头文字色，默认 '#FFFFFF'" },
        },
        required: ["sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "delete_blank_rows",
      description:
        "删除区域内整行为空的行（默认不删表头）。脏数据里的空行用它，不要逐行 delete_rows。会请求确认；可用任务检查点回退。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "省略则用已用区域" },
          has_headers: { type: "boolean", description: "默认 true，第一行即使为空也不删" },
        },
        required: ["sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "remove_duplicates",
      description:
        "按 Excel「删除重复项」去掉重复行，保留第一次出现的行。用户说去重时用它，不要自己扫描再逐行 delete_rows。会请求确认；可用任务检查点回退。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "含表头的区域；省略则用已用区域" },
          columns: {
            type: "array",
            items: {},
            description: "用来判断重复的列：区域内 0 基列号或列字母（如 'A'）。省略则所有列都参与",
          },
          has_headers: { type: "boolean", description: "第一行是否为表头，默认 true" },
        },
        required: ["sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "rename_sheet",
      description: "重命名一个 Sheet。",
      parameters: {
        type: "object",
        properties: {
          name: { type: "string", description: "当前名称" },
          new_name: { type: "string", description: "新名称，最长 31 字符" },
        },
        required: ["name", "new_name"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "copy_sheet",
      description: "复制整个 Sheet（含格式、公式、图表）。需要整表副本时用它，不要单元格级复制。",
      parameters: {
        type: "object",
        properties: {
          name: { type: "string", description: "源 Sheet 名称" },
          new_name: { type: "string", description: "副本名称，省略则由 Excel 自动命名" },
        },
        required: ["name"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "column_stats",
      description:
        "在 Excel 里计算各列统计：非空、空白、数值个数、合计/最小/最大/平均；行数不太多时还给出唯一值个数。分析大表时优先用它，不要 read_range 整表。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "含表头的区域；省略则用已用区域" },
          columns: {
            type: "array",
            items: {},
            description: "要统计的列：表头名、列字母或 0 基列号。省略则统计前 24 列",
          },
          has_headers: { type: "boolean", description: "第一行是否为表头，默认 true" },
        },
        required: ["sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "apply_conditional_format",
      description:
        "给区域加条件格式：重复值标色、色阶、数据条、大于/小于/等于/介于。用户说标黄重复项、热力图时用它，不要逐格 format_range。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string" },
          rule: {
            type: "string",
            enum: ["duplicate", "unique", "greater_than", "less_than", "equal", "between", "color_scale", "data_bar"],
            description: "默认 duplicate",
          },
          value: { description: "greater_than/less_than/equal/between 的比较值" },
          value2: { description: "between 的上限" },
          fill_color: { type: "string", description: "填充色，如 '#FFEB9C'" },
          font_color: { type: "string" },
        },
        required: ["sheet", "range"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "create_chart",
      description: "根据数据区域创建图表。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "数据所在 Sheet" },
          data_range: { type: "string", description: "含表头的数据区域，如 'A1:B13'" },
          chart_type: {
            type: "string",
            enum: ["ColumnClustered", "Line", "Pie", "BarClustered", "Area", "XYScatter", "Doughnut"],
          },
          title: { type: "string", description: "图表标题" },
          position_start_cell: { type: "string", description: "图表左上角放置位置，如 'G2'" },
        },
        required: ["sheet", "data_range", "chart_type"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "list_external_files",
      description:
        "列出用户附加到任务窗格的外部文件（Excel/CSV）及其编号（文件一、文件二…）、原文件名和工作表结构（表头、行数、样本）。当用户提到当前工作簿之外的文件/数据，或说「文件一」「文件2」时先调用它。",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
  {
    type: "function",
    function: {
      name: "read_external_sheet",
      description:
        "读取附加的外部文件中某个工作表的数据（单次上限 5000 单元格）。可以按行读取，也可以用 range 精确读取一个区域（例如用户在预览器中框选的区域）。",
      parameters: {
        type: "object",
        properties: {
          file: { type: "string", description: "外部文件编号或文件名，如「文件一」「文件1」或原文件名（可模糊匹配）" },
          sheet: { type: "string", description: "工作表名，省略则取第一个" },
          range: { type: "string", description: "A1 表示法区域，如 'A2:D50'；Ctrl 多选时为 'A2:A50,C2:C50'。提供时优先于 start_row/max_rows" },
          start_row: { type: "number", description: "起始行（从 1 开始），默认 1" },
          max_rows: { type: "number", description: "返回行数，默认 100，最大 1000" },
        },
        required: ["file"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "import_external_sheet",
      description:
        "把附加的外部文件中的一个工作表（或其中一个区域）导入当前工作簿，成为新 Sheet（自动避免重名）。导入后即可用常规工具处理。这是跨文件操作的推荐方式。",
      parameters: {
        type: "object",
        properties: {
          file: { type: "string", description: "外部文件编号或文件名，如「文件一」「文件1」或原文件名（可模糊匹配）" },
          sheet: { type: "string", description: "工作表名，省略则取第一个" },
          range: { type: "string", description: "只导入该 A1 区域，如 'A2:D50' 或 Ctrl 多选的 'A2:A50,C2:C50'；省略则导入整表" },
          target_sheet_name: { type: "string", description: "导入后的 Sheet 名，省略则自动生成" },
        },
        required: ["file"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "export_file",
      description:
        "把当前工作簿导出为下载文件。用户说「导出/下载/另存为/发给我一份」时用它。xlsx 且不指定单个区域时，导出整个工作簿原件（含格式、图表、透视）；只要一张表或 CSV 时再指定 sheet/format。插件不能直接覆盖磁盘上的原路径，导出会触发浏览器下载，请让用户另存。",
      parameters: {
        type: "object",
        properties: {
          format: { type: "string", enum: ["xlsx", "csv"], description: "默认 xlsx" },
          sheet: {
            type: "string",
            description: "Sheet 名；省略或 all 表示整个工作簿（xlsx）。CSV 只能单表，省略则用活动表",
          },
          range: { type: "string", description: "只导出该 A1 区域；省略则整张表或整个工作簿" },
          filename: { type: "string", description: "下载文件名，可省略" },
        },
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "inspect_range",
      description:
        "查看指定区域（或当前选区）的值、公式和错误（#N/A、#REF! 等）。解释公式、排查报错、修公式前先调用它。省略 sheet/range 则查看用户当前选区。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "Sheet 名称，省略则用活动表" },
          range: { type: "string", description: "A1 区域，省略则用当前选区" },
        },
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "list_tables",
      description: "列出工作簿里所有 Excel 表（ListObject）及其所在 Sheet、区域和列名。",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
  {
    type: "function",
    function: {
      name: "create_table",
      description: "把一块数据区域转成 Excel 表（带表头、自动筛选）。用户说「转成表格」「做成表」时用它，不要只用格式化冒充。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "含表头的数据区域；省略则用该 Sheet 的已用区域" },
          name: { type: "string", description: "表名，省略则自动生成" },
          has_headers: { type: "boolean", description: "第一行是否为表头，默认 true" },
        },
        required: ["sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "apply_autofilter",
      description:
        "对区域启用自动筛选并按某一列过滤。values 为要保留的取值列表；criterion1 用于自定义条件（如 '>100'、'*北京*'）。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "含表头的区域；省略则用已用区域" },
          column_index: { type: "number", description: "区域内的列偏移，0 表示第一列" },
          values: { type: "array", items: { type: "string" }, description: "要显示的取值" },
          criterion1: { type: "string", description: "自定义条件，如 '>100' 或 '*关键字*'" },
          criterion2: { type: "string" },
          filter_on: { type: "string", enum: ["values", "custom"], description: "默认 values" },
        },
        required: ["sheet", "column_index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "clear_autofilter",
      description: "清除或移除指定 Sheet 上的自动筛选。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          remove: { type: "boolean", description: "true 则完全移除筛选，false 只清条件，默认 false" },
        },
        required: ["sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "create_pivot_table",
      description:
        "根据源数据区域创建透视表，默认写到新 Sheet。行/列/值字段必须是源表表头里的真实列名。用户说「做透视」「交叉汇总」时用它，不要手写 SUMIF 来冒充透视。",
      parameters: {
        type: "object",
        properties: {
          source_sheet: { type: "string" },
          source_range: { type: "string", description: "含表头的源数据区域，如 'A1:F200'" },
          rows: { type: "array", items: { type: "string" }, description: "行字段（列名）" },
          columns: { type: "array", items: { type: "string" }, description: "列字段（列名）" },
          values: {
            type: "array",
            description: "值字段。元素可以是列名字符串，或 {field, aggregation}，aggregation 为 sum/count/average/max/min",
            minItems: 1,
            items: {
              anyOf: [
                { type: "string" },
                {
                  type: "object",
                  properties: {
                    field: { type: "string" },
                    aggregation: { type: "string", enum: ["sum", "count", "average", "max", "min"] },
                  },
                  required: ["field"],
                },
              ],
            },
          },
          name: { type: "string", description: "透视表名称" },
          target_sheet: { type: "string", description: "目标 Sheet，省略则新建" },
          target_cell: { type: "string", description: "透视表左上角，默认 A3" },
        },
        required: ["source_sheet", "source_range", "values"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "group_summary",
      description:
        "按一列或多列分组，对整张表算合计/计数/平均/最大/最小/去重计数，结果写到新 Sheet 的普通表（静态值，末尾带合计行）。插件分块读全表计算，不受 5000 单元格抽样限制。「每个地区/每月/每个销售的金额、单数」这类汇总用它，不要 read_range 抽样后自己算，也不要逐个写 SUMIF。按月/季度/年汇总日期列时给 date_column + date_part。需要可拖拽刷新的交叉表才用 create_pivot_table。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "源数据 Sheet" },
          range: { type: "string", description: "含表头的源区域，如 A1:F500 或 A:F；省略则用整张表的已用区域" },
          group_by: { type: "array", items: {}, description: "分组列（表头名、列字母或 0 基列号），按顺序成为结果的前几列" },
          date_column: { description: "按日期分段的列。若也在 group_by 里则原位替换成分段值，否则放在最前" },
          date_part: { type: "string", enum: ["year", "quarter", "month", "week", "day"], description: "日期分段粒度，默认 month；week 为 ISO 周" },
          metrics: {
            type: "array",
            description: "要算的指标，省略则只统计每组行数。客单价这类比值：先汇总金额合计和行数，再在结果表用公式相除",
            items: {
              type: "object",
              properties: {
                column: { description: "指标列；agg=count_rows 时省略" },
                agg: { type: "string", enum: ["sum", "count", "count_rows", "average", "min", "max", "count_distinct"] },
                header: { type: "string", description: "结果列名，默认「列名_合计」这类" },
              },
              required: ["agg"],
            },
          },
          sort: { type: "string", enum: ["key", "desc", "asc"], description: "key=按分组值升序（默认）；desc/asc=按第一个指标降序/升序" },
          total_row: { type: "boolean", description: "末尾加合计行，默认 true" },
          parse_text_numbers: { type: "boolean", description: "把「¥1,200」「800元」这类文本数字也计入，默认 false（只计真数字，并在 warnings 报告跳过个数）" },
          has_headers: { type: "boolean", description: "默认 true" },
          dest_sheet: { type: "string", description: "结果 Sheet 名，默认「源表_汇总」；重名自动加序号，不会覆盖" },
        },
        required: ["sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "reconcile_tables",
      description:
        "两表全量对账：按一个或多个键（如 订单号+行号）匹配，找出仅在左表、仅在右表、键相同但比对列不一致的记录，写到一张新 Sheet（「状态」列可筛选，数值列带差额 = 左 − 右，可设误差）。键比较忽略首尾空格和全角半角；同一个键出现多次时数值按键求和。「两表核对 / 对账 / 找差异 / 哪些单没对上」用它，不要把两张表读进对话逐行比。两表完全一致时不建新表。",
      parameters: {
        type: "object",
        properties: {
          left_sheet: { type: "string", description: "左表（通常是本方账、系统导出）" },
          right_sheet: { type: "string", description: "右表（对方账、银行流水等）" },
          left_range: { type: "string", description: "左表区域，省略则用已用区域" },
          right_range: { type: "string", description: "右表区域，省略则用已用区域" },
          keys: { type: "array", items: {}, minItems: 1, description: "左表的匹配键列（表头名/列字母/0 基列号），多列时组合成一个键" },
          right_keys: { type: "array", items: {}, description: "右表对应的键列，顺序与 keys 一致；省略则按 keys 的表头名在右表中找" },
          compare_columns: { type: "array", items: {}, description: "要比对的左表列（金额、数量、状态…）；省略则只看两边有没有这条" },
          right_compare_columns: { type: "array", items: {}, description: "右表对应的比对列，顺序与 compare_columns 一致；省略则按表头名找" },
          tolerance: { type: "number", description: "数值允许的误差，如 0.01；默认 0" },
          ignore_case: { type: "boolean", description: "键和文本比对忽略大小写，默认 false" },
          include_matched: { type: "boolean", description: "结果里也列出完全一致的记录，默认 false" },
          has_headers: { type: "boolean", description: "默认 true" },
          dest_sheet: { type: "string", description: "结果 Sheet 名，默认「对账结果」" },
        },
        required: ["left_sheet", "right_sheet", "keys"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "stack_sheets",
      description:
        "把多张结构相近的表（各月、各门店、各导入文件）按表头名对齐后纵向追加成一张新表：列顺序不同也能对齐，某表缺的列留空，默认最前加一列「来源」（Sheet 名）。原表不动，结果为静态值。附加文件先 import_external_sheet 导入再合并；按键横向匹配请用 add_lookup_column。",
      parameters: {
        type: "object",
        properties: {
          sheets: { type: "array", items: { type: "string" }, minItems: 2, description: "要合并的 Sheet，按此顺序追加" },
          columns: { type: "array", items: { type: "string" }, description: "只输出这些列并按此顺序；省略则取所有表头的并集（按首次出现顺序）" },
          source_column: { type: "string", description: "来源列表头，默认「来源」；传空字符串则不加" },
          has_headers: { type: "boolean", description: "默认 true" },
          dest_sheet: { type: "string", description: "结果 Sheet 名，默认「合并结果」" },
        },
        required: ["sheets"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "split_by_column",
      description:
        "按某一列的取值把表拆成多张新 Sheet（每个取值一张，带表头和数字格式，写入的是值不带公式），返回每张表的行数。原表不动；空值归入「(空白)」。「按地区/部门/类别拆表」用它，不要循环 create_sheet + copy_range。取值超过 max_sheets 会报错并给出取值数，先问用户再调大。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          column: { description: "分类列：表头名、列字母或 0 基列号" },
          range: { type: "string", description: "源区域，省略则用已用区域" },
          name_prefix: { type: "string", description: "新 Sheet 名前缀，默认无（Sheet 名就是分类值）" },
          max_sheets: { type: "integer", description: "最多新建多少张，默认 30，上限 100" },
          has_headers: { type: "boolean", description: "默认 true" },
        },
        required: ["sheet", "column"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "suggest_followups",
      description:
        "任务基本完成后调用，给出 2～3 个用户可一点就继续的下一步（做成图表、做透视、再清洗等）。在最终文字回复之前调用。简单问答不要调用。",
      parameters: {
        type: "object",
        properties: {
          actions: {
            type: "array",
            minItems: 2,
            maxItems: 3,
            items: {
              type: "object",
              properties: {
                label: { type: "string", description: "按钮上的短文案，不超过 8 个字" },
                prompt: { type: "string", description: "用户点下去之后发给你的完整指令，要具体、可执行" },
              },
              required: ["label", "prompt"],
            },
          },
        },
        required: ["actions"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "remember",
      description:
        "把一条长期约定写入记忆文件。全局习惯跨工作簿生效；workbook 只对当前文件有效。同一件事用相同 key 更新。用户说「记住」「以后都…」、或你发现稳定约定（关联键、输出位置、列含义）时调用。不要记一次性任务或密码/API Key。",
      parameters: {
        type: "object",
        properties: {
          content: { type: "string", description: "一句完整的约定，例如「销售表用订单号和 Sheet2 关联」" },
          scope: { type: "string", enum: ["global", "workbook"], description: "默认 global" },
          key: { type: "string", description: "稳定短键，如 join_key / output_sheet，重复写入会覆盖" },
        },
        required: ["content"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "forget",
      description: "从记忆文件删除一条。用户说「忘掉…」「不要再按这个做」时调用。",
      parameters: {
        type: "object",
        properties: {
          id: { type: "string", description: "记忆 id" },
          key: { type: "string", description: "记忆 key，与 id 二选一" },
        },
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "list_memories",
      description: "列出当前记忆文件。通常不必调用，因为每轮消息已附带记忆；只在需要核对 id/key 再 forget 时用。",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
  {
    type: "function",
    function: {
      name: "save_skill",
      description:
        "把一套可复用的工作流程沉淀为技能，下次相似任务会自动检索并注入。用户说「记住这个做法/以后都这样做」，或你完成了一个值得复用的多步流程时调用。steps 写成工具级泛化步骤（不含具体文件名/单元格值）。同名技能会被更新。技能需用户在技能库确认后才生效，调用后请告诉用户去确认。",
      parameters: {
        type: "object",
        properties: {
          name: { type: "string", description: "技能名，≤12 字，如「跨表订单合并」" },
          description: { type: "string", description: "一句话适用场景" },
          triggers: { type: "array", items: { type: "string" }, description: "3~6 个触发关键词" },
          steps: { type: "string", description: "编号的步骤与注意点" },
        },
        required: ["name", "steps"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "list_skills",
      description: "列出技能库中的所有技能（名称、场景、触发词）。相关技能已自动注入上下文，通常不必调用；用户问「你会什么」或要管理技能时用。",
      parameters: { type: "object", properties: {}, required: [] },
    },
  },
  {
    type: "function",
    function: {
      name: "ask_user_choice",
      description:
        "向用户展示 2～4 个可选方向并等待点选。任务有多种合理做法、目标不够明确、或即将做较大改动时先调用；用户点选后再执行。简单明确的指令不要调用。选项必须具体、可区分，并尽量基于已看到的真实 Sheet/列名。",
      parameters: {
        type: "object",
        properties: {
          title: { type: "string", description: "简短标题，例如「你想怎么处理这张销售表？」" },
          question: { type: "string", description: "一两句背景说明，帮助用户理解为什么要选" },
          options: {
            type: "array",
            minItems: 2,
            maxItems: 4,
            items: {
              type: "object",
              properties: {
                id: { type: "string", description: "稳定的短 id，如 summarize / clean / chart" },
                label: { type: "string", description: "选项标题，几个字说清方向" },
                description: { type: "string", description: "这一步具体会做什么、结果在哪" },
                recommended: { type: "boolean", description: "是否为你最推荐的方向，最多标一个" },
              },
              required: ["label"],
            },
          },
        },
        required: ["title", "options"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "audit_data",
      description:
        "数据体检：扫描指定区域（默认活动 Sheet 已用区域），检测缺失值、重复行、类型不一致、数值离群值、首尾空格五类问题，生成「数据体检报告」新 Sheet，并在界面给出可勾选的问题卡片。用户说「体检 / 审计 / 检查数据质量」时用它。体检只报告不修改原表，等用户勾选后你再修复。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "Sheet 名称，省略则用活动表" },
          range: { type: "string", description: "A1 区域，省略则用整个已用区域" },
          header_row: { type: "number", description: "表头所在行号（从 1 开始），默认 1" },
        },
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "diagnose_errors",
      description:
        "扫描整个 Sheet（或指定区域）的公式错误（#N/A、#REF!、#DIV/0! 等），按「错误类型 + 公式模式」聚类，返回每簇的数量、代表地址、代表公式和受影响范围。用户要排查整表错误时用它（选区级排查用 inspect_range）。拿到结果先解释根因和修复方案，等确认后再 set_formulas 修复。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "Sheet 名称，省略则用活动表" },
          range: { type: "string", description: "A1 区域，省略则用整个已用区域" },
        },
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "run_script",
      description:
        "在隔离沙箱里运行一段 JavaScript 计算，专门处理其他专用工具和 Excel 公式都做不到的逻辑：复杂文本解析/正则提取、模糊匹配、多条件规则打标、自定义排名或分摊、跨多表的复杂合并、生成序列等。" +
        "插件把 inputs 指定的区域完整读入（不抽样、不进对话），脚本返回结果；带 output 时结果直接写表。沙箱不能联网、不能访问 Excel 对象，只能纯计算。" +
        "能用 group_summary / reconcile_tables / add_lookup_column / clean_column 等专用工具或公式完成的不要用它。",
      parameters: {
        type: "object",
        properties: {
          code: {
            type: "string",
            description:
              "JS 函数体（不要写 function 包裹），必须 return 结果。可用变量：inputs.<name> 为二维数组（含表头行；日期是 Excel 序列号，数字是数字）；" +
              "helpers.toObjects(grid) 首行作表头转对象数组；helpers.toGrid(objs, headers?) 转回带表头二维数组；helpers.num(v) 把「¥1,200」「800元」「12%」转数字（失败为 null）；" +
              "helpers.excelDate(serial) 序列号转 yyyy-mm-dd；helpers.toSerial(text) 反向；log(...) 调试输出（会随结果返回）。不能用 import、fetch、eval、setTimeout。",
          },
          inputs: {
            type: "array",
            items: {
              type: "object",
              properties: {
                name: { type: "string", description: "脚本里的变量名，如 orders" },
                sheet: { type: "string" },
                range: { type: "string", description: "A1 区域，省略则读整个已用区域" },
              },
              required: ["name", "sheet"],
            },
            description: "要读入的区域，最多 8 个；不读表的纯计算可省略",
          },
          output: {
            type: "object",
            properties: {
              new_sheet: { type: "string", description: "写到新建 Sheet（重名自动加后缀），推荐" },
              sheet: { type: "string", description: "写到已有 Sheet（与 new_sheet 二选一）" },
              start_cell: { type: "string", description: "配合 sheet 使用，默认 A1；覆盖已有内容会请求确认" },
              header: { type: "boolean", description: "new_sheet 时首行是否作表头（加粗冻结），默认 true" },
            },
            description: "省略则把返回值交给你查看（数组最多返回前 100 项）；要落表时返回二维数组或对象数组",
          },
          purpose: { type: "string", description: "一句话说明这段脚本做什么，用于操作记录" },
          timeout_seconds: { type: "number", description: "超时秒数，默认 15，最大 60" },
        },
        required: ["code"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "set_data_validation",
      description:
        "设置数据验证：下拉列表、整数/小数/日期/文本长度范围、自定义公式校验，或清除验证。用户说「做个下拉框 / 只能填 1～100 / 限制日期 / 防止重复录入」时用它。只约束今后的输入，不修改已有值。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "要加验证的区域，如 'C2:C500'（通常不含表头）" },
          type: { type: "string", enum: ["list", "whole_number", "decimal", "date", "text_length", "custom", "clear"] },
          items: { type: "array", items: { type: "string" }, description: "type=list：选项，如 [\"已完成\",\"进行中\"]；选项多或含逗号时改用 source" },
          source: { type: "string", description: "type=list：选项来源区域，如 '=参数!$A$2:$A$30'" },
          operator: {
            type: "string",
            enum: ["between", "not_between", "equal", "not_equal", "greater_than", "less_than", "greater_or_equal", "less_or_equal"],
            description: "数值/日期/长度类的比较方式，默认 between",
          },
          min: { description: "between 的下限，或单边比较的值；日期写 yyyy-mm-dd；也可写 =公式" },
          max: { description: "between 的上限" },
          formula: { type: "string", description: "type=custom：返回 TRUE 才允许输入的公式，相对区域左上角书写，如 =COUNTIF($A:$A,A2)=1" },
          allow_blank: { type: "boolean", description: "是否允许留空，默认 true" },
          reject_invalid: { type: "boolean", description: "true（默认）拒绝无效输入；false 只警告" },
          error_message: { type: "string", description: "输入无效时的提示文字" },
          prompt_message: { type: "string", description: "选中单元格时显示的输入提示" },
          in_cell_dropdown: { type: "boolean", description: "type=list 时是否显示下拉箭头，默认 true" },
        },
        required: ["sheet", "range", "type"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "set_visibility",
      description: "隐藏或取消隐藏行、列或整张工作表。",
      parameters: {
        type: "object",
        properties: {
          target: { type: "string", enum: ["rows", "columns", "sheet"] },
          sheet: { type: "string" },
          range: { type: "string", description: "rows 写行号如 '3:10'，columns 写列字母如 'C:F'；target=sheet 时省略" },
          hidden: { type: "boolean", description: "true 隐藏（默认），false 取消隐藏" },
        },
        required: ["target", "sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "unpivot_columns",
      description:
        "宽表转长表（逆透视）：把「1月、2月…12月」「各门店」这类横着排的多列转成两列「属性 + 值」，每个原始行 × 每个值列展开成一行，标识列（姓名、产品…）原样带上，结果写到新 Sheet。用户说「转成一维表 / 逆透视 / 列转行 / 宽表变长表 / 方便做透视」时用它，不要 read_range 后手工拼。省略 value_columns 时，除 id_columns 外的列都展开。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "含表头的源区域，省略则用已用区域" },
          id_columns: { type: "array", items: {}, description: "保持不动的标识列（表头名/列字母/0 基列号），如 [\"产品\",\"地区\"]" },
          value_columns: { type: "array", items: {}, description: "要转成行的列，如 [\"1月\",\"2月\",\"3月\"]；省略则为除 id_columns 外的所有列" },
          attribute_header: { type: "string", description: "存放原列名的新列表头，默认「属性」，如「月份」" },
          value_header: { type: "string", description: "存放数值的新列表头，默认「值」，如「销售额」" },
          skip_blanks: { type: "boolean", description: "值为空的格不生成行，默认 true" },
          dest_sheet: { type: "string", description: "结果 Sheet 名，默认「源表_长表」；重名自动加序号" },
        },
        required: ["sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "fill_blank_cells",
      description:
        "取消合并单元格并把每格补上原来的值；或把指定列的空白格按上方（或左侧、固定值）填满。用户说「取消合并并填充 / 合并单元格拆开 / 空白处填上面的值 / 空格补 0」时用它，不要逐格 write_range。只补竖向跨多行的合并（分类、部门这类），单行横向合并的标题保持原样；整行为空的行不填。会请求确认，可撤销（合并状态需用任务检查点回退）。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "含表头的区域，省略则用已用区域" },
          columns: { type: "array", items: {}, description: "要填空白的列（表头名/列字母/0 基列号），只处理与这些列相交的合并区域。省略则只拆开并补齐区域里的合并单元格，不动普通空白格" },
          mode: { type: "string", enum: ["above", "left", "value"], description: "空白格取值方式：above=上方最近的值（默认）；left=左侧的值；value=固定值" },
          value: { description: "mode=value 时填入的值，如 0 或 \"无\"" },
          unmerge: { type: "boolean", description: "是否取消合并并补齐合并区域，默认 true" },
          has_headers: { type: "boolean", description: "第一行是表头（不填充），默认 true" },
        },
        required: ["sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "extract_text",
      description:
        "从一列文本里提取信息写到右侧新列（静态值）：手机号、邮箱、身份证号、数字/金额、日期、网址、中文、英文；由身份证号推出出生日期、性别、周岁年龄并做校验位检查（存成数字的身份证会提示精度已丢失）。「从备注里提取手机号 / 根据身份证算年龄性别 / 把地址里的数字抠出来」用它，不要读进对话手抄。自定义正则或复杂规则改用 run_script。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          column: { description: "源列：表头名、列字母或 0 基列号" },
          targets: {
            type: "array",
            items: {
              type: "string",
              enum: ["mobile", "email", "id_card", "number", "date", "url", "chinese", "english", "birth_date", "gender", "age", "id_check"],
            },
            description: "要提取什么，每项生成一列，按顺序排列，如 [\"birth_date\",\"gender\",\"age\"]",
          },
          all_matches: { type: "boolean", description: "一格里有多个匹配时全部取出并用 separator 连接，默认只取第一个" },
          separator: { type: "string", description: "all_matches 时的连接符，默认「、」" },
          dest_start_column: { description: "写入起始列（字母或表头）。省略则追加到已用区域右侧" },
          dest_headers: { type: "array", items: { type: "string" }, description: "新列表头，省略则用「手机号」「出生日期」这类名称" },
          has_headers: { type: "boolean", description: "默认 true" },
          range: { type: "string", description: "数据区域，省略则用已用区域" },
        },
        required: ["sheet", "column", "targets"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "compare_sheets",
      description:
        "比较同一张表的两个版本（改前/改后、上月/本月、两份导出）逐格找出改了哪些值、新增和删除了哪些行，差异清单写到新 Sheet（变化、行、列、原值、新值、位置）；highlight=true 时在新版本里把改过的格标黄、新增行标绿。给 keys（如 工号）按键对齐行，行顺序变了、插入删除行也能比；不给则按行位置比。列按表头名对齐。附加文件先 import_external_sheet 导入。按键核对金额汇总差额用 reconcile_tables。",
      parameters: {
        type: "object",
        properties: {
          old_sheet: { type: "string", description: "旧版本 Sheet" },
          new_sheet: { type: "string", description: "新版本 Sheet" },
          old_range: { type: "string", description: "旧版本区域，省略则用已用区域" },
          new_range: { type: "string", description: "新版本区域，省略则用已用区域" },
          keys: { type: "array", items: {}, description: "旧版本里的行键列（表头名/列字母/0 基列号），如 [\"工号\"]；省略则按行位置对齐" },
          new_keys: { type: "array", items: {}, description: "新版本的键列，顺序与 keys 一致；省略则按 keys 的表头名找" },
          tolerance: { type: "number", description: "数值允许的误差，默认 0" },
          ignore_case: { type: "boolean", description: "文本比较忽略大小写，默认 false" },
          highlight: { type: "boolean", description: "在新版本 Sheet 上标色：改过的格黄色、新增行绿色。默认 false（会改动新版本的格式）" },
          has_headers: { type: "boolean", description: "默认 true；false 时按列位置对齐" },
          dest_sheet: { type: "string", description: "差异清单 Sheet 名，默认「版本差异」" },
        },
        required: ["old_sheet", "new_sheet"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "set_page_layout",
      description:
        "设置工作表的打印格式：横向/纵向、纸张、缩放到一页宽或一页内、每页重复打印表头行、打印区域、页边距、水平居中、打印网格线、页眉页脚和「第 X 页，共 Y 页」页码。用户说「打印设置 / 打印时每页都有表头 / 调成一页宽 / 横着打 / 加页码」时用它。返回设置后的状态。",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "省略则用活动表" },
          orientation: { type: "string", enum: ["portrait", "landscape"], description: "portrait=纵向，landscape=横向" },
          paper_size: { type: "string", enum: ["A4", "A3", "A5", "B4", "B5", "Letter", "Legal"] },
          fit_to: { type: "string", enum: ["width", "page", "none"], description: "width=所有列缩到一页宽（行数不限）；page=整表缩到一页；none=取消缩放（100%）" },
          scale: { type: "number", description: "缩放百分比 10～400，与 fit_to 二选一" },
          print_title_rows: { type: "string", description: "每页重复打印的表头行，如 '1' 或 '1:2'" },
          print_title_columns: { type: "string", description: "每页重复打印的左侧列，如 'A' 或 'A:B'" },
          print_area: { type: "string", description: "只打印该区域，如 'A1:H50'" },
          margins: { type: "string", enum: ["normal", "narrow", "wide"], description: "页边距预设：常规 / 窄 / 宽" },
          center_horizontally: { type: "boolean", description: "页面水平居中" },
          gridlines: { type: "boolean", description: "打印网格线" },
          page_numbers: { type: "boolean", description: "页脚中间加「第 X 页，共 Y 页」" },
          header_text: { type: "string", description: "页眉中间的文字，如报表标题；空字符串清除" },
          footer_text: { type: "string", description: "页脚文字；同时开 page_numbers 时放在页脚左侧，否则放中间；空字符串清除" },
        },
        required: [],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "create_sheet_index",
      description:
        "生成目录页：新建（或刷新已有的）「目录」Sheet 放在最前面，列出每张工作表的名称（点一下跳转的超链接）、数据区域和行数。Sheet 多、找表麻烦，用户说「做个目录 / 导航页 / 工作表索引」时用它。",
      parameters: {
        type: "object",
        properties: {
          name: { type: "string", description: "目录 Sheet 名，默认「目录」；同名且是之前生成的目录则原地刷新" },
          include_hidden: { type: "boolean", description: "是否列出隐藏的 Sheet，默认 false" },
        },
        required: [],
      },
    },
  },
];
