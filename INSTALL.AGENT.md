# Excel AI Agent — Agent 安装说明书

**给要安装本插件的 Agent：先把本文读完，再执行任何命令。**  
不要凭通用 Office 插件、Yeoman、VSTO、VBA 经验跳步。本插件是 **Windows 桌面 Excel** 的 Office.js 任务窗格，页面必须从本机 `https://localhost:3000` 加载。

人类用户可看 [`README.md`](README.md)。安装、排障、换机器、移动目录，以本文为准。

---

## 1. 这是什么 / 不是什么

| 是 | 不是 |
| --- | --- |
| Excel **桌面版**任务窗格 Web 插件 | VSTO、COM、xlam、VBA 宏 |
| 源码用 webpack 打到 `dist/`，由本机 HTTPS 静态服务提供 | 把 `dist/` 直接丢进 Excel 就能用 |
| 清单 `manifest.xml` 固定指向 `https://localhost:3000` | 改端口或改成 `http://` 就能装 |
| 持久安装：开发者注册 + 信任目录 + 开机自启静态服务 | 往 XLSTART 放 `Book.xltx` 当加载器 |

Office.js **只能操作当前打开的工作簿**。跨文件靠任务窗格「附加」Excel/CSV，不是去打开用户磁盘上的另一个 xlsx。

---

## 2. 环境要求（先检查再装）

在项目根目录执行命令。本仓库路径常含空格（如 `E:\excel agent`），**所有路径必须加引号**。

| 项 | 要求 | 怎么确认 |
| --- | --- | --- |
| 操作系统 | Windows 10/11 | `win32`。持久安装脚本用了 PowerShell / SMB / 启动文件夹，不要在 macOS/Linux 上跑 `install-addin` |
| Node.js | **>= 18**（建议 20 LTS） | `node -v` |
| npm | 随 Node 安装 | `npm -v` |
| Excel | 桌面版，支持 ExcelApi 1.9（Microsoft 365 或 2019+） | 能打开桌面 Excel。网页版 Excel 只能手动上传清单，不是本文主路径 |
| 权限 | 当前用户能写 `%APPDATA%`、`%LOCALAPPDATA%`；创建 SMB 共享时会弹 **UAC** | 用户点「是」。拒绝 UAC 时共享会失败，但开发者注册通常仍可用 |
| 网络 | **安装不需要外网、不需要 npm 源**。离线包已含 `dist/`、本地 Office.js、`runtime/node.exe` | 对话功能另算：要能访问用户在设置里填的模型 API，或内网 Ollama |

**不要**用管理员身份长期开 Excel。用当前用户安装即可；只有 SMB 共享那一步会短暂提权。

---

## 3. 两种包，别装错

| 包 | 怎么来 | 目标环境 | 安装方式 |
| --- | --- | --- | --- |
| **离线安装包** `excel-ai-agent-<version>.zip` | 在能上网的开发机 `npm run pack` | **公司内网（默认）** | 解压后双击 `install.cmd`，**不要** `npm install` |
| 源码包 `excel-ai-agent-<version>-src.zip` | `npm run pack:src` | 能访问 npm 的开发机 | `npm install` 再 `npm run install-addin` |

离线包解压后应看到：

```
install.cmd
uninstall.cmd
INSTALL.AGENT.md
manifest.xml
dist/taskpane.html
dist/office-js/office.js
scripts/
runtime/node.exe
```

没有 `runtime/node.exe` 时，目标机需已安装 Node 18+（内网软件中心即可，**不必能连 npm**）。

---

## 4. 内网标准安装（必须按这个走）

Excel 若已打开，先做第 6 节「彻底退出 Excel」。

### 4.1 解压到固定目录

路径不要经常移动（开机脚本会记下这个路径）。路径含空格必须加引号。

### 4.2 双击安装

双击 `install.cmd`。或：

```powershell
cd "解压后的目录"
.\runtime\node.exe scripts\install-persistent.js
```

**不要**执行 `npm install` / `npm run install-addin` / `npx`。内网会卡在拉包。

这一步会：

1. 用包内已构建的 `dist/`（不再 webpack、不再访问 npm）  
2. 放行 WebView localhost 回环  
3. 写注册表注册开发者清单，并关掉 WebView 调试  
4. 拷贝清单到 `%LOCALAPPDATA%\ExcelAIAgent\catalog\`，尽量创建 SMB 共享（**可能弹 UAC，让用户点「是」**）  
5. 写开机启动 VBS，用包内 `runtime\node.exe` 在 `https://localhost:3000` 提供页面  
6. 如需要，生成本机 localhost 证书；若弹出「要安装证书吗」，让用户点「是」

成功时末尾类似：

```
已设置为自动加载。请完全退出 Excel 再重新打开。
之后每次开 Excel，开始选项卡都应有 AI Agent。
```

### 4.3 彻底退出 Excel 再打开

必须结束**所有** `EXCEL.EXE`。再打开桌面 Excel。功能区 **AI Agent** 选项卡应有「打开助手」。窗格应是对话界面，不是白屏。

### 4.4 升级已用很久的机器（保留记忆和技能）

**不要先卸载。** 旧版的记忆、技能、API Key 在 `https://localhost:3000` 的本机存储里；先卸载或清 Office 缓存才可能丢。

1. 把新的离线包解压到那台机器（覆盖原目录，或解压到新目录都可以，清单地址不变）。
2. 双击 `install.cmd`（不要 `npm install`）。
3. 按第 6 节彻底退出 Excel，再打开，点一次 AI Agent。

第一次打开新版本时，会把旧数据写到 `%LOCALAPPDATA%\ExcelAIAgent\userdata\`，之后再换版本也从这里继承。本工作簿约定仍跟 xlsx 走。

### 4.5 配置模型

窗格右上角 ⚙：选服务商 → 填 API Key →「测试连接」→ 保存。  
Key 只存在本机，不要写入仓库。  
内网若也访问不了公有模型 API，改用内网 Ollama / 公司网关，并把 API 地址填成内网地址。

### 4.6 能上网的开发机（补充）

仅当机器能访问 npm 且你拿的是源码时：

```powershell
npm install
npm run install-addin
```

不要把这条当内网步骤。

---

## 5. 验收清单（装完必须做，不要只看退出码）

按顺序勾：

1. `dist/taskpane.html` 和 `dist/office-js/office.js` 都存在。  
2. `https://localhost:3000/taskpane.html` 能打开（证书可接受，不是连接被拒绝）。可用：

   ```powershell
   curl.exe -k -sS -o NUL -w "HTTP=%{http_code}`n" "https://localhost:3000/taskpane.html"
   ```

   期望 `HTTP=200`。  
3. Excel 已完全重启，功能区能看到插件按钮。  
4. 点开任务窗格不是白屏；设置页能打开。  
5. 没有每次启动都弹「文件格式或扩展名无效」（见 §8.6）。

任一项失败：停在「故障排除」，不要重复盲跑 `install-addin`。

---

## 6. 彻底退出 Excel

```powershell
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
```

或让用户关闭所有工作簿后再看任务管理器。只关任务窗格**不够**，清单/信任目录变更要整进程重启才生效。

---

## 7. 卸载

离线包：双击 `uninstall.cmd`，或 `.\runtime\node.exe scripts\install-persistent.js remove`。

开发机源码：`npm run uninstall-addin`。

会移除开机 VBS、尝试 unregister、删信任目录项、尽量删 SMB 共享，并清掉旧版 XLSTART 残留（`Book.xltx`、`ExcelAIAgentWeb.xlam`）。**不会**删除 `%LOCALAPPDATA%\ExcelAIAgent\userdata\`（记忆/技能档案）。升级请走 §4.4，不要先卸载。

然后再彻底退出 Excel。

---

## 8. 故障排除

### 8.1 内网不能访问 npm / `npm install` 失败

这是预期情况。请改用 **离线安装包** + `install.cmd`，不要配 npm 镜像硬连外网。

- 确认解压目录里有 `dist/taskpane.html`、`dist/office-js/office.js`、`runtime/node.exe`。  
- 没有这些文件说明拿到的是源码包，回到能上网的机器执行 `npm run pack`，把新的 zip 拷进来。  
- 路径空格没加引号：`cd` 和命令里的路径加上 `"..."`。

### 8.2 `npm run build` 失败 / `dist/ 不完整`

- 先 `npm install` 成功。  
- `postinstall` 若被跳过（`--ignore-scripts`），手动执行：`node scripts/gen-icons.js`。  
- 报错原文交给用户；不要删 `src` 凑合。

`install-addin` 在没有 `dist/taskpane.html` 时会直接抛错。

### 8.3 证书错误、浏览器锁死 localhost

现象：窗格白屏、Excel 说无法加载加载项、证书不受信任。

离线安装会写当前用户「受信任的根证书颁发机构」。若弹出安全警告，必须点「是」。点「否」后 WebView 会拒收 `https://localhost:3000`。

删掉 `%LOCALAPPDATA%\ExcelAIAgent\certs\` 后重新跑 `install.cmd` 可再生成一次。  
开发机若已装 `office-addin-dev-certs`，服务会优先用那套证书，一般不必再跑 `npx`。

### 8.4 端口 3000 被占用

`manifest.xml` **必须**继续用 `https://localhost:3000`。不要改端口。

本机只允许 **一种** 服务占 3000：

| 进程 | 用途 |
| --- | --- |
| `node scripts/serve-dist.js` | 日常使用（`install-addin` 装的开机服务） |
| `webpack serve`（`npm start` / `npm run dev-server`） | 开发热更新 |

同时开两个时，后启动的会因 `EADDRINUSE` 退出。静态服务遇到占用会**静默退出**，容易误以为「已经在跑」。

查占用：

```powershell
netstat -ano | findstr ":3000"
```

日常给用户用插件：不要跑 `npm start`。改完代码要生效：先 `npm run build`，再关掉任务窗格重新打开（静态服务读 `dist/`，一般不用重启服务）。

正在用 `npm start` 调试：不要再 `build` 指望静态服务，Excel 读的是 webpack 内存包；关窗格再开即可。

### 8.5 功能区没有 AI Agent 按钮

1. 确认跑过 `npm run install-addin` 且 register 没有报错。  
2. **彻底退出**所有 Excel 再开（§6）。  
3. 确认打开的是**桌面 Excel**，不是 Excel 网页版、WPS、或「Excel 在 Microsoft Edge 中」。  
4. 功能区可能被收进 **AI Agent** 自定义选项卡，不只在「开始」。  
5. 文件 → 选项 → 加载项 → COM/Web 加载项里看是否被禁用。  
6. 再执行一次：

   ```powershell
   npx office-addin-dev-settings register manifest.xml
   ```

### 8.6 Excel 启动提示「文件格式或扩展名无效」

旧版安装脚本曾把损坏的 `Book.xltx` 放进 XLSTART。**现行脚本不会再写这个文件。** 若用户机器上还有残留：

```powershell
$book = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART\Book.xltx"
if (Test-Path $book) { Remove-Item -LiteralPath $book -Force }
```

也可检查同目录是否还有 `ExcelAIAgentWeb.xlam`，有则删除。然后彻底重启 Excel。  
**禁止**为了「自动加载」再往 XLSTART 写 xlsx/xltx。

### 8.7 任务窗格白屏 / 一直转圈

1. `curl.exe -k` 测 `https://localhost:3000/taskpane.html` 是否 200。不是 200 → 启动服务：

   ```powershell
   node scripts/serve-dist.js
   ```

   （先保证已 `npm run build`。）  
2. 证书问题见 §8.3。  
3. **WebView localhost 回环**被拦。`npm start` 会先跑 `scripts/ensure-loopback.js`。若仍失败，用**管理员**命令提示符执行：

   ```text
   CheckNetIsolation.exe LoopbackExempt -a -n=Microsoft.win32webviewhost_cw5n1h2txyewy
   ```

4. 离线包必须带 `dist/office-js/office.js`。缺这个文件时窗格会白屏（以前依赖微软 CDN，内网会被拦）。重新 `npm run pack`。  
5. 旧调试器残留：

   ```powershell
   npx office-addin-dev-settings debugging manifest.xml --disable
   ```

### 8.8 UAC / SMB 共享失败

`install-persistent.js` 创建 `\\<主机名>\ExcelAIAgentCatalog` 时可能弹管理员确认。

- 用户点「否」：脚本会警告「创建 SMB 共享失败」，**开发者注册 + 开机服务通常仍够用**。不要为此中止，继续做 §5 验收。  
- 企业策略禁止 SMB：忽略共享，依赖 `office-addin-dev-settings register`。  
- 共享已存在但路径不同：脚本会删掉旧共享再建。失败则手动在「计算机管理 → 共享文件夹」里检查 `ExcelAIAgentCatalog`。

### 8.9 弹出「Allow localhost loopback for Microsoft Edge WebView? (Y/n)」

这是 `office-addin-debugging` 的交互提问，**Agent 环境无法回答 Y/n**。  
日常请用 `npm run install-addin`，不要用 `npm start` 给最终用户装。  
必须开发调试时：`npm start` 已前置 `ensure-loopback.js`，一般不再问。若仍问，先跑 §8.7 的 `CheckNetIsolation`，再 `npm start`。

### 8.10 移动了项目文件夹后插件挂了

开机 VBS 写死了**当时的项目路径**和 **node.exe 路径**。目录一搬，自启服务还指向旧路径。

处理：在**新目录**重新执行：

```powershell
npm install
npm run install-addin
```

不要只改 VBS；以脚本重写为准。

### 8.11 换电脑 / 解压到新机器

必须在新机器重新 `npm install` 和 `npm run install-addin`。  
不要复制旧机器的 `node_modules` 或 `%LOCALAPPDATA%\ExcelAIAgent` 凑合。  
开发证书也是按机器签发的，新机器要重新 `npx office-addin-dev-certs install`。

### 8.12 `office-addin-dev-settings register` 失败

- 先 `npm install`，保证 `office-addin-dev-settings` 在本地。  
- 用 `npx office-addin-dev-settings register manifest.xml`，工作目录必须是清单所在根目录。  
- Excel 正在运行时有时注册了也不显示 → §6 重启。  
- 执行策略过严时，脚本里已用 `-ExecutionPolicy Bypass` 调 PowerShell；不要去改用户全局 ExecutionPolicy。

### 8.13 网页版 Excel / 非 Windows

- **Excel 网页版**：插入 → 加载项 → 上传我的加载项 → 选 `manifest.xml`。本机仍必须有进程在 `https://localhost:3000` 提供 `dist/`，且浏览器能信任该证书。这不是主安装路径。  
- **macOS**：不要跑 `install-persistent.js`。只能按 Office 文档做 sideload，并自己用受信任证书提供清单里的 URL。本文不覆盖。

### 8.14 改代码后界面还是旧的

先分清 3000 上是谁（§8.4）。

- 静态服务：`npm run build`，然后**关掉任务窗格再打开**（Office WebView 会缓存）。  
- webpack：保存后关窗格再开，不要只 `build` 却盯着旧的 `dist/`。

### 8.15 卸载不干净

再跑一次 `npm run uninstall-addin`。手动核对：

- `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\excel-ai-agent-server.vbs` 应不存在  
- `%APPDATA%\Microsoft\Excel\XLSTART\` 下无本插件残留  
- 任务管理器里可结束仍占 3000 的旧 `node.exe`（确认命令行是本项目的 `serve-dist.js` 再杀）

---

## 9. 开发调试（不要和「给用户安装」搞混）

```powershell
npm start
```

会：放行 localhost 回环 → 用 `office-addin-debugging` 拉起 Excel 和 webpack（同样是 3000）。  
**仅开发用。** 和开机静态服务互斥。给用户装机器用 §4，不要用 `npm start` 当安装程序。

停调试：

```powershell
npm stop
```

然后如需恢复日常使用：确保 `dist/` 是新的，再 `node scripts/serve-dist.js` 或重新 `npm run install-addin`。

---

## 10. Agent 禁止事项

- 修改 `manifest.xml` 的 Id、端口、或把 HTTPS 改成 HTTP。  
- 往 XLSTART 生成 `Book.xltx` / 把 xlsx 改后缀当模板。  
- 在内网对离线包执行 `npm install` / `npx`。  
- 在未读本文的情况下执行「网上抄来的」`office-addin-debugging` 交互流程。  
- 把用户的 API Key 写入仓库、聊天记录以外的文件，或提交 `.env`。  
- 打包时带上 `node_modules`、`dist`、本机工作簿、`.zcode`。  
- 安装失败后对用户只说「再试一次」，不附命令和报错原文。

---

## 11. 给 Agent 的最短命令序列

**内网 / 离线包（默认）：**

```powershell
cd "解压后的目录"
.\runtime\node.exe scripts\install-persistent.js
curl.exe -k -sS -o NUL -w "HTTP=%{http_code}`n" "https://localhost:3000/taskpane.html"
```

或让用户双击 `install.cmd`。不要运行 `npm install`。

**能访问 npm 的开发机（源码）：**

```powershell
cd "项目根目录"
npm install
npm run install-addin
curl.exe -k -sS -o NUL -w "HTTP=%{http_code}`n" "https://localhost:3000/taskpane.html"
```

然后请用户彻底退出并打开桌面 Excel，按 §5 验收。  
UAC 和证书信任弹窗必须由用户点击，Agent 无法代点。
