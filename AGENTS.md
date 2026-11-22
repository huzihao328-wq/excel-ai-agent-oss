# Agent 须知

本仓库是 **Excel 桌面版 Office.js 任务窗格插件**，不是普通 Web 应用，也不是 VSTO / VBA / xlam。

## 安装 / 卸载 / 打包（强制）

在执行下列任何操作之前，**必须先完整阅读** [`INSTALL.AGENT.md`](INSTALL.AGENT.md)，并严格按其中的步骤和验收标准执行：

- 安装、侧载、注册、开机自启、信任目录
- **公司内网 / 不能访问 npm**：解压离线包后双击 `install.cmd`，或 `runtime\node.exe scripts\install-persistent.js`。不要跑 `npm install`
- 能上网的开发机才用 `npm install` / `npm run install-addin` / `npm start`
- 卸载、排障、换机器、移动项目目录
- `npm run pack`（默认打离线包）或 `npm run pack:src`

**禁止**凭通用 Office 插件经验跳步、改端口、改 `manifest.xml` 里的 `https://localhost:3000`、往 XLSTART 写 `Book.xltx`，或同时跑 `npm start` 与静态服务。内网环境不要要求用户配置 npm 镜像或连外网。

遇到报错时，先对照 `INSTALL.AGENT.md` 的「故障排除」逐条处理，再向用户汇报：已执行的命令、完整报错、已尝试的修复。

## Cursor Cloud specific instructions

Cloud agents run on Ubuntu VMs without Excel desktop. Do not run `npm start`, `npm run install-addin`, `install.cmd`, or any sideload/register steps.

- Install deps: `npm install`
- Build: `npm run build`
- Validate cases: `npm run validate-cases`
- Validate manifest: `npm run validate`

Excel add-in install and sideload are Windows-only; do those on a local machine per `INSTALL.AGENT.md`.

