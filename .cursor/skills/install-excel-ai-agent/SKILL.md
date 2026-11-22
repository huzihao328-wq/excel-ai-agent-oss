---
name: install-excel-ai-agent
description: Installs, sideloads, registers, uninstalls, or packs the Excel AI Agent Office add-in on Windows, including company intranet/offline installs that cannot reach npm. Use when the user asks to 安装插件, 侧载, 内网, 离线, install.cmd, install-addin, npm start, uninstall, 打包, pack, or when npm is blocked, the add-in button is missing, localhost:3000 fails, or the certificate is untrusted. MUST read INSTALL.AGENT.md first. On intranet use the offline zip and install.cmd; do not run npm install.
---

# 安装 Excel AI Agent

在跑任何安装相关命令之前，**必须完整阅读**仓库根目录的 `INSTALL.AGENT.md`。

## 硬性要求

- 不要凭通用 Office / Yeoman 插件经验跳步或改端口。
- 不要修改 `manifest.xml` 里的 `https://localhost:3000`。
- 不要往 `%APPDATA%\Microsoft\Excel\XLSTART` 写入 `Book.xltx`。
- 路径含空格必须加引号。
- 安装后必须按文档做验收（证书、页面、Excel 按钮），不要只看命令退出码。
- 失败时按文档「故障排除」处理，再向用户报告命令、原文报错、已尝试的修复。
