# 任务窗格动效约定

本项目使用 React 18、CSS transitions 和浏览器 Web Animations API；不新增动画依赖。业务操作立即执行，视觉收敛由公共组件负责。

## 参数与复用

参数唯一来源是 `src/taskpane/motion.ts`，入口在首帧前安装对应 CSS 变量。

| 层级 | 时长 | 用途 |
| --- | --- | --- |
| 按压 / 微交互 | 100 / 140ms | 按钮、输入框增高、短状态反馈 |
| 退出 / 常规 | 180 / 240ms | 菜单退出、列表 FLIP、分类指示背景 |
| 展开 / 面板 | 280 / 320ms | 详情高度、抽屉、弹窗 |
| 首屏 | 错开 40ms | 欢迎页逐层入场 |

位移统一为 1 / 4 / 8 / 12 / 24px。按压缩放 0.98，小图标 0.94。

### 物理弹簧

进场与松手回弹使用真实阻尼振子（与 SwiftUI 相同的 response / dampingFraction 参数），由 `motion.ts` 采样成 CSS `linear()` 曲线，CSS 通过 `--spring-<name>`（"时长 曲线"，可直接放进 `transition` / `animation` 简写）使用，WAAPI 通过 `SPRING.<name>` 使用。不支持 `linear()` 的旧 WebView 自动退回等价 cubic-bezier。

| 名称 | response / 阻尼 | 约时长 | 用途 |
| --- | --- | --- | --- |
| `smooth` | 0.38 / 1.0 | 480ms | 面板、对话框、消息与卡片入场、展开高度（无过冲） |
| `snappy` | 0.32 / 0.78 | 410ms | 按压回弹、菜单、选中指示、箭头旋转（约 2% 过冲） |
| `bouncy` | 0.40 / 0.66 | 490ms | 图标、状态勾、发送按钮、思考气泡（约 6% 过冲） |
| `list` | 0.28 / 0.88 | 280ms | `MotionList` FLIP 重排、增删、高度 |

退场仍用较短的 ease-in-out（`--presence-exit-duration`），按下 100ms ease-out，松手走弹簧。“回到最新”使用 `springScroll` 弹簧滚动，持续追踪流式输出的底部，滚轮 / 触摸 / 按键立即中断。

- `Presence`：进场、可反向退出、延迟卸载、关闭时 inert；`focus="dialog"` 提供焦点约束和返回，`focus="popover"` 保留非模态操作路径。
- `Collapse`：按真实内容高度展开/收起；覆盖工具详情、技能、模板编辑、操作记录、批量步骤、模型列表、撤销及思考状态。收起编辑区会将焦点返回入口。
- `MotionList`：读取 React 提交前的视觉位置，提交后批量测量目标位置，使用 FLIP 连接增删和重排。同一 key 在退出期间返回会复用节点。删除副本不可交互。
- `MotionTabs`：共用选中背景；分类和预览工作表支持方向键、Home / End，选中状态立即生效。
- 输入框增高和长回复展开从当前尺寸接续动画。流式文本更新保持原有滚动策略，不为每个 token 启动平滑滚动。

## 性能与可访问性边界

- 列表普通内容更新不测量；每个列表最多为 36 个可见条目创建动画。大跨度筛选不逐帧改变整个长列表高度；动画完成后释放资源。
- 恢复长对话只对最近六项应用入场动画；用户、助手消息及工具卡片使用 `memo`，避免随其他消息的 token 更新重复渲染。
- 移除面板背景模糊、公共按钮阴影动画；不用长期 `will-change`。尺寸动画只用于局部展开和有限高度区域。
- 滚动区域预留滚动条位置。输入区域变高时，仅在用户原本停在底部时跟随；翻阅历史时不抢滚动位置。
- 预览鼠标框选按帧合并，相同单元格不重复更新；触屏轻点选择，拖动交给原生平移行为。
- 减少动态效果可在运行时切换：CSS 延迟清零，有限 WAAPI 动画结束，退出副本立即清理。

## 验证方式

生产构建后，设置 `PLAYWRIGHT_MODULE` 为已有 Playwright 模块目录，再运行：

```text
node tests/ui-motion.cjs
node tests/ui-motion-list.cjs
node tests/ui-motion-flows.cjs
node tests/ui-workspace.cjs
node tests/ui-agent-flow.cjs
node tests/ui-stability.cjs
node tests/ui-interaction.cjs
```

`ui-motion-list` 在生产 React 中直接渲染公共组件，测量重排中途反向的连续性。`ui-motion-flows` 运行实际构建页面，覆盖连续开关、键盘、触屏轻点、鼠标框选、尺寸变化、400 条消息、500 条历史记录、300/380/600px 窗口及实时减少动态效果，并输出 `tests/motion-performance.json`。

帧间隔与长任务是本机无头 Edge 下的观测值，测试性能时应单独运行 `ui-motion-flows`。Office 和模型由测试模拟，不能当作真实 Excel WebView、模型响应或不同硬件帧率的证明。当前主机的无头 Edge 对注入的原生触屏滑动没有产生滚动，在独立的普通滚动容器中也可复现；真实触屏惯性滚动仍需设备验收。
