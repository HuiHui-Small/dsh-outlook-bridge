# 变更记录 — dsh-outlook-bridge

> 运行形态：宿主 `index.js`（Graph/令牌/HTTP 路由）+ 浏览器 `client.js`（`sidebar.panellist` 图标 + `main` 面板）

---

## v0.3 — 2026-11-10（用户侧继续迭代，含 AI 指令栏、欢迎页天气、全部已读、拖拽）

规模：`client.js` 163 KB → 227 KB，`index.js` 64.6 KB → 80.8 KB。目录里留有迭代快照：`client.js.bak-move-tabs / -drag-event / -ai-chatbar / -ai-command-route / -ai-bar-base`、`index.js.bak-ai-chatbar / -ai-command-route`。

### 宿主新增接口

| 接口 | 实现要点 |
|---|---|
| `POST /ai/command` | `parseAiCommand(ctx, text, signal)`：把自然语言指令交给 Harness 默认模型，注入 Asia/Shanghai 的「今天/现在」再解析成结构化动作；90 秒超时、客户端断开即 abort；AI 不可用返回 503（`ai_unavailable`）。 |
| `POST /mail/read-all` | `markFolderRead(folder)`：**先把所有分页收齐**再改（边分页边改会漏邮件），每批 20 条，90 秒超时，校验 `nextLink` 必须属于 Graph 域。 |
| `GET /weather/cities?q=` | Open-Meteo geocoding，返回 `{id,name,latitude,longitude}`，名称拼 `城市 · 省/州 · 国家`。 |
| `GET /weather/today?lat&lon` | Open-Meteo forecast：当前温度 + 天气码 + 当日最低/最高；经纬度做范围校验。 |
| `htmlToPlainText()` | 邮件正文 HTML → 纯文本（去 style/script/head、块级换行、实体解码），配合 `/mail/detail`。 |

### 客户端新增

| 组件/函数 | 作用 |
|---|---|
| `MailWelcome` | 欢迎/空态页：定位或搜索城市显示今日天气，选择记在 localStorage；`savedWelcomeCity`、`welcomeWeatherLabel`、`shanghaiTodayParts` 配套。 |
| `parseAiCommandFallback` | AI 不可用时的本地兜底解析（不依赖模型也能识别简单指令）。 |
| `addDaysYmd` / `addMinutesHm` | 日期加减工具，供指令解析与时间轴用。 |
| 日程拖拽 | 日历里加入 `draggable` / `onDragStart` / `onDrop`（快照名 `bak-drag-event` 对应此功能）。 |
| 标签页顺序调整 | 快照 `bak-move-tabs` 对应页签重排。 |

### 校验
- 两侧 `node --check` 均通过；离线激活壳（已桩 `react-dom`）`apply()` 正常注册 `outlook`（panellist order 21）+ `main`。
- 天气上游实测：`上海 → 31.2,121.5`，`7.3°C / code 3 / 7~14.7°C`，字段解析与宿主一致。

---

## v0.2.0 — 2026-11-09（由用户在 ChatGPT 侧修改后合入）

本版新增 **AI 能力** 与 **OWA 式邮件工作区**，宿主与客户端两侧同时改动（`index.js` 64.6 KB / `client.js` 163.6 KB）。

### 宿主新增

| 位置 | 说明 |
|---|---|
| `fetchMailDetail(id, signal)` + `GET /mail/detail?id=` | 取单封邮件全文。请求 Graph 时带 `Prefer: outlook.body-content-type="text"`；若返回的 `contentType` 不是 `text` 则**主动报错**，不把 HTML 送进模型。`demo-` 前缀走示例数据。 |
| `mailAiStatus(ctx)` + `GET /ai/status` | 探测 Harness 的默认模型：`ctx.get('llm')` 有 `prepareCall` 且 `agentDefaultModel.currentSelection()` 有 provider/model 才可用，否则返回"请先在 Harness 模型设置里配置默认模型"。 |
| `generateMailAi(ctx, input, signal)` + `POST /mail/ai` | `action: read` 解读邮件（邮件概要 / 需要处理 / 建议回复），`action: reply` 起草回复正文。 |
| AI 约束 | system prompt 把邮件标题/发件人/正文全部当作**不可信数据**，禁止执行邮件里的指令、禁止输出验证码/密码/一次性登录码、声明自己无工具且不执行发送、不编造附件或日期；正文超 28000 字截断并告知模型；输出超 16000 字报错；`finish.kind !== 'stop'` 视为失败。 |

### 客户端新增组件

| 组件 | 作用 |
|---|---|
| `MailWorkspace` | OWA 式左栏：`收藏夹`（收件箱/已发送/草稿）+ 账户分组（收件箱/垃圾邮件/草稿/已发送/已删除/存档…），支持折叠与右键菜单；文件夹名带未读数。 |
| `MailReadingPane` | 右侧阅读窗格：加载全文、AI「解读」与「起草回复」、意图输入框、已有草稿时给出建议、内联回复发送、错误提示。 |
| `EventEditor` | 日程编辑器弹窗：时间轴、只允许选可写日历（`canEdit`）、改过未保存时二次确认、Esc/遮罩关闭。 |
| `ContextMenu` / `openContextMenu` | 通过 `window` 自定义事件 `outlook-context-menu` 打开的统一右键菜单（邮件行、文件夹等多处复用）。 |
| `copyOutlookText` | 统一的复制封装，剪贴板不可用时报可读错误。 |
| `MailIcon` / `Overlay` | 邮件图标与遮罩层小工具。 |

### 客户端调用的接口（当前全量）

```
/state  /settings  /snapshot  /folders  /mail  /mail/detail  /mail/ai  /ai/status
/mail/send  /mail/reply  /mail/delete  /mail/move  /mail/flag  /mail/read
/calendar  /calendar/create  /task/complete
/connect/start  /connect/poll  /connect/open  /disconnect
/auth/url  /auth/open  /auth/complete
```

---


### v0.2.1 — 日程编辑 / 删除（2026-11-09 追加）

| 位置 | 变更 |
|---|---|
| 宿主 `createEvent` | 支持 `eventId`：带 id 时改为 `PATCH /me/events/{id}` 并对 id 做正则校验；未带 id 仍是 `POST` 新建（保留 calendarId / attendees / reminder / showAs / transactionId） |
| 宿主 `GET /calendar/event?id=` | 取单条日程完整字段（subject/body/start/end/isAllDay/location/attendees/showAs/reminderMinutesBeforeStart/calendarId），供编辑器回填 |
| 宿主 `POST /calendar/update` | 走同一套校验后 PATCH |
| 宿主 `POST /calendar/delete` | `DELETE /me/events/{id}` |
| 客户端 `EventEditor` | 新增 `onDelete`；标题按 `form.eventId` 显示"编辑日程 / 新建活动"；工具栏多一个「删除日程」 |
| 客户端月视图 | 日程块 **点击即编辑**（`stopPropagation`，不触发新建）；右键菜单新增「编辑日程」；删除带确认 |

实测（Graph 直连）：`POST 201` → `PATCH 200`（标题与时间均生效）→ `DELETE 204`。

---

## v0.1.x — 2026-11-07（首版与当天多轮迭代）

### 初始交付
- `sidebar.panellist` 入口（`outlook`，`order: 21`，紧跟「上下文洞察」），Outlook 蓝信封图标 + 未读角标（>99 显示 `99+`）。
- `main` 面板：左邮箱、右日历 + 待办。
- 宿主 `/outlook/api` 前缀路由，闸门：回环 Host + 同源 Origin + `x-outlook-client: 1` + `ctx.connection.requestRejection()`。
- 演示数据（未连接时可见），令牌存 `$DSH_HOME/outlook-bridge/token.json`（0600）。

### 修复与演进（按发生顺序）
1. **开机崩溃**：`apply()` 在 `ctx.locale` 就绪前访问它 → 现在 `inject: ['slots','locale']` + `ctx.get('locale')` + try/catch，并给 `dsh.client.inject` 加 `@deepseek-ai/dsh-client-locale`。
2. **AADSTS65002**：微软禁止第一方（内置 Office）客户端读个人账户邮件 → 改用自己的 Entra 应用注册（`dsh-outlook-bridge`，多租户 + 个人账户，公共客户端流）。
3. **授权 URL 被截断**：`cmd /c start <url>` 把 `&` 当分隔符，`response_type/scope/redirect_uri` 丢失 → Windows 下改为写 `.url` 快捷方式文件再打开。
4. **设备码在 `consumers` 端点只发 8 位**，输码页不认 → 改用 `common` 的 9 位码。
5. 登录链路改为**宿主后台轮询**（`startAuthPolling` / `resumeAuth`），`pending.json` 落盘，应用重启可续；`auth-log.json` 记录每一步 AAD 返回值。
6. 账户强制 FIDO（Windows Hello），最后一步只能由本人完成。
7. **日历空**：个人账户的真实日程在自定义日历里 → `fetchRange()` 合并全部日历（`/me/calendars` → 逐个 `calendarView`）。
8. **To Do 不可用**：`/me/todo` 对个人账户返回 400 → 标记 `todoUnsupported`，界面改为提示而非报错。
9. 界面迭代：三页签 → 日历独立页 → 界面设置（透明度/模糊度，居中弹窗）→ Outlook 月视图（7×6 网格、今天药丸、左侧日历色块勾选）→ 新建日程 → 写邮件 → 按内容分类（安全/账单/工作/平台/推广/其他）→ 文件夹树 + 重点/其他 + 头像色块 + 分组列表。
10. 权限扩到 `Mail.Send` + `Calendars.ReadWrite`（一次重新授权），实测：发信 `202`、建日程 `201`、删除 `204`。

### 已知边界
- 个人 Microsoft 账户无 To Do Graph 接口。
- 事件编辑器目前只**新建**（宿主仅有 `/calendar/create`，没有 update/delete 路由；删除只能在 Outlook 侧或后续补接口）。
- AI 功能依赖 Harness 默认模型；未配置时按钮会提示去设置。
