/**
 * dsh-outlook-bridge 鈥?browser half.
 *
 * One `sidebar.panellist` entry (an Outlook-blue envelope with an unread badge)
 * placed directly under the context panel, and the `main` page it opens: the mailbox
 * block on the left, calendar and to-do stacked on the right.
 */

window.__ModuleLoader__.load({
  id: 'dsh-outlook-bridge',
  factory(require) {
    const React = require('react')
    const ReactDOM = require('react-dom')
    const h = React.createElement

    const NS = 'outlook-bridge'
    const PANEL_ID = 'outlook'
    const inlineReplyDrafts = new Map()

    /* ------------------------------------------------------------ locales */

    const DICT = {
      zh: {
        panel: 'Outlook',
        title: 'Outlook',
        subtitle: '邮件 · 日历 · 待办',
        'chip.demo': '示例数据',
        'chip.offline': '未连接',
        'action.refresh': '刷新',
        'action.settings': '设置',
        'action.connect': '连接 Outlook',
        'action.disconnect': '断开连接',
        'action.reauth': '重新授权',
        'action.save': '保存',
        'action.close': '关闭',
        'action.open': '在 Outlook 中打开',
        'action.markRead': '标记已读',
        'cat.security': '安全验证',
        'cat.billing': '账单订单',
        'cat.work': '工作事务',
        'cat.social': '平台通知',
        'cat.promo': '推广订阅',
        'cat.other': '其他',
        'mail.cat.all': '全部',
        'mail.compose': '写邮件',
        'mail.to': '收件人（多个用逗号分隔）',
        'mail.subject': '主题',
        'mail.body': '正文',
        'mail.send': '发送',
        'mail.sent': '邮件已发送',
        'mail.folders': '文件夹',
        'mail.focused': '重点',
        'mail.other': '其他',
        'mail.today': '今天',
        'mail.yesterday': '昨天',
        'mail.week': '本周',
        'mail.month': '本月',
        'mail.earlier': '更早',
        'mail.delete': '删除',
        'mail.archive': '存档',
        'mail.toggleRead': '标记已读/未读',
        'mail.title': '邮箱',
        'mail.unread': '{n} 封未读',
        'mail.allRead': '全部已读',
        'mail.all': '全部',
        'mail.unreadOnly': '未读',
        'mail.empty': '收件箱是空的',
        'mail.emptyUnread': '没有未读邮件',
        'mail.attachment': '附件',
        'calendar.title': '日历',
        'calendar.empty': '近期没有日程',
        'calendar.today': '今天',
        'calendar.tomorrow': '明天',
        'calendar.yesterday': '昨天',
        'calendar.allDay': '全天',
        'todo.title': '计划',
        'todo.empty': '没有未完成的计划',
        'todo.summary': '{count} 项未完成',
        'todo.overdue': '已逾期',
        'todo.noOverdue': '没有逾期的计划',
        'todo.all': '所有项目',
        'todo.categories': '分类',
        'todo.categoryFilter': '按分类或计划清单筛选',
        'todo.completed': '已完成',
        'todo.sortCompleted': '已完成计划排序',
        'todo.sortCompletion': '完成时间',
        'todo.sortCreation': '创建时间',
        'todo.sortAscending': '按时间升序排列',
        'todo.stale': '正在显示上次同步的计划。',
        'todo.noCompleted': '暂无已完成的计划',
        'todo.complete': '完成计划：{title}',
        'todo.restore': '恢复计划：{title}',
        'todo.menuComplete': '标记为已完成',
        'todo.menuRestore': '恢复为未完成',
        'todo.menuDue': '截止日期设为{day}',
        'todo.menuClearDue': '清除截止日期',
        'todo.menuCopyTitle': '复制计划标题',
        'todo.menuDelete': '删除计划',
        'todo.dueDate': '截止日期',
        'todo.dragHint': '拖动卡片调整日期或分类，右键查看更多操作',
        'todo.dropNoDate': '放下后清除截止日期',
        'todo.dropDate': '放下后将截止日期设为 {date}',
        'todo.dropCategory': '放下后归入「{name}」',
        'todo.due': '截止 {time}',
        'todo.list': '计划清单',
        'todo.noLists': '暂无可用清单',
        'todo.chooseList': '请选择清单',
        'todo.listMissing': 'Microsoft To Do 找不到所选清单。请刷新或重新选择清单后再添加。',
        'todo.taskMissing': 'Microsoft To Do 找不到这项计划。请刷新列表后重试。',
        'todo.taskGone': 'Microsoft To Do 已找不到这项计划，列表已刷新。',
        'todo.new': '输入计划标题',
        'todo.newWithDate': '输入计划标题，可在末尾加 /月.日 或 /年份.月.日',
        'todo.dateHint': '在标题末尾输入 /月.日 或 /年份.月.日 设置截止日期，省略年份时按今年计算。',
        'todo.add': '添加计划',
        'todo.demo': '正在查看演示计划。连接 Outlook 后可新建计划并同步到 Microsoft To Do。',
        'settings.title': '连接设置',
        'settings.clientId': 'Microsoft 应用（客户端）ID（可选）',
        'settings.clientSecret': '客户端密码（可选，机密客户端用）',
        'settings.clientSecretHint': '如果应用里找不到「允许公共客户端流」，就改走机密客户端：左侧「证书和密码」→ 新客户端密码 → 添加 → 立刻复制那串「值」，粘到这里。留空表示不修改。',
        'settings.clientIdHint': '个人邮箱遇到 AADSTS65002 时，可以在 Entra 注册自己的应用，把「应用程序（客户端）ID」填到这里。注册和启用公共客户端流的步骤见顶部「流程指引」。',
        'settings.tenant': '租户',
        'settings.demo': '未连接时显示示例数据',
        'settings.mailTop': '邮件条数',
        'settings.eventDays': '日历天数',
        'auth.step1': '① 点「打开登录页」，用浏览器登录你的微软账户（工作/学校账户或个人账户都行）。',
        'auth.step2': '② 登录成功后浏览器会停在一个空白页，地址栏里带 code=... 。把那条完整网址粘贴到下面，点「完成登录」。',
        'auth.open': '打开登录页',
        'auth.openSystem': '改用系统浏览器打开',
        'auth.complete': '完成登录',
        'auth.orDevice': '备用方式：一次性代码（可选）',
        'cal.new': '新建日程',
        'cal.create': '创建',
        'cal.created': '日程已创建',
        'cal.createHint': '点击单元格新建日程',
        'cal.subject': '日程标题',
        'cal.location': '地点（可选）',
        'cal.mine': '我的日历',
        'cal.today': '今天',
        'cal.loading': '加载中…',
        'tab.mail': '邮件',
        'tab.calendar': '日历',
        'tab.todo': '计划',
        'tab.reply': '回复',
        'tab.appearance': '界面',
        'guide.open': '流程指引',
        'guide.title': '连接 Outlook 的流程指引',
        'guide.version': '插件版本',
        'guide.intro': '请用自己的 Microsoft 账号登录。若内置登录遇到 AADSTS65002，可按下面的步骤配置自己的 Entra 应用。',
        'guide.register': '注册 Entra 应用',
        'guide.registerHelp': '在 Microsoft Entra 管理中心打开「Entra ID → 应用注册 → 新注册」。个人邮箱请选择支持个人 Microsoft 账户的类型。注册后，在「概述」复制应用程序（客户端）ID。',
        'guide.openEntra': '打开 Entra 管理中心',
        'guide.publicClient': '启用公共客户端流',
        'guide.publicClientHelp': '在「Entra ID → 应用注册」列表里，点击第 1 步创建的应用名称，进入应用详情。列表为空时，先完成新注册，或切换到创建该应用的目录。进入详情后，在左侧「管理」下打开「身份验证（预览版） / Authentication (Preview)」，再点「设置」选项卡。如果显示的是「身份验证」，在页面下方找「高级设置」。开启「允许公共客户端流」并保存，设备代码登录时可将客户端密钥留空。',
        'guide.configure': '填写插件设置',
        'guide.configureHelp': '粘贴自己的客户端 ID。租户按注册时选择的账户类型填写。仅个人账户填 consumers，同时支持组织和个人账户填 common，单租户组织应用填目录（租户）ID。填好后点「保存」。',
        'guide.openSettings': '打开插件设置',
        'guide.authorize': '登录并授权',
        'guide.authorizeHelp': '点击连接或重新授权，在微软登录页按提示输入设备代码，用自己的邮箱账号登录并同意授权，然后返回 Outlook 面板。',
        'ui.panel': '主面板',
        'ui.mail': '邮件列表',
        'ui.calendar': '日历月视图',
        'ui.calendarList': '日历列表',
        'ui.todoCreate': '计划添加区',
        'ui.todoCategories': '分类',
        'ui.todoCompleted': '已完成',
        'ui.todoAll': '所有项目',
        'ui.todoToday': '今天',
        'ui.todoTomorrow': '明天',
        'ui.mailFolders': '邮件文件夹',
        'ui.mailReading': '邮件阅读区',
        'ui.toolbar': '工具栏',
        'ui.aiBar': 'AI 输入栏',
        'ui.header': '顶部栏',
        'ui.opacity': '不透明度',
        'ui.blur': '模糊度',
        'ui.title': '界面设置',
        'ui.surfaces': '主界面',
        'ui.hint': '按区域调整不透明度和模糊度。改动即时预览，满意后保存外观。',
        'ui.save': '保存外观',
        'ui.reset': '恢复默认',
        'ui.saved': '外观已保存',
        'reply.title': '回复整理',
        'reply.open': '写回复',
        'reply.back': '返回日历',
        'reply.to': '收件人',
        'reply.subject': '主题',
        'reply.quote': '原文',
        'reply.placeholder': '在这里写回复，发送时 Outlook 会自动带上原文引用与线程关系…',
        'reply.send': '发送回复',
        'reply.clear': '清空',
        'reply.hint': '通过 Microsoft Graph 以你的账户直接发送',
        'reply.sent': '回复已发送',
        'reply.pickFirst': '先在左侧选择一封邮件',
        'connect.title': '连接 Outlook',
        'connect.step1': '① 点击「登录微软账户」拿到一次性代码。',
        'connect.step2': '② 打开登录页，粘贴代码，用工作或学校账户登录并同意授权。',
        'connect.step3': '③ 授权完成后本窗口自动连接，不需要注册任何应用。',
        'connect.start': '登录微软账户',
        'connect.code': '一次性代码',
        'connect.copy': '复制代码',
        'connect.copied': '代码已复制，粘贴到登录页即可。',
        'connect.opened': '已在系统默认浏览器里打开登录页，代码已自动填入，直接登录即可。',
        'connect.remain': '有效期约 {n} 分钟',
        'connect.openPlain': '打开 microsoft.com/devicelogin',
        'connect.manual': '手动打开',
        'connect.openInApp': '在应用内打开登录页',
        'connect.open': '用系统浏览器打开',
        'connect.waiting': '等待你在浏览器中完成登录…',
        'connect.builtin': '当前用的是内置客户端。若登录页报 AADSTS65002，说明微软策略拦截了内置客户端，请在「设置」里填一个自己的客户端 ID（免费注册，步骤见设置页）。',
        'connect.timeout': '等了太久还没收到授权结果。重新点一次「登录微软账户」即可，宿主会在后台继续等待。',
        'connect.done': '已连接 Outlook',
        'connect.missing': '请先在「设置」里填写 Microsoft 应用（客户端）ID。',
        'notice.demo': '当前展示示例数据。连接 Outlook 后会自动切换为你的真实邮件、日程与待办。',
        'error.title': '数据加载失败',
        'error.retry': '重试',
        'loading': '加载中…',
        'week.0': '周日',
        'week.1': '周一',
        'week.2': '周二',
        'week.3': '周三',
        'week.4': '周四',
        'week.5': '周五',
        'week.6': '周六',
      },
      en: {
        panel: 'Outlook',
        title: 'Outlook',
        subtitle: 'Mail / Calendar / To-do',
        'chip.demo': 'Sample data',
        'chip.offline': 'Not connected',
        'action.refresh': 'Refresh',
        'action.settings': 'Settings',
        'action.connect': 'Connect Outlook',
        'action.disconnect': 'Disconnect',
        'action.reauth': 'Re-authorize',
        'action.save': 'Save',
        'action.close': 'Close',
        'action.open': 'Open in Outlook',
        'action.markRead': 'Mark as read',
        'cat.security': 'Security',
        'cat.billing': 'Billing',
        'cat.work': 'Work',
        'cat.social': 'Platform',
        'cat.promo': 'Promotions',
        'cat.other': 'Other',
        'mail.cat.all': 'All',
        'mail.compose': 'Compose',
        'mail.to': 'Recipients (comma separated)',
        'mail.subject': 'Subject',
        'mail.body': 'Message',
        'mail.send': 'Send',
        'mail.sent': 'Message sent',
        'mail.folders': 'Folders',
        'mail.focused': 'Focused',
        'mail.other': 'Other',
        'mail.today': 'Today',
        'mail.yesterday': 'Yesterday',
        'mail.week': 'This week',
        'mail.month': 'This month',
        'mail.earlier': 'Earlier',
        'mail.delete': 'Delete',
        'mail.archive': 'Archive',
        'mail.toggleRead': 'Mark read / unread',
        'mail.title': 'Mailbox',
        'mail.unread': '{n} unread',
        'mail.allRead': 'All read',
        'mail.all': 'All',
        'mail.unreadOnly': 'Unread',
        'mail.empty': 'Your inbox is empty',
        'mail.emptyUnread': 'No unread mail',
        'mail.attachment': 'Attachment',
        'calendar.title': 'Calendar',
        'calendar.empty': 'Nothing scheduled',
        'calendar.today': 'Today',
        'calendar.tomorrow': 'Tomorrow',
        'calendar.yesterday': 'Yesterday',
        'calendar.allDay': 'All day',
        'todo.title': 'Plan',
        'todo.empty': 'Nothing outstanding',
        'todo.summary': '{count} open',
        'todo.overdue': 'Overdue',
        'todo.noOverdue': 'No overdue plans',
        'todo.all': 'All items',
        'todo.categories': 'Categories',
        'todo.categoryFilter': 'Filter by category or task list',
        'todo.completed': 'Completed',
        'todo.sortCompleted': 'Sort completed plans',
        'todo.sortCompletion': 'Completion time',
        'todo.sortCreation': 'Creation time',
        'todo.sortAscending': 'Sort oldest first',
        'todo.stale': 'Showing plans from the last sync.',
        'todo.noCompleted': 'No completed plans yet',
        'todo.complete': 'Complete plan: {title}',
        'todo.restore': 'Restore plan: {title}',
        'todo.menuComplete': 'Mark as completed',
        'todo.menuRestore': 'Restore as incomplete',
        'todo.menuDue': 'Set due date to {day}',
        'todo.menuClearDue': 'Clear due date',
        'todo.menuCopyTitle': 'Copy plan title',
        'todo.menuDelete': 'Delete plan',
        'todo.dueDate': 'Due date',
        'todo.dragHint': 'Drag a card to change its due date or category. Right-click for more actions.',
        'todo.dropNoDate': 'Drop to clear the due date',
        'todo.dropDate': 'Drop to set the due date to {date}',
        'todo.dropCategory': 'Drop to assign “{name}”',
        'todo.due': 'Due {time}',
        'todo.list': 'Task list',
        'todo.noLists': 'No task lists available',
        'todo.chooseList': 'Choose a list',
        'todo.listMissing': 'Microsoft To Do could not find the selected list. Refresh or choose a list before adding the task again.',
        'todo.taskMissing': 'Microsoft To Do could not find this task. Refresh the list and try again.',
        'todo.taskGone': 'Microsoft To Do could not find this task. The list has been refreshed.',
        'todo.new': 'Enter a task title',
        'todo.newWithDate': 'Enter a task title, optionally ending in /month.day or /year.month.day',
        'todo.dateHint': 'Add /month.day or /year.month.day to the end of the title to set a due date. If you omit the year, the current year is used.',
        'todo.add': 'Add task',
        'todo.demo': 'Showing demo tasks. Connect Outlook to add tasks and sync them to Microsoft To Do.',
        'settings.title': 'Connection settings',
        'settings.clientId': 'Microsoft app (client) ID - optional',
        'settings.clientSecret': 'Client secret (optional, confidential client)',
        'settings.clientSecretHint': 'If the app has no "allow public client flows" switch, use a confidential client instead: Certificates & secrets -> New client secret -> copy the Value -> paste it here. Leave empty to keep the stored one.',
        'settings.clientIdHint': 'If a personal mailbox returns AADSTS65002, register your own Entra app and enter its Application (client) ID here. See the Setup guide for registration and public client flow settings.',
        'settings.tenant': 'Tenant',
        'settings.demo': 'Show sample data while disconnected',
        'settings.mailTop': 'Messages',
        'settings.eventDays': 'Calendar days',
        'auth.step1': '1. Press "Open sign-in page" and sign in with your Microsoft account (work, school or personal).',
        'auth.step2': '2. After signing in the browser stops on a blank page whose address contains code=... . Paste that whole URL below and press "Finish sign-in".',
        'auth.open': 'Open sign-in page',
        'auth.openSystem': 'Open in system browser',
        'auth.complete': 'Finish sign-in',
        'auth.orDevice': 'Fallback: one-time code (optional)',
        'cal.new': 'New event',
        'cal.create': 'Create',
        'cal.created': 'Event created',
        'cal.createHint': 'Click a day to create an event',
        'cal.subject': 'Event title',
        'cal.location': 'Location (optional)',
        'cal.mine': 'My calendars',
        'cal.today': 'Today',
        'cal.loading': 'Loading…',
        'tab.mail': 'Mail',
        'tab.calendar': 'Calendar',
        'tab.todo': 'Plan',
        'tab.reply': 'Reply',
        'tab.appearance': 'Interface',
        'guide.open': 'Setup guide',
        'guide.title': 'Connect Outlook',
        'guide.version': 'Plugin version',
        'guide.intro': 'Sign in with your own Microsoft account. If the built-in sign-in returns AADSTS65002, follow these steps to configure your own Entra app.',
        'guide.register': 'Register an Entra app',
        'guide.registerHelp': 'In the Microsoft Entra admin center, go to Entra ID → App registrations → New registration. For a personal mailbox, choose an account type that supports personal Microsoft accounts. Copy the Application (client) ID from Overview.',
        'guide.openEntra': 'Open the Entra admin center',
        'guide.publicClient': 'Enable public client flows',
        'guide.publicClientHelp': 'In Entra ID → App registrations, select the app you created in step 1 to open its details. If the list is empty, complete the registration or switch to the directory where you created the app. Under Manage, open Authentication (Preview) and select Settings. If you see Authentication, look under Advanced settings. Enable Allow public client flows and save. Leave the client secret blank for device code sign-in.',
        'guide.configure': 'Configure the plugin',
        'guide.configureHelp': 'Paste your client ID. Use consumers for personal accounts only, common for both organizational and personal accounts, or the Directory (tenant) ID for a single-tenant organizational app. Select Save.',
        'guide.openSettings': 'Open plugin settings',
        'guide.authorize': 'Sign in and authorize',
        'guide.authorizeHelp': 'Select Connect or Reauthorize. Follow the Microsoft sign-in page’s device code instructions, sign in with your own mailbox account and grant consent, then return to Outlook.',
        'ui.panel': 'Main panel',
        'ui.mail': 'Mail list',
        'ui.calendar': 'Calendar month view',
        'ui.calendarList': 'Calendar list',
        'ui.todoCreate': 'Add plan area',
        'ui.todoCategories': 'Categories',
        'ui.todoCompleted': 'Completed',
        'ui.todoAll': 'All items',
        'ui.todoToday': 'Today',
        'ui.todoTomorrow': 'Tomorrow',
        'ui.mailFolders': 'Mail folders',
        'ui.mailReading': 'Mail reading pane',
        'ui.toolbar': 'Toolbar',
        'ui.aiBar': 'AI input bar',
        'ui.header': 'Header',
        'ui.opacity': 'Opacity',
        'ui.blur': 'Blur',
        'ui.title': 'Interface settings',
        'ui.surfaces': 'Main interface',
        'ui.hint': 'Adjust opacity and blur by region. Changes preview immediately. Save to keep them.',
        'ui.save': 'Save appearance',
        'ui.reset': 'Reset',
        'ui.saved': 'Appearance saved',
        'reply.title': 'Reply',
        'reply.open': 'Write reply',
        'reply.back': 'Back to calendar',
        'reply.to': 'To',
        'reply.subject': 'Subject',
        'reply.quote': 'Original',
        'reply.placeholder': 'Write your reply here; Outlook keeps the quoted thread and subject.',
        'reply.send': 'Send reply',
        'reply.clear': 'Clear',
        'reply.hint': 'Sent from your account through Microsoft Graph',
        'reply.sent': 'Reply sent',
        'reply.pickFirst': 'Pick a message on the left first',
        'connect.title': 'Connect Outlook',
        'connect.step1': '1. Press "Log in with Microsoft" to get a one-time code.',
        'connect.step2': '2. Open the sign-in page, paste the code, and approve access with your work or school account.',
        'connect.step3': '3. The panel connects by itself. No app registration anywhere.',
        'connect.start': 'Log in with Microsoft',
        'connect.code': 'One-time code',
        'connect.copy': 'Copy code',
        'connect.copied': 'Code copied - paste it on the sign-in page.',
        'connect.opened': 'Opened the sign-in page in your default browser with the code filled in.',
        'connect.remain': 'Valid for about {n} minutes',
        'connect.openPlain': 'Open microsoft.com/devicelogin',
        'connect.manual': 'Open manually',
        'connect.openInApp': 'Open sign-in inside the app',
        'connect.open': 'Open in system browser',
        'connect.waiting': 'Waiting for you to finish signing in...',
        'connect.builtin': 'Using the built-in client. If the sign-in page reports AADSTS65002, Microsoft policy blocked it - put your own client ID in Settings (free, steps are on the settings page).',
        'connect.timeout': 'No authorization result arrived in time. Press "Log in with Microsoft" again; the Host keeps waiting in the background.',
        'connect.done': 'Outlook connected',
        'connect.missing': 'Fill in the Microsoft app (client) ID in Settings first.',
        'notice.demo': 'Showing sample rows. Connect Outlook to switch to your real mail, calendar and to-do.',
        'error.title': 'Could not load data',
        'error.retry': 'Retry',
        'loading': 'Loading...',
        'week.0': 'Sun',
        'week.1': 'Mon',
        'week.2': 'Tue',
        'week.3': 'Wed',
        'week.4': 'Thu',
        'week.5': 'Fri',
        'week.6': 'Sat',
      },
    }

    /** Dictionary fallback used before (or without) the host locale service. */
    function fbT(key, params) {
      let text = DICT.zh[key] ?? key
      if (params && typeof params === 'object') {
        for (const [name, value] of Object.entries(params)) {
          text = text.split(`{${name}}`).join(String(value))
        }
      }
      return text
    }

    /* -------------------------------------------------------------- store */

    const store = {
      snapshot: null,
      loading: false,
      error: null,
      subs: new Set(),
      timer: null,
      inflight: null,

      subscribe(fn) {
        store.subs.add(fn)
        if (!store.timer) {
          store.timer = setInterval(() => {
            store.load(false)
          }, 60_000)
          store.load(false)
        }
        return () => {
          store.subs.delete(fn)
          if (store.subs.size === 0 && store.timer) {
            clearInterval(store.timer)
            store.timer = null
          }
        }
      },

      emit() {
        for (const fn of [...store.subs]) {
          try {
            fn()
          } catch {
            /* a broken subscriber must not stop the rest */
          }
        }
      },

      async load(force) {
        if (store.inflight) return store.inflight
        store.loading = true
        store.emit()
        store.inflight = (async () => {
          try {
            const data = await api(`/snapshot${force ? '?refresh=1' : ''}`)
            const previous = store.snapshot
            if (data.taskError && data.connected && previous?.connected &&
                data.account?.mail && data.account.mail === previous.account?.mail &&
                previous.taskLists?.length && (!previous.taskError || previous.taskStale)) {
              data.taskLists = previous.taskLists
              data.tasks = previous.tasks
              data.taskStale = true
            }
            store.snapshot = data
            store.error = null
          } catch (error) {
            store.error = error?.message || String(error)
          } finally {
            store.loading = false
            store.inflight = null
            store.emit()
          }
        })()
        return store.inflight
      },

      patch(patch) {
        store.snapshot = { ...(store.snapshot ?? {}), ...patch }
        store.emit()
      },

      dispose() {
        if (store.timer) {
          clearInterval(store.timer)
          store.timer = null
        }
        store.subs.clear()
      },
    }

    let recentCalendarRange = null
    let calendarRangeRevision = 0

    async function api(path, options) {
      const init = {
        method: options?.method ?? 'GET',
        credentials: 'include',
        cache: 'no-store',
        ...(options?.signal ? { signal: options.signal } : {}),
        headers: { 'x-outlook-client': '1' },
      }
      if (options?.body !== undefined) {
        init.headers['content-type'] = 'application/json'
        init.body = JSON.stringify(options.body)
      }
      const changesCalendar = init.method === 'POST' && /^\/calendar\/(create|update|delete)$/.test(path)
      if (changesCalendar) {
        recentCalendarRange = null
        calendarRangeRevision += 1
      }
      const res = await fetch(`/outlook/api${path}`, init)
      if (changesCalendar) {
        recentCalendarRange = null
        calendarRangeRevision += 1
      }
      let data = null
      try {
        data = await res.json()
      } catch {
        data = null
      }
      if (!res.ok) {
        const error = new Error(data?.error || `HTTP ${res.status}`)
        error.status = res.status
        throw error
      }
      if (data && data.ok === false) throw new Error(data.error || 'request failed')
      return data
    }

    function useSnapshot() {
      const [, bump] = React.useState(0)
      React.useEffect(() => store.subscribe(() => bump((n) => n + 1)), [])
      return store.snapshot
    }

    /* -------------------------------------------------------------- utils */

    function pad(value) {
      return String(value).padStart(2, '0')
    }

    function htmlToPlainText(value) {
      let text = String(value ?? '')
      if (!text) return ''
      if (/<\s*[a-z!]/i.test(text)) {
        text = text
          .replace(/<style[\s\S]*?<\/style>/gi, '')
          .replace(/<script[\s\S]*?<\/script>/gi, '')
          .replace(/<head[\s\S]*?<\/head>/gi, '')
          .replace(/<br\s*\/?>/gi, '\n')
          .replace(/<\/?(p|div|tr|li|h[1-6])\b[^>]*>/gi, (m) => (m.startsWith('</') ? '\n' : ''))
          .replace(/<[^>]+>/g, '')
      }
      text = text
        .replace(/&nbsp;/gi, ' ')
        .replace(/&#160;/g, ' ')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;|&apos;/g, "'")
        .replace(/&#(\d+);/g, (_, n) => {
          const code = Number(n)
          return Number.isFinite(code) ? String.fromCharCode(code) : ''
        })
        .replace(/&#x([0-9a-f]+);/gi, (_, n) => {
          const code = parseInt(n, 16)
          return Number.isFinite(code) ? String.fromCharCode(code) : ''
        })
      return text.replace(/\u00a0/g, ' ').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim()
    }


    function parseDate(value) {
      if (!value) return null
      const text = String(value)
      if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(text) && !/[zZ]|[+-]\d{2}:\d{2}$/.test(text)) {
        // Graph returns wall-clock time without an offset.
        return new Date(text.slice(0, 19))
      }
      const parsed = new Date(text)
      return Number.isNaN(parsed.getTime()) ? null : parsed
    }

    function dayKey(date) {
      return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
    }

    function startOfDay(date) {
      const copy = new Date(date)
      copy.setHours(0, 0, 0, 0)
      return copy
    }

    function fmtTime(date) {
      return date ? `${pad(date.getHours())}:${pad(date.getMinutes())}` : ''
    }

    function fmtDay(date, t) {
      const today = startOfDay(new Date())
      const target = startOfDay(date)
      const diff = Math.round((target.getTime() - today.getTime()) / 86_400_000)
      if (diff === 0) return t('calendar.today')
      if (diff === 1) return t('calendar.tomorrow')
      const week = t(`week.${date.getDay()}`)
      return `${date.getMonth() + 1}/${date.getDate()} ${week}`
    }

    function fmtMailTime(date, t) {
      if (!date) return ''
      const today = startOfDay(new Date())
      const target = startOfDay(date)
      const diff = Math.round((target.getTime() - today.getTime()) / 86_400_000)
      if (diff === 0) return fmtTime(date)
      if (diff === -1) return t('calendar.yesterday')
      return `${date.getMonth() + 1}/${date.getDate()}`
    }

    function initial(name) {
      const text = String(name || '?').trim()
      return text ? text.slice(0, 1).toUpperCase() : '?'
    }

    /* ------------------------------------------------------------- artwork */

    function mark(size) {
      return h(
        'svg',
        {
          viewBox: '0 0 24 24',
          width: size,
          height: size,
          'aria-hidden': 'true',
          style: { display: 'block', flex: 'none' },
        },
        h(
          'defs',
          null,
          h(
            'linearGradient',
            { id: 'ob-mark-grad', x1: '0', y1: '0', x2: '1', y2: '1' },
            h('stop', { offset: '0', stopColor: '#2B8FE6' }),
            h('stop', { offset: '1', stopColor: '#0A56A6' }),
          ),
        ),
        h('rect', { x: 1, y: 3, width: 22, height: 18, rx: 4.5, fill: 'url(#ob-mark-grad)' }),
        h('path', {
          d: 'M3 7.2 12 13.6l9-6.4',
          fill: 'none',
          stroke: '#ffffff',
          strokeWidth: 1.7,
          strokeLinecap: 'round',
          strokeLinejoin: 'round',
        }),
        h('path', {
          d: 'M3 7.2v10.3c0 .6.4 1 1 1h16c.6 0 1-.4 1-1V7.2',
          fill: 'none',
          stroke: 'rgba(255,255,255,0.5)',
          strokeWidth: 1.1,
          strokeLinecap: 'round',
        }),
      )
    }

    const CSS = `
.ob-root{position:relative;display:flex;flex-direction:column;height:100%;min-height:0;color:var(--dsw-alias-label-primary);font-size:13px;line-height:1.45}
.ob-head{display:flex;align-items:center;gap:10px;padding:12px 18px;border-bottom:1px solid var(--dsw-alias-border-l1);flex:none;flex-wrap:wrap}
.ob-head-mark{width:30px;height:30px;border-radius:9px;display:flex;align-items:center;justify-content:center;background:var(--dsw-alias-bg-layer-2);flex:none}
.ob-head-text{display:flex;flex-direction:column;min-width:0}
.ob-title{font-size:14.5px;font-weight:600;letter-spacing:.2px}
.ob-sub{font-size:11.5px;color:var(--dsw-alias-label-tertiary);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:42vw}
.ob-spacer{flex:1}
.ob-chip{font-size:11px;padding:2px 8px;border-radius:999px;border:1px solid var(--dsw-alias-border-l1);color:var(--dsw-alias-label-secondary);white-space:nowrap}
.ob-chip.demo{border-color:transparent;background:var(--dsw-alias-bg-layer-2);color:var(--dsw-alias-label-secondary)}
.ob-chip.ok{border-color:transparent;background:color-mix(in srgb,var(--dsw-alias-state-success-primary) 16%,transparent);color:var(--dsw-alias-state-success-primary)}
.ob-btn{height:27px;padding:0 10px;border-radius:8px;border:1px solid var(--dsw-alias-border-l1);background:transparent;color:var(--dsw-alias-label-primary);font-size:12px;cursor:pointer;display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
.ob-btn:hover{background:var(--dsw-alias-interactive-bg-hover)}
.ob-btn:disabled{opacity:.55;cursor:default}
.ob-btn.primary{background:var(--dsw-alias-brand-primary);border-color:transparent;color:#fff}
.ob-btn.primary:hover{background:var(--dsw-alias-button-primary-hover,var(--dsw-alias-brand-primary))}
.ob-body{flex:1;min-height:0;display:flex;flex-direction:column;gap:12px;padding:12px 18px 16px;overflow:auto}
.ob-grid{flex:1;min-height:0;display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:12px}
.ob-col{display:flex;flex-direction:column;gap:12px;min-height:0}
.ob-card{display:flex;flex-direction:column;min-height:0;background:var(--dsw-alias-bg-layer-1);border:1px solid var(--dsw-alias-border-l1);border-radius:12px;overflow:hidden}
.ob-card.grow{flex:1}
.ob-card-head{display:flex;align-items:center;gap:8px;padding:9px 12px;border-bottom:1px solid var(--dsw-alias-border-l1);flex:none}
.ob-card-title{font-size:12.5px;font-weight:600}
.ob-count{font-size:11px;color:var(--dsw-alias-label-tertiary)}
.ob-seg{display:inline-flex;gap:2px;padding:2px;border-radius:8px;background:var(--dsw-alias-bg-layer-2)}
.ob-seg button{height:20px;padding:0 8px;border:0;border-radius:6px;background:transparent;color:var(--dsw-alias-label-tertiary);font-size:11px;cursor:pointer}
.ob-seg button.on{background:var(--dsw-alias-bg-layer-1);color:var(--dsw-alias-label-primary);box-shadow:0 1px 2px rgba(0,0,0,.08)}
.ob-split{display:flex;flex:1;min-height:0}
.ob-list{flex:1;min-height:0;overflow:auto}
.ob-mail{display:flex;gap:8px;align-items:flex-start;width:100%;text-align:left;border:0;border-bottom:1px solid var(--dsw-alias-border-l1);background:transparent;padding:9px 12px;cursor:pointer;color:inherit;font:inherit}
.ob-mail:last-child{border-bottom:0}
.ob-mail:hover{background:var(--dsw-alias-interactive-bg-hover)}
.ob-mail.sel{background:var(--dsw-alias-interactive-bg-active)}
.ob-dot{width:6px;height:6px;border-radius:50%;background:var(--dsw-alias-brand-primary);margin-top:7px;flex:none}
.ob-dot.read{background:transparent}
.ob-avatar{width:24px;height:24px;border-radius:50%;flex:none;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;background:var(--dsw-alias-bg-layer-2);color:var(--dsw-alias-label-secondary)}
.ob-mail-main{flex:1;min-width:0}
.ob-row1{display:flex;align-items:baseline;gap:8px}
.ob-from{font-size:12.5px;font-weight:600;max-width:55%;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ob-mail.unread .ob-from{color:var(--dsw-alias-label-primary)}
.ob-time{margin-left:auto;font-size:10.5px;color:var(--dsw-alias-label-quaternary);flex:none}
.ob-subj{font-size:12.5px;margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--dsw-alias-label-secondary)}
.ob-mail.unread .ob-subj{color:var(--dsw-alias-label-primary);font-weight:600}
.ob-prev{font-size:11.5px;color:var(--dsw-alias-label-tertiary);margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ob-detail{width:46%;min-width:0;border-left:1px solid var(--dsw-alias-border-l1);display:flex;flex-direction:column;min-height:0}
.ob-detail-head{padding:12px 14px 8px;border-bottom:1px solid var(--dsw-alias-border-l1)}
.ob-detail-subject{font-size:13.5px;font-weight:600;line-height:1.35;word-break:break-word}
.ob-detail-meta{font-size:11.5px;color:var(--dsw-alias-label-tertiary);margin-top:6px;word-break:break-all}
.ob-detail-body{flex:1;min-height:0;overflow:auto;padding:12px 14px;font-size:12.5px;color:var(--dsw-alias-label-secondary);white-space:pre-wrap;word-break:break-word}
.ob-detail-foot{display:flex;gap:8px;padding:10px 14px;border-top:1px solid var(--dsw-alias-border-l1);flex-wrap:wrap}
.ob-daygroup{padding:4px 0}
.ob-dayhead{position:sticky;top:0;z-index:1;padding:6px 12px;font-size:11px;font-weight:600;color:var(--dsw-alias-label-tertiary);background:var(--dsw-alias-bg-layer-1)}
.ob-event{display:flex;gap:10px;padding:7px 12px;align-items:flex-start}
.ob-event-time{width:42px;flex:none;font-size:11.5px;color:var(--dsw-alias-label-tertiary);font-variant-numeric:tabular-nums;padding-top:1px}
.ob-event-main{min-width:0;flex:1}
.ob-event-title{font-size:12.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ob-event-meta{font-size:11px;color:var(--dsw-alias-label-quaternary);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ob-bar{width:2px;border-radius:2px;background:var(--dsw-alias-brand-primary);align-self:stretch;flex:none;opacity:.7}
.ob-todo{display:flex;gap:9px;align-items:flex-start;padding:8px 12px;border-block-end:1px solid var(--ob-line,#ddd)}
.ob-todo:hover{background:color-mix(in srgb,var(--ob-blue) 4%,transparent)}
.ob-todo[draggable=true]{cursor:grab}
.ob-todo[draggable=true]:active{cursor:grabbing}
.ob-todo.ob-plan-dragging{opacity:.38}
.ob-plan-grip{flex:none;color:var(--dsw-alias-label-tertiary,#777)}
.ob-todo-main{min-width:0;flex:1}
.ob-todo-title{font-size:12.5px;word-break:break-word}
.ob-todo-completed .ob-todo-title{text-decoration:line-through;color:var(--dsw-alias-label-secondary,#666)}
.ob-todo-completion-time{display:flex;flex:none;flex-direction:column;align-self:center;text-align:center;white-space:nowrap;font-size:10.5px;font-variant-numeric:tabular-nums;line-height:1.3;opacity:.55}
.ob-todo-meta{font-size:10.5px;color:var(--dsw-alias-label-quaternary);margin-top:1px}
.ob-todo-page{display:flex;flex-direction:column;flex:1;min-height:0;min-width:0;overflow:auto}
.ob-todo-page .ob-card{background:var(--ob-paper);border:0;border-radius:0}
.ob-todo-page > :is(.ob-error,.ob-banner){flex:none;border-radius:0;border-block-end:1px solid var(--ob-line)}
.ob-todo-form{display:flex;flex-wrap:wrap;align-items:center;gap:8px;flex:none}
.ob-todo-form label{display:flex;align-items:center;gap:8px;max-width:100%}
.ob-todo-form select{min-width:0;max-width:210px}
.ob-todo-form .ob-todo-input{flex:1 1 180px;min-width:0}
.ob-todo-form input[type=date]{min-width:0;max-width:100%;box-sizing:border-box}
.ob-todo-form input[type=date]:enabled{cursor:pointer;user-select:none}
.ob-todo-due-row{display:flex;flex:0 0 100%;min-width:0;flex-wrap:wrap;align-items:center;gap:8px}
.ob-todo-due-row .ob-btn[aria-pressed=true]{color:var(--ob-blue);border-color:var(--ob-blue);background:color-mix(in srgb,var(--ob-blue) 8%,transparent)}
.ob-todo-page input[type=checkbox]{accent-color:var(--ob-blue,#0f6cbd)}
.ob-plan-header{flex:none;display:grid;gap:10px;padding:12px;background:var(--ob-paper);border-block-end:2px solid var(--ob-line)}
.ob-plan-heading{display:grid;grid-template-columns:max-content minmax(0,1fr);align-items:center;column-gap:12px}
.ob-plan-heading h2{margin:0;font-size:38px;line-height:1.08}
.ob-plan-heading-details{display:grid;gap:4px;min-width:0}
.ob-plan-hints{display:flex;flex-wrap:wrap;gap:4px 22px}
.ob-plan-summary{justify-self:start;padding:2px 7px;border:1px solid var(--ob-line,#dedede);border-radius:6px;font-size:12px;line-height:1.25;color:var(--dsw-alias-label-secondary,#555)}
.ob-plan-grid{flex:1 0 290px;min-height:290px;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));grid-template-rows:minmax(0,1fr)}
.ob-plan-grid > :not(:last-child){border-inline-end:1px solid var(--ob-line)}
.ob-plan-panel{min-width:0}
.ob-plan-side{display:grid;grid-template-rows:minmax(0,1fr) minmax(0,1fr);min-width:0;min-height:0}
.ob-plan-side > .ob-plan-panel:first-child{border-block-end:1px solid var(--ob-line)}
.ob-plan-panel .ob-card-head{justify-content:space-between}
.ob-plan-sort{display:flex;align-items:center;gap:4px;min-width:0;max-width:65%;margin-inline-start:auto}
.ob-plan-sort select,.ob-plan-sort-direction,.ob-plan-category-filter{box-sizing:border-box;height:25px;border:1px solid var(--ob-line);border-radius:5px;background:var(--ob-paper);color:inherit}
.ob-plan-sort select,.ob-plan-category-filter{min-width:0;padding:0 5px;font:inherit;font-size:11px}
.ob-plan-category-filter{max-width:65%;margin-inline-start:auto}
.ob-plan-sort .ob-plan-sort-direction{flex:none;width:25px;padding:0;justify-content:center;font-size:14px}
.ob-plan-sync-error{display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:8px}
.ob-plan-panel h3{margin:0;font-size:13px;font-weight:650}
.ob-plan-panel .ob-count{padding:2px 7px;border-radius:8px;background:var(--ob-ground);font-variant-numeric:tabular-nums}
.ob-todo-page .ob-plan-overdue{border-inline-start:3px solid var(--dsw-alias-state-error-primary,#b42318)}
.ob-plan-category-filter :is(option,optgroup){color:var(--dsw-alias-label-primary,#242424)}
.ob-plan-overdue :is(h3,.ob-plan-category-filter,.ob-count),.ob-plan-category-filter option[value=overdue]{color:var(--dsw-alias-state-error-primary,#b42318)}
.ob-plan-panel.ob-plan-drop-hot{outline:2px solid var(--ob-blue,#0f6cbd);outline-offset:-2px}
.ob-plan-drop-preview{pointer-events:none;opacity:.62;flex:none;padding:7px 12px;border:1px dashed var(--ob-blue,#0f6cbd);background:color-mix(in srgb,var(--ob-blue,#0f6cbd) 12%,var(--ob-paper))}
.ob-plan-drop-preview strong{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ob-plan-drop-preview span,.ob-plan-drag-hint{font-size:11px;color:var(--dsw-alias-label-secondary,#555)}
@container(max-width:1000px){
  .ob-plan-grid{flex:none;grid-template-columns:repeat(2,minmax(0,1fr));grid-template-rows:none;grid-auto-rows:280px}
  .ob-plan-side{grid-row:span 2}
  .ob-plan-grid > :not(:last-child){border-block-end:1px solid var(--ob-line)}
  .ob-plan-grid > .ob-plan-panel{border-inline-end:0}
  .ob-plan-grid > :last-child{grid-column:1/-1}
}
@container(max-width:570px){
  .ob-plan-grid{grid-template-columns:minmax(0,1fr);grid-auto-rows:220px}
  .ob-plan-grid > .ob-plan-side{border-inline-end:0}
}
.ob-tabs{display:flex;gap:4px;padding:6px 0 0;flex:none}
.ob-tab{height:26px;padding:0 12px;border:0;border-radius:8px 8px 0 0;background:transparent;color:var(--dsw-alias-label-tertiary);font-size:12px;cursor:pointer}
.ob-tab:hover{background:var(--dsw-alias-interactive-bg-hover);color:var(--dsw-alias-label-secondary)}
.ob-tab.on{background:var(--dsw-alias-bg-layer-1);color:var(--dsw-alias-label-primary);font-weight:600}
.ob-grid-calendar{grid-template-columns:minmax(0,1.5fr) minmax(0,1fr)}
.ob-modal-mask{position:fixed;inset:0;z-index:60;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,0.28);backdrop-filter:blur(2px)}
.ob-modal{width:min(760px,92vw);max-height:82vh;display:flex;flex-direction:column;gap:10px;padding:14px;border-radius:14px;border:1px solid var(--dsw-alias-border-l1);background:var(--dsw-alias-bg-overlay);box-shadow:0 18px 48px rgba(0,0,0,0.28)}
.ob-modal-head{display:flex;align-items:center;justify-content:space-between;font-size:13px;font-weight:600;flex:none}
.ob-guide-entry{display:flex;align-items:center;gap:8px}
.ob-guide-dialog{box-sizing:border-box;width:min(600px,calc(100vw - 24px))}
.ob-guide-body{min-height:0;overflow:auto;padding:0 4px 4px}
.ob-guide-body p{margin:0;line-height:1.55;color:var(--dsw-alias-label-secondary,#555)}
.ob-guide-steps{display:grid;gap:8px;margin:12px 0 0;padding:0;list-style:none}
.ob-guide-steps li{display:grid;grid-template-columns:24px minmax(0,1fr);gap:10px;padding:10px;border:1px solid var(--dsw-alias-border-l1,#ddd);border-radius:9px}
.ob-guide-number{display:grid;place-items:center;width:24px;height:24px;border-radius:50%;background:color-mix(in srgb,#0f6cbd 12%,transparent);color:#0f6cbd;font-weight:700}
.ob-guide-steps h3{margin:0 0 4px;font-size:13px;line-height:24px}
.ob-guide-steps .ob-btn{margin-top:8px;text-decoration:none}
.ob-cal-cell{cursor:pointer}
.ob-modal-sm{width:min(380px,90vw)}
.ob-form{display:flex;flex-direction:column;gap:8px}
.ob-form-row{display:flex;align-items:center;gap:8px}
.ob-input{height:30px;padding:0 10px;border-radius:8px;border:1px solid var(--dsw-alias-border-l1);background:var(--dsw-alias-bg-base);color:var(--dsw-alias-label-primary);font-size:12.5px;font-family:inherit}
.ob-input:focus{outline:none;border-color:var(--dsw-alias-brand-primary)}
.ob-cal-notice{align-self:flex-end;margin:6px 10px 8px;padding:5px 10px;border-radius:8px;background:var(--dsw-alias-bg-layer-2);font-size:11.5px;color:var(--dsw-alias-label-secondary);flex:none}
.ob-grid-mail{grid-template-columns:150px minmax(0,1.05fr) minmax(0,1fr)}
.ob-folder{display:flex;align-items:center;gap:6px;width:100%;height:28px;padding:0 8px;border:0;border-radius:7px;background:transparent;color:var(--dsw-alias-label-secondary);font-size:12px;cursor:pointer;text-align:left}
.ob-folder:hover{background:var(--dsw-alias-interactive-bg-hover)}
.ob-folder.on{background:var(--dsw-alias-bg-layer-2);color:var(--dsw-alias-label-primary);font-weight:600}
.ob-folder-name{flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ob-folder-count{font-size:10.5px;color:var(--dsw-alias-brand-primary);font-weight:600}
.ob-tabs2{display:flex;gap:14px;align-items:center;padding:0 12px;border-bottom:1px solid var(--dsw-alias-border-l1);flex:none}
.ob-tab2{height:30px;border:0;background:transparent;color:var(--dsw-alias-label-tertiary);font-size:12px;cursor:pointer;border-bottom:2px solid transparent;padding:0 2px}
.ob-tab2.on{color:var(--dsw-alias-label-primary);font-weight:600;border-bottom-color:var(--dsw-alias-brand-primary)}
.ob-mail-avatar{width:26px;height:26px;border-radius:50%;flex:none;display:flex;align-items:center;justify-content:center;color:#fff;font-size:12px;font-weight:600}
.ob-clip{font-size:11px;opacity:.6;flex:none;align-self:center}
.ob-mail.unread{box-shadow:inset 3px 0 0 var(--dsw-alias-brand-primary)}
.ob-cats{display:flex;gap:6px;padding:7px 12px;border-bottom:1px solid var(--dsw-alias-border-l1);flex-wrap:wrap;flex:none}
.ob-cat{display:inline-flex;align-items:center;gap:5px;height:22px;padding:0 9px;border-radius:999px;border:1px solid var(--dsw-alias-border-l1);background:transparent;color:var(--dsw-alias-label-tertiary);font-size:11px;cursor:pointer}
.ob-cat:hover{background:var(--dsw-alias-interactive-bg-hover);color:var(--dsw-alias-label-secondary)}
.ob-cat.on{background:var(--dsw-alias-bg-layer-2);color:var(--dsw-alias-label-primary);border-color:transparent;font-weight:600}
.ob-cat-dot{width:7px;height:7px;border-radius:2px}
.ob-tag{font-size:10px;border:1px solid currentColor;border-radius:4px;padding:0 4px;opacity:.85;flex:none}
.ob-modal-md{width:min(520px,92vw)}
.ob-cal-layout{display:grid;grid-template-columns:190px minmax(0,1fr);gap:12px;min-height:0;flex:1}
.ob-cal-rail{display:flex;flex-direction:column;gap:6px;min-height:0;overflow:auto;padding-right:2px}
.ob-rail-title{display:flex;align-items:center;justify-content:space-between;font-size:11.5px;font-weight:600;color:var(--dsw-alias-label-tertiary);margin:2px 0 4px}
.ob-rail-item{display:flex;align-items:center;gap:7px;font-size:12px;padding:3px 4px;border-radius:7px;cursor:pointer;color:var(--dsw-alias-label-secondary)}
.ob-rail-item:hover{background:var(--dsw-alias-interactive-bg-hover)}
.ob-rail-item input{accent-color:var(--dsw-alias-brand-primary);margin:0}
.ob-rail-dot{width:9px;height:9px;border-radius:3px;flex:none}
.ob-rail-name{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ob-rail-todo{margin-top:10px;border-top:1px solid var(--dsw-alias-border-l1);padding-top:8px;display:flex;flex-direction:column;gap:2px;min-height:0;overflow:auto}
.ob-cal-head{display:flex;align-items:center;gap:8px;padding:9px 12px;border-bottom:1px solid var(--dsw-alias-border-l1);flex:none}
.ob-cal-month{font-size:13.5px;font-weight:600;margin-left:4px}
.ob-icon-btn{width:24px;height:24px;border:1px solid var(--dsw-alias-border-l1);border-radius:7px;background:transparent;color:var(--dsw-alias-label-secondary);cursor:pointer;font-size:13px;line-height:1;padding:0}
.ob-icon-btn:hover{background:var(--dsw-alias-interactive-bg-hover)}
.ob-error-inline{font-size:11px;color:var(--dsw-alias-state-error-primary)}
.ob-cal-weekdays{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));border-bottom:1px solid var(--dsw-alias-border-l1);flex:none}
.ob-cal-weekdays span{padding:6px 8px;font-size:11px;color:var(--dsw-alias-label-tertiary);text-align:left}
.ob-cal-grid{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));grid-auto-rows:minmax(74px,1fr);flex:1;min-height:0;overflow:auto}
.ob-cal-cell{border-right:1px solid var(--dsw-alias-border-l1);border-bottom:1px solid var(--dsw-alias-border-l1);padding:4px 5px;display:flex;flex-direction:column;gap:3px;min-width:0;overflow:hidden}
.ob-cal-cell:nth-child(7n){border-right:0}
.ob-cal-cell.outside .ob-cal-daynum{color:var(--dsw-alias-label-quaternary)}
.ob-cal-cell.today{background:color-mix(in srgb,var(--dsw-alias-brand-primary) 6%,transparent)}
.ob-cal-daynum{font-size:11.5px;color:var(--dsw-alias-label-secondary);display:flex}
.ob-cal-daynum .pill{background:var(--dsw-alias-brand-primary);color:#fff;border-radius:999px;min-width:18px;height:18px;display:inline-flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;padding:0 5px}
.ob-cal-events{display:flex;flex-direction:column;gap:2px;min-height:0;overflow:hidden}
.ob-cal-chip{font-size:10.5px;line-height:1.35;padding:2px 5px;border-left:3px solid var(--dsw-alias-brand-primary);border-radius:4px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--dsw-alias-label-primary)}
.ob-cal-more{font-size:10px;color:var(--dsw-alias-label-quaternary);padding-left:4px}
.ob-appearance-dialog{--ob-settings-paper:var(--dsw-alias-bg-overlay,#fff);--ob-settings-accent:#7b7073;--ob-settings-line:color-mix(in srgb,currentColor 13%,transparent);--ob-settings-muted:color-mix(in srgb,currentColor 7%,var(--ob-settings-paper));box-sizing:border-box;width:min(660px,100%);max-height:min(100%,720px);gap:0;padding:0;overflow:hidden;color:var(--dsw-alias-label-primary,#202020);background:color-mix(in srgb,var(--ob-settings-paper) 94%,currentColor)}
.ob-appearance-dialog .ob-modal-head,.ob-ui-foot{padding:12px clamp(12px,4vw,22px);background:var(--ob-settings-paper)}
.ob-appearance-dialog .ob-modal-head{border-bottom:1px solid var(--ob-settings-line)}
.ob-appearance-dialog .ob-btn{min-height:32px;border-color:var(--ob-settings-line);font-weight:500}
.ob-ui-page{display:flex;flex-direction:column;gap:16px;min-height:0;overflow:auto;padding:18px clamp(12px,4vw,22px) 20px}
.ob-ui-kicker{color:var(--dsw-alias-label-secondary,#625b5d);font-size:10px;font-weight:700;letter-spacing:.16em}
.ob-ui-head h2{margin:3px 0 4px;font-size:21px;line-height:1.35;font-weight:700}
.ob-ui-head p{margin:0;font-size:12px;color:var(--dsw-alias-label-secondary,#555)}
.ob-appearance-tabs{display:flex;flex-wrap:wrap;gap:6px}
.ob-appearance-tabs [aria-pressed=true]{background:var(--ob-settings-paper);border-color:var(--ob-settings-accent);box-shadow:0 1px 2px #0001;font-weight:650}
.ob-surface-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,250px),1fr));gap:12px}
.ob-surface-card{display:flex;flex-direction:column;gap:8px;min-width:0;padding:16px;border:1px solid var(--ob-settings-line);border-radius:13px;background:var(--ob-settings-paper)}
.ob-surface-title{display:flex;align-items:center;gap:10px;margin-bottom:4px}
.ob-surface-title h3{flex:1;margin:0;font-size:13px;font-weight:700}
.ob-surface-icon{display:grid;flex:none;place-items:center;width:32px;height:32px;border-radius:9px;background:var(--ob-settings-muted);color:var(--ob-settings-accent)}
.ob-surface-badge{flex:none;border-radius:7px;padding:2px 8px;background:var(--ob-settings-muted);font-size:11px;font-variant-numeric:tabular-nums}
.ob-slider{display:grid;grid-template-columns:minmax(0,1fr) 42px;align-items:center;column-gap:10px;font-size:12px}
.ob-slider>span:first-child{grid-column:1 / -1;font-weight:600}
.ob-slider input[type=range]{width:100%;min-width:0;height:22px;margin:0;cursor:pointer;accent-color:var(--ob-settings-accent)}
.ob-slider b{font-size:11px;font-weight:500;text-align:right;font-variant-numeric:tabular-nums}
.ob-appearance-dialog :is(button,input):focus-visible{outline:2px solid var(--ob-settings-accent);outline-offset:2px}
.ob-ui-foot{display:flex;flex:none;justify-content:flex-end;gap:8px;border-top:1px solid var(--ob-settings-line)}
.ob-appearance-dialog .ob-ui-foot{flex-wrap:wrap;align-items:center}
.ob-ui-feedback{flex:1;min-width:100px;overflow-wrap:anywhere;font-size:12px;color:var(--dsw-alias-label-secondary,#555)}
.ob-ui-feedback[role=alert]{color:var(--dsw-alias-state-error-primary,#b42318)}
.ob-appearance-dialog .ob-btn.primary{background:var(--ob-settings-accent);border-color:transparent}
.ob-appearance-dialog .ob-btn.primary:hover:not(:disabled){background:#63595c}
.ob-reply{display:flex;flex-direction:column;gap:10px;padding:12px;flex:1;min-height:0}
.ob-reply-meta{display:flex;flex-direction:column;gap:4px;font-size:11.5px;color:var(--dsw-alias-label-secondary)}
.ob-reply-row{display:flex;gap:8px;align-items:baseline;min-width:0}
.ob-reply-label{color:var(--dsw-alias-label-quaternary);flex:none;width:44px}
.ob-reply-quote{color:var(--dsw-alias-label-tertiary);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ob-reply-body{flex:1;min-height:160px;font-family:inherit;font-size:12.5px}
.ob-reply-foot{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.ob-empty{padding:26px 16px;text-align:center;color:var(--dsw-alias-label-tertiary);font-size:12px}
.ob-banner{display:flex;align-items:center;gap:10px;padding:9px 12px;border-radius:10px;font-size:12px;background:var(--dsw-alias-bg-layer-2);color:var(--dsw-alias-label-secondary);flex:none}
.ob-connect{display:flex;flex-direction:column;gap:12px;padding:18px;border:1px solid var(--dsw-alias-border-l1);border-radius:12px;background:var(--dsw-alias-bg-layer-1);max-width:620px}
.ob-connect h3{margin:0;font-size:13.5px}
.ob-connect p{margin:0;font-size:12px;color:var(--dsw-alias-label-secondary)}
.ob-textarea{width:100%;box-sizing:border-box;padding:9px 10px;border-radius:9px;border:1px solid var(--dsw-alias-border-l1);background:var(--dsw-alias-bg-base);color:var(--dsw-alias-label-primary);font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:11.5px;resize:vertical;line-height:1.5}
.ob-textarea:focus{outline:none;border-color:var(--dsw-alias-brand-primary)}
.ob-divider{display:flex;align-items:center;gap:8px;margin:4px 0;color:var(--dsw-alias-label-quaternary);font-size:11px}
.ob-divider::after{content:'';flex:1;height:1px;background:var(--dsw-alias-border-l1)}
.ob-code{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:22px;letter-spacing:3px;font-weight:600;padding:10px 14px;border-radius:10px;background:var(--dsw-alias-bg-layer-2);align-self:flex-start}
.ob-field{display:flex;flex-direction:column;gap:5px;font-size:12px}
.ob-field label{color:var(--dsw-alias-label-secondary)}
.ob-field input{height:30px;padding:0 10px;border-radius:8px;border:1px solid var(--dsw-alias-border-l1);background:var(--dsw-alias-bg-base);color:var(--dsw-alias-label-primary);font-size:12.5px;font-family:inherit}
.ob-field small{color:var(--dsw-alias-label-quaternary);font-size:11px;line-height:1.5}
.ob-settings{display:flex;flex-direction:column;gap:12px;padding:16px;border:1px solid var(--dsw-alias-border-l1);border-radius:12px;background:var(--dsw-alias-bg-layer-1);max-width:620px}
.ob-settings-row{display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap}
.ob-switch{display:inline-flex;align-items:center;gap:8px;font-size:12px;color:var(--dsw-alias-label-secondary);cursor:pointer}
.ob-switch input{accent-color:var(--dsw-alias-brand-primary)}
.ob-error{padding:10px 12px;border-radius:10px;font-size:12px;background:color-mix(in srgb,var(--dsw-alias-state-error-primary) 12%,transparent);color:var(--dsw-alias-state-error-primary)}
.ob-glyph{position:relative;display:inline-flex;align-items:center;justify-content:center}
.ob-badge{position:absolute;top:-4px;right:-6px;min-width:14px;height:14px;padding:0 3px;border-radius:8px;background:var(--dsw-alias-state-error-primary);color:#fff;font-size:9px;line-height:14px;font-weight:700;text-align:center;font-variant-numeric:tabular-nums;box-shadow:0 0 0 1.5px var(--dsw-specific-sidebar-fill);pointer-events:none}

/* Outlook mail workspace: contiguous navigation, list and reading surfaces. */
:is(.ob-outlook-mail,.ob-outlook-calendar){--ob-blue:#0f6cbd;--ob-line:var(--dsw-alias-border-l1,#dedede);font-family:"Segoe UI","Microsoft YaHei",sans-serif;container-type:inline-size}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-head{padding:8px 16px;gap:8px;border-bottom-color:var(--ob-line,#dedede)}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-chip.ok{border-color:transparent;background:color-mix(in srgb,var(--ob-blue) 16%,transparent);color:var(--ob-blue)}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-btn.primary{background:var(--ob-blue);border-color:transparent;color:#fff}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-btn.primary:hover{background:color-mix(in srgb,var(--ob-blue) 88%,#000)}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-head-mark{width:28px;height:28px;background:transparent}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-body{padding:0;gap:0;overflow:hidden}
.ob-head .ob-tabs{padding:0;gap:6px;align-items:center;border-bottom:0;background:transparent;margin-left:18px}
.ob-head .ob-tab{height:27px;padding:0 12px;border-radius:8px;border:1px solid var(--dsw-alias-border-l1);background:transparent}
.ob-head .ob-tab.on{background:var(--dsw-alias-bg-layer-1)}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-head .ob-tabs{gap:7px;padding:0;background:transparent;border-bottom:0;margin-left:18px}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-head .ob-tab{height:27px;min-width:60px;padding:0 14px;border:1px solid var(--ob-line);border-radius:5px;background:var(--ob-paper);color:var(--dsw-alias-label-secondary);box-shadow:0 1px 3px #00000012}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-head .ob-tab.on{color:var(--ob-blue);border-color:color-mix(in srgb,var(--ob-blue) 30%,var(--ob-line));background:var(--ob-paper);box-shadow:0 2px 5px #00000017;font-weight:600}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-head .ob-tab:hover{background:color-mix(in srgb,var(--ob-blue) 4%,var(--ob-paper));box-shadow:0 2px 5px #00000019}
.ob-outlook-mail .ob-body>.ob-connect{margin:8px 12px}
.ob-mail-workspace{display:flex;flex:1;flex-direction:column;min-height:0;min-width:0}
.ob-commandbar{display:flex;align-items:center;gap:4px;min-height:48px;padding:0 8px;border-bottom:1px solid var(--ob-line);background:var(--ob-paper);flex:none;overflow-x:auto}
.ob-mail-command-notice{flex:none;padding:5px 12px;border-bottom:1px solid var(--ob-line);background:var(--ob-paper);color:var(--dsw-alias-label-secondary);font-size:12px;line-height:1.5}
.ob-mail-command-notice.is-error{color:var(--dsw-alias-state-error-primary,#b42318)}
.ob-command{display:inline-flex;align-items:center;justify-content:center;gap:7px;min-height:32px;padding:0 10px;border:0;border-radius:4px;background:transparent;color:inherit;font:inherit;font-size:13px;white-space:nowrap;cursor:pointer;text-decoration:none}
.ob-command svg{flex:none}.ob-command:hover:not(:disabled){background:var(--ob-ground)}
.ob-command:disabled{opacity:.35;cursor:default}
.ob-new-mail{background:var(--ob-blue);color:white;padding:0 14px;font-weight:600;min-height:34px}
.ob-new-mail:hover:not(:disabled){background:#115ea3}
.ob-command-divider{height:23px;width:1px;background:var(--ob-line);margin:0 5px;flex:none}
.ob-outlook-mail .ob-grid-mail{grid-template-columns:204px 350px minmax(0,1fr);gap:0;overflow:hidden}
.ob-mail-folders{padding:12px 8px;overflow:auto;min-height:0;background:var(--ob-ground)}
.ob-folder-heading{display:flex;gap:8px;align-items:center;width:100%;padding:0 8px;height:30px;border:0;background:transparent;color:inherit;text-align:left;font:inherit;font-size:14px;font-weight:600;cursor:pointer}
.ob-folder-heading>span:first-child{display:flex;flex:none}.ob-folder-heading .ob-folder-name{min-width:0}
.ob-account-heading{margin-top:22px}.ob-rotate{transform:rotate(-90deg)}
.ob-outlook-mail .ob-folder{height:36px;padding:0 10px 0 28px;gap:9px;border-radius:4px;font-size:13px}
.ob-outlook-mail .ob-folder svg{width:17px;height:17px;flex:none}
.ob-outlook-mail .ob-folder.on{background:color-mix(in srgb,currentColor 8%,transparent)}
.ob-outlook-mail .ob-folder-count{font-size:12px;color:var(--ob-blue)}
.ob-outlook-mail .ob-mail-list{border-radius:0;border-top:0;border-bottom:0;border-left:1px solid var(--ob-line);border-right:1px solid var(--ob-line);min-width:0}
.ob-outlook-mail .ob-tabs2{height:46px;padding:0 16px;gap:16px;background:var(--ob-paper)}
.ob-outlook-mail .ob-tab2{height:46px;font-size:14px;padding:0 0 2px;position:relative;border:0}
.ob-outlook-mail .ob-tab2.on:after{content:"";position:absolute;bottom:6px;left:0;right:0;height:3px;border-radius:2px;background:var(--ob-blue)}
.ob-list-action{display:flex;align-items:center;justify-content:center;border:0;background:transparent;color:inherit;width:28px;height:30px;padding:0;border-radius:3px;cursor:pointer;flex:none}
.ob-list-action:hover,.ob-list-action.active{background:var(--ob-ground);color:var(--ob-blue)}
.ob-outlook-mail .ob-tabs2 .ob-spacer{min-width:0}.ob-outlook-mail .ob-tabs2 .ob-list-action+.ob-list-action{margin-left:-10px}
.ob-mail-filters{display:flex;flex-wrap:wrap;gap:12px;padding:10px 16px;border-bottom:1px solid var(--ob-line)}
.ob-mail-filters label{display:flex;align-items:center;gap:6px;font-size:12px}
.ob-mail-filters select{color:inherit;background:var(--ob-paper);border:1px solid var(--ob-line);padding:3px 4px;border-radius:4px}
.ob-outlook-mail .ob-list{scrollbar-width:thin;scrollbar-color:#999 transparent;overflow-x:hidden}
.ob-mail-dayhead{display:flex;align-items:center;gap:16px;position:sticky;top:0;z-index:1;width:100%;height:36px;padding:0 24px;border:0;border-bottom:1px solid var(--ob-line);background:var(--ob-ground);color:inherit;text-align:left;font:inherit;font-size:12px;font-weight:600;cursor:pointer}
.ob-mail-dayhead.collapsed svg{transform:rotate(-90deg)}
.ob-outlook-mail .ob-mail{min-height:81px;padding:10px 16px 10px 30px;gap:9px;position:relative;background:var(--ob-paper);border-bottom:1px solid var(--ob-line)}
.ob-outlook-mail .ob-mail:hover{background:color-mix(in srgb,var(--ob-blue) 5%,var(--ob-paper))}
.ob-outlook-mail .ob-mail.sel{outline:1px solid var(--ob-blue);outline-offset:-1px;background:color-mix(in srgb,var(--ob-blue) 4%,var(--ob-paper))}
.ob-outlook-mail .ob-mail.unread{box-shadow:inset 3px 0 0 var(--ob-blue)}
.ob-mail-avatar{width:32px;height:32px;font-size:13px;font-weight:500;margin-top:2px}
.ob-outlook-mail .ob-from{font-weight:400;font-size:14px;max-width:100%;line-height:20px}
.ob-outlook-mail .ob-mail.unread .ob-from{font-weight:600}
.ob-subject-line{display:flex;align-items:baseline;gap:10px;min-width:0}
.ob-outlook-mail .ob-subj{min-width:0;flex:1;font-size:13px;line-height:19px;margin:0}
.ob-outlook-mail .ob-time{font-size:11px;color:var(--dsw-alias-label-secondary,#555)}
.ob-outlook-mail .ob-mail.unread .ob-subj,.ob-outlook-mail .ob-mail.unread .ob-time{color:var(--ob-blue);font-weight:600}
.ob-outlook-mail .ob-prev{font-size:13px;line-height:19px;margin:0}
.ob-reading-pane{min-height:0;min-width:0;overflow:auto;padding:22px;background:var(--ob-ground)}
.ob-welcome{box-sizing:border-box;min-height:100%;display:flex;align-items:center;justify-content:center;padding:28px 12px;text-align:center;color:color-mix(in srgb,var(--dsw-alias-label-primary,#242424) 60%,transparent)}
.ob-welcome-content{width:min(100%,460px)}
.ob-welcome-weather{font-size:13px;line-height:1.7}
.ob-welcome-temperature{display:block;margin:2px 0;font-size:clamp(23px,2.7cqw,36px);font-weight:500;line-height:1.35}
.ob-welcome-heading{margin:24px 0 6px;font-size:clamp(22px,2.6cqw,34px);font-weight:600;line-height:1.3}
.ob-welcome-copy{margin:0;font-size:13px;line-height:1.7}
.ob-welcome-options{display:flex;gap:12px;align-items:center;justify-content:center;flex-wrap:wrap;margin-top:8px;font-size:11px}
.ob-welcome-options button,.ob-welcome-options a{border:0;background:transparent;color:inherit;font:inherit;text-decoration:underline;text-underline-offset:3px;cursor:pointer;padding:3px 0}
.ob-welcome-options button:disabled{cursor:default;text-decoration:none}
.ob-welcome-search{display:flex;justify-content:center;align-items:center;flex-wrap:wrap;gap:6px;margin:8px auto 0}
.ob-welcome-search input{width:min(180px,100%);height:28px;border:1px solid var(--ob-line);border-radius:4px;background:var(--ob-paper);color:var(--dsw-alias-label-primary);padding:0 8px;font:inherit;font-size:12px}
.ob-welcome-cities{display:flex;justify-content:center;flex-wrap:wrap;gap:5px;margin-top:7px}
.ob-welcome-cities button{height:auto;min-height:26px;white-space:normal;text-align:left}
.ob-welcome-status{margin:6px 0 0;font-size:11px;line-height:1.5}
.ob-reading-card{background:var(--ob-paper);border:1px solid var(--ob-line);border-radius:5px;padding:24px;max-width:960px;margin:0 auto}
.ob-reading-title{font-size:20px;line-height:1.4;font-weight:600;margin:0 0 24px;overflow-wrap:anywhere}
.ob-reading-sender{display:flex;align-items:center;gap:12px}.ob-reading-sender .ob-command{font-size:22px;margin-left:auto}
.ob-reading-identity{display:flex;flex-direction:column;min-width:0;gap:3px}.ob-reading-identity strong{font-size:14px;font-weight:600}.ob-reading-identity span{font-size:12px;color:var(--dsw-alias-label-tertiary);overflow-wrap:anywhere}
.ob-reading-date{font-size:12px;color:var(--dsw-alias-label-tertiary);margin:12px 0 24px}
.ob-reading-content{font-size:14px;line-height:1.9;white-space:pre-wrap;overflow-wrap:anywhere;min-height:100px}
.ob-reading-links{margin-top:16px;font-size:12px}.ob-reading-links a{color:var(--ob-blue)}
.ob-reading-actions{border-top:1px solid var(--ob-line);padding-top:16px;margin-top:24px}.ob-reading-actions .ob-command{border:1px solid var(--ob-line)}
.ob-outlook-mail button:focus-visible,.ob-outlook-mail a:focus-visible{outline:2px solid var(--ob-blue);outline-offset:-2px}
@container (max-width:1050px){.ob-outlook-mail .ob-grid-mail{grid-template-columns:180px 320px minmax(0,1fr)}.ob-reading-pane{padding:12px}.ob-reading-card{padding:18px}}
@container (max-width:800px){.ob-outlook-mail .ob-grid-mail{grid-template-columns:170px minmax(0,1fr)}.ob-reading-pane{display:none}.ob-mail-workspace.has-selection .ob-mail-list{display:none}.ob-mail-workspace.has-selection .ob-reading-pane{display:block}}
@container (max-width:500px){.ob-outlook-mail .ob-grid-mail{grid-template-columns:124px minmax(0,1fr)}.ob-outlook-mail .ob-folder{padding-left:8px}.ob-folder-heading{font-size:12px;padding:0}.ob-outlook-mail .ob-tabs2{padding:0 8px;gap:12px}.ob-outlook-mail .ob-mail{padding-left:12px}.ob-outlook-mail .ob-head{padding:8px}.ob-outlook-mail .ob-head .ob-sub{max-width:200px}}

/* Outlook calendar: left navigator and a full-height six-week month. */
.ob-calendar-workspace{display:flex;flex:1;flex-direction:column;min-height:0;overflow:hidden}
.ob-calendar-toolbar{min-height:43px;gap:5px;flex:none}
.ob-calendar-view-label{display:inline-flex;align-items:center;gap:6px;height:32px;padding:0 9px;border:1px solid var(--ob-line);border-radius:3px;background:var(--ob-ground);font-size:13px}
.ob-calendar-toolbar .ob-command.active{color:var(--ob-blue)}
.ob-calendar-hint{font-size:11px;color:var(--dsw-alias-label-tertiary);padding-right:10px;white-space:nowrap}
.ob-calendar-layout{display:grid;grid-template-columns:232px minmax(0,1fr);flex:1;min-height:0}
.sidebar-hidden .ob-calendar-layout{grid-template-columns:minmax(0,1fr)}
.ob-calendar-sidebar{overflow-y:auto;min-height:0;padding:8px 14px 16px;background:var(--ob-ground);scrollbar-width:thin}
.ob-mini-head{display:flex;align-items:center;gap:4px;padding:0 3px 8px;font-size:13px}
.ob-mini-head strong{font-weight:600;flex:1}.ob-mini-head button{border:0;color:inherit;background:transparent;font-size:20px;cursor:pointer;width:24px;height:26px}
.ob-mini-week,.ob-mini-grid{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:1px}
.ob-mini-week span{font-size:11px;line-height:25px;text-align:center;color:var(--dsw-alias-label-secondary)}
.ob-mini-day{display:flex;align-items:center;justify-content:center;padding:0;height:27px;border:1px solid transparent;background:transparent;color:inherit;font:inherit;font-size:11px;cursor:pointer;border-radius:0}
.ob-mini-day span{display:flex;align-items:center;justify-content:center;width:21px;height:21px;border-radius:50%}
.ob-mini-day.outside{color:var(--dsw-alias-label-quaternary)}
.ob-mini-day:hover{background:color-mix(in srgb,var(--ob-blue) 8%,transparent)}
.ob-mini-day.selected{border-color:var(--ob-blue);background:color-mix(in srgb,var(--ob-blue) 9%,transparent)}
.ob-mini-day.today span{color:white;background:var(--ob-blue);font-weight:600}
.ob-calendar-sidebar-divider{height:1px;background:var(--ob-line);margin:20px -6px 18px}
.ob-calendar-group{margin-bottom:20px}.ob-calendar-group-head{display:flex;align-items:center;height:32px;gap:6px;font-size:13px;font-weight:600}
.ob-calendar-collapse{display:flex;align-items:center;justify-content:center;width:18px;height:24px;border:0;background:transparent;color:inherit;cursor:pointer;padding:0}
.ob-calendar-collapse.closed svg{transform:rotate(-90deg)}
.ob-calendar-group-head label{display:flex;align-items:center;gap:9px;cursor:pointer}
.ob-calendar-sidebar input[type=checkbox]{width:14px;height:14px;accent-color:var(--ob-blue);margin:0;flex:none}
.ob-calendar-choice{display:flex;align-items:center;gap:11px;min-height:36px;padding:0 5px 0 30px;font-size:13px;cursor:pointer;border-radius:3px}
.ob-calendar-choice:hover{background:color-mix(in srgb,currentColor 5%,transparent)}
.ob-calendar-choice span{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ob-calendar-show-all{border:0;background:transparent;color:var(--ob-blue);font:inherit;font-size:12px;cursor:pointer;margin:0 0 12px 26px;padding:6px}
.ob-calendar-sidebar-empty{font-size:12px;color:var(--dsw-alias-label-tertiary);padding:8px}
.ob-calendar-tasks{margin-top:12px;border-top:1px solid var(--ob-line);padding-top:12px;font-size:12px}
.ob-calendar-tasks summary{cursor:pointer}.ob-calendar-task{display:flex;gap:8px;align-items:flex-start;padding:9px 0}
.ob-calendar-board{display:flex;flex-direction:column;min-width:0;min-height:0;border-inline-start:1px solid var(--ob-line);background:var(--ob-paper);overflow:hidden}
.sidebar-hidden .ob-calendar-board{border-inline-start:0}
.ob-calendar-month-head{display:flex;align-items:center;gap:8px;height:48px;flex:none;padding:0 12px}
.ob-calendar-today{border:1px solid var(--ob-line);border-radius:4px;min-width:90px;gap:14px;min-height:31px;font-weight:600}
.ob-calendar-nav{width:28px;height:30px;background:transparent;border:0;color:var(--dsw-alias-label-secondary);font-size:21px;cursor:pointer}
.ob-calendar-month-picker{position:relative;display:flex;align-items:center;gap:9px;font-size:19px;font-weight:600;margin-left:6px;cursor:pointer;white-space:nowrap}
.ob-calendar-month-picker input{position:absolute;inset:0;width:100%;height:100%;opacity:0;cursor:pointer}
.ob-calendar-month-picker input::-webkit-calendar-picker-indicator{position:absolute;inset:0;width:100%;height:100%;cursor:pointer}
.ob-calendar-month-picker:focus-within{outline:2px solid var(--ob-blue);outline-offset:3px}
.ob-calendar-loading{font-size:12px;color:var(--dsw-alias-label-tertiary)}
.ob-calendar-weekdays{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));flex:none;border-top:1px solid var(--ob-line);border-bottom:1px solid var(--ob-line)}
.ob-calendar-weekdays span{font-size:11px;line-height:27px;border-right:1px solid var(--ob-line);padding:0 8px}
.ob-calendar-scroll{display:flex;flex-direction:column;flex:1;min-height:0;overflow:auto;scrollbar-width:thin}
.ob-calendar-scroll .ob-calendar-weekdays{position:sticky;top:0;z-index:3;background:var(--ob-paper)}
.ob-calendar-month-grid{flex:1;min-height:648px;display:grid;grid-template-columns:repeat(7,minmax(0,1fr));grid-template-rows:repeat(6,minmax(108px,1fr))}
.ob-calendar-day{min-width:0;min-height:0;position:relative;display:flex;flex-direction:column;border-right:1px solid var(--ob-line);border-bottom:1px solid var(--ob-line);padding:3px 0 5px;background:transparent;cursor:pointer}
.ob-calendar-day.outside{background:color-mix(in srgb,currentColor 2%,transparent)}
.ob-calendar-day:hover{background:color-mix(in srgb,var(--ob-blue) 3%,transparent)}
.ob-calendar-day.selected{outline:1px solid var(--ob-blue);outline-offset:-1px;z-index:1}
.ob-calendar-date{display:flex;align-items:center;align-self:flex-start;min-height:25px;padding:0 8px;background:transparent;border:0;color:var(--dsw-alias-label-secondary);font:inherit;font-size:13px;cursor:pointer;flex:none}
.ob-calendar-date .today-dot{display:flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;background:var(--ob-blue);color:#fff;font-weight:600;margin-left:-4px}
.ob-calendar-events{display:flex;flex-direction:column;gap:2px;padding-top:2px;min-height:0;overflow:hidden}
.ob-calendar-event{display:block;text-align:left;flex:none;margin:0 7px 0 0;max-width:100%;min-height:22px;line-height:20px;border:0;border-left:4px solid;border-radius:3px;padding:0 6px;font:inherit;font-size:11px;color:var(--dsw-alias-label-secondary);overflow:hidden;white-space:nowrap;text-overflow:ellipsis;cursor:pointer}
.ob-calendar-event:hover{filter:brightness(.96)}
.ob-calendar-event{border-left-color:var(--ob-event-color,var(--ob-blue));background:color-mix(in srgb,var(--ob-event-color,var(--ob-blue)) 24%,var(--ob-paper))}
.ob-calendar-event[draggable=true]{cursor:grab}.ob-calendar-event[draggable=true]:active{cursor:grabbing}.ob-calendar-event.dragging{opacity:.38}.ob-calendar-event.drag-preview{opacity:.5;pointer-events:none;cursor:default}.ob-calendar-day.drag-over{background:color-mix(in srgb,var(--ob-blue) 10%,transparent)}.ob-calendar-day.drag-over .ob-calendar-events{min-height:24px}
.ob-calendar-more{border:0;background:transparent;text-align:left;color:var(--ob-blue);font:inherit;font-size:11px;cursor:pointer;padding:1px 8px;flex:none}
.ob-calendar-status{flex:none;padding:5px 12px;font-size:12px;border-top:1px solid var(--ob-line);color:var(--dsw-alias-label-secondary)}
.ob-calendar-status.error{color:var(--dsw-alias-state-error-primary,#b42318)}
.ob-calendar-agenda{width:min(440px,90vw);overflow:auto}.ob-calendar-agenda-item{display:flex;flex-direction:column;gap:6px;border-left:3px solid;padding:8px 12px;margin:4px 0;font-size:12px}.ob-calendar-agenda-item strong{font-size:14px}.ob-calendar-agenda-item a{color:var(--ob-blue)}
.ob-event-peek{position:fixed;z-index:1100;width:min(320px,calc(100vw - 16px));background:var(--ob-paper,#fff);color:var(--dsw-alias-label-primary,#242424);border:1px solid var(--ob-line,#ddd);border-radius:8px;box-shadow:0 8px 28px #0000002e;overflow:hidden;font-family:"Segoe UI","Microsoft YaHei",sans-serif}.ob-event-peek-accent{height:4px;background:#13a10e}.ob-event-peek-body{padding:12px 14px 8px;display:flex;flex-direction:column;gap:8px}.ob-event-peek-cat{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--dsw-alias-label-secondary,#555)}.ob-event-peek-dot{width:10px;height:10px;border-radius:2px;flex:none}.ob-event-peek-title{margin:0;font-size:18px;line-height:1.3;font-weight:600}.ob-event-peek-row{display:flex;align-items:flex-start;gap:8px;font-size:12.5px;color:var(--dsw-alias-label-secondary,#555)}.ob-event-peek-row svg{flex:none;margin-top:1px}.ob-event-peek-actions{display:flex;align-items:center;gap:4px;padding:6px 8px 10px;border-top:1px solid var(--ob-line,#eee)}.ob-event-peek-actions .ob-command{min-height:30px;padding:0 10px;font-size:12.5px}.ob-event-peek-expand{margin-left:auto}.ob-event-peek-extra{padding:0 14px 12px;font-size:12px;color:var(--dsw-alias-label-tertiary,#666);white-space:pre-wrap}.ob-day-peek{position:fixed;z-index:1100;width:min(300px,calc(100vw - 16px));max-height:min(360px,70vh);overflow:auto;background:var(--ob-paper,#fff);border:1px solid var(--ob-line,#ddd);border-radius:8px;box-shadow:0 8px 28px #0000002e;padding:8px}.ob-day-peek-head{display:flex;align-items:center;justify-content:space-between;padding:4px 6px 8px;font-weight:600;font-size:13px}.ob-day-peek-item{display:flex;flex-direction:column;gap:2px;width:100%;text-align:left;border:0;border-left:3px solid;background:transparent;padding:8px 10px;margin:2px 0;border-radius:4px;cursor:pointer;font:inherit}.ob-day-peek-item:hover{background:var(--ob-ground,#f5f5f5)}.ob-day-peek-item strong{font-size:13px}.ob-day-peek-item span{font-size:11.5px;color:var(--dsw-alias-label-tertiary,#666)}
.ob-outlook-calendar button:focus-visible,.ob-outlook-calendar a:focus-visible,.ob-outlook-calendar input:focus-visible{outline:2px solid var(--ob-blue);outline-offset:-2px}
.ob-outlook-calendar .ob-btn.primary{background:var(--ob-blue)}
.ob-outlook-calendar .ob-body-feedback{padding:5px 10px}
@container (max-width:1000px){.ob-calendar-layout{grid-template-columns:200px minmax(0,1fr)}.ob-calendar-sidebar{padding:8px}.ob-calendar-hint{display:none}.ob-calendar-month-picker{font-size:17px}.ob-calendar-choice{padding-left:26px}}
@container (max-width:700px){.ob-calendar-layout{grid-template-columns:166px minmax(0,1fr)}.ob-calendar-mini-head{font-size:12px}.ob-calendar-weekdays,.ob-calendar-month-grid{min-width:570px}.ob-calendar-sidebar{padding:6px}.ob-mini-day{height:25px}.ob-calendar-choice{padding-left:15px;font-size:12px}.ob-calendar-month-head{gap:4px;overflow-x:auto}.ob-calendar-toolbar{gap:0}}
.ob-context-menu{position:fixed;z-index:1200;min-width:214px;max-width:calc(100vw - 12px);max-height:calc(100vh - 12px);overflow:auto;padding:5px;background:var(--dsw-alias-bg-overlay,#fff);color:var(--dsw-alias-label-primary,#242424);border:1px solid var(--dsw-alias-border-l1,#ddd);border-radius:7px;box-shadow:0 8px 28px #0003;font-family:"Segoe UI","Microsoft YaHei",sans-serif}
.ob-context-menu button{display:flex;align-items:center;gap:10px;min-height:34px;width:100%;padding:6px 12px;border:0;border-radius:3px;background:transparent;color:inherit;text-align:left;font:inherit;font-size:13px;cursor:pointer}.ob-context-menu button:hover,.ob-context-menu button:focus-visible{background:var(--dsw-alias-bg-layer-2,#f1f1f1);outline:0}.ob-context-menu button:disabled{opacity:.4;cursor:default}.ob-context-menu button.danger{color:#bd2727}.ob-context-separator{height:1px;background:var(--dsw-alias-border-l1,#ddd);margin:4px 7px}
.ob-context-toast{position:fixed;z-index:1300;bottom:22px;left:50%;transform:translateX(-50%);padding:12px 18px;background:var(--dsw-alias-bg-overlay,#fff);border:1px solid var(--dsw-alias-border-l1,#ddd);box-shadow:0 4px 20px #0003;border-radius:6px;max-width:80vw}.ob-context-toast button{border:0;background:transparent;color:inherit;margin-left:14px;cursor:pointer}
.ob-event-mask{position:fixed;inset:0;z-index:100;display:flex;align-items:center;justify-content:center;padding:12px;background:#0005;backdrop-filter:blur(2px)}
.ob-event-editor{--ob-blue:#0f6cbd;--ob-paper:var(--dsw-alias-bg-layer-1,#fff);--ob-line:var(--dsw-alias-border-l1,#ddd);--ob-ground:var(--dsw-alias-bg-layer-2,#f5f5f5);display:flex;flex-direction:column;width:min(1180px,98vw);height:min(850px,96vh);margin:0;background:var(--ob-ground);color:var(--dsw-alias-label-primary,#242424);border:1px solid var(--ob-line);border-radius:6px;box-shadow:0 12px 60px #0004;overflow:hidden;font-family:"Segoe UI","Microsoft YaHei",sans-serif}
.ob-event-titlebar{display:flex;align-items:center;justify-content:space-between;min-height:40px;padding:0 16px;font-size:13px;background:var(--ob-ground);flex:none}.ob-event-titlebar button{font-size:24px;border:0;background:transparent;color:inherit;cursor:pointer;width:32px;height:32px}.ob-event-titlebar button:hover{background:#e15b5b;color:white}
.ob-event-toolbar{display:flex;align-items:center;gap:12px;min-height:42px;flex:none;margin:6px 14px 8px;padding:4px 8px;border:1px solid var(--ob-line);border-radius:8px;background:var(--ob-paper);box-shadow:0 1px 3px #0001;overflow-x:auto}.ob-event-toolbar-label{display:flex;align-items:center;gap:7px;font-size:13px;white-space:nowrap}.ob-event-toolbar select{border:0;background:transparent;font:inherit;color:inherit;padding:4px;max-width:150px}
.ob-event-layout{display:grid;grid-template-columns:minmax(0,2.1fr) minmax(270px,1fr);gap:14px;padding:0 14px 14px;flex:1;min-height:0}.ob-event-left{display:flex;flex-direction:column;gap:8px;min-width:0;overflow:auto}.ob-event-fields{border:1px solid var(--ob-line);border-radius:6px;background:var(--ob-paper);padding:10px 14px;margin:0;min-width:0;flex:none}.ob-event-field{display:flex;gap:16px;align-items:flex-start;min-height:46px}.ob-event-field>svg{flex:none;margin:14px 0 0 4px;color:var(--dsw-alias-label-secondary,#555)}.ob-event-field-content{flex:1;min-width:0;border-bottom:1px solid #8887;padding:4px 0}.ob-event-field:last-child .ob-event-field-content{border:0}
.ob-event-field-content>input:not([type=checkbox]),.ob-event-field-content>select{width:100%;height:37px;padding:4px 10px;border:0;border-radius:0;background:transparent;color:inherit;font:inherit;font-size:13px;box-sizing:border-box}.ob-event-subject .ob-event-field-content{border-bottom:2px solid var(--ob-blue)}.ob-event-subject input{font-size:17px!important}.ob-event-fields input:focus,.ob-event-fields select:focus{outline:1px solid var(--ob-blue);outline-offset:-1px}.ob-event-dates{display:flex;flex-direction:column;gap:8px;padding:6px 4px 4px}.ob-event-date-row{display:grid;grid-template-columns:40px minmax(0,1fr) auto;align-items:center;gap:8px}.ob-event-date-row>span{font-size:12px;color:var(--dsw-alias-label-secondary,#555);font-weight:600}.ob-event-date-inputs{display:flex;align-items:center;gap:8px;min-width:0;flex-wrap:wrap}.ob-event-duration{display:flex;align-items:center;gap:6px;flex-wrap:wrap;padding:2px 0 0 48px}.ob-event-duration-label{font-size:11px;color:var(--dsw-alias-label-tertiary,#777);margin-right:2px}.ob-event-duration-btn{height:28px;padding:0 10px;border:1px solid var(--ob-line,#ddd);border-radius:999px;background:transparent;color:var(--dsw-alias-label-secondary,#555);font:inherit;font-size:12px;cursor:pointer}.ob-event-duration-btn:hover:not(:disabled){background:color-mix(in srgb,var(--ob-blue) 6%,transparent);border-color:color-mix(in srgb,var(--ob-blue) 35%,var(--ob-line,#ddd))}.ob-event-duration-btn.on{background:color-mix(in srgb,var(--ob-blue) 14%,transparent);border-color:var(--ob-blue);color:var(--ob-blue);font-weight:600}.ob-event-duration-btn:disabled{opacity:.4;cursor:default}.ob-event-date-options{display:flex;align-items:center;gap:14px;padding:4px 0 2px 48px;font-size:11px;color:var(--dsw-alias-label-tertiary,#666)}.ob-event-date-options label{display:flex;align-items:center;gap:6px}.ob-event-fields input[type=checkbox]{accent-color:var(--ob-blue)}
.ob-event-body-card{display:flex;align-items:flex-start;gap:16px;padding:12px 14px;flex:1;min-height:160px;border:1px solid var(--ob-line);background:var(--ob-paper);border-radius:6px}.ob-event-body-card>svg{margin:10px 0 0 4px;flex:none;color:var(--dsw-alias-label-secondary,#555)}.ob-event-body-card textarea{align-self:stretch;flex:1;min-width:0;border:1px solid var(--ob-line);background:var(--ob-paper);color:inherit;border-radius:3px;resize:none;padding:12px;font:inherit;font-size:14px;line-height:1.7}
.ob-event-schedule{display:flex;flex-direction:column;min-height:0;border:1px solid var(--ob-line);border-radius:6px;background:var(--ob-paper);overflow:hidden}.ob-event-schedule-head{display:flex;align-items:center;gap:8px;padding:10px 8px;min-height:48px;border-bottom:1px solid var(--ob-line);flex:none}.ob-event-schedule-head strong{font-size:13px}.ob-event-schedule-head button{border:0;background:transparent;color:inherit;cursor:pointer;font-size:20px;padding:0 4px}.ob-event-timeline-status{min-height:24px;padding:4px 8px;border-bottom:1px solid var(--ob-line);font-size:10px;color:var(--dsw-alias-label-tertiary,#777);flex:none}.ob-event-timeline-scroll{overflow:auto;flex:1;min-height:0;scrollbar-width:thin}.ob-event-timeline{height:1152px;position:relative;margin-right:12px}.ob-event-hour{display:block;position:relative;width:100%;height:48px;border:0;border-bottom:1px solid var(--ob-line);background:transparent;text-align:left;padding:0;color:var(--dsw-alias-label-tertiary,#666);font-size:11px;cursor:pointer}.ob-event-hour:after{content:"";position:absolute;left:44px;right:0;top:24px;border-top:1px dashed var(--ob-line)}.ob-event-hour span{display:block;width:43px;border-right:1px solid var(--ob-line);height:100%;padding:4px 0 0 17px;box-sizing:border-box}.ob-event-hour:hover{background:color-mix(in srgb,var(--ob-blue) 5%,transparent)}
.ob-event-existing,.ob-event-draft-block,.ob-event-drag-ghost{position:absolute;left:0;right:0;padding:3px 8px;box-sizing:border-box;border-left:4px solid;border-radius:3px;overflow:hidden;font-size:11px}.ob-event-existing{opacity:.8;color:inherit;pointer-events:none}.ob-event-existing.interactive{pointer-events:auto;cursor:grab;touch-action:none}.ob-event-existing.interactive:active,.ob-event-draft-block:active{cursor:grabbing}.ob-event-existing.drag-origin{opacity:.32}.ob-event-draft-block{z-index:2;background:#48c274;color:#0b2616;border-left-color:#249a4e;box-shadow:0 3px 8px #0002;display:flex;flex-direction:column;gap:3px;pointer-events:auto;cursor:grab;touch-action:none}.ob-event-draft-block.dragging{opacity:.9;z-index:5;box-shadow:0 6px 14px #0003}.ob-event-draft-block strong{font-size:12px;white-space:nowrap}.ob-event-draft-block span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.ob-event-drag-ghost{z-index:4;pointer-events:none;opacity:.5;display:flex;flex-direction:column;gap:3px;box-shadow:0 4px 10px #0002}.ob-event-drag-ghost.draft{background:#48c274;color:#0b2616;border-left-color:#249a4e}.ob-event-hour.slot-hot{background:color-mix(in srgb,var(--ob-blue) 8%,transparent)}
.ob-event-existing.resizing{opacity:1;z-index:3}
.ob-guide-version{margin-inline-start:10px;white-space:nowrap;font-size:11px;font-weight:400;color:var(--dsw-alias-label-secondary,#666)}
.ob-plan-panel .ob-plan-empty{opacity:.55;user-select:none;-webkit-user-select:none;cursor:default}
.ob-plan-heading-details .ob-plan-summary,.ob-plan-heading-details .ob-plan-drag-hint{opacity:.55}
.ob-todo-page,.ob-plan-quick-add-layer{-webkit-user-select:none;user-select:none}
.ob-todo-page input,.ob-todo-page textarea,.ob-plan-quick-add textarea{-webkit-user-select:text;user-select:text}
.ob-plan-quick-add-layer{position:fixed;inset:0;z-index:1250}
.ob-plan-quick-add{position:absolute;width:min(310px,calc(100vw - 16px));box-sizing:border-box;border:1px solid var(--ob-line,#ddd);border-radius:10px;background:var(--ob-paper,#fff);color:var(--dsw-alias-label-primary,#242424);box-shadow:0 6px 22px #0003;font:14px/1.4 "Segoe UI","Microsoft YaHei",sans-serif}
.ob-plan-quick-add textarea{display:block;box-sizing:border-box;width:100%;height:46px;min-height:46px;padding:13px 16px;border:0;border-radius:inherit;background:transparent;color:inherit;font:inherit;outline:none;resize:none;white-space:pre-wrap;overflow-wrap:anywhere}
.ob-plan-quick-add:focus-within{border-color:var(--ob-blue,#0f6cbd)}
.ob-plan-quick-add textarea::placeholder{color:currentColor;opacity:.55;user-select:none;-webkit-user-select:none}
.ob-reading-heading{display:flex;align-items:flex-start;gap:16px;margin-bottom:24px}.ob-reading-heading .ob-reading-title{flex:1;min-width:0;margin:0}.ob-reading-heading .ob-btn{flex:none;white-space:nowrap}
.ob-event-schedule-head strong{white-space:nowrap}.ob-event-schedule-head .ob-event-timeline-status{margin-left:auto;min-height:0;min-width:0;flex:1;padding:0;border:0;text-align:right;line-height:1.4;overflow-wrap:anywhere}
.ob-event-calendar-field .ob-event-field-content{border-bottom:0}.ob-event-calendar-field .ob-event-field-content>select{border:1px solid var(--ob-line,#ccc);border-radius:5px;background:var(--ob-paper,#fff);padding:4px 10px;cursor:pointer}.ob-event-calendar-field .ob-event-field-content>select:focus{border-color:var(--ob-blue);outline:1px solid var(--ob-blue);outline-offset:0}
.ob-day-peek-item{background:color-mix(in srgb,var(--ob-event-color,#0f6cbd) 24%,var(--ob-paper,#fff));color:inherit}.ob-day-peek-item:hover{background:color-mix(in srgb,var(--ob-event-color,#0f6cbd) 32%,var(--ob-paper,#fff))}
.ob-event-confirm-mask{position:fixed;inset:0;z-index:1500;display:flex;align-items:center;justify-content:center;background:#0005;padding:20px}.ob-event-confirm-dialog{width:min(420px,100%);box-sizing:border-box;padding:24px;border:1px solid var(--ob-line,#ddd);border-radius:12px;background:var(--ob-paper,#fff);color:var(--dsw-alias-label-primary,#242424);box-shadow:0 14px 50px #0004;font:14px/1.6 "Segoe UI","Microsoft YaHei",sans-serif}.ob-event-confirm-dialog strong{font-size:18px}.ob-event-confirm-dialog p{margin:12px 0 20px}.ob-event-confirm-actions{display:flex;justify-content:flex-end;gap:8px}
.ob-event-close-notice{margin:8px 0 0 auto;width:min(100%,360px);box-sizing:border-box;padding:12px 14px;border:1px solid #f0df9d;border-radius:6px;background:#fff3cd;color:#665010;font-size:12px;line-height:1.6;white-space:normal}.ob-event-close-notice.failed{background:#fce8e8;border-color:#edb5b5;color:#a4262c}
.ob-event-existing.interactive,.ob-event-draft-block,.ob-event-resize-handle{user-select:none;-webkit-user-select:none}
.ob-event-block-track{position:absolute;inset:0 0 0 44px;pointer-events:none}
.ob-event-draft-block,.ob-event-drag-ghost.draft{border-left-color:var(--ob-event-color,#0f6cbd);background:color-mix(in srgb,var(--ob-event-color,#0f6cbd) 24%,var(--ob-paper,#fff));color:var(--dsw-alias-label-primary,#222)}
.ob-event-existing,.ob-event-drag-ghost.existing{border-left-color:var(--ob-event-color,#0f6cbd);background:color-mix(in srgb,var(--ob-event-color,#0f6cbd) 12%,var(--ob-paper,#fff))}
.ob-event-drag-ghost.existing{background:color-mix(in srgb,var(--ob-event-color,#0f6cbd) 18%,var(--ob-paper,#fff))}
.ob-event-agenda-stack{align-self:stretch;display:flex;flex:1;flex-direction:column;gap:8px;min-width:0;padding:8px;border:1px solid var(--ob-line);border-radius:3px;background:var(--ob-paper);box-sizing:border-box}
.ob-event-agenda-stack:focus-within{border-color:var(--ob-blue)}
.ob-event-agenda-stack textarea{min-height:90px;border:0;border-radius:0;padding:6px 4px;outline:none}
.ob-event-agenda-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.ob-event-agenda-panel{display:grid;gap:6px;padding:9px;border:1px solid var(--ob-line);border-radius:6px;background:var(--ob-ground)}
.ob-event-agenda-hint{margin:6px 0 0;font-size:11px;color:var(--dsw-alias-label-secondary,#555)}
.ob-event-agenda-panel .ob-ai-output{max-height:180px;overflow:auto}
.ob-event-date-picker{position:relative;display:inline-flex;align-items:center;gap:5px;height:34px;padding:3px 4px;border:1px solid var(--ob-line,#ddd);border-radius:8px;background:var(--ob-paper,#fff);box-sizing:border-box;font-size:13px;user-select:none;-webkit-user-select:none}
.ob-event-date-picker button,.ob-event-date-picker .ob-date-segment{height:26px;padding:0;border:0;border-radius:4px;background:transparent;color:inherit;font:inherit;text-align:center}
.ob-event-date-picker button{cursor:pointer}
.ob-event-date-picker>.ob-date-year,.ob-event-date-picker .ob-date-segment,.ob-event-date-picker>.ob-date-calendar{border:1px solid var(--ob-line,#ddd);border-radius:6px;background:var(--ob-paper,#fff);box-sizing:border-box}
.ob-event-date-picker button:hover,.ob-event-date-picker .ob-date-segment:focus{background:color-mix(in srgb,var(--ob-blue) 8%,transparent)}
.ob-event-date-picker .ob-date-segment-wrap{position:relative;display:inline-flex}
.ob-event-date-picker .ob-date-segment{width:43px;padding:0 17px 0 4px;user-select:none;-webkit-user-select:none}
.ob-event-date-picker .ob-date-unit{font-size:11px;opacity:.55;user-select:none;-webkit-user-select:none;pointer-events:none}
.ob-event-date-picker .ob-date-segment-wrap>.ob-date-unit{position:absolute;right:5px;top:50%;transform:translateY(-50%)}
.ob-event-date-picker .ob-date-segment::placeholder{color:inherit;opacity:.55}
.ob-event-date-picker .ob-date-segment[data-wheel-value]::placeholder{opacity:1}
.ob-event-date-picker .ob-date-segment::selection{color:inherit;background:transparent}
.ob-event-date-picker .ob-date-year{display:inline-flex;align-items:center;justify-content:center;gap:3px;width:58px}
.ob-event-date-picker .ob-date-calendar{display:grid;place-items:center;width:30px}
.ob-event-date-picker .ob-date-native{position:absolute;right:0;bottom:0;width:1px;height:1px;padding:0;border:0;opacity:0;pointer-events:none}
.ob-date-year-wheel{position:absolute;z-index:8;top:calc(100% + 4px);left:0;width:76px;height:142px;padding:56px 4px;box-sizing:border-box;overflow-y:auto;overscroll-behavior:contain;scroll-snap-type:y mandatory;scrollbar-width:none;border:1px solid var(--ob-line,#ddd);border-radius:6px;background:var(--ob-paper,#fff);box-shadow:0 4px 12px #0002}
.ob-date-year-wheel .ob-date-year-option{display:block;width:100%;height:28px;scroll-snap-align:center;color:var(--dsw-alias-label-secondary,#555)}
.ob-date-year-wheel .ob-date-year-option[aria-selected=true]{background:color-mix(in srgb,var(--ob-blue) 13%,transparent);color:var(--ob-blue);font-weight:600}
.ob-event-resize-handle{appearance:none;position:absolute;z-index:1;left:0;right:0;width:100%;height:min(7px,50%);margin:0;padding:0;border:0;border-radius:0;background:transparent;color:inherit;cursor:ns-resize;touch-action:none}
.ob-event-resize-handle.start{top:0}.ob-event-resize-handle.end{bottom:0}
.ob-event-resize-handle::after{content:"";display:block;margin:auto;width:22px;height:2px;border-radius:2px;background:currentColor;opacity:.45}
.ob-event-resize-handle:hover,.ob-event-resize-handle:focus-visible{background:color-mix(in srgb,currentColor 12%,transparent)}
.ob-event-resize-handle:focus-visible{outline:1px solid currentColor;outline-offset:-1px}
.ob-event-error,.ob-event-discard{margin:0 14px 8px;padding:8px 12px;border-radius:4px;background:#d1343810;color:var(--dsw-alias-state-error-primary,#a4262c);font-size:12px;flex:none}.ob-event-discard{display:flex;align-items:center;gap:12px;background:#fff3cd;color:#665010}.ob-event-discard button{cursor:pointer;padding:4px 10px}.ob-event-error-text{font-size:11px;color:#b42318;padding:4px 8px}.ob-event-editor button:focus-visible{outline:2px solid var(--ob-blue)}
@media(max-width:850px){.ob-event-layout{grid-template-columns:minmax(0,1fr) 260px;gap:8px;padding:0 8px 8px}.ob-event-field{gap:8px}.ob-event-fields{padding:8px}.ob-event-date-row{grid-template-columns:36px minmax(0,1fr)}.ob-event-duration{padding-left:36px}.ob-event-date-options{padding-left:36px}.ob-event-toolbar{margin:4px 8px 8px;gap:6px}}
@media(max-width:650px){.ob-event-layout{display:flex;flex-direction:column;overflow:auto}.ob-event-left{flex:none;overflow:visible}.ob-event-schedule{min-height:320px;flex:none}.ob-event-mask{padding:4px}.ob-event-body-card{min-height:170px}}

/* Raised mail/calendar switches and an in-flow reading/reply workspace. */

.ob-mail-assistant-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.ob-ai-button{color:var(--ob-blue);background:color-mix(in srgb,var(--ob-blue) 4%,var(--ob-paper))}
.ob-inline-hint{font-size:11px;line-height:1.6;color:var(--dsw-alias-label-tertiary,#777);margin:7px 0}
.ob-inline-error{font-size:12px;line-height:1.6;padding:8px 10px;margin:8px 0;border-radius:4px;background:#c52c2c0b;color:var(--dsw-alias-state-error-primary,#b42318)}
.ob-inline-success{padding:8px 0;color:#168042;font-size:12px}
.ob-ai-summary{margin-top:16px;padding:14px 16px;background:color-mix(in srgb,var(--ob-blue) 4%,var(--ob-paper));border:1px solid color-mix(in srgb,var(--ob-blue) 14%,var(--ob-line));border-radius:6px}
.ob-inline-section-head{display:flex;justify-content:space-between;align-items:center;gap:12px;font-size:13px;margin-bottom:10px}
.ob-ai-output{font-size:13px;line-height:1.8;white-space:pre-wrap;overflow-wrap:anywhere}
.ob-inline-reply{position:static;margin-top:18px;padding-top:18px;border-top:1px solid var(--ob-line)}
.ob-inline-intent{display:flex;align-items:center;gap:10px;margin:10px 0;font-size:12px;color:var(--dsw-alias-label-secondary)}
.ob-inline-intent input{flex:1;min-width:0;height:34px;padding:6px 10px;border:1px solid var(--ob-line);border-radius:4px;background:var(--ob-paper);color:inherit;font:inherit}
.ob-inline-editor{display:block;width:100%;box-sizing:border-box;min-height:220px;resize:vertical;padding:14px;border:1px solid var(--ob-line);border-radius:5px;background:var(--ob-paper);color:inherit;font:inherit;font-size:14px;line-height:1.75;outline:none}
.ob-ai-wrap{flex:none;display:flex;flex-direction:column;padding:8px 14px;background:#f3f3f3;border-bottom:1px solid var(--ob-line,#dedede);box-sizing:border-box}
.ob-ai-bar{display:flex;align-items:center;gap:8px;margin:0;padding:4px 8px 4px 12px;min-height:38px;border-radius:20px;background:#fff;border:1px solid #e8e8e8;box-shadow:0 2px 8px rgba(0,0,0,.12);box-sizing:border-box}
.ob-ai-bar-mark{flex:none;font-size:11px;font-weight:700;letter-spacing:.3px;color:#0f6cbd;padding:2px 6px;border-radius:999px;background:color-mix(in srgb,#0f6cbd 10%,transparent)}
.ob-ai-bar-input{flex:1;min-width:0;height:30px;border:0;outline:none;background:transparent;color:inherit;font:inherit;font-size:13px}
.ob-ai-bar-input::placeholder{color:#8a8a8a}
.ob-ai-bar-send{flex:none;height:28px;padding:0 12px;border:0;border-radius:999px;background:#0f6cbd;color:#fff;font:inherit;font-size:12px;cursor:pointer}
.ob-ai-bar-send:disabled{opacity:.55;cursor:default}
.ob-ai-bar-send:hover:not(:disabled){filter:brightness(1.05)}
.ob-ai-bar-msg{flex:none;max-width:42%;font-size:11.5px;line-height:1.3;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:#168042}
.ob-ai-bar-msg.err{color:#b42318}
.ob-ai-bar-msg.busy{color:#666}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-ai-wrap{padding-inline:12px;background:#f3f3f3;border-bottom-color:var(--ob-line,#dedede)}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-ai-bar{margin:0;background:#fff;border-color:#e8e8e8;box-shadow:0 2px 8px rgba(0,0,0,.12)}

.ob-inline-editor:focus,.ob-inline-intent input:focus{border-color:var(--ob-blue);box-shadow:0 0 0 1px color-mix(in srgb,var(--ob-blue) 15%,transparent)}
.ob-inline-reply-footer{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:12px}
.ob-ai-suggestion{margin-top:12px;padding:12px;border:1px solid var(--ob-line);border-radius:5px;background:var(--ob-ground);font-size:12px}.ob-ai-suggestion .ob-ai-output{margin:8px 0}
.ob-reading-pane .ob-reading-card{overflow:visible}

/* Paint each region once so a solid workspace cannot hide local transparency. */
.ob-root{
  --ob-paper-base:var(--dsw-alias-bg-layer-1,#fff);
  --ob-ground-base:var(--dsw-alias-bg-layer-2,#f5f5f5);
  --ob-paper:color-mix(in srgb,var(--ob-paper-base) var(--ob-panel-opacity,100%),transparent);
  --ob-ground:color-mix(in srgb,var(--ob-ground-base) var(--ob-panel-opacity,100%),transparent)
}
.ob-root :is(.ob-head,.ob-ai-wrap,.ob-commandbar,.ob-mail-folders,.ob-mail-list,.ob-reading-pane,.ob-calendar-sidebar,.ob-calendar-board,.ob-plan-header,.ob-plan-panel){
  --ob-region-opacity:calc(var(--ob-panel-opacity,100%) * var(--ob-region-strength,1));
  --ob-paper:color-mix(in srgb,var(--ob-paper-base) var(--ob-region-opacity),transparent);
  --ob-ground:color-mix(in srgb,var(--ob-ground-base) var(--ob-region-opacity),transparent);
  backdrop-filter:var(--ob-region-filter,var(--ob-panel-filter,none))
}
.ob-root :is(.ob-head,.ob-ai-wrap,.ob-mail-folders,.ob-reading-pane){background:var(--ob-ground)}
.ob-body-feedback{flex:none;display:grid;gap:8px;padding:8px 12px;background:var(--ob-ground);backdrop-filter:var(--ob-panel-filter,none)}
.ob-body-feedback>.ob-banner{background:transparent;border:1px solid var(--ob-line,#dedede)}
.ob-body-feedback:has(+.ob-todo-page){padding:4px 14px}
.ob-body-feedback:has(+.ob-todo-page)>.ob-banner{padding-block:4px}
@container(max-width:570px){.ob-body-feedback:has(+.ob-todo-page){padding-inline:8px}}
.ob-root .ob-mail-command-notice{backdrop-filter:var(--ob-panel-filter,none)}
.ob-root .ob-ai-bar{background:var(--ob-paper)}
.ob-root .ob-mail-list{background:var(--ob-paper)}
.ob-root .ob-mail-list :is(.ob-tabs2,.ob-mail:not(:hover):not(.sel)){background:transparent}
.ob-root:not(.ob-outlook-mail,.ob-outlook-calendar) .ob-body{background:var(--ob-paper);backdrop-filter:var(--ob-panel-filter,none)}
.ob-appearance-mask{position:absolute;box-sizing:border-box;padding:12px;background:#0002;backdrop-filter:none}
@container(max-width:800px){.ob-inline-intent{align-items:flex-start;flex-direction:column}.ob-inline-intent input{width:100%;box-sizing:border-box}.ob-inline-editor{min-height:190px}}

`

    function Sheet() {
      return h('style', { 'data-outlook-bridge': 'styles' }, CSS)
    }

    /* ------------------------------------------------------------- glyph */

    function PanelGlyph(props) {
      const size = typeof props?.size === 'number' ? props.size : 18
      const snap = useSnapshot()
      const unread = Number(snap?.unread || 0)
      return h(
        'span',
        { className: 'ob-glyph' },
        h(Sheet, null),
        mark(size),
        unread > 0
          ? h('span', { className: 'ob-badge', 'aria-hidden': 'true' }, unread > 99 ? '99+' : String(unread))
          : null,
      )
    }

    /* -------------------------------------------------------- mail block */

    function MailIcon({ name, size = 18 }) {
      const paths = {
        save: 'M4 3h14l3 3v15H3V3h1ZM7 3v6h10V3M7 21v-8h10v8',
        clock: 'M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18ZM12 7v5h4',
        check: 'm5 12 4 4L19 6',
        people: 'M9 3a3 3 0 1 0 0 6 3 3 0 0 0 0-6ZM3 21v-5a4 4 0 0 1 4-4h4a4 4 0 0 1 4 4v5M17 4a3 3 0 0 1 0 6M18 13a4 4 0 0 1 3 4v4',
        location: 'M12 22s8-8 8-13a8 8 0 0 0-16 0c0 5 8 13 8 13ZM12 6a3 3 0 1 0 0 6 3 3 0 0 0 0-6Z',
        video: 'M3 5h12v14H3V5Zm12 5 6-4v12l-6-4',
        notes: 'M4 3h16v18H4V3ZM8 7h8M8 11h8M8 15h5',
        panel: 'M4 4h16v16H4z',
        header: 'M4 4h16v16H4zM4 9h16',
        columns: 'M4 4h16v16H4zM10 4v16',
        calendar: 'M3 5h18v16H3V5ZM3 10h18M7 2v6M17 2v6M7 14h3M14 14h3M7 18h3',
        refresh: 'M20 7A9 9 0 0 0 4 6L2 9m0-6v6h6M4 17a9 9 0 0 0 16 1l2-3m0 6v-6h-6',
        compose: 'M13 4H5a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2h13a2 2 0 0 0 2-2v-8M16 3l5 5M10 14l2-5 7-7 5 5-7 7-5 2',
        inbox: 'M4 4h16l2 12v4H2v-4L4 4ZM2 15h6l2 3h4l2-3h6',
        sent: 'm3 3 19 9-19 9 4-9-4-9Zm4 9h15',
        drafts: 'M4 5h9M4 10h6M4 15h4m4 3 2-5 6-6 3 3-6 6-5 2Z',
        deleteditems: 'M3 6h18M9 6V3h6v3M5 6l1 15h12l1-15M10 10v7M14 10v7',
        archive: 'M3 3h18v5H3V3Zm2 5v13h14V8M9 12h6',
        junkemail: 'M12 2 3 6v6c0 5 9 10 9 10s9-5 9-10V6l-9-4ZM12 7v6m0 3v1',
        mail: 'M3 5h18v14H3V5Zm0 1 9 7 9-7',
        read: 'm3 9 9-7 9 7v12H3V9Zm0 0 9 7 9-7',
        reply: 'm9 5-7 7 7 7M2 12h12c5 0 7 3 7 7',
        filter: 'M3 6h18M6 12h12M10 18h4',
        sort: 'M7 3v18m-4-4 4 4 4-4M17 21V3m-4 4 4-4 4 4',
        chevron: 'm7 10 5 5 5-5',
        edit: 'M4 20h4L18 10l-4-4L4 16v4Zm11-13 4 4',
        trash: 'M3 6h18M9 6V3h6v3M5 6l1 15h12l1-15M10 10v7M14 10v7',
        attachment: 'm8 13 7-7a3 3 0 0 1 4 4l-9 9a5 5 0 0 1-7-7l9-9',
      }
      return h('svg', { width: size, height: size, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 1.25, strokeLinecap: 'round', strokeLinejoin: 'round', 'aria-hidden': true },
        h('path', { d: paths[name] || paths.mail }))
    }

    function Overlay({ children }) {
      return ReactDOM.createPortal(children, document.body)
    }


    function shanghaiTodayParts(base = new Date()) {
      const fmt = new Intl.DateTimeFormat('en-CA', {
        timeZone: 'Asia/Shanghai',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hourCycle: 'h23',
        weekday: 'short',
      })
      return Object.fromEntries(fmt.formatToParts(base).filter((p) => p.type !== 'literal').map((p) => [p.type, p.value]))
    }

    function addDaysYmd(ymd, days) {
      const date = new Date(`${ymd}T12:00:00+08:00`)
      date.setTime(date.getTime() + days * 86400000)
      const parts = shanghaiTodayParts(date)
      return `${parts.year}-${parts.month}-${parts.day}`
    }

    const EVENT_DURATION_MS = 30 * 60_000
    const EVENT_BUFFER_MS = 15 * 60_000

    function calendarTimestamp(stamp) {
      return new Date(/(?:Z|[+-]\d{2}:\d{2})$/i.test(stamp || '')
        ? stamp
        : `${String(stamp || '').slice(0, 19)}+08:00`).getTime()
    }

    function eventSlotFields(start) {
      const fieldsAt = (instant) => {
        const { year, month, day, hour, minute } = shanghaiTodayParts(new Date(instant))
        return { date: `${year}-${month}-${day}`, time: `${hour}:${minute}` }
      }
      const from = fieldsAt(start)
      const to = fieldsAt(start + EVENT_DURATION_MS)
      return { date: from.date, startTime: from.time, endDate: to.date, endTime: to.time }
    }

    function nextEventSlot(start, events) {
      const intervals = events.filter((event) => !event.allDay)
        .map((event) => ({ start: calendarTimestamp(event.start), end: calendarTimestamp(event.end) }))
        .filter((event) => event.end > event.start)
        .sort((a, b) => a.start - b.start)
      for (const event of intervals) {
        if (event.start >= start + EVENT_DURATION_MS) break
        if (event.end > start) start = Math.ceil((event.end + EVENT_BUFFER_MS) / 60_000) * 60_000
      }
      return start
    }

    function fetchTimelineDay(from) {
      const to = new Date(from.getTime() + 86400000)
      return api(`/calendar?from=${encodeURIComponent(from.toISOString())}&to=${encodeURIComponent(to.toISOString())}`)
    }

    function addMinutesHm(hm, minutes) {
      const match = String(hm || '').match(/^([01]?\d|2[0-3]):([0-5]\d)$/)
      if (!match) return hm
      let total = Number(match[1]) * 60 + Number(match[2]) + minutes
      if (total < 0) total = 0
      if (total >= 24 * 60) total = 24 * 60 - 1
      return `${String(Math.floor(total / 60)).padStart(2, '0')}:${String(total % 60).padStart(2, '0')}`
    }

    function parseAiCommandFallback(text) {
      const raw = String(text || '').trim()
      const parts = shanghaiTodayParts()
      const today = `${parts.year}-${parts.month}-${parts.day}`
      const taskPrefix = /^(?:(?:请\s*)?(?:帮我\s*)?(?:添加|新建|创建|记下)\s*(?:一个|一项|一条)?\s*(?:计划|待办|任务)(?!清单|列表)|(?:计划|待办|任务)(?=\s*[:：]))/u.exec(raw)
      if (taskPrefix) {
        const input = raw.slice(taskPrefix[0].length).replace(/^[\s:：,，]+/u, '')
        if (!input) return { intent: 'unknown', reply: '请输入要添加的计划标题', fallback: true }
        const task = parsePlanTitle(input, Number(parts.year))
        const relative = /^(今天|今日|明天|后天)[\s:：,，]*(.+)$/u.exec(task.title)
        if (!task.dueDate && relative) {
          task.dueDate = addDaysYmd(today, { 今天: 0, 今日: 0, 明天: 1, 后天: 2 }[relative[1]])
          task.title = relative[2].trim()
        }
        return { intent: 'task', task, fallback: true }
      }
      const lower = raw.toLowerCase()
      const emailMatch = raw.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i)
      const wantsMail = /邮件|写信|发给|写给|发送给|发信|compose|email|mail/i.test(raw) || Boolean(emailMatch && /写|发|给/.test(raw))
      const wantsCal = /日程|开会|会议|提醒|约会|日历|安排|预约|见面/i.test(raw)

      if (wantsMail) {
        let subject = ''
        let body = ''
        const about = raw.match(/(?:关于|主题(?:是|为)?|标题(?:是|为)?)\s*[「"']?([^」"'\n]+)[」"']?/)
        if (about) subject = about[1].trim()
        if (!subject) {
          const confirm = raw.match(/写一封(.{0,40}?)(?:的)?邮件/)
          if (confirm) subject = confirm[1].replace(/给\s*\S+@\S+\s*/, '').trim() || '邮件'
        }
        if (!subject) subject = '邮件'
        body = raw
        return {
          intent: 'mail',
          mail: { to: emailMatch ? emailMatch[0] : '', subject, body },
          calendar: null,
          reply: 'AI 暂不可用，已尽量填好表单',
          fallback: true,
        }
      }

      if (wantsCal) {
        let date = today
        if (/明天/.test(raw)) date = addDaysYmd(today, 1)
        else if (/后天/.test(raw)) date = addDaysYmd(today, 2)
        else if (/今天|今日/.test(raw)) date = today
        else {
          const md = raw.match(/(\d{1,2})\s*月\s*(\d{1,2})\s*日/)
          if (md) {
            date = `${parts.year}-${String(md[1]).padStart(2, '0')}-${String(md[2]).padStart(2, '0')}`
          }
        }

        let startTime = '09:00'
        const hm = raw.match(/(?:下午|晚上)?\s*(\d{1,2})\s*(?:[:：点时])\s*(\d{1,2})?/)
        if (hm) {
          let hour = Number(hm[1])
          const minute = hm[2] != null ? Number(hm[2]) : 0
          if (/下午|晚上/.test(raw) && hour < 12) hour += 12
          if (/中午/.test(raw) && hour < 12) hour = 12
          if (/早上|上午|清晨/.test(raw) && hour === 12) hour = 0
          startTime = `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`
        }

        let duration = 30
        const durH = raw.match(/(\d+(?:\.\d+)?)\s*(?:个)?\s*小时/)
        const durM = raw.match(/(\d+)\s*分钟/)
        if (durH) duration = Math.round(Number(durH[1]) * 60)
        else if (durM) duration = Number(durM[1])
        else if (/半小时|半个小时/.test(raw)) duration = 30
        else if (/一小时|1小时/.test(raw)) duration = 60

        const endTime = addMinutesHm(startTime, duration)
        let subject = raw
          .replace(/明天|后天|今天|今日/g, '')
          .replace(/\d{1,2}\s*月\s*\d{1,2}\s*日/g, '')
          .replace(/(?:下午|晚上|早上|上午|中午|清晨)?\s*\d{1,2}\s*(?:[:：点时])\s*\d{0,2}\s*分?/g, '')
          .replace(/(\d+(?:\.\d+)?)\s*(?:个)?\s*小时|\d+\s*分钟|半小时|半个小时/g, '')
          .replace(/帮我|请|一下|安排|添加|新建|创建/g, '')
          .trim() || '会议'
        subject = subject.replace(/^(开会|会议|日程|提醒)/, (m) => m).slice(0, 80)

        const loc = raw.match(/(?:地点|在)\s*[「"']?([^，。；\s"']+)[」"']?/)
        return {
          intent: 'calendar',
          calendar: {
            subject,
            date,
            endDate: date,
            startTime,
            endTime,
            allDay: /全天/.test(raw),
            location: loc ? loc[1] : '',
            body: '',
          },
          mail: null,
          reply: 'AI 暂不可用，已尽量填好表单',
          fallback: true,
        }
      }

      return {
        intent: 'unknown',
        calendar: null,
        mail: null,
        reply: 'AI 暂不可用，请尝试「添加计划：明天交周报」、「明天下午3点开会」或「给 xxx@yy.com 写邮件」',
        fallback: true,
      }
    }

    function openContextMenu(event, items) {
      event.preventDefault()
      event.stopPropagation()
      const rect = event.currentTarget?.getBoundingClientRect?.()
      window.dispatchEvent(new CustomEvent('outlook-context-menu', { detail: {
        x: event.clientX || rect?.left || 8, y: event.clientY || rect?.bottom || 8, items,
      } }))
    }

    async function copyOutlookText(text) {
      if (!navigator.clipboard?.writeText) throw new Error('当前窗口不支持复制，请手动选择文本复制。')
      await navigator.clipboard.writeText(String(text || ''))
    }

    function ContextMenu() {
      const [menu, setMenu] = React.useState(null)
      const [failure, setFailure] = React.useState('')
      const ref = React.useRef(null)
      React.useEffect(() => {
        const open = (event) => { setFailure(''); setMenu(event.detail) }
        window.addEventListener('outlook-context-menu', open)
        return () => window.removeEventListener('outlook-context-menu', open)
      }, [])
      React.useEffect(() => {
        if (!menu || !ref.current) return
        const oldFocus = document.activeElement
        const element = ref.current
        const rect = element.getBoundingClientRect()
        element.style.left = `${Math.max(6, Math.min(menu.x, window.innerWidth - rect.width - 6))}px`
        element.style.top = `${Math.max(6, Math.min(menu.y, window.innerHeight - rect.height - 6))}px`
        element.querySelector('button:not(:disabled)')?.focus()
        const close = () => setMenu(null)
        const outside = (event) => { if (!element.contains(event.target)) close() }
        const scroll = (event) => { if (!element.contains(event.target)) close() }
        const key = (event) => {
          if (event.key === 'Escape') { event.preventDefault(); close(); oldFocus?.focus?.() }
          if (event.key === 'Tab') close()
          if (['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) {
            event.preventDefault()
            const buttons = Array.from(element.querySelectorAll('button:not(:disabled)'))
            const index = buttons.indexOf(document.activeElement)
            const next = event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length - 1 : (index + (event.key === 'ArrowUp' ? -1 : 1) + buttons.length) % buttons.length
            buttons[next]?.focus()
          }
        }
        document.addEventListener('pointerdown', outside, true)
        document.addEventListener('keydown', key, true)
        window.addEventListener('resize', close)
        window.addEventListener('scroll', scroll, true)
        return () => {
          document.removeEventListener('pointerdown', outside, true)
          document.removeEventListener('keydown', key, true)
          window.removeEventListener('resize', close)
          window.removeEventListener('scroll', scroll, true)
        }
      }, [menu])
      return h(Overlay, null,
        menu ? h('div', { ref, className: 'ob-context-menu', role: 'menu', 'aria-label': '快捷操作',
          style: { left: menu.x, top: menu.y }, onContextMenu: (event) => event.preventDefault() },
          menu.items.filter(Boolean).map((item, index) => item.separator
            ? h('div', { className: 'ob-context-separator', role: 'separator', key: index })
            : h('button', { type: 'button', role: 'menuitem', key: index, disabled: item.disabled,
                className: item.danger ? 'danger' : '', onClick: async () => {
                  setMenu(null)
                  try { await item.run() } catch (error) { setFailure(error.message || String(error)) }
                } }, h(MailIcon, { name: item.icon || 'calendar', size: 16 }), item.label))) : null,
        failure ? h('div', { className: 'ob-context-toast', role: 'alert' }, failure,
          h('button', { type: 'button', onClick: () => setFailure(''), 'aria-label': '关闭提示' }, '×')) : null)
    }

    function eventMonthLength(year, month) {
      const end = new Date(0)
      end.setUTCFullYear(year, month, 0)
      return end.getUTCDate()
    }

    function eventUnitLabel(unit) {
      return h('span', { className: 'ob-date-unit', 'aria-hidden': true }, unit)
    }

    function EventNumberSegment({ label, unit, value, min, max, wheelMin = min, wheelMax = max, onInteract, onChange }) {
      const [editing, setEditing] = React.useState(null)
      const inputRef = React.useRef(null)
      const cancelInput = React.useRef(false)
      React.useEffect(() => {
        const input = inputRef.current
        if (!input) return
        const onWheel = (event) => {
          if (input.ownerDocument.activeElement !== input || event.ctrlKey ||
            Math.abs(event.deltaY) <= Math.abs(event.deltaX)) return
          event.preventDefault()
          const direction = Math.sign(event.deltaY)
          setEditing((current) => {
            const draft = current?.text || current?.wheelValue
            const next = Math.max(wheelMin, Math.min(wheelMax, Number(draft ?? value) + direction))
            return { text: '', wheelValue: next }
          })
        }
        input.addEventListener('wheel', onWheel, { passive: false })
        return () => input.removeEventListener('wheel', onWheel)
      }, [value, wheelMin, wheelMax])

      const wheelValue = editing?.wheelValue
      return h('span', { className: 'ob-date-segment-wrap' }, h('input', {
        ref: inputRef, className: 'ob-date-segment', type: 'text', inputMode: 'numeric', autoComplete: 'off',
        'aria-label': label, 'data-wheel-value': wheelValue,
        value: editing?.text ?? pad(value), placeholder: pad(wheelValue ?? value),
        onFocus: () => { onInteract(); setEditing({ text: '' }) },
        onChange: (event) => setEditing({ text: event.target.value.replace(/\D/g, '').slice(0, 2) }),
        onSelect: (event) => {
          const input = event.currentTarget
          if (input.selectionStart !== input.selectionEnd) input.setSelectionRange(input.selectionEnd, input.selectionEnd)
        },
        onBlur: (event) => {
          const entered = Number(event.currentTarget.value || wheelValue)
          setEditing(null)
          if (cancelInput.current) { cancelInput.current = false; return }
          if (entered >= min && entered <= max && entered !== value) onChange(entered)
        },
        onKeyDown: (event) => {
          if (event.key !== 'Enter' && event.key !== 'Escape') return
          event.preventDefault()
          event.stopPropagation()
          cancelInput.current = event.key === 'Escape'
          event.currentTarget.blur()
        },
      }), eventUnitLabel(unit))
    }

    function EventNativePicker({ type, label, value, min, onOpen, onChange }) {
      const picker = React.useRef(null)
      return h(React.Fragment, null,
        h('button', {
          className: 'ob-date-calendar', type: 'button', 'aria-label': label,
          onClick: () => { onOpen(); picker.current?.showPicker?.() },
        }, h(MailIcon, { name: type === 'date' ? 'calendar' : 'clock', size: 14 })),
        h('input', {
          className: 'ob-date-native', type, ref: picker, value, min, tabIndex: -1, 'aria-hidden': true,
          onChange: (event) => {
            const next = event.target.value
            if (next && next !== value && (!min || next >= min)) onChange(next)
          },
        }))
    }

    function EventDatePicker({ label, value, min, onChange, onInteract }) {
      const [year, month, day] = value.split('-').map(Number)
      const [wheelYear, setWheelYear] = React.useState(null)
      const [highlightedYear, setHighlightedYear] = React.useState(year)
      const wheel = React.useRef(null)
      const yearButton = React.useRef(null)
      const [earliestYear, earliestMonth, earliestDay] = (min || '').split('-').map(Number)
      const rowHeight = 28
      const minYear = Math.max(min ? Number(min.slice(0, 4)) : 1, wheelYear - 100)
      const maxYear = Math.min(9999, wheelYear + 100)
      const wheelId = `ob-event-${min ? 'end' : 'start'}-year-wheel`
      React.useEffect(() => {
        if (wheelYear == null || !wheel.current) return
        wheel.current.scrollTop = (wheelYear - minYear) * rowHeight
        wheel.current.focus({ preventScroll: true })
      }, [wheelYear, minYear])

      const apply = (nextYear, nextMonth, nextDay) => {
        const date = `${String(nextYear).padStart(4, '0')}-${pad(nextMonth)}-${pad(Math.min(nextDay, eventMonthLength(nextYear, nextMonth)))}`
        const next = min && date < min ? min : date
        if (next !== value) onChange(next)
      }
      const closeWheel = () => {
        setWheelYear(null)
        yearButton.current?.focus()
      }
      const chooseYear = (nextYear) => {
        apply(nextYear, month, day)
        closeWheel()
      }
      const readWheelYear = (node) => Math.max(minYear, Math.min(maxYear, minYear + Math.round(node.scrollTop / rowHeight)))
      const onWheelKey = (event) => {
        if (event.key === 'Escape' || event.key === 'Enter' || event.key === ' ') {
          event.preventDefault()
          event.stopPropagation()
          if (event.key === 'Escape') closeWheel()
          else chooseYear(readWheelYear(event.currentTarget))
          return
        }
        const step = { ArrowUp: -1, ArrowDown: 1, PageUp: -5, PageDown: 5 }[event.key]
        if (!step) return
        event.preventDefault()
        const next = Math.max(minYear, Math.min(maxYear, readWheelYear(event.currentTarget) + step))
        event.currentTarget.scrollTop = (next - minYear) * rowHeight
        setHighlightedYear(next)
      }
      const interactWithDate = () => {
        onInteract()
        setWheelYear(null)
      }
      return h('div', {
        className: 'ob-event-date-picker', role: 'group', 'aria-label': `${label}日期`,
        onBlur: (event) => { if (!event.currentTarget.contains(event.relatedTarget)) setWheelYear(null) },
      },
        h('button', {
          ref: yearButton, className: 'ob-date-year', type: 'button', title: '滚动选择年份',
          'aria-label': `${label}年份 ${year}`, 'aria-haspopup': 'listbox', 'aria-expanded': wheelYear != null,
          'aria-controls': wheelYear != null ? wheelId : undefined,
          onClick: () => {
            if (wheelYear != null) { closeWheel(); return }
            onInteract()
            setHighlightedYear(year)
            setWheelYear(year)
          },
        }, String(year).padStart(4, '0'), eventUnitLabel('年')),
        h(EventNumberSegment, {
          label: `${label}月份`, unit: '月', value: month, min: 1, max: 12,
          wheelMin: earliestYear === year ? earliestMonth : 1,
          onInteract: interactWithDate, onChange: (next) => apply(year, next, day),
        }),
        h(EventNumberSegment, {
          label: `${label}日期中的日`, unit: '日', value: day, min: 1, max: 31,
          wheelMin: earliestYear === year && earliestMonth === month ? earliestDay : 1,
          wheelMax: eventMonthLength(year, month),
          onInteract: interactWithDate, onChange: (next) => apply(year, month, next),
        }),
        h(EventNativePicker, { type: 'date', label: `打开${label}日期日历`, value, min, onOpen: interactWithDate, onChange }),
        wheelYear != null ? h('div', {
          ref: wheel, className: 'ob-date-year-wheel', id: wheelId, role: 'listbox', tabIndex: 0,
          'aria-label': `${label}年份`, 'aria-activedescendant': `${wheelId}-${highlightedYear}`,
          onScroll: (event) => setHighlightedYear(readWheelYear(event.currentTarget)),
          onKeyDown: onWheelKey,
        }, Array.from({ length: maxYear - minYear + 1 }, (_, index) => {
          const option = minYear + index
          return h('button', {
            className: 'ob-date-year-option', type: 'button', key: option, id: `${wheelId}-${option}`,
            tabIndex: -1, role: 'option', 'aria-selected': option === highlightedYear,
            onMouseDown: (event) => event.preventDefault(),
            onClick: () => chooseYear(option),
          }, String(option).padStart(4, '0'))
        })) : null)
    }

    function EventTimePicker({ label, value, onChange, onInteract }) {
      const [hour, minute] = value.split(':').map(Number)
      return h('div', { className: 'ob-event-date-picker', role: 'group', 'aria-label': `${label}时间` },
        h(EventNumberSegment, {
          label: `${label}小时`, unit: '时', value: hour, min: 0, max: 23, onInteract,
          onChange: (next) => onChange(`${pad(next)}:${pad(minute)}`),
        }),
        h(EventNumberSegment, {
          label: `${label}分钟`, unit: '分', value: minute, min: 0, max: 59, onInteract,
          onChange: (next) => onChange(`${pad(hour)}:${pad(next)}`),
        }),
        h(EventNativePicker, { type: 'time', label: `打开${label}时间选择器`, value, onOpen: onInteract, onChange }))
    }

    function EventAgendaEditor({ subject, body, durationMinutes, saving, patch }) {
      const [instruction, setInstruction] = React.useState('')
      const [draft, setDraft] = React.useState(null)
      const [busy, setBusy] = React.useState(false)
      const [error, setError] = React.useState('')
      const request = React.useRef(null)
      const cancel = () => {
        const pending = request.current
        request.current = null
        pending?.abort()
      }
      React.useEffect(() => cancel, [])
      const cancelGeneration = () => {
        cancel()
        setBusy(false)
      }
      const canGenerate = !saving && !busy && Boolean(instruction.trim() || subject.trim() || body.trim())
      const generate = async () => {
        if (!canGenerate || request.current) return
        const controller = new AbortController()
        request.current = controller
        setBusy(true)
        setError('')
        setDraft(null)
        try {
          const result = await api('/calendar/ai/draft', {
            method: 'POST',
            signal: controller.signal,
            body: { instruction, subject, body: body.slice(0, 6000), durationMinutes },
          })
          if (request.current === controller) setDraft(result)
        } catch (failure) {
          if (request.current === controller) setError(failure.message)
        } finally {
          if (request.current === controller) {
            request.current = null
            setBusy(false)
          }
        }
      }
      const applyDraft = (includeTitle) => {
        if (!draft || saving || busy) return
        patch(includeTitle ? { subject: draft.subject, body: draft.body } : { body: draft.body })
        setDraft(null)
      }
      return h('div', { className: 'ob-event-agenda-stack' },
        h('div', { className: 'ob-ai-bar', role: 'group', 'aria-label': 'AI 帮写日程和议程' },
            h('span', { className: 'ob-ai-bar-mark', 'aria-hidden': true }, 'AI'),
            h('input', {
              className: 'ob-ai-bar-input',
              value: instruction, maxLength: 2000, disabled: busy || saving,
              'aria-label': '活动要点',
              placeholder: '试试：项目周会，回顾进度、讨论风险、确认下周分工',
              title: '参考活动要点、当前标题、说明和时长生成草稿',
              onChange: (event) => setInstruction(event.target.value),
              onKeyDown: (event) => {
                if (event.key !== 'Enter' || event.nativeEvent?.isComposing) return
                event.preventDefault()
                void generate()
              },
            }),
            h('button', { type: 'button', className: 'ob-ai-bar-send', disabled: !canGenerate, onClick: generate }, busy ? '生成中…' : '生成草稿')),
          busy ? h('button', { type: 'button', className: 'ob-btn', onClick: cancelGeneration }, '取消生成') : null,
          busy ? h('p', { className: 'ob-event-agenda-hint', role: 'status' }, '正在起草标题和议程…') : null,
          error ? h('div', { className: 'ob-inline-error', role: 'alert' }, error) : null,
          draft ? h('div', { className: 'ob-event-agenda-panel' },
            h('strong', null, draft.subject),
            h('div', { className: 'ob-ai-output' }, draft.body),
            h('div', { className: 'ob-event-agenda-actions' },
              h('button', { type: 'button', className: 'ob-btn', disabled: saving, onClick: () => applyDraft(true) }, '采用标题和议程'),
              h('button', { type: 'button', className: 'ob-btn', disabled: saving, onClick: () => applyDraft(false) }, '仅采用议程')),
            subject.trim() || body.trim()
              ? h('p', { className: 'ob-event-agenda-hint' }, '采用草稿会替换对应字段。') : null) : null,
        h('textarea', { disabled: saving, 'aria-label': '活动说明', placeholder: '添加说明或议程…',
          value: body, onChange: (event) => patch({ body: event.target.value }) }))
    }

    function EventEditor({ form, setForm, calendars, saving, error, onSave, onDelete, onDeleteTimelineEvent, onClose, onTimelineChange, autoTime, defaultCalendarAction }) {
      const timeline = React.useRef(null)
      const dialog = React.useRef(null)
      const [timelineData, setTimelineData] = React.useState([])
      const [timelineStatus, setTimelineStatus] = React.useState('')
      const [closeError, setCloseError] = React.useState('')
      const [manualEdited, setManualEdited] = React.useState(false)
      const [confirmClose, setConfirmClose] = React.useState(false)
      const initial = React.useRef(form)
      const autoTimePending = React.useRef(autoTime)
      const [checkingTime, setCheckingTime] = React.useState(Boolean(autoTime))
      const writable = calendars.filter((calendar) => calendar.canEdit !== false)
      const sourceCalendar = form.eventId
        ? calendars.find((calendar) => calendar.id === form.sourceCalendarId) || form.sourceCalendar
        : null
      const selectedCalendar = calendars.find((calendar) => calendar.id === form.calendarId) ||
        (sourceCalendar?.id === form.calendarId ? sourceCalendar : null)
      const readOnly = sourceCalendar?.canEdit === false
      const editorTitle = readOnly ? '查看日程' : form.eventId ? '编辑日程' : '新建活动'
      const calendarOptions = readOnly ? [] : writable
      const includeSelected = !calendarOptions.some((calendar) => calendar.id === form.calendarId)
      const movingCalendar = Boolean(form.eventId && form.sourceCalendarId !== form.calendarId)
      const calendarColor = selectedCalendar?.color || '#0f6cbd'
      const timelineEventColor = (item) => calendars.find((calendar) => calendar.id === item.calendarId)?.color || item.color || '#0f6cbd'
      const stopAutoTime = () => {
        if (!autoTimePending.current) return
        autoTimePending.current = false
        setCheckingTime(false)
      }
      const patch = (values, fromTimeline = false) => {
        if (!fromTimeline) setManualEdited(true)
        setConfirmClose(false)
        setCloseError('')
        if (['date', 'endDate', 'startTime', 'endTime', 'allDay'].some((key) => key in values)) stopAutoTime()
        setForm((old) => ({ ...old, ...values }))
      }
      const editableValues = (value) => {
        const { timelineChanges, timelineSaved, ...fields } = value
        return JSON.stringify(fields)
      }
      const currentChanged = editableValues(form) !== editableValues(initial.current)
      const pendingTimeline = Object.keys(form.timelineChanges || {}).length > 0
      const close = () => {
        if (saving) return
        if (!currentChanged && !pendingTimeline) { onClose(); return }
        if (!form.eventId || manualEdited) { setConfirmClose(true); return }
        const saveCurrent = currentChanged
        if (saveCurrent && !canSave) {
          setCloseError('暂时无法保存，请检查标题、起止时间和所属日历后再关闭。')
          return
        }
        setCloseError('')
        void onSave({ saveCurrent })
      }
      const scrollToStart = (time) => {
        if (timeline.current) timeline.current.scrollTop = Math.max(0, (Number(time.slice(0,2)) - 2) * 48)
      }
      React.useEffect(() => {
        const old = document.activeElement
        dialog.current?.querySelector(readOnly ? 'button' : 'input[name=subject]')?.focus()
        scrollToStart(form.startTime)
        return () => old?.focus?.()
      }, [])
      React.useEffect(() => {
        let alive = true
        setTimelineData([])
        if (!/^\d{4}-\d{2}-\d{2}$/.test(form.date)) { setTimelineStatus('请选择日期'); return }
        setTimelineStatus('正在加载当天日程…')
        const from = new Date(`${form.date}T00:00:00+08:00`)
        fetchTimelineDay(from)
          .then(async (data) => {
            if (!alive) return
            setTimelineData(data.events ?? [])
            const warnings = [...(data.warnings ?? [])]
            setTimelineStatus(warnings.join('；'))
            if (!autoTimePending.current) return

            let start = nextEventSlot(calendarTimestamp(`${form.date}T${form.startTime}:00`), data.events ?? [])
            let checkedUntil = from.getTime() + 86400000
            for (let day = 0; day < 6 && start + EVENT_DURATION_MS > checkedUntil; day += 1) {
              const slotDay = calendarTimestamp(`${eventSlotFields(start).date}T00:00:00`)
              const nextFrom = Math.max(checkedUntil, slotDay)
              const nextDay = await fetchTimelineDay(new Date(nextFrom))
              if (!alive || !autoTimePending.current) return
              warnings.push(...(nextDay.warnings ?? []))
              start = nextEventSlot(start, nextDay.events ?? [])
              checkedUntil = nextFrom + 86400000
            }
            if (start + EVENT_DURATION_MS > checkedUntil) throw new Error('暂未找到空闲时段，请手动选择时间')
            const slot = eventSlotFields(start)
            initial.current = { ...initial.current, ...slot }
            setForm((old) => ({ ...old, ...slot }))
            scrollToStart(slot.startTime)
            stopAutoTime()
            setTimelineStatus([...new Set(warnings)].join('；'))
          })
          .catch((failure) => {
            if (!alive) return
            setTimelineStatus(`${autoTimePending.current ? '自动选时失败' : '当天日程加载失败'}：${failure.message}`)
            stopAutoTime()
          })
        return () => { alive = false }
      }, [form.date])
      const keyboard = (event) => {
        if (event.key === 'Escape') { event.preventDefault(); close() }
        if (event.key !== 'Tab') return
        const focusable = Array.from(dialog.current.querySelectorAll('button:not(:disabled), input:not(:disabled), select:not(:disabled), textarea:not(:disabled), a[href], [tabindex="0"]'))
          .filter((element) => element.tabIndex >= 0)
        const first = focusable[0], last = focusable[focusable.length - 1]
        if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus() }
        if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus() }
      }
      const start = new Date(`${form.date}T${form.allDay ? '00:00' : form.startTime}:00+08:00`)
      const end = new Date(`${form.endDate}T${form.allDay ? '23:59' : form.endTime}:00+08:00`)
      const validRange = Number.isFinite(+start) && Number.isFinite(+end) && end > start
      const canSave = !saving && !readOnly && !checkingTime && validRange && Boolean(form.subject.trim()) &&
        Boolean(selectedCalendar && selectedCalendar.canEdit !== false)
      const mins = (value) => { const [hour, minute] = value.split(':').map(Number); return hour*60+minute }
      const startMinute = form.allDay ? 0 : mins(form.startTime)
      const endMinute = form.allDay || form.endDate > form.date ? 1440 : mins(form.endTime)
      const changeDate = (date) => {
        const days = Math.max(0, Math.round((new Date(form.endDate) - new Date(form.date))/86400000)) || 0
        const next = new Date(`${date}T12:00:00`)
        if (!Number.isFinite(+next)) return
        next.setDate(next.getDate() + days)
        patch({ date, endDate: dayKey(next) })
      }
      const shiftDate = (direction) => { const date = new Date(`${form.date}T12:00:00`); date.setDate(date.getDate()+direction); changeDate(dayKey(date)) }
      const PX_PER_MIN = 0.8
      const SNAP_MIN = 15
      const [timelineDrag, setTimelineDrag] = React.useState(null)
      const timelineDragRef = React.useRef(null)
      const timelineDragCleanup = React.useRef(null)
      const suppressHourClick = React.useRef(false)
      const movingExisting = React.useRef(false)
      const timelineDate = React.useRef(form.date)
      timelineDate.current = form.date
      const labelOf = (minute) => `${pad(Math.floor(Math.max(0, minute) / 60) % 24)}:${pad(Math.floor(Math.max(0, minute) % 60))}`
      const pointOf = (minute) => ({
        date: addDaysYmd(form.date, Math.floor(minute / 1440)),
        time: labelOf(minute),
      })
      const snapMinute = (value) => Math.round(value / SNAP_MIN) * SNAP_MIN
      const clampStart = (start, duration) => Math.max(0, Math.min(1440 - Math.max(SNAP_MIN, duration), start))
      const clientYToMinute = (clientY) => {
        const scroll = timeline.current
        const track = scroll?.querySelector?.('.ob-event-timeline')
        if (!track) return 0
        const y = clientY - track.getBoundingClientRect().top
        return y / PX_PER_MIN
      }
      const realTimelineId = (item) => {
        const id = String(item?.id || '')
        // List rows use `${calendarId}:${graphEventId}`; calendarId itself has no colon.
        const cut = id.indexOf(':')
        return cut >= 0 ? id.slice(cut + 1) : id
      }
      const isEditedTimelineItem = (item) => {
        const editId = String(form.eventId || '').trim()
        const listKey = String(form.listEventId || '').trim()
        if (!editId && !listKey) return false
        const full = String(item?.id || '')
        const raw = realTimelineId(item)
        if (listKey && (full === listKey || raw === listKey)) return true
        if (!editId) return false
        if (raw === editId || full === editId) return true
        if (full.endsWith(`:${editId}`)) return true
        // Graph sometimes returns ids that only match as a suffix of the composite key.
        if (raw && editId && (raw.endsWith(editId) || editId.endsWith(raw))) return true
        if (form.calendarId && item?.calendarId && form.calendarId === item.calendarId && raw && editId && raw === editId) return true
        return false
      }
      const wallMinute = (stamp) => {
        const offset = (calendarTimestamp(stamp) - calendarTimestamp(`${form.date}T00:00:00`)) / 60_000
        return Number.isFinite(offset) ? offset : null
      }
      const beginTimelineDrag = (event, spec) => {
        if (event.button !== 0) return
        event.preventDefault()
        event.stopPropagation()
        if (timelineDragRef.current || saving || (spec.mode === 'draft' && form.allDay)) return
        if (spec.mode === 'existing' && movingExisting.current) {
          setTimelineStatus('上一条调整正在保存，请稍候再拖动')
          return
        }
        if (spec.mode === 'draft') stopAutoTime()
        suppressHourClick.current = false
        const pointer = clientYToMinute(event.clientY)
        const anchor = spec.edge === 'end' ? spec.originEnd : spec.originStart
        const next = {
          ...spec,
          grabOffset: pointer - anchor,
          followStart: spec.originStart,
          previewStart: spec.originStart,
          previewEnd: spec.originEnd,
          pointerId: event.pointerId,
        }
        timelineDragRef.current = next
        setTimelineDrag(next)
        // Track the whole gesture at window capture phase. Pointer capture is
        // only an optimization: losing it must not discard an active drag.
        const host = timeline.current?.ownerDocument?.defaultView || window
        let lastPoint = { pointerId: event.pointerId, clientY: event.clientY, buttons: event.buttons }
        const move = (point) => {
          if (point.pointerId !== next.pointerId) return
          lastPoint = point
          onTimelinePointerMove(point)
        }
        const up = (point) => {
          if (point.pointerId !== next.pointerId) return
          // Include the release position even if the browser coalesced the last move.
          onTimelinePointerMove(point)
          finishTimelineDrag(point)
        }
        const cancel = () => clearTimelineDrag({ pointerId: next.pointerId })
        const scroll = () => onTimelinePointerMove(lastPoint)
        const track = timeline.current
        host.addEventListener('pointermove', move, true)
        host.addEventListener('pointerup', up, true)
        host.addEventListener('pointercancel', clearTimelineDrag, true)
        host.addEventListener('blur', cancel)
        track?.addEventListener('scroll', scroll)
        timelineDragCleanup.current = () => {
          host.removeEventListener('pointermove', move, true)
          host.removeEventListener('pointerup', up, true)
          host.removeEventListener('pointercancel', clearTimelineDrag, true)
          host.removeEventListener('blur', cancel)
          track?.removeEventListener('scroll', scroll)
          timelineDragCleanup.current = null
        }
        // Capture on the stable scroll container, not a handle/card that can
        // change position or disappear while React updates the drag preview.
        try { timeline.current?.setPointerCapture(event.pointerId) } catch (_) {}
      }
      const resizePreview = (drag, minute) => {
        const minDuration = 5
        if (drag.edge === 'start') {
          const max = Math.min(1440 - minDuration, Math.floor(drag.originEnd - minDuration))
          return { ...drag, previewStart: max < 0 ? drag.originStart : Math.max(0, Math.min(max, minute)) }
        }
        const min = Math.max(0, Math.ceil(drag.originStart + minDuration))
        return { ...drag, previewEnd: min > 1440 ? drag.originEnd : Math.max(min, Math.min(1440, minute)) }
      }
      const onTimelinePointerMove = (event) => {
        const drag = timelineDragRef.current
        if (!drag || drag.pointerId !== event.pointerId) return
        const raw = clientYToMinute(event.clientY) - drag.grabOffset
        let next
        if (drag.edge) {
          const origin = drag.edge === 'start' ? drag.originStart : drag.originEnd
          next = Math.abs(raw - origin) * PX_PER_MIN < 2
            ? { ...drag, previewStart: drag.originStart, previewEnd: drag.originEnd }
            : resizePreview(drag, snapMinute(raw))
        } else {
          const followStart = clampStart(raw, drag.duration)
          const previewStart = clampStart(snapMinute(raw), drag.duration)
          next = { ...drag, followStart, previewStart }
        }
        timelineDragRef.current = next
        setTimelineDrag(next)
      }
      const clearTimelineDrag = (event) => {
        const drag = timelineDragRef.current
        if (!drag || drag.pointerId !== event.pointerId) return null
        timelineDragRef.current = null
        timelineDragCleanup.current?.()
        setTimelineDrag(null)
        try {
          if (timeline.current?.hasPointerCapture(event.pointerId)) timeline.current.releasePointerCapture(event.pointerId)
        } catch (_) {}
        return drag
      }
      React.useEffect(() => () => {
        timelineDragCleanup.current?.()
        const drag = timelineDragRef.current
        timelineDragRef.current = null
        try { if (drag && timeline.current?.hasPointerCapture(drag.pointerId)) timeline.current.releasePointerCapture(drag.pointerId) } catch (_) {}
      }, [])
      const commitTimelineDrag = async (drag) => {
        const start = drag.previewStart
        const end = drag.edge ? drag.previewEnd : start + drag.duration
        const startChanged = drag.edge !== 'end' && start !== drag.originStart
        const endChanged = drag.edge === 'end' && end !== drag.originEnd
        if (!startChanged && !endChanged) return
        const startPoint = pointOf(start)
        const endPoint = pointOf(end)
        const startFields = { date: startPoint.date, startTime: startPoint.time }
        const endFields = { endDate: endPoint.date, endTime: endPoint.time }
        if (drag.mode === 'draft') {
          patch({
            ...(drag.edge !== 'end' ? startFields : {}),
            ...(drag.edge !== 'start' ? endFields : {}),
          }, true)
          return
        }
        const updated = {
          ...drag.event,
          ...(drag.edge !== 'end' ? { start: `${startPoint.date}T${startPoint.time}:00` } : {}),
          ...(drag.edge !== 'start' ? { end: `${endPoint.date}T${endPoint.time}:00` } : {}),
        }
        setForm((old) => ({ ...old, timelineChanges: { ...old.timelineChanges, [updated.id]: updated } }))
        setTimelineStatus('调整已暂存')
      }
      const finishTimelineDrag = (event) => {
        const drag = clearTimelineDrag(event)
        if (!drag) return
        // Swallow the click that browsers synthesize after a drag.
        suppressHourClick.current = true
        setTimeout(() => { suppressHourClick.current = false }, 0)
        void commitTimelineDrag(drag)
      }
      const resizeHandle = (spec, edge) => {
        const minute = edge === 'start' ? spec.originStart : spec.originEnd
        if (!Number.isFinite(minute) || minute < 0 || minute > 1440 ||
          (edge === 'start' ? minute === 1440 : minute === 0)) return null
        const label = edge === 'start' ? '调整开始时间' : '调整结束时间'
        const resize = { ...spec, edge }
        return h('button', {
          type: 'button',
          className: `ob-event-resize-handle ${edge}`,
          title: `拖动${label}`,
          'aria-label': label,
          onPointerDown: (event) => beginTimelineDrag(event, resize),
          onKeyDown: (event) => {
            if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return
            event.preventDefault()
            if (saving || timelineDragRef.current || (spec.mode === 'existing' && movingExisting.current)) return
            const direction = event.key === 'ArrowUp' ? -1 : 1
            const next = (direction < 0 ? Math.ceil(minute / SNAP_MIN) : Math.floor(minute / SNAP_MIN)) * SNAP_MIN + direction * SNAP_MIN
            void commitTimelineDrag(resizePreview({
              ...resize, previewStart: spec.originStart, previewEnd: spec.originEnd,
            }, next))
          },
        })
      }

      const selectHour = (hour) => {
        if (readOnly) return
        if (suppressHourClick.current || timelineDragRef.current) { suppressHourClick.current = false; return }
        const duration = !form.allDay && validRange ? Math.max(SNAP_MIN, endMinute - startMinute) : 30
        const start = hour * 60
        const end = pointOf(Math.min(1440, start + duration))
        patch({ allDay: false, startTime: labelOf(start), endTime: end.time, endDate: end.date })
      }
      const DURATION_PRESETS = [15, 30, 60, 120]
      const currentDuration = (!form.allDay && form.endDate === form.date)
        ? (mins(form.endTime) - mins(form.startTime))
        : null
      const applyDuration = (minutes) => {
        if (form.allDay || saving) return
        const start = mins(form.startTime)
        if (!Number.isFinite(start)) return
        const total = start + minutes
        const dayOffset = Math.floor(total / 1440)
        const endMin = total % 1440
        const next = new Date(`${form.date}T12:00:00`)
        if (!Number.isFinite(+next)) return
        next.setDate(next.getDate() + dayOffset)
        patch({
          allDay: false,
          endDate: dayKey(next),
          endTime: labelOf(endMin),
        })
      }
      const timelineBlocks = timelineData.map((item) => form.timelineChanges?.[item.id] || form.timelineSaved?.[item.id] || item).filter((item) => !isEditedTimelineItem(item)).flatMap((item) => {
        const start = wallMinute(item.start), end = wallMinute(item.end)
        if (start == null || end == null) return []
        const dragging = timelineDrag?.mode === 'existing' && timelineDrag.event.id === item.id
        const resizing = dragging && timelineDrag.edge
        const top = Math.max(0, resizing ? timelineDrag.previewStart : start)
        const bottom = Math.min(1440, resizing ? timelineDrag.previewEnd : end)
        return bottom > top ? [{ item, start, end, top, bottom, dragging, resizing }] : []
      })
      const draftDuration = Math.max(SNAP_MIN, endMinute - startMinute)
      const draggingDraft = timelineDrag?.mode === 'draft'
      const resizingDraft = draggingDraft && timelineDrag.edge
      const draftStart = resizingDraft ? timelineDrag.previewStart : draggingDraft ? timelineDrag.followStart : startMinute
      const draftEnd = resizingDraft ? timelineDrag.previewEnd : draggingDraft ? draftStart + draftDuration : endMinute
      const draftBlock = validRange ? { top: draftStart, bottom: Math.min(1440, draftEnd) } : null
      const ghostBlock = timelineDrag && !timelineDrag.edge
        ? { top: timelineDrag.previewStart, bottom: Math.min(1440, timelineDrag.previewStart + timelineDrag.duration) }
        : null
      const blocks = timelineBlocks.filter(({ dragging, resizing }) => !dragging || resizing)
      if (draftBlock) blocks.push(draftBlock)
      if (ghostBlock && !draggingDraft) blocks.push(ghostBlock)

      // An overlap group shares its columns. A new group starts at the previous group's end.
      let group = { lanes: [], end: 0 }
      for (const block of blocks.sort((a, b) => a.top - b.top)) {
        if (block.top >= group.end) {
          group.nextTop = block.top
          group = { lanes: [], end: block.bottom }
        }
        const available = group.lanes.findIndex((last) => last.bottom <= block.top)
        const lane = available < 0 ? group.lanes.length : available
        if (available >= 0) group.lanes[lane].nextTop = block.top
        block.lane = lane
        block.group = group
        group.lanes[lane] = block
        group.end = Math.max(group.end, block.bottom)
      }
      const laneStyle = (block) => {
        const count = block?.group?.lanes.length ?? 1
        const lane = block?.lane ?? 0
        return { left: `${lane / count * 100}%`, right: `${(count - lane - 1) / count * 100}%` }
      }
      const blockStyle = (block, minHeight) => {
        const limit = (Math.min(block.nextTop ?? 1440, block.group?.nextTop ?? 1440) - block.top) * PX_PER_MIN
        const height = Math.min(Math.max(minHeight, (block.bottom - block.top) * PX_PER_MIN), limit)
        return { ...laneStyle(block), top: block.top * PX_PER_MIN, height, ...(height < 18 ? { paddingBlock: 0 } : {}) }
      }
      const field = (icon, children, cls = '', onContextMenu) => h('div', { className: `ob-event-field ${cls}`, onContextMenu },
        h(MailIcon, { name: icon, size: 19 }), h('div', { className: 'ob-event-field-content' }, children))
      return h(Overlay, null, h('div', { className: 'ob-event-mask' },
        h('form', { ref: dialog, className: 'ob-event-editor', role: 'dialog', 'aria-modal': true, 'aria-label': editorTitle, onKeyDown: keyboard,
          onSubmit: (event) => { event.preventDefault(); if (canSave) onSave() } },
          h('header', { className: 'ob-event-titlebar' }, h('span', null, `${editorTitle} - 日历`),
            h('button', { type: 'button', disabled: saving, onClick: close, 'aria-label': `关闭${editorTitle}` }, '×')),
          h('div', { className: 'ob-event-toolbar' },
            h('button', { type: 'submit', className: 'ob-command ob-new-mail', disabled: !canSave },
              h(MailIcon, { name: 'save' }), saving ? '保存中…' : movingCalendar ? '移到所选日历' : form.attendees.trim() ? '发送邀请' : '保存'),
            form.eventId && onDelete
              ? h('button', { type: 'button', className: 'ob-command', disabled: saving || readOnly, onClick: onDelete },
                  h(MailIcon, { name: 'trash', size: 16 }), '删除日程')
              : null,
            h('label', { className: 'ob-event-toolbar-label' }, h(MailIcon, { name: 'clock', size: 16 }),
              h('select', { 'aria-label': '提醒', disabled: saving || readOnly, value: form.reminder, onChange: (event) => patch({ reminder: event.target.value }) },
                [['-1','不提醒'],['0','开始时'],['5','提前 5 分钟'],['15','提前 15 分钟'],['30','提前 30 分钟'],['60','提前 1 小时'],['1440','提前 1 天']].map(([value,label])=>h('option',{key:value,value},label))))),
          h('div',{className:'ob-event-layout'},
            h('div',{className:'ob-event-left'},
              h('fieldset',{className:'ob-event-fields',disabled:saving || readOnly},
                field('calendar', h('select',{
                  value:form.calendarId, 'aria-label':'保存到日历', disabled:readOnly,
                  title:readOnly ? '所属日历为只读' : undefined,
                  onChange:(event)=>patch({calendarId:event.target.value}),
                },
                  includeSelected ? h('option',{value:form.calendarId,disabled:!selectedCalendar || readOnly},
                    selectedCalendar?.name || (form.eventId ? '所属日历不可用' : writable.length ? '请选择日历' : '没有可写入的日历')) : null,
                  calendarOptions.map((item)=>h('option',{key:item.id,value:item.id},item.name))),
                  'ob-event-calendar-field', selectedCalendar && defaultCalendarAction
                    ? (event) => openContextMenu(event, [defaultCalendarAction(selectedCalendar)]) : undefined),
                field('compose',h('input',{name:'subject',required:true,maxLength:255,placeholder:'添加标题','aria-label':'活动标题',value:form.subject,onChange:(event)=>patch({subject:event.target.value})}),'ob-event-subject'),
                field('people',h('input',{placeholder:'邀请必需的与会者（邮箱，多个用逗号分隔）','aria-label':'与会者邮箱',value:form.attendees,onChange:(event)=>patch({attendees:event.target.value})})),
                field('clock',[
                  h('div',{className:'ob-event-dates',key:'dates'},
                    h('div',{className:'ob-event-date-row',key:'start'},
                      h('span',null,'开始'),
                      h('div',{className:'ob-event-date-inputs'},
                        h(EventDatePicker,{label:'开始',value:form.date,onChange:changeDate,onInteract:stopAutoTime}),
                        form.allDay ? null : h(EventTimePicker,{label:'开始',value:form.startTime,onChange:(time)=>patch({startTime:time}),onInteract:stopAutoTime}))),
                    h('div',{className:'ob-event-date-row',key:'end'},
                      h('span',null,'结束'),
                      h('div',{className:'ob-event-date-inputs'},
                        h(EventDatePicker,{label:'结束',value:form.endDate,min:form.date,onChange:(date)=>patch({endDate:date}),onInteract:stopAutoTime}),
                        form.allDay ? null : h(EventTimePicker,{label:'结束',value:form.endTime,onChange:(time)=>patch({endTime:time}),onInteract:stopAutoTime}))),
                    form.allDay ? null : h('div',{className:'ob-event-duration',key:'duration','aria-label':'时长预设'},
                      h('span',{className:'ob-event-duration-label'},'时长'),
                      DURATION_PRESETS.map((minutes)=>h('button',{
                        type:'button',
                        key:minutes,
                        className:`ob-event-duration-btn${currentDuration === minutes ? ' on' : ''}`,
                        disabled:saving,
                        'aria-pressed': currentDuration === minutes,
                        onClick:()=>applyDuration(minutes),
                      }, minutes === 120 ? '2 小时' : `${minutes} 分钟`)))),
                  h('div',{className:'ob-event-date-options',key:'options'},h('label',null,h('input',{type:'checkbox',checked:form.allDay,onChange:(event)=>patch({allDay:event.target.checked})}),'全天'),h('span',null,'北京时间 (UTC+08:00)')),
                  !validRange ? h('span',{className:'ob-event-error-text',key:'error'},'结束时间需要晚于开始时间') : null]),
                field('location',h('input',{placeholder:'添加位置','aria-label':'位置',value:form.location,onChange:(event)=>patch({location:event.target.value})})),
                error || closeError ? h('div',{className:'ob-event-error',role:'alert'},error || closeError) : null,
                confirmClose ? h(Overlay,null,h('div',{className:'ob-event-confirm-mask',onKeyDown:(event)=>{
                  event.stopPropagation()
                  if(event.key === 'Escape' && !saving) { event.preventDefault(); setConfirmClose(false) }
                }},h('div',{className:'ob-event-confirm-dialog',role:'alertdialog','aria-modal':true,'aria-label':'保存日程修改'},
                  h('strong',null,'保存日程修改？'),
                  h('p',null,saving ? '正在保存，请稍候…' : '有未保存的修改。点击保存后提交，或放弃本次修改。'),
                  error || closeError ? h('p',{role:'alert'},error || closeError) : null,
                  h('div',{className:'ob-event-confirm-actions'},
                    h('button',{type:'button',className:'ob-btn',disabled:saving,onClick:()=>setConfirmClose(false),autoFocus:true},'继续编辑'),
                    h('button',{type:'button',className:'ob-btn',disabled:saving,onClick:onClose},'放弃'),
                    h('button',{type:'button',className:'ob-btn primary',disabled:!canSave,onClick:()=>onSave()},saving ? '保存中…' : '保存'))))) : null),
              h('div',{className:'ob-event-body-card'},h(MailIcon,{name:'notes',size:19}),
                h(EventAgendaEditor, {
                  subject: form.subject, body: form.body, saving: saving || readOnly, patch,
                  durationMinutes: !form.allDay && validRange ? Math.round((end - start) / 60_000) : null,
                }))),
            h('aside',{className:'ob-event-schedule'},
              h('div',{className:'ob-event-schedule-head'},h('button',{type:'button',disabled:saving || readOnly,'aria-label':'前一天',onClick:()=>shiftDate(-1)},'‹'),h(MailIcon,{name:'calendar',size:17}),
                h('button',{type:'button',disabled:saving || readOnly,'aria-label':'后一天',onClick:()=>shiftDate(1)},'›'),h('strong',null,form.date.replaceAll('-',' / ')),
                h('div',{className:'ob-event-timeline-status',role:'status'},checkingTime ? '正在查找可用时段…' : timelineStatus)),
              h('div',{className:'ob-event-timeline-scroll',ref:timeline,
                onDragStart:(event)=>event.preventDefault(),
              },h('div',{className:'ob-event-timeline'},
                Array.from({length:24},(_,hour)=>h('button',{key:hour,type:'button',className:`ob-event-hour${timelineDrag && Math.floor((timelineDrag.previewStart ?? 0)/60)===hour ? ' slot-hot' : ''}`,disabled:saving,'aria-label':`从 ${hour}:00 开始`,onClick:()=>selectHour(hour)},h('span',null,hour))),
                h('div',{className:'ob-event-block-track'},timelineBlocks.map((block)=>{
                  const { item, start: a0, end: b0, dragging, resizing } = block
                  const canDelete = calendars.find((calendar) => calendar.id === item.calendarId)?.canEdit !== false
                  const editable = !item.allDay && canDelete
                  const resize = { mode: 'existing', event: item, originStart: a0, originEnd: b0 }
                  return h('div',{
                    key:item.id,
                    className:`ob-event-existing${editable ? ' interactive' : ''}${resizing ? ' resizing' : dragging ? ' drag-origin' : ''}`,
                    title: editable ? `${item.subject}（拖动边缘可调整时长）` : item.subject,
                    style:{...blockStyle(block, 18),'--ob-event-color':timelineEventColor(item),pointerEvents:'auto'},
                    onClick: editable || saving ? undefined : (event) => {
                      const hour = Math.floor(clientYToMinute(event.clientY) / 60)
                      if (hour >= 0 && hour < 24) selectHour(hour)
                    },
                    onContextMenu: (event) => openContextMenu(event, [
                      { label: '复制日程标题', icon: 'notes', run: () => copyOutlookText(item.subject) },
                      { separator: true },
                      { label: '删除日程', icon: 'trash', danger: true, disabled: !canDelete || saving || movingExisting.current, run: async () => {
                        if (await onDeleteTimelineEvent(item, true)) {
                          setTimelineData((rows) => rows.filter((row) => row.id !== item.id))
                          setForm((old) => { const changes = { ...old.timelineChanges }; delete changes[item.id]; return { ...old, timelineChanges: changes } })
                        }
                      } },
                    ]),
                    onPointerDown: editable ? (event) => beginTimelineDrag(event, {
                      mode: 'existing',
                      event: item,
                      duration: Math.max(SNAP_MIN, b0 - a0),
                      originStart: Math.max(0, a0),
                    }) : undefined,
                  }, resizing ? `${labelOf(timelineDrag.previewStart)} - ${labelOf(timelineDrag.previewEnd)} · ${item.subject}` : item.subject,
                    editable && resizeHandle(resize, 'start'),
                    editable && resizeHandle(resize, 'end'))
                }),
                draftBlock ? (() => {
                  const resize = { mode: 'draft', originStart: startMinute, originEnd: (end - new Date(`${form.date}T00:00:00+08:00`)) / 60000 }
                  const editable = !form.allDay && !saving && !readOnly
                  return h('div',{
                    className:`ob-event-draft-block${draggingDraft ? ' dragging' : ''}`,
                    style:{...blockStyle(draftBlock, 24),'--ob-event-color':calendarColor},
                    onPointerDown: editable ? (event) => beginTimelineDrag(event, {
                      mode: 'draft',
                      duration: draftDuration,
                      originStart: startMinute,
                    }) : undefined,
                  },
                    h('strong',null,form.allDay?'全天':`${draggingDraft ? labelOf(draftStart) : form.startTime} - ${draggingDraft && resizingDraft !== 'start' ? (draftEnd === 1440 ? '次日 00:00' : labelOf(draftEnd)) : (form.endDate>form.date?`${form.endDate} `:'') + form.endTime}`),
                    h('span',null,form.subject||'新活动'),
                    editable && resizeHandle(resize, 'start'),
                    editable && resizeHandle(resize, 'end'))
                })() : null,
                ghostBlock ? h('div',{
                  className:`ob-event-drag-ghost ${timelineDrag.mode === 'draft' ? 'draft' : 'existing'}`,
                  style:{
                    ...blockStyle(ghostBlock, 24),
                    ...(draggingDraft ? laneStyle(draftBlock) : {}),
                    '--ob-event-color': draggingDraft ? calendarColor : timelineEventColor(timelineDrag.event),
                  },
                },
                  h('strong',null,`${labelOf(timelineDrag.previewStart)} - ${labelOf((timelineDrag.previewStart ?? 0) + Math.max(SNAP_MIN, timelineDrag.duration))}`),
                  h('span',null,timelineDrag.mode === 'draft' ? (form.subject||'新活动') : (timelineDrag.event?.subject || ''))
                ) : null))))))))
    }

    function MailRow(props) {
      const { row, active, t, onSelect, time } = props
      const color = avatarColor(row.from)
      return h(
        'button',
        {
          type: 'button',
          className: `ob-mail${row.isRead ? '' : ' unread'}${active ? ' sel' : ''}`,
          onClick: () => onSelect(row),
          onContextMenu: (event) => props.onContextRow?.(event, row),
          'aria-pressed': Boolean(active),
        },
        h('span', { className: 'ob-mail-avatar', style: { background: color, color: '#48423d' } }, initial(row.from)),
        h(
          'span',
          { className: 'ob-mail-main' },
          h(
            'span',
            { className: 'ob-row1' },
            h('span', { className: 'ob-from' }, row.from),
            
          ),
          h('div', { className: 'ob-subject-line' }, h('span', { className: 'ob-subj' }, row.subject), h('span', { className: 'ob-time' }, time)),
          h('div', { className: 'ob-prev' }, row.preview || ''),
        ),
        row.hasAttachments ? h('span', { className: 'ob-clip', title: t('mail.attachment') }, h(MailIcon, { name: 'attachment', size: 14 })) : null,
      )
    }

    function avatarColor(name) {
      let hash = 0
      const text = String(name || '?')
      for (let index = 0; index < text.length; index += 1) hash = (hash * 31 + text.charCodeAt(index)) % 360
      return `hsl(${hash} 65% 82%)`
    }

    function dayGroupLabel(date, t) {
      const today = startOfDay(new Date())
      const target = startOfDay(date)
      const diff = Math.round((today - target) / 86_400_000)
      if (diff <= 0) return t('mail.today')
      if (diff === 1) return t('mail.yesterday')
      if (diff < 7) return t('mail.week')
      if (target.getMonth() === today.getMonth() && target.getFullYear() === today.getFullYear()) return t('mail.month')
      return t('mail.earlier')
    }

    function MailBlock(props) {
      const { rows, filter, setFilter, category, setCategory, focus, setFocus,
        selected, onSelect, t, folder, surface } = props
      const [filtersOpen, setFiltersOpen] = React.useState(false)
      const [ascending, setAscending] = React.useState(false)
      const [collapsed, setCollapsed] = React.useState({})
      const listDate = (row) => parseDate(folder === 'deleteditems' && row.modifiedAt ? row.modifiedAt : row.receivedAt)
      const visible = rows
        .filter((row) => filter !== 'unread' || !row.isRead)
        .filter((row) => category === 'all' || (row.category ?? 'other') === category)
        .filter((row) => focus === 'all' || (focus === 'focused' ? row.focused !== false : row.focused === false))
        .sort((a, b) => (ascending ? 1 : -1) * ((listDate(a)?.getTime() || 0) - (listDate(b)?.getTime() || 0)))
      const groups = []
      for (const row of visible) {
        const label = dayGroupLabel(listDate(row) ?? new Date(), t)
        let group = groups[groups.length - 1]
        if (!group || group.label !== label) groups.push(group = { label, items: [] })
        group.items.push(row)
      }
      return h('section', { className: 'ob-card ob-mail-list', style: surface, 'aria-label': t('mail.title') },
        h('div', { className: 'ob-tabs2' },
          ['focused', 'other', 'all'].map((key) => h('button', {
            key, type: 'button', className: `ob-tab2${focus === key ? ' on' : ''}`,
            onClick: () => setFocus(key), 'aria-pressed': focus === key,
          }, t(key === 'focused' ? 'mail.focused' : key === 'other' ? 'mail.other' : 'mail.all'))),
          h('span', { className: 'ob-spacer' }),
          h('button', { type: 'button', className: `ob-list-action${filtersOpen || filter !== 'all' || category !== 'all' ? ' active' : ''}`,
            title: '筛选邮件', 'aria-label': '筛选邮件', 'aria-expanded': filtersOpen, onClick: () => setFiltersOpen(!filtersOpen) }, h(MailIcon, { name: 'filter' })),
          h('button', { type: 'button', className: 'ob-list-action', title: ascending ? '最早邮件在前，点击切换' : '最新邮件在前，点击切换',
            'aria-label': ascending ? '按最新邮件排序' : '按最早邮件排序', onClick: () => setAscending(!ascending) }, h(MailIcon, { name: 'sort' }))),
        filtersOpen ? h('div', { className: 'ob-mail-filters' },
          h('label', null, '显示', h('select', { value: filter, onChange: (e) => setFilter(e.target.value) },
            h('option', { value: 'all' }, t('mail.all')), h('option', { value: 'unread' }, t('mail.unreadOnly')))),
          h('label', null, '分类', h('select', { value: category, onChange: (e) => setCategory(e.target.value) },
            h('option', { value: 'all' }, t('mail.all')), CATEGORY_META.map(([key]) => h('option', { key, value: key }, t(`cat.${key}`))))),
        ) : null,
        h('div', { className: 'ob-list' }, visible.length === 0
          ? h('div', { className: 'ob-empty' }, filter === 'unread' ? t('mail.emptyUnread') : '此视图中没有邮件')
          : groups.map((group) => h('div', { key: group.label },
              h('button', { type: 'button', className: `ob-mail-dayhead${collapsed[group.label] ? ' collapsed' : ''}`,
                'aria-expanded': !collapsed[group.label], onClick: () => setCollapsed({ ...collapsed, [group.label]: !collapsed[group.label] }) },
                h(MailIcon, { name: 'chevron', size: 14 }), group.label),
              collapsed[group.label] ? null : group.items.map((row) => h(MailRow, { key: row.id, row,
                active: selected?.id === row.id, t, time: fmtMailTime(listDate(row), t), onSelect, onContextRow: props.onContextRow }))))))
    }

    function MailWorkspace(props) {
      const { t, selected, busy, folders, folder, onFolder, account, onCompose, onDelete,
        onArchive, onToggleRead, onMarkFolderRead, readAllFeedback, connected, unread,
        toolbarSurface, folderSurface, readingSurface } = props
      const [favoritesOpen, setFavoritesOpen] = React.useState(true)
      const [accountOpen, setAccountOpen] = React.useState(true)
      const names = { inbox: '收件箱', drafts: '草稿', sentitems: '已发送邮件', deleteditems: '已删除邮件', archive: '存档', junkemail: '垃圾邮件' }
      const allFolders = folders.length ? folders : Object.entries(names).map(([key, name]) => ({ key, name, unread: 0 }))
      const folderButton = (entry) => h('button', { key: entry.key, type: 'button', className: `ob-folder${folder === entry.key ? ' on' : ''}`,
        'aria-current': folder === entry.key ? 'page' : undefined, onClick: () => onFolder(entry.key),
        onContextMenu: (event) => openContextMenu(event, [
          { label: '打开文件夹', icon: entry.key, run: () => onFolder(entry.key) },
          { label: '刷新文件夹', icon: 'refresh', disabled: busy, run: () => props.onRefreshFolder(entry.key) },
          { label: '新邮件', icon: 'compose', disabled: busy, run: onCompose },
        ]) },
        h(MailIcon, { name: entry.key }), h('span', { className: 'ob-folder-name' }, names[entry.key] || entry.name),
        (entry.key === 'inbox' ? unread : entry.unread) > 0 ? h('span', { className: 'ob-folder-count' }, entry.key === 'inbox' ? unread : entry.unread) : null)
      const onContextRow = (event, row) => openContextMenu(event, [
        { label: '打开邮件', icon: 'mail', run: () => props.onSelect(row) },
        { label: '回复', icon: 'reply', disabled: busy, run: () => props.onReplyRow(row) },
        { separator: true },
        { label: row.isRead ? '标记为未读' : '标记为已读', icon: 'read', disabled: busy, run: () => onToggleRead(row) },
        { label: '存档', icon: 'archive', disabled: busy, run: () => onArchive(row) },
        { label: '删除', icon: 'deleteditems', danger: true, disabled: busy, run: () => onDelete(row) },
        { separator: true },
        { label: '复制发件人邮箱', icon: 'people', disabled: !row.fromAddress, run: () => copyOutlookText(row.fromAddress) },
        row.webLink ? { label: '在 Outlook 中打开', icon: 'mail', run: () => window.open(row.webLink, '_blank', 'noopener,noreferrer') } : null,
      ])
      const command = (name, label, action, disabled = !selected || busy) => h('button', {
        type: 'button', className: 'ob-command', disabled, onClick: action,
      }, h(MailIcon, { name }), h('span', null, label))
      return h('div', { className: `ob-mail-workspace${selected ? ' has-selection' : ''}` },
        h('div', { className: 'ob-commandbar',
          style: toolbarSurface, role: 'toolbar', 'aria-label': '邮件操作' },
          h('button', { type: 'button', className: 'ob-command ob-new-mail', disabled: busy, onClick: onCompose }, h(MailIcon, { name: 'compose' }), '新邮件'),
          h('span', { className: 'ob-command-divider' }),
          command('deleteditems', t('mail.delete'), () => onDelete(selected)),
          command('archive', t('mail.archive'), () => onArchive(selected)),
          h('span', { className: 'ob-command-divider' }),
          command(selected?.isRead ? 'mail' : 'read', selected?.isRead ? '标记为未读' : '标记为已读', () => onToggleRead(selected)),
          command('read', '一键已读', onMarkFolderRead, busy || !connected)),
        readAllFeedback?.folder === folder ? h('div', {
          className: `ob-mail-command-notice${readAllFeedback.error ? ' is-error' : ''}`,
          role: readAllFeedback.error ? 'alert' : 'status',
        }, readAllFeedback.message) : null,
        h('div', { className: 'ob-grid ob-grid-mail' },
          h('aside', { className: 'ob-mail-folders',
            style: folderSurface, 'aria-label': '邮件文件夹' },
            h('button', { type: 'button', className: 'ob-folder-heading', 'aria-expanded': favoritesOpen, onClick: () => setFavoritesOpen(!favoritesOpen) },
              h('span', { className: favoritesOpen ? '' : 'ob-rotate' }, h(MailIcon, { name: 'chevron', size: 14 })), '收藏夹'),
            favoritesOpen ? ['inbox', 'sentitems', 'drafts'].map((key) => allFolders.find((f) => f.key === key)).filter(Boolean).map(folderButton) : null,
            h('button', { type: 'button', className: 'ob-folder-heading ob-account-heading', title: account, 'aria-expanded': accountOpen, onClick: () => setAccountOpen(!accountOpen) },
              h('span', { className: accountOpen ? '' : 'ob-rotate' }, h(MailIcon, { name: 'chevron', size: 14 })), h('span', { className: 'ob-folder-name' }, account || 'Outlook')),
            accountOpen ? ['inbox', 'junkemail', 'drafts', 'sentitems', 'deleteditems', 'archive'].map((key) => allFolders.find((f) => f.key === key)).filter(Boolean).map(folderButton) : null),
          h(MailBlock, { ...props, key: folder, onContextRow }),
          h('section', { className: 'ob-reading-pane',
            style: readingSurface, 'aria-label': selected ? '邮件阅读区' : '邮件欢迎页',
            onContextMenu: (event) => { if (selected && !event.target.closest('input,textarea,[contenteditable]') && !window.getSelection()?.toString()) onContextRow(event, selected) } },
            selected ? h(MailReadingPane, { key: selected.id, selected, t,
              replyRequest: props.replyRequest, onSend: props.onSendReply, busy, onClose: () => props.onSelect(null) }) : h(MailWelcome, null))))
    }

    /* ------------------------------------------------------- right column */

    const welcomeCityKey = 'outlook-bridge.weather-city'
    let welcomeLocation = null
    let welcomeWeatherCache = null

    function savedWelcomeCity() {
      if (welcomeLocation) return welcomeLocation
      try {
        const city = JSON.parse(window.localStorage.getItem(welcomeCityKey))
        if (typeof city?.name === 'string' && Number.isFinite(city.latitude) && Number.isFinite(city.longitude)) return city
      } catch {
        // Weather still works for this session if local storage is unavailable.
      }
      return null
    }

    function welcomeWeatherLabel(code) {
      if (code === 0) return '晴'
      if (code === 1 || code === 2) return '多云'
      if (code === 3) return '阴'
      if ([45, 48].includes(code)) return '雾'
      if ([51, 53, 55, 56, 57].includes(code)) return '毛毛雨'
      if ([61, 63, 65, 66, 67].includes(code)) return '雨'
      if ([71, 73, 75, 77, 85, 86].includes(code)) return '雪'
      if ([80, 81, 82].includes(code)) return '阵雨'
      if ([95, 96, 99].includes(code)) return '雷阵雨'
      return ''
    }

    async function welcomeWeatherRequest(kind, params, signal) {
      try {
        return await api(`/weather/${kind}?${params}`, { signal })
      } catch (error) {
        if (!/^unknown route \/weather\/(?:today|cities)$/i.test(error.message)) throw error
      }

      // The desktop UI may reload while an older plugin host is still running.
      const citySearch = kind === 'cities'
      const url = new URL(citySearch
        ? 'https://geocoding-api.open-meteo.com/v1/search'
        : 'https://api.open-meteo.com/v1/forecast')
      url.search = citySearch
        ? new URLSearchParams({ name: params.get('q'), count: '6', language: 'zh' })
        : new URLSearchParams({
          latitude: params.get('lat'),
          longitude: params.get('lon'),
          current: 'temperature_2m,weather_code',
          daily: 'temperature_2m_min,temperature_2m_max',
          timezone: 'auto',
          forecast_days: '1',
        })
      try {
        const response = await fetch(url, {
          credentials: 'omit',
          referrerPolicy: 'no-referrer',
          signal: AbortSignal.any([signal, AbortSignal.timeout(8_000)]),
        })
        if (!response.ok) throw new Error('Weather service returned an error')
        const data = await response.json()
        if (citySearch) {
          const cities = (data.results ?? [])
            .filter((row) => row.name && Number.isFinite(row.latitude) && Number.isFinite(row.longitude))
            .map((row) => ({
              id: row.id,
              name: [...new Set([row.name, row.admin1, row.country].filter(Boolean))].join(' · '),
              latitude: Number(row.latitude.toFixed(1)),
              longitude: Number(row.longitude.toFixed(1)),
            }))
          return { cities }
        }
        const temperature = data.current?.temperature_2m
        const low = data.daily?.temperature_2m_min?.[0]
        const high = data.daily?.temperature_2m_max?.[0]
        if (![temperature, low, high].every(Number.isFinite)) throw new Error('Missing temperatures')
        return { weather: { temperature, low, high, code: data.current?.weather_code ?? null } }
      } catch {
        throw new Error('天气暂时无法加载。请完全退出并重新打开 DeepSeek 助理后重试。')
      }
    }

    async function findWelcomeLocality({ latitude, longitude }) {
      const url = new URL('https://api.bigdatacloud.net/data/reverse-geocode-client')
      url.search = new URLSearchParams({
        latitude: latitude.toFixed(2),
        longitude: longitude.toFixed(2),
        localityLanguage: 'zh',
      })
      try {
        const response = await fetch(url, {
          credentials: 'omit',
          referrerPolicy: 'no-referrer',
          redirect: 'error',
          signal: AbortSignal.timeout(5_000),
        })
        if (!response.ok) return ''
        const result = await response.json()
        if (/^ip\b/i.test(result.lookupSource ?? '')) return ''
        const city = result.city?.trim() || result.principalSubdivision?.trim() || ''
        const locality = result.locality?.trim() || ''
        const district = locality.replace(/市$/, '') === city.replace(/市$/, '') ? '' : locality
        return [city, district].filter(Boolean).join(' · ')
      } catch {
        return ''
      }
    }

    function MailWelcome() {
      const [place, setPlace] = React.useState(savedWelcomeCity)
      const [report, setReport] = React.useState(null)
      const [weatherError, setWeatherError] = React.useState('')
      const [locationError, setLocationError] = React.useState('')
      const [locating, setLocating] = React.useState(false)
      const [searchOpen, setSearchOpen] = React.useState(false)
      const [query, setQuery] = React.useState('')
      const [cities, setCities] = React.useState([])
      const [searching, setSearching] = React.useState(false)
      const alive = React.useRef(false)
      const locatingRef = React.useRef(false)
      const searchRequest = React.useRef(null)
      const placeKey = place ? `${place.latitude},${place.longitude}` : ''

      const choosePlace = React.useCallback((next, remember) => {
        welcomeLocation = next
        setPlace(next)
        setSearchOpen(false)
        setCities([])
        setLocationError('')
        try {
          if (remember) window.localStorage.setItem(welcomeCityKey, JSON.stringify(next))
          else window.localStorage.removeItem(welcomeCityKey)
        } catch {
          // The choice is already active in memory.
        }
      }, [])

      const locate = React.useCallback((automatic = false) => {
        if (locatingRef.current || (automatic && welcomeLocation)) return
        if (!navigator.geolocation) {
          if (!automatic) setLocationError('当前窗口无法定位，可以手动选择城市。')
          return
        }
        locatingRef.current = true
        setLocating(true)
        setLocationError('')
        navigator.geolocation.getCurrentPosition(
          async ({ coords }) => {
            if (!alive.current) return
            const locality = await findWelcomeLocality(coords)
            locatingRef.current = false
            if (!alive.current) return
            setLocating(false)
            if (automatic && welcomeLocation) return
            choosePlace({
              name: locality ? `当前位置 · ${locality}` : '当前位置',
              latitude: Number(coords.latitude.toFixed(1)),
              longitude: Number(coords.longitude.toFixed(1)),
            }, false)
          },
          () => {
            locatingRef.current = false
            if (!alive.current) return
            setLocating(false)
            if (!automatic) setLocationError('未能获取当前位置，可以手动选择城市。')
          },
          { enableHighAccuracy: false, timeout: 8000, maximumAge: 600000 },
        )
      }, [choosePlace])

      React.useEffect(() => {
        alive.current = true
        let disposed = false
        if (!place && navigator.permissions?.query) {
          navigator.permissions.query({ name: 'geolocation' }).then((permission) => {
            if (!disposed && permission.state === 'granted') locate(true)
          }).catch(() => {})
        }
        return () => {
          disposed = true
          alive.current = false
          searchRequest.current?.abort()
        }
      }, [locate, place])

      React.useEffect(() => {
        if (!place) return
        const request = new AbortController()
        const refreshMs = 20 * 60_000
        async function refresh() {
          setWeatherError('')
          try {
            let result = welcomeWeatherCache
            if (result?.key !== placeKey || Date.now() - result.time >= refreshMs) {
              const params = new URLSearchParams({ lat: place.latitude, lon: place.longitude })
              const { weather } = await welcomeWeatherRequest('today', params, request.signal)
              result = { key: placeKey, time: Date.now(), weather }
            }
            if (!request.signal.aborted) {
              welcomeWeatherCache = result
              setReport(result)
            }
          } catch (error) {
            if (!request.signal.aborted) {
              setReport(null)
              setWeatherError(error.message)
            }
          }
        }
        refresh()
        const interval = window.setInterval(refresh, refreshMs)
        return () => {
          request.abort()
          window.clearInterval(interval)
        }
      }, [place, placeKey])

      const searchCities = async (event) => {
        event.preventDefault()
        if (query.trim().length < 2 || searchRequest.current) return
        const request = new AbortController()
        searchRequest.current = request
        setSearching(true)
        setCities([])
        setLocationError('')
        try {
          const result = await welcomeWeatherRequest('cities', new URLSearchParams({ q: query.trim() }), request.signal)
          if (request.signal.aborted) return
          setCities(result.cities)
          if (!result.cities.length) setLocationError('未找到这个城市，可以试试其他名称。')
        } catch (error) {
          if (!request.signal.aborted) setLocationError(error.message)
        } finally {
          if (searchRequest.current === request) {
            searchRequest.current = null
            if (alive.current) setSearching(false)
          }
        }
      }

      const weather = report?.key === placeKey ? report.weather : null
      const hour = new Date().getHours()
      const greeting = hour < 5 || hour >= 18 ? '晚上好' : hour < 12 ? '上午好' : '下午好'
      return h('div', { className: 'ob-welcome' },
        h('div', { className: 'ob-welcome-content' },
          h('div', { className: 'ob-welcome-weather', 'aria-live': 'polite' },
            h('div', null, place ? `今日天气 · ${place.name}` : '今日天气'),
            weather ? h(React.Fragment, null,
              h('strong', { className: 'ob-welcome-temperature' },
                [welcomeWeatherLabel(weather.code), `${Math.round(weather.temperature)} °C`].filter(Boolean).join(' · ')),
              h('div', null, `最低 ${Math.round(weather.low)}° · 最高 ${Math.round(weather.high)}°`))
              : place ? h('div', null, weatherError || '正在读取当地天气…') : null),
          h('div', { className: 'ob-welcome-options' },
            h('button', { type: 'button', disabled: locating || searching, onClick: () => locate() }, locating ? '正在定位…' : '使用当前位置'),
            h('button', { type: 'button', 'aria-expanded': searchOpen, onClick: () => {
              if (searchOpen) searchRequest.current?.abort()
              setSearchOpen(!searchOpen)
              setLocationError('')
            } }, searchOpen ? '收起城市搜索' : place ? '更换城市' : '选择城市'),
            weatherError && place ? h('button', { type: 'button', onClick: () => setPlace({ ...place }) }, '重试天气') : null,
            weather ? h('a', { href: 'https://open-meteo.com/', target: '_blank', rel: 'noopener noreferrer' }, 'Open-Meteo') : null),
          searchOpen ? h(React.Fragment, null,
            h('form', { className: 'ob-welcome-search', onSubmit: searchCities },
              h('input', { value: query, maxLength: 60, placeholder: '输入城市名', 'aria-label': '城市名称', onChange: (event) => setQuery(event.target.value) }),
              h('button', { className: 'ob-btn', type: 'submit', disabled: locating || searching || query.trim().length < 2 }, searching ? '搜索中…' : '搜索')),
            cities.length ? h('div', { className: 'ob-welcome-cities' },
              cities.map((city) => h('button', { key: city.id, type: 'button', className: 'ob-btn', disabled: locating, onClick: () => choosePlace(city, true) }, city.name))) : null) : null,
          locationError ? h('p', { className: 'ob-welcome-status', role: 'status' }, locationError) : null,
          h('h2', { className: 'ob-welcome-heading' }, `${greeting}，欢迎回来`),
          h('p', { className: 'ob-welcome-copy' }, '从左侧选一封邮件，开始今天的工作。')))
    }

    function MailReadingPane({ selected, t, replyRequest, onSend, busy, onClose }) {
      const [detail, setDetail] = React.useState(null)
      const [detailError, setDetailError] = React.useState('')
      const [detailLoading, setDetailLoading] = React.useState(true)
      const [opened, setOpened] = React.useState(false)
      const [draft, setDraft] = React.useState('')
      const [intent, setIntent] = React.useState('')
      const [summary, setSummary] = React.useState(null)
      const [translation, setTranslation] = React.useState(null)
      const [showTranslation, setShowTranslation] = React.useState(false)
      const [suggestion, setSuggestion] = React.useState(null)
      const [aiStatus, setAiStatus] = React.useState(null)
      const [aiMeta, setAiMeta] = React.useState(null)
      const [aiBusy, setAiBusy] = React.useState('')
      const [aiError, setAiError] = React.useState('')
      const [replyError, setReplyError] = React.useState('')
      const [sent, setSent] = React.useState(false)
      const request = React.useRef(null)
      const sending = React.useRef(false)
      const draftRef = React.useRef('')
      const editor = React.useRef(null)
      const mounted = React.useRef(true)
      const cacheKey = `outlook-reply:${selected.id}`
      const updateDraft = (value) => {
        draftRef.current = value
        setDraft(value)
        // Drafts remain in memory for this plugin session, never in shared browser storage.
        if (value) inlineReplyDrafts.set(cacheKey, value)
        else inlineReplyDrafts.delete(cacheKey)
      }
      React.useEffect(() => {
        mounted.current = true
        updateDraft(inlineReplyDrafts.get(cacheKey) || '')
        const controller = new AbortController()
        api(`/mail/detail?id=${encodeURIComponent(selected.id)}`, { signal: controller.signal })
          .then((data) => { if (mounted.current) { setDetail(data.mail); setDetailError('') } })
          .catch((failure) => { if (!controller.signal.aborted && mounted.current) setDetailError(failure.message) })
          .finally(() => { if (mounted.current) setDetailLoading(false) })
        api('/ai/status', { signal: controller.signal }).then((data) => { if (mounted.current) setAiStatus(data) })
          .catch(() => { if (mounted.current) setAiStatus({ available: false, message: 'AI 服务暂不可用，请重新打开邮件重试。' }) })
        return () => { mounted.current = false; controller.abort(); request.current?.abort() }
      }, [selected.id])
      React.useEffect(() => {
        if (replyRequest?.id === selected.id) setOpened(true)
      }, [replyRequest, selected.id])
      React.useEffect(() => {
        if (opened) editor.current?.focus({ preventScroll: true })
      }, [opened])
      const runAi = async (action) => {
        if (request.current) return
        const controller = new AbortController()
        request.current = controller
        setAiBusy(action)
        setAiError('')
        if (action === 'reply') { setOpened(true); setSent(false) }
        try {
          const result = await api('/mail/ai', { method: 'POST', signal: controller.signal, body: { id: selected.id, action, instruction: intent } })
          if (!mounted.current || controller.signal.aborted) return
          setAiMeta(result)
          if (action === 'translate') { setTranslation(result); setShowTranslation(true) }
          else if (action === 'read') setSummary(result)
          else if (!draftRef.current.trim()) updateDraft(result.text)
          else setSuggestion(result)
        } catch (failure) {
          if (mounted.current && !controller.signal.aborted) setAiError(failure.message)
        } finally {
          if (request.current === controller) request.current = null
          if (mounted.current) setAiBusy('')
        }
      }
      const send = async () => {
        if (sending.current || busy || !draftRef.current.trim()) return
        sending.current = true
        setReplyError('')
        const text = draftRef.current
        try {
          await onSend(selected, text)
          // A successful send clears only the exact draft that was submitted.
          if (inlineReplyDrafts.get(cacheKey) === text) inlineReplyDrafts.delete(cacheKey)
          if (mounted.current) { updateDraft(''); setOpened(false); setSuggestion(null); setSent(true) }
        } catch (failure) {
          if (mounted.current) setReplyError(failure.message)
        } finally { sending.current = false }
      }
      const mail = detail || selected
      const aiDisabled = Boolean(aiBusy) || !aiStatus?.available || busy
      return h('article', { className: 'ob-reading-card' },
        h('div', { className: 'ob-reading-heading' },
          h('h2', { className: 'ob-reading-title' }, showTranslation && translation ? translation.subject : selected.subject),
          h('button', { type:'button', className:'ob-btn', disabled:translation ? Boolean(aiBusy) : aiDisabled,
            title:showTranslation ? '切回原始邮件' : '使用已配置的 AI 模型翻译标题和正文为中文',
            onClick:()=>translation ? setShowTranslation(!showTranslation) : runAi('translate'),
          }, aiBusy === 'translate' ? '翻译中…' : showTranslation ? '查看原文' : translation ? '查看译文' : '翻译')),
        aiBusy === 'translate' ? h('div',{className:'ob-inline-hint',role:'status'},'正在翻译邮件…',
          h('button',{type:'button',className:'ob-command',onClick:()=>request.current?.abort()},'取消')) : null,
        showTranslation && translation?.truncated ? h('div',{className:'ob-inline-hint',role:'status'},'邮件较长，本次仅翻译正文前 28000 字；查看原文可阅读完整内容。') : null,
        h('div', { className: 'ob-reading-sender' },
          h('span', { className: 'ob-mail-avatar', style: { background: avatarColor(selected.from), color: '#48423d' } }, initial(selected.from)),
          h('div', { className: 'ob-reading-identity' }, h('strong', null, selected.from), h('span', null, selected.fromAddress)),
          h('button', { type: 'button', className: 'ob-command', onClick: onClose, 'aria-label': '关闭邮件' }, '×')),
        h('div', { className: 'ob-reading-date' }, parseDate(selected.receivedAt)?.toLocaleString() || ''),
        detailLoading ? h('div', { className: 'ob-inline-hint', role: 'status' }, '正在加载完整邮件…') : null,
        detailError ? h('div', { className: 'ob-inline-error', role: 'alert' }, `完整正文加载失败：${detailError}。下方为邮件预览。`) : null,
        h('div', { className: 'ob-reading-content' }, showTranslation && translation ? translation.text : mail.bodyText ?? mail.preview ?? ''),
        selected.webLink ? h('div', { className: 'ob-reading-links' }, h('a', { href: selected.webLink, target: '_blank', rel: 'noopener noreferrer' }, '在 Outlook 中查看完整邮件')) : null,
        h('div', { className: 'ob-reading-actions ob-mail-assistant-actions' },
          h('button', { type: 'button', className: 'ob-command', disabled: busy, onClick: () => { setOpened(true); setSent(false) } }, h(MailIcon, { name: 'reply' }), '回复'),
          h('button', { type: 'button', className: 'ob-command ob-ai-button', disabled: aiDisabled, onClick: () => runAi('read') }, h(MailIcon, { name: 'notes' }), aiBusy === 'read' ? '正在解读…' : 'AI 阅读邮件'),
          h('button', { type: 'button', className: 'ob-command ob-ai-button', disabled: aiDisabled, onClick: () => runAi('reply') }, h(MailIcon, { name: 'compose' }), aiBusy === 'reply' ? '正在生成…' : 'AI 快速回复'),
          aiBusy ? h('button', { type: 'button', className: 'ob-command', onClick: () => request.current?.abort() }, '取消生成') : null),
        h('div', { className: 'ob-inline-hint' }, aiStatus?.available
          ? `AI 使用 Harness 默认模型 ${aiStatus.model}；点击后会将当前邮件正文交给该模型处理。`
          : aiStatus?.message || '正在读取 AI 配置…'),
        aiError ? h('div', { className: 'ob-inline-error', role: 'alert' }, aiError) : null,
        aiMeta?.truncated && aiMeta.action === 'reply' ? h('div', { className: 'ob-inline-hint' }, '邮件较长，本次回复草稿仅依据正文前 28000 字生成。') : null,
        summary ? h('section', { className: 'ob-ai-summary', 'aria-label': 'AI 邮件解读' },
          h('div', { className: 'ob-inline-section-head' }, h('strong', null, 'AI 邮件解读'), h('button', { type: 'button', className: 'ob-command', onClick: () => setSummary(null), 'aria-label': '关闭解读' }, '×')),
          h('div', { className: 'ob-ai-output' }, summary.text),
          summary.truncated ? h('div', { className: 'ob-inline-hint' }, '邮件较长，本次只解读了正文前 28000 字。') : null) : null,
        sent ? h('div', { className: 'ob-inline-success', role: 'status' }, '回复已发送') : null,
        opened ? h('section', { className: 'ob-inline-reply', 'aria-label': '回复邮件' },
          h('div', { className: 'ob-inline-section-head' }, h('strong', null, `回复给 ${selected.from}`),
            h('button', { type: 'button', className: 'ob-command', disabled: busy, onClick: () => setOpened(false) }, '收起')),
          h('label', { className: 'ob-inline-intent' }, '回复要点', h('input', { value: intent, maxLength: 2000, disabled: Boolean(aiBusy), placeholder: '可选，例如：感谢对方，说明周五前给出结果', onChange: (event) => setIntent(event.target.value) })),
          h('textarea', { ref: editor, className: 'ob-inline-editor', 'aria-label': '回复正文', value: draft, disabled: busy,
            placeholder: '在这里输入回复，或点击 AI 快速回复生成草稿…', onChange: (event) => { updateDraft(event.target.value); setSent(false) } }),
          suggestion ? h('div', { className: 'ob-ai-suggestion' }, h('strong', null, 'AI 建议草稿'), h('div', { className: 'ob-ai-output' }, suggestion.text),
            h('button', { type: 'button', className: 'ob-command', disabled: busy, onClick: () => { updateDraft(suggestion.text); setSuggestion(null) } }, '采用草稿（替换当前回复）'),
            h('button', { type: 'button', className: 'ob-command', onClick: () => setSuggestion(null) }, '忽略')) : null,
          replyError ? h('div', { className: 'ob-inline-error', role: 'alert' }, replyError) : null,
          h('div', { className: 'ob-inline-reply-footer' },
            h('button', { type: 'button', className: 'ob-command ob-new-mail', disabled: busy || Boolean(aiBusy) || !draft.trim(), onClick: send }, busy ? '发送中…' : '发送回复'),
            h('button', { type: 'button', className: 'ob-command', disabled: aiDisabled, onClick: () => runAi('reply') }, aiBusy === 'reply' ? '生成中…' : 'AI 快速回复'),
            h('span', { className: 'ob-inline-hint' }, 'AI 草稿可编辑，点击发送后才会发出。'))) : null)
    }

    const CATEGORY_META = [
      ['security', '#d64545'],
      ['billing', '#e8833a'],
      ['work', '#3a8fd9'],
      ['social', '#22a06b'],
      ['promo', '#8b5cf6'],
      ['other', '#8a8f98'],
    ]

    const WEEK_LABELS = [1, 2, 3, 4, 5, 6, 0]

    function monthGridStart(cursor) {
      const first = new Date(cursor.getFullYear(), cursor.getMonth(), 1)
      const offset = (first.getDay() + 6) % 7
      const start = new Date(first)
      start.setDate(start.getDate() - offset)
      start.setHours(0, 0, 0, 0)
      return start
    }

    function MonthCalendar(props) {
      const { t, tasks = [], onToggle, busy, surface, listSurface, accountMail,
        toolbarSurface, pendingCreate, onPendingCreateConsumed, calendarNonce } = props
      const [cursor, setCursor] = React.useState(() => new Date(new Date().getFullYear(), new Date().getMonth(), 1))
      const [data, setData] = React.useState(null)
      const [loading, setLoading] = React.useState(false)
      const [error, setError] = React.useState(null)
      const [hidden, setHidden] = React.useState(() => new Set())
      const defaultCalendarKey = accountMail ? `outlook:default-create-calendar:${accountMail.toLowerCase()}` : ''
      const [preferredCalendarId, setPreferredCalendarId] = React.useState(() => {
        try { return defaultCalendarKey ? window.localStorage.getItem(defaultCalendarKey) || '' : '' }
        catch { return '' }
      })
      const [creating, setCreating] = React.useState(null)
      const editorOpen = React.useRef(false)
      const refreshAfterEditor = React.useRef(false)
      editorOpen.current = Boolean(creating)
      const [form, setForm] = React.useState({ subject: '', startTime: '09:00', endTime: '10:00', allDay: false, location: '' })
      const [saving, setSaving] = React.useState(false)
      const [notice, setNotice] = React.useState(null)
      const [selectedDay, setSelectedDay] = React.useState(() => startOfDay(new Date()))
      const [showSidebar, setShowSidebar] = React.useState(true)
      const [closedGroups, setClosedGroups] = React.useState({})
      const [eventPeek, setEventPeek] = React.useState(null)
      const [dayPeek, setDayPeek] = React.useState(null)
      const eventClickTimer = React.useRef(null)
      const requestVersion = React.useRef(0)
      const savingRequest = React.useRef(false)
      const [createError, setCreateError] = React.useState(null)
      const [dragState, setDragState] = React.useState(null)
      const dragMoved = React.useRef(false)

      const load = React.useCallback(async () => {
        // Keep the editor's calendars and cards stable throughout an editing
        // session. Drag saves already update their local row immediately.
        if (editorOpen.current) { refreshAfterEditor.current = true; return }
        refreshAfterEditor.current = false
        const version = ++requestVersion.current
        setLoading(true)
        try {
          const start = monthGridStart(cursor)
          const end = new Date(start)
          end.setDate(end.getDate() + 42)
          const path = `/calendar?from=${encodeURIComponent(start.toISOString())}&to=${encodeURIComponent(end.toISOString())}`
          const cacheKey = accountMail ? `${accountMail}:${path}` : ''
          const revision = calendarRangeRevision
          const cached = recentCalendarRange
          if (cacheKey && cached?.key === cacheKey && Date.now() - cached.at < 20_000) {
            setData(cached.data)
            setError(null)
          }
          const result = await api(path)
          if (version !== requestVersion.current) return
          if (editorOpen.current) { refreshAfterEditor.current = true; return }
          if (cacheKey && revision === calendarRangeRevision && !result.warnings?.length) {
            recentCalendarRange = { key: cacheKey, data: result, at: Date.now() }
          }
          setData(result)
          setError(null)
        } catch (failure) {
          if (version === requestVersion.current) {
            if (editorOpen.current) refreshAfterEditor.current = true
            else setError(failure.message)
          }
        } finally {
          if (version === requestVersion.current) setLoading(false)
        }
      }, [cursor, accountMail])

      React.useEffect(() => {
        void load()
        return () => { requestVersion.current += 1 }
      }, [load, calendarNonce])

      React.useEffect(() => {
        if (!creating && refreshAfterEditor.current) void load()
      }, [creating, load])

      const calendars = data?.calendars ?? []
      const calendarsById = new Map(calendars.map((calendar) => [calendar.id, calendar]))
      const preferredCalendar = calendarsById.get(preferredCalendarId)
      const createCalendar = preferredCalendar && preferredCalendar.canEdit !== false ? preferredCalendar :
        calendars.find((calendar) => calendar.isDefault && calendar.canEdit !== false) ||
        calendars.find((calendar) => calendar.canEdit !== false)
      const defaultCalendarAction = (calendar) => ({
        label: calendar.id === createCalendar?.id ? '当前默认创建日历' : '设为默认创建日程',
        icon: 'check',
        disabled: calendar.canEdit === false || calendar.id === createCalendar?.id,
        run: () => {
          if (defaultCalendarKey) window.localStorage.setItem(defaultCalendarKey, calendar.id)
          setPreferredCalendarId(calendar.id)
          setNotice(`新建日程默认使用「${calendar.name}」`)
        },
      })
      const eventColor = (event) => calendarsById.get(event.calendarId)?.color || event.color || '#0f6cbd'
      const events = (data?.events ?? []).filter((event) => !hidden.has(event.calendarId))
      const byDay = new Map()
      for (const event of events) {
        const start = parseDate(event.start)
        if (!start) continue
        const gridFrom = monthGridStart(cursor)
        const gridTo = new Date(gridFrom)
        gridTo.setDate(gridTo.getDate() + 42)
        const end = parseDate(event.end) ?? start
        const last = end > start ? new Date(end.getTime() - 1) : start
        const day = startOfDay(start < gridFrom ? gridFrom : start)
        for (; day <= last && day < gridTo; day.setDate(day.getDate() + 1)) {
          const key = dayKey(day)
          if (!byDay.has(key)) byDay.set(key, [])
          byDay.get(key).push(event)
        }
      }

      const gridStart = monthGridStart(cursor)
      const cells = []
      for (let index = 0; index < 42; index += 1) {
        const day = new Date(gridStart)
        day.setDate(day.getDate() + index)
        cells.push(day)
      }
      const today = startOfDay(new Date())
      const monthLabel = `${cursor.getFullYear()} 年 ${cursor.getMonth() + 1} 月`

      const toggleCalendar = (id) => {
        const next = new Set(hidden)
        if (next.has(id)) next.delete(id)
        else next.add(id)
        setHidden(next)
      }

      const moveMonth = (offset) => {
        const next = new Date(cursor.getFullYear(), cursor.getMonth() + offset, 1)
        setCursor(next)
        setSelectedDay(next)
      }
      const chooseDay = (day) => {
        setSelectedDay(day)
        if (day.getMonth() !== cursor.getMonth() || day.getFullYear() !== cursor.getFullYear()) {
          setCursor(new Date(day.getFullYear(), day.getMonth(), 1))
        }
      }
      const goToday = () => {
        setSelectedDay(today)
        setCursor(new Date(today.getFullYear(), today.getMonth(), 1))
      }
      const groupOf = (key, label, items) => items.length === 0 ? null : h('div', { className: 'ob-calendar-group', key },
        h('div', { className: 'ob-calendar-group-head' },
          h('button', { type: 'button', className: `ob-calendar-collapse${closedGroups[key] ? ' closed' : ''}`, 'aria-label': `${closedGroups[key] ? '展开' : '收起'}${label}`,
            'aria-expanded': !closedGroups[key], onClick: () => setClosedGroups({ ...closedGroups, [key]: !closedGroups[key] }) }, h(MailIcon, { name: 'chevron', size: 16 })),
          h('label', null, h('input', { type: 'checkbox', checked: items.every((item) => !hidden.has(item.id)),
            onChange: () => {
              const hideAll = items.every((item) => !hidden.has(item.id))
              setHidden((old) => { const next = new Set(old); items.forEach((item) => hideAll ? next.add(item.id) : next.delete(item.id)); return next })
            } }), label)),
        closedGroups[key] ? null : items.map((calendar) => h('label', { className: 'ob-calendar-choice', key: calendar.id, onContextMenu: (event) => openContextMenu(event, [
          { label: '在此日历新建活动', disabled: calendar.canEdit === false, run: () => openCreate(selectedDay, { calendarId: calendar.id }) },
          defaultCalendarAction(calendar),
          { label: hidden.has(calendar.id) ? '显示日历' : '隐藏日历', run: () => toggleCalendar(calendar.id) },
          { label: '只显示此日历', run: () => setHidden(new Set(calendars.filter((item) => item.id !== calendar.id).map((item) => item.id))) },
          { label: '显示全部日历', run: () => setHidden(new Set()) },
        ]) },
          h('input', { type: 'checkbox', checked: !hidden.has(calendar.id), style: { accentColor: calendar.color }, onChange: () => toggleCalendar(calendar.id) }),
          h('span', { title: calendar.id === createCalendar?.id ? `${calendar.name}（默认创建日历）` : calendar.name }, calendar.name))))

      const openCreate = (day, overrides = {}) => {
        const date = dayKey(day)
        const autoTime = !overrides.allDay && !overrides.startTime && !overrides.endTime
        const { hour, minute } = shanghaiTodayParts()
        const slot = autoTime
          ? eventSlotFields(calendarTimestamp(`${date}T${hour}:${minute}:00`) + EVENT_BUFFER_MS)
          : { date, endDate: date, startTime: '09:00', endTime: '09:30' }
        setForm({ subject: '', ...slot, allDay: false,
          location: '', attendees: '', body: '', showAs: 'busy', reminder: '15',
          calendarId: createCalendar?.id || '',
          eventId: '', listEventId: '',
          transactionId: window.crypto?.randomUUID?.() || `event-${Date.now()}-${Math.random().toString(16).slice(2)}`, ...overrides })
        setCreateError(null)
        setCreating({ autoTime })
      }

      React.useEffect(() => {
        if (!pendingCreate || (!data && !error)) return
        const day = parseDate(`${pendingCreate.date || dayKey(new Date())}T12:00:00`) || new Date()
        openCreate(day, {
          subject: pendingCreate.subject || '',
          date: pendingCreate.date || dayKey(day),
          endDate: pendingCreate.endDate || pendingCreate.date || dayKey(day),
          startTime: pendingCreate.startTime || '09:00',
          endTime: pendingCreate.endTime || '09:30',
          allDay: Boolean(pendingCreate.allDay),
          location: pendingCreate.location || '',
          body: pendingCreate.body || '',
        })
        onPendingCreateConsumed?.()
      }, [pendingCreate, data, error])


      const dateContext = (event, day) => openContextMenu(event, [
        { label: '新建活动', run: () => openCreate(day) },
        { label: '新建全天活动', run: () => openCreate(day, { allDay: true }) },
        { label: '查看当天日程', icon: 'notes', run: () => openDayPeek(day) },
        { separator: true },
        { label: '回到今天', run: goToday },
        { label: '刷新日历', icon: 'refresh', disabled: loading, run: load },
      ])
      const eventContext = (event, item, day) => openContextMenu(event, [
        { label: '查看日程详情', icon: 'notes', run: () => openEventPeek(item, day) },
        { label: '在当天新建活动', disabled: calendars.find((calendar) => calendar.id === item.calendarId)?.canEdit === false, run: () => openCreate(day, { calendarId: item.calendarId }) },
        { label: '编辑日程', icon: 'edit', disabled: calendars.find((calendar) => calendar.id === item.calendarId)?.canEdit === false, run: () => openEdit(item) },
        { label: '复制日程标题', icon: 'notes', run: () => copyOutlookText(item.subject) },
        item.webLink ? { label: '在 Outlook 中打开', run: () => window.open(item.webLink, '_blank', 'noopener,noreferrer') } : null,
        { separator: true },
        { label: '删除日程', icon: 'trash', danger: true, disabled: saving || savingRequest.current || !eventCanEdit(item), run: () => deletePeekEvent(item, true) },
      ])
      const openEdit = async (item) => {
        const realId = String(item.id).includes(':') ? String(item.id).slice(String(item.id).indexOf(':') + 1) : String(item.id)
        const day = parseDate(item.start) || new Date()
        setCreateError(null)
        try {
          const params = new URLSearchParams({ id: realId, calendarId: item.calendarId || '' })
          const detail = await api(`/calendar/event?${params}`)
          const startDate = parseDate(detail.event.start) || day
          const endDate = parseDate(detail.event.end) || startDate
          const sourceCalendarId = detail.event.calendarId || item.calendarId || ''
          const sourceCalendar = calendarsById.get(sourceCalendarId) || {
            id: sourceCalendarId, name: item.calendar, color: item.color, canEdit: false,
          }
          setForm({
            subject: detail.event.subject || '',
            date: dayKey(startDate),
            endDate: detail.event.allDay && endDate > startDate ? addDaysYmd(dayKey(endDate), -1) : dayKey(endDate),
            startTime: fmtTime(startDate),
            endTime: fmtTime(endDate),
            allDay: Boolean(detail.event.allDay),
            location: detail.event.location || '',
            attendees: detail.event.attendees || '',
            body: htmlToPlainText(detail.event.body || ''),
            showAs: detail.event.showAs || 'busy',
            reminder: String(detail.event.reminder ?? '15'),
            calendarId: sourceCalendarId,
            sourceCalendarId,
            sourceCalendar,
            eventId: detail.event.id || realId,
            listEventId: item.id,
            transactionId: window.crypto?.randomUUID?.() || `event-${Date.now()}`,
          })
          setCreating({ autoTime: false })
        } catch (failure) {
          setCreateError(failure.message)
        }
      }

      const removeEvent = async () => {
        if (!form.eventId || savingRequest.current) return
        if (!window.confirm('确定删除这条日程吗？此操作不可撤销。')) return
        savingRequest.current = true
        setSaving(true)
        try {
          await api('/calendar/delete', { method: 'POST', body: { id: form.eventId } })
          setCreating(null)
          setNotice('日程已删除')
          await load()
        } catch (failure) {
          setCreateError(failure.message)
        } finally {
          savingRequest.current = false
          setSaving(false)
        }
      }

      const submitCreate = async ({ saveCurrent = true } = {}) => {
        if (!creating || savingRequest.current) return
        savingRequest.current = true
        setSaving(true)
        setCreateError(null)
        try {
          const moving = Boolean(saveCurrent && form.eventId && form.sourceCalendarId !== form.calendarId)
          if (moving) {
            let capabilities
            try { capabilities = await api('/calendar/capabilities') } catch {}
            if (capabilities?.verifiedTransfer !== 1) {
              throw new Error('日历后台版本尚未更新，请完全退出并重新打开 DeepSeek 后再转移。此次没有提交保存。')
            }
          }
          // Timeline adjustments are local until this explicit Save. Remove each
          // confirmed update so retrying a partial failure does not replay it.
          for (const row of Object.values(form.timelineChanges || {})) {
            await api('/calendar/update', { method: 'POST', body: {
              eventId: String(row.id).slice(String(row.id).indexOf(':') + 1),
              subject: row.subject || '(无主题)', calendarId: row.calendarId,
              date: row.start.slice(0, 10), startTime: row.start.slice(11, 16),
              endDate: row.end.slice(0, 10), endTime: row.end.slice(11, 16),
              datesOnly: true, allDay: false,
            } })
            refreshAfterEditor.current = true
            setForm((old) => { const changes = { ...old.timelineChanges }; delete changes[row.id]; return { ...old, timelineChanges: changes, timelineSaved: { ...old.timelineSaved, [row.id]: row } } })
          }
          const result = saveCurrent
            ? await api(form.eventId ? '/calendar/update' : '/calendar/create', { method: 'POST', body: form })
            : { updated: true }
          if (moving && (!result.moved || result.calendarId !== form.calendarId)) {
            throw new Error('未收到目标日历的转移确认，请刷新日历核对；本次不能确认转移成功。')
          }
          setCreating(null)
          setNotice(result.moved ? '日程已移到所选日历' : !saveCurrent || form.eventId ? '日程已更新' : t('cal.created'))
          const savedDay = new Date(`${form.date}T12:00:00`)
          setSelectedDay(savedDay)
          if (savedDay.getMonth() !== cursor.getMonth() || savedDay.getFullYear() !== cursor.getFullYear()) {
            setCursor(new Date(savedDay.getFullYear(), savedDay.getMonth(), 1))
          } else await load()
        } catch (failure) {
          setCreateError(failure.message)
        } finally {
          savingRequest.current = false
          setSaving(false)
        }
      }


      const eventCanEdit = (item) => calendars.find((calendar) => calendar.id === item.calendarId)?.canEdit !== false
      const realEventId = (item) => {
        const id = String(item.id)
        return id.includes(':') ? id.slice(id.indexOf(':') + 1) : id
      }
      const placePeek = (rect) => {
        const width = 320
        const left = Math.max(8, Math.min((rect?.right || 40) + 8, window.innerWidth - width - 8))
        const top = Math.max(8, Math.min(rect?.top || 40, window.innerHeight - 220))
        return { left, top }
      }
      const openEventPeek = (item, day, anchorEl) => {
        const rect = anchorEl?.getBoundingClientRect?.() || null
        setDayPeek(null)
        setEventPeek({ item, day, ...placePeek(rect), expanded: false })
      }
      const openDayPeek = (day, anchorEl) => {
        const rect = anchorEl?.getBoundingClientRect?.() || null
        setEventPeek(null)
        setDayPeek({ day, ...placePeek(rect) })
      }
      const scheduleEventClick = (item, day, anchorEl, delay = 220) => {
        if (eventClickTimer.current) clearTimeout(eventClickTimer.current)
        eventClickTimer.current = setTimeout(() => {
          eventClickTimer.current = null
          openEventPeek(item, day, anchorEl)
        }, delay)
      }
      const closeDayPeek = () => {
        if (eventClickTimer.current) clearTimeout(eventClickTimer.current)
        eventClickTimer.current = null
        setDayPeek(null)
      }
      const handleEventDoubleClick = (item) => {
        if (eventClickTimer.current) {
          clearTimeout(eventClickTimer.current)
          eventClickTimer.current = null
        }
        setEventPeek(null)
        setDayPeek(null)
        openEdit(item)
      }
      const deletePeekEvent = async (item, fromMenu = false) => {
        if (!item || savingRequest.current) return
        if (!eventCanEdit(item)) {
          setNotice('此日历为只读，无法删除日程')
          return
        }
        if (!window.confirm('确定删除这条日程吗？此操作不可撤销。')) return
        savingRequest.current = true
        try {
          await api('/calendar/delete', { method: 'POST', body: { id: realEventId(item) } })
          setData((previous) => previous ? { ...previous, events: (previous.events ?? []).filter((event) => event.id !== item.id) } : previous)
          setEventPeek(null)
          setNotice('日程已删除')
          await load()
          return true
        } catch (failure) {
          if (fromMenu) throw new Error(`删除失败：${failure.message}`)
          setNotice(`删除失败：${failure.message}`)
        } finally {
          savingRequest.current = false
        }
      }
      const daysBetweenKeys = (fromKey, toKey) => {
        const from = new Date(`${fromKey}T12:00:00`)
        const to = new Date(`${toKey}T12:00:00`)
        return Math.round((to.getTime() - from.getTime()) / 86_400_000)
      }
      const shiftStamp = (stamp, deltaDays, allDay) => {
        const parsed = parseDate(stamp)
        if (!parsed) return stamp
        const next = new Date(parsed)
        next.setDate(next.getDate() + deltaDays)
        return allDay ? `${dayKey(next)}T00:00:00` : `${dayKey(next)}T${fmtTime(next)}:00`
      }
      const moveEventToDay = async (event, fromKey, toKey) => {
        if (!event || !fromKey || !toKey || fromKey === toKey) return
        if (!eventCanEdit(event)) {
          setNotice('此日历为只读，无法移动日程')
          return
        }
        const delta = daysBetweenKeys(fromKey, toKey)
        if (!delta || savingRequest.current) return
        const snapshot = data
        const optimisticStart = shiftStamp(event.start, delta, event.allDay)
        const optimisticEnd = shiftStamp(event.end, delta, event.allDay)
        setData((prev) =>
          prev
            ? {
                ...prev,
                events: (prev.events || []).map((row) =>
                  row.id === event.id ? { ...row, start: optimisticStart, end: optimisticEnd } : row,
                ),
              }
            : prev,
        )
        setSelectedDay(new Date(`${toKey}T12:00:00`))
        setNotice('正在移动日程…')
        savingRequest.current = true
        try {
          const realId = realEventId(event)
          const detailStart = parseDate(event.start)
          const detailEnd = parseDate(event.end)
          if (!detailStart || !detailEnd) throw new Error('无法解析日程时间')
          const movedStart = new Date(detailStart)
          movedStart.setDate(movedStart.getDate() + delta)
          const movedEnd = new Date(detailEnd)
          movedEnd.setDate(movedEnd.getDate() + delta)
          const allDay = Boolean(event.allDay)
          const endDate = allDay ? dayKey(new Date(movedEnd.getTime() - 1)) : dayKey(movedEnd)
          // datesOnly: server PATCHes only start/end/isAllDay; calendarId is for canEdit check, never Graph body
          await api('/calendar/update', {
            method: 'POST',
            body: {
              eventId: realId,
              subject: event.subject || '(无主题)',
              date: dayKey(movedStart),
              endDate,
              startTime: allDay ? '00:00' : fmtTime(movedStart),
              endTime: allDay ? '23:59' : fmtTime(movedEnd),
              allDay,
              datesOnly: true,
              calendarId: event.calendarId || '',
            },
          })
          setNotice('日程已移动')
          await load()
        } catch (failure) {
          setData(snapshot)
          setNotice(`移动失败：${failure.message}`)
        } finally {
          savingRequest.current = false
        }
      }

      return h(
        'div', { className: `ob-calendar-workspace${showSidebar ? '' : ' sidebar-hidden'}` },
        h('div', { className: 'ob-commandbar ob-calendar-toolbar',
          style: toolbarSurface, role: 'toolbar', 'aria-label': '日历操作' },
          h('button', { type: 'button', className: 'ob-command ob-new-mail', disabled: busy || saving, onClick: () => openCreate(selectedDay) }, h(MailIcon, { name: 'calendar' }), '新事件'),
          h('span', { className: 'ob-command-divider' }),
          h('span', { className: 'ob-calendar-view-label' }, h(MailIcon, { name: 'calendar', size: 17 }), '月'),
          h('button', { type: 'button', className: `ob-command${showSidebar ? ' active' : ''}`, 'aria-pressed': showSidebar, onClick: () => setShowSidebar(!showSidebar) }, h(MailIcon, { name: 'filter' }), '日历筛选'),
          h('span', { className: 'ob-command-divider' }),
          h('button', { type: 'button', className: 'ob-command', disabled: loading, onClick: load }, h(MailIcon, { name: 'refresh', size: 17 }), t('action.refresh')),
          h('span', { className: 'ob-spacer' }),
          h('span', { className: 'ob-calendar-hint' }, '双击日期新建日程')),
        h('div', { className: 'ob-calendar-layout' },
          showSidebar ? h('aside', { className: 'ob-calendar-sidebar',
            style: listSurface, 'aria-label': '日历导航' },
            h('div', { className: 'ob-mini-head' }, h('strong', null, monthLabel),
              h('button', { type: 'button', 'aria-label': '上个月', onClick: () => moveMonth(-1) }, '‹'),
              h('button', { type: 'button', 'aria-label': '下个月', onClick: () => moveMonth(1) }, '›')),
            h('div', { className: 'ob-mini-week' }, ['一', '二', '三', '四', '五', '六', '日'].map((label) => h('span', { key: label }, label))),
            h('div', { className: 'ob-mini-grid' }, cells.map((day) => {
              const key = dayKey(day)
              return h('button', { key, type: 'button', className: `ob-mini-day${day.getMonth() !== cursor.getMonth() ? ' outside' : ''}${key === dayKey(today) ? ' today' : ''}${key === dayKey(selectedDay) ? ' selected' : ''}`,
                'aria-label': `${day.getFullYear()}年${day.getMonth()+1}月${day.getDate()}日`, 'aria-pressed': key === dayKey(selectedDay), onClick: () => chooseDay(day), onContextMenu: (event) => dateContext(event, day) }, h('span', null, day.getDate()))
            })),
            h('div', { className: 'ob-calendar-sidebar-divider' }),
            groupOf('mine', '我的日历', calendars.filter((item) => !item.isDefault)),
            groupOf('other', '其他日历', calendars.filter((item) => item.isDefault)),
            calendars.length === 0 ? h('div', { className: 'ob-calendar-sidebar-empty' }, loading ? t('cal.loading') : error ? '日历加载失败，请点击刷新' : '暂无日历') : null,
            calendars.length ? h('button', { type: 'button', className: 'ob-calendar-show-all', onClick: () => setHidden(new Set()) }, '全部显示') : null,
            tasks.length ? h('details', { className: 'ob-calendar-tasks' }, h('summary', null, `待办 · ${tasks.length}`), tasks.slice(0,8).map((task) => h('label', { className: 'ob-calendar-task', key: `${task.listId}:${task.id}` },
              h('input', { type: 'checkbox', checked: Boolean(task.completed), disabled: busy || task.completed, onChange: () => onToggle(task) }), h('span', null, task.title)))) : null) : null,
          h('section', { className: 'ob-calendar-board', style: surface, 'aria-label': monthLabel },
            h('div', { className: 'ob-calendar-month-head' },
              h('button', { type: 'button', className: 'ob-command ob-calendar-today', onClick: goToday }, h(MailIcon, { name: 'calendar', size: 16 }), '今天'),
              h('button', { type: 'button', className: 'ob-calendar-nav', 'aria-label': '上个月', onClick: () => moveMonth(-1) }, '‹'),
              h('button', { type: 'button', className: 'ob-calendar-nav', 'aria-label': '下个月', onClick: () => moveMonth(1) }, '›'),
              h('label', { className: 'ob-calendar-month-picker', title: '跳转到月份' },
                h('span', null, monthLabel), h(MailIcon, { name: 'chevron', size: 14 }),
                h('input', { type: 'month', 'aria-label': '跳转到月份', value: `${cursor.getFullYear()}-${pad(cursor.getMonth()+1)}`, onChange: (event) => {
                  const match = /^(\d{4})-(\d{2})$/.exec(event.target.value)
                  if (match) { const next = new Date(Number(match[1]), Number(match[2])-1, 1); setCursor(next); setSelectedDay(next) }
                } })),
              h('span', { className: 'ob-spacer' }),
              loading ? h('span', { className: 'ob-calendar-loading', role: 'status' }, t('cal.loading')) : null),
            error ? h('div', { className: 'ob-calendar-status error', role: 'alert' }, error) : null,
            (data?.warnings ?? []).length ? h('div', { className: 'ob-calendar-status', role: 'status' }, data.warnings.join('；')) : null,
            h('div', { className: 'ob-calendar-scroll' },
            h('div', { className: 'ob-calendar-weekdays' }, ['星期一','星期二','星期三','星期四','星期五','星期六','星期日'].map((label) => h('span', { key: label }, label))),
            h('div', { className: 'ob-calendar-month-grid' }, cells.map((day, index) => {
              const key = dayKey(day)
              const dayEvents = byDay.get(key) ?? []
              const isToday = key === dayKey(today)
              const selected = key === dayKey(selectedDay)
              const dateLabel = day.getDate() === 1 || index === 0 ? `${day.getMonth()+1} 月 ${day.getDate()} 日` : String(day.getDate())
              return h('div', { key, className: `ob-calendar-day${day.getMonth() !== cursor.getMonth() ? ' outside' : ''}${isToday ? ' today' : ''}${selected ? ' selected' : ''}${dragState && dragState.previewKey === key && dragState.fromKey !== key ? ' drag-over' : ''}`,
                  onClick: () => setSelectedDay(day), onDoubleClick: () => openCreate(day), onContextMenu: (event) => dateContext(event, day),
                  onDragOver: (mouse) => {
                    if (!dragState) return
                    mouse.preventDefault()
                    mouse.dataTransfer.dropEffect = 'move'
                    if (dragState.previewKey !== key) {
                      dragMoved.current = true
                      setDragState({ ...dragState, previewKey: key })
                    }
                  },
                  onDrop: (mouse) => {
                    mouse.preventDefault()
                    mouse.stopPropagation()
                    const payload = dragState
                    setDragState(null)
                    if (payload) moveEventToDay(payload.event, payload.fromKey, key)
                  } },
                h('button', { type: 'button', className: 'ob-calendar-date', 'aria-label': `${day.getFullYear()}年${day.getMonth()+1}月${day.getDate()}日，新建日程`,
                  onClick: (event) => { event.stopPropagation(); setSelectedDay(day); if (event.detail === 0) openCreate(day) } },
                  h('span', { className: isToday ? 'today-dot' : '' }, isToday ? day.getDate() : dateLabel)),
                h('div', { className: 'ob-calendar-events' }, [
                  ...dayEvents.slice(0,2).map((event) => {
                    const editable = eventCanEdit(event)
                    const dragging = dragState && dragState.event && dragState.event.id === event.id
                    return h('button', {
                      type: 'button', key: event.id, className: `ob-calendar-event${dragging ? ' dragging' : ''}`,
                      draggable: editable,
                      onDragStart: (mouse) => {
                        if (!editable) { mouse.preventDefault(); return }
                        dragMoved.current = false
                        mouse.stopPropagation()
                        mouse.dataTransfer.effectAllowed = 'move'
                        mouse.dataTransfer.setData('text/plain', String(event.id))
                        try { mouse.dataTransfer.setDragImage(mouse.currentTarget, 12, 10) } catch (_) {}
                        setDragState({ event, fromKey: key, previewKey: key })
                      },
                      onDragEnd: () => { setDragState(null) },
                      onContextMenu: (mouse) => eventContext(mouse, event, day),
                      style: { '--ob-event-color': eventColor(event) },
                      title: [event.subject, calendarsById.get(event.calendarId)?.name || event.calendar, event.location].filter(Boolean).join(' · ') +
                        (editable ? '（拖到其他日期可移动）' : ''),
                      onClick: (e) => {
                        e.stopPropagation()
                        if (dragMoved.current) { dragMoved.current = false; return }
                        setSelectedDay(day)
                        scheduleEventClick(event, day, e.currentTarget)
                      },
                      onDoubleClick: (e) => {
                        e.stopPropagation()
                        e.preventDefault()
                        setSelectedDay(day)
                        handleEventDoubleClick(event)
                      },
                    }, event.allDay ? event.subject : `${fmtTime(parseDate(event.start))} ${event.subject}`)
                  }),
                  dragState && dragState.previewKey === key && dragState.fromKey !== key
                    ? h('button', {
                        type: 'button', key: `preview:${dragState.event.id}`, className: 'ob-calendar-event drag-preview',
                        tabIndex: -1,
                        style: { '--ob-event-color': eventColor(dragState.event) },
                      }, dragState.event.allDay ? dragState.event.subject : `${fmtTime(parseDate(dragState.event.start))} ${dragState.event.subject}`)
                    : null,
                  dayEvents.length > 2 ? h('button', { type: 'button', className: 'ob-calendar-more', onDoubleClick: (e) => e.stopPropagation(),
                    onClick: (e) => { e.stopPropagation(); setSelectedDay(day); openDayPeek(day, e.currentTarget) } }, `还有 ${dayEvents.length-2} 个日程`) : null,
                ].filter(Boolean)))
            }))),
            notice ? h('div', { className: 'ob-calendar-status', role: 'status' }, notice) : null)),
        eventPeek ? h(Overlay, null,
          h('div', {
            className: 'ob-event-peek-dismiss',
            style: { position: 'fixed', inset: 0, zIndex: 1090 },
            onClick: () => setEventPeek(null),
            onContextMenu: (e) => { e.preventDefault(); setEventPeek(null) },
          },
            h('div', {
              className: 'ob-event-peek',
              role: 'dialog',
              'aria-label': '日程详情',
              style: { left: eventPeek.left, top: eventPeek.top },
              onClick: (e) => e.stopPropagation(),
              onContextMenu: (event) => eventContext(event, eventPeek.item, eventPeek.day),
            },
              h('div', { className: 'ob-event-peek-accent', style: { background: eventColor(eventPeek.item) } }),
              h('div', { className: 'ob-event-peek-body' },
                h('div', { className: 'ob-event-peek-cat' },
                  h('span', { className: 'ob-event-peek-dot', style: { background: eventColor(eventPeek.item) } }),
                  h('span', null, calendarsById.get(eventPeek.item.calendarId)?.name || eventPeek.item.calendar || '日历')),
                h('h3', { className: 'ob-event-peek-title' }, eventPeek.item.subject || '(无主题)'),
                h('div', { className: 'ob-event-peek-row' },
                  h(MailIcon, { name: 'clock', size: 16 }),
                  h('span', null, eventPeek.item.allDay
                    ? `${eventPeek.day.getFullYear()}/${eventPeek.day.getMonth()+1}/${eventPeek.day.getDate()} · 全天`
                    : `${eventPeek.day.getFullYear()}/${eventPeek.day.getMonth()+1}/${eventPeek.day.getDate()} ${fmtTime(parseDate(eventPeek.item.start))} – ${fmtTime(parseDate(eventPeek.item.end))}`)),
                eventPeek.item.location
                  ? h('div', { className: 'ob-event-peek-row' }, h(MailIcon, { name: 'location', size: 16 }), h('span', null, eventPeek.item.location))
                  : null),
              eventPeek.expanded
                ? h('div', { className: 'ob-event-peek-extra' }, [eventPeek.item.organizer, eventPeek.item.webLink ? '可在 Outlook 中打开' : ''].filter(Boolean).join('\\n') || '暂无更多信息')
                : null,
              h('div', { className: 'ob-event-peek-actions' },
                h('button', {
                  type: 'button', className: 'ob-command',
                  disabled: !eventCanEdit(eventPeek.item),
                  onClick: () => { const item = eventPeek.item; setEventPeek(null); openEdit(item) },
                }, h(MailIcon, { name: 'edit', size: 15 }), '编辑'),
                h('button', {
                  type: 'button', className: 'ob-command',
                  disabled: !eventCanEdit(eventPeek.item),
                  onClick: () => deletePeekEvent(eventPeek.item),
                }, h(MailIcon, { name: 'trash', size: 15 }), '删除'),
                h('button', {
                  type: 'button', className: 'ob-command ob-event-peek-expand',
                  'aria-label': eventPeek.expanded ? '收起' : '展开',
                  onClick: () => setEventPeek({ ...eventPeek, expanded: !eventPeek.expanded }),
                }, h(MailIcon, { name: 'chevron', size: 16 })))))
        ) : null,
        dayPeek ? h(Overlay, null,
          h('div', {
            style: { position: 'fixed', inset: 0, zIndex: 1090 },
            onClick: closeDayPeek,
          },
            h('div', {
              className: 'ob-day-peek',
              role: 'dialog',
              'aria-label': '当天日程',
              style: { left: dayPeek.left, top: dayPeek.top },
              onClick: (e) => e.stopPropagation(),
            },
              h('div', { className: 'ob-day-peek-head' },
                `${dayPeek.day.getMonth()+1} 月 ${dayPeek.day.getDate()} 日`,
                h('button', { type: 'button', className: 'ob-btn', onClick: closeDayPeek }, t('action.close'))),
              (byDay.get(dayKey(dayPeek.day)) ?? []).map((event) => h('button', {
                type: 'button',
                key: event.id,
                className: 'ob-day-peek-item',
                style: { borderLeftColor: eventColor(event), '--ob-event-color': eventColor(event) },
                onClick: (e) => { e.stopPropagation(); scheduleEventClick(event, dayPeek.day, e.currentTarget, 500) },
                onDoubleClick: (e) => { e.preventDefault(); e.stopPropagation(); handleEventDoubleClick(event) },
                onContextMenu: (mouse) => {
                  if (eventClickTimer.current) clearTimeout(eventClickTimer.current)
                  eventClickTimer.current = null
                  eventContext(mouse, event, dayPeek.day)
                },
              },
                h('strong', null, event.subject),
                h('span', null, event.allDay ? '全天' : `${fmtTime(parseDate(event.start))} – ${fmtTime(parseDate(event.end))}`)))))
        ) : null,
        creating ? h(EventEditor, { key: form.transactionId, form, setForm, calendars, saving, error: createError, onSave: submitCreate, onDelete: removeEvent, onDeleteTimelineEvent: deletePeekEvent, onTimelineChange: load, autoTime: creating.autoTime, defaultCalendarAction, onClose: () => { if (!saving) setCreating(null) } }) : null,
      )
    }

    const appearanceGroups = {
      mail: ['mailPanel', 'mailHeader', 'mailAiBar', 'mailToolbar', 'mailFolders', 'mail', 'mailReading'],
      calendar: ['calendarPanel', 'calendarHeader', 'calendarAiBar', 'calendarToolbar', 'calendarList', 'calendar'],
      todo: ['todoPanel', 'todoHeader', 'todoAiBar', 'todoCreate', 'todoCategories', 'todoCompleted', 'todoAll', 'todoToday', 'todoTomorrow'],
    }

    const appearanceSharedLabels = {
      mailPanel: 'panel', calendarPanel: 'panel', todoPanel: 'panel',
      mailHeader: 'header', calendarHeader: 'header', todoHeader: 'header',
      mailAiBar: 'aiBar', calendarAiBar: 'aiBar', todoAiBar: 'aiBar',
      mailToolbar: 'toolbar', calendarToolbar: 'toolbar',
    }

    const appearanceIcons = {
      panel: 'panel', header: 'header', aiBar: 'compose', toolbar: 'filter', mailFolders: 'inbox', mail: 'mail', mailReading: 'notes',
      calendarList: 'columns', calendar: 'calendar',
      todoCreate: 'compose', todoCategories: 'filter', todoCompleted: 'check', todoAll: 'columns', todoToday: 'calendar', todoTomorrow: 'calendar',
    }

    function AppearanceCard({ t, region, values, onChange, disabled }) {
      const opacity = values?.opacity ?? 100
      const blur = values?.blur ?? 0
      const labelKey = appearanceSharedLabels[region] ?? region
      const label = t(`ui.${labelKey}`)
      const slider = (field, value, max, unit) => h('label', { className: 'ob-slider' },
        h('span', null, t(`ui.${field}`)),
        h('input', {
          type: 'range', min: 0, max, step: 1, value, disabled,
          'aria-valuetext': `${value}${unit}`,
          onInput: (event) => onChange(region, field, event.currentTarget.valueAsNumber),
        }),
        h('b', { 'aria-hidden': true }, `${value}${unit}`))
      return h('section', { className: 'ob-surface-card', 'aria-label': label },
        h('div', { className: 'ob-surface-title' },
          h('span', { className: 'ob-surface-icon' }, h(MailIcon, { name: appearanceIcons[labelKey], size: 18 })),
          h('h3', null, label),
          h('span', { className: 'ob-surface-badge', 'aria-hidden': true }, `${opacity}%`)),
        slider('opacity', opacity, 100, '%'),
        slider('blur', blur, 40, 'px'))
    }

    function AppearancePage(props) {
      const { t, view, onView, appearance, setAppearance, onSave, onReset, onClose, busy } = props
      const [feedback, setFeedback] = React.useState(null)
      const update = (key, field, value) => {
        setFeedback(null)
        setAppearance((current) => ({ ...current, [key]: { ...current?.[key], [field]: value } }))
      }
      const save = async () => {
        setFeedback(null)
        try {
          await onSave()
          setFeedback({ text: t('ui.saved') })
        } catch (failure) {
          setFeedback({ text: failure.message, error: true })
        }
      }
      return h('div', {
        className: 'ob-modal-mask ob-appearance-mask',
        onClick: (event) => { if (event.target === event.currentTarget) onClose() },
      },
        h('div', {
          className: 'ob-modal ob-appearance-dialog', role: 'dialog', 'aria-modal': true, 'aria-labelledby': 'ob-ui-title',
          onKeyDown: (event) => { if (event.key === 'Escape') onClose() },
        },
          h('header', { className: 'ob-modal-head' },
            h('span', { id: 'ob-ui-title' }, t('ui.title')),
            h('button', { type: 'button', className: 'ob-btn', onClick: onClose }, t('action.close'))),
          h('div', { className: 'ob-ui-page' },
            h('div', { className: 'ob-ui-head' },
              h('div', { className: 'ob-ui-kicker' }, 'SURFACES'),
              h('h2', null, t('ui.surfaces')),
              h('p', null, t('ui.hint'))),
            h('div', { className: 'ob-appearance-tabs' }, Object.keys(appearanceGroups).map((key) =>
              h('button', { key, type: 'button', className: 'ob-btn', 'aria-pressed': view === key,
                onClick: () => onView(key) }, t(`tab.${key}`)))),
            h('div', { className: 'ob-surface-grid' }, appearanceGroups[view].map((key) =>
              h(AppearanceCard, { key, t, region: key, values: appearance?.[key], onChange: update, disabled: busy })))),
          h('footer', { className: 'ob-ui-foot' },
            feedback ? h('span', { className: 'ob-ui-feedback', role: feedback.error ? 'alert' : 'status' }, feedback.text) : null,
            h('button', { type: 'button', className: 'ob-btn', disabled: busy,
              onClick: () => { setFeedback(null); onReset() } }, t('ui.reset')),
            h('button', { type: 'button', className: 'ob-btn primary', disabled: busy, onClick: save }, t('ui.save')))))
    }

    function CalendarBlock(props) {
      const { events, t } = props
      const groups = []
      const index = new Map()
      for (const event of events) {
        const start = parseDate(event.start)
        if (!start) continue
        const key = dayKey(start)
        if (!index.has(key)) {
          const group = { key, date: start, items: [] }
          index.set(key, group)
          groups.push(group)
        }
        index.get(key).items.push({ event, start })
      }
      return h(
        'section',
        { className: 'ob-card grow' },
        h(
          'div',
          { className: 'ob-card-head' },
          h('span', { className: 'ob-card-title' }, t('calendar.title')),
          h('span', { className: 'ob-count' }, groups.length > 0 ? `${events.length}` : ''),
        ),
        h(
          'div',
          { className: 'ob-list' },
          groups.length === 0
            ? h('div', { className: 'ob-empty' }, t('calendar.empty'))
            : groups.map((group) =>
                h(
                  'div',
                  { className: 'ob-daygroup', key: group.key },
                  h('div', { className: 'ob-dayhead' }, fmtDay(group.date, t)),
                  group.items.map((item) =>
                    h(
                      'div',
                      { className: 'ob-event', key: item.event.id },
                      h('span', { className: 'ob-event-time' }, item.event.allDay ? t('calendar.allDay') : fmtTime(item.start)),
                      h('span', { className: 'ob-bar' }),
                      h(
                        'span',
                        { className: 'ob-event-main' },
                        h('div', { className: 'ob-event-title' }, item.event.subject),
                        item.event.location || item.event.organizer
                          ? h(
                              'div',
                              { className: 'ob-event-meta' },
                              [item.event.location, item.event.organizer].filter(Boolean).join(' 路 '),
                            )
                          : null,
                      ),
                    ),
                  ),
                ),
              ),
        ),
      )
    }

    /* --------------------------------------------------------- todo block */

    function parsePlanTitle(input, currentYear) {
      const title = input.trim()
      const suffix = /\/(?:(\d{4})\.)?(\d{1,2})\.(\d{1,2})$/.exec(title)
      if (!suffix) return { title, dueDate: '' }
      const name = title.slice(0, suffix.index).trimEnd()
      const year = suffix[1] ?? String(currentYear).padStart(4, '0')
      const dueDate = `${year}-${pad(suffix[2])}-${pad(suffix[3])}`
      const date = new Date(`${dueDate}T00:00:00Z`)
      if (!name || year === '0000' || !Number.isFinite(+date) || date.toISOString().slice(0, 10) !== dueDate) {
        return { title, dueDate: '' }
      }
      return { title: name, dueDate }
    }

    function planTimestamp(value) {
      if (!value) return NaN
      // Graph's default completion time is UTC and may omit the offset.
      return Date.parse(/(?:Z|[+-]\d{2}:\d{2})$/i.test(value) ? value : `${value}Z`)
    }

    function groupPlanTasks(tasks, date, { field, ascending }) {
      const today = dayKey(date)
      const nextDay = new Date(date)
      nextDay.setDate(nextDay.getDate() + 1)
      const tomorrow = dayKey(nextDay)
      const groups = { all: [], overdue: [], today: [], tomorrow: [], completed: [], categories: new Map() }
      for (const task of tasks) {
        if (task.completed) { groups.completed.push(task); continue }
        groups.all.push(task)
        const due = task.due?.slice(0, 10)
        if (due === today) groups.today.push(task)
        if (due === tomorrow) groups.tomorrow.push(task)
        if (due && due < today) groups.overdue.push(task)
        for (const name of new Set(task.categories ?? [])) {
          if (!groups.categories.has(name)) groups.categories.set(name, [])
          groups.categories.get(name).push(task)
        }
      }
      groups.completed.sort((a, b) => {
        const first = planTimestamp(a[field])
        const second = planTimestamp(b[field])
        if (Number.isNaN(first)) return Number.isNaN(second) ? 0 : 1
        if (Number.isNaN(second)) return -1
        return ascending ? first - second : second - first
      })
      return groups
    }

    function PlanCompletionTime({ value }) {
      const timestamp = planTimestamp(value)
      if (!Number.isFinite(timestamp)) return null
      const date = new Date(timestamp)
      return h('time', { className: 'ob-todo-completion-time', dateTime: date.toISOString() },
        h('span', null, `${fmtTime(date)}:${pad(date.getSeconds())}`),
        h('span', null, `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`))
    }

    function TodoRows({ tasks, t, onToggle, busy, canDrag, dragged, sourceKey, sourceCategory, onStartDrag, onEndDrag, onContextTask }) {
      return tasks.map((task) => h('div', {
        className: `ob-todo${task.completed ? ' ob-todo-completed' : ''}${dragged?.id === task.id && dragged.listId === task.listId ? ' ob-plan-dragging' : ''}`,
        key: `${task.listId}:${task.id}`,
        draggable: canDrag && !task.completed,
        onDragStart: (event) => task.completed ? event.preventDefault() : onStartDrag(event, task, sourceKey, sourceCategory),
        onDragEnd: onEndDrag,
        onContextMenu: (event) => onContextTask(event, task, sourceKey),
      },
        h('input', {
          type: 'checkbox', checked: Boolean(task.completed),
          'aria-label': t(task.completed ? 'todo.restore' : 'todo.complete', { title: task.title }),
          title: t(task.completed ? 'todo.restore' : 'todo.complete', { title: task.title }),
          disabled: busy, onChange: () => onToggle(task),
        }),
        h('div', { className: 'ob-todo-main' },
          h('div', { className: 'ob-todo-title' }, task.title),
          h('div', { className: 'ob-todo-meta' },
            [task.due ? t('todo.due', { time: task.due.slice(0, 10) }) : '', task.listName].filter(Boolean).join(' · '))),
        task.completed ? h(PlanCompletionTime, { value: task.completedAt }) : null,
        canDrag && !task.completed ? h('span', { className: 'ob-plan-grip', 'aria-hidden': true }, '⠿') : null))
    }

    function TodoPanel({ label, count, empty, overdue, target, headerControl, surface, addEvents, children }) {
      return h('section', {
        className: `ob-card ob-plan-panel${overdue ? ' ob-plan-overdue' : ''}${target?.preview ? ' ob-plan-drop-hot' : ''}`,
        'aria-label': label, style: surface, ...target?.events, ...addEvents,
      },
        h('div', { className: 'ob-card-head' }, h('h3', null, label), headerControl, h('span', { className: 'ob-count' }, count)),
        target?.preview,
        h('div', { className: 'ob-list' }, children ?? (target?.preview ? null : h('div', { className: 'ob-empty ob-plan-empty' }, empty))))
    }

    function TodoPage({ t, lists, tasks, error, stale, connected, busy, onCreate, onToggle, onMove, onDelete, onRefresh, completedSort, onCompletedSort, surfaceOf }) {
      const [listId, setListId] = React.useState('')
      const [title, setTitle] = React.useState('')
      const [dueDate, setDueDate] = React.useState('')
      const [categoryFilter, setCategoryFilter] = React.useState('overdue')
      const [dragged, setDragged] = React.useState(null)
      const [hover, setHover] = React.useState('')
      const dragRef = React.useRef(null)
      const [quickAdd, setQuickAdd] = React.useState(null)
      const [quickTitle, setQuickTitle] = React.useState('')
      const quickSubmitting = React.useRef(false)
      const quickField = React.useRef(null)
      React.useLayoutEffect(() => {
        const field = quickField.current
        if (!quickAdd || !field) return
        const resize = () => {
          field.style.height = '0px'
          const contentHeight = field.scrollHeight
          const height = Math.max(46, Math.min(contentHeight, window.innerHeight - 18))
          field.style.height = `${height}px`
          field.style.overflowY = contentHeight > height ? 'auto' : 'hidden'
          const box = field.form
          box.style.top = `${Math.max(8, Math.min(quickAdd.top, window.innerHeight - box.offsetHeight - 8))}px`
          box.style.left = `${Math.max(8, Math.min(quickAdd.left, window.innerWidth - box.offsetWidth - 8))}px`
        }
        resize()
        window.addEventListener('resize', resize)
        return () => window.removeEventListener('resize', resize)
      }, [quickAdd, quickTitle])
      const selectedList = listId ? lists.find((list) => list.id === listId) : lists[0]
      const defaultTaskList = lists.find((list) => list.isDefault) || lists.find((list) => /^(任务|Tasks)$/i.test(list.name))
      const canQuickAdd = connected && !busy && Boolean(defaultTaskList)
      const canAdd = connected && Boolean(selectedList) && !busy
      const canDrag = connected && !busy
      const now = new Date()
      const today = dayKey(now)
      const tomorrow = addDaysYmd(today, 1)
      const draftPlan = parsePlanTitle(title, now.getFullYear())
      const selectedDueDate = draftPlan.dueDate || dueDate
      const chooseDueDate = (date) => {
        if (draftPlan.dueDate) setTitle(draftPlan.title)
        setDueDate(date)
      }
      const quickAddDates = new Map([['all', ''], ['today', today], ['tomorrow', tomorrow]])
      const prepareQuickAdd = (key, x, y) => {
        if (!canQuickAdd || !quickAddDates.has(key) || quickSubmitting.current) return
        setQuickTitle('')
        setQuickAdd({
          date: quickAddDates.get(key),
          hint: key === 'today' ? '添加计划到今天' : key === 'tomorrow' ? '添加计划到明天' : '添加计划到所有项目',
          left: Math.max(8, Math.min(x, window.innerWidth - 318)),
          top: Math.max(8, Math.min(y, window.innerHeight - 56)),
        })
      }
      const quickAddItem = (key, { clientX, clientY }) => ({
        label: t('todo.add'), icon: 'compose', disabled: !canQuickAdd, run: () => prepareQuickAdd(key, clientX, clientY),
      })
      const quickAddEvents = (key) => ({
        onDoubleClick: (event) => {
          if (!event.target.closest('.ob-todo,button,input,select')) prepareQuickAdd(key, event.clientX, event.clientY)
        },
        onContextMenu: (event) => {
          if (!event.target.closest('.ob-todo,button,input,select')) openContextMenu(event, [quickAddItem(key, event)])
        },
      })
      const groups = groupPlanTasks(tasks, now, completedSort)
      const categoryNames = [...groups.categories.keys()].filter(Boolean).sort((a, b) => a.localeCompare(b))
      const categoryName = categoryNames.find((name) => categoryFilter === `category:${name}`)
      const categoryList = lists.find((list) => categoryFilter === `list:${list.id}`)
      const categoryKey = categoryList ? `list:${categoryList.id}` : categoryName ? `category:${categoryName}` : 'overdue'
      const showOverdue = categoryKey === 'overdue'
      const categoryTasks = categoryList
        ? groups.all.filter((task) => task.listId === categoryList.id)
        : showOverdue ? groups.overdue : groups.categories.get(categoryName)
      const onContextTask = (event, task, sourceKey) => {
        const currentDue = task.due?.slice(0, 10) ?? null
        const setDue = (label, dueDate) => ({
          label, icon: 'calendar', disabled: !canDrag || currentDue === dueDate,
          run: () => onMove(task, { dueDate }),
        })
        openContextMenu(event, [
          ...(quickAddDates.has(sourceKey) ? [quickAddItem(sourceKey, event), { separator: true }] : []),
          {
            label: t(task.completed ? 'todo.menuRestore' : 'todo.menuComplete'),
            icon: task.completed ? 'reply' : 'check', disabled: !canDrag,
            run: () => onToggle(task),
          },
          ...(!task.completed ? [
            { separator: true },
            setDue(t('todo.menuDue', { day: t('calendar.today') }), today),
            setDue(t('todo.menuDue', { day: t('calendar.tomorrow') }), tomorrow),
            setDue(t('todo.menuClearDue'), null),
          ] : []),
          { separator: true },
          { label: t('todo.menuCopyTitle'), icon: 'notes', run: () => copyOutlookText(task.title) },
          { separator: true },
          { label: t('todo.menuDelete'), icon: 'trash', danger: true, disabled: !canDrag || !lists.some((list) => list.id === task.listId && list.canDeleteTasks), run: () => onDelete(task) },
        ])
      }
      const endDrag = () => {
        dragRef.current = null
        setDragged(null)
        setHover('')
      }
      const startDrag = (event, task, sourceKey, sourceCategory) => {
        if (!canDrag) { event.preventDefault(); return }
        event.dataTransfer.effectAllowed = 'move'
        event.dataTransfer.setData('application/x-outlook-plan', 'move')
        dragRef.current = { task, sourceKey, sourceCategory }
        setDragged(task)
      }
      const dropTarget = (key, change, hint) => ({
        preview: hover === key && dragged
          ? h('div', { className: 'ob-plan-drop-preview', 'aria-hidden': true }, h('strong', null, dragged.title), h('span', null, hint))
          : null,
        events: {
          onDragOver: (event) => {
            if (!canDrag || !dragRef.current || dragRef.current.sourceKey === key) return
            event.preventDefault()
            event.dataTransfer.dropEffect = 'move'
            setHover(key)
          },
          onDragLeave: (event) => {
            if (!event.currentTarget.contains(event.relatedTarget)) setHover((current) => current === key ? '' : current)
          },
          onDrop: (event) => {
            const active = dragRef.current
            if (!canDrag || !active || active.sourceKey === key) return
            event.preventDefault()
            endDrag()
            const { task, sourceCategory } = active
            if ('dueDate' in change && (task.due?.slice(0, 10) ?? null) === change.dueDate) return
            if ('category' in change && (change.category === sourceCategory || (change.category === null && !task.categories?.length))) return
            onMove(task, { ...change, sourceCategory })
          },
        },
      })
      const dateTarget = (key, date) => dropTarget(key, { dueDate: date },
        date ? t('todo.dropDate', { date }) : t('todo.dropNoDate'))
      let categoryTarget = null
      if (showOverdue) categoryTarget = dateTarget('overdue', addDaysYmd(today, -1))
      else if (categoryName) categoryTarget = dropTarget(categoryKey, { category: categoryName },
        t('todo.dropCategory', { name: categoryName }))
      const rows = (items, sourceKey, sourceCategory) => items.length ? h(TodoRows, {
        tasks: items, t, busy, onToggle, canDrag, dragged, sourceKey, sourceCategory, onStartDrag: startDrag, onEndDrag: endDrag,
        onContextTask,
      }) : null
      const panel = (label, items, key, date, region) => h(TodoPanel, {
        label: t(label), count: items.length, empty: t('todo.empty'), target: dateTarget(key, date), surface: surfaceOf(region),
        addEvents: quickAddEvents(key),
      }, rows(items, key))
      return h('section', { className: 'ob-todo-page', 'aria-label': t('todo.title') },
        error ? h('div', { className: 'ob-error ob-plan-sync-error', role: 'alert' },
          h('span', null, error, stale ? ` ${t('todo.stale')}` : ''),
          h('button', { type: 'button', className: 'ob-btn', disabled: busy, onClick: onRefresh }, t('action.refresh'))) : null,
        connected ? null : h('div', { className: 'ob-banner' }, t('todo.demo')),
        h('header', { className: 'ob-plan-header', style: surfaceOf('todoCreate') },
          h('div', { className: 'ob-plan-heading' },
            h('h2', null, t('todo.title')),
            h('div', { className: 'ob-plan-heading-details' },
              h('span', { className: 'ob-plan-summary' }, t('todo.summary', { count: groups.all.length })),
              connected ? h('div', { className: 'ob-plan-hints' },
                h('span', { className: 'ob-plan-drag-hint' }, t('todo.dragHint')),
                h('span', { className: 'ob-plan-drag-hint' }, t('todo.dateHint'))) : null)),
          h('form', { className: 'ob-todo-form', onSubmit: async (event) => {
            event.preventDefault()
            if (!canAdd || !draftPlan.title) return
            setListId(selectedList.id)
            if (await onCreate(selectedList, draftPlan.title, selectedDueDate)) {
              setTitle('')
              setDueDate('')
            }
          } },
            h('label', null, t('todo.list'),
              h('select', { className: 'ob-input', value: selectedList?.id ?? '', disabled: busy || lists.length === 0,
                onChange: (event) => setListId(event.target.value) },
                selectedList ? null : h('option', { value: '', disabled: true }, t(lists.length ? 'todo.chooseList' : 'todo.noLists')),
                lists.map((list) => h('option', { key: list.id, value: list.id }, list.name)))),
            h('input', { className: 'ob-input ob-todo-input', value: title, maxLength: 255, disabled: !canAdd,
              placeholder: t('todo.newWithDate'), 'aria-label': t('todo.new'), onChange: (event) => setTitle(event.target.value) }),
            h('button', { type: 'submit', className: 'ob-btn primary', disabled: !canAdd || !draftPlan.title }, t('todo.add')),
            h('div', { className: 'ob-todo-due-row' },
              h('label', null, t('todo.dueDate'), h('input', { className: 'ob-input', type: 'date', value: selectedDueDate,
                disabled: !canAdd, onChange: (event) => chooseDueDate(event.target.value),
                onMouseDown: (event) => {
                  if (event.button === 0 && typeof event.currentTarget.showPicker === 'function') event.preventDefault()
                },
                onClick: (event) => {
                  const field = event.currentTarget
                  if (field.disabled || typeof field.showPicker !== 'function') return
                  try {
                    field.showPicker()
                    event.preventDefault()
                  } catch { /* Let the native calendar button handle the click. */ }
                },
              })),
              [[today, 'calendar.today'], [tomorrow, 'calendar.tomorrow']].map(([date, label]) =>
                h('button', { key: label, type: 'button', className: 'ob-btn', disabled: !canAdd,
                  'aria-pressed': selectedDueDate === date, onClick: () => chooseDueDate(date) }, t(label)))))),
        h('div', { className: 'ob-plan-grid' },
          h('div', { className: 'ob-plan-side' },
            h(TodoPanel, { label: t('todo.categories'), count: categoryTasks.length,
              surface: surfaceOf('todoCategories'),
              empty: t(showOverdue ? 'todo.noOverdue' : 'todo.empty'), overdue: showOverdue, target: categoryTarget,
              headerControl: h('select', { className: 'ob-plan-category-filter', value: categoryKey,
                'aria-label': t('todo.categoryFilter'), onChange: (event) => setCategoryFilter(event.target.value) },
                h('option', { value: 'overdue' }, t('todo.overdue')),
                lists.map((list) => h('option', { key: list.id, value: `list:${list.id}` }, list.name)),
                categoryNames.length ? h('optgroup', { label: t('todo.categories') }, categoryNames.map((name) =>
                  h('option', { key: name, value: `category:${name}` }, name))) : null),
            }, rows(categoryTasks, categoryKey, categoryName)),
            h(TodoPanel, { label: t('todo.completed'), count: groups.completed.length, empty: t('todo.noCompleted'),
              surface: surfaceOf('todoCompleted'),
              headerControl: h('div', { className: 'ob-plan-sort' },
                h('select', { value: completedSort.field, 'aria-label': t('todo.sortCompleted'), onChange: (event) => {
                  const field = event.target.value
                  onCompletedSort((current) => ({ ...current, field }))
                } },
                  h('option', { value: 'completedAt' }, t('todo.sortCompletion')),
                  h('option', { value: 'createdAt' }, t('todo.sortCreation'))),
                h('button', { type: 'button', className: 'ob-btn ob-plan-sort-direction',
                  'aria-label': t('todo.sortAscending'), 'aria-pressed': completedSort.ascending,
                  onClick: () => onCompletedSort((current) => ({ ...current, ascending: !current.ascending })),
                }, completedSort.ascending ? '↑' : '↓')),
            },
              rows(groups.completed, 'completed'))),
          panel('todo.all', groups.all, 'all', null, 'todoAll'),
          panel('calendar.today', groups.today, 'today', today, 'todoToday'),
          panel('calendar.tomorrow', groups.tomorrow, 'tomorrow', tomorrow, 'todoTomorrow')),
        quickAdd ? h(Overlay, null,
          h('div', { className: 'ob-plan-quick-add-layer', onPointerDown: (event) => {
            if (event.target === event.currentTarget && !quickSubmitting.current) setQuickAdd(null)
          } }, h('form', {
            className: 'ob-plan-quick-add', style: { left: quickAdd.left, top: quickAdd.top },
            'aria-label': quickAdd.hint,
            onSubmit: async (event) => {
              event.preventDefault()
              const name = quickTitle.trim()
              if (!name || !canQuickAdd || quickSubmitting.current) return
              quickSubmitting.current = true
              try {
                if (await onCreate(defaultTaskList, name, quickAdd.date)) setQuickAdd(null)
              } finally { quickSubmitting.current = false }
            },
          }, h('textarea', {
            ref: quickField, rows: 1, autoFocus: true, maxLength: 255, value: quickTitle, readOnly: busy,
            'aria-label': quickAdd.hint, placeholder: quickAdd.hint,
            onChange: (event) => setQuickTitle(event.target.value),
            onKeyDown: (event) => {
              if (event.nativeEvent?.isComposing || event.keyCode === 229) return
              if (event.key === 'Enter') {
                event.preventDefault()
                event.currentTarget.form.requestSubmit()
                return
              }
              if (event.key === 'Escape' && !quickSubmitting.current) {
                event.preventDefault()
                event.stopPropagation()
                setQuickAdd(null)
              }
            },
          })))) : null)
    }

    /* ------------------------------------------------------ connect views */

    function ConnectPanel(props) {
      const {
        t, builtin, code, error, busy, paste,
        onPasteChange, onOpenInApp, onOpenSystem, onComplete,
        onStart, onCopy, onOpenSettings,
      } = props
      return h(
        'div',
        { className: 'ob-connect' },
        h('h3', null, t('connect.title')),
        h('p', null, t('auth.step1')),
        h(
          'div',
          { style: { display: 'flex', gap: '8px', flexWrap: 'wrap' } },
          h(
            'button',
            { type: 'button', className: 'ob-btn primary', disabled: busy, onClick: onOpenInApp },
            t('auth.open'),
          ),
        ),
        h('p', null, t('auth.step2')),
        h('textarea', {
          className: 'ob-textarea',
          rows: 3,
          spellCheck: false,
          value: paste,
          placeholder: 'https://login.microsoftonline.com/common/oauth2/nativeclient?code=...',
          onChange: (event) => onPasteChange(event.target.value),
        }),
        h(
          'div',
          { style: { display: 'flex', gap: '8px', flexWrap: 'wrap' } },
          h(
            'button',
            {
              type: 'button',
              className: 'ob-btn primary',
              disabled: busy || !String(paste || '').trim(),
              onClick: onComplete,
            },
            t('auth.complete'),
          ),
          h(
            'button',
            { type: 'button', className: 'ob-btn', onClick: onOpenSystem },
            t('auth.openSystem'),
          ),
        ),
        builtin ? h('p', null, t('connect.builtin')) : null,
        error ? h('div', { className: 'ob-error' }, error) : null,
        h('div', { className: 'ob-divider' }, t('auth.orDevice')),
        code
          ? h(
              'div',
              { style: { display: 'flex', flexDirection: 'column', gap: '10px' } },
              h('div', { className: 'ob-sub' }, t('connect.code')),
              h(
                'div',
                { style: { display: 'flex', gap: '8px', alignItems: 'center', flexWrap: 'wrap' } },
                h('div', { className: 'ob-code' }, code.userCode),
                h(
                  'button',
                  { type: 'button', className: 'ob-btn', onClick: () => onCopy(code.userCode) },
                  t('connect.copy'),
                ),
                h(
                  'span',
                  { className: 'ob-sub' },
                  t('connect.remain', { n: Math.max(1, Math.round((code.expiresIn || 900) / 60)) }),
                ),
              ),
              h('div', { className: 'ob-sub' }, t('connect.waiting')),
            )
          : null,
        h(
          'div',
          { style: { display: 'flex', gap: '8px', flexWrap: 'wrap' } },
          h('button', { type: 'button', className: 'ob-btn', disabled: busy, onClick: onStart }, t('connect.start')),
          h('button', { type: 'button', className: 'ob-btn', onClick: onOpenSettings }, t('action.settings')),
        ),
      )
    }

    function SetupGuide({ t, version, connected, busy, onOpenEntra, onSettings, onConnect, onClose }) {
      const entraUrl = 'https://entra.microsoft.com/'
      const steps = [
        ['register', h('a', { className: 'ob-btn', href: entraUrl, target: '_blank', rel: 'noopener noreferrer',
          onClick: (event) => { if (onOpenEntra(entraUrl)) event.preventDefault() } }, t('guide.openEntra'))],
        ['publicClient'],
        ['configure', h('button', { type: 'button', className: 'ob-btn', onClick: onSettings }, t('guide.openSettings'))],
        ['authorize', h('button', { type: 'button', className: 'ob-btn primary', disabled: busy, onClick: onConnect },
          t(connected ? 'action.reauth' : 'action.connect'))],
      ]
      return h('div', { className: 'ob-modal-mask',
        onClick: (event) => { if (event.target === event.currentTarget) onClose() } },
        h('section', { className: 'ob-modal ob-guide-dialog', role: 'dialog', 'aria-modal': true, 'aria-labelledby': 'ob-guide-title',
          onKeyDown: (event) => { if (event.key === 'Escape') onClose() } },
          h('header', { className: 'ob-modal-head' },
            h('div', null, h('span', { id: 'ob-guide-title' }, t('guide.title')),
              version ? h('span', { className: 'ob-guide-version' }, `${t('guide.version')} v${version}`) : null),
            h('button', { type: 'button', className: 'ob-btn', autoFocus: true, onClick: onClose }, t('action.close'))),
          h('div', { className: 'ob-guide-body' },
            h('p', null, t('guide.intro')),
            h('ol', { className: 'ob-guide-steps' }, steps.map(([key, action], index) =>
              h('li', { key },
                h('span', { className: 'ob-guide-number', 'aria-hidden': true }, index + 1),
                h('div', null, h('h3', null, t(`guide.${key}`)), h('p', null, t(`guide.${key}Help`)), action)))))))
    }

    function SettingsPanel(props) {
      const { t, draft, setDraft, onSave, onClose, busy } = props
      const field = (label, key, type) =>
        h(
          'div',
          { className: 'ob-field' },
          h('label', { htmlFor: `ob-${key}` }, label),
          h('input', {
            id: `ob-${key}`,
            type: type || 'text',
            value: draft[key] ?? '',
            onChange: (event) => setDraft({ ...draft, [key]: event.target.value }),
            placeholder: key === 'clientId' ? '00000000-0000-0000-0000-000000000000' : '',
          }),
        )
      return h(
        'div',
        { className: 'ob-settings' },
        h('h3', { style: { margin: 0, fontSize: '13.5px' } }, t('settings.title')),
        field(t('settings.clientId'), 'clientId'),
        h('small', { style: { color: 'var(--dsw-alias-label-quaternary)', fontSize: '11px' } }, t('settings.clientIdHint')),
        field(t('settings.clientSecret'), 'clientSecret', 'password'),
        h('small', { style: { color: 'var(--dsw-alias-label-quaternary)', fontSize: '11px' } }, t('settings.clientSecretHint')),
        h(
          'div',
          { className: 'ob-settings-row' },
          field(t('settings.tenant'), 'tenant'),
          field(t('settings.mailTop'), 'mailTop', 'number'),
          field(t('settings.eventDays'), 'eventDays', 'number'),
        ),
        h(
          'label',
          { className: 'ob-switch' },
          h('input', {
            type: 'checkbox',
            checked: Boolean(draft.demo),
            onChange: (event) => setDraft({ ...draft, demo: event.target.checked }),
          }),
          t('settings.demo'),
        ),
        h(
          'div',
          { style: { display: 'flex', gap: '8px' } },
          h('button', { type: 'button', className: 'ob-btn primary', disabled: busy, onClick: onSave }, t('action.save')),
          h('button', { type: 'button', className: 'ob-btn', onClick: onClose }, t('action.close')),
        ),
      )
    }

    /* ----------------------------------------------------------- the page */

    function MainPage(props) {
      const t = props?.t || fbT
      const openInApp = props?.openInApp || (() => false)
      const snap = useSnapshot()
      const [filter, setFilter] = React.useState('all')
      const [selected, setSelected] = React.useState(null)
      const [showSettings, setShowSettings] = React.useState(false)
      const [code, setCode] = React.useState(null)
      const [busy, setBusy] = React.useState(false)
      const [error, setError] = React.useState(null)
      const [notice, setNotice] = React.useState(null)
      const [draft, setDraft] = React.useState({ clientId: '', clientSecret: '', tenant: 'consumers', demo: true, mailTop: 12, eventDays: 7 })
      const [paste, setPaste] = React.useState('')
      const [view, setView] = React.useState('mail')
      const [completedSort, setCompletedSort] = React.useState({ field: 'completedAt', ascending: false })
      const [showUi, setShowUi] = React.useState(false)
      const [showGuide, setShowGuide] = React.useState(false)
      const [appearance, setAppearance] = React.useState(null)
      const [replyRequest, setReplyRequest] = React.useState(null)
      const [category, setCategory] = React.useState('all')
      const [focus, setFocus] = React.useState('focused')
      const [folder, setFolder] = React.useState('inbox')
      const [readAllFeedback, setReadAllFeedback] = React.useState(null)
      const activeFolder = React.useRef(folder)
      activeFolder.current = folder
      const folderLoads = React.useRef(new Map())
      const [folders, setFolders] = React.useState([])
      const [folderMail, setFolderMail] = React.useState(null)
      const [compose, setCompose] = React.useState(null)
      const [aiText, setAiText] = React.useState('')
      const [aiBusy, setAiBusy] = React.useState(false)
      const [aiMsg, setAiMsg] = React.useState(null)
      const [aiErr, setAiErr] = React.useState(null)
      const [pendingCreate, setPendingCreate] = React.useState(null)
      const [calendarNonce, setCalendarNonce] = React.useState(0)

      React.useEffect(() => {
        let alive = true
        api('/state')
          .then((state) => {
            if (!alive || !state?.settings) return
            if (state.settings.appearance) {
              const stored = state.settings.appearance
              setAppearance((current) => Object.fromEntries(Object.values(appearanceGroups).flat().map((key) =>
                [key, { ...stored[key], ...current?.[key] }])))
            }
            setDraft({
              clientId: state.settings.clientId ?? '',
              clientSecret: '',
              tenant: state.settings.tenant ?? 'common',
              demo: Boolean(state.settings.demo),
              mailTop: state.settings.mailTop ?? 12,
              eventDays: state.settings.eventDays ?? 7,
            })
          })
          .catch(() => {})
        return () => {
          alive = false
        }
      }, [])

      const openSettings = React.useCallback(async () => {
        setShowSettings(true)
        try {
          const state = await api('/state')
          setDraft({
            clientId: state.settings?.clientId ?? '',
            tenant: state.settings?.tenant ?? 'common',
            demo: Boolean(state.settings?.demo),
            mailTop: state.settings?.mailTop ?? 12,
            eventDays: state.settings?.eventDays ?? 7,
          })
        } catch (failure) {
          setError(failure.message)
        }
      }, [])

      const startConnect = React.useCallback(async () => {
        setBusy(true)
        setError(null)
        try {
          const info = await api('/connect/start', { method: 'POST' })
          setCode(info)
          setNotice(null)
          try {
            await navigator.clipboard.writeText(info.userCode)
            setNotice(t('connect.copied'))
          } catch {
            /* clipboard permission is optional */
          }
          if (!openInApp(info.signInUrl)) {
            try {
              await api('/connect/open', { method: 'POST', body: { url: info.signInUrl } })
            } catch {
              /* the panel still shows the code and the URL */
            }
          }
        } catch (failure) {
          setError(failure.message)
        } finally {
          setBusy(false)
        }
      }, [])

      const copyCode = React.useCallback(async (value) => {
        try {
          await navigator.clipboard.writeText(value)
          setNotice(t('connect.copied'))
        } catch {
          setNotice(value)
        }
      }, [t])

      const openAuthorizeInApp = React.useCallback(async () => {
        setError(null)
        try {
          const info = await api('/auth/url')
          if (!openInApp(info.url)) await api('/auth/open', { method: 'POST' })
        } catch (failure) {
          setError(failure.message)
        }
      }, [])

      const openAuthorizeSystem = React.useCallback(async () => {
        setError(null)
        try {
          await api('/auth/open', { method: 'POST' })
        } catch (failure) {
          setError(failure.message)
        }
      }, [])

      const completeAuth = React.useCallback(async () => {
        setBusy(true)
        setError(null)
        try {
          const result = await api('/auth/complete', { method: 'POST', body: { text: paste } })
          if (result.ok) {
            setPaste('')
            setNotice(t('connect.done'))
            await store.load(true)
          } else {
            setError(result.error || 'failed')
          }
        } catch (failure) {
          setError(failure.message)
        } finally {
          setBusy(false)
        }
      }, [paste, t])

      const openSignIn = React.useCallback(async (url) => {
        try {
          await api('/connect/open', { method: 'POST', body: { url } })
          setNotice(t('connect.opened'))
        } catch (failure) {
          setError(failure.message)
        }
      }, [t])

      React.useEffect(() => {
        if (!code) return
        let stopped = false
        let handle = null
        let tries = 0
        const tick = async () => {
          if (stopped) return
          tries += 1
          try {
            const status = await api('/state')
            if (stopped) return
            if (status.connected) {
              setCode(null)
              setError(null)
              setNotice(t('connect.done'))
              store.load(true)
              return
            }
            const auth = status.auth || {}
            if (auth.status === 'error' || auth.status === 'expired' || auth.status === 'declined') {
              setCode(null)
              setError(auth.message || auth.status)
              return
            }
            if (auth.status === 'idle' && !auth.pending) {
              setCode(null)
              return
            }
          } catch {
            /* transient failures keep the loop alive */
          }
          if (tries > 150) {
            setCode(null)
            setError(t('connect.timeout'))
            return
          }
          handle = setTimeout(tick, 3000)
        }
        handle = setTimeout(tick, 2500)
        return () => {
          stopped = true
          if (handle) clearTimeout(handle)
        }
      }, [code, t])

      const refresh = React.useCallback(async () => {
        setBusy(true)
        await store.load(true)
        setBusy(false)
      }, [])

      const saveSettings = React.useCallback(async () => {
        setBusy(true)
        setError(null)
        try {
          await api('/settings', {
            method: 'POST',
            body: {
              clientId: draft.clientId,
              tenant: draft.tenant,
              demo: Boolean(draft.demo),
              mailTop: Number(draft.mailTop) || 12,
              eventDays: Number(draft.eventDays) || 7,
              ...(String(draft.clientSecret || '').trim()
                ? { clientSecret: String(draft.clientSecret).trim() }
                : {}),
            },
          })
          setShowSettings(false)
          await store.load(true)
        } catch (failure) {
          setError(failure.message)
        } finally {
          setBusy(false)
        }
      }, [draft])

      const disconnect = React.useCallback(async () => {
        setBusy(true)
        try {
          await api('/disconnect', { method: 'POST' })
          setSelected(null)
          await store.load(true)
        } catch (failure) {
          setError(failure.message)
        } finally {
          setBusy(false)
        }
      }, [])

      const markRead = React.useCallback(
        async (row) => {
          setBusy(true)
          try {
            await api('/mail/read', { method: 'POST', body: { id: row.id } })
            setSelected((current) => current?.id === row.id ? { ...current, isRead: true } : current)
            setFolderMail((current) => current?.map((item) => item.id === row.id ? { ...item, isRead: true } : item) ?? current)
            store.patch({
              mail: (store.snapshot?.mail ?? []).map((item) =>
                item.id === row.id ? { ...item, isRead: true } : item,
              ),
              unread: Math.max(0, Number(store.snapshot?.unread || 0) - 1),
            })
          } catch (failure) {
            setError(failure.message)
          } finally {
            setBusy(false)
          }
        },
        [],
      )

      const loadFolders = React.useCallback(async () => {
        try {
          const result = await api('/folders')
          setFolders(result.folders ?? [])
        } catch {
          setFolders([])
        }
      }, [])

      const loadFolderMail = React.useCallback(async (key) => {
        const request = Symbol()
        folderLoads.current.set(key, request)
        const isCurrent = () => activeFolder.current === key && folderLoads.current.get(key) === request
        try {
          const result = await api(`/mail?folder=${encodeURIComponent(key)}`)
          const rows = result.mail ?? []
          if (isCurrent()) setFolderMail(rows)
          return rows
        } catch (failure) {
          if (isCurrent()) setError(failure.message)
        }
      }, [])

      React.useEffect(() => {
        loadFolders()
      }, [loadFolders])

      React.useEffect(() => {
        loadFolderMail(folder)
      }, [folder, loadFolderMail])

      const afterMailAction = React.useCallback(async () => {
        setSelected(null)
        await loadFolderMail(activeFolder.current)
        await loadFolders()
        store.load(true)
      }, [loadFolderMail, loadFolders])

      const selectMailFolder = (key) => {
        if (key !== folder) {
          setFolder(key)
          setFolderMail(null)
        }
        setSelected(null)
        setFocus(key === 'inbox' ? 'focused' : 'all')
        setCategory('all')
        if (key === 'deleteditems') setFilter('all')
      }

      const deleteMail = React.useCallback(async (row) => {
        if (!row) return
        setBusy(true)
        try {
          await api('/mail/delete', { method: 'POST', body: { id: row.id, folder } })
          await afterMailAction()
        } catch (failure) {
          setError(failure.message)
        } finally {
          setBusy(false)
        }
      }, [afterMailAction, folder])

      const archiveMail = React.useCallback(async (row) => {
        if (!row) return
        setBusy(true)
        try {
          await api('/mail/move', { method: 'POST', body: { id: row.id, destination: 'archive' } })
          await afterMailAction()
        } catch (failure) {
          setError(failure.message)
        } finally {
          setBusy(false)
        }
      }, [afterMailAction])

      const toggleRead = React.useCallback(async (row) => {
        if (!row) return
        setBusy(true)
        try {
          await api('/mail/flag', { method: 'POST', body: { id: row.id, read: !row.isRead } })
          await afterMailAction()
        } catch (failure) {
          setError(failure.message)
        } finally {
          setBusy(false)
        }
      }, [afterMailAction])

      const markAllRead = React.useCallback(async () => {
        setBusy(true)
        setError(null)
        setReadAllFeedback(null)
        try {
          const result = await api('/mail/read-all', { method: 'POST', body: { folder } })
          const rows = await loadFolderMail(folder)
          if (activeFolder.current === folder) {
            setSelected((current) => {
              const refreshed = rows?.find((row) => row.id === current?.id)
              return refreshed ? { ...current, isRead: refreshed.isRead } : current
            })
          }
          await loadFolders()
          if (store.inflight) await store.inflight
          await store.load(true)
          if (result.unconfirmed) {
            setReadAllFeedback({
              folder,
              error: true,
              message: `已标记 ${result.marked} 封。${result.unconfirmed} 封状态未确认，可以再次点击一键已读。`,
            })
          } else {
            setReadAllFeedback({
              folder,
              message: result.marked ? `已将当前文件夹的 ${result.marked} 封邮件标为已读。` : '当前文件夹没有未读邮件。',
            })
          }
        } catch (failure) {
          const message = /^unknown route \/mail\/read-all$/i.test(failure.message)
            ? '请完全退出并重新打开 DeepSeek 助理，再使用一键已读。'
            : failure.message
          setReadAllFeedback({ folder, message, error: true })
        } finally {
          setBusy(false)
        }
      }, [folder, loadFolderMail, loadFolders])

      const sendMail = React.useCallback(async () => {
        setBusy(true)
        setError(null)
        try {
          await api('/mail/send', { method: 'POST', body: compose })
          setCompose(null)
          setNotice(t('mail.sent'))
        } catch (failure) {
          setError(failure.message)
        } finally {
          setBusy(false)
        }
      }, [compose, t])

      const sendReply = React.useCallback(
        async (row, text) => {
          setBusy(true)
          setError(null)
          try {
            await api('/mail/reply', { method: 'POST', body: { id: row.id, comment: text } })
            setNotice(t('reply.sent'))
          } catch (failure) {
            throw failure
          } finally {
            setBusy(false)
          }
        },
        [t],
      )

      const surfaceOf = (key) => {
        const row = appearance?.[key]
        if (!row) return undefined
        const opacity = Number.isFinite(Number(row.opacity)) ? Number(row.opacity) : 100
        const blur = Number.isFinite(Number(row.blur)) ? Number(row.blur) : 0
        const style = {}
        const isPanel = appearanceSharedLabels[key] === 'panel'
        style[isPanel ? '--ob-panel-opacity' : '--ob-region-strength'] = isPanel ? `${opacity}%` : opacity / 100
        if (blur > 0) {
          style[isPanel ? '--ob-panel-filter' : '--ob-region-filter'] = `blur(${blur}px)`
        }
        return style
      }

      const saveAppearance = React.useCallback(async () => {
        setBusy(true)
        try {
          await api('/settings', { method: 'POST', body: { appearance } })
        } finally {
          setBusy(false)
        }
      }, [appearance])

      const resetAppearance = React.useCallback(() => {
        setAppearance(Object.fromEntries(Object.values(appearanceGroups).flat().map((key) =>
          [key, { opacity: 100, blur: 0 }])))
      }, [])

      const addTask = React.useCallback(async (list, title, dueDate) => {
        setBusy(true)
        setError(null)
        setNotice(null)
        try {
          const { task } = await api('/task/create', { method: 'POST', body: { listId: list.id, title, dueDate } })
          await store.inflight
          const tasks = store.snapshot?.tasks ?? []
          if (!tasks.some((row) => row.id === task.id && row.listId === list.id)) {
            store.patch({ tasks: [{ ...task, listName: list.name }, ...tasks] })
          }
          return true
        } catch (failure) {
          if (failure.status === 404) {
            await store.inflight
            await store.load(true)
            setError(t('todo.listMissing'))
          } else {
            setError(failure.message)
          }
          return false
        } finally {
          setBusy(false)
        }
      }, [t])

      const handleTaskError = React.useCallback(async (task, failure) => {
        if (failure.status !== 404) {
          setError(failure.message)
          return
        }
        await store.inflight
        await store.load(true)
        const snap = store.snapshot
        const stillListed = snap?.tasks?.some((item) => item.id === task.id && item.listId === task.listId)
        if (!store.error && !snap?.taskError && !stillListed) setNotice(t('todo.taskGone'))
        else setError(t('todo.taskMissing'))
      }, [t])

      const movePlanTask = React.useCallback(async (task, changes) => {
        setBusy(true)
        setError(null)
        setNotice(null)
        try {
          const result = await api('/task/move', {
            method: 'POST', body: { listId: task.listId, taskId: task.id, ...changes },
          })
          await store.inflight
          store.patch({ tasks: (store.snapshot?.tasks ?? []).map((item) =>
            item.id === task.id && item.listId === task.listId
              ? { ...item, ...result.task, listName: item.listName } : item) })
        } catch (failure) {
          await handleTaskError(task, failure)
        } finally {
          setBusy(false)
        }
      }, [handleTaskError])

      const deletePlanTask = React.useCallback(async (task) => {
        setBusy(true)
        setError(null)
        setNotice(null)
        try {
          await api('/task/delete', { method: 'POST', body: { listId: task.listId, taskId: task.id } })
          await store.inflight
          const tasks = (store.snapshot?.tasks ?? []).filter((item) => item.id !== task.id || item.listId !== task.listId)
          store.patch({ tasks })
        } catch (failure) {
          await handleTaskError(task, failure)
        } finally {
          setBusy(false)
        }
      }, [handleTaskError])

      const toggleTask = React.useCallback(async (task) => {
        setBusy(true)
        setError(null)
        setNotice(null)
        try {
          const completed = !task.completed
          const result = await api('/task/complete', {
            method: 'POST',
            body: { listId: task.listId, taskId: task.id, completed },
          })
          await store.inflight
          store.patch({
            tasks: (store.snapshot?.tasks ?? []).map(
              (item) => item.id === task.id && item.listId === task.listId
                ? { ...item, ...result.task, listName: item.listName, completed,
                    completedAt: completed ? result.task?.completedAt || new Date().toISOString() : null }
                : item,
            ),
          })
        } catch (failure) {
          await handleTaskError(task, failure)
        } finally {
          setBusy(false)
        }
      }, [handleTaskError])


      const submitAiCommand = React.useCallback(async () => {
        const text = String(aiText || '').trim()
        if (!text || aiBusy) return
        const initialSnapshot = store.snapshot
        setAiBusy(true)
        setAiErr(null)
        setAiMsg('处理中…')
        let result = null
        let usedFallback = false
        try {
          result = await api('/ai/command', { method: 'POST', body: { text } })
        } catch (failure) {
          usedFallback = true
          result = parseAiCommandFallback(text)
          if (result.intent === 'unknown') {
            setAiBusy(false)
            setAiMsg(null)
            const raw = (failure && failure.message) || ''
            const hint = /unknown route\s*\/?ai\/command/i.test(raw)
              ? 'AI 路由未加载，请重启 DeepSeek 助理后再试'
              : (raw || result.reply || '无法处理指令')
            setAiErr(hint)
            return
          }
        }
        try {
          const intent = result?.intent
          if (intent === 'calendar' && result.calendar) {
            const cal = { ...result.calendar }
            if (!cal.subject || !cal.date) throw new Error('未能解析出完整日程，请补充标题和日期')
            if (!cal.startTime) cal.startTime = '09:00'
            if (!cal.endTime) cal.endTime = addMinutesHm(cal.startTime, 30)
            if (!cal.endDate) cal.endDate = cal.date
            if (usedFallback || result.fallback) {
              setView('calendar')
              setPendingCreate({ ...cal })
              setAiText('')
              setAiMsg(result.reply || 'AI 暂不可用，已尽量填好表单')
            } else {
              await api('/calendar/create', {
                method: 'POST',
                body: {
                  subject: cal.subject,
                  date: cal.date,
                  endDate: cal.endDate || cal.date,
                  startTime: cal.allDay ? '00:00' : cal.startTime,
                  endTime: cal.allDay ? '00:00' : cal.endTime,
                  allDay: Boolean(cal.allDay),
                  location: cal.location || '',
                  body: cal.body || '',
                },
              })
              setAiText('')
              setAiMsg(result.reply || '日程已创建')
              setView('calendar')
              setCalendarNonce((n) => n + 1)
              try { await store.load(true) } catch { /* ignore */ }
            }
          } else if (intent === 'task' && result.task) {
            const current = store.snapshot
            if (!initialSnapshot?.connected || !current?.connected || current.account?.mail !== initialSnapshot.account?.mail) {
              throw new Error('请连接当前 Outlook 账户后再添加计划')
            }
            const list = current.taskLists?.find((item) => item.isDefault)
            if (!list) throw new Error('找不到默认「任务」清单，请刷新后再试')
            const { title, dueDate } = result.task
            if (!await addTask(list, title, dueDate)) throw new Error('计划未创建，请查看错误提示后重试')
            setAiText('')
            setAiMsg(dueDate ? `计划已创建，截止 ${dueDate}` : '计划已创建')
            setView('todo')
          } else if (intent === 'mail' && result.mail) {
            const mail = result.mail
            setCompose({ to: mail.to || '', subject: mail.subject || '', body: mail.body || '' })
            setAiText('')
            setAiMsg(usedFallback || result.fallback ? '已预填邮件草稿，请核对后发送' : '已打开邮件草稿，请核对后发送')
          } else {
            setAiMsg(null)
            setAiErr(result?.reply || '未能识别意图，请试试添加计划、添加日程或写邮件')
          }
        } catch (failure) {
          setAiMsg(null)
          setAiErr(failure?.message || String(failure))
        } finally {
          setAiBusy(false)
        }
      }, [aiText, aiBusy, addTask])

      const connected = Boolean(snap?.connected)
      const layoutView = view === 'todo' ? 'mail' : view
      const demo = Boolean(snap?.demo)
      const needsConnect = snap !== null && !connected && !demo
      const subtitle = connected
        ? snap?.account?.mail || snap?.account?.displayName || t('subtitle')
        : t('subtitle')

      const body = (() => {
        if (showSettings) {
          return h(SettingsPanel, {
            t,
            draft,
            setDraft,
            onSave: saveSettings,
            onClose: () => setShowSettings(false),
            busy,
          })
        }
        if (needsConnect) {
          return h(ConnectPanel, {
            t,
            builtin: !String(draft.clientId || '').trim(),
            code,
            error,
            busy,
            paste,
            onPasteChange: setPaste,
            onOpenInApp: openAuthorizeInApp,
            onOpenSystem: openAuthorizeSystem,
            onComplete: completeAuth,
            onStart: startConnect,
            onCopy: copyCode,
            onOpen: openSignIn,
            onOpenSettings: openSettings,
          })
        }
        if (!snap) {
          return h('div', { className: 'ob-empty' }, t('loading'))
        }
        return h(
          React.Fragment,
          null,
          error || notice || demo ? h('div', { className: 'ob-body-feedback' },
            error ? h('div', { className: 'ob-error', role: 'alert' }, error) : null,
            notice ? h('div', { className: 'ob-banner', role: 'status' }, notice) : null,
            demo ? h('div', { className: 'ob-banner' },
              h('span', null, t('notice.demo')),
              h('span', { className: 'ob-spacer' }),
              h('button', { type: 'button', className: 'ob-btn primary', onClick: startConnect, disabled: busy },
                t('action.connect'))) : null) : null,
          code
            ? h(ConnectPanel, {
                t,
                builtin: !String(draft.clientId || '').trim(),
                code,
                error,
                busy,
                paste,
                onPasteChange: setPaste,
                onOpenInApp: openAuthorizeInApp,
                onOpenSystem: openAuthorizeSystem,
                onComplete: completeAuth,
                onStart: startConnect,
                onCopy: copyCode,
                onOpen: openSignIn,
                onOpenSettings: openSettings,
              })
            : null,
          view === 'todo'
            ? h(TodoPage, {
                t, lists: snap.taskLists ?? [], tasks: snap.tasks ?? [], error: snap.taskError,
                stale: snap.taskStale, connected, busy, onCreate: addTask, onToggle: toggleTask, onMove: movePlanTask,
                onDelete: deletePlanTask, onRefresh: refresh, completedSort, onCompletedSort: setCompletedSort,
                surfaceOf,
              })
            : view === 'calendar'
            ? h(MonthCalendar, {
                t,
                key: snap.account?.mail || '',
                accountMail: snap.account?.mail,
                tasks: (snap.tasks ?? []).filter((task) => !task.completed),
                onToggle: toggleTask,
                busy,
                toolbarSurface: surfaceOf('calendarToolbar'),
                surface: surfaceOf('calendar'),
                listSurface: surfaceOf('calendarList'),
                pendingCreate,
                onPendingCreateConsumed: () => setPendingCreate(null),
                calendarNonce,
              })
            : h(MailWorkspace, {
                  t, selected, busy, connected, folders, folder, filter, setFilter, focus, setFocus, category, setCategory,
                  account: snap.account?.mail || snap.account?.displayName || 'Outlook',
                  rows: folderMail ?? (folder === 'inbox' ? snap.mail ?? [] : []), unread: Number(snap.unread || 0),
                  onFolder: selectMailFolder,
                  onSelect: (row) => { setSelected(row); if (row && !row.isRead && !busy) markRead(row) },
                  onCompose: () => setCompose({ to: '', subject: '', body: '' }),
                  onDelete: deleteMail, onArchive: archiveMail, onToggleRead: toggleRead,
                  onMarkFolderRead: markAllRead, readAllFeedback,
                  onReplyRow: (row) => { setSelected(row); setReplyRequest({ id: row.id, at: Date.now() }) },
                  replyRequest, onSendReply: sendReply,
                  onRefreshFolder: async (key) => {
                    selectMailFolder(key)
                    if (key === folder) await loadFolderMail(key)
                    await loadFolders()
                  },
                  surface: surfaceOf('mail'),
                  toolbarSurface: surfaceOf('mailToolbar'),
                  folderSurface: surfaceOf('mailFolders'),
                  readingSurface: surfaceOf('mailReading'),
                }),
          (snap.warnings ?? []).length > 0
            ? h('div', { className: 'ob-body-feedback' },
                h('div', { className: 'ob-error', role: 'alert' }, snap.warnings.join('；')))
            : null,
        )
      })()

      return h(
        'div',
        { className: `ob-root${!showSettings && !needsConnect ? ` ob-outlook-${layoutView}` : ''}`,
          style: surfaceOf(`${view}Panel`) },
        h(Sheet, null),
        h(ContextMenu, null),
        h(
          'header',
          { className: 'ob-head', style: surfaceOf(`${view}Header`) },
          h('span', { className: 'ob-head-mark' }, mark(18)),
          h(
            'div',
            { className: 'ob-head-text' },
            h('div', { className: 'ob-title' }, t('title')),
            h('div', { className: 'ob-sub' }, subtitle),
          ),
          Boolean(snap) && !showSettings && !needsConnect
            ? h(
                'div',
                { className: 'ob-tabs', role: 'tablist', 'aria-label': t('title') },
                [
                  ['mail', t('tab.mail')],
                  ['calendar', t('tab.calendar')],
                  ['todo', t('tab.todo')],
                ].map(([key, label]) =>
                  h(
                    'button',
                    {
                      key,
                      type: 'button',
                      role: 'tab',
                      'aria-selected': view === key,
                      className: `ob-tab${view === key ? ' on' : ''}`,
                      onClick: () => setView(key),
                    },
                    label,
                  ),
                ),
              )
            : null,
          h('span', { className: 'ob-spacer' }),
          demo ? h('span', { className: 'ob-chip demo' }, t('chip.demo')) : null,
          h('div', { className: 'ob-guide-entry' },
            connected
              ? h('span', { className: 'ob-chip ok' }, 'Outlook')
              : h('span', { className: 'ob-chip' }, t('chip.offline')),
            h('button', { type: 'button', className: 'ob-btn', 'aria-haspopup': 'dialog', onClick: () => setShowGuide(true) },
              t('guide.open'))),
          h(
            'button',
            {
              type: 'button',
              className: `ob-btn${showUi ? ' primary' : ''}`,
              onClick: () => setShowUi((value) => !value),
            },
            t('tab.appearance'),
          ),
          h(
            'button',
            { type: 'button', className: 'ob-btn', onClick: refresh, disabled: busy },
            t('action.refresh'),
          ),
          h('button', { type: 'button', className: 'ob-btn', onClick: openSettings }, t('action.settings')),
          connected
            ? h('button', { type: 'button', className: 'ob-btn', onClick: startConnect, disabled: busy }, t('action.reauth'))
            : h('button', { type: 'button', className: 'ob-btn primary', onClick: startConnect, disabled: busy }, t('action.connect')),
          connected
            ? h('button', { type: 'button', className: 'ob-btn', onClick: disconnect, disabled: busy }, t('action.disconnect'))
            : null,
        ),

        Boolean(snap) && !showSettings && !needsConnect
          ? h(
              'div',
              { className: 'ob-ai-wrap', style: surfaceOf(`${view}AiBar`) },
              h(
                'div',
                { className: 'ob-ai-bar', role: 'search' },
                h('span', { className: 'ob-ai-bar-mark', 'aria-hidden': 'true' }, 'AI'),
              h('input', {
                className: 'ob-ai-bar-input',
                type: 'text',
                value: aiText,
                disabled: aiBusy,
                placeholder: '试试：添加计划，明天交周报 / 明天下午3点开会 / 给 xxx 写邮件',
                'aria-label': 'AI 对话：添加计划、添加日程或写邮件',
                onChange: (event) => {
                  setAiText(event.target.value)
                  if (aiErr) setAiErr(null)
                  if (aiMsg && aiMsg !== '处理中…') setAiMsg(null)
                },
                onKeyDown: (event) => {
                  if (event.key === 'Enter' && !event.shiftKey && !event.nativeEvent?.isComposing && event.keyCode !== 229) {
                    event.preventDefault()
                    submitAiCommand()
                  }
                },
              }),
              aiBusy
                ? h('span', { className: 'ob-ai-bar-msg busy' }, '处理中…')
                : aiErr
                  ? h('span', { className: 'ob-ai-bar-msg err', title: aiErr }, aiErr)
                  : aiMsg
                    ? h('span', { className: 'ob-ai-bar-msg', title: aiMsg }, aiMsg)
                    : null,
                h(
                  'button',
                  {
                    type: 'button',
                    className: 'ob-ai-bar-send',
                    disabled: aiBusy || !String(aiText || '').trim(),
                    onClick: submitAiCommand,
                  },
                  '发送',
                ),
              ),
            )
          : null,

        h('div', { className: 'ob-body' }, body),
        compose
          ? h(
              'div',
              {
                className: 'ob-modal-mask',
                onClick: (event) => {
                  if (event.target === event.currentTarget) setCompose(null)
                },
              },
              h(
                'div',
                { className: 'ob-modal ob-modal-md' },
                h(
                  'div',
                  { className: 'ob-modal-head' },
                  h('span', null, t('mail.compose')),
                  h('button', { type: 'button', className: 'ob-btn', onClick: () => setCompose(null) }, t('action.close')),
                ),
                h(
                  'div',
                  { className: 'ob-form' },
                  h('input', {
                    className: 'ob-input',
                    placeholder: t('mail.to'),
                    value: compose.to,
                    autoFocus: true,
                    onChange: (event) => setCompose({ ...compose, to: event.target.value }),
                  }),
                  h('input', {
                    className: 'ob-input',
                    placeholder: t('reply.subject'),
                    value: compose.subject,
                    onChange: (event) => setCompose({ ...compose, subject: event.target.value }),
                  }),
                  h('textarea', {
                    className: 'ob-textarea',
                    rows: 8,
                    placeholder: t('mail.body'),
                    value: compose.body,
                    onChange: (event) => setCompose({ ...compose, body: event.target.value }),
                  }),
                ),
                h(
                  'div',
                  { className: 'ob-ui-foot' },
                  h(
                    'button',
                    {
                      type: 'button',
                      className: 'ob-btn primary',
                      disabled: busy || !String(compose.to || '').trim(),
                      onClick: sendMail,
                    },
                    t('mail.send'),
                  ),
                ),
              ),
            )
          : null,
        showGuide ? h(SetupGuide, {
          t, version: snap?.pluginVersion, connected, busy: busy || showSettings, onOpenEntra: openInApp,
          onSettings: () => { setShowGuide(false); if (!showSettings) openSettings() },
          onConnect: () => { setShowGuide(false); startConnect() },
          onClose: () => setShowGuide(false),
        }) : null,
        showUi
          ? h(AppearancePage, {
              t,
              view,
              onView: setView,
              appearance,
              setAppearance,
              onSave: saveAppearance,
              onReset: resetAppearance,
              onClose: () => setShowUi(false),
              busy,
            })
          : null,
      )
    }

    /* ------------------------------------------------------------ activation */

    return {
      inject: ['slots', 'locale'],
      apply(ctx) {
        // The dictionary service may still be missing on an early activation
        // pass: read it through `get` and fall back to the packaged dictionary
        // instead of throwing, or a boot-order race takes the whole page down.
        let bound = null
        try {
          const locale = ctx.get('locale')
          if (locale) {
            ctx.effect(() => locale.register(NS, DICT), 'outlook-bridge: dictionaries')
            bound = locale.bind(NS)
          }
        } catch (error) {
          console.error('[outlook-bridge] dictionary registration failed:', error)
        }

        const translate = (key, params) => {
          if (bound) {
            const text = bound(key)
            if (text && text !== key) {
              if (!params || typeof params !== 'object') return text
              return Object.entries(params).reduce(
                (acc, [name, value]) => acc.split(`{${name}}`).join(String(value)),
                text,
              )
            }
          }
          return fbT(key, params)
        }

        try {
          // The desktop shell owns an embedded browser whose <webview> can render
        // Microsoft's sign-in page; the page's own window.open cannot.
        const openInApp = (url) => {
          try {
            const sidebarRight = ctx.get('sidebarRight')
            if (sidebarRight && typeof sidebarRight.openTab === 'function') {
              sidebarRight.openTab('browser', { params: { url } })
              return true
            }
          } catch (error) {
            console.error('[outlook-bridge] in-app browser unavailable:', error)
          }
          return false
        }

        ctx.effect(() => () => store.dispose(), 'outlook-bridge: snapshot store')

          ctx.slots.inject('sidebar.panellist', () =>
            ctx.slots.register(
              {
                name: 'sidebar.panellist',
                id: PANEL_ID,
                order: 21,
                label: () => translate('panel'),
              },
              PanelGlyph,
            ),
          )

          ctx.slots.inject('main', () =>
            ctx.slots.register(
              {
                name: 'main',
                key: PANEL_ID,
                inject: () => ({ t: translate, openInApp }),
              },
              MainPage,
            ),
          )
        } catch (error) {
          console.error('[outlook-bridge] slot registration failed:', error)
        }
      },
    }
  },
})
