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
        'todo.title': '待办',
        'todo.empty': '没有未完成的待办',
        'todo.due': '截止 {time}',
        'todo.unsupported': '个人 Microsoft 账户不支持微软待办（To Do）接口；换成工作/学校账户即可显示。',
        'settings.title': '连接设置',
        'settings.clientId': 'Microsoft 应用（客户端）ID（可选）',
        'settings.clientSecret': '客户端密码（可选，机密客户端用）',
        'settings.clientSecretHint': '如果应用里找不到「允许公共客户端流」，就改走机密客户端：左侧「证书和密码」→ 新客户端密码 → 添加 → 立刻复制那串「值」，粘到这里。留空表示不修改。',
        'settings.clientIdHint': '微软禁止内置（第一方）客户端读取个人账户邮件，会直接报 AADSTS65002。免费注册一个自己的应用即可：portal.azure.com → 应用注册 → 新注册（账户类型选「个人 Microsoft 账户」）→ 复制「应用程序(客户端) ID」→ 在「身份验证」里勾选「允许公共客户端流」。把 ID 填到这里保存。',
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
        'tab.reply': '回复',
        'tab.appearance': '界面',
        'ui.panel': '主面板',
        'ui.mail': '邮件列表',
        'ui.calendar': '日历',
        'ui.todo': '待办',
        'ui.reply': '回复整理',
        'ui.opacity': '透明度',
        'ui.blur': '模糊度',
        'ui.title': '界面设置',
        'ui.hint': '按区域单独调节透明度与模糊度，立刻生效；确认后点保存。',
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
        'todo.title': 'To-do',
        'todo.empty': 'Nothing outstanding',
        'todo.due': 'Due {time}',
        'todo.unsupported': 'Personal Microsoft accounts have no To Do API; sign in with a work or school account to see tasks.',
        'settings.title': 'Connection settings',
        'settings.clientId': 'Microsoft app (client) ID - optional',
        'settings.clientSecret': 'Client secret (optional, confidential client)',
        'settings.clientSecretHint': 'If the app has no "allow public client flows" switch, use a confidential client instead: Certificates & secrets -> New client secret -> copy the Value -> paste it here. Leave empty to keep the stored one.',
        'settings.clientIdHint': 'Microsoft blocks its built-in (first-party) client from reading personal-account mail with AADSTS65002. Register a free app: portal.azure.com -> App registrations -> New registration (personal Microsoft accounts) -> copy the Application (client) ID -> Authentication -> allow public client flows. Paste the ID here.',
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
        'tab.reply': 'Reply',
        'tab.appearance': 'Interface',
        'ui.panel': 'Main panel',
        'ui.mail': 'Mail list',
        'ui.calendar': 'Calendar',
        'ui.todo': 'To-do',
        'ui.reply': 'Reply',
        'ui.opacity': 'Opacity',
        'ui.blur': 'Blur',
        'ui.title': 'Interface settings',
        'ui.hint': 'Tune opacity and blur per region; changes apply immediately, press Save to keep them.',
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
      const res = await fetch(`/outlook/api${path}`, init)
      let data = null
      try {
        data = await res.json()
      } catch {
        data = null
      }
      if (!res.ok) throw new Error(data?.error || `HTTP ${res.status}`)
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

    function fmtDue(date, t) {
      if (!date) return ''
      const today = startOfDay(new Date())
      const target = startOfDay(date)
      const diff = Math.round((target.getTime() - today.getTime()) / 86_400_000)
      const time = fmtTime(date)
      if (diff === 0) return `${t('calendar.today')} ${time}`
      if (diff === 1) return `${t('calendar.tomorrow')} ${time}`
      return `${date.getMonth() + 1}/${date.getDate()} ${time}`
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
.ob-root{display:flex;flex-direction:column;height:100%;min-height:0;color:var(--dsw-alias-label-primary);font-size:13px;line-height:1.45}
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
.ob-todo{display:flex;gap:9px;align-items:flex-start;padding:8px 12px;border-bottom:1px solid var(--dsw-alias-border-l1)}
.ob-todo:last-child{border-bottom:0}
.ob-check{width:15px;height:15px;border-radius:50%;border:1.5px solid var(--dsw-alias-border-l2);background:transparent;flex:none;margin-top:1px;cursor:pointer;padding:0;display:flex;align-items:center;justify-content:center}
.ob-check:hover{border-color:var(--dsw-alias-brand-primary)}
.ob-check.done{background:var(--dsw-alias-brand-primary);border-color:var(--dsw-alias-brand-primary)}
.ob-todo-main{min-width:0;flex:1}
.ob-todo-title{font-size:12.5px;word-break:break-word}
.ob-todo-title.done{text-decoration:line-through;color:var(--dsw-alias-label-tertiary)}
.ob-todo-meta{font-size:10.5px;color:var(--dsw-alias-label-quaternary);margin-top:1px}
.ob-tabs{display:flex;gap:4px;padding:6px 0 0;flex:none}
.ob-tab{height:26px;padding:0 12px;border:0;border-radius:8px 8px 0 0;background:transparent;color:var(--dsw-alias-label-tertiary);font-size:12px;cursor:pointer}
.ob-tab:hover{background:var(--dsw-alias-interactive-bg-hover);color:var(--dsw-alias-label-secondary)}
.ob-tab.on{background:var(--dsw-alias-bg-layer-1);color:var(--dsw-alias-label-primary);font-weight:600}
.ob-grid-calendar{grid-template-columns:minmax(0,1.5fr) minmax(0,1fr)}
.ob-modal-mask{position:fixed;inset:0;z-index:60;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,0.28);backdrop-filter:blur(2px)}
.ob-modal{width:min(760px,92vw);max-height:82vh;display:flex;flex-direction:column;gap:10px;padding:14px;border-radius:14px;border:1px solid var(--dsw-alias-border-l1);background:var(--dsw-alias-bg-overlay);box-shadow:0 18px 48px rgba(0,0,0,0.28)}
.ob-modal-head{display:flex;align-items:center;justify-content:space-between;font-size:13px;font-weight:600;flex:none}
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
.ob-ui-page{display:flex;flex-direction:column;gap:12px;min-height:0;overflow:auto}
.ob-ui-head{font-size:11.5px;color:var(--dsw-alias-label-tertiary)}
.ob-surface-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:10px}
.ob-surface-card{display:flex;flex-direction:column;gap:10px;padding:12px;border:1px solid var(--dsw-alias-border-l1);border-radius:12px;background:var(--dsw-alias-bg-layer-1)}
.ob-surface-title{display:flex;align-items:center;justify-content:space-between;font-size:12.5px;font-weight:600}
.ob-surface-badge{font-size:11px;font-weight:500;color:var(--dsw-alias-label-secondary);background:var(--dsw-alias-bg-layer-2);border-radius:6px;padding:1px 6px}
.ob-slider{display:grid;grid-template-columns:52px 1fr 40px;align-items:center;gap:8px;font-size:11.5px;color:var(--dsw-alias-label-secondary)}
.ob-slider input[type=range]{width:100%;accent-color:var(--dsw-alias-brand-primary)}
.ob-slider b{font-weight:500;text-align:right;color:var(--dsw-alias-label-tertiary)}
.ob-ui-foot{display:flex;gap:8px}
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
:is(.ob-outlook-mail,.ob-outlook-calendar){--ob-blue:#0f6cbd;--ob-line:var(--dsw-alias-border-l1,#dedede);--ob-paper:var(--dsw-alias-bg-layer-1,#fff);--ob-ground:var(--dsw-alias-bg-layer-2,#f5f5f5);font-family:"Segoe UI","Microsoft YaHei",sans-serif;container-type:inline-size}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-head{padding:8px 16px;gap:8px;background:#f3f3f3;border-bottom-color:var(--ob-line,#dedede)}
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
.ob-outlook-mail .ob-body>.ob-banner,.ob-outlook-mail .ob-body>.ob-error,.ob-outlook-mail .ob-body>.ob-connect{margin:8px 12px}
.ob-mail-workspace{display:flex;flex:1;flex-direction:column;min-height:0;min-width:0;background:var(--ob-ground)}
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
.ob-calendar-workspace{display:flex;flex:1;flex-direction:column;min-height:0;background:var(--ob-ground);overflow:hidden}
.ob-calendar-toolbar{min-height:43px;gap:5px;flex:none}
.ob-calendar-view-label{display:inline-flex;align-items:center;gap:6px;height:32px;padding:0 9px;border:1px solid var(--ob-line);border-radius:3px;background:var(--ob-ground);font-size:13px}
.ob-calendar-toolbar .ob-command.active{color:var(--ob-blue)}
.ob-calendar-hint{font-size:11px;color:var(--dsw-alias-label-tertiary);padding-right:10px;white-space:nowrap}
.ob-calendar-layout{display:grid;grid-template-columns:232px minmax(0,1fr);flex:1;min-height:0;gap:0;padding-top:7px}
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
.ob-calendar-board{display:flex;flex-direction:column;min-width:0;min-height:0;margin-right:8px;border:1px solid var(--ob-line);border-bottom:0;border-radius:5px 5px 0 0;background:var(--ob-paper);overflow:hidden}
.ob-calendar-month-head{display:flex;align-items:center;gap:8px;height:48px;flex:none;padding:0 12px}
.ob-calendar-today{border:1px solid var(--ob-line);border-radius:4px;min-width:90px;gap:14px;min-height:31px;font-weight:600}
.ob-calendar-nav{width:28px;height:30px;background:transparent;border:0;color:var(--dsw-alias-label-secondary);font-size:21px;cursor:pointer}
.ob-calendar-month-picker{position:relative;display:flex;align-items:center;gap:9px;font-size:19px;font-weight:600;margin-left:6px;cursor:pointer;white-space:nowrap}
.ob-calendar-month-picker input{position:absolute;inset:0;width:100%;height:100%;opacity:0;cursor:pointer}
.ob-calendar-month-picker input::-webkit-calendar-picker-indicator{position:absolute;inset:0;width:100%;height:100%;cursor:pointer}
.ob-calendar-month-picker:focus-within{outline:2px solid var(--ob-blue);outline-offset:3px}
.ob-calendar-loading{font-size:12px;color:var(--dsw-alias-label-tertiary)}
.ob-calendar-weekdays{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));flex:none;border-top:1px solid var(--ob-line);border-bottom:1px solid var(--ob-line);margin-right:14px}
.ob-calendar-weekdays span{font-size:11px;line-height:27px;border-right:1px solid var(--ob-line);padding:0 8px}
.ob-calendar-scroll{display:flex;flex-direction:column;flex:1;min-height:0;overflow:auto;scrollbar-width:thin}
.ob-calendar-scroll .ob-calendar-weekdays{position:sticky;top:0;z-index:3;background:var(--ob-paper)}
.ob-calendar-month-grid{flex:1;min-height:648px;display:grid;grid-template-columns:repeat(7,minmax(0,1fr));grid-template-rows:repeat(6,minmax(108px,1fr));margin-right:14px}
.ob-calendar-day{min-width:0;min-height:0;position:relative;display:flex;flex-direction:column;border-right:1px solid var(--ob-line);border-bottom:1px solid var(--ob-line);padding:3px 0 5px;background:transparent;cursor:pointer}
.ob-calendar-day.outside{background:color-mix(in srgb,currentColor 2%,transparent)}
.ob-calendar-day:hover{background:color-mix(in srgb,var(--ob-blue) 3%,transparent)}
.ob-calendar-day.selected{outline:1px solid var(--ob-blue);outline-offset:-1px;z-index:1}
.ob-calendar-date{display:flex;align-items:center;align-self:flex-start;min-height:25px;padding:0 8px;background:transparent;border:0;color:var(--dsw-alias-label-secondary);font:inherit;font-size:13px;cursor:pointer;flex:none}
.ob-calendar-date .today-dot{display:flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;background:var(--ob-blue);color:#fff;font-weight:600;margin-left:-4px}
.ob-calendar-events{display:flex;flex-direction:column;gap:2px;padding-top:2px;min-height:0;overflow:hidden}
.ob-calendar-event{display:block;text-align:left;flex:none;margin:0 7px 0 0;max-width:100%;min-height:22px;line-height:20px;border:0;border-left:4px solid;border-radius:3px;padding:0 6px;font:inherit;font-size:11px;color:var(--dsw-alias-label-secondary);overflow:hidden;white-space:nowrap;text-overflow:ellipsis;cursor:pointer}
.ob-calendar-event:hover{filter:brightness(.96)}
.ob-calendar-event[draggable=true]{cursor:grab}.ob-calendar-event[draggable=true]:active{cursor:grabbing}.ob-calendar-event.dragging{opacity:.38}.ob-calendar-event.drag-preview{opacity:.5;pointer-events:none;cursor:default}.ob-calendar-day.drag-over{background:color-mix(in srgb,var(--ob-blue) 10%,transparent)}.ob-calendar-day.drag-over .ob-calendar-events{min-height:24px}
.ob-calendar-more{border:0;background:transparent;text-align:left;color:var(--ob-blue);font:inherit;font-size:11px;cursor:pointer;padding:1px 8px;flex:none}
.ob-calendar-status{flex:none;padding:5px 12px;font-size:12px;border-top:1px solid var(--ob-line);color:var(--dsw-alias-label-secondary)}
.ob-calendar-status.error{color:var(--dsw-alias-state-error-primary,#b42318)}
.ob-calendar-agenda{width:min(440px,90vw);overflow:auto}.ob-calendar-agenda-item{display:flex;flex-direction:column;gap:6px;border-left:3px solid;padding:8px 12px;margin:4px 0;font-size:12px}.ob-calendar-agenda-item strong{font-size:14px}.ob-calendar-agenda-item a{color:var(--ob-blue)}
.ob-event-peek{position:fixed;z-index:1100;width:min(320px,calc(100vw - 16px));background:var(--ob-paper,#fff);color:var(--dsw-alias-label-primary,#242424);border:1px solid var(--ob-line,#ddd);border-radius:8px;box-shadow:0 8px 28px #0000002e;overflow:hidden;font-family:"Segoe UI","Microsoft YaHei",sans-serif}.ob-event-peek-accent{height:4px;background:#13a10e}.ob-event-peek-body{padding:12px 14px 8px;display:flex;flex-direction:column;gap:8px}.ob-event-peek-cat{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--dsw-alias-label-secondary,#555)}.ob-event-peek-dot{width:10px;height:10px;border-radius:2px;flex:none}.ob-event-peek-title{margin:0;font-size:18px;line-height:1.3;font-weight:600}.ob-event-peek-row{display:flex;align-items:flex-start;gap:8px;font-size:12.5px;color:var(--dsw-alias-label-secondary,#555)}.ob-event-peek-row svg{flex:none;margin-top:1px}.ob-event-peek-actions{display:flex;align-items:center;gap:4px;padding:6px 8px 10px;border-top:1px solid var(--ob-line,#eee)}.ob-event-peek-actions .ob-command{min-height:30px;padding:0 10px;font-size:12.5px}.ob-event-peek-expand{margin-left:auto}.ob-event-peek-extra{padding:0 14px 12px;font-size:12px;color:var(--dsw-alias-label-tertiary,#666);white-space:pre-wrap}.ob-day-peek{position:fixed;z-index:1100;width:min(300px,calc(100vw - 16px));max-height:min(360px,70vh);overflow:auto;background:var(--ob-paper,#fff);border:1px solid var(--ob-line,#ddd);border-radius:8px;box-shadow:0 8px 28px #0000002e;padding:8px}.ob-day-peek-head{display:flex;align-items:center;justify-content:space-between;padding:4px 6px 8px;font-weight:600;font-size:13px}.ob-day-peek-item{display:flex;flex-direction:column;gap:2px;width:100%;text-align:left;border:0;border-left:3px solid;background:transparent;padding:8px 10px;margin:2px 0;border-radius:4px;cursor:pointer;font:inherit}.ob-day-peek-item:hover{background:var(--ob-ground,#f5f5f5)}.ob-day-peek-item strong{font-size:13px}.ob-day-peek-item span{font-size:11.5px;color:var(--dsw-alias-label-tertiary,#666)}
.ob-outlook-calendar button:focus-visible,.ob-outlook-calendar a:focus-visible,.ob-outlook-calendar input:focus-visible{outline:2px solid var(--ob-blue);outline-offset:-2px}
.ob-outlook-calendar .ob-btn.primary{background:var(--ob-blue)}
.ob-outlook-calendar .ob-body>.ob-banner,.ob-outlook-calendar .ob-body>.ob-error{margin:5px 10px}
@container (max-width:1000px){.ob-calendar-layout{grid-template-columns:200px minmax(0,1fr)}.ob-calendar-sidebar{padding:8px}.ob-calendar-hint{display:none}.ob-calendar-month-picker{font-size:17px}.ob-calendar-choice{padding-left:26px}}
@container (max-width:700px){.ob-calendar-layout{grid-template-columns:166px minmax(0,1fr)}.ob-calendar-mini-head{font-size:12px}.ob-calendar-board{overflow-x:auto}.ob-calendar-month-head,.ob-calendar-weekdays,.ob-calendar-month-grid{min-width:570px}.ob-calendar-sidebar{padding:6px}.ob-mini-day{height:25px}.ob-calendar-choice{padding-left:15px;font-size:12px}.ob-calendar-month-head{gap:4px}.ob-calendar-toolbar{gap:0}}
.ob-context-menu{position:fixed;z-index:1200;min-width:214px;max-width:calc(100vw - 12px);max-height:calc(100vh - 12px);overflow:auto;padding:5px;background:var(--dsw-alias-bg-overlay,#fff);color:var(--dsw-alias-label-primary,#242424);border:1px solid var(--dsw-alias-border-l1,#ddd);border-radius:7px;box-shadow:0 8px 28px #0003;font-family:"Segoe UI","Microsoft YaHei",sans-serif}
.ob-context-menu button{display:flex;align-items:center;gap:10px;min-height:34px;width:100%;padding:6px 12px;border:0;border-radius:3px;background:transparent;color:inherit;text-align:left;font:inherit;font-size:13px;cursor:pointer}.ob-context-menu button:hover,.ob-context-menu button:focus-visible{background:var(--dsw-alias-bg-layer-2,#f1f1f1);outline:0}.ob-context-menu button:disabled{opacity:.4;cursor:default}.ob-context-menu button.danger{color:#bd2727}.ob-context-separator{height:1px;background:var(--dsw-alias-border-l1,#ddd);margin:4px 7px}
.ob-context-toast{position:fixed;z-index:1300;bottom:22px;left:50%;transform:translateX(-50%);padding:12px 18px;background:var(--dsw-alias-bg-overlay,#fff);border:1px solid var(--dsw-alias-border-l1,#ddd);box-shadow:0 4px 20px #0003;border-radius:6px;max-width:80vw}.ob-context-toast button{border:0;background:transparent;color:inherit;margin-left:14px;cursor:pointer}
.ob-event-mask{position:fixed;inset:0;z-index:100;display:flex;align-items:center;justify-content:center;padding:12px;background:#0005;backdrop-filter:blur(2px)}
.ob-event-editor{--ob-blue:#0f6cbd;--ob-paper:var(--dsw-alias-bg-layer-1,#fff);--ob-line:var(--dsw-alias-border-l1,#ddd);--ob-ground:var(--dsw-alias-bg-layer-2,#f5f5f5);display:flex;flex-direction:column;width:min(1180px,98vw);height:min(850px,96vh);margin:0;background:var(--ob-ground);color:var(--dsw-alias-label-primary,#242424);border:1px solid var(--ob-line);border-radius:6px;box-shadow:0 12px 60px #0004;overflow:hidden;font-family:"Segoe UI","Microsoft YaHei",sans-serif}
.ob-event-titlebar{display:flex;align-items:center;justify-content:space-between;min-height:40px;padding:0 16px;font-size:13px;background:var(--ob-ground);flex:none}.ob-event-titlebar button{font-size:24px;border:0;background:transparent;color:inherit;cursor:pointer;width:32px;height:32px}.ob-event-titlebar button:hover{background:#e15b5b;color:white}
.ob-event-toolbar{display:flex;align-items:center;gap:12px;min-height:42px;flex:none;margin:6px 14px 8px;padding:4px 8px;border:1px solid var(--ob-line);border-radius:8px;background:var(--ob-paper);box-shadow:0 1px 3px #0001;overflow-x:auto}.ob-event-toolbar-label{display:flex;align-items:center;gap:7px;font-size:13px;white-space:nowrap}.ob-event-toolbar select{border:0;background:transparent;font:inherit;color:inherit;padding:4px;max-width:150px}
.ob-event-layout{display:grid;grid-template-columns:minmax(0,2.1fr) minmax(270px,1fr);gap:14px;padding:0 14px 14px;flex:1;min-height:0}.ob-event-left{display:flex;flex-direction:column;gap:8px;min-width:0;overflow:auto}.ob-event-fields{border:1px solid var(--ob-line);border-radius:6px;background:var(--ob-paper);padding:10px 14px;margin:0;min-width:0;flex:none}.ob-event-field{display:flex;gap:16px;align-items:flex-start;min-height:46px}.ob-event-field>svg{flex:none;margin:14px 0 0 4px;color:var(--dsw-alias-label-secondary,#555)}.ob-event-field-content{flex:1;min-width:0;border-bottom:1px solid #8887;padding:4px 0}.ob-event-field:last-child .ob-event-field-content{border:0}
.ob-event-field-content>input:not([type=checkbox]),.ob-event-field-content>select{width:100%;height:37px;padding:4px 10px;border:0;border-radius:0;background:transparent;color:inherit;font:inherit;font-size:13px;box-sizing:border-box}.ob-event-subject .ob-event-field-content{border-bottom:2px solid var(--ob-blue)}.ob-event-subject input{font-size:17px!important}.ob-event-fields input:focus,.ob-event-fields select:focus{outline:1px solid var(--ob-blue);outline-offset:-1px}.ob-event-dates{display:flex;flex-direction:column;gap:8px;padding:6px 4px 4px}.ob-event-date-row{display:grid;grid-template-columns:40px minmax(0,1fr) auto;align-items:center;gap:8px}.ob-event-date-row>span{font-size:12px;color:var(--dsw-alias-label-secondary,#555);font-weight:600}.ob-event-date-inputs{display:flex;align-items:center;gap:8px;min-width:0;flex-wrap:wrap}.ob-event-date-inputs input[type=date],.ob-event-date-inputs input[type=time]{height:32px;padding:0 8px;border:1px solid var(--ob-line,#ddd);border-radius:6px;background:var(--ob-paper,#fff);color:inherit;font:inherit;font-size:13px;min-width:0}.ob-event-date-inputs input[type=date]{width:148px}.ob-event-date-inputs input[type=time]{width:110px}.ob-event-duration{display:flex;align-items:center;gap:6px;flex-wrap:wrap;padding:2px 0 0 48px}.ob-event-duration-label{font-size:11px;color:var(--dsw-alias-label-tertiary,#777);margin-right:2px}.ob-event-duration-btn{height:28px;padding:0 10px;border:1px solid var(--ob-line,#ddd);border-radius:999px;background:transparent;color:var(--dsw-alias-label-secondary,#555);font:inherit;font-size:12px;cursor:pointer}.ob-event-duration-btn:hover:not(:disabled){background:color-mix(in srgb,var(--ob-blue) 6%,transparent);border-color:color-mix(in srgb,var(--ob-blue) 35%,var(--ob-line,#ddd))}.ob-event-duration-btn.on{background:color-mix(in srgb,var(--ob-blue) 14%,transparent);border-color:var(--ob-blue);color:var(--ob-blue);font-weight:600}.ob-event-duration-btn:disabled{opacity:.4;cursor:default}.ob-event-date-options{display:flex;align-items:center;gap:14px;padding:4px 0 2px 48px;font-size:11px;color:var(--dsw-alias-label-tertiary,#666)}.ob-event-date-options label{display:flex;align-items:center;gap:6px}.ob-event-fields input[type=checkbox]{accent-color:var(--ob-blue)}
.ob-event-body-card{display:flex;align-items:flex-start;gap:16px;padding:12px 14px;flex:1;min-height:160px;border:1px solid var(--ob-line);background:var(--ob-paper);border-radius:6px}.ob-event-body-card>svg{margin:10px 0 0 4px;flex:none;color:var(--dsw-alias-label-secondary,#555)}.ob-event-body-card textarea{align-self:stretch;flex:1;min-width:0;border:1px solid var(--ob-line);background:var(--ob-paper);color:inherit;border-radius:3px;resize:none;padding:12px;font:inherit;font-size:14px;line-height:1.7}
.ob-event-schedule{display:flex;flex-direction:column;min-height:0;border:1px solid var(--ob-line);border-radius:6px;background:var(--ob-paper);overflow:hidden}.ob-event-schedule-head{display:flex;align-items:center;gap:8px;padding:10px 8px;min-height:48px;border-bottom:1px solid var(--ob-line);flex:none}.ob-event-schedule-head strong{font-size:13px}.ob-event-schedule-head button{border:0;background:transparent;color:inherit;cursor:pointer;font-size:20px;padding:0 4px}.ob-event-timeline-status{min-height:24px;padding:4px 8px;border-bottom:1px solid var(--ob-line);font-size:10px;color:var(--dsw-alias-label-tertiary,#777);flex:none}.ob-event-timeline-scroll{overflow:auto;flex:1;min-height:0;scrollbar-width:thin}.ob-event-timeline{height:1152px;position:relative;margin-right:12px}.ob-event-hour{display:block;position:relative;width:100%;height:48px;border:0;border-bottom:1px solid var(--ob-line);background:transparent;text-align:left;padding:0;color:var(--dsw-alias-label-tertiary,#666);font-size:11px;cursor:pointer}.ob-event-hour:after{content:"";position:absolute;left:44px;right:0;top:24px;border-top:1px dashed var(--ob-line)}.ob-event-hour span{display:block;width:43px;border-right:1px solid var(--ob-line);height:100%;padding:4px 0 0 17px;box-sizing:border-box}.ob-event-hour:hover{background:color-mix(in srgb,var(--ob-blue) 5%,transparent)}
.ob-event-existing,.ob-event-draft-block,.ob-event-drag-ghost{position:absolute;left:44px;right:0;padding:3px 8px;box-sizing:border-box;border-left:4px solid;border-radius:3px;overflow:hidden;font-size:11px}.ob-event-existing{background:color-mix(in srgb,var(--ob-blue) 12%,var(--ob-paper));opacity:.8;right:40%;color:inherit;pointer-events:none}.ob-event-existing.interactive{pointer-events:auto;cursor:grab;touch-action:none}.ob-event-existing.interactive:active,.ob-event-draft-block:active{cursor:grabbing}.ob-event-existing.drag-origin{opacity:.32}.ob-event-draft-block{left:55%;z-index:2;background:#48c274;color:#0b2616;border-left-color:#249a4e;box-shadow:0 3px 8px #0002;display:flex;flex-direction:column;gap:3px;pointer-events:auto;cursor:grab;touch-action:none}.ob-event-draft-block.dragging{opacity:.9;z-index:5;box-shadow:0 6px 14px #0003}.ob-event-draft-block strong{font-size:12px;white-space:nowrap}.ob-event-draft-block span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.ob-event-drag-ghost{z-index:4;pointer-events:none;opacity:.5;display:flex;flex-direction:column;gap:3px;box-shadow:0 4px 10px #0002}.ob-event-drag-ghost.draft{left:55%;background:#48c274;color:#0b2616;border-left-color:#249a4e}.ob-event-drag-ghost.existing{right:40%;background:color-mix(in srgb,var(--ob-blue) 18%,var(--ob-paper));border-left-color:var(--ob-blue)}.ob-event-hour.slot-hot{background:color-mix(in srgb,var(--ob-blue) 8%,transparent)}
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
.ob-ai-wrap{flex:none;display:flex;flex-direction:column;padding:8px 14px 10px;background:#f3f3f3;border-bottom:1px solid var(--ob-line,#dedede);box-sizing:border-box}
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
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-ai-wrap{padding:8px 12px 10px;background:#f3f3f3;border-bottom-color:var(--ob-line,#dedede)}
:is(.ob-outlook-mail,.ob-outlook-calendar) .ob-ai-bar{margin:0;background:#fff;border-color:#e8e8e8;box-shadow:0 2px 8px rgba(0,0,0,.12)}

.ob-inline-editor:focus,.ob-inline-intent input:focus{border-color:var(--ob-blue);box-shadow:0 0 0 1px color-mix(in srgb,var(--ob-blue) 15%,transparent)}
.ob-inline-reply-footer{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:12px}
.ob-ai-suggestion{margin-top:12px;padding:12px;border:1px solid var(--ob-line);border-radius:5px;background:var(--ob-ground);font-size:12px}.ob-ai-suggestion .ob-ai-output{margin:8px 0}
.ob-reading-pane .ob-reading-card{overflow:visible}
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
        people: 'M9 3a3 3 0 1 0 0 6 3 3 0 0 0 0-6ZM3 21v-5a4 4 0 0 1 4-4h4a4 4 0 0 1 4 4v5M17 4a3 3 0 0 1 0 6M18 13a4 4 0 0 1 3 4v4',
        location: 'M12 22s8-8 8-13a8 8 0 0 0-16 0c0 5 8 13 8 13ZM12 6a3 3 0 1 0 0 6 3 3 0 0 0 0-6Z',
        video: 'M3 5h12v14H3V5Zm12 5 6-4v12l-6-4',
        notes: 'M4 3h16v18H4V3ZM8 7h8M8 11h8M8 15h5',
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
        hour12: false,
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
        reply: 'AI 暂不可用，请尝试「明天下午3点开会30分钟」或「给 xxx@yy.com 写一封…邮件」',
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

    function EventEditor({ form, setForm, calendars, saving, error, onSave, onDelete, onClose }) {
      const timeline = React.useRef(null)
      const dialog = React.useRef(null)
      const [timelineData, setTimelineData] = React.useState([])
      const [timelineStatus, setTimelineStatus] = React.useState('')
      const [discarding, setDiscarding] = React.useState(false)
      const initial = React.useRef(JSON.stringify(form))
      const writable = calendars.filter((calendar) => calendar.canEdit !== false)
      const patch = (values) => setForm((old) => ({ ...old, ...values }))
      const close = () => { if (!saving) { if (JSON.stringify(form) !== initial.current) setDiscarding(true); else onClose() } }
      React.useEffect(() => {
        const old = document.activeElement
        dialog.current?.querySelector('input[name=subject]')?.focus()
        if (timeline.current) timeline.current.scrollTop = Math.max(0, (Number(form.startTime.slice(0,2)) - 2) * 48)
        return () => old?.focus?.()
      }, [])
      React.useEffect(() => {
        let alive = true
        setTimelineData([])
        if (!/^\d{4}-\d{2}-\d{2}$/.test(form.date)) { setTimelineStatus('请选择日期'); return }
        setTimelineStatus('正在加载当天日程…')
        const from = new Date(`${form.date}T00:00:00+08:00`)
        const to = new Date(from.getTime() + 86400000)
        api(`/calendar?from=${encodeURIComponent(from.toISOString())}&to=${encodeURIComponent(to.toISOString())}`)
          .then((data) => { if (alive) { setTimelineData(data.events ?? []); setTimelineStatus(data.warnings?.join('；') || '') } })
          .catch((failure) => { if (alive) setTimelineStatus(`当天日程加载失败：${failure.message}`) })
        return () => { alive = false }
      }, [form.date])
      const keyboard = (event) => {
        if (event.key === 'Escape') { event.preventDefault(); close() }
        if (event.key !== 'Tab') return
        const focusable = Array.from(dialog.current.querySelectorAll('button:not(:disabled), input:not(:disabled), select:not(:disabled), textarea:not(:disabled), a[href]'))
        const first = focusable[0], last = focusable[focusable.length - 1]
        if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus() }
        if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus() }
      }
      const start = new Date(`${form.date}T${form.allDay ? '00:00' : form.startTime}:00+08:00`)
      const end = new Date(`${form.endDate}T${form.allDay ? '23:59' : form.endTime}:00+08:00`)
      const validRange = Number.isFinite(+start) && Number.isFinite(+end) && end > start
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
      const suppressHourClick = React.useRef(false)
      const movingExisting = React.useRef(false)
      React.useEffect(() => { timelineDragRef.current = timelineDrag }, [timelineDrag])
      const minOf = (value) => {
        const [hour, minute] = String(value || '0:0').split(':').map(Number)
        return (Number.isFinite(hour) ? hour : 0) * 60 + (Number.isFinite(minute) ? minute : 0)
      }
      const labelOf = (minute) => `${pad(Math.floor(Math.max(0, minute) / 60) % 24)}:${pad(Math.max(0, minute) % 60)}`
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
        const dayStart = new Date(`${form.date}T00:00:00+08:00`)
        const wall = new Date(/[zZ]|[+-]\d{2}:\d{2}$/.test(stamp || '') ? stamp : `${String(stamp || '').slice(0, 19)}+08:00`)
        if (!Number.isFinite(+wall) || !Number.isFinite(+dayStart)) return null
        return (wall - dayStart) / 60000
      }
      const beginTimelineDrag = (event, spec) => {
        if (saving || form.allDay) return
        event.preventDefault()
        event.stopPropagation()
        suppressHourClick.current = false
        const pointer = clientYToMinute(event.clientY)
        const next = {
          ...spec,
          grabOffset: pointer - spec.originStart,
          followStart: spec.originStart,
          previewStart: spec.originStart,
          pointerId: event.pointerId,
        }
        timelineDragRef.current = next
        setTimelineDrag(next)
        try { event.currentTarget.setPointerCapture(event.pointerId) } catch (_) {}
      }
      const onTimelinePointerMove = (event) => {
        const drag = timelineDragRef.current
        if (!drag) return
        const duration = Math.max(SNAP_MIN, drag.duration)
        const raw = clientYToMinute(event.clientY) - drag.grabOffset
        const followStart = clampStart(raw, duration)
        const previewStart = clampStart(snapMinute(raw), duration)
        if (Math.abs(followStart - drag.originStart) >= 1) suppressHourClick.current = true
        const next = { ...drag, followStart, previewStart }
        timelineDragRef.current = next
        setTimelineDrag(next)
      }
      const finishTimelineDrag = async (event) => {
        const drag = timelineDragRef.current
        if (!drag) return
        try { event.currentTarget.releasePointerCapture?.(drag.pointerId) } catch (_) {}
        const duration = Math.max(SNAP_MIN, drag.duration)
        const start = clampStart(snapMinute(drag.previewStart ?? drag.followStart ?? drag.originStart), duration)
        const end = start + duration
        timelineDragRef.current = null
        setTimelineDrag(null)
        if (start === drag.originStart) {
          suppressHourClick.current = false
          return
        }
        // Swallow the click that browsers synthesize after a drag.
        suppressHourClick.current = true
        setTimeout(() => { suppressHourClick.current = false }, 0)
        if (drag.mode === 'draft') {
          patch({ allDay: false, startTime: labelOf(start), endTime: labelOf(end), endDate: form.date, date: form.date })
          return
        }
        if (drag.mode !== 'existing' || !drag.event || movingExisting.current) return
        if (calendars.find((calendar) => calendar.id === drag.event.calendarId)?.canEdit === false) {
          setTimelineStatus('此日历为只读，无法移动日程')
          return
        }
        const snapshot = timelineData
        const newStartIso = `${form.date}T${labelOf(start)}:00`
        const newEndIso = `${form.date}T${labelOf(end)}:00`
        setTimelineData((prev) =>
          (prev || []).map((row) => (row.id === drag.event.id ? { ...row, start: newStartIso, end: newEndIso } : row)),
        )
        setTimelineStatus('正在移动日程…')
        movingExisting.current = true
        try {
          await api('/calendar/update', {
            method: 'POST',
            body: {
              eventId: realTimelineId(drag.event),
              subject: drag.event.subject || '(无主题)',
              date: form.date,
              endDate: form.date,
              startTime: labelOf(start),
              endTime: labelOf(end),
              allDay: false,
              datesOnly: true,
              calendarId: drag.event.calendarId || '',
            },
          })
          setTimelineStatus('日程已移动')
          const from = new Date(`${form.date}T00:00:00+08:00`)
          const to = new Date(from.getTime() + 86400000)
          const data = await api(`/calendar?from=${encodeURIComponent(from.toISOString())}&to=${encodeURIComponent(to.toISOString())}`)
          setTimelineData(data.events ?? [])
        } catch (failure) {
          setTimelineData(snapshot)
          setTimelineStatus(`移动失败：${failure.message}`)
        } finally {
          movingExisting.current = false
        }
      }

      const selectHour = (hour) => {
        if (suppressHourClick.current || timelineDragRef.current) { suppressHourClick.current = false; return }
        const duration = !form.allDay && validRange ? Math.max(SNAP_MIN, endMinute - startMinute) : 30
        const start = hour * 60
        const end = Math.min(1440, start + duration)
        patch({ allDay: false, startTime: labelOf(start), endTime: labelOf(end), endDate: form.date })
      }
      const DURATION_PRESETS = [5, 15, 30, 60]
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
      const visibleTimeline = (timelineData || []).filter((item) => !isEditedTimelineItem(item))
      const field = (icon, children, cls = '') => h('div', { className: `ob-event-field ${cls}` }, h(MailIcon, { name: icon, size: 19 }), h('div', { className: 'ob-event-field-content' }, children))
      return h(Overlay, null, h('div', { className: 'ob-event-mask' },
        h('form', { ref: dialog, className: 'ob-event-editor', role: 'dialog', 'aria-modal': true, 'aria-label': '新建活动', onKeyDown: keyboard,
          onSubmit: (event) => { event.preventDefault(); if (!saving && validRange) onSave() } },
          h('header', { className: 'ob-event-titlebar' }, h('span', null, form.eventId ? '编辑日程 - 日历' : '新建活动 - 日历'),
            h('button', { type: 'button', disabled: saving, onClick: close, 'aria-label': '关闭新建活动' }, '×')),
          h('div', { className: 'ob-event-toolbar' },
            h('button', { type: 'submit', className: 'ob-command ob-new-mail', disabled: saving || !form.subject.trim() || !validRange || !writable.length },
              h(MailIcon, { name: 'save' }), saving ? '保存中…' : form.attendees.trim() ? '发送邀请' : '保存'),
            form.eventId && onDelete
              ? h('button', { type: 'button', className: 'ob-command', disabled: saving, onClick: onDelete },
                  h(MailIcon, { name: 'trash', size: 16 }), '删除日程')
              : null,
            h('span', { className: 'ob-event-toolbar-label' }, h(MailIcon, { name: 'calendar' }), '活动'),
            h('label', { className: 'ob-event-toolbar-label' }, h(MailIcon, { name: 'calendar', size: 16 }),
              h('select', { 'aria-label': '显示为', disabled: saving, value: form.showAs, onChange: (event) => patch({ showAs: event.target.value }) },
                [['busy','忙碌'],['free','空闲'],['tentative','暂定'],['oof','外出'],['workingElsewhere','在别处工作']].map(([value,label]) => h('option',{key:value,value},label)))),
            h('label', { className: 'ob-event-toolbar-label' }, h(MailIcon, { name: 'clock', size: 16 }),
              h('select', { 'aria-label': '提醒', disabled: saving, value: form.reminder, onChange: (event) => patch({ reminder: event.target.value }) },
                [['-1','不提醒'],['0','开始时'],['5','提前 5 分钟'],['15','提前 15 分钟'],['30','提前 30 分钟'],['60','提前 1 小时'],['1440','提前 1 天']].map(([value,label])=>h('option',{key:value,value},label))))),
          error ? h('div',{className:'ob-event-error',role:'alert'},error) : null,
          discarding ? h('div',{className:'ob-event-discard',role:'alert'},'放弃未保存的活动？',
            h('button',{type:'button',onClick:onClose},'放弃'),h('button',{type:'button',onClick:()=>setDiscarding(false)},'继续编辑')) : null,
          h('div',{className:'ob-event-layout'},
            h('div',{className:'ob-event-left'},
              h('fieldset',{className:'ob-event-fields',disabled:saving},
                field('calendar', h('select',{value:form.calendarId,'aria-label':'保存到日历',onChange:(event)=>patch({calendarId:event.target.value})},
                  writable.length ? writable.map((item)=>h('option',{key:item.id,value:item.id},item.name)) : h('option',{value:''},'没有可写入的日历'))),
                field('compose',h('input',{name:'subject',required:true,maxLength:255,placeholder:'添加标题','aria-label':'活动标题',value:form.subject,onChange:(event)=>patch({subject:event.target.value})}),'ob-event-subject'),
                field('people',h('input',{placeholder:'邀请必需的与会者（邮箱，多个用逗号分隔）','aria-label':'与会者邮箱',value:form.attendees,onChange:(event)=>patch({attendees:event.target.value})})),
                field('clock',[
                  h('div',{className:'ob-event-dates',key:'dates'},
                    h('div',{className:'ob-event-date-row',key:'start'},
                      h('span',null,'开始'),
                      h('div',{className:'ob-event-date-inputs'},
                        h('input',{type:'date',required:true,'aria-label':'开始日期',value:form.date,onChange:(event)=>changeDate(event.target.value)}),
                        form.allDay ? null : h('input',{type:'time',required:true,'aria-label':'开始时间',value:form.startTime,onChange:(event)=>patch({startTime:event.target.value})}))),
                    h('div',{className:'ob-event-date-row',key:'end'},
                      h('span',null,'结束'),
                      h('div',{className:'ob-event-date-inputs'},
                        h('input',{type:'date',required:true,min:form.date,'aria-label':'结束日期',value:form.endDate,onChange:(event)=>patch({endDate:event.target.value})}),
                        form.allDay ? null : h('input',{type:'time',required:true,'aria-label':'结束时间',value:form.endTime,onChange:(event)=>patch({endTime:event.target.value})}))),
                    form.allDay ? null : h('div',{className:'ob-event-duration',key:'duration','aria-label':'时长预设'},
                      h('span',{className:'ob-event-duration-label'},'时长'),
                      DURATION_PRESETS.map((minutes)=>h('button',{
                        type:'button',
                        key:minutes,
                        className:`ob-event-duration-btn${currentDuration === minutes ? ' on' : ''}`,
                        disabled:saving,
                        'aria-pressed': currentDuration === minutes,
                        onClick:()=>applyDuration(minutes),
                      }, minutes >= 60 ? '60 分钟' : `${minutes} 分钟`)))),
                  h('div',{className:'ob-event-date-options',key:'options'},h('label',null,h('input',{type:'checkbox',checked:form.allDay,onChange:(event)=>patch({allDay:event.target.checked})}),'全天'),h('span',null,'北京时间 (UTC+08:00)')),
                  !validRange ? h('span',{className:'ob-event-error-text',key:'error'},'结束时间需要晚于开始时间') : null]),
                field('location',h('input',{placeholder:'添加位置','aria-label':'位置',value:form.location,onChange:(event)=>patch({location:event.target.value})}))),
              h('div',{className:'ob-event-body-card'},h(MailIcon,{name:'notes',size:19}),
                h('textarea',{disabled:saving,'aria-label':'活动说明',placeholder:'添加说明或议程…',value:form.body,onChange:(event)=>patch({body:event.target.value})}))),
            h('aside',{className:'ob-event-schedule'},
              h('div',{className:'ob-event-schedule-head'},h('button',{type:'button',disabled:saving,'aria-label':'前一天',onClick:()=>shiftDate(-1)},'‹'),h(MailIcon,{name:'calendar',size:17}),
                h('button',{type:'button',disabled:saving,'aria-label':'后一天',onClick:()=>shiftDate(1)},'›'),h('strong',null,form.date.replaceAll('-',' / '))),
              h('div',{className:'ob-event-timeline-status',role:'status'},timelineStatus || (form.eventId ? '当天日程 · 拖动事件块改时间，点击空白处换到该时段' : '当天日程 · 拖动绿色块改时间，点击空白处设置开始时间')),
              h('div',{className:'ob-event-timeline-scroll',ref:timeline},h('div',{className:'ob-event-timeline'},
                Array.from({length:24},(_,hour)=>h('button',{key:hour,type:'button',className:`ob-event-hour${timelineDrag && Math.floor((timelineDrag.previewStart ?? 0)/60)===hour ? ' slot-hot' : ''}`,disabled:saving,'aria-label':`从 ${hour}:00 开始`,onClick:()=>selectHour(hour)},h('span',null,hour))),
                visibleTimeline.map((item)=>{
                  const a0 = wallMinute(item.start), b0 = wallMinute(item.end)
                  if (a0 == null || b0 == null) return null
                  const dragging = timelineDrag && timelineDrag.mode === 'existing' && timelineDrag.event && timelineDrag.event.id === item.id
                  const a = Math.max(0, dragging ? timelineDrag.originStart : a0)
                  const b = Math.min(1440, dragging ? timelineDrag.originStart + timelineDrag.duration : b0)
                  const editable = !item.allDay && calendars.find((calendar) => calendar.id === item.calendarId)?.canEdit !== false
                  return b>a ? h('div',{
                    key:item.id,
                    className:`ob-event-existing${editable ? ' interactive' : ''}${dragging ? ' drag-origin' : ''}`,
                    title: editable ? `${item.subject}（拖动可改时间）` : item.subject,
                    style:{top:a*PX_PER_MIN,height:Math.max(18,(b-a)*PX_PER_MIN),borderLeftColor:item.color||'#0f6cbd'},
                    onPointerDown: editable ? (event) => beginTimelineDrag(event, {
                      mode: 'existing',
                      event: item,
                      duration: Math.max(SNAP_MIN, b0 - a0),
                      originStart: Math.max(0, a0),
                    }) : undefined,
                    onPointerMove: editable ? onTimelinePointerMove : undefined,
                    onPointerUp: editable ? finishTimelineDrag : undefined,
                    onPointerCancel: editable ? finishTimelineDrag : undefined,
                  }, item.subject) : null
                }),
                validRange ? (() => {
                  const duration = Math.max(SNAP_MIN, endMinute - startMinute)
                  const dragging = timelineDrag && timelineDrag.mode === 'draft'
                  const top = (dragging ? timelineDrag.followStart : startMinute) * PX_PER_MIN
                  const height = Math.max(24, duration * PX_PER_MIN)
                  return h('div',{
                    className:`ob-event-draft-block${dragging ? ' dragging' : ''}`,
                    style:{left:visibleTimeline.length?'55%':44,top,height},
                    onPointerDown: form.allDay || saving ? undefined : (event) => beginTimelineDrag(event, {
                      mode: 'draft',
                      duration,
                      originStart: startMinute,
                    }),
                    onPointerMove: onTimelinePointerMove,
                    onPointerUp: finishTimelineDrag,
                    onPointerCancel: finishTimelineDrag,
                  },
                    h('strong',null,form.allDay?'全天':`${dragging ? labelOf(timelineDrag.followStart) : form.startTime} - ${form.endDate>form.date?`${form.endDate} `:''}${dragging ? labelOf(timelineDrag.followStart + duration) : form.endTime}`),
                    h('span',null,form.subject||'新活动'))
                })() : null,
                timelineDrag ? h('div',{
                  className:`ob-event-drag-ghost ${timelineDrag.mode === 'draft' ? 'draft' : 'existing'}`,
                  style:{
                    top: (timelineDrag.previewStart ?? 0) * PX_PER_MIN,
                    height: Math.max(24, Math.max(SNAP_MIN, timelineDrag.duration) * PX_PER_MIN),
                    borderLeftColor: timelineDrag.mode === 'draft' ? '#249a4e' : (timelineDrag.event?.color || '#0f6cbd'),
                    ...(timelineDrag.mode === 'draft' && !visibleTimeline.length ? { left: 44 } : {}),
                  },
                },
                  h('strong',null,`${labelOf(timelineDrag.previewStart)} - ${labelOf((timelineDrag.previewStart ?? 0) + Math.max(SNAP_MIN, timelineDrag.duration))}`),
                  h('span',null,timelineDrag.mode === 'draft' ? (form.subject||'新活动') : (timelineDrag.event?.subject || ''))
                ) : null)))))))
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
        selected, onSelect, t, timeOf, surface } = props
      const [filtersOpen, setFiltersOpen] = React.useState(false)
      const [ascending, setAscending] = React.useState(false)
      const [collapsed, setCollapsed] = React.useState({})
      const visible = rows
        .filter((row) => filter !== 'unread' || !row.isRead)
        .filter((row) => category === 'all' || (row.category ?? 'other') === category)
        .filter((row) => focus === 'all' || (focus === 'focused' ? row.focused !== false : row.focused === false))
        .sort((a, b) => (ascending ? 1 : -1) * ((parseDate(a.receivedAt)?.getTime() || 0) - (parseDate(b.receivedAt)?.getTime() || 0)))
      const groups = []
      for (const row of visible) {
        const label = dayGroupLabel(parseDate(row.receivedAt) ?? new Date(), t)
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
                active: selected?.id === row.id, t, time: timeOf(row), onSelect, onContextRow: props.onContextRow }))))))
    }

    function MailWorkspace(props) {
      const { t, selected, busy, folders, folder, onFolder, account, onCompose, onDelete,
        onArchive, onToggleRead, onMarkFolderRead, readAllFeedback, connected, surface, readingSurface, unread } = props
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
      return h('div', { className: `ob-mail-workspace${selected ? ' has-selection' : ''}`, style: surface },
        h('div', { className: 'ob-commandbar', role: 'toolbar', 'aria-label': '邮件操作' },
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
          h('aside', { className: 'ob-mail-folders', 'aria-label': '邮件文件夹' },
            h('button', { type: 'button', className: 'ob-folder-heading', 'aria-expanded': favoritesOpen, onClick: () => setFavoritesOpen(!favoritesOpen) },
              h('span', { className: favoritesOpen ? '' : 'ob-rotate' }, h(MailIcon, { name: 'chevron', size: 14 })), '收藏夹'),
            favoritesOpen ? ['inbox', 'sentitems', 'drafts'].map((key) => allFolders.find((f) => f.key === key)).filter(Boolean).map(folderButton) : null,
            h('button', { type: 'button', className: 'ob-folder-heading ob-account-heading', title: account, 'aria-expanded': accountOpen, onClick: () => setAccountOpen(!accountOpen) },
              h('span', { className: accountOpen ? '' : 'ob-rotate' }, h(MailIcon, { name: 'chevron', size: 14 })), h('span', { className: 'ob-folder-name' }, account || 'Outlook')),
            accountOpen ? ['inbox', 'junkemail', 'drafts', 'sentitems', 'deleteditems', 'archive'].map((key) => allFolders.find((f) => f.key === key)).filter(Boolean).map(folderButton) : null),
          h(MailBlock, { ...props, onContextRow }),
          h('section', { className: 'ob-reading-pane', 'aria-label': selected ? '邮件阅读区' : '邮件欢迎页',
            onContextMenu: (event) => { if (selected && !event.target.closest('input,textarea,[contenteditable]') && !window.getSelection()?.toString()) onContextRow(event, selected) } },
            selected ? h(MailReadingPane, { key: selected.id, selected, t, surface: readingSurface,
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

    function MailReadingPane({ selected, t, surface, replyRequest, onSend, busy, onClose }) {
      const [detail, setDetail] = React.useState(null)
      const [detailError, setDetailError] = React.useState('')
      const [detailLoading, setDetailLoading] = React.useState(true)
      const [opened, setOpened] = React.useState(false)
      const [draft, setDraft] = React.useState('')
      const [intent, setIntent] = React.useState('')
      const [summary, setSummary] = React.useState(null)
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
          if (action === 'read') setSummary(result)
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
      return h('article', { className: 'ob-reading-card', style: surface },
        h('h2', { className: 'ob-reading-title' }, selected.subject),
        h('div', { className: 'ob-reading-sender' },
          h('span', { className: 'ob-mail-avatar', style: { background: avatarColor(selected.from), color: '#48423d' } }, initial(selected.from)),
          h('div', { className: 'ob-reading-identity' }, h('strong', null, selected.from), h('span', null, selected.fromAddress)),
          h('button', { type: 'button', className: 'ob-command', onClick: onClose, 'aria-label': '关闭邮件' }, '×')),
        h('div', { className: 'ob-reading-date' }, parseDate(selected.receivedAt)?.toLocaleString() || ''),
        detailLoading ? h('div', { className: 'ob-inline-hint', role: 'status' }, '正在加载完整邮件…') : null,
        detailError ? h('div', { className: 'ob-inline-error', role: 'alert' }, `完整正文加载失败：${detailError}。下方为邮件预览。`) : null,
        h('div', { className: 'ob-reading-content' }, mail.bodyText ?? mail.preview ?? ''),
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
      const { t, tasks = [], unsupported, onToggle, busy, surface, pendingCreate, onPendingCreateConsumed, calendarNonce } = props
      const [cursor, setCursor] = React.useState(() => new Date(new Date().getFullYear(), new Date().getMonth(), 1))
      const [data, setData] = React.useState(null)
      const [loading, setLoading] = React.useState(false)
      const [error, setError] = React.useState(null)
      const [hidden, setHidden] = React.useState(() => new Set())
      const [creating, setCreating] = React.useState(null)
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
        const version = ++requestVersion.current
        setLoading(true)
        try {
          const start = monthGridStart(cursor)
          const end = new Date(start)
          end.setDate(end.getDate() + 42)
          const result = await api(
            `/calendar?from=${encodeURIComponent(start.toISOString())}&to=${encodeURIComponent(end.toISOString())}`,
          )
          if (version !== requestVersion.current) return
          setData(result)
          setError(null)
        } catch (failure) {
          if (version === requestVersion.current) setError(failure.message)
        } finally {
          if (version === requestVersion.current) setLoading(false)
        }
      }, [cursor])

      React.useEffect(() => {
        load()
        return () => { requestVersion.current += 1 }
      }, [load])

      const calendars = data?.calendars ?? []
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
          { label: hidden.has(calendar.id) ? '显示日历' : '隐藏日历', run: () => toggleCalendar(calendar.id) },
          { label: '只显示此日历', run: () => setHidden(new Set(calendars.filter((item) => item.id !== calendar.id).map((item) => item.id))) },
          { label: '显示全部日历', run: () => setHidden(new Set()) },
        ]) },
          h('input', { type: 'checkbox', checked: !hidden.has(calendar.id), style: { accentColor: calendar.color }, onChange: () => toggleCalendar(calendar.id) }),
          h('span', { title: calendar.name }, calendar.name))))

      const openCreate = (day, overrides = {}) => {
        setForm({ subject: '', date: dayKey(day), endDate: dayKey(day), startTime: '09:00', endTime: '09:30', allDay: false,
          location: '', attendees: '', body: '', showAs: 'busy', reminder: '15',
          calendarId: calendars.find((item) => item.isDefault && item.canEdit !== false)?.id || calendars.find((item) => item.canEdit !== false)?.id || '',
          eventId: '', listEventId: '',
          transactionId: window.crypto?.randomUUID?.() || `event-${Date.now()}-${Math.random().toString(16).slice(2)}`, ...overrides })
        setCreateError(null)
        setCreating(day)
      }

      React.useEffect(() => {
        if (!pendingCreate) return
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
      }, [pendingCreate])

      React.useEffect(() => {
        if (!calendarNonce) return
        load()
      }, [calendarNonce, load])


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
      ])
      const openEdit = async (item) => {
        const realId = String(item.id).includes(':') ? String(item.id).slice(String(item.id).indexOf(':') + 1) : String(item.id)
        const day = parseDate(item.start) || new Date()
        setCreateError(null)
        try {
          const detail = await api(`/calendar/event?id=${encodeURIComponent(realId)}`)
          const startDate = parseDate(detail.event.start) || day
          const endDate = parseDate(detail.event.end) || startDate
          setForm({
            subject: detail.event.subject || '',
            date: dayKey(startDate),
            endDate: dayKey(endDate),
            startTime: fmtTime(startDate),
            endTime: fmtTime(endDate),
            allDay: Boolean(detail.event.allDay),
            location: detail.event.location || '',
            attendees: detail.event.attendees || '',
            body: htmlToPlainText(detail.event.body || ''),
            showAs: detail.event.showAs || 'busy',
            reminder: String(detail.event.reminder ?? '15'),
            calendarId: detail.event.calendarId || item.calendarId || '',
            eventId: detail.event.id || realId,
            listEventId: item.id,
            transactionId: window.crypto?.randomUUID?.() || `event-${Date.now()}`,
          })
          setCreating(startDate)
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

      const submitCreate = async () => {
        if (!creating || savingRequest.current) return
        savingRequest.current = true
        setSaving(true)
        setCreateError(null)
        try {
          await api(form.eventId ? '/calendar/update' : '/calendar/create', { method: 'POST', body: form })
          setCreating(null)
          setNotice(form.eventId ? '日程已更新' : t('cal.created'))
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
      const scheduleEventClick = (item, day, anchorEl) => {
        if (eventClickTimer.current) clearTimeout(eventClickTimer.current)
        eventClickTimer.current = setTimeout(() => {
          eventClickTimer.current = null
          openEventPeek(item, day, anchorEl)
        }, 220)
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
      const deletePeekEvent = async (item) => {
        if (!item || savingRequest.current) return
        if (!eventCanEdit(item)) {
          setNotice('此日历为只读，无法删除日程')
          return
        }
        if (!window.confirm('确定删除这条日程吗？此操作不可撤销。')) return
        savingRequest.current = true
        try {
          await api('/calendar/delete', { method: 'POST', body: { id: realEventId(item) } })
          setEventPeek(null)
          setNotice('日程已删除')
          await load()
        } catch (failure) {
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
        h('div', { className: 'ob-commandbar ob-calendar-toolbar', role: 'toolbar', 'aria-label': '日历操作' },
          h('button', { type: 'button', className: 'ob-command ob-new-mail', disabled: busy || saving, onClick: () => openCreate(selectedDay) }, h(MailIcon, { name: 'calendar' }), '新事件'),
          h('span', { className: 'ob-command-divider' }),
          h('span', { className: 'ob-calendar-view-label' }, h(MailIcon, { name: 'calendar', size: 17 }), '月'),
          h('button', { type: 'button', className: `ob-command${showSidebar ? ' active' : ''}`, 'aria-pressed': showSidebar, onClick: () => setShowSidebar(!showSidebar) }, h(MailIcon, { name: 'filter' }), '日历筛选'),
          h('span', { className: 'ob-command-divider' }),
          h('button', { type: 'button', className: 'ob-command', disabled: loading, onClick: load }, h(MailIcon, { name: 'refresh', size: 17 }), t('action.refresh')),
          h('span', { className: 'ob-spacer' }),
          h('span', { className: 'ob-calendar-hint' }, '双击日期新建日程')),
        h('div', { className: 'ob-calendar-layout' },
          showSidebar ? h('aside', { className: 'ob-calendar-sidebar', 'aria-label': '日历导航' },
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
                      style: { borderLeftColor: event.color || '#0f6cbd', background: `color-mix(in srgb, ${event.color || '#0f6cbd'} 24%, var(--ob-paper))` },
                      title: `${event.subject}${event.location ? ` · ${event.location}` : ''}${editable ? '（拖到其他日期可移动）' : ''}`,
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
                        style: {
                          borderLeftColor: dragState.event.color || '#0f6cbd',
                          background: `color-mix(in srgb, ${dragState.event.color || '#0f6cbd'} 24%, var(--ob-paper))`,
                        },
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
            },
              h('div', { className: 'ob-event-peek-accent', style: { background: eventPeek.item.color || '#13a10e' } }),
              h('div', { className: 'ob-event-peek-body' },
                h('div', { className: 'ob-event-peek-cat' },
                  h('span', { className: 'ob-event-peek-dot', style: { background: eventPeek.item.color || '#13a10e' } }),
                  h('span', null, eventPeek.item.calendar || '日历')),
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
            onClick: () => setDayPeek(null),
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
                h('button', { type: 'button', className: 'ob-btn', onClick: () => setDayPeek(null) }, t('action.close'))),
              (byDay.get(dayKey(dayPeek.day)) ?? []).map((event) => h('button', {
                type: 'button',
                key: event.id,
                className: 'ob-day-peek-item',
                style: { borderLeftColor: event.color || '#0f6cbd' },
                onClick: (e) => { openEventPeek(event, dayPeek.day, e.currentTarget); setDayPeek(null) },
                onDoubleClick: (e) => { e.preventDefault(); setDayPeek(null); handleEventDoubleClick(event) },
                onContextMenu: (mouse) => eventContext(mouse, event, dayPeek.day),
              },
                h('strong', null, event.subject),
                h('span', null, event.allDay ? '全天' : `${fmtTime(parseDate(event.start))} – ${fmtTime(parseDate(event.end))}`)))))
        ) : null,
        creating ? h(EventEditor, { form, setForm, calendars, saving, error: createError, onSave: submitCreate, onDelete: removeEvent, onClose: () => { if (!saving) setCreating(null) } }) : null,
      )
    }

    function CalendarPage(props) {
      const { t, events, tasks, unsupported, onToggle, busy, surface, pendingCreate, onPendingCreateConsumed, calendarNonce } = props
      return h(MonthCalendar, { t, tasks, unsupported, onToggle, busy, surface, pendingCreate, onPendingCreateConsumed, calendarNonce })
    }

    function AppearancePage(props) {
      const { t, appearance, setAppearance, onSave, onReset, onClose, busy } = props
      const regions = [
        ['panel', t('ui.panel')],
        ['mail', t('ui.mail')],
        ['calendar', t('ui.calendar')],
        ['todo', t('ui.todo')],
        ['reply', t('ui.reply')],
      ]
      const valueOf = (key, field, fallback) => {
        const row = appearance?.[key] ?? {}
        const raw = row[field]
        return Number.isFinite(Number(raw)) ? Number(raw) : fallback
      }
      const update = (key, field, value) => {
        const next = { ...(appearance ?? {}) }
        next[key] = { ...(next[key] ?? {}), [field]: Number(value) }
        setAppearance(next)
      }
      return h(
        'div',
        {
          className: 'ob-modal-mask',
          onClick: (event) => {
            if (event.target === event.currentTarget) onClose()
          },
        },
        h(
          'div',
          { className: 'ob-modal' },
          h(
            'div',
            { className: 'ob-modal-head' },
            h('span', null, t('ui.title')),
            h('button', { type: 'button', className: 'ob-btn', onClick: onClose }, t('action.close')),
          ),
          h(
            'div',
            { className: 'ob-ui-page' },
            h('div', { className: 'ob-ui-head' }, t('ui.hint')),
        h(
          'div',
          { className: 'ob-surface-grid' },
          regions.map(([key, label]) =>
            h(
              'div',
              { className: 'ob-surface-card', key },
              h(
                'div',
                { className: 'ob-surface-title' },
                label,
                h('span', { className: 'ob-surface-badge' }, `${valueOf(key, 'opacity', 100)}%`),
              ),
              h(
                'label',
                { className: 'ob-slider' },
                h('span', null, t('ui.opacity')),
                h('input', {
                  type: 'range',
                  min: 0,
                  max: 100,
                  step: 1,
                  value: valueOf(key, 'opacity', 100),
                  onChange: (event) => update(key, 'opacity', event.target.value),
                }),
                h('b', null, `${valueOf(key, 'opacity', 100)}%`),
              ),
              h(
                'label',
                { className: 'ob-slider' },
                h('span', null, t('ui.blur')),
                h('input', {
                  type: 'range',
                  min: 0,
                  max: 40,
                  step: 1,
                  value: valueOf(key, 'blur', 0),
                  onChange: (event) => update(key, 'blur', event.target.value),
                }),
                h('b', null, `${valueOf(key, 'blur', 0)}px`),
              ),
            ),
          ),
        ),
            h(
              'div',
              { className: 'ob-ui-foot' },
              h(
                'button',
                { type: 'button', className: 'ob-btn primary', disabled: busy, onClick: onSave },
                t('ui.save'),
              ),
              h('button', { type: 'button', className: 'ob-btn', disabled: busy, onClick: onReset }, t('ui.reset')),
            ),
          ),
        ),
      )
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

    function TodoBlock(props) {
      const { tasks, t, onToggle, busy, unsupported, surface } = props
      return h(
        'section',
        { className: 'ob-card grow', style: surface },
        h(
          'div',
          { className: 'ob-card-head' },
          h('span', { className: 'ob-card-title' }, t('todo.title')),
          h('span', { className: 'ob-count' }, tasks.length > 0 ? `${tasks.length}` : ''),
        ),
        h(
          'div',
          { className: 'ob-list' },
          tasks.length === 0
            ? h('div', { className: 'ob-empty' }, unsupported ? t('todo.unsupported') : t('todo.empty'))
            : tasks.map((task) =>
                h(
                  'div',
                  { className: 'ob-todo', key: `${task.listId}:${task.id}` },
                  h('button', {
                    type: 'button',
                    className: `ob-check${task.completed ? ' done' : ''}`,
                    'aria-label': task.title,
                    disabled: busy,
                    onClick: () => onToggle(task),
                  }),
                  h(
                    'span',
                    { className: 'ob-todo-main' },
                    h('div', { className: `ob-todo-title${task.completed ? ' done' : ''}` }, task.title),
                    h(
                      'div',
                      { className: 'ob-todo-meta' },
                      [task.due ? t('todo.due', { time: fmtDue(parseDate(task.due), t) }) : '', task.listName]
                        .filter(Boolean)
                        .join(' 路 '),
                    ),
                  ),
                ),
              ),
        ),
      )
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
      const [showUi, setShowUi] = React.useState(false)
      const [appearance, setAppearance] = React.useState(null)
      const [replyRequest, setReplyRequest] = React.useState(null)
      const [category, setCategory] = React.useState('all')
      const [focus, setFocus] = React.useState('focused')
      const [folder, setFolder] = React.useState('inbox')
      const [readAllFeedback, setReadAllFeedback] = React.useState(null)
      const activeFolder = React.useRef(folder)
      activeFolder.current = folder
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
            if (state.settings.appearance) setAppearance(state.settings.appearance)
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
        try {
          const result = await api(`/mail?folder=${encodeURIComponent(key)}`)
          const rows = result.mail ?? []
          if (activeFolder.current === key) setFolderMail(rows)
          return rows
        } catch (failure) {
          setError(failure.message)
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
        await loadFolderMail(folder)
        await loadFolders()
        store.load(true)
      }, [folder, loadFolderMail, loadFolders])

      const deleteMail = React.useCallback(async (row) => {
        if (!row) return
        setBusy(true)
        try {
          await api('/mail/delete', { method: 'POST', body: { id: row.id } })
          await afterMailAction()
        } catch (failure) {
          setError(failure.message)
        } finally {
          setBusy(false)
        }
      }, [afterMailAction])

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
        const style = {
          background: `color-mix(in srgb, var(--dsw-alias-bg-layer-1) ${opacity}%, transparent)`,
        }
        if (blur > 0) {
          style.backdropFilter = `blur(${blur}px)`
          style.WebkitBackdropFilter = `blur(${blur}px)`
        }
        return style
      }

      const saveAppearance = React.useCallback(async () => {
        setBusy(true)
        setError(null)
        try {
          await api('/settings', { method: 'POST', body: { appearance } })
          setNotice(t('ui.saved'))
        } catch (failure) {
          setError(failure.message)
        } finally {
          setBusy(false)
        }
      }, [appearance, t])

      const resetAppearance = React.useCallback(() => {
        const blank = {}
        for (const key of ['panel', 'mail', 'calendar', 'todo', 'reply']) blank[key] = { opacity: 100, blur: 0 }
        setAppearance(blank)
      }, [])

      const toggleTask = React.useCallback(async (task) => {
        setBusy(true)
        try {
          await api('/task/complete', {
            method: 'POST',
            body: { listId: task.listId, taskId: task.id, completed: true },
          })
          store.patch({
            tasks: (store.snapshot?.tasks ?? []).filter(
              (item) => !(item.id === task.id && item.listId === task.listId),
            ),
          })
        } catch (failure) {
          setError(failure.message)
        } finally {
          setBusy(false)
        }
      }, [])


      const submitAiCommand = React.useCallback(async () => {
        const text = String(aiText || '').trim()
        if (!text || aiBusy) return
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
          } else if (intent === 'mail' && result.mail) {
            const mail = result.mail
            setCompose({ to: mail.to || '', subject: mail.subject || '', body: mail.body || '' })
            setAiText('')
            setAiMsg(usedFallback || result.fallback ? '已预填邮件草稿，请核对后发送' : '已打开邮件草稿，请核对后发送')
          } else {
            setAiMsg(null)
            setAiErr(result?.reply || '未能识别意图，请试试添加日程或写邮件')
          }
        } catch (failure) {
          setAiMsg(null)
          setAiErr(failure?.message || String(failure))
        } finally {
          setAiBusy(false)
        }
      }, [aiText, aiBusy])

      const connected = Boolean(snap?.connected)
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
          error ? h('div', { className: 'ob-error' }, error) : null,
          notice ? h('div', { className: 'ob-banner' }, notice) : null,
          demo
            ? h(
                'div',
                { className: 'ob-banner' },
                h('span', null, t('notice.demo')),
                h('span', { className: 'ob-spacer' }),
                h(
                  'button',
                  { type: 'button', className: 'ob-btn primary', onClick: startConnect, disabled: busy },
                  t('action.connect'),
                ),
              )
            : null,
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
          view === 'calendar'
            ? h(CalendarPage, {
                t,
                events: snap.events ?? [],
                tasks: snap.tasks ?? [],
                unsupported: snap.todoUnsupported,
                onToggle: toggleTask,
                busy,
                surface: surfaceOf('calendar'),
                pendingCreate,
                onPendingCreateConsumed: () => setPendingCreate(null),
                calendarNonce,
              })
            : h(MailWorkspace, {
                  t, selected, busy, connected, folders, folder, filter, setFilter, focus, setFocus, category, setCategory,
                  account: snap.account?.mail || snap.account?.displayName || 'Outlook',
                  rows: folderMail ?? (folder === 'inbox' ? snap.mail ?? [] : []), unread: Number(snap.unread || 0),
                  onFolder: (key) => { setFolder(key); setSelected(null); setFolderMail(null); setFocus(key === 'inbox' ? 'focused' : 'all'); setCategory('all') },
                  onSelect: (row) => { setSelected(row); if (row && !row.isRead && !busy) markRead(row) },
                  onCompose: () => setCompose({ to: '', subject: '', body: '' }),
                  onDelete: deleteMail, onArchive: archiveMail, onToggleRead: toggleRead,
                  onMarkFolderRead: markAllRead, readAllFeedback,
                  onReplyRow: (row) => { setSelected(row); setReplyRequest({ id: row.id, at: Date.now() }) },
                  replyRequest, onSendReply: sendReply,
                  onRefreshFolder: async (key) => { setFolder(key); setSelected(null); setFocus(key === 'inbox' ? 'focused' : 'all'); await loadFolderMail(key); await loadFolders() },
                  surface: surfaceOf('mail'), readingSurface: surfaceOf('reply'),
                  timeOf: (row) => fmtMailTime(parseDate(row.receivedAt), t),
                }),
          (snap.warnings ?? []).length > 0
            ? h('div', { className: 'ob-error' }, (snap.warnings ?? []).join('; '))
            : null,
        )
      })()

      return h(
        'div',
        { className: `ob-root${!showSettings && !needsConnect ? view === 'mail' ? ' ob-outlook-mail' : view === 'calendar' ? ' ob-outlook-calendar' : '' : ''}` },
        h(Sheet, null),
        h(ContextMenu, null),
        h(
          'header',
          { className: 'ob-head' },
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
          connected
            ? h('span', { className: 'ob-chip ok' }, 'Outlook')
            : h('span', { className: 'ob-chip' }, t('chip.offline')),
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
              { className: 'ob-ai-wrap' },
              h(
                'div',
                { className: 'ob-ai-bar', role: 'search' },
                h('span', { className: 'ob-ai-bar-mark', 'aria-hidden': 'true' }, 'AI'),
              h('input', {
                className: 'ob-ai-bar-input',
                type: 'text',
                value: aiText,
                disabled: aiBusy,
                placeholder: '试试：明天下午3点开会一小时 / 给 xxx 写一封关于…的邮件',
                'aria-label': 'AI 对话：添加日程或写邮件',
                onChange: (event) => {
                  setAiText(event.target.value)
                  if (aiErr) setAiErr(null)
                  if (aiMsg && aiMsg !== '处理中…') setAiMsg(null)
                },
                onKeyDown: (event) => {
                  if (event.key === 'Enter' && !event.shiftKey) {
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
        showUi
          ? h(AppearancePage, {
              t,
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
