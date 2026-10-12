import { Bell, CheckCheck, MessageSquare, ShieldAlert, Workflow } from 'lucide-react'
import { useMemo, useState } from 'react'

export type NotificationItem = {
  id: string
  category: string
  title: string
  body: string
  priority: string
  readAt: string | null
  createdAt: string
}

export function NotificationCenter({
  items,
  onMarkAllRead,
  onToggleRead,
}: {
  items: NotificationItem[]
  onMarkAllRead: () => Promise<void>
  onToggleRead: (id: string, read: boolean) => Promise<void>
}) {
  const [tab, setTab] = useState('all')
  const [message, setMessage] = useState('')
  const [busyId, setBusyId] = useState<string | null>(null)
  const [markingAll, setMarkingAll] = useState(false)
  const visible = useMemo(() => items.filter((item) => tab === 'all' || item.category === tab), [items, tab])
  const grouped = useMemo(() => groupNotifications(visible), [visible])
  const unreadCount = items.filter((item) => !item.readAt).length
  const mark = async () => {
    setMarkingAll(true)
    try {
      await onMarkAllRead()
      setMessage('All notifications marked as read.')
    } finally {
      setMarkingAll(false)
    }
  }
  const toggle = async (item: NotificationItem) => {
    setBusyId(item.id)
    try {
      await onToggleRead(item.id, !item.readAt)
      setMessage(item.readAt ? 'Notification marked unread.' : 'Notification marked read.')
    } finally {
      setBusyId(null)
    }
  }
  return (
    <div className="page-content notification-page">
      <section className="page-heading">
        <div>
          <p className="kicker">YOUR WORK PULSE</p>
          <h1>Notification Center</h1>
          <p>Approval, collaboration, security and system signals in one place.</p>
        </div>
        <button className="secondary-button" onClick={mark} disabled={markingAll || unreadCount === 0}>
          <CheckCheck size={17} /> {markingAll ? 'Marking…' : `Mark all read (${unreadCount})`}
        </button>
      </section>
      <div className="section-tabs" role="tablist" aria-label="Notification categories">
        {['all', 'workflow', 'collaboration', 'security', 'system'].map((value) => (
          <button
            role="tab"
            aria-selected={tab === value}
            className={tab === value ? 'active' : ''}
            onClick={() => setTab(value)}
            key={value}
          >
            {value === 'workflow' ? (
              <Workflow size={16} />
            ) : value === 'collaboration' ? (
              <MessageSquare size={16} />
            ) : value === 'security' ? (
              <ShieldAlert size={16} />
            ) : (
              <Bell size={16} />
            )}{' '}
            {value[0].toUpperCase() + value.slice(1)}
          </button>
        ))}
      </div>
      <p className="sr-live" aria-live="polite">
        {message}
      </p>
      <section className="notification-list">
        {Object.entries(grouped).map(([label, rows]) => (
          <div key={label}>
            <h2>{label}</h2>
            {rows.map((item) => (
              <article className={item.readAt ? 'notification-row read' : 'notification-row'} key={item.id} aria-busy={busyId === item.id}>
                <span className={`notification-icon ${item.priority}`}>
                  <Bell size={18} />
                </span>
                <div>
                  <span className={item.readAt ? 'read-state read' : 'read-state'}>{item.readAt ? 'Read' : 'Unread'}</span>
                  <h3>{item.title}</h3>
                  <p>{item.body}</p>
                  <time dateTime={item.createdAt} title={new Date(item.createdAt).toLocaleString()}>
                    {new Intl.RelativeTimeFormat('en', { numeric: 'auto' }).format(
                      relativeDays(item.createdAt),
                      'day',
                    )}
                  </time>
                </div>
                <button
                  className="text-button"
                  disabled={busyId === item.id || markingAll}
                  onClick={() => void toggle(item)}
                  aria-label={item.readAt ? 'Mark notification unread' : 'Mark notification read'}
                >
                  {busyId === item.id ? 'Saving…' : item.readAt ? 'Mark unread' : 'Mark read'}
                </button>
              </article>
            ))}
          </div>
        ))}
      </section>
      {!visible.length && (
        <div className="empty-state panel">
          <Bell size={32} />
          <h2>No notifications here</h2>
          <p>New activity will appear in this view.</p>
        </div>
      )}
    </div>
  )
}

function groupNotifications(items: NotificationItem[]) {
  const groups: Record<string, NotificationItem[]> = {}
  for (const item of items) {
    const days = Math.abs(relativeDays(item.createdAt))
    const label = days === 0 ? 'Today' : days <= 7 ? 'This week' : 'Earlier'
    ;(groups[label] ??= []).push(item)
  }
  return groups
}
function relativeDays(value: string) {
  return Math.round((new Date(value).getTime() - Date.now()) / 86_400_000)
}
