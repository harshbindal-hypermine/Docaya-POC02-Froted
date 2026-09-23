import { local } from './translations'
import { useCallback, useEffect, useMemo, useState } from 'react'
import {
  AlertTriangle,
  ArrowDownToLine,
  ArrowRight,
  Bell,
  BookOpen,
  CalendarClock,
  ChartNoAxesCombined,
  Check,
  ChevronRight,
  CircleHelp,
  CloudUpload,
  FileText,
  FolderOpen,
  Globe2,
  LayoutDashboard,
  LockKeyhole,
  LogOut,
  Mail,
  Menu,
  MessageSquare,
  Plus,
  Search,
  Settings2,
  ShieldCheck,
  Workflow,
  X,
} from 'lucide-react'
import {
  addEvent,
  canSee,
  capabilities,
  changeRecord,
  csv,
  departments,
  downloadFile,
  loadWorkspace,
  people,
  roles,
  seedWorkspace,
  STORAGE_KEY,
  type Language,
  type RecordItem,
  type Role,
  type Workspace,
} from './model'
import { DocumentTable, Empty, Field, Modal, type Translate } from './ui'
import { Registration } from './Registration'
import { RecordDetail } from './RecordDetail'
import {
  AdminPage,
  AuditPage,
  CorrespondencePage,
  DocumentLibraryPage,
  Operations,
  RecordsPage,
  ReportsPage,
  type PageProps,
} from './Management'
import { clearOriginals, readOriginal } from './files'
import { SearchPage, AssistantDock } from './Discovery'
import { ApprovalQueue } from './ApprovalQueue'
import { approvalQueue, type ApprovalView } from './approval-queue'
import { DocayaMark, DocayaAIMark } from './BrandMarks'
import { developmentSignIn, uploadSharePointFile } from '../services/api'
import './prototype.css'
import './experience.css'
import './typography.css'
import './approval-queue.css'
import './brand.css'
type Page =
  | 'home'
  | 'documents'
  | 'library'
  | 'workflows'
  | 'operations'
  | 'search'
  | 'correspondence'
  | 'records'
  | 'notifications'
  | 'reports'
  | 'audit'
  | 'admin'
  | 'workshop'
const navigation: {
  id: Page
  en: string
  ar: string
  icon: typeof LayoutDashboard | typeof DocayaAIMark
  group: number
}[] = [
  { id: 'home', en: 'Overview', ar: 'نظرة عامة', icon: LayoutDashboard, group: 0 },
  { id: 'documents', en: 'Document register', ar: 'سجل الوثائق', icon: FolderOpen, group: 1 },
  { id: 'library', en: 'Document Library', ar: 'مكتبة الوثائق', icon: BookOpen, group: 1 },
  { id: 'workflows', en: 'Approvals', ar: 'الاعتمادات', icon: Workflow, group: 1 },
  { id: 'operations', en: 'Staging & publishing', ar: 'التجهيز والنشر', icon: CloudUpload, group: 1 },
  { id: 'correspondence', en: 'Correspondence', ar: 'المراسلات', icon: Mail, group: 1 },
  { id: 'search', en: 'Search & Ask Docaya', ar: 'البحث واسأل دوكايا', icon: DocayaAIMark, group: 2 },
  { id: 'records', en: 'Records & retention', ar: 'السجلات والحفظ', icon: LockKeyhole, group: 2 },
  {
    id: 'reports',
    en: 'Reports & analytics',
    ar: 'التقارير والتحليلات',
    icon: ChartNoAxesCombined,
    group: 2,
  },
  { id: 'audit', en: 'Audit trail', ar: 'سجل التدقيق', icon: ShieldCheck, group: 2 },
  { id: 'admin', en: 'Admin center', ar: 'مركز الإدارة', icon: Settings2, group: 3 },
  { id: 'workshop', en: 'Client workshop', ar: 'ورشة العميل', icon: MessageSquare, group: 3 },
]
const modules = [
  [
    'M1',
    'Capture & registration',
    'الالتقاط والتسجيل',
    'documents',
    'Select or drop a file, inspect its fingerprint and save a draft.',
    'اختر ملفاً وافحص بصمته واحفظ مسودة.',
  ],
  [
    'M2',
    'AI-assisted classification',
    'التصنيف بمساعدة الذكاء الاصطناعي',
    'documents',
    'Review automatic suggestions, override any field, and confirm the classification before submission.',
    'راجع الاقتراحات التلقائية وعدل أي حقل وأكد التصنيف قبل الإرسال.',
  ],
  [
    'M3',
    'Workflow & approval',
    'سير العمل والاعتماد',
    'workflows',
    'Try single, sequential, parallel and conditional approvals; return with comments or delegate.',
    'جرّب مسارات الاعتماد والإرجاع مع تعليق أو التفويض.',
  ],
  [
    'M4',
    'Promotion & publishing',
    'النشر والترقية',
    'operations',
    'Approve to publish the original format and record its version; retry a failed promotion.',
    'اعتمد لنشر الصيغة الأصلية وتسجيل الإصدار أو أعد محاولة النشر.',
  ],
  [
    'M5',
    'Staging lifecycle',
    'دورة حياة التجهيز',
    'operations',
    'Compare retained, returned and published staging records; configure purge or retain.',
    'قارن سجلات التجهيز واضبط الحذف أو الاحتفاظ.',
  ],
  [
    'M6',
    'AI search & RAG',
    'البحث الذكي',
    'search',
    'Run a demo indexing batch, then ask a question and open a cited published record.',
    'شغّل دفعة فهرسة واسأل سؤالاً وافتح السجل المستشهد به.',
  ],
  [
    'M7',
    'Notification center',
    'مركز الإشعارات',
    'notifications',
    'Review lifecycle alerts, read/unread controls and mandatory security preferences.',
    'راجع تنبيهات الدورة والقراءة وتفضيلات الأمان الإلزامية.',
  ],
  [
    'M8',
    'Correspondence',
    'المراسلات',
    'correspondence',
    'Register incoming/outgoing items, reply, forward and acknowledge a circular.',
    'سجل الوارد والصادر وأضف رداً وحول وأكد قراءة تعميم.',
  ],
  [
    'M9',
    'Retention & legal hold',
    'الحفظ والحجز القانوني',
    'records',
    'Apply a hold, see disposition blocked, release with reason and archive.',
    'طبق حجزاً وراجع منع الإتلاف ثم ارفع الحجز بسبب وأرشف.',
  ],
  [
    'M10',
    'Search & discovery',
    'البحث والاكتشاف',
    'search',
    'Filter published records by class, department and sensitivity; save a search.',
    'صفِّ السجلات بالفئة والإدارة والحساسية واحفظ البحث.',
  ],
  [
    'M11',
    'Identity & integration',
    'الهوية والتكامل',
    'admin',
    'Explore the UAE PASS, Entra and governed SharePoint service identity design.',
    'استكشف تصميم الهوية الوطنية وإنترّا وهوية خدمة شيربوينت.',
  ],
  [
    'M12',
    'Access control',
    'التحكم بالوصول',
    'admin',
    'Switch personas and departments to see the demo access rules change results.',
    'بدل الأدوار والإدارات لمشاهدة تأثير الصلاحيات على النتائج.',
  ],
  [
    'M13',
    'Admin center',
    'مركز الإدارة',
    'admin',
    'Add a class; change library routing, SLA, retention and notification defaults.',
    'أضف فئة وعدل التوجيه ومدد الاعتماد والحفظ والتنبيهات.',
  ],
  [
    'M14',
    'Reporting & analytics',
    'التقارير والتحليلات',
    'reports',
    'Inspect metrics calculated from demo records; export CSV or print to PDF.',
    'راجع مؤشرات السجلات التجريبية وصدّر CSV أو اطبع PDF.',
  ],
  [
    'M15',
    'Audit & compliance',
    'التدقيق والامتثال',
    'audit',
    'Trace decisions, holds and changes; export events and a demo signature certificate.',
    'تتبع القرارات والحجوزات والتغييرات وصدّر الأحداث وشهادة تجريبية.',
  ],
] as const
export default function PrototypeApp() {
  const [state, setState] = useState<Workspace>(loadWorkspace)
  const [lang, setLang] = useState<Language>(() =>
    localStorage.getItem('docaya-prototype-language') === 'ar' ? 'ar' : 'en',
  )
  const [signedIn, setSignedIn] = useState(() => sessionStorage.getItem('docaya-prototype-session') === 'yes')
  const [role, setRole] = useState<Role>('System Administrator')
  const [department, setDepartment] = useState('Operations')
  const [page, setPage] = useState<Page>('home')
  const [mobile, setMobile] = useState(false)
  const [selectedId, setSelectedId] = useState('')
  const [approvalView, setApprovalView] = useState<ApprovalView>('pending')
  const [approvalQuery, setApprovalQuery] = useState('')
  const [approvalOverdue, setApprovalOverdue] = useState(false)
  const [registration, setRegistration] = useState<false | 'new' | RecordItem>(false)
  const [toast, setToast] = useState('')
  const [storageError, setStorageError] = useState('')
  const [reset, setReset] = useState(false)
  const [globalQuery, setGlobalQuery] = useState('')
  const t: Translate = useCallback((en, ar) => (lang === 'ar' ? ar : en), [lang])
  const rights = capabilities(role)
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state))
      setStorageError('')
    } catch {
      setStorageError(
        t(
          'Browser storage is full. Changes will last only for this session; export your feedback before leaving.',
          'مساحة المتصفح ممتلئة. التغييرات لهذه الجلسة فقط؛ صدّر الملاحظات قبل المغادرة.',
        ),
      )
    }
  }, [state, t])
  useEffect(() => {
    document.documentElement.lang = lang
    document.documentElement.dir = lang === 'ar' ? 'rtl' : 'ltr'
    document.title =
      lang === 'ar' ? 'دوكايا | ذكاء إدارة الوثائق' : 'Docaya | Document Management Intelligence'
    localStorage.setItem('docaya-prototype-language', lang)
  }, [lang])
  useEffect(() => {
    if (toast) {
      const timer = setTimeout(() => setToast(''), 6500)
      return () => clearTimeout(timer)
    }
  }, [toast])
  useEffect(() => {
    if (!signedIn) return
    let cancelled = false
    developmentSignIn().catch(() => {
      if (!cancelled) {
        setToast(
          t(
            'Backend API is not available. SharePoint actions will not run until the API is started and configured.',
            'واجهة الخادم الخلفية غير متاحة. لن تعمل إجراءات شيربوينت حتى يتم تشغيلها وتهيئتها.',
          ),
        )
      }
    })
    return () => {
      cancelled = true
    }
  }, [signedIn, t])
  const visible = useMemo(
    () => state.documents.filter((r) => canSee(r, role, department)),
    [state.documents, role, department],
  )
  const selected = visible.find((r) => r.id === selectedId)
  const reviewQueue = approvalQueue(visible, approvalView, approvalQuery, approvalOverdue)
  const resetApprovalView = () => {
    setApprovalView('pending')
    setApprovalQuery('')
    setApprovalOverdue(false)
  }
  const notify = (text: string) => setToast(text)
  const navigate = (next: Page) => {
    if ((next === 'admin' && !rights.admin) || (next === 'audit' && !rights.audit)) {
      notify(
        t(
          'Choose an administrator or auditor persona to explore this module.',
          'اختر دور مدير أو مدقق لاستكشاف هذه الوحدة.',
        ),
      )
      return
    }
    setPage(next)
    if (next === 'workflows') resetApprovalView()
    setSelectedId('')
    setMobile(false)
    setGlobalQuery('')
    window.scrollTo(0, 0)
  }
  const open = (record: RecordItem) => {
    if (page === 'home' && record.status === 'PendingApproval') {
      resetApprovalView()
      setPage('workflows')
    }
    setSelectedId(record.id)
    setState((s) => addEvent(s, role, 'Document viewed', 'تم عرض الوثيقة', record.id))
  }
  const props: PageProps = {
    state,
    setState,
    role,
    department,
    lang,
    t,
    visible,
    onOpen: open,
    notify,
    selectedId: selected?.id || '',
  }
  const notices = state.notices.filter((n) => !n.recordId || visible.some((r) => r.id === n.recordId))
  const unread = notices.filter((n) => !n.read).length
  const pending = visible.filter((r) => r.status === 'PendingApproval')
  const overdueApprovals = pending.filter((r) => new Date(r.due).getTime() < Date.now())
  const expiringDocs = visible
    .filter(
      (r) =>
        !['Archived', 'Disposed', 'Superseded'].includes(r.status) &&
        new Date(r.expiry).getTime() - Date.now() < 30 * 86400000,
    )
    .sort((a, b) => a.expiry.localeCompare(b.expiry))
  const attention = overdueApprovals.length + expiringDocs.length
  const allowedNav = navigation
    .filter((n) => n.id !== 'admin' || rights.admin)
    .filter((n) => n.id !== 'audit' || rights.audit)
  const heading = navigation.find((n) => n.id === page)
  const switchRole = (next: Role) => {
    setRole(next)
    setPage('home')
    setSelectedId('')
    setState((s) => addEvent(s, next, 'Demo persona selected', 'تم اختيار دور العرض', next))
  }
  const saveRecord = (record: RecordItem) => {
    setState((s) =>
      addEvent(
        {
          ...s,
          documents: s.documents.some((r) => r.id === record.id)
            ? s.documents.map((r) => (r.id === record.id ? record : r))
            : [record, ...s.documents],
          notices: [
            {
              id: crypto.randomUUID(),
              title: `${record.status === 'PendingApproval' ? 'Approval requested' : 'Registration saved'} · ${record.title}`,
              titleAr: `تم حفظ التسجيل · ${record.titleAr || record.title}`,
              recordId: record.id,
              category: 'Registration',
              read: false,
              time: new Date().toISOString(),
            },
            ...s.notices,
          ],
        },
        role,
        record.status === 'PendingApproval' ? 'Registered, classified and submitted' : 'Registration saved',
        'تم حفظ تسجيل الوثيقة',
        record.id,
      ),
    )
    setRegistration(false)
    setPage(record.status === 'PendingApproval' ? 'workflows' : 'documents')
    if (record.status === 'PendingApproval') resetApprovalView()
    notify(t(`Saved ${record.id}`, `تم حفظ ${record.id}`))
  }
  const publishApprovedOriginal = async (record: RecordItem) => {
    if (!record.sharePointFolderId) {
      notify(t('No SharePoint folder is selected for this document.', 'لم يتم اختيار مجلد شيربوينت لهذه الوثيقة.'))
      return
    }
    try {
      const original = await readOriginal(record.id)
      const file = original || new File(
        [
          `DOCAYA APPROVED DOCUMENT\n${record.id}\n${record.title}\n\n${record.summary}\n`,
        ],
        ensureApprovedSharePointFileName(record, 'txt'),
        { type: 'text/plain' },
      )
      const uploadName = ensureApprovedSharePointFileName(record, file.name.split('.').pop() || 'txt')
      const approvedFile = new File([file], uploadName, { type: file.type || record.mime || 'application/octet-stream' })
      const response = await uploadSharePointFile(approvedFile, record.sharePointFolderId, {
        documentId: record.id,
        documentTitle: record.title,
        documentVersion: record.version || 1,
      })
      setState((current) =>
        addEvent(
          {
            ...current,
            documents: current.documents.map((item) =>
              item.id === record.id
                ? {
                    ...item,
                    location: `SharePoint / ${record.sharePointFolderName || 'Selected folder'} / ${response.item.name}`,
                    note: `Submitted to SharePoint as ${response.item.name}`,
                    staging: 'Published / Approved',
                  }
                : item,
            ),
          },
          role,
          'Approved original submitted to SharePoint',
          'تم إرسال الأصل المعتمد إلى شيربوينت',
          record.id,
          response.item.name,
        ),
      )
      notify(t('Approved document submitted to SharePoint.', 'تم إرسال الوثيقة المعتمدة إلى شيربوينت.'))
    } catch (problem) {
      notify(
        problem instanceof Error
          ? problem.message
          : t('SharePoint submission failed. Please try again.', 'فشل الإرسال إلى شيربوينت. حاول مرة أخرى.'),
      )
    }
  }
  const act = (action: string, reason = '', step = '') => {
    try {
      const before = state.documents.find((record) => record.id === selectedId)
      const next = changeRecord(state, selectedId, role, department, action, reason, step)
      setState(next)
      notify(t('Document updated. The activity trail has been recorded.', 'تم تحديث الوثيقة وتسجيل النشاط.'))
      const after = next.documents.find((record) => record.id === selectedId)
      if (
        action === 'approve' &&
        before?.status === 'PendingApproval' &&
        after?.status === 'Published'
      ) {
        void publishApprovedOriginal(after)
      }
      return true
    } catch (e) {
      notify(e instanceof Error ? e.message : String(e))
      return false
    }
  }
  const enter = async () => {
    try {
      await developmentSignIn()
    } catch {
      notify(
        t(
          'Backend API is not available. SharePoint actions will not run until the API is started and configured.',
          'واجهة الخادم الخلفية غير متاحة. لن تعمل إجراءات شيربوينت حتى يتم تشغيلها وتهيئتها.',
        ),
      )
    }
    sessionStorage.setItem('docaya-prototype-session', 'yes')
    setSignedIn(true)
    if (state.settings.defaultArabic) setLang('ar')
  }
  if (!signedIn)
    return (
      <div className="p-app p-signin">
        <div className="p-signin-story">
          <Brand />
          <div>
            <p className="p-eyebrow">
              {t('A NEW CHAPTER IN DOCUMENT INTELLIGENCE', 'فصل جديد في ذكاء الوثائق')}
            </p>
            <h1>
              {t('Every document.', 'كل وثيقة.')}
              <br />
              {t('A clear purpose.', 'غاية واضحة.')}
              <br />
              <em>{t('An intelligent future.', 'ومستقبل ذكي.')}</em>
            </h1>
            <p>
              {t(
                'Bring your records, decisions and institutional knowledge into one governed, bilingual workspace.',
                'اجمع سجلاتك وقراراتك ومعرفتك المؤسسية في مساحة موحدة ومحكومة وثنائية اللغة.',
              )}
            </p>
            <div className="p-signin-journey">
              {[
                t('Capture', 'التقاط'),
                t('Classify', 'تصنيف'),
                t('Approve', 'اعتماد'),
                t('Discover', 'اكتشاف'),
              ].map((v, i) => (
                <span key={v}>
                  <b>0{i + 1}</b>
                  {v}
                </span>
              ))}
            </div>
          </div>
          <small>
            {t('Designed for the UAE · Arabic & English', 'مصمم لدولة الإمارات · العربية والإنجليزية')}
          </small>
        </div>
        <main className="p-signin-form">
          <button className="p-btn p-language" onClick={() => setLang(lang === 'en' ? 'ar' : 'en')}>
            <Globe2 size={17} />
            {lang === 'en' ? 'العربية' : 'English'}
          </button>
          <div>
            <img
              className="d-welcome-art"
              src={`${import.meta.env.BASE_URL}brand/docaya-welcome.png`}
              width={260}
              height={173}
              alt=""
              draggable={false}
            />
            <span className="p-badge internal">{t('CLIENT EXPERIENCE PROTOTYPE', 'نموذج تجربة العميل')}</span>
            <h2>{t('Welcome to Docaya', 'مرحباً بك في دوكايا')}</h2>
            <p className="p-ar-brand" lang="ar">
              دوكايا · ذكاء إدارة الوثائق
            </p>
            <p>
              {t(
                'Explore the complete document lifecycle with realistic sample records and interactive workflows.',
                'استكشف دورة حياة الوثائق كاملة بسجلات تجريبية واقعية ومسارات تفاعلية.',
              )}
            </p>
            <Field label={t('Choose your demo persona', 'اختر دورك في العرض')}>
              <select value={role} onChange={(e) => setRole(e.target.value as Role)}>
                {roles.map((r) => (
                  <option key={r} value={r}>
                    {local(r, lang)}
                  </option>
                ))}
              </select>
            </Field>
            <button className="p-btn primary p-full" onClick={enter}>
              <ShieldCheck size={19} />
              {t('Enter demo workspace', 'الدخول إلى مساحة العرض')}
              <ArrowRight size={18} />
            </button>
            <div className="p-signin-divider">
              <span>{t('Planned sign-in experience', 'تجربة الدخول المخططة')}</span>
            </div>
            <button className="p-btn p-full" onClick={enter}>
              <ShieldCheck size={20} />
              {t('UAE PASS · simulated sign-in', 'الهوية الرقمية · دخول تجريبي')}
            </button>
            <p className="p-subtle p-center">
              {t(
                'Sample data only. No UAE PASS, Entra, SharePoint or AI service is connected. Changes are stored in this browser.',
                'بيانات تجريبية فقط. لا توجد خدمات هوية أو شيربوينت أو ذكاء اصطناعي متصلة. تحفظ التغييرات في المتصفح.',
              )}
            </p>
          </div>
          <small>DOCAYA © 2026 · {t('Document Management Intelligence', 'ذكاء إدارة الوثائق')}</small>
        </main>
      </div>
    )
  return (
    <div
      className={`p-app ${page === 'workflows' ? 'd-approval-page' : ''} ${selected && page === 'workflows' && !registration ? 'd-has-review' : ''}`}
    >
      <a className="skip-link" href="#prototype-main">
        {t('Skip to content', 'انتقل إلى المحتوى')}
      </a>
      <aside className={`p-sidebar ${mobile ? 'open' : ''}`}>
        <Brand />
        <button
          className="p-mobile-close p-icon"
          aria-label={t('Close navigation', 'إغلاق التنقل')}
          onClick={() => setMobile(false)}
        >
          <X />
        </button>
        <div className="p-workspace-label">
          <span className="p-status-dot" />
          <div>
            {t('UAE · MOI demonstration', 'الإمارات · عرض الداخلية')}
            <small>{t('Enterprise workspace', 'مساحة العمل المؤسسية')}</small>
          </div>
          <span className="p-demo-chip">DEMO</span>
        </div>
        <nav>
          {[0, 1, 2, 3].map((group) => (
            <div className="p-nav-group" key={group}>
              <p>
                {
                  [
                    t('WORKSPACE', 'مساحة العمل'),
                    t('DOCUMENT LIFECYCLE', 'دورة حياة الوثائق'),
                    t('INTELLIGENCE & GOVERNANCE', 'الذكاء والحوكمة'),
                    t('CONFIGURATION', 'الإعدادات'),
                  ][group]
                }
              </p>
              {allowedNav
                .filter((n) => n.group === group)
                .map(({ id, en, ar, icon: Icon }) => (
                  <button
                    key={id}
                    aria-current={page === id ? 'page' : undefined}
                    className={page === id ? 'active' : ''}
                    onClick={() => navigate(id)}
                  >
                    <Icon size={18} />
                    <span>{t(en, ar)}</span>
                    {id === 'workflows' && pending.length > 0 && <b>{pending.length}</b>}
                  </button>
                ))}
            </div>
          ))}
        </nav>
        <div className="p-sidebar-bottom">
          <div>
            <ShieldCheck size={20} />
            <span>
              {t('Intelligence with accountability', 'ذكاء مع مسؤولية')}
              <small>{t('Arabic-first. Governance-led.', 'العربية أولاً. الحوكمة أساساً.')}</small>
            </span>
          </div>
          <button onClick={() => navigate('workshop')}>
            <CircleHelp size={16} />
            {t('Demo guide & feedback', 'دليل العرض والملاحظات')}
            <ChevronRight size={15} />
          </button>
        </div>
      </aside>
      {mobile && <div className="p-scrim" onClick={() => setMobile(false)} />}
      <div className="p-main">
        <header className="p-topbar">
          <button
            className="p-icon p-menu"
            aria-label={t('Open navigation', 'فتح التنقل')}
            onClick={() => setMobile(true)}
          >
            <Menu />
          </button>
          <div className="p-breadcrumb">
            {t('Workspace', 'مساحة العمل')}
            <ChevronRight size={13} />
            <strong>
              {page === 'notifications'
                ? t('Notifications', 'الإشعارات')
                : heading
                  ? t(heading.en, heading.ar)
                  : ''}
            </strong>
          </div>
          <form
            className="p-global-search"
            onSubmit={(e) => {
              e.preventDefault()
              setSelectedId('')
              setPage('search')
            }}
          >
            <Search size={16} />
            <input
              value={globalQuery}
              onChange={(e) => setGlobalQuery(e.target.value)}
              placeholder={t('Search approved knowledge…', 'ابحث في المعرفة المعتمدة…')}
              aria-label={t('Global search', 'البحث العام')}
            />
            <kbd>↵</kbd>
          </form>
          <div className="p-topbar-actions">
            <div className="d-moitag">
              <span aria-hidden="true">
                <i />
                <i />
                <i />
                <i />
              </span>
              <span>
                <strong>{t('UAE · Ministry of Interior', 'الإمارات · وزارة الداخلية')}</strong>
                <small>{t('Synthetic demonstration', 'عرض ببيانات تجريبية')}</small>
              </span>
            </div>
            <button className="p-language" onClick={() => setLang(lang === 'en' ? 'ar' : 'en')}>
              <Globe2 size={17} />
              {lang === 'en' ? 'العربية' : 'English'}
            </button>
            <button
              className="p-icon p-notification-button"
              aria-label={t('Notifications', 'الإشعارات')}
              onClick={() => navigate('notifications')}
            >
              <Bell size={19} />
              {unread + attention > 0 && <i>{unread + attention}</i>}
            </button>
            <span className="p-avatar">
              {people[role]
                .split(' ')
                .slice(0, 2)
                .map((s) => s[0])
                .join('')}
            </span>
            <button
              className="p-icon"
              aria-label={t('Sign out', 'تسجيل الخروج')}
              onClick={() => {
                sessionStorage.removeItem('docaya-prototype-session')
                setSignedIn(false)
              }}
            >
              <LogOut size={17} />
            </button>
          </div>
        </header>
        <div className="p-demo-bar">
          <span>
            <i />
            {t('Interactive prototype', 'نموذج تفاعلي')}
            <small>
              {t('Synthetic MOI data · integrations simulated', 'بيانات داخلية تجريبية · تكاملات محاكاة')}
            </small>
          </span>
          <div>
            <label>
              {t('View as', 'العرض بصفة')}
              <select
                aria-label={t('Demo persona', 'دور العرض')}
                value={role}
                onChange={(e) => switchRole(e.target.value as Role)}
              >
                {roles.map((r) => (
                  <option key={r} value={r}>
                    {local(r, lang)}
                  </option>
                ))}
              </select>
            </label>
            {page !== 'admin' && (
              <select
                aria-label={t('Demo department', 'إدارة العرض')}
                value={department}
                onChange={(e) => {
                  setDepartment(e.target.value)
                  setSelectedId('')
                }}
              >
                {departments.map((v) => (
                  <option key={v} value={v}>
                    {local(v, lang)}
                  </option>
                ))}
              </select>
            )}
          </div>
        </div>
        <main id="prototype-main" className="p-content" tabIndex={-1}>
          {storageError && (
            <div className="p-error" role="alert">
              {storageError}
            </div>
          )}
          <div className="p-page-heading">
            <div>
              <p className="p-eyebrow">
                {page === 'home'
                  ? new Intl.DateTimeFormat(lang === 'ar' ? 'ar-AE' : 'en-AE', {
                      weekday: 'long',
                      day: 'numeric',
                      month: 'long',
                      year: 'numeric',
                    }).format(new Date())
                  : t('DOCAYA / DOCUMENT MANAGEMENT INTELLIGENCE', 'دوكايا / ذكاء إدارة الوثائق')}
              </p>
              <h1>
                {page === 'home'
                  ? t(`Good to see you, ${people[role].split(' ')[0]}`, 'مرحباً بك في دوكايا')
                  : page === 'notifications'
                    ? t('Notification center', 'مركز الإشعارات')
                    : heading
                      ? t(heading.en, heading.ar)
                      : ''}
              </h1>
              <p>
                {page === 'home'
                  ? t(
                      'A clear view of your documents, decisions and what comes next.',
                      'رؤية واضحة لوثائقك وقراراتك وما يأتي بعدها.',
                    )
                  : page === 'workflows'
                    ? t(
                        'Review documents and make a decision, one at a time.',
                        'راجع الوثائق واتخذ قرارك لكل وثيقة على حدة.',
                      )
                    : page === 'workshop'
                      ? t(
                          'Walk through the experience. Shape what comes next together.',
                          'استعرض التجربة وشارك في تحديد الخطوات التالية.',
                        )
                      : t(
                          'One connected workspace for governed information.',
                          'مساحة عمل مترابطة للمعلومات المحكومة.',
                        )}
              </p>
            </div>
            {rights.register && ['home', 'documents'].includes(page) && (
              <button className="p-btn primary" onClick={() => setRegistration('new')}>
                <Plus size={18} />
                {t('Register document', 'تسجيل وثيقة')}
              </button>
            )}
          </div>
          {page === 'home' && (
            <>
              <div className="p-four p-metrics">
                {[
                  [
                    FolderOpen,
                    visible.length,
                    t('Registered documents', 'وثائق مسجلة'),
                    t('A complete chain of custody', 'سلسلة عهدة متكاملة'),
                    'documents',
                  ],
                  [
                    Workflow,
                    pending.length,
                    t('Awaiting approval', 'بانتظار الاعتماد'),
                    t('Decisions that keep work moving', 'قرارات تدفع العمل للأمام'),
                    'workflows',
                  ],
                  [
                    ShieldCheck,
                    visible.filter((r) => r.status === 'Published').length,
                    t('Published records', 'سجلات منشورة'),
                    t('Approved and ready to discover', 'معتمدة وجاهزة للاكتشاف'),
                    'search',
                  ],
                  [
                    LockKeyhole,
                    visible.filter((r) => r.hold).length,
                    t('Protected by legal hold', 'محمية بحجز قانوني'),
                    t('Retention controls in effect', 'ضوابط الحفظ سارية'),
                    'records',
                  ],
                ].map(([Icon, count, label, sub, target]) => {
                  const MetricIcon = Icon as typeof FolderOpen
                  return (
                    <button
                      className="p-card p-metric"
                      key={String(label)}
                      onClick={() => navigate(target as Page)}
                    >
                      <div>
                        <span>{String(label)}</span>
                        <MetricIcon size={20} />
                      </div>
                      <strong>{String(count)}</strong>
                      <small>
                        {String(sub)}
                        <ArrowRight size={14} />
                      </small>
                    </button>
                  )
                })}
              </div>
              <div className="p-home-grid">
                <section className="p-hero">
                  <div className="p-hero-grid" />
                  <div className="p-hero-content">
                    <p className="p-eyebrow">
                      {t('GOVERNED FROM THE VERY FIRST STEP', 'محكوم من الخطوة الأولى')}
                    </p>
                    <h2>
                      {t('Your documents.', 'وثائقك.')}
                      <br />
                      <em>{t('Connected. Controlled. Intelligent.', 'مترابطة. محكومة. ذكية.')}</em>
                    </h2>
                    <p>
                      {t(
                        'From the first upload to the right answer. Turn everyday documents into trusted institutional knowledge.',
                        'من أول رفع إلى الإجابة الصحيحة. حوّل وثائق العمل اليومية إلى معرفة مؤسسية موثوقة.',
                      )}
                    </p>
                    <button className="p-btn light" onClick={() => navigate('workshop')}>
                      <BookOpen size={17} />
                      {t('Explore the demo journey', 'استكشف رحلة العرض')}
                      <ArrowRight size={17} />
                    </button>
                  </div>
                  <div className="p-hero-seal">
                    <span>د</span>
                    <small>DOCAYA</small>
                  </div>
                  <div className="p-hero-footer">
                    <span>
                      <ShieldCheck size={15} />
                      {t('Governance by design', 'الحوكمة بالتصميم')}
                    </span>
                    <span>
                      <Globe2 size={15} />
                      {t('Arabic & English', 'العربية والإنجليزية')}
                    </span>
                    <span>
                      <FileText size={15} />
                      {t('Original formats preserved', 'حفظ الصيغ الأصلية')}
                    </span>
                  </div>
                </section>
                <section className="p-card p-action-queue">
                  <div className="p-card-head">
                    <div>
                      <p className="p-eyebrow">{t('NEEDS YOUR ATTENTION', 'يتطلب انتباهك')}</p>
                      <h3>{t('The next right action', 'الإجراء التالي')}</h3>
                    </div>
                    <span className="p-count">{pending.length}</span>
                  </div>
                  {pending.slice(0, 3).map((r) => (
                    <button
                      className="p-queue-item"
                      key={r.id}
                      onClick={() => {
                        open(r)
                      }}
                    >
                      <span className="p-queue-icon">
                        <FileText size={19} />
                      </span>
                      <span>
                        <strong>{lang === 'ar' ? r.titleAr : r.title}</strong>
                        <small>
                          {local(r.department, lang)} · {local(r.flow, lang)}
                        </small>
                        <em className={new Date(r.due).getTime() < Date.now() ? 'overdue' : ''}>
                          {new Date(r.due).getTime() < Date.now()
                            ? t('SLA overdue', 'تجاوز المدة المحددة')
                            : t('Awaiting your review', 'بانتظار المراجعة')}
                        </em>
                      </span>
                      <ChevronRight size={15} />
                    </button>
                  ))}
                  {!pending.length && <Empty t={t} />}
                  <button className="p-link" onClick={() => navigate('workflows')}>
                    {t('View approval queue', 'عرض قائمة الاعتماد')}
                    <ArrowRight size={15} />
                  </button>
                </section>
              </div>
              <section className="p-card p-lifecycle-card">
                <div className="p-card-head">
                  <div>
                    <p className="p-eyebrow">{t('THE DOCAYA LIFECYCLE', 'دورة حياة دوكايا')}</p>
                    <h3>{t('Every step accounted for', 'كل خطوة موثقة')}</h3>
                  </div>
                  <span className="p-subtle">
                    {t(
                      'AI suggestions → human decisions → trusted knowledge',
                      'اقتراحات ذكية ← قرارات بشرية ← معرفة موثوقة',
                    )}
                  </span>
                </div>
                <div className="p-lifecycle">
                  {[
                    ['Capture', 'الالتقاط', 'documents', CloudUpload],
                    ['Classify', 'التصنيف', 'documents', FolderOpen],
                    ['Approve', 'الاعتماد', 'workflows', ShieldCheck],
                    ['Publish', 'النشر', 'operations', ArrowDownToLine],
                    ['Discover', 'الاكتشاف', 'search', DocayaAIMark],
                  ].map(([en, ar, target, Icon], i) => {
                    const StepIcon = Icon as typeof CloudUpload | typeof DocayaAIMark
                    return (
                      <button key={String(en)} onClick={() => navigate(target as Page)}>
                        <span>
                          <StepIcon size={21} />
                        </span>
                        <div>
                          <small>0{i + 1}</small>
                          <strong>{t(String(en), String(ar))}</strong>
                        </div>
                        <ChevronRight size={18} />
                      </button>
                    )
                  })}
                </div>
              </section>
              <div className="p-card no-pad">
                <div className="p-card-head">
                  <h3>{t('Recently registered', 'المسجلة مؤخراً')}</h3>
                  <button className="p-link" onClick={() => navigate('documents')}>
                    {t('View all documents', 'عرض كل الوثائق')}
                    <ArrowRight size={15} />
                  </button>
                </div>
                <DocumentTable
                  rows={[...visible].sort((a, b) => b.captured.localeCompare(a.captured)).slice(0, 4)}
                  lang={lang}
                  t={t}
                  onOpen={open}
                  selectedId={selected?.id || ''}
                />
              </div>
            </>
          )}
          {page === 'documents' && <RegisterPage {...props} />}
          {page === 'library' && <DocumentLibraryPage {...props} />}
          {page === 'workflows' && (
            <ApprovalQueue
              rows={reviewQueue}
              pendingCount={pending.length}
              returnedCount={visible.filter((record) => record.status === 'Returned').length}
              selectedId={selected?.id || ''}
              onOpen={open}
              lang={lang}
              t={t}
              query={approvalQuery}
              onQuery={setApprovalQuery}
              view={approvalView}
              onView={(view) => {
                setApprovalView(view)
                setApprovalOverdue(false)
                setSelectedId('')
              }}
              overdueOnly={approvalOverdue}
              onOverdue={setApprovalOverdue}
            />
          )}
          {page === 'operations' && <Operations {...props} />}
          {page === 'correspondence' && <CorrespondencePage {...props} />}
          {page === 'records' && <RecordsPage {...props} />}
          {page === 'reports' && <ReportsPage {...props} />}
          {page === 'audit' && rights.audit && <AuditPage {...props} />}
          {page === 'admin' && rights.admin && <AdminPage key={role} {...props} />}
          {page === 'search' && <SearchPage {...props} initialQuery={globalQuery} />}
          {page === 'notifications' && (
            <>
              <div className="p-four p-metrics">
                {[
                  [
                    Workflow,
                    pending.length,
                    t('Pending my approval', 'بانتظار اعتمادي'),
                    t('Decisions waiting on you', 'قرارات بانتظارك'),
                    () => navigate('workflows'),
                  ],
                  [
                    AlertTriangle,
                    overdueApprovals.length,
                    t('Overdue approvals', 'اعتمادات متأخرة'),
                    t('Due date passed, still not approved', 'انقضى الموعد ولم تُعتمد بعد'),
                    () => {
                      navigate('workflows')
                      setApprovalOverdue(true)
                    },
                  ],
                  [
                    CalendarClock,
                    expiringDocs.length,
                    t('Expiring within 30 days', 'تنتهي خلال ٣٠ يوماً'),
                    t('Review or renew before expiry', 'راجع أو جدّد قبل الانتهاء'),
                    () => navigate('records'),
                  ],
                  [
                    Bell,
                    unread,
                    t('Unread notifications', 'إشعارات غير مقروءة'),
                    t('System and workflow updates', 'تحديثات النظام ومسارات العمل'),
                    undefined,
                  ],
                ].map(([Icon, count, label, sub, go]) => {
                  const MetricIcon = Icon as typeof Bell
                  return (
                    <button
                      className="p-card p-metric"
                      key={String(label)}
                      onClick={go as (() => void) | undefined}
                    >
                      <div>
                        <span>{String(label)}</span>
                        <MetricIcon size={20} />
                      </div>
                      <strong>{String(count)}</strong>
                      <small>
                        {String(sub)}
                        {go ? <ArrowRight size={14} /> : null}
                      </small>
                    </button>
                  )
                })}
              </div>
              {attention > 0 && (
                <div className="p-card no-pad">
                  <div className="p-card-head">
                    <div>
                      <p className="p-eyebrow">{t('NEEDS ATTENTION', 'يتطلب الانتباه')}</p>
                      <h3>{t('Important highlights', 'أبرز التنبيهات المهمة')}</h3>
                    </div>
                    <span className="p-count">{attention}</span>
                  </div>
                  {overdueApprovals.map((r) => (
                    <div className="p-notice-row alert" key={`overdue-${r.id}`}>
                      <span className="p-queue-icon alert">
                        <AlertTriangle size={18} />
                      </span>
                      <div>
                        <strong>
                          {t('Approval overdue:', 'اعتماد متأخر:')} {lang === 'ar' ? r.titleAr : r.title}
                        </strong>
                        <small>
                          {local(r.department, lang)} · {t('Due', 'الاستحقاق')}{' '}
                          {new Date(r.due).toLocaleString(lang)}
                        </small>
                        <button className="p-link" onClick={() => open(r)}>
                          {t('Open document', 'فتح الوثيقة')}
                        </button>
                      </div>
                    </div>
                  ))}
                  {expiringDocs.map((r) => {
                    const days = Math.ceil((new Date(r.expiry).getTime() - Date.now()) / 86400000)
                    return (
                      <div className="p-notice-row alert" key={`expiry-${r.id}`}>
                        <span className="p-queue-icon alert">
                          <CalendarClock size={18} />
                        </span>
                        <div>
                          <strong>
                            {days < 0
                              ? t('Document expired:', 'انتهت صلاحية الوثيقة:')
                              : t('Expiring soon:', 'قرب انتهاء الصلاحية:')}{' '}
                            {lang === 'ar' ? r.titleAr : r.title}
                          </strong>
                          <small>
                            {local(r.department, lang)} ·{' '}
                            {days < 0
                              ? t(`Expired ${Math.abs(days)} day(s) ago`, `انتهت منذ ${Math.abs(days)} يوم`)
                              : t(`${days} day(s) remaining`, `متبقٍ ${days} يوم`)}{' '}
                            · {new Date(r.expiry).toLocaleDateString(lang)}
                          </small>
                          <button className="p-link" onClick={() => open(r)}>
                            {t('Open document', 'فتح الوثيقة')}
                          </button>
                        </div>
                      </div>
                    )
                  })}
                </div>
              )}
              <div className="p-toolbar">
                <span>
                  {unread} {t('unread notifications', 'إشعارات غير مقروءة')}
                </span>
                <button
                  className="p-btn"
                  onClick={() =>
                    setState((s) => ({
                      ...s,
                      notices: s.notices.map((n) =>
                        notices.some((v) => v.id === n.id) ? { ...n, read: true } : n,
                      ),
                    }))
                  }
                >
                  <Check size={16} />
                  {t('Mark all read', 'قراءة الكل')}
                </button>
              </div>
              <div className="p-card no-pad">
                {notices.map((n) => (
                  <div className={`p-notice-row ${n.read ? '' : 'unread'}`} key={n.id}>
                    <span className="p-queue-icon">
                      <Bell size={18} />
                    </span>
                    <div>
                      <strong>{lang === 'ar' ? n.titleAr : n.title}</strong>
                      <small>
                        {local(n.category, lang)} · {new Date(n.time).toLocaleString(lang)}
                      </small>
                      {n.recordId && (
                        <button
                          className="p-link"
                          onClick={() => {
                            const r = visible.find((r) => r.id === n.recordId)
                            if (r) open(r)
                          }}
                        >
                          {t('Open document', 'فتح الوثيقة')}
                        </button>
                      )}
                    </div>
                    <button
                      className="p-btn"
                      onClick={() =>
                        setState((s) => ({
                          ...s,
                          notices: s.notices.map((v) => (v.id === n.id ? { ...v, read: !v.read } : v)),
                        }))
                      }
                    >
                      {n.read ? t('Mark unread', 'غير مقروء') : t('Mark read', 'مقروء')}
                    </button>
                  </div>
                ))}
                {!notices.length && <Empty t={t} />}
              </div>
              <div className="p-note">
                <LockKeyhole size={17} />
                {t(
                  'Security alerts are mandatory. Email and Teams delivery, quiet hours and push channels are simulated. Administrators can adjust demo preferences in Admin center.',
                  'تنبيهات الأمان إلزامية. البريد وتيمز وساعات الهدوء وقنوات الدفع محاكاة. يمكن للمدير ضبط تفضيلات العرض في مركز الإدارة.',
                )}
              </div>
            </>
          )}
          {page === 'workshop' && <Workshop {...props} navigate={navigate} onReset={() => setReset(true)} />}
          <footer className="p-footer">
            <span>
              DOCAYA <i>/</i> {t('Document Management Intelligence', 'ذكاء إدارة الوثائق')}
            </span>
            <span>{t('Client experience prototype · v2.0', 'نموذج تجربة العميل · الإصدار ٢.٠')}</span>
          </footer>
        </main>
      </div>
      {toast && (
        <div className="p-toast" role="status">
          <Check size={19} />
          <span>{toast}</span>
          <button className="p-icon" aria-label={t('Dismiss', 'إغلاق')} onClick={() => setToast('')}>
            <X size={16} />
          </button>
        </div>
      )}
      {registration && (
        <Registration
          state={{ ...state, settings: { ...state.settings, department } }}
          role={role}
          lang={lang}
          t={t}
          existing={registration === 'new' ? undefined : registration}
          onClose={() => setRegistration(false)}
          onSave={saveRecord}
        />
      )}
      {selected && !registration && (
        <RecordDetail
          key={selected.id}
          record={selected}
          state={state}
          role={role}
          lang={lang}
          t={t}
          onClose={() => setSelectedId('')}
          onAction={act}
          queue={page === 'workflows' ? reviewQueue : visible}
          onSelect={open}
          reviewMode={page === 'workflows'}
          onEdit={() => {
            setRegistration(selected)
            setSelectedId('')
          }}
          onLog={(action, ar) => setState((s) => addEvent(s, role, action, ar, selected.id))}
        />
      )}
      {!registration && !reset && !(selected && page === 'workflows') && (
        <AssistantDock
          visible={visible}
          lang={lang}
          t={t}
          context={selected}
          onOpen={open}
          onSearch={(query) => {
            setSelectedId('')
            setGlobalQuery(query)
            setPage('search')
          }}
        />
      )}
      {reset && (
        <Modal
          title={t('Reset the demo workspace?', 'إعادة ضبط مساحة العرض؟')}
          onClose={() => setReset(false)}
        >
          <div className="p-modal-body">
            <p>
              {t(
                'This clears prototype changes, uploaded files and workshop notes from this browser, then restores the sample records. Export your notes first.',
                'يمسح تغييرات النموذج والملفات والملاحظات من المتصفح ثم يعيد السجلات التجريبية. صدّر ملاحظاتك أولاً.',
              )}
            </p>
            <div className="p-actions">
              <button className="p-btn" onClick={() => setReset(false)}>
                {t('Cancel', 'إلغاء')}
              </button>
              <button
                className="p-btn danger"
                onClick={() => {
                  void clearOriginals()
                    .then(() => {
                      setState(seedWorkspace())
                      setReset(false)
                      setPage('home')
                      notify(t('Demo restored.', 'تمت استعادة العرض.'))
                    })
                    .catch(() =>
                      notify(
                        t(
                          'Could not clear stored files. Please retry.',
                          'تعذر مسح الملفات المحفوظة. أعد المحاولة.',
                        ),
                      ),
                    )
                }}
              >
                {t('Reset demo', 'إعادة ضبط العرض')}
              </button>
            </div>
          </div>
        </Modal>
      )}
    </div>
  )
}

function ensureApprovedSharePointFileName(record: RecordItem, extension: string) {
  const cleanedExtension = extension.replace(/[^a-z0-9]/gi, '').toLowerCase() || 'txt'
  const slug = record.title
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/gi, '-')
    .replace(/^-+|-+$/g, '')
    .toLowerCase()
    .slice(0, 80) || 'document'
  return `DOCAYA__${record.id}__v${record.version || 1}__${slug}.${cleanedExtension}`
}
function Brand() {
  return (
    <div className="p-brand">
      <span className="p-brand-mark">
        <DocayaMark size={60} />
      </span>
      <div>
        <strong>
          <span lang="en">Docaya</span> <b lang="ar">دوكايا</b>
        </strong>
        <small>Document intelligence</small>
      </div>
    </div>
  )
}
function RegisterPage({ visible, lang, t, onOpen, selectedId }: PageProps) {
  const [query, setQuery] = useState('')
  const [status, setStatus] = useState('all')
  const [department, setDepartment] = useState('all')
  const rows = visible.filter(
    (r) =>
      (status === 'all' || r.status === status) &&
      (department === 'all' || r.department === department) &&
      `${r.title} ${r.titleAr} ${r.id} ${r.reference}`.toLowerCase().includes(query.toLowerCase()),
  )
  return (
    <>
      <div className="p-toolbar">
        <div className="p-input-icon">
          <Search size={17} />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            aria-label={t('Search register', 'بحث السجل')}
            placeholder={t(
              'Search by title, reference or tracking number…',
              'ابحث بالعنوان أو المرجع أو رقم التتبع…',
            )}
          />
        </div>
        <select
          aria-label={t('Status filter', 'تصفية الحالة')}
          value={status}
          onChange={(e) => setStatus(e.target.value)}
        >
          <option value="all">{t('All statuses', 'كل الحالات')}</option>
          {[
            'Draft',
            'Classified',
            'PendingApproval',
            'Returned',
            'Approved',
            'Published',
            'Archived',
            'Disposed',
          ].map((v) => (
            <option value={v} key={v}>
              {local(v, lang)}
            </option>
          ))}
        </select>
        <select
          aria-label={t('Department filter', 'تصفية الإدارة')}
          value={department}
          onChange={(e) => setDepartment(e.target.value)}
        >
          <option value="all">{t('All departments', 'كل الإدارات')}</option>
          {departments.map((v) => (
            <option value={v} key={v}>
              {local(v, lang)}
            </option>
          ))}
        </select>
      </div>
      <div className="p-card no-pad">
        <div className="p-card-head">
          <h3>{t('Controlled document register', 'سجل الوثائق المحكومة')}</h3>
          <span className="p-subtle">
            {rows.length} {t('documents', 'وثائق')}
          </span>
        </div>
        <DocumentTable rows={rows} lang={lang} t={t} onOpen={onOpen} selectedId={selectedId} />
      </div>
    </>
  )
}

function Workshop({
  state,
  setState,
  t,
  navigate,
  onReset,
  notify,
}: PageProps & { navigate: (p: Page) => void; onReset: () => void }) {
  const [module, setModule] = useState('M1')
  const [note, setNote] = useState('')
  const [priority, setPriority] = useState('Must have')
  return (
    <>
      <section className="p-dark-banner">
        <div>
          <p className="p-eyebrow">{t('A CONVERSATION, NOT JUST A PRESENTATION', 'حوار وليس مجرد عرض')}</p>
          <h2>{t('Build the right Docaya. Together.', 'نبني دوكايا المناسبة. معاً.')}</h2>
          <p>
            {t(
              'Explore all 15 modules, try the lifecycle, and capture the client’s expectations while they are fresh.',
              'استكشف الوحدات الخمس عشرة وجرّب دورة الحياة وسجل توقعات العميل أثناء الحوار.',
            )}
          </p>
        </div>
        <button className="p-btn light" onClick={onReset}>
          {t('Reset demo data', 'إعادة البيانات التجريبية')}
        </button>
      </section>
      <div className="p-note">
        {t(
          'Suggested 15-minute walkthrough: register a sample → classify → approve → publish → run indexing → ask a question → apply a legal hold → switch to Viewer → capture feedback. System Administrator can demonstrate every step.',
          'عرض مقترح خلال ١٥ دقيقة: سجل عينة ← صنف ← اعتمد ← انشر ← فهرس ← اسأل ← طبق حجزاً ← بدل إلى قارئ ← سجل الملاحظات. يمكن لمدير النظام عرض جميع الخطوات.',
        )}
      </div>
      <div className="p-workshop-grid">
        <section>
          <div className="p-module-list">
            {modules.map(([id, en, ar, page, desc, descAr]) => (
              <button key={id} onClick={() => navigate(page as Page)}>
                <span>{id}</span>
                <div>
                  <strong>{t(en, ar)}</strong>
                  <p>{t(desc, descAr)}</p>
                </div>
                <ChevronRight size={18} />
              </button>
            ))}
          </div>
        </section>
        <aside>
          <section className="p-card p-feedback">
            <h3>{t('Capture client expectations', 'تسجيل توقعات العميل')}</h3>
            <p>
              {t(
                'What should change? What is essential for day one?',
                'ما الذي يحتاج تغييراً؟ ما الضروري منذ اليوم الأول؟',
              )}
            </p>
            <form
              onSubmit={(e) => {
                e.preventDefault()
                if (!note.trim()) return
                setState((s) => ({
                  ...s,
                  feedback: [{ id: crypto.randomUUID(), module, note: note.trim(), priority }, ...s.feedback],
                }))
                setNote('')
                notify(t('Client feedback captured.', 'تم تسجيل ملاحظات العميل.'))
              }}
            >
              <Field label={t('Module', 'الوحدة')}>
                <select value={module} onChange={(e) => setModule(e.target.value)}>
                  {modules.map(([id, en, ar]) => (
                    <option key={id} value={id}>
                      {id} · {t(en, ar)}
                    </option>
                  ))}
                </select>
              </Field>
              <Field label={t('Priority', 'الأولوية')}>
                <select value={priority} onChange={(e) => setPriority(e.target.value)}>
                  {[
                    ['Must have', 'ضروري'],
                    ['Should have', 'مهم'],
                    ['Future idea', 'فكرة مستقبلية'],
                  ].map(([en, ar]) => (
                    <option key={en} value={en}>
                      {t(en, ar)}
                    </option>
                  ))}
                </select>
              </Field>
              <Field label={t('Expectation / decision', 'التوقع / القرار')}>
                <textarea
                  required
                  rows={4}
                  value={note}
                  onChange={(e) => setNote(e.target.value)}
                  placeholder={t('Capture the client’s own words…', 'سجل كلمات العميل…')}
                />
              </Field>
              <button className="p-btn primary p-full">
                <Plus size={16} />
                {t('Save feedback', 'حفظ الملاحظة')}
              </button>
            </form>
            <button
              className="p-btn p-full"
              disabled={!state.feedback.length}
              onClick={() =>
                downloadFile(
                  'docaya-client-feedback.csv',
                  csv([
                    ['Module', 'Priority', 'Expectation'],
                    ...state.feedback.map((f) => [f.module, f.priority, f.note]),
                  ]),
                )
              }
            >
              {t('Export workshop notes', 'تصدير ملاحظات الورشة')}
            </button>
            {state.feedback.map((f) => (
              <div className="p-feedback-item" key={f.id}>
                <small>
                  {f.module} · {f.priority}
                </small>
                <p>{f.note}</p>
              </div>
            ))}
          </section>
          <section className="p-card">
            <h3>{t('Prototype boundaries', 'حدود النموذج')}</h3>
            <p>
              {t(
                'Working in this browser: navigation, registration, approvals, revisions, holds, correspondence, search, settings, reports and workshop notes.',
                'يعمل داخل المتصفح: التنقل والتسجيل والاعتماد والإصدارات والحجوزات والمراسلات والبحث والإعدادات والتقارير والملاحظات.',
              )}
            </p>
            <p>
              {t(
                'Simulated: UAE PASS / Entra, Azure SQL / Blob, SharePoint, e-signature, malware / DLP, email / Teams and AI indexing. Compliance and performance targets are not certifications.',
                'محاكاة: الهوية وإنترّا وقواعد البيانات والتخزين وشيربوينت والتوقيع والفحص والبريد وتيمز والفهرسة الذكية. أهداف الامتثال والأداء ليست شهادات.',
              )}
            </p>
            <p className="p-subtle">
              {t(
                'Production design targets: 5 GB resumable uploads, UAE data residency, trusted audit, WCAG 2.2 AA and high availability. Bulk capture, scan/email ingestion, scheduled exports and full rule designers remain future implementation work.',
                'أهداف الإنتاج: رفع حتى ٥ غيغابايت وإقامة البيانات في الإمارات وتدقيق موثوق وإتاحة عالية وإمكانية وصول. الالتقاط الجماعي والمسح والبريد والتصدير المجدول ومصممو القواعد تتطلب تنفيذاً لاحقاً.',
              )}
            </p>
          </section>
        </aside>
      </div>
    </>
  )
}
