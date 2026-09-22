export type Language = 'en' | 'ar'
export type Status =
  | 'Draft'
  | 'Classified'
  | 'PendingApproval'
  | 'Returned'
  | 'Approved'
  | 'Published'
  | 'Superseded'
  | 'Archived'
  | 'Disposed'
export type Role =
  | 'System Administrator'
  | 'Department Administrator'
  | 'Contributor'
  | 'Approver'
  | 'Records Officer'
  | 'Viewer'
  | 'Auditor'
export type Flow = 'Single' | 'Sequential' | 'Parallel' | 'Conditional'
export type Sensitivity = 'Public' | 'Internal' | 'Confidential' | 'Secret'
export type SharePointSyncMode = 'Publish approved originals' | 'Metadata only' | 'Read-only discovery'
export type RecordItem = {
  id: string
  sampleKey?: string
  title: string
  titleAr: string
  kind: string
  department: string
  sensitivity: Sensitivity
  status: Status
  sharePointFolderId?: string
  sharePointFolderName?: string
  sharePointFolderPath?: string
  owner: string
  summary: string
  summaryAr: string
  reference: string
  revises?: string
  fileName: string
  mime: string
  size: number
  hash: string
  captured: string
  expiry: string
  flow: Flow
  priority: string
  scope: string
  approvals: string[]
  version: number
  indexed: boolean
  hold: string
  declared: boolean
  location: string
  staging: string
  note: string
  due: string
  delegated: boolean
  signature: string
}
export type Notice = {
  id: string
  title: string
  titleAr: string
  recordId: string
  category: string
  read: boolean
  time: string
}
export type AuditEvent = {
  id: string
  time: string
  actor: string
  role: Role
  action: string
  actionAr: string
  object: string
  detail: string
}
export type Correspondence = {
  id: string
  subject: string
  subjectAr: string
  direction: string
  party: string
  assignee: string
  due: string
  status: string
  thread: string[]
  acknowledged: boolean
}
export type Settings = {
  classes: string[]
  department: string
  library: string
  sharePoint: {
    enabled: boolean
    tenant: string
    siteUrl: string
    libraryName: string
    publishPath: string
    syncMode: SharePointSyncMode
    versioning: boolean
    metadataPush: boolean
    aclTrim: boolean
    lastValidated: string | null
  }
  sla: number
  cleanup: string
  retentionYears: number
  cadence: string
  simulateFailure: boolean
  email: boolean
  teams: boolean
  quietHours: string
  defaultArabic: boolean
}
export type Workspace = {
  schema: 1
  seedVersion?: number
  documents: RecordItem[]
  notices: Notice[]
  audit: AuditEvent[]
  correspondence: Correspondence[]
  settings: Settings
  savedSearches: string[]
  feedback: { id: string; module: string; note: string; priority: string }[]
}
export const roles: Role[] = [
  'System Administrator',
  'Department Administrator',
  'Contributor',
  'Approver',
  'Records Officer',
  'Viewer',
  'Auditor',
]
export const departments = [
  'Operations',
  'Finance',
  'Human Resources',
  'Legal',
  'Civil Defence',
  'Traffic & Patrols',
  'Forensic Sciences',
  'Residency & Identity',
  'Strategy & Governance',
]
export const documentClasses = [
  'Correspondence',
  'Memo',
  'Circular',
  'Contract',
  'HR Record',
  'Report',
  'Policy',
  'Procedure',
  'Work Instruction',
  'Manual',
  'Form',
  'Standard',
]
export const SEED_VERSION = 2
export const sensitivities: Sensitivity[] = ['Public', 'Internal', 'Confidential', 'Secret']
export const flows: Flow[] = ['Single', 'Sequential', 'Parallel', 'Conditional']
export const defaultSharePointSettings = (): Settings['sharePoint'] => ({
  enabled: true,
  tenant: 'moi.gov.ae',
  siteUrl: 'https://moi.sharepoint.com/sites/DocayaRecords',
  libraryName: 'Official Records',
  publishPath: 'Official Records',
  syncMode: 'Publish approved originals',
  versioning: true,
  metadataPush: true,
  aclTrim: true,
  lastValidated: null,
})
export const people: Record<Role, string> = {
  'System Administrator': 'Mariam Al Mansoori',
  'Department Administrator': 'Omar Al Ali',
  Contributor: 'Noor Al Shamsi',
  Approver: 'Ahmed Al Nuaimi',
  'Records Officer': 'Fatima Al Kaabi',
  Viewer: 'Khalid Al Ali',
  Auditor: 'Sara Al Mazrouei',
}
export function capabilities(role: Role) {
  return {
    register: !['Viewer', 'Auditor'].includes(role),
    approve: ['System Administrator', 'Approver', 'Records Officer'].includes(role),
    records: ['System Administrator', 'Records Officer'].includes(role),
    admin: ['System Administrator', 'Department Administrator'].includes(role),
    audit: ['System Administrator', 'Department Administrator', 'Records Officer', 'Auditor'].includes(role),
  }
}
export function canSee(record: RecordItem, role: Role, department: string) {
  if (['System Administrator', 'Records Officer', 'Auditor'].includes(role)) return true
  if (record.sensitivity === 'Secret') return false
  if (
    record.department !== department &&
    (record.sensitivity === 'Confidential' || record.scope === 'Department-only')
  )
    return false
  if (role === 'Viewer' && !['Published', 'Archived'].includes(record.status)) return false
  return true
}
export function approvalSteps(record: RecordItem) {
  if (record.flow === 'Single') return ['Department head']
  if (record.flow === 'Conditional')
    return record.sensitivity === 'Secret' ? ['Senior director', 'Compliance officer'] : ['Department head']
  return ['Department head', 'Compliance officer']
}
export function addEvent(
  state: Workspace,
  role: Role,
  action: string,
  actionAr: string,
  object: string,
  detail = '',
): Workspace {
  return {
    ...state,
    audit: [
      {
        id: crypto.randomUUID(),
        time: new Date().toISOString(),
        actor: people[role],
        role,
        action,
        actionAr,
        object,
        detail,
      },
      ...state.audit,
    ],
  }
}
export function changeRecord(
  state: Workspace,
  id: string,
  role: Role,
  department: string,
  action: string,
  reason = '',
  step = '',
): Workspace {
  const record = state.documents.find((r) => r.id === id)
  if (!record || !canSee(record, role, department)) throw new Error('Access denied / الوصول مرفوض')
  const rights = capabilities(role)
  let changed = { ...record }
  let label = action,
    arabic = action
  if (action === 'submit') {
    if (!rights.register || !['Draft', 'Classified', 'Returned'].includes(record.status))
      throw new Error('Submission is unavailable / الإرسال غير متاح')
    if (!record.title.trim() || !record.summary.trim() || !record.kind || !record.fileName)
      throw new Error('Complete the registration fields / أكمل بيانات التسجيل')
    changed = {
      ...changed,
      status: 'PendingApproval',
      approvals: [],
      note: '',
      due: new Date(Date.now() + state.settings.sla * 3600000).toISOString(),
    }
    label = 'Submitted for approval'
    arabic = 'تم الإرسال للاعتماد'
  } else if (action === 'approve') {
    if (!rights.approve || record.status !== 'PendingApproval')
      throw new Error('Approval is unavailable / الاعتماد غير متاح')
    const steps = approvalSteps(record)
    const target = step || steps.find((s) => !record.approvals.includes(s))!
    if (!steps.includes(target) || record.approvals.includes(target))
      throw new Error('This step is already complete / هذه الخطوة مكتملة')
    if (record.flow !== 'Parallel' && target !== steps.find((s) => !record.approvals.includes(s)))
      throw new Error('Approve in sequence / اعتمد بالترتيب')
    changed.approvals = [...record.approvals, target]
    changed.signature = `${people[role]} · ${new Date().toISOString()} · DEMO-${crypto.randomUUID().slice(0, 8)}`
    label = `Approved: ${target}`
    arabic = 'تم اعتماد خطوة المراجعة'
    if (changed.approvals.length === steps.length) {
      changed.status = state.settings.simulateFailure ? 'Approved' : 'Published'
      changed.note = state.settings.simulateFailure
        ? 'Simulated promotion failure; original retained for retry.'
        : ''
      if (changed.status === 'Published') changed = publish(changed, state.settings)
      label = changed.status === 'Published' ? 'Approved and published' : 'Approved; promotion needs retry'
      arabic = changed.status === 'Published' ? 'تم الاعتماد والنشر' : 'تم الاعتماد؛ إعادة النشر مطلوبة'
    }
  } else if (action === 'return') {
    if (!rights.approve || record.status !== 'PendingApproval' || !reason.trim())
      throw new Error('A return reason is required / سبب الإرجاع مطلوب')
    changed = { ...changed, status: 'Returned', note: reason, approvals: [], indexed: false }
    label = 'Returned for changes'
    arabic = 'تم الإرجاع للتعديل'
  } else if (action === 'retry') {
    if (role !== 'System Administrator' || record.status !== 'Approved')
      throw new Error('Retry is unavailable / إعادة المحاولة غير متاحة')
    changed = publish(changed, state.settings)
    label = 'Promotion replay completed'
    arabic = 'تمت إعادة النشر'
  } else if (action === 'hold') {
    if (!rights.records || !reason.trim() || record.status === 'Disposed')
      throw new Error('A legal hold reason is required / سبب الحجز القانوني مطلوب')
    changed.hold = reason
    label = 'Legal hold applied'
    arabic = 'تم تطبيق الحجز القانوني'
  } else if (action === 'release') {
    if (!rights.records || !reason.trim())
      throw new Error('A release reason is required / سبب رفع الحجز مطلوب')
    changed.hold = ''
    label = 'Legal hold released'
    arabic = 'تم رفع الحجز القانوني'
  } else if (action === 'archive' || action === 'dispose') {
    if (
      !rights.records ||
      record.hold ||
      !reason.trim() ||
      (action === 'archive' ? record.status !== 'Published' : record.status !== 'Archived')
    )
      throw new Error(
        'Disposition blocked; check status, hold and reason / التصرف محظور؛ تحقق من الحالة والحجز والسبب',
      )
    changed = {
      ...changed,
      status: action === 'archive' ? 'Archived' : 'Disposed',
      indexed: false,
      location: action === 'archive' ? 'Archive / ' + record.kind : 'Disposed — registry retained',
    }
    label = action === 'archive' ? 'Record archived' : 'Disposition approved (simulation)'
    arabic = action === 'archive' ? 'تمت أرشفة السجل' : 'تم اعتماد الإتلاف التجريبي'
  } else if (action === 'declare') {
    if (!rights.records || record.status !== 'Published')
      throw new Error('Declaration unavailable / إعلان السجل غير متاح')
    changed.declared = true
    label = 'Declared as a record'
    arabic = 'تم إعلان الوثيقة كسجل'
  } else if (action === 'delegate') {
    if (!rights.approve || record.status !== 'PendingApproval' || !reason.trim())
      throw new Error('Enter a delegate / أدخل المفوض')
    changed.delegated = true
    changed.note = reason
    label = 'Approval delegated'
    arabic = 'تم تفويض الاعتماد'
  } else if (action === 'version') {
    if (!rights.register || record.status !== 'Published' || record.declared || record.hold)
      throw new Error('Record is frozen / السجل مجمد')
    const draft = {
      ...record,
      id: nextNumber(state),
      status: 'Draft' as Status,
      version: record.version,
      approvals: [],
      indexed: false,
      staging: 'Active',
      location: 'Staging',
      reference: record.id,
      revises: record.id,
      signature: '',
      captured: new Date().toISOString(),
    }
    return addEvent(
      { ...state, documents: [draft, ...state.documents] },
      role,
      'Revision draft created',
      'تم إنشاء مسودة إصدار',
      draft.id,
      record.id,
    )
  } else throw new Error('Unknown action')
  const previous = state.documents.find((r) => r.id === changed.revises)
  if (changed.status === 'Published' && previous && (previous.hold || previous.declared))
    throw new Error('The previous record is frozen / الإصدار السابق مجمد')
  const next = {
    ...state,
    documents: state.documents.map(
      (r): RecordItem =>
        r.id === id
          ? changed
          : changed.status === 'Published' && r.id === changed.revises
            ? { ...r, status: 'Superseded', indexed: false }
            : r,
    ),
    notices: [
      {
        id: crypto.randomUUID(),
        title: `${label} · ${record.title}`,
        titleAr: `${arabic} · ${record.titleAr || record.title}`,
        recordId: id,
        category: ['hold', 'release'].includes(action) ? 'Security' : 'Workflow',
        read: false,
        time: new Date().toISOString(),
      },
      ...state.notices,
    ],
  }
  return addEvent(next, role, label, arabic, id, reason || changed.note)
}
function publish(record: RecordItem, settings: Settings): RecordItem {
  const targetRoot = settings.sharePoint?.enabled
    ? `SharePoint Online / ${settings.sharePoint.libraryName || settings.sharePoint.publishPath || 'Official Records'}`
    : settings.library
  const targetFolder = record.sharePointFolderPath || `${record.department}/${record.kind}`
  return {
    ...record,
    status: 'Published',
    version: record.version + 1,
    indexed: false,
    note: '',
    location: `${targetRoot}/${targetFolder}/${record.fileName}`,
    staging: settings.cleanup === 'Purge' ? 'Purged' : 'Published / Approved',
  }
}
export function nextNumber(state: Workspace) {
  return `DOC-${new Date().getFullYear()}-${String(Math.max(123, ...state.documents.map((r) => Number(r.id.split('-').at(-1)) || 0)) + 1).padStart(6, '0')}`
}
// Fictional records for a UAE public-service workshop, not ministry policies or operational guidance.
function ministrySamples(): RecordItem[] {
  type Sample = Pick<
    RecordItem,
    'title' | 'titleAr' | 'kind' | 'department' | 'sensitivity' | 'status' | 'summary' | 'summaryAr' | 'flow'
  >
  const samples: Sample[] = [
    {
      title: 'Civil Defence readiness review',
      titleAr: 'مراجعة جاهزية الدفاع المدني',
      kind: 'Report',
      department: 'Civil Defence',
      sensitivity: 'Internal',
      status: 'PendingApproval',
      flow: 'Parallel',
      summary:
        'Synthetic UAE MOI workshop report for the Civil Defence service quality team. Summarises a fictional public-facility readiness exercise, accessibility checks, staff briefing attendance and follow-up ownership. The department head reviews completeness while the compliance officer checks document controls in parallel. The demonstration asks reviewers to confirm action owners and the next review date; no real facilities, vulnerabilities or emergency procedures are included.',
      summaryAr:
        'تقرير اصطناعي لورشة عرض دوكايا في سياق وزارة الداخلية الإماراتية وفريق جودة خدمات الدفاع المدني. يلخص تمرين جاهزية افتراضياً لمرفق عام وفحوص سهولة الوصول وحضور إحاطات الموظفين ومسؤوليات المتابعة. يراجع رئيس الإدارة اكتمال التقرير ومسؤول الامتثال ضوابط الوثيقة بالتوازي. يؤكد المراجعون أصحاب الإجراءات وموعد المراجعة القادمة. لا يتضمن مرافق حقيقية أو نقاط ضعف أو إجراءات طوارئ فعلية.',
    },
    {
      title: 'Road safety awareness campaign circular',
      titleAr: 'تعميم حملة التوعية بالسلامة المرورية',
      kind: 'Circular',
      department: 'Traffic & Patrols',
      sensitivity: 'Public',
      status: 'Published',
      flow: 'Single',
      summary:
        'Synthetic public-awareness circular for a fictional UAE road safety campaign. Coordinates bilingual school engagement materials, seat-belt awareness posters, accessible digital content and community feedback collection. The approved communication pack is available to all demonstration departments. Campaign dates, attendance and outreach targets are illustrative workshop values, with no enforcement instructions or real incident information.',
      summaryAr:
        'تعميم توعوي اصطناعي لحملة إماراتية افتراضية للسلامة المرورية. ينسق مواد التواصل مع المدارس باللغتين وملصقات التوعية بحزام الأمان والمحتوى الرقمي الميسر وجمع ملاحظات المجتمع. حزمة التواصل المعتمدة متاحة لجميع إدارات العرض. تواريخ الحملة والحضور والمستهدفات أمثلة توضيحية للورشة، دون تعليمات إنفاذ أو معلومات حوادث حقيقية.',
    },
    {
      title: 'Evidence custody documentation procedure',
      titleAr: 'إجراء توثيق عهدة الأدلة',
      kind: 'Procedure',
      department: 'Forensic Sciences',
      sensitivity: 'Confidential',
      status: 'PendingApproval',
      flow: 'Sequential',
      summary:
        'Synthetic document-control procedure for the Forensic Sciences demonstration library. Uses the fictional training reference TRAINING-CASE-042 to illustrate required metadata, version history, acknowledgement fields and supervisor sign-off on an evidence custody form. Department review must precede compliance review. This is a UI workflow example only: it contains no actual evidence, personal information, forensic techniques or authoritative chain-of-custody requirements.',
      summaryAr:
        'إجراء اصطناعي لضبط الوثائق في مكتبة عرض العلوم الجنائية. يستخدم المرجع التدريبي الافتراضي TRAINING-CASE-042 لعرض البيانات الوصفية المطلوبة وسجل الإصدارات وحقول الإقرار واعتماد المشرف على نموذج عهدة الأدلة. تسبق مراجعة الإدارة مراجعة الامتثال. هذا مثال لعرض سير عمل الواجهة فقط، ولا يتضمن أدلة فعلية أو بيانات شخصية أو تقنيات جنائية أو متطلبات رسمية لسلسلة العهدة.',
    },
    {
      title: 'Residency customer service quality standard',
      titleAr: 'معيار جودة خدمة متعاملي الإقامة',
      kind: 'Standard',
      department: 'Residency & Identity',
      sensitivity: 'Internal',
      status: 'Published',
      flow: 'Sequential',
      summary:
        'Synthetic service-quality standard for a fictional residency customer happiness centre. Covers bilingual guidance, accessible appointment information, document completeness checks, feedback handling and escalation ownership. Illustrative measures track first-contact resolution and satisfaction without storing applicant data. This workshop standard is not immigration policy and does not describe actual eligibility, fees or processing commitments.',
      summaryAr:
        'معيار جودة اصطناعي لمركز افتراضي لإسعاد متعاملي خدمات الإقامة. يغطي الإرشاد باللغتين وإتاحة معلومات المواعيد والتحقق من اكتمال الوثائق ومعالجة الملاحظات ومسؤولية التصعيد. تقيس المؤشرات التوضيحية الحل من أول تواصل والرضا دون تخزين بيانات المتقدمين. هذا معيار للورشة وليس سياسة هجرة ولا يصف أهلية أو رسوماً أو التزامات معالجة فعلية.',
    },
    {
      title: 'Community safety outreach manual',
      titleAr: 'دليل التواصل المجتمعي للسلامة',
      kind: 'Manual',
      department: 'Strategy & Governance',
      sensitivity: 'Public',
      status: 'Published',
      flow: 'Single',
      summary:
        'Synthetic bilingual outreach manual for a UAE community safety demonstration. Organises awareness topics, inclusive event invitations, approved presentation templates and post-event feedback forms. Sample audiences include families, schools and community partners. All event details are fictional. The manual demonstrates a searchable, reusable public communication library with clear ownership and periodic review.',
      summaryAr:
        'دليل تواصل اصطناعي باللغتين لعرض توضيحي عن السلامة المجتمعية في الإمارات. ينظم موضوعات التوعية ودعوات الفعاليات الشاملة وقوالب العروض المعتمدة ونماذج الملاحظات بعد الفعاليات. تشمل الفئات التوضيحية الأسر والمدارس والشركاء المجتمعيين. جميع تفاصيل الفعاليات افتراضية. يوضح الدليل مكتبة اتصال عامة قابلة للبحث وإعادة الاستخدام مع مسؤوليات واضحة ومراجعة دورية.',
    },
    {
      title: 'Joint emergency exercise coordination memo',
      titleAr: 'مذكرة تنسيق تمرين الطوارئ المشترك',
      kind: 'Memo',
      department: 'Operations',
      sensitivity: 'Internal',
      status: 'PendingApproval',
      flow: 'Single',
      summary:
        'Synthetic coordination memo for the fictional Al Waha tabletop workshop. Requests approval of the meeting agenda, attendee roles, observer feedback template and consolidated lessons-learned report. Civil Defence and service-centre coordinators contribute administrative observations. Exercise locations and dates are fictional and the memo contains no response tactics, deployment details or real emergency information.',
      summaryAr:
        'مذكرة تنسيق اصطناعية لورشة التمرين المكتبي الافتراضية «الواحة». تطلب اعتماد جدول الاجتماع وأدوار المشاركين ونموذج ملاحظات المراقبين وتقرير الدروس المستفادة الموحد. يساهم منسقو الدفاع المدني ومراكز الخدمة بملاحظات إدارية. مواقع التمرين وتواريخه افتراضية ولا تحتوي المذكرة على تكتيكات استجابة أو تفاصيل انتشار أو معلومات طوارئ حقيقية.',
    },
    {
      title: 'Traffic service request work instruction',
      titleAr: 'تعليمات عمل طلبات الخدمات المرورية',
      kind: 'Work Instruction',
      department: 'Traffic & Patrols',
      sensitivity: 'Internal',
      status: 'PendingApproval',
      flow: 'Single',
      summary:
        'Synthetic work instruction for routing general traffic-service enquiries in the demonstration workspace. Shows how an intake coordinator captures a fictional request reference, verifies required attachments, assigns a service owner and records a customer update. Reviewers check plain-language guidance and bilingual consistency. All turnaround targets are workshop examples, with no official service promises or real driver data.',
      summaryAr:
        'تعليمات عمل اصطناعية لتوجيه استفسارات الخدمات المرورية العامة في مساحة العرض. توضح كيف يسجل منسق الاستقبال مرجع طلب افتراضياً ويتحقق من المرفقات المطلوبة ويعين مسؤول الخدمة ويسجل تحديث المتعامل. يتحقق المراجعون من وضوح الإرشادات واتساق اللغتين. جميع مدد الإنجاز أمثلة للورشة دون وعود خدمة رسمية أو بيانات سائقين حقيقية.',
    },
    {
      title: 'Ministry document control procedure',
      titleAr: 'إجراء ضبط وثائق الوزارة',
      kind: 'Procedure',
      department: 'Strategy & Governance',
      sensitivity: 'Internal',
      status: 'Published',
      flow: 'Sequential',
      summary:
        'Synthetic procedure for the Docaya UAE MOI client workshop. AI-assisted intake suggests the document class, department and confidentiality; the document controller reviews or overrides each suggestion before submission. Approved versions enter the controlled library and superseded versions remain traceable. The sample describes the prototype workflow and provides no official records-management or retention mandate.',
      summaryAr:
        'إجراء اصطناعي لورشة عملاء دوكايا في سياق وزارة الداخلية الإماراتية. يقترح الإدخال بمساعدة الذكاء الاصطناعي فئة الوثيقة والإدارة والسرية، ويراجع مراقب الوثائق كل اقتراح أو يعدله قبل الإرسال. تدخل الإصدارات المعتمدة المكتبة المضبوطة مع بقاء الإصدارات المستبدلة قابلة للتتبع. تصف العينة سير عمل النموذج الأولي ولا تقدم تفويضاً رسمياً لإدارة السجلات أو مدد حفظها.',
    },
    {
      title: 'Civil Defence facility review checklist',
      titleAr: 'قائمة مراجعة مرافق الدفاع المدني',
      kind: 'Form',
      department: 'Civil Defence',
      sensitivity: 'Internal',
      status: 'Classified',
      flow: 'Single',
      summary:
        'Synthetic form template for recording an administrative facility review in the client demo. Contains fictional site identification, review date, document availability, staff training record checks and follow-up ownership fields. Classification is complete and the contributor can confirm the metadata before routing for approval. No actual facility assessments, inspection findings or safety certifications are represented.',
      summaryAr:
        'قالب نموذج اصطناعي لتوثيق مراجعة إدارية لمرفق في عرض العميل. يتضمن تعريف موقع افتراضي وتاريخ المراجعة وتوافر الوثائق والتحقق من سجلات تدريب الموظفين وحقول مسؤولية المتابعة. اكتمل التصنيف ويمكن للمساهم تأكيد البيانات الوصفية قبل الإرسال للاعتماد. لا يمثل تقييمات مرافق فعلية أو نتائج تفتيش أو شهادات سلامة.',
    },
    {
      title: 'Customer data protection review report',
      titleAr: 'تقرير مراجعة حماية بيانات المتعاملين',
      kind: 'Report',
      department: 'Legal',
      sensitivity: 'Confidential',
      status: 'Returned',
      flow: 'Sequential',
      summary:
        'Synthetic governance review for the fictional customer services programme. Describes document access responsibilities, purpose descriptions and a sample retention decision log. Returned to the author to add the accountable owner and clarify the review scope. All examples use invented references without personal data. The report is a demonstration of review feedback and is not a legal assessment or compliance determination.',
      summaryAr:
        'مراجعة حوكمة اصطناعية لبرنامج افتراضي لخدمات المتعاملين. تصف مسؤوليات الوصول للوثائق وأوصاف الأغراض وسجل قرارات حفظ توضيحي. أعيدت للمؤلف لإضافة المسؤول وتوضيح نطاق المراجعة. تستخدم جميع الأمثلة مراجع مختلقة دون بيانات شخصية. التقرير عرض توضيحي لملاحظات المراجعة وليس تقييماً قانونياً أو قرار امتثال.',
    },
    {
      title: 'Inclusive public service charter',
      titleAr: 'ميثاق الخدمات العامة الشاملة',
      kind: 'Policy',
      department: 'Residency & Identity',
      sensitivity: 'Public',
      status: 'PendingApproval',
      flow: 'Parallel',
      summary:
        'Synthetic customer happiness charter for a fictional UAE service centre. Sets out respectful communication, Arabic and English guidance, accessible channels and a clear feedback route. Department and compliance reviewers approve the draft in parallel. Workshop participants can compare the document preview with its metadata and request changes to the proposed service commitments before publication.',
      summaryAr:
        'ميثاق اصطناعي لإسعاد المتعاملين في مركز خدمة إماراتي افتراضي. يوضح التواصل باحترام والإرشاد بالعربية والإنجليزية والقنوات الميسرة ومسار الملاحظات الواضح. يعتمد مراجعو الإدارة والامتثال المسودة بالتوازي. يمكن للمشاركين في الورشة مقارنة معاينة الوثيقة ببياناتها الوصفية وطلب تعديل التزامات الخدمة المقترحة قبل النشر.',
    },
    {
      title: 'Service excellence training register',
      titleAr: 'سجل التدريب على التميز في الخدمة',
      kind: 'HR Record',
      department: 'Human Resources',
      sensitivity: 'Internal',
      status: 'Archived',
      flow: 'Single',
      summary:
        'Synthetic archived register for a completed service-excellence training cycle. Uses anonymous workshop groups rather than employee identities to demonstrate course completion, document ownership and archive review. The archive date and retention review are illustrative configuration examples. A records officer can inspect the history and record a disposition reason without processing any actual personnel records.',
      summaryAr:
        'سجل مؤرشف اصطناعي لدورة تدريبية مكتملة عن التميز في الخدمة. يستخدم مجموعات ورشة مجهولة بدلاً من هويات الموظفين لعرض إكمال الدورات وملكية الوثائق ومراجعة الأرشيف. تاريخ الأرشفة ومراجعة الحفظ أمثلة توضيحية للإعدادات. يمكن لمسؤول السجلات فحص السجل وتوثيق سبب التصرف دون معالجة ملفات موظفين فعلية.',
    },
  ]
  const now = Date.now()
  return samples.map((sample, index): RecordItem => {
    const released = ['Published', 'Archived'].includes(sample.status)
    const record: RecordItem = {
      ...sample,
      id: `DOC-2026-${String(201 + index).padStart(6, '0')}`,
      sampleKey: `uae-moi-workshop-${index + 1}`,
      owner: index % 2 ? 'Noor Al Shamsi' : 'Mariam Al Mansoori',
      reference: `MOI-DEMO-${String(index + 1).padStart(3, '0')}`,
      fileName: `${sample.title.toLowerCase().replaceAll(' ', '-')}-sample.pdf`,
      mime: 'application/pdf',
      size: 146432 + index * 16384,
      hash: 'Synthetic workshop fixture — upload an original to calculate SHA-256',
      captured: new Date(now - (index + 3) * 86400000).toISOString(),
      expiry: new Date(now + (sample.status === 'Archived' ? -1 : 120 + index * 15) * 86400000)
        .toISOString()
        .slice(0, 10),
      priority: index === 0 || index === 5 ? 'Priority' : 'Routine',
      scope: 'Clearance level',
      approvals: [],
      version: released ? 1 : 0,
      indexed: sample.status === 'Published',
      hold: '',
      declared: sample.status === 'Archived',
      location: released
        ? `${sample.status === 'Archived' ? 'Archive' : 'SharePoint / Official Records'}/${sample.department}/${sample.kind}`
        : 'Staging',
      staging: released ? 'Published / Approved' : 'Active',
      note:
        sample.status === 'Returned'
          ? 'Please identify the accountable owner and clarify the review scope.'
          : '',
      due: new Date(now + (index === 0 ? 4 : 18 + index * 3) * 3600000).toISOString(),
      delegated: false,
      signature: released ? 'Synthetic workshop approval · DEMO-SIGNATURE' : '',
    }
    if (released) record.approvals = approvalSteps(record)
    return record
  })
}
export function seedWorkspace(): Workspace {
  const seeds = [
    [
      'Document governance policy',
      'سياسة حوكمة الوثائق',
      'Policy',
      'Operations',
      'Internal',
      'Published',
      'All official records require manual classification and approval before publication. Retain policy records for seven years. Legal holds override disposition.',
      'تتطلب السجلات الرسمية التصنيف اليدوي والاعتماد قبل النشر. تحفظ السياسات لمدة سبع سنوات ويتجاوز الحجز القانوني قرارات الإتلاف.',
    ],
    [
      'Supplier framework agreement',
      'اتفاقية إطار الموردين',
      'Contract',
      'Finance',
      'Confidential',
      'PendingApproval',
      'Annual supplier framework for facilities services. Finance and compliance review are required before execution.',
      'اتفاقية سنوية لخدمات المرافق تتطلب مراجعة المالية والامتثال قبل التنفيذ.',
    ],
    [
      'Business continuity playbook',
      'دليل استمرارية الأعمال',
      'Report',
      'Operations',
      'Internal',
      'Published',
      'Department coordinators maintain continuity plans, review them quarterly and record annual exercise outcomes.',
      'يحافظ منسقو الإدارات على خطط الاستمرارية مع مراجعة ربع سنوية وتوثيق نتائج التمارين السنوية.',
    ],
    [
      'Annual leave circular',
      'تعميم الإجازات السنوية',
      'Circular',
      'Human Resources',
      'Public',
      'Published',
      'Annual leave requests must be submitted through the employee portal and approved by the line manager before travel.',
      'ترسل طلبات الإجازة عبر بوابة الموظفين وتعتمد من المدير المباشر قبل السفر.',
    ],
    [
      'Service improvement memorandum',
      'مذكرة تحسين الخدمات',
      'Memo',
      'Operations',
      'Internal',
      'Returned',
      'Proposed service improvements for the customer experience programme.',
      'تحسينات مقترحة للخدمات ضمن برنامج تجربة المتعاملين.',
    ],
    [
      'Quarterly budget review',
      'مراجعة الميزانية الفصلية',
      'Report',
      'Finance',
      'Confidential',
      'Draft',
      'Quarterly review of operating expenditure and procurement commitments.',
      'مراجعة فصلية للمصروفات التشغيلية والتزامات المشتريات.',
    ],
    [
      'Restricted case assessment',
      'تقييم قضية سرية',
      'Report',
      'Legal',
      'Secret',
      'PendingApproval',
      'Restricted legal case assessment for senior director and compliance approval.',
      'تقييم قضية سرية لاعتماد المدير والامتثال.',
    ],
    [
      'Records retention schedule',
      'جدول مدد حفظ السجلات',
      'Policy',
      'Legal',
      'Internal',
      'Published',
      'Contracts are retained for seven years after closure. Records under legal hold cannot be deleted. A records officer approves disposition.',
      'تحفظ العقود لسبع سنوات بعد الإغلاق. يمنع إتلاف السجلات المحجوزة قانونياً ويعتمد مسؤول السجلات التصرف النهائي.',
    ],
    [
      'Facilities inspection report',
      'تقرير تفتيش المرافق',
      'Report',
      'Operations',
      'Internal',
      'Approved',
      'Inspection findings and maintenance follow-up actions for operational facilities.',
      'نتائج التفتيش وإجراءات متابعة الصيانة للمرافق التشغيلية.',
    ],
    [
      'Employee onboarding checklist',
      'قائمة إجراءات انضمام الموظف',
      'HR Record',
      'Human Resources',
      'Confidential',
      'Published',
      'Onboarding checklist covering induction, access provisioning and mandatory training.',
      'قائمة تشمل التعريف بالمؤسسة ومنح الصلاحيات والتدريب الإلزامي.',
    ],
    [
      'Archived procurement register',
      'سجل المشتريات المؤرشف',
      'Report',
      'Finance',
      'Internal',
      'Archived',
      'Historical procurement register awaiting disposition review.',
      'سجل مشتريات تاريخي بانتظار مراجعة التصرف النهائي.',
    ],
    [
      'Interdepartmental service charter',
      'ميثاق الخدمات بين الإدارات',
      'Policy',
      'Operations',
      'Internal',
      'PendingApproval',
      'Shared service commitments with parallel department and compliance sign-off.',
      'التزامات الخدمات المشتركة مع اعتماد متواز من الإدارة والامتثال.',
    ],
  ]
  const documents = seeds.map(
    (s, i): RecordItem => ({
      id: `DOC-2026-${String(112 + i).padStart(6, '0')}`,
      title: s[0],
      titleAr: s[1],
      kind: s[2],
      department: s[3],
      sensitivity: s[4] as Sensitivity,
      status: s[5] as Status,
      summary: s[6],
      summaryAr: s[7],
      owner: i % 2 ? 'Noor Al Shamsi' : 'Mariam Al Mansoori',
      reference: '',
      fileName: `${s[0].toLowerCase().replaceAll(' ', '-')}.${i === 1 ? 'docx' : i === 5 ? 'xlsx' : 'pdf'}`,
      mime:
        i === 1
          ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
          : 'application/pdf',
      size: 184320 + i * 25812,
      hash: 'Sample fixture — upload a file to calculate SHA-256',
      captured: new Date(Date.now() - (i + 1) * 86400000).toISOString(),
      expiry: new Date(Date.now() + (i === 7 ? 7 : 90 + i * 25) * 86400000).toISOString().slice(0, 10),
      flow: i === 6 ? 'Conditional' : i === 11 ? 'Parallel' : i === 1 ? 'Sequential' : 'Single',
      priority: i === 6 ? 'Urgent' : 'Routine',
      scope: 'Clearance level',
      approvals: [],
      version: s[5] === 'Published' || s[5] === 'Archived' ? 1 : 0,
      indexed: s[5] === 'Published',
      hold: i === 7 ? 'LH-2026-014 · Records review investigation' : '',
      declared: i === 7,
      location: ['Published', 'Archived'].includes(s[5]) ? `SharePoint / ${s[3]} / ${s[2]}` : 'Staging',
      staging: s[5] === 'Published' ? 'Published / Approved' : 'Active',
      note:
        s[5] === 'Returned'
          ? 'Please add the implementation timeline and accountable owner.'
          : s[5] === 'Approved'
            ? 'Simulated connector timeout; original is retained for replay.'
            : '',
      due: new Date(Date.now() + (i === 1 ? -2 : 24) * 3600000).toISOString(),
      delegated: false,
      signature: '',
    }),
  )
  return {
    schema: 1,
    seedVersion: SEED_VERSION,
    documents: [...documents, ...ministrySamples()],
    notices: [
      {
        id: 'n1',
        title: 'Supplier agreement approval is overdue',
        titleAr: 'تأخر اعتماد اتفاقية الموردين',
        recordId: documents[1].id,
        category: 'Workflow',
        read: false,
        time: new Date().toISOString(),
      },
      {
        id: 'n2',
        title: 'Retention schedule expires in 7 days',
        titleAr: 'ينتهي جدول الحفظ خلال ٧ أيام',
        recordId: documents[7].id,
        category: 'Lifecycle',
        read: false,
        time: new Date().toISOString(),
      },
      {
        id: 'n3',
        title: 'Facilities report needs promotion replay',
        titleAr: 'تقرير المرافق يحتاج إعادة النشر',
        recordId: documents[8].id,
        category: 'System',
        read: false,
        time: new Date().toISOString(),
      },
    ],
    audit: documents.slice(0, 5).map((r) => ({
      id: crypto.randomUUID(),
      time: r.captured,
      actor: r.owner,
      role: 'System Administrator',
      action: `Sample record: ${r.status}`,
      actionAr: 'سجل تجريبي',
      object: r.id,
      detail: 'Seeded demonstration event',
    })),
    correspondence: [
      {
        id: 'COR-2026-0042',
        subject: 'Request for annual performance report',
        subjectAr: 'طلب تقرير الأداء السنوي',
        direction: 'Incoming',
        party: 'Executive Office',
        assignee: 'Operations',
        due: new Date(Date.now() + 3 * 86400000).toISOString().slice(0, 10),
        status: 'Open',
        thread: ['Executive Office → Operations: Please provide the approved annual report.'],
        acknowledged: false,
      },
      {
        id: 'COR-2026-0041',
        subject: 'Updated service hours circular',
        subjectAr: 'تعميم تحديث ساعات الخدمة',
        direction: 'Circular',
        party: 'All departments',
        assignee: 'Operations',
        due: new Date().toISOString().slice(0, 10),
        status: 'Open',
        thread: ['Distribution: All departments'],
        acknowledged: false,
      },
    ],
    settings: {
      classes: [...documentClasses],
      department: 'Operations',
      library: 'SharePoint / Official Records',
      sharePoint: defaultSharePointSettings(),
      sla: 48,
      cleanup: 'Retain',
      retentionYears: 7,
      cadence: 'Every 15 minutes',
      simulateFailure: false,
      email: true,
      teams: true,
      quietHours: '20:00–07:00',
      defaultArabic: false,
    },
    savedSearches: [],
    feedback: [],
  }
}
export const STORAGE_KEY = 'docaya-prototype-v1'
export function migrateWorkspace(value: Workspace): Workspace {
  if ((value.seedVersion ?? 1) >= SEED_VERSION && value.settings.sharePoint) return value
  const ids = new Set(value.documents.map((record) => record.id))
  const sampleKeys = new Set(value.documents.map((record) => record.sampleKey).filter(Boolean))
  return {
    ...value,
    seedVersion: SEED_VERSION,
    documents: [
      ...value.documents,
      ...ministrySamples().filter((record) => !ids.has(record.id) && !sampleKeys.has(record.sampleKey)),
    ],
    settings: {
      ...value.settings,
      classes: [...new Set([...(value.settings.classes ?? []), ...documentClasses])],
      sharePoint: { ...defaultSharePointSettings(), ...(value.settings.sharePoint ?? {}) },
    },
  }
}
export function loadWorkspace(): Workspace {
  try {
    const value = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null')
    if (value?.schema === 1 && Array.isArray(value.documents) && value.settings && Array.isArray(value.audit))
      return migrateWorkspace(value)
  } catch {
    /* Start a clean demo if storage is invalid. */
  }
  return seedWorkspace()
}
export function downloadFile(name: string, content: string, type = 'text/csv;charset=utf-8') {
  const url = URL.createObjectURL(new Blob(['\uFEFF', content], { type }))
  const a = document.createElement('a')
  a.href = url
  a.download = name
  a.click()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}
export function csv(rows: string[][]) {
  return rows
    .map((row) =>
      row
        .map((cell) => `"${(/^[=+@\-\t\r]/.test(cell) ? "'" + cell : cell).replaceAll('"', '""')}"`)
        .join(','),
    )
    .join('\r\n')
}
