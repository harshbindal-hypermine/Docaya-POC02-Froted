import { useEffect, useRef, useState, type FormEvent } from 'react'
import { Check, CloudUpload, FileText, ShieldCheck, Sparkles, ScanLine, ChevronRight } from 'lucide-react'
import { local } from './translations'
import {
  departments,
  flows,
  nextNumber,
  people,
  sensitivities,
  type Language,
  type RecordItem,
  type Role,
  type Workspace,
} from './model'
import { Field, Modal, type Translate } from './ui'
import { fingerprint, saveOriginal } from './files'
import { DocumentPreview } from './DocumentPreview'
import { ApiError, getSharePointFoldersByDepartment } from '../services/api'
import { SharePointFolderTree } from '../components/SharePointFolderTree'
import type { SharePointFolder } from '../types'
import {
  applyClassification,
  suggestClassification,
  type ClassificationSuggestion,
  type SuggestedFields,
} from './classification'
import './registration-ai.css'

const suggestedKeys: (keyof SuggestedFields)[] = [
  'title',
  'kind',
  'department',
  'sensitivity',
  'flow',
  'summary',
]
const samples = [
  {
    title: 'Traffic safety circular',
    titleAr: 'تعميم السلامة المرورية',
    filename: 'moi-traffic-safety-circular.txt',
    text: 'SYNTHETIC MOI DEMO DOCUMENT\nTraffic and Patrols — public road safety circular\nCoordinate an awareness campaign on school-zone speed limits, pedestrian crossings and safe arrival. Department coordinators will review the proposed campaign before publication. All locations and activities in this example are fictional.',
  },
  {
    title: 'Civil Defence procedure',
    titleAr: 'إجراءات الدفاع المدني',
    filename: 'civil-defence-evacuation-procedure.txt',
    text: 'SYNTHETIC MOI DEMO DOCUMENT\nCivil Defence — evacuation procedure\nInternal workshop exercise for a fictional ministry service centre. Verify exit routes, assign floor marshals, record assembly-point attendance and document the post-exercise review.',
  },
  {
    title: 'Evidence handling standard',
    titleAr: 'معيار التعامل مع الأدلة',
    filename: 'forensic-evidence-handling-standard.txt',
    text: 'SYNTHETIC MOI DEMO DOCUMENT\nForensic Sciences — evidence handling standard\nConfidential. Use fictional evidence references only. Record each chain of custody transfer, verify the package seal and obtain sequential supervisory and compliance approval. This is sample content, not an official operational procedure.',
  },
]

export function Registration({
  state,
  role,
  lang,
  t,
  existing,
  onClose,
  onSave,
}: {
  state: Workspace
  role: Role
  lang: Language
  t: Translate
  existing?: RecordItem
  onClose: () => void
  onSave: (record: RecordItem) => void
}) {
  const [step, setStep] = useState(existing ? 1 : 0)
  const [file, setFile] = useState<File>()
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [suggestion, setSuggestion] = useState<ClassificationSuggestion>()
  const [confirmed, setConfirmed] = useState(false)
  const [sampleIndex, setSampleIndex] = useState(0)
  const [sharePointFolders, setSharePointFolders] = useState<SharePointFolder[]>([])
  const [sharePointLoading, setSharePointLoading] = useState(false)
  const [sharePointError, setSharePointError] = useState('')
  const generation = useRef(0)
  const dirty = useRef(new Set<keyof RecordItem>(existing ? suggestedKeys : []))
  const [record, setRecord] = useState<RecordItem>(
    () =>
      existing || {
        id: nextNumber(state),
        title: '',
        titleAr: '',
        kind: '',
        department: state.settings.department,
        sensitivity: 'Internal',
        status: 'Draft',
        owner: people[role],
        summary: '',
        summaryAr: '',
        reference: '',
        fileName: '',
        mime: '',
        size: 0,
        hash: '',
        captured: new Date().toISOString(),
        expiry: new Date(new Date().setFullYear(new Date().getFullYear() + state.settings.retentionYears))
          .toISOString()
          .slice(0, 10),
        flow: 'Single',
        priority: 'Routine',
        scope: 'Department-only',
        approvals: [],
        version: 0,
        indexed: false,
        hold: '',
        declared: false,
        location: 'Staging',
        staging: 'Active',
        note: '',
        due: '',
        delegated: false,
        signature: '',
      },
  )
  useEffect(
    () => () => {
      generation.current += 1
    },
    [],
  )
  useEffect(() => {
    if (!record.department) {
      setSharePointFolders([])
      setSharePointError('')
      setSharePointLoading(false)
      return
    }
    const controller = new AbortController()
    setSharePointLoading(true)
    setSharePointError('')
    getSharePointFoldersByDepartment(record.department, controller.signal)
      .then((response) => {
        const folders = response.items || response.folders || []
        setSharePointFolders(folders)
        if (
          record.sharePointFolderId &&
          !folders.some((folder) => folder.id === record.sharePointFolderId)
        ) {
          setRecord((current) => ({
            ...current,
            sharePointFolderId: '',
            sharePointFolderName: '',
            sharePointFolderPath: '',
          }))
          setConfirmed(false)
        }
      })
      .catch((problem) => {
        if (controller.signal.aborted) return
        setSharePointFolders([])
        setSharePointError(
          problem instanceof ApiError
            ? problem.message
            : t('SharePoint folders could not be loaded. Please try again.', 'تعذر تحميل مجلدات شيربوينت. حاول مرة أخرى.'),
        )
      })
      .finally(() => {
        if (!controller.signal.aborted) setSharePointLoading(false)
      })
    return () => controller.abort()
  }, [record.department, record.sharePointFolderId, t])
  const update = <K extends keyof RecordItem>(key: K, value: RecordItem[K]) => {
    dirty.current.add(key)
    setConfirmed(false)
    setRecord((r) => ({ ...r, [key]: value }))
  }
  const updateDepartment = (department: string) => {
    dirty.current.add('department')
    setConfirmed(false)
    setRecord((r) => ({
      ...r,
      department,
      sharePointFolderId: '',
      sharePointFolderName: '',
      sharePointFolderPath: '',
    }))
  }
  const updateSharePointFolder = (folder?: SharePointFolder | null) => {
    setConfirmed(false)
    setRecord((r) => ({
      ...r,
      sharePointFolderId: folder?.id || '',
      sharePointFolderName: folder?.name || '',
      sharePointFolderPath: folder?.path || '',
    }))
  }
  const capture = async (selected?: File) => {
    if (!selected) return
    const request = ++generation.current
    setError('')
    setBusy(true)
    setConfirmed(false)
    try {
      if (selected.size > 20 * 1024 * 1024)
        throw new Error(
          t('This browser demo accepts files up to 20 MB.', 'يقبل العرض ملفات حتى ٢٠ ميغابايت.'),
        )
      const hash = await fingerprint(selected)
      if (request !== generation.current) return
      if (state.documents.some((r) => r.hash === hash && r.id !== record.id))
        throw new Error(
          t(
            'This file is already registered. Open the existing record to create a revision.',
            'هذا الملف مسجل بالفعل. افتح السجل الحالي لإنشاء إصدار.',
          ),
        )
      const content = selected.name.toLowerCase().endsWith('.txt')
        ? await selected.slice(0, 64000).text()
        : ''
      if (request !== generation.current) return
      const result = suggestClassification(selected.name, content, {
        classes: state.settings.classes,
        departments,
        defaultDepartment: state.settings.department,
      })
      const protectedFields = new Set(dirty.current)
      if (role === 'Department Administrator') protectedFields.add('department')
      setSuggestion(result)
      setFile(selected)
      setRecord((r) => ({
        ...applyClassification(r, result, protectedFields),
        fileName: selected.name,
        mime: selected.type || 'application/octet-stream',
        size: selected.size,
        hash,
      }))
    } catch (e) {
      if (request === generation.current) setError(e instanceof Error ? e.message : String(e))
    } finally {
      if (request === generation.current) setBusy(false)
    }
  }
  const save = async (status: RecordItem['status']) => {
    if (busy) return
    if (
      status === 'PendingApproval' &&
      (!confirmed || !record.title.trim() || !record.kind || !record.summary.trim() || !record.sharePointFolderId)
    ) {
      setError(
        t(
          'Review the required metadata and confirm the classification before submitting.',
          'راجع البيانات المطلوبة وأكد التصنيف قبل الإرسال.',
        ),
      )
      return
    }
    setBusy(true)
    setError('')
    try {
      if (file) await saveOriginal(record.id, file)
      onSave({
        ...record,
        status,
        approvals: [],
        note: '',
        due:
          status === 'PendingApproval'
            ? new Date(Date.now() + state.settings.sla * 3600000).toISOString()
            : '',
      })
    } catch {
      setError(
        t(
          'The browser could not store this file. Free some storage and try again.',
          'تعذر حفظ الملف في المتصفح. وفر مساحة وحاول مجدداً.',
        ),
      )
      setBusy(false)
    }
  }
  const advance = (e: FormEvent) => {
    e.preventDefault()
    if (busy || !record.fileName) return
    if (step === 1 && !record.sharePointFolderId) {
      setError(t('Select a SharePoint folder before review.', 'اختر مجلد شيربوينت قبل المراجعة.'))
      return
    }
    if (step === 2) void save('PendingApproval')
    else {
      setError('')
      setStep(step + 1)
    }
  }
  const overrides = suggestion ? suggestedKeys.filter((key) => suggestion.values[key] !== record[key]) : []
  const classificationSummary = suggestion && (
    <section
      className="p-ai-classification"
      aria-label={t('Automatic classification suggestions', 'اقتراحات التصنيف التلقائي')}
    >
      <div className="p-ai-classification-heading">
        <span className="p-ai-classification-icon">
          <Sparkles size={20} />
        </span>
        <div>
          <strong>{t('Docaya has prepared your metadata', 'أعد دوكايا البيانات الوصفية')}</strong>
          <small>
            {t('Local demo suggestions · your review is required', 'اقتراحات تجريبية محلية · تتطلب مراجعتك')}
          </small>
        </div>
        <span className="p-ai-confidence">
          {suggestion.confidence}%<small>{t('rule confidence', 'ثقة القواعد')}</small>
        </span>
      </div>
      <div className="p-ai-suggestion-grid">
        {(['kind', 'department', 'sensitivity', 'flow'] as const).map((key) => (
          <div key={key} className={overrides.includes(key) ? 'overridden' : ''}>
            <small>
              {key === 'kind'
                ? t('Class', 'الفئة')
                : key === 'department'
                  ? t('Department', 'الإدارة')
                  : key === 'sensitivity'
                    ? t('Sensitivity', 'الحساسية')
                    : t('Approval', 'الاعتماد')}
            </small>
            <strong>{record[key] ? local(record[key], lang) : t('Choose a class', 'اختر الفئة')}</strong>
            {overrides.includes(key) && <span>{t('Your value', 'اختيارك')}</span>}
          </div>
        ))}
      </div>
      <details>
        <summary>{t('Why these suggestions?', 'لماذا هذه الاقتراحات؟')}</summary>
        <p>
          {suggestion.source === 'filename-and-text'
            ? t(
                'Read from the filename and the first 12,000 characters of local TXT content.',
                'استناداً إلى اسم الملف وأول ١٢٬٠٠٠ حرف من محتوى الملف النصي المحلي.',
              )
            : t(
                'Read from the filename only. PDF, Office and image contents have not been analysed.',
                'استناداً إلى اسم الملف فقط. لم يتم تحليل محتوى ملفات PDF أو Office أو الصور.',
              )}
        </p>
        <ul>
          {suggestion.reasons.map((reason) => (
            <li key={reason.en}>{t(reason.en, reason.ar)}</li>
          ))}
        </ul>
        <p>
          {t(
            'Confidence is a rule-based demo score, not a calibrated AI probability. No OCR or external AI service is connected.',
            'الثقة نتيجة قواعد تجريبية وليست احتمالاً معايراً للذكاء الاصطناعي. لا توجد خدمة خارجية للذكاء الاصطناعي أو التعرف على النصوص.',
          )}
        </p>
      </details>
      {overrides.length > 0 && (
        <p className="p-ai-overrides">
          <Check size={14} />
          {t(
            `${overrides.length} field(s) kept as your choice.`,
            `تم الاحتفاظ باختيارك في ${overrides.length} حقول.`,
          )}
        </p>
      )}
      <p className="p-ai-edit-hint">
        {t(
          'Every suggestion can be edited in Classify. Your corrections are preserved when you replace the file.',
          'يمكنك تعديل جميع الاقتراحات في التصنيف. تُحفظ تصحيحاتك عند استبدال الملف.',
        )}
      </p>
    </section>
  )
  return (
    <Modal
      title={existing ? t('Edit registration', 'تعديل التسجيل') : t('Register a document', 'تسجيل وثيقة')}
      onClose={onClose}
      wide
      variant="drawer"
    >
      <div className="p-registration-layout">
        <aside className="p-registration-preview">
          <div className="p-registration-preview-heading">
            <ScanLine size={18} />
            <div>
              <strong>{t('Original document', 'الوثيقة الأصلية')}</strong>
              <small>{t('Preview alongside your metadata', 'معاينة بجانب البيانات الوصفية')}</small>
            </div>
            <span>{t('Preserved', 'محفوظة')}</span>
          </div>
          {record.fileName ? (
            <DocumentPreview record={record} lang={lang} file={file} />
          ) : (
            <div className="p-registration-preview-empty">
              <FileText size={48} />
              <h3>{t('Your document, in view', 'وثيقتك أمامك')}</h3>
              <p>
                {t(
                  'Upload a file to preview it here while Docaya suggests its classification.',
                  'ارفع ملفاً لمعاينته هنا أثناء اقتراح دوكايا لتصنيفه.',
                )}
              </p>
              <div className="p-preview-paper-placeholder">
                <span />
                <span />
                <span />
                <span />
                <span />
              </div>
            </div>
          )}
          <p className="p-registration-original-note">
            <ShieldCheck size={15} />
            {t(
              'Original bytes and SHA-256 retained. No conversion.',
              'حفظ البيانات الأصلية وبصمة SHA-256 دون تحويل.',
            )}
          </p>
        </aside>
        <div className="p-registration-controls">
          <div className="p-stepper">
            {[
              t('Capture', 'الالتقاط'),
              t('Classify', 'التصنيف'),
              t('Review & submit', 'المراجعة والإرسال'),
            ].map((label, i) => (
              <div className={step >= i ? 'active' : ''} key={label}>
                <span>{step > i ? <Check size={14} /> : i + 1}</span>
                {label}
              </div>
            ))}
          </div>
          <form onSubmit={advance} className="p-modal-body p-registration-form">
            <fieldset disabled={busy} className="p-registration-fields">
              {step === 0 && (
                <>
                  <div
                    className="p-drop"
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={(e) => {
                      e.preventDefault()
                      if (!busy) void capture(e.dataTransfer.files[0])
                    }}
                  >
                    <CloudUpload size={34} />
                    <h3>{t('Drop it in. Let Docaya organise it.', 'ارفع وثيقتك ودع دوكايا ينظمها.')}</h3>
                    <p>
                      {t(
                        'A preview and editable classification, together.',
                        'معاينة وتصنيف قابل للتعديل في مكان واحد.',
                      )}
                    </p>
                    <input
                      aria-label={t('Choose file', 'اختيار ملف')}
                      type="file"
                      accept=".pdf,.docx,.xlsx,.pptx,.png,.jpg,.jpeg,.tif,.tiff,.txt"
                      onChange={(e) => {
                        void capture(e.target.files?.[0])
                        e.target.value = ''
                      }}
                    />
                    <small>PDF · DOCX · XLSX · PPTX · PNG · JPEG · TIFF · TXT · 20 MB</small>
                  </div>
                  <div className="p-registration-sample">
                    <Field label={t('Explore an MOI scenario', 'استكشف سيناريو لوزارة الداخلية')}>
                      <select value={sampleIndex} onChange={(e) => setSampleIndex(Number(e.target.value))}>
                        {samples.map((sample, i) => (
                          <option value={i} key={sample.filename}>
                            {t(sample.title, sample.titleAr)}
                          </option>
                        ))}
                      </select>
                    </Field>
                    <button
                      type="button"
                      className="p-btn"
                      onClick={() => {
                        const sample = samples[sampleIndex]
                        void capture(new File([sample.text], sample.filename, { type: 'text/plain' }))
                      }}
                    >
                      {t('Use a sample document', 'استخدام وثيقة تجريبية')}
                      <ChevronRight size={15} />
                    </button>
                  </div>
                  {record.fileName && (
                    <div className="p-file-summary">
                      <FileText />
                      <div>
                        <strong>{record.fileName}</strong>
                        <small>
                          {(record.size / 1024).toFixed(1)} KB ·{' '}
                          {t(
                            'File signature checked · SHA-256 calculated',
                            'تم فحص توقيع الملف وحساب البصمة',
                          )}
                        </small>
                      </div>
                      <Check size={19} />
                    </div>
                  )}
                  {classificationSummary}
                  <p className="p-subtle">
                    {t(
                      'Synthetic MOI scenarios for client discussion. Malware scanning, DLP and AI services are simulated.',
                      'سيناريوهات افتراضية لوزارة الداخلية لمناقشتها مع العميل. فحص الملفات ومنع تسرب البيانات وخدمات الذكاء الاصطناعي محاكاة.',
                    )}
                  </p>
                </>
              )}
              {step === 1 && (
                <>
                  {classificationSummary}
                  <div className="p-form-grid">
                    <Field label={t('Business title *', 'عنوان الوثيقة *')}>
                      <input
                        required
                        value={record.title}
                        onChange={(e) => update('title', e.target.value)}
                      />
                    </Field>
                    <Field label={t('Arabic title', 'العنوان العربي')}>
                      <input
                        dir="rtl"
                        value={record.titleAr}
                        onChange={(e) => update('titleAr', e.target.value)}
                      />
                    </Field>
                    <Field label={t('Document class *', 'فئة الوثيقة *')}>
                      <select required value={record.kind} onChange={(e) => update('kind', e.target.value)}>
                        <option value="">{t('Select a class', 'اختر الفئة')}</option>
                        {state.settings.classes.map((v) => (
                          <option key={v} value={v}>
                            {local(v, lang)}
                          </option>
                        ))}
                      </select>
                    </Field>
                    <Field label={t('Owning department *', 'الإدارة المالكة *')}>
                      <select
                        value={record.department}
                        disabled={role === 'Department Administrator'}
                        onChange={(e) => updateDepartment(e.target.value)}
                      >
                        {departments.map((v) => (
                          <option key={v} value={v}>
                            {local(v, lang)}
                          </option>
                        ))}
                      </select>
                    </Field>
                    <div className="p-span">
                      <Field label={t('SharePoint Folder *', 'مجلد شيربوينت *')}>
                        <SharePointFolderTree
                          library={{ id: 'assigned-folders', name: t('Assigned folders', 'المجلدات المرتبطة') }}
                          roots={sharePointFolders}
                          selectedId={record.sharePointFolderId || ''}
                          loading={sharePointLoading}
                          error={sharePointError}
                          emptyMessage={t(
                            'No SharePoint folders are available for this department.',
                            'لا توجد مجلدات شيربوينت متاحة لهذه الإدارة.',
                          )}
                          loadingMessage={t('Loading SharePoint folders...', 'جار تحميل مجلدات شيربوينت...')}
                          errorMessage={t(
                            'SharePoint folders could not be loaded. Please try again.',
                            'تعذر تحميل مجلدات شيربوينت. حاول مرة أخرى.',
                          )}
                          selectLabel={t('Select SharePoint folder', 'اختر مجلد شيربوينت')}
                          onSelect={updateSharePointFolder}
                        />
                        {sharePointError && (
                          <small className="p-error-text">
                            {t(
                              'SharePoint folders could not be loaded. Please try again.',
                              'تعذر تحميل مجلدات شيربوينت. حاول مرة أخرى.',
                            )}
                          </small>
                        )}
                      </Field>
                    </div>
                    <Field label={t('Sensitivity *', 'درجة الحساسية *')}>
                      <select
                        value={record.sensitivity}
                        onChange={(e) => update('sensitivity', e.target.value as RecordItem['sensitivity'])}
                      >
                        {sensitivities.map((v) => (
                          <option key={v} value={v}>
                            {local(v, lang)}
                          </option>
                        ))}
                      </select>
                    </Field>
                    <Field label={t('Approval flow *', 'مسار الاعتماد *')}>
                      <select
                        value={record.flow}
                        onChange={(e) => update('flow', e.target.value as RecordItem['flow'])}
                      >
                        {flows.map((v) => (
                          <option key={v} value={v}>
                            {local(v, lang)}
                          </option>
                        ))}
                      </select>
                    </Field>
                    <Field label={t('Priority', 'الأولوية')}>
                      <select value={record.priority} onChange={(e) => update('priority', e.target.value)}>
                        {['Routine', 'Priority', 'Urgent'].map((v) => (
                          <option key={v} value={v}>
                            {local(v, lang)}
                          </option>
                        ))}
                      </select>
                    </Field>
                    <Field label={t('Access scope', 'نطاق الوصول')}>
                      <select value={record.scope} onChange={(e) => update('scope', e.target.value)}>
                        {['Department-only', 'Clearance level'].map((v) => (
                          <option key={v} value={v}>
                            {local(v, lang)}
                          </option>
                        ))}
                      </select>
                    </Field>
                    <Field label={t('Reference / keywords', 'المرجع / الكلمات المفتاحية')}>
                      <input value={record.reference} onChange={(e) => update('reference', e.target.value)} />
                    </Field>
                    <Field label={t('Expiry / review date', 'تاريخ الانتهاء / المراجعة')}>
                      <input
                        required
                        type="date"
                        value={record.expiry}
                        onChange={(e) => update('expiry', e.target.value)}
                      />
                    </Field>
                    <div className="p-span">
                      <Field label={t('Description *', 'الوصف *')}>
                        <textarea
                          required
                          rows={4}
                          value={record.summary}
                          onChange={(e) => update('summary', e.target.value)}
                        />
                      </Field>
                    </div>
                    <div className="p-span">
                      <Field label={t('Arabic description', 'الوصف العربي')}>
                        <textarea
                          dir="rtl"
                          rows={2}
                          value={record.summaryAr}
                          onChange={(e) => update('summaryAr', e.target.value)}
                        />
                      </Field>
                    </div>
                  </div>
                </>
              )}
              {step === 2 && (
                <>
                  <div className="p-review-card">
                    <p className="p-eyebrow">{record.id}</p>
                    <h2>{lang === 'ar' ? record.titleAr || record.title : record.title}</h2>
                    <p>{lang === 'ar' ? record.summaryAr || record.summary : record.summary}</p>
                    <dl className="p-facts">
                      <dt>{t('Classification', 'التصنيف')}</dt>
                      <dd>
                        {local(record.kind, lang)} · {local(record.sensitivity, lang)}
                      </dd>
                      <dt>{t('Approval', 'الاعتماد')}</dt>
                      <dd>
                        {local(record.flow, lang)} · {state.settings.sla} {t('hour SLA', 'ساعة للمعالجة')}
                      </dd>
                      <dt>{t('Owning Department', 'الإدارة المالكة')}</dt>
                      <dd>{local(record.department, lang)}</dd>
                      <dt>{t('SharePoint Folder', 'مجلد شيربوينت')}</dt>
                      <dd>{record.sharePointFolderName || t('Not selected', 'لم يتم الاختيار')}</dd>
                      <dt>{t('Destination after approval', 'الوجهة بعد الاعتماد')}</dt>
                      <dd>{record.sharePointFolderName || t('Selected folder', 'المجلد المحدد')}</dd>
                      <dt>{t('Original file', 'الملف الأصلي')}</dt>
                      <dd>
                        {record.fileName} · {(record.size / 1024).toFixed(1)} KB
                      </dd>
                      <dt>SHA-256</dt>
                      <dd className="p-hash">{record.hash}</dd>
                    </dl>
                  </div>
                  <label className="p-classification-confirm">
                    <input
                      type="checkbox"
                      checked={confirmed}
                      onChange={(e) => setConfirmed(e.target.checked)}
                    />
                    <span>
                      <strong>
                        {t('I have reviewed the document and classification', 'راجعت الوثيقة وتصنيفها')}
                      </strong>
                      <small>
                        {t(
                          'The title, department, sensitivity and approval route reflect my review. AI suggestions do not approve or publish this document.',
                          'العنوان والإدارة والحساسية ومسار الاعتماد تعكس مراجعتي. الاقتراحات لا تعتمد الوثيقة أو تنشرها.',
                        )}
                      </small>
                    </span>
                  </label>
                  <div className="p-note">
                    <ShieldCheck size={18} />
                    {t(
                      'Submission sends the original to the approval queue. Publication follows every required approval.',
                      'يرسل الملف الأصلي إلى قائمة الاعتماد. يتم النشر بعد استكمال جميع الموافقات المطلوبة.',
                    )}
                  </div>
                </>
              )}
            </fieldset>
            {busy && (
              <div className="p-ai-processing" role="status">
                <Sparkles size={17} />
                {t(
                  'Reading your document and preparing metadata…',
                  'جارٍ قراءة الوثيقة وإعداد البيانات الوصفية…',
                )}
              </div>
            )}
            {error && (
              <p className="p-error" role="alert">
                {error}
              </p>
            )}
            <footer className="p-modal-actions">
              <button
                type="button"
                className="p-btn"
                disabled={busy}
                onClick={() => (step ? setStep(step - 1) : onClose())}
              >
                {step ? t('Back', 'رجوع') : t('Cancel', 'إلغاء')}
              </button>
              <span />
              {record.fileName && (
                <button
                  type="button"
                  className="p-btn"
                  disabled={busy}
                  onClick={() => void save(step === 2 ? 'Classified' : 'Draft')}
                >
                  {t('Save draft', 'حفظ المسودة')}
                </button>
              )}
              <button
                className="p-btn primary"
                disabled={busy || !record.fileName || (step === 2 && !confirmed)}
              >
                {busy
                  ? t('Processing…', 'جارٍ المعالجة…')
                  : step === 2
                    ? t('Submit for approval', 'إرسال للاعتماد')
                    : t('Continue', 'متابعة')}
              </button>
            </footer>
          </form>
        </div>
      </div>
    </Modal>
  )
}
