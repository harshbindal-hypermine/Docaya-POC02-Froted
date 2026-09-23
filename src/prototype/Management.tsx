import { SharePointFolderTree, foldersOnly } from '../components/SharePointFolderTree'
import { local } from './translations'
import { useEffect, useState, type Dispatch, type SetStateAction } from 'react'
import {
  ArrowRight,
  Check,
  Download,
  FileText,
  FolderOpen,
  LockKeyhole,
  Plus,
  RefreshCw,
  Send,
  Settings2,
  ShieldCheck,
  Trash2,
} from 'lucide-react'
import {
  addEvent,
  capabilities,
  csv,
  departments,
  downloadFile,
  people,
  roles,
  type Correspondence,
  type Language,
  type RecordItem,
  type Role,
  type Workspace,
} from './model'
import { Badge, DocumentTable, Empty, Field, Modal, type Translate } from './ui'
import {
  addDepartmentSharePointFolderAccess,
  downloadSharePointFile,
  getSharePointPreview,
  getSharePointFoldersByDepartment,
  listSharePointItems,
  listAdminSharePointFolders,
  listDepartmentSharePointFolderAccess,
  removeDepartmentSharePointFolderAccess,
} from '../services/api'
import type { DepartmentSharePointFolder, SharePointFolder, SharePointItem } from '../types'
type PreviewState = {
  file: SharePointItem
  url: string
  mimeType: string
  text: string
}
export type PageProps = {
  state: Workspace
  setState: Dispatch<SetStateAction<Workspace>>
  role: Role
  lang: Language
  department: string
  t: Translate
  visible: RecordItem[]
  onOpen: (r: RecordItem) => void
  notify: (text: string) => void
  selectedId?: string
}
export function DocumentLibraryPage({ department, lang, t }: PageProps) {
  const [folders, setFolders] = useState<DepartmentSharePointFolder[]>([])
  const [drive, setDrive] = useState<{ name: string; webUrl?: string | null } | null>(null)
  const [site, setSite] = useState<{ hostname: string; path: string } | null>(null)
  const [selectedFolder, setSelectedFolder] = useState<SharePointFolder | null>(null)
  const [selectedFile, setSelectedFile] = useState<SharePointItem | null>(null)
  const [preview, setPreview] = useState<PreviewState | null>(null)
  const [previewBusy, setPreviewBusy] = useState(false)
  const [previewError, setPreviewError] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')

  const loadFolders = async (signal?: AbortSignal) => {
    setBusy(true)
    setError('')
    try {
      const response = await getSharePointFoldersByDepartment(department, signal)
      setFolders(response.items || response.folders || [])
      setDrive(response.drive)
      setSite(response.site)
      setSelectedFolder(null)
      setSelectedFile(null)
      setPreview(null)
      setPreviewError('')
    } catch (problem) {
      if (signal?.aborted) return
      setFolders([])
      setDrive(null)
      setSite(null)
      setError(
        problem instanceof Error
          ? problem.message
          : t('Assigned SharePoint folders could not be loaded.', 'تعذر تحميل مجلدات شيربوينت المرتبطة.'),
      )
    } finally {
      if (!signal?.aborted) setBusy(false)
    }
  }

  useEffect(() => {
    const controller = new AbortController()
    void loadFolders(controller.signal)
    return () => controller.abort()
  }, [department])

  const siteRoot = site ? `${site.hostname}${site.path}` : ''
  const libraryName = drive?.name || t('Documents', 'المستندات')
  const loadChildren = async (folderId: string) => (await listSharePointItems(folderId)).items
  const previewKind = preview?.mimeType || selectedFile?.mimeType || ''
  const previewName = preview?.file.name || selectedFile?.name || ''
  const canEmbedPreview =
    previewKind.startsWith('image/') || previewKind === 'application/pdf' || previewKind.startsWith('text/')

  const openFilePreview = async (file: SharePointItem) => {
    setSelectedFile(file)
    setSelectedFolder(null)
    setPreview(null)
    setPreviewError('')
    setPreviewBusy(true)
    try {
      await getSharePointPreview(file.id)
      const blob = await downloadSharePointFile(file.id)
      const mimeType = file.mimeType || blob.type || 'application/octet-stream'
      const url = URL.createObjectURL(blob)
      const text = mimeType.startsWith('text/') ? await blob.text() : ''
      setPreview((current) => {
        if (current?.url) URL.revokeObjectURL(current.url)
        return { file, url, mimeType, text }
      })
    } catch (problem) {
      setPreviewError(
        problem instanceof Error
          ? problem.message
          : t('Document preview could not be loaded.', 'تعذر تحميل معاينة الوثيقة.'),
      )
    } finally {
      setPreviewBusy(false)
    }
  }

  useEffect(() => {
    return () => {
      if (preview?.url) URL.revokeObjectURL(preview.url)
    }
  }, [preview?.url])

  return (
    <section className="p-card p-library-card">
      <div className="p-card-head">
        <div>
          <h3>{t('Assigned SharePoint folders', 'مجلدات شيربوينت المرتبطة')}</h3>
          <p>
            {t(
              `Folders available to ${department} from the configured SharePoint library.`,
              `المجلدات المتاحة لإدارة ${local(department, lang)} من مكتبة شيربوينت المهيأة.`,
            )}
          </p>
        </div>
        <button className="p-btn" disabled={busy} onClick={() => void loadFolders()}>
          <RefreshCw size={16} className={busy ? 'p-spin' : ''} />
          {t('Refresh', 'تحديث')}
        </button>
      </div>
      <div className="p-library-summary">
        <span>
          <strong>{local(department, lang)}</strong>
          <small>{t('Selected department', 'الإدارة المحددة')}</small>
        </span>
        <span>
          <strong>{libraryName}</strong>
          <small>{t('Assigned SharePoint folders for this department', 'مجلدات شيربوينت المرتبطة بهذه الإدارة')}</small>
        </span>
        <span>
          <strong>{folders.length}</strong>
          <small>{t('Assigned folders', 'مجلدات مرتبطة')}</small>
        </span>
      </div>
      {error && <div className="p-error">{error}</div>}
      {!error && folders.length === 0 && !busy ? (
        <div className="p-empty">
          <FolderOpen size={32} />
          <h3>{t('No folders assigned yet', 'لا توجد مجلدات مرتبطة بعد')}</h3>
          <p>
            {t(
              'Use Admin center -> Folder access to assign SharePoint folders to this department.',
              'استخدم مركز الإدارة ← صلاحية المجلدات لربط مجلدات شيربوينت بهذه الإدارة.',
            )}
          </p>
        </div>
      ) : (
        <>
          <div className="p-library-browser">
            <SharePointFolderTree
              library={{ id: 'sharepoint-drive', name: libraryName }}
              roots={folders}
              selectedId={selectedFile?.id || selectedFolder?.id}
              loading={busy}
              error={error}
              emptyMessage={t(
                `No SharePoint folders have been assigned to ${department} yet.`,
                `لم يتم ربط مجلدات شيربوينت بإدارة ${local(department, lang)} بعد.`,
              )}
              loadingMessage={t('Loading SharePoint folders...', 'جار تحميل مجلدات شيربوينت...')}
              errorMessage={t(
                'Unable to load SharePoint folders. Please check the SharePoint connection or try again.',
                'تعذر تحميل مجلدات شيربوينت. تحقق من الاتصال أو حاول مرة أخرى.',
              )}
              selectLabel={t('Open', 'فتح')}
              showFiles
              onSelect={(folder) => {
                setSelectedFolder(folder)
                setSelectedFile(null)
                setPreviewError('')
              }}
              onFileSelect={(file) => void openFilePreview(file)}
              onLoadChildren={loadChildren}
            />
            <aside className="p-library-preview" aria-live="polite">
              <div className="p-library-preview-head">
                <FileText size={18} />
                <div>
                  <strong>{previewName || t('Select a file', 'اختر ملفاً')}</strong>
                  <small>
                    {selectedFile?.path ||
                      t(
                        'Preview SharePoint file content inside Docaya.',
                        'معاينة محتوى ملف شيربوينت داخل دوكايا.',
                      )}
                  </small>
                </div>
              </div>
              {previewBusy && <div className="p-library-preview-state">{t('Loading preview...', 'جار تحميل المعاينة...')}</div>}
              {previewError && <div className="p-error">{previewError}</div>}
              {!previewBusy && !previewError && !selectedFile && (
                <div className="p-library-preview-state">
                  {t('Choose a file from the folder tree to preview it here.', 'اختر ملفاً من شجرة المجلدات لمعاينته هنا.')}
                </div>
              )}
              {!previewBusy && !previewError && preview && canEmbedPreview && previewKind.startsWith('image/') && (
                <img src={preview.url} alt={preview.file.name} />
              )}
              {!previewBusy && !previewError && preview && canEmbedPreview && previewKind === 'application/pdf' && (
                <iframe title={preview.file.name} src={preview.url} />
              )}
              {!previewBusy && !previewError && preview && canEmbedPreview && previewKind.startsWith('text/') && (
                <pre>{preview.text}</pre>
              )}
              {!previewBusy && !previewError && preview && !canEmbedPreview && (
                <div className="p-library-preview-state">
                  <strong>{t('Preview is not available for this file type.', 'المعاينة غير متاحة لهذا النوع من الملفات.')}</strong>
                  {preview.file.webUrl && (
                    <a href={preview.file.webUrl} target="_blank" rel="noreferrer">
                      {t('Open in SharePoint', 'فتح في شيربوينت')}
                    </a>
                  )}
                  <a href={preview.url} download={preview.file.name}>
                    {t('Download file', 'تنزيل الملف')}
                  </a>
                </div>
              )}
            </aside>
          </div>
          {selectedFolder && (
            <div className="p-folder-selection">
              <FolderOpen size={17} />
              <span>{selectedFolder.name}</span>
            </div>
          )}
        </>
      )}
    </section>
  )
}
export function Operations({
  state,
  setState,
  role,
  lang,
  t,
  visible,
  onOpen,
  notify,
  selectedId,
}: PageProps) {
  const [tab, setTab] = useState('staging')
  const [running, setRunning] = useState(false)
  const eligible = visible.filter((r) => r.status === 'Published' && !r.indexed)
  const runIndex = () => {
    setRunning(true)
    window.setTimeout(() => {
      setState((s) =>
        addEvent(
          {
            ...s,
            documents: s.documents.map((r) => (r.status === 'Published' ? { ...r, indexed: true } : r)),
            notices: [
              {
                id: crypto.randomUUID(),
                title: 'Published-document indexing batch completed',
                titleAr: 'اكتملت دفعة فهرسة الوثائق المنشورة',
                recordId: '',
                category: 'System',
                read: false,
                time: new Date().toISOString(),
              },
              ...s.notices,
            ],
          },
          role,
          'Indexing batch completed (simulation)',
          'اكتملت دفعة الفهرسة التجريبية',
          'RAG batch',
        ),
      )
      setRunning(false)
      notify(
        t(
          'Published records are now searchable. Drafts and returned documents were excluded.',
          'أصبحت السجلات المنشورة قابلة للبحث. استبعدت المسودات والوثائق المعادة.',
        ),
      )
    }, 1000)
  }
  return (
    <>
      <div className="p-tabs">
        {[
          ['staging', 'Staging & promotion', 'التجهيز والنشر'],
          ['index', 'Search indexing', 'فهرسة البحث'],
        ].map(([id, en, ar]) => (
          <button key={id} className={tab === id ? 'active' : ''} onClick={() => setTab(id)}>
            {t(en, ar)}
          </button>
        ))}
      </div>
      {tab === 'staging' ? (
        <>
          <div className="p-three">
            <div className="p-card">
              <span className="p-eyebrow">{t('STAGING ZONE', 'منطقة التجهيز')}</span>
              <h2>{visible.filter((r) => r.staging === 'Active').length}</h2>
              <p>{t('Original files awaiting governed publication', 'ملفات أصلية بانتظار النشر المحكوم')}</p>
            </div>
            <div className="p-card">
              <span className="p-eyebrow">{t('PROMOTION RECOVERY', 'استعادة النشر')}</span>
              <h2>{visible.filter((r) => r.status === 'Approved').length}</h2>
              <p>
                {t(
                  'Approved records ready for idempotent replay',
                  'سجلات معتمدة جاهزة لإعادة النشر دون تكرار',
                )}
              </p>
            </div>
            <div className="p-card">
              <span className="p-eyebrow">{t('CLEANUP POLICY', 'سياسة التنظيف')}</span>
              <h2>{local(state.settings.cleanup, lang)}</h2>
              <p>
                {t(
                  'After publication · returned originals stay in staging',
                  'بعد النشر · تبقى الملفات المعادة في التجهيز',
                )}
              </p>
            </div>
          </div>
          <div className="p-card no-pad">
            <div className="p-card-head">
              <h3>{t('Staging register', 'سجل التجهيز')}</h3>
              <span className="p-subtle">
                {t('Open a record to review or retry', 'افتح سجلاً للمراجعة أو إعادة المحاولة')}
              </span>
            </div>
            <DocumentTable
              rows={visible.filter((r) => !['Published', 'Archived', 'Disposed'].includes(r.status))}
              lang={lang}
              t={t}
              onOpen={onOpen}
              selectedId={selectedId}
            />
          </div>
          <div className="p-note">
            {t(
              'Abandoned drafts: 30-day expiry and owner reminders are shown as a proposed policy. Automatic cleanup, durable retries and cloud storage are simulated.',
              'المسودات المهجورة: انتهاء بعد ٣٠ يوماً وتنبيهات المالك سياسة مقترحة. التنظيف التلقائي وإعادة المحاولة والتخزين السحابي محاكاة.',
            )}
          </div>
        </>
      ) : (
        <>
          <section className="p-dark-banner">
            <div>
              <p className="p-eyebrow">{t('PUBLISHED KNOWLEDGE ONLY', 'المعرفة المنشورة فقط')}</p>
              <h2>{t('From an approved record to a grounded answer.', 'من سجل معتمد إلى إجابة موثقة.')}</h2>
              <p>
                {t(
                  'Metadata and approved content come together in a permission-aware search index.',
                  'تجتمع البيانات والمحتوى المعتمد في فهرس بحث يراعي الصلاحيات.',
                )}
              </p>
            </div>
            <button
              className="p-btn light"
              disabled={running || role !== 'System Administrator'}
              onClick={runIndex}
            >
              <RefreshCw size={16} className={running ? 'p-spin' : ''} />
              {running ? t('Indexing…', 'جارٍ الفهرسة…') : t('Run demo batch', 'تشغيل دفعة تجريبية')}
            </button>
          </section>
          <div className="p-pipeline">
            {[
              ['Collect', 'جمع'],
              ['Extract', 'استخراج'],
              ['Chunk', 'تجزئة'],
              ['Embed', 'تضمين'],
              ['Index', 'فهرسة'],
              ['Serve', 'إجابة'],
            ].map(([en, ar], i) => (
              <div key={en}>
                <span>{i + 1}</span>
                <strong>{t(en, ar)}</strong>
                <small>{t('Simulated', 'محاكاة')}</small>
              </div>
            ))}
          </div>
          <div className="p-note">
            {t(
              `${eligible.length} published records waiting · Schedule: ${state.settings.cadence}. The demo uses registered sample text, not OCR, embeddings or a live language model.`,
              `${eligible.length} سجلات منشورة بانتظار الفهرسة · الجدول: ${local(state.settings.cadence, lang)}. يستخدم العرض النص التجريبي المسجل وليس الاستخراج الضوئي أو نموذجاً لغوياً حياً.`,
            )}
          </div>
          <div className="p-card no-pad">
            <DocumentTable
              rows={visible.filter((r) => r.status === 'Published')}
              lang={lang}
              t={t}
              onOpen={onOpen}
              selectedId={selectedId}
            />
          </div>
        </>
      )}
    </>
  )
}
export function CorrespondencePage({ state, setState, role, lang, t, notify }: PageProps) {
  const [selected, setSelected] = useState('')
  const [creating, setCreating] = useState(false)
  const [reply, setReply] = useState('')
  const [query, setQuery] = useState('')
  const canWrite = capabilities(role).register
  const item = state.correspondence.find((c) => c.id === selected)
  const change = (id: string, patch: Partial<Correspondence>, action: string, ar: string) => {
    setState((s) =>
      addEvent(
        { ...s, correspondence: s.correspondence.map((c) => (c.id === id ? { ...c, ...patch } : c)) },
        role,
        action,
        ar,
        id,
      ),
    )
    notify(t('Correspondence updated in this demo.', 'تم تحديث المراسلة في العرض.'))
  }
  return (
    <>
      <div className="p-toolbar">
        <input
          aria-label={t('Search correspondence', 'بحث المراسلات')}
          placeholder={t('Search subject or reference…', 'ابحث بالموضوع أو المرجع…')}
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
        {canWrite && (
          <button className="p-btn primary" onClick={() => setCreating(true)}>
            <Plus size={16} />
            {t('New correspondence', 'مراسلة جديدة')}
          </button>
        )}
      </div>
      <div className="p-card no-pad">
        <div className="p-table-wrap">
          <table className="p-table">
            <thead>
              <tr>
                {[
                  t('Reference / subject', 'المرجع / الموضوع'),
                  t('Direction', 'الاتجاه'),
                  t('Assigned to', 'مسند إلى'),
                  t('Due date', 'الموعد'),
                  t('Status', 'الحالة'),
                ].map((h) => (
                  <th key={h}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {state.correspondence
                .filter((c) =>
                  `${c.subject} ${c.subjectAr} ${c.id}`.toLowerCase().includes(query.toLowerCase()),
                )
                .map((c) => (
                  <tr key={c.id}>
                    <td>
                      <button className="p-document-link" onClick={() => setSelected(c.id)}>
                        <span>
                          <strong>{lang === 'ar' ? c.subjectAr || c.subject : c.subject}</strong>
                          <small>
                            {c.id} · {c.party}
                          </small>
                        </span>
                      </button>
                    </td>
                    <td>
                      <Badge value={c.direction} lang={lang} />
                    </td>
                    <td>{local(c.assignee, lang)}</td>
                    <td>{c.due}</td>
                    <td>
                      <Badge value={c.status} lang={lang} />
                      {c.acknowledged && (
                        <small className="p-indexed">{t('Acknowledged', 'تم تأكيد القراءة')}</small>
                      )}
                    </td>
                  </tr>
                ))}
            </tbody>
          </table>
        </div>
      </div>
      <div className="p-note">
        {t(
          'Incoming, outgoing, memos and circulars share a reference and conversation thread. Replies and distribution are simulated locally; no messages are sent.',
          'الوارد والصادر والمذكرات والتعاميم تشترك في مرجع وسلسلة محادثة. الردود والتوزيع محاكاة محلية ولا يتم إرسال رسائل.',
        )}
      </div>
      {creating && (
        <Modal title={t('New correspondence', 'مراسلة جديدة')} onClose={() => setCreating(false)}>
          <form
            className="p-modal-body"
            onSubmit={(e) => {
              e.preventDefault()
              const f = new FormData(e.currentTarget)
              const newItem: Correspondence = {
                id: `COR-${new Date().getFullYear()}-${String(Date.now()).slice(-6)}`,
                subject: String(f.get('subject')),
                subjectAr: String(f.get('subjectAr')),
                party: String(f.get('party')),
                direction: String(f.get('direction')),
                assignee: String(f.get('assignee')),
                due: String(f.get('due')),
                status: 'Open',
                thread: [String(f.get('message'))],
                acknowledged: false,
              }
              setState((s) =>
                addEvent(
                  { ...s, correspondence: [newItem, ...s.correspondence] },
                  role,
                  'Correspondence registered',
                  'تم تسجيل المراسلة',
                  newItem.id,
                ),
              )
              setCreating(false)
              setSelected(newItem.id)
            }}
          >
            <div className="p-form-grid">
              <Field label={t('Subject *', 'الموضوع *')}>
                <input name="subject" required />
              </Field>
              <Field label={t('Arabic subject', 'الموضوع العربي')}>
                <input name="subjectAr" dir="rtl" />
              </Field>
              <Field label={t('Direction', 'الاتجاه')}>
                <select name="direction">
                  {['Incoming', 'Outgoing', 'Memo', 'Circular'].map((v) => (
                    <option key={v} value={v}>
                      {local(v, lang)}
                    </option>
                  ))}
                </select>
              </Field>
              <Field label={t('Sender / recipient *', 'المرسل / المستلم *')}>
                <input name="party" required />
              </Field>
              <Field label={t('Assign department', 'إسناد إلى الإدارة')}>
                <select name="assignee">
                  {departments.map((v) => (
                    <option key={v} value={v}>
                      {local(v, lang)}
                    </option>
                  ))}
                </select>
              </Field>
              <Field label={t('Due date *', 'الموعد *')}>
                <input type="date" name="due" required />
              </Field>
            </div>
            <Field label={t('Message *', 'الرسالة *')}>
              <textarea name="message" required rows={3} />
            </Field>
            <footer className="p-modal-actions">
              <button type="button" className="p-btn" onClick={() => setCreating(false)}>
                {t('Cancel', 'إلغاء')}
              </button>
              <button className="p-btn primary">{t('Register correspondence', 'تسجيل المراسلة')}</button>
            </footer>
          </form>
        </Modal>
      )}
      {item && (
        <Modal
          title={lang === 'ar' ? item.subjectAr || item.subject : item.subject}
          onClose={() => setSelected('')}
        >
          <div className="p-modal-body">
            <div className="p-detail-meta">
              <span>{item.id}</span>
              <Badge value={item.direction} lang={lang} />
              <Badge value={item.status} lang={lang} />
            </div>
            <div className="p-thread">
              {item.thread.map((message, i) => (
                <div key={i}>
                  <small>{i ? t('Reply / action', 'رد / إجراء') : item.party}</small>
                  <p>{message}</p>
                </div>
              ))}
            </div>
            {canWrite && (
              <>
                <Field label={t('Reply / action note', 'رد / ملاحظة إجراء')}>
                  <textarea value={reply} onChange={(e) => setReply(e.target.value)} />
                </Field>
                <div className="p-actions">
                  <button
                    className="p-btn primary"
                    disabled={!reply.trim() || item.status === 'Closed'}
                    onClick={() => {
                      change(
                        item.id,
                        { thread: [...item.thread, `${people[role]}: ${reply}`] },
                        'Correspondence reply recorded',
                        'تم تسجيل الرد',
                      )
                      setReply('')
                    }}
                  >
                    <Send size={15} />
                    {t('Record reply', 'تسجيل الرد')}
                  </button>
                  <button
                    className="p-btn"
                    onClick={() =>
                      change(
                        item.id,
                        { status: item.status === 'Open' ? 'Closed' : 'Open' },
                        'Correspondence status changed',
                        'تم تغيير حالة المراسلة',
                      )
                    }
                  >
                    {item.status === 'Open'
                      ? t('Close action', 'إغلاق الإجراء')
                      : t('Reopen action', 'إعادة فتح الإجراء')}
                  </button>
                </div>
                <Field label={t('Forward / assign to', 'تحويل / إسناد إلى')}>
                  <select
                    value={item.assignee}
                    onChange={(e) =>
                      change(
                        item.id,
                        {
                          assignee: e.target.value,
                          thread: [
                            ...item.thread,
                            `${t('Assigned to', 'مسند إلى')}: ${local(e.target.value, lang)}`,
                          ],
                        },
                        'Correspondence forwarded',
                        'تم تحويل المراسلة',
                      )
                    }
                  >
                    {departments.map((v) => (
                      <option key={v} value={v}>
                        {local(v, lang)}
                      </option>
                    ))}
                  </select>
                </Field>
              </>
            )}
            <button
              className="p-btn"
              disabled={item.acknowledged}
              onClick={() => change(item.id, { acknowledged: true }, 'Read acknowledgement', 'تأكيد القراءة')}
            >
              <Check size={15} />
              {item.acknowledged
                ? t('Reading acknowledged', 'تم تأكيد القراءة')
                : t('Acknowledge reading', 'تأكيد القراءة')}
            </button>
          </div>
        </Modal>
      )}
    </>
  )
}
export function RecordsPage(props: PageProps) {
  const { visible, lang, t, onOpen, state, selectedId } = props
  const [filter, setFilter] = useState('all')
  const rows = visible
    .filter((r) => ['Published', 'Archived', 'Disposed'].includes(r.status))
    .filter((r) =>
      filter === 'holds'
        ? !!r.hold
        : filter === 'due'
          ? new Date(r.expiry).getTime() < Date.now() + 30 * 86400000
          : filter === 'archive'
            ? r.status === 'Archived'
            : true,
    )
  return (
    <>
      <div className="p-three">
        <div className="p-card">
          <LockKeyhole size={22} />
          <h2>{visible.filter((r) => r.hold).length}</h2>
          <p>{t('Active legal holds', 'حجوزات قانونية نشطة')}</p>
        </div>
        <div className="p-card">
          <ShieldCheck size={22} />
          <h2>
            {state.settings.retentionYears} {t('years', 'سنوات')}
          </h2>
          <p>{t('Default retention · review before disposal', 'الحفظ الافتراضي · مراجعة قبل الإتلاف')}</p>
        </div>
        <div className="p-card">
          <Settings2 size={22} />
          <h2>{visible.filter((r) => r.declared).length}</h2>
          <p>{t('Declared records · metadata frozen', 'سجلات معلنة · بيانات مجمدة')}</p>
        </div>
      </div>
      <div className="p-tabs">
        {[
          ['all', 'All records', 'كل السجلات'],
          ['holds', 'Legal holds', 'الحجوزات القانونية'],
          ['due', 'Due within 30 days', 'مستحق خلال ٣٠ يوماً'],
          ['archive', 'Disposition queue', 'قائمة التصرف النهائي'],
        ].map(([id, en, ar]) => (
          <button key={id} className={filter === id ? 'active' : ''} onClick={() => setFilter(id)}>
            {t(en, ar)}
          </button>
        ))}
      </div>
      <div className="p-card no-pad">
        <DocumentTable rows={rows} lang={lang} t={t} onOpen={onOpen} selectedId={selectedId} />
      </div>
      <div className="p-note">
        {t(
          'Open a record → Governance to apply a hold, declare, archive, or approve simulated disposition. Legal holds override retention. Production deletion requires the agreed records policy.',
          'افتح السجل ← الحوكمة لتطبيق حجز أو إعلان السجل أو الأرشفة أو اعتماد الإتلاف التجريبي. الحجز القانوني يتجاوز الحفظ. الحذف الإنتاجي يتطلب سياسة سجلات معتمدة.',
        )}
      </div>
    </>
  )
}
export function AdminPage({ state, setState, role, lang, t, notify }: PageProps) {
  const [tab, setTab] = useState('routing')
  const [newClass, setNewClass] = useState('')
  const [settings, setSettings] = useState(state.settings)
  const [folderDepartment, setFolderDepartment] = useState(state.settings.department)
  const [folderRoots, setFolderRoots] = useState<SharePointItem[]>([])
  const [selectedFolder, setSelectedFolder] = useState<SharePointFolder | null>(null)
  const [folderLibraryName, setFolderLibraryName] = useState(t('Documents', 'المستندات'))
  const [folderMappings, setFolderMappings] = useState<DepartmentSharePointFolder[]>([])
  const [folderBusy, setFolderBusy] = useState(false)
  const [folderError, setFolderError] = useState('')
  const mappedFolders = folderMappings.filter((item) => item.department === folderDepartment)
  const selectedFolderAssigned = Boolean(
    selectedFolder && mappedFolders.some((item) => item.sharePointFolderId === selectedFolder.id),
  )
  const loadFolderAccess = async () => {
    setFolderBusy(true)
    setFolderError('')
    try {
      const [folders, mappings] = await Promise.all([
        listAdminSharePointFolders(),
        listDepartmentSharePointFolderAccess(),
      ])
      setFolderRoots(folders.items)
      setFolderLibraryName(folders.drive?.name || t('Documents', 'المستندات'))
      setFolderMappings(mappings)
      setSelectedFolder(null)
    } catch (problem) {
      setFolderError(
        problem instanceof Error
          ? problem.message
          : t('SharePoint folders could not be loaded. Please try again.', 'تعذر تحميل مجلدات شيربوينت. حاول مرة أخرى.'),
      )
    } finally {
      setFolderBusy(false)
    }
  }
  useEffect(() => {
    if (tab === 'folder-access') void loadFolderAccess()
  }, [tab])
  const loadAdminChildren = async (folderId: string) => (await listAdminSharePointFolders(folderId)).items
  const assignSharePointFolder = async () => {
    if (!selectedFolder || selectedFolderAssigned) return
    setFolderBusy(true)
    setFolderError('')
    try {
      const item = await addDepartmentSharePointFolderAccess(folderDepartment, selectedFolder)
      setFolderMappings((current) => [item, ...current])
      setSelectedFolder(null)
      notify(t('SharePoint folder assigned to this department.', 'تم ربط مجلد شيربوينت بهذه الإدارة.'))
    } catch (problem) {
      setFolderError(problem instanceof Error ? problem.message : String(problem))
    } finally {
      setFolderBusy(false)
    }
  }
  const removeSharePointFolder = async (folder: DepartmentSharePointFolder) => {
    setFolderBusy(true)
    setFolderError('')
    try {
      await removeDepartmentSharePointFolderAccess(folderDepartment, folder.sharePointFolderId)
      setFolderMappings((current) => current.filter((item) => item.mappingId !== folder.mappingId))
      notify(t('SharePoint folder access removed.', 'تمت إزالة صلاحية مجلد شيربوينت.'))
    } catch (problem) {
      setFolderError(problem instanceof Error ? problem.message : String(problem))
    } finally {
      setFolderBusy(false)
    }
  }
  const save = () => {
    setState((s) =>
      addEvent(
        {
          ...s,
          settings:
            role === 'Department Administrator'
              ? {
                  ...s.settings,
                  classes: settings.classes,
                  department: settings.department,
                  library: settings.library,
                }
              : settings,
        },
        role,
        'Configuration saved',
        'تم حفظ الإعدادات',
        'Admin Center',
      ),
    )
    notify(
      t(
        'Settings saved. New registrations use these defaults.',
        'تم حفظ الإعدادات. تستخدم التسجيلات الجديدة هذه القيم.',
      ),
    )
  }
  return (
    <>
      <div className="p-tabs">
        {[
          ['routing', 'Types & routing', 'الفئات والتوجيه'],
          ['policies', 'Workflow & lifecycle', 'سير العمل ودورة الحياة'],
          ['identity', 'Identity & access', 'الهوية والوصول'],
          ['folder-access', 'Folder access', 'صلاحية المجلدات'],
          ['integrations', 'Integrations', 'التكاملات'],
        ]
          .filter(([id]) => role !== 'Department Administrator' || id === 'routing')
          .map(([id, en, ar]) => (
            <button key={id} className={tab === id ? 'active' : ''} onClick={() => setTab(id)}>
              {t(en, ar)}
            </button>
          ))}
      </div>
      {tab === 'routing' && (
        <div className="p-two">
          <section className="p-card">
            <h3>{t('Document classes', 'فئات الوثائق')}</h3>
            <p>
              {t(
                'Controlled vocabulary for classification suggestions and human review.',
                'قيم مضبوطة تستخدم أثناء التصنيف اليدوي.',
              )}
            </p>
            <div className="p-tags">
              {settings.classes.map((c) => (
                <span key={c}>{local(c, lang)}</span>
              ))}
            </div>
            <form
              className="p-inline"
              onSubmit={(e) => {
                e.preventDefault()
                if (newClass.trim() && !settings.classes.includes(newClass.trim()))
                  setSettings((s) => ({ ...s, classes: [...s.classes, newClass.trim()] }))
                setNewClass('')
              }}
            >
              <input
                aria-label={t('New document class', 'فئة وثيقة جديدة')}
                value={newClass}
                onChange={(e) => setNewClass(e.target.value)}
                placeholder={t('Add a document class', 'إضافة فئة وثيقة')}
                required
              />
              <button className="p-btn">
                <Plus size={16} />
                {t('Add', 'إضافة')}
              </button>
            </form>
          </section>
          <section className="p-card">
            <h3>{t('Library routing', 'توجيه المكتبة')}</h3>
            <Field label={t('Default department', 'الإدارة الافتراضية')}>
              <select
                value={settings.department}
                disabled={role === 'Department Administrator'}
                onChange={(e) => setSettings((s) => ({ ...s, department: e.target.value }))}
              >
                {departments.map((v) => (
                  <option key={v} value={v}>
                    {local(v, lang)}
                  </option>
                ))}
              </select>
            </Field>
            <Field label={t('Target library root', 'جذر مكتبة الوجهة')}>
              <input
                value={settings.library}
                onChange={(e) => setSettings((s) => ({ ...s, library: e.target.value }))}
              />
            </Field>
            <div className="p-route">
              <span>{t('Department', 'الإدارة')}</span>
              <ArrowRight size={15} />
              <span>{t('Document class', 'فئة الوثيقة')}</span>
              <ArrowRight size={15} />
              <span>{t('Approved original', 'الأصل المعتمد')}</span>
            </div>
            <small>
              {t('Path preview', 'معاينة المسار')}: {settings.library}/{settings.department}/Policy
            </small>
          </section>
        </div>
      )}
      {tab === 'policies' && (
        <div className="p-two">
          <section className="p-card">
            <h3>{t('Workflow & retention defaults', 'إعدادات سير العمل والحفظ')}</h3>
            <Field label={t('Approval SLA (hours)', 'مدة الاعتماد (ساعات)')}>
              <input
                type="number"
                min="1"
                max="720"
                value={settings.sla}
                onChange={(e) =>
                  setSettings((s) => ({ ...s, sla: Math.max(1, Math.min(720, Number(e.target.value))) }))
                }
              />
            </Field>
            <Field label={t('Retention (years)', 'مدة الحفظ (سنوات)')}>
              <input
                type="number"
                min="1"
                max="100"
                value={settings.retentionYears}
                onChange={(e) =>
                  setSettings((s) => ({
                    ...s,
                    retentionYears: Math.max(1, Math.min(100, Number(e.target.value))),
                  }))
                }
              />
            </Field>
            <Field label={t('Staging after publication', 'التجهيز بعد النشر')}>
              <select
                value={settings.cleanup}
                onChange={(e) => setSettings((s) => ({ ...s, cleanup: e.target.value }))}
              >
                {['Retain', 'Purge'].map((v) => (
                  <option key={v} value={v}>
                    {local(v, lang)}
                  </option>
                ))}
              </select>
            </Field>
            <Field label={t('Indexing cadence (simulated)', 'جدول الفهرسة (محاكاة)')}>
              <select
                value={settings.cadence}
                onChange={(e) => setSettings((s) => ({ ...s, cadence: e.target.value }))}
              >
                {['Every 15 minutes', 'Hourly', 'Daily'].map((v) => (
                  <option key={v} value={v}>
                    {local(v, lang)}
                  </option>
                ))}
              </select>
            </Field>
          </section>
          <section className="p-card">
            <h3>{t('Notification & demo defaults', 'إعدادات التنبيهات والعرض')}</h3>
            {(
              [
                ['email', 'Email delivery simulation', 'محاكاة البريد الإلكتروني'],
                ['teams', 'Teams delivery simulation', 'محاكاة تيمز'],
                ['defaultArabic', 'Arabic on next demo sign-in', 'العربية عند الدخول التالي'],
                ['simulateFailure', 'Simulate publication failure', 'محاكاة فشل النشر'],
              ] as const
            ).map(([key, en, ar]) => (
              <label className="p-toggle" key={key}>
                <span>{t(en, ar)}</span>
                <input
                  type="checkbox"
                  checked={settings[key]}
                  onChange={(e) => setSettings((s) => ({ ...s, [key]: e.target.checked }))}
                />
              </label>
            ))}
            <Field label={t('Quiet hours (demo preference)', 'ساعات الهدوء (تفضيل تجريبي)')}>
              <input
                value={settings.quietHours}
                onChange={(e) => setSettings((s) => ({ ...s, quietHours: e.target.value }))}
              />
            </Field>
            <p className="p-note">
              <LockKeyhole size={16} />
              {t('Security notifications always remain enabled.', 'تبقى إشعارات الأمان مفعلة دائماً.')}
            </p>
          </section>
        </div>
      )}
      {tab === 'identity' && (
        <>
          <section className="p-dark-banner">
            <div>
              <p className="p-eyebrow">UAE PASS → MICROSOFT ENTRA ID → DOCAYA</p>
              <h2>{t('One identity. Governed access.', 'هوية واحدة. وصول محكوم.')}</h2>
              <p>
                {t(
                  'Persona switching demonstrates role and department permissions. No real authentication is performed.',
                  'تبديل الأدوار يوضح صلاحيات الدور والإدارة. لا يتم تنفيذ مصادقة حقيقية.',
                )}
              </p>
            </div>
            <ShieldCheck size={56} />
          </section>
          <div className="p-card no-pad">
            <div className="p-table-wrap">
              <table className="p-table">
                <thead>
                  <tr>
                    <th>{t('Persona', 'الدور')}</th>
                    {['Register', 'Approve', 'Records', 'Admin', 'Audit'].map((v, i) => (
                      <th key={v}>{t(v, ['تسجيل', 'اعتماد', 'سجلات', 'إدارة', 'تدقيق'][i])}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {roles.map((r) => (
                    <tr key={r}>
                      <td>
                        <strong>{local(r, lang)}</strong>
                        <small>{people[r]}</small>
                      </td>
                      {Object.values(capabilities(r)).map((v, i) => (
                        <td key={i}>{v ? <Check size={17} /> : '—'}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
          <div className="p-note">
            {t(
              'Demo ABAC: Secret is restricted to system administrators, records officers and auditors. Confidential and department-only records require the selected department. Viewers see published and archived records. This browser model is not a security boundary.',
              'قواعد العرض: السري للغاية متاح لمدير النظام ومسؤول السجلات والمدقق. السجلات السرية أو الخاصة بالإدارة تتطلب الإدارة المحددة. يرى القارئ المنشور والمؤرشف. نموذج المتصفح ليس حداً أمنياً إنتاجياً.',
            )}
          </div>
        </>
      )}
      {tab === 'folder-access' && (
        <div className="p-two p-folder-access">
          <section className="p-card">
            <div className="p-card-head">
              <div>
                <h3>{t('Department folder access', 'صلاحية مجلدات الإدارة')}</h3>
                <p>{t('Assign real SharePoint folders to each department.', 'اربط مجلدات شيربوينت الفعلية بكل إدارة.')}</p>
              </div>
              <button className="p-btn" disabled={folderBusy} onClick={() => void loadFolderAccess()}>
                <RefreshCw size={16} className={folderBusy ? 'p-spin' : ''} />
                {t('Refresh', 'تحديث')}
              </button>
            </div>
            <Field label={t('Department', 'الإدارة')}>
              <select value={folderDepartment} onChange={(event) => setFolderDepartment(event.target.value)}>
                {departments.map((department) => (
                  <option key={department} value={department}>
                    {local(department, lang)}
                  </option>
                ))}
              </select>
            </Field>
            {folderError && <div className="p-error">{folderError}</div>}
            <SharePointFolderTree
              library={{ id: 'sharepoint-drive', name: folderLibraryName }}
              roots={folderRoots}
              loading={folderBusy}
              error=""
              emptyMessage={t('No folders found in this SharePoint library.', 'لا توجد مجلدات في مكتبة شيربوينت هذه.')}
              loadingMessage={t('Loading SharePoint folders...', 'جار تحميل مجلدات شيربوينت...')}
              errorMessage={t(
                'SharePoint folders could not be loaded. Please try again.',
                'تعذر تحميل مجلدات شيربوينت. حاول مرة أخرى.',
              )}
              selectLabel={t('Select folder to assign', 'اختر مجلداً للربط')}
              showFiles
              onSelect={setSelectedFolder}
              onLoadChildren={loadAdminChildren}
            />
            <div className="p-actions end">
              <button
                className="p-btn primary"
                disabled={folderBusy || !selectedFolder || selectedFolderAssigned}
                onClick={assignSharePointFolder}
              >
                {selectedFolderAssigned ? t('Already assigned', 'مرتبط بالفعل') : t('Assign folder', 'ربط المجلد')}
              </button>
            </div>
          </section>
          <section className="p-card">
            <h3>{t('Assigned folders', 'المجلدات المرتبطة')}</h3>
            <p>{t('Registration only shows these mapped folders for the selected department.', 'يعرض التسجيل هذه المجلدات المرتبطة فقط للإدارة المحددة.')}</p>
            <div className="p-folder-list assigned">
              {mappedFolders.length === 0 ? (
                <p className="p-subtle">{t('No SharePoint folders assigned yet.', 'لم يتم ربط أي مجلدات شيربوينت بعد.')}</p>
              ) : (
                mappedFolders.map((folder) => (
                  <div className="p-folder-row" key={folder.mappingId || folder.sharePointFolderId}>
                    <span>
                      <strong>{folder.name}</strong>
                    </span>
                    <button className="p-btn" disabled={folderBusy} onClick={() => void removeSharePointFolder(folder)}>
                      <Trash2 size={15} />
                      {t('Remove', 'إزالة')}
                    </button>
                  </div>
                ))
              )}
            </div>
          </section>
        </div>
      )}
      {tab === 'integrations' && (
        <>
          <div className="p-two">
            {[
              [
                'UAE PASS + Entra ID',
                'National identity · OIDC / PKCE · MFA',
                'الهوية الوطنية · المصادقة متعددة العوامل',
              ],
              [
                'Azure SQL + Blob',
                'Registry and audit · staging-first capture',
                'السجل والتدقيق · الالتقاط في منطقة التجهيز',
              ],
              [
                'Azure AI Search + OpenAI',
                'Approved-content retrieval · citations · ACL trimming',
                'استرجاع المحتوى المعتمد · الاستشهادات · الصلاحيات',
              ],
              [
                'Teams + Email + SignalR',
                'Lifecycle notifications · durable delivery',
                'تنبيهات دورة الحياة · توصيل مستدام',
              ],
              [
                'Purview + Defender + Monitor',
                'Retention · malware / DLP · operational telemetry',
                'الحفظ · فحص البرمجيات الخبيثة · المراقبة',
              ],
            ].map(([name, en, ar]) => (
              <section className="p-card" key={name}>
              <div className="p-card-head">
                <h3>{name}</h3>
                <span className="p-badge internal">{t('Simulated', 'محاكاة')}</span>
              </div>
              <p>{t(en, ar)}</p>
              <small>
                {t(
                  'Production connection requires client configuration.',
                  'الربط الإنتاجي يتطلب إعدادات العميل.',
                )}
              </small>
            </section>
            ))}
          </div>
        </>
      )}
      {['routing', 'policies', 'integrations'].includes(tab) && (
        <div className="p-actions end">
          <button className="p-btn primary" disabled={!settings.library.trim()} onClick={save}>
            {t('Save configuration', 'حفظ الإعدادات')}
          </button>
        </div>
      )}
    </>
  )
}
export function ReportsPage({ state, visible, lang, t, role }: PageProps) {
  const [department, setDepartment] = useState('all')
  const rows = visible.filter((r) => department === 'all' || r.department === department)
  const published = rows.filter((r) => r.status === 'Published').length
  const pending = rows.filter((r) => r.status === 'PendingApproval')
  return (
    <>
      <div className="p-toolbar">
        <select
          aria-label={t('Report department', 'إدارة التقرير')}
          value={department}
          onChange={(e) => setDepartment(e.target.value)}
        >
          <option value="all">{t('All accessible departments', 'كل الإدارات المتاحة')}</option>
          {departments.map((v) => (
            <option key={v} value={v}>
              {local(v, lang)}
            </option>
          ))}
        </select>
        <div className="p-actions">
          <button
            className="p-btn"
            onClick={() =>
              downloadFile(
                'docaya-records-report.csv',
                csv([
                  [
                    t('Tracking number', 'رقم التتبع'),
                    t('Title', 'العنوان'),
                    t('Department', 'الإدارة'),
                    t('Status', 'الحالة'),
                    t('Sensitivity', 'الحساسية'),
                    t('Version', 'الإصدار'),
                    t('Legal hold', 'الحجز القانوني'),
                  ],
                  ...rows.map((r) => [
                    r.id,
                    lang === 'ar' ? r.titleAr || r.title : r.title,
                    local(r.department, lang),
                    local(r.status, lang),
                    local(r.sensitivity, lang),
                    String(r.version),
                    r.hold,
                  ]),
                ]),
              )
            }
          >
            <Download size={15} />
            {t('Export CSV', 'تصدير CSV')}
          </button>
          <button className="p-btn" onClick={() => window.print()}>
            {t('Print / save PDF', 'طباعة / حفظ PDF')}
          </button>
        </div>
      </div>
      <div className="p-four">
        {[
          [rows.length, t('Registered documents', 'وثائق مسجلة')],
          [published, t('Published records', 'سجلات منشورة')],
          [
            pending.filter((r) => new Date(r.due).getTime() < Date.now()).length,
            t('Overdue approvals', 'موافقات متأخرة'),
          ],
          [rows.filter((r) => !!r.hold).length, t('Active legal holds', 'حجوزات نشطة')],
        ].map(([n, label]) => (
          <div className="p-card" key={label}>
            <h2>{n}</h2>
            <p>{label}</p>
          </div>
        ))}
      </div>
      <div className="p-two">
        <section className="p-card">
          <h3>{t('Volume by department', 'الحجم حسب الإدارة')}</h3>
          <div className="p-horizontal-chart">
            {departments.map((d) => (
              <div key={d}>
                <span>{local(d, lang)}</span>
                <div>
                  <i
                    style={{
                      width: `${(rows.filter((r) => r.department === d).length / Math.max(1, rows.length)) * 100}%`,
                    }}
                  />
                </div>
                <strong>{rows.filter((r) => r.department === d).length}</strong>
              </div>
            ))}
          </div>
        </section>
        <section className="p-card">
          <h3>{t('Lifecycle distribution', 'توزيع دورة الحياة')}</h3>
          <div className="p-horizontal-chart">
            {['Draft', 'PendingApproval', 'Returned', 'Approved', 'Published', 'Archived'].map((status) => (
              <div key={status}>
                <span>{local(status, lang)}</span>
                <div>
                  <i
                    style={{
                      width: `${(rows.filter((r) => r.status === status).length / Math.max(1, rows.length)) * 100}%`,
                    }}
                  />
                </div>
                <strong>{rows.filter((r) => r.status === status).length}</strong>
              </div>
            ))}
          </div>
        </section>
        <section className="p-card">
          <h3>{t('Governance indicators', 'مؤشرات الحوكمة')}</h3>
          <dl className="p-facts">
            <dt>{t('Metadata completeness', 'اكتمال البيانات')}</dt>
            <dd>
              {Math.round(
                (rows.filter((r) => r.title && r.kind && r.summary).length / Math.max(1, rows.length)) * 100,
              )}
              %
            </dd>
            <dt>{t('Published records indexed', 'السجلات المنشورة المفهرسة')}</dt>
            <dd>
              {rows.filter((r) => r.indexed && r.status === 'Published').length} / {published}
            </dd>
            <dt>{t('Publication backlog', 'قائمة انتظار النشر')}</dt>
            <dd>{rows.filter((r) => r.status === 'Approved').length}</dd>
            <dt>{t('Report scope', 'نطاق التقرير')}</dt>
            <dd>{local(role, lang)}</dd>
          </dl>
        </section>
        <section className="p-card">
          <h3>{t('Reporting roadmap', 'خطة التقارير')}</h3>
          <p>
            {t(
              'Cycle-time trends, access anomalies, scheduled email reports and native XLSX exports require production event history and connected services.',
              'اتجاهات زمن المعالجة وحالات الوصول غير المعتادة والتقارير المجدولة وتصدير XLSX تتطلب سجل أحداث إنتاجياً وخدمات متصلة.',
            )}
          </p>
          <p className="p-subtle">
            {t(
              `This report is calculated from ${rows.length} accessible demo records and ${state.audit.length} local events.`,
              `يحسب التقرير من ${rows.length} سجلاً تجريبياً متاحاً و${state.audit.length} حدثاً محلياً.`,
            )}
          </p>
        </section>
      </div>
    </>
  )
}
export function AuditPage({ state, visible, lang, t }: PageProps) {
  const [query, setQuery] = useState('')
  const [view, setView] = useState<'document' | 'user'>('document')
  const [actor, setActor] = useState('')
  const [doc, setDoc] = useState('')
  const accessible = state.audit.filter(
    (a) => !a.object.startsWith('DOC-') || visible.some((r) => r.id === a.object),
  )
  const actors = [...new Set(accessible.map((a) => a.actor))].sort()
  const documents = [
    ...new Set(accessible.filter((a) => a.object.startsWith('DOC-')).map((a) => a.object)),
  ].sort()
  const rows = accessible.filter(
    (a) =>
      (view === 'document'
        ? a.object.startsWith('DOC-') && (!doc || a.object === doc)
        : !actor || a.actor === actor) &&
      `${a.actor} ${a.action} ${a.object} ${a.detail}`.toLowerCase().includes(query.toLowerCase()),
  )
  return (
    <>
      <div className="p-note">
        <ShieldCheck size={20} />
        {t(
          'Demo activity log with actor, role, tracking reference and event ID. Production hash chaining, trusted timestamps and immutable storage are design requirements, not provided by browser storage.',
          'سجل نشاط تجريبي يتضمن الفاعل والدور والمرجع ومعرّف الحدث. تسلسل البصمات والطوابع الزمنية الموثوقة والتخزين غير القابل للتغيير متطلبات إنتاجية لا يوفرها المتصفح.',
        )}
      </div>
      <div className="p-tabs">
        {(
          [
            ['document', 'Document audit trail', 'سجل تدقيق الوثائق'],
            ['user', 'User audit trail', 'سجل تدقيق المستخدمين'],
          ] as const
        ).map(([id, en, ar]) => (
          <button key={id} className={view === id ? 'active' : ''} onClick={() => setView(id)}>
            {t(en, ar)}
          </button>
        ))}
      </div>
      <div className="p-toolbar">
        {view === 'document' ? (
          <select
            value={doc}
            onChange={(e) => setDoc(e.target.value)}
            aria-label={t('Filter by document', 'تصفية حسب الوثيقة')}
          >
            <option value="">{t('All documents', 'كل الوثائق')}</option>
            {documents.map((id) => (
              <option key={id} value={id}>
                {id} ·{' '}
                {(() => {
                  const r = visible.find((x) => x.id === id)
                  return r ? (lang === 'ar' ? r.titleAr || r.title : r.title) : ''
                })()}
              </option>
            ))}
          </select>
        ) : (
          <select
            value={actor}
            onChange={(e) => setActor(e.target.value)}
            aria-label={t('Filter by user', 'تصفية حسب المستخدم')}
          >
            <option value="">{t('All users', 'كل المستخدمين')}</option>
            {actors.map((name) => (
              <option key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
        )}
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          aria-label={t('Filter audit log', 'تصفية سجل التدقيق')}
          placeholder={t(
            'Filter by actor, action or tracking number…',
            'تصفية حسب الفاعل أو الإجراء أو رقم التتبع…',
          )}
        />
        <button
          className="p-btn"
          onClick={() =>
            downloadFile(
              `docaya-audit-${view}.csv`,
              csv([
                ['Event ID', 'Timestamp', 'Actor', 'Role', 'Action', 'Object', 'Detail'],
                ...rows.map((a) => [a.id, a.time, a.actor, a.role, a.action, a.object, a.detail]),
              ]),
            )
          }
        >
          <Download size={15} />
          {t('Export audit', 'تصدير التدقيق')}
        </button>
      </div>
      <div className="p-card no-pad">
        <div className="p-table-wrap">
          <table className="p-table">
            <thead>
              <tr>
                {[
                  t('Timestamp / event', 'الوقت / الحدث'),
                  t('Actor', 'الفاعل'),
                  t('Action', 'الإجراء'),
                  t('Object / detail', 'العنصر / التفاصيل'),
                ].map((h) => (
                  <th key={h}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((a) => (
                <tr key={a.id}>
                  <td>
                    {new Date(a.time).toLocaleString(lang)}
                    <small>{a.id.slice(0, 8)}</small>
                  </td>
                  <td>
                    {a.actor}
                    <small>{local(a.role, lang)}</small>
                  </td>
                  <td>{lang === 'ar' ? a.actionAr : a.action}</td>
                  <td>
                    {a.object}
                    <small>{a.detail}</small>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {!rows.length && <Empty t={t} />}
      </div>
    </>
  )
}
