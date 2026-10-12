import { Bot, RefreshCw, Send, X } from 'lucide-react'
import { useRef, useState, type FormEvent } from 'react'
import { useDialogFocus } from '../../hooks/useDialogFocus'
import { askAssistant, ApiError, type AssistantReference } from '../../services/api'

type Exchange = {
  id: string
  question: string
  answer?: string
  references?: AssistantReference[]
  error?: string
}

const examplePrompts = [
  'Which documents need review soon?',
  'Who owns the Data Protection Policy?',
  'Show restricted documents awaiting approval',
]

export function AskDocayaPanel({ onClose }: { onClose: () => void }) {
  const dialogRef = useDialogFocus(onClose)
  const [question, setQuestion] = useState('')
  const [busy, setBusy] = useState(false)
  const [exchanges, setExchanges] = useState<Exchange[]>([])
  const listRef = useRef<HTMLDivElement>(null)

  const ask = async (prompt: string, existingId?: string) => {
    const trimmed = prompt.trim()
    if (!trimmed || busy) return
    const id = existingId || `exchange-${Date.now()}`
    if (existingId) {
      setExchanges((rows) => rows.map((row) => (row.id === id ? { id, question: trimmed } : row)))
    } else {
      setExchanges((rows) => [...rows, { id, question: trimmed }])
    }
    setQuestion('')
    setBusy(true)
    try {
      const { answer, references } = await askAssistant(trimmed)
      setExchanges((rows) => rows.map((row) => (row.id === id ? { ...row, answer, references } : row)))
    } catch (problem) {
      const message =
        problem instanceof ApiError && problem.code === 'ASSISTANT_UNAVAILABLE'
          ? 'Ask Docaya is not configured for this environment yet. An administrator needs to connect an Azure OpenAI deployment.'
          : problem instanceof Error
            ? problem.message
            : 'Ask Docaya could not answer that question.'
      setExchanges((rows) => rows.map((row) => (row.id === id ? { ...row, error: message } : row)))
    } finally {
      setBusy(false)
      requestAnimationFrame(() => listRef.current?.scrollTo({ top: listRef.current.scrollHeight }))
    }
  }

  const submit = (event: FormEvent) => {
    event.preventDefault()
    void ask(question)
  }

  return (
    <div
      className="modal-backdrop"
      role="presentation"
      onMouseDown={(event) => event.target === event.currentTarget && onClose()}
    >
      <section
        ref={dialogRef}
        className="assistant-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="ask-docaya-title"
      >
        <header className="modal-header">
          <div>
            <p className="kicker">GOVERNANCE ASSISTANT</p>
            <h2 id="ask-docaya-title">
              <Bot size={20} /> Ask Docaya
            </h2>
            <span className="number-preview">Answers are grounded in your controlled documents</span>
          </div>
          <button className="close-button" onClick={onClose} aria-label="Close Ask Docaya">
            <X size={20} />
          </button>
        </header>
        <div className="assistant-body" ref={listRef}>
          {!exchanges.length && (
            <div className="assistant-empty">
              <p className="empty-copy">
                Ask about controlled documents, status, ownership, review dates or governance rules. Answers use
                only records returned by the authorized API.
              </p>
              <div className="prompt-chips" aria-label="Example Ask Docaya prompts">
                {examplePrompts.map((prompt) => (
                  <button type="button" key={prompt} onClick={() => void ask(prompt)} disabled={busy}>
                    {prompt}
                  </button>
                ))}
              </div>
            </div>
          )}
          {exchanges.map((exchange) => (
            <div className="assistant-exchange" key={exchange.id}>
              <p className="assistant-question">{exchange.question}</p>
              {exchange.error ? (
                <div className="assistant-error">
                  <p className="form-message" role="alert">
                    {exchange.error}
                  </p>
                  <button type="button" className="text-button" disabled={busy} onClick={() => void ask(exchange.question, exchange.id)}>
                    <RefreshCw size={14} /> Retry
                  </button>
                </div>
              ) : exchange.answer ? (
                <div className="assistant-answer">
                  <p>{exchange.answer}</p>
                  {exchange.references && exchange.references.length > 0 && (
                    <ul className="assistant-references">
                      {exchange.references.map((reference) => (
                        <li key={reference.id}>
                          <strong>{reference.number}</strong>
                          <span>{reference.title}</span>
                          <small>
                            {reference.status} · {reference.department} · Next review {reference.nextReview}
                          </small>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              ) : (
                <p className="assistant-answer assistant-thinking" role="status">
                  Thinking…
                </p>
              )}
            </div>
          ))}
        </div>
        <form className="assistant-composer" onSubmit={submit}>
          <input
            value={question}
            onChange={(event) => setQuestion(event.target.value)}
            placeholder="Ask about a controlled document…"
            aria-label="Ask Docaya a question"
            disabled={busy}
          />
          <button type="submit" className="gold-button" disabled={busy || !question.trim()}>
            <Send size={16} /> Send
          </button>
        </form>
      </section>
    </div>
  )
}
