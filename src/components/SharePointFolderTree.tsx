import { ChevronDown, ChevronRight, FileText, Folder, FolderOpen } from 'lucide-react'
import { useState } from 'react'
import type { SharePointFolder, SharePointItem } from '../types'

type LibraryRoot = {
  id: string
  name: string
}

type Props = {
  library: LibraryRoot
  roots: SharePointItem[]
  selectedId?: string
  loading?: boolean
  error?: string
  emptyMessage: string
  loadingMessage: string
  errorMessage: string
  childEmptyMessage?: string
  selectLabel?: string
  showFiles?: boolean
  onSelect?: (folder: SharePointFolder) => void
  onLoadChildren?: (folderId: string) => Promise<SharePointItem[]>
}

type NodeState = {
  expanded: boolean
  loading: boolean
  error: string
  children?: SharePointItem[]
}

export function SharePointFolderTree({
  library,
  roots,
  selectedId = '',
  loading = false,
  error = '',
  emptyMessage,
  loadingMessage,
  errorMessage,
  childEmptyMessage = 'No child folders.',
  selectLabel = 'Select folder',
  showFiles = false,
  onSelect,
  onLoadChildren,
}: Props) {
  const [nodes, setNodes] = useState<Record<string, NodeState>>({})

  const toggle = async (folder: SharePointFolder) => {
    const current = nodes[folder.id]
    if (current?.expanded) {
      setNodes((state) => ({ ...state, [folder.id]: { ...current, expanded: false } }))
      return
    }
    if (!onLoadChildren) {
      setNodes((state) => ({ ...state, [folder.id]: { ...(current || emptyNode), expanded: true } }))
      return
    }
    if (current?.children) {
      setNodes((state) => ({ ...state, [folder.id]: { ...current, expanded: true } }))
      return
    }
    setNodes((state) => ({ ...state, [folder.id]: { expanded: true, loading: true, error: '' } }))
    try {
      const children = await onLoadChildren(folder.id)
      setNodes((state) => ({
        ...state,
        [folder.id]: { expanded: true, loading: false, error: '', children },
      }))
    } catch (problem) {
      setNodes((state) => ({
        ...state,
        [folder.id]: {
          expanded: true,
          loading: false,
          error: problem instanceof Error ? problem.message : errorMessage,
          children: [],
        },
      }))
    }
  }

  const renderItem = (item: SharePointItem, depth: number, trail: SharePointFolder[]) => {
    if (item.type === 'file') {
      if (!showFiles) return null
      return (
        <li key={item.id}>
          <div className="sp-folder-row sp-file-row" style={{ '--sp-depth': depth } as React.CSSProperties}>
            <span className="sp-folder-toggle" aria-hidden="true" />
            <span className="sp-folder-name sp-file-name">
              <FileText size={17} />
              <span>{item.name}</span>
            </span>
          </div>
        </li>
      )
    }

    const state = nodes[item.id] || emptyNode
    const selected = selectedId === item.id
    const nextTrail = [...trail, item]
    return (
      <li key={item.id}>
        <div
          className={`sp-folder-row ${selected ? 'selected' : ''}`}
          style={{ '--sp-depth': depth } as React.CSSProperties}
        >
          {onLoadChildren ? (
            <button
              type="button"
              className="sp-folder-toggle"
              aria-label={`${state.expanded ? 'Collapse' : 'Expand'} ${item.name}`}
              aria-expanded={state.expanded}
              onClick={() => void toggle(item)}
            >
              {state.expanded ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
            </button>
          ) : (
            <span className="sp-folder-toggle" aria-hidden="true" />
          )}
          <button
            type="button"
            className="sp-folder-name"
            aria-label={`${selectLabel}: ${nextTrail.map((item) => item.name).join(' / ')}`}
            aria-current={selected ? 'true' : undefined}
            onClick={() => onSelect?.(item)}
          >
            {state.expanded ? <FolderOpen size={18} /> : <Folder size={18} />}
            <span>{item.name}</span>
          </button>
        </div>
        {state.expanded && (
          <div className="sp-folder-children">
            {state.loading && <p className="sp-folder-state">{loadingMessage}</p>}
            {state.error && <p className="sp-folder-state error">{state.error}</p>}
            {!state.loading && !state.error && state.children?.length === 0 && (
              <p className="sp-folder-state">{childEmptyMessage}</p>
            )}
            {Boolean(state.children?.length) && (
              <ul>{state.children?.map((child) => renderItem(child, depth + 1, nextTrail))}</ul>
            )}
          </div>
        )}
      </li>
    )
  }

  return (
    <div className="sp-folder-tree">
      <div className="sp-folder-root">
        <FolderOpen size={18} />
        <span>{library.name}</span>
      </div>
      {loading && <p className="sp-folder-state">{loadingMessage}</p>}
      {error && <p className="sp-folder-state error">{error}</p>}
      {!loading && !error && roots.length === 0 && <p className="sp-folder-state">{emptyMessage}</p>}
      {!loading && !error && roots.length > 0 && <ul>{roots.map((item) => renderItem(item, 0, []))}</ul>}
    </div>
  )
}

export const foldersOnly = (items: SharePointItem[]) =>
  items.filter((item): item is SharePointFolder => item.type === 'folder')

const emptyNode: NodeState = { expanded: false, loading: false, error: '' }
