import type { SharePointFolder, SharePointItem } from '../types'

export const foldersOnly = (items: SharePointItem[]) =>
  items.filter((item): item is SharePointFolder => item.type === 'folder')
