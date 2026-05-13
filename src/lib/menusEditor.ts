import type {
  MenuItemResponse,
  MenuNavigationType,
  MenuPageOption,
  SaveMenuItemRequest,
} from '../api/menusApi'

export type MenuEditorItem = {
  client_id: string
  label: string
  navigation_type: MenuNavigationType
  page_id: number | null
  external_url: string
  open_in_new_tab: boolean
  children: MenuEditorItem[]
}

export type MenuItemDraft = {
  label: string
  navigation_type: MenuNavigationType
  page_id: string
  external_url: string
  parent_page_id: string
  open_in_new_tab: boolean
}

export type PageParentMenuOption = {
  value: string
  label: string
}

type MenuItemLocation = {
  item: MenuEditorItem
  path: number[]
  parentPath: number[]
}

export function createDefaultMenuItemDraft(): MenuItemDraft {
  return {
    label: '',
    navigation_type: 'pages',
    page_id: '',
    external_url: '',
    parent_page_id: '',
    open_in_new_tab: false,
  }
}

export function buildMenuEditorItems(items: MenuItemResponse[]): MenuEditorItem[] {
  return items.map((item) => ({
    client_id: createClientId(),
    label: item.label,
    navigation_type: item.navigation_type,
    page_id: item.page_id ?? null,
    external_url: item.external_url ?? '',
    open_in_new_tab: Boolean(item.open_in_new_tab),
    children: buildMenuEditorItems(item.children ?? []),
  }))
}

export function buildMenuItemDraft(
  item: MenuEditorItem,
  parentPageId: string = '',
): MenuItemDraft {
  return {
    label: item.label,
    navigation_type: item.navigation_type,
    page_id: item.page_id ? String(item.page_id) : '',
    external_url: item.external_url,
    parent_page_id: item.navigation_type === 'external_link' ? parentPageId : '',
    open_in_new_tab: item.open_in_new_tab,
  }
}

export function buildSaveMenuItems(items: MenuEditorItem[]): SaveMenuItemRequest[] {
  return items.map((item) => ({
    label: item.label.trim(),
    navigation_type: item.navigation_type,
    page_id: item.navigation_type === 'pages' ? item.page_id : null,
    external_url: item.navigation_type === 'external_link' ? item.external_url.trim() : '',
    open_in_new_tab: item.navigation_type === 'external_link' ? item.open_in_new_tab : false,
    children: buildSaveMenuItems(item.children),
  }))
}

export function serializeMenuItems(items: MenuEditorItem[]) {
  return JSON.stringify(buildSaveMenuItems(items))
}

export function collectUsedPageIds(
  items: MenuEditorItem[],
  excludeClientId?: string,
): Set<number> {
  const used = new Set<number>()

  const walk = (items: MenuEditorItem[]) => {
    for (const item of items) {
      if (item.client_id !== excludeClientId && item.page_id) {
        used.add(item.page_id)
      }
      walk(item.children)
    }
  }
  walk(items)

  return used
}

export function listPageParentMenuOptions(items: MenuEditorItem[]): PageParentMenuOption[] {
  const options: PageParentMenuOption[] = []

  const walk = (items: MenuEditorItem[], depth: number) => {
    for (const item of items) {
      if (item.navigation_type === 'pages' && item.page_id) {
        const prefix = depth > 0 ? `${'  '.repeat(depth)}- ` : ''
        options.push({
          value: String(item.page_id),
          label: `${prefix}${item.label}`,
        })
      }
      walk(item.children, depth + 1)
    }
  }
  walk(items, 0)

  return options
}

export function findMenuItemByClientId(
  items: MenuEditorItem[],
  clientId: string,
): MenuItemLocation | null {
  return findMenuItem(items, clientId, 'client')
}

export function findMenuItemByPageId(
  items: MenuEditorItem[],
  pageId: number,
): MenuItemLocation | null {
  return findMenuItem(items, pageId, 'page')
}

export function findParentPageIdForItem(
  items: MenuEditorItem[],
  clientId: string,
): string {
  const location = findMenuItemByClientId(items, clientId)
  if (!location || location.parentPath.length === 0) {
    return ''
  }

  const parent = getItemAtPath(items, location.parentPath)
  if (!parent || !parent.page_id) {
    return ''
  }

  return String(parent.page_id)
}

export function validateMenuItemDraft(
  draft: MenuItemDraft,
  items: MenuEditorItem[],
  pageOptions: MenuPageOption[],
  editingClientId?: string,
  hasLockedChildren: boolean = false,
): string | null {
  if (!draft.label.trim()) {
    return 'Menu label is required.'
  }

  if (draft.navigation_type === 'pages') {
    const pageId = parseNumericId(draft.page_id)
    if (!pageId) {
      return 'Select a published page.'
    }

    const page = pageOptions.find((option) => option.id === pageId)
    if (!page) {
      return 'Select a valid published page.'
    }

    const usedPageIds = collectUsedPageIds(items, editingClientId)
    if (usedPageIds.has(pageId)) {
      return 'This page has already been added to the menu.'
    }

    if (page.parent_id) {
      const parentItem = findMenuItemByPageId(items, page.parent_id)
      const editingLocation =
        editingClientId ? findMenuItemByClientId(items, editingClientId) : null
      const isSameParent =
        editingLocation &&
        editingLocation.parentPath.length > 0 &&
        getItemAtPath(items, editingLocation.parentPath)?.page_id === page.parent_id
      if (!parentItem && !isSameParent) {
        return 'Add the parent page to the menu before adding this sub-page.'
      }
    }

    if (hasLockedChildren && editingClientId) {
      const existing = findMenuItemByClientId(items, editingClientId)
      if (existing?.item.page_id !== pageId) {
        return 'Items with child links can only change their label.'
      }
    }

    return null
  }

  if (!isValidHttpUrl(draft.external_url)) {
    return 'Enter a valid absolute URL starting with http:// or https://.'
  }

  if (hasLockedChildren) {
    return 'Items with child links can only change their label.'
  }

  const parentPageId = parseNumericId(draft.parent_page_id)
  if (parentPageId) {
    const parentItem = findMenuItemByPageId(items, parentPageId)
    const editingLocation =
      editingClientId ? findMenuItemByClientId(items, editingClientId) : null
    const isSameParent =
      editingLocation &&
      editingLocation.parentPath.length > 0 &&
      getItemAtPath(items, editingLocation.parentPath)?.page_id === parentPageId
    if (!parentItem && !isSameParent) {
      return 'Select a parent page that already exists in the menu.'
    }
  }

  return null
}

export function applyMenuItemDraft(
  items: MenuEditorItem[],
  draft: MenuItemDraft,
  pageOptions: MenuPageOption[],
  editingClientId?: string,
): MenuEditorItem[] {
  const existing = editingClientId ? findMenuItemByClientId(items, editingClientId) : null
  const baseItems = editingClientId ? removeMenuItem(items, editingClientId) : cloneMenuItems(items)
  const pageId = draft.navigation_type === 'pages' ? parseNumericId(draft.page_id) : null
  const nextItem: MenuEditorItem = {
    client_id: existing?.item.client_id ?? createClientId(),
    label: draft.label.trim(),
    navigation_type: draft.navigation_type,
    page_id: pageId,
    external_url: draft.navigation_type === 'external_link' ? draft.external_url.trim() : '',
    open_in_new_tab:
      draft.navigation_type === 'external_link' ? Boolean(draft.open_in_new_tab) : false,
    children:
      existing?.item.children && draft.navigation_type === 'pages'
        ? cloneMenuItems(existing.item.children)
        : [],
  }

  return insertMenuItem(baseItems, nextItem, draft, pageOptions)
}

export function removeMenuItem(items: MenuEditorItem[], clientId: string): MenuEditorItem[] {
  return items
    .filter((item) => item.client_id !== clientId)
    .map((item) => ({
      ...item,
      children: removeMenuItem(item.children, clientId),
    }))
}

export function moveMenuItem(
  items: MenuEditorItem[],
  clientId: string,
  direction: 'up' | 'down',
): MenuEditorItem[] {
  const location = findMenuItemByClientId(items, clientId)
  if (!location) {
    return items
  }

  const nextItems = cloneMenuItems(items)
  const siblings = getSiblingArray(nextItems, location.parentPath)
  const index = location.path[location.path.length - 1]
  const targetIndex = direction === 'up' ? index - 1 : index + 1
  if (targetIndex < 0 || targetIndex >= siblings.length) {
    return items
  }

  const [moved] = siblings.splice(index, 1)
  siblings.splice(targetIndex, 0, moved)
  return nextItems
}

export function reorderMenuItems(
  items: MenuEditorItem[],
  draggedClientId: string,
  targetClientId: string,
): MenuEditorItem[] {
  const dragged = findMenuItemByClientId(items, draggedClientId)
  const target = findMenuItemByClientId(items, targetClientId)
  if (!dragged || !target || draggedClientId === targetClientId) {
    return items
  }

  if (dragged.parentPath.join('.') !== target.parentPath.join('.')) {
    return items
  }

  const nextItems = cloneMenuItems(items)
  const siblings = getSiblingArray(nextItems, dragged.parentPath)
  const fromIndex = dragged.path[dragged.path.length - 1]
  const toIndex = target.path[target.path.length - 1]
  const [moved] = siblings.splice(fromIndex, 1)
  siblings.splice(toIndex, 0, moved)
  return nextItems
}

export function canMoveMenuItem(
  items: MenuEditorItem[],
  clientId: string,
  direction: 'up' | 'down',
): boolean {
  const location = findMenuItemByClientId(items, clientId)
  if (!location) {
    return false
  }

  const siblings = getSiblingArray(items, location.parentPath)
  const index = location.path[location.path.length - 1]
  if (direction === 'up') {
    return index > 0
  }
  return index < siblings.length - 1
}

export function getMenuItemSummary(
  item: MenuEditorItem,
  pageOptions: MenuPageOption[],
): string {
  if (item.navigation_type === 'external_link') {
    return item.external_url
  }

  const page = pageOptions.find((option) => option.id === item.page_id)
  if (!page) {
    return 'Linked page unavailable'
  }
  return `Pages: ${page.page_title} (${page.url_slug})`
}

function insertMenuItem(
  items: MenuEditorItem[],
  nextItem: MenuEditorItem,
  draft: MenuItemDraft,
  pageOptions: MenuPageOption[],
): MenuEditorItem[] {
  const nextItems = cloneMenuItems(items)
  const parentPageId = resolveParentPageId(draft, pageOptions)

  if (!parentPageId) {
    nextItems.push(nextItem)
    return nextItems
  }

  const parentLocation = findMenuItemByPageId(nextItems, parentPageId)
  if (!parentLocation) {
    throw new Error('Parent page is not present in the current menu tree.')
  }

  parentLocation.item.children = [...parentLocation.item.children, nextItem]
  return nextItems
}

function resolveParentPageId(
  draft: MenuItemDraft,
  pageOptions: MenuPageOption[],
): number | null {
  if (draft.navigation_type === 'external_link') {
    return parseNumericId(draft.parent_page_id)
  }

  const pageId = parseNumericId(draft.page_id)
  if (!pageId) {
    return null
  }

  return pageOptions.find((option) => option.id === pageId)?.parent_id ?? null
}

function findMenuItem(
  items: MenuEditorItem[],
  value: string | number,
  mode: 'client' | 'page',
  parentPath: number[] = [],
): MenuItemLocation | null {
  for (const [index, item] of items.entries()) {
    const path = [...parentPath, index]
    const isMatch =
      mode === 'client'
        ? item.client_id === value
        : typeof value === 'number' && item.page_id === value
    if (isMatch) {
      return {
        item,
        path,
        parentPath,
      }
    }

    const childMatch = findMenuItem(item.children, value, mode, path)
    if (childMatch) {
      return childMatch
    }
  }

  return null
}

function getItemAtPath(items: MenuEditorItem[], path: number[]): MenuEditorItem | null {
  let currentItems = items
  let current: MenuEditorItem | null = null

  for (const index of path) {
    current = currentItems[index] ?? null
    if (!current) {
      return null
    }
    currentItems = current.children
  }

  return current
}

function getSiblingArray(items: MenuEditorItem[], parentPath: number[]): MenuEditorItem[] {
  if (parentPath.length === 0) {
    return items
  }

  const parent = getItemAtPath(items, parentPath)
  return parent?.children ?? items
}

function cloneMenuItems(items: MenuEditorItem[]): MenuEditorItem[] {
  return items.map((item) => ({
    ...item,
    children: cloneMenuItems(item.children),
  }))
}

function parseNumericId(value: string): number | null {
  const trimmed = value.trim()
  if (!trimmed) {
    return null
  }

  const parsed = Number.parseInt(trimmed, 10)
  return Number.isFinite(parsed) && parsed > 0 ? parsed : null
}

function isValidHttpUrl(value: string) {
  try {
    const parsed = new URL(value.trim())
    return parsed.protocol === 'http:' || parsed.protocol === 'https:'
  } catch {
    return false
  }
}

function createClientId() {
  if (typeof globalThis.crypto?.randomUUID === 'function') {
    return globalThis.crypto.randomUUID()
  }
  return `menu-item-${Math.random().toString(36).slice(2, 10)}`
}
