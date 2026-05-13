import Dialog from '@mui/material/Dialog'
import DialogActions from '@mui/material/DialogActions'
import DialogContent from '@mui/material/DialogContent'
import DialogContentText from '@mui/material/DialogContentText'
import DialogTitle from '@mui/material/DialogTitle'
import DeleteOutlineOutlined from '@mui/icons-material/DeleteOutlineOutlined'
import DragIndicatorOutlined from '@mui/icons-material/DragIndicatorOutlined'
import EditOutlined from '@mui/icons-material/EditOutlined'
import KeyboardArrowDownOutlined from '@mui/icons-material/KeyboardArrowDownOutlined'
import KeyboardArrowUpOutlined from '@mui/icons-material/KeyboardArrowUpOutlined'
import { useMemo, useState } from 'react'
import { useTranslation } from 'react-i18next'
import type { MenuPageOption } from '../../api/menusApi'
import { AddIcon, CheckCircleIcon, LinkIcon } from '../icons'
import {
  applyMenuItemDraft,
  buildMenuItemDraft,
  canMoveMenuItem,
  createDefaultMenuItemDraft,
  findMenuItemByClientId,
  findParentPageIdForItem,
  getMenuItemSummary,
  listPageParentMenuOptions,
  moveMenuItem,
  removeMenuItem,
  reorderMenuItems,
  validateMenuItemDraft,
  type MenuEditorItem,
  type MenuItemDraft,
} from '../../lib/menusEditor'
import styles from './MenuBuilder.module.css'

type MenuBuilderProps = {
  title: string
  subtitle: string
  items: MenuEditorItem[]
  pageOptions: MenuPageOption[]
  isDirty: boolean
  isSaving: boolean
  saveError?: string | null
  onChange: (items: MenuEditorItem[]) => void
  onSave: () => void
  onDiscard: () => void
}

type DialogState =
  | { mode: 'create' }
  | {
      mode: 'edit'
      itemId: string
    }

type RemoveCandidate = {
  clientId: string
  label: string
  descendants: number
}

export function MenuBuilder({
  title,
  subtitle,
  items,
  pageOptions,
  isDirty,
  isSaving,
  saveError,
  onChange,
  onSave,
  onDiscard,
}: MenuBuilderProps) {
  const { t } = useTranslation()
  const [dialogState, setDialogState] = useState<DialogState | null>(null)
  const [draft, setDraft] = useState<MenuItemDraft>(createDefaultMenuItemDraft())
  const [draftError, setDraftError] = useState<string | null>(null)
  const [removeCandidate, setRemoveCandidate] = useState<RemoveCandidate | null>(null)
  const [draggedClientId, setDraggedClientId] = useState<string | null>(null)
  const [dragOverClientId, setDragOverClientId] = useState<string | null>(null)

  const editingLocation = useMemo(() => {
    if (!dialogState || dialogState.mode !== 'edit') {
      return null
    }
    return findMenuItemByClientId(items, dialogState.itemId)
  }, [dialogState, items])

  const hasLockedChildren = Boolean(editingLocation?.item.children.length)
  const externalParentOptions = useMemo(
    () => listPageParentMenuOptions(items),
    [items],
  )

  function openCreateDialog() {
    setDraft(createDefaultMenuItemDraft())
    setDraftError(null)
    setDialogState({ mode: 'create' })
  }

  function openEditDialog(item: MenuEditorItem) {
    setDraft(buildMenuItemDraft(item, findParentPageIdForItem(items, item.client_id)))
    setDraftError(null)
    setDialogState({
      mode: 'edit',
      itemId: item.client_id,
    })
  }

  function closeDialog() {
    setDialogState(null)
    setDraftError(null)
  }

  function submitDraft() {
    const error = validateMenuItemDraft(
      draft,
      items,
      pageOptions,
      dialogState?.mode === 'edit' ? dialogState.itemId : undefined,
      hasLockedChildren,
    )
    if (error) {
      setDraftError(error)
      return
    }

    const nextItems = applyMenuItemDraft(
      items,
      draft,
      pageOptions,
      dialogState?.mode === 'edit' ? dialogState.itemId : undefined,
    )
    onChange(nextItems)
    closeDialog()
  }

  function confirmRemove(item: MenuEditorItem) {
    setRemoveCandidate({
      clientId: item.client_id,
      label: item.label,
      descendants: countDescendants(item),
    })
  }

  function handleRemoveConfirmed() {
    if (!removeCandidate) {
      return
    }
    onChange(removeMenuItem(items, removeCandidate.clientId))
    setRemoveCandidate(null)
  }

  function handleDragStart(clientId: string) {
    setDraggedClientId(clientId)
    setDragOverClientId(null)
  }

  function handleDrop(targetClientId: string) {
    if (!draggedClientId || draggedClientId === targetClientId) {
      setDraggedClientId(null)
      setDragOverClientId(null)
      return
    }

    const dragged = findMenuItemByClientId(items, draggedClientId)
    const target = findMenuItemByClientId(items, targetClientId)
    const canReorder =
      dragged &&
      target &&
      dragged.parentPath.join('.') === target.parentPath.join('.')

    if (canReorder) {
      onChange(reorderMenuItems(items, draggedClientId, targetClientId))
    }

    setDraggedClientId(null)
    setDragOverClientId(null)
  }

  function renderItems(branch: MenuEditorItem[], depth: number) {
    return branch.map((item) => {
      const isDragOver =
        dragOverClientId === item.client_id &&
        draggedClientId !== null &&
        draggedClientId !== item.client_id
      const hasChildren = item.children.length > 0
      const itemSummary = getMenuItemSummary(item, pageOptions)

      return (
        <div key={item.client_id} className={styles.treeNode}>
          <div
            className={[
              styles.itemRow,
              isDragOver ? styles.itemRowDragOver : '',
              draggedClientId === item.client_id ? styles.itemRowDragging : '',
            ].join(' ')}
            style={{ marginLeft: `${depth * 24}px` }}
            draggable
            onDragStart={() => handleDragStart(item.client_id)}
            onDragEnd={() => {
              setDraggedClientId(null)
              setDragOverClientId(null)
            }}
            onDragOver={(event) => {
              event.preventDefault()
              const dragged = draggedClientId
                ? findMenuItemByClientId(items, draggedClientId)
                : null
              const target = findMenuItemByClientId(items, item.client_id)
              if (
                dragged &&
                target &&
                dragged.parentPath.join('.') === target.parentPath.join('.')
              ) {
                setDragOverClientId(item.client_id)
              }
            }}
            onDrop={(event) => {
              event.preventDefault()
              handleDrop(item.client_id)
            }}
          >
            <span className={styles.dragHandle} aria-hidden="true">
              <DragIndicatorOutlined fontSize="small" />
            </span>

            <div className={styles.itemBody}>
              <div className={styles.itemHeader}>
                <span className={styles.itemLabel}>{item.label}</span>
                <span
                  className={[
                    styles.typeBadge,
                    item.navigation_type === 'pages' ? styles.pageBadge : styles.externalBadge,
                  ].join(' ')}
                >
                  {item.navigation_type === 'pages'
                    ? t('menus.builder.types.page')
                    : t('menus.builder.types.external')}
                </span>

                {hasChildren && (
                  <span className={styles.childCount}>
                    {t('menus.builder.childCount', { count: item.children.length })}
                  </span>
                )}
              </div>

              <p className={styles.itemMeta} title={itemSummary}>
                {itemSummary}
              </p>
            </div>

            <div className={styles.itemActions}>
              <button
                type="button"
                className={styles.iconButton}
                aria-label={t('menus.actions.moveUp')}
                disabled={!canMoveMenuItem(items, item.client_id, 'up')}
                onClick={() => onChange(moveMenuItem(items, item.client_id, 'up'))}
              >
                <KeyboardArrowUpOutlined fontSize="small" />
              </button>
              <button
                type="button"
                className={styles.iconButton}
                aria-label={t('menus.actions.moveDown')}
                disabled={!canMoveMenuItem(items, item.client_id, 'down')}
                onClick={() => onChange(moveMenuItem(items, item.client_id, 'down'))}
              >
                <KeyboardArrowDownOutlined fontSize="small" />
              </button>
              <button
                type="button"
                className={styles.iconButton}
                aria-label={t('menus.actions.edit')}
                onClick={() => openEditDialog(item)}
              >
                <EditOutlined fontSize="small" />
              </button>
              <button
                type="button"
                className={styles.iconButtonDanger}
                aria-label={t('menus.actions.remove')}
                onClick={() => confirmRemove(item)}
              >
                <DeleteOutlineOutlined fontSize="small" />
              </button>
            </div>
          </div>

          {hasChildren && <div className={styles.childGroup}>{renderItems(item.children, depth + 1)}</div>}
        </div>
      )
    })
  }

  return (
    <section className={styles.builder}>
      <header className={styles.header}>
        <div>
          <h2 className={styles.title}>{title}</h2>
          <p className={styles.subtitle}>{subtitle}</p>
        </div>

        <button type="button" className={styles.primaryButton} onClick={openCreateDialog}>
          <AddIcon size={16} />
          {t('menus.actions.add')}
        </button>
      </header>

      {saveError && <div className={styles.errorPanel}>{saveError}</div>}

      <div className={styles.board}>
        {items.length > 0 ? (
          renderItems(items, 0)
        ) : (
          <div className={styles.emptyState}>
            <div className={styles.emptyStateIcon}>
              <LinkIcon size={18} />
            </div>
            <p className={styles.emptyStateTitle}>{t('menus.empty.title')}</p>
            <p className={styles.emptyStateText}>{t('menus.empty.text')}</p>
          </div>
        )}
      </div>

      <footer className={styles.footer}>
        <div className={styles.stageStatus}>
          <CheckCircleIcon size={16} />
          {isDirty ? t('menus.status.staged') : t('menus.status.saved')}
        </div>

        <div className={styles.footerActions}>
          <button
            type="button"
            className={styles.secondaryButton}
            disabled={!isDirty || isSaving}
            onClick={onDiscard}
          >
            {t('menus.actions.discard')}
          </button>
          <button
            type="button"
            className={styles.primaryButton}
            disabled={isSaving}
            onClick={onSave}
          >
            {isSaving ? t('menus.common.saving') : t('menus.actions.save')}
          </button>
        </div>
      </footer>

      <Dialog open={Boolean(dialogState)} onClose={closeDialog} fullWidth maxWidth="sm">
        <DialogTitle>
          {dialogState?.mode === 'edit'
            ? t('menus.dialog.editTitle')
            : t('menus.dialog.createTitle')}
        </DialogTitle>
        <DialogContent>
          <DialogContentText className={styles.dialogText}>
            {t('menus.dialog.description')}
          </DialogContentText>

          <div className={styles.dialogFields}>
            <label className={styles.field}>
              <span>{t('menus.dialog.fields.menuLabel')}</span>
              <input
                type="text"
                value={draft.label}
                onChange={(event) => {
                  setDraft((current) => ({
                    ...current,
                    label: event.target.value,
                  }))
                  setDraftError(null)
                }}
                placeholder={t('menus.dialog.fields.menuLabelPlaceholder')}
              />
            </label>

            <div className={styles.radioField}>
              <span>{t('menus.dialog.fields.linkType')}</span>
              <label className={styles.choiceCard}>
                <input
                  type="radio"
                  name="navigation_type"
                  checked={draft.navigation_type === 'pages'}
                  disabled={hasLockedChildren}
                  onChange={() => {
                    setDraft((current) => ({
                      ...current,
                      navigation_type: 'pages',
                      external_url: '',
                      parent_id: '',
                      open_in_new_tab: false,
                    }))
                    setDraftError(null)
                  }}
                />
                <div>
                  <strong>{t('menus.dialog.fields.linkTypePage')}</strong>
                  <p>{t('menus.dialog.fields.linkTypePageHint')}</p>
                </div>
              </label>
              <label className={styles.choiceCard}>
                <input
                  type="radio"
                  name="navigation_type"
                  checked={draft.navigation_type === 'external_link'}
                  disabled={hasLockedChildren}
                  onChange={() => {
                    setDraft((current) => ({
                      ...current,
                      navigation_type: 'external_link',
                      page_id: '',
                    }))
                    setDraftError(null)
                  }}
                />
                <div>
                  <strong>{t('menus.dialog.fields.linkTypeExternal')}</strong>
                  <p>{t('menus.dialog.fields.linkTypeExternalHint')}</p>
                </div>
              </label>
            </div>

            {draft.navigation_type === 'pages' ? (
              <label className={styles.field}>
                <span>{t('menus.dialog.fields.targetPage')}</span>
                <select
                  value={draft.page_id}
                  disabled={hasLockedChildren}
                  onChange={(event) => {
                    setDraft((current) => ({
                      ...current,
                      page_id: event.target.value,
                    }))
                    setDraftError(null)
                  }}
                >
                  <option value="">{t('menus.dialog.fields.selectPage')}</option>
                  {pageOptions.map((page) => (
                    <option key={page.id} value={page.id}>
                      {`${page.page_title} (${page.url_slug})`}
                    </option>
                  ))}
                </select>
                <p className={styles.fieldHint}>{t('menus.dialog.fields.targetPageHint')}</p>
              </label>
            ) : (
              <>
                <label className={styles.field}>
                  <span>{t('menus.dialog.fields.externalUrl')}</span>
                  <input
                    type="url"
                    value={draft.external_url}
                    onChange={(event) => {
                      setDraft((current) => ({
                        ...current,
                        external_url: event.target.value,
                      }))
                      setDraftError(null)
                    }}
                    placeholder={t('menus.dialog.fields.externalUrlPlaceholder')}
                  />
                </label>

                <label className={styles.field}>
                  <span>{t('menus.dialog.fields.parentPage')}</span>
                  <select
                    value={draft.parent_id}
                    onChange={(event) => {
                      setDraft((current) => ({
                        ...current,
                        parent_id: event.target.value,
                      }))
                      setDraftError(null)
                    }}
                  >
                    <option value="">{t('menus.dialog.fields.noParentPage')}</option>
                    {externalParentOptions.map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                  <p className={styles.fieldHint}>{t('menus.dialog.fields.parentPageHint')}</p>
                </label>

                <label className={styles.toggleField}>
                  <span>{t('menus.dialog.fields.openInNewTab')}</span>
                  <input
                    type="checkbox"
                    checked={draft.open_in_new_tab}
                    onChange={(event) => {
                      setDraft((current) => ({
                        ...current,
                        open_in_new_tab: event.target.checked,
                      }))
                    }}
                  />
                </label>
              </>
            )}

            {hasLockedChildren && (
              <p className={styles.lockedHint}>{t('menus.dialog.fields.lockedHint')}</p>
            )}

            {draftError && <p className={styles.dialogError}>{draftError}</p>}
          </div>
        </DialogContent>
        <DialogActions>
          <button type="button" className={styles.secondaryButton} onClick={closeDialog}>
            {t('menus.actions.cancel')}
          </button>
          <button type="button" className={styles.primaryButton} onClick={submitDraft}>
            {dialogState?.mode === 'edit' ? t('menus.actions.update') : t('menus.actions.addToMenu')}
          </button>
        </DialogActions>
      </Dialog>

      <Dialog open={Boolean(removeCandidate)} onClose={() => setRemoveCandidate(null)} maxWidth="xs" fullWidth>
        <DialogTitle>{t('menus.removeDialog.title')}</DialogTitle>
        <DialogContent>
          <DialogContentText>
            {removeCandidate?.descendants
              ? t('menus.removeDialog.withChildren', {
                  label: removeCandidate.label,
                  count: removeCandidate.descendants,
                })
              : t('menus.removeDialog.single', {
                  label: removeCandidate?.label ?? '',
                })}
          </DialogContentText>
        </DialogContent>
        <DialogActions>
          <button
            type="button"
            className={styles.secondaryButton}
            onClick={() => setRemoveCandidate(null)}
          >
            {t('menus.actions.cancel')}
          </button>
          <button type="button" className={styles.dangerButton} onClick={handleRemoveConfirmed}>
            {t('menus.actions.remove')}
          </button>
        </DialogActions>
      </Dialog>
    </section>
  )
}

function countDescendants(item: MenuEditorItem): number {
  return item.children.reduce((count, child) => count + 1 + countDescendants(child), 0)
}
