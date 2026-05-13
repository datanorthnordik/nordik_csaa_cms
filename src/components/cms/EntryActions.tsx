import { useState, type ReactNode } from 'react'
import { useTranslation } from 'react-i18next'
import { ConfirmDialog } from './ConfirmDialog'
import type { PublishStatus } from './StatusBadge'
import styles from './EntryActions.module.css'

export type EntryType = 'event' | 'press'

export type EntryActionsProps = {
  status: PublishStatus
  entryType?: EntryType
  publishLabel?: string
  saveDraftLabel?: string
  deleteLabel?: string
  deleteConfirmTitle?: string
  deleteConfirmBody?: ReactNode
  deleteConfirmLabel?: string
  onPublish: () => void
  onSaveDraft: () => void
  onDelete?: () => void
  isSubmitting?: boolean
  isDeleting?: boolean
  className?: string
}

export function EntryActions({
  status,
  entryType,
  publishLabel,
  saveDraftLabel,
  deleteLabel,
  deleteConfirmTitle,
  deleteConfirmBody,
  deleteConfirmLabel,
  onPublish,
  onSaveDraft,
  onDelete,
  isSubmitting = false,
  isDeleting = false,
  className,
}: EntryActionsProps) {
  const { t } = useTranslation()
  const [confirmOpen, setConfirmOpen] = useState(false)

  const resolvedPublish =
    publishLabel ??
    (status === 'published' ? t('publishing.saveChanges') : t('publishing.publish'))
  const resolvedSaveDraft = saveDraftLabel ?? t('publishing.saveDraft')
  const defaultDeleteLabel =
    entryType === 'event'
      ? t('events.list.delete')
      : entryType === 'press'
        ? t('press.editor.actions.delete')
        : t('publishing.delete')
  const resolvedDelete = deleteLabel ?? defaultDeleteLabel
  const resolvedConfirmTitle =
    deleteConfirmTitle ?? t('publishing.deleteConfirm.defaultTitle')
  const resolvedConfirmLabel = deleteConfirmLabel ?? resolvedDelete

  function handleRequestDelete() {
    if (!onDelete || isDeleting) return
    setConfirmOpen(true)
  }

  function handleConfirmDelete() {
    if (!onDelete) return
    onDelete()
  }

  function handleCloseConfirm() {
    if (isDeleting) return
    setConfirmOpen(false)
  }

  const wrapClasses = [styles.wrap, className].filter(Boolean).join(' ')

  return (
    <>
      <div className={wrapClasses}>
        <button
          type="button"
          className={`${styles.button} ${styles.primary}`}
          disabled={isSubmitting}
          onClick={onPublish}
        >
          {isSubmitting ? t('publishing.working') : resolvedPublish}
        </button>
        <button
          type="button"
          className={`${styles.button} ${styles.secondary}`}
          disabled={isSubmitting}
          onClick={onSaveDraft}
        >
          {resolvedSaveDraft}
        </button>

        {onDelete && (
          <button
            type="button"
            className={styles.danger}
            disabled={isDeleting || isSubmitting}
            onClick={handleRequestDelete}
          >
            <TrashIcon />
            {resolvedDelete}
          </button>
        )}
      </div>

      {onDelete && (
        <ConfirmDialog
          open={confirmOpen}
          title={resolvedConfirmTitle}
          body={deleteConfirmBody}
          confirmLabel={resolvedConfirmLabel}
          destructive
          busy={isDeleting}
          onConfirm={handleConfirmDelete}
          onClose={handleCloseConfirm}
        />
      )}
    </>
  )
}

function TrashIcon() {
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden="true">
      <path
        d="M2.8 4.4h10.4M6 4.4V3.2c0-.66.54-1.2 1.2-1.2h1.6c.66 0 1.2.54 1.2 1.2v1.2M4.4 4.4l.6 8.2c.04.66.6 1.2 1.26 1.2h3.48c.66 0 1.22-.54 1.26-1.2l.6-8.2"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}
