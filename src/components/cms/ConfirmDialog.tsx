import type { ReactNode } from 'react'
import Dialog from '@mui/material/Dialog'
import DialogActions from '@mui/material/DialogActions'
import DialogContent from '@mui/material/DialogContent'
import DialogContentText from '@mui/material/DialogContentText'
import DialogTitle from '@mui/material/DialogTitle'
import { useTranslation } from 'react-i18next'
import styles from './ConfirmDialog.module.css'

type ConfirmDialogProps = {
  open: boolean
  title: string
  body?: ReactNode
  confirmLabel?: string
  cancelLabel?: string
  destructive?: boolean
  busy?: boolean
  onConfirm: () => void
  onClose: () => void
}

export function ConfirmDialog({
  open,
  title,
  body,
  confirmLabel,
  cancelLabel,
  destructive = false,
  busy = false,
  onConfirm,
  onClose,
}: ConfirmDialogProps) {
  const { t } = useTranslation()
  const resolvedConfirm = confirmLabel ?? t('confirmDialog.confirm')
  const resolvedCancel = cancelLabel ?? t('confirmDialog.cancel')
  const busyLabel = t('confirmDialog.working')

  const confirmClasses = [
    styles.button,
    destructive ? styles.destructive : styles.confirm,
  ].join(' ')

  function handleClose() {
    if (busy) {
      return
    }
    onClose()
  }

  return (
    <Dialog
      open={open}
      onClose={handleClose}
      slotProps={{ paper: { className: styles.paper } }}
    >
      <DialogTitle className={styles.title}>{title}</DialogTitle>
      {body && (
        <DialogContent className={styles.content}>
          <DialogContentText className={styles.body}>{body}</DialogContentText>
        </DialogContent>
      )}
      <DialogActions className={styles.actions}>
        <button
          type="button"
          className={`${styles.button} ${styles.cancel}`}
          disabled={busy}
          onClick={handleClose}
        >
          {resolvedCancel}
        </button>
        <button
          type="button"
          className={confirmClasses}
          disabled={busy}
          onClick={onConfirm}
        >
          {busy ? busyLabel : resolvedConfirm}
        </button>
      </DialogActions>
    </Dialog>
  )
}
