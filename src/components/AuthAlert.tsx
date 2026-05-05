import styles from './AuthAlert.module.css'

export type AlertState = {
  type: 'error' | 'success'
  message: string
}

type AuthAlertProps = {
  alert: AlertState | null
}

export function AuthAlert({ alert }: AuthAlertProps) {
  if (!alert) return null
  const className = [
    styles.alert,
    alert.type === 'success' ? styles.success : styles.error,
  ].join(' ')
  return (
    <div className={className} role="alert">
      {alert.message}
    </div>
  )
}
