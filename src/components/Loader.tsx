import styles from './Loader.module.css'

type LoaderProps = {
  label?: string
  fullscreen?: boolean
}

export function Loader({ label, fullscreen = false }: LoaderProps) {
  return (
    <div
      className={[styles.loader, fullscreen ? styles.fullscreen : '']
        .filter(Boolean)
        .join(' ')}
      role="status"
      aria-live="polite"
    >
      <span className={styles.spinner} aria-hidden="true" />
      <span className={styles.label}>{label ?? 'Loading...'}</span>
    </div>
  )
}
