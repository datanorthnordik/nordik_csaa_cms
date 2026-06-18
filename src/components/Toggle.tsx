import styles from './Toggle.module.css'

interface ToggleProps {
  checked: boolean
  onChange: (checked: boolean) => void
  disabled?: boolean
  id?: string
}

export function Toggle({ checked, onChange, disabled = false, id }: ToggleProps) {
  return (
    <button
      id={id}
      type="button"
      role="switch"
      aria-checked={checked}
      className={`${styles.toggle} ${checked ? styles.toggleActive : ''} ${disabled ? styles.toggleDisabled : ''}`}
      onClick={() => !disabled && onChange(!checked)}
      disabled={disabled}
    >
      <span className={styles.toggleThumb} />
    </button>
  )
}
