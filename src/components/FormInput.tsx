import { useId, useState } from 'react'
import type { InputHTMLAttributes } from 'react'
import { useTranslation } from 'react-i18next'
import styles from './FormInput.module.css'

type FormInputProps = Omit<InputHTMLAttributes<HTMLInputElement>, 'id'> & {
  label?: string
  variant?: 'outlined' | 'filled'
}

export function FormInput({
  label,
  variant = 'outlined',
  type = 'text',
  className,
  ...rest
}: FormInputProps) {
  const id = useId()
  const { t } = useTranslation()
  const isPassword = type === 'password'
  const [revealed, setRevealed] = useState(false)
  const effectiveType = isPassword && revealed ? 'text' : type

  return (
    <div className={[styles.field, className].filter(Boolean).join(' ')}>
      {label && (
        <label htmlFor={id} className={styles.label}>
          {label}
        </label>
      )}
      <div
        className={[
          styles.inputWrap,
          variant === 'filled' ? styles.filled : styles.outlined,
        ].join(' ')}
      >
        <input
          id={id}
          type={effectiveType}
          className={[styles.input, isPassword ? styles.hasToggle : '']
            .filter(Boolean)
            .join(' ')}
          {...rest}
        />
        {isPassword && (
          <button
            type="button"
            className={styles.toggle}
            onClick={() => setRevealed((value) => !value)}
            aria-label={
              revealed
                ? t('auth.formInput.hidePassword')
                : t('auth.formInput.showPassword')
            }
            disabled={rest.disabled}
            tabIndex={-1}
          >
            <EyeIcon open={revealed} />
          </button>
        )}
      </div>
    </div>
  )
}

function EyeIcon({ open }: { open: boolean }) {
  return (
    <svg
      width="18"
      height="18"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z" />
      <circle cx="12" cy="12" r="3" />
      {!open && <line x1="3" y1="3" x2="21" y2="21" />}
    </svg>
  )
}
