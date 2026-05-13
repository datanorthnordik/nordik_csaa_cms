import type { FormEvent } from 'react'
import { useTranslation } from 'react-i18next'
import styles from './SearchFilterBar.module.css'

export type FilterOption = {
  value: string
  label: string
}

export type FilterFieldConfig =
  | {
      type: 'select'
      key: string
      label: string
      value: string
      options: FilterOption[]
      onChange: (value: string) => void
      disabled?: boolean
    }
  | {
      type: 'date'
      key: string
      label: string
      value: string
      onChange: (value: string) => void
      disabled?: boolean
    }
  | {
      type: 'multi-pills'
      key: string
      label: string
      values: string[]
      options: FilterOption[]
      onToggle: (value: string) => void
    }

export type SearchFilterBarProps = {
  searchValue: string
  onSearchChange: (value: string) => void
  searchPlaceholder?: string
  searchLabel?: string
  fields?: FilterFieldConfig[]
  onApply?: () => void
  onReset?: () => void
  applyLabel?: string
  resetLabel?: string
  compact?: boolean
  className?: string
}

export function SearchFilterBar({
  searchValue,
  onSearchChange,
  searchPlaceholder,
  searchLabel,
  fields = [],
  onApply,
  onReset,
  applyLabel,
  resetLabel,
  compact = false,
  className,
}: SearchFilterBarProps) {
  const { t } = useTranslation()
  const resolvedSearchLabel = searchLabel ?? t('searchFilterBar.search')
  const resolvedSearchPlaceholder =
    searchPlaceholder ?? t('searchFilterBar.searchPlaceholder')
  const resolvedApplyLabel = applyLabel ?? t('searchFilterBar.apply')
  const resolvedResetLabel = resetLabel ?? t('searchFilterBar.reset')

  function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    onApply?.()
  }

  const classNames = [styles.panel, compact ? styles.compact : '', className]
    .filter(Boolean)
    .join(' ')

  return (
    <form className={classNames} onSubmit={handleSubmit} role="search">
      <div className={styles.row}>
        <label className={styles.searchField}>
          <span style={{ display: 'none' }}>{resolvedSearchLabel}</span>
          <span className={styles.searchInputWrap}>
            <span className={styles.searchIcon} aria-hidden="true">
              <SearchIcon />
            </span>
            <input
              type="search"
              value={searchValue}
              placeholder={resolvedSearchPlaceholder}
              aria-label={resolvedSearchLabel}
              onChange={(event) => onSearchChange(event.target.value)}
              className={styles.searchInput}
            />
          </span>
        </label>

        {fields.map((field) => {
          if (field.type === 'select') {
            return (
              <label key={field.key} className={styles.field}>
                <span>{field.label}</span>
                <select
                  value={field.value}
                  disabled={field.disabled}
                  onChange={(event) => field.onChange(event.target.value)}
                >
                  {field.options.map((option) => (
                    <option key={option.value} value={option.value}>
                      {option.label}
                    </option>
                  ))}
                </select>
              </label>
            )
          }

          if (field.type === 'date') {
            return (
              <label key={field.key} className={styles.field}>
                <span>{field.label}</span>
                <input
                  type="date"
                  value={field.value}
                  disabled={field.disabled}
                  onChange={(event) => field.onChange(event.target.value)}
                />
              </label>
            )
          }

          return (
            <div key={field.key} className={styles.pillsField}>
              <span>{field.label}</span>
              <div className={styles.pillsRow}>
                {field.options.map((option) => (
                  <label key={option.value} className={styles.pill}>
                    <input
                      type="checkbox"
                      checked={field.values.includes(option.value)}
                      onChange={() => field.onToggle(option.value)}
                    />
                    <span>{option.label}</span>
                  </label>
                ))}
              </div>
            </div>
          )
        })}
      </div>

      {(onApply || onReset) && (
        <div className={styles.actions}>
          {onReset && (
            <button type="button" className={styles.secondary} onClick={onReset}>
              {resolvedResetLabel}
            </button>
          )}
          {onApply && (
            <button type="submit" className={styles.primary}>
              {resolvedApplyLabel}
            </button>
          )}
        </div>
      )}
    </form>
  )
}

function SearchIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <circle cx="7" cy="7" r="5" stroke="currentColor" strokeWidth="1.6" />
      <path
        d="M11 11l3 3"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
      />
    </svg>
  )
}
