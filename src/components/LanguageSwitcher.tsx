import { useTranslation } from 'react-i18next'
import {
  defaultLanguage,
  supportedLanguages,
  type SupportedLanguageCode,
} from '../i18n/config'
import styles from './LanguageSwitcher.module.css'

export function LanguageSwitcher() {
  const { i18n, t } = useTranslation()
  const activeLanguage =
    supportedLanguages.find(({ code }) =>
      i18n.resolvedLanguage?.startsWith(code),
    )?.code ?? defaultLanguage

  const handleLanguageChange = (language: SupportedLanguageCode) => {
    void i18n.changeLanguage(language)
  }

  return (
    <div
      className={styles.switcher}
      role="group"
      aria-label={t('languageSwitcher.label')}
    >
      {supportedLanguages.map(({ code, label }) => (
        <button
          key={code}
          type="button"
          className={styles.option}
          aria-pressed={activeLanguage === code}
          onClick={() => handleLanguageChange(code)}
        >
          {label}
        </button>
      ))}
    </div>
  )
}
