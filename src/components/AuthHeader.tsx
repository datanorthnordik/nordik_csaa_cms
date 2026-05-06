import { useTranslation } from 'react-i18next'
import csaaLogo from '../assets/csaa_logo.png'
import nordikLogo from '../assets/nordik_logo.png'
import { LanguageSwitcher } from './LanguageSwitcher'
import styles from './AuthHeader.module.css'

export function AuthHeader() {
  const { t } = useTranslation()

  return (
    <div className={styles.header}>
      <div className={styles.brand}>
        <img
          src={csaaLogo}
          alt={t('auth.header.csaaAlt')}
          className={styles.logo}
        />
        <img
          src={nordikLogo}
          alt={t('auth.header.nordikAlt')}
          className={styles.logo}
        />
      </div>
      <div className={styles.controls}>
        <LanguageSwitcher />
      </div>
    </div>
  )
}
