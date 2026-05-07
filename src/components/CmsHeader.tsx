import csaaLogo from '../assets/csaa_logo.png'
import nordikLogo from '../assets/nordik_logo.png'
import { useTranslation } from 'react-i18next'
import { BellIcon, HelpIcon, MenuIcon } from './icons'
import { LanguageSwitcher } from './LanguageSwitcher'
import styles from './CmsHeader.module.css'

type CmsHeaderProps = {
  onMenuClick?: () => void
}

export function CmsHeader({ onMenuClick }: CmsHeaderProps) {
  const { t } = useTranslation()

  return (
    <header className={styles.header}>
      <div className={styles.brand}>
        <button
          type="button"
          aria-label={t('cmsHeader.menu')}
          className={styles.menuButton}
          onClick={onMenuClick}
        >
          <MenuIcon />
        </button>
        <img
          src={csaaLogo}
          alt={t('cmsHeader.csaaAlt')}
          className={styles.logo}
        />
        <img
          src={nordikLogo}
          alt={t('cmsHeader.nordikAlt')}
          className={styles.logo}
        />
      </div>

      <div className={styles.actions}>
        <LanguageSwitcher />
        <button
          type="button"
          aria-label={t('cmsHeader.notifications')}
          className={styles.iconButton}
        >
          <BellIcon />
        </button>
        <button
          type="button"
          aria-label={t('cmsHeader.help')}
          className={styles.iconButton}
        >
          <HelpIcon />
        </button>
      </div>
    </header>
  )
}
