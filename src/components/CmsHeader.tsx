import csaaLogo from '../assets/csaa_logo.png'
import nordikLogo from '../assets/nordik_logo.png'
import { BellIcon, HelpIcon, MenuIcon } from './icons'
import styles from './CmsHeader.module.css'

type CmsHeaderProps = {
  onMenuClick?: () => void
}

export function CmsHeader({ onMenuClick }: CmsHeaderProps) {
  return (
    <header className={styles.header}>
      <div className={styles.brand}>
        <button
          type="button"
          aria-label="Open navigation menu"
          className={styles.menuButton}
          onClick={onMenuClick}
        >
          <MenuIcon />
        </button>
        <img
          src={csaaLogo}
          alt="Children of Shingwauk Alumni Association"
          className={styles.logo}
        />
        <img
          src={nordikLogo}
          alt="Nordik Institute"
          className={styles.logo}
        />
      </div>

      <div className={styles.actions}>
        <button type="button" aria-label="Notifications" className={styles.iconButton}>
          <BellIcon />
        </button>
        <button type="button" aria-label="Help" className={styles.iconButton}>
          <HelpIcon />
        </button>
      </div>
    </header>
  )
}
