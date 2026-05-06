import csaaLogo from '../assets/csaa_logo.png'
import nordikLogo from '../assets/nordik_logo.png'
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
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <line x1="3" y1="6" x2="21" y2="6" />
            <line x1="3" y1="12" x2="21" y2="12" />
            <line x1="3" y1="18" x2="21" y2="18" />
          </svg>
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
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <path d="M18 16v-5a6 6 0 0 0-12 0v5l-2 2v1h16v-1l-2-2z" />
            <path d="M10 21a2 2 0 0 0 4 0" />
          </svg>
        </button>
        <button type="button" aria-label="Help" className={styles.iconButton}>
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <circle cx="12" cy="12" r="10" />
            <path d="M9.5 9a2.5 2.5 0 0 1 5 0c0 1.5-2.5 2-2.5 4" />
            <line x1="12" y1="17" x2="12" y2="17.01" />
          </svg>
        </button>
      </div>
    </header>
  )
}
