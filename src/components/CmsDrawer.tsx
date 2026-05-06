import type { ReactNode } from 'react'
import styles from './CmsDrawer.module.css'

export type DrawerNavItem = {
  key: string
  label: string
  icon: ReactNode
  href?: string
  active?: boolean
  onClick?: () => void
}

type CmsDrawerProps = {
  userName?: string
  userRole?: string
  navItems: DrawerNavItem[]
  footerItems?: DrawerNavItem[]
  open?: boolean
  onClose?: () => void
  onSubmitTicket?: () => void
}

export function CmsDrawer({
  userName,
  userRole,
  navItems,
  footerItems = [],
  open = false,
  onClose,
  onSubmitTicket,
}: CmsDrawerProps) {
  function renderNavItem(item: DrawerNavItem) {
    const className = `${styles.navItem} ${item.active ? styles.navItemActive : ''}`

    if (item.href) {
      return (
        <a
          key={item.key}
          href={item.href}
          className={className}
          onClick={item.onClick}
        >
          <span className={styles.navIcon}>{item.icon}</span>
          <span>{item.label}</span>
        </a>
      )
    }

    return (
      <button
        key={item.key}
        type="button"
        className={className}
        onClick={item.onClick}
      >
        <span className={styles.navIcon}>{item.icon}</span>
        <span>{item.label}</span>
      </button>
    )
  }

  return (
    <>
      <div
        className={`${styles.backdrop} ${open ? styles.backdropOpen : ''}`}
        onClick={onClose}
        aria-hidden="true"
      />
      <aside
        className={`${styles.drawer} ${open ? styles.drawerOpen : ''}`}
        aria-label="Primary navigation"
      >
        {(userName || userRole) && (
          <div className={styles.profile}>
            {userName && <span className={styles.profileName}>{userName}</span>}
            {userRole && <span className={styles.profileRole}>{userRole}</span>}
          </div>
        )}

        <nav className={styles.nav}>
          {navItems.map(renderNavItem)}
        </nav>

        <div className={styles.footer}>
          {onSubmitTicket && (
            <button
              type="button"
              className={styles.submitTicket}
              onClick={onSubmitTicket}
            >
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                <path d="M21 7v4a2 2 0 0 0-2 2 2 2 0 0 0 2 2v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4a2 2 0 0 0 2-2 2 2 0 0 0-2-2V7a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" />
                <line x1="13" y1="5" x2="13" y2="7" />
                <line x1="13" y1="11" x2="13" y2="13" />
                <line x1="13" y1="17" x2="13" y2="19" />
              </svg>
              Submit Ticket
            </button>
          )}
          {footerItems.map(renderNavItem)}
        </div>
      </aside>
    </>
  )
}
