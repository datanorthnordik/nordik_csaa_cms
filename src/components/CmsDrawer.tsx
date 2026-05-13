import type { ReactNode } from 'react'
import { useTranslation } from 'react-i18next'
import { TicketIcon } from './icons'
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
  const { t } = useTranslation()
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

        <div className={styles.navWrapper}>
          <nav className={styles.nav}>
            {navItems.map(renderNavItem)}
          </nav>
        </div>

        <div className={styles.footer}>
          {onSubmitTicket && (
            <button
              type="button"
              className={styles.submitTicket}
              onClick={onSubmitTicket}
            >
              <TicketIcon />
              {t('nav.submitTicket')}
            </button>
          )}
          {footerItems.map(renderNavItem)}
        </div>
      </aside>
    </>
  )
}
