import { useState, type ReactNode } from 'react'
import { CmsHeader } from './CmsHeader'
import { CmsDrawer, type DrawerNavItem } from './CmsDrawer'
import styles from './CmsLayout.module.css'

type CmsLayoutProps = {
  children: ReactNode
  userName?: string
  userRole?: string
  navItems: DrawerNavItem[]
  footerItems?: DrawerNavItem[]
  onSubmitTicket?: () => void
}

export function CmsLayout({
  children,
  userName,
  userRole,
  navItems,
  footerItems,
  onSubmitTicket,
}: CmsLayoutProps) {
  const [drawerOpen, setDrawerOpen] = useState(false)

  return (
    <div className={styles.layout}>
      <CmsHeader onMenuClick={() => setDrawerOpen(true)} />
      <div className={styles.body}>
        <CmsDrawer
          userName={userName}
          userRole={userRole}
          navItems={navItems}
          footerItems={footerItems}
          open={drawerOpen}
          onClose={() => setDrawerOpen(false)}
          onSubmitTicket={onSubmitTicket}
        />
        <main className={styles.main}>{children}</main>
      </div>
    </div>
  )
}
