import type { ReactNode } from 'react'
import { useNavigate } from 'react-router-dom'
import { CmsLayout } from './CmsLayout'
import type { DrawerNavItem } from './CmsDrawer'
import {
  cmsNavItems,
  cmsFooterItems,
  type CmsNavItemConfig,
} from '../config/cmsNav'

type CmsAppShellProps = {
  children?: ReactNode
  activeKey?: string
  userName?: string
  userRole?: string
  onSubmitTicket?: () => void
  onLogout?: () => void
}

export function CmsAppShell({
  children,
  activeKey,
  userName = 'Christopher Masha',
  userRole = 'System Administrator',
  onSubmitTicket,
  onLogout,
}: CmsAppShellProps) {
  const navigate = useNavigate()

  function toDrawerItem(item: CmsNavItemConfig): DrawerNavItem {
    let onClick: (() => void) | undefined
    if (item.key === 'logout' && onLogout) {
      onClick = onLogout
    } else if (item.path) {
      onClick = () => navigate(item.path!)
    }

    return {
      key: item.key,
      label: item.label,
      icon: item.icon,
      active: item.key === activeKey,
      onClick,
    }
  }

  return (
    <CmsLayout
      userName={userName}
      userRole={userRole}
      navItems={cmsNavItems.map(toDrawerItem)}
      footerItems={cmsFooterItems.map(toDrawerItem)}
      onSubmitTicket={onSubmitTicket}
    >
      {children}
    </CmsLayout>
  )
}
