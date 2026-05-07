import type { ReactNode } from 'react'
import { useNavigate } from 'react-router-dom'
import { useTranslation } from 'react-i18next'
import { CmsLayout } from './CmsLayout'
import type { DrawerNavItem } from './CmsDrawer'
import {
  cmsNavItems,
  cmsFooterItems,
  type CmsNavItemConfig,
} from '../config/cmsNav'
import { selectAuthSession } from '../store/authSlice'
import { useAppSelector } from '../store/hooks'

type CmsAppShellProps = {
  children?: ReactNode
  activeKey?: string
  onSubmitTicket?: () => void
  onLogout?: () => void
}

export function CmsAppShell({
  children,
  activeKey,
  onSubmitTicket,
  onLogout,
}: CmsAppShellProps) {
  const navigate = useNavigate()
  const { t } = useTranslation()
  const session = useAppSelector(selectAuthSession)
  const userName = session
    ? `${session.firstname} ${session.lastname}`.trim() || undefined
    : undefined
  const userRole = session?.role

  function toDrawerItem(item: CmsNavItemConfig): DrawerNavItem {
    let onClick: (() => void) | undefined
    if (item.key === 'logout' && onLogout) {
      onClick = onLogout
    } else if (item.path) {
      onClick = () => navigate(item.path!)
    }

    return {
      key: item.key,
      label: t(`nav.${item.key}`),
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
