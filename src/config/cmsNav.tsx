import type { ReactNode } from 'react'
import {
  DashboardIcon,
  MenusIcon,
  PagesIcon,
  EventsIcon,
  NewslettersIcon,
  PressIcon,
  MemorialIcon,
  ResourcesIcon,
  MediaIcon,
  SettingsIcon,
  LogoutIcon,
} from '../components/icons'

export type CmsNavItemConfig = {
  key: string
  icon: ReactNode
  path?: string
}

export const cmsNavItems: CmsNavItemConfig[] = [
  { key: 'dashboard', icon: <DashboardIcon />, path: '/dashboard' },
  { key: 'menus', icon: <MenusIcon />, path: '/menus' },
  { key: 'pages', icon: <PagesIcon />, path: '/pages' },
  { key: 'events', icon: <EventsIcon />, path: '/events' },
  { key: 'newsletters', icon: <NewslettersIcon />, path: '' },
  { key: 'press', icon: <PressIcon />, path: '/press' },
  { key: 'memorial', icon: <MemorialIcon />, path: '' },
  { key: 'resources', icon: <ResourcesIcon />, path: '' },
  { key: 'media', icon: <MediaIcon />, path: '/media-library' },
]

export const cmsFooterItems: CmsNavItemConfig[] = [
  { key: 'settings', icon: <SettingsIcon />, path: '' },
  { key: 'logout', icon: <LogoutIcon /> },
]
