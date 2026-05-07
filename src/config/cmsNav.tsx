import type { ReactNode } from 'react'
import {
  DashboardIcon,
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
  { key: 'dashboard', icon: <DashboardIcon />, path: '' },
  { key: 'pages', icon: <PagesIcon />, path: '' },
  { key: 'events', icon: <EventsIcon />, path: '' },
  { key: 'newsletters', icon: <NewslettersIcon />, path: '' },
  { key: 'press', icon: <PressIcon />, path: '' },
  { key: 'memorial', icon: <MemorialIcon />, path: '' },
  { key: 'resources', icon: <ResourcesIcon />, path: '' },
  { key: 'media', icon: <MediaIcon />, path: '' },
]

export const cmsFooterItems: CmsNavItemConfig[] = [
  { key: 'settings', icon: <SettingsIcon />, path: '' },
  { key: 'logout', icon: <LogoutIcon /> },
]
