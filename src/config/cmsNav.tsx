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
  label: string
  icon: ReactNode
  path?: string
}

export const cmsNavItems: CmsNavItemConfig[] = [
  { key: 'dashboard', label: 'Dashboard', icon: <DashboardIcon />, path: '' },
  { key: 'pages', label: 'Pages', icon: <PagesIcon />, path: '' },
  { key: 'events', label: 'Events', icon: <EventsIcon />, path: '' },
  { key: 'newsletters', label: 'Newsletters', icon: <NewslettersIcon />, path: '' },
  { key: 'press', label: 'Press Entries', icon: <PressIcon />, path: '' },
  { key: 'memorial', label: 'Memorial Entries', icon: <MemorialIcon />, path: '' },
  { key: 'resources', label: 'Resources', icon: <ResourcesIcon />, path: '' },
  { key: 'media', label: 'Media Library', icon: <MediaIcon />, path: '' },
]

export const cmsFooterItems: CmsNavItemConfig[] = [
  { key: 'settings', label: 'Settings', icon: <SettingsIcon />, path: '' },
  { key: 'logout', label: 'Logout', icon: <LogoutIcon /> },
]
