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
  { key: 'dashboard', label: 'Dashboard', icon: <DashboardIcon />, path: '/dashboard' },
  { key: 'pages', label: 'Pages', icon: <PagesIcon />, path: '/pages' },
  { key: 'events', label: 'Events', icon: <EventsIcon />, path: '/events' },
  { key: 'newsletters', label: 'Newsletters', icon: <NewslettersIcon />, path: '/newsletters' },
  { key: 'press', label: 'Press Entries', icon: <PressIcon />, path: '/press' },
  { key: 'memorial', label: 'Memorial Entries', icon: <MemorialIcon />, path: '/memorial' },
  { key: 'resources', label: 'Resources', icon: <ResourcesIcon />, path: '/resources' },
  { key: 'media', label: 'Media Library', icon: <MediaIcon />, path: '/media' },
]

export const cmsFooterItems: CmsNavItemConfig[] = [
  { key: 'settings', label: 'Settings', icon: <SettingsIcon />, path: '/settings' },
  { key: 'logout', label: 'Logout', icon: <LogoutIcon /> },
]
