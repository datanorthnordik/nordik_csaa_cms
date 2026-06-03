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
  { key: 'newsletters', icon: <NewslettersIcon />, path: '/newsletters' },
  { key: 'press', icon: <PressIcon />, path: '/press' },
  { key: 'memorial', icon: <MemorialIcon />, path: '/memorial' },
  { key: 'resources', icon: <ResourcesIcon />, path: '/resources' },
  { key: 'media', icon: <MediaIcon />, path: '/media-library' },
]

export const cmsFooterItems: CmsNavItemConfig[] = [
  { key: 'logout', icon: <LogoutIcon /> },
]
