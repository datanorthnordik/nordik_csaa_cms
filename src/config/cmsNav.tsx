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
  VideoIcon,
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
  { key: 'blogs', icon: <PressIcon />, path: '/blogs' },
  { key: 'events', icon: <EventsIcon />, path: '/events' },
  { key: 'newsletters', icon: <NewslettersIcon />, path: '/newsletters' },
  { key: 'books', icon: <ResourcesIcon />, path: '/books' },
  { key: 'knowledgeCenter', icon: <ResourcesIcon />, path: '/knowledge-center' },
  { key: 'press', icon: <PressIcon />, path: '/press' },
  { key: 'memorial', icon: <MemorialIcon />, path: '/memorial' },
  { key: 'resources', icon: <ResourcesIcon />, path: '/resources' },
  { key: 'bookshelf', icon: <ResourcesIcon />, path: '/bookshelf' },
  { key: 'media', icon: <MediaIcon />, path: '/media-library' },
  { key: 'videos', icon: <VideoIcon />, path: '/videos' },
]

export const cmsFooterItems: CmsNavItemConfig[] = [
  { key: 'logout', icon: <LogoutIcon /> },
]
