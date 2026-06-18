import { describe, expect, it, vi } from 'vitest'

vi.mock('../components/icons', () => ({
  DashboardIcon: () => <svg data-testid="dashboard-icon" />,
  MenusIcon: () => <svg data-testid="menus-icon" />,
  PagesIcon: () => <svg data-testid="pages-icon" />,
  EventsIcon: () => <svg data-testid="events-icon" />,
  NewslettersIcon: () => <svg data-testid="newsletters-icon" />,
  PressIcon: () => <svg data-testid="press-icon" />,
  MemorialIcon: () => <svg data-testid="memorial-icon" />,
  ResourcesIcon: () => <svg data-testid="resources-icon" />,
  MediaIcon: () => <svg data-testid="media-icon" />,
  VideoIcon: () => <svg data-testid="video-icon" />,
  LogoutIcon: () => <svg data-testid="logout-icon" />,
}))

import { cmsFooterItems, cmsNavItems } from './cmsNav'

describe('cmsNav', () => {
  it('includes the menus entry in the drawer navigation', () => {
    expect(cmsNavItems[1]).toMatchObject({
      key: 'menus',
      path: '/menus',
    })
    expect(cmsNavItems[1]?.icon).not.toBeNull()
  })

  it('wires the resources entry to the library route', () => {
    expect(cmsNavItems.find((item) => item.key === 'resources')).toMatchObject({
      key: 'resources',
      path: '/resources',
    })
  })

  it('includes a knowledge center entry that links to its review queue', () => {
    expect(cmsNavItems.find((item) => item.key === 'knowledgeCenter')).toMatchObject({
      key: 'knowledgeCenter',
      path: '/knowledge-center',
    })
  })

  it('includes a videos entry that links to the video library', () => {
    expect(cmsNavItems.find((item) => item.key === 'videos')).toMatchObject({
      key: 'videos',
      path: '/videos',
    })
  })

  it('keeps only logout in the footer navigation', () => {
    expect(cmsFooterItems.map((item) => item.key)).toEqual(['logout'])
  })
})
