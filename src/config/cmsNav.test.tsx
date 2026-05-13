import { render, screen } from '@testing-library/react'
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
  SettingsIcon: () => <svg data-testid="settings-icon" />,
  LogoutIcon: () => <svg data-testid="logout-icon" />,
}))

import { cmsFooterItems, cmsNavItems } from './cmsNav'

describe('cmsNav', () => {
  it('includes the menus entry in the drawer navigation', () => {
    expect(cmsNavItems[1]).toMatchObject({
      key: 'menus',
      path: '/menus',
    })

    render(<>{cmsNavItems[1]?.icon}</>)
    expect(screen.getByTestId('menus-icon')).toBeDefined()
  })

  it('keeps settings and logout in the footer navigation', () => {
    expect(cmsFooterItems.map((item) => item.key)).toEqual(['settings', 'logout'])
  })
})
