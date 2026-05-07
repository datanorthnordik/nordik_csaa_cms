import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter } from 'react-router-dom'
import { I18nextProvider } from 'react-i18next'
import { describe, expect, it, vi } from 'vitest'
import { CmsDrawer, type DrawerNavItem } from './CmsDrawer'
import { CmsHeader } from './CmsHeader'
import { CmsLayout } from './CmsLayout'
import { CmsAppShell } from './CmsAppShell'
import { createAppStore } from '../store/store'
import i18n from '../i18n'

// Helper to render components that use i18n
function renderWithI18n(ui: React.ReactElement) {
  return render(ui, {
    wrapper: ({ children }) => <I18nextProvider i18n={i18n}>{children}</I18nextProvider>,
  })
}

const navItems: DrawerNavItem[] = [
  {
    key: 'home',
    label: 'Home',
    icon: <span>H</span>,
    href: '/home',
    onClick: vi.fn(),
  },
  {
    key: 'profile',
    label: 'Profile',
    icon: <span>P</span>,
    onClick: vi.fn(),
  },
]

const footerItems: DrawerNavItem[] = [
  {
    key: 'support',
    label: 'Support',
    icon: <span>S</span>,
    onClick: vi.fn(),
  },
]

describe('CmsHeader', () => {
  it('renders logos and action buttons', () => {
    const onMenuClick = vi.fn()
    renderWithI18n(<CmsHeader onMenuClick={onMenuClick} />)

    expect(screen.getByLabelText(/open navigation menu/i)).toBeDefined()
    expect(screen.getByLabelText(/notifications/i)).toBeDefined()
    expect(screen.getByLabelText(/help/i)).toBeDefined()
    expect(screen.getByAltText(/children of shingwauk alumni association/i)).toBeDefined()
    expect(screen.getByAltText(/nordik institute/i)).toBeDefined()
  })

  it('calls onMenuClick when the menu button is clicked', () => {
    const onMenuClick = vi.fn()
    renderWithI18n(<CmsHeader onMenuClick={onMenuClick} />)

    fireEvent.click(screen.getByLabelText(/open navigation menu/i))
    expect(onMenuClick).toHaveBeenCalledTimes(1)
  })
})

describe('CmsDrawer', () => {
  it('renders profile, navigation, and footer items', () => {
    renderWithI18n(
      <CmsDrawer
        userName="Test User"
        userRole="Administrator"
        navItems={navItems}
        footerItems={footerItems}
        open
      />
    )

    expect(screen.getByText('Test User')).toBeDefined()
    expect(screen.getByText('Administrator')).toBeDefined()
    expect(screen.getByRole('link', { name: /home/i })).toBeDefined()
    expect(screen.getByRole('button', { name: /profile/i })).toBeDefined()
    expect(screen.getByRole('button', { name: /support/i })).toBeDefined()
  })

  it('calls onClose when the backdrop is clicked and onSubmitTicket when ticket button is clicked', () => {
    const onClose = vi.fn()
    const onSubmitTicket = vi.fn()

    const { container } = renderWithI18n(
      <CmsDrawer
        userName="Test User"
        userRole="Administrator"
        navItems={navItems}
        footerItems={footerItems}
        open
        onClose={onClose}
        onSubmitTicket={onSubmitTicket}
      />
    )

    const backdrop = container.querySelector('div[aria-hidden="true"]')
    expect(backdrop).not.toBeNull()
    if (backdrop) {
      fireEvent.click(backdrop)
      expect(onClose).toHaveBeenCalledTimes(1)
    }

    // Use a more flexible selector to find the submit ticket button
    const submitButton = screen.getByRole('button', { name: /ticket/i })
    fireEvent.click(submitButton)
    expect(onSubmitTicket).toHaveBeenCalledTimes(1)
  })
})

describe('CmsLayout', () => {
  it('renders children and opens the drawer when the header menu is clicked', () => {
    renderWithI18n(
      <CmsLayout
        userName="Layout User"
        userRole="Team Lead"
        navItems={navItems}
        footerItems={footerItems}
      >
        <div>Page content</div>
      </CmsLayout>
    )

    expect(screen.getByText('Page content')).toBeDefined()
    expect(screen.getByLabelText(/primary navigation/i)).toBeDefined()

    fireEvent.click(screen.getByLabelText(/open navigation menu/i))
    const drawer = screen.getByLabelText(/primary navigation/i)
    expect(drawer.className).toContain('drawerOpen')
  })
})

describe('CmsAppShell', () => {
  it('renders the CMS shell and calls onLogout when logout is clicked', async () => {
    const onLogout = vi.fn()
    const onSubmitTicket = vi.fn()
    const store = createAppStore()

    render(
      <I18nextProvider i18n={i18n}>
        <Provider store={store}>
          <MemoryRouter>
            <CmsAppShell onLogout={onLogout} onSubmitTicket={onSubmitTicket}>
              <div>Shell content</div>
            </CmsAppShell>
          </MemoryRouter>
        </Provider>
      </I18nextProvider>
    )

    expect(screen.getByText('Shell content')).toBeDefined()

    await waitFor(() => {
      expect(screen.getByRole('button', { name: /submit ticket/i })).toBeDefined()
    })

    fireEvent.click(screen.getByRole('button', { name: /submit ticket/i }))
    expect(onSubmitTicket).toHaveBeenCalledTimes(1)

    fireEvent.click(screen.getByRole('button', { name: /logout/i }))
    expect(onLogout).toHaveBeenCalledTimes(1)
  })
})
