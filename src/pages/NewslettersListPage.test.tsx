import type { ReactNode } from 'react'
import { fireEvent, render, screen, waitFor, within } from '@testing-library/react'
import { MemoryRouter } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import { NewslettersListPage } from './NewslettersListPage'

const {
  mockNavigate,
  mockFetch,
  mockRemove,
  toastError,
  toastSuccess,
  useNewsletterEntriesMock,
} = vi.hoisted(() => ({
  mockNavigate: vi.fn(),
  mockFetch: vi.fn(),
  mockRemove: vi.fn(),
  toastError: vi.fn(),
  toastSuccess: vi.fn(),
  useNewsletterEntriesMock: vi.fn(),
}))

vi.mock('../hooks/useNewsletterEntries', () => ({
  useNewsletterEntries: (...args: unknown[]) => useNewsletterEntriesMock(...args),
}))

vi.mock('../components/CmsAppShell', () => ({
  CmsAppShell: ({ children }: { children?: ReactNode }) => (
    <div data-testid="cms-shell">{children}</div>
  ),
}))

vi.mock('../components/Loader', () => ({
  Loader: () => <div>Loading newsletters</div>,
}))

vi.mock('react-hot-toast', () => ({
  default: {
    error: toastError,
    success: toastSuccess,
  },
}))

vi.mock('react-router-dom', async () => {
  const actual =
    await vi.importActual<typeof import('react-router-dom')>('react-router-dom')

  return {
    ...actual,
    useNavigate: () => mockNavigate,
  }
})

function buildHookState(overrides: Record<string, unknown> = {}) {
  return {
    entries: [
      {
        id: '12',
        title: 'Community Update',
        category: 'csaa',
        sendDate: '2026-06-12',
        contentHtml: '<p>Hello</p>',
        status: 'published',
        visibility: 'public',
        publishAt: null,
        media: [],
        createdAt: '2026-06-01T00:00:00Z',
        updatedAt: '2026-06-02T00:00:00Z',
      },
    ],
    loading: false,
    error: null,
    totalPages: 3,
    total: 12,
    fetch: mockFetch,
    remove: mockRemove,
    ...overrides,
  }
}

function renderPage() {
  return render(
    <MemoryRouter>
      <NewslettersListPage />
    </MemoryRouter>,
  )
}

describe('NewslettersListPage', () => {
  beforeEach(async () => {
    vi.clearAllMocks()
    await i18n.changeLanguage('en')
    mockFetch.mockResolvedValue(undefined)
    useNewsletterEntriesMock.mockReturnValue(buildHookState())
  })

  it('renders newsletter rows, navigates actions, and refetches on filter changes', async () => {
    renderPage()

    await waitFor(() => {
      expect(mockFetch).toHaveBeenCalledWith(1, 10, {
        status: undefined,
        search: undefined,
        sortBy: 'sendDate',
        sortOrder: 'desc',
      })
    })

    const breadcrumb = screen.getByRole('navigation', { name: /breadcrumb/i })
    expect(within(breadcrumb).getByText(/^newsletters$/i)).toBeDefined()
    expect(screen.getByRole('heading', { name: /^newsletters$/i })).toBeDefined()
    expect(screen.getByText('CSAA')).toBeDefined()
    expect(screen.getByRole('button', { name: 'Community Update' })).toBeDefined()
    expect(screen.getByText(/updated /i)).toBeDefined()

    fireEvent.change(screen.getByRole('searchbox', { name: /search newsletters/i }), {
      target: { value: 'June' },
    })

    await waitFor(() => {
      expect(mockFetch).toHaveBeenLastCalledWith(1, 10, {
        status: undefined,
        search: 'June',
        sortBy: 'sendDate',
        sortOrder: 'desc',
      })
    })

    fireEvent.change(screen.getByLabelText(/sort by/i), {
      target: { value: 'title' },
    })

    await waitFor(() => {
      expect(mockFetch).toHaveBeenLastCalledWith(1, 10, {
        status: undefined,
        search: 'June',
        sortBy: 'title',
        sortOrder: 'desc',
      })
    })

    fireEvent.click(screen.getByRole('button', { name: /create new newsletter/i }))
    expect(mockNavigate).toHaveBeenCalledWith('/newsletters/new')

    fireEvent.click(screen.getByRole('button', { name: 'Community Update' }))
    expect(mockNavigate).toHaveBeenCalledWith('/newsletters/12/edit')
  })

  it('shows error and empty states when the hook reports no results', () => {
    useNewsletterEntriesMock.mockReturnValue(
      buildHookState({
        entries: [],
        error: 'Query failed',
        total: 0,
      }),
    )

    renderPage()

    expect(screen.getByText('Query failed')).toBeDefined()
    expect(screen.getByText('No newsletters found')).toBeDefined()
    expect(
      screen.getByText(/adjust your filters, or create a new newsletter/i),
    ).toBeDefined()
  })

  it('shows the loading state while newsletters are being fetched', () => {
    useNewsletterEntriesMock.mockReturnValue(
      buildHookState({
        entries: [],
        loading: true,
        total: 0,
      }),
    )

    renderPage()

    expect(screen.getByText('Loading newsletters')).toBeDefined()
  })

  it('opens the delete dialog and removes the selected newsletter', async () => {
    mockRemove.mockResolvedValue(undefined)

    renderPage()

    fireEvent.click(screen.getAllByRole('button', { name: /^delete$/i })[0])

    const dialog = await screen.findByRole('dialog')
    expect(within(dialog).getByRole('heading', { name: /delete newsletter\?/i })).toBeDefined()
    expect(within(dialog).getByText('"Community Update"', { selector: 'strong' })).toBeDefined()

    fireEvent.click(within(dialog).getByRole('button', { name: /^delete$/i }))

    await waitFor(() => {
      expect(mockRemove).toHaveBeenCalledWith('12')
    })
    expect(toastSuccess).toHaveBeenCalledWith('Newsletter deleted')
  })

  it('shows a toast when the initial fetch request rejects', async () => {
    mockFetch.mockRejectedValueOnce(new Error('Load failed'))
    useNewsletterEntriesMock.mockReturnValue(
      buildHookState({
        entries: [],
        total: 0,
      }),
    )

    renderPage()

    await waitFor(() => {
      expect(toastError).toHaveBeenCalledWith('Load failed')
    })
  })
})
