import type { ReactNode } from 'react'
import { fireEvent, render, screen, waitFor, within } from '@testing-library/react'
import { MemoryRouter } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import { MemorialEntriesListPage } from './MemorialEntriesListPage'

const {
  mockNavigate,
  mockFetch,
  mockRemove,
  toastError,
  toastSuccess,
  useMemorialEntriesMock,
} = vi.hoisted(() => ({
  mockNavigate: vi.fn(),
  mockFetch: vi.fn(),
  mockRemove: vi.fn(),
  toastError: vi.fn(),
  toastSuccess: vi.fn(),
  useMemorialEntriesMock: vi.fn(),
}))

vi.mock('../hooks/useMemorialEntries', () => ({
  useMemorialEntries: (...args: unknown[]) => useMemorialEntriesMock(...args),
}))

vi.mock('../components/CmsAppShell', () => ({
  CmsAppShell: ({ children }: { children?: ReactNode }) => (
    <div data-testid="cms-shell">{children}</div>
  ),
}))

vi.mock('../components/Loader', () => ({
  Loader: ({ label }: { label?: string }) => <div>{label ?? 'Loading'}</div>,
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
    items: [
      {
        id: '12',
        fullName: 'Ada Lovelace',
        affiliation: 'Analytical Engine',
        category: 'founder',
        categoryLabel: 'Founder',
        status: 'published',
        dateOfBirth: '1815-12-10',
        dateOfPassing: '1852-11-27',
        createdAt: '2026-05-20T00:00:00Z',
        updatedAt: '2026-05-21T00:00:00Z',
        publishedAt: '2026-05-21T00:00:00Z',
      },
    ],
    pagination: {
      page: 1,
      pageSize: 10,
      totalItems: 12,
      totalPages: 3,
      hasNext: true,
      hasPrev: false,
    },
    loading: false,
    error: null,
    fetch: mockFetch,
    remove: mockRemove,
    ...overrides,
  }
}

function renderPage() {
  return render(
    <MemoryRouter>
      <MemorialEntriesListPage />
    </MemoryRouter>,
  )
}

describe('MemorialEntriesListPage', () => {
  beforeEach(async () => {
    vi.clearAllMocks()
    await i18n.changeLanguage('en')
    mockFetch.mockResolvedValue(undefined)
    mockRemove.mockResolvedValue(undefined)
    useMemorialEntriesMock.mockReturnValue(buildHookState())
  })

  it('renders entries, navigates actions, and refetches on search and filter changes', async () => {
    renderPage()

    await waitFor(() => {
      expect(mockFetch).toHaveBeenCalledWith({
        searchTerm: '',
        status: 'all',
        category: '',
        page: 1,
        pageSize: 10,
      })
    })

    const breadcrumb = screen.getByRole('navigation', { name: /breadcrumb/i })
    expect(within(breadcrumb).getByText(/^memorial entries$/i)).toBeDefined()
    expect(screen.getByRole('heading', { name: 'Memorial Entries' })).toBeDefined()
    expect(screen.getByText('Analytical Engine')).toBeDefined()
    expect(screen.getByText('Founder')).toBeDefined()
    expect(screen.getByText('Published')).toBeDefined()

    fireEvent.change(screen.getByRole('searchbox', { name: 'Search entries' }), {
      target: { value: 'Ada' },
    })

    await waitFor(() => {
      expect(mockFetch).toHaveBeenLastCalledWith({
        searchTerm: 'Ada',
        status: 'all',
        category: '',
        page: 1,
        pageSize: 10,
      })
    })

    fireEvent.click(screen.getByRole('button', { name: /filters/i }))

    fireEvent.change(screen.getByLabelText('Status'), {
      target: { value: 'review' },
    })

    await waitFor(() => {
      expect(mockFetch).toHaveBeenLastCalledWith({
        searchTerm: 'Ada',
        status: 'review',
        category: '',
        page: 1,
        pageSize: 10,
      })
    })

    fireEvent.change(screen.getByLabelText('Category'), {
      target: { value: 'friend' },
    })

    await waitFor(() => {
      expect(mockFetch).toHaveBeenLastCalledWith({
        searchTerm: 'Ada',
        status: 'review',
        category: 'friend',
        page: 1,
        pageSize: 10,
      })
    })

    fireEvent.click(screen.getByRole('button', { name: 'Add Memorial Entry' }))
    expect(mockNavigate).toHaveBeenCalledWith('/memorial/new')

    fireEvent.click(screen.getByRole('button', { name: 'Ada Lovelace' }))
    expect(mockNavigate).toHaveBeenCalledWith('/memorial/12/edit')

    fireEvent.click(screen.getByRole('button', { name: 'Edit entry' }))
    expect(mockNavigate).toHaveBeenCalledWith('/memorial/12/edit')
  })

  it('shows error and empty states when there are no results', () => {
    useMemorialEntriesMock.mockReturnValue(
      buildHookState({
        items: [],
        error: 'Query failed',
        pagination: {
          page: 1,
          pageSize: 10,
          totalItems: 0,
          totalPages: 0,
          hasNext: false,
          hasPrev: false,
        },
      }),
    )

    renderPage()

    expect(screen.getByText('Query failed')).toBeDefined()
    expect(screen.getByText('No memorial entries found')).toBeDefined()
    expect(screen.getByText('Create your first entry to get started.')).toBeDefined()
  })

  it('shows the loading state while entries are being fetched', () => {
    useMemorialEntriesMock.mockReturnValue(
      buildHookState({
        items: [],
        loading: true,
        pagination: {
          page: 1,
          pageSize: 10,
          totalItems: 0,
          totalPages: 0,
          hasNext: false,
          hasPrev: false,
        },
      }),
    )

    renderPage()

    expect(screen.getByText('Loading...')).toBeDefined()
  })

  it('opens the delete dialog, removes the selected entry, and refetches the previous page when needed', async () => {
    useMemorialEntriesMock.mockReturnValue(
      buildHookState({
        pagination: {
          page: 2,
          pageSize: 10,
          totalItems: 11,
          totalPages: 2,
          hasNext: false,
          hasPrev: true,
        },
      }),
    )

    renderPage()

    fireEvent.click(screen.getByRole('button', { name: 'Delete entry' }))

    const dialog = await screen.findByRole('dialog')
    expect(within(dialog).getByRole('heading', { name: 'Delete memorial entry?' })).toBeDefined()
    expect(within(dialog).getByText('Ada Lovelace', { selector: 'strong' })).toBeDefined()

    fireEvent.click(within(dialog).getByRole('button', { name: 'Delete' }))

    await waitFor(() => {
      expect(mockRemove).toHaveBeenCalledWith('12')
    })
    await waitFor(() => {
      expect(mockFetch).toHaveBeenLastCalledWith({
        searchTerm: '',
        status: 'all',
        category: '',
        page: 1,
        pageSize: 10,
      })
    })
    expect(toastSuccess).toHaveBeenCalledWith('Memorial entry deleted.')
  })

  it('shows a toast when the initial fetch request rejects', async () => {
    mockFetch.mockRejectedValueOnce(new Error('Load failed'))
    useMemorialEntriesMock.mockReturnValue(
      buildHookState({
        items: [],
        pagination: {
          page: 1,
          pageSize: 10,
          totalItems: 0,
          totalPages: 0,
          hasNext: false,
          hasPrev: false,
        },
      }),
    )

    renderPage()

    await waitFor(() => {
      expect(toastError).toHaveBeenCalledWith('Load failed')
    })
  })
})
