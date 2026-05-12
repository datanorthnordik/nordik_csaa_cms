import { fireEvent, render, screen, waitFor, within } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { apiClient } from '../api/apiClient'
import i18n from '../i18n'
import { createAppStore } from '../store/store'
import { PagesListPage } from './PagesListPage'

vi.mock('../api/apiClient', () => ({
  apiClient: {
    get: vi.fn(),
    delete: vi.fn(),
  },
}))

const mockedGet = vi.mocked(apiClient.get)

function renderPage() {
  const store = createAppStore()
  return render(
    <Provider store={store}>
      <MemoryRouter>
        <PagesListPage />
      </MemoryRouter>
    </Provider>,
  )
}

const mockPageItem = {
  id: 3,
  page_title: 'About Us',
  url_slug: '/about-us',
  status: 'published',
  last_modified: '2026-05-01T10:00:00Z',
  modified_by: 1,
  modified_by_name: 'Admin',
  created_at: '2026-04-01T10:00:00Z',
  updated_at: '2026-05-01T10:00:00Z',
}

const mockListResponse = {
  data: {
    items: [mockPageItem],
    pagination: {
      page: 1,
      page_size: 10,
      total_items: 1,
      total_pages: 1,
      has_next: false,
      has_prev: false,
    },
    applied_filters: {
      page: 1,
      page_size: 10,
      search_term: '',
      status: '',
      sort_by: 'last_modified',
      sort_order: 'desc',
    },
  },
}

beforeEach(async () => {
  mockedGet.mockReset()
  window.localStorage.clear()
  window.sessionStorage.clear()
  await i18n.changeLanguage('en')

  mockedGet.mockResolvedValue(mockListResponse)
})

describe('PagesListPage', () => {
  it('renders the breadcrumb above the pages list', async () => {
    renderPage()

    const breadcrumb = await screen.findByRole('navigation', {
      name: /breadcrumb/i,
    })
    expect(within(breadcrumb).getByText(/^pages$/i)).toBeDefined()
  })

  it('renders the create new page button', async () => {
    renderPage()
    expect(
      await screen.findByRole('button', { name: /create new page/i }),
    ).toBeDefined()
  })

  it('renders page items returned by the API', async () => {
    renderPage()
    const items = await screen.findAllByText(/about us/i)
    expect(items.length).toBeGreaterThan(0)
  })

  it('shows the page URL slug in the list', async () => {
    renderPage()
    const slugElements = await screen.findAllByText('/about-us')
    expect(slugElements.length).toBeGreaterThan(0)
  })

  it('shows the empty state when the API returns no items', async () => {
    mockedGet.mockResolvedValue({
      data: {
        ...mockListResponse.data,
        items: [],
        pagination: {
          ...mockListResponse.data.pagination,
          total_items: 0,
        },
      },
    })
    renderPage()
    expect(await screen.findByText(/no pages found/i)).toBeDefined()
  })

  it('opens a delete confirmation dialog when the delete button is clicked', async () => {
    renderPage()

    await screen.findAllByText(/about us/i)

    const deleteButtons = screen.getAllByRole('button', { name: /delete page/i })
    fireEvent.click(deleteButtons[0])

    const dialog = screen.getByRole('dialog')
    expect(
      within(dialog).getByRole('heading', { name: /delete page\?/i }),
    ).toBeDefined()
    expect(screen.getByRole('button', { name: /^cancel$/i })).toBeDefined()
  })

  it('closes the delete dialog when cancel is clicked', async () => {
    renderPage()

    await screen.findAllByText(/about us/i)

    const deleteButtons = screen.getAllByRole('button', { name: /delete page/i })
    fireEvent.click(deleteButtons[0])

    expect(screen.getByRole('dialog')).toBeDefined()

    fireEvent.click(screen.getByRole('button', { name: /^cancel$/i }))

    await waitFor(() => {
      expect(screen.queryByRole('dialog')).toBeNull()
    })
  })

  it('renders a search input for filtering pages', async () => {
    renderPage()

    await screen.findAllByText(/about us/i)

    expect(
      screen.getByRole('searchbox', { name: /search pages/i }),
    ).toBeDefined()
  })
})
