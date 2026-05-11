import { fireEvent, render, screen, within } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { apiClient } from '../api/apiClient'
import i18n from '../i18n'
import { createAppStore } from '../store/store'
import { EventsListPage } from './EventsListPage'

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
        <EventsListPage />
      </MemoryRouter>
    </Provider>,
  )
}

beforeEach(async () => {
  mockedGet.mockReset()
  window.localStorage.clear()
  window.sessionStorage.clear()
  await i18n.changeLanguage('en')

  mockedGet.mockResolvedValue({
    data: {
      items: [
        {
          id: 7,
          title: 'Spring Feast',
          categories: ['Community'],
          status: 'published',
          published: true,
          event_type: 'single_day_all_day',
          start_at: '2026-05-08T00:00:00Z',
          end_at: null,
          date_display: 'May 8, 2026',
          created_at: '2026-05-01T00:00:00Z',
          updated_at: '2026-05-02T00:00:00Z',
        },
      ],
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
        statuses: [],
        start_date: null,
        end_date: null,
        date_range: 'custom',
        sort_by: 'start_at',
        sort_order: 'desc',
      },
    },
  })
})

describe('EventsListPage', () => {
  it('renders the shared breadcrumb above the events list', async () => {
    renderPage()

    const breadcrumb = await screen.findByRole('navigation', { name: /breadcrumb/i })

    expect(within(breadcrumb).getByText(/^events$/i)).toBeDefined()
    expect(within(breadcrumb).getByText(/^list$/i)).toBeDefined()
  })

  it('shows a confirmation dialog before deleting an event', async () => {
    renderPage()

    const triggers = await screen.findAllByRole('button', {
      name: /open event actions/i,
    })

    fireEvent.click(triggers[0])

    expect(screen.getAllByRole('menuitem', { name: /view event/i }).length).toBeGreaterThan(0)
    expect(screen.getAllByRole('menuitem', { name: /edit event/i }).length).toBeGreaterThan(0)

    fireEvent.click(screen.getAllByRole('menuitem', { name: /delete event/i })[0])

    const dialog = screen.getByRole('dialog')

    expect(within(dialog).getByRole('heading', { name: /delete event\?/i })).toBeDefined()
    expect(
      within(dialog).getByText((_, element) =>
        (element?.tagName === 'P' &&
          element.textContent?.includes(
          'This will permanently delete "Spring Feast" from the event list.',
        )) ??
        false,
      ),
    ).toBeDefined()
    expect(within(dialog).getByText('"Spring Feast"', { selector: 'strong' })).toBeDefined()
    expect(screen.getByRole('button', { name: /^cancel$/i })).toBeDefined()
    expect(screen.getAllByRole('button', { name: /delete event/i }).length).toBeGreaterThan(0)
  })
})
