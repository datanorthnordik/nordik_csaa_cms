import { render, screen } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { apiClient } from '../api/apiClient'
import i18n from '../i18n'
import { createAppStore } from '../store/store'
import { EventEditorPage } from './EventEditorPage'

vi.mock('../api/apiClient', () => ({
  apiClient: {
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
    delete: vi.fn(),
  },
}))

const mockedGet = vi.mocked(apiClient.get)

function renderPage(initialEntry = '/events/new') {
  const store = createAppStore()

  return render(
    <Provider store={store}>
      <MemoryRouter initialEntries={[initialEntry]}>
        <Routes>
          <Route path="/events/new" element={<EventEditorPage />} />
          <Route path="/events/:id" element={<EventEditorPage mode="view" />} />
          <Route path="/events/:id/edit" element={<EventEditorPage />} />
        </Routes>
      </MemoryRouter>
    </Provider>,
  )
}

beforeEach(async () => {
  mockedGet.mockReset()
  await i18n.changeLanguage('en')

  mockedGet.mockResolvedValue({ data: { items: [] } })
})

describe('EventEditorPage', () => {
  it('renders the shared breadcrumb on the create event route', async () => {
    renderPage()

    const breadcrumb = await screen.findByRole('navigation', { name: /breadcrumb/i })

    expect(screen.getByRole('heading', { name: /create new event/i })).toBeDefined()
    expect(breadcrumb.textContent).toContain('Events')
    expect(breadcrumb.textContent).toContain('Create new event')
  })

  it('renders Publish and Save Draft action buttons on the create route', async () => {
    renderPage()
    await screen.findByRole('heading', { name: /create new event/i })
    expect(screen.getByRole('button', { name: /publish/i })).toBeDefined()
    expect(screen.getByRole('button', { name: /save draft/i })).toBeDefined()
  })

  it('does not show a delete button on the create route (no entry to delete)', async () => {
    renderPage()
    await screen.findByRole('heading', { name: /create new event/i })
    expect(screen.queryByRole('button', { name: /delete event/i })).toBeNull()
  })

  it('does not render action buttons in view mode', async () => {
    // fetchEventLookups makes parallel calls to /api/events/locations and /api/events/galleries
    // before fetchEventById — use mockImplementation to route by URL instead of once-queue
    mockedGet.mockImplementation((url: string) => {
      if (url === '/api/events/1') {
        return Promise.resolve({
          data: {
            id: 1, title: 'Test Event', published: true, categories: [],
            event_type: 'single_day_all_day', start_at: '2026-05-01T00:00:00Z',
            privacy_type: 'public', private_audiences: [], request_review: false,
            review_email_list: [], teaser: '', description_html: '', contact_name: '',
            contact_email: '', contact_phone: '', contact_ext: '', contact_fax: '',
            location_mode: 'none', show_display_image_when_viewing: false,
            registration_enabled: false, registration_url: '', repeat_enabled: false,
            recurrence_interval: 1, occurrences: [], attachments: [],
            created_at: '2026-01-01T00:00:00Z', updated_at: '2026-01-02T00:00:00Z',
            show_title: true,
          },
        })
      }
      return Promise.resolve({ data: { items: [] } })
    })
    renderPage('/events/1')
    await screen.findByRole('navigation', { name: /breadcrumb/i })
    expect(screen.queryByRole('button', { name: /publish/i })).toBeNull()
    expect(screen.queryByRole('button', { name: /save draft/i })).toBeNull()
  })
})
