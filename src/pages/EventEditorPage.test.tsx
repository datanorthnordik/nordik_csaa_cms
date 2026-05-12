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
})
