import { render, screen } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { apiClient } from '../api/apiClient'
import i18n from '../i18n'
import { createAppStore } from '../store/store'
import { PageEditorPage } from './PageEditorPage'

vi.mock('../api/apiClient', () => ({
  apiClient: {
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
    delete: vi.fn(),
  },
}))

const mockedGet = vi.mocked(apiClient.get)

function renderPage(initialEntry = '/pages/new') {
  const store = createAppStore()

  return render(
    <Provider store={store}>
      <MemoryRouter initialEntries={[initialEntry]}>
        <Routes>
          <Route path="/pages/new" element={<PageEditorPage />} />
          <Route path="/pages/:id" element={<PageEditorPage mode="view" />} />
          <Route path="/pages/:id/edit" element={<PageEditorPage />} />
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

describe('PageEditorPage', () => {
  it('renders the create breadcrumb on the new page route', async () => {
    renderPage()

    const breadcrumb = await screen.findByRole('navigation', {
      name: /breadcrumb/i,
    })

    expect(
      screen.getByRole('heading', { name: /create new page/i }),
    ).toBeDefined()
    expect(breadcrumb.textContent).toContain('Pages')
    expect(breadcrumb.textContent).toContain('Create New Page')
  })

  it('renders a back-to-list button on the create route', async () => {
    renderPage()

    await screen.findByRole('heading', { name: /create new page/i })

    expect(
      screen.getByRole('button', { name: /back to pages/i }),
    ).toBeDefined()
  })

  it('renders save draft and publish buttons in edit mode', async () => {
    renderPage()

    await screen.findByRole('heading', { name: /create new page/i })

    expect(
      screen.getByRole('button', { name: /save as draft/i }),
    ).toBeDefined()
    expect(
      screen.getByRole('button', { name: /publish page/i }),
    ).toBeDefined()
  })

  it('renders the view breadcrumb when mode is view', async () => {
    mockedGet.mockResolvedValue({
      data: {
        id: 5,
        page_title: 'About Us',
        url_slug: '/about-us',
        status: 'published',
        hero_image_enabled: false,
        hero_image_url: '',
        hero_image_object_key: '',
        hero_image_fetch_url: '',
        seo_page_title: '',
        seo_page_description: '',
        created_by: null,
        created_by_name: '',
        modified_by: null,
        modified_by_name: '',
        last_modified: '2026-05-01T10:00:00Z',
        created_at: '2026-04-01T10:00:00Z',
        updated_at: '2026-05-01T10:00:00Z',
      },
    })

    renderPage('/pages/5')

    const breadcrumb = await screen.findByRole('navigation', {
      name: /breadcrumb/i,
    })
    expect(breadcrumb.textContent).toContain('View Page')
  })

  it('shows an invalid link message for a non-numeric page id', async () => {
    renderPage('/pages/not-a-number/edit')

    expect(
      await screen.findByRole('heading', { name: /invalid page link/i }),
    ).toBeDefined()
  })
})
