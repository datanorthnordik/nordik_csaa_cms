import { render, screen } from '@testing-library/react'
import { I18nextProvider } from 'react-i18next'
import { Provider } from 'react-redux'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it } from 'vitest'
import i18n from '../i18n'
import { PRESS_ENTRIES_STORAGE_KEY } from '../lib/pressEntriesStorage'
import { createAppStore } from '../store/store'
import { PressEditorPage } from './PressEditorPage'
import type { PressEntry } from '../lib/pressTypes'

function makeEntry(overrides: Partial<PressEntry> = {}): PressEntry {
  return {
    id: 'press-edit-1',
    title: 'Existing Release',
    releaseDate: '2026-03-15',
    categoryId: null,
    sourceUrl: 'https://example.com',
    contentHtml: '<p>Content</p>',
    status: 'published',
    visibility: 'public',
    publishAt: null,
    media: [],
    createdAt: '2026-01-01T00:00:00.000Z',
    updatedAt: '2026-01-02T00:00:00.000Z',
    ...overrides,
  }
}

function renderPage(initialEntry = '/press/new') {
  const store = createAppStore()
  return render(
    <I18nextProvider i18n={i18n}>
      <Provider store={store}>
        <MemoryRouter initialEntries={[initialEntry]}>
          <Routes>
            <Route path="/press/new" element={<PressEditorPage />} />
            <Route path="/press/:id/edit" element={<PressEditorPage mode="edit" />} />
          </Routes>
        </MemoryRouter>
      </Provider>
    </I18nextProvider>,
  )
}

beforeEach(async () => {
  await i18n.changeLanguage('en')
  window.localStorage.setItem(
    PRESS_ENTRIES_STORAGE_KEY,
    JSON.stringify([makeEntry()]),
  )
})

describe('PressEditorPage', () => {
  it('renders the breadcrumb on the create route', async () => {
    renderPage()
    const breadcrumb = await screen.findByRole('navigation', { name: /breadcrumb/i })
    expect(breadcrumb.textContent).toContain('Press')
  })

  it('shows the create heading in create mode', async () => {
    renderPage()
    expect(
      await screen.findByRole('heading', { name: /create new press entry/i }),
    ).toBeDefined()
  })

  it('shows the edit heading in edit mode', async () => {
    renderPage('/press/press-edit-1/edit')
    expect(
      await screen.findByRole('heading', { name: /edit press entry/i }),
    ).toBeDefined()
  })

  it('renders Publish and Save Draft action buttons via EntryActions', async () => {
    renderPage()
    expect(await screen.findByRole('button', { name: /publish press entry/i })).toBeDefined()
    expect(screen.getByRole('button', { name: /save draft/i })).toBeDefined()
  })

  it('shows "Delete Press Entry" button in edit mode', async () => {
    renderPage('/press/press-edit-1/edit')
    expect(
      await screen.findByRole('button', { name: /delete press entry/i }),
    ).toBeDefined()
  })

  it('does not show a delete button in create mode', async () => {
    renderPage()
    await screen.findByRole('button', { name: /publish/i })
    expect(screen.queryByRole('button', { name: /delete/i })).toBeNull()
  })

  it('does not render action buttons inside the sidebar PublishingControls', async () => {
    renderPage()
    await screen.findByRole('button', { name: /publish/i })
    const sidebar = screen.getByRole('complementary', { name: /publishing/i })
    expect(sidebar.querySelector('button[type="button"]')).toBeNull()
  })
})
