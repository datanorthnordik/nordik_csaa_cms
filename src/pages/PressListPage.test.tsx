import { fireEvent, render, screen, within } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter } from 'react-router-dom'
import { beforeEach, describe, expect, it } from 'vitest'
import i18n from '../i18n'
import { PRESS_ENTRIES_STORAGE_KEY } from '../lib/pressEntriesStorage'
import { createAppStore } from '../store/store'
import { PressListPage } from './PressListPage'
import type { PressEntry } from '../lib/pressTypes'
import { I18nextProvider } from 'react-i18next'

function makeEntry(overrides: Partial<PressEntry> = {}): PressEntry {
  return {
    id: 'entry-1',
    title: 'Community Update',
    releaseDate: '2026-04-01',
    categoryId: null,
    sourceUrl: '',
    contentHtml: '',
    status: 'published',
    visibility: 'public',
    publishAt: null,
    media: [],
    createdAt: '2026-01-01T00:00:00.000Z',
    updatedAt: '2026-01-02T00:00:00.000Z',
    ...overrides,
  }
}

function renderPage() {
  const store = createAppStore()
  return render(
    <I18nextProvider i18n={i18n}>
      <Provider store={store}>
        <MemoryRouter>
          <PressListPage />
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

describe('PressListPage', () => {
  it('renders the breadcrumb navigation', async () => {
    renderPage()
    const breadcrumb = await screen.findByRole('navigation', { name: /breadcrumb/i })
    expect(breadcrumb.textContent).toContain('Press')
  })

  it('renders press entries from localStorage', async () => {
    renderPage()
    expect(await screen.findByText('Community Update')).toBeDefined()
  })

  it('renders multiple entries', async () => {
    window.localStorage.setItem(
      PRESS_ENTRIES_STORAGE_KEY,
      JSON.stringify([
        makeEntry({ id: 'a', title: 'Entry Alpha' }),
        makeEntry({ id: 'b', title: 'Entry Beta' }),
      ]),
    )
    renderPage()
    expect(await screen.findByText('Entry Alpha')).toBeDefined()
    expect(await screen.findByText('Entry Beta')).toBeDefined()
  })

  it('filters entries by search term', async () => {
    window.localStorage.setItem(
      PRESS_ENTRIES_STORAGE_KEY,
      JSON.stringify([
        makeEntry({ id: 'x', title: 'Reunion News' }),
        makeEntry({ id: 'y', title: 'Spring Festival' }),
      ]),
    )
    renderPage()
    await screen.findByText('Reunion News')
    const searchInput = screen.getByRole('searchbox')
    fireEvent.change(searchInput, { target: { value: 'Reunion' } })
    expect(screen.queryByText('Spring Festival')).toBeNull()
    expect(screen.getByText('Reunion News')).toBeDefined()
  })

  it('shows a confirmation dialog before deleting an entry', async () => {
    renderPage()
    await screen.findByText('Community Update')

    const deleteButtons = screen.getAllByRole('button', { name: /delete/i })
    fireEvent.click(deleteButtons[0])

    const dialog = await screen.findByRole('dialog')
    expect(within(dialog).getByRole('heading', { name: /delete/i })).toBeDefined()
    expect(within(dialog).getByRole('button', { name: /cancel/i })).toBeDefined()
  })
})
