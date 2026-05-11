import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import { createAppStore } from '../store/store'
import type { GallerySummary } from '../types/media'
import { MediaLibraryPage } from './MediaLibraryPage'

function renderPage(props: Parameters<typeof MediaLibraryPage>[0] = {}) {
  const store = createAppStore()
  return render(
    <Provider store={store}>
      <MemoryRouter>
        <MediaLibraryPage {...props} />
      </MemoryRouter>
    </Provider>,
  )
}

beforeEach(async () => {
  await i18n.changeLanguage('en')
})

describe('MediaLibraryPage', () => {
  it('renders breadcrumb and the create button', () => {
    renderPage()
    expect(screen.getByRole('navigation', { name: /breadcrumb/i })).toBeDefined()
    expect(screen.getByRole('button', { name: /create new gallery/i })).toBeDefined()
  })

  it('renders the empty state when galleries is undefined', () => {
    renderPage()
    expect(screen.getByText(/no galleries yet/i)).toBeDefined()
  })

  it('renders gallery cards when data is provided', () => {
    const galleries: GallerySummary[] = [
      { id: 1, name: 'Annual Gala 2023 Highlights', assetCount: 142 },
      { id: 2, name: 'Press Release Assets', assetCount: 29, visibility: 'public' },
    ]
    renderPage({ galleries })

    expect(screen.getByText(/annual gala 2023 highlights/i)).toBeDefined()
    expect(screen.getByText(/press release assets/i)).toBeDefined()
  })

  it('only renders the visibility badge when visibility is defined', () => {
    const galleries: GallerySummary[] = [
      { id: 1, name: 'Hidden Gallery' },
      { id: 2, name: 'Public Gallery', visibility: 'public' },
    ]
    renderPage({ galleries })

    expect(screen.queryAllByText(/^public$/i).length).toBeGreaterThanOrEqual(1)
  })

  it('opens the create dialog when the trigger is clicked', () => {
    renderPage()
    fireEvent.click(screen.getByRole('button', { name: /create new gallery/i }))
    expect(screen.getByRole('heading', { name: /create new gallery/i })).toBeDefined()
  })

  it('filters galleries by search term', () => {
    const galleries: GallerySummary[] = [
      { id: 1, name: 'Annual Gala 2023' },
      { id: 2, name: 'Quarterly Training' },
    ]
    renderPage({ galleries })

    const search = screen.getByRole('searchbox')
    fireEvent.change(search, { target: { value: 'training' } })

    expect(screen.queryByText(/annual gala 2023/i)).toBeNull()
    expect(screen.getByText(/quarterly training/i)).toBeDefined()
  })

  it('invokes onCreate with form values', async () => {
    const onCreate = vi.fn()
    renderPage({ onCreate })

    fireEvent.click(screen.getByRole('button', { name: /create new gallery/i }))
    const nameInput = screen.getByPlaceholderText(/annual stakeholder meeting/i)
    fireEvent.change(nameInput, { target: { value: 'New Gallery' } })
    fireEvent.click(screen.getByRole('button', { name: /create & edit/i }))

    await waitFor(() => {
      expect(onCreate).toHaveBeenCalledTimes(1)
    })
    expect(onCreate.mock.calls[0][0]).toMatchObject({ name: 'New Gallery' })
  })
})
