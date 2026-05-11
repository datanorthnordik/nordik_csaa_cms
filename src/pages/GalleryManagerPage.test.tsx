import { fireEvent, render, screen } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import { createAppStore } from '../store/store'
import type { GalleryDetail } from '../types/media'
import { GalleryManagerPage } from './GalleryManagerPage'

function renderPage(
  props: Parameters<typeof GalleryManagerPage>[0] = {},
  initialEntry = '/media-library/1',
) {
  const store = createAppStore()
  return render(
    <Provider store={store}>
      <MemoryRouter initialEntries={[initialEntry]}>
        <Routes>
          <Route
            path="/media-library/:galleryId"
            element={<GalleryManagerPage {...props} />}
          />
        </Routes>
      </MemoryRouter>
    </Provider>,
  )
}

beforeEach(async () => {
  await i18n.changeLanguage('en')
})

const baseGallery: GalleryDetail = {
  id: 1,
  name: 'Summer Collection 2024',
  assetLimit: 20,
  assets: [],
}

describe('GalleryManagerPage', () => {
  it('renders the not-found state when no gallery is provided', () => {
    renderPage()
    expect(screen.getByText(/gallery unavailable/i)).toBeDefined()
  })

  it('renders gallery name, breadcrumb, and asset counter', () => {
    renderPage({
      gallery: { ...baseGallery, assets: [{ id: 9, fileName: 'a.jpg', fileUrl: '/a.jpg' }] },
    })
    expect(screen.getByRole('heading', { name: /summer collection 2024/i })).toBeDefined()
    expect(screen.getByText(/1 \/ 20 assets in this gallery/i)).toBeDefined()
  })

  it('does not render description when it is missing', () => {
    renderPage({ gallery: baseGallery })
    expect(
      screen.queryByText(/this collection features high-resolution/i),
    ).toBeNull()
  })

  it('renders description only when present', () => {
    renderPage({
      gallery: { ...baseGallery, description: 'My gallery description' },
    })
    expect(screen.getByText(/my gallery description/i)).toBeDefined()
  })

  it('selecting an asset opens the SelectedImagePanel and conditionally renders blocks', () => {
    renderPage({
      gallery: {
        ...baseGallery,
        assets: [
          {
            id: 1,
            fileName: 'sunset_01.jpg',
            fileUrl: '/sunset.jpg',
            dimensions: { width: 4200, height: 2800 },
            altText: 'A serene sunset',
          },
        ],
      },
    })

    expect(screen.queryByText(/^selected image$/i)).toBeNull()
    fireEvent.click(screen.getByRole('button', { name: /a serene sunset/i }))
    expect(screen.getByRole('complementary', { name: /selected image/i })).toBeDefined()
    expect(screen.getByText(/dimensions/i)).toBeDefined()
    expect(screen.getByText(/4200 × 2800 px/i)).toBeDefined()
    expect(screen.queryByText(/usage tracking/i)).toBeNull()
  })

  it('renders usage tracking only when present', () => {
    renderPage({
      gallery: {
        ...baseGallery,
        assets: [
          {
            id: 2,
            fileName: 'banner.jpg',
            fileUrl: '/banner.jpg',
            usageTracking: [{ label: '/home', path: '/home' }],
          },
        ],
      },
    })

    fireEvent.click(screen.getByRole('button', { name: /banner\.jpg/i }))
    expect(screen.getByText(/usage tracking/i)).toBeDefined()
    expect(screen.getByRole('link', { name: '/home' })).toBeDefined()
  })

  it('disables upload submit until both files and alt-text are provided', () => {
    const onUploadAssets = vi.fn()
    renderPage({ gallery: baseGallery, onUploadAssets })

    const submit = screen.getByRole('button', { name: /^upload$/i })
    expect((submit as HTMLButtonElement).disabled).toBe(true)
  })
})
