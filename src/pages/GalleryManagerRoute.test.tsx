import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import { createAppStore } from '../store/store'
import type { GalleryDetail } from '../types/media'
import { GalleryManagerRoute } from './GalleryManagerRoute'

const {
  toastSuccess,
  getGalleryMock,
  updateGalleryMock,
  listGalleriesMock,
  fetchGalleryContentMock,
} = vi.hoisted(() => ({
  toastSuccess: vi.fn(),
  getGalleryMock: vi.fn(),
  updateGalleryMock: vi.fn(),
  listGalleriesMock: vi.fn(),
  fetchGalleryContentMock: vi.fn(),
}))

vi.mock('react-hot-toast', () => ({
  default: {
    success: toastSuccess,
    error: vi.fn(),
  },
}))

vi.mock('../api/mediaApi', () => ({
  mediaApi: {
    listGalleries: listGalleriesMock,
    getGallery: getGalleryMock,
    fetchGalleryContent: fetchGalleryContentMock,
    createGallery: vi.fn(),
    updateGallery: updateGalleryMock,
    deleteGallery: vi.fn(),
    uploadGalleryImages: vi.fn(),
    updateGalleryImage: vi.fn(),
    reorderGalleryImages: vi.fn(),
    deleteGalleryImages: vi.fn(),
  },
}))

function renderRoute(initialEntry = '/media-library/1') {
  return render(
    <Provider store={createAppStore()}>
      <MemoryRouter initialEntries={[initialEntry]}>
        <Routes>
          <Route
            path="/media-library/:galleryId"
            element={<GalleryManagerRoute />}
          />
        </Routes>
      </MemoryRouter>
    </Provider>,
  )
}

describe('GalleryManagerRoute', () => {
  beforeEach(async () => {
    await i18n.changeLanguage('en')
    toastSuccess.mockReset()
    getGalleryMock.mockReset()
    updateGalleryMock.mockReset()
    listGalleriesMock.mockReset()
    fetchGalleryContentMock.mockReset()

    if (typeof URL.createObjectURL !== 'function') {
      Object.defineProperty(URL, 'createObjectURL', {
        configurable: true,
        writable: true,
        value: vi.fn(() => `blob:mock-${Math.random().toString(36).slice(2)}`),
      })
    }
    if (typeof URL.revokeObjectURL !== 'function') {
      Object.defineProperty(URL, 'revokeObjectURL', {
        configurable: true,
        writable: true,
        value: vi.fn(),
      })
    }

    const gallery: GalleryDetail = {
      id: 1,
      name: 'Summer Collection 2024',
      assetLimit: 20,
      visibility: 'draft',
      assets: [],
    }

    getGalleryMock.mockResolvedValue(gallery)
    updateGalleryMock.mockResolvedValue({
      message: 'Gallery updated successfully',
      gallery: {
        id: 1,
        name: 'Summer Collection 2024',
        published: true,
      },
    })
    listGalleriesMock.mockResolvedValue([])
    fetchGalleryContentMock.mockResolvedValue(
      new globalThis.Blob(['image'], { type: 'image/jpeg' }),
    )
  })

  it('shows a success toast when publishing a gallery', async () => {
    renderRoute()

    await screen.findByDisplayValue('Summer Collection 2024')
    fireEvent.click(screen.getByRole('button', { name: /publish gallery/i }))

    await waitFor(() => {
      expect(updateGalleryMock).toHaveBeenCalledTimes(1)
    })
    expect(updateGalleryMock).toHaveBeenCalledWith(1, {
      name: 'Summer Collection 2024',
      description: undefined,
      published: true,
      cover_image: undefined,
      remove_cover_image: false,
    })
    await waitFor(() => {
      expect(toastSuccess).toHaveBeenCalledWith(
        'Gallery published successfully.',
      )
    })
  })
})
