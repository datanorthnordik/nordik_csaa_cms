import { fireEvent, render, screen, waitFor, within } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import { createAppStore } from '../store/store'
import type { GallerySummary } from '../types/media'
import { MediaLibraryRoute } from './MediaLibraryRoute'

const { toastSuccess, listGalleriesMock, deleteGalleryMock } = vi.hoisted(() => ({
  toastSuccess: vi.fn(),
  listGalleriesMock: vi.fn(),
  deleteGalleryMock: vi.fn(),
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
    getGallery: vi.fn(),
    fetchGalleryContent: vi.fn(),
    createGallery: vi.fn(),
    updateGallery: vi.fn(),
    deleteGallery: deleteGalleryMock,
    uploadGalleryImages: vi.fn(),
    updateGalleryImage: vi.fn(),
    reorderGalleryImages: vi.fn(),
    deleteGalleryImages: vi.fn(),
  },
}))

function renderRoute(initialEntry = '/media-library') {
  return render(
    <Provider store={createAppStore()}>
      <MemoryRouter initialEntries={[initialEntry]}>
        <Routes>
          <Route path="/media-library" element={<MediaLibraryRoute />} />
        </Routes>
      </MemoryRouter>
    </Provider>,
  )
}

describe('MediaLibraryRoute', () => {
  beforeEach(async () => {
    await i18n.changeLanguage('en')
    toastSuccess.mockReset()
    listGalleriesMock.mockReset()
    deleteGalleryMock.mockReset()

    const galleries: GallerySummary[] = [
      { id: 1, name: 'Annual Gala 2023', assetCount: 12, visibility: 'draft' },
    ]

    listGalleriesMock.mockResolvedValue(galleries)
    deleteGalleryMock.mockResolvedValue({ message: 'Gallery deleted successfully' })
  })

  it('calls the delete gallery api only after delete is confirmed', async () => {
    renderRoute()

    await screen.findByText(/annual gala 2023/i)

    fireEvent.click(screen.getByRole('button', { name: /delete gallery/i }))
    expect(deleteGalleryMock).not.toHaveBeenCalled()

    const dialog = screen.getByRole('dialog')
    fireEvent.click(within(dialog).getByRole('button', { name: /^delete gallery$/i }))

    await waitFor(() => {
      expect(deleteGalleryMock).toHaveBeenCalledWith(1)
    })
    expect(toastSuccess).toHaveBeenCalledWith('Gallery deleted successfully')
  })
})
