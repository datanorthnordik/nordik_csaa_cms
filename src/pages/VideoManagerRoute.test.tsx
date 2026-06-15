import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import { createAppStore } from '../store/store'
import { VideoManagerRoute } from './VideoManagerRoute'

const {
  addVideoItemsMock,
  createVideoPackageMock,
  deleteVideoItemMock,
  deleteVideoPackageMock,
  fetchVideoTeaserContentMock,
  getVideoPackageMock,
  toastError,
  toastSuccess,
  updateVideoItemMock,
  updateVideoPackageMock,
} = vi.hoisted(() => ({
  addVideoItemsMock: vi.fn(),
  createVideoPackageMock: vi.fn(),
  deleteVideoItemMock: vi.fn(),
  deleteVideoPackageMock: vi.fn(),
  fetchVideoTeaserContentMock: vi.fn(),
  getVideoPackageMock: vi.fn(),
  toastError: vi.fn(),
  toastSuccess: vi.fn(),
  updateVideoItemMock: vi.fn(),
  updateVideoPackageMock: vi.fn(),
}))

vi.mock('react-hot-toast', () => ({
  default: {
    success: toastSuccess,
    error: toastError,
  },
}))

vi.mock('../api/videoApi', () => ({
  videoApi: {
    addVideoItems: addVideoItemsMock,
    createVideoPackage: createVideoPackageMock,
    deleteVideoItem: deleteVideoItemMock,
    deleteVideoPackage: deleteVideoPackageMock,
    fetchVideoTeaserContent: fetchVideoTeaserContentMock,
    getVideoPackage: getVideoPackageMock,
    listVideoPackages: vi.fn(),
    updateVideoItem: updateVideoItemMock,
    updateVideoPackage: updateVideoPackageMock,
  },
}))

function renderRoute(initialEntry = '/videos/new') {
  return render(
    <Provider store={createAppStore()}>
      <MemoryRouter initialEntries={[initialEntry]}>
        <Routes>
          <Route path="/videos/new" element={<VideoManagerRoute />} />
          <Route path="/videos/:videoId" element={<VideoManagerRoute />} />
        </Routes>
      </MemoryRouter>
    </Provider>,
  )
}

describe('VideoManagerRoute', () => {
  beforeEach(async () => {
    await i18n.changeLanguage('en')
    addVideoItemsMock.mockReset()
    createVideoPackageMock.mockReset()
    deleteVideoItemMock.mockReset()
    deleteVideoPackageMock.mockReset()
    fetchVideoTeaserContentMock.mockReset()
    getVideoPackageMock.mockReset()
    toastError.mockReset()
    toastSuccess.mockReset()
    updateVideoItemMock.mockReset()
    updateVideoPackageMock.mockReset()

    Object.defineProperty(URL, 'createObjectURL', {
      configurable: true,
      writable: true,
      value: vi.fn(() => 'blob:video-preview'),
    })
    Object.defineProperty(URL, 'revokeObjectURL', {
      configurable: true,
      writable: true,
      value: vi.fn(),
    })
  })

  it('shows the teaser preview immediately after a file is selected', async () => {
    const { container } = renderRoute()
    const teaserFile = new File(['teaser'], 'teaser.png', { type: 'image/png' })

    fireEvent.change(screen.getByLabelText(/drop teaser image here or browse/i), {
      target: { files: [teaserFile] },
    })

    await waitFor(() => {
      expect(screen.getByText(/teaser preview/i)).toBeTruthy()
    })

    expect(container.querySelector('img[src="blob:video-preview"]')).not.toBeNull()
    expect(createVideoPackageMock).not.toHaveBeenCalled()
  })

  it('renders a playable YouTube preview for a valid video link', async () => {
    renderRoute()

    fireEvent.change(screen.getByLabelText(/youtube link/i), {
      target: { value: 'https://www.youtube.com/watch?v=welcome123' },
    })

    const iframe = await screen.findByTitle(/video preview/i)
    expect(iframe.getAttribute('src')).toBe('https://www.youtube.com/embed/welcome123?rel=0')
  })
})
