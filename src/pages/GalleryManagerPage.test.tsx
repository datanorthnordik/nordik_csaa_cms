import { fireEvent, render, screen } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import { GALLERY_IMAGE_UPLOAD_MAX_FILE_SIZE_BYTES } from '../lib/resourceUpload'
import { createAppStore } from '../store/store'
import type { GalleryDetail } from '../types/media'
import { GalleryManagerPage } from './GalleryManagerPage'

const { toastError } = vi.hoisted(() => ({
  toastError: vi.fn(),
}))

vi.mock('react-hot-toast', () => ({
  default: {
    error: toastError,
  },
}))

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
  toastError.mockReset()
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
})

const baseGallery: GalleryDetail = {
  id: 1,
  name: 'Summer Collection 2024',
  assetLimit: 20,
  assets: [],
}

function makeAsset(id: number, fileName: string) {
  return { id, fileName, fileUrl: `/${fileName}` }
}

function fileFromName(name: string) {
  return new File(['x'], name, { type: 'image/jpeg' })
}

function oversizedImageFromName(name: string) {
  const file = new File(['x'], name, { type: 'image/jpeg' })
  Object.defineProperty(file, 'size', {
    configurable: true,
    value: GALLERY_IMAGE_UPLOAD_MAX_FILE_SIZE_BYTES + 1,
  })
  return file
}

describe('GalleryManagerPage', () => {
  it('renders the not-found state when no gallery is provided', () => {
    renderPage()
    expect(screen.getByText(/gallery unavailable/i)).toBeDefined()
  })

  it('renders gallery name as an inline-editable input and asset counter', () => {
    renderPage({
      gallery: { ...baseGallery, assets: [makeAsset(9, 'a.jpg')] },
    })
    expect(screen.getByDisplayValue('Summer Collection 2024')).toBeDefined()
    expect(screen.getByText(/1 \/ 20 assets in this gallery/i)).toBeDefined()
  })

  it('renders an empty description textarea when description is missing', () => {
    renderPage({ gallery: baseGallery })
    const description = screen.getByRole('textbox', {
      name: /describe the gallery/i,
    }) as HTMLTextAreaElement
    expect(description.value).toBe('')
  })

  it('commits name changes via onUpdateGallery on blur', () => {
    const onUpdateGallery = vi.fn()
    renderPage({ gallery: baseGallery, onUpdateGallery })

    const nameInput = screen.getByDisplayValue('Summer Collection 2024')
    fireEvent.change(nameInput, { target: { value: 'New Name' } })
    fireEvent.blur(nameInput)

    expect(onUpdateGallery).toHaveBeenCalledWith({ name: 'New Name' })
  })

  it('does not call onUpdateGallery when name is unchanged on blur', () => {
    const onUpdateGallery = vi.fn()
    renderPage({ gallery: baseGallery, onUpdateGallery })

    const nameInput = screen.getByDisplayValue('Summer Collection 2024')
    fireEvent.blur(nameInput)

    expect(onUpdateGallery).not.toHaveBeenCalled()
  })

  it('commits description changes via onUpdateGallery on blur', () => {
    const onUpdateGallery = vi.fn()
    renderPage({ gallery: baseGallery, onUpdateGallery })

    const descInput = screen.getByRole('textbox', {
      name: /describe the gallery/i,
    })
    fireEvent.change(descInput, { target: { value: 'A summer mood board.' } })
    fireEvent.blur(descInput)

    expect(onUpdateGallery).toHaveBeenCalledWith({
      description: 'A summer mood board.',
    })
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
            title: 'Sunset Hero',
            details: 'A serene sunset',
          },
        ],
      },
    })

    expect(screen.queryByText(/^selected image$/i)).toBeNull()
    fireEvent.click(screen.getByRole('button', { name: /sunset hero/i }))
    expect(
      screen.getByRole('complementary', { name: /selected image/i }),
    ).toBeDefined()
    expect(screen.getByText(/dimensions/i)).toBeDefined()
    expect(screen.getByText(/4200/i)).toBeDefined()
    expect(screen.getByText(/2800 px/i)).toBeDefined()
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
            title: 'Homepage Banner',
            usageTracking: [{ label: '/home', path: '/home' }],
          },
        ],
      },
    })

    fireEvent.click(screen.getByRole('button', { name: /homepage banner/i }))
    expect(screen.getByText(/usage tracking/i)).toBeDefined()
    expect(screen.getByRole('link', { name: '/home' })).toBeDefined()
  })

  it('disables upload submit when no files are pending', () => {
    renderPage({ gallery: baseGallery, onUploadAssets: vi.fn() })

    const submit = screen.getByRole('button', { name: /^upload$/i })
    expect((submit as HTMLButtonElement).disabled).toBe(true)
  })

  it('renders title, details, and optional link fields for each pending upload', () => {
    renderPage({ gallery: baseGallery, onUploadAssets: vi.fn() })

    const fileInputs = document.querySelectorAll('input[type="file"]')
    const dropInput = fileInputs[fileInputs.length - 1] as HTMLInputElement
    expect(dropInput).toBeTruthy()
    fireEvent.change(dropInput, {
      target: { files: [fileFromName('a.jpg'), fileFromName('b.jpg')] },
    })

    expect(
      screen.getByRole('textbox', { name: /image title for a\.jpg/i }),
    ).toBeDefined()
    expect(
      screen.getByRole('textbox', { name: /image details for a\.jpg/i }),
    ).toBeDefined()
    expect(
      screen.getByRole('textbox', { name: /click-through link for a\.jpg/i }),
    ).toBeDefined()
    expect(
      screen.getByRole('textbox', { name: /image title for b\.jpg/i }),
    ).toBeDefined()
    expect(
      screen.getByRole('textbox', { name: /image details for b\.jpg/i }),
    ).toBeDefined()
    expect(
      screen.getByRole('textbox', { name: /click-through link for b\.jpg/i }),
    ).toBeDefined()
  })

  it('rejects oversized gallery uploads before they are queued', async () => {
    renderPage({ gallery: baseGallery, onUploadAssets: vi.fn() })

    const fileInputs = document.querySelectorAll('input[type="file"]')
    const dropInput = fileInputs[fileInputs.length - 1] as HTMLInputElement
    fireEvent.change(dropInput, {
      target: { files: [oversizedImageFromName('too-large.jpg')] },
    })

    expect(await screen.findByText('This file exceeds the 5MB limit.')).toBeDefined()
    expect(toastError).toHaveBeenCalledWith('This file exceeds the 5MB limit.')
    expect(screen.queryByText('too-large.jpg')).toBeNull()
  })

  it('rejects non-image gallery uploads before they are queued', async () => {
    renderPage({ gallery: baseGallery, onUploadAssets: vi.fn() })

    const fileInputs = document.querySelectorAll('input[type="file"]')
    const dropInput = fileInputs[fileInputs.length - 1] as HTMLInputElement
    fireEvent.change(dropInput, {
      target: {
        files: [new File(['pdf'], 'guide.pdf', { type: 'application/pdf' })],
      },
    })

    expect(
      await screen.findByText('Only SVG, PNG, JPG, and WEBP are supported.'),
    ).toBeDefined()
    expect(toastError).toHaveBeenCalledWith(
      'Only SVG, PNG, JPG, and WEBP are supported.',
    )
    expect(screen.queryByText('guide.pdf')).toBeNull()
  })

  it('only enables submit when every pending upload has title and details', () => {
    const onUploadAssets = vi.fn()
    renderPage({ gallery: baseGallery, onUploadAssets })

    const fileInputs = document.querySelectorAll('input[type="file"]')
    const dropInput = fileInputs[fileInputs.length - 1] as HTMLInputElement
    fireEvent.change(dropInput, {
      target: { files: [fileFromName('a.jpg'), fileFromName('b.jpg')] },
    })

    const submit = screen.getByRole('button', { name: /^upload$/i })
    expect((submit as HTMLButtonElement).disabled).toBe(true)

    const firstTitle = screen.getByRole('textbox', {
      name: /image title for a\.jpg/i,
    }) as HTMLInputElement
    expect(firstTitle.value).toBe('a')

    fireEvent.change(firstTitle, { target: { value: '' } })
    expect((submit as HTMLButtonElement).disabled).toBe(true)

    fireEvent.change(firstTitle, { target: { value: 'Summer Hero' } })
    expect((submit as HTMLButtonElement).disabled).toBe(true)

    fireEvent.change(
      screen.getByRole('textbox', { name: /image details for a\.jpg/i }),
      { target: { value: 'first' } },
    )
    expect((submit as HTMLButtonElement).disabled).toBe(true)

    fireEvent.change(
      screen.getByRole('textbox', { name: /click-through link for a\.jpg/i }),
      { target: { value: 'https://partner.example.com' } },
    )

    fireEvent.change(
      screen.getByRole('textbox', { name: /image details for b\.jpg/i }),
      { target: { value: 'second' } },
    )
    expect((submit as HTMLButtonElement).disabled).toBe(false)

    fireEvent.click(submit)
    expect(onUploadAssets).toHaveBeenCalledTimes(1)
    const arg = onUploadAssets.mock.calls[0][0]
    expect(arg).toHaveLength(2)
    expect(arg[0].title).toBe('Summer Hero')
    expect(arg[0].details).toBe('first')
    expect(arg[0].linkUrl).toBe('https://partner.example.com')
    expect(arg[1].title).toBe('b')
    expect(arg[1].details).toBe('second')
    expect(arg[1].linkUrl).toBeUndefined()
  })

  it('enforces the asset limit by clamping dropped files', () => {
    const fullAssets = Array.from({ length: 19 }, (_, index) =>
      makeAsset(index + 1, `asset_${index + 1}.jpg`),
    )
    renderPage({
      gallery: { ...baseGallery, assets: fullAssets },
      onUploadAssets: vi.fn(),
    })

    const fileInputs = document.querySelectorAll('input[type="file"]')
    const dropInput = fileInputs[fileInputs.length - 1] as HTMLInputElement
    fireEvent.change(dropInput, {
      target: {
        files: [
          fileFromName('a.jpg'),
          fileFromName('b.jpg'),
          fileFromName('c.jpg'),
        ],
      },
    })

    expect(
      screen.getAllByRole('textbox', { name: /image title for/i }),
    ).toHaveLength(1)
    expect(
      screen.getAllByRole('textbox', { name: /image details for/i }),
    ).toHaveLength(1)
  })

  it('disables the dropzone input when the gallery is already full', () => {
    const fullAssets = Array.from({ length: 20 }, (_, index) =>
      makeAsset(index + 1, `asset_${index + 1}.jpg`),
    )
    renderPage({
      gallery: { ...baseGallery, assets: fullAssets },
      onUploadAssets: vi.fn(),
    })

    expect(screen.getByRole('alert')).toBeDefined()
    const fileInputs = document.querySelectorAll('input[type="file"]')
    const dropInput = fileInputs[fileInputs.length - 1] as HTMLInputElement
    expect(dropInput.disabled).toBe(true)
  })

  it('renders status at the top and the upload limit as helper text', () => {
    renderPage({
      gallery: { ...baseGallery, visibility: 'published' },
    })

    expect(screen.getByText(/^published$/i)).toBeDefined()
    expect(screen.getByText(/maximum 20 images per gallery/i)).toBeDefined()
    expect(screen.queryByRole('heading', { name: /^actions$/i })).toBeNull()
    expect(screen.queryByText(/max 20 uploads/i)).toBeNull()
  })

  it('does not render a Save changes button', () => {
    renderPage({
      gallery: baseGallery,
      onSaveDraft: vi.fn(),
      onPublish: vi.fn(),
    })

    expect(screen.queryByRole('button', { name: /save changes/i })).toBeNull()
  })

  it('invokes Save draft and Publish gallery', () => {
    const onSaveDraft = vi.fn()
    const onPublish = vi.fn()
    renderPage({
      gallery: baseGallery,
      onSaveDraft,
      onPublish,
    })

    fireEvent.click(screen.getByRole('button', { name: /save draft/i }))
    fireEvent.click(screen.getByRole('button', { name: /publish gallery/i }))

    expect(onSaveDraft).toHaveBeenCalledTimes(1)
    expect(onPublish).toHaveBeenCalledTimes(1)
  })

  it('renders move/delete overlay buttons only when handlers are provided', () => {
    const onMoveAsset = vi.fn()
    const onDeleteAsset = vi.fn()
    renderPage({
      gallery: {
        ...baseGallery,
        assets: [
          makeAsset(1, 'a.jpg'),
          makeAsset(2, 'b.jpg'),
          makeAsset(3, 'c.jpg'),
        ],
      },
      onMoveAsset,
      onDeleteAsset,
    })

    expect(
      screen.getAllByRole('button', { name: /move asset earlier/i }).length,
    ).toBe(2)
    expect(
      screen.getAllByRole('button', { name: /move asset later/i }).length,
    ).toBe(2)
    expect(
      screen.getAllByRole('button', { name: /delete asset/i }).length,
    ).toBe(3)

    fireEvent.click(
      screen.getAllByRole('button', { name: /move asset later/i })[0],
    )
    expect(onMoveAsset).toHaveBeenCalledWith(
      expect.objectContaining({ id: 1 }),
      1,
    )
  })

  it('renders the Cover & Description card with an upload dropzone when no cover is set', () => {
    renderPage({ gallery: baseGallery, onSetCover: vi.fn() })

    expect(
      screen.getByRole('heading', { name: /cover & description/i }),
    ).toBeDefined()
    expect(screen.getByText(/drop cover image here or browse/i)).toBeDefined()
  })

  it('renders Replace/Remove buttons when a cover image is set', () => {
    const onSetCover = vi.fn()
    renderPage({
      gallery: {
        ...baseGallery,
        coverImage: {
          id: 99,
          fileName: 'cover.jpg',
          fileUrl: '/cover.jpg',
        },
      },
      onSetCover,
    })

    expect(screen.getByRole('button', { name: /remove cover/i })).toBeDefined()
    expect(screen.getByText(/replace cover/i)).toBeDefined()
    fireEvent.click(screen.getByRole('button', { name: /remove cover/i }))
    expect(onSetCover).toHaveBeenCalledWith(null)
  })

  it('invokes onSetCover with a file when dropzone receives one', () => {
    const onSetCover = vi.fn()
    renderPage({ gallery: baseGallery, onSetCover })

    const fileInputs = document.querySelectorAll('input[type="file"]')
    const coverInput = fileInputs[0] as HTMLInputElement
    const file = fileFromName('new-cover.jpg')
    fireEvent.change(coverInput, { target: { files: [file] } })

    expect(onSetCover).toHaveBeenCalledTimes(1)
    expect(onSetCover.mock.calls[0][0]).toBeInstanceOf(File)
    expect((onSetCover.mock.calls[0][0] as File).name).toBe('new-cover.jpg')
  })

  it('rejects invalid cover uploads before invoking onSetCover', async () => {
    const onSetCover = vi.fn()
    renderPage({ gallery: baseGallery, onSetCover })

    const fileInputs = document.querySelectorAll('input[type="file"]')
    const coverInput = fileInputs[0] as HTMLInputElement
    fireEvent.change(coverInput, {
      target: {
        files: [new File(['pdf'], 'cover.pdf', { type: 'application/pdf' })],
      },
    })

    expect(
      await screen.findByText('Only SVG, PNG, JPG, and WEBP are supported.'),
    ).toBeDefined()
    expect(onSetCover).not.toHaveBeenCalled()
    expect(toastError).toHaveBeenCalledWith(
      'Only SVG, PNG, JPG, and WEBP are supported.',
    )
  })

  it('lets users edit an existing asset title, details, and click-through link', () => {
    const onUpdateAsset = vi.fn()
    renderPage({
      gallery: {
        ...baseGallery,
        assets: [
          {
            id: 7,
            fileName: 'sunset_01.jpg',
            fileUrl: '/sunset.jpg',
            title: 'Sunset Hero',
            details: 'Warm evening light over the lake.',
            linkUrl: 'https://old.example.com',
          },
        ],
      },
      onUpdateAsset,
    })

    fireEvent.click(screen.getByRole('button', { name: /sunset hero/i }))
    fireEvent.click(screen.getByRole('button', { name: /^edit$/i }))

    fireEvent.change(screen.getByRole('textbox', { name: /^title$/i }), {
      target: { value: 'Sunset Story' },
    })
    fireEvent.change(screen.getByRole('textbox', { name: /^details$/i }), {
      target: { value: 'Updated description for the hero image.' },
    })
    fireEvent.change(screen.getByRole('textbox', { name: /click-through link/i }), {
      target: { value: 'https://partner.example.com' },
    })
    fireEvent.click(screen.getByRole('button', { name: /save changes/i }))

    expect(onUpdateAsset).toHaveBeenCalledWith(
      expect.objectContaining({ id: 7 }),
      {
        title: 'Sunset Story',
        details: 'Updated description for the hero image.',
        linkUrl: 'https://partner.example.com',
      },
    )
  })

  it('makes thumbnails draggable only when onReorderAsset is provided', () => {
    const { rerender } = renderPage({
      gallery: {
        ...baseGallery,
        assets: [makeAsset(1, 'a.jpg'), makeAsset(2, 'b.jpg')],
      },
    })

    let selectButtons = screen.getAllByRole('button', {
      name: /^(a|b)$/i,
    })
    expect(selectButtons.length).toBe(2)
    expect(
      (selectButtons[0].parentElement as HTMLElement).getAttribute('draggable'),
    ).toBe('false')

    rerender(
      <Provider store={createAppStore()}>
        <MemoryRouter initialEntries={['/media-library/1']}>
          <Routes>
            <Route
              path="/media-library/:galleryId"
              element={
                <GalleryManagerPage
                  gallery={{
                    ...baseGallery,
                    assets: [makeAsset(1, 'a.jpg'), makeAsset(2, 'b.jpg')],
                  }}
                  onReorderAsset={vi.fn()}
                />
              }
            />
          </Routes>
        </MemoryRouter>
      </Provider>,
    )

    selectButtons = screen.getAllByRole('button', {
      name: /^(a|b)$/i,
    })
    expect(
      (selectButtons[0].parentElement as HTMLElement).getAttribute('draggable'),
    ).toBe('true')
  })

  it('invokes onReorderAsset on drop', () => {
    const onReorderAsset = vi.fn()
    renderPage({
      gallery: {
        ...baseGallery,
        assets: [
          makeAsset(1, 'a.jpg'),
          makeAsset(2, 'b.jpg'),
          makeAsset(3, 'c.jpg'),
        ],
      },
      onReorderAsset,
    })

    const thumbs = screen
      .getAllByRole('button', { name: /^(a|b|c)$/i })
      .map((btn) => btn.parentElement as HTMLElement)

    const dataTransfer = {
      data: new Map<string, string>(),
      effectAllowed: '',
      dropEffect: '',
      setData(format: string, value: string) {
        this.data.set(format, value)
      },
      getData(format: string) {
        return this.data.get(format) ?? ''
      },
    }

    fireEvent.dragStart(thumbs[0], { dataTransfer })
    fireEvent.dragEnter(thumbs[2], { dataTransfer })
    fireEvent.dragOver(thumbs[2], { dataTransfer })
    fireEvent.drop(thumbs[2], { dataTransfer })

    expect(onReorderAsset).toHaveBeenCalledTimes(1)
    expect(onReorderAsset.mock.calls[0][0]).toMatchObject({ id: 1 })
    expect(onReorderAsset.mock.calls[0][1]).toMatchObject({ id: 3 })
  })
})
