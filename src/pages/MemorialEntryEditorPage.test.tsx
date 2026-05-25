import type { ReactNode } from 'react'
import { fireEvent, render, screen, waitFor, within } from '@testing-library/react'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { memorialApi } from '../api/memorialApi'
import { useMemorialEntries } from '../hooks/useMemorialEntries'
import i18n from '../i18n'
import type { MemorialEntry } from '../lib/memorialTypes'
import { MemorialEntryEditorPage } from './MemorialEntryEditorPage'

const {
  mockNavigate,
  toastSuccess,
  toastError,
} = vi.hoisted(() => ({
  mockNavigate: vi.fn(),
  toastSuccess: vi.fn(),
  toastError: vi.fn(),
}))

vi.mock('../hooks/useMemorialEntries', () => ({
  useMemorialEntries: vi.fn(),
}))

vi.mock('../api/memorialApi', () => ({
  memorialApi: {
    getMemorial: vi.fn(),
    getMemorialPortraitContent: vi.fn(),
    getMemorialGalleryImageContent: vi.fn(),
  },
}))

vi.mock('../components/CmsAppShell', () => ({
  CmsAppShell: ({ children }: { children?: ReactNode }) => (
    <div data-testid="cms-shell">{children}</div>
  ),
}))

vi.mock('../components/Loader', () => ({
  Loader: ({ label }: { label?: string }) => <div>{label ?? 'Loading'}</div>,
}))

vi.mock('../components/cms/RichTextEditor', () => ({
  RichTextEditor: ({
    value,
    onChange,
    placeholder,
  }: {
    value: string
    onChange: (value: string) => void
    placeholder?: string
  }) => (
    <textarea
      aria-label="Rich text editor"
      placeholder={placeholder}
      value={value}
      onChange={(event) => onChange(event.target.value)}
    />
  ),
}))

vi.mock('../components/media/UploadDropzone', () => ({
  UploadDropzone: ({
    label,
    multiple,
    onFiles,
  }: {
    label: string
    multiple?: boolean
    onFiles: (files: File[]) => void
  }) => (
    <button
      type="button"
      onClick={() =>
        onFiles(
          multiple
            ? [
                new File(['gallery-one'], 'fresh-one.png', { type: 'image/png' }),
                new File(['gallery-two'], 'fresh-two.webp', { type: 'image/webp' }),
              ]
            : [new File(['portrait'], 'portrait-upload.jpg', { type: 'image/jpeg' })],
        )
      }
    >
      {label}
    </button>
  ),
}))

vi.mock('../components/cms/EntryActions', () => ({
  EntryActions: ({
    onSaveDraft,
    onPublish,
    onDelete,
    isSubmitting,
    isDeleting,
  }: {
    onSaveDraft: () => void
    onPublish: () => void
    onDelete?: () => void
    isSubmitting?: boolean
    isDeleting?: boolean
  }) => (
    <div>
      <button type="button" disabled={isSubmitting} onClick={onSaveDraft}>
        Save Draft Action
      </button>
      <button type="button" disabled={isSubmitting} onClick={onPublish}>
        Publish Action
      </button>
      {onDelete ? (
        <button type="button" disabled={isDeleting} onClick={onDelete}>
          Delete Action
        </button>
      ) : null}
    </div>
  ),
}))

vi.mock('../components/cms/PublishingControls', () => ({
  PublishingControls: ({
    status,
    publishOn,
  }: {
    status: string
    publishOn?: string
  }) => (
    <div data-testid="publishing-controls">
      {status}
      {publishOn ? ` | ${publishOn}` : ''}
    </div>
  ),
}))

vi.mock('react-hot-toast', () => ({
  default: {
    success: toastSuccess,
    error: toastError,
  },
}))

vi.mock('react-router-dom', async () => {
  const actual =
    await vi.importActual<typeof import('react-router-dom')>('react-router-dom')

  return {
    ...actual,
    useNavigate: () => mockNavigate,
  }
})

const mockedUseMemorialEntries = vi.mocked(useMemorialEntries)
const mockedGetMemorial = vi.mocked(memorialApi.getMemorial)
const mockedGetPortraitContent = vi.mocked(memorialApi.getMemorialPortraitContent)
const mockedGetGalleryImageContent = vi.mocked(memorialApi.getMemorialGalleryImageContent)

const mockCreate = vi.fn()
const mockUpdate = vi.fn()
const mockRemove = vi.fn()

const sampleEntry: MemorialEntry = {
  id: '18',
  fullName: 'Grace Hopper',
  affiliation: 'US Navy',
  category: 'veteran',
  categoryLabel: 'Veteran',
  status: 'published',
  biography: '<p>Pioneer</p>',
  dateOfBirth: '1906-12-09',
  dateOfPassing: '1992-01-01',
  createdAt: '2026-05-18T00:00:00Z',
  updatedAt: '2026-05-19T00:00:00Z',
  publishedAt: '2026-05-20T00:00:00Z',
  portrait: {
    fileName: 'portrait.jpg',
    mimeType: 'image/jpeg',
    fileSize: 2048,
  },
  galleryImages: [
    {
      id: '91',
      fileName: 'gallery-one.png',
      mimeType: 'image/png',
      fileSize: 2048,
    },
    {
      id: '92',
      fileName: 'gallery-two.webp',
      mimeType: 'image/webp',
      fileSize: 512,
    },
  ],
}

function renderPage(initialEntry = '/memorial/new') {
  return render(
    <MemoryRouter initialEntries={[initialEntry]}>
      <Routes>
        <Route path="/memorial/new" element={<MemorialEntryEditorPage />} />
        <Route
          path="/memorial/:id/edit"
          element={<MemorialEntryEditorPage mode="edit" />}
        />
      </Routes>
    </MemoryRouter>,
  )
}

describe('MemorialEntryEditorPage', () => {
  beforeEach(async () => {
    vi.clearAllMocks()
    mockCreate.mockReset()
    mockUpdate.mockReset()
    mockRemove.mockReset()
    mockedUseMemorialEntries.mockReturnValue({
      create: mockCreate,
      update: mockUpdate,
      remove: mockRemove,
    } as never)
    await i18n.changeLanguage('en')

    if (!('createObjectURL' in URL)) {
      Object.defineProperty(URL, 'createObjectURL', {
        configurable: true,
        writable: true,
        value: vi.fn(),
      })
    }
    if (!('revokeObjectURL' in URL)) {
      Object.defineProperty(URL, 'revokeObjectURL', {
        configurable: true,
        writable: true,
        value: vi.fn(),
      })
    }

    let objectUrlIndex = 0
    vi.mocked(URL.createObjectURL).mockImplementation(
      () => `blob:memorial-${++objectUrlIndex}`,
    )
    vi.mocked(URL.revokeObjectURL).mockImplementation(() => undefined)
  })

  it('shows validation errors on create when required fields are missing', async () => {
    renderPage()

    fireEvent.click(screen.getByRole('button', { name: 'Save Draft Action' }))

    await waitFor(() => {
      expect(toastError).toHaveBeenCalledWith('Please fix the highlighted memorial fields.')
    })
    expect(screen.getAllByText('Full name is required.').length).toBeGreaterThan(0)
    expect(screen.getAllByText('Category is required.').length).toBeGreaterThan(0)
    expect(mockCreate).not.toHaveBeenCalled()
  })

  it('creates a memorial entry with uploaded portrait and gallery files, then navigates to edit mode', async () => {
    mockCreate.mockResolvedValue({
      ...sampleEntry,
      id: '77',
      fullName: 'Ada Lovelace',
      affiliation: 'Class of 1984',
      category: 'founder',
      categoryLabel: 'Founder',
      status: 'published',
      biography: 'A life remembered.',
      dateOfBirth: '1815-12-10',
      dateOfPassing: '1852-11-27',
      portrait: undefined,
      galleryImages: [],
    })

    const { container } = renderPage()

    fireEvent.change(screen.getByLabelText('Full Name'), {
      target: { value: ' Ada Lovelace ' },
    })
    fireEvent.change(screen.getByLabelText('Category'), {
      target: { value: 'founder' },
    })
    fireEvent.change(screen.getByLabelText('Affiliation / Class Year'), {
      target: { value: ' Class of 1984 ' },
    })
    fireEvent.change(screen.getByLabelText('Rich text editor'), {
      target: { value: 'A life remembered.' },
    })

    const [dateOfBirthInput, dateOfPassingInput] =
      container.querySelectorAll('input[type="date"]')
    fireEvent.change(dateOfBirthInput as HTMLInputElement, {
      target: { value: '1815-12-10' },
    })
    fireEvent.change(dateOfPassingInput as HTMLInputElement, {
      target: { value: '1852-11-27' },
    })

    fireEvent.click(screen.getByRole('button', { name: 'Featured Portrait' }))
    fireEvent.click(screen.getByRole('button', { name: 'Add Gallery Photos' }))

    await screen.findByText('fresh-one.png')
    await screen.findByText('fresh-two.webp')
    expect(screen.getByText('2 images')).toBeDefined()

    fireEvent.click(screen.getByRole('button', { name: 'Publish Action' }))

    await waitFor(() => {
      expect(mockCreate).toHaveBeenCalled()
    })

    const [input, portraitFile, galleryFiles] = mockCreate.mock.calls[0]
    expect(input).toEqual({
      full_name: 'Ada Lovelace',
      affiliation: 'Class of 1984',
      category: 'founder',
      status: 'published',
      biography: 'A life remembered.',
      date_of_birth: '1815-12-10',
      date_of_passing: '1852-11-27',
      remove_portrait: false,
      remove_gallery_image_ids: [],
    })
    expect((portraitFile as File).name).toBe('portrait-upload.jpg')
    expect((galleryFiles as File[]).map((file) => file.name)).toEqual([
      'fresh-one.png',
      'fresh-two.webp',
    ])
    expect(toastSuccess).toHaveBeenCalledWith('Memorial entry created.')
    expect(mockNavigate).toHaveBeenCalledWith('/memorial/77/edit', {
      replace: true,
    })
  })

  it('loads an existing entry, supports media removal, and saves updates', async () => {
    mockedGetMemorial.mockResolvedValue(sampleEntry)
    mockedGetPortraitContent.mockResolvedValue(
      new Blob(['portrait'], { type: 'image/jpeg' }),
    )
    mockedGetGalleryImageContent
      .mockResolvedValueOnce(new Blob(['gallery-one'], { type: 'image/png' }))
      .mockRejectedValueOnce(new Error('preview unavailable'))
    mockUpdate.mockResolvedValue({
      ...sampleEntry,
      fullName: 'Grace Hopper Updated',
      status: 'review',
      portrait: undefined,
      galleryImages: [],
    })

    renderPage('/memorial/18/edit')

    await waitFor(() => {
      expect(screen.getByDisplayValue('Grace Hopper')).toBeDefined()
    })

    expect(mockedGetPortraitContent).toHaveBeenCalledWith('18')
    expect(mockedGetGalleryImageContent).toHaveBeenCalledWith('18', '91')
    expect(mockedGetGalleryImageContent).toHaveBeenCalledWith('18', '92')
    expect(screen.getByText('gallery-one.png')).toBeDefined()
    expect(screen.getByText('gallery-two.webp')).toBeDefined()
    expect(screen.getByText('PNG | 2.0 KB')).toBeDefined()
    expect(screen.getByTestId('publishing-controls').textContent).toContain(
      'published',
    )

    fireEvent.click(screen.getByRole('button', { name: 'Remove portrait' }))
    await waitFor(() => {
      expect(screen.queryByRole('button', { name: 'Remove portrait' })).toBeNull()
    })

    const firstExistingCard = screen.getByText('gallery-one.png').closest('article')
    if (!firstExistingCard) {
      throw new Error('Expected gallery card for gallery-one.png')
    }
    fireEvent.click(
      within(firstExistingCard).getByRole('button', { name: 'Remove image' }),
    )
    await waitFor(() => {
      expect(screen.queryByText('gallery-one.png')).toBeNull()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Add Gallery Photos' }))
    await screen.findByText('fresh-one.png')
    await screen.findByText('fresh-two.webp')

    const secondNewCard = screen.getByText('fresh-two.webp').closest('article')
    if (!secondNewCard) {
      throw new Error('Expected gallery card for fresh-two.webp')
    }
    fireEvent.click(
      within(secondNewCard).getByRole('button', { name: 'Remove image' }),
    )
    await waitFor(() => {
      expect(screen.queryByText('fresh-two.webp')).toBeNull()
    })

    fireEvent.change(screen.getByLabelText('Full Name'), {
      target: { value: 'Grace Hopper Updated' },
    })
    fireEvent.change(screen.getByLabelText('Status'), {
      target: { value: 'review' },
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save Draft Action' }))

    await waitFor(() => {
      expect(mockUpdate).toHaveBeenCalled()
    })

    const [id, input, portraitFile, galleryFiles] = mockUpdate.mock.calls[0]
    expect(id).toBe('18')
    expect(input).toEqual({
      full_name: 'Grace Hopper Updated',
      affiliation: 'US Navy',
      category: 'veteran',
      status: 'review',
      biography: '<p>Pioneer</p>',
      date_of_birth: '1906-12-09',
      date_of_passing: '1992-01-01',
      remove_portrait: true,
      remove_gallery_image_ids: [91],
    })
    expect(portraitFile).toBeNull()
    expect((galleryFiles as File[]).map((file) => file.name)).toEqual(['fresh-one.png'])
    expect(toastSuccess).toHaveBeenCalledWith('Memorial entry updated.')
    await waitFor(() => {
      expect(screen.getByText('No gallery photos yet')).toBeDefined()
    })
  })

  it('deletes an existing entry and navigates back to the memorial list', async () => {
    mockedGetMemorial.mockResolvedValue(sampleEntry)
    mockedGetPortraitContent.mockResolvedValue(
      new Blob(['portrait'], { type: 'image/jpeg' }),
    )
    mockedGetGalleryImageContent.mockResolvedValue(
      new Blob(['gallery'], { type: 'image/png' }),
    )
    mockRemove.mockResolvedValue(undefined)

    renderPage('/memorial/18/edit')

    await waitFor(() => {
      expect(screen.getByDisplayValue('Grace Hopper')).toBeDefined()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Delete Action' }))

    await waitFor(() => {
      expect(mockRemove).toHaveBeenCalledWith('18')
    })
    expect(toastSuccess).toHaveBeenCalledWith('Memorial entry deleted.')
    expect(mockNavigate).toHaveBeenCalledWith('/memorial', { replace: true })
  })

  it('shows a toast when loading an existing entry fails', async () => {
    mockedGetMemorial.mockRejectedValue(new Error('Backend unavailable'))

    renderPage('/memorial/501/edit')

    await waitFor(() => {
      expect(toastError).toHaveBeenCalledWith('Backend unavailable')
    })
    expect(screen.getByRole('heading', { name: 'Edit Memorial Entry' })).toBeDefined()
  })
})
