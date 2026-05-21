import type { ReactNode } from 'react'
import axios from 'axios'
import type { AxiosError } from 'axios'
import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import {
  addNewsletterMedia,
  deleteNewsletterMedia,
  fetchNewsletterEntry,
  getNewsletterMediaContent,
} from '../api/newslettersApi'
import i18n from '../i18n'
import type { NewsletterEntry } from '../lib/newsletterTypes'
import {
  NewsletterEditorPage,
  buildDocumentMeta,
  canPreviewDocument,
  entryToFormState,
  formatFileSize,
  normalizeDateInputValue,
  resolveDocumentTypeLabel,
  stripFileExtension,
  validate,
} from './NewsletterEditorPage'

const {
  mockNavigate,
  mockCreate,
  mockUpdate,
  mockRemove,
  toastSuccess,
  toastError,
} = vi.hoisted(() => ({
  mockNavigate: vi.fn(),
  mockCreate: vi.fn(),
  mockUpdate: vi.fn(),
  mockRemove: vi.fn(),
  toastSuccess: vi.fn(),
  toastError: vi.fn(),
}))

vi.mock('../hooks/useNewsletterEntries', () => ({
  useNewsletterEntries: () => ({
    create: mockCreate,
    update: mockUpdate,
    remove: mockRemove,
  }),
}))

vi.mock('../api/newslettersApi', () => ({
  addNewsletterMedia: vi.fn(),
  deleteNewsletterMedia: vi.fn(),
  fetchNewsletterEntry: vi.fn(),
  getNewsletterMediaContent: vi.fn(),
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
    onFiles,
  }: {
    onFiles: (files: File[]) => void
  }) => (
    <button
      type="button"
      onClick={() =>
        onFiles([
          new File(['newsletter'], 'newsletter-preview.pdf', {
            type: 'application/pdf',
          }),
        ])
      }
    >
      Mock upload file
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
  PublishingControls: ({ status }: { status: string }) => (
    <div data-testid="publishing-controls">{status}</div>
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

const mockedFetchNewsletterEntry = vi.mocked(fetchNewsletterEntry)
const mockedAddNewsletterMedia = vi.mocked(addNewsletterMedia)
const mockedDeleteNewsletterMedia = vi.mocked(deleteNewsletterMedia)
const mockedGetNewsletterMediaContent = vi.mocked(getNewsletterMediaContent)

const sampleEntry: NewsletterEntry = {
  id: '18',
  title: 'July Digest',
  category: 'csaa',
  sendDate: '2026-07-10T00:00:00Z',
  contentHtml: '<p>Server content</p>',
  status: 'published',
  visibility: 'public',
  publishAt: null,
  media: [
    {
      id: '91',
      fileName: 'digest.pdf',
      mimeType: 'application/pdf',
      fileSize: 4096,
    },
  ],
  createdAt: '2026-07-01T00:00:00Z',
  updatedAt: '2026-07-02T00:00:00Z',
}

function renderPage(initialEntry = '/newsletters/new') {
  return render(
    <MemoryRouter initialEntries={[initialEntry]}>
      <Routes>
        <Route path="/newsletters/new" element={<NewsletterEditorPage />} />
        <Route
          path="/newsletters/:id/edit"
          element={<NewsletterEditorPage mode="edit" />}
        />
      </Routes>
    </MemoryRouter>,
  )
}

describe('NewsletterEditorPage helpers', () => {
  it('normalizes dates and maps loaded entries into form state', () => {
    const formState = entryToFormState(sampleEntry)

    expect(normalizeDateInputValue('2026-07-10T15:00:00Z')).toBe('2026-07-10')
    expect(normalizeDateInputValue('')).toBe('')
    expect(formState).toMatchObject({
      title: 'July Digest',
      category: 'csaa',
      sendDate: '2026-07-10',
      media: sampleEntry.media,
    })
  })

  it('validates required fields for save actions', () => {
    const t = (key: string) =>
      ({
        'newsletters.editor.validation.titleRequired': 'Title is required.',
        'newsletters.editor.validation.sendDateRequired': 'Send date is required.',
      })[key] ?? key

    expect(
      validate(
        {
          title: '  ',
          category: '',
          sendDate: '',
          contentHtml: '',
          visibility: 'public',
          publishAt: '',
          media: [],
        },
        t,
      ),
    ).toEqual({
      title: 'Title is required.',
      sendDate: 'Send date is required.',
    })
  })

  it('formats document metadata and detects previewable file types', () => {
    const t = (key: string) =>
      key === 'newsletters.editor.media.pendingUploadShort' ? 'Pending upload' : key

    expect(stripFileExtension('digest.final.pdf')).toBe('digest.final')
    expect(formatFileSize(512)).toBe('512 B')
    expect(formatFileSize(1536)).toBe('1.5 KB')
    expect(formatFileSize(2 * 1024 * 1024)).toBe('2.0 MB')
    expect(resolveDocumentTypeLabel('', 'newsletter.docx')).toBe('DOCX')
    expect(resolveDocumentTypeLabel('application/pdf', '')).toBe('PDF')
    expect(resolveDocumentTypeLabel('text/plain', '')).toBe('FILE')
    expect(
      buildDocumentMeta({
        isPendingUpload: true,
        fileTypeLabel: 'PDF',
        fileSize: 2048,
        t,
      }),
    ).toBe('Pending upload | PDF | 2.0 KB')
    expect(canPreviewDocument('application/pdf', 'newsletter.pdf')).toBe(true)
    expect(canPreviewDocument('image/png', 'newsletter.png')).toBe(true)
    expect(canPreviewDocument('application/msword', 'newsletter.doc')).toBe(false)
  })
})

describe('NewsletterEditorPage', () => {
  beforeEach(async () => {
    vi.clearAllMocks()
    mockNavigate.mockReset()
    mockCreate.mockReset()
    mockUpdate.mockReset()
    mockRemove.mockReset()
    toastSuccess.mockReset()
    toastError.mockReset()
    mockedFetchNewsletterEntry.mockReset()
    mockedAddNewsletterMedia.mockReset()
    mockedDeleteNewsletterMedia.mockReset()
    mockedGetNewsletterMediaContent.mockReset()
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
      () => `blob:newsletter-${++objectUrlIndex}`,
    )
    vi.mocked(URL.revokeObjectURL).mockImplementation(() => undefined)
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => undefined)
    vi.spyOn(axios, 'isAxiosError').mockImplementation(
      (value): value is AxiosError =>
        Boolean(value) && typeof value === 'object' && 'response' in value,
    )
    window.open = vi.fn() as unknown as typeof window.open
  })

  it('shows validation errors on create when required fields are missing', async () => {
    renderPage()

    fireEvent.click(screen.getByRole('button', { name: 'Save Draft Action' }))

    await waitFor(() => {
      expect(toastError).toHaveBeenCalledWith('Please correct the highlighted fields.')
    })
    expect(screen.getAllByText('Title is required.').length).toBeGreaterThan(0)
    expect(screen.getAllByText('Send date is required.').length).toBeGreaterThan(0)
    expect(mockCreate).not.toHaveBeenCalled()
  })

  it('creates a newsletter, uploads pending files, previews documents, and navigates to edit mode', async () => {
    mockCreate.mockResolvedValue({
      id: '77',
      title: 'Monthly Update',
      category: 'cst',
      sendDate: '2026-08-20',
      contentHtml: '',
      status: 'draft',
      visibility: 'public',
      publishAt: null,
      media: [],
      createdAt: '2026-08-01T00:00:00Z',
      updatedAt: '2026-08-01T00:00:00Z',
    })
    mockedAddNewsletterMedia.mockResolvedValue({
      message: 'uploaded',
      uploadedCount: 1,
    })

    const { container } = renderPage()

    fireEvent.change(screen.getByLabelText('Newsletter title'), {
      target: { value: ' Monthly Update ' },
    })
    fireEvent.change(screen.getByLabelText('Category'), {
      target: { value: 'cst' },
    })

    const dateInput = container.querySelector('input[type="date"]') as HTMLInputElement
    fireEvent.change(dateInput, { target: { value: '2026-08-20' } })

    fireEvent.click(screen.getByRole('button', { name: 'Mock upload file' }))

    await screen.findByText('newsletter-preview')
    expect(screen.getByText(/pending upload/i)).toBeDefined()

    fireEvent.click(screen.getByRole('button', { name: /preview/i }))
    expect(window.open).toHaveBeenCalledWith(
      'blob:newsletter-1',
      '_blank',
      'noopener,noreferrer',
    )

    fireEvent.click(screen.getByRole('button', { name: /download/i }))
    await waitFor(() => {
      expect(HTMLAnchorElement.prototype.click).toHaveBeenCalled()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save Draft Action' }))

    await waitFor(() => {
      expect(mockCreate).toHaveBeenCalledWith({
        title: 'Monthly Update',
        category: 'cst',
        sendDate: '2026-08-20',
        contentHtml: '',
        visibility: 'public',
        publishAt: null,
        status: 'draft',
      })
    })

    const [entryId, files, metadata] = mockedAddNewsletterMedia.mock.calls[0]
    expect(entryId).toBe('77')
    expect(files).toHaveLength(1)
    expect((files[0] as File).name).toBe('newsletter-preview.pdf')
    expect(metadata).toEqual([
      {
        display_name: 'newsletter-preview.pdf',
        file_name: 'newsletter-preview.pdf',
      },
    ])
    expect(toastSuccess).toHaveBeenCalledWith('Newsletter saved')
    expect(mockNavigate).toHaveBeenCalledWith('/newsletters/77/edit', {
      replace: true,
    })
  })

  it('loads an existing newsletter, prepopulates fields, and removes saved media', async () => {
    mockedFetchNewsletterEntry
      .mockResolvedValueOnce(sampleEntry)
      .mockResolvedValueOnce({
        ...sampleEntry,
        media: [],
      })
    mockedGetNewsletterMediaContent.mockResolvedValue(
      new Blob(['server-pdf'], { type: 'application/pdf' }),
    )
    mockedDeleteNewsletterMedia.mockResolvedValue({
      message: 'deleted',
      deletedCount: 1,
    })
    mockRemove.mockResolvedValue(undefined)

    const { container } = renderPage('/newsletters/18/edit')

    await waitFor(() => {
      expect(screen.getByDisplayValue('July Digest')).toBeDefined()
    })

    expect((screen.getByLabelText('Category') as HTMLSelectElement).value).toBe(
      'csaa',
    )
    const dateInput = container.querySelector('input[type="date"]') as HTMLInputElement
    expect(dateInput.value).toBe('2026-07-10')
    expect(mockedGetNewsletterMediaContent).toHaveBeenCalledWith('18', '91')

    await screen.findByTitle('digest')

    fireEvent.click(screen.getByRole('button', { name: /preview/i }))
    expect(window.open).toHaveBeenCalledWith(
      'blob:newsletter-1',
      '_blank',
      'noopener,noreferrer',
    )

    fireEvent.click(screen.getByRole('button', { name: /download/i }))
    await waitFor(() => {
      expect(HTMLAnchorElement.prototype.click).toHaveBeenCalled()
    })

    fireEvent.click(screen.getByRole('button', { name: /remove file/i }))
    await waitFor(() => {
      expect(mockedDeleteNewsletterMedia).toHaveBeenCalledWith('18', [91])
    })
    await waitFor(() => {
      expect(screen.queryByText('digest')).toBeNull()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Delete Action' }))
    await waitFor(() => {
      expect(mockRemove).toHaveBeenCalledWith('18')
    })
    expect(toastSuccess).toHaveBeenCalledWith('Newsletter deleted')
    expect(mockNavigate).toHaveBeenCalledWith('/newsletters', { replace: true })
  })

  it('shows a not found state for missing edit entries', async () => {
    mockedFetchNewsletterEntry.mockRejectedValueOnce({
      response: { status: 404 },
    })

    renderPage('/newsletters/404/edit')

    await screen.findByRole('heading', { name: /newsletter not found/i })
    expect(toastError).not.toHaveBeenCalled()

    fireEvent.click(screen.getByRole('button', { name: /back to newsletters/i }))
    expect(mockNavigate).toHaveBeenCalledWith('/newsletters')
  })

  it('shows a load error state when fetching an entry fails generically', async () => {
    mockedFetchNewsletterEntry.mockRejectedValueOnce(
      new Error('Backend unavailable'),
    )

    renderPage('/newsletters/501/edit')

    await screen.findByRole('heading', { name: /could not load this newsletter/i })
    expect(screen.getByText('Backend unavailable')).toBeDefined()
    expect(toastError).toHaveBeenCalledWith('Backend unavailable')
  })
})
