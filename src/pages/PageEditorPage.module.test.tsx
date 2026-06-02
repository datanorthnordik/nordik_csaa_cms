import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import type { ReactNode } from 'react'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import toast from 'react-hot-toast'
import i18n from '../i18n'
import { PageEditorPage, getSectionDragAutoScrollDelta } from './PageEditorPage'

const {
  dispatchMock,
  navigateMock,
  useAppSelectorMock,
  listPageParentOptionsMock,
  listGalleriesMock,
  fetchPageDocumentContentMock,
  fetchPageHeroImageContentMock,
  fetchPageCTABannerImageContentMock,
} = vi.hoisted(() => ({
  dispatchMock: vi.fn(),
  navigateMock: vi.fn(),
  useAppSelectorMock: vi.fn(),
  listPageParentOptionsMock: vi.fn(),
  listGalleriesMock: vi.fn(),
  fetchPageDocumentContentMock: vi.fn(),
  fetchPageHeroImageContentMock: vi.fn(),
  fetchPageCTABannerImageContentMock: vi.fn(),
}))

vi.mock('react-hot-toast', () => ({
  default: {
    success: vi.fn(),
    error: vi.fn(),
  },
}))

vi.mock('react-router-dom', async () => {
  const actual = await vi.importActual<typeof import('react-router-dom')>('react-router-dom')
  return {
    ...actual,
    useNavigate: () => navigateMock,
    useParams: () => ({ id: '12' }),
  }
})

vi.mock('../api/pagesApi', async () => {
  const actual = await vi.importActual<typeof import('../api/pagesApi')>('../api/pagesApi')
  return {
    ...actual,
    pagesApi: {
      ...actual.pagesApi,
      listPageParentOptions: listPageParentOptionsMock,
      fetchPageDocumentContent: fetchPageDocumentContentMock,
      fetchPageHeroImageContent: fetchPageHeroImageContentMock,
      fetchPageCTABannerImageContent: fetchPageCTABannerImageContentMock,
    },
  }
})

vi.mock('../api/mediaApi', () => ({
  mediaApi: {
    listGalleries: listGalleriesMock,
  },
}))

vi.mock('../components/Breadcrumb', () => ({
  Breadcrumb: () => <nav>Breadcrumb</nav>,
}))

vi.mock('../components/CmsAppShell', () => ({
  CmsAppShell: ({ children }: { children?: ReactNode }) => <div>{children}</div>,
}))

vi.mock('../components/Loader', () => ({
  Loader: ({ label }: { label: string }) => <div>{label}</div>,
}))

vi.mock('../components/media/UploadDropzone', () => ({
  UploadDropzone: ({
    accept,
    disabled,
    hint,
    label,
    onFiles,
  }: {
    accept?: string
    disabled?: boolean
    hint?: string
    label: string
    onFiles?: (files: File[]) => void
  }) => (
    <div>
      <span>{label}</span>
      {hint ? <span>{hint}</span> : null}
      <input
        type="file"
        aria-label={label}
        accept={accept}
        disabled={disabled}
        onChange={(event) => onFiles?.(Array.from(event.target.files ?? []))}
      />
    </div>
  ),
}))

vi.mock('../store/hooks', () => ({
  useAppDispatch: () => dispatchMock,
  useAppSelector: (selector: (state: unknown) => unknown) => useAppSelectorMock(selector),
}))

function setEditablePageState(
  sections: unknown[] = [],
  overrides: Partial<Record<string, unknown>> = {},
) {
  useAppSelectorMock.mockImplementation((selector: (state: unknown) => unknown) =>
    selector({
      pages: {
        detail: {
          item: {
            id: 12,
            page_title: 'Community Care',
            url_slug: '/community-care',
            page_type: 'page',
            parent_id: null,
            parent_page_title: '',
            parent_page_url_slug: '',
            status: 'draft',
            hero_image_enabled: false,
            hero_image_url: '',
            hero_image_object_key: '',
            hero_image_fetch_url: '',
            seo_page_title: 'Community Care',
            seo_page_description: 'Community Care page',
            created_by: null,
            created_by_name: '',
            modified_by: null,
            modified_by_name: '',
            last_modified: '2026-05-13T00:00:00Z',
            created_at: '2026-05-13T00:00:00Z',
            updated_at: '2026-05-13T00:00:00Z',
            page_detail: {
              id: 3,
              page_id: 12,
              template_key: 'default',
              schema_version: 1,
              sections,
            },
            ...overrides,
          },
          status: 'succeeded',
          error: null,
        },
        save: {
          status: 'idle',
          error: null,
          lastResult: null,
        },
      },
    }),
  )
}

function getSectionCard(title: string) {
  const card = screen.getByRole('heading', { name: title }).closest('article')
  if (!card) {
    throw new Error(`Expected section card for ${title}`)
  }

  return card as HTMLElement
}

function mockElementRect(
  element: HTMLElement,
  rect: Partial<Pick<DOMRect, 'top' | 'left' | 'width' | 'height'>>,
) {
  return vi.spyOn(element, 'getBoundingClientRect').mockReturnValue({
    top: rect.top ?? 0,
    left: rect.left ?? 0,
    width: rect.width ?? 320,
    height: rect.height ?? 120,
    right: (rect.left ?? 0) + (rect.width ?? 320),
    bottom: (rect.top ?? 0) + (rect.height ?? 120),
    x: rect.left ?? 0,
    y: rect.top ?? 0,
    toJSON: () => ({}),
  } as DOMRect)
}

describe('PageEditorPage module pages', () => {
  beforeEach(async () => {
    await i18n.changeLanguage('en')
    dispatchMock.mockReset()
    navigateMock.mockReset()
    listPageParentOptionsMock.mockReset()
    listGalleriesMock.mockReset()
    fetchPageDocumentContentMock.mockReset()
    fetchPageHeroImageContentMock.mockReset()
    fetchPageCTABannerImageContentMock.mockReset()
    vi.mocked(toast.error).mockReset()
    vi.mocked(toast.success).mockReset()
    listPageParentOptionsMock.mockResolvedValue([])
    listGalleriesMock.mockResolvedValue([])
    fetchPageDocumentContentMock.mockResolvedValue(new Blob(['document-preview']))
    URL.createObjectURL = vi.fn(() => 'blob:preview')
    URL.revokeObjectURL = vi.fn()
    globalThis.window.scrollBy = vi.fn()
    useAppSelectorMock.mockImplementation((selector: (state: unknown) => unknown) =>
      selector({
        pages: {
          detail: {
            item: {
              id: 12,
              page_title: 'Events',
              url_slug: '/events',
              page_type: 'module',
              parent_id: null,
              parent_page_title: '',
              parent_page_url_slug: '',
              status: 'published',
              hero_image_enabled: false,
              hero_image_url: '',
              hero_image_object_key: '',
              hero_image_fetch_url: '',
              seo_page_title: 'Events',
              seo_page_description: 'Events module landing page',
              created_by: null,
              created_by_name: '',
              modified_by: null,
              modified_by_name: '',
              last_modified: '2026-05-13T00:00:00Z',
              created_at: '2026-05-13T00:00:00Z',
              updated_at: '2026-05-13T00:00:00Z',
            },
            status: 'succeeded',
            error: null,
          },
          save: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
        },
      }),
    )
  })

  it('renders module pages as read-only even on the edit route', async () => {
    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    expect(screen.getByRole('heading', { name: 'View Page' })).toBeDefined()
    expect(screen.getByText('Module-managed page')).toBeDefined()
    expect(
      screen.getByText(
        'This page belongs to a CMS module. Its content is managed in that module and cannot be edited here.',
      ),
    ).toBeDefined()
    expect(screen.queryByRole('button', { name: 'Save as Draft' })).toBeNull()
    expect(screen.queryByRole('button', { name: 'Publish Page' })).toBeNull()
    expect(screen.queryByRole('button', { name: 'Edit Page' })).toBeNull()
    expect(screen.queryByRole('heading', { name: 'Content Modules' })).toBeNull()
  })

  it('moves the first section to the end when dropped on the lower half of the last section', async () => {
    setEditablePageState([
      {
        id: 1,
        section_name: 'First Section',
        section_type: 'header',
        sort_order: 0,
        is_enabled: true,
        header: {
          main_header_text: 'First',
          sub_header_text: '',
          description: '',
          hierarchy: 'h2_section',
          text_align: 'left',
          underline_enabled: false,
        },
        created_at: '2026-05-13T00:00:00Z',
        updated_at: '2026-05-13T00:00:00Z',
      },
      {
        id: 2,
        section_name: 'Middle Section',
        section_type: 'quote',
        sort_order: 1,
        is_enabled: true,
        quote: {
          quote_content: 'Middle quote',
          attribution: '',
        },
        created_at: '2026-05-13T00:00:00Z',
        updated_at: '2026-05-13T00:00:00Z',
      },
      {
        id: 3,
        section_name: 'Last Section',
        section_type: 'cta_banner',
        sort_order: 2,
        is_enabled: true,
        cta_banner: {
          banner_heading: 'Last CTA',
          banner_message: '',
          button_text: 'Learn more',
          button_url: 'https://example.com',
          open_in_new_tab: true,
        },
        created_at: '2026-05-13T00:00:00Z',
        updated_at: '2026-05-13T00:00:00Z',
      },
    ])

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    const firstCard = getSectionCard('First Section')
    const lastCard = getSectionCard('Last Section')
    const dataTransfer = { dropEffect: '' }
    const rectSpy = mockElementRect(lastCard, { top: 0, height: 120 })

    fireEvent.dragStart(firstCard)
    fireEvent.dragOver(lastCard, { clientY: 100, dataTransfer })
    fireEvent.drop(lastCard, { clientY: 100, dataTransfer })

    expect(screen.getAllByRole('heading', { level: 3 }).map((node) => node.textContent)).toEqual([
      'Middle Section',
      'Last Section',
      'First Section',
    ])

    rectSpy.mockRestore()
  })

  it('returns scroll offsets while dragging near the viewport edges', () => {
    expect(getSectionDragAutoScrollDelta(20, 900)).toBe(-32)
    expect(getSectionDragAutoScrollDelta(880, 900)).toBe(32)
    expect(getSectionDragAutoScrollDelta(300, 900)).toBe(0)
  })

  it('rejects unsupported hero image uploads', async () => {
    useAppSelectorMock.mockImplementation((selector: (state: unknown) => unknown) =>
      selector({
        pages: {
          detail: {
            item: {
              id: 12,
              page_title: 'Community Care',
              url_slug: '/community-care',
              page_type: 'page',
              parent_id: null,
              parent_page_title: '',
              parent_page_url_slug: '',
              status: 'draft',
              hero_image_enabled: false,
              hero_image_url: '',
              hero_image_object_key: '',
              hero_image_fetch_url: '',
              seo_page_title: 'Community Care',
              seo_page_description: 'Community Care page',
              created_by: null,
              created_by_name: '',
              modified_by: null,
              modified_by_name: '',
              last_modified: '2026-05-13T00:00:00Z',
              created_at: '2026-05-13T00:00:00Z',
              updated_at: '2026-05-13T00:00:00Z',
            },
            status: 'succeeded',
            error: null,
          },
          save: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
        },
      }),
    )

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    fireEvent.click(screen.getByRole('checkbox', { name: /enable hero image/i }))
    fireEvent.change(screen.getByLabelText('Drop image here or browse'), {
      target: {
        files: [new File(['pdf'], 'hero.pdf', { type: 'application/pdf' })],
      },
    })

    expect(
      (
        await screen.findAllByText('Hero image must be a PNG, JPG, or WEBP file.')
      ).length,
    ).toBeGreaterThan(0)
    expect(toast.error).toHaveBeenCalledWith(
      'Hero image must be a PNG, JPG, or WEBP file.',
    )
  })

  it('rejects hero images larger than 5MB', async () => {
    useAppSelectorMock.mockImplementation((selector: (state: unknown) => unknown) =>
      selector({
        pages: {
          detail: {
            item: {
              id: 12,
              page_title: 'Community Care',
              url_slug: '/community-care',
              page_type: 'page',
              parent_id: null,
              parent_page_title: '',
              parent_page_url_slug: '',
              status: 'draft',
              hero_image_enabled: false,
              hero_image_url: '',
              hero_image_object_key: '',
              hero_image_fetch_url: '',
              seo_page_title: 'Community Care',
              seo_page_description: 'Community Care page',
              created_by: null,
              created_by_name: '',
              modified_by: null,
              modified_by_name: '',
              last_modified: '2026-05-13T00:00:00Z',
              created_at: '2026-05-13T00:00:00Z',
              updated_at: '2026-05-13T00:00:00Z',
            },
            status: 'succeeded',
            error: null,
          },
          save: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
        },
      }),
    )

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    const largeHeroImage = new File(['image'], 'hero.png', { type: 'image/png' })
    Object.defineProperty(largeHeroImage, 'size', {
      configurable: true,
      value: 5 * 1024 * 1024 + 1,
    })

    fireEvent.click(screen.getByRole('checkbox', { name: /enable hero image/i }))
    fireEvent.change(screen.getByLabelText('Drop image here or browse'), {
      target: {
        files: [largeHeroImage],
      },
    })

    expect((await screen.findAllByText('Hero image must be 5MB or smaller.')).length).toBeGreaterThan(0)
    expect(toast.error).toHaveBeenCalledWith('Hero image must be 5MB or smaller.')
  })

  it('shows field-level validation for page title and url slug', async () => {
    setEditablePageState()

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    fireEvent.change(screen.getByRole('textbox', { name: 'Page Title' }), {
      target: { value: '' },
    })
    fireEvent.change(screen.getByPlaceholderText('slug-name'), {
      target: { value: '' },
    })

    expect((await screen.findAllByText('Page title is required.')).length).toBeGreaterThan(0)
    expect((await screen.findAllByText('URL slug is required.')).length).toBeGreaterThan(0)
  })

  it('shows a field-level error and toast when the header module is missing main header text', async () => {
    setEditablePageState([
      {
        id: 4,
        section_name: 'Page heading',
        section_type: 'header',
        sort_order: 0,
        is_enabled: true,
        header: {
          main_header_text: '',
          sub_header_text: '',
          description: '',
          hierarchy: 'h1_hero',
          text_align: 'left',
          underline_enabled: false,
        },
        created_at: '2026-05-13T00:00:00Z',
        updated_at: '2026-05-13T00:00:00Z',
      },
    ])

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save as Draft' }))

    expect(
      (
        await screen.findAllByText('Main header text is required for the Header Module.')
      ).length,
    ).toBeGreaterThan(0)
    expect(toast.error).toHaveBeenCalledWith(
      'Main header text is required for the Header Module.',
    )
  })

  it('shows a field-level error and toast when the typography module content is empty', async () => {
    setEditablePageState([
      {
        id: 4,
        section_name: 'Story body',
        section_type: 'typography',
        sort_order: 0,
        is_enabled: true,
        typography: {
          html_content: '',
          text_content: '',
          text_align: 'left',
        },
        created_at: '2026-05-13T00:00:00Z',
        updated_at: '2026-05-13T00:00:00Z',
      },
    ])

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save as Draft' }))

    expect((await screen.findAllByText('Typography content is required.')).length).toBeGreaterThan(0)
    expect(toast.error).toHaveBeenCalledWith('Typography content is required.')
  })

  it('shows a field-level error and toast when the quote module is missing quote text', async () => {
    setEditablePageState([
      {
        id: 4,
        section_name: 'Quote spotlight',
        section_type: 'quote',
        sort_order: 0,
        is_enabled: true,
        quote: {
          quote_content: '',
          attribution: '',
        },
        created_at: '2026-05-13T00:00:00Z',
        updated_at: '2026-05-13T00:00:00Z',
      },
    ])

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save as Draft' }))

    expect((await screen.findAllByText('Quote text is required.')).length).toBeGreaterThan(0)
    expect(toast.error).toHaveBeenCalledWith('Quote text is required.')
  })

  it('does not re-fetch document previews when only document heading text changes', async () => {
    setEditablePageState([
      {
        id: 4,
        section_name: 'Documents',
        section_type: 'document',
        sort_order: 0,
        is_enabled: true,
        documents: {
          items: [
            {
              id: 9,
              display_name: 'Annual Report',
              description: '',
              original_file_name: 'annual-report.pdf',
              file_name: 'annual-report.pdf',
              file_url: 'gs://bucket/annual-report.pdf',
              fetch_url: '/api/pages/documents/9/content',
              storage_uri: 'gs://bucket/annual-report.pdf',
              mime_type: 'application/pdf',
              gcp_object_key: 'pages/documents/annual-report.pdf',
              file_size: 1024,
            },
          ],
        },
        created_at: '2026-05-13T00:00:00Z',
        updated_at: '2026-05-13T00:00:00Z',
      },
    ])

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(fetchPageDocumentContentMock).toHaveBeenCalledTimes(1)
    })

    fireEvent.change(screen.getByDisplayValue('Annual Report'), {
      target: { value: 'Updated Annual Report' },
    })

    await waitFor(() => {
      expect(fetchPageDocumentContentMock).toHaveBeenCalledTimes(1)
    })
  })

  it('shows a field-level error and toast when the document module has no documents', async () => {
    setEditablePageState([
      {
        id: 4,
        section_name: 'Documents',
        section_type: 'document',
        sort_order: 0,
        is_enabled: true,
        documents: {
          items: [],
        },
        created_at: '2026-05-13T00:00:00Z',
        updated_at: '2026-05-13T00:00:00Z',
      },
    ])

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save as Draft' }))

    expect(
      (
        await screen.findAllByText(
          'At least one document is required for the Document Module.',
        )
      ).length,
    ).toBeGreaterThan(0)
    expect(toast.error).toHaveBeenCalledWith(
      'At least one document is required for the Document Module.',
    )
  })

  it('rejects unsupported document uploads with the press-entry file rules', async () => {
    setEditablePageState([
      {
        id: 4,
        section_name: 'Documents',
        section_type: 'document',
        sort_order: 0,
        is_enabled: true,
        documents: {
          items: [],
        },
        created_at: '2026-05-13T00:00:00Z',
        updated_at: '2026-05-13T00:00:00Z',
      },
    ])

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    fireEvent.change(screen.getByLabelText('Drag and drop documents here or browse'), {
      target: {
        files: [new File(['zip'], 'archive.zip', { type: 'application/zip' })],
      },
    })

    expect(
      (
        await screen.findAllByText(
          'Only PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP are supported.',
        )
      ).length,
    ).toBeGreaterThan(0)
    expect(toast.error).toHaveBeenCalledWith(
      'Only PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP are supported.',
    )
  })

  it('shows a field-level error and toast when a document heading is missing', async () => {
    setEditablePageState([
      {
        id: 4,
        section_name: 'Documents',
        section_type: 'document',
        sort_order: 0,
        is_enabled: true,
        documents: {
          items: [
            {
              id: 9,
              display_name: '',
              description: '',
              original_file_name: 'annual-report.pdf',
              file_name: 'annual-report.pdf',
              file_url: 'gs://bucket/annual-report.pdf',
              fetch_url: '/api/pages/documents/9/content',
              storage_uri: 'gs://bucket/annual-report.pdf',
              mime_type: 'application/pdf',
              gcp_object_key: 'pages/documents/annual-report.pdf',
              file_size: 1024,
            },
          ],
        },
        created_at: '2026-05-13T00:00:00Z',
        updated_at: '2026-05-13T00:00:00Z',
      },
    ])

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save as Draft' }))

    expect((await screen.findAllByText('Document display name is required.')).length).toBeGreaterThan(0)
    expect(toast.error).toHaveBeenCalledWith('Document display name is required.')
  })

  it('shows field-level errors and a toast when CTA required fields are missing', async () => {
    setEditablePageState([
      {
        id: 4,
        section_name: 'CTA Banner',
        section_type: 'cta_banner',
        sort_order: 0,
        is_enabled: true,
        cta_banner: {
          banner_heading: '',
          banner_message: '',
          button_text: '',
          button_url: '',
          open_in_new_tab: false,
        },
        created_at: '2026-05-13T00:00:00Z',
        updated_at: '2026-05-13T00:00:00Z',
      },
    ])

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    fireEvent.click(screen.getByRole('button', { name: 'Save as Draft' }))

    expect(
      (
        await screen.findAllByText(
          'Banner heading is required for the CTA Banner Module.',
        )
      ).length,
    ).toBeGreaterThan(0)
    expect(
      (await screen.findAllByText('Button text is required for the CTA Banner Module.'))
        .length,
    ).toBeGreaterThan(0)
    expect(
      (await screen.findAllByText('Button URL is required for the CTA Banner Module.'))
        .length,
    ).toBeGreaterThan(0)
    expect(toast.error).toHaveBeenCalledWith(
      'Banner heading is required for the CTA Banner Module.',
    )
  })

  it('shows the icons gallery view option for editable page gallery sections', async () => {
    listGalleriesMock.mockResolvedValue([{ id: 14, name: 'Partner Logos' }])
    useAppSelectorMock.mockImplementation((selector: (state: unknown) => unknown) =>
      selector({
        pages: {
          detail: {
            item: {
              id: 12,
              page_title: 'Partners',
              url_slug: '/partners',
              page_type: 'page',
              parent_id: null,
              parent_page_title: '',
              parent_page_url_slug: '',
              status: 'draft',
              hero_image_enabled: false,
              hero_image_url: '',
              hero_image_object_key: '',
              hero_image_fetch_url: '',
              seo_page_title: 'Partners',
              seo_page_description: 'Partner logos',
              created_by: null,
              created_by_name: '',
              modified_by: null,
              modified_by_name: '',
              last_modified: '2026-05-13T00:00:00Z',
              created_at: '2026-05-13T00:00:00Z',
              updated_at: '2026-05-13T00:00:00Z',
              page_detail: {
                id: 3,
                page_id: 12,
                template_key: 'default',
                schema_version: 1,
                sections: [
                  {
                    id: 4,
                    section_name: 'Partner logos',
                    section_type: 'gallery',
                    sort_order: 0,
                    is_enabled: true,
                    gallery: {
                      gallery_id: 14,
                      view_mode: 'icons',
                    },
                    created_at: '2026-05-13T00:00:00Z',
                    updated_at: '2026-05-13T00:00:00Z',
                  },
                ],
              },
            },
            status: 'succeeded',
            error: null,
          },
          save: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
        },
      }),
    )

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    expect(screen.getByRole('button', { name: 'Icons' })).toBeDefined()
  })

  it('shows gallery caption and carousel auto-scroll toggles for editable gallery sections', async () => {
    listGalleriesMock.mockResolvedValue([{ id: 14, name: 'Partner Logos' }])
    useAppSelectorMock.mockImplementation((selector: (state: unknown) => unknown) =>
      selector({
        pages: {
          detail: {
            item: {
              id: 12,
              page_title: 'Partners',
              url_slug: '/partners',
              page_type: 'page',
              parent_id: null,
              parent_page_title: '',
              parent_page_url_slug: '',
              status: 'draft',
              hero_image_enabled: false,
              hero_image_url: '',
              hero_image_object_key: '',
              hero_image_fetch_url: '',
              seo_page_title: 'Partners',
              seo_page_description: 'Partner logos',
              created_by: null,
              created_by_name: '',
              modified_by: null,
              modified_by_name: '',
              last_modified: '2026-05-13T00:00:00Z',
              created_at: '2026-05-13T00:00:00Z',
              updated_at: '2026-05-13T00:00:00Z',
              page_detail: {
                id: 3,
                page_id: 12,
                template_key: 'default',
                schema_version: 1,
                sections: [
                  {
                    id: 4,
                    section_name: 'Partner logos',
                    section_type: 'gallery',
                    sort_order: 0,
                    is_enabled: true,
                    gallery: {
                      gallery_id: 14,
                      view_mode: 'carousel',
                      show_title_description: false,
                      auto_scroll_enabled: true,
                    },
                    created_at: '2026-05-13T00:00:00Z',
                    updated_at: '2026-05-13T00:00:00Z',
                  },
                ],
              },
            },
            status: 'succeeded',
            error: null,
          },
          save: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
        },
      }),
    )

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    expect(
      screen
        .getByRole('switch', {
          name: /display titles and descriptions/i,
        })
        .getAttribute('aria-checked'),
    ).toBe('false')
    expect(
      screen
        .getByRole('switch', {
          name: /automatic carousel scroll/i,
        })
        .getAttribute('aria-checked'),
    ).toBe('true')
  })

  it('disables automatic carousel scroll outside carousel view', async () => {
    listGalleriesMock.mockResolvedValue([{ id: 14, name: 'Partner Logos' }])
    useAppSelectorMock.mockImplementation((selector: (state: unknown) => unknown) =>
      selector({
        pages: {
          detail: {
            item: {
              id: 12,
              page_title: 'Partners',
              url_slug: '/partners',
              page_type: 'page',
              parent_id: null,
              parent_page_title: '',
              parent_page_url_slug: '',
              status: 'draft',
              hero_image_enabled: false,
              hero_image_url: '',
              hero_image_object_key: '',
              hero_image_fetch_url: '',
              seo_page_title: 'Partners',
              seo_page_description: 'Partner logos',
              created_by: null,
              created_by_name: '',
              modified_by: null,
              modified_by_name: '',
              last_modified: '2026-05-13T00:00:00Z',
              created_at: '2026-05-13T00:00:00Z',
              updated_at: '2026-05-13T00:00:00Z',
              page_detail: {
                id: 3,
                page_id: 12,
                template_key: 'default',
                schema_version: 1,
                sections: [
                  {
                    id: 4,
                    section_name: 'Partner logos',
                    section_type: 'gallery',
                    sort_order: 0,
                    is_enabled: true,
                    gallery: {
                      gallery_id: 14,
                      view_mode: 'grid',
                      show_title_description: true,
                      auto_scroll_enabled: true,
                    },
                    created_at: '2026-05-13T00:00:00Z',
                    updated_at: '2026-05-13T00:00:00Z',
                  },
                ],
              },
            },
            status: 'succeeded',
            error: null,
          },
          save: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
        },
      }),
    )

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    const autoScrollSwitch = screen.getByRole('switch', {
      name: /automatic carousel scroll/i,
    }) as HTMLButtonElement

    expect(autoScrollSwitch.disabled).toBe(true)
    expect(autoScrollSwitch.getAttribute('aria-checked')).toBe('false')
    expect(
      screen.getByText('Switch the gallery view to Carousel to enable automatic scrolling.'),
    ).toBeDefined()
  })

  it('shows header description, underline, and text alignment controls for editable header sections', async () => {
    useAppSelectorMock.mockImplementation((selector: (state: unknown) => unknown) =>
      selector({
        pages: {
          detail: {
            item: {
              id: 12,
              page_title: 'About',
              url_slug: '/about',
              page_type: 'page',
              parent_id: null,
              parent_page_title: '',
              parent_page_url_slug: '',
              status: 'draft',
              hero_image_enabled: false,
              hero_image_url: '',
              hero_image_object_key: '',
              hero_image_fetch_url: '',
              seo_page_title: 'About',
              seo_page_description: 'About page',
              created_by: null,
              created_by_name: '',
              modified_by: null,
              modified_by_name: '',
              last_modified: '2026-05-13T00:00:00Z',
              created_at: '2026-05-13T00:00:00Z',
              updated_at: '2026-05-13T00:00:00Z',
              page_detail: {
                id: 4,
                page_id: 12,
                template_key: 'default',
                schema_version: 1,
                sections: [
                  {
                    id: 5,
                    section_name: 'Page heading',
                    section_type: 'header',
                    sort_order: 0,
                    is_enabled: true,
                    header: {
                      main_header_text: 'About Us',
                      sub_header_text: 'Who we are',
                      description: 'Community stories',
                      hierarchy: 'h2_section',
                      text_align: 'center',
                      underline_enabled: true,
                    },
                    created_at: '2026-05-13T00:00:00Z',
                    updated_at: '2026-05-13T00:00:00Z',
                  },
                ],
              },
            },
            status: 'succeeded',
            error: null,
          },
          save: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
        },
      }),
    )

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    expect(screen.getByText('Text Align')).toBeDefined()
    expect(
      (
        screen.getByRole('textbox', {
          name: /header description/i,
        }) as HTMLTextAreaElement
      ).value,
    ).toBe('Community stories')
    expect(
      screen
        .getByRole('switch', {
          name: /optional underline/i,
        })
        .getAttribute('aria-checked'),
    ).toBe('true')
    expect(screen.getByRole('button', { name: 'Left' })).toBeDefined()
    expect(screen.getByRole('button', { name: 'Center' })).toBeDefined()
    expect(screen.getByRole('button', { name: 'Right' })).toBeDefined()
  })

  it('disables h1 hero selection when another header section already uses h1', async () => {
    useAppSelectorMock.mockImplementation((selector: (state: unknown) => unknown) =>
      selector({
        pages: {
          detail: {
            item: {
              id: 12,
              page_title: 'About',
              url_slug: '/about',
              page_type: 'page',
              parent_id: null,
              parent_page_title: '',
              parent_page_url_slug: '',
              status: 'draft',
              hero_image_enabled: false,
              hero_image_url: '',
              hero_image_object_key: '',
              hero_image_fetch_url: '',
              seo_page_title: 'About',
              seo_page_description: 'About page',
              created_by: null,
              created_by_name: '',
              modified_by: null,
              modified_by_name: '',
              last_modified: '2026-05-13T00:00:00Z',
              created_at: '2026-05-13T00:00:00Z',
              updated_at: '2026-05-13T00:00:00Z',
              page_detail: {
                id: 4,
                page_id: 12,
                template_key: 'default',
                schema_version: 1,
                sections: [
                  {
                    id: 5,
                    section_name: 'Page heading',
                    section_type: 'header',
                    sort_order: 0,
                    is_enabled: true,
                    header: {
                      main_header_text: 'About Us',
                      sub_header_text: 'Who we are',
                      description: '',
                      hierarchy: 'h1_hero',
                      text_align: 'left',
                      underline_enabled: false,
                    },
                    created_at: '2026-05-13T00:00:00Z',
                    updated_at: '2026-05-13T00:00:00Z',
                  },
                  {
                    id: 6,
                    section_name: 'Secondary heading',
                    section_type: 'header',
                    sort_order: 1,
                    is_enabled: true,
                    header: {
                      main_header_text: 'Resources',
                      sub_header_text: '',
                      description: '',
                      hierarchy: 'h2_section',
                      text_align: 'left',
                      underline_enabled: false,
                    },
                    created_at: '2026-05-13T00:00:00Z',
                    updated_at: '2026-05-13T00:00:00Z',
                  },
                ],
              },
            },
            status: 'succeeded',
            error: null,
          },
          save: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
        },
      }),
    )

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    const h1Buttons = screen.getAllByRole('button', { name: 'H1 Hero' }) as HTMLButtonElement[]

    expect(h1Buttons).toHaveLength(2)
    expect(h1Buttons[0].disabled).toBe(false)
    expect(h1Buttons[1].disabled).toBe(true)
    expect(
      screen.getByText(
        'This page already has an H1 Hero header. Change that section to H2 before using H1 here.',
      ),
    ).toBeDefined()
  })

  it('disables optional underline outside h2 hierarchy', async () => {
    useAppSelectorMock.mockImplementation((selector: (state: unknown) => unknown) =>
      selector({
        pages: {
          detail: {
            item: {
              id: 12,
              page_title: 'About',
              url_slug: '/about',
              page_type: 'page',
              parent_id: null,
              parent_page_title: '',
              parent_page_url_slug: '',
              status: 'draft',
              hero_image_enabled: false,
              hero_image_url: '',
              hero_image_object_key: '',
              hero_image_fetch_url: '',
              seo_page_title: 'About',
              seo_page_description: 'About page',
              created_by: null,
              created_by_name: '',
              modified_by: null,
              modified_by_name: '',
              last_modified: '2026-05-13T00:00:00Z',
              created_at: '2026-05-13T00:00:00Z',
              updated_at: '2026-05-13T00:00:00Z',
              page_detail: {
                id: 4,
                page_id: 12,
                template_key: 'default',
                schema_version: 1,
                sections: [
                  {
                    id: 5,
                    section_name: 'Page heading',
                    section_type: 'header',
                    sort_order: 0,
                    is_enabled: true,
                    header: {
                      main_header_text: 'About Us',
                      sub_header_text: 'Who we are',
                      description: '',
                      hierarchy: 'h1_hero',
                      text_align: 'center',
                      underline_enabled: true,
                    },
                    created_at: '2026-05-13T00:00:00Z',
                    updated_at: '2026-05-13T00:00:00Z',
                  },
                ],
              },
            },
            status: 'succeeded',
            error: null,
          },
          save: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
        },
      }),
    )

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    const underlineSwitch = screen.getByRole('switch', {
      name: /optional underline/i,
    }) as HTMLButtonElement

    expect(underlineSwitch.disabled).toBe(true)
    expect(underlineSwitch.getAttribute('aria-checked')).toBe('false')
    expect(
      screen.getByText('Switch the hierarchy to H2 Section to enable the underline accent.'),
    ).toBeDefined()
  })

  it('uses the shared typography editor without the image action', async () => {
    useAppSelectorMock.mockImplementation((selector: (state: unknown) => unknown) =>
      selector({
        pages: {
          detail: {
            item: {
              id: 12,
              page_title: 'Story',
              url_slug: '/story',
              page_type: 'page',
              parent_id: null,
              parent_page_title: '',
              parent_page_url_slug: '',
              status: 'draft',
              hero_image_enabled: false,
              hero_image_url: '',
              hero_image_object_key: '',
              hero_image_fetch_url: '',
              seo_page_title: 'Story',
              seo_page_description: 'Story page',
              created_by: null,
              created_by_name: '',
              modified_by: null,
              modified_by_name: '',
              last_modified: '2026-05-13T00:00:00Z',
              created_at: '2026-05-13T00:00:00Z',
              updated_at: '2026-05-13T00:00:00Z',
              page_detail: {
                id: 5,
                page_id: 12,
                template_key: 'default',
                schema_version: 1,
                sections: [
                  {
                    id: 6,
                    section_name: 'Story body',
                    section_type: 'typography',
                    sort_order: 0,
                    is_enabled: true,
                    typography: {
                      html_content: '<p>Story body</p>',
                      text_content: 'Story body',
                      text_align: 'left',
                    },
                    created_at: '2026-05-13T00:00:00Z',
                    updated_at: '2026-05-13T00:00:00Z',
                  },
                ],
              },
            },
            status: 'succeeded',
            error: null,
          },
          save: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
        },
      }),
    )

    render(<PageEditorPage />)

    await waitFor(() => {
      expect(listPageParentOptionsMock).toHaveBeenCalledTimes(1)
      expect(listGalleriesMock).toHaveBeenCalledTimes(1)
    })

    expect(screen.getByRole('button', { name: 'Font size' })).toBeDefined()
    expect(screen.getByRole('button', { name: 'Insert link' })).toBeDefined()
    expect(screen.queryByRole('button', { name: 'Insert image' })).toBeNull()

    fireEvent.click(screen.getByRole('button', { name: 'Insert link' }))
    expect(screen.getByRole('textbox', { name: 'Link destination' })).toBeDefined()
  })
})
