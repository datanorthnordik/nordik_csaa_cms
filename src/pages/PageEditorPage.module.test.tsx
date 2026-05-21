import { render, screen, waitFor } from '@testing-library/react'
import type { ReactNode } from 'react'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import { PageEditorPage } from './PageEditorPage'

const {
  dispatchMock,
  navigateMock,
  useAppSelectorMock,
  listPageParentOptionsMock,
  listGalleriesMock,
  fetchPageHeroImageContentMock,
} = vi.hoisted(() => ({
  dispatchMock: vi.fn(),
  navigateMock: vi.fn(),
  useAppSelectorMock: vi.fn(),
  listPageParentOptionsMock: vi.fn(),
  listGalleriesMock: vi.fn(),
  fetchPageHeroImageContentMock: vi.fn(),
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
      fetchPageHeroImageContent: fetchPageHeroImageContentMock,
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
  UploadDropzone: () => <div>Upload dropzone</div>,
}))

vi.mock('../store/hooks', () => ({
  useAppDispatch: () => dispatchMock,
  useAppSelector: (selector: (state: unknown) => unknown) => useAppSelectorMock(selector),
}))

describe('PageEditorPage module pages', () => {
  beforeEach(async () => {
    await i18n.changeLanguage('en')
    dispatchMock.mockReset()
    navigateMock.mockReset()
    listPageParentOptionsMock.mockReset()
    listGalleriesMock.mockReset()
    fetchPageHeroImageContentMock.mockReset()
    listPageParentOptionsMock.mockResolvedValue([])
    listGalleriesMock.mockResolvedValue([])
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
})
