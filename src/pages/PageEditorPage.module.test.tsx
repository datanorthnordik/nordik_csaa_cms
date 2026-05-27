import { fireEvent, render, screen, waitFor } from '@testing-library/react'
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
