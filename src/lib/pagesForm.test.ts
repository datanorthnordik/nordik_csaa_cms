import { describe, expect, it } from 'vitest'
import { API_BASE_URL } from '../constants/api'
import {
  buildFullPageUrlSlug,
  buildPageFormStateFromDetail,
  buildSavePageRequest,
  createDefaultDocumentState,
  createDefaultPageFormState,
  createDefaultSectionState,
  getDisallowedParentPageIds,
  reorderPageSections,
  validatePageDocumentFields,
  validatePageForm,
  validatePageSectionFields,
} from './pagesForm'

describe('buildSavePageRequest', () => {
  it('returns multipart-ready hero image metadata without base64 content', () => {
    const heroImageFile = new File(['hero-image'], 'hero.png', {
      type: 'image/png',
    })
    const form = createDefaultPageFormState()

    form.pageTitle = 'Homepage'
    form.urlSlug = 'home'
    form.heroImageEnabled = true
    form.heroImageFile = heroImageFile
    form.seoPageTitle = 'Homepage SEO'
    form.seoPageDescription = 'Description'

    const request = buildSavePageRequest(form)

    expect(request.heroImageFile).toBe(heroImageFile)
    expect(request.hero_image).toEqual({
      file_name: 'hero.png',
      mime_type: 'image/png',
    })
    expect(request.parent_id).toBeNull()
    expect(request.hero_image).not.toHaveProperty('data_base64')
  })

  it('prefixes the parent page slug and includes the parent id when selected', () => {
    const form = createDefaultPageFormState()

    form.pageTitle = 'Team'
    form.urlSlug = 'team'
    form.parentPageId = '7'

    const request = buildSavePageRequest(form, '/about')

    expect(request.parent_id).toBe(7)
    expect(request.url_slug).toBe('/about/team')
  })

  it('collects document section files and keeps the saved section ordering', () => {
    const attachment = new File(['content'], 'board-policy.pdf', {
      type: 'application/pdf',
    })
    const form = createDefaultPageFormState()
    const documentSection = createDefaultSectionState('document')

    form.pageTitle = 'Policies'
    form.urlSlug = 'policies'
    documentSection.documents.items = [createDefaultDocumentState(attachment)]
    form.sections = [documentSection]

    const request = buildSavePageRequest(form)

    expect(request.page_detail?.sections).toHaveLength(1)
    expect(request.page_detail?.sections[0]).toMatchObject({
      section_type: 'document',
      sort_order: 0,
      documents: {
        items: [
          {
            display_name: 'board-policy',
            file_name: 'board-policy.pdf',
            mime_type: 'application/pdf',
            original_file_name: 'board-policy.pdf',
          },
        ],
      },
    })
    expect(request.documentFiles).toEqual([
      {
        sectionIndex: 0,
        documentIndex: 0,
        file: attachment,
      },
    ])
  })

  it('preserves existing CTA image references and collects new CTA image uploads', () => {
    const ctaImageFile = new File(['cta-image'], 'cta.png', {
      type: 'image/png',
    })
    const form = createDefaultPageFormState()
    const existingCTASection = createDefaultSectionState('cta_banner')
    const uploadedCTASection = createDefaultSectionState('cta_banner')

    form.pageTitle = 'Community Support'
    form.urlSlug = 'community-support'

    existingCTASection.ctaBanner = {
      bannerHeading: 'We are here for the Community',
      bannerMessage: 'Our mission is rooted in honouring survivors.',
      buttonText: 'Learn more',
      buttonUrl: 'https://example.com/community-support',
      openInNewTab: true,
      imageFile: null,
      existingImageFetchUrl: `${API_BASE_URL}/api/pages/sections/10/cta-image/content`,
      existingImageStorageUrl: 'gs://drive-bucket/pages/sections/10/cta.png',
      existingImageObjectKey: 'pages/sections/10/cta.png',
    }

    uploadedCTASection.ctaBanner = {
      bannerHeading: 'Learn more about the CST',
      bannerMessage: 'Support for survivors, families, and community members.',
      buttonText: 'Read more',
      buttonUrl: 'https://example.com/cst',
      openInNewTab: true,
      imageFile: ctaImageFile,
      existingImageFetchUrl: '',
      existingImageStorageUrl: '',
      existingImageObjectKey: '',
    }

    form.sections = [existingCTASection, uploadedCTASection]

    const request = buildSavePageRequest(form)

    expect(request.page_detail?.sections[0]).toMatchObject({
      section_type: 'cta_banner',
      cta_banner: {
        open_in_new_tab: true,
        image: {
          file_url: 'gs://drive-bucket/pages/sections/10/cta.png',
          storage_uri: 'gs://drive-bucket/pages/sections/10/cta.png',
          gcp_object_key: 'pages/sections/10/cta.png',
        },
      },
    })
    expect(request.page_detail?.sections[1]).toMatchObject({
      section_type: 'cta_banner',
      cta_banner: {
        open_in_new_tab: true,
        image: {
          file_name: 'cta.png',
          mime_type: 'image/png',
        },
      },
    })
    expect(request.ctaBannerImageFiles).toEqual([
      {
        sectionIndex: 1,
        file: ctaImageFile,
      },
    ])
  })

  it('preserves the icons gallery view mode and disables carousel-only scrolling', () => {
    const form = createDefaultPageFormState()
    const gallerySection = createDefaultSectionState('gallery')

    form.pageTitle = 'Partners'
    form.urlSlug = 'partners'
    gallerySection.gallery = {
      galleryId: '14',
      viewMode: 'icons',
      showTitleDescription: false,
      autoScrollEnabled: true,
    }
    form.sections = [gallerySection]

    const request = buildSavePageRequest(form)

    expect(request.page_detail?.sections[0]).toMatchObject({
      section_type: 'gallery',
      sort_order: 0,
      gallery: {
        gallery_id: 14,
        view_mode: 'icons',
        show_title_description: false,
        auto_scroll_enabled: false,
      },
    })
  })

  it('includes header description, underline, and text alignment in page section payloads', () => {
    const form = createDefaultPageFormState()
    const headerSection = createDefaultSectionState('header')

    form.pageTitle = 'About'
    form.urlSlug = 'about'
    headerSection.header = {
      mainHeaderText: 'About Us',
      subHeaderText: 'Who we are',
      description: 'Community stories',
      hierarchy: 'h2_section',
      textAlign: 'center',
      underlineEnabled: true,
    }
    form.sections = [headerSection]

    const request = buildSavePageRequest(form)

    expect(request.page_detail?.sections[0]).toMatchObject({
      section_type: 'header',
      sort_order: 0,
      header: {
        main_header_text: 'About Us',
        sub_header_text: 'Who we are',
        description: 'Community stories',
        hierarchy: 'h2_section',
        text_align: 'center',
        underline_enabled: true,
      },
    })
  })

  it('forces header underline off when the hierarchy is not an h2 section', () => {
    const form = createDefaultPageFormState()
    const headerSection = createDefaultSectionState('header')

    form.pageTitle = 'About'
    form.urlSlug = 'about'
    headerSection.header = {
      mainHeaderText: 'About Us',
      subHeaderText: 'Who we are',
      description: '',
      hierarchy: 'h1_hero',
      textAlign: 'center',
      underlineEnabled: true,
    }
    form.sections = [headerSection]

    const request = buildSavePageRequest(form)

    expect(request.page_detail?.sections[0]).toMatchObject({
      section_type: 'header',
      sort_order: 0,
      header: {
        main_header_text: 'About Us',
        sub_header_text: 'Who we are',
        description: '',
        hierarchy: 'h1_hero',
        text_align: 'center',
        underline_enabled: false,
      },
    })
  })
})

describe('buildPageFormStateFromDetail', () => {
  it('strips the parent page slug back to the child segment when editing', () => {
    const form = buildPageFormStateFromDetail(
      {
        id: 3,
        page_title: 'Team',
        url_slug: '/about/team',
        page_type: 'page',
        parent_id: 7,
        status: 'draft',
        hero_image_enabled: false,
        hero_image_url: '',
        hero_image_object_key: '',
        hero_image_fetch_url: '',
        seo_page_title: '',
        seo_page_description: '',
        created_by: null,
        created_by_name: 'Admin',
        modified_by: null,
        modified_by_name: 'Admin',
        last_modified: '2026-05-12T00:00:00Z',
        created_at: '2026-05-12T00:00:00Z',
        updated_at: '2026-05-12T00:00:00Z',
      },
      {
        parentPageSlug: '/about',
      },
    )

    expect(form.parentPageId).toBe('7')
    expect(form.urlSlug).toBe('team')
  })

  it('hydrates content modules from page detail responses', () => {
    const form = buildPageFormStateFromDetail({
      id: 9,
      page_title: 'About',
      url_slug: '/about',
      page_type: 'page',
      parent_id: null,
      status: 'draft',
      hero_image_enabled: false,
      hero_image_url: '',
      hero_image_object_key: '',
      hero_image_fetch_url: '',
      seo_page_title: '',
      seo_page_description: '',
      created_by: null,
      created_by_name: 'Admin',
      modified_by: null,
      modified_by_name: 'Admin',
      last_modified: '2026-05-12T00:00:00Z',
      created_at: '2026-05-12T00:00:00Z',
      updated_at: '2026-05-12T00:00:00Z',
      page_detail: {
        id: 12,
        page_id: 9,
        template_key: 'default',
        schema_version: 1,
        sections: [
          {
            id: 3,
            section_name: 'Intro',
            section_type: 'typography',
            sort_order: 0,
            is_enabled: true,
            typography: {
              html_content: '<p>Hello world</p>',
              text_content: 'Hello world',
              text_align: 'center',
            },
            created_at: '2026-05-12T00:00:00Z',
            updated_at: '2026-05-12T00:00:00Z',
          },
        ],
      },
    })

    expect(form.templateKey).toBe('default')
    expect(form.sections).toHaveLength(1)
    expect(form.sections[0]).toMatchObject({
      id: 3,
      sectionName: 'Intro',
      sectionType: 'typography',
      typography: {
        htmlContent: '<p>Hello world</p>',
        textAlign: 'center',
      },
    })
  })

  it('hydrates gallery sections with the icons view mode', () => {
    const form = buildPageFormStateFromDetail({
      id: 10,
      page_title: 'Partners',
      url_slug: '/partners',
      page_type: 'page',
      parent_id: null,
      status: 'draft',
      hero_image_enabled: false,
      hero_image_url: '',
      hero_image_object_key: '',
      hero_image_fetch_url: '',
      seo_page_title: '',
      seo_page_description: '',
      created_by: null,
      created_by_name: 'Admin',
      modified_by: null,
      modified_by_name: 'Admin',
      last_modified: '2026-05-12T00:00:00Z',
      created_at: '2026-05-12T00:00:00Z',
      updated_at: '2026-05-12T00:00:00Z',
      page_detail: {
        id: 21,
        page_id: 10,
        template_key: 'default',
        schema_version: 1,
        sections: [
          {
            id: 8,
            section_name: 'Partner logos',
            section_type: 'gallery',
            sort_order: 0,
            is_enabled: true,
            gallery: {
              gallery_id: 14,
              view_mode: 'icons',
              show_title_description: false,
              auto_scroll_enabled: true,
            },
            created_at: '2026-05-12T00:00:00Z',
            updated_at: '2026-05-12T00:00:00Z',
          },
        ],
      },
    })

    expect(form.sections[0]).toMatchObject({
      id: 8,
      sectionName: 'Partner logos',
      sectionType: 'gallery',
      gallery: {
        galleryId: '14',
        viewMode: 'icons',
        showTitleDescription: false,
        autoScrollEnabled: true,
      },
    })
  })

  it('hydrates CTA image metadata from page detail responses', () => {
    const form = buildPageFormStateFromDetail({
      id: 9,
      page_title: 'Support',
      url_slug: '/support',
      page_type: 'page',
      parent_id: null,
      status: 'draft',
      hero_image_enabled: false,
      hero_image_url: '',
      hero_image_object_key: '',
      hero_image_fetch_url: '',
      seo_page_title: '',
      seo_page_description: '',
      created_by: null,
      created_by_name: 'Admin',
      modified_by: null,
      modified_by_name: 'Admin',
      last_modified: '2026-05-12T00:00:00Z',
      created_at: '2026-05-12T00:00:00Z',
      updated_at: '2026-05-12T00:00:00Z',
      page_detail: {
        id: 12,
        page_id: 9,
        template_key: 'default',
        schema_version: 1,
        sections: [
          {
            id: 3,
            section_name: 'CTA Banner',
            section_type: 'cta_banner',
            sort_order: 0,
            is_enabled: true,
            cta_banner: {
              banner_heading: 'Support',
              banner_message: 'We are here for the community.',
              button_text: 'Learn more',
              button_url: 'https://example.com/support',
              open_in_new_tab: false,
              image: {
                file_url: '/api/pages/sections/3/cta-image/content',
                fetch_url: '/api/pages/sections/3/cta-image/content',
                storage_uri: 'gs://drive-bucket/pages/sections/3/cta.png',
                gcp_object_key: 'pages/sections/3/cta.png',
              },
            },
            created_at: '2026-05-12T00:00:00Z',
            updated_at: '2026-05-12T00:00:00Z',
          },
        ],
      },
    })

    expect(form.sections[0]?.ctaBanner.existingImageFetchUrl).toBe(
      `${API_BASE_URL}/api/pages/sections/3/cta-image/content`,
    )
    expect(form.sections[0]?.ctaBanner.existingImageStorageUrl).toBe(
      'gs://drive-bucket/pages/sections/3/cta.png',
    )
    expect(form.sections[0]?.ctaBanner.existingImageObjectKey).toBe(
      'pages/sections/3/cta.png',
    )
    expect(form.sections[0]?.ctaBanner.openInNewTab).toBe(true)
  })

  it('forces auto-scroll off when the gallery is not saved as a carousel', () => {
    const form = createDefaultPageFormState()
    const gallerySection = createDefaultSectionState('gallery')

    form.pageTitle = 'Partners'
    form.urlSlug = 'partners'
    gallerySection.gallery = {
      galleryId: '14',
      viewMode: 'grid',
      showTitleDescription: true,
      autoScrollEnabled: true,
    }
    form.sections = [gallerySection]

    const request = buildSavePageRequest(form)

    expect(request.page_detail?.sections[0]).toMatchObject({
      section_type: 'gallery',
      sort_order: 0,
      gallery: {
        gallery_id: 14,
        view_mode: 'grid',
        show_title_description: true,
        auto_scroll_enabled: false,
      },
    })
  })

  it('defaults legacy gallery option flags when they are missing from section state', () => {
    const form = createDefaultPageFormState()
    const gallerySection = createDefaultSectionState('gallery')

    form.pageTitle = 'Partners'
    form.urlSlug = 'partners'
    gallerySection.gallery.galleryId = '14'
    gallerySection.gallery.viewMode = 'grid'
    Reflect.deleteProperty(gallerySection.gallery as Record<string, unknown>, 'showTitleDescription')
    Reflect.deleteProperty(gallerySection.gallery as Record<string, unknown>, 'autoScrollEnabled')
    form.sections = [gallerySection]

    const request = buildSavePageRequest(form)

    expect(request.page_detail?.sections[0]).toMatchObject({
      section_type: 'gallery',
      gallery: {
        gallery_id: 14,
        view_mode: 'grid',
        show_title_description: true,
        auto_scroll_enabled: false,
      },
    })
  })

  it('hydrates header sections with description, underline, and text alignment', () => {
    const form = buildPageFormStateFromDetail({
      id: 11,
      page_title: 'Mission',
      url_slug: '/mission',
      page_type: 'page',
      parent_id: null,
      status: 'draft',
      hero_image_enabled: false,
      hero_image_url: '',
      hero_image_object_key: '',
      hero_image_fetch_url: '',
      seo_page_title: '',
      seo_page_description: '',
      created_by: null,
      created_by_name: 'Admin',
      modified_by: null,
      modified_by_name: 'Admin',
      last_modified: '2026-05-12T00:00:00Z',
      created_at: '2026-05-12T00:00:00Z',
      updated_at: '2026-05-12T00:00:00Z',
      page_detail: {
        id: 22,
        page_id: 11,
        template_key: 'default',
        schema_version: 1,
        sections: [
          {
            id: 9,
            section_name: 'Mission heading',
            section_type: 'header',
            sort_order: 0,
            is_enabled: true,
            header: {
              main_header_text: 'Our Mission',
              sub_header_text: 'Building stronger communities',
              description: 'Community stories',
              hierarchy: 'h2_section',
              text_align: 'right',
              underline_enabled: true,
            },
            created_at: '2026-05-12T00:00:00Z',
            updated_at: '2026-05-12T00:00:00Z',
          },
        ],
      },
    })

    expect(form.sections[0]).toMatchObject({
      id: 9,
      sectionName: 'Mission heading',
      sectionType: 'header',
      header: {
        mainHeaderText: 'Our Mission',
        subHeaderText: 'Building stronger communities',
        description: 'Community stories',
        hierarchy: 'h2_section',
        textAlign: 'right',
        underlineEnabled: true,
      },
    })
  })

  it('resolves relative asset fetch urls to the API origin', () => {
    const form = buildPageFormStateFromDetail({
      id: 9,
      page_title: 'About',
      url_slug: '/about',
      page_type: 'page',
      parent_id: null,
      status: 'draft',
      hero_image_enabled: true,
      hero_image_url: 'gs://bucket/pages/9/hero.png',
      hero_image_object_key: 'pages/9/hero.png',
      hero_image_fetch_url: '/api/pages/9/hero/content',
      seo_page_title: '',
      seo_page_description: '',
      created_by: null,
      created_by_name: 'Admin',
      modified_by: null,
      modified_by_name: 'Admin',
      last_modified: '2026-05-12T00:00:00Z',
      created_at: '2026-05-12T00:00:00Z',
      updated_at: '2026-05-12T00:00:00Z',
      page_detail: {
        id: 12,
        page_id: 9,
        template_key: 'default',
        schema_version: 1,
        sections: [
          {
            id: 6,
            section_name: 'Downloads',
            section_type: 'document',
            sort_order: 0,
            is_enabled: true,
            documents: {
              items: [
                {
                  id: 17,
                  display_name: 'Board Agenda',
                  description: '',
                  original_file_name: 'agenda.pdf',
                  file_name: 'agenda.pdf',
                  file_url: 'gs://bucket/page-documents/agenda.pdf',
                  fetch_url: '/api/pages/documents/17/content',
                  storage_uri: 'gs://bucket/page-documents/agenda.pdf',
                  gcp_object_key: 'page-documents/agenda.pdf',
                  mime_type: 'application/pdf',
                  file_size: 1024,
                  sort_order: 0,
                  created_at: '2026-05-12T00:00:00Z',
                  updated_at: '2026-05-12T00:00:00Z',
                },
              ],
            },
            created_at: '2026-05-12T00:00:00Z',
            updated_at: '2026-05-12T00:00:00Z',
          },
        ],
      },
    })

    expect(form.existingHeroImageFetchUrl).toBe(
      `${API_BASE_URL}/api/pages/9/hero/content`,
    )
    expect(form.sections[0].documents.items[0].existingFetchUrl).toBe(
      `${API_BASE_URL}/api/pages/documents/17/content`,
    )
  })
})

describe('buildFullPageUrlSlug', () => {
  it('combines the parent slug and child slug into the final path', () => {
    expect(buildFullPageUrlSlug('team', '/about')).toBe('/about/team')
  })
})

describe('parent page validation', () => {
  it('disallows selecting a descendant page as the parent', () => {
    const pageOptions = [
      { id: 1, page_title: 'A', url_slug: '/a', parent_id: null },
      { id: 2, page_title: 'B', url_slug: '/a/b', parent_id: 1 },
    ]

    expect(getDisallowedParentPageIds(1, pageOptions)).toEqual(new Set([1, 2]))
  })

  it('returns a validation error when the selected parent would create a cycle', () => {
    const form = createDefaultPageFormState()
    form.pageTitle = 'A'
    form.urlSlug = 'a'
    form.parentPageId = '2'

    const errors = validatePageForm(form, (key) => key, {
      currentPageId: 1,
      pageOptions: [
        { id: 1, page_title: 'A', url_slug: '/a', parent_id: null },
        { id: 2, page_title: 'B', url_slug: '/a/b', parent_id: 1 },
      ],
    })

    expect(errors.parentPageId).toBe('pages.validation.parentPageCycle')
  })

  it('returns a validation error when a document item has neither a file nor saved storage metadata', () => {
    const form = createDefaultPageFormState()
    const documentSection = createDefaultSectionState('document')

    form.pageTitle = 'Policies'
    form.urlSlug = 'policies'
    documentSection.documents.items = [createDefaultDocumentState()]
    form.sections = [documentSection]

    const errors = validatePageForm(form, (key) => key)

    expect(errors.sections).toBe('pages.validation.documentFileRequired')
  })

  it('returns a validation error when more than one header section uses h1 hero', () => {
    const form = createDefaultPageFormState()
    const firstHeader = createDefaultSectionState('header')
    const secondHeader = createDefaultSectionState('header')

    form.pageTitle = 'About'
    form.urlSlug = 'about'
    firstHeader.header.hierarchy = 'h1_hero'
    secondHeader.header.hierarchy = 'h1_hero'
    form.sections = [firstHeader, secondHeader]

    const errors = validatePageForm(form, (key) => key)

    expect(errors.sections).toBe('pages.validation.singleHeroHeader')
  })

  it('returns field-level validation errors for required header, typography, and quote content', () => {
    const headerSection = createDefaultSectionState('header')
    const typographySection = createDefaultSectionState('typography')
    const quoteSection = createDefaultSectionState('quote')

    headerSection.header.mainHeaderText = '   '
    typographySection.typography.htmlContent = '<p><br></p>'
    quoteSection.quote.quoteContent = '   '

    const errors = validatePageSectionFields(
      [headerSection, typographySection, quoteSection],
      (key) => key,
    )

    expect(errors[headerSection.clientId]?.mainHeaderText).toBe(
      'pages.validation.mainHeaderTextRequired',
    )
    expect(errors[typographySection.clientId]?.htmlContent).toBe(
      'pages.validation.typographyContentRequired',
    )
    expect(errors[quoteSection.clientId]?.quoteContent).toBe(
      'pages.validation.quoteContentRequired',
    )
  })

  it('returns field-level validation errors for required document and CTA fields', () => {
    const documentSection = createDefaultSectionState('document')
    const quoteSection = createDefaultSectionState('cta_banner')

    quoteSection.ctaBanner.bannerHeading = '   '
    quoteSection.ctaBanner.buttonText = '   '
    quoteSection.ctaBanner.buttonUrl = '   '

    const sectionErrors = validatePageSectionFields(
      [documentSection, quoteSection],
      (key) => key,
    )

    expect(sectionErrors[documentSection.clientId]?.documents).toBe(
      'pages.validation.documentsRequired',
    )
    expect(sectionErrors[quoteSection.clientId]?.bannerHeading).toBe(
      'pages.validation.bannerHeadingRequired',
    )
    expect(sectionErrors[quoteSection.clientId]?.buttonText).toBe(
      'pages.validation.buttonTextRequired',
    )
    expect(sectionErrors[quoteSection.clientId]?.buttonUrl).toBe(
      'pages.validation.buttonUrlRequired',
    )
  })

  it('returns field-level validation errors for required document display names', () => {
    const documentSection = createDefaultSectionState('document')
    const document = createDefaultDocumentState(
      new File(['content'], 'agenda.pdf', { type: 'application/pdf' }),
    )

    document.displayName = '   '
    documentSection.documents.items = [document]

    const errors = validatePageDocumentFields([documentSection], (key) => key)

    expect(errors[document.clientId]?.displayName).toBe(
      'pages.validation.documentDisplayNameRequired',
    )
  })

  it('reorders sections before or after the drop target', () => {
    const first = createDefaultSectionState('header')
    const middle = createDefaultSectionState('typography')
    const last = createDefaultSectionState('cta_banner')
    const sections = [first, middle, last]

    expect(
      reorderPageSections(sections, last.clientId, first.clientId, 'before').map(
        (section) => section.clientId,
      ),
    ).toEqual([last.clientId, first.clientId, middle.clientId])

    expect(
      reorderPageSections(sections, first.clientId, last.clientId, 'after').map(
        (section) => section.clientId,
      ),
    ).toEqual([middle.clientId, last.clientId, first.clientId])
  })
})
