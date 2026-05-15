import { describe, expect, it } from 'vitest'
import {
  buildFullPageUrlSlug,
  buildPageFormStateFromDetail,
  buildSavePageRequest,
  createDefaultDocumentState,
  createDefaultPageFormState,
  createDefaultSectionState,
  getDisallowedParentPageIds,
  validatePageForm,
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
})
