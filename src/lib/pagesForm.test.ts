import { describe, expect, it } from 'vitest'
import type { PageDetailResponse } from '../api/pagesApi'
import {
  buildPageFormStateFromDetail,
  buildSavePageRequest,
  createDefaultPageFormState,
  normalizePageSlugInput,
  toPageUrlSlug,
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
    expect(request.hero_image).not.toHaveProperty('data_base64')
  })

  it('omits heroImageFile from the request when no file is selected', () => {
    const form = createDefaultPageFormState()
    form.pageTitle = 'About'
    form.urlSlug = 'about'

    const request = buildSavePageRequest(form)

    expect(request.heroImageFile).toBeUndefined()
    expect(request.hero_image).toBeUndefined()
  })

  it('sets remove_hero_image to true when hero image is disabled', () => {
    const form = createDefaultPageFormState()
    form.heroImageEnabled = false

    const request = buildSavePageRequest(form)

    expect(request.remove_hero_image).toBe(true)
  })
})

describe('validatePageForm', () => {
  it('reports errors when title and slug are both empty', () => {
    const form = createDefaultPageFormState()
    const errors = validatePageForm(form, (key) => key)

    expect(errors.pageTitle).toBeDefined()
    expect(errors.urlSlug).toBeDefined()
  })

  it('reports no errors when required fields are provided', () => {
    const form = createDefaultPageFormState()
    form.pageTitle = 'About Us'
    form.urlSlug = 'about-us'

    const errors = validatePageForm(form, (key) => key)

    expect(errors.pageTitle).toBeUndefined()
    expect(errors.urlSlug).toBeUndefined()
  })

  it('reports a title error but not a slug error when only the title is missing', () => {
    const form = createDefaultPageFormState()
    form.urlSlug = 'about-us'

    const errors = validatePageForm(form, (key) => key)

    expect(errors.pageTitle).toBeDefined()
    expect(errors.urlSlug).toBeUndefined()
  })

  it('reports a slug error when the slug normalises to an empty string', () => {
    const form = createDefaultPageFormState()
    form.pageTitle = 'Valid Title'
    form.urlSlug = '   '

    const errors = validatePageForm(form, (key) => key)

    expect(errors.urlSlug).toBeDefined()
    expect(errors.pageTitle).toBeUndefined()
  })
})

describe('normalizePageSlugInput', () => {
  it('lowercases and trims the input', () => {
    expect(normalizePageSlugInput('  About Us  ')).toBe('about-us')
  })

  it('replaces spaces with hyphens', () => {
    expect(normalizePageSlugInput('my cool page')).toBe('my-cool-page')
  })

  it('removes invalid characters', () => {
    expect(normalizePageSlugInput('hello!@world')).toBe('helloworld')
  })

  it('preserves forward slashes for nested slugs', () => {
    expect(normalizePageSlugInput('events/annual')).toBe('events/annual')
  })

  it('collapses multiple slashes into one', () => {
    expect(normalizePageSlugInput('a//b///c')).toBe('a/b/c')
  })

  it('strips leading and trailing slashes', () => {
    expect(normalizePageSlugInput('/about/')).toBe('about')
  })

  it('returns an empty string for whitespace-only input', () => {
    expect(normalizePageSlugInput('   ')).toBe('')
  })

  it('converts backslashes to forward slashes', () => {
    expect(normalizePageSlugInput('section\\subsection')).toBe('section/subsection')
  })
})

describe('toPageUrlSlug', () => {
  it('prepends a slash to a normalised slug', () => {
    expect(toPageUrlSlug('about-us')).toBe('/about-us')
  })

  it('returns an empty string for empty input', () => {
    expect(toPageUrlSlug('')).toBe('')
  })

  it('normalises the input before prepending the slash', () => {
    expect(toPageUrlSlug('  My Page  ')).toBe('/my-page')
  })

  it('returns empty string when input normalises to empty', () => {
    expect(toPageUrlSlug('!!!')).toBe('')
  })
})

describe('buildPageFormStateFromDetail', () => {
  function makePageDetail(
    overrides: Partial<PageDetailResponse> = {},
  ): PageDetailResponse {
    return {
      id: 5,
      page_title: 'About Us',
      url_slug: '/about-us',
      status: 'published',
      hero_image_enabled: true,
      hero_image_url: 'https://storage.example.com/hero.png',
      hero_image_object_key: 'pages/5/hero.png',
      hero_image_fetch_url: '/api/pages/5/hero/content',
      seo_page_title: 'About Us – Company',
      seo_page_description: 'Learn more about us',
      created_by: 1,
      created_by_name: 'Admin',
      modified_by: 2,
      modified_by_name: 'Editor',
      last_modified: '2026-05-01T10:00:00Z',
      created_at: '2026-04-01T10:00:00Z',
      updated_at: '2026-05-01T10:00:00Z',
      ...overrides,
    }
  }

  it('hydrates all fields from the API response', () => {
    const form = buildPageFormStateFromDetail(makePageDetail())

    expect(form.pageTitle).toBe('About Us')
    expect(form.urlSlug).toBe('about-us')
    expect(form.status).toBe('published')
    expect(form.heroImageEnabled).toBe(true)
    expect(form.existingHeroImageFetchUrl).toBe('/api/pages/5/hero/content')
    expect(form.existingHeroImageStorageUrl).toBe(
      'https://storage.example.com/hero.png',
    )
    expect(form.seoPageTitle).toBe('About Us – Company')
    expect(form.seoPageDescription).toBe('Learn more about us')
  })

  it('initialises mutable file fields to their defaults', () => {
    const form = buildPageFormStateFromDetail(makePageDetail())

    expect(form.heroImageFile).toBeNull()
    expect(form.removeHeroImage).toBe(false)
  })

  it('strips the leading slash from url_slug', () => {
    const form = buildPageFormStateFromDetail(
      makePageDetail({ url_slug: '/events/annual' }),
    )
    expect(form.urlSlug).toBe('events/annual')
  })

  it('handles a draft page with hero image disabled', () => {
    const form = buildPageFormStateFromDetail(
      makePageDetail({ status: 'draft', hero_image_enabled: false }),
    )
    expect(form.status).toBe('draft')
    expect(form.heroImageEnabled).toBe(false)
  })
})
