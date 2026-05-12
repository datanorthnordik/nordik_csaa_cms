import { describe, expect, it } from 'vitest'
import {
  buildSavePageRequest,
  createDefaultPageFormState,
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
})
