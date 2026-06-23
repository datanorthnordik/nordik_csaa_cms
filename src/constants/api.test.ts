import { afterEach, describe, expect, it, vi } from 'vitest'
import { API_BASE_URL, API_ROUTES, resolveApiBaseUrl } from './api'

describe('API_ROUTES', () => {
  afterEach(() => {
    window.__APP_CONFIG__ = undefined
    vi.unstubAllEnvs()
  })

  it('falls back to the default API base URL', () => {
    expect(API_BASE_URL).toBe(
      'https://nordikcsaaapi-724838782318.us-west1.run.app',
    )
  })

  it('prefers the Cloud Run runtime API base URL when present', () => {
    window.__APP_CONFIG__ = {
      API_BASE_URL: 'https://api.example.com/',
    }

    expect(resolveApiBaseUrl()).toBe('https://api.example.com')
  })

  it('falls back to the Vite API base URL for local builds', () => {
    vi.stubEnv('VITE_API_BASE_URL', 'https://local.example.com/')

    expect(resolveApiBaseUrl()).toBe('https://local.example.com')
  })

  it('builds route paths for the CMS APIs', () => {
    expect(API_ROUTES.login).toBe('/api/user/login')
    expect(API_ROUTES.signup).toBe('/api/user/signup')
    expect(API_ROUTES.refresh).toBe('/api/user/refresh')
    expect(API_ROUTES.forgotPassword).toBe('/api/user/forgot-password')
    expect(API_ROUTES.resetPassword).toBe('/api/user/reset-password')
    expect(API_ROUTES.events).toBe('/api/events')
    expect(API_ROUTES.eventLocations).toBe('/api/events/locations')
    expect(API_ROUTES.eventGalleries).toBe('/api/events/galleries')
    expect(API_ROUTES.eventById(42)).toBe('/api/events/42')
    expect(API_ROUTES.eventMediaById(42, 7)).toBe('/api/events/42/media/7/content')
    expect(API_ROUTES.eventDocumentById(42)).toBe('/api/events/42/document')
    expect(API_ROUTES.eventPhotoById(42)).toBe('/api/events/42/photo')
    expect(API_ROUTES.resources).toBe('/api/resources')
    expect(API_ROUTES.resourceById(8)).toBe('/api/resources/8')
    expect(API_ROUTES.resourceContentById(8)).toBe('/api/resources/8/content')
    expect(API_ROUTES.knowledgeCenterSubmissions).toBe(
      '/api/knowledge-center/submissions',
    )
    expect(API_ROUTES.knowledgeCenterSubmissionById(8)).toBe(
      '/api/knowledge-center/submissions/8',
    )
    expect(API_ROUTES.knowledgeCenterSubmissionComplete(8)).toBe(
      '/api/knowledge-center/submissions/8/complete',
    )
    expect(API_ROUTES.memorial).toBe('/api/memorial')
    expect(API_ROUTES.memorialById(8)).toBe('/api/memorial/8')
    expect(API_ROUTES.memorialPortraitById(8)).toBe('/api/memorial/8/portrait/content')
    expect(API_ROUTES.memorialGalleryImageContentById(8, 4)).toBe(
      '/api/memorial/8/gallery/4/content',
    )
    expect(API_ROUTES.pages).toBe('/api/pages')
    expect(API_ROUTES.pageById(12)).toBe('/api/pages/12')
    expect(API_ROUTES.pageHeroById(12)).toBe('/api/pages/12/hero/content')
    expect(API_ROUTES.menuByKey('main')).toBe('/api/menus/main')
    expect(API_ROUTES.menuPageOptionsByKey('main')).toBe('/api/menus/main/page-options')
    expect(API_ROUTES.galleries).toBe('/api/galleries')
    expect(API_ROUTES.galleryById(5)).toBe('/api/galleries/5')
    expect(API_ROUTES.galleryCoverById(5)).toBe('/api/galleries/5/cover/content')
    expect(API_ROUTES.galleryImagesById(5)).toBe('/api/galleries/5/images')
    expect(API_ROUTES.galleryImageById(5, 9)).toBe('/api/galleries/5/images/9')
    expect(API_ROUTES.galleryImageContentById(5, 9)).toBe(
      '/api/galleries/5/images/9/content',
    )
    expect(API_ROUTES.galleryImageOrderById(5)).toBe('/api/galleries/5/images/order')
  })
})
