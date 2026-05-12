import type {
  PageDetailResponse,
  PageStatus,
  SavePagePayload,
} from '../api/pagesApi'
import { fileToBase64 } from './fileUpload'

export type PageFormState = {
  pageTitle: string
  urlSlug: string
  status: PageStatus
  heroImageEnabled: boolean
  heroImageFile: File | null
  removeHeroImage: boolean
  existingHeroImageFetchUrl: string
  existingHeroImageStorageUrl: string
  seoPageTitle: string
  seoPageDescription: string
}

export type PageFormErrors = Partial<Record<keyof PageFormState, string>>

export function createDefaultPageFormState(): PageFormState {
  return {
    pageTitle: '',
    urlSlug: '',
    status: 'draft',
    heroImageEnabled: false,
    heroImageFile: null,
    removeHeroImage: false,
    existingHeroImageFetchUrl: '',
    existingHeroImageStorageUrl: '',
    seoPageTitle: '',
    seoPageDescription: '',
  }
}

export function buildPageFormStateFromDetail(detail: PageDetailResponse): PageFormState {
  return {
    pageTitle: detail.page_title ?? '',
    urlSlug: stripLeadingSlash(detail.url_slug ?? ''),
    status: detail.status,
    heroImageEnabled: detail.hero_image_enabled,
    heroImageFile: null,
    removeHeroImage: false,
    existingHeroImageFetchUrl: detail.hero_image_fetch_url ?? '',
    existingHeroImageStorageUrl: detail.hero_image_url ?? '',
    seoPageTitle: detail.seo_page_title ?? '',
    seoPageDescription: detail.seo_page_description ?? '',
  }
}

export function normalizePageSlugInput(value: string) {
  return value
    .trim()
    .toLowerCase()
    .replace(/\\/g, '/')
    .replace(/\s+/g, '-')
    .replace(/[^a-z0-9/-]+/g, '')
    .replace(/\/+/g, '/')
    .replace(/^\/+/, '')
    .replace(/\/+$/, '')
}

export function toPageUrlSlug(value: string) {
  const normalized = normalizePageSlugInput(value)
  return normalized ? `/${normalized}` : ''
}

export function validatePageForm(
  form: PageFormState,
  t: (key: string) => string,
): PageFormErrors {
  const errors: PageFormErrors = {}

  if (!form.pageTitle.trim()) {
    errors.pageTitle = t('pages.validation.pageTitleRequired')
  }

  if (!normalizePageSlugInput(form.urlSlug)) {
    errors.urlSlug = t('pages.validation.urlSlugRequired')
  }

  return errors
}

export async function buildSavePagePayload(form: PageFormState): Promise<SavePagePayload> {
  let heroImage = undefined

  if (form.heroImageEnabled && form.heroImageFile) {
    heroImage = {
      file_name: form.heroImageFile.name,
      mime_type: form.heroImageFile.type,
      data_base64: await fileToBase64(form.heroImageFile),
    }
  }

  return {
    page_title: form.pageTitle.trim(),
    url_slug: toPageUrlSlug(form.urlSlug),
    status: form.status,
    hero_image_enabled: form.heroImageEnabled,
    hero_image: heroImage,
    remove_hero_image: !form.heroImageEnabled || form.removeHeroImage,
    seo_page_title: form.seoPageTitle.trim(),
    seo_page_description: form.seoPageDescription.trim(),
  }
}

function stripLeadingSlash(value: string) {
  return value.replace(/^\/+/, '')
}
