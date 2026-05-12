import type {
<<<<<<< HEAD
=======
  PageParentOption,
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
  PageDetailResponse,
  PageStatus,
  PageUploadInput,
  SavePagePayload,
  SavePageRequest,
} from '../api/pagesApi'
<<<<<<< HEAD
=======
import {
  resolvePageParentId as getPageParentId,
  resolvePageParentSlug as getPageParentSlug,
} from '../api/pagesApi'
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578

export type PageFormState = {
  pageTitle: string
  urlSlug: string
<<<<<<< HEAD
=======
  parentPageId: string
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
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
<<<<<<< HEAD
=======
    parentPageId: '',
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
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

<<<<<<< HEAD
export function buildPageFormStateFromDetail(detail: PageDetailResponse): PageFormState {
  return {
    pageTitle: detail.page_title ?? '',
    urlSlug: stripLeadingSlash(detail.url_slug ?? ''),
=======
export function buildPageFormStateFromDetail(
  detail: PageDetailResponse,
  parentPageSlug = '',
): PageFormState {
  const resolvedParentId = getPageParentId(detail)
  const effectiveParentSlug = parentPageSlug || getPageParentSlug(detail)

  return {
    pageTitle: detail.page_title ?? '',
    urlSlug: stripParentSlugPrefix(detail.url_slug ?? '', effectiveParentSlug),
    parentPageId: resolvedParentId ? String(resolvedParentId) : '',
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
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

<<<<<<< HEAD
export function validatePageForm(
  form: PageFormState,
  t: (key: string) => string,
=======
export function buildFullPageUrlSlug(value: string, parentPageSlug = '') {
  const normalizedValue = normalizePageSlugInput(value)
  const normalizedParentPath = normalizePagePath(parentPageSlug)

  if (!normalizedValue) {
    return normalizedParentPath
  }

  if (!normalizedParentPath) {
    return `/${normalizedValue}`
  }

  return `${normalizedParentPath}/${normalizedValue}`.replace(/\/+/g, '/')
}

export function stripParentSlugPrefix(value: string, parentPageSlug = '') {
  const normalizedValue = normalizePagePath(value)
  const normalizedParentPath = normalizePagePath(parentPageSlug)

  if (!normalizedParentPath) {
    return stripLeadingSlash(normalizedValue)
  }

  if (normalizedValue === normalizedParentPath) {
    return ''
  }

  const parentPrefix = `${normalizedParentPath}/`
  if (normalizedValue.startsWith(parentPrefix)) {
    return stripLeadingSlash(normalizedValue.slice(parentPrefix.length))
  }

  return stripLeadingSlash(normalizedValue)
}

export function getDisallowedParentPageIds(
  currentPageId: number,
  pageOptions: PageParentOption[],
) {
  const pageMap = new Map(pageOptions.map((page) => [page.id, page]))
  const disallowedIds = new Set<number>([currentPageId])

  for (const page of pageOptions) {
    if (wouldCreatePageCycle(currentPageId, page.id, pageMap)) {
      disallowedIds.add(page.id)
    }
  }

  return disallowedIds
}

export function validatePageForm(
  form: PageFormState,
  t: (key: string) => string,
  options: {
    currentPageId?: number | null
    pageOptions?: PageParentOption[]
  } = {},
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
): PageFormErrors {
  const errors: PageFormErrors = {}

  if (!form.pageTitle.trim()) {
    errors.pageTitle = t('pages.validation.pageTitleRequired')
  }

  if (!normalizePageSlugInput(form.urlSlug)) {
    errors.urlSlug = t('pages.validation.urlSlugRequired')
  }

<<<<<<< HEAD
  return errors
}

export function buildSavePagePayload(form: PageFormState): SavePagePayload {
=======
  if (
    options.currentPageId &&
    form.parentPageId &&
    options.pageOptions?.length &&
    getDisallowedParentPageIds(options.currentPageId, options.pageOptions).has(
      Number.parseInt(form.parentPageId, 10),
    )
  ) {
    errors.parentPageId = t('pages.validation.parentPageCycle')
  }

  return errors
}

export function buildSavePagePayload(
  form: PageFormState,
  parentPageSlug = '',
): SavePagePayload {
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
  let heroImage: PageUploadInput | undefined

  if (form.heroImageEnabled && form.heroImageFile) {
    heroImage = {
      file_name: form.heroImageFile.name,
      mime_type: form.heroImageFile.type || 'application/octet-stream',
    }
  }

  return {
    page_title: form.pageTitle.trim(),
<<<<<<< HEAD
    url_slug: toPageUrlSlug(form.urlSlug),
=======
    url_slug: buildFullPageUrlSlug(form.urlSlug, parentPageSlug),
    parent_page_id: form.parentPageId ? Number.parseInt(form.parentPageId, 10) : null,
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
    status: form.status,
    hero_image_enabled: form.heroImageEnabled,
    hero_image: heroImage,
    remove_hero_image: !form.heroImageEnabled || form.removeHeroImage,
    seo_page_title: form.seoPageTitle.trim(),
    seo_page_description: form.seoPageDescription.trim(),
  }
}

<<<<<<< HEAD
export function buildSavePageRequest(form: PageFormState): SavePageRequest {
  const payload = buildSavePagePayload(form)
=======
export function buildSavePageRequest(
  form: PageFormState,
  parentPageSlug = '',
): SavePageRequest {
  const payload = buildSavePagePayload(form, parentPageSlug)
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578

  return {
    ...payload,
    ...(form.heroImageEnabled && form.heroImageFile
      ? {
          heroImageFile: form.heroImageFile,
        }
      : {}),
  }
}

<<<<<<< HEAD
=======
function wouldCreatePageCycle(
  currentPageId: number,
  candidateParentId: number,
  pageMap: Map<number, PageParentOption>,
) {
  let nextPageId: number | null = candidateParentId
  const visited = new Set<number>()

  while (nextPageId !== null) {
    if (nextPageId === currentPageId) {
      return true
    }

    if (visited.has(nextPageId)) {
      return true
    }

    visited.add(nextPageId)
    nextPageId = pageMap.get(nextPageId)?.parent_page_id ?? null
  }

  return false
}

function normalizePagePath(value: string) {
  const normalized = normalizePageSlugInput(value)
  return normalized ? `/${normalized}` : ''
}

>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
function stripLeadingSlash(value: string) {
  return value.replace(/^\/+/, '')
}
