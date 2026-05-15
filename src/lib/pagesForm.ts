import type {
  PageParentOption,
  PageDetailResponse,
  PageStatus,
  PageUploadInput,
  SavePagePayload,
  SavePageRequest,
} from '../api/pagesApi'

export type PageFormState = {
  pageTitle: string
  urlSlug: string
  parentPageId: string
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
    parentPageId: '',
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

export function buildPageFormStateFromDetail(
  detail: PageDetailResponse,
  options: {
    parentPageSlug?: string
  } = {},
): PageFormState {
  const resolvedParentId =
    typeof detail.parent_id === 'number' ? detail.parent_id : null
  const effectiveParentSlug = options.parentPageSlug ?? ''

  return {
    pageTitle: detail.page_title ?? '',
    urlSlug: stripParentSlugPrefix(detail.url_slug ?? '', effectiveParentSlug),
    parentPageId: resolvedParentId ? String(resolvedParentId) : '',
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
): PageFormErrors {
  const errors: PageFormErrors = {}

  if (!form.pageTitle.trim()) {
    errors.pageTitle = t('pages.validation.pageTitleRequired')
  }

  if (!normalizePageSlugInput(form.urlSlug)) {
    errors.urlSlug = t('pages.validation.urlSlugRequired')
  }

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
  let heroImage: PageUploadInput | undefined

  if (form.heroImageEnabled && form.heroImageFile) {
    heroImage = {
      file_name: form.heroImageFile.name,
      mime_type: form.heroImageFile.type || 'application/octet-stream',
    }
  }

  return {
    page_title: form.pageTitle.trim(),
    url_slug: buildFullPageUrlSlug(form.urlSlug, parentPageSlug),
    parent_id: form.parentPageId ? Number.parseInt(form.parentPageId, 10) : null,
    status: form.status,
    hero_image_enabled: form.heroImageEnabled,
    hero_image: heroImage,
    remove_hero_image: !form.heroImageEnabled || form.removeHeroImage,
    seo_page_title: form.seoPageTitle.trim(),
    seo_page_description: form.seoPageDescription.trim(),
  }
}

export function buildSavePageRequest(
  form: PageFormState,
  parentPageSlug = '',
): SavePageRequest {
  const payload = buildSavePagePayload(form, parentPageSlug)

  return {
    ...payload,
    ...(form.heroImageEnabled && form.heroImageFile
      ? {
          heroImageFile: form.heroImageFile,
        }
      : {}),
  }
}

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
    nextPageId = pageMap.get(nextPageId)?.parent_id ?? null
  }

  return false
}

function normalizePagePath(value: string) {
  const normalized = normalizePageSlugInput(value)
  return normalized ? `/${normalized}` : ''
}

function stripLeadingSlash(value: string) {
  return value.replace(/^\/+/, '')
}
