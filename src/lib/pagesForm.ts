import type {
  PageCTABannerSectionResponse,
  PageCTABannerImageMultipartFile,
  PageContentDetailResponse,
  PageDetailResponse,
  PageDocumentMultipartFile,
  PageDocumentResponse,
  PageGallerySectionResponse,
  PageGalleryViewMode,
  PageHeaderHierarchy,
  PageHeaderSectionResponse,
  PageParentOption,
  PageQuoteSectionResponse,
  PageSectionResponse,
  PageSectionType,
  PageVideoSectionResponse,
  PageTypographySectionResponse,
  PageTypographyTextAlign,
  PageStatus,
  SavePageCTABannerSectionPayload,
  SavePageDetailPayload,
  SavePageDocumentPayload,
  SavePageDocumentsSectionPayload,
  SavePageGallerySectionPayload,
  SavePageHeaderSectionPayload,
  SavePagePayload,
  SavePageQuoteSectionPayload,
  SavePageRequest,
  SavePageSectionPayload,
  SavePageVideoSectionPayload,
  SavePageTypographySectionPayload,
  PageUploadInput,
} from '../api/pagesApi'
import { API_BASE_URL } from '../constants/api'

export type PageSectionDocumentState = {
  clientId: string
  id?: number
  displayName: string
  description: string
  originalFileName: string
  existingFetchUrl: string
  existingStorageUri: string
  existingMimeType: string
  existingFileSize?: number
  existingObjectKey: string
  file: File | null
}

export type PageSectionState = {
  clientId: string
  id?: number
  sectionName: string
  sectionType: PageSectionType
  isEnabled: boolean
  isCollapsed: boolean
  settings: Record<string, unknown>
  header: {
    mainHeaderText: string
    subHeaderText: string
    description: string
    hierarchy: PageHeaderHierarchy
    textAlign: PageTypographyTextAlign
    underlineEnabled: boolean
  }
  typography: {
    htmlContent: string
    textContent: string
    textAlign: PageTypographyTextAlign
  }
  gallery: {
    galleryId: string
    viewMode: PageGalleryViewMode
    showTitleDescription: boolean
    autoScrollEnabled: boolean
  }
  video: {
    videoPackageId: string
  }
  quote: {
    quoteContent: string
    attribution: string
  }
  ctaBanner: {
    bannerHeading: string
    bannerMessage: string
    buttonText: string
    buttonUrl: string
    openInNewTab: boolean
    imageFile: File | null
    existingImageFetchUrl: string
    existingImageStorageUrl: string
    existingImageObjectKey: string
  }
  documents: {
    items: PageSectionDocumentState[]
  }
}

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
  templateKey: string
  sections: PageSectionState[]
}

export type PageFormErrors = Partial<
  Record<'pageTitle' | 'urlSlug' | 'parentPageId' | 'heroImageFile' | 'sections', string>
>

export type PageSectionFieldErrors = Partial<
  Record<
    | 'mainHeaderText'
    | 'htmlContent'
    | 'videoPackage'
    | 'quoteContent'
    | 'documents'
    | 'bannerHeading'
    | 'buttonText'
    | 'buttonUrl',
    string
  >
>

export type PageSectionValidationErrors = Record<string, PageSectionFieldErrors>

export type PageDocumentFieldErrors = Partial<Record<'displayName', string>>

export type PageDocumentValidationErrors = Record<string, PageDocumentFieldErrors>

export type PageSectionDropPlacement = 'before' | 'after'

export function createPageEditorId(prefix: string) {
  if (typeof globalThis.crypto?.randomUUID === 'function') {
    return `${prefix}-${globalThis.crypto.randomUUID()}`
  }
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}`
}

export function createDefaultSectionState(sectionType: PageSectionType): PageSectionState {
  return {
    clientId: createPageEditorId('page-section'),
    sectionName: defaultSectionName(sectionType),
    sectionType,
    isEnabled: true,
    isCollapsed: false,
    settings: {},
    header: {
      mainHeaderText: '',
      subHeaderText: '',
      description: '',
      hierarchy: 'h1_hero',
      textAlign: 'left',
      underlineEnabled: false,
    },
    typography: {
      htmlContent: '',
      textContent: '',
      textAlign: 'left',
    },
    gallery: {
      galleryId: '',
      viewMode: 'grid',
      showTitleDescription: true,
      autoScrollEnabled: false,
    },
    video: {
      videoPackageId: '',
    },
    quote: {
      quoteContent: '',
      attribution: '',
    },
    ctaBanner: {
      bannerHeading: '',
      bannerMessage: '',
      buttonText: '',
      buttonUrl: '',
      openInNewTab: true,
      imageFile: null,
      existingImageFetchUrl: '',
      existingImageStorageUrl: '',
      existingImageObjectKey: '',
    },
    documents: {
      items: [],
    },
  }
}

export function createDefaultDocumentState(file?: File): PageSectionDocumentState {
  return {
    clientId: createPageEditorId('page-document'),
    displayName: file ? suggestDocumentDisplayName(file.name) : '',
    description: '',
    originalFileName: file?.name ?? '',
    existingFetchUrl: '',
    existingStorageUri: '',
    existingMimeType: file?.type ?? '',
    existingObjectKey: '',
    file: file ?? null,
  }
}

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
    templateKey: 'default',
    sections: [],
  }
}

export function resolvePageAssetUrl(value: string) {
  const trimmed = value.trim()
  if (!trimmed) {
    return ''
  }

  if (/^[a-z][a-z0-9+.-]*:/i.test(trimmed) || trimmed.startsWith('//')) {
    return trimmed
  }

  try {
    return new URL(trimmed, `${API_BASE_URL}/`).toString()
  } catch {
    return trimmed
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
  const pageDetail = normalizePageDetail(detail.page_detail)

  return {
    pageTitle: detail.page_title ?? '',
    urlSlug: stripParentSlugPrefix(detail.url_slug ?? '', effectiveParentSlug),
    parentPageId: resolvedParentId ? String(resolvedParentId) : '',
    status: detail.status,
    heroImageEnabled: detail.hero_image_enabled,
    heroImageFile: null,
    removeHeroImage: false,
    existingHeroImageFetchUrl: resolvePageAssetUrl(detail.hero_image_fetch_url ?? ''),
    existingHeroImageStorageUrl: detail.hero_image_url ?? '',
    seoPageTitle: detail.seo_page_title ?? '',
    seoPageDescription: detail.seo_page_description ?? '',
    templateKey: pageDetail.template_key,
    sections: pageDetail.sections
      .slice()
      .sort((left, right) => left.sort_order - right.sort_order)
      .map(mapSectionResponseToState),
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

  if (countHeroHeaderSections(form.sections) > 1) {
    errors.sections = t('pages.validation.singleHeroHeader')
  }

  const hasInvalidDocument = form.sections.some((section) =>
    section.sectionType === 'document' &&
    section.documents.items.some(
      (item) => !item.file && !item.existingStorageUri.trim() && !item.existingObjectKey.trim(),
    ),
  )
  if (!errors.sections && hasInvalidDocument) {
    errors.sections = t('pages.validation.documentFileRequired')
  }

  return errors
}

export function validatePageSectionFields(
  sections: PageSectionState[],
  t: (key: string) => string,
): PageSectionValidationErrors {
  const errors: PageSectionValidationErrors = {}

  for (const section of sections) {
    switch (section.sectionType) {
      case 'header':
        if (!section.header.mainHeaderText.trim()) {
          errors[section.clientId] = {
            ...errors[section.clientId],
            mainHeaderText: t('pages.validation.mainHeaderTextRequired'),
          }
        }
        break

      case 'typography':
        if (isTypographyContentEmpty(section.typography.htmlContent)) {
          errors[section.clientId] = {
            ...errors[section.clientId],
            htmlContent: t('pages.validation.typographyContentRequired'),
          }
        }
        break

      case 'quote':
        if (!section.quote.quoteContent.trim()) {
          errors[section.clientId] = {
            ...errors[section.clientId],
            quoteContent: t('pages.validation.quoteContentRequired'),
          }
        }
        break

      case 'video':
        if (!section.video.videoPackageId.trim()) {
          errors[section.clientId] = {
            ...errors[section.clientId],
            videoPackage: t('pages.validation.videoPackageRequired'),
          }
        }
        break

      case 'document': {
        const hasDocumentItems = section.documents.items.length > 0
        const hasInvalidDocumentFile = section.documents.items.some(
          (item) =>
            !item.file && !item.existingStorageUri.trim() && !item.existingObjectKey.trim(),
        )

        if (!hasDocumentItems) {
          errors[section.clientId] = {
            ...errors[section.clientId],
            documents: t('pages.validation.documentsRequired'),
          }
        } else if (hasInvalidDocumentFile) {
          errors[section.clientId] = {
            ...errors[section.clientId],
            documents: t('pages.validation.documentFileRequired'),
          }
        }
        break
      }

      case 'cta_banner':
        if (!section.ctaBanner.bannerHeading.trim()) {
          errors[section.clientId] = {
            ...errors[section.clientId],
            bannerHeading: t('pages.validation.bannerHeadingRequired'),
          }
        }

        if (!section.ctaBanner.buttonText.trim()) {
          errors[section.clientId] = {
            ...errors[section.clientId],
            buttonText: t('pages.validation.buttonTextRequired'),
          }
        }

        if (!section.ctaBanner.buttonUrl.trim()) {
          errors[section.clientId] = {
            ...errors[section.clientId],
            buttonUrl: t('pages.validation.buttonUrlRequired'),
          }
        }
        break

      default:
        break
    }
  }

  return errors
}

export function getPageSectionValidationMessages(
  errors: PageSectionValidationErrors,
) {
  return Object.values(errors).flatMap((sectionErrors) =>
    Object.values(sectionErrors).filter((message): message is string => Boolean(message)),
  )
}

export function validatePageDocumentFields(
  sections: PageSectionState[],
  t: (key: string) => string,
): PageDocumentValidationErrors {
  const errors: PageDocumentValidationErrors = {}

  sections.forEach((section) => {
    if (section.sectionType !== 'document') {
      return
    }

    section.documents.items.forEach((document) => {
      if (!document.displayName.trim()) {
        errors[document.clientId] = {
          displayName: t('pages.validation.documentDisplayNameRequired'),
        }
      }
    })
  })

  return errors
}

export function getPageDocumentValidationMessages(
  errors: PageDocumentValidationErrors,
) {
  return Object.values(errors).flatMap((documentErrors) =>
    Object.values(documentErrors).filter((message): message is string => Boolean(message)),
  )
}

export function countHeroHeaderSections(sections: PageSectionState[]) {
  return sections.filter(
    (section) => section.sectionType === 'header' && section.header.hierarchy === 'h1_hero',
  ).length
}

export function reorderPageSections(
  sections: PageSectionState[],
  draggedClientId: string,
  targetClientId: string,
  placement: PageSectionDropPlacement = 'before',
) {
  const fromIndex = sections.findIndex((section) => section.clientId === draggedClientId)
  if (fromIndex < 0) {
    return sections
  }

  const next = sections.slice()
  const [moved] = next.splice(fromIndex, 1)
  const targetIndex = next.findIndex((section) => section.clientId === targetClientId)
  if (!moved || targetIndex < 0) {
    return sections
  }

  const insertIndex = placement === 'after' ? targetIndex + 1 : targetIndex
  next.splice(insertIndex, 0, moved)
  return next
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
    page_detail: buildSavePageDetailPayload(form),
  }
}

export function buildSavePageRequest(
  form: PageFormState,
  parentPageSlug = '',
): SavePageRequest {
  const payload = buildSavePagePayload(form, parentPageSlug)
  const request: SavePageRequest = {
    ...payload,
  }

  if (form.heroImageEnabled && form.heroImageFile) {
    request.heroImageFile = form.heroImageFile
  }

  const documentFiles = collectDocumentFiles(form.sections)
  if (documentFiles.length > 0) {
    request.documentFiles = documentFiles
  }

  const ctaBannerImageFiles = collectCTABannerImageFiles(form.sections)
  if (ctaBannerImageFiles.length > 0) {
    request.ctaBannerImageFiles = ctaBannerImageFiles
  }

  return request
}

function buildSavePageDetailPayload(form: PageFormState): SavePageDetailPayload {
  return {
    template_key: form.templateKey.trim() || 'default',
    settings: {},
    sections: form.sections.map((section, index) => buildSaveSectionPayload(section, index)),
  }
}

function buildSaveSectionPayload(
  section: PageSectionState,
  sortOrder: number,
): SavePageSectionPayload {
  return {
    ...(section.id ? { id: section.id } : {}),
    section_name: section.sectionName.trim() || defaultSectionName(section.sectionType),
    section_type: section.sectionType,
    sort_order: sortOrder,
    is_enabled: section.isEnabled,
    settings: section.settings ?? {},
    ...(section.sectionType === 'header'
      ? {
          header: {
            main_header_text: section.header.mainHeaderText.trim(),
            sub_header_text: section.header.subHeaderText.trim(),
            description: section.header.description.trim(),
            hierarchy: section.header.hierarchy,
            text_align: section.header.textAlign,
            underline_enabled:
              section.header.hierarchy === 'h2_section'
                ? (section.header.underlineEnabled ?? false)
                : false,
          } satisfies SavePageHeaderSectionPayload,
        }
      : {}),
    ...(section.sectionType === 'typography'
      ? {
          typography: {
            html_content: section.typography.htmlContent,
            text_content: buildPlainTextFromHtml(section.typography.htmlContent),
            text_align: section.typography.textAlign,
          } satisfies SavePageTypographySectionPayload,
        }
      : {}),
    ...(section.sectionType === 'gallery'
      ? {
          gallery: {
            gallery_id: section.gallery.galleryId
              ? Number.parseInt(section.gallery.galleryId, 10)
              : null,
            view_mode: section.gallery.viewMode,
            show_title_description: section.gallery.showTitleDescription ?? true,
            auto_scroll_enabled:
              section.gallery.viewMode === 'carousel'
                ? (section.gallery.autoScrollEnabled ?? false)
                : false,
          } satisfies SavePageGallerySectionPayload,
        }
      : {}),
    ...(section.sectionType === 'video'
      ? {
          video: {
            video_package_id: section.video.videoPackageId
              ? Number.parseInt(section.video.videoPackageId, 10)
              : null,
          } satisfies SavePageVideoSectionPayload,
        }
      : {}),
    ...(section.sectionType === 'quote'
      ? {
          quote: {
            quote_content: section.quote.quoteContent.trim(),
            attribution: section.quote.attribution.trim(),
          } satisfies SavePageQuoteSectionPayload,
        }
      : {}),
    ...(section.sectionType === 'cta_banner'
      ? {
          cta_banner: {
            banner_heading: section.ctaBanner.bannerHeading.trim(),
            banner_message: section.ctaBanner.bannerMessage.trim(),
            button_text: section.ctaBanner.buttonText.trim(),
            button_url: section.ctaBanner.buttonUrl.trim(),
            open_in_new_tab: true,
            ...(buildCTABannerImagePayload(section)
              ? { image: buildCTABannerImagePayload(section) }
              : {}),
          } satisfies SavePageCTABannerSectionPayload,
        }
      : {}),
    ...(section.sectionType === 'document'
      ? {
          documents: {
            items: section.documents.items.map(buildSaveDocumentPayload),
          } satisfies SavePageDocumentsSectionPayload,
        }
      : {}),
  }
}

function buildSaveDocumentPayload(document: PageSectionDocumentState): SavePageDocumentPayload {
  const existingStorageUri = document.existingStorageUri.trim()

  return {
    ...(document.id ? { id: document.id } : {}),
    display_name:
      document.displayName.trim() || suggestDocumentDisplayName(document.originalFileName),
    description: document.description.trim(),
    original_file_name:
      document.file?.name || document.originalFileName || document.displayName.trim(),
    ...(document.file
      ? {
          file_name: document.file.name,
          mime_type: document.file.type || 'application/octet-stream',
        }
      : {}),
    ...(!document.file && existingStorageUri
      ? {
          file_url: existingStorageUri,
          storage_uri: existingStorageUri,
          gcp_object_key: document.existingObjectKey || undefined,
          mime_type: document.existingMimeType || undefined,
        }
      : {}),
  }
}

function collectDocumentFiles(sections: PageSectionState[]): PageDocumentMultipartFile[] {
  const files: PageDocumentMultipartFile[] = []

  sections.forEach((section, sectionIndex) => {
    if (section.sectionType !== 'document') {
      return
    }

    section.documents.items.forEach((document, documentIndex) => {
      if (!document.file) {
        return
      }
      files.push({
        sectionIndex,
        documentIndex,
        file: document.file,
      })
    })
  })

  return files
}

function collectCTABannerImageFiles(
  sections: PageSectionState[],
): PageCTABannerImageMultipartFile[] {
  const files: PageCTABannerImageMultipartFile[] = []

  sections.forEach((section, sectionIndex) => {
    if (section.sectionType !== 'cta_banner' || !section.ctaBanner.imageFile) {
      return
    }

    files.push({
      sectionIndex,
      file: section.ctaBanner.imageFile,
    })
  })

  return files
}

function buildCTABannerImagePayload(section: PageSectionState): PageUploadInput | undefined {
  if (section.sectionType !== 'cta_banner') {
    return undefined
  }

  if (section.ctaBanner.imageFile) {
    return {
      file_name: section.ctaBanner.imageFile.name,
      mime_type: section.ctaBanner.imageFile.type || 'application/octet-stream',
    }
  }

  const existingStorageUri = section.ctaBanner.existingImageStorageUrl.trim()
  const existingObjectKey = section.ctaBanner.existingImageObjectKey.trim()

  if (!existingStorageUri && !existingObjectKey) {
    return undefined
  }

  return {
    ...(existingStorageUri
      ? {
          file_url: existingStorageUri,
          storage_uri: existingStorageUri,
        }
      : {}),
    ...(existingObjectKey ? { gcp_object_key: existingObjectKey } : {}),
  }
}

function mapSectionResponseToState(section: PageSectionResponse): PageSectionState {
  const next = createDefaultSectionState(section.section_type)

  next.id = section.id
  next.sectionName = section.section_name || defaultSectionName(section.section_type)
  next.sectionType = section.section_type
  next.isEnabled = section.is_enabled
  next.settings = section.settings ?? {}
  next.header = mapHeaderResponse(section.header)
  next.typography = mapTypographyResponse(section.typography)
  next.gallery = mapGalleryResponse(section.gallery)
  next.video = mapVideoResponse(section.video)
  next.quote = mapQuoteResponse(section.quote)
  next.ctaBanner = mapCTABannerResponse(section.cta_banner)
  next.documents.items = (section.documents?.items ?? []).map(mapDocumentResponseToState)

  return next
}

function mapDocumentResponseToState(document: PageDocumentResponse): PageSectionDocumentState {
  return {
    clientId: createPageEditorId('page-document'),
    id: document.id,
    displayName: document.display_name,
    description: document.description,
    originalFileName: document.original_file_name || document.file_name,
    existingFetchUrl: resolvePageAssetUrl(document.fetch_url || document.file_url),
    existingStorageUri: document.storage_uri,
    existingMimeType: document.mime_type,
    existingFileSize: document.file_size,
    existingObjectKey: document.gcp_object_key,
    file: null,
  }
}

function mapHeaderResponse(header?: PageHeaderSectionResponse | null) {
  return {
    mainHeaderText: header?.main_header_text ?? '',
    subHeaderText: header?.sub_header_text ?? '',
    description: header?.description ?? '',
    hierarchy: header?.hierarchy ?? 'h1_hero',
    textAlign: header?.text_align ?? 'left',
    underlineEnabled: header?.underline_enabled ?? false,
  }
}

function mapTypographyResponse(typography?: PageTypographySectionResponse | null) {
  return {
    htmlContent: typography?.html_content ?? '',
    textContent: typography?.text_content ?? '',
    textAlign: typography?.text_align ?? 'left',
  }
}

function mapGalleryResponse(gallery?: PageGallerySectionResponse | null) {
  return {
    galleryId:
      typeof gallery?.gallery_id === 'number' ? String(gallery.gallery_id) : '',
    viewMode: gallery?.view_mode ?? 'grid',
    showTitleDescription: gallery?.show_title_description ?? true,
    autoScrollEnabled: gallery?.auto_scroll_enabled ?? false,
  }
}

function mapVideoResponse(video?: PageVideoSectionResponse | null) {
  return {
    videoPackageId:
      typeof video?.video_package_id === 'number' ? String(video.video_package_id) : '',
  }
}

function mapQuoteResponse(quote?: PageQuoteSectionResponse | null) {
  return {
    quoteContent: quote?.quote_content ?? '',
    attribution: quote?.attribution ?? '',
  }
}

function mapCTABannerResponse(ctaBanner?: PageCTABannerSectionResponse | null) {
  return {
    bannerHeading: ctaBanner?.banner_heading ?? '',
    bannerMessage: ctaBanner?.banner_message ?? '',
    buttonText: ctaBanner?.button_text ?? '',
    buttonUrl: ctaBanner?.button_url ?? '',
    openInNewTab: true,
    imageFile: null,
    existingImageFetchUrl: resolvePageAssetUrl(ctaBanner?.image?.fetch_url ?? ''),
    existingImageStorageUrl: ctaBanner?.image?.storage_uri ?? '',
    existingImageObjectKey: ctaBanner?.image?.gcp_object_key ?? '',
  }
}

function normalizePageDetail(pageDetail?: PageContentDetailResponse | null) {
  return {
    template_key: pageDetail?.template_key ?? 'default',
    sections: pageDetail?.sections ?? [],
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

function defaultSectionName(sectionType: PageSectionType) {
  switch (sectionType) {
    case 'header':
      return 'Header Module'
    case 'typography':
      return 'Typography'
    case 'gallery':
      return 'Gallery Module'
    case 'video':
      return 'Video Module'
    case 'document':
      return 'Document Module'
    case 'quote':
      return 'Quote Module'
    case 'cta_banner':
      return 'CTA Banner'
    default:
      return 'Content Section'
  }
}

function suggestDocumentDisplayName(fileName: string) {
  const trimmed = fileName.trim()
  if (!trimmed) {
    return 'Document'
  }
  return trimmed.replace(/\.[^.]+$/, '')
}

function buildPlainTextFromHtml(html: string) {
  return html
    .replace(/<[^>]*>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

export function isTypographyContentEmpty(html: string) {
  return !buildPlainTextFromHtml(html)
}
