import { useEffect, useMemo, useRef, useState, type DragEvent, type ReactNode } from 'react'
import DragIndicatorOutlinedIcon from '@mui/icons-material/DragIndicatorOutlined'
import ExpandLessOutlinedIcon from '@mui/icons-material/ExpandLessOutlined'
import ExpandMoreOutlinedIcon from '@mui/icons-material/ExpandMoreOutlined'
import KeyboardArrowDownOutlinedIcon from '@mui/icons-material/KeyboardArrowDownOutlined'
import KeyboardArrowUpOutlinedIcon from '@mui/icons-material/KeyboardArrowUpOutlined'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import { getApiErrorMessage } from '../api/apiError'
import { mediaApi } from '../api/mediaApi'
import { videoApi } from '../api/videoApi'
import {
  isModulePage,
  pagesApi,
  resolvePageParentId,
  resolvePageParentSlug,
  type PageParentOption,
  type PageSectionType,
} from '../api/pagesApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { RichTextEditor } from '../components/cms/RichTextEditor'
import {
  AddIcon,
  CloudUploadIcon,
  DeleteIcon,
  DownloadIcon,
  EditIcon,
  GalleryIcon,
  PagesIcon,
  VideoIcon,
} from '../components/icons'
import { Loader } from '../components/Loader'
import { UploadDropzone } from '../components/media/UploadDropzone'
import {
  buildFullPageUrlSlug,
  buildPageFormStateFromDetail,
  buildSavePageRequest,
  countHeroHeaderSections,
  createDefaultDocumentState,
  createDefaultPageFormState,
  createDefaultSectionState,
  getPageDocumentValidationMessages,
  getDisallowedParentPageIds,
  getPageSectionValidationMessages,
  isTypographyContentEmpty,
  normalizePageSlugInput,
  reorderPageSections,
  validatePageDocumentFields,
  validatePageForm,
  validatePageSectionFields,
  type PageDocumentFieldErrors,
  type PageDocumentValidationErrors,
  type PageFormErrors,
  type PageSectionFieldErrors,
  type PageSectionDropPlacement,
  type PageFormState,
  type PageSectionDocumentState,
  type PageSectionState,
  type PageSectionValidationErrors,
} from '../lib/pagesForm'
import {
  RESOURCE_FILE_ACCEPT,
  getResourceUploadValidationErrorMessage,
  validateResourceUploadFile,
} from '../lib/resourceUpload'
import {
  clearCurrentPage,
  clearPageSaveState,
  createPage,
  fetchPageById,
  selectCurrentPage,
  selectPageDetail,
  selectPageSave,
  updatePage,
} from '../store/pagesSlice'
import { useAppDispatch, useAppSelector } from '../store/hooks'
import styles from '../styles/PageEditorPage.module.css'

type SubmitMode = 'draft' | 'publish'

type PageEditorPageProps = {
  mode?: 'edit' | 'view'
}

type PageOptionsStatus = 'idle' | 'loading' | 'succeeded' | 'failed'
type GalleryOptionsStatus = 'idle' | 'loading' | 'succeeded' | 'failed'
type VideoOptionsStatus = 'idle' | 'loading' | 'succeeded' | 'failed'

type PageGalleryOption = {
  id: number
  name: string
}

type PageVideoOption = {
  id: number
  title: string
}

type ModuleTypeOption = {
  type: PageSectionType
  label: string
  description: string
  icon: ReactNode
}

const DOCUMENT_UPLOAD_ACCEPT = RESOURCE_FILE_ACCEPT
const SECTION_DRAG_AUTO_SCROLL_EDGE_PX = 96
const SECTION_DRAG_AUTO_SCROLL_STEP_PX = 32

export function getSectionDragAutoScrollDelta(clientY: number, viewportHeight: number) {
  if (clientY <= SECTION_DRAG_AUTO_SCROLL_EDGE_PX) {
    return -SECTION_DRAG_AUTO_SCROLL_STEP_PX
  }

  if (viewportHeight - clientY <= SECTION_DRAG_AUTO_SCROLL_EDGE_PX) {
    return SECTION_DRAG_AUTO_SCROLL_STEP_PX
  }

  return 0
}

const PAGE_HERO_IMAGE_MAX_FILE_SIZE_MB = 5
const PAGE_HERO_IMAGE_MAX_FILE_SIZE_BYTES =
  PAGE_HERO_IMAGE_MAX_FILE_SIZE_MB * 1024 * 1024
const HERO_IMAGE_SUPPORTED_MIME_TYPES = new Set([
  'image/png',
  'image/jpeg',
  'image/jpg',
  'image/webp',
])
const HERO_IMAGE_SUPPORTED_EXTENSIONS = new Set([
  '.png',
  '.jpg',
  '.jpeg',
  '.webp',
])
const IMAGE_UPLOAD_ACCEPT = [
  ...HERO_IMAGE_SUPPORTED_MIME_TYPES,
  ...HERO_IMAGE_SUPPORTED_EXTENSIONS,
].join(',')

export function PageEditorPage({ mode = 'edit' }: PageEditorPageProps) {
  const dispatch = useAppDispatch()
  const navigate = useNavigate()
  const params = useParams()
  const { t } = useTranslation()
  const currentPage = useAppSelector(selectCurrentPage)
  const detailState = useAppSelector(selectPageDetail)
  const saveState = useAppSelector(selectPageSave)

  const parsedPageId = params.id ? Number.parseInt(params.id, 10) : null
  const isEditMode = parsedPageId !== null && Number.isFinite(parsedPageId)
  const isViewMode = mode === 'view' && isEditMode
  const isInvalidEditId = params.id !== undefined && !isEditMode
  const isModuleManaged = currentPage ? isModulePage(currentPage) : false
  const isReadOnlyMode = isViewMode || isModuleManaged

  const [form, setForm] = useState<PageFormState>(createDefaultPageFormState())
  const [errors, setErrors] = useState<PageFormErrors>({})
  const [sectionFieldErrors, setSectionFieldErrors] = useState<PageSectionValidationErrors>({})
  const [documentFieldErrors, setDocumentFieldErrors] = useState<PageDocumentValidationErrors>({})
  const [slugTouched, setSlugTouched] = useState(false)
  const [heroImagePreviewUrl, setHeroImagePreviewUrl] = useState<string | null>(null)
  const [existingHeroObjectUrl, setExistingHeroObjectUrl] = useState<string | null>(null)
  const [parentPageOptions, setParentPageOptions] = useState<PageParentOption[]>([])
  const [parentPageOptionsStatus, setParentPageOptionsStatus] =
    useState<PageOptionsStatus>('idle')
  const [parentPageOptionsError, setParentPageOptionsError] = useState<string | null>(null)
  const [galleryOptions, setGalleryOptions] = useState<PageGalleryOption[]>([])
  const [galleryOptionsStatus, setGalleryOptionsStatus] =
    useState<GalleryOptionsStatus>('idle')
  const [galleryOptionsError, setGalleryOptionsError] = useState<string | null>(null)
  const [videoOptions, setVideoOptions] = useState<PageVideoOption[]>([])
  const [videoOptionsStatus, setVideoOptionsStatus] =
    useState<VideoOptionsStatus>('idle')
  const [videoOptionsError, setVideoOptionsError] = useState<string | null>(null)
  const [initializedPageId, setInitializedPageId] = useState<number | null>(null)
  const [modulePickerOpen, setModulePickerOpen] = useState(false)
  const [draggedSectionClientId, setDraggedSectionClientId] = useState<string | null>(null)
  const [dragOverSectionClientId, setDragOverSectionClientId] = useState<string | null>(null)
  const [dragOverSectionPlacement, setDragOverSectionPlacement] =
    useState<PageSectionDropPlacement | null>(null)
  const draggedSectionClientIdRef = useRef<string | null>(null)
  const dragOverSectionClientIdRef = useRef<string | null>(null)
  const dragOverSectionPlacementRef = useRef<PageSectionDropPlacement | null>(null)
  const [documentPreviewUrls, setDocumentPreviewUrls] = useState<Record<string, string>>({})
  const [activeDocumentAction, setActiveDocumentAction] = useState<
    Record<string, 'preview' | 'download' | undefined>
  >({})
  const documentPreviewUrlsRef = useRef<Record<string, string>>({})
  const [ctaImagePreviewUrls, setCtaImagePreviewUrls] = useState<Record<string, string>>({})
  const ctaImagePreviewUrlsRef = useRef<Record<string, string>>({})

  useEffect(() => {
    let cancelled = false

    async function loadParentPageOptions() {
      setParentPageOptionsStatus('loading')
      setParentPageOptionsError(null)

      try {
        const options = await pagesApi.listPageParentOptions()
        if (cancelled) {
          return
        }

        setParentPageOptions(options)
        setParentPageOptionsStatus('succeeded')
      } catch (error) {
        if (cancelled) {
          return
        }

        setParentPageOptions([])
        setParentPageOptionsStatus('failed')
        setParentPageOptionsError(getApiErrorMessage(error))
      }
    }

    async function loadGalleryOptions() {
      setGalleryOptionsStatus('loading')
      setGalleryOptionsError(null)

      try {
        const options = await mediaApi.listGalleries()
        if (cancelled) {
          return
        }

        setGalleryOptions(
          options.map((gallery) => ({
            id: gallery.id,
            name: gallery.name,
          })),
        )
        setGalleryOptionsStatus('succeeded')
      } catch (error) {
        if (cancelled) {
          return
        }

        setGalleryOptions([])
        setGalleryOptionsStatus('failed')
        setGalleryOptionsError(getApiErrorMessage(error))
      }
    }

    async function loadVideoOptions() {
      setVideoOptionsStatus('loading')
      setVideoOptionsError(null)

      try {
        const options = await videoApi.listVideoPackages()
        if (cancelled) {
          return
        }

        setVideoOptions(
          options.map((item) => ({
            id: item.id,
            title: item.title,
          })),
        )
        setVideoOptionsStatus('succeeded')
      } catch (error) {
        if (cancelled) {
          return
        }

        setVideoOptions([])
        setVideoOptionsStatus('failed')
        setVideoOptionsError(getApiErrorMessage(error))
      }
    }

    void loadParentPageOptions()
    void loadGalleryOptions()
    void loadVideoOptions()

    return () => {
      cancelled = true
    }
  }, [])

  useEffect(() => {
    if (isEditMode && parsedPageId) {

      setInitializedPageId(null)
      void dispatch(fetchPageById(parsedPageId))
      return
    }

    dispatch(clearCurrentPage())
    setForm(createDefaultPageFormState())
    setErrors({})
    setSectionFieldErrors({})
    setDocumentFieldErrors({})
    setSlugTouched(false)
    setInitializedPageId(null)
  }, [dispatch, isEditMode, parsedPageId])

  const parentPageOptionsById = useMemo(
    () => new Map(parentPageOptions.map((page) => [page.id, page])),
    [parentPageOptions],
  )
  const currentPageParentId = currentPage ? resolvePageParentId(currentPage) : null
  const currentPageParentSlug =
    (currentPageParentId ? parentPageOptionsById.get(currentPageParentId)?.url_slug : '') ||
    (currentPage ? resolvePageParentSlug(currentPage) : '')
  const isWaitingForParentContext = Boolean(
    isEditMode &&
      currentPage &&
      currentPageParentId &&
      !currentPageParentSlug &&
      parentPageOptionsStatus === 'loading',
  )

  useEffect(() => {
    if (!isEditMode || !currentPage || initializedPageId === currentPage.id) {
      return
    }

    if (isWaitingForParentContext) {
      return
    }

    // eslint-disable-next-line react-hooks/set-state-in-effect
    setForm(
      buildPageFormStateFromDetail(currentPage, {
        parentPageSlug: currentPageParentSlug,
      }),
    )
    setErrors({})
    setSectionFieldErrors({})
    setDocumentFieldErrors({})
    setSlugTouched(true)
    setInitializedPageId(currentPage.id)
  }, [
    currentPage,
    currentPageParentId,
    currentPageParentSlug,
    initializedPageId,
    isEditMode,
    isWaitingForParentContext,
  ])

  const selectedParentPageId = form.parentPageId ? Number.parseInt(form.parentPageId, 10) : null
  const selectedParentPageSlug =
    (selectedParentPageId ? parentPageOptionsById.get(selectedParentPageId)?.url_slug : '') ||
    (selectedParentPageId === currentPageParentId ? currentPageParentSlug : '')
  const effectiveUrlSlug = buildFullPageUrlSlug(form.urlSlug, selectedParentPageSlug)
  const slugPrefix = selectedParentPageSlug
    ? `${selectedParentPageSlug.replace(/\/+$/, '')}/`
    : '/'
  const disallowedParentPageIds = useMemo(
    () =>
      isEditMode && parsedPageId
        ? getDisallowedParentPageIds(parsedPageId, parentPageOptions)
        : new Set<number>(),
    [isEditMode, parsedPageId, parentPageOptions],
  )
  const availableParentPageOptions = useMemo(
    () => parentPageOptions.filter((page) => page.id !== parsedPageId),
    [parentPageOptions, parsedPageId],
  )

  const moduleTypeOptions = useMemo<ModuleTypeOption[]>(
    () => [
      {
        type: 'typography',
        label: t('pages.modules.types.typography'),
        description: t('pages.modules.descriptions.typography'),
        icon: <TypographyModuleIcon />,
      },
      {
        type: 'gallery',
        label: t('pages.modules.types.gallery'),
        description: t('pages.modules.descriptions.gallery'),
        icon: <GalleryModuleIcon />,
      },
      {
        type: 'video',
        label: t('pages.modules.types.video'),
        description: t('pages.modules.descriptions.video'),
        icon: <VideoModuleIcon />,
      },
      {
        type: 'document',
        label: t('pages.modules.types.document'),
        description: t('pages.modules.descriptions.document'),
        icon: <DocumentModuleIcon />,
      },
      {
        type: 'quote',
        label: t('pages.modules.types.quote'),
        description: t('pages.modules.descriptions.quote'),
        icon: <QuoteModuleIcon />,
      },
      {
        type: 'cta_banner',
        label: t('pages.modules.types.cta_banner'),
        description: t('pages.modules.descriptions.cta_banner'),
        icon: <CtaModuleIcon />,
      },
      {
        type: 'header',
        label: t('pages.modules.types.header'),
        description: t('pages.modules.descriptions.header'),
        icon: <HeaderModuleIcon />,
      },
    ],
    [t],
  )

  useEffect(() => {
    return () => {
      dispatch(clearPageSaveState())
    }
  }, [dispatch])

  useEffect(() => {
    if (!form.heroImageFile) {

      setHeroImagePreviewUrl(null)
      return
    }

    const objectUrl = URL.createObjectURL(form.heroImageFile)
    setHeroImagePreviewUrl(objectUrl)

    return () => {
      URL.revokeObjectURL(objectUrl)
    }
  }, [form.heroImageFile])

  useEffect(() => {
    if (
      !form.heroImageEnabled ||
      form.removeHeroImage ||
      form.heroImageFile ||
      !form.existingHeroImageFetchUrl
    ) {

      setExistingHeroObjectUrl((current) => {
        if (current) {
          URL.revokeObjectURL(current)
        }
        return null
      })
      return
    }

    let cancelled = false
    let objectUrlToRevoke: string | null = null

    async function loadHeroImage() {
      try {
        const blob = await pagesApi.fetchPageHeroImageContent(form.existingHeroImageFetchUrl)
        const objectUrl = URL.createObjectURL(blob)
        objectUrlToRevoke = objectUrl

        if (cancelled) {
          URL.revokeObjectURL(objectUrl)
          return
        }

        setExistingHeroObjectUrl((current) => {
          if (current) {
            URL.revokeObjectURL(current)
          }
          return objectUrl
        })
      } catch {
        return
      }
    }

    void loadHeroImage()

    return () => {
      cancelled = true
      if (objectUrlToRevoke) {
        URL.revokeObjectURL(objectUrlToRevoke)
      }
    }
  }, [
    form.existingHeroImageFetchUrl,
    form.heroImageEnabled,
    form.heroImageFile,
    form.removeHeroImage,
  ])

  const isBusy = saveState.status === 'loading'
  const isInitialLoad =
    (isEditMode && detailState.status === 'loading' && !currentPage) || isWaitingForParentContext
  const errorMessages = useMemo(
    () =>
      Array.from(
        new Set([
          ...Object.values(errors).filter((message): message is string => Boolean(message)),
          ...getPageSectionValidationMessages(sectionFieldErrors),
          ...getPageDocumentValidationMessages(documentFieldErrors),
        ]),
      ),
    [documentFieldErrors, errors, sectionFieldErrors],
  )
  const currentHeroPreview =
    heroImagePreviewUrl ??
    (form.heroImageEnabled && !form.removeHeroImage ? existingHeroObjectUrl : null)
  const documentPreviewSources = useMemo(
    () =>
      form.sections
        .flatMap((section) => section.documents.items)
        .filter((document) =>
          canPreviewDocument(
            document.file?.type || document.existingMimeType,
            document.originalFileName || document.file?.name || '',
          ),
        )
        .map((document) => ({
          clientId: document.clientId,
          file: document.file,
          fetchUrl: document.existingFetchUrl,
          sourceKey: document.file
            ? [
                document.file.name,
                document.file.size,
                document.file.lastModified,
                document.file.type,
              ].join(':')
            : document.existingFetchUrl,
        }))
        .sort((left, right) => left.clientId.localeCompare(right.clientId)),
    [form.sections],
  )
  const documentPreviewSignature = useMemo(
    () =>
      JSON.stringify(
        documentPreviewSources.map((source) => ({
          clientId: source.clientId,
          sourceKey: source.sourceKey,
        })),
      ),
    [documentPreviewSources],
  )
  const ctaImagePreviewSources = useMemo(
    () =>
      form.sections
        .filter(
          (section) =>
            section.sectionType === 'cta_banner' &&
            (Boolean(section.ctaBanner.imageFile) ||
              Boolean(section.ctaBanner.existingImageFetchUrl)),
        )
        .map((section) => ({
          clientId: section.clientId,
          file: section.ctaBanner.imageFile,
          fetchUrl: section.ctaBanner.existingImageFetchUrl,
          sourceKey: section.ctaBanner.imageFile
            ? [
                section.ctaBanner.imageFile.name,
                section.ctaBanner.imageFile.size,
                section.ctaBanner.imageFile.lastModified,
                section.ctaBanner.imageFile.type,
              ].join(':')
            : section.ctaBanner.existingImageFetchUrl,
        }))
        .sort((left, right) => left.clientId.localeCompare(right.clientId)),
    [form.sections],
  )
  const ctaImagePreviewSignature = useMemo(
    () =>
      JSON.stringify(
        ctaImagePreviewSources.map((source) => ({
          clientId: source.clientId,
          sourceKey: source.sourceKey,
        })),
      ),
    [ctaImagePreviewSources],
  )

  useEffect(() => {
    let cancelled = false

    async function loadDocumentPreviews() {
      const nextEntries = await Promise.all(
        documentPreviewSources.map(async (source) => {
          try {
            if (source.file) {
              return [source.clientId, URL.createObjectURL(source.file)] as const
            }

            if (!source.fetchUrl) {
              return [source.clientId, ''] as const
            }

            const blob = await pagesApi.fetchPageDocumentContent(source.fetchUrl)
            return [source.clientId, URL.createObjectURL(blob)] as const
          } catch {
            return [source.clientId, ''] as const
          }
        }),
      )

      if (cancelled) {
        nextEntries.forEach(([, url]) => {
          if (url) {
            URL.revokeObjectURL(url)
          }
        })
        return
      }

      setDocumentPreviewUrls((current) => {
        Object.values(current).forEach((url) => {
          URL.revokeObjectURL(url)
        })

        const next = Object.fromEntries(nextEntries.filter(([, url]) => Boolean(url)))
        documentPreviewUrlsRef.current = next
        return next
      })
    }

    if (!documentPreviewSources.length) {
      setDocumentPreviewUrls((current) => {
        Object.values(current).forEach((url) => {
          URL.revokeObjectURL(url)
        })
        documentPreviewUrlsRef.current = {}
        return {}
      })
      return
    }

    void loadDocumentPreviews()

    return () => {
      cancelled = true
    }
  }, [documentPreviewSignature])

  useEffect(() => {
    function handleWindowDragOver(event: globalThis.DragEvent) {
      if (!draggedSectionClientIdRef.current || isReadOnlyMode || isBusy) {
        return
      }

      const delta = getSectionDragAutoScrollDelta(event.clientY, globalThis.window.innerHeight)
      if (delta !== 0) {
        globalThis.window.scrollBy({
          top: delta,
          behavior: 'auto',
        })
      }
    }

    globalThis.window.addEventListener('dragover', handleWindowDragOver)

    return () => {
      globalThis.window.removeEventListener('dragover', handleWindowDragOver)
    }
  }, [isBusy, isReadOnlyMode])

  useEffect(() => {
    return () => {
      Object.values(documentPreviewUrlsRef.current).forEach((url) => {
        URL.revokeObjectURL(url)
      })
      documentPreviewUrlsRef.current = {}
    }
  }, [])

  useEffect(() => {
    let cancelled = false

    async function loadCTAImagePreviews() {
      const nextEntries = await Promise.all(
        ctaImagePreviewSources.map(async (source) => {
          try {
            if (source.file) {
              return [source.clientId, URL.createObjectURL(source.file)] as const
            }

            if (!source.fetchUrl) {
              return [source.clientId, ''] as const
            }

            const blob = await pagesApi.fetchPageCTABannerImageContent(source.fetchUrl)
            return [source.clientId, URL.createObjectURL(blob)] as const
          } catch {
            return [source.clientId, ''] as const
          }
        }),
      )

      if (cancelled) {
        nextEntries.forEach(([, url]) => {
          if (url) {
            URL.revokeObjectURL(url)
          }
        })
        return
      }

      setCtaImagePreviewUrls((current) => {
        Object.values(current).forEach((url) => {
          URL.revokeObjectURL(url)
        })

        const next = Object.fromEntries(nextEntries.filter(([, url]) => Boolean(url)))
        ctaImagePreviewUrlsRef.current = next
        return next
      })
    }

    if (!ctaImagePreviewSources.length) {
      setCtaImagePreviewUrls((current) => {
        Object.values(current).forEach((url) => {
          URL.revokeObjectURL(url)
        })
        ctaImagePreviewUrlsRef.current = {}
        return {}
      })
      return
    }

    void loadCTAImagePreviews()

    return () => {
      cancelled = true
    }
  }, [ctaImagePreviewSignature])

  useEffect(() => {
    return () => {
      Object.values(ctaImagePreviewUrlsRef.current).forEach((url) => {
        URL.revokeObjectURL(url)
      })
      ctaImagePreviewUrlsRef.current = {}
    }
  }, [])

  function clearErrors(...keys: Array<string>) {
    setErrors((current) => {
      const next = { ...current }
      for (const key of keys) {
        delete next[key as keyof PageFormErrors]
      }
      return next
    })
  }

  function setFieldError<K extends keyof PageFormErrors>(
    key: K,
    message?: PageFormErrors[K],
  ) {
    setErrors((current) => {
      if (!message) {
        if (!(key in current)) {
          return current
        }

        const next = { ...current }
        delete next[key]
        return next
      }

      return {
        ...current,
        [key]: message,
      }
    })
  }

  function setSectionFieldError(
    sectionClientId: string,
    key: keyof PageSectionFieldErrors,
    message?: string,
  ) {
    setSectionFieldErrors((current) => {
      const currentSectionErrors = current[sectionClientId] ?? {}

      if (!message) {
        if (!(key in currentSectionErrors)) {
          return current
        }

        const nextSectionErrors = { ...currentSectionErrors }
        delete nextSectionErrors[key]

        if (Object.keys(nextSectionErrors).length === 0) {
          const next = { ...current }
          delete next[sectionClientId]
          return next
        }

        return {
          ...current,
          [sectionClientId]: nextSectionErrors,
        }
      }

      return {
        ...current,
        [sectionClientId]: {
          ...currentSectionErrors,
          [key]: message,
        },
      }
    })
  }

  function getSectionFieldError(
    sectionClientId: string,
    key: keyof PageSectionFieldErrors,
  ) {
    return sectionFieldErrors[sectionClientId]?.[key]
  }

  function setDocumentFieldError(
    documentClientId: string,
    key: keyof PageDocumentFieldErrors,
    message?: string,
  ) {
    setDocumentFieldErrors((current) => {
      const currentDocumentErrors = current[documentClientId] ?? {}

      if (!message) {
        if (!(key in currentDocumentErrors)) {
          return current
        }

        const nextDocumentErrors = { ...currentDocumentErrors }
        delete nextDocumentErrors[key]

        if (Object.keys(nextDocumentErrors).length === 0) {
          const next = { ...current }
          delete next[documentClientId]
          return next
        }

        return {
          ...current,
          [documentClientId]: nextDocumentErrors,
        }
      }

      return {
        ...current,
        [documentClientId]: {
          ...currentDocumentErrors,
          [key]: message,
        },
      }
    })
  }

  function getDocumentFieldError(
    documentClientId: string,
    key: keyof PageDocumentFieldErrors,
  ) {
    return documentFieldErrors[documentClientId]?.[key]
  }

  function updateField<K extends keyof PageFormState>(key: K, value: PageFormState[K]) {
    setForm((current) => ({
      ...current,
      [key]: value,
    }))
    clearErrors(key)
  }

  function updateSection(
    sectionClientId: string,
    updater: (section: PageSectionState) => PageSectionState,
  ) {
    setForm((current) => ({
      ...current,
      sections: current.sections.map((section) =>
        section.clientId === sectionClientId ? updater(section) : section,
      ),
    }))
    clearErrors('sections')
  }

  function updateDocument(
    sectionClientId: string,
    documentClientId: string,
    updater: (document: PageSectionDocumentState) => PageSectionDocumentState,
  ) {
    updateSection(sectionClientId, (section) => ({
      ...section,
      documents: {
        items: section.documents.items.map((document) =>
          document.clientId === documentClientId ? updater(document) : document,
        ),
      },
    }))
  }

  function handlePageTitleChange(value: string) {
    const nextUrlSlug = slugTouched ? form.urlSlug : normalizePageSlugInput(value)

    setForm((current) => ({
      ...current,
      pageTitle: value,
      urlSlug: slugTouched ? current.urlSlug : nextUrlSlug,
    }))
    setFieldError(
      'pageTitle',
      value.trim() ? undefined : t('pages.validation.pageTitleRequired'),
    )

    if (!slugTouched) {
      setFieldError(
        'urlSlug',
        nextUrlSlug ? undefined : t('pages.validation.urlSlugRequired'),
      )
    }
  }

  function handleSlugChange(value: string) {
    setSlugTouched(true)
    const normalizedSlug = normalizePageSlugInput(value)
    setForm((current) => ({
      ...current,
      urlSlug: normalizedSlug,
    }))
    setFieldError(
      'urlSlug',
      normalizedSlug ? undefined : t('pages.validation.urlSlugRequired'),
    )
  }

  function handleParentPageChange(value: string) {
    updateField('parentPageId', value)
  }

  function handleStatusToggle(nextPublished: boolean) {
    updateField('status', nextPublished ? 'published' : 'draft')
  }

  function handleHeroToggle(enabled: boolean) {
    setForm((current) => ({
      ...current,
      heroImageEnabled: enabled,
      heroImageFile: enabled ? current.heroImageFile : null,
      removeHeroImage: enabled ? false : true,
    }))
    clearErrors('heroImageFile')
  }

  function handleHeroFiles(files: File[]) {
    const file = files[0]
    if (!file) {
      return
    }

    const validationMessage = getHeroImageValidationMessage(file, t)
    if (validationMessage) {
      setErrors((current) => ({
        ...current,
        heroImageFile: validationMessage,
      }))
      toast.error(validationMessage)
      return
    }

    setForm((current) => ({
      ...current,
      heroImageEnabled: true,
      heroImageFile: file,
      removeHeroImage: false,
    }))
    clearErrors('heroImageFile')
  }

  function removeHeroImage() {
    setForm((current) => ({
      ...current,
      heroImageFile: null,
      removeHeroImage: true,
    }))
    clearErrors('heroImageFile')
  }

  function handleCTABannerImageFiles(sectionClientId: string, files: File[]) {
    const file = files[0]
    if (!file) {
      return
    }

    updateSection(sectionClientId, (current) => ({
      ...current,
      ctaBanner: {
        ...current.ctaBanner,
        imageFile: file,
      },
    }))
  }

  function removeCTABannerImage(sectionClientId: string) {
    updateSection(sectionClientId, (current) => ({
      ...current,
      ctaBanner: {
        ...current.ctaBanner,
        imageFile: null,
        existingImageFetchUrl: '',
        existingImageStorageUrl: '',
        existingImageObjectKey: '',
      },
    }))
  }

  function handleAddSection(sectionType: PageSectionType) {
    setForm((current) => {
      const nextSection = createDefaultSectionState(sectionType)
      if (sectionType === 'header' && countHeroHeaderSections(current.sections) > 0) {
        nextSection.header.hierarchy = 'h2_section'
      }
      return {
        ...current,
        sections: [...current.sections, nextSection],
      }
    })
    setModulePickerOpen(false)
    clearErrors('sections')
  }

  function handleRemoveSection(sectionClientId: string) {
    const documentClientIds =
      form.sections.find((section) => section.clientId === sectionClientId)?.documents.items.map(
        (document) => document.clientId,
      ) ?? []

    setForm((current) => ({
      ...current,
      sections: current.sections.filter((section) => section.clientId !== sectionClientId),
    }))
    setSectionFieldErrors((current) => {
      if (!(sectionClientId in current)) {
        return current
      }

      const next = { ...current }
      delete next[sectionClientId]
      return next
    })
    setDocumentFieldErrors((current) => {
      if (!documentClientIds.length) {
        return current
      }

      const next = { ...current }
      let changed = false

      documentClientIds.forEach((documentClientId) => {
        if (documentClientId in next) {
          delete next[documentClientId]
          changed = true
        }
      })

      return changed ? next : current
    })
    clearErrors('sections')
  }

  function handleMoveSection(sectionClientId: string, direction: -1 | 1) {
    setForm((current) => ({
      ...current,
      sections: moveItemByClientId(current.sections, sectionClientId, direction),
    }))
    clearErrors('sections')
  }

  function handleSectionDragStart(sectionClientId: string) {
    if (isReadOnlyMode || isBusy) {
      return
    }
    draggedSectionClientIdRef.current = sectionClientId
    dragOverSectionClientIdRef.current = sectionClientId
    dragOverSectionPlacementRef.current = 'before'
    setDraggedSectionClientId(sectionClientId)
    setDragOverSectionClientId(sectionClientId)
    setDragOverSectionPlacement('before')
  }

  function resolveSectionDropPlacement(event: DragEvent<HTMLElement>): PageSectionDropPlacement {
    const bounds = event.currentTarget.getBoundingClientRect()
    if (bounds.height <= 0) {
      return 'after'
    }

    return event.clientY - bounds.top <= bounds.height / 2 ? 'before' : 'after'
  }

  function handleSectionDragOver(event: DragEvent<HTMLElement>, sectionClientId: string) {
    const activeDraggedSectionClientId = draggedSectionClientIdRef.current
    if (
      isReadOnlyMode ||
      isBusy ||
      !activeDraggedSectionClientId ||
      activeDraggedSectionClientId === sectionClientId
    ) {
      return
    }

    event.preventDefault()
    event.dataTransfer.dropEffect = 'move'
    const placement = resolveSectionDropPlacement(event)
    dragOverSectionClientIdRef.current = sectionClientId
    dragOverSectionPlacementRef.current = placement
    setDragOverSectionClientId(sectionClientId)
    setDragOverSectionPlacement(placement)
  }

  function handleSectionDrop(event: DragEvent<HTMLElement>, sectionClientId: string) {
    const activeDraggedSectionClientId = draggedSectionClientIdRef.current
    if (
      isReadOnlyMode ||
      isBusy ||
      !activeDraggedSectionClientId ||
      activeDraggedSectionClientId === sectionClientId
    ) {
      return
    }

    event.preventDefault()
    const placement =
      dragOverSectionClientIdRef.current === sectionClientId && dragOverSectionPlacementRef.current
        ? dragOverSectionPlacementRef.current
        : resolveSectionDropPlacement(event)

    setForm((current) => ({
      ...current,
      sections: reorderPageSections(
        current.sections,
        activeDraggedSectionClientId,
        sectionClientId,
        placement,
      ),
    }))
    draggedSectionClientIdRef.current = null
    dragOverSectionClientIdRef.current = null
    dragOverSectionPlacementRef.current = null
    setDraggedSectionClientId(null)
    setDragOverSectionClientId(null)
    setDragOverSectionPlacement(null)
    clearErrors('sections')
  }

  function handleSectionDragEnd() {
    draggedSectionClientIdRef.current = null
    dragOverSectionClientIdRef.current = null
    dragOverSectionPlacementRef.current = null
    setDraggedSectionClientId(null)
    setDragOverSectionClientId(null)
    setDragOverSectionPlacement(null)
  }

  function handleDocumentFiles(sectionClientId: string, files: File[]) {
    if (!files.length) {
      return
    }

    const validFiles: File[] = []
    let validationMessage: string | null = null

    files.forEach((file) => {
      const nextValidationMessage = getDocumentUploadValidationMessage(file)
      if (nextValidationMessage) {
        validationMessage ??= nextValidationMessage
        return
      }

      validFiles.push(file)
    })

    if (!validFiles.length) {
      if (validationMessage) {
        setSectionFieldError(sectionClientId, 'documents', validationMessage)
        toast.error(validationMessage)
      }
      return
    }

    updateSection(sectionClientId, (section) => ({
      ...section,
      documents: {
        items: [...section.documents.items, ...validFiles.map(createDefaultDocumentState)],
      },
    }))
    setSectionFieldError(sectionClientId, 'documents', undefined)

    if (validationMessage) {
      toast.error(validationMessage)
    }
  }

  function handleReplaceDocumentFile(
    sectionClientId: string,
    documentClientId: string,
    file: File,
  ) {
    const validationMessage = getDocumentUploadValidationMessage(file)
    if (validationMessage) {
      setSectionFieldError(sectionClientId, 'documents', validationMessage)
      toast.error(validationMessage)
      return
    }

    updateDocument(sectionClientId, documentClientId, (document) => ({
      ...document,
      file,
      originalFileName: file.name,
      displayName: document.displayName.trim() || stripFileExtension(file.name),
      existingMimeType: file.type || document.existingMimeType,
      existingFileSize: file.size,
    }))
    setSectionFieldError(sectionClientId, 'documents', undefined)
    setDocumentFieldError(documentClientId, 'displayName', undefined)
  }

  function handleRemoveDocument(sectionClientId: string, documentClientId: string) {
    const remainingDocuments =
      (form.sections.find((section) => section.clientId === sectionClientId)?.documents.items
        .length ?? 0) - 1

    updateSection(sectionClientId, (section) => ({
      ...section,
      documents: {
        items: section.documents.items.filter((document) => document.clientId !== documentClientId),
      },
    }))
    setSectionFieldError(
      sectionClientId,
      'documents',
      remainingDocuments > 0 ? undefined : t('pages.validation.documentsRequired'),
    )
    setDocumentFieldErrors((current) => {
      if (!(documentClientId in current)) {
        return current
      }

      const next = { ...current }
      delete next[documentClientId]
      return next
    })
  }

  function handleMoveDocument(
    sectionClientId: string,
    documentClientId: string,
    direction: -1 | 1,
  ) {
    updateSection(sectionClientId, (section) => ({
      ...section,
      documents: {
        items: moveItemByClientId(section.documents.items, documentClientId, direction),
      },
    }))
  }

  async function handlePreviewDocument(document: PageSectionDocumentState) {
    setActiveDocumentAction((current) => ({
      ...current,
      [document.clientId]: 'preview',
    }))

    try {
      const previewUrl =
        documentPreviewUrls[document.clientId] ||
        (await createTemporaryDocumentObjectUrl(document))

      if (!previewUrl) {
        throw new Error('Preview URL unavailable')
      }

      window.open(previewUrl, '_blank', 'noopener,noreferrer')

      if (!documentPreviewUrls[document.clientId]) {
        scheduleObjectUrlRevoke(previewUrl)
      }
    } catch (error) {
      toast.error(getApiErrorMessage(error) || t('pages.feedback.documentPreviewFailed'))
    } finally {
      setActiveDocumentAction((current) => ({
        ...current,
        [document.clientId]: undefined,
      }))
    }
  }

  async function handleDownloadDocument(
    document: PageSectionDocumentState,
    fileName: string,
  ) {
    setActiveDocumentAction((current) => ({
      ...current,
      [document.clientId]: 'download',
    }))

    try {
      const downloadUrl = await createTemporaryDocumentObjectUrl(document)
      triggerFileDownload(downloadUrl, fileName)
      scheduleObjectUrlRevoke(downloadUrl)
    } catch (error) {
      toast.error(getApiErrorMessage(error) || t('pages.feedback.documentDownloadFailed'))
    } finally {
      setActiveDocumentAction((current) => ({
        ...current,
        [document.clientId]: undefined,
      }))
    }
  }

  async function submitForm(submitMode: SubmitMode) {
    const submissionState: PageFormState = {
      ...form,
      status: submitMode === 'publish' ? 'published' : 'draft',
    }

    const heroImageValidationMessage =
      submissionState.heroImageEnabled && submissionState.heroImageFile
        ? getHeroImageValidationMessage(submissionState.heroImageFile, t)
        : null
    const validationErrors = {
      ...validatePageForm(submissionState, t, {
        currentPageId: isEditMode ? parsedPageId : null,
        pageOptions: parentPageOptions,
      }),
      ...(heroImageValidationMessage
        ? { heroImageFile: heroImageValidationMessage }
        : {}),
    }
    const nextSectionFieldErrors = validatePageSectionFields(submissionState.sections, t)
    const nextDocumentFieldErrors = validatePageDocumentFields(submissionState.sections, t)
    const validationMessages = [
      ...Object.values(validationErrors).filter((message): message is string => Boolean(message)),
      ...getPageSectionValidationMessages(nextSectionFieldErrors),
      ...getPageDocumentValidationMessages(nextDocumentFieldErrors),
    ]

    if (validationMessages.length > 0) {
      setErrors(validationErrors)
      setSectionFieldErrors(nextSectionFieldErrors)
      setDocumentFieldErrors(nextDocumentFieldErrors)
      toast.error(validationMessages[0] ?? t('pages.feedback.validation'))
      return
    }

    setErrors({})
    setSectionFieldErrors({})
    setDocumentFieldErrors({})

    try {
      const payload = buildSavePageRequest(submissionState, selectedParentPageSlug)
      const result =
        isEditMode && parsedPageId
          ? await dispatch(updatePage({ id: parsedPageId, payload })).unwrap()
          : await dispatch(createPage(payload)).unwrap()

      if (isEditMode && parsedPageId) {
        await dispatch(fetchPageById(parsedPageId)).unwrap()
        toast.success(
          submitMode === 'publish'
            ? t('pages.feedback.published')
            : t('pages.feedback.draftSaved'),
        )
      } else {
        toast.success(t('pages.feedback.created'))
        navigate(`/pages/${result.page.id}/edit`, { replace: true })
      }
    } catch {
      return
    }
  }

  function renderSectionBody(section: PageSectionState) {
    switch (section.sectionType) {
      case 'header':
        const heroHierarchyUnavailable =
          section.header.hierarchy !== 'h1_hero' && countHeroHeaderSections(form.sections) > 0
        const underlineAvailable = section.header.hierarchy === 'h2_section'
        const underlineEnabled = underlineAvailable
          ? (section.header.underlineEnabled ?? false)
          : false
        const setHeaderHierarchy = (hierarchy: PageSectionState['header']['hierarchy']) => {
          if (hierarchy === 'h1_hero' && heroHierarchyUnavailable) {
            return
          }

          updateSection(section.clientId, (current) => ({
            ...current,
            header: {
              ...current.header,
              hierarchy,
              underlineEnabled:
                hierarchy === 'h2_section' && current.header.hierarchy === 'h2_section'
                  ? (current.header.underlineEnabled ?? false)
                  : false,
            },
          }))
        }

        return (
          <div className={[styles.fieldStack, styles.documentSection].join(' ')}>
            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.sectionName')}</span>
                <input
                  type="text"
                  value={section.sectionName}
                  placeholder={t('pages.modules.placeholders.sectionName')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      sectionName: event.target.value,
                    }))
                  }
                />
              </label>

              <label className={styles.field}>
                <span>{t('pages.modules.fields.mainHeaderText')}</span>
                <input
                  type="text"
                  value={section.header.mainHeaderText}
                  placeholder={t('pages.modules.placeholders.mainHeaderText')}
                  onChange={(event) => {
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      header: {
                        ...current.header,
                        mainHeaderText: event.target.value,
                      },
                    }))
                    setSectionFieldError(
                      section.clientId,
                      'mainHeaderText',
                      event.target.value.trim()
                        ? undefined
                        : t('pages.validation.mainHeaderTextRequired'),
                    )
                  }}
                />
                <FieldError message={getSectionFieldError(section.clientId, 'mainHeaderText')} />
              </label>
            </div>

            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.subHeaderText')}</span>
                <input
                  type="text"
                  value={section.header.subHeaderText}
                  placeholder={t('pages.modules.placeholders.subHeaderText')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      header: {
                        ...current.header,
                        subHeaderText: event.target.value,
                      },
                    }))
                  }
                />
              </label>

              <div className={styles.field}>
                <span>{t('pages.modules.fields.hierarchy')}</span>
                <div className={styles.segmentedControl}>
                  <SegmentedButton
                    active={section.header.hierarchy === 'h1_hero'}
                    onClick={() => setHeaderHierarchy('h1_hero')}
                    label={t('pages.modules.options.h1Hero')}
                    disabled={heroHierarchyUnavailable}
                  />
                  <SegmentedButton
                    active={section.header.hierarchy === 'h2_section'}
                    onClick={() => setHeaderHierarchy('h2_section')}
                    label={t('pages.modules.options.h2Section')}
                  />
                </div>
                {heroHierarchyUnavailable ? (
                  <p className={styles.fieldHint}>{t('pages.modules.hints.h1HeroUnavailable')}</p>
                ) : null}
              </div>
            </div>

            <label className={styles.field}>
              <span>{t('pages.modules.fields.headerDescription')}</span>
              <textarea
                rows={4}
                value={section.header.description}
                placeholder={t('pages.modules.placeholders.headerDescription')}
                onChange={(event) =>
                  updateSection(section.clientId, (current) => ({
                    ...current,
                    header: {
                      ...current.header,
                      description: event.target.value,
                    },
                  }))
                }
              />
              <p className={styles.fieldHint}>{t('pages.modules.hints.headerDescription')}</p>
            </label>

            <div className={styles.field}>
              <span>{t('pages.modules.fields.textAlign')}</span>
              <div className={styles.segmentedControl}>
                <SegmentedButton
                  active={section.header.textAlign === 'left'}
                  onClick={() =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      header: {
                        ...current.header,
                        textAlign: 'left',
                      },
                    }))
                  }
                  label={t('pages.modules.options.alignLeft')}
                />
                <SegmentedButton
                  active={section.header.textAlign === 'center'}
                  onClick={() =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      header: {
                        ...current.header,
                        textAlign: 'center',
                      },
                    }))
                  }
                  label={t('pages.modules.options.alignCenter')}
                />
                <SegmentedButton
                  active={section.header.textAlign === 'right'}
                  onClick={() =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      header: {
                        ...current.header,
                        textAlign: 'right',
                      },
                    }))
                  }
                  label={t('pages.modules.options.alignRight')}
                />
              </div>
            </div>

            <div className={styles.optionGridSingle}>
              <ModuleOptionCard
                eyebrow={t('pages.modules.options.sectionAccent')}
                title={t('pages.modules.fields.underlineEnabled')}
                description={
                  underlineAvailable
                    ? t('pages.modules.hints.underlineEnabled')
                    : t('pages.modules.hints.underlineUnavailable')
                }
                checkedLabel={t('pages.modules.states.on')}
                uncheckedLabel={t('pages.modules.states.off')}
                disabledLabel={t('pages.modules.states.h2Only')}
                checked={underlineEnabled}
                disabled={!underlineAvailable}
                onToggle={(nextChecked) =>
                  updateSection(section.clientId, (current) => ({
                    ...current,
                    header: {
                      ...current.header,
                      underlineEnabled: nextChecked,
                    },
                  }))
                }
              />
            </div>
          </div>
        )

      case 'typography':
        return (
          <div className={styles.fieldStack}>
            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.sectionName')}</span>
                <input
                  type="text"
                  value={section.sectionName}
                  placeholder={t('pages.modules.placeholders.sectionName')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      sectionName: event.target.value,
                    }))
                  }
                />
              </label>

              <div className={styles.field}>
                <span>{t('pages.modules.fields.textAlign')}</span>
                <div className={styles.segmentedControl}>
                  <SegmentedButton
                    active={section.typography.textAlign === 'left'}
                    onClick={() =>
                      updateSection(section.clientId, (current) => ({
                        ...current,
                        typography: {
                          ...current.typography,
                          textAlign: 'left',
                        },
                      }))
                    }
                    label={t('pages.modules.options.alignLeft')}
                  />
                  <SegmentedButton
                    active={section.typography.textAlign === 'center'}
                    onClick={() =>
                      updateSection(section.clientId, (current) => ({
                        ...current,
                        typography: {
                          ...current.typography,
                          textAlign: 'center',
                        },
                      }))
                    }
                    label={t('pages.modules.options.alignCenter')}
                  />
                  <SegmentedButton
                    active={section.typography.textAlign === 'right'}
                    onClick={() =>
                      updateSection(section.clientId, (current) => ({
                        ...current,
                        typography: {
                          ...current.typography,
                          textAlign: 'right',
                        },
                      }))
                    }
                    label={t('pages.modules.options.alignRight')}
                  />
                </div>
              </div>
            </div>

            <div className={styles.field}>
              <span>{t('pages.modules.fields.htmlContent')}</span>
              <RichTextEditor
                value={section.typography.htmlContent}
                onChange={(html) => {
                  updateSection(section.clientId, (current) => ({
                    ...current,
                    typography: {
                      ...current.typography,
                      htmlContent: html,
                    },
                  }))
                  setSectionFieldError(
                    section.clientId,
                    'htmlContent',
                    isTypographyContentEmpty(html)
                      ? t('pages.validation.typographyContentRequired')
                      : undefined,
                  )
                }}
                placeholder={t('pages.modules.placeholders.htmlContent')}
                disabled={isReadOnlyMode || isBusy}
                allowImages={false}
              />
              <FieldError message={getSectionFieldError(section.clientId, 'htmlContent')} />
            </div>
          </div>
        )

      case 'gallery':
        const showTitleDescriptionEnabled = section.gallery.showTitleDescription ?? true
        const autoScrollAvailable = section.gallery.viewMode === 'carousel'
        const autoScrollEnabled = autoScrollAvailable
          ? (section.gallery.autoScrollEnabled ?? false)
          : false
        const setGalleryViewMode = (viewMode: PageSectionState['gallery']['viewMode']) =>
          updateSection(section.clientId, (current) => ({
            ...current,
            gallery: {
              ...current.gallery,
              viewMode,
              autoScrollEnabled:
                viewMode === 'carousel' && current.gallery.viewMode === 'carousel'
                  ? (current.gallery.autoScrollEnabled ?? false)
                  : false,
            },
          }))

        return (
          <div className={styles.fieldStack}>
            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.sectionName')}</span>
                <input
                  type="text"
                  value={section.sectionName}
                  placeholder={t('pages.modules.placeholders.sectionName')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      sectionName: event.target.value,
                    }))
                  }
                />
              </label>

              <label className={styles.field}>
                <span>{t('pages.modules.fields.gallery')}</span>
                <select
                  value={section.gallery.galleryId}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      gallery: {
                        ...current.gallery,
                        galleryId: event.target.value,
                      },
                    }))
                  }
                  disabled={galleryOptionsStatus === 'loading'}
                >
                  <option value="">{t('pages.modules.placeholders.gallery')}</option>
                  {galleryOptions.map((gallery) => (
                    <option key={gallery.id} value={gallery.id}>
                      {gallery.name}
                    </option>
                  ))}
                </select>
                <p className={styles.fieldHint}>
                  {galleryOptionsError || t('pages.modules.hints.gallery')}
                </p>
              </label>
            </div>

            <div className={styles.field}>
              <span>{t('pages.modules.fields.viewMode')}</span>
              <div className={styles.segmentedControl}>
                <SegmentedButton
                  active={section.gallery.viewMode === 'grid'}
                  onClick={() => setGalleryViewMode('grid')}
                  label={t('pages.modules.options.grid')}
                />
                <SegmentedButton
                  active={section.gallery.viewMode === 'carousel'}
                  onClick={() => setGalleryViewMode('carousel')}
                  label={t('pages.modules.options.carousel')}
                />
                <SegmentedButton
                  active={section.gallery.viewMode === 'masonry'}
                  onClick={() => setGalleryViewMode('masonry')}
                  label={t('pages.modules.options.masonry')}
                />
                <SegmentedButton
                  active={section.gallery.viewMode === 'focus'}
                  onClick={() => setGalleryViewMode('focus')}
                  label={t('pages.modules.options.focus')}
                />
                <SegmentedButton
                  active={section.gallery.viewMode === 'icons'}
                  onClick={() => setGalleryViewMode('icons')}
                  label={t('pages.modules.options.icons')}
                />
              </div>
            </div>

            <div className={styles.optionGrid}>
              <ModuleOptionCard
                eyebrow={t('pages.modules.options.captionDisplay')}
                title={t('pages.modules.fields.showTitleDescription')}
                description={t('pages.modules.hints.showTitleDescription')}
                checkedLabel={t('pages.modules.states.on')}
                uncheckedLabel={t('pages.modules.states.off')}
                checked={showTitleDescriptionEnabled}
                disabled={false}
                onToggle={(nextChecked) =>
                  updateSection(section.clientId, (current) => ({
                    ...current,
                    gallery: {
                      ...current.gallery,
                      showTitleDescription: nextChecked,
                    },
                  }))
                }
              />

              <ModuleOptionCard
                eyebrow={t('pages.modules.options.carouselMotion')}
                title={t('pages.modules.fields.autoScrollEnabled')}
                description={
                  autoScrollAvailable
                    ? t('pages.modules.hints.autoScrollEnabled')
                    : t('pages.modules.hints.autoScrollUnavailable')
                }
                checkedLabel={t('pages.modules.states.on')}
                uncheckedLabel={t('pages.modules.states.off')}
                disabledLabel={t('pages.modules.states.carouselOnly')}
                checked={autoScrollEnabled}
                disabled={!autoScrollAvailable}
                onToggle={(nextChecked) =>
                  updateSection(section.clientId, (current) => ({
                    ...current,
                    gallery: {
                      ...current.gallery,
                      autoScrollEnabled: nextChecked,
                    },
                  }))
                }
              />
            </div>
          </div>
        )

      case 'video':
        return (
          <div className={styles.fieldStack}>
            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.sectionName')}</span>
                <input
                  type="text"
                  value={section.sectionName}
                  placeholder={t('pages.modules.placeholders.sectionName')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      sectionName: event.target.value,
                    }))
                  }
                />
              </label>

              <label className={styles.field}>
                <span>{t('pages.modules.fields.videoPackage')}</span>
                <select
                  value={section.video.videoPackageId}
                  onChange={(event) => {
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      video: {
                        ...current.video,
                        videoPackageId: event.target.value,
                      },
                    }))
                    setSectionFieldError(
                      section.clientId,
                      'videoPackage',
                      event.target.value.trim()
                        ? undefined
                        : t('pages.validation.videoPackageRequired'),
                    )
                  }}
                  disabled={videoOptionsStatus === 'loading'}
                >
                  <option value="">{t('pages.modules.placeholders.videoPackage')}</option>
                  {videoOptions.map((item) => (
                    <option key={item.id} value={item.id}>
                      {item.title}
                    </option>
                  ))}
                </select>
                <p className={styles.fieldHint}>
                  {videoOptionsError || t('pages.modules.hints.videoPackage')}
                </p>
                <FieldError message={getSectionFieldError(section.clientId, 'videoPackage')} />
              </label>
            </div>
          </div>
        )

      case 'document':
        return (
          <div className={styles.fieldStack}>
            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.sectionName')}</span>
                <input
                  type="text"
                  value={section.sectionName}
                  placeholder={t('pages.modules.placeholders.sectionName')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      sectionName: event.target.value,
                    }))
                  }
                />
              </label>

              <div className={styles.field}>
                <span>{t('pages.modules.fields.documents')}</span>
                <p className={styles.fieldHint}>{t('pages.modules.hints.documents')}</p>
                <FieldError message={getSectionFieldError(section.clientId, 'documents')} />
              </div>
            </div>

            <UploadDropzone
              multiple
              accept={DOCUMENT_UPLOAD_ACCEPT}
              variant="compact"
              className={styles.documentDropzone}
              disabled={isReadOnlyMode || isBusy}
              icon={<CloudUploadIcon size={20} />}
              label={t('pages.modules.document.dropLabel')}
              hint={t('pages.modules.document.dropHint')}
              onFiles={(files) => handleDocumentFiles(section.clientId, files)}
            />

            {section.documents.items.length > 0 && (
              <div className={styles.documentList}>
                {section.documents.items.map((document, documentIndex) => {
                  const existingFileName =
                    document.originalFileName ||
                    document.file?.name ||
                    t('pages.modules.document.untitled')
                  const activeFileSize = document.file?.size ?? document.existingFileSize
                  const documentTitle =
                    document.displayName || stripFileExtension(existingFileName)
                  const documentTypeLabel = resolveDocumentTypeLabel(
                    document.file?.type || document.existingMimeType,
                    existingFileName,
                  )
                  const canPreview = canPreviewDocument(
                    document.file?.type || document.existingMimeType,
                    existingFileName,
                  )
                  const previewUrl = documentPreviewUrls[document.clientId] || ''
                  const documentMeta = buildDocumentMeta({
                    isPendingUpload: Boolean(document.file),
                    fileTypeLabel: documentTypeLabel,
                    fileSize: activeFileSize,
                    t,
                  })

                  return (
                    <article key={document.clientId} className={styles.documentItem}>
                      <div className={styles.documentHeader}>
                        <div className={styles.documentHeaderContent}>
                          <p className={styles.documentLabel}>
                            {t('pages.modules.document.itemLabel', {
                              index: documentIndex + 1,
                            })}
                          </p>
                          <div className={styles.documentPreviewRow}>
                            <div className={styles.documentPreviewCard} aria-hidden="true">
                              {canPreview && previewUrl ? (
                                <iframe
                                  src={previewUrl}
                                  title={documentTitle}
                                  className={styles.documentPreviewFrame}
                                />
                              ) : (
                                <>
                                  <PagesIcon size={28} className={styles.documentPreviewIcon} />
                                  <span className={styles.documentPreviewBadge}>
                                    {documentTypeLabel}
                                  </span>
                                </>
                              )}
                            </div>
                            <div className={styles.documentPreviewContent}>
                              <h4 className={styles.documentName}>{documentTitle}</h4>
                              {documentMeta ? (
                                <p className={styles.documentMeta}>{documentMeta}</p>
                              ) : null}
                              <div className={styles.documentPreviewActions}>
                                {canPreview ? (
                                  <button
                                    type="button"
                                    className={styles.secondaryButton}
                                    disabled={activeDocumentAction[document.clientId] !== undefined}
                                    onClick={() => void handlePreviewDocument(document)}
                                  >
                                    {activeDocumentAction[document.clientId] === 'preview'
                                      ? t('pages.common.loading')
                                      : t('pages.modules.actions.previewDocument')}
                                  </button>
                                ) : null}
                                <button
                                  type="button"
                                  className={styles.secondaryButton}
                                  disabled={activeDocumentAction[document.clientId] !== undefined}
                                  onClick={() =>
                                    void handleDownloadDocument(document, existingFileName)
                                  }
                                >
                                  <DownloadIcon size={16} />
                                  {activeDocumentAction[document.clientId] === 'download'
                                    ? t('pages.common.loading')
                                    : t('pages.modules.actions.downloadDocument')}
                                </button>
                              </div>
                            </div>
                          </div>
                        </div>

                        {!isReadOnlyMode && (
                          <div className={styles.documentActionRow}>
                            <label className={styles.secondaryButton}>
                              <input
                                type="file"
                                hidden
                                accept={DOCUMENT_UPLOAD_ACCEPT}
                                onChange={(event) => {
                                  const file = event.target.files?.[0]
                                  if (file) {
                                    handleReplaceDocumentFile(
                                      section.clientId,
                                      document.clientId,
                                      file,
                                    )
                                  }
                                  event.target.value = ''
                                }}
                              />
                              <EditIcon size={14} />
                              {t('pages.modules.actions.replaceFile')}
                            </label>
                            <button
                              type="button"
                              className={styles.iconButton}
                              disabled={documentIndex === 0}
                              aria-label={t('pages.modules.actions.moveUp')}
                              onClick={() =>
                                handleMoveDocument(section.clientId, document.clientId, -1)
                              }
                            >
                              <ArrowUpIcon />
                            </button>
                            <button
                              type="button"
                              className={styles.iconButton}
                              disabled={documentIndex === section.documents.items.length - 1}
                              aria-label={t('pages.modules.actions.moveDown')}
                              onClick={() =>
                                handleMoveDocument(section.clientId, document.clientId, 1)
                              }
                            >
                              <ArrowDownIcon />
                            </button>
                            <button
                              type="button"
                              className={styles.iconButtonDanger}
                              aria-label={t('pages.modules.actions.removeDocument')}
                              onClick={() =>
                                handleRemoveDocument(section.clientId, document.clientId)
                              }
                            >
                              <DeleteIcon size={16} />
                            </button>
                          </div>
                        )}
                      </div>

                      <div className={styles.fieldGrid}>
                        <label className={styles.field}>
                          <span>{t('pages.modules.fields.documentDisplayName')}</span>
                          <input
                            type="text"
                            value={document.displayName}
                            placeholder={t('pages.modules.placeholders.documentDisplayName')}
                            onChange={(event) => {
                              updateDocument(section.clientId, document.clientId, (current) => ({
                                ...current,
                                displayName: event.target.value,
                              }))
                              setDocumentFieldError(
                                document.clientId,
                                'displayName',
                                event.target.value.trim()
                                  ? undefined
                                  : t('pages.validation.documentDisplayNameRequired'),
                              )
                            }}
                          />
                          <FieldError
                            message={getDocumentFieldError(document.clientId, 'displayName')}
                          />
                        </label>

                        <label className={styles.field}>
                          <span>{t('pages.modules.fields.documentDescription')}</span>
                          <input
                            type="text"
                            value={document.description}
                            placeholder={t('pages.modules.placeholders.documentDescription')}
                            onChange={(event) =>
                              updateDocument(section.clientId, document.clientId, (current) => ({
                                ...current,
                                description: event.target.value,
                              }))
                            }
                          />
                        </label>
                      </div>
                    </article>
                  )
                })}
              </div>
            )}
          </div>
        )

      case 'quote':
        return (
          <div className={styles.fieldStack}>
            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.sectionName')}</span>
                <input
                  type="text"
                  value={section.sectionName}
                  placeholder={t('pages.modules.placeholders.sectionName')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      sectionName: event.target.value,
                    }))
                  }
                />
              </label>

              <label className={styles.field}>
                <span>{t('pages.modules.fields.attribution')}</span>
                <input
                  type="text"
                  value={section.quote.attribution}
                  placeholder={t('pages.modules.placeholders.attribution')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      quote: {
                        ...current.quote,
                        attribution: event.target.value,
                      },
                    }))
                  }
                />
              </label>
            </div>

            <label className={styles.field}>
              <span>{t('pages.modules.fields.quoteContent')}</span>
              <textarea
                rows={5}
                value={section.quote.quoteContent}
                placeholder={t('pages.modules.placeholders.quoteContent')}
                onChange={(event) => {
                  updateSection(section.clientId, (current) => ({
                    ...current,
                    quote: {
                      ...current.quote,
                      quoteContent: event.target.value,
                    },
                  }))
                  setSectionFieldError(
                    section.clientId,
                    'quoteContent',
                    event.target.value.trim()
                      ? undefined
                      : t('pages.validation.quoteContentRequired'),
                  )
                }}
              />
              <FieldError message={getSectionFieldError(section.clientId, 'quoteContent')} />
            </label>
          </div>
        )

      case 'cta_banner': {
        const currentCTAImagePreview = ctaImagePreviewUrls[section.clientId] || ''
        const hasCTABannerImage = Boolean(
          section.ctaBanner.imageFile || section.ctaBanner.existingImageFetchUrl,
        )

        return (
          <div className={styles.fieldStack}>
            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.sectionName')}</span>
                <input
                  type="text"
                  value={section.sectionName}
                  placeholder={t('pages.modules.placeholders.sectionName')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      sectionName: event.target.value,
                    }))
                  }
                />
              </label>

              <label className={styles.field}>
                <span>{t('pages.modules.fields.bannerHeading')}</span>
                <input
                  type="text"
                  value={section.ctaBanner.bannerHeading}
                  placeholder={t('pages.modules.placeholders.bannerHeading')}
                  onChange={(event) => {
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      ctaBanner: {
                        ...current.ctaBanner,
                        bannerHeading: event.target.value,
                      },
                    }))
                    setSectionFieldError(
                      section.clientId,
                      'bannerHeading',
                      event.target.value.trim()
                        ? undefined
                        : t('pages.validation.bannerHeadingRequired'),
                    )
                  }}
                />
                <FieldError message={getSectionFieldError(section.clientId, 'bannerHeading')} />
              </label>
            </div>

            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.bannerMessage')}</span>
                <input
                  type="text"
                  value={section.ctaBanner.bannerMessage}
                  placeholder={t('pages.modules.placeholders.bannerMessage')}
                  onChange={(event) =>
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      ctaBanner: {
                        ...current.ctaBanner,
                        bannerMessage: event.target.value,
                      },
                    }))
                  }
                />
              </label>

              <label className={styles.field}>
                <span>{t('pages.modules.fields.buttonText')}</span>
                <input
                  type="text"
                  value={section.ctaBanner.buttonText}
                  placeholder={t('pages.modules.placeholders.buttonText')}
                  onChange={(event) => {
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      ctaBanner: {
                        ...current.ctaBanner,
                        buttonText: event.target.value,
                      },
                    }))
                    setSectionFieldError(
                      section.clientId,
                      'buttonText',
                      event.target.value.trim()
                        ? undefined
                        : t('pages.validation.buttonTextRequired'),
                    )
                  }}
                />
                <FieldError message={getSectionFieldError(section.clientId, 'buttonText')} />
              </label>
            </div>

            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.modules.fields.buttonUrl')}</span>
                <input
                  type="url"
                  value={section.ctaBanner.buttonUrl}
                  placeholder={t('pages.modules.placeholders.buttonUrl')}
                  onChange={(event) => {
                    updateSection(section.clientId, (current) => ({
                      ...current,
                      ctaBanner: {
                        ...current.ctaBanner,
                        buttonUrl: event.target.value,
                      },
                    }))
                    setSectionFieldError(
                      section.clientId,
                      'buttonUrl',
                      event.target.value.trim()
                        ? undefined
                        : t('pages.validation.buttonUrlRequired'),
                    )
                  }}
                />
                <FieldError message={getSectionFieldError(section.clientId, 'buttonUrl')} />
              </label>

              <label className={styles.toggleRow}>
                <div>
                  <span>{t('pages.modules.fields.openInNewTab')}</span>
                  <p>{t('pages.modules.hints.openInNewTab')}</p>
                </div>
                <input
                  type="checkbox"
                  checked
                  disabled
                  readOnly
                />
              </label>
            </div>

            <div className={styles.fieldStack}>
              <p className={styles.fieldHint}>{t('pages.modules.hints.ctaImage')}</p>
              <div className={styles.heroPanel}>
                {hasCTABannerImage ? (
                  <div className={styles.heroPreviewCard}>
                    {currentCTAImagePreview ? (
                      <img
                        src={currentCTAImagePreview}
                        alt={section.ctaBanner.bannerHeading || section.sectionName || 'CTA image'}
                        className={styles.heroPreview}
                      />
                    ) : (
                      <div className={styles.mediaPreviewPlaceholder} />
                    )}
                    {!isReadOnlyMode && (
                      <div className={styles.heroActions}>
                        <label className={styles.secondaryButton}>
                          <input
                            type="file"
                            accept={IMAGE_UPLOAD_ACCEPT}
                            hidden
                            onChange={(event) => {
                              const file = event.target.files?.[0]
                              if (file) {
                                handleCTABannerImageFiles(section.clientId, [file])
                              }
                              event.target.value = ''
                            }}
                          />
                          {t('pages.editor.replaceHeroImage')}
                        </label>
                        <button
                          type="button"
                          className={styles.dangerButton}
                          onClick={() => removeCTABannerImage(section.clientId)}
                        >
                          {t('pages.editor.removeHeroImage')}
                        </button>
                      </div>
                    )}
                  </div>
                ) : (
                  <UploadDropzone
                    disabled={isReadOnlyMode || isBusy}
                    icon={<CloudUploadIcon size={24} />}
                    accept={IMAGE_UPLOAD_ACCEPT}
                    label={t('pages.modules.fields.ctaImage')}
                    hint={t('pages.modules.hints.ctaImage')}
                    onFiles={(files) => handleCTABannerImageFiles(section.clientId, files)}
                  />
                )}
              </div>
            </div>
          </div>
        )
      }
    }
  }

  if (isInvalidEditId) {
    return (
      <CmsAppShell activeKey="pages">
        <section className={styles.errorPanel}>
          <h1>{t('pages.editor.invalidIdTitle')}</h1>
          <p>{t('pages.editor.invalidIdText')}</p>
          <button
            type="button"
            className={styles.secondaryButton}
            onClick={() => navigate('/pages')}
          >
            {t('pages.editor.backToList')}
          </button>
        </section>
      </CmsAppShell>
    )
  }

  if (isInitialLoad) {
    return (
      <CmsAppShell activeKey="pages">
        <div className={styles.loaderWrap}>
          <Loader label={t('pages.common.loading')} />
        </div>
      </CmsAppShell>
    )
  }

  if (isEditMode && detailState.error && !currentPage) {
    return (
      <CmsAppShell activeKey="pages">
        <section className={styles.errorPanel}>
          <h1>{t('pages.editor.loadErrorTitle')}</h1>
          <p>{detailState.error}</p>
          <button
            type="button"
            className={styles.secondaryButton}
            onClick={() => navigate('/pages')}
          >
            {t('pages.editor.backToList')}
          </button>
        </section>
      </CmsAppShell>
    )
  }

  return (
    <CmsAppShell activeKey="pages">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: t('pages.breadcrumb.pages'), to: '/pages' },
            {
              label: isReadOnlyMode
                ? t('pages.editor.breadcrumbView')
                : isEditMode
                  ? t('pages.editor.breadcrumbEdit')
                  : t('pages.editor.breadcrumbCreate'),
            },
          ]}
        />

        <header className={styles.header}>
          <div>
            <h1 className={styles.title}>
              {isReadOnlyMode
                ? t('pages.editor.titleView')
                : isEditMode
                  ? t('pages.editor.titleEdit')
                  : t('pages.editor.titleCreate')}
            </h1>
            <p className={styles.subtitle}>{t('pages.editor.subtitle')}</p>
          </div>

          <div className={styles.headerActions}>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => navigate('/pages')}
            >
              {t('pages.editor.backToList')}
            </button>
            {isViewMode && parsedPageId && !isModuleManaged && (
              <button
                type="button"
                className={styles.primaryButton}
                onClick={() => navigate(`/pages/${parsedPageId}/edit`)}
              >
                <EditIcon size={14} />
                {t('pages.editor.editPage')}
              </button>
            )}
          </div>
        </header>

        {isModuleManaged && (
          <div className={styles.infoSummary} role="status">
            <p className={styles.infoSummaryTitle}>{t('pages.editor.moduleManagedTitle')}</p>
            <p className={styles.infoSummaryText}>{t('pages.editor.moduleManagedNotice')}</p>
          </div>
        )}

        {!isReadOnlyMode && errorMessages.length > 0 && (
          <div className={styles.errorSummary} role="alert">
            <p className={styles.errorSummaryTitle}>{t('pages.feedback.validation')}</p>
            <ul>
              {errorMessages.map((message) => (
                <li key={message}>{message}</li>
              ))}
            </ul>
          </div>
        )}

        {!isReadOnlyMode && saveState.error && (
          <div className={styles.errorSummary} role="alert">
            <p className={styles.errorSummaryTitle}>{saveState.error}</p>
          </div>
        )}

        <fieldset className={styles.readOnlyFieldset} disabled={isReadOnlyMode || isBusy}>
          <section className={styles.card}>
            <div className={styles.cardHeader}>
              <h2>{t('pages.sections.pageSettings')}</h2>
              <div className={styles.statusSwitch}>
                <span
                  className={form.status === 'draft' ? styles.statusLabelActive : styles.statusLabel}
                >
                  {t('pages.status.draft')}
                </span>
                <button
                  type="button"
                  className={styles.switch}
                  role="switch"
                  aria-checked={form.status === 'published'}
                  onClick={() => handleStatusToggle(form.status !== 'published')}
                >
                  <span
                    className={[
                      styles.switchThumb,
                      form.status === 'published' ? styles.switchThumbOn : '',
                    ].join(' ')}
                  />
                </button>
                <span
                  className={
                    form.status === 'published' ? styles.statusLabelActive : styles.statusLabel
                  }
                >
                  {t('pages.status.live')}
                </span>
              </div>
            </div>

            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>{t('pages.fields.pageTitle')}</span>
                <input
                  type="text"
                  value={form.pageTitle}
                  placeholder={t('pages.fields.pageTitlePlaceholder')}
                  onChange={(event) => handlePageTitleChange(event.target.value)}
                />
                <FieldError message={errors.pageTitle} />
              </label>

              <label className={styles.field}>
                <span>{t('pages.fields.parentPage')}</span>
                <select
                  value={form.parentPageId}
                  onChange={(event) => handleParentPageChange(event.target.value)}
                  disabled={parentPageOptionsStatus === 'loading'}
                >
                  <option value="">{t('pages.fields.noParentPage')}</option>
                  {availableParentPageOptions.map((page) => (
                    <option
                      key={page.id}
                      value={page.id}
                      disabled={disallowedParentPageIds.has(page.id)}
                    >
                      {`${page.page_title} (${page.url_slug})`}
                    </option>
                  ))}
                </select>
                <p className={styles.fieldHint}>{t('pages.fields.parentPageHint')}</p>
                <FieldError message={errors.parentPageId ?? parentPageOptionsError ?? undefined} />
              </label>

              <label className={styles.field}>
                <span>{t('pages.fields.urlSlug')}</span>
                <div className={styles.slugField}>
                  <span className={styles.slugPrefix}>{slugPrefix}</span>
                  <input
                    type="text"
                    value={form.urlSlug}
                    placeholder={t('pages.fields.urlSlugPlaceholder')}
                    onChange={(event) => handleSlugChange(event.target.value)}
                  />
                </div>
                <p className={styles.fieldHint}>
                  {t('pages.fields.fullUrlSlugPreview', {
                    slug: effectiveUrlSlug || '/',
                  })}
                </p>
                <FieldError message={errors.urlSlug} />
              </label>
            </div>

            <div className={styles.heroSection}>
              <label className={styles.toggleRow}>
                <div>
                  <span>{t('pages.fields.enableHeroImage')}</span>
                  <p>{t('pages.fields.enableHeroImageHint')}</p>
                </div>
                <input
                  type="checkbox"
                  checked={form.heroImageEnabled}
                  onChange={(event) => handleHeroToggle(event.target.checked)}
                />
              </label>

              {form.heroImageEnabled && (
                <div className={styles.heroPanel}>
                  {currentHeroPreview ? (
                    <div className={styles.heroPreviewCard}>
                      <img
                        src={currentHeroPreview}
                        alt={form.pageTitle || t('pages.fields.heroImageAlt')}
                        className={styles.heroPreview}
                      />
                      {!isReadOnlyMode && (
                        <div className={styles.heroActions}>
                          <label className={styles.secondaryButton}>
                            <input
                              type="file"
                              accept={IMAGE_UPLOAD_ACCEPT}
                              hidden
                              onChange={(event) => {
                                const file = event.target.files?.[0]
                                if (file) {
                                  handleHeroFiles([file])
                                }
                                event.target.value = ''
                              }}
                            />
                            {t('pages.editor.replaceHeroImage')}
                          </label>
                          <button
                            type="button"
                            className={styles.dangerButton}
                            onClick={removeHeroImage}
                          >
                            {t('pages.editor.removeHeroImage')}
                          </button>
                        </div>
                      )}
                    </div>
                  ) : (
                    <UploadDropzone
                      disabled={isReadOnlyMode || isBusy}
                      icon={<CloudUploadIcon size={24} />}
                      accept={IMAGE_UPLOAD_ACCEPT}
                      label={t('pages.fields.heroImageLabel')}
                      hint={t('pages.fields.heroImageHint')}
                      onFiles={handleHeroFiles}
                    />
                  )}
                  <FieldError message={errors.heroImageFile} />
                </div>
              )}
            </div>
          </section>

          <section className={styles.card}>
            <div className={styles.cardHeaderBlock}>
              <h2>{t('pages.sections.metadata')}</h2>
              <p>{t('pages.sections.metadataHint')}</p>
            </div>

            <div className={styles.fieldStack}>
              <label className={styles.field}>
                <span>{t('pages.fields.seoPageTitle')}</span>
                <input
                  type="text"
                  value={form.seoPageTitle}
                  placeholder={t('pages.fields.seoPageTitlePlaceholder')}
                  onChange={(event) => updateField('seoPageTitle', event.target.value)}
                />
              </label>

              <label className={styles.field}>
                <span>{t('pages.fields.seoPageDescription')}</span>
                <textarea
                  rows={5}
                  value={form.seoPageDescription}
                  placeholder={t('pages.fields.seoPageDescriptionPlaceholder')}
                  onChange={(event) => updateField('seoPageDescription', event.target.value)}
                />
              </label>
            </div>
          </section>

          {!isModuleManaged && (
            <section className={styles.card}>
            <div className={styles.cardHeaderBlock}>
              <h2>{t('pages.sections.contentModules')}</h2>
              <p>{t('pages.sections.contentModulesHint')}</p>
            </div>

            {form.sections.length === 0 ? (
              <div className={styles.emptyState}>
                <div className={styles.emptyStateIcon}>
                  <PagesIcon size={18} />
                </div>
                <div>
                  <p className={styles.emptyStateTitle}>
                    {t('pages.sections.emptyModulesTitle')}
                  </p>
                  <p className={styles.emptyStateText}>
                    {t('pages.sections.emptyModulesText')}
                  </p>
                </div>
              </div>
            ) : (
              <div className={styles.moduleList}>
                {form.sections.map((section, sectionIndex) => {
                  const moduleType = moduleTypeOptions.find(
                    (option) => option.type === section.sectionType,
                  )
                  const isDragOverTarget =
                    dragOverSectionClientId === section.clientId &&
                    draggedSectionClientId !== section.clientId

                  return (
                    <article
                      key={section.clientId}
                      className={[
                        styles.moduleCard,
                        draggedSectionClientId === section.clientId ? styles.moduleCardDragging : '',
                        isDragOverTarget ? styles.moduleCardDragOver : '',
                        isDragOverTarget && dragOverSectionPlacement === 'before'
                          ? styles.moduleCardDragOverBefore
                          : '',
                        isDragOverTarget && dragOverSectionPlacement === 'after'
                          ? styles.moduleCardDragOverAfter
                          : '',
                      ]
                        .filter(Boolean)
                        .join(' ')}
                      draggable={!isReadOnlyMode && !isBusy}
                      onDragStart={() => handleSectionDragStart(section.clientId)}
                      onDragOver={(event) => handleSectionDragOver(event, section.clientId)}
                      onDrop={(event) => handleSectionDrop(event, section.clientId)}
                      onDragEnd={handleSectionDragEnd}
                    >
                      <div className={styles.moduleHeader}>
                        <div className={styles.moduleHeaderMain}>
                          {!isReadOnlyMode && (
                            <button
                              type="button"
                              className={styles.dragHandle}
                              aria-label={t('pages.modules.actions.dragToReorder')}
                            >
                              <DragHandleIcon />
                            </button>
                          )}

                          <div className={styles.moduleHeadingBlock}>
                            <div className={styles.moduleEyebrowRow}>
                              <span className={styles.moduleEyebrow}>
                                {moduleType?.label ?? section.sectionType}
                              </span>
                              {!section.isEnabled && (
                                <span className={styles.moduleBadge}>
                                  {t('pages.modules.badges.hidden')}
                                </span>
                              )}
                            </div>
                            <h3 className={styles.moduleTitle}>
                              {section.sectionName || moduleType?.label}
                            </h3>
                          </div>
                        </div>

                        <div className={styles.moduleHeaderActions}>
                          <label className={styles.moduleToggle}>
                            <span>{t('pages.modules.fields.enabled')}</span>
                            <input
                              type="checkbox"
                              checked={section.isEnabled}
                              onChange={(event) =>
                                updateSection(section.clientId, (current) => ({
                                  ...current,
                                  isEnabled: event.target.checked,
                                }))
                              }
                            />
                          </label>

                          {!isReadOnlyMode && (
                            <>
                              <button
                                type="button"
                                className={styles.iconButton}
                                disabled={sectionIndex === 0}
                                aria-label={t('pages.modules.actions.moveUp')}
                                onClick={() => handleMoveSection(section.clientId, -1)}
                              >
                                <ArrowUpIcon />
                              </button>
                              <button
                                type="button"
                                className={styles.iconButton}
                                disabled={sectionIndex === form.sections.length - 1}
                                aria-label={t('pages.modules.actions.moveDown')}
                                onClick={() => handleMoveSection(section.clientId, 1)}
                              >
                                <ArrowDownIcon />
                              </button>
                            </>
                          )}

                          <button
                            type="button"
                            className={styles.iconButton}
                            aria-label={
                              section.isCollapsed
                                ? t('pages.modules.actions.expand')
                                : t('pages.modules.actions.collapse')
                            }
                            onClick={() =>
                              updateSection(section.clientId, (current) => ({
                                ...current,
                                isCollapsed: !current.isCollapsed,
                              }))
                            }
                          >
                            {section.isCollapsed ? <ChevronDownIcon /> : <ChevronUpIcon />}
                          </button>

                          {!isReadOnlyMode && (
                            <button
                              type="button"
                              className={styles.iconButtonDanger}
                              aria-label={t('pages.modules.actions.removeSection')}
                              onClick={() => handleRemoveSection(section.clientId)}
                            >
                              <DeleteIcon size={16} />
                            </button>
                          )}
                        </div>
                      </div>

                      {!section.isCollapsed && (
                        <div className={styles.moduleBody}>
                          {renderSectionBody(section)}
                        </div>
                      )}
                    </article>
                  )
                })}
              </div>
            )}

            {!isReadOnlyMode && (
              <div className={styles.modulePickerWrap}>
                <button
                  type="button"
                  className={styles.primaryButton}
                  onClick={() => setModulePickerOpen((current) => !current)}
                >
                  <AddIcon size={16} />
                  {t('pages.modules.actions.addModule')}
                </button>

                {modulePickerOpen && (
                  <div className={styles.modulePicker}>
                    <p className={styles.modulePickerLabel}>
                      {t('pages.modules.actions.selectComponent')}
                    </p>
                    <div className={styles.modulePickerGrid}>
                      {moduleTypeOptions.map((option) => (
                        <button
                          key={option.type}
                          type="button"
                          className={styles.modulePickerItem}
                          onClick={() => handleAddSection(option.type)}
                        >
                          <span className={styles.modulePickerIcon}>{option.icon}</span>
                          <span className={styles.modulePickerText}>
                            <strong>{option.label}</strong>
                            <span>{option.description}</span>
                          </span>
                        </button>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}
            </section>
          )}
        </fieldset>

        {!isReadOnlyMode && (
          <section className={styles.actionsCard}>
            <div className={styles.actionRow}>
              <button
                type="button"
                className={styles.secondaryButton}
                disabled={isBusy}
                onClick={() => submitForm('draft')}
              >
                {t('pages.editor.saveDraft')}
              </button>
              <button
                type="button"
                className={styles.primaryButton}
                disabled={isBusy}
                onClick={() => submitForm('publish')}
              >
                {isBusy ? t('pages.common.loading') : t('pages.editor.publish')}
              </button>
            </div>
          </section>
        )}
      </div>
    </CmsAppShell>
  )
}

function FieldError({ message }: { message?: string }) {
  if (!message) {
    return null
  }

  return (
    <p role="alert" className={styles.fieldError}>
      {message}
    </p>
  )
}

function SegmentedButton({
  active,
  label,
  onClick,
  disabled = false,
}: {
  active: boolean
  label: string
  onClick: () => void
  disabled?: boolean
}) {
  return (
    <button
      type="button"
      className={[styles.segmentedButton, active ? styles.segmentedButtonActive : '']
        .filter(Boolean)
        .join(' ')}
      onClick={onClick}
      disabled={disabled}
    >
      {label}
    </button>
  )
}

function ModuleOptionCard({
  eyebrow,
  title,
  description,
  checkedLabel,
  uncheckedLabel,
  disabledLabel,
  checked,
  disabled,
  onToggle,
}: {
  eyebrow: string
  title: string
  description: string
  checkedLabel: string
  uncheckedLabel: string
  disabledLabel?: string
  checked: boolean
  disabled: boolean
  onToggle: (nextChecked: boolean) => void
}) {
  const statusLabel = disabled
    ? (disabledLabel ?? uncheckedLabel)
    : checked
      ? checkedLabel
      : uncheckedLabel

  return (
    <div
      className={[
        styles.optionCard,
        checked && !disabled ? styles.optionCardActive : '',
        disabled ? styles.optionCardDisabled : '',
      ]
        .filter(Boolean)
        .join(' ')}
    >
      <div className={styles.optionCardContent}>
        <span className={styles.optionCardEyebrow}>{eyebrow}</span>
        <div className={styles.optionCardHeader}>
          <h4 className={styles.optionCardTitle}>{title}</h4>
          <span
            className={[
              styles.optionStatus,
              disabled
                ? styles.optionStatusDisabled
                : checked
                  ? styles.optionStatusActive
                  : styles.optionStatusInactive,
            ].join(' ')}
          >
            {statusLabel}
          </span>
        </div>
        <p className={styles.optionCardDescription}>{description}</p>
      </div>

      <button
        type="button"
        className={[
          styles.optionSwitch,
          checked && !disabled ? styles.optionSwitchActive : '',
        ]
          .filter(Boolean)
          .join(' ')}
        role="switch"
        aria-checked={checked}
        aria-label={title}
        disabled={disabled}
        onClick={() => onToggle(!checked)}
      >
        <span
          className={[
            styles.optionSwitchThumb,
            checked && !disabled ? styles.optionSwitchThumbActive : '',
          ]
            .filter(Boolean)
            .join(' ')}
        />
      </button>
    </div>
  )
}

function moveItemByClientId<T extends { clientId: string }>(
  items: T[],
  clientId: string,
  direction: -1 | 1,
) {
  const fromIndex = items.findIndex((item) => item.clientId === clientId)
  const toIndex = fromIndex + direction
  if (fromIndex < 0 || toIndex < 0 || toIndex >= items.length) {
    return items
  }

  const next = items.slice()
  const [moved] = next.splice(fromIndex, 1)
  next.splice(toIndex, 0, moved)
  return next
}

function stripFileExtension(value: string) {
  return value.replace(/\.[^.]+$/, '')
}

function formatFileSize(size: number) {
  if (size < 1024) {
    return `${size} B`
  }
  if (size < 1024 * 1024) {
    return `${(size / 1024).toFixed(1)} KB`
  }
  return `${(size / (1024 * 1024)).toFixed(1)} MB`
}

function resolveDocumentTypeLabel(mimeType: string, fileName: string) {
  const extension = fileName.split('.').pop()?.trim().toUpperCase() ?? ''
  if (extension) {
    return extension.length > 5 ? extension.slice(0, 5) : extension
  }

  const normalizedMimeType = mimeType.trim().toLowerCase()
  if (normalizedMimeType.includes('pdf')) {
    return 'PDF'
  }
  if (normalizedMimeType.includes('word')) {
    return 'DOC'
  }
  if (normalizedMimeType.includes('excel') || normalizedMimeType.includes('sheet')) {
    return 'XLS'
  }
  if (
    normalizedMimeType.includes('powerpoint') ||
    normalizedMimeType.includes('presentation')
  ) {
    return 'PPT'
  }

  return 'FILE'
}

function buildDocumentMeta({
  isPendingUpload,
  fileTypeLabel,
  fileSize,
  t,
}: {
  isPendingUpload: boolean
  fileTypeLabel: string
  fileSize?: number
  t: (key: string) => string
}) {
  const parts = [
    isPendingUpload ? t('pages.modules.document.pendingUploadShort') : '',
    fileTypeLabel,
    typeof fileSize === 'number' && fileSize > 0 ? formatFileSize(fileSize) : '',
  ].filter(Boolean)

  return parts.join(' • ')
}

function getHeroImageValidationMessage(
  file: Pick<File, 'name' | 'size' | 'type'>,
  t: (key: string, options?: Record<string, unknown>) => string,
) {
  const validationError = validateHeroImageFile(file)
  if (!validationError) {
    return null
  }

  if (validationError === 'file-too-large') {
    return t('pages.validation.heroImageTooLarge', {
      maxSizeMb: PAGE_HERO_IMAGE_MAX_FILE_SIZE_MB,
    })
  }

  return t('pages.validation.heroImageUnsupported')
}

function getDocumentUploadValidationMessage(file: Pick<File, 'name' | 'size' | 'type'>) {
  const validationError = validateResourceUploadFile(file)
  if (!validationError) {
    return null
  }

  return getResourceUploadValidationErrorMessage(validationError)
}

function canPreviewDocument(mimeType: string, fileName: string) {
  const normalizedMimeType = mimeType.trim().toLowerCase()
  const extension = fileName.split('.').pop()?.trim().toLowerCase() ?? ''

  return (
    normalizedMimeType.includes('pdf') ||
    normalizedMimeType.startsWith('image/') ||
    ['pdf', 'png', 'jpg', 'jpeg', 'webp', 'gif'].includes(extension)
  )
}

async function createTemporaryDocumentObjectUrl(document: PageSectionDocumentState) {
  if (document.file) {
    return URL.createObjectURL(document.file)
  }

  if (!document.existingFetchUrl) {
    throw new Error('Document URL unavailable')
  }

  const blob = await pagesApi.fetchPageDocumentContent(document.existingFetchUrl)
  return URL.createObjectURL(blob)
}

function triggerFileDownload(url: string, fileName: string) {
  const link = globalThis.document.createElement('a')
  link.href = url
  link.download = fileName
  link.rel = 'noreferrer'
  globalThis.document.body.appendChild(link)
  link.click()
  link.remove()
}

function scheduleObjectUrlRevoke(url: string) {
  globalThis.setTimeout(() => {
    URL.revokeObjectURL(url)
  }, 60_000)
}

function validateHeroImageFile(file: Pick<File, 'name' | 'size' | 'type'>) {
  if (!isSupportedHeroImageFile(file)) {
    return 'unsupported-file-type' as const
  }

  if (file.size > PAGE_HERO_IMAGE_MAX_FILE_SIZE_BYTES) {
    return 'file-too-large' as const
  }

  return null
}

function isSupportedHeroImageFile(file: Pick<File, 'name' | 'type'>) {
  const normalizedMimeType = file.type.trim().toLowerCase()
  if (normalizedMimeType && HERO_IMAGE_SUPPORTED_MIME_TYPES.has(normalizedMimeType)) {
    return true
  }

  const extension = getFileExtension(file.name)
  return extension ? HERO_IMAGE_SUPPORTED_EXTENSIONS.has(extension) : false
}

function getFileExtension(fileName: string) {
  const lastDotIndex = fileName.lastIndexOf('.')
  if (lastDotIndex < 0) {
    return ''
  }

  return fileName.slice(lastDotIndex).trim().toLowerCase()
}

function DragHandleIcon() {
  return <DragIndicatorOutlinedIcon fontSize="inherit" aria-hidden="true" />
}

function ArrowUpIcon() {
  return <KeyboardArrowUpOutlinedIcon fontSize="inherit" aria-hidden="true" />
}

function ArrowDownIcon() {
  return <KeyboardArrowDownOutlinedIcon fontSize="inherit" aria-hidden="true" />
}

function ChevronUpIcon() {
  return <ExpandLessOutlinedIcon fontSize="inherit" aria-hidden="true" />
}

function ChevronDownIcon() {
  return <ExpandMoreOutlinedIcon fontSize="inherit" aria-hidden="true" />
}

function TypographyModuleIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <path
        d="M3 4h10M8 4v8M5.2 12h5.6"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function GalleryModuleIcon() {
  return <GalleryIcon size={16} />
}

function VideoModuleIcon() {
  return <VideoIcon size={16} />
}

function DocumentModuleIcon() {
  return <PagesIcon size={16} />
}

function QuoteModuleIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <path
        d="M6.2 5.2H4.7A1.7 1.7 0 0 0 3 6.9v1.7a1.7 1.7 0 0 0 1.7 1.7h1.5A1.7 1.7 0 0 0 7.9 8.6V6.9A1.7 1.7 0 0 0 6.2 5.2ZM11.3 5.2H9.8A1.7 1.7 0 0 0 8.1 6.9v1.7a1.7 1.7 0 0 0 1.7 1.7h1.5A1.7 1.7 0 0 0 13 8.6V6.9a1.7 1.7 0 0 0-1.7-1.7Z"
        stroke="currentColor"
        strokeWidth="1.2"
      />
    </svg>
  )
}

function CtaModuleIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <rect x="2.5" y="4" width="11" height="8" rx="1.7" stroke="currentColor" strokeWidth="1.2" />
      <path
        d="M5 8h6M9 6.2 11 8 9 9.8"
        stroke="currentColor"
        strokeWidth="1.2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function HeaderModuleIcon() {
  return (
    <svg viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
      <path
        d="M4 3.5v9M12 3.5v9M4 8h8"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
      />
    </svg>
  )
}
