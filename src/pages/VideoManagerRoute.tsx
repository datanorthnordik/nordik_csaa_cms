import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import { getApiErrorMessage } from '../api/apiError'
import { type VideoItem, type VideoPackageDetail, type VideoPackageType, videoApi } from '../api/videoApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { ConfirmDialog } from '../components/cms/ConfirmDialog'
import { AddIcon, CloudUploadIcon, DeleteIcon } from '../components/icons'
import { Loader } from '../components/Loader'
import { UploadDropzone } from '../components/media/UploadDropzone'
import {
  getGalleryImageUploadValidationErrorMessage,
  RESOURCE_IMAGE_FILE_ACCEPT,
  validateGalleryImageUploadFile,
} from '../lib/resourceUpload'
import { getYouTubeEmbedUrl } from '../lib/videoPreview'
import styles from '../styles/VideoManagerPage.module.css'

type SingleVideoDraft = {
  youtubeUrl: string
  description: string
  teaserImageFile: File | null
  existingTeaserPath: string
  existingStorageUri: string
  existingObjectKey: string
}

type CollectionVideoDraft = {
  clientId: string
  id?: number
  title: string
  youtubeUrl: string
  description: string
  teaserImageFile: File | null
  existingTeaserPath: string
  existingStorageUri: string
  existingObjectKey: string
}

function createDraftId(prefix: string) {
  if (typeof globalThis.crypto?.randomUUID === 'function') {
    return `${prefix}-${globalThis.crypto.randomUUID()}`
  }
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}`
}

function createEmptySingleVideoDraft(): SingleVideoDraft {
  return {
    youtubeUrl: '',
    description: '',
    teaserImageFile: null,
    existingTeaserPath: '',
    existingStorageUri: '',
    existingObjectKey: '',
  }
}

function createCollectionVideoDraft(item?: VideoItem): CollectionVideoDraft {
  return {
    clientId: createDraftId('video-item'),
    id: item?.id,
    title: item?.title ?? '',
    youtubeUrl: item?.youtubeUrl ?? '',
    description: item?.description ?? '',
    teaserImageFile: null,
    existingTeaserPath: item?.teaserImagePath ?? '',
    existingStorageUri: item?.storageUri ?? '',
    existingObjectKey: item?.objectKey ?? '',
  }
}

export function VideoManagerRoute() {
  const { t } = useTranslation()
  const navigate = useNavigate()
  const { videoId } = useParams()
  const isCreateMode = !videoId
  const numericId = videoId ? Number.parseInt(videoId, 10) : Number.NaN

  const [videoPackage, setVideoPackage] = useState<VideoPackageDetail | null>(null)
  const [loading, setLoading] = useState(!isCreateMode)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [deleteDialogOpen, setDeleteDialogOpen] = useState(false)
  const [packageType, setPackageType] = useState<VideoPackageType>('single')
  const [packageTitle, setPackageTitle] = useState('')
  const [singleVideo, setSingleVideo] = useState<SingleVideoDraft>(createEmptySingleVideoDraft())
  const [existingItems, setExistingItems] = useState<CollectionVideoDraft[]>([])
  const [pendingItems, setPendingItems] = useState<CollectionVideoDraft[]>([])
  const [storedPreviewUrls, setStoredPreviewUrls] = useState<Record<string, string>>({})
  const [localPreviewUrls, setLocalPreviewUrls] = useState<Record<string, string>>({})

  function applyLoadedPackage(detail: VideoPackageDetail) {
    setVideoPackage(detail)
    setPackageType(detail.packageType)

    if (detail.packageType === 'single') {
      const source = detail.singleVideo ?? detail.videos[0]
      setPackageTitle(source?.title ?? detail.title)
      setSingleVideo({
        youtubeUrl: source?.youtubeUrl ?? '',
        description: source?.description ?? '',
        teaserImageFile: null,
        existingTeaserPath: source?.teaserImagePath ?? '',
        existingStorageUri: source?.storageUri ?? '',
        existingObjectKey: source?.objectKey ?? '',
      })
      setExistingItems([])
      setPendingItems([])
      return
    }

    setPackageTitle(detail.title)
    setSingleVideo(createEmptySingleVideoDraft())
    setExistingItems((detail.videos ?? []).map(createCollectionVideoDraft))
    setPendingItems([])
  }

  async function loadVideoPackage() {
    if (isCreateMode) {
      setVideoPackage(null)
      setLoading(false)
      setError(null)
      return
    }

    if (!Number.isFinite(numericId)) {
      setVideoPackage(null)
      setLoading(false)
      setError(t('videos.manager.notFound.text', { id: videoId ?? '' }))
      return
    }

    setLoading(true)
    setError(null)

    try {
      const detail = await videoApi.getVideoPackage(numericId)
      applyLoadedPackage(detail)
    } catch (loadError) {
      setVideoPackage(null)
      setError(getApiErrorMessage(loadError))
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    if (isCreateMode) {
      setPackageType('single')
      setPackageTitle('')
      setSingleVideo(createEmptySingleVideoDraft())
      setExistingItems([])
      setPendingItems([])
      setVideoPackage(null)
    }

    void loadVideoPackage()
  }, [isCreateMode, numericId, t, videoId])

  const previewSources = useMemo(() => {
    const sources: Array<{ key: string; path: string }> = []

    if (packageType === 'single' && singleVideo.existingTeaserPath) {
      sources.push({ key: 'single', path: singleVideo.existingTeaserPath })
    }

    for (const item of existingItems) {
      if (item.existingTeaserPath) {
        sources.push({ key: item.clientId, path: item.existingTeaserPath })
      }
    }

    return sources
  }, [existingItems, packageType, singleVideo.existingTeaserPath])

  useEffect(() => {
    const previewFiles: Array<{ key: string; file: File }> = []

    if (packageType === 'single' && singleVideo.teaserImageFile) {
      previewFiles.push({ key: 'single', file: singleVideo.teaserImageFile })
    }

    for (const item of existingItems) {
      if (item.teaserImageFile) {
        previewFiles.push({ key: item.clientId, file: item.teaserImageFile })
      }
    }

    for (const item of pendingItems) {
      if (item.teaserImageFile) {
        previewFiles.push({ key: item.clientId, file: item.teaserImageFile })
      }
    }

    if (!previewFiles.length) {
      setLocalPreviewUrls((current) => {
        Object.values(current).forEach((url) => URL.revokeObjectURL(url))
        return {}
      })
      return
    }

    const nextUrls = previewFiles.reduce<Record<string, string>>((result, entry) => {
      result[entry.key] = URL.createObjectURL(entry.file)
      return result
    }, {})

    setLocalPreviewUrls((current) => {
      Object.values(current).forEach((url) => URL.revokeObjectURL(url))
      return nextUrls
    })

    return () => {
      Object.values(nextUrls).forEach((url) => URL.revokeObjectURL(url))
    }
  }, [existingItems, packageType, pendingItems, singleVideo.teaserImageFile])

  useEffect(() => {
    if (!previewSources.length) {
      setStoredPreviewUrls((current) => {
        Object.values(current).forEach((url) => URL.revokeObjectURL(url))
        return {}
      })
      return
    }

    let cancelled = false

    async function loadPreviews() {
      const entries = await Promise.all(
        previewSources.map(async (entry) => {
          try {
            const blob = await videoApi.fetchVideoTeaserContent(entry.path)
            return [entry.key, URL.createObjectURL(blob)] as const
          } catch {
            return null
          }
        }),
      )

      const nextUrls = entries.reduce<Record<string, string>>((result, entry) => {
        if (entry) {
          result[entry[0]] = entry[1]
        }
        return result
      }, {})

      if (cancelled) {
        Object.values(nextUrls).forEach((url) => URL.revokeObjectURL(url))
        return
      }

      setStoredPreviewUrls((current) => {
        Object.values(current).forEach((url) => URL.revokeObjectURL(url))
        return nextUrls
      })
    }

    void loadPreviews()

    return () => {
      cancelled = true
    }
  }, [previewSources])

  function showError(message: string) {
    setError(message)
    toast.error(message)
  }

  function validateTeaserFile(file: File) {
    const validationError = validateGalleryImageUploadFile(file)
    if (!validationError) {
      return null
    }

    return getGalleryImageUploadValidationErrorMessage(validationError)
  }

  function updateExistingItem(
    clientId: string,
    updater: (item: CollectionVideoDraft) => CollectionVideoDraft,
  ) {
    setExistingItems((current) =>
      current.map((item) => (item.clientId === clientId ? updater(item) : item)),
    )
  }

  function updatePendingItem(
    clientId: string,
    updater: (item: CollectionVideoDraft) => CollectionVideoDraft,
  ) {
    setPendingItems((current) =>
      current.map((item) => (item.clientId === clientId ? updater(item) : item)),
    )
  }

  function validateSingleVideoDraft(requireTeaser: boolean) {
    if (!packageTitle.trim()) {
      return t('videos.validation.videoTitleRequired')
    }
    if (!singleVideo.youtubeUrl.trim()) {
      return t('videos.validation.youtubeUrlRequired')
    }
    if (
      requireTeaser &&
      !singleVideo.teaserImageFile &&
      !singleVideo.existingStorageUri.trim() &&
      !singleVideo.existingObjectKey.trim() &&
      !singleVideo.existingTeaserPath.trim()
    ) {
      return t('videos.validation.teaserRequired')
    }

    return null
  }

  function validateCollectionVideoDraft(item: CollectionVideoDraft, requireTeaser: boolean) {
    if (!item.title.trim()) {
      return t('videos.validation.videoTitleRequired')
    }
    if (!item.youtubeUrl.trim()) {
      return t('videos.validation.youtubeUrlRequired')
    }
    if (
      requireTeaser &&
      !item.teaserImageFile &&
      !item.existingStorageUri.trim() &&
      !item.existingObjectKey.trim() &&
      !item.existingTeaserPath.trim()
    ) {
      return t('videos.validation.teaserRequired')
    }

    return null
  }

  async function handleCreatePackage() {
    if (packageType === 'single') {
      const validationMessage = validateSingleVideoDraft(true)
      if (validationMessage) {
        showError(validationMessage)
        return
      }
    } else if (!packageTitle.trim()) {
      showError(t('videos.validation.packageTitleRequired'))
      return
    }

    const pendingValidation = pendingItems
      .map((item) => validateCollectionVideoDraft(item, true))
      .find(Boolean)
    if (pendingValidation) {
      showError(pendingValidation)
      return
    }

    setBusy(true)
    setError(null)

    try {
      const result =
        packageType === 'single'
          ? await videoApi.createVideoPackage({
              title: packageTitle.trim(),
              package_type: 'single',
              single_video: {
                title: packageTitle.trim(),
                youtube_url: singleVideo.youtubeUrl.trim(),
                description: singleVideo.description.trim(),
                ...(singleVideo.teaserImageFile
                  ? {
                      file_name: singleVideo.teaserImageFile.name,
                      mime_type:
                        singleVideo.teaserImageFile.type || 'application/octet-stream',
                    }
                  : {}),
              },
              singleVideoFile: singleVideo.teaserImageFile,
            })
          : await videoApi.createVideoPackage({
              title: packageTitle.trim(),
              package_type: 'collection',
              videos: pendingItems.map((item) => ({
                title: item.title.trim(),
                youtube_url: item.youtubeUrl.trim(),
                description: item.description.trim(),
                ...(item.teaserImageFile
                  ? {
                      file_name: item.teaserImageFile.name,
                      mime_type: item.teaserImageFile.type || 'application/octet-stream',
                    }
                  : {}),
              })),
              videoFiles: pendingItems.map((item) => item.teaserImageFile),
            })

      toast.success(result.message)
      navigate(`/videos/${result.video.id}`, { replace: true })
    } catch (saveError) {
      setError(getApiErrorMessage(saveError))
    } finally {
      setBusy(false)
    }
  }

  async function handleSaveCollectionTitle() {
    if (!Number.isFinite(numericId)) {
      return
    }
    if (!packageTitle.trim()) {
      showError(t('videos.validation.packageTitleRequired'))
      return
    }

    setBusy(true)
    setError(null)

    try {
      const result = await videoApi.updateVideoPackage(numericId, {
        title: packageTitle.trim(),
      })
      toast.success(result.message)
      await loadVideoPackage()
    } catch (saveError) {
      setError(getApiErrorMessage(saveError))
    } finally {
      setBusy(false)
    }
  }

  async function handleSaveSingleVideo() {
    if (!Number.isFinite(numericId) || !videoPackage?.videos[0]?.id) {
      return
    }

    const validationMessage = validateSingleVideoDraft(false)
    if (validationMessage) {
      showError(validationMessage)
      return
    }

    setBusy(true)
    setError(null)

    try {
      const result = await videoApi.updateVideoItem(numericId, videoPackage.videos[0].id, {
        title: packageTitle.trim(),
        youtube_url: singleVideo.youtubeUrl.trim(),
        description: singleVideo.description.trim(),
        ...(singleVideo.teaserImageFile
          ? {
              file_name: singleVideo.teaserImageFile.name,
              mime_type: singleVideo.teaserImageFile.type || 'application/octet-stream',
              teaserImageFile: singleVideo.teaserImageFile,
            }
          : {}),
      })
      toast.success(result.message)
      await loadVideoPackage()
    } catch (saveError) {
      setError(getApiErrorMessage(saveError))
    } finally {
      setBusy(false)
    }
  }

  async function handleSaveExistingItem(item: CollectionVideoDraft) {
    if (!Number.isFinite(numericId) || !item.id) {
      return
    }

    const validationMessage = validateCollectionVideoDraft(item, false)
    if (validationMessage) {
      showError(validationMessage)
      return
    }

    setBusy(true)
    setError(null)

    try {
      const result = await videoApi.updateVideoItem(numericId, item.id, {
        title: item.title.trim(),
        youtube_url: item.youtubeUrl.trim(),
        description: item.description.trim(),
        ...(item.teaserImageFile
          ? {
              file_name: item.teaserImageFile.name,
              mime_type: item.teaserImageFile.type || 'application/octet-stream',
              teaserImageFile: item.teaserImageFile,
            }
          : {}),
      })
      toast.success(result.message)
      await loadVideoPackage()
    } catch (saveError) {
      setError(getApiErrorMessage(saveError))
    } finally {
      setBusy(false)
    }
  }

  async function handleAddItems() {
    if (!Number.isFinite(numericId) || pendingItems.length === 0) {
      return
    }

    const validationMessage = pendingItems
      .map((item) => validateCollectionVideoDraft(item, true))
      .find(Boolean)
    if (validationMessage) {
      showError(validationMessage)
      return
    }

    setBusy(true)
    setError(null)

    try {
      const result = await videoApi.addVideoItems(numericId, {
        videos: pendingItems.map((item) => ({
          title: item.title.trim(),
          youtube_url: item.youtubeUrl.trim(),
          description: item.description.trim(),
          ...(item.teaserImageFile
            ? {
                file_name: item.teaserImageFile.name,
                mime_type: item.teaserImageFile.type || 'application/octet-stream',
              }
            : {}),
        })),
        videoFiles: pendingItems.map((item) => item.teaserImageFile),
      })
      toast.success(result.message)
      setPendingItems([])
      await loadVideoPackage()
    } catch (saveError) {
      setError(getApiErrorMessage(saveError))
    } finally {
      setBusy(false)
    }
  }

  async function handleDeleteItem(item: CollectionVideoDraft) {
    if (!Number.isFinite(numericId) || !item.id) {
      return
    }

    setBusy(true)
    setError(null)

    try {
      const result = await videoApi.deleteVideoItem(numericId, item.id)
      toast.success(result.message)
      await loadVideoPackage()
    } catch (deleteError) {
      setError(getApiErrorMessage(deleteError))
    } finally {
      setBusy(false)
    }
  }

  async function handleDeletePackage() {
    if (!Number.isFinite(numericId)) {
      return
    }

    setBusy(true)
    setError(null)

    try {
      const result = await videoApi.deleteVideoPackage(numericId)
      toast.success(result.message)
      navigate('/videos', { replace: true })
    } catch (deleteError) {
      setError(getApiErrorMessage(deleteError))
    } finally {
      setBusy(false)
      setDeleteDialogOpen(false)
    }
  }

  function handlePackageTypeChange(nextType: VideoPackageType) {
    setPackageType(nextType)
    setPackageTitle('')
    setSingleVideo(createEmptySingleVideoDraft())
    setExistingItems([])
    setPendingItems(nextType === 'collection' ? [createCollectionVideoDraft()] : [])
    setError(null)
  }

  function handleSingleTeaserFile(file?: File) {
    if (!file) {
      return
    }

    const validationMessage = validateTeaserFile(file)
    if (validationMessage) {
      showError(validationMessage)
      return
    }

    setSingleVideo((current) => ({
      ...current,
      teaserImageFile: file,
    }))
    setError(null)
  }

  function handleCollectionTeaserFile(
    scope: 'existing' | 'pending',
    clientId: string,
    file?: File,
  ) {
    if (!file) {
      return
    }

    const validationMessage = validateTeaserFile(file)
    if (validationMessage) {
      showError(validationMessage)
      return
    }

    const update = (item: CollectionVideoDraft) => ({
      ...item,
      teaserImageFile: file,
    })

    if (scope === 'existing') {
      updateExistingItem(clientId, update)
    } else {
      updatePendingItem(clientId, update)
    }
    setError(null)
  }

  function renderVideoPreview(youtubeUrl: string, title: string) {
    const embedUrl = getYouTubeEmbedUrl(youtubeUrl)
    if (!embedUrl) {
      return null
    }

    const previewTitle = title.trim() || t('videos.manager.untitledVideo')

    return (
      <div className={styles.previewCard}>
        <span className={styles.previewLabel}>{t('videos.manager.videoPreview')}</span>
        <div className={styles.videoPreviewFrameWrap}>
          <iframe
            src={embedUrl}
            title={`${t('videos.manager.videoPreview')}: ${previewTitle}`}
            className={styles.videoPreviewFrame}
            allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share"
            allowFullScreen
            referrerPolicy="strict-origin-when-cross-origin"
          />
        </div>
      </div>
    )
  }

  function renderSingleVideoSection() {
    const teaserPreview = localPreviewUrls.single ?? storedPreviewUrls.single

    return (
      <section className={styles.card}>
        <div className={styles.cardHeader}>
          <div>
            <h2>{t('videos.manager.sections.singleVideo')}</h2>
            <p>{t('videos.manager.sectionHint')}</p>
          </div>
        </div>

        <div className={styles.fieldGrid}>
          <label className={styles.field}>
            <span>{t('videos.manager.fields.videoTitle')}</span>
            <input
              type="text"
              value={packageTitle}
              placeholder={t('videos.manager.placeholders.videoTitle')}
              onChange={(event) => setPackageTitle(event.target.value)}
              disabled={busy}
            />
          </label>

          <label className={styles.field}>
            <span>{t('videos.manager.fields.youtubeUrl')}</span>
            <input
              type="url"
              value={singleVideo.youtubeUrl}
              placeholder={t('videos.manager.placeholders.youtubeUrl')}
              onChange={(event) =>
                setSingleVideo((current) => ({
                  ...current,
                  youtubeUrl: event.target.value,
                }))
              }
              disabled={busy}
            />
          </label>
        </div>

        <label className={styles.field}>
          <span>{t('videos.manager.fields.description')}</span>
          <textarea
            rows={5}
            value={singleVideo.description}
            placeholder={t('videos.manager.placeholders.description')}
            onChange={(event) =>
              setSingleVideo((current) => ({
                ...current,
                description: event.target.value,
              }))
            }
            disabled={busy}
          />
        </label>

        <div className={styles.mediaBlock}>
          {renderVideoPreview(singleVideo.youtubeUrl, packageTitle)}
          {teaserPreview ? (
            <div className={styles.previewCard}>
              <span className={styles.previewLabel}>{t('videos.manager.teaserPreview')}</span>
              <img src={teaserPreview} alt="" className={styles.previewImage} />
            </div>
          ) : null}

          <UploadDropzone
            accept={RESOURCE_IMAGE_FILE_ACCEPT}
            icon={<CloudUploadIcon size={20} />}
            label={t('videos.manager.dropLabel')}
            hint={t('videos.manager.dropHint')}
            disabled={busy}
            onFiles={(files) => handleSingleTeaserFile(files[0])}
          />
          {singleVideo.teaserImageFile ? (
            <p className={styles.selectedFile}>
              {t('videos.manager.selectedTeaser', {
                name: singleVideo.teaserImageFile.name,
              })}
            </p>
          ) : null}
        </div>

        <div className={styles.actionsRow}>
          <button
            type="button"
            className={styles.primaryButton}
            disabled={busy}
            onClick={() =>
              void (isCreateMode ? handleCreatePackage() : handleSaveSingleVideo())
            }
          >
            {isCreateMode
              ? t('videos.manager.actions.createSingle')
              : t('videos.manager.actions.saveSingle')}
          </button>

          {!isCreateMode ? (
            <button
              type="button"
              className={styles.dangerButton}
              disabled={busy}
              onClick={() => setDeleteDialogOpen(true)}
            >
              <DeleteIcon size={14} />
              {t('videos.manager.actions.deletePackage')}
            </button>
          ) : null}
        </div>
      </section>
    )
  }

  function renderCollectionItemCard(
    item: CollectionVideoDraft,
    options: {
      scope: 'existing' | 'pending'
      onUpdate: (updater: (current: CollectionVideoDraft) => CollectionVideoDraft) => void
      onPrimaryAction?: () => void
      primaryLabel?: string
      onRemove: () => void
      removeLabel: string
    },
  ) {
    const teaserPreview = localPreviewUrls[item.clientId] ?? storedPreviewUrls[item.clientId]

    return (
      <article key={item.clientId} className={styles.videoCard}>
        <div className={styles.videoCardHeader}>
          <h3>{item.title.trim() || t('videos.manager.untitledVideo')}</h3>
          <button
            type="button"
            className={styles.iconDangerButton}
            disabled={busy}
            onClick={options.onRemove}
          >
            <DeleteIcon size={14} />
            {options.removeLabel}
          </button>
        </div>

        <div className={styles.fieldGrid}>
          <label className={styles.field}>
            <span>{t('videos.manager.fields.videoTitle')}</span>
            <input
              type="text"
              value={item.title}
              placeholder={t('videos.manager.placeholders.videoTitle')}
              onChange={(event) =>
                options.onUpdate((current) => ({
                  ...current,
                  title: event.target.value,
                }))
              }
              disabled={busy}
            />
          </label>

          <label className={styles.field}>
            <span>{t('videos.manager.fields.youtubeUrl')}</span>
            <input
              type="url"
              value={item.youtubeUrl}
              placeholder={t('videos.manager.placeholders.youtubeUrl')}
              onChange={(event) =>
                options.onUpdate((current) => ({
                  ...current,
                  youtubeUrl: event.target.value,
                }))
              }
              disabled={busy}
            />
          </label>
        </div>

        <label className={styles.field}>
          <span>{t('videos.manager.fields.description')}</span>
          <textarea
            rows={4}
            value={item.description}
            placeholder={t('videos.manager.placeholders.description')}
            onChange={(event) =>
              options.onUpdate((current) => ({
                ...current,
                description: event.target.value,
              }))
            }
            disabled={busy}
          />
        </label>

        <div className={styles.mediaBlock}>
          {renderVideoPreview(item.youtubeUrl, item.title)}
          {teaserPreview ? (
            <div className={styles.previewCard}>
              <span className={styles.previewLabel}>{t('videos.manager.teaserPreview')}</span>
              <img src={teaserPreview} alt="" className={styles.previewImage} />
            </div>
          ) : null}

          <UploadDropzone
            accept={RESOURCE_IMAGE_FILE_ACCEPT}
            variant="compact"
            icon={<CloudUploadIcon size={18} />}
            label={t('videos.manager.dropLabel')}
            hint={t('videos.manager.dropHint')}
            disabled={busy}
            onFiles={(files) =>
              handleCollectionTeaserFile(options.scope, item.clientId, files[0])
            }
          />
          {item.teaserImageFile ? (
            <p className={styles.selectedFile}>
              {t('videos.manager.selectedTeaser', {
                name: item.teaserImageFile.name,
              })}
            </p>
          ) : null}
        </div>

        {options.onPrimaryAction && options.primaryLabel ? (
          <div className={styles.actionsRow}>
            <button
              type="button"
              className={styles.primaryButton}
              disabled={busy}
              onClick={options.onPrimaryAction}
            >
              {options.primaryLabel}
            </button>
          </div>
        ) : null}
      </article>
    )
  }

  function renderCollectionSection() {
    return (
      <>
        <section className={styles.card}>
          <div className={styles.cardHeader}>
            <div>
              <h2>{t('videos.manager.sections.collectionDetails')}</h2>
              <p>{t('videos.manager.collectionHint')}</p>
            </div>
          </div>

          <label className={styles.field}>
            <span>{t('videos.manager.fields.collectionTitle')}</span>
            <input
              type="text"
              value={packageTitle}
              placeholder={t('videos.manager.placeholders.collectionTitle')}
              onChange={(event) => setPackageTitle(event.target.value)}
              disabled={busy}
            />
          </label>

          <div className={styles.actionsRow}>
            <button
              type="button"
              className={styles.primaryButton}
              disabled={busy}
              onClick={() =>
                void (isCreateMode ? handleCreatePackage() : handleSaveCollectionTitle())
              }
            >
              {isCreateMode
                ? t('videos.manager.actions.createCollection')
                : t('videos.manager.actions.saveCollectionTitle')}
            </button>

            {!isCreateMode ? (
              <button
                type="button"
                className={styles.dangerButton}
                disabled={busy}
                onClick={() => setDeleteDialogOpen(true)}
              >
                <DeleteIcon size={14} />
                {t('videos.manager.actions.deletePackage')}
              </button>
            ) : null}
          </div>
        </section>

        {!isCreateMode ? (
          <section className={styles.card}>
            <div className={styles.cardHeader}>
              <div>
                <h2>{t('videos.manager.sections.savedVideos')}</h2>
                <p>{t('videos.manager.savedVideosHint')}</p>
              </div>
            </div>

            {existingItems.length === 0 ? (
              <div className={styles.emptyState}>
                <p className={styles.emptyTitle}>{t('videos.manager.empty.title')}</p>
                <p className={styles.emptyText}>{t('videos.manager.empty.text')}</p>
              </div>
            ) : (
              <div className={styles.videoList}>
                {existingItems.map((item) =>
                  renderCollectionItemCard(item, {
                    scope: 'existing',
                    onUpdate: (updater) => updateExistingItem(item.clientId, updater),
                    onPrimaryAction: () => void handleSaveExistingItem(item),
                    primaryLabel: t('videos.manager.actions.saveItem'),
                    onRemove: () => void handleDeleteItem(item),
                    removeLabel: t('videos.manager.actions.deleteItem'),
                  }),
                )}
              </div>
            )}
          </section>
        ) : null}

        <section className={styles.card}>
          <div className={styles.cardHeader}>
            <div>
              <h2>{t('videos.manager.sections.newVideos')}</h2>
              <p>{t('videos.manager.newVideosHint')}</p>
            </div>

            <button
              type="button"
              className={styles.secondaryButton}
              disabled={busy}
              onClick={() =>
                setPendingItems((current) => [...current, createCollectionVideoDraft()])
              }
            >
              <AddIcon size={14} />
              {t('videos.manager.actions.addVideo')}
            </button>
          </div>

          {pendingItems.length === 0 ? (
            <div className={styles.emptyStateCompact}>
              <p className={styles.emptyText}>{t('videos.manager.pendingEmpty')}</p>
            </div>
          ) : (
            <div className={styles.videoList}>
              {pendingItems.map((item) =>
                renderCollectionItemCard(item, {
                  scope: 'pending',
                  onUpdate: (updater) => updatePendingItem(item.clientId, updater),
                  onRemove: () =>
                    setPendingItems((current) =>
                      current.filter((entry) => entry.clientId !== item.clientId),
                    ),
                  removeLabel: t('videos.manager.actions.removeVideo'),
                }),
              )}
            </div>
          )}

          {pendingItems.length > 0 ? (
            <div className={styles.actionsRow}>
              <button
                type="button"
                className={styles.primaryButton}
                disabled={busy}
                onClick={() => void (isCreateMode ? handleCreatePackage() : handleAddItems())}
              >
                {isCreateMode
                  ? t('videos.manager.actions.createCollection')
                  : t('videos.manager.actions.addItems')}
              </button>
            </div>
          ) : null}
        </section>
      </>
    )
  }

  const breadcrumbItems = [
    { label: t('videos.breadcrumb.library'), to: '/videos' },
    {
      label: isCreateMode
        ? t('videos.breadcrumb.create')
        : videoPackage?.title || t('videos.breadcrumb.edit'),
    },
  ]

  return (
    <CmsAppShell activeKey="videos">
      <div className={styles.page}>
        <Breadcrumb items={breadcrumbItems} />

        {loading ? (
          <div className={styles.loaderWrap}>
            <Loader label={t('videos.manager.loading')} />
          </div>
        ) : !isCreateMode && !videoPackage ? (
          <div className={styles.emptyState}>
            <p className={styles.emptyTitle}>{t('videos.manager.notFound.title')}</p>
            <p className={styles.emptyText}>
              {error || t('videos.manager.notFound.text', { id: videoId ?? '' })}
            </p>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => navigate('/videos')}
            >
              {t('videos.manager.notFound.back')}
            </button>
          </div>
        ) : (
          <>
            <header className={styles.pageHeader}>
              <div>
                <h1 className={styles.title}>
                  {isCreateMode
                    ? t('videos.manager.createTitle')
                    : t('videos.manager.editTitle')}
                </h1>
                <p className={styles.subtitle}>{t('videos.manager.subtitle')}</p>
              </div>
              <span className={styles.packageTypeBadge}>
                {t(`videos.types.${packageType}`)}
              </span>
            </header>

            {error ? <p className={styles.errorText}>{error}</p> : null}

            {isCreateMode ? (
              <section className={styles.card}>
                <div className={styles.cardHeader}>
                  <div>
                    <h2>{t('videos.manager.sections.packageType')}</h2>
                    <p>{t('videos.manager.packageTypeHint')}</p>
                  </div>
                </div>

                <div className={styles.segmentedControl}>
                  <button
                    type="button"
                    className={[
                      styles.segmentedButton,
                      packageType === 'single' ? styles.segmentedButtonActive : '',
                    ]
                      .filter(Boolean)
                      .join(' ')}
                    disabled={busy}
                    onClick={() => handlePackageTypeChange('single')}
                  >
                    {t('videos.types.single')}
                  </button>
                  <button
                    type="button"
                    className={[
                      styles.segmentedButton,
                      packageType === 'collection' ? styles.segmentedButtonActive : '',
                    ]
                      .filter(Boolean)
                      .join(' ')}
                    disabled={busy}
                    onClick={() => handlePackageTypeChange('collection')}
                  >
                    {t('videos.types.collection')}
                  </button>
                </div>
              </section>
            ) : null}

            {packageType === 'single' ? renderSingleVideoSection() : renderCollectionSection()}
          </>
        )}
      </div>

      <ConfirmDialog
        open={deleteDialogOpen}
        title={t('videos.manager.delete.title')}
        body={t('videos.manager.delete.description', { title: packageTitle.trim() })}
        confirmLabel={t('videos.manager.delete.confirm')}
        destructive
        busy={busy}
        onConfirm={() => void handleDeletePackage()}
        onClose={() => setDeleteDialogOpen(false)}
      />
    </CmsAppShell>
  )
}
