import { useEffect, useMemo, useRef, useState } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import { resourcesApi } from '../api/resourcesApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { UploadDropzone } from '../components/media/UploadDropzone'
import { useResources } from '../hooks/useResources'
import {
  resourceCategoryOptions,
  type ResourceCategory,
  type ResourceEntry,
  type ResourceFormErrors,
  type ResourceFormState,
} from '../lib/resourceTypes'
import styles from '../styles/ResourceEditorPage.module.css'

const RESOURCE_FILE_ACCEPT = [
  'application/pdf',
  'application/msword',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.ms-excel',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-powerpoint',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'image/*',
  '.pdf',
  '.doc',
  '.docx',
  '.xls',
  '.xlsx',
  '.ppt',
  '.pptx',
  '.svg',
  '.png',
  '.jpg',
  '.jpeg',
  '.webp',
].join(',')

type ResourceEditorPageProps = {
  mode?: 'create' | 'edit'
}

function emptyFormState(): ResourceFormState {
  return {
    name: '',
    category: '',
    visibility: 'public',
  }
}

function resourceToFormState(resource: ResourceEntry): ResourceFormState {
  return {
    name: resource.name,
    category: resource.category,
    visibility: resource.visibility,
  }
}

export function ResourceEditorPage({ mode = 'create' }: ResourceEditorPageProps) {
  const { t, i18n } = useTranslation()
  const navigate = useNavigate()
  const { id } = useParams()
  const { create, update, remove } = useResources()

  const isEditMode = mode === 'edit' && Boolean(id)
  const [currentResource, setCurrentResource] = useState<ResourceEntry | null>(null)
  const [isLoading, setIsLoading] = useState(isEditMode)
  const [form, setForm] = useState<ResourceFormState>(emptyFormState())
  const [selectedFile, setSelectedFile] = useState<File | null>(null)
  const [documentPreviewUrl, setDocumentPreviewUrl] = useState<string | null>(null)
  const previewUrlRef = useRef<string | null>(null)
  const [errors, setErrors] = useState<ResourceFormErrors>({})
  const [isSubmitting, setIsSubmitting] = useState(false)
  const [isDeleting, setIsDeleting] = useState(false)
  const [activeDocumentAction, setActiveDocumentAction] = useState<'preview' | 'download' | null>(null)

  const dateFormatter = useMemo(
    () =>
      new Intl.DateTimeFormat(i18n.language, {
        month: 'short',
        day: '2-digit',
        year: 'numeric',
      }),
    [i18n.language],
  )

  useEffect(() => {
    if (!isEditMode || !id) {
      return
    }

    const loadResource = async () => {
      try {
        setIsLoading(true)
        const resource = await resourcesApi.getResource(id)
        setCurrentResource(resource)
      } catch (loadError) {
        toast.error(
          loadError instanceof Error
            ? loadError.message
            : t('resources.editor.loadError'),
        )
      } finally {
        setIsLoading(false)
      }
    }

    void loadResource()
  }, [id, isEditMode, t])

  useEffect(() => {
    if (isEditMode && !currentResource) {
      return
    }

    setForm(currentResource ? resourceToFormState(currentResource) : emptyFormState())
    setSelectedFile(null)
  }, [currentResource, isEditMode])

  useEffect(() => {
    let cancelled = false
    let nextPreviewUrl: string | null = null

    async function loadPreview() {
      if (selectedFile) {
        nextPreviewUrl = URL.createObjectURL(selectedFile)
      } else if (currentResource && canPreviewDocument(currentResource.mimeType, currentResource.fileName)) {
        try {
          const blob = await resourcesApi.getResourceContent(currentResource.id)
          nextPreviewUrl = URL.createObjectURL(blob)
        } catch {
          nextPreviewUrl = null
        }
      }

      if (cancelled) {
        if (nextPreviewUrl) {
          URL.revokeObjectURL(nextPreviewUrl)
        }
        return
      }

      if (previewUrlRef.current) {
        URL.revokeObjectURL(previewUrlRef.current)
      }
      previewUrlRef.current = nextPreviewUrl
      setDocumentPreviewUrl(nextPreviewUrl)
    }

    void loadPreview()

    return () => {
      cancelled = true
    }
  }, [currentResource, selectedFile])

  useEffect(() => {
    return () => {
      if (previewUrlRef.current) {
        URL.revokeObjectURL(previewUrlRef.current)
      }
      previewUrlRef.current = null
    }
  }, [])

  function updateField<K extends keyof ResourceFormState>(key: K, value: ResourceFormState[K]) {
    setForm((current) => ({ ...current, [key]: value }))
    if (errors[key]) {
      setErrors((current) => {
        const next = { ...current }
        delete next[key]
        return next
      })
    }
  }

  function validate() {
    const nextErrors: ResourceFormErrors = {}
    if (!form.name.trim()) {
      nextErrors.name = t('resources.editor.validation.nameRequired')
    }
    if (!form.category) {
      nextErrors.category = t('resources.editor.validation.categoryRequired')
    }
    if (!isEditMode && !selectedFile) {
      nextErrors.file = t('resources.editor.validation.fileRequired')
    }
    return nextErrors
  }

  async function handleSave() {
    const validationErrors = validate()
    if (Object.keys(validationErrors).length > 0) {
      setErrors(validationErrors)
      toast.error(t('resources.editor.validation.summary'))
      return
    }

    setErrors({})
    setIsSubmitting(true)
    try {
      const input = {
        name: form.name.trim(),
        category: form.category as ResourceCategory,
        visibility: form.visibility,
      }

      if (isEditMode && currentResource) {
        const updated = await update(currentResource.id, input, selectedFile ?? undefined)
        setCurrentResource(updated)
        setSelectedFile(null)
        toast.success(t('resources.feedback.saved'))
      } else if (selectedFile) {
        const created = await create(input, selectedFile)
        toast.success(t('resources.feedback.created'))
        navigate(`/resources/${created.id}/edit`, { replace: true })
      }
    } catch (saveError) {
      toast.error(
        saveError instanceof Error
          ? saveError.message
          : t('resources.feedback.saveError'),
      )
    } finally {
      setIsSubmitting(false)
    }
  }

  async function handleDelete() {
    if (!currentResource) {
      return
    }
    setIsDeleting(true)
    try {
      await remove(currentResource.id)
      toast.success(t('resources.feedback.deleted'))
      navigate('/resources', { replace: true })
    } catch (deleteError) {
      toast.error(
        deleteError instanceof Error
          ? deleteError.message
          : t('resources.feedback.deleteError'),
      )
    } finally {
      setIsDeleting(false)
    }
  }

  function handleFileSelect(files: File[]) {
    const [file] = files
    if (!file) {
      return
    }
    setSelectedFile(file)
    if (errors.file) {
      setErrors((current) => {
        const next = { ...current }
        delete next.file
        return next
      })
    }
  }

  function handleRemoveSelectedFile() {
    setSelectedFile(null)
  }

  async function handlePreviewDocument() {
    const fileName = selectedFile?.name || currentResource?.fileName
    const mimeType = selectedFile?.type || currentResource?.mimeType || ''

    if (!fileName || !canPreviewDocument(mimeType, fileName)) {
      return
    }

    setActiveDocumentAction('preview')
    try {
      const previewUrl = await createTemporaryResourceObjectUrl(currentResource?.id, selectedFile)
      window.open(previewUrl, '_blank', 'noopener,noreferrer')
      if (!selectedFile) {
        scheduleObjectUrlRevoke(previewUrl)
      }
    } catch {
      toast.error(t('resources.feedback.previewError'))
    } finally {
      setActiveDocumentAction(null)
    }
  }

  async function handleDownloadDocument() {
    setActiveDocumentAction('download')
    try {
      const downloadUrl = await createTemporaryResourceObjectUrl(currentResource?.id, selectedFile)
      const fileName = selectedFile?.name || currentResource?.fileName || 'resource-file'
      triggerFileDownload(downloadUrl, fileName)
      if (!selectedFile) {
        scheduleObjectUrlRevoke(downloadUrl)
      }
    } catch {
      toast.error(t('resources.feedback.downloadError'))
    } finally {
      setActiveDocumentAction(null)
    }
  }

  if (isLoading) {
    return (
      <CmsAppShell activeKey="resources">
        <div className={styles.loaderWrap}>
          <Loader label={t('resources.common.loading')} />
        </div>
      </CmsAppShell>
    )
  }

  const documentName = selectedFile?.name || currentResource?.fileName || ''
  const documentMimeType = selectedFile?.type || currentResource?.mimeType || ''
  const documentSize = selectedFile?.size ?? currentResource?.fileSize ?? 0
  const canPreview = canPreviewDocument(documentMimeType, documentName)
  const errorMessages = Array.from(new Set(Object.values(errors).filter(Boolean)))

  return (
    <CmsAppShell activeKey="resources">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: t('resources.breadcrumb.library'), to: '/resources' },
            {
              label: isEditMode
                ? t('resources.breadcrumb.edit')
                : t('resources.breadcrumb.create'),
            },
          ]}
        />

        <div className={styles.header}>
          <div className={styles.headerText}>
            <h1>
              {isEditMode
                ? t('resources.editor.titleEdit')
                : t('resources.editor.titleCreate')}
            </h1>
            <p>{t('resources.editor.subtitle')}</p>
          </div>
          <button
            type="button"
            className={styles.backLink}
            onClick={() => navigate('/resources')}
          >
            {t('resources.editor.backToList')}
          </button>
        </div>

        {errorMessages.length > 0 && (
          <div className={styles.errorSummary} role="alert">
            <p>{t('resources.editor.validation.summary')}</p>
            <ul>
              {errorMessages.map((message) => (
                <li key={message}>{message}</li>
              ))}
            </ul>
          </div>
        )}

        <div className={styles.layout}>
          <div className={styles.mainColumn}>
            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <h2>{t('resources.editor.sections.details')}</h2>
              </div>

              <div className={styles.fieldGrid}>
                <label className={`${styles.field} ${styles.fieldFull}`}>
                  <span>{t('resources.editor.fields.name')}</span>
                  <input
                    type="text"
                    value={form.name}
                    placeholder={t('resources.editor.fields.namePlaceholder')}
                    onChange={(event) => updateField('name', event.target.value)}
                  />
                  {errors.name && <p className={styles.fieldError}>{errors.name}</p>}
                </label>

                <label className={styles.field}>
                  <span>{t('resources.editor.fields.category')}</span>
                  <select
                    value={form.category}
                    onChange={(event) =>
                      updateField('category', event.target.value as ResourceFormState['category'])
                    }
                  >
                    <option value="">
                      {t('resources.editor.fields.categoryPlaceholder')}
                    </option>
                    {resourceCategoryOptions.map((category) => (
                      <option key={category.id} value={category.id}>
                        {category.label}
                      </option>
                    ))}
                  </select>
                  {errors.category && (
                    <p className={styles.fieldError}>{errors.category}</p>
                  )}
                </label>

                <fieldset className={`${styles.field} ${styles.visibilityField}`}>
                  <legend>{t('resources.editor.fields.visibility')}</legend>
                  <label className={styles.radioOption}>
                    <input
                      type="radio"
                      name="resource-visibility"
                      checked={form.visibility === 'public'}
                      onChange={() => updateField('visibility', 'public')}
                    />
                    <span>{t('resources.editor.fields.visibilityPublic')}</span>
                  </label>
                  <label className={styles.radioOption}>
                    <input
                      type="radio"
                      name="resource-visibility"
                      checked={form.visibility === 'internal'}
                      onChange={() => updateField('visibility', 'internal')}
                    />
                    <span>{t('resources.editor.fields.visibilityInternal')}</span>
                  </label>
                </fieldset>
              </div>
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <h2>{t('resources.editor.sections.document')}</h2>
                <span>{t('resources.editor.sections.documentLimit')}</span>
              </div>

              <div className={styles.mediaSection}>
                <UploadDropzone
                  accept={RESOURCE_FILE_ACCEPT}
                  icon={<UploadCloudIcon />}
                  label={t('resources.editor.document.dropLabel')}
                  hint={t('resources.editor.document.dropHint')}
                  onFiles={handleFileSelect}
                />
                {errors.file && <p className={styles.fieldError}>{errors.file}</p>}

                {(selectedFile || currentResource) && (
                  <article className={styles.documentCard}>
                    <div className={styles.documentPreviewCard} aria-hidden="true">
                      {canPreview && documentPreviewUrl ? (
                        <iframe
                          src={documentPreviewUrl}
                          title={documentName}
                          className={styles.documentPreviewFrame}
                        />
                      ) : (
                        <>
                          <DocumentIcon />
                          <span className={styles.documentPreviewBadge}>
                            {resolveDocumentTypeLabel(documentMimeType, documentName)}
                          </span>
                        </>
                      )}
                    </div>

                    <div className={styles.documentContent}>
                      <h3>{form.name || documentName || t('resources.editor.document.untitled')}</h3>
                      <p className={styles.documentMeta}>
                        {buildDocumentMeta({
                          isPendingUpload: Boolean(selectedFile),
                          fileTypeLabel: resolveDocumentTypeLabel(documentMimeType, documentName),
                          fileSize: documentSize,
                          t,
                        })}
                      </p>

                      <div className={styles.documentActions}>
                        {canPreview ? (
                          <button
                            type="button"
                            className={styles.actionButton}
                            disabled={activeDocumentAction !== null}
                            onClick={() => void handlePreviewDocument()}
                          >
                            {activeDocumentAction === 'preview'
                              ? t('resources.common.loading')
                              : t('resources.editor.document.preview')}
                          </button>
                        ) : null}
                        <button
                          type="button"
                          className={styles.actionButton}
                          disabled={activeDocumentAction !== null}
                          onClick={() => void handleDownloadDocument()}
                        >
                          {activeDocumentAction === 'download'
                            ? t('resources.common.loading')
                            : t('resources.editor.document.download')}
                        </button>
                        {selectedFile ? (
                          <button
                            type="button"
                            className={styles.dangerButton}
                            onClick={handleRemoveSelectedFile}
                          >
                            {t('resources.editor.document.removeSelected')}
                          </button>
                        ) : null}
                      </div>
                    </div>
                  </article>
                )}
              </div>
            </section>

            <div className={styles.actionBar}>
              <button
                type="button"
                className={styles.primaryButton}
                disabled={isSubmitting}
                onClick={() => void handleSave()}
              >
                {isSubmitting
                  ? t('resources.common.loading')
                  : isEditMode
                    ? t('resources.editor.actions.save')
                    : t('resources.editor.actions.create')}
              </button>
              {isEditMode && (
                <button
                  type="button"
                  className={styles.deleteButton}
                  disabled={isDeleting || isSubmitting}
                  onClick={() => void handleDelete()}
                >
                  {isDeleting
                    ? t('resources.common.loading')
                    : t('resources.editor.actions.delete')}
                </button>
              )}
            </div>
          </div>

          <aside className={styles.sideColumn}>
            <div className={styles.metaCard}>
              <span className={styles.metaLabel}>{t('resources.editor.meta.visibility')}</span>
              <strong>
                {t(
                  form.visibility === 'public'
                    ? 'resources.editor.fields.visibilityPublic'
                    : 'resources.editor.fields.visibilityInternal',
                )}
              </strong>
            </div>

            {currentResource && (
              <div className={styles.metaCard}>
                <span className={styles.metaLabel}>{t('resources.editor.meta.created')}</span>
                <strong>{dateFormatter.format(new Date(currentResource.createdAt))}</strong>
                <span className={styles.metaLabel}>{t('resources.editor.meta.updated')}</span>
                <strong>{dateFormatter.format(new Date(currentResource.updatedAt))}</strong>
              </div>
            )}

            <div className={styles.helpCard}>
              <h3>{t('resources.editor.help.title')}</h3>
              <p>{t('resources.editor.help.body')}</p>
            </div>
          </aside>
        </div>
      </div>
    </CmsAppShell>
  )
}

async function createTemporaryResourceObjectUrl(resourceId?: string, file?: File | null) {
  if (file) {
    return URL.createObjectURL(file)
  }
  if (!resourceId) {
    throw new Error('Document URL unavailable')
  }
  const blob = await resourcesApi.getResourceContent(resourceId)
  return URL.createObjectURL(blob)
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
    isPendingUpload ? t('resources.editor.document.pendingUpload') : t('resources.editor.document.savedUpload'),
    fileTypeLabel,
    typeof fileSize === 'number' && fileSize > 0 ? formatFileSize(fileSize) : '',
  ].filter(Boolean)

  return parts.join(' | ')
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

function canPreviewDocument(mimeType: string, fileName: string) {
  const normalizedMimeType = mimeType.trim().toLowerCase()
  const extension = fileName.split('.').pop()?.trim().toLowerCase() ?? ''

  return (
    normalizedMimeType.includes('pdf') ||
    normalizedMimeType.startsWith('image/') ||
    normalizedMimeType.includes('svg') ||
    ['pdf', 'png', 'jpg', 'jpeg', 'webp', 'gif', 'svg'].includes(extension)
  )
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
  if (normalizedMimeType.includes('word') || normalizedMimeType.includes('document')) {
    return 'DOC'
  }
  if (normalizedMimeType.includes('powerpoint') || normalizedMimeType.includes('presentation')) {
    return 'PPT'
  }
  if (normalizedMimeType.includes('excel') || normalizedMimeType.includes('sheet')) {
    return 'XLS'
  }
  if (normalizedMimeType.includes('svg')) {
    return 'SVG'
  }

  return 'FILE'
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

function UploadCloudIcon() {
  return (
    <svg viewBox="0 0 24 24" width="22" height="22" fill="none" aria-hidden="true">
      <path d="M8.5 18H7a4 4 0 1 1 .7-7.94A5.5 5.5 0 0 1 18.38 12H19a3 3 0 1 1 0 6h-2.5" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
      <path d="M12 10v8m0-8 3 3m-3-3-3 3" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  )
}

function DocumentIcon() {
  return (
    <svg viewBox="0 0 24 24" width="28" height="28" fill="none" aria-hidden="true">
      <path d="M8 3.5h6l4 4V20.5H8z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" />
      <path d="M14 3.5v4h4" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" />
    </svg>
  )
}
