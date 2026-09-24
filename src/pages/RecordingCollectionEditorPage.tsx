import { type Dispatch, type SetStateAction, useCallback, useEffect, useState } from 'react'
import toast from 'react-hot-toast'
import { Trans, useTranslation } from 'react-i18next'
import { useNavigate, useParams } from 'react-router-dom'
import { getApiErrorMessage } from '../api/apiError'
import {
  type RecordingItem,
  type RecordingItemInput,
  recordingsApi,
} from '../api/recordingsApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { ConfirmDialog } from '../components/cms/ConfirmDialog'
import { AddIcon, CloudUploadIcon, DeleteIcon } from '../components/icons'
import { Loader } from '../components/Loader'
import { UploadDropzone } from '../components/media/UploadDropzone'
import {
  getRecordingUploadValidationErrorMessage,
  RECORDING_FILE_ACCEPT,
  RECORDING_UPLOAD_MAX_FILE_SIZE_MB,
  RECORDING_UPLOAD_SUPPORTED_FORMATS_LABEL,
  validateRecordingUploadFile,
} from '../lib/recordingUpload'
import styles from '../styles/RecordingCollectionEditorPage.module.css'

type RecordingItemDraft = {
  clientId: string
  id?: number
  title: string
  description: string
  recordingFile: File | null
  recordingUrl?: string
}

function createDraftId(prefix: string) {
  if (typeof globalThis.crypto?.randomUUID === 'function') {
    return `${prefix}-${globalThis.crypto.randomUUID()}`
  }
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}`
}

function createEmptyItemDraft(): RecordingItemDraft {
  return {
    clientId: createDraftId('recording-item'),
    title: '',
    description: '',
    recordingFile: null,
  }
}

function draftFromItem(item: RecordingItem): RecordingItemDraft {
  return {
    clientId: createDraftId('recording-item'),
    id: item.id,
    title: item.title,
    description: item.description,
    recordingFile: null,
    recordingUrl: item.recordingUrl,
  }
}

function toItemInput(draft: RecordingItemDraft): RecordingItemInput {
  const description = draft.description.trim()
  return {
    title: draft.title.trim(),
    description: description || undefined,
  }
}

export function RecordingCollectionEditorPage() {
  const { t } = useTranslation()
  const navigate = useNavigate()
  const { collectionId } = useParams()
  const isCreateMode = !collectionId
  const numericId = collectionId ? Number.parseInt(collectionId, 10) : Number.NaN

  const [loading, setLoading] = useState(!isCreateMode)
  const [busy, setBusy] = useState(false)
  const [itemBusyId, setItemBusyId] = useState<number | null>(null)
  const [pendingItemBusyId, setPendingItemBusyId] = useState<string | null>(null)
  const [notFound, setNotFound] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [collectionName, setCollectionName] = useState('')
  const [existingItems, setExistingItems] = useState<RecordingItemDraft[]>([])
  const [pendingItems, setPendingItems] = useState<RecordingItemDraft[]>([])
  const [deleteCollectionOpen, setDeleteCollectionOpen] = useState(false)
  const [deleteItemTarget, setDeleteItemTarget] = useState<RecordingItemDraft | null>(null)

  const loadCollection = useCallback(async () => {
    if (!Number.isFinite(numericId)) {
      setNotFound(true)
      setLoading(false)
      return
    }

    setLoading(true)
    setError(null)
    setNotFound(false)

    try {
      const detail = await recordingsApi.getRecordingCollection(numericId)
      setCollectionName(detail.name)
      setExistingItems(detail.items.map(draftFromItem))
    } catch (loadError) {
      setNotFound(true)
      setError(getApiErrorMessage(loadError))
    } finally {
      setLoading(false)
    }
  }, [numericId])

  useEffect(() => {
    if (isCreateMode) {
      setCollectionName('')
      setExistingItems([])
      setPendingItems([])
      setNotFound(false)
      setError(null)
      setLoading(false)
      return
    }

    void loadCollection()
  }, [isCreateMode, loadCollection])

  function updateItemDraft(
    setter: Dispatch<SetStateAction<RecordingItemDraft[]>>,
    clientId: string,
    patch: Partial<RecordingItemDraft>,
  ) {
    setter((previous) =>
      previous.map((item) => (item.clientId === clientId ? { ...item, ...patch } : item)),
    )
  }

  function collectPendingItems() {
    return pendingItems.filter(
      (item) => item.title.trim() || item.description.trim() || item.recordingFile,
    )
  }

  function allPendingItemsHaveTitles(items: RecordingItemDraft[]) {
    return items.every((item) => item.title.trim().length > 0)
  }

  function allPendingItemsHaveRecordings(items: RecordingItemDraft[]) {
    return items.every((item) => Boolean(item.recordingFile))
  }

  function handleRecordingFile(
    setter: Dispatch<SetStateAction<RecordingItemDraft[]>>,
    clientId: string,
    file?: File,
  ) {
    if (!file) {
      return
    }

    const validationError = validateRecordingUploadFile(file)
    if (validationError) {
      setError(getRecordingUploadValidationErrorMessage(validationError))
      return
    }

    setError(null)
    updateItemDraft(setter, clientId, { recordingFile: file })
  }

  async function handleCreate() {
    const trimmedName = collectionName.trim()
    if (!trimmedName) {
      setError(t('recordings.validation.collectionNameRequired'))
      return
    }

    const items = collectPendingItems()
    if (!allPendingItemsHaveTitles(items)) {
      setError(t('recordings.validation.itemTitleRequired'))
      return
    }
    if (!allPendingItemsHaveRecordings(items)) {
      setError(t('recordings.validation.recordingRequired'))
      return
    }

    setBusy(true)
    setError(null)
    try {
      const result = await recordingsApi.createRecordingCollection({
        name: trimmedName,
        items: items.map(toItemInput),
        recordingFiles: items.map((item) => item.recordingFile),
      })
      toast.success(result.message)
      navigate(`/recordings/${result.recording.id}`)
    } catch (createError) {
      setError(getApiErrorMessage(createError))
    } finally {
      setBusy(false)
    }
  }

  async function handleSaveName() {
    const trimmedName = collectionName.trim()
    if (!trimmedName) {
      setError(t('recordings.validation.collectionNameRequired'))
      return
    }

    setBusy(true)
    setError(null)
    try {
      const result = await recordingsApi.updateRecordingCollection(numericId, { name: trimmedName })
      toast.success(result.message)
      setCollectionName(trimmedName)
    } catch (saveError) {
      setError(getApiErrorMessage(saveError))
    } finally {
      setBusy(false)
    }
  }

  async function handleAddItems() {
    const items = collectPendingItems()
    if (items.length === 0) {
      return
    }
    if (!allPendingItemsHaveTitles(items)) {
      setError(t('recordings.validation.itemTitleRequired'))
      return
    }
    if (!allPendingItemsHaveRecordings(items)) {
      setError(t('recordings.validation.recordingRequired'))
      return
    }

    setBusy(true)
    setError(null)
    try {
      const result = await recordingsApi.addRecordingItems(numericId, {
        items: items.map(toItemInput),
        recordingFiles: items.map((item) => item.recordingFile),
      })
      toast.success(result.message)
      setPendingItems([])
      await loadCollection()
    } catch (addError) {
      setError(getApiErrorMessage(addError))
    } finally {
      setBusy(false)
    }
  }

  async function handleSavePendingItem(item: RecordingItemDraft) {
    if (!Number.isFinite(numericId)) {
      return
    }
    if (!item.title.trim()) {
      setError(t('recordings.validation.itemTitleRequired'))
      return
    }
    if (!item.recordingFile) {
      setError(t('recordings.validation.recordingRequired'))
      return
    }

    setPendingItemBusyId(item.clientId)
    setError(null)
    try {
      const result = await recordingsApi.addRecordingItems(numericId, {
        items: [toItemInput(item)],
        recordingFiles: [item.recordingFile],
      })
      setPendingItems((previous) =>
        previous.filter((candidate) => candidate.clientId !== item.clientId),
      )
      toast.success(result.message)

      const detail = await recordingsApi.getRecordingCollection(numericId)
      setExistingItems(detail.items.map(draftFromItem))
    } catch (saveError) {
      setError(getApiErrorMessage(saveError))
    } finally {
      setPendingItemBusyId(null)
    }
  }

  async function handleSaveExistingItem(item: RecordingItemDraft) {
    if (!item.id) {
      return
    }
    if (!item.title.trim()) {
      setError(t('recordings.validation.itemTitleRequired'))
      return
    }
    if (!item.recordingUrl && !item.recordingFile) {
      setError(t('recordings.validation.recordingRequired'))
      return
    }

    setItemBusyId(item.id)
    setError(null)
    try {
      const result = await recordingsApi.updateRecordingItem(numericId, item.id, {
        ...toItemInput(item),
        recordingFile: item.recordingFile,
      })
      toast.success(result.message)
      setExistingItems((previous) =>
        previous.map((candidate) =>
          candidate.clientId === item.clientId
            ? {
                ...candidate,
                title: result.item.title,
                description: result.item.description,
                recordingFile: null,
                recordingUrl: result.item.recordingUrl,
              }
            : candidate,
        ),
      )
    } catch (saveError) {
      setError(getApiErrorMessage(saveError))
    } finally {
      setItemBusyId(null)
    }
  }

  async function handleDeleteItemConfirm() {
    if (!deleteItemTarget?.id) {
      return
    }

    setBusy(true)
    setError(null)
    try {
      const result = await recordingsApi.deleteRecordingItem(numericId, deleteItemTarget.id)
      toast.success(result.message)
      setExistingItems((previous) =>
        previous.filter((item) => item.clientId !== deleteItemTarget.clientId),
      )
      setDeleteItemTarget(null)
    } catch (deleteError) {
      setError(getApiErrorMessage(deleteError))
    } finally {
      setBusy(false)
    }
  }

  async function handleDeleteCollectionConfirm() {
    setBusy(true)
    setError(null)
    try {
      const result = await recordingsApi.deleteRecordingCollection(numericId)
      toast.success(result.message)
      setDeleteCollectionOpen(false)
      navigate('/recordings')
    } catch (deleteError) {
      setError(getApiErrorMessage(deleteError))
    } finally {
      setBusy(false)
    }
  }

  if (loading) {
    return (
      <CmsAppShell activeKey="recordings">
        <div className={styles.loaderWrap}>
          <Loader label={t('recordings.manager.loading')} />
        </div>
      </CmsAppShell>
    )
  }

  if (notFound) {
    return (
      <CmsAppShell activeKey="recordings">
        <div className={styles.page}>
          <Breadcrumb items={[{ label: t('recordings.breadcrumb.library'), to: '/recordings' }]} />
          <div className={styles.emptyState}>
            <p className={styles.emptyTitle}>{t('recordings.manager.notFound.title')}</p>
            <p className={styles.emptyText}>{t('recordings.manager.notFound.text')}</p>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => navigate('/recordings')}
            >
              {t('recordings.manager.notFound.back')}
            </button>
          </div>
        </div>
      </CmsAppShell>
    )
  }

  return (
    <CmsAppShell activeKey="recordings">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: t('recordings.breadcrumb.library'), to: '/recordings' },
            {
              label: isCreateMode
                ? t('recordings.breadcrumb.create')
                : t('recordings.breadcrumb.edit'),
            },
          ]}
        />

        <header className={styles.pageHeader}>
          <div>
            <h1 className={styles.title}>
              {isCreateMode
                ? t('recordings.manager.createTitle')
                : t('recordings.manager.editTitle')}
            </h1>
            <p className={styles.subtitle}>{t('recordings.manager.subtitle')}</p>
          </div>
        </header>

        {error ? <p className={styles.errorText}>{error}</p> : null}

        <section className={styles.card}>
          <h2 className={styles.sectionTitle}>
            {t('recordings.manager.sections.collectionDetails')}
          </h2>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>
              {t('recordings.manager.fields.collectionName')}
            </span>
            <input
              type="text"
              className={styles.input}
              value={collectionName}
              onChange={(event) => setCollectionName(event.target.value)}
              placeholder={t('recordings.manager.placeholders.collectionName')}
            />
          </label>
          {!isCreateMode ? (
            <div className={styles.sectionActions}>
              <button
                type="button"
                className={styles.primaryButton}
                onClick={() => void handleSaveName()}
                disabled={busy}
              >
                {t('recordings.manager.actions.saveCollectionName')}
              </button>
            </div>
          ) : null}
        </section>

        {!isCreateMode ? (
          <section className={styles.card}>
            <h2 className={styles.sectionTitle}>{t('recordings.manager.sections.savedItems')}</h2>
            <p className={styles.sectionHint}>{t('recordings.manager.savedItemsHint')}</p>

            {existingItems.length === 0 ? (
              <p className={styles.emptyText}>{t('recordings.manager.empty.text')}</p>
            ) : (
              <div className={styles.itemList}>
                {existingItems.map((item) => (
                  <div key={item.clientId} className={styles.itemCard}>
                    <label className={styles.field}>
                      <span className={styles.fieldLabel}>
                        {t('recordings.manager.fields.itemTitle')}
                      </span>
                      <input
                        type="text"
                        className={styles.input}
                        value={item.title}
                        onChange={(event) =>
                          updateItemDraft(setExistingItems, item.clientId, {
                            title: event.target.value,
                          })
                        }
                        placeholder={t('recordings.manager.placeholders.itemTitle')}
                      />
                    </label>
                    <label className={styles.field}>
                      <span className={styles.fieldLabel}>
                        {t('recordings.manager.fields.description')}
                      </span>
                      <textarea
                        className={styles.textarea}
                        value={item.description}
                        onChange={(event) =>
                          updateItemDraft(setExistingItems, item.clientId, {
                            description: event.target.value,
                          })
                        }
                        placeholder={t('recordings.manager.placeholders.description')}
                        rows={3}
                      />
                    </label>
                    <div className={styles.recordingField}>
                      <span className={styles.fieldLabel}>
                        {t('recordings.manager.fields.recording')}
                      </span>
                      {item.recordingUrl ? (
                        <audio
                          className={styles.audioPreview}
                          controls
                          preload="none"
                          src={item.recordingUrl}
                          aria-label={t('recordings.manager.currentRecording', {
                            title: item.title,
                          })}
                        />
                      ) : null}
                      <UploadDropzone
                        accept={RECORDING_FILE_ACCEPT}
                        variant="compact"
                        icon={<CloudUploadIcon size={18} />}
                        label={t(
                          item.recordingUrl
                            ? 'recordings.manager.replaceRecording'
                            : 'recordings.manager.uploadRecording',
                        )}
                        hint={t('recordings.manager.recordingHint', {
                          formats: RECORDING_UPLOAD_SUPPORTED_FORMATS_LABEL,
                          maxSize: RECORDING_UPLOAD_MAX_FILE_SIZE_MB,
                        })}
                        disabled={busy}
                        onFiles={(files) =>
                          handleRecordingFile(setExistingItems, item.clientId, files[0])
                        }
                      />
                      {item.recordingFile ? (
                        <p className={styles.selectedFile}>
                          {t('recordings.manager.selectedRecording', {
                            name: item.recordingFile.name,
                          })}
                        </p>
                      ) : null}
                    </div>
                    <div className={styles.itemActions}>
                      <button
                        type="button"
                        className={styles.primaryButton}
                        onClick={() => void handleSaveExistingItem(item)}
                        disabled={busy || itemBusyId === item.id}
                      >
                        {t('recordings.manager.actions.saveItem')}
                      </button>
                      <button
                        type="button"
                        className={styles.dangerButton}
                        onClick={() => setDeleteItemTarget(item)}
                        disabled={busy}
                      >
                        <DeleteIcon size={14} />
                        {t('recordings.manager.actions.deleteItem')}
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </section>
        ) : null}

        <section className={styles.card}>
          <h2 className={styles.sectionTitle}>{t('recordings.manager.sections.newItems')}</h2>
          <p className={styles.sectionHint}>
            {t(
              isCreateMode
                ? 'recordings.manager.newItemsCreateHint'
                : 'recordings.manager.newItemsHint',
            )}
          </p>

          <div className={styles.itemList}>
            {pendingItems.map((item) => (
              <div key={item.clientId} className={styles.itemCard}>
                <label className={styles.field}>
                  <span className={styles.fieldLabel}>
                    {t('recordings.manager.fields.itemTitle')}
                  </span>
                  <input
                    type="text"
                    className={styles.input}
                    value={item.title}
                    onChange={(event) =>
                      updateItemDraft(setPendingItems, item.clientId, {
                        title: event.target.value,
                      })
                    }
                    placeholder={t('recordings.manager.placeholders.itemTitle')}
                  />
                </label>
                <label className={styles.field}>
                  <span className={styles.fieldLabel}>
                    {t('recordings.manager.fields.description')}
                  </span>
                  <textarea
                    className={styles.textarea}
                    value={item.description}
                    onChange={(event) =>
                      updateItemDraft(setPendingItems, item.clientId, {
                        description: event.target.value,
                      })
                    }
                    placeholder={t('recordings.manager.placeholders.description')}
                    rows={3}
                  />
                </label>
                <div className={styles.recordingField}>
                  <span className={styles.fieldLabel}>
                    {t('recordings.manager.fields.recording')}
                  </span>
                  <UploadDropzone
                    accept={RECORDING_FILE_ACCEPT}
                    variant="compact"
                    icon={<CloudUploadIcon size={18} />}
                    label={t('recordings.manager.uploadRecording')}
                    hint={t('recordings.manager.recordingHint', {
                      formats: RECORDING_UPLOAD_SUPPORTED_FORMATS_LABEL,
                      maxSize: RECORDING_UPLOAD_MAX_FILE_SIZE_MB,
                    })}
                    disabled={busy}
                    onFiles={(files) =>
                      handleRecordingFile(setPendingItems, item.clientId, files[0])
                    }
                  />
                  {item.recordingFile ? (
                    <p className={styles.selectedFile}>
                      {t('recordings.manager.selectedRecording', {
                        name: item.recordingFile.name,
                      })}
                    </p>
                  ) : null}
                </div>
                <div className={styles.itemActions}>
                  {!isCreateMode ? (
                    <button
                      type="button"
                      className={styles.primaryButton}
                      onClick={() => void handleSavePendingItem(item)}
                      disabled={busy || pendingItemBusyId !== null}
                    >
                      {pendingItemBusyId === item.clientId
                        ? t('recordings.manager.actions.savingItem')
                        : t('recordings.manager.actions.saveItem')}
                    </button>
                  ) : null}
                  <button
                    type="button"
                    className={styles.ghostButton}
                    onClick={() =>
                      setPendingItems((previous) =>
                        previous.filter((candidate) => candidate.clientId !== item.clientId),
                      )
                    }
                    disabled={busy || pendingItemBusyId !== null}
                  >
                    {t('recordings.manager.actions.removeItem')}
                  </button>
                </div>
              </div>
            ))}
          </div>

          <div className={styles.sectionActions}>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() => setPendingItems((previous) => [...previous, createEmptyItemDraft()])}
            >
              <AddIcon size={16} />
              {t('recordings.manager.actions.addItem')}
            </button>
            {isCreateMode ? null : (
              <button
                type="button"
                className={styles.primaryButton}
                onClick={() => void handleAddItems()}
                disabled={
                  busy || pendingItemBusyId !== null || collectPendingItems().length === 0
                }
              >
                {t('recordings.manager.actions.addItems')}
              </button>
            )}
          </div>
        </section>

        <div className={styles.footer}>
          {isCreateMode ? (
            <button
              type="button"
              className={styles.primaryButton}
              onClick={() => void handleCreate()}
              disabled={busy}
            >
              {t('recordings.manager.actions.createCollection')}
            </button>
          ) : (
            <button
              type="button"
              className={styles.dangerButton}
              onClick={() => setDeleteCollectionOpen(true)}
              disabled={busy}
            >
              <DeleteIcon size={14} />
              {t('recordings.manager.actions.deleteCollection')}
            </button>
          )}
        </div>
      </div>

      <ConfirmDialog
        open={Boolean(deleteItemTarget)}
        title={t('recordings.manager.deleteItem.title')}
        body={t('recordings.manager.deleteItem.description')}
        confirmLabel={t('recordings.manager.deleteItem.confirm')}
        destructive
        busy={busy}
        onConfirm={() => void handleDeleteItemConfirm()}
        onClose={() => setDeleteItemTarget(null)}
      />

      <ConfirmDialog
        open={deleteCollectionOpen}
        title={t('recordings.manager.delete.title')}
        body={
          <Trans
            i18nKey="recordings.manager.delete.description"
            values={{ name: collectionName }}
            components={{ collectionName: <strong /> }}
          />
        }
        confirmLabel={t('recordings.manager.delete.confirm')}
        destructive
        busy={busy}
        onConfirm={() => void handleDeleteCollectionConfirm()}
        onClose={() => setDeleteCollectionOpen(false)}
      />
    </CmsAppShell>
  )
}
