import { useEffect, useMemo, useState } from 'react'
import type { InputHTMLAttributes } from 'react'
import { DateInput } from '../components/cms/DateInput'
import toast from 'react-hot-toast'
import { Trans, useTranslation } from 'react-i18next'
import { Link, useNavigate, useParams } from 'react-router-dom'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { UploadDropzone } from '../components/media/UploadDropzone'
import { PublishingControls } from '../components/cms/PublishingControls'
import { EntryActions } from '../components/cms/EntryActions'
import { StatusBadge } from '../components/cms/StatusBadge'
import type {
  EventFormErrors,
  EventFormState,
  EventOccurrenceForm,
} from '../lib/eventsForm'
import {
  buildEventFormStateFromDetail,
  buildLocationSummary,
  buildSaveEventRequest,
  createDefaultEventFormState,
  createOccurrenceFromMain,
  usesSeparateEndDate,
  usesTime,
  validateEventForm,
} from '../lib/eventsForm'
import {
  RESOURCE_FILE_ACCEPT,
  RESOURCE_IMAGE_FILE_ACCEPT,
  getResourceImageUploadValidationErrorMessage,
  getResourceUploadValidationErrorMessage,
  validateResourceImageUploadFile,
  validateResourceUploadFile,
} from '../lib/resourceUpload'
import {
  clearCurrentEvent,
  clearEventSaveState,
  createEvent,
  deleteEvent as deleteEventAction,
  deleteEventDocument,
  deleteEventPhoto,
  fetchEventById,
  fetchEventLookups,
  removeLocalAttachment,
  removeLocalDisplayImage,
  selectCurrentEvent,
  selectEventDeletingIds,
  selectEventDetail,
  selectEventGalleries,
  selectEventLookups,
  selectEventSave,
  selectSavedLocations,
  updateEvent,
} from '../store/eventsSlice'
import { useAppDispatch, useAppSelector } from '../store/hooks'
import type { EventMedia, EventType, RecurrenceFrequency } from '../api/eventsApi'
import { fetchEventMediaObjectUrl } from '../lib/eventMedia'
import styles from '../styles/EventEditorPage.module.css'

const eventTypeOptions: EventType[] = [
  'single_day_all_day',
  'single_day_partial',
  'multi_day_all_day',
  'multi_day_partial',
]

const recurrenceFrequencyOptions: RecurrenceFrequency[] = [
  'daily',
  'weekly',
  'monthly',
  'yearly',
]

type SubmitMode = 'save' | 'draft' | 'publish'

type EventEditorPageProps = {
  mode?: 'edit' | 'view'
}

type EventTouchedFields = Partial<Record<string, boolean>>

function getUniqueErrorMessages(errors: EventFormErrors) {
  return Array.from(new Set(Object.values(errors).filter(Boolean)))
}

function filterVisibleEventErrors(
  errors: EventFormErrors,
  touchedFields: EventTouchedFields,
  showAllErrors: boolean,
) {
  if (showAllErrors) {
    return errors
  }

  const visibleErrors: EventFormErrors = {}

  for (const [key, message] of Object.entries(errors)) {
    if (message && shouldShowEventError(key, touchedFields)) {
      visibleErrors[key] = message
    }
  }

  return visibleErrors
}

function shouldShowEventError(key: string, touchedFields: EventTouchedFields) {
  if (touchedFields[key]) {
    return true
  }

  if (key.startsWith('scheduledOccurrences.')) {
    return Boolean(
      touchedFields.repeatEnabled ||
        touchedFields.recurrenceType ||
        touchedFields.scheduledOccurrences,
    )
  }

  switch (key) {
    case 'startTime':
    case 'endTime':
    case 'endDate':
      return Boolean(touchedFields.eventType)
    case 'privateAudiences':
      return Boolean(
        touchedFields.privacyType ||
          touchedFields.audienceMembers ||
          touchedFields.audiencePublic,
      )
    case 'reviewEmailsText':
      return Boolean(touchedFields.requestReview || touchedFields.reviewEmailsText)
    case 'selectedLocationId':
    case 'locationName':
    case 'addressLine1':
    case 'city':
    case 'provinceState':
    case 'postalCode':
    case 'country':
      return Boolean(touchedFields.locationChoice)
    case 'registrationStartDate':
    case 'registrationStartTime':
    case 'registrationEndDate':
    case 'registrationEndTime':
    case 'registrationUrl':
      return Boolean(touchedFields.registrationEnabled)
    case 'recurrenceType':
      return Boolean(touchedFields.repeatEnabled || touchedFields.recurrenceType)
    case 'recurrenceFrequency':
    case 'recurrenceInterval':
    case 'scheduledOccurrences':
      return Boolean(
        touchedFields.repeatEnabled ||
          touchedFields.recurrenceType ||
          touchedFields.scheduledOccurrences,
      )
    default:
      return false
  }
}

export function EventEditorPage({ mode = 'edit' }: EventEditorPageProps) {
  const dispatch = useAppDispatch()
  const navigate = useNavigate()
  const params = useParams()
  const { t } = useTranslation()
  const currentEvent = useAppSelector(selectCurrentEvent)
  const detailState = useAppSelector(selectEventDetail)
  const lookupsState = useAppSelector(selectEventLookups)
  const saveState = useAppSelector(selectEventSave)
  const savedLocations = useAppSelector(selectSavedLocations)
  const galleries = useAppSelector(selectEventGalleries)
  const deletingIds = useAppSelector(selectEventDeletingIds)

  const parsedEventId = params.id ? Number.parseInt(params.id, 10) : null
  const isEditMode = parsedEventId !== null && Number.isFinite(parsedEventId)
  const isViewMode = mode === 'view' && isEditMode
  const isInvalidEditId = params.id !== undefined && !isEditMode
  const [form, setForm] = useState<EventFormState>(createDefaultEventFormState())
  const [errors, setErrors] = useState<EventFormErrors>({})
  const [touchedFields, setTouchedFields] = useState<EventTouchedFields>({})
  const [hasAttemptedSubmit, setHasAttemptedSubmit] = useState(false)
  const [displayImagePreviewUrl, setDisplayImagePreviewUrl] = useState<string | null>(null)
  const [attachmentPreviewUrls, setAttachmentPreviewUrls] = useState<string[]>([])
  const [existingMediaObjectUrls, setExistingMediaObjectUrls] = useState<
    Record<number, string>
  >({})
  const [isDeletingEvent, setIsDeletingEvent] = useState(false)

  useEffect(() => {
    void dispatch(fetchEventLookups())
  }, [dispatch])

  useEffect(() => {
    if (isEditMode && parsedEventId) {
      void dispatch(fetchEventById(parsedEventId))
      return
    }

    dispatch(clearCurrentEvent())
    setForm(createDefaultEventFormState())
    setTouchedFields({})
    setHasAttemptedSubmit(false)
  }, [dispatch, isEditMode, parsedEventId])

  useEffect(() => {
    if (isEditMode && currentEvent) {
      setForm(buildEventFormStateFromDetail(currentEvent))
      setTouchedFields({})
      setHasAttemptedSubmit(false)
      setErrors({})
    }
  }, [currentEvent, isEditMode])

  useEffect(() => {
    if (isViewMode) {
      setErrors({})
      return
    }

    const validationErrors = validateEventForm(form, t)
    setErrors(
      filterVisibleEventErrors(validationErrors, touchedFields, hasAttemptedSubmit),
    )
  }, [form, hasAttemptedSubmit, isViewMode, t, touchedFields])

  useEffect(() => {
    return () => {
      dispatch(clearEventSaveState())
    }
  }, [dispatch])

  useEffect(() => {
    if (!form.displayImageFile) {
      setDisplayImagePreviewUrl(null)
      return
    }

    const objectUrl = URL.createObjectURL(form.displayImageFile)
    setDisplayImagePreviewUrl(objectUrl)

    return () => {
      URL.revokeObjectURL(objectUrl)
    }
  }, [form.displayImageFile])

  useEffect(() => {
    if (!form.attachmentFiles.length) {
      setAttachmentPreviewUrls([])
      return
    }

    const objectUrls = form.attachmentFiles.map((file) => URL.createObjectURL(file))
    setAttachmentPreviewUrls(objectUrls)

    return () => {
      objectUrls.forEach((objectUrl) => URL.revokeObjectURL(objectUrl))
    }
  }, [form.attachmentFiles])

  const existingMedia = useMemo(
    () => [
      ...(form.existingDisplayImage ? [form.existingDisplayImage] : []),
      ...form.existingAttachments,
    ],
    [form.existingDisplayImage, form.existingAttachments],
  )

  useEffect(() => {
    if (!existingMedia.length) {
      setExistingMediaObjectUrls((current) => {
        Object.values(current).forEach((objectUrl) => URL.revokeObjectURL(objectUrl))
        return {}
      })
      return
    }

    let cancelled = false

    async function loadExistingMedia() {
      const loadedEntries = await Promise.all(
        existingMedia.map(async (media) => {
          try {
            const objectUrl = await fetchEventMediaObjectUrl(media)
            return [media.id, objectUrl] as const
          } catch {
            return null
          }
        }),
      )

      const nextUrls = loadedEntries.reduce<Record<number, string>>((result, entry) => {
        if (entry) {
          result[entry[0]] = entry[1]
        }
        return result
      }, {})

      if (cancelled) {
        Object.values(nextUrls).forEach((objectUrl) => URL.revokeObjectURL(objectUrl))
        return
      }

      setExistingMediaObjectUrls((current) => {
        Object.values(current).forEach((objectUrl) => URL.revokeObjectURL(objectUrl))
        return nextUrls
      })
    }

    void loadExistingMedia()

    return () => {
      cancelled = true
    }
  }, [existingMedia])

  const isBusy = saveState.status === 'loading'
  const isInitialLoad =
    (isEditMode && detailState.status === 'loading' && !currentEvent) ||
    (lookupsState.status === 'loading' && !savedLocations.length && !galleries.length)
  const existingDisplayImageUrl = getExistingMediaObjectUrl(
    form.existingDisplayImage,
    existingMediaObjectUrls,
  )
  const errorMessages = useMemo(() => getUniqueErrorMessages(errors), [errors])
  const breadcrumbItems = useMemo(
    () => [
      { label: t('events.breadcrumb.events'), to: '/events' },
      {
        label: isViewMode
          ? t('events.breadcrumb.view')
          : isEditMode
            ? t('events.breadcrumb.edit')
            : t('events.breadcrumb.create'),
      },
    ],
    [isEditMode, isViewMode, t],
  )

  function touchFields(...keys: string[]) {
    setTouchedFields((current) => {
      const next = { ...current }
      for (const key of keys) {
        next[key] = true
      }
      return next
    })
  }

  function updateField<K extends keyof EventFormState>(
    key: K,
    value: EventFormState[K],
    options: { touch?: boolean; touchKeys?: string[] } = {},
  ) {
    const { touch = true, touchKeys = [String(key)] } = options
    if (touch) {
      touchFields(...touchKeys)
    }

    setForm((current) => ({
      ...current,
      [key]: value,
    }))
  }

  function handleEventTypeChange(eventType: EventType) {
    setForm((current) => {
      const next: EventFormState = {
        ...current,
        eventType,
        startTime: usesTime(eventType) ? current.startTime : '',
        endTime: usesTime(eventType) ? current.endTime : '',
        endDate: usesSeparateEndDate(eventType) ? current.endDate : '',
      }

      if (current.repeatEnabled && current.recurrenceType === 'scheduled') {
        next.scheduledOccurrences = current.scheduledOccurrences.length
          ? current.scheduledOccurrences.map((occurrence) => ({
              ...occurrence,
              startTime: usesTime(eventType) ? occurrence.startTime : '',
              endTime: usesTime(eventType) ? occurrence.endTime : '',
              endDate: usesSeparateEndDate(eventType) ? occurrence.endDate : '',
            }))
          : [createOccurrenceFromMain(next)]
      }

      return next
    })

    touchFields('eventType')
  }

  function handleRepeatEnabledChange(checked: boolean) {
    setForm((current) => {
      if (!checked) {
        return {
          ...current,
          repeatEnabled: false,
          recurrenceType: '',
          recurrenceFrequency: '',
          recurrenceInterval: '1',
          recurrenceUntilDate: '',
          scheduledOccurrences: [],
        }
      }

      return {
        ...current,
        repeatEnabled: true,
        recurrenceType: current.recurrenceType || 'scheduled',
        scheduledOccurrences:
          current.recurrenceType === 'scheduled' || !current.recurrenceType
            ? current.scheduledOccurrences.length
              ? current.scheduledOccurrences
              : [createOccurrenceFromMain(current)]
            : current.scheduledOccurrences,
      }
    })
    touchFields('repeatEnabled')
  }

  function handleRecurrenceTypeChange(value: 'scheduled' | 'recurring') {
    setForm((current) => ({
      ...current,
      recurrenceType: value,
      recurrenceFrequency: value === 'recurring' ? current.recurrenceFrequency : '',
      scheduledOccurrences:
        value === 'scheduled'
          ? current.scheduledOccurrences.length
            ? current.scheduledOccurrences
            : [createOccurrenceFromMain(current)]
          : [],
    }))
    touchFields('recurrenceType')
  }

  function handleOccurrenceChange(
    occurrenceId: string,
    key: keyof EventOccurrenceForm,
    value: string,
  ) {
    const occurrenceIndex = form.scheduledOccurrences.findIndex(
      (occurrence) => occurrence.id === occurrenceId,
    )
    if (occurrenceIndex >= 0) {
      touchFields(`scheduledOccurrences.${occurrenceIndex}.${String(key)}`)
    }

    setForm((current) => ({
      ...current,
      scheduledOccurrences: current.scheduledOccurrences.map((occurrence) =>
        occurrence.id === occurrenceId
          ? {
              ...occurrence,
              [key]: value,
            }
          : occurrence,
      ),
    }))
  }

  function addOccurrence() {
    setForm((current) => ({
      ...current,
      scheduledOccurrences: [
        ...current.scheduledOccurrences,
        createOccurrenceFromMain(current),
      ],
    }))
    touchFields('scheduledOccurrences')
  }

  function removeOccurrence(occurrenceId: string) {
    touchFields('scheduledOccurrences')
    setForm((current) => ({
      ...current,
      scheduledOccurrences: current.scheduledOccurrences.filter(
        (occurrence) => occurrence.id !== occurrenceId,
      ),
    }))
  }

  function handleDisplayImageFiles(files: File[]) {
    const [file] = files
    if (!file) {
      updateField('displayImageFile', null, { touch: false })
      return
    }

    const validationMessage = getEventDisplayImageValidationMessage(file)
    if (validationMessage) {
      setErrors((current) => ({ ...current, displayImage: validationMessage }))
      toast.error(validationMessage)
      return
    }

    setErrors((current) => ({ ...current, displayImage: undefined }))
    updateField('displayImageFile', file, { touch: false })
  }

  function handleAttachmentFiles(files: File[]) {
    if (!files.length) {
      return
    }

    const validFiles: File[] = []
    let validationMessage: string | null = null

    files.forEach((file) => {
      const nextValidationMessage = getEventAttachmentValidationMessage(file)
      if (nextValidationMessage) {
        validationMessage ??= nextValidationMessage
        return
      }

      validFiles.push(file)
    })

    if (validFiles.length > 0) {
      setErrors((current) => ({ ...current, attachments: undefined }))
      setForm((current) => ({
        ...current,
        attachmentFiles: [...current.attachmentFiles, ...validFiles],
      }))
    }

    if (validationMessage) {
      setErrors((current) => ({ ...current, attachments: validationMessage }))
      toast.error(validationMessage)
    }
  }

  function removeNewAttachment(index: number) {
    setForm((current) => ({
      ...current,
      attachmentFiles: current.attachmentFiles.filter((_, currentIndex) => currentIndex !== index),
    }))
  }

  async function removeExistingAttachment(attachment: EventMedia) {
    if (!isEditMode || !parsedEventId) {
      return
    }

    try {
      const storageUrl = attachment.storage_url ?? attachment.storage_uri
      if (!storageUrl) {
        return
      }

      await dispatch(
        deleteEventDocument({
          eventId: parsedEventId,
          mediaId: attachment.id,
          storageUrl,
        }),
      ).unwrap()
      dispatch(removeLocalAttachment(attachment.id))
      setForm((current) => ({
        ...current,
        existingAttachments: current.existingAttachments.filter(
          (item) => item.id !== attachment.id,
        ),
      }))
      toast.success(t('events.feedback.attachmentRemoved'))
    } catch {
      return
    }
  }

  async function removeExistingDisplayImage(media: EventMedia) {
    if (!isEditMode || !parsedEventId) {
      return
    }

    try {
      const storageUrl = media.storage_url ?? media.storage_uri
      if (!storageUrl) {
        return
      }

      await dispatch(
        deleteEventPhoto({
          eventId: parsedEventId,
          mediaId: media.id,
          storageUrl,
        }),
      ).unwrap()
      dispatch(removeLocalDisplayImage(media.id))
      setForm((current) => ({
        ...current,
        existingDisplayImage: current.existingDisplayImage?.id === media.id
          ? null
          : current.existingDisplayImage,
      }))
      toast.success(t('events.feedback.displayImageRemoved'))
    } catch {
      return
    }
  }

  async function submitForm(mode: SubmitMode) {
    const submissionState: EventFormState =
      mode === 'draft'
        ? {
            ...form,
            published: false,
          }
        : mode === 'publish'
          ? {
              ...form,
              published: true,
              requestReview: false,
              reviewEmailsText: '',
            }
          : form

    const validationErrors = validateEventForm(submissionState, t)
    if (Object.keys(validationErrors).length > 0) {
      setHasAttemptedSubmit(true)
      setErrors(validationErrors)
      toast.error(t('events.feedback.validation'))
      return
    }

    setHasAttemptedSubmit(false)
    setErrors({})

    try {
      const payload = buildSaveEventRequest(submissionState)
      const result =
        isEditMode && parsedEventId
          ? await dispatch(updateEvent({ id: parsedEventId, payload })).unwrap()
          : await dispatch(createEvent(payload)).unwrap()

      toast.success(
        t(isEditMode ? 'events.feedback.updated' : 'events.feedback.created'),
      )

      if (isEditMode && parsedEventId) {
        await dispatch(fetchEventById(parsedEventId)).unwrap()
      } else {
        navigate(`/events/${result.event.id}/edit`, { replace: true })
      }
    } catch {
      return
    }
  }

  async function handleDeleteEvent() {
    if (!isEditMode || !parsedEventId) {
      return
    }

    setIsDeletingEvent(true)
    try {
      await dispatch(deleteEventAction(parsedEventId)).unwrap()
      toast.success(t('events.feedback.deleted'))
      navigate('/events', { replace: true })
    } catch {
      setIsDeletingEvent(false)
    }
  }

  if (isInvalidEditId) {
    return (
      <CmsAppShell activeKey="events">
        <section className={styles.errorPanel}>
          <h1>{t('events.editor.invalidIdTitle')}</h1>
          <p>{t('events.editor.invalidIdText')}</p>
          <button
            type="button"
            className={styles.backLink}
            onClick={() => navigate('/events')}
          >
            {t('events.editor.backToList')}
          </button>
        </section>
      </CmsAppShell>
    )
  }

  if (isInitialLoad) {
    return (
      <CmsAppShell activeKey="events">
        <div className={styles.loaderWrap}>
          <Loader label={t('events.common.loading')} />
        </div>
      </CmsAppShell>
    )
  }

  if (isEditMode && detailState.error && !currentEvent) {
    return (
      <CmsAppShell activeKey="events">
        <section className={styles.errorPanel}>
          <h1>{t('events.editor.loadErrorTitle')}</h1>
          <p>{detailState.error}</p>
          <button
            type="button"
            className={styles.backLink}
            onClick={() => navigate('/events')}
          >
            {t('events.editor.backToList')}
          </button>
        </section>
      </CmsAppShell>
    )
  }

  return (
    <CmsAppShell activeKey="events">
      <div className={styles.page}>
        <Breadcrumb items={breadcrumbItems} />

        <div className={styles.header}>
          <div>
            <h1 className={styles.title}>
              {isViewMode
                ? t('events.editor.titleView')
                : isEditMode
                  ? t('events.editor.titleEdit')
                  : t('events.editor.titleCreate')}
            </h1>
          </div>
          <button
            type="button"
            className={styles.backLink}
            onClick={() => navigate('/events')}
          >
            {t('events.editor.backToList')}
          </button>
        </div>

        {!isViewMode && hasAttemptedSubmit && errorMessages.length > 0 && (
          <div className={styles.errorSummary} role="alert">
            <p className={styles.errorSummaryTitle}>{t('events.feedback.validation')}</p>
            <ul>
              {errorMessages.map((message) => (
                <li key={message}>{message}</li>
              ))}
            </ul>
          </div>
        )}

        {!isViewMode && saveState.error && (
          <div className={styles.errorSummary} role="alert">
            <p className={styles.errorSummaryTitle}>{saveState.error}</p>
          </div>
        )}

        <fieldset className={styles.readOnlyFieldset} disabled={isViewMode}>
          <div className={styles.layout}>
          <div className={styles.mainColumn}>
            <section className={styles.card}>
              <div className={styles.sectionHeader}>
                <h2>{t('events.sections.general')}</h2>
                <p>{t('events.sections.generalHint')}</p>
              </div>

              <div className={styles.fieldGrid}>
                <label className={styles.fieldFull}>
                  <span>{t('events.fields.title')}</span>
                  <input
                    type="text"
                    value={form.title}
                    onChange={(event) => updateField('title', event.target.value)}
                    onBlur={() => touchFields('title')}
                  />
                  <FieldError message={errors.title} />
                </label>

                <label className={styles.toggleRow}>
                  <div>
                    <span>{t('events.fields.showTitle')}</span>
                    <p>{t('events.fields.showTitleHint')}</p>
                  </div>
                  <input
                    type="checkbox"
                    checked={form.showTitle}
                    onChange={(event) => updateField('showTitle', event.target.checked)}
                  />
                </label>

                <label className={styles.fieldFull}>
                  <span>{t('events.fields.categories')}</span>
                  <input
                    type="text"
                    value={form.categoriesText}
                    placeholder={t('events.fields.categoriesPlaceholder')}
                    onChange={(event) => updateField('categoriesText', event.target.value)}
                    onBlur={() => touchFields('categoriesText')}
                  />
                  <FieldError message={errors.categoriesText} />
                </label>
              </div>
            </section>

            <section className={styles.card}>
              <div className={styles.sectionHeader}>
                <h2>{t('events.sections.contact')}</h2>
                <p>{t('events.sections.contactHint')}</p>
              </div>

              <div className={styles.fieldGrid}>
                <LabeledInput
                  label={t('events.fields.contactName')}
                  value={form.contactName}
                  onChange={(value) => updateField('contactName', value)}
                  onBlur={() => touchFields('contactName')}
                />
                <LabeledInput
                  label={t('events.fields.contactEmail')}
                  type="email"
                  value={form.contactEmail}
                  onChange={(value) => updateField('contactEmail', value)}
                  onBlur={() => touchFields('contactEmail')}
                  error={errors.contactEmail}
                />
                <LabeledInput
                  label={t('events.fields.contactPhone')}
                  type="tel"
                  inputMode="tel"
                  value={form.contactPhone}
                  onChange={(value) => updateField('contactPhone', value)}
                  onBlur={() => touchFields('contactPhone')}
                  error={errors.contactPhone}
                />
                <LabeledInput
                  label={t('events.fields.contactExt')}
                  inputMode="numeric"
                  maxLength={6}
                  value={form.contactExt}
                  onChange={(value) => updateField('contactExt', value)}
                  onBlur={() => touchFields('contactExt')}
                  error={errors.contactExt}
                />
                <LabeledInput
                  label={t('events.fields.contactFax')}
                  type="tel"
                  inputMode="tel"
                  value={form.contactFax}
                  onChange={(value) => updateField('contactFax', value)}
                  onBlur={() => touchFields('contactFax')}
                  error={errors.contactFax}
                />
              </div>
            </section>

            <section className={styles.card}>
              <div className={styles.sectionHeader}>
                <h2>{t('events.sections.dateAndTime')}</h2>
                <p>{t('events.sections.dateAndTimeHint')}</p>
              </div>

              <div className={styles.optionGrid}>
                {eventTypeOptions.map((option) => (
                  <label
                    key={option}
                    className={[
                      styles.optionCard,
                      form.eventType === option ? styles.optionCardActive : '',
                    ].join(' ')}
                  >
                    <input
                      type="radio"
                      name="eventType"
                      checked={form.eventType === option}
                      onChange={() => handleEventTypeChange(option)}
                    />
                    <div>
                      <strong>{t(`events.eventTypes.${option}`)}</strong>
                    </div>
                  </label>
                ))}
              </div>

              <div className={styles.fieldGrid}>
                <label>
                  <span>{t('events.fields.startDate')}</span>
                  <DateInput
                    value={form.startDate}
                    onChange={(value) => updateField('startDate', value)}
                    onBlur={() => touchFields('startDate')}
                  />
                  <FieldError message={errors.startDate} />
                </label>

                {usesSeparateEndDate(form.eventType) && (
                  <label>
                    <span>{t('events.fields.endDate')}</span>
                    <DateInput
                      value={form.endDate}
                      onChange={(value) => updateField('endDate', value)}
                      onBlur={() => touchFields('endDate')}
                    />
                    <FieldError message={errors.endDate} />
                  </label>
                )}

                {usesTime(form.eventType) && (
                  <label>
                    <span>{t('events.fields.startTime')}</span>
                    <input
                      type="time"
                      value={form.startTime}
                      onChange={(event) => updateField('startTime', event.target.value)}
                      onBlur={() => touchFields('startTime')}
                    />
                    <FieldError message={errors.startTime} />
                  </label>
                )}

                {usesTime(form.eventType) && (
                  <label>
                    <span>{t('events.fields.endTime')}</span>
                    <input
                      type="time"
                      value={form.endTime}
                      onChange={(event) => updateField('endTime', event.target.value)}
                      onBlur={() => touchFields('endTime')}
                    />
                    <FieldError message={errors.endTime} />
                  </label>
                )}
              </div>

              <label className={styles.toggleRow}>
                <div>
                  <span>{t('events.fields.repeatEnabled')}</span>
                  <p>{t('events.fields.repeatEnabledHint')}</p>
                </div>
                <input
                  type="checkbox"
                  checked={form.repeatEnabled}
                  onChange={(event) => handleRepeatEnabledChange(event.target.checked)}
                />
              </label>

              {form.repeatEnabled && (
                <div className={styles.repeatPanel}>
                  <div className={styles.repeatTypeGrid}>
                    {(['scheduled', 'recurring'] as const).map((option) => (
                      <label
                        key={option}
                        className={[
                          styles.optionCard,
                          form.recurrenceType === option ? styles.optionCardActive : '',
                        ].join(' ')}
                      >
                        <input
                          type="radio"
                          name="recurrenceType"
                          checked={form.recurrenceType === option}
                          onChange={() => handleRecurrenceTypeChange(option)}
                        />
                        <div>
                          <strong>{t(`events.recurrenceTypes.${option}.label`)}</strong>
                          <p>{t(`events.recurrenceTypes.${option}.hint`)}</p>
                        </div>
                      </label>
                    ))}
                  </div>
                  <FieldError message={errors.recurrenceType} />

                  {form.recurrenceType === 'recurring' && (
                    <div className={styles.fieldGrid}>
                      <label>
                        <span>{t('events.fields.recurrenceFrequency')}</span>
                        <select
                          value={form.recurrenceFrequency}
                          onChange={(event) =>
                            updateField(
                              'recurrenceFrequency',
                              event.target.value as EventFormState['recurrenceFrequency'],
                            )
                          }
                          onBlur={() => touchFields('recurrenceFrequency')}
                        >
                          <option value="">{t('events.common.select')}</option>
                          {recurrenceFrequencyOptions.map((option) => (
                            <option key={option} value={option}>
                              {t(`events.recurrenceFrequency.${option}`)}
                            </option>
                          ))}
                        </select>
                        <FieldError message={errors.recurrenceFrequency} />
                      </label>

                      <label>
                        <span>{t('events.fields.recurrenceInterval')}</span>
                        <input
                          type="number"
                          min="1"
                          value={form.recurrenceInterval}
                          onChange={(event) =>
                            updateField('recurrenceInterval', event.target.value)
                          }
                          onBlur={() => touchFields('recurrenceInterval')}
                        />
                        <FieldError message={errors.recurrenceInterval} />
                      </label>

                      <label>
                        <span>{t('events.fields.recurrenceUntilDate')}</span>
                        <DateInput
                          value={form.recurrenceUntilDate}
                          onChange={(value) => updateField('recurrenceUntilDate', value)}
                          onBlur={() => touchFields('recurrenceUntilDate')}
                        />
                      </label>
                    </div>
                  )}

                  {form.recurrenceType === 'scheduled' && (
                    <div className={styles.occurrencePanel}>
                      <div className={styles.occurrenceHeader}>
                        <div>
                          <h3>{t('events.fields.scheduledOccurrences')}</h3>
                          <p>{t('events.fields.scheduledOccurrencesHint')}</p>
                        </div>
                        <button
                          type="button"
                          className={styles.secondaryButton}
                          onClick={addOccurrence}
                        >
                          {t('events.editor.addOccurrence')}
                        </button>
                      </div>

                      <FieldError message={errors.scheduledOccurrences} />

                      {form.scheduledOccurrences.map((occurrence, index) => (
                        <div key={occurrence.id} className={styles.occurrenceCard}>
                          <div className={styles.occurrenceHeader}>
                            <strong>
                              {t('events.editor.occurrenceLabel', { number: index + 1 })}
                            </strong>
                            {form.scheduledOccurrences.length > 1 && (
                              <button
                                type="button"
                                className={styles.linkButton}
                                onClick={() => removeOccurrence(occurrence.id)}
                              >
                                {t('events.editor.removeOccurrence')}
                              </button>
                            )}
                          </div>

                          <div className={styles.fieldGrid}>
                            <OccurrenceField
                              label={t('events.fields.startDate')}
                              type="date"
                              value={occurrence.startDate}
                              onChange={(value) =>
                                handleOccurrenceChange(occurrence.id, 'startDate', value)
                              }
                              onBlur={() =>
                                touchFields(`scheduledOccurrences.${index}.startDate`)
                              }
                              error={errors[`scheduledOccurrences.${index}.startDate`]}
                            />

                            {usesSeparateEndDate(form.eventType) && (
                              <OccurrenceField
                                label={t('events.fields.endDate')}
                                type="date"
                                value={occurrence.endDate}
                                onChange={(value) =>
                                  handleOccurrenceChange(occurrence.id, 'endDate', value)
                                }
                                onBlur={() =>
                                  touchFields(`scheduledOccurrences.${index}.endDate`)
                                }
                                error={errors[`scheduledOccurrences.${index}.endDate`]}
                              />
                            )}

                            {usesTime(form.eventType) && (
                              <OccurrenceField
                                label={t('events.fields.startTime')}
                                type="time"
                                value={occurrence.startTime}
                                onChange={(value) =>
                                  handleOccurrenceChange(occurrence.id, 'startTime', value)
                                }
                                onBlur={() =>
                                  touchFields(`scheduledOccurrences.${index}.startTime`)
                                }
                                error={errors[`scheduledOccurrences.${index}.startTime`]}
                              />
                            )}

                            {usesTime(form.eventType) && (
                              <OccurrenceField
                                label={t('events.fields.endTime')}
                                type="time"
                                value={occurrence.endTime}
                                onChange={(value) =>
                                  handleOccurrenceChange(occurrence.id, 'endTime', value)
                                }
                                onBlur={() =>
                                  touchFields(`scheduledOccurrences.${index}.endTime`)
                                }
                                error={errors[`scheduledOccurrences.${index}.endTime`]}
                              />
                            )}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </section>

            <section className={styles.card}>
              <div className={styles.sectionHeader}>
                <h2>{t('events.sections.registration')}</h2>
                <p>{t('events.sections.registrationHint')}</p>
              </div>

              <label className={styles.toggleRow}>
                <div>
                  <span>{t('events.fields.registrationEnabled')}</span>
                  <p>{t('events.fields.registrationEnabledHint')}</p>
                </div>
                <input
                  type="checkbox"
                  checked={form.registrationEnabled}
                  onChange={(event) =>
                    updateField('registrationEnabled', event.target.checked)
                  }
                />
              </label>

              {form.registrationEnabled && (
                <div className={styles.fieldGrid}>
                  <label>
                    <span>{t('events.fields.registrationStartDate')}</span>
                    <DateInput
                      value={form.registrationStartDate}
                      onChange={(value) => updateField('registrationStartDate', value)}
                      onBlur={() => touchFields('registrationStartDate')}
                    />
                    <FieldError message={errors.registrationStartDate} />
                  </label>
                  <label>
                    <span>{t('events.fields.registrationStartTime')}</span>
                    <input
                      type="time"
                      value={form.registrationStartTime}
                      onChange={(event) =>
                        updateField('registrationStartTime', event.target.value)
                      }
                      onBlur={() => touchFields('registrationStartTime')}
                    />
                    <FieldError message={errors.registrationStartTime} />
                  </label>
                  <label>
                    <span>{t('events.fields.registrationEndDate')}</span>
                    <DateInput
                      value={form.registrationEndDate}
                      onChange={(value) => updateField('registrationEndDate', value)}
                      onBlur={() => touchFields('registrationEndDate')}
                    />
                    <FieldError message={errors.registrationEndDate} />
                  </label>
                  <label>
                    <span>{t('events.fields.registrationEndTime')}</span>
                    <input
                      type="time"
                      value={form.registrationEndTime}
                      onChange={(event) =>
                        updateField('registrationEndTime', event.target.value)
                      }
                      onBlur={() => touchFields('registrationEndTime')}
                    />
                    <FieldError message={errors.registrationEndTime} />
                  </label>
                  <label className={styles.fieldFull}>
                    <span>{t('events.fields.registrationUrl')}</span>
                    <input
                      type="url"
                      value={form.registrationUrl}
                      onChange={(event) => updateField('registrationUrl', event.target.value)}
                      onBlur={() => touchFields('registrationUrl')}
                    />
                    <FieldError message={errors.registrationUrl} />
                  </label>
                </div>
              )}
            </section>

            <section className={styles.card}>
              <div className={styles.sectionHeader}>
                <h2>{t('events.sections.media')}</h2>
                <p>{t('events.sections.mediaHint')}</p>
              </div>

              <div className={styles.fieldGrid}>
                <div className={styles.fieldFull}>
                  <span>{t('events.fields.displayImage')}</span>
                  <UploadDropzone
                    accept={RESOURCE_IMAGE_FILE_ACCEPT}
                    label={t('events.fields.displayImage')}
                    onFiles={handleDisplayImageFiles}
                  />
                  <FieldError message={errors.displayImage} />
                </div>

                {(form.existingDisplayImage || form.displayImageFile) && (
                  <div className={`${styles.mediaList} ${styles.displayImageList}`}>
                    {form.existingDisplayImage && (
                      <div className={`${styles.mediaItem} ${styles.mediaPreviewCard}`}>
                        <div className={styles.mediaPreviewSurface}>
                          {existingDisplayImageUrl ? (
                            <img
                              className={styles.mediaPreviewImage}
                              src={existingDisplayImageUrl}
                              alt={form.existingDisplayImage.display_name || t('events.fields.displayImage')}
                            />
                          ) : (
                            <div className={styles.mediaPreviewFallback} aria-hidden="true" />
                          )}
                        </div>
                        <MediaPreviewFooter
                          label={t('events.editor.currentDisplayImage')}
                          title={
                            form.existingDisplayImage.display_name ||
                            t('events.fields.displayImage')
                          }
                          href={existingDisplayImageUrl}
                          actionLabel={t('events.editor.openMedia')}
                        />
                        <button
                          type="button"
                          className={`${styles.linkButton} ${styles.mediaPreviewAction}`}
                          disabled={deletingIds.includes(form.existingDisplayImage.id)}
                          onClick={() =>
                            removeExistingDisplayImage(form.existingDisplayImage!)
                          }
                        >
                          {deletingIds.includes(form.existingDisplayImage.id)
                            ? t('events.common.loading')
                            : t('events.editor.removeMedia')}
                        </button>
                      </div>
                    )}

                    {form.displayImageFile && (
                      <div className={`${styles.mediaItem} ${styles.mediaPreviewCard}`}>
                        <div className={styles.mediaPreviewSurface}>
                          {displayImagePreviewUrl ? (
                            <img
                              className={styles.mediaPreviewImage}
                              src={displayImagePreviewUrl}
                              alt={form.displayImageFile.name || t('events.fields.displayImage')}
                            />
                          ) : (
                            <div className={styles.mediaPreviewFallback} aria-hidden="true" />
                          )}
                        </div>
                        <MediaPreviewFooter
                          label={t('events.editor.newUpload')}
                          title={form.displayImageFile.name || t('events.fields.displayImage')}
                        />
                        <button
                          type="button"
                          className={`${styles.linkButton} ${styles.mediaPreviewAction}`}
                          onClick={() => updateField('displayImageFile', null)}
                        >
                          {t('events.editor.removeMedia')}
                        </button>
                      </div>
                    )}
                  </div>
                )}

                <label className={`${styles.toggleRow} ${styles.displayImageToggle}`}>
                  <div>
                    <span>{t('events.fields.showDisplayImage')}</span>
                    <p>{t('events.fields.showDisplayImageHint')}</p>
                  </div>
                  <input
                    type="checkbox"
                    checked={form.showDisplayImageWhenViewing}
                    onChange={(event) =>
                      updateField('showDisplayImageWhenViewing', event.target.checked)
                    }
                  />
                </label>

                <label className={styles.fieldFull}>
                  <span>{t('events.fields.gallery')}</span>
                  <select
                    value={form.galleryId}
                    onChange={(event) => updateField('galleryId', event.target.value)}
                  >
                    <option value="">{t('events.common.none')}</option>
                    {galleries.map((gallery) => (
                      <option key={gallery.id} value={gallery.id}>
                        {gallery.name}
                      </option>
                    ))}
                  </select>
                </label>

                <label className={styles.fieldFull}>
                  <span>{t('events.fields.teaser')}</span>
                  <textarea
                    rows={3}
                    value={form.teaser}
                    onChange={(event) => updateField('teaser', event.target.value)}
                    onBlur={() => touchFields('teaser')}
                  />
                  <FieldError message={errors.teaser} />
                </label>

                <label className={styles.fieldFull}>
                  <span>{t('events.fields.description')}</span>
                  <textarea
                    rows={8}
                    value={form.descriptionHtml}
                    onChange={(event) => updateField('descriptionHtml', event.target.value)}
                    onBlur={() => touchFields('descriptionHtml')}
                  />
                </label>

                <div className={styles.fieldFull}>
                  <span>{t('events.fields.attachments')}</span>
                  <UploadDropzone
                    multiple
                    accept={RESOURCE_FILE_ACCEPT}
                    label={t('events.fields.attachments')}
                    onFiles={handleAttachmentFiles}
                  />
                  <FieldError message={errors.attachments} />
                </div>

                {(form.existingAttachments.length > 0 || form.attachmentFiles.length > 0) && (
                  <div className={styles.mediaGrid}>
                    {form.existingAttachments.map((attachment) => (
                      <div
                        key={attachment.id}
                        className={`${styles.mediaItem} ${styles.mediaPreviewCard}`}
                      >
                        <div className={styles.mediaPreviewSurface}>
                          {isPreviewableImage(
                            attachment.display_name,
                            attachment.mime_type,
                          ) && existingMediaObjectUrls[attachment.id] ? (
                            <img
                              className={styles.mediaPreviewImage}
                              src={existingMediaObjectUrls[attachment.id]}
                              alt={attachment.display_name || t('events.fields.attachments')}
                            />
                          ) : (
                            <div className={styles.documentPreviewFallback} aria-hidden="true">
                              <span className={styles.documentTypeBadge}>
                                {getFileTypeBadge(
                                  attachment.display_name,
                                  attachment.mime_type,
                                )}
                              </span>
                            </div>
                          )}
                        </div>
                        <MediaPreviewFooter
                          label={t('events.editor.currentAttachment')}
                          title={
                            attachment.display_name || t('events.fields.attachments')
                          }
                          href={existingMediaObjectUrls[attachment.id]}
                          actionLabel={t('events.editor.openMedia')}
                        />
                        <button
                          type="button"
                          className={`${styles.linkButton} ${styles.mediaPreviewAction}`}
                          disabled={deletingIds.includes(attachment.id)}
                          onClick={() => removeExistingAttachment(attachment)}
                        >
                          {deletingIds.includes(attachment.id)
                            ? t('events.common.loading')
                            : t('events.editor.removeMedia')}
                        </button>
                      </div>
                    ))}

                    {form.attachmentFiles.map((file, index) => (
                      <div
                        key={`${file.name}-${index}`}
                        className={`${styles.mediaItem} ${styles.mediaPreviewCard}`}
                      >
                        <div className={styles.mediaPreviewSurface}>
                          {isPreviewableImage(file.name, file.type) &&
                          attachmentPreviewUrls[index] ? (
                            <img
                              className={styles.mediaPreviewImage}
                              src={attachmentPreviewUrls[index]}
                              alt={file.name || t('events.fields.attachments')}
                            />
                          ) : (
                            <div className={styles.documentPreviewFallback} aria-hidden="true">
                              <span className={styles.documentTypeBadge}>
                                {getFileTypeBadge(file.name, file.type)}
                              </span>
                            </div>
                          )}
                        </div>
                        <MediaPreviewFooter
                          label={t('events.editor.newUpload')}
                          title={file.name || t('events.fields.attachments')}
                        />
                        <button
                          type="button"
                          className={`${styles.linkButton} ${styles.mediaPreviewAction}`}
                          onClick={() => removeNewAttachment(index)}
                        >
                          {t('events.editor.removeMedia')}
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </section>

            {!isViewMode && (
              <EntryActions
                entryType="event"
                status={form.published ? 'published' : 'draft'}
                publishLabel={
                  form.published
                    ? t('events.editor.saveChanges')
                    : t('events.editor.publish')
                }
                deleteLabel={t('events.list.delete')}
                onSaveDraft={() => void submitForm('draft')}
                onPublish={() =>
                  void submitForm(form.published ? 'save' : 'publish')
                }
                onDelete={isEditMode ? () => void handleDeleteEvent() : undefined}
                deleteConfirmTitle={t('events.list.deleteDialogTitle')}
                deleteConfirmBody={
                  <Trans
                    i18nKey="events.list.deleteDialogDescription"
                    values={{ title: form.title || t('events.editor.titleCreate') }}
                    components={{ eventName: <strong /> }}
                  />
                }
                deleteConfirmLabel={t('events.list.delete')}
                isSubmitting={isBusy}
                isDeleting={isDeletingEvent}
              />
            )}
          </div>

          <aside className={styles.sideColumn}>
            {isViewMode && parsedEventId ? (
              <section className={styles.card}>
                <div className={styles.sectionHeader}>
                  <h2>{t('events.sections.actions')}</h2>
                </div>
                <div className={styles.statusChipRow}>
                  <StatusBadge status={form.published ? 'published' : 'draft'} />
                </div>
                <div className={styles.actionStack}>
                  <Link
                    className={`${styles.primaryButton} ${styles.buttonLink}`}
                    to={`/events/${parsedEventId}/edit`}
                  >
                    {t('events.editor.editEvent')}
                  </Link>
                </div>
              </section>
            ) : (
              <PublishingControls
                status={form.published ? 'published' : 'draft'}
                visibility={form.privacyType}
                onVisibilityChange={(value) => updateField('privacyType', value)}
                extraSlot={
                  <>
                    {form.privacyType === 'private' && (
                      <div className={styles.stack}>
                        <span className={styles.groupLabel}>
                          {t('events.fields.privateAudiences')}
                        </span>
                        <label className={styles.checkboxRow}>
                          <input
                            type="checkbox"
                            checked={form.audienceMembers}
                            onChange={(event) =>
                              updateField('audienceMembers', event.target.checked)
                            }
                          />
                          <span>{t('events.audiences.members')}</span>
                        </label>
                        <label className={styles.checkboxRow}>
                          <input
                            type="checkbox"
                            checked={form.audiencePublic}
                            onChange={(event) =>
                              updateField('audiencePublic', event.target.checked)
                            }
                          />
                          <span>{t('events.audiences.public')}</span>
                        </label>
                        <FieldError message={errors.privateAudiences} />
                      </div>
                    )}

                    {!form.published && (
                      <>
                        <label className={styles.toggleRow}>
                          <div>
                            <span>{t('events.fields.requestReview')}</span>
                            <p>{t('events.fields.requestReviewHint')}</p>
                          </div>
                          <input
                            type="checkbox"
                            checked={form.requestReview}
                            onChange={(event) =>
                              updateField('requestReview', event.target.checked)
                            }
                          />
                        </label>

                        {form.requestReview && (
                          <label className={styles.reviewEmailsLabel}>
                            <span>{t('events.fields.reviewEmails')}</span>
                            <textarea
                              className={styles.reviewEmailsTextarea}
                              rows={3}
                              value={form.reviewEmailsText}
                              placeholder={t(
                                'events.fields.reviewEmailsPlaceholder',
                              )}
                              onChange={(event) =>
                                updateField(
                                  'reviewEmailsText',
                                  event.target.value,
                                )
                              }
                              onBlur={() => touchFields('reviewEmailsText')}
                            />
                            <FieldError message={errors.reviewEmailsText} />
                          </label>
                        )}
                      </>
                    )}
                  </>
                }
              />
            )}

            <section className={styles.card}>
              <div className={styles.sectionHeader}>
                <h2>{t('events.sections.location')}</h2>
                <p>{t('events.sections.locationHint')}</p>
              </div>

              <div className={styles.stack}>
                <label>
                  <span>{t('events.fields.locationMode')}</span>
                  <select
                    value={form.locationChoice}
                    onChange={(event) =>
                      updateField(
                        'locationChoice',
                        event.target.value as EventFormState['locationChoice'],
                      )
                    }
                    onBlur={() => touchFields('locationChoice')}
                  >
                    <option value="none">{t('events.locationMode.none')}</option>
                    <option value="to_be_determined">
                      {t('events.locationMode.to_be_determined')}
                    </option>
                    <option value="saved">{t('events.locationMode.saved')}</option>
                    <option value="new">{t('events.locationMode.new')}</option>
                  </select>
                </label>

                {form.locationChoice === 'saved' && (
                  <label>
                    <span>{t('events.fields.savedLocation')}</span>
                    <select
                      value={form.selectedLocationId}
                      onChange={(event) =>
                        updateField('selectedLocationId', event.target.value)
                      }
                      onBlur={() => touchFields('selectedLocationId')}
                    >
                      <option value="">{t('events.common.select')}</option>
                      {savedLocations.map((location) => (
                        <option key={location.id} value={location.id}>
                          {buildLocationSummary(location)}
                        </option>
                      ))}
                    </select>
                    <FieldError message={errors.selectedLocationId} />
                  </label>
                )}

                {form.locationChoice === 'new' && (
                  <div className={styles.fieldGrid}>
                    <LabeledInput
                      label={t('events.fields.locationName')}
                      value={form.locationName}
                      onChange={(value) => updateField('locationName', value)}
                      onBlur={() => touchFields('locationName')}
                      error={errors.locationName}
                    />
                    <LabeledInput
                      label={t('events.fields.addressLine1')}
                      value={form.addressLine1}
                      onChange={(value) => updateField('addressLine1', value)}
                      onBlur={() => touchFields('addressLine1')}
                      error={errors.addressLine1}
                    />
                    <LabeledInput
                      label={t('events.fields.addressLine2')}
                      value={form.addressLine2}
                      onChange={(value) => updateField('addressLine2', value)}
                      onBlur={() => touchFields('addressLine2')}
                    />
                    <LabeledInput
                      label={t('events.fields.city')}
                      value={form.city}
                      onChange={(value) => updateField('city', value)}
                      onBlur={() => touchFields('city')}
                      error={errors.city}
                    />
                    <LabeledInput
                      label={t('events.fields.provinceState')}
                      value={form.provinceState}
                      onChange={(value) => updateField('provinceState', value)}
                      onBlur={() => touchFields('provinceState')}
                      error={errors.provinceState}
                    />
                    <LabeledInput
                      label={t('events.fields.postalCode')}
                      value={form.postalCode}
                      onChange={(value) => updateField('postalCode', value)}
                      onBlur={() => touchFields('postalCode')}
                      error={errors.postalCode}
                    />
                    <LabeledInput
                      label={t('events.fields.country')}
                      value={form.country}
                      onChange={(value) => updateField('country', value)}
                      onBlur={() => touchFields('country')}
                      error={errors.country}
                    />

                    <label className={styles.toggleRow}>
                      <div>
                        <span>{t('events.fields.saveAddress')}</span>
                        <p>{t('events.fields.saveAddressHint')}</p>
                      </div>
                      <input
                        type="checkbox"
                        checked={form.saveAddress}
                        onChange={(event) => updateField('saveAddress', event.target.checked)}
                      />
                    </label>
                  </div>
                )}
              </div>
            </section>

          </aside>
        </div>
        </fieldset>
      </div>
    </CmsAppShell>
  )
}

type LabeledInputProps = {
  label: string
  value: string
  type?: string
  inputMode?: InputHTMLAttributes<HTMLInputElement>['inputMode']
  autoComplete?: string
  maxLength?: number
  error?: string
  onBlur?: () => void
  onChange: (value: string) => void
}

function LabeledInput({
  label,
  value,
  type = 'text',
  inputMode,
  autoComplete,
  maxLength,
  error,
  onBlur,
  onChange,
}: LabeledInputProps) {
  return (
    <label>
      <span>{label}</span>
      <input
        type={type}
        inputMode={inputMode}
        autoComplete={autoComplete}
        maxLength={maxLength}
        value={value}
        onChange={(event) => onChange(event.target.value)}
        onBlur={onBlur}
      />
      <FieldError message={error} />
    </label>
  )
}

type OccurrenceFieldProps = {
  label: string
  type: 'date' | 'time'
  value: string
  error?: string
  onBlur?: () => void
  onChange: (value: string) => void
}

function OccurrenceField({
  label,
  type,
  value,
  error,
  onBlur,
  onChange,
}: OccurrenceFieldProps) {
  return (
    <label>
      <span>{label}</span>
      {type === 'date' ? (
        <DateInput value={value} onChange={onChange} onBlur={onBlur} />
      ) : (
        <input
          type={type}
          value={value}
          onChange={(event) => onChange(event.target.value)}
          onBlur={onBlur}
        />
      )}
      <FieldError message={error} />
    </label>
  )
}

function FieldError({ message }: { message?: string }) {
  if (!message) {
    return null
  }

  return <span className={styles.fieldError}>{message}</span>
}

function getEventAttachmentValidationMessage(file: Pick<File, 'name' | 'size' | 'type'>) {
  const validationError = validateResourceUploadFile(file)
  if (!validationError) {
    return null
  }

  return getResourceUploadValidationErrorMessage(validationError)
}

function getEventDisplayImageValidationMessage(file: Pick<File, 'name' | 'size' | 'type'>) {
  const validationError = validateResourceImageUploadFile(file)
  if (!validationError) {
    return null
  }

  return getResourceImageUploadValidationErrorMessage(validationError)
}

type MediaPreviewFooterProps = {
  label: string
  title: string
  href?: string | null
  actionLabel?: string
}

function MediaPreviewFooter({
  label,
  title,
  href,
  actionLabel,
}: MediaPreviewFooterProps) {
  return (
    <div className={styles.mediaPreviewFooter}>
      <div className={styles.mediaPreviewMeta}>
        <p className={styles.mediaPreviewLabel}>{label}</p>
        <strong className={styles.mediaPreviewTitle}>{title}</strong>
      </div>
      {href && actionLabel ? (
        <a
          className={`${styles.linkButton} ${styles.mediaPreviewLink}`}
          href={href}
          target="_blank"
          rel="noreferrer"
        >
          {actionLabel}
        </a>
      ) : null}
    </div>
  )
}

function isPreviewableImage(fileName: string, mimeType?: string) {
  if (mimeType?.startsWith('image/')) {
    return true
  }

  const extension = getFileExtension(fileName)
  return ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg', 'avif'].includes(
    extension,
  )
}

function getFileTypeBadge(fileName: string, mimeType?: string) {
  const extension = getFileExtension(fileName)
  if (extension) {
    return extension.slice(0, 4).toUpperCase()
  }

  if (mimeType === 'application/pdf') {
    return 'PDF'
  }
  if (
    mimeType?.includes('word') ||
    mimeType?.includes('document')
  ) {
    return 'DOC'
  }
  if (
    mimeType?.includes('sheet') ||
    mimeType?.includes('excel')
  ) {
    return 'XLS'
  }
  if (
    mimeType?.includes('presentation') ||
    mimeType?.includes('powerpoint')
  ) {
    return 'PPT'
  }

  return 'FILE'
}

function getFileExtension(fileName: string) {
  const [, extension = ''] = fileName.toLowerCase().match(/\.([^.]+)$/) ?? []
  return extension
}

function getExistingMediaObjectUrl(
  media: EventMedia | null | undefined,
  mediaUrls: Record<number, string>,
) {
  if (!media) {
    return null
  }

  return mediaUrls[media.id] ?? null
}
