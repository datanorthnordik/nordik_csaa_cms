import {
  type EventAddress,
  type EventAddressInput,
  type EventDetailResponse,
  type EventMedia,
  type EventOccurrence,
  type EventOccurrenceInput,
  type EventType,
  type RecurrenceFrequency,
  type SaveEventPayload,
} from '../api/eventsApi'
import { isAllDayEventType } from './eventsDate'

export type EventLocationChoice = 'none' | 'to_be_determined' | 'saved' | 'new'

export type EventOccurrenceForm = {
  id: string
  startDate: string
  endDate: string
  startTime: string
  endTime: string
}

export type EventFormState = {
  title: string
  showTitle: boolean
  categoriesText: string
  eventType: EventType
  startDate: string
  endDate: string
  startTime: string
  endTime: string
  privacyType: 'public' | 'private'
  audienceMembers: boolean
  audiencePublic: boolean
  published: boolean
  requestReview: boolean
  reviewEmailsText: string
  teaser: string
  descriptionHtml: string
  contactName: string
  contactEmail: string
  contactPhone: string
  contactExt: string
  contactFax: string
  locationChoice: EventLocationChoice
  selectedLocationId: string
  locationName: string
  addressLine1: string
  addressLine2: string
  city: string
  provinceState: string
  postalCode: string
  country: string
  saveAddress: boolean
  showDisplayImageWhenViewing: boolean
  galleryId: string
  registrationEnabled: boolean
  registrationStartDate: string
  registrationStartTime: string
  registrationEndDate: string
  registrationEndTime: string
  registrationUrl: string
  repeatEnabled: boolean
  recurrenceType: '' | 'scheduled' | 'recurring'
  recurrenceFrequency: '' | RecurrenceFrequency
  recurrenceInterval: string
  recurrenceUntilDate: string
  scheduledOccurrences: EventOccurrenceForm[]
  displayImageFile: File | null
  existingDisplayImage: EventMedia | null
  attachmentFiles: File[]
  existingAttachments: EventMedia[]
}

export type EventFormErrors = Partial<Record<string, string>>

type TranslateFn = (key: string, options?: Record<string, unknown>) => string
type InputDateOptions = {
  preserveCalendarDate?: boolean
}

export function createDefaultEventFormState(): EventFormState {
  return {
    title: '',
    showTitle: true,
    categoriesText: '',
    eventType: 'single_day_all_day',
    startDate: '',
    endDate: '',
    startTime: '',
    endTime: '',
    privacyType: 'public',
    audienceMembers: false,
    audiencePublic: false,
    published: false,
    requestReview: false,
    reviewEmailsText: '',
    teaser: '',
    descriptionHtml: '',
    contactName: '',
    contactEmail: '',
    contactPhone: '',
    contactExt: '',
    contactFax: '',
    locationChoice: 'none',
    selectedLocationId: '',
    locationName: '',
    addressLine1: '',
    addressLine2: '',
    city: '',
    provinceState: '',
    postalCode: '',
    country: 'Canada',
    saveAddress: false,
    showDisplayImageWhenViewing: true,
    galleryId: '',
    registrationEnabled: false,
    registrationStartDate: '',
    registrationStartTime: '',
    registrationEndDate: '',
    registrationEndTime: '',
    registrationUrl: '',
    repeatEnabled: false,
    recurrenceType: '',
    recurrenceFrequency: '',
    recurrenceInterval: '1',
    recurrenceUntilDate: '',
    scheduledOccurrences: [],
    displayImageFile: null,
    existingDisplayImage: null,
    attachmentFiles: [],
    existingAttachments: [],
  }
}

export function buildEventFormStateFromDetail(
  detail: EventDetailResponse,
): EventFormState {
  const base = createDefaultEventFormState()
  const address = detail.address
  const preserveCalendarDate = isAllDayEventType(detail.event_type)
  const startDate = toInputDate(detail.start_at, { preserveCalendarDate })
  const endDate = toInputDate(detail.end_at, { preserveCalendarDate })
  const startTime = toInputTime(detail.start_at)
  const endTime = toInputTime(detail.end_at)

  return {
    ...base,
    title: detail.title,
    showTitle: detail.show_title,
    categoriesText: detail.categories.join(', '),
    eventType: detail.event_type,
    startDate,
    endDate,
    startTime,
    endTime,
    privacyType: detail.privacy_type,
    audienceMembers: detail.private_audiences.includes('members'),
    audiencePublic: detail.private_audiences.includes('public'),
    published: detail.published,
    requestReview: detail.request_review,
    reviewEmailsText: detail.review_email_list.join(', '),
    teaser: detail.teaser,
    descriptionHtml: detail.description_html,
    contactName: detail.contact_name,
    contactEmail: detail.contact_email,
    contactPhone: detail.contact_phone,
    contactExt: detail.contact_ext,
    contactFax: detail.contact_fax,
    locationChoice: resolveLocationChoice(detail.location_mode, address),
    selectedLocationId:
      detail.location_mode === 'address' && address?.is_saved ? String(address.id) : '',
    locationName: address?.name ?? '',
    addressLine1: address?.address_line_1 ?? '',
    addressLine2: address?.address_line_2 ?? '',
    city: address?.city ?? '',
    provinceState: address?.province_state ?? '',
    postalCode: address?.postal_code ?? '',
    country: address?.country ?? 'Canada',
    saveAddress: address?.is_saved ?? false,
    showDisplayImageWhenViewing: detail.show_display_image_when_viewing,
    galleryId: detail.gallery_id ? String(detail.gallery_id) : '',
    registrationEnabled: detail.registration_enabled,
    registrationStartDate: toInputDate(detail.registration_start_at),
    registrationStartTime: toInputTime(detail.registration_start_at),
    registrationEndDate: toInputDate(detail.registration_end_at),
    registrationEndTime: toInputTime(detail.registration_end_at),
    registrationUrl: detail.registration_url,
    repeatEnabled: detail.repeat_enabled,
    recurrenceType:
      detail.repeat_enabled && detail.recurrence_type
        ? (detail.recurrence_type as EventFormState['recurrenceType'])
        : '',
    recurrenceFrequency:
      detail.repeat_enabled && detail.recurrence_frequency
        ? (detail.recurrence_frequency as EventFormState['recurrenceFrequency'])
        : '',
    recurrenceInterval: String(detail.recurrence_interval || 1),
    recurrenceUntilDate: toInputDate(detail.recurrence_until),
    scheduledOccurrences:
      detail.repeat_enabled && detail.recurrence_type === 'scheduled'
        ? detail.occurrences.map((occurrence) =>
            buildOccurrenceFormFromOccurrence(occurrence, detail.event_type),
          )
        : [],
    displayImageFile: null,
    existingDisplayImage: detail.display_image ?? null,
    attachmentFiles: [],
    existingAttachments: detail.attachments ?? [],
  }
}

export function createOccurrenceFromMain(
  form: Pick<
    EventFormState,
    'eventType' | 'startDate' | 'endDate' | 'startTime' | 'endTime'
  >,
): EventOccurrenceForm {
  return {
    id: nextOccurrenceId(),
    startDate: form.startDate,
    endDate: usesSeparateEndDate(form.eventType) ? form.endDate : '',
    startTime: usesTime(form.eventType) ? form.startTime : '',
    endTime: usesTime(form.eventType) ? form.endTime : '',
  }
}

export function buildLocationSummary(address: EventAddress) {
  return [
    address.name,
    address.address_line_1,
    address.city,
    address.province_state,
    address.country,
  ]
    .filter(Boolean)
    .join(', ')
}

export function usesTime(eventType: EventType) {
  return eventType === 'single_day_partial' || eventType === 'multi_day_partial'
}

export function usesSeparateEndDate(eventType: EventType) {
  return eventType === 'multi_day_all_day' || eventType === 'multi_day_partial'
}

export function validateEventForm(
  form: EventFormState,
  t: TranslateFn,
): EventFormErrors {
  const errors: EventFormErrors = {}

  if (!form.title.trim()) {
    errors.title = t('events.validation.titleRequired')
  }
  if (!splitCommaList(form.categoriesText).length) {
    errors.categoriesText = t('events.validation.categoryRequired')
  }
  if (!form.startDate) {
    errors.startDate = t('events.validation.startDateRequired')
  }
  if (usesTime(form.eventType) && !form.startTime) {
    errors.startTime = t('events.validation.startTimeRequired')
  }

  switch (form.eventType) {
    case 'single_day_partial': {
      if (!form.endTime) {
        errors.endTime = t('events.validation.endTimeRequired')
      }
      const startAt = parseLocalDateTime(form.startDate, form.startTime)
      const endAt = parseLocalDateTime(form.startDate, form.endTime)
      if (startAt && endAt && endAt < startAt) {
        errors.endTime = t('events.validation.endAfterStart')
      }
      break
    }
    case 'multi_day_all_day': {
      if (!form.endDate) {
        errors.endDate = t('events.validation.endDateRequired')
      } else if (form.startDate && form.endDate <= form.startDate) {
        errors.endDate = t('events.validation.multiDayEndDate')
      }
      break
    }
    case 'multi_day_partial': {
      if (!form.endDate) {
        errors.endDate = t('events.validation.endDateRequired')
      }
      if (!form.endTime) {
        errors.endTime = t('events.validation.endTimeRequired')
      }
      const startAt = parseLocalDateTime(form.startDate, form.startTime)
      const endAt = parseLocalDateTime(form.endDate, form.endTime)
      if (form.startDate && form.endDate && form.endDate <= form.startDate) {
        errors.endDate = t('events.validation.multiDayEndDate')
      } else if (startAt && endAt && endAt < startAt) {
        errors.endTime = t('events.validation.endAfterStart')
      }
      break
    }
  }

  if (
    form.privacyType === 'private' &&
    !form.audienceMembers &&
    !form.audiencePublic
  ) {
    errors.privateAudiences = t('events.validation.privateAudienceRequired')
  }

  if (!form.published && form.requestReview) {
    if (!splitCommaList(form.reviewEmailsText).length) {
      errors.reviewEmailsText = t('events.validation.reviewEmailsRequired')
    } else if (
      splitCommaList(form.reviewEmailsText).some((email) => !isValidEmail(email))
    ) {
      errors.reviewEmailsText = t('events.validation.reviewEmailsInvalid')
    }
  }

  if (form.locationChoice === 'saved' && !form.selectedLocationId) {
    errors.selectedLocationId = t('events.validation.savedLocationRequired')
  }

  if (form.locationChoice === 'new') {
    if (!form.locationName.trim()) {
      errors.locationName = t('events.validation.locationNameRequired')
    }
    if (!form.addressLine1.trim()) {
      errors.addressLine1 = t('events.validation.addressLine1Required')
    }
    if (!form.city.trim()) {
      errors.city = t('events.validation.cityRequired')
    }
    if (!form.provinceState.trim()) {
      errors.provinceState = t('events.validation.provinceRequired')
    }
    if (!form.postalCode.trim()) {
      errors.postalCode = t('events.validation.postalCodeRequired')
    }
    if (!form.country.trim()) {
      errors.country = t('events.validation.countryRequired')
    }
  }

  if (form.registrationEnabled) {
    if (!form.registrationStartDate) {
      errors.registrationStartDate = t('events.validation.registrationStartRequired')
    }
    if (!form.registrationStartTime) {
      errors.registrationStartTime = t('events.validation.registrationStartTimeRequired')
    }
    if (!form.registrationEndDate) {
      errors.registrationEndDate = t('events.validation.registrationEndRequired')
    }
    if (!form.registrationEndTime) {
      errors.registrationEndTime = t('events.validation.registrationEndTimeRequired')
    }
    if (!form.registrationUrl.trim()) {
      errors.registrationUrl = t('events.validation.registrationUrlRequired')
    }

    const registrationStart = parseLocalDateTime(
      form.registrationStartDate,
      form.registrationStartTime,
    )
    const registrationEnd = parseLocalDateTime(
      form.registrationEndDate,
      form.registrationEndTime,
    )
    if (registrationStart && registrationEnd && registrationEnd < registrationStart) {
      errors.registrationEndTime = t('events.validation.registrationRange')
    }
  }

  if (form.repeatEnabled) {
    if (!form.recurrenceType) {
      errors.recurrenceType = t('events.validation.repeatTypeRequired')
    }

    if (form.recurrenceType === 'recurring') {
      if (!form.recurrenceFrequency) {
        errors.recurrenceFrequency = t('events.validation.recurrenceFrequencyRequired')
      }
      const interval = Number.parseInt(form.recurrenceInterval, 10)
      if (!Number.isFinite(interval) || interval < 1) {
        errors.recurrenceInterval = t('events.validation.recurrenceIntervalRequired')
      }
    }

    if (form.recurrenceType === 'scheduled') {
      if (!form.scheduledOccurrences.length) {
        errors.scheduledOccurrences = t('events.validation.scheduledOccurrencesRequired')
      }

      form.scheduledOccurrences.forEach((occurrence, index) => {
        const prefix = `scheduledOccurrences.${index}`
        if (!occurrence.startDate) {
          errors[`${prefix}.startDate`] = t('events.validation.occurrenceStartDateRequired')
        }
        if (usesTime(form.eventType) && !occurrence.startTime) {
          errors[`${prefix}.startTime`] = t('events.validation.startTimeRequired')
        }
        if (form.eventType === 'single_day_partial') {
          if (!occurrence.endTime) {
            errors[`${prefix}.endTime`] = t('events.validation.endTimeRequired')
          }
          const startAt = parseLocalDateTime(occurrence.startDate, occurrence.startTime)
          const endAt = parseLocalDateTime(occurrence.startDate, occurrence.endTime)
          if (startAt && endAt && endAt < startAt) {
            errors[`${prefix}.endTime`] = t('events.validation.endAfterStart')
          }
        }
        if (form.eventType === 'multi_day_all_day') {
          if (!occurrence.endDate) {
            errors[`${prefix}.endDate`] = t('events.validation.endDateRequired')
          } else if (
            occurrence.startDate &&
            occurrence.endDate <= occurrence.startDate
          ) {
            errors[`${prefix}.endDate`] = t('events.validation.multiDayEndDate')
          }
        }
        if (form.eventType === 'multi_day_partial') {
          if (!occurrence.endDate) {
            errors[`${prefix}.endDate`] = t('events.validation.endDateRequired')
          }
          if (!occurrence.endTime) {
            errors[`${prefix}.endTime`] = t('events.validation.endTimeRequired')
          }
          const startAt = parseLocalDateTime(
            occurrence.startDate,
            occurrence.startTime,
          )
          const endAt = parseLocalDateTime(occurrence.endDate, occurrence.endTime)
          if (
            occurrence.startDate &&
            occurrence.endDate &&
            occurrence.endDate <= occurrence.startDate
          ) {
            errors[`${prefix}.endDate`] = t('events.validation.multiDayEndDate')
          } else if (startAt && endAt && endAt < startAt) {
            errors[`${prefix}.endTime`] = t('events.validation.endAfterStart')
          }
        }
      })
    }
  }

  return errors
}

export async function buildSaveEventPayload(
  form: EventFormState,
): Promise<SaveEventPayload> {
  const categories = splitCommaList(form.categoriesText)
  const privateAudiences = [
    form.audienceMembers ? 'members' : '',
    form.audiencePublic ? 'public' : '',
  ].filter(Boolean)
  const reviewEmailList = splitCommaList(form.reviewEmailsText)
  const published = form.published
  const startAt = buildMainStartAt(form)
  const endAt = buildMainEndAt(form)
  const registrationStartAt = form.registrationEnabled
    ? toLocalIsoString(form.registrationStartDate, form.registrationStartTime)
    : null
  const registrationEndAt = form.registrationEnabled
    ? toLocalIsoString(form.registrationEndDate, form.registrationEndTime)
    : null
  const recurrenceType = form.repeatEnabled ? form.recurrenceType : ''
  const recurrenceFrequency =
    form.repeatEnabled && form.recurrenceType === 'recurring'
      ? form.recurrenceFrequency
      : ''
  const recurrenceInterval = Math.max(
    1,
    Number.parseInt(form.recurrenceInterval, 10) || 1,
  )
  const recurrenceUntil =
    form.repeatEnabled && form.recurrenceUntilDate
      ? toLocalIsoString(form.recurrenceUntilDate, '23:59')
      : null

  return {
    title: form.title.trim(),
    show_title: form.showTitle,
    categories,
    event_type: form.eventType,
    start_at: startAt,
    end_at: endAt,
    privacy_type: form.privacyType,
    private_audiences: form.privacyType === 'private' ? privateAudiences : [],
    published,
    request_review: published ? false : form.requestReview,
    review_email_list: published || !form.requestReview ? [] : reviewEmailList,
    teaser: form.teaser.trim(),
    description_html: form.descriptionHtml.trim(),
    contact_name: form.contactName.trim(),
    contact_email: form.contactEmail.trim(),
    contact_phone: form.contactPhone.trim(),
    contact_ext: form.contactExt.trim(),
    contact_fax: form.contactFax.trim(),
    location_mode: mapLocationMode(form.locationChoice),
    address: buildAddressInput(form),
    show_display_image_when_viewing: form.showDisplayImageWhenViewing,
    gallery_id: form.galleryId ? Number.parseInt(form.galleryId, 10) : null,
    registration_enabled: form.registrationEnabled,
    registration_start_at: registrationStartAt,
    registration_end_at: registrationEndAt,
    registration_url: form.registrationEnabled ? form.registrationUrl.trim() : '',
    repeat_enabled: form.repeatEnabled,
    recurrence_type: recurrenceType,
    recurrence_frequency: recurrenceFrequency,
    recurrence_interval: form.repeatEnabled ? recurrenceInterval : 1,
    recurrence_until: recurrenceUntil,
    recurrence_rule: undefined,
    occurrences:
      form.repeatEnabled && form.recurrenceType === 'scheduled'
        ? buildOccurrences(form.scheduledOccurrences, form.eventType)
        : [],
    display_image: form.displayImageFile
      ? await fileToUploadInput(form.displayImageFile)
      : undefined,
    attachments: await Promise.all(
      form.attachmentFiles.map((file) => fileToUploadInput(file)),
    ),
  }
}

export function toInputDate(
  value?: string | null,
  options: InputDateOptions = {},
) {
  if (!value) {
    return ''
  }

  if (options.preserveCalendarDate) {
    const matchedDate = value.match(/^(\d{4}-\d{2}-\d{2})/)
    if (matchedDate) {
      return matchedDate[1]
    }
  }

  const date = new Date(value)
  if (Number.isNaN(date.getTime())) {
    return ''
  }

  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate()),
  ].join('-')
}

export function toInputTime(value?: string | null) {
  if (!value) {
    return ''
  }

  const date = new Date(value)
  if (Number.isNaN(date.getTime())) {
    return ''
  }

  return [pad(date.getHours()), pad(date.getMinutes())].join(':')
}

function resolveLocationChoice(
  locationMode: string,
  address?: EventAddress | null,
): EventLocationChoice {
  switch (locationMode) {
    case 'to_be_determined':
      return 'to_be_determined'
    case 'address':
      return address?.is_saved ? 'saved' : 'new'
    default:
      return 'none'
  }
}

function buildOccurrenceFormFromOccurrence(
  occurrence: EventOccurrence,
  eventType: EventType,
): EventOccurrenceForm {
  const preserveCalendarDate = isAllDayEventType(eventType)

  return {
    id: nextOccurrenceId(),
    startDate: toInputDate(occurrence.occurrence_start_at, {
      preserveCalendarDate,
    }),
    endDate: toInputDate(occurrence.occurrence_end_at, {
      preserveCalendarDate,
    }),
    startTime: toInputTime(occurrence.occurrence_start_at),
    endTime: toInputTime(occurrence.occurrence_end_at),
  }
}

function buildMainStartAt(form: EventFormState) {
  return toLocalIsoString(
    form.startDate,
    usesTime(form.eventType) ? form.startTime : '00:00',
  )
}

function buildMainEndAt(form: EventFormState) {
  switch (form.eventType) {
    case 'single_day_all_day':
      return null
    case 'single_day_partial':
      return toLocalIsoString(form.startDate, form.endTime)
    case 'multi_day_all_day':
      return toLocalIsoString(form.endDate, '00:00')
    case 'multi_day_partial':
      return toLocalIsoString(form.endDate, form.endTime)
  }
}

function buildAddressInput(form: EventFormState): EventAddressInput | undefined {
  switch (form.locationChoice) {
    case 'saved':
      return {
        id: Number.parseInt(form.selectedLocationId, 10),
        name: '',
        address_line_1: '',
        address_line_2: '',
        city: '',
        province_state: '',
        postal_code: '',
        country: '',
        is_saved: true,
      }
    case 'new':
      return {
        name: form.locationName.trim(),
        address_line_1: form.addressLine1.trim(),
        address_line_2: form.addressLine2.trim(),
        city: form.city.trim(),
        province_state: form.provinceState.trim(),
        postal_code: form.postalCode.trim(),
        country: form.country.trim(),
        is_saved: form.saveAddress,
      }
    default:
      return undefined
  }
}

function buildOccurrences(
  occurrences: EventOccurrenceForm[],
  eventType: EventType,
): EventOccurrenceInput[] {
  return occurrences.map((occurrence) => ({
    occurrence_start_at: toLocalIsoString(
      occurrence.startDate,
      usesTime(eventType) ? occurrence.startTime : '00:00',
    ),
    occurrence_end_at: buildOccurrenceEndAt(occurrence, eventType),
    occurrence_kind: 'scheduled',
  }))
}

function buildOccurrenceEndAt(
  occurrence: EventOccurrenceForm,
  eventType: EventType,
) {
  switch (eventType) {
    case 'single_day_all_day':
      return null
    case 'single_day_partial':
      return toLocalIsoString(occurrence.startDate, occurrence.endTime)
    case 'multi_day_all_day':
      return toLocalIsoString(occurrence.endDate, '00:00')
    case 'multi_day_partial':
      return toLocalIsoString(occurrence.endDate, occurrence.endTime)
  }
}

async function fileToUploadInput(file: File) {
  return {
    display_name: file.name,
    file_name: file.name,
    mime_type: file.type || 'application/octet-stream',
    data_base64: await fileToBase64(file),
  }
}

function fileToBase64(file: File) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onerror = () => reject(new Error('Could not read file'))
    reader.onload = () => {
      const result = reader.result
      if (typeof result !== 'string') {
        reject(new Error('Could not read file'))
        return
      }
      const [, base64 = ''] = result.split(',')
      resolve(base64)
    }
    reader.readAsDataURL(file)
  })
}

function mapLocationMode(choice: EventLocationChoice) {
  switch (choice) {
    case 'saved':
    case 'new':
      return 'address' as const
    case 'to_be_determined':
      return 'to_be_determined' as const
    default:
      return 'none' as const
  }
}

function splitCommaList(value: string) {
  return value
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean)
}

function parseLocalDateTime(dateValue: string, timeValue: string) {
  if (!dateValue || !timeValue) {
    return null
  }

  const [year, month, day] = dateValue.split('-').map(Number)
  const [hours, minutes] = timeValue.split(':').map(Number)
  if (!year || !month || !day || !Number.isFinite(hours) || !Number.isFinite(minutes)) {
    return null
  }

  return new Date(year, month - 1, day, hours, minutes, 0, 0)
}

function toLocalIsoString(dateValue: string, timeValue: string) {
  const [year, month, day] = dateValue.split('-').map(Number)
  const [hours, minutes] = timeValue.split(':').map(Number)
  const date = new Date(year, month - 1, day, hours, minutes, 0, 0)
  const offsetMinutes = -date.getTimezoneOffset()
  const sign = offsetMinutes >= 0 ? '+' : '-'
  const absoluteOffset = Math.abs(offsetMinutes)
  const offsetHours = Math.floor(absoluteOffset / 60)
  const offsetRemainder = absoluteOffset % 60

  return `${dateValue}T${pad(hours)}:${pad(minutes)}:00${sign}${pad(offsetHours)}:${pad(offsetRemainder)}`
}

function isValidEmail(value: string) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)
}

function pad(value: number) {
  return String(value).padStart(2, '0')
}

function nextOccurrenceId() {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID()
  }

  return `occurrence-${Date.now()}-${Math.round(Math.random() * 100000)}`
}
