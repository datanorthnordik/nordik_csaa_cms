import { describe, expect, it } from 'vitest'
import type { EventDetailResponse, EventType } from '../api/eventsApi'
import {
  buildEventFormStateFromDetail,
  buildSaveEventRequest,
  createDefaultEventFormState,
  toInputDate,
  validateEventForm,
} from './eventsForm'

function createEventDetail(
  eventType: EventType,
  overrides: Partial<EventDetailResponse> = {},
): EventDetailResponse {
  return {
    id: 1,
    title: 'Spring Gathering',
    show_title: true,
    categories: ['Community'],
    event_type: eventType,
    start_at: '2026-05-07T00:00:00Z',
    end_at: '2026-05-09T00:00:00Z',
    privacy_type: 'public',
    private_audiences: [],
    published: false,
    request_review: false,
    review_email_list: [],
    teaser: 'Teaser',
    description_html: '<p>Description</p>',
    contact_name: '',
    contact_email: '',
    contact_phone: '',
    contact_ext: '',
    contact_fax: '',
    location_mode: 'none',
    address: null,
    show_display_image_when_viewing: true,
    gallery_id: null,
    registration_enabled: false,
    registration_start_at: null,
    registration_end_at: null,
    registration_url: '',
    repeat_enabled: false,
    recurrence_type: '',
    recurrence_frequency: '',
    recurrence_interval: 1,
    recurrence_until: null,
    recurrence_rule: undefined,
    occurrences: [],
    display_image: null,
    attachments: [],
    created_by: null,
    created_at: '2026-05-01T12:00:00Z',
    updated_at: '2026-05-02T12:00:00Z',
    ...overrides,
  }
}

describe('toInputDate', () => {
  it('preserves the API calendar date for all-day timestamps', () => {
    expect(
      toInputDate('2026-05-07T00:00:00Z', { preserveCalendarDate: true }),
    ).toBe('2026-05-07')
  })

  it('returns empty string for null or undefined input', () => {
    expect(toInputDate(null)).toBe('')
    expect(toInputDate(undefined)).toBe('')
  })

  it('returns empty string for an invalid date string', () => {
    expect(toInputDate('not-a-date')).toBe('')
  })
})

describe('buildEventFormStateFromDetail', () => {
  it('keeps all-day event dates stable when hydrating the edit form', () => {
    const detail = createEventDetail('multi_day_all_day', {
      repeat_enabled: true,
      recurrence_type: 'scheduled',
      occurrences: [
        {
          id: 10,
          event_id: 1,
          occurrence_start_at: '2026-05-12T00:00:00Z',
          occurrence_end_at: '2026-05-14T00:00:00Z',
          occurrence_kind: 'scheduled',
          created_at: '2026-05-01T12:00:00Z',
          updated_at: '2026-05-02T12:00:00Z',
        },
      ],
    })

    const form = buildEventFormStateFromDetail(detail)

    expect(form.startDate).toBe('2026-05-07')
    expect(form.endDate).toBe('2026-05-09')
    expect(form.scheduledOccurrences[0]?.startDate).toBe('2026-05-12')
    expect(form.scheduledOccurrences[0]?.endDate).toBe('2026-05-14')
  })

  it('maps private audiences correctly', () => {
    const detail = createEventDetail('single_day_all_day', {
      privacy_type: 'private',
      private_audiences: ['members', 'public'],
    })

    const form = buildEventFormStateFromDetail(detail)

    expect(form.privacyType).toBe('private')
    expect(form.audienceMembers).toBe(true)
    expect(form.audiencePublic).toBe(true)
  })

  it('sets galleryId as a string from a numeric gallery_id', () => {
    const detail = createEventDetail('single_day_all_day', { gallery_id: 42 })
    const form = buildEventFormStateFromDetail(detail)

    expect(form.galleryId).toBe('42')
  })

  it('sets galleryId to empty string when gallery_id is null', () => {
    const detail = createEventDetail('single_day_all_day', { gallery_id: null })
    const form = buildEventFormStateFromDetail(detail)

    expect(form.galleryId).toBe('')
  })

  it('populates review email text from the list', () => {
    const detail = createEventDetail('single_day_all_day', {
      request_review: true,
      review_email_list: ['a@test.com', 'b@test.com'],
    })
    const form = buildEventFormStateFromDetail(detail)

    expect(form.reviewEmailsText).toBe('a@test.com, b@test.com')
  })
})

describe('validateEventForm', () => {
  it('does not require teaser text when the required fields are present', () => {
    const form = createDefaultEventFormState()
    form.title = 'Spring Gathering'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.teaser = '   '

    const errors = validateEventForm(form, (key) => key)

    expect(errors.teaser).toBeUndefined()
  })

  it('requires a title', () => {
    const form = createDefaultEventFormState()
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'

    const errors = validateEventForm(form, (key) => key)

    expect(errors.title).toBeDefined()
  })

  it('requires at least one category', () => {
    const form = createDefaultEventFormState()
    form.title = 'Test Event'
    form.startDate = '2026-05-07'

    const errors = validateEventForm(form, (key) => key)

    expect(errors.categoriesText).toBeDefined()
  })

  it('requires a start date', () => {
    const form = createDefaultEventFormState()
    form.title = 'Test Event'
    form.categoriesText = 'Community'

    const errors = validateEventForm(form, (key) => key)

    expect(errors.startDate).toBeDefined()
  })

  it('requires an end date for multi_day_all_day events', () => {
    const form = createDefaultEventFormState()
    form.title = 'Test'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.eventType = 'multi_day_all_day'

    const errors = validateEventForm(form, (key) => key)

    expect(errors.endDate).toBeDefined()
  })

  it('requires end date to be after start date for multi_day_all_day events', () => {
    const form = createDefaultEventFormState()
    form.title = 'Test'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-10'
    form.endDate = '2026-05-07'
    form.eventType = 'multi_day_all_day'

    const errors = validateEventForm(form, (key) => key)

    expect(errors.endDate).toBeDefined()
  })

  it('requires start and end times for single_day_partial events', () => {
    const form = createDefaultEventFormState()
    form.title = 'Workshop'
    form.categoriesText = 'Training'
    form.startDate = '2026-05-07'
    form.eventType = 'single_day_partial'

    const errors = validateEventForm(form, (key) => key)

    expect(errors.startTime).toBeDefined()
    expect(errors.endTime).toBeDefined()
  })

  it('requires end time to be after start time for single_day_partial events', () => {
    const form = createDefaultEventFormState()
    form.title = 'Workshop'
    form.categoriesText = 'Training'
    form.startDate = '2026-05-07'
    form.startTime = '14:00'
    form.endTime = '09:00'
    form.eventType = 'single_day_partial'

    const errors = validateEventForm(form, (key) => key)

    expect(errors.endTime).toBeDefined()
  })

  it('requires at least one private audience when privacy is private', () => {
    const form = createDefaultEventFormState()
    form.title = 'Members Only'
    form.categoriesText = 'Internal'
    form.startDate = '2026-05-07'
    form.privacyType = 'private'
    form.audienceMembers = false
    form.audiencePublic = false

    const errors = validateEventForm(form, (key) => key)

    expect(errors.privateAudiences).toBeDefined()
  })

  it('does not require audience selection when privacy is public', () => {
    const form = createDefaultEventFormState()
    form.title = 'Public Event'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.privacyType = 'public'

    const errors = validateEventForm(form, (key) => key)

    expect(errors.privateAudiences).toBeUndefined()
  })

  it('requires valid review emails when request_review is enabled', () => {
    const form = createDefaultEventFormState()
    form.title = 'Draft Event'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.published = false
    form.requestReview = true
    form.reviewEmailsText = 'not-an-email'

    const errors = validateEventForm(form, (key) => key)

    expect(errors.reviewEmailsText).toBeDefined()
  })

  it('requires review emails to be provided when request_review is enabled', () => {
    const form = createDefaultEventFormState()
    form.title = 'Draft Event'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.published = false
    form.requestReview = true
    form.reviewEmailsText = ''

    const errors = validateEventForm(form, (key) => key)

    expect(errors.reviewEmailsText).toBeDefined()
  })

  it('does not validate review emails when the event is already published', () => {
    const form = createDefaultEventFormState()
    form.title = 'Published Event'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.published = true
    form.requestReview = true
    form.reviewEmailsText = ''

    const errors = validateEventForm(form, (key) => key)

    expect(errors.reviewEmailsText).toBeUndefined()
  })

  it('requires a recurrence type when repeat is enabled', () => {
    const form = createDefaultEventFormState()
    form.title = 'Recurring'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.repeatEnabled = true
    form.recurrenceType = ''

    const errors = validateEventForm(form, (key) => key)

    expect(errors.recurrenceType).toBeDefined()
  })

  it('requires recurrence frequency and interval for recurring events', () => {
    const form = createDefaultEventFormState()
    form.title = 'Weekly'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.repeatEnabled = true
    form.recurrenceType = 'recurring'
    form.recurrenceFrequency = ''
    form.recurrenceInterval = '0'

    const errors = validateEventForm(form, (key) => key)

    expect(errors.recurrenceFrequency).toBeDefined()
    expect(errors.recurrenceInterval).toBeDefined()
  })

  it('requires at least one scheduled occurrence for scheduled recurrence', () => {
    const form = createDefaultEventFormState()
    form.title = 'Scheduled'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.repeatEnabled = true
    form.recurrenceType = 'scheduled'
    form.scheduledOccurrences = []

    const errors = validateEventForm(form, (key) => key)

    expect(errors.scheduledOccurrences).toBeDefined()
  })

  it('requires all registration fields when registration is enabled', () => {
    const form = createDefaultEventFormState()
    form.title = 'Ticketed Event'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.registrationEnabled = true

    const errors = validateEventForm(form, (key) => key)

    expect(errors.registrationStartDate).toBeDefined()
    expect(errors.registrationStartTime).toBeDefined()
    expect(errors.registrationEndDate).toBeDefined()
    expect(errors.registrationEndTime).toBeDefined()
    expect(errors.registrationUrl).toBeDefined()
  })

  it('requires a saved location id when location choice is saved', () => {
    const form = createDefaultEventFormState()
    form.title = 'Location Test'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.locationChoice = 'saved'
    form.selectedLocationId = ''

    const errors = validateEventForm(form, (key) => key)

    expect(errors.selectedLocationId).toBeDefined()
  })

  it('requires address fields when location choice is new', () => {
    const form = createDefaultEventFormState()
    form.title = 'New Location'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.locationChoice = 'new'
    form.country = ''

    const errors = validateEventForm(form, (key) => key)

    expect(errors.locationName).toBeDefined()
    expect(errors.addressLine1).toBeDefined()
    expect(errors.city).toBeDefined()
    expect(errors.provinceState).toBeDefined()
    expect(errors.postalCode).toBeDefined()
    expect(errors.country).toBeDefined()
  })
})

describe('buildSaveEventRequest', () => {
  it('returns multipart-ready event media metadata without base64 content', () => {
    const displayImageFile = new File(['poster'], 'poster.png', {
      type: 'image/png',
    })
    const attachmentFile = new File(['agenda'], 'agenda.pdf', {
      type: 'application/pdf',
    })
    const form = createDefaultEventFormState()

    form.title = 'Spring Fair'
    form.categoriesText = 'Events'
    form.startDate = '2026-05-01'
    form.displayImageFile = displayImageFile
    form.attachmentFiles = [attachmentFile]

    const request = buildSaveEventRequest(form)

    expect(request.displayImageFile).toBe(displayImageFile)
    expect(request.attachmentFiles).toEqual([attachmentFile])
    expect(request.display_image).toEqual({
      display_name: 'poster.png',
      file_name: 'poster.png',
      mime_type: 'image/png',
    })
    expect(request.attachments).toEqual([
      {
        display_name: 'agenda.pdf',
        file_name: 'agenda.pdf',
        mime_type: 'application/pdf',
      },
    ])
    expect(request.display_image).not.toHaveProperty('data_base64')
    expect(request.attachments[0]).not.toHaveProperty('data_base64')
  })

  it('omits displayImageFile and attachmentFiles when no files are selected', () => {
    const form = createDefaultEventFormState()
    form.title = 'No Files'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-01'

    const request = buildSaveEventRequest(form)

    expect(request.displayImageFile).toBeUndefined()
    expect(request.attachmentFiles).toBeUndefined()
    expect(request.display_image).toBeUndefined()
    expect(request.attachments).toEqual([])
  })

  it('strips review emails and sets request_review to false when published', () => {
    const form = createDefaultEventFormState()
    form.title = 'Published'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-01'
    form.published = true
    form.requestReview = true
    form.reviewEmailsText = 'editor@example.com'

    const request = buildSaveEventRequest(form)

    expect(request.request_review).toBe(false)
    expect(request.review_email_list).toEqual([])
  })

  it('clears private audiences when privacy is public', () => {
    const form = createDefaultEventFormState()
    form.title = 'Open'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-01'
    form.privacyType = 'public'
    form.audienceMembers = true

    const request = buildSaveEventRequest(form)

    expect(request.private_audiences).toEqual([])
  })
})
