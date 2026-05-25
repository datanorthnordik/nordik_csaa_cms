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

  it('rejects invalid contact details and registration url values', () => {
    const form = createDefaultEventFormState()
    form.title = 'Spring Gathering'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.contactEmail = 'not-an-email'
    form.contactPhone = 'call me maybe'
    form.contactExt = '12A'
    form.contactFax = 'fax machine'
    form.registrationEnabled = true
    form.registrationStartDate = '2026-05-07'
    form.registrationStartTime = '09:00'
    form.registrationEndDate = '2026-05-07'
    form.registrationEndTime = '10:00'
    form.registrationUrl = 'not-a-url'

    const errors = validateEventForm(form, (key) => key)

    expect(errors.contactEmail).toBe('events.validation.contactEmailInvalid')
    expect(errors.contactPhone).toBe('events.validation.contactPhoneInvalid')
    expect(errors.contactExt).toBe('events.validation.contactExtInvalid')
    expect(errors.contactFax).toBe('events.validation.contactFaxInvalid')
    expect(errors.registrationUrl).toBe('events.validation.registrationUrlInvalid')
  })

  it('allows valid contact details and registration url values', () => {
    const form = createDefaultEventFormState()
    form.title = 'Spring Gathering'
    form.categoriesText = 'Community'
    form.startDate = '2026-05-07'
    form.contactEmail = 'contact@example.com'
    form.contactPhone = '+1 (705) 555-1234'
    form.contactExt = '204'
    form.contactFax = '705-555-9999'
    form.registrationEnabled = true
    form.registrationStartDate = '2026-05-07'
    form.registrationStartTime = '09:00'
    form.registrationEndDate = '2026-05-07'
    form.registrationEndTime = '10:00'
    form.registrationUrl = 'https://example.com/register'

    const errors = validateEventForm(form, (key) => key)

    expect(errors.contactEmail).toBeUndefined()
    expect(errors.contactPhone).toBeUndefined()
    expect(errors.contactExt).toBeUndefined()
    expect(errors.contactFax).toBeUndefined()
    expect(errors.registrationUrl).toBeUndefined()
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
})
