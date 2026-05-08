import { describe, expect, it } from 'vitest'
import type { EventDetailResponse, EventType } from '../api/eventsApi'
import {
  buildEventFormStateFromDetail,
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
})
