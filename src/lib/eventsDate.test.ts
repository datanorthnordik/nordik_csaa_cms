import { describe, expect, it, vi } from 'vitest'
import { formatEventStartDateLabel } from './eventsDate'

describe('formatEventStartDateLabel', () => {
  it('uses the API calendar date for all-day events', () => {
    const localFormatMock = vi.fn(() => 'local')
    let formattedDate: Date | undefined
    const calendarFormatMock = vi.fn((date: Date) => {
      formattedDate = date
      return 'calendar'
    })
    const localFormatter = {
      format: localFormatMock,
    }
    const calendarFormatter = {
      format: calendarFormatMock,
    }

    const result = formatEventStartDateLabel(
      '2026-05-07T00:00:00Z',
      'single_day_all_day',
      {
        localFormatter,
        calendarFormatter,
      },
    )

    expect(result).toBe('calendar')
    expect(calendarFormatMock).toHaveBeenCalledTimes(1)
    expect(formattedDate?.toISOString()).toBe('2026-05-07T00:00:00.000Z')
    expect(localFormatMock).not.toHaveBeenCalled()
  })

  it('uses the local formatter for timed events', () => {
    const localFormatMock = vi.fn(() => 'local')
    const calendarFormatMock = vi.fn(() => 'calendar')
    const localFormatter = {
      format: localFormatMock,
    }
    const calendarFormatter = {
      format: calendarFormatMock,
    }

    const result = formatEventStartDateLabel(
      '2026-05-07T09:30:00-04:00',
      'single_day_partial',
      {
        localFormatter,
        calendarFormatter,
      },
    )

    expect(result).toBe('local')
    expect(localFormatMock).toHaveBeenCalledTimes(1)
    expect(calendarFormatMock).not.toHaveBeenCalled()
  })
})
