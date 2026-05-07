import type { EventType } from '../api/eventsApi'

type DateFormatter = Pick<Intl.DateTimeFormat, 'format'>

export type EventDateFormatters = {
  localFormatter: DateFormatter
  calendarFormatter: DateFormatter
}

export function isAllDayEventType(eventType: EventType) {
  return eventType === 'single_day_all_day' || eventType === 'multi_day_all_day'
}

export function formatEventStartDateLabel(
  startAt: string,
  eventType: EventType,
  formatters: EventDateFormatters,
) {
  if (isAllDayEventType(eventType)) {
    const calendarDate = parseApiCalendarDate(startAt)
    return calendarDate
      ? formatters.calendarFormatter.format(calendarDate)
      : startAt
  }

  const date = new Date(startAt)
  return Number.isNaN(date.getTime())
    ? startAt
    : formatters.localFormatter.format(date)
}

function parseApiCalendarDate(value: string) {
  const matchedDate = value.match(/^(\d{4})-(\d{2})-(\d{2})/)
  if (!matchedDate) {
    return null
  }

  const year = Number(matchedDate[1])
  const month = Number(matchedDate[2])
  const day = Number(matchedDate[3])
  if (!year || !month || !day) {
    return null
  }

  return new Date(Date.UTC(year, month - 1, day))
}
