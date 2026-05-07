import { useEffect, useMemo, useState } from 'react'
import { useTranslation } from 'react-i18next'
import { useNavigate } from 'react-router-dom'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import {
  defaultEventListFilters,
  fetchEvents,
  selectEventList,
  selectEventListFilters,
  setEventListFilters,
} from '../store/eventsSlice'
import { useAppDispatch, useAppSelector } from '../store/hooks'
import type {
  EventDateRange,
  EventListFilters,
  EventSortBy,
  EventSortOrder,
  EventStatus,
  EventType,
} from '../api/eventsApi'
import { formatEventStartDateLabel } from '../lib/eventsDate'
import styles from '../styles/EventsListPage.module.css'

const dateRangeOptions: EventDateRange[] = [
  'custom',
  'next_30_days',
  'last_30_days',
  'today',
  'this_month',
  'upcoming',
]

const sortByOptions: EventSortBy[] = [
  'start_at',
  'title',
  'created_at',
  'updated_at',
  'published',
]

const sortOrderOptions: EventSortOrder[] = ['desc', 'asc']

const pageSizeOptions = [10, 20, 50]
const statusOptions: EventStatus[] = ['published', 'draft']

export function EventsListPage() {
  const dispatch = useAppDispatch()
  const navigate = useNavigate()
  const { i18n, t } = useTranslation()
  const filters = useAppSelector(selectEventListFilters)
  const { items, pagination, status, error } = useAppSelector(selectEventList)
  const [draftFilters, setDraftFilters] = useState<EventListFilters>(filters)

  useEffect(() => {
    setDraftFilters(filters)
  }, [filters])

  useEffect(() => {
    void dispatch(fetchEvents(filters))
  }, [dispatch, filters])

  const isLoading = status === 'loading'
  const formatter = useMemo(
    () =>
      new Intl.DateTimeFormat(i18n.language, {
        month: 'short',
        day: '2-digit',
        year: 'numeric',
      }),
    [i18n.language],
  )
  const calendarFormatter = useMemo(
    () =>
      new Intl.DateTimeFormat(i18n.language, {
        month: 'short',
        day: '2-digit',
        year: 'numeric',
        timeZone: 'UTC',
      }),
    [i18n.language],
  )

  function updateDraft<K extends keyof EventListFilters>(
    key: K,
    value: EventListFilters[K],
  ) {
    setDraftFilters((current) => ({
      ...current,
      [key]: value,
    }))
  }

  function toggleStatus(statusValue: EventStatus) {
    setDraftFilters((current) => {
      const hasStatus = current.statuses.includes(statusValue)
      return {
        ...current,
        statuses: hasStatus
          ? current.statuses.filter((status) => status !== statusValue)
          : [...current.statuses, statusValue],
      }
    })
  }

  function applyFilters(event?: React.FormEvent<HTMLFormElement>) {
    event?.preventDefault()

    dispatch(
      setEventListFilters({
        ...draftFilters,
        page: 1,
        startDate: draftFilters.dateRange === 'custom' ? draftFilters.startDate : '',
        endDate: draftFilters.dateRange === 'custom' ? draftFilters.endDate : '',
      }),
    )
  }

  function resetFilters() {
    setDraftFilters(defaultEventListFilters)
    dispatch(setEventListFilters(defaultEventListFilters))
  }

  function changePage(nextPage: number) {
    if (!pagination || nextPage < 1 || nextPage > pagination.total_pages) {
      return
    }

    dispatch(setEventListFilters({ page: nextPage }))
  }

  function changePageSize(pageSize: number) {
    dispatch(setEventListFilters({ pageSize, page: 1 }))
  }

  function formatEventDate(startAt: string, eventType: EventType) {
    return formatEventStartDateLabel(startAt, eventType, {
      localFormatter: formatter,
      calendarFormatter,
    })
  }

  const totalItems = pagination?.total_items ?? 0
  const rangeStart =
    pagination && totalItems > 0
      ? (pagination.page - 1) * pagination.page_size + 1
      : 0
  const rangeEnd =
    pagination && totalItems > 0
      ? Math.min(pagination.page * pagination.page_size, totalItems)
      : 0

  return (
    <CmsAppShell activeKey="events">
      <div className={styles.page}>
        <div className={styles.pageHeader}>
          <div>
            <p className={styles.eyebrow}>{t('events.list.eyebrow')}</p>
            <h1 className={styles.title}>{t('events.list.title')}</h1>
            <p className={styles.subtitle}>{t('events.list.subtitle')}</p>
          </div>
          <button
            type="button"
            className={styles.primaryButton}
            onClick={() => navigate('/events/new')}
          >
            {t('events.list.create')}
          </button>
        </div>

        <form className={styles.filterPanel} onSubmit={applyFilters}>
          <div className={styles.filterGrid}>
            <label className={styles.field}>
              <span>{t('events.filters.search')}</span>
              <input
                type="search"
                value={draftFilters.searchTerm}
                placeholder={t('events.filters.searchPlaceholder')}
                onChange={(event) => updateDraft('searchTerm', event.target.value)}
              />
            </label>

            <label className={styles.field}>
              <span>{t('events.filters.dateRange')}</span>
              <select
                value={draftFilters.dateRange}
                onChange={(event) =>
                  updateDraft('dateRange', event.target.value as EventDateRange)
                }
              >
                {dateRangeOptions.map((option) => (
                  <option key={option} value={option}>
                    {t(`events.dateRanges.${option}`)}
                  </option>
                ))}
              </select>
            </label>

            <label className={styles.field}>
              <span>{t('events.filters.startDate')}</span>
              <input
                type="date"
                value={draftFilters.startDate}
                disabled={draftFilters.dateRange !== 'custom'}
                onChange={(event) => updateDraft('startDate', event.target.value)}
              />
            </label>

            <label className={styles.field}>
              <span>{t('events.filters.endDate')}</span>
              <input
                type="date"
                value={draftFilters.endDate}
                disabled={draftFilters.dateRange !== 'custom'}
                onChange={(event) => updateDraft('endDate', event.target.value)}
              />
            </label>

            <label className={styles.field}>
              <span>{t('events.filters.sortBy')}</span>
              <select
                value={draftFilters.sortBy}
                onChange={(event) =>
                  updateDraft('sortBy', event.target.value as EventSortBy)
                }
              >
                {sortByOptions.map((option) => (
                  <option key={option} value={option}>
                    {t(`events.sortBy.${option}`)}
                  </option>
                ))}
              </select>
            </label>

            <label className={styles.field}>
              <span>{t('events.filters.sortOrder')}</span>
              <select
                value={draftFilters.sortOrder}
                onChange={(event) =>
                  updateDraft('sortOrder', event.target.value as EventSortOrder)
                }
              >
                {sortOrderOptions.map((option) => (
                  <option key={option} value={option}>
                    {t(`events.sortOrder.${option}`)}
                  </option>
                ))}
              </select>
            </label>
          </div>

          <div className={styles.statusRow}>
            <div className={styles.statusFilters}>
              <span className={styles.statusLabel}>{t('events.filters.status')}</span>
              {statusOptions.map((statusOption) => (
                <label key={statusOption} className={styles.checkboxPill}>
                  <input
                    type="checkbox"
                    checked={draftFilters.statuses.includes(statusOption)}
                    onChange={() => toggleStatus(statusOption)}
                  />
                  <span>{t(`events.status.${statusOption}`)}</span>
                </label>
              ))}
            </div>

            <div className={styles.filterActions}>
              <button type="button" className={styles.secondaryButton} onClick={resetFilters}>
                {t('events.filters.reset')}
              </button>
              <button type="submit" className={styles.primaryButton}>
                {t('events.filters.apply')}
              </button>
            </div>
          </div>
        </form>

        <section className={styles.resultsPanel}>
          <div className={styles.resultsHeader}>
            <div>
              <p className={styles.resultsLabel}>
                {t('events.list.resultsLabel', {
                  start: rangeStart,
                  end: rangeEnd,
                  total: totalItems,
                })}
              </p>
              {error && <p className={styles.errorText}>{error}</p>}
            </div>

            <label className={styles.pageSizeField}>
              <span>{t('events.filters.pageSize')}</span>
              <select
                value={filters.pageSize}
                onChange={(event) => changePageSize(Number.parseInt(event.target.value, 10))}
              >
                {pageSizeOptions.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
            </label>
          </div>

          {isLoading && !items.length ? (
            <div className={styles.loaderWrap}>
              <Loader label={t('events.common.loading')} />
            </div>
          ) : items.length ? (
            <>
              <div className={styles.tableWrap}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>{t('events.list.columns.title')}</th>
                      <th>{t('events.list.columns.categories')}</th>
                      <th>{t('events.list.columns.status')}</th>
                      <th>{t('events.list.columns.date')}</th>
                      <th>{t('events.list.columns.actions')}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((item) => (
                      <tr key={item.id}>
                        <td>
                          <div className={styles.eventTitle}>{item.title}</div>
                        </td>
                        <td>{item.categories.join(', ') || '—'}</td>
                        <td>
                          <span
                            className={[
                              styles.badge,
                              item.status === 'published'
                                ? styles.badgePublished
                                : styles.badgeDraft,
                            ].join(' ')}
                          >
                            {t(`events.status.${item.status}`)}
                          </span>
                        </td>
                        <td>{formatEventDate(item.start_at, item.event_type)}</td>
                        <td>
                          <button
                            type="button"
                            className={styles.linkButton}
                            onClick={() => navigate(`/events/${item.id}/edit`)}
                          >
                            {t('events.list.edit')}
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className={styles.cardList}>
                {items.map((item) => (
                  <article key={item.id} className={styles.card}>
                    <div className={styles.cardHeader}>
                      <div>
                        <p className={styles.cardTitle}>{item.title}</p>
                        <p className={styles.cardMeta}>
                          {formatEventDate(item.start_at, item.event_type)}
                        </p>
                      </div>
                      <span
                        className={[
                          styles.badge,
                          item.status === 'published'
                            ? styles.badgePublished
                            : styles.badgeDraft,
                        ].join(' ')}
                      >
                        {t(`events.status.${item.status}`)}
                      </span>
                    </div>
                    <p className={styles.cardMeta}>
                      {item.categories.join(', ') || t('events.list.noCategory')}
                    </p>
                    <button
                      type="button"
                      className={styles.linkButton}
                      onClick={() => navigate(`/events/${item.id}/edit`)}
                    >
                      {t('events.list.edit')}
                    </button>
                  </article>
                ))}
              </div>
            </>
          ) : (
            <div className={styles.emptyState}>
              <p className={styles.emptyTitle}>{t('events.list.emptyTitle')}</p>
              <p className={styles.emptyText}>{t('events.list.emptyText')}</p>
            </div>
          )}

          {pagination && pagination.total_pages > 1 && (
            <div className={styles.pagination}>
              <button
                type="button"
                className={styles.secondaryButton}
                disabled={!pagination.has_prev}
                onClick={() => changePage(filters.page - 1)}
              >
                {t('events.list.previous')}
              </button>
              <span className={styles.pageNumber}>
                {t('events.list.pageOf', {
                  page: pagination.page,
                  total: pagination.total_pages,
                })}
              </span>
              <button
                type="button"
                className={styles.secondaryButton}
                disabled={!pagination.has_next}
                onClick={() => changePage(filters.page + 1)}
              >
                {t('events.list.next')}
              </button>
            </div>
          )}
        </section>
      </div>
    </CmsAppShell>
  )
}
