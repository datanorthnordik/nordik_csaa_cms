import IconButton from '@mui/material/IconButton'
import Menu from '@mui/material/Menu'
import MenuItem from '@mui/material/MenuItem'
import MoreVertIcon from '@mui/icons-material/MoreVert'
import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { Trans, useTranslation } from 'react-i18next'
import { useNavigate } from 'react-router-dom'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { AddIcon } from '../components/icons'
import { ConfirmDialog } from '../components/cms/ConfirmDialog'
import { PaginationControls } from '../components/cms/PaginationControls'
import { SearchFilterBar } from '../components/cms/SearchFilterBar'
import { StatusBadge } from '../components/cms/StatusBadge'
import {
  defaultEventListFilters,
  deleteEvent as deleteEventAction,
  fetchEvents,
  selectEventList,
  selectEventListFilters,
  setEventListFilters,
} from '../store/eventsSlice'
import { useAppDispatch, useAppSelector } from '../store/hooks'
import type {
  EventDateRange,
  EventListFilters,
  EventListItem,
  EventSortBy,
  EventSortOrder,
  EventStatus,
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

const statusOptions: EventStatus[] = ['published', 'draft']

export function EventsListPage() {
  const dispatch = useAppDispatch()
  const navigate = useNavigate()
  const { i18n, t } = useTranslation()
  const filters = useAppSelector(selectEventListFilters)
  const { items, pagination, status, error } = useAppSelector(selectEventList)
  const [draftFilters, setDraftFilters] = useState<EventListFilters>(filters)
  const [actionMenuAnchorEl, setActionMenuAnchorEl] = useState<HTMLElement | null>(null)
  const [activeActionItem, setActiveActionItem] = useState<EventListItem | null>(null)
  const [deleteCandidate, setDeleteCandidate] = useState<EventListItem | null>(null)
  const [deletingEventId, setDeletingEventId] = useState<number | null>(null)

  useEffect(() => {
    setDraftFilters(filters)
  }, [filters])

  useEffect(() => {
    void dispatch(fetchEvents(filters))
  }, [dispatch, filters])

  useEffect(() => {
    if (activeActionItem && !items.some((item) => item.id === activeActionItem.id)) {
      closeActionMenu()
    }
  }, [activeActionItem, items])

  useEffect(() => {
    if (deleteCandidate && !items.some((item) => item.id === deleteCandidate.id)) {
      closeDeleteDialog()
    }
  }, [deleteCandidate, items])

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

  function openActionMenu(
    event: React.MouseEvent<HTMLButtonElement>,
    item: EventListItem,
  ) {
    setActionMenuAnchorEl(event.currentTarget)
    setActiveActionItem(item)
  }

  function closeActionMenu() {
    setActionMenuAnchorEl(null)
    setActiveActionItem(null)
  }

  function openDeleteDialog(item: EventListItem) {
    closeActionMenu()
    setDeleteCandidate(item)
  }

  function closeDeleteDialog() {
    if (deleteCandidate && deletingEventId === deleteCandidate.id) {
      return
    }

    setDeleteCandidate(null)
  }

  function formatEventDate(item: EventListItem) {
    if (item.date_display?.trim()) {
      return item.date_display
    }

    return formatEventStartDateLabel(item.start_at, item.event_type, {
      localFormatter: formatter,
      calendarFormatter,
    })
  }

  async function handleDeleteEvent(item: EventListItem) {
    setDeletingEventId(item.id)

    try {
      await dispatch(deleteEventAction(item.id)).unwrap()
      setDeleteCandidate(null)

      const nextPage =
        pagination && pagination.page > 1 && items.length === 1
          ? pagination.page - 1
          : filters.page

      dispatch(setEventListFilters({ page: nextPage }))
      toast.success(t('events.feedback.deleted'))
    } catch {
      return
    } finally {
      setDeletingEventId(null)
    }
  }

  function renderEventActions(item: EventListItem) {
    return (
      <span className={styles.actionMenu}>
        <IconButton
          id={`event-actions-trigger-${item.id}`}
          size="small"
          className={styles.actionMenuTrigger}
          aria-controls={
            activeActionItem?.id === item.id && actionMenuAnchorEl
              ? 'event-actions-menu'
              : undefined
          }
          aria-expanded={activeActionItem?.id === item.id && actionMenuAnchorEl ? true : undefined}
          aria-haspopup="menu"
          aria-label={t('events.list.openActions')}
          disabled={deletingEventId === item.id}
          onClick={(event) => openActionMenu(event, item)}
        >
          <MoreVertIcon fontSize="small" />
        </IconButton>
      </span>
    )
  }

  const totalItems = pagination?.total_items ?? 0
  const isActionMenuOpen = Boolean(actionMenuAnchorEl && activeActionItem)
  const isDeleteDialogOpen = Boolean(deleteCandidate)
  const isDeleteDialogBusy = Boolean(
    deleteCandidate && deletingEventId === deleteCandidate.id,
  )
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
        <Breadcrumb
          items={[
            { label: t('events.breadcrumb.events'), to: '/events' },
            { label: t('events.breadcrumb.list') },
          ]}
        />

        <div className={styles.pageHeader}>
          <div>
            <h1 className={styles.title}>{t('events.list.title')}</h1>
            <p className={styles.subtitle}>{t('events.list.subtitle')}</p>
          </div>
          <button
            type="button"
            className={styles.primaryButton}
            onClick={() => navigate('/events/new')}
          >
            <AddIcon size={16} />
            {t('events.list.create')}
          </button>
        </div>

        <SearchFilterBar
          searchValue={draftFilters.searchTerm}
          onSearchChange={(value) => updateDraft('searchTerm', value)}
          searchPlaceholder={t('events.filters.searchPlaceholder')}
          searchLabel={t('events.filters.search')}
          applyLabel={t('events.filters.apply')}
          resetLabel={t('events.filters.reset')}
          onApply={() => applyFilters()}
          onReset={resetFilters}
          fields={[
            {
              type: 'select',
              key: 'dateRange',
              label: t('events.filters.dateRange'),
              value: draftFilters.dateRange,
              onChange: (value) => updateDraft('dateRange', value as EventDateRange),
              options: dateRangeOptions.map((option) => ({
                value: option,
                label: t(`events.dateRanges.${option}`),
              })),
            },
            {
              type: 'date',
              key: 'startDate',
              label: t('events.filters.startDate'),
              value: draftFilters.startDate,
              disabled: draftFilters.dateRange !== 'custom',
              onChange: (value) => updateDraft('startDate', value),
            },
            {
              type: 'date',
              key: 'endDate',
              label: t('events.filters.endDate'),
              value: draftFilters.endDate,
              disabled: draftFilters.dateRange !== 'custom',
              onChange: (value) => updateDraft('endDate', value),
            },
            {
              type: 'select',
              key: 'sortBy',
              label: t('events.filters.sortBy'),
              value: draftFilters.sortBy,
              onChange: (value) => updateDraft('sortBy', value as EventSortBy),
              options: sortByOptions.map((option) => ({
                value: option,
                label: t(`events.sortBy.${option}`),
              })),
            },
            {
              type: 'select',
              key: 'sortOrder',
              label: t('events.filters.sortOrder'),
              value: draftFilters.sortOrder,
              onChange: (value) =>
                updateDraft('sortOrder', value as EventSortOrder),
              options: sortOrderOptions.map((option) => ({
                value: option,
                label: t(`events.sortOrder.${option}`),
              })),
            },
            {
              type: 'multi-pills',
              key: 'statuses',
              label: t('events.filters.status'),
              values: draftFilters.statuses,
              onToggle: (value) => toggleStatus(value as EventStatus),
              options: statusOptions.map((option) => ({
                value: option,
                label: t(`events.status.${option}`),
              })),
            },
          ]}
        />

        <section className={styles.resultsPanel}>
          {totalItems > 0 && (
            <div className={styles.resultsHeader}>
              <span className={styles.resultsLabel}>
                {t('events.list.resultsLabel', {
                  start: rangeStart,
                  end: rangeEnd,
                  total: totalItems,
                })}
              </span>
              {pagination && pagination.total_pages > 1 && (
                <PaginationControls
                  page={pagination.page}
                  totalPages={pagination.total_pages}
                  onChange={changePage}
                />
              )}
            </div>
          )}
          {error && <p className={styles.errorText}>{error}</p>}

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
                        <td>{item.categories.join(', ') || t('events.list.noCategory')}</td>
                        <td>
                          <StatusBadge
                            status={item.status}
                            label={t(`events.status.${item.status}`)}
                          />
                        </td>
                        <td>{formatEventDate(item)}</td>
                        <td className={styles.actionsCell}>{renderEventActions(item)}</td>
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
                          {formatEventDate(item)}
                        </p>
                      </div>
                      <StatusBadge
                        status={item.status}
                        label={t(`events.status.${item.status}`)}
                      />
                    </div>
                    <p className={styles.cardMeta}>
                      {item.categories.join(', ') || t('events.list.noCategory')}
                    </p>
                    <div className={styles.cardActionRow}>{renderEventActions(item)}</div>
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

        </section>

        <Menu
          id="event-actions-menu"
          anchorEl={actionMenuAnchorEl}
          open={isActionMenuOpen}
          onClose={closeActionMenu}
          anchorOrigin={{
            vertical: 'bottom',
            horizontal: 'right',
          }}
          transformOrigin={{
            vertical: 'top',
            horizontal: 'right',
          }}
          slotProps={{
            paper: {
              className: styles.actionMenuPaper,
            },
            list: {
              'aria-labelledby': activeActionItem
                ? `event-actions-trigger-${activeActionItem.id}`
                : undefined,
              className: styles.actionMenuList,
            },
          }}
        >
          <MenuItem
            className={styles.actionMenuItem}
            onClick={() => {
              if (!activeActionItem) {
                return
              }

              const item = activeActionItem
              closeActionMenu()
              navigate(`/events/${item.id}`)
            }}
          >
            {t('events.list.view')}
          </MenuItem>
          <MenuItem
            className={styles.actionMenuItem}
            onClick={() => {
              if (!activeActionItem) {
                return
              }

              const item = activeActionItem
              closeActionMenu()
              navigate(`/events/${item.id}/edit`)
            }}
          >
            {t('events.list.edit')}
          </MenuItem>
          <MenuItem
            className={`${styles.actionMenuItem} ${styles.actionMenuDanger}`}
            disabled={activeActionItem ? deletingEventId === activeActionItem.id : false}
            onClick={() => {
              if (!activeActionItem) {
                return
              }

              openDeleteDialog(activeActionItem)
            }}
          >
            {activeActionItem && deletingEventId === activeActionItem.id
              ? t('events.common.loading')
              : t('events.list.delete')}
          </MenuItem>
        </Menu>

        <ConfirmDialog
          open={isDeleteDialogOpen}
          title={t('events.list.deleteDialogTitle')}
          body={
            <Trans
              i18nKey="events.list.deleteDialogDescription"
              values={{ title: deleteCandidate?.title ?? '' }}
              components={{
                eventName: <strong className={styles.confirmDialogEventName} />,
              }}
            />
          }
          confirmLabel={t('events.list.delete')}
          cancelLabel={t('events.list.cancelDelete')}
          destructive
          busy={isDeleteDialogBusy}
          onConfirm={() => {
            if (!deleteCandidate) {
              return
            }
            void handleDeleteEvent(deleteCandidate)
          }}
          onClose={closeDeleteDialog}
        />
      </div>
    </CmsAppShell>
  )
}
