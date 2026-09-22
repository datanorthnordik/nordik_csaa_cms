import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { Trans, useTranslation } from 'react-i18next'
import { useNavigate } from 'react-router-dom'
import { getApiErrorMessage } from '../api/apiError'
import {
  type RecordingCollectionSummary,
  recordingsApi,
} from '../api/recordingsApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { ConfirmDialog } from '../components/cms/ConfirmDialog'
import { AddIcon, DeleteIcon, EditIcon, RecordingIcon, SearchIcon } from '../components/icons'
import { Loader } from '../components/Loader'
import styles from '../styles/RecordingsListPage.module.css'

export function RecordingsListPage() {
  const { t } = useTranslation()
  const navigate = useNavigate()
  const [collections, setCollections] = useState<RecordingCollectionSummary[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [searchTerm, setSearchTerm] = useState('')
  const [deleteCandidate, setDeleteCandidate] = useState<RecordingCollectionSummary | null>(null)
  const [deleting, setDeleting] = useState(false)

  async function loadCollections() {
    setLoading(true)
    setError(null)

    try {
      const items = await recordingsApi.listRecordingCollections()
      setCollections(items)
    } catch (loadError) {
      setCollections([])
      setError(getApiErrorMessage(loadError))
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    void loadCollections()
  }, [])

  const filteredCollections = useMemo(() => {
    const trimmed = searchTerm.trim().toLowerCase()
    if (!trimmed) {
      return collections
    }

    return collections.filter((item) => item.name.toLowerCase().includes(trimmed))
  }, [collections, searchTerm])

  async function handleDeleteConfirm() {
    if (!deleteCandidate) {
      return
    }

    setDeleting(true)
    try {
      const result = await recordingsApi.deleteRecordingCollection(deleteCandidate.id)
      toast.success(result.message)
      setDeleteCandidate(null)
      await loadCollections()
    } catch (deleteError) {
      setError(getApiErrorMessage(deleteError))
    } finally {
      setDeleting(false)
    }
  }

  return (
    <CmsAppShell activeKey="recordings">
      <div className={styles.page}>
        <Breadcrumb items={[{ label: t('recordings.breadcrumb.library') }]} />

        <header className={styles.pageHeader}>
          <div>
            <h1 className={styles.title}>{t('recordings.library.title')}</h1>
            <p className={styles.subtitle}>{t('recordings.library.subtitle')}</p>
          </div>
          <button
            type="button"
            className={styles.primaryButton}
            onClick={() => navigate('/recordings/new')}
          >
            <AddIcon size={16} />
            {t('recordings.library.create')}
          </button>
        </header>

        <div className={styles.searchRow}>
          <span className={styles.searchIcon} aria-hidden="true">
            <SearchIcon />
          </span>
          <input
            type="search"
            className={styles.searchInput}
            value={searchTerm}
            onChange={(event) => setSearchTerm(event.target.value)}
            placeholder={t('recordings.library.searchPlaceholder')}
            aria-label={t('recordings.library.searchPlaceholder')}
          />
        </div>

        {error ? <p className={styles.errorText}>{error}</p> : null}

        {loading ? (
          <div className={styles.loaderWrap}>
            <Loader label={t('recordings.library.loading')} />
          </div>
        ) : filteredCollections.length === 0 ? (
          <div className={styles.emptyState}>
            <p className={styles.emptyTitle}>
              {searchTerm
                ? t('recordings.library.empty.noMatchTitle')
                : t('recordings.library.empty.title')}
            </p>
            <p className={styles.emptyText}>
              {searchTerm
                ? t('recordings.library.empty.noMatchText', { term: searchTerm })
                : t('recordings.library.empty.text')}
            </p>
          </div>
        ) : (
          <div className={styles.grid}>
            {filteredCollections.map((item) => (
              <article key={item.id} className={styles.card}>
                <div className={styles.cardBody}>
                  <p className={styles.cardEyebrow}>
                    <RecordingIcon size={14} />
                    {t('recordings.library.card.eyebrow')}
                  </p>
                  <h2 className={styles.cardTitle}>{item.name}</h2>
                  <p className={styles.countText}>
                    {t('recordings.library.card.itemCount', { count: item.itemCount })}
                  </p>
                  <p className={styles.cardUpdated}>
                    {t('recordings.library.card.updatedAt', {
                      date: new Date(item.updatedAt).toLocaleDateString(),
                    })}
                  </p>

                  <div className={styles.cardActions}>
                    <button
                      type="button"
                      className={styles.secondaryButton}
                      onClick={() => navigate(`/recordings/${item.id}`)}
                    >
                      <EditIcon size={14} />
                      {t('recordings.library.card.manage')}
                    </button>
                    <button
                      type="button"
                      className={styles.dangerButton}
                      onClick={() => setDeleteCandidate(item)}
                    >
                      <DeleteIcon size={14} />
                      {t('recordings.library.card.delete')}
                    </button>
                  </div>
                </div>
              </article>
            ))}
          </div>
        )}
      </div>

      <ConfirmDialog
        open={Boolean(deleteCandidate)}
        title={t('recordings.library.delete.title')}
        body={
          <Trans
            i18nKey="recordings.library.delete.description"
            values={{ name: deleteCandidate?.name ?? '' }}
            components={{ collectionName: <strong /> }}
          />
        }
        confirmLabel={t('recordings.library.delete.confirm')}
        destructive
        busy={deleting}
        onConfirm={() => void handleDeleteConfirm()}
        onClose={() => setDeleteCandidate(null)}
      />
    </CmsAppShell>
  )
}
