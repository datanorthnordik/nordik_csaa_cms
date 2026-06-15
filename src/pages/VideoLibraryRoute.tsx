import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { Trans, useTranslation } from 'react-i18next'
import { useNavigate } from 'react-router-dom'
import { getApiErrorMessage } from '../api/apiError'
import { type VideoPackageSummary, videoApi } from '../api/videoApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { ConfirmDialog } from '../components/cms/ConfirmDialog'
import { AddIcon, DeleteIcon, EditIcon, SearchIcon, VideoIcon } from '../components/icons'
import { Loader } from '../components/Loader'
import styles from '../styles/VideoLibraryPage.module.css'

export function VideoLibraryRoute() {
  const { t } = useTranslation()
  const navigate = useNavigate()
  const [packages, setPackages] = useState<VideoPackageSummary[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [searchTerm, setSearchTerm] = useState('')
  const [deleteCandidate, setDeleteCandidate] = useState<VideoPackageSummary | null>(null)
  const [deleting, setDeleting] = useState(false)
  const [previewUrls, setPreviewUrls] = useState<Record<number, string>>({})

  async function loadPackages() {
    setLoading(true)
    setError(null)

    try {
      const items = await videoApi.listVideoPackages()
      setPackages(items)
    } catch (loadError) {
      setPackages([])
      setError(getApiErrorMessage(loadError))
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    void loadPackages()
  }, [])

  useEffect(() => {
    const previewablePackages = packages.filter((item) => item.frontImagePath)
    if (!previewablePackages.length) {
      setPreviewUrls((current) => {
        Object.values(current).forEach((url) => URL.revokeObjectURL(url))
        return {}
      })
      return
    }

    let cancelled = false

    async function loadPreviews() {
      const entries = await Promise.all(
        previewablePackages.map(async (item) => {
          if (!item.frontImagePath) {
            return null
          }

          try {
            const blob = await videoApi.fetchVideoTeaserContent(item.frontImagePath)
            return [item.id, URL.createObjectURL(blob)] as const
          } catch {
            return null
          }
        }),
      )

      const nextUrls = entries.reduce<Record<number, string>>((result, entry) => {
        if (entry) {
          result[entry[0]] = entry[1]
        }
        return result
      }, {})

      if (cancelled) {
        Object.values(nextUrls).forEach((url) => URL.revokeObjectURL(url))
        return
      }

      setPreviewUrls((current) => {
        Object.values(current).forEach((url) => URL.revokeObjectURL(url))
        return nextUrls
      })
    }

    void loadPreviews()

    return () => {
      cancelled = true
    }
  }, [packages])

  const filteredPackages = useMemo(() => {
    const trimmed = searchTerm.trim().toLowerCase()
    if (!trimmed) {
      return packages
    }

    return packages.filter((item) => item.title.toLowerCase().includes(trimmed))
  }, [packages, searchTerm])

  async function handleDeleteConfirm() {
    if (!deleteCandidate) {
      return
    }

    setDeleting(true)
    try {
      const result = await videoApi.deleteVideoPackage(deleteCandidate.id)
      toast.success(result.message)
      setDeleteCandidate(null)
      await loadPackages()
    } catch (deleteError) {
      setError(getApiErrorMessage(deleteError))
    } finally {
      setDeleting(false)
    }
  }

  return (
    <CmsAppShell activeKey="videos">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: t('videos.breadcrumb.library') },
          ]}
        />

        <header className={styles.pageHeader}>
          <div>
            <h1 className={styles.title}>{t('videos.library.title')}</h1>
            <p className={styles.subtitle}>{t('videos.library.subtitle')}</p>
          </div>
          <button
            type="button"
            className={styles.primaryButton}
            onClick={() => navigate('/videos/new')}
          >
            <AddIcon size={16} />
            {t('videos.library.create')}
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
            placeholder={t('videos.library.searchPlaceholder')}
            aria-label={t('videos.library.searchPlaceholder')}
          />
        </div>

        {error ? <p className={styles.errorText}>{error}</p> : null}

        {loading ? (
          <div className={styles.loaderWrap}>
            <Loader label={t('videos.library.loading')} />
          </div>
        ) : filteredPackages.length === 0 ? (
          <div className={styles.emptyState}>
            <p className={styles.emptyTitle}>
              {searchTerm
                ? t('videos.library.empty.noMatchTitle')
                : t('videos.library.empty.title')}
            </p>
            <p className={styles.emptyText}>
              {searchTerm
                ? t('videos.library.empty.noMatchText', { term: searchTerm })
                : t('videos.library.empty.text')}
            </p>
          </div>
        ) : (
          <div className={styles.grid}>
            {filteredPackages.map((item) => (
              <article key={item.id} className={styles.card}>
                <div className={styles.cardMedia}>
                  {previewUrls[item.id] ? (
                    <img
                      src={previewUrls[item.id]}
                      alt=""
                      className={styles.cardImage}
                    />
                  ) : (
                    <div className={styles.cardPlaceholder}>
                      <VideoIcon size={22} />
                    </div>
                  )}
                </div>

                <div className={styles.cardBody}>
                  <div className={styles.cardMetaRow}>
                    <span className={styles.typeBadge}>
                      {t(`videos.types.${item.packageType}`)}
                    </span>
                    <span className={styles.countText}>
                      {t('videos.library.card.videoCount', { count: item.videoCount })}
                    </span>
                  </div>

                  <h2 className={styles.cardTitle}>{item.title}</h2>
                  <p className={styles.cardUpdated}>
                    {t('videos.library.card.updatedAt', {
                      date: new Date(item.updatedAt).toLocaleDateString(),
                    })}
                  </p>

                  <div className={styles.cardActions}>
                    <button
                      type="button"
                      className={styles.secondaryButton}
                      onClick={() => navigate(`/videos/${item.id}`)}
                    >
                      <EditIcon size={14} />
                      {t('videos.library.card.manage')}
                    </button>
                    <button
                      type="button"
                      className={styles.dangerButton}
                      onClick={() => setDeleteCandidate(item)}
                    >
                      <DeleteIcon size={14} />
                      {t('videos.library.card.delete')}
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
        title={t('videos.library.delete.title')}
        body={
          <Trans
            i18nKey="videos.library.delete.description"
            values={{ title: deleteCandidate?.title ?? '' }}
            components={{ packageName: <strong /> }}
          />
        }
        confirmLabel={t('videos.library.delete.confirm')}
        destructive
        busy={deleting}
        onConfirm={() => void handleDeleteConfirm()}
        onClose={() => setDeleteCandidate(null)}
      />
    </CmsAppShell>
  )
}
