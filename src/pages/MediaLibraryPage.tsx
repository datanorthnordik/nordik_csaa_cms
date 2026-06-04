import { useMemo, useState } from 'react'
import { Trans, useTranslation } from 'react-i18next'
import { useNavigate } from 'react-router-dom'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { ConfirmDialog } from '../components/cms/ConfirmDialog'
import { AddIcon, SearchIcon } from '../components/icons'
import { Loader } from '../components/Loader'
import {
  CreateGalleryDialog,
  type CreateGalleryFormValues,
} from '../components/media/CreateGalleryDialog'
import { GalleryCard } from '../components/media/GalleryCard'
import type { GallerySummary } from '../types/media'
import styles from '../styles/MediaLibraryPage.module.css'

type MediaLibraryPageProps = {
  galleries?: GallerySummary[]
  loading?: boolean
  error?: string
  creating?: boolean
  deleting?: boolean
  onCreate?: (
    values: CreateGalleryFormValues,
    frontImage?: File,
  ) => void | boolean | Promise<void | boolean>
  onDelete?: (gallery: GallerySummary) => void | boolean | Promise<void | boolean>
}

export function MediaLibraryPage({
  galleries,
  loading = false,
  error,
  creating = false,
  deleting = false,
  onCreate,
  onDelete,
}: MediaLibraryPageProps = {}) {
  const { t } = useTranslation()
  const navigate = useNavigate()
  const [searchTerm, setSearchTerm] = useState('')
  const [isCreateOpen, setIsCreateOpen] = useState(false)
  const [deleteCandidate, setDeleteCandidate] = useState<GallerySummary | null>(null)

  const filteredGalleries = useMemo(() => {
    if (!galleries) {
      return undefined
    }

    const trimmed = searchTerm.trim().toLowerCase()
    if (!trimmed) {
      return galleries
    }

    return galleries.filter((gallery) =>
      gallery.name.toLowerCase().includes(trimmed),
    )
  }, [galleries, searchTerm])

  async function handleCreateSubmit(
    values: CreateGalleryFormValues,
    frontImage?: File,
  ) {
    const result = await onCreate?.(values, frontImage)
    if (result !== false) {
      setIsCreateOpen(false)
    }
  }

  async function handleDeleteConfirm() {
    if (!deleteCandidate) {
      return
    }

    const result = await onDelete?.(deleteCandidate)
    if (result !== false) {
      setDeleteCandidate(null)
    }
  }

  return (
    <CmsAppShell activeKey="media">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: t('mediaLibrary.breadcrumb.media') },
            {
              label: t('mediaLibrary.breadcrumb.galleryLibrary'),
            },
          ]}
        />

        <header className={styles.pageHeader}>
          <div>
            <h1 className={styles.title}>{t('mediaLibrary.title')}</h1>
            <p className={styles.subtitle}>{t('mediaLibrary.subtitle')}</p>
          </div>
          <button
            type="button"
            className={styles.primaryButton}
            onClick={() => setIsCreateOpen(true)}
          >
            <AddIcon size={16} />
            {t('mediaLibrary.create.trigger')}
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
            placeholder={t('mediaLibrary.searchPlaceholder')}
            aria-label={t('mediaLibrary.searchPlaceholder')}
          />
        </div>

        {error && <p className={styles.errorText}>{error}</p>}

        {loading && !galleries ? (
          <div className={styles.loaderWrap}>
            <Loader label={t('mediaLibrary.loading')} />
          </div>
        ) : filteredGalleries === undefined ? (
          <EmptyState
            title={t('mediaLibrary.empty.title')}
            text={t('mediaLibrary.empty.text')}
          />
        ) : filteredGalleries.length === 0 ? (
          <EmptyState
            title={
              searchTerm
                ? t('mediaLibrary.empty.noMatchTitle')
                : t('mediaLibrary.empty.title')
            }
            text={
              searchTerm
                ? t('mediaLibrary.empty.noMatchText', { term: searchTerm })
                : t('mediaLibrary.empty.text')
            }
          />
        ) : (
          <div className={styles.grid}>
            {filteredGalleries.map((gallery) => (
              <GalleryCard
                key={gallery.id}
                gallery={gallery}
                onManage={(item) => navigate(`/media-library/${item.id}`)}
                onDelete={(item) => setDeleteCandidate(item)}
              />
            ))}
          </div>
        )}

        <CreateGalleryDialog
          open={isCreateOpen}
          onClose={() => setIsCreateOpen(false)}
          onSubmit={handleCreateSubmit}
          submitting={creating}
        />
        <ConfirmDialog
          open={Boolean(deleteCandidate)}
          title={t('mediaLibrary.delete.title')}
          body={
            <Trans
              i18nKey="mediaLibrary.delete.description"
              values={{ title: deleteCandidate?.name ?? '' }}
              components={{ galleryName: <strong /> }}
            />
          }
          confirmLabel={t('mediaLibrary.delete.confirm')}
          destructive
          busy={deleting}
          onConfirm={() => void handleDeleteConfirm()}
          onClose={() => setDeleteCandidate(null)}
        />
      </div>
    </CmsAppShell>
  )
}

function EmptyState({ title, text }: { title: string; text: string }) {
  return (
    <div className={styles.emptyState}>
      <p className={styles.emptyTitle}>{title}</p>
      <p className={styles.emptyText}>{text}</p>
    </div>
  )
}
