import Dialog from '@mui/material/Dialog'
import { useState } from 'react'
import { useForm } from 'react-hook-form'
import { useTranslation } from 'react-i18next'
import { CloseIcon, CloudUploadIcon } from '../icons'
import type { MediaVisibility } from '../../types/media'
import { UploadDropzone } from './UploadDropzone'
import styles from './CreateGalleryDialog.module.css'

export type CreateGalleryFormValues = {
  name: string
  description: string
  visibility: Extract<MediaVisibility, 'draft' | 'published'>
}

type CreateGalleryDialogProps = {
  open: boolean
  onClose: () => void
  onSubmit?: (values: CreateGalleryFormValues, frontImage?: File) => void
  submitting?: boolean
}

const defaultValues: CreateGalleryFormValues = {
  name: '',
  description: '',
  visibility: 'draft',
}

export function CreateGalleryDialog({
  open,
  onClose,
  onSubmit,
  submitting = false,
}: CreateGalleryDialogProps) {
  const { t } = useTranslation()
  const {
    register,
    handleSubmit,
    reset,
    formState: { errors, isSubmitting },
  } = useForm<CreateGalleryFormValues>({ defaultValues })
  const [frontImage, setFrontImage] = useState<File | null>(null)

  function onValid(values: CreateGalleryFormValues) {
    onSubmit?.(values, frontImage ?? undefined)
  }

  const isBusy = submitting || isSubmitting

  function handleClose() {
    if (isBusy) {
      return
    }
    reset(defaultValues)
    setFrontImage(null)
    onClose()
  }

  return (
    <Dialog
      open={open}
      onClose={handleClose}
      slotProps={{
        paper: { className: styles.paper },
      }}
    >
      <form onSubmit={handleSubmit(onValid)} noValidate>
        <header className={styles.header}>
          <h2 className={styles.title}>{t('mediaLibrary.create.title')}</h2>
          <button
            type="button"
            className={styles.closeButton}
            onClick={handleClose}
            disabled={isBusy}
            aria-label={t('mediaLibrary.create.close')}
          >
            <CloseIcon />
          </button>
        </header>

        <div className={styles.body}>
          <label className={styles.field}>
            <span className={styles.fieldLabel}>
              {t('mediaLibrary.create.fields.name')}
              <span className={styles.required} aria-hidden="true">
                *
              </span>
            </span>
            <input
              type="text"
              className={styles.input}
              placeholder={t('mediaLibrary.create.fields.namePlaceholder')}
              disabled={isBusy}
              aria-invalid={errors.name ? true : undefined}
              {...register('name', {
                required: t('mediaLibrary.create.errors.nameRequired'),
              })}
            />
            {errors.name?.message && (
              <span className={styles.errorText} role="alert">
                {errors.name.message}
              </span>
            )}
          </label>

          <label className={styles.field}>
            <span className={styles.fieldLabel}>
              {t('mediaLibrary.create.fields.description')}
            </span>
            <textarea
              className={styles.textarea}
              rows={3}
              placeholder={t('mediaLibrary.create.fields.descriptionPlaceholder')}
              disabled={isBusy}
              {...register('description')}
            />
          </label>

          <fieldset className={styles.field}>
            <legend className={styles.fieldLabel}>
              {t('mediaLibrary.create.fields.visibility')}
            </legend>
            <div className={styles.radioRow}>
              <label className={styles.radioOption}>
                <input
                  type="radio"
                  value="draft"
                  disabled={isBusy}
                  {...register('visibility')}
                />
                <span>{t('mediaLibrary.visibility.draft')}</span>
              </label>
              <label className={styles.radioOption}>
                <input
                  type="radio"
                  value="published"
                  disabled={isBusy}
                  {...register('visibility')}
                />
                <span>{t('mediaLibrary.visibility.published')}</span>
              </label>
            </div>
          </fieldset>

          <div className={styles.field}>
            <span className={styles.fieldLabel}>
              {t('mediaLibrary.create.fields.frontImage')}
            </span>
            <UploadDropzone
              icon={<CloudUploadIcon />}
              accept="image/png,image/jpeg"
              label={t('mediaLibrary.create.fields.frontImageLabel')}
              hint={t('mediaLibrary.create.fields.frontImageHint')}
              disabled={isBusy}
              onFiles={(files) => setFrontImage(files[0] ?? null)}
            />
            {frontImage && (
              <p className={styles.fileName}>{frontImage.name}</p>
            )}
          </div>
        </div>

        <footer className={styles.footer}>
          <button
            type="button"
            className={styles.secondaryButton}
            onClick={handleClose}
            disabled={isBusy}
          >
            {t('mediaLibrary.create.cancel')}
          </button>
          <button
            type="submit"
            className={styles.primaryButton}
            disabled={isBusy}
          >
            {isBusy
              ? t('mediaLibrary.create.submitting')
              : t('mediaLibrary.create.submit')}
          </button>
        </footer>
      </form>
    </Dialog>
  )
}
