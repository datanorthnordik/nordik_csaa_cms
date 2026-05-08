import { useId, useRef, useState, type DragEvent, type ReactNode } from 'react'
import styles from './UploadDropzone.module.css'

type UploadDropzoneProps = {
  onFiles?: (files: File[]) => void
  multiple?: boolean
  accept?: string
  label?: ReactNode
  hint?: ReactNode
  icon?: ReactNode
  variant?: 'default' | 'compact'
  disabled?: boolean
  className?: string
}

export function UploadDropzone({
  onFiles,
  multiple = false,
  accept,
  label,
  hint,
  icon,
  variant = 'default',
  disabled = false,
  className,
}: UploadDropzoneProps) {
  const inputRef = useRef<HTMLInputElement>(null)
  const inputId = useId()
  const [isDragOver, setIsDragOver] = useState(false)

  function emit(fileList: FileList | null) {
    if (!fileList || !onFiles) {
      return
    }

    const files = Array.from(fileList)
    if (files.length) {
      onFiles(files)
    }
  }

  function handleDragOver(event: DragEvent<HTMLLabelElement>) {
    event.preventDefault()
    if (disabled) {
      return
    }
    setIsDragOver(true)
  }

  function handleDragLeave() {
    setIsDragOver(false)
  }

  function handleDrop(event: DragEvent<HTMLLabelElement>) {
    event.preventDefault()
    setIsDragOver(false)
    if (disabled) {
      return
    }
    emit(event.dataTransfer.files)
  }

  const classNames = [
    styles.dropzone,
    variant === 'compact' ? styles.compact : '',
    isDragOver ? styles.dragOver : '',
    disabled ? styles.disabled : '',
    className,
  ]
    .filter(Boolean)
    .join(' ')

  return (
    <label
      htmlFor={inputId}
      className={classNames}
      onDragOver={handleDragOver}
      onDragEnter={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      <input
        ref={inputRef}
        id={inputId}
        type="file"
        className={styles.input}
        multiple={multiple}
        accept={accept}
        disabled={disabled}
        onChange={(event) => emit(event.target.files)}
      />
      {icon && <span className={styles.icon}>{icon}</span>}
      {label && <span className={styles.label}>{label}</span>}
      {hint && <span className={styles.hint}>{hint}</span>}
    </label>
  )
}
