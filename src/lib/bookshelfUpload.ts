const BOOK_FILE_EXTENSIONS = new Set(['.pdf', '.epub', '.doc', '.docx'])
const AUTHOR_IMAGE_FILE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.webp', '.svg'])
const COVER_FILE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.webp', '.svg'])

const BOOK_FILE_MIME_TYPES = new Set([
  'application/pdf',
  'application/epub+zip',
  'application/msword',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
])

const AUTHOR_IMAGE_FILE_MIME_TYPES = new Set([
  'image/png',
  'image/jpeg',
  'image/webp',
  'image/svg+xml',
])

const COVER_FILE_MIME_TYPES = new Set([
  'image/png',
  'image/jpeg',
  'image/webp',
  'image/svg+xml',
])

export const BOOKSHELF_BOOK_ACCEPT = '.pdf,.epub,.doc,.docx'
export const BOOKSHELF_AUTHOR_IMAGE_ACCEPT = '.png,.jpg,.jpeg,.webp,.svg'
export const BOOKSHELF_COVER_ACCEPT = '.png,.jpg,.jpeg,.webp,.svg'
export const BOOKSHELF_BOOK_MAX_FILE_SIZE_MB = 20
export const BOOKSHELF_AUTHOR_IMAGE_MAX_FILE_SIZE_MB = 5
export const BOOKSHELF_COVER_MAX_FILE_SIZE_MB = 5

export type BookshelfUploadValidationError = 'file-too-large' | 'file-unsupported'

export function validateBookshelfBookFile(
  file: Pick<File, 'name' | 'size' | 'type'>,
): BookshelfUploadValidationError | null {
  return validateFile(
    file,
    BOOKSHELF_BOOK_MAX_FILE_SIZE_MB,
    BOOK_FILE_EXTENSIONS,
    BOOK_FILE_MIME_TYPES,
  )
}

export function validateBookshelfCoverFile(
  file: Pick<File, 'name' | 'size' | 'type'>,
): BookshelfUploadValidationError | null {
  return validateFile(
    file,
    BOOKSHELF_COVER_MAX_FILE_SIZE_MB,
    COVER_FILE_EXTENSIONS,
    COVER_FILE_MIME_TYPES,
  )
}

export function validateBookshelfAuthorImageFile(
  file: Pick<File, 'name' | 'size' | 'type'>,
): BookshelfUploadValidationError | null {
  return validateFile(
    file,
    BOOKSHELF_AUTHOR_IMAGE_MAX_FILE_SIZE_MB,
    AUTHOR_IMAGE_FILE_EXTENSIONS,
    AUTHOR_IMAGE_FILE_MIME_TYPES,
  )
}

export function assertValidBookshelfBookFile(file: Pick<File, 'name' | 'size' | 'type'>) {
  const validationError = validateBookshelfBookFile(file)
  if (validationError) {
    throw new Error(validationError)
  }
}

export function assertValidBookshelfAuthorImageFile(
  file: Pick<File, 'name' | 'size' | 'type'>,
) {
  const validationError = validateBookshelfAuthorImageFile(file)
  if (validationError) {
    throw new Error(validationError)
  }
}

export function assertValidBookshelfCoverFile(file: Pick<File, 'name' | 'size' | 'type'>) {
  const validationError = validateBookshelfCoverFile(file)
  if (validationError) {
    throw new Error(validationError)
  }
}

function validateFile(
  file: Pick<File, 'name' | 'size' | 'type'>,
  maxSizeMb: number,
  allowedExtensions: Set<string>,
  allowedMimeTypes: Set<string>,
): BookshelfUploadValidationError | null {
  if (file.size > maxSizeMb * 1024 * 1024) {
    return 'file-too-large'
  }

  const extension = file.name.includes('.')
    ? `.${file.name.split('.').pop()?.trim().toLowerCase() ?? ''}`
    : ''
  const normalizedMimeType = file.type.trim().toLowerCase()

  if (allowedExtensions.has(extension)) {
    return null
  }
  if (normalizedMimeType && allowedMimeTypes.has(normalizedMimeType)) {
    return null
  }

  return 'file-unsupported'
}
