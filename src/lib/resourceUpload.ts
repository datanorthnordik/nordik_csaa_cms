export const RESOURCE_UPLOAD_MAX_FILE_SIZE_MB = 20
export const RESOURCE_UPLOAD_MAX_FILE_SIZE_BYTES =
  RESOURCE_UPLOAD_MAX_FILE_SIZE_MB * 1024 * 1024
export const RESOURCE_UPLOAD_SUPPORTED_FORMATS_LABEL =
  'PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP'
export const RESOURCE_UPLOAD_SUPPORTED_IMAGE_FORMATS_LABEL =
  'SVG, PNG, JPG, and WEBP'

const RESOURCE_SUPPORTED_MIME_TYPES = new Set([
  'application/pdf',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'image/svg+xml',
  'image/png',
  'image/jpeg',
  'image/jpg',
  'image/webp',
])

const RESOURCE_SUPPORTED_IMAGE_MIME_TYPES = new Set([
  'image/svg+xml',
  'image/png',
  'image/jpeg',
  'image/jpg',
  'image/webp',
])

const RESOURCE_SUPPORTED_EXTENSIONS = new Set([
  '.pdf',
  '.docx',
  '.xlsx',
  '.pptx',
  '.svg',
  '.png',
  '.jpg',
  '.jpeg',
  '.webp',
])

const RESOURCE_SUPPORTED_IMAGE_EXTENSIONS = new Set([
  '.svg',
  '.png',
  '.jpg',
  '.jpeg',
  '.webp',
])

export const RESOURCE_FILE_ACCEPT = [
  ...RESOURCE_SUPPORTED_MIME_TYPES,
  ...RESOURCE_SUPPORTED_EXTENSIONS,
].join(',')

export const RESOURCE_IMAGE_FILE_ACCEPT = [
  ...RESOURCE_SUPPORTED_IMAGE_MIME_TYPES,
  ...RESOURCE_SUPPORTED_IMAGE_EXTENSIONS,
].join(',')

export type ResourceUploadValidationError =
  | 'file-too-large'
  | 'unsupported-file-type'

export function validateResourceUploadFile(file: Pick<File, 'name' | 'size' | 'type'>) {
  return validateFileAgainstAllowedTypes(
    file,
    RESOURCE_SUPPORTED_MIME_TYPES,
    RESOURCE_SUPPORTED_EXTENSIONS,
  )
}

export function validateResourceImageUploadFile(file: Pick<File, 'name' | 'size' | 'type'>) {
  return validateFileAgainstAllowedTypes(
    file,
    RESOURCE_SUPPORTED_IMAGE_MIME_TYPES,
    RESOURCE_SUPPORTED_IMAGE_EXTENSIONS,
  )
}

export function getResourceUploadValidationErrorMessage(
  validationError: ResourceUploadValidationError,
) {
  if (validationError === 'file-too-large') {
    return `This file exceeds the ${RESOURCE_UPLOAD_MAX_FILE_SIZE_MB}MB limit.`
  }

  return `Only ${RESOURCE_UPLOAD_SUPPORTED_FORMATS_LABEL} are supported.`
}

export function getResourceImageUploadValidationErrorMessage(
  validationError: ResourceUploadValidationError,
) {
  if (validationError === 'file-too-large') {
    return `This file exceeds the ${RESOURCE_UPLOAD_MAX_FILE_SIZE_MB}MB limit.`
  }

  return `Only ${RESOURCE_UPLOAD_SUPPORTED_IMAGE_FORMATS_LABEL} are supported.`
}

export function assertValidResourceUploadFile(file: Pick<File, 'name' | 'size' | 'type'>) {
  const validationError = validateResourceUploadFile(file)
  if (!validationError) {
    return
  }

  throw new Error(getResourceUploadValidationErrorMessage(validationError))
}

export function assertValidResourceImageUploadFile(file: Pick<File, 'name' | 'size' | 'type'>) {
  const validationError = validateResourceImageUploadFile(file)
  if (!validationError) {
    return
  }

  throw new Error(getResourceImageUploadValidationErrorMessage(validationError))
}

function validateFileAgainstAllowedTypes(
  file: Pick<File, 'name' | 'size' | 'type'>,
  supportedMimeTypes: Set<string>,
  supportedExtensions: Set<string>,
) {
  if (!isSupportedFile(file, supportedMimeTypes, supportedExtensions)) {
    return 'unsupported-file-type' satisfies ResourceUploadValidationError
  }

  if (file.size > RESOURCE_UPLOAD_MAX_FILE_SIZE_BYTES) {
    return 'file-too-large' satisfies ResourceUploadValidationError
  }

  return null
}

function isSupportedFile(
  file: Pick<File, 'name' | 'type'>,
  supportedMimeTypes: Set<string>,
  supportedExtensions: Set<string>,
) {
  const normalizedMimeType = file.type.trim().toLowerCase()
  if (normalizedMimeType && supportedMimeTypes.has(normalizedMimeType)) {
    return true
  }

  const extension = getFileExtension(file.name)
  return extension ? supportedExtensions.has(extension) : false
}

function getFileExtension(fileName: string) {
  const lastDotIndex = fileName.lastIndexOf('.')
  if (lastDotIndex < 0) {
    return ''
  }

  return fileName.slice(lastDotIndex).trim().toLowerCase()
}
