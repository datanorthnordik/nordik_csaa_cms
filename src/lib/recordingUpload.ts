export const RECORDING_UPLOAD_MAX_FILE_SIZE_MB = 100
export const RECORDING_UPLOAD_MAX_FILE_SIZE_BYTES =
  RECORDING_UPLOAD_MAX_FILE_SIZE_MB * 1024 * 1024
export const RECORDING_UPLOAD_SUPPORTED_FORMATS_LABEL =
  'MP3, WAV, M4A, AAC, OGG, FLAC, and WEBM audio'

const RECORDING_MIME_TYPES = new Set([
  'audio/aac',
  'audio/flac',
  'audio/m4a',
  'audio/mp4',
  'audio/mpeg',
  'audio/ogg',
  'audio/wav',
  'audio/webm',
  'audio/x-m4a',
  'audio/x-wav',
])

const RECORDING_EXTENSIONS = new Set([
  '.aac',
  '.flac',
  '.m4a',
  '.mp3',
  '.oga',
  '.ogg',
  '.wav',
  '.webm',
])

export const RECORDING_FILE_ACCEPT = [
  ...RECORDING_MIME_TYPES,
  ...RECORDING_EXTENSIONS,
].join(',')

export type RecordingUploadValidationError =
  | 'file-too-large'
  | 'unsupported-file-type'

export function validateRecordingUploadFile(
  file: Pick<File, 'name' | 'size' | 'type'>,
): RecordingUploadValidationError | null {
  const mimeType = file.type.trim().toLowerCase()
  const extensionIndex = file.name.lastIndexOf('.')
  const extension = extensionIndex >= 0
    ? file.name.slice(extensionIndex).trim().toLowerCase()
    : ''

  if (!RECORDING_MIME_TYPES.has(mimeType) && !RECORDING_EXTENSIONS.has(extension)) {
    return 'unsupported-file-type'
  }

  if (file.size > RECORDING_UPLOAD_MAX_FILE_SIZE_BYTES) {
    return 'file-too-large'
  }

  return null
}

export function getRecordingUploadValidationErrorMessage(
  error: RecordingUploadValidationError,
) {
  if (error === 'file-too-large') {
    return `This recording exceeds the ${RECORDING_UPLOAD_MAX_FILE_SIZE_MB}MB limit.`
  }

  return `Only ${RECORDING_UPLOAD_SUPPORTED_FORMATS_LABEL} are supported.`
}

export function assertValidRecordingUploadFile(
  file: Pick<File, 'name' | 'size' | 'type'>,
) {
  const error = validateRecordingUploadFile(file)
  if (error) {
    throw new Error(getRecordingUploadValidationErrorMessage(error))
  }
}
