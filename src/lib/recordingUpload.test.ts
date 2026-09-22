import { describe, expect, it } from 'vitest'
import {
  RECORDING_UPLOAD_MAX_FILE_SIZE_BYTES,
  validateRecordingUploadFile,
} from './recordingUpload'

describe('recordingUpload', () => {
  it('accepts supported audio MIME types and extensions', () => {
    expect(validateRecordingUploadFile({ name: 'story.mp3', size: 10, type: 'audio/mpeg' })).toBeNull()
    expect(validateRecordingUploadFile({ name: 'story.m4a', size: 10, type: '' })).toBeNull()
  })

  it('rejects non-audio files', () => {
    expect(validateRecordingUploadFile({ name: 'notes.pdf', size: 10, type: 'application/pdf' })).toBe(
      'unsupported-file-type',
    )
  })

  it('rejects recordings over the size limit', () => {
    expect(
      validateRecordingUploadFile({
        name: 'story.mp3',
        size: RECORDING_UPLOAD_MAX_FILE_SIZE_BYTES + 1,
        type: 'audio/mpeg',
      }),
    ).toBe('file-too-large')
  })
})
