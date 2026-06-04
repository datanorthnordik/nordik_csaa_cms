import { describe, expect, it } from 'vitest'
import {
  GALLERY_IMAGE_UPLOAD_MAX_FILE_SIZE_BYTES,
  GALLERY_IMAGE_UPLOAD_MAX_FILE_SIZE_MB,
  RESOURCE_FILE_ACCEPT,
  RESOURCE_IMAGE_FILE_ACCEPT,
  RESOURCE_UPLOAD_MAX_FILE_SIZE_BYTES,
  RESOURCE_UPLOAD_MAX_FILE_SIZE_MB,
  RESOURCE_UPLOAD_SUPPORTED_IMAGE_FORMATS_LABEL,
  validateGalleryImageUploadFile,
  validateResourceUploadFile,
  validateResourceImageUploadFile,
} from './resourceUpload'

describe('resourceUpload', () => {
  it('keeps the max upload size in a single MB and byte constant', () => {
    expect(RESOURCE_UPLOAD_MAX_FILE_SIZE_MB).toBe(20)
    expect(RESOURCE_UPLOAD_MAX_FILE_SIZE_BYTES).toBe(20 * 1024 * 1024)
    expect(GALLERY_IMAGE_UPLOAD_MAX_FILE_SIZE_MB).toBe(9)
    expect(GALLERY_IMAGE_UPLOAD_MAX_FILE_SIZE_BYTES).toBe(9 * 1024 * 1024)
  })

  it('accepts supported office and image file types only', () => {
    expect(validateResourceUploadFile(new File(['doc'], 'guide.docx', {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }))).toBeNull()

    expect(validateResourceUploadFile(new File(['image'], 'poster.jpeg', {
      type: '',
    }))).toBeNull()

    expect(validateResourceUploadFile(new File(['video'], 'clip.mp4', {
      type: 'video/mp4',
    }))).toBe('unsupported-file-type')

    expect(validateResourceUploadFile(new File(['audio'], 'voice.mp3', {
      type: 'audio/mpeg',
    }))).toBe('unsupported-file-type')
  })

  it('rejects supported files that exceed the size limit', () => {
    const largePdf = new File(['pdf'], 'large.pdf', {
      type: 'application/pdf',
    })
    Object.defineProperty(largePdf, 'size', {
      value: RESOURCE_UPLOAD_MAX_FILE_SIZE_BYTES + 1,
    })

    expect(validateResourceUploadFile(largePdf)).toBe('file-too-large')
  })

  it('accepts only the supported image types for image-only uploads', () => {
    expect(validateResourceImageUploadFile(new File(['image'], 'cover.webp', {
      type: 'image/webp',
    }))).toBeNull()

    expect(validateResourceImageUploadFile(new File(['document'], 'cover.pdf', {
      type: 'application/pdf',
    }))).toBe('unsupported-file-type')
  })

  it('applies the stricter gallery image size limit', () => {
    const largeImage = new File(['image'], 'gallery-cover.png', {
      type: 'image/png',
    })
    Object.defineProperty(largeImage, 'size', {
      value: GALLERY_IMAGE_UPLOAD_MAX_FILE_SIZE_BYTES + 1,
    })

    expect(validateGalleryImageUploadFile(largeImage)).toBe('file-too-large')
    expect(
      validateGalleryImageUploadFile(
        new File(['image'], 'gallery-cover.webp', {
          type: 'image/webp',
        }),
      ),
    ).toBeNull()
  })

  it('exposes an exact accept list instead of broad wildcards', () => {
    expect(RESOURCE_FILE_ACCEPT).toContain('application/pdf')
    expect(RESOURCE_FILE_ACCEPT).toContain('image/webp')
    expect(RESOURCE_FILE_ACCEPT).toContain('.docx')
    expect(RESOURCE_FILE_ACCEPT).not.toContain('image/*')
    expect(RESOURCE_FILE_ACCEPT).not.toContain('.doc,')
    expect(RESOURCE_FILE_ACCEPT).not.toContain('.xls,')
    expect(RESOURCE_FILE_ACCEPT).not.toContain('.ppt,')
    expect(RESOURCE_IMAGE_FILE_ACCEPT).toContain('image/png')
    expect(RESOURCE_IMAGE_FILE_ACCEPT).toContain('.webp')
    expect(RESOURCE_IMAGE_FILE_ACCEPT).not.toContain('application/pdf')
    expect(RESOURCE_UPLOAD_SUPPORTED_IMAGE_FORMATS_LABEL).toBe('SVG, PNG, JPG, and WEBP')
  })
})
