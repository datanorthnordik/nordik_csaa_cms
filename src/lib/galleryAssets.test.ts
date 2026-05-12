import { describe, expect, it } from 'vitest'
import {
  getGalleryAssetDetails,
  getGalleryAssetTitle,
  suggestGalleryAssetTitle,
} from './galleryAssets'

describe('suggestGalleryAssetTitle', () => {
  it('strips the file extension and replaces underscores with spaces', () => {
    expect(suggestGalleryAssetTitle('summer_sunset.jpg')).toBe('summer sunset')
  })

  it('replaces hyphens with spaces', () => {
    expect(suggestGalleryAssetTitle('photo-banner-2024.png')).toBe('photo banner 2024')
  })

  it('collapses consecutive separators', () => {
    expect(suggestGalleryAssetTitle('file--name__test.jpeg')).toBe('file name test')
  })

  it('returns the original filename when the result would be empty', () => {
    expect(suggestGalleryAssetTitle('.jpg')).toBe('.jpg')
  })

  it('handles a file name with no extension', () => {
    expect(suggestGalleryAssetTitle('portrait')).toBe('portrait')
  })

  it('handles a file name with multiple dots', () => {
    expect(suggestGalleryAssetTitle('event.2024.photo.jpg')).toBe('event.2024.photo')
  })
})

describe('getGalleryAssetTitle', () => {
  it('returns the trimmed title when present', () => {
    expect(
      getGalleryAssetTitle({ fileName: 'photo.jpg', title: '  My Photo  ' }),
    ).toBe('My Photo')
  })

  it('falls back to the suggested title when title is absent', () => {
    expect(getGalleryAssetTitle({ fileName: 'hero_image.jpg' })).toBe('hero image')
  })

  it('falls back to the suggested title when title is an empty string', () => {
    expect(
      getGalleryAssetTitle({ fileName: 'banner.jpg', title: '' }),
    ).toBe('banner')
  })

  it('falls back to the suggested title when title is whitespace only', () => {
    expect(
      getGalleryAssetTitle({ fileName: 'cover.jpg', title: '   ' }),
    ).toBe('cover')
  })
})

describe('getGalleryAssetDetails', () => {
  it('returns details when present', () => {
    expect(
      getGalleryAssetDetails({ details: 'A description', altText: 'Alt' }),
    ).toBe('A description')
  })

  it('falls back to altText when details is absent', () => {
    expect(
      getGalleryAssetDetails({ altText: 'Fallback alt text' }),
    ).toBe('Fallback alt text')
  })

  it('returns an empty string when neither details nor altText is present', () => {
    expect(getGalleryAssetDetails({})).toBe('')
  })

  it('returns an empty string when both details and altText are whitespace only', () => {
    expect(getGalleryAssetDetails({ details: '   ', altText: '  ' })).toBe('')
  })

  it('trims details before returning', () => {
    expect(getGalleryAssetDetails({ details: '  caption  ' })).toBe('caption')
  })
})
