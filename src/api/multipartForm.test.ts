import { describe, expect, it } from 'vitest'
import { buildMultipartPayload } from './multipartForm'

describe('buildMultipartPayload', () => {
  it('appends the payload as a JSON string', () => {
    const payload = { title: 'Test', count: 3 }
    const form = buildMultipartPayload(payload, [])
    expect(form.get('payload')).toBe(JSON.stringify(payload))
  })

  it('appends a file with its field name and file name', () => {
    const file = new File(['content'], 'photo.png', { type: 'image/png' })
    const form = buildMultipartPayload({}, [
      { fieldName: 'image_file', file, fileName: 'photo.png' },
    ])
    const appended = form.get('image_file') as File
    expect(appended).toBeInstanceOf(File)
    expect(appended.name).toBe('photo.png')
  })

  it('appends a file without a custom file name when fileName is omitted', () => {
    const file = new File(['content'], 'unnamed.png', { type: 'image/png' })
    const form = buildMultipartPayload({}, [{ fieldName: 'raw_file', file }])
    expect(form.get('raw_file')).toBeInstanceOf(Blob)
  })

  it('skips entries where file is null', () => {
    const form = buildMultipartPayload({ title: 'ok' }, [
      { fieldName: 'missing_file', file: null },
    ])
    expect(form.get('missing_file')).toBeNull()
    expect(form.get('payload')).toBe(JSON.stringify({ title: 'ok' }))
  })

  it('skips entries where file is undefined', () => {
    const form = buildMultipartPayload({ title: 'ok' }, [
      { fieldName: 'also_missing', file: undefined },
    ])
    expect(form.get('also_missing')).toBeNull()
  })

  it('handles multiple file entries with indexed field names', () => {
    const file1 = new File(['a'], 'a.jpg', { type: 'image/jpeg' })
    const file2 = new File(['b'], 'b.jpg', { type: 'image/jpeg' })
    const form = buildMultipartPayload({ count: 2 }, [
      { fieldName: 'images[0].file', file: file1, fileName: 'a.jpg' },
      { fieldName: 'images[1].file', file: file2, fileName: 'b.jpg' },
    ])
    expect((form.get('images[0].file') as File).name).toBe('a.jpg')
    expect((form.get('images[1].file') as File).name).toBe('b.jpg')
  })

  it('serialises complex payload objects correctly', () => {
    const payload = {
      name: 'Gallery',
      published: false,
      images: [{ title: 'img1', alt_text: 'alt' }],
    }
    const form = buildMultipartPayload(payload, [])
    expect(form.get('payload')).toBe(JSON.stringify(payload))
  })
})
