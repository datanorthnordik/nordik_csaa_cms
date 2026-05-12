import { beforeEach, describe, expect, it, vi } from 'vitest'
import { API_ROUTES } from '../constants/api'

const { deleteMock, getMock, patchMock, postMock, putMock } = vi.hoisted(() => ({
  deleteMock: vi.fn(),
  getMock: vi.fn(),
  patchMock: vi.fn(),
  postMock: vi.fn(),
  putMock: vi.fn(),
}))

vi.mock('./apiClient', () => ({
  apiClient: {
    delete: deleteMock,
    get: getMock,
    patch: patchMock,
    post: postMock,
    put: putMock,
  },
}))

import { eventsApi } from './eventsApi'
import { mediaApi } from './mediaApi'
import { pagesApi } from './pagesApi'

describe('upload request bodies', () => {
  beforeEach(() => {
    deleteMock.mockReset()
    getMock.mockReset()
    patchMock.mockReset()
    postMock.mockReset()
    putMock.mockReset()

    postMock.mockResolvedValue({ data: { message: 'ok' } })
    putMock.mockResolvedValue({ data: { message: 'ok' } })
  })

  it('sends page hero uploads as multipart payload plus hero_image_file', async () => {
    const heroImageFile = new File(['hero'], 'hero.png', { type: 'image/png' })

    await pagesApi.createPage({
      page_title: 'Homepage',
      url_slug: '/home',
<<<<<<< HEAD
<<<<<<< HEAD
=======
      parent_page_id: null,
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
=======
      parent_page_id: null,
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
      status: 'draft',
      hero_image_enabled: true,
      hero_image: {
        file_name: 'hero.png',
        mime_type: 'image/png',
      },
      remove_hero_image: false,
      seo_page_title: 'Homepage SEO',
      seo_page_description: 'Description',
      heroImageFile,
    })

    expect(postMock).toHaveBeenCalledTimes(1)
    expect(postMock.mock.calls[0]?.[0]).toBe(API_ROUTES.pages)

    const body = postMock.mock.calls[0]?.[1]
    expect(body).toBeInstanceOf(FormData)
    expect((body as FormData).get('payload')).toBe(
      JSON.stringify({
        page_title: 'Homepage',
        url_slug: '/home',
<<<<<<< HEAD
<<<<<<< HEAD
=======
        parent_page_id: null,
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
=======
        parent_page_id: null,
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
        status: 'draft',
        hero_image_enabled: true,
        hero_image: {
          file_name: 'hero.png',
          mime_type: 'image/png',
        },
        remove_hero_image: false,
        seo_page_title: 'Homepage SEO',
        seo_page_description: 'Description',
      }),
    )
    expect(((body as FormData).get('hero_image_file') as File).name).toBe('hero.png')
  })

<<<<<<< HEAD
<<<<<<< HEAD
  it('sends page hero update as multipart PUT with hero_image_file', async () => {
    const heroImageFile = new File(['hero'], 'updated-hero.png', {
      type: 'image/png',
    })

    await pagesApi.updatePage(5, {
      page_title: 'About',
      url_slug: '/about',
      status: 'published',
      hero_image_enabled: true,
      hero_image: {
        file_name: 'updated-hero.png',
        mime_type: 'image/png',
      },
      remove_hero_image: false,
      seo_page_title: 'About SEO',
      seo_page_description: '',
      heroImageFile,
    })

    expect(putMock).toHaveBeenCalledTimes(1)
    expect(putMock.mock.calls[0]?.[0]).toBe(API_ROUTES.pageById(5))

    const body = putMock.mock.calls[0]?.[1] as FormData
    expect(body).toBeInstanceOf(FormData)
    expect(((body).get('hero_image_file') as File).name).toBe('updated-hero.png')
  })

=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
  it('sends gallery image uploads as multipart payload with indexed file fields', async () => {
    const bannerFile = new File(['banner'], 'banner.png', { type: 'image/png' })
    const detailFile = new File(['detail'], 'detail.png', { type: 'image/png' })

    await mediaApi.uploadGalleryImages(7, {
      images: [
        {
          title: 'Opening banner',
          alt_text: 'Banner',
          file_name: 'banner.png',
          mime_type: 'image/png',
        },
        {
          title: 'Detail shot',
          alt_text: 'Detail',
          file_name: 'detail.png',
          mime_type: 'image/png',
        },
      ],
      imageFiles: [bannerFile, detailFile],
    })

    expect(postMock).toHaveBeenCalledTimes(1)
    expect(postMock.mock.calls[0]?.[0]).toBe(API_ROUTES.galleryImagesById(7))

    const body = postMock.mock.calls[0]?.[1] as FormData
    expect(body).toBeInstanceOf(FormData)
    expect((body.get('images[0].file') as File).name).toBe('banner.png')
    expect((body.get('images[1].file') as File).name).toBe('detail.png')
    expect(body.get('payload')).toBe(
      JSON.stringify({
        images: [
          {
            title: 'Opening banner',
            alt_text: 'Banner',
            file_name: 'banner.png',
            mime_type: 'image/png',
          },
          {
            title: 'Detail shot',
            alt_text: 'Detail',
            file_name: 'detail.png',
            mime_type: 'image/png',
          },
        ],
      }),
    )
  })

  it('sends event display images and attachments with the expected multipart field names', async () => {
    const posterFile = new File(['poster'], 'poster.png', { type: 'image/png' })
    const agendaFile = new File(['agenda'], 'agenda.pdf', {
      type: 'application/pdf',
    })

    await eventsApi.updateEvent(11, {
      title: 'Spring Fair',
      show_title: true,
      categories: ['Events'],
      event_type: 'single_day_all_day',
      start_at: '2026-05-01T10:00:00Z',
      end_at: null,
      privacy_type: 'public',
      private_audiences: [],
      published: false,
      request_review: false,
      review_email_list: [],
      teaser: 'Welcome!',
      description_html: '',
      contact_name: '',
      contact_email: '',
      contact_phone: '',
      contact_ext: '',
      contact_fax: '',
      location_mode: 'none',
      address: undefined,
      show_display_image_when_viewing: true,
      gallery_id: null,
      registration_enabled: false,
      registration_start_at: null,
      registration_end_at: null,
      registration_url: '',
      repeat_enabled: false,
      recurrence_type: '',
      recurrence_frequency: '',
      recurrence_interval: 1,
      recurrence_until: null,
      recurrence_rule: undefined,
      occurrences: [],
      display_image: {
        display_name: 'poster.png',
        file_name: 'poster.png',
        mime_type: 'image/png',
      },
      attachments: [
        {
          display_name: 'agenda.pdf',
          file_name: 'agenda.pdf',
          mime_type: 'application/pdf',
        },
      ],
      displayImageFile: posterFile,
      attachmentFiles: [agendaFile],
    })

    expect(putMock).toHaveBeenCalledTimes(1)
    expect(putMock.mock.calls[0]?.[0]).toBe(API_ROUTES.eventById(11))

    const body = putMock.mock.calls[0]?.[1] as FormData
    expect(body).toBeInstanceOf(FormData)
    expect((body.get('display_image_file') as File).name).toBe('poster.png')
    expect((body.get('attachments[0].file') as File).name).toBe('agenda.pdf')
    expect(body.get('payload')).toBe(
      JSON.stringify({
        title: 'Spring Fair',
        show_title: true,
        categories: ['Events'],
        event_type: 'single_day_all_day',
        start_at: '2026-05-01T10:00:00Z',
        end_at: null,
        privacy_type: 'public',
        private_audiences: [],
        published: false,
        request_review: false,
        review_email_list: [],
        teaser: 'Welcome!',
        description_html: '',
        contact_name: '',
        contact_email: '',
        contact_phone: '',
        contact_ext: '',
        contact_fax: '',
        location_mode: 'none',
        address: undefined,
        show_display_image_when_viewing: true,
        gallery_id: null,
        registration_enabled: false,
        registration_start_at: null,
        registration_end_at: null,
        registration_url: '',
        repeat_enabled: false,
        recurrence_type: '',
        recurrence_frequency: '',
        recurrence_interval: 1,
        recurrence_until: null,
        recurrence_rule: undefined,
        occurrences: [],
        display_image: {
          display_name: 'poster.png',
          file_name: 'poster.png',
          mime_type: 'image/png',
        },
        attachments: [
          {
            display_name: 'agenda.pdf',
            file_name: 'agenda.pdf',
            mime_type: 'application/pdf',
          },
        ],
      }),
    )
  })

<<<<<<< HEAD
<<<<<<< HEAD
  it('sends createEvent with display image as a multipart POST', async () => {
    const posterFile = new File(['poster'], 'poster.png', { type: 'image/png' })

    await eventsApi.createEvent({
      title: 'New Event',
      show_title: true,
      categories: ['Community'],
      event_type: 'single_day_all_day',
      start_at: '2026-06-01T00:00:00Z',
      end_at: null,
      privacy_type: 'public',
      private_audiences: [],
      published: false,
      request_review: false,
      review_email_list: [],
      teaser: '',
      description_html: '',
      contact_name: '',
      contact_email: '',
      contact_phone: '',
      contact_ext: '',
      contact_fax: '',
      location_mode: 'none',
      address: undefined,
      show_display_image_when_viewing: true,
      gallery_id: null,
      registration_enabled: false,
      registration_start_at: null,
      registration_end_at: null,
      registration_url: '',
      repeat_enabled: false,
      recurrence_type: '',
      recurrence_frequency: '',
      recurrence_interval: 1,
      recurrence_until: null,
      recurrence_rule: undefined,
      occurrences: [],
      display_image: {
        display_name: 'poster.png',
        file_name: 'poster.png',
        mime_type: 'image/png',
      },
      attachments: [],
      displayImageFile: posterFile,
    })

    expect(postMock).toHaveBeenCalledTimes(1)
    expect(postMock.mock.calls[0]?.[0]).toBe(API_ROUTES.events)

    const body = postMock.mock.calls[0]?.[1] as FormData
    expect(body).toBeInstanceOf(FormData)
    expect((body.get('display_image_file') as File).name).toBe('poster.png')
  })

  it('sends createGallery cover image as a multipart POST', async () => {
    const coverFile = new File(['cover'], 'cover.jpg', { type: 'image/jpeg' })

    await mediaApi.createGallery({
      name: 'Summer 2026',
      published: false,
      coverImageFile: coverFile,
    })

    expect(postMock).toHaveBeenCalledTimes(1)
    expect(postMock.mock.calls[0]?.[0]).toBe(API_ROUTES.galleries)

    const body = postMock.mock.calls[0]?.[1] as FormData
    expect(body).toBeInstanceOf(FormData)
    expect((body.get('cover_image_file') as File).name).toBe('cover.jpg')
=======
=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
  it('deletes event documents using the storage_url query parameter', async () => {
    await eventsApi.deleteEventDocument(16, 'gs://bucket/events/16/documents/agenda.pdf')

    expect(deleteMock).toHaveBeenCalledTimes(1)
    expect(deleteMock).toHaveBeenCalledWith(
      '/api/events/16/document',
      {
        params: {
          storage_url: 'gs://bucket/events/16/documents/agenda.pdf',
        },
      },
    )
  })

  it('deletes event photos using the storage_url query parameter', async () => {
    await eventsApi.deleteEventPhoto(16, 'gs://bucket/events/16/photo_274.jpg')

    expect(deleteMock).toHaveBeenCalledTimes(1)
    expect(deleteMock).toHaveBeenCalledWith(
      '/api/events/16/photo',
      {
        params: {
          storage_url: 'gs://bucket/events/16/photo_274.jpg',
        },
      },
    )
<<<<<<< HEAD
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
=======
>>>>>>> 4890b41c5b79edd78ad76b508a3f852018316578
  })
})
