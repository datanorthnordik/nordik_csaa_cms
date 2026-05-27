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
    const documentFile = new File(['policy'], 'policy.pdf', { type: 'application/pdf' })

    await pagesApi.createPage({
      page_title: 'Homepage',
      url_slug: '/home',
      parent_id: null,
      status: 'draft',
      hero_image_enabled: true,
      hero_image: {
        file_name: 'hero.png',
        mime_type: 'image/png',
      },
      remove_hero_image: false,
      seo_page_title: 'Homepage SEO',
      seo_page_description: 'Description',
      page_detail: {
        template_key: 'default',
        settings: {},
        sections: [
          {
            section_name: 'Policy Documents',
            section_type: 'document',
            sort_order: 0,
            is_enabled: true,
            settings: {},
            documents: {
              items: [
                {
                  display_name: 'Board Policy',
                  description: 'Latest policy',
                  original_file_name: 'policy.pdf',
                  file_name: 'policy.pdf',
                  mime_type: 'application/pdf',
                },
              ],
            },
          },
        ],
      },
      heroImageFile,
      documentFiles: [
        {
          sectionIndex: 0,
          documentIndex: 0,
          file: documentFile,
        },
      ],
    })

    expect(postMock).toHaveBeenCalledTimes(1)
    expect(postMock.mock.calls[0]?.[0]).toBe(API_ROUTES.pages)

    const body = postMock.mock.calls[0]?.[1]
    expect(body).toBeInstanceOf(FormData)
    expect((body as FormData).get('payload')).toBe(
      JSON.stringify({
        page_title: 'Homepage',
        url_slug: '/home',
        parent_id: null,
        status: 'draft',
        hero_image_enabled: true,
        hero_image: {
          file_name: 'hero.png',
          mime_type: 'image/png',
        },
        remove_hero_image: false,
        seo_page_title: 'Homepage SEO',
        seo_page_description: 'Description',
        page_detail: {
          template_key: 'default',
          settings: {},
          sections: [
            {
              section_name: 'Policy Documents',
              section_type: 'document',
              sort_order: 0,
              is_enabled: true,
              settings: {},
              documents: {
                items: [
                  {
                    display_name: 'Board Policy',
                    description: 'Latest policy',
                    original_file_name: 'policy.pdf',
                    file_name: 'policy.pdf',
                    mime_type: 'application/pdf',
                  },
                ],
              },
            },
          ],
        },
      }),
    )
    expect(((body as FormData).get('hero_image_file') as File).name).toBe('hero.png')
    expect(
      ((body as FormData).get('page_detail.sections[0].documents.items[0].file') as File).name,
    ).toBe('policy.pdf')
  })

  it('sends CTA banner image uploads with the section-specific multipart field', async () => {
    const ctaImageFile = new File(['cta-image'], 'cta.png', { type: 'image/png' })

    await pagesApi.createPage({
      page_title: 'Community Support',
      url_slug: '/community-support',
      parent_id: null,
      status: 'draft',
      hero_image_enabled: false,
      remove_hero_image: false,
      seo_page_title: 'Community Support',
      seo_page_description: 'Support details',
      page_detail: {
        template_key: 'default',
        settings: {},
        sections: [
          {
            section_name: 'CTA Banner',
            section_type: 'cta_banner',
            sort_order: 0,
            is_enabled: true,
            settings: {},
            cta_banner: {
              banner_heading: 'We are here for the Community',
              banner_message: 'Our mission is rooted in honouring survivors.',
              button_text: 'Learn more',
              button_url: 'https://example.com/community-support',
              open_in_new_tab: false,
              image: {
                file_name: 'cta.png',
                mime_type: 'image/png',
              },
            },
          },
        ],
      },
      ctaBannerImageFiles: [
        {
          sectionIndex: 0,
          file: ctaImageFile,
        },
      ],
    })

    expect(postMock).toHaveBeenCalledTimes(1)

    const body = postMock.mock.calls[0]?.[1] as FormData
    expect(body).toBeInstanceOf(FormData)
    expect((body.get('page_detail.sections[0].cta_banner.image.file') as File).name).toBe(
      'cta.png',
    )
    expect(body.get('payload')).toBe(
      JSON.stringify({
        page_title: 'Community Support',
        url_slug: '/community-support',
        parent_id: null,
        status: 'draft',
        hero_image_enabled: false,
        remove_hero_image: false,
        seo_page_title: 'Community Support',
        seo_page_description: 'Support details',
        page_detail: {
          template_key: 'default',
          settings: {},
          sections: [
            {
              section_name: 'CTA Banner',
              section_type: 'cta_banner',
              sort_order: 0,
              is_enabled: true,
              settings: {},
              cta_banner: {
                banner_heading: 'We are here for the Community',
                banner_message: 'Our mission is rooted in honouring survivors.',
                button_text: 'Learn more',
                button_url: 'https://example.com/community-support',
                open_in_new_tab: false,
                image: {
                  file_name: 'cta.png',
                  mime_type: 'image/png',
                },
              },
            },
          ],
        },
      }),
    )
  })

  it('sends gallery image uploads as multipart payload with indexed file fields', async () => {
    const bannerFile = new File(['banner'], 'banner.png', { type: 'image/png' })
    const detailFile = new File(['detail'], 'detail.png', { type: 'image/png' })

    await mediaApi.uploadGalleryImages(7, {
      images: [
        {
          title: 'Opening banner',
          alt_text: 'Banner',
          link_url: 'https://partner.example.com',
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
            link_url: 'https://partner.example.com',
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

  it('sends gallery image metadata updates with an optional link_url', async () => {
    patchMock.mockResolvedValue({
      data: {
        message: 'Gallery image updated successfully',
        image: {
          id: 9,
          gallery_id: 7,
          title: 'Opening banner',
          alt_text: 'Banner',
          link_url: 'https://partner.example.com',
          file_name: 'banner.png',
          file_url: '/api/galleries/7/images/9/content',
          storage_uri: 'gs://bucket/galleries/7/images/banner.png',
          mime_type: 'image/png',
          file_size: 1234,
          sort_order: 0,
          created_at: '2026-05-20T12:00:00Z',
          updated_at: '2026-05-20T12:05:00Z',
        },
      },
    })

    await mediaApi.updateGalleryImage(7, 9, {
      title: 'Opening banner',
      alt_text: 'Banner',
      link_url: 'https://partner.example.com',
    })

    expect(patchMock).toHaveBeenCalledTimes(1)
    expect(patchMock).toHaveBeenCalledWith(
      API_ROUTES.galleryImageById(7, 9),
      {
        title: 'Opening banner',
        alt_text: 'Banner',
        link_url: 'https://partner.example.com',
      },
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
  })
})
