export const API_BASE_URL =
  'https://nordikcsaaapi-724838782318.us-west1.run.app'

export const API_ROUTES = {
  login: '/api/user/login',
  signup: '/api/user/signup',
  refresh: '/api/user/refresh',
  events: '/api/events',
  eventLocations: '/api/events/locations',
  eventGalleries: '/api/events/galleries',
  eventById: (id: number | string) => `/api/events/${id}`,
  eventMediaById: (id: number | string, mediaId: number | string) =>
    `/api/events/${id}/media/${mediaId}/content`,
  eventDocumentById: (id: number | string) => `/api/events/${id}/document`,
  eventPhotoById: (id: number | string) => `/api/events/${id}/photo`,
  pages: '/api/pages',
  pageById: (id: number | string) => `/api/pages/${id}`,
  pageHeroById: (id: number | string) => `/api/pages/${id}/hero/content`,
  galleries: '/api/galleries',
  galleryById: (id: number | string) => `/api/galleries/${id}`,
  galleryCoverById: (id: number | string) => `/api/galleries/${id}/cover/content`,
  galleryImagesById: (id: number | string) => `/api/galleries/${id}/images`,
  galleryImageById: (id: number | string, imageId: number | string) =>
    `/api/galleries/${id}/images/${imageId}`,
  galleryImageContentById: (id: number | string, imageId: number | string) =>
    `/api/galleries/${id}/images/${imageId}/content`,
  galleryImageOrderById: (id: number | string) =>
    `/api/galleries/${id}/images/order`,
} as const
