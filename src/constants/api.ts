export const API_BASE_URL =
  'https://nordikcsaaapi-724838782318.us-west1.run.app'

export const API_ROUTES = {
  login: '/api/user/login',
  signup: '/api/user/signup',
  refresh: '/api/user/refresh',
  forgotPassword: '/api/user/forgot-password',
  resetPassword: '/api/user/reset-password',
  events: '/api/events',
  eventLocations: '/api/events/locations',
  eventGalleries: '/api/events/galleries',
  eventById: (id: number | string) => `/api/events/${id}`,
  eventMediaById: (id: number | string, mediaId: number | string) =>
    `/api/events/${id}/media/${mediaId}/content`,
  press: '/api/press',
  pressById: (id: string) => `/api/press/${id}`,
  pressCoverById: (id: number | string) => `/api/press/${id}/cover/content`,
  resources: '/api/resources',
  resourceById: (id: number | string) => `/api/resources/${id}`,
  resourceContentById: (id: number | string) => `/api/resources/${id}/content`,
  newsletters: '/api/newsletters',
  newsletterById: (id: number | string) => `/api/newsletters/${id}`,
  newsletterMediaById: (id: number | string, mediaId: number | string) =>
    `/api/newsletters/${id}/media/${mediaId}/content`,
  eventDocumentById: (id: number | string) => `/api/events/${id}/document`,
  eventPhotoById: (id: number | string) => `/api/events/${id}/photo`,
  pages: '/api/pages',
  pageById: (id: number | string) => `/api/pages/${id}`,
  pageHeroById: (id: number | string) => `/api/pages/${id}/hero/content`,
  pageDocumentById: (id: number | string) => `/api/pages/documents/${id}/content`,
  menuByKey: (key: string) => `/api/menus/${key}`,
  menuPageOptionsByKey: (key: string) => `/api/menus/${key}/page-options`,
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
