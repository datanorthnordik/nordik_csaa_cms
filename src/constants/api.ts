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
  eventDocumentById: (id: number | string, mediaId: number | string) =>
    `/api/events/${id}/documents/${mediaId}`,
  eventPhotoById: (id: number | string, mediaId: number | string) =>
    `/api/events/${id}/photos/${mediaId}`,
} as const
