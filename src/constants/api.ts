const DEFAULT_API_BASE_URL =
  'https://nordikcsaaapi-724838782318.us-west1.run.app'

const normalizeApiBaseUrl = (value: string | undefined) => {
  const trimmed = value?.trim()
  if (!trimmed) {
    return undefined
  }

  return trimmed.replace(/\/+$/, '')
}

const getRuntimeApiBaseUrl = () => {
  if (typeof window === 'undefined') {
    return undefined
  }

  return window.__APP_CONFIG__?.API_BASE_URL
}

export const resolveApiBaseUrl = () =>
  normalizeApiBaseUrl(getRuntimeApiBaseUrl()) ??
  normalizeApiBaseUrl(import.meta.env.VITE_API_BASE_URL) ??
  DEFAULT_API_BASE_URL

export const API_BASE_URL = resolveApiBaseUrl()

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
  blogs: '/api/blogs',
  blogById: (id: number | string) => `/api/blogs/${id}`,
  blogCoverById: (id: number | string) => `/api/blogs/${id}/cover/content`,
  resources: '/api/resources',
  resourceById: (id: number | string) => `/api/resources/${id}`,
  resourceContentById: (id: number | string) => `/api/resources/${id}/content`,
  bookshelf: '/api/bookshelf',
  bookshelfById: (id: number | string) => `/api/bookshelf/${id}`,
  bookshelfBookContentById: (id: number | string) => `/api/bookshelf/${id}/book/content`,
  bookshelfAuthorImageContentById: (id: number | string) =>
    `/api/bookshelf/${id}/author-image/content`,
  bookshelfCoverContentById: (id: number | string) => `/api/bookshelf/${id}/cover/content`,
  knowledgeCenterSubmissions: '/api/knowledge-center/submissions',
  knowledgeCenterSubmissionById: (id: number | string) =>
    `/api/knowledge-center/submissions/${id}`,
  knowledgeCenterSubmissionComplete: (id: number | string) =>
    `/api/knowledge-center/submissions/${id}/complete`,
  memorial: '/api/memorial',
  memorialById: (id: number | string) => `/api/memorial/${id}`,
  memorialPortraitById: (id: number | string) => `/api/memorial/${id}/portrait/content`,
  memorialGalleryImageContentById: (id: number | string, imageId: number | string) =>
    `/api/memorial/${id}/gallery/${imageId}/content`,
  newsletters: '/api/newsletters',
  newsletterById: (id: number | string) => `/api/newsletters/${id}`,
  newsletterMediaById: (id: number | string, mediaId: number | string) =>
    `/api/newsletters/${id}/media/${mediaId}/content`,
  books: '/api/books',
  bookById: (id: number | string) => `/api/books/${id}`,
  bookVersions: (bookId: number | string) => `/api/books/${bookId}/versions`,
  bookVersionById: (bookId: number | string, versionId: number | string) =>
    `/api/books/${bookId}/versions/${versionId}`,
  bookVersionActivate: (bookId: number | string, versionId: number | string) =>
    `/api/books/${bookId}/versions/${versionId}/activate`,
  bookVersionGenerated: (bookId: number | string, versionId: number | string) =>
    `/api/books/${bookId}/versions/${versionId}/generated`,
  bookVersionSourceContent: (bookId: number | string, versionId: number | string) =>
    `/api/books/${bookId}/versions/${versionId}/source/content`,
  bookVersionGeneratedContent: (bookId: number | string, versionId: number | string) =>
    `/api/books/${bookId}/versions/${versionId}/generated/content`,
  bookSubmissions: (bookId: number | string) => `/api/books/${bookId}/submissions`,
  bookSubmissionById: (bookId: number | string, submissionId: number | string) =>
    `/api/books/${bookId}/submissions/${submissionId}`,
  bookSubmissionApprove: (bookId: number | string, submissionId: number | string) =>
    `/api/books/${bookId}/submissions/${submissionId}/approve`,
  bookSubmissionReject: (bookId: number | string, submissionId: number | string) =>
    `/api/books/${bookId}/submissions/${submissionId}/reject`,
  bookSubmissionImageContent: (bookId: number | string, submissionId: number | string) =>
    `/api/books/${bookId}/submissions/${submissionId}/image/content`,
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
  videos: '/api/videos',
  videoById: (id: number | string) => `/api/videos/${id}`,
  videoItemById: (id: number | string, itemId: number | string) =>
    `/api/videos/${id}/items/${itemId}`,
  videoItemTeaserById: (id: number | string, itemId: number | string) =>
    `/api/videos/${id}/items/${itemId}/teaser/content`,
} as const
