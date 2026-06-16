import { apiClient } from './apiClient'
import { buildMultipartPayload } from './multipartForm'
import { API_BASE_URL, API_ROUTES } from '../constants/api'

export type BookFieldInputType = 'single_line' | 'rich_text'
export type BookFieldPlacement = 'heading' | 'body'
export type BookSubmissionStatus = 'pending' | 'approved' | 'rejected'

export type BookSummary = {
  id: number
  title: string
  description: string
  activeVersionId?: number
  activeVersionNumber?: number
  versionCount: number
  pendingSubmissionCount: number
  createdAt: string
  updatedAt: string
}

export type BookVersionSummary = {
  id: number
  versionNumber: number
  isActive: boolean
  sourcePageCount: number
  sectionsCount: number
  fieldsCount: number
  approvedSubmissionCount: number
  pendingSubmissionCount: number
  createdAt: string
  updatedAt: string
  lastGeneratedAt?: string
}

export type BookDetail = {
  id: number
  title: string
  description: string
  adminNotificationEmails: string[]
  activeVersionId?: number
  versions: BookVersionSummary[]
  createdAt: string
  updatedAt: string
}

export type BookVersionSection = {
  id: number
  name: string
  sourceStartPage?: number
  sourceEndPage?: number
  currentStartPage: number
  currentEndPage: number
  sortOrder: number
  createdAt: string
  updatedAt: string
}

export type BookVersionField = {
  id: number
  label: string
  inputType: BookFieldInputType
  placement: BookFieldPlacement
  showLabel: boolean
  isRequired: boolean
  isEmailField: boolean
  sortOrder: number
  createdAt: string
  updatedAt: string
}

export type BookSubmissionImage = {
  fileName: string
  mimeType: string
  fileSize: number
  fetchUrl: string
}

export type BookSubmissionValue = {
  fieldId: number
  label: string
  inputType: BookFieldInputType
  placement: BookFieldPlacement
  showLabel: boolean
  isEmailField: boolean
  value: string
}

export type BookSubmission = {
  id: number
  bookId: number
  bookVersionId: number
  targetSectionId?: number
  targetSectionName: string
  newSectionName: string
  status: BookSubmissionStatus
  submitterEmail: string
  image?: BookSubmissionImage
  fieldValues: BookSubmissionValue[]
  reviewedBy?: number
  reviewedAt?: string
  rejectionReason: string
  createdAt: string
  updatedAt: string
}

export type BookVersionDetail = {
  id: number
  bookId: number
  versionNumber: number
  isActive: boolean
  sourcePageCount: number
  contentTemplatePageNumber: number
  sectionTemplatePageNumber: number
  allowPageImage: boolean
  allowNewSections: boolean
  layoutSettings: Record<string, unknown>
  sourcePdfFetchUrl: string
  generatedPdfFetchUrl: string
  sections: BookVersionSection[]
  fields: BookVersionField[]
  approvedSubmissions: BookSubmission[]
  lastGeneratedAt?: string
  createdAt: string
  updatedAt: string
}

export type BookSaveInput = {
  title: string
  description: string
  adminNotificationEmails: string[]
}

export type BookVersionSectionInput = {
  id?: number
  name: string
  sourceStartPage?: number
  sourceEndPage?: number
}

export type BookVersionFieldInput = {
  id?: number
  label: string
  inputType: BookFieldInputType
  placement: BookFieldPlacement
  showLabel: boolean
  isRequired: boolean
  isEmailField: boolean
}

export type BookVersionSaveInput = {
  sourcePageCount: number
  contentTemplatePageNumber: number
  sectionTemplatePageNumber: number
  allowPageImage: boolean
  allowNewSections: boolean
  layoutSettings: Record<string, unknown>
  sections: BookVersionSectionInput[]
  fields: BookVersionFieldInput[]
  activateImmediately: boolean
}

export type BookSubmissionSaveInput = {
  targetSectionId?: number
  newSectionName: string
  fieldValues: Array<{
    fieldId: number
    value: string
  }>
  removeImage?: boolean
}

type ApiListBooksResponse = {
  books: ApiBookSummary[]
}

type ApiBookSummary = {
  id: number
  title: string
  description: string
  active_version_id?: number
  active_version_number?: number
  version_count: number
  pending_submission_count: number
  created_at: string
  updated_at: string
}

type ApiBookDetailResponse = {
  book: ApiBookDetail
}

type ApiBookDetail = {
  id: number
  title: string
  description: string
  admin_notification_emails: string[]
  active_version_id?: number
  versions: ApiBookVersionSummary[]
  created_at: string
  updated_at: string
}

type ApiBookVersionSummary = {
  id: number
  version_number: number
  is_active: boolean
  source_page_count: number
  sections_count: number
  fields_count: number
  approved_submission_count: number
  pending_submission_count: number
  created_at: string
  updated_at: string
  last_generated_at?: string
}

type ApiBookMutationResponse = {
  book: {
    id: number
    title: string
    description: string
    updated_at: string
  }
}

type ApiBookVersionMutationResponse = {
  version: {
    id: number
    book_id: number
    version_number: number
    is_active: boolean
    updated_at: string
  }
}

type ApiBookVersionDetailResponse = {
  version: ApiBookVersionDetail
}

type ApiBookVersionDetail = {
  id: number
  book_id: number
  version_number: number
  is_active: boolean
  source_page_count: number
  content_template_page_number: number
  section_template_page_number: number
  allow_page_image: boolean
  allow_new_sections: boolean
  layout_settings: Record<string, unknown>
  source_pdf_fetch_url: string
  generated_pdf_fetch_url: string
  sections: ApiBookVersionSection[]
  fields: ApiBookVersionField[]
  approved_submissions: ApiBookSubmission[]
  last_generated_at?: string
  created_at: string
  updated_at: string
}

type ApiBookVersionSection = {
  id: number
  name: string
  source_start_page?: number
  source_end_page?: number
  current_start_page: number
  current_end_page: number
  sort_order: number
  created_at: string
  updated_at: string
}

type ApiBookVersionField = {
  id: number
  label: string
  input_type: BookFieldInputType
  placement: BookFieldPlacement
  show_label: boolean
  is_required: boolean
  is_email_field: boolean
  sort_order: number
  created_at: string
  updated_at: string
}

type ApiBookSubmissionImage = {
  file_name: string
  mime_type: string
  file_size: number
  fetch_url: string
}

type ApiBookSubmissionValue = {
  field_id: number
  label: string
  input_type: BookFieldInputType
  placement: BookFieldPlacement
  show_label: boolean
  is_email_field: boolean
  value: string
}

type ApiBookSubmission = {
  id: number
  book_id: number
  book_version_id: number
  target_section_id?: number
  target_section_name: string
  new_section_name: string
  status: BookSubmissionStatus
  submitter_email: string
  image?: ApiBookSubmissionImage
  field_values: ApiBookSubmissionValue[]
  reviewed_by?: number
  reviewed_at?: string
  rejection_reason: string
  created_at: string
  updated_at: string
}

type ApiSubmissionListResponse = {
  submissions: ApiBookSubmission[]
}

type ApiSubmissionMutationResponse = {
  submission: {
    id: number
    status: BookSubmissionStatus
    updated_at: string
  }
}

function mapBookSummary(item: ApiBookSummary): BookSummary {
  return {
    id: item.id,
    title: item.title,
    description: item.description,
    activeVersionId: item.active_version_id,
    activeVersionNumber: item.active_version_number,
    versionCount: item.version_count,
    pendingSubmissionCount: item.pending_submission_count,
    createdAt: item.created_at,
    updatedAt: item.updated_at,
  }
}

function mapBookVersionSummary(item: ApiBookVersionSummary): BookVersionSummary {
  return {
    id: item.id,
    versionNumber: item.version_number,
    isActive: item.is_active,
    sourcePageCount: item.source_page_count,
    sectionsCount: item.sections_count,
    fieldsCount: item.fields_count,
    approvedSubmissionCount: item.approved_submission_count,
    pendingSubmissionCount: item.pending_submission_count,
    createdAt: item.created_at,
    updatedAt: item.updated_at,
    lastGeneratedAt: item.last_generated_at,
  }
}

function mapBookDetail(item: ApiBookDetail): BookDetail {
  return {
    id: item.id,
    title: item.title,
    description: item.description,
    adminNotificationEmails: item.admin_notification_emails ?? [],
    activeVersionId: item.active_version_id,
    versions: (item.versions ?? []).map(mapBookVersionSummary),
    createdAt: item.created_at,
    updatedAt: item.updated_at,
  }
}

function mapSection(item: ApiBookVersionSection): BookVersionSection {
  return {
    id: item.id,
    name: item.name,
    sourceStartPage: item.source_start_page,
    sourceEndPage: item.source_end_page,
    currentStartPage: item.current_start_page,
    currentEndPage: item.current_end_page,
    sortOrder: item.sort_order,
    createdAt: item.created_at,
    updatedAt: item.updated_at,
  }
}

function mapField(item: ApiBookVersionField): BookVersionField {
  return {
    id: item.id,
    label: item.label,
    inputType: item.input_type,
    placement: item.placement,
    showLabel: item.show_label,
    isRequired: item.is_required,
    isEmailField: item.is_email_field,
    sortOrder: item.sort_order,
    createdAt: item.created_at,
    updatedAt: item.updated_at,
  }
}

function mapSubmission(item: ApiBookSubmission): BookSubmission {
  return {
    id: item.id,
    bookId: item.book_id,
    bookVersionId: item.book_version_id,
    targetSectionId: item.target_section_id,
    targetSectionName: item.target_section_name,
    newSectionName: item.new_section_name,
    status: item.status,
    submitterEmail: item.submitter_email,
    image: item.image
      ? {
          fileName: item.image.file_name,
          mimeType: item.image.mime_type,
          fileSize: item.image.file_size,
          fetchUrl: item.image.fetch_url,
        }
      : undefined,
    fieldValues: (item.field_values ?? []).map((value) => ({
      fieldId: value.field_id,
      label: value.label,
      inputType: value.input_type,
      placement: value.placement,
      showLabel: value.show_label,
      isEmailField: value.is_email_field,
      value: value.value,
    })),
    reviewedBy: item.reviewed_by,
    reviewedAt: item.reviewed_at,
    rejectionReason: item.rejection_reason,
    createdAt: item.created_at,
    updatedAt: item.updated_at,
  }
}

function mapVersionDetail(item: ApiBookVersionDetail): BookVersionDetail {
  return {
    id: item.id,
    bookId: item.book_id,
    versionNumber: item.version_number,
    isActive: item.is_active,
    sourcePageCount: item.source_page_count,
    contentTemplatePageNumber: item.content_template_page_number,
    sectionTemplatePageNumber: item.section_template_page_number,
    allowPageImage: item.allow_page_image,
    allowNewSections: item.allow_new_sections,
    layoutSettings: item.layout_settings ?? {},
    sourcePdfFetchUrl: item.source_pdf_fetch_url,
    generatedPdfFetchUrl: item.generated_pdf_fetch_url,
    sections: (item.sections ?? []).map(mapSection),
    fields: (item.fields ?? []).map(mapField),
    approvedSubmissions: (item.approved_submissions ?? []).map(mapSubmission),
    lastGeneratedAt: item.last_generated_at,
    createdAt: item.created_at,
    updatedAt: item.updated_at,
  }
}

function resolveApiPath(pathname: string) {
  const trimmed = pathname.trim()
  if (!trimmed) {
    return trimmed
  }
  if (/^https?:\/\//i.test(trimmed)) {
    return trimmed
  }
  if (trimmed.startsWith('/api/')) {
    return `${API_BASE_URL}${trimmed}`
  }
  return `${API_BASE_URL}/${trimmed.replace(/^\/+/, '')}`
}

function buildVersionPayload(input: BookVersionSaveInput, sourcePdfFile?: File | null) {
  const payload = {
    source_page_count: input.sourcePageCount,
    content_template_page_number: input.contentTemplatePageNumber,
    section_template_page_number: input.sectionTemplatePageNumber,
    allow_page_image: input.allowPageImage,
    allow_new_sections: input.allowNewSections,
    layout_settings: input.layoutSettings,
    sections: input.sections.map((section) => ({
      id: section.id,
      name: section.name,
      source_start_page: section.sourceStartPage,
      source_end_page: section.sourceEndPage,
    })),
    fields: input.fields.map((field) => ({
      id: field.id,
      label: field.label,
      input_type: field.inputType,
      placement: field.placement,
      show_label: field.showLabel,
      is_required: field.isRequired,
      is_email_field: field.isEmailField,
    })),
    activate_immediately: input.activateImmediately,
    source_pdf: sourcePdfFile
      ? {
          file_name: sourcePdfFile.name,
          mime_type: sourcePdfFile.type || 'application/pdf',
          file_size: sourcePdfFile.size,
        }
      : undefined,
  }

  if (!sourcePdfFile) {
    return payload
  }

  return buildMultipartPayload(payload, [
    {
      fieldName: 'source_pdf_file',
      file: sourcePdfFile,
      fileName: sourcePdfFile.name,
    },
  ])
}

function buildSubmissionPayload(input: BookSubmissionSaveInput, imageFile?: File | null) {
  const payload = {
    target_section_id: input.targetSectionId,
    new_section_name: input.newSectionName,
    remove_image: Boolean(input.removeImage),
    field_values: input.fieldValues.map((value) => ({
      field_id: value.fieldId,
      value: value.value,
    })),
    image: imageFile
      ? {
          file_name: imageFile.name,
          mime_type: imageFile.type,
          file_size: imageFile.size,
        }
      : undefined,
  }

  if (!imageFile) {
    return payload
  }

  return buildMultipartPayload(payload, [
    {
      fieldName: 'image_file',
      file: imageFile,
      fileName: imageFile.name,
    },
  ])
}

export const booksApi = {
  async listBooks() {
    const response = await apiClient.get<ApiListBooksResponse>(API_ROUTES.books)
    return (response.data.books ?? []).map(mapBookSummary)
  },

  async getBook(bookId: number) {
    const response = await apiClient.get<ApiBookDetailResponse>(API_ROUTES.bookById(bookId))
    return mapBookDetail(response.data.book)
  },

  async createBook(input: BookSaveInput) {
    const response = await apiClient.post<ApiBookMutationResponse>(API_ROUTES.books, {
      title: input.title,
      description: input.description,
      admin_notification_emails: input.adminNotificationEmails,
    })
    return response.data.book
  },

  async updateBook(bookId: number, input: BookSaveInput) {
    const response = await apiClient.put<ApiBookMutationResponse>(API_ROUTES.bookById(bookId), {
      title: input.title,
      description: input.description,
      admin_notification_emails: input.adminNotificationEmails,
    })
    return response.data.book
  },

  async getVersion(bookId: number, versionId: number) {
    const response = await apiClient.get<ApiBookVersionDetailResponse>(
      API_ROUTES.bookVersionById(bookId, versionId),
    )
    return mapVersionDetail(response.data.version)
  },

  async createVersion(bookId: number, input: BookVersionSaveInput, sourcePdfFile: File) {
    const response = await apiClient.post<ApiBookVersionMutationResponse>(
      API_ROUTES.bookVersions(bookId),
      buildVersionPayload(input, sourcePdfFile),
    )
    return response.data.version
  },

  async updateVersion(
    bookId: number,
    versionId: number,
    input: BookVersionSaveInput,
    sourcePdfFile?: File | null,
  ) {
    const response = await apiClient.put<ApiBookVersionMutationResponse>(
      API_ROUTES.bookVersionById(bookId, versionId),
      buildVersionPayload(input, sourcePdfFile),
    )
    return response.data.version
  },

  async activateVersion(bookId: number, versionId: number) {
    const response = await apiClient.post<ApiBookVersionMutationResponse>(
      API_ROUTES.bookVersionActivate(bookId, versionId),
      {},
    )
    return response.data.version
  },

  async fetchSourcePdfBlob(bookId: number, versionId: number) {
    const response = await apiClient.get<Blob>(API_ROUTES.bookVersionSourceContent(bookId, versionId), {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },

  async uploadGeneratedPdf(bookId: number, versionId: number, pdfBlob: Blob, fileName: string) {
    const file = new File([pdfBlob], fileName, { type: 'application/pdf' })
    const payload = buildMultipartPayload(
      {
        generated_pdf: {
          file_name: file.name,
          mime_type: file.type,
          file_size: file.size,
        },
      },
      [{ fieldName: 'generated_pdf_file', file, fileName: file.name }],
    )

    const response = await apiClient.put<ApiBookVersionMutationResponse>(
      API_ROUTES.bookVersionGenerated(bookId, versionId),
      payload,
    )
    return response.data.version
  },

  async listSubmissions(bookId: number, versionId: number, status = '') {
    const params = new URLSearchParams()
    params.set('version_id', String(versionId))
    if (status) {
      params.set('status', status)
    }

    const response = await apiClient.get<ApiSubmissionListResponse>(API_ROUTES.bookSubmissions(bookId), {
      params,
    })
    return (response.data.submissions ?? []).map(mapSubmission)
  },

  async updateSubmission(
    bookId: number,
    submissionId: number,
    input: BookSubmissionSaveInput,
    imageFile?: File | null,
  ) {
    const response = await apiClient.put<ApiSubmissionMutationResponse>(
      API_ROUTES.bookSubmissionById(bookId, submissionId),
      buildSubmissionPayload(input, imageFile),
    )
    return response.data.submission
  },

  async approveSubmission(bookId: number, submissionId: number) {
    const response = await apiClient.post<ApiSubmissionMutationResponse>(
      API_ROUTES.bookSubmissionApprove(bookId, submissionId),
      {},
    )
    return response.data.submission
  },

  async rejectSubmission(bookId: number, submissionId: number, rejectionReason: string) {
    const response = await apiClient.post<ApiSubmissionMutationResponse>(
      API_ROUTES.bookSubmissionReject(bookId, submissionId),
      { rejection_reason: rejectionReason },
    )
    return response.data.submission
  },

  async fetchSubmissionImage(fetchUrl: string) {
    const response = await apiClient.get<Blob>(resolveApiPath(fetchUrl), {
      responseType: 'blob',
      skipErrorToast: true,
    })
    return response.data
  },
}
