import { API_ROUTES } from '../constants/api'
import { apiClient } from './apiClient'

export type KnowledgeCenterSubmissionType = 'post' | 'video' | 'both'
export type KnowledgeCenterSubmissionStatus = 'open' | 'completed'

export type KnowledgeCenterCompletedBy = {
  id: number
  name: string
  email: string
}

export type KnowledgeCenterSubmission = {
  id: number
  submitterName: string
  submitterEmail: string
  submitterPhone: string
  submissionType: KnowledgeCenterSubmissionType
  message: string
  status: KnowledgeCenterSubmissionStatus
  completionNotes: string
  completedBy?: KnowledgeCenterCompletedBy
  completedAt?: string
  createdAt: string
  updatedAt: string
}

export type KnowledgeCenterListFilters = {
  page: number
  pageSize: number
  searchTerm: string
  status: KnowledgeCenterSubmissionStatus
}

export type KnowledgeCenterListPageMeta = {
  page: number
  pageSize: number
  totalItems: number
  totalPages: number
  hasNext: boolean
  hasPrev: boolean
}

export type KnowledgeCenterListSummary = {
  openCount: number
  completedCount: number
}

type KnowledgeCenterApiSubmission = {
  id: number
  submitter_name: string
  submitter_email: string
  submitter_phone: string
  submission_type: KnowledgeCenterSubmissionType
  message: string
  status: KnowledgeCenterSubmissionStatus
  completion_notes: string
  completed_by?: KnowledgeCenterCompletedBy
  completed_at?: string
  created_at: string
  updated_at: string
}

type KnowledgeCenterApiListResponse = {
  items: KnowledgeCenterApiSubmission[]
  pagination: {
    page: number
    page_size: number
    total_items: number
    total_pages: number
    has_next: boolean
    has_prev: boolean
  }
  summary: {
    open_count: number
    completed_count: number
  }
  applied_filters: {
    page: number
    page_size: number
    search_term: string
    status: KnowledgeCenterSubmissionStatus
  }
}

type KnowledgeCenterApiDetailResponse = {
  submission: KnowledgeCenterApiSubmission
}

type KnowledgeCenterApiMutationResponse = {
  message: string
  submission: KnowledgeCenterApiSubmission
}

function buildListQuery(filters: KnowledgeCenterListFilters) {
  const params = new URLSearchParams()
  params.set('page', String(filters.page))
  params.set('page_size', String(filters.pageSize))
  params.set('status', filters.status)

  if (filters.searchTerm.trim()) {
    params.set('search', filters.searchTerm.trim())
  }

  return params
}

function mapSubmission(
  submission: KnowledgeCenterApiSubmission,
): KnowledgeCenterSubmission {
  return {
    id: submission.id,
    submitterName: submission.submitter_name,
    submitterEmail: submission.submitter_email,
    submitterPhone: submission.submitter_phone,
    submissionType: submission.submission_type,
    message: submission.message,
    status: submission.status,
    completionNotes: submission.completion_notes,
    completedBy: submission.completed_by,
    completedAt: submission.completed_at,
    createdAt: submission.created_at,
    updatedAt: submission.updated_at,
  }
}

export const knowledgeCenterApi = {
  async listSubmissions(filters: KnowledgeCenterListFilters) {
    const response = await apiClient.get<KnowledgeCenterApiListResponse>(
      API_ROUTES.knowledgeCenterSubmissions,
      {
        params: buildListQuery(filters),
        skipErrorToast: true,
      },
    )

    return {
      items: (response.data.items ?? []).map(mapSubmission),
      pagination: {
        page: response.data.pagination.page,
        pageSize: response.data.pagination.page_size,
        totalItems: response.data.pagination.total_items,
        totalPages: response.data.pagination.total_pages,
        hasNext: response.data.pagination.has_next,
        hasPrev: response.data.pagination.has_prev,
      } satisfies KnowledgeCenterListPageMeta,
      summary: {
        openCount: response.data.summary.open_count,
        completedCount: response.data.summary.completed_count,
      } satisfies KnowledgeCenterListSummary,
      appliedFilters: {
        page: response.data.applied_filters.page,
        pageSize: response.data.applied_filters.page_size,
        searchTerm: response.data.applied_filters.search_term,
        status: response.data.applied_filters.status,
      },
    }
  },

  async getSubmission(id: number) {
    const response = await apiClient.get<KnowledgeCenterApiDetailResponse>(
      API_ROUTES.knowledgeCenterSubmissionById(id),
      {
        skipErrorToast: true,
      },
    )

    return mapSubmission(response.data.submission)
  },

  async completeSubmission(id: number, completionNotes: string) {
    const response = await apiClient.post<KnowledgeCenterApiMutationResponse>(
      API_ROUTES.knowledgeCenterSubmissionComplete(id),
      {
        completion_notes: completionNotes,
      },
      {
        skipErrorToast: true,
      },
    )

    return mapSubmission(response.data.submission)
  },
}
