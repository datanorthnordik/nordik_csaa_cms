import axios from 'axios'

type ApiErrorPayload = {
  error?:
    | string
    | {
        code?: string
        message?: string
        details?: Array<{
          field?: string
          message?: string
        }>
      }
  message?: string
}

const fallbackErrorMessage = 'Something went wrong. Please try again.'

export function getApiErrorMessage(error: unknown) {
  if (axios.isAxiosError<ApiErrorPayload>(error)) {
    const errorPayload = error.response?.data?.error
    if (typeof errorPayload === 'object' && errorPayload) {
      const detailMessage = errorPayload.details
        ?.map((detail) => detail.message)
        .filter(Boolean)
        .join(', ')

      return detailMessage
        ? `${errorPayload.message ?? fallbackErrorMessage}: ${detailMessage}`
        : (errorPayload.message ?? fallbackErrorMessage)
    }

    return (
      (typeof errorPayload === 'string' ? errorPayload : undefined) ??
      error.response?.data?.message ??
      error.message ??
      fallbackErrorMessage
    )
  }

  if (error instanceof Error) {
    return error.message
  }

  return fallbackErrorMessage
}
