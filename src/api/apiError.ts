import axios from 'axios'

type ApiErrorPayload = {
  error?: string
  message?: string
}

const fallbackErrorMessage = 'Something went wrong. Please try again.'

export function getApiErrorMessage(error: unknown) {
  if (axios.isAxiosError<ApiErrorPayload>(error)) {
    return (
      error.response?.data?.error ??
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
