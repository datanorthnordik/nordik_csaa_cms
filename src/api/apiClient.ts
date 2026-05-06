import axios, {
  AxiosHeaders,
  type AxiosError,
  type RawAxiosHeaders,
  type RawAxiosRequestHeaders,
} from 'axios'
import toast from 'react-hot-toast'
import { API_BASE_URL, API_ROUTES } from '../constants/api'
import {
  clearAuthSession,
  selectAccessToken,
  selectRefreshToken,
  updateAccessToken,
} from '../store/authSlice'
import { store } from '../store/store'
import { getApiErrorMessage } from './apiError'

type RefreshResponse = {
  message: string
  accessToken: string
}

const apiClient = axios.create({
  baseURL: API_BASE_URL,
  headers: {
    'Content-Type': 'application/json',
  },
})

let refreshPromise: Promise<string> | null = null

function setAuthorizationHeader(
  headers:
    | AxiosHeaders
    | RawAxiosHeaders
    | RawAxiosRequestHeaders
    | Partial<RawAxiosRequestHeaders>
    | undefined,
  token: string,
) {
  const nextHeaders = AxiosHeaders.from(
    headers as AxiosHeaders | RawAxiosHeaders | undefined,
  )
  nextHeaders.set('Authorization', `Bearer ${token}`)
  return nextHeaders
}

function shouldSkipAuth(url?: string, skipAuth?: boolean) {
  if (skipAuth) {
    return true
  }

  return url === API_ROUTES.login || url === API_ROUTES.signup
}

async function refreshAccessToken() {
  const refreshToken = selectRefreshToken(store.getState())
  if (!refreshToken) {
    store.dispatch(clearAuthSession())
    throw new Error('Your session has expired. Please sign in again.')
  }

  if (!refreshPromise) {
    refreshPromise = axios
      .post<RefreshResponse>(`${API_BASE_URL}${API_ROUTES.refresh}`, null, {
        headers: {
          Authorization: `Bearer ${refreshToken}`,
        },
      })
      .then(({ data }) => {
        store.dispatch(updateAccessToken(data.accessToken))
        return data.accessToken
      })
      .catch((error) => {
        store.dispatch(clearAuthSession())
        throw error
      })
      .finally(() => {
        refreshPromise = null
      })
  }

  return refreshPromise
}

apiClient.interceptors.request.use((config) => {
  if (shouldSkipAuth(config.url, config.skipAuth)) {
    return config
  }

  const accessToken = selectAccessToken(store.getState())
  if (!accessToken) {
    return config
  }

  config.headers = setAuthorizationHeader(config.headers, accessToken)
  return config
})

apiClient.interceptors.response.use(
  (response) => response,
  async (error: AxiosError) => {
    const originalRequest = error.config

    if (!originalRequest) {
      toast.error(getApiErrorMessage(error))
      return Promise.reject(error)
    }

    if (
      error.response?.status === 401 &&
      !originalRequest._retry &&
      !shouldSkipAuth(originalRequest.url, originalRequest.skipAuth) &&
      originalRequest.url !== API_ROUTES.refresh
    ) {
      originalRequest._retry = true

      try {
        const nextAccessToken = await refreshAccessToken()
        originalRequest.headers = setAuthorizationHeader(
          originalRequest.headers,
          nextAccessToken,
        )
        return apiClient(originalRequest)
      } catch (refreshError) {
        if (!originalRequest.skipErrorToast) {
          toast.error(getApiErrorMessage(refreshError))
        }
        return Promise.reject(refreshError)
      }
    }

    if (!originalRequest.skipErrorToast) {
      toast.error(getApiErrorMessage(error))
    }

    return Promise.reject(error)
  },
)

export { apiClient }
