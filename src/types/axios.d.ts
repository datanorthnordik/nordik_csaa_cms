import 'axios'

declare module 'axios' {
  export interface AxiosRequestConfig<_D = unknown> {
    skipAuth?: boolean
    skipErrorToast?: boolean
    _retry?: boolean
  }

  export interface InternalAxiosRequestConfig<_D = unknown> {
    skipAuth?: boolean
    skipErrorToast?: boolean
    _retry?: boolean
  }
}
