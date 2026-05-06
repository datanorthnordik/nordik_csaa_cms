import 'axios'

declare module 'axios' {
  export interface AxiosRequestConfig<D = any> {
    skipAuth?: boolean
    skipErrorToast?: boolean
    _retry?: boolean
  }

  export interface InternalAxiosRequestConfig<D = any> {
    skipAuth?: boolean
    skipErrorToast?: boolean
    _retry?: boolean
  }
}
