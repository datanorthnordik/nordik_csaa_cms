import { API_ROUTES } from '../constants/api'
import type { AuthSession } from '../store/authSlice'
import { apiClient } from './apiClient'

export type LoginRequest = {
  email: string
  password: string
  rememberMe: boolean
}

export type SignupRequest = {
  firstname: string
  lastname: string
  email: string
  password: string
}

export type LoginResponse = {
  message: string
  data: AuthSession
}

export type SignupResponse = {
  message: string
  user: {
    id: number
    firstname: string
    lastname: string
    email: string
    role: string
  }
}

export const authApi = {
  async login(payload: LoginRequest) {
    const response = await apiClient.post<LoginResponse>(API_ROUTES.login, payload, {
      skipAuth: true,
    })
    return response.data
  },
  async signup(payload: SignupRequest) {
    const response = await apiClient.post<SignupResponse>(
      API_ROUTES.signup,
      payload,
      {
        skipAuth: true,
      },
    )
    return response.data
  },
}
