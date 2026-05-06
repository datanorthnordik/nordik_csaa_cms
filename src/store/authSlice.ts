import { createSlice, type PayloadAction } from '@reduxjs/toolkit'

export type AuthSession = {
  accessToken: string
  refreshToken: string
  firstname: string
  lastname: string
  id: number
  email: string
  role: string
}

export type AuthState = {
  session: AuthSession | null
  rememberMe: boolean
}

const initialState: AuthState = {
  session: null,
  rememberMe: false,
}

const authSlice = createSlice({
  name: 'auth',
  initialState,
  reducers: {
    setAuthSession(
      state,
      action: PayloadAction<{ session: AuthSession; rememberMe: boolean }>,
    ) {
      state.session = action.payload.session
      state.rememberMe = action.payload.rememberMe
    },
    updateAccessToken(state, action: PayloadAction<string>) {
      if (!state.session) {
        return
      }

      state.session.accessToken = action.payload
    },
    clearAuthSession(state) {
      state.session = null
      state.rememberMe = false
    },
  },
})

export const { setAuthSession, updateAccessToken, clearAuthSession } =
  authSlice.actions

export const selectAuthSession = (state: { auth: AuthState }) => state.auth.session
export const selectAccessToken = (state: { auth: AuthState }) =>
  state.auth.session?.accessToken ?? null
export const selectRefreshToken = (state: { auth: AuthState }) =>
  state.auth.session?.refreshToken ?? null
export const selectIsAuthenticated = (state: { auth: AuthState }) =>
  Boolean(state.auth.session?.accessToken)

export default authSlice.reducer
