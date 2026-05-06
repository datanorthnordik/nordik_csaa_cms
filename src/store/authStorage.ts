import type { AuthState, AuthSession } from './authSlice'

const AUTH_STORAGE_KEY = 'nordik_csaa_auth'

const emptyAuthState: AuthState = {
  session: null,
  rememberMe: false,
}

function isBrowser() {
  return typeof window !== 'undefined'
}

function isValidSession(value: unknown): value is AuthSession {
  if (!value || typeof value !== 'object') {
    return false
  }

  const session = value as Partial<AuthSession>
  return Boolean(session.accessToken && session.refreshToken && session.email)
}

function parseStoredAuthState(rawValue: string | null): AuthState | null {
  if (!rawValue) {
    return null
  }

  try {
    const parsedValue = JSON.parse(rawValue) as Partial<AuthState>
    if (!isValidSession(parsedValue.session)) {
      return null
    }

    return {
      session: parsedValue.session,
      rememberMe: Boolean(parsedValue.rememberMe),
    }
  } catch {
    return null
  }
}

export function loadStoredAuthState(): AuthState {
  if (!isBrowser()) {
    return emptyAuthState
  }

  const localState = parseStoredAuthState(
    window.localStorage.getItem(AUTH_STORAGE_KEY),
  )
  if (localState) {
    return localState
  }

  const sessionState = parseStoredAuthState(
    window.sessionStorage.getItem(AUTH_STORAGE_KEY),
  )
  return sessionState ?? emptyAuthState
}

export function clearStoredAuthState() {
  if (!isBrowser()) {
    return
  }

  window.localStorage.removeItem(AUTH_STORAGE_KEY)
  window.sessionStorage.removeItem(AUTH_STORAGE_KEY)
}

export function persistAuthState(state: AuthState) {
  if (!isBrowser()) {
    return
  }

  if (!state.session) {
    clearStoredAuthState()
    return
  }

  const payload = JSON.stringify({
    session: state.session,
    rememberMe: state.rememberMe,
  })

  if (state.rememberMe) {
    window.localStorage.setItem(AUTH_STORAGE_KEY, payload)
    window.sessionStorage.removeItem(AUTH_STORAGE_KEY)
    return
  }

  window.sessionStorage.setItem(AUTH_STORAGE_KEY, payload)
  window.localStorage.removeItem(AUTH_STORAGE_KEY)
}
