import {
  combineReducers,
  configureStore,
  createListenerMiddleware,
  isAnyOf,
} from '@reduxjs/toolkit'
import authReducer, {
  clearAuthSession,
  setAuthSession,
  updateAccessToken,
} from './authSlice'
import { loadStoredAuthState, persistAuthState } from './authStorage'

const rootReducer = combineReducers({
  auth: authReducer,
})

export type RootState = ReturnType<typeof rootReducer>

const authPersistenceListener = createListenerMiddleware<RootState>()

authPersistenceListener.startListening({
  matcher: isAnyOf(setAuthSession, updateAccessToken, clearAuthSession),
  effect: (_, listenerApi) => {
    persistAuthState(listenerApi.getState().auth)
  },
})

export function createAppStore(preloadedState?: Partial<RootState>) {
  return configureStore({
    reducer: rootReducer,
    middleware: (getDefaultMiddleware) =>
      getDefaultMiddleware().prepend(authPersistenceListener.middleware),
    preloadedState:
      preloadedState ??
      ({
      auth: loadStoredAuthState(),
      } satisfies RootState),
  })
}

export const store = createAppStore()

export type AppStore = ReturnType<typeof createAppStore>
export type AppDispatch = AppStore['dispatch']
