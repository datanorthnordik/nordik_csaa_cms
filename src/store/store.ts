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
import eventsReducer from './eventsSlice'
import mediaReducer from './mediaSlice'
import pagesReducer from './pagesSlice'

const rootReducer = combineReducers({
  auth: authReducer,
  events: eventsReducer,
  media: mediaReducer,
  pages: pagesReducer,
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
      getDefaultMiddleware({
        serializableCheck: {
          ignoredActionPaths: ['meta.arg'],
        },
      }).prepend(authPersistenceListener.middleware),
    preloadedState:
      preloadedState ??
      ({
        auth: loadStoredAuthState(),
      } satisfies Partial<RootState>),
  })
}

export const store = createAppStore()

export type AppStore = ReturnType<typeof createAppStore>
export type AppDispatch = AppStore['dispatch']
