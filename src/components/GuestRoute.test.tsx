import { describe, it, expect } from 'vitest'
import { render, screen } from '@testing-library/react'
import { BrowserRouter, Routes, Route } from 'react-router-dom'
import { Provider } from 'react-redux'
import { configureStore } from '@reduxjs/toolkit'
import { GuestRoute } from './GuestRoute'
import authReducer from '../store/authSlice'

const createMockStore = (isAuthenticated: boolean) =>
  configureStore({
    reducer: {
      auth: authReducer,
    },
    preloadedState: {
      auth: {
        user: isAuthenticated ? { id: 1, email: 'test@example.com' } : null,
        status: 'idle',
        error: null,
        session: null,
        rememberMe: false,
      },
    },
  })

describe('GuestRoute', () => {
  it('renders outlet when user is not authenticated', () => {
    const store = createMockStore(false)
    render(
      <Provider store={store}>
        <BrowserRouter>
          <Routes>
            <Route element={<GuestRoute />}>
              <Route index element={<div>Guest Content</div>} />
            </Route>
          </Routes>
        </BrowserRouter>
      </Provider>
    )

    expect(screen.getByText('Guest Content')).toBeInTheDocument()
  })

  it('redirects when user is authenticated', () => {
    const store = createMockStore(true)
    render(
      <Provider store={store}>
        <BrowserRouter>
          <Routes>
            <Route element={<GuestRoute />}>
              <Route index element={<div>Guest Content</div>} />
            </Route>
            <Route path="/events" element={<div>Events Page</div>} />
          </Routes>
        </BrowserRouter>
      </Provider>
    )

    expect(screen.getByText('Events Page')).toBeInTheDocument()
  })
})
