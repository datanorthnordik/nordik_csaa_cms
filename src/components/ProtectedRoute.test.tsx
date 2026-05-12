import { describe, it, expect, vi } from 'vitest'
import { render, screen } from '@testing-library/react'
import { BrowserRouter } from 'react-router-dom'
import { Provider } from 'react-redux'
import { configureStore } from '@reduxjs/toolkit'
import { ProtectedRoute } from './ProtectedRoute'
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
      },
    },
  })

const TestComponent = () => <div>Protected Content</div>

describe('ProtectedRoute', () => {
  it('renders protected content when user is authenticated', () => {
    const store = createMockStore(true)
    render(
      <Provider store={store}>
        <BrowserRouter>
          <ProtectedRoute>
            <TestComponent />
          </ProtectedRoute>
        </BrowserRouter>
      </Provider>
    )

    expect(screen.getByText('Protected Content')).toBeInTheDocument()
  })

  it('redirects to login when user is not authenticated', () => {
    const store = createMockStore(false)
    render(
      <Provider store={store}>
        <BrowserRouter>
          <ProtectedRoute>
            <TestComponent />
          </ProtectedRoute>
        </BrowserRouter>
      </Provider>
    )

    expect(screen.queryByText('Protected Content')).not.toBeInTheDocument()
  })
})
