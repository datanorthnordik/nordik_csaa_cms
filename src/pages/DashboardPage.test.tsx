import { describe, it, expect, vi, beforeEach } from 'vitest'
import { render, screen, waitFor } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { Provider } from 'react-redux'
import { BrowserRouter } from 'react-router-dom'
import { configureStore } from '@reduxjs/toolkit'
import { DashboardPage } from './DashboardPage'
import authReducer from '../store/authSlice'

const createMockStore = () =>
  configureStore({
    reducer: {
      auth: authReducer,
    },
    preloadedState: {
      auth: {
        user: {
          id: 1,
          email: 'test@example.com',
          name: 'Test User',
        },
        status: 'idle',
        error: null,
      },
    },
  })

const renderPage = () => {
  const store = createMockStore()
  return render(
    <Provider store={store}>
      <BrowserRouter>
        <DashboardPage />
      </BrowserRouter>
    </Provider>
  )
}

describe('DashboardPage', () => {
  it('renders dashboard page', () => {
    renderPage()

    expect(screen.getByRole('main')).toBeInTheDocument()
  })

  it('displays welcome message', () => {
    renderPage()

    expect(screen.getByText(/dashboard|welcome/i)).toBeInTheDocument()
  })

  it('displays navigation links to main features', async () => {
    renderPage()

    await waitFor(() => {
      const mainElement = screen.getByRole('main')
      expect(mainElement).toBeInTheDocument()
    })
  })
})
