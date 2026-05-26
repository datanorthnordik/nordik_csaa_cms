import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import type { ReactElement } from 'react'
import { Provider } from 'react-redux'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import App from './App'
import i18n from './i18n'
import { ForgotPasswordPage } from './pages/ForgotPasswordPage'
import { LoginPage } from './pages/LoginPage'
import { ResetPasswordPage } from './pages/ResetPasswordPage'
import { SignupPage } from './pages/SignupPage'
import { createAppStore } from './store/store'

function renderWithProviders(ui: ReactElement) {
  const store = createAppStore()
  return render(<Provider store={store}>{ui}</Provider>)
}

function renderPage(ui: ReactElement) {
  return renderWithProviders(<MemoryRouter>{ui}</MemoryRouter>)
}

beforeEach(async () => {
  window.localStorage.clear()
  window.sessionStorage.clear()
  window.history.pushState({}, '', '/')
  await i18n.changeLanguage('en')
})

describe('App Component', () => {
  it('renders the login page by default', () => {
    renderWithProviders(<App />)
    expect(
      screen.getByRole('heading', { name: /sign in with email address/i })
    ).toBeDefined()
    expect(screen.getByRole('button', { name: 'FR' })).toBeDefined()
  })

  it('renders the signup page when navigating to /signup', () => {
    window.history.pushState({}, '', '/signup')
    renderWithProviders(<App />)
    expect(
      screen.getByRole('heading', { name: /create your account/i })
    ).toBeDefined()
  })

  it('renders the forgot password page when navigating to /forgot-password', () => {
    window.history.pushState({}, '', '/forgot-password')
    renderWithProviders(<App />)
    expect(
      screen.getByRole('heading', { name: /forgot your password/i })
    ).toBeDefined()
  })

  it('renders the reset password page when navigating to /reset-password', () => {
    window.history.pushState({}, '', '/reset-password?token=abc123')
    renderWithProviders(<App />)
    expect(
      screen.getByRole('heading', { name: /set a new password/i })
    ).toBeDefined()
  })
})

describe('LoginPage', () => {
  it('shows field errors when the form is submitted without credentials', async () => {
    renderPage(<LoginPage />)

    fireEvent.click(screen.getByRole('button', { name: /sign in/i }))

    await screen.findByText(/email is required/i)
    expect(screen.getByText(/password is required/i)).toBeDefined()
  })

  it('shows field error when the email address is invalid', async () => {
    renderPage(<LoginPage />)

    fireEvent.change(screen.getByPlaceholderText(/email address/i), {
      target: { value: 'invalid-email' },
    })
    fireEvent.change(screen.getByPlaceholderText(/password/i), {
      target: { value: 'Password123!' },
    })

    fireEvent.click(screen.getByRole('button', { name: /sign in/i }))

    await screen.findByText(/please enter a valid email address/i)
    expect(screen.getByText(/please enter a valid email address/i)).toBeDefined()
  })

  it('submits successfully with valid credentials', async () => {
    const onSubmit = vi.fn().mockResolvedValue(undefined)
    renderPage(<LoginPage onSubmit={onSubmit} />)

    fireEvent.change(screen.getByPlaceholderText(/email address/i), {
      target: { value: 'user@example.com' },
    })
    fireEvent.change(screen.getByPlaceholderText(/password/i), {
      target: { value: 'Password123!' },
    })

    fireEvent.click(screen.getByRole('button', { name: /sign in/i }))

    await screen.findByRole('alert')
    expect(screen.getByRole('alert').textContent).toMatch(/signing you in/i)
    expect(onSubmit).toHaveBeenCalledWith({
      email: 'user@example.com',
      password: 'Password123!',
      remember: false,
    })
  })

  it('switches the login page copy to french from the header toggle', async () => {
    renderPage(<LoginPage />)

    fireEvent.click(screen.getByRole('button', { name: 'FR' }))

    expect(
      await screen.findByRole('heading', {
        name: /se connecter avec une adresse courriel/i,
      })
    ).toBeDefined()
    expect(
      screen.getByRole('button', { name: /creer un nouveau compte/i })
    ).toBeDefined()
  })
})

describe('SignupPage', () => {
  it('shows field errors when required fields are missing', async () => {
    renderPage(<SignupPage />)

    fireEvent.click(screen.getByRole('button', { name: /create account/i }))

    await screen.findByText(/first name is required/i)
    expect(screen.getByText(/email is required/i)).toBeDefined()
    expect(screen.getByText(/password is required/i)).toBeDefined()
  })

  it('shows field error when the email address is invalid', async () => {
    renderPage(<SignupPage />)

    fireEvent.change(screen.getByLabelText(/first name/i), {
      target: { value: 'Jane' },
    })
    fireEvent.change(screen.getByLabelText(/last name/i), {
      target: { value: 'Doe' },
    })
    fireEvent.change(screen.getByLabelText(/work email address/i), {
      target: { value: 'invalid-email' },
    })
    fireEvent.change(screen.getByLabelText(/^password$/i), {
      target: { value: 'Password123!' },
    })
    fireEvent.change(screen.getByLabelText(/confirm password/i), {
      target: { value: 'Password123!' },
    })

    fireEvent.click(screen.getByRole('button', { name: /create account/i }))

    await screen.findByText(/please enter a valid work email address/i)
    expect(
      screen.getByText(/please enter a valid work email address/i)
    ).toBeDefined()
  })

  it('shows field error when passwords do not match', async () => {
    renderPage(<SignupPage />)

    fireEvent.change(screen.getByLabelText(/first name/i), {
      target: { value: 'Jane' },
    })
    fireEvent.change(screen.getByLabelText(/last name/i), {
      target: { value: 'Doe' },
    })
    fireEvent.change(screen.getByLabelText(/work email address/i), {
      target: { value: 'jane.doe@example.com' },
    })
    fireEvent.change(screen.getByLabelText(/^password$/i), {
      target: { value: 'Password123!' },
    })
    fireEvent.change(screen.getByLabelText(/confirm password/i), {
      target: { value: 'Different123!' },
    })

    fireEvent.click(screen.getByRole('button', { name: /create account/i }))

    await screen.findByText(/passwords do not match/i)
    expect(screen.getByText(/passwords do not match/i)).toBeDefined()
  })

  it('shows field error when password is too short', async () => {
    renderPage(<SignupPage />)

    fireEvent.change(screen.getByLabelText(/first name/i), {
      target: { value: 'Jane' },
    })
    fireEvent.change(screen.getByLabelText(/last name/i), {
      target: { value: 'Doe' },
    })
    fireEvent.change(screen.getByLabelText(/work email address/i), {
      target: { value: 'jane.doe@example.com' },
    })
    fireEvent.change(screen.getByLabelText(/^password$/i), {
      target: { value: '123' },
    })
    fireEvent.change(screen.getByLabelText(/confirm password/i), {
      target: { value: '123' },
    })

    fireEvent.click(screen.getByRole('button', { name: /create account/i }))

    await screen.findByText(/password must be at least 6 characters/i)
    expect(screen.getByText(/password must be at least 6 characters/i)).toBeDefined()
  })

  it('submits valid signup data and displays a success message', async () => {
    const onSubmit = vi.fn().mockResolvedValue(undefined)
    renderPage(<SignupPage onSubmit={onSubmit} />)

    fireEvent.change(screen.getByLabelText(/first name/i), {
      target: { value: 'Jane' },
    })
    fireEvent.change(screen.getByLabelText(/last name/i), {
      target: { value: 'Doe' },
    })
    fireEvent.change(screen.getByLabelText(/work email address/i), {
      target: { value: 'jane.doe@example.com' },
    })
    fireEvent.change(screen.getByLabelText(/^password$/i), {
      target: { value: 'Password123!' },
    })
    fireEvent.change(screen.getByLabelText(/confirm password/i), {
      target: { value: 'Password123!' },
    })

    fireEvent.click(screen.getByRole('button', { name: /create account/i }))

    await waitFor(() => expect(onSubmit).toHaveBeenCalledTimes(1))
    expect(onSubmit).toHaveBeenCalledWith({
      firstName: 'Jane',
      lastName: 'Doe',
      email: 'jane.doe@example.com',
      password: 'Password123!',
      confirmPassword: 'Password123!',
    })

    await screen.findByRole('alert')
    expect(screen.getByRole('alert').textContent).toMatch(
      /account request submitted/i
    )
  })

  it('navigates to login when the sign in button is clicked', () => {
    renderWithProviders(
      <MemoryRouter initialEntries={['/signup']}>
        <SignupPage />
      </MemoryRouter>,
    )

    expect(screen.getByRole('button', { name: /sign in/i })).toBeDefined()
  })

  it('shows the language toggle on the signup page', () => {
    renderPage(<SignupPage />)

    expect(screen.getByRole('button', { name: 'EN' })).toBeDefined()
    expect(screen.getByRole('button', { name: 'FR' })).toBeDefined()
  })
})

describe('ForgotPasswordPage', () => {
  it('shows a field error when the email address is missing', async () => {
    renderPage(<ForgotPasswordPage />)

    fireEvent.click(screen.getByRole('button', { name: /send reset code/i }))

    await screen.findByText(/email is required/i)
  })

  it('navigates to reset password after submitting a valid email address', async () => {
    const onSubmit = vi.fn().mockResolvedValue(undefined)
    renderWithProviders(
      <MemoryRouter initialEntries={['/forgot-password']}>
        <Routes>
          <Route
            path="/forgot-password"
            element={<ForgotPasswordPage onSubmit={onSubmit} />}
          />
          <Route path="/reset-password" element={<ResetPasswordPage />} />
        </Routes>
      </MemoryRouter>,
    )

    fireEvent.change(screen.getByLabelText(/email address/i), {
      target: { value: 'user@example.com' },
    })

    fireEvent.click(screen.getByRole('button', { name: /send reset code/i }))

    await waitFor(() => expect(onSubmit).toHaveBeenCalledTimes(1))
    expect(onSubmit).toHaveBeenCalledWith({
      email: 'user@example.com',
    })

    expect(
      await screen.findByRole('heading', { name: /set a new password/i })
    ).toBeDefined()
    expect(screen.getByRole('alert').textContent).toMatch(
      /reset code has been sent/i
    )
  })

  it('lets users open the reset password page when they already have a code', async () => {
    renderWithProviders(
      <MemoryRouter initialEntries={['/forgot-password']}>
        <Routes>
          <Route path="/forgot-password" element={<ForgotPasswordPage />} />
          <Route path="/reset-password" element={<ResetPasswordPage />} />
        </Routes>
      </MemoryRouter>,
    )

    fireEvent.click(screen.getByRole('button', { name: /i already have a code/i }))

    expect(
      await screen.findByRole('heading', { name: /set a new password/i })
    ).toBeDefined()
  })
})

describe('ResetPasswordPage', () => {
  it('shows a field error when the email address is missing', async () => {
    renderPage(<ResetPasswordPage />)

    fireEvent.change(screen.getByLabelText(/reset code/i), {
      target: { value: '834697' },
    })
    fireEvent.change(screen.getByLabelText(/^new password$/i), {
      target: { value: 'Password123!' },
    })
    fireEvent.change(screen.getByLabelText(/^confirm new password$/i), {
      target: { value: 'Password123!' },
    })

    fireEvent.click(screen.getByRole('button', { name: /reset password/i }))

    await screen.findByText(/email is required/i)
  })

  it('shows a field error when the reset code is missing', async () => {
    renderPage(<ResetPasswordPage />)

    fireEvent.change(screen.getByLabelText(/email address/i), {
      target: { value: 'user@example.com' },
    })
    fireEvent.change(screen.getByLabelText(/^new password$/i), {
      target: { value: 'Password123!' },
    })
    fireEvent.change(screen.getByLabelText(/^confirm new password$/i), {
      target: { value: 'Password123!' },
    })

    fireEvent.click(screen.getByRole('button', { name: /reset password/i }))

    await screen.findByText(/reset code is required/i)
  })

  it('uses the code and email from state and submits a new password', async () => {
    const onSubmit = vi.fn().mockResolvedValue(undefined)
    renderWithProviders(
      <MemoryRouter
        initialEntries={[
          {
            pathname: '/reset-password',
            search: '?otp=test-token',
            state: { email: 'user@example.com' },
          },
        ]}
      >
        <ResetPasswordPage onSubmit={onSubmit} />
      </MemoryRouter>,
    )

    expect(
      (screen.getByLabelText(/email address/i) as HTMLInputElement).value
    ).toBe('user@example.com')
    expect(
      (screen.getByLabelText(/reset code/i) as HTMLInputElement).value
    ).toBe('test-token')

    fireEvent.change(screen.getByLabelText(/^new password$/i), {
      target: { value: 'Password123!' },
    })
    fireEvent.change(screen.getByLabelText(/^confirm new password$/i), {
      target: { value: 'Password123!' },
    })

    fireEvent.click(screen.getByRole('button', { name: /reset password/i }))

    await waitFor(() => expect(onSubmit).toHaveBeenCalledTimes(1))
    expect(onSubmit).toHaveBeenCalledWith({
      email: 'user@example.com',
      otp: 'test-token',
      password: 'Password123!',
      confirmPassword: 'Password123!',
    })

    await screen.findByRole('alert')
    expect(screen.getByRole('alert').textContent).toMatch(
      /your password has been reset/i
    )
  })
})
