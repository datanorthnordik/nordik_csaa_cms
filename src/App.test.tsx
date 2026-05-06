import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { MemoryRouter } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import App from './App'
import i18n from './i18n'
import { LoginPage } from './pages/LoginPage'
import { SignupPage } from './pages/SignupPage'

beforeEach(async () => {
  window.localStorage.clear()
  window.history.pushState({}, '', '/')
  await i18n.changeLanguage('en')
})

describe('App Component', () => {
  it('renders the login page by default', () => {
    render(<App />)
    expect(
      screen.getByRole('heading', { name: /sign in with email address/i })
    ).toBeDefined()
    expect(screen.getByRole('button', { name: 'FR' })).toBeDefined()
  })

  it('renders the signup page when navigating to /signup', () => {
    window.history.pushState({}, '', '/signup')
    render(<App />)
    expect(
      screen.getByRole('heading', { name: /create your account/i })
    ).toBeDefined()
  })
})

describe('LoginPage', () => {
  it('shows field errors when the form is submitted without credentials', async () => {
    render(
      <MemoryRouter>
        <LoginPage />
      </MemoryRouter>
    )

    fireEvent.click(screen.getByRole('button', { name: /sign in/i }))

    await screen.findByText(/email is required/i)
    expect(screen.getByText(/password is required/i)).toBeDefined()
  })

  it('shows field error when the email address is invalid', async () => {
    render(
      <MemoryRouter>
        <LoginPage />
      </MemoryRouter>
    )

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
    render(
      <MemoryRouter>
        <LoginPage onSubmit={onSubmit} />
      </MemoryRouter>
    )

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
    render(
      <MemoryRouter>
        <LoginPage />
      </MemoryRouter>
    )

    fireEvent.click(screen.getByRole('button', { name: 'FR' }))

    expect(
      await screen.findByRole('heading', {
        name: /se connecter avec une adresse courriel/i,
      })
    ).toBeDefined()
    expect(
      screen.getByRole('button', { name: /créer un nouveau compte/i })
    ).toBeDefined()
  })
})

describe('SignupPage', () => {
  it('shows field errors when required fields are missing', async () => {
    render(
      <MemoryRouter>
        <SignupPage />
      </MemoryRouter>
    )

    fireEvent.click(screen.getByRole('button', { name: /create account/i }))

    await screen.findByText(/first name is required/i)
    expect(screen.getByText(/email is required/i)).toBeDefined()
    expect(screen.getByText(/password is required/i)).toBeDefined()
  })

  it('shows field error when the email address is invalid', async () => {
    render(
      <MemoryRouter>
        <SignupPage />
      </MemoryRouter>
    )

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
    render(
      <MemoryRouter>
        <SignupPage />
      </MemoryRouter>
    )

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
    render(
      <MemoryRouter>
        <SignupPage />
      </MemoryRouter>
    )

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

    await screen.findByText(/password must be at least 8 characters/i)
    expect(screen.getByText(/password must be at least 8 characters/i)).toBeDefined()
  })

  it('submits valid signup data and displays a success message', async () => {
    const onSubmit = vi.fn().mockResolvedValue(undefined)
    render(
      <MemoryRouter>
        <SignupPage onSubmit={onSubmit} />
      </MemoryRouter>
    )

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
    render(
      <MemoryRouter initialEntries={['/signup']}>
        <SignupPage />
      </MemoryRouter>
    )

    expect(screen.getByRole('button', { name: /sign in/i })).toBeDefined()
  })

  it('shows the language toggle on the signup page', () => {
    render(
      <MemoryRouter initialEntries={['/signup']}>
        <SignupPage />
      </MemoryRouter>
    )

    expect(screen.getByRole('button', { name: 'EN' })).toBeDefined()
    expect(screen.getByRole('button', { name: 'FR' })).toBeDefined()
  })
})
