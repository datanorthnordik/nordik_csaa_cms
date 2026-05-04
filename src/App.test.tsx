import { describe, it, expect, vi } from 'vitest'
import { fireEvent, render, screen } from '@testing-library/react'
import App from './App'
import { LoginPage } from './pages/LoginPage'
import { SignupPage } from './pages/SignupPage'

describe('App Component', () => {
  it('renders the login page by default', () => {
    render(<App />)
    expect(screen.getByRole('heading', { name: /sign in with email address/i })).toBeDefined()
  })

  it('switches to the signup page when the create account button is clicked', () => {
    render(<App />)
    fireEvent.click(screen.getByRole('button', { name: /create new account/i }))
    expect(screen.getByRole('heading', { name: /create your account/i })).toBeDefined()
  })
})

describe('LoginPage', () => {
  it('shows an error when the form is submitted without credentials', () => {
    render(<LoginPage />)
    fireEvent.click(screen.getByRole('button', { name: /sign in/i }))
    expect(screen.getByRole('alert').textContent).toMatch(/please enter your email and password/i)
  })

  it('shows an error when the email address is invalid', () => {
    render(<LoginPage />)
    fireEvent.change(screen.getByPlaceholderText(/email address/i), {
      target: { value: 'invalid-email' },
    })
    fireEvent.change(screen.getByPlaceholderText(/password/i), {
      target: { value: 'Password123!' },
    })
    fireEvent.click(screen.getByRole('button', { name: /sign in/i }))
    expect(screen.getByRole('alert').textContent).toMatch(/please enter a valid email address/i)
  })
})

describe('SignupPage', () => {
  it('shows an error when required fields are missing', () => {
    render(<SignupPage />)
    fireEvent.click(screen.getByRole('button', { name: /create account/i }))
    expect(screen.getByRole('alert').textContent).toMatch(/please fill in all fields/i)
  })

  it('shows an error when the email address is invalid', () => {
    render(<SignupPage />)
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
    expect(screen.getByRole('alert').textContent).toMatch(/please enter a valid work email address/i)
  })

  it('shows an error when passwords do not match', () => {
    render(<SignupPage />)
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
    expect(screen.getByRole('alert').textContent).toMatch(/passwords do not match/i)
  })

  it('submits valid signup data and displays a success message', async () => {
    const onSubmit = vi.fn().mockResolvedValue(undefined)
    render(<SignupPage onSubmit={onSubmit} />)

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
    expect(onSubmit).toHaveBeenCalledTimes(1)
    expect(onSubmit).toHaveBeenCalledWith({
      firstName: 'Jane',
      lastName: 'Doe',
      email: 'jane.doe@example.com',
      password: 'Password123!',
      confirmPassword: 'Password123!',
    })
    expect((await screen.findByRole('alert')).textContent).toMatch(
      /account request submitted/i,
    )
  })

  it('calls onSignIn when the sign in button is clicked', () => {
    const onSignIn = vi.fn()
    render(<SignupPage onSignIn={onSignIn} />)
    fireEvent.click(screen.getByRole('button', { name: /sign in/i }))
    expect(onSignIn).toHaveBeenCalledTimes(1)
  })
})
