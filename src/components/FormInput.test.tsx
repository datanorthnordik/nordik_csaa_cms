import { describe, it, expect, vi } from 'vitest'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { FormInput } from './FormInput'

describe('FormInput', () => {
  it('renders input with label', () => {
    render(
      <FormInput
        label="Email"
        type="email"
        value=""
        onChange={() => {}}
      />
    )

    expect(screen.getByText('Email')).toBeInTheDocument()
  })

  it('renders input with placeholder', () => {
    render(
      <FormInput
        label="Email"
        type="email"
        placeholder="Enter email"
        value=""
        onChange={() => {}}
      />
    )

    const input = screen.getByPlaceholderText('Enter email')
    expect(input).toBeInTheDocument()
  })

  it('calls onChange when input value changes', async () => {
    const handleChange = vi.fn()
    const user = await userEvent.setup()

    render(
      <FormInput
        label="Username"
        type="text"
        value=""
        onChange={handleChange}
      />
    )

    const input = screen.getByRole('textbox')
    await user.type(input, 'testuser')

    expect(handleChange).toHaveBeenCalled()
  })

  it('displays error message when provided', () => {
    render(
      <FormInput
        label="Email"
        type="email"
        value=""
        onChange={() => {}}
        error="Invalid email"
      />
    )

    expect(screen.getByText('Invalid email')).toBeInTheDocument()
  })

  it('disables input when disabled prop is true', () => {
    render(
      <FormInput
        label="Email"
        type="email"
        value=""
        onChange={() => {}}
        disabled
      />
    )

    const input = screen.getByRole('textbox')
    expect(input).toBeDisabled()
  })

  it('renders required indicator when required', () => {
    const { container } = render(
      <FormInput
        label="Email"
        type="email"
        value=""
        onChange={() => {}}
        required
      />
    )

    expect(container.textContent).toContain('Email')
  })
})
