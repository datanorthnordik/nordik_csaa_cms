import { describe, it, expect } from 'vitest'
import { render, screen } from '@testing-library/react'
import { AuthLayout } from './AuthLayout'

describe('AuthLayout', () => {
  it('renders auth layout', () => {
    const { container } = render(
      <AuthLayout>
        <div>Login Content</div>
      </AuthLayout>
    )

    expect(container).toBeInTheDocument()
  })

  it('renders children', () => {
    render(
      <AuthLayout>
        <div>Login Content</div>
      </AuthLayout>
    )

    expect(screen.getByText('Login Content')).toBeInTheDocument()
  })

  it('applies proper layout classes', () => {
    const { container } = render(
      <AuthLayout>
        <div>Login Content</div>
      </AuthLayout>
    )

    const layout = container.querySelector('[class*="layout"], [class*="auth"]')
    expect(layout).toBeInTheDocument()
  })
})
