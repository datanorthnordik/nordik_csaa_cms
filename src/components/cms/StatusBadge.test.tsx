import { describe, it, expect, vi } from 'vitest'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { StatusBadge } from './StatusBadge'

describe('StatusBadge', () => {
  it('renders draft status badge', () => {
    render(<StatusBadge status="draft" />)

    expect(screen.getByText(/draft/i)).toBeInTheDocument()
  })

  it('renders published status badge', () => {
    render(<StatusBadge status="published" />)

    expect(screen.getByText(/published/i)).toBeInTheDocument()
  })

  it('applies correct styling for draft status', () => {
    const { container } = render(<StatusBadge status="draft" />)

    const badge = container.querySelector('[class*="badge"]')
    expect(badge).toBeInTheDocument()
    expect(badge?.textContent).toContain('draft')
  })

  it('applies correct styling for published status', () => {
    const { container } = render(<StatusBadge status="published" />)

    const badge = container.querySelector('[class*="badge"]')
    expect(badge).toBeInTheDocument()
    expect(badge?.textContent).toContain('published')
  })

  it('renders with custom className', () => {
    const { container } = render(
      <StatusBadge status="draft" className="custom-class" />
    )

    const badge = container.querySelector('.custom-class')
    expect(badge).toBeInTheDocument()
  })
})
