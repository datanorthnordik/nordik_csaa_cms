import { describe, it, expect } from 'vitest'
import { render, screen } from '@testing-library/react'
import { Loader } from './Loader'

describe('Loader', () => {
  it('renders loader element', () => {
    const { container } = render(<Loader />)
    expect(container.querySelector('[role="status"]')).toBeInTheDocument()
  })

  it('displays label when provided', () => {
    render(<Loader label="Loading data..." />)
    expect(screen.getByText('Loading data...')).toBeInTheDocument()
  })

  it('does not display label when not provided', () => {
    const { container } = render(<Loader />)
    expect(container.querySelector('[role="status"]')).toBeInTheDocument()
  })

  it('applies custom className when provided', () => {
    const { container } = render(<Loader className="custom-loader" />)
    expect(container.querySelector('.custom-loader')).toBeInTheDocument()
  })

  it('has accessible aria-label or aria-live', () => {
    const { container } = render(<Loader label="Loading" />)
    const loader = container.querySelector('[role="status"]')
    expect(loader).toHaveAttribute('aria-live', 'polite')
  })
})
