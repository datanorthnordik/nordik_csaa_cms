import { describe, it, expect, vi } from 'vitest'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { CreateGalleryDialog } from './CreateGalleryDialog'

describe('CreateGalleryDialog', () => {
  const defaultProps = {
    open: true,
    onClose: vi.fn(),
    onSubmit: vi.fn(),
  }

  it('does not render when closed', () => {
    const { container } = render(
      <CreateGalleryDialog {...defaultProps} open={false} />
    )

    const dialog = container.querySelector('dialog, [role="dialog"]')
    if (dialog) {
      expect(dialog).not.toBeVisible()
    }
  })

  it('renders dialog form when open', () => {
    render(<CreateGalleryDialog {...defaultProps} />)

    expect(screen.getByRole('dialog')).toBeInTheDocument()
  })

  it('renders form fields', () => {
    render(<CreateGalleryDialog {...defaultProps} />)

    expect(screen.getByLabelText(/gallery name|name/i)).toBeInTheDocument()
  })

  it('calls onClose when close button is clicked', async () => {
    const handleClose = vi.fn()
    const user = await userEvent.setup()

    render(
      <CreateGalleryDialog {...defaultProps} onClose={handleClose} />
    )

    const closeButton = screen.getByRole('button', { name: /close|cancel/i })
    await user.click(closeButton)

    expect(handleClose).toHaveBeenCalled()
  })

  it('displays loading state when submitting', () => {
    render(
      <CreateGalleryDialog {...defaultProps} submitting={true} />
    )

    const submitButton = screen.getByRole('button', { name: /create|submit/i })
    expect(submitButton).toBeDisabled()
  })

  it('renders cancel button', () => {
    render(<CreateGalleryDialog {...defaultProps} />)

    expect(screen.getByRole('button', { name: /cancel/i })).toBeInTheDocument()
  })

  it('renders submit button', () => {
    render(<CreateGalleryDialog {...defaultProps} />)

    expect(screen.getByRole('button', { name: /create|submit/i })).toBeInTheDocument()
  })
})
