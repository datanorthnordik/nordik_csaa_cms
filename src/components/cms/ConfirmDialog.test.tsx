import { describe, it, expect, vi } from 'vitest'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { ConfirmDialog } from './ConfirmDialog'

describe('ConfirmDialog', () => {
  const defaultProps = {
    open: true,
    title: 'Confirm Action',
    body: 'Are you sure?',
    confirmLabel: 'Confirm',
    onConfirm: vi.fn(),
    onClose: vi.fn(),
  }

  it('renders dialog when open', () => {
    render(<ConfirmDialog {...defaultProps} />)

    expect(screen.getByRole('dialog')).toBeInTheDocument()
  })

  it('displays title', () => {
    render(<ConfirmDialog {...defaultProps} />)

    expect(screen.getByText('Confirm Action')).toBeInTheDocument()
  })

  it('displays body content', () => {
    render(<ConfirmDialog {...defaultProps} />)

    expect(screen.getByText('Are you sure?')).toBeInTheDocument()
  })

  it('displays confirm button with custom label', () => {
    render(
      <ConfirmDialog {...defaultProps} confirmLabel="Delete" />
    )

    expect(screen.getByRole('button', { name: 'Delete' })).toBeInTheDocument()
  })

  it('calls onConfirm when confirm button is clicked', async () => {
    const handleConfirm = vi.fn()
    const user = await userEvent.setup()

    render(
      <ConfirmDialog {...defaultProps} onConfirm={handleConfirm} />
    )

    const confirmButton = screen.getByRole('button', { name: defaultProps.confirmLabel })
    await user.click(confirmButton)

    expect(handleConfirm).toHaveBeenCalled()
  })

  it('calls onClose when cancel button is clicked', async () => {
    const handleClose = vi.fn()
    const user = await userEvent.setup()

    render(
      <ConfirmDialog {...defaultProps} onClose={handleClose} />
    )

    const cancelButton = screen.getByRole('button', { name: /cancel|close/i })
    await user.click(cancelButton)

    expect(handleClose).toHaveBeenCalled()
  })

  it('displays destructive styling when destructive prop is true', () => {
    const { container } = render(
      <ConfirmDialog {...defaultProps} destructive={true} />
    )

    const confirmButton = screen.getByRole('button', { name: defaultProps.confirmLabel })
    expect(confirmButton).toHaveClass(expect.stringMatching(/destructive|danger/i))
  })

  it('disables buttons when busy', () => {
    render(<ConfirmDialog {...defaultProps} busy={true} />)

    const confirmButton = screen.getByRole('button', { name: defaultProps.confirmLabel })
    expect(confirmButton).toBeDisabled()
  })

  it('does not render when open is false', () => {
    const { container } = render(
      <ConfirmDialog {...defaultProps} open={false} />
    )

    const dialog = container.querySelector('dialog, [role="dialog"]')
    if (dialog) {
      expect(dialog).not.toBeVisible()
    }
  })
})
