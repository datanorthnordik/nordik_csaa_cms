import { fireEvent, render, screen, within } from '@testing-library/react'
import { I18nextProvider } from 'react-i18next'
import { describe, expect, it, vi } from 'vitest'
import i18n from '../../i18n'
import { ConfirmDialog } from './ConfirmDialog'

function renderDialog(overrides: Partial<Parameters<typeof ConfirmDialog>[0]> = {}) {
  const props = {
    open: true,
    title: 'Confirm action',
    onConfirm: vi.fn(),
    onClose: vi.fn(),
    ...overrides,
  }
  return render(
    <I18nextProvider i18n={i18n}>
      <ConfirmDialog {...props} />
    </I18nextProvider>,
  )
}

describe('ConfirmDialog', () => {
  it('renders the dialog title when open is true', () => {
    renderDialog()
    expect(screen.getByRole('dialog')).toBeDefined()
    expect(screen.getByRole('heading', { name: /confirm action/i })).toBeDefined()
  })

  it('does not render the dialog when open is false', () => {
    renderDialog({ open: false })
    expect(screen.queryByRole('dialog')).toBeNull()
  })

  it('renders body text when provided', () => {
    renderDialog({ body: 'This cannot be undone.' })
    expect(screen.getByText('This cannot be undone.')).toBeDefined()
  })

  it('does not render body section when body is omitted', () => {
    renderDialog()
    expect(screen.queryByText(/cannot be undone/i)).toBeNull()
  })

  it('calls onConfirm when the confirm button is clicked', () => {
    const onConfirm = vi.fn()
    renderDialog({ onConfirm })
    fireEvent.click(within(screen.getByRole('dialog')).getByRole('button', { name: /confirm/i }))
    expect(onConfirm).toHaveBeenCalledTimes(1)
  })

  it('calls onClose when the cancel button is clicked', () => {
    const onClose = vi.fn()
    renderDialog({ onClose })
    fireEvent.click(within(screen.getByRole('dialog')).getByRole('button', { name: /cancel/i }))
    expect(onClose).toHaveBeenCalledTimes(1)
  })

  it('uses custom confirmLabel and cancelLabel', () => {
    renderDialog({ confirmLabel: 'Yes, proceed', cancelLabel: 'No thanks' })
    const dialog = screen.getByRole('dialog')
    expect(within(dialog).getByRole('button', { name: /yes, proceed/i })).toBeDefined()
    expect(within(dialog).getByRole('button', { name: /no thanks/i })).toBeDefined()
  })

  it('shows "Working…" and disables buttons when busy is true', () => {
    renderDialog({ busy: true })
    const dialog = screen.getByRole('dialog')
    const buttons = within(dialog).getAllByRole('button')
    buttons.forEach((btn) => {
      expect((btn as HTMLButtonElement).disabled).toBe(true)
    })
    expect(within(dialog).getByText(/working/i)).toBeDefined()
  })

  it('does not call onClose when busy and cancel is clicked', () => {
    const onClose = vi.fn()
    renderDialog({ busy: true, onClose })
    fireEvent.click(within(screen.getByRole('dialog')).getByRole('button', { name: /cancel/i }))
    expect(onClose).not.toHaveBeenCalled()
  })
})
