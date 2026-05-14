import { fireEvent, render, screen, waitFor, within } from '@testing-library/react'
import { I18nextProvider } from 'react-i18next'
import { describe, expect, it, vi } from 'vitest'
import i18n from '../../i18n'
import { EntryActions } from './EntryActions'
import type { EntryActionsProps } from './EntryActions'

function renderActions(overrides: Partial<EntryActionsProps> = {}) {
  const props: EntryActionsProps = {
    status: 'draft',
    onPublish: vi.fn(),
    onSaveDraft: vi.fn(),
    ...overrides,
  }
  return render(
    <I18nextProvider i18n={i18n}>
      <EntryActions {...props} />
    </I18nextProvider>,
  )
}

describe('EntryActions', () => {
  it('renders publish and save draft buttons', () => {
    renderActions()
    expect(screen.getByRole('button', { name: /publish/i })).toBeDefined()
    expect(screen.getByRole('button', { name: /save draft/i })).toBeDefined()
  })

  it('shows "Save changes" as the primary label when status is published', () => {
    renderActions({ status: 'published' })
    expect(screen.getByRole('button', { name: /save changes/i })).toBeDefined()
  })

  it('does not render a delete button when onDelete is not provided', () => {
    renderActions()
    expect(screen.queryByRole('button', { name: /trash/i })).toBeNull()
  })

  it('renders a delete button when onDelete is provided', () => {
    renderActions({ onDelete: vi.fn() })
    // default label with no entryType is "Move to Trash"
    expect(screen.getByRole('button', { name: /move to trash/i })).toBeDefined()
  })

  it('shows "Delete event" when entryType is event', () => {
    renderActions({ entryType: 'event', onDelete: vi.fn() })
    expect(screen.getByRole('button', { name: /delete event/i })).toBeDefined()
  })

  it('shows "Delete Press Entry" when entryType is press', () => {
    renderActions({ entryType: 'press', onDelete: vi.fn() })
    expect(screen.getByRole('button', { name: /delete press entry/i })).toBeDefined()
  })

  it('explicit deleteLabel overrides entryType default', () => {
    renderActions({ entryType: 'event', deleteLabel: 'Remove item', onDelete: vi.fn() })
    expect(screen.getByRole('button', { name: /remove item/i })).toBeDefined()
    expect(screen.queryByRole('button', { name: /delete event/i })).toBeNull()
  })

  it('falls back to "Move to Trash" when no entryType or deleteLabel is given', () => {
    renderActions({ onDelete: vi.fn() })
    expect(screen.getByRole('button', { name: /move to trash/i })).toBeDefined()
  })

  it('calls onPublish when the primary button is clicked', () => {
    const onPublish = vi.fn()
    renderActions({ onPublish })
    fireEvent.click(screen.getByRole('button', { name: /publish/i }))
    expect(onPublish).toHaveBeenCalledTimes(1)
  })

  it('calls onSaveDraft when the save draft button is clicked', () => {
    const onSaveDraft = vi.fn()
    renderActions({ onSaveDraft })
    fireEvent.click(screen.getByRole('button', { name: /save draft/i }))
    expect(onSaveDraft).toHaveBeenCalledTimes(1)
  })

  it('disables publish and save draft when isSubmitting is true', () => {
    renderActions({ isSubmitting: true })
    const primary = screen.getByRole('button', { name: /working/i })
    expect((primary as HTMLButtonElement).disabled).toBe(true)
    const saveDraft = screen.getByRole('button', { name: /save draft/i })
    expect((saveDraft as HTMLButtonElement).disabled).toBe(true)
  })

  it('disables the delete button when isDeleting is true', () => {
    renderActions({ onDelete: vi.fn(), entryType: 'event', isDeleting: true })
    const deleteBtn = screen.getByRole('button', { name: /delete event/i })
    expect((deleteBtn as HTMLButtonElement).disabled).toBe(true)
  })

  it('opens a confirm dialog when the delete button is clicked', async () => {
    renderActions({
      onDelete: vi.fn(),
      entryType: 'event',
      deleteConfirmTitle: 'Really delete?',
    })
    fireEvent.click(screen.getByRole('button', { name: /delete event/i }))
    await waitFor(() => expect(screen.getByRole('dialog')).toBeDefined())
    expect(within(screen.getByRole('dialog')).getByText('Really delete?')).toBeDefined()
  })

  it('calls onDelete when the confirm dialog is confirmed', async () => {
    const onDelete = vi.fn()
    renderActions({
      onDelete,
      entryType: 'event',
      deleteConfirmTitle: 'Confirm delete',
      deleteConfirmLabel: 'Yes, delete',
    })
    fireEvent.click(screen.getByRole('button', { name: /delete event/i }))
    await waitFor(() => screen.getByRole('dialog'))
    fireEvent.click(within(screen.getByRole('dialog')).getByRole('button', { name: /yes, delete/i }))
    expect(onDelete).toHaveBeenCalledTimes(1)
  })

  it('closes the confirm dialog when cancel is clicked without calling onDelete', async () => {
    const onDelete = vi.fn()
    renderActions({ onDelete, entryType: 'event' })
    fireEvent.click(screen.getByRole('button', { name: /delete event/i }))
    await waitFor(() => screen.getByRole('dialog'))
    fireEvent.click(within(screen.getByRole('dialog')).getByRole('button', { name: /cancel/i }))
    expect(onDelete).not.toHaveBeenCalled()
  })
})
