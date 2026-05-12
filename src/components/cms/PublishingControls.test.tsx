import { describe, it, expect, vi } from 'vitest'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { PublishingControls } from './PublishingControls'

describe('PublishingControls', () => {
  const defaultProps = {
    status: 'draft' as const,
    publishLabel: 'Publish',
    onSaveDraft: vi.fn(),
    onPublish: vi.fn(),
  }

  it('renders publishing controls', () => {
    render(<PublishingControls {...defaultProps} />)

    expect(screen.getByRole('button', { name: /publish/i })).toBeInTheDocument()
    expect(screen.getByRole('button', { name: /save draft/i })).toBeInTheDocument()
  })

  it('displays draft status badge', () => {
    render(<PublishingControls {...defaultProps} status="draft" />)

    expect(screen.getByText(/draft/i)).toBeInTheDocument()
  })

  it('displays published status badge', () => {
    render(<PublishingControls {...defaultProps} status="published" />)

    expect(screen.getByText(/published/i)).toBeInTheDocument()
  })

  it('calls onPublish when publish button is clicked', async () => {
    const handlePublish = vi.fn()
    const user = await userEvent.setup()

    render(
      <PublishingControls
        {...defaultProps}
        onPublish={handlePublish}
      />
    )

    const publishButton = screen.getByRole('button', { name: /publish/i })
    await user.click(publishButton)

    expect(handlePublish).toHaveBeenCalled()
  })

  it('calls onSaveDraft when save draft button is clicked', async () => {
    const handleSaveDraft = vi.fn()
    const user = await userEvent.setup()

    render(
      <PublishingControls
        {...defaultProps}
        onSaveDraft={handleSaveDraft}
      />
    )

    const saveDraftButton = screen.getByRole('button', { name: /save draft/i })
    await user.click(saveDraftButton)

    expect(handleSaveDraft).toHaveBeenCalled()
  })

  it('disables buttons when isSubmitting is true', () => {
    render(
      <PublishingControls
        {...defaultProps}
        isSubmitting={true}
      />
    )

    const publishButton = screen.getByRole('button', { name: /publish/i })
    const saveDraftButton = screen.getByRole('button', { name: /save draft/i })

    expect(publishButton).toBeDisabled()
    expect(saveDraftButton).toBeDisabled()
  })

  it('renders visibility control when provided', () => {
    render(
      <PublishingControls
        {...defaultProps}
        visibility="public"
        onVisibilityChange={vi.fn()}
      />
    )

    expect(screen.getByRole('combobox')).toBeInTheDocument()
  })

  it('calls onDelete when delete button is clicked', async () => {
    const handleDelete = vi.fn()
    const user = await userEvent.setup()

    render(
      <PublishingControls
        {...defaultProps}
        onDelete={handleDelete}
      />
    )

    const deleteButton = screen.getByRole('button', { name: /delete/i })
    await user.click(deleteButton)

    // Should open confirm dialog
    expect(screen.getByRole('dialog')).toBeInTheDocument()
  })
})
