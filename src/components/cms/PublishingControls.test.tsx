import { fireEvent, render, screen } from '@testing-library/react'
import { I18nextProvider } from 'react-i18next'
import { describe, expect, it, vi } from 'vitest'
import i18n from '../../i18n'
import { PublishingControls } from './PublishingControls'
import type { PublishingControlsProps } from './PublishingControls'

function renderControls(overrides: Partial<PublishingControlsProps> = {}) {
  const props: PublishingControlsProps = {
    status: 'draft',
    ...overrides,
  }
  return render(
    <I18nextProvider i18n={i18n}>
      <PublishingControls {...props} />
    </I18nextProvider>,
  )
}

describe('PublishingControls', () => {
  it('renders the status badge', () => {
    renderControls({ status: 'published' })
    expect(screen.getByText(/published/i)).toBeDefined()
  })

  it('renders "Draft" status badge for draft entries', () => {
    renderControls({ status: 'draft' })
    // Use exact match to distinguish the "Draft" badge from the "Save Draft" button
    expect(screen.getByText('Draft')).toBeDefined()
  })

  it('does not render a visibility row when visibility prop is omitted', () => {
    renderControls()
    expect(screen.queryByLabelText(/visibility/i)).toBeNull()
    expect(screen.queryByRole('combobox')).toBeNull()
  })

  it('renders a visibility select when visibility and onVisibilityChange are provided', () => {
    renderControls({
      visibility: 'public',
      onVisibilityChange: vi.fn(),
    })
    const select = screen.getByRole('combobox', { name: /visibility/i })
    expect(select).toBeDefined()
    expect((select as HTMLSelectElement).value).toBe('public')
  })

  it('calls onVisibilityChange when the select value changes', () => {
    const onVisibilityChange = vi.fn()
    renderControls({
      visibility: 'public',
      onVisibilityChange,
    })
    fireEvent.change(screen.getByRole('combobox', { name: /visibility/i }), {
      target: { value: 'private' },
    })
    expect(onVisibilityChange).toHaveBeenCalledWith('private')
  })

  it('displays a static visibility text when onVisibilityChange is not provided', () => {
    renderControls({ visibility: 'private' })
    expect(screen.queryByRole('combobox')).toBeNull()
    expect(screen.getByText(/private/i)).toBeDefined()
  })

  it('uses a custom heading when heading prop is provided', () => {
    renderControls({ heading: 'Custom Heading' })
    expect(screen.getByRole('heading', { name: 'Custom Heading' })).toBeDefined()
  })

  it('renders extra slot content when extraSlot is provided', () => {
    renderControls({ extraSlot: <div>Extra content</div> })
    expect(screen.getByText('Extra content')).toBeDefined()
  })

  it('does not render action buttons when no action handlers are provided', () => {
    renderControls()
    expect(screen.queryByRole('button', { name: /publish/i })).toBeNull()
    expect(screen.queryByRole('button', { name: /save draft/i })).toBeNull()
    expect(screen.queryByRole('button', { name: /delete/i })).toBeNull()
  })

  it('renders Publish and Save Draft buttons when onPublish and onSaveDraft are provided', () => {
    renderControls({
      publishLabel: 'Publish Now',
      onPublish: vi.fn(),
      onSaveDraft: vi.fn(),
    })
    expect(screen.getByRole('button', { name: /publish now/i })).toBeDefined()
    expect(screen.getByRole('button', { name: /save draft/i })).toBeDefined()
  })

  it('does not render Delete button when onDelete is not provided', () => {
    renderControls({
      publishLabel: 'Publish',
      onPublish: vi.fn(),
      onSaveDraft: vi.fn(),
    })
    expect(screen.queryByRole('button', { name: /delete/i })).toBeNull()
  })

  it('renders Delete button when onDelete is provided', () => {
    renderControls({
      publishLabel: 'Publish',
      onPublish: vi.fn(),
      onSaveDraft: vi.fn(),
      onDelete: vi.fn(),
    })
    // default delete label is "Move to Trash" from publishing.delete translation
    expect(screen.getByRole('button', { name: /move to trash/i })).toBeDefined()
  })
})
