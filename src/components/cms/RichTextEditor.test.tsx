import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { Editor } from '@tiptap/core'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../../i18n'
import { RichTextEditor, createRichTextEditorExtensions } from './RichTextEditor'

describe('RichTextEditor', () => {
  beforeEach(async () => {
    await i18n.changeLanguage('en')

    if (!('ResizeObserver' in window)) {
      Object.defineProperty(window, 'ResizeObserver', {
        configurable: true,
        writable: true,
        value: class {
          observe() {}
          unobserve() {}
          disconnect() {}
        },
      })
    }

    if (!('requestAnimationFrame' in window)) {
      Object.defineProperty(window, 'requestAnimationFrame', {
        configurable: true,
        writable: true,
        value: (callback: FrameRequestCallback) => setTimeout(callback, 0),
      })
    }

    if (!('cancelAnimationFrame' in window)) {
      Object.defineProperty(window, 'cancelAnimationFrame', {
        configurable: true,
        writable: true,
        value: (id: number) => clearTimeout(id),
      })
    }
  })

  it('does not render the image insertion action in the toolbar', async () => {
    render(<RichTextEditor value="<p>Hello</p>" onChange={vi.fn()} />)

    await waitFor(() => {
      expect(screen.getByRole('toolbar', { name: 'Formatting toolbar' })).toBeDefined()
    })

    expect(screen.queryByRole('button', { name: 'Insert image' })).toBeNull()
  })

  it('keeps italic support in the shared editor extension stack', () => {
    const editor = new Editor({
      extensions: createRichTextEditorExtensions(),
      content: '<p>Hello world</p>',
      element: document.createElement('div'),
    })

    editor.commands.setTextSelection({ from: 1, to: 6 })

    expect(editor.chain().focus().toggleItalic().run()).toBe(true)
    expect(editor.getHTML()).toContain('<em>Hello</em>')

    editor.destroy()
  })

  it('opens the link editor when the toolbar link button is clicked', async () => {
    render(<RichTextEditor value="<p>Hello</p>" onChange={vi.fn()} />)

    await waitFor(() => {
      expect(screen.getByRole('button', { name: 'Insert link' })).toBeDefined()
    })

    fireEvent.click(screen.getByRole('button', { name: 'Insert link' }))

    expect(screen.getByRole('textbox', { name: 'Link destination' })).toBeDefined()
  })
})
