import { useCallback, useEffect, useState, type ReactNode } from 'react'
import { useTranslation } from 'react-i18next'
import { EditorContent, useEditor, type Editor } from '@tiptap/react'
import StarterKit from '@tiptap/starter-kit'
import Underline from '@tiptap/extension-underline'
import Link from '@tiptap/extension-link'
import Image from '@tiptap/extension-image'
import styles from './RichTextEditor.module.css'

type RichTextEditorProps = {
  value: string
  onChange: (html: string) => void
  placeholder?: string
  disabled?: boolean
  className?: string
}

export function RichTextEditor({
  value,
  onChange,
  placeholder,
  disabled = false,
  className,
}: RichTextEditorProps) {
  const { t } = useTranslation()
  const [isFocused, setIsFocused] = useState(false)
  const resolvedPlaceholder = placeholder ?? t('richText.placeholder')

  const editor = useEditor({
    extensions: [
      StarterKit.configure({ heading: false }),
      Underline,
      Link.configure({
        openOnClick: false,
        HTMLAttributes: { rel: 'noopener noreferrer', target: '_blank' },
      }),
      Image,
    ],
    content: value,
    editable: !disabled,
    onUpdate: ({ editor: currentEditor }) => {
      onChange(currentEditor.getHTML())
    },
    onFocus: () => setIsFocused(true),
    onBlur: () => setIsFocused(false),
    editorProps: {
      attributes: {
        class: styles.editor,
        'data-placeholder': resolvedPlaceholder,
      },
    },
  })

  useEffect(() => {
    if (!editor) {
      return
    }
    if (editor.getHTML() === value) {
      return
    }
    editor.commands.setContent(value || '', { emitUpdate: false })
  }, [editor, value])

  useEffect(() => {
    editor?.setEditable(!disabled)
  }, [editor, disabled])

  const handleSetLink = useCallback(() => {
    if (!editor) {
      return
    }
    const previous = editor.getAttributes('link').href as string | undefined
    const url = window.prompt(t('richText.linkPrompt'), previous ?? 'https://')
    if (url === null) {
      return
    }
    if (url === '') {
      editor.chain().focus().extendMarkRange('link').unsetLink().run()
      return
    }
    editor.chain().focus().extendMarkRange('link').setLink({ href: url }).run()
  }, [editor, t])

  const handleSetImage = useCallback(() => {
    if (!editor) {
      return
    }
    const url = window.prompt(t('richText.imagePrompt'), 'https://')
    if (!url) {
      return
    }
    editor.chain().focus().setImage({ src: url }).run()
  }, [editor, t])

  if (!editor) {
    return null
  }

  const classes = [
    styles.wrapper,
    isFocused ? styles.focused : '',
    disabled ? styles.disabled : '',
    className,
  ]
    .filter(Boolean)
    .join(' ')

  return (
    <div className={classes}>
      <div className={styles.toolbar} role="toolbar" aria-label={t('richText.toolbar')}>
        <div className={styles.toolbarGroup}>
          <ToolbarButton
            editor={editor}
            label={t('richText.bold')}
            isActive={editor.isActive('bold')}
            onClick={() => editor.chain().focus().toggleBold().run()}
          >
            <strong>B</strong>
          </ToolbarButton>
          <ToolbarButton
            editor={editor}
            label={t('richText.italic')}
            isActive={editor.isActive('italic')}
            onClick={() => editor.chain().focus().toggleItalic().run()}
          >
            <span className={styles.italic}>I</span>
          </ToolbarButton>
          <ToolbarButton
            editor={editor}
            label={t('richText.underline')}
            isActive={editor.isActive('underline')}
            onClick={() => editor.chain().focus().toggleUnderline().run()}
          >
            <span className={styles.underline}>U</span>
          </ToolbarButton>
        </div>

        <div className={styles.toolbarGroup}>
          <ToolbarButton
            editor={editor}
            label={t('richText.bulletList')}
            isActive={editor.isActive('bulletList')}
            onClick={() => editor.chain().focus().toggleBulletList().run()}
          >
            <BulletListIcon />
          </ToolbarButton>
          <ToolbarButton
            editor={editor}
            label={t('richText.orderedList')}
            isActive={editor.isActive('orderedList')}
            onClick={() => editor.chain().focus().toggleOrderedList().run()}
          >
            <OrderedListIcon />
          </ToolbarButton>
        </div>

        <div className={styles.toolbarGroup}>
          <ToolbarButton
            editor={editor}
            label={t('richText.link')}
            isActive={editor.isActive('link')}
            onClick={handleSetLink}
          >
            <LinkIcon />
          </ToolbarButton>
          <ToolbarButton
            editor={editor}
            label={t('richText.image')}
            isActive={false}
            onClick={handleSetImage}
          >
            <ImageIcon />
          </ToolbarButton>
          <ToolbarButton
            editor={editor}
            label={t('richText.quote')}
            isActive={editor.isActive('blockquote')}
            onClick={() => editor.chain().focus().toggleBlockquote().run()}
          >
            <QuoteIcon />
          </ToolbarButton>
        </div>
      </div>

      <EditorContent editor={editor} />
    </div>
  )
}

type ToolbarButtonProps = {
  editor: Editor
  label: string
  isActive: boolean
  onClick: () => void
  children: ReactNode
}

function ToolbarButton({ editor, label, isActive, onClick, children }: ToolbarButtonProps) {
  return (
    <button
      type="button"
      aria-label={label}
      aria-pressed={isActive}
      title={label}
      disabled={!editor.isEditable}
      onClick={onClick}
      className={[styles.toolbarButton, isActive ? styles.toolbarButtonActive : '']
        .filter(Boolean)
        .join(' ')}
    >
      {children}
    </button>
  )
}

function BulletListIcon() {
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden="true">
      <circle cx="2.5" cy="4" r="1" fill="currentColor" />
      <circle cx="2.5" cy="8" r="1" fill="currentColor" />
      <circle cx="2.5" cy="12" r="1" fill="currentColor" />
      <path d="M5.5 4h8M5.5 8h8M5.5 12h8" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
    </svg>
  )
}

function OrderedListIcon() {
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden="true">
      <text x="1" y="6" fontSize="5" fill="currentColor" fontWeight="700">1</text>
      <text x="1" y="11" fontSize="5" fill="currentColor" fontWeight="700">2</text>
      <text x="1" y="16" fontSize="5" fill="currentColor" fontWeight="700">3</text>
      <path d="M6 4h8M6 9h8M6 14h8" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
    </svg>
  )
}

function LinkIcon() {
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden="true">
      <path
        d="M6.5 9.5a3 3 0 0 0 4.24 0l1.76-1.76a3 3 0 0 0-4.24-4.24l-1 1M9.5 6.5a3 3 0 0 0-4.24 0L3.5 8.26a3 3 0 0 0 4.24 4.24l1-1"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

function ImageIcon() {
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden="true">
      <rect x="2" y="3" width="12" height="10" rx="1.5" stroke="currentColor" strokeWidth="1.4" />
      <circle cx="6" cy="7" r="1.2" fill="currentColor" />
      <path d="m3 12 3-3 2 2 2.5-2.5L14 12" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  )
}

function QuoteIcon() {
  return (
    <svg viewBox="0 0 16 16" width="14" height="14" fill="none" aria-hidden="true">
      <path
        d="M4 5c-1.2 0-2 .9-2 2 0 1.2.8 2 2 2 .4 0 .6-.1.6-.1-.1 1-.7 1.7-1.6 2l.6 1c1.8-.5 2.9-2 2.9-4 0-1.7-.9-2.9-2.5-2.9Zm6.5 0c-1.2 0-2 .9-2 2 0 1.2.8 2 2 2 .4 0 .6-.1.6-.1-.1 1-.7 1.7-1.6 2l.6 1c1.8-.5 2.9-2 2.9-4 0-1.7-.9-2.9-2.5-2.9Z"
        fill="currentColor"
      />
    </svg>
  )
}
