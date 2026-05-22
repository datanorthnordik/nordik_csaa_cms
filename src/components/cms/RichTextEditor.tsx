import { useCallback, useEffect, useState, type FormEvent, type ReactNode } from 'react'
import { useTranslation } from 'react-i18next'
import { Mark, mergeAttributes } from '@tiptap/core'
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
  allowImages?: boolean
}

const FONT_SIZE_OPTIONS = [
  { value: '', label: 'Default' },
  { value: '8pt', label: '8' },
  { value: '9pt', label: '9' },
  { value: '10pt', label: '10' },
  { value: '11pt', label: '11' },
  { value: '12pt', label: '12' },
  { value: '14pt', label: '14' },
  { value: '16pt', label: '16' },
  { value: '18pt', label: '18' },
  { value: '20pt', label: '20' },
  { value: '22pt', label: '22' },
  { value: '24pt', label: '24' },
  { value: '26pt', label: '26' },
  { value: '28pt', label: '28' },
  { value: '32pt', label: '32' },
  { value: '36pt', label: '36' },
  { value: '40pt', label: '40' },
  { value: '44pt', label: '44' },
  { value: '48pt', label: '48' },
  { value: '54pt', label: '54' },
  { value: '60pt', label: '60' },
  { value: '66pt', label: '66' },
  { value: '72pt', label: '72' },
] as const

const FontSize = Mark.create({
  name: 'fontSize',

  addAttributes() {
    return {
      size: {
        default: null,
        parseHTML: (element) => {
          const fontSize = element.style.fontSize
          return fontSize || null
        },
        renderHTML: (attributes) => {
          if (!attributes.size) {
            return {}
          }

          return {
            style: `font-size: ${attributes.size}`,
          }
        },
      },
    }
  },

  parseHTML() {
    return [
      {
        style: 'font-size',
        getAttrs: (value) => {
          if (typeof value !== 'string' || !value.trim()) {
            return false
          }

          return {
            size: value,
          }
        },
      },
    ]
  },

  renderHTML({ HTMLAttributes }) {
    return ['span', mergeAttributes(HTMLAttributes), 0]
  },
})

export function RichTextEditor({
  value,
  onChange,
  placeholder,
  disabled = false,
  className,
  allowImages = true,
}: RichTextEditorProps) {
  const { t } = useTranslation()
  const [isFocused, setIsFocused] = useState(false)
  const [isLinkEditorOpen, setIsLinkEditorOpen] = useState(false)
  const [linkUrl, setLinkUrl] = useState('')
  const resolvedPlaceholder = placeholder ?? t('richText.placeholder')

  const editor = useEditor({
    extensions: [
      StarterKit.configure({ heading: false, link: false, underline: false }),
      Underline,
      FontSize,
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

  const handleToggleLinkEditor = useCallback(() => {
    if (!editor) {
      return
    }

    const previous = editor.getAttributes('link').href as string | undefined
    setLinkUrl(previous ?? '')
    setIsLinkEditorOpen((current) => !current)
  }, [editor])

  const handleApplyLink = useCallback(
    (event: FormEvent<HTMLFormElement>) => {
      event.preventDefault()
      if (!editor) {
        return
      }

      const trimmedUrl = linkUrl.trim()
      if (!trimmedUrl) {
        editor.chain().focus().extendMarkRange('link').unsetLink().run()
        setIsLinkEditorOpen(false)
        return
      }

      editor.chain().focus().extendMarkRange('link').setLink({ href: trimmedUrl }).run()
      setIsLinkEditorOpen(false)
    },
    [editor, linkUrl],
  )

  const handleRemoveLink = useCallback(() => {
    if (!editor) {
      return
    }

    editor.chain().focus().extendMarkRange('link').unsetLink().run()
    setLinkUrl('')
    setIsLinkEditorOpen(false)
  }, [editor])

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
  const currentFontSize = (editor.getAttributes('fontSize').size as string | undefined) ?? ''

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
            isActive={editor.isActive('link') || isLinkEditorOpen}
            onClick={handleToggleLinkEditor}
          >
            <LinkIcon />
          </ToolbarButton>
          {allowImages ? (
            <ToolbarButton
              editor={editor}
              label={t('richText.image')}
              isActive={false}
              onClick={handleSetImage}
            >
              <ImageIcon />
            </ToolbarButton>
          ) : null}
          <ToolbarButton
            editor={editor}
            label={t('richText.quote')}
            isActive={editor.isActive('blockquote')}
            onClick={() => editor.chain().focus().toggleBlockquote().run()}
          >
            <QuoteIcon />
          </ToolbarButton>
        </div>

        <div className={styles.toolbarGroup}>
          <FontSizeSelector
            editor={editor}
            currentSize={currentFontSize}
            t={t}
          />
        </div>
      </div>

      {isLinkEditorOpen ? (
        <form className={styles.linkEditor} onSubmit={handleApplyLink}>
          <label className={styles.linkField}>
            <span>{t('richText.linkUrl')}</span>
            <input
              type="url"
              value={linkUrl}
              disabled={!editor.isEditable}
              placeholder={t('richText.linkUrlPlaceholder')}
              onChange={(event) => setLinkUrl(event.target.value)}
            />
          </label>
          <div className={styles.linkActions}>
            <button
              type="submit"
              className={styles.linkActionButton}
              disabled={!editor.isEditable}
            >
              {t('richText.applyLink')}
            </button>
            <button
              type="button"
              className={styles.linkActionButtonSecondary}
              disabled={!editor.isEditable}
              onClick={handleRemoveLink}
            >
              {t('richText.removeLink')}
            </button>
          </div>
        </form>
      ) : null}

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

type FontSizeSelectorProps = {
  editor: Editor
  currentSize: string
  t: ReturnType<typeof useTranslation>['t']
}

function FontSizeSelector({ editor, currentSize, t }: FontSizeSelectorProps) {
  const [isOpen, setIsOpen] = useState(false)
  const displaySize = currentSize || t('richText.fontSizeDefault')
  
  const handleSelectSize = (value: string) => {
    if (!value) {
      editor.chain().focus().unsetMark('fontSize').run()
    } else {
      editor.chain().focus().setMark('fontSize', { size: value }).run()
    }
    setIsOpen(false)
  }

  return (
    <div className={styles.fontSizeContainer}>
      <button
        type="button"
        aria-label={t('richText.fontSize')}
        disabled={!editor.isEditable}
        onClick={() => setIsOpen(!isOpen)}
        className={styles.fontSizeButton}
        title={`${t('richText.fontSize')}: ${displaySize}`}
      >
        <span className={styles.fontSizeDisplay}>{displaySize}</span>
        <svg viewBox="0 0 12 8" width="10" height="8" className={styles.fontSizeDropdownIcon} aria-hidden="true">
          <path d="M1 1l5 5 5-5" stroke="currentColor" strokeWidth="1.5" fill="none" strokeLinecap="round" strokeLinejoin="round" />
        </svg>
      </button>
      
      {isOpen && (
        <div className={styles.fontSizeDropdown}>
          {FONT_SIZE_OPTIONS.map(({ value, label }) => (
            <button
              key={value || 'default'}
              type="button"
              className={[
                styles.fontSizeOption,
                currentSize === value ? styles.fontSizeOptionActive : '',
              ]
                .filter(Boolean)
                .join(' ')}
              onClick={() => handleSelectSize(value)}
              title={label}
            >
              <span className={styles.fontSizeOptionLabel}>{label}</span>
            </button>
          ))}
        </div>
      )}
    </div>
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
