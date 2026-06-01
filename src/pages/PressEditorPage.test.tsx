import type { ReactNode } from 'react'
import { fireEvent, render, screen } from '@testing-library/react'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import { RESOURCE_UPLOAD_MAX_FILE_SIZE_BYTES } from '../lib/resourceUpload'
import { PressEditorPage } from './PressEditorPage'

const {
  mockNavigate,
  mockCreate,
  mockUpdate,
  mockRemove,
  toastSuccess,
  toastError,
} = vi.hoisted(() => ({
  mockNavigate: vi.fn(),
  mockCreate: vi.fn(),
  mockUpdate: vi.fn(),
  mockRemove: vi.fn(),
  toastSuccess: vi.fn(),
  toastError: vi.fn(),
}))

vi.mock('../hooks/usePressEntries', () => ({
  usePressEntries: () => ({
    create: mockCreate,
    update: mockUpdate,
    remove: mockRemove,
  }),
}))

vi.mock('../components/CmsAppShell', () => ({
  CmsAppShell: ({ children }: { children?: ReactNode }) => (
    <div data-testid="cms-shell">{children}</div>
  ),
}))

vi.mock('../components/Loader', () => ({
  Loader: ({ label }: { label?: string }) => <div>{label ?? 'Loading'}</div>,
}))

vi.mock('../components/cms/RichTextEditor', () => ({
  RichTextEditor: ({
    value,
    onChange,
    placeholder,
  }: {
    value: string
    onChange: (value: string) => void
    placeholder?: string
  }) => (
    <textarea
      aria-label="Rich text editor"
      placeholder={placeholder}
      value={value}
      onChange={(event) => onChange(event.target.value)}
    />
  ),
}))

vi.mock('../components/cms/EntryActions', () => ({
  EntryActions: () => <div data-testid="entry-actions" />,
}))

vi.mock('../components/cms/PublishingControls', () => ({
  PublishingControls: ({ status }: { status: string }) => (
    <div data-testid="publishing-controls">{status}</div>
  ),
}))

vi.mock('react-hot-toast', () => ({
  default: {
    success: toastSuccess,
    error: toastError,
  },
}))

vi.mock('react-router-dom', async () => {
  const actual =
    await vi.importActual<typeof import('react-router-dom')>('react-router-dom')

  return {
    ...actual,
    useNavigate: () => mockNavigate,
  }
})

function renderPage(initialEntry = '/press/new') {
  return render(
    <MemoryRouter initialEntries={[initialEntry]}>
      <Routes>
        <Route path="/press/new" element={<PressEditorPage />} />
        <Route path="/press/:id/edit" element={<PressEditorPage mode="edit" />} />
      </Routes>
    </MemoryRouter>,
  )
}

function uploadFileForLabel(labelText: string, file: File) {
  const label = screen.getByText(labelText).closest('label')
  if (!label) {
    throw new Error(`Expected upload dropzone label for ${labelText}`)
  }

  const input = label.querySelector('input[type="file"]') as HTMLInputElement | null
  if (!input) {
    throw new Error(`Expected file input for ${labelText}`)
  }

  Object.defineProperty(input, 'files', {
    configurable: true,
    value: [file],
  })
  fireEvent.change(input)
}

beforeEach(async () => {
  mockNavigate.mockReset()
  mockCreate.mockReset()
  mockUpdate.mockReset()
  mockRemove.mockReset()
  toastSuccess.mockReset()
  toastError.mockReset()
  await i18n.changeLanguage('en')
})

describe('PressEditorPage', () => {
  it('rejects unsupported attachment uploads with the resources message', async () => {
    renderPage()

    uploadFileForLabel(
      'Drag and drop files here',
      new File(['audio'], 'voice.mp3', { type: 'audio/mpeg' }),
    )

    expect(
      (
        await screen.findAllByText(
          'Only PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP are supported.',
        )
      ).length,
    ).toBeGreaterThan(0)
    expect(toastError).toHaveBeenCalledWith(
      'Only PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP are supported.',
    )
    expect(screen.queryByText('voice.mp3')).toBeNull()
  })

  it('rejects oversized attachment uploads before they are added', async () => {
    renderPage()

    const largePdf = new File(['pdf'], 'large.pdf', {
      type: 'application/pdf',
    })
    Object.defineProperty(largePdf, 'size', {
      configurable: true,
      value: RESOURCE_UPLOAD_MAX_FILE_SIZE_BYTES + 1,
    })

    uploadFileForLabel('Drag and drop files here', largePdf)

    expect((await screen.findAllByText('This file exceeds the 20MB limit.')).length).toBeGreaterThan(0)
    expect(toastError).toHaveBeenCalledWith('This file exceeds the 20MB limit.')
    expect(screen.queryByText('large.pdf')).toBeNull()
  })

  it('rejects unsupported cover image uploads with the image-only message', async () => {
    renderPage()

    uploadFileForLabel(
      'Drag and drop a cover image here',
      new File(['gif'], 'cover.gif', { type: 'image/gif' }),
    )

    expect(
      (await screen.findAllByText('Only SVG, PNG, JPG, and WEBP are supported.')).length,
    ).toBeGreaterThan(0)
    expect(toastError).toHaveBeenCalledWith('Only SVG, PNG, JPG, and WEBP are supported.')
    expect(screen.queryByText('cover.gif')).toBeNull()
  })
})
