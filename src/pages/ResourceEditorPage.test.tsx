import type { ReactNode } from 'react'
import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import toast from 'react-hot-toast'
import i18n from '../i18n'
import { RESOURCE_UPLOAD_MAX_FILE_SIZE_BYTES } from '../lib/resourceUpload'
import { ResourceEditorPage } from './ResourceEditorPage'

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

vi.mock('../hooks/useResources', () => ({
  useResources: () => ({
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

function renderPage(initialEntry = '/resources/new') {
  return render(
    <MemoryRouter initialEntries={[initialEntry]}>
      <Routes>
        <Route path="/resources/new" element={<ResourceEditorPage />} />
        <Route
          path="/resources/:id/edit"
          element={<ResourceEditorPage mode="edit" />}
        />
      </Routes>
    </MemoryRouter>,
  )
}

function uploadFile(file: File) {
  const input = document.querySelector('input[type="file"]') as HTMLInputElement
  Object.defineProperty(input, 'files', {
    configurable: true,
    value: [file],
  })
  fireEvent.change(input)
}

function dropFile(file: File) {
  const dropzone = screen.getByText('Drag and drop files here').closest('label')
  if (!dropzone) {
    throw new Error('Expected upload dropzone label')
  }

  fireEvent.drop(dropzone, {
    dataTransfer: {
      files: [file],
    },
  })
}

beforeEach(async () => {
  mockNavigate.mockReset()
  mockCreate.mockReset()
  mockUpdate.mockReset()
  mockRemove.mockReset()
  toastSuccess.mockReset()
  toastError.mockReset()
  mockCreate.mockResolvedValue({ id: '123' })
  URL.createObjectURL = vi.fn(() => 'blob:resource-preview')
  URL.revokeObjectURL = vi.fn()
  window.localStorage.clear()
  window.sessionStorage.clear()
  await i18n.changeLanguage('en')
})

describe('ResourceEditorPage', () => {
  it('rejects files larger than the configured size limit', async () => {
    renderPage()

    const largePdf = new File(['pdf'], 'large.pdf', {
      type: 'application/pdf',
    })
    Object.defineProperty(largePdf, 'size', {
      configurable: true,
      value: RESOURCE_UPLOAD_MAX_FILE_SIZE_BYTES + 1,
    })

    uploadFile(largePdf)

    expect((await screen.findAllByText('This file exceeds the 20MB limit.')).length).toBeGreaterThan(0)
    expect(toast.error).toHaveBeenCalledWith('This file exceeds the 20MB limit.')
    expect(screen.queryByText('large.pdf')).toBeNull()
  })

  it('rejects unsupported audio and video uploads', async () => {
    renderPage()

    uploadFile(new File(['video'], 'clip.mp4', { type: 'video/mp4' }))

    expect(
      (
        await screen.findAllByText(
          'Only PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP are supported.',
        )
      ).length,
    ).toBeGreaterThan(0)
    expect(toast.error).toHaveBeenCalledWith(
      'Only PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP are supported.',
    )
    expect(screen.queryByText('clip.mp4')).toBeNull()
  })

  it('rejects unsupported files dropped onto the upload area', async () => {
    renderPage()

    dropFile(new File(['audio'], 'voice.mp3', { type: 'audio/mpeg' }))

    expect(
      (
        await screen.findAllByText(
          'Only PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP are supported.',
        )
      ).length,
    ).toBeGreaterThan(0)
    expect(toast.error).toHaveBeenCalledWith(
      'Only PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP are supported.',
    )
    expect(screen.queryByText('voice.mp3')).toBeNull()
  })

  it('submits supported files within the limit', async () => {
    renderPage()

    fireEvent.change(screen.getByPlaceholderText('Enter file display name'), {
      target: { value: 'Board package' },
    })
    fireEvent.change(
      screen.getByPlaceholderText('Describe what this resource is for'),
      {
        target: { value: 'Quarterly board package' },
      },
    )
    fireEvent.change(screen.getByRole('combobox'), {
      target: { value: 'report' },
    })

    const validFile = new File(['png'], 'board-package.png', {
      type: 'image/png',
    })
    uploadFile(validFile)

    fireEvent.click(screen.getByRole('button', { name: 'Start Upload' }))

    await waitFor(() => {
      expect(mockCreate).toHaveBeenCalledWith(
        {
          name: 'Board package',
          description: 'Quarterly board package',
          category: 'report',
          visibility: 'public',
          linkUrl: '',
        },
        validFile,
      )
    })
    expect(mockNavigate).toHaveBeenCalledWith('/resources/123/edit', {
      replace: true,
    })
  })
})
