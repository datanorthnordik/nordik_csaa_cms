import { fireEvent, render, screen } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { apiClient } from '../api/apiClient'
import i18n from '../i18n'
import { createAppStore } from '../store/store'
import { EventEditorPage } from './EventEditorPage'

const { toastError, toastSuccess } = vi.hoisted(() => ({
  toastError: vi.fn(),
  toastSuccess: vi.fn(),
}))

vi.mock('../api/apiClient', () => ({
  apiClient: {
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
    delete: vi.fn(),
  },
}))

vi.mock('react-hot-toast', () => ({
  default: {
    success: toastSuccess,
    error: toastError,
  },
}))

const mockedGet = vi.mocked(apiClient.get)
const mockedPost = vi.mocked(apiClient.post)

function renderPage(initialEntry = '/events/new') {
  const store = createAppStore()

  return render(
    <Provider store={store}>
      <MemoryRouter initialEntries={[initialEntry]}>
        <Routes>
          <Route path="/events/new" element={<EventEditorPage />} />
          <Route path="/events/:id" element={<EventEditorPage mode="view" />} />
          <Route path="/events/:id/edit" element={<EventEditorPage />} />
        </Routes>
      </MemoryRouter>
    </Provider>,
  )
}

beforeEach(async () => {
  mockedGet.mockReset()
  mockedPost.mockReset()
  toastError.mockReset()
  toastSuccess.mockReset()
  await i18n.changeLanguage('en')

  mockedGet.mockResolvedValue({ data: { items: [] } })
})

describe('EventEditorPage', () => {
  it('renders the shared breadcrumb on the create event route', async () => {
    renderPage()

    const breadcrumb = await screen.findByRole('navigation', { name: /breadcrumb/i })

    expect(screen.getByRole('heading', { name: /create new event/i })).toBeDefined()
    expect(breadcrumb.textContent).toContain('Events')
    expect(breadcrumb.textContent).toContain('Create new event')
  })

  it('renders Publish and Save Draft action buttons on the create route', async () => {
    renderPage()
    await screen.findByRole('heading', { name: /create new event/i })
    expect(screen.getByRole('button', { name: /publish/i })).toBeDefined()
    expect(screen.getByRole('button', { name: /save draft/i })).toBeDefined()
  })

  it('does not show a delete button on the create route (no entry to delete)', async () => {
    renderPage()
    await screen.findByRole('heading', { name: /create new event/i })
    expect(screen.queryByRole('button', { name: /delete event/i })).toBeNull()
  })

  it('does not render action buttons in view mode', async () => {
    // fetchEventLookups makes parallel calls to /api/events/locations and /api/events/galleries
    // before fetchEventById — use mockImplementation to route by URL instead of once-queue
    mockedGet.mockImplementation((url: string) => {
      if (url === '/api/events/1') {
        return Promise.resolve({
          data: {
            id: 1, title: 'Test Event', published: true, categories: [],
            event_type: 'single_day_all_day', start_at: '2026-05-01T00:00:00Z',
            privacy_type: 'public', private_audiences: [], request_review: false,
            review_email_list: [], teaser: '', description_html: '', contact_name: '',
            contact_email: '', contact_phone: '', contact_ext: '', contact_fax: '',
            location_mode: 'none', show_display_image_when_viewing: false,
            registration_enabled: false, registration_url: '', repeat_enabled: false,
            recurrence_interval: 1, occurrences: [], attachments: [],
            created_at: '2026-01-01T00:00:00Z', updated_at: '2026-01-02T00:00:00Z',
            show_title: true,
          },
        })
      }
      return Promise.resolve({ data: { items: [] } })
    })
    renderPage('/events/1')
    await screen.findByRole('navigation', { name: /breadcrumb/i })
    expect(screen.queryByRole('button', { name: /publish/i })).toBeNull()
    expect(screen.queryByRole('button', { name: /save draft/i })).toBeNull()
  })

  it('shows required validation messages inline when submit is attempted', async () => {
    renderPage()
    await screen.findByRole('heading', { name: /create new event/i })

    fireEvent.click(screen.getByRole('button', { name: /publish/i }))

    expect(mockedPost).not.toHaveBeenCalled()
    expect(toastError).toHaveBeenCalledWith(
      'Please fix the highlighted fields before continuing.',
    )
    expect((await screen.findAllByText('Title is required.')).length).toBeGreaterThan(0)
    expect((await screen.findAllByText('At least one category is required.')).length).toBeGreaterThan(0)
    expect((await screen.findAllByText('Start date is required.')).length).toBeGreaterThan(0)
  })

  it('shows inline errors for invalid contact email and phone values before submit', async () => {
    renderPage()
    await screen.findByRole('heading', { name: /create new event/i })

    fireEvent.change(screen.getByLabelText(/^title$/i), {
      target: { value: 'Spring Gathering' },
    })
    fireEvent.change(screen.getByLabelText(/^categories$/i), {
      target: { value: 'Community' },
    })
    fireEvent.change(screen.getByLabelText(/^start date$/i), {
      target: { value: '2026-05-07' },
    })
    fireEvent.change(screen.getByLabelText(/^email$/i), {
      target: { value: 'contact@example..com' },
    })
    fireEvent.change(screen.getByLabelText(/^phone$/i), {
      target: { value: '1234567' },
    })

    expect(mockedPost).not.toHaveBeenCalled()
    expect(toastError).not.toHaveBeenCalled()
    expect((await screen.findAllByText('Enter a valid contact email address.')).length).toBeGreaterThan(0)
    expect((await screen.findAllByText('Enter a valid contact phone number.')).length).toBeGreaterThan(0)
  })

  it('rejects unsupported display image and attachment files from picker and drag-and-drop', async () => {
    renderPage()
    await screen.findByRole('heading', { name: /create new event/i })

    const invalidDisplayImage = new File(['gif'], 'cover.gif', { type: 'image/gif' })
    fireEvent.change(screen.getByLabelText(/^display image$/i), {
      target: { files: [invalidDisplayImage] },
    })

    expect(toastError).toHaveBeenNthCalledWith(
      1,
      'Only SVG, PNG, JPG, and WEBP are supported.',
    )
    expect(screen.getByText('Only SVG, PNG, JPG, and WEBP are supported.')).toBeDefined()
    expect(screen.queryByText('cover.gif')).toBeNull()

    const invalidAttachment = new File(['exe'], 'agenda.exe', {
      type: 'application/octet-stream',
    })
    const attachmentDropzone = screen
      .getAllByText(/^additional files$/i)
      .map((element) => element.closest('label'))
      .find((element): element is HTMLLabelElement => element instanceof HTMLLabelElement)
    if (!attachmentDropzone) {
      throw new Error('Expected attachment dropzone')
    }

    fireEvent.drop(attachmentDropzone, {
      dataTransfer: { files: [invalidAttachment] },
    })

    expect(toastError).toHaveBeenNthCalledWith(
      2,
      'Only PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP are supported.',
    )
    expect(
      screen.getByText('Only PDF, DOCX, PPTX, XLSX, SVG, PNG, JPG, and WEBP are supported.'),
    ).toBeDefined()
    expect(screen.queryByText('agenda.exe')).toBeNull()
  })
})
