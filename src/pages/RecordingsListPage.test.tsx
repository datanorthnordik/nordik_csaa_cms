import { fireEvent, render, screen, waitFor, within } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import { createAppStore } from '../store/store'
import { RecordingsListPage } from './RecordingsListPage'

const { deleteRecordingCollectionMock, listRecordingCollectionsMock, toastSuccess } = vi.hoisted(
  () => ({
    deleteRecordingCollectionMock: vi.fn(),
    listRecordingCollectionsMock: vi.fn(),
    toastSuccess: vi.fn(),
  }),
)

vi.mock('react-hot-toast', () => ({
  default: {
    success: toastSuccess,
    error: vi.fn(),
  },
}))

vi.mock('../api/recordingsApi', () => ({
  recordingsApi: {
    deleteRecordingCollection: deleteRecordingCollectionMock,
    listRecordingCollections: listRecordingCollectionsMock,
  },
}))

function renderPage() {
  return render(
    <Provider store={createAppStore()}>
      <MemoryRouter initialEntries={['/recordings']}>
        <RecordingsListPage />
      </MemoryRouter>
    </Provider>,
  )
}

describe('RecordingsListPage', () => {
  beforeEach(async () => {
    await i18n.changeLanguage('en')
    deleteRecordingCollectionMock.mockReset()
    listRecordingCollectionsMock.mockReset()
    toastSuccess.mockReset()
  })

  it('renders each recording collection with its item count', async () => {
    listRecordingCollectionsMock.mockResolvedValue([
      {
        id: 2,
        name: 'Board Meetings',
        itemCount: 4,
        createdAt: '2026-06-15T00:00:00Z',
        updatedAt: '2026-06-15T01:00:00Z',
      },
      {
        id: 1,
        name: 'Annual General Meeting',
        itemCount: 1,
        createdAt: '2026-06-14T00:00:00Z',
        updatedAt: '2026-06-14T01:00:00Z',
      },
    ])

    renderPage()

    expect(await screen.findByText('Board Meetings')).toBeTruthy()
    expect(screen.getByText('Annual General Meeting')).toBeTruthy()
    expect(screen.getByText('4 items')).toBeTruthy()
    expect(screen.getByText('1 item')).toBeTruthy()
  })

  it('filters collections by the search term', async () => {
    listRecordingCollectionsMock.mockResolvedValue([
      { id: 2, name: 'Board Meetings', itemCount: 4, createdAt: '', updatedAt: '2026-06-15T01:00:00Z' },
      {
        id: 1,
        name: 'Annual General Meeting',
        itemCount: 1,
        createdAt: '',
        updatedAt: '2026-06-14T01:00:00Z',
      },
    ])

    renderPage()
    await screen.findByText('Board Meetings')

    fireEvent.change(screen.getByRole('searchbox'), { target: { value: 'board' } })

    expect(screen.getByText('Board Meetings')).toBeTruthy()
    expect(screen.queryByText('Annual General Meeting')).toBeNull()
  })

  it('deletes the selected collection after confirmation', async () => {
    listRecordingCollectionsMock.mockResolvedValue([
      { id: 2, name: 'Board Meetings', itemCount: 4, createdAt: '', updatedAt: '2026-06-15T01:00:00Z' },
    ])
    deleteRecordingCollectionMock.mockResolvedValue({ message: 'Recording collection deleted' })

    renderPage()
    await screen.findByText('Board Meetings')

    fireEvent.click(screen.getByRole('button', { name: /delete collection/i }))

    const dialog = await screen.findByRole('dialog')
    fireEvent.click(within(dialog).getByRole('button', { name: /delete collection/i }))

    await waitFor(() => {
      expect(deleteRecordingCollectionMock).toHaveBeenCalledWith(2)
    })
  })
})
