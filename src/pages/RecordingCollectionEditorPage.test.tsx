import { fireEvent, render, screen, waitFor, within } from '@testing-library/react'
import { Provider } from 'react-redux'
import { MemoryRouter, Route, Routes } from 'react-router-dom'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import { createAppStore } from '../store/store'
import { RecordingCollectionEditorPage } from './RecordingCollectionEditorPage'

const {
  addRecordingItemsMock,
  createRecordingCollectionMock,
  deleteRecordingCollectionMock,
  deleteRecordingItemMock,
  getRecordingCollectionMock,
  toastSuccess,
  updateRecordingCollectionMock,
  updateRecordingItemMock,
} = vi.hoisted(() => ({
  addRecordingItemsMock: vi.fn(),
  createRecordingCollectionMock: vi.fn(),
  deleteRecordingCollectionMock: vi.fn(),
  deleteRecordingItemMock: vi.fn(),
  getRecordingCollectionMock: vi.fn(),
  toastSuccess: vi.fn(),
  updateRecordingCollectionMock: vi.fn(),
  updateRecordingItemMock: vi.fn(),
}))

vi.mock('react-hot-toast', () => ({
  default: {
    success: toastSuccess,
    error: vi.fn(),
  },
}))

vi.mock('../api/recordingsApi', () => ({
  recordingsApi: {
    addRecordingItems: addRecordingItemsMock,
    createRecordingCollection: createRecordingCollectionMock,
    deleteRecordingCollection: deleteRecordingCollectionMock,
    deleteRecordingItem: deleteRecordingItemMock,
    getRecordingCollection: getRecordingCollectionMock,
    updateRecordingCollection: updateRecordingCollectionMock,
    updateRecordingItem: updateRecordingItemMock,
  },
}))

const existingDetail = {
  id: 9,
  name: 'Town Halls',
  itemCount: 1,
  items: [
    {
      id: 14,
      recordingCollectionId: 9,
      title: 'Spring Town Hall',
      description: 'Notes',
      sortOrder: 0,
      createdAt: '2026-06-15T00:00:00Z',
      updatedAt: '2026-06-15T01:00:00Z',
    },
  ],
  createdAt: '2026-06-15T00:00:00Z',
  updatedAt: '2026-06-15T01:00:00Z',
}

function renderRoute(initialEntry = '/recordings/new') {
  return render(
    <Provider store={createAppStore()}>
      <MemoryRouter initialEntries={[initialEntry]}>
        <Routes>
          <Route path="/recordings/new" element={<RecordingCollectionEditorPage />} />
          <Route path="/recordings/:collectionId" element={<RecordingCollectionEditorPage />} />
        </Routes>
      </MemoryRouter>
    </Provider>,
  )
}

describe('RecordingCollectionEditorPage', () => {
  beforeEach(async () => {
    await i18n.changeLanguage('en')
    addRecordingItemsMock.mockReset()
    createRecordingCollectionMock.mockReset()
    deleteRecordingCollectionMock.mockReset()
    deleteRecordingItemMock.mockReset()
    getRecordingCollectionMock.mockReset()
    toastSuccess.mockReset()
    updateRecordingCollectionMock.mockReset()
    updateRecordingItemMock.mockReset()
  })

  it('creates a collection with a name and title/description items only', async () => {
    createRecordingCollectionMock.mockResolvedValue({
      message: 'Recording collection created successfully',
      recording: { id: 9, name: 'Weekly Updates' },
    })
    getRecordingCollectionMock.mockResolvedValue(existingDetail)

    renderRoute('/recordings/new')

    fireEvent.change(screen.getByPlaceholderText(/enter a collection name/i), {
      target: { value: '  Weekly Updates  ' },
    })
    fireEvent.click(screen.getByRole('button', { name: /add another item/i }))
    fireEvent.change(screen.getByPlaceholderText(/enter an item title/i), {
      target: { value: 'Week 1' },
    })
    fireEvent.change(screen.getByPlaceholderText(/optional supporting description/i), {
      target: { value: 'Kickoff recording' },
    })

    fireEvent.click(screen.getByRole('button', { name: /create collection/i }))

    await waitFor(() => {
      expect(createRecordingCollectionMock).toHaveBeenCalledWith({
        name: 'Weekly Updates',
        items: [{ title: 'Week 1', description: 'Kickoff recording' }],
      })
    })
  })

  it('blocks creation when the collection name is blank', async () => {
    renderRoute('/recordings/new')

    fireEvent.click(screen.getByRole('button', { name: /create collection/i }))

    expect(await screen.findByText('Collection name is required.')).toBeTruthy()
    expect(createRecordingCollectionMock).not.toHaveBeenCalled()
  })

  it('rejects a new item that has a description but no title', async () => {
    renderRoute('/recordings/new')

    fireEvent.change(screen.getByPlaceholderText(/enter a collection name/i), {
      target: { value: 'Weekly Updates' },
    })
    fireEvent.click(screen.getByRole('button', { name: /add another item/i }))
    fireEvent.change(screen.getByPlaceholderText(/optional supporting description/i), {
      target: { value: 'Description without a title' },
    })

    fireEvent.click(screen.getByRole('button', { name: /create collection/i }))

    expect(await screen.findByText('Each item needs a title.')).toBeTruthy()
    expect(createRecordingCollectionMock).not.toHaveBeenCalled()
  })

  it('loads a saved collection and updates a single item', async () => {
    getRecordingCollectionMock.mockResolvedValue(existingDetail)
    updateRecordingItemMock.mockResolvedValue({
      message: 'Recording item updated successfully',
      item: {
        id: 14,
        recordingCollectionId: 9,
        title: 'Renamed Item',
        description: 'Notes',
        sortOrder: 0,
        createdAt: '2026-06-15T00:00:00Z',
        updatedAt: '2026-06-15T02:00:00Z',
      },
    })

    renderRoute('/recordings/9')

    const titleInput = await screen.findByDisplayValue('Spring Town Hall')
    fireEvent.change(titleInput, { target: { value: 'Renamed Item' } })

    fireEvent.click(screen.getByRole('button', { name: /save item/i }))

    await waitFor(() => {
      expect(updateRecordingItemMock).toHaveBeenCalledWith(9, 14, {
        title: 'Renamed Item',
        description: 'Notes',
      })
    })
  })

  it('deletes a saved item only after the confirmation dialog', async () => {
    getRecordingCollectionMock.mockResolvedValue(existingDetail)
    deleteRecordingItemMock.mockResolvedValue({
      message: 'Recording item deleted',
      deletedCount: 1,
    })

    renderRoute('/recordings/9')
    await screen.findByDisplayValue('Spring Town Hall')

    fireEvent.click(screen.getByRole('button', { name: /delete item/i }))
    expect(deleteRecordingItemMock).not.toHaveBeenCalled()

    const dialog = await screen.findByRole('dialog')
    fireEvent.click(within(dialog).getByRole('button', { name: /delete item/i }))

    await waitFor(() => {
      expect(deleteRecordingItemMock).toHaveBeenCalledWith(9, 14)
    })
  })
})
