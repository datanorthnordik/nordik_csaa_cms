import { describe, it, expect, vi, beforeEach } from 'vitest'
import { render, screen, waitFor } from '@testing-library/react'
import { Provider } from 'react-redux'
import { BrowserRouter } from 'react-router-dom'
import { configureStore } from '@reduxjs/toolkit'
import { MediaLibraryRoute } from './MediaLibraryRoute'
import mediaReducer from '../store/mediaSlice'

const createMockStore = () =>
  configureStore({
    reducer: {
      media: mediaReducer,
    },
    preloadedState: {
      media: {
        list: {
          items: [],
          status: 'idle',
          error: null,
        },
        detail: {
          item: null,
          currentId: null,
          status: 'idle',
          error: null,
        },
        create: {
          status: 'idle',
          error: null,
          lastResult: null,
        },
        save: {
          status: 'idle',
          error: null,
          lastResult: null,
        },
        upload: {
          status: 'idle',
          error: null,
          lastResult: null,
        },
        assetUpdate: {
          status: 'idle',
          error: null,
          lastResult: null,
        },
        reorder: {
          status: 'idle',
          error: null,
          lastResult: null,
        },
        assetDelete: {
          status: 'idle',
          error: null,
          lastResult: null,
        },
        deleteGallery: {
          status: 'idle',
          error: null,
          lastResult: null,
        },
      },
    },
  })

describe('MediaLibraryRoute', () => {
  it('renders media library page', async () => {
    const store = createMockStore()
    render(
      <Provider store={store}>
        <BrowserRouter>
          <MediaLibraryRoute />
        </BrowserRouter>
      </Provider>
    )

    await waitFor(() => {
      expect(screen.getByRole('main')).toBeInTheDocument()
    })
  })

  it('displays loading state initially', async () => {
    const store = configureStore({
      reducer: { media: mediaReducer },
      preloadedState: {
        media: {
          list: {
            items: [],
            status: 'loading',
            error: null,
          },
          detail: {
            item: null,
            currentId: null,
            status: 'idle',
            error: null,
          },
          create: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
          save: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
          upload: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
          assetUpdate: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
          reorder: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
          assetDelete: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
          deleteGallery: {
            status: 'idle',
            error: null,
            lastResult: null,
          },
        },
      },
    })

    render(
      <Provider store={store}>
        <BrowserRouter>
          <MediaLibraryRoute />
        </BrowserRouter>
      </Provider>
    )

    await waitFor(() => {
      expect(screen.getByText(/loading/i)).toBeInTheDocument()
    })
  })
})
