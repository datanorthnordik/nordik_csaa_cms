import {
  createAsyncThunk,
  createSlice,
  type PayloadAction,
} from '@reduxjs/toolkit'
import { getApiErrorMessage } from '../api/apiError'
import {
  pagesApi,
  type PageDetailResponse,
  type PageListFilters,
  type PageListPageMeta,
  type PageListResponse,
  type PageListItem,
  type PageMutationResponse,
  type SavePagePayload,
} from '../api/pagesApi'
import type { RootState } from './store'

export const defaultPageListFilters: PageListFilters = {
  page: 1,
  pageSize: 10,
  searchTerm: '',
  status: '',
  sortBy: 'last_modified',
  sortOrder: 'desc',
}

type AsyncState = 'idle' | 'loading' | 'succeeded' | 'failed'

type PagesState = {
  list: {
    items: PageListItem[]
    pagination: PageListPageMeta | null
    appliedFilters: PageListResponse['applied_filters'] | null
    filters: PageListFilters
    status: AsyncState
    error: string | null
  }
  detail: {
    item: PageDetailResponse | null
    status: AsyncState
    error: string | null
  }
  save: {
    status: AsyncState
    error: string | null
    lastResult: PageMutationResponse | null
  }
}

const initialState: PagesState = {
  list: {
    items: [],
    pagination: null,
    appliedFilters: null,
    filters: defaultPageListFilters,
    status: 'idle',
    error: null,
  },
  detail: {
    item: null,
    status: 'idle',
    error: null,
  },
  save: {
    status: 'idle',
    error: null,
    lastResult: null,
  },
}

export const fetchPages = createAsyncThunk<
  PageListResponse,
  PageListFilters,
  { rejectValue: string }
>('pages/fetchPages', async (filters, thunkApi) => {
  try {
    return await pagesApi.listPages(filters)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const fetchPageById = createAsyncThunk<
  PageDetailResponse,
  number,
  { rejectValue: string }
>('pages/fetchPageById', async (id, thunkApi) => {
  try {
    return await pagesApi.getPage(id)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const createPage = createAsyncThunk<
  PageMutationResponse,
  SavePagePayload,
  { rejectValue: string }
>('pages/createPage', async (payload, thunkApi) => {
  try {
    return await pagesApi.createPage(payload)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const updatePage = createAsyncThunk<
  PageMutationResponse,
  { id: number; payload: SavePagePayload },
  { rejectValue: string }
>('pages/updatePage', async ({ id, payload }, thunkApi) => {
  try {
    return await pagesApi.updatePage(id, payload)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const deletePage = createAsyncThunk<
  number,
  number,
  { rejectValue: string }
>('pages/deletePage', async (id, thunkApi) => {
  try {
    await pagesApi.deletePage(id)
    return id
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

const pagesSlice = createSlice({
  name: 'pages',
  initialState,
  reducers: {
    setPageListFilters(state, action: PayloadAction<Partial<PageListFilters>>) {
      state.list.filters = {
        ...state.list.filters,
        ...action.payload,
      }
    },
    clearCurrentPage(state) {
      state.detail.item = null
      state.detail.status = 'idle'
      state.detail.error = null
    },
    clearPageSaveState(state) {
      state.save.status = 'idle'
      state.save.error = null
      state.save.lastResult = null
    },
  },
  extraReducers: (builder) => {
    builder
      .addCase(fetchPages.pending, (state) => {
        state.list.status = 'loading'
        state.list.error = null
      })
      .addCase(fetchPages.fulfilled, (state, action) => {
        state.list.status = 'succeeded'
        state.list.items = action.payload.items
        state.list.pagination = action.payload.pagination
        state.list.appliedFilters = action.payload.applied_filters
      })
      .addCase(fetchPages.rejected, (state, action) => {
        state.list.status = 'failed'
        state.list.error = action.payload ?? 'Could not load pages.'
      })
      .addCase(fetchPageById.pending, (state) => {
        state.detail.status = 'loading'
        state.detail.error = null
      })
      .addCase(fetchPageById.fulfilled, (state, action) => {
        state.detail.status = 'succeeded'
        state.detail.item = action.payload
      })
      .addCase(fetchPageById.rejected, (state, action) => {
        state.detail.status = 'failed'
        state.detail.error = action.payload ?? 'Could not load page details.'
      })
      .addCase(createPage.pending, (state) => {
        state.save.status = 'loading'
        state.save.error = null
      })
      .addCase(createPage.fulfilled, (state, action) => {
        state.save.status = 'succeeded'
        state.save.lastResult = action.payload
      })
      .addCase(createPage.rejected, (state, action) => {
        state.save.status = 'failed'
        state.save.error = action.payload ?? 'Could not create the page.'
      })
      .addCase(updatePage.pending, (state) => {
        state.save.status = 'loading'
        state.save.error = null
      })
      .addCase(updatePage.fulfilled, (state, action) => {
        state.save.status = 'succeeded'
        state.save.lastResult = action.payload
      })
      .addCase(updatePage.rejected, (state, action) => {
        state.save.status = 'failed'
        state.save.error = action.payload ?? 'Could not update the page.'
      })
      .addCase(deletePage.fulfilled, (state, action) => {
        state.list.items = state.list.items.filter((item) => item.id !== action.payload)
        if (state.detail.item?.id === action.payload) {
          state.detail.item = null
        }
      })
  },
})

export const {
  setPageListFilters,
  clearCurrentPage,
  clearPageSaveState,
} = pagesSlice.actions

export const selectPageList = (state: RootState) => state.pages.list
export const selectPageListFilters = (state: RootState) => state.pages.list.filters
export const selectPageDetail = (state: RootState) => state.pages.detail
export const selectPageSave = (state: RootState) => state.pages.save
export const selectCurrentPage = (state: RootState) => state.pages.detail.item

export default pagesSlice.reducer
