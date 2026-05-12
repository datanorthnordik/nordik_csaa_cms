import {
  createAsyncThunk,
  createSlice,
  type PayloadAction,
} from '@reduxjs/toolkit'
import { getApiErrorMessage } from '../api/apiError'
import {
  eventsApi,
  type EventDetailResponse,
  type EventListFilters,
  type EventListPageMeta,
  type EventListResponse,
  type EventListItem,
  type EventMedia,
  type EventMutationResponse,
  type EventAddress,
  type Gallery,
  type SaveEventRequest,
} from '../api/eventsApi'
import type { RootState } from './store'

export const defaultEventListFilters: EventListFilters = {
  page: 1,
  pageSize: 10,
  searchTerm: '',
  statuses: [],
  startDate: '',
  endDate: '',
  dateRange: 'custom',
  sortBy: 'start_at',
  sortOrder: 'desc',
}

type AsyncState = 'idle' | 'loading' | 'succeeded' | 'failed'

type EventsState = {
  list: {
    items: EventListItem[]
    pagination: EventListPageMeta | null
    appliedFilters: EventListResponse['applied_filters'] | null
    filters: EventListFilters
    status: AsyncState
    error: string | null
  }
  detail: {
    item: EventDetailResponse | null
    status: AsyncState
    error: string | null
  }
  lookups: {
    locations: EventAddress[]
    galleries: Gallery[]
    status: AsyncState
    error: string | null
  }
  save: {
    status: AsyncState
    error: string | null
    lastResult: EventMutationResponse | null
  }
  mediaDelete: {
    deletingIds: number[]
    error: string | null
  }
}

const initialState: EventsState = {
  list: {
    items: [],
    pagination: null,
    appliedFilters: null,
    filters: defaultEventListFilters,
    status: 'idle',
    error: null,
  },
  detail: {
    item: null,
    status: 'idle',
    error: null,
  },
  lookups: {
    locations: [],
    galleries: [],
    status: 'idle',
    error: null,
  },
  save: {
    status: 'idle',
    error: null,
    lastResult: null,
  },
  mediaDelete: {
    deletingIds: [],
    error: null,
  },
}

export const fetchEvents = createAsyncThunk<
  EventListResponse,
  EventListFilters,
  { rejectValue: string }
>('events/fetchEvents', async (filters, thunkApi) => {
  try {
    return await eventsApi.listEvents(filters)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const fetchEventById = createAsyncThunk<
  EventDetailResponse,
  number,
  { rejectValue: string }
>('events/fetchEventById', async (id, thunkApi) => {
  try {
    return await eventsApi.getEvent(id)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const fetchEventLookups = createAsyncThunk<
  { locations: EventAddress[]; galleries: Gallery[] },
  void,
  { rejectValue: string }
>('events/fetchEventLookups', async (_, thunkApi) => {
  try {
    const [locations, galleries] = await Promise.all([
      eventsApi.getSavedLocations(),
      eventsApi.getGalleries(),
    ])

    return {
      locations,
      galleries,
    }
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const createEvent = createAsyncThunk<
  EventMutationResponse,
  SaveEventRequest,
  { rejectValue: string }
>('events/createEvent', async (payload, thunkApi) => {
  try {
    return await eventsApi.createEvent(payload)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const updateEvent = createAsyncThunk<
  EventMutationResponse,
  { id: number; payload: SaveEventRequest },
  { rejectValue: string }
>('events/updateEvent', async ({ id, payload }, thunkApi) => {
  try {
    return await eventsApi.updateEvent(id, payload)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const deleteEvent = createAsyncThunk<
  number,
  number,
  { rejectValue: string }
>('events/deleteEvent', async (id, thunkApi) => {
  try {
    await eventsApi.deleteEvent(id)
    return id
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const deleteEventDocument = createAsyncThunk<
  { eventId: number; mediaId: number },
  { eventId: number; mediaId: number },
  { rejectValue: string }
>('events/deleteEventDocument', async ({ eventId, mediaId }, thunkApi) => {
  try {
    await eventsApi.deleteEventDocument(eventId, mediaId)
    return { eventId, mediaId }
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const deleteEventPhoto = createAsyncThunk<
  { eventId: number; mediaId: number },
  { eventId: number; mediaId: number },
  { rejectValue: string }
>('events/deleteEventPhoto', async ({ eventId, mediaId }, thunkApi) => {
  try {
    await eventsApi.deleteEventPhoto(eventId, mediaId)
    return { eventId, mediaId }
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

const eventsSlice = createSlice({
  name: 'events',
  initialState,
  reducers: {
    setEventListFilters(state, action: PayloadAction<Partial<EventListFilters>>) {
      state.list.filters = {
        ...state.list.filters,
        ...action.payload,
      }
    },
    clearCurrentEvent(state) {
      state.detail.item = null
      state.detail.status = 'idle'
      state.detail.error = null
    },
    clearEventSaveState(state) {
      state.save.status = 'idle'
      state.save.error = null
      state.save.lastResult = null
    },
    removeLocalAttachment(state, action: PayloadAction<number>) {
      if (!state.detail.item) {
        return
      }
      state.detail.item.attachments = state.detail.item.attachments.filter(
        (item) => item.id !== action.payload,
      )
    },
    removeLocalDisplayImage(state, action: PayloadAction<number>) {
      if (state.detail.item?.display_image?.id === action.payload) {
        state.detail.item.display_image = null
      }
    },
  },
  extraReducers: (builder) => {
    builder
      .addCase(fetchEvents.pending, (state) => {
        state.list.status = 'loading'
        state.list.error = null
      })
      .addCase(fetchEvents.fulfilled, (state, action) => {
        state.list.status = 'succeeded'
        state.list.items = action.payload.items
        state.list.pagination = action.payload.pagination
        state.list.appliedFilters = action.payload.applied_filters
      })
      .addCase(fetchEvents.rejected, (state, action) => {
        state.list.status = 'failed'
        state.list.error = action.payload ?? 'Could not load events.'
      })
      .addCase(fetchEventById.pending, (state) => {
        state.detail.status = 'loading'
        state.detail.error = null
      })
      .addCase(fetchEventById.fulfilled, (state, action) => {
        state.detail.status = 'succeeded'
        state.detail.item = action.payload
      })
      .addCase(fetchEventById.rejected, (state, action) => {
        state.detail.status = 'failed'
        state.detail.error = action.payload ?? 'Could not load event details.'
      })
      .addCase(fetchEventLookups.pending, (state) => {
        state.lookups.status = 'loading'
        state.lookups.error = null
      })
      .addCase(fetchEventLookups.fulfilled, (state, action) => {
        state.lookups.status = 'succeeded'
        state.lookups.locations = action.payload.locations
        state.lookups.galleries = action.payload.galleries
      })
      .addCase(fetchEventLookups.rejected, (state, action) => {
        state.lookups.status = 'failed'
        state.lookups.error = action.payload ?? 'Could not load event lookups.'
      })
      .addCase(createEvent.pending, (state) => {
        state.save.status = 'loading'
        state.save.error = null
      })
      .addCase(createEvent.fulfilled, (state, action) => {
        state.save.status = 'succeeded'
        state.save.lastResult = action.payload
      })
      .addCase(createEvent.rejected, (state, action) => {
        state.save.status = 'failed'
        state.save.error = action.payload ?? 'Could not save the event.'
      })
      .addCase(updateEvent.pending, (state) => {
        state.save.status = 'loading'
        state.save.error = null
      })
      .addCase(updateEvent.fulfilled, (state, action) => {
        state.save.status = 'succeeded'
        state.save.lastResult = action.payload
      })
      .addCase(updateEvent.rejected, (state, action) => {
        state.save.status = 'failed'
        state.save.error = action.payload ?? 'Could not update the event.'
      })
      .addCase(deleteEvent.fulfilled, (state, action) => {
        state.list.items = state.list.items.filter(
          (item) => item.id !== action.payload,
        )
        if (state.detail.item?.id === action.payload) {
          state.detail.item = null
        }
      })
      .addCase(deleteEventDocument.pending, (state, action) => {
        state.mediaDelete.error = null
        state.mediaDelete.deletingIds.push(action.meta.arg.mediaId)
      })
      .addCase(deleteEventDocument.fulfilled, (state, action) => {
        state.mediaDelete.deletingIds = state.mediaDelete.deletingIds.filter(
          (id) => id !== action.payload.mediaId,
        )
        if (state.detail.item) {
          state.detail.item.attachments = state.detail.item.attachments.filter(
            (item) => item.id !== action.payload.mediaId,
          )
        }
      })
      .addCase(deleteEventDocument.rejected, (state, action) => {
        state.mediaDelete.deletingIds = state.mediaDelete.deletingIds.filter(
          (id) => id !== action.meta.arg.mediaId,
        )
        state.mediaDelete.error =
          action.payload ?? 'Could not remove the attachment.'
      })
      .addCase(deleteEventPhoto.pending, (state, action) => {
        state.mediaDelete.error = null
        state.mediaDelete.deletingIds.push(action.meta.arg.mediaId)
      })
      .addCase(deleteEventPhoto.fulfilled, (state, action) => {
        state.mediaDelete.deletingIds = state.mediaDelete.deletingIds.filter(
          (id) => id !== action.payload.mediaId,
        )
        if (state.detail.item?.display_image?.id === action.payload.mediaId) {
          state.detail.item.display_image = null
        }
      })
      .addCase(deleteEventPhoto.rejected, (state, action) => {
        state.mediaDelete.deletingIds = state.mediaDelete.deletingIds.filter(
          (id) => id !== action.meta.arg.mediaId,
        )
        state.mediaDelete.error =
          action.payload ?? 'Could not remove the display image.'
      })
  },
})

export const {
  setEventListFilters,
  clearCurrentEvent,
  clearEventSaveState,
  removeLocalAttachment,
  removeLocalDisplayImage,
} = eventsSlice.actions

export const selectEventList = (state: RootState) => state.events.list
export const selectEventListFilters = (state: RootState) => state.events.list.filters
export const selectEventDetail = (state: RootState) => state.events.detail
export const selectEventLookups = (state: RootState) => state.events.lookups
export const selectEventSave = (state: RootState) => state.events.save
export const selectEventMediaDelete = (state: RootState) => state.events.mediaDelete
export const selectEventDeletingIds = (state: RootState) =>
  state.events.mediaDelete.deletingIds
export const selectCurrentEvent = (state: RootState) => state.events.detail.item
export const selectSavedLocations = (state: RootState) =>
  state.events.lookups.locations
export const selectEventGalleries = (state: RootState) =>
  state.events.lookups.galleries
export const selectExistingAttachments = (state: RootState) =>
  state.events.detail.item?.attachments ?? []
export const selectExistingDisplayImage = (state: RootState): EventMedia | null =>
  state.events.detail.item?.display_image ?? null

export default eventsSlice.reducer
