import { createAsyncThunk, createSlice } from '@reduxjs/toolkit'
import { getApiErrorMessage } from '../api/apiError'
import {
  mediaApi,
  type CreateGalleryResponse,
  type DeleteGalleryImagesPayload,
  type DeleteGalleryImagesResponse,
  type DeleteGalleryResponse,
  type ReorderGalleryImagesPayload,
  type ReorderGalleryImagesResponse,
  type SaveGalleryRequest,
  type UpdateGalleryImagePayload,
  type UpdateGalleryImageResponse,
  type UpdateGalleryResponse,
  type UploadGalleryImagesRequest,
  type UploadGalleryImagesResponse,
} from '../api/mediaApi'
import type { GalleryDetail, GallerySummary } from '../types/media'
import type { RootState } from './store'

type AsyncState = 'idle' | 'loading' | 'succeeded' | 'failed'

type MediaState = {
  list: {
    items: GallerySummary[]
    status: AsyncState
    error: string | null
  }
  detail: {
    item: GalleryDetail | null
    currentId: number | null
    status: AsyncState
    error: string | null
  }
  create: {
    status: AsyncState
    error: string | null
    lastResult: CreateGalleryResponse | null
  }
  save: {
    status: AsyncState
    error: string | null
    lastResult: UpdateGalleryResponse | null
  }
  upload: {
    status: AsyncState
    error: string | null
    lastResult: UploadGalleryImagesResponse | null
  }
  assetUpdate: {
    status: AsyncState
    error: string | null
    lastResult: UpdateGalleryImageResponse | null
  }
  reorder: {
    status: AsyncState
    error: string | null
    lastResult: ReorderGalleryImagesResponse | null
  }
  assetDelete: {
    status: AsyncState
    error: string | null
    lastResult: DeleteGalleryImagesResponse | null
  }
  deleteGallery: {
    status: AsyncState
    error: string | null
    lastResult: DeleteGalleryResponse | null
  }
}

const initialState: MediaState = {
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
}

export const fetchMediaLibrary = createAsyncThunk<
  GallerySummary[],
  void,
  { rejectValue: string }
>('media/fetchMediaLibrary', async (_, thunkApi) => {
  try {
    return await mediaApi.listGalleries()
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const fetchGalleryById = createAsyncThunk<
  GalleryDetail,
  number,
  { rejectValue: string }
>('media/fetchGalleryById', async (id, thunkApi) => {
  try {
    return await mediaApi.getGallery(id)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const createGallery = createAsyncThunk<
  CreateGalleryResponse,
  SaveGalleryRequest,
  { rejectValue: string }
>('media/createGallery', async (payload, thunkApi) => {
  try {
    return await mediaApi.createGallery(payload)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const saveGallery = createAsyncThunk<
  UpdateGalleryResponse,
  { id: number; payload: SaveGalleryRequest },
  { rejectValue: string }
>('media/saveGallery', async ({ id, payload }, thunkApi) => {
  try {
    return await mediaApi.updateGallery(id, payload)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const deleteGallery = createAsyncThunk<
  { id: number; result: DeleteGalleryResponse },
  number,
  { rejectValue: string }
>('media/deleteGallery', async (id, thunkApi) => {
  try {
    const result = await mediaApi.deleteGallery(id)
    return { id, result }
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const uploadGalleryImages = createAsyncThunk<
  UploadGalleryImagesResponse,
  { id: number; payload: UploadGalleryImagesRequest },
  { rejectValue: string }
>('media/uploadGalleryImages', async ({ id, payload }, thunkApi) => {
  try {
    return await mediaApi.uploadGalleryImages(id, payload)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const updateGalleryAsset = createAsyncThunk<
  UpdateGalleryImageResponse,
  { galleryId: number; imageId: number; payload: UpdateGalleryImagePayload },
  { rejectValue: string }
>('media/updateGalleryAsset', async ({ galleryId, imageId, payload }, thunkApi) => {
  try {
    return await mediaApi.updateGalleryImage(galleryId, imageId, payload)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const reorderGalleryAssets = createAsyncThunk<
  ReorderGalleryImagesResponse,
  { id: number; payload: ReorderGalleryImagesPayload },
  { rejectValue: string }
>('media/reorderGalleryAssets', async ({ id, payload }, thunkApi) => {
  try {
    return await mediaApi.reorderGalleryImages(id, payload)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

export const deleteGalleryAssets = createAsyncThunk<
  DeleteGalleryImagesResponse,
  { id: number; payload: DeleteGalleryImagesPayload },
  { rejectValue: string }
>('media/deleteGalleryAssets', async ({ id, payload }, thunkApi) => {
  try {
    return await mediaApi.deleteGalleryImages(id, payload)
  } catch (error) {
    return thunkApi.rejectWithValue(getApiErrorMessage(error))
  }
})

const mediaSlice = createSlice({
  name: 'media',
  initialState,
  reducers: {
    clearCurrentGallery(state) {
      state.detail.item = null
      state.detail.currentId = null
      state.detail.status = 'idle'
      state.detail.error = null
    },
  },
  extraReducers: (builder) => {
    builder
      .addCase(fetchMediaLibrary.pending, (state) => {
        state.list.status = 'loading'
        state.list.error = null
      })
      .addCase(fetchMediaLibrary.fulfilled, (state, action) => {
        state.list.status = 'succeeded'
        state.list.items = action.payload
      })
      .addCase(fetchMediaLibrary.rejected, (state, action) => {
        state.list.status = 'failed'
        state.list.error = action.payload ?? 'Could not load galleries.'
      })
      .addCase(fetchGalleryById.pending, (state, action) => {
        state.detail.status = 'loading'
        state.detail.error = null
        state.detail.currentId = action.meta.arg
        if (state.detail.item?.id !== action.meta.arg) {
          state.detail.item = null
        }
      })
      .addCase(fetchGalleryById.fulfilled, (state, action) => {
        state.detail.status = 'succeeded'
        state.detail.item = action.payload
        state.detail.currentId = action.payload.id
      })
      .addCase(fetchGalleryById.rejected, (state, action) => {
        state.detail.status = 'failed'
        state.detail.error = action.payload ?? 'Could not load the gallery.'
        state.detail.item = null
      })
      .addCase(createGallery.pending, (state) => {
        state.create.status = 'loading'
        state.create.error = null
      })
      .addCase(createGallery.fulfilled, (state, action) => {
        state.create.status = 'succeeded'
        state.create.lastResult = action.payload
      })
      .addCase(createGallery.rejected, (state, action) => {
        state.create.status = 'failed'
        state.create.error = action.payload ?? 'Could not create the gallery.'
      })
      .addCase(saveGallery.pending, (state) => {
        state.save.status = 'loading'
        state.save.error = null
      })
      .addCase(saveGallery.fulfilled, (state, action) => {
        state.save.status = 'succeeded'
        state.save.lastResult = action.payload
      })
      .addCase(saveGallery.rejected, (state, action) => {
        state.save.status = 'failed'
        state.save.error = action.payload ?? 'Could not save the gallery.'
      })
      .addCase(uploadGalleryImages.pending, (state) => {
        state.upload.status = 'loading'
        state.upload.error = null
      })
      .addCase(uploadGalleryImages.fulfilled, (state, action) => {
        state.upload.status = 'succeeded'
        state.upload.lastResult = action.payload
      })
      .addCase(uploadGalleryImages.rejected, (state, action) => {
        state.upload.status = 'failed'
        state.upload.error = action.payload ?? 'Could not upload gallery images.'
      })
      .addCase(updateGalleryAsset.pending, (state) => {
        state.assetUpdate.status = 'loading'
        state.assetUpdate.error = null
      })
      .addCase(updateGalleryAsset.fulfilled, (state, action) => {
        state.assetUpdate.status = 'succeeded'
        state.assetUpdate.lastResult = action.payload
      })
      .addCase(updateGalleryAsset.rejected, (state, action) => {
        state.assetUpdate.status = 'failed'
        state.assetUpdate.error = action.payload ?? 'Could not update the gallery image.'
      })
      .addCase(reorderGalleryAssets.pending, (state) => {
        state.reorder.status = 'loading'
        state.reorder.error = null
      })
      .addCase(reorderGalleryAssets.fulfilled, (state, action) => {
        state.reorder.status = 'succeeded'
        state.reorder.lastResult = action.payload
      })
      .addCase(reorderGalleryAssets.rejected, (state, action) => {
        state.reorder.status = 'failed'
        state.reorder.error = action.payload ?? 'Could not reorder gallery images.'
      })
      .addCase(deleteGalleryAssets.pending, (state) => {
        state.assetDelete.status = 'loading'
        state.assetDelete.error = null
      })
      .addCase(deleteGalleryAssets.fulfilled, (state, action) => {
        state.assetDelete.status = 'succeeded'
        state.assetDelete.lastResult = action.payload
      })
      .addCase(deleteGalleryAssets.rejected, (state, action) => {
        state.assetDelete.status = 'failed'
        state.assetDelete.error = action.payload ?? 'Could not delete the gallery image.'
      })
      .addCase(deleteGallery.pending, (state) => {
        state.deleteGallery.status = 'loading'
        state.deleteGallery.error = null
      })
      .addCase(deleteGallery.fulfilled, (state, action) => {
        state.deleteGallery.status = 'succeeded'
        state.deleteGallery.lastResult = action.payload.result
        state.list.items = state.list.items.filter((item) => item.id !== action.payload.id)
        if (state.detail.item?.id === action.payload.id) {
          state.detail.item = null
          state.detail.currentId = null
        }
      })
      .addCase(deleteGallery.rejected, (state, action) => {
        state.deleteGallery.status = 'failed'
        state.deleteGallery.error = action.payload ?? 'Could not delete the gallery.'
      })
  },
})

export const { clearCurrentGallery } = mediaSlice.actions

export const selectMediaLibrary = (state: RootState) => state.media.list
export const selectMediaGalleryDetail = (state: RootState) => state.media.detail
export const selectCurrentGallery = (state: RootState) => state.media.detail.item
export const selectMediaCreate = (state: RootState) => state.media.create
export const selectMediaIsSaving = (state: RootState) =>
  [
    state.media.save.status,
    state.media.upload.status,
    state.media.assetUpdate.status,
    state.media.reorder.status,
    state.media.assetDelete.status,
    state.media.deleteGallery.status,
  ].some((status) => status === 'loading')

export default mediaSlice.reducer
