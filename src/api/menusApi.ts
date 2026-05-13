import { API_ROUTES } from '../constants/api'
import { apiClient } from './apiClient'

export type MenuNavigationType = 'pages' | 'external_link'

export type MenuPageReference = {
  id: number
  page_title: string
  url_slug: string
  parent_id: number | null
  status: string
}

export type MenuItemResponse = {
  id: number
  parent_id: number | null
  label: string
  navigation_type: MenuNavigationType
  page_id: number | null
  external_url: string
  open_in_new_tab: boolean
  sort_order: number
  href: string
  page?: MenuPageReference | null
  children: MenuItemResponse[]
}

export type MenuResponse = {
  id: number
  menu_key: string
  name: string
  items: MenuItemResponse[]
}

export type MenuPageOption = {
  id: number
  page_title: string
  url_slug: string
  parent_id: number | null
  parent_page_title: string
  status: string
}

export type MenuPageOptionsResponse = {
  items: MenuPageOption[]
}

export type SaveMenuItemRequest = {
  label: string
  navigation_type: MenuNavigationType
  page_id: number | null
  external_url: string
  open_in_new_tab: boolean
  children: SaveMenuItemRequest[]
}

export type SaveMenuRequest = {
  name: string
  items: SaveMenuItemRequest[]
}

export type SaveMenuResponse = {
  message: string
  menu: MenuResponse
}

export const menusApi = {
  async getMenu(key: string) {
    const response = await apiClient.get<MenuResponse>(API_ROUTES.menuByKey(key))
    return response.data
  },

  async listMenuPageOptions(key: string) {
    const response = await apiClient.get<MenuPageOptionsResponse>(
      API_ROUTES.menuPageOptionsByKey(key),
    )
    return Array.isArray(response.data.items) ? response.data.items : []
  },

  async saveMenu(key: string, payload: SaveMenuRequest) {
    const response = await apiClient.put<SaveMenuResponse>(API_ROUTES.menuByKey(key), payload)
    return response.data
  },
}
