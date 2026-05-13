import { beforeEach, describe, expect, it, vi } from 'vitest'
import { menusApi, type MenuResponse, type SaveMenuRequest } from './menusApi'

const { getMock, putMock } = vi.hoisted(() => ({
  getMock: vi.fn(),
  putMock: vi.fn(),
}))

vi.mock('./apiClient', () => ({
  apiClient: {
    get: getMock,
    put: putMock,
  },
}))

function createMenuResponse(): MenuResponse {
  return {
    id: 1,
    menu_key: 'main',
    name: 'Main Website Navigation',
    items: [
      {
        id: 10,
        parent_id: null,
        label: 'About Us',
        navigation_type: 'pages',
        page_id: 100,
        external_url: '',
        open_in_new_tab: false,
        sort_order: 1,
        href: '/about-us',
        children: [],
      },
    ],
  }
}

describe('menusApi', () => {
  beforeEach(() => {
    getMock.mockReset()
    putMock.mockReset()
  })

  it('loads a menu by key', async () => {
    const menu = createMenuResponse()
    getMock.mockResolvedValue({ data: menu })

    const response = await menusApi.getMenu('main')

    expect(getMock).toHaveBeenCalledTimes(1)
    expect(getMock).toHaveBeenCalledWith('/api/menus/main')
    expect(response).toEqual(menu)
  })

  it('returns page options and falls back to an empty array when items is missing', async () => {
    const options = [
      {
        id: 100,
        page_title: 'About Us',
        url_slug: '/about-us',
        parent_id: null,
        parent_page_title: '',
        status: 'published',
      },
    ]
    getMock.mockResolvedValueOnce({
      data: {
        items: options,
      },
    })
    getMock.mockResolvedValueOnce({
      data: {},
    })

    const populated = await menusApi.listMenuPageOptions('main')
    const empty = await menusApi.listMenuPageOptions('main')

    expect(getMock).toHaveBeenNthCalledWith(1, '/api/menus/main/page-options')
    expect(getMock).toHaveBeenNthCalledWith(2, '/api/menus/main/page-options')
    expect(populated).toEqual(options)
    expect(empty).toEqual([])
  })

  it('saves a menu payload', async () => {
    const payload: SaveMenuRequest = {
      name: 'Main Website Navigation',
      items: [
        {
          label: 'About Us',
          navigation_type: 'pages',
          page_id: 100,
          external_url: '',
          open_in_new_tab: false,
          children: [],
        },
      ],
    }
    const response = {
      message: 'Navigation saved successfully.',
      menu: createMenuResponse(),
    }
    putMock.mockResolvedValue({ data: response })

    const result = await menusApi.saveMenu('main', payload)

    expect(putMock).toHaveBeenCalledTimes(1)
    expect(putMock).toHaveBeenCalledWith('/api/menus/main', payload)
    expect(result).toEqual(response)
  })
})
