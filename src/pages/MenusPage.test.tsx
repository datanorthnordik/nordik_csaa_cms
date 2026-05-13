import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import type { ReactNode } from 'react'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import type { MenuPageOption, MenuResponse } from '../api/menusApi'
import type { MenuEditorItem } from '../lib/menusEditor'
import { MenusPage } from './MenusPage'

const {
  getMenuMock,
  listMenuPageOptionsMock,
  saveMenuMock,
  toastSuccess,
  toastError,
  getApiErrorMessageMock,
  menuBuilderPropsSpy,
} = vi.hoisted(() => ({
  getMenuMock: vi.fn(),
  listMenuPageOptionsMock: vi.fn(),
  saveMenuMock: vi.fn(),
  toastSuccess: vi.fn(),
  toastError: vi.fn(),
  getApiErrorMessageMock: vi.fn(),
  menuBuilderPropsSpy: vi.fn(),
}))

let nextBuilderItems: MenuEditorItem[] = []

vi.mock('react-hot-toast', () => ({
  default: {
    success: toastSuccess,
    error: toastError,
  },
}))

vi.mock('../api/apiError', () => ({
  getApiErrorMessage: getApiErrorMessageMock,
}))

vi.mock('../api/menusApi', () => ({
  menusApi: {
    getMenu: getMenuMock,
    listMenuPageOptions: listMenuPageOptionsMock,
    saveMenu: saveMenuMock,
  },
}))

vi.mock('../components/Breadcrumb', () => ({
  Breadcrumb: ({ items }: { items: Array<{ label: string }> }) => (
    <nav data-testid="breadcrumb">{items.map((item) => item.label).join(' / ')}</nav>
  ),
}))

vi.mock('../components/CmsAppShell', () => ({
  CmsAppShell: ({
    activeKey,
    children,
  }: {
    activeKey?: string
    children?: ReactNode
  }) => (
    <div data-testid="app-shell" data-active-key={activeKey}>
      {children}
    </div>
  ),
}))

vi.mock('../components/Loader', () => ({
  Loader: ({ label }: { label: string }) => <div data-testid="loader">{label}</div>,
}))

vi.mock('../components/navigation/MenuBuilder', () => ({
  MenuBuilder: (props: {
    title: string
    subtitle: string
    isDirty: boolean
    isSaving: boolean
    saveError?: string | null
    onChange: (items: MenuEditorItem[]) => void
    onSave: () => void
    onDiscard: () => void
  }) => {
    menuBuilderPropsSpy(props)
    return (
      <section data-testid="menu-builder">
        <h2>{props.title}</h2>
        <p>{props.subtitle}</p>
        <p data-testid="dirty-state">{String(props.isDirty)}</p>
        <p data-testid="saving-state">{String(props.isSaving)}</p>
        <p data-testid="save-error">{props.saveError ?? ''}</p>
        <button type="button" onClick={() => props.onChange(nextBuilderItems)}>
          change items
        </button>
        <button type="button" onClick={props.onSave}>
          save menu
        </button>
        <button type="button" onClick={props.onDiscard}>
          discard menu
        </button>
      </section>
    )
  },
}))

function createMenuResponse(label = 'About Us'): MenuResponse {
  return {
    id: 1,
    menu_key: 'main',
    name: 'Main Website Navigation',
    items: [
      {
        id: 10,
        parent_id: null,
        label,
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

function createPageOptions(): MenuPageOption[] {
  return [
    {
      id: 100,
      page_title: 'About Us',
      url_slug: '/about-us',
      parent_id: null,
      parent_page_title: '',
      status: 'published',
    },
  ]
}

function createChangedItems(): MenuEditorItem[] {
  return [
    {
      client_id: 'changed-about',
      label: 'About Us Updated',
      navigation_type: 'pages',
      page_id: 100,
      external_url: '',
      open_in_new_tab: false,
      children: [],
    },
  ]
}

function latestMenuBuilderProps() {
  return menuBuilderPropsSpy.mock.calls[menuBuilderPropsSpy.mock.calls.length - 1]?.[0]
}

describe('MenusPage', () => {
  beforeEach(async () => {
    await i18n.changeLanguage('en')
    getMenuMock.mockReset()
    listMenuPageOptionsMock.mockReset()
    saveMenuMock.mockReset()
    toastSuccess.mockReset()
    toastError.mockReset()
    menuBuilderPropsSpy.mockReset()
    nextBuilderItems = []
    getApiErrorMessageMock.mockReset()
    getApiErrorMessageMock.mockImplementation((error: unknown) =>
      error instanceof Error ? error.message : 'Unknown error',
    )
  })

  it('loads the menu editor and passes the fetched data into MenuBuilder', async () => {
    getMenuMock.mockResolvedValue(createMenuResponse())
    listMenuPageOptionsMock.mockResolvedValue(createPageOptions())

    render(<MenusPage />)

    expect(screen.getByTestId('loader').textContent).toContain('Loading navigation...')
    await screen.findByTestId('menu-builder')

    expect(getMenuMock).toHaveBeenCalledWith('main')
    expect(listMenuPageOptionsMock).toHaveBeenCalledWith('main')
    expect(screen.getByTestId('app-shell').getAttribute('data-active-key')).toBe('menus')
    expect(screen.getByTestId('breadcrumb').textContent).toContain('Menus')
    expect(latestMenuBuilderProps().title).toBe('Main Website Navigation')
    expect(latestMenuBuilderProps().items).toHaveLength(1)
    expect(latestMenuBuilderProps().pageOptions).toEqual(createPageOptions())
    expect(latestMenuBuilderProps().isDirty).toBe(false)
  })

  it('shows a load error state with retry guidance', async () => {
    getMenuMock.mockRejectedValue(new Error('Load failed'))
    listMenuPageOptionsMock.mockResolvedValue(createPageOptions())

    render(<MenusPage />)

    await screen.findByRole('heading', { name: 'We could not load the menu editor' })
    expect(screen.getByText('Load failed')).toBeDefined()
    expect(screen.getByRole('button', { name: 'Retry' })).toBeDefined()
  })

  it('saves changed menu items and clears the dirty state after success', async () => {
    let resolveSave: ((value: unknown) => void) | null = null

    getMenuMock.mockResolvedValue(createMenuResponse())
    listMenuPageOptionsMock.mockResolvedValue(createPageOptions())
    saveMenuMock.mockImplementation(
      () =>
        new Promise((resolve) => {
          resolveSave = resolve
        }),
    )

    render(<MenusPage />)
    await screen.findByTestId('menu-builder')

    nextBuilderItems = createChangedItems()
    fireEvent.click(screen.getByRole('button', { name: 'change items' }))
    await waitFor(() => {
      expect(latestMenuBuilderProps().isDirty).toBe(true)
    })

    fireEvent.click(screen.getByRole('button', { name: 'save menu' }))

    await waitFor(() => {
      expect(saveMenuMock).toHaveBeenCalledWith('main', {
        name: 'Main Website Navigation',
        items: [
          {
            label: 'About Us Updated',
            navigation_type: 'pages',
            page_id: 100,
            external_url: '',
            open_in_new_tab: false,
            children: [],
          },
        ],
      })
    })
    await waitFor(() => {
      expect(latestMenuBuilderProps().isSaving).toBe(true)
    })

    resolveSave?.({
      message: 'Navigation saved successfully.',
      menu: createMenuResponse('About Us Updated'),
    })

    await waitFor(() => {
      expect(toastSuccess).toHaveBeenCalledWith('Navigation saved successfully.')
    })
    await waitFor(() => {
      expect(latestMenuBuilderProps().isDirty).toBe(false)
    })
  })

  it('shows save errors and restores the last saved state when discarding', async () => {
    getMenuMock.mockResolvedValue(createMenuResponse())
    listMenuPageOptionsMock.mockResolvedValue(createPageOptions())
    saveMenuMock.mockRejectedValue(new Error('Save failed'))

    render(<MenusPage />)
    await screen.findByTestId('menu-builder')

    nextBuilderItems = createChangedItems()
    fireEvent.click(screen.getByRole('button', { name: 'change items' }))
    await waitFor(() => {
      expect(latestMenuBuilderProps().isDirty).toBe(true)
    })

    fireEvent.click(screen.getByRole('button', { name: 'save menu' }))

    await waitFor(() => {
      expect(toastError).toHaveBeenCalledWith('Save failed')
    })
    expect(latestMenuBuilderProps().saveError).toBe('Save failed')

    fireEvent.click(screen.getByRole('button', { name: 'discard menu' }))

    await waitFor(() => {
      expect(latestMenuBuilderProps().isDirty).toBe(false)
    })
    expect(latestMenuBuilderProps().saveError).toBeNull()
  })
})
