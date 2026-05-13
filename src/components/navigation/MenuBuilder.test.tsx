import { fireEvent, render, screen, within } from '@testing-library/react'
import type { ComponentProps, ReactNode } from 'react'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../../i18n'
import type { MenuPageOption } from '../../api/menusApi'
import type { MenuEditorItem } from '../../lib/menusEditor'
import { MenuBuilder } from './MenuBuilder'

vi.mock('@mui/material/Dialog', () => ({
  default: ({
    open,
    children,
  }: {
    open: boolean
    children?: ReactNode
  }) => (open ? <div data-testid="dialog">{children}</div> : null),
}))

vi.mock('@mui/material/DialogActions', () => ({
  default: ({ children }: { children?: ReactNode }) => <div>{children}</div>,
}))

vi.mock('@mui/material/DialogContent', () => ({
  default: ({ children }: { children?: ReactNode }) => <div>{children}</div>,
}))

vi.mock('@mui/material/DialogContentText', () => ({
  default: ({ children }: { children?: ReactNode }) => <div>{children}</div>,
}))

vi.mock('@mui/material/DialogTitle', () => ({
  default: ({ children }: { children?: ReactNode }) => <div>{children}</div>,
}))

vi.mock('../icons', () => ({
  AddIcon: () => <span data-testid="add-icon" />,
  CheckCircleIcon: () => <span data-testid="check-icon" />,
  LinkIcon: () => <span data-testid="link-icon" />,
}))

const pageOptions: MenuPageOption[] = [
  {
    id: 10,
    page_title: 'Gatherings Overview',
    url_slug: '/gatherings',
    parent_id: null,
    parent_page_title: '',
    status: 'published',
  },
  {
    id: 11,
    page_title: 'Conference 2024',
    url_slug: '/gatherings/conference-2024',
    parent_id: 10,
    parent_page_title: 'Gatherings Overview',
    status: 'published',
  },
  {
    id: 12,
    page_title: 'About the Association',
    url_slug: '/about-us',
    parent_id: null,
    parent_page_title: '',
    status: 'published',
  },
  {
    id: 13,
    page_title: 'Regional Hub',
    url_slug: '/regional-hub',
    parent_id: null,
    parent_page_title: '',
    status: 'published',
  },
]

function createItems(): MenuEditorItem[] {
  return [
    {
      client_id: 'gatherings',
      label: 'Gatherings',
      navigation_type: 'pages',
      page_id: 10,
      external_url: '',
      open_in_new_tab: false,
      children: [
        {
          client_id: 'annual-conference',
          label: 'Annual Conference',
          navigation_type: 'pages',
          page_id: 11,
          external_url: '',
          open_in_new_tab: false,
          children: [],
        },
      ],
    },
    {
      client_id: 'about-us',
      label: 'About Us',
      navigation_type: 'pages',
      page_id: 12,
      external_url: '',
      open_in_new_tab: false,
      children: [],
    },
  ]
}

function renderMenuBuilder(overrides: Partial<ComponentProps<typeof MenuBuilder>> = {}) {
  const onChange = vi.fn()
  const onSave = vi.fn()
  const onDiscard = vi.fn()
  const props = {
    title: 'Main Website Navigation',
    subtitle: 'Configure the structure and hierarchy of the primary site menu.',
    items: createItems(),
    pageOptions,
    isDirty: true,
    isSaving: false,
    saveError: null,
    onChange,
    onSave,
    onDiscard,
    ...overrides,
  }

  return {
    ...render(<MenuBuilder {...props} />),
    props,
    onChange,
    onSave,
    onDiscard,
  }
}

function getItemRow(label: string) {
  return screen.getByText(label).closest('[draggable="true"]') as HTMLElement
}

describe('MenuBuilder', () => {
  beforeEach(async () => {
    await i18n.changeLanguage('en')
  })

  it('renders the empty state, save controls, and saving feedback', () => {
    const { onSave, props, rerender } = renderMenuBuilder({
      items: [],
      isDirty: false,
      saveError: 'Unable to save right now.',
    })

    expect(screen.getByText('Need a new section?')).toBeDefined()
    expect(screen.getByText('Unable to save right now.')).toBeDefined()
    expect(
      (screen.getByRole('button', { name: 'Discard Changes' }) as HTMLButtonElement).disabled,
    ).toBe(true)

    fireEvent.click(screen.getByRole('button', { name: 'Save Navigation' }))
    expect(onSave).toHaveBeenCalledTimes(1)

    fireEvent.click(screen.getByRole('button', { name: 'Add Menu Item' }))
    expect(
      screen.getByText('Create or update an item for your website navigation.'),
    ).toBeDefined()

    fireEvent.click(screen.getByRole('button', { name: 'Cancel' }))
    expect(
      screen.queryByText('Create or update an item for your website navigation.'),
    ).toBeNull()

    rerender(
      <MenuBuilder
        {...props}
        items={[]}
        isDirty={false}
        isSaving
        saveError={null}
      />,
    )

    expect(screen.getByRole('button', { name: 'Saving...' })).toBeDefined()
    expect(
      (screen.getByRole('button', { name: 'Saving...' }) as HTMLButtonElement).disabled,
    ).toBe(true)
  })

  it('validates the draft and adds a new page item', () => {
    const { onChange } = renderMenuBuilder({
      items: [],
    })

    fireEvent.click(screen.getByRole('button', { name: 'Add Menu Item' }))
    fireEvent.click(screen.getByRole('button', { name: 'Add to Menu' }))
    expect(screen.getByText('Menu label is required.')).toBeDefined()

    fireEvent.change(screen.getByPlaceholderText('e.g., Community Outreach'), {
      target: { value: 'About Us' },
    })
    fireEvent.click(screen.getByRole('button', { name: 'Add to Menu' }))
    expect(screen.getByText('Select a published page.')).toBeDefined()

    fireEvent.change(screen.getByRole('combobox'), {
      target: { value: '12' },
    })
    fireEvent.click(screen.getByRole('button', { name: 'Add to Menu' }))

    expect(onChange).toHaveBeenCalledTimes(1)
    expect(onChange.mock.calls[0]?.[0]).toHaveLength(1)
    expect(onChange.mock.calls[0]?.[0]?.[0]).toMatchObject({
      label: 'About Us',
      navigation_type: 'pages',
      page_id: 12,
    })
  })

  it('adds an external link beneath a parent page', () => {
    const { onChange } = renderMenuBuilder()

    fireEvent.click(screen.getByRole('button', { name: 'Add Menu Item' }))
    fireEvent.change(screen.getByPlaceholderText('e.g., Community Outreach'), {
      target: { value: 'Regional Hub' },
    })
    fireEvent.click(screen.getAllByRole('radio')[1] as HTMLInputElement)
    fireEvent.change(screen.getByPlaceholderText('https://example.com/resource'), {
      target: { value: 'https://example.com/resource' },
    })
    fireEvent.change(screen.getByRole('combobox'), {
      target: { value: '10' },
    })
    fireEvent.click(screen.getByRole('checkbox'))
    fireEvent.click(screen.getByRole('button', { name: 'Add to Menu' }))

    expect(onChange).toHaveBeenCalledTimes(1)
    expect(onChange.mock.calls[0]?.[0]?.[0]?.children).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          label: 'Regional Hub',
          navigation_type: 'external_link',
          page_id: null,
          external_url: 'https://example.com/resource',
          open_in_new_tab: true,
        }),
      ]),
    )
  })

  it('locks parent items with children to label-only edits', () => {
    const { onChange } = renderMenuBuilder()

    fireEvent.click(
      within(getItemRow('Gatherings')).getByRole('button', {
        name: 'Edit menu item',
      }),
    )

    expect(
      screen.getByText(
        'This item already has child links, so only the label can be changed here.',
      ),
    ).toBeDefined()
    expect((screen.getAllByRole('radio')[0] as HTMLInputElement).disabled).toBe(true)
    expect((screen.getByRole('combobox') as HTMLSelectElement).disabled).toBe(true)

    fireEvent.change(screen.getByDisplayValue('Gatherings'), {
      target: { value: 'Gatherings Updated' },
    })
    fireEvent.click(screen.getByRole('button', { name: 'Update Item' }))

    const updatedItem = onChange.mock.calls[0]?.[0]?.find(
      (item: MenuEditorItem) => item.label === 'Gatherings Updated',
    )

    expect(updatedItem).toBeDefined()
    expect(updatedItem?.children).toHaveLength(1)
  })

  it('moves and removes items at the correct level', () => {
    const { onChange } = renderMenuBuilder()

    fireEvent.click(
      within(getItemRow('Gatherings')).getByRole('button', {
        name: 'Move menu item down',
      }),
    )
    expect(onChange.mock.calls[0]?.[0]?.map((item: MenuEditorItem) => item.label)).toEqual([
      'About Us',
      'Gatherings',
    ])

    fireEvent.click(
      within(getItemRow('Gatherings')).getByRole('button', {
        name: 'Remove',
      }),
    )
    expect(
      screen.getByText('This will remove "Gatherings" and its 1 nested child items from the menu.'),
    ).toBeDefined()
    fireEvent.click(screen.getByRole('button', { name: 'Cancel' }))
    expect(
      screen.queryByText('This will remove "Gatherings" and its 1 nested child items from the menu.'),
    ).toBeNull()

    fireEvent.click(
      within(getItemRow('Annual Conference')).getByRole('button', {
        name: 'Remove',
      }),
    )
    expect(screen.getByText('This will remove "Annual Conference" from the menu.')).toBeDefined()

    const removeButtons = screen.getAllByRole('button', { name: 'Remove' })
    fireEvent.click(removeButtons[removeButtons.length - 1] as HTMLButtonElement)

    expect(onChange.mock.calls[1]?.[0]?.[0]?.children).toHaveLength(0)
  })

  it('reorders sibling items through drag and drop without crossing parents', () => {
    const { onChange } = renderMenuBuilder()
    const gatheringsRow = getItemRow('Gatherings')
    const aboutRow = getItemRow('About Us')
    const conferenceRow = getItemRow('Annual Conference')

    fireEvent.dragStart(gatheringsRow)
    fireEvent.dragOver(aboutRow)
    fireEvent.drop(aboutRow)

    expect(onChange).toHaveBeenCalledTimes(1)
    expect(onChange.mock.calls[0]?.[0]?.map((item: MenuEditorItem) => item.label)).toEqual([
      'About Us',
      'Gatherings',
    ])

    onChange.mockClear()

    fireEvent.dragStart(conferenceRow)
    fireEvent.dragOver(aboutRow)
    fireEvent.drop(aboutRow)
    fireEvent.dragEnd(conferenceRow)

    expect(onChange).not.toHaveBeenCalled()
  })
})
