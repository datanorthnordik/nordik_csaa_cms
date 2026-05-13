import { describe, expect, it } from 'vitest'
import type { MenuPageOption } from '../api/menusApi'
import {
  applyMenuItemDraft,
  buildMenuEditorItems,
  buildMenuItemDraft,
  buildSaveMenuItems,
  canMoveMenuItem,
  collectUsedPageIds,
  createDefaultMenuItemDraft,
  findParentPageIdForItem,
  getMenuItemSummary,
  listPageParentMenuOptions,
  moveMenuItem,
  removeMenuItem,
  reorderMenuItems,
  serializeMenuItems,
  validateMenuItemDraft,
  type MenuEditorItem,
} from './menusEditor'

const pageOptions: MenuPageOption[] = [
  {
    id: 10,
    page_title: 'About Us',
    url_slug: '/about',
    parent_id: null,
    parent_page_title: '',
    status: 'published',
  },
  {
    id: 11,
    page_title: 'Team',
    url_slug: '/about/team',
    parent_id: 10,
    parent_page_title: 'About Us',
    status: 'published',
  },
  {
    id: 12,
    page_title: 'Contact',
    url_slug: '/contact',
    parent_id: null,
    parent_page_title: '',
    status: 'published',
  },
]

const initialItems: MenuEditorItem[] = [
  {
    client_id: 'root-about',
    label: 'About',
    navigation_type: 'pages',
    page_id: 10,
    external_url: '',
    open_in_new_tab: false,
    children: [],
  },
  {
    client_id: 'root-contact',
    label: 'Contact',
    navigation_type: 'pages',
    page_id: 12,
    external_url: '',
    open_in_new_tab: false,
    children: [],
  },
]

describe('menusEditor', () => {
  it('adds a child page under its parent page automatically', () => {
    const draft = {
      ...createDefaultMenuItemDraft(),
      label: 'Team',
      navigation_type: 'pages' as const,
      page_id: '11',
    }

    const error = validateMenuItemDraft(draft, initialItems, pageOptions)
    expect(error).toBeNull()

    const nextItems = applyMenuItemDraft(initialItems, draft, pageOptions)
    expect(nextItems[0].children).toHaveLength(1)
    expect(nextItems[0].children[0]).toMatchObject({
      label: 'Team',
      page_id: 11,
      navigation_type: 'pages',
    })
  })

  it('prevents duplicate page usage in the menu', () => {
    const draft = {
      ...createDefaultMenuItemDraft(),
      label: 'About Duplicate',
      navigation_type: 'pages' as const,
      page_id: '10',
    }

    expect(validateMenuItemDraft(draft, initialItems, pageOptions)).toContain(
      'already been added',
    )
  })

  it('serializes nested items for save requests', () => {
    const items: MenuEditorItem[] = [
      {
        client_id: 'root-about',
        label: 'About',
        navigation_type: 'pages',
        page_id: 10,
        external_url: '',
        open_in_new_tab: false,
        children: [
          {
            client_id: 'child-link',
            label: 'Partner Site',
            navigation_type: 'external_link',
            page_id: null,
            external_url: 'https://example.com',
            open_in_new_tab: true,
            children: [],
          },
        ],
      },
    ]

    expect(buildSaveMenuItems(items)).toEqual([
      {
        label: 'About',
        navigation_type: 'pages',
        page_id: 10,
        external_url: '',
        open_in_new_tab: false,
        children: [
          {
            label: 'Partner Site',
            navigation_type: 'external_link',
            page_id: null,
            external_url: 'https://example.com',
            open_in_new_tab: true,
            children: [],
          },
        ],
      },
    ])
    expect(serializeMenuItems(items)).toContain('Partner Site')
  })

  it('reorders only within the same parent group', () => {
    const items: MenuEditorItem[] = [
      {
        client_id: 'root-about',
        label: 'About',
        navigation_type: 'pages',
        page_id: 10,
        external_url: '',
        open_in_new_tab: false,
        children: [
          {
            client_id: 'child-team',
            label: 'Team',
            navigation_type: 'pages',
            page_id: 11,
            external_url: '',
            open_in_new_tab: false,
            children: [],
          },
        ],
      },
      {
        client_id: 'root-contact',
        label: 'Contact',
        navigation_type: 'pages',
        page_id: 12,
        external_url: '',
        open_in_new_tab: false,
        children: [],
      },
    ]

    const unchanged = reorderMenuItems(items, 'child-team', 'root-contact')
    expect(unchanged[0].children[0].client_id).toBe('child-team')

    const reordered = reorderMenuItems(items, 'root-contact', 'root-about')
    expect(reordered[0].client_id).toBe('root-contact')
  })

  it('tracks used page ids and move availability', () => {
    const used = collectUsedPageIds(initialItems)
    expect(Array.from(used)).toEqual([10, 12])
    expect(canMoveMenuItem(initialItems, 'root-about', 'up')).toBe(false)
    expect(canMoveMenuItem(initialItems, 'root-about', 'down')).toBe(true)
  })

  it('builds editor drafts, parent menu options, and readable summaries', () => {
    const responseItems = [
      {
        id: 10,
        parent_id: null,
        label: 'About',
        navigation_type: 'pages' as const,
        page_id: 10,
        external_url: '',
        open_in_new_tab: false,
        sort_order: 1,
        href: '/about',
        children: [
          {
            id: 20,
            parent_id: 10,
            label: 'Partner Site',
            navigation_type: 'external_link' as const,
            page_id: null,
            external_url: 'https://example.com',
            open_in_new_tab: true,
            sort_order: 1,
            href: 'https://example.com',
            children: [],
          },
        ],
      },
    ]

    const editorItems = buildMenuEditorItems(responseItems)
    const externalChild = editorItems[0].children[0]

    expect(buildMenuItemDraft(externalChild, '10')).toEqual({
      label: 'Partner Site',
      navigation_type: 'external_link',
      page_id: '',
      external_url: 'https://example.com',
      parent_page_id: '10',
      open_in_new_tab: true,
    })
    expect(listPageParentMenuOptions(initialItems)).toEqual([
      { value: '10', label: 'About' },
      { value: '12', label: 'Contact' },
    ])
    expect(findParentPageIdForItem(editorItems, externalChild.client_id)).toBe('10')
    expect(getMenuItemSummary(externalChild, pageOptions)).toBe('https://example.com')
    expect(
      getMenuItemSummary(
        {
          client_id: 'missing-page',
          label: 'Missing Page',
          navigation_type: 'pages',
          page_id: 999,
          external_url: '',
          open_in_new_tab: false,
          children: [],
        },
        pageOptions,
      ),
    ).toBe('Linked page unavailable')
  })

  it('removes nested items, preserves no-op moves, and validates external parents', () => {
    const nestedItems: MenuEditorItem[] = [
      {
        client_id: 'root-about',
        label: 'About',
        navigation_type: 'pages',
        page_id: 10,
        external_url: '',
        open_in_new_tab: false,
        children: [
          {
            client_id: 'child-link',
            label: 'Partner Site',
            navigation_type: 'external_link',
            page_id: null,
            external_url: 'https://example.com',
            open_in_new_tab: true,
            children: [],
          },
        ],
      },
    ]

    expect(removeMenuItem(nestedItems, 'child-link')[0].children).toHaveLength(0)
    expect(moveMenuItem(initialItems, 'missing-item', 'down')).toBe(initialItems)
    expect(moveMenuItem(initialItems, 'root-about', 'up')).toBe(initialItems)

    expect(
      validateMenuItemDraft(
        {
          label: 'External Resource',
          navigation_type: 'external_link',
          page_id: '',
          external_url: 'https://example.com/resource',
          parent_page_id: '999',
          open_in_new_tab: false,
        },
        initialItems,
        pageOptions,
      ),
    ).toBe('Select a parent page that already exists in the menu.')
  })
})
