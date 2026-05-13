import { render, screen, within } from '@testing-library/react'
import type { ReactNode } from 'react'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import i18n from '../i18n'
import type { PageListItem } from '../api/pagesApi'
import { PagesListPage } from './PagesListPage'

const { dispatchMock, navigateMock, useAppSelectorMock } = vi.hoisted(() => ({
  dispatchMock: vi.fn(),
  navigateMock: vi.fn(),
  useAppSelectorMock: vi.fn(),
}))

vi.mock('react-hot-toast', () => ({
  default: {
    success: vi.fn(),
    error: vi.fn(),
  },
}))

vi.mock('react-router-dom', async () => {
  const actual = await vi.importActual<typeof import('react-router-dom')>('react-router-dom')
  return {
    ...actual,
    useNavigate: () => navigateMock,
  }
})

vi.mock('../components/Breadcrumb', () => ({
  Breadcrumb: () => <nav>Breadcrumb</nav>,
}))

vi.mock('../components/CmsAppShell', () => ({
  CmsAppShell: ({ children }: { children?: ReactNode }) => <div>{children}</div>,
}))

vi.mock('../components/Loader', () => ({
  Loader: ({ label }: { label: string }) => <div>{label}</div>,
}))

vi.mock('../store/hooks', () => ({
  useAppDispatch: () => dispatchMock,
  useAppSelector: (selector: (state: unknown) => unknown) => useAppSelectorMock(selector),
}))

function createItem(
  id: number,
  pageTitle: string,
  pageType: PageListItem['page_type'],
): PageListItem {
  return {
    id,
    page_title: pageTitle,
    url_slug: `/${pageTitle.toLowerCase().replace(/\s+/g, '-')}`,
    page_type: pageType,
    parent_id: null,
    parent_page_title: '',
    parent_page_url_slug: '',
    status: 'published',
    last_modified: '2026-05-13T00:00:00Z',
    modified_by: null,
    modified_by_name: 'Admin',
    created_at: '2026-05-13T00:00:00Z',
    updated_at: '2026-05-13T00:00:00Z',
  }
}

describe('PagesListPage', () => {
  beforeEach(async () => {
    await i18n.changeLanguage('en')
    dispatchMock.mockReset()
    navigateMock.mockReset()
    useAppSelectorMock.mockImplementation((selector: (state: unknown) => unknown) =>
      selector({
        pages: {
          list: {
            items: [
              createItem(1, 'About Us', 'page'),
              createItem(2, 'Events', 'module'),
            ],
            pagination: null,
            appliedFilters: null,
            filters: {
              page: 1,
              pageSize: 10,
              searchTerm: '',
              status: '',
              sortBy: 'last_modified',
              sortOrder: 'desc',
            },
            status: 'succeeded',
            error: null,
          },
        },
      }),
    )
  })

  it('shows module pages in the list without edit or delete actions', () => {
    render(<PagesListPage />)

    const moduleRow = screen
      .getAllByText('Events')
      .map((element) => element.closest('tr'))
      .find((row) => row !== null)
    const standardRow = screen
      .getAllByText('About Us')
      .map((element) => element.closest('tr'))
      .find((row) => row !== null)

    expect(moduleRow).not.toBeNull()
    expect(standardRow).not.toBeNull()

    expect(within(moduleRow as HTMLElement).getByText('Module')).toBeDefined()
    expect(
      within(moduleRow as HTMLElement).getByRole('button', { name: 'View page' }),
    ).toBeDefined()
    expect(
      within(moduleRow as HTMLElement).queryByRole('button', { name: 'Edit page' }),
    ).toBeNull()
    expect(
      within(moduleRow as HTMLElement).queryByRole('button', { name: 'Delete page' }),
    ).toBeNull()

    expect(within(standardRow as HTMLElement).getByText('Page')).toBeDefined()
    expect(
      within(standardRow as HTMLElement).getByRole('button', { name: 'Edit page' }),
    ).toBeDefined()
    expect(
      within(standardRow as HTMLElement).getByRole('button', { name: 'Delete page' }),
    ).toBeDefined()
  })
})
