import { describe, it, expect, vi } from 'vitest'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { BrowserRouter } from 'react-router-dom'
import { PressListPage } from './PressListPage'

describe('PressListPage', () => {
  it('renders press list page', () => {
    render(
      <BrowserRouter>
        <PressListPage loading={false} pressEntries={[]} />
      </BrowserRouter>
    )

    expect(screen.getByRole('main')).toBeInTheDocument()
  })

  it('displays loading state', () => {
    render(
      <BrowserRouter>
        <PressListPage loading={true} pressEntries={[]} />
      </BrowserRouter>
    )

    expect(screen.getByText(/loading|loading press/i)).toBeInTheDocument()
  })

  it('displays empty state when no press entries', () => {
    render(
      <BrowserRouter>
        <PressListPage loading={false} pressEntries={[]} />
      </BrowserRouter>
    )

    expect(screen.getByText(/no press|empty/i)).toBeInTheDocument()
  })

  it('displays press entries when data is available', () => {
    const mockEntries = [
      {
        id: 1,
        title: 'Press Release 1',
        status: 'published',
      },
      {
        id: 2,
        title: 'Press Release 2',
        status: 'draft',
      },
    ]

    render(
      <BrowserRouter>
        <PressListPage loading={false} pressEntries={mockEntries} />
      </BrowserRouter>
    )

    expect(screen.getByText('Press Release 1')).toBeInTheDocument()
    expect(screen.getByText('Press Release 2')).toBeInTheDocument()
  })

  it('displays error message when provided', () => {
    render(
      <BrowserRouter>
        <PressListPage
          loading={false}
          pressEntries={[]}
          error="Failed to load press releases"
        />
      </BrowserRouter>
    )

    expect(screen.getByText('Failed to load press releases')).toBeInTheDocument()
  })

  it('calls onCreateNew when create button is clicked', async () => {
    const handleCreate = vi.fn()
    const user = await userEvent.setup()

    render(
      <BrowserRouter>
        <PressListPage
          loading={false}
          pressEntries={[]}
          onCreateNew={handleCreate}
        />
      </BrowserRouter>
    )

    const createButton = screen.getByRole('button', { name: /create|new/i })
    await user.click(createButton)

    expect(handleCreate).toHaveBeenCalled()
  })
})
