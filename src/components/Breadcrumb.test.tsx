import { describe, it, expect } from 'vitest'
import { render, screen } from '@testing-library/react'
import { BrowserRouter } from 'react-router-dom'
import { Breadcrumb } from './Breadcrumb'

describe('Breadcrumb', () => {
  it('renders breadcrumb items', () => {
    const items = [
      { label: 'Home', path: '/' },
      { label: 'Dashboard', path: '/dashboard' },
    ]
    render(
      <BrowserRouter>
        <Breadcrumb items={items} />
      </BrowserRouter>
    )

    expect(screen.getByText('Home')).toBeInTheDocument()
    expect(screen.getByText('Dashboard')).toBeInTheDocument()
  })

  it('renders last item as current page without link', () => {
    const items = [
      { label: 'Home', path: '/' },
      { label: 'Current Page' },
    ]
    render(
      <BrowserRouter>
        <Breadcrumb items={items} />
      </BrowserRouter>
    )

    const currentPage = screen.getByText('Current Page')
    expect(currentPage).toBeInTheDocument()
    expect(currentPage.closest('a')).not.toBeInTheDocument()
  })

  it('renders navigation links for non-last items', () => {
    const items = [
      { label: 'Home', path: '/' },
      { label: 'Dashboard', path: '/dashboard' },
    ]
    render(
      <BrowserRouter>
        <Breadcrumb items={items} />
      </BrowserRouter>
    )

    const homeLink = screen.getByRole('link', { name: 'Home' })
    expect(homeLink).toHaveAttribute('href', '/')
  })

  it('renders empty when no items provided', () => {
    const { container } = render(
      <BrowserRouter>
        <Breadcrumb items={[]} />
      </BrowserRouter>
    )

    expect(container.querySelector('nav')).toBeInTheDocument()
  })
})
