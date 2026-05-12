import { describe, it, expect, vi } from 'vitest'
import { render, screen } from '@testing-library/react'
import userEvent from '@testing-library/user-event'
import { GalleryCard } from './GalleryCard'
import type { GallerySummary } from '../../types/media'

describe('GalleryCard', () => {
  const mockGallery: GallerySummary = {
    id: 1,
    name: 'Test Gallery',
    assetCount: 5,
    visibility: 'draft',
    updatedAt: '2024-01-01T00:00:00Z',
  }

  it('renders gallery card with name', () => {
    render(<GalleryCard gallery={mockGallery} />)

    expect(screen.getByText('Test Gallery')).toBeInTheDocument()
  })

  it('displays asset count', () => {
    render(<GalleryCard gallery={mockGallery} />)

    expect(screen.getByText(/5.*asset/i)).toBeInTheDocument()
  })

  it('displays visibility status', () => {
    render(<GalleryCard gallery={mockGallery} />)

    expect(screen.getByText(/draft/i)).toBeInTheDocument()
  })

  it('renders manage button', () => {
    render(<GalleryCard gallery={mockGallery} />)

    expect(screen.getByRole('button', { name: /manage/i })).toBeInTheDocument()
  })

  it('calls onManage when manage button is clicked', async () => {
    const handleManage = vi.fn()
    const user = await userEvent.setup()

    render(
      <GalleryCard
        gallery={mockGallery}
        onManage={handleManage}
      />
    )

    const manageButton = screen.getByRole('button', { name: /manage/i })
    await user.click(manageButton)

    expect(handleManage).toHaveBeenCalledWith(mockGallery)
  })

  it('displays front image when provided', () => {
    const galleryWithImage: GallerySummary = {
      ...mockGallery,
      frontImageUrl: 'https://example.com/image.jpg',
    }

    render(<GalleryCard gallery={galleryWithImage} />)

    const image = screen.getByRole('img')
    expect(image).toHaveAttribute('src', 'https://example.com/image.jpg')
  })

  it('displays placeholder when no front image', () => {
    render(<GalleryCard gallery={mockGallery} />)

    expect(screen.getByRole('img')).toBeInTheDocument()
  })

  it('formats relative time when formatter provided', () => {
    const formatRelativeTime = vi.fn().mockReturnValue('2 hours ago')

    render(
      <GalleryCard
        gallery={mockGallery}
        formatRelativeTime={formatRelativeTime}
      />
    )

    expect(screen.getByText(/2 hours ago/i)).toBeInTheDocument()
    expect(formatRelativeTime).toHaveBeenCalledWith(mockGallery.updatedAt)
  })

  it('renders published status when visibility is published', () => {
    const publishedGallery: GallerySummary = {
      ...mockGallery,
      visibility: 'published',
    }

    render(<GalleryCard gallery={publishedGallery} />)

    expect(screen.getByText(/published/i)).toBeInTheDocument()
  })
})
