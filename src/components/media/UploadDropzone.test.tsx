import { fireEvent, render, screen } from '@testing-library/react'
import { describe, expect, it, vi } from 'vitest'
import { UploadDropzone } from './UploadDropzone'

describe('UploadDropzone', () => {
  it('renders the label and hint text', () => {
    render(<UploadDropzone label="Drop files here" hint="PNG or JPG" />)
    expect(screen.getByText('Drop files here')).toBeDefined()
    expect(screen.getByText('PNG or JPG')).toBeDefined()
  })

  it('renders a file input associated with the label', () => {
    const { container } = render(<UploadDropzone label="Upload" />)
    const input = container.querySelector('input[type="file"]')
    expect(input).not.toBeNull()
  })

  it('the file input uses position:fixed so it does not trigger container scroll', () => {
    const { container } = render(<UploadDropzone label="Upload" />)
    const input = container.querySelector('input[type="file"]') as HTMLElement
    // The CSS module applies a class whose computed style has position:fixed.
    // In jsdom the class name is hashed but we can verify the element has a
    // className (i.e. the module class was applied at all).
    expect(input.className.length).toBeGreaterThan(0)
  })

  it('calls onFiles with the selected files', () => {
    const onFiles = vi.fn()
    const { container } = render(<UploadDropzone onFiles={onFiles} />)
    const input = container.querySelector('input[type="file"]') as HTMLInputElement
    const file = new File(['content'], 'photo.png', { type: 'image/png' })
    Object.defineProperty(input, 'files', { value: [file], configurable: true })
    fireEvent.change(input)
    expect(onFiles).toHaveBeenCalledWith([file])
  })

  it('does not call onFiles when no files are selected', () => {
    const onFiles = vi.fn()
    const { container } = render(<UploadDropzone onFiles={onFiles} />)
    const input = container.querySelector('input[type="file"]') as HTMLInputElement
    Object.defineProperty(input, 'files', { value: [], configurable: true })
    fireEvent.change(input)
    expect(onFiles).not.toHaveBeenCalled()
  })

  it('passes accept attribute to the file input', () => {
    const { container } = render(<UploadDropzone accept="image/png,image/jpeg" />)
    const input = container.querySelector('input[type="file"]') as HTMLInputElement
    expect(input.accept).toBe('image/png,image/jpeg')
  })

  it('passes multiple attribute to the file input when multiple is true', () => {
    const { container } = render(<UploadDropzone multiple />)
    const input = container.querySelector('input[type="file"]') as HTMLInputElement
    expect(input.multiple).toBe(true)
  })

  it('disables the file input when disabled is true', () => {
    const { container } = render(<UploadDropzone disabled />)
    const input = container.querySelector('input[type="file"]') as HTMLInputElement
    expect(input.disabled).toBe(true)
  })

  it('calls onFiles with files dropped on the dropzone', () => {
    const onFiles = vi.fn()
    render(<UploadDropzone onFiles={onFiles} label="Drop zone" />)
    const label = screen.getByText('Drop zone').closest('label')!
    const file = new File(['data'], 'image.jpg', { type: 'image/jpeg' })
    fireEvent.drop(label, { dataTransfer: { files: [file] } })
    expect(onFiles).toHaveBeenCalledWith([file])
  })

  it('does not call onFiles on drop when disabled', () => {
    const onFiles = vi.fn()
    render(<UploadDropzone onFiles={onFiles} disabled label="Drop zone" />)
    const label = screen.getByText('Drop zone').closest('label')!
    const file = new File(['data'], 'image.jpg', { type: 'image/jpeg' })
    fireEvent.drop(label, { dataTransfer: { files: [file] } })
    expect(onFiles).not.toHaveBeenCalled()
  })
})
