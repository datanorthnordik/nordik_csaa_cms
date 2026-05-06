import { describe, it, expect } from 'vitest'
import { render, screen } from '@testing-library/react'
import './i18n'
import App from './App'

describe('App Component', () => {
  it('renders without crashing', () => {
    render(<App />)
    expect(screen.getByRole('heading', { name: 'Get started' })).toBeDefined()
    expect(screen.getByRole('button', { name: 'EN' })).toBeDefined()
  })
})
