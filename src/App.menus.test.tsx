import { render, screen } from '@testing-library/react'
import { beforeEach, describe, expect, it, vi } from 'vitest'

vi.mock('./components/GuestRoute', async () => {
  const { Outlet } = await import('react-router-dom')
  return {
    GuestRoute: () => <Outlet />,
  }
})

vi.mock('./components/ProtectedRoute', async () => {
  const { Outlet } = await import('react-router-dom')
  return {
    ProtectedRoute: () => <Outlet />,
  }
})

vi.mock('./pages/LoginPage', () => ({
  LoginPage: () => <div>Login Page</div>,
}))

vi.mock('./pages/ForgotPasswordPage', () => ({
  ForgotPasswordPage: () => <div>Forgot Password Page</div>,
}))

vi.mock('./pages/ResetPasswordPage', () => ({
  ResetPasswordPage: () => <div>Reset Password Page</div>,
}))

vi.mock('./pages/SignupPage', () => ({
  SignupPage: () => <div>Signup Page</div>,
}))

vi.mock('./pages/DashboardPage', () => ({
  DashboardPage: () => <div>Dashboard Page</div>,
}))

vi.mock('./pages/EventsListPage', () => ({
  EventsListPage: () => <div>Events List Page</div>,
}))

vi.mock('./pages/EventEditorPage', () => ({
  EventEditorPage: () => <div>Event Editor Page</div>,
}))

vi.mock('./pages/MediaLibraryRoute', () => ({
  MediaLibraryRoute: () => <div>Media Library Route</div>,
}))

vi.mock('./pages/GalleryManagerRoute', () => ({
  GalleryManagerRoute: () => <div>Gallery Manager Route</div>,
}))

vi.mock('./pages/VideoLibraryRoute', () => ({
  VideoLibraryRoute: () => <div>Video Library Route</div>,
}))

vi.mock('./pages/VideoManagerRoute', () => ({
  VideoManagerRoute: () => <div>Video Manager Route</div>,
}))

vi.mock('./pages/PagesListPage', () => ({
  PagesListPage: () => <div>Pages List Page</div>,
}))

vi.mock('./pages/PageEditorPage', () => ({
  PageEditorPage: () => <div>Page Editor Page</div>,
}))

vi.mock('./pages/MenusPage', () => ({
  MenusPage: () => <h1>Menus Page</h1>,
}))

import App from './App'

describe('App menus route', () => {
  beforeEach(() => {
    window.history.pushState({}, '', '/menus')
  })

  it('renders the menus page when visiting /menus', () => {
    render(<App />)

    expect(screen.getByRole('heading', { name: 'Menus Page' })).toBeDefined()
  })
})
