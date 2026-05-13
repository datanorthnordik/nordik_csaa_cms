import { BrowserRouter, Navigate, Route, Routes } from 'react-router-dom'
import { GuestRoute } from './components/GuestRoute'
import { ProtectedRoute } from './components/ProtectedRoute'
import { LoginPage } from './pages/LoginPage'
import { SignupPage } from './pages/SignupPage'
import { DashboardPage } from './pages/DashboardPage'
import { EventsListPage } from './pages/EventsListPage'
import { EventEditorPage } from './pages/EventEditorPage'
import { MediaLibraryRoute } from './pages/MediaLibraryRoute'
import { GalleryManagerRoute } from './pages/GalleryManagerRoute'
import { PagesListPage } from './pages/PagesListPage'
import { PageEditorPage } from './pages/PageEditorPage'
import { MenusPage } from './pages/MenusPage'
import './App.css'

function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route element={<GuestRoute />}>
          <Route path="/" element={<LoginPage />} />
          <Route path="/signup" element={<SignupPage />} />
        </Route>

        <Route element={<ProtectedRoute />}>
          <Route path="/dashboard" element={<DashboardPage />} />
          <Route path="/menus" element={<MenusPage />} />
          <Route path="/pages" element={<PagesListPage />} />
          <Route path="/pages/new" element={<PageEditorPage />} />
          <Route path="/pages/:id" element={<PageEditorPage mode="view" />} />
          <Route path="/pages/:id/edit" element={<PageEditorPage />} />
          <Route path="/events" element={<EventsListPage />} />
          <Route path="/events/new" element={<EventEditorPage />} />
          <Route path="/events/:id" element={<EventEditorPage mode="view" />} />
          <Route path="/events/:id/edit" element={<EventEditorPage />} />
          <Route path="/media-library" element={<MediaLibraryRoute />} />
          <Route path="/media-library/:galleryId" element={<GalleryManagerRoute />} />
        </Route>

        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </BrowserRouter>
  )
}

export default App
