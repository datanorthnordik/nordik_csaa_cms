import { BrowserRouter, Navigate, Route, Routes } from 'react-router-dom'
import { GuestRoute } from './components/GuestRoute'
import { ProtectedRoute } from './components/ProtectedRoute'
import { ForgotPasswordPage } from './pages/ForgotPasswordPage'
import { LoginPage } from './pages/LoginPage'
import { ResetPasswordPage } from './pages/ResetPasswordPage'
import { SignupPage } from './pages/SignupPage'
import { DashboardPage } from './pages/DashboardPage'
import { EventsListPage } from './pages/EventsListPage'
import { EventEditorPage } from './pages/EventEditorPage'
import { MediaLibraryRoute } from './pages/MediaLibraryRoute'
import { GalleryManagerRoute } from './pages/GalleryManagerRoute'
import { VideoLibraryRoute } from './pages/VideoLibraryRoute'
import { VideoManagerRoute } from './pages/VideoManagerRoute'
import { NewslettersListPage } from './pages/NewslettersListPage'
import { NewsletterEditorPage } from './pages/NewsletterEditorPage'
import { PressListPage } from './pages/PressListPage'
import { PressEditorPage } from './pages/PressEditorPage'
import { ResourcesListPage } from './pages/ResourcesListPage'
import { ResourceEditorPage } from './pages/ResourceEditorPage'
import { PagesListPage } from './pages/PagesListPage'
import { PageEditorPage } from './pages/PageEditorPage'
import { MenusPage } from './pages/MenusPage'
import { MemorialEntriesListPage } from './pages/MemorialEntriesListPage'
import { MemorialEntryEditorPage } from './pages/MemorialEntryEditorPage'
import './App.css'

function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route element={<GuestRoute />}>
          <Route path="/" element={<LoginPage />} />
          <Route path="/forgot-password" element={<ForgotPasswordPage />} />
          <Route path="/reset-password" element={<ResetPasswordPage />} />
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
          <Route path="/videos" element={<VideoLibraryRoute />} />
          <Route path="/videos/new" element={<VideoManagerRoute />} />
          <Route path="/videos/:videoId" element={<VideoManagerRoute />} />
          <Route path="/newsletters" element={<NewslettersListPage />} />
          <Route path="/newsletters/new" element={<NewsletterEditorPage mode="create" />} />
          <Route path="/newsletters/:id/edit" element={<NewsletterEditorPage mode="edit" />} />
          <Route path="/press" element={<PressListPage />} />
          <Route path="/press/new" element={<PressEditorPage mode="create" />} />
          <Route path="/press/:id/edit" element={<PressEditorPage mode="edit" />} />
          <Route path="/memorial" element={<MemorialEntriesListPage />} />
          <Route path="/memorial/new" element={<MemorialEntryEditorPage mode="create" />} />
          <Route path="/memorial/:id/edit" element={<MemorialEntryEditorPage mode="edit" />} />
          <Route path="/resources" element={<ResourcesListPage />} />
          <Route path="/resources/new" element={<ResourceEditorPage mode="create" />} />
          <Route path="/resources/:id/edit" element={<ResourceEditorPage mode="edit" />} />
        </Route>

        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </BrowserRouter>
  )
}

export default App
