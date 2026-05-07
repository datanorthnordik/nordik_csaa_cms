import { Navigate, Outlet, useLocation } from 'react-router-dom'
import { selectIsAuthenticated } from '../store/authSlice'
import { useAppSelector } from '../store/hooks'

export function ProtectedRoute() {
  const isAuthenticated = useAppSelector(selectIsAuthenticated)
  const location = useLocation()

  if (!isAuthenticated) {
    return <Navigate to="/" replace state={{ from: location }} />
  }

  return <Outlet />
}
