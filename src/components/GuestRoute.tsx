import { Navigate, Outlet } from 'react-router-dom'
import { selectIsAuthenticated } from '../store/authSlice'
import { useAppSelector } from '../store/hooks'

export function GuestRoute() {
  const isAuthenticated = useAppSelector(selectIsAuthenticated)

  if (isAuthenticated) {
    return <Navigate to="/events" replace />
  }

  return <Outlet />
}
