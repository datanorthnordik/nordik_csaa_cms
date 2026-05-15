import type { PressCategory } from './pressTypes'

const FIXED_PRESS_CATEGORIES: PressCategory[] = [
  { id: '1', name: 'Statement', createdAt: '2024-01-01T00:00:00.000Z' },
  { id: '2', name: 'News', createdAt: '2024-01-01T00:00:00.000Z' },
  { id: '3', name: 'Announcement', createdAt: '2024-01-01T00:00:00.000Z' },
  {
    id: '4',
    name: 'Community Outreach',
    createdAt: '2024-01-01T00:00:00.000Z',
  },
]

export function loadPressCategories(): PressCategory[] {
  return FIXED_PRESS_CATEGORIES
}
