export type MemorialStatus = 'published' | 'draft' | 'review'

export type MemorialCategory = 'alumnus' | 'veteran' | 'founder' | 'friend'

export type MemorialEntry = {
  id: number
  fullName: string
  affiliation: string
  category: string
  status: MemorialStatus
  dateAdded: string
  biography: string
  dateOfBirth?: string
  dateOfPassing?: string
  featuredPortraitUrl?: string
  galleryImageUrls?: string[]
}

export type MemorialFormState = {
  fullName: string
  biography: string
  category: string
  affiliation: string
  status: MemorialStatus
  publishAt: string
  dateOfBirth: string
  dateOfPassing: string
}

export type MemorialFormErrors = {
  fullName?: string
}

export type MemorialListFilters = {
  searchTerm: string
  status: MemorialStatus | 'all'
  category: string
}

export const defaultMemorialListFilters: MemorialListFilters = {
  searchTerm: '',
  status: 'all',
  category: '',
}

export const MEMORIAL_CATEGORIES: MemorialCategory[] = [
  'alumnus',
  'veteran',
  'founder',
  'friend',
]

export const MOCK_MEMORIAL_ENTRIES: MemorialEntry[] = [
  {
    id: 1,
    fullName: 'James Montgomery',
    affiliation: 'Class of 1964',
    category: 'alumnus',
    status: 'published',
    dateAdded: '2023-10-12',
    biography: '<p>James Montgomery was a proud alumnus of the class of 1964 and dedicated his life to the advancement of the association.</p>',
    dateOfBirth: '1946-03-14',
    dateOfPassing: '2023-09-01',
  },
  {
    id: 2,
    fullName: 'Sarah Whitaker',
    affiliation: 'Army Nurse Corps',
    category: 'veteran',
    status: 'published',
    dateAdded: '2023-11-04',
    biography: '<p>Sarah Whitaker served with distinction in the Army Nurse Corps and was deeply respected by all who knew her.</p>',
    dateOfBirth: '1950-07-22',
    dateOfPassing: '2023-10-15',
  },
  {
    id: 3,
    fullName: 'Robert Thompson',
    affiliation: 'CSAA Board 1980',
    category: 'founder',
    status: 'draft',
    dateAdded: '2024-01-15',
    biography: '<p>Robert Thompson was a founding member of the CSAA Board in 1980 and helped shape the organization for generations to come.</p>',
    dateOfBirth: '1938-11-05',
    dateOfPassing: '2024-01-02',
  },
  {
    id: 4,
    fullName: 'Eleanor Miller',
    affiliation: 'Class of 1972',
    category: 'alumnus',
    status: 'review',
    dateAdded: '2024-02-02',
    biography: '<p>Eleanor Miller was a beloved member of the class of 1972 and a tireless volunteer for community initiatives.</p>',
    dateOfBirth: '1954-04-18',
    dateOfPassing: '2024-01-20',
  },
  {
    id: 5,
    fullName: 'George Harrington',
    affiliation: 'Class of 1958',
    category: 'alumnus',
    status: 'published',
    dateAdded: '2024-03-10',
    biography: '<p>George Harrington served the association with honour for over four decades.</p>',
    dateOfBirth: '1940-02-28',
    dateOfPassing: '2024-02-14',
  },
  {
    id: 6,
    fullName: 'Patricia Levesque',
    affiliation: 'Faculty 1975–2001',
    category: 'friend',
    status: 'published',
    dateAdded: '2024-04-05',
    biography: '<p>Patricia Levesque dedicated 26 years as a faculty member and mentor to hundreds of students.</p>',
    dateOfBirth: '1948-09-11',
    dateOfPassing: '2024-03-22',
  },
]
