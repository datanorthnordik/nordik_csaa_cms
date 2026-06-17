import { describe, expect, it } from 'vitest'
import {
  buildPendingRequestRows,
  filterPendingRequestRows,
  type BookRequestRow,
} from './BooksListPage'
import type { BookSubmission, BookSummary } from '../api/booksApi'

describe('BooksListPage request helpers', () => {
  const books: BookSummary[] = [
    {
      id: 1,
      title: 'Cook Book',
      description: 'Community recipes',
      activeVersionId: 11,
      activeVersionNumber: 2,
      versionCount: 2,
      pendingSubmissionCount: 1,
      createdAt: '2026-06-16T10:00:00Z',
      updatedAt: '2026-06-17T09:00:00Z',
    },
    {
      id: 2,
      title: 'Memory Book',
      description: 'Shared memories',
      activeVersionId: 21,
      activeVersionNumber: 4,
      versionCount: 4,
      pendingSubmissionCount: 1,
      createdAt: '2026-06-15T10:00:00Z',
      updatedAt: '2026-06-17T08:00:00Z',
    },
  ]

  const submissionsByBook = new Map<number, BookSubmission[]>([
    [
      1,
      [
        {
          id: 101,
          bookId: 1,
          bookVersionId: 11,
          bookVersionNumber: 2,
          targetSectionName: 'Desserts',
          targetSectionId: 8,
          newSectionName: '',
          status: 'pending',
          submitterEmail: 'chef@example.com',
          fieldValues: [],
          rejectionReason: '',
          createdAt: '2026-06-17T09:15:00Z',
          updatedAt: '2026-06-17T09:15:00Z',
        },
      ],
    ],
    [
      2,
      [
        {
          id: 88,
          bookId: 2,
          bookVersionId: 21,
          bookVersionNumber: 4,
          targetSectionName: '',
          newSectionName: 'New Elders',
          status: 'pending',
          submitterEmail: 'stories@example.com',
          fieldValues: [],
          rejectionReason: '',
          createdAt: '2026-06-16T18:30:00Z',
          updatedAt: '2026-06-16T18:30:00Z',
        },
      ],
    ],
  ])

  it('builds request rows with book context sorted by newest submission first', () => {
    const rows = buildPendingRequestRows(books, submissionsByBook)

    expect(rows).toEqual<BookRequestRow[]>([
      {
        submissionId: 101,
        bookId: 1,
        bookTitle: 'Cook Book',
        bookDescription: 'Community recipes',
        bookVersionId: 11,
        bookVersionNumber: 2,
        requestedSection: 'Desserts',
        submitterEmail: 'chef@example.com',
        createdAt: '2026-06-17T09:15:00Z',
      },
      {
        submissionId: 88,
        bookId: 2,
        bookTitle: 'Memory Book',
        bookDescription: 'Shared memories',
        bookVersionId: 21,
        bookVersionNumber: 4,
        requestedSection: 'New Elders',
        submitterEmail: 'stories@example.com',
        createdAt: '2026-06-16T18:30:00Z',
      },
    ])
  })

  it('filters request rows by book, version, section, request number, and email', () => {
    const rows = buildPendingRequestRows(books, submissionsByBook)

    expect(filterPendingRequestRows(rows, 'cook')).toHaveLength(1)
    expect(filterPendingRequestRows(rows, 'version 4')).toHaveLength(1)
    expect(filterPendingRequestRows(rows, 'new elders')).toHaveLength(1)
    expect(filterPendingRequestRows(rows, 'request 101')).toHaveLength(1)
    expect(filterPendingRequestRows(rows, 'stories@example.com')).toHaveLength(1)
  })
})
