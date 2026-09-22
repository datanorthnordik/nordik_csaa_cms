import { describe, expect, it } from 'vitest'
import { cmsNavItems } from './cmsNav'

describe('cmsNav recordings entry', () => {
  it('links the recordings entry to the recordings library', () => {
    const entry = cmsNavItems.find((item) => item.key === 'recordings')

    expect(entry).toMatchObject({ key: 'recordings', path: '/recordings' })
    expect(entry?.icon).not.toBeNull()
  })
})
