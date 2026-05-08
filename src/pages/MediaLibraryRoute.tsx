import { useMockMediaStore } from '../lib/mockMediaStore'
import { MediaLibraryPage } from './MediaLibraryPage'

export function MediaLibraryRoute() {
  const { galleries, createGallery } = useMockMediaStore()

  return (
    <MediaLibraryPage
      galleries={galleries}
      onCreate={(values, frontImage) =>
        createGallery({
          name: values.name,
          description: values.description,
          visibility: values.visibility,
          frontImage,
        })
      }
    />
  )
}
