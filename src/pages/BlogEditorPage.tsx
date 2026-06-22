import ExpandLessOutlinedIcon from '@mui/icons-material/ExpandLessOutlined'
import ExpandMoreOutlinedIcon from '@mui/icons-material/ExpandMoreOutlined'
import KeyboardArrowDownOutlinedIcon from '@mui/icons-material/KeyboardArrowDownOutlined'
import KeyboardArrowUpOutlinedIcon from '@mui/icons-material/KeyboardArrowUpOutlined'
import DragIndicatorOutlinedIcon from '@mui/icons-material/DragIndicatorOutlined'
import { useEffect, useMemo, useRef, useState, type DragEvent } from 'react'
import toast from 'react-hot-toast'
import { useNavigate, useParams } from 'react-router-dom'
import { blogsApi, resolveBlogAssetUrl, type BlogActionType, type BlogAnimationImagePosition, type BlogAnimationNavigation, type BlogDetailResponse, type BlogSectionResponse, type BlogSectionType, type SaveBlogRequest } from '../api/blogsApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { RichTextEditor } from '../components/cms/RichTextEditor'
import { AddIcon, CloudUploadIcon, DeleteIcon } from '../components/icons'
import { Loader } from '../components/Loader'
import { UploadDropzone } from '../components/media/UploadDropzone'
import styles from '../styles/PageEditorPage.module.css'

type BlogEditorPageProps = {
  mode?: 'edit' | 'view'
}

type AnimationItemState = {
  clientId: string
  id?: number
  heading: string
  subHeading: string
  description: string
  file: File | null
  existingImageUrl: string
  existingObjectKey: string
}

type BlogSectionState = {
  clientId: string
  id?: number
  sectionName: string
  sectionType: BlogSectionType
  isEnabled: boolean
  isCollapsed: boolean
  heading: {
    headingText: string
    underlineEnabled: boolean
  }
  image: {
    caption: string
    file: File | null
    existingImageUrl: string
    existingObjectKey: string
  }
  typography: {
    htmlContent: string
  }
  action: {
    text: string
    actionType: BlogActionType
    targetUrl: string
  }
  video: {
    youtubeUrl: string
    caption: string
  }
  animation: {
    navigation: BlogAnimationNavigation
    imagePosition: BlogAnimationImagePosition
    items: AnimationItemState[]
  }
}

type BlogFormState = {
  publishDate: string
  heading: string
  description: string
  coverImageFile: File | null
  existingCoverImageUrl: string
  existingCoverObjectKey: string
  sections: BlogSectionState[]
}

type FormErrors = Partial<Record<'publishDate' | 'heading' | 'description' | 'coverImage' | 'sections', string>>
type SectionDropPlacement = 'before' | 'after'

const IMAGE_ACCEPT = '.png,.jpg,.jpeg,.webp,image/png,image/jpeg,image/webp'

export function BlogEditorPage({ mode = 'edit' }: BlogEditorPageProps) {
  const navigate = useNavigate()
  const params = useParams()
  const parsedBlogId = params.id ? Number.parseInt(params.id, 10) : null
  const isEditMode = parsedBlogId !== null && Number.isFinite(parsedBlogId)
  const isViewMode = mode === 'view' && isEditMode
  const isReadOnlyMode = isViewMode

  const [form, setForm] = useState<BlogFormState>(createDefaultForm())
  const [errors, setErrors] = useState<FormErrors>({})
  const [loadingState, setLoadingState] = useState<'idle' | 'loading' | 'ready' | 'error'>(
    isEditMode ? 'loading' : 'ready',
  )
  const [loadError, setLoadError] = useState<string | null>(null)
  const [isSaving, setIsSaving] = useState(false)
  const [draggedSectionClientId, setDraggedSectionClientId] = useState<string | null>(null)
  const [dragOverSectionClientId, setDragOverSectionClientId] = useState<string | null>(null)
  const [dragOverSectionPlacement, setDragOverSectionPlacement] =
    useState<SectionDropPlacement | null>(null)
  const draggedSectionClientIdRef = useRef<string | null>(null)
  const dragOverSectionClientIdRef = useRef<string | null>(null)
  const dragOverSectionPlacementRef = useRef<SectionDropPlacement | null>(null)

  useEffect(() => {
    let cancelled = false

    if (!isEditMode || !parsedBlogId) {
      setForm(createDefaultForm())
      setLoadingState('ready')
      setLoadError(null)
      return () => {
        cancelled = true
      }
    }

    async function loadBlog() {
      setLoadingState('loading')
      setLoadError(null)

      try {
        const detail = await blogsApi.getBlog(parsedBlogId!)
        if (cancelled) {
          return
        }
        setForm(buildFormFromDetail(detail))
        setLoadingState('ready')
      } catch {
        if (cancelled) {
          return
        }
        setLoadingState('error')
        setLoadError('Could not load this blog right now.')
      }
    }

    void loadBlog()

    return () => {
      cancelled = true
    }
  }, [isEditMode, parsedBlogId])

  const moduleOptions = useMemo(
    () =>
      [
        { type: 'heading', label: 'Heading' },
        { type: 'image', label: 'Image' },
        { type: 'typography', label: 'Typography' },
        { type: 'action', label: 'Action' },
        { type: 'video', label: 'Video' },
        { type: 'animation', label: 'Animation' },
      ] satisfies Array<{ type: BlogSectionType; label: string }>,
    [],
  )

  const coverImagePreviewUrl = form.coverImageFile
    ? URL.createObjectURL(form.coverImageFile)
    : form.existingCoverImageUrl

  function updateSection(clientId: string, recipe: (section: BlogSectionState) => BlogSectionState) {
    setForm((current) => ({
      ...current,
      sections: current.sections.map((section) =>
        section.clientId === clientId ? recipe(section) : section,
      ),
    }))
  }

  function addSection(sectionType: BlogSectionType) {
    setForm((current) => ({
      ...current,
      sections: [...current.sections, createDefaultSection(sectionType)],
    }))
  }

  function removeSection(clientId: string) {
    setForm((current) => ({
      ...current,
      sections: current.sections.filter((section) => section.clientId !== clientId),
    }))
  }

  function moveSection(clientId: string, direction: -1 | 1) {
    setForm((current) => {
      const index = current.sections.findIndex((section) => section.clientId === clientId)
      if (index < 0) {
        return current
      }
      const nextIndex = index + direction
      if (nextIndex < 0 || nextIndex >= current.sections.length) {
        return current
      }
      const nextSections = current.sections.slice()
      const [moved] = nextSections.splice(index, 1)
      nextSections.splice(nextIndex, 0, moved)
      return { ...current, sections: nextSections }
    })
  }

  function handleSectionDragStart(clientId: string) {
    if (isReadOnlyMode || isSaving) {
      return
    }
    draggedSectionClientIdRef.current = clientId
    dragOverSectionClientIdRef.current = clientId
    dragOverSectionPlacementRef.current = 'before'
    setDraggedSectionClientId(clientId)
    setDragOverSectionClientId(clientId)
    setDragOverSectionPlacement('before')
  }

  function resolveSectionDropPlacement(event: DragEvent<HTMLElement>): SectionDropPlacement {
    const bounds = event.currentTarget.getBoundingClientRect()
    return event.clientY - bounds.top <= bounds.height / 2 ? 'before' : 'after'
  }

  function handleSectionDragOver(event: DragEvent<HTMLElement>, clientId: string) {
    const draggedId = draggedSectionClientIdRef.current
    if (isReadOnlyMode || isSaving || !draggedId || draggedId === clientId) {
      return
    }
    event.preventDefault()
    event.dataTransfer.dropEffect = 'move'
    const placement = resolveSectionDropPlacement(event)
    dragOverSectionClientIdRef.current = clientId
    dragOverSectionPlacementRef.current = placement
    setDragOverSectionClientId(clientId)
    setDragOverSectionPlacement(placement)
  }

  function handleSectionDrop(event: DragEvent<HTMLElement>, targetId: string) {
    const draggedId = draggedSectionClientIdRef.current
    if (isReadOnlyMode || isSaving || !draggedId || draggedId === targetId) {
      return
    }
    event.preventDefault()
    const placement =
      dragOverSectionClientIdRef.current === targetId && dragOverSectionPlacementRef.current
        ? dragOverSectionPlacementRef.current
        : resolveSectionDropPlacement(event)
    setForm((current) => ({
      ...current,
      sections: reorderSections(current.sections, draggedId, targetId, placement),
    }))
    handleSectionDragEnd()
  }

  function handleSectionDragEnd() {
    draggedSectionClientIdRef.current = null
    dragOverSectionClientIdRef.current = null
    dragOverSectionPlacementRef.current = null
    setDraggedSectionClientId(null)
    setDragOverSectionClientId(null)
    setDragOverSectionPlacement(null)
  }

  function toggleSectionCollapse(clientId: string) {
    updateSection(clientId, (section) => ({
      ...section,
      isCollapsed: !section.isCollapsed,
    }))
  }

  function addAnimationItem(sectionClientId: string) {
    updateSection(sectionClientId, (section) => ({
      ...section,
      animation: {
        ...section.animation,
        items: [...section.animation.items, createDefaultAnimationItem()],
      },
    }))
  }

  function updateAnimationItem(
    sectionClientId: string,
    itemClientId: string,
    recipe: (item: AnimationItemState) => AnimationItemState,
  ) {
    updateSection(sectionClientId, (section) => ({
      ...section,
      animation: {
        ...section.animation,
        items: section.animation.items.map((item) =>
          item.clientId === itemClientId ? recipe(item) : item,
        ),
      },
    }))
  }

  function removeAnimationItem(sectionClientId: string, itemClientId: string) {
    updateSection(sectionClientId, (section) => ({
      ...section,
      animation: {
        ...section.animation,
        items: section.animation.items.filter((item) => item.clientId !== itemClientId),
      },
    }))
  }

  function moveAnimationItem(sectionClientId: string, itemClientId: string, direction: -1 | 1) {
    updateSection(sectionClientId, (section) => {
      const index = section.animation.items.findIndex((item) => item.clientId === itemClientId)
      if (index < 0) {
        return section
      }
      const nextIndex = index + direction
      if (nextIndex < 0 || nextIndex >= section.animation.items.length) {
        return section
      }
      const nextItems = section.animation.items.slice()
      const [moved] = nextItems.splice(index, 1)
      nextItems.splice(nextIndex, 0, moved)
      return {
        ...section,
        animation: {
          ...section.animation,
          items: nextItems,
        },
      }
    })
  }

  function validateForm() {
    const nextErrors: FormErrors = {}

    if (!form.publishDate.trim()) {
      nextErrors.publishDate = 'Publish date is required.'
    }
    if (!form.heading.trim()) {
      nextErrors.heading = 'Heading is required.'
    }
    if (!form.description.trim()) {
      nextErrors.description = 'Description is required.'
    }
    if (!form.coverImageFile && !form.existingCoverImageUrl && !form.existingCoverObjectKey) {
      nextErrors.coverImage = 'Cover image is required.'
    }

    const hasInvalidSection = form.sections.some((section) => {
      switch (section.sectionType) {
        case 'heading':
          return !section.heading.headingText.trim()
        case 'image':
          return !section.image.file && !section.image.existingImageUrl && !section.image.existingObjectKey
        case 'typography':
          return stripHtml(section.typography.htmlContent) === ''
        case 'action':
          return !section.action.text.trim() || !section.action.targetUrl.trim()
        case 'video':
          return !section.video.youtubeUrl.trim()
        case 'animation':
          return (
            section.animation.items.length === 0 ||
            section.animation.items.some(
              (item) =>
                !item.heading.trim() ||
                !item.subHeading.trim() ||
                !item.description.trim() ||
                (!item.file && !item.existingImageUrl && !item.existingObjectKey),
            )
          )
        default:
          return false
      }
    })

    if (hasInvalidSection) {
      nextErrors.sections = 'Fill in every required module field before saving.'
    }

    setErrors(nextErrors)
    return Object.keys(nextErrors).length === 0
  }

  async function handleSave() {
    if (isReadOnlyMode || isSaving) {
      return
    }
    if (!validateForm()) {
      return
    }

    setIsSaving(true)
    try {
      const request = buildSaveRequest(form)
      if (isEditMode && parsedBlogId) {
        const response = await blogsApi.updateBlog(parsedBlogId, request)
        toast.success('Blog updated successfully.')
        navigate(`/blogs/${response.blog.id}`, { replace: true })
      } else {
        const response = await blogsApi.createBlog(request)
        toast.success('Blog created successfully.')
        navigate(`/blogs/${response.blog.id}/edit`, { replace: true })
      }
    } catch {
      toast.error('Could not save this blog right now.')
    } finally {
      setIsSaving(false)
    }
  }

  if (params.id !== undefined && !isEditMode) {
    return (
      <CmsAppShell activeKey="blogs">
        <div className={styles.page}>
          <div className={styles.errorPanel}>
            <h1>Invalid blog ID</h1>
            <p>The blog record you requested could not be found.</p>
            <button type="button" className={styles.secondaryButton} onClick={() => navigate('/blogs')}>
              Back to blogs
            </button>
          </div>
        </div>
      </CmsAppShell>
    )
  }

  if (loadingState === 'loading') {
    return (
      <CmsAppShell activeKey="blogs">
        <div className={styles.page}>
          <div className={styles.loaderWrap}>
            <Loader label="Loading blog" />
          </div>
        </div>
      </CmsAppShell>
    )
  }

  if (loadingState === 'error') {
    return (
      <CmsAppShell activeKey="blogs">
        <div className={styles.page}>
          <div className={styles.errorPanel}>
            <h1>Could not load this blog</h1>
            <p>{loadError}</p>
            <button type="button" className={styles.secondaryButton} onClick={() => navigate('/blogs')}>
              Back to blogs
            </button>
          </div>
        </div>
      </CmsAppShell>
    )
  }

  return (
    <CmsAppShell activeKey="blogs">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: 'Blogs', to: '/blogs' },
            { label: isViewMode ? 'View' : isEditMode ? 'Edit' : 'Create' },
          ]}
        />

        <header className={styles.header}>
          <div>
            <h1 className={styles.title}>
              {isViewMode ? 'View blog' : isEditMode ? 'Edit blog' : 'Create blog'}
            </h1>
            <p className={styles.subtitle}>
              Manage the story metadata and arrange blog modules in the same order they
              should appear on the website.
            </p>
          </div>
          <div className={styles.headerActions}>
            <button type="button" className={styles.secondaryButton} onClick={() => navigate('/blogs')}>
              Back to blogs
            </button>
            {isViewMode ? (
              <button
                type="button"
                className={styles.primaryButton}
                onClick={() => parsedBlogId && navigate(`/blogs/${parsedBlogId}/edit`)}
              >
                Edit blog
              </button>
            ) : (
              <button type="button" className={styles.primaryButton} disabled={isSaving} onClick={() => void handleSave()}>
                {isSaving ? 'Saving...' : isEditMode ? 'Save changes' : 'Create blog'}
              </button>
            )}
          </div>
        </header>

        {Object.keys(errors).length > 0 && (
          <div className={styles.errorSummary}>
            <p className={styles.errorSummaryTitle}>Please fix the following before saving:</p>
            <ul>
              {Object.values(errors).map((message) =>
                message ? <li key={message}>{message}</li> : null,
              )}
            </ul>
          </div>
        )}

        <fieldset className={styles.readOnlyFieldset} disabled={isReadOnlyMode || isSaving}>
          <section className={styles.card}>
            <div className={styles.cardHeaderBlock}>
              <h2>Blog settings</h2>
              <p>These fields drive the story card and the top of the blog detail view.</p>
            </div>

            <div className={styles.fieldGrid}>
              <label className={styles.field}>
                <span>Publish date</span>
                <input
                  type="date"
                  value={form.publishDate}
                  onChange={(event) => setForm((current) => ({ ...current, publishDate: event.target.value }))}
                />
                {errors.publishDate && <strong className={styles.fieldError}>{errors.publishDate}</strong>}
              </label>

              <label className={styles.field}>
                <span>Heading</span>
                <input
                  type="text"
                  value={form.heading}
                  onChange={(event) => setForm((current) => ({ ...current, heading: event.target.value }))}
                  placeholder="Shirley: A Residential School Story"
                />
                {errors.heading && <strong className={styles.fieldError}>{errors.heading}</strong>}
              </label>
            </div>

            <label className={styles.field}>
              <span>Description</span>
              <textarea
                rows={4}
                value={form.description}
                onChange={(event) => setForm((current) => ({ ...current, description: event.target.value }))}
                placeholder="This description appears on the story card and supports the detail page."
              />
              {errors.description && <strong className={styles.fieldError}>{errors.description}</strong>}
            </label>

            <div className={styles.heroSection}>
              <div className={styles.cardHeaderBlock}>
                <h2>Cover image</h2>
                <p>Used on the story card and as the lead visual when the story opens.</p>
              </div>

              <div className={styles.heroPanel}>
                {coverImagePreviewUrl ? (
                  <div className={styles.heroPreviewCard}>
                    <img src={coverImagePreviewUrl} alt={form.heading || 'Blog cover'} className={styles.heroPreview} />
                    {!isReadOnlyMode && (
                      <div className={styles.heroActions}>
                        <button
                          type="button"
                          className={styles.secondaryButton}
                          onClick={() =>
                            setForm((current) => ({
                              ...current,
                              coverImageFile: null,
                              existingCoverImageUrl: '',
                              existingCoverObjectKey: '',
                            }))
                          }
                        >
                          Remove cover image
                        </button>
                      </div>
                    )}
                  </div>
                ) : (
                  <UploadDropzone
                    label="Upload cover image"
                    hint="PNG, JPG, or WebP."
                    accept={IMAGE_ACCEPT}
                    icon={<CloudUploadIcon size={26} />}
                    onFiles={(files) =>
                      files[0] &&
                      setForm((current) => ({
                        ...current,
                        coverImageFile: files[0],
                        existingCoverImageUrl: '',
                        existingCoverObjectKey: '',
                      }))
                    }
                    disabled={isReadOnlyMode}
                  />
                )}
                {errors.coverImage && <strong className={styles.fieldError}>{errors.coverImage}</strong>}
              </div>
            </div>
          </section>

          <section className={styles.card}>
            <div className={styles.cardHeaderBlock}>
              <h2>Content modules</h2>
              <p>Add the sections in the same order they should appear on the website.</p>
            </div>

            {form.sections.length ? (
              <div className={styles.moduleList}>
                {form.sections.map((section, index) => {
                  const isDragOverTarget =
                    dragOverSectionClientId === section.clientId &&
                    draggedSectionClientId !== section.clientId

                  return (
                  <article
                    key={section.clientId}
                    className={[
                      styles.moduleCard,
                      draggedSectionClientId === section.clientId
                        ? styles.moduleCardDragging
                        : '',
                      isDragOverTarget ? styles.moduleCardDragOver : '',
                      isDragOverTarget && dragOverSectionPlacement === 'before'
                        ? styles.moduleCardDragOverBefore
                        : '',
                      isDragOverTarget && dragOverSectionPlacement === 'after'
                        ? styles.moduleCardDragOverAfter
                        : '',
                    ].filter(Boolean).join(' ')}
                    draggable={!isReadOnlyMode && !isSaving}
                    onDragStart={() => handleSectionDragStart(section.clientId)}
                    onDragOver={(event) => handleSectionDragOver(event, section.clientId)}
                    onDrop={(event) => handleSectionDrop(event, section.clientId)}
                    onDragEnd={handleSectionDragEnd}
                  >
                    <div className={styles.moduleHeader}>
                      <div className={styles.moduleHeaderMain}>
                        {!isReadOnlyMode && (
                          <button
                            type="button"
                            className={styles.dragHandle}
                            aria-label="Drag to reorder module"
                          >
                            <DragIndicatorOutlinedIcon fontSize="small" />
                          </button>
                        )}
                        <div className={styles.moduleHeadingBlock}>
                          <div className={styles.moduleEyebrowRow}>
                            <span className={styles.moduleEyebrow}>{section.sectionType}</span>
                            {!section.isEnabled && <span className={styles.moduleBadge}>Hidden</span>}
                          </div>
                          <p className={styles.moduleTitle}>{section.sectionName || 'Untitled module'}</p>
                        </div>
                      </div>

                      <div className={styles.moduleHeaderActions}>
                        <label className={styles.moduleToggle}>
                          <span>Enabled</span>
                          <input
                            type="checkbox"
                            checked={section.isEnabled}
                            onChange={(event) =>
                              updateSection(section.clientId, (current) => ({
                                ...current,
                                isEnabled: event.target.checked,
                              }))
                            }
                          />
                        </label>
                        <button
                          type="button"
                          className={styles.iconButton}
                          aria-label="Move module up"
                          disabled={index === 0}
                          onClick={() => moveSection(section.clientId, -1)}
                        >
                          <KeyboardArrowUpOutlinedIcon fontSize="small" />
                        </button>
                        <button
                          type="button"
                          className={styles.iconButton}
                          aria-label="Move module down"
                          disabled={index === form.sections.length - 1}
                          onClick={() => moveSection(section.clientId, 1)}
                        >
                          <KeyboardArrowDownOutlinedIcon fontSize="small" />
                        </button>
                        <button
                          type="button"
                          className={styles.iconButton}
                          aria-label={section.isCollapsed ? 'Expand module' : 'Collapse module'}
                          onClick={() => toggleSectionCollapse(section.clientId)}
                        >
                          {section.isCollapsed ? (
                            <ExpandMoreOutlinedIcon fontSize="small" />
                          ) : (
                            <ExpandLessOutlinedIcon fontSize="small" />
                          )}
                        </button>
                        {!isReadOnlyMode && (
                          <button
                            type="button"
                            className={styles.iconButtonDanger}
                            aria-label="Remove module"
                            onClick={() => removeSection(section.clientId)}
                          >
                            <DeleteIcon size={16} />
                          </button>
                        )}
                      </div>
                    </div>

                    {!section.isCollapsed && (
                      <div className={styles.moduleBody}>
                        <div className={styles.fieldStack}>
                          <label className={styles.field}>
                            <span>Module label</span>
                            <input
                              type="text"
                              value={section.sectionName}
                              onChange={(event) =>
                                updateSection(section.clientId, (current) => ({
                                  ...current,
                                  sectionName: event.target.value,
                                }))
                              }
                            />
                          </label>

                          {section.sectionType === 'heading' && (
                            <>
                              <label className={styles.field}>
                                <span>Heading text</span>
                                <input
                                  type="text"
                                  value={section.heading.headingText}
                                  onChange={(event) =>
                                    updateSection(section.clientId, (current) => ({
                                      ...current,
                                      heading: {
                                        ...current.heading,
                                        headingText: event.target.value,
                                      },
                                    }))
                                  }
                                />
                              </label>
                              <label className={styles.moduleToggle}>
                                <span>Underline</span>
                                <input
                                  type="checkbox"
                                  checked={section.heading.underlineEnabled}
                                  onChange={(event) =>
                                    updateSection(section.clientId, (current) => ({
                                      ...current,
                                      heading: {
                                        ...current.heading,
                                        underlineEnabled: event.target.checked,
                                      },
                                    }))
                                  }
                                />
                              </label>
                            </>
                          )}

                          {section.sectionType === 'image' && (
                            <>
                              <label className={styles.field}>
                                <span>Caption</span>
                                <textarea
                                  rows={3}
                                  value={section.image.caption}
                                  onChange={(event) =>
                                    updateSection(section.clientId, (current) => ({
                                      ...current,
                                      image: {
                                        ...current.image,
                                        caption: event.target.value,
                                      },
                                    }))
                                  }
                                />
                              </label>
                              {section.image.file || section.image.existingImageUrl ? (
                                <div className={styles.heroPreviewCard}>
                                  <img
                                    src={
                                      section.image.file
                                        ? URL.createObjectURL(section.image.file)
                                        : section.image.existingImageUrl
                                    }
                                    alt={section.sectionName || 'Section image'}
                                    className={styles.heroPreview}
                                  />
                                  {!isReadOnlyMode && (
                                    <div className={styles.heroActions}>
                                      <button
                                        type="button"
                                        className={styles.secondaryButton}
                                        onClick={() =>
                                          updateSection(section.clientId, (current) => ({
                                            ...current,
                                            image: {
                                              ...current.image,
                                              file: null,
                                              existingImageUrl: '',
                                              existingObjectKey: '',
                                            },
                                          }))
                                        }
                                      >
                                        Remove image
                                      </button>
                                    </div>
                                  )}
                                </div>
                              ) : (
                                <UploadDropzone
                                  label="Upload module image"
                                  hint="PNG, JPG, or WebP."
                                  accept={IMAGE_ACCEPT}
                                  icon={<CloudUploadIcon size={26} />}
                                  onFiles={(files) =>
                                    files[0] &&
                                    updateSection(section.clientId, (current) => ({
                                      ...current,
                                      image: {
                                        ...current.image,
                                        file: files[0],
                                        existingImageUrl: '',
                                        existingObjectKey: '',
                                      },
                                    }))
                                  }
                                  disabled={isReadOnlyMode}
                                />
                              )}
                            </>
                          )}

                          {section.sectionType === 'typography' && (
                            <RichTextEditor
                              value={section.typography.htmlContent}
                              onChange={(html) =>
                                updateSection(section.clientId, (current) => ({
                                  ...current,
                                  typography: {
                                    ...current.typography,
                                    htmlContent: html,
                                  },
                                }))
                              }
                              placeholder="Write the story content here..."
                              disabled={isReadOnlyMode}
                            />
                          )}

                          {section.sectionType === 'action' && (
                            <>
                              <label className={styles.field}>
                                <span>Button text</span>
                                <input
                                  type="text"
                                  value={section.action.text}
                                  onChange={(event) =>
                                    updateSection(section.clientId, (current) => ({
                                      ...current,
                                      action: {
                                        ...current.action,
                                        text: event.target.value,
                                      },
                                    }))
                                  }
                                />
                              </label>
                              <label className={styles.field}>
                                <span>Action type</span>
                                <select
                                  value={section.action.actionType}
                                  onChange={(event) =>
                                    updateSection(section.clientId, (current) => ({
                                      ...current,
                                      action: {
                                        ...current.action,
                                        actionType: event.target.value as BlogActionType,
                                      },
                                    }))
                                  }
                                >
                                  <option value="link">Open a link</option>
                                  <option value="video">Play a video</option>
                                </select>
                              </label>
                              <label className={styles.field}>
                                <span>{section.action.actionType === 'video' ? 'Video URL' : 'Link URL'}</span>
                                <input
                                  type="text"
                                  value={section.action.targetUrl}
                                  onChange={(event) =>
                                    updateSection(section.clientId, (current) => ({
                                      ...current,
                                      action: {
                                        ...current.action,
                                        targetUrl: event.target.value,
                                      },
                                    }))
                                  }
                                />
                              </label>
                            </>
                          )}

                          {section.sectionType === 'video' && (
                            <>
                              <label className={styles.field}>
                                <span>YouTube URL</span>
                                <input
                                  type="text"
                                  value={section.video.youtubeUrl}
                                  onChange={(event) =>
                                    updateSection(section.clientId, (current) => ({
                                      ...current,
                                      video: {
                                        ...current.video,
                                        youtubeUrl: event.target.value,
                                      },
                                    }))
                                  }
                                />
                              </label>
                              <label className={styles.field}>
                                <span>Caption</span>
                                <textarea
                                  rows={3}
                                  value={section.video.caption}
                                  onChange={(event) =>
                                    updateSection(section.clientId, (current) => ({
                                      ...current,
                                      video: {
                                        ...current.video,
                                        caption: event.target.value,
                                      },
                                    }))
                                  }
                                />
                              </label>
                            </>
                          )}

                          {section.sectionType === 'animation' && (
                            <>
                              <div className={styles.fieldGrid}>
                                <label className={styles.field}>
                                  <span>Navigation</span>
                                  <select
                                    value={section.animation.navigation}
                                    onChange={(event) =>
                                      updateSection(section.clientId, (current) => ({
                                        ...current,
                                        animation: {
                                          ...current.animation,
                                          navigation: event.target.value as BlogAnimationNavigation,
                                        },
                                      }))
                                    }
                                  >
                                    <option value="vertical">Vertical</option>
                                    <option value="horizontal">Horizontal</option>
                                  </select>
                                </label>
                                <label className={styles.field}>
                                  <span>Image position</span>
                                  <select
                                    value={section.animation.imagePosition}
                                    onChange={(event) =>
                                      updateSection(section.clientId, (current) => ({
                                        ...current,
                                        animation: {
                                          ...current.animation,
                                          imagePosition: event.target.value as BlogAnimationImagePosition,
                                        },
                                      }))
                                    }
                                  >
                                    <option value="left">Left</option>
                                    <option value="right">Right</option>
                                  </select>
                                </label>
                              </div>

                              <div className={styles.fieldStack}>
                                {section.animation.items.map((item, itemIndex) => (
                                  <div
                                    key={item.clientId}
                                    className={[styles.optionCard, styles.animationItemCard].join(' ')}
                                  >
                                    <div
                                      className={[styles.optionCardContent, styles.animationItemContent].join(' ')}
                                    >
                                      <div className={styles.optionCardHeader}>
                                        <p className={styles.optionCardTitle}>Animation item {itemIndex + 1}</p>
                                      </div>

                                      <div className={styles.animationItemFields}>
                                        <label className={styles.field}>
                                          <span>Heading</span>
                                          <input
                                            type="text"
                                            value={item.heading}
                                            onChange={(event) =>
                                              updateAnimationItem(section.clientId, item.clientId, (current) => ({
                                                ...current,
                                                heading: event.target.value,
                                              }))
                                            }
                                          />
                                        </label>

                                        <label className={styles.field}>
                                          <span>Sub heading</span>
                                          <input
                                            type="text"
                                            value={item.subHeading}
                                            onChange={(event) =>
                                              updateAnimationItem(section.clientId, item.clientId, (current) => ({
                                                ...current,
                                                subHeading: event.target.value,
                                              }))
                                            }
                                          />
                                        </label>

                                        <label
                                          className={[styles.field, styles.animationItemDescriptionField].join(' ')}
                                        >
                                          <span>Description</span>
                                          <textarea
                                            rows={3}
                                            value={item.description}
                                            onChange={(event) =>
                                              updateAnimationItem(section.clientId, item.clientId, (current) => ({
                                                ...current,
                                                description: event.target.value,
                                              }))
                                            }
                                          />
                                        </label>
                                      </div>

                                      {item.file || item.existingImageUrl ? (
                                        <div
                                          className={[styles.heroPreviewCard, styles.animationItemMedia].join(' ')}
                                        >
                                          <img
                                            src={item.file ? URL.createObjectURL(item.file) : item.existingImageUrl}
                                            alt={item.heading || 'Animation item image'}
                                            className={[styles.heroPreview, styles.animationItemPreview].join(' ')}
                                          />
                                          {!isReadOnlyMode && (
                                            <div className={styles.heroActions}>
                                              <button
                                                type="button"
                                                className={styles.secondaryButton}
                                                onClick={() =>
                                                  updateAnimationItem(section.clientId, item.clientId, (current) => ({
                                                    ...current,
                                                    file: null,
                                                    existingImageUrl: '',
                                                    existingObjectKey: '',
                                                  }))
                                                }
                                              >
                                                Remove image
                                              </button>
                                            </div>
                                          )}
                                        </div>
                                      ) : (
                                        <div className={styles.animationItemMedia}>
                                          <UploadDropzone
                                            label="Upload item image"
                                            hint="PNG, JPG, or WebP."
                                            accept={IMAGE_ACCEPT}
                                            icon={<CloudUploadIcon size={26} />}
                                            onFiles={(files) =>
                                              files[0] &&
                                              updateAnimationItem(section.clientId, item.clientId, (current) => ({
                                                ...current,
                                                file: files[0],
                                                existingImageUrl: '',
                                                existingObjectKey: '',
                                              }))
                                            }
                                            disabled={isReadOnlyMode}
                                          />
                                        </div>
                                      )}
                                    </div>

                                    {!isReadOnlyMode && (
                                      <div className={[styles.actionRow, styles.animationItemActions].join(' ')}>
                                        <button
                                          type="button"
                                          className={styles.iconButton}
                                          aria-label="Move item up"
                                          disabled={itemIndex === 0}
                                          onClick={() => moveAnimationItem(section.clientId, item.clientId, -1)}
                                        >
                                          <KeyboardArrowUpOutlinedIcon fontSize="small" />
                                        </button>
                                        <button
                                          type="button"
                                          className={styles.iconButton}
                                          aria-label="Move item down"
                                          disabled={itemIndex === section.animation.items.length - 1}
                                          onClick={() => moveAnimationItem(section.clientId, item.clientId, 1)}
                                        >
                                          <KeyboardArrowDownOutlinedIcon fontSize="small" />
                                        </button>
                                        <button
                                          type="button"
                                          className={styles.iconButtonDanger}
                                          aria-label="Remove animation item"
                                          onClick={() => removeAnimationItem(section.clientId, item.clientId)}
                                        >
                                          <DeleteIcon size={16} />
                                        </button>
                                      </div>
                                    )}
                                  </div>
                                ))}
                              </div>

                              {!isReadOnlyMode && (
                                <button
                                  type="button"
                                  className={styles.secondaryButton}
                                  onClick={() => addAnimationItem(section.clientId)}
                                >
                                  <AddIcon size={16} />
                                  Add animation item
                                </button>
                              )}
                            </>
                          )}
                        </div>
                      </div>
                    )}
                  </article>
                  )
                })}
              </div>
            ) : (
              <div className={styles.emptyState}>
                <div>
                  <p className={styles.emptyStateTitle}>No modules added yet</p>
                  <p className={styles.emptyStateText}>
                    Add a heading, image, typography, action, video, or animation block
                    to start building this story.
                  </p>
                </div>
              </div>
            )}

            {!isReadOnlyMode && (
              <div className={styles.actionsCard}>
                <div className={styles.segmentedControl}>
                  {moduleOptions.map((option) => (
                    <button
                      key={option.type}
                      type="button"
                      className={styles.secondaryButton}
                      onClick={() => addSection(option.type)}
                    >
                      <AddIcon size={16} />
                      {option.label}
                    </button>
                  ))}
                </div>
              </div>
            )}
          </section>
        </fieldset>
      </div>
    </CmsAppShell>
  )
}

function createEditorId(prefix: string) {
  if (typeof globalThis.crypto?.randomUUID === 'function') {
    return `${prefix}-${globalThis.crypto.randomUUID()}`
  }
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}`
}

function createDefaultAnimationItem(): AnimationItemState {
  return {
    clientId: createEditorId('animation-item-client'),
    heading: '',
    subHeading: '',
    description: '',
    file: null,
    existingImageUrl: '',
    existingObjectKey: '',
  }
}

function createDefaultSection(sectionType: BlogSectionType): BlogSectionState {
  return {
    clientId: createEditorId('blog-section-client'),
    sectionName: defaultSectionName(sectionType),
    sectionType,
    isEnabled: true,
    isCollapsed: false,
    heading: {
      headingText: '',
      underlineEnabled: false,
    },
    image: {
      caption: '',
      file: null,
      existingImageUrl: '',
      existingObjectKey: '',
    },
    typography: {
      htmlContent: '',
    },
    action: {
      text: '',
      actionType: 'link',
      targetUrl: '',
    },
    video: {
      youtubeUrl: '',
      caption: '',
    },
    animation: {
      navigation: 'vertical',
      imagePosition: 'left',
      items: [],
    },
  }
}

function createDefaultForm(): BlogFormState {
  return {
    publishDate: '',
    heading: '',
    description: '',
    coverImageFile: null,
    existingCoverImageUrl: '',
    existingCoverObjectKey: '',
    sections: [],
  }
}

function buildFormFromDetail(detail: BlogDetailResponse): BlogFormState {
  return {
    publishDate: detail.publish_date.slice(0, 10),
    heading: detail.heading,
    description: detail.description,
    coverImageFile: null,
    existingCoverImageUrl: resolveBlogAssetUrl(detail.cover_image_fetch_url || detail.cover_image_url),
    existingCoverObjectKey: detail.cover_image_object_key ?? '',
    sections: (detail.blog_detail?.sections ?? [])
      .slice()
      .sort((left, right) => left.sort_order - right.sort_order)
      .map(mapSectionFromResponse),
  }
}

function mapSectionFromResponse(section: BlogSectionResponse): BlogSectionState {
  const next = createDefaultSection(section.section_type)

  next.id = section.id
  next.sectionName = section.section_name
  next.sectionType = section.section_type
  next.isEnabled = section.is_enabled
  next.heading.headingText = section.heading?.heading_text ?? ''
  next.heading.underlineEnabled = section.heading?.underline_enabled ?? false
  next.image.caption = section.image?.caption ?? ''
  next.image.existingImageUrl = resolveBlogAssetUrl(section.image?.asset?.fetch_url ?? '')
  next.image.existingObjectKey = section.image?.asset?.gcp_object_key ?? ''
  next.typography.htmlContent = section.typography?.html_content ?? ''
  next.action.text = section.action?.text ?? ''
  next.action.actionType = section.action?.action_type ?? 'link'
  next.action.targetUrl = section.action?.target_url ?? ''
  next.video.youtubeUrl = section.video?.youtube_url ?? ''
  next.video.caption = section.video?.caption ?? ''
  next.animation.navigation = section.animation?.navigation ?? 'vertical'
  next.animation.imagePosition = section.animation?.image_position ?? 'left'
  next.animation.items = (section.animation?.items ?? [])
    .slice()
    .sort((left, right) => left.sort_order - right.sort_order)
    .map((item) => ({
      clientId: createEditorId('animation-item-client'),
      id: item.id,
      heading: item.heading,
      subHeading: item.sub_heading,
      description: item.description,
      file: null,
      existingImageUrl: resolveBlogAssetUrl(item.image?.fetch_url ?? ''),
      existingObjectKey: item.image?.gcp_object_key ?? '',
    }))

  return next
}

function buildSaveRequest(form: BlogFormState): SaveBlogRequest {
  const request: SaveBlogRequest = {
    publish_date: form.publishDate.trim(),
    heading: form.heading.trim(),
    description: form.description.trim(),
    remove_cover_image: false,
    blog_detail: {
      sections: form.sections.map((section, index) => buildSectionPayload(section, index)),
    },
  }

  if (form.coverImageFile) {
    request.cover_image = {
      file_name: form.coverImageFile.name,
      mime_type: form.coverImageFile.type || 'application/octet-stream',
    }
    request.coverImageFile = form.coverImageFile
  } else if (form.existingCoverImageUrl || form.existingCoverObjectKey) {
    request.cover_image = {
      file_url: form.existingCoverImageUrl,
      storage_uri: form.existingCoverImageUrl,
      gcp_object_key: form.existingCoverObjectKey || undefined,
    }
  }

  const sectionImageFiles = form.sections
    .map((section, sectionIndex) =>
      section.sectionType === 'image' && section.image.file
        ? {
            sectionIndex,
            file: section.image.file,
          }
        : null,
    )
    .filter((entry): entry is { sectionIndex: number; file: File } => Boolean(entry))
  if (sectionImageFiles.length) {
    request.sectionImageFiles = sectionImageFiles
  }

  const animationItemImageFiles = form.sections.flatMap((section, sectionIndex) =>
    section.sectionType === 'animation'
      ? section.animation.items
          .map((item, itemIndex) =>
            item.file
              ? {
                  sectionIndex,
                  itemIndex,
                  file: item.file,
                }
              : null,
          )
          .filter((entry): entry is { sectionIndex: number; itemIndex: number; file: File } => Boolean(entry))
      : [],
  )
  if (animationItemImageFiles.length) {
    request.animationItemImageFiles = animationItemImageFiles
  }

  return request
}

function buildSectionPayload(section: BlogSectionState, sortOrder: number) {
  return {
    id: section.id,
    section_name: section.sectionName.trim() || defaultSectionName(section.sectionType),
    section_type: section.sectionType,
    sort_order: sortOrder,
    is_enabled: section.isEnabled,
    settings: {},
    ...(section.sectionType === 'heading'
      ? {
          heading: {
            heading_text: section.heading.headingText.trim(),
            underline_enabled: section.heading.underlineEnabled,
          },
        }
      : {}),
    ...(section.sectionType === 'image'
      ? {
          image: {
            caption: section.image.caption.trim(),
            asset: buildImageAsset(section.image.file, section.image.existingImageUrl, section.image.existingObjectKey),
          },
        }
      : {}),
    ...(section.sectionType === 'typography'
      ? {
          typography: {
            html_content: section.typography.htmlContent,
            text_content: stripHtml(section.typography.htmlContent),
          },
        }
      : {}),
    ...(section.sectionType === 'action'
      ? {
          action: {
            text: section.action.text.trim(),
            action_type: section.action.actionType,
            target_url: section.action.targetUrl.trim(),
          },
        }
      : {}),
    ...(section.sectionType === 'video'
      ? {
          video: {
            youtube_url: section.video.youtubeUrl.trim(),
            caption: section.video.caption.trim(),
          },
        }
      : {}),
    ...(section.sectionType === 'animation'
      ? {
          animation: {
            navigation: section.animation.navigation,
            image_position: section.animation.imagePosition,
            items: section.animation.items.map((item, itemIndex) => ({
              id: item.id,
              sort_order: itemIndex,
              heading: item.heading.trim(),
              sub_heading: item.subHeading.trim(),
              description: item.description.trim(),
              image: buildImageAsset(item.file, item.existingImageUrl, item.existingObjectKey),
            })),
          },
        }
      : {}),
  }
}

function buildImageAsset(file: File | null, existingUrl: string, existingObjectKey: string) {
  if (file) {
    return {
      file_name: file.name,
      mime_type: file.type || 'application/octet-stream',
    }
  }
  return {
    file_url: existingUrl || undefined,
    storage_uri: existingUrl || undefined,
    gcp_object_key: existingObjectKey || undefined,
  }
}

function defaultSectionName(sectionType: BlogSectionType) {
  switch (sectionType) {
    case 'heading':
      return 'Heading Module'
    case 'image':
      return 'Image Module'
    case 'typography':
      return 'Typography Module'
    case 'action':
      return 'Action Module'
    case 'video':
      return 'Video Module'
    case 'animation':
      return 'Animation Module'
    default:
      return 'Content Module'
  }
}

function stripHtml(value: string) {
  return value.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim()
}

function reorderSections(
  sections: BlogSectionState[],
  draggedId: string,
  targetId: string,
  placement: SectionDropPlacement,
) {
  const fromIndex = sections.findIndex((section) => section.clientId === draggedId)
  const targetIndex = sections.findIndex((section) => section.clientId === targetId)
  if (fromIndex < 0 || targetIndex < 0 || fromIndex === targetIndex) {
    return sections
  }

  const next = sections.slice()
  const [dragged] = next.splice(fromIndex, 1)
  const adjustedTargetIndex = next.findIndex((section) => section.clientId === targetId)
  const insertIndex = placement === 'after' ? adjustedTargetIndex + 1 : adjustedTargetIndex
  next.splice(insertIndex, 0, dragged)
  return next
}
