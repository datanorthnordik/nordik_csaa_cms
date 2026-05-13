import { useEffect, useMemo, useState } from 'react'
import toast from 'react-hot-toast'
import { useTranslation } from 'react-i18next'
import { getApiErrorMessage } from '../api/apiError'
import { menusApi, type MenuResponse, type MenuPageOption } from '../api/menusApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import { MenuBuilder } from '../components/navigation/MenuBuilder'
import {
  buildSaveMenuItems,
  buildMenuEditorItems,
  serializeMenuItems,
  type MenuEditorItem,
} from '../lib/menusEditor'
import styles from '../styles/MenusPage.module.css'

type AsyncState = 'idle' | 'loading' | 'succeeded' | 'failed'

const MENU_KEY = 'main'

export function MenusPage() {
  const { t } = useTranslation()
  const [menu, setMenu] = useState<MenuResponse | null>(null)
  const [items, setItems] = useState<MenuEditorItem[]>([])
  const [pageOptions, setPageOptions] = useState<MenuPageOption[]>([])
  const [savedSnapshot, setSavedSnapshot] = useState(serializeMenuItems([]))
  const [loadStatus, setLoadStatus] = useState<AsyncState>('loading')
  const [loadError, setLoadError] = useState<string | null>(null)
  const [saveStatus, setSaveStatus] = useState<AsyncState>('idle')
  const [saveError, setSaveError] = useState<string | null>(null)

  useEffect(() => {
    let cancelled = false

    async function loadMenuEditor() {
      setLoadStatus('loading')
      setLoadError(null)

      try {
        const [menuResponse, pageOptionsResponse] = await Promise.all([
          menusApi.getMenu(MENU_KEY),
          menusApi.listMenuPageOptions(MENU_KEY),
        ])
        if (cancelled) {
          return
        }

        const nextItems = buildMenuEditorItems(menuResponse.items ?? [])
        setMenu(menuResponse)
        setItems(nextItems)
        setPageOptions(pageOptionsResponse)
        setSavedSnapshot(serializeMenuItems(nextItems))
        setLoadStatus('succeeded')
      } catch (error) {
        if (cancelled) {
          return
        }

        setLoadError(getApiErrorMessage(error))
        setLoadStatus('failed')
      }
    }

    void loadMenuEditor()

    return () => {
      cancelled = true
    }
  }, [])

  const isDirty = useMemo(
    () => serializeMenuItems(items) !== savedSnapshot,
    [items, savedSnapshot],
  )

  async function handleSave() {
    if (!menu) {
      return
    }

    setSaveStatus('loading')
    setSaveError(null)

    try {
      const response = await menusApi.saveMenu(MENU_KEY, {
        name: menu.name,
        items: buildSaveMenuItems(items),
      })

      const nextItems = buildMenuEditorItems(response.menu.items ?? [])
      setMenu(response.menu)
      setItems(nextItems)
      setSavedSnapshot(serializeMenuItems(nextItems))
      setSaveStatus('succeeded')
      toast.success(t('menus.feedback.saved'))
    } catch (error) {
      const message = getApiErrorMessage(error)
      setSaveStatus('failed')
      setSaveError(message)
      toast.error(message)
    }
  }

  function handleDiscard() {
    if (!menu) {
      return
    }

    const nextItems = buildMenuEditorItems(menu.items ?? [])
    setItems(nextItems)
    setSavedSnapshot(serializeMenuItems(nextItems))
    setSaveError(null)
    setSaveStatus('idle')
  }

  if (loadStatus === 'loading' && !menu) {
    return (
      <CmsAppShell activeKey="menus">
        <div className={styles.loaderWrap}>
          <Loader label={t('menus.common.loading')} />
        </div>
      </CmsAppShell>
    )
  }

  if (loadStatus === 'failed' && !menu) {
    return (
      <CmsAppShell activeKey="menus">
        <section className={styles.errorPanel}>
          <h1>{t('menus.loadError.title')}</h1>
          <p>{loadError}</p>
          <button type="button" className={styles.secondaryButton} onClick={() => window.location.reload()}>
            {t('menus.actions.retry')}
          </button>
        </section>
      </CmsAppShell>
    )
  }

  return (
    <CmsAppShell activeKey="menus">
      <div className={styles.page}>
        <Breadcrumb items={[{ label: t('menus.breadcrumb.title') }]} />

        <MenuBuilder
          title={menu?.name ?? t('menus.title')}
          subtitle={t('menus.subtitle')}
          items={items}
          pageOptions={pageOptions}
          isDirty={isDirty}
          isSaving={saveStatus === 'loading'}
          saveError={saveError ?? loadError}
          onChange={setItems}
          onSave={handleSave}
          onDiscard={handleDiscard}
        />
      </div>
    </CmsAppShell>
  )
}
