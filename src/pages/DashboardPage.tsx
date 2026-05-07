import { useTranslation } from 'react-i18next'
import { CmsAppShell } from '../components/CmsAppShell'
import styles from '../styles/DashboardPage.module.css'

export function DashboardPage() {
  const { t } = useTranslation()

  return (
    <CmsAppShell activeKey="dashboard">
      <section className={styles.panel}>
        <p className={styles.eyebrow}>{t('dashboard.eyebrow')}</p>
        <h1 className={styles.title}>{t('dashboard.title')}</h1>
        <p className={styles.subtitle}>{t('dashboard.subtitle')}</p>
      </section>
    </CmsAppShell>
  )
}
