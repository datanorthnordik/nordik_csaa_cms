import type { ReactNode } from 'react'
import { AuthHeader } from './AuthHeader'
import styles from './AuthLayout.module.css'

type AuthLayoutProps = {
  children: ReactNode
}

export function AuthLayout({ children }: AuthLayoutProps) {
  return (
    <div className={styles.page}>
      <div className={styles.card}>
        <AuthHeader />
        <div className={styles.body}>{children}</div>
      </div>
    </div>
  )
}
