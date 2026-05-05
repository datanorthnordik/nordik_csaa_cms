import csaaLogo from '../assets/csaa_logo.png'
import nordikLogo from '../assets/nordik_logo.png'
import styles from './AuthHeader.module.css'

export function AuthHeader() {
  return (
    <div className={styles.header}>
      <img
        src={csaaLogo}
        alt="Children of Shingwauk Alumni Association"
        className={styles.logo}
      />
      <img
        src={nordikLogo}
        alt="Nordik Institute"
        className={styles.logo}
      />
    </div>
  )
}
