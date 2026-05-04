import { useState } from 'react'
import type { FormEvent } from 'react'
import { AuthLayout } from '../components/AuthLayout'
import { FormInput } from '../components/FormInput'
import { AuthAlert } from '../components/AuthAlert'
import type { AlertState } from '../components/AuthAlert'
import { isValidEmail } from '../lib/validation'
import styles from '../styles/LoginPage.module.css'

type LoginValues = { email: string; password: string; remember: boolean }

type LoginPageProps = {
  onSubmit?: (values: LoginValues) => void | Promise<void>
  onCreateAccount?: () => void
  onForgotPassword?: () => void
}

export function LoginPage({
  onSubmit,
  onCreateAccount,
  onForgotPassword,
}: LoginPageProps) {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [remember, setRemember] = useState(false)
  const [alert, setAlert] = useState<AlertState | null>(null)

  async function handleSubmit(e: FormEvent<HTMLFormElement>) {
    e.preventDefault()
    setAlert(null)

    if (!email || !password) {
      setAlert({
        type: 'error',
        message: 'Please enter your email and password.',
      })
      return
    }

    if (!isValidEmail(email)) {
      setAlert({
        type: 'error',
        message: 'Please enter a valid email address.',
      })
      return
    }

    try {
      await onSubmit?.({ email, password, remember })
      setAlert({ type: 'success', message: 'Signing you in…' })
    } catch (err) {
      setAlert({
        type: 'error',
        message:
          err instanceof Error ? err.message : 'Sign in failed. Please try again.',
      })
    }
  }

  return (
    <AuthLayout>
      <h1 className={styles.title}>Sign in with Email Address</h1>

      <AuthAlert alert={alert} />

      <form className={styles.form} onSubmit={handleSubmit} noValidate>
        <FormInput
          type="email"
          placeholder="Email Address"
          variant="filled"
          autoComplete="email"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
        />

        <FormInput
          type="password"
          placeholder="Password"
          variant="filled"
          autoComplete="current-password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
        />

        <label className={styles.remember}>
          <input
            type="checkbox"
            checked={remember}
            onChange={(e) => setRemember(e.target.checked)}
          />
          <span>Remember Me</span>
        </label>

        <button type="submit" className={styles.primary}>
          SIGN IN
        </button>

        <button
          type="button"
          className={styles.forgot}
          onClick={onForgotPassword}
        >
          Forgot password
        </button>
      </form>

      <div className={styles.divider}>
        <span>OR</span>
      </div>

      <button
        type="button"
        className={styles.secondary}
        onClick={onCreateAccount}
      >
        CREATE NEW ACCOUNT
      </button>
    </AuthLayout>
  )
}
