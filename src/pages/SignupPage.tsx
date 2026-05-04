import { useState } from 'react'
import type { FormEvent } from 'react'
import { AuthLayout } from '../components/AuthLayout'
import { FormInput } from '../components/FormInput'
import { AuthAlert } from '../components/AuthAlert'
import type { AlertState } from '../components/AuthAlert'
import { isValidEmail } from '../lib/validation'
import styles from '../styles/SignupPage.module.css'

export type SignupValues = {
  firstName: string
  lastName: string
  email: string
  password: string
  confirmPassword: string
}

type SignupPageProps = {
  onSubmit?: (values: SignupValues) => void | Promise<void>
  onSignIn?: () => void
}

export function SignupPage({ onSubmit, onSignIn }: SignupPageProps) {
  const [values, setValues] = useState<SignupValues>({
    firstName: '',
    lastName: '',
    email: '',
    password: '',
    confirmPassword: '',
  })
  const [alert, setAlert] = useState<AlertState | null>(null)

  function update<K extends keyof SignupValues>(key: K, value: SignupValues[K]) {
    setValues((prev) => ({ ...prev, [key]: value }))
  }

  async function handleSubmit(e: FormEvent<HTMLFormElement>) {
    e.preventDefault()
    setAlert(null)

    const { firstName, lastName, email, password, confirmPassword } = values

    if (!firstName || !lastName || !email || !password || !confirmPassword) {
      setAlert({ type: 'error', message: 'Please fill in all fields.' })
      return
    }
    if (!isValidEmail(email)) {
      setAlert({
        type: 'error',
        message: 'Please enter a valid work email address.',
      })
      return
    }
    if (password.length < 8) {
      setAlert({
        type: 'error',
        message: 'Password must be at least 8 characters.',
      })
      return
    }
    if (password !== confirmPassword) {
      setAlert({ type: 'error', message: 'Passwords do not match.' })
      return
    }

    try {
      await onSubmit?.(values)
      setAlert({
        type: 'success',
        message:
          'Account request submitted. You will receive an email once approved.',
      })
    } catch (err) {
      setAlert({
        type: 'error',
        message:
          err instanceof Error
            ? err.message
            : 'Could not create account. Please try again.',
      })
    }
  }

  return (
    <AuthLayout>
      <div className={styles.intro}>
        <h1 className={styles.title}>Create Your Account</h1>
        <p className={styles.subtitle}>
          Request access to the administrative console.
        </p>
      </div>

      <AuthAlert alert={alert} />

      <form className={styles.form} onSubmit={handleSubmit} noValidate>
        <div className={styles.row}>
          <FormInput
            label="First Name"
            autoComplete="given-name"
            value={values.firstName}
            onChange={(e) => update('firstName', e.target.value)}
          />
          <FormInput
            label="Last Name"
            autoComplete="family-name"
            value={values.lastName}
            onChange={(e) => update('lastName', e.target.value)}
          />
        </div>

        <FormInput
          label="Work Email Address"
          type="email"
          autoComplete="email"
          value={values.email}
          onChange={(e) => update('email', e.target.value)}
        />

        <div className={styles.row}>
          <FormInput
            label="Password"
            type="password"
            autoComplete="new-password"
            value={values.password}
            onChange={(e) => update('password', e.target.value)}
          />
          <FormInput
            label="Confirm Password"
            type="password"
            autoComplete="new-password"
            value={values.confirmPassword}
            onChange={(e) => update('confirmPassword', e.target.value)}
          />
        </div>

        <button type="submit" className={styles.primary}>
          CREATE ACCOUNT
        </button>
      </form>

      <div className={styles.divider} />

      <p className={styles.footer}>
        Already have an account?{' '}
        <button type="button" className={styles.link} onClick={onSignIn}>
          Sign In
        </button>
      </p>
    </AuthLayout>
  )
}
