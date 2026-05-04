import { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { useForm } from 'react-hook-form'
import { AuthLayout } from '../components/AuthLayout'
import { FormInput } from '../components/FormInput'
import { AuthAlert } from '../components/AuthAlert'
import type { AlertState } from '../components/AuthAlert'
import styles from '../styles/LoginPage.module.css'

type LoginValues = { email: string; password: string; remember: boolean }

type LoginPageProps = {
  onSubmit?: (values: LoginValues) => void | Promise<void>
}

export function LoginPage({ onSubmit }: LoginPageProps) {
  const navigate = useNavigate()
  const [alert, setAlert] = useState<AlertState | null>(null)
  const { register, handleSubmit, formState: { errors }, watch } = useForm<LoginValues>({
    defaultValues: {
      email: '',
      password: '',
      remember: false
    },
    mode: 'onSubmit'
  })

  const emailValue = watch('email')
  const passwordValue = watch('password')

  async function onSubmitForm(values: LoginValues) {
    setAlert(null)

    try {
      await onSubmit?.(values)
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

      <form className={styles.form} onSubmit={handleSubmit(onSubmitForm)} noValidate>
        <FormInput
          type="email"
          placeholder="Email Address"
          variant="filled"
          autoComplete="email"
          {...register('email', { 
            required: 'Email is required',
            pattern: {
              value: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
              message: 'Please enter a valid email address'
            }
          })}
        />
        {errors.email && (
          <div className={styles.error} role="alert">
            {errors.email.message}
          </div>
        )}

        <FormInput
          type="password"
          placeholder="Password"
          variant="filled"
          autoComplete="current-password"
          {...register('password', { required: 'Password is required' })}
        />
        {errors.password && (
          <div className={styles.error} role="alert">
            {errors.password.message}
          </div>
        )}

        <label className={styles.remember}>
          <input
            type="checkbox"
            {...register('remember')}
          />
          <span>Remember Me</span>
        </label>

        <button type="submit" className={styles.primary}>
          SIGN IN
        </button>

        <button
          type="button"
          className={styles.forgot}
          onClick={() => console.log('forgot password')}
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
        onClick={() => navigate('/signup')}
      >
        CREATE NEW ACCOUNT
      </button>
    </AuthLayout>
  )
}
