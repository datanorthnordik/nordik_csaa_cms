import { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { useForm } from 'react-hook-form'
import { AuthLayout } from '../components/AuthLayout'
import { FormInput } from '../components/FormInput'
import { AuthAlert } from '../components/AuthAlert'
import type { AlertState } from '../components/AuthAlert'
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
}

export function SignupPage({ onSubmit }: SignupPageProps) {
  const navigate = useNavigate()
  const [alert, setAlert] = useState<AlertState | null>(null)
  const { register, handleSubmit, formState: { errors }, watch } = useForm<SignupValues>({
    defaultValues: {
      firstName: '',
      lastName: '',
      email: '',
      password: '',
      confirmPassword: '',
    },
    mode: 'onSubmit'
  })

  const passwordValue = watch('password')
  const confirmPasswordValue = watch('confirmPassword')

  async function onSubmitForm(values: SignupValues) {
    setAlert(null)

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

      <form className={styles.form} onSubmit={handleSubmit(onSubmitForm)} noValidate>
        <div className={styles.row}>
          <FormInput
            label="First Name"
            autoComplete="given-name"
            {...register('firstName', { required: 'First name is required' })}
          />
          <FormInput
            label="Last Name"
            autoComplete="family-name"
            {...register('lastName', { required: 'Last name is required' })}
          />
        </div>
        {(errors.firstName || errors.lastName) && (
          <div className={styles.error} role="alert">
            {errors.firstName?.message || errors.lastName?.message}
          </div>
        )}

        <FormInput
          label="Work Email Address"
          type="email"
          autoComplete="email"
          {...register('email', { 
            required: 'Email is required',
            pattern: {
              value: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
              message: 'Please enter a valid work email address'
            }
          })}
        />
        {errors.email && (
          <div className={styles.error} role="alert">
            {errors.email.message}
          </div>
        )}

        <div className={styles.row}>
          <FormInput
            label="Password"
            type="password"
            autoComplete="new-password"
            {...register('password', { 
              required: 'Password is required',
              minLength: {
                value: 8,
                message: 'Password must be at least 8 characters'
              }
            })}
          />
          <FormInput
            label="Confirm Password"
            type="password"
            autoComplete="new-password"
            {...register('confirmPassword', { 
              required: 'Please confirm your password',
              validate: (value) => value === watch('password') || 'Passwords do not match'
            })}
          />
        </div>
        {(errors.password || errors.confirmPassword) && (
          <div className={styles.error} role="alert">
            {errors.password?.message || errors.confirmPassword?.message}
          </div>
        )}

        <button type="submit" className={styles.primary}>
          CREATE ACCOUNT
        </button>
      </form>

      <div className={styles.divider} />

      <p className={styles.footer}>
        Already have an account?{' '}
        <button type="button" className={styles.link} onClick={() => navigate('/')}>
          Sign In
        </button>
      </p>
    </AuthLayout>
  )
}
