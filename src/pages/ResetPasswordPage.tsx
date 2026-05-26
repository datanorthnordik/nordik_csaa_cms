import axios from 'axios'
import { useEffect, useMemo, useState } from 'react'
import { useForm } from 'react-hook-form'
import { useTranslation } from 'react-i18next'
import { useLocation, useNavigate, useSearchParams } from 'react-router-dom'
import { authApi } from '../api/authApi'
import { AuthAlert } from '../components/AuthAlert'
import type { AlertState } from '../components/AuthAlert'
import { AuthLayout } from '../components/AuthLayout'
import { FormInput } from '../components/FormInput'
import { Loader } from '../components/Loader'
import { isValidEmail } from '../lib/validation'
import styles from '../styles/AuthActionPage.module.css'

type ResetPasswordValues = {
  email: string
  otp: string
  password: string
  confirmPassword: string
}

type ResetPasswordPageProps = {
  onSubmit?: (values: ResetPasswordValues) => void | Promise<void>
}

type ResetPasswordLocationState = {
  email?: string
  resetRequested?: boolean
} | null

function getResetToken(searchParams: URLSearchParams) {
  for (const key of ['otp', 'token', 'reset_token', 'resetToken', 'code']) {
    const value = searchParams.get(key)?.trim()
    if (value) {
      return value
    }
  }

  return ''
}

function getResetEmail(
  searchParams: URLSearchParams,
  locationState: ResetPasswordLocationState,
) {
  const emailFromState =
    typeof locationState?.email === 'string' ? locationState.email.trim() : ''

  if (emailFromState) {
    return emailFromState
  }

  for (const key of ['email']) {
    const value = searchParams.get(key)?.trim()
    if (value) {
      return value
    }
  }

  return ''
}

async function submitResetPassword(values: ResetPasswordValues) {
  await authApi.resetPassword({
    email: values.email,
    otp: values.otp,
    password: values.password,
  })
}

export function ResetPasswordPage({ onSubmit }: ResetPasswordPageProps) {
  const location = useLocation()
  const navigate = useNavigate()
  const [searchParams] = useSearchParams()
  const { i18n, t } = useTranslation()
  const [alert, setAlert] = useState<AlertState | null>(null)
  const locationState = location.state as ResetPasswordLocationState
  const emailFromState = useMemo(
    () => getResetEmail(searchParams, locationState),
    [locationState, searchParams],
  )
  const tokenFromLink = useMemo(() => getResetToken(searchParams), [searchParams])
  const entryAlert =
    locationState?.resetRequested
      ? {
          type: 'success' as const,
          message: locationState.email
            ? t('auth.resetPassword.feedback.codeSentWithEmail', {
                email: locationState.email,
              })
            : t('auth.resetPassword.feedback.codeSent'),
        }
      : null
  const {
    register,
    handleSubmit,
    formState: { errors, isSubmitting },
    reset,
    trigger,
    watch,
  } = useForm<ResetPasswordValues>({
    defaultValues: {
      email: emailFromState,
      otp: tokenFromLink,
      password: '',
      confirmPassword: '',
    },
    mode: 'onSubmit',
  })
  const hasErrors = Object.keys(errors).length > 0
  const passwordValue = watch('password')
  const handleResetPassword = onSubmit ?? submitResetPassword

  useEffect(() => {
    if (hasErrors) {
      void trigger()
    }
  }, [hasErrors, i18n.language, trigger])

  async function onSubmitForm(values: ResetPasswordValues) {
    setAlert(null)

    try {
      await handleResetPassword(values)
      reset({
        email: values.email,
        otp: '',
        password: '',
        confirmPassword: '',
      })
      setAlert({
        type: 'success',
        message: t('auth.resetPassword.feedback.success'),
      })
    } catch (error) {
      if (axios.isAxiosError(error)) {
        return
      }

      setAlert({
        type: 'error',
        message:
          error instanceof Error
            ? error.message
            : t('auth.resetPassword.feedback.errorGeneric'),
      })
    }
  }

  return (
    <AuthLayout>
      {isSubmitting && <Loader fullscreen label={t('auth.loading')} />}

      <div className={styles.intro}>
        <h1 className={styles.title}>{t('auth.resetPassword.title')}</h1>
        <p className={styles.subtitle}>{t('auth.resetPassword.subtitle')}</p>
      </div>

      <AuthAlert alert={alert ?? entryAlert} />

      <form className={styles.form} onSubmit={handleSubmit(onSubmitForm)} noValidate>
        <FormInput
          label={t('auth.resetPassword.fields.email')}
          type="email"
          autoComplete="email"
          disabled={isSubmitting}
          {...register('email', {
            required: t('auth.validation.emailRequired'),
            validate: (value) =>
              isValidEmail(value) || t('auth.validation.emailInvalid'),
          })}
        />
        {errors.email && (
          <div className={styles.error} role="alert">
            {errors.email.message}
          </div>
        )}

        <FormInput
          label={t('auth.resetPassword.fields.token')}
          autoComplete="one-time-code"
          disabled={isSubmitting}
          {...register('otp', {
            required: t('auth.validation.resetTokenRequired'),
          })}
        />
        {errors.otp && (
          <div className={styles.error} role="alert">
            {errors.otp.message}
          </div>
        )}

        <FormInput
          label={t('auth.resetPassword.fields.password')}
          type="password"
          autoComplete="new-password"
          disabled={isSubmitting}
          {...register('password', {
            required: t('auth.validation.passwordRequired'),
            minLength: {
              value: 6,
              message: t('auth.validation.passwordMinLength'),
            },
          })}
        />
        {errors.password && (
          <div className={styles.error} role="alert">
            {errors.password.message}
          </div>
        )}

        <FormInput
          label={t('auth.resetPassword.fields.confirmPassword')}
          type="password"
          autoComplete="new-password"
          disabled={isSubmitting}
          {...register('confirmPassword', {
            required: t('auth.validation.confirmPasswordRequired'),
            validate: (value) =>
              value === passwordValue || t('auth.validation.passwordsDoNotMatch'),
          })}
        />
        {errors.confirmPassword && (
          <div className={styles.error} role="alert">
            {errors.confirmPassword.message}
          </div>
        )}

        <button
          type="submit"
          className={styles.primary}
          disabled={isSubmitting}
        >
          {t('auth.resetPassword.actions.resetPassword')}
        </button>
      </form>

      <p className={styles.helper}>{t('auth.resetPassword.helper')}</p>

      <button
        type="button"
        className={styles.secondary}
        disabled={isSubmitting}
        onClick={() => navigate('/')}
      >
        {t('auth.resetPassword.actions.backToSignIn')}
      </button>
    </AuthLayout>
  )
}
