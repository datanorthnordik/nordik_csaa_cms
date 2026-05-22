import axios from 'axios'
import { useEffect, useMemo, useState } from 'react'
import { useForm } from 'react-hook-form'
import { useTranslation } from 'react-i18next'
import { useNavigate, useSearchParams } from 'react-router-dom'
import { authApi } from '../api/authApi'
import { AuthAlert } from '../components/AuthAlert'
import type { AlertState } from '../components/AuthAlert'
import { AuthLayout } from '../components/AuthLayout'
import { FormInput } from '../components/FormInput'
import { Loader } from '../components/Loader'
import styles from '../styles/AuthActionPage.module.css'

type ResetPasswordValues = {
  token: string
  password: string
  confirmPassword: string
}

type ResetPasswordPageProps = {
  onSubmit?: (values: ResetPasswordValues) => void | Promise<void>
}

function getResetToken(searchParams: URLSearchParams) {
  for (const key of ['token', 'reset_token', 'resetToken', 'code']) {
    const value = searchParams.get(key)?.trim()
    if (value) {
      return value
    }
  }

  return ''
}

async function submitResetPassword(values: ResetPasswordValues) {
  await authApi.resetPassword({
    token: values.token,
    password: values.password,
  })
}

export function ResetPasswordPage({ onSubmit }: ResetPasswordPageProps) {
  const navigate = useNavigate()
  const [searchParams] = useSearchParams()
  const { i18n, t } = useTranslation()
  const [alert, setAlert] = useState<AlertState | null>(null)
  const tokenFromLink = useMemo(() => getResetToken(searchParams), [searchParams])
  const {
    register,
    handleSubmit,
    formState: { errors, isSubmitting },
    reset,
    trigger,
    watch,
  } = useForm<ResetPasswordValues>({
    defaultValues: {
      token: tokenFromLink,
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
        token: values.token,
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

      <AuthAlert alert={alert} />

      <form className={styles.form} onSubmit={handleSubmit(onSubmitForm)} noValidate>
        <FormInput
          label={t('auth.resetPassword.fields.token')}
          autoComplete="one-time-code"
          disabled={isSubmitting}
          {...register('token', {
            required: t('auth.validation.resetTokenRequired'),
          })}
        />
        {errors.token && (
          <div className={styles.error} role="alert">
            {errors.token.message}
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
