import axios from 'axios'
import { useEffect, useState } from 'react'
import { useForm } from 'react-hook-form'
import { useTranslation } from 'react-i18next'
import { useLocation, useNavigate } from 'react-router-dom'
import { authApi } from '../api/authApi'
import { AuthAlert } from '../components/AuthAlert'
import type { AlertState } from '../components/AuthAlert'
import { AuthLayout } from '../components/AuthLayout'
import { FormInput } from '../components/FormInput'
import { Loader } from '../components/Loader'
import { isValidEmail } from '../lib/validation'
import styles from '../styles/AuthActionPage.module.css'

type ForgotPasswordValues = {
  email: string
}

type ForgotPasswordPageProps = {
  onSubmit?: (values: ForgotPasswordValues) => void | Promise<void>
}

type ForgotPasswordLocationState = {
  email?: string
} | null

async function submitForgotPassword(values: ForgotPasswordValues) {
  await authApi.forgotPassword({
    email: values.email,
  })
}

export function ForgotPasswordPage({ onSubmit }: ForgotPasswordPageProps) {
  const location = useLocation()
  const navigate = useNavigate()
  const { i18n, t } = useTranslation()
  const [alert, setAlert] = useState<AlertState | null>(null)
  const locationState = location.state as ForgotPasswordLocationState
  const defaultEmail =
    typeof locationState?.email === 'string' ? locationState.email : ''
  const {
    register,
    handleSubmit,
    formState: { errors, isSubmitting },
    trigger,
  } = useForm<ForgotPasswordValues>({
    defaultValues: {
      email: defaultEmail,
    },
    mode: 'onSubmit',
  })
  const hasErrors = Object.keys(errors).length > 0
  const handleForgotPassword = onSubmit ?? submitForgotPassword

  useEffect(() => {
    if (hasErrors) {
      void trigger()
    }
  }, [hasErrors, i18n.language, trigger])

  async function onSubmitForm(values: ForgotPasswordValues) {
    setAlert(null)

    try {
      await handleForgotPassword(values)
      setAlert({
        type: 'success',
        message: t('auth.forgotPassword.feedback.success'),
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
            : t('auth.forgotPassword.feedback.errorGeneric'),
      })
    }
  }

  return (
    <AuthLayout>
      {isSubmitting && <Loader fullscreen label={t('auth.loading')} />}

      <div className={styles.intro}>
        <h1 className={styles.title}>{t('auth.forgotPassword.title')}</h1>
        <p className={styles.subtitle}>{t('auth.forgotPassword.subtitle')}</p>
      </div>

      <AuthAlert alert={alert} />

      <form className={styles.form} onSubmit={handleSubmit(onSubmitForm)} noValidate>
        <FormInput
          label={t('auth.forgotPassword.fields.email')}
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

        <button type="submit" className={styles.primary} disabled={isSubmitting}>
          {t('auth.forgotPassword.actions.sendLink')}
        </button>
      </form>

      <p className={styles.helper}>{t('auth.forgotPassword.helper')}</p>

      <button
        type="button"
        className={styles.secondary}
        disabled={isSubmitting}
        onClick={() => navigate('/')}
      >
        {t('auth.forgotPassword.actions.backToSignIn')}
      </button>
    </AuthLayout>
  )
}
