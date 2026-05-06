import { useEffect, useState } from 'react'
import { useForm } from 'react-hook-form'
import { useTranslation } from 'react-i18next'
import { useNavigate } from 'react-router-dom'
import { AuthAlert } from '../components/AuthAlert'
import type { AlertState } from '../components/AuthAlert'
import { AuthLayout } from '../components/AuthLayout'
import { FormInput } from '../components/FormInput'
import { isValidEmail } from '../lib/validation'
import styles from '../styles/LoginPage.module.css'

type LoginValues = { email: string; password: string; remember: boolean }

type LoginPageProps = {
  onSubmit?: (values: LoginValues) => void | Promise<void>
}

export function LoginPage({ onSubmit }: LoginPageProps) {
  const navigate = useNavigate()
  const { i18n, t } = useTranslation()
  const [alert, setAlert] = useState<AlertState | null>(null)
  const {
    register,
    handleSubmit,
    formState: { errors },
    trigger,
  } = useForm<LoginValues>({
    defaultValues: {
      email: '',
      password: '',
      remember: false,
    },
    mode: 'onSubmit',
  })
  const hasErrors = Object.keys(errors).length > 0

  useEffect(() => {
    if (hasErrors) {
      void trigger()
    }
  }, [hasErrors, i18n.language, trigger])

  async function onSubmitForm(values: LoginValues) {
    setAlert(null)

    try {
      await onSubmit?.(values)
      setAlert({ type: 'success', message: t('auth.login.feedback.success') })
    } catch (err) {
      setAlert({
        type: 'error',
        message:
          err instanceof Error
            ? err.message
            : t('auth.login.feedback.errorGeneric'),
      })
    }
  }

  return (
    <AuthLayout>
      <h1 className={styles.title}>{t('auth.login.title')}</h1>

      <AuthAlert alert={alert} />

      <form className={styles.form} onSubmit={handleSubmit(onSubmitForm)} noValidate>
        <FormInput
          type="email"
          placeholder={t('auth.login.fields.emailPlaceholder')}
          variant="filled"
          autoComplete="email"
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
          type="password"
          placeholder={t('auth.login.fields.passwordPlaceholder')}
          variant="filled"
          autoComplete="current-password"
          {...register('password', {
            required: t('auth.validation.passwordRequired'),
          })}
        />
        {errors.password && (
          <div className={styles.error} role="alert">
            {errors.password.message}
          </div>
        )}

        <label className={styles.remember}>
          <input type="checkbox" {...register('remember')} />
          <span>{t('auth.login.rememberMe')}</span>
        </label>

        <button type="submit" className={styles.primary}>
          {t('auth.login.actions.signIn')}
        </button>

        <button
          type="button"
          className={styles.forgot}
          onClick={() => console.log('forgot password')}
        >
          {t('auth.login.actions.forgotPassword')}
        </button>
      </form>

      <div className={styles.divider}>
        <span>{t('auth.login.divider')}</span>
      </div>

      <button
        type="button"
        className={styles.secondary}
        onClick={() => navigate('/signup')}
      >
        {t('auth.login.actions.createAccount')}
      </button>
    </AuthLayout>
  )
}
