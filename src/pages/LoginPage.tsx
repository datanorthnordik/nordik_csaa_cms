import axios from 'axios'
import { useEffect, useState } from 'react'
import { useForm } from 'react-hook-form'
import { useTranslation } from 'react-i18next'
import { useNavigate } from 'react-router-dom'
import { authApi } from '../api/authApi'
import { AuthAlert } from '../components/AuthAlert'
import type { AlertState } from '../components/AuthAlert'
import { AuthLayout } from '../components/AuthLayout'
import { FormInput } from '../components/FormInput'
import { Loader } from '../components/Loader'
import { isValidEmail } from '../lib/validation'
import { setAuthSession, type AuthSession } from '../store/authSlice'
import { useAppDispatch } from '../store/hooks'
import styles from '../styles/LoginPage.module.css'

type LoginValues = { email: string; password: string; remember: boolean }

type LoginSubmitResult = {
  session?: AuthSession
}

type LoginPageProps = {
  onSubmit?: (
    values: LoginValues,
  ) => void | LoginSubmitResult | Promise<void | LoginSubmitResult>
}

async function submitLogin(values: LoginValues): Promise<LoginSubmitResult> {
  const response = await authApi.login({
    email: values.email,
    password: values.password,
    rememberMe: values.remember,
  })

  return {
    session: response.data,
  }
}

export function LoginPage({ onSubmit }: LoginPageProps) {
  const dispatch = useAppDispatch()
  const navigate = useNavigate()
  const { i18n, t } = useTranslation()
  const [alert, setAlert] = useState<AlertState | null>(null)
  const {
    register,
    handleSubmit,
    formState: { errors, isSubmitting },
    getValues,
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
  const handleLogin = onSubmit ?? submitLogin

  useEffect(() => {
    if (hasErrors) {
      void trigger()
    }
  }, [hasErrors, i18n.language, trigger])

  async function onSubmitForm(values: LoginValues) {
    setAlert(null)

    try {
      const result = await handleLogin(values)

      if (result?.session) {
        dispatch(
          setAuthSession({
            session: result.session,
            rememberMe: values.remember,
          }),
        )
        navigate('/events', { replace: true })
        return
      }

      setAlert({ type: 'success', message: t('auth.login.feedback.success') })
    } catch (error) {
      if (axios.isAxiosError(error)) {
        return
      }

      setAlert({
        type: 'error',
        message:
          error instanceof Error
            ? error.message
            : t('auth.login.feedback.errorGeneric'),
      })
    }
  }

  return (
    <AuthLayout>
      {isSubmitting && <Loader fullscreen label={t('auth.loading')} />}

      <h1 className={styles.title}>{t('auth.login.title')}</h1>

      <AuthAlert alert={alert} />

      <form className={styles.form} onSubmit={handleSubmit(onSubmitForm)} noValidate>
        <FormInput
          type="email"
          placeholder={t('auth.login.fields.emailPlaceholder')}
          variant="filled"
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
          type="password"
          placeholder={t('auth.login.fields.passwordPlaceholder')}
          variant="filled"
          autoComplete="current-password"
          disabled={isSubmitting}
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
          <input type="checkbox" disabled={isSubmitting} {...register('remember')} />
          <span>{t('auth.login.rememberMe')}</span>
        </label>

        <button type="submit" className={styles.primary} disabled={isSubmitting}>
          {t('auth.login.actions.signIn')}
        </button>

        <button
          type="button"
          className={styles.forgot}
          disabled={isSubmitting}
          onClick={() =>
            navigate('/forgot-password', {
              state: { email: getValues('email') },
            })
          }
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
        disabled={isSubmitting}
        onClick={() => navigate('/signup')}
      >
        {t('auth.login.actions.createAccount')}
      </button>
    </AuthLayout>
  )
}
