import { useEffect, useState } from 'react'
import { useForm } from 'react-hook-form'
import { useTranslation } from 'react-i18next'
import { useNavigate } from 'react-router-dom'
import { AuthAlert } from '../components/AuthAlert'
import type { AlertState } from '../components/AuthAlert'
import { AuthLayout } from '../components/AuthLayout'
import { FormInput } from '../components/FormInput'
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
}

export function SignupPage({ onSubmit }: SignupPageProps) {
  const navigate = useNavigate()
  const { i18n, t } = useTranslation()
  const [alert, setAlert] = useState<AlertState | null>(null)
  const {
    register,
    handleSubmit,
    formState: { errors },
    trigger,
    watch,
  } = useForm<SignupValues>({
    defaultValues: {
      firstName: '',
      lastName: '',
      email: '',
      password: '',
      confirmPassword: '',
    },
    mode: 'onSubmit',
  })
  const hasErrors = Object.keys(errors).length > 0
  const passwordValue = watch('password')

  useEffect(() => {
    if (hasErrors) {
      void trigger()
    }
  }, [hasErrors, i18n.language, trigger])

  async function onSubmitForm(values: SignupValues) {
    setAlert(null)

    try {
      await onSubmit?.(values)
      setAlert({
        type: 'success',
        message: t('auth.signup.feedback.success'),
      })
    } catch (err) {
      setAlert({
        type: 'error',
        message:
          err instanceof Error
            ? err.message
            : t('auth.signup.feedback.errorGeneric'),
      })
    }
  }

  return (
    <AuthLayout>
      <div className={styles.intro}>
        <h1 className={styles.title}>{t('auth.signup.title')}</h1>
        <p className={styles.subtitle}>{t('auth.signup.subtitle')}</p>
      </div>

      <AuthAlert alert={alert} />

      <form className={styles.form} onSubmit={handleSubmit(onSubmitForm)} noValidate>
        <div className={styles.row}>
          <FormInput
            label={t('auth.signup.fields.firstName')}
            autoComplete="given-name"
            {...register('firstName', {
              required: t('auth.validation.firstNameRequired'),
            })}
          />
          <FormInput
            label={t('auth.signup.fields.lastName')}
            autoComplete="family-name"
            {...register('lastName', {
              required: t('auth.validation.lastNameRequired'),
            })}
          />
        </div>
        {(errors.firstName || errors.lastName) && (
          <div className={styles.error} role="alert">
            {errors.firstName?.message || errors.lastName?.message}
          </div>
        )}

        <FormInput
          label={t('auth.signup.fields.email')}
          type="email"
          autoComplete="email"
          {...register('email', {
            required: t('auth.validation.emailRequired'),
            validate: (value) =>
              isValidEmail(value) || t('auth.validation.workEmailInvalid'),
          })}
        />
        {errors.email && (
          <div className={styles.error} role="alert">
            {errors.email.message}
          </div>
        )}

        <div className={styles.row}>
          <FormInput
            label={t('auth.signup.fields.password')}
            type="password"
            autoComplete="new-password"
            {...register('password', {
              required: t('auth.validation.passwordRequired'),
              minLength: {
                value: 8,
                message: t('auth.validation.passwordMinLength'),
              },
            })}
          />
          <FormInput
            label={t('auth.signup.fields.confirmPassword')}
            type="password"
            autoComplete="new-password"
            {...register('confirmPassword', {
              required: t('auth.validation.confirmPasswordRequired'),
              validate: (value) =>
                value === passwordValue ||
                t('auth.validation.passwordsDoNotMatch'),
            })}
          />
        </div>
        {(errors.password || errors.confirmPassword) && (
          <div className={styles.error} role="alert">
            {errors.password?.message || errors.confirmPassword?.message}
          </div>
        )}

        <button type="submit" className={styles.primary}>
          {t('auth.signup.actions.createAccount')}
        </button>
      </form>

      <div className={styles.divider} />

      <p className={styles.footer}>
        {t('auth.signup.footer.prompt')}{' '}
        <button type="button" className={styles.link} onClick={() => navigate('/')}>
          {t('auth.signup.footer.signIn')}
        </button>
      </p>
    </AuthLayout>
  )
}
