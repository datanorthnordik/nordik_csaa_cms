import i18n from 'i18next'
import { initReactI18next } from 'react-i18next'
import { defaultLanguage, fallbackNamespace, supportedLanguages } from './config'
import enCommon from './locales/en/common.json'
import frCommon from './locales/fr/common.json'

const resources = {
  en: {
    common: enCommon,
  },
  fr: {
    common: frCommon,
  },
} as const

const isSupportedLanguage = (language: string): language is keyof typeof resources =>
  supportedLanguages.some(({ code }) => code === language)

const getInitialLanguage = () => {
  if (typeof window === 'undefined') {
    return defaultLanguage
  }

  const savedLanguage = window.localStorage.getItem('language')
  if (savedLanguage && isSupportedLanguage(savedLanguage)) {
    return savedLanguage
  }

  const browserLanguage = window.navigator.language.split('-')[0]
  if (isSupportedLanguage(browserLanguage)) {
    return browserLanguage
  }

  return defaultLanguage
}

if (!i18n.isInitialized) {
  void i18n.use(initReactI18next).init({
    resources,
    lng: getInitialLanguage(),
    fallbackLng: defaultLanguage,
    supportedLngs: supportedLanguages.map(({ code }) => code),
    defaultNS: fallbackNamespace,
    initAsync: false,
    interpolation: {
      escapeValue: false,
    },
  })

  if (typeof window !== 'undefined') {
    i18n.on('languageChanged', (language) => {
      window.localStorage.setItem('language', language)
    })
  }
}

export default i18n
