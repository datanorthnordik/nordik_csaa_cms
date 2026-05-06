export const defaultLanguage = 'en'
export const fallbackNamespace = 'common'

export const supportedLanguages = [
  { code: 'en', label: 'EN' },
  { code: 'fr', label: 'FR' },
] as const

export type SupportedLanguageCode = (typeof supportedLanguages)[number]['code']
