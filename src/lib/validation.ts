const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/
const PHONE_ALLOWED_CHARS_RE = /^\+?[\d\s().-]+$/
const EXTENSION_RE = /^\d{1,6}$/

export function isValidEmail(email: string): boolean {
  const trimmed = email.trim()
  if (!trimmed || !EMAIL_RE.test(trimmed) || trimmed.includes('..')) {
    return false
  }

  const [localPart, domain] = trimmed.split('@')
  if (!localPart || !domain) {
    return false
  }

  if (
    localPart.startsWith('.') ||
    localPart.endsWith('.') ||
    domain.startsWith('.') ||
    domain.endsWith('.')
  ) {
    return false
  }

  return domain
    .split('.')
    .every((label) => label && !label.startsWith('-') && !label.endsWith('-'))
}

export function isValidPhoneNumber(phoneNumber: string): boolean {
  const trimmed = phoneNumber.trim()
  if (!trimmed || !PHONE_ALLOWED_CHARS_RE.test(trimmed)) {
    return false
  }

  const digitsOnly = trimmed.replace(/\D/g, '')
  return digitsOnly.length >= 10 && digitsOnly.length <= 15
}

export function isValidExtension(extension: string): boolean {
  return EXTENSION_RE.test(extension.trim())
}

export function isValidHttpUrl(value: string): boolean {
  const trimmed = value.trim()
  if (!trimmed) {
    return false
  }

  if (trimmed.startsWith('/')) {
    return !trimmed.startsWith('//')
  }

  try {
    const url = new URL(trimmed)
    return url.protocol === 'http:' || url.protocol === 'https:'
  } catch {
    return false
  }
}
