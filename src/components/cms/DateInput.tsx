import type { InputHTMLAttributes } from 'react'

type DateInputProps = Omit<InputHTMLAttributes<HTMLInputElement>, 'type' | 'onChange'> & {
  value: string
  onChange: (value: string) => void
}

export function DateInput({ value, onChange, max = '9999-12-31', ...rest }: DateInputProps) {
  function handleChange(e: React.ChangeEvent<HTMLInputElement>) {
    const raw = e.target.value
    if (!raw) {
      onChange(raw)
      return
    }
    const dashIdx = raw.indexOf('-')
    if (dashIdx > 4) {
      onChange('9999' + raw.slice(dashIdx))
      return
    }
    onChange(raw)
  }

  return (
    <input
      type="date"
      value={value}
      max={max}
      onChange={handleChange}
      {...rest}
    />
  )
}
