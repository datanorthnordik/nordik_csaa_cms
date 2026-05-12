type MultipartFileField = {
  fieldName: string
  file?: Blob | null
  fileName?: string
}

export function buildMultipartPayload(
  payload: unknown,
  files: MultipartFileField[],
) {
  const form = new FormData()
  form.append('payload', JSON.stringify(payload))

  for (const entry of files) {
    if (!entry.file) {
      continue
    }

    if (entry.fileName) {
      form.append(entry.fieldName, entry.file, entry.fileName)
      continue
    }

    form.append(entry.fieldName, entry.file)
  }

  return form
}
