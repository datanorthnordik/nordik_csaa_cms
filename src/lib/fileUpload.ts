export function fileToBase64(file: File) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onerror = () => reject(new Error('Could not read file'))
    reader.onload = () => {
      const result = reader.result
      if (typeof result !== 'string') {
        reject(new Error('Could not read file'))
        return
      }

      const [, base64 = ''] = result.split(',')
      resolve(base64)
    }

    reader.readAsDataURL(file)
  })
}
