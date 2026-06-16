import { PDFDocument, StandardFonts, rgb, type PDFPage } from 'pdf-lib'
import type {
  BookSubmission,
  BookSubmissionValue,
  BookVersionDetail,
  BookVersionField,
} from '../api/booksApi'

type RectArea = {
  x?: number
  y?: number
  width?: number
  height?: number
  background_color?: string
  font_size?: number
  line_height?: number
  text_align?: 'left' | 'center' | 'right'
}

type LayoutSettings = {
  content_mask?: RectArea
  heading_area?: RectArea
  body_area?: RectArea
  image_area?: RectArea
  section_mask?: RectArea
  section_title_area?: RectArea
}

type GenerateBookPdfOptions = {
  version: BookVersionDetail
  sourcePdfBytes: Uint8Array
  fetchImageBytes: (submission: BookSubmission) => Promise<Uint8Array | null>
}

export async function generateBookPdf({
  version,
  sourcePdfBytes,
  fetchImageBytes,
}: GenerateBookPdfOptions) {
  const layout = version.layoutSettings as LayoutSettings
  const sourceDoc = await PDFDocument.load(sourcePdfBytes)
  const outputDoc = await PDFDocument.create()
  const headingFont = await outputDoc.embedFont(StandardFonts.TimesRomanBold)
  const bodyFont = await outputDoc.embedFont(StandardFonts.TimesRoman)
  const approvedBySectionId = buildApprovedSubmissionMap(version.approvedSubmissions)

  const sourceBackedSections = version.sections.filter(
    (section) => typeof section.sourceStartPage === 'number' && typeof section.sourceEndPage === 'number',
  )
  const generatedSections = version.sections.filter(
    (section) => typeof section.sourceStartPage !== 'number' || typeof section.sourceEndPage !== 'number',
  )

  let sourceCursor = 1
  for (const section of sourceBackedSections) {
    const sectionStart = section.sourceStartPage ?? 1
    const sectionEnd = section.sourceEndPage ?? sectionStart

    if (sourceCursor <= sectionEnd) {
      const copiedPages = await outputDoc.copyPages(
        sourceDoc,
        range(
          Math.max(sourceCursor, 1) - 1,
          Math.min(sectionEnd, sourceDoc.getPageCount()) - 1,
        ),
      )
      copiedPages.forEach((page) => outputDoc.addPage(page))
      sourceCursor = sectionEnd + 1
    }

    const submissions = approvedBySectionId.get(section.id) ?? []
    for (const submission of submissions) {
      const page = await cloneTemplatePage(
        outputDoc,
        sourceDoc,
        version.contentTemplatePageNumber,
      )
      renderContentPage(page, submission, version.fields, layout, headingFont, bodyFont)
      await renderSubmissionImage(page, submission, fetchImageBytes, outputDoc, layout)
      outputDoc.addPage(page)
    }
  }

  if (sourceCursor <= sourceDoc.getPageCount()) {
    const copiedPages = await outputDoc.copyPages(
      sourceDoc,
      range(sourceCursor - 1, sourceDoc.getPageCount() - 1),
    )
    copiedPages.forEach((page) => outputDoc.addPage(page))
  }

  for (const section of generatedSections) {
    const sectionPage = await cloneTemplatePage(
      outputDoc,
      sourceDoc,
      version.sectionTemplatePageNumber,
    )
    renderSectionPage(sectionPage, section.name, layout, headingFont)
    outputDoc.addPage(sectionPage)

    const submissions = approvedBySectionId.get(section.id) ?? []
    for (const submission of submissions) {
      const page = await cloneTemplatePage(
        outputDoc,
        sourceDoc,
        version.contentTemplatePageNumber,
      )
      renderContentPage(page, submission, version.fields, layout, headingFont, bodyFont)
      await renderSubmissionImage(page, submission, fetchImageBytes, outputDoc, layout)
      outputDoc.addPage(page)
    }
  }

  return new Blob([toArrayBuffer(await outputDoc.save())], { type: 'application/pdf' })
}

function buildApprovedSubmissionMap(submissions: BookSubmission[]) {
  const grouped = new Map<number, BookSubmission[]>()
  submissions.forEach((submission) => {
    if (!submission.targetSectionId) {
      return
    }
    const current = grouped.get(submission.targetSectionId) ?? []
    current.push(submission)
    grouped.set(submission.targetSectionId, current)
  })
  return grouped
}

async function cloneTemplatePage(outputDoc: PDFDocument, sourceDoc: PDFDocument, pageNumber: number) {
  const templateIndex = Math.max(pageNumber - 1, 0)
  const [copiedPage] = await outputDoc.copyPages(sourceDoc, [templateIndex])
  return copiedPage
}

function renderContentPage(
  page: PDFPage,
  submission: BookSubmission,
  fields: BookVersionField[],
  layout: LayoutSettings,
  headingFont: Awaited<ReturnType<PDFDocument['embedFont']>>,
  bodyFont: Awaited<ReturnType<PDFDocument['embedFont']>>,
) {
  drawMask(page, layout.content_mask)

  const valuesByFieldId = new Map<number, BookSubmissionValue>(
    submission.fieldValues.map((value) => [value.fieldId, value]),
  )
  const headingText = buildSubmissionText(fields, valuesByFieldId, 'heading')
  const bodyText = buildSubmissionText(fields, valuesByFieldId, 'body')

  drawWrappedText(page, headingText, layout.heading_area, headingFont, {
    fontSize: layout.heading_area?.font_size ?? 18,
    color: rgb(0.17, 0.14, 0.1),
    lineHeight: layout.heading_area?.line_height ?? 1.2,
    textAlign: layout.heading_area?.text_align ?? 'left',
  })

  drawWrappedText(page, bodyText, layout.body_area, bodyFont, {
    fontSize: layout.body_area?.font_size ?? 11,
    color: rgb(0.18, 0.18, 0.18),
    lineHeight: layout.body_area?.line_height ?? 1.35,
    textAlign: layout.body_area?.text_align ?? 'left',
  })
}

function renderSectionPage(
  page: PDFPage,
  sectionName: string,
  layout: LayoutSettings,
  font: Awaited<ReturnType<PDFDocument['embedFont']>>,
) {
  drawMask(page, layout.section_mask)
  drawWrappedText(page, sectionName, layout.section_title_area, font, {
    fontSize: layout.section_title_area?.font_size ?? 28,
    color: rgb(0.18, 0.12, 0.09),
    lineHeight: layout.section_title_area?.line_height ?? 1.1,
    textAlign: layout.section_title_area?.text_align ?? 'center',
  })
}

async function renderSubmissionImage(
  page: PDFPage,
  submission: BookSubmission,
  fetchImageBytes: (submission: BookSubmission) => Promise<Uint8Array | null>,
  outputDoc: PDFDocument,
  layout: LayoutSettings,
) {
  if (!submission.image || !layout.image_area) {
    return
  }

  const imageBytes = await fetchImageBytes(submission)
  if (!imageBytes) {
    return
  }

  try {
    const embedded = await embedImageBytes(outputDoc, imageBytes, submission.image.mimeType)
    if (!embedded) {
      return
    }

    const area = layout.image_area
    const width = area.width ?? 110
    const height = area.height ?? 110
    const fitted = embedded.scaleToFit(width, height)
    page.drawImage(embedded, {
      x: area.x ?? 0,
      y: area.y ?? 0,
      width: fitted.width,
      height: fitted.height,
    })
  } catch {
    // Ignore image rendering failures so the PDF can still be generated.
  }
}

async function embedImageBytes(outputDoc: PDFDocument, bytes: Uint8Array, mimeType: string) {
  const normalizedMimeType = mimeType.trim().toLowerCase()
  if (normalizedMimeType === 'image/png') {
    return outputDoc.embedPng(bytes)
  }
  if (normalizedMimeType === 'image/jpeg' || normalizedMimeType === 'image/jpg') {
    return outputDoc.embedJpg(bytes)
  }

  const blob = new Blob([toArrayBuffer(bytes)], { type: mimeType || 'image/png' })
  const pngBytes = await convertImageBlobToPngBytes(blob)
  if (!pngBytes) {
    return null
  }
  return outputDoc.embedPng(pngBytes)
}

async function convertImageBlobToPngBytes(blob: Blob) {
  const objectUrl = URL.createObjectURL(blob)

  try {
    const image = await new Promise<HTMLImageElement>((resolve, reject) => {
      const element = new Image()
      element.onload = () => resolve(element)
      element.onerror = () => reject(new Error('Unable to load image'))
      element.src = objectUrl
    })

    const canvas = document.createElement('canvas')
    canvas.width = image.naturalWidth || image.width
    canvas.height = image.naturalHeight || image.height
    const context = canvas.getContext('2d')
    if (!context) {
      return null
    }
    context.drawImage(image, 0, 0)
    const dataUrl = canvas.toDataURL('image/png')
    const payload = dataUrl.split(',')[1] ?? ''
    const binary = atob(payload)
    const bytes = new Uint8Array(binary.length)
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index)
    }
    return bytes
  } finally {
    URL.revokeObjectURL(objectUrl)
  }
}

function buildSubmissionText(
  fields: BookVersionField[],
  valuesByFieldId: Map<number, BookSubmissionValue>,
  placement: 'heading' | 'body',
) {
  return fields
    .filter((field) => field.placement === placement)
    .map((field) => {
      const value = valuesByFieldId.get(field.id)?.value ?? ''
      const text = field.inputType === 'rich_text' ? htmlToPlainText(value) : value.trim()
      if (!text) {
        return ''
      }
      return field.showLabel ? `${field.label}: ${text}` : text
    })
    .filter(Boolean)
    .join(placement === 'heading' ? '\n' : '\n\n')
}

function htmlToPlainText(value: string) {
  const trimmed = value.trim()
  if (!trimmed) {
    return ''
  }
  const documentFragment = new DOMParser().parseFromString(trimmed, 'text/html')
  return collectNodeText(documentFragment.body)
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .trim()
}

function collectNodeText(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent ?? ''
  }
  if (node.nodeType !== Node.ELEMENT_NODE) {
    return ''
  }

  const element = node as HTMLElement
  if (element.tagName === 'BR') {
    return '\n'
  }

  const childText = Array.from(element.childNodes).map(collectNodeText).join('')

  if (element.tagName === 'LI') {
    return `• ${childText.trim()}\n`
  }
  if (['P', 'DIV', 'SECTION', 'ARTICLE', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6'].includes(element.tagName)) {
    return `${childText.trim()}\n\n`
  }
  if (['UL', 'OL'].includes(element.tagName)) {
    return `${childText.trim()}\n\n`
  }

  return childText
}

function drawMask(page: PDFPage, area?: RectArea) {
  if (!area || typeof area.width !== 'number' || typeof area.height !== 'number') {
    return
  }
  page.drawRectangle({
    x: area.x ?? 0,
    y: area.y ?? 0,
    width: area.width,
    height: area.height,
    color: hexToRgb(area.background_color ?? '#ffffff'),
  })
}

function drawWrappedText(
  page: PDFPage,
  text: string,
  area: RectArea | undefined,
  font: Awaited<ReturnType<PDFDocument['embedFont']>>,
  options: {
    fontSize: number
    lineHeight: number
    color: ReturnType<typeof rgb>
    textAlign: 'left' | 'center' | 'right'
  },
) {
  if (!area || !text.trim()) {
    return
  }

  const width = area.width ?? 0
  const height = area.height ?? 0
  if (width <= 0 || height <= 0) {
    return
  }

  const lines = wrapText(text, width, font, options.fontSize)
  const lineHeight = options.fontSize * options.lineHeight
  const maxLines = Math.max(Math.floor(height / lineHeight), 1)
  const trimmedLines = lines.slice(0, maxLines)

  trimmedLines.forEach((line, index) => {
    const textWidth = font.widthOfTextAtSize(line, options.fontSize)
    const x =
      options.textAlign === 'center'
        ? (area.x ?? 0) + Math.max((width - textWidth) / 2, 0)
        : options.textAlign === 'right'
          ? (area.x ?? 0) + Math.max(width - textWidth, 0)
          : area.x ?? 0

    page.drawText(line, {
      x,
      y: (area.y ?? 0) + height - options.fontSize - index * lineHeight,
      size: options.fontSize,
      font,
      color: options.color,
      maxWidth: width,
      lineHeight,
    })
  })
}

function wrapText(
  text: string,
  maxWidth: number,
  font: Awaited<ReturnType<PDFDocument['embedFont']>>,
  fontSize: number,
) {
  const paragraphs = text.split(/\n+/)
  const lines: string[] = []

  paragraphs.forEach((paragraph, paragraphIndex) => {
    const words = paragraph.split(/\s+/).filter(Boolean)
    if (!words.length) {
      if (paragraphIndex < paragraphs.length - 1) {
        lines.push('')
      }
      return
    }

    let currentLine = ''
    words.forEach((word) => {
      const candidate = currentLine ? `${currentLine} ${word}` : word
      const candidateWidth = font.widthOfTextAtSize(candidate, fontSize)
      if (candidateWidth <= maxWidth || !currentLine) {
        currentLine = candidate
      } else {
        lines.push(currentLine)
        currentLine = word
      }
    })

    if (currentLine) {
      lines.push(currentLine)
    }

    if (paragraphIndex < paragraphs.length - 1) {
      lines.push('')
    }
  })

  return lines
}

function hexToRgb(value: string) {
  const normalized = value.trim().replace('#', '')
  const full = normalized.length === 3
    ? normalized
        .split('')
        .map((char) => `${char}${char}`)
        .join('')
    : normalized.padEnd(6, '0')

  const red = Number.parseInt(full.slice(0, 2), 16) / 255
  const green = Number.parseInt(full.slice(2, 4), 16) / 255
  const blue = Number.parseInt(full.slice(4, 6), 16) / 255
  return rgb(red, green, blue)
}

function range(start: number, end: number) {
  if (end < start) {
    return []
  }
  return Array.from({ length: end - start + 1 }, (_, index) => start + index)
}

function toArrayBuffer(bytes: Uint8Array) {
  return Uint8Array.from(bytes).buffer
}
