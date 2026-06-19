import { useEffect, useState } from 'react'
import toast from 'react-hot-toast'
import { useNavigate, useParams } from 'react-router-dom'
import {
  knowledgeCenterApi,
  type KnowledgeCenterSubmission,
} from '../api/knowledgeCenterApi'
import { Breadcrumb } from '../components/Breadcrumb'
import { CmsAppShell } from '../components/CmsAppShell'
import { Loader } from '../components/Loader'
import styles from '../styles/BookRequestReviewPage.module.css'

export function KnowledgeCenterReviewPage() {
  const navigate = useNavigate()
  const { id } = useParams()
  const submissionID = id ? Number.parseInt(id, 10) : Number.NaN

  const [submission, setSubmission] = useState<KnowledgeCenterSubmission | null>(
    null,
  )
  const [completionNotes, setCompletionNotes] = useState('')
  const [isLoading, setIsLoading] = useState(true)
  const [isWorking, setIsWorking] = useState(false)

  useEffect(() => {
    if (Number.isNaN(submissionID)) {
      setIsLoading(false)
      return
    }

    void loadSubmission(submissionID)
  }, [submissionID])

  async function loadSubmission(targetID: number) {
    try {
      setIsLoading(true)
      const detail = await knowledgeCenterApi.getSubmission(targetID)
      setSubmission(detail)
      setCompletionNotes(detail.completionNotes)
    } catch (error) {
      toast.error(
        error instanceof Error
          ? error.message
          : 'Unable to load the knowledge center request.',
      )
    } finally {
      setIsLoading(false)
    }
  }

  async function handleMarkCompleted() {
    if (!submission) {
      return
    }
    if (!completionNotes.trim()) {
      toast.error('Completion notes are required.')
      return
    }

    try {
      setIsWorking(true)
      await knowledgeCenterApi.completeSubmission(
        submission.id,
        completionNotes.trim(),
      )
      toast.success('Request marked as completed.')
      navigate('/knowledge-center?tab=completed')
    } catch (error) {
      toast.error(
        error instanceof Error
          ? error.message
          : 'Unable to complete the request.',
      )
    } finally {
      setIsWorking(false)
    }
  }

  const isCompleted = submission?.status === 'completed'

  return (
    <CmsAppShell activeKey="knowledgeCenter">
      <div className={styles.page}>
        <Breadcrumb
          items={[
            { label: 'Knowledge Center', to: '/knowledge-center' },
            { label: `Request #${submission?.id ?? submissionID}` },
          ]}
        />

        <header className={styles.header}>
          <div>
            <h1 className={styles.title}>
              {submission
                ? `Request #${submission.id}`
                : 'Knowledge Center Request'}
            </h1>
            <p className={styles.subtitle}>
              Review the public contribution details from the Living History Hub
              and capture completion notes once the request has been handled.
            </p>
          </div>

          <div className={styles.headerActions}>
            <button
              type="button"
              className={styles.secondaryButton}
              onClick={() =>
                navigate(
                  submission?.status === 'completed'
                    ? '/knowledge-center?tab=completed'
                    : '/knowledge-center',
                )
              }
            >
              Back to Requests
            </button>
          </div>
        </header>

        {isLoading ? (
          <div className={styles.loaderWrap}>
            <Loader />
          </div>
        ) : !submission ? (
          <section className={styles.errorPanel}>
            <h2>Request not found</h2>
            <p>The selected knowledge center request could not be loaded.</p>
          </section>
        ) : (
          <>
            <section className={styles.summaryGrid}>
              <article className={styles.summaryCard}>
                <span className={styles.summaryLabel}>Submitter</span>
                <strong>{submission.submitterName}</strong>
                <p>{submission.submitterEmail}</p>
              </article>

              <article className={styles.summaryCard}>
                <span className={styles.summaryLabel}>Contribution Type</span>
                <strong>{formatSubmissionType(submission.submissionType)}</strong>
                <p>{submission.submitterPhone || 'No phone number provided'}</p>
              </article>

              <article className={styles.summaryCard}>
                <span className={styles.summaryLabel}>Timeline</span>
                <strong>{formatDateTime(submission.createdAt)}</strong>
                <p>
                  {isCompleted && submission.completedAt
                    ? `Completed ${formatDateTime(submission.completedAt)}`
                    : 'Still open for review'}
                </p>
              </article>
            </section>

            <section className={styles.card}>
              <div className={styles.cardHeader}>
                <div>
                  <h2>Submission Details</h2>
                  <p>
                    This is exactly what the visitor shared from the Living
                    History Hub contribution form.
                  </p>
                </div>

                <span className={styles.statusBadge}>
                  {isCompleted ? 'Completed' : 'Open'}
                </span>
              </div>

              <div className={styles.metaRow}>
                <span className={styles.metaPill}>
                  {formatSubmissionType(submission.submissionType)}
                </span>
                <span className={styles.metaPill}>
                  {submission.submitterEmail}
                </span>
                {submission.submitterPhone ? (
                  <span className={styles.metaTag}>
                    {submission.submitterPhone}
                  </span>
                ) : null}
              </div>

              <div className={styles.fieldStack}>
                <div className={styles.requestField}>
                  <label className={styles.field}>
                    <span>Request Message</span>
                    <textarea
                      rows={10}
                      readOnly
                      value={submission.message}
                    />
                  </label>
                </div>
              </div>
            </section>

            <section className={styles.actionCard}>
              <div className={styles.cardHeader}>
                <div>
                  <h2>
                    {isCompleted ? 'Completion Notes' : 'Mark as Completed'}
                  </h2>
                  <p>
                    {isCompleted
                      ? 'This request has already been completed and the reviewer trail is stored below.'
                      : 'Add internal notes about what was done before moving this request into the completed tab.'}
                  </p>
                </div>
              </div>

              <label className={styles.field}>
                <span>Completion Notes</span>
                <textarea
                  rows={6}
                  value={completionNotes}
                  readOnly={isCompleted}
                  onChange={(event) => setCompletionNotes(event.target.value)}
                />
              </label>

              {submission.completedBy ? (
                <div className={styles.metaRow}>
                  <span className={styles.metaPill}>
                    {submission.completedBy.name}
                  </span>
                  <span className={styles.metaPill}>
                    {submission.completedBy.email}
                  </span>
                  {submission.completedAt ? (
                    <span className={styles.metaTag}>
                      {formatDateTime(submission.completedAt)}
                    </span>
                  ) : null}
                </div>
              ) : null}

              <div className={styles.actionRow}>
                <button
                  type="button"
                  className={styles.secondaryButton}
                  onClick={() =>
                    navigate(
                      isCompleted
                        ? '/knowledge-center?tab=completed'
                        : '/knowledge-center',
                    )
                  }
                >
                  Back to List
                </button>
                {!isCompleted ? (
                  <button
                    type="button"
                    className={styles.primaryButton}
                    disabled={isWorking || !completionNotes.trim()}
                    onClick={() => void handleMarkCompleted()}
                  >
                    {isWorking ? 'Saving...' : 'Mark Completed'}
                  </button>
                ) : null}
              </div>
            </section>
          </>
        )}
      </div>
    </CmsAppShell>
  )
}

function formatSubmissionType(value: KnowledgeCenterSubmission['submissionType']) {
  switch (value) {
    case 'post':
      return 'Story'
    case 'video':
      return 'Video'
    case 'both':
      return 'Story + Video'
    default:
      return value
  }
}

function formatDateTime(value: string) {
  const parsed = Date.parse(value)
  if (Number.isNaN(parsed)) {
    return value
  }

  return new Intl.DateTimeFormat('en-CA', {
    month: 'short',
    day: '2-digit',
    year: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
  }).format(new Date(parsed))
}
