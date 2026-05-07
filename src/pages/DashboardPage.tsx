import { CmsAppShell } from '../components/CmsAppShell'

export function DashboardPage() {
  return (
    <CmsAppShell
      activeKey="dashboard"
      onSubmitTicket={() => console.log('submit ticket')}
      onLogout={() => console.log('logout')}
    >
      {/* Empty white demo page for previewing the header and drawer */}
    </CmsAppShell>
  )
}
