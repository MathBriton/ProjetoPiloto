import { useState } from 'react'
import { Outlet } from 'react-router-dom'
import { Sidebar } from './Sidebar'
import { TopBar } from './TopBar'
import { PlantaProvider } from '@/context/PlantaContext'

export function AppShell() {
  const [sidebarExpandida, setSidebarExpandida] = useState(true)

  return (
    <PlantaProvider>
      <div className="flex h-screen bg-slate-50 dark:bg-slate-900">
        <Sidebar expandida={sidebarExpandida} aoAlternar={() => setSidebarExpandida((v) => !v)} />
        <div className="flex flex-1 flex-col overflow-hidden">
          <TopBar />
          <main className="flex-1 overflow-y-auto p-6">
            <Outlet />
          </main>
        </div>
      </div>
    </PlantaProvider>
  )
}
