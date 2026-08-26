import { useAuth } from '@/context/AuthContext'
import { usePlanta } from '@/context/PlantaContext'

export function Home() {
  const { usuario } = useAuth()
  const { plantaAtual } = usePlanta()

  return (
    <div className="space-y-2">
      <h1 className="text-2xl font-bold text-slate-800 dark:text-white">
        Olá, {usuario?.nomeCompleto.split(' ')[0]} 👋
      </h1>
      <p className="text-slate-500 dark:text-slate-400">
        Você está visualizando a planta <strong>{plantaAtual?.nome ?? '—'}</strong>.
      </p>
    </div>
  )
}
