import { createContext, useContext, useState, useEffect, type ReactNode } from 'react'
import { useAuth } from './AuthContext'
import type { Planta } from '@/types'

interface PlantaContextValor {
  plantaAtual: Planta | null
  plantas: Planta[]
  trocarPlanta: (plantaId: number) => void
}

const PlantaContext = createContext<PlantaContextValor | undefined>(undefined)

const CHAVE_STORAGE = 'swnet:plantaAtualId'

export function PlantaProvider({ children }: { children: ReactNode }) {
  const { usuario } = useAuth()
  const plantas = usuario?.plantas ?? []
  const [plantaAtualId, setPlantaAtualId] = useState<number | null>(null)

  useEffect(() => {
    if (!usuario) {
      setPlantaAtualId(null)
      return
    }
    const salva = Number(localStorage.getItem(CHAVE_STORAGE))
    const existe = usuario.plantas.some((p) => p.id === salva)
    setPlantaAtualId(existe ? salva : usuario.plantaAtualId)
  }, [usuario])

  const trocarPlanta = (plantaId: number) => {
    setPlantaAtualId(plantaId)
    localStorage.setItem(CHAVE_STORAGE, String(plantaId))
    // TODO: quando o backend real estiver de pé, revalidar token/contexto de dados na troca de planta
  }

  const plantaAtual = plantas.find((p) => p.id === plantaAtualId) ?? null

  return (
    <PlantaContext.Provider value={{ plantaAtual, plantas, trocarPlanta }}>
      {children}
    </PlantaContext.Provider>
  )
}

export function usePlanta() {
  const contexto = useContext(PlantaContext)
  if (!contexto) throw new Error('usePlanta precisa estar dentro de <PlantaProvider>')
  return contexto
}
