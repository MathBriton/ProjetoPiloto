import { createContext, useContext, useEffect, useState, type ReactNode } from 'react'

type Tema = 'light' | 'dark'

interface ThemeContextValor {
  tema: Tema
  alternarTema: () => void
}

const ThemeContext = createContext<ThemeContextValor | undefined>(undefined)

const CHAVE_STORAGE = 'swnet:tema'

function temaInicial(): Tema {
  const salvo = localStorage.getItem(CHAVE_STORAGE) as Tema | null
  if (salvo) return salvo
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
}

export function ThemeProvider({ children }: { children: ReactNode }) {
  const [tema, setTema] = useState<Tema>(temaInicial)

  useEffect(() => {
    document.documentElement.classList.toggle('dark', tema === 'dark')
    localStorage.setItem(CHAVE_STORAGE, tema)
  }, [tema])

  const alternarTema = () => setTema((atual) => (atual === 'dark' ? 'light' : 'dark'))

  return <ThemeContext.Provider value={{ tema, alternarTema }}>{children}</ThemeContext.Provider>
}

export function useTheme() {
  const contexto = useContext(ThemeContext)
  if (!contexto) throw new Error('useTheme precisa estar dentro de <ThemeProvider>')
  return contexto
}
