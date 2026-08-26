import {
  Ban,
  BadgePercent,
  Boxes,
  Building2,
  ClipboardCheck,
  ClipboardList,
  Coins,
  Contact,
  FileBarChart,
  FileSearch,
  FileSpreadsheet,
  Files,
  Factory,
  Headset,
  History,
  KeyRound,
  Landmark,
  LayoutDashboard,
  ListChecks,
  ListTree,
  Package,
  Power,
  Recycle,
  ScrollText,
  Settings2,
  ShieldCheck,
  ShoppingCart,
  Sliders,
  Truck,
  UserCog,
  UserPlus,
  UserRoundPlus,
  Users,
  Wrench,
  type LucideIcon
} from 'lucide-react'

/**
 * Registro único dos menus do sidebar.
 *
 * A Sidebar e as rotas do App leem daqui, então adicionar uma tela é mexer
 * em um lugar só. Cada grupo espelha uma pasta de
 * `Projetos-A-Refatorar/Menu-Serviços-Sidebar` e cada item, uma subpasta.
 *
 * `pendente: true` marca as telas ainda não refatoradas — a rota existe e
 * abre um aviso de pendência apontando a pasta de origem.
 */

export interface ItemMenu {
  rotulo: string
  caminho: string
  icone: LucideIcon
  permissaoNecessaria?: string
  pendente?: boolean
  /** Pasta de origem no material legado, para rastreabilidade */
  origem?: string
}

export interface GrupoMenu {
  rotulo: string
  icone: LucideIcon
  filhos: ItemMenu[]
}

export type EntradaMenu = ItemMenu | GrupoMenu

export function ehGrupo(entrada: EntradaMenu): entrada is GrupoMenu {
  return 'filhos' in entrada
}

/** Atalho para declarar as telas ainda não refatoradas */
function pendente(rotulo: string, caminho: string, icone: LucideIcon, origem: string): ItemMenu {
  return { rotulo, caminho, icone, pendente: true, origem }
}

export const ITENS_MENU: EntradaMenu[] = [
  { rotulo: 'Início', caminho: '/', icone: LayoutDashboard },

  {
    rotulo: 'Administrador',
    icone: ShieldCheck,
    filhos: [
      {
        rotulo: 'Acessos Usuários',
        caminho: '/administrador/acessos-usuarios',
        icone: Users,
        permissaoNecessaria: 'acessos-usuarios.visualizar',
        origem: 'Administrador/Acessos-Usuarios-SWNET'
      },
      {
        rotulo: 'Alterar Usuário',
        caminho: '/administrador/alterar-usuario',
        icone: UserCog,
        permissaoNecessaria: 'usuarios.alterar',
        origem: 'Administrador/Alterar-Usuario-Swnet'
      },
      {
        rotulo: 'Ativar / Desativar',
        caminho: '/administrador/ativar-desativar',
        icone: Power,
        permissaoNecessaria: 'usuarios.ativar-desativar',
        origem: 'Administrador/Ativar-Desativar-Usuarios'
      },
      {
        rotulo: 'Cadastro Usuário',
        caminho: '/administrador/cadastro-usuario',
        icone: UserPlus,
        permissaoNecessaria: 'usuarios.incluir',
        origem: 'Administrador/Cadastro-Usuario-SWNET'
      },
      {
        rotulo: 'Cancelar Usuários',
        caminho: '/administrador/cancelar-usuarios',
        icone: Ban,
        permissaoNecessaria: 'usuarios.cancelar',
        origem: 'Administrador/Cancelar-Usuarios'
      },
      {
        rotulo: 'Criação Usuário',
        caminho: '/administrador/criacao-usuario-ad',
        icone: UserRoundPlus,
        permissaoNecessaria: 'usuarios.criacao-ad',
        origem: 'Administrador/Criacao-Usuarios-AD'
      },
      {
        rotulo: 'Gerenciar Menus',
        caminho: '/administrador/gerenciar-menus',
        icone: ListTree,
        permissaoNecessaria: 'menus.gerenciar',
        origem: 'Administrador/Gerenciar-Menus'
      },
      {
        rotulo: 'Permissões Especiais',
        caminho: '/administrador/permissoes-especiais',
        icone: KeyRound,
        permissaoNecessaria: 'permissoes-especiais.editar',
        origem: 'Administrador/Permissoes-Especiais'
      },
      {
        rotulo: 'Relatório de Permissões',
        caminho: '/administrador/relatorio-permissoes',
        icone: ScrollText,
        permissaoNecessaria: 'permissoes.relatorio',
        origem: 'Administrador/Relatorio-De-Permissoes-De-Usuarios'
      },
      {
        rotulo: 'Revisão de Usuário',
        caminho: '/administrador/revisao-usuarios',
        icone: History,
        permissaoNecessaria: 'usuarios.revisao',
        origem: 'Administrador/Revisao-De-Usuarios'
      }
    ]
  },

  {
    rotulo: 'Cadastros FNC',
    icone: ClipboardList,
    filhos: [
      {
        rotulo: 'Aprovador FNC',
        caminho: '/cadastros-fnc/aprovadores',
        icone: ClipboardCheck,
        origem: 'Cadastros-FNC/Cadastro-Aprovador-FNC'
      },
      {
        rotulo: 'Defeitos de FNC',
        caminho: '/cadastros-fnc/defeitos',
        icone: Wrench,
        origem: 'Cadastros-FNC/Cadastro-Defeitos-De-FNC'
      },
      {
        rotulo: 'Responsável FNC',
        caminho: '/cadastros-fnc/responsaveis',
        icone: Contact,
        origem: 'Cadastros-FNC/Cadastro-Responsavel-FNC'
      },
      {
        rotulo: 'Setores de FNC',
        caminho: '/cadastros-fnc/setores',
        icone: Building2,
        origem: 'Cadastros-FNC/Cadastro-Setores-De-FNC'
      }
    ]
  },

  {
    rotulo: 'Cadastros Orçamento',
    icone: Coins,
    filhos: [
      {
        rotulo: 'Analistas de Orçamento',
        caminho: '/cadastros-orcamento/analistas',
        icone: Users,
        origem: 'Cadastros-Orçamento/Cadastro-Analistas-De-Orcamento'
      },
      {
        rotulo: 'Aprovadores de Orçamento',
        caminho: '/cadastros-orcamento/aprovadores',
        icone: ClipboardCheck,
        origem: 'Cadastros-Orçamento/Cadastro-Aprovadores-De-Orcamento'
      },
      {
        rotulo: 'Faixas de Aprovador',
        caminho: '/cadastros-orcamento/faixas-aprovador',
        icone: BadgePercent,
        origem: 'Cadastros-Orçamento/Faixas-De-Aprovador-De-Orcamento'
      }
    ]
  },

  {
    rotulo: 'Cavaco / Aparas',
    icone: Recycle,
    filhos: [
      {
        rotulo: 'Fornecedor de Aparas',
        caminho: '/cavaco-aparas/fornecedores',
        icone: Factory,
        origem: 'Cavaco-Aparas/Cadastro-De-Fornecedor-Aparas'
      },
      {
        rotulo: 'Controle de Recebimento',
        caminho: '/cavaco-aparas/controle-recebimento',
        icone: ClipboardList,
        origem: 'Cavaco-Aparas/Controle-De-Recebimento'
      },
      {
        rotulo: 'Relatório de Recebimento',
        caminho: '/cavaco-aparas/relatorio-recebimento',
        icone: FileSpreadsheet,
        origem: 'Cavaco-Aparas/Controle-De-Recebimento'
      }
    ]
  },

  {
    rotulo: 'Cliente',
    icone: Contact,
    filhos: [
      pendente('Alterar Dados do Cliente', '/cliente/alterar-dados', UserCog, 'Cliente/Alterar-Dados-Do-Cliente'),
      pendente('Alterar Status do Cliente', '/cliente/alterar-status', Power, 'Cliente/Alterar-Status-Do-Cliente'),
      pendente('Cadastro Unificado V2', '/cliente/cadastro-unificado-v2', Files, 'Cliente/Cadastro-Unificado-V2'),
      pendente(
        'Transferência / Extensão',
        '/cliente/transferencia-extensao',
        Truck,
        'Cliente/Transferencia-Extensao-De-Cliente'
      )
    ]
  },

  {
    rotulo: 'Compras',
    icone: ShoppingCart,
    filhos: [
      {
        rotulo: 'Classificar Notas',
        caminho: '/compras/classificar-notas',
        icone: FileSearch,
        origem: 'Compras/Classificar-Notas'
      }
    ]
  },

  {
    rotulo: 'Desenvolvimento',
    icone: Package,
    filhos: [
      {
        rotulo: 'Alteração de Produtos',
        caminho: '/desenvolvimento/alteracao-produtos',
        icone: Wrench,
        origem: 'Desenvolvimento/Alteracao-De-Produtos'
      },
      {
        rotulo: 'Aprovação Comercial SPD',
        caminho: '/desenvolvimento/aprovacao-comercial-spd',
        icone: ClipboardCheck,
        origem: 'Desenvolvimento/Aprovacao-Comercial-SPD'
      },
      {
        rotulo: 'Lote Piloto V2',
        caminho: '/desenvolvimento/lote-piloto-v2',
        icone: Boxes,
        origem: 'Desenvolvimento/Lote-Piloto-V2'
      },
      {
        rotulo: 'Solicitação de SPD V2',
        caminho: '/desenvolvimento/solicitacao-spd-v2',
        icone: ClipboardList,
        origem: 'Desenvolvimento/Solicitacao-De-SPD-V2'
      }
    ]
  },

  {
    rotulo: 'Expedição',
    icone: Truck,
    filhos: [
      {
        rotulo: 'Checklist Transporte',
        caminho: '/expedicao/checklist-transporte',
        icone: ListChecks,
        origem: 'Expedicao/Checklist-Transporte'
      },
      {
        rotulo: 'Simulador de Cargas',
        caminho: '/expedicao/simulador-cargas',
        icone: Boxes,
        origem: 'Expedicao/Simulador-De-Cargas'
      }
    ]
  },

  {
    rotulo: 'Financeiro',
    icone: Landmark,
    filhos: [
      {
        rotulo: 'Alterar Risco',
        caminho: '/financeiro/alterar-risco',
        icone: Sliders,
        origem: 'Financeiro/Alterar-Risco'
      },
      {
        rotulo: 'Cadastro Canhoto',
        caminho: '/financeiro/cadastro-canhoto',
        icone: Files,
        origem: 'Financeiro/Cadastro-Canhoto'
      },
      {
        rotulo: 'Validadores de Crédito',
        caminho: '/financeiro/validadores-credito',
        icone: ClipboardCheck,
        origem: 'Financeiro/Cadastro-Validadores-De-Credito'
      },
      {
        rotulo: 'Solicitar Aumento de Crédito',
        caminho: '/financeiro/aumento-credito',
        icone: Coins,
        origem: 'Financeiro/Solicitar-Aumento-De-Credito'
      },
      {
        rotulo: 'Transferência de Crédito',
        caminho: '/financeiro/transferencia-credito',
        icone: Coins,
        origem: 'Financeiro/Solicitar-Transferencia-De-Credito'
      }
    ]
  },

  {
    rotulo: 'Pedidos',
    icone: ClipboardList,
    filhos: [
      pendente('Alterar Pedidos Brasil', '/pedidos/alterar-brasil', UserCog, 'Pedidos/Alterar-Pedidos-Brasil'),
      pendente('Alterar Pedidos PiraUbe', '/pedidos/alterar-piraube', UserCog, 'Pedidos/Alterar-Pedidos-PiraUbe'),
      pendente(
        'Entrega no Fim de Semana',
        '/pedidos/entrega-fim-de-semana',
        Truck,
        'Pedidos/Definir-Entrega-No-Fim-De-Semana'
      ),
      pendente('Inclusão OBS NF', '/pedidos/inclusao-obs-nf', Files, 'Pedidos/Inclusao-OBS-NF'),
      pendente('Orçamento Brasil', '/pedidos/orcamento-brasil', Coins, 'Pedidos/Orçamento-Brasil'),
      pendente('Parâmetros Pedidos', '/pedidos/parametros', Settings2, 'Pedidos/Parametros-Pedidos'),
      pendente('Pedidos Brasil V2', '/pedidos/brasil-v2', ClipboardList, 'Pedidos/Pedidos-Brasil-v2'),
      pendente(
        'Pedidos Digitados Brasil',
        '/pedidos/digitados-brasil',
        ClipboardList,
        'Pedidos/Pedidos-Digitados-Brasil'
      ),
      pendente(
        'Pedidos Digitados PiraUbe',
        '/pedidos/digitados-piraube',
        ClipboardList,
        'Pedidos/Pedidos-Digitados-PiraUbe'
      ),
      pendente('Pedidos PiraUbe', '/pedidos/piraube', ClipboardList, 'Pedidos/PedidosPiraUbe')
    ]
  },

  {
    rotulo: 'RH',
    icone: Headset,
    filhos: [
      pendente('Atendentes RH', '/rh/atendentes', Users, 'RH/Cadastro-Atendentes-RH'),
      pendente('Categorias de Chamado', '/rh/categorias-chamado', ListTree, 'RH/Categorias-De-Chamado'),
      pendente('Chamados RH', '/rh/chamados', Headset, 'RH/Chamados-RH'),
      pendente('Departamentos de Chamado', '/rh/departamentos-chamado', Building2, 'RH/Departamentos-De-Chamado-RH'),
      pendente('Importar do ADP', '/rh/importar-adp', FileSpreadsheet, 'RH/Importar-Do-ADP'),
      pendente('Meus Chamados RH', '/rh/meus-chamados', ClipboardList, 'RH/Meus-Chamados-RH'),
      pendente('Novo Chamado RH', '/rh/novo-chamado', UserPlus, 'RH/Novo-Chamado-RH'),
      pendente('Prioridades de Chamado', '/rh/prioridades-chamado', Sliders, 'RH/Prioridades-De-Chamado-RH')
    ]
  },

  {
    rotulo: 'SGQ',
    icone: ClipboardCheck,
    filhos: [
      { rotulo: 'FNC Brasil', caminho: '/sgq/fnc-brasil', icone: ClipboardList, origem: 'SGQ/FNC-Brasil' },
      { rotulo: 'Ficha de QM', caminho: '/sgq/ficha-qm', icone: Files, origem: 'SGQ/Ficha-De-QM' }
    ]
  },

  {
    rotulo: 'Utilitários',
    icone: Settings2,
    filhos: [
      {
        rotulo: 'Alterar Senha',
        caminho: '/utilitarios/alterar-senha',
        icone: KeyRound,
        origem: 'Utilitarios/Alterar-Senha'
      }
    ]
  },

  {
    rotulo: 'Validações',
    icone: ClipboardCheck,
    filhos: [
      { rotulo: 'Papel', caminho: '/validacoes/papel', icone: Files, origem: 'Validacoes/Papel' },
      {
        rotulo: 'Solicitações de Crédito',
        caminho: '/validacoes/solicitacoes-credito',
        icone: Coins,
        origem: 'Validacoes/Solicitacoes-De-Credito'
      },
      {
        rotulo: 'Validar Pedido',
        caminho: '/validacoes/validar-pedido',
        icone: ClipboardCheck,
        origem: 'Validacoes/Validar-Pedido'
      },
      {
        rotulo: 'Validar Pedido Brasil',
        caminho: '/validacoes/validar-pedido-brasil',
        icone: ClipboardCheck,
        origem: 'Validacoes/Validar-Pedido-Brasil'
      }
    ]
  },

  { rotulo: 'Relatórios', caminho: '/relatorios', icone: FileBarChart }
]

/** Todos os itens em lista plana — usado para gerar as rotas */
export const ITENS_PLANOS: ItemMenu[] = ITENS_MENU.flatMap((entrada) =>
  ehGrupo(entrada) ? entrada.filhos : [entrada]
)

/** Só as telas ainda sem fonte ASP, que abrem o aviso de pendência */
export const ITENS_PENDENTES = ITENS_PLANOS.filter((item) => item.pendente)
