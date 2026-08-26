<!-- V 1.02 -->
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%
session.lcid = 1046
Response.CodePage = 65001
Response.CharSet = "utf-8"
'Response.Write "Serviço interrompido temporariamente."
'Response.End
%>
<!DOCTYPE html>
<html lang="pt-BR">
    <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Estruturas do Papelão Ondulado — Pirapetinga</title>

        <!-- Bootstrap CSS -->
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
              rel="stylesheet"
              integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
              crossorigin="anonymous" />

        <!-- Bootstrap Icons -->
        <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
              rel="stylesheet" />

        <!--#include file="../../../../Funcoes/funcoes.asp" -->
        <!--#include file="../../../../includes/variaveis.asp" -->
        <!--#include file="../../../../cnn/cnn.asp" -->
    <%
     'DADOSADV
    'call ValidaPermissaoRelatorios("sfa/pedidos/consultas/estruturas",SESSION("CODUSER"),SESSION("FILIAL"))
    %>

        <style>
            :root {
                /* Neutros */
                --color-bg:            hsl(240, 9%, 98%);
                --color-surface:       hsl(0, 0%, 100%);
                --color-surface-muted: hsl(240, 9%, 97%);
                --color-border:        hsl(240, 6%, 90%);
                --color-border-strong: hsl(240, 6%, 82%);
                --color-text:          hsl(240, 6%, 10%);
                --color-text-muted:    hsl(240, 4%, 46%);
                --color-text-faint:    hsl(240, 4%, 64%);

                /* Accent único */
                --color-accent:        hsl(210, 100%, 45%);
                --color-accent-hover:  hsl(210, 100%, 40%);
                --color-accent-soft:   hsl(210, 100%, 96%);

                /* Tipografia */
                --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;
                --font-mono: "SF Mono", "Cascadia Code", Consolas, monospace;

                /* Raio */
                --radius-sm: 6px;
                --radius-md: 10px;
                --radius-lg: 14px;

                /* Sombra */
                --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
                --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);

                /* Transição */
                --transition: 160ms ease;
            }

            body {
                background-color: var(--color-bg);
                color: var(--color-text);
                font-family: var(--font-sans);
                min-height: 100vh;
                -webkit-font-smoothing: antialiased;
            }

            .page-title {
                font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
                font-weight: 600;
                letter-spacing: -0.01em;
                color: var(--color-text);
            }

            /* Card de seção */
            .section-card {
                background-color: var(--color-surface);
                border: 1px solid var(--color-border);
                border-radius: var(--radius-md);
                box-shadow: var(--shadow-sm);
                overflow: hidden;
                height: 100%;
            }
            .section-card-header {
                display: flex;
                align-items: center;
                gap: .5rem;
                padding: .7rem 1.25rem;
                border-bottom: 1px solid var(--color-border);
                font-weight: 600;
                font-size: .85rem;
                color: var(--color-text);
            }
            .section-card-body { padding: 1.1rem 1.25rem 1.25rem; }

            .table-wrapper { overflow-x: auto; }

            /* Tabela — quiet UI */
            .table { margin-bottom: 0; }
            .table thead th {
                background-color: var(--color-surface-muted);
                color: var(--color-text-muted);
                font-size: .76rem;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: .03em;
                text-align: center;
                white-space: nowrap;
                vertical-align: middle;
                border-color: var(--color-border);
                padding: .6rem .75rem;
            }
            .table tbody td,
            .table tbody th {
                font-size: .85rem;
                text-align: center;
                color: var(--color-text);
                vertical-align: middle;
                border-color: var(--color-border);
                padding: .55rem .75rem;
            }
            .table tbody tr:hover td,
            .table tbody tr:hover th {
                background-color: var(--color-surface-muted);
            }

            /* Célula-rótulo (substitui o antigo fundo escuro rgba(50,50,50,.75)) */
            .cell-label {
                background-color: var(--color-surface-muted);
                font-weight: 600;
                color: var(--color-text);
                white-space: nowrap;
            }
            .cell-code { font-family: var(--font-mono); font-weight: 600; }

            .subhead {
                font-size: .78rem;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: .03em;
                color: var(--color-text-muted);
                margin: 0 0 .6rem;
            }

            /* Nota / callout */
            .note {
                background-color: var(--color-surface-muted);
                border: 1px solid var(--color-border);
                border-left: 3px solid var(--color-accent);
                border-radius: var(--radius-sm);
                padding: .8rem 1rem;
                font-size: .82rem;
                color: var(--color-text-muted);
                text-align: justify;
            }
            .note p { margin-bottom: .6rem; }
            .note p:last-child { margin-bottom: 0; }
            .note strong { color: var(--color-text); }

            /* Lista de medidas auto-conjugáveis */
            .measure-list { display: flex; flex-wrap: wrap; gap: .5rem; }
            .measure-chip {
                font-size: .82rem;
                font-family: var(--font-mono);
                padding: .3rem .7rem;
                background-color: var(--color-surface-muted);
                border: 1px solid var(--color-border);
                border-radius: var(--radius-sm);
                color: var(--color-text);
                white-space: nowrap;
            }

            @media (prefers-reduced-motion: reduce) {
                *, *::before, *::after {
                    animation-duration: 0.01ms !important;
                    transition-duration: 0.01ms !important;
                }
            }
        </style>
    </head>

    <body>

    <%
    if 1 <> 1 then
    %>
        <div class="container-fluid py-5">
            <div class="text-center">
                <span class="page-title d-block mb-2">Tabela de Estrutura do Papelão Ondulado</span>
                <p class="text-muted mb-0">Temporariamente indisponível!</p>
            </div>
        </div>
    <%
    else

        'BUSCANDO A REVISÃO DA TABELA
            SQL = "SELECT "
        SQL = SQL & "   Z1_ESTRUT AS ESTRUTURA, "
        SQL = SQL & "   Z1_STATUS, "
        SQL = SQL & "   (Z1_GESTPIR) AS GRAMATPIR, "
        SQL = SQL & "   (Z1_GESTCOM + Z1_GE1CCOM + Z1_GE2CCOM) AS GRAMTCOM, "
        SQL = SQL & "   Z1_COLUCOM "
        SQL = SQL & "FROM SZ1010 "
        SQL = SQL & "WHERE (Z1_STATUS IN ('I','T')) "
        SQL = SQL & "   AND D_E_L_E_T_ <> '*' "
        SQL = SQL & "   AND Z1_STATUS <> 'T' "
        SQL = SQL & "   AND SUBSTRING(Z1_ESTRUT,11,12) = 'A' "
        SQL = SQL & "ORDER BY Z1_ESTRUT  "

        Dados.Open(SQL)

    %>
    <div class="container-fluid py-3 px-3 px-lg-4">

        <div class="mb-3">
            <span class="page-title">Estruturas do Papelão Ondulado — Pirapetinga</span>
        </div>

        <div class="row g-4 mb-4">

            <!-- Estruturas sem tratamento -->
            <div class="col-lg-6">
                <div class="section-card">
                    <div class="section-card-header">
                        <i class="bi bi-grid-3x3-gap me-1" style="color: var(--color-text-muted);"></i>Estruturas sem tratamento
                    </div>
                    <div class="table-wrapper">
                        <table class="table table-bordered table-sm align-middle mb-0">
                            <thead>
                                <tr>
                                    <th>Nomenclatura</th>
                                    <th>Gramatura (g/m²)</th>
                                </tr>
                            </thead>
                            <tbody>
                            <%
                            x = 1
                            While not Dados.Eof
                            %>
                                <tr>
                                    <td><%=estrutura(Dados("ESTRUTURA"))%></td>
                                    <%if Dados("GRAMTCOM") <> 0 then %>
                                    <td><%=Dados("GRAMTCOM")%></td>
                                    <%else %>
                                    <td>-</td>
                                    <%end if%>
                                </tr>
                            <%
                                Dados.MoveNext()
                                x = x + 1
                            Wend

                            Dados.Close()
                            %>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <%
            SQL = "SELECT "
            SQL = SQL & "Z1_ESTRUT AS ESTRUTURA, "
            SQL = SQL & "Z1_STATUS, "
            SQL = SQL & "(Z1_GESTPIR) AS GRAMATPIR, "
            SQL = SQL & "(Z1_GESTCOM + Z1_GE1CCOM + Z1_GE2CCOM) AS GRAMTCOM, "
            SQL = SQL & "Z1_COLUCOM "
            SQL = SQL & "FROM SZ1010 "
            SQL = SQL & "WHERE (Z1_STATUS IN ('I','T')) "
            SQL = SQL & "AND D_E_L_E_T_ <> '*' "
            SQL = SQL & "AND Z1_STATUS = 'T' "
            SQL = SQL & "   AND SUBSTRING(Z1_ESTRUT,11,12) = 'A' "
            SQL = SQL & "ORDER BY Z1_ESTRUT "

            Dados.Open(SQL)
            %>

            <!-- Estruturas com tratamento -->
            <div class="col-lg-6">
                <div class="section-card">
                    <div class="section-card-header">
                        <i class="bi bi-grid-3x3-gap-fill me-1" style="color: var(--color-text-muted);"></i>Estruturas com tratamento
                    </div>
                    <div class="table-wrapper">
                        <table class="table table-bordered table-sm align-middle mb-0">
                            <thead>
                                <tr>
                                    <th>Nomenclatura</th>
                                    <th>Gramatura (g/m²)</th>
                                </tr>
                            </thead>
                            <tbody>
                            <%
                            x = 1
                            While not Dados.Eof
                            %>
                                <tr>
                                    <td><%=estrutura(Dados("ESTRUTURA"))%></td>
                                    <%if Dados("GRAMTCOM") <> 0 then %>
                                    <td><%=Dados("GRAMTCOM")%></td>
                                    <%else %>
                                    <td>-</td>
                                    <%end if%>
                                </tr>
                            <%
                                Dados.MoveNext()
                                x = x + 1
                            Wend

                            Dados.Close()
                            %>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

        </div>

        <!-- Legenda e dados construtivos -->
        <div class="section-card">
            <div class="section-card-header">
                <i class="bi bi-info-circle me-1" style="color: var(--color-text-muted);"></i>Legenda e dados construtivos
            </div>
            <div class="section-card-body">

                <div class="row g-4">
                    <!-- Legenda dos papéis -->
                    <div class="col-lg-6">
                        <p class="subhead">Legenda dos papéis</p>
                        <div class="table-wrapper">
                            <table class="table table-bordered table-sm align-middle mb-0">
                                <thead>
                                    <tr>
                                        <th style="width: 6rem;">Sigla</th>
                                        <th>Descrição</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr><th class="cell-label cell-code">TL</th><td>Test Liner</td></tr>
                                    <tr><th class="cell-label cell-code">KL</th><td>Kraft Liner (Terceiro)</td></tr>
                                    <tr><th class="cell-label cell-code">WT</th><td>Branco</td></tr>
                                    <tr><th class="cell-label cell-code">MI</th><td>Miolo</td></tr>
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <!-- Garantia dos testes -->
                    <div class="col-lg-6">
                        <p class="subhead">Garantia dos testes</p>
                        <div class="note">
                            <p>
                                <strong>Umidade da Chapa:</strong> 5,0 a 8,5%<br />
                                <strong>Umidade relativa do Ar:</strong> 57% +/- 8%<br />
                                <strong>Temperatura Ambiente:</strong> 25°C +/- 2°C
                            </p>
                            <p>
                                *Testes físicos: devem ser consideradas as condições de climatização controladas
                                em temperatura de 25°C +/- 2°C e umidade relativa de 57% +/- 8%
                                (<em>NBR NM-ISO 187</em>). Considerando o clima dos países tropicais.
                            </p>
                        </div>
                    </div>
                </div>

                <div class="row g-4 mt-1">
                    <!-- Tipo de tratamento -->
                    <div class="col-lg-6">
                        <p class="subhead">Tipo de tratamento</p>
                        <div class="table-wrapper">
                            <table class="table table-bordered table-sm align-middle mb-0">
                                <thead>
                                    <tr>
                                        <th style="width: 6rem;">Tipo</th>
                                        <th>Descrição</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr><th rowspan="2" class="cell-label cell-code">1</th><td>Cobb Interno de 12-16</td></tr>
                                    <tr><td>Cobb Externo de 16-22</td></tr>
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <!-- Tratamento das capas -->
                    <div class="col-lg-6">
                        <p class="subhead">Tratamento das capas</p>
                        <div class="table-wrapper">
                            <table class="table table-bordered table-sm align-middle mb-0">
                                <thead>
                                    <tr>
                                        <th style="width: 6rem;">Código</th>
                                        <th>Descrição</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr><th class="cell-label cell-code">001</th><td>Capa interna</td></tr>
                                    <tr><th class="cell-label cell-code">101</th><td>Capa interna e externa</td></tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- Dados construtivos -->
                <p class="subhead mt-4">Dados construtivos</p>
                <div class="table-wrapper">
                    <table class="table table-bordered table-sm align-middle mb-0">
                        <thead>
                            <tr>
                                <th colspan="2">Parâmetro</th>
                                <th>Pirapetinga</th>
                                <th>Uberaba</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr><th rowspan="3" class="cell-label">TAKEUP</th><td>Onda B</td><td>1,320</td><td>1,340</td></tr>
                            <tr><td>Onda C</td><td>1,400</td><td>1,400</td></tr>
                            <tr><td>Onda E</td><td>1,280</td><td>1,280</td></tr>
                            <tr><th rowspan="2" class="cell-label">COLA (g/m²)</th><td>Parede Simples</td><td>10</td><td>10</td></tr>
                            <tr><td>Parede Dupla</td><td>20</td><td>20</td></tr>
                            <tr><th rowspan="2" class="cell-label">Tratamento (g/m²)</th><td>Uma Capa</td><td>8</td><td>8</td></tr>
                            <tr><td>Duas Capas</td><td>16</td><td>16</td></tr>
                        </tbody>
                    </table>
                </div>

                <!-- Legenda — unidade de fabricação -->
                <p class="subhead mt-4">Legenda — unidade de fabricação</p>
                <div class="table-wrapper">
                    <table class="table table-bordered table-sm align-middle mb-0">
                        <thead>
                            <tr>
                                <th style="width: 14rem;">Nomenclatura</th>
                                <th>Unidade</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr><th class="cell-label cell-code">xx.xx.xxx.xxx.A</th><td>Pirapetinga</td></tr>
                            <tr><th class="cell-label cell-code">xx.xx.xxx.xxx.B</th><td>Uberaba</td></tr>
                        </tbody>
                    </table>
                </div>

                <!-- Informações complementares -->
                <p class="subhead mt-4">Informações complementares</p>
                <div class="note mb-3">
                    <strong>Objetivo:</strong> favorecer a conjugação de estruturas tratadas e acima de 7,0 coluna.
                </div>

                <p class="subhead">Medidas auto-conjugáveis (largura da chapa com refiles)</p>
                <div class="measure-list mb-3">
                    <span class="measure-chip">324 a 347 mm</span>
                    <span class="measure-chip">378 a 405 mm</span>
                    <span class="measure-chip">454 a 486 mm</span>
                    <span class="measure-chip">568 a 607 mm</span>
                    <span class="measure-chip">757 a 810 mm</span>
                    <span class="measure-chip">1135 a 1215 mm</span>
                    <span class="measure-chip">2270 a 2430 mm</span>
                </div>

                <div class="note">
                    *Testes de Coluna (ECT) para papelão com tratamento: em função da absorção da resina,
                    deve ser considerado teste com 1 kgf/cm abaixo do valor da estrutura normal sem tratamento.
                </div>

            </div>
        </div>

    </div>
    <%
    function estrutura(estrut)
        estrutura = mid(estrut,1,2)&"."&mid(estrut,3,2)&"."&mid(estrut, 5,3)&"."&mid(estrut,8,3)&"."&mid(estrut,11,1)
    end function

    end if
    %>

    <!-- Bootstrap JS Bundle -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>
    </body>
</html>
