<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../../../cnn2/cnn2.asp" -->
<!--#include file="../../../../funcoes/funcoes.asp" -->
<!--#include file="../../../../includes/variaveis.asp" -->
<!--#include file="../../../../../Funcoes/FuncoesAsp.asp" -->
<%
Response.Buffer  = true
Response.Expires = 0
Session.LCID     = 1046

'Recuperando dados do formulário da página anterior (alteração)

'Dados que não podem ser alterados
idSolicitacao = trim(Request("IdSolicitacao"))
Pedido        = trim(Request("hPedido"))
Cliente       = trim(Request("hCliente"))
Loja          = trim(Request("hLoja"))
ExisteExcessaoCargaMaquina = 0

SQL = "" &_
    "SELECT " &_
        "SA1010.A1_NOME AS Nome " &_
    "FROM SA1010 " &_
    "WHERE Rtrim(SA1010.A1_COD) = '" & Cliente & "' " &_
      "AND Rtrim(SA1010.A1_LOJA) = '" & Loja & "' " &_
      "AND D_E_L_E_T_ = ''"

SET bNomeCliente = Conecta("PIR-DADOSADV", SQL)

if not isObject(bNomeCliente) or (bNomeCliente IS NOTHING) then
    response.end
end if

NomeCliente = iif( not bNomeCliente.EOF, bNomeCliente("Nome"), "" )

bNomeCliente.Close()
Set bNomeCliente = NOTHING

SQL = "" &_
    "SELECT " &_
        "C5_DESTINO, " &_
        "C5_LOJDEST " &_
    "FROM SC5010 " &_
    "WHERE C5_NUM = '" & Pedido & "' " &_
      "AND D_E_L_E_T_ = '' " &_
      "AND C5_REMESSA = 'F'"

SET Remessa = Conecta("PIR-DADOSADV", SQL)

if not isObject(Remessa) or (Remessa IS NOTHING) then
    response.end
end if

if not Remessa.EOF then
    clienteRemessa = trim(Remessa("C5_DESTINO"))
else
    clienteRemessa = ""
end if

if not Remessa.EOF then
    lojaRemessa = trim(Remessa("C5_LOJDEST"))
else
    lojaRemessa = ""
end if

Remessa.Close()
SET Remessa = Nothing

NomeTipoOperacao = trim(Request("hNomeTipoOperacao"))
TipoOperacao     = trim(Request("hTipoOperacao"))
Produto          = trim(Request("hProduto"))
DescricaoProduto = trim(Request("hDescri"))
Quantidade       = Replace(Replace(trim(Request("hQuantidade")),".",""),",",".")
QuantidadePronta = trim(Request("hQuantidadepronta"))
Preco            = trim(Request("hPreco"))
Total            = trim(Request("hTotal"))
Entrega1         = trim(Request("hEntrega1"))
Entrega2         = trim(Request("hEntrega2"))
Emissao          = trim(Request("hEmissao"))
TipoFrete        = trim(Request("hTPFrete"))
DataCadastro     = trim(Request("hDataCadastro"))
HoraCadastro     = trim(Request("hHoraCadastro"))
QuemCadastrou    = trim(Request("hQuemCadastrou"))
DataLiberado     = trim(Request("hDataLiberado"))
HoraLiberado     = trim(Request("hHoraLiberado"))
QuemLiberou      = trim(Request("hQuemLiberou"))
CondPag          = trim(Request("hCodCondPag"))
Estatus          = trim(Request("hEstatus"))
PedidoCliente    = trim(Request("hPedidoCliente"))
Filial           = trim(Request("Filial"))

ValorFrete = 0

if (trim(Request("valorfrete")) <> "" or trim(Request("NovoFreteUnitario")) <> "") and left(Produto,3) = "700" then
    ValorFrete = iif( trim(Request("valorfrete")) <> "", trim(Request("valorfrete")), trim(Request("NovoFreteUnitario")) )
end if

'Dados de registro
DataAlteracao = InverteData(Date())
HoraAlteracao = Time()
Usuario = Usr

if mid(Produto,1,3) = "720" then

    SQL = "" &_
        "SELECT " &_
            "C6_LARG, " &_
            "C6_COMP " &_
        "FROM SC6010 " &_
        "WHERE C6_NUM = '" & Pedido & "' " &_
          "AND D_E_L_E_T_ = '' "

    SET PedidoChapa = Conecta("PIR-DADOSADV", SQL)

    if not isobject(PedidoChapa) or (PedidoChapa IS NOTHING) then
        response.end
    end if

    if not PedidoChapa.EOF then
        Largura     = PedidoChapa("C6_LARG")
        Comprimento = PedidoChapa("C6_COMP")
    else
        Largura     = 0
        Comprimento = 0
    end if

    PedidoChapa.Close
    SET PedidoChapa = Nothing

else

    Largura     = 0
    Comprimento = 0

end if

'Se o tipo de operação for diferente de 6 - condições de pagamento
'conecta na tabela de condições de pagamento para pegar o nome pelo código
if TipoOperacao <> 6 then

    SQL = "" &_
        "SELECT " &_
            "SE4010.E4_DESCRI AS Nome " &_
        "FROM SE4010 " &_
        "WHERE SE4010.E4_CODIGO = '" & CondPag & "'"

    SET CondPagamento = Conecta("PIR-DADOSADV", SQL)

    if not isobject(CondPagamento) or (CondPagamento IS NOTHING) then
        response.end
    end if

    NomeCondPag = iif( not CondPagamento.EOF, NomeCondPag = CondPagamento("Nome"), "" )

end if

'Dados que poderão sofrer alterações
select case TipoOperacao

case 1 'TIPO DE FRETE

    NovoTipoFrete = trim(Request("NovoTPFrete"))
    De            = TipoFrete
    Para          = NovoTipoFrete
    NovoTotal     = 0

    if (trim(Para) = "") Or (trim(Para) <> "C" And trim(Para) <> "F") then
        Aviso("Preencha o novo tipo de frete.")
        response.end
    end if

    if (trim(Para) = "C" and (ValorFrete = "0" or ValorFrete = "") and left(Produto,3) = "700") then
        Aviso("Preencha o valor do frete.")
        response.end
    end if

case 2 'PREÇO

    NovoPreco = trim(Request("NovoPreco"))
    De        = Preco
    Para      = NovoPreco

    if (trim(Para) = "") then
        Aviso("Preencha o novo preço.")
        response.end
    end if

    SQL = "" &_
        "SELECT " &_
            "CASE F4_IPI WHEN 'S' THEN " &_
                "ISNULL((C6_QTDVEN - C6_QTDENT) * " & replace(trim(Para),",",".") & ",0)* (1+(B1_IPI/100)) " &_
            "ELSE " &_
                "ISNULL((C6_QTDVEN - C6_QTDENT) * " & replace(trim(Para),",",".") & ",0) " &_
            "END NOVO_SALDO " &_
        "FROM SC6010 " &_
            "INNER JOIN SC5010 ON SC6010.C6_NUM = SC5010.C5_NUM AND SC5010.D_E_L_E_T_ = '' " &_
            "INNER JOIN SF4010 ON SF4010.F4_CODIGO = SC6010.C6_TES AND SF4010.D_E_L_E_T_ = '' " &_
            "INNER JOIN SB1010 ON B1_COD = C6_PRODUTO AND SB1010.D_E_L_E_T_ = '' AND SUBSTRING(B1_COD,1,3) <> '120' " &_
        "WHERE C6_QTDVEN - C6_QTDENT > 0 " &_
          "AND C5_ESTATUS NOT IN ('L','C') " &_
          "AND (C5_NUM BETWEEN '000000' AND '999999' OR C5_NUM BETWEEN 'AA0000' AND 'ZZ9999') " &_
          "AND SC6010.D_E_L_E_T_ = '' " &_
          "AND C5_TIPO NOT IN ('D','B') " &_
          "AND C5_NUM = '" & trim(Pedido) & "'"

    SET NovoValor = Conecta("PIR-DADOSADV", SQL)

    if not isobject(NovoValor) or (NovoValor IS NOTHING) then
        response.end
    end if

    if not NovoValor.EOF then
        NovoTotal = NovoValor("NOVO_SALDO")
    else
        NovoTotal = cdbl(NovoPreco) * cdbl(Quantidade)
    end if

    if Para > De then
        if NovoValor.EOF Or (EVAL(NovoTotal - Total) > EVAL(NovoValor("NOVO_SALDO"))) then
            Aviso("O limite de crédito do cliente foi ultrapassado. (1)")
            NovoValor.Close()
            SET NovoValor = NOTHING
            response.end
        end if
    end if

    NovoValor.Close()
    SET NovoValor = NOTHING

case 3,12 'QUANTIDADE

    NovaQuantidade = Replace(FormatNumber(trim(Request("NovaQuantidade")),0),".","")
    De             = Quantidade
    Para           = NovaQuantidade

    SQL = "" &_
        "SELECT " &_
            "CASE F4_IPI WHEN 'S' THEN " &_
                "ISNULL((" & (Request("NovaQuantidade") - Quantidade) & " - C6_QTDENT) * C6_PRCVEN,0)* (1+(B1_IPI/100)) " &_
            "ELSE " &_
                "ISNULL((" & (Request("NovaQuantidade") - Quantidade) & " - C6_QTDENT) * C6_PRCVEN,0) " &_
            "END NOVO_SALDO," &_
            "CASE F4_IPI WHEN 'S' THEN " &_
                "ISNULL((" & (Request("NovaQuantidade")) & " - C6_QTDENT) * C6_PRCVEN,0)* (1+(B1_IPI/100)) " &_
            "ELSE " &_
                "ISNULL((" & (Request("NovaQuantidade")) & " - C6_QTDENT) * C6_PRCVEN,0) " &_
            "END NOVO_TOTAL, " &_
            "CASE F4_IPI WHEN 'S' THEN " &_
                "ISNULL((C6_QTDVEN - C6_QTDENT) * C6_PRCVEN,0)* (1+(B1_IPI/100)) " &_
            "ELSE " &_
                "ISNULL((C6_QTDVEN - C6_QTDENT) * C6_PRCVEN,0) " &_
            "END TOTAL, " &_
            Request("NovaQuantidade") & " - C6_QTDENT AS SALDO_QUANTIDADE " &_
        "FROM SC6010 " &_
            "INNER JOIN SC5010 ON SC6010.C6_NUM = SC5010.C5_NUM AND SC5010.D_E_L_E_T_ = '' " &_
            "INNER JOIN SF4010 ON SF4010.F4_CODIGO = SC6010.C6_TES AND SF4010.D_E_L_E_T_ = '' " &_
            "INNER JOIN SB1010 ON B1_COD = C6_PRODUTO AND SB1010.D_E_L_E_T_ = '' AND SUBSTRING(B1_COD,1,3) <> '120' " &_
        "WHERE C6_QTDVEN - C6_QTDENT > 0 " &_
            "AND C5_ESTATUS NOT IN ('L','C') " &_
            "AND SC6010.D_E_L_E_T_ = '' " &_
            "AND C5_TIPO NOT IN ('D','B') " &_
            "AND C5_NUM = '" & trim(Pedido) & "' " &_
            "AND F4_DUPLIC = 'S' "

    SET NovoValor = Conecta("PIR-DADOSADV", SQL)

    if not isobject(NovoValor) or (NovoValor IS NOTHING) then
        response.end
    end if

    if not NovoValor.EOF then

        NovoTotal       = NovoValor("NOVO_TOTAL")
        DiferencaTotal  = NovoValor("NOVO_SALDO")
        SaldoQuantidade = Cdbl(NovoValor("SALDO_QUANTIDADE"))
        totalComIPI     = NovoValor("TOTAL")

        if usr = "adriana.silva" then
           response.write SaldoQuantidade & " - 1<br />"
           response.write SQL
        end if

    else

        if Cdbl(replace(NovaQuantidade,".",",")) > Cdbl(replace(Quantidade,".",",")) then
            NovoTotal      = Cdbl(Preco) * Cdbl(NovaQuantidade)
            DiferencaTotal = Cdbl(Preco) * Cdbl(NovaQuantidade - Cdbl(replace(Quantidade,".",",")))
            SaldoQuantidade = Cdbl(NovaQuantidade - Cdbl(replace(Quantidade,".",",")))
            totalComIPI    = Total
        else
            NovoTotal      = Cdbl(Preco) * Cdbl(NovaQuantidade)
            DiferencaTotal = Cdbl(Preco) * (Cdbl(replace(Quantidade,".",","))-Cdbl(NovaQuantidade))
            SaldoQuantidade = Cdbl(replace(Quantidade,".",","))-Cdbl(NovaQuantidade)
            totalComIPI    = Total
        end if

    end if

    NovoValor.Close()
    SET NovoValor = NOTHING

    if (trim(Para) = "") then
        Aviso("Preencha a nova quantidade.")
        response.end
    end if

    if EVAL(SaldoQuantidade) < 0 then
        Aviso("Não é possível alterar a quantidade para um valor menor que o já expedido.")
        response.end
    end if

    SQL = "{call dbo.SP_CONSULTA_CREDITO_CLIENTE('" & Cliente & "',0)}"

    SET LimiteCredito = Conecta("PIR-DADOSADV", SQL)

    if not isobject(LimiteCredito) or (LimiteCredito IS NOTHING) then
        response.end
    end if

    if LimiteCredito.EOF Or (EVAL(NovoTotal - totalComIPI) > EVAL(LimiteCredito("SALDO_CREDITO"))) then
        Aviso("O limite de crédito do cliente foi ultrapassado. (2)")
        response.end
    end if

    if Cdbl(DiferencaTotal) > Cdbl(LimiteCredito("SALDO_CREDITO")) then
        Aviso("O limite de crédito do cliente foi ultrapassado. (3)")
        response.end
    end if

    Total = totalComIPI

    LimiteCredito.Close()
    SET LimiteCredito = Nothing

case 4 'QUANTIDADE PRONTO

    NovaQuantidadePronta = trim(Request("NovaQuantidadePronta"))
    De        = QuantidadePronta
    Para      = NovaQuantidadePronta
    NovoTotal = 0

    if (trim(Para) = "") then
        Aviso("Preencha a nova quantidade do produto pronto.")
        response.end
    end if

case 5 'DATA DE ENTREGA

    NovaEntrega1 = trim(Request("NovaEntrega1"))
    NovaEntrega2 = trim(Request("NovaEntrega2"))

    if inverteData(NovaEntrega1) > invertedata(NovaEntrega2) then
        Aviso("A primeira data de entrega deve ser maior ou igual a segunda data.")
        response.end
    end if

    ExisteDisponibilidadeMaquina    = 0
    DataDisponibilidadeCargaMaquina = NovaEntrega1
    De        = AcertaData(Entrega1) & " - " & AcertaData(Entrega2)
    Para      = ConverterData(NovaEntrega1,"DD/MM/AAAA") & " - " & ConverterData(NovaEntrega2,"DD/MM/AAAA")
    NovoTotal = 0

    if (trim(NovaEntrega1) = "") Or (trim(NovaEntrega2) = "") then
        Aviso("Preencha a nova data de entrega corretamente. A primeira e a segunda, mesmo que seja igual a anterior.")
        response.end
    end if

case 6 'CONDIÇÕES DE PAGAMENTO

    if NovaCondPag = "090" then
        Aviso("Esta opção está temporariamente desativada para alteração de pedidos.")
        response.end
    end if

    NomeCondPag    = trim(Request("hNomeCondPag"))
    NovaCondPag    = trim(Request("NovaCondPag"))
    NomeNovaCondPag = trim(Request("NomeNovaCondPag"))
    De             = NomeCondPag
    Para           = NomeNovaCondPag
    AlteraCondPag  = NovaCondPag
    NovoTotal      = 0

case 7 'CANCELAMENTO

    NovoEstatus = trim(Request("NovoEstatus"))
    De          = Estatus
    Para        = NovoEstatus
    NovoTotal   = 0

    if (trim(Para) = "") or (trim(Para) <> "C") then
        Aviso("Marque a opção de cancelar na página anterior.")
        response.end
    end if

case 8 'PEDIDO DO CLIENTE

    NovoPedidoCliente = trim(Request("NovoPedidoCliente"))
    De        = PedidoCliente
    Para      = NovoPedidoCliente
    NovoTotal = 0

    if (trim(Para) = "") then
        Aviso("Informe o novo pedido do cliente na página anterior.")
        response.end
    end if

case 9,13 'ALTERAÇÃO DE STATUS (menos cancelamento)

    NovoEstatus2 = trim(Request("NovoEstatus2"))
    De           = Estatus
    Para         = NovoEstatus2
    NovoTotal    = 0

    SQL = "" &_
        "SELECT " &_
            "(C6_QTDVEN - C6_QTDENT) * C6_PRCVEN AS TOTAL " &_
        "FROM SC6010 " &_
        "WHERE C6_NUM = '" & trim(Pedido) & "' " &_
          "AND D_E_L_E_T_ <> '*' " &_
          "AND SUBSTRING(C6_PRODUTO,1,3) <> '120' "

    SET rsPedido = Conecta("PIR-DADOSADV", SQL)

    if not isobject(rsPedido) or (rsPedido IS NOTHING) then
        response.end
    end if

    Total = iif( not rsPedido.EOF, rsPedido("TOTAL"), 0 )

    if (De = "L" Or De = "C") And (Para <> "L" And Para <> "C") then

        SQL = "EXECUTE dbo.SP_CONSULTA_CREDITO_CLIENTE '" & Cliente & "', 0 "

        SET LimiteCredito = Conecta("PIR-DADOSADV", SQL)

        if not isobject(LimiteCredito) or (LimiteCredito IS NOTHING) then
            response.end
        end if

        if LimiteCredito.EOF or ( Total > LimiteCredito("SALDO_CREDITO") ) then
            Aviso("O limite de crédito do cliente foi ultrapassado. (4)")
            response.end
        end if

        LimiteCredito.Close()
        Set LimiteCredito = Nothing

    end if

    rsPedido.Close()
    SET rsPedido = NOTHING

case 10 'ALTERAÇÃO DE FILIAL

    De        = Filial
    Para      = trim(Request("NovaFilial"))
    NovoTotal = 0

    SQL = "" &_
        "SELECT " &_
            "CASE '" & Para & "' WHEN 'I' THEN " &_
                "B1_FERRPIR " &_
            "ELSE " &_
                "B1_FERRUBE " &_
            "END AS FERRAMENTAL, " &_
            "B1_DESCNF " &_
        "FROM DADOSADV..SB1010 " &_
        "WHERE B1_COD = '" & Produto & "'"

    SET rsFerramental = Conecta("PIR-DADOSADV", SQL)

    if not isobject(rsFerramental) or (rsFerramental IS NOTHING) then
        response.end
    end if

    if rsFerramental.EOF or (rsFerramental("FERRAMENTAL") <> "S" and (rsFerramental("B1_DESCNF") = "901" or rsFerramental("B1_DESCNF") = "902")) then
        Aviso("Ferramental para o produto não disponível na Filial solicitada.<br>Favor entrar em contato com Gilvan Junior - Gerente de PCP")
        response.end
    end if

    rsFerramental.Close()
    SET rsFerramental = NOTHING

case 11 'ALTERAÇÃO MOTIVO DO ATRASO

    De        = request("MotivoAtrazo")
    Para      = Request("NovoMotivoAtrazo")
    NovoTotal = 0

case 14 'ALTERAÇÃO DE FRETE

    if (ValorFrete = "0" or ValorFrete = "") then
        Aviso("Preencha o valor do frete.")
        response.end
    end if

    De        = trim(Request("FreteUnitario"))
    Para      = ValorFrete
    NovoTotal = 0

end select
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>SWnet - Confirmação da Alteração do Pedido</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css"
          rel="stylesheet"
          integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB"
          crossorigin="anonymous" />
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css"
          rel="stylesheet" />

    <style>
        /* ── Tokens de Design ─────────────────────────────────────────────────── */
        :root {
            --color-bg:            hsl(240, 9%, 98%);
            --color-surface:       hsl(0, 0%, 100%);
            --color-surface-muted: hsl(240, 9%, 97%);
            --color-border:        hsl(240, 6%, 90%);
            --color-border-strong: hsl(240, 6%, 82%);

            --color-text:          hsl(240, 6%, 10%);
            --color-text-muted:    hsl(240, 4%, 46%);
            --color-text-faint:    hsl(240, 4%, 64%);

            --color-accent:        hsl(210, 100%, 45%);
            --color-accent-hover:  hsl(210, 100%, 40%);
            --color-accent-soft:   hsl(210, 100%, 96%);

            --color-success:       hsl(145, 63%, 32%);
            --color-success-soft:  hsl(145, 60%, 95%);
            --color-danger:        hsl(3, 90%, 58%);
            --color-danger-soft:   hsl(3, 100%, 96%);
            --color-warning:       hsl(35, 100%, 48%);
            --color-warning-soft:  hsl(40, 100%, 95%);

            --font-sans: -apple-system, BlinkMacSystemFont, "Segoe UI", "Inter", Roboto, Helvetica, Arial, sans-serif;

            --radius-sm:  6px;
            --radius-md:  10px;
            --radius-lg:  14px;

            --shadow-sm: 0 1px 2px hsla(240, 6%, 10%, .04);
            --shadow-md: 0 4px 16px hsla(240, 6%, 10%, .08);

            --transition: 160ms ease;
        }

        /* ── Base ─────────────────────────────────────────────────────────────── */
        *, *::before, *::after { box-sizing: border-box; }

        body {
            background-color: var(--color-bg);
            font-family: var(--font-sans);
            color: var(--color-text);
            min-height: 100vh;
            margin: 0;
        }

        /* ── Título da página ─────────────────────────────────────────────────── */
        .page-title {
            font-size: clamp(1rem, 0.9rem + 0.4vw, 1.25rem);
            font-weight: 600;
            letter-spacing: -0.01em;
            color: var(--color-text);
        }

        /* ── Card principal ───────────────────────────────────────────────────── */
        .main-card {
            background-color: var(--color-surface);
            border: 1px solid var(--color-border);
            border-radius: var(--radius-md);
            box-shadow: var(--shadow-md);
        }

        .main-card-header {
            background-color: var(--color-surface);
            border-bottom: 1px solid var(--color-border);
            border-radius: var(--radius-md) var(--radius-md) 0 0;
            padding: .875rem 1.25rem;
        }

        .main-card-body {
            padding: 1.25rem 1.25rem 1.5rem;
        }

        /* ── Filter card ──────────────────────────────────────────────────────── */
        .filter-card {
            background-color: var(--color-surface);
            border: 1px solid var(--color-border);
            border-radius: var(--radius-md);
            box-shadow: var(--shadow-sm);
        }

        .filter-card-header {
            background-color: var(--color-surface);
            color: var(--color-text);
            padding: .7rem 1.1rem;
            border-radius: var(--radius-md) var(--radius-md) 0 0;
            border-bottom: 1px solid var(--color-border);
            font-weight: 600;
            font-size: .85rem;
        }

        .filter-card-body {
            padding: 1.1rem 1.25rem 1.35rem;
        }

        /* ── Dados inline (label + valor) ─────────────────────────────────────── */
        .dado-label {
            font-size: .78rem;
            font-weight: 600;
            color: var(--color-text-muted);
            letter-spacing: .01em;
        }

        .dado-valor {
            font-size: .875rem;
            color: var(--color-text);
        }

        /* ── Bloco De / Para ──────────────────────────────────────────────────── */
        .de-para-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 1rem;
        }

        @media (max-width: 576px) {
            .de-para-grid { grid-template-columns: 1fr; }
        }

        .de-para-item {
            background-color: var(--color-surface-muted);
            border: 1px solid var(--color-border);
            border-radius: var(--radius-sm);
            padding: .875rem 1rem;
        }

        .de-para-item.para {
            background-color: var(--color-accent-soft);
            border-color: hsl(210, 100%, 85%);
        }

        .de-para-badge {
            display: inline-block;
            font-size: .7rem;
            font-weight: 700;
            letter-spacing: .06em;
            text-transform: uppercase;
            padding: .15rem .5rem;
            border-radius: var(--radius-sm);
            margin-bottom: .4rem;
        }

        .de-para-badge.de {
            background-color: var(--color-border);
            color: var(--color-text-muted);
        }

        .de-para-badge.para {
            background-color: var(--color-accent);
            color: hsl(0, 0%, 100%);
        }

        .de-para-value {
            font-size: .925rem;
            font-weight: 500;
            color: var(--color-text);
        }

        /* ── Alerta de atenção ────────────────────────────────────────────────── */
        .alerta-atencao {
            background-color: var(--color-warning-soft);
            border: 1px solid hsl(35, 100%, 78%);
            border-radius: var(--radius-sm);
            padding: .875rem 1rem;
            font-size: .875rem;
            color: hsl(30, 80%, 30%);
            display: flex;
            gap: .6rem;
            align-items: flex-start;
        }

        .alerta-atencao .bi {
            flex-shrink: 0;
            margin-top: .1rem;
            color: var(--color-warning);
        }

        /* ── Divisor ──────────────────────────────────────────────────────────── */
        hr {
            border: none;
            border-top: 1px solid var(--color-border);
            margin: 1.25rem 0;
        }

        /* ── Botões ───────────────────────────────────────────────────────────── */
        .btn {
            border-radius: var(--radius-sm);
            font-weight: 500;
            transition: background-color var(--transition), border-color var(--transition), box-shadow var(--transition);
        }

        .btn:focus-visible {
            outline: none;
            box-shadow: 0 0 0 3px var(--color-accent-soft);
        }

        .btn:active { transform: scale(.98); }

        .btn-primary {
            background-color: var(--color-accent);
            border-color: var(--color-accent);
            color: hsl(0, 0%, 100%);
        }

        .btn-primary:hover {
            background-color: var(--color-accent-hover);
            border-color: var(--color-accent-hover);
        }

        .btn-primary:disabled,
        .btn-primary[disabled] {
            opacity: .5;
            cursor: not-allowed;
            pointer-events: none;
        }

        .btn-secondary-outline {
            background-color: var(--color-surface);
            border: 1px solid var(--color-border);
            color: var(--color-text);
        }

        .btn-secondary-outline:hover {
            background-color: var(--color-surface-muted);
            border-color: var(--color-border-strong);
            color: var(--color-text);
        }

        /* ── Loader ───────────────────────────────────────────────────────────── */
        #overlay-loader {
            display: none;
            position: fixed;
            inset: 0;
            background-color: hsla(0, 0%, 100%, .75);
            z-index: 1100;
            align-items: center;
            justify-content: center;
        }

        #overlay-loader.active {
            display: flex;
        }

        .loader-spinner {
            width: 44px;
            height: 44px;
            border: 4px solid var(--color-border);
            border-top-color: var(--color-accent);
            border-radius: 50%;
            animation: spin 700ms linear infinite;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        /* ── Responsividade ───────────────────────────────────────────────────── */
        @media (max-width: 768px) {
            .main-card-body { padding: 1rem; }
        }

        /* ── Acessibilidade ───────────────────────────────────────────────────── */
        @media (prefers-reduced-motion: reduce) {
            *, *::before, *::after {
                animation-duration: 0.01ms !important;
                transition-duration: 0.01ms !important;
            }
        }
    </style>
</head>

<body onload="doOnLoad()">

    <!-- ── Overlay de carregamento ──────────────────────────────────────────── -->
    <div id="overlay-loader" role="status" aria-live="polite">
        <div class="loader-spinner"></div>
    </div>

    <div class="container-fluid py-4 px-3 px-lg-4" style="max-width: 960px;">

        <!-- ── Cabeçalho ─────────────────────────────────────────────────────── -->
        <div class="mb-3">
            <span class="page-title">Alteração de Pedido</span>
        </div>

        <!-- ── Card principal ────────────────────────────────────────────────── -->
        <div class="main-card">

            <div class="main-card-header d-flex align-items-center gap-2">
                <i class="bi bi-patch-check" style="color: var(--color-text-muted);"></i>
                <span style="font-weight: 600; font-size: .9rem; color: var(--color-text);">Confirmação da Alteração</span>
            </div>

            <div class="main-card-body">

                <form id="form" name="form" method="post" action="AlteraPedido04.asp">

                    <!-- Hidden fields -->
                    <input type="hidden" name="IdSolicitacao"       id="IdSolicitacao"      value="<%=IdSolicitacao %>" />
                    <input type="hidden" name="Pedido"              id="Pedido"             value="<%=Pedido%>" />
                    <input type="hidden" name="Quantidade"          id="Quantidade"         value="<%=Quantidade%>" />
                    <input type="hidden" name="QuantidadePronta"    id="QuantidadePronta"   value="<%=QuantidadePronta%>" />
                    <input type="hidden" name="Preco"               id="Preco"              value="<%=Preco%>" />
                    <input type="hidden" name="Total"               id="Total"              value="<%=Total%>" />
                    <input type="hidden" name="Entrega1"            id="Entrega1"           value="<%=Entrega1%>" />
                    <input type="hidden" name="Entrega2"            id="Entrega2"           value="<%=Entrega2%>" />
                    <input type="hidden" name="Emissao"             id="Emissao"            value="<%=Emissao%>" />
                    <input type="hidden" name="TipoFrete"           id="TipoFrete"          value="<%=TipoFrete%>" />
                    <input type="hidden" name="NomeCondPag"         id="NomeCondPag"        value="<%=NomeCondPag%>" />
                    <input type="hidden" name="Estatus"             id="Estatus"            value="<%=Estatus%>" />
                    <input type="hidden" name="SiglaEstatusTrim"    id="SiglaEstatusTrim"   value="<%=SiglaEstatusTrim%>" />
                    <input type="hidden" name="QuemCadastrou"       id="QuemCadastrou"      value="<%=QuemCadastrou%>" />
                    <input type="hidden" name="DataCadastro"        id="DataCadastro"       value="<%=DataCadastro%>" />
                    <input type="hidden" name="HoraCadastro"        id="HoraCadastro"       value="<%=HoraCadastro%>" />
                    <input type="hidden" name="QuemLiberou"         id="QuemLiberou"        value="<%=QuemLiberou%>" />
                    <input type="hidden" name="DataLiberado"        id="DataLiberado"       value="<%=DataLiberado%>" />
                    <input type="hidden" name="HoraLiberado"        id="HoraLiberado"       value="<%=HoraLiberado%>" />
                    <input type="hidden" name="Cliente"             id="Cliente"            value="<%=Cliente%>" />
                    <input type="hidden" name="Loja"                id="Loja"               value="<%=Loja%>" />
                    <input type="hidden" name="NomeCliente"         id="NomeCliente"        value="<%=NomeCliente%>" />
                    <input type="hidden" name="ClienteRemessa"      id="ClienteRemessa"     value="<%=clienteRemessa%>" />
                    <input type="hidden" name="LojaRemessa"         id="LojaRemessa"        value="<%=lojaRemessa%>" />
                    <input type="hidden" name="Produto"             id="Produto"            value="<%=Produto%>" />
                    <input type="hidden" name="DescricaoProduto"    id="DescricaoProduto"   value="<%=DescricaoProduto%>" />
                    <input type="hidden" name="TipoOperacao"        id="TipoOperacao"       value="<%=TipoOperacao%>" />
                    <input type="hidden" name="NomeTipoOperacao"    id="NomeTipoOperacao"   value="<%=NomeTipoOperacao%>" />
                    <input type="hidden" name="PedidoCliente"       id="PedidoCliente"      value="<%=trim(PedidoCliente)%>" />
                    <input type="hidden" name="De"                  id="De"                 value="<%=De%>" />
                    <input type="hidden" name="Para"                id="Para"               value="<%=trim(Para)%>" />
                    <input type="hidden" name="NovoTotal"           id="NovoTotal"          value="<%=NovoTotal%>" />
                    <input type="hidden" name="AlteraCondPag"       id="AlteraCondPag"      value="<%=AlteraCondPag%>" />
                    <input type="hidden" name="DataAlteracao"       id="DataAlteracao"      value="<%=DataAlteracao%>" />
                    <input type="hidden" name="HoraAlteracao"       id="HoraAlteracao"      value="<%=HoraAlteracao%>" />
                    <input type="hidden" name="NovaEntrega1"        id="NovaEntrega1"       value="<%=NovaEntrega1%>" />
                    <input type="hidden" name="NovaEntrega2"        id="NovaEntrega2"       value="<%=NovaEntrega2%>" />
                    <input type="hidden" name="NovaFilial"          id="NovaFilial"         value="<%=NovaFilial%>" />
                    <input type="hidden" name="valorfrete"          id="valorfrete"         value="<%=ValorFrete%>" />
                    <input type="hidden" name="Largura"             id="Largura"            value="<%=Largura%>" />
                    <input type="hidden" name="Comprimento"         id="Comprimento"        value="<%=Comprimento%>" />
                    <input type="hidden" name="FreteUnitario"       id="FreteUnitario"      value="<%=trim(Request("FreteUnitario"))%>" />

                    <!-- ── Dados do Pedido ───────────────────────────────────────────── -->
                    <div class="filter-card mb-3">
                        <div class="filter-card-header">
                            <i class="bi bi-info-circle me-2" style="color: var(--color-text-muted);"></i>Dados do Pedido
                        </div>
                        <div class="filter-card-body">
                            <div class="row g-3">
                                <div class="col-sm-6 col-lg-4">
                                    <span class="dado-label d-block">Pedido</span>
                                    <span class="dado-valor fw-semibold"><%=Pedido%></span>
                                </div>
                                <div class="col-sm-6 col-lg-4">
                                    <span class="dado-label d-block">Cliente</span>
                                    <span class="dado-valor"><%=Cliente & " - " & Loja & " - " & NomeCliente%></span>
                                </div>
                                <div class="col-sm-6 col-lg-4">
                                    <span class="dado-label d-block">Tipo de Operação</span>
                                    <span class="dado-valor"><%=NomeTipoOperacao%></span>
                                </div>
                            </div>
                        </div>
                    </div>

<%
if TipoOperacao = 5 then 'DATA DE ENTREGA

    if trim(Replace(Entrega2,"/","")) <> "" then
        if DateDiff("d",AcertaData(Entrega1),NovaEntrega1) = 0 And DateDiff("d",AcertaData(Entrega2),NovaEntrega2) = 0 then
            Mensagem      = "<b>ATENÇÃO</b><br>As datas de entregas não foram alteradas.<br>A operação não poderá ser realizada."
            HabilitaBotao = "disabled"
        else
            Mensagem      = ""
            HabilitaBotao = ""
        end if
    end if

    Mensagem = DateDiff("d",AcertaData(DataAlteracao),NovaEntrega1)

    if DateDiff("d",AcertaData(DataAlteracao),NovaEntrega1) < 3 Or DateDiff("d",AcertaData(DataAlteracao),NovaEntrega2) < 3 then
        Mensagem      = "<b>ATENÇÃO</b><br>A data limite para alteração de data de entrega está estourada.<br>A alteração da data de entrega deste pedido não é aconselhada."
        HabilitaBotao = ""
    else
        Mensagem      = ""
        HabilitaBotao = ""
    end if

end if

if Left(Produto,3) > "700" then

    SQL = "{call dbo.HistStatusPedidoInpanet('" & Pedido & "')}"

    SET EstatusTrim = Conecta("PIR-TRIMBOX", SQL)

    if not isobject(EstatusTrim) or (EstatusTrim IS NOTHING) then
        response.end
    end if

    if not EstatusTrim.EOF then
        stTrimBox = Estatustrim("StatusPedido")
    else
        sTrimBox = ""
    end if

    EstatusTrim.Close
    SET EstatusTrim = Nothing

    select case trim(stTrimBox)
    case "Pronto"     : SiglaEstatusTrim = "Pt" : Mensagem2 = "<b>ATENÇÃO!</b><br>Este pedido já está pronto!"
    case "Convertido" : SiglaEstatusTrim = "Co" : Mensagem2 = "<b>ATENÇÃO!</b><br>Este pedido já está convertido na impressora!"
    case "Ondulado"   : SiglaEstatusTrim = "On" : Mensagem2 = "<b>ATENÇÃO!</b><br>Este pedido já está ondulado e pronto para ser convertido."
    case "Programado" : SiglaEstatusTrim = "Pr" : Mensagem2 = "<b>ATENÇÃO!</b><br>Este pedido está Programado para produção."
    case "Aberto"     : SiglaEstatusTrim = "Ab" : Mensagem2 = ""
    case else
        SiglaEstatusTrim = "Ab"
        Mensagem2        = ""
    end select

else

    SiglaEstatusTrim = "NA"

end if

MensagemFinal = trim(Mensagem & Mensagem2)
%>
<%
if MensagemFinal <> "" then
%>
                    <!-- ── Alerta ───────────────────────────────────────────────────────── -->
                    <div class="alerta-atencao mb-3">
                        <i class="bi bi-exclamation-triangle-fill"></i>
                        <div><%=MensagemFinal%></div>
                    </div>
<%
end if
%>

                    <!-- ── Alteração a ser Confirmada ────────────────────────────────── -->
                    <div class="filter-card mb-3">
                        <div class="filter-card-header">
                            <i class="bi bi-arrow-left-right me-2" style="color: var(--color-text-muted);"></i>Alteração a ser Confirmada
                        </div>
                        <div class="filter-card-body">
                            <div class="de-para-grid">
                                <div class="de-para-item">
                                    <span class="de-para-badge de">De</span>
                                    <div class="de-para-value"><%=trim(De)%></div>
                                </div>
                                <div class="de-para-item para">
                                    <span class="de-para-badge para">Para</span>
                                    <div class="de-para-value"><%=trim(Para)%></div>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- ── Ações ─────────────────────────────────────────────────────── -->
                    <div class="d-flex justify-content-end gap-2 mt-3">
                        <a href="#" class="btn btn-secondary-outline btn-sm" onclick="history.go(-1); return false;">
                            <i class="bi bi-arrow-left me-1"></i>Voltar
                        </a>
                        <button type="button" class="btn btn-primary btn-sm" onclick="verifica();" <%=HabilitaBotao%>>
                            <i class="bi bi-check-lg me-1"></i>Confirmar Alteração
                        </button>
                    </div>

                </form>

            </div>
        </div>

    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"
            integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI"
            crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>

    <script>
        function verifica() {
            document.getElementById('overlay-loader').classList.add('active');
            document.form.submit();
        }

        function RetornaPagina() {
            const url = window.location.protocol + '//' +
                        window.location.hostname + '/corporativo/frameset/sfa/pedidos/alteracoes/AlteraPedido02.asp';
            window.location.href = url;
        }

        function doOnLoad() {
            document.getElementById('overlay-loader').classList.remove('active');
        }
    </script>

</body>
<%
function Aviso( msg )

    mostraAviso = "" &_
        "<div class='d-flex justify-content-center align-items-center' style='min-height:200px;'>" &_
            "<div style='max-width:480px; width:100%; background-color:hsl(3,100%,96%); border:1px solid hsl(3,80%,85%); border-radius:10px; padding:1.5rem 1.75rem; text-align:center;'>" &_
                "<i class='bi bi-exclamation-triangle-fill' style='font-size:1.5rem; color:hsl(3,90%,58%); display:block; margin-bottom:.75rem;'></i>" &_
                "<p style='font-weight:600; font-size:.9rem; color:hsl(3,90%,40%); margin-bottom:.5rem;'>Aviso</p>" &_
                "<p style='font-size:.875rem; color:hsl(240,6%,10%); margin-bottom:1rem;'>" & msg & "</p>" &_
                "<a href='javascript:window.history.back();' style='display:inline-flex; align-items:center; gap:.35rem; font-size:.875rem; font-weight:500; color:hsl(3,90%,58%); text-decoration:none;'>" &_
                    "<i class='bi bi-arrow-left'></i> Voltar" &_
                "</a>" &_
            "</div>" &_
        "</div>"

    response.write mostraAviso

end function
%>
</html>
