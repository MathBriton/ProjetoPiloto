<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Gravar Status</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" />
    <!-- <link href="../Emissoes/style.css" rel="stylesheet" type="text/css" /> -->
    <!--#include file="../../includes/variaveis.asp" -->
    <!--#include file="../../funcoes/funcoes.asp" -->
    <!--#include file="../../cnn2/cnn2.asp" -->
    <style>
        body {
            background-color: #f5f5f5;
            font-family: Arial, Verdana, Helvetica, sans-serif;
            font-size: 13px;
        }

        .alerta-wrapper {
            display: flex;
            justify-content: center;
            align-items: flex-start;
            min-height: 100vh;
            padding: 40px 16px;
        }

        .alerta-card {
            width: 100%;
            max-width: 480px;
            border-radius: 6px;
            overflow: hidden;
            box-shadow: 0 4px 16px rgba(0,0,0,.15);
        }

        .alerta-header {
            background-color: #dc3545;
            color: #fff;
            font-weight: 700;
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: .5px;
            padding: 10px 16px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .alerta-body {
            background-color: #fff;
            padding: 16px 20px;
            font-size: 12px;
            color: #333;
            line-height: 1.6;
            text-align: center;
        }

        .alerta-body a {
            display: inline-block;
            margin-top: 12px;
            padding: 6px 18px;
            background-color: #dc3545;
            color: #fff;
            font-size: 12px;
            font-weight: 600;
            border-radius: 4px;
            text-decoration: none;
        }

        .alerta-body a:hover {
            background-color: #a71d2a;
            color: #fff;
        }
    </style>
</head>
<body>
<%
'==========================================================================
' HELPER — substitui os blocos de tabela vermelha espalhados pelo arquivo.
' Uso: MostraErro "Mensagem aqui.<br>Clique <a ...>aqui</a> para voltar."
'==========================================================================
Sub MostraErro(msg)
    Response.Write "<div class=""alerta-wrapper"">"
    Response.Write "  <div class=""alerta-card"">"
    Response.Write "    <div class=""alerta-header"">"
    Response.Write "      <i class=""bi bi-exclamation-triangle-fill""></i> Aten&ccedil;&atilde;o"
    Response.Write "    </div>"
    Response.Write "    <div class=""alerta-body"">" & msg & "</div>"
    Response.Write "  </div>"
    Response.Write "</div>"
    Response.Write "<script src=""https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js""></script>"
End Sub

Session.LCID = 1046
PROD = SESSION("PROD")
If Session("Nivel") > 0 Then
    MostraErro "Voc&ecirc; n&atilde;o tem permiss&atilde;o para acessar esta fun&ccedil;&atilde;o.<br><a href=""#"" onclick=""history.back(1);"">Voltar</a>"
    Response.End
End If

'Set rsPARAM = Server.CreateObject("ADODB.Recordset")
'VarSQLPARAM = "SELECT Dias, ISNULL(DeMaleta,'') DeMaleta, ISNULL(AteMaleta,'') AteMaleta, ISNULL(DeCorteVinco,'') DeCorteVinco, ISNULL(AteCorteVinco,'') AteCorteVinco, ISNULL(DeChapa,'') DeChapa, ISNULL(AteChapa,'') AteChapa, ISNULL(DePapel,'') DePapel, DiasViagem,DiaUtil,DiasBloqueiaP1, isnull(BloqueiaMaquinaChapa,0) BloqueiaMaquinaChapa, isnull(BloqueiaMaquinaCaixa,0) BloqueiaMaquinaCaixa, isnull(BloqueiaMaquinaCorteVinco,0) BloqueiaMaquinaCorteVinco,CargaMaquina, DiasAcrescimoPA, DiasAcrescimoCX, DiasAcrescimoCH, PesoMN, PesoCV, PesoCH   "
'VarSQLPARAM = VarSQLPARAM&"FROM CORPORATIVO..parametrosPedidos "
'VarSQLPARAM = VarSQLPARAM&"WHERE id = 1"
'Set rsPARAM = Conexao2.Execute(VarSQLPARAM)

VarSQLPARAM = "SELECT Dias, ISNULL(DeMaleta,'') DeMaleta, ISNULL(AteMaleta,'') AteMaleta, ISNULL(DeCorteVinco,'') DeCorteVinco, ISNULL(AteCorteVinco,'') AteCorteVinco, ISNULL(DeChapa,'') DeChapa, ISNULL(AteChapa,'') AteChapa, ISNULL(DePapel,'') DePapel, DiasViagem,DiaUtil,DiasBloqueiaP1, isnull(BloqueiaMaquinaChapa,0) BloqueiaMaquinaChapa, isnull(BloqueiaMaquinaCaixa,0) BloqueiaMaquinaCaixa, isnull(BloqueiaMaquinaCorteVinco,0) BloqueiaMaquinaCorteVinco,CargaMaquina, DiasAcrescimoPA, DiasAcrescimoCX, DiasAcrescimoCH, PesoMN, PesoCV, PesoCH   "
'VarSQLPARAM = VarSQLPARAM&"FROM CORPORATIVOHMG..parametrosPedidos "
VarSQLPARAM = VarSQLPARAM&"FROM CORPORATIVO..PARAMETROS_PEDIDO "
VarSQLPARAM = VarSQLPARAM&"WHERE id = 1"

SET rsPARAM = Conecta("PIR-CORPORATIVO", VarSQLPARAM)
if not isobject(rsPARAM) or (rsPARAM IS NOTHING) then
      response.end
end if

if not rsPARAM.eof then
    paramDias = rsPARAM("Dias")
    paramDiaUtil = rsPARAM("DiaUtil")
    paramDiasBloqueioP1 = rsPARAM("DiasBloqueiaP1")
    paramDiasViagem = rsPARAM("DiasViagem")
    paramCargaMaquina = rsPARAM("CargaMaquina")
    paramDiasAcrescimoPA = rsPARAM("DiasAcrescimoPA")
    paramDiasAcrescimoCX = rsPARAM("DiasAcrescimoCX")
    paramDiasAcrescimoCH = rsPARAM("DiasAcrescimoCH") 
    paramPesoMN = rsPARAM("PesoMN")
    paramPesoCV = rsPARAM("PesoCV")
    paramPesoCH = rsPARAM("PesoCH")
    paramBloqueiaMaquinaChapa = rsPARAM("BloqueiaMaquinaChapa")
    paramBloqueiaMaquinaCaixa = rsPARAM("BloqueiaMaquinaCaixa")
    paramBloqueiaMaquinaCaixaCorteVinco = rsPARAM("BloqueiaMaquinaCorteVinco")
    dataBloqueiaMaquinaMaletaDe = rsPARAM("DeMaleta")
    dataBloqueiaMaquinaMaletaAte = rsPARAM("AteMaleta")
    dataBloqueiaMaquinaCorteVincoDe = rsPARAM("DeCorteVinco")
    dataBloqueiaMaquinaCorteVincoAte = rsPARAM("AteCorteVinco")
    dataBloqueiaMaquinaChapaDe = rsPARAM("DeChapa")
    dataBloqueiaMaquinaChapaAte = rsPARAM("AteChapa")
else
    paramDias = 2
    paramDiaUtil = "True"
    paramDiasBloqueioP1 = "True"
    paramDiasViagem = "True"
    paramCargaMaquina = "True"
    paramDiasAcrescimoPA = 0
    paramDiasAcrescimoCX = 0
    paramDiasAcrescimoCH = 0
    paramPesoMN = 0
    paramPesoCV = 0
    paramPesoCH = 0
    paramBloqueiaMaquinaChapa = "False"
    paramBloqueiaMaquinaCaixa = "False"
    dataBloqueiaMaquinaMaletaDe = ""
    dataBloqueiaMaquinaMaletaAte = ""
    dataBloqueiaMaquinaCorteVincoDe = ""
    dataBloqueiaMaquinaCorteVincoAte = ""
    dataBloqueiaMaquinaChapaDe = ""
    dataBloqueiaMaquinaChapaAte = ""
end if
rsPARAM.Close
Set rsPARAM = Nothing
ExisteExcessaoCargaMaquina = 0
Status = request("radiobutton")
NumPed = session("pedido")
Session.Contents.Remove("NumPed")

If Status = "PR" then
	CodBloq  = 17
    DescBloq = "Preco menor que do Pedido anterior"
    Mens = "PEDIDO " & NumPed & " BLOQUEADO POR PREÇO !!"
	RnaResultado = "-1"
end if

if Status = "DE" then
	CodBloq  = 35
    DescBloq = "Data de Entrega" 
    Mens = "PEDIDO " & NumPed & " BLOQUEADO POR DATA DE ENTREGA !!"
	RnaResultado = "-1"
end if

if Status = "CP" then
	CodBloq  = 30
    DescBloq = "Condição de Pagamento" 
    Mens = "PEDIDO " & NumPed & " BLOQUEADO POR COND. PAGTO !!"
	RnaResultado = "-1"
end if

if Status = "LI" then
	CodBloq  = " "
    DescBloq = " "
    Mens = "PEDIDO " & NumPed & " LIBERADO !!"
	RnaResultado = "1"
end if

if Status = "LC" then
	CodBloq  = 40
    DescBloq = "Bloqueado por Limite de Crédito" 
    Mens = "CLIENTE BLOQUEADO POR LIMITE DE CRÉDITO !!"
	RnaResultado = "-1"
end if

PR_PEDIDO = Request.Form("PRPEDIDO")
ULTIMOPRECO = Request.Form("ULTIMOPRECO")
VPESOTOTAL = Request.Form("VPESOTOTAL")
TIPOPRODUTO = Request.Form("TIPOPRODUTO")
quantidade = Trim(Request.Form("Quantidade"))
largura = Trim(Request.Form("Largura"))
comprimento = Trim(Request.Form("Comprimento"))
TES = Request("Tes")

IF TRIM(TIPOPRODUTO) = "" THEN
	TIPOPRODUTO = "903"
END IF

ESTRUTURA = Request.Form("ESTRUTURA")
LOJACLI = TRIM(REQUEST.FORM("LOJACLI"))
CODCLI = TRIM(REQUEST.FORM("CODCLI"))
LOJACLI2 = TRIM(REQUEST.FORM("LOJACLI2"))
CODCLI2 = TRIM(REQUEST.FORM("CODCLI2"))

IF trim(Status) = "" then
	MostraErro "Op&ccedil;&atilde;o de valida&ccedil;&atilde;o n&atilde;o selecionada.<br>Selecione uma das op&ccedil;&otilde;es dispon&iacute;veis.<br><a href=""#"" onMouseDown=""history.go(-1);"">Voltar</a>"
	Response.End 
end if

IF (PR_PEDIDO < ULTIMOPRECO) AND (Usr = "adriana") THEN
	MostraErro "VOC&Ecirc; N&Atilde;O TEM PERMISS&Atilde;O PARA LIBERAR ESTE PEDIDO.<br>Pedido com o pre&ccedil;o menor que o pedido anterior.<br>Pre&ccedil;o deste pedido: " & PR_PEDIDO & "<br>Pre&ccedil;o do pedido anterior: " & ULTIMOPRECO & "<br><a href=""#"" onMouseDown=""history.go(-1);"">Voltar</a>"
	Response.End 
END IF

IF TRIM(CODCLI) <> "0893" AND TRIM(CODCLI) <> "1684" AND TRIM(CODCLI) <> "0005" AND TRIM(CODCLI) <> "1732" AND TRIM(CODCLI) <> "1733" AND TRIM(CODCLI) <> "1131" AND TRIM(CODCLI) <> "1505" AND TRIM(CODCLI) <> "0004" AND TRIM(CODCLI) <> "0059" AND TRIM(CODCLI) <> "1013" AND TRIM(CODCLI) <> "0001" AND TRIM(CODCLI) <> "1629" AND TRIM(CODCLI) <> "3265"AND TRIM(CODCLI) <> "0211" AND TRIM(CODCLI) <> "0211" AND TRIM(CODCLI) <> "0002" AND TRIM(CODCLI) <> "0620" AND TRIM(CODCLI) <> "3470" AND TRIM(CODCLI) <> "0571" AND TRIM(CODCLI) <> "3786" AND TRIM(CODCLI) <> "3782" AND TRIM(CODCLI) <> "0014" THEN  
	IF (PRECOMENOR <> "S") AND (PR_PEDIDO < ULTIMOPRECO) AND (ULTIMOPRECO > 0) AND (Status = "LI") THEN
		MostraErro "VOC&Ecirc; N&Atilde;O TEM PERMISS&Atilde;O PARA LIBERAR ESTE PEDIDO.<br>Pedido com o pre&ccedil;o menor que o pedido anterior.<br>Pre&ccedil;o deste pedido: " & PR_PEDIDO & "<br>Pre&ccedil;o do pedido anterior: " & ULTIMOPRECO & "<br><a href=""#"" onMouseDown=""history.go(-1);"">Voltar</a>"
		Response.End 
	END IF

	'BLOQUEANDO POR PESO DE CHAPA ONDA BC (ONDA DUPLA) MENOR QUE O DEFINIDO PELO CARLOS SOBRINHO
	IF (PESOMENOR <> "S") AND (TIPOPRODUTO="903") AND (ESTRUTURA="SIMPLES") AND (VPESOTOTAL < 250) AND (Status = "LI") THEN
		MostraErro "VOC&Ecirc; N&Atilde;O TEM PERMISS&Atilde;O PARA LIBERAR ESTE PEDIDO.<br>Peso total do pedido inferior que o m&iacute;nimo permitido.<br>Peso do pedido: " & VPESOTOTAL & "<br>Peso m&iacute;nimo permitido: 250Kg<br><a href=""#"" onMouseDown=""history.go(-1);"">Voltar</a>"
		Response.End 
	END IF

	'BLOQUEANDO POR PESO DE CHAPA ONDA B OU C (ONDA SIMPLES) MENOR QUE O DEFINIDO PELO CARLOS SOBRINHO
	IF (PESOMENOR <> "S") AND (TIPOPRODUTO="903") AND (ESTRUTURA="DUPLA") AND (VPESOTOTAL < paramPesoCH) AND (Status = "LI") THEN
		MostraErro "VOC&Ecirc; N&Atilde;O TEM PERMISS&Atilde;O PARA LIBERAR ESTE PEDIDO.<br>Peso total do pedido inferior que o m&iacute;nimo permitido.<br>Peso do pedido: " & VPESOTOTAL & "<br>Peso m&iacute;nimo permitido: 500Kg<br><a href=""#"" onMouseDown=""history.go(-1);"">Voltar</a>"
		Response.End 
	END IF

	'BLOQUEANDO POR PESO DE CAIXAS MALETA NORMAL MENOR QUE O DEFINIDO PELO CARLOS SOBRINHO
	IF (PESOMENOR <> "S") AND (TIPOPRODUTO="901") AND (VPESOTOTAL < paramPesoCV) AND (Status = "LI") THEN
		MostraErro "VOC&Ecirc; N&Atilde;O TEM PERMISS&Atilde;O PARA LIBERAR ESTE PEDIDO.<br>Peso total do pedido inferior que o m&iacute;nimo permitido.<br>Peso do pedido: " & VPESOTOTAL & "<br>Peso m&iacute;nimo permitido: 1.200Kg<br><a href=""#"" onMouseDown=""history.go(-1);"">Voltar</a>"
		Response.End 
	END IF

	'BLOQUEANDO POR PESO DE CAIXAS CORTE E VINCO MENOR QUE O DEFINIDO PELO CARLOS SOBRINHO
	IF (PESOMENOR <> "S") AND (TIPOPRODUTO="902") AND (VPESOTOTAL < paramPesoMN) AND (Status = "LI") THEN
		MostraErro "VOC&Ecirc; N&Atilde;O TEM PERMISS&Atilde;O PARA LIBERAR ESTE PEDIDO.<br>Peso total do pedido inferior que o m&iacute;nimo permitido.<br>Peso do pedido: " & VPESOTOTAL & "<br>Peso m&iacute;nimo permitido: 2.500Kg<br><a href=""#"" onMouseDown=""history.go(-1);"">Voltar</a>"
		Response.End 
	END IF
END IF

If (Usr <> "CARLOSS" And Usr <> "sandro" And Usr <> "Adriana.Bazzarello" And Usr <> "Eduardo") AND (Tes = "580" Or Tes="581" Or Tes="600" Or Tes="601" Or Tes="715" Or Tes="716" Or Tes="717" Or Tes="718" Or Tes="710") Then 
	MostraErro "VOC&Ecirc; N&Atilde;O TEM PERMISS&Atilde;O PARA LIBERAR ESTE PEDIDO.<br>TES: " & TES & "<br>Trata-se de um pedido livre de d&eacute;bito.<br><a href=""#"" onMouseDown=""history.go(-1);"">Voltar</a>"
	Response.End 
End If		

'CHAMA DATA DA FUNÇÃO INVERTEDATASIGA AAAAMMDD E HORA DO TIME LCID 1046
Hoje = Trim(InverteDataSiga(Date()))
Hora = Trim(Time())

'.......................................................................................................
'VERIFICANDO SE O CLIENTE ESTÁ BLOQUEADO

'RECUPERANDO A CÓDIGO DO CLIENTE
' Set rsRIS = Server.CreateObject ("ADODB.Recordset")
' VarSQLRIS = "Execute SP_RISCO ["&CodCli&"] "
' set rsRIS = Conexao2.Execute(VarSQLRIS)

VarSQLRIS = "Execute SP_RISCO ["&CodCli&"] "

SET rsRIS = Conecta("PIR-DADOSADV", VarSQLRIS)
if not isobject(rsRIS) or (rsRIS IS NOTHING) then
      response.end
end if
if NOT rsRIS.EOF then
	MostraErro "Cliente bloqueado pelo Departamento Financeiro!<br>Motivo: Risco E!<br><a href=""#"" onMouseDown=""history.go(-1);"">Voltar</a>"
	rsRIS.Close
	Set rsRIS = Nothing
	Response.End
End if

'Fechando o objeto
rsRIS.Close
Set rsRIS = Nothing

'*************************************************************************************
'VERIFICANDO SE O PRODUTO PERTENCE AO CLIENTE
'*************************************************************************************
'RECUPERANDO A CÓDIGO DO CLIENTE
Dim rsProdCli, VarSQLrsProdCli

' Set rsProdCli = Server.CreateObject ("ADODB.Recordset")
' VarSQLrsProdCli = "Execute SP_VERIFICA_PRODUTO_PERTENCE_CLIENTE '" & CodCli & "', '" & PROD & "'"
' set rsProdCli = Conexao2.Execute(VarSQLrsProdCli)

VarSQLrsProdCli = "Execute SP_VERIFICA_PRODUTO_PERTENCE_CLIENTE '" & CodCli & "', '" & PROD & "'"

SET rsProdCli = Conecta("PIR-DADOSADV", VarSQLrsProdCli)
if not isobject(rsProdCli) or (rsProdCli IS NOTHING) then
      response.end
end if

if uCase(rsProdCli("RETORNO")) = "NAO" then
	MostraErro "Produto n&atilde;o pertence ao cliente.<br><a href=""#"" onMouseDown=""history.go(-1);"">Voltar</a>"
	rsProdCli.Close
	Set rsProdCli = Nothing
	Response.End
End if

'Fechando o objeto
rsProdCli.Close
Set rsProdCli = Nothing

DAT1 = SESSION("DAT1")
DAT2 = SESSION("DAT2")

EntregaSabadoDomingo = Request.Form("SD")
dEntreg = DAT1
dEntreg1 = Mid(dEntreg,7,2) &"/"& Mid(dEntreg,5,2) &"/"& Mid(dEntreg,1,4)

select case TIPOPRODUTO
    case "900"
        condicao = "PA"
		paramDias = paramDiasAcrescimoPA
    case "903"
        condicao = "CH"
		paramDias = paramDiasAcrescimoCH
    case Else
        condicao = "CX"
	    paramDias = paramDiasAcrescimoCX 
end select

IF LEN(DAT2) < 5 THEN
	DAT2 = DAT1
END IF
'response.write DiaUtil(Date(),dEntreg1)
'response.write paramDias
'response.write paramDiaUtil
if paramDiaUtil = "True" or paramDiaUtil = "Verdadeiro" or paramDiaUtil = "VERDADEIO" Then
    if paramDias > DiaUtil(Date(),dEntreg1) then
		MostraErro "O pedido n&atilde;o pode ser emitido pois a data esta fora do limite m&iacute;nimo permitido.<br><a href=""#"" onMouseDown=""history.go(-1);"">Voltar</a>"
		Response.End
    end if
else
    if paramDias > datediff("d",Date(),dEntreg1) then
		MostraErro "O pedido n&atilde;o pode ser emitido pois a data esta fora do limite m&iacute;nimo permitido.<br><a href=""#"" onMouseDown=""history.go(-1);"">Voltar</a>"
		Response.End
    end if
end if

' Validando o pedido
' Set rsPEDIDO = Server.CreateObject ("ADODB.Recordset")
' VarSQLPED = "Execute SP_LIBERAPEDIDO ["&NumPed&"],["&Status&"],["&CodBloq&"],["&DescBloq&"],["&USR&"],["&Hoje&"],["&Hora&"]"
' set rsPEDIDO = Conexao2.Execute(VarSQLPED)
' Set rsPEDIDO = Nothing

VarSQLPED = "Execute SP_LIBERAPEDIDO ["&NumPed&"],["&Status&"],["&CodBloq&"],["&DescBloq&"],["&USR&"],["&Hoje&"],["&Hora&"]"

SET rsPEDIDO = Conecta("PIR-DADOSADV", VarSQLPED)
if not isobject(rsPEDIDO) or (rsPEDIDO IS NOTHING) then
      response.end
end if

' Criando a Proposta de venda
' Set rsPedidoProposta = Server.CreateObject ("ADODB.Recordset")
' sqlPedidoProposta = "Execute SP_CONLIBERASZF '" & NumPed &"' "
' set rsPedidoProposta = Conexao2.Execute(sqlPedidoProposta)

sqlPedidoProposta = "Execute SP_CONLIBERASZF '" & NumPed &"' "
SET rsPedidoProposta = Conecta("PIR-DADOSADV", sqlPedidoProposta)
if not isobject(rsPedidoProposta) or (rsPedidoProposta IS NOTHING) then
      response.end
end if

if Not rsPedidoProposta.eof and Trim(Request("IncluiProposta")) = "1" then
    msgPesoProposta = "O Pedido foi validado com exito.\nPorém não foi possivel identificar o tipo desta Caixa, por este motivo não criamos a proposta."
    If rsPedidoProposta("ZF_DESCNF") = "901" then
        pesoProposta = 1200
        msgPesoProposta = ""
    end if
    If rsPedidoProposta("ZF_DESCNF") = "902" then
        pesoProposta = 2500
        msgPesoProposta = ""
    end if

    if msgPesoProposta = "" then
        ' Set rsProposta = Server.CreateObject ("ADODB.Recordset")
        ' sqlValidaProposta = "SELECT COUNT(*) EXISTE FROM CORPORATIVO..Proposta WHERE Ativa = 1 AND Produto = '"&Trim(rsPedidoProposta("ZF_PRODUTO"))&"' AND Cliente = '"&Trim(rsPedidoProposta("ZF_CLIENTE"))&"' AND Loja = '"&Trim(rsPedidoProposta("ZF_LOJACLI"))&"' "
        ' set rsValidaProposta = Conexao2.Execute(sqlValidaProposta)

		sqlValidaProposta = "SELECT COUNT(*) EXISTE FROM CORPORATIVO..Proposta WHERE Ativa = 1 AND Produto = '"&Trim(rsPedidoProposta("ZF_PRODUTO"))&"' AND Cliente = '"&Trim(rsPedidoProposta("ZF_CLIENTE"))&"' AND Loja = '"&Trim(rsPedidoProposta("ZF_LOJACLI"))&"' "
		SET rsValidaProposta = Conecta("PIR-CORPORATIVO", sqlValidaProposta)
		if not isobject(rsValidaProposta) or (rsValidaProposta IS NOTHING) then
			response.end
		end if
        if Cint(rsValidaProposta("EXISTE")) = 0 then
            'Set rsProposta = Server.CreateObject("ADODB.Recordset")
            sqlProposta = "INSERT INTO CORPORATIVO..Proposta(IdProposta, De, Ate, Produto, Estrutura, Tipo, Cliente, Loja, Pallet, Condicao, TipoFrete, Preco, Ativa, Usuario, DataCadastro, HoraCadastro, Quantidade) Values"
            sqlProposta = sqlProposta &"(ISNULL((SELECT MAX(IdProposta) FROM CORPORATIVO..Proposta),0)+1, '"&InverteData(Date())&"', '"&InverteData(Dateadd("M",6,Date()))&"', "
            sqlProposta = sqlProposta &"'"&Trim(rsPedidoProposta("ZF_PRODUTO"))&"', '"&Trim(rsPedidoProposta("ZF_ESTRUTU"))&"', 'CX', "
            sqlProposta = sqlProposta &"'"&Trim(rsPedidoProposta("ZF_CLIENTE"))&"', '"&Trim(rsPedidoProposta("ZF_LOJACLI"))&"', 'S', "
            sqlProposta = sqlProposta &"'"&Trim(rsPedidoProposta("ZF_CONDPAG"))&"', '"&Trim(rsPedidoProposta("ZF_TPFRETE"))&"', "
            sqlProposta = sqlProposta &""&replace(Trim(rsPedidoProposta("ZF_PRECO")),",",".")&", 1, '"&Trim(USR)&" - Validação', "
            sqlProposta = sqlProposta &"'"&InverteData(Date())&"', '"&Mid(Time,1,8)&"', 0) "
            'set rsProposta = Conexao2.Execute(sqlProposta)
            'set rsProposta = Nothing
		    SET rsProposta = Conecta("PIR-CORPORATIVO", sqlProposta)
		    if not isobject(rsProposta) or (rsProposta IS NOTHING) then
     			response.end
	    	end if

            ' Set rsProposta = Server.CreateObject ("ADODB.Recordset")
            ' sqlProposta = "SELECT IdProposta FROM CORPORATIVO..Proposta WHERE Ativa = 1 AND Produto = '"&Trim(rsPedidoProposta("ZF_PRODUTO"))&"' AND Cliente = '"&Trim(rsPedidoProposta("ZF_CLIENTE"))&"' AND Loja = '"&Trim(rsPedidoProposta("ZF_LOJACLI"))&"' AND Condicao = '"&Trim(rsPedidoProposta("ZF_CONDPAG"))&"' AND TipoFrete = '"&Trim(rsPedidoProposta("ZF_TPFRETE"))&"' AND De = '"&InverteData(Date())&"' AND Ate = '"&InverteData(Dateadd("M",6,Date()))&"' AND Preco = "&Replace(Trim(rsPedidoProposta("ZF_PRECO")),",",".")
            ' set rsProposta = Conexao2.Execute(sqlProposta)

			sqlProposta = "SELECT IdProposta FROM CORPORATIVO..Proposta WHERE Ativa = 1 AND Produto = '"&Trim(rsPedidoProposta("ZF_PRODUTO"))&"' AND Cliente = '"&Trim(rsPedidoProposta("ZF_CLIENTE"))&"' AND Loja = '"&Trim(rsPedidoProposta("ZF_LOJACLI"))&"' AND Condicao = '"&Trim(rsPedidoProposta("ZF_CONDPAG"))&"' AND TipoFrete = '"&Trim(rsPedidoProposta("ZF_TPFRETE"))&"' AND De = '"&InverteData(Date())&"' AND Ate = '"&InverteData(Dateadd("M",6,Date()))&"' AND Preco = "&Replace(Trim(rsPedidoProposta("ZF_PRECO")),",",".")
		    SET rsProposta = Conecta("PIR-CORPORATIVO", sqlProposta)
		    if not isobject(rsProposta) or (rsProposta IS NOTHING) then
     			response.end
	    	end if

            if not rsProposta.eof  then
                Response.write "<script>alert('Proposta "&rsProposta("IdProposta")&" incluída com sucesso!')</script>"
            else
                Response.write "<script>alert('Houve algum problema e a proposta não pode ser inserida.')</script>"
            end if
            rsProposta.Close
            Set rsProposta = Nothing
        else
            Response.write "<script>alert('Existe um proposta ativa para este cliente, produto e loja\nA Proposta não foi inserida.')</script>"
        end if
        rsValidaProposta.close
        set rsValidaProposta = Nothing
    Else
        Response.write "<script>alert('"&msgPesoProposta&"')</script>"
    End if
end if
rsPedidoProposta.Close
Set rsPedidoProposta = Nothing

'Conexao2.Close
'Set Conexao2 = Nothing

response.Redirect "../validacao/ListaPedidosAValidar.asp"

Dim Ano, Mes, Dia
Function InverteDataSiga(InverteData)
  Ano = Year(InverteData)
  Mes = Month(InverteData)
  Dia = Day(InverteData)
  If Len(Ano) > 4 Then
    Ano = "20" & Ano
  End If
  If Len(Mes) =1 Then
    Mes = "0" & Mes
  End If
  If Len(Dia) = 1 Then
    Dia = "0" & Dia
  End If
  InverteDataSiga = Ano & Mes & Dia
End Function

dim data_inici,data_fina
Function DiaUtil(data_inici,data_fina)
    Dim difer,nume,week,contad
	difer = DateDiff("d",data_inici,data_fina)
	nume = Weekday(data_inici)
    if nume = 1 or nume = 7 then nume = nume + 1
	'For contad = 0 to difer
		'If nume = 1 Then week = week + 1
		'If nume = 7 Then week = week + 1
			'nume = nume + 1
		'If nume > 7 Then nume = 1
	'Next
	'DiaUtil = difer - week
	DiaUtil = difer
End Function

%>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
