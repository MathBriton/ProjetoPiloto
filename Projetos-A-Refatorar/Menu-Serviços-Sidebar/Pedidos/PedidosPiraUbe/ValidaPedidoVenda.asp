<!DOCTYPE html>
<!--#include file="../../cnn/cnn.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->
<!--#include file="../../includes/variaveis.asp" -->
<html lang="pt-br">
  <head>
    <meta charset="utf-8" />
    <title></title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <style>
      body {margin:50px;}
    </style>
  </head>
  <body>
<%

SESSION.LCID=1046

codsiga   = coduser
VLARG     = 0
VCOMP     = 0
VVINCO_01 = 0
VVINCO_02 = 0
VVINCO_03 = 0
VVINCO_04 = 0
VVINCO_05 = 0
COMPAP    = 0
GRAM      = 0
FORMATO   = 0
DIAM      = 0
TUBETE    = 0
TRANSPA   = " "
Frete_Unit= 0
eShop = Trim(Request("eshop"))

if Trim(Request("FreteUnitario")) <> "" and Trim(Request("FreteUnitario")) <> "0" then

  Frete_Unit = Replace(Trim(Request("FreteUnitario")),",",".")

end if

PalletCH = " "
ExisteDisponibilidadeMaquina = 0
DataDisponibilidadeCargaMaquina = ""
ExisteExcessaoCargaMaquina = 0

condicao            = Trim(Request("condicao"))
session("condicao") = condicao
filial              = Trim(Request("hidPlanta"))
session("filial")   = Trim(Request("hidPlanta"))
codcli              = Request("codcli")
lojacli             = Request("lojacli")
'tipo_frete          = "F"
'tipo_frete          = iif( (Trim(tipo_frete) = ""), Trim(tipoFrete), Trim(Request("tipoFrete")) )
tipo_frete=Trim(Request("tipo_frete"))
tipoFrete = Trim(Request("tipoFrete"))

if Trim(tipo_frete) = "" then
  tipo_frete = Trim(tipoFrete)
end if

cond_pgto           = Request("cond_pgto")
condicaoPagamento   = Request("CondicaoPagamento")

if Trim(cond_pgto) = "" then

  cond_pgto = condicaoPagamento

end if

moedaExportacao = Trim(Request("MoedaExportacao"))

if Trim(moedaExportacao) = "" then

  moedaExportacao = "1"

end if

desconto        = Request("desconto")
tipo_nf         = Request("enable")
codcli2         = Request("codcli2")
lojacli2        = Request("lojacli2")
enviaemailcli   = Request("enviaemailcli")
descarregamento = trim(request("descarregamento"))
horaDescarga    = trim(request("horaDescarga"))
dataDescarga    = trim(request("dataDescarga"))
obs_venda       = Request("obs_venda")
NumFnc          = Request("NumFnc")
produto         = Trim(Request("produto"))
quantidade      = Request("quantidade")
preco_unit      = Trim(replace(Request("preco_unit"),".",","))
tes             = Request("tes")
DATA1           = Request("data1")
DATA2           = Request("data2")
data5           = Request("data5")
ped_cliente     = Trim(RemoveAcentos(Request("ped_cliente")))
ItemPedCli      = Trim(Request("itempedcli"))
nf_pallet       = Request("nf_pallet")
CODFAT          = Request("codfat")
CODFAT2         = Request("codfat2")
Estrutura       = Trim(Request("estrutura"))

vlarg = iif( (Trim(Request("LARG")) = ""), 0, Request("LARG") )
VCOMP = iif( (Trim(Request("COMP")) = ""), 0, Request("COMP") )
larguraMaxima = request("hidLarguraProduto")
comprimentoMaximo = request("hidComprimentoMax")

if condicao="CH" then

	VVINCO_01=Request("VINCO_01")
	VVINCO_02=Request("VINCO_02")
	VVINCO_03=Request("VINCO_03")
	VVINCO_04=Request("VINCO_04")
	VVINCO_05=Request("VINCO_05")

	if vvinco_01="" then vvinco_01=0 end if
	if vvinco_02="" then vvinco_02=0 end if
	if vvinco_03="" then vvinco_03=0 end if
	if vvinco_04="" then vvinco_04=0 end if
	if vvinco_05="" then vvinco_05=0 end if

  PalletCH = Request("pallet")

end if

EntregaSabadoDomingo = iif( (Trim(Request("SD")) = "on"), 1, 0 )

if obs_venda="" then
	obs_venda=" "
end if

if NumFnc <> "" then
  obs_venda= " FNC: " & NumFnc
end if  

if ped_cliente = "" then
  ItemPedCli = ""
else
  if ItemPedCli <> "" then
     ItemPedCli = replace(replace(ItemPedCli,".",""),",","")
  end if
end if

CodRepre = iif( (Request("CodRepre") <> "" and Trim(Request("condicao")) = "PA"), Request("CodRepre"), " " )
pedmae   = iif( (request("PM") = "on"), "S", " " )

if Trim(Request("MsgNota")) <> "" then
  MsgNota = UCase(Trim(Request("MsgNota")))
  MsgNota = RemoveAcentos(MsgNota)
else
  MsgNota = " "
end if

if Trim(Request("MsgNota2")) <> "" then
  MsgNota2 = UCase(Trim(Request("MsgNota2")))
  MsgNota2 = RemoveAcentos(MsgNota2)
else
  MsgNota2 = " "
end if

if Trim(Request("MsgNota3")) <> "" then
  MsgNota3 = UCase(Trim(Request("MsgNota3")))
  MsgNota3 = RemoveAcentos(MsgNota3)
else
  MsgNota3 = " "
end if

if Trim(Request("MsgNota4")) <> "" then
  MsgNota4 = UCase(Trim(Request("MsgNota4")))
  MsgNota4 = RemoveAcentos(MsgNota4)
else
  MsgNota4 = " "
end if

obsLotePiloto = Request("obs_lotepiloto")
lotePiloto = iif( (Trim(Request("LotePiloto")) = "on"), "S", "N" )

'RESPONSE.WRITE "Request Lote Piloto  " & Request("LotePiloto")  & " lotePiloto redefinido  " & lotePiloto
session("refaturamento")  = iif( (Trim(Request("checkRefaturamento")) = "on"), "R", "N" )
'response.end 

if lotePiloto = "S" and eShop = "1" then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>eShop!</strong> Não é possível emitir um lote piloto pra pedidos eShop. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
  response.end
end if

if Nivel <> 0 and eShop = 1 then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>eShop!</strong> Você não tem permissão para emitir pedidos para eShop. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
  response.end
end if

if moedaExportacao > "1" and eShop = "1" then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>eShop!</strong> Tipo de moeda inválida para pedidos eShop. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
  response.end
end if

if Trim(condicao) = "" then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Tipo de Pedido!</strong> O Tipo de Pedido deve ser informado. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
  response.end
end if

if (Trim(condicao) <> "CX" And Trim(condicao) <> "CH") and eShop = "1" then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>eShop!</strong> Para eShop, por favor selecione ""Tipo de pedido: Caixa"". Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
  response.end
end if

if Cdbl(quantidade) <= 0 then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Quantidade!</strong> A Quantidade deve ser maior que zero. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
  response.end
end if

if Cdbl(quantidade) > 909090 then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Quantidade!</strong> Por favor, altere a quantidade informada, pois ultrapassou o máximo permitido, que é de 909090. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
  response.end
end if

if Trim(DATA1) = "" then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Data Entrega!</strong> A Data de entrega deve ser informada. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
  response.end
end if

if Trim(cond_pgto) = "" then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Condição de pagamento!</strong>Por favor, informe a condição de pagamento. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
	Response.End
end if

if Trim(cond_pgto) = "090" And Len(trim(data5)) <> 10 then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Condição de pagamento!</strong>A condição de pagamento INDICAR A DATA DE VENCIMENTO MANUAL, foi escolhida, porém a data de Vencimento não foi preenchida. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
	Response.End
end if

if Trim(cond_pgto) = "090" And Len(trim(data5)) = 10 then
	if DateDiff("d",vhoje1,trim(data5)) < 1 then
    response.write "<div class='alert alert-danger' role='alert'>"
    response.write "<strong>Data vencimento!</strong> A data de vencimento deve ser superior a 1 dia, Data mínima: "&DateAdd("d",Date(),1)&". Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
    response.write "</div>"
    Response.End
  end if
end if

if Trim(produto) = "" then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Produto!</strong> O Produto deve ser informado. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
  response.end
end if

if Trim(tes) = "" then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>TES!</strong> A TES deve ser informada. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
  response.end
end if

if condicao="PA"  then
  if Trim(Request("data2")) = "" then
    response.write "<div class='alert alert-danger' role='alert'>"
    response.write "<strong>Segunda data!</strong>Por favor informe a segunda data de entrega. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
    response.write "</div>"
    response.end
  end if

	gram      =	mid(Request("gramatura"),1,3) & mid(Request("gramatura"),5,1)
	diam      =	formatnumber(replace(Request("formato2"),".",","),2)
	formato   =	0'replace(Request("formato"),".",",")
	tubete    =	mid(Request("tubete"),1,1)& mid(Request("tubete"),3,1)
	nf_pallet =	"N"

  if dateadd("d",2,Date()) > cDate(DATA1) then
    response.write "<div class='alert alert-danger' role='alert'>"
    response.write "<strong>Data de entrega!</strong> Data de entrega insuficiente. A data de entrega mais proxima disponível é "&dateadd("d",2,Date())&". Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
    response.write "</div>"
    response.end
  end if
else  
  SQL = "SELECT Id " 
  SQL = SQL & " FROM CORPORATIVO..EstruturaRestricao " 
  SQL = SQL & "WHERE ISNULL(Excluido,0) = 0 "
  SQL = SQL &   "AND Estrutura = '" & Estrutura & "' "
  SQL = SQL &   "AND (" & vlarg & " BETWEEN LarguraDe AND LarguraAte) "
  SQL = SQL &   "AND Tipo = '" & condicao & "' "
  Dados1.Open(SQL)
  if not Dados1.eof then
    response.write "<div class='alert alert-danger' role='alert'>"
    response.write "<strong>Estrutura e Largura!</strong> Existe uma restrição para produção desta estrutura "&Estrutura&" na Largura informada "&vlarg&". Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
    response.write "</div>"
    response.end
  end if
  Dados1.Close
end if


SQL = "SELECT Dias, " 
SQL = SQL & "ISNULL(DeMaleta,'') DeMaleta, " 
SQL = SQL & "ISNULL(AteMaleta,'') AteMaleta, " 
SQL = SQL & "ISNULL(DeCorteVinco,'') DeCorteVinco, " 
SQL = SQL & "ISNULL(AteCorteVinco,'') AteCorteVinco, " 
SQL = SQL & "ISNULL(DeChapa,'') DeChapa, " 
SQL = SQL & "ISNULL(AteChapa,'') AteChapa, " 
SQL = SQL & "ISNULL(DePapel,'') DePapel, " 
SQL = SQL & "DiasViagem, " 
SQL = SQL & "DiaUtil, " 
SQL = SQL & "DiasBloqueiaP1, " 
SQL = SQL & "ISNULL(BloqueiaMaquinaChapa,0) BloqueiaMaquinaChapa, " 
SQL = SQL & "ISNULL(BloqueiaMaquinaCaixa,0) BloqueiaMaquinaCaixa, " 
SQL = SQL & "ISNULL(BloqueiaMaquinaCorteVinco,0) BloqueiaMaquinaCorteVinco, " 
SQL = SQL & "CargaMaquina, " 
SQL = SQL & "DiasAcrescimoPA, " 
SQL = SQL & "DiasAcrescimoCX, " 
SQL = SQL & "DiasAcrescimoCH, " 
SQL = SQL & "PesoMN, " 
SQL = SQL & "PesoCV, " 
SQL = SQL & "PesoCH, " 
SQL = SQL & "MtsCVmin, " 
SQL = SQL & "MtsCXmin, " 
SQL = SQL & "MtsCHmin, " 
SQL = SQL & "PrecoKgCorteVinco, " 
SQL = SQL & "PrecoKgMaleta, " 
SQL = SQL & "PrecoKgChapa, " 
SQL = SQL & "EmailAbaixoPrecoKg " 
SQL = SQL & " FROM CORPORATIVO..PARAMETROS_PEDIDO " 
SQL = SQL & " WHERE Planta = " & filial
Dados1.Open(SQL)

if not Dados1.eof then
  paramDias                           = Dados1("Dias")
  paramDiaUtil                        = Dados1("DiaUtil")
  paramDiasBloqueioP1                 = Dados1("DiasBloqueiaP1")
  paramDiasViagem                     = Dados1("DiasViagem")
  paramCargaMaquina                   = Dados1("CargaMaquina")
  paramDiasAcrescimoPA                = Dados1("DiasAcrescimoPA")
  paramDiasAcrescimoCX                = Dados1("DiasAcrescimoCX")
  paramDiasAcrescimoCH                = Dados1("DiasAcrescimoCH") 
  paramPesoMN                         = Dados1("PesoMN")
  paramPesoCV                         = Dados1("PesoCV")
  paramPesoCH                         = Dados1("PesoCH")
  ParamAreaMN                         = Dados1("MtsCXmin")
  ParamAreaCV                         = Dados1("MtsCVmin")
  ParamAreaCH                         = Dados1("MtsCHmin")
  paramBloqueiaMaquinaChapa           = Dados1("BloqueiaMaquinaChapa")
  paramBloqueiaMaquinaCaixa           = Dados1("BloqueiaMaquinaCaixa")
  paramBloqueiaMaquinaCaixaCorteVinco = Dados1("BloqueiaMaquinaCorteVinco")
  dataBloqueiaMaquinaMaletaDe         = Dados1("DeMaleta")
  dataBloqueiaMaquinaMaletaAte        = Dados1("AteMaleta")
  dataBloqueiaMaquinaCorteVincoDe     = Dados1("DeCorteVinco")
  dataBloqueiaMaquinaCorteVincoAte    = Dados1("AteCorteVinco")
  dataBloqueiaMaquinaChapaDe          = Dados1("DeChapa")
  dataBloqueiaMaquinaChapaAte         = Dados1("AteChapa")
  precoKgCorteVinco                   = FormatNumber(trim(Dados1("PrecoKgCorteVinco")),2)
  precoKgMaleta                       = FormatNumber(trim(Dados1("PrecoKgMaleta")),2)
  precoKgChapa                        = FormatNumber(trim(Dados1("PrecoKgChapa")),2)
  emailAbaixoPrecoKg                  = trim(Dados1("EmailAbaixoPrecoKg"))
else
  paramDias                         = 2
  paramDiaUtil                      = "1"
  paramDiasBloqueioP1               = "1"
  paramDiasViagem                   = "1"
  paramCargaMaquina                 = "1"
  paramDiasAcrescimoPA              = 0
  paramDiasAcrescimoCX              = 0
  paramDiasAcrescimoCH              = 0
  paramPesoMN                       = 0
  paramPesoCV                       = 0
  paramPesoCH                       = 0
  ParamAreaMN                       = 0
  ParamAreaCV                       = 0
  ParamAreaCH                       = 0
  paramBloqueiaMaquinaChapa         = "0"
  paramBloqueiaMaquinaCaixa         = "0"
  dataBloqueiaMaquinaMaletaDe       = ""
  dataBloqueiaMaquinaMaletaAte      = ""
  dataBloqueiaMaquinaCorteVincoDe   = ""
  dataBloqueiaMaquinaCorteVincoAte  = ""
  dataBloqueiaMaquinaChapaDe        = ""
  dataBloqueiaMaquinaChapaAte       = ""
  precoKgCorteVinco                 = 0
  precoKgMaleta                     = 0
  precoKgChapa                      = 0
  emailAbaixoPrecoKg                = ""
end if
Dados1.Close

if moedaExportacao > 1 then
  SQL = "SELECT A1_COD FROM SA1010 WHERE A1_COD = '"&codcli&"' AND A1_LOJA = '"&lojacli&"' AND A1_EST = 'EX' AND D_E_L_E_T_ = '' "
  Dados1.Open(SQL)
  if Dados1.Eof then
    response.write "<div class='alert alert-danger' role='alert'>"
    response.write "<strong>Exportação!</strong>Não é possivel inserir moeda de exportação para este pedido. Cliente não configurado para exportação. Entre em contato com o Fiscal. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
    response.write "</div>"
    response.end
  end if
	Dados1.Close
end if

if 	Cint(VCOMP) > larguraMaxima then     
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Largura Maxima!</strong>A largura maxima de chapas permitida é " & larguraMaxima & ". Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
  response.end
end if

if 	Cint(VCOMP) > comprimentoMaximo then     
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Comprimento Maximo!</strong>O comprimento maximo de chapas permitido é " & comprimentoMaximo & ". Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
  response.end
end if

if codcli=codcli2 And lojacli=lojacli2 then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Cliente remessa!</strong>Cliente de remessa deve ser diferente do cliente de destino. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
  response.write "</div>"
	Response.End
end if

if nivel < 2 then 
  Dados1.Open("Execute SP_VALIDACLIENTE ["&nivel&"],["&codsiga&"],["&CODCLI&"],[" &LOJACLI&"]")
  if Dados1.Eof then
    response.write "<div class='alert alert-danger' role='alert'>"
    response.write "<strong>Cliente!</strong>Cliente não encontrado, está bloqueado ou não pertence à sua carteira. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
    response.write "</div>"
		Response.End    
	end if	
	Dados1.Close
end if

dim QuantDig, QuantSug1, QuantSug2

if condicao = "CX" then
  Dados1.Open("Execute SP_VERIFICA_PECAS_POR_PALETE ["&Produto&"]")
  if Dados1("PECAS") > 0 then
    if int(Quantidade / Dados1("PECAS")) <> Quantidade / Dados1("PECAS") then
      QuantSug1 = Int(Quantidade / Dados1("PECAS")) * Dados1("PECAS")
      QuantSug2 = (Int(Quantidade / Dados1("PECAS")) + 1) * Dados1("PECAS")
      response.write "<div class='alert alert-danger' role='alert'>"
      response.write "<strong>Palete Completo!</strong>Este produto exige quantidade que complete os paletes. Quantidades sugeridas: " & QuantSug1 & ", ou " & QuantSug2 & ". Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
      response.write "</div>"
      Response.End
    end if
  end if
  Dados1.Close
end if

if Trim(codcli2) <> "" And Trim(Lojacli2) <> "" then 
  Dados1.Open("Execute SP_VALIDACLIENTEBAIXADO ["&CODCLI2&"],[" &LOJACLI2&"]")
  if Dados1.Eof then
    response.write "<div class='alert alert-danger' role='alert'>"
    response.write "<strong>Cliente inválido!</strong>Este Cliente se encontra baixado. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
    response.write "</div>"
		Response.End
	end if

  Dados1.Close
end if

if tipo_nf = "F" then
  'SÓ GARANTINDO A NOVA VALIDAÇÃO DE CLIENTE REMESSA - SANDRO
  dim Remessa, existe, valor

  SQL = "SELECT TOP 1 A1_NFFAT " 
  SQL = SQL & " FROM SA1010 " 
  SQL = SQL & " WHERE A1_COD = '" & Trim(CODCLI) & "' " 
  SQL = SQL & "AND D_E_L_E_T_ <> '*' " 
  SQL = SQL & "AND A1_MSBLQL <> '1' " 
  SQL = SQL & "AND A1_LOJA = '" & trim(LOJACLI) & "' " 
  SQL = SQL & iif(trim(nivel) = "1"," AND A1_VEND = '" & trim(codsiga) & "' ", "" ) 
  SQL = SQL & "AND 1 = 1 "
  Dados1.Open(SQL)

  if not Dados1.Eof and Trim(Dados1("A1_NFFAT")) <> "" then
    Remessa = Split(Replace(Trim(Dados1("A1_NFFAT")),",",""),"-")
    existe = 0

    for each valor in Remessa 
      if Trim(valor) = Trim(CODCLI2) then
         existe = 1
      end if
    next 

    if existe = 1 then
      SQL = "SELECT A1_COD, " 
      SQL = SQL & "A1_LOJA " 
      SQL = SQL & "FROM SA1010 " 
      SQL = SQL & "WHERE A1_COD = '"&trim(CODCLI2)&"' " 
      SQL = SQL & "AND D_E_L_E_T_ <> '*' " 
      SQL = SQL & "AND A1_MSBLQL <> '1' " 
      SQL = SQL & "AND A1_LOJA = '"&trim(LOJACLI2)&"' "
      Dados2.Open(SQL)

      if Dados2.Eof then
        response.write "<div class='alert alert-danger' role='alert'>"
        response.write "<strong>Cliente remessa!</strong>Cliente remessa não foi encontrado ou esta bloqueado. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
        response.write "</div>"
        Response.end
      end if
      Dados2.Close
    else
      response.write "<div class='alert alert-danger' role='alert'>"
      response.write "<strong>Cliente remessa!</strong>O Cliente não possui remessa cadastrada para este destino, faça a solicitação de cadastro pelo menu 'Solicitar Cliente Remessa'. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
      response.write "</div>"
      Response.end
    end if
  else
    response.write "<div class='alert alert-danger' role='alert'>"
    response.write "<strong>Cliente inválido!</strong>Este Cliente não foi encontrado, esta bloqueado pelo fiscal, ou não pertence à sua carteira. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
    response.write "</div>"
    Response.end
  end if
  Dados1.Close
end if

if (nivel = 2) And (CODCLI <> codsiga) then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Cliente inválido!</strong>Não foi possivel emitir pedido para este Cliente. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
  response.write "</div>"
	Response.End
end if

varB5_TOLMA = 0
if condicao = "CX" and eShop <> "1" then
	Dados1.Open("Execute SP_EXISTE_FT_PARA_O_CLIENTE ["&CodCli&"],["&LojaCli&"],["&produto&"]")
	if Dados1.EOF then
    response.write "<div class='alert alert-danger' role='alert'>"
    response.write "<strong>Produto!</strong> Este produto não foi encontrado, esta bloqueado pelo fiscal, ou não pertence à carteira deste Cliente. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
    response.write "</div>"
		Response.End
  end if

  varB5_TOLMA = Dados1("B5_TOLMA")
  
  Dados1.Close
else
  if eShop = "1" then
    SQL = "" & _
      "SELECT " & _
        "SB1010.B1_COD, " & _
        "SB1010.B1_DESC, " & _
        "SB5010.B5_TOLMA, " & _
        "SB5010.B5_TOLME, " & _
        "SB1010.B1_UPRC, " & _
        "SB1010.B1_UCOM " & _
      "FROM SB1010 with(nolock) " & _
        "LEFT JOIN SB5010 with(nolock) ON SUBSTRING(SB1010.B1_COD,4,6) = SUBSTRING(SB5010.B5_COD,4,6) AND SB5010.D_E_L_E_T_ <> '*' " & _
      "WHERE SB1010.B1_COD = '" & produto & "' " & _
        "AND SB1010.D_E_L_E_T_ <> '*' " & _
        "AND SB1010.B1_STATUS NOT IN ('C', 'P') " & _
        "AND B1_MSBLQL <> '1' "
    Dados1.Open(SQL)

    if Dados1.EOF then
      response.write "<div class='alert alert-danger' role='alert'>"
      response.write "<strong>Produto!</strong> Este produto não foi encontrado, esta bloqueado pelo fiscal ou não esta validado pelo DDP. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
      response.write "</div>"
      Response.End
    end if
    varB5_TOLMA = Dados1("B5_TOLMA")
    Dados1.Close
  end if
end if

if DESCONTO < 0 then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Desconto incorreto!</strong>Não é possível um desconto negativo. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
	Response.End
end if

Dados1.Open("Execute SP_RISCO ["&CodCli&"] ")

if not Dados1.EOF then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Cliente Bloqueado!</strong>Cliente bloqueado pelo departamento financeiro. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
  response.write "</div>"
	Response.End
end if

Dados1.Close
SQL = "SELECT A1_RISCO " & _
      "FROM SA1010 " & _
      "WHERE A1_RISCO = 'J' " & _
        "AND A1_COD = '"&CodCli&"' " & _
        "AND A1_LOJA = '"&LojaCli&"' " & _
        "AND D_E_L_E_T_ <> '*'"
Dados1.Open(SQL)

if not Dados1.EOF then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Loja Bloqueada!</strong>Esta loja do Cliente está bloqueada pelo departamento financeiro. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
  response.write "</div>"
	Response.End
end if
Dados1.Close

if Cstr(preco_unit) = "" then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Preço Unitário!</strong>Favor informar o preço unitário do produto. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
	Response.End
end if

if Cdbl(preco_unit) <= 0 then
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Preço Unitário!</strong>Favor informar o preço unitário do produto. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
  response.write "</div>"
	Response.End
end if

' TIRAR OS VALORES FIXOS PARA A PRODUCAO
SQL = "SELECT A1_ACEITOL, A1_PERCTOL, A1_TOLMAIS "
SQL = SQL & "FROM SA1010 "
SQL = SQL & "WHERE A1_COD = '" & CodCli & "' "
SQL = SQL & "AND A1_LOJA = '" & LojaCli & "' "
SQL = SQL & "AND D_E_L_E_T_ <> '*'"
Dados1.Open(SQL)

varA1_ACEITOL = ""
varA1_PERCTOL = 0
varA1_TOLMAIS = 0

if NOT Dados1.Eof then
  varA1_ACEITOL = Dados1("A1_ACEITOL")
  varA1_PERCTOL = Dados1("A1_PERCTOL")
  varA1_TOLMAIS = Dados1("A1_TOLMAIS")

end if
Dados1.Close

if (condicao="PA") then 
  PercAcrescimo = 1.1
else
  if varA1_ACEITOL = "S" then
     PercAcrescimo = 1 + ( (varA1_TOLMAIS + varA1_PERCTOL) / 100 )
  else
     PercAcrescimo = 1 + (varA1_TOLMAIS / 100 )
  end if
end if

varQTDECOMTOL = ( Cdbl(QUANTIDADE) * PercAcrescimo )
varVLRTOTPED = (varQTDECOMTOL * Cdbl(preco_unit))

SQL = "SELECT * FROM fq_CONSULTA_CREDITO_INC_ALT_PED( '" & CodCli & "', '', '" & produto & "' ,'" & tes & "', " & replace(varVLRTOTPED,",",".") & " )"
Dados1.Open( SQL )

'RESPONSE.WRITE "varA1_TOLMAIS " & varA1_TOLMAIS   & "  varQTDECOMTOL " & varQTDECOMTOL & " varVLRTOTPED " & varVLRTOTPED & "SALDO_CREDITO " & Dados1("SALDO_CREDITO")
if ( Dados1.EOF or Cdbl(Dados1("SALDO_CREDITO")) < 0 ) and eShop <> "1" then
  varLIMITEDISP = Dados1("LIMITE_ATUAL") - Dados1("SALDO_TITULOS_ABERTOS") - Dados1("SALDO_PED_VENDAS") - Dados1("SALDO_PED_N_VALIDADOS")
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Limite de crédito!</strong><br>" &_
                  "O pedido não pode ser gravado, pois o saldo de crédito para este cliente é inferior ao valor do pedido.<br>" &_
                  "LIMITE..................= R$ " & FormatNumber(Dados1("LIMITE_ATUAL"),2) & "<br>" &_
                  "SALDO TÍTULOS ABERTOS...= R$ " & FormatNumber(Dados1("SALDO_TITULOS_ABERTOS"),2) & "<br>" &_
                  "SALDO PEDIDOS ABERTOS...= R$ " & FormatNumber(Dados1("SALDO_PED_VENDAS"),2) & "<br>" &_
                  "SALDO PEDIDOS A VALIDAR.= R$ " & FormatNumber(Dados1("SALDO_PED_N_VALIDADOS"),2) & "<br>" &_
                  "LIMITE DISPONÍVEL.......= R$ " & FormatNumber(varLIMITEDISP,2) & "<br>" &_
                  "VALOR DO PEDIDO + IPI...= R$ " & FormatNumber(Dados1("VALOR_PEDIDO"),2) & "<br><br>" &_
                  "Obs.: VALOR DO PEDIDO CONSIDERANDO QUANTIDADE + TOLERÂNCIAS + IPI <br><br>" &_
                  "Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
  response.end
end if
Dados1.Close
'response.end 

if Trim(codcli) = "0754" And Trim(Lojacli) = "01" then 
  response.write "<div class='alert alert-danger' role='alert'>"
  response.write "<strong>Loja Bloqueada!</strong>A Loja 01 deste Cliente esta bloqueada para emissão de pedidos. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
  response.write "</div>"
	Response.End
end if

'================================ NOVA VALIDAÇÃO DE TES ===========================================
dim TesConsulta,validou,ites
validou = 0

if eShop <> "1" then

  SQL = "" &_
    "SELECT " &_
      "A1_GERTES, "  &_
      "A1_TESEXCE, " &_
      "A1_DTVLTES " &_
    "FROM SA1010 " &_ 
    "WHERE A1_COD = '"&CODCLI&"' " &_ 
      "AND A1_LOJA = '"&LOJACLI&"' " &_ 
      "AND D_E_L_E_T_ = '' " &_ 
      "AND A1_MSBLQL = '2'" 
  Dados1.Open(SQL)

  if not Dados1.Eof then
    if Trim(Dados1("A1_GERTES")) <> "" then
      TesConsulta = split(Dados1("A1_GERTES"),",")

      for ites = 0 to Ubound(TesConsulta)
        if TesConsulta(ites) = trim(tes) then
          validou = 1
        end if
      next

      if validou <> 1 then
        if Dados1("A1_DTVLTES") = InverteDataSiga(Date()) then
          if Trim(Dados1("A1_TESEXCE")) <> tes then
            response.write "<div class='alert alert-danger' role='alert'>"
            response.write "<strong>TES!</strong>Tes incompatível com o cliente, verifique se esta TES foi informada no cadastro do Cliente. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
            response.write "</div>"
            response.end
          end if
        else
            response.write "<div class='alert alert-danger' role='alert'>"
            response.write "<strong>TES!</strong>Tes incompatível com o cliente, verifique se esta TES foi informada no cadastro do Cliente. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
            response.write "</div>"
            response.end
        end if
      end if
    else
      if Dados1("A1_DTVLTES") = InverteDataSiga(Date()) then
        if Trim(Dados1("A1_TESEXCE")) <> tes then
          response.write "<div class='alert alert-danger' role='alert'>"
          response.write "<strong>TES!</strong>Tes incompatível com o cliente, verifique se esta TES foi informada no cadastro do Cliente. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
          response.write "</div>"
          response.end
        end if
      else
          response.write "<div class='alert alert-danger' role='alert'>"
          response.write "<strong>TES!</strong>Tes incompatível com o cliente, verifique se esta TES foi informada no cadastro do Cliente. Clique <a href=# onclick='history.back(1);'>aqui</a> para continuar emitindo pedidos."
          response.write "</div>"
          response.end
      end if
    end if
  end if
  Dados1.Close
else
  validou = 1   
end if        

vHoje2      = now()
vHoraWork   = vhoje2
DoisPontos  = InStr(1,vHoraWork,":")
vHora       = Mid(vHoraWork,DoisPontos-2,2) & Mid(vHoraWork,DoisPontos,6) 
Espaco      = InStr(1,vHoje2, " ")
vHoje1      = Mid(vHoje2,1,Espaco-1)

select case len(vHoje1)
      case 10 vHoje = mid(vHoje1,7,4) & mid(vHoje1,4,2) & mid(vHoje1,1,2)
      case  9 vHoje = mid(vHoje1,6,4) & "0" & mid(vHoje1,4,1) & mid(vHoje1,1,2)
      case  8 vHoje = mid(vHoje1,5,4) & "0" & mid(vHoje1,3,1)& "0" &mid(vHoje1,1,1)
end select

vHoje = Trim(InverteDataSiga(Date()))
vHora = Trim(Time())

SQL = "SELECT A1_CLASCLI, " 
SQL = SQL & "A1_DTEMENO, " 
SQL = SQL & "A1_DTEMAIO, " 
SQL = SQL & "A1_DIASPRO, " 
SQL = SQL & "A1_DTENTLI " 
SQL = SQL & "FROM SA1010 " 
SQL = SQL & "WHERE A1_COD = '" & CODCLI & "' " 
SQL = SQL & "AND A1_LOJA = '" & LOJACLI & "'"
Dados1.Open(SQL)

if not Dados1.Eof then
  ClasCli = Dados1("A1_CLASCLI")
  DiasPro = Dados1("A1_DIASPRO")

  if condicao = "CX" or condicao = "CH" then
    PrimeiraData = DATA1
    DiasMaio = CInt(Dados1("A1_DTEMAIO"))
    SegundaData = DateAdd("d",DiasMaio,DATA1)
    

    '-----------------------------------------------------------------------------------------------
    '										EXPLICANDO VALIDAÇÕES
    'SE O CLIENTE FOR 0001 - NESTLE OU 
    '0005 - SADIA, O SISTEMA NÃO PASSA PELA VALIDAÇÃO DE SABADO E DOMINGO, CASO CONTRÁRIO ELE FARÁ
    'ESTA VALIDAÇÃO.  
    '-----------------------------------------------------------------------------------------------	
    If Not (CODCLI = "0001" Or CODCLI = "0005") Then
      '-------------------------------------------------------------------------------------
      '							VALIDANDO PRIMEIRA DATA SE SABADO
      '						ACRESCENTA MAIS DOIS DIAS À PRIMEIRA DATA
      '-------------------------------------------------------------------------------------
        'if WeekDay(PrimeiraData) = 7 then
        '  PrimeiraData = DateAdd("d",2,PrimeiraData)
        '  MsgData = "Sua PRIMEIRA data de entrega foi alterada para " & PrimeiraData
        'end if
      
      if not inArray( codcli, array( "010050", "013592" ) ) then
        'if codcli <> "010050" then

        '-------------------------------------------------------------------------------------
        ' A REGRA PARA DATA DE ENTREGA NÃO SER AOS DOMINGOS, NÃO SE APLICA PARA O CLIENTE 
        ' 010050 - PANDURATA
        ' 013592 - '
        ' SOLICITAÇÃO FEITA POR MARIA GUIMARÃES 30/04/2025, CHAMADO I2504-1208            
        '-------------------------------------------------------------------------------------

        '-------------------------------------------------------------------------------------
        '							VALIDANDO PRIMEIRA DATA SE DOMINGO
        '						  ACRESCENTA MAIS UM DIA À PRIMEIRA DATA
        '-------------------------------------------------------------------------------------

        if WeekDay(PrimeiraData) = 1 then
          PrimeiraData = DateAdd("d",1,PrimeiraData)
          MsgData = "Sua PRIMEIRA data de entrega foi alterada para " & PrimeiraData
        end if
        
        ' Conforme autorizado pelo Carlos Sobrinho em 05/08/19, por Sandro Cardoso 
        '-------------------------------------------------------------------------------------
        '							VALIDANDO SEGUNDA DATA SE SABADO
        '						ACRESCENTA MAIS DOIS DIAS À SEGUNDA DATA
        '-------------------------------------------------------------------------------------
        if WeekDay(SegundaData) = 7 then
          SegundaData = DateAdd("d",2,SegundaData)
          MsgData = "Sua SEGUNDA data de entrega foi alterada para " & SegundaData
        end if
        '-------------------------------------------------------------------------------------
        '							VALIDANDO SEGUNDA DATA SE DOMINGO
        '						  ACRESCENTA MAIS UM DIA À SEGUNDA DATA
        '-------------------------------------------------------------------------------------
        if WeekDay(SegundaData) = 1 then
          SegundaData = DateAdd("d",1,SegundaData)
          MsgData = "Sua SEGUNDA data de entrega foi alterada para " & SegundaData
        end if

      end if

    end if
  else
    PrimeiraData = DATA1
    SegundaData = DATA2
  end if
end if 

if (condicao <> "PA") then
    DATA1 = PrimeiraData
    DATA2 = SegundaData

    dim Datverifica2
    Datverifica2 = iif( (Trim(data2) = ""), DATA1, DATA2)

    if DateDiff("d",date(),DATA1) > 365 Or DateDiff("d",date(),Datverifica2) > 365 then
      response.write "<div class='alert alert-danger' role='alert'>"
      response.write "<strong>Data Entrega!</strong>Pedido emitido com entrega superior a um ano. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
      response.write "</div>"
      Response.End
    end if

    if data2 <> "" then
      if DateDiff("d",DATA1,DATA2) < 0 then
        response.write "<div class='alert alert-danger' role='alert'>"
        response.write "<strong>Segunda data de entrega!</strong>A segunda data de entrega não pode ser igual ou anterior à primeira. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar e corrigir."
        response.write "</div>"
        Response.End
      end if
    end if

    SQL = "EXECUTE SP_VERIFICAPEDIDOIGUAL [" & CodCli & "], " & _
            "[" & LojaCli & "], " & _
            "[" & Tipo_Frete & "], " & _
            "[" & Cond_Pgto & "], " & _
            "[" & vHoje & "], " & _
            "[" & Desconto & "], " & _
            "[" & Tipo_NF & "], " & _
            iif( (tipo_nf = "F"), "[" & codcli2 & "], " , "[ ], " ) & _
            iif( (tipo_nf = "F"), "[" & Lojacli2 & "], ", "[ ], " ) & _
            "[" & Left(Obs_venda,50) & "], " & _
            "[" & Produto & "], " & _
            "[" & TES & "], " & _
            "[" & quantidade & "], " & _
            "[" & Trim(Replace(Preco_unit,",",".")) & "], "

    vdata1 = mid(DATA1,7,4)& mid(DATA1,4,2) & mid(DATA1,1,2)
    vdata2 = iif( (DATA2 = ""), 0, mid(DATA2,7,4)& mid(DATA2,4,2) & mid(DATA2,1,2) )
    
    SQL = SQL & "[" & vData1 &  "], " & _
                "[" & vData2 &  "], " & _
                iif( (trim(Ped_Cliente) <> ""), "[" & Ped_Cliente & "], ", "[ ], " ) & _
                "[" & vLARG & "], " & _
                "[" & vCOMP & "], " & _
                "[" & NF_Pallet & "], " & _
                "[" & vVINCO_01 & "], " & _
                "[" & vVINCO_02 & "], " & _
                "[" & vVINCO_03 & "], " & _
                "[" & vVinco_04 & "], " & _
                "[" & vVinco_05 & "] "
    Dados2.Open(SQL)
    if not Dados2.Eof then 
      response.write "<div class='alert alert-danger' role='alert'>"
      response.write "<strong>Pedido já Cadastrado!</strong> O pedido não poderá ser emitido, pois um pedido semelhante já se encontra cadastrado em nosso sistema. Clique <a href=# onclick='history.back(1);'>aqui</a> para emitir outro pedido."
      response.write "</div>"
      Response.End
    end if
    Dados2.Close
else
    vdata1 = mid(DATA1,7,4)& mid(DATA1,4,2) & mid(DATA1,1,2)
    vdata2 = iif( (DATA2 = ""), 0, mid(DATA2,7,4)& mid(DATA2,4,2) & mid(DATA2,1,2) )
end if

  vdata5 = " "
  if TRIM(cond_pgto) = "090" And Len(trim(data5)) = 10 then
    vdata5 = mid(data5,7,4)& mid(data5,4,2) & mid(data5,1,2)
  else
    data5 = ""
  end if

  vData1 = InverteDataSiga(DATA1)
  vData2 = InverteDataSiga(DATA2)

  if Trim(Vdata2) = "" Or Trim(Vdata2) = 0 then
    Vdata2 = Vdata1
  end if

  SQL = "SELECT TOP 1 B1_DESCNF " & _
        "FROM SB1010 " & _
        "WHERE D_E_L_E_T_ = '' " & _
          "AND B1_COD = '" & produto & "' " & _
          "AND D_E_L_E_T_ = '' " & _
          "AND B1_DESCNF IN ('901','902','903','900') "
  Dados2.Open(SQL)
  if not Dados2.eof then
     descNf = Dados2("B1_DESCNF") 
  else
    response.write "" &_
      "<div class='alert alert-danger' role='alert'>" &_
        "<strong>Tipo desconhecido!</strong> " &_
        "Não foi possível identificar qual o tipo deste produto. " &_
        "Tipos conhecidos, 900, 901, 902, 903. " &_
        "Por favor entrar em contato com o Desenvolvimento de produtos." &_
      "</div>"
    response.end
  end if
  Dados2.Close

  ' VALIDAÇÃO POR CARGA DE MÁQUINA ----------------------------------------------------------------
  area = 0
  'Area total serve para validar os parametros dos pedidos
  areaTotal = 0
  diasViagem = 0

  if condicao = "CH" then
        'alterado em 14/07/2026
        'area = quantidade * (vLarg/1000)*(vComp/1000)
        area = (vLarg/1000)*(vComp/1000)
        areaTotal = quantidade * (vLarg/1000)*(vComp/1000)
  else
        Dados2.Open("SELECT dbo.funcCalculaArea('" & Produto & "', '') AS AREA")

        if not Dados2.eof then
          'alterado em 14/07/2026
          'area = quantidade * Dados2("AREA")
          area =  Dados2("AREA")
          areaTotal = quantidade * Dados2("AREA")
        end if

        Dados2.Close
  end if

  if Tipo_NF = "N" then
      SQL = "" &_
        "SELECT " & _
          "CASE '" & filial & "' WHEN '04' THEN A1_DIASVIU ELSE A1_DIASVIA END AS DIASVIAGEM " & _
        "FROM SA1010 " & _
        "WHERE A1_COD = '" & codcli & "' " & _
          "AND A1_LOJA = '" & lojacli & "'"
  else
      SQL = "" &_
        "SELECT " & _
          "CASE '" & filial & "' WHEN '04' THEN A1_DIASVIU ELSE A1_DIASVIA END AS DIASVIAGEM " & _
        "FROM SA1010 " & _
        "WHERE A1_COD = '" & codcli2 & "' " & _
          "AND A1_LOJA = '" & lojacli2 & "'"
  end if
  Dados2.Open(SQL)

  if not Dados2.eof then
      diasViagem = Dados2("DIASVIAGEM")
  end if
  Dados2.Close
  
Dados1.Close
if diasViagem = "" then diasViagem = 0
peso = 0

Dados1.Open( "Execute SP_CALCULAPESO '" & produto & "'" )   
if not Dados1.eof then
  peso = iif( (condicao = "CH"), (vLarg * vComp * Dados1("NPESO")) / 1000000, Dados1("NPESO") )
end if

Dados1.Close

if descNf = "900" then
  peso = quantidade
end if	

if peso = 0 then
  response.write "" &_
    "<div class='alert alert-danger' role='alert'>" &_
      "<strong>Peso!</strong> Existe um problema com o produto que não permite o calculo do peso. Favor verificar o cadastro do produto.&nbsp;&nbsp;" &_
      "<a href='#' onclick='window.history.back();' class='btn btn-danger'>Voltar</a>" &_
    "</div>"
  response.end
end if

if descNf > "900" then
  pesoTotal = peso * quantidade
  PrecoKg   = preco_unit / peso
else
  pesoTotal = peso
  PrecoKg   = preco_unit / quantidade
end if


select case condicao
      case "CX"
           if descNf = "901" then
              if Cdbl(pesoTotal) < Cdbl(paramPesoMN) and lotePiloto = "N" then 
                response.write "" &_
                "<div class='alert alert-danger' role='alert'>" &_
                "<strong>Peso Minimo!</strong><br>" &_
                "Peso Total menor que o peso mínimo de " & paramPesoMN &" KG. " &_
                "Clique <a href=# onclick='history.back(1);'>aqui</a> e verifique o peso do produto." &_
                "</div>"
                response.end
              end if
           end if

           if descNf = "902" then
              if Cdbl(pesoTotal) < Cdbl(paramPesoCV) and lotePiloto = "N" then 
                  response.write "" &_
                "<div class='alert alert-danger' role='alert'>" &_
                "<strong>Peso Minimo!</strong><br>" &_
                "Peso Total menor que o peso mínimo de " & paramPesoCV &" KG. " &_
                "Clique <a href=# onclick='history.back(1);'>aqui</a> e verifique o peso do produto." &_
                "</div>"
                response.end
              end if
          end if

      case "CH"
         
          if Cdbl(pesoTotal) < Cdbl(paramPesoCH) and lotePiloto = "N" then 
               response.write "" &_
               "<div class='alert alert-danger' role='alert'>" &_
               "<strong>Peso Minimo!</strong><br>" &_
               "Peso Total menor que o peso mínimo de " & paramPesoCH &" KG. " &_
               "Clique <a href=# onclick='history.back(1);'>aqui</a> e verifique o peso do produto." &_
               "</div>"
               response.end

           end if
      'case "PA"
           'if Cdbl(pesoTotal) < Cdbl(pesoTotalPapel) then 
               'response.write "" &_
               '"<div class='alert alert-danger' role='alert'>" &_
               '"<strong>Peso Minimo!</strong><br>" &_
               '"Peso Total menor que o peso mínimo de " & pesoTotalPapel &" KG. " &_
               '"Clique <a href=# onclick='history.back(1);'>aqui</a> e verifique o peso do produto." &_
               '"</div>"
               'response.end

           'end if


end select  

select case descNf
case "900" PrecoKgParametro = 0
case "901" PrecoKgParametro = precoKgMaleta
case "902" PrecoKgParametro = precoKgCorteVinco
case "903" PrecoKgParametro = precoKgChapa
end select

'Validando Area total 
select case condicao
      case "CX"
           if descNf = "901" then
              if Cdbl(areaTotal) < Cdbl(ParamAreaMN) and lotePiloto = "N" then 
                response.write "" &_
                "<div class='alert alert-danger' role='alert'>" &_
                 "<strong>Area Minima!</strong><br>" &_
                "Area Total menor que o area mínima de " & ParamAreaMN &" M2. " &_
                "Clique <a href=# onclick='history.back(1);'>aqui</a> e verifique a area do produto." &_
                "</div>"
                response.end
              end if
           end if

           if descNf = "902" then
              if Cdbl(areaTotal) < Cdbl(ParamAreaCV) and lotePiloto = "N" then 
                  response.write "" &_
                "<div class='alert alert-danger' role='alert'>" &_
                 "<strong>Area Minima!</strong><br>" &_
                "Area Total menor que o area mínima de " & ParamAreaCV &" M2. " &_
                "Clique <a href=# onclick='history.back(1);'>aqui</a> e verifique o area do produto." &_
                "</div>"
                response.end
              end if
          end if

      case "CH"
          if Cdbl(areaTotal) < Cdbl(ParamAreaCH) and lotePiloto = "N" then 
               response.write "" &_
               "<div class='alert alert-danger' role='alert'>" &_
               "<strong>Area Minima!</strong><br>" &_
               "Area Total menor que o area mínima de " & ParamAreaCH &" M2. " &_
                "Clique <a href=# onclick='history.back(1);'>aqui</a> e verifique o area do produto." &_
               "</div>"
               response.end

           end if
end select  


existeAlertaPrecoPorQuilo = 0
if Cdbl(PrecoKgParametro) > Cdbl(PrecoKg) then
  existeAlertaPrecoPorQuilo = 1
end if

if trim(descarregamento) <> "" then
  while len(trim(descarregamento)) < 10
    descarregamento = "*"&descarregamento
  wend
end if

descNf = 0
Session("CodCli")               = CodCli
Session("LojaCli")              = LojaCli
Session("CodRepre")             = CodRepre
Session("Tipo_Frete")           = Tipo_Frete
Session("Cond_Pgto")            = Cond_Pgto
Session("vHoje")                = vHoje
Session("Desconto")             = Desconto
Session("Tipo_NF")              = Tipo_NF
Session("codcli2")              = codcli2
Session("Lojacli2")             = Lojacli2
Session("Obs_venda")            = RemoveAcentos(Obs_venda)
Session("Produto")              = Produto
Session("TES")                  = TES
Session("quantidade")           = quantidade
Session("Preco_unit")           = replace(Preco_unit,",",".")
Session("vData1")               = vData1
Session("vData2")               = vData2
Session("vData5")               = vData5
Session("EntregaSabadoDomingo") = EntregaSabadoDomingo
Session("Ped_Cliente")          = iif( (trim(Ped_Cliente) <> ""), Trim(RemoveAcentos(Ped_Cliente)), "" )
Session("ItemPedCli")           = iif( (trim(ItemPedCli) <> "") , Trim(RemoveAcentos(ItemPedCli)) , "" )
Session("vLARG")                = vLARG
Session("vCOMP")                = vCOMP
Session("NF_Pallet")            = NF_Pallet
Session("usr")                  = usr
Session("vVINCO_01")            = vVINCO_01
Session("vVINCO_02")            = vVINCO_02
Session("vVINCO_03")            = vVINCO_03
Session("vVINCO_04")            = vVINCO_04
Session("vVINCO_05")            = vVINCO_05
Session("vHora")                = vHora
Session("compap")               = compap
Session("GRAM")                 = GRAM
Session("FORMATO")              = FORMATO
Session("DIAM")                 = DIAM
Session("TUBETE")               = TUBETE
Session("TRANSPA")              = TRANSPA
Session("FRETE_UNIT")           = Frete_Unit  
Session("enviaemailcli")        = enviaemailcli
Session("MsgData")              = MsgData
Session("descarregamento")      = descarregamento
Session("CODFAT")               = CODFAT
Session("CODFAT2")              = CODFAT2
Session("PM")                   = pedmae
Session("PALLETCH")             = PalletCH
Session("horaDescarga")         = horaDescarga
Session("dataDescarga")         = dataDescarga
Session("MsgNota")              = MsgNota
Session("MsgNota2")             = MsgNota2
Session("MsgNota3")             = MsgNota3
Session("MsgNota4")             = MsgNota4
Session("lotePiloto")           = LotePiloto
Session("estrutura")            = Estrutura
Session("area")                 = area
Session("peso")                 = peso
Session("diasViagem")           = diasViagem
Session("descNf")               = descNf
session("DescCom")              = DescCom
session("pesoTotal")            = pesoTotal
session("moedaExportacao")      = moedaExportacao
session("existeAlertaPrecoPorQuilo") = existeAlertaPrecoPorQuilo
session("emailAbaixoPrecoKg")   = emailAbaixoPrecoKg
session("eShop")                = eShop   
session("obs_lotepiloto")       = obsLotePiloto 

%>
<html>
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
  <script>
    location.href = "../gravacao/GravarPedido.asp";
  </script>
</html>
<%

dim Ano, Mes, Dia
function InverteDataSiga(InverteData)
  Ano = Year(InverteData)
  Mes = Month(InverteData)
  Dia = Day(InverteData)

  if Len(Ano) > 4 then
    Ano = "20" & Ano
  end if		

  if Len(Mes) =1 then 
    Mes = "0" & Mes
  end if

  if Len(Dia) = 1 then
    Dia = "0" & Dia
  end if

  InverteDataSiga = Ano & Mes & Dia
end function

'Função que coloca o nos dias e meses da data normal
function AcertaData(InverteData)
  Ano = Year(InverteData)
  Mes = Month(InverteData)
  Dia = Day(InverteData)

  if Len(Ano) > 4 then
    Ano = "20" & Ano
  end if		

  if Len(Mes) =1 then 
    Mes = "0" & Mes
  end if

  if Len(Dia) = 1 then
    Dia = "0" & Dia
  end if

  AcertaData = Dia & "/" & Mes & "/" & Ano
end function

dim NomeDia
function MostraNomeData(MostraData)
  select case WeekDay(MostraData)
        Case 1 NomeDia = "Domingo"
        Case 2 NomeDia = "Segunda-feira"
        Case 3 NomeDia = "Terça-feira"
        Case 4 NomeDia = "Quarta-feira"
        Case 5 NomeDia = "Quinta-feira"
        Case 6 NomeDia = "Sexta-feira"
        Case 7 NomeDia = "Sábado"
  end select

  MostraNomeData = NomeDia
end function 
 
dim Texto 
function DataIncorreta(DiasMinimo,Tipo)
  tipo1 = "<td class=texto_simples2>" & _
            "Primeira data de entrega incorreta!</br>" & _
            "A primeira data de entrega tem que ser superior a <b>" & DiasMinimo - 1 & " dias</b>.</br>" & _
            "Data mínima: " & DateAdd("d",now,4) & ".<br>" & _
            "<a href=# onclick='history.back(1);'>Clique aqui para voltar e corrigir.</a>" & _
          "</td>"

  tipoX = "<td class=texto_simples2>" & _
            "Primeira data de entrega incorreta!</br>" & _
            "A primeira data de entrega tem que ser superior a <b>" & DiasMinimo - 1 & " dias Uteis</b>." & _
            "<a href=# onclick='history.back(1);'>Clique aqui para voltar e corrigir.</a>" & _
          "</td>"

  Texto = "" & _
    "</br>" & _
	  "<table width=450 border=0 cellpadding=0 cellspacing=0 bgcolor=#FF0000 align=center>" & _
	    "<tr>" & _
	      "<td align=center valign=top class=texto_atencao style='padding=10px'>ATEN&Ccedil;&Atilde;O<br>" & _
	        "<table width=100% border=0 cellspacing=0 cellpadding=0>" & _
	          "<tr>" & _
	            "<td align=center bgcolor=#FFFFFF style='padding=10px'>" & _
	              "<table width=100% border=0 cellspacing=0 cellpadding=0>" & _
	                "<tr>" & _
	                  "<td style=padding=2px><img src=../../images/alerta_val_user.gif width=31 height=31 align=middle></td>" & _
                    iif( (Tipo = 1), tipo1, tipoX ) & _
                	"</tr>" & _
	              "</table>" & _                
	            "</td>" & _
	          "</tr>" & _
	        "</table>" & _
        "</td>" & _
	    "</tr>" & _
	  "</table>"
	DataIncorreta = Texto
end function

dim diferenca,numero,weekend,contador,data_inicial,data_final

function DiaUtil(data_inicial,data_final)
	diferenca = DateDiff("d",data_inicial,data_final)
	numero = Weekday(data_inicial)
  if (numero = 1 or numero = 7) then 
    numero = numero + 1
  end if

  for contador = 0 to diferenca
    if numero = 1 then 
      weekend = weekend + 1
    end if

    if numero = 7 then 
      weekend = weekend + 1
    end if

    numero = numero + 1
    if numero > 7 then 
      numero = 1
    end if
  next

  DiaUtil = diferenca - weekend
end function
%>