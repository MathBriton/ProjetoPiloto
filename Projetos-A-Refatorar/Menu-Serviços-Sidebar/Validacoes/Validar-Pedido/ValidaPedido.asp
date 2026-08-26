<!DOCTYPE html>
<!-- v2.0 Data: 14-06-21-->
<!-- Alterações decritas no Chpl 197 -->
<html lang="pt-br" xmlns="http://www.w3.org/1999/xhtml">

<%
'on error resume next
'response.write (p(10))
'Response.write(Err.Number)
'response.write("<br>")
'response.write(Err.description)
'response.write("<br>")
%>
<head>
    <meta charset="utf-8" />
    <title>Validação de Pedido</title>
    <!--#include file="../../includes/variaveis.asp" -->
    <!--#Include File = "../../cnn/cnn.asp"-->
    <!--#include file="../../funcoes/funcoes.asp" -->
    <%
    call ValidaPermissaoFormularios("../comercial/validacao/ListaPedidosAValidar",SESSION("CODUSER"),SESSION("FILIAL"))
    %>
    <!--[if lt IE 9]>
    <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
    <script src="https://oss.maxcdn.com/libs/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->

    <script src="../../../ajax/libs/jquery/1.11.3/jquery.min.js"></script>
    <script src="../../../bootstrap/js/bootstrap.min.js"></script>
    <script type="text/javascript">
        function BloqueioLimiteCredito() {
            form.submit();
        }
        function validaBloqueio() {
            if ((document.form.radiobutton[0].checked == true) ||
                (document.form.radiobutton[1].checked == true) ||
                (document.form.radiobutton[2].checked == true) ||
                (document.form.radiobutton[3].checked == true) ||
                (document.form.radiobutton[4].checked == true)) {
                form.submit();
            }
        }
    </script>
      <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />
    <link href="../../bootstrap/css/bootstrap.min.css" rel="stylesheet" />
    <style type="text/css">
        body {
            margin-top: 20px;
        }

        .margemLegenda {
            margin-bottom: 20px;
        }

        .textoAtencao {
            color: red;
        }

        .linhaAtencao {
            background-color: yellow;
        }

        .margemTopo {
            margin-top: 20px;
        }
    </style>
</head>
<%
      SQL = "SELECT "
SQL = SQL &   "Dias, ISNULL(DiaUtil,0) DiaUtil,DiasBloqueiaP1, ISNULL(BloqueiaMaquinaChapa,0) BloqueiaMaquinaChapa, ISNULL(BloqueiaMaquinaCaixa,0) BloqueiaMaquinaCaixa, ISNULL(BloqueiaMaquinaCorteVinco,0) BloqueiaMaquinaCorteVinco, "
SQL = SQL &   "ISNULL(DeMaleta,'') DeMaleta, ISNULL(AteMaleta,'') AteMaleta, ISNULL(DeCorteVinco,'') DeCorteVinco, ISNULL(AteCorteVinco,'') AteCorteVinco, ISNULL(DeChapa,'') DeChapa, ISNULL(AteChapa,'') AteChapa, ISNULL(DePapel,'') DePapel, "
SQL = SQL &   "ISNULL(AtePapel,'') AtePapel, DiasViagem, CargaMaquina, DiasAcrescimoPA, DiasAcrescimoCX, DiasAcrescimoCH, PesoMN, PesoCV, PesoCH, PrecoKgCorteVinco, PrecoKgMaleta, PrecoKgChapa, EmailAbaixoPrecoKg "
SQL = SQL & "FROM CORPORATIVO..parametrosPedidos "
SQL = SQL & "WHERE id = 1"
 
Dados.Open(SQL)

if not Dados.eof then

    paramDias = Dados("Dias")
    paramDiaUtil = Dados("DiaUtil")
    paramDiasBloqueioP1 = Dados("DiasBloqueiaP1")
    paramDiasViagem = Dados("DiasViagem")
    paramCargaMaquina = Dados("CargaMaquina")
    paramDiasAcrescimoPA = Dados("DiasAcrescimoPA")
    paramDiasAcrescimoCX = Dados("DiasAcrescimoCX")
    paramDiasAcrescimoCH = Dados("DiasAcrescimoCH")
    paramPesoMN = Dados("PesoMN")
    paramPesoCV = Dados("PesoCV")
    paramPesoCH = Dados("PesoCH")
    precoKgCorteVinco = FormatNumber(trim(Dados("PrecoKgCorteVinco")),2)
    precoKgMaleta = FormatNumber(trim(Dados("PrecoKgMaleta")),2)
    precoKgChapa = FormatNumber(trim(Dados("PrecoKgChapa")),2)
    emailAbaixoPrecoKg = trim(Dados("EmailAbaixoPrecoKg"))

else

    'DEFAULT
    paramDias = 4
    paramDiaUtil = "False"
    paramDiasBloqueioP1 = "False"
    paramDiasViagem = "True"
    paramCargaMaquina = "True"
    paramDiasAcrescimoPA = 0
    paramDiasAcrescimoCX = 4
    paramDiasAcrescimoCH = 3
    paramPesoMN = 2500
    paramPesoCV = 1200
    paramPesoCH = 500
    precoKgCorteVinco = 0
    precoKgMaleta = 0
    precoKgChapa = 0
    emailAbaixoPrecoKg = ""

end if

Dados.Close()

dataHora = Now
filial = Request("filial")

If filial = "MATRIZ" Then
    planta = "00"
Else
    planta = "04"
End If 

SELECT CASE planta
CASE "00"
    set consulta = Dados51
    set consulta1 = Dados52
    plantaDescricao = "Pirapetinga"
CASE "04"
    set consulta = Dados61
    set consulta1 = Dados62
    plantaDescricao = "Uberaba"
END SELECT

%>
<body>
    <div class="container-fluid">
<%
        NumPed = Request("pedido")

        SQL = "Execute SP_CONLIBERASZF '" & NumPed &"' "
        Dados.Open(SQL)

        if Dados.Eof then
%>
            <div class="alert alert-warning">
                <strong>Alerta!</strong> Pedido não encontrado.
            </div>
<%
            Dados.Close()
            response.end

        end if

        ExisteExcessaoCargaMaquina = 0
        ultimoPreco = Dados("B1_UPRC")
        quantidade = Dados("ZF_QUANT")
        codigoCliente = Trim(Dados("ZF_CLIENTE"))
        lojaCli = Dados("ZF_LOJACLI")
        areaChapa = 0
        precoAreaChapa = 0
        areaTotalChapa = 0
        TES = Dados("ZF_TES")
        produto = Dados("ZF_PRODUTO")
        precoPedido = Dados("ZF_PRECO")
        ultimoPreco = Dados("B1_UPRC")
        dataEmissao = AcertaData(Dados("ZF_EMISSAO"))
        dataEntrega1 = AcertaData(Dados("ZF_ENT1"))
        dataEntrega2 = AcertaData(Dados("ZF_ENT2"))
        entregaSabadoDomingo = Dados("ZF_SD")
        digitado = Dados("ZF_DIGITAD")
        desconto = Dados("ZF_DESC")
        p1 = Trim(Dados("A1_CLASCLI"))
        descricaoTipoProduto = ""
        
        SELECT CASE Dados("B1_DESCNF")
        CASE 900
            descricaoTipoProduto = "Papel"
            PrecoKgParametro = 0
        CASE 901
            descricaoTipoProduto = "Caixa Maleta"
            PrecoKgParametro = precoKgMaleta
        CASE 902
            descricaoTipoProduto = "Corte e Vinco"
            PrecoKgParametro = precoKgCorteVinco
        CASE 903
            descricaoTipoProduto = "Chapas"
            PrecoKgParametro = precoKgChapa
        END SELECT

        avisoTES = ""

        if TES > 502 then

            avisoTES = "ATENÇÃO PARA ESTA TES !!"

        end if

        avisoTES2 = ""

        if TES = 580 or TES = 581 or TES = 715 or TES = 716 or TES = 626 then

            avisoTES2 = "TES PARA PEDIDO LIVRE DE DEBITO - ATENÇÃO"

        end if

        If Dados("B1_INPUBE") = "U" Then

            fabrica = "UBERABA"

        Else

            fabrica = "MATRIZ"

        End If

        SQL = "Execute SP_CONULTIMASCOMPRAS '" & produto &"' "

        Dados2.Open(SQL)

        SQL = "Execute SP_CALCULAPESO '" & produto & "'"

        Dados3.Open(SQL)

        if Not Dados3.Eof then

            tipoProduto = Dados3("TIPO")
            estrutura = Trim(Dados3("ESTRUTURA"))
            peso = Dados3("NPESO")

            if Mid(estrutura,7,1) = "" THEN

                estrutura = "SIMPLES"

            else

                estrutura = "DUPLA"

            end if

        else

            tipoProduto = ""
            estrutura = ""
            peso = 0

        end if

        Dados3.Close()

        if left(produto,3) = "720" then

            SQL = "SELECT dbo.funcCalculaArea('"&produto&"', '"&NumPed&"') AS AREA"

            Dados3.Open(SQL)

            if Not Dados3.eof then

                areaChapa = Dados3("AREA")
                precoAreaChapa = precoPedido/areaChapa
                areaTotalChapa = areaChapa * Dados("ZF_QUANT")

            end if

            Dados3.Close()

            if peso > 0 then

                PKG = precoPedido/((Dados("ZF_LARG")*Dados("ZF_COMP")*peso)/1000000)
                pesoTotal = Dados("ZF_QUANT")*((Dados("ZF_LARG")*Dados("ZF_COMP")*peso)/1000000)

            else

                PKG = 0
                pesoTotal = 0

            end if

        else

            if peso > 0 then

                PKG = precoPedido/peso
                pesoTotal = Dados("ZF_QUANT") * peso

            else

                PKG = 0
                pesoTotal = 0

            end if

        end if

        alertaPrecoPorQuiloAMenor = ""
        classPrecoPorKgMenor = ""

        if Cdbl(PrecoKgParametro) > Cdbl(PKG) then

            alertaPrecoPorQuiloAMenor = "<strong>Isto não é um Bloqueio!!!</strong> O Preço por quilo deste produto esta abaixo do parametro estipulado que é de "&PrecoKgParametro&". Entre em contato com "&emailAbaixoPrecoKg&"."
            classPrecoPorKgMenor = "text-warning"

        end if

        alertaPeso = ""

        if pesoTotal < paramPesoCV then

            alertaPeso = "<span class='text-danger'>Atenção! Pedido com menos de "&paramPesoCV&" kilos!("&FormatNumber(pesoTotal,0)&" Kg)</span>"

        else

            alertaPeso = FormatNumber(pesoTotal,0)&" Kg"

        end if

        precoKG = CStr(PKG)
        precoKG = Left(precoKG,InStr(1,precoKG,",")+5)

        if (paramDiaUtil = "True" or paramDiaUtil = "Verdadeiro" or paramDiaUtil = "Verdadeiro") And 1 <> 1 Then

            if paramDias > DiaUtil(Date(),dataEntrega1) then

                response.write "<div class='alert alert-danger'>"
                response.write "<strong>Data Bloqueada!.</strong> O pedido não poderá ser emitido para esta data, "&dataEntrega1&", a máquina esta sobrecarregada. Clique <a href=# onclick='history.back(1);'>aqui</a> e tente outra data."
                response.write "</div>"
                response.end

            end if

        else

            if paramDias > datediff("d",Date(),dataEntrega1) And 1 <> 1 then

                response.write "<div class='alert alert-danger'>"
                response.write "<strong>Data Bloqueada!!!</strong> O pedido não poderá ser emitido para esta data, "&dataEntrega1&", a máquina esta sobrecarregada. Clique <a href=# onclick='history.back(1);'>aqui</a> e tente outra data."
                response.write "</div>"
                response.end

            end if

        end if

        SQL = "SELECT Id FROM CORPORATIVO..AbrirMaquina Where Cliente = '"&Trim(codigoCliente)&"' AND Loja = '"&Trim(lojacli)&"' AND Data = '"&InverteData(Trim(dataEntrega1))&"' AND ISNULL(Excluido,0) = 0 AND Tipo = 0"
        Dados3.Open(SQL)

        if Not Dados3.Eof then

            ExisteExcessaoCargaMaquina = 1

        end if

        Dados3.Close

        If (paramCargaMaquina = "True" or paramCargaMaquina = "Verdadeiro") and ExisteExcessaoCargaMaquina = "0" Then

            SQL = "SELECT  dbo.VerificaCargaMaquina('"&NumPed&"','"&produto&"', '"&codigoCliente&"', '"&lojaCli&"', '"&InverteData(Trim(dataEntrega1))&"', '"&quantidade&"', "&Dados("ZF_LARG")&", "&Dados("ZF_COMP")&", 0,'',0, '"&Trim(Dados("ZF_DESTINO"))&"', '"&Trim(Dados("ZF_LOJADES"))&"') AS CargaMaquina" 
            Dados3.Open(SQL)

            if Dados3("CargaMaquina") = "0" then

                response.write "<div class='alert alert-danger'>"
                response.write "<strong>Data Bloqueada!!</strong> O pedido não poderá entrar na carteira para esta data, "&dataEntrega1&", a máquina esta sobrecarregada. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar."
                response.write "</div>"
                response.end

            end if

            Dados3.Close()

        End If

        'VERIFICANDO O LIMITE DE CRÉDITO DO CLIENTE
        SQL = "Execute SP_CONSULTA_CREDITO_CLIENTE '"&codigoCliente&"', '0' "
        Dados3.open(SQL)

        if Session("NIVEL") > 0 then
%>
        <div class="alert alert-warning">
            <strong>Alerta!</strong> Você não tem permissão para acessar esta opção.
        </div>
<%
            response.end 
        end if 
%>
        <form name="form" method="post" action="../Gravacao/GravaStatus.asp">
            <div class="row">
                <div class="col-xs-10 col-sm-10 col-md-10 col-xs-offset-1 col-sm-offset-1 col-md-offset-1 text-center bg-primary">
                    <h4>Detalhes do pedido <strong><%=NumPed%></strong> emitido em <%=dataEmissao %>, por <%=digitado %></h4>
                </div>
            </div>
            <div class="row margemTopo">
                <div class="col-xs-10 col-sm-10 col-md-10 col-xs-offset-1 col-sm-offset-1 col-md-offset-1">
                    <table class="table table-condensed table-hover">
                        <thead>
                            <%if Trim(Dados("ZF_LOTEPIL")) = "S" then %>
                            <tr>
                                <td colspan="5" class="alert-danger text-center">
                                    <strong>ATENÇÃO!</strong> ESTE PEDIDO É UM LOTE PILOTO.
                                </td>
                            </tr>
                            <%End if %>
                            <%if Trim(alertaPrecoPorQuiloAMenor) <> "" then %>
                            <tr>
                                <td colspan="5" class="alert-warning text-center">
                                    <%=alertaPrecoPorQuiloAMenor %>
                                </td>
                            </tr>
                            <%End if %>
                            <tr>
                                <td colspan="5">
                                    Cliente: <%=codigoCliente %>/<%=lojaCli %> - <%=Trim(Dados("A1_NOME"))%> | <strong>Saldo de crédito: <%=FormatNumber(Dados3("SALDO_CREDITO"),2) %></strong>
                                </td>
                            </tr>
                            <%if Trim(Dados("ZF_REMESSA")) <> "N" then %>
                            <tr>
                                <td class="alert-warning"><strong>Local de Entrega:</strong></td>
                                <td class="alert-warning" colspan="4"><strong><%=Trim(Dados("ZF_DESTINO")) & " - " & Trim(Dados("ZF_LOJADES")) & " - " & Trim(Dados("CLIENTEREMESSA")) & " - " & Trim(Dados("MUNICIPIOREMESSA"))%></strong></td>
                            </tr>
                            <%end if %>
                            <tr>
                                <td>Produto:</td>
                                <td><%=produto %> - <%=descricaoTipoProduto %> </td>
                                <td>Referência:</td>
                                <td colspan="2"><%=Trim(Dados("B1_DESC"))%></td>
                            </tr>
                            <tr>
                                <td>Fabricação:</td>
                                <td><%=fabrica %></td>
                                <td>Ultima Compra:</td>
                                <td colspan="2"><%=AcertaData(Dados("B1_UCOM")) %></td>
                            </tr>
                            <tr>
                                <td>Quantidade:</td>
                                <td><%=quantidade%></td>
                                <td>Peso:</td>
                                <td colspan="2"><%=alertaPeso%></td>
                            </tr>
                            <tr>
                                <td>Condição de Pagamento:</td>
                                <td>
                                    <%=Dados("E4_DESCRI")%>
                                    <% If trim(Dados("E4_TIPO")) = "9" then %>
                                    - <strong>Vencimento: <%=AcertaData(Dados("ZF_DATA5"))%></strong><br />
                                    <strong class="text-warning">Atenção: Data de Vencimento especificada pelo Vendedor</strong>
                                    <% End If %>
                                </td>
                                <td>TES</td>
                                <td>
                                    <%=TES%><strong class="alert-warning"> <%=AvisoTES%><br /><%=AvisoTES2 %></strong>
                                </td>
                            </tr>
                            <tr>
                                <td>Preço:</td>
                                <td><%="R$ " & FormatNumber(precoPedido,5)%></td>
                                <td>Desconto:</td>
                                <td colspan="2"><%=Dados("ZF_DESC")&"%"%></td>
                            </tr>

                            <tr>
                                <td>Ultimo Preço:</td>
                                <td><%="R$ " & FormatNumber(ultimoPreco,5)%></td>
                                <td class="<%=classPrecoPorKgMenor %>">Preço/Kg:</td>
                                <td colspan="2" class="<%=classPrecoPorKgMenor %>"><%="R$ " & FormatNumber(precoKG,5)%></td>
                            </tr>
                            <tr>
                                <td>Entrega 1:</td>
                                <td>
                                    <%=dataEntrega1%>&nbsp;
                                    <%if DiffDias < paramDias then%>
                                    <strong class="text-danger"><%=AlertaDias%></strong>
                                    <%else%>
                                    <strong class="text-success"><%=AlertaDias%></strong>
                                    <%end if%>
                                </td>
                                <td>Entrega 2:</td>
                                <td colspan="2"><%=dataEntrega2%></td>
                            </tr>
                            <%if left(produto,3) = "720" then %>
                            <tr>
                                <td>Area(m²):</td>
                                <td>
                                    <%=FormatNumber(areaTotalChapa,0) %>
                                </td>
                                <td>Preço/m²:</td>
                                <td colspan="2"><%="R$ " & FormatNumber(precoAreaChapa,5)%></td>
                            </tr>
                            <%end if %>
                            <tr>
                                <td>Obs. Venda:</td>
                                <td colspan="4"><%=Dados("ZF_OBSVEND")%></td>
                            </tr>
                            <%
                            Response.Flush()
                            Response.End

                            SQL = "SELECT Orcamento,Erro,Usuario,DataHora,Planta, Referencia,Cliente,Preco,IPI,ICMS,Margem,Fechamento,TipoCalculo,Pedido FROM OrcamentosTrimbox Where Pedido = '"&NumPed&"' "
                            Dados11.Open(SQL)

                            if Dados11.eof then

                                A1_MUN = trim(Dados("A1_MUN"))
                                A1_EST = trim(Dados("A1_EST"))

                                if Trim(Dados("ZF_REMESSA")) <> "N" then 

                                    A1_MUN = Trim(Dados("MUNICIPIOREMESSA"))
                                    A1_EST = Trim(Dados("UFREMESSA"))

                                end if

                                SQL1 = "BEGIN TRANSACTION "
                                SQL1 = SQL1 & "DECLARE @Carcteres varchar(100), @return_value int, @proximoproduto varchar(20), @erro varchar(8000) EXEC @return_value = [dbo].[pg_ProxCodPreFTPelaInpanet] @proximoproduto = @proximoproduto OUTPUT, "
                                SQL1 = SQL1 & "@erro = @erro OUTPUT DECLARE @Item Varchar(20) DECLARE @Conjunto Varchar(20) SET @Conjunto = @proximoproduto SET @Item = @Conjunto "
                                SQL1 = SQL1 & "SET @Carcteres = 'abcdefghijklmnopqrstuvwxyz' "
                                SQL1 = SQL1 & "Insert TSK_IN_PreFTs(Cliente,NomeCliente,IDMunicipio,UF  ,Referencia , "
                                SQL1 = SQL1 & "CodRepresentante,CodFechamento, Maquina1, Maquina2,Maquina3,Maquina4, "
                                SQL1 = SQL1 & "Cor1   ,Cor2   ,Cor3,Cor4,ConsumoCor1,ConsumoCor2,ConsumoCor3,ConsumoCor4,"
                                SQL1 = SQL1 & "IDPalete,Composicao   ,TipoABNT,Largura,Comprimento,LarguraInterna,ComprimentoInterno,AlturaInterna,"
                                SQL1 = SQL1 & "AreaBrutaPeca,Usuario,IsentoICMS,IsentoIPI,Item,Conjunto, "
                                SQL1 = SQL1 & "Estado,MultLarg,MultComp,PecasPorConjunto,Paletizado, "
                                SQL1 = SQL1 & "ResinaInterna,ResinaExterna,EmpilhamentoMaximo,PecasPorPacote,PecasPorUnidade,"
                                SQL1 = SQL1 & "PacotesLargura, PacotesComprimento, PacotesAltura, Filme) "
                                
                                if left(produto,3) = "720" then

                                    SQL1 = SQL1 & "" & GeraConsultaMargem720(Dados("ZF_LARG"),Dados("ZF_COMP"),A1_EST,A1_MUN)

                                else

                                    SQL1 = SQL1 & "Select '"&Dados("ZF_CLIENTE")&"','"&Trim(Dados("A1_NREDUZ"))&" ','"&Municipo(A1_EST,A1_MUN)&"','"&A1_EST&"'  ,Referencia , "
                                    SQL1 = SQL1 & "CodRepresentante,Fechamento, Maquina1, Maquina2,Maquina3,Maquina4, "
                                    SQL1 = SQL1 & "Cor1   ,Cor2   ,Cor3,Cor4,"
                                    SQL1 = SQL1 & "CAST(ISNULL(ConsumoCor1,0) AS FLOAT)/10000 * 100/AreaBrutaPeca, "
                                    SQL1 = SQL1 & "CAST(ISNULL(ConsumoCor2,0) AS FLOAT)/10000 * 100/AreaBrutaPeca, "
                                    SQL1 = SQL1 & "CAST(ISNULL(ConsumoCor3,0) AS FLOAT)/10000 * 100/AreaBrutaPeca, "
                                    SQL1 = SQL1 & "CAST(ISNULL(ConsumoCor4,0) AS FLOAT)/10000 * 100/AreaBrutaPeca, "
                                    SQL1 = SQL1 & "IdPalete,Composicao,TipoABNT, "

                                    if left(produto,3) <> "720" then

                                        SQL1 = SQL1 & "Largura,Comprimento, "

                                    else

                                        SQL1 = SQL1 &""& Dados("ZF_LARG") & "," & Dados("ZF_COMP") & ", "

                                    end if

                                    SQL1 = SQL1 & "LarguraInterna,ComprimentoInterno,AlturaInterna, "
                                    SQL1 = SQL1 & "AreaBrutaPeca,0,0,0,@Conjunto + SUBSTRING(@Carcteres, ROW_NUMBER() over(order by Conjunto),1),@Conjunto, "
                                    SQL1 = SQL1 & "1,MultLarg,MultComp,PecasPorConjunto,Paletizado, "
                                    SQL1 = SQL1 & "ResinaInterna,ResinaExterna,EmpilhamentoMaximo,PecasPorPacote,PecasPorUnidade, "
                                    SQL1 = SQL1 & "PacotesLargura, PacotesComprimento, PacotesAltura, Filme "
                                    SQL1 = SQL1 & "from vw_ConsultaFTs where Conjunto = '"&Trim(Dados("ZF_PRODUTO"))&"' "

                                end if

                               ''''SQL1 = SQL1 & "VALUES('" & Dados1("A1_COD") & "','" & Dados1("A1_NREDUZ") & "','" & CodMunicipio & "','" & Dados1("A1_EST") & "','" & Dados1("G1_COMP") & "', "                                              'Cliente, NomeCliente, IDMunicipio, UF  ,Referencia
                               '''''SQL1 = SQL1 & "'" & Dados1("A1_VEND") & "','" & Dados1("B5_FECHAM") & "','" & Dados1("B5_MAQ1") & "','" & Dados1("B5_MAQ2") & "','" & Dados1("B5_MAQ3") & "','" & Dados1("B5_MAQ4") & "', "                  'CodRepresentante, CodFechamento, Maquina1, Maquina2,Maquina3,Maquina4,
                               ''''SQL1 = SQL1 & "'" & Dados1("A1_VEND") & "','" & Dados1("B5_FECHAM") & "','041','039','029','" & Dados1("B5_MAQ4") & "', "
                               ''''SQL1 = SQL1 & "'" & Dados1("B5_COR1") & "','" & Dados1("B5_COR2") & "','" & Dados1("B5_COR3") & "','" & Dados1("B5_COR4") & "',"                                                                             'Cor1, Cor2, Cor3, Cor4, 
                               ''''SQL1 = SQL1 & "'" & FormatNumber(Dados1("B5_ARECOR1"),0) & "','" & FormatNumber(Dados1("B5_ARECOR2"),0) & "','" & FormatNumber(Dados1("B5_ARECOR3"),0) & "','" & FormatNumber(Dados1("B5_ARECOR4"),0) & "'," 'ConsumoCor1, ConsumoCor2,ConsumoCor3,ConsumoCor4
                               ''''SQL1 = SQL1 & "'" & Trim(Dados1("B5_TPPALE")) & "','" & Dados1("ZF_ESTRUTU") & "','" & Estilo & "'," & Dados1("B5_LARGCH") & "," & Dados1("B5_COMPCH") & ","                                                       'IDPalete, Composicao, TipoABNT,Largura,Comprimento,
                               '''''SQL1 = SQL1 & "'254','" & Dados1("ZF_ESTRUTU") & "','" & Estilo & "'," & Dados1("B5_LARGCH") & "," & Dados1("B5_COMPCH") & ","
                               ''''SQL1 = SQL1 & "" & Dados1("B5_LARG") & "," & Dados1("B5_COMPR") & "," & Dados1("B5_ALTU") & "," & Replace(Dados1("B5_AREA"),",",".") & ","                                                                   'LarguraInterna,ComprimentoInterno,AlturaInterna,AreaBrutaPeca
                               ''''SQL1 = SQL1 & "'',0,0,@Item+'" & letra & "',@Conjunto,"                                                                                                                                                      ',Usuario,IsentoICMS,IsentoIPI,Item,Conjunto
                               ''''SQL1 = SQL1 & "1," & Dados1("B5_ARRL") & "," & Dados1("B5_ARRC") & "," & Dados1("TOTAL_PECAS") & "," & Palete & ","                                                                                           'Estado,MultLarg,MultComp,PecasPorConjunto,Paletizado,
                               ''''SQL1 = SQL1 & "0,0,'2'," & Dados1("B5_QTDPCT") & "," & Dados1("B5_PCSPAL") & ", " 'ResinaInterna,ResinaExterna,EmpilhamentoMaximo,PecasPorPacote,PecasPorUnidade
                               ''''SQL1 = SQL1 & "" & Dados1("B5_PLLAST") & ",1," & Dados1("B5_ARRPA") & "," & Filme & ") "     'PacotesLargura, PacotesComprimento (sempre 1), PacotesAltura, Filme (sempre 1)                                                                                                                     
                               ''''Quanti = Dados1("ZF_QUANT")
                               ''''CondPg = Dados1("ZF_CONDPAG")
                               ''''contador = contador + 1
                               ''''Dados1.MoveNext()
                               ''''Wend

                               ''''Dados1.MoveFirst()
                                SQL1 = SQL1 & "DECLARE @orcamento int "

                                tipoCalculo = 2
                                tipoProdutoMargem = 1
                                tipoCaminhao = "CARRETA 14,00 | 5 EI"

                                if left(produto,3) = "720" Then

                                    tipoProdutoMargem = 3

                                end if

                                SQL1 = SQL1 & "EXEC @return_value = [dbo].[ssp_ConsultaOrcamentoProdutosPeloInpanet_SK] @conjunto = @Conjunto, "
                                SQL1 = SQL1 & "@tipoproduto = "&tipoProdutoMargem&", "
                                SQL1 = SQL1 & "@quantidade = " & Dados("ZF_QUANT") & ", "
                                SQL1 = SQL1 & "@condpag = N'" & Dados("ZF_CONDPAG") & "', "
                                SQL1 = SQL1 & "@tipocaminhao = N'"&tipoCaminhao&"', "
                                SQL1 = SQL1 & "@tipocalculo = "&tipoCalculo&", "
                                SQL1 = SQL1 & "@argumento = " & Replace(Replace(Dados("ZF_PRECO") * 1000,".",""),",",".") & ", "
                                SQL1 = SQL1 & "@orcamento = @orcamento OUTPUT, "
                                SQL1 = SQL1 & "@erro = @erro OUTPUT "
                                SQL1 = SQL1 & "If object_id('OrcamentosTrimInpanetSK') is null "
                                SQL1 = SQL1 & "SELECT @orcamento as orcamento, @erro as erro, '"&USR&"' as Usuario, '"&dataHora&"' as DataHora, '"&NumPed&"' as Pedido into OrcamentosTrimInpanetSK Else "
                                SQL1 = SQL1 & "Begin insert into OrcamentosTrimInpanetSK(Orcamento,erro,Usuario,DataHora,Pedido) "
                                SQL1 = SQL1 & "SELECT @orcamento as Orcamento, @erro as Erro, '"&USR&"' as Usuario, '"&dataHora&"' as DataHora, '"&NumPed&"' as Pedido End "
                            
                                SQL1 = SQL1 & "IF @@ERROR <> 0 ROLLBACK TRANSACTION ELSE "
                                SQL1 = SQL1 & "COMMIT TRANSACTION "
                                preco = Dados("ZF_PRECO") * 1000
                                response.write SQL1

                                consulta.Open(SQL1)

                                SQL = "SELECT ISNULL(Orcamento,0) Orcamento, Erro, Usuario, DataHora from OrcamentosTrimInpanetSK WITH(NOLOCK) WHERE Usuario = '"&USR&"' AND DataHora = '"&dataHora&"' and Pedido = '"&NumPed&"' "
                                consulta.open(SQL)

                                if not consulta.eof then

                                    orcamento = consulta("Orcamento")
                                    msgErroOrcamento = consulta("Erro")

                                else

                                    orcamento = 0
                                    msgErroOrcamento = "Houve um erro não foi possivel realizar a consulta"

                                end if

                                msgErroOrcamento = Replace(msgErroOrcamento,"'","")
                                consulta.Close()

                                'response.write SQL1

                                margem = 0

                                if Trim(msgErroOrcamento) = "" then

                                    if left(produto,3) = "750" then

                                        SQL1 = "SELECT Cabecalho.orcamento,  Cabecalho.DataCriacao, Itens.Conjunto, SUM(CASE PARCELA WHEN 700 THEN VALOR ELSE 0 END) / SUM(CASE PARCELA WHEN 100 THEN VALOR ELSE 0 END) * 100 PORCVALORVENDA "
                                        SQL1 = SQL1 & "FROM vw_ConsultaOrcamentosPelaInpaNet AS Cabecalho "
                                        SQL1 = SQL1 & "INNER JOIN vw_ConsultaItensOrcamentosPelaInpaNet AS Itens ON Cabecalho.Orcamento = Itens.Orcamento "
                                        SQL1 = SQL1 & "INNER JOIN vw_ConsultaCustosItensOrcamentosPelaInpaNet As Orcamentos ON Itens.Orcamento = Orcamentos.Orcamento AND Itens.Produto = Orcamentos.Produto "
                                        SQL1 = SQL1 & "INNER JOIN TSK_IN_PreFTs As Interface WITH(NOLOCK) ON Interface.Item = Itens.Produto "
                                        SQL1 = SQL1 & "WHERE Cabecalho.Orcamento = "&orcamento&" AND PARCELA IN (100,700) "
                                        SQL1 = SQL1 & "GROUP BY Cabecalho.orcamento, Cabecalho.DataCriacao, Itens.Conjunto "

                                        consulta.Open(SQL1)

                                        If Not consulta.EOF Then

                                            margem = FormatNumber(consulta("PORCVALORVENDA"),2)
                                            conjunto = consulta("Conjunto")

                                        else

                                            margem = 0
                                            conjunto = ""

                                        End If

                                        consulta.Close()

                                        SQL1 = "Select "
	                                    SQL1 = SQL1 & "    sum(precoorcado * pecasporconjunto) as PrecoTotal, "
	                                    SQL1 = SQL1 & "    sum(precoorcado * ResultOperacional / 100  * pecasporconjunto) as ResultadoTotal, "
	                                    SQL1 = SQL1 & "    sum(precoorcado * ResultOperacional / 100  * pecasporconjunto) / sum(precoorcado * pecasporconjunto) as ResulPorcentagemMedio "
                                        SQL1 = SQL1 & "from itemorc where orcamento = "&orcamento&" "

                                        if conjunto <> "" then

                                            SQL1 = SQL1 & "and conjunto = '"&conjunto&"' "

                                        END IF
 
                                        consulta.Open(SQL1)

                                        If Not consulta.EOF Then

                                            resultadoOperacional = formatNumber(consulta("ResulPorcentagemMedio"),2)

                                        Else

                                            resultadoOperacional = 0

                                        End If

                                        consulta.Close()

                                    Else

                                               SQL1 = "SELECT "
                                        SQL1 = SQL1 & "sum(CASE Parcela when 700 then PorcValorVenda else 0 end) Margem, "
                                        SQL1 = SQL1 & "sum(CASE Parcela when 900 then PorcValorVenda else 0 end) Resultado "
                                        SQL1 = SQL1 & "FROM vw_ConsultaCustosItensOrcamentosPelaInpaNet WHERE Orcamento = "&orcamento&" AND Parcela IN (700,900) "

                                        consulta.Open(SQL1)

                                        If Not consulta.EOF Then

                                            margem = formatNumber(consulta("Margem"),2)
                                            resultadoOperacional = formatNumber(consulta("Resultado"),2)

                                        Else

                                            margem = 0
                                            resultadoOperacional = 0

                                        End If

                                        consulta.Close()

                                    End if
                                    
                                    valorVendaCalculaSpread = 0
                                    papelCalculaSpread = 0
                                    colaOndCalculaSpread = 0
                                    resinaOndCalculaSpread = 0
                                    tintaCalculaSpread = 0
                                    filmeCalculaSpread = 0
                                    materialEmbalagemCalculaSpread = 0
                                    freteCalculaSpread = 0
                                    contadorValorVendaCalculaSpread = 0

                                    SQL = "Select Parcela,Isnull(Sum(Valor),0) Valor from vw_ConsultaCustosItensOrcamentosPelaInpaNet where Orcamento = "&orcamento&"  "
                                    SQL = SQL &"and Parcela in (100,610,621,622,623,625,628,402) Group By Parcela"

                                    consulta.open(SQL)

                                    while not consulta.eof

                                        if consulta("Parcela") = "100" then

                                            if tipoCalculo = "2" then

                                                valorVendaCalculaSpread =  FormatNumber(consulta("Valor"),2)

                                            else

                                                valorVendaCalculaSpread = valorVendaCalculaSpread + FormatNumber(consulta("Valor"),2)

                                            end if

                                        end if

                                        if consulta("Parcela") = "610" then
                                            papelCalculaSpread = papelCalculaSpread + FormatNumber(consulta("Valor"),2)
                                        end if

                                        if consulta("Parcela") = "621" then
                                            colaOndCalculaSpread = colaOndCalculaSpread + FormatNumber(consulta("Valor"),2)
                                        end if

                                        if consulta("Parcela") = "622" then
                                            resinaOndCalculaSpread = resinaOndCalculaSpread + FormatNumber(consulta("Valor"),2)
                                        end if

                                        if consulta("Parcela") = "623" then
                                            tintaCalculaSpread = tintaCalculaSpread + FormatNumber(consulta("Valor"),2)
                                        end if

                                        if consulta("Parcela") = "625" then
                                            filmeCalculaSpread = filmeCalculaSpread + FormatNumber(consulta("Valor"),2)
                                        end if

                                        if consulta("Parcela") = "628" then
                                            materialEmbalagemCalculaSpread = materialEmbalagemCalculaSpread + FormatNumber(consulta("Valor"),2)
                                        end if

                                        if consulta("Parcela") = "402" then
                                            freteCalculaSpread = freteCalculaSpread + FormatNumber(consulta("Valor"),2)
                                        end if

                                        consulta.MoveNext()

                                    wend

                                    consulta.Close()

                                    spread = valorVendaCalculaSpread - (papelCalculaSpread + colaOndCalculaSpread + resinaOndCalculaSpread + tintaCalculaSpread + filmeCalculaSpread + materialEmbalagemCalculaSpread + freteCalculaSpread)

                                    SQL = "INSERT INTO CORPORATIVO..OrcamentosTrimbox(Orcamento,Erro,Usuario,DataHora,Planta, Referencia,Cliente,Preco,IPI,ICMS,Margem,ResultadoOperacional,Spread,Fechamento,TipoCalculo,Pedido) "
                                    SQL = SQL &"VALUES("&orcamento&",'"&msgErroOrcamento&"','"&USR&"','"&dataHora&"','"&trim(planta)&"','"&Trim(Dados("B1_DESC"))&"','"&Trim(Dados("ZF_CLIENTE"))&"',"&Replace(Replace(Dados("ZF_PRECO") * 1000,".",""),",",".")&",0,0,"&Replace(Replace(Trim(margem),".",""),",",".")&","&Replace(Replace(Trim(resultadoOperacional),".",""),",",".")&","&Replace(Replace(Trim(spread),".",""),",",".")&",'',"&tipoCalculo&",'"&NumPed&"')"

                                    Dados01.Open(SQL)

                                    cor = ClassificaMargem(margem)
%>
                                    <tr class="alert-<%=cor%>">
                                        <td>Margem:</td>
                                        <td colspan="4"><span title="Margem de Contribuição Calculada" style="cursor: pointer"><b><%=Margem%></b></span>
                                            <a class="btn btn-xs btn-<%=cor %>" href="http://br-pir1-m012/corporativo/comercial/Orcamento/Orcamento.asp?Id=<%=base64_encode(orcamento)&""&base64_encode(planta)&""&base64_encode(dataHora)%>&preco=<%=Replace(Replace(Dados("ZF_PRECO") * 1000,".",""),",",".")%>&ipi=0&icms=0&margem=<%=Replace(Replace(Trim(margem),".",""),",",".")%>&fechamento=&tipocalculo=<%=TipoCalculo%>" target="_blank" title="Visualizar orçamento">
                                                <i class="far fa-eye"></i>
                                            </a>
                                        </td>
                                    </tr>
<%
                                else
%>
                                    <tr class="alert-danger">
                                        <td colspan="5"><span title="Margem de Contribuição Calculada" style="cursor: pointer">Não foi possível calcular a Margem.</span>
                                            <p><em><%=msgErroOrcamento %></em></p>
                                        </td>
                                    </tr>
<%
                                end if
                            else
                                cor = ClassificaMargem(Dados11("Margem"))
%>
                                    <tr class="alert-<%=cor%>">
                                        <td>Margem:</td>
                                        <td colspan="4"><span title="Margem de Contribuição Calculada" style="cursor: pointer"><b><%=FormatNumber(Dados11("Margem"),2)%></b></span>
                                            <a class="btn btn-xs btn-<%=cor %>" href="http://br-pir1-m012/corporativo/comercial/Orcamento/Orcamento.asp?Id=<%=base64_encode(Dados11("Orcamento"))&""&base64_encode(Dados11("Planta"))&""&base64_encode(Dados11("DataHora"))%>&preco=<%=Dados11("Preco")%>&ipi=0&icms=0&margem=<%=Dados11("Margem")%>&fechamento=&tipocalculo=<%=Dados11("TipoCalculo")%>" target="_blank" title="Visualizar orçamento">
                                                <i class="far fa-eye"></i>
                                            </a>
                                        </td>
                                    </tr>
<%
                            end if

                            Dados11.Close() 
                            if Trim(Dados("ZF_MENNOTA")) <> "" then
%>
                            <tr>
                                <td colspan="5" class="text-center alert-warning"><strong>Atenção - 1&ordf; Observação para Nota Fiscal</strong></td>
                            </tr>

                            <tr>
                                <td colspan="5"><%=Trim(Dados("ZF_MENNOTA")) %></td>
                            </tr>
<%
                            end if 

                            if Trim(Dados("ZF_MENNOT2")) <> "" then
%>
                            <tr>
                                <td colspan="5" class="text-center alert-warning"><strong>Atenção - 2&ordf; Observação para Nota Fiscal</strong></td>
                            </tr>
                            <tr>
                                <td colspan="5"><%=Trim(Dados("ZF_MENNOT2")) %></td>
                            </tr>
<%
                            end if 

                            if Trim(Dados("ZF_MENNOT3")) <> "" then
%>
                            <tr>
                                <td colspan="5" class="text-center alert-warning"><strong>Atenção - 3&ordf; Observação para Nota Fiscal</strong></td>
                            </tr>
                            <tr>
                                <td colspan="5"><%=Trim(Dados("ZF_MENNOT3")) %></td>
                            </tr>
<%
                            end if 
%>

<%
                            if Trim(Dados("ZF_MENNOT4")) <> "" then
%>
                            <tr>
                                <td colspan="5" class="text-center alert-warning"><strong>Atenção - 4&ordf; Observação para Nota Fiscal</strong></td>
                            </tr>
                            <tr>
                                <td colspan="5"><%=Trim(Dados("ZF_MENNOT4")) %></td>
                            </tr>
<%
                            end if 
%>
                        </thead>
                    </table>
                </div>
<%
                if not Dados2.eof then 
%>
                <div class="col-xs-10 col-sm-10 col-md-10 col-xs-offset-1 col-sm-offset-1 col-md-offset-1">
                    <table class="table table-condensed table-striped">
                        <thead>
                            <tr>
                                <th colspan="12" class="text-center bg-primary"><strong>ULTIMOS PEDIDOS</strong></th>
                            </tr>
                            <tr>
                                <th>Cliente</th>
                                <th>Pedido</th>
                                <th>Referencia</th>
                                <th>Emissao</th>
                                <th>Entrega</th>
                                <th>TES</th>
                                <th>Quantidade</th>
                                <th>Preço</th>
                                <th>Preço/Kg</th>
                                <%if left(produto,3) = "720" then %>
                                <th>Preço/m²</th>
                                <%end if %>
                                <th>Cond. Pagamento</th>
                                <th>%(1)/(2)</th>
                            </tr>
<%
                            while not Dados2.eof

                                if left(produto,3) = "720" then
                                    if peso <> 0 then
                                        larguraPeloComprimento = Dados2("C6_LARG")*Dados2("C6_COMP")
                                        if larguraPeloComprimento = 0 then
                                            PKGPedido = 0
                                        else
                                            PKGPedido = Dados2("C6_PRCVEN")/((larguraPeloComprimento*peso)/1000000)
                                        end if
                                        if  Cdbl(Dados2("AREA")) > 0 then
                                            precoArea = Dados2("C6_PRCVEN")/Dados2("AREA")
                                        end if
                                    end if
                                else
                                    if peso <> 0 then
                                        PKGPedido = Dados2("C6_PRCVEN")/peso
                                    end if
                                End if

                                PrecoKGPedido = CStr(PKGPedido)
                                PrecoKGPedido = Left(PrecoKGPedido,InStr(1,PrecoKGPedido,",")+5)
                                Percentual = PercPreco(precoPedido,Dados2("C6_PRCVEN"))
%>

                            <tr>
                                <td><%=Dados2("A1_NREDUZ") %></td>
                                <td class="text-center"><%=Dados2("C6_NUM") %></td>
                                <td><%=Dados2("C6_DESCRI") %></td>
                                <td class="text-center"><%=AcertaData(Dados2("C6_DTVALID"))%></td>
                                <td class="text-center"><%=AcertaData(Dados2("C6_ENTREG")) %></td>
                                <td class="text-center"><%=Dados2("C6_TES") %></td>
                                <td class="text-right"><%=Dados2("C6_QTDVEN") %></td>
                                <td class="text-right"><%=Dados2("C6_PRCVEN") %></td>
                                <td class="text-right"><%=PrecoKGPedido %></td>
                                <%if left(produto,3) = "720" then %>
                                <td class="text-right"><%=FormatNumber(precoArea,5) %></td>
                                <%end if %>
                                <td><%=Dados2("E4_DESCRI") %></td>
                                <td class="text-right"><%=FormatNumber(Percentual,0) %>%</td>
                            </tr>
<%
                                Dados2.MoveNext() 
                            wend 
%>
                        </thead>
                    </table>
<%
                        else 
%>
                    <div class="col-xs-10 col-sm-10 col-md-10 col-xs-offset-1 col-sm-offset-1 col-md-offset-1">
                        <div class="alert alert-warning">
                            <strong>Alerta!</strong> Não há pedidos cadastrados.
                        </div>
                    </div>
<%
                        end if 
%>
                </div>
<%
                session("PROD") = produto
                session("pedido") = NumPed
                session("DAT1") = InverteData(dataEntrega1)
                session("DAT2") = InverteData(dataEntrega2)

                if (Dados3.EOF) or (Eval(Dados3("SALDO_CREDITO")) < 0) then
%>
                <input type="hidden" name="radiobutton" value="LC" />
                <input name="CODCLI" type="hidden" id="CODCLI" value="<%=codigoCliente %>" />
                <div class="col-xs-10 col-sm-10 col-md-10 col-xs-offset-1 col-sm-offset-1 col-md-offset-1 alert alert-warning" >
                    <strong>Alerta!</strong> Cliente sem limite de crédito. Clique <a href="#" onclick="BloqueioLimiteCredito()">aqui</a> para concluir o bloqueio.
                </div>
                </div>
                <%Response.end %>
                <%end if %>
                <%Dados3.Close() %>
            </div>
            <hr />

            <div class="row">
                <div class="col-xs-2 col-sm-2 col-md-2 col-xs-offset-1 col-sm-offset-1 col-md-offset-1">
                    <label class="radio-inline"><input type="radio" name="radiobutton" value="PR"<% if precoPedido < ultimoPreco then%> checked="checked" <%end if %> />Bloqueado por preço</label>
                </div>
                <div class="col-xs-2 col-sm-2 col-md-2">
                    <label class="radio-inline"><input type="radio" name="radiobutton" value="LC" />Bloqueado por lim. crédito</label>
                </div>
                <div class="col-xs-2 col-sm-2 col-md-2">
                    <label class="radio-inline"><input type="radio" name="radiobutton" value="DE" />Bloqueado por Data Entrega</label>
                </div>
                <div class="col-xs-2 col-sm-2 col-md-2">
                    <label class="radio-inline"><input type="radio" name="radiobutton" value="CP" />Bloqueado por cond. pagto.</label>
                </div>
                <div class="col-xs-2 col-sm-2 col-md-2">
                    <label class="radio-inline"><input type="radio" name="radiobutton" value="LI"<% if precoPedido < ultimoPreco then%> checked="checked" <%end if %> /> Liberado</label>
                </div>
            </div>
            <%if (Dados("B1_DESCNF") = 901 or Dados("B1_DESCNF") = 902) and Trim(Dados("ZF_LOTEPIL")) <> "S"  and Trim(Dados("ZF_REMESSA")) = "N" and (Trim(Dados("ZF_PROPOST")) = "" or Trim(Dados("ZF_PROPOST")) = "0") and Trim(Dados("ZF_ESTATUS")) <> "LI" then %>
            <div class="row">
                <div class="col-xs-2 col-sm-2 col-md-2 col-xs-offset-1 col-sm-offset-1 col-md-offset-1">
                    <div class="checkbox">
                        <label><input type="checkbox" name="IncluiProposta" id="IncluiProposta" value="1" /><strong>Incluir Proposta para este pedido</strong></label>
                    </div>
                </div>
            </div>
            <%end if %>
            <hr />
            <input name="CODCLI" type="hidden" id="CODCLI" value="<%=codigoCliente %>" />
            <input name="LOJACLI" type="hidden" id="LOJACLI" value="<%=lojaCli%>" />
            <input name="CODCLI2" type="hidden" id="CODCLI2" value="<%=Trim(Dados("zf_destino"))%>" />
            <input name="LOJACLI2" type="hidden" id="LOJACLI2" value="<%=Trim(Dados("zf_lojades"))%>" />
            <input name="PRPEDIDO" type="hidden" id="PRPEDIDO" value="<%=precoPedido%>" />
            <input name="ULTIMOPRECO" type="hidden" id="ULTIMOPRECO" value="<%=ultimoPreco%>" />
            <input name="VPESOTOTAL" type="hidden" id="VPESOTOTAL" value="<%=pesoTotal%>" />
            <input name="TIPOPRODUTO" type="hidden" id="TIPOPRODUTO" value="<%=tipoProduto%>" />
            <input name="ESTRUTURA" type="hidden" id="ESTRUTURA" value="<%=estrutura%>" />
            <input name="Quantidade" type="hidden" id="Quantidade" value="<%=quantidade%>" />
            <input name="Largura" type="hidden" id="Largura" value="<%=Dados("zf_larg")%>" />
            <input name="Comprimento" type="hidden" id="Comprimento" value="<%=Dados("zf_comp")%>" />
            <input name="TES" type="hidden" id="TES" value="<%=TES%>" />
            <div class="row margemTopo">
                <div class="col-xs-10 col-sm-10 col-md-10 col-xs-offset-1 col-sm-offset-1 col-md-offset-1 text-center">
                    <a href="#" class="btn btn-success" onclick="validaBloqueio()">
                        <span class="glyphicon glyphicon-check"></span>
                        Confirmar
                    </a>
                </div>
            </div>
            <%'Else
            '  if Trim(Dados("ZF_PROPOST")) <> "" then
            '      response.write "<div class="" row"">
                '"
                '      response.write "<div class="" col-xs-10 col-sm-10 col-md-10 col-xs-offset-1 col-sm-offset-1 col-md-offset-1"">
                    '"
                    '      response.write "<div class='alert alert-info'>
                        '"
                        '      response.write "<strong>Pedido já Validado!</strong> Este pedido já foi validado automaticamente pela proposta n° "&Trim(Dados("ZF_PROPOST"))&" . Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar."
                        '      response.write "
                    '</div>"
                    '      response.write "
                '</div>"
                '      response.write "
            '</div>"
            '  else
            '      response.write "<div class="" row"">
                '"
                '      response.write "<div class="" col-xs-10 col-sm-10 col-md-10 col-xs-offset-1 col-sm-offset-1 col-md-offset-1"">
                    '"
                    '      response.write "<div class='alert alert-info'>
                        '"
                        '      response.write "<strong>Pedido já Validado!</strong> Este pedido já foi validado. Clique <a href=# onclick='history.back(1);'>aqui</a> para voltar."
                        '      response.write "
                    '</div>"
                    '      response.write "
                '</div>"
                '      response.write "
            '</div>"
            '  end if
            '  response.end
            'End if %>
        </form>
    </div>
</body>
</html>
<%
Function PercPreco(PP,PV)

   if PV = 0 then
		PercPreco = 1
   else
		v_PercPreco = ((PP/PV)*100) - 100
		PercPreco = FormatNumber(((PP/PV)*100) - 100,0)
   end if

End Function

Function ClassificaMargem(valor)

    If Margem < 20 Then
        ClassificaMargem = "danger"
    ElseIf Margem > 40 Then
        ClassificaMargem = "success"
    Else
        ClassificaMargem = "warning"
    End If

end Function

Function GeraConsultaMargem720(larguraChapa,comprimentoChapa,estadoMargem,municipioMargem)

          SQL = "SELECT "
    SQL = SQL &   "ZF_FILIAL, ZF_NUMERO, ZF_CLIENTE, ZF_LOJACLI, ZF_TPFRETE, ZF_CONDPAG, ZF_EMISSAO, ZF_PRODUTO, ZF_QUANT, ZF_PRECO, ZF_ESTRUTU, ZF_AREAUNI, ZF_PESOUNI, "
    SQL = SQL &   "A1_COD, A1_LOJA, A1_NOME, A1_NREDUZ, A1_EST, A1_COD_MUN, A1_MUN, A1_VEND, B1_PICM, B1_IPI, B1_COD, "
    SQL = SQL &   "ISNULL(B5_COMPR,0) B5_COMPR, ISNULL(B5_LARG,0) B5_LARG, ISNULL(B5_ALTU,0) B5_ALTU, ISNULL(B5_ESTILO,'0903') B5_ESTILO, CASE ISNULL(B5_LARGCH,0) WHEN 0 THEN "&larguraChapa&" ELSE B5_LARGCH END B5_LARGCH, CASE ISNULL(B5_COMPCH,0) WHEN 0 THEN "&comprimentoChapa&" ELSE B5_COMPCH END B5_COMPCH, B5_BLANKL, B5_BLANKC, ISNULL(B5_AREA,0) B5_AREA, 'S' B5_FECHAM, "
    SQL = SQL &   "REPLACE(B5_COR1,'S IMP   ','') B5_COR1, REPLACE(B5_COR2,'S IMP   ','') B5_COR2, REPLACE(B5_COR3,'S IMP   ','') B5_COR3, REPLACE(B5_COR4,'S IMP   ','') B5_COR4, "
    SQL = SQL &   "B5_ARECOR1 / B5_AREA / 100 B5_ARECOR1, B5_ARECOR2 / B5_AREA / 100 B5_ARECOR2, B5_ARECOR3 / B5_AREA / 100 B5_ARECOR3, B5_ARECOR4 / B5_AREA / 100 B5_ARECOR4, "
    SQL = SQL &   "CASE ISNULL(B5_QTDPCT,0) WHEN 0 THEN 25 ELSE B5_QTDPCT END B5_QTDPCT , B5_PALETE, ISNULL(B5_TPPALE,'1') B5_TPPALE, B5_MAQ1, B5_MAQ2, B5_MAQ3, B5_MAQ4, B5_MAQ5, B5_MAQ6, 1 B5_ARRL, 1 B5_ARRC, 1 TOTAL_PECAS, "
    SQL = SQL &   "B5_FILME, B5_COBINT, B5_COBEXT, B5_AMARRAD, CASE ISNULL(B5_PCSPAL,0) WHEN 0 THEN 25 ELSE B5_PCSPAL END B5_PCSPAL, isnull(B5_PLLAST,0) B5_PLLAST, isnull(B5_ARRPA,0) B5_ARRPA "
    SQL = SQL & "FROM SZF010 inner join SA1010 on ZF_CLIENTE = A1_COD AND A1_LOJA = ZF_LOJACLI inner join SB1010 ON B1_COD = ZF_PRODUTO LEFT JOIN SB5010 ON B1_COD = B5_COD AND ISNULL(SB5010.D_E_L_E_T_,'') = '' "
    SQL = SQL & "WHERE "
    SQL = SQL &   "SZF010.D_E_L_E_T_ <> '*' AND "
    SQL = SQL &   "SA1010.D_E_L_E_T_ <> '*' AND "
    SQL = SQL &   "SB1010.D_E_L_E_T_ <> '*' AND "
    SQL = SQL &   "ZF_NUMERO = '" & Request("pedido") & "' "
    'SQL = "SELECT LTRIM(RTRIM(A1_COD)),RTRIM(A1_NREDUZ), '"&Municipo(estadoMargem,municipioMargem)&"', A1_EST,RTRIM(B1_COD), " 'Cliente, NomeCliente, IDMunicipio, UF  ,Referencia
    'SQL = SQL &"RTRIM(A1_VEND),'',ISNULL(B5_MAQ1,'041'), ISNULL(B5_MAQ2,''), ISNULL(B5_MAQ3,''), ISNULL(B5_MAQ4,''), " 'CodRepresentante, CodFechamento, Maquina1, Maquina2,Maquina3,Maquina4,
    'SQL = SQL &"'','','','', " 'Cor1, Cor2, Cor3, Cor4, 
    'SQL = SQL & "'0','0','0','0'," 'ConsumoCor1, ConsumoCor2,ConsumoCor3,ConsumoCor4
    'SQL = SQL & "LTRIM(RTRIM(ISNULL(B5_TPPALE,''))),ZF_ESTRUTU,ISNULL(B5_ESTILO,''), CASE ISNULL(B5_LARGCH,0) WHEN 0 THEN "&larguraChapa&" ELSE B5_LARGCH END, CASE ISNULL(B5_COMPCH,0) WHEN 0 THEN "&comprimentoChapa&" ELSE B5_COMPCH END B5_COMPCH, " 'IDPalete, Composicao, TipoABNT,Largura,Comprimento,
    'SQL = SQL & "ISNULL(B5_LARG,0), ISNULL(B5_COMPR,0), ISNULL(B5_ALTU,0), ISNULL(B5_AREA,0), "  'LarguraInterna,ComprimentoInterno,AlturaInterna,AreaBrutaPeca
    'SQL = SQL & "0,0,0,@Item,@Conjunto, " ',Usuario,IsentoICMS,IsentoIPI,Item,Conjunto
    'SQL = SQL & "1, ISNULL(B5_ARRL,1), ISNULL(B5_ARRC,1), ISNULL(TOTAL_PECAS,1), CASE ISNULL(B5_PALETE,'') WHEN 'S' THEN 1 ELSE 0 END"  'Estado,MultLarg,MultComp,PecasPorConjunto,Paletizado,
    'SQL = SQL & "CASE SUBSTRING(ZF_ESTRUTU,10,1) WHEN 2 THEN 1 ELSE CASE WHEN SUBSTRING(ZF_ESTRUTU,10,1) > 1 THEN 1 ELSE SUBSTRING(ZF_ESTRUTU,10,1) END END, CASE SUBSTRING(ZF_ESTRUTU,8,1) WHEN 2 THEN 1 ELSE CASE WHEN SUBSTRING(ZF_ESTRUTU,8,1) > 1 THEN 1 ELSE SUBSTRING(ZF_ESTRUTU,8,1) END END, " 'ResinaInterna,ResinaExterna,
    'SQL = SQL & "'1',CASE ISNULL(B5_QTDPCT,0) WHEN 0 THEN 25 ELSE B5_QTDPCT END , CASE ISNULL(B5_PCSPAL,0) WHEN 0 THEN 25 ELSE B5_PCSPAL END, " 'EmpilhamentoMaximo,PecasPorPacote,PecasPorUnidade
    'SQL = SQL & "ISNULL(B5_PLLAST,0), isnull(B5_ARRPA,0), CASE ISNULL(B5_FILME,'') WHEN 'S' THEN 1 ELSE 0 END" 'PacotesLargura, PacotesComprimento, PacotesAltura, Filme 
   
    Dados01.Open(SQL)
    
    if Not Dados01.Eof then
        If Dados01("B5_PALETE") = "S" Then
            Palete = "1"
        Else
            Palete = "0"
        End If
        If Dados01("B5_FILME") = "S" Then
            Filme = "1"
        Else
            Filme = "0"
        End If
        resinaInternaConsulta = ResinaInterna(Trim(Dados01("ZF_ESTRUTU")))
        resinaExternaConsulta = ResinaExterna(Trim(Dados01("ZF_ESTRUTU")))

        if resinaInternaConsulta > "1" then
            resinaInternaConsulta = 1
        end if

        if resinaExternaConsulta > "1" then
            resinaExternaConsulta = 1
        end if
        SQL = "VALUES('" & Trim(Dados01("A1_COD")) & "','" & Trim(Dados01("A1_NREDUZ")) & "','"&Municipo(estadoMargem,municipioMargem)&"','" & Dados01("A1_EST") & "','" & Trim(Dados01("B1_COD")) & "', "                                             'Cliente, NomeCliente, IDMunicipio, UF  ,Referencia
        SQL = SQL & "'" & Trim(Dados01("A1_VEND")) & "','','041','','','', "        
        SQL = SQL & "'','','','',"                                                                             'Cor1, Cor2, Cor3, Cor4, 
        SQL = SQL & "'0','0','0','0'," 'ConsumoCor1, ConsumoCor2,ConsumoCor3,ConsumoCor4
        SQL = SQL & "'" & Trim(Dados01("B5_TPPALE")) & "','" & Dados01("ZF_ESTRUTU") & "','" & Dados01("B5_ESTILO") & "'," & Dados01("B5_LARGCH") & "," & Dados01("B5_COMPCH") & ","                                                       'IDPalete, Composicao, TipoABNT,Largura,Comprimento,
        SQL = SQL & "" & Dados01("B5_LARG") & "," & Dados01("B5_COMPR") & "," & Dados01("B5_ALTU") & "," & Replace(Dados01("B5_AREA"),",",".") & ","   'LarguraInterna,ComprimentoInterno,AlturaInterna,AreaBrutaPeca
        SQL = SQL & "0,0,0,@Item,@Conjunto,"  ',Usuario,IsentoICMS,IsentoIPI,Item,Conjunto
        SQL = SQL & "1," & Dados01("B5_ARRL") & "," & Dados01("B5_ARRC") & "," & Dados01("TOTAL_PECAS") & "," & Palete & ","                                                                                           'Estado,MultLarg,MultComp,PecasPorConjunto,Paletizado,
        SQL = SQL & "" & resinaInternaConsulta &","& resinaExternaConsulta &",'1'," & Dados01("B5_QTDPCT") & "," & Dados01("B5_PCSPAL") & ", " 'ResinaInterna,ResinaExterna,EmpilhamentoMaximo,PecasPorPacote,PecasPorUnidade
        SQL = SQL & "" & Dados01("B5_PLLAST") & ",1," & Dados01("B5_ARRPA") & "," & Filme & ") "     'PacotesLargura, PacotesComprimento (sempre 1), PacotesAltura, Filme (sempre 1)  
    
        GeraConsultaMargem720 = SQL
    Else
        GeraConsultaMargem720 = ""
    end if

    Dados01.Close()

End function

function ResinaInterna(composicao)

    if Mid(composicao,10,1) = "2" then
        ResinaInterna = "1"
    else
        ResinaInterna = Mid(composicao,10,1)
    end if

end function

function ResinaExterna(composicao)

    if Mid(composicao,8,1) = "2" then
        ResinaExterna = "1"
    else
        ResinaExterna = Mid(composicao,8,1)    
    end if

end function

function Municipo(estadoMargem,municipioMargem)

    SQL = "SELECT TOP 1 IDMunicipio From vw_ConsultaMunicipiosPelaInpaNet WHERE UF = '" & estadoMargem & "' AND Municipio = '" & municipioMargem & "' "

    consulta.Open(SQL)

    If Not consulta.EOF Then

        Municipo = consulta("IDMunicipio")

    else

        Municipo = 0

    End If

    consulta.Close()

end function
%>