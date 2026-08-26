<%session.LCID = 1046 %>
<!--#include file = "../cnn/DADOSADV.asp"-->
<!--#include file = "../funcoes/funcoes.asp"-->
<%
validou = 1
PesoCalc = 0
AreaCalc = 0
msg = ""
msgvalidou = ""
largura=request("larg")
comprimeto=request("comp")

if trim(request("cliente")) <> "" and trim(request("loja")) <> "" and trim(request("produto")) <> "" and trim(request("quantidade")) <> "" and trim(request("condicao")) <> "" then
    
  if trim(request("codsiga")) = "" and trim(request("nivel")) <> "" then
    response.write "ErroLogof|"&trim(request("tipo"))
    response.End
  end if

  if trim(request("nivel")) <> "" then

    if trim(request("condicao")) = "CH" and (trim(request("larg")) = "" or trim(request("comp")) = "") then  
      response.write "ErroParametro|"&trim(request("tipo"))
      response.end
    end if
    
    if Trim(request("filial")) <> "" then
      filial = request("filial")    'iif( (Trim(request("filial")) = "00"), "I", "U" )
    else
      response.write "ErroParametro|"&trim(request("tipo"))
      response.end
    end if

    ' VERIFICA A MEDIA DE QTDE COMPRADAS ------------------------------------------------- CHPL 374
    ' qtdeMediaVendida = verMediaDeQuantidadeVendida( trim(request("cliente")), trim(request("loja")), trim(request("produto")), 10 ) 
    ' if Cdbl(trim(request("quantidade"))) > qtdeMediaVendida then

    '   validou = 0 
    '   response.write "|"&trim(request("tipo"))&"|"&validou&"|A Quantidade Média Vendida é de : " & formatnumber(qtdeMediaVendida,0)
    '   response.End

    ' else

    '   validou = 1
    '   response.write "|"&trim(request("tipo"))&"|"&validou&"|"

    ' end if

    ' ---------------------------------------------------------------------------------------------
    SQL = "SELECT " &_
            "A1_CLASCLI " &_
          "FROM SA1010 " &_
          "WHERE A1_COD = '"&trim(request("cliente"))&"'  "
    Dados3.Open(SQL)      
    A1_CLASCLI = iif( (not Dados3.EOF), Trim(Dados3("A1_CLASCLI")), "" )
    Dados3.Close()

    paramPesoMN = 0
    paramPesoCV = 0
    paramPesoCH = 0

    SQL = "SELECT " &_
            "Dias, " &_
            "DiasViagem, " &_
            "CargaMaquina, " &_
            "DiasAcrescimoPA, " &_
            "DiasAcrescimoCX, " &_
            "DiasAcrescimoCH, " &_
            "PesoMN, " &_
            "PesoCV, " &_
            "PesoCH, " &_
            "MtsCVmin, " &_
            "MtsCXmin, " &_
            "MtsCHmin " &_
          "FROM CORPORATIVOHMG..PARAMETROS_PEDIDO " &_
          "WHERE Planta = " & filial
    Dados3.Open(SQL)            

    if Not Dados3.Eof then
      paramPesoMN = Dados3("PesoMN")
      paramPesoCV = Dados3("PesoCV")
      paramPesoCH = Dados3("PesoCH")

      paramAreaCV = Dados3("MtsCVmin")
      paramAreaMN = Dados3("MtsCXmin")  'maleta
      paramAreaCH = Dados3("MtsCHmin")
    else
       response.write "FINAL DE ARQUIVO"    
    end if
    Dados3.Close()
    'response.write "paramAreaCV " & paramPesoMN & "paramAreaMN " & paramPesoCV  & "paramAreaCH" & paramPesoCH
      
    select case trim(request("condicao"))
        case "CX" ' ---------------------------------------------------------------------------------

            SQL = "SELECT dbo.funcCalculaPeso ('"&trim(request("produto"))&"','','C') AS PESO"
            Dados.Open(SQL)
            PesoCalc = Cdbl(trim(request("quantidade"))) * Dados("PESO")
            Dados.Close()

        case "CH" ' ---------------------------------------------------------------------------------

            SQL = "SELECT dbo.CalculaPesoCHAPA ('"&trim(request("produto"))&"',"&Cdbl(trim(request("larg")))&","&Cdbl(trim(request("comp")))&") AS PESO"
            Dados.Open(SQL)
            PesoCalc = Cdbl(trim(request("quantidade"))) * Dados("PESO")
            Dados.Close()

        case "PA" : PesoCalc = trim(request("quantidade"))
        case "M2" : PesoCalc = trim(request("quantidade"))
    end select   

    if trim(request("condicao")) <> "CH" and trim(request("condicao")) <> "PA"  Then
        AreaCalc = CalculaArea(trim(request("produto")))
    end if    

    'if A1_CLASCLI = "P1" then
    
      SQL = "SELECT " &_
              "B1_DESCNF AS TIPO, " &_
              "B1_ESTRUCH AS ESTRUTURA " &_
            "FROM SB1010 " &_
            "WHERE B1_COD = '"&trim(request("produto"))&"' " &_
              "AND D_E_L_E_T_ <> '*'"
      Dados2.Open(SQL)
      ' Verificando se para CAIXA o peso está abaixo do estipulado ------------------------------
      if trim(request("condicao")) = "CX" then

        '* Se for MALETA ------------------------------------------------------------------------
        if (Dados2("TIPO") = "901" or Trim(Dados2("TIPO")) = "") and PesoCalc < paramPesoMN then
          validou = 0
          msg = "|"&trim(request("tipo"))&"|"&validou&"|Peso Calc.: "&formatnumber(PesoCalc,0)&"Kg. Peso Minimo: "&paramPesoMN&"Kg"
          
          if  cdbl(formatnumber(AreaCalc,0)) < cdbl(formatnumber(paramAreaMN,0))  then
              msg = msg  & " Area Calc.: "&formatnumber(AreaCalc,0)&" Mts. Area Minima: "&paramAreaMN&" Mts"
          end if 
          
          response.write msg 
        end if
       
        '* Se for CORTE E VINCO
        'Alterado para 2500 KG em 13/04/2015 por Sandro, conforme solicitação de Carlos Sobrinho, Wagner, Eduardo Couto
        'if Dados2("TIPO") = "902" And PesoCalc < 2000 Then

        if (Dados2("TIPO") = "902" or Trim(Dados2("TIPO")) = "") And PesoCalc < paramPesoCV Then
          validou = 0
          msg =  "|"&trim(request("tipo"))&"|"&validou&"|Peso Calculado: "&formatnumber(PesoCalc,0)&"Kg. Peso Minimo: "&paramPesoCV&"Kg"
          
          if  cdbl(formatnumber(AreaCalc,0)) < cdbl(formatnumber(paramAreaCX,0))  then
              msg = msg  & " Area Calc.: "&formatnumber(AreaCalc,0)&" Mts. Area Minima: "&paramAreaCX&" Mts"
          end if 
          
          response.write msg 
        end if  
      end if

      ' Verificando se para CHAPA o peso está abaixo do estipulado --------------------------------
      if trim(request("condicao")) = "CH" Then
         largura=request("larg")
         comprimento=request("comp")
         AreaCalc = 0
         if largura <> "" and largura <> 0 and comprimento <> "" and comprimento <> 0 then 
             AreaCalc = Cdbl(trim(request("quantidade"))) * ((largura/1000 * comprimento/1000))
         end if     

        '* Se for TRIPLEX ( ULTIMA ATUALIZAÇÃO SOLICITADA POR CARLOS SOBRINHO EM 24/09/2012 ) EDUARDO COUTO
        if mid(Dados2("ESTRUTURA"),3,2) = "BC" and PesoCalc < 500 Then   
          validou = 0 
          msg = "|"&trim(request("tipo"))&"|"&validou&"|Peso Calculado: "&formatnumber(PesoCalc,0)&"Kg. Peso Minimo: "&paramPesoCH&"Kg"
        end if
        '* Se for DUPLEX
        if mid(Dados2("ESTRUTURA"),3,2) <> "BC" and PesoCalc < paramPesoCH Then   
          validou = 0 
          msg = "|"&trim(request("tipo"))&"|"&validou&"|Peso Calculado: "&formatnumber(PesoCalc,0)&"Kg. Peso Minimo: "&paramPesoCH&"Kg"
        end if
        
        if cdbl(formatnumber(AreaCalc,0)) < cdbl(formatnumber(paramAreaCH,0))  Then   
           validou = 0 
           if msg = "" then 
              msg  = "|"&trim(request("tipo"))&"|"&validou&"|"
           end if
           msg = msg  & "Area Calc.: "&formatnumber(AreaCalc,0)&"Mts. Area Minima: "&paramAreaCH&"Mts"
        else    
           'msg =  "Area Calculada: "&formatnumber(AreaCalc,0)
        end if
        if msg <> "" then 
            response.write msg 
        end if 

      end if
      if validou <> 0 then
         SQL = "SELECT dbo.FUNC_QUANTIDADE_PALETE('"&trim(request("produto"))&"',"&Cdbl(trim(request("quantidade")))&") AS QTD_PALETE"
         Dados3.Open(SQL)
         QtdPalete = iif( (not Dados3.EOF), Trim(Dados3("QTD_PALETE")), "0" )
         Dados3.Close()

         response.write "|"&trim(request("tipo"))&"|"&validou&"|Peso Calculado: "&formatnumber(PesoCalc,0)&"Kg.  Area Calc.: "&formatnumber(AreaCalc,0)&" Mts. Qtd. Paletes: "&QtdPalete
      end if

    'else
    '  validou = 1
    '  response.write "|"&trim(request("tipo"))&"|"&validou&"|Cliente P1 - Peso Calculado: "&formatnumber(PesoCalc,0)&"Kg."   
    'end if
  else
    response.write "ErroLogof|"&trim(request("tipo"))
  end if
else
  response.write "ErroParametro|"&trim(request("tipo"))
end if

' VERIFICA A MEDIA DE QUANTIDADE DE VENDAS DO PRODUTO PARA O CLIENTE ------------------------------
function verMediaDeQuantidadeVendida( pCliente, pLoja, pProduto, pPercentual )

    SQL = "SELECT " &_
            "ROUND(AVG(C6_QTDVEN),0) AS MEDIA " &_
          "FROM SC6010 " &_
          "WHERE C6_CLI = '" & pCliente & "' " &_
            "AND C6_LOJA = '" & pLoja & "' " &_
            "AND C6_PRODUTO = '" & pProduto & "' " &_
            "AND D_E_L_E_T_ = '' " 
    Dados.Open(SQL)

    percentual    = (pPercentual / 10000)
    mediaDeVenda  = Dados("MEDIA")

    if mediaDeVenda > 0 then 

      mediaMaisPerc = (mediaDeVenda * percentual) * 100
      qtdeMaisPerc  = mediaDeVenda + mediaMaisPerc
    
    else
      
      qtdeMaisPerc = 0

    end if

    Dados.Close()

    verMediaDeQuantidadeVendida = qtdeMaisPerc

end function

function CalculaArea(produto)

      AreaTotal = 0
      SQL = "SELECT SUM(isnull(AREA,0)) as AREA from VW_BI_PRODUTOS where CIDADE = 'PIRAPETINGA' AND  COD_CONJUNTO = '" & produto & "'"
      Dados.Open(SQL)
      if not Dados.eof then 
          AreaTotal = Cdbl(trim(request("quantidade"))) * Dados("AREA")
      end if     
      CalculaArea = AreaTotal
      'response.write "  AreaTotal " & AreaTotal
      Dados.Close()

end function 
%>
