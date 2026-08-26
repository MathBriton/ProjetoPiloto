<!--#include file = "../cnn2/cnn2.asp"-->
<!--#include file = "../includes/variaveis.asp" -->
<!--#include file = "../funcoes/funcoes.asp" -->
<%
  if (trim(request("cliente")) <> "") or (trim(request("tipo")) = "Grupo") then

    if trim(request("codsiga")) = "" and trim(request("nivel")) <> "0" and trim(request("nivel")) <> "" then

      response.write "ErroLogof"
      response.End
      
    end if

    'loja = replace(trim(request("loja")),"'","")

    codsiga = trim(request("codsiga"))
    Nivel = trim(request("nivel"))

    if trim(request("nivel")) <> "" then

      select case request("planta")
      case "05" 'BENTO ----------------------------------------------------------------------------------------------------------------------------------

         SQL = "" &_
            "SELECT " &_
              "T1.COD_CLIENTE AS A1_COD, " &_ 
              "T1.LOJA_CLIENTE AS A1_LOJA, " &_ 
              "COD_SAP_CLIENTE AS A1_CODSAP, " &_ 
              "'2' AS A1_MSBLQL, " &_ 
              "T1.NOME_CLIENTE AS A1_NOME, " &_ 
              "T1.FANTASIA_CLIENTE AS A1_NREDUZ, " &_ 
              "T1.MUNICIPIO_CLIENTE AS A1_MUN, " &_ 
              "T1.ESTADO_CLIENTE AS A1_EST, " &_
              "T1.A1_BAIRRO AS  A1_BAIRRO, " &_ 
              "T1.COD_VENDEDOR AS A1_VEND, " &_ 
              "T1.CNPJ_CLIENTE AS Cnpj, " &_ 
              "T1.CONDICAO_PAGAMENTO, " &_ 
              "T1.DIAS_TRANSPORTE, " &_ 
              "CASE WHEN T2.TIPO_FRETE = 'C' THEN '0' ELSE '1' END as TIPO_FRETE " &_
            "FROM VW_BI_CLIENTES T1 " &_ 
            "LEFT JOIN VW_CLIENTES_PEDIDO T2 ON T2.COD_CLIENTE = T1.COD_CLIENTE " &_       
            "WHERE T1.CODIGO_ERP = '" & trim(Request("cliente")) & "' "  &_
              "AND T1.STATUS = 'ATIVO' " &_
              iif( (Nivel <> 0), "AND T1.COD_VENDEDOR = '"&codsiga&"' ", "" ) &_ 
            "ORDER BY T1.FANTASIA_CLIENTE "

      case "00", "04" 'PIRAPETINGA, UBERABA  ------------------------------------------------------------------------------------------------------------

        SQL = "" &_
          "SELECT " &_
            "A1_COD, " &_
            "A1_LOJA, " &_
            "A1_CODSAP, " &_
            "A1_MSBLQL, " &_
            "A1_NOME, " &_
            "A1_NREDUZ, " &_
            "A1_MUN, " &_
            "A1_EST, " &_
            "A1_BAIRRO, " &_
            "A1_VEND, " &_
            "A1_CGC Cnpj, " &_
            "0 as CONDICAO_PAGAMENTO " &_
          "FROM SA1010 " &_
          "WHERE A1_COD = '"&trim(Request("cliente"))&"' " &_
            "AND D_E_L_E_T_ = '' " &_
            "AND A1_MSBLQL <> '1' AND A1_LOJA = '" &trim(Request("loja"))& "' " &_
            iif( (Nivel <> 0), "AND A1_VEND = '" & codsiga & "' ", "" ) &_
          "ORDER BY A1_NREDUZ "

      case "06" 'FORTALEZA ------------------------------------------------------------------------------------------------------------------------------

          SQL = "" &_
            "SELECT " &_
              "T1.COD_CLIENTE AS A1_COD, " &_ 
              "T1.LOJA_CLIENTE AS A1_LOJA, " &_ 
              "COD_SAP_CLIENTE AS A1_CODSAP, " &_ 
              "'2' AS A1_MSBLQL, " &_ 
              "T1.NOME_CLIENTE AS A1_NOME, " &_ 
              "T1.FANTASIA_CLIENTE AS A1_NREDUZ, " &_ 
              "T1.MUNICIPIO_CLIENTE AS A1_MUN, " &_ 
              "T1.ESTADO_CLIENTE AS A1_EST, " &_
              "T1.Bairro AS  A1_BAIRRO, " &_ 
              "T1.COD_VENDEDOR AS A1_VEND, " &_ 
              "T1.CNPJ_CLIENTE AS Cnpj, " &_ 
              "T1.CONDICAO_PAGAMENTO, " &_ 
              "T1.DIAS_TRANSPORTE, " &_ 
              "CASE WHEN T2.TIPO_FRETE = 'C' THEN '0' ELSE '1' END as TIPO_FRETE " &_
            "FROM VW_BI_CLIENTES T1 " &_ 
            "LEFT JOIN VW_CLIENTES_PEDIDO T2 ON T2.COD_CLIENTE = T1.COD_CLIENTE " &_       
            "WHERE T1.CODIGO_ERP = '" & trim(Request("cliente")) & "' "  &_
              "AND T1.STATUS = 'ATIVO' " &_
              iif( (Nivel <> 0), "AND T1.COD_VENDEDOR = '"&codsiga&"' ", "" ) &_ 
            "ORDER BY T1.FANTASIA_CLIENTE "

      case "07" 'SAQUAREMA ------------------------------------------------------------------------------------------------------------------------------'

        SQL = "" &_
          "SELECT " &_
            "COD_CLIENTE AS A1_COD, " &_
            "LOJA_CLIENTE AS A1_LOJA, " &_
            "'COD_SAP_CLIENTE' AS A1_CODSAP, " &_
            "'2' AS A1_MSBLQL, " &_
            "NOME_CLIENTE AS A1_NOME, " &_
            "FANTASIA_CLIENTE AS A1_NREDUZ, " &_
            "MUNICIPIO_CLIENTE AS A1_MUN, " &_
            "ESTADO_CLIENTE AS A1_EST, " &_
            "'' AS  A1_BAIRRO, " &_
            "COD_VENDEDOR AS A1_VEND, " &_
            "CNPJ_CLIENTE AS Cnpj, " &_
            "CONDICAO_PAGAMENTO " &_
          "FROM VW_BI_CLIENTES " &_
          "WHERE COD_CLIENTE = '" & trim(Request("cliente")) & "' " &_
            "AND STATUS = 'ATIVO' " &_
            iif( (Nivel <> 0), "AND COD_VENDEDOR = '"&codsiga&"' ", "" ) &_
          "ORDER BY FANTASIA_CLIENTE "

      case else

        SQL = "" &_
          "SELECT " &_
            "NOME_CLIENTE AS A1_NOME, " &_
            "COD_CLIENTE AS A1_COD, " &_
            "COD_SAP_CLIENTE AS A1_CODSAP, " &_
            "COD_VENDEDOR AS A1_VEND, " &_
            "'2' as A1_MSBLQL, " &_
            "LOJA_CLIENTE AS A1_LOJA, " &_
            "MUNICIPIO_CLIENTE AS A1_MUN, " &_
            "ESTADO_CLIENTE AS A1_EST, " &_
            "'' AS A1_BAIRRO, " &_
            "CNPJ_CLIENTE AS Cnpj, " &_
            "0 CONDICAO_PAGAMENTO "  &_
          "FROM dbo.VW_CLIENTES Clientes " &_
          "WHERE STATUS = 'ATIVO' " &_
            "AND cast(COD_CLIENTE as varchar(20)) = '" & trim(Request("parametro")) & "' " &_
            iif( (Nivel <> 0), "AND COD_VENDEDOR = '" & codsiga & "' ", "" ) &_
          "ORDER BY NOME_CLIENTE, LOJA_CLIENTE "

      end select 
      

      ' SQL = "" &_
      '   "SELECT " &_
      '     "TOP 1 A1_NOME " &_
      '   "FROM SA1010 " &_
      '   "WHERE A1_COD = '"&Trim(request("cliente"))&"' " &_
      '     "AND A1_CODSAP <> '' " &_
      '     "AND D_E_L_E_T_ <> '*' " &_
      '     "AND A1_MSBLQL <> '1' " &_
      '     iif( trim(request("tipo")) <> "Grupo", "AND A1_LOJA = '" & loja & "' ", "" ) &_
      '     iif( (trim(request("nivel")) = "1" and trim(request("tipo")) <> "Remessa"), "AND A1_VEND = '" & trim(request("codsiga")) & "' ", "" ) &_
      '     "AND 1 = 1 "
      
      select case request("planta")
      case "00" SET Consulta = Conecta( "PIR-DADOSADV", SQL )
      case "04" SET Consulta = Conecta( "PIR-DADOSADV", SQL )
      case "05" SET Consulta = Conecta( "BEN-SOLPEN", SQL )
      case "06" SET Consulta = Conecta( "FOR-DADOSADV", SQL )
      case "07" SET Consulta = Conecta( "SAQ-DADOSADV", SQL )
      case else SET Consulta = Conecta( "SPO-BI", SQL )
      end select

      if vartype(Consulta) = 0 or vartype(Consulta) = 1 then 
        response.end
      end if                        

      if not Consulta.Eof then

        response.write "|" & trim(request("tipo")) & "|" & trim(Consulta("A1_NOME"))

        Consulta.Close()
        SET Consulta = NOTHING

      else

        response.write "ErroNEncontrado|" & trim(request("tipo"))

      end if

    else

      response.write "ErroLogof|" & trim(request("tipo"))

    end if

  else

    response.write "ErroCliente"

  end if
%>
