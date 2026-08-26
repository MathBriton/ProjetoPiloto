<!-- #include file = "../../cnn2/cnn2.asp" -->
<!-- #include file = "../../funcoes/funcoes.asp" -->
<!-- #include file = "../../includes/variaveis.asp" -->
<%
  if (trim(request("cliente")) <> "" and trim(request("loja")) <> "") or (trim(request("cliente")) <> "" and trim(request("tipo")) = "Grupo") then

    if trim(request("codsiga")) = "" and trim(request("nivel")) <> "0" and trim(request("nivel")) <> "" then
      response.write "ErroLogof"
      response.End
    end if

    loja = REPLACE(trim(request("loja")),"'","")

    select case request("planta")
    case "00" 'PIRAPETINGA, UBERABA  ------------------------------------------------------------

    SQL = "" &_
      "SELECT " &_
        "A1_NOME AS NOME_CLIENTE " &_
      "FROM SA1010 " &_
      "WHERE A1_COD = '" & Trim(request("cliente")) & "' " &_
        "AND D_E_L_E_T_ = '' " &_
        "AND A1_MSBLQL <> '1' " &_
        iif( (trim(request("nivel")) <> 0), "AND A1_VEND = '" & codsiga & "' ", "" ) &_
      "ORDER BY A1_NREDUZ "

    case "05", "06" 'BENTO - MARANGUAPE ---------------------------------------------------------

        SQL = "" &_
          "SELECT " &_
            "T1.NOME_CLIENTE " &_ 
          "FROM VW_BI_CLIENTES T1 " &_ 
          "LEFT JOIN VW_CLIENTES_PEDIDO T2 ON T2.COD_CLIENTE = T1.CODIGO_ERP " &_       
          "WHERE T1.CODIGO_ERP = " & trim(request("cliente")) & " " &_
            "AND T1.STATUS = 'ATIVO' " &_
            iif( (trim(request("nivel")) <> 0), "AND T1.COD_VENDEDOR = '" & codsiga & "' ", "" ) &_ 
          "ORDER BY T1.FANTASIA_CLIENTE "

    case "07" 'SAQUAREMA ------------------------------------------------------------------------

      SQL = "" &_
        "SELECT " &_
          "NOME_CLIENTE " &_
        "FROM VW_BI_CLIENTES " &_
        "WHERE COD_CLIENTE = '" & trim(request("cliente")) & "' " &_
          "AND STATUS = 'ATIVO' " &_
          iif( (trim(request("nivel")) <> 0), "AND COD_VENDEDOR = '" & codsiga & "' ", "" ) &_
        "ORDER BY FANTASIA_CLIENTE "

    end select 

    'RESPONSE.WRITE SQL

    select case request("planta")
    case "00" SET CONSULTA = Conecta( "PIR-DADOSADV", SQL )
    case "04" SET CONSULTA = Conecta( "PIR-DADOSADV", SQL )
    case "05" SET CONSULTA = Conecta( "BEN-DADOSADV", SQL )
    case "06" SET CONSULTA = Conecta( "FOR-DADOSADV", SQL )
    case "06" SET CONSULTA = Conecta( "FOR-DADOSADV", SQL )
    case "07" SET CONSULTA = Conecta( "SAQ-DADOSADV", SQL )
    case else SET CONSULTA = Conecta( "SPO-BI", SQL )
    end select

    if not CONSULTA.Eof then
      response.write "|" & trim(request("tipo")) & "|" & trim(CONSULTA("NOME_CLIENTE"))
    else
      response.write "ErroNEncontrado|" & trim(request("tipo"))
    end if

  else

    response.write "ErroCliente"

  end if
%>
