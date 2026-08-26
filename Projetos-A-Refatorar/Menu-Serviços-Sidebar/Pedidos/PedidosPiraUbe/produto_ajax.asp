<!--#Include File = "../cnn2/cnn2.asp"-->
<%
if Trim(request("produto")) <> "" then

  if Trim(request("codsiga")) = "" and Trim(request("nivel")) <> "0" and Trim(request("nivel")) <> "" then

    response.write "ErroLogof"
    response.End

  end if

  if Trim(request("nivel")) <> "" then

    if Trim(Request("tppedido")) <> "CX" THEN ' TIPO CHAPA OU PAPEL -------------------------------

      SQL = "" &_
        "SELECT " &_
          "B1_COD, " &_
          "ISNULL(B5_CLI,'') B5_CLI, " &_
          "CASE WHEN B1_INPUBE = 'U' THEN " &_
            "'UBERABA' " &_
          "ELSE " &_
            "'PIRAPETINGA' " &_
          "END AS B1_INPUBE, " &_
          "ISNULL(B5_COMPR,'') B5_COMPR, " &_
          "ISNULL(B5_LARG,'') B5_LARG, " &_
          "ISNULL(B5_LARGCH,'') B5_LARGCH, " &_
          "ISNULL(B5_COMPCH,'') B5_COMPCH, " &_
          "ISNULL(B5_VINCO1,'') B5_VINCO1, " &_
          "ISNULL(B5_VINCO2,'') B5_VINCO2, " &_
          "ISNULL(B5_VINCO3,'') B5_VINCO3, " &_
          "ISNULL(B5_VINCO4,'') B5_VINCO4, " &_
          "ISNULL(B5_VINCO5,'') B5_VINCO5, " &_
          "B1_DESC AS DESCRICAO, " &_
          iif( (Trim(Request("tppedido")) = "CH" OR Trim(Request("tppedido")) = "PA" OR Trim(Request("tppedido")) = "M2"), "CASE WHEN RTrim(ISNULL(B5_ESTRUTU,'')) = '' THEN B1_ESTRUCH ELSE B5_ESTRUTU END AS ESTRUTURA, ", "") &_
          iif( (Trim(Request("tppedido")) = "CH" OR Trim(Request("tppedido")) = "PA" OR Trim(Request("tppedido")) = "M2"), "0 AS PRECO, ", "" ) &_
          "ISNULL(B5_PALETE,'') B5_PALETE, " &_
          "ISNULL(B5_TPPALE,'') B5_TPPALE, " &_
          "ISNULL(B5_DESCPAL,'') B5_DESCPAL " &_
        "FROM	SB1010 SB1 " &_
          "LEFT JOIN SB5010 SB5 ON B5_COD  = B1_COD AND ISNULL(SB5.D_E_L_E_T_,'') = '' AND 1=1 " &_
        "WHERE (ISNULL(SB1.D_E_L_E_T_,'') = '')	" &_
          "AND (SB1.B1_STATUS IN ('V','A','W'))	" &_
          "AND (B1_MSBLQL <> '1')	" &_
          "AND (B1_COD = '"&Trim(request("produto"))& "') " &_
          "AND SUBSTRING(SB1.B1_COD,1,2) <> '77' " &_
          iif( (Trim(Request("tppedido")) = "CH"), "AND (B1_GRUPO =  '72') AND (ISNULL(B5_CLI,'') = '"&Trim(request("cliente"))& "' OR ISNULL(B5_CLI,'') = '') ", "" ) &_
          iif( (Trim(Request("tppedido")) = "PA" OR Trim(Request("tppedido")) = "M2"), "AND (B1_GRUPO =  '70') ", "" ) &_
        "ORDER BY B1_INPUBE, B1_DESC "

    else ' TIPO CAIXA -----------------------------------------------------------------------------

      SQL = "" &_
        "SELECT DISTINCT " &_
          "SB1010.B1_DESC AS DESCRICAO, " &_
          "ISNULL(B5_CLI,'') B5_CLI, " &_
          "SB1010.B1_COD , " &_
          "MIN(B5_ESTRUTU) AS ESTRUTURA, " &_
          "ISNULL((SELECT TOP 1 C6_PRCVEN " &_
                  "FROM SC6010 " &_
                  "WHERE C6_PRODUTO = B1_COD " &_
                    "AND C6_CLI = '" & Trim(request("cliente")) & "' " &_
                    "AND D_E_L_E_T_ <> '*' " &_
                  "ORDER BY C6_DTVALID DESC),0) AS PRECO, " &_
          "CASE WHEN B1_INPUBE = 'U' THEN " &_
            "'UBERABA' " &_
          "ELSE " &_
            "'PIRAPETINGA' " &_
          "END AS B1_INPUBE " &_
        "FROM SB5010 " &_
          "RIGHT OUTER JOIN SB1010 ON SUBSTRING (SB5010.B5_COD, 4, 6) = SUBSTRING (SB1010.B1_COD, 4, 6) " &_
                                 "AND (SB5010.B5_TIPO = 'C' OR SB5010.B5_TIPO = 'D' OR SB5010.B5_TIPO = 'S' OR SB5010.B5_TIPO = 'T') " &_
        "WHERE (SB1010.D_E_L_E_T_ <> '*') " &_
          "AND SB1010.B1_STATUS IN ('V','A','W') "  &_
          "AND SB1010.B1_MSBLQL <> '1' " &_
          "AND (SB1010.B1_COD ='"&Trim(request("produto"))& "') " &_
          "AND SUBSTRING(ISNULL(B1_COD,''),1,2) <> '77' " &_
          "AND B1_SUBGRUP NOT IN ('750','760') " &_
          iif( (Trim(request("eshop")) = "0"), "AND SB5010.B5_CLI = '"&Trim(request("cliente"))& "' ", "AND (SB5010.B5_CLI IN ('9099', '9990') OR ISNULL(B5_CLI,'') = '') " ) &_
        "GROUP BY SB1010.B1_DESC, " &_
                 "ISNULL(B5_CLI,''), " &_
                 "SB1010.B1_COD, " &_
                 "B1_INPUBE " &_
      "--UNION " &_
      "--AQUI PEGO SÓ OS PRODUTOS 750 ---------------------------------------------------------------------------------------------- " &_
      "SELECT " &_
        "B1_DESC AS DESCRICAO, " &_
        "B5_CLI, " &_
        "B1_COD, " &_
        "( SELECT TOP 1 B5_ESTRUTU " &_
          "FROM SB5010 AS SB5 " &_
          "WHERE SB5.B5_COD = MIN(G1_COMP) " &_
            "AND SB5.D_E_L_E_T_ = '') AS ESTRUTURA, " &_
        "ISNULL((SELECT TOP 1 C6_PRCVEN " &_
                "FROM SC6010 WITH(NOLOCK) " &_
                "WHERE C6_PRODUTO = B1_COD " &_
                  "AND C6_CLI = '" & Trim(request("cliente")) & "' " &_
                  "AND D_E_L_E_T_ = '' " &_
                "ORDER BY C6_DTVALID DESC),0) AS PRECO, " &_
        "CASE WHEN B1_INPUBE = 'U' THEN 'UBERABA' ELSE 'PIRAPETINGA' END AS B1_INPUBE " &_
      "FROM SB1010 " &_
      "INNER JOIN SG1010 ON G1_COD = B1_COD AND SG1010.D_E_L_E_T_ = '' " &_
      "INNER JOIN SB5010 ON B5_COD = G1_COMP AND SB5010.D_E_L_E_T_ = '' " &_
      "WHERE (SB1010.D_E_L_E_T_ = '') " &_
        "AND SB1010.B1_STATUS IN ('V','A','W') " &_
        "AND SB1010.B1_MSBLQL <> '1' " &_
        "AND SB1010.B1_COD = '" & Trim(request("produto")) & "' " &_
        "AND SUBSTRING(SB1010.B1_COD,1,2) <> '77' " &_
        "AND SB1010.B1_SUBGRUP IN ('750','760') " &_
        iif( (Trim(request("eshop")) = "0"), "AND SB5010.B5_CLI = '" & Trim(request("cliente")) & "' ", "AND (SB5010.B5_CLI IN ('9099', '9990') OR ISNULL(SB5010.B5_CLI,'') = '')" ) &_
      "GROUP BY B1_COD, B1_DESC, B5_CLI, B1_INPUBE " &_
      "ORDER BY SB1010.B1_DESC "

    end if

    %><script>console.log(<%=SQL%>);</script><%

    set PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )
    
    if not PIR_DADOSADV.Eof then

      if Trim(Request("tppedido")) <> "CX" THEN

        response.Write "|"&Trim(Request("tipo"))& "|"&Trim(PIR_DADOSADV("DESCRICAO"))& "|"&Trim(PIR_DADOSADV("ESTRUTURA"))& "|"&Trim(PIR_DADOSADV("B1_INPUBE"))& "|"&Trim(PIR_DADOSADV("B5_LARGCH"))& "|"&Trim(PIR_DADOSADV("B5_COMPCH"))& "|"&Trim(PIR_DADOSADV("B5_VINCO1"))& "|"&Trim(PIR_DADOSADV("B5_VINCO2"))& "|"&Trim(PIR_DADOSADV("B5_VINCO3"))& "|"&Trim(PIR_DADOSADV("B5_VINCO4"))& "|"&Trim(PIR_DADOSADV("B5_VINCO5"))& "|"&Trim(PIR_DADOSADV("B5_TPPALE"))& "|"&Trim(PIR_DADOSADV("B5_TPPALE"))& " - "&Trim(PIR_DADOSADV("B5_DESCPAL"))& "|"&Trim(Request("tppedido"))

      else

        if Trim(PIR_DADOSADV("ESTRUTURA")) <> "" then

          response.Write "|"&Trim(Request("tipo"))& "|"&Trim(PIR_DADOSADV("DESCRICAO"))& "|"&Trim(PIR_DADOSADV("ESTRUTURA"))& "|"&Trim(PIR_DADOSADV("B1_INPUBE"))& "|0|0|0|0|0|0|0|0|0|CX"

        else

          response.Write "ErroNEncontrado|"&Trim(request("tipo"))

        end if

      end if

    else

      response.Write "ErroNEncontrado|"&Trim(request("tipo"))

    end if

    PIR_DADOSADV.Close
    set PIR_DADOSADV = nothing

  else

    response.write "ErroLogof|"&Trim(request("tipo"))

  end if

else

  response.Write "ErroProduto"

end if
%>
