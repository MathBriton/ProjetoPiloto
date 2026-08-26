<!--#include file="../cnn/cnn.asp" -->
<!--#Include File="../includes/variaveis.asp"-->
<%
if trim(Request("Id")) <> "" then
  SQL = "SELECT " & _
          "CRM_COD, " & _
          "CASE CONVERT(VARCHAR(10), " & _
          "CONVERT(smalldatetime,CRM_DATACES),103) WHEN '01/01/1900' THEN '' ELSE CONVERT(VARCHAR(10), CONVERT(smalldatetime,CRM_DATACES),103) END ULTIMOACESSO, " & _
          "HP_ADMHP, " & _
          "CRM_NOME, " & _
          "RA_SEXO, " & _
          "RA_NASC, " & _
          "CRM_LOGIN, " & _
          "CRM_DATACES, " & _
          "CASE ISNULL(CRM_D_E_L_E_T,'') WHEN '*' THEN 'EXCLUÍDO' ELSE 'ATIVO' END SITUACAO, " & _
          "CRM_SENHA, " & _
          "CRM_EMAILPRIMARIO, " & _
          "CRM_RAMAL, " & _
          "EMP_CODIGO, " & _
          "CRM_TIPO, " & _
          "CAST(ISNULL(Bloqueado,0) AS INT) Bloqueado " & _
        "FROM CRM_DTI " & _
        "WHERE CRM_R_E_C_N_O = " & Trim(Request("Id")) & " "
  Dados21.Open(SQL)
  if Not Dados21.Eof Then
    Response.write "0|0|"&Dados21("CRM_COD")&"|"&Dados21("CRM_NOME")&"|"&Dados21("EMP_CODIGO")&"|"
  else
    Response.write "1|0|"
  end if
  Dados21.Close()
else
  Response.write "1|0|"
end if
%>