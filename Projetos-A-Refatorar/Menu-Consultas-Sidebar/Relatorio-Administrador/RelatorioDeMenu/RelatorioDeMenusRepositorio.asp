<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Response.ContentType = "application/json"
Session.LCID = 1046
%>
<!--#include file="../cnn2/cnn2.asp"-->
<!--#include file="../includes/variaveis.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%
On Error Resume Next
Response.Clear

Function JsonEscape(str)
  If IsNull(str) Or IsEmpty(str) Then JsonEscape = "" : Exit Function
  Dim r : r = CStr(str)
  r = Replace(r, "\", "\\")
  r = Replace(r, """", "\""")
  r = Replace(r, vbCrLf, "\n")
  r = Replace(r, vbCr, "\n")
  r = Replace(r, vbLf, "\n")
  r = Replace(r, vbTab, "\t")
  JsonEscape = Trim(r)
End Function

Function SqlEscape(str)
  If IsNull(str) Or IsEmpty(str) Then SqlEscape = "" : Exit Function
  SqlEscape = Replace(Trim(CStr(str)), "'", "''")
End Function

Dim acao : acao = Trim(Request("acao"))

If acao = "" Then
  Response.Write "{""sucesso"":false,""dados"":null,""mensagem"":""Ação não informada.""}"
  Response.End
End If

Dim RS, SQL

Select Case acao

  Case "listarMenus"
    Dim tipoMenu : tipoMenu = Trim(Request("tipo"))
    Dim jsonMenus : jsonMenus = "{""sucesso"":true,""dados"":["
    Dim primeiroMenu : primeiroMenu = True

    If tipoMenu = "servico" Then
      SQL = "SELECT L.LIN_COD, L.LIN_DESC, S.SUB_MENU " & _
            "FROM LINKS L " & _
            "INNER JOIN SUBMENU S ON LTRIM(RTRIM(L.SUB_COD)) = LTRIM(RTRIM(S.SUB_COD)) " & _
            "WHERE ISNULL(LTRIM(RTRIM(L.LIN_D_E_L_E_T)),'') <> '*' " & _
            "AND ISNULL(LTRIM(RTRIM(S.SUB_D_E_L_E_T)),'') <> '*' " & _
            "ORDER BY S.SUB_MENU, L.LIN_DESC"
    Else
      SQL = "SELECT R.REL_COD, " & _
            "CASE WHEN ISNULL(LTRIM(RTRIM(R.REL_QUERY)),'') <> '' THEN LTRIM(RTRIM(R.REL_QUERY)) " & _
            "ELSE LTRIM(RTRIM(R.REL_NOME)) END AS REL_NOME_FINAL, " & _
            "C.CAT_NOME " & _
            "FROM RELATORIOSCAPA R " & _
            "INNER JOIN CAT_RELATORIOS C ON LTRIM(RTRIM(R.CAT_COD)) = LTRIM(RTRIM(C.CAT_COD)) " & _
            "WHERE ISNULL(LTRIM(RTRIM(R.REL_D_E_L_E_T_E_D)),'') <> '*' " & _
            "AND (ISNULL(LTRIM(RTRIM(R.REL_NOME)),'') <> '' OR ISNULL(LTRIM(RTRIM(R.REL_QUERY)),'') <> '') " & _
            "ORDER BY C.CAT_NOME, REL_NOME_FINAL"
    End If

    On Error Resume Next
    Response.Clear
    Set RS = Conecta("PIR-HELPDESK", SQL)

    If Not IsObject(RS) Then
      Response.Write "{""sucesso"":false,""dados"":null,""mensagem"":""Erro ao conectar ao banco de dados.""}"
      Response.End
    End If

    While Not RS.EOF
      If Not primeiroMenu Then jsonMenus = jsonMenus & ","
      If tipoMenu = "servico" Then
        jsonMenus = jsonMenus & "{""codigo"":""" & JsonEscape(RS("LIN_COD")) & """,""nome"":""" & JsonEscape(RS("LIN_DESC")) & """,""categoria"":""" & JsonEscape(RS("SUB_MENU")) & """}"
      Else
        jsonMenus = jsonMenus & "{""codigo"":""" & JsonEscape(RS("REL_COD")) & """,""nome"":""" & JsonEscape(RS("REL_NOME_FINAL")) & """,""categoria"":""" & JsonEscape(RS("CAT_NOME")) & """}"
      End If
      primeiroMenu = False
      RS.MoveNext()
    Wend
    RS.Close() : Set RS = Nothing

    Response.Write jsonMenus & "],""mensagem"":""""}"

  Case "buscarUsuarios"
    Dim tipoBusca : tipoBusca = Trim(Request("tipo"))
    Dim codigoMenu : codigoMenu = SqlEscape(Trim(Request("codigoMenu")))
    Dim apenasAtivos : apenasAtivos = Trim(Request("apenasAtivos"))
    Dim nomeMenuResult : nomeMenuResult = ""

    If codigoMenu = "" Then
      Response.Write "{""sucesso"":false,""dados"":null,""mensagem"":""Código do menu não informado.""}"
      Response.End
    End If

    Dim filtroAtivo : filtroAtivo = ""
    If apenasAtivos = "1" Then
      filtroAtivo = "AND U.CRM_D_E_L_E_T <> '*' AND ISNULL(U.DataBloqueio,'') = '' "
    End If

    ' Apenas registros onde a filial do usuário coincide com a filial da permissão
    If tipoBusca = "servico" Then
      SQL = "SELECT U.CRM_COD, U.CRM_NOME, U.CRM_LOGIN, U.EMP_CODIGO, " & _
            "U.CRM_D_E_L_E_T, U.DataBloqueio, " & _
            "L.LIN_DESC AS NOME_MENU, M.FILIAL " & _
            "FROM MONTAMENU M " & _
            "INNER JOIN CRM_DTI U ON LTRIM(RTRIM(M.CRM_COD)) = LTRIM(RTRIM(U.CRM_COD)) " & _
            "INNER JOIN LINKS L ON LTRIM(RTRIM(M.LIN_COD)) = LTRIM(RTRIM(L.LIN_COD)) " & _
            "WHERE LTRIM(RTRIM(M.LIN_COD)) = '" & codigoMenu & "' " & _
            "AND ISNULL(LTRIM(RTRIM(M.MON_D_E_L_E_T)),'') = '' " & _
            "AND M.CRM_COD <> '' " & _
            "AND LTRIM(RTRIM(U.EMP_CODIGO)) = LTRIM(RTRIM(M.FILIAL)) " & _
            filtroAtivo & _
            "ORDER BY U.CRM_NOME"
    Else
      SQL = "SELECT U.CRM_COD, U.CRM_NOME, U.CRM_LOGIN, U.EMP_CODIGO, " & _
            "U.CRM_D_E_L_E_T, U.DataBloqueio, " & _
            "CASE WHEN ISNULL(LTRIM(RTRIM(R.REL_QUERY)),'') <> '' THEN LTRIM(RTRIM(R.REL_QUERY)) " & _
            "ELSE LTRIM(RTRIM(R.REL_NOME)) END AS NOME_MENU, M.FILIAL " & _
            "FROM MONTARELAT M " & _
            "INNER JOIN CRM_DTI U ON LTRIM(RTRIM(M.CRM_COD)) = LTRIM(RTRIM(U.CRM_COD)) " & _
            "INNER JOIN RELATORIOSCAPA R ON LTRIM(RTRIM(M.REL_COD)) = LTRIM(RTRIM(R.REL_COD)) " & _
            "WHERE LTRIM(RTRIM(M.REL_COD)) = '" & codigoMenu & "' " & _
            "AND ISNULL(LTRIM(RTRIM(M.MON_D_E_L_E_T_E_D)),'') = '' " & _
            "AND M.CRM_COD <> '' " & _
            "AND LTRIM(RTRIM(U.EMP_CODIGO)) = LTRIM(RTRIM(M.FILIAL)) " & _
            filtroAtivo & _
            "ORDER BY U.CRM_NOME"
    End If

    On Error Resume Next
    Response.Clear
    Set RS = Conecta("PIR-HELPDESK", SQL)

    If Not IsObject(RS) Then
      Response.Write "{""sucesso"":false,""dados"":null,""mensagem"":""Erro ao conectar ao banco de dados.""}"
      Response.End
    End If

    Dim jsonUsuarios : jsonUsuarios = "{""sucesso"":true,""dados"":{""usuarios"":["
    Dim primeiroUsr : primeiroUsr = True
    Dim totalUsr : totalUsr = 0

    While Not RS.EOF
      If nomeMenuResult = "" Then nomeMenuResult = Trim(RS("NOME_MENU") & "")
      If Not primeiroUsr Then jsonUsuarios = jsonUsuarios & ","
      Dim isAtivo : isAtivo = "0"
      If LTRIM(RTRIM(RS("CRM_D_E_L_E_T") & "")) <> "*" And LTRIM(RTRIM(RS("DataBloqueio") & "")) = "" Then isAtivo = "1"
      jsonUsuarios = jsonUsuarios & "{" & _
        """matricula"":""" & JsonEscape(RS("CRM_COD")) & """," & _
        """nome"":""" & JsonEscape(RS("CRM_NOME")) & """," & _
        """login"":""" & JsonEscape(RS("CRM_LOGIN")) & """," & _
        """filial"":""" & JsonEscape(RS("EMP_CODIGO")) & """," & _
        """filialPermissao"":""" & JsonEscape(RS("FILIAL")) & """," & _
        """ativo"":""" & isAtivo & """}"
      primeiroUsr = False
      totalUsr = totalUsr + 1
      RS.MoveNext()
    Wend
    RS.Close() : Set RS = Nothing

    jsonUsuarios = jsonUsuarios & "]," & _
      """nomeMenu"":""" & JsonEscape(nomeMenuResult) & """," & _
      """tipoMenu"":""" & JsonEscape(tipoBusca) & """," & _
      """totalRegistros"":" & totalUsr & "}," & _
      """mensagem"":""""}"

    Response.Write jsonUsuarios

  Case Else
    Response.Write "{""sucesso"":false,""dados"":null,""mensagem"":""Ação inválida.""}"

End Select

Response.End
%>
