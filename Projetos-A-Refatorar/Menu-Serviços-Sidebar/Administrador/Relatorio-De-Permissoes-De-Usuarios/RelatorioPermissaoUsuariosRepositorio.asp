<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../cnn2/cnn2.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->
<!--#include file="../../includes/variaveis.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046
Response.ContentType = "application/json"

Dim usuarioLogado
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id", Session("CODIGO")
usuarioLogado.Add "codigo", Session("CODUSER")
usuarioLogado.Add "login", Session("USR")
usuarioLogado.Add "nome", Session("NOME")
usuarioLogado.Add "email", Session("EMAIL")
usuarioLogado.Add "nivel", SESSION("NIVEL")

Dim acao
acao = Request("acao")

If acao = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Ação não informada""}")
    Response.End()
End If

Dim resultado
Select Case acao
    Case "listarUsuarios"
        resultado = ListarUsuarios()
    Case "listarSubcategorias"
        resultado = ListarSubcategorias()
    Case "listarAcessos"
        resultado = ListarAcessos()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)


Function ListarUsuarios()
    id = Request("filtroId")
    codPlanta = Request("filtroFilial")
    matricula = Request("filtroMatricula")
    email = Request("filtroEmail")
    idCategoria = Request("filtroCategoria")
    idsSubcategoria = Request("filtroSubcategoria")

    SQL = "SELECT"
    SQL = SQL & "   CRM_R_E_C_N_O as Id, "
    SQL = SQL & "   CRM_LOGIN AS Login, "
    SQL = SQL & "   CRM_NOME AS Nome, "
    SQL = SQL & "   CRM_EMAILPRIMARIO AS Email, "
    SQL = SQL & "   CRM_FUNCAO AS Funcao,"
    SQL = SQL & "   EMP_CODIGO AS CodPlanta, "
    SQL = SQL & "   CRM_COD as Matricula, "
    SQL = SQL & "   CASE WHEN ISDATE(CRM_DATCAD) = 1 "
    SQL = SQL & "       THEN CONVERT(DATE, CRM_DATCAD, 108) "
    SQL = SQL & "   ELSE "
    SQL = SQL & "      NULL "
    SQL = SQL & "   END AS DataCadastro, "

    SQL = SQL & " ("
    SQL = SQL &	"   SELECT DISTINCT"
    SQL = SQL & "   LINKS.LIN_COD as CodSubcategoria, "
    SQL = SQL & "   SUBMENU.SUB_MENU as NomeSubCategoria, "
    SQL = SQL & "   LINKS.LIN_DESC as Descricao "
    SQL = SQL & "   FROM MONTAMENU "
    SQL = SQL & "   INNER JOIN LINKS ON MONTAMENU.LIN_COD = LINKS.LIN_COD  "
    SQL = SQL & "   INNER JOIN SUBMENU ON LINKS.SUB_COD = SUBMENU.SUB_COD "
    SQL = SQL & "   WHERE MONTAMENU.MON_D_E_L_E_T <> '*'  "
    SQL = SQL & "   AND c.CRM_COD = MONTAMENU.CRM_COD "
    SQL = SQL & "   AND LINKS.LIN_D_E_L_E_T <> '*'"
    SQL = SQL & "   AND CRM_D_E_L_E_T = ''"
    SQL = SQL & "   AND DataBloqueio IS NULL"

    If idCategoria = "1" AND idsSubcategoria <> "" Then
        SQL = SQL & "   AND LINKS.SUB_COD IN (" & idsSubcategoria & ")"
    End If

    SQL = SQL & "   ORDER BY LINKS.LIN_COD "
    SQL = SQL & "   FOR JSON PATH"
    SQL = SQL & ") AS servicos, "

    SQL = SQL & "("
    SQL = SQL &	"   SELECT DISTINCT"
    SQL = SQL & "   RELATORIOSCAPA.REL_COD as CodSubcategoria, "
    SQL = SQL & "   CAT_RELATORIOS.CAT_NOME as NomeSubCategoria, "
    SQL = SQL & "   RELATORIOSCAPA.REL_NOME as Descricao "
    SQL = SQL & "   FROM MONTARELAT "
    SQL = SQL & "   INNER JOIN RELATORIOSCAPA ON RELATORIOSCAPA.REL_COD = MONTARELAT.REL_COD "
    SQL = SQL & "   INNER JOIN CAT_RELATORIOS ON RELATORIOSCAPA.CAT_COD = CAT_RELATORIOS.CAT_COD "
    SQL = SQL & "   WHERE MONTARELAT.MON_D_E_L_E_T_E_D <> '*' "
    SQL = SQL & "   AND c.CRM_COD = MONTARELAT.CRM_COD "
    SQL = SQL & "   AND RELATORIOSCAPA.REL_D_E_L_E_T_E_D <> '*'"
    SQL = SQL & "   AND CRM_D_E_L_E_T = ''"
    SQL = SQL & "   AND DataBloqueio IS NULL"

    If idCategoria = "2" AND idsSubcategoria <> "" Then
        SQL = SQL & "   AND RELATORIOSCAPA.CAT_COD IN (" & idsSubcategoria & ")"
    End If

    SQL = SQL & "   ORDER BY RELATORIOSCAPA.REL_COD "
    SQL = SQL & "   FOR JSON PATH"
    SQL = SQL & ") AS relatorios "
    
    SQL = SQL & "FROM CRM_DTI c "
    SQL = SQL & "WHERE 1=1"

    If codPlanta <> "" Then
        SQL = SQL & " AND EMP_CODIGO = '" & codPlanta & "'"
    End If

    If id <> "" Then
        SQL = SQL & " AND CRM_R_E_C_N_O LIKE '%" & id & "%'"
    End If

    If matricula <> "" Then
        SQL = SQL & " AND CRM_COD LIKE '%" & matricula & "%'"
    End If

    If email <> "" Then
        SQL = SQL & " AND CRM_EMAILPRIMARIO LIKE '%" & email & "%'"
    End If

    On Error Resume Next
    SET rsUsuarios = Conecta("PIR-HELPDESK", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsUsuarios) Then
        ListarUsuarios = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de respostas: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsUsuarios.EOF Then
        Do While Not rsUsuarios.EOF
            servicos = "[]"
            relatorios = "[]"
            
            If (idCategoria = "" OR idCategoria = "1") AND NOT IsNull(rsUsuarios("servicos")) Then
                servicos = rsUsuarios("servicos")
            End If

            If (idCategoria = "" OR idCategoria = "2") AND NOT IsNull(rsUsuarios("relatorios")) Then
                relatorios = rsUsuarios("relatorios")
            End If
            
            If servicos <> "[]" OR relatorios <> "[]" Then
                If Not primeiroItem Then jsonArray = jsonArray & "," End If

                jsonArray = jsonArray & "{"
                jsonArray = jsonArray & """id"": """ & EscapeJSON(rsUsuarios("Id")) & """, "
                jsonArray = jsonArray & """login"": """ & EscapeJSON(rsUsuarios("Login")) & """, "
                jsonArray = jsonArray & """funcao"": """ & EscapeJSON(rsUsuarios("Funcao")) & """, "
                jsonArray = jsonArray & """nome"": """ & EscapeJSON(rsUsuarios("Nome")) & """, "
                jsonArray = jsonArray & """email"": """ & EscapeJSON(rsUsuarios("Email")) & """, "
                jsonArray = jsonArray & """codPlanta"": """ & EscapeJSON(rsUsuarios("CodPlanta")) & """, "
                jsonArray = jsonArray & """nomePlanta"": """ & EscapeJSON(ObterNomePlantas("""[" & rsUsuarios("CodPlanta") & """]")) & """, "
                jsonArray = jsonArray & """matricula"": """ & EscapeJSON(rsUsuarios("Matricula")) & """, "
                jsonArray = jsonArray & """dataCadastro"": """ & EscapeJSON(rsUsuarios("DataCadastro")) & """, "
                jsonArray = jsonArray & """servicos"": " & servicos  & ", "
                jsonArray = jsonArray & """relatorios"": " & relatorios  & ""
                jsonArray = jsonArray & "}"
                
                primeiroItem = False
            End If
    
            rsUsuarios.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsUsuarios.Close()
    SET rsUsuarios = NOTHING
    
    ListarUsuarios = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarSubcategorias()
    idCategoria = Request("filtroCategoria")

    If idCategoria = "" OR idCategoria = "1" Then
        SQL = "SELECT "
        SQL = SQL & "SUB_COD AS Id, "
        SQL = SQL & "'1' as IdCategoria, "
        SQL = SQL & "SUB_MENU AS Descricao "
        SQL = SQL & "FROM SUBMENU WHERE SUB_D_E_L_E_T <> '*' "
    End If

    If idCategoria = "" Then
        SQL = SQL & "UNION "
    End If

    If idCategoria = "" OR idCategoria = "2" Then
        SQL = SQL & "SELECT "
        SQL = SQL & "CAT_COD AS Id, "
        SQL = SQL & "'2' as IdCategoria, "
        SQL = SQL & "CAT_NOME AS Descricao "
        SQL = SQL & "FROM CAT_RELATORIOS WHERE CAT_D_E_L_E_T_E_D <> '*'"
    End If

    On Error Resume Next
    SET rsSubcategorias = Conecta("PIR-HELPDESK", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsSubcategorias) Then
        ListarSubcategorias = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de subcategorias: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsSubcategorias.EOF Then
        Do While Not rsSubcategorias.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsSubcategorias("Id")) & """, "
            jsonArray = jsonArray & """idCategoria"": """ & EscapeJSON(rsSubcategorias("IdCategoria")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsSubcategorias("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsSubcategorias.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsSubcategorias.Close()
    SET rsSubcategorias = NOTHING
    
    ListarSubcategorias = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function
%>