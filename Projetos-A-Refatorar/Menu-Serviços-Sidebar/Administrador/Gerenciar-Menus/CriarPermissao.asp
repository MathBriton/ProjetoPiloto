<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer  = True
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Session.LCID = 1046
%>
<!--#include file="../cnn2/cnn2.asp"-->
<!--#include file="../includes/variaveis.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%
Dim Conexao, Dados
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open BuscaStringConexao("PIR-HELPDESK")
Set Dados = Server.CreateObject("ADODB.Recordset")
Dados.ActiveConnection = Conexao

Call ValidaPermissaoFormularios("../CriarCategoriaPermissao/CriarPermissao", Session("CODUSER"), Session("FILIAL"))

If Not Dados Is Nothing Then
    If Dados.State <> 0 Then Dados.Close
End If
If Not Conexao Is Nothing Then
    If Conexao.State <> 0 Then Conexao.Close
End If
Set Dados = Nothing
Set Conexao = Nothing

Function JsonEscape(str)
    If IsNull(str) Or IsEmpty(str) Then JsonEscape = "" : Exit Function
    Dim r : r = CStr(str)
    r = Replace(r, "\", "\\")
    r = Replace(r, """", "\""")
    r = Replace(r, vbCrLf, "\n")
    r = Replace(r, vbCr,   "\n")
    r = Replace(r, vbLf,   "\n")
    r = Replace(r, vbTab,  "\t")
    JsonEscape = Trim(r)
End Function

Function SqlEscape(str)
    If IsNull(str) Or IsEmpty(str) Then SqlEscape = "" : Exit Function
    SqlEscape = Replace(Trim(CStr(str)), "'", "''")
End Function

Function ExecutaEscrita(strConexao, sql)
    Dim cn : Set cn = Server.CreateObject("ADODB.Connection")
    cn.Open BuscaStringConexao(UCase(strConexao))
    On Error Resume Next
    Err.Clear
    cn.Execute sql
    Dim numErr : Dim descErr
    numErr  = Err.Number
    descErr = Err.Description
    On Error GoTo 0
    cn.Close : Set cn = Nothing
    If numErr <> 0 Then ExecutaEscrita = descErr Else ExecutaEscrita = ""
End Function

Dim acao : acao = Trim(Request("acao"))

If acao <> "" Then
    Response.ContentType = "application/json"
    On Error Resume Next

    Select Case acao

    Case "listarTudo"

        Dim json : json = "{""success"":true,"
        json = json & """servicos"":{""categorias"":["

        Dim arrServCod(), arrServNome()
        Dim contServ : contServ = 0
        Dim SQL

        SQL = "SELECT s.SUB_COD, s.SUB_MENU " & _
              "FROM SUBMENU s " & _
              "WHERE ISNULL(LTRIM(RTRIM(s.SUB_D_E_L_E_T)), '') <> '*' " & _
              "AND EXISTS ( " & _
              "   SELECT 1 FROM LINKS l " & _
              "   WHERE LTRIM(RTRIM(l.SUB_COD)) = LTRIM(RTRIM(s.SUB_COD)) " & _
              "   AND ISNULL(LTRIM(RTRIM(l.LIN_D_E_L_E_T)), '') <> '*' " & _
              ") " & _
              "ORDER BY s.SUB_MENU"

        Dim RS : Set RS = Conecta("PIR-HELPDESK", SQL)
        While Not RS.EOF
            contServ = contServ + 1
            ReDim Preserve arrServCod(contServ)
            ReDim Preserve arrServNome(contServ)
            arrServCod(contServ)  = Trim(RS("SUB_COD")  & "")
            arrServNome(contServ) = Trim(RS("SUB_MENU") & "")
            RS.MoveNext()
        Wend
        If IsObject(RS) Then RS.Close() : Set RS = Nothing

        Dim iServ
        For iServ = 1 To contServ
            If iServ > 1 Then json = json & ","
            json = json & "{""codigo"":""" & JsonEscape(arrServCod(iServ)) & ""","
            json = json & """nome"":"""   & JsonEscape(arrServNome(iServ)) & ""","
            json = json & """permissoes"":["

            SQL = "SELECT LIN_COD, LIN_DESC, LIN_URL, LIN_EXTEN " & _
                  "FROM LINKS " & _
                  "WHERE LTRIM(RTRIM(SUB_COD)) = '" & SqlEscape(arrServCod(iServ)) & "' " & _
                  "AND ISNULL(LTRIM(RTRIM(LIN_D_E_L_E_T)), '') <> '*' " & _
                  "ORDER BY LIN_DESC"

            Set RS = Conecta("PIR-HELPDESK", SQL)
            Dim primPerm : primPerm = True
            While Not RS.EOF
                If Not primPerm Then json = json & ","
                json = json & "{""codigo"":"""   & JsonEscape(RS("LIN_COD"))  & ""","
                json = json & """nome"":"""      & JsonEscape(RS("LIN_DESC")) & ""","
                json = json & """url"":"""       & JsonEscape(RS("LIN_URL"))  & ""","
                json = json & """extensao"":""" & JsonEscape(RS("LIN_EXTEN")) & """}"
                primPerm = False
                RS.MoveNext()
            Wend
            If IsObject(RS) Then RS.Close() : Set RS = Nothing
            json = json & "]}"
        Next

        json = json & "]},"
        json = json & """consultas"":{""categorias"":["

        Dim arrConsCod(), arrConsNome()
        Dim contCons : contCons = 0

        SQL = "SELECT c.CAT_COD, c.CAT_NOME " & _
              "FROM CAT_RELATORIOS c " & _
              "WHERE EXISTS ( " & _
              "   SELECT 1 FROM RELATORIOSCAPA r " & _
              "   WHERE LTRIM(RTRIM(r.CAT_COD)) = LTRIM(RTRIM(c.CAT_COD)) " & _
              "   AND ISNULL(LTRIM(RTRIM(r.REL_D_E_L_E_T_E_D)), '') <> '*' " & _
              "   AND (ISNULL(LTRIM(RTRIM(r.REL_NOME)), '') <> '' OR ISNULL(LTRIM(RTRIM(r.REL_QUERY)), '') <> '') " & _
              ") " & _
              "ORDER BY c.CAT_NOME"

        Set RS = Conecta("PIR-HELPDESK", SQL)
        While Not RS.EOF
            contCons = contCons + 1
            ReDim Preserve arrConsCod(contCons)
            ReDim Preserve arrConsNome(contCons)
            arrConsCod(contCons)  = Trim(RS("CAT_COD")  & "")
            arrConsNome(contCons) = Trim(RS("CAT_NOME") & "")
            RS.MoveNext()
        Wend
        If IsObject(RS) Then RS.Close() : Set RS = Nothing

        Dim iCons
        For iCons = 1 To contCons
            If iCons > 1 Then json = json & ","
            json = json & "{""codigo"":""" & JsonEscape(arrConsCod(iCons)) & ""","
            json = json & """nome"":"""   & JsonEscape(arrConsNome(iCons)) & ""","
            json = json & """permissoes"":["

            SQL = "SELECT REL_COD, REL_NOME, REL_QUERY, REL_URL, REL_EXTEN, REL_BANCO " & _
                  "FROM RELATORIOSCAPA " & _
                  "WHERE LTRIM(RTRIM(CAT_COD)) = '" & SqlEscape(arrConsCod(iCons)) & "' " & _
                  "AND ISNULL(LTRIM(RTRIM(REL_D_E_L_E_T_E_D)), '') <> '*' " & _
                  "AND (ISNULL(LTRIM(RTRIM(REL_NOME)), '') <> '' OR ISNULL(LTRIM(RTRIM(REL_QUERY)), '') <> '') " & _
                  "ORDER BY REL_NOME"

            Set RS = Conecta("PIR-HELPDESK", SQL)
            Dim primPermC : primPermC = True
            While Not RS.EOF
                Dim nomeRel : nomeRel = Trim(RS("REL_NOME") & "")
                If nomeRel = "" Then nomeRel = Trim(RS("REL_QUERY") & "")
                If Not primPermC Then json = json & ","
                json = json & "{""codigo"":"""   & JsonEscape(RS("REL_COD"))  & ""","
                json = json & """nome"":"""      & JsonEscape(nomeRel)        & ""","
                json = json & """url"":"""       & JsonEscape(RS("REL_URL"))  & ""","
                json = json & """extensao"":""" & JsonEscape(RS("REL_EXTEN")) & ""","
                json = json & """banco"":"""    & JsonEscape(RS("REL_BANCO")) & """}"
                primPermC = False
                RS.MoveNext()
            Wend
            If IsObject(RS) Then RS.Close() : Set RS = Nothing
            json = json & "]}"
        Next

        json = json & "]}}"
        Response.Write json

    Case "criarPermissao"

        Dim tipoCriarPerm, categoriaCriarPerm, nomeCriarPerm, urlCriarPerm, bancoCriarPerm
        tipoCriarPerm      = Trim(Request("tipo"))
        categoriaCriarPerm = Trim(Request("codigoCategoria"))
        nomeCriarPerm      = Left(Trim(Request("nome")), 50)
        urlCriarPerm       = Trim(Request("url"))
        bancoCriarPerm     = Trim(Request("banco"))

        If nomeCriarPerm = "" Then
            Response.Write "{""success"":false,""mensagem"":""Nome da permiss" & Chr(227) & "o n" & Chr(227) & "o pode ser vazio.""}"
            Response.End
        End If

        If tipoCriarPerm = "servico" Then
            If urlCriarPerm = "" Then
                Response.Write "{""success"":false,""mensagem"":""Informe a URL (LIN_URL) para Servi" & Chr(231) & "os.""}"
                Response.End
            End If
        Else
            If urlCriarPerm = "" Then
                Response.Write "{""success"":false,""mensagem"":""Informe a URL (REL_URL) para Consultas.""}"
                Response.End
            End If
            If bancoCriarPerm = "" Then
                Response.Write "{""success"":false,""mensagem"":""Informe o Banco (REL_BANCO) para Consultas.""}"
                Response.End
            End If
        End If

        Dim novoCodigoPerm, sqlCriar, erroCriar

        If tipoCriarPerm = "servico" Then
            SQL = "SELECT ISNULL(MAX(CAST(LIN_COD AS INT)), 0) + 1 AS PROXIMO " & _
                  "FROM LINKS WHERE ISNUMERIC(LIN_COD) = 1"
            Set RS = Conecta("PIR-HELPDESK", SQL)
            If Not RS.EOF Then novoCodigoPerm = CStr(RS("PROXIMO")) Else novoCodigoPerm = "1"
            If IsObject(RS) Then RS.Close() : Set RS = Nothing

            sqlCriar = "INSERT INTO LINKS (LIN_COD, SUB_COD, LIN_DESC, SRV_CODSERV, LIN_NOME, LIN_ORDEM, LIN_URL, LIN_EXTEN, LIN_D_E_L_E_T) " & _
                       "VALUES ('" & SqlEscape(novoCodigoPerm) & "', " & _
                       "'" & SqlEscape(categoriaCriarPerm) & "', " & _
                       "'" & SqlEscape(nomeCriarPerm) & "', " & _
                       "'1', '" & SqlEscape(nomeCriarPerm) & "', " & _
                       "'20', '" & SqlEscape(urlCriarPerm) & "', '.asp', '')"
        Else
            SQL = "SELECT ISNULL(MAX(CAST(REL_COD AS INT)), 0) + 1 AS PROXIMO " & _
                  "FROM RELATORIOSCAPA WHERE ISNUMERIC(REL_COD) = 1"
            Set RS = Conecta("PIR-HELPDESK", SQL)
            If Not RS.EOF Then novoCodigoPerm = CStr(RS("PROXIMO")) Else novoCodigoPerm = "1"
            If IsObject(RS) Then RS.Close() : Set RS = Nothing

            sqlCriar = "INSERT INTO RELATORIOSCAPA (REL_COD, CAT_COD, REL_NOME, REL_QUERY, REL_URL, REL_EXTEN, REL_BANCO, REL_D_E_L_E_T_E_D, REL_ATALHO) " & _
                       "VALUES ('" & SqlEscape(novoCodigoPerm) & "', " & _
                       "'" & SqlEscape(categoriaCriarPerm) & "', " & _
                       "'" & SqlEscape(nomeCriarPerm) & "', " & _
                       "'" & SqlEscape(nomeCriarPerm) & "', " & _
                       "'" & SqlEscape(urlCriarPerm) & "', '.asp', '" & SqlEscape(bancoCriarPerm) & "', '', 'N')"
        End If

        erroCriar = ExecutaEscrita("PIR-HELPDESK", sqlCriar)

        If erroCriar = "" Then
            Response.Write "{""success"":true,""mensagem"":""Permiss" & Chr(227) & "o criada com sucesso."",""codigo"":""" & novoCodigoPerm & """}"
        Else
            Response.Write "{""success"":false,""mensagem"":""Erro ao criar permiss" & Chr(227) & "o: " & JsonEscape(erroCriar) & """}"
        End If

    Case "editarPermissao"

        Dim tipoEditPerm, codigoEditPerm, novoNomePerm, urlEditPerm, bancoEditPerm, codigoCategoriaEdit
        Dim sqlFinal, erroEdit

        tipoEditPerm        = Trim(Request("tipo"))
        codigoEditPerm      = Trim(Request("codigo"))
        novoNomePerm        = Left(Trim(Request("nome")), 50)
        urlEditPerm         = Trim(Request("url"))
        bancoEditPerm       = Trim(Request("banco"))
        codigoCategoriaEdit = Trim(Request("codigoCategoria"))

        If tipoEditPerm <> "servico" And tipoEditPerm <> "consulta" Then
            Response.Write "{""success"":false,""mensagem"":""Tipo inv" & Chr(225) & "lido. Use servico ou consulta.""}"
            Response.End
        End If
        If novoNomePerm = "" Then
            Response.Write "{""success"":false,""mensagem"":""Nome da permiss" & Chr(227) & "o n" & Chr(227) & "o pode ser vazio.""}"
            Response.End
        End If
        If codigoEditPerm = "" Then
            Response.Write "{""success"":false,""mensagem"":""C" & Chr(243) & "digo da permiss" & Chr(227) & "o n" & Chr(227) & "o informado.""}"
            Response.End
        End If
        If codigoCategoriaEdit = "" Then
            Response.Write "{""success"":false,""mensagem"":""Categoria destino n" & Chr(227) & "o informada.""}"
            Response.End
        End If

        If tipoEditPerm = "servico" Then
            If urlEditPerm = "" Then
                Response.Write "{""success"":false,""mensagem"":""Informe a URL (LIN_URL) para Servi" & Chr(231) & "os.""}"
                Response.End
            End If
        Else
            If urlEditPerm = "" Then
                Response.Write "{""success"":false,""mensagem"":""Informe a URL (REL_URL) para Consultas.""}"
                Response.End
            End If
            If bancoEditPerm = "" Then
                Response.Write "{""success"":false,""mensagem"":""Informe o Banco (REL_BANCO) para Consultas.""}"
                Response.End
            End If
        End If

        If tipoEditPerm = "servico" Then
            SQL = "SELECT TOP 1 1 AS OK FROM SUBMENU s " & _
                  "WHERE LTRIM(RTRIM(s.SUB_COD)) = '" & SqlEscape(codigoCategoriaEdit) & "' " & _
                  "AND ISNULL(LTRIM(RTRIM(s.SUB_D_E_L_E_T)), '') <> '*'"
            Set RS = Conecta("PIR-HELPDESK", SQL)
            If RS.EOF Then
                If IsObject(RS) Then RS.Close() : Set RS = Nothing
                Response.Write "{""success"":false,""mensagem"":""Categoria de Servi" & Chr(231) & "o inv" & Chr(225) & "lida ou inativa.""}"
                Response.End
            End If
            If IsObject(RS) Then RS.Close() : Set RS = Nothing

            sqlFinal = "UPDATE LINKS SET " & _
                       "SUB_COD = '" & SqlEscape(codigoCategoriaEdit) & "', " & _
                       "LIN_DESC = '" & SqlEscape(novoNomePerm) & "', " & _
                       "LIN_NOME = '" & SqlEscape(novoNomePerm) & "', " & _
                       "LIN_URL = '"  & SqlEscape(urlEditPerm)   & "' " & _
                       "WHERE LIN_COD = '" & SqlEscape(codigoEditPerm) & "'"
        Else
            SQL = "SELECT TOP 1 1 AS OK FROM CAT_RELATORIOS c " & _
                  "WHERE LTRIM(RTRIM(c.CAT_COD)) = '" & SqlEscape(codigoCategoriaEdit) & "'"
            Set RS = Conecta("PIR-HELPDESK", SQL)
            If RS.EOF Then
                If IsObject(RS) Then RS.Close() : Set RS = Nothing
                Response.Write "{""success"":false,""mensagem"":""Categoria de Consulta inv" & Chr(225) & "lida.""}"
                Response.End
            End If
            If IsObject(RS) Then RS.Close() : Set RS = Nothing

            sqlFinal = "UPDATE RELATORIOSCAPA SET " & _
                       "CAT_COD = '"   & SqlEscape(codigoCategoriaEdit) & "', " & _
                       "REL_QUERY = '" & SqlEscape(novoNomePerm)        & "', " & _
                       "REL_NOME = '"  & SqlEscape(novoNomePerm)        & "', " & _
                       "REL_URL = '"   & SqlEscape(urlEditPerm)          & "', " & _
                       "REL_BANCO = '" & SqlEscape(bancoEditPerm)        & "' " & _
                       "WHERE REL_COD = '" & SqlEscape(codigoEditPerm) & "'"
        End If

        If Trim(sqlFinal) = "" Then
            Response.Write "{""success"":false,""mensagem"":""SQL vazio (n" & Chr(227) & "o foi montado).""}"
            Response.End
        End If

        erroEdit = ExecutaEscrita("PIR-HELPDESK", sqlFinal)

        If erroEdit = "" Then
            Response.Write "{""success"":true,""mensagem"":""Permiss" & Chr(227) & "o atualizada com sucesso.""}"
        Else
            Response.Write "{""success"":false,""mensagem"":""Erro ao atualizar permiss" & Chr(227) & "o: " & JsonEscape(erroEdit) & """}"
        End If

    Case "excluirPermissao"

        Dim tipoExclPerm, codigoExclPerm, sqlExcl, erroExcl
        tipoExclPerm   = Trim(Request("tipo"))
        codigoExclPerm = Trim(Request("codigo"))

        If tipoExclPerm = "servico" Then
            sqlExcl = "UPDATE LINKS SET LIN_D_E_L_E_T = '*' WHERE LIN_COD = '" & SqlEscape(codigoExclPerm) & "'"
        Else
            sqlExcl = "UPDATE RELATORIOSCAPA SET REL_D_E_L_E_T_E_D = '*' WHERE REL_COD = '" & SqlEscape(codigoExclPerm) & "'"
        End If

        erroExcl = ExecutaEscrita("PIR-HELPDESK", sqlExcl)

        If erroExcl = "" Then
            Response.Write "{""success"":true,""mensagem"":""Permiss" & Chr(227) & "o exclu" & Chr(237) & "da com sucesso.""}"
        Else
            Response.Write "{""success"":false,""mensagem"":""Erro ao excluir permiss" & Chr(227) & "o: " & JsonEscape(erroExcl) & """}"
        End If

    Case "listarInativos"

        Dim escopoInativo : escopoInativo = Trim(Request("escopo"))
        Dim jsonInativo : jsonInativo = "{""success"":true,"

        jsonInativo = jsonInativo & """servicos"":{""categorias"":["
        If escopoInativo = "todos" Or escopoInativo = "servicos" Then
            Dim arrServCodI(), arrServNomeI()
            Dim contServI : contServI = 0
            SQL = "SELECT s.SUB_COD, s.SUB_MENU " & _
                  "FROM SUBMENU s " & _
                  "WHERE ISNULL(LTRIM(RTRIM(s.SUB_D_E_L_E_T)), '') <> '*' " & _
                  "AND EXISTS ( " & _
                  "   SELECT 1 FROM LINKS l " & _
                  "   WHERE LTRIM(RTRIM(l.SUB_COD)) = LTRIM(RTRIM(s.SUB_COD)) " & _
                  "   AND LTRIM(RTRIM(ISNULL(l.LIN_D_E_L_E_T, ''))) = '*' " & _
                  ") " & _
                  "ORDER BY s.SUB_MENU"
            Set RS = Conecta("PIR-HELPDESK", SQL)
            While Not RS.EOF
                contServI = contServI + 1
                ReDim Preserve arrServCodI(contServI)
                ReDim Preserve arrServNomeI(contServI)
                arrServCodI(contServI)  = Trim(RS("SUB_COD")  & "")
                arrServNomeI(contServI) = Trim(RS("SUB_MENU") & "")
                RS.MoveNext()
            Wend
            If IsObject(RS) Then RS.Close() : Set RS = Nothing
            Dim iServI
            For iServI = 1 To contServI
                If iServI > 1 Then jsonInativo = jsonInativo & ","
                jsonInativo = jsonInativo & "{""codigo"":""" & JsonEscape(arrServCodI(iServI)) & ""","
                jsonInativo = jsonInativo & """nome"":"""   & JsonEscape(arrServNomeI(iServI)) & ""","
                jsonInativo = jsonInativo & """permissoes"":["
                SQL = "SELECT LIN_COD, LIN_DESC, LIN_URL, LIN_EXTEN " & _
                      "FROM LINKS " & _
                      "WHERE LTRIM(RTRIM(SUB_COD)) = '" & SqlEscape(arrServCodI(iServI)) & "' " & _
                      "AND LTRIM(RTRIM(ISNULL(LIN_D_E_L_E_T, ''))) = '*' " & _
                      "ORDER BY LIN_DESC"
                Set RS = Conecta("PIR-HELPDESK", SQL)
                Dim primPermI : primPermI = True
                While Not RS.EOF
                    If Not primPermI Then jsonInativo = jsonInativo & ","
                    jsonInativo = jsonInativo & "{""codigo"":"""   & JsonEscape(RS("LIN_COD"))  & ""","
                    jsonInativo = jsonInativo & """nome"":"""      & JsonEscape(RS("LIN_DESC")) & ""","
                    jsonInativo = jsonInativo & """url"":"""       & JsonEscape(RS("LIN_URL"))  & ""","
                    jsonInativo = jsonInativo & """extensao"":""" & JsonEscape(RS("LIN_EXTEN")) & """}"
                    primPermI = False
                    RS.MoveNext()
                Wend
                If IsObject(RS) Then RS.Close() : Set RS = Nothing
                jsonInativo = jsonInativo & "]}"
            Next
        End If
        jsonInativo = jsonInativo & "]},"

        jsonInativo = jsonInativo & """consultas"":{""categorias"":["
        If escopoInativo = "todos" Or escopoInativo = "consultas" Then
            Dim arrConsCodI(), arrConsNomeI()
            Dim contConsI : contConsI = 0
            SQL = "SELECT c.CAT_COD, c.CAT_NOME " & _
                  "FROM CAT_RELATORIOS c " & _
                  "WHERE EXISTS ( " & _
                  "   SELECT 1 FROM RELATORIOSCAPA r " & _
                  "   WHERE LTRIM(RTRIM(r.CAT_COD)) = LTRIM(RTRIM(c.CAT_COD)) " & _
                  "   AND LTRIM(RTRIM(ISNULL(r.REL_D_E_L_E_T_E_D, ''))) = '*' " & _
                  ") " & _
                  "ORDER BY c.CAT_NOME"
            Set RS = Conecta("PIR-HELPDESK", SQL)
            While Not RS.EOF
                contConsI = contConsI + 1
                ReDim Preserve arrConsCodI(contConsI)
                ReDim Preserve arrConsNomeI(contConsI)
                arrConsCodI(contConsI)  = Trim(RS("CAT_COD")  & "")
                arrConsNomeI(contConsI) = Trim(RS("CAT_NOME") & "")
                RS.MoveNext()
            Wend
            If IsObject(RS) Then RS.Close() : Set RS = Nothing
            Dim iConsI
            For iConsI = 1 To contConsI
                If iConsI > 1 Then jsonInativo = jsonInativo & ","
                jsonInativo = jsonInativo & "{""codigo"":""" & JsonEscape(arrConsCodI(iConsI)) & ""","
                jsonInativo = jsonInativo & """nome"":"""   & JsonEscape(arrConsNomeI(iConsI)) & ""","
                jsonInativo = jsonInativo & """permissoes"":["
                SQL = "SELECT REL_COD, REL_NOME, REL_QUERY, REL_URL, REL_EXTEN, REL_BANCO " & _
                      "FROM RELATORIOSCAPA " & _
                      "WHERE LTRIM(RTRIM(CAT_COD)) = '" & SqlEscape(arrConsCodI(iConsI)) & "' " & _
                      "AND LTRIM(RTRIM(ISNULL(REL_D_E_L_E_T_E_D, ''))) = '*' " & _
                      "ORDER BY REL_NOME"
                Set RS = Conecta("PIR-HELPDESK", SQL)
                Dim primPermCI : primPermCI = True
                While Not RS.EOF
                    Dim nomeRelI : nomeRelI = Trim(RS("REL_NOME") & "")
                    If nomeRelI = "" Then nomeRelI = Trim(RS("REL_QUERY") & "")
                    If Not primPermCI Then jsonInativo = jsonInativo & ","
                    jsonInativo = jsonInativo & "{""codigo"":"""   & JsonEscape(RS("REL_COD"))   & ""","
                    jsonInativo = jsonInativo & """nome"":"""      & JsonEscape(nomeRelI)         & ""","
                    jsonInativo = jsonInativo & """url"":"""       & JsonEscape(RS("REL_URL"))   & ""","
                    jsonInativo = jsonInativo & """extensao"":""" & JsonEscape(RS("REL_EXTEN"))  & ""","
                    jsonInativo = jsonInativo & """banco"":"""    & JsonEscape(RS("REL_BANCO"))  & """}"
                    primPermCI = False
                    RS.MoveNext()
                Wend
                If IsObject(RS) Then RS.Close() : Set RS = Nothing
                jsonInativo = jsonInativo & "]}"
            Next
        End If
        jsonInativo = jsonInativo & "]}}"
        Response.Write jsonInativo

    Case "reativarPermissao"

        Dim tipoReat, codigoReat, sqlReat, erroReat
        tipoReat   = Trim(Request("tipo"))
        codigoReat = Trim(Request("codigo"))

        If tipoReat = "servico" Then
            sqlReat = "UPDATE LINKS SET LIN_D_E_L_E_T = '' WHERE LIN_COD = '" & SqlEscape(codigoReat) & "'"
        Else
            sqlReat = "UPDATE RELATORIOSCAPA SET REL_D_E_L_E_T_E_D = '' WHERE REL_COD = '" & SqlEscape(codigoReat) & "'"
        End If

        erroReat = ExecutaEscrita("PIR-HELPDESK", sqlReat)

        If erroReat = "" Then
            Response.Write "{""success"":true,""mensagem"":""Menu reativado com sucesso.""}"
        Else
            Response.Write "{""success"":false,""mensagem"":""Erro ao reativar: " & JsonEscape(erroReat) & """}"
        End If

    Case Else
        Response.Write "{""success"":false,""mensagem"":""A" & Chr(231) & Chr(227) & "o inv" & Chr(225) & "lida.""}"

    End Select

    If Err.Number <> 0 Then
        Response.Write "{""success"":false,""mensagem"":""Erro interno: " & JsonEscape(Err.Description) & """}"
        Err.Clear
    End If

    Response.End
End If
%>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SWnet - Gerenciar Menus</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
  <style>
    body {
      background: #f0f0f0;
      color: #212529;
      min-height: 100vh;
    }

    .page-title {
      font-size: 1.25rem;
      color: #4a4a4a;
      margin: 0;
    }

    .form-label {
      font-size: 0.82rem;
      font-weight: 600;
      margin-bottom: 0.25rem;
      color: #212529;
    }
    .form-control,
    .form-select {
      font-size: 0.85rem;
    }
    .form-control:focus,
    .form-select:focus {
      border-color: #4a4a4a;
      box-shadow: 0 0 0 0.2rem rgba(74, 74, 74, 0.15);
    }
    .form-control[readonly] {
      background-color: #f8f9fa;
    }

    .action-bar {
      background: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 1px 3px rgba(0, 0, 0, .08);
    }

    .panel {
      background: #fff;
      border: 1px solid #dee2e6;
      border-radius: 6px;
    }
    .panel-header {
      background: #fff;
      color: #212529;
      border-bottom: 1px solid #dee2e6;
      padding: 0.6rem 1rem;
      border-radius: 6px 6px 0 0;
      font-size: 0.8rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.3px;
      display: flex;
      align-items: center;
      gap: 0.4rem;
    }
    .panel-body {
      padding: 1rem;
    }

    .section-header {
      background: #fff;
      color: #212529;
      padding: .75rem .875rem;
      border-bottom: 1px solid #dee2e6;
      border-radius: .5rem .5rem 0 0;
      user-select: none;
    }
    .section-header .section-title {
      font-size: .8125rem;
      font-weight: 600;
      color: #212529;
      margin: 0;
    }
    .section-header i { color: #212529; }
    .section-header .badge {
      font-size: .5625rem;
      background: #e9ecef !important;
      color: #495057 !important;
    }
    .section-header[role="button"] { cursor: pointer; }

    .section-toggle-icon {
      transition: transform .25s ease;
      color: #212529;
      font-size: .75rem;
    }
    .section-toggle-icon.collapsed { transform: rotate(-90deg); }

    .section-body {
      border: 1px solid #dee2e6;
      border-top: none;
      border-radius: 0 0 .5rem .5rem;
      background: #f0f0f0;
    }

    .scrollable-container {
      max-height: calc(100vh - 200px);
      overflow-y: auto;
      padding: .75rem;
    }
    .scrollable-container::-webkit-scrollbar { width: 5px; }
    .scrollable-container::-webkit-scrollbar-track { background: #f0f0f0; border-radius: 3px; }
    .scrollable-container::-webkit-scrollbar-thumb { background: #c0c0c0; border-radius: 3px; }
    .scrollable-container::-webkit-scrollbar-thumb:hover { background: #a0a0a0; }

    .category-card {
      border: 1px solid #dee2e6;
      border-radius: 4px;
      overflow: hidden;
      margin-bottom: .625rem;
      transition: box-shadow .2s ease;
      background: #fff;
    }
    .category-card:hover { box-shadow: 0 2px 8px rgba(0, 0, 0, .08); }

    .category-header {
      cursor: pointer;
      user-select: none;
      background: #fff;
      color: #212529;
      padding: .5rem .75rem;
      border-bottom: 1px solid #dee2e6;
      transition: background .15s ease;
    }
    .category-header:hover { background: #f8f9fa; }

    .category-title {
      font-size: .75rem;
      font-weight: 600;
      color: #212529;
    }

    .toggle-icon {
      transition: transform .25s ease;
      color: #4a4a4a;
      font-size: .6rem;
    }

    .count-badge { font-size: .5625rem; padding: .15rem .4rem; }

    .btn-criar-perm {
      font-size: .75rem;
      padding: 0;
      background: none;
      color: #198754;
      border: none;
      transition: color .15s ease, transform .15s ease;
      line-height: 1;
      display: inline-flex;
      align-items: center;
    }
    .btn-criar-perm:hover {
      color: #146c43;
      transform: scale(1.2);
    }
    .btn-criar-perm i { font-size: .875rem; }

    .category-body { background: #fff !important; }

    .permission-item {
      padding: .4rem .75rem;
      border-bottom: 1px solid #e8e8e8;
      display: flex;
      align-items: center;
      justify-content: space-between;
      background: #fff;
      transition: all .15s ease;
      cursor: default;
    }
    .permission-item:last-child { border-bottom: none; }
    .permission-item:hover {
      background: rgba(0, 0, 0, .03);
      border-color: #dee2e6;
      transform: translateX(2px);
    }

    .permission-name { font-size: .6875rem; color: #212529; font-weight: 400; }
    .permission-item .badge { font-size: .5625rem; padding: .125rem .35rem; }

    .btn-acao-group { display: inline-flex; align-items: center; gap: .35rem; }
    .btn-acao {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: .15rem .35rem;
      font-size: .625rem;
      border-radius: 4px;
      line-height: 1;
      border: none;
      transition: opacity .15s;
      cursor: pointer;
    }
    .btn-acao:hover { opacity: .82; }
    .btn-acao-editar  { background: #fff; color: #e6a817; border: 1px solid #ffc107 !important; }
    .btn-acao-editar:hover  { background: #ffc107; color: #212529; opacity: 1; }
    .btn-acao-excluir { background: #fff; color: #dc3545; border: 1px solid #dc3545 !important; }
    .btn-acao-excluir:hover { background: #dc3545; color: #fff; opacity: 1; }
    .btn-acao-reativar { background: #fff; color: #198754; border: 1px solid #198754 !important; }
    .btn-acao-reativar:hover { background: #198754; color: #fff; opacity: 1; }

    .filtro-inativos-wrap { display: inline-flex; align-items: center; gap: 0; }
    .filtro-inativos-opcoes {
      display: inline-flex;
      align-items: center;
      gap: .25rem;
      max-width: 0;
      overflow: hidden;
      opacity: 0;
      transition: max-width .3s ease, opacity .25s ease, margin .3s ease;
      margin-left: 0;
      white-space: nowrap;
    }
    .filtro-inativos-opcoes.aberto {
      max-width: 260px;
      opacity: 1;
      margin-left: .35rem;
    }
    .btn-filtro-opcao {
      font-size: .6875rem;
      padding: .25rem .6rem;
      border-radius: 4px;
      border: 1px solid #dc3545;
      background: #fff;
      color: #dc3545;
      cursor: pointer;
      transition: background .15s, color .15s, text-decoration .1s;
      line-height: 1.4;
      text-decoration: none;
    }
    .btn-filtro-opcao:hover { text-decoration: underline; }
    .btn-filtro-opcao.ativo { background: #dc3545; color: #fff; text-decoration: none; }
    .btn-filtro-opcao.ativo:hover { text-decoration: underline; }

    #btnToggleInativos { transition: background .15s, color .15s, border-color .15s; }
    #btnToggleInativos:hover { text-decoration: underline; }

    .badge-modo-inativo {
      font-size: .5rem;
      padding: .1rem .3rem;
      background: #dc3545 !important;
      color: #fff !important;
      border-radius: 3px;
      letter-spacing: .02em;
      text-transform: uppercase;
      animation: sk-fadeIn .2s ease;
    }

    .count-badge-inativo.count-badge {
      background: #dc3545 !important;
      color: #ffe0e3 !important;
    }

    .perm-sigla-inativa {
      background: #f8d7da !important;
      color: #dc3545 !important;
      border: 1px solid #f5c2c7;
    }

    .modal-content {
      border: 1px solid #dee2e6;
      border-radius: 6px;
      box-shadow: 0 8px 30px rgba(0, 0, 0, .2);
    }
    .modal-header-sk {
      background: #fff;
      color: #212529;
      padding: .65rem 1rem;
      border-bottom: 1px solid #dee2e6;
      border-radius: 6px 6px 0 0;
    }
    .modal-header-sk .modal-title { font-size: .8125rem; font-weight: 600; color: #212529; }
    .modal-header-sk .btn-close-white { filter: none; opacity: .7; }
    .modal-header-sk .btn-close-white:hover { opacity: 1; }

    .modal-body { padding: 1.25rem; font-size: .6875rem; }
    .modal-body .form-label { font-size: .6875rem; font-weight: 600; color: #212529; margin-bottom: 0.25rem; }
    .modal-body .form-control,
    .modal-body .form-select { font-size: .6875rem; padding: .35rem .6rem; }
    .modal-body .form-control:focus,
    .modal-body .form-select:focus { border-color: #4a4a4a; box-shadow: 0 0 0 0.2rem rgba(74, 74, 74, 0.15); }
    .modal-body .form-control-plaintext { font-size: .6875rem; }
    .modal-body .form-control[readonly] { background-color: #f8f9fa; }
    .modal-body .input-group-text { font-size: .6875rem; padding: .35rem .6rem; }
    .modal-body .form-text { font-size: .625rem; }
    .modal-body .badge { font-size: .5625rem; }

    .modal-footer-sk {
      background: #f8f9fa;
      border-top: 1px solid #dee2e6;
      padding: .6rem 1rem;
    }
    .modal-footer-sk .btn { font-size: .6875rem; padding: .3rem .75rem; }

    .modal-header-warning-sk {
      background: #fff;
      color: #212529;
      padding: .65rem 1rem;
      border-bottom: 1px solid #dee2e6;
      border-radius: 6px 6px 0 0;
    }
    .modal-header-warning-sk .modal-title { font-size: .8125rem; font-weight: 600; color: #212529; }
    .modal-header-warning-sk i { color: #dc3545; }

    @keyframes shimmer {
      from { background-position: -600px 0; }
      to   { background-position:  600px 0; }
    }
    .sk-block {
      background: linear-gradient(90deg, #e8e8e8 25%, #f5f5f5 50%, #e8e8e8 75%);
      background-size: 600px 100%;
      animation: shimmer 1.4s infinite linear;
      border-radius: 4px;
    }
    .sk-cat-wrap { padding: .5rem .75rem; border-bottom: 1px solid #dee2e6; }
    .sk-cat-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: .5rem .75rem;
      border-bottom: 1px solid #dee2e6;
    }
    .sk-cat-header-left { display: flex; align-items: center; gap: .5rem; }
    .sk-perm-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: .4rem .75rem;
      border-bottom: 1px solid #e8e8e8;
    }
    .sk-perm-left { display: flex; align-items: center; gap: .5rem; }

    .empty-state { padding: 1.5rem; text-align: center; color: #888; }
    .empty-state i { font-size: 1.5rem; margin-bottom: .375rem; display: block; }
    .empty-state p { font-size: .6875rem; margin: 0; }

    @keyframes sk-fadeIn { from { opacity: 0; transform: translateY(-4px); } to { opacity: 1; transform: translateY(0); } }
    .sk-fade-in { animation: sk-fadeIn .25s ease; }

    .sk-footer {
      background: #fff;
      border-top: 1px solid #dee2e6;
      margin-top: 2rem;
      padding-block: .75rem;
      text-align: center;
    }
    .sk-footer p { margin: 0; font-size: .6875rem; color: #a0a0a0; letter-spacing: .02em; }

    .toast-wrapper {
      position: fixed;
      top: 1rem;
      right: 1rem;
      z-index: 11000;
      display: flex;
      flex-direction: column;
      gap: .5rem;
      min-width: 300px;
      max-width: 380px;
      pointer-events: none;
    }
    .toast-item {
      background: #fff;
      border-radius: 6px;
      overflow: hidden;
      box-shadow: 0 4px 16px rgba(0, 0, 0, .12);
      border: 1px solid #e8e8e8;
      pointer-events: all;
      animation: entrar-toast .2s ease;
    }
    .toast-item-success { border-left: 4px solid #198754; }
    .toast-item-warning { border-left: 4px solid #ffc107; }
    .toast-item-danger  { border-left: 4px solid #dc3545; }
    .toast-item-info    { border-left: 4px solid #6c757d; }
    @keyframes entrar-toast {
      from { opacity: 0; transform: translateX(20px); }
      to   { opacity: 1; transform: translateX(0); }
    }
    .toast-corpo {
      display: flex;
      align-items: center;
      gap: .6rem;
      padding: .65rem .85rem;
      font-size: .85rem;
      color: #333;
    }
    .toast-fechar {
      margin-left: auto;
      background: none;
      border: none;
      font-size: 1rem;
      color: #999;
      cursor: pointer;
      padding: 0 0 0 .5rem;
    }
    .toast-fechar:hover { color: #333; }
    .toast-barra { height: 3px; animation: encolher linear forwards; }
    .toast-barra-success { background: #198754; }
    .toast-barra-warning { background: #ffc107; }
    .toast-barra-danger  { background: #dc3545; }
    .toast-barra-info    { background: #6c757d; }
    @keyframes encolher { from { width: 100%; } to { width: 0%; } }

    @media (max-width: 1440px) { .scrollable-container { max-height: calc(100vh - 180px); } }
    @media (min-width: 1920px) {
      .scrollable-container { max-height: calc(90vh - 220px); }
      .section-header { padding: .875rem 1rem; }
      .category-header { padding: .625rem .875rem; }
    }
  </style>
</head>
<body class="d-flex flex-column min-vh-100">

  <div class="toast-wrapper" id="toastWrapper"></div>

  <div class="d-flex align-items-center justify-content-between mb-3 mt-3 px-3">
    <h4 class="fw-bold mb-0 page-title">Gerenciar Menus</h4>
    <div class="d-flex gap-2">
      <a href="../Usuario/AcessosUsuarios.asp" class="btn btn-outline-secondary btn-sm">
        <i class="bi bi-arrow-left me-1"></i> Acessos de Usuário
      </a>
    </div>
  </div>

  <main class="flex-grow-1">
    <div class="py-3 px-3">

      <div class="action-bar p-2 mb-2">
        <div class="d-flex justify-content-between align-items-center gap-2">

          <div class="filtro-inativos-wrap">
            <button type="button" class="btn btn-sm btn-outline-danger" id="btnToggleInativos">
              <i class="bi bi-eye-slash me-1"></i>Inativos<i class="bi bi-chevron-right ms-1" id="iconChevronInativos" style="font-size:.6rem;transition:transform .3s ease;"></i>
            </button>
            <div class="filtro-inativos-opcoes" id="filtroInativosOpcoes">
              <button type="button" class="btn-filtro-opcao" data-escopo="todos">Todos</button>
              <button type="button" class="btn-filtro-opcao" data-escopo="servicos">Serviços</button>
              <button type="button" class="btn-filtro-opcao" data-escopo="consultas">Consultas</button>
            </div>
          </div>

          <div style="max-width:320px;width:100%;">
            <div class="input-group input-group-sm">
              <span class="input-group-text bg-white border-end-0">
                <i class="bi bi-search text-muted"></i>
              </span>
              <input type="text" class="form-control border-start-0" id="inputPesquisa"
                     placeholder="Pesquisar permissões..." autocomplete="off">
            </div>
          </div>

        </div>
      </div>

      <div class="row g-3" id="gridPrincipal">
        <div class="col-lg-6">
          <div class="card border-0 shadow-sm overflow-hidden rounded-3">
            <div class="section-header d-flex justify-content-between align-items-center"
                 id="headerServicos" role="button" data-secao="servicos">
              <div class="d-flex align-items-center gap-2">
                <i class="bi bi-chevron-down section-toggle-icon" id="iconToggleServicos"></i>
                <i class="bi bi-gear"></i>
                <span class="section-title">Serviços</span>
                <span class="badge rounded-pill" id="totalServicos">0</span>
              </div>
            </div>
            <div class="section-body">
              <div id="containerServicos" class="scrollable-container">
                <div id="skeletonServicos">
                  <div class="sk-cat-header">
                    <div class="sk-cat-header-left">
                      <div class="sk-block" style="width:10px;height:10px;border-radius:2px;"></div>
                      <div class="sk-block" style="width:120px;height:10px;"></div>
                      <div class="sk-block" style="width:22px;height:14px;border-radius:10px;"></div>
                    </div>
                    <div class="sk-block" style="width:18px;height:18px;border-radius:50%;"></div>
                  </div>
                  <div class="sk-perm-row"><div class="sk-perm-left"><div class="sk-block" style="width:18px;height:14px;border-radius:10px;"></div><div class="sk-block" style="width:180px;height:9px;"></div></div><div class="sk-block" style="width:52px;height:18px;border-radius:4px;"></div></div>
                  <div class="sk-perm-row"><div class="sk-perm-left"><div class="sk-block" style="width:18px;height:14px;border-radius:10px;"></div><div class="sk-block" style="width:140px;height:9px;"></div></div><div class="sk-block" style="width:52px;height:18px;border-radius:4px;"></div></div>
                  <div class="sk-perm-row"><div class="sk-perm-left"><div class="sk-block" style="width:18px;height:14px;border-radius:10px;"></div><div class="sk-block" style="width:210px;height:9px;"></div></div><div class="sk-block" style="width:52px;height:18px;border-radius:4px;"></div></div>
                  <div class="sk-cat-header" style="margin-top:.25rem;">
                    <div class="sk-cat-header-left">
                      <div class="sk-block" style="width:10px;height:10px;border-radius:2px;"></div>
                      <div class="sk-block" style="width:95px;height:10px;"></div>
                      <div class="sk-block" style="width:22px;height:14px;border-radius:10px;"></div>
                    </div>
                    <div class="sk-block" style="width:18px;height:18px;border-radius:50%;"></div>
                  </div>
                  <div class="sk-perm-row"><div class="sk-perm-left"><div class="sk-block" style="width:18px;height:14px;border-radius:10px;"></div><div class="sk-block" style="width:160px;height:9px;"></div></div><div class="sk-block" style="width:52px;height:18px;border-radius:4px;"></div></div>
                  <div class="sk-perm-row"><div class="sk-perm-left"><div class="sk-block" style="width:18px;height:14px;border-radius:10px;"></div><div class="sk-block" style="width:200px;height:9px;"></div></div><div class="sk-block" style="width:52px;height:18px;border-radius:4px;"></div></div>
                </div>
              </div>
            </div>
          </div>
        </div>

        <div class="col-lg-6">
          <div class="card border-0 shadow-sm overflow-hidden rounded-3">
            <div class="section-header d-flex justify-content-between align-items-center"
                 id="headerConsultas" role="button" data-secao="consultas">
              <div class="d-flex align-items-center gap-2">
                <i class="bi bi-chevron-down section-toggle-icon" id="iconToggleConsultas"></i>
                <i class="bi bi-bar-chart"></i>
                <span class="section-title">Consultas</span>
                <span class="badge rounded-pill" id="totalConsultas">0</span>
              </div>
            </div>
            <div class="section-body">
              <div id="containerConsultas" class="scrollable-container">
                <div id="skeletonConsultas">
                  <div class="sk-cat-header">
                    <div class="sk-cat-header-left">
                      <div class="sk-block" style="width:10px;height:10px;border-radius:2px;"></div>
                      <div class="sk-block" style="width:110px;height:10px;"></div>
                      <div class="sk-block" style="width:22px;height:14px;border-radius:10px;"></div>
                    </div>
                    <div class="sk-block" style="width:18px;height:18px;border-radius:50%;"></div>
                  </div>
                  <div class="sk-perm-row"><div class="sk-perm-left"><div class="sk-block" style="width:18px;height:14px;border-radius:10px;"></div><div class="sk-block" style="width:190px;height:9px;"></div></div><div class="sk-block" style="width:52px;height:18px;border-radius:4px;"></div></div>
                  <div class="sk-perm-row"><div class="sk-perm-left"><div class="sk-block" style="width:18px;height:14px;border-radius:10px;"></div><div class="sk-block" style="width:130px;height:9px;"></div></div><div class="sk-block" style="width:52px;height:18px;border-radius:4px;"></div></div>
                  <div class="sk-cat-header" style="margin-top:.25rem;">
                    <div class="sk-cat-header-left">
                      <div class="sk-block" style="width:10px;height:10px;border-radius:2px;"></div>
                      <div class="sk-block" style="width:140px;height:10px;"></div>
                      <div class="sk-block" style="width:22px;height:14px;border-radius:10px;"></div>
                    </div>
                    <div class="sk-block" style="width:18px;height:18px;border-radius:50%;"></div>
                  </div>
                  <div class="sk-perm-row"><div class="sk-perm-left"><div class="sk-block" style="width:18px;height:14px;border-radius:10px;"></div><div class="sk-block" style="width:175px;height:9px;"></div></div><div class="sk-block" style="width:52px;height:18px;border-radius:4px;"></div></div>
                  <div class="sk-perm-row"><div class="sk-perm-left"><div class="sk-block" style="width:18px;height:14px;border-radius:10px;"></div><div class="sk-block" style="width:220px;height:9px;"></div></div><div class="sk-block" style="width:52px;height:18px;border-radius:4px;"></div></div>
                  <div class="sk-perm-row"><div class="sk-perm-left"><div class="sk-block" style="width:18px;height:14px;border-radius:10px;"></div><div class="sk-block" style="width:155px;height:9px;"></div></div><div class="sk-block" style="width:52px;height:18px;border-radius:4px;"></div></div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

    </div>
  </main>

  <template id="tplPermissaoItem">
    <div class="permission-item sk-fade-in"
         data-codigo="" data-tipo="" data-url="" data-banco="" data-extensao="">
      <div class="d-flex align-items-center gap-2">
        <span class="badge rounded-pill text-bg-secondary perm-sigla"
              style="font-size:.5rem;padding:.1rem .3rem;"></span>
        <span class="permission-name perm-nome"></span>
      </div>
      <div class="btn-acao-group">
        <button type="button" class="btn-acao btn-acao-editar" title="Editar">
          <i class="bi bi-pencil-fill"></i>
        </button>
        <button type="button" class="btn-acao btn-acao-excluir" title="Excluir">
          <i class="bi bi-trash-fill"></i>
        </button>
      </div>
    </div>
  </template>

  <template id="tplCategoriaCard">
    <div class="category-card sk-fade-in" data-tipo="" data-codigo="" data-nome="">
      <div class="category-header d-flex justify-content-between align-items-center"
           role="button" aria-expanded="false">
        <div class="d-flex align-items-center gap-2">
          <i class="bi bi-chevron-right toggle-icon" style="font-size:.6rem;"></i>
          <span class="category-title cat-nome"></span>
          <span class="badge rounded-pill count-badge cat-badge"></span>
        </div>
        <div>
          <button type="button" class="btn-criar-perm" title="Nova Permissão">
            <i class="bi bi-plus-circle-fill"></i>
          </button>
        </div>
      </div>
      <div class="collapse">
        <div class="category-body"></div>
      </div>
    </div>
  </template>

  <div class="modal fade" id="modalCriarPermissao" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered modal-lg">
      <div class="modal-content">
        <div class="modal-header-sk position-relative">
          <h5 class="modal-title d-flex align-items-center gap-2 mb-0">
            <i class="bi bi-key"></i>
            <span id="modalPermTitulo">Nova Permissão</span>
          </h5>
          <button type="button" class="btn-close btn-close-white position-absolute top-50 end-0 translate-middle-y me-3"
                  data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body">
          <input type="hidden" id="modalPermModo" value="criar">
          <input type="hidden" id="modalPermCodigo" value="">
          <input type="hidden" id="modalPermTipo" value="">
          <input type="hidden" id="modalPermCategoria" value="">

          <div class="mb-3">
            <div class="d-flex align-items-center justify-content-between">
              <label class="form-label fw-semibold mb-1" for="inputCategoriaPerm">Categoria</label>
              <span class="badge" style="background-color:#cfe2ff;color:#084298;">Somente leitura</span>
            </div>
            <div class="bg-light border rounded px-3 py-2">
              <input type="text" id="inputCategoriaPerm" class="form-control-plaintext py-0" readonly>
            </div>
          </div>

          <div class="mb-3 d-none" id="grupoMoverCategoria">
            <label class="form-label fw-semibold" for="selectCategoriaPerm">
              Mover para categoria <span class="badge ms-1" style="background-color:#fff3cd;color:#842029;">Obrigatório</span>
            </label>
            <div class="input-group">
              <span class="input-group-text bg-white">
                <i class="bi bi-folder2-open text-muted"></i>
              </span>
              <select id="selectCategoriaPerm" class="form-select">
                <option value="">Selecione...</option>
              </select>
            </div>
          </div>

          <div class="mb-3">
            <label class="form-label fw-semibold" for="inputNomePermissao">
              Nome da Permissão <span class="badge ms-1" style="background-color:#fff3cd;color:#842029;">Obrigatório</span>
            </label>
            <div class="input-group">
              <span class="input-group-text bg-white">
                <i class="bi bi-tag text-muted"></i>
              </span>
              <input type="text" id="inputNomePermissao" class="form-control"
                     placeholder="Ex: Consulta de Estoque" maxlength="50">
            </div>
          </div>

          <div class="mb-3">
            <label class="form-label fw-semibold" for="inputPermUrl">
              URL do Arquivo <span class="badge ms-1" style="background-color:#fff3cd;color:#842029;">Obrigatório</span>
            </label>
            <div class="input-group">
              <span class="input-group-text bg-white">
                <i class="bi bi-link-45deg text-muted"></i>
              </span>
              <input type="text" id="inputPermUrl" class="form-control"
                     placeholder="Ex: ../../corporativo/CriarPermissao/CriarPermissao"
                     maxlength="300" autocomplete="off" autocapitalize="off" spellcheck="false">
            </div>
            <div class="form-text text-danger fw-semibold">
              Atenção: NÃO inserir o (.asp).
            </div>
          </div>

          <div class="mb-3 d-none" id="grupoBancoPerm">
            <label class="form-label fw-semibold" for="inputPermBanco">
              Banco <span class="badge ms-1" style="background-color:#fff3cd;color:#842029;">Obrigatório</span>
            </label>
            <div class="input-group">
              <span class="input-group-text bg-white">
                <i class="bi bi-database text-muted"></i>
              </span>
              <input type="text" id="inputPermBanco" class="form-control"
                     placeholder="Ex: HELPDESK" maxlength="50">
            </div>
            <div class="form-text">Obrigatório somente para a seção de Consultas (REL_BANCO).</div>
          </div>

          <div class="mb-3 d-none" id="grupoExtensaoPerm">
            <label class="form-label fw-semibold" for="inputPermExtensao">
              Extensão do Arquivo
            </label>
            <div class="input-group">
              <span class="input-group-text bg-white">
                <i class="bi bi-file-code text-muted"></i>
              </span>
              <input type="text" id="inputPermExtensao" class="form-control bg-light"
                     readonly tabindex="-1">
            </div>
            <div class="form-text">
              Esse campo é apenas para refletir se a coluna no banco está preenchida com o arquivo <code>.asp</code> ou não.
            </div>
          </div>
        </div>
        <div class="modal-footer-sk d-flex justify-content-end">
          <button type="button" class="btn btn-outline-secondary" data-bs-dismiss="modal">Cancelar</button>
          <button type="button" class="btn btn-success ms-2" id="btnSalvarPermissao"><span id="textoBotaoSalvarPerm">Confirmar</span></button>
        </div>
      </div>
    </div>
  </div>

  <div class="modal fade" id="modalConfirmacao" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered">
      <div class="modal-content">
        <div class="modal-header-warning-sk position-relative">
          <h5 class="modal-title d-flex align-items-center gap-2 mb-0">
            <i class="bi bi-trash-fill"></i>
            <span id="tituloConfirmacao">Confirmação</span>
          </h5>
          <button type="button" class="btn-close position-absolute top-50 end-0 translate-middle-y me-3"
                  data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body">
          <p class="mb-0" id="textoConfirmacao">Tem certeza que deseja continuar?</p>
        </div>
        <div class="modal-footer-sk d-flex justify-content-end">
          <button type="button" class="btn btn-outline-secondary" data-bs-dismiss="modal">Cancelar</button>
          <button type="button" class="btn btn-danger ms-2" id="btnConfirmarAcao">Confirmar</button>
        </div>
      </div>
    </div>
  </div>

  <div class="modal fade" id="modalReativar" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered">
      <div class="modal-content">
        <div class="modal-header-warning-sk position-relative">
          <h5 class="modal-title d-flex align-items-center gap-2 mb-0">
            <i class="bi bi-arrow-counterclockwise text-success"></i>
            <span>Reativar Menu</span>
          </h5>
          <button type="button" class="btn-close position-absolute top-50 end-0 translate-middle-y me-3"
                  data-bs-dismiss="modal" aria-label="Fechar"></button>
        </div>
        <div class="modal-body">
          <p class="mb-0" id="textoReativar">Ao <strong>Confirmar</strong> a reativação do <strong>Menu</strong>, o Menu será <strong>reativado</strong>, tem certeza que deseja continuar?</p>
        </div>
        <div class="modal-footer-sk d-flex justify-content-end">
          <button type="button" class="btn btn-outline-secondary" data-bs-dismiss="modal">Fechar</button>
          <button type="button" class="btn btn-success ms-2" id="btnConfirmarReativar">Confirmar</button>
        </div>
      </div>
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"></script>
  <script>
    (($) => {
      "use strict";

      const API_URL    = "CriarPermissao.asp";
      const TOAST_CFG  = Object.freeze({ delay: 4500 });
      const TOAST_ICON = Object.freeze({
        success: { cls: "bi-check-circle-fill",         cor: "#198754" },
        warning: { cls: "bi-exclamation-triangle-fill", cor: "#ffc107" },
        danger:  { cls: "bi-x-circle-fill",             cor: "#dc3545" },
        info:    { cls: "bi-info-circle-fill",          cor: "#6c757d" },
      });

      const State = (() => {
        let _dados        = { servicos: { categorias: [] }, consultas: { categorias: [] } };
        let _sessao       = { servicos: false, consultas: false };
        let _pesquisa     = null;
        let _modoInativos = null;
        return Object.freeze({
          getDados:          ()      => _dados,
          setDados:          (d)     => { _dados = d; },
          getSessao:         (s)     => _sessao[s],
          setSessao:         (s, v)  => { _sessao[s] = v; },
          getPesquisa:       ()      => _pesquisa,
          setPesquisa:       (v)     => { _pesquisa = v; },
          getModoInativos:   ()      => _modoInativos,
          setModoInativos:   (v)     => { _modoInativos = v; },
        });
      })();

      const Html = Object.freeze({
        esc(txt) {
          if (!txt) return "";
          return String(txt)
            .replace(/&/g, "&amp;").replace(/</g, "&lt;")
            .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
        },
        escAttr(txt) {
          return String(txt ?? "")
            .replaceAll("&", "&amp;").replaceAll('"', "&quot;")
            .replaceAll("'", "&#039;").replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;");
        },
      });

      const UI = Object.freeze({

        notificar(mensagem, tipo = "info") {
          const cfg  = TOAST_ICON[tipo] ?? TOAST_ICON.info;
          const item = $(`
            <div class="toast-item toast-item-${tipo}">
              <div class="toast-corpo">
                <i class="bi ${cfg.cls}" style="color:${cfg.cor};font-size:1rem;flex-shrink:0;"></i>
                <span class="flex-grow-1">${mensagem}</span>
                <button class="toast-fechar" aria-label="Fechar">&times;</button>
              </div>
              <div class="toast-barra toast-barra-${tipo}" style="animation-duration:${TOAST_CFG.delay}ms;"></div>
            </div>
          `);
          item.find(".toast-fechar").on("click", () => item.remove());
          $("#toastWrapper").append(item);
          setTimeout(() => item.fadeOut(300, () => item.remove()), TOAST_CFG.delay);
        },

        fecharModal(id) {
          bootstrap.Modal.getInstance(document.getElementById(id))?.hide();
        },

        confirmar(opcoes) {
          const cfg = {
            titulo: "Confirmação",
            mensagem: "Tem certeza que deseja continuar?",
            textoBotao: "Confirmar",
            classeBotao: "btn-dark",
            iconeBotao: "bi-check-lg",
            aoConfirmar: null,
            ...opcoes,
          };

          const modalEl = document.getElementById("modalConfirmacao");
          const modal   = bootstrap.Modal.getOrCreateInstance(modalEl, { backdrop: "static", keyboard: false });

          $("#tituloConfirmacao").text(cfg.titulo);
          $("#textoConfirmacao").html(cfg.mensagem);

          const $btn = $("#btnConfirmarAcao");
          $btn.removeClass("btn-primary btn-success btn-danger btn-warning btn-dark btn-outline-danger")
              .addClass(cfg.classeBotao)
              .html(`${cfg.iconeBotao ? `<i class="bi ${cfg.iconeBotao} me-1"></i> ` : ""}${cfg.textoBotao}`);

          $btn.off("click.confirmar").on("click.confirmar", () => {
            modal.hide();
            cfg.aoConfirmar?.();
          });

          modal.show();
        },
      });

      const Api = Object.freeze({
        post(payload) {
          return new Promise((resolve, reject) => {
            $.ajax({
              url:         API_URL,
              type:        "POST",
              data:        payload,
              dataType:    "text",
              success(raw) {
                try { resolve(JSON.parse(raw)); }
                catch { reject(new Error("Resposta inválida do servidor.")); }
              },
              error() {
                reject(new Error("Erro de comunicação com o servidor."));
              },
            });
          });
        },
      });

      const Render = Object.freeze({
        _permissaoItem(tipo, perm, modoInativo) {
          const tpl = document.getElementById("tplPermissaoItem").content.cloneNode(true);
          const el  = tpl.querySelector(".permission-item");

          el.dataset.codigo   = perm.codigo  ?? "";
          el.dataset.tipo     = tipo;
          el.dataset.url      = perm.url      ?? "";
          el.dataset.banco    = perm.banco    ?? "";
          el.dataset.extensao = perm.extensao ?? "";

          el.querySelector(".perm-sigla").textContent = tipo === "servico" ? "S" : "C";
          el.querySelector(".perm-nome").textContent  = perm.nome ?? "";

          if (modoInativo) {
            el.querySelector(".perm-sigla").classList.add("perm-sigla-inativa");
            el.querySelector(".btn-acao-group").innerHTML =
              `<button type="button" class="btn-acao btn-acao-reativar" title="Reativar"><i class="bi bi-arrow-counterclockwise"></i></button>`;
          }

          return el;
        },

        _categoriaCard(tipo, cat, idx, modoInativo) {
          const collapseId = `collapse-${tipo}-${idx}`;
          const permissoes = cat.permissoes ?? [];
          const tpl        = document.getElementById("tplCategoriaCard").content.cloneNode(true);
          const card       = tpl.querySelector(".category-card");
          const header     = card.querySelector(".category-header");
          const collapse   = card.querySelector(".collapse");
          const body       = card.querySelector(".category-body");

          card.dataset.tipo   = tipo;
          card.dataset.codigo = cat.codigo;
          card.dataset.nome   = cat.nome;

          collapse.id = collapseId;
          header.setAttribute("data-bs-toggle", "collapse");
          header.setAttribute("data-bs-target", `#${collapseId}`);

          card.querySelector(".cat-nome").textContent  = cat.nome ?? "";
          const badge = card.querySelector(".cat-badge");
          badge.textContent = permissoes.length;
          if (modoInativo) {
            badge.className = `badge rounded-pill count-badge count-badge-inativo`;
          } else {
            badge.className = `badge rounded-pill count-badge ${permissoes.length > 0 ? "text-bg-dark" : "text-bg-secondary"}`;
          }

          const btnCriar = card.querySelector(".btn-criar-perm");
          if (modoInativo) {
            btnCriar.remove();
          } else {
            btnCriar.dataset.tipo   = tipo;
            btnCriar.dataset.codigo = cat.codigo;
            btnCriar.dataset.nome   = cat.nome;
          }

          if (permissoes.length) {
            permissoes.forEach((p) => body.appendChild(this._permissaoItem(tipo, p, modoInativo)));
          } else {
            body.innerHTML = `<div class="empty-state"><i class="bi bi-inbox"></i><p>Nenhuma permissão nesta categoria</p></div>`;
          }

          return card;
        },

        secao(containerId, totalId, tipo, categorias, modoInativo) {
          const $container = $(containerId);
          $container.empty();

          const headerBadgeId = tipo === "servico" ? "#totalServicos" : "#totalConsultas";
          const $headerEl     = $(totalId).closest(".section-header").find(".badge-modo-inativo");
          $headerEl.remove();
          if (modoInativo) {
            $(totalId).after(`<span class="badge-modo-inativo ms-1">inativos</span>`);
          }

          if (!categorias?.length) {
            $container.html(`<div class="empty-state"><i class="bi bi-inbox"></i><p>${modoInativo ? "Nenhum inativo nesta seção" : "Nenhuma categoria cadastrada"}</p></div>`);
            $(totalId).text("0");
            return;
          }

          const total    = categorias.reduce((acc, cat) => acc + (cat.permissoes?.length ?? 0), 0);
          const fragment = document.createDocumentFragment();
          categorias.forEach((cat, idx) => fragment.appendChild(this._categoriaCard(tipo, cat, idx, modoInativo)));
          $container[0].appendChild(fragment);
          $(totalId).text(total);
        },
      });

      const Dados = Object.freeze({
        async carregar() {
          try {
            const data = await Api.post({ acao: "listarTudo" });
            if (!data.success) { UI.notificar(`Erro ao carregar dados: ${data.mensagem}`, "danger"); return; }
            State.setDados(data);
            $("#skeletonServicos").remove();
            $("#skeletonConsultas").remove();
            Render.secao("#containerServicos",  "#totalServicos",  "servico",  data.servicos.categorias,  false);
            Render.secao("#containerConsultas", "#totalConsultas", "consulta", data.consultas.categorias, false);
          } catch (err) {
            UI.notificar(err.message, "danger");
            $("#skeletonServicos, #skeletonConsultas").remove();
            $("#containerServicos, #containerConsultas").html(
              `<div class="empty-state"><i class="bi bi-exclamation-triangle"></i><p>Erro ao carregar</p></div>`
            );
          }
        },

        async carregarInativos(escopo) {
          try {
            const data = await Api.post({ acao: "listarInativos", escopo });
            if (!data.success) { UI.notificar(`Erro: ${data.mensagem}`, "danger"); return; }

            if (escopo === "todos" || escopo === "servicos") {
              Render.secao("#containerServicos",  "#totalServicos",  "servico",  data.servicos.categorias,  true);
            }
            if (escopo === "todos" || escopo === "consultas") {
              Render.secao("#containerConsultas", "#totalConsultas", "consulta", data.consultas.categorias, true);
            }
          } catch (err) {
            UI.notificar(err.message, "danger");
          }
        },
      });

      const Modal = Object.freeze({
        _resetForm() {
          $("#inputNomePermissao, #inputPermUrl, #inputPermBanco, #inputPermExtensao").val("");
          $("#selectCategoriaPerm").off("change").val("");
          $("#grupoMoverCategoria, #grupoExtensaoPerm").addClass("d-none");
        },

        _abrir(cfg) {
          this._resetForm();
          $("#modalPermTitulo").text(cfg.titulo);
          $("#textoBotaoSalvarPerm").text(cfg.textoBotao);
          $("#modalPermModo").val(cfg.modo);
          $("#modalPermCodigo").val(cfg.codigo ?? "");
          $("#modalPermTipo").val(cfg.tipo);
          $("#modalPermCategoria").val(cfg.codigoCategoria);
          $("#inputCategoriaPerm").val(cfg.nomeCategoria);
          $("#grupoBancoPerm").toggleClass("d-none", cfg.tipo !== "consulta");
          new bootstrap.Modal(document.getElementById("modalCriarPermissao")).show();
          setTimeout(() => $("#inputNomePermissao").trigger("focus"), 250);
        },

        abrirCriar(tipo, codigoCategoria, nomeCategoria) {
          const label = tipo === "servico" ? "Serviço" : "Consulta";
          this._abrir({ titulo: `Nova Permissão de ${label}`, textoBotao: "Confirmar", modo: "criar", tipo, codigoCategoria, nomeCategoria });
        },

        abrirEditar(tipo, codigo) {
          const $item = $(`.permission-item[data-tipo="${tipo}"][data-codigo="${codigo}"]`);
          const $card = $item.closest(".category-card");
          const cats  = tipo === "servico"
            ? (State.getDados().servicos?.categorias  ?? [])
            : (State.getDados().consultas?.categorias ?? []);

          const nomeAtual       = $item.find(".perm-nome").text().trim();
          const urlAtual        = $item.data("url")      ?? "";
          const bancoAtual      = $item.data("banco")    ?? "";
          const extensaoAtual   = $item.data("extensao") ?? "";
          const categoriaCodigo = $card.data("codigo")   ?? "";
          const categoriaNome   = ($card.data("nome") ?? $card.find(".cat-nome").first().text()).trim();

          this._abrir({ titulo: "Editar Permissão", textoBotao: "Confirmar", modo: "editar", tipo, codigo, codigoCategoria: categoriaCodigo, nomeCategoria: categoriaNome || categoriaCodigo });

          const $select = $("#selectCategoriaPerm");
          $select.empty().append('<option value="">Selecione...</option>');
          cats.forEach((c) => $select.append(`<option value="${Html.escAttr(c.codigo)}">${Html.esc(c.nome)}</option>`));
          $select.val(categoriaCodigo);
          $select.on("change", function () { $("#modalPermCategoria").val($(this).val()); });

          $("#inputNomePermissao").val(nomeAtual);
          $("#inputPermUrl").val(urlAtual);
          $("#inputPermBanco").val(bancoAtual);
          $("#inputPermExtensao").val(extensaoAtual);
          $("#grupoExtensaoPerm, #grupoMoverCategoria").removeClass("d-none");

          setTimeout(() => $("#inputNomePermissao").trigger("focus").trigger("select"), 200);
        },
      });

      const Permissao = Object.freeze({
        async _enviar(payload, msgSucesso) {
          try {
            const data = await Api.post(payload);
            if (data.success) {
              const duracao = payload.acao === "criarPermissao" ? 8000 : 4000;
              UI.notificar(msgSucesso, "success");
              await Dados.carregar();
            } else {
              UI.notificar(`Erro: ${data.mensagem}`, "danger");
            }
          } catch (err) {
            UI.notificar(err.message, "danger");
          }
        },

        confirmarSalvar() {
          const modo            = $("#modalPermModo").val();
          const tipo            = $("#modalPermTipo").val();
          const codigoCategoria = $("#modalPermCategoria").val();
          const codigo          = $("#modalPermCodigo").val();
          const nome            = $("#inputNomePermissao").val().trim();
          const url             = $("#inputPermUrl").val().trim();
          const banco           = $("#inputPermBanco").val().trim();

          if (!nome) { UI.notificar("Informe o nome da permissão.", "warning"); $("#inputNomePermissao").trigger("focus"); return; }
          if (!url)  { UI.notificar("Informe a URL do arquivo.", "warning");    $("#inputPermUrl").trigger("focus");      return; }
          if (modo === "editar" && !codigoCategoria) { UI.notificar("Selecione a categoria destino.", "warning"); $("#selectCategoriaPerm").trigger("focus"); return; }
          if (tipo === "consulta" && !banco)          { UI.notificar("Informe o Banco (REL_BANCO).", "warning");  $("#inputPermBanco").trigger("focus");      return; }

          const isCriar = modo === "criar";

          UI.confirmar({
            titulo:      "Confirmar",
            mensagem:    isCriar ? "Uma nova permissão será criada. Deseja continuar?" : "A permissão será atualizada. Deseja continuar?",
            textoBotao:  "Confirmar",
            classeBotao: "btn-dark",
            iconeBotao:  "bi-check-lg",
            aoConfirmar: async () => {
              UI.fecharModal("modalCriarPermissao");
              const acao    = isCriar ? "criarPermissao" : "editarPermissao";
              const payload = { acao, tipo, nome, url, banco, codigoCategoria };
              if (!isCriar) payload.codigo = codigo;
              const msg = isCriar
                ? "Permissão criada com sucesso! Vá em Acessos de Usuários e atribua ao usuário."
                : "Permissão atualizada com sucesso!";
              await this._enviar(payload, msg);
            },
          });
        },

        confirmarExcluir(tipo, codigo, nome) {
          UI.confirmar({
            titulo:      "Excluir Permissão",
            mensagem:    `Deseja realmente excluir a permissão <strong>"${Html.esc(nome)}"</strong>?`,
            textoBotao:  "Excluir",
            classeBotao: "btn-danger",
            iconeBotao:  "",
            aoConfirmar: async () => {
              await this._enviar({ acao: "excluirPermissao", tipo, codigo }, "Permissão excluída com sucesso!");
            },
          });
        },
      });

      const Pesquisa = Object.freeze({
        _filtrarSecao($container, termo) {
          $container.find(".permission-item").each(function () {
            const nomeMatch = $(this).find(".perm-nome").text().toLowerCase().includes(termo);
            $(this).toggleClass("d-none", !nomeMatch);
          });

          $container.find(".category-card").each(function () {
            const catMatch = $(this).find(".cat-nome").text().toLowerCase().includes(termo);
            const temMatch = $(this).find(".permission-item:not(.d-none)").length > 0;

            if (catMatch || temMatch) {
              $(this).removeClass("d-none");
              $(this).find(".collapse").collapse("show");
            } else {
              $(this).addClass("d-none");
            }
          });
        },

        configurar() {
          let timeout;
          $("#inputPesquisa").on("input", function () {
            clearTimeout(timeout);
            const termo = $(this).val().trim().toLowerCase();

            timeout = setTimeout(() => {
              if (!termo) {
                $(".category-card, .permission-item").removeClass("d-none");
                const estado = State.getPesquisa();
                if (estado) {
                  $(".category-card .collapse").each(function () {
                    $(this).collapse(estado[$(this).attr("id")] ? "show" : "hide");
                  });
                  State.setPesquisa(null);
                }
                return;
              }

              if (!State.getPesquisa()) {
                const snap = {};
                $(".category-card .collapse").each(function () {
                  snap[$(this).attr("id")] = $(this).hasClass("show");
                });
                State.setPesquisa(snap);
              }

              Pesquisa._filtrarSecao($("#containerServicos"),  termo);
              Pesquisa._filtrarSecao($("#containerConsultas"), termo);
            }, 250);
          });
        },
      });

      const Secao = Object.freeze({
        toggle(secao) {
          const containerId = secao === "servicos" ? "#containerServicos" : "#containerConsultas";
          const iconId      = secao === "servicos" ? "#iconToggleServicos" : "#iconToggleConsultas";
          const expandido   = State.getSessao(secao);
          $(containerId).find(".category-card .collapse").collapse(expandido ? "hide" : "show");
          $(iconId).toggleClass("collapsed", !expandido);
          State.setSessao(secao, !expandido);
        },
      });

      const FiltroInativos = (() => {
        let _aberto = false;

        return Object.freeze({

        _sincronizarBotoes(escopoAtivo) {
          $("#filtroInativosOpcoes .btn-filtro-opcao").each(function () {
            $(this).toggleClass("ativo", $(this).data("escopo") === escopoAtivo);
          });
        },

        _sincronizarTodos() {
          const servAtivo  = State.getModoInativos() === "servicos"  || State.getModoInativos() === "todos";
          const consAtivo  = State.getModoInativos() === "consultas" || State.getModoInativos() === "todos";
          const ambosAtivo = servAtivo && consAtivo;
          $("#filtroInativosOpcoes .btn-filtro-opcao").each(function () {
            const escopo = $(this).data("escopo");
            if (escopo === "todos") {
              $(this).toggleClass("ativo", ambosAtivo);
            } else if (escopo === "servicos") {
              $(this).toggleClass("ativo", servAtivo && !ambosAtivo);
            } else if (escopo === "consultas") {
              $(this).toggleClass("ativo", consAtivo && !ambosAtivo);
            }
          });
        },

        _aplicarEstadoBotao(ativo) {
          const $btn = $("#btnToggleInativos");
          if (ativo) {
            $btn.removeClass("btn-outline-danger").addClass("btn-danger");
            $btn.find(".bi-eye-slash").removeClass("bi-eye-slash").addClass("bi-eye");
          } else {
            $btn.removeClass("btn-danger").addClass("btn-outline-danger");
            $btn.find(".bi-eye").removeClass("bi-eye").addClass("bi-eye-slash");
          }
          $("#iconChevronInativos").css("transform", _aberto ? "rotate(-90deg)" : "rotate(0deg)");
        },

        async selecionarEscopo(escopo) {
          const modoAtual = State.getModoInativos();
          let novoEscopo  = escopo;

          if (escopo === "servicos" && modoAtual === "consultas") {
            novoEscopo = "todos";
          } else if (escopo === "consultas" && modoAtual === "servicos") {
            novoEscopo = "todos";
          }

          State.setModoInativos(novoEscopo);
          this._sincronizarBotoes(novoEscopo);
          this._aplicarEstadoBotao(true);
          await Dados.carregarInativos(novoEscopo);

          if (modoAtual === "todos" && novoEscopo !== "todos") {
            const dadosAtivos = State.getDados();
            if (novoEscopo === "servicos") {
              Render.secao("#containerConsultas", "#totalConsultas", "consulta", dadosAtivos.consultas?.categorias ?? [], false);
            } else {
              Render.secao("#containerServicos", "#totalServicos", "servico", dadosAtivos.servicos?.categorias ?? [], false);
            }
          }
        },

        async desativar() {
          State.setModoInativos(null);
          this._sincronizarBotoes(null);
          this._aplicarEstadoBotao(false);
          await Dados.carregar();
        },

        toggle() {
          _aberto = !_aberto;
          $("#filtroInativosOpcoes").toggleClass("aberto", _aberto);

          if (!_aberto && State.getModoInativos()) {
            this.desativar();
          } else {
            this._aplicarEstadoBotao(_aberto || !!State.getModoInativos());
          }
        },

        configurar() {
          $("#btnToggleInativos").on("click", () => this.toggle());

          $("#filtroInativosOpcoes").on("click", ".btn-filtro-opcao", (e) => {
            const escopo = $(e.currentTarget).data("escopo");
            this.selecionarEscopo(escopo);
          });
        },
        });
      })();

      function anexarEventos() {
        $(".section-header[data-secao]").on("click", function () {
          Secao.toggle($(this).data("secao"));
        });

        $("#btnSalvarPermissao").on("click", () => Permissao.confirmarSalvar());

        $("#inputNomePermissao, #inputPermUrl, #inputPermBanco").on("keypress", (e) => {
          if (e.which === 13) { e.preventDefault(); Permissao.confirmarSalvar(); }
        });

        $("#modalCriarPermissao").on("shown.bs.modal", () => {
          if ($("#modalPermModo").val() === "criar") $("#inputPermUrl, #inputPermBanco").val("");
        });

        $(document).on("click", ".btn-acao-editar", function () {
          const $item = $(this).closest(".permission-item");
          Modal.abrirEditar($item.data("tipo"), $item.data("codigo"));
        });

        $(document).on("click", ".btn-acao-excluir", function () {
          const $item = $(this).closest(".permission-item");
          Permissao.confirmarExcluir($item.data("tipo"), $item.data("codigo"), $item.find(".perm-nome").text().trim());
        });

        $(document).on("click", ".btn-criar-perm", function () {
          Modal.abrirCriar($(this).data("tipo"), $(this).data("codigo"), $(this).data("nome"));
        });

        $(document).on("click", ".btn-acao-reativar", function () {
          const $item   = $(this).closest(".permission-item");
          const tipo    = $item.data("tipo");
          const codigo  = $item.data("codigo");
          const nome    = $item.find(".perm-nome").text().trim();
          const modalEl = document.getElementById("modalReativar");
          const modal   = bootstrap.Modal.getOrCreateInstance(modalEl, { backdrop: "static", keyboard: false });

          $("#textoReativar").html(`Ao <strong>Confirmar</strong> a reativação do <strong>Menu</strong> <em>"${Html.esc(nome)}"</em>, o Menu será <strong>reativado</strong>, tem certeza que deseja continuar?`);

          $("#btnConfirmarReativar").off("click.reativar").on("click.reativar", async () => {
            modal.hide();
            try {
              const data = await Api.post({ acao: "reativarPermissao", tipo, codigo });
              if (data.success) {
                UI.notificar("Menu reativado com sucesso!", "success");
                const escopoAtual = State.getModoInativos();
                if (escopoAtual) await Dados.carregarInativos(escopoAtual);
              } else {
                UI.notificar(`Erro: ${data.mensagem}`, "danger");
              }
            } catch (err) {
              UI.notificar(err.message, "danger");
            }
          });

          modal.show();
        });

        $(document).on("show.bs.collapse", ".category-card .collapse", function () {
          $(this).closest(".category-card").find(".toggle-icon")
            .removeClass("bi-chevron-right").addClass("bi-chevron-down");
        });

        $(document).on("hide.bs.collapse", ".category-card .collapse", function () {
          $(this).closest(".category-card").find(".toggle-icon")
            .removeClass("bi-chevron-down").addClass("bi-chevron-right");
        });
      }

      $(async () => {
        await Dados.carregar();
        Pesquisa.configurar();
        FiltroInativos.configurar();
        anexarEventos();
      });

    })(jQuery);
  </script>

</body>
</html>
