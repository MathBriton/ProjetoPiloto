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
Dim SQL, RS, SQL1, PIR_HELPDESK, TESTE_CONN

Dim urlPerm, sqlPerm, rsPerm, existePerm
urlPerm = "../Usuario/AcessosUsuarios"
sqlPerm = "SELECT COUNT(*) EXISTE FROM LINKS MENU " & _
          "INNER JOIN MONTAMENU PERMISSAO ON MENU.LIN_COD=PERMISSAO.LIN_COD AND ISNULL(MENU.LIN_D_E_L_E_T,'')='' " & _
          "WHERE LIN_URL='" & urlPerm & "' AND CRM_COD='" & Session("CODUSER") & "' AND FILIAL='" & Session("FILIAL") & "' " & _
          "AND ISNULL(MON_D_E_L_E_T,'')='' AND CRM_COD<>'' AND FILIAL<>''"
Set rsPerm = Conecta("PIR-HELPDESK", sqlPerm)
If Not IsObject(rsPerm) Then Response.End
existePerm = 0
If Not rsPerm.EOF Then existePerm = CInt(rsPerm("EXISTE"))
rsPerm.Close() : Set rsPerm = Nothing
If existePerm = 0 Then Response.Redirect "../login/login.asp" : Response.End

Dim IdUsuarioLogado, LoginUsuarioLogado
IdUsuarioLogado    = Session("CODIGO")
LoginUsuarioLogado = Session("USR")

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

Dim acao : acao = Trim(Request("acao"))
If acao <> "" Then
  Response.ContentType = "application/json"
  Response.CharSet = "UTF-8"

  Select Case acao
    Case "listarCategorias"
      Dim jsonCat, primeiroServ, primeiroCons
      jsonCat = "{""success"":true,""servicos"":["
      Set RS = Conecta("PIR-HELPDESK", "SELECT SUB_COD, SUB_MENU FROM SUBMENU WHERE ISNULL(SUB_D_E_L_E_T,'')='' ORDER BY SUB_MENU")
      primeiroServ = True
      While Not RS.EOF
        If Not primeiroServ Then jsonCat = jsonCat & ","
        jsonCat = jsonCat & "{""codigo"":""" & JsonEscape(RS("SUB_COD")) & """,""nome"":""" & JsonEscape(RS("SUB_MENU")) & """}"
        primeiroServ = False
        RS.MoveNext()
      Wend
      RS.Close() : Set RS = Nothing

      jsonCat = jsonCat & "],""consultas"":["
      Set RS = Conecta("PIR-HELPDESK", "SELECT CAT_COD, CAT_NOME FROM CAT_RELATORIOS ORDER BY CAT_NOME")
      primeiroCons = True
      While Not RS.EOF
        If Not primeiroCons Then jsonCat = jsonCat & ","
        jsonCat = jsonCat & "{""codigo"":""" & JsonEscape(RS("CAT_COD")) & """,""nome"":""" & JsonEscape(RS("CAT_NOME")) & """}"
        primeiroCons = False
        RS.MoveNext()
      Wend
      RS.Close() : Set RS = Nothing
      Response.Write jsonCat & "]}"

    Case "listarPermissoes"
      Dim tipo, codigoCategoria, jsonPerm, primeiro
      tipo = Trim(Request("tipo"))
      codigoCategoria = Trim(Request("codigoCategoria"))

      If tipo = "servico" Then
        SQL = "SELECT LIN_COD, LIN_DESC FROM LINKS " & _
              "WHERE RTRIM(LTRIM(SUB_COD))=RTRIM(LTRIM('" & Replace(codigoCategoria,"'","''") & "')) " & _
              "AND ISNULL(LIN_D_E_L_E_T,'')='' " & _
              "ORDER BY CONVERT(INT,LIN_COD) DESC"
      Else
        SQL = "SELECT REL_COD, REL_QUERY FROM RELATORIOSCAPA " & _
              "WHERE RTRIM(LTRIM(CAT_COD))=RTRIM(LTRIM('" & Replace(codigoCategoria,"'","''") & "')) " & _
              "AND ISNULL(REL_D_E_L_E_T_E_D,'')='' " & _
              "ORDER BY CONVERT(INT,REL_COD) DESC"
      End If

      Set RS = Conecta("PIR-HELPDESK", SQL)
      jsonPerm = "{""success"":true,""permissoes"":["
      primeiro = True
      While Not RS.EOF
        If Not primeiro Then jsonPerm = jsonPerm & ","
        If tipo = "servico" Then
          jsonPerm = jsonPerm & "{""codigo"":""" & JsonEscape(RS("LIN_COD")) & """,""nome"":""" & JsonEscape(RS("LIN_DESC")) & """}"
        Else
          jsonPerm = jsonPerm & "{""codigo"":""" & JsonEscape(RS("REL_COD")) & """,""nome"":""" & JsonEscape(RS("REL_QUERY")) & """}"
        End If
        primeiro = False
        RS.MoveNext()
      Wend
      RS.Close() : Set RS = Nothing
      Response.Write jsonPerm & "]}"

    Case "criarPermissao"
      Dim tipoCriar, categoriaCriar, nomeCriar, novoCodigo, sqlInsert, CONN_INSERT
      tipoCriar = Trim(Request("tipo"))
      categoriaCriar = Trim(Request("codigoCategoria"))
      nomeCriar = Trim(Request("nome"))

      If nomeCriar = "" Then
        Response.Write "{""success"":false,""mensagem"":""Nome da permissão não pode ser vazio.""}"
        Response.End
      End If

      On Error Resume Next
      If tipoCriar = "servico" Then
        Set RS = Conecta("PIR-HELPDESK", "SELECT ISNULL(MAX(CONVERT(INT,LIN_COD)),0)+1 AS PROXIMO FROM LINKS")
        If Not RS.EOF Then novoCodigo = CStr(RS("PROXIMO")) Else novoCodigo = "1"
        RS.Close() : Set RS = Nothing

        sqlInsert = "INSERT INTO LINKS (LIN_COD,SUB_COD,LIN_DESC,SRV_CODSERV,LIN_NOME,LIN_ORDEM,LIN_URL,LIN_EXTEN,LIN_D_E_L_E_T) " & _
                    "VALUES('" & novoCodigo & "','" & Replace(categoriaCriar,"'","''") & "','" & Replace(nomeCriar,"'","''") & "',0,'" & Replace(nomeCriar,"'","''") & "',0,'','','')"
      Else
        Set RS = Conecta("PIR-HELPDESK", "SELECT ISNULL(MAX(CONVERT(INT,REL_COD)),0)+1 AS PROXIMO FROM RELATORIOSCAPA")
        If Not RS.EOF Then novoCodigo = CStr(RS("PROXIMO")) Else novoCodigo = "1"
        RS.Close() : Set RS = Nothing

        sqlInsert = "INSERT INTO RELATORIOSCAPA (REL_COD,CAT_COD,REL_NOME,REL_QUERY,REL_URL,REL_EXTEN,REL_BANCO,REL_D_E_L_E_T_E_D,REL_ATALHO) " & _
                    "VALUES('" & novoCodigo & "','" & Replace(categoriaCriar,"'","''") & "','" & Replace(nomeCriar,"'","''") & "','" & Replace(nomeCriar,"'","''") & "','','','','','')"
      End If

      Set CONN_INSERT = Conecta("PIR-HELPDESK", sqlInsert)
      If Err.Number = 0 And IsObject(CONN_INSERT) Then
        CONN_INSERT.Close() : Set CONN_INSERT = Nothing
        Response.Write "{""success"":true,""mensagem"":""Permissão criada com sucesso."",""codigo"":""" & novoCodigo & """}"
      Else
        Response.Write "{""success"":false,""mensagem"":""Erro ao criar permissão: " & JsonEscape(Err.Description) & """}"
      End If
      On Error GoTo 0

    Case "editarPermissao"
      Dim tipoEditar, codigoEditar, novoNome, sqlUpdate, CONN_UPDATE
      tipoEditar = Trim(Request("tipo"))
      codigoEditar = Trim(Request("codigo"))
      novoNome = Trim(Request("nome"))

      If novoNome = "" Then
        Response.Write "{""success"":false,""mensagem"":""Nome da permissão não pode ser vazio.""}"
        Response.End
      End If

      On Error Resume Next
      If tipoEditar = "servico" Then
        sqlUpdate = "UPDATE LINKS SET LIN_DESC='" & Replace(novoNome,"'","''") & "',LIN_NOME='" & Replace(novoNome,"'","''") & "' WHERE LIN_COD='" & Replace(codigoEditar,"'","''") & "'"
      Else
        sqlUpdate = "UPDATE RELATORIOSCAPA SET REL_QUERY='" & Replace(novoNome,"'","''") & "',REL_NOME='" & Replace(novoNome,"'","''") & "' WHERE REL_COD='" & Replace(codigoEditar,"'","''") & "'"
      End If

      Set CONN_UPDATE = Conecta("PIR-HELPDESK", sqlUpdate)
      If Err.Number = 0 And IsObject(CONN_UPDATE) Then
        CONN_UPDATE.Close() : Set CONN_UPDATE = Nothing
        Response.Write "{""success"":true,""mensagem"":""Permissão atualizada com sucesso.""}"
      Else
        Response.Write "{""success"":false,""mensagem"":""Erro ao atualizar permissão: " & JsonEscape(Err.Description) & """}"
      End If
      On Error GoTo 0

    Case Else
      Response.Write "{""success"":false,""mensagem"":""Ação inválida.""}"
  End Select

  Response.End
End If

Function AdicionarLog(idAlvo, nomeAlvo, tipoItem, nomeItem, acaoLog)
  Dim descricao, acaoTexto, sqlLog, LOG_CONN
  If Trim(idAlvo) = "" Or Not IsNumeric(idAlvo) Then idAlvo = "NULL"

  If acaoLog = "habilitar" Then
    acaoTexto = "Habilitar " & nomeItem
    descricao = "Usuário " & LoginUsuarioLogado & " habilitou menu " & nomeItem & " para o usuário " & nomeAlvo
  Else
    acaoTexto = "Desabilitar " & nomeItem
    descricao = "Usuário " & LoginUsuarioLogado & " desabilitou menu " & nomeItem & " para o usuário " & nomeAlvo
  End If

  sqlLog = "INSERT INTO LogAtividade (IdUsuarioAcao,IdUsuarioAlvo,Rotina,Sistema,Acao,Descricao,DataCriacao) " & _
           "VALUES(" & IdUsuarioLogado & "," & idAlvo & ",'Acessos Usuários','SWNET','" & acaoTexto & "','" & descricao & "',GETDATE())"

  On Error Resume Next
  Set LOG_CONN = Conecta("PIR-CORPORATIVO", sqlLog)
  If Err.Number = 0 And IsObject(LOG_CONN) Then LOG_CONN.Close() : Set LOG_CONN = Nothing
  On Error GoTo 0
End Function

Dim listaUsuarios(), totalUsuarios, idxUsuario
totalUsuarios = 0 : idxUsuario = 0 : ReDim listaUsuarios(0)
Set PIR_HELPDESK = Conecta("PIR-HELPDESK", "SELECT CRM_R_E_C_N_O, CRM_COD, EMP_CODIGO, CRM_NOME, CRM_LOGIN FROM CRM_DTI WHERE CRM_D_E_L_E_T<>'*' ORDER BY CRM_NOME")
While Not PIR_HELPDESK.EOF
  ReDim Preserve listaUsuarios(idxUsuario)
  Set listaUsuarios(idxUsuario) = Server.CreateObject("Scripting.Dictionary")
  listaUsuarios(idxUsuario)("id")        = CStr(PIR_HELPDESK("CRM_R_E_C_N_O"))
  listaUsuarios(idxUsuario)("matricula") = Trim(PIR_HELPDESK("CRM_COD"))
  listaUsuarios(idxUsuario)("filial")    = Trim(PIR_HELPDESK("EMP_CODIGO"))
  listaUsuarios(idxUsuario)("nome")      = Trim(PIR_HELPDESK("CRM_NOME"))
  listaUsuarios(idxUsuario)("login")     = Trim(PIR_HELPDESK("CRM_LOGIN"))
  idxUsuario = idxUsuario + 1
  PIR_HELPDESK.MoveNext()
Wend
totalUsuarios = idxUsuario
PIR_HELPDESK.Close() : Set PIR_HELPDESK = Nothing

' === Usuários INATIVOS (para o modal de Importar Permissões) ===
Dim listaInativos(), totalInativos, idxInativo
totalInativos = 0 : idxInativo = 0 : ReDim listaInativos(0)
Set PIR_HELPDESK = Conecta("PIR-HELPDESK", "SELECT CRM_R_E_C_N_O, CRM_COD, EMP_CODIGO, CRM_NOME, CRM_LOGIN FROM CRM_DTI WHERE CRM_D_E_L_E_T='*' ORDER BY CRM_NOME")
While Not PIR_HELPDESK.EOF
  ReDim Preserve listaInativos(idxInativo)
  Set listaInativos(idxInativo) = Server.CreateObject("Scripting.Dictionary")
  listaInativos(idxInativo)("id")        = CStr(PIR_HELPDESK("CRM_R_E_C_N_O"))
  listaInativos(idxInativo)("matricula") = Trim(PIR_HELPDESK("CRM_COD"))
  listaInativos(idxInativo)("filial")    = Trim(PIR_HELPDESK("EMP_CODIGO"))
  listaInativos(idxInativo)("nome")      = Trim(PIR_HELPDESK("CRM_NOME"))
  listaInativos(idxInativo)("login")     = Trim(PIR_HELPDESK("CRM_LOGIN"))
  idxInativo = idxInativo + 1
  PIR_HELPDESK.MoveNext()
Wend
totalInativos = idxInativo
PIR_HELPDESK.Close() : Set PIR_HELPDESK = Nothing

Dim listaConsultas(), listaServicos(), totalConsultas, totalServicos
ReDim listaConsultas(-1) : ReDim listaServicos(-1)
totalConsultas = 0 : totalServicos = 0

Dim categoriasConsultas, categoriasServicos
Set categoriasConsultas = Server.CreateObject("Scripting.Dictionary")
Set categoriasServicos  = Server.CreateObject("Scripting.Dictionary")

Dim usuarioSelecionado
usuarioSelecionado = (Trim(Request("matricula")) <> "" And Trim(Request("filial")) <> "")

If usuarioSelecionado Then
  Dim mat : mat = Trim(Request("matricula"))
  Dim fil : fil = Trim(Request("filial"))

  Dim idxConsulta, nomeCategoria
  idxConsulta = 0 : ReDim listaConsultas(0)

  SQL = "SELECT REL.REL_COD, REL.REL_QUERY, CATEGORIA.CAT_NOME, ISNULL(ACESSO.SRV_CODSERV,0) AS SRV_CODSERV " & _
        "FROM RELATORIOSCAPA REL INNER JOIN CAT_RELATORIOS CATEGORIA ON REL.CAT_COD=CATEGORIA.CAT_COD " & _
        "LEFT JOIN (SELECT SRV_CODSERV,REL_COD FROM MONTARELAT " & _
        "           WHERE CRM_COD='" & mat & "' AND FILIAL='" & fil & "' AND ISNULL(MON_D_E_L_E_T_E_D,'')='') ACESSO ON REL.REL_COD=ACESSO.REL_COD " & _
        "WHERE ISNULL(REL_D_E_L_E_T_E_D,'')='' ORDER BY CATEGORIA.CAT_NOME, REL.REL_QUERY"

  Set PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
  While Not PIR_HELPDESK.EOF
    ReDim Preserve listaConsultas(idxConsulta)
    Set listaConsultas(idxConsulta) = Server.CreateObject("Scripting.Dictionary")
    listaConsultas(idxConsulta)("codigo")     = CStr(PIR_HELPDESK("REL_COD"))
    listaConsultas(idxConsulta)("nome")       = PIR_HELPDESK("REL_QUERY")
    listaConsultas(idxConsulta)("categoria")  = PIR_HELPDESK("CAT_NOME")
    listaConsultas(idxConsulta)("habilitado") = (PIR_HELPDESK("SRV_CODSERV") = 1)
    nomeCategoria = PIR_HELPDESK("CAT_NOME")
    If Not categoriasConsultas.Exists(nomeCategoria) Then categoriasConsultas.Add nomeCategoria, ""
    idxConsulta = idxConsulta + 1
    PIR_HELPDESK.MoveNext()
  Wend
  totalConsultas = idxConsulta
  PIR_HELPDESK.Close() : Set PIR_HELPDESK = Nothing

  Dim idxServico : idxServico = 0 : ReDim listaServicos(0)

  SQL = "SELECT SERVICO.LIN_COD, SERVICO.LIN_DESC, SUBMENU.SUB_MENU, ISNULL(ACESSO.SRV_CODSERV,0) AS SRV_CODSERV " & _
        "FROM LINKS SERVICO INNER JOIN SUBMENU ON SERVICO.SUB_COD=SUBMENU.SUB_COD " & _
        "LEFT JOIN (SELECT SRV_CODSERV,LIN_COD FROM MONTAMENU " & _
        "           WHERE CRM_COD='" & mat & "' AND FILIAL='" & fil & "' AND ISNULL(MON_D_E_L_E_T,'')='') ACESSO ON SERVICO.LIN_COD=ACESSO.LIN_COD " & _
        "WHERE ISNULL(SERVICO.LIN_D_E_L_E_T,'')='' AND ISNULL(SUBMENU.SUB_D_E_L_E_T,'')='' " & _
        "ORDER BY SUBMENU.SUB_MENU, SERVICO.LIN_DESC"

  Set PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
  While Not PIR_HELPDESK.EOF
    ReDim Preserve listaServicos(idxServico)
    Set listaServicos(idxServico) = Server.CreateObject("Scripting.Dictionary")
    listaServicos(idxServico)("codigo")     = CStr(PIR_HELPDESK("LIN_COD"))
    listaServicos(idxServico)("nome")       = PIR_HELPDESK("LIN_DESC")
    listaServicos(idxServico)("categoria")  = PIR_HELPDESK("SUB_MENU")
    listaServicos(idxServico)("habilitado") = (PIR_HELPDESK("SRV_CODSERV") = 1)
    nomeCategoria = PIR_HELPDESK("SUB_MENU")
    If Not categoriasServicos.Exists(nomeCategoria) Then categoriasServicos.Add nomeCategoria, ""
    idxServico = idxServico + 1
    PIR_HELPDESK.MoveNext()
  Wend
  totalServicos = idxServico
  PIR_HELPDESK.Close() : Set PIR_HELPDESK = Nothing
End If

If Trim(Request("envia")) = "IMPORTAR" Then
  Dim usuarioOrigem, dadosOrigem, matriculaOrigem, filialOrigem, matriculaDestino, filialDestino
  usuarioOrigem = Trim(Request("usuarioOrigem"))

  If usuarioOrigem = "" Then
    Response.Write "<script>alert('Nenhum usuário de origem selecionado!');window.history.back();</script>"
    Response.End
  End If

  dadosOrigem = Split(usuarioOrigem, "|")
  If UBound(dadosOrigem) < 1 Then
    Response.Write "<script>alert('Usuário de origem inválido!');window.history.back();</script>"
    Response.End
  End If

  matriculaOrigem  = Trim(dadosOrigem(0))
  filialOrigem     = Trim(dadosOrigem(1))
  matriculaDestino = Trim(Request("matricula"))
  filialDestino    = Trim(Request("filial"))

  SQL1 = "BEGIN TRANSACTION " & _
         "DELETE FROM MONTARELAT WHERE CRM_COD='" & matriculaDestino & "' AND FILIAL='" & filialDestino & "' " & _
         "DELETE FROM MONTAMENU  WHERE CRM_COD='" & matriculaDestino & "' AND FILIAL='" & filialDestino & "' " & _
         "INSERT INTO MONTARELAT (CRM_COD,SRV_CODSERV,REL_COD,MON_PERSONAL,MON_D_E_L_E_T_E_D,R_E_C_N_O_,CRM_TIPO,FILIAL) " & _
         "SELECT '" & matriculaDestino & "',SRV_CODSERV,REL_COD,MON_PERSONAL,MON_D_E_L_E_T_E_D," & _
         "(SELECT ISNULL(MAX(R_E_C_N_O_),0)+1 FROM MONTARELAT),CRM_TIPO,'" & filialDestino & "' " & _
         "FROM MONTARELAT WHERE CRM_COD='" & matriculaOrigem & "' AND FILIAL='" & filialOrigem & "' " & _
         "INSERT INTO MONTAMENU (CRM_COD,SRV_CODSERV,LIN_COD,MON_PERSONAL,MON_D_E_L_E_T,CRM_TIPO,FILIAL) " & _
         "SELECT '" & matriculaDestino & "',SRV_CODSERV,LIN_COD,MON_PERSONAL,MON_D_E_L_E_T,CRM_TIPO,'" & filialDestino & "' " & _
         "FROM MONTAMENU WHERE CRM_COD='" & matriculaOrigem & "' AND FILIAL='" & filialOrigem & "' " & _
         "IF @@ERROR<>0 ROLLBACK TRANSACTION ELSE COMMIT TRANSACTION"

  On Error Resume Next
  Set TESTE_CONN = Conecta("PIR-HELPDESK", SQL1)
  If Err.Number = 0 And IsObject(TESTE_CONN) Then TESTE_CONN.Close() : Set TESTE_CONN = Nothing
  On Error GoTo 0

  Response.Redirect "AcessosUsuarios.asp?matricula=" & matriculaDestino & "&filial=" & filialDestino & "&nome=" & Server.URLEncode(Request("nome"))
End If

Dim salvouComSucesso : salvouComSucesso = False
If Trim(Request("envia")) = "1" Then
  Dim nomeUsuarioAlvo, idUsuarioAlvo, i, estaHab, estavaSele
  Dim matReq : matReq = Trim(Request("matricula"))
  Dim filReq : filReq = Trim(Request("filial"))

  idUsuarioAlvo   = Trim(Request("idUsuario"))
  nomeUsuarioAlvo = Trim(Request("nome"))

  SQL1 = "BEGIN TRANSACTION " & _
         "DELETE FROM MONTARELAT WHERE CRM_COD='" & matReq & "' AND FILIAL='" & filReq & "' "

  If totalConsultas > 0 Then
    For i = 0 To totalConsultas - 1
      estaHab    = (Trim(Request("consulta" & i)) = "on")
      estavaSele = listaConsultas(i)("habilitado")

      If estaHab Then
        SQL1 = SQL1 & "INSERT INTO MONTARELAT(CRM_COD,SRV_CODSERV,REL_COD,MON_PERSONAL,MON_D_E_L_E_T_E_D,R_E_C_N_O_,CRM_TIPO,FILIAL) " & _
                      "VALUES('" & matReq & "',1," & listaConsultas(i)("codigo") & ",'S','',(SELECT MAX(R_E_C_N_O_)+1 FROM MONTARELAT WITH(NOLOCK)),1,'" & filReq & "') "
        listaConsultas(i)("habilitado") = True
        If Not estavaSele Then AdicionarLog idUsuarioAlvo, nomeUsuarioAlvo, "consulta", listaConsultas(i)("nome"), "habilitar"
      Else
        listaConsultas(i)("habilitado") = False
        If estavaSele Then AdicionarLog idUsuarioAlvo, nomeUsuarioAlvo, "consulta", listaConsultas(i)("nome"), "desabilitar"
      End If
    Next
  End If

  SQL1 = SQL1 & "DELETE FROM MONTAMENU WHERE CRM_COD='" & matReq & "' AND FILIAL='" & filReq & "' "

  If totalServicos > 0 Then
    For i = 0 To totalServicos - 1
      estaHab    = (Trim(Request("servico" & i)) = "on")
      estavaSele = listaServicos(i)("habilitado")

      If estaHab Then
        SQL1 = SQL1 & "INSERT INTO MONTAMENU(CRM_COD,SRV_CODSERV,LIN_COD,MON_PERSONAL,MON_D_E_L_E_T,CRM_TIPO,FILIAL) " & _
                      "VALUES('" & matReq & "',1," & listaServicos(i)("codigo") & ",'S','',1,'" & filReq & "') "
        listaServicos(i)("habilitado") = True
        If Not estavaSele Then AdicionarLog idUsuarioAlvo, nomeUsuarioAlvo, "serviço", listaServicos(i)("nome"), "habilitar"
      Else
        listaServicos(i)("habilitado") = False
        If estavaSele Then AdicionarLog idUsuarioAlvo, nomeUsuarioAlvo, "serviço", listaServicos(i)("nome"), "desabilitar"
      End If
    Next
  End If

  SQL1 = SQL1 & "IF @@ERROR<>0 ROLLBACK TRANSACTION ELSE COMMIT TRANSACTION"

  On Error Resume Next
  Set TESTE_CONN = Conecta("PIR-HELPDESK", SQL1)
  If Err.Number = 0 And IsObject(TESTE_CONN) Then
    TESTE_CONN.Close() : Set TESTE_CONN = Nothing
    salvouComSucesso = True
  End If
  On Error GoTo 0
End If
%>

<!doctype html>
<html lang="pt-BR">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>Gestão de Permissões de Usuários</title>

    <link rel="preconnect" href="https://cdn.jsdelivr.net" crossorigin />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css" integrity="sha384-sRIl4kxILFvY47J16cr9ZwB07vP4J8+LH7qKQnuqkuIAvNWLzeN8tE5YBujZqJLB" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" />
    <style>
      body {
        background: #f0f0f0;
        color: #212529;
        font-size: 0.8125rem;
        min-height: 100vh;
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

      .btn {
        font-size: 0.6875rem;
      }

      .page-title {
        font-size: 1.25rem;
        color: #4a4a4a;
      }

      .search-section {
        max-width: 54rem;
        margin: 1.5rem auto;
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

      .section-header {
        background: #fff;
        color: #212529;
        padding: 0.75rem 0.875rem;
        border-bottom: 1px solid #dee2e6;
        border-radius: 0.5rem 0.5rem 0 0;
        user-select: none;
        cursor: pointer;
      }

      .section-header .section-title {
        font-size: 0.9rem;
        font-weight: 600;
        color: #212529;
        margin: 0;
      }

      .section-toggle-icon {
        transition: transform 0.25s ease;
        color: #212529;
        font-size: 0.75rem;
      }

      .section-toggle-icon.collapsed {
        transform: rotate(-90deg);
      }

      .section-body {
        border: 1px solid #dee2e6;
        border-top: none;
        border-radius: 0 0 0.5rem 0.5rem;
        background: #f0f0f0;
        max-height: calc(100vh - 200px);
        overflow-y: auto;
        padding: 0.75rem;
      }

      .section-body::-webkit-scrollbar {
        width: 5px;
      }

      .section-body::-webkit-scrollbar-track {
        background: #f0f0f0;
        border-radius: 3px;
      }

      .section-body::-webkit-scrollbar-thumb {
        background: #c0c0c0;
        border-radius: 3px;
      }

      .section-body::-webkit-scrollbar-thumb:hover {
        background: #a0a0a0;
      }

      .sk-placeholder {
        padding: 2.5rem 1rem;
        text-align: center;
        color: #6c757d;
      }

      .sk-placeholder i {
        font-size: 2.5rem;
        display: block;
        margin-bottom: 0.5rem;
      }

      .sk-placeholder p {
        font-size: 0.75rem;
        margin: 0;
      }

      .sk-placeholder-loading {
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 0.5rem;
        padding: 2.5rem 1rem;
        color: #6c757d;
      }

      .sk-placeholder-loading .spinner-border {
        width: 2rem;
        height: 2rem;
        color: #4a4a4a;
      }

      .sk-placeholder-loading p {
        font-size: 0.75rem;
        margin: 0;
      }

      .user-info-bar {
        border-left: 0.25rem solid #6c757d;
      }

      .user-avatar {
        width: 2.25rem;
        height: 2.25rem;
        background: #6c757d;
        border-radius: 50%;
        color: #fff;
        font-size: 1rem;
        flex-shrink: 0;
        display: flex;
        align-items: center;
        justify-content: center;
      }

      .user-details .user-login-display {
        font-size: 0.72rem;
        color: #6c757d;
        margin: 0;
      }

      .category-card {
        border: 1px solid #dee2e6;
        border-radius: 4px;
        overflow: hidden;
        margin-bottom: 0.625rem;
        transition: box-shadow 0.2s ease;
        background: #fff;
      }

      .category-card:hover {
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
      }

      .category-header {
        cursor: pointer;
        user-select: none;
        background: #fff;
        color: #212529;
        padding: 0.5rem 0.75rem;
        border-bottom: 1px solid #dee2e6;
        transition: background 0.15s ease;
      }

      .category-header:hover {
        background: #f8f9fa;
      }

      .category-title {
        font-size: 0.75rem;
        font-weight: 600;
        color: #212529;
      }

      .toggle-icon {
        transition: transform 0.25s ease;
        color: #4a4a4a;
        font-size: 0.6rem;
      }

      .count-badge {
        font-size: 0.5625rem !important;
        padding: 0.15rem 0.4rem !important;
        border-radius: 0.75rem !important;
        background: #e9ecef !important;
        color: #495057 !important;
      }

      .items-body {
        background: #fff;
      }

      .permission-item {
        padding: 0.4rem 0.75rem;
        border-bottom: 1px solid #e8e8e8;
        display: flex;
        align-items: center;
        gap: 0.5rem;
        background: #fff;
        transition: all 0.15s ease;
        cursor: pointer;
      }

      .permission-item:last-child {
        border-bottom: none;
      }

      .permission-item:hover {
        background: rgba(0, 0, 0, 0.03);
        border-color: #dee2e6;
        transform: translateX(2px);
      }

      .permission-item.is-checked {
        background: rgba(0, 0, 0, 0.03);
      }

      .permission-item label {
        margin: 0;
        cursor: pointer;
        flex: 1;
        font-size: 0.6875rem;
        color: #212529;
        line-height: 1.35;
      }

      .permission-item.is-checked label {
        font-weight: 400;
      }

      input[type="checkbox"].perm-check,
      input[type="checkbox"].category-checkbox {
        width: 1rem;
        height: 1rem;
        margin: 0;
        cursor: pointer;
        accent-color: #0d47a1;
      }

      .search-input-compact {
        max-width: 18rem;
      }

      .search-input-compact .form-control {
        font-size: 0.6875rem;
        padding: 0.2rem 0.5rem;
      }

      .search-input-compact .input-group-text {
        font-size: 0.625rem;
        padding: 0.2rem 0.5rem;
      }

      .search-results-inline small {
        font-size: 0.625rem;
        color: #6c757d;
        white-space: nowrap;
      }

      .modal-content {
        border: 1px solid #dee2e6;
        border-radius: 6px;
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.2);
      }

      .modal-header-sk {
        background: #fff;
        color: #212529;
        padding: 0.65rem 1rem;
        border-bottom: 1px solid #dee2e6;
        border-radius: 6px 6px 0 0;
      }

      .modal-header-sk .modal-title {
        font-size: 0.8125rem;
        font-weight: 600;
        color: #212529;
      }

      .modal-footer-sk {
        background: #f8f9fa;
        border-top: 1px solid #dee2e6;
        padding: 0.6rem 1rem;
      }

      .modal-importar {
        max-width: 640px;
        width: 100%;
      }

      .modal-importar .modal-body {
        min-height: 260px;
        display: flex;
        flex-direction: column;
      }

      .modal-importar .alert-smoking {
        flex-grow: 1;
        margin-bottom: 0;
        display: flex;
        flex-direction: column;
        justify-content: center;
        text-align: center;
      }

      #switchIncluirInativos {
        transform: scale(1.4);
        transform-origin: left center;
        margin-right: 0.5rem;
      }

      .modal-footer-sk .btn {
        font-size: 0.6875rem;
        padding: 0.3rem 0.75rem;
      }

      .alert-smoking {
        background: #fffbeb;
        border: 1px solid #fde68a;
        border-radius: 4px;
        padding: 0.625rem 0.75rem;
        font-size: 0.75rem;
      }

      .alert-smoking-title {
        font-size: 0.75rem;
        font-weight: 700;
        color: #78350f;
      }

      .alert-smoking p {
        color: #92400e;
        font-size: 0.72rem;
      }

      .overlay-loader {
        position: fixed;
        inset: 0;
        background: rgba(255, 255, 255, 0.75);
        z-index: 9999;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        gap: 0.75rem;
      }

      .overlay-loader-spinner {
        width: 2.5rem;
        height: 2.5rem;
        border: 3px solid #e0e0e0;
        border-top-color: #4a4a4a;
        border-radius: 50%;
        animation: spin 0.75s linear infinite;
      }

      @keyframes spin {
        to { transform: rotate(360deg); }
      }

      .overlay-loader-label {
        font-size: 0.75rem;
        color: #6c757d;
        letter-spacing: 0.02em;
      }

      .toast-wrapper {
        position: fixed;
        top: 1rem;
        right: 1rem;
        z-index: 11000;
        display: flex;
        flex-direction: column;
        gap: 0.5rem;
        min-width: 300px;
        max-width: 380px;
        pointer-events: none;
      }

      .toast-item {
        background: #fff;
        border-radius: 6px;
        overflow: hidden;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.12);
        border: 1px solid #e8e8e8;
        pointer-events: all;
        animation: entrar-toast 0.2s ease;
      }

      .toast-item-success {
        border-left: 4px solid #198754;
      }

      .toast-item-warning {
        border-left: 4px solid #ffc107;
      }

      .toast-item-danger {
        border-left: 4px solid #dc3545;
      }

      .toast-item-info {
        border-left: 4px solid #6c757d;
      }

      @keyframes entrar-toast {
        from {
          opacity: 0;
          transform: translateX(20px);
        }

        to {
          opacity: 1;
          transform: translateX(0);
        }
      }

      .toast-corpo {
        display: flex;
        align-items: center;
        gap: 0.6rem;
        padding: 0.65rem 0.85rem;
        font-size: 0.85rem;
        color: #333;
      }

      .toast-fechar {
        margin-left: auto;
        background: none;
        border: none;
        font-size: 1rem;
        color: #999;
        cursor: pointer;
        padding: 0 0 0 0.5rem;
      }

      .toast-fechar:hover {
        color: #333;
      }

      .toast-barra {
        height: 3px;
        animation: encolher linear forwards;
      }

      .toast-barra-success {
        background: #198754;
      }

      .toast-barra-warning {
        background: #ffc107;
      }

      .toast-barra-danger {
        background: #dc3545;
      }

      .toast-barra-info {
        background: #6c757d;
      }

      @keyframes encolher {
        from {
          width: 100%;
        }

        to {
          width: 0%;
        }
      }

      .filial-badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 3px;
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.2px;
      }

      .filial-00 {
        background: #dbeafe;
        color: #1e40af;
      }

      .filial-01 {
        background: #fce7f3;
        color: #9d174d;
      }

      .filial-02 {
        background: #d1fae5;
        color: #065f46;
      }

      .filial-03 {
        background: #ede9fe;
        color: #4c1d95;
      }

      .filial-04 {
        background: #fef3c7;
        color: #78350f;
      }

      .filial-05 {
        background: #cffafe;
        color: #155e75;
      }

      .filial-06 {
        background: #fee2e2;
        color: #7f1d1d;
      }

      .filial-07 {
        background: #d1fae5;
        color: #064e3b;
      }

      .filial-08 {
        background: #ffedd5;
        color: #7c2d12;
      }

      .filial-09 {
        background: #e0e7ff;
        color: #1e1b4b;
      }

      .filial-10 {
        background: #fef9c3;
        color: #713f12;
      }

      .filial-11 {
        background: #f0fdf4;
        color: #14532d;
      }

      .filial-12 {
        background: #fdf4ff;
        color: #581c87;
      }

      .filial-13 {
        background: #ecfdf5;
        color: #065f46;
      }

      .sk-search-wrapper {
        position: relative;
      }

      .sk-dropdown {
        position: absolute;
        top: 100%;
        left: 0;
        right: 0;
        max-height: 280px;
        overflow-y: auto;
        background: #fff;
        border: 1px solid #dee2e6;
        border-top: none;
        border-radius: 0 0 6px 6px;
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.12);
        z-index: 1050;
        display: none;
      }

      .sk-dropdown.show {
        display: block;
      }

      .sk-dropdown-item {
        padding: 9px 14px;
        cursor: pointer;
        border-bottom: 1px solid #f5f5f5;
        transition: background 0.15s;
      }

      .sk-dropdown-item:last-child {
        border-bottom: none;
      }

      .sk-dropdown-item:hover,
      .sk-dropdown-item.active {
        background: #f0f0f0;
      }

      .sk-dropdown-login {
        font-size: 0.78rem;
      }

      .sk-dropdown-counter {
        padding: 5px 14px;
        font-size: 0.72rem;
        background: #fafafa;
        border-top: 1px solid #dee2e6;
        color: #888;
        text-align: right;
        position: sticky;
        bottom: 0;
      }

      @media (max-width: 992px) {
        .search-section {
          max-width: 100%;
          padding: 0 0.75rem;
        }

        .search-input-compact {
          max-width: 100%;
          flex-grow: 1;
        }
      }

      @media (max-width: 768px) {
        .search-section {
          margin: 1rem auto;
        }
      }

      @media (max-width: 1440px) {
        .section-body {
          max-height: calc(100vh - 180px);
        }
      }

      @media (min-width: 1920px) {
        .section-body {
          max-height: calc(90vh - 220px);
        }
      }

      @media print {
        .btn,
        .search-bar-wrapper {
          display: none;
        }

        .card {
          box-shadow: none;
          border: 1px solid #ccc;
          page-break-inside: avoid;
        }

        .collapse {
          display: block;
          height: auto;
        }
      }
    </style>
  </head>

  <body class="d-flex flex-column min-vh-100">
    <div class="toast-wrapper" id="toastWrapper"></div>

    <div id="loader" class="overlay-loader d-none">
      <div class="overlay-loader-spinner"></div>
      <span class="overlay-loader-label">Carregando...</span>
    </div>

    <input type="hidden" id="idUsuarioLogado" value="<%=IdUsuarioLogado%>" />
    <input type="hidden" id="loginUsuarioLogado" value="<%=LoginUsuarioLogado%>" />

    <div class="container-fluid px-3 my-3">
      <form action="AcessosUsuarios.asp" method="post" id="AcessosUsuario" autocomplete="off">
        <input type="hidden" name="envia" id="envia" value="0" />
        <input type="hidden" name="importar" id="importar" value="0" />
        <input type="hidden" name="idUsuario" id="idUsuario" value="<%=Trim(Request("idUsuario"))%>" />
        <input type="hidden" name="usuarioOrigem" id="usuarioOrigemHidden" value="" />
        <input type="hidden" name="matricula" id="matricula" value="<%=Trim(Request("matricula"))%>" />
        <input type="hidden" name="filial" id="filial" value="<%=Trim(Request("filial"))%>" />
        <input type="hidden" name="nome" id="nome" value="<%=Trim(Request("nome"))%>" />
        <input type="hidden" name="crmLogin" id="crmLogin" value="<%=Trim(Request("crmLogin"))%>" />
        <input type="hidden" name="contaServico" id="contaServico" value="0" />
        <input type="hidden" name="contaConsulta" id="contaConsulta" value="0" />

        <% If usuarioSelecionado Then %>

        <%
          Sub RenderizarBloco(tipoBloco, tituloBloco, iconeBloco, listaItens, totalItens, categoriasDict)
            Dim iconId : iconId = "toggleIcon" & UCase(Left(tipoBloco,1)) & Mid(tipoBloco,2)
            Dim contId : If tipoBloco = "consulta" Then contId = "consultasContainer" Else contId = "servicosContainer"
        %>
        <div class="col-lg-6">
          <section class="card shadow-sm h-100">
            <div
              class="section-header d-flex align-items-center gap-2"
              role="button"
              tabindex="0"
              onclick="toggleSessao('<%=tipoBloco%>')"
              onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();toggleSessao('<%=tipoBloco%>');}"
              aria-controls="<%=contId%>"
              aria-expanded="true"
            >
              <i class="bi bi-chevron-down section-toggle-icon" id="<%=iconId%>" aria-hidden="true"></i>
              <i class="<%=iconeBloco%>" aria-hidden="true"></i>
              <span class="section-title"><%=tituloBloco%></span>
            </div>

            <div class="section-body" id="<%=contId%>">
              <%
                Dim catKeys, catIdx, nomeCat
                catKeys = categoriasDict.Keys
                For catIdx = 0 To categoriasDict.Count - 1
                  nomeCat = catKeys(catIdx)
              %>

              <article class="category-card mb-2" data-tipo="<%=tipoBloco%>" data-categoria="<%=nomeCat%>">
                <div
                  class="category-header d-flex justify-content-between align-items-center"
                  role="button"
                  tabindex="0"
                  data-bs-toggle="collapse"
                  data-bs-target="#collapse-<%=tipoBloco%>-<%=catIdx%>"
                  aria-controls="collapse-<%=tipoBloco%>-<%=catIdx%>"
                  aria-expanded="true"
                  onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();this.click();}"
                >
                  <div class="d-flex align-items-center gap-2">
                    <input
                      type="checkbox"
                      class="form-check-input category-checkbox"
                      data-tipo="<%=tipoBloco%>"
                      data-categoria="<%=nomeCat%>"
                      onclick="event.stopPropagation();"
                      aria-label="Selecionar categoria <%=nomeCat%>"
                    />
                    <span class="category-title"><%=nomeCat%></span>
                    <span class="badge count-badge" aria-label="Itens marcados/total">0/0</span>
                  </div>
                  <i class="bi bi-chevron-down toggle-icon" aria-hidden="true"></i>
                </div>

                <div class="collapse show" id="collapse-<%=tipoBloco%>-<%=catIdx%>">
                  <div class="items-body">
                    <%
                      Dim itemIdx
                      For itemIdx = 0 To totalItens - 1
                        If listaItens(itemIdx)("categoria") = nomeCat Then
                    %>
                    <div
                      class="permission-item"
                      data-nome="<%=listaItens(itemIdx)("nome")%>"
                      data-categoria="<%=nomeCat%>"
                    >
                      <%Dim chkAttr : chkAttr = "" : If listaItens(itemIdx)("habilitado") Then chkAttr = "checked" End If%>
                      <input
                        type="checkbox"
                        class="form-check-input perm-check"
                        name="<%=tipoBloco%><%=itemIdx%>"
                        id="<%=tipoBloco%><%=itemIdx%>"
                        data-tipo="<%=tipoBloco%>"
                        data-categoria="<%=nomeCat%>"
                        <%=chkAttr%>
                      />
                      <label for="<%=tipoBloco%><%=itemIdx%>"><%=listaItens(itemIdx)("nome")%></label>
                    </div>
                    <%
                        End If
                      Next
                    %>
                  </div>
                </div>
              </article>

              <%
                Next
              %>
            </div>
          </section>
        </div>
        <% End Sub %>

        <div class="d-flex align-items-center justify-content-between mb-3 px-1">
          <h4 class="fw-bold mb-0 page-title">Acessos Usuários</h4>
        </div>

        <section id="permissionsSection" aria-label="Permissões do usuário">
          <div class="user-info-bar card shadow-sm mb-3">
            <div class="card-body py-2 px-3 d-flex align-items-center flex-wrap gap-2">
              <div class="user-avatar" aria-hidden="true">
                <i class="bi bi-person-fill"></i>
              </div>

              <div class="user-details">
                <h1 id="userNameDisplay" class="mb-0 h6"><%=Trim(Request("nome"))%></h1>
                <p class="mt-1 user-login-display mb-0">
                  <i class="bi bi-person-badge me-1" aria-hidden="true"></i>
                  <span id="userLoginDisplay"><%=Trim(Request("crmLogin"))%></span>
                   &bull; 
                  <span id="userFilialDisplay"></span>
                </p>
              </div>

              <div class="d-flex gap-2 flex-wrap ms-auto align-items-center">
                <button type="button" class="btn btn-sm btn-success" onclick="Salvar()" aria-label="Salvar permissões" title="Salvar">
                  <i class="bi bi-floppy" aria-hidden="true"></i>
                </button>
                <button type="button" class="btn btn-sm btn-primary" onclick="VoltarFormBusca()" aria-label="Voltar para busca de usuário" title="Buscar">
                  <i class="bi bi-search" aria-hidden="true"></i>
                </button>
                <button type="button" class="btn btn-sm btn-secondary" onclick="abrirModalImportar()" aria-label="Copiar permissões de outro usuário" title="Importar Permissões">
                  <i class="bi bi-copy" aria-hidden="true"></i>
                </button>
                <button type="button" class="btn btn-sm btn-danger" onclick="gerarPDF()" aria-label="Gerar PDF" title="Baixar PDF">
                  <i class="bi bi-file-pdf" aria-hidden="true"></i>
                </button>

                <div class="search-bar-wrapper ms-1">
                  <div class="d-flex align-items-center gap-2">
                    <div id="searchResults" class="search-results-inline" style="display: none">
                      <small><span id="searchCount">0</span> resultados</small>
                    </div>

                    <div class="input-group input-group-sm search-input-compact">
                      <span class="input-group-text"><i class="bi bi-search" aria-hidden="true"></i></span>
                      <input
                        type="text"
                        class="form-control"
                        id="searchPermissions"
                        placeholder="Pesquisar permissões"
                        autocomplete="off"
                        inputmode="search"
                      />
                      <button class="btn btn-outline-secondary" type="button" id="clearSearch" aria-label="Limpar pesquisa">
                        <i class="bi bi-x-lg" aria-hidden="true"></i>
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div class="row g-3">
            <% Call RenderizarBloco("servico",  "Serviços",  "bi bi-gear",      listaServicos,  totalServicos,  categoriasServicos)  %>
            <% Call RenderizarBloco("consulta", "Consultas", "bi bi-bar-chart", listaConsultas, totalConsultas, categoriasConsultas) %>
          </div>
        </section>

        <script>
          document.getElementById("contaServico").value = <%=totalServicos%>;
          document.getElementById("contaConsulta").value = <%=totalConsultas%>;
        </script>

        <% Else %>

        <section class="search-section mx-auto" aria-label="Selecionar usuário">
          <div class="card shadow">
            <div class="panel-header">Controle de Acessos</div>

            <div class="card-body pt-4 pb-3 px-4">
              <div class="row g-4 justify-content-center">
                <div class="col-lg-6">
                  <label class="form-label" for="searchUsuario">
                    <i class="bi bi-person me-1" aria-hidden="true"></i>Usuário
                  </label>
                  <input type="hidden" id="ListaUsuario" name="ListaUsuario" />
                  <div class="sk-search-wrapper">
                    <div class="input-group">
                      <span class="input-group-text bg-white border-end-0">
                        <i class="bi bi-search text-muted" aria-hidden="true"></i>
                      </span>
                      <input
                        type="text"
                        id="searchUsuario"
                        class="form-control search-control border-start-0 ps-0"
                        placeholder="Digite o nome ou login do usuário..."
                        autocomplete="off"
                        aria-label="Buscar usuário"
                        aria-autocomplete="list"
                        aria-controls="dropdownUsuario"
                        aria-expanded="false"
                      />
                    </div>
                    <div class="sk-dropdown" id="dropdownUsuario" role="listbox"></div>
                  </div>
                </div>

                <div class="col-lg-6">
                  <label class="form-label" for="filialDescricao">
                    <i class="bi bi-building me-1" aria-hidden="true"></i>Filial
                  </label>
                  <input
                    type="text"
                    id="filialDescricao"
                    class="form-control search-control"
                    placeholder="Filial"
                    readonly
                  />
                </div>
              </div>

              <div class="row justify-content-center mt-3">
                <div class="col-lg-10">
                  <div id="skPlaceholder" class="sk-placeholder">
                    <i class="bi bi-person-circle" aria-hidden="true"></i>
                    <p>Selecione um usuário acima para visualizar os dados e gerenciar a conta.</p>
                  </div>
                  <div id="skPlaceholderLoading" class="sk-placeholder-loading d-none">
                    <div class="spinner-border" role="status" aria-label="Carregando..."></div>
                    <p>Carregando dados do usuário...</p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </section>

        <% End If %>
      </form>
    </div>

    <div class="modal fade" id="modalImportar" tabindex="-1" aria-labelledby="modalImportarLabel">
      <div class="modal-dialog modal-dialog-centered modal-importar">
        <div class="modal-content">
          <div class="modal-header-sk position-relative">
            <h5 class="modal-title d-flex align-items-center gap-2 mb-0" id="modalImportarLabel">
              <i class="bi bi-copy"></i> Importar Permissões
            </h5>
            <button type="button" class="btn-close position-absolute top-50 end-0 translate-middle-y me-3" data-bs-dismiss="modal" aria-label="Fechar"></button>
          </div>

          <div class="modal-body">
            <div class="mb-3">
              <div class="d-flex align-items-center justify-content-between mb-2">
                <label class="form-label mb-0" for="usuarioOrigemSelect">
                  <i class="bi bi-person me-1" aria-hidden="true"></i>Usuário de origem
                </label>
                <div class="form-check form-switch mb-0 d-flex align-items-center gap-2">
                  <input
                    class="form-check-input"
                    type="checkbox"
                    role="switch"
                    id="switchIncluirInativos"
                    style="cursor:pointer;"
                  />
                  <label class="form-check-label text-muted" for="switchIncluirInativos" style="font-size:0.75rem;white-space:nowrap;cursor:pointer;">
                    <i class="bi bi-person-slash me-1" aria-hidden="true"></i>Incluir inativos
                  </label>
                </div>
              </div>

              <select id="usuarioOrigemSelect" class="form-select" style="width:100%;font-size:0.75rem;" data-placeholder="Nome ou login...">
                <option value="">Selecione um usuário...</option>
                <% Dim z2 : For z2 = 0 To totalUsuarios - 1 %>
                <option
                  value="<%=listaUsuarios(z2)("matricula")%>|<%=listaUsuarios(z2)("filial")%>"
                  data-login="<%=listaUsuarios(z2)("login")%>"
                  data-ativo="1"
                ><%=listaUsuarios(z2)("nome")%> — <%=listaUsuarios(z2)("login")%></option>
                <% Next %>

                <% Dim z3 : For z3 = 0 To totalInativos - 1 %>
                <option
                  value="<%=listaInativos(z3)("matricula")%>|<%=listaInativos(z3)("filial")%>"
                  data-login="<%=listaInativos(z3)("login")%>"
                  data-ativo="0"
                  class="usuario-inativo"
                  style="display:none;"
                >⚠ <%=listaInativos(z3)("nome")%> — <%=listaInativos(z3)("login")%> (inativo)</option>
                <% Next %>
              </select>
            </div>

            <div class="alert-smoking" role="alert">
              <span class="alert-smoking-title">Atenção</span>
              <p class="mb-0">Todas as permissões atuais serão substituídas pelas do <strong>Usuário</strong> selecionado.</p>
            </div>
          </div>

          <div class="modal-footer-sk d-flex justify-content-end">
            <button type="button" class="btn btn-outline-secondary" data-bs-dismiss="modal">Cancelar</button>
            <button type="button" class="btn btn-success ms-2" onclick="ConfirmarImportacao()">Confirmar</button>
          </div>
        </div>
      </div>
    </div>

    <div class="modal fade" id="modalConfirmacao" tabindex="-1">
      <div class="modal-dialog modal-dialog-centered modal-sm">
        <div class="modal-content">
          <div class="modal-header-sk position-relative">
            <h5 class="modal-title d-flex align-items-center gap-2 mb-0">
              <i class="bi bi-exclamation-circle"></i>
              <span id="modalConfirmacaoTitulo">Confirmação</span>
            </h5>
            <button type="button" class="btn-close position-absolute top-50 end-0 translate-middle-y me-3" data-bs-dismiss="modal" aria-label="Fechar"></button>
          </div>

          <div class="modal-body">
            <p class="mb-0" id="modalConfirmacaoMsg"></p>
          </div>

          <div class="modal-footer-sk d-flex justify-content-end">
            <button type="button" class="btn btn-outline-secondary" data-bs-dismiss="modal">Cancelar</button>
            <button type="button" class="btn btn-success ms-2" id="modalConfirmacaoBtn">Confirmar</button>
          </div>
        </div>
      </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/jquery@4.0.0/dist/jquery.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js" integrity="sha384-FKyoEForCGlyvwx9Hj09JcYn3nv7wiPVlz7YYwJrWVcXK/BmnVDxM+D2scQbITxI" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"></script>

    <script>
      const idUsuarioLogado = "<%=IdUsuarioLogado%>";
      const loginUsuarioLogado = "<%=LoginUsuarioLogado%>";
      const caminhoLogo = "/corporativo/usuario/Logo/logo-smurfit-black.png";
    </script>

    <script>
      (($) => {
        "use strict";

        const TOAST_CFG = Object.freeze({ delay: 4500, borda: true });

        const TOAST_ICON = Object.freeze({
          success: { cls: "bi-check-circle-fill", cor: "#198754" },
          warning: { cls: "bi-exclamation-triangle-fill", cor: "#ffc107" },
          danger: { cls: "bi-x-circle-fill", cor: "#dc3545" },
          info: { cls: "bi-info-circle-fill", cor: "#6c757d" },
        });

        const AppState = {
  usuarioSelecionado: false,
  estadoCheckboxes: { consultas: {}, servicos: {} },
  categorias: { servicos: {}, consultas: {} },
};

$(document).ready(() => {
  inicializarAplicacao();
  configurarEventListeners();
  configurarPesquisa();
  configurarAtalhosTeclado();
  configurarSelectUsuario();
});

function debounce(fn, wait = 300) {
  let t;
  return function (...args) {
    clearTimeout(t);
    t = setTimeout(() => fn.apply(this, args), wait);
  };
}

function inicializarAplicacao() {
  const filial = $("#filial").val();
  if (filial) $("#userFilialDisplay").html(renderFilialBadge(filial));
  if ($("#permissionsSection").length) {
    inicializarCategorias();
    sincronizarTudo();
  }
}

function sincronizarTudo() {
  $(".permission-item").each(function () {
    $(this).toggleClass(
      "is-checked",
      $(this).find(".perm-check").is(":checked"),
    );
  });
  $(".category-card").each(function () {
    atualizarContagemCategoria($(this));
    const $first = $(this).find(".perm-check").first();
    if ($first.length) atualizarCheckboxCategoria($first[0]);
  });
}

function configurarEventListeners() {
  if (!$("#permissionsSection").length) return;

  $(".category-checkbox").on("change", function () {
    toggleCategoria(this);
  });

  $(document).on("change", ".permission-item .perm-check", function () {
    $(this)
      .closest(".permission-item")
      .toggleClass("is-checked", $(this).is(":checked"));
    atualizarCheckboxCategoria(this);
    atualizarContadores();
  });

  $(document).on("click", ".permission-item", function (e) {
    if ($(e.target).is(".perm-check, label")) return;
    const $cb = $(this).find(".perm-check");
    $cb.prop("checked", !$cb.prop("checked")).trigger("change");
  });

  $(".category-header").on("click", function (e) {
    if (!$(e.target).is(".category-checkbox, .form-check-input")) {
      $(this).find(".toggle-icon").toggleClass("bi-chevron-down bi-chevron-right");
    }
  });

  $(".collapse")
    .on("shown.bs.collapse", function () {
      $(this).closest(".category-card").find(".toggle-icon").removeClass("bi-chevron-right").addClass("bi-chevron-down");
    })
    .on("hidden.bs.collapse", function () {
      $(this).closest(".category-card").find(".toggle-icon").removeClass("bi-chevron-down").addClass("bi-chevron-right");
    });
}

const entrarUsuarioDebounced = debounce(() => {
  if (!$("#filial").val() || !$("#matricula").val()) return;
  mostrarLoadingInline(true);
  $("#AcessosUsuario").submit();
}, 350);

const listaUsuariosJS = [
  <% Dim zJS : For zJS = 0 To totalUsuarios - 1 %>
  {
    id: "<%=listaUsuarios(zJS)("id")%>",
    matricula: "<%=listaUsuarios(zJS)("matricula")%>",
    filial: "<%=listaUsuarios(zJS)("filial")%>",
    nome: "<%=Replace(listaUsuarios(zJS)("nome"), """", "\""")%>",
    login: "<%=Replace(listaUsuarios(zJS)("login"), """", "\""")%>"
  }<% If zJS < totalUsuarios - 1 Then %>,<% End If %>
  <% Next %>
];

const LIMITE_DROPDOWN = 50;
let skActiveIndex = -1;
let skDebounceTimer = null;

function configurarSelectUsuario() {
  const $input = $("#searchUsuario");
  const $dropdown = $("#dropdownUsuario");

  if (!$input.length) return;

  $input.on("input", function () {
    const termo = $(this).val().trim();
    skActiveIndex = -1;
    clearTimeout(skDebounceTimer);
    skDebounceTimer = setTimeout(function () {
      skFiltrarUsuarios(termo);
    }, 200);
  });

  $input.on("focus", function () {
    skFiltrarUsuarios($(this).val().trim());
  });

  $input.on("keydown", function (e) {
    const $items = $(".sk-dropdown-item:visible", $dropdown);
    if (!$items.length) return;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      skActiveIndex = Math.min(skActiveIndex + 1, $items.length - 1);
      skAtualizarItemAtivo($items);
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      skActiveIndex = Math.max(skActiveIndex - 1, 0);
      skAtualizarItemAtivo($items);
    } else if (e.key === "Enter") {
      e.preventDefault();
      if (skActiveIndex >= 0) $items.eq(skActiveIndex).trigger("click");
    } else if (e.key === "Escape") {
      skFecharDropdown();
    }
  });

  $(document).on("click", function (e) {
    if (!$(e.target).closest(".sk-search-wrapper").length) {
      skFecharDropdown();
    }
  });
}

function skFiltrarUsuarios(termo) {
  const $dropdown = $("#dropdownUsuario");
  const termoLower = termo.toLowerCase();
  $dropdown.empty();

  const filtrados = listaUsuariosJS.filter(function (u) {
    return (
      termo === "" ||
      u.nome.toLowerCase().indexOf(termoLower) > -1 ||
      u.login.toLowerCase().indexOf(termoLower) > -1
    );
  });

  if (filtrados.length === 0) {
    $dropdown
      .html('<div class="p-3 text-center text-muted small"><i class="bi bi-emoji-frown me-1"></i>Nenhum usuário encontrado</div>')
      .addClass("show");
    return;
  }

  const exibidos = filtrados.slice(0, LIMITE_DROPDOWN);
  const limitado = filtrados.length > LIMITE_DROPDOWN;

  exibidos.forEach(function (usr) {
    const $item = $(
      '<div class="sk-dropdown-item" role="option">' +
        '<div class="d-flex justify-content-between align-items-center"><div>' +
          '<div class="fw-semibold small">' + usr.nome + "</div>" +
          '<div class="text-muted sk-dropdown-login"><i class="bi bi-person-badge me-1"></i>' + usr.login + "</div>" +
        "</div></div></div>"
    )
      .data("usuario", usr)
      .on("click", function () {
        skSelecionarUsuario(usr);
      });

    $dropdown.append($item);
  });

  if (limitado) {
    $dropdown.append(
      '<div class="sk-dropdown-counter">Mostrando ' +
        LIMITE_DROPDOWN + " de " + filtrados.length +
        "</div>"
    );
  }

  $dropdown.addClass("show");
}

function skAtualizarItemAtivo($items) {
  $items.removeClass("active");
  if (skActiveIndex < 0) return;
  const $ativo = $items.eq(skActiveIndex).addClass("active");
  const container = document.getElementById("dropdownUsuario");
  const item = $ativo[0];
  if (item.offsetTop < container.scrollTop) {
    container.scrollTop = item.offsetTop;
  } else if (item.offsetTop + item.offsetHeight > container.scrollTop + container.clientHeight) {
    container.scrollTop = item.offsetTop + item.offsetHeight - container.clientHeight;
  }
}

function skFecharDropdown() {
  $("#dropdownUsuario").removeClass("show");
  skActiveIndex = -1;
}

function skSelecionarUsuario(usr) {
  $("#searchUsuario").val(usr.nome);
  $("#ListaUsuario").val(usr.id);
  skFecharDropdown();

  $("#idUsuario").val(usr.id);
  $("#matricula").val(usr.matricula);
  $("#nome").val(usr.nome);
  $("#filial").val(usr.filial);
  $("#crmLogin").val(usr.login);
  $("#filialDescricao").val(ObterDescricaoFilial(usr.filial));
  entrarUsuarioDebounced();
}

function SelecionouUsuario(selectElement) {

}

function select2Matcher(params, data) {
  // Ocultar inativos enquanto o switch estiver desligado
  if ($(data.element).data("ativo") === 0 && !$("#switchIncluirInativos").is(":checked")) {
    return null;
  }
  const termo = (params.term || "").trim().toLowerCase();
  if (!termo) return data;
  const texto = (data.text || "").toLowerCase();
  const login = String($(data.element).data("login") || "").toLowerCase();
  return texto.includes(termo) || login.includes(termo) ? data : null;
}

function limparCamposUsuario() {
  $(
    "#idUsuario, #matricula, #nome, #filial, #crmLogin, #filialDescricao, #loginDescricao",
  ).val("");
}

function Buscar() {
  if (!$("#filial").val() || !$("#matricula").val()) {
    mostrarNotificacao("Por favor, selecione um usuário.", "warning");
    return;
  }
  mostrarLoading(true);
  $("#AcessosUsuario").submit();
}

function VoltarFormBusca() {
  mostrarModalConfirmacao(
    "Retornar",
    "Alterações não salvas serão perdidas.",
    () => {
      limparCamposUsuario();
      $("#envia").val("0");
      $("#contaServico, #contaConsulta").val(0);
      AppState.estadoCheckboxes = { consultas: {}, servicos: {} };
      mostrarLoading(true);
      $("#AcessosUsuario").submit();
    },
  );
}

function configurarPesquisa() {
  const $input = $("#searchPermissions");
  const $clear = $("#clearSearch");
  if (!$input.length) return;

  $input.on(
    "input",
    debounce(function () {
      const termo = ($(this).val() || "").trim().toLowerCase();
      termo ? realizarPesquisa(termo) : limparPesquisa();
    }, 250),
  );

  $clear.on("click", () => {
    $input.val("");
    limparPesquisa();
    $input.focus();
  });
  $input.on("keydown", (e) => {
    if (e.key === "Enter") e.preventDefault();
  });
}

function realizarPesquisa(termo) {
  let count = 0;
  $(".permission-item").hide();
  $(".category-card").hide();
  $(".permission-item").each(function () {
    const nome = ($(this).data("nome") || "").toLowerCase();
    const cat = ($(this).data("categoria") || "").toLowerCase();
    if (nome.includes(termo) || cat.includes(termo)) {
      count++;
      $(this).show().closest(".category-card").show();
      const $col = $(this).closest(".collapse");
      if ($col.length && !$col.hasClass("show")) $col.collapse("show");
    }
  });
  $("#searchCount").text(count);
  $("#searchResults").toggle(count > 0);
}

function limparPesquisa() {
  $(".permission-item, .category-card").show();
  $(".collapse").collapse("show");
  $("#searchCount").text("");
  $("#searchResults").hide();
}

function configurarAtalhosTeclado() {
  $(document).on("keydown", (e) => {
    const $inp = $("#searchPermissions");
    if ((e.ctrlKey || e.metaKey) && e.key === "f" && $inp.length) {
      e.preventDefault();
      $inp.focus();
    }
    if (
      (e.ctrlKey || e.metaKey) &&
      e.key === "s" &&
      $("#permissionsSection").length
    ) {
      e.preventDefault();
      Salvar();
    }
    if (e.key === "Escape" && $inp.length && $inp.val()) {
      $inp.val("");
      limparPesquisa();
    }
  });
}

function inicializarCategorias() {
  $(".category-card").each(function () {
    const tipo = $(this).data("tipo"),
      cat = $(this).data("categoria");
    if (!AppState.categorias[tipo]) AppState.categorias[tipo] = {};
    AppState.categorias[tipo][cat] = {
      expandida: true,
      totalItens: 0,
      itensSelecionados: 0,
    };
    atualizarContagemCategoria($(this));
  });
}

function toggleCategoria(checkbox) {
  const $checkbox = $(checkbox),
    checked = $checkbox.is(":checked");
  const $card = $checkbox.closest(".category-card");
  const filtroAtivo = $("#searchResults").is(":visible");
  const $itens = filtroAtivo
    ? $card.find(".permission-item:visible .perm-check")
    : $card.find(".perm-check");
  $itens.each(function () {
    if ($(this).is(":checked") !== checked)
      $(this).prop("checked", checked).trigger("change");
  });
  atualizarContagemCategoria($card);
  atualizarContadores();
}

function atualizarCheckboxCategoria(itemCheckbox) {
  const $card = $(itemCheckbox).closest(".category-card");
  const $all = $card.find(".perm-check"),
    total = $all.length,
    checked = $all.filter(":checked").length;
  $card
    .find(".category-checkbox")
    .prop("checked", checked === total && total > 0)
    .prop("indeterminate", checked > 0 && checked < total);
  atualizarContagemCategoria($card);
}

function atualizarContagemCategoria($card) {
  const $all = $card.find(".perm-check"),
    total = $all.length,
    checked = $all.filter(":checked").length;
  $card.find(".count-badge").text(`${checked}/${total}`);
  const tipo = $card.data("tipo"),
    cat = $card.data("categoria");
  if (AppState.categorias[tipo]?.[cat]) {
    AppState.categorias[tipo][cat].totalItens = total;
    AppState.categorias[tipo][cat].itensSelecionados = checked;
  }
}

function toggleSessao(tipo) {
  const $collapses = $(`.category-card[data-tipo="${tipo}"] .collapse`);
  const expandido = $collapses.filter(".show").length > 0;
  $collapses.collapse(expandido ? "hide" : "show");
  $(`#toggleIcon${tipo.charAt(0).toUpperCase() + tipo.slice(1)}`).toggleClass("collapsed", expandido);
}

function atualizarContadores() {
  $(".category-card").each(function () {
    atualizarContagemCategoria($(this));
  });
}

function Salvar() {
  const cCons = parseInt($("#contaConsulta").val(), 10) || 0;
  const cServ = parseInt($("#contaServico").val(), 10) || 0;

  let total = 0;

  for (let i = 0; i < cCons; i++) {
    AppState.estadoCheckboxes.consultas[i] = $(`#consulta${i}`).is(":checked");
    if (AppState.estadoCheckboxes.consultas[i]) total++;
  }

  for (let i = 0; i < cServ; i++) {
    AppState.estadoCheckboxes.servicos[i] = $(`#servico${i}`).is(":checked");
    if (AppState.estadoCheckboxes.servicos[i]) total++;
  }

  if (total === 0 && !confirm("Nenhuma permissão selecionada. Deseja continuar?")) {
    return;
  }

  mostrarLoading(true);
  $("#envia").val("1");
  $("#AcessosUsuario").submit();
}

function restaurarEstadoCheckboxes() {
  for (const k in AppState.estadoCheckboxes.consultas)
    $(`#consulta${k}`).prop("checked", AppState.estadoCheckboxes.consultas[k]);
  for (const k in AppState.estadoCheckboxes.servicos)
    $(`#servico${k}`).prop("checked", AppState.estadoCheckboxes.servicos[k]);
  sincronizarTudo();
}

function abrirModalImportar() {
  new bootstrap.Modal(document.getElementById("modalImportar")).show();
  const $sel = $("#usuarioOrigemSelect");
  const $switch = $("#switchIncluirInativos");

  // Resetar switch para desligado a cada abertura
  $switch.prop("checked", false);

  // Destruir instância anterior do Select2 para garantir estado limpo
  if ($sel.hasClass("select2-hidden-accessible")) {
    $sel.val("").trigger("change");
    $sel.select2("destroy");
  }

  // Recriar Select2 — o matcher já filtra inativos via estado do switch
  $sel.select2({
    theme: "bootstrap-5",
    width: "100%",
    placeholder: $sel.data("placeholder") || "Nome ou login...",
    allowClear: true,
    dropdownParent: $("#modalImportar .modal-body"),
    matcher: select2Matcher,
  });

  // Switch: ao alternar, forçar o Select2 a re-executar o matcher
  $switch.off("change.importar").on("change.importar", function () {
    // Se desligou com um inativo selecionado, limpar seleção
    if (!$(this).is(":checked")) {
      const $opt = $sel.find("option:selected");
      if ($opt.data("ativo") === 0) {
        $sel.val("").trigger("change");
      }
    }
    // Fechar e reabrir o dropdown para forçar re-render do matcher
    $sel.select2("close");
    $sel.select2("open");
  });
}

function ConfirmarImportacao() {
  const $sel = $("#usuarioOrigemSelect");
  if (!$sel.val()) {
    mostrarNotificacao("Por favor, selecione o usuário de origem.", "warning");
    return;
  }
  bootstrap.Modal.getInstance(document.getElementById("modalImportar"))?.hide();
  mostrarModalConfirmacao(
    "Importar permissões",
    "As permissões atuais serão substituídas pelas do usuário selecionado.",
    () => {
      $("#usuarioOrigemHidden").val($sel.val());
      $("#envia").val("IMPORTAR");
      mostrarLoading(true);
      $("#AcessosUsuario").submit();
    },
  );
}

const PDF = {
  sistema: "SWNet",
  titulo: "Permissões do usuário",
  mg: { top: 38, right: 10, bottom: 20, left: 10 },
  cor: {
    hBg: [255, 255, 255],
    hBd: [200, 200, 200],
    hTx: [0, 0, 0],
    titBg: [240, 240, 245],
    secBg: [59, 59, 59],
    ok: [34, 139, 34],
    no: [220, 53, 69],
    ln: [200, 200, 200],
    sec: [100, 100, 100],
  },
  fs: { h: 16, t: 12, st: 10, b: 7, m: 6 },
};

async function gerarPDF() {
  mostrarLoading(true);
  try {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF("p", "mm", "a4");
    const dados = coletarDadosUsuarioPDF();
    const servs = coletarPermissoesPDF("servico");
    const conss = coletarPermissoesPDF("consulta");
    const logo = await carregarLogoPDF();

    desenharCabecalhoPDF(pdf, dados, logo);
    let y = PDF.mg.top + 5;
    y = adicionarInfoUsuarioPDF(pdf, dados, y);
    if (servs.length) adicionarSecaoPDF(pdf, "SERVIÇOS", servs, y);

    pdf.addPage();
    desenharCabecalhoPDF(pdf, dados, logo);
    y = PDF.mg.top + 5;
    y = adicionarInfoUsuarioPDF(pdf, dados, y);
    if (conss.length) adicionarSecaoPDF(pdf, "CONSULTAS", conss, y);

    adicionarRodapePDF(pdf, dados);
    pdf.save(gerarNomePDF(dados.nomeUsuario));
    mostrarNotificacao("PDF gerado com sucesso!", "success");
  } catch (e) {
    console.error(e);
    mostrarNotificacao("Erro ao gerar PDF. Tente novamente.", "error");
  } finally {
    mostrarLoading(false);
  }
}

function desenharCabecalhoPDF(pdf, dados, logo) {
  const pw = pdf.internal.pageSize.getWidth();
  pdf.setFillColor(...PDF.cor.hBg);
  pdf.setDrawColor(...PDF.cor.hBd);
  pdf.setLineWidth(0.5);
  pdf.rect(0, 0, pw, 30, "FD");

  if (logo?.length > 100) {
    try {
      pdf.addImage(logo, "PNG", 8, 4, 44, 18);
    } catch (e) {}
  }

  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(PDF.fs.h);
  pdf.setTextColor(...PDF.cor.hTx);
  pdf.text(PDF.titulo, pw / 2, 12, { align: "center" });

  pdf.setFont("helvetica", "normal");
  pdf.setFontSize(PDF.fs.b);
  pdf.setTextColor(...PDF.cor.sec);

  const sep = "  |  ";
  const txts = [
    "Sistema: " + PDF.sistema,
    "Usuário logado: " + dados.usuarioLogado,
    "Emitido em: " + new Date().toLocaleString("pt-BR"),
  ];

  const blocoW =
    txts.reduce((a, t) => a + pdf.getTextWidth(t), 0) +
    pdf.getTextWidth(sep) * 2;

  let x = (pw - blocoW) / 2;

  txts.forEach((t, i) => {
    pdf.text(t, x, 19);
    x += pdf.getTextWidth(t);
    if (i < 2) {
      pdf.text(sep, x, 19);
      x += pdf.getTextWidth(sep);
    }
  });
}

function adicionarInfoUsuarioPDF(pdf, dados, y) {
  const pw = pdf.internal.pageSize.getWidth(),
    lm = PDF.mg.left,
    cw = pw - lm * 2;

  pdf.setFillColor(...PDF.cor.titBg);
  pdf.rect(lm, y, cw, 25, "F");
  pdf.setDrawColor(180, 180, 180);
  pdf.setLineWidth(0.3);
  pdf.rect(lm, y, cw, 25, "S");

  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(PDF.fs.st);
  pdf.setTextColor(0, 0, 0);
  pdf.text("DADOS DO USUÁRIO", pw / 2, y + 7, { align: "center" });

  pdf.setFontSize(PDF.fs.b);
  pdf.text("Filial: " + dados.filialUsuario, pw / 2, y + 15, { align: "center" });
  pdf.text("Nome: " + dados.nomeUsuario, pw / 2, y + 21, { align: "center" });

  return y + 29;
}

function adicionarSecaoPDF(pdf, titulo, permissoes, y) {
  const pw = pdf.internal.pageSize.getWidth(),
    lm = PDF.mg.left,
    cw = pw - 20;
  pdf.setFillColor(...PDF.cor.secBg);
  pdf.rect(lm, y, cw, 8, "F");
  pdf.setTextColor(255, 255, 255);
  pdf.setFontSize(PDF.fs.m);
  pdf.setFont("helvetica", "normal");
  pdf.text(
    `(${permissoes.filter((p) => p.status === "Sim").length}/${permissoes.length} habilitados)`,
    lm + 3,
    y + 5.5,
  );
  pdf.setFontSize(PDF.fs.t);
  pdf.setFont("helvetica", "bold");
  pdf.text(titulo, pw / 2, y + 5.5, { align: "center" });
  y += 11;

  const linhas = [];
  for (let i = 0; i < permissoes.length; i += 3) {
    const l = [];
    for (let j = 0; j < 3; j++)
      l.push(
        i + j < permissoes.length ? permissoes[i + j].nome : "",
        i + j < permissoes.length ? permissoes[i + j].status : "",
      );
    linhas.push(l);
  }
  const lg = cw / 3,
    ln = lg * 0.75,
    ls = lg * 0.25;
  pdf.autoTable({
    startY: y,
    head: [
      ["Permissão", "Status", "Permissão", "Status", "Permissão", "Status"],
    ],
    body: linhas,
    margin: { left: lm, right: PDF.mg.right },
    styles: {
      fontSize: PDF.fs.m,
      cellPadding: 2,
      overflow: "linebreak",
      lineWidth: 0.2,
      lineColor: [180, 180, 180],
    },
    headStyles: {
      fillColor: PDF.cor.secBg,
      textColor: [255, 255, 255],
      fontStyle: "bold",
      halign: "center",
    },
    columnStyles: {
      0: { cellWidth: ln },
      1: { cellWidth: ls, halign: "center", fontStyle: "bold" },
      2: { cellWidth: ln },
      3: { cellWidth: ls, halign: "center", fontStyle: "bold" },
      4: { cellWidth: ln },
      5: { cellWidth: ls, halign: "center", fontStyle: "bold" },
    },
    alternateRowStyles: { fillColor: [250, 248, 255] },
    didParseCell(data) {
      if (data.column.index % 2 === 1 && data.section === "body") {
        const v = data.cell.raw.toString().toUpperCase();
        if (v === "SIM") {
          data.cell.styles.textColor = PDF.cor.ok;
          data.cell.styles.fontStyle = "bold";
        } else if (v === "NÃO") {
          data.cell.styles.textColor = PDF.cor.no;
          data.cell.styles.fontStyle = "bold";
        }
      }
    },
    didDrawPage(data) {
      if (!data?.table || !data?.settings) return;
      const ml = Number(data.settings.margin.left);
      const g1 = ml + ln + ls,
        g2 = g1 + ln + ls;
      const sy = data.table.head?.[0]
        ? Number(data.table.head[0].y + data.table.head[0].height)
        : null;
      const ey = data.cursor?.y ? Number(data.cursor.y) : null;
      if (!sy || !ey || ey <= sy) return;
      pdf.setDrawColor(147, 112, 219);
      pdf.setLineWidth(0.8);
      pdf.line(g1, sy, g1, ey);
      pdf.line(g2, sy, g2, ey);
    },
  });
}

function adicionarRodapePDF(pdf, dados) {
  const pw = pdf.internal.pageSize.getWidth(),
    ph = pdf.internal.pageSize.getHeight(),
    tp = pdf.internal.getNumberOfPages();
  for (let i = 1; i <= tp; i++) {
    pdf.setPage(i);
    pdf.setDrawColor(...PDF.cor.ln);
    pdf.setLineWidth(0.3);
    pdf.line(10, ph - 12, pw - 10, ph - 12);
    pdf.setTextColor(...PDF.cor.sec);
    pdf.setFontSize(8);
    pdf.setFont("helvetica", "normal");
    pdf.text(PDF.sistema + " - Usuário: " + dados.usuarioLogado, 10, ph - 6);
    pdf.text(`Página ${i} de ${tp}`, pw / 2, ph - 6, { align: "center" });
    pdf.text(
      new Date().toLocaleString("pt-BR", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
      }),
      pw - 10,
      ph - 6,
      { align: "right" },
    );
  }
}

async function carregarLogoPDF() {
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = "Anonymous";
    img.onload = function () {
      const scale = img.width > 300 ? 300 / img.width : 1;
      const canvas = document.createElement("canvas");
      canvas.width = Math.round(img.width * scale);
      canvas.height = Math.round(img.height * scale);
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#FFFFFF";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      resolve(canvas.toDataURL("image/png"));
    };
    img.onerror = () => resolve(null);
    img.src = caminhoLogo;
  });
}

function coletarDadosUsuarioPDF() {
  const filial = $("#filial").val() || "";
  return {
    nomeUsuario:
      $("#userNameDisplay").text().trim() ||
      $("#nome").val() ||
      "Não informado",
    filialUsuario: ObterDescricaoFilial(filial) || filial || "Não informado",
    usuarioLogado:
      typeof loginUsuarioLogado !== "undefined"
        ? loginUsuarioLogado
        : $("#loginUsuarioLogado").val() || "Sistema",
    idUsuarioLogado:
      typeof idUsuarioLogado !== "undefined"
        ? idUsuarioLogado
        : $("#idUsuarioLogado").val() || "N/A",
  };
}

function coletarPermissoesPDF(tipo) {
  return $(
    `#${tipo === "servico" ? "servicos" : "consultas"}Container .permission-item`,
  )
    .map(function () {
      const nome =
        $(this).find("label").text().trim() || $(this).data("nome") || "";
      return nome
        ? {
            nome,
            status: $(this).find(".perm-check").is(":checked") ? "Sim" : "Não",
            categoria: $(this).data("categoria") || "",
          }
        : null;
    })
    .get()
    .filter(Boolean);
}

function gerarNomePDF(nomeUsuario) {
  const n = nomeUsuario
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "_")
    .replace(/[^a-zA-Z0-9_]/g, "")
    .substring(0, 30);
  const d = new Date();
  return `Permissoes_${n}_${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}_${String(d.getHours()).padStart(2, "0")}${String(d.getMinutes()).padStart(2, "0")}.pdf`;
}

const FILIAIS = Object.freeze({
  "00": { nome: "Pirapetinga", cls: "filial-00" },
  "01": { nome: "Rio de Janeiro", cls: "filial-01" },
  "02": { nome: "São Paulo", cls: "filial-02" },
  "03": { nome: "Corporativo", cls: "filial-03" },
  "04": { nome: "Uberaba", cls: "filial-04" },
  "05": { nome: "Bento Gonçalves", cls: "filial-05" },
  "06": { nome: "Fortaleza", cls: "filial-06" },
  "07": { nome: "Saquarema", cls: "filial-07" },
  "08": { nome: "Araçatuba", cls: "filial-08" },
  "09": { nome: "Blumenau", cls: "filial-09" },
  10: { nome: "Pacajus", cls: "filial-10" },
  11: { nome: "Porto Feliz", cls: "filial-11" },
  12: { nome: "Três Barras", cls: "filial-12" },
  13: { nome: "Florestal", cls: "filial-13" },
});

function ObterDescricaoFilial(cod) {
  return FILIAIS[cod] ? FILIAIS[cod].nome : "";
}

function renderFilialBadge(cod) {
  const f = FILIAIS[cod];
  if (!f) return "";
  return `<span class="filial-badge ${f.cls}">${f.nome}</span>`;
}

function mostrarLoading(show) {
  if (show) {
    $("#loader").removeClass("d-none");
  } else {
    $("#loader").addClass("d-none");
  }
}

function mostrarLoadingInline(show) {
  if (show) {
    $("#skPlaceholder").addClass("d-none");
    $("#skPlaceholderLoading").removeClass("d-none");
  } else {
    $("#skPlaceholder").removeClass("d-none");
    $("#skPlaceholderLoading").addClass("d-none");
  }
}

function mostrarNotificacao(msg, tipo = "info") {
  const cfg = TOAST_ICON[tipo] ?? TOAST_ICON.info;
  const tipoClass = tipo === "error" ? "danger" : tipo;
  const item = $(`
    <div class="toast-item toast-item-${tipoClass}">
      <div class="toast-corpo">
        <i class="bi ${cfg.cls}" style="color:${cfg.cor};font-size:1rem;flex-shrink:0;"></i>
        <span class="flex-grow-1">${msg}</span>
        <button class="toast-fechar" aria-label="Fechar">&times;</button>
      </div>
      <div class="toast-barra toast-barra-${tipoClass}" style="animation-duration:${TOAST_CFG.delay}ms;"></div>
    </div>
  `);
  item.find(".toast-fechar").on("click", () => item.remove());
  $("#toastWrapper").append(item);
  setTimeout(() => item.fadeOut(300, () => item.remove()), TOAST_CFG.delay);
}

function mostrarModalConfirmacao(titulo, mensagem, cb) {
  const $modal = $("#modalConfirmacao");
  $modal.find("#modalConfirmacaoTitulo").text(titulo);
  $modal.find("#modalConfirmacaoMsg").text(mensagem);
  $modal
    .find("#modalConfirmacaoBtn")
    .off("click")
    .on("click", function () {
      bootstrap.Modal.getInstance($modal[0])?.hide();
      if (typeof cb === "function") cb();
    });
  new bootstrap.Modal($modal[0]).show();
}

    Object.assign(window, {
          AppState,
          SelecionouUsuario,
          Buscar,
          VoltarFormBusca,
          Salvar,
          abrirModalImportar,
          ConfirmarImportacao,
          gerarPDF,
          toggleSessao,
          ObterDescricaoFilial,
          renderFilialBadge,
          mostrarModalConfirmacao,
        });
      })(jQuery);
    </script>
  </body>
</html>