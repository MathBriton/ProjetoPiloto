<!--#include file="../../cnn2/cnn2.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->
<!--#include file="../../includes/variaveis.asp"-->

<%
If UCase(Trim(Request("acao"))) = "GET" Then

  Response.ContentType = "application/json"
  Response.Charset     = "UTF-8"

  Dim getCodigo
  getCodigo = Trim(Request("codigo"))

  If getCodigo = "" Then
    Response.Write "{""sucesso"":false,""mensagem"":""Codigo nao informado.""}"
    Response.End
  End If

  Dim sqlGet
  sqlGet = "" & _
    "SELECT CODIGO, CNPJ, NOME " & _
    "FROM RECEBIMENTO_FORNECEDORES " & _
    "WHERE CODIGO = '" & Replace(getCodigo, "'", "''") & "' " & _
      "AND ISNULL(EXCLUIDO,'') <> '*'"

  Dim rsGet
  SET rsGet = Conecta("PIR-CORPORATIVO", sqlGet)

  If isObject(rsGet) And Not (rsGet Is Nothing) And Not rsGet.EOF Then
    Dim jCodigo, jCnpj, jNome
    jCodigo = Replace(rsGet("CODIGO") & "", """", "\""")
    jCnpj   = Replace(rsGet("CNPJ")   & "", """", "\""")
    jNome   = Replace(rsGet("NOME")   & "", """", "\""")
    Response.Write "{""sucesso"":true,""dados"":{""CODIGO"":""" & jCodigo & """,""CNPJ"":""" & jCnpj & """,""NOME"":""" & jNome & """}}"
    rsGet.Close()
  Else
    Response.Write "{""sucesso"":false,""mensagem"":""Registro nao encontrado.""}"
  End If

  Response.End

End If

Select Case Request("acao")

  Case "ALT"

    SQLalt = "" & _
      "BEGIN TRANSACTION TRANS; " & _
        "UPDATE RECEBIMENTO_FORNECEDORES SET " & _
          "CNPJ = '" & Replace(Trim(Request("edtCNPJ")), "'", "''") & "'" & _
          ", NOME = '" & Replace(UCase(Trim(Request("edtNOME"))), "'", "''") & "'" & _
          ", DATA_ALTERACAO = getdate()" & _
          ", QUEM_ALTEROU = '" & Session("CODUSER") & "' " & _
        "WHERE CODIGO = '" & Replace(Trim(Request("edtCODIGO")), "'", "''") & "'; " & _
      "IF @@ERROR <> 0 " & _
        "ROLLBACK TRANSACTION TRANS " & _
      "ELSE " & _
        "COMMIT TRANSACTION TRANS "

    Dim altErro : altErro = False
    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQLalt)
    If Err.Number <> 0 Then altErro = True
    On Error GoTo 0
    SET PIR_CORPORATIVO = Nothing

    If Request.ServerVariables("HTTP_X_REQUESTED_WITH") = "XMLHttpRequest" Then
      Response.ContentType = "application/json"
      Response.Charset     = "UTF-8"
      If altErro Then
        Response.Write "{""sucesso"":false,""mensagem"":""Erro ao atualizar o registro.""}"
      Else
        Response.Write "{""sucesso"":true}"
      End If
      Response.End
    End If

  Case "INS"

    cadastro = "N"

    SQLsel = "" & _
      "SELECT CNPJ " & _
      "FROM RECEBIMENTO_FORNECEDORES " & _
      "WHERE CNPJ = '" & Trim(Request("edtCNPJ")) & "' " & _
        "AND ISNULL(EXCLUIDO,'') <> '*'"

    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQLsel)

    If Not isObject(PIR_CORPORATIVO) Or (PIR_CORPORATIVO Is Nothing) Then
      Response.End
    End If

    If Not PIR_CORPORATIVO.EOF Then
      cadastro = "S"
      %>
      <script>alert("C.N.P.J. já cadastrado !!!");</script>
      <%
    End If

    SQLins = "" & _
      "BEGIN TRANSACTION TRANS; " & _
        "INSERT INTO RECEBIMENTO_FORNECEDORES ( " & _
           "CODIGO " & _
          ",CNPJ " & _
          ",NOME " & _
          ",QUEM_INCLUIU " & _
          ",DATA_INCLUSAO " & _
          ",QUEM_ALTEROU " & _
          ",DATA_ALTERACAO " & _
          ",QUEM_EXCLUIU " & _
          ",DATA_EXCLUSAO " & _
          ",EXCLUIDO " & _
        ") VALUES ( " & _
          "'" & Trim(Request("edtCODIGO")) & "' " & _
          ",'" & Trim(Request("edtCNPJ")) & "' " & _
          ",'" & Trim(Request("edtNOME")) & "' " & _
          ",'" & Session("CODUSER") & "' " & _
          ",getdate() " & _
          ",'' " & _
          ",'' " & _
          ",'' " & _
          ",'' " & _
          ",'' " & _
        "); " & _
      "IF @@ERROR <> 0 " & _
        "ROLLBACK TRANSACTION TRANS " & _
      "ELSE " & _
        "COMMIT TRANSACTION TRANS "

    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQLins)
    SET PIR_CORPORATIVO = Nothing

    If Request.ServerVariables("HTTP_X_REQUESTED_WITH") = "XMLHttpRequest" Then
      Response.ContentType = "application/json"
      Response.Charset     = "UTF-8"
      If cadastro = "S" Then
        Response.Write "{""sucesso"":false,""mensagem"":""C.N.P.J. j\u00e1 cadastrado.""}"
      Else
        Response.Write "{""sucesso"":true}"
      End If
      Response.End
    End If

  Case "DEL"

    Dim delCodigo
    delCodigo = Trim(Request("codigo"))

    If delCodigo = "" Then
      Response.ContentType = "application/json"
      Response.Charset     = "UTF-8"
      Response.Write "{""sucesso"":false,""mensagem"":""Codigo nao informado.""}"
      Response.End
    End If

    SQLdel = "" & _
      "BEGIN TRANSACTION TRANS; " & _
        "UPDATE RECEBIMENTO_FORNECEDORES SET " & _
          "DATA_EXCLUSAO = getdate(), " & _
          "QUEM_EXCLUIU = '" & Session("CODUSER") & "', " & _
          "EXCLUIDO = '*' " & _
        "WHERE CODIGO = '" & Replace(delCodigo, "'", "''") & "'; " & _ 
      "IF @@ERROR <> 0 " & _
        "ROLLBACK TRANSACTION TRANS " & _
      "ELSE " & _
        "COMMIT TRANSACTION TRANS "

    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQLdel)
    SET PIR_CORPORATIVO = Nothing

End Select

Select Case acao

  Case "INS"
    If cadastrado <> 1 Then
      Response.Write "<script>location.href = 'grid_fornecedores.asp?acao=INS&aviso=salvou'</script>"
    Else
      Response.Write "<script>history.go(-1);</script>"
    End If

  Case "ALT"
    Response.Write "<script>location.href = 'grid_fornecedores.asp?acao=ALT&aviso=ok'</script>"

  Case "APB"
    Response.Write "<script>location.href = 'grid_fornecedores.asp?acao=APB&aviso=ok'</script>"

  Case "DEL"
    Response.Write "<script>location.href = 'grid_fornecedores.asp?acao=DEL&aviso=ok'</script>"

  Case "ERR"
    Response.Write "<script>location.href = 'grid_fornecedores.asp?acao=INS&aviso=erro'</script>"

  Case Else
    Response.Write "<script>location.href = 'grid_fornecedores.asp?acao=" & acao & "&erro=vazio'</script>"

End Select

Response.End
%>
