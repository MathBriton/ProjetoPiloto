<!-- #include file="../../cnn2/cnn2.asp" -->
<!-- #include file="../../includes/variaveis.asp" -->
<!-- #include file="../../funcoes/funcoes.asp" -->
<%
' ══════════════════════════════════════════════════════════════════
'  rpst_validadoresdecredito.asp
'  Repositório de ValidadoresDeCredito — API JSON
'
'  Aceita POST com Content-Type: application/json
'  Responde sempre JSON: { "sucesso": true|false, "mensagem": "..." }
'
'  Payload esperado (serializado pelo form_validadoresdecredito.asp):
'  {
'    "acao":    "INS" | "ALT" | "DEL",
'    "vldcod":  "<id do validador>",
'    "nivel":   "10" | "20" | ... | "60",
'    "valMin":  0 ou "1.500,00",
'    "valMax":  9999999 ou "50.000,00",
'    "plantas": ["01","03",10,11],   <- strings OU números
'    "planta":  "<filtro de contexto>",
'    "page":    "<página do grid>"
'  }
' ══════════════════════════════════════════════════════════════════

Response.Buffer      = True
Response.Expires     = 0
Response.ContentType = "application/json"
Response.Charset     = "utf-8"
Session.lcId         = 1046

' ── Leitura do corpo JSON ──────────────────────────────────────────
Dim rawBody, rpstJson
rawBody  = Request.BinaryRead(Request.TotalBytes)
rpstJson = BytesToStr(rawBody)

' ── Parse do payload — prefixo rpst_ evita colisão com includes ───
Dim rpstAcao, rpstVldcod, rpstNivel, rpstValMinRaw, rpstValMaxRaw, rpstPlantasRaw
rpstAcao       = JsonVal(rpstJson, "acao")
rpstVldcod     = JsonVal(rpstJson, "codigo")   ' form envia "codigo", não "vldcod"
rpstNivel      = JsonVal(rpstJson, "nivel")
rpstValMinRaw  = JsonVal(rpstJson, "valMin")
rpstValMaxRaw  = JsonVal(rpstJson, "valMax")
rpstPlantasRaw = JsonArr(rpstJson, "plantas")  ' array de strings (normalizado)

' ── Conversão dos valores monetários ──────────────────────────────
'    Front pode enviar número (0, 9999999) ou string mascarada ("1.500,00")
Dim rpstValMin, rpstValMax
rpstValMin = ConverteValor(rpstValMinRaw)
rpstValMax = ConverteValor(rpstValMaxRaw)

' ── Mapa de plantas selecionadas ──────────────────────────────────
Dim rpstPlantaMap
Set rpstPlantaMap = CreateObject("Scripting.Dictionary")

Dim rpstCodigos(13)
rpstCodigos(0)  = "00" : rpstCodigos(1)  = "01" : rpstCodigos(2)  = "02"
rpstCodigos(3)  = "03" : rpstCodigos(4)  = "04" : rpstCodigos(5)  = "05"
rpstCodigos(6)  = "06" : rpstCodigos(7)  = "07" : rpstCodigos(8)  = "08"
rpstCodigos(9)  = "09" : rpstCodigos(10) = "10" : rpstCodigos(11) = "11"
rpstCodigos(12) = "12" : rpstCodigos(13) = "13"

Dim rpstK
For rpstK = 0 To 13
  rpstPlantaMap.Add rpstCodigos(rpstK), "0"
Next

Dim rpstP
For Each rpstP In rpstPlantasRaw
  ' Normaliza: "10" e 10 ambos viram "10" com zero à esquerda se necessário
  Dim rpstPNorm
  rpstPNorm = Right("0" & Trim(rpstP), 2)
  If rpstPlantaMap.Exists(rpstPNorm) Then rpstPlantaMap.Item(rpstPNorm) = "1"
Next

' ══════════════════════════════════════════════════════════════════
'  LÓGICA PRINCIPAL
' ══════════════════════════════════════════════════════════════════
Dim rpstSQL, rpstMsg, rpstOk
rpstOk  = False
rpstMsg = ""

' Guarda obrigatória — aborta antes de qualquer SQL se o id vier vazio
If Trim(rpstVldcod) = "" Then
  Response.Status = "400 Bad Request"
  Response.Write "{""sucesso"":false,""mensagem"":""ID do validador não informado.""}"
  Response.End
End If

If rpstAcao = "INS" Or rpstAcao = "ALT" Then

  ' ── MERGE por planta — zero SELECTs em VBScript ────────────────
  ' Para cada planta: se marcada → MERGE (INSERT ou UPDATE conforme existência)
  '                   se desmarcada → soft-delete se existir
  rpstSQL = "BEGIN TRANSACTION TRANS; "
  For rpstK = 0 To 13
    rpstSQL = rpstSQL & MergePlanta(rpstCodigos(rpstK))
  Next
  rpstSQL = rpstSQL & Finaliza()
  ExecutaSQL rpstSQL
  rpstOk  = True
  rpstMsg = iif(rpstAcao = "INS", "Validador de Crédito inserido com sucesso!", "Validador de Crédito alterado com sucesso!")

ElseIf rpstAcao = "DEL" Then

  rpstSQL = ApagaValidador("")
  ExecutaSQL rpstSQL
  rpstOk  = True
  rpstMsg = "Validador de Crédito apagado com sucesso!"

Else

  rpstMsg = "Ação inválida: " & rpstAcao

End If

' ── Resposta JSON ──────────────────────────────────────────────────
If rpstOk Then
  Response.Write "{""sucesso"":true,""mensagem"":""" & EscapeJson(rpstMsg) & """}"
Else
  Response.Status = "400 Bad Request"
  Response.Write "{""sucesso"":false,""mensagem"":""" & EscapeJson(rpstMsg) & """}"
End If

Response.End

' ══════════════════════════════════════════════════════════════════
'  FUNÇÕES AUXILIARES
' ══════════════════════════════════════════════════════════════════

Sub ExecutaSQL(sqlStr)
  Dim rpstRs
  SET rpstRs = Conecta("PIR-CORPORATIVO", sqlStr)
  SET rpstRs = Nothing
End Sub

Function Finaliza()
  Finaliza = "IF @@ERROR <> 0 " & _
               "ROLLBACK TRANSACTION TRANS " & _
             "ELSE " & _
               "COMMIT TRANSACTION TRANS;"
End Function

' ── MergePlanta — toda a decisão INSERT/UPDATE/DELETE ocorre no T-SQL ──
' Elimina SELECTs em VBScript e conexões aninhadas dentro de loop.
'
' Planta MARCADA:
'   IF EXISTS → UPDATE (reativa se estava excluído, atualiza valores)
'   ELSE      → INSERT
'
' Planta DESMARCADA:
'   IF EXISTS (e não excluído) → soft-delete
'   ELSE                       → nada (sem registro, sem ação)
Function MergePlanta(aPlanta)
  Dim marcada, sql
  marcada = (rpstPlantaMap.Item(aPlanta) = "1")
  sql = ""

  If marcada Then
    sql = "" & _
      "IF EXISTS (" & _
        "SELECT 1 FROM ValidadoresCredito " & _
        "WHERE IdValidador = '" & rpstVldcod & "' AND Planta = '" & aPlanta & "'" & _
      ") " & _
      "  UPDATE ValidadoresCredito SET " & _
      "    Nivel = "           & rpstNivel  & ", " & _
      "    Minimo = "          & rpstValMin & ", " & _
      "    Maximo = "          & rpstValMax & ", " & _
      "    QuemAlterou = "     & IdUsuario  & ", " & _
      "    DataAlteracao = GETDATE(), " & _
      "    Excluido = NULL, QuemExcluiu = NULL, DataExclusao = NULL " & _
      "  WHERE IdValidador = '" & rpstVldcod & "' AND Planta = '" & aPlanta & "'; " & _
      "ELSE " & _
      "  INSERT INTO ValidadoresCredito " & _
      "    (IdValidador, Planta, Nivel, Minimo, Maximo, QuemIncluiu, DataInclusao) " & _
      "  VALUES (" & _
      "    '" & rpstVldcod & "', '" & aPlanta & "', " & _
      "    " & rpstNivel & ", " & rpstValMin & ", " & rpstValMax & ", " & _
      "    " & IdUsuario & ", GETDATE()" & _
      "  ); "
  Else
    sql = "" & _
      "IF EXISTS (" & _
        "SELECT 1 FROM ValidadoresCredito " & _
        "WHERE IdValidador = '" & rpstVldcod & "' AND Planta = '" & aPlanta & "' " & _
          "AND (ISNULL(Excluido,0) = 0 OR Excluido = '0')" & _
      ") " & _
      "  UPDATE ValidadoresCredito SET " & _
      "    Excluido = '1', " & _
      "    QuemExcluiu = '" & IdUsuario & "', " & _
      "    DataExclusao = GETDATE() " & _
      "  WHERE IdValidador = '" & rpstVldcod & "' AND Planta = '" & aPlanta & "'; "
  End If

  MergePlanta = sql
End Function

Function ApagaValidador(aPlanta)
  ApagaValidador = "UPDATE ValidadoresCredito SET " & _
      "Excluido = '1', " & _
      "QuemExcluiu = '" & IdUsuario & "', " & _
      "DataExclusao = GETDATE() " & _
    "WHERE IdValidador = '" & rpstVldcod & "' " & _
    iif(aPlanta <> "", "AND Planta = '" & aPlanta & "'; ", "; ")
End Function

' ══════════════════════════════════════════════════════════════════
'  UTILITÁRIOS DE JSON / ENCODING
' ══════════════════════════════════════════════════════════════════

Function BytesToStr(bytes)
  Dim rpstStream
  Set rpstStream = Server.CreateObject("ADODB.Stream")
  rpstStream.Type    = 1
  rpstStream.Open
  rpstStream.Write bytes
  rpstStream.Position = 0
  rpstStream.Type    = 2
  rpstStream.Charset = "utf-8"
  BytesToStr = rpstStream.ReadText
  rpstStream.Close
  Set rpstStream = Nothing
End Function

' Extrai valor escalar de JSON — funciona com string e número
Function JsonVal(json, chave)
  Dim pat, pos, valStr, ch
  pat = """" & chave & """"
  pos = InStr(json, pat)
  If pos = 0 Then JsonVal = "" : Exit Function

  pos = pos + Len(pat)
  Do While pos <= Len(json)
    ch = Mid(json, pos, 1)
    If ch = ":" Or ch = " " Then pos = pos + 1 Else Exit Do
  Loop

  If Mid(json, pos, 1) = """" Then
    pos    = pos + 1
    valStr = ""
    Do While pos <= Len(json)
      ch = Mid(json, pos, 1)
      If ch = """" Then Exit Do
      valStr = valStr & ch
      pos    = pos + 1
    Loop
    JsonVal = valStr
  Else
    valStr = ""
    Do While pos <= Len(json)
      ch = Mid(json, pos, 1)
      If ch = "," Or ch = "}" Or ch = " " Or ch = Chr(13) Or ch = Chr(10) Then Exit Do
      valStr = valStr & ch
      pos    = pos + 1
    Loop
    JsonVal = Trim(valStr)
  End If
End Function

' Extrai array do JSON — aceita itens com aspas ("01") OU sem aspas (10)
' e devolve sempre array de strings normalizadas
Function JsonArr(json, chave)
  Dim pat, pos, ch, item, items(), cnt
  pat = """" & chave & """"
  pos = InStr(json, pat)

  ReDim items(0)
  cnt = 0

  If pos = 0 Then JsonArr = items : Exit Function

  pos = InStr(pos, json, "[")
  If pos = 0 Then JsonArr = items : Exit Function

  pos = pos + 1
  Do While pos <= Len(json)
    ch = Mid(json, pos, 1)

    If ch = "]" Then Exit Do

    If ch = """" Then
      ' Item com aspas — lê até fechar aspas
      pos  = pos + 1
      item = ""
      Do While pos <= Len(json)
        ch = Mid(json, pos, 1)
        If ch = """" Then Exit Do
        item = item & ch
        pos  = pos + 1
      Loop
      ReDim Preserve items(cnt)
      items(cnt) = item
      cnt = cnt + 1

    ElseIf ch >= "0" And ch <= "9" Then
      ' Item numérico sem aspas — lê até , ou ]
      item = ""
      Do While pos <= Len(json)
        ch = Mid(json, pos, 1)
        If ch = "," Or ch = "]" Or ch = " " Then Exit Do
        item = item & ch
        pos  = pos + 1
      Loop
      ReDim Preserve items(cnt)
      items(cnt) = item
      cnt = cnt + 1
      pos = pos - 1  ' recua para o delimitador ser relido pelo loop externo
    End If

    pos = pos + 1
  Loop

  JsonArr = items
End Function

' Converte valor monetário pt-BR ("1.500,00") ou número (9999999) para inteiro
Function ConverteValor(raw)
  Dim v
  v = Trim(CStr(raw))
  If v = "" Or v = "0" Then ConverteValor = 0 : Exit Function

  If InStr(v, ",") > 0 Then
    v = Replace(v, ".", "")
    Dim rpstPos
    rpstPos = InStr(v, ",")
    v = Left(v, rpstPos - 1)
  End If

  If IsNumeric(v) Then
    ConverteValor = CLng(v)
  Else
    ConverteValor = 0
  End If
End Function

Function EscapeJson(s)
  s = Replace(s, "\",  "\\")
  s = Replace(s, """", "\""")
  s = Replace(s, "/",  "\/")
  s = Replace(s, Chr(13) & Chr(10), "\n")
  s = Replace(s, Chr(13), "\n")
  s = Replace(s, Chr(10), "\n")
  EscapeJson = s
End Function
%>
