<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Buffer = True

' ── Sanitização do código ────────────────────────────────────────────────────
Dim codigo
codigo = Trim(Request.QueryString("cod"))

' Remove qualquer tentativa de path traversal
codigo = Replace(codigo, "..", "")
codigo = Replace(codigo, "/",  "")
codigo = Replace(codigo, "\",  "")
codigo = Replace(codigo, ":",  "")
codigo = Replace(codigo, "%",  "")

' Código vazio → 400
If codigo = "" Then
  Response.Status      = "400 Bad Request"
  Response.ContentType = "text/plain"
  Response.Write "Parâmetro 'cod' não informado."
  Response.End
End If

' ── Monta caminhos físicos via Server.MapPath ─────────────────────────────────
Dim caminhoJpg, caminhoPng
caminhoJpg = "c:\Inetpub\wwwroot\ft\" & codigo & ".jpg"
caminhoPng = "c:\Inetpub\wwwroot\ft\" & codigo & ".png"

' ── Detecta qual extensão existe ─────────────────────────────────────────────
Dim fso, caminho, contentType
Set fso = Server.CreateObject("Scripting.FileSystemObject")

If fso.FileExists(caminhoJpg) Then
  caminho     = caminhoJpg
  contentType = "image/jpeg"
ElseIf fso.FileExists(caminhoPng) Then
  caminho     = caminhoPng
  contentType = "image/png"
Else
  Set fso = Nothing
  Response.Status      = "404 Not Found"
  Response.ContentType = "text/plain"
  Response.Write "Imagem não encontrada."
  Response.End
End If

Set fso = Nothing

' ── Lê e serve o arquivo binário ────────────────────────────────────────────
Dim stream
Set stream = Server.CreateObject("ADODB.Stream")

On Error Resume Next
stream.Open
stream.Type = 1  ' adTypeBinary
stream.LoadFromFile caminho

If Err.Number <> 0 Then
  stream.Close
  Set stream = Nothing
  Response.Status      = "500 Internal Server Error"
  Response.ContentType = "text/plain"
  Response.Write "Erro ao ler o arquivo."
  Response.End
End If
On Error GoTo 0

Response.ContentType = contentType
Response.AddHeader "Cache-Control", "max-age=3600, private"
Response.BinaryWrite stream.Read

stream.Close
Set stream = Nothing
%>
