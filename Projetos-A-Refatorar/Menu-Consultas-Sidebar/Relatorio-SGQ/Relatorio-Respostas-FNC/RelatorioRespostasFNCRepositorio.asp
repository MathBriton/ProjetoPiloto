<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../cnn2/cnn2.asp" -->
<!--#include file="../../../funcoes/funcoes.asp" -->
<!--#include file="../../../includes/variaveis.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046
Response.ContentType = "application/json"

Dim acao
acao = Request("acao")

If acao = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Ação não informada""}")
    Response.End()
End If

Dim resultado
Select Case acao
    Case "listarRespostas"
        resultado = ListarRespostas()
    Case "listarDefeitos"
        resultado = ListarDefeitos()
    Case "listarDepartamentos"
        resultado = ListarDepartamentos()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)

Function ListarRespostas()
    Dim SQL, rsRespostas, jsonArray, primeiroItem

    idFNC = Request("filtroIdFNC")
    codCliente = Request("filtroCodCliente")
    codLoja = Request("filtroCodLoja")
    idDepartamento = Request("filtroDepartamento")
    idDefeito = Request("filtroDefeito")
    codPlanta = Request("filtroFilial")
    situacao = Request("filtroSituacao")
    dataInclusaoDe = Request("filtroDataInclusaoDe")
    dataInclusaoAte = Request("filtroDataInclusaoAte")

    If codPlanta = "" Then
        ListarRespostas = "{""sucesso"": false, ""mensagem"": ""Planta não informada""}"
        Exit Function
    End If

    SQL = "SELECT r.IdFNC, f.Planta as Filial, f.Cliente as CodCliente, f.Loja as CodLoja, f.Situacao, "

    SQL = SQL & " STRING_ESCAPE(CONVERT(VARCHAR(MAX), f.Descricao), 'json') AS Descricao, "
    SQL = SQL & " STRING_ESCAPE(CONVERT(VARCHAR(MAX), r.Justificativa), 'json') AS Justificativa, "
    SQL = SQL & " r.QuemIncluiu as IdUsuarioInclusao, c1.CRM_NOME as NomeUsuarioInclusao, CLI.NOME_CLIENTE AS NomeCliente,"
    SQL = SQL & " r.QuemFinalizou as IdUsuarioFinalizacao, c2.CRM_NOME as NomeUsuarioFinalizacao,  s.Descricao as Setor, "
    SQL = SQL & " f.IdSetor as IdDepartamento, f.IdDefeito, r.DataInclusao, r.DataFinalizacao, def.Tipo as TipoDefeito"
    SQL = SQL & " FROM Resposta r "
    SQL = SQL & " LEFT JOIN FNC f on f.Id = r.IdFNC "
    SQL = SQL & " LEFT JOIN CRM_DTI c1 on c1.CRM_R_E_C_N_O = r.QuemIncluiu "
    SQL = SQL & " LEFT JOIN CRM_DTI c2 on c2.CRM_R_E_C_N_O = r.QuemFinalizou "
    SQL = SQL & " LEFT JOIN Setor s ON f.IdSetor = s.Id " 
    SQL = SQL & "INNER JOIN Defeito def WITH(NOLOCK) ON f.IdDefeito = def.Id "
    SQL = SQL & "OUTER APPLY (  SELECT " &_
                                    "TOP 1 NOME_CLIENTE " &_
                                "FROM BI_CLIENTES " &_
                                "WHERE CIDADE = " &_
                                "CASE f.Planta " &_
                                                "WHEN '00' THEN 'PIRAPETINGA' " &_
                                                "WHEN '04' THEN 'UBERABA' " &_
                                                "WHEN '05' THEN 'BENTO' " &_
                                                "WHEN '06' THEN 'MARANGUAPE' " &_
                                                "WHEN '07' THEN 'SAQUAREMA' " &_
                                                "WHEN '08' THEN 'ARACATUBA' " &_
                                                "WHEN '09' THEN 'BLUMENAU' " &_
                                                "WHEN '10' THEN 'PACAJUS' " &_
                                                "WHEN '11' THEN 'PORTO FELIZ' " &_
                                                "WHEN '12' THEN 'TRES BARRAS' " &_
                                                "WHEN '13' THEN 'FLORESTAL' " &_
                                "END AND " &_
                                "CASE WHEN f.Planta IN ('06','05') THEN CODIGO_ERP ELSE COD_CLIENTE END = f.Cliente COLLATE SQL_Latin1_General_CP1_CI_AS AND " &_
                                "CASE WHEN f.Planta IN ('09') THEN '01' ELSE LOJA_CLIENTE END = f.Loja COLLATE SQL_Latin1_General_CP1_CI_AS) AS CLI "

    If codPlanta <> "" Then
        SQL = SQL & " WHERE f.Planta = '" & codPlanta & "'"
    End If

    If codCliente <> "" Then
        SQL = SQL & " AND f.Cliente = '" & codCliente & "'"
    End If

    If codLoja <> "" Then
        SQL = SQL & " AND f.Loja = '" & codLoja & "'"
    End If

    If idDepartamento <> "" Then
        SQL = SQL & " AND f.IdSetor = '" & idDepartamento & "'"
    End If

    If idDefeito <> "" Then
        SQL = SQL & " AND f.IdDefeito = '" & idDefeito & "'"
    End If

    If idFNC <> "" Then
        SQL = SQL & " AND r.IdFNC = '" & idFNC & "'"
    End If

    If situacao <> "" Then
        SQL = SQL & " AND f.Situacao = '" & situacao & "'"
    End If

    If dataInclusaoDe <> "" Then
        SQL = SQL & " AND r.DataInclusao >= '" & Replace(dataInclusaoDe, "-", "") & "'"
    End If

    If dataInclusaoAte <> "" Then
        SQL = SQL & " AND r.DataInclusao <= '" & Replace(dataInclusaoAte, "-", "") & "'"
    End If

    SQL = SQL & " ORDER BY r.IdFNC DESC"

    On Error Resume Next
    SET rsRespostas = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsRespostas) Then
        ListarRespostas = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de respostas: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsRespostas.EOF Then
        Do While Not rsRespostas.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """idFNC"": """ & EscapeJSON(rsRespostas("IdFNC")) & """, "
            jsonArray = jsonArray & """filial"": """ & EscapeJSON(rsRespostas("Filial")) & """, "
            jsonArray = jsonArray & """tipoDefeito"": """ & EscapeJSON(rsRespostas("TipoDefeito")) & """, "
            jsonArray = jsonArray & """setor"": """ & EscapeJSON(rsRespostas("Setor")) & """, "
            jsonArray = jsonArray & """codCliente"": """ & EscapeJSON(rsRespostas("CodCliente")) & """, "
            jsonArray = jsonArray & """codLoja"": """ & EscapeJSON(rsRespostas("CodLoja")) & """, "
            jsonArray = jsonArray & """nomeCliente"": """ & EscapeJSON(rsRespostas("NomeCliente")) & """, "
            jsonArray = jsonArray & """situacao"": """ & EscapeJSON(rsRespostas("Situacao")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsRespostas("Descricao")) & """, "
            jsonArray = jsonArray & """justificativa"": """ & EscapeJSON(rsRespostas("Justificativa")) & """, "
            jsonArray = jsonArray & """idUsuarioInclusao"": """ & EscapeJSON(rsRespostas("IdUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """nomeUsuarioInclusao"": """ & EscapeJSON(rsRespostas("NomeUsuarioInclusao")) & """, "
            jsonArray = jsonArray & """idUsuarioFinalizacao"": """ & EscapeJSON(rsRespostas("IdUsuarioFinalizacao")) & """, "
            jsonArray = jsonArray & """nomeUsuarioFinalizacao"": """ & EscapeJSON(rsRespostas("NomeUsuarioFinalizacao")) & """, "
            jsonArray = jsonArray & """idDepartamento"": """ & EscapeJSON(rsRespostas("IdDepartamento")) & """, "
            jsonArray = jsonArray & """idDefeito"": """ & EscapeJSON(rsRespostas("IdDefeito")) & """, "
            jsonArray = jsonArray & """dataInclusao"": """ & EscapeJSON(rsRespostas("DataInclusao")) & """, "
            jsonArray = jsonArray & """dataFinalizacao"": """ & EscapeJSON(rsRespostas("DataFinalizacao")) & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsRespostas.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsRespostas.Close()
    SET rsRespostas = NOTHING
    
    ListarRespostas = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarDefeitos()
    Dim SQL, rsDefeitos, jsonArray, primeiroItem

    SQL = "SELECT Id, Descricao FROM Defeito"

    On Error Resume Next
    SET rsDefeitos = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDefeitos) Then
        ListarDefeitos = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de defeitos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsDefeitos.EOF Then
        Do While Not rsDefeitos.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsDefeitos("Id")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsDefeitos("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsDefeitos.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsDefeitos.Close()
    SET rsDefeitos = NOTHING
    
    ListarDefeitos = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarDepartamentos()
    Dim SQL, rsDepartamentos, jsonArray, primeiroItem

    SQL = "SELECT Id, Descricao FROM Setor"

    On Error Resume Next
    SET rsDepartamentos = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDepartamentos) Then
        ListarDepartamentos = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de departamentos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsDepartamentos.EOF Then
        Do While Not rsDepartamentos.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsDepartamentos("Id")) & """, "
            jsonArray = jsonArray & """descricao"": """ & EscapeJSON(rsDepartamentos("Descricao")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsDepartamentos.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsDepartamentos.Close()
    SET rsDepartamentos = NOTHING
    
    ListarDepartamentos = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function
%>