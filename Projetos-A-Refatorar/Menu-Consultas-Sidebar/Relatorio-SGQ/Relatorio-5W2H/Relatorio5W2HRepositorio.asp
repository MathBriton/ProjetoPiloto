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
    Case "listarAcoes"
        resultado = ListarAcoes()
    Case "listarDefeitos"
        resultado = ListarDefeitos()
    Case "listarDepartamentos"
        resultado = ListarDepartamentos()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)

Function ListarAcoes()
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
        ListarAcoes = "{""sucesso"": false, ""mensagem"": ""Planta não informada""}"
        Exit Function
    End If

    SQL = "SELECT "
    SQL = SQL & "   a.Id, a.IdFNC, a.DataCriacao, s.Descricao as Setor, def.Tipo as TipoDefeito, "
    SQL = SQL & "   STRING_ESCAPE(CONVERT(VARCHAR(MAX), a.Acao), 'json') AS Acao, "
    SQL = SQL & "   STRING_ESCAPE(CONVERT(VARCHAR(MAX), a.Oque), 'json') AS Oque, "
    SQL = SQL & "   STRING_ESCAPE(CONVERT(VARCHAR(MAX), a.Porque), 'json') AS Porque, "
    SQL = SQL & "   STRING_ESCAPE(CONVERT(VARCHAR(MAX), a.Onde), 'json') AS Onde, "
    SQL = SQL & "   STRING_ESCAPE(CONVERT(VARCHAR(MAX), a.Quando), 'json') AS Quando, "
    SQL = SQL & "   STRING_ESCAPE(CONVERT(VARCHAR(MAX), a.Quem), 'json') AS Quem, "
    SQL = SQL & "   STRING_ESCAPE(CONVERT(VARCHAR(MAX), a.Como), 'json') AS Como, "
    SQL = SQL & "   STRING_ESCAPE(CONVERT(VARCHAR(MAX), a.Quanto), 'json') AS Quanto, "
    SQL = SQL & "   cli.NOME_CLIENTE AS NomeCliente, cli.COD_CLIENTE AS CodigoCliente, "
    SQL = SQL & "   f.Situacao, d.Descricao AS DescricaoTipoDefeito "
    SQL = SQL & "FROM FNC5W2HAcoes a "
    SQL = SQL & "LEFT JOIN FNC f on f.Id = a.IdFNC "
    SQL = SQL & "LEFT JOIN Setor s ON f.IdSetor = s.Id " 
    SQL = SQL & "LEFT JOIN Defeito d ON d.Id = f.IdDefeito "
    SQL = SQL & "INNER JOIN Defeito def WITH(NOLOCK) ON f.IdDefeito = def.Id "
    SQL = SQL & "OUTER APPLY (SELECT " &_
                                "TOP 1 NOME_CLIENTE, COD_CLIENTE " &_
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
                            "CASE WHEN f.Planta IN ('09') THEN '01' ELSE LOJA_CLIENTE END = f.Loja COLLATE SQL_Latin1_General_CP1_CI_AS) AS cli "
    SQL = SQL & "WHERE 1=1 "

    If codPlanta <> "" Then
        SQL = SQL & " AND f.Planta = '" & codPlanta & "'"
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
        SQL = SQL & " AND a.IdFNC = '" & idFNC & "'"
    End If

    If situacao <> "" Then
        SQL = SQL & " AND f.Situacao = '" & situacao & "'"
    End If

    If dataInclusaoDe <> "" Then
        SQL = SQL & " AND a.DataInclusao >= '" & Replace(dataInclusaoDe, "-", "") & "'"
    End If

    If dataInclusaoAte <> "" Then
        SQL = SQL & " AND a.DataInclusao <= '" & Replace(dataInclusaoAte, "-", "") & "'"
    End If

    SQL = SQL & " ORDER BY a.IdFNC DESC"

    On Error Resume Next
    SET rsAcoes = Conecta("BEN-CORPORATIVO", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsAcoes) Then
        ListarAcoes = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de AÇÕES: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsAcoes.EOF Then
        Do While Not rsAcoes.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If

            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsAcoes("Id")) & """, "
            jsonArray = jsonArray & """idFNC"": """ & EscapeJSON(rsAcoes("IdFNC")) & """, "
            jsonArray = jsonArray & """dataCriacao"": """ & EscapeJSON(rsAcoes("DataCriacao")) & """, "
            jsonArray = jsonArray & """setor"": """ & EscapeJSON(rsAcoes("Setor")) & """, "
            jsonArray = jsonArray & """situacao"": """ & EscapeJSON(rsAcoes("Situacao")) & """, "
            jsonArray = jsonArray & """tipoDefeito"": """ & EscapeJSON(rsAcoes("TipoDefeito")) & """, "
            jsonArray = jsonArray & """descricaoTipoDefeito"": """ & EscapeJSON(rsAcoes("DescricaoTipoDefeito")) & """, "
            jsonArray = jsonArray & """nomeCliente"": """ & EscapeJSON(rsAcoes("NomeCliente")) & """, "
            jsonArray = jsonArray & """codigoCliente"": """ & EscapeJSON(rsAcoes("CodigoCliente")) & """, "
            jsonArray = jsonArray & """acao"": """ & EscapeJSON(rsAcoes("Acao")) & """, "
            jsonArray = jsonArray & """oque"": """ & EscapeJSON(rsAcoes("Oque")) & """, "
            jsonArray = jsonArray & """porque"": """ & EscapeJSON(rsAcoes("Porque")) & """, "
            jsonArray = jsonArray & """onde"": """ & EscapeJSON(rsAcoes("Onde")) & """, "
            jsonArray = jsonArray & """quando"": """ & EscapeJSON(rsAcoes("Quando")) & """, "
            jsonArray = jsonArray & """quem"": """ & EscapeJSON(rsAcoes("Quem")) & """, "
            jsonArray = jsonArray & """como"": """ & EscapeJSON(rsAcoes("Como")) & """, "
            jsonArray = jsonArray & """quanto"": """ & EscapeJSON(rsAcoes("Quanto")) & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsAcoes.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsAcoes.Close()
    SET rsAcoes = NOTHING
    
    ListarAcoes = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
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