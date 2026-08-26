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
    Case "listar"
        resultado = Listar()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)


Function Listar()
    retiraHref = Request("filtroRetiraHref")
    codCliente = Request("filtroCodCliente")
    codLoja = Request("filtroCodLoja")
    mes = Request("filtroMes")
    ano = Request("filtroAno")
    nomePlanta  = descricaoDaPlanta(Request("filtroFilial"))
    representante = Request("filtroRepresentante")
    valores = iif(Request("filtroTipoProduto") = "1", 1, request("filtroValores")) 
    tipoProduto = parametrotipoProduto(Request("filtroTipoProduto"))
    tipoPapelao = parametrotipoPapelao(Request("filtroTipoPapelao"))
    clienteGrupo = Request("filtroClienteGrupo")
    impostos =  Request("filtroImpostos")
    
    SQL = "SET NOCOUNT ON; "
    
    SQL = SQL & "EXEC SPR_RESUMOVENDASCARTEIRA_V14 " 
    SQL = SQL & "@filialSelecionada='" & nomePlanta & "', " 
    SQL = SQL & "@datainformada='" & right(ano, 4) & "-" & left(mes, 2) & "', " 
    SQL = SQL & "@cboClienteGrupo='" & clienteGrupo & "', "  
    SQL = SQL & "@Representante='" & representante & "', "  
    SQL = SQL & "@codigoCliente='" & codCliente & "', "  
    SQL = SQL & "@comOuSemImposto='" & impostos & "', "  
    SQL = SQL & "@tipoProduto='" & tipoProduto & "', "  
    SQL = SQL & "@tipoPapelao='" & tipoPapelao & "', "  
    SQL = SQL & "@FLAGIPI='" & impostos & "', "  
    SQL = SQL & "@comutaColuna='" & valores & "', " 
    SQL = SQL & "@retiraHref='" & retiraHref & "', " 
    SQL = SQL & "@usuario='" & usuarioLogado("id") & "' "
    
    SQL = SQL & "SELECT ("
    SQL = SQL & "   SELECT * FROM RESUMOVENDACARTEIRA_" & usuarioLogado("id") 
    SQL = SQL & "   ORDER BY NOME_VENDEDOR, NOME_CLIENTE " 
    SQL = SQL & "   FOR JSON PATH"
    SQL = SQL & ") AS dados"

    On Error Resume Next
    SET rsDados = Conecta("SPO-BI", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDados) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar dados: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    dados = ""
    If Not rsDados.EOF Then
       dados = Iif(isnull(rsDados("dados")), "[]", rsDados("dados"))
    End If

    rsDados.Close()
    SET rsDados = NOTHING 

    // Dropa tabela temporária
    SQL = "DROP TABLE RESUMOVENDACARTEIRA_" & usuarioLogado("id")
    
    On Error Resume Next
    SET rsDropTable = Conecta("SPO-BI", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDropTable) Then
        Listar = "{""sucesso"": false, ""mensagem"": ""Erro ao remover tabela: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    Listar = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & dados & "}"
End Function


' Auxiliares
Function descricaoDaPlanta(codigo)
    Select Case codigo
        Case "00" descricaoDaPlanta = "PIRAPETINGA"
        Case "04" descricaoDaPlanta = "UBERABA"
        Case "05" descricaoDaPlanta = "BENTO"
        Case "06" descricaoDaPlanta = "MARANGUAPE"
        Case "07" descricaoDaPlanta = "SAQUAREMA"
        Case "10" descricaoDaPlanta = "ESHOP"
    End Select
End Function

Function parametrocomOuSemImposto(argcomOuSemImposto)
    Select Case argcomOuSemImposto
        Case "0" parametrocomOuSemImposto = "Com IPI"
        Case "1" parametrocomOuSemImposto = "Sem IPI"
        Case "2" parametrocomOuSemImposto = "Sem Imposto (ICM, PIS, COF)"
    End Select
End Function

Function parametrotipoProduto(argTipoProduto)
    Select Case argTipoProduto
        Case "0" parametrotipoProduto = "AMBOS"
        Case "1" parametrotipoProduto = "PAPEL"
        Case "2" parametrotipoProduto = "PAPELAO"
    End Select
End Function

Function parametrotipoPapelao(argTipoPapelao)
    Select Case argTipoPapelao
        Case "0" parametrotipoPapelao = "AMBOS"
        Case "1" parametrotipoPapelao = "CAIXA"
        Case "2" parametrotipoPapelao = "CHAPA"
        Case "3" parametrotipoPapelao = "CORTE VINCO"
        Case "4" parametrotipoPapelao = "MALETA"
    End Select
End Function

Function parametroValores(argValores)
    Select Case argValores
        Case "0" ParametroValores = "Todos"
        Case "1" ParametroValores = "Só Peso"
        Case "2" ParametroValores = "Só M2"
    End Select
End Function

Function parametroClienteGrupo(argClienteGrupo)
    Select Case argClienteGrupo
        Case "T" ParametroClienteGrupo = "Todos"
        Case "S" ParametroClienteGrupo = "Somente Grupo"
        Case "S" ParametroClienteGrupo = "Somente Clientes"
    End Select             
End Function 
%>