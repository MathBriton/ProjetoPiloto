<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../cnn2/cnn2.asp"-->
<!--#Include File="../../funcoes/funcoes.asp"-->
<!--#Include File="../../includes/variaveis.asp"-->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046
Response.ContentType = "application/json"

' Variáveis globais
Dim acao
Dim usuarioLogado
Set usuarioLogado = CreateObject("Scripting.Dictionary")
usuarioLogado.Add "id", Session("CODIGO")
usuarioLogado.Add "login", Session("USR")
usuarioLogado.Add "nome", Session("NOME")
usuarioLogado.Add "nivel", Session("NIVEL")
usuarioLogado.Add "cod", Session("CODUSER")

acao = Request("acao")

If acao = "" Then
    Response.Write("{""sucesso"": false, ""mensagem"": ""Ação não informada""}")
    Response.End()
End If

Dim resultado
Select Case acao
    Case "listar"
        resultado = Listar()
    Case "criar"
        resultado = Criar()
    Case "editar"
        resultado = Editar()
    Case "excluir"
        resultado = Excluir()
    Case "cancelar"
        resultado = Cancelar()
    Case "reverterCancelamento"
        resultado = ReverterCancelamento()
    Case "obterPorId"
        resultado = ObterPorId(Request("id"), Request("codPlanta"), 1)
    Case "alterarStatusAndamento"
        resultado = AlterarStatusAndamento()
    Case "aprovacaoComercial"
        resultado = AprovacaoComercial()
    Case "rejeicaoComercial"
        resultado = RejeicaoComercial()
    Case "aprovacaoDesenvolvimento"
        resultado = AprovacaoDesenvolvimento()
    Case "rejeicaoDesenvolvimento"
        resultado = RejeicaoDesenvolvimento()
    Case "aprovacaoPCP"
        resultado = AprovacaoPCP()
    Case "rejeicaoPCP"
        resultado = RejeicaoPCP()
    Case "editarTestesFisicos"
        resultado = EditarTestesFisicos()
    Case "listarFichasTecnicas"
        resultado = ListarFichasTecnicas()
    Case "listarDesenvolvedores"
        resultado = ListarDesenvolvedores()
    Case "listarEstilos"
        resultado = ListarEstilos()
    Case "listarEstruturasPapelao"
        resultado = ListarEstruturasPapelao()
    Case "listarConteudosEmbalado"
        resultado = ListarConteudosEmbalado()
    Case "listarTiposProdutoPrimario"
        resultado = ListarTiposProdutoPrimario()
    Case "listarPaletes"
        resultado = ListarPaletes()
    Case "listarStatusAndamento"
        resultado = ListarStatusAndamento()
    Case "listarMotivosReprovacaoAmostra"
        resultado = ListarMotivosReprovacaoAmostra()
    Case "listarMotivosRejeicao"
        resultado = ListarMotivosRejeicao()
    Case "listarProcessosAprovacao"
        resultado = ListarProcessosAprovacao()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Ação inválida: " & acao & """}"
End Select

' Retorna resposta
Response.Write(resultado)

'----------------------------------------------
'                MÉTODOS CRUD
'----------------------------------------------
Function Listar()
    id = Request("filtroIdSolicitacao")
    codPlanta = Request("filtroFilial")
    status = Request("filtroStatus")
    ftOrigem = Request("filtroFtOrigem")
    ftNova = Request("filtroFtNova")
    codCliente = Request("filtroCodCliente")
    codLoja = Request("filtroCodLoja")
    tipoSolicitacao = Request("filtroTipoSolicitacao")
    nivelUsuario = SESSION("NIVELSPD")
    andamentoProcesso = Request("filtroAndamentoProcesso")
    flagApenasPendentes = Request("flagApenasPendentes")

    SQL = "SELECT TOP 100 " 
    SQL = SQL & "s.FT1_ID AS Id, "
    SQL = SQL & "CONCAT(FORMAT(CAST(s.FT1_DATA AS DATE), 'dd/MM/yyyy'), ' ', s.FT1_HORA) AS DataEmissao, "
    SQL = SQL & "s.FT1_USU as LoginUsuarioCriacao, "
    SQL = SQL & "s.FT1_TIPO as TipoSolicitacao, "
    SQL = SQL & "s.FT1_STATUS as Status, "
    SQL = SQL & "s.FT1_PLANTA AS CodPlanta, "
    SQL = SQL & "s.FT1_B5CLI as CodCliente, "
    SQL = SQL & "s.FT1_LOJA as CodLoja, "
    SQL = SQL & "s.FT1_B5NOMECLI as NomeCliente, "
    SQL = SQL & "s.FT1_FT as FtOriginal, "
    SQL = SQL & "s.FT1_FTNOVA as FtNova, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), s.FT1_B5REF), CHAR(7), ''), CHAR(34), ''), CHAR(0), ''), CHAR(9), '') AS Referencia, "
    SQL = SQL & "s.FT1_B5TOLMA as ToleranciaMaxima, "
    SQL = SQL & "s.FT1_B5TOLME as ToleranciaMinima, "
    SQL = SQL & "s.FT1_B5ESTILO as Estilo, "
    SQL = SQL & "s.FT1_B1CONTA as Conta, "
    SQL = SQL & "s.FT1_B1DESCNF as DescNF, "
    SQL = SQL & "s.FT1_B1NATUREZ as Natureza, "
    SQL = SQL & "s.FT1_OBSP as Observacoes, "
    SQL = SQL & "s.FT1_B1INPUBE as UnidadeFabricacao, "
    SQL = SQL & "s.FT1_KMUNID as DistanciaUnidade, "
    SQL = SQL & "s.FT1_B5COMPR as Comprimento, "
    SQL = SQL & "s.FT1_B5LARG as Largura, "
    SQL = SQL & "s.FT1_B5ALTU as Altura, "
    SQL = SQL & "s.FT1_B5NUMFACA as NumeroMatriz, "
    SQL = SQL & "s.FT1_TARA as ClienteUsaTara, "
    SQL = SQL & "s.FT1_TOLPESO as ToleranciaPeso, "
    SQL = SQL & "s.FT1_B5ESTRUTU as EstruturaPapelao, "
    SQL = SQL & "s.FT1_INKJET as InkJet, "
    SQL = SQL & "s.FT1_B5GRAM as GramaturaReferencial, "
    SQL = SQL & "s.FT1_B5ESMAG as Esmagamento, "
    SQL = SQL & "s.FT1_B5COLUNA as EctColuna, "
    SQL = SQL & "s.FT1_MULLEN as MullenReferencial, "
    SQL = SQL & "s.FT1_B5ESPES2 as Espessura, "
    SQL = SQL & "s.FT1_B5COMPRE as BctCompressao, "
    SQL = SQL & "s.FT1_B5COBINT as CobbInterno, "
    SQL = SQL & "s.FT1_B5COBEXT as CobbExterno, "
    SQL = SQL & "s.FT1_B5FECHAM as TipoFechamento, "
    SQL = SQL & "s.FT1_B5PALETE as EhPaletizado, "
    SQL = SQL & "s.FT1_B5AMARRAD as EhAmarrado, "
    SQL = SQL & "s.FT1_ALTPAL as AlturaMaximaPalete, "
    SQL = SQL & "s.FT1_FITILH as QuantidadeFitilhos, "
    SQL = SQL & "s.FT1_FITPOS as TipoFitilho, "
    SQL = SQL & "s.FT1_B5QTDPCT as PecasPorUnidade, "
    SQL = SQL & "s.FT1_B5TPPALE as Palete, "
    SQL = SQL & "s.FT1_B5FACOP as TemFacesOpostas, "
    SQL = SQL & "s.FT1_B5CANTON as TemCantinho, "
    SQL = SQL & "s.FT1_B5TAMPA as TemGradeMadeira, "
    SQL = SQL & "s.FT1_B5FILME as TemFilmeStrech, "
    SQL = SQL & "s.FT1_RACK as ClienteUsaRack, "
    SQL = SQL & "s.FT1_OVERHANG as Overhang, "
    SQL = SQL & "COALESCE(s.FT1_B5ESPELHO, 0) as QuantidadeEspelhos, "
    SQL = SQL & "s.FT1_B1ALIMENT as ConteudoEmbalado, "
    SQL = SQL & "s.FT1_CONSMES as ConsumoMensal, "
    SQL = SQL & "s.FT1_LINAUTO as TipoLinhaProducaoCliente, "
    SQL = SQL & "s.FT1_TIPOPRD as TipoProdutoPrimario, "
    SQL = SQL & "s.FT1_PESOPRI as PesoPrimario, "
    SQL = SQL & "s.FT1_QTDPRI as QuantidadePrimarios, "
    SQL = SQL & "s.FT1_TEMPOEST as TempoEstocagem, "
    SQL = SQL & "s.FT1_UNIREL as UmidadeRelativa, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), s.FT1_OBSERV), CHAR(7), ''), CHAR(34), ''), CHAR(0), ''), CHAR(9), '') AS InformacoesAdicionais, "
    SQL = SQL & "s.FT1_QTDAMO as QuantidadeAmostras, "
    SQL = SQL & "s.FT1_CEP as AmostraCep, "
    SQL = SQL & "s.FT1_CON as AmostraResponsavel, "
    SQL = SQL & "s.FT1_END as AmostraEndereco, "
    SQL = SQL & "s.FT1_BAI as AmostraBairro, "
    SQL = SQL & "s.FT1_MUN as AmostraMunicipio, "
    SQL = SQL & "s.FT1_EST as AmostraUF, "
    SQL = SQL & "s.FT1_TEL as AmostraTelefone, "
    SQL = SQL & "s.FT1_CNPJ as AmostraCNPJ, "
    SQL = SQL & "s.FT1_EMAIL as AmostraEmail, "
    SQL = SQL & "s.FT1_ObservacaoStatus2 as AmostraObservacoes, "
    SQL = SQL & "s.FT1_FormaEnvioAmostra as AmostraFormaEnvio, "
    SQL = SQL & "s.FT1_MOTIVOREPROVACAOAMOSTRA as MotivoReprovacaoAmostra, "
    SQL = SQL & "s.FT1_MOTIVOREPROVACAOAMOSTRA2 as MotivoReprovacaoAmostra2, "
    SQL = SQL & "s.FT1_B5FITLAR as LarguraFitilho, "
    SQL = SQL & "s.FT1_B5FITCOM as ComprimentoFitilho, "
    SQL = SQL & "s.FT1_COR as Cor, "
    SQL = SQL & "s.FT1_ICMS as ICMS, "
    SQL = SQL & "s.FT1_IPI as IPI, "
    SQL = SQL & "s.FT1_TIPOPREVISTO as TipoPrevisto, "
    SQL = SQL & "s.FT1_ORCAMENTO as Orcamento, "
    SQL = SQL & "s.FT1_LARGURACHAPA as LarguraChapa, "
    SQL = SQL & "s.FT1_COMPRIMENTOCHAPA as ComprimentoChapa, "
    SQL = SQL & "s.FT1_TESTECOLUNA as TesteColuna, "
    SQL = SQL & "s.FT1_TESTEGRAMATURA as TesteGramatura, "
    SQL = SQL & "s.FT1_TESTECOBINTERNO as TesteCobInterno, "
    SQL = SQL & "s.FT1_TESTECOBEXTERNO as TesteCobExterno, "
    SQL = SQL & "s.FT1_TESTEMULLENINTERNOCAPA as TesteDescolamentoCapaInterna, "
    SQL = SQL & "s.FT1_TESTEMULLENINTERNOMIOLO as TesteDescolamentoMiolo1, "
    SQL = SQL & "s.FT1_TESTEMULLENINTERNOCAPAINTERMEDIARIA as TesteDescolamentoMiolo2, "
    SQL = SQL & "s.FT1_TESTEMULLENEXTERNOCAPA as TesteDescolamentoMiolo3, "
    SQL = SQL & "s.FT1_TESTEMULLENEXTERNOMIOLO as TesteDescolamentoCapaExterna, "
    SQL = SQL & "s.FT1_COMPRESSAO1 as TesteBctCompressao1, "
    SQL = SQL & "s.FT1_COMPRESSAO2 as TesteBctCompressao2, "
    SQL = SQL & "s.FT1_COMPRESSAO3 as TesteBctCompressao3, "
    SQL = SQL & "s.FT1_PININTERNO as TestePinInterno, "
    SQL = SQL & "s.FT1_PINEXTERNO as TestePinExterno, "
    SQL = SQL & "s.FT1_FORCAVINCO1 as TesteForcaVinco1, "
    SQL = SQL & "s.FT1_FORCAVINCO2 as TesteForcaVinco2, "
    SQL = SQL & "s.FT1_FORCAVINCO3 as TesteForcaVinco3, "
    SQL = SQL & "s.FT1_TipoAlteraCancela as Processo, "
    SQL = SQL & "s.FT1_TECNICO as IdDesenvolvedor, "
    SQL = SQL & "s.FT1_Margem as MargemContribuicao, "
    SQL = SQL & "s.FT1_Resultado as Resultado, "
    SQL = SQL & "s.FT1_QTDCOR as QuantidadeCores, "
    SQL = SQL & "s.FT1_MOTCANC as JustificativaRejeicao, "
    SQL = SQL & "s.FT1_IDMunicipio as IdMunicipioCliente, "
    SQL = SQL & "s.FT1_Z1GRAMC as Gramatura, "
    SQL = SQL & "s.FT1_SAVING as TipoSavings, "
    SQL = SQL & "s.FT1_SAVINGVALK as ValorSavingSK, "
    SQL = SQL & "s.FT1_SAVINGVALC as ValorSavingCliente, "
    SQL = SQL & "s.FT1_A3NREDUZ as NomeVendedor, "
    SQL = SQL & "s.FT1_A1VEND as IdVendedor, "
    SQL = SQL & "s.FT1_STATUS2 as StatusAndamento, "
    SQL = SQL & "s.FT1_DTCONFEC as DataConfeccaoFerramental, "
    SQL = SQL & "s.FT1_AcompanharLotePiloto as NecessitaAcompanhamento, "
    SQL = SQL & "s.FT1_B5AREA as AreaLiquida, "
    SQL = SQL & "s.FT1_SUPORTEPRIMARIO as SuportacaoPrimario, "
    SQL = SQL & "s.FT1_PRECOKG as PrecoKG, "
    SQL = SQL & "s.FT1_VALORCLICHE AS ValorEstimadoCliche, "
    SQL = SQL & "s.FT1_VALORMATRIZ as ValorEstimadoMatriz, "
    SQL = SQL & "s.FT1_NIVELLOTEPILOTO as NivelLotePiloto, "
    SQL = SQL & "s.FT1_PRECADASTRO as EhPrecadastro, "
    SQL = SQL & "s.FT1_DEL as Deletado, "
    SQL = SQL & "("
	SQL = SQL &	"   SELECT "
    SQL = SQL & "   IdHistorico, "
	SQL = SQL & "   IdProcesso, "
	SQL = SQL & "   IsNull(DataInsercaoProcesso,'') DataInsercaoProcesso, "
	SQL = SQL & "   DataAlteracaoProcesso, "
    SQL = SQL & "   DATEDIFF( "
    SQL = SQL & "       DAY, "
    SQL = SQL & "       DataInsercaoProcesso, "
    SQL = SQL & "       CASE WHEN DataAlteracaoProcesso IS NULL THEN GETDATE() ELSE DataAlteracaoProcesso END "
    SQL = SQL & "   ) AS LeadTime, "
	SQL = SQL & "   QuemAlterou, "
	SQL = SQL & "   IsNull(Complemento,'') Complemento "
	SQL = SQL & "   FROM HistoricoAlteracaoProcessoSap "
	SQL = SQL & "   WHERE IdSap = s.FT1_ID "
    SQL = SQL & "   FOR JSON PATH"
    SQL = SQL & ") AS alteracoes "
    SQL = SQL & "FROM SKFT1 s "
    SQL = SQL & "WHERE FT1_DEL = 0 "
    
    If id <> "" Then
        SQL = SQL & "AND FT1_ID = '" & id & "' "
    End If    

    If codPlanta <> "" Then
        SQL = SQL & "AND FT1_PLANTA = '" & codPlanta & "' "
    End If

    If status <> "" Then
        SQL = SQL & "AND FT1_STATUS = '" & status & "' "
    End If

    If ftNova <> "" then
        SQL = SQL & "AND (FT1_FTNOVA LIKE '%" & UCase(ftNova) & "%' OR FT1_B5REF LIKE '%" & ftNova & "%') "
    End If

    If ftOrigem <> "" then
        SQL = SQL & "AND (FT1_FT = '" & UCase(ftOrigem) & "' OR FT1_B5REF LIKE '%" & ftOrigem & "%') "
    End If

    If codCliente <> "" then
        SQL = SQL & "AND FT1_B5CLI = '" & codCliente & "' "
    End If

    If usuarioLogado("nivel") = 1 Then
        SQL = SQL & "AND FT1_A1VEND = '" & usuarioLogado("cod") & "' "
    End If

    If flagApenasPendentes = "1" Then
        If nivelUsuario = "D" Then
            ' Desenvolvimento
            SQL = SQL & "AND FT1_STATUS = 'A' "
        ElseIf nivelUsuario = "C" Then
            ' Comercial
            SQL = SQL & "AND FT1_STATUS = 'S' "
        ElseIf nivelUsuario = "P" Then
            ' PCP
            SQL = SQL & "AND FT1_STATUS = 'D' "
        End If
    End If

    If andamentoProcesso <> "" then
        SQL = SQL & "AND FT1_STATUS2 = '" & andamentoProcesso & "' "
    End If

    SQL = SQL & "ORDER BY FT1_ID DESC"

    On Error Resume Next
    SET rsListar = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        ListarFNCs = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de solicitações: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsListar.EOF Then
        Do While Not rsListar.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If

            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsListar("Id")) & """, "
            jsonArray = jsonArray & """alteracoes"": " & Iif(IsNull(rsListar("alteracoes")), "[]", rsListar("alteracoes"))  & ", "
            jsonArray = jsonArray & """dataEmissao"": """ & EscapeJSON(rsListar("DataEmissao")) & """, "
            jsonArray = jsonArray & """loginUsuarioCriacao"": """ & EscapeJSON(rsListar("LoginUsuarioCriacao")) & """, "
            jsonArray = jsonArray & """tipoSolicitacao"": """ & EscapeJSON(rsListar("TipoSolicitacao")) & """, "
            jsonArray = jsonArray & """status"": """ & EscapeJSON(rsListar("Status")) & """, "
            jsonArray = jsonArray & """codPlanta"": """ & EscapeJSON(rsListar("CodPlanta")) & """, "
            jsonArray = jsonArray & """codCliente"": """ & EscapeJSON(rsListar("CodCliente")) & """, "
            jsonArray = jsonArray & """codLoja"": """ & EscapeJSON(rsListar("CodLoja")) & """, "
            jsonArray = jsonArray & """nomeCliente"": """ & EscapeJSON(rsListar("NomeCliente")) & """, "
            jsonArray = jsonArray & """ftOriginal"": """ & EscapeJSON(rsListar("FtOriginal")) & """, "
            jsonArray = jsonArray & """ftNova"": """ & EscapeJSON(rsListar("FtNova")) & """, "
            jsonArray = jsonArray & """referencia"": """ & EscapeJSON(rsListar("Referencia")) & """, "
            jsonArray = jsonArray & """toleranciaMaxima"": """ & EscapeJSON(rsListar("ToleranciaMaxima")) & """, "
            jsonArray = jsonArray & """toleranciaMinima"": """ & EscapeJSON(rsListar("ToleranciaMinima")) & """, "
            jsonArray = jsonArray & """estilo"": """ & EscapeJSON(rsListar("Estilo")) & """, "
            jsonArray = jsonArray & """conta"": """ & EscapeJSON(rsListar("Conta")) & """, "
            jsonArray = jsonArray & """descNF"": """ & EscapeJSON(rsListar("DescNF")) & """, "
            jsonArray = jsonArray & """natureza"": """ & EscapeJSON(rsListar("Natureza")) & """, "
            jsonArray = jsonArray & """observacoes"": """ & EscapeJSON(rsListar("Observacoes")) & """, "
            jsonArray = jsonArray & """unidadeFabricacao"": """ & EscapeJSON(rsListar("UnidadeFabricacao")) & """, "
            jsonArray = jsonArray & """distanciaUnidade"": """ & EscapeJSON(rsListar("DistanciaUnidade")) & """, "
            jsonArray = jsonArray & """comprimento"": """ & EscapeJSON(rsListar("Comprimento")) & """, "
            jsonArray = jsonArray & """largura"": """ & EscapeJSON(rsListar("Largura")) & """, "
            jsonArray = jsonArray & """altura"": """ & EscapeJSON(rsListar("Altura")) & """, "
            jsonArray = jsonArray & """numeroMatriz"": """ & EscapeJSON(rsListar("NumeroMatriz")) & """, "
            jsonArray = jsonArray & """clienteUsaTara"": """ & EscapeJSON(rsListar("ClienteUsaTara")) & """, "
            jsonArray = jsonArray & """toleranciaPeso"": """ & EscapeJSON(rsListar("ToleranciaPeso")) & """, "
            jsonArray = jsonArray & """estruturaPapelao"": """ & EscapeJSON(rsListar("EstruturaPapelao")) & """, "
            jsonArray = jsonArray & """inkJet"": """ & EscapeJSON(rsListar("InkJet")) & """, "
            jsonArray = jsonArray & """gramaturaReferencial"": """ & EscapeJSON(rsListar("GramaturaReferencial")) & """, "
            jsonArray = jsonArray & """esmagamento"": """ & EscapeJSON(rsListar("Esmagamento")) & """, "
            jsonArray = jsonArray & """ectColuna"": """ & EscapeJSON(rsListar("EctColuna")) & """, "
            jsonArray = jsonArray & """mullenReferencial"": """ & EscapeJSON(rsListar("MullenReferencial")) & """, "
            jsonArray = jsonArray & """espessura"": """ & EscapeJSON(rsListar("Espessura")) & """, "
            jsonArray = jsonArray & """bctCompressao"": """ & EscapeJSON(rsListar("BctCompressao")) & """, "
            jsonArray = jsonArray & """cobbInterno"": """ & EscapeJSON(rsListar("CobbInterno")) & """, "
            jsonArray = jsonArray & """cobbExterno"": """ & EscapeJSON(rsListar("CobbExterno")) & """, "
            jsonArray = jsonArray & """tipoFechamento"": """ & EscapeJSON(rsListar("TipoFechamento")) & """, "
            jsonArray = jsonArray & """ehPaletizado"": """ & EscapeJSON(rsListar("EhPaletizado")) & """, "
            jsonArray = jsonArray & """ehAmarrado"": """ & EscapeJSON(rsListar("EhAmarrado")) & """, "
            jsonArray = jsonArray & """alturaMaximaPalete"": """ & EscapeJSON(rsListar("AlturaMaximaPalete")) & """, "
            jsonArray = jsonArray & """quantidadeFitilhos"": """ & EscapeJSON(rsListar("QuantidadeFitilhos")) & """, "
            jsonArray = jsonArray & """tipoFitilho"": """ & EscapeJSON(rsListar("TipoFitilho")) & """, "
            jsonArray = jsonArray & """pecasPorUnidade"": """ & EscapeJSON(rsListar("PecasPorUnidade")) & """, "
            jsonArray = jsonArray & """palete"": """ & EscapeJSON(rsListar("Palete")) & """, "
            jsonArray = jsonArray & """temFacesOpostas"": """ & EscapeJSON(rsListar("TemFacesOpostas")) & """, "
            jsonArray = jsonArray & """temCantinho"": """ & EscapeJSON(rsListar("TemCantinho")) & """, "
            jsonArray = jsonArray & """temGradeMadeira"": """ & EscapeJSON(rsListar("TemGradeMadeira")) & """, "
            jsonArray = jsonArray & """temFilmeStrech"": """ & EscapeJSON(rsListar("TemFilmeStrech")) & """, "
            jsonArray = jsonArray & """clienteUsaRack"": """ & EscapeJSON(rsListar("ClienteUsaRack")) & """, "
            jsonArray = jsonArray & """overhang"": """ & EscapeJSON(rsListar("Overhang")) & """, "
            jsonArray = jsonArray & """quantidadeEspelhos"": """ & EscapeJSON(rsListar("QuantidadeEspelhos")) & """, "
            jsonArray = jsonArray & """conteudoEmbalado"": """ & EscapeJSON(rsListar("ConteudoEmbalado")) & """, "
            jsonArray = jsonArray & """consumoMensal"": """ & EscapeJSON(rsListar("ConsumoMensal")) & """, "
            jsonArray = jsonArray & """tipoLinhaProducaoCliente"": """ & EscapeJSON(rsListar("TipoLinhaProducaoCliente")) & """, "
            jsonArray = jsonArray & """tipoProdutoPrimario"": """ & EscapeJSON(rsListar("TipoProdutoPrimario")) & """, "
            jsonArray = jsonArray & """pesoPrimario"": """ & EscapeJSON(rsListar("PesoPrimario")) & """, "
            jsonArray = jsonArray & """quantidadePrimarios"": """ & EscapeJSON(rsListar("QuantidadePrimarios")) & """, "
            jsonArray = jsonArray & """tempoEstocagem"": """ & EscapeJSON(rsListar("TempoEstocagem")) & """, "
            jsonArray = jsonArray & """umidadeRelativa"": """ & EscapeJSON(rsListar("UmidadeRelativa")) & """, "
            jsonArray = jsonArray & """informacoesAdicionais"": """ & EscapeJSON(rsListar("InformacoesAdicionais")) & """, "
            jsonArray = jsonArray & """quantidadeAmostras"": """ & EscapeJSON(rsListar("QuantidadeAmostras")) & """, "
            jsonArray = jsonArray & """amostraCep"": """ & EscapeJSON(rsListar("AmostraCep")) & """, "
            jsonArray = jsonArray & """amostraResponsavel"": """ & EscapeJSON(rsListar("AmostraResponsavel")) & """, "
            jsonArray = jsonArray & """amostraEndereco"": """ & EscapeJSON(rsListar("AmostraEndereco")) & """, "
            jsonArray = jsonArray & """amostraBairro"": """ & EscapeJSON(rsListar("AmostraBairro")) & """, "
            jsonArray = jsonArray & """amostraMunicipio"": """ & EscapeJSON(rsListar("AmostraMunicipio")) & """, "
            jsonArray = jsonArray & """amostraUF"": """ & EscapeJSON(rsListar("AmostraUF")) & """, "
            jsonArray = jsonArray & """amostraTelefone"": """ & EscapeJSON(rsListar("AmostraTelefone")) & """, "
            jsonArray = jsonArray & """amostraCNPJ"": """ & EscapeJSON(rsListar("AmostraCNPJ")) & """, "
            jsonArray = jsonArray & """amostraEmail"": """ & EscapeJSON(rsListar("AmostraEmail")) & """, "
            jsonArray = jsonArray & """statusAndamento"": """ & EscapeJSON(rsListar("StatusAndamento")) & """, "
            jsonArray = jsonArray & """amostraObservacoes"": """ & EscapeJSON(rsListar("AmostraObservacoes")) & """, "
            jsonArray = jsonArray & """amostraFormaEnvio"": """ & EscapeJSON(rsListar("AmostraFormaEnvio")) & """, "
            jsonArray = jsonArray & """motivoReprovacaoAmostra"": """ & EscapeJSON(rsListar("MotivoReprovacaoAmostra")) & """, "
            jsonArray = jsonArray & """motivoReprovacaoAmostra2"": """ & EscapeJSON(rsListar("MotivoReprovacaoAmostra2")) & """, "
            jsonArray = jsonArray & """larguraFitilho"": """ & EscapeJSON(rsListar("LarguraFitilho")) & """, "
            jsonArray = jsonArray & """comprimentoFitilho"": """ & EscapeJSON(rsListar("ComprimentoFitilho")) & """, "
            jsonArray = jsonArray & """cor"": """ & EscapeJSON(rsListar("Cor")) & """, "
            jsonArray = jsonArray & """icms"": """ & EscapeJSON(rsListar("ICMS")) & """, "
            jsonArray = jsonArray & """ipi"": """ & EscapeJSON(rsListar("IPI")) & """, "
            jsonArray = jsonArray & """tipoPrevisto"": """ & EscapeJSON(rsListar("TipoPrevisto")) & """, "
            jsonArray = jsonArray & """orcamento"": """ & EscapeJSON(rsListar("Orcamento")) & """, "
            jsonArray = jsonArray & """larguraChapa"": """ & EscapeJSON(rsListar("LarguraChapa")) & """, "
            jsonArray = jsonArray & """comprimentoChapa"": """ & EscapeJSON(rsListar("ComprimentoChapa")) & """, "
            jsonArray = jsonArray & """testeColuna"": """ & EscapeJSON(rsListar("TesteColuna")) & """, "
            jsonArray = jsonArray & """testeGramatura"": """ & EscapeJSON(rsListar("TesteGramatura")) & """, "
            jsonArray = jsonArray & """testeCobInterno"": """ & EscapeJSON(rsListar("TesteCobInterno")) & """, "
            jsonArray = jsonArray & """testeCobExterno"": """ & EscapeJSON(rsListar("TesteCobExterno")) & """, "
            jsonArray = jsonArray & """testeDescolamentoCapaInterna"": """ & EscapeJSON(rsListar("TesteDescolamentoCapaInterna")) & """, "
            jsonArray = jsonArray & """testeDescolamentoMiolo1"": """ & EscapeJSON(rsListar("TesteDescolamentoMiolo1")) & """, "
            jsonArray = jsonArray & """testeDescolamentoMiolo2"": """ & EscapeJSON(rsListar("TesteDescolamentoMiolo2")) & """, "
            jsonArray = jsonArray & """testeDescolamentoMiolo3"": """ & EscapeJSON(rsListar("TesteDescolamentoMiolo3")) & """, "
            jsonArray = jsonArray & """testeDescolamentoCapaExterna"": """ & EscapeJSON(rsListar("TesteDescolamentoCapaExterna")) & """, "
            jsonArray = jsonArray & """testeBctCompressao1"": """ & EscapeJSON(rsListar("TesteBctCompressao1")) & """, "
            jsonArray = jsonArray & """testeBctCompressao2"": """ & EscapeJSON(rsListar("TesteBctCompressao2")) & """, "
            jsonArray = jsonArray & """testeBctCompressao3"": """ & EscapeJSON(rsListar("TesteBctCompressao3")) & """, "
            jsonArray = jsonArray & """processo"": """ & EscapeJSON(rsListar("Processo")) & """, "
            jsonArray = jsonArray & """idDesenvolvedor"": """ & EscapeJSON(rsListar("IdDesenvolvedor")) & """, "
            jsonArray = jsonArray & """margemContribuicao"": """ & EscapeJSON(rsListar("MargemContribuicao")) & """, "
            jsonArray = jsonArray & """resultado"": """ & EscapeJSON(rsListar("Resultado")) & """, "
            jsonArray = jsonArray & """quantidadeCores"": """ & EscapeJSON(rsListar("QuantidadeCores")) & """, "
            jsonArray = jsonArray & """justificativaRejeicao"": """ & EscapeJSON(rsListar("JustificativaRejeicao")) & """, "
            jsonArray = jsonArray & """idMunicipioCliente"": """ & EscapeJSON(rsListar("IdMunicipioCliente")) & """, "
            jsonArray = jsonArray & """gramatura"": """ & EscapeJSON(rsListar("Gramatura")) & """, "
            jsonArray = jsonArray & """tipoSavings"": """ & EscapeJSON(rsListar("TipoSavings")) & """, "
            jsonArray = jsonArray & """valorSavingSK"": """ & EscapeJSON(rsListar("ValorSavingSK")) & """, "
            jsonArray = jsonArray & """valorSavingCliente"": """ & EscapeJSON(rsListar("ValorSavingCliente")) & """, "
            jsonArray = jsonArray & """nomeVendedor"": """ & EscapeJSON(rsListar("NomeVendedor")) & """, "
            jsonArray = jsonArray & """idVendedor"": """ & EscapeJSON(rsListar("IdVendedor")) & """, "
            jsonArray = jsonArray & """statusAndamento"": """ & EscapeJSON(rsListar("StatusAndamento")) & """, "
            jsonArray = jsonArray & """dataConfeccaoFerramental"": """ & EscapeJSON(rsListar("DataConfeccaoFerramental")) & """, "
            jsonArray = jsonArray & """necessitaAcompanhamento"": """ & EscapeJSON(rsListar("NecessitaAcompanhamento")) & """, "
            jsonArray = jsonArray & """areaLiquida"": """ & EscapeJSON(rsListar("AreaLiquida")) & """, "
            jsonArray = jsonArray & """suportacaoPrimario"": """ & EscapeJSON(rsListar("SuportacaoPrimario")) & """, "
            jsonArray = jsonArray & """precoKG"": """ & EscapeJSON(rsListar("PrecoKG")) & """, "
            jsonArray = jsonArray & """valorEstimadoCliche"": """ & EscapeJSON(rsListar("ValorEstimadoCliche")) & """, "
            jsonArray = jsonArray & """valorEstimadoMatriz"": """ & EscapeJSON(rsListar("ValorEstimadoMatriz")) & """, "
            jsonArray = jsonArray & """nivelLotePiloto"": """ & EscapeJSON(rsListar("NivelLotePiloto")) & """, "
            jsonArray = jsonArray & """testePinInterno"": """ & EscapeJSON(rsListar("TestePinInterno")) & """, "
            jsonArray = jsonArray & """testePinExterno"": """ & EscapeJSON(rsListar("TestePinExterno")) & """, "
            jsonArray = jsonArray & """testeForcaVinco1"": """ & EscapeJSON(rsListar("TesteForcaVinco1")) & """, "
            jsonArray = jsonArray & """testeForcaVinco2"": """ & EscapeJSON(rsListar("TesteForcaVinco2")) & """, "
            jsonArray = jsonArray & """testeForcaVinco3"": """ & EscapeJSON(rsListar("TesteForcaVinco3")) & """, "
            jsonArray = jsonArray & """ehPrecadastro"": """ & EscapeJSON(rsListar("EhPrecadastro")) & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsListar.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsListar.Close()
    SET rsListar = NOTHING
    
    Listar = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ObterPorId(id, codPlanta, carregarRelacionamentos)
    If id = "" Then
        Response.Write("{""sucesso"": false, ""mensagem"": ""Id da solicitação não informado""}")
        Exit Function
    End If

    If codPlanta = "" Then
        Response.Write("{""sucesso"": false, ""mensagem"": ""Planta nao informada""}")
        Exit Function
    End If

    SQL = "SELECT " 
    SQL = SQL & "s.FT1_ID AS Id, "
    SQL = SQL & "Len(Replace(Replace(Substring(s.FT1_B5ESTRUTU, 3, 2),'0',''),'O','')) TipoOnda,"
    SQL = SQL & "s.FT1_DATA AS Data, "
    SQL = SQL & "s.FT1_HORA AS Hora, "
    SQL = SQL & "CONCAT(FORMAT(CAST(s.FT1_DATA AS DATE), 'dd/MM/yyyy'), ' ', s.FT1_HORA) AS DataEmissao, "
    SQL = SQL & "s.FT1_USU as LoginUsuarioCriacao, "
    SQL = SQL & "s.FT1_TIPO as TipoSolicitacao, "
    SQL = SQL & "s.FT1_STATUS as Status, "
    SQL = SQL & "s.FT1_PLANTA AS CodPlanta, "
    SQL = SQL & "s.FT1_B5CLI as CodCliente, "
    SQL = SQL & "s.FT1_LOJA as CodLoja, "
    SQL = SQL & "s.FT1_B5NOMECLI as NomeCliente, "
    SQL = SQL & "s.FT1_FT as FtOriginal, "
    SQL = SQL & "s.FT1_FTNOVA as FtNova, "
    SQL = SQL & "REPLACE(REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), s.FT1_B5REF), CHAR(7), ''), CHAR(34), ''), CHAR(0), ''), CHAR(9), '') AS Referencia, "
    SQL = SQL & "s.FT1_B5TOLMA as ToleranciaMaxima, "
    SQL = SQL & "s.FT1_B5TOLME as ToleranciaMinima, "
    SQL = SQL & "s.FT1_B5ESTILO as Estilo, "
    SQL = SQL & "s.FT1_B1CONTA as Conta, "
    SQL = SQL & "s.FT1_B1DESCNF as DescNF, "
    SQL = SQL & "s.FT1_B1NATUREZ as Natureza, "
    SQL = SQL & "s.FT1_OBSP as Observacoes, "
    SQL = SQL & "s.FT1_B1INPUBE as UnidadeFabricacao, "
    SQL = SQL & "s.FT1_KMUNID as DistanciaUnidade, "
    SQL = SQL & "s.FT1_B5COMPR as Comprimento, "
    SQL = SQL & "s.FT1_B5LARG as Largura, "
    SQL = SQL & "s.FT1_B5ALTU as Altura, "
    SQL = SQL & "s.FT1_B5NUMFACA as NumeroMatriz, "
    SQL = SQL & "s.FT1_TARA as ClienteUsaTara, "
    SQL = SQL & "s.FT1_TOLPESO as ToleranciaPeso, "
    SQL = SQL & "s.FT1_B5ESTRUTU as EstruturaPapelao, "
    SQL = SQL & "s.FT1_INKJET as InkJet, "
    SQL = SQL & "s.FT1_B5GRAM as GramaturaReferencial, "
    SQL = SQL & "s.FT1_B5ESMAG as Esmagamento, "
    SQL = SQL & "s.FT1_B5COLUNA as EctColuna, "
    SQL = SQL & "s.FT1_MULLEN as MullenReferencial, "
    SQL = SQL & "s.FT1_B5ESPES2 as Espessura, "
    SQL = SQL & "s.FT1_B5COMPRE as BctCompressao, "
    SQL = SQL & "s.FT1_B5COBINT as CobbInterno, "
    SQL = SQL & "s.FT1_B5COBEXT as CobbExterno, "
    SQL = SQL & "s.FT1_B5FECHAM as TipoFechamento, "
    SQL = SQL & "s.FT1_B5PALETE as EhPaletizado, "
    SQL = SQL & "s.FT1_B5AMARRAD as EhAmarrado, "
    SQL = SQL & "s.FT1_ALTPAL as AlturaMaximaPalete, "
    SQL = SQL & "s.FT1_FITILH as QuantidadeFitilhos, "
    SQL = SQL & "s.FT1_FITPOS as TipoFitilho, "
    SQL = SQL & "s.FT1_B5QTDPCT as PecasPorUnidade, "
    SQL = SQL & "s.FT1_B5TPPALE as Palete, "
    SQL = SQL & "s.FT1_B5FACOP as TemFacesOpostas, "
    SQL = SQL & "s.FT1_B5CANTON as TemCantinho, "
    SQL = SQL & "s.FT1_B5TAMPA as TemGradeMadeira, "
    SQL = SQL & "s.FT1_B5FILME as TemFilmeStrech, "
    SQL = SQL & "s.FT1_RACK as ClienteUsaRack, "
    SQL = SQL & "s.FT1_OVERHANG as Overhang, "
    SQL = SQL & "COALESCE(s.FT1_B5ESPELHO, 0) as QuantidadeEspelhos, "
    SQL = SQL & "s.FT1_B1ALIMENT as ConteudoEmbalado, "
    SQL = SQL & "s.FT1_CONSMES as ConsumoMensal, "
    SQL = SQL & "s.FT1_LINAUTO as TipoLinhaProducaoCliente, "
    SQL = SQL & "s.FT1_TIPOPRD as TipoProdutoPrimario, "
    SQL = SQL & "s.FT1_PESOPRI as PesoPrimario, "
    SQL = SQL & "s.FT1_QTDPRI as QuantidadePrimarios, "
    SQL = SQL & "s.FT1_TEMPOEST as TempoEstocagem, "
    SQL = SQL & "s.FT1_UNIREL as UmidadeRelativa, "
    SQL = SQL & "s.FT1_OBSERV as InformacoesAdicionais, "
    SQL = SQL & "s.FT1_QTDAMO as QuantidadeAmostras, "
    SQL = SQL & "s.FT1_CEP as AmostraCep, "
    SQL = SQL & "s.FT1_CON as AmostraResponsavel, "
    SQL = SQL & "s.FT1_END as AmostraEndereco, "
    SQL = SQL & "s.FT1_BAI as AmostraBairro, "
    SQL = SQL & "s.FT1_MUN as AmostraMunicipio, "
    SQL = SQL & "s.FT1_EST as AmostraUF, "
    SQL = SQL & "s.FT1_TEL as AmostraTelefone, "
    SQL = SQL & "s.FT1_CNPJ as AmostraCNPJ, "
    SQL = SQL & "s.FT1_EMAIL as AmostraEmail, "
    SQL = SQL & "s.FT1_ObservacaoStatus2 as AmostraObservacoes, "
    SQL = SQL & "s.FT1_FormaEnvioAmostra as AmostraFormaEnvio, "
    SQL = SQL & "s.FT1_MOTIVOREPROVACAOAMOSTRA as MotivoReprovacaoAmostra, "
    SQL = SQL & "s.FT1_MOTIVOREPROVACAOAMOSTRA2 as MotivoReprovacaoAmostra2, "
    SQL = SQL & "s.FT1_B5FITLAR as LarguraFitilho, "
    SQL = SQL & "s.FT1_B5FITCOM as ComprimentoFitilho, "
    SQL = SQL & "s.FT1_COR as Cor, "
    SQL = SQL & "s.FT1_ICMS as ICMS, "
    SQL = SQL & "s.FT1_IPI as IPI, "
    SQL = SQL & "s.FT1_TIPOPREVISTO as TipoPrevisto, "
    SQL = SQL & "CASE WHEN s.FT1_ORCAMENTO = '1' THEN '1' ELSE '0' END as Orcamento, "
    SQL = SQL & "s.FT1_LARGURACHAPA as LarguraChapa, "
    SQL = SQL & "s.FT1_COMPRIMENTOCHAPA as ComprimentoChapa, "
    SQL = SQL & "s.FT1_TESTECOLUNA as TesteColuna, "
    SQL = SQL & "s.FT1_TESTEGRAMATURA as TesteGramatura, "
    SQL = SQL & "s.FT1_TESTECOBINTERNO as TesteCobInterno, "
    SQL = SQL & "s.FT1_TESTECOBEXTERNO as TesteCobExterno, "
    SQL = SQL & "s.FT1_TESTEMULLENINTERNOCAPA as TesteDescolamentoCapaInterna, "
    SQL = SQL & "s.FT1_TESTEMULLENINTERNOMIOLO as TesteDescolamentoMiolo1, "
    SQL = SQL & "s.FT1_TESTEMULLENINTERNOCAPAINTERMEDIARIA as TesteDescolamentoMiolo2, "
    SQL = SQL & "s.FT1_TESTEMULLENEXTERNOCAPA as TesteDescolamentoMiolo3, "
    SQL = SQL & "s.FT1_TESTEMULLENEXTERNOMIOLO as TesteDescolamentoCapaExterna, "
    SQL = SQL & "s.FT1_COMPRESSAO1 as TesteBctCompressao1, "
    SQL = SQL & "s.FT1_COMPRESSAO2 as TesteBctCompressao2, "
    SQL = SQL & "s.FT1_COMPRESSAO3 as TesteBctCompressao3, "
    SQL = SQL & "s.FT1_PININTERNO as TestePinInterno, "
    SQL = SQL & "s.FT1_PINEXTERNO as TestePinExterno, "
    SQL = SQL & "s.FT1_FORCAVINCO1 as TesteForcaVinco1, "
    SQL = SQL & "s.FT1_FORCAVINCO2 as TesteForcaVinco2, "
    SQL = SQL & "s.FT1_FORCAVINCO3 as TesteForcaVinco3, "
    SQL = SQL & "s.FT1_TipoAlteraCancela as Processo, "
    SQL = SQL & "s.FT1_TECNICO as IdDesenvolvedor, "
    SQL = SQL & "s.FT1_Margem as MargemContribuicao, "
    SQL = SQL & "s.FT1_Resultado as Resultado, "
    SQL = SQL & "s.FT1_QTDCOR as QuantidadeCores, "
    SQL = SQL & "s.FT1_MOTCANC as JustificativaRejeicao, "
    SQL = SQL & "s.FT1_IDMunicipio as IdMunicipioCliente, "
    SQL = SQL & "s.FT1_Z1GRAMC as Gramatura, "
    SQL = SQL & "s.FT1_SAVING as TipoSavings, "
    SQL = SQL & "s.FT1_SAVINGVALK as ValorSavingSK, "
    SQL = SQL & "s.FT1_SAVINGVALC as ValorSavingCliente, "
    SQL = SQL & "s.FT1_A3NREDUZ as NomeVendedor, "
    SQL = SQL & "s.FT1_A1VEND as IdVendedor, "
    SQL = SQL & "s.FT1_STATUS2 as StatusAndamento, "
    SQL = SQL & "s.FT1_DTCONFEC as DataConfeccaoFerramental, "
    SQL = SQL & "s.FT1_AcompanharLotePiloto as NecessitaAcompanhamento, "
    SQL = SQL & "s.FT1_B5AREA as AreaLiquida, "
    SQL = SQL & "s.FT1_SUPORTEPRIMARIO as SuportacaoPrimario, "
    SQL = SQL & "s.FT1_PRECOKG as PrecoKG, "
    SQL = SQL & "s.FT1_VALORCLICHE AS ValorEstimadoCliche, "
    SQL = SQL & "s.FT1_VALORMATRIZ as ValorEstimadoMatriz, "
    SQL = SQL & "s.FT1_NIVELLOTEPILOTO as NivelLotePiloto, "
    SQL = SQL & "s.FT1_PRECADASTRO as EhPrecadastro, "
    SQL = SQL & "s.FT1_DataEnvioAmostra as DataEnvioAmostra, "
    SQL = SQL & "s.FT1_QuemEnviouAmostra as QuemEnviouAmostra, "
    SQL = SQL & "s.FT1_DEL as Deletado "
    SQL = SQL & "FROM SKFT1 s "
    SQL = SQL & "WHERE FT1_DEL = 0 "
    SQL = SQL & "AND FT1_ID = " & id

    On Error Resume Next
    SET rsObter = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObter) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    ' Busca lista de acessórios
    SQL = "SELECT FT2_ID AS Id, FT2_ITEM as Item, FT2_ACEQTD as Quantidade, "
    SQL = SQL & "FT2_ACEFEFCO as Estilo, FT2_ACECORTE as NumeroCortes, FT2_ACELARG as Largura, "
    SQL = SQL & "FT2_ACECOMP as Comprimento, FT2_ACEESTR as Estrutura " 
    SQL = SQL & "FROM SKFT2 s WHERE FT2_ID = " & id
    
    On Error Resume Next
    SET rsAcessorios = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsAcessorios) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    jsonAcessorios = "["
    primeiroItem = True

    If Not rsAcessorios.EOF Then
        Do While Not rsAcessorios.EOF
            If Not primeiroItem Then jsonAcessorios = jsonAcessorios & "," End If
            
            jsonAcessorios = jsonAcessorios & "{"
            jsonAcessorios = jsonAcessorios & """item"": """ & EscapeJSON(rsAcessorios("Item")) & """, "
            jsonAcessorios = jsonAcessorios & """quantidade"": """ & EscapeJSON(rsAcessorios("Quantidade")) & """, "
            jsonAcessorios = jsonAcessorios & """estilo"": """ & EscapeJSON(rsAcessorios("Estilo")) & """, "
            jsonAcessorios = jsonAcessorios & """numeroCortes"": """ & EscapeJSON(rsAcessorios("NumeroCortes")) & """, "
            jsonAcessorios = jsonAcessorios & """largura"": """ & EscapeJSON(rsAcessorios("Largura")) & """, "
            jsonAcessorios = jsonAcessorios & """comprimento"": """ & EscapeJSON(rsAcessorios("Comprimento")) & """, "
            jsonAcessorios = jsonAcessorios & """estrutura"": """ & EscapeJSON(rsAcessorios("Estrutura")) & """ "
            jsonAcessorios = jsonAcessorios & "}"

            primeiroItem = False
            rsAcessorios.MoveNext
        Loop
    End If

    jsonAcessorios = jsonAcessorios & "]"

    ' Busca Histórico de atualizações de status
    SQL = "SELECT IdProcesso, DataInsercaoProcesso, DataAlteracaoProcesso, QuemAlterou, Complemento "
    SQL = SQL & "FROM HistoricoAlteracaoProcessoSap "
    SQL = SQL & "WHERE IdSap = " & id

    On Error Resume Next
    SET rsAlteracoes = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsAlteracoes) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    jsonAlteracoes = "["
    primeiroItem = True

    If Not rsAlteracoes.EOF Then
        Do While Not rsAlteracoes.EOF
            If Not primeiroItem Then jsonAlteracoes = jsonAlteracoes & "," End If
            
            jsonAlteracoes = jsonAlteracoes & "{"
            jsonAlteracoes = jsonAlteracoes & """idProcesso"": """ & EscapeJSON(rsAlteracoes("IdProcesso")) & """, "
            jsonAlteracoes = jsonAlteracoes & """dataInsercaoProcesso"": """ & EscapeJSON(rsAlteracoes("DataInsercaoProcesso")) & """, "
            jsonAlteracoes = jsonAlteracoes & """dataAlteracaoProcesso"": """ & EscapeJSON(rsAlteracoes("DataAlteracaoProcesso")) & """, "
            jsonAlteracoes = jsonAlteracoes & """quemAlterou"": """ & EscapeJSON(rsAlteracoes("QuemAlterou")) & """, "
            jsonAlteracoes = jsonAlteracoes & """complemento"": """ & EscapeJSON(rsAlteracoes("Complemento")) & """ "
            jsonAlteracoes = jsonAlteracoes & "}"

            primeiroItem = False
            rsAlteracoes.MoveNext
        Loop
    End If

    jsonAlteracoes = jsonAlteracoes & "]"

    ' Obtem lista de anexos
    SQL = "SELECT Id, IdSolicitacaoProjeto, Nome, Caminho, DataCriacao "
    SQL = SQL & "FROM SolicitacaoProjetoDocumento "
    SQL = SQL & "WHERE IdSolicitacaoProjeto = " & id
    SQL = SQL & "ORDER BY DataCriacao Desc"

    On Error Resume Next
    SET rsAnexos = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsAnexos) Then
        ObterPorId = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de setores: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonAnexos = "["

    If Not rsAnexos.EOF Then
        Do While Not rsAnexos.EOF
            If Not primeiroItem Then jsonAnexos = jsonAnexos & "," End If
            
            jsonAnexos = jsonAnexos & "{"
            jsonAnexos = jsonAnexos & """id"": """ & EscapeJSON(rsAnexos("Id")) & """, "
            jsonAnexos = jsonAnexos & """idSolicitacaoProjeto"": """ & EscapeJSON(rsAnexos("IdSolicitacaoProjeto")) & """, "
            jsonAnexos = jsonAnexos & """nome"": """ & EscapeJSON(rsAnexos("Nome")) & """, "
            jsonAnexos = jsonAnexos & """caminho"": """ & EscapeJSON(rsAnexos("Caminho")) & """, "
            jsonAnexos = jsonAnexos & """dataCriacao"": """ & EscapeJSON(rsAnexos("DataCriacao")) & """ "
            jsonAnexos = jsonAnexos & "}"
            
            primeiroItem = False
            rsAnexos.MoveNext
        Loop
    End If

    jsonAnexos = jsonAnexos & "]"

    rsAnexos.Close()
    SET rsAnexos = NOTHING

    jsonObjeto = "{"

    If Not rsObter.EOF Then
        jsonObjeto = jsonObjeto & """id"": """ & EscapeJSON(rsObter("Id")) & """, "
        jsonObjeto = jsonObjeto & """tipoOnda"": """ & EscapeJSON(rsObter("TipoOnda")) & """, "
        jsonObjeto = jsonObjeto & """data"": """ & EscapeJSON(rsObter("Data")) & """, "
        jsonObjeto = jsonObjeto & """hora"": """ & EscapeJSON(rsObter("Hora")) & """, "
        jsonObjeto = jsonObjeto & """dataEmissao"": """ & EscapeJSON(rsObter("DataEmissao")) & """, "
        jsonObjeto = jsonObjeto & """loginUsuarioCriacao"": """ & EscapeJSON(rsObter("LoginUsuarioCriacao")) & """, "
        jsonObjeto = jsonObjeto & """tipoSolicitacao"": """ & EscapeJSON(rsObter("TipoSolicitacao")) & """, "
        jsonObjeto = jsonObjeto & """status"": """ & EscapeJSON(rsObter("Status")) & """, "
        jsonObjeto = jsonObjeto & """codPlanta"": """ & EscapeJSON(rsObter("CodPlanta")) & """, "
        jsonObjeto = jsonObjeto & """codCliente"": """ & EscapeJSON(rsObter("CodCliente")) & """, "
        jsonObjeto = jsonObjeto & """codLoja"": """ & EscapeJSON(rsObter("CodLoja")) & """, "
        jsonObjeto = jsonObjeto & """nomeCliente"": """ & EscapeJSON(rsObter("NomeCliente")) & """, "
        jsonObjeto = jsonObjeto & """ftOriginal"": """ & EscapeJSON(rsObter("FtOriginal")) & """, "
        jsonObjeto = jsonObjeto & """ftNova"": """ & EscapeJSON(rsObter("FtNova")) & """, "
        jsonObjeto = jsonObjeto & """referencia"": """ & EscapeJSON(rsObter("Referencia")) & """, "
        jsonObjeto = jsonObjeto & """toleranciaMaxima"": """ & EscapeJSON(rsObter("ToleranciaMaxima")) & """, "
        jsonObjeto = jsonObjeto & """toleranciaMinima"": """ & EscapeJSON(rsObter("ToleranciaMinima")) & """, "
        jsonObjeto = jsonObjeto & """estilo"": """ & EscapeJSON(rsObter("Estilo")) & """, "
        jsonObjeto = jsonObjeto & """conta"": """ & EscapeJSON(rsObter("Conta")) & """, "
        jsonObjeto = jsonObjeto & """descNF"": """ & EscapeJSON(rsObter("DescNF")) & """, "
        jsonObjeto = jsonObjeto & """natureza"": """ & EscapeJSON(rsObter("Natureza")) & """, "
        jsonObjeto = jsonObjeto & """observacoes"": """ & EscapeJSON(rsObter("Observacoes")) & """, "
        jsonObjeto = jsonObjeto & """unidadeFabricacao"": """ & EscapeJSON(rsObter("UnidadeFabricacao")) & """, "
        jsonObjeto = jsonObjeto & """distanciaUnidade"": """ & EscapeJSON(rsObter("DistanciaUnidade")) & """, "
        jsonObjeto = jsonObjeto & """comprimento"": """ & EscapeJSON(rsObter("Comprimento")) & """, "
        jsonObjeto = jsonObjeto & """largura"": """ & EscapeJSON(rsObter("Largura")) & """, "
        jsonObjeto = jsonObjeto & """altura"": """ & EscapeJSON(rsObter("Altura")) & """, "
        jsonObjeto = jsonObjeto & """numeroMatriz"": """ & EscapeJSON(rsObter("NumeroMatriz")) & """, "
        jsonObjeto = jsonObjeto & """clienteUsaTara"": """ & EscapeJSON(rsObter("ClienteUsaTara")) & """, "
        jsonObjeto = jsonObjeto & """toleranciaPeso"": """ & EscapeJSON(rsObter("ToleranciaPeso")) & """, "
        jsonObjeto = jsonObjeto & """estruturaPapelao"": """ & EscapeJSON(rsObter("EstruturaPapelao")) & """, "
        jsonObjeto = jsonObjeto & """inkJet"": """ & EscapeJSON(rsObter("InkJet")) & """, "
        jsonObjeto = jsonObjeto & """gramaturaReferencial"": """ & EscapeJSON(rsObter("GramaturaReferencial")) & """, "
        jsonObjeto = jsonObjeto & """esmagamento"": """ & EscapeJSON(rsObter("Esmagamento")) & """, "
        jsonObjeto = jsonObjeto & """ectColuna"": """ & EscapeJSON(rsObter("EctColuna")) & """, "
        jsonObjeto = jsonObjeto & """mullenReferencial"": """ & EscapeJSON(rsObter("MullenReferencial")) & """, "
        jsonObjeto = jsonObjeto & """espessura"": """ & EscapeJSON(rsObter("Espessura")) & """, "
        jsonObjeto = jsonObjeto & """bctCompressao"": """ & EscapeJSON(rsObter("BctCompressao")) & """, "
        jsonObjeto = jsonObjeto & """cobbInterno"": """ & EscapeJSON(rsObter("CobbInterno")) & """, "
        jsonObjeto = jsonObjeto & """cobbExterno"": """ & EscapeJSON(rsObter("CobbExterno")) & """, "
        jsonObjeto = jsonObjeto & """tipoFechamento"": """ & EscapeJSON(rsObter("TipoFechamento")) & """, "
        jsonObjeto = jsonObjeto & """ehPaletizado"": """ & EscapeJSON(rsObter("EhPaletizado")) & """, "
        jsonObjeto = jsonObjeto & """ehAmarrado"": """ & EscapeJSON(rsObter("EhAmarrado")) & """, "
        jsonObjeto = jsonObjeto & """alturaMaximaPalete"": """ & EscapeJSON(rsObter("AlturaMaximaPalete")) & """, "
        jsonObjeto = jsonObjeto & """quantidadeFitilhos"": """ & EscapeJSON(rsObter("QuantidadeFitilhos")) & """, "
        jsonObjeto = jsonObjeto & """tipoFitilho"": """ & EscapeJSON(rsObter("TipoFitilho")) & """, "
        jsonObjeto = jsonObjeto & """pecasPorUnidade"": """ & EscapeJSON(rsObter("PecasPorUnidade")) & """, "
        jsonObjeto = jsonObjeto & """palete"": """ & EscapeJSON(rsObter("Palete")) & """, "
        jsonObjeto = jsonObjeto & """temFacesOpostas"": """ & EscapeJSON(rsObter("TemFacesOpostas")) & """, "
        jsonObjeto = jsonObjeto & """temCantinho"": """ & EscapeJSON(rsObter("TemCantinho")) & """, "
        jsonObjeto = jsonObjeto & """temGradeMadeira"": """ & EscapeJSON(rsObter("TemGradeMadeira")) & """, "
        jsonObjeto = jsonObjeto & """temFilmeStrech"": """ & EscapeJSON(rsObter("TemFilmeStrech")) & """, "
        jsonObjeto = jsonObjeto & """clienteUsaRack"": """ & EscapeJSON(rsObter("ClienteUsaRack")) & """, "
        jsonObjeto = jsonObjeto & """overhang"": """ & EscapeJSON(rsObter("Overhang")) & """, "
        jsonObjeto = jsonObjeto & """quantidadeEspelhos"": """ & EscapeJSON(rsObter("QuantidadeEspelhos")) & """, "
        jsonObjeto = jsonObjeto & """conteudoEmbalado"": """ & EscapeJSON(rsObter("ConteudoEmbalado")) & """, "
        jsonObjeto = jsonObjeto & """consumoMensal"": """ & EscapeJSON(rsObter("ConsumoMensal")) & """, "
        jsonObjeto = jsonObjeto & """tipoLinhaProducaoCliente"": """ & EscapeJSON(rsObter("TipoLinhaProducaoCliente")) & """, "
        jsonObjeto = jsonObjeto & """tipoProdutoPrimario"": """ & EscapeJSON(rsObter("TipoProdutoPrimario")) & """, "
        jsonObjeto = jsonObjeto & """pesoPrimario"": """ & EscapeJSON(rsObter("PesoPrimario")) & """, "
        jsonObjeto = jsonObjeto & """quantidadePrimarios"": """ & EscapeJSON(rsObter("QuantidadePrimarios")) & """, "
        jsonObjeto = jsonObjeto & """tempoEstocagem"": """ & EscapeJSON(rsObter("TempoEstocagem")) & """, "
        jsonObjeto = jsonObjeto & """umidadeRelativa"": """ & EscapeJSON(rsObter("UmidadeRelativa")) & """, "
        jsonObjeto = jsonObjeto & """informacoesAdicionais"": """ & EscapeJSON(rsObter("InformacoesAdicionais")) & """, "
        jsonObjeto = jsonObjeto & """quantidadeAmostras"": """ & EscapeJSON(rsObter("QuantidadeAmostras")) & """, "
        jsonObjeto = jsonObjeto & """amostraCep"": """ & EscapeJSON(rsObter("AmostraCep")) & """, "
        jsonObjeto = jsonObjeto & """amostraResponsavel"": """ & EscapeJSON(rsObter("AmostraResponsavel")) & """, "
        jsonObjeto = jsonObjeto & """amostraEndereco"": """ & EscapeJSON(rsObter("AmostraEndereco")) & """, "
        jsonObjeto = jsonObjeto & """amostraBairro"": """ & EscapeJSON(rsObter("AmostraBairro")) & """, "
        jsonObjeto = jsonObjeto & """amostraMunicipio"": """ & EscapeJSON(rsObter("AmostraMunicipio")) & """, "
        jsonObjeto = jsonObjeto & """amostraUF"": """ & EscapeJSON(rsObter("AmostraUF")) & """, "
        jsonObjeto = jsonObjeto & """amostraTelefone"": """ & EscapeJSON(rsObter("AmostraTelefone")) & """, "
        jsonObjeto = jsonObjeto & """amostraCNPJ"": """ & EscapeJSON(rsObter("AmostraCNPJ")) & """, "
        jsonObjeto = jsonObjeto & """amostraEmail"": """ & EscapeJSON(rsObter("AmostraEmail")) & """, "
        jsonObjeto = jsonObjeto & """statusAndamento"": """ & EscapeJSON(rsObter("StatusAndamento")) & """, "
        jsonObjeto = jsonObjeto & """amostraObservacoes"": """ & EscapeJSON(rsObter("AmostraObservacoes")) & """, "
        jsonObjeto = jsonObjeto & """amostraFormaEnvio"": """ & EscapeJSON(rsObter("AmostraFormaEnvio")) & """, "
        jsonObjeto = jsonObjeto & """motivoReprovacaoAmostra"": """ & EscapeJSON(rsObter("MotivoReprovacaoAmostra")) & """, "
        jsonObjeto = jsonObjeto & """motivoReprovacaoAmostra2"": """ & EscapeJSON(rsObter("MotivoReprovacaoAmostra2")) & """, "
        jsonObjeto = jsonObjeto & """larguraFitilho"": """ & EscapeJSON(rsObter("LarguraFitilho")) & """, "
        jsonObjeto = jsonObjeto & """comprimentoFitilho"": """ & EscapeJSON(rsObter("ComprimentoFitilho")) & """, "
        jsonObjeto = jsonObjeto & """cor"": """ & EscapeJSON(rsObter("Cor")) & """, "
        jsonObjeto = jsonObjeto & """icms"": """ & EscapeJSON(rsObter("ICMS")) & """, "
        jsonObjeto = jsonObjeto & """ipi"": """ & EscapeJSON(rsObter("IPI")) & """, "
        jsonObjeto = jsonObjeto & """tipoPrevisto"": """ & EscapeJSON(rsObter("TipoPrevisto")) & """, "
        jsonObjeto = jsonObjeto & """orcamento"": """ &  EscapeJSON(rsObter("Orcamento")) & """, "
        jsonObjeto = jsonObjeto & """larguraChapa"": """ & EscapeJSON(rsObter("LarguraChapa")) & """, "
        jsonObjeto = jsonObjeto & """comprimentoChapa"": """ & EscapeJSON(rsObter("ComprimentoChapa")) & """, "
        jsonObjeto = jsonObjeto & """testeColuna"": """ & EscapeJSON(rsObter("TesteColuna")) & """, "
        jsonObjeto = jsonObjeto & """testeGramatura"": """ & EscapeJSON(rsObter("TesteGramatura")) & """, "
        jsonObjeto = jsonObjeto & """testeCobInterno"": """ & EscapeJSON(rsObter("TesteCobInterno")) & """, "
        jsonObjeto = jsonObjeto & """testeCobExterno"": """ & EscapeJSON(rsObter("TesteCobExterno")) & """, "
        jsonObjeto = jsonObjeto & """testeDescolamentoCapaInterna"": """ & EscapeJSON(rsObter("TesteDescolamentoCapaInterna")) & """, "
        jsonObjeto = jsonObjeto & """testeDescolamentoMiolo1"": """ & EscapeJSON(rsObter("TesteDescolamentoMiolo1")) & """, "
        jsonObjeto = jsonObjeto & """testeDescolamentoMiolo2"": """ & EscapeJSON(rsObter("TesteDescolamentoMiolo2")) & """, "
        jsonObjeto = jsonObjeto & """testeDescolamentoMiolo3"": """ & EscapeJSON(rsObter("TesteDescolamentoMiolo3")) & """, "
        jsonObjeto = jsonObjeto & """testeDescolamentoCapaExterna"": """ & EscapeJSON(rsObter("TesteDescolamentoCapaExterna")) & """, "
        jsonObjeto = jsonObjeto & """testeBctCompressao1"": """ & EscapeJSON(rsObter("TesteBctCompressao1")) & """, "
        jsonObjeto = jsonObjeto & """testeBctCompressao2"": """ & EscapeJSON(rsObter("TesteBctCompressao2")) & """, "
        jsonObjeto = jsonObjeto & """testeBctCompressao3"": """ & EscapeJSON(rsObter("TesteBctCompressao3")) & """, "
        jsonObjeto = jsonObjeto & """testePinInterno"": """ & trim(rsObter("TestePinInterno")) & """, "
        jsonObjeto = jsonObjeto & """testePinExterno"": """ & trim(rsObter("TestePinExterno")) & """, "
        jsonObjeto = jsonObjeto & """testeForcaVinco1"": """ & trim(rsObter("TesteForcaVinco1")) & """, "
        jsonObjeto = jsonObjeto & """testeForcaVinco2"": """ & trim(rsObter("TesteForcaVinco2")) & """, "
        jsonObjeto = jsonObjeto & """testeForcaVinco3"": """ & trim(rsObter("TesteForcaVinco3")) & """, "
        jsonObjeto = jsonObjeto & """processo"": """ & EscapeJSON(rsObter("Processo")) & """, "
        jsonObjeto = jsonObjeto & """idDesenvolvedor"": """ & EscapeJSON(rsObter("IdDesenvolvedor")) & """, "
        jsonObjeto = jsonObjeto & """margemContribuicao"": """ & EscapeJSON(rsObter("MargemContribuicao")) & """, "
        jsonObjeto = jsonObjeto & """resultado"": """ & EscapeJSON(rsObter("Resultado")) & """, "
        jsonObjeto = jsonObjeto & """quantidadeCores"": """ & EscapeJSON(rsObter("QuantidadeCores")) & """, "
        jsonObjeto = jsonObjeto & """justificativaRejeicao"": """ & EscapeJSON(rsObter("JustificativaRejeicao")) & """, "
        jsonObjeto = jsonObjeto & """idMunicipioCliente"": """ & EscapeJSON(rsObter("IdMunicipioCliente")) & """, "
        jsonObjeto = jsonObjeto & """gramatura"": """ & trim(rsObter("Gramatura")) & """, "
        jsonObjeto = jsonObjeto & """tipoSavings"": """ & trim(rsObter("TipoSavings")) & """, "
        jsonObjeto = jsonObjeto & """valorSavingSK"": """ & trim(rsObter("ValorSavingSK")) & """, "
        jsonObjeto = jsonObjeto & """valorSavingCliente"": """ & trim(rsObter("ValorSavingCliente")) & """, "
        jsonObjeto = jsonObjeto & """nomeVendedor"": """ & trim(rsObter("NomeVendedor")) & """, "
        jsonObjeto = jsonObjeto & """idVendedor"": """ & trim(rsObter("IdVendedor")) & """, "
        jsonObjeto = jsonObjeto & """statusAndamento"": """ & trim(rsObter("StatusAndamento")) & """, "
        jsonObjeto = jsonObjeto & """dataConfeccaoFerramental"": """ & trim(rsObter("DataConfeccaoFerramental")) & """, "
        jsonObjeto = jsonObjeto & """necessitaAcompanhamento"": """ & trim(rsObter("NecessitaAcompanhamento")) & """, "
        jsonObjeto = jsonObjeto & """areaLiquida"": """ & trim(rsObter("AreaLiquida")) & """, "
        jsonObjeto = jsonObjeto & """suportacaoPrimario"": """ & trim(rsObter("SuportacaoPrimario")) & """, "
        jsonObjeto = jsonObjeto & """precoKG"": """ & trim(rsObter("PrecoKG")) & """, "
        jsonObjeto = jsonObjeto & """valorEstimadoCliche"": """ & trim(rsObter("ValorEstimadoCliche")) & """, "
        jsonObjeto = jsonObjeto & """valorEstimadoMatriz"": """ & trim(rsObter("ValorEstimadoMatriz")) & """, "
        jsonObjeto = jsonObjeto & """nivelLotePiloto"": """ & trim(rsObter("NivelLotePiloto")) & """, "
        jsonObjeto = jsonObjeto & """dataEnvioAmostra"": """ & trim(rsObter("DataEnvioAmostra")) & """, "
        jsonObjeto = jsonObjeto & """quemEnviouAmostra"": """ & trim(rsObter("QuemEnviouAmostra")) & """, "
        jsonObjeto = jsonObjeto & """ehPrecadastro"": """ & trim(rsObter("EhPrecadastro")) & """"

        If carregarRelacionamentos = 1 Then
            jsonObjeto = jsonObjeto & ",""acessorios"": " & jsonAcessorios & ""
            jsonObjeto = jsonObjeto & ",""alteracoes"": " & jsonAlteracoes & ""
            jsonObjeto = jsonObjeto & ",""anexos"": " & jsonAnexos & ""
        End If
    End If
    
    jsonObjeto = jsonObjeto & "}"

    rsObter.Close()
    SET rsObter = NOTHING
    
    ObterPorId = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonObjeto & "}"
End Function

Function Criar()    
    nomeCliente = Left(trim(Request("nomeCliente")), 30)
    codPlanta = Request("codPlanta")
    idDesenvolvedor = Request("idDesenvolvedor")
    codCliente = Request("codCliente")
    codLoja = Right("0" & Request("codLoja"), 2) 
    referencia = Request("referencia")
    necessitaAcompanhamento = Request("necessitaAcompanhamento")
    nivelLotePiloto = Iif (Request("nivelLotePiloto") = "", 0, Request("nivelLotePiloto")) 
    toleranciaMaxima = Request("toleranciaMaxima")
    toleranciaMinima = Request("toleranciaMinima")
    estilo = Request("estilo")
    observacoes = Request("observacoes")
    comprimento = Request("comprimento")
    largura = Request("largura")
    altura = Request("altura")
    numeroMatriz = Request("numeroMatriz")
    clienteUsaTara = Request("clienteUsaTara")
    toleranciaPeso = Iif (Request("toleranciaPeso") = "", 0, Request("toleranciaPeso")) 
    estruturaPapelao = Request("estruturaPapelao")
    gramatura = Request("gramatura")
    larguraChapa = Request("larguraChapa")
    comprimentoChapa = Request("comprimentoChapa")
    areaBlank = Request("areaBlank")
    areaLiquida = Request("areaLiquida")
    peso = Request("peso")
    inkJet = Request("inkJet")
    gramaturaReferencial = Request("gramaturaReferencial")
    esmagamento = Request("esmagamento")
    ectColuna = Request("ectColuna")
    mullenReferencial = Request("mullenReferencial")
    espessura = Request("espessura")
    bctCompressao = Request("bctCompressao") 
    cobbInterno = Request("cobbInterno")
    cobbExterno = Request("cobbExterno")
    tipoFechamento = Request("tipoFechamento")
    ehAmarrado = Request("ehAmarrado")
    quantidadeFitilhos = Iif (Request("quantidadeFitilhos") = "", 0, Request("quantidadeFitilhos"))
    tipoFitilho = Request("tipoFitilho")
    pecasPorUnidade = Request("pecasPorUnidade")
    ehPaletizado = Request("ehPaletizado")
    alturaMaximaPalete = Iif (Request("alturaMaximaPalete") = "", 0, Request("alturaMaximaPalete"))
    palete = Request("palete")
    temFacesOpostas = Request("temFacesOpostas")
    temCantinho = Request("temCantinho")
    temGradeMadeira = Request("temGradeMadeira")
    temfilmeStrech = Request("temFilmeStrech")
    clienteUsaRack = Request("clienteUsaRack")
    overhang = Request("overhang")
    quantidadeEspelhos = Iif (Request("quantidadeEspelhos") = "", 0, Request("quantidadeEspelhos"))
    acessoriosJSON = Request("acessorios")
    conteudoEmbalado = Request("conteudoEmbalado")
    consumoMensal = Request("consumoMensal")
    tipoLinhaProducaoCliente = Request("tipoLinhaProducaoCliente")
    tipoProdutoPrimario = Request("tipoProdutoPrimario")
    pesoPrimario = Request("pesoPrimario")
    quantidadePrimarios = Request("quantidadePrimarios")
    suportacaoPrimario = Request("suportacaoPrimario")
    tempoEstocagem = Request("tempoEstocagem")
    umidadeRelativa = Request("umidadeRelativa")
    quantidadeCores = Request("quantidadeCores")
    precoKG = Request("precoKG")
    orcamento = Iif (Request("orcamento") = "", 0, Request("orcamento"))
    margemContribuicao = Request("margemContribuicao")
    resultado = Request("resultado")
    tipoSavings = Request("tipoSavings")
    valorSavingSK = Request("valorSavingSK")
    valorSavingCliente = Request("valorSavingCliente")
    valorEstimadoCliche = Request("valorEstimadoCliche")
    valorEstimadoMatriz = Request("valorEstimadoMatriz")
    valorEstimadoPapelao = Request("valorEstimadoPapelao")
    informacoesAdicionais = Request("informacoesAdicionais")
    quantidadeAmostras = Request("quantidadeAmostras")
    amostraCep = Request("amostraCep")
    amostraResponsavel = Request("amostraResponsavel")
    amostraEndereco = Request("amostraEndereco")
    amostraBairro = Request("amostraBairro")
    amostraMunicipio = Request("amostraMunicipio")
    amostraUF = Request("amostraUF")
    amostraTelefone = Request("amostraTelefone")
    amostraCNPJ = Request("amostraCNPJ") 
    amostraEmail = Request("amostraEmail")
    amostraObservacoes = Request("amostraObservacoes")
    numeroFTCopia = Request("numeroFTCopia")
    larguraFitilho = Iif(tipoFitilho = "L", quantidadeFitilhos, 0)
    comprimentoFitilho = Iif(tipoFitilho = "C", quantidadeFitilhos, 0)
    ehPreCadastro = Request("ehPreCadastro")
    tipoSolicitacao = Request("tipoSolicitacao")
    processo = 0
    codClienteSAP = Iif(codPlanta = "07", codCliente & codPlanta, codCliente)
    anexosJSON = Request("anexos")

    flagAprovacaoComercial = ObterFlagAprovacaoComercial(codPlanta)

    idNovaSolicitacao = GerarProximoIdSolicitacao(codPlanta)
    If idNovaSolicitacao = "" Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao obter próximo id""}"
        Exit Function
    End If

    Set dadosContabilidade = ObterDadosContabilidade(codPlanta, estilo)
    conta = dadosContabilidade("conta")
    descNF = dadosContabilidade("descNF")
    natureza = dadosContabilidade("natureza")

    Set dadosVendedor = ObterDadosVendedor(codPlanta, codCliente, codLoja)
    idMunicipioCliente = 0 ' Verificar necessidade
    idVendedor = dadosVendedor("idVendedor")
    nomeVendedor = dadosVendedor("nomeVendedor")

    ' Verificar necessidade destes campos
    unidadeFabricacao = ObterUnidadeFabricacao(codPlanta)

    ' Campos de outros formulários (validar)
    justificativaRejeicao = ""
    amostraFormaEnvio = ""
    cor = "" 

    icms = 0
    ipi = 0 
    orcamentista = 0

    tipoPrevisto = 0
    dataConfeccaoFerramental = "NULL" 

    SQL = "INSERT INTO SKFT1 (FT1_ID, FT1_DATA, FT1_HORA, FT1_USU, FT1_TIPO, FT1_STATUS, FT1_B5CLI, "
    SQL = SQL & "FT1_B5NOMECLI, FT1_LOJA, FT1_FT, FT1_B5REF, FT1_B5TOLMA, FT1_B5TOLME, FT1_B5ESTILO, "
    SQL = SQL & "FT1_B1CONTA, FT1_B1DESCNF, FT1_B1NATUREZ, FT1_OBSP, FT1_B1INPUBE, FT1_KMUNID, "
    SQL = SQL & "FT1_B5COMPR, FT1_B5LARG, FT1_B5ALTU, FT1_B5NUMFACA, FT1_TARA, FT1_TOLPESO, "
    SQL = SQL & "FT1_B5ESTRUTU, FT1_INKJET, FT1_B5GRAM, FT1_B5ESMAG, FT1_B5COLUNA, FT1_MULLEN, "
    SQL = SQL & "FT1_B5ESPES2, FT1_B5COMPRE, FT1_B5COBINT, FT1_B5COBEXT, FT1_B5FECHAM, FT1_B5PALETE, "
    SQL = SQL & "FT1_B5AMARRAD, FT1_ALTPAL, FT1_FITILH, FT1_FITPOS, FT1_B5QTDPCT, FT1_B5TPPALE, "
    SQL = SQL & "FT1_B5FACOP, FT1_B5CANTON, FT1_B5TAMPA, FT1_B5FILME, FT1_RACK, FT1_OVERHANG, "
    SQL = SQL & "FT1_B5ESPELHO, FT1_B1ALIMENT, FT1_CONSMES, FT1_LINAUTO, FT1_TIPOPRD, FT1_PESOPRI, "
    SQL = SQL & "FT1_QTDPRI, FT1_TEMPOEST, FT1_UNIREL, FT1_OBSERV, FT1_QTDAMO, FT1_CEP, FT1_CON, "
    SQL = SQL & "FT1_END, FT1_BAI, FT1_MUN, FT1_EST, FT1_TEL, FT1_CNPJ, FT1_EMAIL, FT1_DEL, "
    SQL = SQL & "FT1_B5FITLAR, FT1_B5FITCOM, FT1_FTNOVA, FT1_COR, FT1_ICMS, FT1_IPI, FT1_ORCAMENTISTA, "
    SQL = SQL & "FT1_TIPOPREVISTO, FT1_ORCAMENTO, FT1_STATUSAMOSTRA, FT1_MOTIVOREPROVACAOAMOSTRA, "
    SQL = SQL & "FT1_LARGURACHAPA, FT1_COMPRIMENTOCHAPA, FT1_TESTECOLUNA, FT1_TESTEGRAMATURA, "
    SQL = SQL & "FT1_TESTECOBINTERNO, FT1_TESTECOBEXTERNO, FT1_TESTEMULLENINTERNOCAPA, "
    SQL = SQL & "FT1_TESTEMULLENINTERNOMIOLO, FT1_TESTEMULLENINTERNOCAPAINTERMEDIARIA, "
    SQL = SQL & "FT1_TESTEMULLENEXTERNOCAPA, FT1_TESTEMULLENEXTERNOMIOLO, FT1_COMPRESSAO1, "
    SQL = SQL & "FT1_COMPRESSAO2, FT1_COMPRESSAO3, FT1_TipoAlteraCancela, FT1_FILIAL, FT1_TECNICO, "
    SQL = SQL & "FT1_Margem, FT1_Resultado, FT1_QTDCOR, FT1_MOTCANC, FT1_ObservacaoStatus2, "
    SQL = SQL & "FT1_FormaEnvioAmostra, FT1_IDMunicipio, FT1_Z1GRAMC, FT1_MOTIVOREPROVACAOAMOSTRA2, "
    SQL = SQL & "FT1_SAVING, FT1_SAVINGVALK, FT1_SAVINGVALC, FT1_PLANTA, FT1_A3NREDUZ, FT1_A1VEND, "
    SQL = SQL & "FT1_STATUS2, FT1_DTCONFEC, FT1_AcompanharLotePiloto, FT1_B5AREA, FT1_SUPORTEPRIMARIO, "
    SQL = SQL & "FT1_PRECOKG, FT1_VALORCLICHE, FT1_VALORMATRIZ, FT1_NIVELLOTEPILOTO, FT1_PININTERNO, "
    SQL = SQL & "FT1_PINEXTERNO, FT1_FORCAVINCO1, FT1_FORCAVINCO2, FT1_FORCAVINCO3, FT1_PRECADASTRO "
    SQL = SQL & ") VALUES (" 
    SQL = SQL & "'" & idNovaSolicitacao & "', "
    SQL = SQL & "FORMAT(GETDATE(), 'yyyyMMdd'),"
    SQL = SQL & "FORMAT(GETDATE(), 'HH:mm:ss'),"
    SQL = SQL & "'" & usuarioLogado("login") & "', "
    SQL = SQL & "'" & tipoSolicitacao & "', "
    SQL = SQL & Iif(flagAprovacaoComercial, "'S', ", "'A', ")
    SQL = SQL & "'" & codClienteSAP & "', "
    SQL = SQL & "'" & nomeCliente & "', "
    SQL = SQL & "'" & codLoja & "', "
    SQL = SQL & "'" & numeroFTCopia & "', "
    SQL = SQL & "'" & referencia & "', "
    SQL = SQL & "" & toleranciaMaxima & ", "
    SQL = SQL & "" & toleranciaMinima & ", "
    SQL = SQL & "'" & estilo & "', "
    SQL = SQL & "'" & conta & "', "
    SQL = SQL & "'" & descNF & "', "
    SQL = SQL & "'" & natureza & "', "
    SQL = SQL & "'" & observacoes & "', "
    SQL = SQL & "'" & unidadeFabricacao & "', "
    SQL = SQL & "'" & 0 & "', "
    SQL = SQL & "" & comprimento & ", "
    SQL = SQL & "" & largura & ", "
    SQL = SQL & "" & altura & ", "
    SQL = SQL & "'" & numeroMatriz & "', "
    SQL = SQL & "" & clienteUsaTara & ", "
    SQL = SQL & "" & toleranciaPeso & ", "
    SQL = SQL & "'" & estruturaPapelao & "', "
    SQL = SQL & "'" & inkJet & "', "
    SQL = SQL & "" & gramaturaReferencial & ", "
    SQL = SQL & "" & esmagamento & ", "
    SQL = SQL & "" & ectColuna & ", "
    SQL = SQL & "" & mullenReferencial & ", "
    SQL = SQL & "" & espessura & ", "
    SQL = SQL & "" & bctCompressao & ", "
    SQL = SQL & "" & cobbInterno & ", "
    SQL = SQL & "" & cobbExterno & ", "
    SQL = SQL & "'" & tipoFechamento & "', "
    SQL = SQL & "'" & ehPaletizado & "', "
    SQL = SQL & "'" & ehAmarrado & "', "
    SQL = SQL & "" & alturaMaximaPalete & ", "
    SQL = SQL & "" & quantidadeFitilhos & ", "
    SQL = SQL & "'" & tipoFitilho & "', "
    SQL = SQL & "" & pecasPorUnidade & ", "
    SQL = SQL & "'" & palete & "', "
    SQL = SQL & "'" & temFacesOpostas & "', "
    SQL = SQL & "'" & temCantinho & "', "
    SQL = SQL & "'" & temGradeMadeira & "', "
    SQL = SQL & "'" & temfilmeStrech & "', "
    SQL = SQL & "'" & clienteUsaRack & "', "
    SQL = SQL & "'" & overhang & "', "
    SQL = SQL & "'" & quantidadeEspelhos & "', "
    SQL = SQL & "'" & conteudoEmbalado & "', "
    SQL = SQL & "" & consumoMensal & ", "
    SQL = SQL & "'" & tipoLinhaProducaoCliente & "', "
    SQL = SQL & "'" & tipoProdutoPrimario & "', "
    SQL = SQL & "" & pesoPrimario & ", "
    SQL = SQL & "" & quantidadePrimarios & ", "
    SQL = SQL & "" & tempoEstocagem & ", "
    SQL = SQL & "" & umidadeRelativa & ", "
    SQL = SQL & "'" & informacoesAdicionais & "', " 
    SQL = SQL & "" & quantidadeAmostras & ", "
    SQL = SQL & "'" & amostraCep & "', "
    SQL = SQL & "'" & amostraResponsavel & "', "
    SQL = SQL & "'" & amostraEndereco & "', "
    SQL = SQL & "'" & amostraBairro & "', "
    SQL = SQL & "'" & amostraMunicipio & "', "
    SQL = SQL & "'" & amostraUF & "', "
    SQL = SQL & "'" & amostraTelefone & "', "
    SQL = SQL & "'" & amostraCNPJ & "', "
    SQL = SQL & "'" & amostraEmail & "', "
    SQL = SQL & "0, "
    SQL = SQL & "" & larguraFitilho & ", " 
    SQL = SQL & "" & comprimentoFitilho & ", " 
    SQL = SQL & "'', " 
    SQL = SQL & "'" & cor & "', " 
    SQL = SQL & "" & icms & ", " 
    SQL = SQL & "" & ipi & ", " 
    SQL = SQL & "" & orcamentista & ", " 
    SQL = SQL & "'" & tipoPrevisto & "', " 
    SQL = SQL & "" & orcamento & ", " 
    SQL = SQL & "'" & "" & "', " 
    SQL = SQL & "'" & "" & "', " 
    SQL = SQL & "" & larguraChapa & ", " 
    SQL = SQL & "" & comprimentoChapa & ", " 
    SQL = SQL & "" & 0 & ", " 
    SQL = SQL & "" & 0 & ", " 
    SQL = SQL & "" & 0 & ", " 
    SQL = SQL & "" & 0 & ", " 
    SQL = SQL & "" & 0 & ", " 
    SQL = SQL & "" & 0 & ", " 
    SQL = SQL & "" & 0 & ", " 
    SQL = SQL & "" & 0 & ", " 
    SQL = SQL & "" & 0 & ", " 
    SQL = SQL & "" & 0 & ", " 
    SQL = SQL & "" & 0 & ", " 
    SQL = SQL & "" & 0 & ", " 
    SQL = SQL & "'" & processo  & "', "
    SQL = SQL & "'" & codPlanta & "', " 
    SQL = SQL & "'" & idDesenvolvedor & "', " 
    SQL = SQL & "" & margemContribuicao & ", " 
    SQL = SQL & "" & resultado & ", " 
    SQL = SQL & "" & quantidadeCores & ", " 
    SQL = SQL & "'" & justificativaRejeicao & "', " 
    SQL = SQL & "'" & amostraObservacoes & "', " 
    SQL = SQL & "'" & amostraFormaEnvio & "', " 
    SQL = SQL & "" & idMunicipioCliente & ", " 
    SQL = SQL & "" & gramatura & ", " 
    SQL = SQL & "'', "
    SQL = SQL & "'" & tipoSavings & "', " 
    SQL = SQL & "" & valorSavingSK & ", " 
    SQL = SQL & "" & valorSavingCliente & ", " 
    SQL = SQL & "'" & codPlanta & "', " 
    SQL = SQL & "'" & nomeVendedor & "', " 
    SQL = SQL & "'" & idVendedor & "', " 
    SQL = SQL & "'" & "" & "', " 
    SQL = SQL & "" & dataConfeccaoFerramental & ", " 
    SQL = SQL & "'" & necessitaAcompanhamento & "', " 
    SQL = SQL & "" & areaLiquida & ", " 
    SQL = SQL & "'" & suportacaoPrimario & "', " 
    SQL = SQL & "" & precoKG & ", " 
    SQL = SQL & "" & valorEstimadoCliche & ", " 
    SQL = SQL & "" & valorEstimadoMatriz & ", " 
    SQL = SQL & "'" & nivelLotePiloto & "', " 
    SQL = SQL & "'" & 0 & "', " 
    SQL = SQL & "'" & 0 & "', " 
    SQL = SQL & "'" & 0 & "', " 
    SQL = SQL & "'" & 0 & "', " 
    SQL = SQL & "'" & 0 & "', " 
    SQL = SQL & "'" & ehPreCadastro & "')" 

    On Error Resume Next
    SET rsCriar = Conecta(ObterConexaoCorporativo(codPlanta), SQL)

    If Err.Number <> 0 OR Not IsObject(rsCriar) Then
        Criar = "{""sucesso"": false, ""mensagem"": ""Erro ao criar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsCriar.Close()
    SET rsCriar = NOTHING
    On Error Goto 0

    If acessoriosJSON <> "" Then
        Call InserirAcessorios(idNovaSolicitacao, acessoriosJSON, codPlanta)
    End If

    If anexosJSON <> "" Then
        Call InserirAnexos(idNovaSolicitacao, anexosJSON, codPlanta)
    End If

    ' Se tipo de solicitação for altera/cancela deve cancelar a FT passada
    If tipoSolicitacao = "3" Then
        If codPlanta = "00" OR codPlanta = "04" Then
            'Pirapetinga e Uberaba usam protheus
            cancelouFT = CancelarFTProtheus(numeroFTCopia, codPlanta)
        Else
            'Outras plantas usam trimbox
            cancelouFT = CancelarFTTrimbox(numeroFTCopia, codPlanta)
        End If

        If Not cancelouFT Then
            AprovacaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Erro ao cancelar FT""}"
        End If
    End If

    Call EnviarEmailCriacao(idNovaSolicitacao, codPlanta, codClienteSAP, nomeCliente, codLoja)
    Call RegistrarLog("criacao", idNovaSolicitacao, "", "")
    Criar = "{""sucesso"": true, ""mensagem"": ""Solicitação criada com sucesso!""}"
End Function

Function Editar()
    id = Request("id")
    codPlanta = Request("codPlanta")

    If id = "" Then
        Editar = "{""sucesso"": false, ""mensagem"": ""ID da solicitação não informado""}"
        Exit Function
    End If

    If codPlanta = "" Then
        Editar = "{""sucesso"": false, ""mensagem"": ""Planta da solicitação nao informada""}"
        Exit Function
    End If
    
    anexosJSON = Request("anexos")
    flagEnviarPCP = Request("flagEnviarPCP")
    idDesenvolvedor = Request("idDesenvolvedor")
    referencia = Request("referencia")
    necessitaAcompanhamento = Request("necessitaAcompanhamento")
    nivelLotePiloto = Request("nivelLotePiloto")
    toleranciaMaxima = Request("toleranciaMaxima")
    toleranciaMinima = Request("toleranciaMinima")
    estilo = Request("estilo")
    observacoes = Request("observacoes")
    comprimento = Request("comprimento")
    largura = Request("largura")
    altura = Request("altura")
    numeroMatriz = Request("numeroMatriz")
    clienteUsaTara = Request("clienteUsaTara")
    toleranciaPeso = Iif(Request("toleranciaPeso") = "", 0, Request("toleranciaPeso"))
    estruturaPapelao = Request("estruturaPapelao")
    gramatura = Request("gramatura")
    larguraChapa = Request("larguraChapa")
    comprimentoChapa = Request("comprimentoChapa")
    areaBlank = Request("areaBlank")
    areaLiquida = Request("areaLiquida")
    peso = Request("peso")
    inkJet = Request("inkJet")
    gramaturaReferencial = Request("gramaturaReferencial")
    esmagamento = Request("esmagamento")
    ectColuna = Request("ectColuna")
    mullenReferencial = Request("mullenReferencial")
    espessura = Request("espessura")
    bctCompressao = Request("bctCompressao")
    cobbInterno = Request("cobbInterno")
    cobbExterno = Request("cobbExterno")
    tipoFechamento = Request("tipoFechamento")
    ehAmarrado = Request("ehAmarrado")
    quantidadeFitilhos = Iif(Request("quantidadeFitilhos") = "", 0, Request("quantidadeFitilhos"))
    tipoFitilho = Request("tipoFitilho")
    pecasPorUnidade = Request("pecasPorUnidade")
    ehPaletizado = Request("ehPaletizado")
    alturaMaximaPalete = Iif (Request("alturaMaximaPalete") = "", 0, Request("alturaMaximaPalete"))
    palete = Request("palete")
    temFacesOpostas = Request("temFacesOpostas")
    temCantinho = Request("temCantinho")
    temGradeMadeira = Request("temGradeMadeira")
    temfilmeStrech = Request("temFilmeStrech")
    clienteUsaRack = Request("clienteUsaRack")
    overhang = Request("overhang")
    quantidadeEspelhos = Request("quantidadeEspelhos")
    acessoriosJSON = Request("acessorios")
    conteudoEmbalado = Request("conteudoEmbalado")
    consumoMensal = Request("consumoMensal")
    tipoLinhaProducaoCliente = Request("tipoLinhaProducaoCliente")
    tipoProdutoPrimario = Request("tipoProdutoPrimario")
    pesoPrimario = Request("pesoPrimario")
    quantidadePrimarios = Request("quantidadePrimarios")
    suportacaoPrimario = Request("suportacaoPrimario")
    tempoEstocagem = Request("tempoEstocagem")
    umidadeRelativa = Request("umidadeRelativa")
    quantidadeCores = Request("quantidadeCores")
    precoKG = Request("precoKG")
    orcamento = Iif (Request("orcamento") = "", 0, Request("orcamento"))
    margemContribuicao = Request("margemContribuicao")
    resultado = Request("resultado")
    tipoSavings = Request("tipoSavings")
    valorSavingSK = Request("valorSavingSK")
    valorSavingCliente = Request("valorSavingCliente")
    valorEstimadoCliche = Request("valorEstimadoCliche")
    valorEstimadoMatriz = Request("valorEstimadoMatriz")
    valorEstimadoPapelao = Request("valorEstimadoPapelao")
    informacoesAdicionais = Request("informacoesAdicionais")
    quantidadeAmostras = Request("quantidadeAmostras")
    amostraCep = Request("amostraCep")
    amostraResponsavel = Request("amostraResponsavel")
    amostraEndereco = Request("amostraEndereco")
    amostraBairro = Request("amostraBairro")
    amostraMunicipio = Request("amostraMunicipio")
    amostraUF = Request("amostraUF")
    amostraTelefone = Request("amostraTelefone")
    amostraCNPJ = Request("amostraCNPJ")
    amostraEmail = Request("amostraEmail")
    amostraObservacoes = Request("amostraObservacoes")
    larguraFitilho = Iif(tipoFitilho = "L", quantidadeFitilhos, 0)
    comprimentoFitilho = Iif(tipoFitilho = "C", quantidadeFitilhos, 0)

    Set dadosContabilidade = ObterDadosContabilidade(codPlanta, estilo)
    conta = dadosContabilidade("conta")
    descNF = dadosContabilidade("descNF")
    natureza = dadosContabilidade("natureza")
    
    unidadeFabricacao = ObterUnidadeFabricacao(codPlanta)

    SQL = "UPDATE SKFT1 SET "
    SQL = SQL & Iif(flagEnviarPCP = 1, "FT1_STATUS = 'D', ",  "")
    SQL = SQL & "FT1_B5REF = '" & referencia & "', "
    SQL = SQL & "FT1_B5TOLMA = " & toleranciaMaxima & ", "
    SQL = SQL & "FT1_B5TOLME = " & toleranciaMinima & ", "
    SQL = SQL & "FT1_B5ESTILO = '" & estilo & "', "
    SQL = SQL & "FT1_B1CONTA = '" & conta & "', "
    SQL = SQL & "FT1_B1DESCNF = '" & descNF & "', "
    SQL = SQL & "FT1_B1NATUREZ = '" & natureza & "', "
    SQL = SQL & "FT1_OBSP = '" & observacoes & "', "
    SQL = SQL & "FT1_B1INPUBE = '" & unidadeFabricacao & "', "
    SQL = SQL & "FT1_KMUNID = '" & 0 & "', "
    SQL = SQL & "FT1_B5COMPR = " & comprimento & ", "
    SQL = SQL & "FT1_B5LARG = " & largura & ", "
    SQL = SQL & "FT1_B5ALTU = " & altura & ", "
    SQL = SQL & "FT1_B5NUMFACA = '" & numeroMatriz & "', "
    SQL = SQL & "FT1_TARA = " & clienteUsaTara & ", "
    SQL = SQL & "FT1_TOLPESO = " & toleranciaPeso & ", "
    SQL = SQL & "FT1_B5ESTRUTU = '" & estruturaPapelao & "', "
    SQL = SQL & "FT1_INKJET = '" & inkJet & "', "
    SQL = SQL & "FT1_B5GRAM = " & gramaturaReferencial & ", "
    SQL = SQL & "FT1_B5ESMAG = " & esmagamento & ", "
    SQL = SQL & "FT1_B5COLUNA = " & ectColuna & ", "
    SQL = SQL & "FT1_MULLEN = " & mullenReferencial & ", "
    SQL = SQL & "FT1_B5ESPES2 = " & espessura & ", "
    SQL = SQL & "FT1_B5COMPRE = " & bctCompressao & ", "
    SQL = SQL & "FT1_B5COBINT = " & cobbInterno & ", "
    SQL = SQL & "FT1_B5COBEXT = " & cobbExterno & ", "
    SQL = SQL & "FT1_B5FECHAM = '" & tipoFechamento & "', "
    SQL = SQL & "FT1_B5PALETE = '" & ehPaletizado & "', "
    SQL = SQL & "FT1_B5AMARRAD = '" & ehAmarrado & "', "
    SQL = SQL & "FT1_ALTPAL = " & alturaMaximaPalete & ", "
    SQL = SQL & "FT1_FITILH = " & quantidadeFitilhos & ", "
    SQL = SQL & "FT1_FITPOS = '" & tipoFitilho & "', "
    SQL = SQL & "FT1_B5QTDPCT = " & pecasPorUnidade & ", "
    SQL = SQL & "FT1_B5TPPALE = '" & palete & "', "
    SQL = SQL & "FT1_B5FACOP = '" & temFacesOpostas & "', "
    SQL = SQL & "FT1_B5CANTON = '" & temCantinho & "', "
    SQL = SQL & "FT1_B5TAMPA = '" & temGradeMadeira & "', "
    SQL = SQL & "FT1_B5FILME = '" & temfilmeStrech & "', "
    SQL = SQL & "FT1_RACK = '" & clienteUsaRack & "', "
    SQL = SQL & "FT1_OVERHANG = '" & overhang & "', "
    SQL = SQL & "FT1_B5ESPELHO = '" & quantidadeEspelhos & "', "
    SQL = SQL & "FT1_B1ALIMENT = '" & conteudoEmbalado & "', "
    SQL = SQL & "FT1_CONSMES = " & consumoMensal & ", "
    SQL = SQL & "FT1_LINAUTO = '" & tipoLinhaProducaoCliente & "', "
    SQL = SQL & "FT1_TIPOPRD = '" & tipoProdutoPrimario & "', "
    SQL = SQL & "FT1_PESOPRI = " & pesoPrimario & ", "
    SQL = SQL & "FT1_QTDPRI = " & quantidadePrimarios & ", "
    SQL = SQL & "FT1_TEMPOEST = " & tempoEstocagem & ", "
    SQL = SQL & "FT1_UNIREL = " & umidadeRelativa & ", "
    SQL = SQL & "FT1_OBSERV = '" & informacoesAdicionais & "', "
    SQL = SQL & "FT1_QTDAMO = " & quantidadeAmostras & ", "
    SQL = SQL & "FT1_CEP = '" & amostraCep & "', "
    SQL = SQL & "FT1_CON = '" & amostraResponsavel & "', "
    SQL = SQL & "FT1_END = '" & amostraEndereco & "', "
    SQL = SQL & "FT1_BAI = '" & amostraBairro & "', "
    SQL = SQL & "FT1_MUN = '" & amostraMunicipio & "', "
    SQL = SQL & "FT1_EST = '" & amostraUF & "', "
    SQL = SQL & "FT1_TEL = '" & amostraTelefone & "', "
    SQL = SQL & "FT1_CNPJ = '" & amostraCNPJ & "', "
    SQL = SQL & "FT1_EMAIL = '" & amostraEmail & "', "
    SQL = SQL & "FT1_B5FITLAR = " & larguraFitilho & ", "
    SQL = SQL & "FT1_B5FITCOM = " & comprimentoFitilho & ", "
    SQL = SQL & "FT1_ORCAMENTO = " & orcamento & ", "
    SQL = SQL & "FT1_LARGURACHAPA = " & larguraChapa & ", "
    SQL = SQL & "FT1_COMPRIMENTOCHAPA = " & comprimentoChapa & ", "
    SQL = SQL & "FT1_TECNICO = '" & idDesenvolvedor & "', "
    SQL = SQL & "FT1_Margem = " & margemContribuicao & ", "
    SQL = SQL & "FT1_Resultado = " & resultado & ", "
    SQL = SQL & "FT1_QTDCOR = " & quantidadeCores & ", "
    SQL = SQL & "FT1_ObservacaoStatus2 = '" & amostraObservacoes & "', "
    SQL = SQL & "FT1_Z1GRAMC = " & gramatura & ", "
    SQL = SQL & "FT1_SAVING = '" & tipoSavings & "', "
    SQL = SQL & "FT1_SAVINGVALK = " & valorSavingSK & ", "
    SQL = SQL & "FT1_SAVINGVALC = " & valorSavingCliente & ", "
    SQL = SQL & "FT1_AcompanharLotePiloto = '" & necessitaAcompanhamento & "', "
    SQL = SQL & "FT1_B5AREA = " & areaLiquida & ", "
    SQL = SQL & "FT1_SUPORTEPRIMARIO = '" & suportacaoPrimario & "', "
    SQL = SQL & "FT1_PRECOKG = " & precoKG & ", "
    SQL = SQL & "FT1_VALORCLICHE = " & valorEstimadoCliche & ", "
    SQL = SQL & "FT1_VALORMATRIZ = " & valorEstimadoMatriz & ", "
    SQL = SQL & "FT1_NIVELLOTEPILOTO = '" & nivelLotePiloto & "' "
    SQL = SQL & "WHERE FT1_ID = '" & id & "' AND FT1_PLANTA = '" & codPlanta & "'"

    On Error Resume Next
    SET rsEditar = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsEditar) Then
        Editar = "{""sucesso"": false, ""mensagem"": ""Erro ao editar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsEditar.Close()
    SET rsEditar = NOTHING
    On Error Goto 0

    If acessoriosJSON <> "" Then
        Call InserirAcessorios(id, acessoriosJSON, codPlanta)
    End If

    If anexosJSON <> "" Then
        Call InserirAnexos(id, anexosJSON, codPlanta)
    End If

    Call RegistrarLog("edicao", id, "", "")
    Editar = "{""sucesso"": true, ""mensagem"": ""Solicitação editada com sucesso!""}"
End Function

Function EditarTestesFisicos()
    id = Request("id")
    codPlanta = Request("codPlanta")

    If id = "" Then
        EditarTestesFisicos = "{""sucesso"": false, ""mensagem"": ""ID da solicitação não informado""}"
        Exit Function
    End If

    If codPlanta = "" Then
        EditarTestesFisicos = "{""sucesso"": false, ""mensagem"": ""Planta da solicitação nao informada""}"
        Exit Function
    End If

    testeColuna = Request("testeColuna")
    testeGramatura = Request("testeGramatura")
    testeCobInterno = Request("testeCobInterno")
    testeCobExterno = Request("testeCobExterno")
    testeDescolamentoCapaInterna = Request("testeDescolamentoCapaInterna")
    testeDescolamentoMiolo1 = Request("testeDescolamentoMiolo1")
    testeDescolamentoMiolo2 = Request("testeDescolamentoMiolo2")
    testeDescolamentoMiolo3 = Request("testeDescolamentoMiolo3")
    testeDescolamentoCapaExterna = Request("testeDescolamentoCapaExterna")
    testeBctCompressao1 = Request("testeBctCompressao1")
    testeBctCompressao2 = Request("testeBctCompressao2")
    testeBctCompressao3 = Request("testeBctCompressao3")
    testePinInterno = Request("testePinInterno")
    testePinExterno = Request("testePinExterno")
    testeForcaVinco1 = Request("testeForcaVinco1")
    testeForcaVinco2 = Request("testeForcaVinco2")
    testeForcaVinco3 = Request("testeForcaVinco3")

    SQL = "UPDATE SKFT1 SET "
    SQL = SQL & "FT1_TESTECOLUNA = " & testeColuna & ", "
    SQL = SQL & "FT1_TESTEGRAMATURA = " & testeGramatura & ", "
    SQL = SQL & "FT1_TESTECOBINTERNO = " & testeCobInterno & ", "
    SQL = SQL & "FT1_TESTECOBEXTERNO = " & testeCobExterno & ", " 
    SQL = SQL & "FT1_TESTEMULLENINTERNOCAPA = " & testeDescolamentoCapaInterna & ", "
    SQL = SQL & "FT1_TESTEMULLENINTERNOMIOLO = " & testeDescolamentoMiolo1 & ", "
    SQL = SQL & "FT1_TESTEMULLENINTERNOCAPAINTERMEDIARIA = " & testeDescolamentoMiolo2 & ", "
    SQL = SQL & "FT1_TESTEMULLENEXTERNOCAPA = " & testeDescolamentoMiolo3 & ", "
    SQL = SQL & "FT1_TESTEMULLENEXTERNOMIOLO = " & testeDescolamentoCapaExterna & ", "
    SQL = SQL & "FT1_COMPRESSAO1 = " & testeBctCompressao1 & ", "
    SQL = SQL & "FT1_COMPRESSAO2 = " & testeBctCompressao2 & ", " 
    SQL = SQL & "FT1_COMPRESSAO3 = " & testeBctCompressao3 & ", "
    SQL = SQL & "FT1_PININTERNO = " & testePinInterno & ", "
    SQL = SQL & "FT1_PINEXTERNO = " & testePinExterno & ", "
    SQL = SQL & "FT1_FORCAVINCO1 = " & testeForcaVinco1 & ", "
    SQL = SQL & "FT1_FORCAVINCO2 = " & testeForcaVinco2 & ", "
    SQL = SQL & "FT1_FORCAVINCO3 = " & testeForcaVinco3
    SQL = SQL & "WHERE FT1_ID = '" & id & "' AND FT1_PLANTA = '" & codPlanta & "'"

    On Error Resume Next
    SET rsEditarTestesFisicos = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsEditarTestesFisicos) Then
        EditarTestesFisicos = "{""sucesso"": false, ""mensagem"": ""Erro ao editar testes físicos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsEditarTestesFisicos.Close()
    SET rsEditarTestesFisicos = NOTHING
    On Error Goto 0

    If acessoriosJSON <> "" Then
        Call InserirAcessorios(id, acessoriosJSON, codPlanta)
    End If

    Call RegistrarLog("edicao_teste_fisico", id, "", "")
    EditarTestesFisicos = "{""sucesso"": true, ""mensagem"": ""Testes físicos editados com sucesso!""}"
End Function

Function AprovacaoDesenvolvimento() 
    id = Request("id")
    codPlanta = Request("codPlanta")
    processo = Request("processoAprovacao")
    
    If id = "" Then
        AprovacaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    If codPlanta = "" Then
        AprovacaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Planta nao informada""}"
        Exit Function
    End If

    respostaObter = ObterPorId(id, codPlanta, 0)
    Set solicitacaoObj = JsonToDictionary(respostaObter)
    
    If Not solicitacaoObj("sucesso") Then
        AprovacaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Solicitação inválida""}"
        Exit Function
    End If

    Set solicitacao = solicitacaoObj("dados")

    If solicitacao("status") <> "F" Then
        AprovacaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Solicitação já foi aprovada/rejeitada anteriormente""}"
        Exit Function
    End If

    If solicitacao("ehPrecadastro") = "P" Then
        AprovacaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Não é possível aprovar uma solicitação de cliente pre-cadastro""}"
        Exit Function
    End If

    numeroFT = ""
    If solicitacao("codPlanta") = "00" OR solicitacao("codPlanta") = "04" Then
        numeroFT = InserirFTProtheus(solicitacao)
    Else
        numeroFT = InserirFTTrimbox(solicitacao)
    End If

    If numeroFT = "" OR numeroFT = "0" Then
        AprovacaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Erro ao gerar FT""}"
        Exit Function
    End If

    SQL = "UPDATE SKFT1 SET "
    SQL = SQL & "FT1_STATUS = 'P', "
    SQL = SQL & "FT1_TipoAlteraCancela = '" & processo & "', "
    SQL = SQL & "FT1_FTNOVA = " & numeroFT & ", "
    SQL = SQL & "FT1_DataGeracaoFT = GETDATE(), "
    SQL = SQL & "FT1_QuemGerouFT = '" & usuarioLogado("login") & "' "
    SQL = SQL & "WHERE FT1_ID = " & solicitacao("id") & "" 

    On Error Resume Next
    SET rsAprovar = Conecta(ObterConexaoCorporativo(solicitacao("codPlanta")), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsAprovar) Then
        AprovacaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Erro ao Aprovar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsAprovar.Close()
    SET rsAprovar = NOTHING
    On Error Goto 0

    Call EnviarEmailAprovacaoDesenvolvimento(solicitacao, numeroFT)
    Call RegistrarLog("aprovar", id, "", "")
    AprovacaoDesenvolvimento = "{""sucesso"": true, ""mensagem"": ""Solicitação aprovada com sucesso!""}"
End Function

Function RejeicaoDesenvolvimento()
    id = Request("id")
    codPlanta = Request("codPlanta")
    motivoRejeicao = Request("motivoRejeicao")
    
    If id = "" Then
        RejeicaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    If codPlanta = "" Then
        RejeicaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Planta nao informada""}"
        Exit Function
    End If

    If motivoRejeicao = "" Then
        RejeicaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Motivo de rejeicao nao informado""}"
        Exit Function
    End If

    respostaObter = ObterPorId(id, codPlanta, 0)
    Set solicitacaoObj = JsonToDictionary(respostaObter)
    
    If Not solicitacaoObj("sucesso") Then
        RejeicaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Solicitação inválida""}"
        Exit Function
    End If
    
    Set solicitacao = solicitacaoObj("dados")

    If solicitacao("status") <> "F" Then
        RejeicaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Solicitação já foi aprovada/rejeitada anteriormente""}"
        Exit Function
    End If
    
    SQL = "UPDATE SKFT1 SET "
    SQL = SQL & "FT1_STATUS = 'X', "
    SQL = SQL & "FT1_MOTCANC = '" & motivoRejeicao & "' "
    SQL = SQL & "WHERE FT1_ID = " & id
    
    On Error Resume Next
    SET rsRejeitar = Conecta(ObterConexaoCorporativo(solicitacao("codPlanta")), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsRejeitar) Then
        RejeicaoDesenvolvimento = "{""sucesso"": false, ""mensagem"": ""Erro ao Rejeitar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsRejeitar.Close()
    SET rsRejeitar = NOTHING
    On Error Goto 0

    Call EnviarEmailRejeicao(solicitacao, motivoRejeicao)
    Call RegistrarLog("rejeitar", id, "", "")
    RejeicaoDesenvolvimento = "{""sucesso"": true, ""mensagem"": ""Solicitação rejeitada com sucesso!""}"
End Function

Function AprovacaoComercial()
    id = Request("id")
    codPlanta = Request("codPlanta")
    
    If id = "" Then
        AprovacaoComercial = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    If codPlanta = "" Then
        AprovacaoComercial = "{""sucesso"": false, ""mensagem"": ""Planta nao informada""}"
        Exit Function
    End If

    respostaObter = ObterPorId(id, codPlanta, 0)
    Set solicitacaoObj = JsonToDictionary(respostaObter)
    
    If Not solicitacaoObj("sucesso") Then
        AprovacaoComercial = "{""sucesso"": false, ""mensagem"": ""Solicitação inválida""}"
        Exit Function
    End If
    
    Set solicitacao = solicitacaoObj("dados")

    If solicitacao("status") <> "S" Then
        AprovacaoComercial = "{""sucesso"": false, ""mensagem"": ""Solicitação já foi aprovada/rejeitada anteriormente""}"
        Exit Function
    End If

    nomeUsuario = Mid(usuarioLogado("nome"), 1, 20)

    SQL = "INSERT INTO WorkFlowSapComercial (idFT, Usuario, Status, MotivoReprovacao, DataStatus, HoraStatus) VALUES ("
    SQL = SQL & "'" & solicitacao("id") & "', "
    SQL = SQL & "'" & nomeUsuario & "', "
    SQL = SQL & "'C', "
    SQL = SQL & "'', "
    SQL = SQL & "FORMAT(GETDATE(), 'yyyyMMdd'), "
    SQL = SQL & "FORMAT(GETDATE(), 'HH:mm:ss')"
    SQL = SQL & "); "

    SQL = SQL & "UPDATE SKFT1 SET " 
    SQL = SQL & "FT1_STATUS = 'A' "
    SQL = SQL & "WHERE FT1_ID = " & solicitacao("id") & " "

    On Error Resume Next
    SET rsAprovar = Conecta(ObterConexaoCorporativo(solicitacao("codPlanta")), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsAprovar) Then
        AprovacaoComercial = "{""sucesso"": false, ""mensagem"": ""Erro ao Aprovar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsAprovar.Close()
    SET rsAprovar = NOTHING
    On Error Goto 0

    Call EnviarEmailAprovacaoComercial(solicitacao)
    Call RegistrarLog("aprovar_comercial", id, "", "")
    AprovacaoComercial = "{""sucesso"": true, ""mensagem"": ""Solicitação aprovada com sucesso!""}"
End Function

Function RejeicaoComercial()
    id = Request("id")
    codPlanta = Request("codPlanta")
    motivoRejeicao = Request("motivoRejeicao")
    
    If id = "" Then
        RejeicaoComercial = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    If codPlanta = "" Then
        RejeicaoComercial = "{""sucesso"": false, ""mensagem"": ""Planta nao informada""}"
        Exit Function
    End If

    If motivoRejeicao = "" Then
        RejeicaoComercial = "{""sucesso"": false, ""mensagem"": ""Motivo de rejeicao nao informado""}"
        Exit Function
    End If

    respostaObter = ObterPorId(id, codPlanta, 0)
    Set solicitacaoObj = JsonToDictionary(respostaObter)
    
    If Not solicitacaoObj("sucesso") Then
        RejeicaoComercial = "{""sucesso"": false, ""mensagem"": ""Solicitação inválida""}"
        Exit Function
    End If
    
    Set solicitacao = solicitacaoObj("dados")

    If solicitacao("status") <> "S" Then
        RejeicaoComercial = "{""sucesso"": false, ""mensagem"": ""Solicitação já foi aprovada/rejeitada anteriormente""}"
        Exit Function
    End If

    nomeUsuario = Mid(usuarioLogado("nome"), 1, 20)

    SQL = "INSERT INTO WorkFlowSapComercial (idFT, Usuario, Status, MotivoReprovacao, DataStatus, HoraStatus) VALUES ("
    SQL = SQL & "'" & id & "', "
    SQL = SQL & "'" & nomeUsuario & "', "
    SQL = SQL & "'V', "
    SQL = SQL & "'" & motivoRejeicao & "', "
    SQL = SQL & "FORMAT(GETDATE(), 'yyyyMMdd'), "
    SQL = SQL & "FORMAT(GETDATE(), 'HH:mm:ss')"
    SQL = SQL & "); "

    SQL = SQL & "UPDATE SKFT1 SET " 
    SQL = SQL & "FT1_STATUS = 'V' "
    SQL = SQL & "WHERE FT1_ID = " & id & " "

    On Error Resume Next
    SET rsRejeitar = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsRejeitar) Then
        RejeicaoComercial = "{""sucesso"": false, ""mensagem"": ""Erro ao rejeitar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsRejeitar.Close()
    SET rsRejeitar = NOTHING
    On Error Goto 0

    Call EnviarEmailRejeicao(solicitacao, motivoRejeicao)
    Call RegistrarLog("rejeitar_comercial", id, "", "")
    RejeicaoComercial = "{""sucesso"": true, ""mensagem"": ""Solicitação rejeitada com sucesso!""}"
End Function

Function AprovacaoPCP()
    id = Request("id")
    codPlanta = Request("codPlanta")
    
    If id = "" Then
        AprovacaoPCP = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    If codPlanta = "" Then
        AprovacaoPCP = "{""sucesso"": false, ""mensagem"": ""Planta nao informada""}"
        Exit Function
    End If

    respostaObter = ObterPorId(id, codPlanta, 0)
    Set solicitacaoObj = JsonToDictionary(respostaObter)
    
    If Not solicitacaoObj("sucesso") Then
        AprovacaoPCP = "{""sucesso"": false, ""mensagem"": ""Solicitação inválida""}"
        Exit Function
    End If
    
    Set solicitacao = solicitacaoObj("dados")

    If solicitacao("status") <> "D" Then
        AprovacaoPCP = "{""sucesso"": false, ""mensagem"": ""Solicitação já foi aprovada/rejeitada anteriormente""}"
        Exit Function
    End If

    nomeUsuario = Mid(usuarioLogado("nome"), 1, 20)

    SQL = "INSERT INTO WorkFlowSapComercial (idFT, Usuario, Status, MotivoReprovacao, DataStatus, HoraStatus) VALUES ("
    SQL = SQL & "'" & solicitacao("id") & "', "
    SQL = SQL & "'" & nomeUsuario & "', "
    SQL = SQL & "'C', "
    SQL = SQL & "'', "
    SQL = SQL & "FORMAT(GETDATE(), 'yyyyMMdd'), "
    SQL = SQL & "FORMAT(GETDATE(), 'HH:mm:ss')"
    SQL = SQL & "); "

    SQL = SQL & "UPDATE SKFT1 SET " 
    SQL = SQL & "FT1_STATUS = 'F' "
    SQL = SQL & "WHERE FT1_ID = " & solicitacao("id") & " "

    On Error Resume Next
    SET rsAprovar = Conecta(ObterConexaoCorporativo(solicitacao("codPlanta")), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsAprovar) Then
        AprovacaoPCP = "{""sucesso"": false, ""mensagem"": ""Erro ao Aprovar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsAprovar.Close()
    SET rsAprovar = NOTHING
    On Error Goto 0

    Call EnviarEmailAprovacaoPCP(solicitacao)
    Call RegistrarLog("aprovar_pcp", id, "", "")
    AprovacaoPCP = "{""sucesso"": true, ""mensagem"": ""Solicitação aprovada com sucesso!""}"
End Function

Function RejeicaoPCP()
    id = Request("id")
    codPlanta = Request("codPlanta")
    motivoRejeicao = Request("motivoRejeicao")
    
    If id = "" Then
        RejeicaoPCP = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    If codPlanta = "" Then
        RejeicaoPCP = "{""sucesso"": false, ""mensagem"": ""Planta nao informada""}"
        Exit Function
    End If

    If motivoRejeicao = "" Then
        RejeicaoPCP = "{""sucesso"": false, ""mensagem"": ""Motivo de rejeicao nao informado""}"
        Exit Function
    End If

    respostaObter = ObterPorId(id, codPlanta, 0)
    Set solicitacaoObj = JsonToDictionary(respostaObter)
    
    If Not solicitacaoObj("sucesso") Then
        RejeicaoPCP = "{""sucesso"": false, ""mensagem"": ""Solicitação inválida""}"
        Exit Function
    End If
    
    Set solicitacao = solicitacaoObj("dados")

    If solicitacao("status") <> "D" Then
        RejeicaoPCP = "{""sucesso"": false, ""mensagem"": ""Solicitação já foi aprovada/rejeitada anteriormente""}"
        Exit Function
    End If

    nomeUsuario = Mid(usuarioLogado("nome"), 1, 20)

    SQL = "INSERT INTO WorkFlowSapComercial (idFT, Usuario, Status, MotivoReprovacao, DataStatus, HoraStatus) VALUES ("
    SQL = SQL & "'" & id & "', "
    SQL = SQL & "'" & nomeUsuario & "', "
    SQL = SQL & "'V', "
    SQL = SQL & "'" & motivoRejeicao & "', "
    SQL = SQL & "FORMAT(GETDATE(), 'yyyyMMdd'), "
    SQL = SQL & "FORMAT(GETDATE(), 'HH:mm:ss')"
    SQL = SQL & "); "

    SQL = SQL & "UPDATE SKFT1 SET " 
    SQL = SQL & "FT1_STATUS = 'G' "
    SQL = SQL & "WHERE FT1_ID = " & id & " "

    On Error Resume Next
    SET rsRejeitar = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsRejeitar) Then
        RejeicaoPCP = "{""sucesso"": false, ""mensagem"": ""Erro ao rejeitar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsRejeitar.Close()
    SET rsRejeitar = NOTHING
    On Error Goto 0

    Call EnviarEmailRejeicaoPCP(solicitacao, motivoRejeicao)
    Call RegistrarLog("rejeitar_comercial", id, "", "")
    RejeicaoPCP = "{""sucesso"": true, ""mensagem"": ""Solicitação rejeitada com sucesso!""}"
End Function

Function Excluir()
    id = Request("id")
    codPlanta = Request("codPlanta")
    
    If id = "" Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    If codPlanta = "" Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Planta nao informada""}"
        Exit Function
    End If

    SQL = "UPDATE SKFT1 SET "
    SQL = SQL & "FT1_DEL = 1, "
    SQL = SQL & "FT1_QuemExcluiu = '" & usuarioLogado("login") & "', "
    SQL = SQL & "FT1_DataExclusao = GETDATE() "
    SQL = SQL & "WHERE FT1_ID = " & id
    
    On Error Resume Next
    SET rsExcluir = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsExcluir) Then
        Excluir = "{""sucesso"": false, ""mensagem"": ""Erro ao excluir solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsExcluir.Close()
    SET rsExcluir = NOTHING
    On Error Goto 0
    
    Call RegistrarLog("exclusao", id, "", "")
    Excluir = "{""sucesso"": true, ""mensagem"": ""Solicitação excluída com sucesso!""}"
End Function

Function Cancelar()
    id = Request("id")
    codPlanta = Request("codPlanta")
    
    If id = "" Then
        Cancelar = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    If codPlanta = "" Then
        Cancelar = "{""sucesso"": false, ""mensagem"": ""Planta nao informada""}"
        Exit Function
    End If

    SQL = "UPDATE SKFT1 SET "
    SQL = SQL & "FT1_STATUS = 'K' "
    SQL = SQL & "WHERE FT1_ID = " & id
    
    On Error Resume Next
    SET rsReverter = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsReverter) Then
        Cancelar = "{""sucesso"": false, ""mensagem"": ""Erro ao cancelar solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsReverter.Close()
    SET rsReverter = NOTHING
    On Error Goto 0
    
    Call RegistrarLog("cancelar", id, "", "")
    Cancelar = "{""sucesso"": true, ""mensagem"": ""Solicitação cancelada com sucesso!""}"
End Function

Function ReverterCancelamento()
    id = Request("id")
    codPlanta = Request("codPlanta")
    
    If id = "" Then
        ReverterCancelamento = "{""sucesso"": false, ""mensagem"": ""Id não informado""}"
        Exit Function
    End If

    If codPlanta = "" Then
        ReverterCancelamento = "{""sucesso"": false, ""mensagem"": ""Planta nao informada""}"
        Exit Function
    End If

    SQL = "UPDATE SKFT1 SET "
    SQL = SQL & "FT1_STATUS = 'A' "
    SQL = SQL & "WHERE FT1_ID = " & id
    
    On Error Resume Next
    SET rsReverter = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsReverter) Then
        ReverterCancelamento = "{""sucesso"": false, ""mensagem"": ""Erro ao reverter cancelamento: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsReverter.Close()
    SET rsReverter = NOTHING
    On Error Goto 0
    
    Call RegistrarLog("reversao_cancelamento", id, "", "")
    ReverterCancelamento = "{""sucesso"": true, ""mensagem"": ""Solicitação revertida com sucesso!""}"
End Function


Function AlterarStatusAndamento()
    id = Request("id")
    codPlanta = Request("codPlanta")
    novoStatus = Request("status")

    If id = "" Then
        AlterarStatusAndamento = "{""sucesso"": false, ""mensagem"": ""ID da solicitação não informado""}"
        Exit Function
    End If

    If codPlanta = "" Then
        AlterarStatusAndamento = "{""sucesso"": false, ""mensagem"": ""Planta nao informada""}"
        Exit Function
    End If

    If novoStatus = "" Then
        AlterarStatusAndamento = "{""sucesso"": false, ""mensagem"": ""Status da solicitação nao informada""}"
        Exit Function
    End If

    dataEnvioAmostra = Request("dataEnvioAmostra")
    quemEnviouAmostra = Request("quemEnviouAmostra")
    motivoReprovacaoAmostra = Request("motivoReprovacaoAmostra")
    dataConfeccaoFerramental = Request("dataConfeccaoFerramental")
    valorCliche = Request("valorCliche")
    valorMatriz = Request("valorMatriz")
    observacoes = Request("observacoes")
    statusAntigo = ""

    ' Validação amostra aprovada
    If novoStatus = "E" AND (dataEnvioAmostra = "" OR quemEnviouAmostra = "") Then
        AlterarStatusAndamento = "{""sucesso"": false, ""mensagem"": ""Data de envio de amostra nao informada ou quem enviou nao informado""}"
        Exit Function
    End If

    ' Validação amostra reprovada
    If novoStatus = "R" AND motivoReprovacaoAmostra = "" Then
        AlterarStatusAndamento = "{""sucesso"": false, ""mensagem"": ""Motivo de reprovação de amostra nao informado""}"
        Exit Function
    End If

    ' Validação confeccao ferramental
    If novoStatus = "J" AND (dataConfeccaoFerramental = "" OR valorCliche = "" OR valorMatriz = "") Then
        AlterarStatusAndamento = "{""sucesso"": false, ""mensagem"": ""Data de confeccao de ferramental nao informada""}"
        Exit Function
    End If

    SQL = "SELECT FT1_STATUS2 FROM SKFT1 WHERE FT1_ID = '" & id & "' AND FT1_PLANTA = '" & codPlanta & "'"

    On Error Resume Next
    SET rsObterStatus = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsObterStatus) Then
        AlterarStatusAndamento = "{""sucesso"": false, ""mensagem"": ""Erro ao alterar status solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    statusAntigo = rsObterStatus("FT1_STATUS2")
    
    rsObterStatus.Close()
    SET rsObterStatus = NOTHING
    On Error Goto 0

    SQL = "BEGIN TRANSACTION TRANS; " 

    SQL = SQL & "UPDATE SKFT1 SET "
    SQL = SQL & "FT1_STATUS2 = '" & novoStatus & "'"

    If novoStatus = "E" Then
        SQL = SQL & ", FT1_DataEnvioAmostra = '" & dataEnvioAmostra & "'"
        SQL = SQL & ", FT1_QuemEnviouAmostra = '" & quemEnviouAmostra & "'"
    End If

    If novoStatus = "R" Then
        SQL = SQL & ", FT1_MOTIVOREPROVACAOAMOSTRA2 = '" & motivoReprovacaoAmostra & "'"
    End If

    If novoStatus = "J" Then
        SQL = SQL & ", FT1_DTCONFEC = '" & dataConfeccaoFerramental & "'"
        SQL = SQL & ", FT1_VALORCLICHE = '" & valorCliche & "'"
        SQL = SQL & ", FT1_VALORMATRIZ = '" & valorMatriz & "'"
    End If

    SQL = SQL & "WHERE FT1_ID = '" & id & "' AND FT1_PLANTA = '" & codPlanta & "'; "

    ' Atualiza registro anterior do histórico com data atual (Fazemos isso para contabilizar o leadtime)
    ' Ou seja, Data de alteração do registro anterior sempre será igual a data de criação do registro atual
    SQL = SQL & "UPDATE HistoricoAlteracaoProcessoSap SET "
    SQL = SQL & "DataAlteracaoProcesso = GETDATE(), "
    SQL = SQL & "QuemAlterou = '" & usuarioLogado("login") & "'"
    SQL = SQL & "WHERE IdHistorico = ("
    SQL = SQL & "   SELECT TOP 1 IdHistorico "
    SQL = SQL & "   FROM HistoricoAlteracaoProcessoSap "
    SQL = SQL & "   WHERE IdSap = '" & id & "'"
    SQL = SQL & "   ORDER BY IdHistorico DESC"
    SQL = SQL & "); "

    ' Insere novo registro de alteração no histórico
    SQL = SQL & "INSERT INTO HistoricoAlteracaoProcessoSap (" 
    SQL = SQL & "IdSap, IdProcesso, DataInsercaoProcesso, " 
    SQL = SQL & "QuemAlterou, Complemento "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'" & id & "',"
    SQL = SQL & "'" & novoStatus & "',"
    SQL = SQL & "GETDATE()" & ","
    SQL = SQL & "'" & usuarioLogado("login") & "',"
    SQL = SQL & "'" & observacoes & "'); "

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION TRANS "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION TRANS "

    On Error Resume Next
    SET rsInserirHistorico = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsInserirHistorico) Then
        AlterarStatusAndamento = "{""sucesso"": false, ""mensagem"": ""Erro ao alterar status solicitação: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    
    rsInserirHistorico.Close()
    SET rsInserirHistorico = NOTHING
    On Error Goto 0

    Call RegistrarLog("alteracao_status", id, novoStatus, status)
    AlterarStatusAndamento = "{""sucesso"": true, ""mensagem"": ""Status alterado com sucesso!""}"
End Function

Function ListarEstilos()
    Dim SQL, rsEstilo, jsonArray, primeiroItem
    Dim filtroCodPlanta, stringConexao

    filtroCodPlanta = Request("filtroCodPlanta")

    If filtroCodPlanta = "" Then
        ListarEstilos = "{""sucesso"": false, ""mensagem"": ""Planta não informada""}"
        Exit Function
    End If

    Select Case filtroCodPlanta
        Case "00"
            SQL = "SELECT Z2_CODIGO AS Codigo, REPLACE(Z2_DESC, '\', '') as Descricao FROM SZ2010 WHERE D_E_L_E_T_ = ''"
            stringConexao = "PIR-DADOSADV"
        Case "04"
            SQL = "SELECT Z2_CODIGO AS Codigo, REPLACE(Z2_DESC, '\', '') as Descricao FROM SZ2010 WHERE D_E_L_E_T_ = ''"
            stringConexao = "PIR-DADOSADV"
        Case "05"
            SQL = "SELECT CodTrimbox as Codigo, REPLACE(Descricao, '\', '') as Descricao FROM vw_ConsultaTipoCaixaPelaInpaNet WHERE Desativado = 0 "
            stringConexao = "BEN-TRIMBOX"
        Case "06"
            SQL = "SELECT CodTrimbox as Codigo, REPLACE(Descricao, '\', '') as Descricao FROM vw_ConsultaTipoCaixaPelaInpaNet WHERE Desativado = 0 "
            stringConexao = "FOR-TRIMBOX"
        Case "07"
            SQL = "SELECT CodTrimbox as Codigo, REPLACE(Descricao, '\', '') as Descricao FROM vw_ConsultaTipoCaixaPelaInpaNet WHERE Desativado = 0 "
            stringConexao = "SAQ-TRIMBOX"
        Case Else
            ListarEstilos = "{""sucesso"": false, ""mensagem"": ""Planta inválida""}"
    End Select

    On Error Resume Next
    SET rsEstilo = Conecta(stringConexao, SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsEstilo) Then
        ListarEstilos = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de estilos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsEstilo.EOF Then
        Do While Not rsEstilo.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """codigo"": """ & EscapeJSON(rsEstilo("Codigo") & "") & """, "
            jsonArray = jsonArray & """descricao"": """ & trim(EscapeJSON(rsEstilo("Descricao")) & "") & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsEstilo.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsEstilo.Close()
    SET rsEstilo = NOTHING
    
    ListarEstilos = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarEstruturasPapelao()
    Dim SQL, rsEstrutura, jsonArray, primeiroItem
    Dim filtroCodPlanta, stringConexao

    filtroCodPlanta = Request("filtroCodPlanta")

    If filtroCodPlanta = "" Then
        ListarEstruturasPapelao = "{""sucesso"": false, ""mensagem"": ""Planta não informada""}"
        Exit Function
    End If

    Select Case filtroCodPlanta
        Case "00"
            stringConexao = "PIR-DADOSADV"
            SQL = "SELECT SUBSTRING(Z1_ESTRUT, 1, LEN(Z1_ESTRUT)-1) + 'A' AS Codigo, '' AS NomeComercial, Z1_GRAM_C AS Gramatura FROM SZ1010 WHERE Z1_STATUS IN ('I','T') AND D_E_L_E_T_ = '' "
        Case "04"
            stringConexao = "PIR-DADOSADV"
            SQL = "SELECT SUBSTRING(Z1_ESTRUT, 1, LEN(Z1_ESTRUT)-1) + 'B' AS Codigo, '' AS NomeComercial, Z1_GRAM_C AS Gramatura FROM SZ1010 WHERE Z1_STATUS IN ('I','T') AND D_E_L_E_T_ = '' "
        Case "05"
            stringConexao = "BEN-TRIMBOX"
            SQL = "SELECT Composicao as Codigo, NomeComercial, Gramatura FROM VW_ESTRUTURA_SAP WHERE Desativada = 0 ORDER BY CodigoERP"
        Case "06"
            stringConexao = "FOR-TRIMBOX"
            SQL = "SELECT Composicao as Codigo, NomeComercial, Gramatura FROM VW_ESTRUTURA_SAP WHERE Desativada = 0 ORDER BY CodigoERP"
        Case "07"
            stringConexao = "SAQ-TRIMBOX"
            SQL = "SELECT Composicao as Codigo, NomeComercial, Gramatura FROM VW_ESTRUTURA_SAP WHERE Desativada = 0 ORDER BY CodigoERP"
        Case Else
            ListarEstruturasPapelao = "{""sucesso"": false, ""mensagem"": ""Planta inválida""}"
    End Select

    On Error Resume Next
    SET rsEstrutura = Conecta(stringConexao, SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsEstrutura) Then
        ListarEstruturasPapelao = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de estruturas de papelão: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsEstrutura.EOF Then
        Do While Not rsEstrutura.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """codigo"": """ & EscapeJSON(rsEstrutura("Codigo") & "") & """, "
            jsonArray = jsonArray & """gramatura"": """ & trim(rsEstrutura("Gramatura")) & """, "
            jsonArray = jsonArray & """nomeComercial"": """ & trim(EscapeJSON(rsEstrutura("NomeComercial")) & "") & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsEstrutura.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsEstrutura.Close()
    SET rsEstrutura = NOTHING
    
    ListarEstruturasPapelao = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarFichasTecnicas()
    filtroCodPlanta = Request("filtroCodPlanta")
    filtroCodCliente = Request("filtroCodCliente")
    busca = Request("busca")
    stringConexao = ""

    If filtroCodPlanta = "" Then
        ListarFichasTecnicas = "{""sucesso"": false, ""mensagem"": ""Planta não informada""}"
        Exit Function
    End If

    If filtroCodCliente = "" Then
        ListarFichasTecnicas = "{""sucesso"": false, ""mensagem"": ""Cliente não informado""}"
        Exit Function
    End If

    If busca = "" Then
        ListarFichasTecnicas = "{""sucesso"": false, ""mensagem"": ""Busca nao informada""}"
        Exit Function
    End If

    Select Case filtroCodPlanta
        Case "05"
            tipoItem = "('65','14','CN','CV')"
        Case "06"
            tipoItem = "('12','27','CN','CV')"
        Case "07"
            tipoItem = "('CN','CV')"
    end Select 

    If filtroCodPlanta = "00" OR filtroCodPlanta = "04" Then
        stringConexao = ObterConexaoProtheus(filtroCodPlanta)

        SQL = "SELECT DISTINCT TOP 100"
        SQL = SQL & "	B5_COD AS Item," 
        SQL = SQL & "	B5_CLI as CodCliente, " 
        SQL = SQL & "   B5_NOMECLI as NomeCliente,"
        SQL = SQL & "	B5_LOJA AS CodLoja, "
        SQL = SQL & "   REPLACE(REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), B5_REF), CHAR(7), ''), CHAR(34), ''), CHAR(0), ''), CHAR(9), '') AS Referencia, "
        SQL = SQL & "	B5_TOLMA as ToleranciaMaxima, "
        SQL = SQL & "	B5_TOLME as ToleranciaMinima, "
        SQL = SQL & "	B5_ESTILO as Estilo, "
        SQL = SQL & "	B5_COMPR as Comprimento,"
        SQL = SQL & "	B5_LARG as Largura,"
        SQL = SQL & "	B5_ALTU as Altura,"
        SQL = SQL & "	B5_NUMFACA as NumeroMatriz,"
        SQL = SQL & "	B5_ESTRUTU as EstruturaPapelao,"
        SQL = SQL & "	B5_GRAM as GramaturaReferencial,"
        SQL = SQL & "	B5_ESMAG as Esmagamento,"
        SQL = SQL & "	B5_COLUNA as EctColuna,"
        SQL = SQL & "	B5_ESPES2 as Espessura," 
        SQL = SQL & "	B5_COMPRE as BctCompressao,"
        SQL = SQL & "	B5_COBINT as CobbInterno," 
        SQL = SQL & "	B5_COBEXT as CobbExterno,"
        SQL = SQL & "	B5_FECHAM as TipoFechamento, " 
        SQL = SQL & "	B5_PALETE as EhPaletizado, "
        SQL = SQL & "	B1_ALIMENT as ConteudoEmbalado, "
        SQL = SQL & "	B1_TIPCONV as TipoLinhaProducaoCliente, "
        SQL = SQL & "	B5_AMARRAD as  EhAmarrado, "
        SQL = SQL & "	0 as AlturaMaximaPalete, "
        SQL = SQL & "	B5_QTDPCT AS PecasPorUnidade, "
        SQL = SQL & "	B5_TPPALE as Palete, "
        SQL = SQL & "	B5_FACOP as TemFacesOpostas, " 
        SQL = SQL & "	B5_CANTON as TemCantinho, " 
        SQL = SQL & "	B5_TAMPA as TemGradeMadeira, " 
        SQL = SQL & "	B5_FILME as TemFilmeStrech, " 
        SQL = SQL & "	B5_ESPELHO as QuantidadeEspelhos, "
        SQL = SQL & "	'' as ConteudoEmbalado, " 
        SQL = SQL & "	B5_FITLAR as LarguraFitilho, " 
        SQL = SQL & "	B5_FITCOM ComprimentoFitilho, "
        SQL = SQL & "	B5_DES as TesteForcaVinco1, "
        SQL = SQL & "	B5_EMB1 as TesteForcaVinco2, "
        SQL = SQL & "	B5_EMB2 as TesteForcaVinco3, "
        SQL = SQL & "	0 as QuantidadeCores, "
        SQL = SQL & "	B5_VEND as IdVendedor, " 
        SQL = SQL & "	B5_NOMVEN as NomeVendedor, "
        SQL = SQL & "   B5_CERT AS TestePinInterno, "
        SQL = SQL & "   B5_ESTMAT AS TestePinExterno, "
        SQL = SQL & "   CASE B1_INPUBE WHEN 'U' THEN B1_INPUBE ELSE 'P' END UnidadeFabricacao "
        SQL = SQL & "FROM SB5010 "
        SQL = SQL & "LEFT JOIN SB1010 ON SB1010.B1_COD = SB5010.B5_COD "
        SQL = SQL & "WHERE (B5_COD LIKE '%" & busca & "%' OR B5_REF LIKE '%" & busca & "%') "
        SQL = SQL & "AND B5_CLI = '" & filtroCodCliente & "' "

        If usuarioLogado("nivel") = 1 Then
            SQL = SQL & "AND A1_VEND = '" & usuarioLogado("cod") & "' "
        End If

        SQL = SQL & "ORDER BY Item DESC"
    Else
        stringConexao = ObterConexaoTrimbox(filtroCodPlanta)

        SQL = "SELECT DISTINCT TOP 100"
        SQL = SQL & "	B1.Conjunto as Item," 
        SQL = SQL & "	A1.CodCliente as CodCliente, " 
        SQL = SQL & "   A1.RazaoSocial as NomeCliente,"
        SQL = SQL & "	A1.CodLoja, "
        SQL = SQL & "	B1.Referencia, "
        SQL = SQL & "	B1.TolMais as ToleranciaMaxima, "
        SQL = SQL & "	B1.TolMenos as ToleranciaMinima, "
        SQL = SQL & "	B1.TipoABNT as Estilo, "
        SQL = SQL & "	B1.ComprimentoInterno as Comprimento,"
        SQL = SQL & "	B1.LarguraInterna as Largura,"
        SQL = SQL & "	B1.AlturaInterna as Altura,"
        SQL = SQL & "	B1.Faca as NumeroMatriz,"
        SQL = SQL & "	B1.Composicao as EstruturaPapelao,"
        SQL = SQL & "	B1.Gramatura as GramaturaReferencial,"
        SQL = SQL & "	B1.Compressao as Esmagamento,"
        SQL = SQL & "	B1.ColunaMinimo as EctColuna,"
        SQL = SQL & "	'' Espessura," 
        SQL = SQL & "	CASE WHEN B1.Compressao > 0 THEN (B1.Compressao / 9.81) ELSE B1.Compressao END as BctCompressao,"
        SQL = SQL & "	B1.CobbIntMaximo as CobbInterno," 
        SQL = SQL & "	B1.CobbExtMaximo as CobbExterno,"
        SQL = SQL & "	Left(B1.Fechamento,1) as TipoFechamento, " 
        SQL = SQL & "	CASE B1.Paletizado WHEN 0 then 'N' ELSE 'S' END EhPaletizado, "
        SQL = SQL & "   CASE WHEN (FitilhosCompPacote <> 0 OR FitilhosLargPacote <> 0) THEN 'S' ELSE '' END AS EhAmarrado,"
        SQL = SQL & "	AlturaPaleteFechado as AlturaMaximaPalete, "
        SQL = SQL & "	B1.PecasPorPacote AS PecasPorUnidade, "
        SQL = SQL & "	B1.IDPalete as Palete, "
        SQL = SQL & "	CASE B1.FacesOpostas WHEN 0 then 'N' ELSE 'S' END TemFacesOpostas, " 
        SQL = SQL & "	CASE B1.Cantoneiras WHEN 0 then 'N' ELSE 'S' END TemCantinho, " 
        SQL = SQL & "	'' as TemGradeMadeira, " 
        SQL = SQL & "	CASE B1.Filme WHEN 0 then 'N' ELSE 'S' END TemFilmeStrech, " 
        SQL = SQL & "	B1.Espelho as QuantidadeEspelhos, "
        SQL = SQL & "	'' as ConteudoEmbalado, " 
        SQL = SQL & "	'' as TipoLinhaProducaoCliente, "
        SQL = SQL & "	isnull(B1.FitilhosLargPacote,0) as LarguraFitilho, " 
        SQL = SQL & "	isnull(B1.FitilhosCompPacote,0) as ComprimentoFitilho, " 
        SQL = SQL & "	isnull(B1.ForcaDeVinco1,0) as TesteForcaVinco1, " 
        SQL = SQL & "	isnull(B1.ForcaDeVinco2,0) as TesteForcaVinco2, " 
        SQL = SQL & "	isnull(B1.ForcaDeVinco3,0) as TesteForcaVinco3, " 
        SQL = SQL & "	0 as TestePinInterno, "
        SQL = SQL & "	0 as TestePinExterno, "
        SQL = SQL & "	B1.NrCores as QuantidadeCores, "
        SQL = SQL & "	B1.CodRepresentante as IdVendedor, "
        SQL = SQL & "   'B' AS UnidadeFabricacao, "  
        SQL = SQL & "	A3.Nome as NomeVendedor "
        SQL = SQL & "FROM vw_ConsultaFTs B1 "
        SQL = SQL & "INNER JOIN vw_ConsultaLojasTrimBox A1 ON A1.CodCliente = B1.CodCliente "
        SQL = SQL & "LEFT JOIN vw_ConsultaRepresentantesPelaInpaNet A3 on B1.CodRepresentante = A3.Codigo "
        SQL = SQL & "WHERE (Item LIKE '%" & busca & "%' OR Referencia LIKE '%" & busca & "%') "
        SQL = SQL & "AND TipoABNT in " & tipoItem & " "
        SQL = SQL & "AND A1.CodCliente = '" & filtroCodCliente & "' "

        If usuarioLogado("nivel") = 1 Then
            SQL = SQL & "AND B1.CodRepresentante = '" & usuarioLogado("cod") & "' "
        End If

        SQL = SQL & "ORDER BY Item DESC"
    End If

    On Error Resume Next
    SET rsFicha = Conecta(stringConexao, SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsFicha) Then
        ListarFichasTecnicas = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de estruturas de papelão: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsFicha.EOF Then
        Do While Not rsFicha.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            quantidadeFitilhos = 0
            tipoFitilho = ""

            If Clng(rsFicha("LarguraFitilho")) > 0 Then
                quantidadeFitilhos = rsFicha("LarguraFitilho")
                tipoFitilho = "L"
            ElseIf Clng(rsFicha("ComprimentoFitilho")) > 0 Then
                quantidadeFitilhos = rsFicha("ComprimentoFitilho")
                tipoFitilho = "C"
            End If

            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """item"": """ & EscapeJSON(rsFicha("Item") & "") & """, "
            jsonArray = jsonArray & """nomeCliente"": """ & EscapeJSON(rsFicha("NomeCliente") & "") & """, "
            jsonArray = jsonArray & """codCliente"": """ & EscapeJSON(rsFicha("CodCliente") & "") & """, "
            jsonArray = jsonArray & """codLoja"": """ & EscapeJSON(rsFicha("CodLoja") & "") & """, "
            jsonArray = jsonArray & """referencia"": """ & EscapeJSON(rsFicha("Referencia") & "") & """, "
            jsonArray = jsonArray & """toleranciaMaxima"": """ & Replace(EscapeJSON(rsFicha("ToleranciaMaxima")), ",", ".") & """, "
            jsonArray = jsonArray & """toleranciaMinima"": """ & Replace(EscapeJSON(rsFicha("ToleranciaMinima")), ",", ".") & """, "
            jsonArray = jsonArray & """estilo"": """ & EscapeJSON(rsFicha("Estilo") & "") & """, "
            jsonArray = jsonArray & """comprimento"": """ & Replace(EscapeJSON(rsFicha("Comprimento")), ",", ".") & """, "
            jsonArray = jsonArray & """largura"": """ & Replace(EscapeJSON(rsFicha("Largura")), ",", ".") & """, "
            jsonArray = jsonArray & """altura"": """ & Replace(EscapeJSON(rsFicha("Altura")), ",", ".") & """, "
            jsonArray = jsonArray & """numeroMatriz"": """ & EscapeJSON(rsFicha("NumeroMatriz") & "") & """, "
            jsonArray = jsonArray & """estruturaPapelao"": """ & EscapeJSON(rsFicha("EstruturaPapelao") & "") & """, "
            jsonArray = jsonArray & """gramaturaReferencial"": """ & Replace(EscapeJSON(rsFicha("GramaturaReferencial")), ",", ".") & """, "
            jsonArray = jsonArray & """esmagamento"": """ & Replace(EscapeJSON(rsFicha("Esmagamento")), ",", ".") & """, "
            jsonArray = jsonArray & """ectColuna"": """ & Replace(EscapeJSON(rsFicha("EctColuna")), ",", ".") & """, "
            jsonArray = jsonArray & """espessura"": """ & Replace(EscapeJSON(rsFicha("Espessura")), ",", ".") & """, "
            jsonArray = jsonArray & """bctCompressao"": """ & Replace(EscapeJSON(rsFicha("BctCompressao")), ",", ".") & """, "
            jsonArray = jsonArray & """cobbInterno"": """ & Replace(EscapeJSON(rsFicha("CobbInterno")), ",", ".") & """, "
            jsonArray = jsonArray & """cobbExterno"": """ & Replace(EscapeJSON(rsFicha("CobbExterno")), ",", ".") & """, "
            jsonArray = jsonArray & """gramaturaReferencial"": """ & Replace(EscapeJSON(rsFicha("GramaturaReferencial")), ",", ".") & """, "
            jsonArray = jsonArray & """gramaturaReferencial"": """ & Replace(EscapeJSON(rsFicha("GramaturaReferencial")), ",", ".") & """, "
            jsonArray = jsonArray & """tipoFechamento"": """ & EscapeJSON(rsFicha("TipoFechamento") & "") & """, "
            jsonArray = jsonArray & """ehPaletizado"": """ & EscapeJSON(rsFicha("EhPaletizado") & "") & """, "
            jsonArray = jsonArray & """ehAmarrado"": """ & EscapeJSON(rsFicha("EhAmarrado") & "") & """, "
            jsonArray = jsonArray & """alturaMaximaPalete"": """ & Replace(EscapeJSON(rsFicha("AlturaMaximaPalete")), ",", ".") & """, "
            jsonArray = jsonArray & """pecasPorUnidade"": """ & EscapeJSON(rsFicha("PecasPorUnidade") & "") & """, "
            jsonArray = jsonArray & """palete"": """ & EscapeJSON(rsFicha("Palete") & "") & """, "
            jsonArray = jsonArray & """temFacesOpostas"": """ & EscapeJSON(rsFicha("TemFacesOpostas") & "") & """, "
            jsonArray = jsonArray & """temCantinho"": """ & EscapeJSON(rsFicha("TemCantinho") & "") & """, "
            jsonArray = jsonArray & """temGradeMadeira"": """ & EscapeJSON(rsFicha("TemGradeMadeira") & "") & """, "
            jsonArray = jsonArray & """temFilmeStrech"": """ & EscapeJSON(rsFicha("TemFilmeStrech") & "") & """, "
            jsonArray = jsonArray & """quantidadeEspelhos"": """ & EscapeJSON(rsFicha("QuantidadeEspelhos") & "") & """, "
            jsonArray = jsonArray & """conteudoEmbalado"": """ & EscapeJSON(rsFicha("ConteudoEmbalado") & "") & """, "
            jsonArray = jsonArray & """larguraFitilho"": """ & Replace(EscapeJSON(rsFicha("LarguraFitilho")), ",", ".") & """, "
            jsonArray = jsonArray & """comprimentoFitilho"": """ & Replace(EscapeJSON(rsFicha("ComprimentoFitilho")), ",", ".") & """, "
            jsonArray = jsonArray & """quantidadeFitilhos"": """ & Replace(EscapeJSON(quantidadeFitilhos), ",", ".") & """, "
            jsonArray = jsonArray & """tipoFitilho"": """ & EscapeJSON(tipoFitilho & "") & """, "
            jsonArray = jsonArray & """testeForcaVinco1"": """ & Replace(EscapeJSON(rsFicha("TesteForcaVinco1")), ",", ".") & """, "
            jsonArray = jsonArray & """testeForcaVinco2"": """ & Replace(EscapeJSON(rsFicha("TesteForcaVinco2")), ",", ".") & """, "
            jsonArray = jsonArray & """testeForcaVinco3"": """ & Replace(EscapeJSON(rsFicha("TesteForcaVinco3")), ",", ".") & """, "
            jsonArray = jsonArray & """testePinInterno"": """ & Replace(EscapeJSON(rsFicha("TestePinInterno")), ",", ".") & """, "
            jsonArray = jsonArray & """testePinExterno"": """ & Replace(EscapeJSON(rsFicha("TestePinExterno")), ",", ".") & """, "
            jsonArray = jsonArray & """quantidadeCores"": """ & EscapeJSON(rsFicha("QuantidadeCores") & "") & """, "
            jsonArray = jsonArray & """unidadeFabricacao"": """ & EscapeJSON(rsFicha("UnidadeFabricacao") & "") & """, "
            jsonArray = jsonArray & """tipoLinhaProducaoCliente"": """ & EscapeJSON(rsFicha("TipoLinhaProducaoCliente") & "") & """, "
            jsonArray = jsonArray & """idVendedor"": """ & EscapeJSON(rsFicha("IdVendedor") & "") & """, "
            jsonArray = jsonArray & """nomeVendedor"": """ & EscapeJSON(rsFicha("NomeVendedor") & "") & """"
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsFicha.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsFicha.Close()
    SET rsFicha = NOTHING
    
    ListarFichasTecnicas = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarConteudosEmbalado()
    codPlanta = Request("filtroCodPlanta")

    If codPlanta = "" Then
        ListarConteudosEmbalado = "{""sucesso"": false, ""mensagem"": ""Planta não informada""}"
        Exit Function
    End If

    SQL = "SELECT COD AS Codigo, DSC as Descricao FROM SKALI WHERE ATV = 1 AND DEL = 0 ORDER BY COD"
    
    On Error Resume Next
    SET rsConteudos = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsConteudos) Then
        ListarConteudosEmbalado = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de conteúdo embalado: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsConteudos.EOF Then
        Do While Not rsConteudos.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """codigo"": """ & EscapeJSON(rsConteudos("Codigo") & "") & """, "
            jsonArray = jsonArray & """descricao"": """ & trim(EscapeJSON(rsConteudos("Descricao")) & "") & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsConteudos.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsConteudos.Close()
    SET rsConteudos = NOTHING
    
    ListarConteudosEmbalado = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarPaletes()
    Dim SQL, rsPalete, jsonArray, primeiroItem
    Dim filtroCodPlanta, stringConexao

    filtroCodPlanta = Request("filtroCodPlanta")

    If filtroCodPlanta = "" Then
        ListarPaletes = "{""sucesso"": false, ""mensagem"": ""Planta não informada""}"
        Exit Function
    End If

    Select Case filtroCodPlanta
        Case "00"
            stringConexao = "PIR-DADOSADV"
            SQL = "SELECT X5_CHAVE as Chave, X5_DESCRI as Descricao FROM SX5010 WHERE D_E_L_E_T_ = '' AND X5_TABELA = '95' AND X5_FILIAL = '00' ORDER BY X5_DESCRI"  
        Case "04"
            stringConexao = "PIR-DADOSADV"
            SQL = "SELECT X5_CHAVE as Chave, X5_DESCRI as Descricao FROM SX5010 WHERE D_E_L_E_T_ = '' AND X5_TABELA = '95' AND X5_FILIAL = '00' ORDER BY X5_DESCRI"  
        Case "05"
            stringConexao = "BEN-TRIMBOX"
            SQL = "SELECT IDPalete as Chave, RTRIM(TipoEstrado) + ' - ' + Descricao as Descricao FROM vw_ConsultaPaletesPelaInpaNet WHERE Desativado = 0 ORDER BY Descricao"  
        Case "06"
            stringConexao = "FOR-TRIMBOX"
            SQL = "SELECT IDPalete as Chave, RTRIM(Tipo) + ' - ' + Descricao as Descricao FROM vw_ConsultaPaletesPelaInpaNet WHERE Desativado = 0 ORDER BY Descricao"  
        Case "07"
            stringConexao = "SAQ-TRIMBOX"
            SQL = "SELECT IDPalete as Chave, RTRIM(TipoEstrado) + ' - ' + Descricao as Descricao FROM vw_ConsultaPaletesPelaInpaNet WHERE Desativado = 0 ORDER BY Descricao"  
        Case Else
            ListarPaletes = "{""sucesso"": false, ""mensagem"": ""Planta inválida""}"
    End Select

    On Error Resume Next
    SET rsPalete = Conecta(stringConexao, SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsPalete) Then
        ListarPaletes = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de estruturas de papelão: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsPalete.EOF Then
        Do While Not rsPalete.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """chave"": """ & trim(rsPalete("Chave")) & "" & """, "
            jsonArray = jsonArray & """descricao"": """ & trim(rsPalete("Descricao") & "") & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsPalete.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsPalete.Close()
    SET rsPalete = NOTHING
    
    ListarPaletes = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarTiposProdutoPrimario()
    codPlanta = Request("filtroCodPlanta")

    If codPlanta = "" Then
        listarTiposProdutoPrimario = "{""sucesso"": false, ""mensagem"": ""Planta não informada""}"
        Exit Function
    End If

    SQL = "SELECT COD AS Codigo, DSC as Descricao FROM SKTIP WHERE ATV = 1 AND DEL = 0 ORDER BY COD"

    On Error Resume Next
    SET rsTiposProduto = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsTiposProduto) Then
        ListarTiposProdutoPrimario = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de conteúdo embala: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsTiposProduto.EOF Then
        Do While Not rsTiposProduto.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """codigo"": """ & EscapeJSON(rsTiposProduto("Codigo") & "") & """, "
            jsonArray = jsonArray & """descricao"": """ & trim(EscapeJSON(rsTiposProduto("Descricao")) & "") & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsTiposProduto.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsTiposProduto.Close()
    SET rsTiposProduto = NOTHING
    
    ListarTiposProdutoPrimario = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarDesenvolvedores()
    Dim SQL, rsDesenvolvedores, jsonArray, primeiroItem
    Dim filtroCodPlanta

    filtroCodPlanta = Request("filtroCodPlanta")

    If filtroCodPlanta = "" Then
        ListarDesenvolvedores = "{""sucesso"": false, ""mensagem"": ""Planta não informada""}"
        Exit Function
    End If

    SQL = "SELECT CRM_COD AS Id, CRM_NOME AS Nome, ISNULL(CRM_NREDUZ,CRM_NOME) NomeReduzido, CRM_EMAILPRIMARIO as Email "
    SQL = SQL & "FROM CRM_DTI "
    SQL = SQL & "WHERE ISNULL(ValidaPreCadastroProduto,0) = 1 AND ISNULL(Bloqueado,0) = 0 AND ISNULL(CRM_D_E_L_E_T,'') = ''  "
    
    If filtroCodPlanta = "00" Or filtroCodPlanta = "04" then
        SQL = SQL & "AND EMP_CODIGO IN ('00','04') "
    Else
        SQL = SQL & "AND EMP_CODIGO = '" & filtroCodPlanta & "' "
    End If
    
    SQL = SQL & "ORDER BY CRM_NOME"

    On Error Resume Next
    SET rsDesenvolvedores = Conecta("PIR-HELPDESK", SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsDesenvolvedores) Then
        ListarDesenvolvedores = "{""sucesso"": false, ""mensagem"": ""Erro ao buscar lista de técnicos: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If
    On Error Goto 0

    primeiroItem = True
    jsonArray = "["

    If Not rsDesenvolvedores.EOF Then
        Do While Not rsDesenvolvedores.EOF
            If Not primeiroItem Then jsonArray = jsonArray & "," End If
            
            jsonArray = jsonArray & "{"
            jsonArray = jsonArray & """id"": """ & EscapeJSON(rsDesenvolvedores("Id") & "") & """, "
            jsonArray = jsonArray & """nome"": """ & EscapeJSON(rsDesenvolvedores("Nome") & "") & """, "
            jsonArray = jsonArray & """nomeReduzido"": """ & EscapeJSON(rsDesenvolvedores("NomeReduzido") & "") & """, "
            jsonArray = jsonArray & """email"": """ & EscapeJSON(rsDesenvolvedores("Email") & "") & """ "
            jsonArray = jsonArray & "}"
            
            primeiroItem = False
            rsDesenvolvedores.MoveNext
        Loop
    End If

    jsonArray = jsonArray & "]"

    rsDesenvolvedores.Close()
    SET rsDesenvolvedores = NOTHING
    
    ListarDesenvolvedores = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarStatusAndamento()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "E" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Amostra Enviada" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "A" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Amostra Aprovada" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "R" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Amostra Reprovada" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "P" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Análise do Projeto" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "G" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Envio de Especificação/Amostra" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "B" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Liberação PPCP/Produção" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "I" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Liberação Análise de Custo" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "M" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Aprovação Custos Gerência" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "N" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Desenvolvimento de Layout" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "O" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Corte de Amostras" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "Q" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Falta de Chapas" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "S" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Orçamento" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "L" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Aprovação de Layout" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "C" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Liberação Comercial - Gerência" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "D" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Liberação Comercial - Diretoria" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "F" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Cadastro de FT" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "H" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Padrão de Cor" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "J" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Confecção do Ferramental" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "K" & """, "
    jsonArray = jsonArray & """descricao"": """ & "FT Liberada" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "T" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Sem Retorno" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarStatusAndamento = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarMotivosReprovacaoAmostra()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "P" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Preço" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "D" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Dimensional em Desacordo" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "E" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Especificação Cliente Desatualizada" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "S" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Especificação de Primário Errado" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "U" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Descontinuidade do Projeto" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "K" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Sem Retorno Comercial" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarMotivosReprovacaoAmostra = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarMotivosRejeicao()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "E" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Emissão em Duplicidade" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "P" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Preço" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "D" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Descontinuidade do Projeto" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "R" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Sem Retorno Comercial" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "C" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Sem Retorno Cliente" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "T" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Proposta Descartada" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarMotivosRejeicao = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function

Function ListarProcessosAprovacao()
    jsonArray = "["

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "1" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Inovação" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "4" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Melhoria Continua - SK" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "5" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Melhoria Continua - Cliente" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "6" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Alteração Layout" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "7" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Alteração Dimensional" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "8" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Alteração Estrutura" & """"
    jsonArray = jsonArray & "}"
    jsonArray = jsonArray & ","

    jsonArray = jsonArray & "{"
    jsonArray = jsonArray & """codigo"": """ & "9" & """, "
    jsonArray = jsonArray & """descricao"": """ & "Nova Referência conforme especificado" & """"
    jsonArray = jsonArray & "}"

    jsonArray = jsonArray & "]"
    
    ListarProcessosAprovacao = "{""sucesso"": true, ""mensagem"": ""Listagem realizada com sucesso"", ""dados"": " & jsonArray & "}"
End Function



'----------------------------------------------
'                 UTILITÁRIOS
'----------------------------------------------

Sub EnviarEmailCriacao(idSolicitacao, codPlanta, codCliente, nomeCliente, lojaCliente)
    emailResponsavelComercial = ""

    SQL = "SELECT Endereco FROM FollowUpSapComercial WHERE Planta = '" & codPlanta & "' " 

    On Error Resume Next
    SET rsEndereco = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number = 0 AND IsObject(rsEndereco) Then
        emailResponsavelComercial = rsEndereco("Endereco")
        rsEndereco.Close()
        SET rsEndereco = NOTHING
    End If
    On Error Goto 0

    assunto = "Uma SAP foi aberta por " & usuarioLogado("nome") & "  com o número " & idSolicitacao & "."

    mensagem = ""
    mensagem = mensagem & "<!DOCTYPE html>"
    mensagem = mensagem & "<html>"
    mensagem = mensagem & "<head>"
    mensagem = mensagem & "<meta charset=""UTF-8"">"
    mensagem = mensagem & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">"
    mensagem = mensagem & "<style>"
    mensagem = mensagem & "body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }"
    mensagem = mensagem & ".content { margin-bottom: 25px; }"
    mensagem = mensagem & ".info-item { margin-bottom: 0px; }"
    mensagem = mensagem & ".footer { border-top: 1px solid #ddd; padding-top: 15px; margin-top: 30px; font-size: 12px; color: #666; }"
    mensagem = mensagem & "</style>"
    mensagem = mensagem & "</head>"
    mensagem = mensagem & "<body>"
    mensagem = mensagem & "<div class=""content"">"
    mensagem = mensagem & "<p>Olá,</p>"
    mensagem = mensagem & "<p>Uma SAP foi aberta por " & usuarioLogado("nome") & " com o ID " & idSolicitacao & ".</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Cliente:</strong> " & codCliente & " - " & nomeCliente & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Loja:</strong> " & codLoja & "</div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    'emailResponsavelComercial = "thiago.pereira@smurfitwestrock.com.br"
    Call EnviaEmail (emailResponsavelComercial, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
End Sub

Sub EnviarEmailAprovacaoComercial(solicitacao)  
    emailResponsavelTecnico = ""

    SQL = "SELECT Endereco FROM WorkFlowPreCadastroProduto WHERE Tipo = 1 AND Planta = '" & solicitacao("codPlanta") & "'"
    
    On Error Resume Next
    SET rsResponsavelTecnico = Conecta(ObterConexaoCorporativo(solicitacao("codPlanta")), SQL)
    
    If Err.Number = 0 AND IsObject(rsResponsavelTecnico) Then
        emailResponsavelTecnico = rsResponsavelTecnico("Endereco")
        rsResponsavelTecnico.Close()
        SET rsResponsavelTecnico = NOTHING
    End If
    On Error Goto 0

    assunto = "A SAP número " & solicitacao("id") & " foi aprovada pelo comercial"

    mensagem = ""
    mensagem = mensagem & "<!DOCTYPE html>"
    mensagem = mensagem & "<html>"
    mensagem = mensagem & "<head>"
    mensagem = mensagem & "<meta charset=""UTF-8"">"
    mensagem = mensagem & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">"
    mensagem = mensagem & "<style>"
    mensagem = mensagem & "body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }"
    mensagem = mensagem & ".content { margin-bottom: 25px; }"
    mensagem = mensagem & ".info-item { margin-bottom: 0px; }"
    mensagem = mensagem & ".footer { border-top: 1px solid #ddd; padding-top: 15px; margin-top: 30px; font-size: 12px; color: #666; }"
    mensagem = mensagem & "</style>"
    mensagem = mensagem & "</head>"
    mensagem = mensagem & "<body>"
    mensagem = mensagem & "<div class=""content"">"
    mensagem = mensagem & "<p>Olá,</p>"
    mensagem = mensagem & "<p>A SAP número <strong>" & solicitacao("id") & "</strong> foi aprovada pelo comercial.</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Cliente:</strong> " & solicitacao("codCliente") & " - " & solicitacao("nomeCliente") & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Loja:</strong> " & solicitacao("codLoja") & "</div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Nome Empresa</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    'emailResponsavelTecnico = "thiago.pereira@smurfitwestrock.com.br"
    Call EnviaEmail (emailResponsavelTecnico, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
End Sub

Sub EnviarEmailAprovacaoPCP(solicitacao)  
    emailResponsavelTecnico = ""

    SQL = "SELECT Endereco FROM WorkFlowPreCadastroProduto WHERE Tipo = 1 AND Planta = '" & solicitacao("codPlanta") & "'"
    
    On Error Resume Next
    SET rsResponsavelTecnico = Conecta(ObterConexaoCorporativo(solicitacao("codPlanta")), SQL)
    
    If Err.Number = 0 AND IsObject(rsResponsavelTecnico) Then
        emailResponsavelTecnico = rsResponsavelTecnico("Endereco")
        rsResponsavelTecnico.Close()
        SET rsResponsavelTecnico = NOTHING
    End If
    On Error Goto 0

    assunto = "A SAP número " & solicitacao("id") & " foi aprovada pelo PCP"

    mensagem = ""
    mensagem = mensagem & "<!DOCTYPE html>"
    mensagem = mensagem & "<html>"
    mensagem = mensagem & "<head>"
    mensagem = mensagem & "<meta charset=""UTF-8"">"
    mensagem = mensagem & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">"
    mensagem = mensagem & "<style>"
    mensagem = mensagem & "body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }"
    mensagem = mensagem & ".content { margin-bottom: 25px; }"
    mensagem = mensagem & ".info-item { margin-bottom: 0px; }"
    mensagem = mensagem & ".footer { border-top: 1px solid #ddd; padding-top: 15px; margin-top: 30px; font-size: 12px; color: #666; }"
    mensagem = mensagem & "</style>"
    mensagem = mensagem & "</head>"
    mensagem = mensagem & "<body>"
    mensagem = mensagem & "<div class=""content"">"
    mensagem = mensagem & "<p>Olá,</p>"
    mensagem = mensagem & "<p>A SAP número <strong>" & solicitacao("id") & "</strong> foi aprovada pelo PCP.</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Cliente:</strong> " & solicitacao("codCliente") & " - " & solicitacao("nomeCliente") & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Loja:</strong> " & solicitacao("codLoja") & "</div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Nome Empresa</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    'emailResponsavelTecnico = "thiago.pereira@smurfitwestrock.com.br"
    Call EnviaEmail (emailResponsavelTecnico, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
End Sub

Sub EnviarEmailRejeicaoPCP(solicitacao, motivoRejeicao)  
    emailResponsavelTecnico = ""

    SQL = "SELECT Endereco FROM WorkFlowPreCadastroProduto WHERE Tipo = 1 AND Planta = '" & solicitacao("codPlanta") & "'"
    
    On Error Resume Next
    SET rsResponsavelTecnico = Conecta(ObterConexaoCorporativo(solicitacao("codPlanta")), SQL)
    
    If Err.Number = 0 AND IsObject(rsResponsavelTecnico) Then
        emailResponsavelTecnico = rsResponsavelTecnico("Endereco")
        rsResponsavelTecnico.Close()
        SET rsResponsavelTecnico = NOTHING
    End If
    On Error Goto 0

    assunto = "A SAP número " & solicitacao("id") & " foi rejeitada pelo PCP"

    mensagem = ""
    mensagem = mensagem & "<!DOCTYPE html>"
    mensagem = mensagem & "<html>"
    mensagem = mensagem & "<head>"
    mensagem = mensagem & "<meta charset=""UTF-8"">"
    mensagem = mensagem & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">"
    mensagem = mensagem & "<style>"
    mensagem = mensagem & "body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }"
    mensagem = mensagem & ".content { margin-bottom: 25px; }"
    mensagem = mensagem & ".info-item { margin-bottom: 0px; }"
    mensagem = mensagem & ".footer { border-top: 1px solid #ddd; padding-top: 15px; margin-top: 30px; font-size: 12px; color: #666; }"
    mensagem = mensagem & "</style>"
    mensagem = mensagem & "</head>"
    mensagem = mensagem & "<body>"
    mensagem = mensagem & "<div class=""content"">"
    mensagem = mensagem & "<p>Olá,</p>"
    mensagem = mensagem & "<p>A SAP número <strong>" & solicitacao("id") & "</strong> foi rejeitada pelo PCP.</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Motivo:</strong> " & motivoRejeicao & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Cliente:</strong> " & solicitacao("codCliente") & " - " & solicitacao("nomeCliente") & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Loja:</strong> " & solicitacao("codLoja") & "</div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Nome Empresa</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    'emailResponsavelTecnico = "thiago.pereira@smurfitwestrock.com.br"
    Call EnviaEmail (emailResponsavelTecnico, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
End Sub

Sub EnviarEmailAprovacaoDesenvolvimento(solicitacao, numeroFT)
    emailUsuarioCriacao = ""
    emailResponsavelTecnico = ""

    SQL = "SELECT TOP 1 CRM_EMAILPRIMARIO FROM HELPDESK..CRM_DTI WHERE CRM_LOGIN = '" & solicitacao("loginUsuarioCriacao") & "' AND ISNULL(CRM_D_E_L_E_T,'') = '' "

    On Error Resume Next
    SET rsUsuarioCriacao = Conecta("PIR-HELPDESK", SQL)
    
    If Err.Number = 0 AND IsObject(rsUsuarioCriacao) Then
        emailUsuarioCriacao = rsUsuarioCriacao("CRM_EMAILPRIMARIO")
        rsUsuarioCriacao.Close()
        SET rsUsuarioCriacao = NOTHING
    End If
    On Error Goto 0

    SQL = "SELECT Endereco FROM WorkFlowPreCadastroProduto WHERE Tipo <> 5 AND Planta = '" & solicitacao("codPlanta") & "'"
    
    On Error Resume Next
    SET rsResponsavelTecnico = Conecta(ObterConexaoCorporativo(solicitacao("codPlanta")), SQL)
    
    If Err.Number = 0 AND IsObject(rsResponsavelTecnico) Then
        emailResponsavelTecnico = rsResponsavelTecnico("Endereco")
        rsResponsavelTecnico.Close()
        SET rsResponsavelTecnico = NOTHING
    End If
    On Error Goto 0

    assunto = "Um cadastro de produto foi aberto por " & usuarioLogado("nome") & " com o numero " & numeroFT & "."

    mensagem = ""
    mensagem = mensagem & "<!DOCTYPE html>"
    mensagem = mensagem & "<html>"
    mensagem = mensagem & "<head>"
    mensagem = mensagem & "<meta charset=""UTF-8"">"
    mensagem = mensagem & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">"
    mensagem = mensagem & "<style>"
    mensagem = mensagem & "body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }"
    mensagem = mensagem & ".content { margin-bottom: 25px; }"
    mensagem = mensagem & ".info-item { margin-bottom: 0px; }"
    mensagem = mensagem & ".footer { border-top: 1px solid #ddd; padding-top: 15px; margin-top: 30px; font-size: 12px; color: #666; }"
    mensagem = mensagem & "</style>"
    mensagem = mensagem & "</head>"
    mensagem = mensagem & "<body>"
    mensagem = mensagem & "<div class=""content"">"
    mensagem = mensagem & "<p>Olá,</p>"
    mensagem = mensagem & "<p>Uma Pre FT foi gerada por " & usuarioLogado("nome") & " com o código " & numeroFT & ".</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Pré Cadastro:</strong> " & solicitacao("id") & "</div>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    'emailUsuarioCriacao = "thiago.pereira@smurfitwestrock.com.br"
    'emailResponsavelTecnico = "thiago.pereira@smurfitwestrock.com.br"
    Call EnviaEmail (emailUsuarioCriacao, emailResponsavelTecnico, "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
End Sub

Sub EnviarEmailRejeicao(solicitacao, motivoRejeicao)
    emailUsuarioCriacao = ""

    SQL = "SELECT TOP 1 CRM_EMAILPRIMARIO FROM HELPDESK..CRM_DTI WHERE CRM_LOGIN = '" & solicitacao("loginUsuarioCriacao") & "' AND ISNULL(CRM_D_E_L_E_T,'') = '' "

    On Error Resume Next
    SET rsUsuarioCriacao = Conecta("PIR-HELPDESK", SQL)
    
    If Err.Number = 0 AND IsObject(rsUsuarioCriacao) Then
        emailUsuarioCriacao = rsUsuarioCriacao("CRM_EMAILPRIMARIO")
        rsUsuarioCriacao.Close()
        SET rsUsuarioCriacao = NOTHING
    End If
    On Error Goto 0

    assunto = "A SAP número " & solicitacao("id") & " foi reprovada"
    motivoRejeicao = ObterMotivoRejeicaoTraduzido(motivoRejeicao)

    mensagem = ""
    mensagem = mensagem & "<!DOCTYPE html>"
    mensagem = mensagem & "<html>"
    mensagem = mensagem & "<head>"
    mensagem = mensagem & "<meta charset=""UTF-8"">"
    mensagem = mensagem & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">"
    mensagem = mensagem & "<style>"
    mensagem = mensagem & "body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; }"
    mensagem = mensagem & ".content { margin-bottom: 25px; }"
    mensagem = mensagem & ".info-item { margin-bottom: 0px; }"
    mensagem = mensagem & ".footer { border-top: 1px solid #ddd; padding-top: 15px; margin-top: 30px; font-size: 12px; color: #666; }"
    mensagem = mensagem & "</style>"
    mensagem = mensagem & "</head>"
    mensagem = mensagem & "<body>"
    mensagem = mensagem & "<div class=""content"">"
    mensagem = mensagem & "<p>Olá,</p>"
    mensagem = mensagem & "<p>A SAP número <strong>" & solicitacao("id") & "</strong> foi reprovada.</p>"
    mensagem = mensagem & "<div class=""info-item""><strong>Motivo:</strong> " & motivoRejeicao & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Cliente:</strong> " & solicitacao("codCliente") & " - " & solicitacao("nomeCliente") & "</div>"
    mensagem = mensagem & "<div class=""info-item""><strong>Loja:</strong> " & solicitacao("codLoja") & "</div>"
    mensagem = mensagem & "<p>Essa SAP não foi enviada para desenvolvimento.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "<div class=""footer"">"
    mensagem = mensagem & "<p><strong>Smurfit Westrock</strong></p>"
    mensagem = mensagem & "<p>Email: noreply.brasil@smurfitkappa.com.br</p>"
    mensagem = mensagem & "<p>Este é um email automático, por favor não responda.</p>"
    mensagem = mensagem & "</div>"
    mensagem = mensagem & "</body>"
    mensagem = mensagem & "</html>"

    'emailUsuarioCriacao = "thiago.pereira@smurfitwestrock.com.br"
    Call EnviaEmail (emailUsuarioCriacao, "", "", "", "noreply.brasil@smurfitkappa.com.br", mensagem, assunto)
End Sub

Sub InserirAcessorios(idSolicitacao, acessoriosJSON, codPlanta)
    Dim SQL, rsAcessorios

    acessoriosJSON = Replace(acessoriosJSON, "[", "")
    acessoriosJSON = Replace(acessoriosJSON, "]", "")
    
    Dim acessoriosArray, acessorio, i
    acessoriosArray = Split(acessoriosJSON, "},{")
    
    Dim valoresInsert
    valoresInsert = ""
    For i = 0 To UBound(acessoriosArray)
        acessorio = acessoriosArray(i)
        
        Dim quantidade, caminho
        quantidade = ExtrairValorJSON(acessorio, "quantidade")
        estilo = ExtrairValorJSON(acessorio, "estilo")
        numeroCortes = ExtrairValorJSON(acessorio, "numeroCortes")
        largura = ExtrairValorJSON(acessorio, "largura")
        comprimento = ExtrairValorJSON(acessorio, "comprimento")
        estrutura = ExtrairValorJSON(acessorio, "estrutura")

        If i > 0 Then valoresInsert = valoresInsert & ","

        valoresInsert = valoresInsert & "("
        valoresInsert = valoresInsert & "'" & idSolicitacao & "', "
        valoresInsert = valoresInsert & "'" & i + 1 & "', "
        valoresInsert = valoresInsert & "'" & quantidade & "', "
        valoresInsert = valoresInsert & "'" & estilo & "', "
        valoresInsert = valoresInsert & "'" & numeroCortes & "', "
        valoresInsert = valoresInsert & "'" & largura & "', "
        valoresInsert = valoresInsert & "'" & comprimento & "', "
        valoresInsert = valoresInsert & "'" & estrutura & "'"
        valoresInsert = valoresInsert & ")"
    Next

    SQL = "BEGIN TRANSACTION TRANS; " 

    SQL = "DELETE FROM SKFT2 WHERE FT2_ID = '" & idSolicitacao & "'; "

    If valoresInsert <> "" Then
        SQL = SQL & "INSERT INTO SKFT2 ("
        SQL = SQL & "   FT2_ID, FT2_ITEM, FT2_ACEQTD, FT2_ACEFEFCO, "
        SQL = SQL & "   FT2_ACECORTE, FT2_ACELARG, FT2_ACECOMP, FT2_ACEESTR"
        SQL = SQL & ") VALUES " & valoresInsert
    End If

    SQL = SQL & "IF @@ERROR <> 0 "
    SQL = SQL & "   ROLLBACK TRANSACTION TRANS "
    SQL = SQL & "ELSE "
    SQL = SQL & "   COMMIT TRANSACTION TRANS "
    
    On Error Resume Next
    SET rsAcessorios = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number = 0 AND IsObject(rsAcessorios) Then
        rsAcessorios.Close()
        SET rsAcessorios = NOTHING
    End If
    On Error Goto 0
End Sub

Sub InserirAnexos(idSolicitacao, anexosJSON, codPlanta)
    Dim SQL, PIR_CORPORATIVO

    On Error Goto 0

    anexosJSON = Replace(anexosJSON, "[", "")
    anexosJSON = Replace(anexosJSON, "]", "")
    
    Dim anexosArray, anexo, i
    anexosArray = Split(anexosJSON, "},{")
    
    For i = 0 To UBound(anexosArray)
        anexo = anexosArray(i)
        
        Dim nome, caminho
        nome = ExtrairValorJSON(anexo, "name")
        caminho = ExtrairValorJSON(anexo, "uploadedName")
        
        SQL = "INSERT INTO SolicitacaoProjetoDocumento "
        SQL = SQL & "(IdSolicitacaoProjeto, Nome, Caminho, DataCriacao) VALUES ("
        SQL = SQL & idSolicitacao & ", "
        SQL = SQL & "'" & nome & "', "
        SQL = SQL & "'" & caminho & "', "
        SQL = SQL & "GETDATE()"
        SQL = SQL & ")"
        
        On Error Resume Next
        SET rsAnexos = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
        
        If Err.Number = 0 AND IsObject(rsAnexos) Then
            rsAnexos.Close()
            SET rsAnexos = NOTHING
        End If
        On Error Goto 0
    Next
End Sub


Function CancelarFTProtheus(numeroFTCancelamento, codPlanta)
    ' Tratar tipo de alteração = 3 (Altera/Cancela)'
    SQL = SQL & "UPDATE SB1010 SET "
    SQL = SQL & "B1_STATUS = 'E' " 
    SQL = SQL & "WHERE ISNULL(D_E_L_E_T_, '') = ''"
    SQL = SQL & "AND B1_COD = '" & numeroFTCancelamento & "'"

    On Error Resume Next
    SET rsCancelar = Conecta(ObterConexaoProtheus(codPlanta), SQL)

    If Err.Number <> 0 OR Not IsObject(rsCancelar) Then
        CancelarFTProtheus = False
        Exit Function
    End If
    
    rsCancelar.Close()
    SET rsCancelar = NOTHING
    On Error Goto 0

    CancelarFTProtheus = True
End Function

Function CancelarFTTrimbox(numeroFTCancelamento, codPlanta)
    ' Tratar tipo de alteração = 3 (Altera/Cancela)'
    SQL = SQL & "UPDATE Itens SET "
    SQL = SQL & "EstadoFT_Detec = '4', ObsVersao = 'Altera/Cancela SPD' " 
    SQL = SQL & "WHERE Conjunto = '" & numeroFTCancelamento & "'"


    On Error Resume Next
    SET rsCancelar = Conecta(ObterConexaoTrimbox(codPlanta), SQL)

    If Err.Number <> 0 OR Not IsObject(rsCancelar) Then
        CancelarFTTrimbox = False
        Exit Function
    End If
    
    rsCancelar.Close()
    SET rsCancelar = NOTHING
    On Error Goto 0

    CancelarFTTrimbox = True
End Function

Function InserirFTProtheus(solicitacao)
    numeroFT = GerarNumeroFTProtheus(solicitacao("codPlanta"))
    idSB1010 = GerarProximoIdSB1010(solicitacao("codPlanta"))
    codPcp = GerarCodPcp(solicitacao("codCliente"), solicitacao("codPlanta"))
    unidadeFabricacao = ObterUnidadeFabricacao(solicitacao("codPlanta"))
    
    ' Tabela de Produto
    SQL = "INSERT INTO SB1010 ("
    SQL = SQL & "   B1_FILIAL, B1_SUBGRUP, B1_GRUPO, B1_FAMGRUP, B1_COD, B1_DESC, B1_STATUS, B1_TIPO, "
    SQL = SQL & "   B1_ALIMENT, B1_TIPCONV, B1_CONTA, B1_NATUREZ, B1_DESCNF, B1_INPUBE, B1_CODPCP, "
    SQL = SQL & "   D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_, B1_UM, B1_LOCPAD, B1_QE, B1_SEGUM, B1_PRV1, B1_EMIN,"
    SQL = SQL & "   B1_CUSTD, B1_UPRC, B1_PESO, B1_ESTSEG, B1_PE, B1_LM, B1_TIPODEC, B1_ORIGEM, B1_DATREF, "
    SQL = SQL & "   B1_UREV, B1_RASTRO, B1_COMIS, B1_MRP,  B1_PERINV, B1_NOTAMIN, B1_PRVALID, B1_NUMCOP, B1_LOCALIZ,"
    SQL = SQL & "   B1_VLREFUS, B1_IMPORT, B1_NUMCQPR, B1_OPAUTOM, B1_ENVOBR, B1_MTBF, B1_MTTR, B1_PCSLL, B1_NROPAG, "
    SQL = SQL & "   B1_PESBRU, B1_POTENCI, B1_QTDACUM, B1_QTDINIC, B1_LOTVEN, B1_UMOEC, B1_UVLRC, B1_PMACNUT, B1_PMICNUT, "
    SQL = SQL & "   B1_PAUTFET, B1_PRFDSUL, B1_CRDEST, B1_FECP, B1_QBP, B1_ALFECOP, B1_ALFECST, B1_CRDPRES,"
    SQL = SQL & "   B1_AFETHAB, B1_AFACS, B1_AFABOV, B1_EMINUBE, B1_ALFUMAC, B1_MSGFSC, B1_PR43080, B1_FECPBA,"
    SQL = SQL & "   B1_DCRII, B1_COEFDCR, B1_PRINCMG, B1_MARKUP, B1_ALFECRN, B1_VLCIF, B1_LOTESBP, B1_GARANT, B1_PERGART,"
    SQL = SQL & "   B1_AFAMAD, B1_AFASEMT, B1_AIMAMT, B1_AFUNDES, B1_IMPNCM, B1_POSIPI, B1_MSBLQL "
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'', "
    SQL = SQL & "'710', "
    SQL = SQL & "'71', "
    SQL = SQL & "'7', "
    SQL = SQL & "'" & numeroFT & "', "
    SQL = SQL & "'" & solicitacao("referencia") & "', "
    SQL = SQL & "'P', "
    SQL = SQL & "'PA', "
    SQL = SQL & "'" & Left(solicitacao("conteudoEmbalado"), 1)  & "', " 
    SQL = SQL & "'" & solicitacao("tipoLinhaProducaoCliente") & "', "
    SQL = SQL & "'" & solicitacao("conta") & "', "
    SQL = SQL & "'" & solicitacao("natureza") & "', "
    SQL = SQL & "'" & solicitacao("descNF") & "', "
    SQL = SQL & "'" & unidadeFabricacao & "', "
    SQL = SQL & "'" & codPcp & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & idSB1010 & "', "
    SQL = SQL & "0, "
    SQL = SQL & "'UN', "
    SQL = SQL & "'01', "
    SQL = SQL & "1, "
    SQL = SQL & "'UN', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'N', "
    SQL = SQL & "0, "
    SQL = SQL & "CONVERT(VARCHAR(8),GETDATE(),112), "
    SQL = SQL & "CONVERT(VARCHAR(8),GETDATE(),112), "
    SQL = SQL & "'N', "
    SQL = SQL & "0, "
    SQL = SQL & "'S', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'N', "
    SQL = SQL & "0, "
    SQL = SQL & "'N', "
    SQL = SQL & "0, "
    SQL = SQL & "'N', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'B', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "2, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'48191000', "
    SQL = SQL & "'1' "
    SQL = SQL & "); "

    ' Tabela de detalhamento de produto
    idSB5010 = GerarProximoIdSB5010(solicitacao("codPlanta"))
    descricaoQuantidadeEspelhos = ""
    descricaoPalete = ObterDescricaoPalete(solicitacao("codPlanta"), solicitacao("palete"))
    b5Lap = ""
    b5Arrl = ""
    b5Arrc = ""
    
    If solicitacao("quantidadeEspelhos")  > 1 Then
        descricaoQuantidadeEspelhos = "Colocar " & solicitacao("quantidadeEspelhos") & " espelhos"
    End If

    if solicitacao("estilo") = "0200" Or _
        solicitacao("estilo") = "0201" Or _
        solicitacao("estilo") = "0202" Or _
        solicitacao("estilo") = "0203" Or _
        solicitacao("estilo") = "0205" Or _
        solicitacao("estilo") = "0209" Or _
        solicitacao("estilo") = "0214" Then
        b5Lap = "L"
    end if

    If solicitacao("descNF") = "901" Then
        b5Arrl = "1"
        b5Arrc = "1"
    End If

    SQL = SQL & "INSERT INTO SB5010 ("
    SQL = SQL & "   B5_FILIAL, B5_COD, B5_CLI, B5_LOJA, B5_NOMECLI, B5_REF, B5_TOLMA, "
    SQL = SQL & "   B5_TOLME, B5_COMPR, B5_LARG, B5_ALTU, B5_NUMFACA, B5_ESTRUTU,"
    SQL = SQL & "   B5_GRAM, B5_ESMAG, B5_MULLEN, B5_COLUNA, B5_ESPES2, B5_COMPRE, "
    SQL = SQL & "   B5_COBINT, B5_COBEXT, B5_FECHAM, B5_PALETE, B5_DESCPAL, B5_ESTILO, "
    SQL = SQL & "   B5_AMARRAD, B5_FITLAR, B5_FITCOM, B5_QTDPCT, B5_TPPALE, B5_FITASL, "
    SQL = SQL & "   B5_FACOP, B5_CANTON, B5_TAMPA, B5_FILME, B5_ESPELHO, D_E_L_E_T_, "
    SQL = SQL & "   R_E_C_N_O_, R_E_C_D_E_L_, B5_QE1, B5_PESOLIQ, B5_VLMATRI, B5_VLCLICH,"
    SQL = SQL & "   B5_PESOESP, B5_PRV2, B5_PRV3, B5_PRV4, B5_PRV5, B5_PRV6, B5_PRV7, B5_USUARIO,"
    SQL = SQL & "   B5_DTCADFT, B5_DTAPRFT, B5_CONVDIP, B5_REVPRO, B5_ALTURLC, B5_COMPRLC, "
    SQL = SQL & "   B5_EMPMAX, B5_FATARMA, B5_LARGLC, B5_QPA, B5_PRODDNF, B5_TMPFABR, "
    SQL = SQL & "   B5_VALFAT, B5_IDADEM, B5_FTIN, B5_FTIF, B5_PROPOR1, B5_PROPOR2, B5_CRDEST1,"
    SQL = SQL & "   B5_CRDEST2, B5_LVLEMB, B5_LPSEMB, B5_PERCUB, B5_INSUMO, B5_VOLRECI, B5_CODISRF, "
    SQL = SQL & "   B5_CONCENT, B5_DENSID, B5_CALORIA, B5_PESO, B5_TOTGOR, B5_TOTGORS, B5_TOTGORT, "
    SQL = SQL & "   B5_ALTURA, B5_EAN141, B5_EAN142, B5_EAN143, B5_EAN144, B5_EAN145, B5_EAN146, "
    SQL = SQL & "   B5_EAN147, B5_EAN148, B5_QEI, B5_QEL, B5_AMMULTO, B5_AMMULTS, B5_ZFMULTO,  "
    SQL = SQL & "   B5_ZFMULTS, B5_CODGNRE, B5_FCIPRV, B5_AM4PORC, B5_ZF4PORC, B5_VOL, "
    SQL = SQL & "   B5_NRB1FPI, B5_NRB2FPI, B5_NRB2FUB, B5_ECPESOE, B5_ECALTEM, B5_DIASHL, "
    SQL = SQL & "   B5_DIASHF, B5_MARGPRE, B5_VLRCID, B5_PRZCQ, B5_ECLARGE, B5_BLQINVD,"
    SQL = SQL & "   B5_DIASES, B5_ECCUBAG, B5_PROTOTI, B5_ECCOMP, B5_ECPROFU, B5_LEADTR, "
    SQL = SQL & "   B5_ECLARGU, B5_NUMCLI, B5_CEME, B5_UMIND, B5_VEND, B5_SAP, B5_NOMVEN, "
    SQL = SQL & "   B5_NUMNCON, B5_LAP, B5_TIPO, B5_PCSCON, B5_DTRECSA, B5_OBS3,"
    SQL = SQL & "   B5_CERT, B5_ESTMAT, B5_DES, B5_EMB1, B5_EMB2, B5_ARRL, B5_ARRC"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'',"
    SQL = SQL & "'" & numeroFT & "',"
    SQL = SQL & "'" & solicitacao("codCliente") & "',"
    SQL = SQL & "'" & solicitacao("codLoja") & "',"
    SQL = SQL & "'" & solicitacao("nomeCliente") & "',"
    SQL = SQL & "'" & solicitacao("referencia") & "',"
    SQL = SQL & "'" & solicitacao("toleranciaMaxima") & "',"
    SQL = SQL & "'" & solicitacao("toleranciaMinima") & "',"
    SQL = SQL & "'" & solicitacao("comprimento") & "',"
    SQL = SQL & "'" & solicitacao("largura") & "',"
    SQL = SQL & "'" & solicitacao("altura") & "',"
    SQL = SQL & "'" & solicitacao("numeroMatriz") & "',"
    SQL = SQL & "'" & solicitacao("estruturaPapelao") & "'," 
    SQL = SQL & "'" & solicitacao("gramaturaReferencial") & "',"
    SQL = SQL & "'" & solicitacao("esmagamento") & "',"
    SQL = SQL & "'" & solicitacao("mullenReferencial") & "',"
    SQL = SQL & "'" & solicitacao("ectColuna") & "',"
    SQL = SQL & "'" & solicitacao("espessura") & "',"
    SQL = SQL & "" & Iif(solicitacao("bctCompressao") = "", 0, solicitacao("bctCompressao")) & " * 9.81 ,"  
    SQL = SQL & "'" & solicitacao("cobbInterno") & "',"
    SQL = SQL & "'" & solicitacao("cobbExterno") & "',"
    SQL = SQL & "'" & solicitacao("tipoFechamento") & "',"
    SQL = SQL & "'" & solicitacao("ehPaletizado") & "',"
    SQL = SQL & "'" & descricaoPalete & "',"
    SQL = SQL & "'" & solicitacao("estilo") & "',"
    SQL = SQL & "'" & solicitacao("ehAmarrado") & "',"
    SQL = SQL & "'" & solicitacao("larguraFitilho") & "',"
    SQL = SQL & "'" & solicitacao("comprimentoFitilho") & "',"
    SQL = SQL & "'" & solicitacao("pecasPorUnidade") & "',"
    SQL = SQL & "'" & solicitacao("palete") & "',"
    SQL = SQL & "'" & Iif(solicitacao("palete") = "", 0, 2) & "',"
    SQL = SQL & "'" & solicitacao("temFacesOpostas") & "',"
    SQL = SQL & "'" & solicitacao("temCantinho") & "',"
    SQL = SQL & "'" & solicitacao("temGradeMadeira") & "',"
    SQL = SQL & "'" & solicitacao("temFilmeStrech") & "',"
    SQL = SQL & "'" & solicitacao("quantidadeEspelhos") & "',"
    SQL = SQL & "'',"  
    SQL = SQL & "'" & idSB5010 & "',"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "SUBSTRING(('" & usuarioLogado("nome") & "'), 1, 15), "
    SQL = SQL & "CONVERT(VARCHAR(8),GETDATE(),112), "
    SQL = SQL & "CONVERT(VARCHAR(8),GETDATE(),112), "
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "'F', "
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "0,"
    SQL = SQL & "RIGHT(" & numeroFT & ", 6)," 
    SQL = SQL & "'" & solicitacao("estilo") & "',"
    SQL = SQL & "1,"
    SQL = SQL & "'" & solicitacao("idVendedor") & "',"
    SQL = SQL & "'" & solicitacao("id") & "',"
    SQL = SQL & "SUBSTRING(('" & solicitacao("nomeVendedor") & "'), 1, 15), "
    SQL = SQL & "1,"
    SQL = SQL & "'" & b5Lap & "',"    
    SQL = SQL & "'C', "
    SQL = SQL & "1,"
    SQL = SQL & "'" & solicitacao("data") & "',"
    SQL = SQL & "'" & descricaoQuantidadeEspelhos & "', "
    SQL = SQL & "'" & solicitacao("testePinInterno") & "',"
    SQL = SQL & "'" & solicitacao("testePinExterno") & "',"
    SQL = SQL & "'" & solicitacao("testeForcaVinco1") & "',"
    SQL = SQL & "'" & solicitacao("testeForcaVinco2") & "',"
    SQL = SQL & "'" & solicitacao("testeForcaVinco3") & "', "
    SQL = SQL & "'" & b5Arrl & "',"
    SQL = SQL & "'" & b5Arrc & "'"
    SQL = SQL & "); "
    
    ' Bloqueia clicheria (Validar pra que serve)
    idSZ9010 = GerarProximoIdSZ9010(solicitacao("codPlanta"))
    SQL = SQL & "INSERT INTO SZ9010 ( "
    SQL = SQL & "   Z9_FILIAL, Z9_LOC, Z9_LOC1, Z9_COD, Z9_BLOQ, Z9_DTSOLIC, "
    SQL = SQL & "   Z9_REF, Z9_CLI, D_E_L_E_T_, R_E_C_N_O_ "
    SQL = SQL & ") VALUES ( "
    SQL = SQL & "'', "
    SQL = SQL & "'CLICHE', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & numeroFT & "', "
    SQL = SQL & "'S', "
    SQL = SQL & "CONVERT(VARCHAR(8),GETDATE(),112), "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & idSZ9010 & "'"
    SQL = SQL & ")"

    On Error Resume Next
    SET rsInserirFT = Conecta(ObterConexaoProtheus(solicitacao("codPlanta")), SQL)
    
    If Err.Number = 0 AND IsObject(rsInserirFT) Then
        rsInserirFT.Close()
        SET rsInserirFT = NOTHING
    End If
    On Error Goto 0

    InserirFTProtheus = numeroFT
End Function

Function InserirFTTrimbox(solicitacao)
    numeroFT = GerarNumeroFTTrimbox(solicitacao("codPlanta"))
    Set especificacoes = ObterEspecificacoesProdutoTrimbox(solicitacao)

    'Response.Write(" | numeroFT: " & numeroFT)
    'Response.Write(" | codCliente: " & solicitacao("codCliente"))
    'Response.Write(" | nomeCliente: " & solicitacao("nomeCliente"))
    'Response.Write(" | cidade: " & solicitacao("amostraEndereco"))
    'Response.Write(" | UF: " & solicitacao("amostraUF"))
    'Response.Write(" | nomeVendedor: " & solicitacao("nomeVendedor"))
    'Response.Write(" | idVendedor: " & solicitacao("idVendedor"))
    'Response.Write(" | tipoFechamento: " & solicitacao("tipoFechamento"))

    ' Alguns campos precisam ser tratados pois não podemos inserir com valor zerado
    ' Deve ser ou maior que 0 ou NULL
    ectColuna = Iif(solicitacao("ectColuna") = 0, "NULL", Replace(solicitacao("ectColuna") & "", ",", "."))
    gramaturaReferencial = Iif(solicitacao("gramaturaReferencial") = 0, "NULL", Replace(solicitacao("gramaturaReferencial") & "", ",", "."))
    mullenReferencial = Iif(solicitacao("mullenReferencial") = 0, "NULL", Replace(solicitacao("mullenReferencial") & "", ",", "."))
    cobbInterno = Iif(solicitacao("cobbInterno") = 0, "NULL", Replace(solicitacao("cobbInterno") & "", ",", "."))
    cobbExterno = Iif(solicitacao("cobbExterno") = 0, "NULL", Replace(solicitacao("cobbExterno") & "", ",", "."))
    espessura = Iif(solicitacao("espessura") = 0, "NULL", Replace(solicitacao("espessura") & "", ",", "."))
    umidadeRelativa = Iif(solicitacao("umidadeRelativa") = 0, "NULL", Replace(solicitacao("umidadeRelativa") & "", ",", "."))
    bctCompressao = Iif(solicitacao("bctCompressao") = 0, "NULL", Replace(solicitacao("bctCompressao") & "", ",", "."))

    ' Largura e comprimento devem ser no mínimo 1
    largura = Iif(solicitacao("largura") = 0, "1", solicitacao("largura")) 
    comprimento = Iif(solicitacao("comprimento") = 0, "1", solicitacao("comprimento")) 

    SQL = SQL & "INSERT INTO TSK_IN_FTs ("
    SQL = SQL & "   Estado, ErroNesteRegistro, CodCliente, Cliente, NomeCliente, Endereco, "
    SQL = SQL & "   Cidade, UF, NomeRepresentante, CodRepresentante, CodFechamento, "
    SQL = SQL & "   Maquina1, Maquina2, Maquina3, Maquina4, CodCorERP1, DescricaoCorERP1, "
    SQL = SQL & "   CodCorERP2, DescricaoCorERP2, CodCorERP3, DescricaoCorERP3, CodCorERP4, "
    SQL = SQL & "   DescricaoCorERP4, IDPalete, ComposicaoERP, DataCriacaoRegistro, Item, "
    SQL = SQL & "   Conjunto, Referencia, CodigoReferencia, CodigoEAN, TipoABNT, TipoABNTERP, "
    SQL = SQL & "   DescricaoABNT, PecasPorUnidade, UnidadesPorConjunto, PecasPorConjunto, ExigeLaudo, "
    SQL = SQL & "   Gramatura, MullenMinimo, CrushMinimo, ColunaMinimo, CobbIntMaximo, "
    SQL = SQL & "   CobbExtMaximo, CobbInt30Maximo, CobbExt30Maximo, WickingMaximo, PAT_Interno, "
    SQL = SQL & "   PAT_Externo, Espessura, Umidade, Compressao, Composicao, FormatoSimplex, Largura, "
    SQL = SQL & "   RefiloLargura, Comprimento, RefiloComprimento, MultLarg, MultComp, AreaBrutaPeca, "
    SQL = SQL & "   RefugoCliente, VincosLargNoRiscador, VincoLarg1, VincoLarg2, VincoLarg3, "
    SQL = SQL & "   VincoLarg4, VincoLarg5, VincoLarg6, VincoLarg7, VincoLarg8, VincoLarg9, VincoLarg10, "
    SQL = SQL & "   VincosCompNoRiscador, VincoComp1, VincoComp2, VincoComp3, VincoComp4, VincoComp5, "
    SQL = SQL & "   VincoComp6, VincoComp7, VincoComp8, VincoComp9, VincoComp10, TearTape1, TearTape2, "
    SQL = SQL & "   TearTape3, TearTape4, Lap, ProlongLap, LapNoComp, LapInterno, LapPrimVinco, Refilado, "
    SQL = SQL & "   Rejeito, RejeitoPecas, ResinaInterna, ResinaExterna, EndurecedorMiolo, Fechamento, "
    SQL = SQL & "   Grampos, Amarrado, PecasPorPacote, FitilhosLargPacote, FitilhosCompPacote, Paletizado, "
    SQL = SQL & "   PacotesLargura, PacotesComprimento, PacotesAltura, FitasLargPalete, FitasCompPalete, "
    SQL = SQL & "   FacesOpostas, Espelho, Cantoneiras, Filme, Faca, CodFI, Cor1, ConsumoCor1, ViscosidadeCor1, "
    SQL = SQL & "   Cor2, ConsumoCor2, ViscosidadeCor2, Cor3, ConsumoCor3, ViscosidadeCor3, Cor4, ConsumoCor4, "
    SQL = SQL & "   ViscosidadeCor4, Cor5, ConsumoCor5, ViscosidadeCor5, Cor6, ConsumoCor6, ViscosidadeCor6, "
    SQL = SQL & "   ImpComRegistro, DetalheCV, TolMais, TolMenos, TipoObs1, Obs1, TipoObs2, Obs2, "
    SQL = SQL & "   TipoObs3, Obs3, TipoObs4, Obs4, TipoObs5, Obs5, TipoObs6, Obs6, TipoObs7, Obs7, TipoObs8, "
    SQL = SQL & "   Obs8, LarguraInterna, ComprimentoInterno, AlturaInterna, LargPeca, CompPeca, "
    SQL = SQL & "   DimAdic1, DimAdic2, DimAdic3, DimAdic4, DimAdic5, DimAdic6, ContraOnda, "
    SQL = SQL & "   ImprimeLogotipo, ComissaoDeVendas, Usuario, PesoAdic, CustoUnitAdic, DataModif, "
    SQL = SQL & "   DataCriacao, CompPacote, LargPacote, AlturaPacote, CompPaleteFechado, "
    SQL = SQL & "   LargPaleteFechado, AlturaPaleteFechado, PorosidadeMinimo, PesoCaixa, Projeto, RefiloOnduladeira, "
    SQL = SQL & "   TipoCustoExtra, TravaComposicao, DescrTravaComposicao, PrecoNegociado, RefiloOndOrcam, "
    SQL = SQL & "   GrupoPaletizacao, PathFiguraDoLastro, PathFiguraDaFT, PressaoDeArqueamento, EmpilhamentoMaximo, "
    SQL = SQL & "   ReferenciaGrupoPaletizacao, FlagEstoqueOuTerceiros"
    SQL = SQL & ") VALUES ("
    SQL = SQL & "'1', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & solicitacao("codCliente") & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & solicitacao("nomeCliente") & "', "
    SQL = SQL & "'" & solicitacao("amostraEndereco") & "', "
    SQL = SQL & "'" & solicitacao("amostraMunicipio") & "', "
    SQL = SQL & "'" & solicitacao("amostraUF") & "', "
    SQL = SQL & "'" & solicitacao("nomeVendedor") & "', "
    SQL = SQL & "'" & solicitacao("idVendedor") & "', "
    SQL = SQL & "'" & solicitacao("tipoFechamento") & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & Iif(solicitacao("palete") = "", "NULL, ", "'" & solicitacao("palete") & "', ") 
    SQL = SQL & "'', "
    SQL = SQL & "GETDATE(), "
    SQL = SQL & "'" & numeroFT & "', "
    SQL = SQL & "'" & numeroFT & "', "
    SQL = SQL & "'" & solicitacao("referencia") & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & solicitacao("estilo") & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & especificacoes("pecasPorUnidade") & "', "
    SQL = SQL & "1, "
    SQL = SQL & "1, "
    SQL = SQL & "0, "
    SQL = SQL & "" & gramaturaReferencial & ", "
    SQL = SQL & "" & mullenReferencial & ", "
    SQL = SQL & "NULL, "
    SQL = SQL & "" & ectColuna & ", "
    SQL = SQL & "" & cobbInterno & ", "
    SQL = SQL & "" & cobbExterno & ", "
    SQL = SQL & "NULL, "
    SQL = SQL & "NULL, "
    SQL = SQL & "NULL, "
    SQL = SQL & "NULL, "
    SQL = SQL & "NULL, "
    SQL = SQL & "" & espessura & ", "
    SQL = SQL & "" & umidadeRelativa & ", "
    SQL = SQL & "" & bctCompressao & ", "
    SQL = SQL & "'" & solicitacao("estruturaPapelao") & "', "
    SQL = SQL & "0, "
    SQL = SQL & "" & largura & ", "
    SQL = SQL & "'" & especificacoes("refiloLargura") & "', "
    SQL = SQL & "" & comprimento & ", "
    SQL = SQL & "'" & especificacoes("refiloComprimento") & "', "
    SQL = SQL & "1, "
    SQL = SQL & "1, "
    SQL = SQL & "0, "
    SQL = SQL & "'" & especificacoes("refugoCliente") & "', "
    SQL = SQL & "'" & especificacoes("vincosLargNoRiscador") & "', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'" & especificacoes("vincosCompNoRiscador") & "', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'" & especificacoes("lap") & "', "
    SQL = SQL & "0, "
    SQL = SQL & "'" & especificacoes("lapNoComp") & "', "
    SQL = SQL & "'" & especificacoes("lapInterno") & "', "
    SQL = SQL & "'" & especificacoes("lapPrimVinco") & "', "
    SQL = SQL & "0, "
    SQL = SQL & "'" & especificacoes("rejeito") & "', "
    SQL = SQL & "'" & especificacoes("rejeitoPecas") & "', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "CASE '" & solicitacao("tipoFechamento") & "' WHEN 'S' then 'SEM FECHAMENTO' WHEN 'C' then 'Cola' WHEN 'G' then 'Grampo' ELSE '0' END, "   
    SQL = SQL & "0, "
    SQL = SQL & "CASE '" & solicitacao("ehAmarrado") & "' WHEN '' then 0 ELSE CASE '" & solicitacao("ehAmarrado") & "' WHEN 'S' then 1 ELSE 0 END END, "
    SQL = SQL & "CASE " & solicitacao("pecasPorUnidade") & " WHEN 0 then 25 ELSE " & solicitacao("pecasPorUnidade") & " END, "
    SQL = SQL & "'" & solicitacao("larguraFitilho") & "', "
    SQL = SQL & "'" & solicitacao("comprimentoFitilho") & "', "
    SQL = SQL & "CASE '" & solicitacao("ehPaletizado") & "' WHEN 'S' then 1 ELSE 0 END, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "CASE '" & solicitacao("temFacesOpostas") & "' WHEN 'S' then 1 ELSE 0 END, "
    SQL = SQL & "'" & solicitacao("quantidadeEspelhos") & "', "
    SQL = SQL & "CASE '" & solicitacao("temCantinho") & "' WHEN 'S' then 1 ELSE 0 END, "
    SQL = SQL & "CASE '" & solicitacao("temFilmeStrech") & "' WHEN 'S' then 1 ELSE 0 END, "
    SQL = SQL & "0, "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'', "
    SQL = SQL & "0, "
    SQL = SQL & "'', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'', "
    SQL = SQL & "'" & solicitacao("toleranciaMaxima") & "', "
    SQL = SQL & "'" & solicitacao("toleranciaMinima") & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'" & solicitacao("largura") & "', "
    SQL = SQL & "'" & solicitacao("comprimento") & "', "
    SQL = SQL & "'" & solicitacao("altura") & "', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "NULL, "
    SQL = SQL & "GETDATE(), "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'" & solicitacao("alturaMaximaPalete") & "', "
    SQL = SQL & "NULL, "
    SQL = SQL & "NULL, "
    SQL = SQL & "'', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "'', "
    SQL = SQL & "0, "
    SQL = SQL & "0, "
    SQL = SQL & "'', "
    SQL = SQL & "'" & especificacoes("flagEstoqueOuTerceiros") & "'"
    SQL = SQL & "); "

    SQL = SQL & "DECLARE @return_value_InterfaceImportacaoFTsPasso1 int, @erro_InterfaceImportacaoFTsPasso1 varchar(8000) "
    SQL = SQL & "EXEC	@return_value_InterfaceImportacaoFTsPasso1 = [ssp_InterfaceImportacaoFTsPasso1_SK] " & numeroFT & ", "
    SQL = SQL & "@erro = @erro_InterfaceImportacaoFTsPasso1 OUTPUT "

    ' Todo tratar erro retornado
    'SQL = SQL = "SET NOCOUNT ON; "
    'SQL = SQL & "DECLARE @retorno int, @erro varchar(8000) "
    'SQL = SQL & "EXEC @retorno = [ssp_InterfaceImportacaoFTsPasso1_SK] " & numeroFT & ", "
    'SQL = SQL & "@erro OUTPUT "
    'SQL = SQL & "SELECT @erro as erro, @retorno as retorno"
    
    'On Error Resume Next
    'SET rsInserirFT = Conecta(ObterConexaoTrimbox(solicitacao("codPlanta")), SQL)

    'If Err.Number = 0 AND IsObject(rsInserirFT) Then
        'Do Until rsInserirFT Is Nothing
            'If rsInserirFT.State = 1 Then 
                'Exit Do
            'End If
            'Set rsInserirFT = rsInserirFT.NextRecordset
        'Loop

        'If Not rsInserirFT Is Nothing Then
            'erro = rsInserirFT("erro")
            'retorno = rsInserirFT("retorno")
            'rsInserirFT.Close()
        'End If
    'End If
    'On Error Goto 0

    'Response.Write(" | Erro: " & erro & " | Retorno: " & retorno)
    'Response.Write(" | conexao: " & ObterConexaoTrimbox(solicitacao("codPlanta")))
    'Response.Write(" | SQL: " & SQL)

    On Error Resume Next
    SET rsInserirFT = Conecta(ObterConexaoTrimbox(solicitacao("codPlanta")), SQL)
    
    If Err.Number = 0 AND IsObject(rsInserirFT) Then
        rsInserirFT.Close()
        SET rsInserirFT = NOTHING
    End If
    On Error Goto 0

    InserirFTTrimbox = numeroFT
End Function

Function GerarProximoIdSolicitacao(codPlanta)
    proximoId = ""

    SQL = "SELECT ISNULL(MAX(FT1_ID), 300000) + 1 AS ProximoId FROM SKFT1"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta(ObterConexaoCorporativo(codPlanta), SQL)
    
    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        If Not PIR_CORPORATIVO.EOF Then
            proximoId = CStr(PIR_CORPORATIVO("ProximoId"))
        End If
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0

    GerarProximoIdSolicitacao = proximoId
End Function

Function GerarProximoIdSB1010(codPlanta)
    proximoId = ""

    SQL = "SELECT ISNULL(MAX(R_E_C_N_O_), 0) + 1 AS ProximoId FROM SB1010"

    On Error Resume Next
    SET rsProximoId = Conecta(ObterConexaoProtheus(codPlanta), SQL)
    
    If Err.Number = 0 AND IsObject(rsProximoId) Then
        If Not rsProximoId.EOF Then
            proximoId = CStr(rsProximoId("ProximoId"))
        End If
        
        rsProximoId.Close()
        SET rsProximoId = NOTHING
    End If
    On Error Goto 0

    GerarProximoIdSB1010 = proximoId
End Function

Function GerarProximoIdSB5010(codPlanta)
    proximoId = ""

    SQL = "SELECT ISNULL(MAX(R_E_C_N_O_), 0) + 1 AS ProximoId FROM SB5010"

    On Error Resume Next
    SET rsProximoId = Conecta(ObterConexaoProtheus(codPlanta), SQL)
    
    If Err.Number = 0 AND IsObject(rsProximoId) Then
        If Not rsProximoId.EOF Then
            proximoId = CStr(rsProximoId("ProximoId"))
        End If

        rsProximoId.Close()
        SET rsProximoId = NOTHING
    End If
    On Error Goto 0

    GerarProximoIdSB5010 = proximoId
End Function

Function GerarProximoIdSZ9010(codPlanta)
    proximoId = ""

    SQL = "SELECT ISNULL(MAX(R_E_C_N_O_), 0) + 1 AS ProximoId FROM SZ9010"

    On Error Resume Next
    SET rsProximoId = Conecta(ObterConexaoProtheus(codPlanta), SQL)
    
    If Err.Number = 0 AND IsObject(rsProximoId) Then
        If Not rsProximoId.EOF Then
            proximoId = CStr(rsProximoId("ProximoId"))
        End If

        rsProximoId.Close()
        SET rsProximoId = NOTHING
    End If
    On Error Goto 0

    GerarProximoIdSZ9010 = proximoId
End Function

Function GerarCodPcp(codCliente, codPlanta)
    cod = ""

    SQL = "DECLARE @codpcp varchar(12); "
    SQL = SQL & "SELECT @codpcp = ISNULL(REPLACE(MAX(B1_CODPCP),'" & codCliente & "',''), 0) + 1 FROM SB1010, SB5010 "
    SQL = SQL & "WHERE B1_COD = B5_COD "
    SQL = SQL & "AND B5_CLI = '" & codCliente & "' "
    SQL = SQL & "AND LEFT(B1_CODPCP,LEN(B5_CLI)) = '" & codCliente & "' "
    SQL = SQL & "AND SUBSTRING(B1_COD,1,3) NOT IN ('770'); "
    SQL = SQL & "Set @codpcp = '" & codCliente & "' + REPLICATE('0', 6-LEN(CONVERT(VARCHAR(6), @codpcp))) + CONVERT(VARCHAR(6), @codpcp); "
    SQL = SQL & "SELECT @codpcp as Cod; "

    On Error Resume Next
    SET rsCod = Conecta(ObterConexaoProtheus(codPlanta), SQL)
    
    If Err.Number = 0 AND IsObject(rsCod) Then
        If Not rsCod.EOF Then
            cod = CStr(rsCod("cod"))
        End If

        rsCod.Close()
        SET rsCod = NOTHING
    End If
    On Error Goto 0

    GerarCodPcp = cod
End Function

Function GerarNumeroFTProtheus(codPlanta)
    numeroFT = 0

    SQL = "DECLARE @numeroFT varchar(9); "

    SQL = SQL & "SELECT @numeroFT = MAX(SUBSTRING(B1_COD, 4, 9)) + 1 FROM SB1010 "
    SQL = SQL & "WHERE LEFT(B1_COD,3) = '710' AND LEN(B1_COD) = '9' AND B1_COD <> '710999998'; "
   
    SQL = SQL & "IF len(@numeroFT) = '5' "
    SQL = SQL & "    SET @numeroFT = '0' + @numeroFT; "
    SQL = SQL & "SET @numeroFT = '710' + @numeroFT; "
    SQL = SQL & "SELECT @numeroFT as numeroFT;"
  
    On Error Resume Next
    SET rsObterFT = Conecta(ObterConexaoProtheus(codPlanta), SQL)
    
    If Err.Number = 0 AND IsObject(rsObterFT) Then
        If Not rsObterFT.EOF Then
            numeroFT = CStr(rsObterFT("numeroFT"))
        End If

        rsObterFT.Close()
        SET rsObterFT = NOTHING
    End If
    On Error Goto 0

    GerarNumeroFTProtheus = numeroFT
End Function

Function GerarNumeroFTTrimbox(codPlanta)
    numeroFT = 0

    SQL = "SET NOCOUNT ON; "
    SQL = SQL & "DECLARE @return_value_ProxCodFTPelaInpanet int, @proximoProduto VARCHAR(20), @erro_ProxCodFTPelaInpanet VARCHAR(8000); "
    SQL = SQL & "EXEC @return_value_ProxCodFTPelaInpanet = pg_ProxCodProdutoPelaInpanet @proximoProduto OUTPUT, @erro_ProxCodFTPelaInpanet OUTPUT; "
    SQL = SQL & "SELECT @proximoProduto as numeroFT; "

    On Error Resume Next
    SET rsObterFT = Conecta(ObterConexaoTrimbox(codPlanta), SQL)

    If Err.Number = 0 AND IsObject(rsObterFT) Then
        If Not rsObterFT.EOF Then
            numeroFT = CStr(rsObterFT("numeroFT"))
        End If

        rsObterFT.Close()
        SET rsObterFT = NOTHING
    End If
    On Error Goto 0

    GerarNumeroFTTrimbox = numeroFT
End Function

Function ObterDadosContabilidade(codPlanta, estilo)
    Dim dados
    Set dados = Server.CreateObject("Scripting.Dictionary")
    
    dados("conta") = ""
    dados("descNF") = ""
    dados("natureza") = ""
    
    ' Pirapetinga e Uberaba
    If codPlanta = "00" Or codPlanta = "04" Then
        ' Grupo de estilos 901
        If InStr(",0200,0201,0202,0203,0205,0209,0214,", "," & estilo & ",") > 0 Then
            dados("conta") = "3111010001"
            dados("descNF") = "901"
            dados("natureza") = "2010103"
        ElseIf estilo = "0110" Then
            dados("conta") = "3111010003"
            dados("descNF") = "903"
            dados("natureza") = "2010102"
        Else
            dados("conta") = "3111010002"
            dados("descNF") = "902"
            dados("natureza") = "2010104"
        End If
    
    ' Bento 
    ElseIf codPlanta = "05" Then
        If estilo = "65" Then
            dados("conta") = "48191000"
            dados("descNF") = "901"
            dados("natureza") = ""
        ElseIf estilo = "14" Then
            dados("conta") = "48191000"
            dados("descNF") = "902"
            dados("natureza") = ""
        End If
    End If
    
    Set ObterDadosContabilidade = dados
End Function

Function ObterEspecificacoesProdutoTrimbox(solicitacao)
    Set dados = Server.CreateObject("Scripting.Dictionary")
    
    dados("refiloLargura") = ""
    dados("refiloComprimento") = ""
    dados("lapInterno") = ""
    dados("lapNoComp") = ""
    dados("lapPrimVinco") = ""
    dados("lap") = ""
    dados("pecasPorUnidade") = ""
    dados("rejeito") = ""
    dados("rejeitoPecas") = ""
    dados("vincosLargNoRiscador") = ""
    dados("vincosCompNoRiscador") = ""
    dados("flagEstoqueOuTerceiros") = ""
    dados("refugoCliente") = ""

    SQL = "DECLARE @RefiloLargura float, @RefiloComprimento float, "
    SQL = SQL & "@LapInterno int, @LapNoComp int, @LapPrimVinco int, @Lap int, "
    SQL = SQL & "@PecasPorUnidade int, @Rejeito int, @RejeitoPecas int, @VincosLargNoRiscador int, "
    SQL = SQL & "@VincosCompNoRiscador int, @FlagEstoqueOuTerceiros int, @RefugoCliente int; "

    SQL = SQL & "SELECT "
    SQL = SQL & "    @RefiloLargura = Isnull(RefiloLarg,0), @RefiloComprimento = Isnull(RefiloComp,0), "
    SQL = SQL & "    @LapInterno = Isnull(LapInterno,0), @LapNoComp = Isnull(LapNoComp,0), "
    SQL = SQL & "    @LapPrimVinco = Isnull(LapPrimVinco,0), @PecasPorUnidade = PecasPorUnidade, "
    SQL = SQL & "    @Rejeito = Rejeito, @RejeitoPecas = RejeitoPecas, @VincosLargNoRiscador = VincosLargNoRiscador, "
    SQL = SQL & "    @VincosCompNoRiscador = VincosCompNoRiscador, @FlagEstoqueOuTerceiros = FlagEstoqueOuTerceiros, "
    SQL = SQL & "    @RefugoCliente = RefugoCliente,"
    SQL = SQL & "    @Lap = CASE '" & solicitacao("tipoOnda") & "' WHEN 1 then Isnull(LapOndaSimples,0) WHEN 2 then Isnull(LapOndaDupla,0) else 0 End "  
    SQL = SQL & "From vw_ConsultaTipoCaixaPelaInpaNet where CodTrimbox = '" & solicitacao("estilo") & "'; "

    SQL = SQL & "SELECT "
    SQL = SQL & "   @RefiloLargura as RefiloLargura, "
    SQL = SQL & "   @RefiloComprimento as RefiloComprimento, "
    SQL = SQL & "   @LapInterno as LapInterno, "
    SQL = SQL & "   @LapNoComp as LapNoComp, "
    SQL = SQL & "   @LapPrimVinco as LapPrimVinco, "
    SQL = SQL & "   @Lap as Lap, "
    SQL = SQL & "   @PecasPorUnidade as PecasPorUnidade, "
    SQL = SQL & "   @Rejeito as Rejeito, "
    SQL = SQL & "   @RejeitoPecas as RejeitoPecas, "
    SQL = SQL & "   @VincosLargNoRiscador as VincosLargNoRiscador, "
    SQL = SQL & "   @VincosCompNoRiscador as VincosCompNoRiscador, "
    SQL = SQL & "   @FlagEstoqueOuTerceiros as FlagEstoqueOuTerceiros, "
    SQL = SQL & "   @RefugoCliente as RefugoCliente "

    On Error Resume Next
    SET rsEspecificacoes = Conecta(ObterConexaoTrimbox(solicitacao("codPlanta")), SQL)

    If Err.Number = 0 AND IsObject(rsEspecificacoes) Then
        If Not rsEspecificacoes.EOF Then
            dados("refiloLargura") = rsEspecificacoes("RefiloLargura")
            dados("refiloComprimento") = rsEspecificacoes("RefiloComprimento")
            dados("lapInterno") = rsEspecificacoes("LapInterno")
            dados("lapNoComp") = rsEspecificacoes("LapNoComp")
            dados("lapPrimVinco") = rsEspecificacoes("LapPrimVinco")
            dados("lap") = rsEspecificacoes("Lap")
            dados("pecasPorUnidade") = rsEspecificacoes("PecasPorUnidade")
            dados("rejeito") = rsEspecificacoes("Rejeito")
            dados("rejeitoPecas") = rsEspecificacoes("RejeitoPecas")
            dados("vincosLargNoRiscador") = rsEspecificacoes("VincosLargNoRiscador")
            dados("vincosCompNoRiscador") = rsEspecificacoes("VincosCompNoRiscador")
            dados("flagEstoqueOuTerceiros") = rsEspecificacoes("FlagEstoqueOuTerceiros")
            dados("refugoCliente") = rsEspecificacoes("RefugoCliente")
        End If

        rsEspecificacoes.Close()
        SET rsEspecificacoes = NOTHING
    End If
    On Error Goto 0

    Set ObterEspecificacoesProdutoTrimbox = dados
End Function

Function ObterDadosVendedor(codPlanta, codCliente, codLoja)
    Set dados = Server.CreateObject("Scripting.Dictionary")
    dados("idVendedor") = ""
    dados("nomeVendedor") = ""

    SQL = "SELECT A1_VEND AS IdVendedor, A3_NREDUZ as NomeVendedor FROM SA1010 "
    SQL = SQL & "INNER JOIN SA3010 ON A1_VEND = A3_COD "

    If codPlanta = "05" Or codPlanta = "06" Then
        'Bento e Maranguape
        SQL = SQL &  "WHERE A1_CDSIADN = '" & codCliente & "' " 
    Else
        ' Outras Plantas
        SQL = SQL & "WHERE A1_COD = '" & codCliente & "' "
        SQL = SQL & "AND A1_LOJA = '" & codLoja & "' "
    End If

    On Error Resume Next
    SET rsDadosVendedor = Conecta(ObterConexaoProtheus(codPlanta), SQL)

    If Err.Number <> 0 OR Not IsObject(rsDadosVendedor) Then
        Exit Function
    End If

    dados("idVendedor") = rsDadosVendedor("IdVendedor")
    dados("nomeVendedor") = rsDadosVendedor("NomeVendedor") 
    
    rsDadosVendedor.Close()
    SET rsDadosVendedor = NOTHING
    On Error Goto 0
    
    Set ObterDadosVendedor = dados
End Function

Function ObterFlagAprovacaoComercial(codPlanta)
    flagAprovacaoComercial = False

    SQL = "SELECT SPDAprovacaoComercial FROM PARAMETROS_PEDIDO WHERE Planta = '" & codPlanta & "' "

    On Error Resume Next
    SET rsFlag = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number <> 0 OR Not IsObject(rsFlag) Then
        ObterFlagAprovacaoComercial = flagAprovacaoComercial
        Exit Function
    End If

    If Not rsFlag.EOF Then
        flagAprovacaoComercial = Iif(rsFlag("SPDAprovacaoComercial"), True, False)
    End If
    
    rsFlag.Close()
    SET rsFlag = NOTHING
    On Error Goto 0
    
    ObterFlagAprovacaoComercial = flagAprovacaoComercial
End Function


Function ObterUnidadeFabricacao(codPlanta)
    unidadeFabricacao = ""

    Select Case codPlanta 
        Case "00" unidadeFabricacao = "I"
        Case "04" unidadeFabricacao = "U"
        Case "05" unidadeFabricacao = "B"
        Case "06" unidadeFabricacao = "F"
        Case "07" unidadeFabricacao = "S"
    end Select

    ObterUnidadeFabricacao = unidadeFabricacao
End Function


Function ObterDescricaoPalete(codPlanta, palete)
    descricaoPalete = ""

    If palete = "" Then
        Exit Function
    End If

    SQL = "SELECT X5_DESCRI FROM SX5010 "
	SQL = SQL & "WHERE D_E_L_E_T_='' "
	SQL = SQL & "AND X5_TABELA = '95' "
    SQL = SQL & "AND X5_FILIAL = '00' "
    SQL = SQL & "AND X5_CHAVE = '" & palete & "' "
    SQL = SQL & "collate SQL_Latin1_General_CP437_BIN"

    SET rsListar = Conecta(ObterConexaoProtheus(codPlanta), SQL)
    
    If Err.Number <> 0 OR Not IsObject(rsListar) Then
        descricaoPalete = ""
        Exit Function
    End If

    descricaoPalete = Iif(isNull(rsListar("X5_DESCRI")), "", rsListar("X5_DESCRI"))
    
    rsListar.Close()
    SET rsListar = NOTHING

    ObterDescricaoPalete = descricaoPalete
End Function

Function ObterConexaoCorporativo(codPlanta)
    Select Case codPlanta
        Case "00": ObterConexaoCorporativo = "PIR-CORPORATIVO"
        Case "04": ObterConexaoCorporativo = "PIR-CORPORATIVO"
        Case "05": ObterConexaoCorporativo = "BEN-CORPORATIVO"
        Case "06": ObterConexaoCorporativo = "FOR-CORPORATIVO"
        Case "07": ObterConexaoCorporativo = "SAQ-CORPORATIVO"
        Case Else: ObterConexaoCorporativo = ""
    End Select
End Function

Function ObterConexaoProtheus(codPlanta)
    Select Case codPlanta
        Case "00": ObterConexaoProtheus = "PIR-DADOSADV"
        Case "04": ObterConexaoProtheus = "PIR-DADOSADV"
        Case "05": ObterConexaoProtheus = "BEN-DADOSADV"
        Case "06": ObterConexaoProtheus = "FOR-DADOSADV"
        Case "07": ObterConexaoProtheus = "SAQ-DADOSADV"
        Case Else: ObterConexaoProtheus = ""
    End Select
End Function

Function ObterConexaoTrimbox(codPlanta)
    Select Case codPlanta
        Case "00": ObterConexaoTrimbox = "PIR-TRIMBOX"
        Case "04": ObterConexaoTrimbox = "UBE-TRIMBOX"
        Case "05": ObterConexaoTrimbox = "BEN-TRIMBOX"
        Case "06": ObterConexaoTrimbox = "FOR-TRIMBOX"
        Case "07": ObterConexaoTrimbox = "SAQ-TRIMBOX"
        Case Else: ObterConexaoTrimbox = ""
    End Select
End Function

Function ObterMotivoRejeicaoTraduzido(codMotivo) 
    motivo = ""

    Select Case UCase(codMotivo)
        Case "E"
            motivo = "Emissão em Duplicidade"
        Case "P"
            motivo = "Preço"
        Case "D"
            motivo = "Descontinuidade do Projeto"
        Case "R"
            motivo = "Sem Retorno Comercial"
        Case "C"
            motivo = "Sem Retorno Cliente"
        Case "T"
            motivo = "Proposta Descartada"
        Case Else
            motivo = ""
    End Select
    
    ObterMotivoRejeicaoTraduzido = motivo
End Function


Sub RegistrarLog(acao, id, statusAntigo, statusNovo)
    Select Case LCase(acao)
        Case "criacao"
            acaoLog = "Criação de SPD"
            descricaoLog = "Usuário " & usuarioLogado("login") & " criou a SPD de id " & id
        Case "edicao"
            acaoLog = "Edição de SPD"
            descricaoLog = "Usuário " & usuarioLogado("login") & " editou a SPD de id " & id
        Case "aprovar"
            acaoLog = "Aprovação de SPD"
            descricaoLog = "Usuário " & usuarioLogado("login") & " aprovou a SPD de id " & id
        Case "rejeitar"
            acaoLog = "Rejeição de SPD"
            descricaoLog = "Usuário " & usuarioLogado("login") & " rejeitou a SPD de id " & id
        Case "cancelar"
            acaoLog = "Cancelamento de SPD"
            descricaoLog = "Usuário " & usuarioLogado("login") & " cancelou a SPD de id " & id
        Case "aprovar_comercial"
            acaoLog = "Aprovação de SPD pelo comercial"
            descricaoLog = "Usuário " & usuarioLogado("login") & " aprovou a SPD de id " & id
        Case "rejeitar_comercial"
            acaoLog = "Rejeição de SPD pelo comercial"
            descricaoLog = "Usuário " & usuarioLogado("login") & " rejeitou a SPD de id " & id
        Case "alteracao_status"
            acaoLog = "Alteração de status de SPD"
            descricaoLog = "Usuário " & usuarioLogado("login") & " alterou o status da SPD de id " & id & " de " & statusAntigo & " para " & statusNovo
         Case "reversao_cancelamento"
            acaoLog = "Reversão de cancelamento de SPD"
            descricaoLog = "Usuário " & usuarioLogado("login") & " reverteu o cancelamento da SPD de id " & id
        Case "edicao_testes_fisicos"
            acaoLog = "Edição de testes físicos de SPD"
            descricaoLog = "Usuário " & usuarioLogado("login") & " editou os testes físicos da SPD de id " & id
        Case "exclusao"
            acaoLog = "Exclusão de SPD"
            descricaoLog = "Usuário " & usuarioLogado("login") & " excluiu a SPD de id " & id
    End Select
    
    descricaoLog = Replace(descricaoLog, "'", "''")
    acaoLog = Replace(acaoLog, "'", "''")
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & usuarioLogado("id") & ", "
    SQL = SQL & "'SPD', "
    SQL = SQL & "'SWNET', "
    SQL = SQL & "'" & acaoLog & "', "
    SQL = SQL & "'" & descricaoLog & "', "
    SQL = SQL & "GETDATE())"
    
    On Error Resume Next
    SET rsLog = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(rsLog) Then
        rsLog.Close()
        SET rsLog = NOTHING
    End If
    On Error Goto 0
End Sub
%>