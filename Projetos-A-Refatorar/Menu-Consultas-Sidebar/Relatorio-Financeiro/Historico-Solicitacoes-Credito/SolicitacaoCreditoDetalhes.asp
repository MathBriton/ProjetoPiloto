<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="./util/Util.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0

Dim idSolicitacao
idSolicitacao = Request("idSolicitacao")

If idSolicitacao = "" Then
    Response.Write "<div class='alert alert-warning'>ID da solicitação não informado.</div>"
    Response.End
End If

Dim solicitacao
Set solicitacao = Server.CreateObject("Scripting.Dictionary")

SQL = "SELECT "
SQL = SQL & "s.NUM as IdSolicitacao, "
SQL = SQL & "s.LOGININPANET as LoginSolicitante, "
SQL = SQL & "s.SOLICITANTE as NomeSolicitante, "
SQL = SQL & "s.NomeCliente, "
SQL = SQL & "s.QUEMVAL as NomeValidador, "
SQL = SQL & "s.EnderecoDocumentacao, "
SQL = SQL & "s.COD as CodCliente, "
SQL = SQL & "s.CODLOJA as CodLoja, "
SQL = SQL & "s.Planta as CodPlanta, "
SQL = SQL & "s.LCATUAL as LimiteAtual, "
SQL = SQL & "s.LCSOLICIT as LimiteSolicitado, "
SQL = SQL & "s.LCNOVO as LimiteNovo, "
SQL = SQL & "s.LCVALID as LimiteValidado, "
SQL = SQL & "s.FaturamentoEstimadoPositivo, "
SQL = SQL & "s.ScoreSerasa, "
SQL = SQL & "s.LimiteSugeridoSerasa, "
SQL = SQL & "s.ValorSeguradora, "
SQL = SQL & "s.VolumeM2, "
SQL = SQL & "s.VolumeUnidades, "
SQL = SQL & "s.VolumeToneladas, "
SQL = SQL & "s.PrecoM2, "
SQL = SQL & "s.PrecoKg, "
SQL = SQL & "s.ValorToneladaTora, "
SQL = SQL & "s.ValorUnitarioSemente, "
SQL = SQL & "s.ValorUnitarioMuda, "
SQL = SQL & "s.CondicaoPagamento, "
SQL = SQL & "s.STATUS as Status, "
SQL = SQL & "s.Observacao, "
SQL = SQL & "s.LimiteSugeridoSW, "
SQL = SQL & "s.ScoreSW, "
SQL = SQL & "s.ClasseRiscoSW, "
SQL = SQL & "s.ClasseRiscoSerasa, "
SQL = SQL & "s.BalancoAno, "
SQL = SQL & "s.ClienteGrupo, "
SQL = SQL & "s.TipoMercado, "
SQL = SQL & "s.CNPJCPF as CNPJ, "
SQL = SQL & "CASE WHEN ISDATE(s.DTVALID) = 1 THEN CONVERT(VARCHAR(10), CONVERT(DATE, s.DTVALID, 112), 126) ELSE NULL END AS DataValidacao, "
SQL = SQL & "CASE WHEN ISDATE(s.DTSOLICIT) = 1 THEN CONVERT(VARCHAR(10), CONVERT(DATE, s.DTSOLICIT, 112), 126) ELSE NULL END AS DataCadastro, "
SQL = SQL & "s.AprovadorPrimeiroNivel as ValidadorFinanceiro, "
SQL = SQL & "s.AprovadorSegundoNivel as ValidadorComercial, "
SQL = SQL & "s.AprovadorTerceiroNivel as ValidadorControladoria, "
SQL = SQL & "s.AprovadorQuartoNivel as ValidadorGerenciaGeral, "
SQL = SQL & "s.AprovadorQuintoNivel as ValidadorDiretoriaComercial, "
SQL = SQL & "s.AprovadorSextoNivel as ValidadorCFO, "
SQL = SQL & "s.AprovadorSetimoNivel as ValidadorCEO, "
SQL = SQL & "s.Observacao1 as ObservacaoFinanceiro, "
SQL = SQL & "s.Observacao2 as ObservacaoComercial, "
SQL = SQL & "s.Observacao3 as ObservacaoControladoria, "
SQL = SQL & "s.Observacao4 as ObservacaoGerenciaGeral, "
SQL = SQL & "s.Observacao5 as ObservacaoDiretoriaComercial, "
SQL = SQL & "s.Observacao6 as ObservacaoCFO, "
SQL = SQL & "s.Observacao7 as ObservacaoCEO, "
SQL = SQL & "CASE WHEN ISDATE(s.DataAprovacaoPrimeiroNivel) = 1 THEN CONVERT(VARCHAR(10), s.DataAprovacaoPrimeiroNivel, 126) ELSE NULL END as DataValidacaoFinanceiro, "
SQL = SQL & "CASE WHEN ISDATE(s.DataAprovacaoSegundoNivel) = 1 THEN CONVERT(VARCHAR(10), s.DataAprovacaoSegundoNivel, 126) ELSE NULL END as DataValidacaoComercial, "
SQL = SQL & "CASE WHEN ISDATE(s.DataAprovacaoTerceiroNivel) = 1 THEN CONVERT(VARCHAR(10), s.DataAprovacaoTerceiroNivel, 126) ELSE NULL END as DataValidacaoControladoria, "
SQL = SQL & "CASE WHEN ISDATE(s.DataAprovacaoQuartoNivel) = 1 THEN CONVERT(VARCHAR(10), s.DataAprovacaoQuartoNivel, 126) ELSE NULL END as DataValidacaoGerenciaGeral, "
SQL = SQL & "CASE WHEN ISDATE(s.DataAprovacaoQuintoNivel) = 1 THEN CONVERT(VARCHAR(10), s.DataAprovacaoQuintoNivel, 126) ELSE NULL END as DataValidacaoDiretoriaComercial, "
SQL = SQL & "CASE WHEN ISDATE(s.DataAprovacaoSextoNivel) = 1 THEN CONVERT(VARCHAR(10), s.DataAprovacaoSextoNivel, 126) ELSE NULL END as DataValidacaoCFO, "
SQL = SQL & "CASE WHEN ISDATE(s.DataAprovacaoSetimoNivel) = 1 THEN CONVERT(VARCHAR(10), s.DataAprovacaoSetimoNivel, 126) ELSE NULL END as DataValidacaoCEO "
SQL = SQL & "FROM SOLCREDITOCLIENTE s "
SQL = SQL & "WHERE s.NUM = " & idSolicitacao

SET RS = Conecta("PIR-CORPORATIVO", SQL)

If NOT RS.EOF Then
    solicitacao("IdSolicitacao") = RS("IdSolicitacao")
    solicitacao("LoginSolicitante") = TrataNull(RS("LoginSolicitante"))
    solicitacao("NomeSolicitante") = TrataNull(RS("NomeSolicitante"))
    solicitacao("NomeCliente") = TrataNull(RS("NomeCliente"))
    solicitacao("NomeValidador") = TrataNull(RS("NomeValidador"))
    solicitacao("CodCliente") = TrataNull(RS("CodCliente"))
    solicitacao("CodLoja") = TrataNull(RS("CodLoja"))
    solicitacao("CodPlanta") = TrataNull(RS("CodPlanta"))
    solicitacao("CNPJ") = TrataNull(RS("CNPJ"))
    solicitacao("LimiteAtual") = RS("LimiteAtual")
    solicitacao("LimiteSolicitado") = RS("LimiteSolicitado")
    solicitacao("LimiteNovo") = RS("LimiteNovo")
    solicitacao("LimiteValidado") = RS("LimiteValidado")
    solicitacao("FaturamentoEstimadoPositivo") = TrataNull(RS("FaturamentoEstimadoPositivo"))
    solicitacao("ScoreSerasa") = TrataNull(RS("ScoreSerasa"))
    solicitacao("LimiteSugeridoSerasa") = RS("LimiteSugeridoSerasa")
    solicitacao("ValorSeguradora") = RS("ValorSeguradora")
    solicitacao("VolumeM2") = TrataNull(RS("VolumeM2"))
    solicitacao("VolumeUnidades") = TrataNull(RS("VolumeUnidades"))
    solicitacao("VolumeToneladas") = TrataNull(RS("VolumeToneladas"))
    solicitacao("PrecoM2") = RS("PrecoM2")
    solicitacao("PrecoKg") = RS("PrecoKg")
    solicitacao("ValorToneladaTora") = RS("ValorToneladaTora")
    solicitacao("ValorUnitarioSemente") = RS("ValorUnitarioSemente")
    solicitacao("ValorUnitarioMuda") = RS("ValorUnitarioMuda")
    solicitacao("CondicaoPagamento") = TrataNull(RS("CondicaoPagamento"))
    solicitacao("Status") = TrataNull(RS("Status"))
    solicitacao("Observacao") = TrataNull(RS("Observacao"))
    solicitacao("LimiteSugeridoSW") = RS("LimiteSugeridoSW")
    solicitacao("ScoreSW") = TrataNull(RS("ScoreSW"))
    solicitacao("ClasseRiscoSW") = TrataNull(RS("ClasseRiscoSW"))
    solicitacao("ClasseRiscoSerasa") = TrataNull(RS("ClasseRiscoSerasa"))
    solicitacao("BalancoAno") = TrataNull(RS("BalancoAno"))
    solicitacao("ClienteGrupo") = TrataNull(RS("ClienteGrupo"))
    solicitacao("TipoMercado") = TrataNull(RS("TipoMercado"))
    solicitacao("DataValidacao") = TrataNull(RS("DataValidacao"))
    solicitacao("DataCadastro") = TrataNull(RS("DataCadastro"))
    solicitacao("ValidadorFinanceiro") = TrataNull(RS("ValidadorFinanceiro"))
    solicitacao("ValidadorComercial") = TrataNull(RS("ValidadorComercial"))
    solicitacao("ValidadorControladoria") = TrataNull(RS("ValidadorControladoria"))
    solicitacao("ValidadorGerenciaGeral") = TrataNull(RS("ValidadorGerenciaGeral"))
    solicitacao("ValidadorDiretoriaComercial") = TrataNull(RS("ValidadorDiretoriaComercial"))
    solicitacao("ValidadorCFO") = TrataNull(RS("ValidadorCFO"))
    solicitacao("ValidadorCEO") = TrataNull(RS("ValidadorCEO"))
    solicitacao("ObservacaoFinanceiro") = TrataNull(RS("ObservacaoFinanceiro"))
    solicitacao("ObservacaoComercial") = TrataNull(RS("ObservacaoComercial"))
    solicitacao("ObservacaoControladoria") = TrataNull(RS("ObservacaoControladoria"))
    solicitacao("ObservacaoGerenciaGeral") = TrataNull(RS("ObservacaoGerenciaGeral"))
    solicitacao("ObservacaoDiretoriaComercial") = TrataNull(RS("ObservacaoDiretoriaComercial"))
    solicitacao("ObservacaoCFO") = TrataNull(RS("ObservacaoCFO"))
    solicitacao("ObservacaoCEO") = TrataNull(RS("ObservacaoCEO"))
    solicitacao("DataValidacaoFinanceiro") = TrataNull(RS("DataValidacaoFinanceiro"))
    solicitacao("DataValidacaoComercial") = TrataNull(RS("DataValidacaoComercial"))
    solicitacao("DataValidacaoControladoria") = TrataNull(RS("DataValidacaoControladoria"))
    solicitacao("DataValidacaoGerenciaGeral") = TrataNull(RS("DataValidacaoGerenciaGeral"))
    solicitacao("DataValidacaoDiretoriaComercial") = TrataNull(RS("DataValidacaoDiretoriaComercial"))
    solicitacao("DataValidacaoCFO") = TrataNull(RS("DataValidacaoCFO"))
    solicitacao("DataValidacaoCEO") = TrataNull(RS("DataValidacaoCEO"))
Else
    Response.Write "<div class='alert alert-warning'>Solicitação não encontrada.</div>"
    Response.End
End If

If isObject(RS) Then
    RS.Close()
    SET RS = NOTHING
End If

Function TrataNull(valor)
    If IsNull(valor) Then
        TrataNull = ""
    Else
        TrataNull = valor
    End If
End Function

Function ExibirValor(valor)
    If valor = "" Or IsNull(valor) Then
        ExibirValor = "-"
    Else
        ExibirValor = valor
    End If
End Function
%>

<script type="text/javascript">
    function exportarDetalhesPDF() {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('portrait', 'mm', 'a4');
        const larguraPagina = doc.internal.pageSize.getWidth();
        const alturaPagina = doc.internal.pageSize.getHeight();
        const margin = 5;
        let posicaoY = 15;
        
        const idSolicitacao = '<%=solicitacao("IdSolicitacao")%>';
        
        function checarNovaPagina(precisaEspaco) {
            if (posicaoY + precisaEspaco > alturaPagina - 15) {
                doc.addPage();
                posicaoY = margin;
                return true;
            }
            return false;
        }
        
        function adicionarTitulo(title) {
            checarNovaPagina(12);
            doc.setFontSize(10);
            doc.setFont('helvetica', 'bold');
            doc.setTextColor(73, 80, 87);
            doc.text(title, margin, posicaoY);
            posicaoY += 2;

            doc.setLineWidth(0.1);
            doc.line(margin, posicaoY, larguraPagina - margin, posicaoY);
            posicaoY += 6;
        }
        
        function adicionarLinha(label, value) {
            checarNovaPagina(8);
            doc.setFontSize(9);
            doc.setFont('helvetica', 'normal');
            doc.setTextColor(108, 117, 125);
            doc.text(label, margin, posicaoY);
            doc.setTextColor(33, 37, 41);
            doc.text(String(value || '-'), margin + 40, posicaoY);
            posicaoY += 6;
        }
        
        function adicionarLinhaDuasColunas(label1, value1, label2, value2) {
            checarNovaPagina(8);
            const metadeLargura = larguraPagina / 2;
            
            doc.setFontSize(9);
            doc.setFont('helvetica', 'normal');
            doc.setTextColor(108, 117, 125);
            doc.text(label1, margin, posicaoY);
            doc.setTextColor(33, 37, 41);
            doc.text(String(value1 || '-'), margin + 50, posicaoY);
            
            doc.setTextColor(108, 117, 125);
            doc.text(label2, margin + metadeLargura + 3, posicaoY);
            doc.setTextColor(33, 37, 41);
            doc.text(String(value2 || '-'), margin + metadeLargura + 50, posicaoY);
            
            posicaoY += 6;
        }
        
        var dados = {};
        $('#modalDetalhesBody [data-label]').each(function() {
            var label = $(this).data('label');
            var value = $(this).text().trim();
            if (label) {
                dados[label] = value;
            }
        });
        
        adicionarTitulo('Dados da Solicitação');

        adicionarLinhaDuasColunas('Nº Solicitação', idSolicitacao, 'Status', dados['Status']);
        adicionarLinhaDuasColunas('Data Solicitação', dados['DataCadastro'], 'Data Validação', dados['DataValidacao']);
        adicionarLinha('Solicitante', dados['NomeSolicitante']);
        adicionarLinha('Último Validador', dados['NomeValidador']);
        posicaoY += 4;
        
        adicionarTitulo('Dados do Cliente');
        
        adicionarLinha('Razão Social / Nome', dados['NomeCliente']);
        adicionarLinhaDuasColunas('CNPJ/CPF', dados['CNPJ'], 'Código Cliente', dados['CodCliente']);
        adicionarLinhaDuasColunas('Código Loja', dados['CodLoja'], 'Planta', dados['Planta']);
        adicionarLinha('Tipo de Mercado', dados['TipoMercado']);
        posicaoY += 4;
        
        adicionarTitulo('Limites de Crédito');
        
        adicionarLinhaDuasColunas('Limite Anterior', dados['LimiteAtual'], 'Limite Solicitado', dados['LimiteSolicitado']);
        adicionarLinha('Limite Validado', dados['LimiteValidado']);
        posicaoY += 4;
        
        adicionarTitulo('Análise de Crédito');
        
        adicionarLinhaDuasColunas('Score Serasa', dados['ScoreSerasa'], 'Score SW', dados['ScoreSW']);
        adicionarLinhaDuasColunas('Classe Risco Serasa', dados['ClasseRiscoSerasa'], 'Classe Risco SW', dados['ClasseRiscoSW']);
        adicionarLinhaDuasColunas('Limite Sug. Serasa', dados['LimiteSugeridoSerasa'], 'Limite Sug. SW', dados['LimiteSugeridoSW']);
        adicionarLinhaDuasColunas('Valor Seguradora', dados['ValorSeguradora'], 'Ano do Balanço', dados['BalancoAno']);
        adicionarLinha('Faturamento Est. Positivo', dados['FaturamentoEstimadoPositivo']);
        posicaoY += 4;
        
        adicionarTitulo('Informações Comerciais');
        
        adicionarLinhaDuasColunas('Volume M²', dados['VolumeM2'], 'Preço M²', dados['PrecoM2']);
        adicionarLinhaDuasColunas('P. Volume mensal (unidades)', dados['VolumeUnidades'], 'P. Volume mensal (toneladas)', dados['VolumeToneladas']);
        adicionarLinhaDuasColunas('Preço Kg', dados['PrecoKg'], 'Valor Ton. Tora', dados['ValorToneladaTora']);
        adicionarLinhaDuasColunas('Valor Unit. Semente', dados['ValorUnitarioSemente'], 'Valor Unit. Muda', dados['ValorUnitarioMuda']);
        adicionarLinha('Condição de Pagamento', dados['CondicaoPagamento']);
        posicaoY += 4;
        
        adicionarTitulo('Fluxo de Aprovação');
        
        var aprovadores = [];
        if (dados['ValidadorFinanceiro'] && dados['ValidadorFinanceiro'] !== '-') {
            aprovadores.push(['Financeiro', dados['ValidadorFinanceiro'], dados['DataValidacaoFinanceiro'] || '-', dados['ObservacaoFinanceiro'] || '-']);
        }
        if (dados['ValidadorComercial'] && dados['ValidadorComercial'] !== '-') {
            aprovadores.push(['Comercial', dados['ValidadorComercial'], dados['DataValidacaoComercial'] || '-', dados['ObservacaoComercial'] || '-']);
        }
        if (dados['ValidadorControladoria'] && dados['ValidadorControladoria'] !== '-') {
            aprovadores.push(['Controladoria', dados['ValidadorControladoria'], dados['DataValidacaoControladoria'] || '-', dados['ObservacaoControladoria'] || '-']);
        }
        if (dados['ValidadorGerenciaGeral'] && dados['ValidadorGerenciaGeral'] !== '-') {
            aprovadores.push(['Gerência Geral', dados['ValidadorGerenciaGeral'], dados['DataValidacaoGerenciaGeral'] || '-', dados['ObservacaoGerenciaGeral'] || '-']);
        }
        if (dados['ValidadorDiretoriaComercial'] && dados['ValidadorDiretoriaComercial'] !== '-') {
            aprovadores.push(['Dir. Comercial', dados['ValidadorDiretoriaComercial'], dados['DataValidacaoDiretoriaComercial'] || '-', dados['ObservacaoDiretoriaComercial'] || '-']);
        }
        if (dados['ValidadorCFO'] && dados['ValidadorCFO'] !== '-') {
            aprovadores.push(['CFO', dados['ValidadorCFO'], dados['DataValidacaoCFO'] || '-', dados['ObservacaoCFO'] || '-']);
        }
        if (dados['ValidadorCEO'] && dados['ValidadorCEO'] !== '-') {
            aprovadores.push(['CEO', dados['ValidadorCEO'], dados['DataValidacaoCEO'] || '-', dados['ObservacaoCEO'] || '-']);
        }
        
        if (aprovadores.length > 0) {
            doc.autoTable({
                startY: posicaoY,
                head: [['Nível', 'Aprovador', 'Data', 'Observação']],
                body: aprovadores,
                theme: 'grid',
                styles: { 
                    font: 'helvetica',
                    fontSize: 8,
                    cellPadding: 2,
                    lineColor: [222, 226, 230],
                    lineWidth: 0.3,
                    textColor: [33, 37, 41]
                },
                headStyles: { 
                    fillColor: [248, 249, 250],
                    textColor: [73, 80, 87],
                    fontStyle: 'bold'
                },
                columnStyles: {
                    0: { cellWidth: 25 },
                    1: { cellWidth: 45 },
                    2: { cellWidth: 22 },
                    3: { cellWidth: 'auto' }
                },
                margin: { left: margin, right: margin }
            });
            posicaoY = doc.lastAutoTable.finalY + 5;
        } else {
            doc.setFontSize(9);
            doc.setFont('helvetica', 'italic');
            doc.setTextColor(108, 117, 125);
            doc.text('Nenhuma aprovação registrada até o momento.', margin, posicaoY);
            posicaoY += 8;
        }
        
        window.open(doc.output('bloburl'), '_blank');
    }
</script>

<div class="container-fluid p-0">
    <div class="border-bottom border-dashed pb-3 mb-3 d-flex justify-content-end d-print-none">
        <button type="button" class="btn btn-danger btn-sm" onclick="exportarDetalhesPDF()">
            <i class="fas fa-file-pdf me-1"></i> Exportar PDF
        </button>
    </div>

    <div class="card mb-3">
        <div class="card-header bg-light fw-semibold text-secondary py-2">
            Dados da Solicitação
        </div>
        <div class="card-body">
            <div class="row g-3">
                <div class="col-md-6">
                    <small class="text-muted text-uppercase d-block">Nº da Solicitação</small>
                    <span class="fw-semibold" data-label="IdSolicitacao"><%=solicitacao("IdSolicitacao")%></span>
                </div>
                <div class="col-md-6">
                    <small class="text-muted text-uppercase d-block">Status</small>
                    <span data-label="Status"><%=ObterNomeStatus(solicitacao("Status"))%></span>
                </div>
                <div class="col-md-6">
                    <small class="text-muted text-uppercase d-block">Solicitante</small>
                    <span data-label="NomeSolicitante"><%=ExibirValor(solicitacao("NomeSolicitante"))%></span>
                </div>
                <div class="col-md-6">
                    <small class="text-muted text-uppercase d-block">Login</small>
                    <span data-label="LoginSolicitante"><%=ExibirValor(solicitacao("LoginSolicitante"))%></span>
                </div>
                <div class="col-md-6">
                    <small class="text-muted text-uppercase d-block">Data da Solicitação</small>
                    <span data-label="DataCadastro"><%=FormatarData(solicitacao("DataCadastro"))%></span>
                </div>
                <div class="col-md-6">
                    <small class="text-muted text-uppercase d-block">Data Validação</small>
                    <span data-label="DataValidacao"><%=FormatarData(solicitacao("DataValidacao"))%></span>
                </div>
                <div class="col-12">
                    <small class="text-muted text-uppercase d-block">Último Validador</small>
                    <span data-label="NomeValidador"><%=ExibirValor(solicitacao("NomeValidador"))%></span>
                </div>
            </div>
        </div>
    </div>

    <div class="card mb-3">
        <div class="card-header bg-light fw-semibold text-secondary py-2">
            Dados do Cliente
        </div>
        <div class="card-body">
            <div class="row g-3">
                <div class="col-md-6">
                    <small class="text-muted text-uppercase d-block">Razão Social / Nome</small>
                    <span class="fw-semibold" data-label="NomeCliente"><%=ExibirValor(solicitacao("NomeCliente"))%></span>
                </div>
                <div class="col-md-6">
                    <small class="text-muted text-uppercase d-block">CNPJ/CPF</small>
                    <span data-label="CNPJ"><%=FormataCnpj(solicitacao("CNPJ"))%></span>
                </div>
                <div class="col-md-3">
                    <small class="text-muted text-uppercase d-block">Código Cliente</small>
                    <span data-label="CodCliente"><%=ExibirValor(solicitacao("CodCliente"))%></span>
                </div>
                <div class="col-md-3">
                    <small class="text-muted text-uppercase d-block">Código Loja</small>
                    <span data-label="CodLoja"><%=ExibirValor(solicitacao("CodLoja"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Planta</small>
                    <span data-label="Planta"><%=ObterNomePlanta(solicitacao("CodPlanta"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Tipo de Mercado</small>
                    <span data-label="TipoMercado"><%=ExibirValor(solicitacao("TipoMercado"))%></span>
                </div>
            </div>
        </div>
    </div>

    <div class="card mb-3">
        <div class="card-header bg-light fw-semibold text-secondary py-2">
            Limites de Crédito
        </div>
        <div class="card-body">
            <div class="row g-3">
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Limite Anterior</small>
                    <span data-label="LimiteAtual"><%=FormatarMoeda(solicitacao("LimiteAtual"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Limite Solicitado</small>
                    <span data-label="LimiteSolicitado"><%=FormatarMoeda(solicitacao("LimiteSolicitado"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Limite Validado</small>
                    <span data-label="LimiteValidado"><%=FormatarMoeda(solicitacao("LimiteValidado"))%></span>
                </div>
            </div>
        </div>
    </div>

    <div class="card mb-3">
        <div class="card-header bg-light fw-semibold text-secondary py-2">
            Análise de Crédito
        </div>
        <div class="card-body">
            <div class="row g-3">
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Score Serasa</small>
                    <span data-label="ScoreSerasa"><%=ExibirValor(solicitacao("ScoreSerasa"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Score SW</small>
                    <span data-label="ScoreSW"><%=ExibirValor(solicitacao("ScoreSW"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Classe de Risco Serasa</small>
                    <span data-label="ClasseRiscoSerasa"><%=ExibirValor(solicitacao("ClasseRiscoSerasa"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Classe de Risco SW</small>
                    <span data-label="ClasseRiscoSW"><%=ExibirValor(solicitacao("ClasseRiscoSW"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Limite Sugerido Serasa</small>
                    <span data-label="LimiteSugeridoSerasa"><%=FormatarMoeda(solicitacao("LimiteSugeridoSerasa"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Limite Sugerido SW</small>
                    <span data-label="LimiteSugeridoSW"><%=FormatarMoeda(solicitacao("LimiteSugeridoSW"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Valor Seguradora</small>
                    <span data-label="ValorSeguradora"><%=FormatarMoeda(solicitacao("ValorSeguradora"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Ano do Balanço</small>
                    <span data-label="BalancoAno"><%=ExibirValor(solicitacao("BalancoAno"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Faturamento Est. Positivo</small>
                    <span data-label="FaturamentoEstimadoPositivo"><%=ExibirValor(solicitacao("FaturamentoEstimadoPositivo"))%></span>
                </div>
            </div>
        </div>
    </div>

    <div class="card mb-3">
        <div class="card-header bg-light fw-semibold text-secondary py-2">
            Informações Comerciais
        </div>
        <div class="card-body">
            <div class="row g-3">
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Volume M²</small>
                    <span data-label="VolumeM2"><%=ExibirValor(solicitacao("VolumeM2"))%></span>
                </div>
                 <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Potencial Volume mensal (unidades)</small>
                    <span data-label="VolumeUnidades"><%=ExibirValor(solicitacao("VolumeUnidades"))%></span>
                </div>
                 <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Potencial Volume mensal (toneladas)</small>
                    <span data-label="VolumeToneladas"><%=ExibirValor(solicitacao("VolumeToneladas"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Preço M²</small>
                    <span data-label="PrecoM2"><%=FormatarMoeda(solicitacao("PrecoM2"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Preço Kg</small>
                    <span data-label="PrecoKg"><%=FormatarMoeda(solicitacao("PrecoKg"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Valor Tonelada Tora</small>
                    <span data-label="ValorToneladaTora"><%=FormatarMoeda(solicitacao("ValorToneladaTora"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Valor Unit. Semente</small>
                    <span data-label="ValorUnitarioSemente"><%=FormatarMoeda(solicitacao("ValorUnitarioSemente"))%></span>
                </div>
                <div class="col-md-4">
                    <small class="text-muted text-uppercase d-block">Valor Unit. Muda</small>
                    <span data-label="ValorUnitarioMuda"><%=FormatarMoeda(solicitacao("ValorUnitarioMuda"))%></span>
                </div>
                <div class="col-12">
                    <small class="text-muted text-uppercase d-block">Condição de Pagamento</small>
                    <span data-label="CondicaoPagamento"><%=ExibirValor(solicitacao("CondicaoPagamento"))%></span>
                </div>
            </div>
        </div>
    </div>

    <div class="card mb-3">
        <div class="card-header bg-light fw-semibold text-secondary py-2">
            Fluxo de Aprovação
        </div>
        <div class="card-body p-0">
            <% 
            Dim temAprovadores
            temAprovadores = False
            
            If solicitacao("ValidadorFinanceiro") <> "" Then temAprovadores = True
            If solicitacao("ValidadorComercial") <> "" Then temAprovadores = True
            If solicitacao("ValidadorControladoria") <> "" Then temAprovadores = True
            If solicitacao("ValidadorGerenciaGeral") <> "" Then temAprovadores = True
            If solicitacao("ValidadorDiretoriaComercial") <> "" Then temAprovadores = True
            If solicitacao("ValidadorCFO") <> "" Then temAprovadores = True
            If solicitacao("ValidadorCEO") <> "" Then temAprovadores = True
            
            If Not temAprovadores Then
            %>
                <div class="p-3">
                    <p class="text-muted fst-italic mb-0">Nenhuma aprovação registrada até o momento.</p>
                </div>
                <span data-label="ValidadorFinanceiro" class="d-none">-</span>
                <span data-label="DataValidacaoFinanceiro" class="d-none">-</span>
                <span data-label="ObservacaoFinanceiro" class="d-none">-</span>
                <span data-label="ValidadorComercial" class="d-none">-</span>
                <span data-label="DataValidacaoComercial" class="d-none">-</span>
                <span data-label="ObservacaoComercial" class="d-none">-</span>
                <span data-label="ValidadorControladoria" class="d-none">-</span>
                <span data-label="DataValidacaoControladoria" class="d-none">-</span>
                <span data-label="ObservacaoControladoria" class="d-none">-</span>
                <span data-label="ValidadorGerenciaGeral" class="d-none">-</span>
                <span data-label="DataValidacaoGerenciaGeral" class="d-none">-</span>
                <span data-label="ObservacaoGerenciaGeral" class="d-none">-</span>
                <span data-label="ValidadorDiretoriaComercial" class="d-none">-</span>
                <span data-label="DataValidacaoDiretoriaComercial" class="d-none">-</span>
                <span data-label="ObservacaoDiretoriaComercial" class="d-none">-</span>
                <span data-label="ValidadorCFO" class="d-none">-</span>
                <span data-label="DataValidacaoCFO" class="d-none">-</span>
                <span data-label="ObservacaoCFO" class="d-none">-</span>
                <span data-label="ValidadorCEO" class="d-none">-</span>
                <span data-label="DataValidacaoCEO" class="d-none">-</span>
                <span data-label="ObservacaoCEO" class="d-none">-</span>
            <% 
            Else 
            %>
                <table class="table table-sm table-bordered mb-0">
                    <thead class="table-light">
                        <tr>
                            <th style="width: 15%;">Nível</th>
                            <th style="width: 25%;">Aprovador</th>
                            <th style="width: 12%;">Data</th>
                            <th>Observação</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% If solicitacao("ValidadorFinanceiro") <> "" Then %>
                        <tr>
                            <td>Financeiro</td>
                            <td data-label="ValidadorFinanceiro"><%=solicitacao("ValidadorFinanceiro")%></td>
                            <td data-label="DataValidacaoFinanceiro"><%=FormatarData(solicitacao("DataValidacaoFinanceiro"))%></td>
                            <td data-label="ObservacaoFinanceiro"><%=ExibirValor(solicitacao("ObservacaoFinanceiro"))%></td>
                        </tr>
                        <% Else %>
                        <span data-label="ValidadorFinanceiro" class="d-none">-</span>
                        <span data-label="DataValidacaoFinanceiro" class="d-none">-</span>
                        <span data-label="ObservacaoFinanceiro" class="d-none">-</span>
                        <% End If %>
                        
                        <% If solicitacao("ValidadorComercial") <> "" Then %>
                        <tr>
                            <td>Comercial</td>
                            <td data-label="ValidadorComercial"><%=solicitacao("ValidadorComercial")%></td>
                            <td data-label="DataValidacaoComercial"><%=FormatarData(solicitacao("DataValidacaoComercial"))%></td>
                            <td data-label="ObservacaoComercial"><%=ExibirValor(solicitacao("ObservacaoComercial"))%></td>
                        </tr>
                        <% Else %>
                        <span data-label="ValidadorComercial" class="d-none">-</span>
                        <span data-label="DataValidacaoComercial" class="d-none">-</span>
                        <span data-label="ObservacaoComercial" class="d-none">-</span>
                        <% End If %>
                        
                        <% If solicitacao("ValidadorControladoria") <> "" Then %>
                        <tr>
                            <td>Controladoria</td>
                            <td data-label="ValidadorControladoria"><%=solicitacao("ValidadorControladoria")%></td>
                            <td data-label="DataValidacaoControladoria"><%=FormatarData(solicitacao("DataValidacaoControladoria"))%></td>
                            <td data-label="ObservacaoControladoria"><%=ExibirValor(solicitacao("ObservacaoControladoria"))%></td>
                        </tr>
                        <% Else %>
                        <span data-label="ValidadorControladoria" class="d-none">-</span>
                        <span data-label="DataValidacaoControladoria" class="d-none">-</span>
                        <span data-label="ObservacaoControladoria" class="d-none">-</span>
                        <% End If %>
                        
                        <% If solicitacao("ValidadorGerenciaGeral") <> "" Then %>
                        <tr>
                            <td>Gerência Geral</td>
                            <td data-label="ValidadorGerenciaGeral"><%=solicitacao("ValidadorGerenciaGeral")%></td>
                            <td data-label="DataValidacaoGerenciaGeral"><%=FormatarData(solicitacao("DataValidacaoGerenciaGeral"))%></td>
                            <td data-label="ObservacaoGerenciaGeral"><%=ExibirValor(solicitacao("ObservacaoGerenciaGeral"))%></td>
                        </tr>
                        <% Else %>
                        <span data-label="ValidadorGerenciaGeral" class="d-none">-</span>
                        <span data-label="DataValidacaoGerenciaGeral" class="d-none">-</span>
                        <span data-label="ObservacaoGerenciaGeral" class="d-none">-</span>
                        <% End If %>
                        
                        <% If solicitacao("ValidadorDiretoriaComercial") <> "" Then %>
                        <tr>
                            <td>Dir. Comercial</td>
                            <td data-label="ValidadorDiretoriaComercial"><%=solicitacao("ValidadorDiretoriaComercial")%></td>
                            <td data-label="DataValidacaoDiretoriaComercial"><%=FormatarData(solicitacao("DataValidacaoDiretoriaComercial"))%></td>
                            <td data-label="ObservacaoDiretoriaComercial"><%=ExibirValor(solicitacao("ObservacaoDiretoriaComercial"))%></td>
                        </tr>
                        <% Else %>
                        <span data-label="ValidadorDiretoriaComercial" class="d-none">-</span>
                        <span data-label="DataValidacaoDiretoriaComercial" class="d-none">-</span>
                        <span data-label="ObservacaoDiretoriaComercial" class="d-none">-</span>
                        <% End If %>
                        
                        <% If solicitacao("ValidadorCFO") <> "" Then %>
                        <tr>
                            <td>CFO</td>
                            <td data-label="ValidadorCFO"><%=solicitacao("ValidadorCFO")%></td>
                            <td data-label="DataValidacaoCFO"><%=FormatarData(solicitacao("DataValidacaoCFO"))%></td>
                            <td data-label="ObservacaoCFO"><%=ExibirValor(solicitacao("ObservacaoCFO"))%></td>
                        </tr>
                        <% Else %>
                        <span data-label="ValidadorCFO" class="d-none">-</span>
                        <span data-label="DataValidacaoCFO" class="d-none">-</span>
                        <span data-label="ObservacaoCFO" class="d-none">-</span>
                        <% End If %>
                        
                        <% If solicitacao("ValidadorCEO") <> "" Then %>
                        <tr>
                            <td>CEO</td>
                            <td data-label="ValidadorCEO"><%=solicitacao("ValidadorCEO")%></td>
                            <td data-label="DataValidacaoCEO"><%=FormatarData(solicitacao("DataValidacaoCEO"))%></td>
                            <td data-label="ObservacaoCEO"><%=ExibirValor(solicitacao("ObservacaoCEO"))%></td>
                        </tr>
                        <% Else %>
                        <span data-label="ValidadorCEO" class="d-none">-</span>
                        <span data-label="DataValidacaoCEO" class="d-none">-</span>
                        <span data-label="ObservacaoCEO" class="d-none">-</span>
                        <% End If %>
                    </tbody>
                </table>
            <% End If %>
        </div>
    </div>
</div>
