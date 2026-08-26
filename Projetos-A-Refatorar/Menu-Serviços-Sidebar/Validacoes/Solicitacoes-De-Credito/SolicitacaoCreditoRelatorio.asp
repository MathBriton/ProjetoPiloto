<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<!--#include file="./util/Util.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer  = true
Response.Expires = 0
Session.LCID = 1046

Dim solicitacoes()
Dim transferencias()

Dim idx
idx = 0
totalSolicitacoes = 0
totalTransferencias = 0

Dim idUsuarioLogado
idUsuarioLogado = Session("CODIGO")

Dim abaSelecionada
abaSelecionada = Request.QueryString("aba")
If abaSelecionada = "" Then
    abaSelecionada = "solicitacoes"
End If

Dim filtroDataSolicitacaoDe, filtroDataSolicitacaoAte
Dim filtroDataAprovacaoDe, filtroDataAprovacaoAte
Dim filtroPlanta, filtroCodCliente, filtroCodLoja, filtroIdSolicitacao, filtroStatus
Dim possuiFiltroData

filtroDataSolicitacaoDe = Request("filtroDataSolicitacaoDe")
filtroDataSolicitacaoAte = Request("filtroDataSolicitacaoAte")
filtroDataAprovacaoDe = Request("filtroDataAprovacaoDe")
filtroDataAprovacaoAte = Request("filtroDataAprovacaoAte")
filtroPlanta = Request("filtroPlanta")
filtroCodCliente = Request("filtroCodCliente")
filtroCodLoja = Trim(Request("filtroCodLoja"))
filtroIdSolicitacao = Request("filtroIdSolicitacao")
filtroStatus = Request("filtroStatus")

Dim filtroDataSolicitacaoDeTransf, filtroDataSolicitacaoAteTransf
Dim filtroDataAprovacaoDeTransf, filtroDataAprovacaoAteTransf
Dim filtroPlantaOrigem, filtroPlantaDestino
Dim filtroCodClienteOrigem, filtroCodClienteDestino
Dim filtroCodLojaOrigem, filtroCodLojaDestino
Dim filtroIdTransferencia, filtroStatusTransf
Dim possuiFiltroDataTransf

filtroDataSolicitacaoDeTransf = Request("filtroDataSolicitacaoDeTransf")
filtroDataSolicitacaoAteTransf = Request("filtroDataSolicitacaoAteTransf")
filtroDataAprovacaoDeTransf = Request("filtroDataAprovacaoDeTransf")
filtroDataAprovacaoAteTransf = Request("filtroDataAprovacaoAteTransf")
filtroPlantaOrigem = Request("filtroPlantaOrigem")
filtroPlantaDestino = Request("filtroPlantaDestino")
filtroCodClienteOrigem = Request("filtroCodClienteOrigem")
filtroCodClienteDestino = Request("filtroCodClienteDestino")
filtroCodLojaOrigem = Trim(Request("filtroCodLojaOrigem"))
filtroCodLojaDestino = Trim(Request("filtroCodLojaDestino"))
filtroIdTransferencia = Request("filtroIdTransferencia")
filtroStatusTransf = Request("filtroStatusTransf")

possuiFiltroData = False
If filtroDataSolicitacaoDe <> "" Or filtroDataSolicitacaoAte <> "" Or filtroDataAprovacaoDe <> "" Or filtroDataAprovacaoAte <> "" Or filtroCodCliente <> "" Then
    possuiFiltroData = True
End If

possuiFiltroDataTransf = False
If filtroDataSolicitacaoDeTransf <> "" Or filtroDataSolicitacaoAteTransf <> "" Or filtroDataAprovacaoDeTransf <> "" Or filtroDataAprovacaoAteTransf <> "" Then
    possuiFiltroDataTransf = True
End If

ReDim solicitacoes(0)
ReDim transferencias(0)
idx = 0

Dim validador, codNivel

Dim usuarioLogado
Set usuarioLogado = CreateObject("Scripting.Dictionary")

If abaSelecionada = "solicitacoes" And possuiFiltroData Then
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
    SQL = SQL & "s.STATUS as Status, "
    SQL = SQL & "s.CNPJCPF as CNPJ, "
    SQL = SQL & "s.BalancoAno, "
    SQL = SQL & "CASE" 
    SQL = SQL & "    WHEN ISDATE(s.DTVALID) = 1 AND STATUS = 'V'"
    SQL = SQL & "    THEN CONVERT(VARCHAR(10), CONVERT(DATE, s.DTVALID, 112), 126) "
    SQL = SQL & "    ELSE NULL "
    SQL = SQL & "END AS DataValidacao, "
    SQL = SQL & "CASE"
    SQL = SQL & "    WHEN ISDATE(s.DTSOLICIT) = 1"
    SQL = SQL & "    THEN CONVERT(VARCHAR(10), CONVERT(DATE, s.DTSOLICIT, 112), 126) "
    SQL = SQL & "    ELSE NULL "
    SQL = SQL & "END AS DataCadastro, "
    SQL = SQL & "(SELECT COUNT(*) FROM SolicitacaoCreditoDocumento scd WHERE scd.IdSolicitacaoCredito = s.NUM) as QuantidadeAnexos "
    SQL = SQL & "FROM SOLCREDITOCLIENTE s "
    SQL = SQL & "WHERE 1=1 "
    
    If filtroIdSolicitacao <> "" Then
        SQL = SQL & "AND s.NUM = " & filtroIdSolicitacao & " "
    End If
    
    If filtroCodCliente <> "" Then
        SQL = SQL & "AND s.COD = '" & filtroCodCliente & "' "
        If filtroCodLoja <> "" Then
            SQL = SQL & "AND s.CODLOJA = '" & filtroCodLoja & "' "
        End If
    End If
    
    If filtroPlanta <> "" And filtroPlanta <> "TODOS" Then
        SQL = SQL & "AND s.Planta = '" & filtroPlanta & "' "
    End If
    
    If filtroStatus <> "" And filtroStatus <> "TODOS" Then
        SQL = SQL & "AND s.STATUS = '" & filtroStatus & "' "
    End If
    
    If filtroDataSolicitacaoDe <> "" Then
        SQL = SQL & "AND CONVERT(DATE, s.DTSOLICIT, 112) >= '" & Replace(filtroDataSolicitacaoDe, "-", "") & "' "
    End If
    
    If filtroDataSolicitacaoAte <> "" Then
        SQL = SQL & "AND CONVERT(DATE, s.DTSOLICIT, 112) <= '" & Replace(filtroDataSolicitacaoAte, "-", "") & "' "
    End If
    
    If filtroDataAprovacaoDe <> "" Then
        SQL = SQL & "AND s.STATUS = 'V' "
        SQL = SQL & "AND CONVERT(DATE, s.DTVALID, 112) >= '" & Replace(filtroDataAprovacaoDe, "-", "") & "' "
    End If
    
    If filtroDataAprovacaoAte <> "" Then
        SQL = SQL & "AND s.STATUS = 'V' "
        SQL = SQL & "AND CONVERT(DATE, s.DTVALID, 112) <= '" & Replace(filtroDataAprovacaoAte, "-", "") & "' "
    End If
    
    SQL = SQL & "ORDER BY s.NUM DESC"

    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    
    idx = 0
    While NOT PIR_CORPORATIVO.EOF
        ReDim Preserve solicitacoes(idx)
        Set solicitacoes(idx) = Server.CreateObject("Scripting.Dictionary")
        solicitacoes(idx)("IdSolicitacao") = PIR_CORPORATIVO("IdSolicitacao")
        solicitacoes(idx)("CodPlanta") = PIR_CORPORATIVO("CodPlanta")
        solicitacoes(idx)("LoginSolicitante") = PIR_CORPORATIVO("LoginSolicitante")
        solicitacoes(idx)("NomeSolicitante") = PIR_CORPORATIVO("NomeSolicitante")
        solicitacoes(idx)("NomeCliente") = PIR_CORPORATIVO("NomeCliente")
        solicitacoes(idx)("NomeValidador") = PIR_CORPORATIVO("NomeValidador")
        solicitacoes(idx)("CNPJ") = PIR_CORPORATIVO("CNPJ")
        solicitacoes(idx)("LimiteAtual") = PIR_CORPORATIVO("LimiteAtual")
        solicitacoes(idx)("LimiteNovo") = PIR_CORPORATIVO("LimiteNovo")
        solicitacoes(idx)("LimiteValidado") = PIR_CORPORATIVO("LimiteValidado")
        solicitacoes(idx)("LimiteSolicitado") = PIR_CORPORATIVO("LimiteSolicitado")
        solicitacoes(idx)("Status") = PIR_CORPORATIVO("Status")
        solicitacoes(idx)("BalancoAno") = PIR_CORPORATIVO("BalancoAno")
        solicitacoes(idx)("DataCadastro") = PIR_CORPORATIVO("DataCadastro")
        solicitacoes(idx)("DataValidacao") = PIR_CORPORATIVO("DataValidacao")
        solicitacoes(idx)("CodCliente") = PIR_CORPORATIVO("CodCliente")
        solicitacoes(idx)("CodLoja") = PIR_CORPORATIVO("CodLoja")
        solicitacoes(idx)("EnderecoDocumentacao") = PIR_CORPORATIVO("EnderecoDocumentacao")
        solicitacoes(idx)("QuantidadeAnexos") = PIR_CORPORATIVO("QuantidadeAnexos")
        
        idx = idx + 1
        PIR_CORPORATIVO.MoveNext()
    Wend
    
    If isObject(PIR_CORPORATIVO) Then
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If

    totalSolicitacoes = idx
End If

If abaSelecionada = "transferencias" And possuiFiltroDataTransf Then
    SQL = "SELECT "
    SQL = SQL & "tc.Id, "
    SQL = SQL & "tc.NomeCliente, "
    SQL = SQL & "tc.NomeSolicitante, "
    SQL = SQL & "tc.NomeAprovador, "
    SQL = SQL & "tc.Status, "
    SQL = SQL & "tc.LimiteAtualOrigem, "
    SQL = SQL & "tc.LimiteAtualDestino, "
    SQL = SQL & "tc.ValorTransferencia, "
    SQL = SQL & "tc.PlantaOrigem, "
    SQL = SQL & "tc.CodClienteOrigem, "
    SQL = SQL & "tc.CodLojaOrigem, "
    SQL = SQL & "tc.PlantaDestino, "
    SQL = SQL & "tc.CodClienteDestino, "
    SQL = SQL & "tc.CodLojaDestino, "
    SQL = SQL & "tc.Observacao, "
    SQL = SQL & "tc.CNPJCPF as CNPJ, "
    SQL = SQL & "CASE "
    SQL = SQL & "    WHEN ISDATE(tc.DataCriacao) = 1 "
    SQL = SQL & "    THEN CONVERT(VARCHAR(10), CONVERT(DATE, tc.DataCriacao, 120), 126) "
    SQL = SQL & "    ELSE NULL "
    SQL = SQL & "END AS DataCriacao, "
    SQL = SQL & "CASE "
    SQL = SQL & "    WHEN ISDATE(tc.DataValidacao) = 1 "
    SQL = SQL & "    THEN CONVERT(VARCHAR(10), CONVERT(DATE, tc.DataValidacao, 120), 126) "
    SQL = SQL & "    ELSE NULL "
    SQL = SQL & "END AS DataValidacao "
    SQL = SQL & "FROM TransferenciaCredito tc "
    SQL = SQL & "WHERE tc.FlagInativo = 0 "
    
    If filtroIdTransferencia <> "" Then
        SQL = SQL & "AND tc.Id = " & filtroIdTransferencia & " "
    End If
    
    If filtroCodClienteOrigem <> "" Then
        SQL = SQL & "AND tc.CodClienteOrigem = '" & filtroCodClienteOrigem & "' "
        If filtroCodLojaOrigem <> "" Then
            SQL = SQL & "AND tc.CodLojaOrigem = '" & filtroCodLojaOrigem & "' "
        End If
    End If
    
    If filtroCodClienteDestino <> "" Then
        SQL = SQL & "AND tc.CodClienteDestino = '" & filtroCodClienteDestino & "' "
        If filtroCodLojaDestino <> "" Then
            SQL = SQL & "AND tc.CodLojaDestino = '" & filtroCodLojaDestino & "' "
        End If
    End If
    
    If filtroPlantaOrigem <> "" And filtroPlantaOrigem <> "TODOS" Then
        SQL = SQL & "AND tc.PlantaOrigem = '" & filtroPlantaOrigem & "' "
    End If
    
    If filtroPlantaDestino <> "" And filtroPlantaDestino <> "TODOS" Then
        SQL = SQL & "AND tc.PlantaDestino = '" & filtroPlantaDestino & "' "
    End If
    
    If filtroStatusTransf <> "" And filtroStatusTransf <> "TODOS" Then
        SQL = SQL & "AND tc.Status = '" & filtroStatusTransf & "' "
    End If
    
    If filtroDataSolicitacaoDeTransf <> "" Then
        SQL = SQL & "AND CONVERT(DATE, tc.DataCriacao, 120) >= '" & filtroDataSolicitacaoDeTransf & "' "
    End If
    
    If filtroDataSolicitacaoAteTransf <> "" Then
        SQL = SQL & "AND CONVERT(DATE, tc.DataCriacao, 120) <= '" & filtroDataSolicitacaoAteTransf & "' "
    End If
    
    If filtroDataAprovacaoDeTransf <> "" Then
        SQL = SQL & "AND tc.Status = 'APROVADO' "
        SQL = SQL & "AND CONVERT(DATE, tc.DataValidacao, 120) >= '" & filtroDataAprovacaoDeTransf & "' "
    End If
    
    If filtroDataAprovacaoAteTransf <> "" Then
        SQL = SQL & "AND tc.Status = 'APROVADO' "
        SQL = SQL & "AND CONVERT(DATE, tc.DataValidacao, 120) <= '" & filtroDataAprovacaoAteTransf & "' "
    End If
    
    SQL = SQL & "ORDER BY tc.Id DESC"
    
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    idx = 0
    While NOT PIR_CORPORATIVO.EOF
        ReDim Preserve transferencias(idx)
        Set transferencias(idx) = Server.CreateObject("Scripting.Dictionary")
        transferencias(idx)("Id") = PIR_CORPORATIVO("Id")
        transferencias(idx)("NomeCliente") = PIR_CORPORATIVO("NomeCliente")
        transferencias(idx)("NomeSolicitante") = PIR_CORPORATIVO("NomeSolicitante")
        transferencias(idx)("NomeAprovador") = PIR_CORPORATIVO("NomeAprovador")
        transferencias(idx)("CNPJ") = PIR_CORPORATIVO("CNPJ")
        transferencias(idx)("CodClienteOrigem") = PIR_CORPORATIVO("CodClienteOrigem")
        transferencias(idx)("LimiteAtualOrigem") = PIR_CORPORATIVO("LimiteAtualOrigem")
        transferencias(idx)("LimiteAtualDestino") = PIR_CORPORATIVO("LimiteAtualDestino")
        transferencias(idx)("PlantaOrigem") = PIR_CORPORATIVO("PlantaOrigem")
        transferencias(idx)("PlantaDestino") = PIR_CORPORATIVO("PlantaDestino")
        transferencias(idx)("DataCriacao") = PIR_CORPORATIVO("DataCriacao")
        transferencias(idx)("DataValidacao") = PIR_CORPORATIVO("DataValidacao")
        transferencias(idx)("Status") = PIR_CORPORATIVO("Status")
        idx = idx + 1
        PIR_CORPORATIVO.MoveNext()
    Wend
    If isObject(PIR_CORPORATIVO) Then
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    totalTransferencias = idx
End If
%>

<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <title>Smurfit Westrock</title>
        <meta charset="utf-8"/>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
        <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />    
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
        <script src="./util/Util.js"></script>
        <script type="text/javascript">
            function alterarAba(aba) {
                window.location.href = window.location.pathname + '?aba=' + aba;
            }

            function limparFiltros() {
                var aba = '<%=abaSelecionada%>';
                if (aba === 'solicitacoes') {
                    document.getElementById('filtroDataSolicitacaoDe').value = '';
                    document.getElementById('filtroDataSolicitacaoAte').value = '';
                    document.getElementById('filtroDataAprovacaoDe').value = '';
                    document.getElementById('filtroDataAprovacaoAte').value = '';
                    document.getElementById('filtroPlanta').value = 'TODOS';
                    document.getElementById('filtroCodCliente').value = '';
                    document.getElementById('filtroCodLoja').value = '';
                    document.getElementById('filtroIdSolicitacao').value = '';
                    document.getElementById('filtroStatus').value = 'TODOS';
                } else {
                    document.getElementById('filtroDataSolicitacaoDeTransf').value = '';
                    document.getElementById('filtroDataSolicitacaoAteTransf').value = '';
                    document.getElementById('filtroDataAprovacaoDeTransf').value = '';
                    document.getElementById('filtroDataAprovacaoAteTransf').value = '';
                    document.getElementById('filtroPlantaOrigem').value = 'TODOS';
                    document.getElementById('filtroPlantaDestino').value = 'TODOS';
                    document.getElementById('filtroCodClienteOrigem').value = '';
                    document.getElementById('filtroCodLojaOrigem').value = '';
                    document.getElementById('filtroCodClienteDestino').value = '';
                    document.getElementById('filtroCodLojaDestino').value = '';
                    document.getElementById('filtroIdTransferencia').value = '';
                    document.getElementById('filtroStatusTransf').value = 'TODOS';
                }
            }

            function validarFiltros() {
                var aba = '<%=abaSelecionada%>';
                
                if (aba === 'solicitacoes') {
                    var dataSolicitacaoDe = document.getElementById('filtroDataSolicitacaoDe').value;
                    var dataSolicitacaoAte = document.getElementById('filtroDataSolicitacaoAte').value;
                    var dataAprovacaoDe = document.getElementById('filtroDataAprovacaoDe').value;
                    var dataAprovacaoAte = document.getElementById('filtroDataAprovacaoAte').value;
                    var filtroCodCliente = document.getElementById('filtroCodCliente').value.trim();
                    var filtroCodLoja = document.getElementById('filtroCodLoja').value.trim();

                    if (!dataSolicitacaoDe && !dataSolicitacaoAte && !dataAprovacaoDe && !dataAprovacaoAte && !filtroCodCliente) {
                        alert('Por favor, informe pelo menos uma data de solicitação ou aprovação para realizar a pesquisa.');
                        return false;
                    }

                    if (filtroCodLoja && !filtroCodCliente) {
                        alert('Preencha o Código do Cliente antes de filtrar pela Loja.');
                        document.getElementById('filtroCodCliente').focus();
                        return false;
                    }

                    if(dataSolicitacaoAte && !dataSolicitacaoDe) {
                        alert('Para pesquisar pela data de solicitação indique a data de início.');
                        return false;
                    }

                    if(dataAprovacaoAte && !dataAprovacaoDe) {
                        alert('Para pesquisar pela data de aprovação indique a data de início.');
                        return false;
                    }
                } else {
                    var dataSolicitacaoDe = document.getElementById('filtroDataSolicitacaoDeTransf').value;
                    var dataSolicitacaoAte = document.getElementById('filtroDataSolicitacaoAteTransf').value;
                    var dataAprovacaoDe = document.getElementById('filtroDataAprovacaoDeTransf').value;
                    var dataAprovacaoAte = document.getElementById('filtroDataAprovacaoAteTransf').value;
                    var codOrigem = document.getElementById('filtroCodClienteOrigem').value.trim();
                    var lojaOrigem = document.getElementById('filtroCodLojaOrigem').value.trim();
                    var codDestino = document.getElementById('filtroCodClienteDestino').value.trim();
                    var lojaDestino = document.getElementById('filtroCodLojaDestino').value.trim();

                    if (!dataSolicitacaoDe && !dataSolicitacaoAte && !dataAprovacaoDe && !dataAprovacaoAte) {
                        alert('Por favor, informe pelo menos uma data de solicitação ou aprovação para realizar a pesquisa.');
                        return false;
                    }

                    if (lojaOrigem && !codOrigem) {
                        alert('Preencha o Código do Cliente Origem antes de filtrar pela Loja Origem.');
                        document.getElementById('filtroCodClienteOrigem').focus();
                        return false;
                    }

                    if (lojaDestino && !codDestino) {
                        alert('Preencha o Código do Cliente Destino antes de filtrar pela Loja Destino.');
                        document.getElementById('filtroCodClienteDestino').focus();
                        return false;
                    }

                    if(dataSolicitacaoAte && !dataSolicitacaoDe) {
                        alert('Para pesquisar pela data de solicitação indique a data de início.');
                        return false;
                    }

                    if(dataAprovacaoAte && !dataAprovacaoDe) {
                        alert('Para pesquisar pela data de aprovação indique a data de início.');
                        return false;
                    }
                }
                
                return true;
            }

            function abrirHistoricoValidacoes(idSolicitacao) {                
                var modal = new bootstrap.Modal(document.getElementById('modalHistorico'));
                modal.show();
                carregarHistoricoValidacoes(idSolicitacao);
            }

            function carregarHistoricoValidacoes(idSolicitacao) {
                $.ajax({
                    url: 'SolicitacaoCreditoHistorico.asp?v=' + Math.random(),
                    type: 'GET',
                    data: { idSolicitacao: idSolicitacao },
                    success: function(response) {
                        $('#modalHistoricoBody').html(response);
                    },
                    error: function(xhr, status, error) {
                        $('#modalHistoricoBody').html('<div class="alert alert-danger">Erro ao carregar histórico: ' + error + '</div>');
                    }
                });
            }

            function abrirModalAnexos(idSolicitacao, enderecoLegado) {
                var modal = new bootstrap.Modal(document.getElementById('modalAnexos'));
                
                $('#modalAnexosBody').html('<div class="text-center py-4"><div class="spinner-border text-primary" role="status"><span class="visually-hidden">Carregando...</span></div><p class="mt-2">Carregando anexos...</p></div>');
                
                modal.show();
                
                $.ajax({
                    url: 'SolicitacaoCreditoAnexos.asp?v=' + Math.random(),
                    type: 'GET',
                    data: { 
                        idSolicitacao: idSolicitacao,
                        enderecoLegado: enderecoLegado || ''
                    },
                    success: function(response) {
                        $('#modalAnexosBody').html(response);
                    },
                    error: function(xhr, status, error) {
                        $('#modalAnexosBody').html('<div class="alert alert-danger">Erro ao carregar anexos: ' + error + '</div>');
                    }
                });
            }

            function abrirModalDetalhes(idSolicitacao) {
                var modal = new bootstrap.Modal(document.getElementById('modalDetalhes'));
                
                $('#modalDetalhesBody').html('<div class="text-center py-4"><div class="spinner-border text-primary" role="status"><span class="visually-hidden">Carregando...</span></div><p class="mt-2">Carregando detalhes...</p></div>');
                $('#modalDetalhes').data('idSolicitacao', idSolicitacao);
                
                modal.show();
                
                $.ajax({
                    url: 'SolicitacaoCreditoDetalhes.asp?v=' + Math.random(),
                    type: 'GET',
                    data: { idSolicitacao: idSolicitacao },
                    success: function(response) {
                        $('#modalDetalhesBody').html(response);
                    },
                    error: function(xhr, status, error) {
                        $('#modalDetalhesBody').html('<div class="alert alert-danger">Erro ao carregar detalhes: ' + error + '</div>');
                    }
                });
            }

            function formatarDataBR(data) {
                if (!data) return '';
                var partes = data.split('-');
                return partes[2] + '/' + partes[1] + '/' + partes[0];
            }

            function obterNomePlantaPorCodigo(codigo) {
                var plantas = {
                    '00': 'Pirapetinga',
                    '04': 'Uberaba',
                    '05': 'Bento',
                    '06': 'Maranguape',
                    '07': 'Saquarema',
                    '08': 'Aracatuba',
                    '09': 'Blumenau',
                    '10': 'Pacajus',
                    '11': 'Porto Feliz',
                    '12': 'Tres Barras',
                    '13': 'Florestal'
                };
                return plantas[codigo] || codigo;
            }

            function obterNomeStatus(codigo) {
                var status = {
                    'A': 'Novo',
                    'D': 'Análise Finalizada',
                    'C': 'Validado Analista',
                    'E': 'Validado Coordenador',
                    'F': 'Validado Gerente de Tesouraria',
                    'G': 'Validado Controller',
                    'H': 'Validado Gerente Planta',
                    'V': 'Aprovado',
                    'B': 'Rejeitado'
                };
                return status[codigo] || codigo;
            }

            // ── Popup buscaCLIENTE ──────────────────────────────────────────
            function abrirBuscaCliente(campoCod, campoLoja, planta) {
                if (!planta || planta === 'TODOS') {
                    alert('Selecione uma Planta específica antes de buscar o cliente.');
                    return;
                }
                var url = '../buscaCLIENTE.asp'
                    + '?id1=' + campoCod
                    + '&id2=' + campoLoja
                    + '&id3='
                    + '&id4=1'
                    + '&id6=' + planta;
                window.open(url, 'buscaCliente', 'width=900,height=600,scrollbars=yes,resizable=yes');
            }

            function exportarPDF() {
                const { jsPDF } = window.jspdf;
                const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' });
                const dataAtual = new Date().toLocaleDateString('pt-BR');
                var aba = '<%=abaSelecionada%>';

                const tabela = document.getElementById(aba === 'solicitacoes' ? 'GradeSolicitacoes' : 'GradeTransferencias');
                const linhas = tabela.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
                
                const dados = [];
                var colunas = [];
                var yPosition = 15;
                
                if (aba === 'solicitacoes') {
                    colunas = ['#', 'Planta', 'Solicitante', 'Cód. Cliente', 'Nome Cliente', 'CNPJ', 'Limite Atual', 'Limite Solicitado', 'Limite Validado', 'Último Validador', 'Status', 'Indicador Ano', 'Data de Solicitação', 'Data de Aprovação'];
                    doc.setFontSize(16);
                    doc.text('Histórico de Solicitações de Crédito', 14, yPosition);
                    yPosition += 7;

                    doc.setFontSize(9);
                    doc.setFont(undefined, 'bold');
                    var filtrosAplicados = [];
                    
                    var filtroDataSolicitacaoDe = '<%=filtroDataSolicitacaoDe%>';
                    var filtroDataSolicitacaoAte = '<%=filtroDataSolicitacaoAte%>';
                    var filtroDataAprovacaoDe = '<%=filtroDataAprovacaoDe%>';
                    var filtroDataAprovacaoAte = '<%=filtroDataAprovacaoAte%>';
                    var filtroPlanta = '<%=filtroPlanta%>';
                    var filtroCodCliente = '<%=filtroCodCliente%>';
                    var filtroIdSolicitacao = '<%=filtroIdSolicitacao%>';
                    var filtroStatus = '<%=filtroStatus%>';

                    if (filtroDataSolicitacaoDe) {
                        filtrosAplicados.push('Data Solicitação (De): ' + formatarDataBR(filtroDataSolicitacaoDe));
                    }
                    if (filtroDataSolicitacaoAte) {
                        filtrosAplicados.push('Data Solicitação (Até): ' + formatarDataBR(filtroDataSolicitacaoAte));
                    }
                    if (filtroDataAprovacaoDe) {
                        filtrosAplicados.push('Data Aprovação (De): ' + formatarDataBR(filtroDataAprovacaoDe));
                    }
                    if (filtroDataAprovacaoAte) {
                        filtrosAplicados.push('Data Aprovação (Até): ' + formatarDataBR(filtroDataAprovacaoAte));
                    }
                    if (filtroPlanta && filtroPlanta !== 'TODOS') {
                        filtrosAplicados.push('Planta: ' + obterNomePlantaPorCodigo(filtroPlanta));
                    }
                    if (filtroCodCliente) {
                        var lojaLabel = '<%=filtroCodLoja%>';
                        filtrosAplicados.push('Cod. Cliente: ' + filtroCodCliente + (lojaLabel ? ' / Loja: ' + lojaLabel : ''));
                    }
                    if (filtroIdSolicitacao) {
                        filtrosAplicados.push('Id Solicitação: ' + filtroIdSolicitacao);
                    }
                    if (filtroStatus && filtroStatus !== 'TODOS') {
                        filtrosAplicados.push('Status: ' + obterNomeStatus(filtroStatus));
                    }

                    if (filtrosAplicados.length > 0) {
                        doc.text('Filtros Aplicados:', 14, yPosition);
                        yPosition += 5;
                        doc.setFont(undefined, 'normal');
                        for (var i = 0; i < filtrosAplicados.length; i++) {
                            doc.text('  • ' + filtrosAplicados[i], 14, yPosition);
                            yPosition += 4;
                        }
                        yPosition += 2;
                    }
                } else {
                    colunas = ['#', 'Solicitante', 'Aprovador', 'Cliente', 'CNPJ', 'Planta Origem', 'Planta Destino', 'Limite Anterior Origem', 'Limite Anterior Destino', 'Status', 'Data Solicitação', 'Data Aprovação'];
                    doc.setFontSize(16);
                    doc.text('Histórico de Transferências de Crédito', 14, yPosition);
                    yPosition += 7;

                    doc.setFontSize(9);
                    doc.setFont(undefined, 'bold');
                    var filtrosAplicados = [];
                    
                    var filtroDataSolicitacaoDeTransf = '<%=filtroDataSolicitacaoDeTransf%>';
                    var filtroDataSolicitacaoAteTransf = '<%=filtroDataSolicitacaoAteTransf%>';
                    var filtroDataAprovacaoDeTransf = '<%=filtroDataAprovacaoDeTransf%>';
                    var filtroDataAprovacaoAteTransf = '<%=filtroDataAprovacaoAteTransf%>';
                    var filtroPlantaOrigem = '<%=filtroPlantaOrigem%>';
                    var filtroPlantaDestino = '<%=filtroPlantaDestino%>';
                    var filtroCodClienteOrigem = '<%=filtroCodClienteOrigem%>';
                    var filtroCodClienteDestino = '<%=filtroCodClienteDestino%>';
                    var filtroIdTransferencia = '<%=filtroIdTransferencia%>';
                    var filtroStatusTransf = '<%=filtroStatusTransf%>';

                    if (filtroDataSolicitacaoDeTransf) {
                        filtrosAplicados.push('Data Solicitação (De): ' + formatarDataBR(filtroDataSolicitacaoDeTransf));
                    }
                    if (filtroDataSolicitacaoAteTransf) {
                        filtrosAplicados.push('Data Solicitação (Até): ' + formatarDataBR(filtroDataSolicitacaoAteTransf));
                    }
                    if (filtroDataAprovacaoDeTransf) {
                        filtrosAplicados.push('Data Aprovação (De): ' + formatarDataBR(filtroDataAprovacaoDeTransf));
                    }
                    if (filtroDataAprovacaoAteTransf) {
                        filtrosAplicados.push('Data Aprovação (Até): ' + formatarDataBR(filtroDataAprovacaoAteTransf));
                    }
                    if (filtroPlantaOrigem && filtroPlantaOrigem !== 'TODOS') {
                        filtrosAplicados.push('Planta Origem: ' + obterNomePlantaPorCodigo(filtroPlantaOrigem));
                    }
                    if (filtroPlantaDestino && filtroPlantaDestino !== 'TODOS') {
                        filtrosAplicados.push('Planta Destino: ' + obterNomePlantaPorCodigo(filtroPlantaDestino));
                    }
                    if (filtroCodClienteOrigem) {
                        var lojaOrigLabel = '<%=filtroCodLojaOrigem%>';
                        filtrosAplicados.push('Cod. Cliente Origem: ' + filtroCodClienteOrigem + (lojaOrigLabel ? ' / Loja: ' + lojaOrigLabel : ''));
                    }
                    if (filtroCodClienteDestino) {
                        var lojaDestLabel = '<%=filtroCodLojaDestino%>';
                        filtrosAplicados.push('Cod. Cliente Destino: ' + filtroCodClienteDestino + (lojaDestLabel ? ' / Loja: ' + lojaDestLabel : ''));
                    }
                    if (filtroIdTransferencia) {
                        filtrosAplicados.push('Id Transferência: ' + filtroIdTransferencia);
                    }
                    if (filtroStatusTransf && filtroStatusTransf !== 'TODOS') {
                        filtrosAplicados.push('Status: ' + filtroStatusTransf);
                    }

                    if (filtrosAplicados.length > 0) {
                        doc.text('Filtros Aplicados:', 14, yPosition);
                        yPosition += 5;
                        doc.setFont(undefined, 'normal');
                        for (var i = 0; i < filtrosAplicados.length; i++) {
                            doc.text('  • ' + filtrosAplicados[i], 14, yPosition);
                            yPosition += 4;
                        }
                        yPosition += 2;
                    }
                }
                
                for (let i = 0; i < linhas.length; i++) {
                    const celulas = linhas[i].getElementsByTagName('td');
                    const linha = [];
                    for (let j = 0; j < celulas.length - 3; j++) {
                        linha.push(celulas[j].innerText);
                    }
                    dados.push(linha);
                }

                yPosition += 3;
                doc.setFontSize(10);
                doc.setFont(undefined, 'normal');
                doc.text('Data de geração: ' + dataAtual + ' | Total de registros: ' + dados.length, 14, yPosition);
                yPosition += 3;
                
                doc.autoTable({
                    head: [colunas],
                    body: dados,
                    startY: yPosition,
                    theme: 'striped',
                    styles: {
                        fontSize: 8,
                        cellPadding: 2,
                        lineWidth: 0.1, 
                        lineColor: [0, 0, 0, 0.4],
                    },
                    headStyles: {
                        fillColor: [211, 211, 212],
                        textColor: [0, 0, 0],
                        fontStyle: 'bold',
                        halign: 'center',
                        lineWidth: 0.1, 
                        lineColor: [0, 0, 0, 0.4],
                    },
                    margin: { top: yPosition }
                });
                              
                window.open(doc.output('bloburl'), '_blank');
            }

            function exportarCSV() {
                var aba = '<%=abaSelecionada%>';
            
                let tabelaId = '';
                let nomeArquivo = '';
                let colunas = [];
                   
                if (aba === 'solicitacoes') {
                    tabelaId = 'GradeSolicitacoes';
                    nomeArquivo = 'Solicitacoes_credito.csv';
                    colunas = ['#', 'Planta', 'Solicitante', 'Cód. Cliente', 'Nome Cliente', 'CNPJ', 'Limite Atual', 'Limite Solicitado', 'Limite Validado', 'Último Validador', 'Status', 'Indicador Ano', 'Data de Solicitação', 'Data de Aprovação'];
                } else {
                    tabelaId = 'GradeTransferencias';
                    nomeArquivo = 'Transferencias_credito.csv';
                    colunas = ['#', 'Solicitante', 'Aprovador', 'Cód. Cliente', 'Nome Cliente', 'CNPJ', 'Planta Origem', 'Planta Destino', 'Limite Anterior Origem', 'Limite Anterior Destino', 'Status', 'Data Solicitação', 'Data Aprovação'];
                }

                let csvContent = '\uFEFF'; 
                csvContent += colunas.join(';') + '\r\n';

                const tabela = document.getElementById(tabelaId);
                const linhas = tabela.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
                
                for (let i = 0; i < linhas.length; i++) {
                    const celulas = linhas[i].getElementsByTagName('td');
                    const linhaDados = [];
                    
                    for (let j = 0; j < celulas.length - 3; j++) {
                        let texto = celulas[j].innerText.trim();
                        texto = texto;
                        linhaDados.push(texto);
                    }
                    
                    csvContent += linhaDados.join(';') + '\r\n';
                }
                
                const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                const link = document.createElement('a');
                
                if (navigator.msSaveBlob) { 
                    navigator.msSaveBlob(blob, nomeArquivo);
                } else {
                    const url = URL.createObjectURL(blob);
                    link.href = url;
                    link.download = nomeArquivo;
                    link.style.display = 'none';
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    URL.revokeObjectURL(url);
                }
            }
        </script>
        <style type="text/css">
            body {
                background-color: #f5f5f5;
                padding-bottom: 20px;
                font-size: .9rem;
            }
            .card {
                padding: 0rem;
                margin-bottom: 20px;
            }
            .container {
                max-width: 100% !important;
            }
            .form-label {
                font-weight: 600;
            }
            .filters-card {
                background-color: white;
                border-radius: 8px;
                padding: 15px;
                margin-bottom: 20px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            }
            .table-wrapper {
                overflow: auto;
                max-height: 70vh;
            }
            table {
                width: max-content;
                min-width: 100%;
            }
            thead {
                vertical-align: middle !important;
                position: sticky; 
                inset-block-start: 0; 
                background: #f5f5f5;
            }
            
            td, th {
                 white-space: nowrap;
                vertical-align: top;
                padding: 8px 12px !important;
            }
            th {
                font-weight: 600;
                color: #495057;
            }
            tr:hover {
                background: #f5f5f5;
            }
            .no-records {
                padding: 10px;
                text-align: center;
                background: white;
            }
            .status-badge {
                display: inline-block;
                text-align: center;
                padding: 6px 12px;
                border-radius: 6px;
                font-size: 0.75rem;
                font-weight: 600;
            }
            .badge-novo {
                background-color: #dbeafe;
                color: #1e40af;
            }
            .badge-rejeitado {
                background-color: #fee2e2;
                color: #991b1b;
            }
            .badge-aprovado {
                background-color: #d1fae5;
                color: #065f46;
            }        
            .badge-validado-analista {
                background-color: #e9d5ff;
                color: #6b21a8;
            }
            .badge-analise-finalizada {
                background-color: #fed7aa;
                color: #c2410c;
            }
            .badge-validado-coordenador {
                background-color: #cffafe;
                color: #155e75;
            }
            .badge-validado-gerente-tesouraria {
                background-color: #e0e7ff;
                color: #3730a3;
            }
            .badge-validado-controller {
                background-color: #fce7f3;
                color: #9f1239;
            }
            .badge-validado-gerente-planta {
                background-color: #fef3c7;
                color: #92400e;
            }
            .badge-default {
                background-color: #f3f4f6;
                color: #374151;
            }
            .badge-pendente {
                background-color: #fff3cd;
                color: #856404;
            }
            .tabs-container {
                display: flex;
                padding: 0;
                gap: 0;
                min-width: max-content;
                background-color: #f8f9fa;
                border-radius: 8px 8px 0 0;
            }
            .tab-button {
                flex: 0 0 auto; 
                padding: 15px 20px;
                background: transparent;
                border: none;
                border-bottom: 3px solid transparent;
                cursor: pointer;
                display: flex;
                align-items: center;
                justify-content: baseline;
                gap: 8px;
                color: #6c757d;
                font-weight: 500;
            }
            .tab-button:hover {
                color: #212529;
            }
            .tab-button.active {
                color: #0d6efd;
                border-bottom-color: #0d6efd;
                background-color: #fff;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h3 class="mt-4 mb-4">Histórico de Solicitações e Transferências</h3>
            
            <div class="tabs-container">
                <button type="button" class="tab-button <%=IIf(abaSelecionada = "solicitacoes", "active", "")%>" onclick="alterarAba('solicitacoes')">
                    <i class="fa fa-list me-2"></i>
                    <span class="tab-title">Solicitações de Crédito</span>
                </button>
                <button type="button" class="tab-button <%=IIf(abaSelecionada = "transferencias", "active", "")%>" onclick="alterarAba('transferencias')">
                    <i class="fas fa-exchange-alt me-2"></i>
                    <span class="tab-title">Transferências de Crédito</span>
                </button>
            </div>
            
            <% If abaSelecionada = "solicitacoes" Then %>
            <div class="filters-card">
                <form method="post" action="?aba=solicitacoes" onsubmit="return validarFiltros();">
                    <div class="row">
                        <div class="col-md-3 mb-3">
                            <label for="filtroDataSolicitacaoDe" class="form-label">Data de Solicitação (De)</label>
                            <input type="date" class="form-control form-control-sm" id="filtroDataSolicitacaoDe" name="filtroDataSolicitacaoDe" value="<%=filtroDataSolicitacaoDe%>">
                        </div>
                        <div class="col-md-3 mb-3">
                            <label for="filtroDataSolicitacaoAte" class="form-label">Data de Solicitação (Até)</label>
                            <input type="date" class="form-control form-control-sm" id="filtroDataSolicitacaoAte" name="filtroDataSolicitacaoAte" value="<%=filtroDataSolicitacaoAte%>">
                        </div>
                        <div class="col-md-3 mb-3">
                            <label for="filtroDataAprovacaoDe" class="form-label">Data de Aprovação (De)</label>
                            <input type="date" class="form-control form-control-sm" id="filtroDataAprovacaoDe" name="filtroDataAprovacaoDe" value="<%=filtroDataAprovacaoDe%>">
                        </div>
                        <div class="col-md-3 mb-3">
                            <label for="filtroDataAprovacaoAte" class="form-label">Data de Aprovação (Até)</label>
                            <input type="date" class="form-control form-control-sm" id="filtroDataAprovacaoAte" name="filtroDataAprovacaoAte" value="<%=filtroDataAprovacaoAte%>">
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-md-3 mb-3">
                            <label for="filtroPlanta" class="form-label">Planta</label>
                            <select class="form-select form-select-sm" id="filtroPlanta" name="filtroPlanta">
                                <option value="TODOS" <%=IIf(filtroPlanta = "TODOS" Or filtroPlanta = "", "selected", "")%>>Todas</option>
                                <option value="00" <%=IIf(filtroPlanta = "00", "selected", "")%>>Pirapetinga</option>
                                <option value="04" <%=IIf(filtroPlanta = "04", "selected", "")%>>Uberaba</option>
                                <option value="05" <%=IIf(filtroPlanta = "05", "selected", "")%>>Bento</option>
                                <option value="06" <%=IIf(filtroPlanta = "06", "selected", "")%>>Maranguape</option>
                                <option value="07" <%=IIf(filtroPlanta = "07", "selected", "")%>>Saquarema</option>
                                <option value="08" <%=IIf(filtroPlanta = "08", "selected", "")%>>Aracatuba</option>
                                <option value="09" <%=IIf(filtroPlanta = "09", "selected", "")%>>Blumenau</option>
                                <option value="10" <%=IIf(filtroPlanta = "10", "selected", "")%>>Pacajus</option>
                                <option value="11" <%=IIf(filtroPlanta = "11", "selected", "")%>>Porto Feliz</option>
                                <option value="12" <%=IIf(filtroPlanta = "12", "selected", "")%>>Tres Barras</option>
                                <option value="13" <%=IIf(filtroPlanta = "13", "selected", "")%>>Florestal</option>
                            </select>
                        </div>
                        <div class="col-md-3 mb-3">
                            <label for="filtroStatus" class="form-label">Status</label>
                            <select class="form-select form-select-sm" id="filtroStatus" name="filtroStatus">
                                <option value="TODOS" <%=IIf(filtroStatus = "TODOS" Or filtroStatus = "", "selected", "")%>>Todos</option>
                                <option value="A" <%=IIf(filtroStatus = "A", "selected", "")%>>Novo</option>
                                <option value="D" <%=IIf(filtroStatus = "D", "selected", "")%>>Análise Finalizada</option>
                                <option value="C" <%=IIf(filtroStatus = "C", "selected", "")%>>Validado Analista</option>
                                <option value="E" <%=IIf(filtroStatus = "E", "selected", "")%>>Validado Coordenador</option>
                                <option value="F" <%=IIf(filtroStatus = "F", "selected", "")%>>Validado Gerente de Tesouraria</option>
                                <option value="G" <%=IIf(filtroStatus = "G", "selected", "")%>>Validado Controller</option>
                                <option value="H" <%=IIf(filtroStatus = "H", "selected", "")%>>Validado Gerente Planta</option>
                                <option value="V" <%=IIf(filtroStatus = "V", "selected", "")%>>Aprovado</option>
                                <option value="B" <%=IIf(filtroStatus = "B", "selected", "")%>>Rejeitado</option>
                            </select>
                        </div>
                        <div class="col-md-3 mb-3">
                            <label class="form-label">Cod. Cliente</label>
                            <div class="input-group input-group-sm">
                                <button type="button"
                                        class="btn btn-secondary btn-sm"
                                        title="Buscar cliente"
                                        onclick="abrirBuscaCliente('filtroCodCliente','filtroCodLoja', document.getElementById('filtroPlanta').value)">
                                    Buscar &nbsp;<i class="fas fa-external-link-alt"></i>
                                </button>
                                <input type="text"
                                       class="form-control form-control-sm"
                                       id="filtroCodCliente"
                                       name="filtroCodCliente"
                                       value="<%=filtroCodCliente%>"
                                       placeholder="Código"
                                       autocomplete="off" />
                                <input type="text"
                                       class="form-control form-control-sm"
                                       id="filtroCodLoja"
                                       name="filtroCodLoja"
                                       value="<%=filtroCodLoja%>"
                                       placeholder="Loja"
                                       autocomplete="off"
                                       style="max-width:60px;" />
                            </div>
                        </div>
                        <div class="col-md-3 mb-3">
                            <label for="filtroIdSolicitacao" class="form-label">Id Solicitação</label>
                            <input type="text" class="form-control form-control-sm" id="filtroIdSolicitacao" name="filtroIdSolicitacao" value="<%=filtroIdSolicitacao%>">
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-12 text-end">
                            <% IF totalSolicitacoes > 0 THEN %>
                            <button type="button" class="btn btn-danger btn-sm" onclick="exportarPDF()">
                                <i class="fas fa-file-pdf me-2"></i> Exportar PDF
                            </button>
                            <button type="button" class="btn btn-success btn-sm" onclick="exportarCSV()">
                                <i class="fas fa-file-excel me-2"></i>Exportar Excel
                            </button>
                            <% END IF %>
                            <button type="button" class="btn btn-sm btn-secondary" onclick="limparFiltros()">
                                <i class="fas fa-eraser me-2"></i> Limpar Filtros
                            </button>
                            <button type="submit" class="btn btn-sm btn-primary">
                                <i class="fas fa-search me-2"></i> Pesquisar
                            </button>
                        </div>
                    </div>
                </form>
            </div>
            <% End If %>
            
            <% If abaSelecionada = "transferencias" Then %>
            <div class="filters-card">
                <form method="post" action="?aba=transferencias" onsubmit="return validarFiltros();">
                    <div class="row">
                        <div class="col-md-3 mb-3">
                            <label for="filtroDataSolicitacaoDeTransf" class="form-label">Data de Solicitação (De)</label>
                            <input type="date" class="form-control form-control-sm" id="filtroDataSolicitacaoDeTransf" name="filtroDataSolicitacaoDeTransf" value="<%=filtroDataSolicitacaoDeTransf%>">
                        </div>
                        <div class="col-md-3 mb-3">
                            <label for="filtroDataSolicitacaoAteTransf" class="form-label">Data de Solicitação (Até)</label>
                            <input type="date" class="form-control form-control-sm" id="filtroDataSolicitacaoAteTransf" name="filtroDataSolicitacaoAteTransf" value="<%=filtroDataSolicitacaoAteTransf%>">
                        </div>
                        <div class="col-md-3 mb-3">
                            <label for="filtroDataAprovacaoDeTransf" class="form-label">Data de Aprovação (De)</label>
                            <input type="date" class="form-control form-control-sm" id="filtroDataAprovacaoDeTransf" name="filtroDataAprovacaoDeTransf" value="<%=filtroDataAprovacaoDeTransf%>">
                        </div>
                        <div class="col-md-3 mb-3">
                            <label for="filtroDataAprovacaoAteTransf" class="form-label">Data de Aprovação (Até)</label>
                            <input type="date" class="form-control form-control-sm" id="filtroDataAprovacaoAteTransf" name="filtroDataAprovacaoAteTransf" value="<%=filtroDataAprovacaoAteTransf%>">
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-md-3 mb-3">
                            <label for="filtroPlantaOrigem" class="form-label">Planta Origem</label>
                            <select class="form-select form-select-sm" id="filtroPlantaOrigem" name="filtroPlantaOrigem">
                                <option value="TODOS" <%=IIf(filtroPlantaOrigem = "TODOS" Or filtroPlantaOrigem = "", "selected", "")%>>Todas</option>
                                <option value="00" <%=IIf(filtroPlantaOrigem = "00", "selected", "")%>>Pirapetinga</option>
                                <option value="04" <%=IIf(filtroPlantaOrigem = "04", "selected", "")%>>Uberaba</option>
                                <option value="05" <%=IIf(filtroPlantaOrigem = "05", "selected", "")%>>Bento</option>
                                <option value="06" <%=IIf(filtroPlantaOrigem = "06", "selected", "")%>>Maranguape</option>
                                <option value="07" <%=IIf(filtroPlantaOrigem = "07", "selected", "")%>>Saquarema</option>
                                <option value="08" <%=IIf(filtroPlantaOrigem = "08", "selected", "")%>>Aracatuba</option>
                                <option value="09" <%=IIf(filtroPlantaOrigem = "09", "selected", "")%>>Blumenau</option>
                                <option value="10" <%=IIf(filtroPlantaOrigem = "10", "selected", "")%>>Pacajus</option>
                                <option value="11" <%=IIf(filtroPlantaOrigem = "11", "selected", "")%>>Porto Feliz</option>
                                <option value="12" <%=IIf(filtroPlantaOrigem = "12", "selected", "")%>>Tres Barras</option>
                                <option value="13" <%=IIf(filtroPlantaOrigem = "13", "selected", "")%>>Florestal</option>
                            </select>
                        </div>
                        <div class="col-md-3 mb-3">
                            <label for="filtroPlantaDestino" class="form-label">Planta Destino</label>
                            <select class="form-select form-select-sm" id="filtroPlantaDestino" name="filtroPlantaDestino">
                                <option value="TODOS" <%=IIf(filtroPlantaDestino = "TODOS" Or filtroPlantaDestino = "", "selected", "")%>>Todas</option>
                                <option value="00" <%=IIf(filtroPlantaDestino = "00", "selected", "")%>>Pirapetinga</option>
                                <option value="04" <%=IIf(filtroPlantaDestino = "04", "selected", "")%>>Uberaba</option>
                                <option value="05" <%=IIf(filtroPlantaDestino = "05", "selected", "")%>>Bento</option>
                                <option value="06" <%=IIf(filtroPlantaDestino = "06", "selected", "")%>>Maranguape</option>
                                <option value="07" <%=IIf(filtroPlantaDestino = "07", "selected", "")%>>Saquarema</option>
                                <option value="08" <%=IIf(filtroPlantaDestino = "08", "selected", "")%>>Aracatuba</option>
                                <option value="09" <%=IIf(filtroPlantaDestino = "09", "selected", "")%>>Blumenau</option>
                                <option value="10" <%=IIf(filtroPlantaDestino = "10", "selected", "")%>>Pacajus</option>
                                <option value="11" <%=IIf(filtroPlantaDestino = "11", "selected", "")%>>Porto Feliz</option>
                                <option value="12" <%=IIf(filtroPlantaDestino = "12", "selected", "")%>>Tres Barras</option>
                                <option value="13" <%=IIf(filtroPlantaDestino = "13", "selected", "")%>>Florestal</option>
                            </select>
                        </div>
                        <div class="col-md-3 mb-3">
                            <label for="filtroStatusTransf" class="form-label">Status</label>
                            <select class="form-select form-select-sm" id="filtroStatusTransf" name="filtroStatusTransf">
                                <option value="TODOS" <%=IIf(filtroStatusTransf = "TODOS" Or filtroStatusTransf = "", "selected", "")%>>Todos</option>
                                <option value="PENDENTE" <%=IIf(filtroStatusTransf = "PENDENTE", "selected", "")%>>Pendente</option>
                                <option value="APROVADO" <%=IIf(filtroStatusTransf = "APROVADO", "selected", "")%>>Aprovado</option>
                                <option value="REJEITADO" <%=IIf(filtroStatusTransf = "REJEITADO", "selected", "")%>>Rejeitado</option>
                            </select>
                        </div>
                        <div class="col-md-3 mb-3">
                            <label class="form-label">Cód. Cliente Origem</label>
                            <div class="input-group input-group-sm">
                                <button type="button"
                                        class="btn btn-secondary btn-sm"
                                        title="Buscar cliente"
                                        onclick="abrirBuscaCliente('filtroCodClienteOrigem','filtroCodLojaOrigem', document.getElementById('filtroPlantaOrigem').value)">
                                    Buscar &nbsp;<i class="fas fa-external-link-alt"></i>
                                </button>
                                <input type="text"
                                       class="form-control form-control-sm"
                                       id="filtroCodClienteOrigem"
                                       name="filtroCodClienteOrigem"
                                       value="<%=filtroCodClienteOrigem%>"
                                       placeholder="Código"
                                       autocomplete="off" />
                                <input type="text"
                                       class="form-control form-control-sm"
                                       id="filtroCodLojaOrigem"
                                       name="filtroCodLojaOrigem"
                                       value="<%=filtroCodLojaOrigem%>"
                                       placeholder="Loja"
                                       autocomplete="off"
                                       style="max-width:60px;" />
                            </div>
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-md-3 mb-3">
                            <label class="form-label">Cód. Cliente Destino</label>
                            <div class="input-group input-group-sm">
                                <button type="button"
                                        class="btn btn-secondary btn-sm"
                                        title="Buscar cliente"
                                        onclick="abrirBuscaCliente('filtroCodClienteDestino','filtroCodLojaDestino', document.getElementById('filtroPlantaDestino').value)">
                                    Buscar &nbsp;<i class="fas fa-external-link-alt"></i>
                                </button>
                                <input type="text"
                                       class="form-control form-control-sm"
                                       id="filtroCodClienteDestino"
                                       name="filtroCodClienteDestino"
                                       value="<%=filtroCodClienteDestino%>"
                                       placeholder="Código"
                                       autocomplete="off" />
                                <input type="text"
                                       class="form-control form-control-sm"
                                       id="filtroCodLojaDestino"
                                       name="filtroCodLojaDestino"
                                       value="<%=filtroCodLojaDestino%>"
                                       placeholder="Loja"
                                       autocomplete="off"
                                       style="max-width:60px;" />
                            </div>
                        </div>
                        <div class="col-md-3 mb-3">
                            <label for="filtroIdTransferencia" class="form-label">Id Transferência</label>
                            <input type="text" class="form-control form-control-sm" id="filtroIdTransferencia" name="filtroIdTransferencia" value="<%=filtroIdTransferencia%>">
                        </div>
                        <div class="col-md-6 mb-3 d-flex align-items-end">
                            <div class="d-flex gap-2">
                                <% IF totalTransferencias > 0 THEN %>
                                <button type="button" class="btn btn-danger btn-sm" onclick="exportarPDF()">
                                    <i class="fas fa-file-pdf me-2"></i> Exportar PDF
                                </button>
                                <button type="button" class="btn btn-success btn-sm" onclick="exportarCSV()">
                                    <i class="fas fa-file-excel me-2"></i> Exportar Excel
                                </button>
                                <% END IF %>
                                <button type="button" class="btn btn-sm btn-secondary" onclick="limparFiltros()">
                                    <i class="fas fa-eraser me-2"></i> Limpar Filtros
                                </button>
                                <button type="submit" class="btn btn-sm btn-primary">
                                    <i class="fas fa-search me-2"></i> Pesquisar
                                </button>
                            </div>
                        </div>
                    </div>
                </form>
            </div>
            <% End If %>

            <% If abaSelecionada = "solicitacoes" Then %>
            <div class="card">
                <div class="table-wrapper">
                    <table class="table" name="GradeSolicitacoes" id="GradeSolicitacoes">
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Planta</th>
                                <th>Solicitante</th>
                                <th>Cód. Cliente</th>
                                <th>Nome Cliente</th>
                                <th>CNPJ</th>
                                <th>Limite Atual</th>
                                <th>Limite Solicitado</th>
                                <th>Limite Validado</th>
                                <th>Último Validador</th>
                                <th>Status</th>
                                <th>Indicador Ano</th>
                                <th>Data de Solicitação</th>
                                <th>Data de Aprovação</th>
                                <th>Detalhes</th>
                                <th>Anexos</th>
                                <th>Histórico</th>
                            </tr>              
                        </thead>

                        <% IF totalSolicitacoes > 0 THEN %>
                        <tbody>
                            <% 
                            Dim i
                            For i = 0 To UBound(solicitacoes)
                                Dim temAnexos, temEnderecoLegado
                                temEnderecoLegado = (solicitacoes(i)("EnderecoDocumentacao") <> "" AND NOT IsNull(solicitacoes(i)("EnderecoDocumentacao")))
                                temAnexos = (solicitacoes(i)("QuantidadeAnexos") > 0) OR temEnderecoLegado
                            %>
                                <tr>   
                                    <td class="text-left"><%=solicitacoes(i)("IdSolicitacao")%></td>
                                    <td class="text-left"><%=ObterNomePlanta(solicitacoes(i)("CodPlanta"))%></td>
                                    <td class="text-left"><%=solicitacoes(i)("NomeSolicitante")%></td>
                                    <td class="text-left"><%=solicitacoes(i)("CodCliente")%></td>
                                    <td class="text-left"><%=solicitacoes(i)("NomeCliente")%></td>
                                    <td class="text-left"><%=FormataCnpj(solicitacoes(i)("CNPJ"))%></td>
                                    <td class="text-right"><%=FormatarMoeda(solicitacoes(i)("LimiteAtual"))%></td>
                                    <td class="text-right"><%=FormatarMoeda(solicitacoes(i)("LimiteSolicitado"))%></td>
                                    <td class="text-right"><%=FormatarMoeda(solicitacoes(i)("LimiteValidado"))%></td>
                                    <td class="text-left">
                                        <% If solicitacoes(i)("NomeValidador") <> "" Then %>
                                            <span><%=solicitacoes(i)("NomeValidador")%></span>
                                        <% Else %>
                                            <span>-</span>
                                        <% End If %>
                                    </td>
                                    <td class="text-left">
                                        <span class="status-badge <%=ObterClasseStatus(solicitacoes(i)("Status"))%>">
                                            <%=ObterNomeStatus(solicitacoes(i)("Status"))%>
                                        </span>
                                    </td>
                                    <td class="text-left"><%=solicitacoes(i)("BalancoAno")%></td>
                                    <td class="text-left"><%=FormatarData(solicitacoes(i)("DataCadastro"))%></td>
                                    <td class="text-left"><%=FormatarData(solicitacoes(i)("DataValidacao"))%></td>
                                    <td class="text-center">
                                        <button type="button" class="btn btn-sm btn-info" onclick="abrirModalDetalhes(<%=solicitacoes(i)("IdSolicitacao")%>)" title="Ver Detalhes">
                                            <i class="fas fa-file-alt"></i>
                                        </button>
                                    </td>
                                    <td class="text-center">
                                        <% If temAnexos Then %>
                                            <button type="button" class="btn btn-sm btn-primary" onclick="abrirModalAnexos(<%=solicitacoes(i)("IdSolicitacao")%>, '<%=solicitacoes(i)("EnderecoDocumentacao")%>')" title="Ver Anexos">
                                                <i class="fas fa-paperclip"></i>
                                            </button>
                                        <% Else %>
                                            <span class="text-muted">-</span>
                                        <% End If %>
                                    </td>
                                    <td class="text-center">
                                        <button type="button" class="btn btn-sm btn-secondary btn-outlined" onclick="abrirHistoricoValidacoes(<%=solicitacoes(i)("IdSolicitacao")%>)" title="Ver Histórico de Etapas">
                                            <i class="fas fa-history"></i>
                                        </button>
                                    </td>
                                </tr>
                            <% Next %>
                        </tbody>
                        <% END IF %>
                    </table>
                    <% IF totalSolicitacoes = 0 AND possuiFiltroData THEN %>
                        <div class="no-records">
                            Nenhum registro encontrado
                        </div>
                    <% END IF %>
                </div>
            </div>
            <% End If %>

            <% If abaSelecionada = "transferencias" Then %>
            <div class="card">
                <div class="table-wrapper">
                    <table class="table" name="GradeTransferencias" id="GradeTransferencias">
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Solicitante</th>
                                <th>Aprovador</th>
                                <th>Cód. Cliente</th>
                                <th>Nome Cliente</th>
                                <th>CNPJ</th>
                                <th>Planta Origem</th>
                                <th>Planta Destino</th>
                                <th>Limite Anterior Origem</th>
                                <th>Limite Anterior Destino</th>
                                <th>Status</th>
                                <th>Data de Solicitação</th>
                                <th>Data de Aprovação</th>
                            </tr>              
                        </thead>

                        <% IF totalTransferencias > 0 THEN %>
                        <tbody>
                            <% 
                            Dim t
                            For t = 0 To UBound(transferencias)
                                Dim classeStatusTransf
                                Select Case transferencias(t)("Status")
                                    Case "PENDENTE"
                                        classeStatusTransf = "badge-pendente"
                                    Case "APROVADO"
                                        classeStatusTransf = "badge-aprovado"
                                    Case "REJEITADO"
                                        classeStatusTransf = "badge-rejeitado"
                                    Case Else
                                        classeStatusTransf = "badge-default"
                                End Select
                            %>
                                <tr>   
                                    <td class="text-left"><%=transferencias(t)("Id")%></td>
                                    <td class="text-left"><%=transferencias(t)("NomeSolicitante")%></td>
                                    <td class="text-left">
                                        <%= Iif(transferencias(t)("NomeAprovador") <> "", transferencias(t)("NomeAprovador"), "-") %>
                                    </td>
                                    <td class="text-left"><%=transferencias(t)("CodClienteOrigem")%></td>
                                    <td class="text-left"><%=transferencias(t)("NomeCliente")%></td>
                                    <td class="text-left"><%=FormataCnpj(transferencias(t)("CNPJ"))%></td>
                                    <td class="text-left"><%=ObterNomePlanta(transferencias(t)("PlantaOrigem"))%></td>
                                    <td class="text-left"><%=ObterNomePlanta(transferencias(t)("PlantaDestino"))%></td>
                                    <td class="text-right"><%=FormatarMoeda(transferencias(t)("LimiteAtualOrigem"))%></td>
                                    <td class="text-right"><%=FormatarMoeda(transferencias(t)("LimiteAtualDestino"))%></td>
                                    <td class="text-left">
                                        <span class="status-badge <%=classeStatusTransf%>">
                                            <%=transferencias(t)("Status")%>
                                        </span>
                                    </td>
                                    <td class="text-left"><%=FormatarData(transferencias(t)("DataCriacao"))%></td>
                                    <td class="text-left"><%=FormatarData(transferencias(t)("DataValidacao"))%></td>
                                </tr>
                            <% Next %>
                        </tbody>
                        <% END IF %>
                    </table>
                    <% IF totalTransferencias = 0 AND possuiFiltroDataTransf THEN %>
                        <div class="no-records">
                            Nenhum registro encontrado
                        </div>
                    <% END IF %>
                </div>
            </div>
            <% End If %>
        </div>

        <div class="modal fade" id="modalHistorico" tabindex="-1" aria-labelledby="modalHistoricoLabel" aria-hidden="true">
            <div class="modal-dialog modal-lg">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="modalHistoricoLabel">
                            <i class="fas fa-history me-2"></i>Histórico
                        </h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body" id="modalHistoricoBody"></div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Fechar</button>
                    </div>
                </div>
            </div>
        </div>

        <div class="modal fade" id="modalAnexos" tabindex="-1" aria-labelledby="modalAnexosLabel" aria-hidden="true">
            <div class="modal-dialog modal-lg">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="modalAnexosLabel">
                            <i class="fas fa-paperclip me-2"></i>Anexos
                        </h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body" id="modalAnexosBody"></div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Fechar</button>
                    </div>
                </div>
            </div>
        </div>

        <div class="modal fade" id="modalDetalhes" tabindex="-1" aria-labelledby="modalDetalhesLabel" aria-hidden="true">
            <div class="modal-dialog modal-xl">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title" id="modalDetalhesLabel">Detalhes da Solicitação</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body" id="modalDetalhesBody"></div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary btn-sm" data-bs-dismiss="modal">Fechar</button>
                    </div>
                </div>
            </div>
        </div>
    </body>
</html>
