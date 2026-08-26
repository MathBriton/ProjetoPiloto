<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../../cnn2/cnn2.asp" -->
<!--#include File = "../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp" -->
<!--#include file="../util/Util.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046

Dim idTransferencia, flagRejeitar, idUsuarioLogado
idTransferencia = Request.QueryString("idTransferencia")
flagRejeitar = Request.QueryString("flagRejeitar")
idUsuarioLogado = Session("IdUsuario")

If idTransferencia = "" Or Not IsNumeric(idTransferencia) Then
    Response.Redirect "TransferenciaCreditoGrade.asp"
    Response.End
End If

If flagRejeitar = "" Then flagRejeitar = "0"

Dim SQL, PIR_CORPORATIVO, codNivel, validadorEncontrado
validadorEncontrado = False

SQL = "SELECT DISTINCT "
SQL = SQL & " CAST(vc.IdValidador AS INT) AS IdValidador,"
SQL = SQL & " u.CRM_NOME AS Nome,"
SQL = SQL & " CAST(vc.Nivel AS NVARCHAR(10)) as CodNivel"
SQL = SQL & " FROM ValidadoresCredito vc"
SQL = SQL & " LEFT JOIN HELPDESK..CRM_DTI u ON vc.IdValidador = u.CRM_R_E_C_N_O"
SQL = SQL & " WHERE vc.IdValidador = " & idUsuarioLogado
SQL = SQL & " GROUP BY vc.IdValidador, u.CRM_NOME, vc.Nivel"
SQL = SQL & " ORDER BY Nome"

On Error Resume Next
SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
    If Not PIR_CORPORATIVO.EOF Then
        codNivel = PIR_CORPORATIVO("CodNivel")
        validadorEncontrado = True
    End If
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
End If
On Error Goto 0

If Not validadorEncontrado OR codNivel <> "10" Then
    Response.write("<span>Você não tem acesso a essa página.</span>")
    Response.End()
End If

Dim TRANSFERENCIA_CONEXAO
SQL = "SELECT DISTINCT " 
SQL = SQL & "t.Id as IdTransferencia, "
SQL = SQL & "t.PlantaOrigem, "
SQL = SQL & "t.CodClienteOrigem, "
SQL = SQL & "t.CodLojaOrigem, "
SQL = SQL & "t.PlantaDestino, "
SQL = SQL & "t.CodClienteDestino, "
SQL = SQL & "t.CodLojaDestino, "
SQL = SQL & "t.ValorTransferencia, "
SQL = SQL & "t.Status, "
SQL = SQL & "t.Observacao, "
SQL = SQL & "CASE" 
SQL = SQL & "    WHEN ISDATE(t.DataCriacao) = 1"
SQL = SQL & "    THEN CONVERT(VARCHAR(10), CONVERT(DATE, t.DataCriacao, 112), 126) "
SQL = SQL & "    ELSE NULL "
SQL = SQL & "END AS DataCriacao "
SQL = SQL & "FROM TransferenciaCredito t " 
SQL = SQL & "WHERE t.Id = " & idTransferencia

On Error Resume Next
SET TRANSFERENCIA_CONEXAO = Conecta("PIR-CORPORATIVO", SQL)
If Err.Number <> 0 OR NOT IsObject(TRANSFERENCIA_CONEXAO) OR TRANSFERENCIA_CONEXAO.EOF Then
    Response.Redirect "TransferenciaCreditoGrade.asp"
    Response.End
End If
On Error Goto 0

Dim plantaOrigem, codClienteOrigem, codLojaOrigem
Dim plantaDestino, codClienteDestino, codLojaDestino
Dim valorTransferencia, statusTransferencia, observacao

plantaOrigem = TRANSFERENCIA_CONEXAO("PlantaOrigem") & ""
codClienteOrigem = TRANSFERENCIA_CONEXAO("CodClienteOrigem") & ""
codLojaOrigem = TRANSFERENCIA_CONEXAO("CodLojaOrigem") & ""
plantaDestino = TRANSFERENCIA_CONEXAO("PlantaDestino") & ""
codClienteDestino = TRANSFERENCIA_CONEXAO("CodClienteDestino") & ""
codLojaDestino = TRANSFERENCIA_CONEXAO("CodLojaDestino") & ""
valorTransferencia = TRANSFERENCIA_CONEXAO("ValorTransferencia")
statusTransferencia = TRANSFERENCIA_CONEXAO("Status") & ""
observacao = TRANSFERENCIA_CONEXAO("Observacao") & ""

If IsObject(TRANSFERENCIA_CONEXAO) Then
    TRANSFERENCIA_CONEXAO.Close()
    SET TRANSFERENCIA_CONEXAO = NOTHING
End If

Function ObterTituloTela()
    If flagRejeitar = "1" Then
        ObterTituloTela = "Rejeitar Transferência de Crédito"
    Else
        ObterTituloTela = "Aprovar Transferência de Crédito"
    End If
    ObterTituloTela = ObterTituloTela & " #" & idTransferencia
End Function

Function ObterTextoBotao()
    If flagRejeitar = "1" Then
        ObterTextoBotao = "Rejeitar Transferência"
    Else
        ObterTextoBotao = "Aprovar Transferência"
    End If
End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Smurfit Kappa - <%= ObterTituloTela() %></title>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" />
    <link href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" rel="stylesheet" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css"> 
    <style type="text/css">
        body {
            background-color: #f5f5f5;
            font-size: .9rem;
        }
        .btn-back {
            display: flex;
            justify-content: center;
            align-items: center;
            cursor: pointer;
            background: transparent;
            height: 35px;
            width: 35px;
            border-radius: 50%;
            border: none;
            margin-right: 15px;
            font-size: 1.4rem;
        }
        .btn-manual {
            display: flex;
            justify-content: center;
            align-items: center;
            color: #fff;
            background: #0d6efd;
            border: 1px solid #0d6efd;
            width: 35px;
            height: 35px;
            border-radius: 7px;
            margin-left: 12px;
            font-size: 1.2rem;
        }
        .card {
            padding: 1rem;
        }
        .form-label {
            font-weight: 600;
        }
        .required-field {
            color: #dc3545;
        }
        .transfer-container {
            display: flex;
            gap: 2rem;
            align-items: stretch;
        }
        .transfer-card {
            flex: 1;
            background: #fafafa;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: .7rem;
        }
        .transfer-header {
            display: flex;
            justify-content: space-between;
            align-items: baseline;
            margin-bottom: 1rem;
            padding-bottom: 0.75rem;
            border-bottom: 1px solid #e0e0e0;
        }
        .transfer-label {
            font-size: 0.75rem;
            color: #757575;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .transfer-title {
            font-size: 1rem;
            font-weight: 600;
            color: #212121;
        }
        .transfer-details {
            display: flex;
            flex-direction: column;
            gap: 0.4rem;
        }
        .transfer-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .transfer-row-label {
            font-size: 0.875rem;
            color: #616161;
        }
        .transfer-row-value {
            font-size: 0.875rem;
            font-weight: 600;
            color: #212121;
        }
        .transfer-row.highlight {
            padding-top: 0.75rem;
            border-top: 1px solid #e0e0e0;
        }
        .transfer-row.highlight .transfer-row-value {
            font-size: 1rem;
            color: #198754;
        }
        .transfer-arrow-container {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 0.5rem;
            min-width: 80px;
        }
        .transfer-arrow {
            width: 32px;
            height: 32px;
            border: 1px solid #bdbdbd;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            background: white;
        }
        .transfer-arrow i {
            color: #757575;
            font-size: 0.875rem;
        }
        @media (max-width: 991px) {
            .transfer-container {
                flex-direction: column;
            }
            .transfer-arrow-container {
                flex-direction: row;
                min-width: auto;
                padding: 1rem 0;
            }
            .transfer-arrow {
                transform: rotate(90deg);
            }
        }
        .divider {
            height: 1px;
            background-color: #dee2e6;
            margin: 20px 0;
        }
        .validation-message {
            background-color: #cfe2ff;
            border: 1px solid #0d6efd;
            color: #084298;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            margin: 40px auto;
            max-width: 600px;
        }
        .validation-message i {
            font-size: 3rem;
            margin-bottom: 15px;
        }
        .validation-message h4 {
            margin-bottom: 10px;
        }
        .validation-message.finished {
            background-color: #d1e7dd;
            border: 1px solid #198754;
            color: #0f5132;
        }
        .info-card {
            background-color: #f8f9fa;
            padding: 1rem;
            border-radius: 0.25rem;
            margin-bottom: 1rem;
            border: 1px solid #e8e9ed;
        }
        .info-card h6 {
            font-size: 0.875rem;
            font-weight: 600;
            color: #495057;
            margin-bottom: 0.5rem;
        }
        .arrow-icon {
            width: 32px;
            height: 32px;
            background-color: white;
            border: 1px solid #667eea;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .arrow-icon i {
            color: #667eea;
            font-size: 1rem;
            font-weight: 700;
        }
    </style>
    
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jquery-maskmoney@3.0.2/dist/jquery.maskMoney.min.js"></script>        
    <script src="../upload/fileUpload.js?v=<%= Now() %>"></script>
    <script src="../util/Util.js?v=<%= Now() %>"></script>

    <script type="text/javascript">
        const flagRejeitar = <%= flagRejeitar %>;
        const idTransferencia = <%= idTransferencia %>;
        const valorTransferenciaOriginal = parseFloat(<%= Replace(CStr(valorTransferencia), ",", ".") %>);
        
        const plantaOrigem = '<%= plantaOrigem %>';
        const codClienteOrigem = '<%= codClienteOrigem %>';
        const codLojaOrigem = '<%= codLojaOrigem %>';
        
        const plantaDestino = '<%= plantaDestino %>';
        const codClienteDestino = '<%= codClienteDestino %>';
        const codLojaDestino = '<%= codLojaDestino %>';

        let limiteAtualOrigem = 0;
        let limiteAtualDestino = 0;

        const PLANTAS = {
            '06': 'Maranguape', 
            '00': 'Pirapetinga',
            '04': 'Uberaba',  
            '05': 'Bento', 
            '07': 'Saquarema'
        };

        function atualizarVisualizacaoTransferencia() {
            const valorTransf = valorTransferenciaOriginal;
            
            $('#inputValorTransferencia').val(formatarMoeda(valorTransf));
            $('#previewOriginAfter').text(formatarMoeda(limiteAtualOrigem - valorTransf));
            $('#previewDestAfter').text(formatarMoeda(limiteAtualDestino + valorTransf));
        }

        function navegarParaGrade() {
            window.location.href = "../SolicitacaoCreditoGrade.asp?aba=transferencias"
        }

        function carregarDadosClienteOrigem() {
            if (!plantaOrigem || !codClienteOrigem) return;

            $.ajax({
                url: '../SolicitacaoCreditoRepositorio.asp?v=' + Math.random(),
                type: 'POST',
                data: {
                    acao: "obterCliente",
                    planta: plantaOrigem.trim(),
                    codCliente: codClienteOrigem.trim(),
                    codLoja: codLojaOrigem.trim()
                },
                dataType: 'json',
                success: function(response) {
                    if(response.sucesso) {
                        $('#inputNomeCliente').val(response.dados.nomeCliente);
                        $('#limiteAtualOrigem').text(formatarMoeda(response.dados.limite));
                        $('#saldoOrigem').text(formatarMoeda(response.dados.saldo));
                        $('#limiteAbertoOrigem').text(formatarMoeda(response.dados.aberto));
                        
                        limiteAtualOrigem = parseFloat(response.dados.limite) || 0;
                        atualizarVisualizacaoTransferencia();
                    } else {
                        console.log('Erro ao carregar dados do cliente de origem: ' + response.mensagem);
                    }
                },
                error: function(xhr, status, error) {
                    console.log('Erro ao obter dados do cliente de origem: ' + error);
                }
            });
        }

        function carregarDadosClienteDestino() {
            if (!plantaDestino || !codClienteDestino) return;

            $.ajax({
                url: '../SolicitacaoCreditoRepositorio.asp?v=' + Math.random(),
                type: 'POST',
                data: {
                    acao: "obterCliente",
                    planta: plantaDestino.trim(),
                    codCliente: codClienteDestino.trim(),
                    codLoja: codLojaDestino.trim()
                },
                dataType: 'json',
                success: function(response) {
                    if(response.sucesso) {
                        $('#limiteAtualDestino').text(formatarMoeda(response.dados.limite));
                        $('#saldoDestino').text(formatarMoeda(response.dados.saldo));
                        $('#limiteAbertoDestino').text(formatarMoeda(response.dados.aberto));
                        
                        limiteAtualDestino = parseFloat(response.dados.limite) || 0;
                        atualizarVisualizacaoTransferencia();
                    } else {
                        console.log('Erro ao carregar dados do cliente de destino: ' + response.mensagem);
                    }
                },
                error: function(xhr, status, error) {
                    console.log('Erro ao obter dados do cliente de destino: ' + error);
                }
            });
        }

        function salvarFormulario() {
            const $btnSalvar = $('#btnSalvar');
            const $iconCheck = $('#iconCheck');
            const $iconLoading = $('#iconLoading');
            
            $btnSalvar.prop('disabled', true);
            $iconCheck.hide();
            $iconLoading.show();

            const payload = {
                acao: "validarTransferencia",
                idTransferencia: idTransferencia,
                flagRejeitar: flagRejeitar,
                observacao: $('#observacao').val().trim()
            };
            
            $.ajax({
                url: 'TransferenciaCreditoRepositorio.asp?v=' + Math.random(),
                type: 'POST',
                data: payload,
                dataType: 'json',
                success: function(response) {
                    if (response.sucesso) {
                        $('#containerFormulario').hide();
                        $('#mensagemFinalizado').show();
                    } else {
                        alert('Erro: ' + response.mensagem);
                        $btnSalvar.prop('disabled', false);
                        $iconLoading.hide();
                        $iconCheck.show();
                    }
                },
                error: function(xhr, status, error) {
                    alert('Erro ao processar a validação: ' + error);
                    $btnSalvar.prop('disabled', false);
                    $iconLoading.hide();
                    $iconCheck.show();
                }
            });
        }

        function validarFormulario(e) {
            e.preventDefault();
            
            const form = document.getElementById('formularioValidacao');
            if (!form.checkValidity()) {
                e.stopPropagation();
                $(form).addClass('was-validated');
                return false;
            }

            const mensagemConfirmacao = flagRejeitar == 1 ? 
                'Deseja realmente rejeitar esta transferência?' : 
                'Deseja realmente aprovar esta transferência?';
            
            if (!confirm(mensagemConfirmacao)) return false;
            
            salvarFormulario();
        }

        $(document).ready(function() {
            carregarDadosClienteOrigem();
            carregarDadosClienteDestino();
            
            $('#formularioValidacao').on('submit', validarFormulario);
        });
    </script>
</head>
<body>
    <div class="container">
        <div id="mensagemFinalizado" style="display: none;">
            <div class="validation-message finished">
                <i class="fas fa-check-circle"></i>
                <h4>Processo Finalizado</h4>
                <p>A validação foi concluída com sucesso e o processo foi finalizado</p>
                <div class="mt-3"><a href="../SolicitacaoCreditoGrade.asp?aba=transferencias">Clique aqui para voltar a lista de transferências</a></div>
            </div>
        </div>

        <div id="containerFormulario">
            <div class="d-flex my-4">
                <button onclick="navegarParaGrade()" class="btn-back me-2"> <i class="fas fa-arrow-left"></i> </button>
                <h3><%= ObterTituloTela() %></h3>
                <button onclick="window.open('/corporativo/manuais/manual_validar_transferencia.pdf', '_blank')" class="btn-manual ms-4" title="Manual de Uso">
                    <i class="fas fa-file-alt"></i>
                </button>
            </div>

            <form name="formularioValidacao" id="formularioValidacao" method="post" class="needs-validation" novalidate>
                <div class="card">
                    
                    <div class="mb-3">
                        <label for="inputNomeCliente" class="form-label">Cliente</label>
                        <input type="text" 
                            class="form-control form-control-sm" 
                            id="inputNomeCliente" 
                            value="Carregando..." 
                            disabled>
                    </div>

                    <div class="mb-3">
                        <label for="inputValorTransferencia" class="form-label">Valor de transferência</label>
                        <input type="text" 
                            class="form-control form-control-sm" 
                            id="inputValorTransferencia" 
                            value="R$ 0,00" 
                            disabled>
                    </div>

                    <div class="transfer-container">
                        <div class="transfer-card">
                            <div class="transfer-header">
                                <div>
                                    <div class="transfer-label">Origem</div>
                                    <div class="transfer-title" id="previewOriginName"><%= ObterNomePlanta(plantaOrigem) %></div>
                                </div>
                                <div style="text-align: right;">
                                    <div class="text-muted">Cód. Cliente: <%= codClienteOrigem %></div>
                                    <div class="text-muted">Cód. Loja: <%= IIf(codLojaOrigem = "", "-", codLojaOrigem) %></div>
                                </div>
                            </div>
                            <div class="transfer-details">
                                <div class="transfer-row">
                                    <span class="transfer-row-label">Limite Atual</span>
                                    <span class="transfer-row-value" id="limiteAtualOrigem">Carregando...</span>
                                </div>
                                <div class="transfer-row">
                                    <span class="transfer-row-label">Em Aberto</span>
                                    <span class="transfer-row-value" id="limiteAbertoOrigem">Carregando...</span>
                                </div>
                                <div class="transfer-row">
                                    <span class="transfer-row-label">Saldo Disponível</span>
                                    <span class="transfer-row-value" id="saldoOrigem">Carregando...</span>
                                </div>
                                <div class="transfer-row highlight">
                                    <span class="transfer-row-label">Após Transferência</span>
                                    <span class="transfer-row-value" id="previewOriginAfter">R$ 0,00</span>
                                </div>
                            </div>
                        </div>

                        <div class="transfer-arrow-container">
                            <div class="transfer-arrow">
                                <i class="fas fa-arrow-right"></i>
                            </div>
                        </div>

                        <div class="transfer-card">
                            <div class="transfer-header">
                                <div>
                                    <div class="transfer-label">Destino</div>
                                    <div class="transfer-title" id="previewDestName"><%= ObterNomePlanta(plantaDestino) %></div>
                                </div>
                                <div style="text-align: right;">
                                    <div class="text-muted">Cód. Cliente: <%= codClienteDestino %></div>
                                    <div class="text-muted">Cód. Loja: <%= IIf(codLojaDestino = "", "-", codLojaDestino) %></div>
                                </div>
                            </div>
                            <div class="transfer-details">
                                <div class="transfer-row">
                                    <span class="transfer-row-label">Limite Atual</span>
                                    <span class="transfer-row-value" id="limiteAtualDestino">Carregando...</span>
                                </div>
                                <div class="transfer-row">
                                    <span class="transfer-row-label">Em Aberto</span>
                                    <span class="transfer-row-value" id="limiteAbertoDestino">Carregando...</span>
                                </div>
                                <div class="transfer-row">
                                    <span class="transfer-row-label">Saldo Disponível</span>
                                    <span class="transfer-row-value" id="saldoDestino">Carregando...</span>
                                </div>
                                <div class="transfer-row highlight">
                                    <span class="transfer-row-label">Após Transferência</span>
                                    <span class="transfer-row-value" id="previewDestAfter">R$ 0,00</span>
                                </div>
                            </div>
                        </div>
                    </div>

                    <div class="mt-4">
                        <label for="observacao" class="form-label">
                            Observação 
                            <%= IIf(flagRejeitar = "1", "<span class='required-field'>*</span>", "") %>
                        </label>
                        <textarea class="form-control form-control-sm" 
                            id="observacao" 
                            name="observacao" 
                            rows="3" 
                            placeholder="Digite uma observação..."
                            <%= IIf(flagRejeitar = "1", "required", "") %>><%= observacao %>
                        </textarea>
                    </div>
                </div>
                
                <div class="d-flex justify-content-end gap-2 mt-4 mb-4">
                    <button type="button" class="btn btn-sm btn-outline-secondary" onclick="navegarParaGrade()">
                        Cancelar
                    </button>
                    <button type="submit" class="btn btn-sm <%= IIf(flagRejeitar = "1", "btn-danger", "btn-primary") %>" id="btnSalvar">
                        <span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true" id="iconLoading" style="display: none;"></span>
                        <i class="fas fa-check me-2" id="iconCheck"></i>
                        <%= ObterTextoBotao() %>
                    </button>
                </div>
            </form>
        </div>
    </div>
</body>
</html>