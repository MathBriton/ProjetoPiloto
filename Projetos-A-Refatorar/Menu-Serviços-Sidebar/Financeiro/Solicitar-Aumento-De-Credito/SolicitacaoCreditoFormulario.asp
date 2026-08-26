<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <link href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" rel="stylesheet" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css"> 
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
   
    <style type="text/css">
        body {
            background-color: #f5f5f5;
            font-size: .9rem;
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
        .card {
            padding: 1rem;
        }
        .divider {
            height: 1px;
            background-color: #dee2e6;
            margin: 20px 0;
        }
        .form-label {
            font-weight: 600;
        }
        .required-field {
            color: #dc3545;
        }
        .error-message {
            background-color: #f8d7da;
            border: 1px solid #dc3545;
            color: #842029;
            padding: 10px;
            border-radius: 8px;
            text-align: center;
        }
        .error-message i {
            font-size: 2rem;
            margin-bottom: 15px;
            color: #dc3545;
        }
        .success-message {
            background-color: #d1e7dd;
            border: 1px solid #198754;
            color: #0f5132;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            margin: 40px auto;
            max-width: 600px;
        }
        .success-message i {
            font-size: 3rem;
            margin-bottom: 15px;
        }
        .success-message h4 {
            margin-bottom: 10px;
        }
        .loading-container {
            text-align: center;
            padding: 30px;
        }
        .loading-container .spinner-border {
            width: 3rem;
            height: 3rem;
            color: #0d6efd;
        }
        .loading-container p {
            margin-top: 15px;
            color: #6c757d;
        }
        .form-switch {
            display: flex;
            align-items: end;
        }
        .form-switch .form-check-input {
            width: 3em;
            height: 1.5em;
            cursor: pointer;
        }
        .form-switch .form-check-label {
            cursor: pointer;
            font-weight: 500;
            padding-left: 8px;
        }
        .switch-container {
            border: 1px solid #ced4da;
            border-radius: 6px;
            padding: 8px;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/jquery-maskmoney@3.0.2/dist/jquery.maskMoney.min.js"></script>        
    <script src="./util/Util.js"></script>
    <script src="/corporativo/components/upload/fileUpload.js?v=<%=Now()%>"></script>

    <script type="text/javascript">
        function PesquisaCliente() {
            const codPlanta = $('#planta').val();
            if (!codPlanta) return alert('Selecione a planta primeiro');
            const url = `${location.protocol}//${location.host}/corporativo/buscaCLIENTE.asp?id1=codCliente&id2=codLoja&id3=nomeCliente&id6=${codPlanta}`;
            window.open(url, '', "width=650,height=400");
        }

        function plantaAlterada() {
            const planta = $('#planta');
            const codLoja = $('#codLoja');
            
            const divCamposPadrao = $('#divCamposPadrao');
            const precoM2 = $('#precoM2');
            const precoKg = $('#precoKg');
            const volumeM2 = $('#volumeM2');
            
            const divCamposFlorestal = $('#divCamposFlorestal');
            const valorToneladaTora = $('#valorToneladaTora');
            const valorUnitarioSemente = $('#valorUnitarioSemente');
            const valorUnitarioMuda = $('#valorUnitarioMuda');

            if (planta.val() == "07") {
                codLoja.prop('required', true);
            } else {
                codLoja.prop('required', false);
            }
            
            if (planta.val() == "13") {
                divCamposFlorestal.show();
                divCamposPadrao.hide();
                
                valorToneladaTora.prop('required', false);
                valorUnitarioSemente.prop('required', false);
                valorUnitarioMuda.prop('required', false);
                
                precoM2.prop('required', false);
                precoKg.prop('required', false);
                volumeM2.prop('required', false);
            } else {
                divCamposPadrao.show();
                divCamposFlorestal.hide();
                
                precoM2.prop('required', true);
                precoKg.prop('required', true);
                volumeM2.prop('required', true);
                
                valorToneladaTora.prop('required', false);
                valorUnitarioSemente.prop('required', false);
                valorUnitarioMuda.prop('required', false);
            }
            
            limparDadosCliente();
            outrasVendasAlterado();
        }
        
        function outrasVendasAlterado() {
            const planta = $('#planta').val();
            const outrasVendas = $('#switchOutrasVendas').is(':checked');
            const precoM2 = $('#precoM2');
            const volumeM2 = $('#volumeM2');
            const volumeUnidades = $('#volumeUnidades');
            const volumeToneladas = $('#volumeToneladas');
            const divPrecoM2 = $('#divPrecoM2');
            const divVolumeM2 = $('#divVolumeM2');
            const divVolumeUnidades = $('#divVolumeUnidades');
            const divVolumeToneladas = $('#divVolumeToneladas');
            const volumeUnidadesFlorestal = $('#volumeUnidadesFlorestal');
            const volumeToneladasFlorestal = $('#volumeToneladasFlorestal');
            const divVolumeUnidadesFlorestal = $('#divVolumeUnidadesFlorestal');
            const divVolumeToneladasFlorestal = $('#divVolumeToneladasFlorestal');
            
            if (planta == "13") {
                if (outrasVendas) {
                    divVolumeUnidadesFlorestal.show();
                    divVolumeToneladasFlorestal.show();
                    
                    volumeUnidadesFlorestal.prop('required', false);
                    volumeToneladasFlorestal.prop('required', false);
                } else {
                    divVolumeUnidadesFlorestal.hide();
                    divVolumeToneladasFlorestal.hide();
                    
                    volumeUnidadesFlorestal.prop('required', false);
                    volumeToneladasFlorestal.prop('required', false);
                }
            } else {
                if (outrasVendas) {
                    divPrecoM2.hide();
                    divVolumeM2.hide();
                    divVolumeUnidades.show();
                    divVolumeToneladas.show();
                    
                    precoM2.prop('required', false);
                    volumeM2.prop('required', false);
                    volumeUnidades.prop('required', false);
                    volumeToneladas.prop('required', false);
                } else {
                    divPrecoM2.show();
                    divVolumeM2.show();
                    divVolumeUnidades.hide();
                    divVolumeToneladas.hide();
                    
                    precoM2.prop('required', true);
                    volumeM2.prop('required', true);
                    volumeUnidades.prop('required', false);
                    volumeToneladas.prop('required', false);
                }
            }
            tipoMercadoAlterado();
        }
        
        function tipoMercadoAlterado() {
            const exportacao = $('#tipoMercado').val() == 'EXPORTACAO';
            const outrasVendas = $('#switchOutrasVendas').is(':checked');
            const planta = $('#planta').val();
            const precoM2 = $('#precoM2');
            const precoKg = $('#precoKg');
            const requiredPrecoM2 = $('#requiredPrecoM2');

            $('#unidadePrecoM2').text(exportacao ? 'Preços USD/m2 - NET' : 'Preços R$/m2 - NET');
            $('#unidadePrecoKg').text(exportacao ? 'Preços USD/kg - NET' : 'Preços R$/kg - NET');

            if (exportacao) {
                precoM2.prop('required', false);
                requiredPrecoM2.removeClass('required-field').text('(opcional)');
            } else {
                const precoM2Obrigatorio = planta != '13' && !outrasVendas;
                precoM2.prop('required', precoM2Obrigatorio);
                requiredPrecoM2.toggleClass('required-field', precoM2Obrigatorio).text(precoM2Obrigatorio ? '*' : '');
            }

            precoKg.prop('required', planta != '13');
        }
        
        function limiteAlterado() {
            const limiteSolicitadoStr = $('#limiteSolicitado').val();
            const limiteSolicitado = parseFloat(desformatarMoeda(limiteSolicitadoStr)) || 0;
            const tituloDocumentos = $('#tituloDocumentos');
            const requiredBalanco = $('#requiredBalanco');
            const requiredContrato = $('#requiredContrato');
            
            requiredContrato.addClass('required-field').text('*');
            
            if (limiteSolicitado > 180000) {
                tituloDocumentos.html('<i class="fas fa-exclamation-circle me-2"></i>Documentos Obrigatórios (Limite acima de R$ 180.000,00)');
                tituloDocumentos.removeClass('text-secondary').addClass('text-primary');
                requiredBalanco.addClass('required-field').text('*');
            } else {
                tituloDocumentos.html('<i class="fas fa-file-upload me-2"></i>Documentos');
                tituloDocumentos.removeClass('text-primary').addClass('text-secondary');
                requiredBalanco.removeClass('required-field').text('(opcional)');
            }
        }

        function obterCliente() {
            const planta = $('#planta').val();
            const codCliente = $('#codCliente').val();
            const codLoja = $('#codLoja').val();

            if (!planta || !codCliente || !codLoja) return;

            mostrarLoadingCliente();

            $.ajax({
                url: 'SolicitacaoCreditoRepositorio.asp?v=' + Math.random(),
                type: 'POST',
                data: {
                    acao: "obterCliente",
                    planta: planta,
                    codCliente: codCliente,
                    codPlanta: planta,
                    codLoja: codLoja.padStart(2, '0')
                },
                dataType: 'json',
                success: function(response) {
                    if (response.sucesso) {
                        $('#nomeCliente').val(response.dados.nomeCliente);
                        $('#limiteAtual').val(formatarMoeda(response.dados.limite).toString());
                        $('#saldo').val(formatarMoeda(response.dados.saldo).toString());
                        $('#limiteAberto').val(formatarMoeda(response.dados.aberto).toString());
                        $('#condicaoPagamento').val(response.dados.condicaoPagamento);
                        $('#codLoja').val(response.dados.codLoja);
                        
                        if (response.dados.codClientePrincipal) {
                            $('#codClientePrincipal').val(response.dados.codClientePrincipal);
                            $('#nomeClientePrincipal').val(response.dados.nomeClientePrincipal);
                            $('#divClientePrincipal, #divNomeClientePrincipal').show();
                        } else {
                            $('#codClientePrincipal').val('');
                            $('#nomeClientePrincipal').val('');
                            $('#divClientePrincipal, #divNomeClientePrincipal').hide();
                        }

                        mostrarDadosCliente();
                    } else {
                        mostrarErroCliente(response.mensagem);
                        $('#nomeCliente').val("");
                        $('#limiteAtual').val("");
                        $('#saldo').val("");
                        $('#limiteAberto').val("");
                        $('#codClientePrincipal').val('');
                        $('#nomeClientePrincipal').val('');
                        $('#divClientePrincipal, #divNomeClientePrincipal').hide();
                    }
                },
                error: function(xhr, status, error) {
                    console.log('Erro ao obter saldo do cliente: ' + error);
                    mostrarErroCliente('Erro ao carregar dados do cliente. Por favor, tente novamente.');
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
            
            const uploadBalancoDRE = document.getElementById('uploadBalancoDRE');
            const uploadContratoSocial = document.getElementById('uploadContratoSocial');
            const uploadDocumentosComplementares = document.getElementById('uploadDocumentosComplementares');

            const arquivosBalanco = uploadBalancoDRE.getFiles();
            const arquivoContrato = uploadContratoSocial.getFile();
            const documentosComplementares = uploadDocumentosComplementares.getFiles();

            const anexos = [];
            if (arquivosBalanco[0]) 
                anexos.push({ name: "balanco_dre_1", uploadedName: arquivosBalanco[0].uploadedName })

            if (arquivosBalanco[1]) 
                anexos.push({ name: "balanco_dre_2", uploadedName: arquivosBalanco[1].uploadedName })
            
            if (arquivoContrato) 
                anexos.push({ name: "contrato_social", uploadedName: arquivoContrato.uploadedName })
            
            if (documentosComplementares.length > 0)
                anexos.push(...documentosComplementares)
            
            const planta = $('#planta').val();
            const outrasVendas = $('#switchOutrasVendas').is(':checked');
            
            const formData = {
                acao: "criarSolicitacao",
                planta: planta,
                codCliente: $('#codCliente').val(),
                codLoja: $('#codLoja').val(),
                tipoMercado: $('#tipoMercado').val(),
                observacaoVendedor: $('#observacaoVendedor').val(),
                limiteSolicitado: desformatarMoeda($('#limiteSolicitado').val()),
                condicaoPagamento: $('#condicaoPagamento').val(),
                anexos: JSON.stringify(anexos),
            };
            
            if (planta == "13") {
                formData.valorToneladaTora = desformatarMoeda($('#valorToneladaTora').val());
                formData.valorUnitarioSemente = desformatarMoeda($('#valorUnitarioSemente').val());
                formData.valorUnitarioMuda = desformatarMoeda($('#valorUnitarioMuda').val());
                
                if (outrasVendas) {
                    formData.outrasVendas = "1";
                    formData.volumeUnidades = $('#volumeUnidadesFlorestal').val();
                    formData.volumeToneladas = $('#volumeToneladasFlorestal').val();
                } else {
                    formData.outrasVendas = "0";
                }
            } else {
                formData.precoKg = desformatarMoeda($('#precoKg').val());
                
                if (outrasVendas) {
                    formData.outrasVendas = "1";
                    formData.volumeUnidades = $('#volumeUnidades').val();
                    formData.volumeToneladas = $('#volumeToneladas').val();
                } else {
                    formData.outrasVendas = "0";
                    formData.volumeM2 = $('#volumeM2').val();
                    formData.precoM2 = desformatarMoeda($('#precoM2').val());
                }
            }
            
            $.ajax({
                url: 'SolicitacaoCreditoRepositorio.asp',
                type: 'POST',
                data: formData,
                dataType: 'json',
                success: function(response) {
                    if (response.sucesso) {
                        mostrarMensagemSucesso();
                    } else {
                        alert('Erro: ' + response.mensagem);
                        resetarBotaoSalvar($btnSalvar, $iconCheck, $iconLoading);
                    }
                },
                error: function(xhr, status, error) {
                    alert('Erro ao processar a requisição: ' + error);
                    resetarBotaoSalvar($btnSalvar, $iconCheck, $iconLoading);
                }
            });
        }

        function resetarBotaoSalvar($btnSalvar, $iconCheck, $iconLoading) {
            $btnSalvar.prop('disabled', false);
            $iconLoading.hide();
            $iconCheck.show();
        }

        function limparDadosCliente() {
            $('#codCliente').val('');
            $('#codLoja').val('');
            $('#nomeCliente').val('');
            $('#limiteAtual').val('');
            $('#saldo').val('');
            $('#limiteAberto').val('');
            $('#codClientePrincipal').val('');
            $('#nomeClientePrincipal').val('');
            $('#divClientePrincipal, #divNomeClientePrincipal').hide();
            $('#divDadosCliente').hide();
            $('#divLoadingCliente').hide();
            $('#divErroCliente').hide();
        }

        function mostrarDadosCliente() {
            $('#divDadosCliente').show();
            $('#divLoadingCliente').hide();
            $('#divErroCliente').hide();
        }
        
        function mostrarLoadingCliente() {
            $('#divLoadingCliente').show();
            $('#divDadosCliente').hide();
            $('#divErroCliente').hide();
        }
        
        function mostrarErroCliente(mensagem) {
            $('#mensagemErroCliente').text(mensagem);
            $('#divErroCliente').show();
            $('#divLoadingCliente').hide();
            $('#divDadosCliente').hide();
        }
                    
        function mostrarMensagemSucesso() {
            $('#mensagemSucesso').show();
            $('#containerFormulario').hide();
        }

        function validarFormulario(e) {
            e.preventDefault();
            
            const form = document.getElementById('formularioSolicitacao');
            if (!form.checkValidity()) {
                e.stopPropagation();
                $(form).addClass('was-validated');
                return false;
            }

            const limiteSolicitado = parseFloat(desformatarMoeda($('#limiteSolicitado').val()));
            const limiteAberto = parseFloat(desformatarMoeda($('#limiteAberto').val()));

            if (limiteSolicitado <= limiteAberto) {
                alert('O limite desejado deve ser maior que o limite em aberto.');
                return false;
            }
            
            const uploadBalancoDRE = document.getElementById('uploadBalancoDRE');
            const uploadContratoSocial = document.getElementById('uploadContratoSocial');
            
            const arquivosBalanco = uploadBalancoDRE.getFiles();
            const arquivoContrato = uploadContratoSocial.getFile();
            
            if (!arquivoContrato) {
                alert('Por favor, anexe o Contrato Social. Este documento é obrigatório.');
                return false;
            }
            
            if (limiteSolicitado > 180000) {
                if (arquivosBalanco.length == 0) {
                    alert('Por favor, anexe o Balanço DRE do exercício para limites acima de R$ 180.000,00.');
                    return false;
                }
            }
            
            salvarFormulario();
        }

        $(document).ready(function() {
            $('.select2').select2({  theme: 'bootstrap-5', width: '100%' });
            $("#limiteSolicitado").maskMoney({thousands:'.', decimal:','});
            $("#precoM2").maskMoney({thousands:'.', decimal:','});
            $("#precoKg").maskMoney({thousands:'.', decimal:','});
            $("#valorToneladaTora").maskMoney({thousands:'.', decimal:','});
            $("#valorUnitarioSemente").maskMoney({thousands:'.', decimal:','});
            $("#valorUnitarioMuda").maskMoney({thousands:'.', decimal:','});

            limiteAlterado();
            plantaAlterada();
            
            $('#formularioSolicitacao').on('submit', validarFormulario);

            $('#uploadDocumentosComplementares, #uploadContratoSocial, #uploadBalancoDRE').on('error', (e) => {
                alert(e.detail.message);
            });
            
            $('#switchOutrasVendas').on('change', outrasVendasAlterado);
        });
    </script>
</head>
<body>
    <div class="container">
        <div id="mensagemSucesso" style="display: none;">
            <div class="success-message">
                <i class="fas fa-check-circle"></i>
                <h4>Solicitação criada com sucesso!</h4>
                <p>A solicitação de crédito foi enviada para análise</p>
                <div class="mt-3"><a href="SolicitacaoCreditoFormulario.asp">Clique aqui para criar uma nova solicitação</a></div>
            </div>
        </div>

        <div id="containerFormulario">
            <div class="d-flex my-4">
                <h3>Nova Solicitação de Crédito</h3>
                <button onclick="window.open('/corporativo/manuais/manual_solicitar_credito.pdf', '_blank')" class="btn-manual ms-4" title="Manual de Uso">
                    <i class="fas fa-file-alt"></i>
                </button>
            </div>

            <form name="formularioSolicitacao" id="formularioSolicitacao" method="post">                              
                <div class="card">
                    <div class="row g-3">
                        <div class="col-12">
                            <label for="planta" class="form-label">Planta <span class="required-field">*</span></label>
                            <select class="form-select form-select-sm select2" name="planta" id="planta" required onchange="plantaAlterada()">
                                <option value="">Selecione a planta</option>
                                <option value="00">Pirapetinga</option>
                                <option value="04">Uberaba</option>
                                <option value="06">Maranguape</option>
                                <option value="05">Bento</option>
                                <option value="07">Saquarema</option>
                                <option value="08">Aracatuba</option>
                                <option value="09">Blumenau</option>
                                <option value="10">Pacajus</option>
                                <option value="11">Porto Feliz</option>
                                <option value="12">Tres Barras</option>
                                <option value="13">Florestal</option>
                            </select>
                        </div>
                        <div class="col-md-6">
                            <label for="codCliente" class="form-label">Cód. Cliente <span class="required-field me-2">*</span></label>
                            <div class="input-group">
                                <span class="btn btn-sm btn-secondary" onclick="PesquisaCliente()">Buscar <i class="fa fa-external-link-alt"></i></span>
                                <input required onblur="obterCliente()" type="text" class="form-control form-control-sm" id="codCliente" name="codCliente" required>
                            </div>
                            <small class="text-muted">Clique no botão para buscar ou digite um cód.</small>
                        </div>
                        <div class="col-md-6">
                            <label for="codLoja" class="form-label">Cód. Loja <span class="required-field">*</span></label>
                            <input required onblur="obterCliente()" type="text" class="form-control form-control-sm" value="" name="codLoja" id="codLoja" size="2" maxlength="2"/>
                        </div>
                    </div>

                    <div id="divLoadingCliente" style="display: none;">
                        <div class="divider"></div>
                        <div class="loading-container">
                            <div class="spinner-border"></div>
                            <p>Carregando dados do cliente...</p>
                        </div>
                    </div>

                    <div id="divErroCliente" style="display: none;">
                        <div class="divider"></div>
                        <div class="error-message">
                            <i class="fas fa-times-circle"></i>
                            <h6>Erro ao Carregar Dados</h6>
                            <div id="mensagemErroCliente"></div>
                        </div>
                    </div>

                    <div id="divDadosCliente" style="display: none;">
                        <div class="divider"></div>

                        <div class="row g-3">
                            <div class="col-md-6">
                                <label for="nomeCliente" class="form-label">Nome do cliente</label>
                                <input disabled class="form-control form-control-sm" name="nomeCliente" id="nomeCliente" />
                            </div>
                            <div class="col-md-6">
                                <label for="limiteAtual" class="form-label">Limite Atual</label>
                                <input disabled type="text" class="form-control form-control-sm" name="limiteAtual" id="limiteAtual" />
                            </div>
                            <div class="col-md-6">
                                <label for="limiteAberto" class="form-label">Limite em aberto</label>
                                <input disabled type="text" class="form-control form-control-sm" name="limiteAberto" id="limiteAberto" />
                            </div>
                            <div class="col-md-6">
                                <label for="saldo" class="form-label">Saldo</label>
                                <input disabled type="text" class="form-control form-control-sm" name="saldo" id="saldo" />
                            </div>
                            <div class="col-md-6" id="divClientePrincipal" style="display: none;">
                                <label for="codClientePrincipal" class="form-label">Código Cliente Principal</label>
                                <input disabled type="text" class="form-control form-control-sm" name="codClientePrincipal" id="codClientePrincipal" />
                            </div>
                            <div class="col-md-6" id="divNomeClientePrincipal" style="display: none;">
                                <label for="nomeClientePrincipal" class="form-label">Nome Cliente Principal</label>
                                <input disabled type="text" class="form-control form-control-sm" name="nomeClientePrincipal" id="nomeClientePrincipal" />
                            </div>
                        </div>
                    </div>

                    <div class="divider"></div>

                    <div class="row g-3">
                        <div id="divCamposPadrao">
                            <div class="row g-3">
                                <div class="col-12">
                                    <div class="switch-container">
                                        <div class="form-check form-switch">
                                            <input class="form-check-input" type="checkbox" id="switchOutrasVendas" name="switchOutrasVendas">
                                            <label class="form-check-label" for="switchOutrasVendas">Outras vendas (Resíduos/Aparas)</label>
                                        </div>
                                    </div>
                                </div>
                                <div class="col-md-4" id="divPrecoM2">
                                    <label for="precoM2" class="form-label"> 
                                        <span id="unidadePrecoM2">Preços R$/m2 - NET</span>
                                       <span class="required-field" id="requiredPrecoM2">*</span>
                                    </label>
                                    <input type="text" class="form-control form-control-sm" name="precoM2" id="precoM2" />
                                </div>
                                <div class="col-md-4">
                                    <label for="precoKg" class="form-label">
                                        <span id="unidadePrecoKg">Preços R$/kg - NET</span>
                                       <span class="required-field">*</span>
                                    </label>
                                    <input type="text" class="form-control form-control-sm" name="precoKg" id="precoKg" />
                                </div>
                                <div class="col-md-4" id="divVolumeM2">
                                    <label for="volumeM2" class="form-label">Potencial Volume de Vendas MSM mensal<span class="required-field">*</span></label>
                                    <input type="number" min="1" max="99999999" class="form-control form-control-sm" name="volumeM2" id="volumeM2" />
                                </div>
                                <div class="col-md-4" id="divVolumeUnidades" style="display: none;">
                                    <label for="volumeUnidades" class="form-label">Potencial Volume mensal (unidades)<span>(opcional)</span></label>
                                    <input type="number" min="0" max="99999999" class="form-control form-control-sm" name="volumeUnidades" id="volumeUnidades" />
                                </div>
                                <div class="col-md-4" id="divVolumeToneladas" style="display: none;">
                                    <label for="volumeToneladas" class="form-label">Potencial Volume mensal (toneladas)<span>(opcional)</span></label>
                                    <input type="number" min="0" max="99999999" class="form-control form-control-sm" name="volumeToneladas" id="volumeToneladas" />
                                </div>
                            </div>
                        </div>
                        
                        <div id="divCamposFlorestal" style="display: none;">
                            <div class="row g-3">
                                <div class="col-12">
                                    <div class="switch-container">
                                        <div class="form-check form-switch">
                                            <input class="form-check-input" type="checkbox" id="switchOutrasVendas" name="switchOutrasVendas">
                                            <label class="form-check-label" for="switchOutrasVendas">Outras vendas (Resíduos/Aparas)</label>
                                        </div>
                                    </div>
                                </div>
                                <div class="col-md-4">
                                    <label for="valorToneladaTora" class="form-label"> 
                                        Tora: Valor R$/ton - BRUTO
                                       <span>(opcional)</span>
                                    </label>
                                    <input type="text" class="form-control form-control-sm" name="valorToneladaTora" id="valorToneladaTora" />
                                </div>
                                <div class="col-md-4">
                                    <label for="valorUnitarioSemente" class="form-label">
                                        Sementes: R$/kg - BRUTO e R$/unidade - BRUTO
                                       <span>(opcional)</span>
                                    </label>
                                    <input type="text" class="form-control form-control-sm" name="valorUnitarioSemente" id="valorUnitarioSemente" />
                                </div>
                                <div class="col-md-4">
                                    <label for="valorUnitarioMuda" class="form-label">
                                        Mudas: R$/unidade - BRUTO
                                       <span>(opcional)</span>
                                    </label>
                                    <input type="text" class="form-control form-control-sm" name="valorUnitarioMuda" id="valorUnitarioMuda" />
                                </div>
                                <div class="col-md-4" id="divVolumeUnidadesFlorestal" style="display: none;">
                                    <label for="volumeUnidadesFlorestal" class="form-label">Potencial Volume mensal (unidades)<span>(opcional)</span></label>
                                    <input type="number" min="0" max="99999999" class="form-control form-control-sm" name="volumeUnidadesFlorestal" id="volumeUnidadesFlorestal" />
                                </div>
                                <div class="col-md-4" id="divVolumeToneladasFlorestal" style="display: none;">
                                    <label for="volumeToneladasFlorestal" class="form-label">Potencial Volume mensal (toneladas)<span>(opcional)</span></label>
                                    <input type="number" min="0" max="99999999" class="form-control form-control-sm" name="volumeToneladasFlorestal" id="volumeToneladasFlorestal" />
                                </div>
                            </div>
                        </div>

                        <div class="col-md-4">
                            <label for="tipoMercado" class="form-label"> 
                                Tipo de Mercado
                               <span class="required-field">*</span>
                            </label>
                           <select class="form-select form-select-sm select2" name="tipoMercado" id="tipoMercado" required onchange="tipoMercadoAlterado()">
                                <option value="">Selecione o tipo de mercado</option>
                                <option value="NACIONAL">Nacional</option>
                                <option value="EXPORTACAO">Exportação</option>
                            </select>
                        </div>
                        
                        <div class="col-md-4">
                            <label for="condicaoPagamento" class="form-label">Condição de pagamento <span class="required-field">*</span></label>
                            <input disabled class="form-control form-control-sm" name="condicaoPagamento" id="condicaoPagamento" required />
                        </div>
                        <div class="col-md-4">
                            <label for="limiteSolicitado" class="form-label">Limite Desejado <span class="required-field">*</span></label>
                            <input onblur="limiteAlterado()" type="text" class="form-control form-control-sm" name="limiteSolicitado" id="limiteSolicitado" required />
                        </div>
                        <div class="col-12">
                            <label for="observacaoVendedor" class="form-label">Observação</label>
                            <textarea class="form-control form-control-sm" 
                                id="observacaoVendedor" 
                                name="observacaoVendedor" 
                                rows="5" 
                                placeholder="Digite uma observação..."></textarea>
                        </div>
                    </div>

                    <div id="divDocumentos"> 
                        <div class="divider"></div>

                        <div class="row g-3">
                            <h6 class="text-secondary mt-3" id="tituloDocumentos">
                                <i class="fas fa-file-upload me-2"></i>
                                Documentos
                            </h6>
                            <div class="col-12">
                                <label class="form-label">Contrato Social <span class="required-field" id="requiredContrato">*</span></label>
                                <file-upload file-code="SOLICITACAO_CREDITO" id="uploadContratoSocial"></file-upload>
                            </div>
                            <div class="col-12">
                                <label class="form-label">Balanço DRE do exercício <span class="required-field" id="requiredBalanco">*</span></label>
                                <file-upload max-files="2" multiple file-code="SOLICITACAO_CREDITO" id="uploadBalancoDRE"></file-upload>
                            </div>
                            <div class="col-12">
                                <label class="form-label">Documentos Complementares <span>(opcional)</span></label>
                                <file-upload 
                                    id="uploadDocumentosComplementares"
                                    file-code="SOLICITACAO_CREDITO"
                                    max-files="5"
                                    multiple>
                                </file-upload>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="d-flex justify-content-end gap-2 mt-4 mb-4">
                    <button type="submit" class="btn btn-sm btn-primary" id="btnSalvar">
                        <span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true" id="iconLoading" style="display: none;"></span>
                        <i class="fa fa-check me-2" id="iconCheck"></i> Salvar
                    </button>
                </div>
            </form>
        </div>
    </div>
</body>
</html>