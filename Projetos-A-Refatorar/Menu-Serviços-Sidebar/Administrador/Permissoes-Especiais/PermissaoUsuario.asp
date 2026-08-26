<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer  = true
Response.Expires = 0
Session.LCID = 1046

Dim idUsuarioLogado
idUsuarioLogado = Session("CODIGO")

Dim usuarios()
Dim idx
idx = 0

ReDim usuarios(0)

idUsuario = Request.Form("idUsuario")

Call CarregarUsuarios()

Sub CarregarUsuarios()
    SQL = "SELECT CRM_R_E_C_N_O, CRM_NOME, CRM_VERTODASFNCS, CRM_ALTERAPO, CRM_ALTERAPAPEL, AtivaUsuario,"
    SQL = SQL & "CRM_NIVELALTERAPEDIDO, CRM_PERMISSAO, ValidaPreCadastroProduto, ETools, NivelSPD, NivelSolicitacaoAlteracaoPedido, "
    SQL = SQL & "NivelLimiteCredito, NivelControladoria, NivelE2OPEN, NivelControlesInternos, NivelFinanceiro, NivelFiscal, NivelDiretorComercial "
    SQL = SQL & "FROM CRM_DTI WHERE CRM_D_E_L_E_T <> '*' ORDER BY CRM_NOME"
    
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    If IsObject(PIR_HELPDESK) Then
        While NOT PIR_HELPDESK.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = PIR_HELPDESK("CRM_R_E_C_N_O")
            usuarios(idx)("Nome") = PIR_HELPDESK("CRM_NOME")
            usuarios(idx)("verTodasFncs") = PIR_HELPDESK("CRM_VERTODASFNCS")
            usuarios(idx)("alteraPo") = PIR_HELPDESK("CRM_ALTERAPO")
            usuarios(idx)("alteraPa") = PIR_HELPDESK("CRM_ALTERAPAPEL")
            usuarios(idx)("ativaUsuario") = PIR_HELPDESK("AtivaUsuario")
            usuarios(idx)("nivelAlteraPedido") = PIR_HELPDESK("CRM_NIVELALTERAPEDIDO")
            usuarios(idx)("nivelSPD") = PIR_HELPDESK("NivelSPD")
            usuarios(idx)("nivelSolicitacaoAlteracaoPedido") = PIR_HELPDESK("NivelSolicitacaoAlteracaoPedido")
            usuarios(idx)("permissao") = PIR_HELPDESK("CRM_PERMISSAO")
            usuarios(idx)("validaPreCadastroProduto") = PIR_HELPDESK("ValidaPreCadastroProduto")
            usuarios(idx)("eTools") = PIR_HELPDESK("ETools")
            usuarios(idx)("nivelLimiteCredito") = PIR_HELPDESK("NivelLimiteCredito")
            usuarios(idx)("nivelControladoria") = PIR_HELPDESK("NivelControladoria")
            usuarios(idx)("nivelE2Open") = PIR_HELPDESK("NivelE2OPEN")
            usuarios(idx)("nivelControlesInternos") = PIR_HELPDESK("NivelControlesInternos")
            usuarios(idx)("nivelFinanceiro") = PIR_HELPDESK("NivelFinanceiro")
            usuarios(idx)("nivelFiscal") = PIR_HELPDESK("NivelFiscal")
            usuarios(idx)("nivelDiretorComercial") = PIR_HELPDESK("NivelDiretorComercial")
            idx = idx + 1
            PIR_HELPDESK.MoveNext()
        Wend
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    End If
End Sub

Function SerializarUsuarios()
    Dim json, i
    json = "{"
    For i = 0 To UBound(usuarios)
        If i > 0 Then json = json & ","
        json = json & """" & usuarios(i)("Id") & """:{"
        json = json & """nome"":""" & usuarios(i)("Nome") & ""","
        json = json & """verTodasFncs"":""" & usuarios(i)("verTodasFncs") & ""","
        json = json & """alteraPo"":""" & usuarios(i)("alteraPo") & ""","
        json = json & """alteraPa"":""" & usuarios(i)("alteraPa") & ""","
        json = json & """ativaUsuario"":""" & usuarios(i)("ativaUsuario") & ""","
        json = json & """nivelAlteraPedido"":""" & usuarios(i)("nivelAlteraPedido") & ""","
        json = json & """permissao"":""" & usuarios(i)("permissao") & ""","
        json = json & """validaPreCadastroProduto"":""" & usuarios(i)("validaPreCadastroProduto") & ""","
        json = json & """eTools"":""" & usuarios(i)("eTools") & ""","
        json = json & """nivelSPD"":""" & usuarios(i)("nivelSPD") & ""","
        json = json & """nivelSolicitacaoAlteracaoPedido"":""" & usuarios(i)("nivelSolicitacaoAlteracaoPedido") & ""","
        json = json & """nivelLimiteCredito"":""" & usuarios(i)("nivelLimiteCredito") & ""","
        json = json & """nivelControladoria"":""" & usuarios(i)("nivelControladoria") & ""","
        json = json & """nivelE2Open"":""" & usuarios(i)("nivelE2Open") & ""","
        json = json & """nivelControlesInternos"":""" & usuarios(i)("nivelControlesInternos") & ""","
        json = json & """nivelFinanceiro"":""" & usuarios(i)("nivelFinanceiro") & ""","
        json = json & """nivelDiretorComercial"":""" & usuarios(i)("nivelDiretorComercial") & ""","
        json = json & """nivelFiscal"":""" & usuarios(i)("nivelFiscal") & """"
        json = json & "}"
    Next
    json = json & "}"
    SerializarUsuarios = json
End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <title>Permissões de Usuário - Smurfit Westrock</title>
        <meta charset="utf-8"/>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
        <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />
        <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" />
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
        <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js"></script>
        <script type="text/javascript">
            var dadosUsuarios = <%= SerializarUsuarios() %>;

            function carregarDadosUsuario(idUsuario) {
                if (!idUsuario || !dadosUsuarios[idUsuario]) {
                    limparFormulario();
                    return;
                }

                var dados = dadosUsuarios[idUsuario];

                $('#nome').val(dados.nome);
                $('#verTodasFncs').prop('checked', dados.verTodasFncs == 'S');
                $('#alteraPo').prop('checked', dados.alteraPo == 'S');
                $('#alteraPa').prop('checked', dados.alteraPa == 'S');
                $('#ativaUsuario').prop('checked', dados.ativaUsuario == 'True');
                $('#nivelAlteraPedido').val(dados.nivelAlteraPedido);
                $('#nivelSPD').val(dados.nivelSPD == "NULL" ? '' : dados.nivelSPD);
                $('#nivelSolicitacaoAlteracaoPedido').val(dados.nivelSolicitacaoAlteracaoPedido == "NULL" ? '' : dados.nivelSolicitacaoAlteracaoPedido);
                $('#permissao').val(dados.permissao);
                $('#validaPreCadastroProduto').prop('checked', dados.validaPreCadastroProduto == 'True');
                $('#eTools').prop('checked', dados.eTools == 'True');
                $('#nivelLimiteCredito').prop('checked', dados.nivelLimiteCredito == 'True');
                $('#nivelControladoria').prop('checked', dados.nivelControladoria == 'True');
                $('#nivelE2Open').prop('checked', dados.nivelE2Open == 'True');
                $('#nivelControlesInternos').prop('checked', dados.nivelControlesInternos == 'True');
                $('#nivelFinanceiro').prop('checked', dados.nivelFinanceiro == 'True');
                $('#nivelDiretorComercial').prop('checked', dados.nivelDiretorComercial == 'True');
                $('#nivelFiscal').prop('checked', dados.nivelFiscal == 'True');
            }

            function limparFormulario() {
                $('input[type="checkbox"]').prop('checked', false);
                $('#nivelAlteraPedido').val('');
                $('#nivelSPD').val('');
                $('#nivelSolicitacaoAlteracaoPedido').val('');
                $('#permissao').val('');
            }

            function salvarFormulario() {
                var $btnSalvar = $('#btnSalvar');
                var $iconCheck = $('#iconCheck');
                var $iconLoading = $('#iconLoading');
                
                $btnSalvar.prop('disabled', true);
                $iconCheck.hide();
                $iconLoading.show();

                var payload = {
                    idUsuario: $('#idUsuario').val(),
                    nome: $('#nome').val(),
                    verTodasFncs: $('#verTodasFncs').is(':checked') ? 'S' : 'N',
                    alteraPo: $('#alteraPo').is(':checked') ? 'S' : 'N',
                    alteraPa: $('#alteraPa').is(':checked') ? 'S' : 'N',
                    ativaUsuario: $('#ativaUsuario').is(':checked') ? '1' : '0',
                    nivelAlteraPedido: $('#nivelAlteraPedido').val(),
                    nivelSPD: $('#nivelSPD').val(),
                    nivelSolicitacaoAlteracaoPedido: $('#nivelSolicitacaoAlteracaoPedido').val(),
                    permissao: $('#permissao').val(),
                    validaPreCadastroProduto: $('#validaPreCadastroProduto').is(':checked') ? '1' : null,
                    eTools: $('#eTools').is(':checked') ? '1' : null,
                    nivelLimiteCredito: $('#nivelLimiteCredito').is(':checked') ? '1' : null,
                    nivelControladoria: $('#nivelControladoria').is(':checked') ? '1' : null,
                    nivelE2Open: $('#nivelE2Open').is(':checked') ? '1' : null,
                    nivelControlesInternos: $('#nivelControlesInternos').is(':checked') ? '1' : null,
                    nivelFinanceiro: $('#nivelFinanceiro').is(':checked') ? '1' : null,
                    nivelDiretorComercial: $('#nivelDiretorComercial').is(':checked') ? '1' : null,
                    nivelFiscal: $('#nivelFiscal').is(':checked') ? '1' : null
                };
                
                $.ajax({
                    url: 'PermissaoUsuarioRepositorio.asp?v=' + Math.random(),
                    type: 'POST',
                    data: payload,
                    dataType: 'json',
                    success: function(response) {
                        if (response.sucesso) {
                            alert(response.mensagem);
                            $btnSalvar.prop('disabled', false);
                            $iconLoading.hide();
                            $iconCheck.show();
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
            
            $(document).ready(function() {
                $('#idUsuario').select2({
                    theme: 'bootstrap-5',
                    placeholder: 'Selecione um usuário',
                    allowClear: true,
                    width: '100%'
                });

                $('#idUsuario').on('change', function() {
                    carregarDadosUsuario($(this).val());
                });

                $('#formularioPermissao').on('submit', function(e) {
                    e.preventDefault();
                    
                    if (!this.checkValidity()) {
                        e.stopPropagation();
                        $(this).addClass('was-validated');
                        return false;
                    }
                    
                    salvarFormulario();
                });
            });
        </script>
        <style type="text/css">
            body{
                background-color: #f5f5f5;
                font-size: .9rem !important;
            }
            .form-label {
                font-weight: 600;
                margin-bottom: 8px;
            }
            .card {
                background: white;
                border-radius: 8px;
                padding: 20px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                border: none;
            }
            .form-switch .form-check-input {
                width: 3em;
                height: 1.5em;
                cursor: pointer;
            }
            .form-switch .form-check-label {
                cursor: pointer;
                font-weight: 500;
                color: #495057;
                padding-left: 8px;
            }
            .switch-container {
                border: 1px solid #dee2e6;
                border-radius: 6px;
                padding: 8px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h3 class="mt-4 mb-4">Permissões de Usuário</h3>

            <form name="formularioPermissao" id="formularioPermissao" method="post" class="needs-validation" novalidate>
                <input type="hidden" name="nome" id="nome" />

                <div class="card mb-4">      
                    <div class="row g-3">
                        <div class="col-12">
                            <label for="idUsuario" class="form-label">Usuário</label>
                            <select class="form-select" name="idUsuario" id="idUsuario" required>
                                <option value="">Selecione um usuário</option>
                                <% 
                                For idx = 0 To UBound(usuarios)
                                %>
                                <option value="<%= usuarios(idx)("Id") %>" <%= IIf(idUsuario = usuarios(idx)("Id"), "selected", "") %>>
                                    <%= usuarios(idx)("Nome") %>
                                </option>
                                <% Next %>
                            </select>
                        </div>

                        <div class="col-lg-3 col-md-6">
                            <label for="nivelAlteraPedido" class="form-label">
                                Nível Altera Pedido
                            </label>
                            <select required class="form-select" name="nivelAlteraPedido" id="nivelAlteraPedido">
                                <option value="">Selecione</option>
                                <option value="0">0 - Nível 0</option>
                                <option value="14">14 - Nível 14</option>
                            </select>
                        </div>

                        <div class="col-lg-3 col-md-6">
                            <label for="nivelSolicitacaoAlteracaoPedido" class="form-label"> Nível Solicitação Alteração Pedido</label>
                            <select class="form-select" name="nivelSolicitacaoAlteracaoPedido" id="nivelSolicitacaoAlteracaoPedido">
                                <option value="">Nenhum</option>
                                <option value="V">Validador</option>
                                <option value="P">PCP</option>
                            </select>
                        </div>

                        <div class="col-lg-3 col-md-6">
                            <label for="nivelSPD" class="form-label"> Nível SPD</label>
                            <select class="form-select" name="nivelSPD" id="nivelSPD">
                                <option value="">Nenhum</option>
                                <option value="C">Comercial</option>
                                <option value="D">Desenvolvimento</option>
                            </select>
                        </div>

                        <div class="col-lg-3 col-md-6">
                            <label for="permissao" class="form-label">
                                Permissão
                            </label>
                            <select required class="form-select" name="permissao" id="permissao">
                                <option value="">Selecione</option>
                                <option value="C">C - Validador</option>
                                <option value="G">G - Representante</option>
                                <option value="L">L - Usuário Interno</option>
                            </select>
                        </div>

                        <div class="col-md-6">
                            <div class="switch-container">
                                <div class="form-check form-switch">
                                    <input class="form-check-input" type="checkbox" name="verTodasFncs" id="verTodasFncs">
                                    <label class="form-check-label" for="verTodasFncs">
                                        Ver Todas FNCs
                                    </label>
                                </div>
                            </div>
                        </div>

                        <div class="col-md-6">
                            <div class="switch-container">
                                <div class="form-check form-switch">
                                    <input class="form-check-input" type="checkbox" name="alteraPo" id="alteraPo">
                                    <label class="form-check-label" for="alteraPo">
                                        Altera PO
                                    </label>
                                </div>
                            </div>
                        </div>

                        <div class="col-md-6">
                            <div class="switch-container">
                                <div class="form-check form-switch">
                                    <input class="form-check-input" type="checkbox" name="alteraPa" id="alteraPa">
                                    <label class="form-check-label" for="alteraPa">
                                        Altera PA
                                    </label>
                                </div>
                            </div>
                        </div>

                        <div class="col-md-6">
                            <div class="switch-container">
                                <div class="form-check form-switch">
                                    <input class="form-check-input" type="checkbox" name="ativaUsuario" id="ativaUsuario">
                                    <label class="form-check-label" for="ativaUsuario">
                                        Ativa Usuário
                                    </label>
                                </div>
                            </div>
                        </div>

                         <div class="col-md-6">
                            <div class="switch-container">
                                <div class="form-check form-switch">
                                    <input class="form-check-input" type="checkbox" name="validaPreCadastroProduto" id="validaPreCadastroProduto" value="1">
                                    <label class="form-check-label" for="validaPreCadastroProduto">
                                        Valida Pré-Cadastro Produto
                                    </label>
                                </div>
                            </div>
                        </div>

                        <div class="col-md-6">
                            <div class="switch-container">
                                <div class="form-check form-switch">
                                    <input class="form-check-input" type="checkbox" name="eTools" id="eTools" value="1">
                                    <label class="form-check-label" for="eTools">
                                        E-Tools
                                    </label>
                                </div>
                            </div>
                        </div>

                        <div class="col-md-6">
                            <div class="switch-container">
                                <div class="form-check form-switch">
                                    <input class="form-check-input" type="checkbox" name="nivelLimiteCredito" id="nivelLimiteCredito" value="1">
                                    <label class="form-check-label" for="nivelLimiteCredito">
                                        Nível Limite Crédito
                                    </label>
                                </div>
                            </div>
                        </div>

                        <div class="col-md-6">
                            <div class="switch-container">
                                <div class="form-check form-switch">
                                    <input class="form-check-input" type="checkbox" name="nivelControladoria" id="nivelControladoria" value="1">
                                    <label class="form-check-label" for="nivelControladoria">
                                        Nível Controladoria
                                    </label>
                                </div>
                            </div>
                        </div>

                        <div class="col-md-6">
                            <div class="switch-container">
                                <div class="form-check form-switch">
                                    <input class="form-check-input" type="checkbox" name="nivelE2Open" id="nivelE2Open" value="1">
                                    <label class="form-check-label" for="nivelE2Open">
                                        Nível E2OPEN
                                    </label>
                                </div>
                            </div>
                        </div>

                        <div class="col-md-6">
                            <div class="switch-container">
                                <div class="form-check form-switch">
                                    <input class="form-check-input" type="checkbox" name="nivelControlesInternos" id="nivelControlesInternos" value="1">
                                    <label class="form-check-label" for="nivelControlesInternos">
                                        Nível Controles Internos
                                    </label>
                                </div>
                            </div>
                        </div>

                         <div class="col-md-6">
                            <div class="switch-container">
                                <div class="form-check form-switch">
                                    <input class="form-check-input" type="checkbox" name="nivelDiretorComercial" id="nivelDiretorComercial" value="1">
                                    <label class="form-check-label" for="nivelDiretorComercial">
                                        Nível Diretor Comercial
                                    </label>
                                </div>
                            </div>
                        </div>

                        <div class="col-md-6">
                            <div class="switch-container">
                                <div class="form-check form-switch">
                                    <input class="form-check-input" type="checkbox" name="nivelFinanceiro" id="nivelFinanceiro" value="1">
                                    <label class="form-check-label" for="nivelFinanceiro">
                                        Nível Financeiro
                                    </label>
                                </div>
                            </div>
                        </div>

                        <div class="col-md-6">
                            <div class="switch-container">
                                <div class="form-check form-switch">
                                    <input class="form-check-input" type="checkbox" name="nivelFiscal" id="nivelFiscal" value="1">
                                    <label class="form-check-label" for="nivelFiscal">
                                        Nível Fiscal
                                    </label>
                                </div>
                            </div>
                        </div>                        
                    </div>
                </div>

                <div class="d-flex justify-content-end gap-2 mt-4 mb-4">
                    <button type="button" class="btn btn-sm btn-outline-secondary" onclick="window.location.reload()">
                        Cancelar
                    </button>
                    <button type="submit" class="btn btn-sm btn-primary" id="btnSalvar">
                        <span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true" id="iconLoading" style="display: none;"></span>
                        <i class="fas fa-check me-2" id="iconCheck"></i>
                        Salvar
                    </button>
                </div>
            </form>
        </div>
    </body>
</html>
