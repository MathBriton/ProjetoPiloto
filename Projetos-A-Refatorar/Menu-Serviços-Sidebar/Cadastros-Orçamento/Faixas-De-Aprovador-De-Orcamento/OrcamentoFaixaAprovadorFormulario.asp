<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<!--#include File = "../Orcamento/include/FuncoesDoOrcamento.asp"-->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer  = true
Response.Expires = 0
Session.LCID = 1046

Dim plantas()
Dim idFaixa, acao
idFaixa = 0
acao = ""

If Request("idFaixa") <> "" Then
    idFaixa = CLng(Request("idFaixa"))
End If

If Request("acao") <> "" Then
    acao = Request("acao")
End If

'Variaveis auxiliares'
Dim idx
idx = 0

'----------------------------------------------
'          Carrega Lista de plantas
'----------------------------------------------
ReDim plantas(0)
SQL = "SELECT Empresa, Razao FROM Empresas ORDER BY Razao"

SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
While NOT PIR_HELPDESK.EOF
    ReDim Preserve plantas(idx)
    Set plantas(idx) = Server.CreateObject("Scripting.Dictionary")
    plantas(idx)("Nome") = PIR_HELPDESK("Razao")
    plantas(idx)("Codigo") = PIR_HELPDESK("Empresa")
    idx = idx + 1
    PIR_HELPDESK.MoveNext()
Wend
If isObject(PIR_HELPDESK) Then
    PIR_HELPDESK.Close()
    SET PIR_HELPDESK = NOTHING
End If

Dim planta, tipoProduto, valorMinKg, valorMaxKg, valorMinM2, valorMaxM2, quantidade
planta = ""
tipoProduto = ""
valorMinKg = 999999
valorMaxKg = 999999
valorMinM2 = 999999
valorMaxM2 = 999999
quantidade = 1

'----------------------------------------------
'          Carrega Registro a ser editado (se for edição)
'----------------------------------------------

IF idFaixa > 0 AND acao = "edicao" THEN
    SQL = "SELECT TOP 1 * FROM FAIXA_APROVADORES WHERE ID = " & idFaixa
    
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    If NOT PIR_CORPORATIVO.EOF Then
        idFaixa = PIR_CORPORATIVO("ID")
        planta = PIR_CORPORATIVO("PLANTA")
        tipoProduto = PIR_CORPORATIVO("TIPO_PRODUTO")
        valorMinKg = PIR_CORPORATIVO("VALOR_MIN_KG")
        valorMaxKg = PIR_CORPORATIVO("VALOR_MAX_KG")
        valorMinM2 = PIR_CORPORATIVO("VALOR_MIN_M2")
        valorMaxM2 = PIR_CORPORATIVO("VALOR_MAX_M2")
        quantidade = PIR_CORPORATIVO("QUANTIDADE")
    End If
    If isObject(PIR_CORPORATIVO) Then
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
END IF


'Função para converter decimal para formato HTML (com ponto)
Function FormatarDecimalParaHTML(valor)
    If IsNumeric(valor) Then
        FormatarDecimalParaHTML = Replace(CStr(valor), ",", ".")
    Else
        FormatarDecimalParaHTML = valor
    End If
End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <title>Smurfit Kappa - Faixa de Aprovador de Orçamento</title>
        <meta charset="utf-8"/>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
        <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />
        <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
        <link href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" rel="stylesheet" />
        
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
        <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
        
        <style type="text/css">
            body{
                font-size:80% !important;
            }
            .header {
                margin: 25px 0;
            }
            .btn {
                font-size: .8rem;
                width: fit-content;
            }
            .form-label {
                font-weight: bold;
            }
            .spinner-border-sm {
                width: 1rem;
                height: 1rem;
                border-width: 0.15em;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header border-bottom">
                <h3><% If acao = "edicao" Then Response.Write("Editar") Else Response.Write("Nova") End If %> Faixa de Aprovador de Orçamento</h3>
            </div>
            
            <form name="formularioFaixa" id="formularioFaixa" method="post">
                <input type="hidden" name="acao" id="acao" value="<%= acao %>" />
                <input type="hidden" name="idFaixa" id="idFaixa" value="<%= idFaixa %>" />
                
                <div class="row mb-3">
                    <div class="col-md-6">
                        <label for="planta" class="form-label">Planta *</label>
                        <select class="form-select" name="planta" id="planta" required>
                            <option value="">Selecione uma planta</option>
                            <%
                            For idx = 0 To UBound(plantas)
                                Dim selectedPlanta
                                selectedPlanta = ""
                                If CStr(plantas(idx)("Codigo")) = CStr(planta) Then
                                    selectedPlanta = " selected"
                                End If
                                Response.Write("<option value='" & plantas(idx)("Codigo") & "'" & selectedPlanta & ">" & plantas(idx)("Nome") & "</option>")
                            Next
                            %>
                        </select>
                    </div>
                    
                    <div class="col-md-6">
                        <label for="tipoProduto" class="form-label">Tipo de Produto *</label>
                        <select class="form-select" name="tipoProduto" id="tipoProduto" required>
                            <option value="">Selecione um tipo</option>
                            <option value="CX" <% If tipoProduto = "CX" Then Response.Write("selected") End If %>>Caixa</option>
                            <option value="CH" <% If tipoProduto = "CH" Then Response.Write("selected") End If %>>Chapa</option>
                            <option value="PA" <% If tipoProduto = "PA" Then Response.Write("selected") End If %>>Papel</option>
                            <option value="CJ" <% If tipoProduto = "CJ" Then Response.Write("selected") End If %>>Conjunto</option>
                        </select>
                    </div>
                </div>
                
                <div class="row mb-3">
                    <div class="col-md-6">
                        <label for="valorMinKg" class="form-label">Valor Mínimo Kg *</label>
                        <input type="number" step="0.00001" min="0" max="999999" class="form-control" name="valorMinKg" id="valorMinKg" value="<%= FormatarDecimalParaHTML(valorMinKg) %>" required />
                    </div>
                    
                    <div class="col-md-6">
                        <label for="valorMaxKg" class="form-label">Valor Máximo Kg *</label>
                        <input type="number" step="0.00001" min="0" max="999999" class="form-control" name="valorMaxKg" id="valorMaxKg" value="<%= FormatarDecimalParaHTML(valorMaxKg) %>" required />
                    </div>
                </div>
                
                <div class="row mb-3">
                    <div class="col-md-6">
                        <label for="valorMinM2" class="form-label">Valor Mínimo M² *</label>
                        <input type="number" step="0.00001" min="0" max="999999" class="form-control" name="valorMinM2" id="valorMinM2" value="<%= FormatarDecimalParaHTML(valorMinM2) %>" required />
                    </div>
                    
                    <div class="col-md-6">
                        <label for="valorMaxM2" class="form-label">Valor Máximo M² *</label>
                        <input type="number" step="0.00001" min="0" max="999999" class="form-control" name="valorMaxM2" id="valorMaxM2" value="<%= FormatarDecimalParaHTML(valorMaxM2) %>" required />
                    </div>
                </div>
                
                <div class="row mb-3">
                    <div class="col-md-6">
                        <label for="quantidade" class="form-label">Qtd. Aprovadores *</label>
                        <input type="number" step="1" min="1" class="form-control" name="quantidade" id="quantidade" value="<%= quantidade %>" required />
                    </div>
                </div>
                
                <div class="row mt-4">
                    <div class="col-12">
                        <button style="margin-right: 10px" type="button" class="btn btn-secondary" id="btnVoltar" onclick="history.back()">
                            <i class="fas fa-arrow-left me-2"></i> Voltar
                        </button>
                        <button type="submit" class="btn btn-primary" id="btnSalvar">
                            <i class="fas fa-check me-2"></i> Salvar
                        </button>
                    </div>
                </div>
            </form>
        </div>
        
        <script type="text/javascript">
            $(document).ready(function() {
                $('#planta').select2({
                    theme: 'bootstrap-5',
                    placeholder: 'Selecione uma planta',
                    allowClear: true,
                    width: '100%'
                });
                
                $('#tipoProduto').select2({
                    theme: 'bootstrap-5',
                    placeholder: 'Selecione um tipo',
                    allowClear: true,
                    width: '100%'
                });
                
                $('#formularioFaixa').on('submit', function(e) {
                    e.preventDefault();
                    
                    if (!this.checkValidity()) {
                        e.stopPropagation();
                        $(this).addClass('was-validated');
                        return false;
                    }
                    
                    var quantidade = parseInt($('#quantidade').val());
                    if (quantidade < 1) {
                        alert('A quantidade deve ser maior que zero.');
                        return false;
                    }
                    
                    enviarFormulario();
                });
                
                function enviarFormulario() {
                    var $btnSalvar = $('#btnSalvar');
                    var $btnVoltar = $('#btnVoltar');
                    var textoOriginal = $btnSalvar.html();
                    
                    $btnSalvar.prop('disabled', true);
                    $btnVoltar.prop('disabled', true);
                    $btnSalvar.html('<span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>Salvar');
                    
                    var formData = {
                        idFaixa: $('#idFaixa').val(),
                        planta: $('#planta').val(),
                        tipoProduto: $('#tipoProduto').val(),
                        valorMinKg: $('#valorMinKg').val(),
                        valorMaxKg: $('#valorMaxKg').val(),
                        valorMinM2: $('#valorMinM2').val(),
                        valorMaxM2: $('#valorMaxM2').val(),
                        quantidade: $('#quantidade').val(),
                        acao: $('#acao').val()
                    };
                    
                    $.ajax({
                        url: 'OrcamentoFaixaAprovadorRepositorio.asp',
                        type: 'POST',
                        data: formData,
                        dataType: 'json',
                        success: function(response) {
                            if (response.sucesso) {
                                alert(response.mensagem);
                                window.location.href = 'OrcamentoFaixaAprovadorGrade.asp';
                            } else {
                                alert('Erro: ' + response.mensagem);
                                $btnSalvar.prop('disabled', false);
                                $btnVoltar.prop('disabled', false);
                                $btnSalvar.html(textoOriginal);
                            }
                        },
                        error: function(xhr, status, error) {
                            alert('Erro ao processar a requisição: ' + error);
                            $btnSalvar.prop('disabled', false);
                            $btnVoltar.prop('disabled', false);
                            $btnSalvar.html(textoOriginal);
                        }
                    });
                }
            });
        </script>
    </body>
</html>