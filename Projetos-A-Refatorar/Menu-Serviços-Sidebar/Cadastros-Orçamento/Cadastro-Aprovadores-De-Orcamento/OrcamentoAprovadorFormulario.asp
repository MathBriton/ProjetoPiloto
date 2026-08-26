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

Dim usuarios()
Dim plantas()
Dim idAprovador, acao
idAprovador = 0
acao = ""

If Request("idAprovador") <> "" Then
    idAprovador = CLng(Request("idAprovador"))
End If

If Request("acao") <> "" Then
    acao = Request("acao")
End If

'Variaveis auxiliares'
Dim idx
idx = 0

'----------------------------------------------
'          Carrega Lista de Usuarios
'----------------------------------------------

ReDim usuarios(0)
SQL = "SELECT CRM_R_E_C_N_O as IdUsuario, CRM_NOME as Nome FROM CRM_DTI WHERE CRM_D_E_L_E_T <> '*' ORDER BY CRM_NOME"

SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
While NOT PIR_HELPDESK.EOF
    ReDim Preserve usuarios(idx)
    Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
    usuarios(idx)("IdUsuario") = PIR_HELPDESK("IdUsuario")
    usuarios(idx)("Nome") = PIR_HELPDESK("Nome")
    idx = idx + 1
    PIR_HELPDESK.MoveNext()
Wend
If isObject(PIR_HELPDESK) Then
    PIR_HELPDESK.Close()
    SET PIR_HELPDESK = NOTHING
End If

'----------------------------------------------
'          Carrega Lista de plantas
'----------------------------------------------
idx = 0
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

'Variaveis para edição
Dim aprovadorNivel
Dim plantasSelecionadas
aprovadorNivel = ""
plantasSelecionadas = "[]"

'----------------------------------------------
'          Carrega Registro a ser editado (se for edição)
'----------------------------------------------

IF idAprovador > 0 AND acao = "edicao" THEN
    SQL = "SELECT TOP 1 a.ID_APROVADOR as IdAprovador, a.NIVEL as Nivel,"
    SQL = SQL & " ISNULL('[' + STUFF(("
    SQL = SQL & "    SELECT ',' + '""' + a2.PLANTA + '""'"
    SQL = SQL & "    FROM APROVADORES_ORCAMENTO a2"
    SQL = SQL & "    WHERE a2.ID_APROVADOR = a.ID_APROVADOR"
    SQL = SQL & "    AND (a2.EXCLUIDO = 0 OR a2.EXCLUIDO IS NULL)"
    SQL = SQL & "    ORDER BY a2.PLANTA"
    SQL = SQL & "    FOR XML PATH(''), TYPE"
    SQL = SQL & ").value('.', 'NVARCHAR(MAX)'), 1, 1, '') + ']', '[]') AS Plantas"
    SQL = SQL & " FROM APROVADORES_ORCAMENTO a"
    SQL = SQL & " WHERE a.ID_APROVADOR = " & idAprovador
    SQL = SQL & " AND (a.EXCLUIDO = 0 OR a.EXCLUIDO IS NULL)"
    
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
    If NOT PIR_CORPORATIVO.EOF Then
        idAprovador = PIR_CORPORATIVO("IdAprovador")
        aprovadorNivel = PIR_CORPORATIVO("Nivel")
        plantasSelecionadas = PIR_CORPORATIVO("Plantas")
    End If
    If isObject(PIR_CORPORATIVO) Then
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
END IF
%>

<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <title>Smurfit Kappa - Aprovadores de Orçamento</title>
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
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header border-bottom">
                <h3><% If acao = "edicao" Then Response.Write("Editar") Else Response.Write("Novo") End If %> Aprovador</h3>
            </div>
            
            <form name="formularioAprovador" id="formularioAprovador" method="post">
                <input type="hidden" name="acao" id="acao" value="<%= acao %>" />
                
                <div class="row mb-3">
                    <div class="col-8">
                        <label for="aprovador" class="form-label">Aprovador *</label>
                        <select <% If acao = "edicao" Then Response.Write("disabled") End If %> class="form-select" name="aprovador" id="aprovador" required>
                            <option value="">Selecione um usuário</option>
                            <%
                            For idx = 0 To UBound(usuarios)
                                Dim selected
                                selected = ""
                                If CStr(usuarios(idx)("IdUsuario")) = CStr(idAprovador) Then
                                    selected = " selected"
                                End If
                                Response.Write("<option value='" & usuarios(idx)("IdUsuario") & "'" & selected & ">" & usuarios(idx)("Nome") & "</option>")
                            Next
                            %>
                        </select>
                    </div>
                    
                    <div class="col-4">
                        <label for="nivel" class="form-label">Nível *</label>
                        <select class="form-select" name="nivel" id="nivel" required>
                            <option value="">Selecione</option>
                            <option value="1"<% If aprovadorNivel = "1" Then Response.Write(" selected") End If %>>Nível 1</option>
                            <option value="2"<% If aprovadorNivel = "2" Then Response.Write(" selected") End If %>>Nível 2</option>
                            <option value="3"<% If aprovadorNivel = "3" Then Response.Write(" selected") End If %>>Nível 3</option>
                        </select>
                    </div>
                </div>
                
                <div class="row mb-3">
                    <div class="col-12">
                        <label for="plantas" class="form-label">Plantas *</label>
                        <select class="form-select" name="plantas" id="plantas" multiple="multiple" required>
                            <%
                            For idx = 0 To UBound(plantas)
                                Response.Write("<option value='" & plantas(idx)("Codigo") & "'>" & plantas(idx)("Nome") & "</option>")
                            Next
                            %>
                        </select>
                        <small class="form-text text-muted">Selecione uma ou mais plantas</small>
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
                $('#plantas').select2({
                    theme: 'bootstrap-5',
                    placeholder: 'Selecione as plantas',
                    allowClear: true,
                    width: '100%'
                });
                
                $('#aprovador').select2({
                    theme: 'bootstrap-5',
                    placeholder: 'Selecione um usuário',
                    allowClear: true,
                    width: '100%'
                });
                
                <% If acao = "edicao" Then %>
                var plantasSelecionadas = <%= plantasSelecionadas %>;
                $('#plantas').val(plantasSelecionadas).trigger('change');
                <% End If %>
                
                $('#formularioAprovador').on('submit', function(e) {
                    e.preventDefault();
                    
                    if (!this.checkValidity()) {
                        e.stopPropagation();
                        $(this).addClass('was-validated');
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
                    $btnSalvar.html('<span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span> Salvar');

                    var formData = {
                        idAprovador: $('#aprovador').val(),
                        nivel: $('#nivel').val(),
                        plantas: JSON.stringify($('#plantas').val()),
                        acao: $('#acao').val()
                    };
                    
                    $.ajax({
                        url: 'OrcamentoAprovadorRepositorio.asp',
                        type: 'POST',
                        data: formData,
                        dataType: 'json',
                        success: function(response) {
                            if (response.sucesso) {
                                alert(response.mensagem);
                                window.location.href = 'OrcamentoAprovadorGrade.asp';
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