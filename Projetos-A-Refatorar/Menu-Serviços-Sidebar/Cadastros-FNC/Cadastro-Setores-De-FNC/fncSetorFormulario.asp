<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../../cnn2/cnn2.asp"-->
<!--#Include File="../../../funcoes/funcoes.asp"-->
<!--#Include File="../../../includes/variaveis.asp"-->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer  = true
Response.Expires = 0
Session.LCID = 1046

Dim plantas()
Dim idSetor, acao, descricao
idSetor = 0
acao = ""
descricao = ""

If Request("idSetor") <> "" Then
    idSetor = CLng(Request("idSetor"))
End If

If Request("acao") <> "" Then
    acao = Request("acao")
End If

'Variaveis auxiliares'
Dim idx
idx = 0

'----------------------------------------------
'          Carrega Registro a ser editado (se for edição)
'----------------------------------------------

IF idSetor > 0 AND acao = "edicao" THEN
    SQL = "SELECT TOP 1 * FROM Setor WHERE Id = " & idSetor
    
    SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
    If NOT BEN_CORPORATIVO.EOF Then
        idSetor = BEN_CORPORATIVO("Id")
        descricao = BEN_CORPORATIVO("Descricao")
    End If
    If isObject(BEN_CORPORATIVO) Then
        BEN_CORPORATIVO.Close()
        SET BEN_CORPORATIVO = NOTHING
    End If
END IF
%>

<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <title>Smurfit Kappa - Setores de FNC</title>
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
                <h3><% If acao = "edicao" Then Response.Write("Editar") Else Response.Write("Novo") End If %> Setor de FNC</h3>
            </div>
            
            <form name="formularioSetor" id="formularioSetor" method="post">
                <input type="hidden" name="acao" id="acao" value="<%= acao %>" />
                <input type="hidden" name="idSetor" id="idSetor" value="<%= idSetor %>" />
                
                <div class="row mb-3">
                    <div class="col-12">
                        <label for="descricao" class="form-label">Descrição *</label>
                        <input class="form-control" name="descricao" id="descricao" value="<%= descricao %>" required />
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
                $('#formularioSetor').on('submit', function(e) {
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
                    $btnSalvar.html('<span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>Salvar');
                    
                    var formData = {
                        idSetor: $('#idSetor').val(),
                        descricao: $('#descricao').val(),
                        acao: $('#acao').val()
                    };
                    
                    $.ajax({
                        url: 'fncSetorRepositorio.asp',
                        type: 'POST',
                        data: formData,
                        dataType: 'json',
                        success: function(response) {
                            if (response.sucesso) {
                                alert(response.mensagem);
                                window.location.href = 'fncSetorGrade.asp';
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