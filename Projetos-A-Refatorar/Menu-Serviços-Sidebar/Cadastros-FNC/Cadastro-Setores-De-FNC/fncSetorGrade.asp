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

Dim setores()

'Variaveis auxiliares'
Dim idx
idx = 0
totalSetores = 0

'----------------------------------------------
'         Carrega Grade principal
'----------------------------------------------

ReDim setores(0)
idx = 0

SQL = "SELECT * FROM Setor WHERE (FlagInativo = 0 OR FlagInativo is null)" 

SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
While NOT BEN_CORPORATIVO.EOF
    ReDim Preserve setores(idx)
    Set setores(idx) = Server.CreateObject("Scripting.Dictionary")
    setores(idx)("IdSetor") = BEN_CORPORATIVO("Id")
    setores(idx)("Descricao") = BEN_CORPORATIVO("Descricao")
    idx = idx + 1
    BEN_CORPORATIVO.MoveNext()
Wend
If isObject(BEN_CORPORATIVO) Then
    BEN_CORPORATIVO.Close()
    SET BEN_CORPORATIVO = NOTHING
End If

totalSetores = idx
%>


<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <title>Smurfit Kappa</title>
        <meta charset="utf-8"/>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
        <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />    
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js"></script>
       <script type="text/javascript">
            function Ajax() {
                var xmlHttp = null;
                if (window.XMLHttpRequest) {
                    xmlHttp = new XMLHttpRequest();
                }
                else if (window.ActiveXObject) {
                    xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
                }
                else {
                    alert("Seu navegador não suporta AJAX");
                }
                return xmlHttp;
            }

            function novoSetor() {
                window.location.href = 'fncSetorFormulario.asp?acao=criacao';
            }

            function editarSetor(idSetor) {
                window.location.href = 'fncSetorFormulario.asp?acao=edicao&idSetor=' + idSetor;
            }

            function excluirSetor(idSetor) {
                if (confirm('Deseja realmente excluir este setor?')) {
                    $.ajax({
                        url: 'fncSetorRepositorio.asp',
                        type: 'POST',
                        data: {
                            idSetor: idSetor,
                            acao: 'exclusao'
                        },
                        dataType: 'json',
                        success: function(response) {
                            if (response.sucesso) {
                                alert(response.mensagem);
                                location.reload();
                            } else {
                                alert('Erro: ' + response.mensagem);
                            }
                        },
                        error: function(xhr, status, error) {
                            alert('Erro ao processar a requisição: ' + error);
                        }
                    });
                }
            }
        </script>
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
            tr:nth-child(even) {
                background-color: #f8f9fa;
            }
            table {
                margin-bottom: 0px !important;
            }
            thead {
                background-color: #f8f9fa;
            }
            .sem-registro {
                border-bottom: solid #dee2e6 1px;
                border-left: solid #dee2e6 1px;
                padding: 10px;
                border-right: solid #dee2e6 1px;
                text-align: center;
            }
            .filtro-container {
                display: flex;
                justify-content: flex-end;
                align-items: flex-end;
                padding: 10px;
                margin-bottom: 10px;
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 0.25rem;
            }
          
        </style>
    </head>
    <body>
        <form name="Pesquisar" id="Pesquisar" method="post">        
            <div class="container">
                <div class="header border-bottom">
                    <h3>Setores de FNC</h3>
                </div>

                <div class="filtro-container">
                    <button type="button" id="btnNovo" name="btnNovo" class="btn btn-sm btn-primary mb-1" title="Novo Setor" onclick="novoSetor()">
                        <i class="fa fa-plus me-2"></i>
                        NOVO
                    </button>
                </div>

                <table class="table table-bordered" name="Grade" id="Grade">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Descrição</th>
                            <th style="width: 100px;">Ações</th>
                        </tr>              
                    </thead>

                    <% IF totalSetores > 0 THEN %>
                    <tbody>
                        <% 
                        Dim i
                        For i = 0 To UBound(setores)
                        %>
                            <tr>   
                                <td class="text-left"><%=setores(i)("IdSetor")%></td>
                                <td class="text-left"><%=setores(i)("Descricao")%></td>
                                <td class="text-left">
                                    <button type="button" class="btn btn-sm btn-primary me-2" title="Editar Setor" onclick="editarSetor(<%=setores(i)("IdSetor")%>)">
                                        <i class="fa fa-pencil-alt"></i>
                                    </button>
                                    <button type="button" class="btn btn-sm btn-danger" title="Excluir Setor" onclick="excluirSetor(<%=setores(i)("IdSetor")%>)">
                                        <i class="fa fa-trash-alt"></i>
                                    </button>
                                </td>
                            </tr>
                        <% Next %>
                    </tbody>
                    <% END IF %>
                </table>
                <% IF totalSetores = 0 THEN %>
                    <div class="sem-registro">
                        Nenhum registro encontrado
                    </div>
                <% END IF %>
            </div>
        </form>
    </body>
</html>