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

Dim defeitos()

'Variaveis auxiliares'
Dim idx
idx = 0
totalDefeitos = 0

'----------------------------------------------
'         Carrega Grade principal
'----------------------------------------------

ReDim defeitos(0)
idx = 0

SQL = "SELECT d.*, s.Descricao as NomeSetor FROM Defeito d " 
SQL = SQL & "LEFT JOIN Setor s ON d.IdSetor = s.Id "
SQL = SQL & "WHERE (d.FlagInativo = 0 OR d.FlagInativo is null)"

SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
While NOT BEN_CORPORATIVO.EOF
    ReDim Preserve defeitos(idx)
    Set defeitos(idx) = Server.CreateObject("Scripting.Dictionary")
    defeitos(idx)("IdDefeito") = BEN_CORPORATIVO("Id")
    defeitos(idx)("Tipo") = BEN_CORPORATIVO("Tipo")
    defeitos(idx)("NomeSetor") = BEN_CORPORATIVO("NomeSetor")
    defeitos(idx)("Descricao") = BEN_CORPORATIVO("Descricao")
    idx = idx + 1
    BEN_CORPORATIVO.MoveNext()
Wend
If isObject(BEN_CORPORATIVO) Then
    BEN_CORPORATIVO.Close()
    SET BEN_CORPORATIVO = NOTHING
End If

totalDefeitos = idx

Function ObterNomeTipo(nivel)
    If nivel = "" Then
        ObterNomeTipo = ""
        Exit Function
    End If

	Select Case nivel
		Case "PRO"
			nomeTipo = "Industrial"
		Case "ADM"
			nomeTipo = "Administrativo"
		Case Else
			nomeTipo = "Desconhecido"
	End Select
    
    ObterNomeTipo = nomeTipo
End Function
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

            function novoDefeito() {
                window.location.href = 'fncDefeitoFormulario.asp?acao=criacao';
            }

            function editarDefeito(idDefeito) {
                window.location.href = 'fncDefeitoFormulario.asp?acao=edicao&idDefeito=' + idDefeito;
            }

            function excluirDefeito(idDefeito) {
                if (confirm('Deseja realmente excluir este defeito?')) {
                    $.ajax({
                        url: 'fncDefeitoRepositorio.asp',
                        type: 'POST',
                        data: {
                            idDefeito: idDefeito,
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
                    <h3>Defeitos de FNC</h3>
                </div>

                <div class="filtro-container">
                    <button type="button" id="btnNovo" name="btnNovo" class="btn btn-sm btn-primary mb-1" title="Novo Defeito" onclick="novoDefeito()">
                        <i class="fa fa-plus me-2"></i>
                        NOVO
                    </button>
                </div>

                <table class="table table-bordered" name="Grade" id="Grade">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Tipo</th>
                            <th>Setor</th>
                            <th>Descrição</th>
                            <th style="width: 100px;">Ações</th>
                        </tr>              
                    </thead>

                    <% IF totalDefeitos > 0 THEN %>
                    <tbody>
                        <% 
                        Dim i
                        For i = 0 To UBound(defeitos)
                        %>
                            <tr>   
                                <td class="text-left"><%=defeitos(i)("IdDefeito")%></td>
                                <td class="text-left"><%=ObterNomeTipo(defeitos(i)("Tipo"))%></td>
                                <td class="text-left"><%=defeitos(i)("NomeSetor")%></td>
                                <td class="text-left"><%=defeitos(i)("Descricao")%></td>
                                <td class="text-left">
                                    <button type="button" class="btn btn-sm btn-primary me-2" title="Editar Defeito" onclick="editarDefeito(<%=defeitos(i)("IdDefeito")%>)">
                                        <i class="fa fa-pencil-alt"></i>
                                    </button>
                                    <button type="button" class="btn btn-sm btn-danger" title="Excluir Defeito" onclick="excluirDefeito(<%=defeitos(i)("IdDefeito")%>)">
                                        <i class="fa fa-trash-alt"></i>
                                    </button>
                                </td>
                            </tr>
                        <% Next %>
                    </tbody>
                    <% END IF %>
                </table>
                <% IF totalDefeitos = 0 THEN %>
                    <div class="sem-registro">
                        Nenhum registro encontrado
                    </div>
                <% END IF %>
            </div>
        </form>
    </body>
</html>