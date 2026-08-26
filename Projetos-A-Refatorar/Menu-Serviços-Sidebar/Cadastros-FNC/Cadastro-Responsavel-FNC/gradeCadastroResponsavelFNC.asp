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

Dim responsaveis()
Dim plantas(), setores()
Dim plantaFiltro, setoresFiltro

'Variaveis auxiliares'
Dim idx
idx = 0
totalResponsaveis = 0

'----------------------------------------------
'          Carrega Items filtro de plantas
'----------------------------------------------

ReDim plantas(0)
SQL = "SELECT Empresa, Razao FROM Empresas"

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

'----------------------------------------------
'          Carrega Items filtro de setores
'----------------------------------------------


ReDim setores(0)
idx = 0

SQL = "SELECT * FROM Setor WHERE (FlagInativo = 0 OR FlagInativo is null)" 

SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
While NOT BEN_CORPORATIVO.EOF
    ReDim Preserve setores(idx)
    Set setores(idx) = Server.CreateObject("Scripting.Dictionary")
    setores(idx)("Id") = BEN_CORPORATIVO("Id")
    setores(idx)("Descricao") = BEN_CORPORATIVO("Descricao")
    idx = idx + 1
    BEN_CORPORATIVO.MoveNext()
Wend
If isObject(BEN_CORPORATIVO) Then
    BEN_CORPORATIVO.Close()
    SET BEN_CORPORATIVO = NOTHING
End If

'----------------------------------------------
'         Carrega Grade principal
'----------------------------------------------

' Captura os filtros
plantaFiltro = Request.Form("plantaFiltro")
If plantaFiltro = "" Then plantaFiltro = "TODOS"

setorFiltro = Request.Form("setorFiltro")
If setorFiltro = "" Then setorFiltro = "TODOS"

ReDim responsaveis(0)
idx = 0

Dim clausulaPlanta
clausulaPlanta = ""
If plantaFiltro <> "" AND plantaFiltro <> "TODOS" Then
    clausulaPlanta = " AND r.Planta = '" & plantaFiltro & "'"
End If

Dim clausulaSetor
clausulaSetor = ""
If setorFiltro <> "" AND setorFiltro <> "TODOS" Then
    clausulaSetor = " AND r.IdSetor = '" & setorFiltro & "'"
End If

SQL = "SELECT r.Id, r.Nome, r.Tipo, r.Email, s.Descricao as NomeSetor,"
SQL = SQL & " ISNULL('[' + STUFF(("
SQL = SQL & "    SELECT ',' + '""' + r2.Planta + '""'"
SQL = SQL & "    FROM Responsavel r2"
SQL = SQL & "    WHERE r2.Id = r.Id"
SQL = SQL & "    AND r2.Tipo = r.Tipo"
SQL = SQL & "    AND r2.Email = r.Email"
SQL = SQL & "    AND (r2.Excluido = 0 OR r2.Excluido IS NULL)"
SQL = SQL & "    ORDER BY r2.Planta"
SQL = SQL & "    FOR XML PATH(''), TYPE"
SQL = SQL & ").value('.', 'NVARCHAR(MAX)'), 1, 1, '') + ']', '[]') AS Plantas,"
SQL = SQL & " MIN(r.DataInclusao) AS DataInclusao"
SQL = SQL & " FROM Responsavel r"
SQL = SQL & " LEFT JOIN Setor s ON r.IdSetor = s.Id "
SQL = SQL & " WHERE (r.Excluido = 0 OR r.Excluido IS NULL)" & clausulaPlanta & clausulaSetor
SQL = SQL & " GROUP BY r.Id, r.Nome, r.Tipo, r.Email, r.IdSetor, s.Descricao"
SQL = SQL & " ORDER BY r.Nome" 

SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)
While NOT BEN_CORPORATIVO.EOF
    ReDim Preserve responsaveis(idx)
    Set responsaveis(idx) = Server.CreateObject("Scripting.Dictionary")
    responsaveis(idx)("IdResponsavel") = BEN_CORPORATIVO("Id")
    responsaveis(idx)("NomeSetor") = BEN_CORPORATIVO("NomeSetor")
    responsaveis(idx)("NomeResponsavel") = BEN_CORPORATIVO("Nome")
	responsaveis(idx)("Email") = BEN_CORPORATIVO("Email")
    responsaveis(idx)("Tipo") = BEN_CORPORATIVO("Tipo")
    responsaveis(idx)("Plantas") = BEN_CORPORATIVO("Plantas")
    responsaveis(idx)("DataInclusao") = BEN_CORPORATIVO("DataInclusao")
    idx = idx + 1
    BEN_CORPORATIVO.MoveNext()
Wend
If isObject(BEN_CORPORATIVO) Then
    BEN_CORPORATIVO.Close()
    SET BEN_CORPORATIVO = NOTHING
End If

totalResponsaveis = idx

Function ObterNomeTipo(tipo)
    If tipo = "" Then
        ObterNomeTipo = ""
        Exit Function
    End If

	Select Case tipo
		Case "1"
			nomeTipo = "Analista"
		Case "2"
			nomeTipo = "Reponsável Setorial"
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

            function novoResponsavel() {
                window.location.href = 'fncCadastroResponsavelFNC.asp?acao=criacao';
            }

            function editarResponsavel(IdResponsavel) {
                window.location.href = 'fncCadastroResponsavelFNC.asp?acao=edicao&IdResponsavel=' + IdResponsavel;
            }

            function excluirResponsavel(IdResponsavel) {
                if (confirm('Deseja realmente excluir este responsável?')) {
                    $.ajax({
                        url: 'fncCadastroResponsavelRepositorio.asp',
                        type: 'POST',
                        data: {
                            IdResponsavel: IdResponsavel,
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

            $(document).ready(function() {
                $('#plantaFiltro').on('change', function() {
                    $('#Pesquisar').submit();
                });

                $('#setorFiltro').on('change', function() {
                    $('#Pesquisar').submit();
                });
            });

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
            .filtro-container {
                display: flex;
                gap: 10px;
                align-items: flex-end;
                padding: 10px;
                margin-bottom: 10px;
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 0.25rem;
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
          
        </style>
    </head>
    <body>
        <form name="Pesquisar" id="Pesquisar" method="post">        
            <div class="container">
                <div class="header border-bottom">
                    <h3>Responsáveis de FNC</h3>
                </div>

                <div class="filtro-container">
                    <div class="col-md-4">
                        <label for="plantaFiltro" class="form-label"><strong>Planta:</strong></label>
                        <select name="plantaFiltro" id="plantaFiltro" class="form-select form-select-sm">
                            <option value="TODOS" <%If Request.Form("plantaFiltro") = "TODOS" OR Request.Form("plantaFiltro") = "" Then Response.Write("selected")%>>Todas Plantas</option>
                            <% 
                            idx = 0
                            For idx = 0 To UBound(plantas)
                                Dim isSelected
                                isSelected = ""
                                If Request.Form("plantaFiltro") = plantas(idx)("Codigo") Then
                                    isSelected = "selected"
                                End If
                            %>
                                <option value="<%=plantas(idx)("Codigo")%>" <%=isSelected%>><%=plantas(idx)("Nome")%></option>
                            <% Next %>
                        </select>
                    </div>

                    <div class="col-md-4">
                        <label for="setorFiltro" class="form-label"><strong>Setor:</strong></label>
                        <select name="setorFiltro" id="setorFiltro" class="form-select form-select-sm">
                            <option value="TODOS" <%If Request.Form("setorFiltro") = "TODOS" OR Request.Form("setorFiltro") = "" Then Response.Write("selected")%>>Todos Setores</option>
                            <% 
                            idx = 0
                            For idx = 0 To UBound(setores)
                                isSelected = ""
                                If Request.Form("setorFiltro") = CStr(setores(idx)("Id")) Then
                                    isSelected = "selected"
                                End If
                            %>
                                <option value="<%=setores(idx)("Id")%>" <%=isSelected%>><%=setores(idx)("Descricao")%></option>
                            <% Next %>
                        </select>
                    </div>

                    <button type="button" id="btnNovo" name="btnNovo" class="btn btn-sm btn-primary mb-1" title="Novo Responsável" onclick="novoResponsavel()">
                        <i class="fa fa-plus me-2"></i>
                        NOVO
                    </button>
                </div>

                <table class="table table-bordered" name="Grade" id="Grade">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Setor</th>
                            <th>Nome</th>
							<th>Email</th>
                            <th>Tipo</th>
                            <th>Plantas</th>
                            <th>Data de Cadastro</th>
                            <th style="width: 100px;">Ações</th>
                        </tr>              
                    </thead>

                    <% IF totalResponsaveis > 0 THEN %>
                    <tbody>
                        <% 
                        Dim i
                        For i = 0 To UBound(responsaveis)
                        %>
                            <tr>   
                                <td class="text-left"><%=responsaveis(i)("IdResponsavel")%></td>
                                <td class="text-left"><%=responsaveis(i)("NomeSetor")%></td>
                                <td class="text-left"><%=responsaveis(i)("NomeResponsavel")%></td>
								<td class="text-left"><%=responsaveis(i)("Email")%></td>
                                <td class="text-left"><%=ObterNomeTipo(responsaveis(i)("Tipo"))%></td>
                                <td class="text-left"><%=ObterNomePlantas(responsaveis(i)("Plantas"))%></td>
                                <td class="text-left"><%=responsaveis(i)("DataInclusao")%></td>
                                <td class="text-left">
                                    <button type="button" class="btn btn-sm btn-primary" title="Editar Responsável" onclick="editarResponsavel(<%=responsaveis(i)("IdResponsavel")%>)">
                                        <i class="fa fa-pencil-alt"></i>
                                    </button>
                                    <button type="button" class="btn btn-sm btn-danger" title="Excluir Responsável" onclick="excluirResponsavel(<%=responsaveis(i)("IdResponsavel")%>)">
                                        <i class="fa fa-trash-alt"></i>
                                    </button>
                                </td>
                            </tr>
                        <% Next %>
                    </tbody>
                    <% END IF %>
                </table>
                <% IF totalResponsaveis = 0 THEN %>
                    <div class="sem-registro">
                        Nenhum registro encontrado
                    </div>
                <% END IF %>
            </div>
        </form>
    </body>
</html>