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

Dim analistas()
Dim plantas()
Dim plantaFiltro

'Variaveis auxiliares'
Dim idx
idx = 0
totalAnalistas = 0

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
'         Carrega Grade principal
'----------------------------------------------

' Captura os filtros
plantaFiltro = Request.Form("plantaFiltro")
If plantaFiltro = "" Then plantaFiltro = "TODOS"

ReDim analistas(0)
idx = 0

Dim clausulaPlanta
clausulaPlanta = ""
If plantaFiltro <> "" AND plantaFiltro <> "TODOS" Then
    clausulaPlanta = " AND a.PLANTA = '" & plantaFiltro & "'"
End If

SQL = "SELECT a.CODIGO_USUARIO, a.NOME_USUARIO, a.EMAIL,"
SQL = SQL & " ISNULL('[' + STUFF(("
SQL = SQL & "    SELECT ',' + '""' + a2.PLANTA + '""'"
SQL = SQL & "    FROM ANALISTAS_ORCAMENTO a2"
SQL = SQL & "    INNER JOIN HELPDESK..Empresas e2 ON a2.PLANTA = e2.Empresa"
SQL = SQL & "    WHERE a2.CODIGO_USUARIO = a.CODIGO_USUARIO"
SQL = SQL & "    AND (a2.EXCLUIDO = 0 OR a2.EXCLUIDO IS NULL)"
SQL = SQL & "    ORDER BY a2.PLANTA"
SQL = SQL & "    FOR XML PATH(''), TYPE"
SQL = SQL & ").value('.', 'NVARCHAR(MAX)'), 1, 1, '') + ']', '[]') AS Plantas,"
SQL = SQL & " MIN(a.DATA_INCLUSAO) AS DATA_INCLUSAO"
SQL = SQL & " FROM ANALISTAS_ORCAMENTO a"
SQL = SQL & " LEFT JOIN HELPDESK..Empresas e ON a.PLANTA = e.Empresa"
SQL = SQL & " WHERE (a.EXCLUIDO = 0 OR a.EXCLUIDO IS NULL)" & clausulaPlanta
SQL = SQL & " GROUP BY a.CODIGO_USUARIO, a.NOME_USUARIO, a.EMAIL"
SQL = SQL & " ORDER BY a.NOME_USUARIO" 

SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
While NOT PIR_CORPORATIVO.EOF
    ReDim Preserve analistas(idx)
    Set analistas(idx) = Server.CreateObject("Scripting.Dictionary")
    analistas(idx)("IdAnalista") = PIR_CORPORATIVO("CODIGO_USUARIO")
    analistas(idx)("NomeAnalista") = PIR_CORPORATIVO("NOME_USUARIO")
    analistas(idx)("Email") = PIR_CORPORATIVO("EMAIL")
    analistas(idx)("Plantas") = PIR_CORPORATIVO("PLANTAS")
    analistas(idx)("DataInclusao") = PIR_CORPORATIVO("DATA_INCLUSAO")
    idx = idx + 1
    PIR_CORPORATIVO.MoveNext()
Wend
If isObject(PIR_CORPORATIVO) Then
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
End If

totalAnalistas = idx

Function ObterNomePlantas(plantasJSON)
    If IsNull(plantasJSON) OR plantasJSON = "" OR plantasJSON = "[]" Then
        ObterNomePlantas = ""
        Exit Function
    End If
    
    Dim plantasStr
    plantasStr = Replace(plantasJSON, "[", "")
    plantasStr = Replace(plantasStr, "]", "")
    plantasStr = Replace(plantasStr, """", "")
    
    Dim plantasArray
    plantasArray = Split(plantasStr, ",")
    
    Dim i, codigo, nomes()
    ReDim nomes(UBound(plantasArray))
    
    For i = 0 To UBound(plantasArray)
        codigo = Trim(plantasArray(i))
        Select Case codigo
            Case "00"
                nomes(i) = "Pirapetinga"
            Case "01"
                nomes(i) = "Corporativo RJ"
            Case "02"
                nomes(i) = "Corporativo SP"
            Case "03"
                nomes(i) = "Corporativo"
            Case "04"
                nomes(i) = "Uberaba"
            Case "05"
                nomes(i) = "Bento"
            Case "06"
                nomes(i) = "Maranguape"
            Case "07"
                nomes(i) = "Saquarema"
            Case Else
                nomes(i) = "Desconhecido"
        End Select
    Next
    
    ObterNomePlantas = Join(nomes, ", ")
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

            function novoAnalista() {
                window.location.href = 'OrcamentoAnalistaFormulario.asp?acao=criacao';
            }

            function editarAnalista(idAnalista) {
                window.location.href = 'OrcamentoAnalistaFormulario.asp?acao=edicao&idAnalista=' + idAnalista;
            }

            function excluirAnalista(idAnalista) {
                if (confirm('Deseja realmente excluir este analista?')) {
                    $.ajax({
                        url: 'OrcamentoAnalistaRepositorio.asp',
                        type: 'POST',
                        data: {
                            idAnalista: idAnalista,
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
                justify-content: space-between;
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
                    <h3>Analistas de Orçamento</h3>
                </div>

                <div class="filtro-container">
                    <div class="col-md-4">
                        <label for="plantaFiltro" class="form-label"><strong>Planta:</strong></label>
                        <select name="plantaFiltro" id="plantaFiltro" class="form-select form-select-sm">
                            <option value="TODOS" <%If Request.Form("plantaFiltro") = "TODOS" OR Request.Form("plantaFiltro") = "" Then Response.Write("selected")%>>Todas Plantas</option>
                            <% 
                            Dim j
                            For j = 0 To UBound(plantas)
                                Dim isSelected
                                isSelected = ""
                                If Request.Form("plantaFiltro") = plantas(j)("Codigo") Then
                                    isSelected = "selected"
                                End If
                            %>
                                <option value="<%=plantas(j)("Codigo")%>" <%=isSelected%>><%=plantas(j)("Nome")%></option>
                            <% Next %>
                        </select>
                    </div>

                    <button type="button" id="btnNovo" name="btnNovo" class="btn bgn-sm btn-primary mb-1" title="Novo Analista" onclick="novoAnalista()">
                        <i class="fa fa-plus me-2"></i>
                        NOVO
                    </button>
                </div>

                <table class="table table-bordered" name="Grade" id="Grade">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Nome</th>
                            <th>Email</th>
                            <th>Plantas</th>
                            <th>Data de Cadastro</th>
                            <th style="width: 100px;">Ações</th>
                        </tr>              
                    </thead>

                    <% IF totalAnalistas > 0 THEN %>
                    <tbody>
                        <% 
                        Dim i
                        For i = 0 To UBound(analistas)
                        %>
                            <tr>   
                                <td class="text-left"><%=analistas(i)("IdAnalista")%></td>
                                <td class="text-left"><%=analistas(i)("NomeAnalista")%></td>
                                <td class="text-left"><%=analistas(i)("Email")%></td>
                                <td class="text-left"><%=ObterNomePlantas(analistas(i)("Plantas"))%></td>
                                <td class="text-left"><%=analistas(i)("DataInclusao")%></td>
                                <td class="text-left">
                                    <button type="button" class="btn btn-sm btn-primary" title="Editar Analista" onclick="editarAnalista(<%=analistas(i)("IdAnalista")%>)">
                                        <i class="fa fa-pencil-alt"></i>
                                    </button>
                                    <button type="button" class="btn btn-sm btn-danger me-2" title="Excluir Analista" onclick="excluirAnalista(<%=analistas(i)("IdAnalista")%>)">
                                        <i class="fa fa-trash-alt"></i>
                                    </button>
                                </td>
                            </tr>
                        <% Next %>
                    </tbody>
                    <% END IF %>
                </table>
                <% IF totalAnalistas = 0 THEN %>
                    <div class="sem-registro">
                        Nenhum registro encontrado
                    </div>
                <% END IF %>
            </div>
        </form>
    </body>
</html>