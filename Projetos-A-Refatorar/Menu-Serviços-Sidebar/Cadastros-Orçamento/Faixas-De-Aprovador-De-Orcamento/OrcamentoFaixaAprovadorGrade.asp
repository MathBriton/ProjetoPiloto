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

Dim faixas()
Dim plantas()
Dim plantaFiltro

'Variaveis auxiliares'
Dim idx
idx = 0
totalFaixas = 0

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

ReDim faixas(0)
idx = 0

Dim clausulaPlanta
clausulaPlanta = ""
If plantaFiltro <> "" AND plantaFiltro <> "TODOS" Then
    clausulaPlanta = " AND PLANTA = '" & plantaFiltro & "'"
End If

SQL = "SELECT * FROM FAIXA_APROVADORES WHERE 1=1" & clausulaPlanta

SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
While NOT PIR_CORPORATIVO.EOF
    ReDim Preserve faixas(idx)
    Set faixas(idx) = Server.CreateObject("Scripting.Dictionary")
    faixas(idx)("IdFaixa") = PIR_CORPORATIVO("ID")
    faixas(idx)("Planta") = PIR_CORPORATIVO("PLANTA")
    faixas(idx)("TipoProduto") = PIR_CORPORATIVO("TIPO_PRODUTO")
    faixas(idx)("ValorMinKg") = PIR_CORPORATIVO("VALOR_MIN_KG")
    faixas(idx)("ValorMaxKg") = PIR_CORPORATIVO("VALOR_MAX_KG")
    faixas(idx)("ValorMinM2") = PIR_CORPORATIVO("VALOR_MIN_M2")
    faixas(idx)("ValorMaxM2") = PIR_CORPORATIVO("VALOR_MAX_M2")
    faixas(idx)("Quantidade") = PIR_CORPORATIVO("QUANTIDADE")
    idx = idx + 1
    PIR_CORPORATIVO.MoveNext()
Wend
If isObject(PIR_CORPORATIVO) Then
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
End If

totalFaixas = idx

Function ObterNomePlanta(codPlanta)
    Dim nomePlanta 
    nomePlanta = ""
    Select Case codPlanta
        Case "00"
            nomePlanta = "Pirapetinga"
        Case "01"
            nomePlanta = "Corporativo RJ"
        Case "02"
            nomePlanta = "Corporativo SP"
        Case "03"
            nomePlanta = "Corporativo"
        Case "04"
            nomePlanta = "Uberaba"
        Case "05"
            nomePlanta = "Bento"
        Case "06"
            nomePlanta = "Maranguape"
        Case "07"
            nomePlanta = "Saquarema"
        Case Else
            nomePlanta = "Desconhecido"
    End Select
    
    ObterNomePlanta = nomePlanta
End Function


Function ObterNomeProduto(codProduto)
    Dim nomeProduto 
    nomeProduto = ""

    Select Case codProduto
        Case "CX"
            nomeProduto = "CAIXA"
        Case "CH"
            nomeProduto = "CHAPA"
        Case "PA"
            nomeProduto = "PAPEL"
        Case "CJ"
            nomeProduto = "Conjunto"
        Case Else
            nomeProduto = "DESCONHECIDO"
    End Select
    
    ObterNomeProduto = nomeProduto
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

            function novaFaixa() {
                window.location.href = 'OrcamentoFaixaAprovadorFormulario.asp?acao=criacao';
            }

            function editarFaixa(idFaixa) {
                window.location.href = 'OrcamentoFaixaAprovadorFormulario.asp?acao=edicao&idFaixa=' + idFaixa;
            }

            function excluirFaixa(idFaixa) {
                if (confirm('Deseja realmente excluir esta faixa?')) {
                    $.ajax({
                        url: 'OrcamentoFaixaAprovadorRepositorio.asp',
                        type: 'POST',
                        data: {
                            idFaixa: idFaixa,
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
                    <h3>Faixas de Aprovadores de Orçamento</h3>
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

                    <button type="button" id="btnNovo" name="btnNovo" class="btn btn-sm btn-primary mb-1" title="Novo Faixa" onclick="novaFaixa()">
                        <i class="fa fa-plus me-2"></i>
                        NOVO
                    </button>
                </div>

                <table class="table table-bordered" name="Grade" id="Grade">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Planta</th>
                            <th>Tipo Produto</th>
                            <th>Valor min. p/kg</th>
                            <th>Valor máx. p/kg</th>
                            <th>Valor min. p/m2</th>
                            <th>Valor máx. p/m2</th>
                            <th>Qtd. Aprovadores</th>
                            <th style="width: 100px;">Ações</th>
                        </tr>              
                    </thead>

                    <% IF totalFaixas > 0 THEN %>
                    <tbody>
                        <% 
                        Dim i
                        For i = 0 To UBound(faixas)
                        %>
                            <tr>   
                                <td class="text-left"><%=faixas(i)("IdFaixa")%></td>
                                <td class="text-left"><%=ObterNomePlanta(faixas(i)("Planta"))%></td>
                                <td class="text-left"><%=ObterNomeProduto(faixas(i)("TipoProduto"))%></td>
                                <td class="text-left"><%=faixas(i)("ValorMinKg")%></td>
                                <td class="text-left"><%=faixas(i)("ValorMaxKg")%></td>
                                <td class="text-left"><%=faixas(i)("ValorMinM2")%></td>
                                <td class="text-left"><%=faixas(i)("ValorMaxM2")%></td>
                                <td class="text-left"><%=faixas(i)("Quantidade")%></td>
                                <td class="text-left">
                                    <button type="button" class="btn btn-sm btn-primary me-2" title="Editar Faixa" onclick="editarFaixa(<%=faixas(i)("IdFaixa")%>)">
                                        <i class="fa fa-pencil-alt"></i>
                                    </button>
                                    <button type="button" class="btn btn-sm btn-danger" title="Excluir Faixa" onclick="excluirFaixa(<%=faixas(i)("IdFaixa")%>)">
                                        <i class="fa fa-trash-alt"></i>
                                    </button>
                                </td>
                            </tr>
                        <% Next %>
                    </tbody>
                    <% END IF %>
                </table>
                <% IF totalFaixas = 0 THEN %>
                    <div class="sem-registro">
                        Nenhum registro encontrado
                    </div>
                <% END IF %>
            </div>
        </form>
    </body>
</html>