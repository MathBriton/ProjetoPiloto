<!DOCTYPE html>
<%
Server.ScriptTimeout = 1000
Session.LCID = 1046 
if Trim(request("excel")) = "1" then
    Response.ContentType="application/vnd.ms-excel"
    response.AddHeader "content-disposition", "inline; filename=relatorio.xls"
end if
%>
<html lang="pt-br" xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="utf-8" />
    <!--MOBILE DEVICE-->
    <title>Custo de produto - Analítico</title>   
     
    <!--#Include File = "../cnn/cnn.asp"-->
    <!--#Include File = "../funcoes/funcoes.asp"-->
    <!--#include file="../includes/variaveis.asp" -->

    <!-- BOOTSTRAP -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <!-- BOOTSTRAP -->
    <script src="../funcoes/funcoesJAVASCRIPT.js"></script>
    <script src="../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.js"></script>

    <script type="text/javascript">
        function imprima() {
            print();
        }
        function doOnLoad() {
            myCalendar = new dhtmlXCalendarObject(["DataIni", "DataFim"]);
            myCalendar.setDate(myCalendar.getDate());
            myCalendar.loadUserLanguage("br");
        }
        function Consular() {
            document.getElementById("excel").value = "0";
            MateriaisSecundariosNaoConferidos.submit();
        }
    </script>

    <style>
        .margimBody {
            margin-top:50px;
        }
        .LinhaCinza {
            background-color:gray;
        }
    </style>

</head>
<%  

if trim(Request("filial")) <> "" then
    if trim(Request("filial")) = "P" then
        Filial = "Pirapetinga"
    else
        Filial = "Uberaba"
    end if
end if

SQL = "SELECT IdNota, Codigo, Razao, NotaFiscal, Serie,Item, DataLancamento, Material,  Descricao,Quantidade, Filial, PONTOS, PESO, ENCONTRADO, CO_PONTOS, UM FROM VwNotasAnalises "
SQL = SQL &"WHERE PONTOS IS NULL AND DataLancamento BETWEEN '"&InverteData(Trim(Request("DataIni")))&"' AND '"&InverteData(Trim(Request("DataFim")))&"' AND "

if Trim(Request("Fornecedor")) <> "" then
    SQL = SQL &"    Codigo = '"&Trim(Request("Fornecedor"))&"' AND "
end if

if Trim(Request("filial")) <> "" then
    if Trim(Request("filial")) = "P" then
        SQL = SQL &"    Filial IN ('00','05','') AND "
    else
        SQL = SQL &"    Filial NOT IN ('00','05','') AND "
    end if
end if

SQL = SQL &" 1=1 "
SQL = SQL &"ORDER BY DataLancamento "

%>
<body class="body margimBody" onload="doOnLoad()">
    <div class="container-fluid">
        <div class="panel panel-primary">
            <div class="panel-heading"><h3 class="text-center">Materiais Secundarios não conferidos</h3></div>
            <div class="panel-body">
                <form name="MateriaisSecundariosNaoConferidos" id="MateriaisSecundariosNaoConferidos" action="MateriaisSecundariosNaoConferidos.asp" method="post" role="form" >
                    <input type="hidden" name="excel" id="excel" value="0" />
                    <div class="row">
                        <div class="form-group-sm">
                            <div class="col-lg-2">
                                <label for="Filial">Filial:</label>
                                <select name="Filial" id="Filial" class="form-control">
                                    <option value="" <%if Trim(Request("filial")) = "" then%> selected="selected" <%end if %>>Selecione</option>
                                    <option value="P" <%if Trim(Request("filial")) = "P" then%> selected="selected" <%end if %>>Pirapetinga</option>
                                    <option value="U" <%if Trim(Request("filial")) = "U" then%> selected="selected" <%end if %>>Uberaba</option>
                                </select>
                            </div>
                            <div class="col-lg-2">
                                <label for="DataIni">Período:</label>
                                <input type="text" name="DataIni" id="DataIni" value="<%=Trim(Request("DataIni")) %>" class="form-control" maxlength="10" placeholder="Inicial" />&nbsp;
                                <input type="text" name="DataFim" id="DataFim" value="<%=Trim(Request("DataFim")) %>" class="form-control" maxlength="10" placeholder="Final"/>
                            </div>

                            <div class="col-lg-2">
                                <label for="Fornecedor">Fornecedor:</label>
                                <input type="text" name="Fornecedor" id="Fornecedor" value="<%=Trim(Request("Fornecedor")) %>" class="form-control" />
                            </div>
                            <div class="col-lg-6">
                                <br />     
                                <input type="button" name="consultar" id="consultar" value="Consultar" onclick="Consular()" class="btn btn-primary btn-sm" />
                                &nbsp;
                                <a href="#" class="btn btn-primary btn-sm" onclick="document.getElementById('excel').value = '1'; MateriaisSecundariosNaoConferidos.submit()" >
                                    <span class="glyphicon glyphicon-export"></span> Excel 
                                </a> 
                                &nbsp;
                                <a href="#" class="btn btn-primary btn-sm" onclick="imprima()" >
                                    <span class="glyphicon glyphicon-print"></span> Imprimir 
                                </a>   
                            </div>
                        </div>

                    </div>
                </form>
            </div>
        </div>
        <div class="row">
            <div class="col-sm-12">
                <table class="table table-striped table-condensed table-hover">
                    <thead>
                        <tr>
                            <th>Filial</th>
                            <th>Fornecedor</th>
                            <th>Nota Fiscal</th>
                            <th>Série</th>
                            <th>Item</th>
                            <th>Codigo</th>
                            <th>Descricao</th>
                            <th>Quantidade</th>
                            <th>UM</th>
                        </tr>
                    </thead>
                    <tbody>
<% 
                if (trim(Request("DataIni")) <> "" and Trim(Request("DataFim")) <> "") then
                    
                    Dados11.source = SQL
                    Dados11.Open()

                    if not Dados11.Eof then
                        Entrada = Trim(Dados11("DataLancamento"))
                    end if

                    x = 0
                    z = 0
                    Quantidade = 0
                    Valor = 0
                    QuantidadeTotal = 0
                    ValorTotal = 0
                    While not Dados11.Eof

                        select case Dados11("Filial")
                            case "","00","05"
                                Filial= "Pirapetinga"
                            case Else
                                Filial = "Uberaba"
                        end select 

                        If x = 0 then
                            response.write "<tr class='success'>"
                            response.write "<td colspan='9'>"
                            response.write "<b>"&AcertaData(Trim(Dados11("DataLancamento")))&"</b>"
                            response.write "</td>"
                            response.write "</tr>"                
                        end if

                        if Entrada <> Trim(Dados11("DataLancamento")) then
                            response.write "<tr style='background-color:LightGray'>"
                            response.write "<th colspan='7'>"
                            response.write "<strong>Total para o dia "&AcertaData(Entrada)&" "&z&" Nota(s) Fiscai(s)</strong>."
                            response.write "</th>"
                            response.write "<th class='text-right' title='Quantidade'>"
                            response.write "<strong>"&formatnumber(Quantidade,0)&"</strong>"
                            response.write "</th>"
                            response.write "<th>"
                            response.write "</th>"
                            response.write "</tr>"     

                            response.write "<tr class='success'>"
                            response.write "<td colspan='9'>"
                            response.write "<b>"&AcertaData(Trim(Dados11("DataLancamento")))&"</b>"
                            response.write "</td>"
                            response.write "</tr>"    
 
                            QuantidadeTotal = QuantidadeTotal + Quantidade
                            ValorTotal = ValorTotal + Valor
                            Quantidade = 0
                            Valor = 0 
                            z = 0 
                        end if
%>
                        <tr>
                            <td class="text-center" title="Filial"><%=Filial %></td>
                            <td title="Fornecedor"><%=Trim(Dados11("Codigo")) %>&nbsp;<%=Trim(Dados11("Razao")) %></td>
                            <td class="text-right" title="Nota Fiscal"><%=Dados11("NotaFiscal") %> </td>
                            <td class="text-center" title="Série"><%=Dados11("Serie") %> </td>
                            <td class="text-center" title="Item NF"><%=Dados11("Item") %> </td>
                            <td class="text-center" title="Produto"><%=Dados11("Material") %> </td>
                            <td title="Descrição"><%=Dados11("Descricao") %> </td>
                            <td class="text-right" title="Quantidade"><%=FormatNumber(Dados11("Quantidade"),0) %> </td>
                            <td title="Unidade de Medida"><%=Dados11("UM") %> </td>
                        </tr>
<%
                        x = x + 1
                        z = z + 1
                        Quantidade = Quantidade + Dados11("Quantidade")
                        Entrada = Dados11("DataLancamento")
                        response.Flush()
                        Dados11.MoveNext()
                    Wend
                    Dados11.Close()
                end if
                response.write "<tr style='background-color:LightGray'>"
                response.write "<th colspan='7'>"
                response.write "<strong>Total para o dia "&AcertaData(Entrada)&" "&z&" Nota(s) Fiscai(s)</strong>."
                response.write "</th>"
                response.write "<th class='text-right' title='Quantidade'>"
                response.write "<strong>"&formatnumber(Quantidade,0)&"</strong>"
                response.write "</th>"
                response.write "<th>"
                response.write "</th>"
                response.write "</tr>"     
 
                QuantidadeTotal = QuantidadeTotal + Quantidade
                ValorTotal = ValorTotal + Valor
                Quantidade = 0
                Valor = 0 
                z = 0 

                response.write "<tfoot>"
                response.write "<tr class='warning'>"
                response.write "<th colspan='7'>"
                response.write "<strong>Total de "&x&" Nota(s) Fiscai(s)</strong>."
                response.write "</th>"
                response.write "<th class='text-right' title='Quantidade'>"
                response.write "<strong>"&formatnumber(QuantidadeTotal,0)&"</strong>"
                response.write "</th>"
                response.write "<th>"
                response.write "</th>"
                response.write "</tr>"
                response.write "</tfoot>"

                QuantidadeTotal = 0
                ValorTotal = 0
                x = 0
%>
                </table>
            </div>
        </div>
    </div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
</body>
</html>