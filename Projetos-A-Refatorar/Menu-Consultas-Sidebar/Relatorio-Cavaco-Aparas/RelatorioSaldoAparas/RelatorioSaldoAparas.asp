<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../../cnn2/cnn2.asp" -->
<!--#include File = "../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer  = true
Response.Expires = 0
Session.LCID = 1046

Dim idUsuarioLogado
idUsuarioLogado = Session("CODIGO")

' Filtros
Dim filtroCodERP
filtroCodERP = Request.Form("filtroCodERP")

' Totais
Dim totalSaldo, totalEntradas, totalSaidas, totalPedidosValidados, totalPedidosValidar
totalSaldo = 0
totalEntradas = 0
totalSaidas = 0
totalPedidosValidados = 0
totalPedidosValidar = 0

Dim abaSelecionada
abaSelecionada = Request.Form("aba")
If abaSelecionada = "" Then
    abaSelecionada = "entradas"
End If

Dim dadosEntradas()
Dim dadosSaidas()
Dim dadosPedidosValidados()
Dim dadosPedidosValidar()

ReDim dadosEntradas(0)
ReDim dadosSaidas(0)
ReDim dadosPedidosValidados(0)
ReDim dadosPedidosValidar(0)

Dim idxEntradas, idxSaidas, idxPedidosValidados, idxPedidosValidar
idxEntradas = 0
idxSaidas = 0
idxPedidosValidados = 0
idxPedidosValidar = 0

'----------------------------------------------
'    Carrega dados conforme aba selecionada
'----------------------------------------------

If filtroCodERP <> "" Then
    Call CarregarTotais()

    Select Case abaSelecionada
        Case "entradas"
            Call CarregarAbaEntradas()
        Case "saidas"
            Call CarregarAbaSaidas()
        Case "pedidosvalidados"
            Call CarregarAbaPedidosValidados()
        Case "pedidosvalidar"
            Call CarregarAbaPedidosValidar()
    End Select
End If

Sub CarregarTotais()
    Dim SQL

    SQL = "SELECT ENTRADA_APARAS, SAIDA_NF, PEDIDOS_VALIDADOS, PEDIDOS_VALIDAR, SALDO_FINAL "
    SQL = SQL & "FROM FUNC_SALDO_APARAS_TABELA('" & filtroCodERP & "')"

    On Error Resume Next
    SET BEN_DADOSADV = Conecta("BEN-DADOSADV", SQL)
    
    If Err.Number = 0 AND IsObject(BEN_DADOSADV) Then
        While NOT BEN_DADOSADV.EOF
            totalSaldo = BEN_DADOSADV("SALDO_FINAL")
            totalEntradas = BEN_DADOSADV("ENTRADA_APARAS")
            totalSaidas = BEN_DADOSADV("SAIDA_NF")
            totalPedidosValidar = BEN_DADOSADV("PEDIDOS_VALIDAR")
            totalPedidosValidados = BEN_DADOSADV("PEDIDOS_VALIDADOS")
            BEN_DADOSADV.MoveNext()
        Wend
        
        If isObject(BEN_DADOSADV) Then
            BEN_DADOSADV.Close()
            SET BEN_DADOSADV = NOTHING
        End If
    End If
    On Error Goto 0
End Sub

Sub CarregarAbaEntradas()
    Dim SQL

    SQL = "SELECT NOME_CLIENTE, PRODUTO AS CODIGO_PRODUTO, CODIGO_CLIENTE, LOJA_CLIENTE, " 
    SQL = SQL & "DESCRICAO_PRODUTO, QUANTIDADE, EMISSAO AS DATA_EMISSAO, SERIE, DOCUMENTO, "
    SQL = SQL & "DATA_DIGITACAO FROM VW_ENTRADA_APARAS WHERE 1=1 "
    SQL = SQL & "AND CODIGO_ERP = '" & filtroCodERP & "' "

    On Error Resume Next
    SET BEN_DADOSADV = Conecta("BEN-DADOSADV", SQL)
    
    If Err.Number = 0 AND IsObject(BEN_DADOSADV) Then
        While NOT BEN_DADOSADV.EOF
            ReDim Preserve dadosEntradas(idxEntradas)
            Set dadosEntradas(idxEntradas) = Server.CreateObject("Scripting.Dictionary")
            dadosEntradas(idxEntradas)("NomeCliente") = BEN_DADOSADV("NOME_CLIENTE")
            dadosEntradas(idxEntradas)("CodigoCliente") = BEN_DADOSADV("CODIGO_CLIENTE")
            dadosEntradas(idxEntradas)("CodigoProduto") = BEN_DADOSADV("CODIGO_PRODUTO")
            dadosEntradas(idxEntradas)("DescricaoProduto") = BEN_DADOSADV("DESCRICAO_PRODUTO")
            dadosEntradas(idxEntradas)("Serie") = BEN_DADOSADV("SERIE")
            dadosEntradas(idxEntradas)("Documento") = BEN_DADOSADV("DOCUMENTO")
            dadosEntradas(idxEntradas)("Quantidade") = BEN_DADOSADV("QUANTIDADE")
            dadosEntradas(idxEntradas)("DataDigitacao") = BEN_DADOSADV("DATA_DIGITACAO")     
            dadosEntradas(idxEntradas)("DataEmissao") = BEN_DADOSADV("DATA_EMISSAO")     

            idxEntradas = idxEntradas + 1
            BEN_DADOSADV.MoveNext()
        Wend
        
        If isObject(BEN_DADOSADV) Then
            BEN_DADOSADV.Close()
            SET BEN_DADOSADV = NOTHING
        End If
    End If
    On Error Goto 0
End Sub

Sub CarregarAbaSaidas()
    Dim SQL

    SQL = "SELECT NUMERO_NF, CODIGO_ERP, NOME_CLIENTE, CODIGO_CLIENTE, PRODUTO AS CODIGO_PRODUTO, " 
    SQL = SQL & "DESCRICAO_PRODUTO, QUANTIDADE, EMISSAO AS DATA_EMISSAO, PEDIDO, PESO "
    SQL = SQL & "FROM VW_SAIDA_APARAS WHERE 1=1 "
    SQL = SQL & "AND CODIGO_ERP = '" & filtroCodERP & "' "
    
    On Error Resume Next
    SET BEN_DADOSADV = Conecta("BEN-DADOSADV", SQL)
    
    If Err.Number = 0 AND IsObject(BEN_DADOSADV) Then
        While NOT BEN_DADOSADV.EOF
            ReDim Preserve dadosSaidas(idxSaidas)
            Set dadosSaidas(idxSaidas) = Server.CreateObject("Scripting.Dictionary")
            dadosSaidas(idxSaidas)("NomeCliente") = BEN_DADOSADV("NOME_CLIENTE")
            dadosSaidas(idxSaidas)("CodigoCliente") = BEN_DADOSADV("CODIGO_CLIENTE")
            dadosSaidas(idxSaidas)("NumeroNf") = BEN_DADOSADV("NUMERO_NF")
            dadosSaidas(idxSaidas)("CodigoProduto") = BEN_DADOSADV("CODIGO_PRODUTO")
            dadosSaidas(idxSaidas)("DescricaoProduto") = BEN_DADOSADV("DESCRICAO_PRODUTO")
            dadosSaidas(idxSaidas)("Pedido") = BEN_DADOSADV("PEDIDO")
            dadosSaidas(idxSaidas)("Quantidade") = BEN_DADOSADV("QUANTIDADE")
            dadosSaidas(idxSaidas)("Peso") = BEN_DADOSADV("PESO")
            dadosSaidas(idxSaidas)("DataEmissao") = BEN_DADOSADV("DATA_EMISSAO")  

            idxSaidas = idxSaidas + 1
            BEN_DADOSADV.MoveNext()
        Wend
        
        If isObject(BEN_DADOSADV) Then
            BEN_DADOSADV.Close()
            SET BEN_DADOSADV = NOTHING
        End If
    End If
    On Error Goto 0
End Sub

Sub CarregarAbaPedidosValidados()
    Dim SQL
    
    SQL = "SELECT PEDIDO AS CODIGO_PEDIDO, NOME_CLIENTE, PRODUTO AS CODIGO_PRODUTO, SALDO_PESO, " 
    SQL = SQL & "DESCRICAO_PRODUTO, QUANTIDADE_VENDIDA, QUANTIDADE_ENTREGUE, EMISSAO AS DATA_EMISSAO "
    SQL = SQL & "FROM VW_PEDIDOS_VALIDADOS_APARAS WHERE 1=1 "
    SQL = SQL & "AND CODIGO_ERP = '" & filtroCodERP & "' "

    On Error Resume Next
    SET BEN_DADOSADV = Conecta("BEN-DADOSADV", SQL)
    
    If Err.Number = 0 AND IsObject(BEN_DADOSADV) Then
        While NOT BEN_DADOSADV.EOF
            ReDim Preserve dadosPedidosValidados(idxPedidosValidados)
            Set dadosPedidosValidados(idxPedidosValidados) = Server.CreateObject("Scripting.Dictionary")
            dadosPedidosValidados(idxPedidosValidados)("NomeCliente") = BEN_DADOSADV("NOME_CLIENTE")
            dadosPedidosValidados(idxPedidosValidados)("CodigoPedido") = BEN_DADOSADV("CODIGO_PEDIDO")
            dadosPedidosValidados(idxPedidosValidados)("CodigoProduto") = BEN_DADOSADV("CODIGO_PRODUTO")
            dadosPedidosValidados(idxPedidosValidados)("DescricaoProduto") = BEN_DADOSADV("DESCRICAO_PRODUTO")
            dadosPedidosValidados(idxPedidosValidados)("QuantidadeVendida") = BEN_DADOSADV("QUANTIDADE_VENDIDA")
            dadosPedidosValidados(idxPedidosValidados)("QuantidadeEntregue") = BEN_DADOSADV("QUANTIDADE_ENTREGUE")
            dadosPedidosValidados(idxPedidosValidados)("SaldoPeso") = BEN_DADOSADV("SALDO_PESO")
            dadosPedidosValidados(idxPedidosValidados)("DataEmissao") = BEN_DADOSADV("DATA_EMISSAO")

            idxPedidosValidados = idxPedidosValidados + 1
            BEN_DADOSADV.MoveNext()
        Wend
        
        If isObject(BEN_DADOSADV) Then
            BEN_DADOSADV.Close()
            SET BEN_DADOSADV = NOTHING
        End If
    End If
    On Error Goto 0
End Sub

Sub CarregarAbaPedidosValidar()
    Dim SQL

    SQL = "SELECT PEDIDO AS CODIGO_PEDIDO, NOME_CLIENTE, PRODUTO AS CODIGO_PRODUTO, " 
    SQL = SQL & "DESCRICAO_PRODUTO, QUANTIDADE, PESO, EMISSAO AS DATA_EMISSAO "
    SQL = SQL & "FROM VW_PEDIDOS_VALIDAR_APARAS WHERE 1=1 "
    SQL = SQL & "AND CODIGO_ERP = '" & filtroCodERP & "' "

    On Error Resume Next
    SET BEN_DADOSADV = Conecta("BEN-DADOSADV", SQL)
    
    If Err.Number = 0 AND IsObject(BEN_DADOSADV) Then
        While NOT BEN_DADOSADV.EOF
            ReDim Preserve dadosPedidosValidar(idxPedidosValidar)
            Set dadosPedidosValidar(idxPedidosValidar) = Server.CreateObject("Scripting.Dictionary")
            dadosPedidosValidar(idxPedidosValidar)("NomeCliente") = BEN_DADOSADV("NOME_CLIENTE")
            dadosPedidosValidar(idxPedidosValidar)("CodigoPedido") = BEN_DADOSADV("CODIGO_PEDIDO")
            dadosPedidosValidar(idxPedidosValidar)("CodigoProduto") = BEN_DADOSADV("CODIGO_PRODUTO")
            dadosPedidosValidar(idxPedidosValidar)("DescricaoProduto") = BEN_DADOSADV("DESCRICAO_PRODUTO")
            dadosPedidosValidar(idxPedidosValidar)("Quantidade") = BEN_DADOSADV("QUANTIDADE")
            dadosPedidosValidar(idxPedidosValidar)("Peso") = BEN_DADOSADV("PESO")
            dadosPedidosValidar(idxPedidosValidar)("DataEmissao") = BEN_DADOSADV("DATA_EMISSAO")

            idxPedidosValidar = idxPedidosValidar + 1
            BEN_DADOSADV.MoveNext()
        Wend
        
        If isObject(BEN_DADOSADV) Then
            BEN_DADOSADV.Close()
            SET BEN_DADOSADV = NOTHING
        End If
    End If
    On Error Goto 0
End Sub

Function FormatarNumero(valor)
    If IsNull(valor) Or valor = "" Then
        FormatarNumero = "0,00"
        Exit Function
    End If
    
    Dim valorString
    valorString = CStr(valor)
    valorString = Replace(valorString, ".", ",")
    
    Dim valorNumerico
    valorNumerico = CDbl(valorString)
    
    Dim valorFormatado
    valorFormatado = FormatNumber(valorNumerico, 2, -1, 0, -1)
    
    FormatarNumero = valorFormatado
End Function

Function ObterNomeAba(aba)
    Select Case aba
        Case "entradas"
            ObterNomeAba = "Entradas"
        Case "saidas"
            ObterNomeAba = "Saídas"
        Case "pedidosvalidados"
            ObterNomeAba = "Pedidos Validados"
        Case "pedidosvalidar"
            ObterNomeAba = "Pedidos a Validar"
        Case Else
            ObterNomeAba = "Relatório"
    End Select
End Function
%>

<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <title>Relatório de Saldo de Aparas - Smurfit Kappa</title>
        <meta charset="utf-8"/>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
        <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />    
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js"></script>
        <script type="text/javascript">
            function PesquisaCliente() {
                urlBuscaCliente = `${location.protocol}//${location.host}/corporativo/buscaCLIENTEBRASIL.asp?id1=filtroCodERP&id2=codCliente&id3=codLoja&id6=05`;
                window.open(urlBuscaCliente, '', "width=650,height=400");
            }
            function alterarAba(aba) {
                const codCliente = $('#filtroCodERP').val();
                $('#abaHidden').val(aba);
                $('#formFiltro').submit();
            }
            function validarFiltros() {
                const codERP = $('#filtroCodERP').val()
                
                if(!codERP) {
                    alert("Código do cliente é obrigatório")
                    return false;
                }
                return true;
            }

            function exportarPDF() {
                const { jsPDF } = window.jspdf;
                const doc = new jsPDF('landscape');
                const dataAtual = new Date().toLocaleDateString('pt-BR');
                const aba = "<%=abaSelecionada%>";
                const nomeAba = "<%=ObterNomeAba(abaSelecionada)%>";
                const codCliente = "<%=filtroCodERP%>";

                let tabelaId = '';
                let colunas = [];
                let temDados = false;

                switch(aba) {
                    case 'entradas':
                        tabelaId = 'GradeEntradas';
                        colunas = ['Cliente', 'Cod. Produto', 'Desc. Produto', 'Serie', 'Documento', 'Quantidade', 'Data Emissão', 'Data Digitação'];
                        temDados = <%=IIf(idxEntradas > 0, "true", "false")%>;
                        break;
                    case 'saidas':
                        tabelaId = 'GradeSaidas';
                        colunas = ['Cliente', 'Número NF', 'Cod. Produto', 'Desc. Produto', 'Pedido', 'Quantidade', 'Peso', 'Data Emissão'];
                        temDados = <%=IIf(idxSaidas > 0, "true", "false")%>;
                        break;
                    case 'pedidosvalidados':
                        tabelaId = 'GradePedidosValidados';
                        colunas = ['Cliente', 'Cod. Pedido', 'Cod. Produto', 'Desc. Produto', 'Qtd. Vendida', 'Qtd. Entregue', 'Saldo Peso', 'Data Emissão'];
                        temDados = <%=IIf(idxPedidosValidados > 0, "true", "false")%>;
                        break;
                    case 'pedidosvalidar':
                        tabelaId = 'GradePedidosValidar';
                        colunas = ['Cliente', 'Cod. Pedido', 'Cod. Produto', 'Desc. Produto', 'Quantidade', 'Peso', 'Data Emissão'];
                        temDados = <%=IIf(idxPedidosValidar > 0, "true", "false")%>;
                        break;
                }

                const dados = [];
                
                if (temDados) {
                    const tabela = document.getElementById(tabelaId);
                    const linhas = tabela.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
                    
                    for (let i = 0; i < linhas.length; i++) {
                        const celulas = linhas[i].getElementsByTagName('td');
                        const linha = [];
                        for (let j = 0; j < celulas.length; j++) {
                            let texto = celulas[j].innerText.trim();
                            linha.push(texto);
                        }
                        dados.push(linha);
                    }
                }
                
                doc.setFontSize(16);
                doc.text(`Relatório de Saldo de Aparas - ${nomeAba}`, 14, 15);
                
                doc.setFontSize(10);
                const statusMsg = temDados ? `${dados.length} registro(s) encontrado(s)` : 'Nenhum registro encontrado';
                doc.text(`Data: ${dataAtual} | Cliente: ${codCliente} | ${statusMsg}`, 14, 22);

                doc.autoTable({
                    head: [colunas],
                    body: dados,
                    startY: 28,
                    theme: 'striped',
                    styles: { 
                        fontSize: 8,
                        cellPadding: 2
                    },
                    headStyles: { 
                        fillColor: [211, 211, 212],
                        textColor: [0, 0, 0],
                        fontStyle: 'bold'
                    },
                    margin: { top: 28 }
                });
                                
                window.open(doc.output('bloburl'), '_blank');
            }

            function exportarCSV() {
                const codERP = $('#filtroCodERP').val();
                const aba = "<%=abaSelecionada%>";
                
                if(!codERP) {
                    alert("Código do cliente é obrigatório para exportação");
                    return false;
                }

                let tabelaId = '';
                let nomeArquivo = '';
                let colunas = [];
                let temDados = false;

                switch(aba) {
                    case 'entradas':
                        tabelaId = 'GradeEntradas';
                        nomeArquivo = 'Entradas_Aparas.csv';
                        colunas = ['Cliente', 'Cod. Produto', 'Desc. Produto', 'Serie', 'Documento', 'Quantidade', 'Data de Emissao', 'Data de Digitacao'];
                        temDados = <%=IIf(idxEntradas > 0, "true", "false")%>;
                        break;
                    case 'saidas':
                        tabelaId = 'GradeSaidas';
                        nomeArquivo = 'Saidas_Aparas.csv';
                        colunas = ['Cliente', 'Numero NF', 'Cod. Produto', 'Desc. Produto', 'Pedido', 'Quantidade', 'Peso', 'Data de Emissao'];
                        temDados = <%=IIf(idxSaidas > 0, "true", "false")%>;
                        break;
                    case 'pedidosvalidados':
                        tabelaId = 'GradePedidosValidados';
                        nomeArquivo = 'Pedidos_Validados_Aparas.csv';
                        colunas = ['Cliente', 'Cod. Pedido', 'Cod. Produto', 'Desc. Produto', 'Quantidade Vendida', 'Quantidade Entregue', 'Saldo Peso', 'Data Emissao'];
                        temDados = <%=IIf(idxPedidosValidados > 0, "true", "false")%>;
                        break;
                    case 'pedidosvalidar':
                        tabelaId = 'GradePedidosValidar';
                        nomeArquivo = 'Pedidos_a_Validar_Aparas.csv';
                        colunas = ['Cliente', 'Cod. Pedido', 'Cod. Produto', 'Desc. Produto', 'Quantidade', 'Peso', 'Data de Emissao'];
                        temDados = <%=IIf(idxPedidosValidar > 0, "true", "false")%>;
                        break;
                }

                let csvContent = '\uFEFF'; 
                csvContent += colunas.join(';') + '\r\n';

                if (temDados) {
                    const tabela = document.getElementById(tabelaId);
                    const linhas = tabela.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
                    
                    for (let i = 0; i < linhas.length; i++) {
                        const celulas = linhas[i].getElementsByTagName('td');
                        const linhaDados = [];
                        
                        for (let j = 0; j < celulas.length; j++) {
                            let texto = celulas[j].innerText.trim();
                            texto = texto;
                            linhaDados.push(texto);
                        }
                        
                        csvContent += linhaDados.join(';') + '\r\n';
                    }
                }

                const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                const link = document.createElement('a');
                
                if (navigator.msSaveBlob) { 
                    navigator.msSaveBlob(blob, nomeArquivo);
                } else {
                    const url = URL.createObjectURL(blob);
                    link.href = url;
                    link.download = nomeArquivo;
                    link.style.display = 'none';
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    URL.revokeObjectURL(url);
                }
            }
        </script>
        <style type="text/css">
            body{
                background-color: #f5f5f5;
                font-size: .9rem !important;
                padding-bottom: 20px;
            }
            .form-label {
                font-weight: 600;
            }
            .card {
                background: white;
                border-radius: 8px;
                padding: 20px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                border: none;
            }
            .card-title {
                font-size: 0.875rem;
                color: #6c757d;
                font-weight: 500;
                margin-bottom: 8px;
            }
            .card-value {
                font-size: 1.75rem;
                font-weight: 600;
                color: #212529;
            }
            .card-icon {
                font-size: 1.2rem;
                opacity: .7;
            }
            .table-wrapper {
                overflow: auto;
                max-height: 60vh;
            }
            thead {
                vertical-align: middle !important;
                position: sticky; 
                inset-block-start: 0; 
                background: #f5f5f5;
            }
            td, th {
                padding: 8px 12px !important;
            }
            th {
                font-weight: 600;
                color: #495057;
            }
            tr:hover {
                background: #f5f5f5;
            }
            .no-records {
                padding: 10px;
                text-align: center;
                background: white;
            }
            .tabs-container {
                display: flex;
                padding: 0;
                gap: 0;
                min-width: max-content;
                background-color: #f8f9fa;
                border-radius: 8px 8px 0 0;
                margin-top: 20px;
            }
            .tab-button {
                flex: 0 0 auto; 
                padding: 15px 20px;
                background: transparent;
                border: none;
                border-bottom: 3px solid transparent;
                cursor: pointer;
                display: flex;
                align-items: center;
                justify-content: baseline;
                gap: 8px;
                color: #6c757d;
                font-weight: 500;
            }
            .tab-button:hover {
                color: #212529;
            }
            .tab-button.active {
                color: #0d6efd;
                border-bottom-color: #0d6efd;
                background-color: #fff;
            }
            .filtros-card {
                background-color: white;
                border-radius: 8px;
                padding: 15px;
                margin-bottom: 20px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            }
            .truncate-text {
                max-width: 200px;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
                cursor: help;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h3 class="mt-4 mb-4">Relatório de Saldo de Aparas</h3>

            <div class="filtros-card">
                <form method="post" action="" onsubmit="return validarFiltros();" id="formFiltro">
                    <input type="hidden" id="abaHidden" name="aba" value="">
                    <input type="hidden" id="codCliente" name="codCliente" value="">
                    <input type="hidden" id="codLoja" name="codLoja" value="">

                    <div class="row">
                        <div class="col-12 col-md-4 mb-3 ">
                            <label for="filtroCodERP" class="form-label">Cliente</label>
                            <div class="input-group">
                                <span class="btn btn-sm btn-secondary" onclick="PesquisaCliente()" > 
                                    Buscar
                                    <i class="fa fa-external-link-alt"></i>
                                </span>
                                <input type="text" class="form-control form-control-sm" id="filtroCodERP" name="filtroCodERP" value="<%=filtroCodERP%>">
                            </div>
                            <small class="text-muted ">Clique no botão para buscar ou digite um cód.</small>
                        </div>
                        <div class="col-md-4"> 
                            <label for="filtroCodERP" class="form-label">
                                Pesquisar
                            </label>
                            <button type="submit" class="btn btn-sm btn-primary form-control">
                                <i class="fas fa-search me-2"></i> Pesquisar
                            </button>
                        </div>
                    </div>
                </form>
            </div>
            
            <div class="row g-3 mb-4">
                <div class="col-12 col-md-4 col-lg-2">
                    <div class="card">
                        <div class="d-flex justify-content-between align-items-start">
                            <div>
                                <div class="card-title">Saldo</div>
                                <div class="card-value"><%=FormatarNumero(totalSaldo)%></div>
                            </div>
                            <i class="fas fa-balance-scale card-icon"></i>
                        </div>
                    </div>
                </div>
                
                <div class="col-12 col-md-4 col-lg-2">
                    <div class="card">
                        <div class="d-flex justify-content-between align-items-start">
                            <div>
                                <div class="card-title">Total de Entradas</div>
                                <div class="card-value"><%=FormatarNumero(totalEntradas)%></div>
                            </div>
                            <i class="fas fa-arrow-down card-icon text-success"></i>
                        </div>
                    </div>
                </div>
                
                <div class="col-12 col-md-4 col-lg-2">
                    <div class="card">
                        <div class="d-flex justify-content-between align-items-start">
                            <div>
                                <div class="card-title">Total de Saídas</div>
                                <div class="card-value"><%=FormatarNumero(totalSaidas)%></div>
                            </div>
                            <i class="fas fa-arrow-up card-icon text-danger"></i>
                        </div>
                    </div>
                </div>
                
                <div class="col-12 col-md-6 col-lg-3">
                    <div class="card">
                        <div class="d-flex justify-content-between align-items-start">
                            <div>
                                <div class="card-title">Total de Pedidos Validados</div>
                                <div class="card-value"><%=FormatarNumero(totalPedidosValidados)%></div>
                            </div>
                           <i class="fas fa-check-circle card-icon text-success"></i>
                        </div>
                    </div>
                </div>
                
                <div class="col-12 col-md-6 col-lg-3">
                    <div class="card">
                        <div class="d-flex justify-content-between align-items-start">
                            <div>
                                <div class="card-title">Total de Pedidos a Validar</div>
                                <div class="card-value"><%=FormatarNumero(totalPedidosValidar)%></div>
                            </div>
                            <i class="fa fa-clock card-icon text-warning"></i>
                        </div>
                    </div>
                </div>
            </div>
            
            <div class="tabs-container">
                <button type="button" class="tab-button <%= IIf(abaSelecionada = "entradas", "active", "") %>" onclick="alterarAba('entradas')">
                    <i class="fas fa-arrow-down"></i>
                    <span class="tab-title">Entradas</span>
                </button>
                <button type="button" class="tab-button <%= IIf(abaSelecionada = "saidas", "active", "") %>" onclick="alterarAba('saidas')">
                    <i class="fas fa-arrow-up"></i>
                    <span class="tab-title">Saídas</span>
                </button>
                <button type="button" class="tab-button <%= IIf(abaSelecionada = "pedidosvalidados", "active", "") %>" onclick="alterarAba('pedidosvalidados')">
                    <i class="fas fa-check-circle"></i>
                    <span class="tab-title">Pedidos Validados</span>
                </button>
                <button type="button" class="tab-button <%= IIf(abaSelecionada = "pedidosvalidar", "active", "") %>" onclick="alterarAba('pedidosvalidar')">
                    <i class="fa fa-clock"></i>
                    <span class="tab-title">Pedidos a Validar</span>
                </button>
            </div>
            
            <div class="card p-0">
                <% If filtroCodERP <> "" Then %>
                <div class="d-flex justify-content-end gap-2 p-3 border-bottom">
                    <button type="button" class="btn btn-success btn-sm" onclick="exportarCSV()">
                        <i class="fas fa-file-excel me-2"></i>Exportar Excel
                    </button>
                    <button type="button" class="btn btn-danger btn-sm" onclick="exportarPDF()">
                        <i class="fas fa-file-pdf me-2"></i>Exportar PDF
                    </button>
                </div>
                <% End If %>
                
                <div class="table-wrapper">
                    <% If abaSelecionada = "entradas" Then %>
                        <table class="table" name="GradeEntradas" id="GradeEntradas">
                            <thead>
                                <tr>
                                    <th>Cliente</th>
                                    <th>Cod. Produto</th>
                                    <th>Desc. Produto</th>
                                    <th>Serie</th>
                                    <th>Documento</th>
                                    <th>Quantidade</th>
                                    <th>Data de emissão</th>
                                    <th>Data de digitação</th>
                                </tr>              
                            </thead>
                            <% IF idxEntradas > 0 THEN %>
                            <tbody>
                                <% 
                                Dim i
                                For i = 0 To UBound(dadosEntradas)
                                %>
                                    <tr>
                                        <td class="text-left">
                                            <div class="truncate-text" title="<%=dadosEntradas(i)("NomeCliente") %>">
                                                <%=dadosEntradas(i)("NomeCliente") %>
                                            </div>
                                        </td>
                                        <td class="text-left"><%=dadosEntradas(i)("CodigoProduto") %></td>
                                        <td class="text-left"><%=dadosEntradas(i)("DescricaoProduto") %></td>
                                        <td class="text-left"><%=dadosEntradas(i)("Serie") %></td>
                                        <td class="text-left"><%=dadosEntradas(i)("Documento") %></td>
                                        <td class="text-left"><%=FormatarNumero(dadosEntradas(i)("Quantidade")) %></td>
                                        <td class="text-left"><%=acertadata(dadosEntradas(i)("DataEmissao")) %></td>
                                        <td class="text-left"><%=acertadata(dadosEntradas(i)("DataDigitacao")) %></td>
                                    </tr>
                                <% Next %>
                            </tbody>
                            <% END IF %>
                        </table>
                        <% IF idxEntradas = 0 THEN %>
                            <div class="no-records">
                                Nenhum registro encontrado
                            </div>
                        <% END IF %>
                    
                    <% ElseIf abaSelecionada = "saidas" Then %>
                        <table class="table" name="GradeSaidas" id="GradeSaidas">
                            <thead>
                                <tr>
                                    <th>Cliente</th>
                                    <th>Número NF</th>
                                    <th>Cod. Produto</th>
                                    <th>Desc. Produto</th>
                                    <th>Pedido</th>
                                    <th>Quantidade</th>
                                    <th>Peso</th>
                                    <th>Data de emissão</th>
                                </tr>              
                            </thead>
                            <% IF idxSaidas > 0 THEN %>
                            <tbody>
                                <% 
                                For i = 0 To UBound(dadosSaidas)
                                %>
                                    <tr>
                                        <td class="text-left">
                                            <div class="truncate-text" title="<%=dadosSaidas(i)("NomeCliente") %>">
                                                <%=dadosSaidas(i)("NomeCliente") %>
                                            </div>
                                        </td>
                                        <td class="text-left"><%=dadosSaidas(i)("NumeroNf") %></td>
                                        <td class="text-left"><%=dadosSaidas(i)("CodigoProduto") %></td>
                                        <td class="text-left"><%=dadosSaidas(i)("DescricaoProduto") %></td>
                                        <td class="text-left"><%=dadosSaidas(i)("Pedido") %></td>
                                        <td class="text-left"><%=FormatarNumero(dadosSaidas(i)("Quantidade")) %></td>
                                        <td class="text-left"><%=FormatarNumero(dadosSaidas(i)("Peso")) %></td>
                                        <td class="text-left"><%=acertadata(dadosSaidas(i)("DataEmissao")) %></td>
                                    </tr>
                                <% Next %>
                            </tbody>
                            <% END IF %>
                        </table>
                        <% IF idxSaidas = 0 THEN %>
                            <div class="no-records">
                                Nenhum registro encontrado
                            </div>
                        <% END IF %>
                    
                    <% ElseIf abaSelecionada = "pedidosvalidados" Then %>
                        <table class="table" name="GradePedidosValidados" id="GradePedidosValidados">
                            <thead>
                                <tr>
                                    <th>Cliente</th>
                                    <th>Cod. Pedido</th>
                                    <th>Cod. Produto</th>
                                    <th>Desc. Produto</th>
                                    <th>Quantidade Vendida</th>
                                    <th>Quantidade Entregue</th>
                                    <th>Saldo Peso</th>
                                    <th>Data Emissão</th>
                                </tr>              
                            </thead>
                            <% IF idxPedidosValidados > 0 THEN %>
                            <tbody>
                                <% For i = 0 To UBound(dadosPedidosValidados) %>
                                    <tr>
                                        <td class="text-left">
                                            <div class="truncate-text" title="<%=dadosPedidosValidados(i)("NomeCliente") %>">
                                                <%=dadosPedidosValidados(i)("NomeCliente") %>
                                            </div>
                                        </td>
                                        <td class="text-left"><%=dadosPedidosValidados(i)("CodigoPedido") %></td>
                                        <td class="text-left"><%=dadosPedidosValidados(i)("CodigoProduto") %></td>
                                        <td class="text-left"><%=dadosPedidosValidados(i)("DescricaoProduto") %></td>
                                        <td class="text-left"><%=FormatarNumero(dadosPedidosValidados(i)("QuantidadeVendida")) %></td>
                                        <td class="text-left"><%=FormatarNumero(dadosPedidosValidados(i)("QuantidadeEntregue")) %></td>
                                        <td class="text-left"><%=FormatarNumero(dadosPedidosValidados(i)("SaldoPeso")) %></td>
                                        <td class="text-left"><%=acertadata(dadosPedidosValidados(i)("DataEmissao")) %></td>
                                    </tr>
                                <% Next %>
                            </tbody>
                            <% END IF %>
                        </table>
                        <% IF idxPedidosValidados = 0 THEN %>
                            <div class="no-records">
                                Nenhum registro encontrado
                            </div>
                        <% END IF %>
                    
                    <% ElseIf abaSelecionada = "pedidosvalidar" Then %>
                        <table class="table" name="GradePedidosValidar" id="GradePedidosValidar">
                            <thead>
                                <tr>
                                    <th>Cliente</th>
                                    <th>Cod. Pedido</th>
                                    <th>Cod. Produto</th>
                                    <th>Desc. Produto</th>
                                    <th>Quantidade</th>
                                    <th>Peso</th>
                                    <th>Data de Emissão</th>
                                </tr>              
                            </thead>
                            <% IF idxPedidosValidar > 0 THEN %>
                            <tbody>
                                <% For i = 0 To UBound(dadosPedidosValidar) %>
                                    <tr>
                                        <td class="text-left">
                                            <div class="truncate-text" title="<%=dadosPedidosValidar(i)("NomeCliente") %>">
                                                <%=dadosPedidosValidar(i)("NomeCliente") %>
                                            </div>
                                        </td>
                                        <td class="text-left"><%=dadosPedidosValidar(i)("CodigoPedido") %></td>
                                        <td class="text-left"><%=dadosPedidosValidar(i)("CodigoProduto") %></td>
                                        <td class="text-left"><%=dadosPedidosValidar(i)("DescricaoProduto") %></td>
                                        <td class="text-left"><%=FormatarNumero(dadosPedidosValidar(i)("Quantidade")) %></td>
                                        <td class="text-left"><%=FormatarNumero(dadosPedidosValidar(i)("Peso")) %></td>
                                        <td class="text-left"><%=acertadata(dadosPedidosValidar(i)("DataEmissao")) %></td>
                                    </tr>
                                <% Next %>
                            </tbody>
                            <% END IF %>
                        </table>
                        <% IF idxPedidosValidar = 0 THEN %>
                            <div class="no-records">
                                Nenhum registro encontrado
                            </div>
                        <% END IF %>
                    <% End If %>
                </div>
            </div>
        </div>
    </body>
</html>