<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../../cnn2/cnn2.asp" -->
<!--#include File = "../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046

Dim idFNC
idFNC = Request.QueryString("idFNC")

If idFNC = "" Or Not IsNumeric(idFNC) Then
    Response.Write("<div class='alert alert-danger'>ID de FNC inválido.</div>")
    Response.End
End If

Dim SQL, rsRespostas

Dim respostas()
ReDim respostas(0)
Dim idxResposta
idxResposta = 0

SQL = "SELECT r.IdFNC, f.Planta as Filial, f.Cliente as CodCliente, f.Loja as CodLoja, f.Situacao, "
SQL = SQL & " REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), f.Descricao), CHAR(7), ''), CHAR(34), ''), CHAR(0), '') AS Descricao, "
SQL = SQL & " REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(MAX), r.Justificativa), CHAR(7), ''), CHAR(34), ''), CHAR(0), '') AS Justificativa,"
SQL = SQL & " r.QuemIncluiu as IdUsuarioInclusao, c1.CRM_NOME as NomeUsuarioInclusao, cli.NOME_CLIENTE AS NomeCliente,"
SQL = SQL & " r.QuemFinalizou as IdUsuarioFinalizacao, c2.CRM_NOME as NomeUsuarioFinalizacao, r.DataInclusao, r.DataFinalizacao"
SQL = SQL & " FROM Resposta r "
SQL = SQL & " LEFT JOIN FNC f on f.Id = r.IdFNC "
SQL = SQL & " LEFT JOIN CRM_DTI c1 on c1.CRM_R_E_C_N_O = r.QuemIncluiu "
SQL = SQL & " LEFT JOIN CRM_DTI c2 on c2.CRM_R_E_C_N_O = r.QuemFinalizou "

SQL = SQL & "OUTER APPLY (  SELECT " &_
                                "TOP 1 NOME_CLIENTE " &_
                            "FROM BI_CLIENTES " &_
                            "WHERE CIDADE = " &_
                            "CASE f.Planta " &_
                                            "WHEN '00' THEN 'PIRAPETINGA' " &_
                                            "WHEN '04' THEN 'UBERABA' " &_
                                            "WHEN '05' THEN 'BENTO' " &_
                                            "WHEN '06' THEN 'MARANGUAPE' " &_
                                            "WHEN '07' THEN 'SAQUAREMA' " &_
                                            "WHEN '08' THEN 'ARACATUBA' " &_
                                            "WHEN '09' THEN 'BLUMENAU' " &_
                                            "WHEN '10' THEN 'PACAJUS' " &_
                                            "WHEN '11' THEN 'PORTO FELIZ' " &_
                                            "WHEN '12' THEN 'TRES BARRAS' " &_
                                            "WHEN '13' THEN 'FLORESTAL' " &_
                            "END AND " &_
                            "CASE WHEN f.Planta IN ('06','05') THEN CODIGO_ERP ELSE COD_CLIENTE END = f.Cliente COLLATE SQL_Latin1_General_CP1_CI_AS AND " &_
                            "CASE WHEN f.Planta IN ('09') THEN '01' ELSE LOJA_CLIENTE END = f.Loja COLLATE SQL_Latin1_General_CP1_CI_AS) AS cli "
SQL = SQL & "WHERE r.IdFNC = " & idFNC

On Error Resume Next
SET rsRespostas = Conecta("BEN-CORPORATIVO", SQL)

If Err.Number <> 0 OR NOT IsObject(rsRespostas) Then
    Response.Write("<div class='alert alert-danger'>Erro ao obter respostas.</div>")
    Response.End
End If

On Error Goto 0

While NOT rsRespostas.EOF
    ReDim Preserve respostas(idxResposta)
    Set respostas(idxResposta) = Server.CreateObject("Scripting.Dictionary")
    respostas(idxResposta)("IdFNC") = EscapeJSON(rsRespostas("IdFNC"))
    respostas(idxResposta)("Filial") = EscapeJSON(rsRespostas("Filial"))
    respostas(idxResposta)("CodCliente") = EscapeJSON(rsRespostas("CodCliente"))
    respostas(idxResposta)("CodLoja") = EscapeJSON(rsRespostas("CodLoja"))
    respostas(idxResposta)("Situacao") = EscapeJSON(rsRespostas("Situacao"))
    respostas(idxResposta)("IdUsuarioFinalizacao") = EscapeJSON(rsRespostas("IdUsuarioFinalizacao"))
    respostas(idxResposta)("IdUsuarioInclusao") = EscapeJSON(rsRespostas("IdUsuarioInclusao"))
    respostas(idxResposta)("Descricao") = Iif(rsRespostas("Descricao") <> "", EscapeJSON(rsRespostas("Descricao")), "-") 
    respostas(idxResposta)("Justificativa") = Iif(rsRespostas("Justificativa") <> "", EscapeJSON(rsRespostas("Justificativa")), "-") 
    respostas(idxResposta)("NomeUsuarioInclusao") = Iif(rsRespostas("NomeUsuarioInclusao") <> "", EscapeJSON(rsRespostas("NomeUsuarioInclusao")), "-") 
    respostas(idxResposta)("NomeCliente") = Iif(rsRespostas("NomeCliente") <> "", EscapeJSON(rsRespostas("NomeCliente")), "-") 
    respostas(idxResposta)("NomeUsuarioFinalizacao") = Iif(rsRespostas("NomeUsuarioFinalizacao") <> "", EscapeJSON(rsRespostas("NomeUsuarioFinalizacao")), "-") 
    respostas(idxResposta)("DataInclusao") = Iif(rsRespostas("DataInclusao") <> "", EscapeJSON(rsRespostas("DataInclusao")), "-") 
    respostas(idxResposta)("DataFinalizacao") = Iif(rsRespostas("DataFinalizacao") <> "", EscapeJSON(rsRespostas("DataFinalizacao")), "-") 

    idxResposta = idxResposta + 1
    rsRespostas.MoveNext()
Wend

If isObject(rsRespostas) Then
    rsRespostas.Close()
    SET rsRespostas = NOTHING
End If

Dim temRegistros
temRegistros = (idxResposta > 0)
%>

<% If temRegistros Then %>
    <div class="mb-3">
        <button type="button" class="btn btn-success btn-sm" onclick="exportarCSV()">
            <i class="bi bi-file-earmark-excel me-1"></i> Exportar Excel
        </button>
        <button type="button" class="btn btn-danger btn-sm" onclick="exportarPDF()">
            <i class="bi bi-file-earmark-pdf me-1"></i> Exportar PDF
        </button>
    </div>
    
    <table class="table table-sm table-hover history-table">
        <thead>
            <tr>
                <th>FNC</th>
                <th>Cliente</th>
                <th>Descrição</th>
                <th>Justificativa</th>
                <th>Incluida Por</th>
                <th>Finalizada Por</th>
                <th>Data Inclusão</th>
                <th>Data Finalização</th>
            </tr>
        </thead>
        <tbody>
        <% 
            Dim j
            For j = 0 To UBound(respostas)
            %>
                <tr>
                    <td><%=respostas(j)("IdFNC")%></td>
                    <td><%=respostas(j)("NomeCliente")%></td>
                    <td>
                        <div class="truncate-text" title="<%=respostas(j)("Descricao")%>">
                            <%=respostas(j)("Descricao")%>
                        </div>
                    </td>
                    <td>
                        <div class="truncate-text" title="<%=respostas(j)("Justificativa")%>">
                            <%=respostas(j)("Justificativa")%>
                        </div>
                    </td>
                    <td><%=respostas(j)("NomeUsuarioInclusao")%></td>
                    <td><%=respostas(j)("NomeUsuarioFinalizacao")%></td>
                    <td><%=FormatarData(respostas(j)("DataInclusao"))%></td>
                    <td><%=FormatarData(respostas(j)("DataFinalizacao"))%></td>
                </tr>
        <% Next %>
        </tbody>
    </table>
<% Else %>
    <div class="alert alert-info mb-0">
        <i class="fas fa-info-circle me-2"></i>
        Nenhuma resposta encontrada.
    </div>
<% End If %>


<script>    
    var dadosExportacao = [
        <% 
        If temRegistros Then
            For j = 0 To UBound(respostas)
                Dim descCSV, justCSV
                descCSV = EscapeJSON(respostas(j)("Descricao"))
                justCSV = EscapeJSON(respostas(j)("Justificativa"))
        %>
        {
            idFNC: "<%=respostas(j)("IdFNC")%>",
            nomeCliente: "<%=respostas(j)("NomeCliente")%>",
            descricao: "<%=descCSV%>",
            justificativa: "<%=justCSV%>",
            situacao: "<%=respostas(j)("Situacao")%>",
            nomeUsuarioInclusao: "<%=respostas(j)("NomeUsuarioInclusao")%>",
            nomeUsuarioFinalizacao: "<%=respostas(j)("NomeUsuarioFinalizacao")%>",
            dataInclusao: "<%=FormatarData(respostas(j)("DataInclusao"))%>",
            dataFinalizacao: "<%=FormatarData(respostas(j)("DataFinalizacao"))%>",
        }
        <%= IIf(j < UBound(respostas), ",", "") %>
        <% 
            Next
        End If
        %>
    ];

    function escapeCSV(valor) {
        if (!valor) return '""';
        valor = valor.replace(/[\r\n]+/g, ' ').trim();
        valor = valor.replace(/"/g, '""');
        return '"' + valor + '"';
    }
    
    function exportarPDF() {
        if (dadosExportacao.length === 0) {
            alert('Não há dados para exportar');
            return;
        }

        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('landscape');
        const dataAtual = new Date().toLocaleDateString('pt-BR');

        const colunas = ['IdFNC', 'Cliente', 'Descrição', 'Justificativa', 'Situação', 'Incluída Por', 'Finalizada Por', 'Data Inclusão', 'Data Finalização'];
        const dados = dadosExportacao.map(item => [
            item.idFNC || '-',
            item.nomeCliente || '-',
            item.descricao || '-',
            item.justificativa || '-',
            obterDescricaoSituacao(item.situacao || '-'),
            item.nomeUsuarioInclusao || '-',
            item.nomeUsuarioFinalizacao || '-',
            item.dataInclusao || '-',
            item.dataFinalizacao || '-'
        ]);

        doc.setFontSize(16);
        doc.text('Relatório de Respostas de FNC', 14, 15);
        
        doc.setFontSize(10);
        doc.text(`Data: ${dataAtual} | ${dados.length} registro(s) encontrado(s)`, 14, 22);

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
            columnStyles: {
                0: { cellWidth: 25 },
                1: { cellWidth: 'auto' },
                2: { cellWidth: 50 },
                3: { cellWidth: 50 }
            },
            margin: { top: 28 }
        });
                        
        window.open(doc.output('bloburl'), '_blank');
    }

    function exportarCSV() {
        if (dadosExportacao.length === 0) {
            alert('Não há dados para exportar');
            return;
        }

        const nomeArquivo = 'Relatorio_Respostas_FNC.csv';
        const colunas = ['IdFNC', 'Cliente', 'Descricao', 'Justificativa', 'Situação', 'Incluida Por', 'Finalizada Por', 'Data Inclusão', 'Data Finalizacao'];

        let csvContent = '\uFEFF'; 
        csvContent += colunas.join(';') + '\r\n';

        dadosExportacao.forEach(item => {
            const linha = [
                item.idFNC || '-',
                item.nomeCliente || '-',
                escapeCSV(item.descricao || '-'),
                escapeCSV(item.justificativa || '-'),
                obterDescricaoSituacao(item.situacao || '-'),
                item.nomeUsuarioInclusao || '-',
                item.nomeUsuarioFinalizacao || '-',
                item.dataInclusao || '-',
                item.dataFinalizacao || '-'
            ];
            csvContent += linha.join(';') + '\r\n';
        });

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

    function obterDescricaoSituacao(situacao) {
        switch(situacao) {
            case '1':
                return 'ABERTA';
            case '2':
                return 'RESPONDIDA';
            case '3':
                return 'ENCERRADA';
            default:
                return '';
        }
    }
</script>