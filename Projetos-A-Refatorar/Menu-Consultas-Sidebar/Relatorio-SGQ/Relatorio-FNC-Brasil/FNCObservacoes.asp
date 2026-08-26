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

Dim observacoes()
ReDim observacoes(0)
Dim idxObservacao
idxObservacao = 0

SQL = "SELECT o.Id, o.IdFnc, o.Descricao, o.IdUsuarioInclusao, o.DataCriacao, u.CRM_LOGIN as LoginUsuarioInclusao, u.CRM_NOME AS NomeUsuarioInclusao "
SQL = SQL & "FROM ObservacaoFnc o "
SQL = SQL & "LEFT JOIN CRM_DTI u ON o.IdUsuarioInclusao = u.CRM_R_E_C_N_O "
SQL = SQL & "WHERE o.IdFnc = " & idFNC & " "
SQL = SQL & "AND o.FlagInativo = 0"

On Error Resume Next
SET rsObservacoes = Conecta("BEN-CORPORATIVO", SQL)

If Err.Number <> 0 OR NOT IsObject(rsObservacoes) Then
    Response.Write("<div class='alert alert-danger'>Erro ao obter observacoes.</div>")
    Response.End
End If

On Error Goto 0

While NOT rsObservacoes.EOF
    ReDim Preserve observacoes(idxObservacao)
    Set observacoes(idxObservacao) = Server.CreateObject("Scripting.Dictionary")
    observacoes(idxObservacao)("Id") = EscapeJSON(rsObservacoes("Id"))
    observacoes(idxObservacao)("IdFNC") = EscapeJSON(rsObservacoes("IdFNC"))
    observacoes(idxObservacao)("Descricao") = EscapeJSON(rsObservacoes("Descricao"))
    observacoes(idxObservacao)("IdUsuarioInclusao") = EscapeJSON(rsObservacoes("IdUsuarioInclusao"))
    observacoes(idxObservacao)("DataCriacao") = EscapeJSON(rsObservacoes("DataCriacao"))
    observacoes(idxObservacao)("LoginUsuarioInclusao") = EscapeJSON(rsObservacoes("LoginUsuarioInclusao"))
    observacoes(idxObservacao)("NomeUsuarioInclusao") = EscapeJSON(rsObservacoes("NomeUsuarioInclusao"))

    idxObservacao = idxObservacao + 1
    rsObservacoes.MoveNext()
Wend

If isObject(rsObservacoes) Then
    rsObservacoes.Close()
    SET rsObservacoes = NOTHING
End If

Dim temRegistros
temRegistros = (idxObservacao > 0)
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
                <th>#</th>
                <th>FNC</th>
                <th>Descrição</th>
                <th>Incluida Por</th>
                <th>Data Inclusão</th>
            </tr>
        </thead>
        <tbody>
        <% 
            Dim j
            For j = 0 To UBound(observacoes)
            %>
                <tr>
                    <td><%=observacoes(j)("Id")%></td>
                    <td><%=observacoes(j)("IdFNC")%></td>
                    <td>
                        <div class="truncate-text" title="<%=observacoes(j)("Descricao")%>">
                            <%=observacoes(j)("Descricao")%>
                        </div>
                    </td>
                    <td><%=observacoes(j)("NomeUsuarioInclusao")%></td>
                    <td><%=FormatarData(observacoes(j)("DataCriacao"))%></td>
                </tr>
        <% Next %>
        </tbody>
    </table>
<% Else %>
    <div class="alert alert-info mb-0">
        <i class="fas fa-info-circle me-2"></i>
        Nenhuma observação encontrada.
    </div>
<% End If %>


<script>    
    var dadosExportacao = [
        <% 
        If temRegistros Then
            For j = 0 To UBound(observacoes)
                Dim descCSV, justCSV
                descCSV = EscapeJSON(observacoes(j)("Descricao"))
        %>
        {
            id: "<%=observacoes(j)("Id")%>",
            idFNC: "<%=observacoes(j)("IdFNC")%>",
            descricao: "<%=descCSV%>",
            nomeUsuarioInclusao: "<%=observacoes(j)("NomeUsuarioInclusao")%>",
            dataCriacao: "<%=FormatarData(observacoes(j)("DataCriacao"))%>",
        }
        <%= IIf(j < UBound(observacoes), ",", "") %>
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

        const colunas = ['Id', 'IdFNC', 'Descrição', 'Incluída Por', 'Data Inclusão'];
        const dados = dadosExportacao.map(item => [
            item.id || '-',
            item.idFNC || '-',
            item.descricao || '-',
            item.nomeUsuarioInclusao || '-',
            item.dataCriacao || '-',
        ]);

        doc.setFontSize(16);
        doc.text('Relatório de Observações de FNC', 14, 15);
        
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

        const nomeArquivo = 'Relatorio_Observacoes_FNC.csv';
        const colunas = ['Id', 'IdFNC', 'Descricao', 'Incluida Por', 'Data Inclusão'];

        let csvContent = '\uFEFF'; 
        csvContent += colunas.join(';') + '\r\n';

        dadosExportacao.forEach(item => {
            const linha = [
                item.id || '-',
                item.idFNC || '-',
                escapeCSV(item.descricao || '-'),
                item.nomeUsuarioInclusao || '-',
                item.dataCriacao || '-',
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
</script>