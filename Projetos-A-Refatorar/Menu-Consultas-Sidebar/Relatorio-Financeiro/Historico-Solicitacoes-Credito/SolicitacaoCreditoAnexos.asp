<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="./util/Util.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0

Dim idSolicitacao
Dim enderecoLegado

idSolicitacao = Request("idSolicitacao")
enderecoLegado = Request("enderecoLegado")

If idSolicitacao = "" Then
    Response.Write "<div class='alert alert-warning'>ID da solicitação não informado.</div>"
    Response.End
End If

' Busca os anexos da solicitação
Dim anexos()
ReDim anexos(0)
Dim idxAnexo
idxAnexo = 0

SQL = "SELECT "
SQL = SQL & "scd.Id, "
SQL = SQL & "scd.Nome, "
SQL = SQL & "scd.Caminho, "
SQL = SQL & "CONVERT(VARCHAR(19), scd.DataCriacao, 120) as DataCriacao "
SQL = SQL & "FROM SolicitacaoCreditoDocumento scd "
SQL = SQL & "WHERE scd.IdSolicitacaoCredito = " & idSolicitacao & " "
SQL = SQL & "ORDER BY scd.DataCriacao DESC"

SET RS = Conecta("PIR-CORPORATIVO", SQL)

While NOT RS.EOF
    ReDim Preserve anexos(idxAnexo)
    Set anexos(idxAnexo) = Server.CreateObject("Scripting.Dictionary")
    anexos(idxAnexo)("Id") = RS("Id")
    anexos(idxAnexo)("Nome") = RS("Nome")
    anexos(idxAnexo)("Caminho") = RS("Caminho")
    anexos(idxAnexo)("DataCriacao") = RS("DataCriacao")
    idxAnexo = idxAnexo + 1
    RS.MoveNext()
Wend

If isObject(RS) Then
    RS.Close()
    SET RS = NOTHING
End If

Dim temAnexosTabela, temEnderecoLegado
temAnexosTabela = (idxAnexo > 0)
temEnderecoLegado = (enderecoLegado <> "" AND enderecoLegado <> "undefined" AND enderecoLegado <> "null")
%>

<style>
    .anexo-item {
        padding: 12px;
        border: 1px solid #e5e7eb;
        border-radius: 6px;
        margin-bottom: 10px;
        transition: background-color 0.2s;
    }
    .anexo-item:hover {
        background-color: #f9fafb;
    }
    .anexo-icon {
        font-size: 1.5rem;
        color: #3b82f6;
        margin-right: 12px;
    }
    .anexo-info {
        flex: 1;
    }
    .anexo-nome {
        font-weight: 600;
        color: #1f2937;
        margin-bottom: 4px;
    }
    .anexo-data {
        font-size: 0.85rem;
        color: #6b7280;
    }
</style>

<% If temAnexosTabela Then %>
    <div class="list-group">
        <% 
        Dim j
        For j = 0 To UBound(anexos)
        %>
            <div class="anexo-item d-flex align-items-center">
                <i class="fas fa-file-alt anexo-icon"></i>
                <div class="anexo-info">
                    <div class="anexo-nome"><%=anexos(j)("Nome")%></div>
                    <div class="anexo-data">Adicionado em: <%=FormatarData(anexos(j)("DataCriacao"))%></div>
                </div>
                <a href="<%="/Clientes_Docs/" & anexos(j)("Caminho")%>" target="_blank" class="btn btn-sm btn-primary">
                    <i class="fas fa-external-link-alt me-1"></i> Abrir
                </a>
            </div>
        <% Next %>
    </div>
    
    <% If temEnderecoLegado Then %>
        <hr>
        <h6 class="text-muted mb-3">Documentação Legada</h6>
        <div class="anexo-item d-flex align-items-center">
            <i class="fas fa-folder anexo-icon" style="color: #f59e0b;"></i>
            <div class="anexo-info">
                <div class="anexo-nome">Documentação</div>
                <div class="anexo-data">Pasta de documentos</div>
            </div>
            <a href="<%=enderecoLegado%>" target="_blank" class="btn btn-sm btn-warning">
                <i class="fas fa-external-link-alt me-1"></i> Abrir Pasta
            </a>
        </div>
    <% End If %>
    
<% ElseIf temEnderecoLegado Then %>
    <div class="anexo-item d-flex align-items-center">
        <i class="fas fa-folder anexo-icon" style="color: #f59e0b;"></i>
        <div class="anexo-info">
            <div class="anexo-nome">Documentação</div>
            <div class="anexo-data">Pasta de documentos</div>
        </div>
        <a href="<%=enderecoLegado%>" target="_blank" class="btn btn-sm btn-warning">
            <i class="fas fa-external-link-alt me-1"></i> Abrir Pasta
        </a>
    </div>
<% Else %>
    <div class="alert alert-info mb-0">
        <i class="fas fa-info-circle me-2"></i>
        Nenhum anexo encontrado para esta solicitação.
    </div>
<% End If %>
