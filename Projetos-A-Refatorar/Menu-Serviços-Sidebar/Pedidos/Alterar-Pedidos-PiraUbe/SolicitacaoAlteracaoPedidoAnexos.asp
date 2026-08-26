<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../cnn2/cnn2.asp" -->
<!--#include file="../../../funcoes/funcoes.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046

Dim idSolicitacao
idSolicitacao = Request.QueryString("idSolicitacao")

If idSolicitacao = "" Or Not IsNumeric(idSolicitacao) Then
    Response.Write "<div class='alert alert-warning'><i class='fas fa-exclamation-triangle mr-2'></i>ID da solicitação inválido.</div>"
    Response.End
End If

Dim SQL, BEN_CORPORATIVO

SQL = "SELECT Id, IdSolicitacaoAlteracaoPedido, Nome, Caminho, CONVERT(VARCHAR(19), DataCriacao, 120) as DataCriacao " 
SQL = SQL & "FROM SolicitacaoAlteracaoPedidoEvidencia "
SQL = SQL & "WHERE IdSolicitacaoAlteracaoPedido = " & idSolicitacao & " " 
SQL = SQL & "ORDER BY DataCriacao DESC"

On Error Resume Next
SET BEN_CORPORATIVO = Conecta("BEN-CORPORATIVO", SQL)

If Err.Number <> 0 OR Not IsObject(BEN_CORPORATIVO) Then
    Response.Write "<div class='alert alert-danger'><i class='fas fa-exclamation-triangle mr-2'></i>Erro ao consultar anexos.</div>"
    On Error Goto 0
    Response.End
End If
On Error Goto 0

If BEN_CORPORATIVO.EOF Then
    Response.Write "<div class='alert alert-info mb-0'><i class='fas fa-info-circle mr-2'></i>Nenhum anexo encontrado para esta solicitação.</div>"
Else
    Response.Write "<div class='list-group'>"
    
    While Not BEN_CORPORATIVO.EOF
        Dim nomeArquivo, caminhoArquivo, dataCriacao, dataFormatada
        nomeArquivo = Trim(BEN_CORPORATIVO("Nome"))
        caminhoArquivo = Trim(BEN_CORPORATIVO("Caminho"))
        dataCriacao = Trim(BEN_CORPORATIVO("DataCriacao"))
        
        If dataCriacao <> "" Then
            dataFormatada = dataCriacao
        Else
            dataFormatada = "-"
        End If
%>
        <div class="anexo-item d-flex align-items-center">
            <i class="fas fa-file-alt anexo-icon"></i>
            <div class="anexo-info">
                <div class="anexo-nome"><%=Server.HTMLEncode(nomeArquivo)%></div>
                <div class="anexo-data">Adicionado em: <%=dataFormatada%></div>
            </div>

            <a href="<%="/Clientes_Docs/" & caminhoArquivo%>" target="_blank" class="btn btn-sm btn-primary">
                <i class="fas fa-external-link-alt mr-1"></i> Abrir
            </a>
        </div>
<%
        BEN_CORPORATIVO.MoveNext
    Wend
    
    Response.Write "</div>"
End If

If IsObject(BEN_CORPORATIVO) Then
    BEN_CORPORATIVO.Close()
    SET BEN_CORPORATIVO = NOTHING
End If
%>
