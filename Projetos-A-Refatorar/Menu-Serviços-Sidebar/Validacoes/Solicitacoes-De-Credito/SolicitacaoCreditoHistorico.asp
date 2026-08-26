<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#include File = "../funcoes/funcoes.asp"-->
<!--#include file="../includes/variaveis.asp" -->
<!--#include file="./util/Util.asp" -->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046

Dim idSolicitacao
idSolicitacao = Request.QueryString("idSolicitacao")

If idSolicitacao = "" Or Not IsNumeric(idSolicitacao) Then
    Response.Write("<div class='alert alert-danger'>ID de solicitação inválido.</div>")
    Response.End
End If

Dim SQL, PIR_CORPORATIVO

SQL = "SELECT DISTINCT "
SQL = SQL & "s.NUM as IdSolicitacao, "
SQL = SQL & "s.LCSOLICIT as LimiteSolicitado, "
SQL = SQL & "s.ClienteGrupo, "
SQL = SQL & "s.AprovadorPrimeiroNivel as ValidadorPrimeiroNivel, "
SQL = SQL & "s.AprovadorSegundoNivel as ValidadorSegundoNivel, "
SQL = SQL & "s.AprovadorTerceiroNivel as ValidadorTerceiroNivel, "
SQL = SQL & "s.AprovadorQuartoNivel as ValidadorQuartoNivel, "
SQL = SQL & "s.AprovadorQuintoNivel as ValidadorQuintoNivel, "
SQL = SQL & "s.AprovadorSextoNivel as ValidadorSextoNivel, "
SQL = SQL & "s.AprovadorSetimoNivel as ValidadorSetimoNivel, "
SQL = SQL & "s.Observacao1 as ObservacaoPrimeiroNivel, "
SQL = SQL & "s.Observacao2 as ObservacaoSegundoNivel, "
SQL = SQL & "s.Observacao3 as ObservacaoTerceiroNivel, "
SQL = SQL & "s.Observacao4 as ObservacaoQuartoNivel, "
SQL = SQL & "s.Observacao5 as ObservacaoQuintoNivel, "
SQL = SQL & "s.Observacao6 as ObservacaoSextoNivel, "
SQL = SQL & "s.Observacao7 as ObservacaoSetimoNivel, "
SQL = SQL & "s.DataAprovacaoPrimeiroNivel as DataValidacaoPrimeiroNivel, "
SQL = SQL & "s.DataAprovacaoSegundoNivel as DataValidacaoSegundoNivel, "
SQL = SQL & "s.DataAprovacaoTerceiroNivel as DataValidacaoTerceiroNivel, "
SQL = SQL & "s.DataAprovacaoQuartoNivel as DataValidacaoQuartoNivel, "
SQL = SQL & "s.DataAprovacaoQuintoNivel as DataValidacaoQuintoNivel, "
SQL = SQL & "s.DataAprovacaoSextoNivel as DataValidacaoSextoNivel, "
SQL = SQL & "s.DataAprovacaoSetimoNivel as DataValidacaoSetimoNivel, "
SQL = SQL & "s.ValidaMigracao "
SQL = SQL & "FROM SOLCREDITOCLIENTE s "
SQL = SQL & "WHERE s.NUM = " & idSolicitacao

On Error Resume Next
SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)

If Err.Number <> 0 OR NOT IsObject(PIR_CORPORATIVO) OR PIR_CORPORATIVO.EOF Then
    Response.Write("<div class='alert alert-danger'>Solicitação não encontrada.</div>")
    Response.End
End If
On Error Goto 0

Function EhDataAnteriorMudancaNomenclaturas()
    If PIR_CORPORATIVO("ValidaMigracao") = "A" Then
        EhDataAnteriorMudancaNomenclaturas = True
    Else
        EhDataAnteriorMudancaNomenclaturas = False
    End If
End Function
%>

<table class="table table-sm table-hover history-table">
    <thead>
        <tr>
            <th>Etapa</th>
            <th>Usuário</th>
            <th>Data</th>
            <th class="text-center">Observação</th>
        </tr>
    </thead>
    <tbody>
        <%
        Dim nomeEtapa1, nomeEtapa2, nomeEtapa3, nomeEtapa4, nomeEtapa5, nomeEtapa6, nomeEtapa7
        Dim deveMostrarEtapa1, deveMostrarEtapa2, deveMostrarEtapa3, deveMostrarEtapa4, deveMostrarEtapa5, deveMostrarEtapa6, deveMostrarEtapa7

        ' Como este componte é utilizado tanto para registros legados quanto para registros atuais, 
        ' devemos realizar tratativas para alterar o nome das etapas e se devemos ou não mostrar a etapa baseado na regra de negócio vigente
        If EhDataAnteriorMudancaNomenclaturas() Then
            nomeEtapa1 = "Financeiro"
            nomeEtapa2 = "Comercial"
            nomeEtapa3 = "Controladoria"
            nomeEtapa4 = "Gerência Geral"
            nomeEtapa5 = "Diretoria Comercial"
            nomeEtapa6 = "CFO"
            nomeEtapa7 = "CEO"

            deveMostrarEtapa1 = True
            deveMostrarEtapa2 = True
            deveMostrarEtapa3 = PIR_CORPORATIVO("LimiteSolicitado") > 50000
            deveMostrarEtapa4 = PIR_CORPORATIVO("LimiteSolicitado") > 100000
            deveMostrarEtapa5 = PIR_CORPORATIVO("LimiteSolicitado") > 300000
            deveMostrarEtapa6 = PIR_CORPORATIVO("LimiteSolicitado") > 1000000
            deveMostrarEtapa7 = Iif(deveMostrarEtapa6 = True, True, False)
        Else
            nomeEtapa1 = "Análise"
            nomeEtapa2 = "Analista de Crédito"
            nomeEtapa3 = "Coordenador de Crédito"
            nomeEtapa4 = "Gerente de Tesouraria"
            nomeEtapa5 = "Divisional Controller"
            nomeEtapa6 = "Gerente de Planta"
            nomeEtapa7 = "CFO"

            deveMostrarEtapa1 = True
            deveMostrarEtapa2 = True
            deveMostrarEtapa3 = PIR_CORPORATIVO("ClienteGrupo") <> 1 AND PIR_CORPORATIVO("LimiteSolicitado") > 180000
            deveMostrarEtapa4 = PIR_CORPORATIVO("ClienteGrupo") <> 1 AND PIR_CORPORATIVO("LimiteSolicitado") > 1000000 
            deveMostrarEtapa5 = PIR_CORPORATIVO("ClienteGrupo") <> 1 AND PIR_CORPORATIVO("LimiteSolicitado") > 3000000 
            deveMostrarEtapa6 = PIR_CORPORATIVO("ClienteGrupo") <> 1 AND Iif(deveMostrarEtapa5 = True, True, False) 
            deveMostrarEtapa7 = PIR_CORPORATIVO("ClienteGrupo") <> 1 AND PIR_CORPORATIVO("LimiteSolicitado") > 5000000
        End If
        %>

        <% If deveMostrarEtapa1 Then %>
        <tr>
            <td><%= nomeEtapa1 %></td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("ValidadorPrimeiroNivel")) Then %>
                    <%= PIR_CORPORATIVO("ValidadorPrimeiroNivel") %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("DataValidacaoPrimeiroNivel")) Then %>
                    <%= FormatarData(PIR_CORPORATIVO("DataValidacaoPrimeiroNivel")) %> 
                    <%= FormatDateTime(PIR_CORPORATIVO("DataValidacaoPrimeiroNivel"), 4) %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td class="text-center">
                <% If Not IsNull(PIR_CORPORATIVO("ObservacaoPrimeiroNivel")) And Trim(PIR_CORPORATIVO("ObservacaoPrimeiroNivel")) <> "" Then 
                    textoComentario = PIR_CORPORATIVO("ObservacaoPrimeiroNivel")
                    textoComentario = Replace(textoComentario, vbCrLf, "\n")
                    textoComentario = Replace(textoComentario, vbCr, "\n")
                    textoComentario = Replace(textoComentario, vbLf, "\n")
                    textoComentario = Replace(textoComentario, "'", "\'")     
                    textoComentario = Replace(textoComentario, """", "\'") 
                %>
                    <button type="button" class="btn btn-sm btn-outline-secondary" 
                        onclick="abrirModalComentario('<%=nomeEtapa1 %>', '<%=textoComentario%>')">
                        <i class="fas fa-comment-alt"></i>
                    </button>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
        </tr>
        <% End If %>
       
        <% If deveMostrarEtapa2 Then %>
        <tr>
            <td><%= nomeEtapa2 %></td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("ValidadorSegundoNivel")) Then %>
                    <%= PIR_CORPORATIVO("ValidadorSegundoNivel") %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("DataValidacaoSegundoNivel")) Then %>
                    <%= FormatarData(PIR_CORPORATIVO("DataValidacaoSegundoNivel")) %> 
                    <%= FormatDateTime(PIR_CORPORATIVO("DataValidacaoSegundoNivel"), 4) %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td class="text-center">
                <% If Not IsNull(PIR_CORPORATIVO("ObservacaoSegundoNivel")) And Trim(PIR_CORPORATIVO("ObservacaoSegundoNivel")) <> "" Then 
                    textoComentario = PIR_CORPORATIVO("ObservacaoSegundoNivel")
                    textoComentario = Replace(textoComentario, vbCrLf, "\n")  
                    textoComentario = Replace(textoComentario, "'", "\'")     
                    textoComentario = Replace(textoComentario, """", "\'") 
                %>
                    <button type="button" class="btn btn-sm btn-outline-secondary" 
                        onclick="abrirModalComentario('<%=nomeEtapa2 %>', '<%=textoComentario%>')">
                        <i class="fas fa-comment-alt"></i>
                    </button>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
        </tr>
        <% End If %>
       
        <% If deveMostrarEtapa3 Then %>
        <tr>
            <td><%= nomeEtapa3 %></td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("ValidadorTerceiroNivel")) Then %>
                    <%= PIR_CORPORATIVO("ValidadorTerceiroNivel") %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("DataValidacaoTerceiroNivel")) Then %>
                    <%= FormatarData(PIR_CORPORATIVO("DataValidacaoTerceiroNivel")) %> 
                    <%= FormatDateTime(PIR_CORPORATIVO("DataValidacaoTerceiroNivel"), 4) %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td class="text-center">
                <% If Not IsNull(PIR_CORPORATIVO("ObservacaoTerceiroNivel")) And Trim(PIR_CORPORATIVO("ObservacaoTerceiroNivel")) <> "" Then 
                    textoComentario = PIR_CORPORATIVO("ObservacaoTerceiroNivel")
                    textoComentario = Replace(textoComentario, vbCrLf, "\n")  
                    textoComentario = Replace(textoComentario, "'", "\'")     
                    textoComentario = Replace(textoComentario, """", "\'") 
                %>
                    <button type="button" class="btn btn-sm btn-outline-secondary" 
                        onclick="abrirModalComentario('<%=nomeEtapa3 %>', '<%=textoComentario%>')">
                        <i class="fas fa-comment-alt"></i>
                    </button>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
        </tr>
        <% End If %>
       
        <% If deveMostrarEtapa4 Then %>
        <tr>
            <td><%= nomeEtapa4 %></td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("ValidadorQuartoNivel")) Then %>
                    <%= PIR_CORPORATIVO("ValidadorQuartoNivel") %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("DataValidacaoQuartoNivel")) Then %>
                    <%= FormatarData(PIR_CORPORATIVO("DataValidacaoQuartoNivel")) %> 
                    <%= FormatDateTime(PIR_CORPORATIVO("DataValidacaoQuartoNivel"), 4) %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td class="text-center">
                <% If Not IsNull(PIR_CORPORATIVO("ObservacaoQuartoNivel")) And Trim(PIR_CORPORATIVO("ObservacaoQuartoNivel")) <> "" Then 
                    textoComentario = PIR_CORPORATIVO("ObservacaoQuartoNivel")
                    textoComentario = Replace(textoComentario, vbCrLf, "\n")  
                    textoComentario = Replace(textoComentario, "'", "\'")     
                    textoComentario = Replace(textoComentario, """", "\'") 
                %>
                    <button type="button" class="btn btn-sm btn-outline-secondary" 
                        onclick="abrirModalComentario('<%=nomeEtapa4 %>', '<%=textoComentario%>')">
                        <i class="fas fa-comment-alt"></i>
                    </button>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
        </tr>
        <% End If %>
       
        <% If deveMostrarEtapa5 Then %>
        <tr>
            <td><%= nomeEtapa5 %></td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("ValidadorQuintoNivel")) Then %>
                    <%= PIR_CORPORATIVO("ValidadorQuintoNivel") %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("DataValidacaoQuintoNivel")) Then %>
                    <%= FormatarData(PIR_CORPORATIVO("DataValidacaoQuintoNivel")) %> 
                    <%= FormatDateTime(PIR_CORPORATIVO("DataValidacaoQuintoNivel"), 4) %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td class="text-center">
                <% If Not IsNull(PIR_CORPORATIVO("ObservacaoQuintoNivel")) And Trim(PIR_CORPORATIVO("ObservacaoQuintoNivel")) <> "" Then 
                    textoComentario = PIR_CORPORATIVO("ObservacaoQuintoNivel")
                    textoComentario = Replace(textoComentario, vbCrLf, "\n")  
                    textoComentario = Replace(textoComentario, "'", "\'")     
                    textoComentario = Replace(textoComentario, """", "\'") 
                %>
                    <button type="button" class="btn btn-sm btn-outline-secondary" 
                        onclick="abrirModalComentario('<%=nomeEtapa5 %>', '<%=textoComentario%>')">
                        <i class="fas fa-comment-alt"></i>
                    </button>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
        </tr>
        <% End If %>

        <% If deveMostrarEtapa6 Then %>
        <tr>
            <td><%= nomeEtapa6 %></td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("ValidadorSextoNivel")) Then %>
                    <%= PIR_CORPORATIVO("ValidadorSextoNivel") %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("DataValidacaoSextoNivel")) Then %>
                    <%= FormatarData(PIR_CORPORATIVO("DataValidacaoSextoNivel")) %> 
                    <%= FormatDateTime(PIR_CORPORATIVO("DataValidacaoSextoNivel"), 4) %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td class="text-center">
                <% If Not IsNull(PIR_CORPORATIVO("ObservacaoSextoNivel")) And Trim(PIR_CORPORATIVO("ObservacaoSextoNivel")) <> "" Then 
                    textoComentario = PIR_CORPORATIVO("ObservacaoSextoNivel")
                    textoComentario = Replace(textoComentario, vbCrLf, "\n")  
                    textoComentario = Replace(textoComentario, "'", "\'")     
                    textoComentario = Replace(textoComentario, """", "\'") 
                %>
                    <button type="button" class="btn btn-sm btn-outline-secondary" 
                        onclick="abrirModalComentario('<%=nomeEtapa6 %>', '<%=textoComentario%>')">
                        <i class="fas fa-comment-alt"></i>
                    </button>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
        </tr>
        <% End If %>

        <% If deveMostrarEtapa7 Then %>
        <tr>
            <td><%= nomeEtapa7 %></td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("ValidadorSetimoNivel")) Then %>
                    <%= PIR_CORPORATIVO("ValidadorSetimoNivel") %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td>
                <% If Not IsNull(PIR_CORPORATIVO("DataValidacaoSetimoNivel")) Then %>
                    <%= FormatarData(PIR_CORPORATIVO("DataValidacaoSetimoNivel")) %> 
                    <%= FormatDateTime(PIR_CORPORATIVO("DataValidacaoSetimoNivel"), 4) %>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
            <td class="text-center">
                <% If Not IsNull(PIR_CORPORATIVO("ObservacaoSetimoNivel")) And Trim(PIR_CORPORATIVO("ObservacaoSetimoNivel")) <> "" Then 
                    textoComentario = PIR_CORPORATIVO("ObservacaoSetimoNivel")
                    textoComentario = Replace(textoComentario, vbCrLf, "\n")  
                    textoComentario = Replace(textoComentario, "'", "\'")     
                    textoComentario = Replace(textoComentario, """", "\'") 
                %>
                    <button type="button" class="btn btn-sm btn-outline-secondary" 
                        onclick="abrirModalComentario('<%=nomeEtapa7 %>', '<%=textoComentario%>')">
                        <i class="fas fa-comment-alt"></i>
                    </button>
                <% Else %>
                    <span class="text-muted">-</span>
                <% End If %>
            </td>
        </tr>
        <% End If %>
    </tbody>
</table>

<!-- Modal para exibir comentário -->
<div class="modal fade" id="modalComentario" tabindex="-1" aria-labelledby="modalComentarioLabel" aria-hidden="true" data-bs-backdrop="static">
    <div class="modal-dialog modal-dialog-centered">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="modalComentarioLabel">Observação</h5>
                <button type="button" class="btn-close" onclick="fecharModalComentario()" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                <h6 id="modalComentarioEtapa" class="text-muted mb-3"></h6>
                <p id="modalComentarioTexto" style="white-space: pre-wrap;"></p>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" onclick="fecharModalComentario()">Fechar</button>
            </div>
        </div>
    </div>
</div>

<script>
    var modalComentario = null;

    function abrirModalComentario(etapa, comentario) {
        document.getElementById('modalComentarioEtapa').innerText = etapa;
        document.getElementById('modalComentarioTexto').innerText = comentario.replace(/\\n/g, '\n');
        
        if (modalComentario === null) {
            modalComentario = new bootstrap.Modal(document.getElementById('modalComentario'), {
                backdrop: 'static',
                keyboard: false
            });
        }
        
        modalComentario.show();
    }

    function fecharModalComentario() {
        if (modalComentario !== null) {
            modalComentario.hide();
        }
    }

    document.getElementById('modalComentario').addEventListener('hidden.bs.modal', function (event) {
        event.stopPropagation();
    });
</script>

<%
If IsObject(PIR_CORPORATIVO) Then
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
End If
%>