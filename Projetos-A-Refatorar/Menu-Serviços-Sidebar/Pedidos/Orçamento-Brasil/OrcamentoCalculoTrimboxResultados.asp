<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../cnn/cnn.asp" -->
<!--#include file = "../cnn2/cnn2.asp" -->
<!--#Include File = "../includes/variaveis.asp"-->
<!--#Include File = "../funcoes/funcoes.asp"-->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046

Dim idOrcamentoSknet, sucessoIds, erroIds
idOrcamentoSknet = Request("idOrcamentoSknet")
sucessoIds = Request("sucessoIds")
erroIds = Request("erroIds")

Dim SQL, PIR_CORPORATIVO
Dim itensSucesso(), itensErro(), idxSucesso, idxErro

Dim idUsuarioLogado, podeSelecionarItem
idUsuarioLogado = Session("CODIGO")
podeSelecionarItem = false

SQL = "SELECT TOP 1 *"
SQL = SQL & " FROM ANALISTAS_ORCAMENTO"
SQL = SQL & " WHERE CODIGO_USUARIO = " & idUsuarioLogado
SQL = SQL & " AND (EXCLUIDO = 0 OR EXCLUIDO IS NULL)"

On Error Resume Next
SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)
If NOT PIR_CORPORATIVO.EOF AND erroIds = "" AND sucessoIds = "" Then
    podeSelecionarItem = true
End If
If isObject(PIR_CORPORATIVO) Then
    PIR_CORPORATIVO.Close()
    SET PIR_CORPORATIVO = NOTHING
End If
On Error Goto 0

ReDim itensSucesso(0)
idxSucesso = 0

If sucessoIds <> "" OR idOrcamentoSknet <> "" Then
    SQL = "SELECT PLANTA as Planta, ORCAMENTO_SKNET AS IdOrcamentoSKnet, ORCAMENTO_TRIM AS IdOrcamentoTrim, "
    SQL = SQL & "CLIENTE AS Cliente, PRODUTO AS Produto, TIPO_CALCULO as TipoCalculo, VALOR_ANALISADO AS ValorAnalisado, "
    SQL = SQL & "DATA_INCLUSAO AS DataInclusao, CASE WHEN VALIDADOR = 1 THEN 1 ELSE 0 END AS Validador FROM ORCAMENTO_SKNET_TRIM WHERE 1=1"

    If sucessoIds <> "" Then
        SQL = SQL & " AND ID IN (" & sucessoIds & ")"
    Else
        SQL = SQL & " AND ORCAMENTO_SKNET = " & idOrcamentoSknet
    End If

    SQL = SQL & " ORDER BY DATA_INCLUSAO DESC"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        While NOT PIR_CORPORATIVO.EOF
            ReDim Preserve itensSucesso(idxSucesso)
            Set itensSucesso(idxSucesso) = Server.CreateObject("Scripting.Dictionary")
            
            itensSucesso(idxSucesso)("Planta") = PIR_CORPORATIVO("Planta") & ""
            itensSucesso(idxSucesso)("IdOrcamentoSKnet") = PIR_CORPORATIVO("IdOrcamentoSKnet") & ""
            itensSucesso(idxSucesso)("IdOrcamentoTrim") = PIR_CORPORATIVO("IdOrcamentoTrim") & ""
            itensSucesso(idxSucesso)("Cliente") = PIR_CORPORATIVO("Cliente") & ""
            itensSucesso(idxSucesso)("Produto") = PIR_CORPORATIVO("Produto") & ""
            itensSucesso(idxSucesso)("TipoCalculo") = PIR_CORPORATIVO("TipoCalculo") & ""
            itensSucesso(idxSucesso)("ValorAnalisado") = Replace(CStr(PIR_CORPORATIVO("ValorAnalisado")), ",", ".")
            itensSucesso(idxSucesso)("DataInclusao") = PIR_CORPORATIVO("DataInclusao") & ""
            itensSucesso(idxSucesso)("Validador") = PIR_CORPORATIVO("Validador") & ""

            idxSucesso = idxSucesso + 1
            PIR_CORPORATIVO.MoveNext()
        Wend
        
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0

End If

ReDim itensErro(0)
idxErro = 0
    
If erroIds <> "" OR idOrcamentoSknet <> "" Then
    SQL = "SELECT PLANTA as Planta, ORCAMENTO_SKNET AS IdOrcamentoSKnet, ORCAMENTO_TRIM AS IdOrcamentoTrim, "
    SQL = SQL & "CLIENTE AS Cliente, PRODUTO AS Produto, TIPO_CUSTO as TipoCusto, CHAVE_CUSTO AS ChaveCusto, "
    SQL = SQL & "DATA_INCLUSAO AS DataInclusao, ESTADO AS Estado, ERRO AS Erro FROM ORCAMENTO_VALIDA_ERRO WHERE 1=1"

    If erroIds <> "" Then
        SQL = SQL & " AND ID IN (" & erroIds & ")"
    Else
        SQL = SQL & " AND ORCAMENTO_SKNET = " & idOrcamentoSknet
    End If

    SQL = SQL & " ORDER BY DATA_INCLUSAO DESC"

    On Error Resume Next
    SET PIR_CORPORATIVO = Conecta("PIR-CORPORATIVO", SQL)

    If Err.Number = 0 AND IsObject(PIR_CORPORATIVO) Then
        While NOT PIR_CORPORATIVO.EOF
            ReDim Preserve itensErro(idxErro)
            Set itensErro(idxErro) = Server.CreateObject("Scripting.Dictionary")
            
            itensErro(idxErro)("Planta") = PIR_CORPORATIVO("Planta") & ""
            itensErro(idxErro)("IdOrcamentoSKnet") = PIR_CORPORATIVO("IdOrcamentoSKnet") & ""
            itensErro(idxErro)("IdOrcamentoTrim") = PIR_CORPORATIVO("IdOrcamentoTrim") & ""
            itensErro(idxErro)("Cliente") = PIR_CORPORATIVO("Cliente") & ""
            itensErro(idxErro)("Produto") = PIR_CORPORATIVO("Produto") & ""
            itensErro(idxErro)("TipoCusto") = PIR_CORPORATIVO("TipoCusto") & ""
            itensErro(idxErro)("ChaveCusto") = PIR_CORPORATIVO("ChaveCusto") & ""
            itensErro(idxErro)("DataInclusao") = PIR_CORPORATIVO("DataInclusao") & ""
            itensErro(idxErro)("Estado") = PIR_CORPORATIVO("Estado") & ""
            itensErro(idxErro)("Erro") = PIR_CORPORATIVO("Erro") & ""

            idxErro = idxErro + 1
            PIR_CORPORATIVO.MoveNext()
        Wend
        
        PIR_CORPORATIVO.Close()
        SET PIR_CORPORATIVO = NOTHING
    End If
    On Error Goto 0

End If

Dim abaSelecionada
abaSelecionada = Request.QueryString("aba")
If abaSelecionada = "" Then
    If idxSucesso > 0 Then
        abaSelecionada = "sucesso"
    ElseIf idxErro > 0 Then
        abaSelecionada = "erro"
    Else
        abaSelecionada = "sucesso"
    End If
End If
%>

<style>
    .tabs-container-modal {
        display: flex;
        padding: 0;
        gap: 0;
        min-width: max-content;
        background-color: #f8f9fa;
        border-bottom: 1px solid #dee2e6;
    }

    .tab-button-modal {
        min-width: 200px;
        padding: 15px 20px;
        background: transparent;
        border: none;
        border-bottom: 3px solid transparent;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 8px;
        color: #6c757d;
        font-weight: 500;
    }

    .tab-button-modal:hover {
        color: #212529;
        background-color: #e9ecef;
    }

    .tab-button-modal.active {
        color: #0d6efd;
        border-bottom-color: #0d6efd;
        background-color: #fff;
    }

    .tab-total-modal {
        display: flex;
        justify-content: center;
        align-items: center;
        background-color: #8d8d8d;
        color: #fff;
        border-radius: 50%;
        width: 20px;
        height: 20px;
        font-size: 0.7rem;
        font-weight: 600;
        padding: 0 6px;
    }

    .tab-total-modal.active {
        background-color: #0d6efd;
    }

    .tab-content-modal {
        background-color: #fff;
    }

    .tab-pane-modal {
        display: none;
    }

    .tab-pane-modal.active {
        display: block;
    }
    
    .no-records {
        padding: 20px;
        text-align: center;
        color: #6c757d;
    }

    .btn-selecionar {
        display: flex;
        justify-content: center;
        align-items: baseline;
        gap: 5px;
        padding: 8px 10px;
        font-size: 0.75rem;
        border: 1px solid #dee2e6;
        background-color: #fff;
        color: #6c757d;
        border-radius: 4px;
        cursor: pointer;
    }

    .btn-selecionar:hover {
        background-color: #0d6efd;
        color: #fff;
        border-color: #0d6efd;
    }

    .btn-selecionar:disabled {
        opacity: 0.5;
        cursor: not-allowed;
    }

    .btn-selecionar.selecionado {
        background-color: #198754;
        color: #fff;
        border-color: #198754;
        cursor: default;
    }

    .btn-selecionar.selecionado:hover {
        background-color: #198754;
        border-color: #198754;
    }
</style>

<div class="tabs-container-modal">
    <button type="button" class="tab-button-modal <%=IIf(abaSelecionada = "sucesso", "active", "")%>" id="tabSucessoBtn" onclick="alterarAbaModalResultados('sucesso')">
        <i class="fas fa-check-circle me-2"></i>
        <span class="tab-title">Calculados com Sucesso</span>
        <div class="tab-total-modal <%=IIf(abaSelecionada = "sucesso", "active", "")%>" id="totalSucessoTab">
            <%=idxSucesso%>
        </div>
    </button>
    <button type="button" class="tab-button-modal <%=IIf(abaSelecionada = "erro", "active", "")%>" id="tabErroBtn" onclick="alterarAbaModalResultados('erro')">
        <i class="fas fa-exclamation-triangle me-2"></i>
        <span class="tab-title">Calculados com Erro</span>
        <div class="tab-total-modal <%=IIf(abaSelecionada = "erro", "active", "")%>" id="totalErroTab">
            <%=idxErro%>
        </div>
    </button>
</div>

<div class="tab-content-modal p-3">
    <div id="secaoSucesso" class="tab-pane-modal <%=IIf(abaSelecionada = "sucesso", "active", "")%>">
        <div class="table-responsive" style="max-height: 550px; overflow-y: auto; border: 1px solid #dee2e6; border-radius: 0.25rem;">
            <table id="tabelaSucesso" class="table table-striped table-sm table-hover mb-0">
                <thead class="thead-light" style="position: sticky; top: 0; z-index: 10; background-color: #f8f9fa;">
                    <tr>
                        <% If podeSelecionarItem Then %>
                        <th style="width: 50px;"></th>
                        <% End If %>
                        <th>ID Sknet</th>
                        <th>ID Trim</th>
                        <th>Produto</th>
                        <th>Cliente</th>
                        <th>Tipo de Cálculo</th>
                        <th>Valor Analisado</th>
                        <th>Data Inclusão</th>
                        <th></th>
                    </tr>
                </thead>
                
                <% IF idxSucesso > 0 THEN %>
                <tbody>
                    <% 
                    Dim s
                    For s = 0 To UBound(itensSucesso)
                    %>
                        <tr data-sknet="<%=itensSucesso(s)("IdOrcamentoSKnet")%>" data-trim="<%=itensSucesso(s)("IdOrcamentoTrim")%>">
                            <% If podeSelecionarItem Then %>
                            <td class="text-center">
                                <% If CStr(itensSucesso(s)("Validador")) = "1" Then %>
                                    <button onclick="selecionarItem('<%=itensSucesso(s)("IdOrcamentoSKnet")%>', '<%=itensSucesso(s)("IdOrcamentoTrim")%>')"
                                        class="btn-selecionar selecionado" 
                                        title="Item Selecionado">
                                        <i class="fas fa-check"></i> 
                                    </button>
                                <% Else %>
                                    <button onclick="selecionarItem('<%=itensSucesso(s)("IdOrcamentoSKnet")%>', '<%=itensSucesso(s)("IdOrcamentoTrim")%>')"
                                        class="btn-selecionar" 
                                        title="Selecionar este item">
                                        <i class="far fa-circle"></i> 
                                    </button>
                                <% End If %>
                            </td>
                            <% End If %>
                            <td><%=itensSucesso(s)("IdOrcamentoSKnet")%></td>
                            <td><%=itensSucesso(s)("IdOrcamentoTrim")%></td>
                            <td><%=itensSucesso(s)("Produto")%></td>
                            <td><%=itensSucesso(s)("Cliente")%></td>
                            <td><%=itensSucesso(s)("TipoCalculo")%></td>
                            <td><%=itensSucesso(s)("ValorAnalisado")%></td>
                            <td><%=itensSucesso(s)("DataInclusao")%></td>
                            <td class="text-center">
                                <button onclick="visualizaDetalhesTrim('<%=itensSucesso(s)("IdOrcamentoSKnet")%>', '<%=itensSucesso(s)("IdOrcamentoTrim")%>', '<%=itensSucesso(s)("Planta")%>')"                                    
                                    class="btn btn-sm btn-primary" 
                                    title="Visualizar Detalhes">
                                    <i class="fas fa-eye"></i>
                                </button>
                            </td>
                        </tr>
                    <% Next %>
                </tbody>
                <% END IF %>
            </table>
            
            <% IF idxSucesso = 0 THEN %>
                <div class="no-records">
                    Nenhum item
                </div>
            <% END IF %>
        </div>
    </div>
    
    <div id="secaoErro" class="tab-pane-modal <%=IIf(abaSelecionada = "erro", "active", "")%>">
        <div class="table-responsive" style="max-height: 550px; overflow-y: auto; border: 1px solid #dee2e6; border-radius: 0.25rem;">
            <table id="tabelaErro" class="table table-striped table-sm table-hover mb-0">
                <thead class="thead-light" style="position: sticky; top: 0; z-index: 10; background-color: #f8f9fa;">
                    <tr>
                        <th>ID Sknet</th>
                        <th>ID Trim</th>
                        <th>Produto</th>
                        <th>Cliente</th>
                        <th>Mensagem de Erro</th>
                        <th>Tipo</th>
                        <th>Chave</th>
                        <th>Data Inclusão</th>
                    </tr>
                </thead>
                
                <% IF idxErro > 0 THEN %>
                <tbody>
                    <% 
                    Dim e
                    For e = 0 To UBound(itensErro)
                    %>
                        <tr>
                            <td><%=itensErro(e)("IdOrcamentoSKnet")%></td>
                            <td><%=itensErro(e)("IdOrcamentoTrim")%></td>
                            <td><%=itensErro(e)("Produto")%></td>
                            <td><%=itensErro(e)("Cliente")%></td>
                            <td><%=itensErro(e)("Erro")%></td>
                            <td><%=itensErro(e)("TipoCusto")%></td>
                            <td><%=itensErro(e)("ChaveCusto")%></td>
                            <td><%=itensErro(e)("DataInclusao")%></td>
                        </tr>
                    <% Next %>
                </tbody>
                <% END IF %>
            </table>
            
            <% IF idxErro = 0 THEN %>
                <div class="no-records">
                    Nenhum item 
                </div>
            <% END IF %>
        </div>
    </div>
</div>

<script>
function alterarAbaModalResultados(aba) {
    var botoes = document.querySelectorAll('.tab-button-modal');
    for (var i = 0; i < botoes.length; i++) {
        botoes[i].classList.remove('active');
    }
    
    var totais = document.querySelectorAll('.tab-total-modal');
    for (var i = 0; i < totais.length; i++) {
        totais[i].classList.remove('active');
    }
    
    var paineis = document.querySelectorAll('.tab-pane-modal');
    for (var i = 0; i < paineis.length; i++) {
        paineis[i].classList.remove('active');
        paineis[i].style.display = 'none';
    }
    
    if (aba === 'sucesso') {
        document.getElementById('tabSucessoBtn').classList.add('active');
        document.getElementById('totalSucessoTab').classList.add('active');
        document.getElementById('secaoSucesso').classList.add('active');
        document.getElementById('secaoSucesso').style.display = 'block';
    } else if (aba === 'erro') {
        document.getElementById('tabErroBtn').classList.add('active');
        document.getElementById('totalErroTab').classList.add('active');
        document.getElementById('secaoErro').classList.add('active');
        document.getElementById('secaoErro').style.display = 'block';
    }
}

function selecionarItem(paramOrcamentoSKNet, paramOrcamentoTrim) {
    var $botoes = $('.btn-selecionar');
    $botoes.prop('disabled', true);

    $.ajax({
        url: 'OrcamentoTrimboxDestacarRepositorio.asp?v=' + Math.random(),
        type: 'POST',
        data: {
            paramOrcamentoSKNet: paramOrcamentoSKNet,
            paramOrcamentoTrim: paramOrcamentoTrim
        },
        success: function(response) {
            if (response.sucesso == false) {
                alert("Erro ao selecionar item")
                return;
            }

            $botoes.removeClass('selecionado')
                   .html('<i class="far fa-circle"></i>')
                   .attr('title', 'Selecionar este item')
                   .prop('disabled', false);
            
            var $linhaSelecionada = $('#tabelaSucesso tbody tr[data-sknet="' + paramOrcamentoSKNet + '"][data-trim="' + paramOrcamentoTrim + '"]');
            
            $linhaSelecionada.find('.btn-selecionar')
                            .addClass('selecionado')
                            .html('<i class="fas fa-check"></i>')
                            .attr('title', 'Item Selecionado');
        },
        error: function(xhr, status, error) {
            alert('Erro ao selecionar o item. Por favor, tente novamente.');
            $botoes.prop('disabled', false);
        }
    });
}
</script>