<!--#include file="../includes/variaveis.asp" -->
<!--#include file="../funcoes/funcoes.asp" -->
<!--#include file="../cnn2/cnn2.asp" -->
<%session.LCID = 1046 

' Variáveis de busca
Dim nomeusuario, emailusuario, etapa
if Request("Nome") <> "" then
   nomeusuario = Request("Nome")
end if
if Request("Email") <> "" then
   emailusuario = Request("Email")  
end if

' Dicionário para armazenar emails únicos encontrados
Dim emailsEncontrados, emailsSelecionados
Set emailsEncontrados = CreateObject("Scripting.Dictionary")
Set emailsSelecionados = CreateObject("Scripting.Dictionary")

Dim palavras, palavra, i
emailusuario = Trim(emailusuario)
emailusuario = RemoverDominioEmail(emailusuario)
nomeusuario = UCase(TRIM(nomeusuario))
nomeusuario = Replace(Replace(Replace(nomeusuario, "'", "''"), "%", "[%]"), "_", "[_]")
palavras = Split(Trim(nomeusuario), " ")

Function CriarCondicaoNome(palavras, campo)
    Dim condicoes, palavra, i
    condicoes = ""
    For i = 0 To UBound(palavras)
        palavra = Trim(palavras(i))
        If palavra <> "" Then
            condicoes = condicoes & " AND " & campo & " LIKE '%" & palavra & "%'"
        End If
    Next
    CriarCondicaoNome = condicoes
End Function

Function RemoverDominioEmail(email)
    Dim posicaoArroba
    posicaoArroba = InStr(email, "@")
    
    If posicaoArroba > 0 Then
        RemoverDominioEmail = Left(email, posicaoArroba - 1)
    Else
        RemoverDominioEmail = email
    End If
End Function


etapa = Request("etapa")
if etapa = "" then
    etapa = "pesquisa"
end if

' Se estamos na etapa de exibição, recupera os emails selecionados
if etapa = "exibir" then
    Dim emailsParam, arrEmails, e
    emailsParam = Request("emailsSelecionados")
    if emailsParam <> "" then
        arrEmails = Split(emailsParam, ",")
        For Each e in arrEmails
            emailsSelecionados(Trim(e)) = True
        Next
    end if
end if


' ==============================================
'       PRÉ-CONSULTA - COLETANDO TODOS OS EMAILS
' ==============================================

IF (nomeusuario <> "" or emailusuario <> "") THEN 
    ' --- PRÉ-CONSULTA SWNET ---
    SQL = "SELECT DISTINCT CRM_EMAILPRIMARIO as Email FROM CRM_DTI WHERE 1=1"
    If nomeusuario <> "" Then SQL = SQL & CriarCondicaoNome(palavras, "CRM_NOME")
    If emailusuario <> "" Then SQL = SQL & " AND CRM_EMAILPRIMARIO LIKE '%" & emailusuario & "%'"
        
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    if isObject(PIR_HELPDESK) then
        While NOT PIR_HELPDESK.EOF
            If Not IsNull(PIR_HELPDESK("Email")) And Trim(PIR_HELPDESK("Email")) <> "" Then
                emailsEncontrados(Trim(PIR_HELPDESK("Email"))) = True
            End If
            PIR_HELPDESK.MoveNext()
        Wend
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    end if

    ' --- PRÉ-CONSULTA TRIMBOX BENTO ---
    SQL = "Select DISTINCT Email FROM Usuario WHERE 1=1"
    If nomeusuario <> "" Then SQL = SQL & CriarCondicaoNome(palavras, "NomeCompleto")
    If emailusuario <> "" Then SQL = SQL & " AND Email LIKE '%" & emailusuario & "%'"
    
    SET BEN_TRIMBOX = Conecta("BEN-TRIMBOX", SQL)
    if isObject(BEN_TRIMBOX) then
        While NOT BEN_TRIMBOX.EOF
            If Not IsNull(BEN_TRIMBOX("Email")) And Trim(BEN_TRIMBOX("Email")) <> "" Then
                emailsEncontrados(Trim(BEN_TRIMBOX("Email"))) = True
            End If
            BEN_TRIMBOX.MoveNext()
        Wend
        BEN_TRIMBOX.Close()
        SET BEN_TRIMBOX = NOTHING
    end if

    ' --- PRÉ-CONSULTA TRIMBOX PIRAPETINGA ---
    SQL = "Select DISTINCT Email FROM Usuario WHERE 1=1"
    If nomeusuario <> "" Then SQL = SQL & CriarCondicaoNome(palavras, "NomeCompleto")
    If emailusuario <> "" Then SQL = SQL & " AND Email LIKE '%" & emailusuario & "%'"
    
    SET PIR_TRIMBOX = Conecta("PIR-TRIMBOX", SQL)
    if isObject(PIR_TRIMBOX) then
        While NOT PIR_TRIMBOX.EOF
            If Not IsNull(PIR_TRIMBOX("Email")) And Trim(PIR_TRIMBOX("Email")) <> "" Then
                emailsEncontrados(Trim(PIR_TRIMBOX("Email"))) = True
            End If
            PIR_TRIMBOX.MoveNext()
        Wend
        PIR_TRIMBOX.Close()
        SET PIR_TRIMBOX = NOTHING
    end if

    ' --- PRÉ-CONSULTA TRIMBOX UBERABA ---
    SQL = "Select DISTINCT Email FROM Usuario WHERE 1=1"
    If nomeusuario <> "" Then SQL = SQL & CriarCondicaoNome(palavras, "NomeCompleto")
    If emailusuario <> "" Then SQL = SQL & " AND Email LIKE '%" & emailusuario & "%'"
    
    SET UBE_TRIMBOX = Conecta("UBE-TRIMBOX", SQL)
    if isObject(UBE_TRIMBOX) then
        While NOT UBE_TRIMBOX.EOF
            If Not IsNull(UBE_TRIMBOX("Email")) And Trim(UBE_TRIMBOX("Email")) <> "" Then
                emailsEncontrados(Trim(UBE_TRIMBOX("Email"))) = True
            End If
            UBE_TRIMBOX.MoveNext()
        Wend
        UBE_TRIMBOX.Close()
        SET UBE_TRIMBOX = NOTHING
    end if

    ' --- PRÉ-CONSULTA TRIMBOX MARANGUAPE ---
    SQL = "Select DISTINCT Email FROM Usuario WHERE 1=1"
    If nomeusuario <> "" Then SQL = SQL & CriarCondicaoNome(palavras, "NomeCompleto")
    If emailusuario <> "" Then SQL = SQL & " AND Email LIKE '%" & emailusuario & "%'"

    SET FOR_TRIMBOX = Conecta("FOR-TRIMBOX", SQL)
    if isObject(FOR_TRIMBOX) then
        While NOT FOR_TRIMBOX.EOF
            If Not IsNull(FOR_TRIMBOX("Email")) And Trim(FOR_TRIMBOX("Email")) <> "" Then
                emailsEncontrados(Trim(FOR_TRIMBOX("Email"))) = True
            End If
            FOR_TRIMBOX.MoveNext()
        Wend
        FOR_TRIMBOX.Close()
        SET FOR_TRIMBOX = NOTHING
    end if

    ' --- PRÉ-CONSULTA TRIMBOX SAQUAREMA ---
    SQL = "Select DISTINCT Email FROM Usuario WHERE 1=1"
    If nomeusuario <> "" Then SQL = SQL & CriarCondicaoNome(palavras, "NomeCompleto")
    If emailusuario <> "" Then SQL = SQL & " AND Email LIKE '%" & emailusuario & "%'"
    
    SET SAQ_TRIMBOX = Conecta("SAQ-TRIMBOX", SQL)
    if isObject(SAQ_TRIMBOX) then
        While NOT SAQ_TRIMBOX.EOF
            If Not IsNull(SAQ_TRIMBOX("Email")) And Trim(SAQ_TRIMBOX("Email")) <> "" Then
                emailsEncontrados(Trim(SAQ_TRIMBOX("Email"))) = True
            End If
            SAQ_TRIMBOX.MoveNext()
        Wend
        SAQ_TRIMBOX.Close()
        SET SAQ_TRIMBOX = NOTHING
    end if

    ' --- PRÉ-CONSULTA TRIMPAPEL BENTO ---
    SQL = "Select DISTINCT Email FROM Usuarios WHERE 1=1"
    If nomeusuario <> "" Then SQL = SQL & CriarCondicaoNome(palavras, "NomeCompleto")
    If emailusuario <> "" Then SQL = SQL & " AND Email LIKE '%" & emailusuario & "%'"
    
    SET BEN_TRIMPAPER = Conecta("BEN-TRIMPAPER", SQL)
    if isObject(BEN_TRIMPAPER) then
        While NOT BEN_TRIMPAPER.EOF
            If Not IsNull(BEN_TRIMPAPER("Email")) And Trim(BEN_TRIMPAPER("Email")) <> "" Then
                emailsEncontrados(Trim(BEN_TRIMPAPER("Email"))) = True
            End If
            BEN_TRIMPAPER.MoveNext()
        Wend
        BEN_TRIMPAPER.Close()
        SET BEN_TRIMPAPER = NOTHING
    end if

    ' --- PRÉ-CONSULTA TRIMPAPEL PIRAPETINGA ---
    SQL = "Select DISTINCT Email FROM Usuarios WHERE 1=1"
    If nomeusuario <> "" Then SQL = SQL & CriarCondicaoNome(palavras, "NomeCompleto")
    If emailusuario <> "" Then SQL = SQL & " AND Email LIKE '%" & emailusuario & "%'"
    
    SET PIR_TRIMPAPER = Conecta("PIR-TRIMPAPER", SQL)
    if isObject(PIR_TRIMPAPER) then
        While NOT PIR_TRIMPAPER.EOF
            If Not IsNull(PIR_TRIMPAPER("Email")) And Trim(PIR_TRIMPAPER("Email")) <> "" Then
                emailsEncontrados(Trim(PIR_TRIMPAPER("Email"))) = True
            End If
            PIR_TRIMPAPER.MoveNext()
        Wend
        PIR_TRIMPAPER.Close()
        SET PIR_TRIMPAPER = NOTHING
    end if

    ' --- PRÉ-CONSULTA TRIMPAPEL UBERABA ---
    SQL = "Select DISTINCT Email FROM Usuarios WHERE 1=1"
    If nomeusuario <> "" Then SQL = SQL & CriarCondicaoNome(palavras, "NomeCompleto")
    If emailusuario <> "" Then SQL = SQL & " AND Email LIKE '%" & emailusuario & "%'"
    
    SET UBE_TRIMPAPER = Conecta("UBE-TRIMPAPER", SQL)
    if isObject(UBE_TRIMPAPER) then
        While NOT UBE_TRIMPAPER.EOF
            If Not IsNull(UBE_TRIMPAPER("Email")) And Trim(UBE_TRIMPAPER("Email")) <> "" Then
                emailsEncontrados(Trim(UBE_TRIMPAPER("Email"))) = True
            End If
            UBE_TRIMPAPER.MoveNext()
        Wend
        UBE_TRIMPAPER.Close()
        SET UBE_TRIMPAPER = NOTHING
    end if

    ' --- PRÉ-CONSULTA PROTHEUS PIRAPETINGA ---
    SQL = "Select DISTINCT USR_EMAIL as Email FROM SYS_USR WHERE 1=1"
    If nomeusuario <> "" Then SQL = SQL & CriarCondicaoNome(palavras, "USR_NOME")
    If emailusuario <> "" Then SQL = SQL & " AND USR_EMAIL LIKE '%" & emailusuario & "%'"
    
    SET PIR_DADOSADV = Conecta("PIR-DADOSADV", SQL)
    if isObject(PIR_DADOSADV) then
        While NOT PIR_DADOSADV.EOF
            If Not IsNull(PIR_DADOSADV("Email")) And Trim(PIR_DADOSADV("Email")) <> "" Then
                emailsEncontrados(Trim(PIR_DADOSADV("Email"))) = True
            End If
            PIR_DADOSADV.MoveNext()
        Wend
        PIR_DADOSADV.Close()
        SET PIR_DADOSADV = NOTHING
    end if

    ' --- PRÉ-CONSULTA PROTHEUS SAQUAREMA ---
    SQL = "Select DISTINCT USR_EMAIL as Email FROM SYS_USR WHERE 1=1"
    If nomeusuario <> "" Then SQL = SQL & CriarCondicaoNome(palavras, "USR_NOME")
    If emailusuario <> "" Then SQL = SQL & " AND USR_EMAIL LIKE '%" & emailusuario & "%'"
    
    SET SAQ_DADOSADV = Conecta("SAQ-DADOSADV", SQL)
    if isObject(SAQ_DADOSADV) then
        While NOT SAQ_DADOSADV.EOF
            If Not IsNull(SAQ_DADOSADV("Email")) And Trim(SAQ_DADOSADV("Email")) <> "" Then
                emailsEncontrados(Trim(SAQ_DADOSADV("Email"))) = True
            End If
            SAQ_DADOSADV.MoveNext()
        Wend
        SAQ_DADOSADV.Close()
        SET SAQ_DADOSADV = NOTHING
    end if

    ' --- PRÉ-CONSULTA PROTHEUS MARANGUAPE ---
    SQL = "Select DISTINCT USR_EMAIL as Email FROM SYS_USR WHERE 1=1"
    If nomeusuario <> "" Then SQL = SQL & CriarCondicaoNome(palavras, "USR_NOME")
    If emailusuario <> "" Then SQL = SQL & " AND USR_EMAIL LIKE '%" & emailusuario & "%'"
    
    SET FOR_DADOSADV = Conecta("FOR-DADOSADV", SQL)
    if isObject(FOR_DADOSADV) then
        While NOT FOR_DADOSADV.EOF
            If Not IsNull(FOR_DADOSADV("Email")) And Trim(FOR_DADOSADV("Email")) <> "" Then
                emailsEncontrados(Trim(FOR_DADOSADV("Email"))) = True
            End If
            FOR_DADOSADV.MoveNext()
        Wend
        FOR_DADOSADV.Close()
        SET FOR_DADOSADV = NOTHING
    end if

    ' --- PRÉ-CONSULTA PROTHEUS BENTO ---
    SQL = "Select DISTINCT USR_EMAIL as Email FROM SYS_USR WHERE 1=1"
    If nomeusuario <> "" Then SQL = SQL & CriarCondicaoNome(palavras, "USR_NOME")
    If emailusuario <> "" Then SQL = SQL & " AND USR_EMAIL LIKE '%" & emailusuario & "%'"
    
    SET BEN_DADOSADV = Conecta("BEN-DADOSADV", SQL)
    if isObject(BEN_DADOSADV) then
        While NOT BEN_DADOSADV.EOF
            If Not IsNull(BEN_DADOSADV("Email")) And Trim(BEN_DADOSADV("Email")) <> "" Then
                emailsEncontrados(Trim(BEN_DADOSADV("Email"))) = True
            End If
            BEN_DADOSADV.MoveNext()
        Wend
        BEN_DADOSADV.Close()
        SET BEN_DADOSADV = NOTHING
    end if

END IF

if emailsEncontrados.Count = 1 AND etapa = "pesquisa" then
    etapa = "exibir"
    Dim unicoEmail
    For Each unicoEmail in emailsEncontrados.Keys
        emailsSelecionados(unicoEmail) = True
    Next
end if

if emailsEncontrados.Count > 1 AND etapa = "pesquisa" then
    etapa = "selecao"
end if

if emailsEncontrados.Count = 0 AND etapa = "pesquisa" then
    etapa = "exibir"
end if
%>

<!DOCTYPE html>
<html lang="pt-br">
    <head>
        <title>Smurfit Kappa</title>
        <meta charset="utf-8"/>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
        <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />    
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

            let selecoes = new Set();

            function SelecionarUsuario(checkbox) {
                const userId = checkbox.dataset.userId;
                const sistema = checkbox.dataset.sistema;
                const key = `${sistema}:${userId}`;
                
                if (checkbox.checked) {
                    selecoes.add(key);
                } else {
                    selecoes.delete(key);
                    document.getElementById('selectAll').checked = false;
                }
                
                atualizarVisibilidadeBotoes();
            }

            function SelecionarTodos(checkboxGeral) {
                const checkboxes = document.querySelectorAll('tbody input[type="checkbox"]:not(:disabled)');
                
                if (checkboxGeral.checked) {
                    checkboxes.forEach(cb => {
                        cb.checked = true;
                        const key = `${cb.dataset.sistema}:${cb.dataset.userId}`;
                        selecoes.add(key);
                    });
                } else {
                    checkboxes.forEach(cb => cb.checked = false);
                    selecoes.clear();
                }
                
                atualizarVisibilidadeBotoes();
            }

            function atualizarVisibilidadeBotoes() {
                const botoes = document.querySelectorAll('.btn-confirmar');
                const existemSelecoes = selecoes.size > 0;
                
                botoes.forEach(btn => {
                    btn.classList.toggle('hidden', !existemSelecoes);
                });
            }

            function agruparPorSistema() {
                const agrupado = {};
                
                selecoes.forEach(key => {
                    const [sistema, userId] = key.split(':');
                    if (!agrupado[sistema]) {
                        agrupado[sistema] = [];
                    }
                    agrupado[sistema].push(userId);
                });
                
                return agrupado;
            }

            function AlterarUsuarios(acao) {
                if (selecoes.size === 0) {
                    alert('Nenhum usuário selecionado');
                    return;
                }
                
                const emailsSelecionados = new Set();
                selecoes.forEach(key => {
                    const checkbox = document.querySelector(`input[data-user-id="${key.split(':')[1]}"][data-sistema="${key.split(':')[0]}"]`);
                    if (checkbox) {
                        emailsSelecionados.add(checkbox.dataset.email.toUpperCase());
                    }
                });
                
                let params = new URLSearchParams({
                    Acao: acao
                });
                
                params.append('emailsSelecionados', Array.from(emailsSelecionados).join(','));
                
                const agrupado = agruparPorSistema();
                Object.entries(agrupado).forEach(([sistema, ids]) => {
                    params.append(sistema, ids.join(','));
                });
                
                const url = `CancelamentoUsuarioRepositorio.asp?${params.toString()}`;
                
                if (confirm(`Deseja realmente ${acao} ${emailsSelecionados.size} usuário(s) único(s) em ${selecoes.size} sistema(s)?`)) {
                    window.location.href = url;
                }
            }

            function PesquisarSistemasUsuario() {
                const nome = document.getElementById('nome').value;
                const email = document.getElementById('Email').value;

                if(nome.length < 3 && email.length < 3) {
                    alert('Informe ao menos 3 caracteres no Nome ou E-mail para realizar a pesquisa.');
                    return;
                }

                const params = new URLSearchParams({ Nome: nome, Email: email });
                window.location.href = `CancelamentoUsuario.asp?${params.toString()}`;
            }
            
            function SelecionarEmails() {
                const checkboxes = document.querySelectorAll('input[name="emailSelecao"]:checked');
                if (checkboxes.length === 0) {
                    alert('Selecione pelo menos um e-mail para continuar.');
                    return;
                }
                
                const emailsSelecionados = [];
                checkboxes.forEach(cb => {
                    emailsSelecionados.push(cb.value);
                });
                
                const params = new URLSearchParams({
                    Nome: document.getElementById('nome').value || '',
                    Email: document.getElementById('Email').value || '',
                    etapa: 'exibir',
                    emailsSelecionados: emailsSelecionados.join(',')
                });
                
                window.location.href = `CancelamentoUsuario.asp?${params.toString()}`;
            }

             function exportarPDF() {
                const { jsPDF } = window.jspdf;
                const doc = new jsPDF('landscape');
                const dataAtual = new Date().toLocaleDateString('pt-BR');
                const acao =  "<%=acao%>";
                const pesquisaNome =  "<%=nomeusuario%>";
                const pesquisaEmail =  "<%=emailusuario%>";

                const tabela = document.getElementById('GradeVazia');
                const linhas = tabela.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
                
                const dados = [];
                for (let i = 0; i < linhas.length; i++) {
                    const celulas = linhas[i].getElementsByTagName('td');
                    const linha = [];
                    for (let j = 0; j < celulas.length; j++) {
                        linha.push(celulas[j].innerText);
                    }
                    dados.push(linha);
                }
                
                doc.setFontSize(16);
                doc.text(`Relatório de Desligamento de Usuário`, 14, 15);
                
                doc.setFontSize(10);
                doc.text('Data de geração: ' + dataAtual + ' | ' + "Nenhum usuário encontrado para os critérios informados", 14, 22);
                doc.text('Email: ' + pesquisaEmail, 14, 30);
                doc.text('Nome: ' + pesquisaNome, 14, 35);

                doc.autoTable({
                    head: [['Sistema', 'Status']],
                    body: dados,
                    startY: 42,
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
        </script>
        <style type="text/css">
            body{
                font-size:80% !important;
            }
            .header {
                margin: 25px 0;
            }
            input[type="checkbox"] {
                width: 1rem;
                height: 1rem;
                cursor: pointer;
            }
           
            .hidden {
                display: none;
            }
            .badge {
                padding: 4px 10px;
                width: fit-content;
                border-radius: 5px;
                font-weight: 700;
                font-size: .7rem;
            }
            .active {
                background: #b5e9b5;
                color: #003b00;
            }
            .inactive {
                background: #e9b5b5ff;
                color: #3b0000ff;
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
                margin-bottom: 20px;
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
        </style>
    </head>
    <body>
        <form name="Pesquisar" id="Pesquisar" method="post">        
            <div class="container">
                <div class="header border-bottom">
                    <h3> Desligamento de Usuário</h3>
                </div>
                
                <div class="filtro-container">                    
                    <div class="col-md-4">
                        <label for="nome" class="form-label"><strong>Nome do Usuário:</strong></label>
                        <input type="text" class="form-control form-control-sm" name="nome" id="nome" value="<%=Trim(Request("cliente"))%>" placeholder="Nome do Usuário..." />
                    </div>

                    <div class="col-md-4">
                        <label for="Email" class="form-label"><strong>Email do Usuário:</strong></label>
                        <input type="text" class="form-control form-control-sm" name="Email" id="Email" value="<%=Trim(Request("codigo"))%>" placeholder="E-mail do Usuário..." />
                    </div>

                    <div class="col-md-4">
                        <input type="button" class="btn btn-sm btn-primary" name="btnPesquisar" id="btnPesquisar" value="Pesquisar" onclick="PesquisarSistemasUsuario()"/>
                    </div>
                </div>

                <% IF etapa = "exibir" AND emailsEncontrados.Count = 0 AND (nomeusuario <> "" or emailusuario <> "") THEN %>
                    <div class="alert alert-warning mt-3" role="alert">
                        <h6 class="alert-heading">Nenhum usuário encontrado para os critérios informados</h6>
                        
                        <% IF nomeusuario <> "" THEN %>
                            <label><strong>Nome</strong>: <%=nomeusuario%></label>
                            <br>
                        <% END IF %>

                        <% IF emailusuario <> "" THEN %>
                            <label><strong>Email</strong>: <%=emailusuario%></label>
                            <br>
                        <% END IF %>
                    </div>
                <% END IF %>
        
                <% IF etapa = "selecao" AND emailsEncontrados.Count > 1 THEN %>
                    <div class="row">
                        <h5>Foram encontrados <%=emailsEncontrados.Count%> usuários únicos. Selecione quais deseja visualizar:</h5>
                    
                        <%
                        Dim emailItem, contador
                        contador = 1
                        For Each emailItem in emailsEncontrados.Keys
                        %>
                            <div style="display: flex; align-items: center; gap: 0.5rem;">
                                <input type="checkbox" name="emailSelecao" value="<%=emailItem%>" id="email_<%=contador%>">
                                <label style="font-size: 1rem" class="form-check-label" for="email_<%=contador%>">
                                    <%=emailItem%>
                                </label>
                            </div>
                        <%
                            contador = contador + 1
                        Next
                        %>
                        <div class="mt-3">
                            <button type="button" class="btn btn-sm btn-success me-2" onclick="SelecionarEmails()">
                                Continuar com os selecionados
                            </button>
                            <button type="button" class="btn btn-sm btn-secondary" onclick="window.location.href='CancelamentoUsuario.asp'">
                                Nova pesquisa
                            </button>
                        </div>
                    </div>
                <% END IF %>

                <% IF etapa = "exibir" AND emailsEncontrados.Count = 0 AND (nomeusuario <> "" or emailusuario <> "") THEN %>
                    <div class="row mb-3 align-middle">
                        <div class="col-md-12 mb-3">
                            <button type="button" class="btn btn-sm btn-danger mr-3" onclick="exportarPDF()">
                                <i class="fas fa-file-pdf"></i> Exportar para PDF
                            </button>
                        </div>

                        <div class="col-md-12">
                            <table class="table table-bordered" name="Grade" id="GradeVazia">
                                <thead>
                                    <tr>
                                        <th class="text-left">Sistema</th>
                                        <th class="text-left">Status</th>
                                    </tr>              
                                </thead>
                                <tbody>
                                    <tr>
                                        <td class="text-left">SWNET</td>
                                        <td class="text-left">SEM REGISTRO</td>
                                    </tr>
                                    <tr>
                                        <td class="text-left">TRIMBOX - BENTO</td>
                                        <td class="text-left">SEM REGISTRO</td>
                                    </tr>
                                    <tr>
                                        <td class="text-left">TRIMBOX - MARANGUAPE</td>
                                        <td class="text-left">SEM REGISTRO</td>
                                    </tr>
                                    <tr>
                                        <td class="text-left">TRIMBOX - PIRAPETINGA</td>
                                        <td class="text-left">SEM REGISTRO</td>
                                    </tr>
                                    <tr>
                                        <td class="text-left">TRIMBOX - SAQUAREMA</td>
                                        <td class="text-left">SEM REGISTRO</td>
                                    </tr>
                                    <tr>
                                        <td class="text-left">TRIMBOX - UBERABA</td>
                                        <td class="text-left">SEM REGISTRO</td>
                                    </tr>
                                    <tr>
                                        <td class="text-left">TRIMPAPER - BENTO</td>
                                        <td class="text-left">SEM REGISTRO</td>
                                    </tr>
                                    <tr>
                                        <td class="text-left">TRIMPAPER - PIRAPETINGA</td>
                                        <td class="text-left">SEM REGISTRO</td>
                                    </tr>
                                    <tr>
                                        <td class="text-left">TRIMPAPER - UBERABA</td>
                                        <td class="text-left">SEM REGISTRO</td>
                                    </tr>
                                    <tr> 
                                        <td class="text-left">PROTHEUS - BENTO</td>
                                        <td class="text-left">SEM REGISTRO</td>
                                    </tr>
                                    <tr>
                                        <td class="text-left">PROTHEUS - MARANGUAPE</td>
                                        <td class="text-left">SEM REGISTRO</td>
                                    </tr>
                                    <tr>   
                                        <td class="text-left">PROTHEUS - PIRAPETINGA</td>
                                        <td class="text-left">SEM REGISTRO</td>
                                    </tr>
                                    <tr>
                                        <td class="text-left">PROTHEUS - SAQUAREMA</td>
                                        <td class="text-left">SEM REGISTRO</td>
                                    </tr>
                                <tbody>
                            </table>
                        </div>
                    </div>
                <% END IF %>

                <% IF etapa = "exibir" AND emailsSelecionados.Count > 0 THEN %>
                    <div class="row mb-3 align-middle">
                        <div class="col-md-12 mb-3">
                            <button type="button" class="btn-confirmar btn btn-sm btn-danger me-2 hidden" name="Confirmar" class="btn btn-primary" onclick="AlterarUsuarios('desabilitar')" title="Confirmar Cancelamento">
                                <i class="fa fa-ban me-2"></i>
                                <span>Desabilitar selecionados</span>
                            </button>
                            <% IF SESSION("AtivaUsuario") THEN %>
                            <button type="button" class="btn-confirmar btn btn-sm btn-success hidden" name="Confirmar" class="btn btn-primary" onclick="AlterarUsuarios('habilitar')" title="Confirmar Habilitação">
                                <i class="fa fa-check me-2"></i>
                                <span>Habilitar selecionados</span>
                            </button>
                            <% END IF %>
                        </div>
                        
                        <div class="col-md-12">
                            <table class="table table-bordered" name="Grade" id="Grade">
                                <thead>
                                    <tr>
                                        <th class="text-left">
                                            <input type="checkbox" id="selectAll" onchange="SelecionarTodos(this)">
                                        </th>
                                        <th class="text-left">#</th>
                                        <th class="text-left">Usuario</th>
                                        <th class="text-left">E-mail</th>
                                        <th class="text-left">Sistema</th>
                                        <th class="text-left">Data de Cadastro</th>
                                        <th class="text-left">Data de Bloqueio</th>
                                        <th class="text-left">Status</th>
                                    </tr>              
                                </thead>
                                <tbody>
                                   <%
                                    '----------------------------------------------
                                    '          USUARIOS HELPDESK - SWNET 
                                    '----------------------------------------------

                                    SQL = "Select CRM_NOME, CRM_EMAILPRIMARIO, CRM_R_E_C_N_O, CRM_D_E_L_E_T, CRM_DATCAD, DataBloqueio, "
                                    SQL = SQL & "CASE CRM_D_E_L_E_T WHEN '*' THEN 'INATIVO' ELSE 'ATIVO' END as Status "
                                    SQL = SQL & "FROM CRM_DTI WHERE CRM_EMAILPRIMARIO IN ('" & Join(emailsSelecionados.Keys, "', '") & "')"

                                    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)

                                    if not isObject(PIR_HELPDESK) OR PIR_HELPDESK.EOF then
                                    %>
                                        <tr>
                                            <td class="text-left"></td>
                                            <td class="text-left">#</td>
                                            <td class="text-left"><%=nomeusuario%></td>
                                            <td class="text-left"><%=emailusuario%></td>
                                            <td class="text-left">SWNET</td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">SEM REGISTRO</td>
                                        </tr>
                                    <%
                                    else
                                        While NOT PIR_HELPDESK.EOF
                                    %>
                                            <tr>                                                
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=PIR_HELPDESK("CRM_R_E_C_N_O")%>" 
                                                        data-sistema="swnet"
                                                        data-email="<%=UCase(Trim(PIR_HELPDESK("CRM_EMAILPRIMARIO")))%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                                
                                                <td class="text-left"><%=PIR_HELPDESK("CRM_R_E_C_N_O")%></td>
                                                <td class="text-left"><%=PIR_HELPDESK("CRM_NOME")%></td>
                                                <td class="text-left"><%=PIR_HELPDESK("CRM_EMAILPRIMARIO")%></td>
                                                <td class="text-left">SWNET</td>
                                                <td class="text-left"><%=acertadata(PIR_HELPDESK("CRM_DATCAD"))%></td>
                                                <td class="text-left"><%=acertadata(PIR_HELPDESK("DataBloqueio"))%></td>
                                            
                                                <td class="text-left">
                                                    <% if PIR_HELPDESK("Status") = "ATIVO" then %>
                                                        <div class="badge active"><%=PIR_HELPDESK("Status")%></div>
                                                    <% else %>
                                                        <div class="badge inactive"><%=PIR_HELPDESK("Status")%></div>
                                                    <% end if %>
                                                </td>
                                            </tr>
                                    <%
                                            PIR_HELPDESK.MoveNext()
                                        Wend
                                    end if

                                    if isObject(PIR_HELPDESK) then
                                        PIR_HELPDESK.Close()
                                        SET PIR_HELPDESK = NOTHING
                                    end if
                                    %>
                                    <tbody>

                                    <%
                                    '----------------------------------------------
                                    '         USUARIOS TRIMBOX - BENTO 
                                    '----------------------------------------------

                                    SQL = "Select NomeCompleto, Email, Usuario, convert(varchar,DataCriacao,103) as DataCriacao, convert(varchar,DataDesativado,103) as DataDesativado, CASE Desativado WHEN '1' THEN 'INATIVO' ELSE 'ATIVO' END as Status "
                                    SQL = SQL & " FROM Usuario "
                                    SQL = SQL & " WHERE 1=1 "
                                    SQL = SQL & " AND Email IN ('" & Join(emailsSelecionados.Keys, "', '") & "')"
                                    SQL = SQL & " ORDER BY NomeCompleto"

                                    SET BEN_TRIMBOX = Conecta( "BEN-TRIMBOX", SQL )

                                    if not isObject(BEN_TRIMBOX) OR BEN_TRIMBOX.EOF then
                                    %>
                                        <tr>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">TRIMBOX BENTO</td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">SEM REGISTRO</td>
                                        </tr>
                                    <%
                                    else
                                        While NOT BEN_TRIMBOX.EOF
                                    %>
                                            <tr>                                                
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=BEN_TRIMBOX("Usuario")%>" 
                                                        data-sistema="benTrimbox"
                                                        data-email="<%=UCase(Trim(BEN_TRIMBOX("Email")))%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                                
                                                <td class="text-left"><%=BEN_TRIMBOX("Usuario")%></td>
                                                <td class="text-left"><%=BEN_TRIMBOX("NomeCompleto")%></td>
                                                <td class="text-left"><%=BEN_TRIMBOX("Email")%></td>
                                                <td class="text-left">TRIMBOX BENTO</td>
                                                <td class="text-left"><%=BEN_TRIMBOX("DataCriacao")%></td>
                                                <td class="text-left"><%=BEN_TRIMBOX("DataDesativado")%></td>
                                                <td class="text-left">
                                                    <%if BEN_TRIMBOX("Status") = "ATIVO" then %>
                                                        <div class="badge active">
                                                    <%else%>
                                                        <div class="badge inactive">
                                                    <%end if%>
                                                        <%=BEN_TRIMBOX("Status")%>
                                                    </div>
                                                </td>

                                            </tr>
                                    <%
                                            BEN_TRIMBOX.MoveNext()
                                        Wend
                                    end if

                                    if isObject(BEN_TRIMBOX) then
                                        BEN_TRIMBOX.Close()
                                        SET BEN_TRIMBOX = NOTHING
                                    end if
                                    %>

                                    <%
                                    '----------------------------------------------
                                    '         USUARIOS TRIMBOX - PIRAPETINGA 
                                    '----------------------------------------------

                                    SQL = "Select NomeCompleto, Email, Usuario, convert(varchar,DataCriacao,103) as DataCriacao, convert(varchar,DataDesativado,103) as DataDesativado, CASE Desativado WHEN '1' THEN 'INATIVO' ELSE 'ATIVO' END as Status "
                                    SQL = SQL & " FROM Usuario "
                                    SQL = SQL & " WHERE 1=1 "
                                    SQL = SQL & " AND Email IN ('" & Join(emailsSelecionados.Keys, "', '") & "')"
                                    SQL = SQL & " ORDER BY NomeCompleto"

                                    SET PIR_TRIMBOX = Conecta( "PIR-TRIMBOX", SQL )

                                    if not isObject(PIR_TRIMBOX) OR PIR_TRIMBOX.EOF then
                                    %>
                                        <tr>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">TRIMBOX PIRAPETINGA</td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">SEM REGISTRO</td>
                                        </tr>
                                    <%
                                    else
                                        While NOT PIR_TRIMBOX.EOF
                                    %>
                                            <tr>                                                
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=PIR_TRIMBOX("Usuario")%>" 
                                                        data-sistema="pirTrimbox"
                                                        data-email="<%=UCase(Trim(PIR_TRIMBOX("Email")))%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                                
                                                <td class="text-left"><%=PIR_TRIMBOX("Usuario")%></td>
                                                <td class="text-left"><%=PIR_TRIMBOX("NomeCompleto")%></td>
                                                <td class="text-left"><%=PIR_TRIMBOX("Email")%></td>
                                                <td class="text-left">TRIMBOX PIRAPETINGA</td>
                                                <td class="text-left"><%=PIR_TRIMBOX("DataCriacao")%></td>
                                                <td class="text-left"><%=PIR_TRIMBOX("DataDesativado")%></td>
                                                <td class="text-left">
                                                    <%if PIR_TRIMBOX("Status") = "ATIVO" then %>
                                                        <div class="badge active">
                                                    <%else%>
                                                        <div class="badge inactive">
                                                    <%end if%>
                                                        <%=PIR_TRIMBOX("Status")%>
                                                    </div>
                                                </td>

                                            </tr>
                                    <%
                                            PIR_TRIMBOX.MoveNext()
                                        Wend
                                    end if

                                    if isObject(PIR_TRIMBOX) then
                                        PIR_TRIMBOX.Close()
                                        SET PIR_TRIMBOX = NOTHING
                                    end if
                                    %>

                                    <%
                                    '----------------------------------------------
                                    '         USUARIOS TRIMBOX - UBERABA 
                                    '----------------------------------------------

                                    SQL = "Select NomeCompleto, Email, Usuario, convert(varchar,DataCriacao,103) as DataCriacao, convert(varchar,DataDesativado,103) as DataDesativado, CASE Desativado WHEN '1' THEN 'INATIVO' ELSE 'ATIVO' END as Status "
                                    SQL = SQL & " FROM Usuario "
                                    SQL = SQL & " WHERE 1=1 "
                                    SQL = SQL & " AND Email IN ('" & Join(emailsSelecionados.Keys, "', '") & "')"
                                    SQL = SQL & " ORDER BY NomeCompleto"

                                    SET UBE_TRIMBOX = Conecta( "UBE-TRIMBOX", SQL )

                                    if not isObject(UBE_TRIMBOX) OR UBE_TRIMBOX.EOF then
                                    %>
                                        <tr>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">TRIMBOX UBERABA</td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">SEM REGISTRO</td>
                                        </tr>
                                    <%
                                    else
                                        While NOT UBE_TRIMBOX.EOF
                                    %>
                                            <tr>                                                
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=UBE_TRIMBOX("Usuario")%>" 
                                                        data-sistema="ubeTrimbox"
                                                        data-email="<%=UCase(Trim(UBE_TRIMBOX("Email")))%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                                
                                                <td class="text-left"><%=UBE_TRIMBOX("Usuario")%></td>
                                                <td class="text-left"><%=UBE_TRIMBOX("NomeCompleto")%></td>
                                                <td class="text-left"><%=UBE_TRIMBOX("Email")%></td>
                                                <td class="text-left">TRIMBOX UBERABA</td>
                                                <td class="text-left"><%=UBE_TRIMBOX("DataCriacao")%></td>
                                                <td class="text-left"><%=UBE_TRIMBOX("DataDesativado")%></td>
                                                <td class="text-left">
                                                    <%if UBE_TRIMBOX("Status") = "ATIVO" then %>
                                                        <div class="badge active">
                                                    <%else%>
                                                        <div class="badge inactive">
                                                    <%end if%>
                                                        <%=UBE_TRIMBOX("Status")%>
                                                    </div>
                                                </td>

                                            </tr>
                                    <%
                                            UBE_TRIMBOX.MoveNext()
                                        Wend
                                    end if

                                    if isObject(UBE_TRIMBOX) then
                                        UBE_TRIMBOX.Close()
                                        SET UBE_TRIMBOX = NOTHING
                                    end if
                                    %>

                                    <%
                                    '----------------------------------------------
                                    '         USUARIOS TRIMBOX - MARANGUAPE 
                                    '----------------------------------------------

                                    SQL = "Select NomeCompleto, Email, Usuario, convert(varchar,DataCriacao,103) as DataCriacao, convert(varchar,DataDesativado,103) as DataDesativado, CASE Desativado WHEN '1' THEN 'INATIVO' ELSE 'ATIVO' END as Status "
                                    SQL = SQL & " FROM Usuario "
                                    SQL = SQL & " WHERE 1=1 "
                                    SQL = SQL & " AND Email IN ('" & Join(emailsSelecionados.Keys, "', '") & "')"
                                    SQL = SQL & " ORDER BY NomeCompleto"

                                    SET FOR_TRIMBOX = Conecta( "FOR-TRIMBOX", SQL )

                                    if not isObject(FOR_TRIMBOX) OR FOR_TRIMBOX.EOF then
                                    %>
                                        <tr>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">TRIMBOX MARANGUAPE</td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">SEM REGISTRO</td>
                                        </tr>
                                    <%
                                    else
                                        While NOT FOR_TRIMBOX.EOF
                                    %>
                                            <tr>                                                
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=FOR_TRIMBOX("Usuario")%>" 
                                                        data-sistema="forTrimbox"
                                                        data-email="<%=UCase(Trim(FOR_TRIMBOX("Email")))%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                                
                                                <td class="text-left"><%=FOR_TRIMBOX("Usuario")%></td>
                                                <td class="text-left"><%=FOR_TRIMBOX("NomeCompleto")%></td>
                                                <td class="text-left"><%=FOR_TRIMBOX("Email")%></td>
                                                <td class="text-left">TRIMBOX MARANGUAPE</td>
                                                <td class="text-left"><%=FOR_TRIMBOX("DataCriacao")%></td>
                                                <td class="text-left"><%=FOR_TRIMBOX("DataDesativado")%></td>
                                                <td class="text-left">
                                                    <%if FOR_TRIMBOX("Status") = "ATIVO" then %>
                                                        <div class="badge active">
                                                    <%else%>
                                                        <div class="badge inactive">
                                                    <%end if%>
                                                        <%=FOR_TRIMBOX("Status")%>
                                                    </div>
                                                </td>

                                            </tr>
                                    <%
                                            FOR_TRIMBOX.MoveNext()
                                        Wend
                                    end if

                                    if isObject(FOR_TRIMBOX) then
                                        FOR_TRIMBOX.Close()
                                        SET FOR_TRIMBOX = NOTHING
                                    end if
                                    %>

                                    <%
                                    '----------------------------------------------
                                    '         USUARIOS TRIMBOX - SAQUAREMA 
                                    '----------------------------------------------

                                    SQL = "Select NomeCompleto, Email, Usuario, convert(varchar,DataCriacao,103) as DataCriacao, convert(varchar,DataDesativado,103) as DataDesativado, CASE Desativado WHEN '1' THEN 'INATIVO' ELSE 'ATIVO' END as Status "
                                    SQL = SQL & " FROM Usuario "
                                    SQL = SQL & " WHERE 1=1 "
                                    SQL = SQL & " AND Email IN ('" & Join(emailsSelecionados.Keys, "', '") & "')"
                                    SQL = SQL & " ORDER BY NomeCompleto"

                                    SET SAQ_TRIMBOX = Conecta( "SAQ-TRIMBOX", SQL )

                                    if not isObject(SAQ_TRIMBOX) OR SAQ_TRIMBOX.EOF then
                                    %>
                                        <tr>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">TRIMBOX SAQUAREMA</td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">SEM REGISTRO</td>
                                        </tr>
                                    <%
                                    else
                                        While NOT SAQ_TRIMBOX.EOF
                                    %>
                                            <tr>                                                
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=SAQ_TRIMBOX("Usuario")%>" 
                                                        data-sistema="saqTrimbox"
                                                        data-email="<%=UCase(Trim(SAQ_TRIMBOX("Email")))%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                                
                                                <td class="text-left"><%=SAQ_TRIMBOX("Usuario")%></td>
                                                <td class="text-left"><%=SAQ_TRIMBOX("NomeCompleto")%></td>
                                                <td class="text-left"><%=SAQ_TRIMBOX("Email")%></td>
                                                <td class="text-left">TRIMBOX SAQUAREMA</td>
                                                <td class="text-left"><%=SAQ_TRIMBOX("DataCriacao")%></td>
                                                <td class="text-left"><%=SAQ_TRIMBOX("DataDesativado")%></td>
                                                <td class="text-left">
                                                    <%if SAQ_TRIMBOX("Status") = "ATIVO" then %>
                                                        <div class="badge active">
                                                    <%else%>
                                                        <div class="badge inactive">
                                                    <%end if%>
                                                        <%=SAQ_TRIMBOX("Status")%>
                                                    </div>
                                                </td>

                                            </tr>
                                    <%
                                            SAQ_TRIMBOX.MoveNext()
                                        Wend
                                    end if

                                    if isObject(SAQ_TRIMBOX) then
                                        SAQ_TRIMBOX.Close()
                                        SET SAQ_TRIMBOX = NOTHING
                                    end if
                                    %>

                                    <%
                                    '----------------------------------------------
                                    '         USUARIOS TRIMPAPER - BENTO 
                                    '----------------------------------------------

                                    SQL = "Select NomeCompleto, Email, IDUsuario, convert(varchar,DataCadastro,103) as DataCadastro, DataDesativado, CASE Desativado WHEN '1' THEN 'INATIVO' ELSE 'ATIVO' END as Status "
                                    SQL = SQL & " FROM Usuarios "
                                    SQL = SQL & " WHERE 1=1 "
                                    SQL = SQL & " AND Email IN ('" & Join(emailsSelecionados.Keys, "', '") & "')"
                                    SQL = SQL & " ORDER BY NomeCompleto"

                                    SET BEN_TRIMPAPER = Conecta( "BEN-TRIMPAPER", SQL )

                                    if not isObject(BEN_TRIMPAPER) OR BEN_TRIMPAPER.EOF then
                                    %>
                                        <tr>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                                <td class="text-left">TRIMPAPER BENTO</td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">SEM REGISTRO</td>
                                        </tr>
                                    <%
                                    else
                                        While NOT BEN_TRIMPAPER.EOF
                                    %>
                                            <tr>                                                
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=BEN_TRIMPAPER("IDUsuario")%>" 
                                                        data-sistema="benTrimpaper"
                                                        data-email="<%=UCase(Trim(BEN_TRIMPAPER("Email")))%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                                
                                                <td class="text-left"><%=BEN_TRIMPAPER("IDUsuario")%></td>
                                                <td class="text-left"><%=BEN_TRIMPAPER("NomeCompleto")%></td>
                                                <td class="text-left"><%=BEN_TRIMPAPER("Email")%></td>
                                                <td class="text-left">TRIMPAPER BENTO</td>
                                                <td class="text-left"><%=BEN_TRIMPAPER("DataCadastro")%></td>
                                                <td class="text-left"><%=BEN_TRIMPAPER("DataDesativado")%></td>
                                                <td class="text-left">
                                                    <%if BEN_TRIMPAPER("Status") = "ATIVO" then %>
                                                        <div class="badge active">
                                                    <%else%>
                                                        <div class="badge inactive">
                                                    <%end if%>
                                                        <%=BEN_TRIMPAPER("Status")%>
                                                    </div>
                                                </td>

                                            </tr>
                                    <%
                                            BEN_TRIMPAPER.MoveNext()
                                        Wend
                                    end if

                                    if isObject(BEN_TRIMPAPER) then
                                        BEN_TRIMPAPER.Close()
                                        SET BEN_TRIMPAPER = NOTHING
                                    end if
                                    %>

                                    <%
                                    '----------------------------------------------
                                    '         USUARIOS TRIMPAPER - PIRAPETINGA 
                                    '----------------------------------------------

                                    SQL = "Select NomeCompleto, Email, IDUsuario, convert(varchar,DataCadastro,103) as DataCadastro, DataDesativado, CASE Desativado WHEN '1' THEN 'INATIVO' ELSE 'ATIVO' END as Status "
                                    SQL = SQL & " FROM Usuarios "
                                    SQL = SQL & " WHERE 1=1 "
                                    SQL = SQL & " AND Email IN ('" & Join(emailsSelecionados.Keys, "', '") & "')"
                                    SQL = SQL & " ORDER BY NomeCompleto"

                                    SET PIR_TRIMPAPER = Conecta( "PIR-TRIMPAPER", SQL )

                                    if not isObject(PIR_TRIMPAPER) OR PIR_TRIMPAPER.EOF then
                                    %>
                                        <tr>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                                <td class="text-left">TRIMPAPER PIRAPETINGA</td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">SEM REGISTRO</td>
                                        </tr>
                                    <%
                                    else
                                        While NOT PIR_TRIMPAPER.EOF
                                    %>
                                            <tr>                                                
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=PIR_TRIMPAPER("IDUsuario")%>" 
                                                        data-sistema="pirTrimpaper"
                                                        data-email="<%=UCase(Trim(PIR_TRIMPAPER("Email")))%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                                
                                                <td class="text-left"><%=PIR_TRIMPAPER("IDUsuario")%></td>
                                                <td class="text-left"><%=PIR_TRIMPAPER("NomeCompleto")%></td>
                                                <td class="text-left"><%=PIR_TRIMPAPER("Email")%></td>
                                                <td class="text-left">TRIMPAPER PIRAPETINGA</td>
                                                <td class="text-left"><%=PIR_TRIMPAPER("DataCadastro")%></td>
                                                <td class="text-left"><%=PIR_TRIMPAPER("DataDesativado")%></td>
                                                <td class="text-left">
                                                    <%if PIR_TRIMPAPER("Status") = "ATIVO" then %>
                                                        <div class="badge active">
                                                    <%else%>
                                                        <div class="badge inactive">
                                                    <%end if%>
                                                        <%=PIR_TRIMPAPER("Status")%>
                                                    </div>
                                                </td>

                                            </tr>
                                    <%
                                            PIR_TRIMPAPER.MoveNext()
                                        Wend
                                    end if

                                    if isObject(PIR_TRIMPAPER) then
                                        PIR_TRIMPAPER.Close()
                                        SET PIR_TRIMPAPER = NOTHING
                                    end if
                                    %>

                                    <%
                                    '----------------------------------------------
                                    '         USUARIOS TRIMPAPER - UBERABA 
                                    '----------------------------------------------

                                    SQL = "Select NomeCompleto, Email, IDUsuario, convert(varchar,DataCadastro,103) as DataCadastro, DataDesativado, CASE Desativado WHEN '1' THEN 'INATIVO' ELSE 'ATIVO' END as Status "
                                    SQL = SQL & " FROM Usuarios "
                                    SQL = SQL & " WHERE 1=1 "
                                    SQL = SQL & " AND Email IN ('" & Join(emailsSelecionados.Keys, "', '") & "')"
                                    SQL = SQL & " ORDER BY NomeCompleto"

                                    SET UBE_TRIMPAPER = Conecta( "UBE-TRIMPAPER", SQL )

                                    if not isObject(UBE_TRIMPAPER) OR UBE_TRIMPAPER.EOF then
                                    %>
                                        <tr>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                                <td class="text-left">TRIMPAPER UBERABA</td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">SEM REGISTRO</td>
                                        </tr>
                                    <%
                                    else
                                        While NOT UBE_TRIMPAPER.EOF
                                    %>
                                            <tr>                                                
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=UBE_TRIMPAPER("IDUsuario")%>" 
                                                        data-sistema="ubeTrimpaper"
                                                        data-email="<%=UCase(Trim(UBE_TRIMPAPER("Email")))%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                                
                                                <td class="text-left"><%=UBE_TRIMPAPER("IDUsuario")%></td>
                                                <td class="text-left"><%=UBE_TRIMPAPER("NomeCompleto")%></td>
                                                <td class="text-left"><%=UBE_TRIMPAPER("Email")%></td>
                                                <td class="text-left">TRIMPAPER UBERABA</td>
                                                <td class="text-left"><%=UBE_TRIMPAPER("DataCadastro")%></td>
                                                <td class="text-left"><%=UBE_TRIMPAPER("DataDesativado")%></td>
                                                <td class="text-left">
                                                    <%if UBE_TRIMPAPER("Status") = "ATIVO" then %>
                                                        <div class="badge active">
                                                    <%else%>
                                                        <div class="badge inactive">
                                                    <%end if%>
                                                        <%=UBE_TRIMPAPER("Status")%>
                                                    </div>
                                                </td>

                                            </tr>
                                    <%
                                            UBE_TRIMPAPER.MoveNext()
                                        Wend
                                    end if

                                    if isObject(UBE_TRIMPAPER) then
                                        UBE_TRIMPAPER.Close()
                                        SET UBE_TRIMPAPER = NOTHING
                                    end if
                                    %>

                                    <%
                                    '----------------------------------------------f
                                    '         USUARIOS PROTHEUS - PIRAPETINGA 
                                    '----------------------------------------------

                                    SQL = "Select USR_NOME as Nome, USR_EMAIL as Email, USR_ID as Id, convert(varchar,USR_DTINC,103) as DataCriacao, convert(varchar,USR_MSBLQD,103) as DataDesativado, CASE USR_MSBLQL WHEN '1' THEN 'INATIVO' ELSE 'ATIVO' END as Status "
                                    SQL = SQL & " FROM SYS_USR "
                                    SQL = SQL & " WHERE 1=1 "
                                    SQL = SQL & " AND USR_EMAIL IN ('" & Join(emailsSelecionados.Keys, "', '") & "')"
                                    SQL = SQL & " AND D_E_L_E_T_ = ''"
                                    SQL = SQL & " ORDER BY USR_NOME"

                                    SET PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

                                    if not isObject(PIR_DADOSADV) OR PIR_DADOSADV.EOF then
                                    %>
                                        <tr>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                                <td class="text-left">PROTHEUS PIRAPETINGA</td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">SEM REGISTRO</td>
                                        </tr>
                                    <%
                                    else
                                        While NOT PIR_DADOSADV.EOF
                                    %>
                                            <tr>                                                
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=PIR_DADOSADV("Id")%>" 
                                                        data-sistema="pirDadosadv" 
                                                        data-email="<%=UCase(Trim(PIR_DADOSADV("Email")))%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                                
                                                <td class="text-left"><%=PIR_DADOSADV("Id")%></td>
                                                <td class="text-left"><%=PIR_DADOSADV("Nome")%></td>
                                                <td class="text-left"><%=PIR_DADOSADV("Email")%></td>
                                                <td class="text-left">PROTHEUS PIRAPETINGA</td>
                                                <td class="text-left"><%=acertadata(PIR_DADOSADV("DataCriacao"))%></td>
                                                <td class="text-left"><%=acertadata(PIR_DADOSADV("DataDesativado"))%></td>
                                                <td class="text-left">
                                                    <%if PIR_DADOSADV("Status") = "ATIVO" then %>
                                                        <div class="badge active">
                                                    <%else%>
                                                        <div class="badge inactive">
                                                    <%end if%>
                                                        <%=PIR_DADOSADV("Status")%>
                                                    </div>
                                                </td>

                                            </tr>
                                    <%
                                            PIR_DADOSADV.MoveNext()
                                        Wend
                                    end if

                                    if isObject(PIR_DADOSADV) then
                                        PIR_DADOSADV.Close()
                                        SET PIR_DADOSADV = NOTHING
                                    end if
                                    %>

                                    <%
                                    '----------------------------------------------
                                    '         USUARIOS PROTHEUS - SAQUAREMA 
                                    '----------------------------------------------

                                    SQL = "Select USR_NOME as Nome, USR_EMAIL as Email, USR_ID as Id, convert(varchar,USR_DTINC,103) as DataCriacao, convert(varchar,USR_MSBLQD,103) as DataDesativado, CASE USR_MSBLQL WHEN '1' THEN 'INATIVO' ELSE 'ATIVO' END as Status "
                                    SQL = SQL & " FROM SYS_USR "
                                    SQL = SQL & " WHERE 1=1 "
                                    SQL = SQL & " AND USR_EMAIL IN ('" & Join(emailsSelecionados.Keys, "', '") & "')"
                                    SQL = SQL & " AND D_E_L_E_T_ = ''"
                                    SQL = SQL & " ORDER BY USR_NOME"

                                    SET SAQ_DADOSADV = Conecta( "SAQ-DADOSADV", SQL )

                                    if not isObject(SAQ_DADOSADV) OR SAQ_DADOSADV.EOF then
                                    %>
                                        <tr>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                                <td class="text-left">PROTHEUS SAQUAREMA</td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">SEM REGISTRO</td>
                                        </tr>
                                    <%
                                    else
                                        While NOT SAQ_DADOSADV.EOF
                                    %>
                                            <tr>                                                
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=SAQ_DADOSADV("Id")%>" 
                                                        data-sistema="saqDadosadv" 
                                                        data-email="<%=UCase(Trim(SAQ_DADOSADV("Email")))%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                                
                                                <td class="text-left"><%=SAQ_DADOSADV("Id")%></td>
                                                <td class="text-left"><%=SAQ_DADOSADV("Nome")%></td>
                                                <td class="text-left"><%=SAQ_DADOSADV("Email")%></td>
                                                <td class="text-left">PROTHEUS SAQUAREMA</td>
                                                <td class="text-left"><%=acertadata(SAQ_DADOSADV("DataCriacao"))%></td>
                                                <td class="text-left"><%=acertadata(SAQ_DADOSADV("DataDesativado"))%></td>
                                                <td class="text-left">
                                                    <%if SAQ_DADOSADV("Status") = "ATIVO" then %>
                                                        <div class="badge active">
                                                    <%else%>
                                                        <div class="badge inactive">
                                                    <%end if%>
                                                        <%=SAQ_DADOSADV("Status")%>
                                                    </div>
                                                </td>

                                            </tr>
                                    <%
                                            SAQ_DADOSADV.MoveNext()
                                        Wend
                                    end if

                                    if isObject(SAQ_DADOSADV) then
                                        SAQ_DADOSADV.Close()
                                        SET SAQ_DADOSADV = NOTHING
                                    end if
                                    %>

                                    <%
                                    '----------------------------------------------
                                    '         USUARIOS PROTHEUS - MARANGUAPE 
                                    '----------------------------------------------

                                    SQL = "Select USR_NOME as Nome, USR_EMAIL as Email, USR_ID as Id, convert(varchar,USR_DTINC,103) as DataCriacao, convert(varchar,USR_MSBLQD,103) as DataDesativado, CASE USR_MSBLQL WHEN '1' THEN 'INATIVO' ELSE 'ATIVO' END as Status "
                                    SQL = SQL & " FROM SYS_USR "
                                    SQL = SQL & " WHERE 1=1 "
                                    SQL = SQL & " AND USR_EMAIL IN ('" & Join(emailsSelecionados.Keys, "', '") & "')"
                                    SQL = SQL & " AND D_E_L_E_T_ = ''"
                                    SQL = SQL & " ORDER BY USR_NOME"

                                    SET FOR_DADOSADV = Conecta( "FOR-DADOSADV", SQL )

                                    if not isObject(FOR_DADOSADV) OR FOR_DADOSADV.EOF then
                                    %>
                                        <tr>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">PROTHEUS MARANGUAPE</td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">SEM REGISTRO</td>
                                        </tr>
                                    <%
                                    else
                                        While NOT FOR_DADOSADV.EOF
                                    %>
                                            <tr>                                                
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=FOR_DADOSADV("Id")%>" 
                                                        data-sistema="forDadosadv" 
                                                        data-email="<%=UCase(Trim(FOR_DADOSADV("Email")))%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                                
                                                <td class="text-left"><%=FOR_DADOSADV("Id")%></td>
                                                <td class="text-left"><%=FOR_DADOSADV("Nome")%></td>
                                                <td class="text-left"><%=FOR_DADOSADV("Email")%></td>
                                                <td class="text-left">PROTHEUS MARANGUAPE</td>
                                                <td class="text-left"><%=acertadata(FOR_DADOSADV("DataCriacao"))%></td>
                                                <td class="text-left"><%=acertadata(FOR_DADOSADV("DataDesativado"))%></td>
                                                <td class="text-left">
                                                    <%if FOR_DADOSADV("Status") = "ATIVO" then %>
                                                        <div class="badge active">
                                                    <%else%>
                                                        <div class="badge inactive">
                                                    <%end if%>
                                                        <%=FOR_DADOSADV("Status")%>
                                                    </div>
                                                </td>

                                            </tr>
                                    <%
                                            FOR_DADOSADV.MoveNext()
                                        Wend
                                    end if

                                    if isObject(FOR_DADOSADV) then
                                        FOR_DADOSADV.Close()
                                        SET FOR_DADOSADV = NOTHING
                                    end if
                                    %>

                                    <%
                                    '----------------------------------------------
                                    '         USUARIOS PROTHEUS - BENTO 
                                    '----------------------------------------------

                                    SQL = "Select USR_NOME as Nome, USR_EMAIL as Email, USR_ID as Id, convert(varchar,USR_DTINC,103) as DataCriacao, convert(varchar,USR_MSBLQD,103) as DataDesativado, CASE USR_MSBLQL WHEN '1' THEN 'INATIVO' ELSE 'ATIVO' END as Status "
                                    SQL = SQL & " FROM SYS_USR "
                                    SQL = SQL & " WHERE 1=1 "
                                    SQL = SQL & " AND USR_EMAIL IN ('" & Join(emailsSelecionados.Keys, "', '") & "')"
                                    SQL = SQL & " AND D_E_L_E_T_ = ''"
                                    SQL = SQL & " ORDER BY USR_NOME"

                                    SET BEN_DADOSADV = Conecta( "BEN-DADOSADV", SQL )

                                    if not isObject(BEN_DADOSADV) OR BEN_DADOSADV.EOF then
                                    %>
                                        <tr>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">PROTHEUS BENTO</td>
                                            <td class="text-left"></td>
                                            <td class="text-left"></td>
                                            <td class="text-left">SEM REGISTRO</td>
                                        </tr>
                                    <%
                                    else
                                        While NOT BEN_DADOSADV.EOF
                                    %>
                                            <tr>                                                
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=BEN_DADOSADV("Id")%>" 
                                                        data-sistema="benDadosadv" 
                                                        data-email="<%=UCase(Trim(BEN_DADOSADV("Email")))%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                                
                                                <td class="text-left"><%=BEN_DADOSADV("Id")%></td>
                                                <td class="text-left"><%=BEN_DADOSADV("Nome")%></td>
                                                <td class="text-left"><%=BEN_DADOSADV("Email")%></td>
                                                <td class="text-left">PROTHEUS BENTO</td>
                                                <td class="text-left"><%=acertadata(BEN_DADOSADV("DataCriacao"))%></td>
                                                <td class="text-left"><%=acertadata(BEN_DADOSADV("DataDesativado"))%></td>
                                                <td class="text-left">
                                                    <%if BEN_DADOSADV("Status") = "ATIVO" then %>
                                                        <div class="badge active">
                                                    <%else%>
                                                        <div class="badge inactive">
                                                    <%end if%>
                                                        <%=BEN_DADOSADV("Status")%>
                                                    </div>
                                                </td>

                                            </tr>
                                    <%
                                            BEN_DADOSADV.MoveNext()
                                        Wend
                                    end if

                                    if isObject(BEN_DADOSADV) then
                                        BEN_DADOSADV.Close()
                                        SET BEN_DADOSADV = NOTHING
                                    end if
                                    %>
                                </tbody>
                            </table>
                        </div>
                    </div>
                <% END IF %>
            </div>
        </form>
    </body>
</html>
