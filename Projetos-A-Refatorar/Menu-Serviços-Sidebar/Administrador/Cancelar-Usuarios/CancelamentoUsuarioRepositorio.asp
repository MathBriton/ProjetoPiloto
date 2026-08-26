<!--#include file="../includes/variaveis.asp" -->
<!--#include file="../funcoes/funcoes.asp" -->
<!--#include file="../cnn2/cnn2.asp" -->

<%session.LCID = 1046  
IdUsuarioLogado = Session("CODIGO")
LoginUsuarioLogado = Session("USR")

acao = Request("Acao")
emailsSelecionados = Request("emailsSelecionados")

swnetIds = Request("swnet")
benTrimboxIds = Request("benTrimbox")
pirTrimboxIds = Request("pirTrimbox")
forTrimboxIds = Request("forTrimbox")
ubeTrimboxIds = Request("ubeTrimbox")
saqTrimboxIds = Request("saqTrimbox")
benTrimpaperIds = Request("benTrimpaper")
pirTrimpaperIds = Request("pirTrimpaper")
ubeTrimpaperIds = Request("ubeTrimpaper")
pirDadosadvIds = Request("pirDadosadv")
saqDadosadvIds = Request("saqDadosadv")
forDadosadvIds = Request("forDadosadv")
benDadosadvIds = Request("benDadosadv")

'==================== VARIÁVEIS PARA IDS FILTRADOS ======================='

Dim swnetIdsAtualizacao 
Dim benTrimboxAtualizacao
Dim pirTrimboxIdsAtualizacao
Dim forTrimboxIdsAtualizacao
Dim ubeTrimboxIdsAtualizacao
Dim saqTrimboxIdsAtualizacao
Dim benTrimpaperIdsAtualizacao
Dim pirTrimpaperIdsAtualizacao
Dim ubeTrimpaperIdsAtualizacao
Dim pirDadosadvIdsAtualizacao
Dim saqDadosadvIdsAtualizacao
Dim forDadosadvIdsAtualizacao
Dim benDadosadvIdsAtualizacao

swnetIdsAtualizacao = ""
benTrimboxAtualizacao = ""
pirTrimboxIdsAtualizacao = ""
forTrimboxIdsAtualizacao = ""
ubeTrimboxIdsAtualizacao = ""
saqTrimboxIdsAtualizacao = ""
benTrimpaperIdsAtualizacao = ""
pirTrimpaperIdsAtualizacao = ""
ubeTrimpaperIdsAtualizacao = ""
pirDadosadvIdsAtualizacao = ""
saqDadosadvIdsAtualizacao = ""
forDadosadvIdsAtualizacao = ""
benDadosadvIdsAtualizacao = ""


'==================== FUNÇÃO PARA MONTAR SQL DE INSERIR LOG ======================='

Dim logValues
logValues = ""

Function AdicionarLog(idAlvo, sistema, nomeAlvo)
    Dim descricao, textoAcao, acaoFormatada
    if acao = "habilitar" then
        textoAcao = "habilitou"
    else
        textoAcao = "desabilitou"
    end if

    acaoFormatada = UCase(Left(acao, 1)) & Mid(acao, 2)
    descricao = "Usuário " & LoginUsuarioLogado & " " & textoAcao & " o usuário " & nomeAlvo
    descricao = Replace(descricao, "'", "''")
    
    if logValues <> "" Then
        logValues = logValues & ", "
    end if
    logValues = logValues & "(" & IdUsuarioLogado & ", " & idAlvo & ", 'Desligamento de usuário', '" & sistema & "', '" & acaoFormatada & " usuário', '" & descricao & "', GETDATE())"
End Function

Dim usuarios()
ReDim usuarios(0)
Dim idx
idx = 0

'==================== BUSCA USUÁRIOS QUE SERÃO ATUALIZADOS E ADICIONA A LISTA DE LOGS ======================='

'----------------------------------------------
'          USUARIOS HELPDESK - SWNET 
'----------------------------------------------

if swnetIds <> "" AND swnetIds <> "undefined" Then
    if(acao = "habilitar") then
        SQL = "Select CRM_R_E_C_N_O as Id, CRM_NOME as Nome, CRM_EMAILPRIMARIO as Email, 'SWNET' as Sistema"
        SQL = SQL & " FROM CRM_DTI WHERE CRM_R_E_C_N_O IN (" & swnetIds & ") AND DataBloqueio IS NOT NULL"
    else
        SQL = "Select CRM_R_E_C_N_O as Id, CRM_NOME as Nome, CRM_EMAILPRIMARIO as Email, 'SWNET' as Sistema"
        SQL = SQL & " FROM CRM_DTI WHERE CRM_R_E_C_N_O IN (" & swnetIds & ") AND DataBloqueio IS NULL"
    end if
    
    On Error Resume Next
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    
    While NOT PIR_HELPDESK.EOF
        AdicionarLog PIR_HELPDESK("Id"), "SWNET", PIR_HELPDESK("Nome")
        
        ' Adiciona o ID à lista de IDs para atualizar
        if swnetIdsAtualizacao <> "" Then
            swnetIdsAtualizacao = swnetIdsAtualizacao & ","
        end if
        swnetIdsAtualizacao = swnetIdsAtualizacao & PIR_HELPDESK("Id")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = PIR_HELPDESK("Id")
        usuarios(idx)("Nome") = PIR_HELPDESK("Nome")
        usuarios(idx)("Email") = PIR_HELPDESK("Email")
        usuarios(idx)("Sistema") = PIR_HELPDESK("Sistema")
        idx = idx + 1
        
        PIR_HELPDESK.MoveNext()
    Wend
    
    if isObject(PIR_HELPDESK) Then
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    end if
    On Error Goto 0
end if


'----------------------------------------------
'          USUARIOS TRIMBOX - BENTO 
'----------------------------------------------

if benTrimboxIds <> "" AND benTrimboxIds <> "undefined" Then
    if(acao = "habilitar") then
        SQL = "SELECT Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - BENTO' as Sistema" 
        SQL = SQL & " FROM Usuario WHERE Usuario IN (" & benTrimboxIds & ") AND Desativado = '1'"
    else
        SQL = "SELECT Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - BENTO' as Sistema" 
        SQL = SQL & " FROM Usuario WHERE Usuario IN (" & benTrimboxIds & ") AND Desativado = '0'"
    end if

    On Error Resume Next
    SET BEN_TRIMBOX = Conecta("BEN-TRIMBOX", SQL)
    
    While NOT BEN_TRIMBOX.EOF
        AdicionarLog BEN_TRIMBOX("Id"), "TRIMBOX - BENTO", BEN_TRIMBOX("Nome")
        
        if benTrimboxAtualizacao <> "" Then
            benTrimboxAtualizacao = benTrimboxAtualizacao & ","
        end if
        benTrimboxAtualizacao = benTrimboxAtualizacao & BEN_TRIMBOX("Id")
        
        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = BEN_TRIMBOX("Id")
        usuarios(idx)("Nome") = BEN_TRIMBOX("Nome")
        usuarios(idx)("Email") = BEN_TRIMBOX("Email")
        usuarios(idx)("Sistema") = BEN_TRIMBOX("Sistema")
        idx = idx + 1
        
        BEN_TRIMBOX.MoveNext()
    Wend
    
    if isObject(BEN_TRIMBOX) Then
        BEN_TRIMBOX.Close()
        SET BEN_TRIMBOX = NOTHING
    end if  
    On Error Goto 0
end if

'----------------------------------------------
'          USUARIOS TRIMBOX - PIRAPETINGA 
'----------------------------------------------

if pirTrimboxIds <> "" AND pirTrimboxIds <> "undefined" Then
    if(acao = "habilitar") then
        SQL = "SELECT Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - PIRAPETINGA' as Sistema" 
        SQL = SQL & " FROM Usuario WHERE Usuario IN (" & pirTrimboxIds & ") AND Desativado = '1'"
    else
        SQL = "SELECT Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - PIRAPETINGA' as Sistema" 
        SQL = SQL & " FROM Usuario WHERE Usuario IN (" & pirTrimboxIds & ") AND Desativado = '0'"
    end if

    On Error Resume Next
    SET PIR_TRIMBOX = Conecta("PIR-TRIMBOX", SQL)

    While NOT PIR_TRIMBOX.EOF
        AdicionarLog PIR_TRIMBOX("Id"), "TRIMBOX - PIRAPETINGA", PIR_TRIMBOX("Nome")
        
        if pirTrimboxIdsAtualizacao <> "" Then
            pirTrimboxIdsAtualizacao = pirTrimboxIdsAtualizacao & ","
        end if
        pirTrimboxIdsAtualizacao = pirTrimboxIdsAtualizacao & PIR_TRIMBOX("Id")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = PIR_TRIMBOX("Id")
        usuarios(idx)("Nome") = PIR_TRIMBOX("Nome")
        usuarios(idx)("Email") = PIR_TRIMBOX("Email")
        usuarios(idx)("Sistema") = PIR_TRIMBOX("Sistema")
        idx = idx + 1
        
        PIR_TRIMBOX.MoveNext()
    Wend

    if isObject(PIR_TRIMBOX) Then
        PIR_TRIMBOX.Close()
        SET PIR_TRIMBOX = NOTHING
    end if
    On Error Goto 0
end if

'----------------------------------------------
'          USUARIOS TRIMBOX - FORTALEZA 
'----------------------------------------------

if forTrimboxIds <> "" AND forTrimboxIds <> "undefined" Then
    if(acao = "habilitar") then
        SQL = "SELECT Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - FORTALEZA' as Sistema" 
        SQL = SQL & " FROM Usuario WHERE Usuario IN (" & forTrimboxIds & ") AND Desativado = '1'"
    else
        SQL = "SELECT Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - FORTALEZA' as Sistema" 
        SQL = SQL & " FROM Usuario WHERE Usuario IN (" & forTrimboxIds & ") AND Desativado = '0'"
    end if

    On Error Resume Next
    SET FOR_TRIMBOX = Conecta("FOR-TRIMBOX", SQL)

    While NOT FOR_TRIMBOX.EOF
        AdicionarLog FOR_TRIMBOX("Id"), "TRIMBOX - FORTALEZA", FOR_TRIMBOX("Nome")
        
        if forTrimboxIdsAtualizacao <> "" Then
            forTrimboxIdsAtualizacao = forTrimboxIdsAtualizacao & ","
        end if
        forTrimboxIdsAtualizacao = forTrimboxIdsAtualizacao & FOR_TRIMBOX("Id")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = FOR_TRIMBOX("Id")
        usuarios(idx)("Nome") = FOR_TRIMBOX("Nome")
        usuarios(idx)("Email") = FOR_TRIMBOX("Email")
        usuarios(idx)("Sistema") = FOR_TRIMBOX("Sistema")
        idx = idx + 1
        
        FOR_TRIMBOX.MoveNext()
    Wend

    if isObject(FOR_TRIMBOX) Then
        FOR_TRIMBOX.Close()
        SET FOR_TRIMBOX = NOTHING
    end if
    On Error Goto 0
end if

'----------------------------------------------
'          USUARIOS TRIMBOX - UBERABA
'----------------------------------------------

if ubeTrimboxIds <> "" AND ubeTrimboxIds <> "undefined" Then
    if(acao = "habilitar") then
        SQL = "SELECT Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - UBERABA' as Sistema" 
        SQL = SQL & " FROM Usuario WHERE Usuario IN (" & ubeTrimboxIds & ") AND Desativado = '1'"
    else
        SQL = "SELECT Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - UBERABA' as Sistema" 
        SQL = SQL & " FROM Usuario WHERE Usuario IN (" & ubeTrimboxIds & ") AND Desativado = '0'"
    end if

    On Error Resume Next
    SET UBE_TRIMBOX = Conecta("UBE-TRIMBOX", SQL)

    While NOT UBE_TRIMBOX.EOF
        AdicionarLog UBE_TRIMBOX("Id"), "TRIMBOX - UBERABA", UBE_TRIMBOX("Nome")
        
        if ubeTrimboxIdsAtualizacao <> "" Then
            ubeTrimboxIdsAtualizacao = ubeTrimboxIdsAtualizacao & ","
        end if
        ubeTrimboxIdsAtualizacao = ubeTrimboxIdsAtualizacao & UBE_TRIMBOX("Id")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = UBE_TRIMBOX("Id")
        usuarios(idx)("Nome") = UBE_TRIMBOX("Nome")
        usuarios(idx)("Email") = UBE_TRIMBOX("Email")
        usuarios(idx)("Sistema") = UBE_TRIMBOX("Sistema")
        idx = idx + 1
        
        UBE_TRIMBOX.MoveNext()
    Wend

    if isObject(UBE_TRIMBOX) Then
        UBE_TRIMBOX.Close()
        SET UBE_TRIMBOX = NOTHING
    end if
    On Error Goto 0
end if

'----------------------------------------------
'          USUARIOS TRIMBOX - SAQUAREMA
'----------------------------------------------

if saqTrimboxIds <> "" AND saqTrimboxIds <> "undefined" Then
    if(acao = "habilitar") then
        SQL = "SELECT Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - SAQUAREMA' as Sistema" 
        SQL = SQL & " FROM Usuario WHERE Usuario IN (" & saqTrimboxIds & ") AND Desativado = '1'"
    else
        SQL = "SELECT Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - SAQUAREMA' as Sistema" 
        SQL = SQL & " FROM Usuario WHERE Usuario IN (" & saqTrimboxIds & ") AND Desativado = '0'"
    end if

    On Error Resume Next
    SET SAQ_TRIMBOX = Conecta("SAQ-TRIMBOX", SQL)
    While NOT SAQ_TRIMBOX.EOF
        AdicionarLog SAQ_TRIMBOX("Id"), "TRIMBOX - SAQUAREMA", SAQ_TRIMBOX("Nome")
        
        if saqTrimboxIdsAtualizacao <> "" Then
            saqTrimboxIdsAtualizacao = saqTrimboxIdsAtualizacao & ","
        end if
        saqTrimboxIdsAtualizacao = saqTrimboxIdsAtualizacao & SAQ_TRIMBOX("Id")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = SAQ_TRIMBOX("Id")
        usuarios(idx)("Nome") = SAQ_TRIMBOX("Nome")
        usuarios(idx)("Email") = SAQ_TRIMBOX("Email")
        usuarios(idx)("Sistema") = SAQ_TRIMBOX("Sistema")
        idx = idx + 1
        
        SAQ_TRIMBOX.MoveNext()
    Wend
    if isObject(SAQ_TRIMBOX) Then
        SAQ_TRIMBOX.Close()
        SET SAQ_TRIMBOX = NOTHING
    end if
    On Error Goto 0
end if

'----------------------------------------------
'          USUARIOS TRIMPAPER - BENTO
'----------------------------------------------

if benTrimpaperIds <> "" AND benTrimpaperIds <> "undefined" Then
    if(acao = "habilitar") then
        SQL = "SELECT IDUsuario as Id, NomeCompleto as Nome, Email, 'TRIMPAPER - BENTO' as Sistema" 
        SQL = SQL & " FROM Usuarios WHERE IDUsuario IN (" & benTrimpaperIds & ") AND Desativado = '1'"
    else
        SQL = "SELECT IDUsuario as Id, NomeCompleto as Nome, Email, 'TRIMPAPER - BENTO' as Sistema" 
        SQL = SQL & " FROM Usuarios WHERE IDUsuario IN (" & benTrimpaperIds & ") AND Desativado = '0'"
    end if

    On Error Resume Next
    SET BEN_TRIMPAPER = Conecta("BEN-TRIMPAPER", SQL)

    While NOT BEN_TRIMPAPER.EOF
        AdicionarLog BEN_TRIMPAPER("Id"), "TRIMPAPER - BENTO", BEN_TRIMPAPER("Nome")
        
        if benTrimpaperIdsAtualizacao <> "" Then
            benTrimpaperIdsAtualizacao = benTrimpaperIdsAtualizacao & ","
        end if
        benTrimpaperIdsAtualizacao = benTrimpaperIdsAtualizacao & BEN_TRIMPAPER("Id")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = BEN_TRIMPAPER("Id")
        usuarios(idx)("Nome") = BEN_TRIMPAPER("Nome")
        usuarios(idx)("Email") = BEN_TRIMPAPER("Email")
        usuarios(idx)("Sistema") = BEN_TRIMPAPER("Sistema")
        idx = idx + 1
        
        BEN_TRIMPAPER.MoveNext()
    Wend

    if isObject(BEN_TRIMPAPER) Then
        BEN_TRIMPAPER.Close()
        SET BEN_TRIMPAPER = NOTHING
    end if
    On Error Goto 0
end if

'----------------------------------------------
'          USUARIOS TRIMPAPER - PIRAPETINGA
'----------------------------------------------

if pirTrimpaperIds <> "" AND pirTrimpaperIds <> "undefined" Then
    if(acao = "habilitar") then
        SQL = "SELECT IDUsuario as Id, NomeCompleto as Nome, Email, 'TRIMPAPER - PIRAPETINGA' as Sistema" 
        SQL = SQL & " FROM Usuarios WHERE IDUsuario IN (" & pirTrimpaperIds & ") AND Desativado = '1'"
    else
        SQL = "SELECT IDUsuario as Id, NomeCompleto as Nome, Email, 'TRIMPAPER - PIRAPETINGA' as Sistema" 
        SQL = SQL & " FROM Usuarios WHERE IDUsuario IN (" & pirTrimpaperIds & ") AND Desativado = '0'"
    end if

    On Error Resume Next
    SET PIR_TRIMPAPER = Conecta("PIR-TRIMPAPER", SQL)

    While NOT PIR_TRIMPAPER.EOF
        AdicionarLog PIR_TRIMPAPER("Id"), "TRIMPAPER - PIRAPETINGA", PIR_TRIMPAPER("Nome")
        
        if pirTrimpaperIdsAtualizacao <> "" Then
            pirTrimpaperIdsAtualizacao = pirTrimpaperIdsAtualizacao & ","
        end if
        pirTrimpaperIdsAtualizacao = pirTrimpaperIdsAtualizacao & PIR_TRIMPAPER("Id")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = PIR_TRIMPAPER("Id")
        usuarios(idx)("Nome") = PIR_TRIMPAPER("Nome")
        usuarios(idx)("Email") = PIR_TRIMPAPER("Email")
        usuarios(idx)("Sistema") = PIR_TRIMPAPER("Sistema")
        idx = idx + 1
        
        PIR_TRIMPAPER.MoveNext()
    Wend

    if isObject(PIR_TRIMPAPER) Then
        PIR_TRIMPAPER.Close()
        SET PIR_TRIMPAPER = NOTHING
    end if
    On Error Goto 0
end if

'----------------------------------------------
'          USUARIOS TRIMPAPER - UBERABA
'----------------------------------------------

if ubeTrimpaperIds <> "" AND ubeTrimpaperIds <> "undefined" Then
     if(acao = "habilitar") then
        SQL = "SELECT IDUsuario as Id, NomeCompleto as Nome, Email, 'TRIMPAPER - UBERABA' as Sistema" 
        SQL = SQL & " FROM Usuarios WHERE IDUsuario IN (" & ubeTrimpaperIds & ") AND Desativado = '1'"
    else
        SQL = "SELECT IDUsuario as Id, NomeCompleto as Nome, Email, 'TRIMPAPER - UBERABA' as Sistema" 
        SQL = SQL & " FROM Usuarios WHERE IDUsuario IN (" & ubeTrimpaperIds & ") AND Desativado = '0'"
    end if

    On Error Resume Next
    SET UBE_TRIMPAPER = Conecta("UBE-TRIMPAPER", SQL)

    While NOT UBE_TRIMPAPER.EOF
        AdicionarLog UBE_TRIMPAPER("Id"), "TRIMPAPER - UBERABA", UBE_TRIMPAPER("Nome")
        
        if ubeTrimpaperIdsAtualizacao <> "" Then
            ubeTrimpaperIdsAtualizacao = ubeTrimpaperIdsAtualizacao & ","
        end if
        ubeTrimpaperIdsAtualizacao = ubeTrimpaperIdsAtualizacao & UBE_TRIMPAPER("Id")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = UBE_TRIMPAPER("Id")
        usuarios(idx)("Nome") = UBE_TRIMPAPER("Nome")
        usuarios(idx)("Email") = UBE_TRIMPAPER("Email")
        usuarios(idx)("Sistema") = UBE_TRIMPAPER("Sistema")
        idx = idx + 1
        
        UBE_TRIMPAPER.MoveNext()
    Wend

    if isObject(UBE_TRIMPAPER) Then
        UBE_TRIMPAPER.Close()
        SET UBE_TRIMPAPER = NOTHING
    end if
    On Error Goto 0
end if

'----------------------------------------------
'          USUARIOS PROTHEUS - PIRAPETINGA
'----------------------------------------------

if pirDadosadvIds <> "" AND pirDadosadvIds <> "undefined" Then
    if(acao = "habilitar") then
        SQL = "SELECT USR_ID as Id, USR_NOME as Nome, USR_EMAIL as Email, 'PROTHEUS - PIRAPETINGA' as Sistema" 
        SQL = SQL & " FROM SYS_USR WHERE USR_ID IN (" & pirDadosadvIds & ") AND USR_MSBLQL = '1'"
    else
        SQL = "SELECT USR_ID as Id, USR_NOME as Nome, USR_EMAIL as Email, 'PROTHEUS - PIRAPETINGA' as Sistema" 
        SQL = SQL & " FROM SYS_USR WHERE USR_ID IN (" & pirDadosadvIds & ") AND USR_MSBLQL = '2'"
    end if

    On Error Resume Next
    SET PIR_DADOSADV = Conecta("PIR-DADOSADV", SQL)

    While NOT PIR_DADOSADV.EOF
        AdicionarLog "'" & PIR_DADOSADV("Id") & "'", "PROTHEUS - PIRAPETINGA", PIR_DADOSADV("Nome")
        
        if pirDadosadvIdsAtualizacao <> "" Then
            pirDadosadvIdsAtualizacao = pirDadosadvIdsAtualizacao & ","
        end if
        pirDadosadvIdsAtualizacao = pirDadosadvIdsAtualizacao & "'" & PIR_DADOSADV("Id") & "'"

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = PIR_DADOSADV("Id")
        usuarios(idx)("Nome") = PIR_DADOSADV("Nome")
        usuarios(idx)("Email") = PIR_DADOSADV("Email")
        usuarios(idx)("Sistema") = PIR_DADOSADV("Sistema")
        idx = idx + 1
        
        PIR_DADOSADV.MoveNext()
    Wend
    
    if isObject(PIR_DADOSADV) Then
        PIR_DADOSADV.Close()
        SET PIR_DADOSADV = NOTHING
    end if
    On Error Goto 0
end if

'----------------------------------------------
'          USUARIOS PROTHEUS - SAQUAREMA
'----------------------------------------------

if saqDadosadvIds <> "" AND saqDadosadvIds <> "undefined" Then
    if(acao = "habilitar") then
        SQL = "SELECT USR_ID as Id, USR_NOME as Nome, USR_EMAIL as Email, 'PROTHEUS - SAQUAREMA' as Sistema" 
        SQL = SQL & " FROM SYS_USR WHERE USR_ID IN (" & saqDadosadvIds & ") AND USR_MSBLQL = '1'"
    else
        SQL = "SELECT USR_ID as Id, USR_NOME as Nome, USR_EMAIL as Email, 'PROTHEUS - SAQUAREMA' as Sistema" 
        SQL = SQL & " FROM SYS_USR WHERE USR_ID IN (" & saqDadosadvIds & ") AND USR_MSBLQL = '2'"
    end if

    On Error Resume Next
    SET SAQ_DADOSADV = Conecta("SAQ-DADOSADV", SQL)

    While NOT SAQ_DADOSADV.EOF
        AdicionarLog "'" & SAQ_DADOSADV("Id") & "'", "PROTHEUS - SAQUAREMA", SAQ_DADOSADV("Nome")
        
        if saqDadosadvIdsAtualizacao <> "" Then
            saqDadosadvIdsAtualizacao = saqDadosadvIdsAtualizacao & ","
        end if
        saqDadosadvIdsAtualizacao = saqDadosadvIdsAtualizacao & "'" & SAQ_DADOSADV("Id") & "'"

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = SAQ_DADOSADV("Id")
        usuarios(idx)("Nome") = SAQ_DADOSADV("Nome")
        usuarios(idx)("Email") = SAQ_DADOSADV("Email")
        usuarios(idx)("Sistema") = SAQ_DADOSADV("Sistema")
        idx = idx + 1
        
        SAQ_DADOSADV.MoveNext()
    Wend

    if isObject(SAQ_DADOSADV) Then
        SAQ_DADOSADV.Close()
        SET SAQ_DADOSADV = NOTHING
    end if
    On Error Goto 0
end if

'----------------------------------------------
'          USUARIOS PROTHEUS - FORTALEZA
'----------------------------------------------

if forDadosadvIds <> "" AND forDadosadvIds <> "undefined" Then
    if(acao = "habilitar") then
        SQL = "SELECT USR_ID as Id, USR_NOME as Nome, USR_EMAIL as Email, 'PROTHEUS - FORTALEZA' as Sistema" 
        SQL = SQL & " FROM SYS_USR WHERE USR_ID IN (" & forDadosadvIds & ") AND USR_MSBLQL = '1'"
    else
        SQL = "SELECT USR_ID as Id, USR_NOME as Nome, USR_EMAIL as Email, 'PROTHEUS - FORTALEZA' as Sistema" 
        SQL = SQL & " FROM SYS_USR WHERE USR_ID IN (" & forDadosadvIds & ") AND USR_MSBLQL = '2'"
    end if

    On Error Resume Next
    SET FOR_DADOSADV = Conecta("FOR-DADOSADV", SQL)
    
    While NOT FOR_DADOSADV.EOF
        AdicionarLog "'" & FOR_DADOSADV("Id") & "'", "PROTHEUS - FORTALEZA", FOR_DADOSADV("Nome")
        
        if forDadosadvIdsAtualizacao <> "" Then
            forDadosadvIdsAtualizacao = forDadosadvIdsAtualizacao & ","
        end if
        forDadosadvIdsAtualizacao = forDadosadvIdsAtualizacao & "'" & FOR_DADOSADV("Id") & "'"

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = FOR_DADOSADV("Id")
        usuarios(idx)("Nome") = FOR_DADOSADV("Nome")
        usuarios(idx)("Email") = FOR_DADOSADV("Email")
        usuarios(idx)("Sistema") = FOR_DADOSADV("Sistema")
        idx = idx + 1
        
        FOR_DADOSADV.MoveNext()
    Wend

    if isObject(FOR_DADOSADV) Then
        FOR_DADOSADV.Close()
        SET FOR_DADOSADV = NOTHING
    end if
    On Error Goto 0
end if


'----------------------------------------------
'          USUARIOS PROTHEUS - BENTO
'----------------------------------------------

if benDadosadvIds <> "" AND benDadosadvIds <> "undefined" Then
    if(acao = "habilitar") then
        SQL = "SELECT USR_ID as Id, USR_NOME as Nome, USR_EMAIL as Email, 'PROTHEUS - BENTO' as Sistema" 
        SQL = SQL & " FROM SYS_USR WHERE USR_ID IN (" & benDadosadvIds & ") AND USR_MSBLQL = '1'"
    else
        SQL = "SELECT USR_ID as Id, USR_NOME as Nome, USR_EMAIL as Email, 'PROTHEUS - BENTO' as Sistema" 
        SQL = SQL & " FROM SYS_USR WHERE USR_ID IN (" & benDadosadvIds & ") AND USR_MSBLQL = '2'"
    end if

    On Error Resume Next
    SET BEN_DADOSADV = Conecta("BEN-DADOSADV", SQL)

    While NOT BEN_DADOSADV.EOF
        AdicionarLog "'" & BEN_DADOSADV("Id") & "'", "PROTHEUS - BENTO", BEN_DADOSADV("Nome")
        
        if benDadosadvIdsAtualizacao <> "" Then
            benDadosadvIdsAtualizacao = benDadosadvIdsAtualizacao & ","
        end if
        benDadosadvIdsAtualizacao = benDadosadvIdsAtualizacao & "'" & BEN_DADOSADV("Id") & "'"

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = BEN_DADOSADV("Id")
        usuarios(idx)("Nome") = BEN_DADOSADV("Nome")
        usuarios(idx)("Email") = BEN_DADOSADV("Email")
        usuarios(idx)("Sistema") = BEN_DADOSADV("Sistema")
        idx = idx + 1
        
        BEN_DADOSADV.MoveNext()
    Wend

    if isObject(BEN_DADOSADV) Then
        BEN_DADOSADV.Close()
        SET BEN_DADOSADV = NOTHING
    end if
    On Error Goto 0
end if

'===================== REALIZA UPDATE COM IDS FILTRADOS (SOMENTE SE HOUVER USUÁRIOS) ====================='

if idx > 0 then
    '----------------------------------------------
    '               USUARIOS HELPDESK - SWNET 
    '----------------------------------------------

    if swnetIdsAtualizacao <> "" then     
        if acao = "habilitar" then
            SQL = "UPDATE CRM_DTI SET "
            SQL = SQL & " CRM_D_E_L_E_T = '', "
            SQL = SQL & " Observacoes = 'USUARIO HABILITADO EM ' + convert(varchar,getDate(),103), "
            SQL = SQL & " DataBloqueio = NULL "
            SQL = SQL & " WHERE CRM_R_E_C_N_O IN (" & swnetIdsAtualizacao & ")"
        else
            SQL = "UPDATE CRM_DTI SET "
            SQL = SQL & " CRM_D_E_L_E_T = '*', "
            SQL = SQL & " Observacoes = 'USUARIO DESABILITADO EM ' + convert(varchar,getDate(),103), "
            SQL = SQL & " DataBloqueio = convert(varchar,getDate(),112) "
            SQL = SQL & " WHERE CRM_R_E_C_N_O IN (" & swnetIdsAtualizacao & ")"
        end if
        
        On Error Resume Next
        SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
        
        if Err.Number = 0 AND isObject(PIR_HELPDESK) then
            PIR_HELPDESK.Close()
            SET PIR_HELPDESK = NOTHING
        end if
        
        On Error Goto 0
    end if 

    '----------------------------------------------
    '            USUARIOS TRIMBOX  - BENTO 
    '----------------------------------------------

    if benTrimboxAtualizacao <> "" then   
        if acao = "habilitar" then
            SQL = "UPDATE Usuario SET "
            SQL = SQL & " Desativado = '0', "
            SQL = SQL & " DataDesativado = NULL "
            SQL = SQL & " WHERE Usuario IN (" & benTrimboxAtualizacao & ")"
        else
            SQL = "UPDATE Usuario SET "
            SQL = SQL & " Desativado = '1', "
            SQL = SQL & " DataDesativado = getDate() "
            SQL = SQL & " WHERE Usuario IN (" & benTrimboxAtualizacao & ")"
        end if

        On Error Resume Next
        SET BEN_TRIMBOX = Conecta("BEN-TRIMBOX", SQL)
        
        if Err.Number = 0 AND isObject(BEN_TRIMBOX) then
            BEN_TRIMBOX.Close()
            SET BEN_TRIMBOX = NOTHING
        end if
        
        On Error Goto 0
    end if 

    '----------------------------------------------
    '           USUARIOS TRIMBOX  - PIRAPETINGA 
    '----------------------------------------------

    if pirTrimboxIdsAtualizacao <> "" then   
        if acao = "habilitar" then
            SQL = "UPDATE Usuario SET "
            SQL = SQL & " Desativado = '0', "
            SQL = SQL & " DataDesativado = NULL "
            SQL = SQL & " WHERE Usuario IN (" & pirTrimboxIdsAtualizacao & ")"
        else
            SQL = "UPDATE Usuario SET "
            SQL = SQL & " Desativado = '1', "
            SQL = SQL & " DataDesativado = getDate() "
            SQL = SQL & " WHERE Usuario IN (" & pirTrimboxIdsAtualizacao & ")"
        end if

        On Error Resume Next
        SET PIR_TRIMBOX = Conecta("PIR-TRIMBOX", SQL)
        
        if Err.Number = 0 AND isObject(PIR_TRIMBOX) then
            PIR_TRIMBOX.Close()
            SET PIR_TRIMBOX = NOTHING
        end if
        
        On Error Goto 0
    end if 

    '----------------------------------------------
    '           USUARIOS TRIMBOX - FORTALEZA 
    '----------------------------------------------

    if forTrimboxIdsAtualizacao <> "" then   
        if acao = "habilitar" then
            SQL = "UPDATE Usuario SET "
            SQL = SQL & " Desativado = '0', "
            SQL = SQL & " DataDesativado = NULL "
            SQL = SQL & " WHERE Usuario IN (" & forTrimboxIdsAtualizacao & ")"
        else
            SQL = "UPDATE Usuario SET "
            SQL = SQL & " Desativado = '1', "
            SQL = SQL & " DataDesativado = getDate() "
            SQL = SQL & " WHERE Usuario IN (" & forTrimboxIdsAtualizacao & ")"
        end if

        On Error Resume Next
        SET FOR_TRIMBOX = Conecta("FOR-TRIMBOX", SQL)
        
        if Err.Number = 0 AND isObject(FOR_TRIMBOX) then
            FOR_TRIMBOX.Close()
            SET FOR_TRIMBOX = NOTHING
        end if
        
        On Error Goto 0
    end if 

    '----------------------------------------------
    '           USUARIOS TRIMBOX - UBERABA 
    '----------------------------------------------

    if ubeTrimboxIdsAtualizacao <> "" then     
        if acao = "habilitar" then
            SQL = "UPDATE Usuario SET "
            SQL = SQL & " Desativado = '0', "
            SQL = SQL & " DataDesativado = NULL "
            SQL = SQL & " WHERE Usuario IN (" & ubeTrimboxIdsAtualizacao & ")"
        else
            SQL = "UPDATE Usuario SET "
            SQL = SQL & " Desativado = '1', "
            SQL = SQL & " DataDesativado = getDate() "
            SQL = SQL & " WHERE Usuario IN (" & ubeTrimboxIdsAtualizacao & ")"
        end if

        On Error Resume Next
        SET UBE_TRIMBOX = Conecta("UBE-TRIMBOX", SQL)
        
        if Err.Number = 0 AND isObject(UBE_TRIMBOX) then
            UBE_TRIMBOX.Close()
            SET UBE_TRIMBOX = NOTHING
        end if
        
        On Error Goto 0
    end if 

    '----------------------------------------------
    '           USUARIOS TRIMBOX - SAQUAREMA 
    '----------------------------------------------

    if saqTrimboxIdsAtualizacao <> "" then     
        if acao = "habilitar" then
            SQL = "UPDATE Usuario SET "
            SQL = SQL & " Desativado = '0', "
            SQL = SQL & " DataDesativado = NULL "
            SQL = SQL & " WHERE Usuario IN (" & saqTrimboxIdsAtualizacao & ")"
        else
            SQL = "UPDATE Usuario SET "
            SQL = SQL & " Desativado = '1', "
            SQL = SQL & " DataDesativado = getDate() "
            SQL = SQL & " WHERE Usuario IN (" & saqTrimboxIdsAtualizacao & ")"
        end if

        On Error Resume Next
        SET SAQ_TRIMBOX = Conecta("SAQ-TRIMBOX", SQL)
        
        if Err.Number = 0 AND isObject(SAQ_TRIMBOX) then
            SAQ_TRIMBOX.Close()
            SET SAQ_TRIMBOX = NOTHING
        end if
        
        On Error Goto 0
    end if 

    '----------------------------------------------
    '           USUARIOS TRIMPAPER  - BENTO 
    '----------------------------------------------

    if benTrimpaperIdsAtualizacao <> "" then     
        if acao = "habilitar" then
            SQL = "UPDATE Usuarios SET "
            SQL = SQL & " Desativado = '0', "
            SQL = SQL & " DataDesativado = NULL "
            SQL = SQL & " WHERE IDUsuario IN (" & benTrimpaperIdsAtualizacao & ")"
        else
            SQL = "UPDATE Usuarios SET "
            SQL = SQL & " Desativado = '1', "
            SQL = SQL & " DataDesativado = getDate() "
            SQL = SQL & " WHERE IDUsuario IN (" & benTrimpaperIdsAtualizacao & ")"
        end if

        On Error Resume Next
        SET BEN_TRIMPAPER = Conecta("BEN-TRIMPAPER", SQL)
        
        if Err.Number = 0 AND isObject(BEN_TRIMPAPER) then
            BEN_TRIMPAPER.Close()
            SET BEN_TRIMPAPER = NOTHING
        end if
        
        On Error Goto 0
    end if 

    '----------------------------------------------
    '           USUARIOS TRIMPAPEL  - PIRAPETINGA 
    '----------------------------------------------

    if pirTrimpaperIdsAtualizacao <> "" then    
        if acao = "habilitar" then
            SQL = "UPDATE Usuarios SET "
            SQL = SQL & " Desativado = '0', "
            SQL = SQL & " DataDesativado = NULL "
            SQL = SQL & " WHERE IDUsuario IN (" & pirTrimpaperIdsAtualizacao & ")"
        else
            SQL = "UPDATE Usuarios SET "
            SQL = SQL & " Desativado = '1', "
            SQL = SQL & " DataDesativado = getDate() "
            SQL = SQL & " WHERE IDUsuario IN (" & pirTrimpaperIdsAtualizacao & ")"
        end if

        On Error Resume Next
        SET PIR_TRIMPAPER = Conecta("PIR-TRIMPAPER", SQL)
        
        if Err.Number = 0 AND isObject(PIR_TRIMPAPER) then
            PIR_TRIMPAPER.Close()
            SET PIR_TRIMPAPER = NOTHING
        end if
        
        On Error Goto 0
    end if 

    '----------------------------------------------
    '           USUARIOS TRIMPAPER - UBERABA 
    '----------------------------------------------

    if ubeTrimpaperIdsAtualizacao <> "" then     
        if acao = "habilitar" then
            SQL = "UPDATE Usuarios SET "
            SQL = SQL & " Desativado = '0', "
            SQL = SQL & " DataDesativado = NULL "
            SQL = SQL & " WHERE IDUsuario IN (" & ubeTrimpaperIdsAtualizacao & ")"
        else
            SQL = "UPDATE Usuarios SET "
            SQL = SQL & " Desativado = '1', "
            SQL = SQL & " DataDesativado = getDate() "
            SQL = SQL & " WHERE IDUsuario IN (" & ubeTrimpaperIdsAtualizacao & ")"
        end if

        On Error Resume Next
        SET UBE_TRIMPAPER = Conecta("UBE-TRIMPAPER", SQL)
        
        if Err.Number = 0 AND isObject(UBE_TRIMPAPER) then
            UBE_TRIMPAPER.Close()
            SET UBE_TRIMPAPER = NOTHING
        end if
        
        On Error Goto 0
    end if 

    '----------------------------------------------
    '           USUARIOS PROTHEUS - PIRAPETINGA 
    '----------------------------------------------

    if pirDadosadvIdsAtualizacao <> "" then     
        if acao = "habilitar" then
            SQL = "UPDATE SYS_USR SET "
            SQL = SQL & " USR_MSBLQL = '2', "
            SQL = SQL & " USR_MSBLQD = '' "
            SQL = SQL & " WHERE USR_ID IN (" & pirDadosadvIdsAtualizacao & ")"
        else
            SQL = "UPDATE SYS_USR SET "
            SQL = SQL & " USR_MSBLQL = '1', "
            SQL = SQL & " USR_MSBLQD = convert(varchar,getDate(),112) "
            SQL = SQL & " WHERE USR_ID IN (" & pirDadosadvIdsAtualizacao & ")"
        end if

        On Error Resume Next
        SET PIR_DADOSADV = Conecta("PIR-DADOSADV", SQL)
        
        if Err.Number = 0 AND isObject(PIR_DADOSADV) then
            PIR_DADOSADV.Close()
            SET PIR_DADOSADV = NOTHING
        end if
        
        On Error Goto 0
    end if 

    '----------------------------------------------
    '           USUARIOS PROTHEUS - SAQUAREMA 
    '----------------------------------------------

    if saqDadosadvIdsAtualizacao <> "" then     
        if acao = "habilitar" then
            SQL = "UPDATE SYS_USR SET "
            SQL = SQL & " USR_MSBLQL = '2', "
            SQL = SQL & " USR_MSBLQD = '' "
            SQL = SQL & " WHERE USR_ID IN (" & saqDadosadvIdsAtualizacao & ")"
        else
            SQL = "UPDATE SYS_USR SET "
            SQL = SQL & " USR_MSBLQL = '1', "
            SQL = SQL & " USR_MSBLQD = convert(varchar,getDate(),112) "
            SQL = SQL & " WHERE USR_ID IN (" & saqDadosadvIdsAtualizacao & ")"
        end if

        On Error Resume Next
        SET SAQ_DADOSADV = Conecta("SAQ-DADOSADV", SQL)
        
        if Err.Number = 0 AND isObject(SAQ_DADOSADV) then
            SAQ_DADOSADV.Close()
            SET SAQ_DADOSADV = NOTHING
        end if
        
        On Error Goto 0
    end if 

    '----------------------------------------------
    '           USUARIOS PROTHEUS - FORTALEZA 
    '----------------------------------------------

    if forDadosadvIdsAtualizacao <> "" then     
        if acao = "habilitar" then
            SQL = "UPDATE SYS_USR SET "
            SQL = SQL & " USR_MSBLQL = '2', "
            SQL = SQL & " USR_MSBLQD = '' "
            SQL = SQL & " WHERE USR_ID IN (" & forDadosadvIdsAtualizacao & ")"
        else
            SQL = "UPDATE SYS_USR SET "
            SQL = SQL & " USR_MSBLQL = '1', "
            SQL = SQL & " USR_MSBLQD = convert(varchar,getDate(),112) "
            SQL = SQL & " WHERE USR_ID IN (" & forDadosadvIdsAtualizacao & ")"
        end if

        On Error Resume Next
        SET FOR_DADOSADV = Conecta("FOR-DADOSADV", SQL)
        
        if Err.Number = 0 AND isObject(FOR_DADOSADV) then
            FOR_DADOSADV.Close()
            SET FOR_DADOSADV = NOTHING
        end if
        
        On Error Goto 0
    end if 

    '----------------------------------------------
    '           USUARIOS PROTHEUS - BENTO 
    '----------------------------------------------

    if benDadosadvIdsAtualizacao <> "" then     
        if acao = "habilitar" then
            SQL = "UPDATE SYS_USR SET "
            SQL = SQL & " USR_MSBLQL = '2', "
            SQL = SQL & " USR_MSBLQD = '' "
            SQL = SQL & " WHERE USR_ID IN (" & benDadosadvIdsAtualizacao & ")"
        else
            SQL = "UPDATE SYS_USR SET "
            SQL = SQL & " USR_MSBLQL = '1', "
            SQL = SQL & " USR_MSBLQD = convert(varchar,getDate(),112) "
            SQL = SQL & " WHERE USR_ID IN (" & benDadosadvIdsAtualizacao & ")"
        end if

        On Error Resume Next
        SET BEN_DADOSADV = Conecta("BEN-DADOSADV", SQL)
        
        if Err.Number = 0 AND isObject(BEN_DADOSADV) then
            BEN_DADOSADV.Close()
            SET BEN_DADOSADV = NOTHING
        end if
        
        On Error Goto 0
    end if 

    '===================== APÓS FINALIZAR UPDATE, INSERE LOGS ====================='

    if logValues <> "" Then
        SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, IdUsuarioAlvo, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES " & logValues
        
        On Error Resume Next
        SET LOG_CONN = Conecta("PIR-CORPORATIVO", SQL)
        
        if Err.Number = 0 AND isObject(LOG_CONN) Then
            LOG_CONN.Close()
            SET LOG_CONN = NOTHING
        end if
        On Error Goto 0
    end if

end if ' Fecha o if idx > 0

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
            function Voltar() {
                var emailsSelecionados = "<%=emailsSelecionados%>";
                
                if (emailsSelecionados && emailsSelecionados !== "") {
                    window.location.href = "CancelamentoUsuario.asp?email=" + emailsSelecionados;
                } else {
                    window.location.href = "CancelamentoUsuario.asp";
                }
            }
        </script>
        <style type="text/css">
            body{
                font-size:70% !important;
            }
            .btn {
                font-size: .8rem;
                width: fit-content;
            }
            h3 {
                margin-left: 0;
            }
        </style>
    </head>
    <body>
        <div class="container mt-4">
            <%if idx = 0 then%>
                <div class="alert alert-warning" role="alert">
                    <h4 class="alert-heading">
                        <i class="fas fa-exclamation-triangle"></i> Atenção!
                    </h4>
                    <hr>
                    <p class="mb-0">
                        Nenhum usuário foi encontrado para 
                        <%if acao = "habilitar" then%>
                            habilitar
                        <%else%>
                            desabilitar
                        <%end if%>
                    </p>
                    <p class="mt-2 mb-0">
                        Os usuários selecionados já estão no estado desejado ou não foram localizados nos sistemas.
                    </p>
                </div>
                
                <div class="row mt-3">
                    <div class="col-md-12">
                        <button type="button" class="btn btn-primary" onclick="Voltar()">
                            <i class="fas fa-arrow-left"></i> Voltar
                        </button>
                    </div>
                </div>

            <%else%>
                <h3 class="alert-heading">
                    <%if acao = "habilitar" then%>
                        Usuários habilitados com sucesso!
                    <%else%>
                        Usuários desabilitados com sucesso!
                    <%end if%>
                </h3>

                <div class="row mb-3">
                    <div class="col-md-12">
                        <button type="button" class="btn btn-danger mr-3" onclick="exportarPDF()">
                            <i class="fas fa-file-pdf"></i> Exportar para PDF
                        </button>
                        <button type="button" class="btn btn-primary" onclick="Voltar()">
                            <i class="fas fa-arrow-left"></i> Voltar
                        </button>
                    </div>
                </div>

                <div class="row">
                    <div class="col-md-12">
                        <table class="table table-striped table-hover table-responsive-sm" name="Grade" id="Grade">
                            <thead style="background-color: #D3D3D4;">
                                <tr>
                                    <th class="text-left">#</th>
                                    <th class="text-left">Nome</th>
                                    <th class="text-left">Email</th>
                                    <th class="text-left">Sistema</th>
                                    <th class="text-left">
                                        <%if acao = "habilitar" then%>
                                            Data de desbloqueio
                                        <%else%>
                                            Data de bloqueio
                                        <%end if%>
                                    </th>
                                </tr>              
                            </thead>
                            <tbody>
                                <% 
                                Dim i
                                For i = 0 To UBound(usuarios)     
                                %>
                                    <tr>                                                
                                        <td class="text-left"><%=usuarios(i)("Id")%></td>
                                        <td class="text-left"><%=usuarios(i)("Nome")%></td>
                                        <td class="text-left"><%=usuarios(i)("Email")%></td>
                                        <td class="text-left"><%=usuarios(i)("Sistema")%></td>
                                        <td class="text-left"><%=FormatDatetime(Date(), 2)%></td>
                                    </tr>
                                <% Next %>
                            </tbody>
                        </table>
                    </div>
                </div>

                <script>
                    function exportarPDF() {
                        const { jsPDF } = window.jspdf;
                        const doc = new jsPDF('landscape');
                        const dataAtual = new Date().toLocaleDateString('pt-BR');
                        const acao =  "<%=acao%>";

                        const tabela = document.getElementById('Grade');
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
                        doc.text(`Relatório de Usuários ${acao == 'habilitar'? 'Habilitados': 'Desabilitados'}`, 14, 15);
                        doc.setFontSize(10);
                        doc.text('Data de geração: ' + dataAtual + ` | Total de usuários ${acao == 'habilitar' ? 'habilitados': 'desabilitados'}: ` + dados.length, 14, 22);
                        
                        doc.autoTable({
                            head: [['#', 'Nome', 'Email', 'Sistema', acao == 'habilitar' ? 'Data de desbloqueio' : 'Data de bloqueio']],
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
                </script>
            <%end if%>
        </div>
    </body>
</html>