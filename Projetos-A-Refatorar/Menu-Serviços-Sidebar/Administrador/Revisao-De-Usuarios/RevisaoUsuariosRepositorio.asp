<!--#include file="../includes/variaveis.asp" -->
<!--#include file="../funcoes/funcoes.asp" -->
<!--#include file="../cnn2/cnn2.asp" -->

<%session.LCID = 1046  

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

IdUsuarioLogado = Session("CODIGO")
LoginUsuarioLogado = Session("USR")

'==================== FUNÇÃO AUXILIAR PARA ADICIONAR A LISTA DE LOGS ======================='

Dim logValues
logValues = ""

Function AdicionarLog(idAlvo, sistema, nomeAlvo)
    Dim descricao
    descricao = "Usuario " & LoginUsuarioLogado & " desabilitou o usuario " & nomeAlvo
    descricao = Replace(descricao, "'", "''")
    
    if logValues <> "" Then
        logValues = logValues & ", "
    end if
    logValues = logValues & "(" & IdUsuarioLogado & ", " & idAlvo & ", 'Revisao de usuarios', '" & sistema & "', 'Desabilitar usuário', '" & descricao & "', GETDATE())"
End Function


'================== BUSCA USUÁRIOS A SEREM ATUALIZADOS E ADICIONA A LISTA DE LOG ====================='

Dim usuarios()
Dim totalUsuarios
totalUsuarios = 0

ReDim usuarios(0)
Dim idx
idx = 0

'----------------------------------------------
'          USUARIOS HELPDESK - SWNET 
'----------------------------------------------
if swnetIds <> "" AND swnetIds <> "undefined" then     
    SQL = "Select CRM_R_E_C_N_O as Id, CRM_NOME as Nome, CRM_EMAILPRIMARIO as Email, 'SWNET' as Sistema,"
    SQL = SQL & " CASE WHEN CRM_DATACES = '' THEN DATEDIFF(DAY,CRM_DATCAD, GETDATE()) ELSE DATEDIFF(DAY,CRM_DATACES, GETDATE()) END AS DiasSemAcessar"
    SQL = SQL & " FROM CRM_DTI WHERE CRM_R_E_C_N_O IN (" & swnetIds & ")"
    
    SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
    While NOT PIR_HELPDESK.EOF
        AdicionarLog PIR_HELPDESK("Id"), "SWNET", PIR_HELPDESK("Nome")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = PIR_HELPDESK("Id")
        usuarios(idx)("Nome") = PIR_HELPDESK("Nome")
        usuarios(idx)("Email") = PIR_HELPDESK("Email")
        usuarios(idx)("Sistema") = PIR_HELPDESK("Sistema")
        usuarios(idx)("DiasSemAcessar") = PIR_HELPDESK("DiasSemAcessar")
        idx = idx + 1
        PIR_HELPDESK.MoveNext()
    Wend

    if isObject(PIR_HELPDESK) Then
        PIR_HELPDESK.Close()
        SET PIR_HELPDESK = NOTHING
    end if
end if

'----------------------------------------------
'         USUARIOS TRIMBOX - PIRAPETINGA 
'----------------------------------------------
if pirTrimboxIds <> "" AND pirTrimboxIds <> "undefined" then     
    SQL = "Select Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - PIRAPETINGA' as Sistema,"
    SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCriacao, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
    SQL = SQL & " FROM Usuario WHERE Usuario in (" & pirTrimboxIds & ")"
    
    SET PIR_TRIMBOX = Conecta("PIR-TRIMBOX", SQL)
    While NOT PIR_TRIMBOX.EOF
        AdicionarLog PIR_TRIMBOX("Id"), "TRIMBOX - PIRAPETINGA", PIR_TRIMBOX("Nome")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = PIR_TRIMBOX("Id")
        usuarios(idx)("Nome") = PIR_TRIMBOX("Nome")
        usuarios(idx)("Email") = PIR_TRIMBOX("Email")
        usuarios(idx)("Sistema") = PIR_TRIMBOX("Sistema")
        usuarios(idx)("DiasSemAcessar") = PIR_TRIMBOX("DiasSemAcessar")
        idx = idx + 1
        PIR_TRIMBOX.MoveNext()
    Wend

    if isObject(PIR_TRIMBOX) Then
        PIR_TRIMBOX.Close()
        SET PIR_TRIMBOX = NOTHING
    end if
end if

'----------------------------------------------
'         USUARIOS TRIMBOX - BENTO 
'----------------------------------------------
if benTrimboxIds <> "" AND benTrimboxIds <> "undefined" then     
    SQL = "Select Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - BENTO' as Sistema,"
    SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCriacao, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
    SQL = SQL & " FROM Usuario WHERE Usuario in (" & benTrimboxIds & ")"
    
    SET BEN_TRIMBOX = Conecta("BEN-TRIMBOX", SQL)
    While NOT BEN_TRIMBOX.EOF
        AdicionarLog BEN_TRIMBOX("Id"), "TRIMBOX - BENTO", BEN_TRIMBOX("Nome")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = BEN_TRIMBOX("Id")
        usuarios(idx)("Nome") = BEN_TRIMBOX("Nome")
        usuarios(idx)("Email") = BEN_TRIMBOX("Email")
        usuarios(idx)("Sistema") = BEN_TRIMBOX("Sistema")
        usuarios(idx)("DiasSemAcessar") = BEN_TRIMBOX("DiasSemAcessar")
        idx = idx + 1
        BEN_TRIMBOX.MoveNext()
    Wend

    if isObject(BEN_TRIMBOX) Then
        BEN_TRIMBOX.Close()
        SET BEN_TRIMBOX = NOTHING
    end if
end if

'----------------------------------------------
'         USUARIOS TRIMBOX - UBERABA 
'----------------------------------------------
if ubeTrimboxIds <> "" AND ubeTrimboxIds <> "undefined" then     
    SQL = "Select Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - UBERABA' as Sistema,"
    SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCriacao, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
    SQL = SQL & " FROM Usuario WHERE Usuario in (" & ubeTrimboxIds & ")"
    
    SET UBE_TRIMBOX = Conecta("UBE-TRIMBOX", SQL)
    While NOT UBE_TRIMBOX.EOF
        AdicionarLog UBE_TRIMBOX("Id"), "TRIMBOX - UBERABA", UBE_TRIMBOX("Nome")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = UBE_TRIMBOX("Id")
        usuarios(idx)("Nome") = UBE_TRIMBOX("Nome")
        usuarios(idx)("Email") = UBE_TRIMBOX("Email")
        usuarios(idx)("Sistema") = UBE_TRIMBOX("Sistema")
        usuarios(idx)("DiasSemAcessar") = UBE_TRIMBOX("DiasSemAcessar")
        idx = idx + 1
        UBE_TRIMBOX.MoveNext()
    Wend

    if isObject(UBE_TRIMBOX) Then
        UBE_TRIMBOX.Close()
        SET UBE_TRIMBOX = NOTHING
    end if
end if

'----------------------------------------------
'         USUARIOS TRIMBOX - MARANGUAPE 
'----------------------------------------------
if forTrimboxIds <> "" AND forTrimboxIds <> "undefined" then     
    SQL = "Select Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - MARANGUAPE' as Sistema,"
    SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCriacao, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
    SQL = SQL & " FROM Usuario WHERE Usuario in (" & forTrimboxIds & ")"

    SET FOR_TRIMBOX = Conecta("FOR-TRIMBOX", SQL)
    While NOT FOR_TRIMBOX.EOF
        AdicionarLog FOR_TRIMBOX("Id"), "TRIMBOX - MARANGUAPE", FOR_TRIMBOX("Nome")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = FOR_TRIMBOX("Id")
        usuarios(idx)("Nome") = FOR_TRIMBOX("Nome")
        usuarios(idx)("Email") = FOR_TRIMBOX("Email")
        usuarios(idx)("Sistema") = FOR_TRIMBOX("Sistema")
        usuarios(idx)("DiasSemAcessar") = FOR_TRIMBOX("DiasSemAcessar")
        idx = idx + 1
        FOR_TRIMBOX.MoveNext()
    Wend

    if isObject(FOR_TRIMBOX) Then
        FOR_TRIMBOX.Close()
        SET FOR_TRIMBOX = NOTHING
    end if
end if

'----------------------------------------------
'         USUARIOS TRIMBOX - SAQUAREMA 
'----------------------------------------------
if saqTrimboxIds <> "" AND saqTrimboxIds <> "undefined" then     
    SQL = "Select Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - SAQUAREMA' as Sistema,"
    SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCriacao, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
    SQL = SQL & " FROM Usuario WHERE Usuario in (" & saqTrimboxIds & ")"
    
    SET SAQ_TRIMBOX = Conecta("SAQ-TRIMBOX", SQL)
    While NOT SAQ_TRIMBOX.EOF
        AdicionarLog SAQ_TRIMBOX("Id"), "TRIMBOX - SAQUAREMA", SAQ_TRIMBOX("Nome")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = SAQ_TRIMBOX("Id")
        usuarios(idx)("Nome") = SAQ_TRIMBOX("Nome")
        usuarios(idx)("Email") = SAQ_TRIMBOX("Email")
        usuarios(idx)("Sistema") = SAQ_TRIMBOX("Sistema")
        usuarios(idx)("DiasSemAcessar") = SAQ_TRIMBOX("DiasSemAcessar")
        idx = idx + 1
        SAQ_TRIMBOX.MoveNext()
    Wend

    if isObject(SAQ_TRIMBOX) Then
        SAQ_TRIMBOX.Close()
        SET SAQ_TRIMBOX = NOTHING
    end if
end if

'----------------------------------------------
'         USUARIOS TRIMPAPER - BENTO 
'----------------------------------------------
if benTrimpaperIds <> "" AND benTrimpaperIds <> "undefined" then     
    SQL = "Select IDUsuario as Id, NomeCompleto as Nome, Email, 'TRIMPAPER - BENTO' as Sistema,"
    SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCadastro, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
    SQL = SQL & " FROM Usuarios WHERE IDUsuario in (" & benTrimpaperIds & ")"
    
    SET BEN_TRIMPAPER = Conecta("BEN-TRIMPAPER", SQL)
    While NOT BEN_TRIMPAPER.EOF
        AdicionarLog BEN_TRIMPAPER("Id"), "TRIMPAPER - BENTO", BEN_TRIMPAPER("Nome")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = BEN_TRIMPAPER("Id")
        usuarios(idx)("Nome") = BEN_TRIMPAPER("Nome")
        usuarios(idx)("Email") = BEN_TRIMPAPER("Email")
        usuarios(idx)("Sistema") = BEN_TRIMPAPER("Sistema")
        usuarios(idx)("DiasSemAcessar") = BEN_TRIMPAPER("DiasSemAcessar")
        idx = idx + 1
        BEN_TRIMPAPER.MoveNext()
    Wend

    if isObject(BEN_TRIMPAPER) Then
        BEN_TRIMPAPER.Close()
        SET BEN_TRIMPAPER = NOTHING
    end if
end if

'----------------------------------------------
'         USUARIOS TRIMPAPER - PIRAPETINGA 
'----------------------------------------------
if pirTrimpaperIds <> "" AND pirTrimpaperIds <> "undefined" then     
    SQL = "Select IDUsuario as Id, NomeCompleto as Nome, Email, 'TRIMPAPER - PIRAPETINGA' as Sistema,"
    SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCadastro, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
    SQL = SQL & " FROM Usuarios WHERE IDUsuario in (" & pirTrimpaperIds & ")"
    
    SET PIR_TRIMPAPER = Conecta("PIR-TRIMPAPER", SQL)
    While NOT PIR_TRIMPAPER.EOF
        AdicionarLog PIR_TRIMPAPER("Id"), "TRIMPAPER - PIRAPETINGA", PIR_TRIMPAPER("Nome")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = PIR_TRIMPAPER("Id")
        usuarios(idx)("Nome") = PIR_TRIMPAPER("Nome")
        usuarios(idx)("Email") = PIR_TRIMPAPER("Email")
        usuarios(idx)("Sistema") = PIR_TRIMPAPER("Sistema")
        usuarios(idx)("DiasSemAcessar") = PIR_TRIMPAPER("DiasSemAcessar")
        idx = idx + 1
        PIR_TRIMPAPER.MoveNext()
    Wend

    if isObject(PIR_TRIMPAPER) Then
        PIR_TRIMPAPER.Close()
        SET PIR_TRIMPAPER = NOTHING
    end if
end if

'----------------------------------------------
'         USUARIOS TRIMPAPER - UBERABA 
'----------------------------------------------
if ubeTrimpaperIds <> "" AND ubeTrimpaperIds <> "undefined" then     
    SQL = "Select IDUsuario as Id, NomeCompleto as Nome, Email, 'TRIMPAPER - UBERABA' as Sistema,"
    SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCadastro, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
    SQL = SQL & " FROM Usuarios WHERE IDUsuario in (" & ubeTrimpaperIds & ")"
    
    SET UBE_TRIMPAPER = Conecta("UBE-TRIMPAPER", SQL)
    While NOT UBE_TRIMPAPER.EOF
        AdicionarLog UBE_TRIMPAPER("Id"), "TRIMPAPER - UBERABA", UBE_TRIMPAPER("Nome")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = UBE_TRIMPAPER("Id")
        usuarios(idx)("Nome") = UBE_TRIMPAPER("Nome")
        usuarios(idx)("Email") = UBE_TRIMPAPER("Email")
        usuarios(idx)("Sistema") = UBE_TRIMPAPER("Sistema")
        usuarios(idx)("DiasSemAcessar") = UBE_TRIMPAPER("DiasSemAcessar")
        idx = idx + 1
        UBE_TRIMPAPER.MoveNext()
    Wend

    if isObject(UBE_TRIMPAPER) Then
        UBE_TRIMPAPER.Close()
        SET UBE_TRIMPAPER = NOTHING
    end if
end if

'----------------------------------------------
'         USUARIOS PROTHEUS - PIRAPETINGA 
'----------------------------------------------
if pirDadosadvIds <> "" AND pirDadosadvIds <> "undefined" then     
    SQL = "Select u.USR_ID as Id, u.USR_NOME as Nome, u.USR_EMAIL as Email, 'PROTHEUS - PIRAPETINGA' as Sistema,"
    SQL = SQL & "CASE WHEN ISNULL(s.USR_DTLOGON,u.USR_DTLOGON) <> '' "
    SQL = SQL & "   THEN DATEDIFF(DAY,CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON), 103), CONVERT(DATE,GETDATE(), 103)) "
    SQL = SQL & "   ELSE DATEDIFF(DAY,CONVERT(DATE,u.USR_DTINC, 103), CONVERT(DATE,GETDATE(), 103) "
    SQL = SQL & ") END AS DiasSemAcessar "
    SQL = SQL & " FROM SYS_USR u"
    SQL = SQL & " LEFT JOIN SYS_USR_LOGIN s on s.USR_ID = u.USR_ID "
    SQL = SQL & " WHERE u.USR_ID in (" & pirDadosadvIds & ")"
    
    SET PIR_DADOSADV = Conecta("PIR-DADOSADV", SQL)
    While NOT PIR_DADOSADV.EOF
        AdicionarLog PIR_DADOSADV("Id"), "PROTHEUS - PIRAPETINGA", PIR_DADOSADV("Nome")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = PIR_DADOSADV("Id")
        usuarios(idx)("Nome") = PIR_DADOSADV("Nome")
        usuarios(idx)("Email") = PIR_DADOSADV("Email")
        usuarios(idx)("Sistema") = PIR_DADOSADV("Sistema")
        usuarios(idx)("DiasSemAcessar") = PIR_DADOSADV("DiasSemAcessar")
        idx = idx + 1
        PIR_DADOSADV.MoveNext()
    Wend

    if isObject(PIR_DADOSADV) Then
        PIR_DADOSADV.Close()
        SET PIR_DADOSADV = NOTHING
    end if
end if

'----------------------------------------------
'         USUARIOS PROTHEUS - SAQUAREMA 
'----------------------------------------------
if saqDadosadvIds <> "" AND saqDadosadvIds <> "undefined" then     
    SQL = "Select u.USR_ID as Id, u.USR_NOME as Nome, u.USR_EMAIL as Email, 'PROTHEUS - SAQUAREMA' as Sistema,"
    SQL = SQL & "CASE WHEN ISNULL(s.USR_DTLOGON,u.USR_DTLOGON) <> '' "
    SQL = SQL & "   THEN DATEDIFF(DAY,CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON), 103), CONVERT(DATE,GETDATE(), 103)) "
    SQL = SQL & "   ELSE DATEDIFF(DAY,CONVERT(DATE,u.USR_DTINC, 103), CONVERT(DATE,GETDATE(), 103) "
    SQL = SQL & ") END AS DiasSemAcessar "
    SQL = SQL & " FROM SYS_USR u"
    SQL = SQL & " LEFT JOIN SYS_USR_LOGIN s on s.USR_ID = u.USR_ID "
    SQL = SQL & " WHERE u.USR_ID in (" & saqDadosadvIds & ")"
    
    SET SAQ_DADOSADV = Conecta("SAQ-DADOSADV", SQL)
    While NOT SAQ_DADOSADV.EOF
        AdicionarLog SAQ_DADOSADV("Id"), "PROTHEUS - SAQUAREMA", SAQ_DADOSADV("Nome")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = SAQ_DADOSADV("Id")
        usuarios(idx)("Nome") = SAQ_DADOSADV("Nome")
        usuarios(idx)("Email") = SAQ_DADOSADV("Email")
        usuarios(idx)("Sistema") = SAQ_DADOSADV("Sistema")
        usuarios(idx)("DiasSemAcessar") = SAQ_DADOSADV("DiasSemAcessar")
        idx = idx + 1
        SAQ_DADOSADV.MoveNext()
    Wend

    if isObject(SAQ_DADOSADV) Then
        SAQ_DADOSADV.Close()
        SET SAQ_DADOSADV = NOTHING
    end if
end if

'----------------------------------------------
'         USUARIOS PROTHEUS - MARANGUAPE 
'----------------------------------------------
if forDadosadvIds <> "" AND forDadosadvIds <> "undefined" then     
    SQL = "Select u.USR_ID as Id, u.USR_NOME as Nome, u.USR_EMAIL as Email, 'PROTHEUS - MARANGUAPE' as Sistema,"
    SQL = SQL & "CASE WHEN ISNULL(s.USR_DTLOGON,u.USR_DTLOGON) <> '' "
    SQL = SQL & "   THEN DATEDIFF(DAY,CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON), 103), CONVERT(DATE,GETDATE(), 103)) "
    SQL = SQL & "   ELSE DATEDIFF(DAY,CONVERT(DATE,u.USR_DTINC, 103), CONVERT(DATE,GETDATE(), 103) "
    SQL = SQL & ") END AS DiasSemAcessar "
    SQL = SQL & " FROM SYS_USR u"
    SQL = SQL & " LEFT JOIN SYS_USR_LOGIN s on s.USR_ID = u.USR_ID "
    SQL = SQL & " WHERE u.USR_ID in (" & forDadosadvIds & ")"
    
    SET FOR_DADOSADV = Conecta("FOR-DADOSADV", SQL)
    While NOT FOR_DADOSADV.EOF
        AdicionarLog FOR_DADOSADV("Id"), "PROTHEUS - MARANGUAPE", FOR_DADOSADV("Nome")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = FOR_DADOSADV("Id")
        usuarios(idx)("Nome") = FOR_DADOSADV("Nome")
        usuarios(idx)("Email") = FOR_DADOSADV("Email")
        usuarios(idx)("Sistema") = FOR_DADOSADV("Sistema")
        usuarios(idx)("DiasSemAcessar") = FOR_DADOSADV("DiasSemAcessar")
        idx = idx + 1
        FOR_DADOSADV.MoveNext()
    Wend

    if isObject(FOR_DADOSADV) Then
        FOR_DADOSADV.Close()
        SET FOR_DADOSADV = NOTHING
    end if
end if

'----------------------------------------------
'         USUARIOS PROTHEUS - BENTO 
'----------------------------------------------
if benDadosadvIds <> "" AND benDadosadvIds <> "undefined" then     
    SQL = "Select u.USR_ID as Id, u.USR_NOME as Nome, u.USR_EMAIL as Email, 'PROTHEUS - BENTO' as Sistema,"
    SQL = SQL & "CASE WHEN ISNULL(s.USR_DTLOGON,u.USR_DTLOGON) <> '' "
    SQL = SQL & "   THEN DATEDIFF(DAY,CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON), 103), CONVERT(DATE,GETDATE(), 103)) "
    SQL = SQL & "   ELSE DATEDIFF(DAY,CONVERT(DATE,u.USR_DTINC, 103), CONVERT(DATE,GETDATE(), 103) "
    SQL = SQL & ") END AS DiasSemAcessar "
    SQL = SQL & " FROM SYS_USR u"
    SQL = SQL & " LEFT JOIN SYS_USR_LOGIN s on s.USR_ID = u.USR_ID "
    SQL = SQL & " WHERE u.USR_ID in (" & benDadosadvIds & ")"
    
    SET BEN_DADOSADV = Conecta("BEN-DADOSADV", SQL)
    While NOT BEN_DADOSADV.EOF
        AdicionarLog BEN_DADOSADV("Id"), "PROTHEUS - BENTO", BEN_DADOSADV("Nome")

        ReDim Preserve usuarios(idx)
        Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
        usuarios(idx)("Id") = BEN_DADOSADV("Id")
        usuarios(idx)("Nome") = BEN_DADOSADV("Nome")
        usuarios(idx)("Email") = BEN_DADOSADV("Email")
        usuarios(idx)("Sistema") = BEN_DADOSADV("Sistema")
        usuarios(idx)("DiasSemAcessar") = BEN_DADOSADV("DiasSemAcessar")
        idx = idx + 1
        BEN_DADOSADV.MoveNext()
    Wend

    if isObject(BEN_DADOSADV) Then
        BEN_DADOSADV.Close()
        SET BEN_DADOSADV = NOTHING
    end if
end if


'==================== REALIZA UPDATE ======================='

'----------------------------------------------
'               USUARIOS HELPDESK - SWNET 
'----------------------------------------------

if swnetIds <> "" AND swnetIds <> "undefined" then     
    SQL = "UPDATE CRM_DTI SET "
    SQL = SQL & " CRM_D_E_L_E_T = '*', "
    SQL = SQL & " Observacoes = 'USUARIO DESABILITADO EM ' + convert(varchar,getDate(),103), "
    SQL = SQL & " DataBloqueio = convert(varchar,getDate(),112) "
    SQL = SQL & " WHERE CRM_R_E_C_N_O IN (" & swnetIds & ")"
    SQL = SQL & " AND DataBloqueio IS NULL"
    
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

if benTrimboxIds <> "" AND benTrimboxIds <> "undefined" then   
    SQL = "UPDATE Usuario SET "
    SQL = SQL & " Desativado = '1', "
    SQL = SQL & " DataDesativado = getDate() "
    SQL = SQL & " WHERE Usuario IN (" & benTrimboxIds & ")"
    SQL = SQL & " AND Desativado = '0'"

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

if pirTrimboxIds <> "" AND pirTrimboxIds <> "undefined" then   
    SQL = "UPDATE Usuario SET "
    SQL = SQL & " Desativado = '1', "
    SQL = SQL & " DataDesativado = getDate() "
    SQL = SQL & " WHERE Usuario IN (" & pirTrimboxIds & ")"
    SQL = SQL & " AND Desativado = '0'"

    On Error Resume Next
    SET PIR_TRIMBOX = Conecta("PIR-TRIMBOX", SQL)
    
    if Err.Number = 0 AND isObject(PIR_TRIMBOX) then
        PIR_TRIMBOX.Close()
        SET PIR_TRIMBOX = NOTHING
    end if
    
    On Error Goto 0
end if 

'----------------------------------------------
'           USUARIOS TRIMBOX - MARANGUAPE 
'----------------------------------------------

if forTrimboxIds <> "" AND forTrimboxIds <> "undefined" then   
    SQL = "UPDATE Usuario SET "
    SQL = SQL & " Desativado = '1', "
    SQL = SQL & " DataDesativado = getDate() "
    SQL = SQL & " WHERE Usuario IN (" & forTrimboxIds & ")"
    SQL = SQL & " AND Desativado = '0'"

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

if ubeTrimboxIds <> "" AND ubeTrimboxIds <> "undefined" then     
    SQL = "UPDATE Usuario SET "
    SQL = SQL & " Desativado = '1', "
    SQL = SQL & " DataDesativado = getDate() "
    SQL = SQL & " WHERE Usuario IN (" & ubeTrimboxIds & ")"
    SQL = SQL & " AND Desativado = '0'"

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

if saqTrimboxIds <> "" AND saqTrimboxIds <> "undefined" then     
    SQL = "UPDATE Usuario SET "
    SQL = SQL & " Desativado = '1', "
    SQL = SQL & " DataDesativado = getDate() "
    SQL = SQL & " WHERE Usuario IN (" & saqTrimboxIds & ")"
    SQL = SQL & " AND Desativado = '0'"

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

if benTrimpaperIds <> "" AND benTrimpaperIds <> "undefined" then     
    SQL = "UPDATE Usuarios SET "
    SQL = SQL & " Desativado = '1', "
    SQL = SQL & " DataDesativado = getDate() "
    SQL = SQL & " WHERE IDUsuario IN (" & benTrimpaperIds & ")"
    SQL = SQL & " AND Desativado = '0'"

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

if pirTrimpaperIds <> "" AND pirTrimpaperIds <> "undefined" then    
    SQL = "UPDATE Usuarios SET "
    SQL = SQL & " Desativado = '1', "
    SQL = SQL & " DataDesativado = getDate() "
    SQL = SQL & " WHERE IDUsuario IN (" & pirTrimpaperIds & ")"
    SQL = SQL & " AND Desativado = '0'"

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

if ubeTrimpaperIds <> "" AND ubeTrimpaperIds <> "undefined" then     
    SQL = "UPDATE Usuarios SET "
    SQL = SQL & " Desativado = '1', "
    SQL = SQL & " DataDesativado = getDate() "
    SQL = SQL & " WHERE IDUsuario IN (" & ubeTrimpaperIds & ")"
    SQL = SQL & " AND Desativado = '0'"

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

if pirDadosadvIds <> "" AND pirDadosadvIds <> "undefined" then     
    SQL = "UPDATE SYS_USR SET "
    SQL = SQL & " USR_MSBLQL = '1', "
    SQL = SQL & " USR_MSBLQD = convert(varchar,getDate(),112) "
    SQL = SQL & " WHERE USR_ID IN (" & pirDadosadvIds & ")"
    SQL = SQL & " AND USR_MSBLQL = '2'"

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

if saqDadosadvIds <> "" AND saqDadosadvIds <> "undefined" then     
    SQL = "UPDATE SYS_USR SET "
    SQL = SQL & " USR_MSBLQL = '1', "
    SQL = SQL & " USR_MSBLQD = convert(varchar,getDate(),112) "
    SQL = SQL & " WHERE USR_ID IN (" & saqDadosadvIds & ")"
    SQL = SQL & " AND USR_MSBLQL = '2'"

    On Error Resume Next
    SET SAQ_DADOSADV = Conecta("SAQ-DADOSADV", SQL)
    
    if Err.Number = 0 AND isObject(SAQ_DADOSADV) then
        SAQ_DADOSADV.Close()
        SET SAQ_DADOSADV = NOTHING
    end if
    
    On Error Goto 0
end if 

'----------------------------------------------
'           USUARIOS PROTHEUS - MARANGUAPE 
'----------------------------------------------

if forDadosadvIds <> "" AND forDadosadvIds <> "undefined" then     
    SQL = "UPDATE SYS_USR SET "
    SQL = SQL & " USR_MSBLQL = '1', "
    SQL = SQL & " USR_MSBLQD = convert(varchar,getDate(),112) "
    SQL = SQL & " WHERE USR_ID IN (" & forDadosadvIds & ")"
    SQL = SQL & " AND USR_MSBLQL = '2'"

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

if benDadosadvIds <> "" AND benDadosadvIds <> "undefined" then     
    SQL = "UPDATE SYS_USR SET "
    SQL = SQL & " USR_MSBLQL = '1', "
    SQL = SQL & " USR_MSBLQD = convert(varchar,getDate(),112) "
    SQL = SQL & " WHERE USR_ID IN (" & benDadosadvIds & ")"
    SQL = SQL & " AND USR_MSBLQL = '2'"

    On Error Resume Next
    SET BEN_DADOSADV = Conecta("BEN-DADOSADV", SQL)
    
    if Err.Number = 0 AND isObject(BEN_DADOSADV) then
        BEN_DADOSADV.Close()
        SET BEN_DADOSADV = NOTHING
    end if
    
    On Error Goto 0
end if 

'==================== INSERE LOGS ======================='

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

totalUsuarios = idx

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
            <h3>Usuários desabilitados com sucesso!</h3>

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
                                <th class="text-left">Dias sem acessar</th>
                                <th class="text-left">Data de bloqueio</th>
                            </tr>              
                        </thead>
                        <% IF totalUsuarios > 0 THEN %>
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
                                    <td class="text-left"><%=usuarios(i)("DiasSemAcessar")%></td>
                                    <td class="text-left"><%=FormatDatetime(Date(), 2)%></td>
                                </tr>
                            <% Next %>
                        </tbody>
                        <% END IF %>
                    </table>
                </div>
            </div>
        </div>

        <script>
            function Voltar() {               
                window.location.href = "RevisaoUsuarios.asp";
            }

            function exportarPDF() {
                const { jsPDF } = window.jspdf;
                const doc = new jsPDF('landscape');
                const dataAtual = new Date().toLocaleDateString('pt-BR');

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
                doc.text('Relatório de Usuários Desabilitados', 14, 15);
                doc.setFontSize(10);
                doc.text('Data de geração: ' + dataAtual + ' | Total de usuários desabilitados: ' + dados.length, 14, 22);
                
                doc.autoTable({
                    head: [['#', 'Nome', 'Email', 'Sistema', 'Dias sem acessar', 'Data de bloqueio']],
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
    </body>
</html>