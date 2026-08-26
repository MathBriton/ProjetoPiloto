<!--#include file="../includes/variaveis.asp" -->
<!--#include file="../funcoes/funcoes.asp" -->
<!--#include file="../cnn2/cnn2.asp" -->
<%
session.LCID = 1046

'----------------------------------------------
'          PROCESSAMENTO DE DADOS
'----------------------------------------------
Dim usuarios()
Dim totalUsuarios
Dim pesquisaExecutada
Dim diasInativos
Dim sistemaFiltro

totalUsuarios = 0
pesquisaExecutada = (Request.Form("btnPesquisar") = "1")

' Captura os filtros
If pesquisaExecutada Then
    diasInativos = Request.Form("diasInativos")
    sistemaFiltro = Request.Form("sistemaFiltro")
    If sistemaFiltro = "" Then sistemaFiltro = "TODOS"
End If

If pesquisaExecutada Then
    ReDim usuarios(0)
    Dim idx
    idx = 0
    
    Dim filtroTempo
    filtroTempo = ""
    If diasInativos <> "" AND diasInativos <> "0" Then
        filtroTempo = " AND DiasSemAcessar >= " & diasInativos
    End If
    
    '----------------------------------------------
    '          USUARIOS HELPDESK - SWNET 
    '----------------------------------------------
    If sistemaFiltro = "TODOS" OR sistemaFiltro = "swnet" Then
        SQL = "SELECT * FROM ("
        SQL = SQL & " Select CRM_R_E_C_N_O as Id, CRM_NOME as Nome, CRM_EMAILPRIMARIO as Email, 'SWNET' as Sistema,"
        SQL = SQL & " CASE WHEN CRM_DATACES = '' THEN CONVERT(DATE, CRM_DATCAD, 103) ELSE CONVERT(DATE, CRM_DATACES, 103) END DataUltimoAcesso,"
        SQL = SQL & " CASE WHEN CRM_DATACES = '' THEN DATEDIFF(DAY,CRM_DATCAD, GETDATE()) ELSE DATEDIFF(DAY,CRM_DATACES, GETDATE()) END AS DiasSemAcessar"
        SQL = SQL & " FROM CRM_DTI WHERE Bloqueado = 0 AND CRM_D_E_L_E_T = ''"
        SQL = SQL & ") AS subquery WHERE 1=1" & filtroTempo
        
        SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
        While NOT PIR_HELPDESK.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = PIR_HELPDESK("Id")
            usuarios(idx)("Nome") = PIR_HELPDESK("Nome")
            usuarios(idx)("Email") = PIR_HELPDESK("Email")
            usuarios(idx)("Sistema") = PIR_HELPDESK("Sistema")
            usuarios(idx)("DataUltimoAcesso") = PIR_HELPDESK("DataUltimoAcesso")
            usuarios(idx)("DiasSemAcessar") = PIR_HELPDESK("DiasSemAcessar")
            usuarios(idx)("SistemaId") = "swnet"
            idx = idx + 1
            PIR_HELPDESK.MoveNext()
        Wend
        If isObject(PIR_HELPDESK) Then
            PIR_HELPDESK.Close()
            SET PIR_HELPDESK = NOTHING
        End If
    End If

    '----------------------------------------------
    '         USUARIOS TRIMBOX - PIRAPETINGA 
    '----------------------------------------------
    If sistemaFiltro = "TODOS" OR sistemaFiltro = "pirTrimbox" Then
        SQL = "SELECT * FROM ("
        SQL = SQL & " Select Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - PIRAPETINGA' as Sistema,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN CONVERT(DATE, DataCriacao, 103) ELSE CONVERT(DATE, DataUltEntrada, 103) END as DataUltimoAcesso,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCriacao, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
        SQL = SQL & " FROM Usuario WHERE Desativado = 0"
        SQL = SQL & ") AS subquery WHERE 1=1" & filtroTempo
        
        SET PIR_TRIMBOX = Conecta("PIR-TRIMBOX", SQL)
        While NOT PIR_TRIMBOX.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = PIR_TRIMBOX("Id")
            usuarios(idx)("Nome") = PIR_TRIMBOX("Nome")
            usuarios(idx)("Email") = PIR_TRIMBOX("Email")
            usuarios(idx)("Sistema") = PIR_TRIMBOX("Sistema")
            usuarios(idx)("DataUltimoAcesso") = PIR_TRIMBOX("DataUltimoAcesso")
            usuarios(idx)("DiasSemAcessar") = PIR_TRIMBOX("DiasSemAcessar")
            usuarios(idx)("SistemaId") = "pirTrimbox"
            idx = idx + 1
            PIR_TRIMBOX.MoveNext()
        Wend
        If isObject(PIR_TRIMBOX) Then
            PIR_TRIMBOX.Close()
            SET PIR_TRIMBOX = NOTHING
        End If
    End If

    '----------------------------------------------
    '         USUARIOS TRIMBOX - BENTO 
    '----------------------------------------------
    If sistemaFiltro = "TODOS" OR sistemaFiltro = "benTrimbox" Then
        SQL = "SELECT * FROM ("
        SQL = SQL & " Select Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - BENTO' as Sistema,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN CONVERT(DATE, DataCriacao, 103) ELSE CONVERT(DATE, DataUltEntrada, 103) END as DataUltimoAcesso,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCriacao, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
        SQL = SQL & " FROM Usuario WHERE Desativado = 0"
        SQL = SQL & ") AS subquery WHERE 1=1" & filtroTempo
        
        SET BEN_TRIMBOX = Conecta("BEN-TRIMBOX", SQL)
        While NOT BEN_TRIMBOX.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = BEN_TRIMBOX("Id")
            usuarios(idx)("Nome") = BEN_TRIMBOX("Nome")
            usuarios(idx)("Email") = BEN_TRIMBOX("Email")
            usuarios(idx)("Sistema") = BEN_TRIMBOX("Sistema")
            usuarios(idx)("DataUltimoAcesso") = BEN_TRIMBOX("DataUltimoAcesso")
            usuarios(idx)("DiasSemAcessar") = BEN_TRIMBOX("DiasSemAcessar")
            usuarios(idx)("SistemaId") = "benTrimbox"
            idx = idx + 1
            BEN_TRIMBOX.MoveNext()
        Wend
        If isObject(BEN_TRIMBOX) Then
            BEN_TRIMBOX.Close()
            SET BEN_TRIMBOX = NOTHING
        End If
    End If

    '----------------------------------------------
    '         USUARIOS TRIMBOX - UBERABA 
    '----------------------------------------------
    If sistemaFiltro = "TODOS" OR sistemaFiltro = "ubeTrimbox" Then
        SQL = "SELECT * FROM ("
        SQL = SQL & " Select Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - UBERABA' as Sistema,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN CONVERT(DATE, DataCriacao, 103) ELSE CONVERT(DATE, DataUltEntrada, 103) END as DataUltimoAcesso,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCriacao, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
        SQL = SQL & " FROM Usuario WHERE Desativado = 0"
        SQL = SQL & ") AS subquery WHERE 1=1" & filtroTempo
        
        SET UBE_TRIMBOX = Conecta("UBE-TRIMBOX", SQL)
        While NOT UBE_TRIMBOX.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = UBE_TRIMBOX("Id")
            usuarios(idx)("Nome") = UBE_TRIMBOX("Nome")
            usuarios(idx)("Email") = UBE_TRIMBOX("Email")
            usuarios(idx)("Sistema") = UBE_TRIMBOX("Sistema")
            usuarios(idx)("DataUltimoAcesso") = UBE_TRIMBOX("DataUltimoAcesso")
            usuarios(idx)("DiasSemAcessar") = UBE_TRIMBOX("DiasSemAcessar")
            usuarios(idx)("SistemaId") = "ubeTrimbox"
            idx = idx + 1
            UBE_TRIMBOX.MoveNext()
        Wend
        If isObject(UBE_TRIMBOX) Then
            UBE_TRIMBOX.Close()
            SET UBE_TRIMBOX = NOTHING
        End If
    End If

    '----------------------------------------------
    '         USUARIOS TRIMBOX - MARANGUAPE 
    '----------------------------------------------
    If sistemaFiltro = "TODOS" OR sistemaFiltro = "forTrimbox" Then
        SQL = "SELECT * FROM ("
        SQL = SQL & " Select Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - MARANGUAPE' as Sistema,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN CONVERT(DATE, DataCriacao, 103) ELSE CONVERT(DATE, DataUltEntrada, 103) END as DataUltimoAcesso,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCriacao, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
        SQL = SQL & " FROM Usuario WHERE Desativado = 0"
        SQL = SQL & ") AS subquery WHERE 1=1" & filtroTempo
        
        SET FOR_TRIMBOX = Conecta("FOR-TRIMBOX", SQL)
        While NOT FOR_TRIMBOX.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = FOR_TRIMBOX("Id")
            usuarios(idx)("Nome") = FOR_TRIMBOX("Nome")
            usuarios(idx)("Email") = FOR_TRIMBOX("Email")
            usuarios(idx)("Sistema") = FOR_TRIMBOX("Sistema")
            usuarios(idx)("DataUltimoAcesso") = FOR_TRIMBOX("DataUltimoAcesso")
            usuarios(idx)("DiasSemAcessar") = FOR_TRIMBOX("DiasSemAcessar")
            usuarios(idx)("SistemaId") = "forTrimbox"
            idx = idx + 1
            FOR_TRIMBOX.MoveNext()
        Wend
        If isObject(FOR_TRIMBOX) Then
            FOR_TRIMBOX.Close()
            SET FOR_TRIMBOX = NOTHING
        End If
    End If

    '----------------------------------------------
    '         USUARIOS TRIMBOX - SAQUAREMA 
    '----------------------------------------------
    If sistemaFiltro = "TODOS" OR sistemaFiltro = "saqTrimbox" Then
        SQL = "SELECT * FROM ("
        SQL = SQL & " Select Usuario as Id, NomeCompleto as Nome, Email, 'TRIMBOX - SAQUAREMA' as Sistema,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN CONVERT(DATE, DataCriacao, 103) ELSE CONVERT(DATE, DataUltEntrada, 103) END as DataUltimoAcesso,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCriacao, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
        SQL = SQL & " FROM Usuario WHERE Desativado = 0"
        SQL = SQL & ") AS subquery WHERE 1=1" & filtroTempo
        
        SET SAQ_TRIMBOX = Conecta("SAQ-TRIMBOX", SQL)
        While NOT SAQ_TRIMBOX.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = SAQ_TRIMBOX("Id")
            usuarios(idx)("Nome") = SAQ_TRIMBOX("Nome")
            usuarios(idx)("Email") = SAQ_TRIMBOX("Email")
            usuarios(idx)("Sistema") = SAQ_TRIMBOX("Sistema")
            usuarios(idx)("DataUltimoAcesso") = SAQ_TRIMBOX("DataUltimoAcesso")
            usuarios(idx)("DiasSemAcessar") = SAQ_TRIMBOX("DiasSemAcessar")
            usuarios(idx)("SistemaId") = "saqTrimbox"
            idx = idx + 1
            SAQ_TRIMBOX.MoveNext()
        Wend
        If isObject(SAQ_TRIMBOX) Then
            SAQ_TRIMBOX.Close()
            SET SAQ_TRIMBOX = NOTHING
        End If
    End If

    '----------------------------------------------
    '         USUARIOS TRIMPAPER - BENTO 
    '----------------------------------------------
    If sistemaFiltro = "TODOS" OR sistemaFiltro = "benTrimpaper" Then
        SQL = "SELECT * FROM ("
        SQL = SQL & " Select IDUsuario as Id, NomeCompleto as Nome, Email, 'TRIMPAPER - BENTO' as Sistema,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN CONVERT(DATE, DataCadastro, 103) ELSE CONVERT(DATE, DataUltEntrada, 103) END as DataUltimoAcesso,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCadastro, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
        SQL = SQL & " FROM Usuarios WHERE Desativado = 0"
        SQL = SQL & ") AS subquery WHERE 1=1" & filtroTempo
        
        SET BEN_TRIMPAPER = Conecta("BEN-TRIMPAPER", SQL)
        While NOT BEN_TRIMPAPER.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = BEN_TRIMPAPER("Id")
            usuarios(idx)("Nome") = BEN_TRIMPAPER("Nome")
            usuarios(idx)("Email") = BEN_TRIMPAPER("Email")
            usuarios(idx)("Sistema") = BEN_TRIMPAPER("Sistema")
            usuarios(idx)("DataUltimoAcesso") = BEN_TRIMPAPER("DataUltimoAcesso")
            usuarios(idx)("DiasSemAcessar") = BEN_TRIMPAPER("DiasSemAcessar")
            usuarios(idx)("SistemaId") = "benTrimpaper"
            idx = idx + 1
            BEN_TRIMPAPER.MoveNext()
        Wend
        If isObject(BEN_TRIMPAPER) Then
            BEN_TRIMPAPER.Close()
            SET BEN_TRIMPAPER = NOTHING
        End If
    End If

    '----------------------------------------------
    '         USUARIOS TRIMPAPER - PIRAPETINGA 
    '----------------------------------------------
    If sistemaFiltro = "TODOS" OR sistemaFiltro = "pirTrimpaper" Then
        SQL = "SELECT * FROM ("
        SQL = SQL & " Select IDUsuario as Id, NomeCompleto as Nome, Email, 'TRIMPAPER - PIRAPETINGA' as Sistema,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN CONVERT(DATE, DataCadastro, 103) ELSE CONVERT(DATE, DataUltEntrada, 103) END as DataUltimoAcesso,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCadastro, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
        SQL = SQL & " FROM Usuarios WHERE Desativado = 0"
        SQL = SQL & ") AS subquery WHERE 1=1" & filtroTempo
        
        SET PIR_TRIMPAPER = Conecta("PIR-TRIMPAPER", SQL)
        While NOT PIR_TRIMPAPER.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = PIR_TRIMPAPER("Id")
            usuarios(idx)("Nome") = PIR_TRIMPAPER("Nome")
            usuarios(idx)("Email") = PIR_TRIMPAPER("Email")
            usuarios(idx)("Sistema") = PIR_TRIMPAPER("Sistema")
            usuarios(idx)("DataUltimoAcesso") = PIR_TRIMPAPER("DataUltimoAcesso")
            usuarios(idx)("DiasSemAcessar") = PIR_TRIMPAPER("DiasSemAcessar")
            usuarios(idx)("SistemaId") = "pirTrimpaper"
            idx = idx + 1
            PIR_TRIMPAPER.MoveNext()
        Wend
        If isObject(PIR_TRIMPAPER) Then
            PIR_TRIMPAPER.Close()
            SET PIR_TRIMPAPER = NOTHING
        End If
    End If

    '----------------------------------------------
    '         USUARIOS TRIMPAPER - UBERABA 
    '----------------------------------------------
    If sistemaFiltro = "TODOS" OR sistemaFiltro = "ubeTrimpaper" Then
        SQL = "SELECT * FROM ("
        SQL = SQL & " Select IDUsuario as Id, NomeCompleto as Nome, Email, 'TRIMPAPER - UBERABA' as Sistema,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN CONVERT(DATE, DataCadastro, 103) ELSE CONVERT(DATE, DataUltEntrada, 103) END as DataUltimoAcesso,"
        SQL = SQL & " CASE WHEN DataUltEntrada IS NULL THEN DATEDIFF(DAY,CONVERT(DATE, DataCadastro, 103), CONVERT(DATE,GETDATE(), 103)) ELSE DATEDIFF(DAY,CONVERT(DATE,DataUltEntrada, 103), CONVERT(DATE,GETDATE(), 103)) END AS DiasSemAcessar"
        SQL = SQL & " FROM Usuarios WHERE Desativado = 0"
        SQL = SQL & ") AS subquery WHERE 1=1" & filtroTempo
        
        SET UBE_TRIMPAPER = Conecta("UBE-TRIMPAPER", SQL)
        While NOT UBE_TRIMPAPER.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = UBE_TRIMPAPER("Id")
            usuarios(idx)("Nome") = UBE_TRIMPAPER("Nome")
            usuarios(idx)("Email") = UBE_TRIMPAPER("Email")
            usuarios(idx)("Sistema") = UBE_TRIMPAPER("Sistema")
            usuarios(idx)("DataUltimoAcesso") = UBE_TRIMPAPER("DataUltimoAcesso")
            usuarios(idx)("DiasSemAcessar") = UBE_TRIMPAPER("DiasSemAcessar")
            usuarios(idx)("SistemaId") = "ubeTrimpaper"
            idx = idx + 1
            UBE_TRIMPAPER.MoveNext()
        Wend
        If isObject(UBE_TRIMPAPER) Then
            UBE_TRIMPAPER.Close()
            SET UBE_TRIMPAPER = NOTHING
        End If
    End If

    '----------------------------------------------
    '         USUARIOS PROTHEUS - PIRAPETINGA 
    '----------------------------------------------
    If sistemaFiltro = "TODOS" OR sistemaFiltro = "pirDadosadv" Then
        SQL = "SELECT * FROM ( "
        SQL = SQL & "	 Select u.USR_ID as Id, u.USR_NOME as Nome, u.USR_EMAIL as Email, 'PROTHEUS - PIRAPETINGA' as Sistema, "
        SQL = SQL & "	 CASE WHEN "
        SQL = SQL & "		CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON),103) <> '' "
        SQL = SQL & "		THEN CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON),103) "
        SQL = SQL & "		ELSE CONVERT(DATE,u.USR_DTINC, 103 "
        SQL = SQL & "	 ) END AS DataUltimoAcesso, "
        SQL = SQL & "	 CASE WHEN ISNULL(s.USR_DTLOGON,u.USR_DTLOGON) <> '' "
        SQL = SQL & "		 THEN DATEDIFF(DAY,CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON), 103), CONVERT(DATE,GETDATE(), 103)) "
        SQL = SQL & "		 ELSE DATEDIFF(DAY,CONVERT(DATE,u.USR_DTINC, 103), CONVERT(DATE,GETDATE(), 103) "
        SQL = SQL & "	 ) END AS DiasSemAcessar "
        SQL = SQL & "	 FROM SYS_USR u "
        SQL = SQL & "	 LEFT JOIN SYS_USR_LOGIN s on s.USR_ID = u.USR_ID "
        SQL = SQL & "	 WHERE u.USR_MSBLQL = '2' "
        SQL = SQL & "	 AND u.USR_ID <> '000000' "
        SQL = SQL & "	 AND u.D_E_L_E_T_ = ''"
        SQL = SQL & ") AS subquery  WHERE 1=1 " & filtroTempo
        
        SET PIR_DADOSADV = Conecta("PIR-DADOSADV", SQL)
        While NOT PIR_DADOSADV.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = PIR_DADOSADV("Id")
            usuarios(idx)("Nome") = PIR_DADOSADV("Nome")
            usuarios(idx)("Email") = PIR_DADOSADV("Email")
            usuarios(idx)("Sistema") = PIR_DADOSADV("Sistema")
            usuarios(idx)("DataUltimoAcesso") = PIR_DADOSADV("DataUltimoAcesso")
            usuarios(idx)("DiasSemAcessar") = PIR_DADOSADV("DiasSemAcessar")
            usuarios(idx)("SistemaId") = "pirDadosadv"
            idx = idx + 1
            PIR_DADOSADV.MoveNext()
        Wend
        If isObject(PIR_DADOSADV) Then
            PIR_DADOSADV.Close()
            SET PIR_DADOSADV = NOTHING
        End If
    End If

    '----------------------------------------------
    '         USUARIOS PROTHEUS - SAQUAREMA 
    '----------------------------------------------
    If sistemaFiltro = "TODOS" OR sistemaFiltro = "saqDadosadv" Then 
        SQL = "SELECT * FROM ( "
        SQL = SQL & "	 Select u.USR_ID as Id, u.USR_NOME as Nome, u.USR_EMAIL as Email, 'PROTHEUS - SAQUAREMA' as Sistema, "
        SQL = SQL & "	 CASE WHEN "
        SQL = SQL & "		CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON),103) <> '' "
        SQL = SQL & "		THEN CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON),103) "
        SQL = SQL & "		ELSE CONVERT(DATE,u.USR_DTINC, 103 "
        SQL = SQL & "	 ) END AS DataUltimoAcesso, "
        SQL = SQL & "	 CASE WHEN ISNULL(s.USR_DTLOGON,u.USR_DTLOGON) <> '' "
        SQL = SQL & "		 THEN DATEDIFF(DAY,CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON), 103), CONVERT(DATE,GETDATE(), 103)) "
        SQL = SQL & "		 ELSE DATEDIFF(DAY,CONVERT(DATE,u.USR_DTINC, 103), CONVERT(DATE,GETDATE(), 103) "
        SQL = SQL & "	 ) END AS DiasSemAcessar "
        SQL = SQL & "	 FROM SYS_USR u "
        SQL = SQL & "	 LEFT JOIN SYS_USR_LOGIN s on s.USR_ID = u.USR_ID "
        SQL = SQL & "	 WHERE u.USR_MSBLQL = '2' "
        SQL = SQL & "	 AND u.USR_ID <> '000000' "
        SQL = SQL & "	 AND u.D_E_L_E_T_ = ''"
        SQL = SQL & ") AS subquery  WHERE 1=1 " & filtroTempo
        
        SET SAQ_DADOSADV = Conecta("SAQ-DADOSADV", SQL)
        While NOT SAQ_DADOSADV.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = SAQ_DADOSADV("Id")
            usuarios(idx)("Nome") = SAQ_DADOSADV("Nome")
            usuarios(idx)("Email") = SAQ_DADOSADV("Email")
            usuarios(idx)("Sistema") = SAQ_DADOSADV("Sistema")
            usuarios(idx)("DataUltimoAcesso") = SAQ_DADOSADV("DataUltimoAcesso")
            usuarios(idx)("DiasSemAcessar") = SAQ_DADOSADV("DiasSemAcessar")
            usuarios(idx)("SistemaId") = "saqDadosadv"
            idx = idx + 1
            SAQ_DADOSADV.MoveNext()
        Wend
        If isObject(SAQ_DADOSADV) Then
            SAQ_DADOSADV.Close()
            SET SAQ_DADOSADV = NOTHING
        End If
    End If

    '----------------------------------------------
    '         USUARIOS PROTHEUS - MARANGUAPE 
    '----------------------------------------------
    If sistemaFiltro = "TODOS" OR sistemaFiltro = "forDadosadv" Then
        SQL = "SELECT * FROM ( "
        SQL = SQL & "	 Select u.USR_ID as Id, u.USR_NOME as Nome, u.USR_EMAIL as Email, 'PROTHEUS - MARANGUAPE' as Sistema, "
        SQL = SQL & "	 CASE WHEN "
        SQL = SQL & "		CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON),103) <> '' "
        SQL = SQL & "		THEN CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON),103) "
        SQL = SQL & "		ELSE CONVERT(DATE,u.USR_DTINC, 103 "
        SQL = SQL & "	 ) END AS DataUltimoAcesso, "
        SQL = SQL & "	 CASE WHEN ISNULL(s.USR_DTLOGON,u.USR_DTLOGON) <> '' "
        SQL = SQL & "		 THEN DATEDIFF(DAY,CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON), 103), CONVERT(DATE,GETDATE(), 103)) "
        SQL = SQL & "		 ELSE DATEDIFF(DAY,CONVERT(DATE,u.USR_DTINC, 103), CONVERT(DATE,GETDATE(), 103) "
        SQL = SQL & "	 ) END AS DiasSemAcessar "
        SQL = SQL & "	 FROM SYS_USR u "
        SQL = SQL & "	 LEFT JOIN SYS_USR_LOGIN s on s.USR_ID = u.USR_ID "
        SQL = SQL & "	 WHERE u.USR_MSBLQL = '2' "
        SQL = SQL & "	 AND u.USR_ID <> '000000' "
        SQL = SQL & "	 AND u.D_E_L_E_T_ = ''"
        SQL = SQL & ") AS subquery  WHERE 1=1 " & filtroTempo
        
        SET FOR_DADOSADV = Conecta("FOR-DADOSADV", SQL)
        While NOT FOR_DADOSADV.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = FOR_DADOSADV("Id")
            usuarios(idx)("Nome") = FOR_DADOSADV("Nome")
            usuarios(idx)("Email") = FOR_DADOSADV("Email")
            usuarios(idx)("Sistema") = FOR_DADOSADV("Sistema")
            usuarios(idx)("DataUltimoAcesso") = FOR_DADOSADV("DataUltimoAcesso")
            usuarios(idx)("DiasSemAcessar") = FOR_DADOSADV("DiasSemAcessar")
            usuarios(idx)("SistemaId") = "forDadosadv"
            idx = idx + 1
            FOR_DADOSADV.MoveNext()
        Wend
        If isObject(FOR_DADOSADV) Then
            FOR_DADOSADV.Close()
            SET FOR_DADOSADV = NOTHING
        End If
    End If

    '----------------------------------------------
    '         USUARIOS PROTHEUS - BENTO 
    '----------------------------------------------
    If sistemaFiltro = "TODOS" OR sistemaFiltro = "benDadosadv" Then
        SQL = "SELECT * FROM ( "
        SQL = SQL & "	 Select u.USR_ID as Id, u.USR_NOME as Nome, u.USR_EMAIL as Email, 'PROTHEUS - BENTO' as Sistema, "
        SQL = SQL & "	 CASE WHEN "
        SQL = SQL & "		CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON),103) <> '' "
        SQL = SQL & "		THEN CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON),103) "
        SQL = SQL & "		ELSE CONVERT(DATE,u.USR_DTINC, 103 "
        SQL = SQL & "	 ) END AS DataUltimoAcesso, "
        SQL = SQL & "	 CASE WHEN ISNULL(s.USR_DTLOGON,u.USR_DTLOGON) <> '' "
        SQL = SQL & "		 THEN DATEDIFF(DAY,CONVERT(DATE,ISNULL(s.USR_DTLOGON,u.USR_DTLOGON), 103), CONVERT(DATE,GETDATE(), 103)) "
        SQL = SQL & "		 ELSE DATEDIFF(DAY,CONVERT(DATE,u.USR_DTINC, 103), CONVERT(DATE,GETDATE(), 103) "
        SQL = SQL & "	 ) END AS DiasSemAcessar "
        SQL = SQL & "	 FROM SYS_USR u "
        SQL = SQL & "	 LEFT JOIN SYS_USR_LOGIN s on s.USR_ID = u.USR_ID "
        SQL = SQL & "	 WHERE u.USR_MSBLQL = '2' "
        SQL = SQL & "	 AND u.USR_ID <> '000000' "
        SQL = SQL & "	 AND u.D_E_L_E_T_ = ''"
        SQL = SQL & ") AS subquery  WHERE 1=1 " & filtroTempo
        
        SET BEN_DADOSADV = Conecta("BEN-DADOSADV", SQL)
        While NOT BEN_DADOSADV.EOF
            ReDim Preserve usuarios(idx)
            Set usuarios(idx) = Server.CreateObject("Scripting.Dictionary")
            usuarios(idx)("Id") = BEN_DADOSADV("Id")
            usuarios(idx)("Nome") = BEN_DADOSADV("Nome")
            usuarios(idx)("Email") = BEN_DADOSADV("Email")
            usuarios(idx)("Sistema") = BEN_DADOSADV("Sistema")
            usuarios(idx)("DataUltimoAcesso") = BEN_DADOSADV("DataUltimoAcesso")
            usuarios(idx)("DiasSemAcessar") = BEN_DADOSADV("DiasSemAcessar")
            usuarios(idx)("SistemaId") = "benDadosadv"
            idx = idx + 1
            BEN_DADOSADV.MoveNext()
        Wend
        If isObject(BEN_DADOSADV) Then
            BEN_DADOSADV.Close()
            SET BEN_DADOSADV = NOTHING
        End If
    End If
    
    totalUsuarios = idx
End If
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

            const nomesSistemas = {
                'swnet': 'SWNET',
                'pirTrimbox': 'TRIMBOX - PIRAPETINGA',
                'benTrimbox': 'TRIMBOX - BENTO',
                'ubeTrimbox': 'TRIMBOX - UBERABA',
                'forTrimbox': 'TRIMBOX - MARANGUAPE',
                'saqTrimbox': 'TRIMBOX - SAQUAREMA',
                'benTrimpaper': 'TRIMPAPER - BENTO',
                'pirTrimpaper': 'TRIMPAPER - PIRAPETINGA',
                'ubeTrimpaper': 'TRIMPAPER - UBERABA',
                'pirDadosadv': 'PROTHEUS - PIRAPETINGA',
                'saqDadosadv': 'PROTHEUS - SAQUAREMA',
                'forDadosadv': 'PROTHEUS - MARANGUAPE',
                'benDadosadv': 'PROTHEUS - BENTO'
            };

            let selecoes = new Set();
            let sistemaAtual = null;

            function SelecionarUsuario(checkbox) {
                const userId = checkbox.dataset.userId;
                const sistema = checkbox.dataset.sistema;
                const key = `${sistema}:${userId}`;
                
                if (checkbox.checked) {
                    // Se é a primeira seleção, define o sistema atual
                    if (sistemaAtual === null) {
                        sistemaAtual = sistema;
                        desabilitarOutrosSistemas(sistema);
                    }
                    // Verifica se está tentando selecionar outro sistema
                    else if (sistemaAtual !== sistema) {
                        alert('Você só pode selecionar usuários de um sistema por vez.');
                        checkbox.checked = false;
                        return;
                    }
                    selecoes.add(key);
                } else {
                    selecoes.delete(key);
                    
                    // Se não há mais seleções, libera todos os sistemas
                    if (selecoes.size === 0) {
                        sistemaAtual = null;
                        habilitarTodosSistemas();
                        document.getElementById('selectAll').checked = false;
                    }
                }
                
                atualizarVisibilidadeBotoes();
            }

            function SelecionarTodos(checkboxGeral) {
                if (checkboxGeral.checked) {
                    // Pega o primeiro checkbox disponível para determinar o sistema
                    const primeiroCheckbox = document.querySelector('tbody input[type="checkbox"]:not(:disabled)');
                    if (primeiroCheckbox) {
                        sistemaAtual = primeiroCheckbox.dataset.sistema;
                    }
                    
                    // Seleciona apenas os checkboxes do sistema atual
                    const checkboxes = document.querySelectorAll(`tbody input[type="checkbox"][data-sistema="${sistemaAtual}"]:not(:disabled)`);
                    checkboxes.forEach(cb => {
                        cb.checked = true;
                        const key = `${cb.dataset.sistema}:${cb.dataset.userId}`;
                        selecoes.add(key);
                    });
                    
                    desabilitarOutrosSistemas(sistemaAtual);
                } else {
                    // Desmarca todos
                    const checkboxes = document.querySelectorAll('tbody input[type="checkbox"]');
                    checkboxes.forEach(cb => cb.checked = false);
                    selecoes.clear();
                    sistemaAtual = null;
                    habilitarTodosSistemas();
                }
                
                atualizarVisibilidadeBotoes();
            }

            function desabilitarOutrosSistemas(sistemaPermitido) {
                const todosCheckboxes = document.querySelectorAll('tbody input[type="checkbox"]');
                todosCheckboxes.forEach(cb => {
                    if (cb.dataset.sistema !== sistemaPermitido) {
                        cb.disabled = true;
                        cb.closest('tr').style.opacity = '0.5';
                    }
                });
            }

            function habilitarTodosSistemas() {
                const todosCheckboxes = document.querySelectorAll(`tbody input[type="checkbox"]:not(:disabled)`);

                todosCheckboxes.forEach(cb => {
                    cb.disabled = false;
                    cb.closest('tr').style.opacity = '1';
                });
            }

            function atualizarVisibilidadeBotoes() {
                const botoes = document.querySelectorAll('.btn-confirmar');
                const existemSelecoes = selecoes.size > 0;
                
                botoes.forEach(btn => {
                    btn.classList.toggle('hidden', !existemSelecoes);
                });
                
                // Atualiza o texto do alerta com o sistema selecionado
                if (existemSelecoes && sistemaAtual) {
                    const infoSistema = document.getElementById('infoSistema');
                    if (infoSistema) {
                        infoSistema.textContent = `Sistema selecionado: ${nomesSistemas[sistemaAtual] || sistemaAtual}`;
                        infoSistema.style.display = 'block';
                    }
                } else {
                    const infoSistema = document.getElementById('infoSistema');
                    if (infoSistema) {
                        infoSistema.style.display = 'none';
                    }
                }
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

            function DesabilitarUsuarios() {
                if (selecoes.size === 0) {
                    alert('Nenhum usuário selecionado');
                    return;
                }
                
                let params = new URLSearchParams();
                
                const agrupado = agruparPorSistema();
                Object.entries(agrupado).forEach(([sistema, ids]) => {
                    params.append(sistema, ids.join(','));
                });
                
                const url = `RevisaoUsuariosRepositorio.asp?${params.toString()}`;
                const nomeSistema = nomesSistemas[sistemaAtual] || sistemaAtual;
                
                if (confirm(`Deseja realmente desabilitar ${selecoes.size} usuário(s) do sistema ${nomeSistema}?`)) {
                    window.location.href = url;
                }
            }

            let ordemAtual = 'asc'; 
            
            function OrdenarPorDias() {
                const tabela = document.getElementById('Grade');
                const tbody = tabela.querySelector('tbody');
                const linhas = Array.from(tbody.querySelectorAll('tr'));
                
                ordemAtual = ordemAtual === 'asc' ? 'desc' : 'asc';
                
                linhas.sort((a, b) => {
                    const diasA = parseInt(a.cells[6].textContent);
                    const diasB = parseInt(b.cells[6].textContent);
                    
                    if (ordemAtual === 'asc') {
                        return diasA - diasB;
                    } else {
                        return diasB - diasA;
                    }
                });
                
                while (tbody.firstChild) {
                    tbody.removeChild(tbody.firstChild);
                }
                
                linhas.forEach(linha => tbody.appendChild(linha));
                
                atualizarIconeOrdenacao();
            }
            
            function atualizarIconeOrdenacao() {
                const icone = document.getElementById('sortIcon');
                if (ordemAtual === 'asc') {
                    icone.className = 'fas fa-sort-up';
                } else {
                    icone.className = 'fas fa-sort-down';
                }
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
                doc.text('Relatório de Revisão de Usuários', 14, 15);
                doc.setFontSize(10);
                doc.text('Data de geração: ' + dataAtual + ' | Total de usuários: ' + dados.length, 14, 22);
                
                doc.autoTable({
                    head: [['', '#', 'Nome', 'Email', 'Sistema', 'Último acesso', 'Dias sem acessar']],
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
            .status-indicator {
                width: 15px;
                height: 15px;
                border-radius: 50%;
                display: inline-block;
                margin: 0 auto;
            }
            .status-verde {
                background-color: #28a745;
            }
            .status-amarelo {
                background-color: #ffc107;
            }
            .status-laranja {
                background-color: #fd7e14;
            }
            .status-vermelho {
                background-color: #dc3545;
            }
            .text-center-content {
                text-align: center;
            }
            .sortable-header {
                cursor: pointer;
                user-select: none;
                position: relative;
                padding-right: 20px;
            }
            .sortable-header:hover {
                background-color: #e9ecef;
            }
            .sortable-header i {
                margin-left: 8px;
                color: #6c757d;
            }
            #infoSistema {
                display: none;
                font-size: 0.9rem;
                color: #0056b3;
                font-weight: 600;
                margin-left: 10px;
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
                    <h3>Revisão de Usuários</h3>
                </div>

                <div class="filtro-container">                    
                    <div class="col-md-4">
                        <label for="diasInativos" class="form-label"><strong>Dias sem acessar:</strong></label>
                        <select name="diasInativos" id="diasInativos" class="form-select form-select-sm">
                            <option value="0" <%If Request.Form("diasInativos") = "0" OR Request.Form("diasInativos") = "" Then Response.Write("selected")%>>Sem filtro</option>
                            <option value="60" <%If Request.Form("diasInativos") = "60" Then Response.Write("selected")%>>60 dias ou mais</option>
                        </select>
                    </div>

                    <div class="col-md-4">
                        <label for="sistemaFiltro" class="form-label"><strong>Sistema:</strong></label>
                        <select name="sistemaFiltro" id="sistemaFiltro" class="form-select form-select-sm">
                            <option value="TODOS" <%If Request.Form("sistemaFiltro") = "TODOS" OR Request.Form("sistemaFiltro") = "" Then Response.Write("selected")%>>Todos os Sistemas</option>
                            <option value="swnet" <%If Request.Form("sistemaFiltro") = "swnet" Then Response.Write("selected")%>>SWNET</option>
                            <option value="benTrimbox" <%If Request.Form("sistemaFiltro") = "benTrimbox" Then Response.Write("selected")%>>TRIMBOX - Bento</option>
                            <option value="forTrimbox" <%If Request.Form("sistemaFiltro") = "forTrimbox" Then Response.Write("selected")%>>TRIMBOX - Maranguape</option>
                            <option value="pirTrimbox" <%If Request.Form("sistemaFiltro") = "pirTrimbox" Then Response.Write("selected")%>>TRIMBOX - Pirapetinga</option>
                            <option value="saqTrimbox" <%If Request.Form("sistemaFiltro") = "saqTrimbox" Then Response.Write("selected")%>>TRIMBOX - Saquarema</option>
                            <option value="ubeTrimbox" <%If Request.Form("sistemaFiltro") = "ubeTrimbox" Then Response.Write("selected")%>>TRIMBOX - Uberaba</option>
                            <option value="benTrimpaper" <%If Request.Form("sistemaFiltro") = "benTrimpaper" Then Response.Write("selected")%>>TRIMPAPER - Bento</option>
                            <option value="pirTrimpaper" <%If Request.Form("sistemaFiltro") = "pirTrimpaper" Then Response.Write("selected")%>>TRIMPAPER - Pirapetinga</option>
                            <option value="ubeTrimpaper" <%If Request.Form("sistemaFiltro") = "ubeTrimpaper" Then Response.Write("selected")%>>TRIMPAPER - Uberaba</option>
                            <option value="benDadosadv" <%If Request.Form("sistemaFiltro") = "benDadosadv" Then Response.Write("selected")%>>PROTHEUS - Bento</option>
                            <option value="forDadosadv" <%If Request.Form("sistemaFiltro") = "forDadosadv" Then Response.Write("selected")%>>PROTHEUS - Maranguape</option>
                            <option value="pirDadosadv" <%If Request.Form("sistemaFiltro") = "pirDadosadv" Then Response.Write("selected")%>>PROTHEUS - Pirapetinga</option>
                            <option value="saqDadosadv" <%If Request.Form("sistemaFiltro") = "saqDadosadv" Then Response.Write("selected")%>>PROTHEUS - Saquarema</option>
                        </select>
                    </div>

                    <div class="col-md-4">
                        <button type="submit" name="btnPesquisar" value="1" class="btn btn-sm btn-primary">
                            <i class="fas fa-search me-2"></i> Pesquisar
                        </button>
                    </div>
                </div>

                <% IF pesquisaExecutada AND totalUsuarios = 0 THEN %>
                    <div class="alert alert-warning mt-3" role="alert">
                        <h5 class="alert-heading">Nenhum usuário encontrado!</h5>
                        <p>Não foram encontrados usuários com os filtros selecionados.</p>
                    </div>
                <% END IF %>

                <% IF pesquisaExecutada AND totalUsuarios > 0 THEN %>
                    <div class="row align-middle">
                        <div class="col-md-12">
                            <div class="alert alert-info">
                                <strong><%=totalUsuarios%></strong> usuário(s) encontrado(s)
                                <% If sistemaFiltro = "TODOS" Then %>
                                    <br><small class="text-muted">Você pode selecionar usuários de apenas um sistema por vez para desabilitar.</small>
                                <% End If %>
                            </div>

                            <div class="col-md-12 mb-3 d-flex gap-2">
                                <button type="button" class="btn btn-danger btn-sm" onclick="exportarPDF()">
                                    <i class="fas fa-file-pdf mr-3"></i> Exportar para PDF
                                </button>

                                <div class="d-flex align-items-center">
                                    <button type="button" class="btn-confirmar btn-sm btn btn-danger hidden" onclick="DesabilitarUsuarios('DESABILITAR')" title="Confirmar Cancelamento">
                                        <i class="fa fa-ban me-2"></i>
                                        <span>Desabilitar selecionados</span>
                                    </button>
                                    <span id="infoSistema" class="ms-3"></span>
                                </div>
                            </div>
                            
                            <table class="table table-bordered" name="Grade" id="Grade">
                                <thead>
                                    <tr>
                                        <th class="text-left">
                                            <input type="checkbox" id="selectAll" onchange="SelecionarTodos(this)">
                                        </th>
                                        <th class="text-left">#</th>
                                        <th class="text-left">Nome</th>
                                        <th class="text-left">Email</th>
                                        <th class="text-left">Sistema</th>
                                        <th class="text-left">Último acesso</th>
                                        <th class="text-left sortable-header" onclick="OrdenarPorDias()" title="Clique para ordenar">
                                            Dias sem acessar
                                            <i id="sortIcon" class="fas fa-sort" style="color: #6c757d;"></i>
                                        </th>
                                        <th class="text-center-content">Alerta</th>
                                    </tr>              
                                </thead>
                                <tbody>
                                    <% 
                                    Dim i, diasSemAcessar, statusClass
                                    For i = 0 To UBound(usuarios)
                                        diasSemAcessar = usuarios(i)("DiasSemAcessar")
                                        
                                        If diasSemAcessar < 60 Then
                                            statusClass = ""
                                        Else
                                            statusClass = "status-vermelho"
                                        End If
                                    %>
                                        <tr>   
                                            <% If diasSemAcessar >= 60 Then %>
                                                <td class="text-left">
                                                    <input type="checkbox" 
                                                        data-user-id="<%=usuarios(i)("Id")%>" 
                                                        data-sistema="<%=usuarios(i)("SistemaId")%>"
                                                        onchange="SelecionarUsuario(this)">
                                                </td>
                                            <% Else %>
                                                <td class="text-left">
                                                    <input type="checkbox" disabled>
                                                </td>
                                            <% End If %>
                                            <td class="text-left"><%=usuarios(i)("Id")%></td>
                                            <td class="text-left"><%=usuarios(i)("Nome")%></td>
                                            <td class="text-left"><%=usuarios(i)("Email")%></td>
                                            <td class="text-left"><%=usuarios(i)("Sistema")%></td>
                                            <td class="text-left"><%=usuarios(i)("DataUltimoAcesso")%></td>
                                            <td class="text-left"><%=usuarios(i)("DiasSemAcessar")%></td>
                                            <td class="text-center-content">
                                                <span class="status-indicator <%=statusClass%>" title="<%=diasSemAcessar%> dias sem acessar"></span>
                                            </td>
                                        </tr>
                                    <% Next %>
                                </tbody>
                            </table>
                        </div>
                    </div>
                <% END IF %>
            </div>
        </form>
    </body>
</html>