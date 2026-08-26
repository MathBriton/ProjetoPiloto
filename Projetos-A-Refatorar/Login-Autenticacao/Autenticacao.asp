<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="corporativo/cnn2/cnn2.asp"-->
<!--#Include File="corporativo/funcoes/funcoes.asp"-->
<%
Response.Charset = "UTF-8"
Response.CodePage = 65001
Response.Buffer = true
Response.Expires = 0
Session.LCID = 1046
Response.ContentType = "application/json"

metodoAutenticacao = Request.Form("metodoAutenticacao")

Select Case metodoAutenticacao
    Case "WINDOWS"
        resultado = AutenticarWindows()
    Case "SISTEMA"
        resultado = AutenticarSistema()
    Case "AD"
        resultado = AutenticarAD()
    Case Else
        resultado = "{""sucesso"": false, ""mensagem"": ""Método de autenticação inválido: " & metodoAutenticacao & """}"
End Select

' Retorna resposta
Response.Write(resultado)

' Autentica com usuário logado no windows
Function AutenticarWindows() 
    usuario = SanitizeSQL(Request.Form("usuario"))
    usuarioWindows = Request.ServerVariables("AUTH_USER")
        
    If usuario = "" Or usuarioWindows = "" Then
        Call AcessoFalhou(usuario, "WINDOWS")
        AutenticarWindows = "{""sucesso"": false, ""mensagem"": ""Por favor atualize a página e tente novamente""}"
        Exit Function
    End If

    If usuario <> usuarioWindows Then
        Call AcessoFalhou(usuario, "WINDOWS")
        AutenticarWindows = "{""sucesso"": false, ""mensagem"": ""Usuário ou senha inválidos""}"
        Exit Function
    End If

    ' Busca usuário no sistema
    Set dadosUsuario = VerificarUsuarioSistema("WINDOWS", usuarioWindows, "")
    If Not dadosUsuario("sucesso") Then
        Call AcessoFalhou(usuario, "WINDOWS")
        AutenticarWindows = "{""sucesso"": false, ""mensagem"": ""Usuário ou senha inválidos""}"
        Exit Function
    End If

    If dadosUsuario("BloqueadoAtivo") = 1 Then
        Call AcessoFalhou(usuario, "WINDOWS")
        AutenticarWindows = "{""sucesso"": false, ""mensagem"": ""Usuário bloqueado""}"
        Exit Function
    End If
    
    Call SalvarDadosSessao(dadosUsuario)
    Call RegistrarLog("login", "WINDOWS", dadosUsuario("LOGIN"), dadosUsuario("CODIGO"))
    AutenticarWindows = "{""sucesso"": true, ""mensagem"": ""Login realizado com sucesso!""}"
End Function


' Autentica com Active Directory
Function AutenticarAD() 
    usuario = SanitizeSQL(Request.Form("usuario"))
    senha = SanitizeSQL(Request.Form("senha"))
    dominio = SanitizeSQL(Request.Form("dominio"))

    If usuario = "" Or senha = "" Then
        Call AcessoFalhou(usuario, "AD")
        AutenticarAD = "{""sucesso"": false, ""mensagem"": ""Usuário ou senha não informados!""}"
        Exit Function
    End If

    If dominio = "" Then
        Call AcessoFalhou(usuario, "AD")
        AutenticarAD = "{""sucesso"": false, ""mensagem"": ""Dominio nao informado!""}"
        Exit Function
    End If

    ' Busca usuário no AD
    encontrouUsuarioAD = VerificarUsuarioAD(usuario, senha, dominio)
    If Not encontrouUsuarioAD Then
        Call AcessoFalhou(usuario, "AD")
        AutenticarAD = "{""sucesso"": false, ""mensagem"": ""Usuário ou senha inválidos""}"
        Exit Function
    End If

    ' Busca usuário no sistema
    Set dadosUsuario = VerificarUsuarioSistema("AD", usuario, "")
    If Not dadosUsuario("sucesso") Then
        Call AcessoFalhou(usuario, "AD")
        AutenticarAD = "{""sucesso"": false, ""mensagem"": ""Usuário ou senha inválidos""}"
        Exit Function
    End If

    If dadosUsuario("BloqueadoAtivo") = 1 Then
        Call AcessoFalhou(usuario, "AD")
        AutenticarAD = "{""sucesso"": false, ""mensagem"": ""Usuário bloqueado""}"
        Exit Function
    End If

    Call SalvarDadosSessao(dadosUsuario)
    Call RegistrarLog("login", "AD", dadosUsuario("LOGIN"), dadosUsuario("CODIGO"))
    AutenticarAD = "{""sucesso"": true, ""mensagem"": ""Login realizado com sucesso!""}"
End Function

' Autentica com credenciais do sistema
Function AutenticarSistema() 
    usuario = SanitizeSQL(Request.Form("usuario"))
    senha = SanitizeSQL(Request.Form("senha"))

    If usuario = "" Or senha = "" Then
        Call AcessoFalhou(usuario, "SISTEMA")
        AutenticarSistema = "{""sucesso"": false, ""mensagem"": ""Usuário ou senha não informados!""}"
        Exit Function
    End If

    ' Busca usuário no sistema
    Set dadosUsuario = VerificarUsuarioSistema("SISTEMA", usuario, senha)
    If Not dadosUsuario("sucesso") Then
        Call AcessoFalhou(usuario, "SISTEMA")
        AutenticarSistema = "{""sucesso"": false, ""mensagem"": ""Usuário ou senha incorretos!""}"
        Exit Function
    End If

    If dadosUsuario("BloqueadoAtivo") = 1 Then
        Call AcessoFalhou(usuario, "SISTEMA")
        AutenticarSistema = "{""sucesso"": false, ""mensagem"": ""Usuário bloqueado""}"
        Exit Function
    End If

    Call SalvarDadosSessao(dadosUsuario)
    Call RegistrarLog("login", "SISTEMA", dadosUsuario("LOGIN"), dadosUsuario("CODIGO"))
    AutenticarSistema = "{""sucesso"": true, ""mensagem"": ""Login realizado com sucesso!""}"
End Function


'----------------------------------------------
'                 UTILITÁRIOS
'----------------------------------------------

Function VerificarUsuarioSistema(metodoAutenticacao, usuario, senha)
    SET dados = Server.CreateObject("Scripting.Dictionary")

    SQL = "SELECT DISTINCT "
    SQL = SQL & "   CRM_DTI.CRM_COD AS CODUSER, "
    SQL = SQL & "   CRM_DTI.Tentativas AS QuantidadeTentativas, "
    SQL = SQL & "   CASE WHEN CRM_DTI.CRM_VALNIVEL = 0 THEN CRM_R_E_C_N_O ELSE CRM_DTI.CRM_COD END AS IdUsuario, "
    SQL = SQL & "   SUBSTRING(CRM_DTI.CRM_DATEXP,7,2)+'/'+SUBSTRING(CRM_DTI.CRM_DATEXP,5,2)+'/'+SUBSTRING(CRM_DTI.CRM_DATEXP,1,4) AS DATEXPIRASENHA, "
    SQL = SQL & "   CRM_DTI.CRM_EMAILPRIMARIO AS EMAIL, "
    SQL = SQL & "   CRM_DTI.CRM_R_E_C_N_O AS CODIGO, "
    SQL = SQL & "   CRM_DTI.CodUsuProtheus AS CODUSUPROTHEUS, "
    SQL = SQL & "   ISNULL(CRM_DTI.HP_ADMHP,'') AS Administrador, "
    SQL = SQL & "   CRM_DTI.CRM_DATHORACES AS HORAACESSO, "
    SQL = SQL & "   CRM_DTI.CRM_DATACES AS DATAACESSO, "
    SQL = SQL & "   CRM_DTI.CRM_ACESSOS AS QTDACESSOS, "
    SQL = SQL & "   CRM_DTI.CRM_TIPO AS TIPO, "
    SQL = SQL & "   CRM_DTI.CRM_VALPESOMENOR AS PESOMENOR, "
    SQL = SQL & "   CRM_DTI.CRM_VALPRECOMENOR AS PRECOMENOR, "
    SQL = SQL & "   CRM_DTI.CRM_VALINFCAD AS VALINFCAD, "
    SQL = SQL & "   CRM_DTI.CRM_FININFCAD AS FININFCAD, "
    SQL = SQL & "   CRM_DTI.CRM_VERTODASFNCS AS VERTODASFNCS, "
    SQL = SQL & "   CRM_DTI.CRM_NOME AS NOMEUSER, "
    SQL = SQL & "   ISNULL(CRM_DTI.VER_FILIAL, 'N') AS VER_FILIAL, "
    SQL = SQL & "   CRM_DTI.EMP_CODIGO AS FILIAL, "
    SQL = SQL & "   CRM_DTI.RA_SEXO AS SEXOUSER, "
    SQL = SQL & "   CRM_DTI.RA_NASC AS DATNASC, "
    SQL = SQL & "   CRM_DTI.CRM_LOGIN AS LOGIN, "
    SQL = SQL & "   CRM_DTI.CRM_USERAD AS USERAD, "
    SQL = SQL & "   CRM_DTI.CRM_EMAILPRIMARIO AS EMAILPRIMARIO, "
    SQL = SQL & "   CRM_DTI.CRM_EMAILSECUNDARIO AS EMAILSECUNDARIO, "
    SQL = SQL & "   CRM_DTI.CRM_VALNIVEL AS NIVEL, "
    SQL = SQL & "   CRM_DTI.CRM_ALTERAPO AS ALTERAPO, "
    SQL = SQL & "   CRM_DTI.AtivaUsuario, "
    SQL = SQL & "   CRM_DTI.NivelSPD AS NivelSPD, "
    SQL = SQL & "   CRM_DTI.CRM_ALTERAPAPEL AS ALTERAPAPEL, "
    SQL = SQL & "   CRM_DTI.CRM_NIVELALTERAPEDIDO AS NIVELALTERAPEDIDO, "
    SQL = SQL & "   CRM_DTI.NivelSolicitacaoAlteracaoPedido AS NIVELSOLICITACAOALTERACAOPEDIDO, " 
    SQL = SQL & "   CRM_DTI.CRM_VALARTNIVEL AS VALARTNIVEL, "
    SQL = SQL & "   CRM_DTI.CRM_PERMISSAO AS PERMISSAO, "
    SQL = SQL & "   CRM_TRATAMENTO.TRA_COD AS CODTRATAMENTO, "
    SQL = SQL & "   CRM_TRATAMENTO.TRA_DESC AS TRATAMENTO, "
    SQL = SQL & "   CRM_CATEGORIA.CAT_NOME AS CATEGORIA, "
    SQL = SQL & "   CRM_DTI.CRM_ASSINATURA AS ASSINATURA, "
    SQL = SQL & "   CRM_DTI.CAT_CODIGO AS CATCODIGO, "
    SQL = SQL & "   ISNULL(TROCARSENHA,'N') AS TROCARSENHA, "
    SQL = SQL & "   CRM_DTI.CRM_DPTO AS DPTO, "
    SQL = SQL & "   CRM_DTI.COD_FORNE_SIGA AS FORNE_SIGA, "
    SQL = SQL & "   CRM_DTI.LOJA_FORNE_SIGA AS LOJA_SIGA, "
    SQL = SQL & "   ISNULL(CRM_DTI.Bloqueado,0) AS Bloqueado, "
    SQL = SQL & "   ISNULL(CRM_DTI.Desbloqueio,'') DataDesbloqueio, "
    SQL = SQL & "   ISNULL(CASE WHEN GETDATE() >= Desbloqueio THEN 1 ELSE 0 END,0) Desbloqueado, "
    SQL = SQL & "   CASE WHEN ISNULL(CRM_DTI.Bloqueado,0) = 0 THEN 0 WHEN CRM_DTI.Desbloqueio IS NULL THEN 1 WHEN GETDATE() < CRM_DTI.Desbloqueio THEN 1 ELSE 0 END AS BloqueadoAtivo, "
    SQL = SQL & "   ISNULL(ValidaPreCadastroProduto,0) ValidaPreCadastroProduto, "
    SQL = SQL & "   ISNULL(AprovaOrcamentoPreCadastroProduto,0) AprovaOrcamentoPreCadastroProduto, "
    SQL = SQL & "   ISNULL(ValidaOrcamentoCusto,0) ValidaOrcamentoCusto, "
    SQL = SQL & "   ISNULL(NivelCEDAG,0) NivelCEDAG, "
    SQL = SQL & "   ISNULL(NivelE2OPEN,0) NivelE2OPEN, "
    SQL = SQL & "   ISNULL(NivelControlesInternos,0) NivelControlesInternos, "
    SQL = SQL & "   ISNULL(NivelControladoria,0) NivelControladoria, "
    SQL = SQL & "   ISNULL(SAPLaboratorista,0) SAPLaboratorista, "
    SQL = SQL & "   ISNULL(SAPClicherista,0) SAPClicherista, "
    SQL = SQL & "   ISNULL(NivelFinanceiro,0) NivelFinanceiro, "
    SQL = SQL & "   ISNULL(NivelFiscal,0) NivelFiscal, "
    SQL = SQL & "   ISNULL(ValidadoresCredito.Nivel, '') NivelValidadorCredito "
    SQL = SQL & "FROM CRM_DTI "
    SQL = SQL & "INNER JOIN CRM_TRATAMENTO ON CRM_DTI.TRA_COD = CRM_TRATAMENTO.TRA_COD "
    SQL = SQL & "INNER JOIN CRM_CATEGORIA ON CRM_DTI.CAT_CODIGO = CRM_CATEGORIA.CAT_CODIGO "
    SQL = SQL & "LEFT JOIN CORPORATIVOHMG..ValidadoresCredito ON CRM_DTI.CRM_R_E_C_N_O = ValidadoresCredito.IdValidador AND ISNULL(ValidadoresCredito.QuemExcluiu, 0) = 0"
    
    If metodoAutenticacao = "WINDOWS" OR metodoAutenticacao = "AD" Then 
        SQL = SQL & "WHERE CRM_DTI.CRM_USERAD = '" & usuario & "' "
    Else
        SQL = SQL & "WHERE ( CRM_DTI.CRM_LOGIN = '" & usuario & "' "
        SQL = SQL & "AND CRM_DTI.Senha = HASHBYTES('SHA1', '" & senha & "')) "
    End If

    SQL = SQL & "AND CRM_DTI.CRM_D_E_L_E_T <> '*' "

    SET rsUsuario = Conecta("PIR-HELPDESK", SQL)
    If not isObject(rsUsuario) OR (rsUsuario IS NOTHING) Then 
        dados("sucesso") = False
        Set VerificarUsuarioSistema = dados
        Exit Function
    End If 

    If rsUsuario.EOF Then
        dados("sucesso") = False
        Set VerificarUsuarioSistema = dados
        Exit Function
    Else
        dados("sucesso") = True
        dados("QuantidadeTentativas") = rsUsuario("QuantidadeTentativas")
        dados("CODUSER") = rsUsuario("CODUSER")
        dados("IdUsuario") = rsUsuario("IdUsuario")
        dados("DATEXPIRASENHA") = rsUsuario("DATEXPIRASENHA")
        dados("EMAIL") = rsUsuario("EMAIL")
        dados("CODIGO") = rsUsuario("CODIGO")
        dados("CODUSUPROTHEUS") = rsUsuario("CODUSUPROTHEUS")
        dados("HORAACESSO") = rsUsuario("HORAACESSO")
        dados("DATAACESSO") = rsUsuario("DATAACESSO")
        dados("QTDACESSOS") = rsUsuario("QTDACESSOS")
        dados("TIPO") = rsUsuario("TIPO")
        dados("PESOMENOR") = rsUsuario("PESOMENOR")
        dados("PRECOMENOR") = rsUsuario("PRECOMENOR")
        dados("VALINFCAD") = rsUsuario("VALINFCAD")
        dados("FININFCAD") = rsUsuario("FININFCAD")
        dados("VERTODASFNCS") = rsUsuario("VERTODASFNCS")
        dados("NOMEUSER") = rsUsuario("NOMEUSER")
        dados("VER_FILIAL") = rsUsuario("VER_FILIAL")
        dados("FILIAL") = rsUsuario("FILIAL")
        dados("SEXOUSER") = rsUsuario("SEXOUSER")
        dados("DATNASC") = rsUsuario("DATNASC")
        dados("LOGIN") = rsUsuario("LOGIN")
        dados("USERAD") = rsUsuario("USERAD")
        dados("EMAILPRIMARIO") = rsUsuario("EMAILPRIMARIO")
        dados("EMAILSECUNDARIO") = rsUsuario("EMAILSECUNDARIO")
        dados("NIVEL") = rsUsuario("NIVEL")
        dados("ALTERAPO") = rsUsuario("ALTERAPO")
        dados("AtivaUsuario") = rsUsuario("AtivaUsuario")
        dados("NivelSPD") = rsUsuario("NivelSPD")
        dados("ALTERAPAPEL") = rsUsuario("ALTERAPAPEL")
        dados("NIVELALTERAPEDIDO") = rsUsuario("NIVELALTERAPEDIDO")
        dados("NIVELSOLICITACAOALTERACAOPEDIDO") = rsUsuario("NIVELSOLICITACAOALTERACAOPEDIDO")
        dados("VALARTNIVEL") = rsUsuario("VALARTNIVEL")
        dados("PERMISSAO") = rsUsuario("PERMISSAO")
        dados("CODTRATAMENTO") = rsUsuario("CODTRATAMENTO")
        dados("TRATAMENTO") = rsUsuario("TRATAMENTO")
        dados("CATEGORIA") = rsUsuario("CATEGORIA")
        dados("ASSINATURA") = rsUsuario("ASSINATURA")
        dados("CATCODIGO") = rsUsuario("CATCODIGO")
        dados("TROCARSENHA") = rsUsuario("TROCARSENHA")
        dados("DPTO") = rsUsuario("DPTO")
        dados("FORNE_SIGA") = rsUsuario("FORNE_SIGA")
        dados("LOJA_SIGA") = rsUsuario("LOJA_SIGA")
        dados("Bloqueado") = rsUsuario("Bloqueado")
        dados("Desbloqueado") = rsUsuario("Desbloqueado")
        dados("DataDesbloqueio") = rsUsuario("DataDesbloqueio")
        dados("BloqueadoAtivo") = rsUsuario("BloqueadoAtivo")
        dados("ValidaPreCadastroProduto") = rsUsuario("ValidaPreCadastroProduto")
        dados("AprovaOrcamentoPreCadastroProduto") = rsUsuario("AprovaOrcamentoPreCadastroProduto")
        dados("ValidaOrcamentoCusto") = rsUsuario("ValidaOrcamentoCusto")
        dados("NivelCEDAG") = rsUsuario("NivelCEDAG")
        dados("NivelE2OPEN") = rsUsuario("NivelE2OPEN")
        dados("NivelControlesInternos") = rsUsuario("NivelControlesInternos")
        dados("NivelControladoria") = rsUsuario("NivelControladoria")
        dados("SAPLaboratorista") = rsUsuario("SAPLaboratorista")
        dados("SAPClicherista") = rsUsuario("SAPClicherista")
        dados("NivelFinanceiro") = rsUsuario("NivelFinanceiro")
        dados("NivelFiscal") = rsUsuario("NivelFiscal")
        dados("NivelValidadorCredito") = rsUsuario("NivelValidadorCredito")
    End If

    rsUsuario.Close()
    SET rsUsuario = nothing

    Set VerificarUsuarioSistema = dados
End Function

Function VerificarUsuarioAD(usuario, senha, dominio)
    controladorDominio = ""
    porta = 389 ' Porta padrão para LDAP
    usuario = trim( "" & usuario)

    ' Define o controlador de domínio e caminho base no Active Directory
    If dominio = "GROUP" Then
        'Legado SW
        usuario = "GROUP\" & usuario
        controladorDominio = "GROUP.WAN"
        baseAD = "DC=GROUP,DC=WAN" 
    Else
        'Westrock
        usuario = "WESTROCK\" & usuario
        controladorDominio = "WESTROCK.COM"
        baseAD = "DC=WESTROCK,DC=COM" 
    End If

    encontrouUsuario = False 
    
    SET oCon = Server.CreateObject("ADODB.Connection")
    oCon.Provider = "ADsDSOObject" ' Define o provedor para o Active Directory
    oCon.Properties("User ID") = usuario ' Define o ID do usuário para a conexão
    oCon.Properties("Password") = senha ' Define a senha para a conexão
    oCon.Open "ADProvider" ' Abre a conexão com o Active Directory

    SET oCmd = Server.CreateObject("ADODB.Command")
    SET oCmd.ActiveConnection = oCon 
    propriedade = "mail" ' Propriedade que queremos recuperar (endereço de e-mail)

    ' Monta a string de comando LDAP
    cmd = "<" & "LDAP://" & controladorDominio & ":" & porta & "/" & baseAD & ">;" & ";" & propriedade & ";subtree"
    oCmd.CommandText = cmd 
    oCmd.Properties("Page Size") = 100 ' Define o tamanho da página para a consulta
    
    SET rsUsuarioAD = oCmd.Execute

    If Err.Number = 0 AND isObject(rsUsuarioAD) Then
        encontrouUsuario = true 
        rsUsuarioAD.Close() 
        SET rsUsuarioAD = NOTHING 
    End If

    set oCmd = NOTHING 

    VerificarUsuarioAD = encontrouUsuario
End Function

Sub AcessoFalhou(usuario, metodoAutenticacao)  
    SQL = SQL & "UPDATE CRM_DTI SET Tentativas = ISNULL(Tentativas, 0) + 1 "
    SQL = SQL & "WHERE (CRM_LOGIN = '" & usuario & "' OR CRM_USERAD = '" & usuario & "') AND CRM_D_E_L_E_T <> '*'"

    On Error Resume Next
    SET rsUpdateTentativas = Conecta("PIR-HELPDESK", SQL)

    If Err.Number = 0 AND isObject(rsUpdateTentativas) Then
        rsUpdateTentativas.Close()
        SET rsUpdateTentativas = NOTHING
    End If
    On Error Goto 0
    
    totalTentativas = 0
    emailUsuario = ""

    SQL = "SELECT TOP 1 Tentativas as totalTentativas, CRM_EMAILPRIMARIO as email "
    SQL = SQL & "FROM CRM_DTI "
    SQL = SQL & "WHERE (CRM_LOGIN = '" & usuario & "' OR CRM_USERAD = '" & usuario & "') AND CRM_D_E_L_E_T <> '*'"
    
    On Error Resume Next
    SET rsTotalTentativas = Conecta("PIR-HELPDESK", SQL)

    If Err.Number = 0 AND isObject(rsTotalTentativas) Then
        totalTentativas = rsTotalTentativas("totalTentativas")
        emailUsuario = rsTotalTentativas("email")
        rsTotalTentativas.Close()
        SET rsTotalTentativas = NOTHING
    End If

    On Error Goto 0

    If totalTentativas = 5 Then
        SQL = SQL & "UPDATE CRM_DTI SET "
        SQL = SQL & "   Bloqueado = 1, "
        SQL = SQL & "   Desbloqueio = Dateadd(hour,1,getdate()), "
        SQL = SQL & "   IpBloqueou = '" & Request.ServerVariables("REMOTE_ADDR") & "' "
        SQL = SQL & "WHERE CRM_LOGIN = '" & usuario & "' OR CRM_USERAD = '" & usuario & "'"
        SQL = SQL & "AND CRM_D_E_L_E_T <> '*' " 
        
        On Error Resume Next
        SET rsTentativas = Conecta("PIR-HELPDESK", SQL)

        If Err.Number = 0 AND isObject(rsTentativas) Then
            rsTentativas.Close()
            SET rsTentativas = NOTHING
        End If

        On Error Goto 0

        If emailUsuario <> "" Then
            Call EnviarEmailBloqueio(emailUsuario)
        End If
    End If

    Call RegistrarLog("tentativa_login", metodoAutenticacao, usuario, "")
End Sub

Sub SalvarDadosSessao(dadosUsuario)
    SESSION("Adm")              = iif( trim(dadosUsuario("Administrador")) = "S", 1, 0 )
    SESSION("IdUsuario")        = dadosUsuario("IdUsuario")
    SESSION("Bloqueao")         = dadosUsuario("Bloqueado")
    SESSION("USR")              = dadosUsuario("LOGIN")
    SESSION("PSW")              = PWD
    SESSION("CODUSER")          = dadosUsuario("CODUSER")
    SESSION("CODIGO")           = dadosUsuario("CODIGO")
    SESSION("FILIAL")           = dadosUsuario("FILIAL")
    SESSION("VER_FILIAL")       = dadosUsuario("VER_FILIAL")
    SESSION("CODTRATAMENTO")    = dadosUsuario("CODTRATAMENTO")
    SESSION("TRATAMENTO")       = dadosUsuario("TRATAMENTO")
    SESSION("NOME")             = dadosUsuario("NOMEUSER")
    SESSION("DATACESSO")        = dadosUsuario("DATAACESSO")
    SESSION("HORACESSO")        = dadosUsuario("HORAACESSO")
    SESSION("ACESSOS")          = dadosUsuario("QTDACESSOS")
    SESSION("PESOMENOR")        = dadosUsuario("PESOMENOR")
    SESSION("PRECOMENOR")       = dadosUsuario("PRECOMENOR")
    SESSION("VALINFCAD")        = dadosUsuario("VALINFCAD")
    SESSION("FININFCAD")        = dadosUsuario("FININFCAD")
    SESSION("VERTODASFNCS")     = Trim(dadosUsuario("VERTODASFNCS"))
    SESSION("EMAIL")            = dadosUsuario("EMAIL")
    SESSION("CATCODIGO")        = dadosUsuario("CATCODIGO")
    SESSION("NIVEL")            = dadosUsuario("NIVEL")
    SESSION("ALTERAPO")         = dadosUsuario("ALTERAPO")
    SESSION("AtivaUsuario")     = dadosUsuario("AtivaUsuario")
    SESSION("NIVELSPD")         = dadosUsuario("NivelSPD")
    SESSION("ALTERAPAPEL")      = dadosUsuario("ALTERAPAPEL")
    SESSION("NIVELALTERAPEDIDO") = dadosUsuario("NIVELALTERAPEDIDO")
    SESSION("NIVELSOLICITACAOALTERACAOPEDIDO") = dadosUsuario("NIVELSOLICITACAOALTERACAOPEDIDO")
    SESSION("VALARTNIVEL")      = dadosUsuario("VALARTNIVEL")
    SESSION("CATEGORIA")        = dadosUsuario("CATCODIGO")
    SESSION("PERMISSAO")        = dadosUsuario("PERMISSAO")
    SESSION("DPTO")             = dadosUsuario("DPTO")
    SESSION("CODSIGA")          = dadosUsuario("CODUSER")
    SESSION("ASSINATURA")       = dadosUsuario("ASSINATURA")
    SESSION("TROCARSENHA")      = dadosUsuario("TROCARSENHA")
    SESSION("DATEXPIRASENHA")   = dadosUsuario("DATEXPIRASENHA")
    SESSION("FORNECEDOR")       = dadosUsuario("FORNE_SIGA")
    SESSION("LOJA")             = dadosUsuario("LOJA_SIGA")
    SESSION("TIPO")             = dadosUsuario("TIPO")
    SESSION("CODUSUPROTHEUS")   = dadosUsuario("CODUSUPROTHEUS")
    SESSION("ValidaPreCadastroProduto") = dadosUsuario("ValidaPreCadastroProduto")
    SESSION("AprovaOrcamentoPreCadastroProduto")  = dadosUsuario("AprovaOrcamentoPreCadastroProduto")
    SESSION("ValidaOrcamentoCusto") = dadosUsuario("ValidaOrcamentoCusto")
    SESSION("CEDAG")            = dadosUsuario("NivelCEDAG")
    SESSION("E2OPEN")           = dadosUsuario("NivelE2OPEN")
    SESSION("ControlesInternos")= dadosUsuario("NivelControlesInternos")
    SESSION("Controladoria")    = dadosUsuario("NivelControladoria")
    SESSION("NivelFinanceiro")  = dadosUsuario("NivelFinanceiro")
    SESSION("NivelFiscal")      = dadosUsuario("NivelFiscal")
    SESSION("NivelValidadorCredito") = dadosUsuario("NivelValidadorCredito")

    'GERANDO COOKIES PARA INTEGRACAO COM OUTROS SISTEMAS CRIPTOGRAFADO EM BASE64
    response.cookies("INPANET_CODUSER")           = base64_encode(trim(dadosUsuario("CODUSER")))
    response.cookies("INPANET_CODUSER").Domain    = Request.ServerVariables("server_name")
    response.cookies("INPANET_CODIGO")            = base64_encode(trim(dadosUsuario("CODIGO")))
    response.cookies("INPANET_CODIGO").Domain     = Request.ServerVariables("server_name")
    response.cookies("INPANET_USR")               = base64_encode(trim(dadosUsuario("LOGIN")))
    response.cookies("INPANET_USR").Domain        = Request.ServerVariables("server_name")
    response.cookies("INPANET_PWD")               = base64_encode(PWD)
    response.cookies("INPANET_PWD").Domain        = Request.ServerVariables("server_name")
    response.cookies("INPANET_NOMEUSR")           = base64_encode(trim(dadosUsuario("NOMEUSER")))
    response.cookies("INPANET_NOMEUSR").Domain    = Request.ServerVariables("server_name")
    response.cookies("INPANET_EMAIL")             = base64_encode(trim(dadosUsuario("EMAIL")))
    response.cookies("INPANET_EMAIL").Domain      = Request.ServerVariables("server_name")
    response.cookies("INPANET_USRNIVEL")          = base64_encode(trim(dadosUsuario("NIVEL")))
    response.cookies("INPANET_USRNIVEL").Domain   = Request.ServerVariables("server_name")
    response.cookies("INPANET_VER_FILIAL")        = base64_encode(trim(dadosUsuario("VER_FILIAL")))
    response.cookies("INPANET_VER_FILIAL").Domain = Request.ServerVariables("server_name")
    response.cookies("INPANET_FILIAL")            = base64_encode(trim(dadosUsuario("FILIAL")))
    response.cookies("INPANET_FILIAL").Domain     = Request.ServerVariables("server_name")
    response.cookies("INPANET_NOME")              = trim(dadosUsuario("LOGIN"))   
    response.cookies("INPANET_NOME").Domain       = Request.ServerVariables("server_name")
    response.cookies("FILIAL_USUARIO")            = trim(dadosUsuario("FILIAL"))
    response.cookies("FILIAL_USUARIO").Domain     = Request.ServerVariables("server_name")

    response.cookies("INPANET_CRM_LOGIN")            = trim(dadosUsuario("LOGIN"))
    response.cookies("INPANET_CRM_LOGIN").Domain     = Request.ServerVariables("server_name")

    response.cookies("INPANET_CRM_COD")              = trim(dadosUsuario("CODUSER"))
    response.cookies("INPANET_CRM_COD").Domain       = Request.ServerVariables("server_name")

    response.cookies("INPANET_CRM_NOME")             = trim(dadosUsuario("NOMEUSER"))
    response.cookies("INPANET_CRM_NOME").Domain      = Request.ServerVariables("server_name")

    response.cookies("INPANET_CRM_VALNIVEL")         = trim(dadosUsuario("NIVEL"))
    response.cookies("INPANET_CRM_VALNIVEL").Domain  = Request.ServerVariables("server_name")

    response.cookies("INPANET_CRM_R_E_C_N_O")        = trim(dadosUsuario("CODIGO"))
    response.cookies("INPANET_CRM_R_E_C_N_O").Domain = Request.ServerVariables("server_name")

    response.cookies("INPANET_EMP_CODIGO")           = trim(dadosUsuario("FILIAL"))
    response.cookies("INPANET_EMP_CODIGO").Domain    = Request.ServerVariables("server_name")

    ' Reseta tentativas
    SQL = SQL & "BEGIN TRANSACTION " 
    SQL = SQL & "UPDATE CRM_DTI SET " 
    SQL = SQL & "Tentativas = 0, " 
    SQL = SQL & "CRM_DTI.CRM_DATACES = CONVERT(VARCHAR(8),GETDATE(),112) " 
    SQL = SQL & "WHERE CRM_USERAD = '" & dadosUsuario("USERAD") & "'" 
    SQL = SQL & "OR CRM_LOGIN = '" & dadosUsuario("LOGIN") & "'"
    SQL = SQL & "IF @@ERROR > 0 " 
    SQL = SQL & "   ROLLBACK TRANSACTION " 
    SQL = SQL & "ELSE " 
    SQL = SQL & "   COMMIT TRANSACTION "

    On Error Resume Next
    SET rsTentativas = Conecta("PIR-HELPDESK", SQL)
    
    If Err.Number = 0 AND IsObject(rsTentativas) Then
        rsTentativas.Close()
        SET rsTentativas = NOTHING
    End If
    On Error Goto 0
End Sub

Sub EnviarEmailBloqueio(email)
    MSG = "" &_
    "<h4 style='font-family:arial'> " &_
        "<strong>Aten&ccedil;&atilde;o:</strong> Seu usu&aacute;rio, do SWnet, acaba de ser Bloqueado por exceder o n&uacute;mero m&aacute;ximo de tentativas de logon inv&aacute;lidas." &_
    "</h4>" &_
    "<p style='font-family:arial'>" &_
        "O desbloqueio ocorrer&aacute; de forma autom&aacute;tica ap&oacute;s o primeiro acesso, que exceda o hor&aacute;rio do bloqueio em uma hora.<br>" &_
        "Por motivo de seguran&ccedil;a tamb&eacute;m ser&aacute; solicitada a troca da senha." &_
    "</p>" &_
    "<p style='font-family:arial'>" &_
        "Atenciosamente;<br />" &_
        "Smurfit Westrock - Depto. de Tecnologia da informa&ccedil;&atilde;o" &_
    "</p>"

    Call EnviaEmail(email, "", "", "", "noreply.brasil@smurfitkappa.com.br", MSG, "Bloqueio de usuario SWnet")
End Sub

Sub RegistrarLog(acao, metodoAutenticacao, login, idUsuario)
    Select Case LCase(acao)
        Case "login"
            acaoLog = "Login bem sucedido"
            descricaoLog = "Usuario " & login & " realizou login com sucesso com o metodo " & metodoAutenticacao
            descricaoLog = descricaoLog & " - IP: " & Request.ServerVariables("REMOTE_ADDR")
        Case "tentativa_login"
            acaoLog = "Login mal sucedido"
            descricaoLog = "Tentativa de login mal sucedida para o usuario " & login & " com o metodo " & metodoAutenticacao
            descricaoLog = descricaoLog & " - IP: " & Request.ServerVariables("REMOTE_ADDR")
    End Select
    
    SQL = "INSERT INTO LogAtividade (IdUsuarioAcao, IdUsuarioAlvo, Rotina, Sistema, Acao, Descricao, DataCriacao) VALUES ("
    SQL = SQL & Iif(idUsuario <> "", idUsuario, "NULL") & ", "
    SQL = SQL & "NULL, "
    SQL = SQL & "'LOGIN', "
    SQL = SQL & "'SWNET', "
    SQL = SQL & "'" & acaoLog & "', "
    SQL = SQL & "'" & descricaoLog & "', "
    SQL = SQL & "GETDATE())"
    
    On Error Resume Next
    SET rsLog = Conecta("PIR-CORPORATIVO", SQL)
    
    If Err.Number = 0 AND IsObject(rsLog) Then
        rsLog.Close()
        SET rsLog = NOTHING
    End If
    On Error Goto 0
End Sub
%>