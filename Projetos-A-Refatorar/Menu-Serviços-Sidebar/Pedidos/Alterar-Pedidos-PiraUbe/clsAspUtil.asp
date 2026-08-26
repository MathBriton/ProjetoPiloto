<%
'#################################################################################################
'##                                                                                             ##
'##    CLASSE COM DIVERSAS FUNÇÕES ÚTEIS                                                        ##
'##                                                                                             ##
'##     ÍNDICE:                                                                                 ##
'##     01 - Criptografia(Texto, Tipo) |  Criptografa e Descriptografia simples                 ##
'##     02 - EncriptaStr(Texto)        |  Criptografa strings                                   ##
'##     03 - DecriptaStr(Texto)        |  DesCriptografa strings                                ##
'##     04 - GerarSenha(MaxNum,Tipo)   |  Gera Senhas aleatórias                                ##
'##     05 - getUsuariosADS(strGrupo)  |  Lista Usuários de um grupo do Active Direcotry        ##
'##     06 - RemoveAcentos(strTexto)   |  Função que remove acentos de textos                   ##
'##     07 - DelHTMLTags(strHTML)      |  Função que remove tags HTML de textos                 ##
'##                                                                                             ##
'#################################################################################################

Class AspUtil
 
'.................................................................................................
  '<----- 01 - FUNÇÃO QUE CRIPTOGRAFA UMA STRING SIMPLES
          'Função de criptofgrafia simples para senhas
		  'Parâmetros:
		  'Texto - String a ser criprografada ou descriptografada
		  'Tipo  - Boolean : True criptografa e false descriptografa
  Function Criptografia(Texto, Tipo)
    'Declara as variáveis usadas dentro da função
	Dim i, a, x, y, Constant, Constant2
	'Monta as constantes
	Constant  = "0,1,2,3,4,5,6,7,8,9,a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
	Constant2 = StrReverse(Constant)
    
	'Varre toda a string enviada como parâmetro
	For i = 1 To Len(Texto)
      'Verifica se é True: Criptografa ou False: Descriptografa
	  If Tipo = True Then
	    'Criptografa
		x = Mid(Texto, i, 1)
		For a = 1 to Len(Constant)
		  If strComp(x,Mid(Constant, a, 1)) = 0 Then
		    y = Y & Mid(Constant2, a, 1)
		  End If
		Next
	  Else
	    'Descriptografa
		x = Mid(Texto, i, 1)
		For a = 1 to Len(Constant2)
		  If strComp(x,Mid(Constant2, a, 1)) = 0 Then
		    y = Y & Mid(Constant, a, 1)
		  End If
		Next	  
	  End If
	Next
    Criptografia = y
	'-----
    'Crip = Criptografia("Marcos", true)
    'Response.Write Crip & "<br>"
    'Response.Write Criptografia(Crip, false)
  End Function
 '................................................................................................. 

  '.................................................................................................
  '<----- 01 - FUNÇÃO QUE CRIPTOGRAFA UMA STRING
          'recebe uma string como parâmetro
		  'Ideal para senhas
  Function EncriptaStr(Texto)
    Const EncC1 = 109
    Const EncC2 = 191
    Const EncKey = 161

    Dim TempStr, TempResult, TempNum, TempChar
    Dim TempKey
    Dim i
    TempStr = Texto
    TempResult = ""
    TempKey = ((EncKey * EncC1) + EncC2) Mod 65536
    For i = 1 To Len(TempStr)
        TempNum = (Asc(Mid(TempStr, i, 1)) Xor (AuxShr(TempKey, 8))) Mod 256
        TempChar = Chr(TempNum)
        TempKey = (((Asc(TempChar) + TempKey) * EncC1) + EncC2) Mod 65536
        TempResult = TempResult & TempChar
    Next
    EncriptaStr = TempResult
    
	'-----
	'CÓDIGO PARA TESTAR A FUNÇÃO
    'VarTeste = EncriptaStr("1 ,2 ,3 testando...")
    '     CheckError (VarTeste& "<br>")
    '     CheckError (DecriptaStr(VarTeste))
  End Function
  '.................................................................................................

  '.................................................................................................
  '<----- 02 - FUNÇÃO PARA DESCRIPTOGRAFIA DE STRINGS
          'recebe uma string criptografada como parâmetro
		  'Ideal para senhas
  Function DecriptaStr(Texto)
    Dim TempStr, TempResult, TempNum, TempChar
    Dim TempKey
    Dim i
    TempStr = Texto
    TempResult = ""
    TempKey = ((EncKey * EncC1) + EncC2) Mod 65536
    For i = 1 To Len(TempStr)
        TempNum = (Asc(Mid(TempStr, i, 1)) Xor (AuxShr(TempKey, 8))) Mod 256
        TempChar = Chr(TempNum)
        TempKey = (((Asc(Mid(TempStr, i, 1)) + TempKey) * EncC1) + EncC2) Mod 65536
        TempResult = TempResult & TempChar
    Next
    DecriptaStr = TempResult
  End Function
  
  Private Function AuxShr(Numero, BShr)
    AuxShr = Int(Numero / (2 ^ BShr))
  End Function
  '-----
  'CÓDIGO PARA TESTAR A FUNÇÃO
  'VarTeste = EncriptaStr("1 ,2 ,3 testando...")
  'CheckError (DecriptaStr(VarTeste))  
  '.................................................................................................
 
  '.................................................................................................
  '<----- 03 - FUNÇÃO QUE GERA SENHA ALEATÓRIA
          'MaxNum: Numero máximo de caracteres da senha gerada
          'Tipo      : Tipo de  caracteres utilizados na senha
          '            1 = Somente Numeros
          '            2 = Numero e Letras
          '            3 = Somente Letras
  Function GerarSenha(MaxNum,Tipo)
    Dim var_valores, xArray, Chave, num
	If Tipo = 1 Then var_valores =    "0,1,2,3,4,5,6,7,8,9,"&_
                                    "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
    If Tipo = 2 Then var_valores =    "0,1,2,3,4,5,6,7,8,9"
    
    If Tipo = 3 Then var_valores =    "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
    xArray = Split(var_valores,",")
    Randomize
    Do While Len(chave) < MaxNum
        num = xArray(Int(Ubound(xArray) * Rnd )) 
        chave = chave + num 
    Loop 
    GerarSenha = Trim(chave)
  	
	'-----
	'CÓDIGO PARA TESTAR A FUNÇÃO
    'Response.Write GerarSenha(8,1)
  End Function
  '.................................................................................................

  '.................................................................................................
  '<----- 04 - LISTAS USUÁRIOS DE UM GRUPO DO ACTIVE DIRETORY
          'Nome: getUsuariosADS(strGrupo)
          'Tipo: Funcao Public
          'Sinopse: Retornar a lista de usuário de um grupo no Active Directory
          'Parametros:
          'strGrupo: Grupo na rede
          'Retorno: Collection (Scripting.Dictionary)
  Function getUsuariosADS(strGrupo)
    'Declarações
    Dim objSysInfo
    Dim objUsuarios
    Dim objUsuario
    Dim objColl
        
    'Criação dos objetos
    Set objSysInfo    = Server.CreateObject("ADSystemInfo")
    Set objColl        = Server.CreateObject("Scripting.Dictionary")
    Set objUsuarios    = GetObject("WinNT://" & objSysInfo.DomainDNSName & "/" & strGrupo)
        
    For Each objUsuario In objUsuarios.Members
        'Adicionar o grupo na collection
        objColl.Add objColl.Count+1, objUsuario.Name
    Next
        
    'Retornando a função
    Set getUsuariosADS = objColl
        
    'Destruindo os objetos
    Set objUsuarios    = Nothing
    Set objSysInfo    = Nothing
	
	'-----
	'CÓDIGO PARA TESTAR A FUNÇÃO
    'Dim Grupo    'Grupo da rede
    'Dim objResultado
    'Set objResultado = Server.CreateObject("Scripting.Dictionary")
    'Grupo = "Administradores" '<----- EDITE O GRUPO AQUI
    'Set objResultado = getUsuariosADS(Grupo)
    'For i = 1 To objResultado.Count
      '     CheckError objResultado.Item(i) & "<br>" 
    'Next
    'Set objResultado = Nothing	
  End Function
  '.................................................................................................

  '.................................................................................................
  '<----- FUNÇÃO QUE REMOVE ACENTOS
          'Recebe um texto como parâmetro e devolve o mesmo sem acentos
  Function RemoveAcentos(strTexto)
    Dim strComAcentos
    Dim strSemAcentos
    Dim strResultado
	Dim intCont
    'Conjunto de Caracteres com acentos
    strComAcentos = "ÁÍÓÚÉÄÏÖÜËÀÌÒÙÈÃÕÂÎÔÛÊáíóúéäïöüëàìòùèãõâîôûêÇç&°º"
    'Conjunto de Caracteres sem acentos
    strSemAcentos = "AIOUEAIOUEAIOUEAOAIOUEaioueaioueaioueaoaioueCcE_-"

    intCont = 0
    strResultado = strTexto
    Do While intCont < Len(strComAcentos)
        intCont = intCont + 1
        strResultado = Replace(strResultado, Mid(strComAcentos, intCont, 1), Mid(strSemAcentos, intCont, 1))
    Loop
    RemoveAcentos = strResultado
	
	'-----
	'CÓDIGO PARA TESTAR A FUNÇÃO
	'Response.Write RemoveAcentos("Análise é o bicho de caiçara")
  End Function
  '.................................................................................................

  '.................................................................................................
  '<----- FUNÇÃO QUE REMOVE TAGS HTML DOS TEXTOS
          'Nome: DelHTMLTags(strHTML)
          'Sinopse: Remove todas as tags HTML de uma string
          'Parametros:
          'strHTML: String com as tags HTML a serem retiradas
          'Retorno: String
  Function DelHTMLTags(strHTML)
    Dim objER
    Dim strTexto

    'Configurando o objeto de Expressão Regular
    Set objER            = New RegExp 
    objER.IgnoreCase    = True
    objER.Global        = True
    objER.Pattern        = "<[^>]*>"
    
    'Substituindo as tags encontradas pela expressão
    strTexto            = strHTML
    strTexto            = objER.Replace(strTexto, "")
    
    Set objER            = Nothing

    'Retornando a função
    DelHTMLTags = strTexto
	
	'-----
	'CÓDIGO PARA TESTAR A FUNÇÃO
	'Dim HTML
    'HTML = "<font face='verdana' size='2'>teste</font> da função de retira as <b>TAGS</b> <font color='red'>HTML</font><br>"
    'Texto com TAG
    'Response.Write HTML
    'Texto sem TAG
    'Response.Write DelHTMLTags(HTML)	
  End Function
  '.................................................................................................

  '.................................................................................................
  '<----- FUNÇÃO VERIFICA A EXISTÊNCIA DE UM ARQUIVO NO SERVIDOR
          'Nome: ArquivoExiste(strCaminhoFisico)
          'Sinopse: Verifica se um determinado arquivo existe no disco do servidor
          'Parametros:
          'strCaminhoFisico: Caminho físico completo
          'Retorno: True ou False
  Function ArquivoExiste(strCaminhoFisico)
    Dim objFSO
    'Criando o objeto FSO
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")    

    'Retornando a função
    ArquivoExiste = objFSO.FileExists (strCaminhoFisico)
    
    'Destruindo o objeto FSO
    Set objFSO = Nothing
	
	'-----
	'CÓDIGO PARA TESTAR A FUNÇÃO	
	'If ArquivoExiste("c:\web\site\images\logo.gif") = True Then Response.Write "Existe"
  End Function
  '.................................................................................................
      
  '.................................................................................................
  '<----- FUNÇÃO PARA DEVOLVER O TRATAMENTO DE ERROS
  Function CheckError()
    Select  Case Err.Number
      Case 3
        CheckError " Return sem GoSub"
      Case 5
        CheckError "Chamada de Procedimento Inválida"
      Case 6
        CheckError "Sobrecarga"
      Case 7
        CheckError "Sem Memória"
      Case 9
        CheckError "SubScript fora de área"
      Case 10
        CheckError "Este Array está fixo ou temporariamente travado"
      Case 11
        CheckError "Divisão Por Zero"
      Case 13
        CheckError "Tipos Incompatíveis"
      Case 14
        CheckError "Fora de Espaço de String"
      Case 16
        CheckError "Expressão muito Complexa"
      Case 17
        CheckError "Não pode recuperar a operação"
      Case 18
        CheckError "Interrupção do usuário ocorrida"
      Case 20
        CheckError "Resume Without Error"
      Case 28
        CheckError "Fora de Espaço de Pilha"
      Case 35
        CheckError "Sub ou Function não Definida"
      Case 47
        CheckError "Muitas DLL na aplicação cliente"
      Case 48
        CheckError "Erro carregando DLL"
      Case 49
        CheckError "DLL com problemas de chamada"
      Case 51
        CheckError "Erro Interno"
      Case 52
        CheckError "Nome ou número do arquivo errado"
      Case 53
        CheckError "Arquivo não Encontrado"
      Case 54
        CheckError "Modo de arquivo errado"
      Case 55
        CheckError "Arquivo já está Aberto"
      Case 57
        CheckError "Device I/O Error"
      Case 58
        CheckError "Arquivo jé existente"
      Case 59
        CheckError "Tamanho do registro errado"
      Case 61
        CheckError "Disco Cheio"
      Case 62
        CheckError "Entrada passa do final do arquivo"
      Case 63
        CheckError "Número de registros errados"
      Case 67
        CheckError "Muitos arquivos"
      Case 68
        CheckError "Ferramenta não disponível"
      Case 70
        CheckError "Permissão Negada"
      Case 71
        CheckError "Disco não Preparado"
      Case 74
        CheckError "Não posso renomear com discos diferentes"
      Case 75
        CheckError "Caminho/Arquivos Erro de acesso"
      Case 76
        CheckError "Caminho não encontrado"
      Case 91
        CheckError "Variável de objeto não definida"
      Case 92
        CheckError "Loop For não foi inicializado"
      Case 94
        CheckError "Uso inválido de Null"
      Case 322
        CheckError "Não posso criar Arquivos temporários nescessários"
      Case 325
        CheckError "Formato inválido no arquivo"
      Case 380
        CheckError "Valor da propriedade inválida"
      Case 400
        CheckError "ERRO HTTP 1.1 --- pedido ruim"
      Case 401.1
        CheckError "ERRO HTTP 1.1 --- não autorizado: falha no logon"
      Case 401.2
        CheckError "ERRO HTTP 1.1 --- não autorizado: falha no logon devido a configuração do servidor"
      Case 401.3
        CheckError "ERRO HTTP 1.1 --- não autorizado: não autorizado devido a ACL no recurso"
      Case 401.4
        CheckError "ERRO HTTP 1.1 --- não autorizado: falha na autorização pelo filtro"
      Case 401.5
        CheckError "ERRO HTTP 1.1 --- não autorizado: falha na autorização por ISAPI/CGI App"
      Case 403.1
        CheckError "ERRO HTTP 1.1 --- proibido: acesso a execução proibido"
      Case 403.2
        CheckError "ERRO HTTP 1.1 --- proibido: acesso de leitura proibido"
      Case 403.3
        CheckError "ERRO HTTP 1.1 --- proibido: acesso de escrever proibido"
      Case 403.4
        CheckError "ERRO HTTP 1.1 --- proibido: requer SSL"
      Case 403.5
        CheckError "ERRO HTTP 1.1 --- proibido: requer SSL 128"
      Case 403.6
        CheckError "ERRO HTTP 1.1 --- proibido: endereço de IP rejeitado"
      Case 403.7
        CheckError "ERRO HTTP 1.1 --- proibido: requer certificação do cliente"
      Case 403.8
        CheckError "ERRO HTTP 1.1 --- proibido: acesso ao site negado"
      Case 403.9
        CheckError "ERRO HTTP 1.1 --- acesso proibido: Muitos usuários estão conectados"
      Case 403.10
        CheckError "ERRO HTTP 1.1 --- acesso proibido: configuração inválida"
      Case 403.11
        CheckError "ERRO HTTP 1.1 --- acesso proibido: senha alterada"
      Case 403.12
        CheckError "ERRO HTTP 1.1 --- acesso proibido: negado acesso ao mapa"
      Case 404
        CheckError "ERRO HTTP 1.1 --- não encontrado"
      Case 405
        CheckError "ERRO HTTP 1.1 --- método não permitido"
      Case 406
        CheckError "ERRO HTTP 1.1 --- não aceitável"
      Case 407
        CheckError "ERRO HTTP 1.1 --- requer autenticação do Proxy"
      Case 412
        CheckError "ERRO HTTP 1.1 --- falha em pré condições"
      Case 414
        CheckError "ERRO HTTP 1.1 --- pedido - URI muito grande"
      Case 423
        CheckError "Propriedade ou metodo não encontrado"
      Case 424
        CheckError "Objeto Requerido"
      Case 429
        CheckError "OLE Automation não pode ser criado no servidor"
      Case 430
        CheckError "Classe não suportada pelo OLE Automation"
      Case 432
        CheckError "Nome do arquivo ou de classe nõ encontrado durante a operação OLE Automation"
      Case 438
        CheckError "Objeto não suporta esta propriedade ou método"
      Case 440
        CheckError "Erro na OLE Automation"
      Case 442
        CheckError "Connection to type library or object library for remote process has been lost. Press OK for dialog to remove reference"
      Case 443
        CheckError "Objeto OLE Automation não contém um valor padrão"
      Case 445
        CheckError "Objeto não suporta esta ação"
      Case 446
        CheckError "Objeto não suporta o nome do argumento"
      Case 447
        CheckError "Objeto não suporta a definição do local atual"
      Case 448
        CheckError "Nome de argumentos não encontrados"
      Case 449
        CheckError "Este argumento não é opcional"
      Case 450
        CheckError "Número de argumentos errado ou definição de propriedade inválida"
      Case 451
        CheckError "Objeto não é uma coleção"
      Case 452
        CheckError "Número ordinal inválido"
      Case 453
        CheckError "Função DLL especificada não foi encontrada"
      Case 454
        CheckError "código de origem não encontrado"
      Case 455
        CheckError "Erro de trava no código"
      Case 457
        CheckError "Esta chave já está associada a um elemento desta coleção"
      Case 458
        CheckError "Tipos de variáveis usadas na OLE Automation não são suportadas pelo Visual Basic"
      Case 462
        CheckError "A máquina do servidor remoto não existe ou não está disponível"
      Case 481
        CheckError "Figura Inválida"
      Case 500
        CheckError "Variável não definida"
      Case 501
        CheckError "Variável não pode ser atribuída"
      Case 502
        CheckError "Objeto não é seguro para script"
      Case 503
        CheckError "Objeto não é seguro para inicialização"
      Case 504
        CheckError "Objeto não é seguro para criação"
      Case 505
        CheckError "Referência inválida ou não qualificada"
      Case 506
        CheckError "Classe não definida"
      Case 1001
        CheckError "Sem memória"
      Case 1002
        CheckError "Erro de Sintaxe"
      Case 1003
        CheckError "Esperado ':'"
      Case 1004
        CheckError "Esperado ';'"
      Case 1005
        CheckError "Esperado '('"
      Case 1006
        CheckError "Esperado ')'"
      Case 1007
        CheckError "Esperado ']'"
      Case 1008
        CheckError "Esperado '{'"
      Case 1009
        CheckError "Esperado '}'"
      Case 1010
        CheckError "Esperado Identificador"
      Case 1011
        CheckError "Esperado '='"
      Case 1012
        CheckError "Esperado 'If'"
      Case 1013
        CheckError "Esperado 'To'"
      Case 1014
        CheckError "Esperado 'End'"
      Case 1015
        CheckError "Esperado 'Function'"
      Case 1016
        CheckError "Esperado 'Sub'"
      Case 1017
        CheckError "Esperado 'Then'"
      Case 1018
        CheckError "Esperado 'Wend'"
      Case 1019
        CheckError "Esperado 'Loop'"
      Case 1020
        CheckError "Esperado 'Next'"
      Case 1021
        CheckError "Esperado '    Case'"
      Case 1022
        CheckError "Esperado 'Select'"
      Case 1023
        CheckError "Esperado expressão"
      Case 1024
        CheckError "Esperado declaração"
      Case 1025
        CheckError "Esperado final da declaração"
      Case 1026
        CheckError "Esperado inteiro constante"
      Case 1027
        CheckError "Esperado 'While' , 'Until'"
      Case 1028
        CheckError "Esperado 'While' , 'Until' ou final de declaração"
      Case 1029
        CheckError "Esperado 'With'"
      Case 1030
        CheckError "Identificador Muito Longo"
      Case 1031
        CheckError "Número Inválido"
      Case 1032
        CheckError "Caracter Inválido"
      Case 1033
        CheckError "Constante de String não Terminada"
      Case 1034
        CheckError "Comentário não Terminado"
      Case 1035
        CheckError "Nested Comment"
      Case 1036
        CheckError "'Me' não pode ser usado como saída de rotina"
      Case 1037
        CheckError "Uso Inválido da Palavra Chave 'Me'"
      Case 1038
        CheckError "'Loop' sem 'Do'"
      Case 1039
        CheckError "Declaração 'Exit' Inválida"
      Case 1040
        CheckError "Variável de Controle de Loop 'for' Inválida"
      Case 1041
        CheckError "Variável Redefinida"
      Case 1042
        CheckError "Tem que ser a primeira declaração da linha"
      Case 1043
        CheckError "Não pode atribuir non-By Val para um argumento"
      Case 1044
        CheckError "Não pode usar parêntesis para chamar uma sub"
      Case 1045
        CheckError "Esperada Constante Literal"
      Case 1046
        CheckError "Esperado 'In'"
      Case 1047
        CheckError "Esperado 'Class'"
      Case 1048
        CheckError "Tem que ser definido dentro de uma Classe"
      Case 1049
        CheckError "Esperado Let ou Set ou Get na declaração de propriedade"
      Case 1050
        CheckError "Esperado 'Property'"
      Case 1051
        CheckError "Número de argumentos tem que ser consistente em especificações de propriedades"
      Case 1052
        CheckError "Não pode haver método/ propriedade padrão múltiplo em uma Classe"
      Case 1053
        CheckError "Class initialize ou terminate não tem argumentos"
      Case 1054
        CheckError "Propriedade Set ou Let tem que ter pelo menos um argumento"
      Case 1055
        CheckError "'Next' inesperado"
      Case 1056
        CheckError "'Default' pode ser especificado somente em 'Property' ou 'Function' ou 'Sub'"
      Case 1057
        CheckError "Especificação 'Default' precisa especificar também 'Public'"
      Case 1058
        CheckError "Especificação 'Default' só pode estar em Property Get"

      Case 3000
        CheckError "O provedor não concluiu a ação pedida"
      Case 3001
        CheckError "A aplicação está usando argumentos do tipo errado, estão fora do âmbito aceitável ou em conflito com alguma outra aplicação"
      Case 3002
        CheckError "Ocorreu um erro durante a abertura do arquivo pedido"
      Case 3003
        CheckError "Erro na leitura do arquivo especificado"
      Case 3004
        CheckError "Erro ao escrever no arquivo"
      Case 3021
        CheckError "BOF ou EOF é True ou o registro atual foi deletado. A operação pedido pela aplicação requer um registro atual"
      Case 3219
        CheckError "A operação pedida pela aplicação não é permitida neste contexto"
      Case 3246
        CheckError "A aplicação não pode fechar explicitamente um objeto connection no meio de uma transação"
      Case 3251
        CheckError "O provedor não oferece suporte a operação pedida pela aplicação"
      Case 3265
        CheckError "ADO não pode achar o objeto na coleção"
      Case 3367
        CheckError "Não é anexar, objeto já está na coleção"
      Case 3420
        CheckError "O objeto referenciado pela aplicação não aponta mais para um objeto válido"
      Case 3421
        CheckError "A aplicação está usando um valor do tipo errado para a aplicação atual"
      Case 3704
        CheckError "A operação pedida pela aplicação não é permitida se o objeto estiver fechado"
      Case 3705
        CheckError "A operação pedida pela aplicação não é permitida se o objeto estiver aberto"
      Case 3706
        CheckError "ADO não pode achar o provedor especificado"
      Case 3707
        CheckError "A aplicação não pode alterar a propriedade ActiveConnect de um objeto Recordset com um objeto Command como fonte"
      Case 3708
        CheckError "A aplicação definiu de modo impróprio um objeto Parameter"
      Case 3709
        CheckError "A aplicação pediu uma operação em um objeto com uma referência a um objeto Connection inválido ou fechado"
      Case 3710
        CheckError "A operação não é reentrante"
      Case 3711
        CheckError "A operação ainda está executando"
      Case 3712
        CheckError "Operação cancelada"
      Case 3713
        CheckError "A operação ainda está conectando"
      Case 3714
        CheckError "A transação é inválida"
      Case 3715
        CheckError "A operação não está sendo executada"
      Case 3716
        CheckError "A operação não é segura sob estas circunstâncias"
      Case 3717
        CheckError "A operação fez com que aparecesse uma caixa de diálogo"
      Case 3718
        CheckError "A operação fez com que aparecesse um cabeçalho de caixa de diálogo"
      Case 3719
        CheckError "A ação falhou devido a uma violação na integridade dos dados"
      Case 3720
        CheckError "O provedor não pode ser modificado"
      Case 3721
        CheckError "Dados longos demais para o tipo de dados apresentados"
      Case 3722
        CheckError "Ação causou uma violação do esquema"
      Case 3723
        CheckError "A expressão continha sinais não coincidentes"
      Case 3724
        CheckError "O valor não pode ser convertido"
      Case 3725
        CheckError "O recurso não pode ser criado"
      Case 3726
        CheckError "A coluna especificada não existe nesta fileira"
      Case 3727
        CheckError "O URL não existe"
      Case 3728
        CheckError "Você não tem permissão para ver a árvore do diretório"
      Case 3729
        CheckError "O URL apresentado é inválido"
      Case 3730
        CheckError "Recurso travado"
      Case 3731
        CheckError "Recurso já existente"
      Case 3732
        CheckError "A ação não pode ser concluída"
      Case 3733
        CheckError "O volume de arquivo não foi encontrado"
      Case 3734
        CheckError "Falha na operação porque o servidor não pode obter espaço suficiente para completar a operação"
      Case 3735
        CheckError "Recurso fora de âmbito"
      Case 3736
        CheckError "Comando não está disponível"
      Case 3737
        CheckError "O URL na fileira identificada não existe"
      Case 3738
        CheckError "O recurso não pode ser deletado porque está fora do escopo permitido"
      Case 3739
        CheckError "Esta propriedade é inválida para a coluna selecionada"
      Case 3740
        CheckError "Você apresentou uma opção inválida para esta propriedade"
      Case 3741
        CheckError "Você apresentou um valor inválido para esta propriedade"
      Case 3742
        CheckError "A definição desta propriedade causou um conflito com outras propriedades"
      Case 3743
        CheckError "Nem todas as propriedades podem ser definidas"
      Case 3744
        CheckError "A propriedade não foi definida"
      Case 3745
        CheckError "A propriedade não pode ser definida"
      Case 3746
        CheckError "A propriedade não tem suporte"
      Case 3747
        CheckError "A ação não pode ser concluída porque o catálogo não está definido"
      Case 3748
        CheckError "A conexão não pode ser alterada"
      Case 3749
        CheckError "O método Update da coleção Fields falhou"
      Case 3750
        CheckError "Não é possível definir permissão Deny porque o provedor não oferece suporte para tanto"
      Case 3751
        CheckError "o provedor não oferece suporte ao tipo de pedido"
    End Select

    '-----	
    'Antes de qualquer código asp
	'On Error Resume Next
    'Depois do último código asp (fim da página por exemplo)
    'IF Err > 0 Then Response.Wirit CheckError()
  End Function
  '.................................................................................................

 '.................................................................................................
  '<----- 01 - FUNÇÃO DE ACERTO DE DATA
          'Função que recebe a data no formato no microsiga aaaammdd e Devolve no formato normal dd/mm/aaaa
		  'Parâmetros:
		  'Data - Data no formato aaaammdd para ser devolvida em dd/mm/aaaa
  Function AcertaData(Data)
    Dim Dia, Mes, Ano
    Dia = Right(Data,2)
    Mes = Mid(Data,5,2)
    Ano = Left(Data,4)
    AcertaData = Dia & "/" & Mes & "/" & Ano
  
    '-----
    'CÓDIGO PARA TESTAR A FUNÇÃO	
    'Response.Write AcertaData('20070101') 
  End Function 
  '.................................................................................................

 '.................................................................................................
  '<----- 01 - FUNÇÃO DE INVERSÃO DE DATA
          'Função que recebe a data no formato no microsiga dd/mm/aaaa e Devolve no formato microsiga aaaammdd
		  'Parâmetros:
		  'Data - Data no formato dd/mm/aaaa para ser devolvida aaaammdd
  Function InverteData(Data)
    Dim Ano, Mes, Dia
    Ano = Year(Data)
    Mes = Month(Data)
    If Len(Mes) < 2 Then
      Mes = "0" & Mes
    End If
    Dia = Day(Data)
    If Len(Dia) < 2 Then
      Dia = "0" & Dia
    End If
    InverteData = Ano & Mes & Dia
  End Function 
 '.................................................................................................

 '.................................................................................................
  '<----- 01 - FUNÇÃO COMPLETA OS CARACTERES DO MES E DIA DA DATA
          'Função que recebe a data no formato no microsiga d/m/aaaa e Devolve no formato dd/mm/aaaa
		  'Parâmetros:
		  'Data - Data no formato d/m/aaaa para ser devolvida dd/mm/aaaa  
  Function CompletaData(Data)
    Dim Ano, Mes, Dia 
    Ano = Year(Data)
    Mes = Month(Data)
    If Len(Mes) < 2 Then
      Mes = "0" & Mes
    End If
    Dia = Day(Data)
    If Len(Dia) < 2 Then
      Dia = "0" & Dia
    End If
    CompletaData = Dia & "/" & Mes & "/" & Ano
  End Function    
 '.................................................................................................

 '.................................................................................................
  '<----- 01 - FUNÇÃO COMPLETA OS CARACTERES DA HORA E TRANSFORMA EM 24 HORAS
          'Função que recebe hora no e completa os caracteres para 2 dígitos  e transforma em 24 horas
		  'Parâmetros:
		  'Hora - Devole no padrão 24 horas com dois dígitos
  Function Acerta24Horas(Horas)
    'Dim Hora, Minuto, Segundo, E24Horas
	'If uCase(Right(Horas,2)) = "PM" Or uCase(Right(Horas,2)) = "AM" Then 
      'E24Horas = False
	'Else
	  'E24Horas = True
	'End If
	
	'Hora    = DatePart("h", Horas)
	'Minuto  = DatePart("n", Horas)
	'Segundo = DatePart("s", Horas)
	
	'If E24Horas = False Then
	  'If uCase(Right(Horas,2)) = "PM" Then
	    'If Hora <> "12" Then
		
		'End If
	  'End If
	'End If
	
	'Acerta24Horas = Minuto
  End Function    
 '.................................................................................................   
End Class
%>

