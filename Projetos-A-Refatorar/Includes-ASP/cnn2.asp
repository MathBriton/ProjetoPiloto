<!--#include file = "../funcoes/funcoes.asp"-->
<% Server.ScriptTimeout = 180 %>
<%
'--------------------------------------------------------------------------------------------------
' FONTE    : cnn2.asp
' CRIAÇÃO  : 09/02/2024
' AUTOR    : JORGE.QUEIROZ
' DESCRIÇÃO: OTIMIZAR AS CONEXÕES COM OS BANCOS DE DADOS;
'            OTIMIZAR A LEITURA DAS STRINGS DE CONEXÃO NOMEADO-AS
'               EX.: PIR-DADOSADV
'            FACILITAR A LEITURA DO FONTE, NOMEANDO AS VARIÁVEIS DO RECORDSET
'               EX.: DADOSADV1 = Conecta( "PIR-DADOSADV", SQL )
'--------------------------------------------------------------------------------------------------
'On Error GoTo 0

Session.TimeOut = 60
Server.ScriptTimeout = 1200

' -------------------------------------------------------------------------------------------------
' FUNÇÃO    : Conecta
' DESCRIÇÃO : Abrir a conexão com o banco de dados específico, conforme o parâmetro [nomeDaStringDeConexao] 
'             e executar a query passada no parâmetro queryAexecutar
' PARÂMETROS: nomeDaStringDeConexao - nome da string de conexão no arquivo conexão.txt
'             aSQL - query que será executada
' -------------------------------------------------------------------------------------------------
function Conecta( nomeDaStringDeConexao, queryAexecutar)

  On Error Resume Next
  
  Const adOpenStatic = 3    
  Const adUseClient = 3
  Const adLockBatchOptimistic = 4 

  DIM aRetorno

  SET Conexao = NOTHING
  SET Dados = NOTHING

  StringDeConexao = BuscaStringConexao( ucase(nomeDaStringDeConexao) )
  
  'response.write "StringDeConexao: " & StringDeConexao & "<br>"

  if trim(StringDeConexao) <> "" then

    On Error Resume Next

    Set Conexao = Server.CreateObject("ADODB.Connection")
    Conexao.CommandTimeout = 300
    Conexao.Open StringDeConexao
   
    if (err.number <> 0) OR (Conexao Is Nothing) then
      'ERRO DE CONEXAO ------------------------------------------------------------------------------------------------------------------------------
      txtErro = "CONEXAO (Conecta) : " & nomeDaStringDeConexao

      call HandleError( err, GetFileName(), txtErro )

      response.write "" &_
        "<html>" &_
          "<head>" &_
            "<meta charset='utf-8' />" &_
            "<meta name='viewport' content='width=device-width, initial-scale=1'/>" &_
            "<link href='https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css' rel='stylesheet' integrity='sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC' crossorigin='anonymous' />" &_
            "<script src='https://code.jquery.com/jquery-3.5.1.slim.min.js' integrity='sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj' crossorigin='anonymous'></script>" &_
            "<link rel='stylesheet' href='https://use.fontawesome.com/releases/v5.7.0/css/all.css' integrity='sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ' crossorigin='anonymous' />" &_
            "<script src='https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.11/jquery.mask.min.js'></script>"  &_
          "</head>" &_
          "<body>" &_
            "<div class='alert alert-danger alert-dismissible' role='alert'>" &_
              "<a href='#' class='close' data-dismiss='alert' aria-label='close'>&times;</a>" &_
              "<br>" &_
              "<br>" & GetFileName() &_
              "<br>" & StringDeConexao &_
              "<br>" & Err.Number & " - " & Err.Description &_
              "<br>" &_
              "<strong>A V I S O !!!</strong> Erro ao conectar ao banco de dados, verifique com a equipe de suporte." &_
            "</div>" &_
          "<body>" &_
        "</html>"              
      response.end

      err.clear

      SET Conecta = NOTHING
      exit function

    else

      On Error Resume Next

      'CONECTOU -------------------------------------------------------------------------------------------------------------------------------------
      Set RecordSet = Server.CreateObject("ADODB.Recordset")
      RecordSet.CursorLocation = adUseClient

      'CALL EnviaEmail( "jorge.queiroz@smurfitkappa.com.br","","","","", queryAexecutar , "TESTE - CNN2" )

      RecordSet.Open queryAexecutar, Conexao ', adOpenStatic, adLockBatchOptimistic

      if err.number <> 0 then
        Response.Write("Erro: " & err.number & " - " & err.description & "<br>")

        txtErro = "ERRO AO EXECUTAR UMA QUERY"

        'response.write "txtErro: " & txtErro & "<br>"

        call HandleError( err, GetFileName(), txtErro, queryAexecutar )
        err.clear

        RecordSet.Close
        Conexao.Close
        SET RecordSet = NOTHING
        SET Conexao = NOTHING


        response.write "" &_
          "<html>" &_
            "<head>" &_
              "<meta charset='utf-8' />" &_
              "<meta name='viewport' content='width=device-width, initial-scale=1'/>" &_
              "<link href='https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css' rel='stylesheet' integrity='sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC' crossorigin='anonymous' />" &_
              "<script src='https://code.jquery.com/jquery-3.5.1.slim.min.js' integrity='sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj' crossorigin='anonymous'></script>" &_
              "<link rel='stylesheet' href='https://use.fontawesome.com/releases/v5.7.0/css/all.css' integrity='sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ' crossorigin='anonymous' />" &_
              "<script src='https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.11/jquery.mask.min.js'></script>"  &_
            "</head>" &_
            "<body>" &_
              "<div class='alert alert-danger alert-dismissible' role='alert'>" &_
                "<a href='#' class='close' data-dismiss='alert' aria-label='close'>&times;</a>" &_
                "<strong>A V I S O !!!</strong> Erro ao executar a consulta, verifique com a equipe de suporte." &_
                queryAexecutar &_
              "</div>" &_
            "<body>" &_
          "</html>"
        response.end

        SET Conecta = NOTHING
        exit function

      else

       'response.write nomeDaStringDeConexao & "<br>" & queryAexecutar & "<br>" 
        SET Conecta = RecordSet
        exit function

      end if

    end if

  else

    err.clear

    response.write "" &_
      "<html>" &_
        "<head>" &_
          "<meta charset='utf-8' />" &_
          "<meta name='viewport' content='width=device-width, initial-scale=1'/>" &_
          "<link href='https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css' rel='stylesheet' integrity='sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC' crossorigin='anonymous' />" &_
          "<script src='https://code.jquery.com/jquery-3.5.1.slim.min.js' integrity='sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj' crossorigin='anonymous'></script>" &_
          "<link rel='stylesheet' href='https://use.fontawesome.com/releases/v5.7.0/css/all.css' integrity='sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ' crossorigin='anonymous' />" &_
          "<script src='https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.11/jquery.mask.min.js'></script>"  &_
        "</head>" &_
        "<body>" &_
          "<div class='alert alert-danger alert-dismissible' role='alert'>" &_
            "<a href='#' class='close' data-dismiss='alert' aria-label='close'>&times;</a>" &_
            "<strong>E R R O !!!</strong> O nome da String de Conexão " & ucase(nomeDaStringDeConexao) & ", Não Existe!!!." &_
          "</div>" &_
        "<body>" &_
      "</html>"
      response.end

      SET Conecta = NOTHING
      exit function

  end if

end function
' -------------------------------------------------------------------------------------------------
' FUNÇÃO    : BuscaStringConexao
' DESCRIÇÃO : Procura a string de conexão conforme o parâmetro [nomeDaStringDeConexao]
' PARÂMETROS: [nomeDaStringDeConexao] - nome da string de conexão no arquivo conexão.txt
' -------------------------------------------------------------------------------------------------
function BuscaStringConexao( aStringDeConexao )

  Dim objFSO, ObjAbreArq, campo, parametros(50,5), strconexao, ArrayLinha, LinCnx

  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set ObjAbreArq = objFSO.OpenTextFile("c:\inetpub\wwwroot\corporativo\cnn2\conexao.txt")

  x = 0
  y = 0

  LinCnx = ""

  while not ObjAbreArq.AtEndOfStream
    LinCnx = ObjAbreArq.ReadLine
    ArrayLinha = split( LinCnx, ";" )
    for y = 0 to 5
      parametros(x,y) = trim(ArrayLinha(y))
    next
    x = x + 1
  wend
  x = 0

  for x = 0 to Ubound(parametros) 
    if parametros(x,0) = aStringDeConexao then
      strconexao = "Provider=" & parametros(x,1) & ";Data Source=" & parametros(x,2) & ";Initial Catalog=" & parametros(x,3) & ";User ID=" & base64_decode(parametros(x,4)) & ";Password=" & base64_decode(parametros(x,5)) & ";"
    end if 
  next

  ObjAbreArq.Close
  Set ObjAbreArq = NOTHING

  BuscaStringConexao = strconexao

end function
' -------------------------------------------------------------------------------------------------
' FUNÇÃO    : ConectaDB
' DESCRIÇÃO : Abrir a conexão com o banco de dados específico, conforme o parâmetro [nomeDaStringDeConexao] 
' PARÂMETROS: nomeDaStringDeConexao - nome da string de conexão no arquivo conexão.txt
' -------------------------------------------------------------------------------------------------
public function ConectaDB( nomeDaStringDeConexao )

  On Error Resume Next

  SET Conexao = Nothing
  SET Dados = Nothing

  DIM aRetorno

  aRetorno = True

  StringDeConexao = BuscaStringConexao( ucase(nomeDaStringDeConexao) )

  'response.write "StringDeConexao: " & StringDeConexao & "<br>"

  if trim(StringDeConexao) <> "" then

    SET Conexao = Server.CreateObject("ADODB.Connection")
    Conexao.CommandTimeout = 60
    Conexao.Open StringDeConexao

    if err.number <> 0 then

      txtErro = "CONEXAO (ConectaDB) : " & ucase(nomeDaStringDeConexao)
      call HandleError( err, GetFileName(), txtErro, "" )
      err.clear

      response.write "" &_
        "<html>" &_
          "<head>" &_
            "<meta charset='utf-8' />" &_
            "<meta name='viewport' content='width=device-width, initial-scale=1'/>" &_
            "<link href='https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css' rel='stylesheet' integrity='sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC' crossorigin='anonymous' />" &_
            "<script src='https://code.jquery.com/jquery-3.5.1.slim.min.js' integrity='sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj' crossorigin='anonymous'></script>" &_
            "<link rel='stylesheet' href='https://use.fontawesome.com/releases/v5.7.0/css/all.css' integrity='sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ' crossorigin='anonymous' />" &_
            "<script src='https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.11/jquery.mask.min.js'></script>"  &_
          "</head>" &_
          "<body>" &_
            "<div class='alert alert-danger alert-dismissible' role='alert'>" &_
              "<a href='javascript:history.back()' class='close' data-dismiss='alert' aria-label='close'>&times;</a>" &_
              "<strong>A V I S O !!!</strong> Não foi possível se conectar ao banco de dados " & ucase(nomeDaStringDeConexao) & ", verifique com a equipe de suporte." &_
            "</div>" &_
          "<body>" &_
        "</html>"
      response.end      

      aRetorno = False

    end if

  else

    response.write "" &_
      "<html>" &_
        "<head>" &_
          "<meta charset='utf-8' />" &_
          "<meta name='viewport' content='width=device-width, initial-scale=1'/>" &_
          "<link href='https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css' rel='stylesheet' integrity='sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC' crossorigin='anonymous' />" &_
          "<script src='https://code.jquery.com/jquery-3.5.1.slim.min.js' integrity='sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj' crossorigin='anonymous'></script>" &_
          "<link rel='stylesheet' href='https://use.fontawesome.com/releases/v5.7.0/css/all.css' integrity='sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ' crossorigin='anonymous' />" &_
          "<script src='https://cdnjs.cloudflare.com/ajax/libs/jquery.mask/1.14.11/jquery.mask.min.js'></script>" &_
        "</head>" &_
        "<body>" &_
          "<div class='alert alert-danger alert-dismissible' role='alert'>" &_
            "<a href='javascript:history.back()' class='close' data-dismiss='alert' aria-label='close'>&times;</a>" &_
            "<strong>A V I S O !!!</strong> O nome da String de Conexão " & ucase(nomeDaStringDeConexao) & ", Não Existe!!!" &_
          "</div>" &_
        "<body>" &_
      "</html>"
    response.end
    
    aRetorno = False

  end if

  'response.write StringDeConexao & "<br>"
  'response.write "agora: " & now() & "<br>"
  'response.write "retorno: " & aRetorno & "<br>"
  'response.write "vartype: " & vartype(aRetorno) & "<br>"
  'response.end

  Conexao.Close

  ConectaDB = aRetorno
  
end function
' -------------------------------------------------------------------------------------------------'
function PesquisaObjetoNaQuery(aQuery)

  DIM aRetorno 
  DIM aComandos
  aComandos = array( "SELECT ", "EXEC ", "FROM " )

  aRetorno = ""

  for i = 0 to UBound(aComandos)

    'response.write aComandos(i) & " - " & inStr( 1, ucase(aQuery), aComandos(i) ) & "<br>"

    pComando = inStr( 1, ucase(aQuery), aComandos(i) )

    if pComando > 0 then

      pProximo = (pComando + len(aComandos(i)) + 1)
      fObject = inStr( pProximo, aQuery, " " )      
      aObject = mid(aQuery, pProximo, fObject - pProximo)

      ' response.write "aQuery: " & aQuery & "<br>"
      ' response.write "pProximo: " & aObject & "<br>"
      ' response.write "fObject: " & fObject & "<br>"
      ' response.write "aObject: " & aObject & "<br>"

      aRetorno = aObject
      exit for

    end if

  next

  'response.write "aRetorno: " & aRetorno

  PesquisaObjetoNaQuery = aRetorno

end function
%>