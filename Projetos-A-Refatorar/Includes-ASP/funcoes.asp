<% 
Base64Chars =	"ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
		          "abcdefghijklmnopqrstuvwxyz" & _
		          "0123456789" & _
		          "+/"

' functions for encoding string to Base64 ---------------------------------------------------------
public function base64_encode( byVal strIn )

	Dim c1, c2, c3, w1, w2, w3, w4, n, strOut
	For n = 1 To Len( strIn ) Step 3
		c1 = Asc( Mid( strIn, n, 1 ) )
		c2 = Asc( Mid( strIn, n + 1, 1 ) + Chr(0) )
		c3 = Asc( Mid( strIn, n + 2, 1 ) + Chr(0) )
		w1 = Int( c1 / 4 ) : w2 = ( c1 And 3 ) * 16 + Int( c2 / 16 )
		if Len( strIn ) >= n + 1 Then 
			w3 = ( c2 And 15 ) * 4 + Int( c3 / 64 ) 
		else 
			w3 = -1
		end if
		if Len( strIn ) >= n + 2 Then 
			w4 = c3 And 63 
		else 
			w4 = -1
		end if
		strOut = strOut + mimeencode( w1 ) + mimeencode( w2 ) + mimeencode( w3 ) + mimeencode( w4 )
	Next
	base64_encode = strOut

end function

' FUNÇÃO mimeencode -------------------------------------------------------------------------------
Private function mimeencode( byVal intIn )
	if intIn >= 0 Then 
		mimeencode = Mid( Base64Chars, intIn + 1, 1 ) 
	else 
		mimeencode = ""
	end if
end function	

' function to decode string from Base64 -----------------------------------------------------------
public function base64_decode( byVal strIn )
	Dim w1, w2, w3, w4, n, strOut
	For n = 1 To Len(strIn) step 4
		w1 = mimedecode(Mid(strIn, n, 1 ))
		w2 = mimedecode(Mid(strIn, n + 1, 1 ))
		w3 = mimedecode(Mid(strIn, n + 2, 1 ))
		w4 = mimedecode(Mid(strIn, n + 3, 1 ))
		if w2 >= 0 Then _
			strOut = strOut + _
      Chr((( w1 * 4 + Int( w2 / 16 )) And 255 ))
		if w3 >= 0 Then _
			strOut = strOut + _
      Chr((( w2 * 16 + Int( w3 / 4 )) And 255 ))
		if w4 >= 0 Then _
			strOut = strOut + _
      Chr(((w3 * 64 + w4) And 255))
	Next
	base64_decode = strOut
end function

'--------------------------------------------------------------------------------------------------
Private function mimedecode( byVal strIn )
	if Len( strIn ) = 0 Then 
		mimedecode = -1 : Exit function
	else
		mimedecode = InStr( Base64Chars, strIn ) - 1
	end if
end function

' -------------------------------------------------------------------------------------------------
'DATA NO FORMATO MICROSIGA
function invertedata(data)

  'dia = day(data)
  dia = mid(data,1,2)
  if len(dia) = 1 then
    dia = "0"&dia
  end if
  'mes = month(data)
  mes = mid(data,4,2)
  if len(mes) = 1 then
    mes = "0"&mes
  end if
  ano = mid(data,7,4)
  'ano = year(data)
  invertedata = ano&""&mes&""&dia

end function   
' -------------------------------------------------------------------------------------------------
'DATA FORMATO NORMAL
function acertadata(data)

  if Trim(data) <> "" then
    Hora = ""
    if len(data) = 8 then
      'FORMATO AAAAMMDD
      dia = mid(data,7,2)
      if len(dia) = 1 then
        dia = "0"&dia
      end if
      mes = mid(data,5,2)
      if len(mes) = 1 then
        mes = "0"&mes
      end if
      ano = mid(data,1,4)
    else
      if len(data) > 10 then
        'FORMATO AAAAMMDD HH:MM:SS
        dia = mid(data,7,2)
        if len(dia) = 1 then
          dia = "0"&dia
        end if
        mes = mid(data,5,2)
        if len(mes) = 1 then
          mes = "0"&mes
        end if
        ano = mid(data,1,4)
        Hora = mid(data,9,9)
      else
        'FORMATO AAAA/MM/DD
        dia = Trim(mid(data,9,2))
        if len(dia) = 1 then
          dia = "0"&dia
        end if
        mes = Trim(mid(data,6,2))
        if len(mes) = 1 then
          mes = "0"&mes
        end if
        ano = mid(data,1,4)
      end if
    end if
    acertadata = dia&"/"&mes&"/"&ano&""&hora
  else
    acertadata = ""
  end if

end function
' -------------------------------------------------------------------------------------------------
function FormataCnpj(cnpj)

  if len(cnpj) > 12 then

    digito = Right(cnpj,2) 
    filial_dig = mid(Right(cnpj,6),1,4)
    seq3 = mid(right(cnpj,9),1,3)
    seq2 = mid(right(cnpj,12),1,3)
    seq1 = mid(cnpj,1,(len(cnpj)-(len(digito)+len(filial_dig)+len(seq3)+len(seq2))))

    FormataCnpj = seq1&"."&seq2&"."&seq3&"/"&filial_dig&"-"&digito

  else

    if len(cpf) > 8 then

      digito = Right(cpf,2) 
      seq3 = mid(right(cpf,9),1,3)
      seq2 = mid(right(cpf,12),1,3)
      seq1 = mid(cpf,1,(len(cpf)-(len(digito)+len(seq3)+len(seq2))))

      FormataCnpj= seq1&"."&seq2&"."&seq3&"-"&digito

    else

      FormataCnpj = cnpj

    end if

  end if

end function
' -------------------------------------------------------------------------------------------------
function FormataCpf(cpf)

  if len(cpf) > 8 then

    digito = Right(cpf,2) 
    seq3 = mid(right(cpf,9),1,3)
    seq2 = mid(right(cpf,12),1,3)
    seq1 = mid(cpf,1,(len(cpf)-(len(digito)+len(seq3)+len(seq2))))

    FormataCpf= seq1&"."&seq2&"."&seq3&"-"&digito

  else

    FormataCpf = cpf

  end if

end function
' -------------------------------------------------------------------------------------------------
function TirarPontoCNPJ(cnpj)

  response.write cnpj
          cnpj = replaceAll(cnpj,".","")
          cnpj = replace(cnpj,"/","")
          cnpj = replace(cnpj,"-","")
          TirarPontoCNPJ = cnpj

end function 
' -------------------------------------------------------------------------------------------------
'FORMATANDO VALOR - SUBSTITUINDO A VIRGULA PELO PONTO
function formatavalor(valor)

  formatavalor = replace(replace(valor,".",""),",",".")

end function  
' -------------------------------------------------------------------------------------------------
function EnviaEmail(Para, Copia1, Copia2, Copia3, De, Mensagem, Assunto)

    DIM myMail

    ' If not EmailValido(Para) Then
    '   exit function 
    ' End If

    if Request.ServerVariables("server_name") = "br-pir-m012" then 

        Para   = "leone.souza@smurfitwestrock.com.br;" &_
                 "jorge.queiroz@smurfitwestrock.com.br;" &_
                 "carlos.souza@smurfitwestrock.com.br;" &_
                 "rodolpho.rocha@smurfitwestrock.com.br"
        Copia1 = ""
        Copia2 = ""
        Copia3 = ""

    end if

    SET myMail = CreateObject("CDO.Message")

    myMail.Subject = Assunto
    myMail.From = "noreply.brasil@smurfitkappa.com.br"
    myMail.To = Para

    if Trim(Copia1) <> "" Then
        copia = Copia1
    end if

    if Trim(Copia2) <> "" Then
        copia = copia & ";" & Copia2
    end if

    if Trim(Copia3) <> "" Then
        copia = copia & ";" & Copia3
    end if

    if Trim(Copia) <> "" then
        myMail.CC = Copia    
    end if

    myMail.HTMLBody = Mensagem

    myMail.BodyPart.Charset = "UTF-8"
    myMail.HTMLBodyPart.Charset = "UTF-8"

    myMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    'Name or IP of remote SMTP server
    myMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "cas.na.smurfitkappa.com"

    myMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

    'Your UserID on the SMTP server
    myMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "group\noreply.brasil"

    'Your password on the SMTP server
    myMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "2@!F8r2j15Of9a%"

    'Server port
    myMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

    'Use SSL for the connection (False or True)
    myMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = false

    'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
    myMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

    myMail.Configuration.Fields.Update()
    myMail.Send

    SET myMail = NOTHING

end function
' -------------------------------------------------------------------------------------------------
function formata_cnpj(cnpj)

  if len(cnpj) > 12 then

    digito = Right(cnpj,2) 
    filial_dig = mid(Right(cnpj,6),1,4)
    seq3 = mid(right(cnpj,9),1,3)
    seq2 = mid(right(cnpj,12),1,3)
    seq1 = mid(cnpj,1,(len(cnpj)-(len(digito)+len(filial_dig)+len(seq3)+len(seq2))))

    formata_cnpj = seq1&"."&seq2&"."&seq3&"/"&filial_dig&"-"&digito

  else

    if len(cpf) > 8 then

      digito = Right(cpf,2) 
      seq3 = mid(right(cpf,9),1,3)
      seq2 = mid(right(cpf,12),1,3)
      seq1 = mid(cpf,1,(len(cpf)-(len(digito)+len(seq3)+len(seq2))))

      formata_cpf= seq1&"."&seq2&"."&seq3&"-"&digito

    else

      formata_cnpj = cnpj

    end if

  end if

end function
' -------------------------------------------------------------------------------------------------
function formata_cpf(cpf)

  if len(cpf) > 8 then

    digito = Right(cpf,2) 
    seq3 = mid(right(cpf,9),1,3)
    seq2 = mid(right(cpf,12),1,3)
    seq1 = mid(cpf,1,(len(cpf)-(len(digito)+len(seq3)+len(seq2))))

    formata_cpf= seq1&"."&seq2&"."&seq3&"-"&digito

  else

    formata_cpf = cpf

  end if

end function
' -------------------------------------------------------------------------------------------------
function ZeroEsquerda( digitos, valor )

  DIM novoValor

  novoValor = trim(valor)

  if len(trim(valor)) < digitos then

    for x = 1 to digitos - len(trim(valor))
      novoValor = "0" & novoValor
    next

  end if

  ZeroEsquerda = trim(novoValor)

end function
' -------------------------------------------------------------------------------------------------
function UltimoDia(data)

  UltimoDia = dateadd("d",-1,cDate("01/" & month(dateadd("m",1,data)) & "/" & year(dateadd("m",1,data))))

end function
' -------------------------------------------------------------------------------------------------
function DiaUtil(data_inicial,data_final)

	diferenca = DateDiff("d",data_inicial,data_final)
	dia = Weekday(data_inicial)

  if dia = 1 or dia = 7 then dia = dia + 1

	for contador = 0 to diferenca

		if dia = 1 Then weekend = weekend + 1
		if dia = 7 Then weekend = weekend + 1
    dia = dia + 1
		if dia > 7 Then dia = 1

	Next

	DiaUtil = diferenca - weekend

end function
' -------------------------------------------------------------------------------------------------
function ProximoDiaUtil(Data)

  Dia = Weekday(Data)
  ProximoDiaUtil = iif( (Dia = 7), DateAdd("d",2,Data), DateAdd("d",1,Data) )

end function
' -------------------------------------------------------------------------------------------------
function PermiteAcesso(nivel)

  if Session("Nivel") > 0 Then

    Texto = "" &_
      "<html lang='pt-BR'>" &_
        "<head>" &_
          "<meta charset='utf-8' />" &_
          "<title>SWnet</title>" &_
          "<meta name='viewport' content='width=device-width, initial-scale=1' />" &_
          "<meta http-equiv='PRAGMA' content='no-cache' />" &_
          "<meta http-equiv='cache-control' content='no-cache' />" &_
          "<link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css' />" &_
          "<link rel='stylesheet' href='https://use.fontawesome.com/releases/v5.7.0/css/all.css' integrity='sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ' crossorigin='anonymous' />" &_
          "<script src='https://code.jquery.com/jquery-3.2.1.slim.min.js' integrity='sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN' crossorigin='anonymous'></script>" &_
          "<script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.3/umd/popper.min.js' integrity='sha384-vFJXuSJphROIrBnz7yo7oB41mKfc8JzQZiCq4NCceLEaO4IHwicKwpJf9c9IpFgh' crossorigin='anonymous'></script>" &_
          "<script src='https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-beta.2/js/bootstrap.min.js' integrity='sha384-alpBpkh1PFOepccYVYDB4do5UnbKysX5WZXm3XxPqe5iKTfUKjNkCk9SaVuEZflJ' crossorigin='anonymous'></script>" &_
          "<style>" &_
          ".customer-logo { border: 0px; vertical-align: bottom; width: 500px; height: 146px; }" &_
          ".nome{" &_
            "font-family: Georgia, serif;" &_
            "font-size: 130px;" &_
            "font-weight: bold;" &_
            "color: transparent;" &_
            "background-color: #555;" &_
            "text-shadow: 1px 1px 3px rgba(255,255,255,.5);" &_
            "-webkit-background-clip: text;" &_
            "-moz-background-clip: text;" &_
            "background-clip: text;" &_
          "}" &_
          "</style>" &_
        "</head>" &_
        "<body style='background-color: white'>" &_
          "<div class='container-fluid'>" &_
            "<div class='text-center' style='color:gray'>" &_
              "<br>" &_
              "<br>" &_
              "<br>" &_
              "<br>" &_
              "<br>" &_
              "<img src='http://" & request.ServerVariables("server_name") & "/corporativo/images/SWlogo500w.png' class='customer-logo'>" &_
              "<br>" &_
              "<span class='nome'>SWnet</span>" &_
              "<br>" &_
              "<br>" &_
              "<h2 style='color:red'>ATEN&Ccedil;&Atilde;O!" &_
              "<br>" &_
              "Você n&atilde;o tem permiss&atilde;o para acessar esta funç&atilde;o." &_
              "<br>" &_
              "<br>" &_
              "Clique <a href='#' onclick='history.go(-1)'>aqui</a> para ir para a p&aacute;gina de login." &_
            "</div>" &_
          "</div>" &_
        "</body>" &_
      "</html>"
    response.end()      

  else

      Texto = "Liberado"

  end if

  PermiteAcesso = Texto

end function
' -------------------------------------------------------------------------------------------------
function AnoDoisDigitos(data)

  if Trim(data) <> "" then

    MesDia = Mid(data,1,5)
    Ano = Right(data,2)
    AnoDoisDigitos = MesDia&"/"&Ano

  else

    AnoDoisDigitos = ""

  end if

end function
' -------------------------------------------------------------------------------------------------
function LimitaTexto(texto,limite)

  NovoTexto = Trim(texto)

  if Len(NovoTexto) > limite then
    NovoTexto = Mid(NovoTexto,1,limite-3)&"..."
  end if

  LimitaTexto = NovoTexto

end function
' -------------------------------------------------------------------------------------------------
function ping(hostname)

  set wshShell = CreateObject("WScript.Shell")
  ping = Not CBool(wshShell.run("ping -n 1 " & hostname,0,True))

end function
' -------------------------------------------------------------------------------------------------
public function Cripto(senha)

  Dim s, calc 
  calc= Senha * 123456
  Cripto = calc 

end function 
' -------------------------------------------------------------------------------------------------
function ArredondarCima(valor,decimais)

  'response.write "<p>Valor: " & valor
  'response.write "<br>decimais: " & decimais

  if (decimais = 0) then

    'nao se desejam decimais, posso devolver a parte inteira
    'porem se o número tem decimais, tenho que arredondar por cima
    if (int(valor)-valor < 0) then
      'havia decimais no valor, devolvo o valor inteiro + 1
      ArredondarCima = int(valor) + 1
    else
      'no havia decimais
      ArredondarCima = valor
    end if

  else

    'foi dado um valor de decimais e há que arredondar com esses decimais
    arredondamento = round(valor, decimais) 
    'se o arredondamento é por baixo ou faço por cima
    if (valor - arredondamento) > 0 then 
      arredondamento = arredondamento + (1/(calcula_base_expoente(10,decimais)))
      arredondamento = round (arredondamento,decimais)
    end if

    arredondamento_cima=arredondamento 

  end if

end function 
' -------------------------------------------------------------------------------------------------
public function DesfazCripto(senha) 

  Dim s, calc 
  calc= Senha / 123456
  senha = Right("000000" & calc , 6)
  DesfazCripto = senha

end function 
' -------------------------------------------------------------------------------------------------
public function MesDescricaoResumida(data) 

  select case Mid(Data,4,2)
  case "01" MesDescricaoResumida = "Jan"
  case "02" MesDescricaoResumida = "Fev"
  case "03" MesDescricaoResumida = "Mar"
  case "04" MesDescricaoResumida = "Abr"
  case "05" MesDescricaoResumida = "Mai"
  case "06" MesDescricaoResumida = "Jun"
  case "07" MesDescricaoResumida = "Jul"
  case "08" MesDescricaoResumida = "Ago"
  case "09" MesDescricaoResumida = "Set"
  case "10" MesDescricaoResumida = "Out"
  case "11" MesDescricaoResumida = "Nov"
  case "12" MesDescricaoResumida = "Dez"
  end select

end function 
' -------------------------------------------------------------------------------------------------
public function MesDescricaoResumidaEnglish(data) 

  select case Mid(Data,4,2)
  case "01" MesDescricaoResumidaEnglish = "Jan"
  case "02" MesDescricaoResumidaEnglish = "Feb"
  case "03" MesDescricaoResumidaEnglish = "Mar"
  case "04" MesDescricaoResumidaEnglish = "Apr"
  case "05" MesDescricaoResumidaEnglish = "May"
  case "06" MesDescricaoResumidaEnglish = "Jun"
  case "07" MesDescricaoResumidaEnglish = "Jul"
  case "08" MesDescricaoResumidaEnglish = "Aug"
  case "09" MesDescricaoResumidaEnglish = "Sep"
  case "10" MesDescricaoResumidaEnglish = "Oct"
  case "11" MesDescricaoResumidaEnglish = "Nov"
  case "12" MesDescricaoResumidaEnglish = "Dec"
  end select

end function 
' -------------------------------------------------------------------------------------------------
public function MesDescricaoDetalhada(mes) 

  select case mes
  case "01" MesDescricaoDetalhada = "Janeiro"
  case "02" MesDescricaoDetalhada = "Fevereiro"
  case "03" MesDescricaoDetalhada = "Marco"
  case "04" MesDescricaoDetalhada = "Abril"
  case "05" MesDescricaoDetalhada = "Maio"
  case "06" MesDescricaoDetalhada = "Junho"
  case "07" MesDescricaoDetalhada = "Julho"
  case "08" MesDescricaoDetalhada = "Agosto"
  case "09" MesDescricaoDetalhada = "Setembro"
  case "10" MesDescricaoDetalhada = "Outubro"
  case "11" MesDescricaoDetalhada = "Novembro"
  case "12" MesDescricaoDetalhada = "Dezembro"
  end select

end function 
' -------------------------------------------------------------------------------------------------
public function MesDescricaoDetalhadaEnglish(data) 

  select case Mid(data,4,2)
  case "01" MesDescricaoDetalhadaEnglish = "January"
  case "02" MesDescricaoDetalhadaEnglish = "February"
  case "03" MesDescricaoDetalhadaEnglish = "March"
  case "04" MesDescricaoDetalhadaEnglish = "April"
  case "05" MesDescricaoDetalhadaEnglish = "May"
  case "06" MesDescricaoDetalhadaEnglish = "June"
  case "07" MesDescricaoDetalhadaEnglish = "July"
  case "08" MesDescricaoDetalhadaEnglish = "August"
  case "09" MesDescricaoDetalhadaEnglish = "September"
  case "10" MesDescricaoDetalhadaEnglish = "October"
  case "11" MesDescricaoDetalhadaEnglish = "November"
  case "12" MesDescricaoDetalhadaEnglish = "December"  
  end select

end function 
' -------------------------------------------------------------------------------------------------
public function VerificaZeroNegativo(numero) 

  if (numero * -1) = 0 then
    VerificaZeroNegativo = numero * -1
  else
    VerificaZeroNegativo = numero
  end if

end function 
' -------------------------------------------------------------------------------------------------
public function SesseaoExpirada()

  RESPONSE.WRITE "<script>SairInpanet()</script>" &_
   "<head>" &_
     "<style>" &_
       "body.center-form {" &_
         "min-height: 100vh;" &_
       "}" &_
       "div.center-form {" &_
         "position: relative;" &_
         "min-height: 100vh;" &_
         "margin-top: 10%;" &_
       "}" &_
       "div.center-form > img {" &_
         "display: block;" &_
         "margin-left: auto;" &_
         "margin-right: auto;" &_
       "}" &_
       "div.center-form > h1 {" &_
         "position: absolute;" &_
         "top: 50%;" &_
         "left: 50%;" &_
         "text-align: center;" &_
         "transform: translateY(-50%) translateX(-50%);" &_
         "color: #333333;" &_
       "}"   &_
       "div.center-form > h4 {" &_
         "position: absolute;" &_
         "top: 50%;" &_
         "left: 50%;" &_
         "text-align: center;" &_
         "transform: translateY(-50%) translateX(-50%);" &_
         "color: #333333;" &_
       "}" &_
     "</style>" &_
   "</head>" &_
   "<body class='center-form' style='background-color:#FFCC00'>" &_
     "<div class='center-form'>" &_
       "<img src='../images/expirou.png' class='center-form'/>" &_
       "</br>" &_
       "<h1>Seu tempo expirou.<br><br>Efetue novamente o login<br>para utilizar nossos servi&ccedil;os.</h1>" &_
     "</div>" &_
   "</body>"

  Response.End()

end function
' -------------------------------------------------------------------------------------------------
public function ValidaPermissaoFormularios(url,matricula,filial)
       urlPermissao = replace(url,"/ambienteteste/corporativo/","../")
       urlPermissao = replace(urlPermissao,"/corporativo/","../")

        SQL = "SELECT COUNT(*) EXISTE "
        SQL = SQL + " FROM HELPDESK..LINKS MENU "
        SQL = SQL + " INNER JOIN HELPDESK..MONTAMENU PERMISSAO ON MENU.LIN_COD = PERMISSAO.LIN_COD AND ISNULL(MENU.LIN_D_E_L_E_T,'') = '' "
        SQL = SQL + " WHERE LIN_URL = '" + urlPermissao + "' "
        SQL = SQL + " AND CRM_COD = '" + matricula + "' "
        SQL = SQL + " AND FILIAL = '" + filial + "' "
        SQL = SQL + " AND ISNULL(MON_D_E_L_E_T,'') = '' "
        SQL = SQL + " AND CRM_COD <> '' "
        SQL = SQL + " AND FILIAL <> '' " 

        Dados.Open(SQL)
        Dados.Close
end function
' -------------------------------------------------------------------------------------------------
public function ValidaPermissaoRelatorios(url,matricula,filial)

  urlPermissao = replace(url,"/ambienteteste/corporativo/","../")
  urlPermissao = replace(urlPermissao,"/corporativo/","../")

  if url = "../financeiro/HistoricoSolicitacaocreditoCliente" then
    'urlPermissao = "../../inpanet/dptos/financeiro/consultar_limite_credito1"
    sqlAdcional = " " &_
      "OR 1 = (" &_
        "SELECT " &_
          "COUNT(*) EXISTE " &_
        "FROM HELPDESK..LINKS MENU " &_
          "INNER JOIN HELPDESK..MONTAMENU PERMISSAO ON MENU.LIN_COD = PERMISSAO.LIN_COD AND ISNULL(MENU.LIN_D_E_L_E_T,'') = '' " &_
        "WHERE LIN_URL = '../../corporativo2/index.php/diversos/limitecredito/listasolicitacoes/' " &_
          "AND CRM_COD = '" & matricula & "' " &_
          "AND FILIAL = '" & filial & "' " &_
          "AND ISNULL(MON_D_E_L_E_T,'') = '' " &_
          "AND CRM_COD <> '' " &_
          "AND FILIAL <> '' ))"
  else
    sqlAdcional = ")"
  end if

  SQL = "" &_
    "SELECT " &_
      "COUNT(*) EXISTE " &_
    "FROM HELPDESK..RELATORIOSCAPA MENU " &_
    "INNER JOIN HELPDESK..MONTARELAT PERMISSAO ON " &_
      "MENU.REL_COD = PERMISSAO.REL_COD AND " &_
      "ISNULL(MENU.REL_D_E_L_E_T_E_D,'') = '' " &_
    "WHERE CRM_COD = '" & matricula & "' " &_
      "AND FILIAL = '" & filial & "' " &_
      "AND ISNULL(MON_D_E_L_E_T_E_D,'') = '' " &_
      "AND CRM_COD <> '' " &_
      "AND FILIAL <> '' " &_
      "AND (REL_URL = '" & urlPermissao & "' " & sqlAdcional

  Dados.Open(SQL)
  Dados.Close

end function
' -------------------------------------------------------------------------------------------------
'Função que remove todos os acentos e caracteres especiais do texto
function RemoveAcentos(strTexto)

  Dim strComAcentos
  Dim strSemAcentos
  Dim strResultado
	Dim intCont
  'Conjunto de Caracteres com acentos
  strComAcentos = "ÁÍÓÚÉÄÏÖÜËÀÌÒÙÈÃÕÂÎÔÛÊáíóúéäïöüëàìòùèãõâîôûêÇç&°º"
  'Conjunto de Caracteres sem acentos
  strSemAcentos = "AIOUEAIOUEAIOUEAOAIOUEaioueaioueaioueaoaioueCcE_:"
  strEspeciais = "!@#$%¨&*()+=§ªº{}[]:;<>?/°|\¢¬£'^~`´"
  intCont = 0
  strResultado = strTexto
  Do While intCont < Len(strComAcentos)
    intCont = intCont + 1
    strResultado = Replace(strResultado, Mid(strComAcentos, intCont, 1), Mid(strSemAcentos, intCont, 1))
  Loop
  intCont = 0
  Do While intCont < Len(strEspeciais)
    intCont = intCont + 1
    strResultado = Replace(strResultado, Mid(strEspeciais,intCont,1), "")
  Loop

  RemoveAcentos = strResultado

end function
' função iif - if em linha ------------------------------------------------------------------------
' parametros: condicao   - expressão lógica
' retorno:    verdadeiro - valor que será retornado, caso a condicao seja verdadeira
'             falso      - valor que será retornado, caso a condicao seja falso
' -------------------------------------------------------------------------------------------------
function iif(condicao, verdadeiro, falso)

  if condicao then
    iif = verdadeiro
  else
    iif = falso
  end if

end function
' função inArray ----------------------------------------------------------------------------------
' parametros: aValor - o parametro que será pesquisado
'             aArray - matriz contém os valores que seram comparados
' retorno   : true   - se encontrar o aValor
'             false  - senão encontrar o aValor
' ------------------------------------------------------------------------------------------------- 
function inArray( aValor, aArray )

  dim p, retorno
  retorno = False
  for p = 0 to UBound(aArray)
    if aArray(p) = aValor then
      retorno = True
      exit for
    end if
  next

  inArray = retorno

end function
' função formataDataYYYYMMDD ----------------------------------------------------------------------------------
' parametros: data - data (recebe a data no formato YYYY-MM-DD)
' retorno   : a data no formato YYYYMMDD
' ------------------------------------------------------------------------------------------------- 
function formataDataYYYYMMDD( data )

  dim dia, mes, ano, ret
  dia = ""
  mes = ""
  ano = ""
  ret = ""
  
  if (trim(data) <> "") then

    if (len( data ) = 10) then
      ano = mid( data, 1, 4 )
      mes = mid( data, 6, 2 )
      dia = mid( data, 9, 2 )
      ret = ano & mes & dia
    end if

  else

    ret = ""
    
  end if

  formataDataYYYYMMDD = ret

end function
' função formataDataUS ----------------------------------------------------------------------------
' parametros: data - data (recebe a data no formato DD/MM/YYYY)
' retorno   : a data no formato YYYYMMDD
' -------------------------------------------------------------------------------------------------
function formataDataUS( data )

  dim dia, mes, ano, ret
  dia = ""
  mes = ""
  ano = ""
  ret = ""
  
  if (trim(data) <> "") then

    if (len( data ) = 10) then
      ano = mid( data, 7, 4 )
      mes = mid( data, 4, 2 )
      dia = mid( data, 1, 2 )
      ret = ano & "-" & mes & "-" & dia
    end if

  else

    ret = ""
    
  end if

  formataDataUS = ret

end function
' função formataDataUS ----------------------------------------------------------------------------
' parametros: data - data (recebe a data no formato DD/MM/YYYY)
' retorno   : a data no formato YYYYMMDD
' -------------------------------------------------------------------------------------------------
function formataDataBR( data )

  dim dia, mes, ano, ret
  dia = ""
  mes = ""
  ano = ""
  ret = ""
  
  if (trim(data) <> "") then

    if (len( data ) = 10) then
      ano = mid( data, 7, 4 )
      mes = mid( data, 4, 2 )
      dia = mid( data, 1, 2 )
      ret = dia & "/" & mes & "/" & ano
    end if

  else

    ret = ""
    
  end if

  formataDataBR = ret

end function
' -------------------------------------------------------------------------------------------------
function datetime(pDate, pHour)

  Dim dd, mm, yy, hh, nn, ss
  Dim datevalue, timevalue, dtsnow, dtsvalue
  on error resume next

  if len(pDate) > 10 then
    dtsnow = pDate 
  else
    dtsnow = pDate & " " & pHour
  end if

  dd = Right("00" & Day(dtsnow), 2)
  mm = Right("00" & Month(dtsnow), 2)
  yy = Year(dtsnow)

  hh = Right("00" & Hour(dtsnow), 2)
  nn = Right("00" & Minute(dtsnow), 2)
  ss = Right("00" & Second(dtsnow), 2)

  datevalue = yy & "-" & mm & "-" & dd
  timevalue = hh & ":" & nn & ":" & ss
  dtsvalue = datevalue & " " & timevalue

  if err.number <> 0 then
    datetime = "01/01/1900 00:00:00.000"
		'response.write err.number & "<br>"
		'response.write err.description & "<br>"
  else
    datetime = dtsvalue
	end if

  on error goto 0

end function
' -------------------------------------------------------------------------------------------------
' Sub OnTransactionAbort()

'   Dim objErr, msgBody, msgAssunto
'   Set objErr = Server.GetLastError()

'   msgBody = "<b>ASP - código...: </b>" & objErr.ASPCode & "<br />" &_
'             "<b>ASP - descrição: </b>" & objErr.ASPDescription & "<br />" &_
'             "<b>Categoria......: </b>" & objErr.Category & "<br />" &_
'             "<b>Coluna..,......: </b>" & objErr.Column & "<br />" &_
'             "<b>Descrição......: </b>" & objErr.Description & "<br />" &_
'             "<b>Arquivo........: </b>" & objErr.File & "<br />" &_
'             "<b>Linha..........: </b>" & objErr.Line & "<br />" &_
'             "<b>Número.........: </b>" & objErr.Number & "<br />" &_
'             "<b>Fonte..........: </b>" & objErr.Source & "<br />" 

'   msgAssunto = "SKnet - TRATAMENTO DE ERROS - " + FormatDateTime( Now(), 3 )

'   CALL EnviaEmail( "jorge.queiroz@smurftkapparock.com.br","","","","", msgBody, msgAssunto )

' End Sub 
' -------------------------------------------------------------------------------------------------
public function HandleError(err, source, complement, query)

  DIM msgBody, msgAssunto

  Set objASPError = Server.GetLastError()

  if IsEmpty(query) then
    query = ""
  end if

  msgBody = "<p style='font-family: Courier'>" &_
            "<b>Usuario.....: </b>" & CODUSER & " - " & USR & "<br>" &_
            "<b>Servidor....: </b>" & Request.ServerVariables("server_name") & "<br>" &_
            "<b>Fonte.......: </b>" & source & "<br>" &_           
            "<b>Erro........: </b>" & " (0x" & Hex(err.Number) & ")" & "<br>" &_
            "<b>Descricao...: </b>" & err.Description & "<br>" &_
            "<b>Complemento.: </b>" & complement & "<br>" &_
            "</p>" &_
            "<br>" &_
            "<br>" &_
            query

  msgAssunto = "SWnet - TRATAMENTO DE ERROS - " + FormatDateTime( Now(), 3 )

  call EnviaEmail( EMAIL, , , , "", msgBody, msgAssunto )

''  call EnviaEmail( "jorge.queiroz@smurfitkappa.com.br", "leone.souza@smurfitkappa.com.br", "carlos.souza@smurfitkappa.com.br", "thiago.pereira@smurfitwestrock.com.br", "", msgBody, msgAssunto )

end function
' -------------------------------------------------------------------------------------------------
public function GetFileName()

  DIM files, url, segments

  'get then current url from the server variables
  url = Request.ServerVariables("path_info")

  segments = split(url,"/")

  'read the last segment
  url = segments(ubound(segments))
  GetFileName = url

end function 
' -------------------------------------------------------------------------------------------------
public function CriaConexao(PosicaoConexao)

  conexaoString=""
  Retorno = ""
  x = 0
  y = 0
  if PosicaoConexao < 0 or PosicaoConexao > 22 then
    SESSION("conexaoString")=""
  else
    Dim objFSO, ObjAbreArq, campo, parametros(50,4), connect(60), ArrayLinha
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set ObjAbreArq = objFSO.OpenTextFile("c:\inetpub\wwwroot\corporativo\cnn\conexao.txt")

    while not ObjAbreArq.AtEndOfStream
      linha = ObjAbreArq.ReadLine
      ArrayLinha = split(linha, ";")
      for y = 0 to 4
          parametros(x,y) = ArrayLinha(y)
      next
      x = x + 1
    wend

    x = 0
    for x=0 to Ubound(parametros) 
      connect(x) = "Provider="&parametros(x,0)&";Data Source="&parametros(x,1)&";Initial Catalog="&parametros(x,2)&";User ID="&base64_decode(parametros(x,3))&";Password="&base64_decode(parametros(x,4))&";"
    next

    if connect(PosicaoConexao) <> "" then
      SESSION("conexaoString")=connect(PosicaoConexao) 
    else
      SESSION("conexaoString")=""
    end if 
  end if 

end function
'--------------------------------------------------------------------------------------------------
function TirarAcento(Palavra) 

  CAcento = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
  SAcento = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
  Texto = ""

  if Palavra <> "" then

    for X = 1 to Len(Palavra)

      Letra = Mid(Palavra,X,1)
      Pos_Acento = InStr(CAcento,Letra)

      if Pos_Acento > 0 then 
        Letra = mid(SAcento,Pos_Acento,1)
      end if

      Texto = Texto & Letra

    next

    TirarAcento = Texto
  end if

end function
'--------------------------------------------------------------------------------------------------
function TrocarAcento(Palavra)

  CAcento = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
  Texto = ""
  if Palavra <> "" then
    for X = 1 to Len(Palavra)
      Letra = Mid(Palavra,X,1)
      Pos_Acento = InStr(CAcento,Letra)
      if Pos_Acento > 0 then
        Letra = "_"
      end if
      Texto = Texto & Letra
    Next
    TrocarAcento = Texto
  end if

end function
'--------------------------------------------------------------------------------------------------
function fNomePlanta(planta)
        select case planta
              case "00" fNomePlanta = "PIRAPETINGA"
              case "04" fNomePlanta = "UBERABA"
              case "05" fNomePlanta = "BENTO"
              case "06" fNomePlanta = "FORTALEZA"
              case "07" fNomePlanta = "SAQUAREMA"
              case "08" fNomePlanta = "ARACATUBA"
              case "09" fNomePlanta = "BLUMENAU"
              case "10" fNomePlanta = "PACAJUS"
              case "11" fNomePlanta = "PORTO FELIZ"
              case "12" fNomePlanta = "TRES BARRAS"
              case "13" fNomePlanta = "FLORESTAL"
        end select
end Function

'-------------------------- função responsavel em trasformar , em . -------------------------------
Function trataValores(argValor)
         if argValor = "" then 
            trataValores = 0
            exit function 
         end if

         trataValores = replace(argValor,".",",")
end function 

'-------------------------- função responsavel em verificar se o BD está sendo atualizado -------------------------------
function atualizandoBD()
    SQL = "SELECT TOP 1 VALIDACAO FROM BI_ATUALIZACAO"
    SET SPO_VALIDACAO = Conecta( "SPO-BI", SQL )
    if not isObject(SPO_VALIDACAO) then
      response.end
    end if

    atualizandoBD = false 
    if not SPO_VALIDACAO.eof then
       atualizandoBD = SPO_VALIDACAO("VALIDACAO")
    end if 
    SPO_VALIDACAO.close()
    SET SPO_VALIDACAO = nothing
end function     

' função formataDataSAP ----------------------------------------------------------------------------------
' parametros: data - data (recebe a data no formato YYYY-MM-DD)
' retorno   : a data no formato DD-MM-YYYY
' ------------------------------------------------------------------------------------------------- 
function formataDataSAP( data )
  dim dia, mes, ano, ret
  dia = ""
  mes = ""
  ano = ""
  ret = ""
  
  if (trim(data) <> "") then
    if (len( data ) = 10) then
      ano = mid( data, 1, 4 )
      mes = mid( data, 6, 2 )
      dia = mid( data, 9, 2 )
      ret = dia & "-" & mes & "-" & ano
    end if
  end if

  formataDataSAP = ret

end function

Function EmailValido(strEmail)
    Dim objRegExp, ret
    Set objRegExp = New RegExp
    
    objRegExp.Pattern = "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$"
    objRegExp.IgnoreCase = True
    objRegExp.Global = False
    
    ret = objRegExp.Test(strEmail)
    Set objRegExp = Nothing
    
    EmailValido = ret
End Function
' Função para converter uma data em string para um formato especificado
' Parâmetros:
'   dataStr: String representando a data (ex: "2023-10-05", "2023/10/05", "20231005")
'   formatoRetorno: String indicando o formato desejado (ex: "DD/MM/AAAA", "DD-MM-AAAA", etc.)
' Retorno: String com a data no formato solicitado, ou "Formato inválido" se o formato não for reconhecido
function ConverterData(dataStr, formatoRetorno)

  Dim dt
  Dim ano, mes, dia

  ' Verificar o formato da data de entrada
  if len(trim(dataStr)) = 8 and instr(dataStr, "-") = 0 and instr(dataStr, "/") = 0 then

    ' Formato AAAAMMDD
    ano = CInt(Left(dataStr, 4))
    mes = CInt(Mid(dataStr, 5, 2))
    dia = CInt(Right(dataStr, 2))
    dt = DateSerial(ano, mes, dia)

  elseif len(trim(dataStr)) = 10 and (instr(dataStr, "-") > 0 or instr(dataStr, "/") > 0) then

    ' Formato AAAA-MM-DD ou AAAA/MM/DD
    dt = CDate(dataStr)

  else

    ' Formato não reconhecido
    ConverterData = dataStr
    exit function

  end if

  ' Formatar a data de acordo com o formatoRetorno
  Dim result

  select case formatoRetorno
  case "DD/MM/AAAA" : result = Right("0" & Day(dt), 2) & "/" & Right("0" & Month(dt), 2) & "/" & Year(dt)
  case "DD-MM-AAAA" : result = Right("0" & Day(dt), 2) & "-" & Right("0" & Month(dt), 2) & "-" & Year(dt)
  case "DDMMAAAA"   : result = Right("0" & Day(dt), 2) & Right("0" & Month(dt), 2) & Year(dt)
  case "AAAA/MM/DD" : result = Year(dt) & "/" & Right("0" & Month(dt), 2) & "/" & Right("0" & Day(dt), 2)
  case "AAAA-MM-DD" : result = Year(dt) & "-" & Right("0" & Month(dt), 2) & "-" & Right("0" & Day(dt), 2)
  case "AAAAMMDD"   : result = Year(dt) & Right("0" & Month(dt), 2) & Right("0" & Day(dt), 2)
  case else         : result = "Formato de retorno inválido"
  end select

  ConverterData = result

end function

Function ObterNomePlantas(plantasJSON)
    If IsNull(plantasJSON) OR plantasJSON = "" OR plantasJSON = "[]" Then
        ObterNomePlantas = ""
        Exit Function
    End If
    
    Dim plantasStr
    plantasStr = Replace(plantasJSON, "[", "")
    plantasStr = Replace(plantasStr, "]", "")
    plantasStr = Replace(plantasStr, """", "")
    
    Dim plantasArray
    plantasArray = Split(plantasStr, ",")
    
    Dim i, codigo, nomes()
    ReDim nomes(UBound(plantasArray))
    
    For i = 0 To UBound(plantasArray)
        codigo = Trim(plantasArray(i))
        Select Case codigo
            Case "00"
                nomes(i) = "Pirapetinga"
            Case "01"
                nomes(i) = "Corporativo RJ"
            Case "02"
                nomes(i) = "Corporativo SP"
            Case "03"
                nomes(i) = "Corporativo"
            Case "04"
                nomes(i) = "Uberaba"
            Case "05"
                nomes(i) = "Bento"
            Case "06"
                nomes(i) = "Maranguape"
            Case "07"
                nomes(i) = "Saquarema"
            Case "08"
                nomes(i) = "Aracatuba"
            Case "09"
                nomes(i) = "Blumenau"
            Case "10"
                nomes(i) = "Pacajus"
            Case "11"
                nomes(i) = "Porto Feliz"
            Case "12"
                nomes(i) = "Três Barras"
            Case "13"
                nomes(i) = "Florestal"
            Case Else
                nomes(i) = "Desconhecido"
        End Select
    Next
    
    ObterNomePlantas = Join(nomes, ", ")
End Function

Function ObterUrlBase()
    Dim protocol, host, appPath
    
    If Request.ServerVariables("HTTPS") = "on" Then
        protocol = "https://"
    Else
        protocol = "http://"
    End If
    
    host = Request.ServerVariables("HTTP_HOST")
    
    appPath = Request.ServerVariables("APPL_MD_PATH")
    appPath = Replace(appPath, "/LM/W3SVC/1/ROOT", "")
    
    ObterUrlBase = protocol & host & appPath
End Function

Function FormatarData(data)
    If IsNull(data) Or data = "" Then
        FormatarData = ""
        Exit Function
    End If
    
    On Error Resume Next
    Dim dataFormatada
    dataFormatada = CDate(data)
    
    If Err.Number <> 0 Then
        FormatarData = ""
        Exit Function
    End If
    
    FormatarData = Right("0" & Day(dataFormatada), 2) & "/" & _
                             Right("0" & Month(dataFormatada), 2) & "/" & _
                             Year(dataFormatada)
    On Error Goto 0
End Function

Function ExtrairValorJSON(jsonStr, chave)
    Dim padrao, regex, matches
    Set regex = New RegExp
    regex.Pattern = """" & chave & """\s*:\s*""([^""]*)""|'" & chave & "'\s*:\s*'([^']*)'"
    regex.IgnoreCase = True
    regex.Global = False
    
    Set matches = regex.Execute(jsonStr)
    If matches.Count > 0 Then
        If matches(0).SubMatches(0) <> "" Then
            ExtrairValorJSON = matches(0).SubMatches(0)
        Else
            ExtrairValorJSON = matches(0).SubMatches(1)
        End If
    Else
        ExtrairValorJSON = ""
    End If
End Function


Function EscapeJSON(str)
    If isnull(str) Then
        str = ""
    End If

    Dim result
    result = str
    resultado = Replace(resultado, """", "")
    resultado = Replace(resultado, "\", "\\")
    resultado = Replace(resultado, "/", "")
    resultado = Replace(resultado, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    result = trim(result)
    EscapeJSON = result
End Function

Function ConverterNumeroParaCalcular(valor)
    On Error Resume Next
    Dim valorLimpo
    
    If IsEmpty(valor) Or IsNull(valor) Or valor = "" Then
        ConverterNumeroParaCalcular = 0
        Exit Function
    End If
    
    valorLimpo = CStr(valor)
    valorLimpo = Trim(valorLimpo)
    
    valorLimpo = Replace(valorLimpo, " ", "")

    If SESSION("Ambiente") = "HOMOLOGACAO" Then
        valorLimpo = Replace(valorLimpo, ".", ",")
    Else 
        valorLimpo = Replace(valorLimpo, ",", ".")
    End If
    
    ConverterNumeroParaCalcular = CDbl(valorLimpo)
    
    If Err.Number <> 0 Then
        ConverterNumeroParaCalcular = 0
        Err.Clear
    End If
    
    On Error Goto 0
End Function

Function FormatarMoeda(valor)
    If IsNull(valor) Or valor = "" Then
        FormatarMoeda = "0,00"
        Exit Function
    End If
    
    Dim valorNumerico
    valorNumerico = CDbl(valor)
    
    Dim valorFormatado
    valorFormatado = FormatNumber(valorNumerico, 2, -1, 0, -1)
    
    FormatarMoeda = valorFormatado
End Function

Function JsonToDictionary(jsonString)
    Dim dict, temp, pairs, i, pair, key, value, inObject, objStart, objEnd
    Dim currentPos, char, braceCount, objectContent
    
    Set dict = Server.CreateObject("Scripting.Dictionary")
    
    temp = Trim(jsonString)
    
    If Left(temp, 1) = "{" Then temp = Mid(temp, 2)
    If Right(temp, 1) = "}" Then temp = Left(temp, Len(temp) - 1)
    
    currentPos = 1
    inObject = False
    braceCount = 0
    
    Do While currentPos <= Len(temp)
        Dim keyStart, keyEnd, colonPos, valueStart, valueEnd
        
        Do While currentPos <= Len(temp) And (Mid(temp, currentPos, 1) = " " Or Mid(temp, currentPos, 1) = vbTab Or Mid(temp, currentPos, 1) = vbCr Or Mid(temp, currentPos, 1) = vbLf)
            currentPos = currentPos + 1
        Loop
        
        If currentPos > Len(temp) Then Exit Do
        
        If Mid(temp, currentPos, 1) = """" Then
            keyStart = currentPos + 1
            keyEnd = InStr(keyStart, temp, """")
            key = Mid(temp, keyStart, keyEnd - keyStart)
            currentPos = keyEnd + 1
        Else
            Exit Do
        End If
        
        Do While currentPos <= Len(temp) And Mid(temp, currentPos, 1) <> ":"
            currentPos = currentPos + 1
        Loop
        currentPos = currentPos + 1 ' Pula o ":"
        
        Do While currentPos <= Len(temp) And (Mid(temp, currentPos, 1) = " " Or Mid(temp, currentPos, 1) = vbTab)
            currentPos = currentPos + 1
        Loop
        
        char = Mid(temp, currentPos, 1)
        
        If char = "{" Then
            objStart = currentPos
            braceCount = 1
            currentPos = currentPos + 1
            
            Do While currentPos <= Len(temp) And braceCount > 0
                char = Mid(temp, currentPos, 1)
                If char = "{" Then
                    braceCount = braceCount + 1
                ElseIf char = "}" Then
                    braceCount = braceCount - 1
                End If
                currentPos = currentPos + 1
            Loop
            
            objEnd = currentPos - 1
            objectContent = Mid(temp, objStart, objEnd - objStart + 1)
            
            Set dict.Item(key) = JsonToDictionary(objectContent)
            
        ElseIf char = """" Then
            valueStart = currentPos + 1
            valueEnd = InStr(valueStart, temp, """")
            value = Mid(temp, valueStart, valueEnd - valueStart)
              
            If NOT dict.Exists(key) Then
              dict.Add key, value
            End If
            currentPos = valueEnd + 1
            
        Else
            valueStart = currentPos
            valueEnd = InStr(currentPos, temp, ",")
            
            If valueEnd = 0 Then
                valueEnd = Len(temp) + 1
            End If
            
            value = Trim(Mid(temp, valueStart, valueEnd - valueStart))
            
            If IsNumeric(Replace(value, ".", ",")) Then
                Dim valueForConversion
                
                If InStr(value, ".") > 0 Then
                    valueForConversion = Replace(value, ".", ",")
                    value = CDbl(valueForConversion)
                Else
                    value = CLng(value)
                End If
            ElseIf LCase(value) = "true" Then
                value = True
            ElseIf LCase(value) = "false" Then
                value = False
            ElseIf LCase(value) = "null" Then
                value = Null
            End If
            
            dict.Add key, value
            currentPos = valueEnd
        End If
        
        Do While currentPos <= Len(temp) And (Mid(temp, currentPos, 1) = "," Or Mid(temp, currentPos, 1) = " " Or Mid(temp, currentPos, 1) = vbTab Or Mid(temp, currentPos, 1) = vbCr Or Mid(temp, currentPos, 1) = vbLf)
            currentPos = currentPos + 1
        Loop
    Loop
    
    Set JsonToDictionary = dict
End Function
%>