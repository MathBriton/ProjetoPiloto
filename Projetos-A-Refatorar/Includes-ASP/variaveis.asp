<%
	Dim Base64Chars
	Base64Chars =	"ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
								"abcdefghijklmnopqrstuvwxyz" & _
								"0123456789" & _
								"+/"

	' Functions for encoding string to Base64
	Public Function base64_encode( byVal strIn )
		Dim c1, c2, c3, w1, w2, w3, w4, n, strOut
		For n = 1 To Len( strIn ) Step 3
			c1 = Asc( Mid( strIn, n, 1 ) )
			c2 = Asc( Mid( strIn, n + 1, 1 ) + Chr(0) )
			c3 = Asc( Mid( strIn, n + 2, 1 ) + Chr(0) )
			w1 = Int( c1 / 4 ) : w2 = ( c1 And 3 ) * 16 + Int( c2 / 16 )
			If Len( strIn ) >= n + 1 Then 
				w3 = ( c2 And 15 ) * 4 + Int( c3 / 64 ) 
			Else 
				w3 = -1
			End If
			If Len( strIn ) >= n + 2 Then 
				w4 = c3 And 63 
			Else 
				w4 = -1
			End If
			strOut = strOut + mimeencode( w1 ) + mimeencode( w2 ) + _
					  mimeencode( w3 ) + mimeencode( w4 )
		Next
		base64_encode = strOut
	End Function

	Private Function mimeencode( byVal intIn )
		If intIn >= 0 Then 
			mimeencode = Mid( Base64Chars, intIn + 1, 1 ) 
		Else 
			mimeencode = ""
		End If
	End Function	

	' Function to decode string from Base64
	Public Function base64_decode( byVal strIn )
		Dim w1, w2, w3, w4, n, strOut
		For n = 1 To Len( strIn ) Step 4
			w1 = mimedecode( Mid( strIn, n, 1 ) )
			w2 = mimedecode( Mid( strIn, n + 1, 1 ) )
			w3 = mimedecode( Mid( strIn, n + 2, 1 ) )
			w4 = mimedecode( Mid( strIn, n + 3, 1 ) )
			If w2 >= 0 Then _
				strOut = strOut + _
					Chr( ( ( w1 * 4 + Int( w2 / 16 ) ) And 255 ) )
			If w3 >= 0 Then _
				strOut = strOut + _
					Chr( ( ( w2 * 16 + Int( w3 / 4 ) ) And 255 ) )
			If w4 >= 0 Then _
				strOut = strOut + _
					Chr( ( ( w3 * 64 + w4 ) And 255 ) )
		Next
		base64_decode = strOut
	End Function

	Private Function mimedecode( byVal strIn )
		If Len( strIn ) = 0 Then 
			mimedecode = -1 : Exit Function
		Else
			mimedecode = InStr( Base64Chars, strIn ) - 1
		End If
	End Function

'DECLARANDO VARI�VEIS DESTA P�GINA
DIM NOME, TRATAMENTO, CODTRATAMENTO, DATACESSO, HORACESSO, ACESSOS, CODUSER, UNIDADE_FILIAL, VER_FILIAL, DATEXPIRASENHA, CODIGOHELPDESK, ValidaOrcamentoCusto,AdmSistemas
DIM USR, MSGANIVERSARIO, SENHAEXPIRADA, EMAIL, CATCODIGO, NIVEL, VALARTNIVEL, CATEGORIA, PERMISSAO, PESOMENOR, PRECOMENOR, VALINFCAD, FININFCAD, VERTODASFNCS
DIM ASSINATURA, ALTERAPO, ALTERAPAPEL, NIVELALTERAPEDIDO, VALCLIENTE, CRMTIPO, FORNECEDOR, LOJA,CODUSUPROTHEUS, validaPreCadastroProduto, AprovaOrcamentoPreCadastroProduto,CEDAG,E2OPEN
DIM controlesInternos,controladoria,sapLaboratorista,sapClicherista,IdUsuario,NivelFinanceiro,NivelFiscal

'RECEBENDO AS VARI�VEIS

AdmSistemas = SESSION("Adm")

NOME = Trim(SESSION("NOME"))
If NOME = "" Then NOME = base64_decode(Request.Cookies("INPANET_NOME")) End If

DATEXPIRASENHA = SESSION("DATEXPIRASENHA")
If DATEXPIRASENHA = "" Then DATEXPIRASENHA = base64_decode(Request.Cookies("INPANET_DATEXPIRASENHA")) End If

VER_FILIAL = SESSION("VER_FILIAL")
If VER_FILIAL = "" Then VER_FILIAL = base64_decode(Request.Cookies("INPANET_VER_FILIAL")) End If

UNIDADE_FILIAL = SESSION("FILIAL")
If UNIDADE_FILIAL = "" Then UNIDADE_FILIAL = base64_decode(Request.Cookies("INPANET_FILIAL")) End If

TRATAMENTO = SESSION("TRATAMENTO")
If TRATAMENTO = "" Then TRATAMENTO = base64_decode(Request.Cookies("INPANET_TRATAMENTO")) End If

CODTRATAMENTO = SESSION("CODTRATAMENTO")
If CODTRATAMENTO = "" Then CODTRATAMENTO = base64_decode(Request.Cookies("INPANET_CODTRATAMENTO")) End If

IF CODTRATAMENTO < 2 THEN
	TRATAMENTO = ""
END IF

CODUSUPROTHEUS = SESSION("CODUSUPROTHEUS")

IdUsuario = SESSION("IdUsuario") 
validaPreCadastroProduto = SESSION("ValidaPreCadastroProduto") 
AprovaOrcamentoPreCadastroProduto = SESSION("AprovaOrcamentoPreCadastroProduto")
ValidaOrcamentoCusto = SESSION("ValidaOrcamentoCusto") 

CEDAG = SESSION("CEDAG")
E2OPEN = SESSION("E2OPEN") 
controlesInternos = SESSION("ControlesInternos")
NivelFinanceiro= SESSION("NivelFinanceiro")
NivelFiscal= SESSION("NivelFiscal")
controladoria = SESSION("Controladoria")

DATACESSO = SESSION("DATACESSO")
HORACESSO = SESSION("HORACESSO")
ACESSOS = SESSION("ACESSOS")

PESOMENOR = SESSION("PESOMENOR")
If PESOMENOR = "" Then PESOMENOR = base64_decode(Request.Cookies("INPANET_PESOMENOR")) End If

PRECOMENOR = SESSION("PRECOMENOR")
If PRECOMENOR = "" Then PRECOMENOR = base64_decode(Request.Cookies("INPANET_PRECOMENOR")) End If

VALINFCAD = SESSION("VALINFCAD")
If VALINFCAD = "" Then VALINFCAD = base64_decode(Request.Cookies("INPANET_VALINFCAD")) End If

FININFCAD = SESSION("FININFCAD")
If FININFCAD = "" Then FININFCAD = base64_decode(Request.Cookies("INPANET_FININFCAD")) End If

VERTODASFNCS = SESSION("VERTODASFNCS")
If VERTODASFNCS = "" Then VERTODASFNCS = base64_decode(Request.Cookies("INPANET_VERTODASFNCS")) End If

CODUSER = Trim(SESSION("CODUSER"))
If CODUSER = "" Then CODUSER = base64_decode(Request.Cookies("INPANET_CODUSER")) End If

CODIGOHELPDESK = Trim(SESSION("CODIGO"))
If CODIGOHELPDESK = "" Then CODIGOHELPDESK = base64_decode(Request.Cookies("INPANET_CODIGO")) End If

USR = Trim(SESSION("USR"))
If USR = "" Then USR = base64_decode(Request.Cookies("INPANET_USR")) End If

MSGANIVERSARIO = SESSION("MSGANIVERSARIO")
'SENHAEXPIRADA = SESSION("SENHAEXPIRADA")

EMAIL = SESSION("EMAIL")
If EMAIL = "" Then EMAIL = base64_decode(Request.Cookies("INPANET_EMAIL")) End If

CATCODIGO = SESSION("CATCODIGO")
If CATCODIGO = "" Then CATCODIGO = base64_decode(Request.Cookies("INPANET_CATCODIGO")) End If

NIVEL = SESSION("NIVEL")
If NIVEL = "" Then NIVEL = base64_decode(Request.Cookies("INPANET_NIVEL")) End If

ALTERAPO = SESSION("ALTERAPO")
If ALTERAPO = "" Then ALTERAPO = base64_decode(Request.Cookies("INPANET_ALTERAPO")) End If

ALTERAPAPEL = SESSION("ALTERAPAPEL")
If ALTERAPAPEL = "" Then ALTERAPAPEL = base64_decode(Request.Cookies("INPANET_ALTERAPAPEL")) End If

NIVELALTERAPEDIDO = SESSION("NIVELALTERAPEDIDO")
If NIVELALTERAPEDIDO = "" Then NIVELALTERAPEDIDO = base64_decode(Request.Cookies("INPANET_NIVELALTERAPEDIDO")) End If

VALARTNIVEL = SESSION("VALARTNIVEL")
If VALARTNIVEL = "" Then VALARTNIVEL = base64_decode(Request.Cookies("INPANET_VALARTNIVEL")) End If

CATEGORIA = SESSION("CATEGORIA")
If CATEGORIA = "" Then CATEGORIA = base64_decode(Request.Cookies("INPANET_CATEGORIA")) End If

PERMISSAO = SESSION("PERMISSAO")
If PERMISSAO = "" Then PERMISSAO = base64_decode(Request.Cookies("INPANET_PERMISSAO")) End If

ASSINATURA = SESSION("ASSINATURA")
If ASSINATURA = "" Then ASSINATURA = base64_decode(Request.Cookies("INPANET_ASSINATURA")) End If

CRMTIPO = SESSION("TIPO")
If CRMTIPO = "" Then CRMTIPO = base64_decode(Request.Cookies("INPANET_TIPO")) End If

FORNECEDOR = SESSION("FORNECEDOR")
If FORNECEDOR = "" Then FORNECEDOR = base64_decode(Request.Cookies("INPANET_FORNECEDOR")) End If

LOJA = SESSION("LOJA")
If LOJA = "" Then LOJA = base64_decode(Request.Cookies("INPANET_LOJA")) End If
'RESPONSE.WRITE CRM_TIPO&"\<BR />"
'PASSANDO AS VARI�VEIS

SESSION("Adm") = AdmSistemas
SESSION("IdUsuario") = IdUsuario
SESSION("NOME") = NOME
SESSION("VER_FILIAL") = VER_FILIAL
SESSION("FILIAL") = UNIDADE_FILIAL
SESSION("TRATAMENTO") = TRATAMENTO
SESSION("DATACESSO") = DATACESSO
SESSION("HORACESSO") = HORACESSO
SESSION("ACESSOS") = ACESSOS
SESSION("PESOMENOR") = PESOMENOR
SESSION("PRECOMENOR") = PRECOMENOR
SESSION("VALINFCAD") = VALINFCAD
SESSION("FININFCAD") = FININFCAD 
SESSION("VERTODASFNCS") = VERTODASFNCS
SESSION("CODUSER") = CODUSER
SESSION("CODIGO") = CODIGOHELPDESK
SESSION("USR") = USR
SESSION("MSGANIVERSARIO") = MSGANIVERSARIO
'SESSION("SENHAEXPIRADA") = SENHAEXPIRADA
SESSION("EMAIL") = EMAIL
SESSION("CATCODIGO") = CATCODIGO
SESSION("NIVEL") = NIVEL
SESSION("ALTERAPO") = ALTERAPO
SESSION("ALTERAPAPEL") = ALTERAPAPEL
SESSION("NIVELALTERAPEDIDO") = NIVELALTERAPEDIDO
SESSION("VALARTNIVEL") = VALARTNIVEL
SESSION("CATEGORIA") = CATEGORIA
SESSION("PERMISSAO") = PERMISSAO
SESSION("ASSINATURA") = ASSINATURA
SESSION("DATEXPIRASENHA") = DATEXPIRASENHA 
SESSION("TIPO") = CRMTIPO 
SESSION("FORNECEDOR") = FORNECEDOR 
SESSION("LOJA") = LOJA 
SESSION("CODUSUPROTHEUS") = CODUSUPROTHEUS 
SESSION("ValidaPreCadastroProduto") = validaPreCadastroProduto
SESSION("AprovaOrcamentoPreCadastroProduto") = AprovaOrcamentoPreCadastroProduto
SESSION("ValidaOrcamentoCusto") = ValidaOrcamentoCusto
SESSION("CEDAG") = CEDAG 
SESSION("E2OPEN")= E2OPEN 
SESSION("ControlesInternos") = controlesInternos
SESSION("Controladoria") = controladoria
SESSION("NivelFinanceiro") = NivelFinanceiro
SESSION("NivelFiscal") = NivelFiscal
SESSION("Ambiente") = "HOMOLOGACAO"

'SE O TIMEOUT ESPIRAR N�O PERMITE MAIS ACESSO

if Trim(SESSION("NOME")) = "" OR Trim(SESSION("CODUSER")) = "" OR Trim(SESSION("USR")) = "" or Trim(SESSION("TIPO")) = "" then
%>
  <script type="text/javascript">
    function sair() {
      window.open("http://" + location.hostname, "_top");
    }
  </script>
<%
	response.write "" &_
    "<html lang='en-US' xmlns='http://www.w3.org/1999/xhtml'>" &_
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
            "<img src='http://" & request.ServerVariables("server_name") & "/corporativo/images/SWlogo500w.png' class='customer-logo'>" &_
            "<br>" &_
            "<span class='nome'>SWnet</span>" &_
            "<br>" &_
            "<i class='fa fa-hourglass-end' aria-hidden='true' style='font-size:150px;color: #333333'></i>" &_
            "<br>" &_
            "<br>" &_
            "<h1 style='color: red'>Seu tempo expirou no sistema!</h1>" &_
            "<br>" &_
            "<br>" &_
            "Efetue novamente o login para utilizar nossos serviços." &_
            "<br>" &_
            "<br>" &_
            "Clique <a href='#' onclick='sair()'>aqui</a> para ir para a página de login" &_
          "</div>" &_
        "</div>" &_
      "</body>" &_
    "</html>"
	response.end
end if

%>