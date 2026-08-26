<%
With Response
 .Buffer = True
 .Expires = -1
 .ExpiresAbsolute = Now() - 1
 .CacheControl = "no-cache; private; no-store; must-revalidate; max-stale=0; post-check=0; pre-check=0; max-age=0"
 .AddHeader "pragma", "no-store"
 .AddHeader "pragma", "no-cache"
 .AddHeader "cache-control", "no-cache; private; no-store; must-revalidate; max-stale=0; post-check=0; pre-check=0; max-age=0"
End With

%>
<!--#Include File = "../cnn2/cnn2.asp"-->
<!--#Include File = "../funcoes/funcoes.asp"-->
<%
if trim(request("produto")) <> "" then

  SQL = "EXEC SP_CALCULAPESO '" & trim(request("produto")) & "'"

  set Consulta = Conecta( "PIR-DADOSADV", SQL )
  if not isobject(Consulta) then
         response.end
  end if

  peso = iif( (Not Consulta.Eof), Consulta("NPESO") * trim(request("quantidade")), 0 )

  Consulta.Close()

  SQL = "" &_
    "SELECT " &_
      "Condicao, " &_
      "TipoFrete, " &_
      "Preco, " &_
      "E4_DESCRI " &_
    "FROM CORPORATIVOHMG..Proposta " &_
    "INNER JOIN SE4010 ON E4_CODIGO = Condicao COLLATE SQL_Latin1_General_CP437_BIN AND SE4010.D_E_L_E_T_ = '' " &_
    "WHERE E4_MSBLQL <> '1' " &_
      "AND Ativa = 1 " &_
      "AND Produto = '"&trim(request("produto"))&"' " &_
      "AND Cliente = '"&trim(request("cliente"))&"' " &_
      "AND Loja = '"&trim(request("loja"))&"' " &_
      "AND Quantidade <= "&replace(replace(peso,".",""),",",".") &" " &_
      "AND De <= '"&InverteData(Trim(Request("dataEntrega")))&"' " &_
      "AND Ate >= '"&InverteData(Trim(Request("dataEntrega")))&"' "

  set Consulta = Conecta( "PIR-DADOSADV", SQL )
  if not isobject(Consulta) then
         response.end
  end if

  if not Consulta.Eof then

    tipoFrete = iif( (trim(Consulta("TipoFrete")) = "C"), "CIF", "FOB" )

    response.Write "|"&Trim(Request("tipo"))&"|"&trim(Consulta("Condicao"))&"|"&trim(Consulta("TipoFrete"))&"|"&trim(Consulta("Preco"))&"|"&trim(Consulta("E4_DESCRI"))&"|"&tipoFrete

  else

    response.Write "ErroNEncontrado|"&trim(request("tipo"))&"|"

  end if

  Consulta.Close()

else

  response.Write "ErroProposta"
  
end if
%>
