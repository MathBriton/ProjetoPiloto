<!-- #include file = "../../cnn2/cnn2.asp" -->
<!-- #include file = "../../funcoes/funcoes.asp" -->

<%
'--------------------------------------------------------------------------------------------------'
function SentencaSQL(aFil, aCli, aLoj)

  if inArray(aFil,array("05","06")) then

    aSub = ",( " &_
              "SELECT " &_
                "A1_NOME " &_
              "FROM SA1010 " &_
              "WHERE ( A1_COD LIKE CONCAT(L.Cliente,'%') ) OR ( A1_CDSIADN LIKE CONCAT(L.Cliente,'%') ) ) as NomeCliente "

  end if

  SQL = "" &_
    "SELECT " &_
      "L.Cliente " &_
      ",L.Limite " &_
      ",L.Saldo as SALDO_CREDITO " &_
      iif( inArray(aFil,array("05","06")), argSub + " /* a filial 06 não o nome do cliente, por isto o subselect para buscar */ ", "" ) &_
      iif( aFil = "07", ",L.Cliente as NomeCliente /* a filial 07 ja tem o nome do cliente, por isto o if coloca no select */ ", "" ) &_
    "FROM LimiteCredito as L " &_
    "WHERE 1 = 1 " &_
      iif( inArray( aFil, array("05","06") ), "AND L.Cliente = '" & aCli & "' ", "" ) &_
      iif( aFil = "07", "AND L.CODCLI = '" & aCli & "' AND LOJACLI = '" & aLoj & "' ", "" )

  SentencaSQL = SQL

end function 
'--------------------------------------------------------------------------------------------------'

select case request("filial")
case "00" ' PIRAPETINGA -------------------------------------------------------------------------

  SQL = "EXEC SP_CONSULTA_CREDITO_CLIENTE '"& request("cliente") & "', '" & codsiga & "'"
  SET CONSULTA = Conecta( "PIR-DADOSADV", SQL )

case "05" ' BENTO GONÇALVES ------------------------------------------------------------

  SQL = "EXEC [spLimiteCredito] '" & request("cliente") & "'"
  SET CONSULTA = Conecta( "BEN-DADOSADV", SQL )

  if (CONSULTA IS NOTHING) then
    response.end
  end if

  SQL = SentencaSQL( request("filial"), request("cliente"), "" )

  SET CONSULTA = Conecta( "BEN-DADOSADV", SQL )  

case "06" ' MARANGUAPE --------------------------------------------------------------------------

  SQL = "EXEC [spLimiteCredito] '" & request("cliente") & "'"
  SET CONSULTA = Conecta( "FOR-DADOSADV", SQL )

  if (CONSULTA IS NOTHING) then
    response.end
  end if

  SQL = SentencaSQL( request("filial"), request("cliente"), "" )

  SET CONSULTA = Conecta( "FOR-DADOSADV", SQL )

case "07" ' SAQUAREMA ---------------------------------------------------------------------------

  SQL = "EXEC [spLimiteCredito] '" & request("cliente") & "', '" & request("loja") & "'"
  SET CONSULTA = Conecta( "SAQ-DADOSADV", SQL )

  if (CONSULTA IS NOTHING) then
    response.end
  end if

  SQL = SentencaSQL( request("filial"), request("cliente"), request("loja") )

  SET CONSULTA = Conecta( "SAQ-DADOSADV", SQL )

end select

if not isObject(CONSULTA) or (CONSULTA IS NOTHING) then
  response.end
end if

if not CONSULTA.EOF then

  LIMITE  = CONSULTA("Limite")
  ABERTO  = iif( request("filial") = "00", FormatNumber(CONSULTA("LIMITE"),2), FormatNumber(CONSULTA("Limite"),2) ) - FormatNumber(CONSULTA("SALDO_CREDITO"),2)
  CREDITO = CONSULTA("SALDO_CREDITO")

  Limite  = iif( (Limite <> ""), FormatNumber(Limite,2), "" )
  Aberto  = iif( (Aberto <> ""), FormatNumber(Aberto,2), "" )
  Credito = iif( (Credito <> ""), FormatNumber(Credito,2), "" )

  CONSULTA.Close()
  SET CONSULTA = NOTHING

  response.write "|LimiteCredito|" & Limite & "|" & Aberto & "|" & Credito

else

  CONSULTA.Close()
  SET CONSULTA = NOTHING

  response.write "ErroNaoEncontrado"

end if

%>