<!--#Include File = "../cnn/DADOSADV.asp"-->


<%
if trim(request("cliente")) <> "" and trim(request("loja")) <> "" and trim(request("clienteremessa")) <> "" and trim(request("lojaremessa")) <> "" then

  if trim(request("codsiga")) = "" and trim(request("nivel")) <> "0"  then
    response.write "ErroLogof"
    response.End
  end if

  if trim(request("nivel")) <> "" then

    SQL = "SELECT TOP 1 A1_NFFAT " & _
          "FROM SA1010 " & _
          "WHERE A1_COD = '" & Trim(request("cliente")) & "' " & _
            "AND D_E_L_E_T_ <> '*' " & _
            "AND A1_MSBLQL <> '1' "

    if trim(request("tipo")) <> "Grupo" then
      SQL = SQL & "AND A1_LOJA = '" & trim(request("loja")) & "' "
    end if

'    if trim(request("nivel")) = "1" and trim(request("codsiga")) <> "24" then

    if trim(request("nivel")) = "1" then
      SQL = SQL & "AND A1_VEND = '"& trim(request("codsiga")) & "' "
    end if

    Dados.Open(SQL)

    if not Dados.Eof and Trim(Dados("A1_NFFAT")) <> "" then

      Remessa = Split(Trim(Dados("A1_NFFAT")),"-")
      existe = 0

      for each valor in Remessa 

        if Trim(valor) = Trim(request("clienteremessa")) then
          existe = 1
        end if

      next 

      if existe = 1 then

        SQL = "SELECT TOP 1 A1_NOME " & _
              "FROM SA1010 " & _
              "WHERE A1_COD = '" & trim(request("clienteremessa")) & "' " & _
                "AND D_E_L_E_T_ <> '*' " & _
                "AND A1_MSBLQL <> '1' "

        if trim(request("tipo")) <> "Grupo" then
          SQL = SQL & "AND A1_LOJA = '" & trim(request("lojaremessa")) & "' "
        end if

        Dados2.Open(SQL)

        if not Dados2.Eof then
          response.Write "|"&trim(request("tipo"))&"|"&trim(Dados2("A1_NOME"))
        else
          response.Write "ErroNEncontrado|"&trim(request("tipo"))
        end if

      else

        response.Write "ErroNEncontrado|"&trim(request("tipo"))

      end if

    else

      response.Write "ErroNEncontrado|"&trim(request("tipo"))

    end if

  else

    response.write "ErroLogof|"&trim(request("tipo"))

  end if

else

  response.Write "ErroLogof"

end if
%>
