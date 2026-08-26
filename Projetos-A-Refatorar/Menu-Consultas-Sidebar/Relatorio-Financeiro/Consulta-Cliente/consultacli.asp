<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
  <head>
  <title>Consulta Cliente</title>
  <link rel="stylesheet" type="text/css" href="css/style.css"/>
    <style>
      body    { font-family: verdana, sans-serif; font-size: 80%; }
      #global { font-family: verdana, sans-serif; font-size: 80%; }
      td      { font-family: verdana, sans-serif; font-size: 80%; }
      th      { font-family: verdana, sans-serif; font-size: 80%; } 
    </style>
  </head>

  <script>
      /*----------------------------------------------------------------------------
      Formatação para qualquer mascara
      -----------------------------------------------------------------------------*/
      function formatar(src, mask) {
          var i = src.value.length;
          var saida = mask.substring(0, 1);
          var texto = mask.substring(i)
          if (texto.substring(0, 1) != saida) {
              src.value += texto.substring(0, 1);
          }
      }
  </script>

  <!--#Include File="cnn/cnn.asp"-->

  <body>
    <form method="post" action="consultacli.asp" />
      <table align="center">
        <tr>
          <td>Digite o nome do cliente ou parte dele ou o c�digo:</td>
          <td><input type="text" name="cli" id="cli" value="<%=Trim(Request("cli")) %>" /></td>
          <td><input type="submit" name="consulta" id="consulta" value="consultar"/></td>
        </tr>
      </table>

      <br/>
  <%
      function formatacnpj(num)
          um  = mid(num,1,2) 
          dois= mid(num,3,3)
          tres= mid(num,6,3)
          barra = mid(num,9,4)
          dig = mid(num,13,2)
          formatacnpj= um&"."&dois&"."&tres&"/"&barra&"-"&dig
      end function

      function FormatCep(cep)
          FormatCep = left(cep,5)+"-"+right(cep,3)
      end function

      function formatacpf(num)
          um  = mid(num,1,3) 
          dois= mid(num,4,3)
          tres= mid(num,7,3)
          dig = mid(num,10,2)
          formatacpf= um&"."&dois&"."&tres&"-"&dig
      end function

      if request("cli")<>"" then
              SQL = "SELECT "
        SQL = SQL &   "A1_LOJA, "
        SQL = SQL &   "A1_COD, "
        SQL = SQL &   "A1_NOME, "
        SQL = SQL &   "A1_CGC, "
        SQL = SQL &   "A1_MUN, "
        SQL = SQL &   "A1_END, "
        SQL = SQL &   "A1_CEP, "
        SQL = SQL &   "A1_RISCO, "
        SQL = SQL &   "CASE WHEN A1_MSBLQL = '1' THEN 'SIM' ELSE 'NAO' END AS BLOQUEADO "
        SQL = SQL & "FROM "
        SQL = SQL &   "SA1010 "
        SQL = SQL & "WHERE "
        SQL = SQL &   "UPPER(A1_NOME) LIKE UPPER('%"&request("cli")&"%') OR "
        SQL = SQL &   "A1_COD = '"&Trim(request("cli"))&"' "
        SQL = SQL & "ORDER BY "
        SQL = SQL &  "A1_NOME, A1_LOJA"

        'RESPONSE.WRITE SQL

        Dados.Open(SQL)
  %>
  <%
        if Dados.eof then
  %>
      <p></p>
      <br/>
      <table style="border:2 solid red; background-color:Yellow" align ="center">
        <tr>
          <td><font collor="red"> Cliente não encontrado!!! </font></td>
        </tr>
      </table>    
  <%
          response.End
        end if
  %>
      <table align="center" cellspacing ="0" cellpadding="4">
        <tr style="background-color:Gray">
          <th align="center" style="border-bottom:1 solid black;border-top:1 solid black;">C�d.</th>
          <th align="center" style="border-bottom:1 solid black;border-top:1 solid black;">Loja</th>
          <th align="center" style="border-bottom:1 solid black;border-top:1 solid black;">Nome</th>
          <th align="center" style="border-bottom:1 solid black;border-top:1 solid black;">Endere�o</th>
          <th align="center" style="border-bottom:1 solid black;border-top:1 solid black;">Munic�pio</th>
          <th align="center" style="border-bottom:1 solid black;border-top:1 solid black;">CEP</th>
          <th align="center" style="border-bottom:1 solid black;border-top:1 solid black;">CNPJ</th>
          <th align="center" style="border-bottom:1 solid black;border-top:1 solid black;">Bloqueado</th>
          <th align="center" style="border-bottom:1 solid black;border-top:1 solid black;">Risco</th>
        </tr>
  <%
        while not Dados.eof 
          cnpj = Dados("A1_CGC")
          if Estilo = "LinhaImpar" then
            Estilo = "LinhaPar"
          else
            Estilo = "LinhaImpar"
          end If
  %>
      
        <tr>    
          <td align="right" style="border-bottom:1 solid black;" class="<%=Estilo%>">&nbsp;&nbsp;<%=Dados("A1_COD")%></td>
          <td align="right" style="border-bottom:1 solid black;" class="<%=Estilo%>">&nbsp;&nbsp;<%=Dados("A1_LOJA")%>&nbsp;&nbsp;</td>
          <td align="left" style="border-bottom:1 solid black;" class="<%=Estilo%>">&nbsp;&nbsp;<%=Dados("A1_NOME")%>&nbsp;&nbsp;</td>
          <td align="left" style="border-bottom:1 solid black;" class="<%=Estilo%>">&nbsp;&nbsp;<%=Dados("A1_END")%>&nbsp;&nbsp;</td>
          <td align="left" style="border-bottom:1 solid black;" class="<%=Estilo%>">&nbsp;&nbsp;<%=Dados("A1_MUN")%>&nbsp;&nbsp;</td>
          <td align="left" style="border-bottom:1 solid black;" class="<%=Estilo%>">&nbsp;&nbsp;<%=FormatCep(Dados("A1_CEP"))%>&nbsp;&nbsp;</td>
          <td align="right" style="border-bottom:1 solid black;" class="<%=Estilo%>">
            &nbsp;&nbsp;
  <%
          if len(trim(cnpj)) = 11 then 
            response.write formatacpf(cnpj)
          else
            response.write formatacnpj(cnpj)
          end if        
  %>
            &nbsp;&nbsp;
          </td>
          <td align="left" style="border-bottom:1 solid black;" class="<%=Estilo%>">&nbsp;&nbsp;<%=Dados("BLOQUEADO")%>&nbsp;&nbsp;</td>        
          <td align="left" style="border-bottom:1 solid black;" class="<%=Estilo%>">&nbsp;&nbsp;<%=Dados("A1_RISCO")%>&nbsp;&nbsp;</td>
        </tr>
  <%  
          Dados.MoveNext
        wend
  %>
      </table>
  <%
    end if
  %>
    </form>
  </body>
</html>
