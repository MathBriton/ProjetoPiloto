<!DOCTYPE html>
<%
  if request("excel") = "1"  then
     response.ContentType="application/vnd.ms-excel"
     response.AddHeader "content-disposition", "inline; filename=RelatorioComissao.xls"
  end if    
%>

<!-- v2.0 Data: 11-12-20-->
<!-- Descrições de alterações realizadas podem ser encontradas no Chpl 179 -->
<html lang="pt-br" xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=2"/>
    <!--#Include File="../../corporativo/cnn/cnn.asp"-->
    <!--#Include File="../../corporativo/funcoes/funcoes.asp"-->
    <!--#Include File="../../corporativo/includes/variaveis.asp"-->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />
    <script type="text/javascript" src="../../corporativo/funcoes/funcoesJAVASCRIPT.js"></script>
    <title>Relatório de comissões</title>
    <style>
        .alert, table{
            width:99% !important;
            margin-left:1%;
        }
        th{
            text-align:center;
        }
        .historico{
            font-size:80%
        }
        .table td, .table th{
            padding:5px;
        }
        td{
            font-size: small
        }
    </style>
   
</head>

<%
'call ValidaPermissaoRelatorios("../../intranet/comissao/parametros",SESSION("CODUSER"),SESSION("FILIAL"))

session.LCID = 1046
filial = Trim(Request("Filial"))
if nivel = 0 then
  vendedor = Trim(request("vendedor"))
else
  vendedor = coduser
end if
parametro = Request("hidParametros")
if parametro = "" then
    parametro = " "
end if
if right(parametro,1) = "," then
    parametro = Left(parametro,len(parametro)-1)    
end if 
parametro = replace(parametro,"|","','") 

if len(Request("DataIni")) = 7 then
    dataIni = "01/" & Trim(Request("DataIni"))
else
    dataIni = Trim(Request("DataIni"))
end if    
if len(Request("dataFim")) = 7 then
    dataFim = UltimoDia(Trim(Request("DataFim")))
else
    dataFim = Trim(Request("dataFim"))
end if   
cliente = Trim(Request("ClieIni"))

if Cdate(Trim(Request("DataIni"))) <= Cdate("30/04/2023") then
  msg = "Para acessar informações anteriores a 05/2023, utilize a  opção Comissão!"
end if

'if Nivel > 0 then
'        SQL = "SELECT "
'  SQL = SQL &   "CRM_COD "
'  SQL = SQL & "FROM "
'  SQL = SQL &   "HELPDESK..CRM_DTI "
'  SQL = SQL & "WHERE "
'  SQL = SQL &   "CRM_COD = '"&CodUser&"' AND "
'  SQL = SQL &   "CRM_ASSINATURA = '"&Trim(Request("assinatura"))&"' " 
'  Dados.Open(SQL)
'  if Dados.Eof then
'    msg = "Representante não encontrado!"
'  end if
'  Dados.Close()
'end if

'      SQL = "SELECT "
'SQL = SQL &   "A3_COD, "
'SQL = SQL &   "A3_NOME AS A3_NREDUZ "
'SQL = SQL & "FROM "
'SQL = SQL &   "SA3010 "
'SQL = SQL & "WHERE "
'SQL = SQL &   "A3_COD = '" & Vendedor & "' AND "
'SQL = SQL &   "D_E_L_E_T_ <> '*' AND "
'SQL = SQL &   "A3_CALCOMI <> 'N' "
Planta = ""
Select Case Parametro 
       case "BPRC','BPRM" 
             Planta = "BPRC"
       case "BPCE" 
             Planta = "BPCE"             
       case "BPIC','BPIM','BPIW" 
             Planta = "BPIC"
       case "BSQC" 
             Planta = "BSQC"            
       case "BUBC','BUBM" 
             Planta = "BPIC"            
End Select        
if Planta = "" then
    msg = "Escolha a Planta novamente! " 
end if

SQL = "SELECT ZC1_CODORI, ZC1_NOME, ZC1_PLANTA,ZC1_GRUPO "
SQL = SQL & "FROM ZC1010  "
SQL = SQL & "WHERE ZC1_CODORI =  '" &  Vendedor  & "' AND "
SQL = SQL & "ZC1_GRUPO = '" & Planta  &"'"
'response.write sql 
'response.end 
Dados.Open(SQL)
if Dados.Eof then
 if msg = "" then  msg = "Representante não encontrado! "  
end if

%>
<body>
    <div class="container container-fluid">
        <%if msg <> "" then %>
            <div class="alert alert-warning">
              <strong>Atenção!</strong> <%=msg %>
              <a href="Parametroscmn.asp" class="btn btn-info" title="Voltar parâmetros"><i class="fas fa-undo"></i></a>
            </div>
            <%response.end %>
        <%else %>
        <table class="table table-striped table-bordered">
            
            <tr>
                <td class="text-center"><a href="Parametroscmn.asp" class="btn btn-info" title="Voltar parâmetros"><i class="fas fa-undo"></i></a></td>
                <td class="text-center" colspan="5"><h2><strong>Relatório de Comissões - <%=filialDescricao %></strong></h2></td>
                <td class="text-center align-middle">
                    <a href="relatorioImpressaocmn.asp?hidParametros=<%=parametro %>&vendedor=<%=vendedor %>&DataIni=<%=dataini %>&DataFim=<%=datafim %>&Cliente=<%=cliente %>&assinatura=<%=Trim(Request("assinatura")) %>" target="_blank" title="Imprimir" class="btn btn-light">
                        <i class="fas fa-print"></i>
                    </a>
                </td>
                <td class="text-center align-middle">
                      <a href="relatoriocmn.asp?Excel=1&hidParametros=<%=parametro %>&vendedor=<%=vendedor %>&DataIni=<%=dataini %>&DataFim=<%=datafim %>&ClieIni=<%=cliente %>&assinatura=<%=Trim(Request("assinatura")) %>"  title="Exportar Excel - Relatorio"  class="btn btn-light">
                      <i class="fas fa-file"></i>
                  </a>
                </td>
                <td class="text-center align-middle">
                  <a href="relatorioExcelLinhacmn.asp?Excel=1&hidParametros=<%=parametro %>&vendedor=<%=vendedor %>&DataIni=<%=dataini %>&DataFim=<%=datafim %>&Cliente=<%=cliente %>&assinatura=<%=Trim(Request("assinatura")) %>"  title="Exportar Excel em Linhas"  class="btn btn-light">
                  <i class="fas fa-file"></i>
              </a>
            </td>

            </tr>
            <tr>
                <td  class="text-center" colspan="10">
			        Titulos baixados entre 
                    <b><%=dataIni%></b> e <b><%=dataFim%></b> 
                    - Vendedor <b><%=vendedor%>-<%=Mid(Dados("ZC1_NOME"),1,20)%></b>
                    - Plantas  <b><%=replace(Parametro,"','"," - ")%></b>
                    <%if cliente <> "" then%>
                        - Cliente: <b><%=cliente%></b>
                    <% end if%>

		        </td>
            </tr>
<%
            Dados.close() 
            if Nivel = 0 then
            SQL = "SELECT ZC9_CODREP "
      SQL = SQL & "FROM ZC9010 "
      SQL = SQL & "WHERE ZC9_EMISSA BETWEEN '" & InverteData(dataIni) & "' AND '" & InverteData(dataFim) & "' and ZC9_CODREP not in (24,200) "
      SQL = SQL &   "AND ZC9_PLANTA IN ('" & Parametro & "') "
      SQL = SQL & "GROUP BY ZC9_CODREP ORDER BY CONVERT(INT, ZC9_CODREP) "
      Dados2.Open(SQL)  
      if not Dados2.eof then
%>            
		        <tr>
			        <td  colspan="10" class="text-center">
                <%
                while Not(Dados2.EOF)%>
                <a href="relatoriocmn.asp?hidParametros=<%=parametro %>&vendedor=<%=Dados2("ZC9_CODREP") %>&DataIni=<%=dataini %>&DataFim=<%=datafim %>&ClieIni=<%=cliente %>&assinatura=<%=Trim(Request("assinatura")) %>"  title="Exportar Excel - Relatorio"  class="btn btn-light">                  
                <%=Trim(Dados2("ZC9_CODREP"))%>
                </a>
<%
                  Dados2.MoveNext
                Wend
%>                
			        </td>
		        </tr>
<%
      end if 
      Dados2.close() 
      end if
                      SQL = "SELECT *, "
            SQL = SQL & " CASE ZC9_ORIGEM WHEN 'A' then 'Modelo Antigo' WHEN 'T' then 'Modelo Novo' WHEN 'P' then 'Percentual Padrão' ELSE 'Outros' END ZC9_ORIGEM, "
            SQL = SQL & " CASE ZC9_PROC WHEN 'C' then 'Automatico' WHEN 'M' then 'Manual' ELSE 'Outros' END ZC9_PROC "
            SQL = SQL & "FROM ZC9010 "
            SQL = SQL & "WHERE "
            SQL = SQL &       "ZC9_PGTO >= '" & InverteData(dataIni) & "' "
            SQL = SQL &   "AND ZC9_PGTO <= '" & InverteData(dataFim) & "' "
            SQL = SQL &   "AND ZC9_CODREP = '" & vendedor & "' "
            SQL = SQL &   "AND ZC9_CODREP <> '24' "
            SQL = SQL &   "AND ZC9_PCOMIS < 100 AND D_E_L_E_T_ <> '*' "
            SQL = SQL &   "AND ZC9_PLANTA IN ('" & Parametro & "') " 
            if cliente <> "" then
	            SQL = SQL & "AND ZC9_CODCLI = '" & Cliente & "'  "
            end if       
            SQL = SQL & "ORDER BY ZC9_NOMCLI, ZC9_PGTO "                
            'response.write  SQL 
            'response.end  
            Dados.open(SQL)   
%>
		    <tr style="line-height:1px;">
			    <td colspan="10" >&nbsp</td>
		    </tr>
		    <tr>
			    <td class="text-center" colspan="10"><b>Comissões liberadas durante o período</b></td>
		    </tr>
<%
            QtdeTitulos = 0
            TotalIPI = 0
            TotalICMS = 0
            TotalCOFINS = 0
            TotalPIS = 0
            TotBasLib = 0
            TotLib = 0
            TotalIRRF = 0
            NumLib = 0
            While Not(Dados.EOF)
                          
%>
<tr style="border: 2px; margin: 5px; background-color: rgba(117, 199, 128, 0.032);">
  <td colspan="7" style="border: 3px solid rgb(141, 138, 138);" width="100%">
    <table class="table table-striped table-bordered">
      <tr>
        <th>Razão Social</th>
        <th>Cliente</th>
        <th>Pagamento</th>
        <th>Emissão</th>
        <th>Título</th>
        <th>Grupo</th>
        <th>Vlr Titulo</th>
        <th>Base</th>
        <th>Perc.</th>
        <th>Comissão</th>
      </tr>
      <tr>
        <td><%=Dados("ZC9_NOMCLI")%></td>
        <td><%=Dados("ZC9_CODCLI")%>|<%=Dados("ZC9_LOJCLI")%></td>
        <td><%=AcertaData(Dados("ZC9_PGTO"))%></td>
        <td><%=AcertaData(Dados("ZC9_EMISSA"))%></td>
        <td><%=Dados("ZC9_DOC")&"/"&Dados("ZC9_SERIE")&"/"&Dados("ZC9_PARCEL")%></td>
        <td><%=Dados("ZC9_GRUPO")%></td>
        <td><%=FormatNumber(Dados("ZC9_VLBASE"),2)%></td>
        <td><%=FormatNumber(Dados("ZC9_VLNET"),2)%></td>
        <td><%=FormatNumber(Dados("ZC9_PCOMIS"),2)%></td>
        <td><%=FormatNumber(Dados("ZC9_VLCOM"),2)%></td>
      </tr>
    </table>
    <table class="table table-striped table-bordered">
      <tr class="warning">
        <th>P. ICMS</th>
        <th>V. ICMS</th>
        <th>P. IPI</th>
        <th>V. IPI.</th>
        <th>P. COFINS</th>
        <th>V. COFINS</th>
        <th>P. PIS</th>
        <th>V. PIS</th>
        <th>V. IRRF</th>
      </tr>

      <tr class="warning">
        <td class="text-right"><%=FormatNumber(Dados("ZC9_PICMS"),2)%>&nbsp;</td>
        <td class="text-right"><%=FormatNumber(Dados("ZC9_VICMS"),2)%>&nbsp;</td>
        <td class="text-right"><%=FormatNumber(Dados("ZC9_PIPI"),2)%>&nbsp;</td>
        <td class="text-right"><%=FormatNumber(Dados("ZC9_VIPI"),2)%>&nbsp;</td>
        <td class="text-right"><%=FormatNumber(Dados("ZC9_PCOFIN"),2)%>&nbsp;</td>
        <td class="text-right"><%=FormatNumber(Dados("ZC9_VCOFIN"),2)%>&nbsp;</td>
        <td class="text-right"><%=FormatNumber(Dados("ZC9_PPIS"),2)%>&nbsp;</td>
        <td class="text-right"><%=FormatNumber(Dados("ZC9_VPIS"),2)%>&nbsp;</td>
        <td class="text-right"><%=FormatNumber(Dados("ZC9_IRRF"),2)%>&nbsp;</td>
      </tr>
    </table>
    <table class="table table-striped table-bordered">
      <tr class="info">
        <th>VL. KG NET</th>
        <th>DIF. DIA</th>
        <th>ORIGEM</th>
        <th>ID. TAB</th>
        <th>CALCULO</th>
        <th>OBSERVAÇÃO</th>
        <th>ARQUIVO ORIGEM</th>
      </tr>

      <tr class="info">
        <td class="text-left"><%=Dados("ZC9_VKGNET")%>&nbsp;</td>
        <td class="text-left"><%=Dados("ZC9_DIFDIA")%>&nbsp;</td>
        <td class="text-left"><%=Dados("ZC9_ORIGEM")%>&nbsp;</td>
        <td class="text-left"><%=Dados("ZC9_IDTAB")%>&nbsp;</td>
        <td class="text-left"><%=Dados("ZC9_PROC")%>&nbsp;</td>
        <td class="text-left"><%=Dados("ZC9_OBS")%>&nbsp;</td>
        <td class="text-left"><%=Dados("ZC9_ARQORI")%>&nbsp;</td>
      </tr>
    </table>
  </td>
</tr>
<tr><td><p></p></td></tr>

<%
                QtdeTitulos =  QtdeTitulos + 1
                TotalIPI = TotalIPI +  Dados("ZC9_VIPI")
                TotalICMS = TotalICMS +  Dados("ZC9_VICMS")
                TotalCOFINS = TotalCOFINS +  Dados("ZC9_VCOFIN")
                TotalPIS = TotalPIS +  Dados("ZC9_VPIS")
                TotBasLib = TotBasLib + Dados("ZC9_VLBASE")
                TotLib    = TotLib    + Dados("ZC9_VLCOM")
                TotalIRRF = TotalIRRF + Dados("ZC9_IRRF")
                NumLib    = NumLib    + 1
              response.flush()
              Dados.MoveNext()
            Wend
            Dados.Close()
%>

<%
      end if
%>
      </table>
      <table class="table table-striped table-bordered">
          <tr>
            <td colspan="4"><b>Total de Vendas:</b></td>
            <td class="text-right"><b><%=FormatNumber(TotBasLib,2)%></b></td>
          </tr>
          <tr>
            <td colspan="4"><b>Total de Comissões:</b></td>
            <td class="text-right"><b><%=FormatNumber(TotLib,2)%></b></td>
          </tr>
          <tr>
            <td colspan="4"><b>Total de Titulos:</b></td>
            <td class="text-right"><b><%=FormatNumber(QtdeTitulos,0)%></b></td>
          </tr>
          <tr>
            <td colspan="4"><b>Total IPI:</b></td>
            <td class="text-right"><b><%=FormatNumber(TotalIPI,2)%></b></td>
          </tr>
          <tr>
            <td colspan="4"><b>Total ICMS:</b></td>
            <td class="text-right"><b><%=FormatNumber(TotalICMS,2)%></b></td>
          </tr>
          <tr>
            <td colspan="4"><b>Total COFINS:</b></td>
            <td class="text-right"><b><%=FormatNumber(TotalCOFINS,2)%></b></td>
          </tr>  
          <tr>
            <td colspan="4"><b>Total PIS:</b></td>
            <td class="text-right"><b><%=FormatNumber(TotalPIS,2)%></b></td>
          </tr>  
          <tr>
            <td colspan="4"><b>Total IRRF:</b></td>
            <td class="text-right"><b><%=FormatNumber(TotalIRRF,2)%></b></td>
          </tr>  
      </table>
    </div>

</body>
</html>