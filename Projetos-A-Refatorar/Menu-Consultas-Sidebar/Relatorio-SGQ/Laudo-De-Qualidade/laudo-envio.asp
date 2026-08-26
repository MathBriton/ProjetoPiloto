<!DOCTYPE html>
<%
session.lcid = 1046

On Error Resume Next
%>

<!--#include file ="../cnn2/cnn2.asp"-->
<!--#include file = "../funcoes/funcoes.asp"-->

<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta charset="utf-8" />
    <title>SWnet - Certificado de Qualidade</title>   
<%
  Dim vNota, vSerie
  vNota  = trim(request("id1"))
  vSerie = trim(request("id2"))
  vOP    = trim(request("id3"))

%>

    <link href="css/css.css" rel="stylesheet" type="text/css" />

    <style type="text/css">
      .auto-style2 {
        width: 421px;
      }
      .auto-style3 {
        width: 150px;
      }
      .Margem{
        margin-top:50px;
      }
      .text{
        font-family:Verdana;
        font-weight:bold;
        text-align:center
      }
      .table-notas {
        padding: 4px;
        border-spacing: 0px;
      }
      .nota1 {
        text-align: justify;
        text-justify: inter-word;
        font-size: 11px;
        vertical-align: top;
      }

      .container {
          display: grid; /* Ativa o Grid */
          place-items: center; /* Centraliza horizontal e verticalmente */
          /*border: 2px solid #ccc; /* Apenas para visualização */
          height: 200px; /* Altura da div */
      }
      img {
          max-width: 100%; /* Garante que a imagem não exceda a largura da div */
          height: auto; /* Mantém a proporção da imagem */
      }

      .filial-unidade {
        text-align: right;
      }
      .filial-endereco {
        text-align: left;
      }
            
    </style>
  </head>

<%
SQL = "" &_
  "SELECT DISTINCT " &_
    "R.OP, " &_
    "A1_GRPVEN GrupoCliente, " &_
    "F2_DOC, " &_
    "D2_COD, " &_
    "F2_CLIENTE, " &_
    "F2_LOJA, " &_
    "F2_SERIE, " &_
    "CASE WHEN F2_SERIE IN ('001','002') THEN " &_
      "'Unidade Fabril - Pirapetinga - MG - Rua Inpa, 186 - B. Centro CEP 36730-000' " &_
    "ELSE " &_
      "'Unidade Fabril - Uberaba - MG - Rodovia BR 50, km 168 - B. Distrito Industrial 2 - CEP 38064-750' " &_
    "END AS END_FABRICACAO, " &_
    "A1_PESSOA, " &_
    "A1_CGC, " &_
    "A1_EMAIL, " &_
    "A1_EMAILCO, " &_
    "8 FlagGap, " &_
    "r.Pedido,  " &_
    "ISNULL(r.Referencia,'') Referencia, " &_
    "ISNULL(r.Composicao,'') Composicao, " &_
    "ISNULL(r.AlturaInterna,'') AlturaInterna, " &_
    "ISNULL(r.ComprimentoInterno,'') ComprimentoInterno, " &_
    "ISNULL(r.LarguraInterna,'') LarguraInterna, " &_
    "ISNULL(r.Gramatura,'') Gramatura, " &_
    "ISNULL(r.ColunaMinimo,'') ColunaMinimo, " &_
    "ISNULL(r.Gramatura,'') Gramatura1, " &_
    "ISNULL(r.Espessura,'') Espessura1, " &_
    "ISNULL(r.CobbIntMaximo,'') CobbIntMaximo, " &_
    "ISNULL(r.CobbExtMaximo,'') CobbExtMaximo, " &_
    "ISNULL(r.Espessura,'') Espessura, " &_
    "ISNULL(r.MullenMinimo,'') MullenMinimo, " &_
    "ISNULL(r.CrushMinimo,'') CrushMinimo, " &_
    "ISNULL(r.TCrush,'') TCrush, " &_
    "ISNULL(r.TColuna,'') TColuna, " &_
    "ISNULL(r.TCobbInt,'') TCobbInt, " &_
    "ISNULL(r.TCobbExt,'') TCobbExt, " &_
    "ISNULL(r.TEspessura,'') TEspessura, " &_
    "ISNULL(r.TMullen,'') TMullen, " &_
    "ISNULL(r.TGramatura,'') TGramatura, " &_
    "ISNULL(r.TComprimentoInterno,'') TComprimentoInterno, " &_
    "ISNULL(r.TLarguraInterna,'') TLarguraInterna, " &_
    "ISNULL(r.TAlturaInterna,'') TAlturaInterna, " &_
    "ISNULL(r.DPGramatura,'') DPGramatura, " &_
    "ISNULL(r.DPColuna,'') DPColuna, " &_
    "ISNULL(r.Responsavel1,'') Responsavel1, " &_
    "ISNULL(F2_DOC,'') NotaFiscal, " &_
    "ISNULL(r.QtdNotaFiscal,'') QtdNotaFiscal, " &_
    "ISNULL(r.Responsavel2,'') Responsavel2, " &_
    "ISNULL(c.Mullen,'') Mullen, " &_
    "ISNULL(c.CobbExt1,'') CobbExt1, " &_
    "ISNULL(c.CobbInt1,'') CobbInt1, " &_
    "ISNULL(c.Coluna,'') Coluna, " &_
    "ISNULL(c.Crush,'') Crush, " &_
    "ISNULL(c.Gramatura,'') Gramatura_Composicao, " &_
    "ISNULL(c.Espessura,'') Espessura_Composicao, " &_
    "ISNULL(m.Municipio,'') Municipio, " &_
    "ISNULL(r.TCompressao,'') TCompressao, " &_
    "ISNULL(r.Compressao,'') Compressao, " &_
    "ISNULL(p.Lap,'') Lap, " &_
    "ISNULL(p.Item,'') Item, " &_
    "ISNULL(r.CodigoBarras,'') CodigoBarras, " &_
    "ISNULL(r.TUmidade,'') TUmidade, " &_
    "ISNULL(r.DataFabricacao,'') DataFabricacao, " &_
    "ISNULL(r.DataCriacao,'') DataCriacao, " &_
    "ISNULL(p.Obs7,'') Obs7, " &_
    "ISNULL(p.PedCliente,'') PedCliente, " &_
    "ISNULL(p.ItemPedido,0) ItemPedido, " &_
    "ISNULL(r.ObsOF,'') ObsOF, " &_
    "ISNULL(p.Chapa,'') Chapa, " &_
    "ISNULL(c.CobbInt0,'') CobbInt0, " &_
    "ISNULL(p.ResinaInterna,'') ResinaInterna, " &_
    "ISNULL(c.CobbExt0,'') CobbExt0, " &_
    "ISNULL(p.ResinaExterna,'') ResinaExterna, " &_
    "ISNULL(cl.CodCliente,'') CodCliente, " &_
    "ISNULL(l.RazaoSocial,'') RazaoSocial, " &_
    "ISNULL(c.Onda,'') Onda " &_
  "FROM BRPIR1M001.DADOSADV.dbo.SF2010 SF2 with(nolock) " &_
    "INNER JOIN BRPIR1M001.DADOSADV.dbo.SD2010 SD2 with(nolock) ON " &_
          "SF2.F2_FILIAL = SD2.D2_FILIAL " &_
      "AND SF2.F2_DOC = SD2.D2_DOC " &_
      "AND SF2.F2_SERIE = SD2.D2_SERIE " &_
      "AND SF2.F2_CLIENTE = SD2.D2_CLIENTE " &_
      "AND ISNULL(SD2.D_E_L_E_T_,'') = '' " &_
    "OUTER APPLY(SELECT G1_COMP FROM BRPIR1M001.DADOSADV.dbo.SG1010 WHERE G1_COD = D2_COD AND SUBSTRING(D2_COD,1,3) IN ('750','760') AND D_E_L_E_T_ = '') SG1" &_
    "INNER JOIN BRPIR1M001.DADOSADV.dbo.SA1010 SA1 with(nolock) ON " &_
          "SF2.F2_CLIENTE = SA1.A1_COD " &_
      "AND SF2.F2_LOJA = SA1.A1_LOJA " &_
      "AND ISNULL(SA1.D_E_L_E_T_,'') = '' " &_
    "INNER JOIN rellaudo r with(nolock) on LTRIM(RTRIM(CONCAT(r.Pedido,'/',r.Item))) COLLATE Latin1_General_CI_AS = LTRIM(RTRIM(CONCAT(D2_PEDIDO,'/',CASE WHEN SUBSTRING(D2_COD,1,3) IN ('750','760') THEN G1_COMP ELSE D2_COD END)))" &_
    "INNER JOIN Pedidos p with(nolock) ON " &_
          "r.Pedido = p.Pedido " &_
      "AND r.Item = p.Item " &_
      "AND SD2.D2_PEDIDO COLLATE Latin1_General_BIN = p.Pedido " &_
    "INNER JOIN Clientes cl with(nolock) ON p.IDCliente=cl.IDCliente " &_
    "INNER JOIN Lojas l with(nolock) ON p.IDCliente=l.IDCliente AND p.IDLoja=l.IDLoja " &_
    "INNER JOIN LocaisEntrega le with(nolock) ON " &_
          "p.IDCliente=le.IDCliente " &_
      "AND p.IDLoja=le.IDLoja " &_
      "AND p.IDLocalEntrega=le.IDLocalEntrega " &_
    "LEFT OUTER JOIN composic c with(nolock) ON " &_
          "CASE WHEN LEFT(r.Composicao, 1) = '*' THEN " &_
            "SUBSTRING(r.Composicao, 2, LEN(r.Composicao)) " &_
          "ELSE " &_
            "r.Composicao " &_
          "END = c.Composicao " &_
    "LEFT OUTER JOIN Municipios m with(nolock) ON m.IDMunicipio = le.IDMunicipio " &_
  "WHERE SF2.D_E_L_E_T_ = '' " &_
    "AND F2_DOC = '" & vNota & "' " &_
    "AND F2_SERIE = '" & vSerie & "' " &_
    "AND OP = '" & vOP & "' "

if (vSerie="001") or (vSerie="002") then 

  filialDescricao = "Pirapetinga"
  filial = iif( (vSerie="001"), "00", "05" )

  SET TRIMBOX = Conecta( "PIR-TRIMBOX", SQL )

end if

if (vSerie="003") or (vSerie="004") then

  filialDescricao = "Uberaba"
  filial = iif( (vSerie="003"), "04", "06" )

  SET TRIMBOX = Conecta( "UBE-TRIMBOX", SQL )

end if

if NOT isObject(TRIMBOX) or (TRIMBOX IS NOTHING) then

  response.end

end if

if not TRIMBOX.EOF then 

  if TRIMBOX("ColunaMinimo") > TRIMBOX("TColuna") then

    TesteColuna = TRIMBOX("ColunaMinimo") + 0.1

  else

    TesteColuna = TRIMBOX("TColuna")
    
  end if
%>
  <body>
    <div id="margem">
      <div id="laudo">
        <table>
          <thead>
            <tr>
              <th colspan="5"><h1><em>Certificado da Qualidade</em></h1></th>
            </tr>
            <tr>
              <td colspan="5" class="input_text_align_right padding_botton">Emissão: <%=Date()%></td>
            </tr>
          </thead>
          <tbody>
            <tr>
              <th>Lote Nr:</th>
              <td colspan="2" class="fonte_13_bold"><%=TRIMBOX("Pedido") %></td>
              <th>OF: <%=Right(Trim(TRIMBOX("OP")),1) %></th>
              <td rowspan="3" >
                <p>
                  <span class="pisicao input_text_align_center">
                    Este pedido foi produzido em<br />
                    conformidade com os requisitos de projeto<br />
                    e em atendimento às seguintes<br />
                    especificações técnicas:
                  </span>
                </p>
              </td>
            </tr>

            <tr>
              <th>Cliente:</th>
              <td colspan="3" class="fonte_13_bold"><%=TRIMBOX("CodCliente")&" - "&TRIMBOX("RazaoSocial") %></td>
            </tr>

            <tr>
              <th>Referência:</th>
              <td colspan="3" class="fonte_13_bold"><%=TRIMBOX("Referencia") %></td>
            </tr>

            <tr>
              <th>Onda:</th>
              <td class="fonte_13_bold"><%=TRIMBOX("Onda") %></td>
              <th>Estrutura:</th>
              <td colspan="2" class="fonte_13_bold"><%=TRIMBOX("Composicao") %></td>
            </tr>

            <tr>
              <th>Nota Fiscal Nr:</th>
              <td class="fonte_13_bold"><%if TRIMBOX("NotaFiscal") <> 0 then %><%=FormatNumber(TRIMBOX("NotaFiscal"),0) %><%else %>&nbsp;<%end if %></td>
              <th>Qtd da Nota Fiscal:</th>
              <td colspan="2" class="fonte_13_bold"><% if TRIMBOX("QtdNotaFiscal") <> 0 then %><%=FormatNumber(TRIMBOX("QtdNotaFiscal"),0) %><% else %>&nbsp;<% end if %></td>
            </tr>

            <tr>
              <th>Ordem de Compra Nr: </th>
              <td colspan="5"><%=TRIMBOX("OP") %></td>
            </tr>

            <tr>
              <br>
              <td colspan="5"><%="OF " & TRIMBOX("OP") %></td>
            </tr>

            <tr>
              <td colspan="2">&nbsp;</td>
              <td class="input_text_align_center" colspan="2"><u>ESPECIFICADO</u></td>
              <td class="input_text_align_center" colspan="2"><u>ENCONTRADO</u></td>
              <td >&nbsp;</td>
            </tr>

            <tr>
              <td >&nbsp;</td>
              <td >Medidas Internas C x L x A (+/- 2mm):</td>
              <td class="fonte_13_bold" colspan="2"><center><%=TRIMBOX("ComprimentoInterno")&"  x  "&TRIMBOX("LarguraInterna")&"  x  "&TRIMBOX("AlturaInterna") %></center></td>
              <td class="fonte_13_bold" colspan="2"><center><%=TRIMBOX("ComprimentoInterno")&"  x  "&TRIMBOX("LarguraInterna")&"  x  "&TRIMBOX("AlturaInterna") %></center></td>
              <!-- <td class="fonte_13_bold" colspan="2"><center><%=TRIMBOX("TComprimentoInterno")&"  x  "&TRIMBOX("TLarguraInterna")&"  x  "&TRIMBOX("TAlturaInterna") %></center></td> -->
              <td >&nbsp;</td>
            </tr>
            <tr>
              <td><b>Obs:</b>&nbsp;</td>
              <td align="center"><u>CARACTERISTICA</u></td>
              <td class="input_text_align_center" colspan="2"><u>ESPECIFICADO</u></td>
              <td class="input_text_align_center" colspan="2"><u>ENCONTRADO</u></td>
              <!--
              <td class="input_text_align_center"><u>ENC.</u></td>
              -->
            </tr>
          </tbody>
        </table>
        <table>         
          <tbody>
            <tr>
              <td>
<%
                response.write iif( trim(TRIMBOX("Obs7")) <> "", trim(TRIMBOX("ObsOF")) & " / " & trim(TRIMBOX("Obs7")), trim(TRIMBOX("ObsOF")) )
%>
              </td>
            </tr>
            <tr>
<%
              if TRIMBOX("Gramatura1") <> "0" then 
%>
              <td class="input_text_align_right"><b>Gramatura (+/- 10%)</b></td>
              <td class="input_text_align_right fonte_13_bold"><%=TRIMBOX("Gramatura1") %></td>
              <td class="input_text_align_right">g/m<sup>2</sup></td>
              <td class="input_text_align_right fonte_13_bold"><%if trim(TRIMBOX("TGramatura")) <> "0" then %><%=TRIMBOX("TGramatura") %><%else %>&nbsp;<%end if %></td>
              <td class="input_text_align_right">g/m<sup>2</sup></td>
              <td class="input_text_align_right fonte_13_bold"><%if formatnumber(TRIMBOX("DPGramatura"),2) <> 0 then %><%=formatnumber(TRIMBOX("DPGramatura"),2) %><%else %>&nbsp;<%end if %></td>
<% 
              else 
%>
              <td>&nbsp;</td>
<% 
              end if 
%>
            </tr>
<%
            if formatnumber(TRIMBOX("ColunaMinimo"),1) <> "0" then 
%>
            <tr>
              <td class="input_text_align_right"><b>Coluna (min)</b></td>
              <td class="input_text_align_right fonte_13_bold"><%=formatnumber(TRIMBOX("ColunaMinimo"),1) %></td>
              <td class="input_text_align_right">kgf/cm</td>
              <td class="input_text_align_right fonte_13_bold"><%if formatnumber(TesteColuna,1) <> 0 then %><%=formatnumber(TesteColuna,1) %><% else %>&nbsp;<% end if %></td>
              <td class="input_text_align_right">kgf/cm</td>
              <!--
              <td class="input_text_align_right fonte_13_bold"><%if formatnumber(TRIMBOX("DPColuna"),2) <> 0 then %><%=formatnumber(TRIMBOX("DPColuna"),2) %><% else %>&nbsp;<% end if %></td>
              -->
            </tr>
<%
            end if 
            if formatnumber(TRIMBOX("MullenMinimo"),2) <> 0 then  
%>
            <tr>
              <td class="input_text_align_right"><b>Arrebentamento (ref)</b></td>
              <td class="input_text_align_right fonte_13_bold"><%=formatnumber(TRIMBOX("MullenMinimo")*6894.744825/1000,0) %></td>
              <td class="input_text_align_right">kPa</td>
              <td class="input_text_align_right fonte_13_bold"><%if formatnumber(TRIMBOX("TMullen"),2) <> 0 then %><%=formatnumber(TRIMBOX("TMullen")*6894.744825/1000,1) %><%else %>&nbsp;<%end if %></td>
              <td class="input_text_align_right">kPa</td>
              <td class="input_text_align_right">&nbsp;</td>
            </tr>
<%
            end if 
            if formatnumber(TRIMBOX("MullenMinimo"),2) <> 0 then  
%>
            <tr>
              <td class="input_text_align_right">&nbsp;</td>
              <td class="input_text_align_right"><%=formatnumber(TRIMBOX("MullenMinimo"),1) %></td>
              <td class="input_text_align_right">lbf/pol<sup>2</sup></td>
              <td class="input_text_align_right"><%if formatnumber(TRIMBOX("TMullen"),1) <> 0 then %><%=formatnumber(TRIMBOX("TMullen"),1) %><%else %>&nbsp;<%end if %></td>
              <td class="input_text_align_right">lbf/pol<sup>2</sup></td>
              <td class="input_text_align_right">&nbsp;</td>
            </tr>
<%
            end if 
            if formatnumber(TRIMBOX("MullenMinimo"),2) <> 0 then  
%>
            <tr>
              <td>&nbsp;</td>
              <td class="input_text_align_right"><%=formatnumber(TRIMBOX("MullenMinimo")*0.070306,1) %></td>
              <td class="input_text_align_right">kgf/cm<sup>2</sup></td>
              <td class="input_text_align_right"><%if formatnumber(TRIMBOX("TMullen"),1) <> 0 then %><%=formatnumber(TRIMBOX("TMullen")*0.070306,1) %><%else %>&nbsp;<%end if %></td>
              <td class="input_text_align_right">kgf/cm<sup>2</sup></td>
              <td class="input_text_align_right">&nbsp;</td>
            </tr>
<%
            end if 
            if formatnumber(TRIMBOX("CrushMinimo"),2)  <> 0 then 
%>
            <tr>
              <td class="input_text_align_right"><b>Esmagamento (mín)</b></td>
              <td class="input_text_align_right fonte_13_bold"><%=formatnumber(TRIMBOX("CrushMinimo"),2) %></td>
              <td class="input_text_align_right">kgf/cm<sup>2</sup></td>
              <td class="input_text_align_right fonte_13_bold"><%if formatnumber(TRIMBOX("TCrush"),2) <> 0 then %><%=formatnumber(TRIMBOX("TCrush"),2) %><%else %>&nbsp;<%end if %></td>
              <td class="input_text_align_right">kgf/cm<sup>2</sup></td>
              <td class="input_text_align_right">&nbsp;</td>
            </tr>
<%
            end if 
            if formatnumber(TRIMBOX("CobbIntMaximo"),0) <> 0 then
%>
            <tr>
              <td class="input_text_align_right"><b>COBB - Absorção de Água Int. (min-máx.)</b></td>
              <td class="input_text_align_right fonte_13_bold">0-<%=formatnumber(TRIMBOX("CobbIntMaximo"),0) %></td>
              <td class="input_text_align_right">g[H<sub>2</sub>O]/m<sup>2</sup></td>
              <td class="input_text_align_right fonte_13_bold"><%if formatnumber(TRIMBOX("TCobbInt"),0) <> 0 then %><%=formatnumber(TRIMBOX("TCobbInt"),0) %><%else %>&nbsp;<%end if %></td>
              <td class="input_text_align_right">g[H<sub>2</sub>O]/m<sup>2</sup></td>
              <td class="input_text_align_right">&nbsp;</td>
            </tr>
<%
            end if 
            if formatnumber(TRIMBOX("CobbExtMaximo"),0) <> 0 then 
%>
            <tr>
              <td class="input_text_align_right"><b>COBB - Absorção de Água Ext. (min-máx.)</b></td>
              <td class="input_text_align_right fonte_13_bold">0-<%=formatnumber(TRIMBOX("CobbExtMaximo"),0) %></td>
              <td class="input_text_align_right">g[H<sub>2</sub>O]/m<sup>2</sup></td>
              <td class="input_text_align_right fonte_13_bold"><%if formatnumber(TRIMBOX("TCobbExt"),0) <> 0 then %><%=formatnumber(TRIMBOX("TCobbExt"),0) %><%else %>&nbsp;<%end if %></td>
              <td class="input_text_align_right">g[H<sub>2</sub>O]/m<sup>2</sup></td>
              <td class="input_text_align_right">&nbsp;</td>
            </tr>
<%
            end if 
            if formatnumber(TRIMBOX("Espessura1"),2) <> 0 then 
%>
            <tr>
              <td class="input_text_align_right"><b>Espessura (mín)</b></td>
              <td class="input_text_align_right fonte_13_bold"><%=formatnumber(TRIMBOX("Espessura1"),2) %> - <%=formatnumber(TRIMBOX("Espessura"),2) %></td>
              <td class="input_text_align_right">mm</td>
              <td class="input_text_align_right fonte_13_bold"><%if formatnumber(TRIMBOX("TEspessura"),2) <> 0 then %><%=formatnumber(TRIMBOX("TEspessura"),2) %><% else %>&nbsp;<% end if %></td>
              <td class="input_text_align_right">mm</td>
              <td class="input_text_align_right">&nbsp;</td>
            </tr>
<%
            end if 
            if formatnumber(TRIMBOX("Compressao"),2) <> 0 then  
%>
            <tr>
              <td class="input_text_align_right"><b>Compressão (mín.)</b></td>
              <td class="input_text_align_right fonte_13_bold"><%=formatnumber(TRIMBOX("Compressao"),2) %></td>
              <td class="input_text_align_right">N</td>
              <td class="input_text_align_right fonte_13_bold"><%if formatnumber(TRIMBOX("TCompressao"),2) <> 0 then %><%=formatnumber(TRIMBOX("TCompressao"),2) %><%else %>&nbsp;<%end if %></td>
              <td class="input_text_align_right">N</td>
              <td class="input_text_align_right">&nbsp;</td>
            </tr>
<%
            end if 
            if formatnumber(TRIMBOX("Compressao"),2) <> 0 then  
%>
            <tr>
              <td>&nbsp;</td>
              <td class="input_text_align_right"><%=formatnumber((TRIMBOX("Compressao")/9.8),2) %></td>
              <td class="input_text_align_right">kgf</td>
              <td class="input_text_align_right"><%if formatnumber(TRIMBOX("TCompressao"),2) <> 0 then %><%=formatnumber((TRIMBOX("TCompressao")/9.8),2) %><%else %>&nbsp;<%end if %></td>
              <td class="input_text_align_right">kgf</td>
              <td class="input_text_align_right">&nbsp;</td>
            </tr>
<% 
            end if 

            'if TRIMBOX("Lap") <> 0 then 
%>
            <tr>
              <td class="input_text_align_right"><b>Tamanho do lap</b></td>
<% 
              if InStr(1,TRIMBOX("Composicao"),"0B") or InStr(1,TRIMBOX("Composicao"),"0C") then 
                tamLap = 35
              elseif InStr(1,TRIMBOX("Composicao"),"BB") then 
                tamLap = 40
              else
                tamLap = TRIMBOX("Lap")
              end if                
%>
              <td class="input_text_align_right  fonte_13_bold"><%=tamLap%></td>
              <td class="input_text_align_right">mm</td>
              <td class="input_text_align_right  fonte_13_bold"><%=tamLap%></td>
              <td class="input_text_align_right">mm</td>
              <td class="input_text_align_right">&nbsp;</td>

            </tr>
<%
'            end if 
%>

            <%'if formatnumber(TRIMBOX("TUmidade"),2) <> 0 and 1<>1 then %>
            <!--
            <tr>
              <td class="input_text_align_right"><b>Umidade</b></td>
              <td class="input_text_align_right  fonte_13_bold">5 - 8</td>
              <td class="input_text_align_right">%</td>
              <td class="input_text_align_right  fonte_13_bold"><%=formatnumber(TRIMBOX("TUmidade"),2)%></td>
              <td class="input_text_align_right">%</td>
              <td class="input_text_align_right">&nbsp;</td>
            </tr>
            -->
            <%'end if %>

            <%'if formatnumber(TRIMBOX("TUmidade"),2) <> 0  then %>
            <!--
            <tr>
              <td class="input_text_align_right"><b>Umidade</b></td>
              <td class="input_text_align_right  fonte_13_bold">5 - 8</td>
              <td class="input_text_align_right">%</td>
              <td class="input_text_align_right  fonte_13_bold"><%=formatnumber(TRIMBOX("TUmidade"),2)%></td>
              <td class="input_text_align_right">%</td>
              <td class="input_text_align_right">&nbsp;</td>
            </tr>
            -->
            <%'end if %>

          </tbody>
        </table>
        <div id="nota">
          <table>
            <tr>
              <td class="nota1">
                Análise Visual / Sensorial: Especificado: Legível / Sem odor estranho Encontrado: OK<br>
                Flap Gap (Abertura entalhe): Especificado (+/- 2mm): 8 (mm) Encontrado: 8 (mm)<br>
                Sentido de onda: Conforme desenho técnico especificado.<br>
                Material 100% reciclável
                <br/><br/>
                NOTAS:
              </td>
              <td class="auto-style3">&nbsp;</td>
              <td><b>Código de Barras:&nbsp;<%=Trim(TRIMBOX("CodigoBarras"))%></b></td>
            </tr>
          </table>

          <table class="table-notas">
            <tr>
              <td class="nota1" width="40%">
                (1) Prazo de Validade = 12 meses da Data de Fabricação mantendo as embalagens em ambiente coberto, seco e arejado;
                <br/><br/>
                (2) Testes Físicos: Devem ser considerados as condições de climatização controladas em temperatura de 25° C ± 2°C e 
                umidade relativa de 57% ± 8% (NBR NM ISO 187). Considerando o clima dos países tropicais;
              <td>
              <td width="1%">
              </td>
              <td class="nota1" width="45%">
                (3) Ensaios realizados de acordo com as Normas Brasileiras para Papelão Ondulado ( ABNT NBR 14979, NBR ISO 3034, NBR 6737, 
                NBR NM-ISO 536, NBR NM-ISO 535 e NBR ISO 3035 ), aplicando-se plano de amostragem por variáveis ( NBR 5429 / ISO 3951 ) e
                amostragem por atributos (NBR 5426 / ISO 2859), considerando NQA de 1,5% para defeitos críticos e 4% para defeitos graves ( NBR NM-ISO 186 );
                <br/><br/>
                (4) Valores encontrados referem-se às médias das medidas dos testes realizados.
              </td>
              <td width="14%">
              </td>
            <tr>
          </table>
        </div>
        <table>
          <tbody>
            <tr>
              <td class="padding_botton20">Responsável pelos testes</td>
              <td class="padding_botton20">&nbsp;</td>
              <td class="padding_botton20">Responsável pela fabricação</td>
            </tr>
            <tr>
              <td class="border_botton assinatura"><%=TRIMBOX("Responsavel1") %></td>
              <td>&nbsp;</td>
              <td class="border_botton assinatura"><%=TRIMBOX("Responsavel2")%></td>
            </tr>
          </tbody>
          <tfoot>
            <tr>
              <td colspan="3" class="padding_top">Data de Fabricação</td>
            </tr>
            <tr>
              <td colspan="3" class="padding_botton130"><b><em><%=TRIMBOX("DataFabricacao") %></em></b></td>
            </tr>
          </tfoot>
        </table>
        <div id="rodape">
          <table>
            <tr>
              <td class="input_text_align_center">Documentos impressos a laser com base em registros e assinaturas eletrônicas</td>
            </tr>
            <tr>
              <td class="revisao">RGQ301rev12<br>Data de Edição: 20/09/2013<br>Data da Revisão: 13/12/2024</td>
            </tr>
            <tr>
              <td class="input_text_align_right"><b>SMURFIT WESTROCK BRASIL INDÚSTRIA DE EMBALAGENS S/A.<br>Unidade de Fabricação&nbsp;<%=filialDescricao %></b></td>
            </tr>
          </table>
          <table>
            <tr>
              <th width="20%">
                <p class="filial-unidade">
                  <b>Matriz</b> - <br>
                  <b>Unidade Fabril</b> - <br>
                  <b>Filial RJ</b> - <br><br>
                  <b>Filial SP</b> - <br>
                  <b>Unidade Fabril</b> - 
                </p>
              </th>
              <th width="80%">
                <p class="filial-endereco">
                  Rua Inpa, 186 - Pirapetinga-MG - CEP 36730 000 - Fone (32)3465-3000 - Fax (32)3465-3001<br>
                  Uberaba-MG - Rodovia BR 50, Km 168 - B. Distrito Indústrial 2 - CEP 38064-750 - Fone (34)2104-3100<br />
                  Av. João Cabri de Nelo Neto, 850 - Bloco 3 - Salas 1306 a 1310 - Cond Ceo Esast Offices, <br>
                  Barra da Tijuca, Rio de Janeiro - RJ, Brasil - CEP 22.775-057 - Fone (21)2136-9000<br />
                  Rua Castilho, 392 - 16º Andar - Brooklin, São Paulo-SP - CEP 04568-010 - Fone (11)2105-9000<br />
                  Saquarema-RJ - R. Durval Souza - Sampaio Correia - CEP 28990-000, Brasil - Fone (21) 98880-6918 
                </p>
              </th>
            </tr>
          </table>
        </div>
      </div>
      <div id="logo" class="container"> 
        <img src="../images/SW_LOGO_BLACK.png" style="opacity: 0.1;" alt="logo" width="950" border="0" style="border:0px"/>
      </div>
    </div>
<% 
else 
%>
  <body class="Margem text">
    <span>Não foi possível encontrar o Laudo informado</span>
<% 
end if 

TRIMBOX.Close()
SET TRIMBOX = NOTHING
%>
  </body>
</html>