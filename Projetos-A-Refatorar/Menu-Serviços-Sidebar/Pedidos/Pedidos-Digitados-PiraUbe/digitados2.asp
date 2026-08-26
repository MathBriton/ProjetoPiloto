
<!--#include file="../../../../includes/variaveis.asp" -->
<!--#include file="../../../../cnn/cnn.asp" --> 
<!--#include file="../../../../includes/atribui_nivel.inc" -->
<html>
  <head>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>SWnet - Consulta / Exclusão a Pedidos Digitados</title>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="PRAGMA" content="no-cache" />
    <meta http-equiv="cache-control" content="no-cache" />

    <link rel="stylesheet" href="../../../../bootstrap/5.2.0/css/bootstrap.min.css" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />

    <script src="../../../../jQuery/3.2.1/jquery-3.2.1.slim.min.js" ></script>
    <script src="../../../../bootstrap/5.2.0/js/bootstrap.min.js" ></script>

    <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>

    <style type="text/css">

      body {
        /*font-size:75%*/
      }

      .CabecalhoDestaque {
        background-color:#555555;
        Color:white;
      }

      #loader {
        position: absolute;
        left: 50%;
        top: 50%;
        z-index: 1;
        width: 120px;
        height: 120px;
        margin: -76px 0 0 -76px;
        border: 16px solid #f3f3f3;
        border-radius: 50%;
        border-top: 16px solid #3498db;
        -webkit-animation: spin 2s linear infinite;
        animation: spin 2s linear infinite;
      }
      
      @-webkit-keyframes spin {
        0% { -webkit-transform: rotate(0deg); }
        100% { -webkit-transform: rotate(360deg); }
      }

      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }

      /* Add animation to "page content" */
      .animate-bottom {
        position: relative;
        -webkit-animation-name: animatebottom;
        -webkit-animation-duration: 1s;
        animation-name: animatebottom;
        animation-duration: 1s
      }

      @-webkit-keyframes animatebottom {
        from { bottom:-100px; opacity:0 } 
        to { bottom:0px; opacity:1 }
      }

      @keyframes animatebottom { 
        from{ bottom:-100px; opacity:0 } 
        to{ bottom:0; opacity:1 }
      }  

      @media print{

          @page{
            size: landscape;
          }

          #image-example{
              position: absolute;
              top: 0;
              bottom: 0;
              left: 0;
              right: 0;
              margin: auto;
          }

      }

      .tableFixHead          { overflow: auto; height: 75vh; }
      .tableFixHead thead th { position: sticky; top: 0; z-index: 1; }

      /* Just common table stuff. Really. */
      table  { border-collapse: collapse; width: 100%; }
      th, td { padding: 8px 16px; }
      th     { background-color: #555555;}      
    </style>
  </head>

  <!-- inicio script link nova pagina --> 
  <script>
    function link(end) {
      window.open(end,"nova","height=250,width=417,left=370,top=0,scrollbars=yes");
    }
  </script>

  <body onload="MostrarPagina();" onafterprint="MostrarPagina();">

    <div class="text-center">
      <h3 style="color: gray">Confirmação de Exclusão de Pedidos Digitados</h3>
      <h6>Pedido(s) Digitado(s) por: <%response.write usr%></h6>
    </div>

<%
              checado=request.form("checados")

              session("checado")=checado

              if checado="" then

                  response.write "" &_
                    "</br>" &_
                    "<table width=450 border=0 cellpadding=0 cellspacing=0 bgcolor=#FF0000 align=center>" &_
                      "<tr>" &_
                        "<td align=center valign=top class=texto_atencao style='padding=10px'>ATEN&Ccedil;&Atilde;O<br>" &_
                          "<table width=100% border=0 cellspacing=0 cellpadding=0>" &_
                            "<tr>" &_
                              "<td align=center bgcolor=#FFFFFF style='padding=10px'>" &_
                                "<table width=100% border=0 cellspacing=0 cellpadding=0>" &_
                                  "<tr>" &_
                                    "<td style=padding=2px><img src=../../../images/alerta_val_user.gif width=31 height=31 align=middle></td>" &_
                                    "<td class=texto_simples2>Nenhum pedido foi selecionado.</b>!<br>Clique <a href=# onMouseDown='history.go(-1);'>aqui</a> para voltar e selecionar.</td>" &_
                                  "</tr>" &_
                                "</table>" &_                 
                              "</td>" &_
                            "</tr>" &_
                          "</table>" &_
                        "</td>" &_
                      "</tr>" &_
                    "</table>" &_
                  "</br>"

                  response.end

              end if

              SQL = "" &_
                "SELECT " &_
                  "SZF010.*, " &_
                  "SB1010.B1_STATUS AS ESTATUS, " &_
                  "SB1010.B1_DESC AS REF," &_
                  "SA1010.A1_NREDUZ AS CLIENTE " &_
                "FROM " &_
                  "SZF010 " &_
                "INNER JOIN SB1010 ON SZF010.ZF_PRODUTO = SB1010.B1_COD " &_
                "INNER JOIN SA1010 ON SZF010.ZF_CLIENTE = SA1010.A1_COD AND SZF010.ZF_LOJACLI = SA1010.A1_LOJA " &_
                "WHERE " &_
                  "(LTRIM(RTRIM(SZF010.ZF_ESTATUS)) NOT IN ('LI','OK')) AND " &_
                  "(SZF010.D_E_L_E_T_ <> '*') AND " &_
                  "(SA1010.D_E_L_E_T_ <> '*') AND " &_
                  "(SB1010.D_E_L_E_T_ <> '*') AND " &_
                  "(SZF010.ZF_NUMERO  IN(" & checado & "))"

              Dados.Open(SQL)

              if Dados.EOF then

                response.write "" &_
                  "</br>" &_
                  "<table width=450 border=0 cellpadding=0 cellspacing=0 bgcolor=#FF0000 align=center>" &_
                    "<tr>" &_
                      "<td align=center valign=top class=texto_atencao style='padding=10px'>ATEN&Ccedil;&Atilde;O<br>" &_
                        "<table width=100% border=0 cellspacing=0 cellpadding=0>" &_
                          "<tr>" &_
                            "<td align=center bgcolor=#FFFFFF style='padding=10px'>" &_
                              "<table width=100% border=0 cellspacing=0 cellpadding=0>" &_
                                "<tr>" &_
                                  "<td style=padding=2px><img src=../../../images/alerta_val_user.gif width=31 height=31 align=middle></td>" &_
                                  "<td class=texto_simples2>" &_
                                    "Nenhum pedido foi encontrado. Verifique na consulta Pedidos Validados</b>!" &_
                                    "<br>" &_
                                    "Clique <a href=# onMouseDown='history.go(-1);'>aqui</a> para voltar." &_
                                  "</td>" &_
                                "</tr>" &_
                              "</table>" &_                 
                            "</td>" &_
                          "</tr>" &_
                        "</table>" &_
                      "</td>" &_
                    "</tr>" &_
                  "</table>" &_
                "</br>"              
                response.end

              end if

              Dados.MoveFirst()
%>
    <div class="container-fluid">
      <div style="box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.1), 0 6px 10px 0 rgba(0, 0, 0, 0.19);">
        <div class="container-fluid" style="background-color: rgb(242, 242, 242);">

          <form method="post" action="digitados3.asp" name="form">

            <div class="table-responsive tableFixHead">

              <br>

              <table class="table table-bordered table-sm table-responsive-sm table-hover" style="font-size:12px;background-color:white">

                <thead bgcolor="#E0E0E0">
                  <td align=center>
                    <i class="fa fa-trash" aria-hidden="true"></i>
                  </td>
                  <td><div align=center>PEDIDO</div></td>
                  <td><div align=center>CLIENTE</div></td>
                  <td><div align=center>NOME</div></td>
                  <td><div align=center>FT</div></td>
                  <td><div align=center>PRODUTO</div></td>
                  <td><div align=center>QTDE.</div></td>
                  <td><div align=center>PREÇO</div></td>
                  <td><div align=center>ENTREGA</div></td>
                  <td><div align=center>MEDIDA</div></td>
                  <td><div align=center>VINCOS</div></td>
                  <td><div align=center>EMISSÃO</div></td>
                </thead>
<%
              Do Until (Dados.EOF)

                vData1 = mid(Dados("ZF_ENT1"),7,2)+"/"+mid(Dados("ZF_ENT1"),5,2)+"/"+ mid(Dados("ZF_ENT1"),1,4) 
                vData2 = mid(Dados("ZF_ENT2"),7,2)+"/"+mid(Dados("ZF_ENT2"),5,2)+"/"+ mid(Dados("ZF_ENT2"),1,4)
                vEmissao = mid(Dados("ZF_EMISSAO"),7,2)+"/"+mid(Dados("ZF_EMISSAO"),5,2)+"/"+ mid(Dados("ZF_EMISSAO"),1,4)
%>
                  <tbody>
                    <td align=center>
                      <i style="color:red" class="fa fa-trash" aria-hidden="true"></i>
                    </td>
                    <td>
                      <div align=center><%=Dados ("ZF_NUMERO")%></div>
                    </td>
                    <td>
                      <div align=center><%=Dados ("ZF_CLIENTE") & " - " & Dados ("ZF_LOJACLI")%></div>
                    </td>
                    <td>
                      <div align=center><%=Dados ("CLIENTE")%></div>
                    </td>
                    <td>
                      <div align=center><%=Dados ("ZF_PRODUTO")%></div>
                    </td>
                    <td>
                      <div align=center><%=Dados ("REF")%></div>
                    </td>
                    <td>
                      <div align=center><%=FormatNumber(Dados ("ZF_QUANT"),0)%></div>
                    </td>
                    <td>
                      <div align=center><%=FormatNumber(Dados ("ZF_PRECO"),5)%></div>
                    </td>
                    <td>
                      <div align=center><%=vData1 & " - " & vData2%></div>
                    </td>
                    <td>
                      <div align=center><%=FormatNumber(Dados ("ZF_LARG"),0) & " x " & FormatNumber(Dados ("ZF_COMP"),0)%></div>
                    </td>
                    <td>
                      <div align=center><%=FormatNumber(Dados ("ZF_VINCO1"),0) & " - " & FormatNumber(Dados ("ZF_VINCO2"),0) & " - " & FormatNumber(Dados ("ZF_VINCO3"),0)%></div>
                    </td>
                    <td>
                      <div align=center><%=vEmissao & " as " &  Dados ("ZF_HORA")%></div>
                    </td>
                  </tbody>
<%              
                Dados.MoveNext()

              Loop

              Dados.Close()
%>
              </table>
            </div>

            <div class="row">
              <div class="col" align="center">
                <input name="Submit2" type="submit" class="btn btn-danger" value="Confirmar Exclusão">
              </div>
            </div>

            <br>

          </form>
        </div>
      </div>
    </div>
  </body>
</html>
