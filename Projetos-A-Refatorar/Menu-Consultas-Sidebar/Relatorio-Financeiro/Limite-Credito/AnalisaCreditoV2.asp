<!--#include file="../../cnn2/cnn2.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->
<%
'--------------------------------------------------------------------------------------------------'
function AddZerosLeft(originalString, totalLength)

  Dim currentLength
  currentLength = Len(originalString)
  
  if currentLength < totalLength then
    Dim zerosToAdd
    zerosToAdd = totalLength - currentLength
    originalString = String(zerosToAdd, "0") & originalString
  end if
  
  AddZerosLeft = originalString

end function

Session("MMCliente") = request("codcli")

Dim NomeCliente

Dim SaldoGeral
Dim SaldoFinanceiro
Dim SaldoFinanceiro1
Dim SaldoFinanceiro2
Dim SaldoPedidoLiberado
Dim SaldoPedidoALiberar

NomePlanta = ""

Cliente = request("codcli")
codLoja = request("codLoja")
Filial  = AddZerosLeft( request("Filial"), 2 )


select case Filial
case "00" : NomePlanta = "PIRAPETINGA"                                
case "04" : NomePlanta = "UBERABA"    
case "05" : NomePlanta = "BENTO"      
case "06" : NomePlanta = "MARANGUAPE" 
case "07" : NomePlanta = "SAQUAREMA" 
end select

select case Filial
case "00", "04" : 

  SQL = "EXEC SP_CONSULTA_CREDITO_CLIENTE '"&trim(request("codcli"))&"', '0'"      
  SET SaldoGeral = Conecta( "PIR-DADOSADV", SQL)

  if (SaldoGeral IS NOTHING) then
    response.end
  end if

  DadosPirapetinga()

case "05", "06" 

  SQL = "EXEC [spLimiteCredito] '" & request("codcli") & "'"

  select case Filial
  case "05" : SET SaldoGeral = Conecta( "BEN-DADOSADV", SQL )
  case "06" : SET SaldoGeral = Conecta( "FOR-DADOSADV", SQL )
  end select

  if (SaldoGeral IS NOTHING) then
    response.end
  end if

  SQL = "SELECT Limite, Saldo as SALDO_CREDITO FROM LimiteCredito "

  select case Filial
  case "05" : SET SaldoGeral = Conecta( "BEN-DADOSADV", SQL )
  case "06" : SET SaldoGeral = Conecta( "FOR-DADOSADV", SQL )
  end select

  if (not isObject(SaldoGeral)) or (SaldoGeral IS NOTHING) then
    response.end
  end if

  DadosFortalezaBentoGoncalves()

case "07"  

  SQL = "EXECUTE spLimiteCredito '" & trim(request("codcli")) & "', '" & trim(codLoja) & "'" 

  SET rstSaldoGeral = Conecta( "SAQ-DADOSADV", SQL )
  if (not isObject(rstSaldoGeral)) or (rstSaldoGeral IS NOTHING) then
    response.end
  end if

  SQL = "SELECT Limite, Saldo as SALDO_CREDITO FROM LimiteCredito" 
  SET SaldoGeral = Conecta( "SAQ-DADOSADV", SQL )
  if (not isObject(SaldoGeral)) or (SaldoGeral IS NOTHING) then
    response.end
  end if

  DadosSaquarema()      

end select

if SaldoGeral.EOF then 
%>   
  <script>
    alert("Sem detalhes a mostrar deste cliente")  
    document.getelentById("AnalisaCreditoV2.asp").style.display = "none";
  </script>
<%   
  response.end
end if 
%>
<html lang="pt-br">
  <head>
    <meta charset="utf-8" />
    <title>SWnet - Análise de Crédito</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="PRAGMA" content="no-cache" />
    <meta http-equiv="cache-control" content="no-cache" />

    <!-- Bootstrap 5.0 CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet">

    <style>
      .custom-tab-content {
          border: 1px solid #dee2e6;
          border-top: none;
          padding: 20px;
          border-radius: 0 0 5px 5px;
      }
      .nav-tabs .nav-link.active {
          background-color: #f8f9fa;
          border-color: #dee2e6 #dee2e6 #f8f9fa;
      }
      .divider {
        color: white;
        background-color: #00205B;
        align-items: center;
        justify-content: center;
        display: flex;
        border-radius: 3.5px;
      }
      label { 
        font-weight: bold; 
      } 
      /* Background escuro e transparente cobrindo toda a tela */
      #loader {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-color: rgba(0, 0, 0, 0.6); /* preto com 60% opacidade */

        display: flex;
        justify-content: center;
        align-items: center;

        z-index: 9999;
      }

      /* Spinner centralizado */
      .spinner {
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

      /* Animação de giro */
      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }

    </style>  

  </head>

  <body>

    <div class="row">

      <div class="col">
        <div class="text-center">
          <h5>DETALHES DA ANÁLISE DE CRÉDITO DO CLIENTE</h5>
          <p><%=request("nomcli")%></p>
        </div>
      </div>

    </div>

    <div class="container-fluid" style="max-width: 880px;">
      <div style="box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.1), 0 6px 10px 0 rgba(0, 0, 0, 0.19);">
        <div class="container-fluid" style="background-color: rgb(242, 242, 242); height: 85vh;">

          <div class="divider text-center">
            <b>Dados Gerais</b>
          </div>

          <div class="table-responsive tableFixHead">
            <table class="table table-bordered table-sm table-responsive-sm table-hover table-striped">
              <tbody>
                <tr>
                  <td><strong>Limite de Crédito:</strong></td>
                  <td style="text-align:right"><%=FormatNumber(SaldoGeral("LIMITE").Value,2)%></td>
                </tr>
                <tr>
                  <td><strong>Saldo Geral:</strong></td>
                  <td style="text-align:right"><%=FormatNumber(SaldoGeral("SALDO_CREDITO").Value,2)%></td>
                </tr>
                <tr>
                  <td><strong>Em Aberto Financeiro - vencidos:</strong></td>
                  <td style="text-align:right"><%=FormatNumber(SaldoFinanceiro1("SALDO_FINANCEIRO").Value,2)%></td>
                </tr>
                <tr>
                  <td><strong>Em Aberto Financeiro - à vencer:</strong></td>
                  <td style="text-align:right"><%=FormatNumber(SaldoFinanceiro2("SALDO_FINANCEIRO").Value,2)%></td>
                </tr>
                <tr>
                  <td><strong>Total em Aberto Financeiro:</strong></td>
                  <td style="text-align:right"><%=FormatNumber(SaldoFinanceiro("SALDO_FINANCEIRO").Value,2)%></td>
                </tr>
                <tr>
                  <td><strong>Pedidos Liberados em Aberto:</strong></td>
                  <td style="text-align:right"><%=FormatNumber(SaldoPedidoLiberado("SALDO_PEDIDO_LIBERADO").Value,2)%></td>
                </tr>
                <tr>
                  <td><strong>Pedidos à Liberar em Aberto:</strong></td>
                  <td style="text-align:right"><%=FormatNumber(SaldoPedidoALiberar("SALDO_PEDIDO_A_LIBERAR").Value,2)%></td>
                </tr>
              </tbody>
            </table>
          </div>

          <div class="divider text-center">
            <b>Analítico</b>
          </div>

          <!-- Navegação por abas -->
          <ul class="nav nav-tabs" id="myTab" role="tablist">
            <li class="nav-item"><a class="nav-link active" id="tvs-tab" data-bs-toggle="tab" href="#tvs" role="tab" aria-controls="tvs" aria-selected="true">Títulos Vencidos</a></li>
            <li class="nav-item"><a class="nav-link"        id="tav-tab" data-bs-toggle="tab" href="#tav" role="tab" aria-controls="tav" aria-selected="false">Títulos à Vencer</a></li>
            <li class="nav-item"><a class="nav-link"        id="pvs-tab" data-bs-toggle="tab" href="#pvs" role="tab" aria-controls="pvs" aria-selected="false">Pedidos Validados</a></li>
            <li class="nav-item"><a class="nav-link"        id="pav-tab" data-bs-toggle="tab" href="#pav" role="tab" aria-controls="pav" aria-selected="false">Pedidos à Validar</a></li>
          </ul>

          <!-- Conteúdo das abas -->
          <div class="tab-content custom-tab-content" id="myTabContent">
            <div class="tab-pane fade show active" id="tvs" role="tabpanel" aria-labelledby="tvs-tab">
              <div class="row">
                <div class="col">
                  <iframe frameborder="0" width="100%" height="400" name="analfin" scrolling="yes" src="AnalisaCreditoAnaliticoVencidos.asp?ParamFilial=<%=Filial%>&ParamCliente=<%=Cliente%>&ParamLoja=<%=codLoja%>&ParamNomePlanta=<%=NomePlanta%>"></iframe>   
                </div>
              </div>           
            </div>
            
            <div class="tab-pane fade" id="tav" role="tabpanel" aria-labelledby="tav-tab">
              <div class="row">
                <div class="col">
                  <iframe frameborder="0" width="100%" height="400" name="analfin" scrolling="yes" src="AnalisaCreditoAnaliticoFinanceiro.asp?ParamFilial=<%=Filial%>&ParamCliente=<%=Cliente%>&ParamLoja=<%=codLoja%>&ParamNomePlanta=<%=NomePlanta%>"></iframe>
                </div>
              </div>              
            </div>
            
            <div class="tab-pane fade" id="pvs" role="tabpanel" aria-labelledby="pvs-tab">
              <div class="row">
                <div class="col">
                  <iframe frameborder="0" width="100%" height="400" name="analfin" scrolling="yes" src="AnalisaCreditoAnaliticoValPed.asp?ParamFilial=<%=Filial%>&ParamCliente=<%=Cliente%>&ParamLoja=<%=codLoja%>"></iframe>
                </div>
              </div>             
            </div>

            <div class="tab-pane fade" id="pav" role="tabpanel" aria-labelledby="pav-tab">
              <div class="row">
                <div class="col">
                  <iframe frameborder="0" width="100%" height="400" name="analfin" scrolling="yes" src="AnalisaCreditoAnaliticoAValPed.asp?ParamFilial=<%=Filial%>&ParamCliente=<%=Cliente%>&ParamLoja=<%=codLoja%>"></iframe>
                </div>
              </div>            
            </div>
          </div>

          <br>

          <div class="row">
            <div class="col text-center">
              <a href="#" class="btn btn-sm btn-danger" title="Fechar" onclick="window.close();" ><i class="fa fa-close"></i>&nbsp;Fechar</a>              
            </div>
          </div>
          <br>
        </div>
      </div>
    </div>

    <!-- Bootstrap 5.0 JS Bundle with Popper -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js"></script>
    <!-- Font Awesome 6 -->
    <script src="https://kit.fontawesome.com/a076d05399.js" crossorigin="anonymous"></script>

  </body>
</html>
<%
SaldoGeral.Close()
Set SaldoGeral = Nothing
 
SaldoFinanceiro.Close()
Set SaldoFinanceiro = Nothing
 
SaldoPedidoLiberado.Close()
Set SaldoPedidoLiberado = Nothing
 
SaldoPedidoALiberar.Close()
Set SaldoPedidoALiberar = Nothing
 
'--------------------------------------------------------------------------------------------------' 
function DadosPirapetinga()

  '------------------------------------------------------------------------------------------------'
  SQL = "" &_
    "SELECT " &_
      "ISNULL(SUM(  CASE SE1980.E1_TIPO " &_
                      "WHEN 'AB-' THEN " &_
                        "(-1*E1_SALDO) " &_
                      "WHEN 'RA' THEN " &_
                        "(-1*E1_SALDO) " &_
                      "ELSE " &_
                        "CASE WHEN SUBSTRING(SE1980.E1_TIPO,3,1) = '-' THEN " &_
                          "(-1*E1_SALDO) " &_
                        "ELSE " &_
                          "E1_SALDO " &_
                        "END " &_
                    "END " &_
        "), 0 ) AS SALDO_FINANCEIRO " &_
    "FROM SE1980 " &_
    "WHERE (D_E_L_E_T_ <> '*') " &_
      "AND (E1_SALDO > 0) " &_
      "AND E1_TIPO <> 'NCC' " &_
      "AND (E1_CLIENTE = '" + replace( request("codcli"), "'", "''" ) + "') "

  SET SaldoFinanceiro = Conecta( "PIR-DADOSADV", SQL )

  if (not isObject(SaldoFinanceiro)) or (SaldoFinanceiro IS NOTHING) then
    reponse.end
  end if

  '------------------------------------------------------------------------------------------------'
  SQL = "" &_
    "SELECT " &_
      "ISNULL( SUM(  CASE SE1980.E1_TIPO " &_
                      "WHEN 'AB-' THEN " &_
                        "(-1*E1_SALDO) " &_
                      "WHEN 'RA' THEN " &_
                        "(-1*E1_SALDO) " &_
                      "ELSE " &_
                        "CASE WHEN SUBSTRING(SE1980.E1_TIPO,3,1) = '-' THEN " &_
                          "(-1*E1_SALDO) " &_
                        "ELSE " &_
                          "E1_SALDO " &_
                        "END " &_
                      "END), 0) AS SALDO_FINANCEIRO " &_
    "FROM SE1980 " &_
    "WHERE (D_E_L_E_T_ <> '*') " &_
      "AND (E1_SALDO > 0) " &_
      "AND E1_TIPO <> 'NCC' " &_
      "AND (E1_CLIENTE = '" + replace(request("codcli"), "'", "''") + "') " &_
      "AND E1_VENCREA < CONVERT(VARCHAR(8),GETDATE(),112) "

  SET SaldoFinanceiro1 = Conecta( "PIR-DADOSADV", SQL )

  if (not isObject(SaldoFinanceiro1)) or (SaldoFinanceiro1 IS NOTHING) then
    reponse.end
  end if

  '------------------------------------------------------------------------------------------------'
  SQL = "" &_
    "SELECT " &_
      "ISNULL( SUM(  CASE SE1980.E1_TIPO " &_
                      "WHEN 'AB-' THEN " &_
                        "(-1*E1_SALDO) " &_
                      "WHEN 'RA' THEN " &_
                        "(-1*E1_SALDO) " &_
                    "ELSE " &_
                      "CASE WHEN SUBSTRING(SE1980.E1_TIPO,3,1) = '-' THEN " &_
                        "(-1*E1_SALDO) " &_
                      "ELSE " &_
                        "E1_SALDO " &_
                      "END " &_
                    "END), 0 ) AS SALDO_FINANCEIRO " &_
    "FROM SE1980 " &_
    "WHERE (D_E_L_E_T_ <> '*') " &_
      "AND (E1_SALDO > 0) " &_
      "AND E1_TIPO <> 'NCC' " &_
      "AND (E1_CLIENTE = '" + replace(request("codcli"), "'", "''") + "') " &_
      "AND E1_VENCREA >= CONVERT(VARCHAR(8),GETDATE(),112) "

  SET SaldoFinanceiro2 = Conecta( "PIR-DADOSADV", SQL )

  if (not isObject(SaldoFinanceiro2)) or (SaldoFinanceiro2 IS NOTHING) then
    reponse.end
  end if

  '------------------------------------------------------------------------------------------------'
  SQL = "" &_
    "SELECT " &_
      "ISNULL(SUM(SALDO_PED),0) AS SALDO_PEDIDO_LIBERADO " &_
    "FROM ( " &_
      "SELECT " &_
        "CASE F4_IPI WHEN 'S' THEN " &_
          "SUM(ISNULL((C6_QTDVEN - C6_QTDENT) * SC6010.C6_PRCVEN,0)* (1+(B1_IPI/100))) " &_
        "ELSE " &_
          "SUM(ISNULL((C6_QTDVEN - C6_QTDENT) * SC6010.C6_PRCVEN,0)) " &_
        "END SALDO_PED " &_
      "FROM SC6010 " &_
        "INNER JOIN SC5010 ON SC6010.C6_NUM = SC5010.C5_NUM AND SC5010.D_E_L_E_T_ = '' " &_
        "INNER JOIN SF4010 ON SF4010.F4_CODIGO = SC6010.C6_TES AND SF4010.D_E_L_E_T_ = '' " &_
        "INNER JOIN SB1010 ON B1_COD = C6_PRODUTO AND SB1010.D_E_L_E_T_ = '' " &_
      "WHERE C6_QTDVEN - C6_QTDENT > 0 " &_
        "AND C5_ESTATUS NOT IN ('L','C') " &_
        "AND (C5_NUM BETWEEN '000000' AND '999999' OR C5_NUM BETWEEN 'AA0000' AND 'ZZ9999') " &_
        "AND C5_TIPO NOT IN ('D','B') " &_
        "AND SC6010.D_E_L_E_T_ = '' " &_
        "AND C6_CLI = '"& replace(trim(request("codcli")), "'", "''") &"' " &_
        "AND SF4010.F4_DUPLIC = 'S' " &_
    "GROUP BY F4_IPI) AS LIMITE "

  SET SaldoPedidoLiberado = Conecta( "PIR-DADOSADV", SQL )

  if (not isObject(SaldoPedidoLiberado)) or (SaldoPedidoLiberado IS NOTHING) then
    reponse.end
  end if

  '------------------------------------------------------------------------------------------------'
  SQL = "" &_
    "SELECT " &_
      "ISNULL(SUM(TOTAL),0) SALDO_PEDIDO_A_LIBERAR " &_
    "FROM ( " &_
      "SELECT " &_
        "CASE F4_IPI WHEN 'S' THEN " &_
          "SUM(ISNULL(ZF_QUANT * ZF_PRECO,0)* (1+(B1_IPI/100))) " &_
        "ELSE " &_
          "SUM(ISNULL(ZF_QUANT * ZF_PRECO,0)) " &_
        "END TOTAL " &_
      "FROM SZF010 " &_
        "INNER JOIN SF4010 ON SF4010.F4_CODIGO = SZF010.ZF_TES " &_
        "INNER JOIN SB1010 ON B1_COD = ZF_PRODUTO AND SB1010.D_E_L_E_T_ = '' " &_
      "WHERE (ZF_ESTATUS <> 'OK') " &_
        "AND (SZF010.D_E_L_E_T_ = '') " &_
        "AND (ZF_CLIENTE = '"&replace(trim(request("codcli")), "'", "''")&"') " &_
        "AND (F4_DUPLIC = 'S') " &_
    "GROUP BY F4_IPI ) AS LIMITE "

  SET SaldoPedidoALiberar = Conecta( "PIR-DADOSADV", SQL )

  if (not isObject(SaldoPedidoALiberar)) or (SaldoPedidoALiberar IS NOTHING) then
    reponse.end
  end if

end function    
'--------------------------------------------------------------------------------------------------'
Function DadosFortalezaBentoGoncalves()
  
  SQL = "" &_
    "SELECT " &_
      "isnull(SALDO_FINANCEIRO1,0)+isnull(SALDO_FINANCEIRO2,0) AS SALDO_FINANCEIRO " &_
    "FROM ( " &_
      "SELECT " &_
        "SUM(SALDO) AS SALDO_FINANCEIRO1, " &_
        "( " &_
          "SELECT " &_
            "SUM(SALDO) " &_
          "FROM VW_DETALHE_TITULOS_LC " &_
          "WHERE COD_CLIENTE = '" &  request("codcli") & "' " &_
            "AND CONVERT(DATE, VENCTO, 103) >= CONVERT(DATE, GETDATE(), 103) " &_
            "AND SALDO > 0 " &_
        ") AS SALDO_FINANCEIRO2 " &_
    "FROM VW_DETALHE_TITULOS_LC " &_
    "WHERE COD_CLIENTE = '" & request("codcli") & "' " &_
      "AND CONVERT(DATE, VENCTO, 103) < CONVERT(DATE, GETDATE(), 103) " &_
      "AND SALDO > 0) as SALDO"

  select case Filial
  case "05" : SET SaldoFinanceiro = Conecta( "BEN-CORPORATIVO", SQL )
  case "06" : SET SaldoFinanceiro = Conecta( "FOR-CORPORATIVO", SQL )
  end select

  if (not isObject(SaldoFinanceiro)) or (SaldoFinanceiro IS NOTHING) then
    response.end
  end if

  SQL = "" &_
    "SELECT " &_
      "isnull(sum(SALDO),0) AS SALDO_FINANCEIRO " &_
    "FROM VW_DETALHE_TITULOS_LC " &_
    "WHERE COD_CLIENTE = '" & request("codcli") & "'" & " " &_
      "AND CONVERT(DATE, VENCTO, 103) < CONVERT(Date, GETDATE(), 103) " &_
      "AND SALDO > 0"

  select case Filial
  case "05" : SET SaldoFinanceiro1 = Conecta( "BEN-CORPORATIVO", SQL )
  case "06" : SET SaldoFinanceiro1 = Conecta( "FOR-CORPORATIVO", SQL )
  end select

  if (not isObject(SaldoFinanceiro1)) or (SaldoFinanceiro1 IS NOTHING) then
    response.end
  end if

  SQL = "" &_
    "SELECT " &_
      "isnull(sum(SALDO),0) as SALDO_FINANCEIRO " &_
    "FROM VW_DETALHE_TITULOS_LC " &_
    "WHERE COD_CLIENTE = '" &  request("codcli") & "' " &_
      "AND CONVERT(DATE, VENCTO, 103) >= CONVERT(Date, GETDATE(), 103) " &_
      "AND SALDO > 0"

  select case Filial
  case "05" : SET SaldoFinanceiro2 = Conecta( "BEN-CORPORATIVO", SQL )
  case "06" : SET SaldoFinanceiro2 = Conecta( "FOR-CORPORATIVO", SQL )
  end select

  if (not isObject(SaldoFinanceiro2)) or (SaldoFinanceiro2 IS NOTHING) then
    response.end
  end if

  SQL = "" &_
    "SELECT " &_
      "ISNULL(SUM(VALOR_SALDO),0) AS SALDO_PEDIDO_LIBERADO " &_
    "FROM VW_DETALHE_PEDIDO_LC "  &_
    "WHERE COD_CLI = '" + request("codcli") & "' " &_
      "AND CIDADE = '" & NomePlanta & "' "

  select case Filial
  case "05" : SET SaldoPedidoLiberado = Conecta( "BEN-CORPORATIVO", SQL )
  case "06" : SET SaldoPedidoLiberado = Conecta( "FOR-CORPORATIVO", SQL )
  end select

  if (not isObject(SaldoPedidoLiberado)) or (SaldoPedidoLiberado IS NOTHING) then
    response.end
  end if

  SQL = "" &_
    "SELECT " &_
      "ISNULL(SUM(VALOR),0) AS SALDO_PEDIDO_A_LIBERAR " &_
    "FROM VW_DETALHE_PEDIDO_VALIDAR_LC " &_
    "WHERE CODIGO_CLIENTE = '" + request("codcli") & "' " &_
      "AND CIDADE = '" & NomePlanta & "' "

  select case Filial
  case "05" : SET SaldoPedidoALiberar = Conecta( "BEN-CORPORATIVO", SQL )
  case "06" : SET SaldoPedidoALiberar = Conecta( "FOR-CORPORATIVO", SQL )
  end select

  if (not isObject(SaldoPedidoALiberar)) or (SaldoPedidoALiberar IS NOTHING) then
    response.end
  end if

end function 
'--------------------------------------------------------------------------------------------------'
function DadosSaquarema()
 
  SQL = "" &_
    "SELECT " &_
    "isnull(SALDO_FINANCEIRO1,0)+isnull(SALDO_FINANCEIRO2,0) AS SALDO_FINANCEIRO " &_
    "FROM ( " &_
      "SELECT SUM(SALDO) AS SALDO_FINANCEIRO1, " &_
      "( " &_
        "SELECT SUM(SALDO) " &_
        "FROM VW_DETALHE_TITULOS_LC " &_
        "WHERE COD_CLIENTE = '" &  request("codcli") & "' " &_
          "AND LOJA_CLIENTE = '" + codLoja & "' " &_
          "AND CONVERT(DATE, VENCTO, 103) >= CONVERT(DATE, GETDATE(), 103) " &_
      ") AS SALDO_FINANCEIRO2 " &_
      "FROM VW_DETALHE_TITULOS_LC " &_
      "WHERE COD_CLIENTE = '" & request("codcli") & "' " &_
        "AND LOJA_CLIENTE = '" + codLoja & "' " &_
        "AND CONVERT(DATE, VENCTO, 103) < CONVERT(DATE, GETDATE(), 103) " &_
    ") AS SALDO"
  
  SET SaldoFinanceiro = Conecta( "SAQ-CORPORATIVO", SQL )

  if (NOT isObject(SaldoFinanceiro)) OR (SaldoFinanceiro IS NOTHING) then
    response.end
  end if

  SQL = "" &_
    "SELECT " &_
      "isnull(sum(SALDO),0) as SALDO_FINANCEIRO " &_
    "FROM VW_DETALHE_TITULOS_LC " &_
    "WHERE COD_CLIENTE = " & + replace(request("codcli"), "'", "''") & " " &_
      "AND LOJA_CLIENTE = '" + codLoja & "' " &_
      "AND CONVERT(DATE, VENCTO, 103) < CONVERT(Date, GETDATE(), 103) "

  SET SaldoFinanceiro1 = Conecta( "SAQ-CORPORATIVO", SQL )

  if (NOT isObject(SaldoFinanceiro1)) OR (SaldoFinanceiro1 IS NOTHING) then
    response.end
  end if

  SQL = "" &_
    "SELECT " &_
      "isnull(sum(SALDO),0) as SALDO_FINANCEIRO " &_
    "FROM VW_DETALHE_TITULOS_LC " &_
    "WHERE COD_CLIENTE = " & + replace(request("codcli"), "'", "''") & " " &_
      "AND LOJA_CLIENTE = '" + codLoja & "'" &_
      "AND CONVERT(DATE, VENCTO, 103) >= CONVERT(Date, GETDATE(), 103) "

  SET SaldoFinanceiro2 = Conecta( "SAQ-CORPORATIVO", SQL )

  if (NOT isObject(SaldoFinanceiro2)) OR (SaldoFinanceiro2 IS NOTHING) then
    response.end
  end if

  SQL = "SELECT " &_
        "ISNULL(SUM(VALOR_SALDO),0) AS SALDO_PEDIDO_LIBERADO " &_
        "FROM VW_DETALHE_PEDIDO_LC " &_
        "WHERE COD_CLI = '" + request("codcli") & codLoja & "' " &_
        "AND LOJA_CLI = '" + codLoja & "'"

  SET SaldoPedidoLiberado = Conecta( "SAQ-CORPORATIVO", SQL )

  if (NOT isObject(SaldoPedidoLiberado)) OR (SaldoPedidoLiberado IS NOTHING) then
    response.end
  end if

  'SQL = "SELECT 0 AS SALDO_PEDIDO_A_LIBERAR"
  SQL = "SELECT " &_
        "ISNULL(SUM(VALOR),0) AS SALDO_PEDIDO_A_LIBERAR " &_
        "FROM VW_DETALHE_PEDIDO_VALIDAR_LC " &_
        "WHERE CODIGO_CLIENTE = '" + request("codcli") & "' " &_
        "AND CIDADE = '" & NomePlanta & "' "
  
  SET SaldoPedidoALiberar = Conecta( "SAQ-CORPORATIVO", SQL )

  if (NOT isObject(SaldoPedidoALiberar)) OR (SaldoPedidoALiberar IS NOTHING) then
    response.end
  end if

end function
%>