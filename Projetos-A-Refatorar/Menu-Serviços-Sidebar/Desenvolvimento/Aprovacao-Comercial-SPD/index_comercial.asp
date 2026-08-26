<!--#Include File="../cnn/cnn.asp"-->
<!--#Include File="../funcoes/funcoes.asp"-->
<!--#Include File="../includes/variaveis.asp"-->
<%session.LCID = 1046
if UNIDADE_FILIAL = "03" or  UNIDADE_FILIAL = "02"  or  UNIDADE_FILIAL = "01" then
    UNIDADE_FILIAL = "00"
end if

if CODUSER = "008097" then
    UNIDADE_FILIAL = "05"
end if

UNIDADE_FILIAL = "00"

Select case UNIDADE_FILIAL
    case "00" set consulta  = Dados11:  set Consulta1 = Dados11
    case "04" set consulta  = Dados11:  set Consulta1 = Dados11
    case "05" set consulta  = Dados91:  set Consulta1 = Dados131
    case "06" set consulta  = Dados121: set Consulta1 = Dados201
    case "07" set consulta  = Dados191: set Consulta1 = Dados181
End select    
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
  <title>Smurfit Kappa</title>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css"/>
  <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />

  <link rel="stylesheet" type="text/css" href="../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.css" />
  <link rel="stylesheet" type="text/css" href="../../dhtmlx/dhtmlxCalendar/codebase/skins/dhtmlxcalendar_dhx_web.css" />
  <script src="../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.js"></script>
  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js"></script>
</head>
<body>
  <%
  'Variáveis com data e hora
  Ano = Year(Date())
  Mes = Month(Date())
  Dia = Day(Date())
  If Len(cStr(Mes)) = 1 Then
    Mes = "0" & cStr(Mes)
  End If
  If Len(cStr(Dia)) = 1 Then
    Dia = "0" & cStr(Dia)
  End If
  Hora = Time()

  Select Case Request("motivo")
    'Motivos Reprovação
    Case "D"
      motivo = "Emissão em Duplicidade"
    Case "P" 
      motivo = "Preço não Aprovado"
    Case "E" 
      motivo = "Descontinuidade Projeto"
    Case "R" 
      motivo = "Sem Retorno do Comercial"
    Case "N" 
      motivo = "Sem Retorno Cliente"
    Case "O" 
      motivo = "Proposta Descartada"
    Case "F" 
      motivo = "Cadastro Cliente Finalizado"
    Case "I" 
      motivo = "Processo Interno SK Teste"
    Case "T" 
      motivo = "Teste"
    'Motivos aprovação
    Case "1"
      motivo = "Cliente Novo - Plano"
    Case "2"
      motivo = "Cliente Novo - Explorar"
    Case "3"
      motivo = "Cliente Novo - Qualificar"
    Case "5"
      motivo = "Cliente Novo - Desenvolver"
    Case "6"
      motivo = "Cliente Novo - Propor"
    Case "7"
      motivo = "Cliente Novo - Fechar"
    Case "8"
      motivo = "Cliente Atual - Melhoria"
    Case "9"
      motivo = "Cliente Atual - Novo SKU"
    Case Else
      motivo = ""
  End Select

  'Cancela a solicitação
  If Request("cancela") <> "" Then
    SQL = "INSERT INTO WorkFlowSapComercial VALUES ('" & Request("cod") & "', '" & Mid(session("NOME"),1,20) & "', 'V', '" & Request("motivo") & "', '" & Ano & Mes & Dia & "','" & Hora & "') "
    Consulta1.Open(SQL)

    SQL = "UPDATE SKFT1 set FT1_STATUS = 'V' WHERE FT1_ID = " & Request("cod") & " "
    Consulta1.Source = SQL
    Consulta1.Open() 
    
    'Enviar e-mail para quem abriu a SAP (vendedor) informando do cancelamento
    SQL = "SELECT FT1_USU,FT1_PLANTA,FT1_B5CLI,FT1_B5NOMECLI,FT1_LOJA FROM SKFT1 "
    SQL = SQL & "WHERE FT1_ID = " & Request("cod") & " "
    Consulta1.Open(SQL)  
    If Not Consulta1.Eof Then
      SQL = "SELECT CRM_EMAILPRIMARIO, CRM_EMAILSECUNDARIO FROM HELPDESK..CRM_DTI WHERE CRM_LOGIN = '"&Consulta1("FT1_USU")&"' and EMP_CODIGO = '"&Consulta1("FT1_PLANTA")&"' "
      Dados21.Source = SQL
      Dados21.Open()
      if not Dados21.eof then
          email = Trim(Dados21("CRM_EMAILPRIMARIO"))
          if Trim(Dados21("CRM_EMAILSECUNDARIO")) <> "" then
            email = email &";"&Trim(Dados21("CRM_EMAILSECUNDARIO"))
          end if
          If email <> "" Then
            Assunto = "A SAP número " & Request("cod") & " foi cancelada"
            mensagem = "<html>"
            mensagem = mensagem & "<head>"
            mensagem = mensagem & "<title></title>"
            mensagem = mensagem & "</head>"
            mensagem = mensagem & "<body style='font-family:Verdana;font-size:12px'>"
            mensagem = mensagem & "<p style='text-align:justify'>A SAP número número <b>" & Request("cod") & "</b> foi cancelada pelo motivo <b>" & motivo & "</b>.<br>"
            mensagem = mensagem & "Essa SAP não foi enviada para desenvolvimento.<br />"
            mensagem = mensagem & "Cliente <b>" & Consulta1("FT1_B5CLI") & " - " & Left(Consulta1("FT1_B5NOMECLI"),30) & "</b>, loja <b>" & Consulta1("FT1_LOJA") & "</b>.</p>"
            mensagem = mensagem & "<p style='text-align:justify'>E-mail enviado automaticamente pelo sistema de automação de processos.<br /><strng>Smurfit Kappa do Brasil</strong></p>"
            CALL EnviaEmail (email, "", "", "", "noreply.brasil@smurfitwestrock.com.br", mensagem, assunto)
          End If  
      end if 
      Dados21.Close()
    End If
    Consulta1.Close()
    Response.Write "<script>alert('Processo de não aprovação de SAP realizado com sucesso')</script>"
    Response.Write "<script>location.href='index_comercial.asp'</script>"
  End If
 
  'Aprova a solicitação
  If Request("aprova") <> "" Then
    SQL = "INSERT INTO WorkFlowSapComercial VALUES ('" & Request("cod") & "', '" & Mid(session("NOME"),1,20) & "', 'C', '" & Request("motivo") & "', '" & Ano & Mes & Dia & "','" & Hora & "'); "
    Consulta1.Source = SQL
    Consulta1.Open()

    SQL = "UPDATE SKFT1 set FT1_STATUS = 'A' WHERE FT1_ID = " & Request("cod") & " "
    Consulta1.Source = SQL
    Consulta1.Open()  

    'Enviar e-mail para pessoal do desenvolvimento
    SQL = "SELECT FT1_PLANTA, FT1_B5CLI, FT1_B5NOMECLI, FT1_LOJA, FT1_B5REF, FT1_A1VEND, FT1_A3NREDUZ, Endereco "
    SQL = SQL & "FROM SKFT1, WorkFlowPreCadastroProduto "
    SQL = SQL & "WHERE Planta = FT1_PLANTA AND Tipo = 1 AND FT1_ID = " & Request("cod") & " "
    Consulta1.Source = SQL
    Consulta1.Open() 

    If Not Consulta1.Eof Then
      email = Trim(Consulta1("Endereco"))
      If email <> "" Then
        assunto = "A SAP número " & Request("cod") & " foi aprovada"
        mensagem = "<html>"
        mensagem = mensagem & "<head>"
        mensagem = mensagem & "<title></title>"
        mensagem = mensagem & "</head>"
        mensagem = mensagem & "<body style='font-family:Verdana;font-size:12px'>"
        mensagem = mensagem & "<p style='text-align:justify'>A SAP número número <b>" & Request("cod") & "</b> foi aprovada. <b>" & motivo & "</b><br />"
        mensagem = mensagem & "Cliente <b>" & Consulta1("FT1_B5CLI") & " - " & Left(Consulta1("FT1_B5NOMECLI"),30) & "</b>, loja <b>" & Consulta1("FT1_LOJA") & "</b>.</p>"
        mensagem = mensagem & "<p style='text-align:justify'>E-mail enviado automaticamente pelo sistema de automação de processos.<br /><strng>Smurfit Kappa do Brasil</strong></p>"
      
       CALL EnviaEmail (email, "", "", "", "noreply.brasil@smurfitwestrock.com.br", mensagem, assunto)
      End If   
    End If
    Consulta1.Close()
    
    Response.Write "<script>alert('Processo aprovação de SAP realizado com sucesso')</script>"
    Response.Write "<script>location.href='index_comercial.asp'</script>"

  End If
  %>

  <% '----Cabeçalho e Filtros------------------%>
  <div class="jumbotron text-center">
    <h3 style="color:gray">Aprovação Comercial - Solicitação de Projeto ao Desenvolvimento (SPD)</h3>
  </div>
  <div class="container-fluid">
  <div style="box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.1), 0 6px 10px 0 rgba(0, 0, 0, 0.19);">
    <div class="container-fluid" style="background-color: rgb(242, 242, 242);">
    <form action="index_comercial.asp" id="formfil">
    <input type="hidden" name="envia" value="s"/>
    
    <div class="row">
      <div class="col-sm-4">
        <label>SAP</label>
        <input type="text" class="form-control" name="Id" id="Id" value="<%=Trim(Request("Id"))%>" placeholder="Informe a SAP..."/>
      </div>
      <div class="col-sm-4">
        <label>FT Origem</label>
        <input type="text" class="form-control" name="ftorigem" id="ftorigem" value="<%=Trim(Request("ftorigem"))%>" placeholder="Informe a FT..."/>
      </div>
    </div>
    
    <div class="row">
      <div class="col-sm-4">
        <label>Cliente</label>
        <input type="text" class="form-control" name="Cliente" id="Cliente" value="<%=Trim(Request("Cliente"))%>" placeholder="Informe o código do cliente..."/>
      </div>
      <div class="col-sm-4">
        <label>Status</label>
        <select name="status" class="form-control" onchange="form.submit()">
          <option value="T" <%If trim(request("status")) = "T" Then%>selected<%End If%>>Status (Todos)</option>
          <option value="" <%If trim(request("status")) = "" Then%>selected<%End If%>>Pendente</option>
          <option value="C" <%If trim(request("status")) = "C" Then%>selected<%End If%>>Aprovado</option>
          <option value="V" <%If trim(request("status")) = "V" Then%>selected<%End If%>>Reprovada</option>

        </select>
      </div>
      <div class="col-sm-4">
        <label>Planta</label>
        <div class="input-group-prepend">
          <select name="Planta" class="form-control col-lg-5" onchange="form.submit()">
            <option value="" <%If Request("Planta") = "" Then%>selected<%End If%>>Todas</option>
            <option value="00" <%If Request("Planta") = "00" Then%>selected<%End If%>>Pirapetinga</option>
            <option value="04" <%If Request("Planta") = "04" Then%>selected<%End If%>>Uberaba</option>
            <option value="05" <%If Request("Planta") = "05" Then%>selected<%End If%>>Bento Gonçalves</option>
            <option value="06" <%If Request("Planta") = "06" Then%>selected<%End If%>>Fortaleza</option>
            <option value="07" <%If Request("Planta") = "07" Then%>selected<%End If%>>Saquarema</option>
          </select>&nbsp;
          <button type="submit" class="btn btn-info" title="Filtrar"><i class="fas fa-filter"></i></button>&nbsp;
        </div>          
      </div>
    </div>
    <div>&nbsp;</div>
    </form>
  </div>
  <% '----Fim Cabeçalho e Filtros------------------%>

  <% '----Exibição das SAPs ------------------%>
  <div class="table-responsive" style="background-color:lightgray">
  <table class="table table-bordered table-hover" style="font-size:12px;background-color:white">
    <tr style="background-color:gray;color:white;font-weight:bold">
      <td class="text-left">SAP</td>
      <td class="text-left vscr">Status</td>
      <td class="text-center vscr">Data / Motivo</td>
      <td class="text-left">Emissão</td>
      <td class="text-left">Cliente</td>
      <td class="text-left vscr">Vendedor</td>
      <td class="text-left vscr">Orçamento</td>
      <td class="text-left vscr">Estrutura</td>
      <td class="text-left vscr">Preço Kg</td>
      <td class="text-center vscr">Ação</td>
    </tr>
  <%
  SQL = "SELECT * FROM "
  SQL = SQL & "(SELECT FT1_ID, FT1_FT, FT1_TIPO, FT1_DATA, FT1_B5REF, FT1_B5CLI, FT1_USU, FT1_STATUS, FT1_STATUS2, FT1_FTNOVA, FT1_B5ESTRUTU, "
  SQL = SQL & "FT1_A1VEND, FT1_ORCAMENTO, FT1_PLANTA, Isnull(FT1_DTCONFEC,'') FT1_DTCONFEC, FT1_MOTCANC, FT1_PRECOKG, "
  SQL = SQL & "ISNULL(FT1_A3NREDUZ,'') AS FT1_A3NREDUZ, FT1_QTDAMO, Isnull(FT1_QuemEnviouAmostra,'') FT1_QuemEnviouAmostra, "
  SQL = SQL & "Isnull(FT1_DataEnvioAmostra,'') FT1_DataEnvioAmostra, FT1_B5NOMECLI, FT1_MOTIVOREPROVACAOAMOSTRA2, "
  SQL = SQL & "ISNULL((SELECT TOP 1 Status FROM WorkFlowSapComercial wf WHERE wf.idFT = SKFT1.FT1_ID ORDER BY DataStatus Desc, HoraStatus Desc),'') FT1_STATUSCOMERCIAL, "
  SQL = SQL & "ISNULL((SELECT TOP 1 DataStatus FROM WorkFlowSapComercial wf WHERE wf.idFT = SKFT1.FT1_ID ORDER BY DataStatus Desc, HoraStatus Desc),'') FT1_DATASTATUS, "
  SQL = SQL & "ISNULL((SELECT TOP 1 HoraStatus FROM WorkFlowSapComercial wf WHERE wf.idFT = SKFT1.FT1_ID ORDER BY DataStatus Desc, HoraStatus Desc),'') FT1_HORASTATUS, "
  SQL = SQL & "ISNULL((SELECT TOP 1 MotivoReprovacao FROM WorkFlowSapComercial wf WHERE wf.idFT = SKFT1.FT1_ID ORDER BY DataStatus Desc, HoraStatus Desc),'') FT1_MOTIVOREPROVACAO, "
  SQL = SQL & "ISNULL((SELECT TOP 1 Usuario FROM WorkFlowSapComercial wf WHERE wf.idFT = SKFT1.FT1_ID ORDER BY DataStatus Desc, HoraStatus Desc),'') FT1_USUARIO "
  SQL = SQL & "FROM SKFT1  "

  SQL = SQL & "WHERE ISNULL(FT1_DEL,0) = 0 AND FT1_STATUS IN ('S','O','C','A','V') "
  If Trim(Request("id")) <> "" Then
    SQL = SQL & "AND (FT1_ID = '" & Request("id") & "') "
  Else
    If Request("Planta") <> "" Then
      SQL = SQL & "AND FT1_PLANTA = '" & Request("Planta") & "' "
    End If

    If Trim(Request("ftorigem")) <> "" Then
      SQL = SQL & "AND (FT1_FT = '" & UCase(Request("ftorigem")) & "' OR FT1_B5REF LIKE '%" & Request("ftorigem") & "%') "
    End If

    If Trim(Request("Cliente")) <> "" Then
      SQL = SQL & "AND (FT1_B5CLI = '" & Request("Cliente") & "') "
    End If

  End If    
  SQL = SQL &") AS CONSULTA "     
  SQL = SQL &"WHERE FT1_DATA >= '20220509' "
      
  If  Trim(Request("status")) <> "T"  Then
      SQL = SQL & "AND  FT1_STATUSCOMERCIAL = '"&Request("status")&"' "
  END IF   
  SQL = SQL & "ORDER BY FT1_ID DESC"
  
  Consulta1.Source = SQL
  Consulta1.Open()
  While Not Consulta1.EOF
    FT1_STATUSCOMERCIAL = Consulta1("FT1_STATUSCOMERCIAL")
    FT1_DATASTATUS = Consulta1("FT1_DATASTATUS")
    FT1_HORASTATUS = Consulta1("FT1_HORASTATUS")
    FT1_MOTIVOREPROVACAO = Consulta1("FT1_MOTIVOREPROVACAO")
    FT1_USUARIO = Consulta1("FT1_USUARIO")

    'Status Comercial
    Select Case FT1_STATUSCOMERCIAL
        Case "C"
        StatusComercial = "Aprovado Comercial"
        statusbg = "alert-success"
        Case "V"
        StatusComercial = "Reprovado Comercial"
        statusbg = "alert-danger"
        Case Else
        StatusComercial = "Pendente Comercial"
        statusbg = "alert-warning"
    End Select
    
    %>
    <tr class="<%=statusbg%>" id="<%=Consulta1("FT1_ID") %>">
      <td class="text-center align-middle"><a href="EspelhoPreCadastro.asp?Id=<%=Consulta1("FT1_ID") %>&Planta=<%=Consulta1("FT1_PLANTA")%>" target="_blank" title="Visualizar espelho"> <%=Consulta1("FT1_ID")%></a></td>
      <td class="text-left align-middle vscr"><%=Consulta1("FT1_STATUSCOMERCIAL") %> - <%=StatusComercial%></td>
      <td class="text-left align-middle vscr"><%=FT1_USUARIO%><br><%=AcertaData(FT1_DATASTATUS)%>-<%=FT1_HORASTATUS%></td>
      <td class="text-center align-middle"><%=AcertaData(Consulta1("FT1_DATA"))%></td>
      <td class="text-left align-middle" title="<%=Consulta1("FT1_B5CLI")%>-<%=Consulta1("FT1_B5NOMECLI")%>"><%=Consulta1("FT1_B5CLI")%>-<%=Left(Consulta1("FT1_B5NOMECLI"),15)%></td>
      <td class="text-left align-middle vscr"><%=Consulta1("FT1_A1VEND")%> - <%=Consulta1("FT1_A3NREDUZ")%></td>
      <td class="text-left align-middle vscr">Não</td>
      <td class="text-left align-middle vscr"><%=Consulta1("FT1_B5ESTRUTU")%></td>
      <td class="text-right align-middle vscr"><%=FormatNumber(Consulta1("FT1_PRECOKG"),2)%></td>
      <td class="text-left text-nowrap align-middle vscr">
        <%if FT1_STATUSCOMERCIAL <> "C" then %>
        <button type="button" class="btn btn-success btn-sm" title="Aprovar" onclick="maprov('true','<%=Consulta1("FT1_ID")%>')">
          <i class="fas fa-check"></i>
        </button>
            <%if FT1_STATUSCOMERCIAL <> "C" then %>
                <button type="button" class="btn btn-danger btn-sm" title="Cancelar" onclick="mcancel('true','<%=Consulta1("FT1_ID")%>')">
                  <i class="fas fa-times"></i>
                </button>
            <%end if %>
        <%end if %>
        <button type="button" class="btn btn-light btn-sm" title="Visualizar" onclick="location.href='index.asp?voltar=1&act=ALT&save=false&cod=<%=Consulta1("FT1_ID")%>&tipo=<%=Consulta1("FT1_TIPO")%>'">
            <i class="far fa-eye"></i>
        </button>
      </td>
    </tr>
    <%
    Consulta1.MoveNext()
  Wend
  Consulta1.Close()
  %>
  </table>
  </div>
  <% '----Fim de Exibição das SAPs ------------------%>

  <%'------------Modal de Cancelamento---------------%>
  <div class="modal fade" id="modal1">
    <div class="modal-dialog">
      <div class="modal-content">
        <div class="modal-header bg-danger">
          <h6 class="modal-title">Não Aprovar SAP</h6>
          <button type="button" class="close" data-dismiss="modal">&times;</button>
        </div>
        <div class="modal-body">
          <input type="hidden" id="sapcod">
          <input type="hidden" id="tc">
          <p><span id="txtcanc"></span></p>
          <select id="FT1_MOTCANC" style="padding: 5px;">
            <option value="" style="display: none;">Selecione a justificativa...</option>
            <option value="D">Emissão em Duplicidade</option>
            <option value="P">Preço Não Aprovado</option>
            <option value="E">Descontinuidade do Projeto</option>
            <option value="R">Sem Retorno Comercial</option>
            <option value="N">Sem Retorno Cliente</option>
            <option value="O">Proposta Descartada</option>
            <option value="F">Cadastro Cliente Finalizado</option>
            <option value="I">Processo Interno SK - Teste</option>
            <option value="T">Teste</option>
          </select>
          <span id="alertmot" style="color: red;display: none;"><b>Justificativa obrigatória</b></span>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" style="width: 100px;" onclick="ccancel()">Sim</button>
          <button type="button" class="btn btn-light" data-dismiss="modal" style="width: 100px;">Não</button>
        </div>
      </div>
    </div>
  </div>

  <script>
    function mcancel(t,id) {
      document.getElementById('alertmot').style.display = 'none';
      document.getElementById('sapcod').value = id;
      document.getElementById('tc').value = t;
      if (t == 'true') {
        var msg = 'Confirma a NÃO aprovação da SAP ' + id + ' ?';
        document.getElementById('FT1_MOTCANC').style.display = '';
      } else {
        var msg = 'Confirma reverter a não aprovação da SAP ' + id + ' ?';
        document.getElementById('FT1_MOTCANC').style.display = 'none';
      }       
      document.getElementById('txtcanc').innerHTML = msg;
      $('#modal1').modal({backdrop: "static"});
    }
    function ccancel() {
      if (document.getElementById('FT1_MOTCANC').value == '' && document.getElementById('tc').value == 'true') {
        document.getElementById('alertmot').style.display = '';
      } else {
        location.href = 'index_comercial.asp?cancela='+document.getElementById('tc').value+'&cod='+document.getElementById('sapcod').value+'&motivo='+document.getElementById('FT1_MOTCANC').value;
      }
    }
  </script>

  <%'------------Modal de Aprovação---------------%>
  <div class="modal fade" id="modal2">
    <div class="modal-dialog">
      <div class="modal-content">
        <!--<div class="modal-header" style="background-color: lightgray;">-->
        <div class="modal-header bg-success">
          <h6 class="modal-title">Aprovar SAP</h6>
          <button type="button" class="close" data-dismiss="modal">&times;</button>
        </div>
        <div class="modal-body">
          <input type="hidden" id="sapcoda">
          <input type="hidden" id="ta">
          <p><span id="txtaprov"></span></p>
          <select id="FT1_MOTAPROV" style="padding: 5px;">
            <option value="" style="display: none;">Pipeline Comercial...</option>
            <optgroup label="Cliente Novo">
              <option value="1">Plano</option>
              <option value="2">Explorar</option>
              <option value="3">Qualificar</option>
              <option value="5">Desenvolver</option>
              <option value="6">Propor</option>
              <option value="7">Fechar</option>
            </optgroup>
            <optgroup label="Cliente Atual">
              <option value="8">Melhoria</option>
              <option value="9">Novo SKU</option>
            </optgroup>
          </select>
          <span id="alertmota" style="color: red;display: none;"><b>Seleção obrigatória</b></span>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" style="width: 100px;" onclick="caprov()">Sim</button>
          <button type="button" class="btn btn-light" data-dismiss="modal" style="width: 100px;">Não</button>
        </div>
      </div>
    </div>
  </div>

  <script>
    function maprov(t,id) {
      document.getElementById('alertmot').style.display = 'none';
      document.getElementById('sapcoda').value = id;
      document.getElementById('ta').value = t;
      if (t == 'true') {
        var msg = 'Confirma a aprovação da SAP ' + id + ' ?';
        document.getElementById('FT1_MOTCANC').style.display = '';
      } else {
        var msg = 'Confirma reverter a aprovação da SAP ' + id + ' ?';
        document.getElementById('FT1_MOTAPROV').style.display = 'none';
      }       
      document.getElementById('txtaprov').innerHTML = msg;
      $('#modal2').modal({backdrop: "static"});
    }
    function caprov() {
      if (document.getElementById('FT1_MOTAPROV').value == '') {
        document.getElementById('alertmota').style.display = '';
      } else {
        location.href = 'index_comercial.asp?aprova='+document.getElementById('ta').value+'&cod='+document.getElementById('sapcoda').value+'&motivo='+document.getElementById('FT1_MOTAPROV').value;
      }
    }
  </script>
</body>
</html>