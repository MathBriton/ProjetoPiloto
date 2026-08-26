<!-- v3.0 Data: 16-07-2021-->
<!-- Alterações decritas no Chpl 201 -->
<!--#Include file="../cnn/cnn.asp"-->
<!--#Include file="../cnn2/cnn2.asp"-->
<!--#Include file="../funcoes/funcoes.asp"-->
<!--#Include file="../includes/variaveis.asp"-->
<!--#Include file="includes/asp/Unificado.asp"-->

<%session.LCID = 1046
temNivelFinanceiro=false
temNivelFiscal=false
dadosEnderecoEntrega = ""
dadosEnderecoCobranca = ""
filialAntiga=""

FilialDoRegistro=Request("paramFilial")
FilialFiltro=Request("Filial")
if FilialFiltro="" then
   FilialFiltro=SESSION("FILIAL") 
end if
TipoSolicitacao = Request("tiposol")
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <title>SWnet - Cadastro Unificado</title>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1"/>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
  <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="Strict" />
  <link rel="stylesheet" type="text/css" href="../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.css" />
  <link rel="stylesheet" type="text/css" href="../../dhtmlx/dhtmlxCalendar/codebase/skins/dhtmlxcalendar_dhx_web.css" />
  <link rel="stylesheet" type="text/css" href="includes/css/Unificado.css" />
  <script src="../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.js"></script>
  <script src="includes/js/index.js?v=<%=Now()%>"></script>
  <script src="includes/js/Unificado.js?v=<%=Now()%>"></script>
  <script src="includes/js/jquery-3.4.1.min.js"></script>
  <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.5.2/dist/js/bootstrap.bundle.min.js"></script>
  <script>
      $(function () {
        $('[data-toggle="tooltip"]').tooltip()
      })
  </script>
  <style>
      body {
        <%if Request("act") <> "" And Request("act") <> "LST" And Request("act") <> "NEW" then%>
          background-color:gray;
        <%end if%>
      }

      window.onload = function() {
        var selectFilial = document.getElementById('Filial');
        selectFilial.dispatchEvent(new Event('change'));
      };
    </style>
</head>
<body onload="doOnLoad('<%=FilialFiltro%>')" <%if Request("act") = "LST" then%> style="background-color: lightgray;"<%end if%>>

<% 
call ValidaPermissaoFormularios("../Unificado/index",SESSION("CODUSER"),SESSION("FILIAL"))
SELECT CASE Request("act")
       CASE "INC" : ChamarINC()
       CASE "GRI" : ChamarGRI()
       CASE "ALT" : ChamarALT()
       CASE "GRA" : ChamarGRA()
       CASE "EXC" : ChamarEXC() 'Exibe os dados para exclusão
       CASE "GRE" : ChamarGRE() 'Executa a exclusão
       CASE "NEW" : NewCAD()
       CASE "VIS" : ChamarVIS(Request("act"))
       CASE Else : VerificaSessaoExpirada() : ListUNI()
END SELECT

' -------------------------------------------------------------------------------------------------
Function UpdateCAD()
  if Request("Filial") = "00" Or Request("Filial") = "04" then
    Empresa = "BR02"
  Elseif Request("Filial") = "05" then
    Empresa = "BR03"
  Elseif Request("Filial") = "06" then
   Empresa = "BR04"
  Elseif Request("Filial") = "07" then
    Empresa = "BR07"
  end if

  SQL = " UPDATE UnificadoCliente SET "
  SQL = SQL &   "GrupoContas = '" & Request("GrupoContas") & "', "
  SQL = SQL &   "Empresa = '" & Empresa & "', "
  SQL = SQL &   "Filial = '" & Request("Filial") & "', "
  SQL = SQL &   "RazaoSocial = '" & removerCaracteresEspeciais(Request("RazaoSocial")) & "', "
  SQL = SQL &   "NomeFantasia = '" & removerCaracteresEspeciais(Request("NomeFantasia")) & "', "
  SQL = SQL &   "Rua = '" & removerCaracteresEspeciais(Request("Rua")) & "', "
  SQL = SQL &   "Numero = '" & Request("Numero") & "', "
  SQL = SQL &   "Bairro = '" & removerCaracteresEspeciais(Request("Bairro")) & "', "
  SQL = SQL &   "Cep = '" & Trim(Replace(Replace(Request("Cep"),".",""),"-","")) & "', "
  SQL = SQL &   "Cidade = '" & removerCaracteresEspeciais(Request("Cidade")) & "', "
  SQL = SQL &   "Estado = '" & removerCaracteresEspeciais(Request("Estado")) & "', "
  SQL = SQL &   "Pais = '" & removerCaracteresEspeciais(Request("Pais")) & "', "
  SQL = SQL &   "Telefone = '" & Trim(Replace(Replace(Replace(Request("Telefone"),"(",""),")",""),"-","")) & "', "
  SQL = SQL &   "Fax = '', "
  SQL = SQL &   "EmailXml = '" & Request("EmailXml") & "', "
  SQL = SQL &   "EmailPagamento = '" & Request("EmailPagamento") & "', "
  SQL = SQL &   "CodigoVendedor = '', "
  SQL = SQL &   "CodigoHfm = '" & Request("CodigoHfm") & "', "
  SQL = SQL &   "Cpf = '" & Trim(Replace(Replace(Request("Cpf"),".",""),"-","")) & "', "
  SQL = SQL &   "Cnpj = '" & Trim(Replace(Replace(Replace(Request("Cnpj"),"/",""),".",""),"-","")) & "', "
  SQL = SQL &   "InscricaoEstadual = '" & Request("InscricaoEstadual") & "', "
  SQL = SQL &   "InscricaoMunicipal = '" & Request("InscricaoMunicipal") & "', "
  SQL = SQL &   "Declarante = '" & Request("Declarante") & "', "
  SQL = SQL &   "Contribuinte = '" & Request("Contribuinte") & "', "
  SQL = SQL &   "NaturezaLegal = '', "
  SQL = SQL &   "TipoIndustria = '" & Request("TipoIndustria") & "', "
  SQL = SQL &   "CodigoSuframa = '', "
  SQL = SQL &   "ContaConciliacao = '" & Request("ContaConciliacao") & "', "
  SQL = SQL &   "ContaBanco = '', "
  SQL = SQL &   "ChaveOrdenacao = '" & Request("ChaveOrdenacao") & "', "
  SQL = SQL &   "CondicaoPagamento = '" & Request("CondicaoPagamento") & "', "
  SQL = SQL &   "MetodoPagamento = '" & Request("MetodoPagamento") & "', "
  SQL = SQL &   "CodigoJuros = '" & Request("CodigoJuros") & "', "
  SQL = SQL &   "CodigoAntigo = '', "
  SQL = SQL &   "Correspondencia = '" & Request("Correspondencia") & "', "
  SQL = SQL &   "LimiteCredito = '" & Request("LimiteCredito") & "', "
  SQL = SQL &   "VolumeVendam2 = '" & Request("VolumeVendam2") & "', "
  SQL = SQL &   "PrecoEstm2 = '" & Request("PrecoEstm2") & "', "
  SQL = SQL &   "PrecoEstTon = '" & Request("PrecoEstTon") & "', "
  SQL = SQL &   "LimiteCredSug = '" & Request("LimiteCredSug") & "', "
  SQL = SQL &   "PrazoPag = '" & Request("PrazoPag") & "', "
  SQL = SQL &   "DtVenda = '" & invertedata(Request("DtVenda")) & "', "
  SQL = SQL &   "Seguimento = '" & Request("Seguimento") & "', "
  SQL = SQL &   "ResponsavelAprovacao = '" & removerCaracteresEspeciais(Request("ResponsavelAprovacao")) & "', "
  SQL = SQL &   "Natureza = '" & Request("Natureza") & "', "
  SQL = SQL &   "Tipo = '" & Request("Tipo") & "', "
  SQL = SQL &   "TipoFrete = '" & Request("TipoFrete") & "', "
  SQL = SQL &   "CodigoVendedorPro = '" & Request("CodigoVendedorProtheus") & "', "
  SQL = SQL &   "CodigoClasProd = '" & Request("CodigoClasProd") & "', "
  SQL = SQL &   "DiasViagem = '" & Request("DiasViagem") & "', "
  SQL = SQL &   "DiasEntregaMaior = " & iif(Request("DiasEntregaMaior")="","null",Request("DiasEntregaMaior")) & ", "
  SQL = SQL &   "DiasEntregaMenor = " & iif(Request("DiasEntregaMenor")="","null",Request("DiasEntregaMenor")) & ", "
  SQL = SQL &   "RuaEntPro = '" & removerCaracteresEspeciais(Request("RuaEntPro")) & "', "
  SQL = SQL &   "NumeroEntPro = '" & Request("NumeroEntPro") & "', "
  SQL = SQL &   "BairroEntPro = '" & removerCaracteresEspeciais(Request("BairroEntPro")) & "', "
  SQL = SQL &   "CepEntPro = '" & Trim(Replace(Replace(Request("CepEntPro"),".",""),"-","")) & "', "
  SQL = SQL &   "CidadeEntPro = '" & removerCaracteresEspeciais(Request("CidadeEntPro")) & "', "
  SQL = SQL &   "EstadoEntPro = '" & removerCaracteresEspeciais(Request("EstadoEntPro")) & "', "
  SQL = SQL &   "RuaCobPro = '" & removerCaracteresEspeciais(Request("RuaCobPro")) & "', "
  SQL = SQL &   "NumeroCobPro = '" & Request("NumeroCobPro") & "', "
  SQL = SQL &   "BairroCobPro = '" & removerCaracteresEspeciais(Request("BairroCobPro")) & "', "
  SQL = SQL &   "CepCobPro = '" & Trim(Replace(Replace(Request("CepCobPro"),".",""),"-","")) & "', "
  SQL = SQL &   "CidadeCobPro = '" & removerCaracteresEspeciais(Request("CidadeCobPro")) & "', "
  SQL = SQL &   "EstadoCobPro = '" & removerCaracteresEspeciais(Request("EstadoCobPro")) & "', "
  SQL = SQL &   "Email = '" & Request("Email") & "', "
  SQL = SQL &   "PrecoPadrao = '" & Request("PrecoPadrao") & "', "
  SQL = SQL &   "TransacaoPedidos = '" & removerCaracteresEspeciais(Request("TransacaoPedidos")) & "', "
  SQL = SQL &   "AtividadePincipal = '" & removerCaracteresEspeciais(Request("AtividadePincipal")) & "', "
  SQL = SQL &   "GrupoEmpresa = '" & removerCaracteresEspeciais(Request("GrupoEmpresa")) & "', "
  SQL = SQL &   "AreaVendas = '" & removerCaracteresEspeciais(Request("AreaVendas")) & "', "
  SQL = SQL &   "NomeRepresentante = '" & removerCaracteresEspeciais(Request("NomeRepresentante")) & "', "
  SQL = SQL &   "EmailRepresentante = '" & Request("EmailRepresentante") & "', "
  SQL = SQL &   "TelefoneRep = '" & Trim(Replace(Replace(Replace(Request("TelefoneRep"),"(",""),")",""),"-","")) & "', "
  SQL = SQL &   "SegmentacaoMercado = '" & removerCaracteresEspeciais(Request("SegmentacaoMercado")) & "', "
  SQL = SQL &   "AvisaEmbarque = '" & removerCaracteresEspeciais(Request("AvisaEmbarque")) & "', "
  SQL = SQL &   "ContEmbalagens = '" & Request("ContEmbalagens") & "', "
  SQL = SQL &   "RegraEmbarque = '" & Request("RegraEmbarque") & "', "
  SQL = SQL &   "DiasTransporte = '" & Request("DiasTransporte") & "', "
  SQL = SQL &   "TolEntregaMenos = '" & Request("TolEntregaMenos") & "', "
  SQL = SQL &   "TolEntregaMais = '" & Request("TolEntregaMais") & "', "
  SQL = SQL &   "JanTolMenos = '" & Request("JanTolMenos") & "', "
  SQL = SQL &   "JanTolMais = '" & Request("JanTolMais") & "', "
  SQL = SQL &   "EmiteCertificado = '" & Request("EmiteCertificado") & "', "
  SQL = SQL &   "Frete = '" & Request("Frete") & "', "
  SQL = SQL &   "DesprezarTarifa = '" & Request("DesprezarTarifa") & "', "
  SQL = SQL &   "DistRomaneio = '" & Request("DistRomaneio") & "', "
  SQL = SQL &   "ObsEmissor = '" & removerCaracteresEspeciais(Request("ObsEmissor")) & "', "
  SQL = SQL &   "ObsCedag = '" & removerCaracteresEspeciais(Request("ObsCedag")) & "', "
  SQL = SQL &   "CodSap = '" & Request("CodSap") & "', "
  SQL = SQL &   "CodPro = '" & Request("CodPro") & "', "
  SQL = SQL &   "Tes = '" & Request("Tes") & "', "
  SQL = SQL &   "Eshop = '" & Request("eshop") & "', "
  SQL = SQL &   "TQtdeMaior = " & iif(Request("TQtdeMaior")="","0",Request("TQtdeMaior")) & ", "
  SQL = SQL &   "TQtdeMenor = " & iif(Request("TQtdeMenor")="","0",Request("TQtdeMenor")) & ", "
  SQL = SQL &   "PreCadastro = 'C',"
  SQL = SQL &   "CodLoja = '" & Request("CodLoja") & "', "
  SQL = SQL &   "KmEndereco = '" & Request("KmEndereco") & "', "
  SQL = SQL &   "KmEnderecoCobranca = '" & Request("KmEnderecoCobranca") & "', "
  SQL = SQL &   "isencaoIPI = " & iif(Request("isencaoIPI")="","null",Request("isencaoIPI")) & ", " 
  SQL = SQL &   "IcmsDiferido = '" & iif(Request("IcmsDiferido")="","null",Request("IcmsDiferido")) & "', " 
  SQL = SQL &   "percentualDiferimento = " & iif(Request("percentualDiferimento")="","null",Request("percentualDiferimento")) & " " 
  SQL = SQL &   "WHERE Codigo = '" & Request("cod") & "'; "
  UpdateCAD = SQL
End Function

' -------------------------------------------------------------------------------------------------
function removerCaracteresEspeciais(txt)
  Dim objRegExp
  Set objRegExp = New Regexp
  
  objRegExp.IgnoreCase = True
  objRegExp.Global = True
  
  if Len(txt) <> 0 then
    objRegExp.Pattern = "\s+"
    txt = objRegExp.Replace(txt, " ")
  
    objRegExp.Pattern = "[(?*"",\\<>&#~%{}+.@:\/!;']+"
    txt = objRegExp.Replace(txt, " ")
  end if

  removerCaracteresEspeciais = Trim(txt)
end Function
' -------------------------------------------------------------------------------------------------
Function ValidaFT1() 'Efetua a validação dos dados
	ValidaFT1 = ""
End Function
' -------------------------------------------------------------------------------------------------
Function fnumber(v)
  if v = "" then
    fnumber = ""
  Else
    fnumber = Replace(Replace(FormatNumber(v,2),".",""),",",".")
  end if
End Function
' -------------------------------------------------------------------------------------------------
Function ListUNI()  
  %>
  <style>
    @media screen and (max-width: 500px) {
      .vscr {
        display:none;
      }
    }
    label {
      font-size: 12px;
    }
    a {
      cursor: pointer;
    }
  </style>
  <div class="jumbotron text-center">
    <h3 style="color:gray">CADASTRO UNIFICADO</h3>
  </div>
  <div class="container">
    <div style="box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.1), 0 6px 10px 0 rgba(0, 0, 0, 0.19);">
      <div class="container" style="background-color: rgb(242, 242, 242);">
        <form action="index.asp" id="formfil">
          <input type="hidden" name="envia" value="s"/> 
            <div class="row">
              <div class="col-sm-3">
                <label>Cliente</label>
                <input type="text" class="form-control" name="cliente" id="cliente" value="<%=Trim(Request("cliente"))%>" placeholder="Informe o cliente..."/>
              </div>
              <div class="col-sm-3">
                <label>Código</label>
                <input type="text" class="form-control" name="codigo" id="codigo" value="<%=Trim(Request("codigo"))%>" placeholder="Informe o código..."/>
              </div>
              <div class="col-sm-3">
                   <label>Status</label>
                  <SELECT name="Statusfil" id="Statusfil" class="form-select">
                    <option value="">TODOS</option>
                    <option value="P" <%if Request("Statusfil") = "P" then%>SELECTed<%end if%>>PENDENTE</option>
                    <option value="S" <%if Request("Statusfil") = "S" then%>SELECTed<%end if%>>FISCAL</option>
                    <option value="C" <%if Request("Statusfil") = "C" Or (Request("envia") = "" And SESSION("Controladoria")) then%>SELECTed<%end if%>>CONTROLADORIA</option>
                    <option value="E" <%if Request("Statusfil") = "E" Or (Request("envia") = "" And SESSION("E2OPEN")) then%>SELECTed<%end if%>>E2OPEN/CDAG</option>
                    <option value="I" <%if Request("Statusfil") = "I" Or (Request("envia") = "" And SESSION("ControlesInternos")) then%>SELECTed<%end if%>>CONTROLES INTERNOS</option>
                    <option value="F" <%if SESSION("NivelFinanceiro")=1 then%>SELECTed<%end if%>>FINANCEIRO/FINALIZADO</option>
                    <option value="R" <%if Request("Statusfil") = "R" then%>SELECTed<%end if%>>DEVOLVIDO</option>
                    <option value="X" <%if Request("Statusfil") = "X" then%>SELECTed<%end if%>>CANCELADO</option>
                    <option value="Z" <%if Request("Statusfil") = "Z" then%>SELECTed<%end if%>>CLIENTE PROSPECT</option>
                  </SELECT>
              </div>
              <div class="col-sm-3">
                <label>&nbsp;</label>
              </div>
            </div>  
            <div class="row">
              <div class="col-sm-3">
                <label>Data Início</label>
                <input type="text" readonly class="form-control" name="dtini" id="dtini" style="cursor:pointer" value="<%=Request("dtini")%>" />
              </div>
              <div class="col-sm-3">
                <label>Data Fim</label>
                <input type="text" readonly class="form-control" name="dtfim" id="dtfim" style="cursor:pointer" value="<%=Request("dtfim")%>" />
              </div> 
              <div class="col-sm-3">
                <label>Planta</label>
                <SELECT name="planta" id="planta" class="form-select" onchange="VendedorPrincipal(this.value,'<%=CodigoVendedor%>')">>
                  <option value="">TODAS</option>
                  <option value="00" <%if FilialFiltro = "00" then%>SELECTED<%end if%>>BR02 - PIRAPETINGA</option>
                  <option value="04" <%if FilialFiltro = "04" then%>SELECTED<%end if%>>BR02 - UBERABA</option>
                  <option value="05" <%if FilialFiltro = "05" then%>SELECTED<%end if%>>BR03 - BENTO GONÇALVES</option>
                  <option value="06" <%if FilialFiltro = "06" then%>SELECTED<%end if%>>BR04 - MARANGUAPE</option>
                  <option value="07" <%if FilialFiltro = "07" then%>SELECTED<%end if%>>BR07 - SAQUAREMA</option>
                </SELECT>
              </div>
              <div class="col-sm-3">
                <label>&nbsp;</label>
                <div class="input-group-prepend">
                  <button type="submit" class="btn btn-info" title="Filtrar"><i class="fas fa-filter"></i></button>&nbsp;
                  <button type="button" class="btn btn-warning" title="Adicionar novo" onclick="location.href='index.asp?act=NEW'"><i class="fas fa-plus"></i></button>&nbsp;
                  <button type="button" class="btn btn-info" title="Limpar os campos dos filtros" onclick="limparFiltro()"><i class="fas fa-sync-alt"></i></button>&nbsp;
                  <a href="manual.pdf" target="_blank" title="Manual" class="btn btn-success" >
                    <i class="fas fa-book-reader"></i>
                  </a>&nbsp;
                  <a href="Procedimento.pdf" target="_blank" title="Procedimento" class="btn btn-success" >
                    <i class="fa fa-address-book" aria-hidden="true"></i>
                  </a>&nbsp;
                </div>
              </div>
            </div>
            <div>&nbsp;</div>
        </form>
      </div>
      <div class="table-responsive" style="background-color:lightgray">
        <table class="table table-bordered table-hover" style="font-size:12px;background-color:white">
          <thead>
            <tr style="background-color:gray;color:white;font-weight:bold">
              <td class="text-left" style="width: 8%;">Código</td>
              <td class="text-center" colspan="2">Status</td>
              <td class="text-left">LeadTime</td>
              <td class="text-left">Emissão</td>
              <td class="text-left">Emitente</td>
              <td class="text-left vscr">Razão Social</td>
              <td class="text-left">Nome Fantasia</td>
              <td class="text-left">Planta</td>
              <td class="text-center">Ação</td>
            </tr>
          </thead>
          <tbody>
            <%
            
            SQL = "SELECT top 100 Case when a.PreCadastro='P' then 'P'+right('000000'+a.Codigo,6) else rtrim(ltrim(a.Codigo)) end as Codigo, "
            SQL = SQL & "a.Codigo as CodigoOriginal, a.DataEmissao, a.RazaoSocial, a.NomeFantasia, a.StatusCli, a.QuemIncluiu, a.CodigoAntigo, a.Filial, a.CodLoja, a.VolumeVendam2, a.PrecoEstm2, a.PrecoEstTon, "
            SQL = SQL & "a.Eshop, isnull(a.PreCadastro,'') as PreCadastro, a.LimiteCredito, a.CodPro, t2.NivelFinanceiro, t2.NivelFiscal, t2.CRM_NOME, t2.CRM_LOGIN, CASE WHEN a.PreCadastro='V' THEN 1 ELSE 0 END FlagNaoLiberaFinanceiro "
            SQL = SQL & ",(SELECT DATEDIFF(DAY,MIN(DataAlt), MAX(DataAlt)) FROM UnificadoHistoricoCliente b WHERE b.CodigoCad = a.Codigo) LeadTime "
            SQL = SQL & "FROM UnificadoCliente a "
            SQL = SQL & "left join HELPDESK..CRM_DTI t2 ON t2.CRM_LOGIN = a.QuemIncluiu "
            SQL = SQL & "left join SOLCREDITOCLIENTE t3 ON t3.COD = a.CodPro "
            SQL = SQL & "WHERE Del = 0 "
            if Nivel = 1 then
               SQL = SQL & "AND QuemIncluiuCod = '" & SESSION("CODIGO") & "' "
            end if

            if Request("envia") = "" then
              if SESSION("CEDAG") And SESSION("E2OPEN") then
                SQL = SQL & "AND (StatusCli = 'D' OR StatusCli = 'A' Or StatusCli = 'E') "
              Elseif SESSION("CEDAG") then
                SQL = SQL & "AND (StatusCli = 'D' OR StatusCli = 'A') "
              Elseif SESSION("E2OPEN") then
                SQL = SQL & "AND StatusCli = 'E' "
              Elseif SESSION("ControlesInternos") then
                SQL = SQL & "AND StatusCli = 'I' "
              Elseif SESSION("Controladoria") then
                SQL = SQL & "AND StatusCli = 'C' "
              Elseif SESSION("NivelFinanceiro") = 1 then
               SQL = SQL & "AND StatusCli = 'F' "
              Elseif SESSION("NivelFiscal") = 1 then
               SQL = SQL & "AND StatusCli = 'S' "               
              else
                 SQL = SQL & "AND StatusCli NOT IN ('F','X') "  
              end if
            Else
              if Request("cliente") <> "" then
                SQL = SQL & "AND RazaoSocial LIKE '%" & Request("cliente") & "%' OR NomeFantasia LIKE '%" & Request("cliente") & "%' "
              end if
              if Request("codigo") <> "" then
                SQL = SQL & "AND Codigo = '" & Request("codigo") & "' "
              end if
              if Request("dtini") <> "" then
                SQL = SQL & "AND DataEmissao >= '" & invertedata(Request("dtini")) & "' "
              end if
              if Request("dtfim") <> "" then
                SQL = SQL & "AND DataEmissao <= '" & invertedata(Request("dtfim")) & "' "
              end if
              if Request("Statusfil") <> "" and Request("Statusfil") <> "Z" then
                SQL = SQL & "AND StatusCli = '" & Request("Statusfil") & "' "
              end if
              if Request("planta") <> "" then
                SQL = SQL & "AND Filial = '" & Request("planta") & "' "
              end if              
            end if

            IF Request("Statusfil") = "Z" then
               SQL = SQL & "AND PreCadastro = 'P' "
            end if
            SQL = SQL & " ORDER BY Codigo Desc"
            Dados11.Open(SQL)

            Dim jaTemVisualizar
            dim temNivelControladoria
            temNivelControladoria=false
            if not Dados11.EOF then
              while not Dados11.EOF
                temNivelFinanceiro = False
                jaTemVisualizar=false
                Emissor = True
                'Status do cadastro
                SELECT CASE Dados11("StatusCli")
                  CASE "P"
                    StatusCliente = "PENDENTE"
                    NovoStatus = "S"
                    titulo = "Enviar para o Fiscal"
                    statusbg = ""
                  CASE "S"
                    StatusCliente = "FISCAL"
                    NovoStatus = "C"
                    titulo = "Enviar para Controladoria"
                    statusbg = ""                    
                  CASE "C"
                    StatusCliente = "CONTROLADORIA"
                    NovoStatus = "E"
                    titulo = "Enviar para E2OPEN/CDAG"
                    statusbg = ""                  
                 CASE "E"
                    StatusCliente = "E2OPEN/CDAG"
                    NovoStatus = "F"
                    titulo = "Finalizar cadastro"
                    statusbg = "alert-info"                    
                 CASE "I"
                    StatusCliente = "CONTROLES INTERNOS"
                    NovoStatus = "E"
                    titulo = "Enviar para E2OPEN/CDAG"
                    statusbg = ""
                  CASE "A"
                    StatusCliente = "APROVADO"
                    NovoStatus = "F"
                    titulo = "Finalizar cadastro"
                    statusbg = "alert-info"
                  CASE "F"
                    StatusCliente = "FINANCEIRO/FINALIZADO"
                    NovoStatus = ""
                    titulo = ""
                    statusbg = "alert-success"
                  CASE "R"
                    StatusCliente = "DEVOLVIDO"
                    NovoStatus = "S"
                    titulo = "Reenviar para FISCAL"
                    statusbg = "alert-danger"
                  CASE "X"
                    StatusCliente = "CANCELADO"
                    NovoStatus = ""
                    titulo = ""
                    statusbg = "alert-danger"
                  CASE Else
                    StatusCliente = ""
                    NovoStatus = ""
                    titulo = ""
                END SELECT

                'Libera ou não os botões de Visto, Devolução e Cancelamento conforme status e perfil
                if Emissor and (Dados11("StatusCli") = "R" Or Dados11("StatusCli") = "P") then
                   visto = True
                   cancela = False
                   devolve = False
                   if SESSION("NivelFiscal")  then
                      temNivelFiscal = True
                   end if   
                 Elseif (SESSION("CEDAG") OR SESSION("E2OPEN")) And Dados11("StatusCli") = "E" then
                        visto = True
                        cancela = True
                        devolve = True
                Elseif SESSION("ControlesInternos") And Dados11("StatusCli") = "I" then
                       visto = True
                       cancela = False
                       devolve = True
                Elseif SESSION("Controladoria") And Dados11("StatusCli") = "C" then
                      visto = True
                      cancela = True
                      devolve = False
                      temNivelControladoria = True
                Elseif SESSION("NivelFinanceiro") then
                       visto = True
                      cancela = False
                      devolve = False
                      temNivelFinanceiro = True
                Elseif SESSION("NivelFiscal") then
                       visto = True
                       cancela = False
                       devolve = False
                       temNivelFiscal = True
                Else
                    visto = False
                    cancela = False
                    devolve = False
                    temNivelFiscal = False
                    temNivelFinanceiro = False
                    temNivelControladoria = False
                end if

                if Dados11("StatusCli") = "F" then
                  finaliza = True
                end if                

                if Dados11("Filial") = "00" then
                   plantaDescricao = "Pirapetinga"
                Elseif Dados11("Filial") = "04" then
                       plantaDescricao = "Uberaba"
                Elseif Dados11("Filial") = "05" then
                      plantaDescricao = "Bento"
                Elseif Dados11("Filial") = "06" then
                      plantaDescricao = "Maranguape"
                Elseif Dados11("Filial") = "07" then
                      plantaDescricao = "Saquarema"
                end if 
                
                %>
                <tr class="<%=statusbg%>">
                  <td class="text-center"><%=Dados11("Codigo")%></td>
                  <td style="width: 1%;">
                    <% if Dados11("StatusCli") <> "P" then %>
                      <i class='far fa-clock' title="Histórico de vistos" onclick="window.open('historicovistos.asp?cod=<%=Dados11("Codigo")%>')" style='color: blue; cursor: pointer;'></i> &nbsp;
                    <% Else %>
                      <i class='far fa-clock' style='color: lightgray;'></i> &nbsp;
                    <% end if %>
                  </td>
                  <td><%=StatusCliente%></td>
                  <td style="text-align: center;"><%=Dados11("LeadTime")%></td>
                  <td><%=AcertaData(Dados11("DataEmissao"))%></td>
                  <td><%=Dados11("CRM_NOME")%></td>
                  <td><%=Dados11("RazaoSocial")%></td>
                  <td><%=Dados11("NomeFantasia")%></td>
                  <td><%=plantaDescricao%></td>
                  <td class="text-nowrap">
                          <button type="button" class="btn btn-info btn-sm" title="Upload De Arquivo" onclick="enviaArquivoParaServidor('<%=Dados11("CodigoOriginal")%>','<%=Dados11("Filial")%>','<%=Dados11("RazaoSocial")%>')"><i class="fas fa-upload"></i></button>

                      <!-- Usuario nivel fiscal -->  
                      <%if Emissor and not temNivelFiscal and Dados11("PreCadastro") <> "P" AND Dados11("StatusCli") = "P" AND (SESSION("USR") = Dados11("QuemIncluiu"))  then
                            if not SESSION("Controladoria") and not (SESSION("CEDAG") OR SESSION("E2OPEN")) and not SESSION("ControlesInternos") and SESSION("NivelFinanceiro")<>1 then%>
                              <button type="button" class="btn btn-success btn-sm" title="<%=titulo%>" onclick="EnviarParaOFiscal('<%=Dados11("CodigoOriginal")%>')"><i class="fas fa-check"></i></button>
                              <button type="button" class="btn btn-light btn-sm" title="Alterar" onclick="location.href='index.asp?act=ALT&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>&fz=<%=finaliza%>'"><i class="fas fa-pencil-alt"></i></button>
                              <button type="button" class="btn btn-danger btn-sm" title="Excluir" onclick="if(confirm('Deseja realmente excluir o cadastro <%=Dados11("CodigoOriginal")%>?')) { excluirRegistro(<%=Dados11("CodigoOriginal")%>); }"><i class="far fa-trash-alt"></i></button>
                            <%end if%>   
                      <% Elseif Emissor and Dados11("PreCadastro") <> "P" AND Dados11("StatusCli") = "S" AND SESSION("NivelFiscal") then%>
                                <button type="button" class="btn btn-success btn-sm" title="<%=titulo%>" onclick="TelaNivelFiscal('<%=Dados11("CodigoOriginal")%>','<%=titulo%>','<%=NovoStatus%>','<%=Dados11("Filial")%>')" ><i class="fas fa-check"></i></button>
                                <button type="button" class="btn btn-warning btn-sm" title="Devolver para o requisitante" onclick="fvisto('<%=Dados11("CodigoOriginal")%>','Devolver para o requisitante','R','<%=Dados11("Filial")%>')"><i class="fas fa-undo-alt"></i></button>  
                      <% Else 
                            if Emissor and not temNivelFiscal and Dados11("PreCadastro") = "P" AND (SESSION("USR") = Dados11("QuemIncluiu")) then
                               if ((Dados11("StatusCli") = "P" OR Dados11("StatusCli") = "R")) then %>
                                  <button type="button" class="btn btn-danger btn-sm" title="Unificar" onclick="location.href='index.asp?act=ALT&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>&fz=<%=finaliza%>'"><i class="fas fa-eraser"></i></button>
                                  <button type="button" class="btn btn-danger btn-sm" title="Cancelar" onclick="fvisto('<%=Dados11("CodigoOriginal")%>','Cancelar','X','<%=Dados11("Filial")%>')"><i class="fas fa-ban"></i></button>
                                  <button type="button" class="btn btn-danger btn-sm" title="Excluir" onclick="if(confirm('Deseja realmente excluir o cadastro <%=Dados11("CodigoOriginal")%>?')) { excluirRegistro(<%=Dados11("CodigoOriginal")%>); }"><i class="far fa-trash-alt"></i></button>
                               <% else
                                     if not jaTemVisualizar then %>
                                        <button type="button" class="btn btn-light btn-sm" title="Visualizar 1" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                        <%jaTemVisualizar = True
                                     end if
                             end if
                      end if%>

                      <!-- Usuario nivel emissor -->  
                      <%if not temNivelFiscal AND Emissor AND Dados11("PreCadastro") <> "P" AND (SESSION("USR") = Dados11("QuemIncluiu")) then
                            if ((Dados11("StatusCli") = "P" OR Dados11("StatusCli") = "R")) AND not SESSION("Controladoria") and not (SESSION("CEDAG") OR SESSION("E2OPEN")) and not SESSION("ControlesInternos") and SESSION("NivelFinanceiro")<>1 then%> 
                               <button type="button" class="btn btn-success btn-sm" title="<%=titulo%>" onclick="fvisto('<%=Dados11("CodigoOriginal")%>','<%=titulo%>','<%=NovoStatus%>','<%=Dados11("Filial")%>')" ><i class="fas fa-check"></i></button>
                               <button type="button" class="btn btn-danger btn-sm" title="Cancelar" onclick="fvisto('<%=Dados11("CodigoOriginal")%>','Cancelar','X','<%=Dados11("Filial")%>')"><i class="fas fa-ban"></i></button>
                               <button type="button" class="btn btn-light btn-sm" title="Alterar" onclick="location.href='index.asp?act=ALT&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>&fz=<%=finaliza%>'"><i class="fas fa-pencil-alt"></i></button>
                           <%Elseif Emissor and Dados11("PreCadastro") = "P" AND (SESSION("USR") <> Dados11("QuemIncluiu")) then
                                    if (Dados11("StatusCli") = "P" OR Dados11("StatusCli") = "R") then
                                        if not jaTemVisualizar then %>
                                          <button type="button" class="btn btn-light btn-sm" title="Visualizar 0" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                          <%jaTemVisualizar = True
                                        end if
                                    end if
                            Elseif visto = False And finaliza = False then
                                   if not jaTemVisualizar then %>
                                      <button type="button" class="btn btn-light btn-sm" title="Visualizar 2" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                      <%jaTemVisualizar = True
                                   end if
                            Elseif Dados11("PreCadastro") = "P" then
                                   if not jaTemVisualizar then %>
                                      <button type="button" class="btn btn-light btn-sm" title="Visualizar 3" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>                            
                                      <%jaTemVisualizar = True
                                   end if
                            end if
                        end if %> 
 
                       <!-- Usuario nivel controladoria -->
                       <%if visto and not temNivelFiscal and Dados11("PreCadastro") <> "P" AND SESSION("Controladoria")  then
                              if Dados11("StatusCli") = "R" then
                                 if not jaTemVisualizar then %>
                                    <button type="button" class="btn btn-light btn-sm" title="Visualizar 4" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                    <%jaTemVisualizar = True
                                 end if
                              elseif Dados11("StatusCli") = "P" then 
                                     if not jaTemVisualizar then %>
                                        <button type="button" class="btn btn-light btn-sm" title="Visualizar 5" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                        <%jaTemVisualizar = True
                                     end if
                              elseif Dados11("StatusCli") = "E" then 
                                     if not jaTemVisualizar then %>
                                        <button type="button" class="btn btn-light btn-sm" title="Visualizar 6" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                        <%jaTemVisualizar = True
                                     end if
                              elseif Dados11("StatusCli") = "I" then
                                     if not jaTemVisualizar then %>
                                        <button type="button" class="btn btn-light btn-sm" title="Visualizar 7" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                        <%jaTemVisualizar = True
                                     end if                                       
                              elseif Dados11("StatusCli") = "C" then%>
                                     <button type="button" class="btn btn-success btn-sm" title="<%=titulo%>" onclick="fvisto('<%=Dados11("CodigoOriginal")%>','<%=titulo%>','<%=NovoStatus%>','<%=Dados11("Filial")%>')" ><i class="fas fa-check"></i></button>
                                     <button type="button" class="btn btn-warning btn-sm" title="Devolver para o requisitante" onclick="fvisto('<%=Dados11("CodigoOriginal")%>','Devolver para o requisitante','R','<%=Dados11("Filial")%>')"><i class="fas fa-undo-alt"></i></button>  
                                    <%if not jaTemVisualizar then %>
                                       <button type="button" class="btn btn-light btn-sm" title="Visualizar 8" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                       <%jaTemVisualizar = True
                                    end if                                      
                                    if cancela then %>
                                       <button type="button" class="btn btn-danger btn-sm" title="Cancelar" onclick="fvisto('<%=Dados11("CodigoOriginal")%>','Cancelar','X','<%=Dados11("Filial")%>')"><i class="fas fa-ban"></i></button>
                                    <% end if
                             end if
                         end if %>
 
                        <!-- Usuario nivel CDAG -->
                        <%if visto and Dados11("PreCadastro") <> "P" AND (SESSION("CEDAG") OR SESSION("E2OPEN")) AND NOT SESSION("NivelFiscal") then
                           if (Dados11("StatusCli") = "R" OR Dados11("StatusCli") = "P") then 
                                 if not jaTemVisualizar then %>
                                   <button type="button" class="btn btn-light btn-sm" title="Visualizar 9" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                   <%jaTemVisualizar = True
                                end if
                          elseif Dados11("StatusCli") = "E" then
                                if not jaTemVisualizar then %>
                                   <button type="button" class="btn btn-light btn-sm" title="Visualizar 10" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                   <%jaTemVisualizar = True
                                end if %>     
                                <button type="button" class="btn btn-warning btn-sm" title="Devolver para o requisitante" onclick="fvisto('<%=Dados11("CodigoOriginal")%>','Devolver para o requisitante','R','<%=Dados11("Filial")%>')"><i class="fas fa-undo-alt"></i></button>  
                                <button type="button" class="btn btn-secondary btn-sm" title="Não Conforme" onclick="fvisto('<%=Dados11("CodigoOriginal")%>','Enviar para Controles Internos','I','<%=Dados11("Filial")%>')"><i class="fas fa-thumbs-down"></i></button>
                                <button type="button" class="btn btn-success btn-sm" title="Alterar e Finalizar" onclick="location.href='index.asp?act=ALT&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>&fz=<%=finaliza%>'"><i class="fas fa-pencil-alt"></i></button>
                                <% if cancela then %>
                                  <button type="button" class="btn btn-danger btn-sm" title="Cancelar" onclick="fvisto('<%=Dados11("CodigoOriginal")%>','Cancelar','X','<%=Dados11("Filial")%>')"><i class="fas fa-ban"></i></button>
                                <% end if
                          elseif Dados11("StatusCli") = "I" then 
                                 if not jaTemVisualizar then %>
                                    <button type="button" class="btn btn-light btn-sm" title="Visualizar 11" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                    <%jaTemVisualizar = True
                                 end if                                      
                          elseif Dados11("StatusCli") = "C" then
                                 if not jaTemVisualizar then %>
                                    <button type="button" class="btn btn-light btn-sm" title="Visualizar 12" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                    <%jaTemVisualizar = True
                                end if                                    
                           end if
                       end if %>
 
                       <!-- Usuario nivel Controles Internos -->
                        <%if visto and Dados11("PreCadastro") <> "P" AND SESSION("ControlesInternos") then
                              if Dados11("StatusCli") = "R" then
                                 if not jaTemVisualizar then %>
                                    <button type="button" class="btn btn-light btn-sm" title="Visualizar 13" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                    <%jaTemVisualizar = True
                                end if        
                          elseif Dados11("StatusCli") = "P" then
                                 if not jaTemVisualizar then %>
                                    <button type="button" class="btn btn-light btn-sm" title="Visualizar 14" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                    <%jaTemVisualizar = True
                                 end if
                          elseif Dados11("StatusCli") = "E" then
                                 if not jaTemVisualizar then %>
                                    <button type="button" class="btn btn-light btn-sm" title="Visualizar 15" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                    <%jaTemVisualizar = True
                                 end if             
                          elseif Dados11("StatusCli") = "I" then
                                  if not jaTemVisualizar then %>
                                     <button type="button" class="btn btn-light btn-sm" title="Visualizar 16" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                     <%jaTemVisualizar = True
                                  end if%>
                                 <button type="button" class="btn btn-warning btn-sm" title="Devolver para o requisitante" onclick="fvisto('<%=Dados11("CodigoOriginal")%>','Devolver para o requisitante','R','<%=Dados11("Filial")%>')"><i class="fas fa-undo-alt"></i></button>
                                 <button type="button" class="btn btn-success btn-sm" title="<%=titulo%>" onclick="fvisto('<%=Dados11("CodigoOriginal")%>','<%=titulo%>','<%=NovoStatus%>','<%=Dados11("Filial")%>')" ><i class="fas fa-check"></i></button>
                          <% elseif Dados11("StatusCli") = "C" then
                                    if not jaTemVisualizar then %>                            
                                      <button type="button" class="btn btn-light btn-sm" title="Visualizar 17" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                      <%jaTemVisualizar = True
                                    end if
                              end if
                          end if %>
 
                        <!-- Usuario Nivel Financeiro -->
                        <%if visto and Dados11("PreCadastro") <> "P" AND temNivelFinanceiro then
                              if Dados11("StatusCli") = "R" then
                                 if not jaTemVisualizar then %>  
                                    <button type="button" class="btn btn-light btn-sm" title="Visualizar 18" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                    <%jaTemVisualizar = True
                                 end if
                          elseif Dados11("StatusCli") = "P" then
                                 if not jaTemVisualizar then %>  
                                    <button type="button" class="btn btn-light btn-sm" title="Visualizar 19" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                    <%jaTemVisualizar = True
                                 end if
                          elseif Dados11("StatusCli") = "E" then
                                 if not jaTemVisualizar then %>  
                                    <button type="button" class="btn btn-light btn-sm" title="Visualizar 20" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                    <%jaTemVisualizar = True
                                 end if
                          elseif Dados11("StatusCli") = "I" then
                                 if not jaTemVisualizar then %>  
                                    <button type="button" class="btn btn-light btn-sm" title="Visualizar 21" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                    <%jaTemVisualizar = True
                                 end if
                          elseif Dados11("StatusCli") = "C" then
                                 if not jaTemVisualizar then %>  
                                    <button type="button" class="btn btn-light btn-sm" title="Visualizar 22" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                                    <%jaTemVisualizar = True
                                 end if
                          elseif Dados11("StatusCli") = "F" AND Dados11("CodPro") <> "" AND Dados11("FlagNaoLiberaFinanceiro") = "0" then %>
                                 <button type="button" class="btn btn-danger btn-sm" title="Solicitar Limite de Credito" onclick="LimiteCreditoFinanceiro('<%=Dados11("CodPro")%>','<%=Dados11("RazaoSocial")%>','<%=Dados11("LimiteCredito")%>','<%=Dados11("Filial")%>','<%=Dados11("CRM_NOME")%>','<%=Dados11("CRM_LOGIN")%>','<%=Dados11("CodLoja")%>','<%=Dados11("VolumeVendam2")%>','<%=Dados11("PrecoEstm2")%>','<%=Dados11("PrecoEstTon")%>')"><i class="fas fa-dollar-sign"></i></button>
                          <% end if
                        end if
                      end if
 
                      if finaliza then
                         if not jaTemVisualizar then %>  
                            <button type="button" class="btn btn-light btn-sm" title="Visualizar 23" onclick="Finalizado('<%=Dados11("CodigoOriginal")%>','<%=visto%>','<%=devolve%>','<%=cancela%>','<%=NovoStatus%>','<%=Dados11("StatusCli")%>','<%=Emissor%>')"><i class="far fa-eye"></i></button>
                            <%jaTemVisualizar = True
                         end if
                      end if
                      if not jaTemVisualizar then %>
                         <button type="button" class="btn btn-light btn-sm" title="Visualizar 24" onclick="location.href='index.asp?act=ALT&save=false&cod=<%=Dados11("CodigoOriginal")%>&vs=<%=visto%>&dv=<%=devolve%>&ca=<%=cancela%>&ns=<%=NovoStatus%>&StatusCli=<%=Dados11("StatusCli")%>&Emissor=<%=Emissor%>'"><i class="far fa-eye"></i></button>
                         <%jaTemVisualizar = True
                      end if %>
                  </td>
                </tr>
                <%
                Dados11.MoveNext()
              wend
            end if
            Dados11.Close()
            %>
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <div class="modal fade" id="myModal3">
    <div class="modal-dialog">
      <div class="modal-content">
        <div class="modal-header" id="modalheader3">
          <h6 class="modal-title" id="modaltitle3">Visto Cadastro Unificado</h6>
          <button type="button" class="close" id="modalbtnx3" data-dismiss="modal">&times;</button>
        </div>
        <div class="modal-body" id="modalbody3">
          <p id="vper">Confirma prosseguir?</p>
          <form action="index.asp" method="POST" id="formVisto" name="formVisto">
            <input type="hidden" name="cod" id="cod">
            <input type="hidden" name="act" id="act" value="VIS">
            <input type="hidden" name="NovoStatus" id="NovoStatus">
            <input type="hidden" name="Filial" id="Filial">
            <div id="divcontrole" style="display: none;">
              <SELECT name="controle" id="controle" class="form-SELECT" onchange="document.getElementById('NovoStatus').value=this.value" style="background-color: lightyellow">
                <option value="" style="display: none;">Selecione...</option>
                <option value="I">CONTROLES INTERNOS</option>
                <option value="C">CONTROLADORIA</option>
              </SELECT><br><br>
            </div>
            <textarea name="vobs" id="vobs" rows="3" style="width:100%" maxlength="100" placeholder="Comentários"></textarea>
          </form>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-sm btn-danger" id="btnsubmitd" onclick="validavisto()" style="width:149px; display: none;">Devolver</button>
          <button type="button" class="btn btn-sm btn-success" id="btnsubmit" onclick="validavisto()" style="width:149px">Confirmar</button>
          <button type="button" class="btn btn-sm btn-light" data-dismiss="modal" style="width:149px">Cancelar</button>
        </div>
      </div>
    </div>
  </div>
  <%
End Function

Function NewCAD()
  %>
  <div class="jumbotron text-center">
    <h3 style="color:gray">CADASTRO UNIFICADO</h3>
  </div>
  <div class="container">    
    <script type="text/javascript">
        function seltipo() {        
            if ( (document.getElementById('tiposolcli').checked == true || 
                  document.getElementById('tiposolclip').checked == true ||
                  document.getElementById('tiposolfor').checked == true || 
                  document.getElementById('tiposolpr1').checked == true || 
                  document.getElementById('tiposolpr2').checked == true ) ) {
                formtipo.submit()
            } else {
                malert('Selecione o tipo de cadastro');
            }
        }
    </script>
    <div class="row justify-content-center">
      <div class="col-lg-5">
        <form action="index.asp" id="formtipo">
          <input type="hidden" name="act" value="INC" />
    
          <h5 style="color:gray">Selecione o tipo de cadastro</h5><br />
          <div class="custom-control custom-radio custom-control">&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="radio" class="custom-control-input" id="tiposolcli" name="tiposol" value="1"/>
            <label class="custom-control-label" for="tiposolcli">Cadastro de Clientes</label>
          </div><br />
          <div class="custom-control custom-radio custom-control">&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="radio" class="custom-control-input" id="tiposolclip" name="tiposol" value="5"/>
            <label class="custom-control-label" for="tiposolclip">Cadastro de Cliente Prospect</label>
          </div><br />          
          <br />
          <br />
          <div class="d-grid gap-2">
            <button type="button" class="btn btn-info" onclick="seltipo()" >A v a n ç a r</button>
            <button type="button" class="btn btn-secondary" onclick="location.href='index.asp?act=LST'">S o l i c i t a ç õ e s</button>
          </div>
        </form>
        <p>&nbsp;</p>
      </div>
    </div>    
  </div>
  <%
End Function

Function FormUNI(Acao, TxtAct, Codigo, GrupoContas, Empresa, RazaoSocial, NomeFantasia, Rua, Numero, Bairro, Cep, Cidade, Estado, Pais, Idioma, Telefone, Fax,EmailXml, EmailPagamento, CodigoHfm, Cpf, Cnpj, InscricaoEstadual, InscricaoMunicipal, Declarante, Contribuinte, NaturezaLegal, TipoIndustria,CodigoSuframa, ContaConciliacao, ContaBanco, ChaveOrdenacao, CondicaoPagamento, MetodoPagamento, CodigoJuros,  Correspondencia, LimiteCredito,VolumeVendam2, PrecoEstm2, PrecoEstTon, LimiteCredSug, PrazoPag, DtVenda, Seguimento, ResponsavelAprovacao, Natureza, Tipo, TipoFrete, CodigoVendedorPro,CodigoClasProd, DiasViagem, DiasEntregaMaior, DiasEntregaMenor, RuaEntPro, NumeroEntPro, BairroEntPro, CepEntPro, CidadeEntPro, EstadoEntPro, RuaCobPro, NumeroCobPro, BairroCobPro,CepCobPro, CidadeCobPro, EstadoCobPro, ContatoCompras, Departamento, Email, PrecoPadrao, TransacaoPedidos, AtividadePincipal, GrupoEmpresa, AreaVendas,NomeRepresentante, EmailRepresentante, TelefoneRep, SegmentacaoMercado, AvisaEmbarque, ContEmbalagens, RegraEmbarque, DiasTransporte, TolEntregaMenos,TolEntregaMais, JanTolMenos, JanTolMais, EmiteCertificado, Frete, DesprezarTarifa, DistRomaneio, ObsEmissor, ObsCedag, CodSap, CodPro, CodSia, Filial,Tes, eshop, dataEmissao,TQtdeMaior, TQtdeMenor, PreCadastro, CodLoja, KmEndereco, KmEnderecoCobranca, isencaoIPI, IcmsDiferido ,percentualDiferimento)
  altvis = False
  
  dadosEnderecoEntrega = "" 
  dadosEnderecoCobranca = ""
  FilialDoRegistro = Filial
  tituloDaPagina="CADASTRO UNIFICADO"

  if Acao = "GRA" then 
      dadosEnderecoEntrega = RuaEntPro + NumeroEntPro + BairroEntPro + CepEntPro + CidadeEntPro + EstadoEntPro
      dadosEnderecoCobranca = RuaCobPro + NumeroCobPro + BairroCobPro + CepCobPro + CidadeCobPro + EstadoCobPro
  elseif Acao="GRI" then
      if trim(TipoSolicitacao)="5" then  
        tituloDaPagina="PRÉ-CADASTRO"
      else
        tituloDaPagina="CADASTRO UNIFICADO"
      end if      
  end if  

  TipoSolicitacao=0
  if PreCadastro="P" and Acao<>"GRA" then
     TipoSolicitacao=5
     LimiteCredito=0
  end if
  
  if Request("save") <> "false" then
     Edita = True

    if Request("vs") = "True" And ( SESSION("CEDAG") Or SESSION("E2OPEN") Or SESSION("ControlesInternos") Or SESSION("Controladoria") ) then
       TxtAct = TxtAct & " e Vistar"
       if PreCadastro="P" then
          TxtAct = "Inserir"
       end if
       altvis = True
    Elseif Request("fz") = "True" And SESSION("CEDAG") then
          TxtAct = TxtAct & " e Finalizar"
          altvis = True
    end if
  Else
     Edita = False
  end if
  %>
  <style>
    .tit {
      font-size:16px;
      padding-top:40px;
      padding-bottom:40px;
      font-weight:bold;
      letter-spacing:1px;      
    }
    @media screen and (min-width: 601px) {
      .win {
        background-color:whitesmoke;
        box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.2), 0 6px 20px 0 rgba(0, 0, 0, 0.19);
        margin-bottom:20px;
      }
    }
    @media screen and (max-width: 600px) {
      .win {
        background-color:whitesmoke;
      }
    }
    .aceia {
      background-color:lavender;
      padding:13px 10px 10px 20px;
    }
    .aceib {
      background-color:lavender;
      padding:10px 10px 10px 10px;
    }
    .req {
      color:red;
      font-weight:bold;
    }
    .divmenu {
      text-align: center;
      letter-spacing: 1px;
      box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.1), 0 6px 10px 0 rgba(0, 0, 0, 0.19);
    }
  </style>
  <form action="index.asp" method="post" name="formInput" id="formInput">
  <input type="hidden" name="act" id="act" value="<%=Acao%>"/>
  <input type="hidden" name="cod" value="<%=Codigo%>"/>
  <input type="hidden" name="formulario" value="s"/>
  <input type="hidden" name="hidTipoSolicitacao" value="<%=TipoSolicitacao%>"/>
  <input type="hidden" name="hidFilialRegistro" value="<%=FilialDoRegistro%>"/>
 
  <div class="container">
    <div class="row justify-content-md-center">
      <div class="col col-lg-8 win">
        <div class="container-fluid text-center" style="height:80px;background-color:#ededed;display: flex;justify-content: center;align-items: center;">
          <h5 style="color:gray"><%=tituloDaPagina%></h5>
        </div>
        
        <ul class="nav nav-tabs nav-justified" role="tablist">
          <li class="nav-item" role="presentation">
            <button class="nav-link active" id="nmenu0" data-bs-toggle="tab" data-bs-target="#menu0" type="button" role="tab" aria-controls="nmenu0" aria-selected="true">Formulário Geral</button>
          </li>
          <li class="nav-item" role="presentation">
            <button class="nav-link" id="nmenu1" data-bs-toggle="tab" data-bs-target="#menu1" type="button" role="tab" aria-controls="nmenu1" aria-selected="false" onclick="CalcularDistanciaPelaAba();ValoresNaturezaPreDefinido('<%=FilialFiltro%>','<%=Natureza%>');VendedorParaProtheus('<%=FilialFiltro%>','<%=CodigoVendedorPro%>');AtivaEnderecos('<%=dadosEnderecoEntrega%>','<%=dadosEnderecoCobranca%>');">PROTHEUS</button>
          </li>
        </ul>
        <div class="container-fluid text-center" style="color:gray;margin-top:1em;letter-spacing:1px;font-size:12px;display:flex;justify-content: center;align-items: center;">
        </div>
        <!-- Tab panes -->
        <div class="tab-content">
          <!--===================================================
                            Formulário Geral
          ====================================================-->
          <div class="tab-pane fade show active" id="menu0" role="tabpanel" aria-labelledby="nmenu0">      
            <div class="tit">
              <div class="alert alert-primary divmenu">Formulário Geral</div>
            </div>
         
             <div class="form-group">
                <label>Eshop<span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
                <SELECT name="eshop" id="eshop" class="form-select" >
                   <option value="0" <%if eshop then%> SELECTed="SELECTed" <%end if %>> Não </option>
                   <option value="1" <%if eshop then%> SELECTed="SELECTed" <%end if %>> Sim </option>
                </SELECT>       
             </div>
               
            <div class="form-group">
              <label>Grupo de Contas <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <SELECT name="GrupoContas" id="GrupoContas" class="form-select" onchange="grpconta()">
                <option value=""></option>
                <option value="Z003" <%if GrupoContas = "Z003" then%>SELECTed<%end if%>>CLIENTES NACIONAIS</option>
                <option value="ZINT" <%if GrupoContas = "ZINT" then%>SELECTed<%end if%>>CLIENTES INTERCOMPANY</option>
              </SELECT>
            </div>

            <div class="form-group">
              <label>Empresa <span class="req">*</span></label>
              <SELECT name="Filial" id="Filial" class="form-select">
                <option value="00" <%if Filial = "00" then%>SELECTED<%end if%>>BR02 - PIRAPETINGA</option>
                <option value="04" <%if Filial = "04" then%>SELECTED<%end if%>>BR02 - UBERABA</option>
                <option value="05" <%if Filial = "05" then%>SELECTED<%end if%>>BR03 - BENTO GONÇALVES</option>
                <option value="06" <%if Filial = "06" then%>SELECTED<%end if%>>BR04 - MARANGUAPE</option>
                <option value="07" <%if Filial = "07" then%>SELECTED<%end if%>>BR07 - SAQUAREMA</option>
              </SELECT>
            </div>

            <div class="form-group">
              <label>Razão Social <span class="req">*</span></label>
              <input type="text"  class="form-control" name="RazaoSocial" id="RazaoSocial" value="<%=RazaoSocial%>" maxlength="50" />
            </div>
          
            <div class="form-group">
              <label>Nome Fantasia <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <input type="text"  class="form-control" name="NomeFantasia" id="NomeFantasia" value="<%=NomeFantasia%>" maxlength="20" />
            </div>

            <div class="form-group">
              <label>Rua <span class="req">*</span></label>
              <input type="text" onchange="endgeral(this.value)" class="form-control" name="Rua" id="Rua" value="<%=Rua%>" maxlength="50" />
            </div>

            <div class="form-group">
              <label>Número <span class="req">*</span></label>
              <input type="text" onchange="numgeral(this.value)" class="form-control" name="Numero" id="Numero" value="<%=Numero%>" maxlength="10" />
            </div>

            <div class="form-group">
              <label>Bairro <span class="req">*</label>
              <input type="text" onchange="bairrogeral(this.value)" class="form-control" name="Bairro" id="Bairro" value="<%=Bairro%>" maxlength="40" />
            </div>

            <div class="form-group">
              <label>CEP <span class="req">*</label>
              <input type="text" autocomplete="off" onchange="formatacep(this.value,this.id);cepgeral(this.value)" onkeypress="return onlynumber();" class="form-control" name="Cep" id="Cep" value="<%=formatcep(Cep)%>" maxlength="8" />
            </div>

            <div class="form-group">
              <label>Estado <span class="req">*</label>
              <SELECT onchange="ufgeral(this.value); municipio(this.value,'Cidade','cidgeral','dCidade')" name="Estado" id="Estado" class="form-select">
                <option value=""></option>
                <%
                SQL2 = "SELECT DISTINCT ESTADO FROM UnificadoCidades ORDER BY ESTADO "
                Dados12.Source = SQL2
                Dados12.Open()
                if not Dados12.EOF then
                  while not Dados12.EOF
                    %>
                    <option value="<%=Dados12("ESTADO")%>" <%if Estado = Dados12("ESTADO") then%>SELECTed<%end if%>><%=Dados12("ESTADO")%></option>
                    <%
                    Dados12.MoveNext()
                  wend
                end if
                Dados12.Close()
                %>
              </SELECT>
            </div>
    
            <div class="form-group" id="dCidade">
              <label>Cidade <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <input type="text"readonly class="form-control" name="Cidade" id="Cidade" value="<%=Cidade%>" readonly />
            </div>

            <div class="form-group">
              <label>País <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <input type="text" readonly  class="form-control" name="Pais" id="Pais" value="Brasil" maxlength="6" />
            </div>

            <div class="form-group">
              <label>KM até o destino <span class="req"><%if TipoSolicitacao<>5 then%>*<%else%><%end if%></span></label>
              <input type="text" class="form-control" name="KmEndereco" id="KmEndereco" value="<%=KmEndereco%>" maxlength="6" />
            </div>            

            <div class="form-group">
              <label>Idioma <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <input type="text" readonly  class="form-control" name="Idioma" id="Idioma" value="Português" maxlength="5" />
            </div>

            <div class="form-group">
              <label>Telefone <span class="req"></span></label>
              <input type="text" onchange="formatatel(this.value,this.id)" onkeypress="return onlynumber();" class="form-control" name="Telefone" id="Telefone" value="<%=Telefone%>" maxlength="11" />
            </div>

            <div class="form-group">
              <label>E-mail para o envio do XML <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <input type="text"  class="form-control" name="EmailXml" id="EmailXml" value="<%=EmailXml%>" maxlength="255" />
            </div>

            <div class="form-group">
              <label>E-mail para envio da ordem de pagamento <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <input type="text"  class="form-control" name="EmailPagamento" id="EmailPagamento" value="<%=EmailPagamento%>" maxlength="255" />
            </div>

            <div class="form-group">
              <label>Código HFM para clientes Intercompany</label>
              <input type="text" class="form-control" name="CodigoHfm" id="CodigoHfm" value="<%=CodigoHfm%>" maxlength="30" />
            </div>

            <div class="form-group">
              <label>CNPJ <span class="req">*</span></label>
              <input name="Cnpj" id="Cnpj" type="text" autocomplete="off" onchange="valida_cnpj(this.value,this.id)" class="form-control" value="<%=formatacnpj(Cnpj)%>" maxlength="19" />
            </div>

            <div class="form-group">
              <label>CPF <span class="req">*</span></label>
              <input name="Cpf" id="Cpf" type="text" autocomplete="off" onkeypress="return onlynumber();" onchange="valida_cpf(this.value,this.id)" class="form-control" value="<%=formatacpf(Cpf)%>" maxlength="14" />
            </div>

            <div class="form-group">
              <label>Inscrição Estadual  <span class="req">*</span></label>
              <input type="text" class="form-control" name="InscricaoEstadual" id="InscricaoEstadual" value="<%=InscricaoEstadual%>" maxlength="18" />
            </div>

            <div class="form-group">
              <label>Inscrição Municipal</label>
              <input type="text" class="form-control" name="InscricaoMunicipal" id="InscricaoMunicipal" value="<%=InscricaoMunicipal%>" maxlength="18" />
            </div>

            <div class="form-group">
              <label>Declarante de PIS/CONFINS</label>
              <SELECT name="Declarante" id="Declarante" class="form-select">
                <option value=""></option>
                <option value="CM" <%if Declarante = "CM" then%>SELECTed<%end if%>>Regime Cumulativo</option>
                <option value="NC" <%if Declarante = "NC" then%>SELECTed<%end if%>>Regime Não Cumulativo</option>
              </SELECT>
            </div>

            <div class="form-group">
              <label>Contribuinte de ICMS</label>
              <SELECT name="Contribuinte" id="Contribuinte" class="form-select">
                <option value=""></option>
                <option value="CI" <%if Contribuinte = "CI" then%>SELECTed<%end if%>>Contribuinte Isento</option>
                <option value="CO" <%if Contribuinte = "CO" then%>SELECTed<%end if%>>Contribuinte de ICMS</option>
                <option value="NA" <%if Contribuinte = "NA" then%>SELECTed<%end if%>>Pessoa Física</option>
                <option value="NC" <%if Contribuinte = "NC" then%>SELECTed<%end if%>>Não contribuinte</option>
                <option value="PR" <%if Contribuinte = "PR" then%>SELECTed<%end if%>>Produtor Rural</option>
              </SELECT>
            </div>

            <div class="form-group">
              <label>ICMS Diferido <%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <SELECT name="IcmsDiferido" id="IcmsDiferido" onchange="NotificarDocumentoICMS()" class="form-select">
                <option value=""></option>
                 <option value="0" <%if IcmsDiferido = "0" then%>SELECTED<%end if%>>Não</option>
                 <option value="1" <%if IcmsDiferido = "1" then%>SELECTED<%end if%>>Sim</option>
              </SELECT>
            </div>

             <div class="form-group">
              <label>% de Diferimento</label>
              <input type="text" autocomplete="off" onchange="this.value=fnum(this.value,2)" onfocus="this.select();" oninput="this.value=this.value.replace(/[^0-9.,]/g, '').replace(/(\..*)\./g, '$1');" class="form-control" name="percentualDiferimento" id="percentualDiferimento" value="<%=percentualDiferimento%>" maxlength="6">
               <label id="lblPercentualDiferimento" style="display:none; color:red;">Atenção: obrigatório inserção de documento comprobatório</label>
            </div>

            <div class="form-group">
              <label>Suspensão/Isenção do IPI <%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <SELECT name="isencaoIPI" id="isencaoIPI" onchange="NotificarCartaDeIsencao()" class="form-select">
                <option value=""></option>
                <option value="0" <%if isencaoIPI = "0" then%>SELECTED<%end if%>>Não</option>
                <option value="1" <%if isencaoIPI = "1" then%>SELECTED<%end if%>>Sim</option>
              </SELECT>
              <label id="lblIsencaoIPI" style="display:none; color:red;">Atenção: É obrigatório anexar a carta de isenção de IPI do cliente</label>
            </div>

            <div class="form-group">
              <label>Tipo de Indústria <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <SELECT name="TipoIndustria" id="TipoIndustria" class="form-select">
                <option value=""></option>
                <option value="CN" <%if TipoIndustria = "CN" then%>SELECTED<%end if%>>Construtora</option>
                <option value="CR" <%if TipoIndustria = "CR" then%>SELECTED<%end if%>>Comércio</option>
                <option value="IN" <%if TipoIndustria = "IN" then%>SELECTED<%end if%>>Indústria</option>
                <option value="PE" <%if TipoIndustria = "PE" then%>SELECTED<%end if%>>Preponderantemente Exportador</option>
                <option value="RA" <%if TipoIndustria = "RA" then%>SELECTED<%end if%>>Revendedor Atacadista</option>
                <option value="RV" <%if TipoIndustria = "RV" then%>SELECTED<%end if%>>Revendedor Varejista</option>
              </SELECT>
            </div>

            <div class="form-group">
              <label>Código Suframa <span class="req" id="reqSuframa"></span></label>
              <input type="text"  class="form-control" name="CodigoSuframa" id="CodigoSuframa" value="<%=CodigoSuframa%>" maxlength="30" />
            </div>

            <div class="form-group">
              <label>Conta Conciliação <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <SELECT name="ContaConciliacao" id="ContaConciliacao" class="form-select">
                <option value=""></option>
                <option value="41100000" <%if ContaConciliacao = "41100000" then%>SELECTed<%end if%>>41100000 - Clientes Nacionais</option>
                <option value="40110100" <%if ContaConciliacao = "40110100" then%>SELECTed<%end if%>>40110100 - Clientes Estrangeiros</option>
                <option value="45500000" <%if ContaConciliacao = "45500000" then%>SELECTed<%end if%>>45500000 - Clientes de Parlamentos Federais</option>
                <option value="42560000" <%if ContaConciliacao = "42560000" then%>SELECTed<%end if%>>42560000 - Funcionários SK</option>
                <option value="94000000" <%if ContaConciliacao = "94000000" then%>SELECTed<%end if%>>94000000 - Clientes grupo que são Divisões da empresa</option>
                <option value="40210000" <%if ContaConciliacao = "40210000" then%>SELECTed<%end if%>>40210000 - Clientes grupo nacionais</option>
              </SELECT>
            </div>

            <div class="form-group">
              <label>Chave de Ordenação <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <input type="text" readonly class="form-control" name="ChaveOrdenacao" id="ChaveOrdenacao" value="0009" maxlength="4" />
            </div>

            <div class="form-group">
              <label>Condição de Pagamento <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <SELECT name="CondicaoPagamento" id="CondicaoPagamento" class="form-select">
                <option value=""></option>
                <%
                  SQL2 = "SELECT Codigo, Descricao FROM UnificadoCondPag ORDER BY Codigo "
                  Dados12.Source = SQL2
                  Dados12.Open()
                  if not Dados12.EOF then
                    while not Dados12.EOF
                      %>
                      <option value="<%=Dados12("Codigo")%>" <%if CondicaoPagamento = Dados12("Codigo") then%>SELECTed<%end if%>>[<%=Dados12("Codigo")%>] <%=Dados12("Descricao")%></option>
                      <%
                      Dados12.MoveNext()
                    wend
                  end if
                  Dados12.Close()
                %>
              </SELECT>
            </div>

            <div class="form-group">
              <label>Método de Pagamento <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <SELECT name="MetodoPagamento" id="MetodoPagamento" class="form-select">
                <option value=""></option>
                <option value="D" <%if MetodoPagamento = "D" then%>SELECTed<%end if%>>Boleto</option>
                <option value="A" <%if MetodoPagamento = "A" then%>SELECTed<%end if%>>Ordem de Pagamento</option>
              </SELECT>
            </div>

            <div class="form-group">
              <label>Código de Juros <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <input type="text" readonly class="form-control" name="CodigoJuros" id="CodigoJuros" value="BR" maxlength="2" />
            </div>

            <div class="form-group">
              <label>Correspondência (Empresa) <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <SELECT name="Correspondencia" id="Correspondencia" class="form-select">
                <option value=""></option>
                <option value="BR02" <%if Correspondencia = "BR02" then%>SELECTed<%end if%>>BR02 - PIRAPETINGA / UBERABA</option>
                <option value="BR03" <%if Correspondencia = "BR03" then%>SELECTed<%end if%>>BR03 - BENTO GONÇALVES</option>
                <option value="BR04" <%if Correspondencia = "BR04" then%>SELECTed<%end if%>>BR04 - MARANGUAPE</option>
                <option value="BR07" <%if Correspondencia = "BR07" then%>SELECTed<%end if%>>BR07 - SAQUAREMA</option>
              </SELECT>
            </div>

            <div class="form-group">
              <label>Limite de Crédito <span class="req"><%if TipoSolicitacao<>5 then%>*<%else%><%end if%></span></label>
              <input type="text" autocomplete="off" onchange="this.value=fnum(this.value,2)" onfocus="this.select();" oninput="this.value=this.value.replace(/[^0-9.,]/g, '').replace(/(\..*)\./g, '$1');" class="form-control" name="LimiteCredito" id="LimiteCredito" value="<%=fnumber(LimiteCredito)%>" maxlength="15" />
            </div>

            <div class="form-group">
              <label>Volume de venda projetado mensal: em M² <span class="req"><%if TipoSolicitacao<>5 then%>*<%else%><%end if%></span></label>
              <input type="text" autocomplete="off" onchange="maiorzero(this.value,this.id)" onkeypress="return onlynumber();" class="form-control" name="VolumeVendam2" id="VolumeVendam2" value="<%=VolumeVendam2%>" maxlength="15" />
            </div>

            <div class="form-group">
              <label>Preço de venda net estimado de venda: R$/M² <span class="req"><%if TipoSolicitacao<>5 then%>*<%else%><%end if%></span></label>
              <input type="text" autocomplete="off" onchange="this.value=fnum(this.value,2);maiorzero(this.value,this.id)" onfocus="this.select();" oninput="this.value=this.value.replace(/[^0-9.,]/g, '').replace(/(\..*)\./g, '$1');" class="form-control" name="PrecoEstm2" id="PrecoEstm2" value="<%=fnumber(PrecoEstm2)%>" maxlength="15" />
            </div>

            <div class="form-group">
              <label>Preço de venda net estimado de venda: R$/KG <span class="req"><%if TipoSolicitacao<>5 then%>*<%else%><%end if%></span></label>
              <input type="text" autocomplete="off" onchange="this.value=fnum(this.value,2);maiorzero(this.value,this.id)" onfocus="this.select();" oninput="this.value=this.value.replace(/[^0-9.,]/g, '').replace(/(\..*)\./g, '$1');" class="form-control" name="PrecoEstTon" id="PrecoEstTon" value="<%=fnumber(PrecoEstTon)%>" maxlength="15" />
            </div>

            <div class="form-group">
              <label>Segmento <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <SELECT name="Seguimento" id="Seguimento" class="form-select">
                <option value=""></option>
                <%if dataEmissao <= "20210730" and dataEmissao <> "" then %>
                    <option value="1" <%if Seguimento = "1" then%>SELECTed<%end if%>>1.Produtos Alimentícios</option>
                    <option value="2" <%if Seguimento = "2" then%>SELECTed<%end if%>>2. Proteína (Abates e produtos de Carne)</option>
                    <option value="3" <%if Seguimento = "3" then%>SELECTed<%end if%>>3.Laticínios</option>
                    <option value="4" <%if Seguimento = "4" then%>SELECTed<%end if%>>4.Químicos e Derivados</option>
                    <option value="5" <%if Seguimento = "5" then%>SELECTed<%end if%>>5.Frutas, Legumes e Vegetais</option>
                    <option value="6" <%if Seguimento = "6" then%>SELECTed<%end if%>>6.Produtos Farmacêuticos</option>
                    <option value="7" <%if Seguimento = "7" then%>SELECTed<%end if%>>7.Plásticos e Borracha</option>
                    <option value="8" <%if Seguimento = "8" then%>SELECTed<%end if%>>8.Bebidas</option>
                    <option value="9" <%if Seguimento = "9" then%>SELECTed<%end if%>>9.Papel e Papelão</option>
                    <option value="10" <%if Seguimento = "10" then%>SELECTed<%end if%>>10.Têxteis, Vestuário e Calçados</option>
                    <option value="11" <%if Seguimento = "11" then%>SELECTed<%end if%>>11.Avicultura (Ovos)</option>
                    <option value="12" <%if Seguimento = "12" then%>SELECTed<%end if%>>12.Materiais Elétricos e Comunicação</option>
                    <option value="13" <%if Seguimento = "13" then%>SELECTed<%end if%>>13.Vidros e Cerâmicas</option>
                    <option value="14" <%if Seguimento = "14" then%>SELECTed<%end if%>>14.Fumos</option>
                    <option value="15" <%if Seguimento = "15" then%>SELECTed<%end if%>>15.Metalúrgica</option>
                    <option value="16" <%if Seguimento = "16" then%>SELECTed<%end if%>>16.Mecânica</option>
                    <option value="17" <%if Seguimento = "17" then%>SELECTed<%end if%>>17.Madeiras e Mobiliários</option>
                    <option value="18" <%if Seguimento = "18" then%>SELECTed<%end if%>>18.Materiais de Transporte</option>
                    <option value="19" <%if Seguimento = "19" then%>SELECTed<%end if%>>19.Outros</option>
                <% else 
                  SQL2 = "SELECT AOV_CODSEG, AOV_DESSEG FROM AOV010 WHERE AOV_MSBLQL = '2' AND ISNULL(AOV010.D_E_L_E_T_,'') = ''"
                  SQL2 = SQL2 & "ORDER BY AOV_DESSEG "
                  Dados1.Source = SQL2
                  Dados1.Open()

                  if not Dados1.EOF then
                    while not Dados1.EOF
                    %>
                      <option value="<%=Trim(Dados1("AOV_CODSEG"))%>" <%if Seguimento = Trim(Dados1("AOV_CODSEG")) then%>selected<%end if%>>[<%=Trim(Dados1("AOV_CODSEG"))%>]  <%=Dados1("AOV_DESSEG")%></option>
                    <%
                    Dados1.MoveNext()
                    wend
                  end if
                  Dados1.Close()
                end if
                %>
              </SELECT>
            </div>

            <div class="form-group">
              <span>Observação de arquivos enviados (Emissor) </span>
              <textarea type="text" class="form-control" name="ObsEmissor" id="ObsEmissor" value="<%=ObsEmissor%>" maxlength="200" <%if Request("Emissor") <> "True" then%> readonly <%end if%>><%=ObsEmissor%></textarea>
            </div>

            <div class="form-group">
              <span>Observação de arquivos recebidos (CDAG) </span>
              <textarea type="text" class="form-control" name="ObsCedag" id="ObsCedag" value="<%=ObsCedag%>" maxlength="200" <%if SESSION("CEDAG") = False then%>readonly<%end if%> ><%=ObsCedag%></textarea>
            </div>

            <div class="form-group">
              <label>Codigo SAP</label>
              <input type="text" class="form-control" name="CodSap" id="CodSap" value="<%=CodSap%>" maxlength="30" <%if SESSION("E2OPEN") = True And Request("StatusCli") = "E" then%><%Else%>readonly<%end if%> />
            </div>

            <div class="form-group">&nbsp;</div>
            <div class="d-grid gap-2">
              <button type="button" class="btn btn-info" onclick="ativaaba('0','1','<%=FilialDoRegistro%>','<%=CodigoVendedorPro%>','<%=Natureza%>')">A v a n ç a r</button>
              <button type="button" class="btn btn-light" onclick="location.href='index.asp?act=NEW'">V o l t a r</button>
              <button type="button" class="btn btn-light" onclick="location.href='index.asp?act=LST'">S o l i c i t a ç õ e s</button>
            </div>        
            <p>&nbsp;</p>
          </div>

          <!--===================================================
                                  PROTHEUS
          ====================================================-->
          <div class="tab-pane fade" id="menu1" role="tabpanel" aria-labelledby="nmenu1">
            <div class="tit">
              <div class="alert alert-info divmenu">Complementar PROTHEUS</div>
            </div>

            <div class="form-group">
              <label>Natureza <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <div id="NaturezaDinamicamente">
              </div>            
            </div>            

            <div class="form-group">
              <label>Tipo</label>
              <SELECT name="Tipo" id="Tipo" class="form-select">
                <option value=""></option>
                <option value="P" <%if Tipo = "P" then%>SELECTED<%end if%>>PRODUTOR RURAL</option>
                <option value="R" <%if Tipo = "R" then%>SELECTED<%end if%>>REVENDEDOR</option>
                <option value="C" <%if Tipo = "C" then%>SELECTED<%end if%>>CONS. FINAL</option>
                <option value="S" <%if Tipo = "S" then%>SELECTED<%end if%>>SOLIDARIO</option>
                <option value="E" <%if Tipo = "E" then%>SELECTED<%end if%>>EXPORTAÇÃO</option>
              </SELECT>
            </div>

            <div class="form-group">
              <label>Tipo de Frete <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <SELECT name="TipoFrete" id="TipoFrete" class="form-select">
                <option value=""></option>
                <option value="C" <%if TipoFrete = "C" then%>SELECTED<%end if%>>CIF</option>
                <option value="F" <%if TipoFrete = "F" then%>SELECTED<%end if%>>FOB</option>
              </SELECT>
            </div>

            <div class="form-group">
              <label>TES</label>
              <input type="text" class="form-control protheus" name="Tes" id="Tes" value="<%=Tes%>" maxlength="3" readonly/>
            </div>

            <div class="form-group">
                 <label>Código do Vendedor <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
                 <div id="divCodigoVendedorProthes">
                 </div> 
            </div>

            <div class="form-group">
              <label>Transit Time (Tempo em Trânsito) <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <button type="button" 
                       class="btn btn-link p-0 m-0 align-baseline" 
                       data-toggle="tooltip" 
                       data-placement="right" 
                       title="Refere-se ao período necessário para que uma carga seja transportada de um ponto de origem até o seu destino final.  
                              Esse tempo inclui todas as etapas do processo de transporte, desde a saída do local inicial, passando por eventuais pontos intermiediarios, 
                              ate a chegada ao destino.">
                       <i class="fas fa-question-circle"></i>
               </button>
              <input type="text" onkeypress="return onlynumber();" class="form-control" name="DiasViagem" id="DiasViagem" value="<%=DiasViagem%>" readonly maxlength="1" />
            </div>

            <div class="form-group">
              <label>Tolerância de dias para Entrega a maior <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <button type="button" 
                       class="btn btn-link p-0 m-0 align-baseline" 
                       data-toggle="tooltip" 
                       data-placement="right" 
                       title="Período adicional permitido além do prazo inicialmente previsto para a entrega de um produto quando acordado com o cliente. 
                             Essas margem serve para acomodar possíveis atrasos inesperados no processo logístico, como imprevistos no transporte, condições climáticas adversas, ou dificuldades operacionais, sem que isso seja considerado uma falha no cumprimento do prazo.">
                       <i class="fas fa-question-circle"></i>
               </button>
               <SELECT name="DiasEntregaMaior" id="DiasEntregaMaior" class="form-select">
                    <option value="">Selecione...</option>
                    <option value="0" <%if DiasEntregaMaior = "0" then%>SELECTed<%end if%>>0</option>
                    <option value="1" <%if DiasEntregaMaior = "1" then%>SELECTed<%end if%>>1</option>
                    <option value="2" <%if DiasEntregaMaior = "2" then%>SELECTed<%end if%>>2</option>
               </SELECT>
            </div>

            <div class="form-group">
              <label>Tolerância de dias para Entrega a menor <span class="req"><%if TipoSolicitacao=5 then%><%else%>*<%end if%></span></label>
              <button type="button" 
                       class="btn btn-link p-0 m-0 align-baseline" 
                       data-toggle="tooltip" 
                       data-placement="right" 
                       title="Período adicional permitido além do prazo inicialmente previsto para a entrega de um produto quando acordado com o cliente. 
                             Essas margem serve para acomodar possíveis atrasos inesperados no processo logístico, como imprevistos no transporte, condições climáticas adversas, ou dificuldades operacionais, sem que isso seja considerado uma falha no cumprimento do prazo.">
                       <i class="fas fa-question-circle"></i>
               </button>
               <SELECT name="DiasEntregaMenor" id="DiasEntregaMenor" class="form-select">
                    <option value="">Selecione...</option>
                    <option value="0" <%if DiasEntregaMenor = "0" then%>SELECTed<%end if%>>0</option>
                    <option value="1" <%if DiasEntregaMenor = "1" then%>SELECTed<%end if%>>-1</option>
                    <option value="2" <%if DiasEntregaMenor = "2" then%>SELECTed<%end if%>>-2</option>
                    <option value="3" <%if DiasEntregaMenor = "3" then%>SELECTed<%end if%>>-3</option>
                    <option value="4" <%if DiasEntregaMenor = "4" then%>SELECTed<%end if%>>-4</option>
               </SELECT>
            </div>

            <div class="form-group">
              <label>Endereço de Entrega</label>
              <div class="form-check form-switch">
                  <input id="cbxEnderecoEntrega" name="cbxEnderecoEntrega" class="form-check-input" type="checkbox" onclick="habilitaEnderecoEntrega('Entrega')">
                  <label for="cbxEnderecoEntrega" class="form-check-label">Endereço de Entrega diferente do cadastro</label>
              </div> 
            </div>
            <div id="divEnderecoEntrega" style="display: none;">
                <div class="form-group">
                  <label>Rua</label>
                  <input type="text" onchange="endent(this.value)" class="form-control" name="RuaEntPro" id="RuaEntPro" value="<%=RuaEntPro%>" maxlength="50" />
                </div>

                <div class="form-group">
                  <label>Número</label>
                  <input type="text" onchange="nument(this.value)" class="form-control" name="NumeroEntPro" id="NumeroEntPro" value="<%=NumeroEntPro%>" maxlength="10" />
                </div>

                <div class="form-group">
                  <label>Bairro</label>
                  <input type="text" onchange="bairroent(this.value)" class="form-control" name="BairroEntPro" id="BairroEntPro" value="<%=BairroEntPro%>" maxlength="40" />
                </div>

                <div class="form-group">
                  <label>CEP</label>
                  <input type="text" autocomplete="off" onchange="formatacep(this.value,this.id);cepent(this.value)" onkeypress="return onlynumber();" class="form-control" name="CepEntPro" id="CepEntPro" value="<%=formatcep(CepEntPro)%>" maxlength="8" />
                </div>

                <div class="form-group">
                  <label>Estado</label>
                  <SELECT onchange="ufent(this.value); municipio(this.value, 'CidadeEntPro', 'cident', 'dCidadeEntPro')" name="EstadoEntPro" id="EstadoEntPro" class="form-select">
                    <option value=""></option>
                    <%
                      SQL2 = "SELECT DISTINCT ESTADO FROM UnificadoCidades ORDER BY ESTADO "
                      Dados12.Source = SQL2
                      Dados12.Open()
                      if not Dados12.EOF then
                        while not Dados12.EOF%>
                          <option value="<%=Dados12("ESTADO")%>" <%if EstadoEntPro = Dados12("ESTADO") then%>SELECTed<%end if%>><%=Dados12("ESTADO")%></option>
                        <%Dados12.MoveNext()
                        wend
                      end if
                      Dados12.Close()
                    %>
                  </SELECT>
                </div>

                <div class="form-group" id="dCidadeEntPro">
                    <label>Cidade </label>
                    <input type="text" readonly class="form-control protheus" name="CidadeEntPro" id="CidadeEntPro" value="<%=CidadeEntPro%>" readonly />
                </div>

                <div class="form-group">
                  <label>KM <span class="req"><%if TipoSolicitacao<>5 then%>*<%else%><%end if%></span></label>
                  <input type="text" class="form-control" name="KmEnderecoCobranca" id="KmEnderecoCobranca" onblur="CalcularDistancia(this)" value="<%=KmEnderecoCobranca%>" maxlength="6" />
                </div>                  
            </div>
            <div class="form-group">
              <label>Endereço de Cobrança</label>
              <div class="form-check form-switch">
                  <input id="cbxEnderecoCobranca" name="cbxEnderecoCobranca" class="form-check-input" type="checkbox" onclick="habilitaEnderecoEntrega('Cobranca')">
                  <label for="cbxEnderecoCobranca" class="form-check-label">Endereço de Cobrança diferente do cadastro</label>
              </div> 
            </div>
            <div id="divEnderecoCobranca" style="display: none;">
                <div class="form-group">
                  <label>Rua</label>
                  <input type="text" onchange="endcob(this.value)" class="form-control" name="RuaCobPro" id="RuaCobPro" value="<%=RuaCobPro%>" maxlength="50" />
                </div>

                <div class="form-group">
                  <label>Número</label>
                  <input type="text" onchange="numcob(this.value)" class="form-control" name="NumeroCobPro" id="NumeroCobPro" value="<%=NumeroCobPro%>" maxlength="10" />
                </div>

                <div class="form-group">
                  <label>Bairro</label>
                  <input type="text" onchange="bairrocob(this.value)" class="form-control" name="BairroCobPro" id="BairroCobPro" value="<%=BairroCobPro%>" maxlength="40" />
                </div>

                <div class="form-group">
                  <label>CEP</label>
                  <input type="text" autocomplete="off" onchange="formatacep(this.value,this.id);cepcob(this.value)" onkeypress="return onlynumber();" class="form-control" name="CepCobPro" id="CepCobPro" value="<%=formatcep(CepCobPro)%>" maxlength="8" />
                </div>

                <div class="form-group">
                  <label>Estado</label>
                  <SELECT onchange="ufcob(this.value);municipio(this.value,'CidadeCobPro','cidcob','dCidadeCobPro')" name="EstadoCobPro" id="EstadoCobPro" class="form-select">
                    <option value=""></option>
                    <%
                      SQL2 = "SELECT DISTINCT ESTADO FROM UnificadoCidades ORDER BY ESTADO "
                      Dados12.Source = SQL2
                      Dados12.Open()
                      if not Dados12.EOF then
                        while not Dados12.EOF%>
                          <option value="<%=Dados12("ESTADO")%>" <%if EstadoCobPro = Dados12("ESTADO") then%>SELECTed<%end if%>><%=Dados12("ESTADO")%></option>
                        <%Dados12.MoveNext()
                        wend
                      end if
                      Dados12.Close()
                    %>
                  </SELECT>
                </div>

                <div class="form-group" id="dCidadeCobPro">
                    <label>Cidade </label>
                    <input type="text" class="form-control protheus" name="CidadeCobPro" id="CidadeCobPro" value="<%=CidadeCobPro%>" readonly />
                </div>
            </div>

            <div class="form-group" id="divTqtdeMaior">
                <label>Tolerância de Quantidade a maior (%)<span class="req"><%if TipoSolicitacao<>5 then%>*<%else%><%end if%></span></label>
                <input type="text" class="form-control protheus" name="TQtdeMaior" id="TQtdeMaior" value="<%=TQtdeMaior%>"  />
            </div>

            <div class="form-group" id="divTqtdeMenor">
                <label>Tolerância de Quantidade a menor (%)<span class="req"><%if TipoSolicitacao<>5 then%>*<%else%><%end if%></span></label>
                <input type="text" class="form-control protheus" name="TQtdeMenor" id="TQtdeMenor" value="<%=TQtdeMenor%>"  />
            </div>                        

            <div class="form-group">
              <label>Codigo PROTHEUS</label>
              <input type="text" class="form-control protheus" name="CodPro" id="CodPro" value="<%=CodPro%>" maxlength="30" <%if SESSION("E2OPEN") = True And Request("StatusCli") = "E" then%><%Else%>readonly<%end if%> />
            </div>

            <div class="form-group">
              <label>Loja</label>
              <input type="text" class="form-control protheus" name="CodLoja" id="CodLoja" value="<%=CodLoja%>" maxlength="05" <%if SESSION("E2OPEN") = True And Request("StatusCli") = "E" then%><%Else%>readonly<%end if%> />
            </div>            

            <div class="form-group">&nbsp;</div>
            <div class="d-grid gap-2">
              <button type="button" class="btn btn-light" onclick="ativaaba('1','0','')">V o l t a r</button>
              <button type="button" class="btn btn-light" onclick="ExecutaSolicitacao('<%=PreCadastro%>')">S o l i c i t a ç õ e s</button>
              <%if Edita then
                if Request("ns")="F" then%>
                  <button type="button" class="btn btn-success btn-block" onclick="validacamposFinanceiro('<%=codigo%>','<%=TipoSolicitacao%>','<%=FilialDoRegistro%>')"><%= TxtAct %></button>
                <% Else %>
                  <button type="button" class="btn btn-success btn-block" onclick="validacampos('<%=altvis%>','<%=Request("ns")%>','','<%=TipoSolicitacao%>')"><%= TxtAct %></button>
                <%end if%>                
              <%end if%>
              <% if Request("vs") = "True" then %>
                <% if Request("ca") = "True" then %>
                  <% if Request("ns") = "C" And Request("StatusCli") = "E" then %>
                    <button type="button" class="btn btn-secondary btn-block" title="Não Conforme" onclick="validacampos('<%=altvis%>','<%=Request("ns")%>','I','<%=TipoSolicitacao%>')" ><i class="fas fa-times"></i> Não Conforme</button>
                  <% Else %>
                    <button type="button" class="btn btn-secondary btn-block" title="Cancelar" onclick="validacampos('<%=altvis%>','<%=Request("ns")%>','X','<%=TipoSolicitacao%>')" ><i class="fas fa-times"></i> Cancelar o cadastro</button>
                  <% end if %>
                <% end if %>
                
                <% if Request("dv") = "True" then %>
                  <button type="button" class="btn btn-warning btn-block" title="Devolver para o emissor" onclick="validacampos('<%=altvis%>','<%=Request("ns")%>','R','<%=TipoSolicitacao%>')"><i class="fas fa-undo-alt"></i> Devolver para o emissor</button>
                <% end if %>
              <% end if %>
            </div>
            <p>&nbsp;</p>
          </div>
        </div> <!-- div tab-content -->
      </div> <!-- div col col-lg-8 -->
          
      <div class="modal" id="myOtherModal"><!-- The Modal -->
        <div class="modal-dialog">
          <div class="modal-content">
      
            <!-- Modal Header -->
            <div class="modal-header">
              <h4 class="modal-title">Tipos de Cliente primário</h4>
              <button type="button" class="close" data-dismiss="modal">&times;</button>
            </div>
        
            <!-- Modal body -->
            <div class="modal-body">
              <img src="Tipos.png" style="border:none" alt="Imagens de tipos suportados." title="Tipos suportados." /> 
            </div>
        
            <!-- Modal footer -->
            <div class="modal-footer">
              <button type="button" class="btn btn-danger" data-dismiss="modal">Close</button>
            </div>
        
          </div>
        </div>
      </div>

      <div class="modal fade" id="myModal2">
        <div class="modal-dialog" style="min-width: 80%;">
          <div class="modal-content">
            <div class="modal-header" id="modalheader2">
              <h6 class="modal-title" id="modaltitle2"></h6>
              <button type="button" class="close" id="modalbtnx2" data-dismiss="modal">&times;</button>
            </div>
            <div class="modal-body table-responsive" style="max-height: 400px;background-color: gray;" id="modalbody2"></div>
            <div class="modal-footer">
              <button type="button" class="btn btn-sm btn-light" id="modalbtnyes" style="display:none;width:149px" onclick="location.href='index.asp?act=LST'">Sim</button>
              <button type="button" class="btn btn-sm btn-light" id="modalbtnnot" data-dismiss="modal" style="display:none;width:149px">Não</button>
              <button type="button" class="btn btn-sm btn-light" id="modalbtnclose2" data-dismiss="modal">Fechar</button>
            </div>
          </div>
        </div>
      </div>

    </div> <!-- div justify-content -->
  </div> <!-- div container -->

  <div class="modal fade" id="myModal5">
    <div class="modal-dialog">
      <div class="modal-content">
        <div class="modal-header" id="modalheader5">
          <h6 class="modal-title" id="modaltitle5">Visto Cadastro Unificado</h6>
          <button type="button" class="close" id="modalbtnx5" data-dismiss="modal">&times;</button>
        </div>
        <div class="modal-body" id="modalbody5">
          <p id="vper">Confirma prosseguir?</p>
          <input type="hidden" name="NovoStatus" id="NovoStatus">
          <div id="divcontrole" style="display: none;">
            <SELECT name="controle" id="controle" class="form-SELECT" onchange="document.getElementById('NovoStatus').value=this.value" style="background-color: lightyellow">
              <option value="" style="display: none;">Selecione...</option>
              <option value="I">CONTROLES INTERNOS</option>
              <option value="C">CONTROLADORIA</option>
            </SELECT><br><br>
          </div>
          <textarea name="vobs" id="vobs" rows="3" style="width:100%" maxlength="100" placeholder="Comentários"></textarea>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-sm btn-danger" id="btnsubmitd" onclick="validavisto('R')" style="width:149px; display: none;">Devolver</button>
          <button type="button" class="btn btn-sm btn-success" id="btnsubmit" onclick="validavisto('<%=Acao%>')" style="width:149px">Confirmar</button>
          <button type="button" class="btn btn-sm btn-light" data-dismiss="modal" style="width:149px">Cancelar</button>
        </div>
      </div>
    </div>
  </div>
  </form>

  <script>
    function grpconta() {
      var op = document.getElementById("ContaConciliacao").getElementsByTagName("option");
      if (document.getElementById('GrupoContas').value == 'Z003') 
      {
        document.getElementById('ContaConciliacao').value = '41100000';
        var b = true;
      } else {
        var b = false;
      }

      for (var i = 0; i < op.length; i++) 
      {
        if (op[i].value != "41100000") 
        {
          op[i].disabled = b;
        }
      }
    }

    <%if Acao <> "GRI" then%>
      grpconta();
    <%end if%>
  </script>
  <%
End Function

Function rp(v)
  v = Trim(v)
  rp = Replace(v,"'","")
End Function

Function rpPonto(v)
  v = Trim(v)
  rpPonto = Replace(Replace(v,".",""),",",".")
End Function

Function fnumber(v)
  if v = "" then
    fnumber = ""
  Else
    fnumber = Replace(Replace(FormatNumber(v,2),".",""),",",".")
  end if
End Function

Function ShowMsg(Mensagem)
  response.Write "<script>alert('" & Mensagem & "')</script>"
End Function
%>

<!-- Modal -->
<div class="modal fade" id="staticBackdrop" data-bs-backdrop="static" data-bs-keyboard="false" tabindex="-1" aria-labelledby="staticBackdropLabel" aria-hidden="true">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header" id="modalheader">
        <h5 class="modal-title" id="staticBackdropLabel">Modal title</h5>
      </div>
      <div class="modal-body" id="modalbody">
        ...
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-sm btn-light" id="modalbtnyes" style="display:none;width:149px">Sim</button>
        <button type="button" class="btn btn-sm btn-light" id="modalbtnnot" data-bs-dismiss="modal" style="display:none;width:149px">Não</button>
        <button type="button" class="btn btn-sm btn-light" id="modalbtnclose" data-bs-dismiss="modal">Fechar</button>
      </div>
    </div>
  </div>
</div>

<%
    Function VerificaSessaoExpirada()
      if Trim(Session("CODUSER")) = "" then
        Call SesseaoExpirada()
      end if
    End function   

    function formatcep(cep)
      if cep <> "" then
        formatcep = mid(cep,1,2)&"."&mid(cep,3,3)&"-"&mid(cep,6,3)
      end if
    end function

    function formatacpf(cpf)
        if len(cpf) = 11 then
            digito = Right(cpf,2) 
            seq3 = mid(cpf,7,3)
            seq2 = mid(cpf,4,3)
            seq1 = mid(cpf,1,3)
            formatacpf= seq1&"."&seq2&"."&seq3&"-"&digito
        else
            formatacpf = cpf
        end if
    end function

    Function FormataData(data)
      if data = null then
         dataAmostra = ""       
      else
        if data <> "1900-01-01" then
          dataAmostra = replace(data,"-","")
          if len(dataAmostra) > 8 then
              dataAmostra = InverteData(dataAmostra)
          end if
        else
          dataAmostra = ""
        end if
      end if
      FormataData = dataAmostra
    End function

    Function AmostraEnviada(statusRegistro, DataAmostra)
      if statusRegistro <> "P" and InverteData(Date()) >= FormataData(DataAmostra) and FormataData(DataAmostra) <> ""  then
        AmostraEnviada = True
      else
        AmostraEnviada = False     
      end if
    end Function 
    
%>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
</body>
</html>
<%
function ChamarINC()
         FormUNI "GRI", "Incluir", "","", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", false, "",10,10,iif(TipoSolicitacao=5,"P",""),"","","","","",""
end function

function ChamarGRI()
         VerificaSessaoExpirada()
         SELECT CASE Request("Filial")
                CASE "00" Empresa = "BR02"
                CASE "04" Empresa = "BR02"
                CASE "05" Empresa = "BR03"
                CASE "06" Empresa = "BR04"
                CASE "07" Empresa = "BR07"
          END SELECT   
          axpercentualDiferimento = iif(Request("percentualDiferimento")="",0,Request("percentualDiferimento"))

          SQL = "BEGIN TRANSACTION "
          SQL = SQL + "INSERT INTO UnificadoCliente "
          SQL = SQL + "(Codigo "
          SQL = SQL + ",GrupoContas "
          SQL = SQL + ",Empresa "
          SQL = SQL + ",RazaoSocial "
          SQL = SQL + ",NomeFantasia "
          SQL = SQL + ",Rua "
          SQL = SQL + ",Numero "
          SQL = SQL + ",Bairro "
          SQL = SQL + ",Cep "
          SQL = SQL + ",Cidade "
          SQL = SQL + ",Estado "
          SQL = SQL + ",Pais "
          SQL = SQL + ",Idioma "
          SQL = SQL + ",Telefone "
          SQL = SQL + ",Fax "
          SQL = SQL + ",EmailXml "
          SQL = SQL + ",EmailPagamento "
          SQL = SQL + ",CodigoVendedor "
          SQL = SQL + ",CodigoHfm "
          SQL = SQL + ",Cpf "
          SQL = SQL + ",Cnpj "
          SQL = SQL + ",InscricaoEstadual "
          SQL = SQL + ",InscricaoMunicipal "
          SQL = SQL + ",Declarante "
          SQL = SQL + ",Contribuinte "
          SQL = SQL + ",NaturezaLegal "
          SQL = SQL + ",TipoIndustria "
          SQL = SQL + ",CodigoSuframa "
          SQL = SQL + ",ContaConciliacao "
          SQL = SQL + ",ContaBanco "
          SQL = SQL + ",ChaveOrdenacao "
          SQL = SQL + ",CondicaoPagamento "
          SQL = SQL + ",MetodoPagamento "
          SQL = SQL + ",CodigoJuros "
          SQL = SQL + ",CodigoAntigo "
          SQL = SQL + ",Correspondencia "
          SQL = SQL + ",LimiteCredito "
          SQL = SQL + ",VolumeVendam2 "
          SQL = SQL + ",PrecoEstm2 "
          SQL = SQL + ",PrecoEstTon "
          SQL = SQL + ",LimiteCredSug "
          SQL = SQL + ",PrazoPag "
          SQL = SQL + ",DtVenda "
          SQL = SQL + ",Seguimento "
          SQL = SQL + ",ResponsavelAprovacao "
          SQL = SQL + ",Natureza "
          SQL = SQL + ",Tipo "
          SQL = SQL + ",TipoFrete "
          SQL = SQL + ",CodigoVendedorPro "
          SQL = SQL + ",CodigoClasProd "
          SQL = SQL + ",DiasViagem "
          SQL = SQL + ",DiasEntregaMaior "
          SQL = SQL + ",DiasEntregaMenor "
          SQL = SQL + ",RuaEntPro "
          SQL = SQL + ",NumeroEntPro "
          SQL = SQL + ",BairroEntPro "
          SQL = SQL + ",CepEntPro "
          SQL = SQL + ",CidadeEntPro "
          SQL = SQL + ",EstadoEntPro "
          SQL = SQL + ",RuaCobPro "
          SQL = SQL + ",NumeroCobPro "
          SQL = SQL + ",BairroCobPro "
          SQL = SQL + ",CepCobPro "
          SQL = SQL + ",CidadeCobPro "
          SQL = SQL + ",EstadoCobPro "
          SQL = SQL + ",ContatoCompras "
          SQL = SQL + ",Departamento "
          SQL = SQL + ",TelefoneSia "
          SQL = SQL + ",Email "
          SQL = SQL + ",PrecoPadrao "
          SQL = SQL + ",TransacaoPedidos "
          SQL = SQL + ",AtividadePincipal "
          SQL = SQL + ",GrupoEmpresa "
          SQL = SQL + ",AreaVendas "
          SQL = SQL + ",NomeRepresentante "
          SQL = SQL + ",EmailRepresentante "
          SQL = SQL + ",TelefoneRep "
          SQL = SQL + ",SegmentacaoMercado "
          SQL = SQL + ",AvisaEmbarque "
          SQL = SQL + ",ContEmbalagens "
          SQL = SQL + ",RegraEmbarque "
          SQL = SQL + ",DiasTransporte "
          SQL = SQL + ",TolEntregaMenos "
          SQL = SQL + ",TolEntregaMais "
          SQL = SQL + ",JanTolMenos "
          SQL = SQL + ",JanTolMais "
          SQL = SQL + ",EmiteCertificado "
          SQL = SQL + ",Frete "
          SQL = SQL + ",DesprezarTarifa "
          SQL = SQL + ",DistRomaneio "
          SQL = SQL + ",Del "
          SQL = SQL + ",QuemIncluiu "
          SQL = SQL + ",DataEmissao "
          SQL = SQL + ",HoraEmissao "
          SQL = SQL + ",StatusCli "
          SQL = SQL + ",QuemIncluiuCod "
          SQL = SQL + ",ObsEmissor "
          SQL = SQL + ",ObsCedag "
          SQL = SQL + ",CodSap "
          SQL = SQL + ",CodPro "
          SQL = SQL + ",CodSia "
          SQL = SQL + ",Filial "
          SQL = SQL + ",Tes "
          SQL = SQL + ",Eshop "
          SQL = SQL + ",TQtdeMaior "
          SQL = SQL + ",TQtdeMenor "
          SQL = SQL + ",PreCadastro "
          SQL = SQL + ",CodLoja "
          SQL = SQL + ",KmEndereco "
          SQL = SQL + ",KmEnderecoCobranca "
          SQL = SQL + ",isencaoIPI "
          SQL = SQL + ",IcmsDiferido "
          SQL = SQL + ",percentualDiferimento) "
          SQL = SQL + "VALUES "
          SQL = SQL + "((SELECT ISNULL(MAX(CAST(Codigo AS bigint)),999)+1 FROM UnificadoCliente) "
          SQL = SQL + ",'" + Request("GrupoContas") + "'"
          SQL = SQL + ",'" + Empresa + "'"
          SQL = SQL + ",'" + ucase(Request("RazaoSocial")) + "'"
          SQL = SQL + ",'" + ucase(Request("NomeFantasia")) + "'"
          SQL = SQL + ",'" + ucase(Request("Rua")) + "'"
          SQL = SQL + ",'" + Request("Numero") + "'"
          SQL = SQL + ",'" + ucase(Request("Bairro")) + "'"
          SQL = SQL + ",'" + Trim(Replace(Replace(Request("Cep"),".",""),"-","")) + "'"
          SQL = SQL + ",'" + ucase(Trim(Request("Cidade"))) + "'"
          SQL = SQL + ",'" + Request("Estado") + "'"
          SQL = SQL + ",'" + Request("Pais") + "'"
          SQL = SQL + ",'PT-BR'"
          SQL = SQL + ",'" + Trim(Replace(Replace(Replace(Request("Telefone"),"(",""),")",""),"-","")) + "'"
          SQL = SQL + ",''"
          SQL = SQL + ",'" + Request("EmailXml") + "'"
          SQL = SQL + ",'" + Request("EmailPagamento") + "'"
          SQL = SQL + ",''"
          SQL = SQL + ",'" + Request("CodigoHfm") + "'"
          SQL = SQL + ",'" + Trim(Replace(Replace(Request("Cpf"),".",""),"-","")) + "'"
          SQL = SQL + ",'" + Trim(Replace(Replace(Replace(Request("Cnpj"),".",""),"-",""),"/","")) + "'"
          SQL = SQL + ",'" + Request("InscricaoEstadual") + "'"
          SQL = SQL + ",'" + Request("InscricaoMunicipal") + "'"
          SQL = SQL + ",'" + Request("Declarante") + "'"
          SQL = SQL + ",'" + Request("Contribuinte") + "'"
          SQL = SQL + ",''"
          SQL = SQL + ",'" + Request("TipoIndustria") + "'"
          SQL = SQL + ",''"
          SQL = SQL + ",'" + Request("ContaConciliacao")  + "'"
          SQL = SQL + ",''"
          SQL = SQL + ",'" + Request("ChaveOrdenacao") + "'"
          SQL = SQL + ",'" + Request("CondicaoPagamento") + "'"
          SQL = SQL + ",'" + Request("MetodoPagamento") + "'"
          SQL = SQL + ",'" + Request("CodigoJuros") + "'"
          SQL = SQL + ",''"
          SQL = SQL + ",'" + Request("Correspondencia") + "'"
          SQL = SQL + ",'" + Request("LimiteCredito") & "'"
          SQL = SQL + ",'" + Request("VolumeVendam2") + "'"
          SQL = SQL + ",'" + Request("PrecoEstm2") + "'"
          SQL = SQL + ",'" + Request("PrecoEstTon") + "'"
          SQL = SQL + ",'" + Request("LimiteCredSug") + "'"
          SQL = SQL + ",'" + Request("PrazoPag") + "'"
          SQL = SQL + ",'" + DataInvertidaSemSeparador(Request("DtVenda")) + "'"
          SQL = SQL + ",'" + Request("Seguimento") + "'"
          SQL = SQL + ",'" + Request("ResponsavelAprovacao") + "'"
          SQL = SQL + ",'" + Request("Natureza") + "'"
          SQL = SQL + ",'" + Request("Tipo") + "'"
          SQL = SQL + ",'" + Request("TipoFrete") + "'"
          SQL = SQL + ",'" + Request("CodigoVendedorProtheus") + "'"
          SQL = SQL + ",'" + Request("CodigoClasProd") + "'"
          SQL = SQL + ",'" + Request("DiasViagem") + "'"
          SQL = SQL + "," + iif(Request("DiasEntregaMaior")="","null",Request("DiasEntregaMaior")) 
          SQL = SQL + "," + iif(Request("DiasEntregaMenor")="","null",Request("DiasEntregaMenor"))
          SQL = SQL + ",'" + Request("RuaEntPro") + "'"
          SQL = SQL + ",'" + Request("NumeroEntPro") + "'"
          SQL = SQL + ",'" + Request("BairroEntPro") + "'"
          SQL = SQL + ",'" + Trim(Replace(Replace(Request("CepEntPro"),".",""),"-","")) + "'"
          SQL = SQL + ",'" + Request("CidadeEntPro") + "'"
          SQL = SQL + ",'" + Request("EstadoEntPro") + "'"
          SQL = SQL + ",'" + Request("RuaCobPro") + "'"
          SQL = SQL + ",'" + Request("NumeroCobPro") + "'"
          SQL = SQL + ",'" + Request("BairroCobPro")  + "'"
          SQL = SQL + ",'" + Trim(Replace(Replace(Request("CepCobPro"),".",""),"-",""))  + "'"
          SQL = SQL + ",'" + Trim(Request("CidadeCobPro")) + "'"
          SQL = SQL + ",'" + Request("EstadoCobPro") + "'"
          SQL = SQL + ",'" + Request("Email") + "'"
          SQL = SQL + ",'" + Request("PrecoPadrao") + "'"
          SQL = SQL + ",'" + Request("TransacaoPedidos") + "'"
          SQL = SQL + ",'" + Request("AtividadePincipal") + "'"
          SQL = SQL + ",'" + Request("GrupoEmpresa") + "'"
          SQL = SQL + ",'" + Request("AreaVendas") + "'"
          SQL = SQL + ",'" + Request("NomeRepresentante") + "'"
          SQL = SQL + ",'" + Request("EmailRepresentante")  + "'"
          SQL = SQL + ",'" + Trim(Replace(Replace(Replace(Request("TelefoneRep"),"(",""),")",""),"-","")) + "'"
          SQL = SQL + ",'" + Request("SegmentacaoMercado") + "'"
          SQL = SQL + ",'" + Request("AvisaEmbarque")  + "'"
          SQL = SQL + ",'" + Request("ContEmbalagens") + "'"
          SQL = SQL + ",'" + Request("RegraEmbarque") + "'"
          SQL = SQL + ",'" + Request("DiasTransporte") + "'"
          SQL = SQL + ",'" + Request("TolEntregaMenos") + "'"
          SQL = SQL + ",'" + Request("TolEntregaMais") + "'"
          SQL = SQL + ",'" + Request("JanTolMenos") + "'"
          SQL = SQL + ",'" + Request("JanTolMais") + "'"
          SQL = SQL + ",'" + Request("EmiteCertificado") + "'"
          SQL = SQL + ",''" '+ Request("Frete") + "'"
          SQL = SQL + ",''" '+ Request("DesprezarTarifa") + "'"
          SQL = SQL + ",''" '+ Request("DistRomaneio") + "'"
          SQL = SQL + ",''" 
          SQL = SQL + ",''" 
          SQL = SQL + ",''" 
          SQL = SQL + ",'0'" 'DEL
          SQL = SQL + ",'" + SESSION("USR") + "'"
          SQL = SQL + ",'" + Hoje() + "'"
          SQL = SQL + ",'" + Agora() + "'"
          SQL = SQL + ",'P'"
          SQL = SQL + ",'" + SESSION("CODIGO")  + "'"
          SQL = SQL + ",'" + Request("ObsEmissor") + "'"
          SQL = SQL + ",'" + Request("ObsCedag") + "'"
          SQL = SQL + ",'" + Request("CodSap") + "'"
          SQL = SQL + ",'" + Request("CodPro") + "'"
          SQL = SQL + ",'" + Request("CodSia") + "'"
          SQL = SQL + ",'" + Request("Filial") + "'"
          SQL = SQL + ",'" + Request("Tes") + "'"
          SQL = SQL + ",'" + Request("eshop") + "'"
          SQL = SQL + "," + Request("TQtdeMaior")
          SQL = SQL + "," + Request("TQtdeMenor") 
          SQL = SQL + ",'" + iif(TipoSolicitacao="5","P","") + "'"
          SQL = SQL + ",'" + Request("CodLoja") + "'" 
          SQL = SQL + ",'" + Request("KmEndereco") + "'" 
          SQL = SQL + ",'" + Request("KmEnderecoCobranca") + "'" 
          SQL = SQL + ","  + iif(Request("isencaoIPI")="","null",Request("isencaoIPI")) 
          SQL = SQL + ",'" + iif(Request("IcmsDiferido")="","0",Request("IcmsDiferido")) + "'" 
          SQL = SQL & "," & iif(Request("percentualDiferimento")="",0,Request("percentualDiferimento")) & ")"
          SQL = SQL + " IF @@ERROR <> 0 "
          SQL = SQL +   "ROLLBACK TRANSACTION "
          SQL = SQL + "ELSE "
          SQL = SQL +   "COMMIT TRANSACTION "

          Dados11.Open(SQL)

          response.Write "<script>location.href='index.asp'</script>"
end function

function ChamarALT()
         VerificaSessaoExpirada()

          SQL = "SELECT Codigo, "
    SQL = SQL & "GrupoContas, "
    SQL = SQL & "Empresa, "
    SQL = SQL & "RazaoSocial, "
    SQL = SQL & "NomeFantasia, "
    SQL = SQL & "Rua, "
    SQL = SQL & "Numero, "
    SQL = SQL & "Bairro, "
    SQL = SQL & "Cep, "
    SQL = SQL & "Cidade, "
    SQL = SQL & "Estado, "
    SQL = SQL & "Pais, "
    SQL = SQL & "Idioma, "
    SQL = SQL & "Telefone, "
    SQL = SQL & "Fax, "
    SQL = SQL & "EmailXml, "
    SQL = SQL & "EmailPagamento, "
    SQL = SQL & "CodigoVendedor, "
    SQL = SQL & "CodigoHfm, "
    SQL = SQL & "Cpf, "
    SQL = SQL & "Cnpj, "
    SQL = SQL & "InscricaoEstadual, "
    SQL = SQL & "InscricaoMunicipal, "
    SQL = SQL & "Declarante, "
    SQL = SQL & "Contribuinte, "
    SQL = SQL & "NaturezaLegal, "
    SQL = SQL & "TipoIndustria, "
    SQL = SQL & "CodigoSuframa, "
    SQL = SQL & "ContaConciliacao, "
    SQL = SQL & "ContaBanco, "
    SQL = SQL & "ChaveOrdenacao, "
    SQL = SQL & "CondicaoPagamento, "
    SQL = SQL & "MetodoPagamento, "
    SQL = SQL & "CodigoJuros,  "
    SQL = SQL & "CodigoAntigo,  "
    SQL = SQL & "Correspondencia,  "
    SQL = SQL & "LimiteCredito,  "
    SQL = SQL & "VolumeVendam2, " 
    SQL = SQL & "PrecoEstm2,  "
    SQL = SQL & "PrecoEstTon,  "
    SQL = SQL & "LimiteCredSug,  "
    SQL = SQL & "PrazoPag,  "
    SQL = SQL & "DtVenda,  "
    SQL = SQL & "Seguimento,  "
    SQL = SQL & "ResponsavelAprovacao,  "
    SQL = SQL & "Natureza,  "
    SQL = SQL & "Tipo,  "
    SQL = SQL & "TipoFrete, "
    SQL = SQL & "CodigoVendedorPro, " 
    SQL = SQL & "CodigoClasProd,  "
    SQL = SQL & "DiasViagem,  "
    SQL = SQL & "DiasEntregaMaior,  "
    SQL = SQL & "DiasEntregaMenor,  "
    SQL = SQL & "RuaEntPro,  "
    SQL = SQL & "NumeroEntPro,  "
    SQL = SQL & "BairroEntPro, " 
    SQL = SQL & "CepEntPro,  "
    SQL = SQL & "CidadeEntPro,  "
    SQL = SQL & "EstadoEntPro, " 
    SQL = SQL & "RuaCobPro,  "
    SQL = SQL & "NumeroCobPro,  "
    SQL = SQL & "BairroCobPro,  "
    SQL = SQL & "CepCobPro, "
    SQL = SQL & "CidadeCobPro,  "
    SQL = SQL & "EstadoCobPro,  "
    SQL = SQL & "ContatoCompras,  "
    SQL = SQL & "Departamento,  "
    SQL = SQL & "Email,  "
    SQL = SQL & "PrecoPadrao, "
    SQL = SQL & "TransacaoPedidos,  "
    SQL = SQL & "AtividadePincipal,  "
    SQL = SQL & "GrupoEmpresa,  "
    SQL = SQL & "AreaVendas,  "
    SQL = SQL & "NomeRepresentante,  "
    SQL = SQL & "EmailRepresentante,  "
    SQL = SQL & "TelefoneRep,  "
    SQL = SQL & "SegmentacaoMercado,  "
    SQL = SQL & "AvisaEmbarque, " 
    SQL = SQL & "ContEmbalagens, " 
    SQL = SQL & "RegraEmbarque,  "
    SQL = SQL & "DiasTransporte, "
    SQL = SQL & "TolEntregaMenos,  "
    SQL = SQL & "TolEntregaMais, " 
    SQL = SQL & "JanTolMenos,  "
    SQL = SQL & "JanTolMais,  "
    SQL = SQL & "EmiteCertificado,  "
    SQL = SQL & "Frete,  "
    SQL = SQL & "DesprezarTarifa, " 
    SQL = SQL & "DistRomaneio,  "
    SQL = SQL & "ObsEmissor,  "
    SQL = SQL & "ObsCedag,  "
    SQL = SQL & "CodSap,  "
    SQL = SQL & "CodPro,  "
    SQL = SQL & "CodSia,  "
    SQL = SQL & "Filial, " 
    SQL = SQL & "Tes,  "
    SQL = SQL & "Eshop,  "
    SQL = SQL & "DataEmissao, "
    SQL = SQL & "TQtdeMaior, "
    SQL = SQL & "TQtdeMenor, "
    SQL = SQL & "PreCadastro, "
    SQL = SQL & "CodLoja, "
    SQL = SQL & "KmEndereco, "
    SQL = SQL & "KmEnderecoCobranca, "
    SQL = SQL & "isencaoIPI, "
    SQL = SQL & "IcmsDiferido, "
    SQL = SQL & "percentualDiferimento "
    SQL = SQL & "FROM UnificadoCliente "
    SQL = SQL & "WHERE Codigo = '" & Request("cod") & "' "
		Dados11.Open(SQL)

    FormUNI "GRA", "Alterar", Dados11("Codigo"), Dados11("GrupoContas"), Dados11("Empresa"), Dados11("RazaoSocial"), Dados11("NomeFantasia"), _
                              Dados11("Rua"), Dados11("Numero"), Dados11("Bairro"), Dados11("Cep"), Dados11("Cidade"), Dados11("Estado"), Dados11("Pais"), _
                              Dados11("Idioma"), Dados11("Telefone"), Dados11("Fax"), Dados11("EmailXml"), Dados11("EmailPagamento"), _ 
                              Dados11("CodigoHfm"), Dados11("Cpf"), Dados11("Cnpj"), Dados11("InscricaoEstadual"), Dados11("InscricaoMunicipal"), Dados11("Declarante"), _
                              Dados11("Contribuinte"), Dados11("NaturezaLegal"), Dados11("TipoIndustria"), Dados11("CodigoSuframa"), Dados11("ContaConciliacao"), _
                              Dados11("ContaBanco"), Dados11("ChaveOrdenacao"), Dados11("CondicaoPagamento"), Dados11("MetodoPagamento"), Dados11("CodigoJuros"), _
                              Dados11("Correspondencia"), Dados11("LimiteCredito"), Dados11("VolumeVendam2"), Dados11("PrecoEstm2"), _
                              Dados11("PrecoEstTon"), Dados11("LimiteCredSug"), Dados11("PrazoPag"), Dados11("DtVenda"), Dados11("Seguimento"), _
                              Dados11("ResponsavelAprovacao"), Dados11("Natureza"), Dados11("Tipo"), Dados11("TipoFrete"), Dados11("CodigoVendedorPro"), _
                              Dados11("CodigoClasProd"), Dados11("DiasViagem"), Dados11("DiasEntregaMaior"), Dados11("DiasEntregaMenor"), Dados11("RuaEntPro"), Dados11("NumeroEntPro"), _
                              Dados11("BairroEntPro"), Dados11("CepEntPro"), Dados11("CidadeEntPro"), Dados11("EstadoEntPro"), Dados11("RuaCobPro"), _
                              Dados11("NumeroCobPro"), Dados11("BairroCobPro"), Dados11("CepCobPro"), Dados11("CidadeCobPro"), Dados11("EstadoCobPro"), _
                              Dados11("ContatoCompras"), Dados11("Departamento"), Dados11("Email"), Dados11("PrecoPadrao"), Dados11("TransacaoPedidos"), _
                              Dados11("AtividadePincipal"), Dados11("GrupoEmpresa"), Dados11("AreaVendas"), Dados11("NomeRepresentante"), Dados11("EmailRepresentante"), _
                              Dados11("TelefoneRep"), Dados11("SegmentacaoMercado"), Dados11("AvisaEmbarque"), Dados11("ContEmbalagens"), Dados11("RegraEmbarque"), _
                              Dados11("DiasTransporte"), Dados11("TolEntregaMenos"), Dados11("TolEntregaMais"), Dados11("JanTolMenos"), Dados11("JanTolMais"), _
                              Dados11("EmiteCertificado"), Dados11("Frete"), Dados11("DesprezarTarifa"), Dados11("DistRomaneio"), Dados11("ObsEmissor"), _
                              Dados11("ObsCedag"), Dados11("CodSap"), Dados11("CodPro"), Dados11("CodSia"), Dados11("Filial"), Dados11("Tes"), Dados11("Eshop"), _
                              Dados11("DataEmissao"), Dados11("TQtdeMaior"), Dados11("TQtdeMenor"), Dados11("PreCadastro"), Dados11("CodLoja"), Dados11("KmEndereco"), _
                              Dados11("KmEnderecoCobranca"), Dados11("isencaoIPI"), Dados11("IcmsDiferido"), Dados11("percentualDiferimento")
end function

function ChamarGRA()
       VerificaSessaoExpirada()
       SQL = "BEGIN TRANSACTION "

       SQL = UpdateCAD()

       SQL = SQL & "IF @@ERROR <> 0 "
       SQL = SQL &   "ROLLBACK TRANSACTION "
       SQL = SQL & "ELSE "
       SQL = SQL &   "COMMIT TRANSACTION "
       Dados11.Open(SQL)

       response.Write "<script>location.href='index.asp'</script>"
end function

function ChamarEXC()
         VerificaSessaoExpirada()

		      SQL = "SELECT Codigo, "
    SQL = SQL + "GrupoContas, "
    SQL = SQL + "Empresa, "
    SQL = SQL + "RazaoSocial, "
    SQL = SQL + "NomeFantasia, "
    SQL = SQL + "Rua, "
    SQL = SQL + "Numero, "
    SQL = SQL + "Bairro, "
    SQL = SQL + "Cep, "
    SQL = SQL + "Cidade, "
    SQL = SQL + "Estado, "
    SQL = SQL + "Pais, "
    SQL = SQL + "Idioma, "
    SQL = SQL + "Telefone, "
    SQL = SQL + "Fax, "
    SQL = SQL + "EmailXml, "
    SQL = SQL + "EmailPagamento, "
    SQL = SQL + "CodigoVendedor, "
    SQL = SQL + "CodigoHfm, "
    SQL = SQL + "Cpf, "
    SQL = SQL + "Cnpj, "
    SQL = SQL + "InscricaoEstadual, "
    SQL = SQL + "InscricaoMunicipal, "
    SQL = SQL + "Declarante, "
    SQL = SQL + "Contribuinte, "
    SQL = SQL + "NaturezaLegal, "
    SQL = SQL + "TipoIndustria, "
    SQL = SQL + "CodigoSuframa, "
    SQL = SQL + "ContaConciliacao, "
    SQL = SQL + "ContaBanco, "
    SQL = SQL + "ChaveOrdenacao, "
    SQL = SQL + "CondicaoPagamento, "
    SQL = SQL + "MetodoPagamento, "
    SQL = SQL + "CodigoJuros, "
    SQL = SQL + "CodigoAntigo, "
    SQL = SQL + "Correspondencia, "
    SQL = SQL + "LimiteCredito, "
    SQL = SQL + "VolumeVendam2, "
    SQL = SQL + "PrecoEstm2, "
    SQL = SQL + "PrecoEstTon, "
    SQL = SQL + "LimiteCredSug, "
    SQL = SQL + "PrazoPag, "
    SQL = SQL + "DtVenda, "
    SQL = SQL + "Seguimento, "
    SQL = SQL + "ResponsavelAprovacao, "
    SQL = SQL + "Natureza, "
    SQL = SQL + "Tipo, "
    SQL = SQL + "TipoFrete, "
    SQL = SQL + "CodigoVendedorPro, "
    SQL = SQL + "CodigoClasProd, "
    SQL = SQL + "DiasViagem, "
    SQL = SQL + "DiasEntregaMaior, "
    SQL = SQL + "DiasEntregaMenor, "
    SQL = SQL + "RuaEntPro, "
    SQL = SQL + "NumeroEntPro, "
    SQL = SQL + "BairroEntPro, "
    SQL = SQL + "CepEntPro, "
    SQL = SQL + "CidadeEntPro, "
    SQL = SQL + "EstadoEntPro, "
    SQL = SQL + "RuaCobPro, "
    SQL = SQL + "NumeroCobPro, "
    SQL = SQL + "BairroCobPro, "
    SQL = SQL + "CepCobPro, "
    SQL = SQL + "CidadeCobPro, "
    SQL = SQL + "EstadoCobPro, "
    SQL = SQL + "ContatoCompras, "
    SQL = SQL + "Departamento, "
    SQL = SQL + "Email, "
    SQL = SQL + "PrecoPadrao, "
    SQL = SQL + "TransacaoPedidos, "
    SQL = SQL + "AtividadePincipal, "
    SQL = SQL + "GrupoEmpresa, "
    SQL = SQL + "AreaVendas, "
    SQL = SQL + "NomeRepresentante, "
    SQL = SQL + "EmailRepresentante, "
    SQL = SQL + "TelefoneRep, "
    SQL = SQL + "SegmentacaoMercado, "
    SQL = SQL + "AvisaEmbarque, "
    SQL = SQL + "ContEmbalagens, "
    SQL = SQL + "RegraEmbarque, "
    SQL = SQL + "DiasTransporte, "
    SQL = SQL + "TolEntregaMenos, "
    SQL = SQL + "TolEntregaMais, "
    SQL = SQL + "JanTolMenos, "
    SQL = SQL + "JanTolMais, "
    SQL = SQL + "EmiteCertificado, "
    SQL = SQL + "Frete, "
    SQL = SQL + "DesprezarTarifa, "
    SQL = SQL + "DistRomaneio, "
    SQL = SQL + "ObsEmissor, "
    SQL = SQL + "ObsCedag, "
    SQL = SQL + "CodSap, "
    SQL = SQL + "CodPro, "
    SQL = SQL + "CodSia, "
    SQL = SQL + "Filial, "
    SQL = SQL + "Tes, "
    SQL = SQL + "Eshop, "
    SQL = SQL + "DataEmissao, "
    SQL = SQL + "TQtdeMaior, "
    SQL = SQL + "TQtdeMenor, "
    SQL = SQL + "PreCadastro, "
    SQL = SQL + "ISNULL(CodLoja,'') AS CodLoja, "
    SQL = SQL + "ISNULL(KmEndereco,'') AS KmEndereco, "
    SQL = SQL + "ISNULL(KmEnderecoCobranca,'') AS KmEnderecoCobranca, "
    SQL = SQL + "ISNULL(isencaoIPI,'') AS isencaoIPI,  "
    SQL = SQL + "ISNULL(IcmsDiferido,'') AS IcmsDiferido,  "
    SQL = SQL + "ISNULL(percentualDiferimento,0) AS percentualDiferimento  "
    SQL = SQL + "FROM UnificadoCliente "
    SQL = SQL + "WHERE Codigo = '" + Request("cod") + "' "

    Dados11.Open(SQL)

    FormUNI "GRE", "Excluir", Dados11("Codigo"), Dados11("GrupoContas"), Dados11("Empresa"), Dados11("RazaoSocial"), Dados11("NomeFantasia"), Dados11("Rua"), _
                              Dados11("Numero"), Dados11("Bairro"), Dados11("Cep"), Dados11("Cidade"), Dados11("Estado"), Dados11("Pais"), Dados11("Idioma"), _
                              Dados11("Telefone"), Dados11("Fax"), Dados11("EmailXml"), Dados11("EmailPagamento"), Dados11("CodigoHfm"),  _
                              Dados11("Cpf"), Dados11("Cnpj"), Dados11("InscricaoEstadual"), Dados11("InscricaoMunicipal"), Dados11("Declarante"), Dados11("Contribuinte"),  _
                              Dados11("NaturezaLegal"), Dados11("TipoIndustria"), Dados11("CodigoSuframa"), Dados11("ContaConciliacao"), Dados11("ContaBanco"),  _
                              Dados11("ChaveOrdenacao"), Dados11("CondicaoPagamento"), Dados11("MetodoPagamento"), Dados11("CodigoJuros"),  _
                              Dados11("Correspondencia"), Dados11("LimiteCredito"), Dados11("VolumeVendam2"), Dados11("PrecoEstm2"), Dados11("PrecoEstTon"),  _
                              Dados11("LimiteCredSug"), Dados11("PrazoPag"), Dados11("DtVenda"), Dados11("Seguimento"), Dados11("ResponsavelAprovacao"),  _
                              Dados11("Natureza"), Dados11("Tipo"), Dados11("TipoFrete"), Dados11("CodigoVendedorPro"), Dados11("CodigoClasProd"),  _
                              Dados11("DiasViagem"), Dados11("DiasEntregaMaior"), Dados11("DiasEntregaMenor"), Dados11("RuaEntPro"), Dados11("NumeroEntPro"), Dados11("BairroEntPro"),  _
                              Dados11("CepEntPro"), Dados11("CidadeEntPro"), Dados11("EstadoEntPro"), Dados11("RuaCobPro"), Dados11("NumeroCobPro"),  _
                              Dados11("BairroCobPro"), Dados11("CepCobPro"), Dados11("CidadeCobPro"), Dados11("EstadoCobPro"), Dados11("ContatoCompras"),  _
                              Dados11("Departamento"), Dados11("Email"), Dados11("PrecoPadrao"), Dados11("TransacaoPedidos"), Dados11("AtividadePincipal"),  _
                              Dados11("GrupoEmpresa"), Dados11("AreaVendas"), Dados11("NomeRepresentante"), Dados11("EmailRepresentante"), Dados11("TelefoneRep"), _ 
                              Dados11("SegmentacaoMercado"), Dados11("AvisaEmbarque"), Dados11("ContEmbalagens"), Dados11("RegraEmbarque"), Dados11("DiasTransporte"), _
                              Dados11("TolEntregaMenos"), Dados11("TolEntregaMais"), Dados11("JanTolMenos"), Dados11("JanTolMais"), Dados11("EmiteCertificado"),  _
                              Dados11("Frete"), Dados11("DesprezarTarifa"), Dados11("DistRomaneio"), Dados11("ObsEmissor"), Dados11("ObsCedag"), Dados11("CodSap"),  _
                              Dados11("CodPro"), Dados11("CodSia"), Dados11("Filial"), Dados11("Tes"), Dados11("Eshop"), Dados11("DataEmissao"), Dados11("TQtdeMaior"), _
                              Dados11("TQtdeMenor"), Dados11("PreCadastro"), Dados11("CodLoja"), Dados11("KmEndereco"), Dados11("KmEnderecoCobranca"), Dados11("isencaoIPI"), _
                              Dados11("IcmsDiferido"), Dados11("percentualDiferimento")
end function

function ChamarGRE()
         VerificaSessaoExpirada()

		      SQL = "UPDATE UnificadoCliente SET "
    SQL = SQL &   "Del = 1, "
    SQL = SQL &   "QuemExcluiu = '" & SESSION("USR") & "', "
    SQL = SQL &   "DataExclusao = GETDATE() "
    SQL = SQL & "WHERE Codigo = '" & Request("cod") & "' "
    Dados11.Open(SQL)

    response.write "<script>alert('Exclusão efetuada com sucesso')</script>"
    response.Redirect "index.asp?act=LST"
end function
' ---------------------------------------------------------------------------------------------------------------------------------------------------'
function ChamarVIS(argAcao)
        ' if argAcao <> "VIS" OR argAcao <> "EXC" then
        '   SQL = "SELECT count(*) as qtde, t2.Tipo, isnull(t2.isencaoIPI,0) as isencaoIPI "
        '   SQL=SQL + "FROM UnificadoClienteDocumento t1 "
        '   SQL=SQL + "LEFT JOIN UnificadoCliente t2 ON t2.Codigo = t1.idUnificadoCliente "
        '   SQL=SQL + "Where t1.idUnificadoCliente = " & Request("cod")
        '   SQL=SQL + "GROUP BY t2.Tipo, t2.isencaoIPI"
        '   Dados11.Open(SQL)

        '   quantidadeDeArquivo=2
        '   temAnexos=0
        '   isencaoIPI=0
        '   tipo=""
        '   if not Dados11.EOF then
        '      isencaoIPI=Dados11("isencaoIPI")
        '      temAnexos=Dados11("qtde")
        '      tipo=Dados11("Tipo")
        '   end if
        '   dados11.Close()

        '   if tipo="P" and isencaoIPI=0 then
        '     quantidadeDeArquivo=2
        '   elseif tipo="P" and isencaoIPI>0 then
        '       quantidadeDeArquivo=3
        '   elseif tipo<>"P" and isencaoIPI=0 then
        '       quantidadeDeArquivo=3
        '   elseif tipo<>"P" and isencaoIPI>0 then
        '       quantidadeDeArquivo=4
        '   end if

        '   if temAnexos = 0 then
        '      response.write "<script type=""text/javascript"">"
        '      response.write "alert('Atenção! Há necessidade de ter pelo menos um documento anexo.');"
        '      response.write "window.location.href = ""index.asp?act=LST"";"
        '      response.write "</script>"
        '      response.End
        '   else
        '       if temAnexos < quantidadeDeArquivo then
        '         response.write "<script type=""text/javascript"">"
        '          response.write "alert('Atenção! Há necessidade de ter pelo menos " & cstr(quantidadeDeArquivo) & " ou mais documentos anexos.');"
        '          response.write "window.location.href = ""index.asp?act=LST"";"
        '          response.write "</script>"
        '          response.End        
        '       end if
        '   end if
        ' end if

        SQL = "BEGIN TRANSACTION "
        if Request("formulario") = "s" and Request("NovoStatus") <> "F" then
           SQL = SQL + UpdateCAD()
        Else
           if Request("NovoStatus") = "F" and Request("formulario") = "s" then
              SQL = SQL & "UPDATE UnificadoCliente SET "
              SQL = SQL &   "CodSap = '" & Request("CodSap") & "', "
              SQL = SQL &   "CodPro = '" & Request("CodPro") & "' "
              SQL = SQL & "WHERE "
              SQL = SQL &   "Codigo = '" & Request("cod") & "'; "
           end if
        end if

        if FilialFiltro = "07" OR Request("paramFilial") = "07" then
           SQL = SQL & "UPDATE UnificadoCliente SET StatusCli = '" & Request("NovoStatus") & "', CodLoja = '" & Request("CodLoja") & "' WHERE Codigo = '" & Request("cod") & "'; "
        else
           SQL = SQL & "UPDATE UnificadoCliente SET StatusCli = '" & Request("NovoStatus") & "', CodLoja = '" & Request("CodLoja") & "' WHERE Codigo = '" & Request("cod") & "'; "
        end if

        'Registra a troca de status no histórico
        SQL = SQL & "INSERT INTO UnificadoHistoricoCliente (Codigo, CodigoCad, Obs, HStatus, DataAlt, Hora, Quem) "
        SQL = SQL & "VALUES ((SELECT ISNULL(MAX(CAST(Codigo AS BIGINT)),0)+1 FROM UnificadoHistoricoCliente), '" & Request("cod") & "', '" & Trim(Request("vobs")) & "', "
        SQL = SQL & "'" & Request("NovoStatus") & "', '" & Hoje() & "', '" & Agora() & "', '" & SESSION("USR") & "') "
        SQL = SQL & "if @@ERROR <> 0 "
        SQL = SQL & "ROLLBACK TRANSACTION "
        SQL = SQL & "ELSE "
        SQL = SQL & "COMMIT TRANSACTION"
        Dados11.Open(SQL)

              SQL = "SELECT Codigo, RazaoSocial, NomeFantasia, LimiteCredito "
        SQL = SQL & "FROM UnificadoCliente "
        SQL = SQL & "WHERE Codigo = '" & Request("cod") & "' AND StatusCli = '" & Request("NovoStatus") & "'"
        Dados12.Open(SQL)

        razaoSocial=""
        nomeFantasia=""
        limiteCredito=0
        if not Dados12.EOF then
          razaoSocial=Dados12("RazaoSocial")
          nomeFantasia=Dados12("NomeFantasia")
          limiteCredito=Dados12("LimiteCredito")
          if Request("NovoStatus") = "F" and Request("formulario") = "s" then
             if not AtualizarSapComOCodigoProtheus(Request("CodPro"), Request("cod"), FilialFiltro, Request("CodLoja")) then
                 response.Redirect "index.asp?act=LST"
             end if
          end if

          'CEDAG devolve para o emissor
          if Request("NovoStatus") = "R" then
            SQL = "SELECT LTRIM(RTRIM(CRM_EMAILPRIMARIO)) AS CRM_EMAILPRIMARIO FROM UnificadoCliente, HELPDESK..CRM_DTI " &_
                  "WHERE QuemIncluiuCod = CRM_COD " &_
                  "AND ISNULL(CRM_D_E_L_E_T,'') = '' " &_
                  "AND Codigo = '" & Request("cod") & "' "
            Dados11.Open(SQL)
            Email = iif( not Dados11.EOF, Dados11("CRM_EMAILPRIMARIO"), "" )

            Dados11.Close()
            if Email <> "" then
              assuntoUnificado = "O Cadastro Unificado de Clientes nº " & Request("cod") & " foi devolvido por "&SESSION("NOME")&" (CDAG)"
              mensagemUnificado = "" &_
                "<html>" &_
                  "<head>" &_
                    "<title></title>" &_
                  "</head>" &_
                  "<body style='font-family:Verdana;font-size:12px'>" &_
                    "<p style='text-align:justify'>" &_
                      "O Cadastro Unificado de Clientes nº " & Request("cod") & " foi devolvido por "&SESSION("NOME")&" (CDAG).<br>" &_
                    "</p>" &_
                    "<p style='text-align:justify'>" &_
                      "E-mail enviado automaticamente pelo sistema de automação de processos." &_
                      "<br>" &_
                      "<strong>Smurfit Westrock do Brasil</strong>" &_
                    "</p>" &_
                  "</body>" &_
                "</html>"

              call EnviaEmail (Email, "", "", "", "noreply.brasil@smurfitwestrock.com.br", mensagemUnificado, assuntoUnificado)
            end if
          response.write "<script>alert('Devolvido com sucesso')</script>"
        'Controles Internos cancelou o cadastro
        elseif Request("NovoStatus") = "X" And SESSION("ControlesInternos") then
          SQL = "SELECT " &_
                  "CRM_EMAILPRIMARIO " &_
                "FROM " &_
                  "UnificadoCliente, HELPDESK..CRM_DTI " &_
                "WHERE QuemIncluiuCod = CRM_COD " &_
                  "AND ISNULL(CRM_D_E_L_E_T,'') = '' " &_
                  "AND Codigo = '" & Request("cod") & "' " &_
                "UNION " &_
                "SELECT " &_
                  "CRM_EMAILPRIMARIO " &_
                "FROM " &_
                  "HELPDESK..CRM_DTI " &_
                "WHERE ( ISNULL(NivelCEDAG,0) = 1 OR ISNULL(NivelE2OPEN,0) = 1 ) " &_
                  "AND ISNULL(CRM_D_E_L_E_T,'') = '' "

          Dados11.Open(SQL)
          if not Dados11.EOF then
            Email = ""
            while not Dados11.EOF
              Email = Email & Trim(Dados11("CRM_EMAILPRIMARIO")) & ";"
              Dados11.MoveNext()
            wend
          end if

          Dados11.Close()
          if Email <> "" then
            Email = Left(Email,Len(Email)-1)

            assuntoUnificado = "O Cadastro Unificado de Clientes nº " & Request("cod") & " foi cancelado por "&SESSION("NOME")&" (Controles Internos)"
            mensagemUnificado = "" &_
              "<html>" &_
                "<head>" &_
                  "<title></title>" &_
                "</head>" &_
                "<body style='font-family:Verdana;font-size:12px'>" &_
                  "<p style='text-align:justify'>" &_
                    "O Cadastro Unificado de Clientes nº " & Request("cod") & " foi cancelado por "&SESSION("NOME")&" (Controles Internos)." &_
                    "<br>" &_
                  "</p>" &_
                  "<p style='text-align:justify'>" &_
                    "E-mail enviado automaticamente pelo sistema de automação de processos." &_
                    "<br>" &_
                    "<strong>Smurfit Westrock do Brasil</strong>" &_
                  "</p>" &_
                "</body>" &_
              "</html>"
            CALL EnviaEmail (Email, "", "", "", "noreply.brasil@smurfitwestrock.com.br", mensagemUnificado, assuntoUnificado)
          end if

          response.write "<script>alert('Cancelado com sucesso')</script>"

         'Controladoria cancelou o cadastro
      Elseif Request("NovoStatus") = "X" And SESSION("Controladoria") then
        SQL = "SELECT CRM_EMAILPRIMARIO "
        SQL = SQL + "FROM UnificadoCliente, HELPDESK..CRM_DTI "
        SQL = SQL + "WHERE QuemIncluiuCod = CRM_COD "
        SQL = SQL + "AND ISNULL(CRM_D_E_L_E_T,'') = '' "
        SQL = SQL + "AND Codigo = '" & Request("cod") & "' "
        SQL = SQL + "UNION "
        SQL = SQL + "SELECT "
        SQL = SQL + "CRM_EMAILPRIMARIO "
        SQL = SQL + "FROM HELPDESK..CRM_DTI "
        SQL = SQL + "WHERE ( ISNULL(NivelCEDAG,0) = 1 OR ISNULL(NivelE2OPEN,0) = 1 OR ISNULL(NivelControlesInternos,0) = 1 ) "
        SQL = SQL + "AND ISNULL(CRM_D_E_L_E_T,'') = '' "

        Dados11.Open(SQL)
        if not Dados11.EOF then
          Email = ""
          while not Dados11.EOF
                Email = Email & Trim(Dados11("CRM_EMAILPRIMARIO")) & ";"
                Dados11.MoveNext()
          wend
        end if

        Dados11.Close()
        if Email <> "" then
          Email = Left(Email,Len(Email)-1)

          assuntoUnificado = "O Cadastro Unificado de Clientes nº " & Request("cod") & " foi cancelado por "&SESSION("NOME")&" (Controladoria)"
          mensagemUnificado = "" &_
            "<html>" &_
              "<head>" &_
                "<title></title>" &_
              "</head>" &_
              "<body style='font-family:Verdana;font-size:12px'>" &_
                "<p style='text-align:justify'>" &_
                  "O Cadastro Unificado de Clientes nº " & Request("cod") & " foi cancelado por "&SESSION("NOME")&" (Controladoria)." &_
                  "<br>" &_
                "</p>" &_
                "<p style='text-align:justify'>" &_
                  "E-mail enviado automaticamente pelo sistema de automação de processos." &_
                  "<br>" &_
                  "<strong>Smurfit Westrock do Brasil</strong>" &_
                "</p>" &_
              "</body>" &_
            "</html>"

          CALL EnviaEmail (Email, "", "", "", "noreply.brasil@smurfitwestrock.com.br", mensagemUnificado, assuntoUnificado)
        end if
        response.write "<script>alert('Cancelado com sucesso')</script>"
      Else
        'Cadastro vistado, aprovado ou finalizado
        SQL = "SELECT CRM_EMAILPRIMARIO, "
        SQL = SQL + "'' EmailXml, "
        SQL = SQL + "'' EmailPagamento "
        SQL = SQL + "FROM UnificadoCliente, HELPDESK..CRM_DTI "
        SQL = SQL + "WHERE QuemIncluiuCod = CRM_COD "
        SQL = SQL + "AND ISNULL(CRM_D_E_L_E_T,'') = '' "
        SQL = SQL + "AND Codigo = '" & Request("cod") & "' "
        SQL = SQL + "UNION "

        if Request("NovoStatus") = "C" then
          if Request("Filial") = "05" then
             SQL = SQL & "SELECT CRM_EMAILPRIMARIO, "
             SQL = SQL & "'' EmailXml, "
             SQL = SQL & "'' EmailPagamento "
             SQL = SQL & "FROM HELPDESK..CRM_DTI "
             SQL = SQL & "WHERE EMP_CODIGO IN ('05','06') "
             SQL = SQL & "AND ISNULL(NivelControladoria,0) = 1 "
             SQL = SQL & "AND ISNULL(CRM_D_E_L_E_T,'') = '' "
          else
             SQL = SQL & ""
             SQL = SQL & "SELECT CRM_EMAILPRIMARIO, "
             SQL = SQL & "'' EmailXml, "
             SQL = SQL & "'' EmailPagamento "
             SQL = SQL & "FROM HELPDESK..CRM_DTI "
             SQL = SQL & "WHERE EMP_CODIGO = '" & Request("Filial") & "' "
             SQL = SQL & "AND ISNULL(NivelControladoria,0) = 1 "
             SQL = SQL & "AND ISNULL(CRM_D_E_L_E_T,'') = '' "
          end if
        Else
           SQL = SQL &  "SELECT " &_
                          "CRM_EMAILPRIMARIO, " &_
                          "'' EmailXml, " &_
                          "'' EmailPagamento " &_
                        "FROM HELPDESK..CRM_DTI " &_
                        "WHERE (ISNULL(NivelE2OPEN,0) = 1 OR ISNULL(NivelControlesInternos,0) = 1 ) " &_
                          "AND ISNULL(CRM_D_E_L_E_T,'') = '' " &_
                        "UNION " &_
                        "SELECT " &_
                          "'' CRM_EMAILPRIMARIO, " &_
                          "EmailXml, " &_
                          "EmailPagamento " &_
                        "FROM UnificadoCliente " &_
                        "WHERE Codigo = '" & Request("cod") & "' "
        end if
        Dados11.Open(SQL)
        if not Dados11.EOF then
          Email = ""
          Email1 = ""
          Email2 = ""

          while not Dados11.EOF
              if Dados11("CRM_EMAILPRIMARIO")<>"" then
                  Email = Email & Trim(Dados11("CRM_EMAILPRIMARIO")) & ";"
              end if

              if Dados11("EmailXml")<>"" then
                 Email1 = Email1 & Trim(Dados11("EmailXml")) & ";"
              end if

              if Dados11("EmailPagamento")<>"" then
                 Email2 = Email2 & Trim(Dados11("EmailPagamento")) & ";"
              end if

              Dados11.MoveNext()
          wend
        end if
        Dados11.Close()

        if Request("NovoStatus") = "C" then 
          nomevisto = "Emissor"
          texto = "enviado"
        Elseif SESSION("Controladoria") then
          if Request("NovoStatus") = "E" then
            nomevisto = "Controladoria"
            texto = "vistado"
          end if
        Elseif SESSION("CEDAG") or SESSION("E2OPEN") then
          if Request("NovoStatus") = "I" then
            texto = "enviado como não conforme"
            nomevisto = "E2OPEN/CDAG"
          Else
            texto = "vistado"
          end if
        Elseif SESSION("ControlesInternos") then
          if Request("NovoStatus") = "E" then
            nomevisto = "Controles Internos"
            texto = "vistado"
          end if
        end if
        
        if Email <> "" and Request("NovoStatus") <> "F" then
          Email = Left(Email,Len(Email)-1)
          
          assuntoUnificado = "O Cadastro Unificado de Clientes nº " & Request("cod") & " foi "& texto &" por "&SESSION("NOME")&" (" & nomevisto & ")"
          mensagemUnificado = "" &_
            "<html>" &_
              "<head>" &_
                "<title></title>" &_
              "</head>" &_
              "<body style='font-family:Verdana;font-size:12px'>" &_
                "<p style='text-align:justify'>" &_
                  "O Cadastro Unificado de Clientes nº " & Request("cod") & " foi "& texto &" por "&SESSION("NOME")&" (" & nomevisto & ")." &_
                  "<br>" &_
                "</p>" &_
                "<p style='text-align:justify'>" &_
                  "E-mail enviado automaticamente pelo sistema de automação de processos." &_
                  "<br>" &_
                  "<strong>Smurfit Westrock do Brasil</strong>" &_
                "</p>" &_
              "</body>" &_
            "</html>"

            CALL EnviaEmail (Email, "", "", "", "noreply.brasil@smurfitwestrock.com.br", mensagemUnificado, assuntoUnificado)
        end if

        EmailCli = ""
        ' FINALIZADO ---------------------------------------------------------------------------------------------'
        if Request("NovoStatus") = "F" and Request("eshop") <> "1" then

          EmailCli = Email1 & Email2
          if right(EmailCli,1)=";" then
             EmailCli = Left(EmailCli,Len(EmailCli)-1)
          end if
          
          assuntoUnificado = "Bem Vindo!"
          mensagemUnificado = "" &_
            "<html>" &_
              "<head>" &_
                "<title></title>" &_
              "</head>" &_
              "<body style='font-family:Verdana;font-size:12px' class='text-center'>" &_
                "<p align='center'>" &_
                  "<img src='http://br-pir1-m010/imagens/onboardingNew.jpg' usemap='#onboard' class='img-fluid'>" &_
                "</p>" &_
                "<br>" &_
                "<p style='text-align:center; font-weight:bold; font-size:12px; color:#00a0e2;'>" &_
                  "Queremos saber a sua opni&atilde;o. <a href='https://survey.binds.co/link/XWdGMRknYNv0AzYGG8PDR0joh26J81bdgpOLx9Jl5wPoVB46'>Clique Aqui!</a>" &_
                "</p>" &_
                "<p style='text-align:left' >" &_
                  "E-mail enviado automaticamente pelo sistema de automação de processos." &_
                  "<br>" &_
                  "<strong>Smurfit Westrock do Brasil</strong>" &_
                "</p>" &_
                "<map name='onboard'>" &_
                  "<area shape='rect' coords='58,1095,747,1178' href='https://survey.binds.co/link/XWdGMRknYNv0AzYGG8PDR0joh26J81bdgpOLx9Jl5wPoVB46' alt='pesquisa' />" &_
                  "<area shape='circle' coords='73,1234,20' href='https://www.facebook.com/Smurfit-Kappa-Brasil-111455303638153/' alt='facebook'/>" &_
                  "<area shape='circle' coords='128,1234,20' href='https://instagram.com/smurfitkappabrasil' alt='insta'/>" &_
                  "<area shape='circle' coords='188,1234,20' href='https://www.linkedin.com/company/smurfit-kappa-group' alt='in'/>" &_
                  "<area shape='circle' coords='240,1234,20' href='https://www.youtube.com/user/SmurfitKappaGroup' alt='youtube'/>" &_
                "</map>" &_
              "</body>" &_
            "</html>"

            CALL EnviaEmail (EmailCli, "", "", "", "noreply.brasil@smurfitwestrock.com.br", mensagemUnificado, assuntoUnificado)
        end if
        
        if Request("NovoStatus") = "F" then
          'ENVIO DE E-MAIL - CADASTRO DE CLIENTE FINALIZADO -----------------------------------------------------------------------------------------
          assuntoUnificado = "CADASTRO DO CLIENTE [ " & Request("cod") & " ] FINALIZADO"
          mensagemUnificado = "<html>" &_
                              "<head>" &_
                                "<title></title>" &_
                              "</head>" &_
                              "<body style='font-family:Verdana;font-size:12px' class='text-center'>" &_
                                "<p style='text-align:center; font-weight:bold; font-size:12px;'>" &_
                                  "O cadastro do cliente: <br>" &_
                                  "<br/><br/>" &_
                                  Request("cod") & " - " & razaoSocial & " ( " & nomeFantansia & " ) " &_
                                  "<br/><br/>" &_
                                  "FOI FINALIZADO!" &_
                                  "<br/><br/>" &_
                                "</p>" &_
                                "<p style='text-align:left' >" &_
                                  "E-mail enviado automaticamente pelo sistema de automação de processos." &_
                                  "<br/>" &_
                                  "<strong>Smurfit Westrock do Brasil</strong>" &_
                                "</p>" &_
                              "</body>" &_
                            "</html>"

          eTO = EmailQuemIncluiuCLI( Request("cod") ) 
          CALL EnviaEmail (eTO, "", "", "", "", mensagemUnificado, assuntoUnificado)
          if limiteCredito > 0 then
              assuntoUnificado = "CADASTRO DO CLIENTE [ " & Request("cod") & " ] FINALIZADO"
              mensagemUnificado = "" &_
                "<html>" &_
                  "<head>" &_
                    "<title></title>" &_
                  "</head>" &_
                  "<body style='font-family:Verdana;font-size:12px' class='text-center'>" &_
                    "<p style='text-align:center; font-weight:bold; font-size:12px;'>" &_
                      "O cadastro do cliente: <br>" &_
                      "<br/>" &_
                      "<br/>" &_
                      Request("cod") & " - " & razaoSocial & " ( " & nomeFantasia & " ) " &_
                      "<br/>" &_
                      "<br/>" &_
                      "PENDENTE DE APROVAÇÃO - LIMITE DE CREDITO!" &_
                      "<br/>" &_
                      "<br/>" &_
                    "</p>" &_
                    "<p style='text-align:left' >" &_
                      "E-mail enviado automaticamente pelo sistema de automação de processos." &_
                      "<br/>" &_
                      "<strong>Smurfit Westrock do Brasil</strong>" &_
                    "</p>" &_
                  "</body>" &_
                "</html>"

              eTO = EmailFinanceiro( Request("cod"))
              if eTO<>"" then
                CALL EnviaEmail (eTO, "", "", "", "", mensagemUnificado, assuntoUnificado)
              end if
          end if
          response.write "<script>alert('Finalizado com sucesso')</script>"
        Else
          response.write "<script>alert('Vistado com sucesso')</script>"
        end if
      end if
    Else
      response.write "<script>alert('ATENÇÃO!! Ocorreu um erro e não foi possível vistar')</script>"
    end if
    Dados12.Close()
    response.Write "<script>location.href='index.asp'</script>"
end function
'----------------------------------------------------------------------------------------------------------------------------------------------------'
function EmailQuemIncluiuCLI( cod )
  SQL = "SELECT CRM_EMAILPRIMARIO " &_
          "FROM HELPDESK..CRM_DTI " &_  
	             "LEFT JOIN UnificadoCliente ON CRM_R_E_C_N_O = QuemIncluiuCod " &_
         "WHERE ISNULL(CRM_D_E_L_E_T,'') = '' " &_
           "AND Codigo =  '" & cod & "' "

  Dados13.Open(SQL)
  EmailQuemIncluiuCLI = iif( (not Dados13.EOF), Trim(Dados13("CRM_EMAILPRIMARIO")), "" )
  Dados13.close()
end function

function EmailFinanceiro( cod )
  SQL = "SELECT ltrim(rtrim(CRM_EMAILPRIMARIO)) as CRM_EMAILPRIMARIO FROM CRM_DTI " &_
         "WHERE ISNULL(CRM_D_E_L_E_T,'') = '' " &_
           "AND ISNULL(NivelFinanceiro,0) = 1 " 

  todosOsEmail=""
  Dados23.Open(SQL)
  Do While not Dados23.EOF
     if Trim(Dados23("CRM_EMAILPRIMARIO")) <> "" then
        todosOsEmail = todosOsEmail + Trim(Dados23("CRM_EMAILPRIMARIO")) + "; "
     end if   
     Dados23.MoveNext
  loop
  Dados23.close()
  EmailFinanceiro = left(todosOsEmail, len(todosOsEmail)-1)
end function

function AtualizarSapComOCodigoProtheus(CodigoProtheus, Codigo, Planta, Loja)
      dim SQL

 On Error Resume Next 
      if CodigoProtheus="" then
         AtualizarSapComOCodigoProtheus = false
         Exit Function
      elseif Codigo="" then
            AtualizarSapComOCodigoProtheus = false
            Exit Function
      elseif Planta="" then
            AtualizarSapComOCodigoProtheus = false
            Exit Function
      end if  

      codSaq=Codigo
      if Planta="07" then
         CodigoProtheus=trim(CodigoProtheus)+trim(Loja)
         codSaq=trim(Codigo)+trim(Loja)
      end if

      SQL="UPDATE SKFT1 " 
      SQL = SQL + " Set FT1_PRECADASTRO = '', FT1_B5CLI = '" + CodigoProtheus + "', FT1_OBSERV = FT1_OBSERV + ' ALTERADO PROSPECT DE: " + Codigo + ", PARA: " + CodigoProtheus + "'" 
      SQL = SQL + " WHERE FT1_PRECADASTRO='P' AND FT1_B5CLI = '" + codSaq + "' AND FT1_FILIAL = '" + Planta + "'"
      select case Planta 
          case "00":set rstAlterar = Conecta( "PIR-CORPORATIVO",SQL )
          case "04":set rstAlterar = Conecta( "PIR-CORPORATIVO",SQL )
          case "05":set rstAlterar = Conecta( "BEN-CORPORATIVO",SQL )
          case "06":set rstAlterar = Conecta( "FOR-CORPORATIVO",SQL )
          case "07":set rstAlterar = Conecta( "SAQ-CORPORATIVO",SQL )
          case else set rstAlterar = Conecta( "FOR-CORPORATIVO",SQL )
      end Select   

      if not isObject(rstAlterar) then
         AtualizarSapComOCodigoProtheus = false
         response.end
      end if

      If Err.Number <> 0 Then
        AtualizarSapComOCodigoProtheus = false
        Err.Clear
        Exit Function
     End If      

   AtualizarSapComOCodigoProtheus = true       
end function
%>