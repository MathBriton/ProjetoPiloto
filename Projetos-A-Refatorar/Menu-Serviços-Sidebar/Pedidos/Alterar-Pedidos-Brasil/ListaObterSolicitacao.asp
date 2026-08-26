<!--#include file="../../cnn2/cnn2.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->
<!--#include file="../../includes/variaveis.asp" -->
<%
session.lcid = 1046
dadosDaQuery = Request.QueryString("selecao")
campoUsuario = campoUsuario & Session("CODUSER")

SQL = "" &_
"SELECT TOP 1000" &_
  "p.Id, " &_
  "p.Pedido, " &_
  "p.Planta, " &_
  "p.CodCliente, " &_
  "p.NomeCliente, " &_
  "p.IdOperacao, " &_
  "p.Operacao, " &_
  "CASE p.IdOperacao WHEN 5 THEN convert(Varchar,p.De,103) " &_
  "ELSE p.De END AS De, " &_
  "CASE p.IdOperacao " &_
    "WHEN 9 THEN " &_
      "CASE p.Para " &_
        "WHEN 'N' THEN 'Normal' " &_
        "WHEN 'C' THEN 'Cancelado' " &_
      "ELSE 'Atendido' END " &_
    "WHEN 5 THEN convert(Varchar,p.Para,103) " &_
    "WHEN 6 THEN " &_
      "(" &_
          "SELECT " &_
            "descricao COLLATE SQL_Latin1_General_CP437_BIN " &_
          "FROM SOLPEN.dbo.interface_inpanet_condicao_pagamento pa " &_
          "WHERE pa.codigo = convert(int,ltrim(rtrim(p.Para)))" &_
      ") " &_
  "ELSE p.Para " &_
  "END AS Para,	" &_
  "p.Observacao, " &_
  "p.Solicitante, " &_
  "convert(varchar,cast(p.Data as Date),103) as Data, " &_
  "h.Status, " &_
  "hd.CRM_NOME as NomeSolicitante, " &_
  "CASE isnull(h.Status,'') " &_
    "WHEN '' THEN 'PENDENTE'  " &_
    "WHEN 1 THEN 'REALIZADA' " &_
  "ELSE 'DESCARTADA' " &_
  "END AS [Status] " &_
"FROM HistoricoSolicitacaoAlteracaoPedido p " &_
  "LEFT JOIN HistoricoAlteracaoPedido h on h.Id = p.Id " &_
  "LEFT JOIN [BRPIR1M001].[HELPDESK].[dbo].[CRM_DTI] hd on hd.CRM_R_E_C_N_O = p.Solicitante " &_
"WHERE Planta = '05' " &_
  iif( Session("Nivel") = "1", "AND Solicitante = '" & SESSION("CODIGO") & "' ", "" ) &_
  iif( dadosDaQuery = "Solicitar", "AND h.Status IS NULL ", "" ) &_
  iif( dadosDaQuery = "Realizada", "AND h.Status = 1 ", "" ) &_
  iif( dadosDaQuery = "Descartadas", "AND h.Status = 2 ", "" ) &_
"ORDER BY p.Data DESC "    

'response.write SQL
'response.end

SET BEN_CORPORATIVO = Conecta( "BEN-CORPORATIVO", SQL )

if not isObject(BEN_CORPORATIVO) or (BEN_CORPORATIVO IS NOTHING) then
  response.end
end if

dgvGrade = "" &_
  "<div class='table-responsive tableFixHead'>" &_
    "<table class='table table-stripedtable-bordered table-sm table-responsive-sm table-hover table-striped'>" &_
      "<thead>" &_
        "<tr>" &_
          "<th class='text-center'>Pedido</th>" &_
          "<th class='text-center'>Planta</th>" &_
          "<th class='text-center'>Cliente</th>" &_
          "<th class='text-center'>Nome</th>" &_
          "<th class='text-center'>Tipo</th>" &_
          "<th class='text-center'>Operação</th>" &_
          "<th class='text-center'>De</th>" &_
          "<th class='text-center'>Para</th>" &_
          "<th class='text-center'>Observação</th>" &_
          "<th class='text-center'>Solicitante</th>" &_
          "<th class='text-center'>Data</th>" &_
          "<th class='text-center'>Status</th>" &_
          "<th class='text-center'>Operação</th>" &_
          "<th class='text-center'><p>  </p></th>" &_
        "</tr>" &_
      "</thead>" &_
      "<tbody>"

SQL = "" &_
  "SELECT " &_
    "b.CRM_COD, hp.CRM_ALTERAPO " &_
  "FROM LINKS a " &_
    "INNER JOIN MONTAMENU b on a.LIN_COD = b.LIN_COD " &_
    "INNER JOIN CRM_DTI hp ON hp.CRM_COD = b.CRM_COD AND ISNULL(hp.CRM_D_E_L_E_T,'') = '' " &_
  "WHERE LIN_URL = '../comercial/bento/ListaSolicitacaoAlteracaoPedido' " &_
    "AND ISNULL(b.MON_D_E_L_E_T,'') = '' " &_
    "AND ISNULL(a.LIN_D_E_L_E_T,'') = '' " &_
    "AND b.CRM_COD = '" & CODUSER & "' " &_
    "AND b.FILIAL = '" & UNIDADE_FILIAL & "' "
 
SET PIR_HELPDESK = Conecta( "PIR-HELPDESK", SQL )

if not isObject(PIR_HELPDESK) or (PIR_HELPDESK IS NOTHING) then
  response.end
end if

ALTERAPO = PIR_HELPDESK("CRM_ALTERAPO")

alteraPedido = iif( not PIR_HELPDESK.EOF, 1, 0 )

ALTERAPO = "S"

PIR_HELPDESK.Close()
SET PIR_HELPDESK = NOTHING

while not BEN_CORPORATIVO.Eof 

  dgvGrade = dgvGrade & "" &_
    "<tr>" &_
      "<td class='text-end' style='display:none'>" & BEN_CORPORATIVO("Id") & "</td>" &_
      "<td class='text-end'>" & BEN_CORPORATIVO("Pedido") & "</td>" &_
      "<td class='text-center'>" & BEN_CORPORATIVO("Planta") & "</td>" &_
      "<td class='text-end'>" & BEN_CORPORATIVO("CodCliente") & "</td>" &_
      "<td class='text-right'>" & BEN_CORPORATIVO("NomeCliente") & "</td>" &_
      "<td class='text-end'>" & BEN_CORPORATIVO("IdOperacao") & "</td>" &_
      "<td class='text-right'>" & BEN_CORPORATIVO("Operacao") & "</td>" &_
      "<td class='text-end'>" & BEN_CORPORATIVO("De") & "</td>" &_
      "<td class='text-end'>" & BEN_CORPORATIVO("Para") & "</td>" &_
      "<td class='text-right'>" & BEN_CORPORATIVO("Observacao") & "</td>" &_
      "<td class='text-right'>" & BEN_CORPORATIVO("NomeSolicitante") & "</td>" &_
      "<td class='text-center'>" & BEN_CORPORATIVO("Data") & "</td>" &_
      "<td class='text-right'>" & BEN_CORPORATIVO("Status") & "</td>"

  if dadosDaQuery <> "Descartadas" AND dadosDaQuery <> "Realizada" and ALTERAPO = "S" then

    dgvGrade = dgvGrade & "" &_
      "<td>" &_
        "<a href='ConfirmaSolicitacao.asp?tipo=" & BEN_CORPORATIVO("Id") & "' class='btn btn-success' alt='Efetivar Solicitação' title='Efetivar Solicitação' role='button'>" &_
          "<i class='bi bi-pencil-fill'></i>Efetivar" &_
        "</a>" &_
      "</td>"

  else

    dgvGrade = dgvGrade & "<td>&nbsp;</td>"

  end if

  if dadosDaQuery <> "Descartadas" AND dadosDaQuery <> "Realizada" and ALTERAPO = "S" then

    dgvGrade = dgvGrade & "" &_
      "<td>" &_
        "<input type='button' class='btn btn-secondary' id='btnExcluir' name='btnExcluir' value='Descartar' onclick='DescartarSolicitacao("& BEN_CORPORATIVO("Id") & ")' alt='Descartar Solicitação' title='Descartar Solicitação' role='button'>" &_
      "</td>" &_
    "</tr>"

  else

    dgvGrade = dgvGrade & "" &_
      "<td>" &_
    "</td>"

  end if

  BEN_CORPORATIVO.MoveNext()

wend

BEN_CORPORATIVO.Close()
SET BEN_CORPORATIVO = NOTHING

dgvGrade = dgvGrade + "</tbody>"

response.write dgvGrade
response.end
%>
