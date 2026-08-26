<%
Session.LCID = 1046
response.Buffer=true
response.AddHeader "cache-control", "no-cache, no-store, must-revalidate"
response.AddHeader "pragma", "no-cache"
response.ExpiresAbsolute = #January 1, 1990 00:00:01#
response.Expires=Now()-1
response.AddHeader "Cache-Control", "must-revalidate"
response.AddHeader "Cache-Control", "no-cache"
response.Addheader "Last-Modified:", Now()

Session.LCID = 1046
Response.ContentType = "text/html; charset=utf-8"
Response.AddHeader "Content-Type", "text/html; charset=utf-8"
%>
<!--#include file = "../Classes/clsAspUtil.asp"-->
<!--#include file = "../../../../../cnn2/cnn2.asp"-->
<!--#include file = "../../../../../includes/variaveis.asp" -->
<!--#include file = "../../../../../funcoes/funcoes.asp" -->
<%
'<----- Recuperando os dados enviados pro querystring da pagina ajax
DIM acao, tipo, pedido, de, para, obs, para2
DIM a, Table, Estilo

filtro = request("filtro")
acao   = request("acao")
tipo   = request("tipo")
pedido = request("pedido")
de     = request("de")
para   = request("para")
obs    = request("obs")
para2  = request("para2")
evidenciasJSON  = request("evidencias")
item   = "01"

'response.write para & "<br>"
'response.end 'filtro & "<br>"

'<----- Validando os campos do formulario
'<----- Instanciando um objeto da classe clsAspUtil, funcao Datas

Dim objUtil
Set objUtil = New AspUtil

select case acao
case "cboTipo"                       call ListaTipos()
case "DePara"                        call MontaDePara()
case "Grava"                         call Grava()
case "BuscaSolicitacoes"             call BuscaSolicitacoes()
case "BuscaSolicitacoesRealizadas"   call BuscaSolicitacoesRealizadas()
case "BuscaSolicitacoesDescartadas"  call BuscaSolicitacoesDescartadas()
case "Exclui"                        call Exclui()
end select

' Monta o combobox dos tipos de alteracao quando a modalbox é carregada -------------------------------------------------------------------------------------------------------------------------------
sub ListaTipos()

  dim Body

  SQL = "EXEC dbo.SP_INPANET_HISTORICO_LISTA_OPERACOES_ALTERACOES_PEDIDO 14"

  set PIR_HELPDESK = Conecta( "PIR-HELPDESK", SQL )

  if PIR_HELPDESK.Eof then

    response.write 1 & "||Não foram encontrados tipos de alterações."
    
  else

    selTipo = "<select name='cboTipo' id='cboTipo' onChange=CallAjax('DePara'); class='form-control' >" &_
                "<option value='0'>Selecione o tipo...</option>"

    while Not PIR_HELPDESK.Eof

      selTipo = selTipo & "<option value=" & PIR_HELPDESK("Id") & ">" & Server.HTMLEncode(PIR_HELPDESK("Operacao")) & "</option>"
      PIR_HELPDESK.MoveNext

    wend

    selTipo = selTipo + "</select>"

    response.write 0 & "||" & selTipo & "||cboTipo||"

  end if

  PIR_HELPDESK.Close()
  set PIR_HELPDESK = nothing

end sub

'Recebe o tipo e o pedido e monta corretamente os campos de e para
Sub MontaDePara()

' DESMARCAR PARA A PRODUÇÃO

  ' if Session("Nivel") > 0 then

  '   Dados.Open( "SELECT C5_NUM FROM SC5010 WHERE C5_NUM = '" & pedido & "' AND C5_VEND1 = '" & Session("CODUSER") & "' AND D_E_L_E_T_ <> '*'" )

  '   if Dados.Eof then
  '     response.write 1 & "||Pedido n o encontrado na base de dados ou o cliente n o pertence   sua carteira."
  '     response.End
  '   end if

  ' end if

	Dim cpoDe, cpoPara
	select case tipo
	case 1 ' Tipo Frete -----------------------------------------------------------------------------

    SQL = "" &_
      "SELECT " &_
        "C5_TPFRETE " &_
      "FROM SC5010 " &_
      "WHERE C5_NUM = '" & pedido & "' " &_
        "AND D_E_L_E_T_ <> '*'"
		
    set PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

		if PIR_DADOSADV.Eof then

      PIR_DADOSADV.Close()
		  response.write 1 & "||Pedido não encontrado na base de dados"
		  response.End

		else

		  cpoDe   = "<input type='text' name='txtDe' id='txtDe' class='form-control' style='width:250px' disabled='disabled' value='" & PIR_DADOSADV("C5_TPFRETE") & "'/>"
		  cpoPara = "<select name='txtPara' id='txtPara' style='width:256px';>" &_
                  "<option value=''>Selecione um tipo de frete...</option>" &_
                  "<option value='C'>C = CIF</option>" &_
                  "<option value='F'>F = FOB</option>" &_
                "</select>"
		  response.write 0 & "||x||DePara||" & cpoDe & "||" & cpoPara
      response.end

		end if	

    PIR_DADOSADV.Close()
    set PIR_DADOSADV = nothing

	case 2 ' Preco ----------------------------------------------------------------------------------

    SQL = "" &_
      "SELECT " &_
        "C6_PRCVEN " &_
      "FROM SC6010 " &_
      "WHERE C6_NUM = '" & pedido & "' " &_
        "AND D_E_L_E_T_ <> '*' " &_
        "AND C6_ITEM = '" & item & "'"
		
    set PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

		if PIR_DADOSADV.Eof then

		  response.write 1 & "||Pedido não encontrado na base de dados"
		  response.End()

		else

		  cpoDe   = "<input type='text' name='txtDe' id='txtDe' class='form-control' style='width:250px' disabled='disabled' value='" & PIR_DADOSADV("C6_PRCVEN") & "'/>"
		  cpoPara = "<input type='text' name='txtPara' id='txtPara' class='form-control' style='width:250px' value=''/>"
		  response.write 0 & "||x||DePara||" & cpoDe & "||" & cpoPara
      response.end

		end if	

    PIR_DADOSADV.Close()
    set PIR_DADOSADV = nothing
    
  case 3 ' Quantidade -----------------------------------------------------------------------------

    SQL = "" &_
      "SELECT " &_
        "C6_QTDVEN " &_
      "FROM SC6010 " &_
      "WHERE C6_NUM = '" & pedido & "' " &_
        "AND D_E_L_E_T_ <> '*' " &_
        "AND C6_ITEM = '" & item & "'" 

    set PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

		if PIR_DADOSADV.EOF then

		  response.write 1 & "||Pedido não encontrado na base de dados"
		  response.End()

		else

		  cpoDe   = "<input type='text' name='txtDe' id='txtDe' class='form-control' style='width:250px' disabled='disabled' value='" & PIR_DADOSADV("C6_QTDVEN") & "'/>"
		  cpoPara = "<input type='text' name='txtPara' id='txtPara' class='form-control' style='width:250px' value=''/>"		  
		  response.write 0 & "||x||DePara||" & cpoDe & "||" & cpoPara
      response.end

		end if

    PIR_DADOSADV.Close()
    set PIR_DADOSADV = nothing

  case 4 ' Quantidade Pronto ----------------------------------------------------------------------

    SQL = "" &_
      "SELECT " &_
        "C6_PRONTO " &_
      "FROM SC6010 " &_
      "WHERE C6_NUM = '" & pedido & "' " &_
        "AND D_E_L_E_T_ <> '*' " &_
        "AND C6_ITEM = '" & item & "' " 
		
    set PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )
     
		if PIR_DADOSADV.Eof then

		  response.write 1 & "||Pedido não encontrado na base de dados"
		  response.End()

		else

		  cpoDe   = "<input type='text' name='txtDe' id='txtDe' class='form-control' style='width:250px' disabled='disabled' value='" & PIR_DADOSADV("C6_PRONTO") & "'/>"
		  cpoPara = "<input type='text' name='txtPara' id='txtPara' class='form-control' style='width:250px' value=''/>"
		  response.write 0 & "||x||DePara||" & cpoDe & "||" & cpoPara
      response.end

		end if	

    PIR_DADOSADV.Close()
    set PIR_DADOSADV = nothing
    		
  case 5 'Data de Entrega -------------------------------------------------------------------------

    'VERIFICANDO SE JA EXISTE ALTERACAO DE DATA PARA ESTE PEDIDO, SE EXISTIR NAO PERMITE NOVA ALTERAÇAO.
		SQL = "" &_
      "SELECT " &_
        "Pedido " &_
      "FROM InHistoricoAlteracoesPedidosSolicitacao " &_
      "WHERE Tipo = 5 " &_
        "AND Pedido = '" & pedido & "' " &_
        "AND Descartado <> 1 " &_
        "AND Realizado = 0"

    set PIR_HELPDESK = Conecta( "PIR-HELPDESK", SQL )

		if not PIR_HELPDESK.EOF then

			response.write 1 & "||Existe alteração de Data de Entrega Cadastrada para este pedido. Não é permitido nova alteração. Entre em contato com o Comercial SK."
			response.End

		end if

    PIR_HELPDESK.Close()
    set PIR_HELPDESK = nothing

    SQL = "" &_
      "SELECT " &_
        "C6_ENTREG, " &_
        "C6_ENTREG2 " &_
      "FROM SC6010 " &_
      "WHERE C6_NUM = '" & pedido & "' " &_
        "AND C6_ITEM = '" & item & "' " &_
        "AND D_E_L_E_T_ <> '*' "

    set PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

		if PIR_DADOSADV.EOF then
      
		  response.write 1 & "||Pedido não encontrado!"
		  response.End

		else

		  cpoDe   = "<input type='text' name='txtDe' id='txtDe' class='form-control' disabled='disabled' value='" & objUtil.AcertaData(PIR_DADOSADV("C6_ENTREG")) & " a " & objUtil.AcertaData(PIR_DADOSADV("C6_ENTREG2")) & "'/>"
      cpoPara = "<div class='input-group-prepend'>" &_
                  "<input type='date' name='txtPara' id='txtPara' class='form-control' value='' required='required' onfocusout='VerificaEntregaNoFinalDeSemana()'/>" &_
                  "&nbsp;&nbsp;<span class='align-middle'>a</span>&nbsp;&nbsp;" &_
                  "<input type='date' name='txtPara2' id='txtPara2' class='form-control' value=''/>" &_
                "</div>"

		  response.write 0 & "||x||DePara||" & cpoDe & "||" & cpoPara
      response.end

		end if

    PIR_DADOSADV.Close()
    set PIR_DADOSADV = nothing

  case 6 ' Condicoes de pagamento ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

    SQL = "SELECT " &_
            "E4_CODIGO, " &_
            "E4_DESCRI " &_
          "FROM ZZ7010 " &_
            "INNER JOIN SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' " &_
            "INNER JOIN SC5010 ON C5_CONDPAG = E4_CODIGO AND SC5010.D_E_L_E_T_ = '' " &_
          "WHERE ZZ7_GRUPO = 'SE4' " &_
            "AND ZZ7010.D_E_L_E_T_ = '' " &_
            "AND E4_LIBVEN <> 'N' " &_
            "AND E4_MSBLQL <> '1' " &_
            "AND E4_CODIGO <> '090' " &_
            "AND C5_NUM = '" & pedido & "' " &_
          "ORDER BY E4_DESCRI"

    set PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )
		
		if PIR_DADOSADV.Eof then

		  response.write 1 & "||Pedido não encontrado na base de dados"
		  response.End()

		else

		  cpoDe = "<input type='text' name='txtDe' id='txtDe' class='form-control' style='width:250px' disabled='disabled' value='" & PIR_DADOSADV("E4_DESCRI") & "'/>"

		  'Busca as condi  es de pagamento
      SQL = "" &_
        "SELECT " &_
          "E4_CODIGO, " &_
          "E4_DESCRI " &_
        "FROM ZZ7010 " &_
          "INNER JOIN SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' " &_
        "WHERE ZZ7_GRUPO = 'SE4' " &_
          "AND ZZ7010.D_E_L_E_T_ = '' " &_
          "AND E4_LIBVEN <> 'N' " &_
          "AND E4_MSBLQL <> '1' " &_
          "AND E4_CODIGO <> '090' " &_
        "ORDER BY E4_DESCRI"

      set PIR_DADOSADV1 = Conecta( "PIR-DADOSADV", SQL )
		  
		  cpoPara = "<select name='txtPara' id='txtPara' class='form-control' style='width:256px';>" &_
                  "<option value=''>Selecione a condição...</option>"

		  while not PIR_DADOSADV1.Eof

        cpoPara = cpoPara & "<option value=" & PIR_DADOSADV1("E4_CODIGO") & ">" & PIR_DADOSADV1("E4_DESCRI") & "</option>"
        PIR_DADOSADV1.MoveNext

      wend

      cpoPara = cpoPara + "</select>"
		  
		  response.write 0 & "||x||DePara||" & cpoDe & "||" & cpoPara
      response.end	

      PIR_DADOSADV1.close()
      set PIR_DADOSADV1 = nothing

		end if	  

    PIR_DADOSADV.close()
    set PIR_DADOSADV = nothing

    response.write para & "<br>"
response.end 'filtro & "<br>"

  case 7 ' Cancelar -------------------------------------------------------------------------------

    'Conecta no banco e resgata o status do pedido para mostrar
    SQL = "" &_
      "SELECT " &_
        "C5_ESTATUS " &_
      "FROM SC5010 " &_
      "WHERE C5_NUM = '" & pedido & "' " &_
        "AND D_E_L_E_T_ <> '*' " 
		
    set PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

		if PIR_DADOSADV.Eof then

		  response.write 1 & "||Pedido não encontrado na base de dados"
		  response.End

		else

		  cpoDe   = "<input type='text' name='txtDe' id='txtDe' class='form-control' style='width:250px' disabled='disabled' value='" & PIR_DADOSADV("C5_ESTATUS") & "'/>"
      cpoPara = "<input type='text' name='txtPara' id='txtPara' class='form-control' style='width:250px' disabled='disabled' value='C'/>"
		  response.write 0 & "||x||DePara||" & cpoDe & "||" & cpoPara
      response.end

		end if

    PIR_DADOSADV.Close()
    set PIR_DADOSADV = nothing

  case 8 ' Pedido do Cliente ----------------------------------------------------------------------

    SQL = "" &_
      "SELECT " &_
        "C6_PEDCLI " &_
      "FROM SC6010 " &_
      "WHERE C6_NUM = '" & pedido & "' " &_
        "AND D_E_L_E_T_ <> '*' " &_
        "AND C6_ITEM = '" & item & "'"
		
    set PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

		if PIR_DADOSADV.Eof then

		  response.write 1 & "||Pedido não encontrado na base de dados"
		  response.End()

		else

		  cpoDe   = "<input type='text' name='txtDe' id='txtDe' class='form-control' style='width:250px' disabled='disabled' value='" & PIR_DADOSADV("C6_PEDCLI") & "'/>"
		  cpoPara = "<input type='text' name='txtPara' id='txtPara' class='form-control' style='width:250px'  maxlength='15' value=''/>"
		  response.write 0 & "||x||DePara||" & cpoDe & "||" & cpoPara
      response.end

		end if	

    PIR_DADOSADV.Close()
    set PIR_DADOSADV = nothing

  case 9 ' Status ---------------------------------------------------------------------------------

    SQL = "" &_
      "SELECT " &_
        "C5_ESTATUS " &_
      "FROM SC5010 " &_
      "WHERE C5_NUM = '" & pedido & "' " &_
        "AND D_E_L_E_T_ <> '*' "

    set PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

		if PIR_DADOSADV.Eof then

		  response.write 1 & "||Pedido não encontrado na base de dados"
		  response.End()

		else

      if PIR_DADOSADV("C5_ESTATUS") <> "S" then

        cpoDe   = "<input type='text' name='txtDe' id='txtDe' class='form-control' style='width:250px' disabled='disabled' value='" & PIR_DADOSADV("C5_ESTATUS") & "'/>"
        cpoPara = "<select name='txtPara' id='txtPara' class='form-control' style='width:256px';>" &_
                    "<option value=''>Selecione um estatus...</option>" &_
                    "<option value='A'>ATIVO</option>" &_
                    "<option value='P'>PROGRAMADO</option>" &_
                    "<option value='C'>CANCELADO</option>" &_
                    "<option value='L'>LIQUIDADO</option>" &_
                  "</select>"
        response.write 0 & "||x||DePara||" & cpoDe & "||" & cpoPara
        response.end

      else

        response.write 1 & "||O Pedido esta SUSPENSO a alteração deste status somente é possível com desbloqueio financeiro do Cliente."

      end if

		end if	

    PIR_DADOSADV.Close()
    set PIR_DADOSADV = nothing

  case 10 ' Filial Fabrica ------------------------------------------------------------------------

    SQL = "SELECT " &_
            "CASE C5_INPUBE WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END C5_INPUBE " &_
          "FROM SC5010 " &_
          "WHERE C5_NUM = '" & pedido & "' AND D_E_L_E_T_ <> '*' " &_
          iif( (session("Nivel") > "0"), "AND C5_VEND1 = '" & Session("CODUSER") & "' ", "" )

    set PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

		if PIR_DADOSADV.Eof then
     
		  response.write 1 & "||Pedido não encontrado na base de dados"
		  response.End

		else

		  cpoDe = "<input type='text' name='txtDe' id='txtDe' class='form-control' style='width:250px' disabled='disabled' value='" & PIR_DADOSADV("C5_INPUBE") & "'/>"
		  cpoPara = "<select name='txtPara' id='txtPara' class='form-control' style='width:256px';>" &_
                  "<option value=''>Selecione uma filial...</option>" &_
		              iif( (PIR_DADOSADV("C5_INPUBE") = "Pirapetinga"), "<option value='U'>Uberaba</option>", "<option value='I'>Pirapetinga</option>" ) &_
                "</select>"		  
		  response.write 0 & "||x||DePara||" & cpoDe & "||" & cpoPara
      response.end

		end if

    PIR_DADOSADV.Close()
    set PIR_DADOSADV = nothing

  case 14

    SQL = "" &_
      "SELECT " &_
        "C5_FRETUNI " &_
      "FROM SC5010 " &_
      "WHERE C5_NUM = '" & pedido & "' " &_
        "AND D_E_L_E_T_ <> '*' " &_
        iif( (session("Nivel") > "0"), "AND C5_VEND1 = '" & Session("CODUSER") & "' ", "" )

    set PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

    if PIR_DADOSADV.Eof then

		  response.write 1 & "||Pedido não encontrado na base de dados"
		  response.End()

		else

      cpoDe = "<input type='text' name='txtDe' id='txtDe' class='form-control' style='width:250px' disabled='disabled' value='" & PIR_DADOSADV("C5_FRETUNI") & "'/>"
		  cpoPara = "<input type='text' name='txtPara' id='txtPara' style='width:250px'/>"
		  response.write 0 & "||x||DePara||" & cpoDe & "||" & cpoPara
      response.end

    end if

    PIR_DADOSADV.Close()
    set PIR_DADOSADV = nothing

	end select

end sub

	
' -------------------------------------------------------------------------------------------------  
Sub Grava()
	Dim Data, Horas, LoginUsu, Msg
	Msg = ""

	if trim(SESSION("USR")) = "" then
		LoginUsu = " "
	else
		LoginUsu = trim(SESSION("USR"))
	end if

	if trim(tipo) = "5" then
		' msgAss = "(1) ALTERA PEDIDO - 1ª DATA DE ENTREGA"
		' msgBod = "<br>" & trim(para) & "<br>" & FormatarDataPTBR( trim(para) )

		' CALL EnviaEmail( "jorge.queiroz@smurfitwestrock.com.br","","","","", msgBod, msgAss  )

		VerEntFDS = VerificaDataDeEntregaNoFinalDeSemana( trim(pedido), trim(para) ) 

		if VerEntFDS <> "" then

		Msg = VerEntFDS & " Solicitação registrada com sucesso."

		else

		Msg = " Solicitação registrada com sucesso."

		end if

		' msgAss = "(2) ALTERA PEDIDO - 1ª DATA DE ENTREGA"
		' msgBod = "<br>" & trim(para) & "<br>" & FormatarDataPTBR( trim(para) )

		' CALL EnviaEmail( "jorge.queiroz@smurfitwestrock.com.br","","","","", msgBod, msgAss  )

		'para = FormatarDataPTBR( trim(para) )
		'para2 = FormatarDataPTBR( trim(para2) )
	else
		Msg = "Solicitação registrada com sucesso."
	end if

	Data = objUtil.InverteData(objUtil.CompletaData(Date()))
	Horas = cstr(Time())

	cmd = "" &_
		"BEGIN TRANSACTION TRANS; " &_
		"INSERT INTO InHistoricoAlteracoesPedidosSolicitacao( " &_
			"Pedido, " &_
			"Tipo, " &_
			"De, " &_
			"Para, " &_
			"Para2, " &_
			"Observacao, " &_
			"Realizado, " &_
			"Descartado, " &_
			"Data, " &_
			"Hora, " &_
			"Login " &_
		") OUTPUT INSERTED.Id VALUES ( " &_
			" '" & trim(pedido) & "', " &_
			" '" & trim(tipo) & "', " &_
			" '" & trim(de) & "', " &_
			" '" & trim(para) & "', " &_
			" '" & trim(para2) & "', " &_
			" '" & trim(obs) & "', " &_
			" 0, " &_
			" 0, " &_
			" '" & trim(Data) & "', " &_ 
			" '" & trim(Horas) & "', " &_ 
			" '" & trim(LoginUsu) & "' " &_
		"); " &_
		"IF @@ERROR <> 0 " &_
			"ROLLBACK TRANSACTION TRANS " &_
		"ELSE " &_
			"COMMIT TRANSACTION TRANS "

	'RESPONSE.WRITE cmd
	'RESPONSE.END

	Dim idSolicitacao
	SET PIR_HELPDESK = Conecta( "PIR-HELPDESK", cmd )
	If Err.Number = 0 AND IsObject(PIR_HELPDESK) Then		
		idSolicitacao = PIR_HELPDESK("Id")
		
		PIR_HELPDESK.Close()
		SET PIR_HELPDESK = NOTHING
	End If
  
	If evidenciasJSON <> "" Then
		Call SalvarEvidencias(idSolicitacao, evidenciasJSON)
	End If
  
	response.write 0 & "||" & Msg & "||Grava||"
End Sub

' -------------------------------------------------------------------------------------------------
sub Exclui()

  LoginUsu = iif( (trim(SESSION("USR")) = ""), " ", trim(SESSION("USR")) )

  SQL = "EXEC ExcluiSolicitacaoAlteracaoPedido @Id = '" & trim(tipo) & "', @LoginUsu = '" & trim(LoginUsu) & "' "
  set PIR_HELPDESK = Conecta( "PIR-HELPDESK", SQL )
  set PIR_HELPDESK = nothing

  response.write 0 & "||Solicitação " & tipo & " excluida da lista com sucesso.||Exclui"

end sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
sub BuscaSolicitacoes()

  if session("Nivel") = 0 then ' NIVEL USUÁRIO ----------------------------------------------------

    SQL = "" &_
      "SELECT " &_
        "a.Id As Id, " &_
        "Pedido As Pedido, " &_
        "A1_NREDUZ As Cliente, " &_
        "b.Id as OperacaoId, " &_
        "Operacao As Operacao, " &_
        "CASE WHEN b.Id = 3 THEN De + ' - ' + REPLACE(cast(cast((DADOSADV.dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(De,'.',''),',','.') as float)) as int) as varchar(10)),'.',',') +'Kg' ELSE De END AS De, " &_
        "Para, " &_
        "Para2, " &_
        "CASE C5_INPUBE WHEN 'U' THEN 'Uberaba' Else 'Pirapetinga' END FILIALDESCRICAO, " &_
        "CASE b.Id " &_
          "WHEN 3 THEN Para + ' - ' + REPLACE(cast(cast((DADOSADV.dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(Para,'.',''),',','.') as float)) as int) as varchar(10)),'.',',') +'Kg'  " &_
          "WHEN 1 THEN CASE Para WHEN 'C' THEN 'CIF' ELSE 'FOB' END " &_
          "WHEN 6 THEN ( SELECT top 1 E4_DESCRI COLLATE Latin1_General_BIN " &_
                        "FROM DADOSADV.dbo.ZZ7010 " &_
                        "INNER JOIN DADOSADV.dbo.SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' " &_
                        "WHERE E4_CODIGO COLLATE Latin1_General_BIN = Para) " &_
          "WHEN 10 THEN CASE Para WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END " &_
        "ELSE Para END AS Descricao, " &_
        "Observacao As Observacao, " &_
        "Realizado As Realizado, " &_
        "Data As Data, " &_
        "Hora As Hora, " &_
        "Login As Login " &_
      "FROM InHistoricoAlteracoesPedidosSolicitacao as a " &_
        "INNER JOIN InHistoricoTipoOperacao as b on a.Tipo = b.Id " &_
        "INNER JOIN DADOSADV.dbo.SC5010 ON DADOSADV.dbo.SC5010.C5_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
        "INNER JOIN DADOSADV.dbo.SC6010 ON DADOSADV.dbo.SC6010.C6_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
        "INNER JOIN DADOSADV.dbo.SA1010 ON DADOSADV.dbo.SA1010.A1_COD = DADOSADV.dbo.SC5010.C5_CLIENTE AND DADOSADV.dbo.SA1010.A1_LOJA = DADOSADV.dbo.SC5010.C5_LOJACLI " &_
      "WHERE " &_
        "Data >= '20230101' " &_
      "ORDER BY Realizado, Data, Hora "

        ' "Data >= '20220101' " &_
        ' "AND Descartado <> 1 " &_
        ' "AND Realizado = 0 " &_
        ' iif( filtro = "CL", "--AND b.Id = 9 AND Para = 'L' ", "--AND Para <> 'L' " ) &_


  else ' NIVEL REPRESENTANTE ----------------------------------------------------------------------

    SQL = "" &_
      "SELECT " &_
        "a.Id As Id, " &_
        "Pedido As Pedido, " &_
        "A1_NREDUZ As Cliente, " &_
        "b.Id as OperacaoId, " &_
        "Operacao As Operacao, " &_
        "CASE WHEN b.Id = 3 THEN De + ' - ' + REPLACE(cast(cast((DADOSADV.dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(De,'.',''),',','.') as float)) as int) as varchar(10)),'.',',') +'Kg' ELSE De ENS AS De, " &_
        "CASE WHEN b.Id = 3 THEN Para + ' - ' + REPLACE(cast(cast((DADOSADV.dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(Para,'.',''),',','.') as float)) as int) as varchar(10)),'.',',') +'Kg' ELSE Para END AS Para, " &_
        "Para2, " &_
        "CASE C5_INPUBE WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END FILIALDESCRICAO, " &_
        "CASE b.Id " &_
          "WHEN 1 THEN CASE Para WHEN 'C' THEN 'CIF' ELSE 'FOB' END " &_
          "WHEN 6 THEN (SELECT top 1 E4_DESCRI COLLATE Latin1_General_BIN FROM DADOSADV.dbo.ZZ7010 INNER JOIN DADOSADV.dbo.SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' WHERE E4_CODIGO COLLATE Latin1_General_BIN = Para) " &_
          "WHEN 10 THEN CASE Para WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END " &_
        "ELSE Para END AS Descricao, " &_
        "Observacao As Observacao, " &_
        "Realizado As Realizado, " &_
        "Data As Data, " &_
        "Hora As Hora, " &_
        "Login As Login " &_
      "FROM InHistoricoAlteracoesPedidosSolicitacao as a " &_
        "INNER JOIN InHistoricoTipoOperacao as b on a.Tipo = b.Id " &_
        "INNER JOIN DADOSADV.dbo.SC5010 ON DADOSADV.dbo.SC5010.C5_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
        "INNER JOIN DADOSADV.dbo.SC6010 ON DADOSADV.dbo.SC6010.C6_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
        "INNER JOIN DADOSADV.dbo.SA1010 ON DADOSADV.dbo.SA1010.A1_COD = DADOSADV.dbo.SC5010.C5_CLIENTE AND DADOSADV.dbo.SA1010.A1_LOJA = DADOSADV.dbo.SC5010.C5_LOJACLI " &_
      "WHERE Descartado <> 1 " &_
        "AND Data >= '20150101' " &_
        "AND Realizado = 0 " &_
        "AND lower(Login) = lower('" & Session("Usr") & "') " &_
        iif( filtro = "SL", "AND b.Id = 9 AND Para = 'L' ", "AND Para <> 'L' " ) &_
      "ORDER BY Realizado, Data, Hora"

  end if

  set PIR_HELPDESK = Conecta( "PIR-HELPDESK", cmd )

  if NOT PIR_HELPDESK.Eof then

    Table = "<table class='table table-sm table-striped table-hover' style='font-size:11px'>" &_
              "<thead>" &_
                "<tr style='font-family:arial; font-size:12px;background-color:#EEEEEE;padding:5px'>" &_
                  "<td align='center' colspan='13'>" &_
                    "<b>Lista de Solicitações de Alteração de Pedidos</b>" &_
                    iif( filtro = "SL", " (sem liquidação)", " (só liquidação)" ) &_
                  "</td>" &_
                "</tr>" &_
                "<tr style='font-family:arial; font-size:12px; background-color:#EEEEEE; padding:5px'>" &_
                  "<td align='center' >&nbsp;</td>" &_
                  "<td align='center' >Pedido</td>" &_
                  "<td align='center' >Planta</td>" &_
                  "<td align='center' >Cliente</td>" &_
                  "<td align='center' >&nbsp;</td>" &_
                  "<td align='center' >Tipo</td>" &_
                  "<td align='center' >De</td>" &_
                  "<td align='center' >Para</td>" &_
                  "<td align='center' >Realizado</td>" &_
                  "<td align='center' >Obs</td>" &_
                  "<td align='center' >Solicitante</td>" &_
                  "<td align='center' >Data</td>" &_
                  "<td align='center' >Hora</td>" &_
                "</tr>" &_
              "</thead>" &_
              "<tbody>"
    a = 1

    while not PIR_HELPDESK.Eof

      if (DateDiff("d", objUtil.AcertaData(PIR_HELPDESK("Data")), Date()) < 4 and PIR_HELPDESK("Realizado") = True) or PIR_HELPDESK("Realizado") = False then

        Table = Table & "<tr>"

        if Session("Nivel") = 0 AND SESSION("PERMISSAO") = "C" then 

          lnk = "AlteraPedido.asp" &_
                  "?Pedido=" & trim(PIR_HELPDESK("Pedido")) &_
                  "&Tipo=" & trim(PIR_HELPDESK("OperacaoId")) &_ 
                  "&De="&trim(PIR_HELPDESK("De")) &_ 
                  "&Para="&trim(PIR_HELPDESK("Para")) &_
                  "&Para2="&trim(PIR_HELPDESK("Para2")) &_ 
                  "&Descricao="&trim(PIR_HELPDESK("Descricao")) &_ 
                  "&Id="&trim(PIR_HELPDESK("Id"))

          Table = Table & "<td align='center'>" &_
                            "<img src='ico_excluir.gif' width='15' height='15' onclick=ExcluiSolicitacao('" & trim(PIR_HELPDESK("Id")) & "'); title='Clique para excluir esta solicitação da lista.' style='cursor:pointer'>" &_
                          "</td>" &_
                          "<td align='center'>" &_
                            "<a href='" & lnk & "'>" & trim(PIR_HELPDESK("Pedido")) & "</a>" &_
                          "</td>"

        else

          Table = Table & "<td align='center'>&nbsp;</td>" &_
                          "<td align='center'>" & trim(PIR_HELPDESK("Pedido")) & "</td>"	  

        end if

        Table = Table & "<td align='center'>" & trim(PIR_HELPDESK("FILIALDESCRICAO")) & "</td>" &_
                        "<td align='center'>" & trim(PIR_HELPDESK("Cliente")) & "</td>" &_
                        "<td align='center'>" & trim(PIR_HELPDESK("OperacaoId")) & "</td>" &_
                        "<td align='left'>"   & trim(PIR_HELPDESK("Operacao")) & "</td>" &_
                        "<td align='center'>"  & trim(PIR_HELPDESK("De")) & "</td>"

        if trim(PIR_HELPDESK("OperacaoId")) = "5" AND trim(PIR_HELPDESK("Para2")) <> "" then          
          Table = Table & "<td align='center'>" & FormatarDataPTBR(trim(PIR_HELPDESK("Para"))) & " à " & FormatarDataPTBR(trim(PIR_HELPDESK("Para2"))) & "</td>"
        else
          Table = Table & "<td align='right'>" & trim(PIR_HELPDESK("Descricao")) & "</td>"
        end if

        if trim(PIR_HELPDESK("Realizado")) then

          Table = Table & "<td align='center'>SIM</td>"

        else

          Table = Table & "<td align='center'>NÃO</td>"

        end if

        Table = Table & "<td align='center'>" & trim(PIR_HELPDESK("Observacao")) & "</td>" &_
                        "<td align='center'>" & trim(PIR_HELPDESK("Login")) & "</td>" &_
                        "<td align='center'>" & objUtil.AcertaData(trim(PIR_HELPDESK("Data"))) & " " & DateDiff("d", objUtil.AcertaData(PIR_HELPDESK("Data")), Date()) & "</td>" &_
                        "<td align='right'>" & trim(PIR_HELPDESK("Hora")) & "</td>" &_
                      "</tr>"
        a = a + 1

      end if

      PIR_HELPDESK.MoveNext

    wend

    Table = Table & "</tbody>" &_
                  "</table>"

'    response.write 0 & "||" & Table & "||BuscaSolicitacoes||"
	  response.write Table

  else

    Table = "<table class='table-responsive table-sm table-striped table-hover'>" &_
              "<tr>" &_
                "<td>" &_
                  "<center>" &_
                    "NÃO HÁ ITENS CADASTRADOS ..." &_
                  "</center>" &_
                "</td>" &_
              "</tr>" &_
            "</table>"

    Lista = 0 & "||" & Table

'    response.write 0 & "||" & Table & "||BuscaSolicitacoes||"
	  response.write Table

  end if

  PIR_HELPDESK.Close()
  set PIR_HELPDESK = nothing

end sub
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
sub BuscaSolicitacoesRealizadas()

  if Session("Nivel") = 0 then ' NIVEL USUARIO ----------------------------------------------------

    SQL = "" &_
      "SELECT " &_
        "a.Id As Id, " &_
        "Pedido As Pedido, " &_
        "A1_NREDUZ As Cliente, " &_
        "b.Id as OperacaoId, " &_
        "Operacao As Operacao, " &_
        "ISNULL(QuemRealizou,'') QuemRealizou, " &_
        "ISNULL(DtRealizou,'') DtRealizou, " &_
        "CASE WHEN b.Id = 3 THEN De + ' - ' + REPLACE(cast(cast((DADOSADV.dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(De,'.',''),',','.') as float)) as int) as varchar(10)),'.',',') +'Kg' ELSE De END AS De, " &_
        "CASE WHEN b.Id = 3 THEN Para + ' - ' + REPLACE(cast(cast((DADOSADV.dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(Para,'.',''),',','.') as float)) as int) as varchar(10)),'.',',') +'Kg' ELSE Para END AS Para, " &_
        "Para2, " &_
        "CASE b.Id	" &_
        "WHEN 1 THEN CASE Para WHEN 'C' THEN 'CIF' ELSE 'FOB' END " &_
        "WHEN 6 THEN ( SELECT top 1 E4_DESCRI COLLATE Latin1_General_BIN " &_
                      "FROM DADOSADV.dbo.ZZ7010 " &_
                        "INNER JOIN DADOSADV.dbo.SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' " &_
                      "WHERE E4_CODIGO COLLATE Latin1_General_BIN = Para) " &_
        "WHEN 10 THEN CASE Para WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END " &_
        "ELSE Para END AS Descricao, " &_
        "Observacao As Observacao, " &_
        "Realizado As Realizado, " &_
        "Data As Data, " &_
        "Hora As Hora, " &_
        "Login As Login " &_
      "FROM InHistoricoAlteracoesPedidosSolicitacao as a " &_
        "INNER JOIN InHistoricoTipoOperacao as b on a.Tipo = b.Id " &_
        "INNER JOIN DADOSADV.dbo.SC5010 ON DADOSADV.dbo.SC5010.C5_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
        "INNER JOIN DADOSADV.dbo.SC6010 ON DADOSADV.dbo.SC6010.C6_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
        "INNER JOIN DADOSADV.dbo.SA1010 ON DADOSADV.dbo.SA1010.A1_COD = DADOSADV.dbo.SC5010.C5_CLIENTE AND DADOSADV.dbo.SA1010.A1_LOJA = DADOSADV.dbo.SC5010.C5_LOJACLI " &_
      "WHERE " &_
        "Data >= '20230101' " &_
      "ORDER BY Realizado, Data, Hora"

      ' "Descartado <> 1 " &_
      '   "AND Data >= convert(varchar(8),dateadd(Day,-720,getdate()),112) " &_
      '   "AND Realizado = 1 " &_
      '   iif( filtro = "CL", "AND b.Id = 9 AND Para = 'L' ", "AND b.Id <> 9 " ) &_

  else ' NIVEL REPRESENTANTE ----------------------------------------------------------------------

    SQL = "" &_
      "SELECT " &_
        "a.Id As Id, " &_
        "Pedido As Pedido, " &_
        "A1_NREDUZ As Cliente, " &_
        "b.Id as OperacaoId, " &_
        "Operacao As Operacao, " &_
        "ISNULL(QuemRealizou,'') QuemRealizou, " &_
        "ISNULL(DtRealizou,'') DtRealizou, " &_
        "CASE WHEN b.Id = 3 then De + ' - ' + REPLACE(cast(cast((DADOSADV.dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(De,'.',''),',','.') as float)) as int) as varchar(10)),'.',',') +'Kg' ELSE De END AS De, " &_
        "CASE WHEN b.Id = 3 then Para + ' - ' + REPLACE(cast(cast((DADOSADV.dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(Para,'.',''),',','.') as float)) as int) as varchar(10)),'.',',') +'Kg' ELSE Para END AS Para, " &_
        "Para2, " &_
        "CASE b.Id " &_
          "WHEN 1 THEN CASE Para WHEN 'C' THEN 'CIF' ELSE 'FOB' END " &_
          "WHEN 6 THEN ( SELECT top 1 E4_DESCRI COLLATE Latin1_General_BIN " &_
                        "FROM DADOSADV.dbo.ZZ7010 " &_
                        "INNER JOIN DADOSADV.dbo.SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' " &_
                        "WHERE E4_CODIGO COLLATE Latin1_General_BIN = Para) " &_
          "WHEN 10 THEN CASE Para WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END " &_
        "ELSE Para END AS Descricao, " &_
        "Observacao As Observacao, " &_
        "Realizado As Realizado, " &_
        "Data As Data, " &_
        "Hora As Hora, " &_
        "Login As Login " &_
      "FROM InHistoricoAlteracoesPedidosSolicitacao as a " &_
        "INNER JOIN InHistoricoTipoOperacao as b on a.Tipo = b.Id " &_
        "INNER JOIN DADOSADV.dbo.SC5010 ON DADOSADV.dbo.SC5010.C5_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
        "INNER JOIN DADOSADV.dbo.SC6010 ON DADOSADV.dbo.SC6010.C6_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
        "INNER JOIN DADOSADV.dbo.SA1010 ON DADOSADV.dbo.SA1010.A1_COD = DADOSADV.dbo.SC5010.C5_CLIENTE AND DADOSADV.dbo.SA1010.A1_LOJA = DADOSADV.dbo.SC5010.C5_LOJACLI " &_
      "WHERE Descartado <> 1 " &_
        "AND Data >= convert(varchar(8),dateadd(Day,-150,getdate()),112) " &_
        "AND Realizado = 1 " &_
        "AND lower(Login) = lower('" & Session("Usr") & "') " &_
        iif( filtro = "CL", "AND b.Id = 9 AND Para = 'L' ", "AND b.Id <> 9 " ) &_
      "ORDER BY Realizado, Data, Hora" 

  end if

' response.write SQL
' response.end

  set PIR_HELPDESK = Conecta( "PIR-HELPDESK", cmd )

  if not PIR_HELPDESK.Eof then

    Table = "" &_
      "<table class='table-responsive table-sm table-striped table-hover'>" &_
        "<thead>" &_
          "<tr style='font-family:arial; font-size:12px;background-color:#EEEEEE;padding:5px'>" &_
            "<th style='font-family:arial; font-size:12px; background-color:#EEEEEE; padding:5px' align='center' colspan=14>" &_
              "<b>Lista de Solicitações de Alteração de Pedidos REALIZADAS nos últimos 15 dias</b>" &_
              iif( filtro = "SL", " (sem liquidação)", " (só liquidação)" ) &_
            "</th>" &_
          "</tr>" &_
          "<tr style='font-family:arial; font-size:12px;background-color:#EEEEEE;padding:5px'>" &_
            "<th align='center' >&nbsp;</th>" &_
            "<th align='center' >Pedido</th>" &_
            "<th align='center' >Cliente</th>" &_
            "<th align='center' >&nbsp;</th>" &_
            "<th align='center' >Tipo</th>" &_
            "<th align='center' >De</th>" &_
            "<th align='center' >Para</th>" &_
            "<th align='center' >Realizado</th>" &_
            "<th align='center' >Obs</th>" &_
            "<th align='center' >Solicitante</th>" &_
            "<th align='center' >Data</th>" &_
            "<th align='center' >Hora</th>" &_
            "<th align='center' >Realizou</th>" &_
            "<th align='center' >Data</th>" &_
          "</tr>" &_
        "</thead>" &_
        "<tbody>"
    
    while not PIR_HELPDESK.EOF

      Table = Table & "" &_
                "<tr>" &_
                  "<td align='center'>&nbsp;</td>" &_
                  "<td align='center'>" & trim(PIR_HELPDESK("Pedido")) & "</td>" &_
                  "<td align='center'>" & trim(PIR_HELPDESK("Cliente")) & "</td>" &_
                  "<td align='center'>" & trim(PIR_HELPDESK("OperacaoId")) & "</td>" &_
                  "<td align='left'>" & trim(PIR_HELPDESK("Operacao")) & "</td>" &_
                  "<td align='center'>" & trim(PIR_HELPDESK("De")) & "</td>"

      if trim(PIR_HELPDESK("OperacaoId")) = "5" AND trim(PIR_HELPDESK("Para2")) <> "" then
        Table = Table & "<td align='center'>" & FormatarDataPTBR(trim(PIR_HELPDESK("Para"))) & " a " & FormatarDataPTBR(trim(PIR_HELPDESK("Para2"))) & "</td>"
      else
        Table = Table & "<td align='center'>" & trim(PIR_HELPDESK("Descricao")) & "</td>"
      end if         

      Table = Table & "" &_                  
                iif( (trim(PIR_HELPDESK("Realizado")) = "1"), "<td align='center'>SIM</td>", "<td align='center'>NÃO</td>" ) &_
                "<td align='center'>" & trim(PIR_HELPDESK("Observacao")) & "</td>" &_
                "<td align='center'>" & trim(PIR_HELPDESK("Login")) & "</td>" 

      if trim(PIR_HELPDESK("Data")) <> "" then
        Table = Table & "<td align='center'>" & objUtil.AcertaData(trim(PIR_HELPDESK("Data"))) & " " & DateDiff("d", objUtil.AcertaData(PIR_HELPDESK("Data")), Date()) & "</td>" 
      else
        Table = Table & "<td align='center'>&nbsp;</td>"
      end if

      Table = Table & "<td align='right'>" & trim(PIR_HELPDESK("Hora")) & "</td>" &_
                  "<td align='right'>" & trim(PIR_HELPDESK("QuemRealizou")) & "</td>"

      if trim(PIR_HELPDESK("DtRealizou")) <> "" then
        Table = Table & "<td align='center'>" & objUtil.AcertaData(trim(PIR_HELPDESK("DtRealizou"))) & " " & DateDiff("d", objUtil.AcertaData(PIR_HELPDESK("DtRealizou")), Date()) & "</td>"
      else
        Table = Table & "<td align='center'>&nbsp;</td>"
      end if		  

      Table = Table & "<td align='center'></td>" &_
		                "</tr>"

      PIR_HELPDESK.MoveNext

    wend

    Table = Table & "</tbody>" &_
                  "</table>"

'    response.write 0 & "||" & Table & "||BuscaSolicitacoesRealizadas||"
	  response.write Table
    
  else

    Table = "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='1'>" &_
              "<tr>" &_
                "<td>" &_
                  "<center>" &_
                    "NÃO HÁ ITENS CADASTRADOS ..." &_
                  "</center>" &_
                "</td>" &_
              "</tr>" &_
            "</table>"

    Lista = 0 & "||" & Table

'	  response.write 0 & "||" & Table & "||BuscaSolicitacoesRealizadas||"
	  response.write Table

  end if
  
  PIR_HELPDESK.Close()
  set PIR_HELPDESK = nothing

end sub
' -------------------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------------------
sub BuscaSolicitacoesDescartadas()

  if Session("Nivel") = 0 then ' NIVEL USUARIO ----------------------------------------------------

    SQL = "" &_
      "SELECT " &_
        "a.Id As Id, " &_
        "Pedido As Pedido, " &_
        "A1_NREDUZ As Cliente, " &_
        "b.Id as OperacaoId, " &_
        "Operacao As Operacao, " &_
        "CASE WHEN b.Id = 3 THEN De + ' - ' + REPLACE(cast(cast((DADOSADV.dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(De,'.',''),',','.') as float)) as int) as varchar(10)),'.',',') +'Kg' ELSE De END AS De, " &_
        "CASE WHEN b.Id = 3 THEN Para + ' - ' + REPLACE(cast(cast((DADOSADV.dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(Para,'.',''),',','.') as float)) as int) as varchar(10)),'.',',') +'Kg' ELSE Para END AS Para, " &_
        "Para2, " &_
        "CASE b.Id " &_
        "WHEN 1 THEN CASE Para WHEN 'C' THEN 'CIF' ELSE 'FOB' END " &_
        "WHEN 6 THEN ( SELECT top 1 E4_DESCRI COLLATE Latin1_General_BIN " &_
                      "FROM DADOSADV.dbo.ZZ7010 " &_
                        "INNER JOIN DADOSADV.dbo.SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' " &_
                      "WHERE E4_CODIGO COLLATE Latin1_General_BIN = Para) " &_
        "WHEN 10 THEN CASE Para WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END " &_
        "ELSE Para END AS Descricao, " &_
        "Observacao As Observacao, " &_
        "Realizado As Realizado, " &_
        "Data As Data, " &_
        "Hora As Hora, " &_
        "Login As Login " &_
      "FROM InHistoricoAlteracoesPedidosSolicitacao as a " &_
        "INNER JOIN InHistoricoTipoOperacao as b on a.Tipo = b.Id " &_
        "INNER JOIN DADOSADV.dbo.SC5010 ON DADOSADV.dbo.SC5010.C5_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
        "INNER JOIN DADOSADV.dbo.SC6010 ON DADOSADV.dbo.SC6010.C6_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
        "INNER JOIN DADOSADV.dbo.SA1010 ON DADOSADV.dbo.SA1010.A1_COD = DADOSADV.dbo.SC5010.C5_CLIENTE AND DADOSADV.dbo.SA1010.A1_LOJA = DADOSADV.dbo.SC5010.C5_LOJACLI " &_
      "WHERE Descartado = 1 " &_
        "AND Data >= convert(varchar(8),dateadd(Day,-15,getdate()),112) " &_
        iif( filtro = "CL", "AND b.Id = 9 AND Para = 'L' ", "AND b.Id <> 9 " ) &_
      "ORDER BY Realizado, Data, Hora"

  else

    SQL = "" &_
      "SELECT " &_
        "a.Id As Id, " &_
        "Pedido As Pedido, " &_
        "A1_NREDUZ As Cliente, " &_
        "b.Id as OperacaoId, " &_
        "Operacao As Operacao, " &_
        "CASE WHEN b.Id = 3 then De + ' - ' + REPLACE(cast(cast((DADOSADV.dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(De,'.',''),',','.') as float)) as int) as varchar(10)),'.',',') +'Kg' ELSE De END AS De, " &_
        "CASE WHEN b.Id = 3 then Para + ' - ' + REPLACE(cast(cast((DADOSADV.dbo.funcCalculaPeso(C6_PRODUTO,C5_NUM,C5_INPUBE)*cast(replace(REPLACE(Para,'.',''),',','.') as float)) as int) as varchar(10)),'.',',') +'Kg' ELSE Para END AS Para, " &_
        "Para2, " &_
        "CASE b.Id " &_
          "WHEN 1 THEN CASE Para WHEN 'C' THEN 'CIF' ELSE 'FOB' END " &_
          "WHEN 6 THEN ( SELECT TOP 1 E4_DESCRI COLLATE Latin1_General_BIN " &_
                        "FROM DADOSADV.dbo.ZZ7010 " &_
                          "INNER JOIN DADOSADV.dbo.SE4010 ON ZZ7_COD = E4_CODIGO AND SE4010.D_E_L_E_T_ = '' " &_
                        "WHERE E4_CODIGO COLLATE Latin1_General_BIN = Para) " &_
          "WHEN 10 THEN CASE Para WHEN 'U' THEN 'Uberaba' ELSE 'Pirapetinga' END " &_
        "ELSE Para END AS Descricao, " &_
        "Observacao As Observacao, " &_
        "Realizado As Realizado, " &_
        "Data As Data, " &_
        "Hora As Hora, " &_
        "Login As Login " &_
      "FROM InHistoricoAlteracoesPedidosSolicitacao as a " &_
        "INNER JOIN InHistoricoTipoOperacao as b on a.Tipo = b.Id " &_
        "INNER JOIN DADOSADV.dbo.SC5010 ON DADOSADV.dbo.SC5010.C5_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
        "INNER JOIN DADOSADV.dbo.SC6010 ON DADOSADV.dbo.SC6010.C6_NUM = a.Pedido COLLATE Latin1_General_BIN " &_
        "INNER JOIN DADOSADV.dbo.SA1010 ON DADOSADV.dbo.SA1010.A1_COD = DADOSADV.dbo.SC5010.C5_CLIENTE AND DADOSADV.dbo.SA1010.A1_LOJA = DADOSADV.dbo.SC5010.C5_LOJACLI " &_
      "WHERE Descartado = 1 " &_
        "AND Data >= convert(varchar(8),dateadd(Day,-15,getdate()),112) " &_
        "AND lower(Login) = lower('" & Session("Usr") & "') " &_
        iif( filtro = "CL", "AND b.Id = 9 AND Para = 'L' ", "AND b.Id <> 9 " ) &_
    "ORDER BY Realizado, Data, Hora" 

  end if

  set PIR_HELPDESK = Conecta( "PIR-HELPDESK", cmd )

  if not PIR_HELPDESK.Eof then

    Table = "<table class='table-responsive table-sm table-striped table-hover'>" &_
              "<thead>" &_
                "<tr style='font-family:arial; font-size:12px;background-color:#EEEEEE;padding:5px'>" &_
                  "<th style='font-size:12px;padding:5px' align='center' colspan=12>" &_
                    "<b>LISTA DE SOLICITAÇÕES DE ALTERAÇÕES DE PEDIDOS DESCARTADAS NOS ÚLTIMOS 15 DIAS</b>" &_
                    iif( filtro = "SL", " (sem liquidação)", " (só liquidação)" ) &_
                  "</th>" &_
                "</tr>" &_
                "<tr style='font-family:arial; font-size:12px;background-color:#cccccc;padding:5px'>" &_
                  "<th align='center' >&nbsp;</th>" &_
                  "<th align='center' >Pedido</th>" &_
                  "<th align='center' >Cliente</th>" &_
                  "<th align='center' >&nbsp;</th>" &_
                  "<th align='center' >Tipo</th>" &_
                  "<th align='center' >De</th>" &_
                  "<th align='center' >Para</th>" &_
                  "<th align='center' >Realizado</th>" &_
                  "<th align='center' >Obs</th>" &_
                  "<th align='center' >Solicitante</th>" &_
                  "<th align='center' >Data</th>" &_
                  "<th align='center' >Hora</th>" &_
                "</tr>" &_
              "</thead>" &_
              "<tbody>"
	  a = 1

    while not PIR_HELPDESK.Eof

		  Table = Table & "" &_
                "<tr>" &_
                  "<td align='center'>&nbsp;</td>" &_
                  "<td align='center'>" & trim(PIR_HELPDESK("Pedido")) & "</td>" &_
                  "<td align='center'>" & trim(PIR_HELPDESK("Cliente")) & "</td>" &_
                  "<td align='center'>" & trim(PIR_HELPDESK("OperacaoId")) & "</td>" &_
                  "<td align='left'>" & trim(PIR_HELPDESK("Operacao")) & "</td>" &_
                  "<td align='right'>" & trim(PIR_HELPDESK("De")) & "</td>"

      if trim(PIR_HELPDESK("OperacaoId")) = "5" AND trim(PIR_HELPDESK("Para2")) <> "" then
        Table = Table & "<td align='center'>" & FormatarDataPTBR(trim(PIR_HELPDESK("Para"))) & " à " & FormatarDataPTBR(trim(PIR_HELPDESK("Para2"))) & "</td>"
      else
        Table = Table & "<td align='center'>" & trim(PIR_HELPDESK("Descricao")) & "</td>"
      end if         

      Table = Table & "" &_                  
                  iif( (trim(PIR_HELPDESK("Realizado")) = "1"), "<td align='center'>SIM</td>", "<td align='center'>NÃO</td>" ) &_
                  "<td align='center'>" & trim(PIR_HELPDESK("Observacao")) & "</td>" &_
                  "<td align='center'>" & trim(PIR_HELPDESK("Login")) & "</td>" &_
                  "<td align='center'>" & objUtil.AcertaData(trim(PIR_HELPDESK("Data"))) & " " & DateDiff("d", objUtil.AcertaData(PIR_HELPDESK("Data")), Date()) & "</td>" &_
                  "<td align='right'>" & trim(PIR_HELPDESK("Hora")) & "</td>" &_
                  "<td align='center'></td>" &_
                "</tr>"
		  a = a + 1
      PIR_HELPDESK.MoveNext

    wend

    Table = Table & "</tbody>" &_
                  "</table>"
    
'    response.write 0 & "||" & Table & "||BuscaSolicitacoesDescartadas||"
	  response.write Table

  else

    Table = "<table class='table-responsive table-sm table-striped table-hover'>" &_
              "<tr>" &_
                "<td>" &_
                  "<center>" &_
                    "NÃO HÁ ITENS CADASTRADOS ..." &_
                  "</center>" &_
                "</td>" &_
              "</tr>" &_
            "</table>"

    Lista = 0 & "||" & Table

'	  response.write 0 & "||" & Table & "||BuscaSolicitacoesDescartadas||"
	  response.write Table

  end if
  
  PIR_HELPDESK.Close()
  set PIR_HELPDESK = nothing

end sub
' Verificando se ocorreram erros no processamento da pagina ---------------------------------------
if Err > 0 then
  response.write 1 & "||Ocorreu um erro: " & escape(Err.Description)
  response.end
end if
'.................................................................................................

function VerificaDataDeEntregaNoFinalDeSemana( pPedido, pData )

  Erro    = "0"
  MsgData = ""

  if pData <> "" then

    SQL = "" &_
      "SELECT " &_
        "SC5.C5_CLIENTE, " &_
        "SC5.C5_NUM, " &_
        "SA1.A1_DTEMENO, " &_
        "SA1.A1_DTEMAIO, " &_
        "SA1.A1_DIASPRO, " &_
        "SA1.A1_DTENTLI, " &_
        "SB1.B1_COD " &_
      "FROM SC5010 SC5 " &_
      "LEFT JOIN SA1010 SA1 ON SC5.C5_CLIENTE = SA1.A1_COD AND SC5.C5_LOJACLI = SA1.A1_LOJA " &_
      "LEFT JOIN SC6010 SC6 ON SC5.C5_NUM = SC6.C6_NUM " &_
      "LEFT JOIN SB1010 SB1 ON SC6.C6_PRODUTO = SB1.B1_COD " &_
      "WHERE SC5.C5_NUM = '" & pPedido & "' "

    SET PIR_DADOSADV = Conecta( "PIR-DADOSADV", SQL )

    if not isObject(PIR_DADOSADV) or (PIR_DADOSADV IS NOTHING) then
      response.end
    end if

    if not PIR_DADOSADV.EOF then

      DiasPro = PIR_DADOSADV("A1_DIASPRO")

      PrimeiraData = pData
      DiasMaio = CInt(PIR_DADOSADV("A1_DTEMAIO"))
      SegundaData = DateAdd( "d", DiasMaio, PrimeiraData )

      SQL = "" &_
        "SELECT " &_
          "ATIVO " &_
        "FROM DEF_ENTREGA_FDS " &_
        "WHERE COD_CLIENTE = '" & trim(PIR_DADOSADV("C5_CLIENTE")) & "' " &_
          "AND (QUEM_EXCLUIU = 0 OR QUEM_EXCLUIU = '') "

      'RESPONSE.WRITE SQL

      SET DEF_ENTREGA_FDS = Conecta( "PIR-CORPORATIVO", SQL )

      if not isObject(DEF_ENTREGA_FDS) or (DEF_ENTREGA_FDS IS NOTHING) then
        response.end
      end if

      ENTREGA_NO_FDS = False

      if (NOT DEF_ENTREGA_FDS.EOF) then 
      
        if (DEF_ENTREGA_FDS("ATIVO")) then 

          ENTREGA_NO_FDS = True

        end if

      end if

      if (NOT ENTREGA_NO_FDS) then

        '-------------------------------------------------------------------------------------
        '							VALIDANDO PRIMEIRA DATA SE SABADO
        '						  ACRESCENTA MAIS UM DIA À PRIMEIRA DATA
        '-------------------------------------------------------------------------------------
        if WeekDay(PrimeiraData) = 7 then

          PrimeiraData = DateAdd( "d", 2, PrimeiraData )
          SegundaData = DateAdd( "d", DiasMaio, PrimeiraData )

          MsgData = "Sua 1ª data de entrega foi alterada para " & PrimeiraData 
          MsgData = MsgData & "Sua 2ª data de entrega foi alterada para " & SegundaData 
          para  = PrimeiraData
          para2 = SegundaData
          
        end if

        '-------------------------------------------------------------------------------------
        '							VALIDANDO PRIMEIRA DATA SE DOMINGO
        '						  ACRESCENTA MAIS UM DIA À PRIMEIRA DATA
        '-------------------------------------------------------------------------------------
        if WeekDay(PrimeiraData) = 1 then

          PrimeiraData = DateAdd( "d", 1, PrimeiraData )
          SegundaData = DateAdd( "d", DiasMaio, PrimeiraData )

          MsgData = MsgData & "Sua 1ª data de entrega foi alterada para " & PrimeiraData & ". "
          MsgData = MsgData & "Sua 2ª data de entrega foi alterada para " & SegundaData & ". "
          para  = PrimeiraData
          para2 = SegundaData

        end if

        ' Conforme autorizado pelo Carlos Sobrinho em 05/08/19, por Sandro Cardoso 
        '-------------------------------------------------------------------------------------
        '							VALIDANDO SEGUNDA DATA SE SABADO
        '						ACRESCENTA MAIS DOIS DIAS À SEGUNDA DATA
        '-------------------------------------------------------------------------------------
        if WeekDay(SegundaData) = 7 then

          SegundaData = DateAdd( "d", 2, SegundaData )
          MsgData = MsgData & "\\n Sua 2ª data de entrega foi alterada para " & SegundaData 
          para2 = SegundaData

        end if

        '-------------------------------------------------------------------------------------
        '							VALIDANDO SEGUNDA DATA SE DOMINGO
        '						  ACRESCENTA MAIS UM DIA À SEGUNDA DATA
        '-------------------------------------------------------------------------------------
        if WeekDay(SegundaData) = 1 then

          SegundaData = DateAdd( "d", 1, SegundaData )
          MsgData = MsgData & "\\n Sua 2ª data de entrega foi alterada para " & SegundaData 
          para2 = SegundaData

        end if

      else

        PrimeiraData = FormatarDataPTBR( request("De") )
        SegundaData = FormatarDataPTBR( SegundaData )

      end if

      DEF_ENTREGA_FDS.Close()
      SET DEF_ENTREGA_FDS = NOTHING

    end if

    PIR_DADOSADV.Close()
    SET PIR_DADOSADV = NOTHING

  end if

  ' RESPONSE.WRITE "Para   : " & para & "<BR>"
  ' RESPONSE.WRITE "DiaMaio: " & DiasMaio & "<BR>"
  ' RESPONSE.WRITE "Para2  : " & para2 & "<BR>"
  ' RESPONSE.WRITE "MsgData: " & MsgData & "<BR>"

  ' msgAss = "(1.1) ALTERA PEDIDO - 1ª DATA DE ENTREGA"
  ' msgBod = "<br>PARA:" & trim(para) & "<br>FormatarDataPTBR: " & FormatarDataPTBR( trim(para) ) & "<br>PrimeiraData: " & PrimeiraData & "<br>SegundaData: " & SegundaData

  ' CALL EnviaEmail( "jorge.queiroz@smurfitwestrock.com.br","","","","", msgBod, msgAss  )

  VerificaDataDeEntregaNoFinalDeSemana = MsgData
end function 
'--------------------------------------------------------------------------------------------------'
function FormatarDataPTBR(aData)

  Dim ano, mes, dia

  if InStr(1, aData, "/") > 0 then

    FormatarDataPTBR = aData

  else

    if InStr(1, aData, "-") > 0 then

      ' Extrai partes da data ex.: 2025-12-31
      ano = Left(aData, 4)
      mes = Mid(aData, 6, 2)
      dia = Right(aData, 2)
      
      ' Retorna no formato DD/MM/YYYY
      FormatarDataPTBR = dia & "/" & mes & "/" & ano

    else

      ' Extrai partes da data ex.: 20251231
      ano = Left(aData, 4)
      mes = Mid(aData, 5, 2)
      dia = Right(aData, 2)
      
      ' Retorna no formato DD/MM/YYYY
      FormatarDataPTBR = dia & "/" & mes & "/" & ano

    end if

  end if

end function

Sub SalvarEvidencias(idSolicitacao, evidenciasJSON)
    Dim SQL, PIR_HELPDESK

    On Error Goto 0

    evidenciasJSON = Replace(evidenciasJSON, "[", "")
    evidenciasJSON = Replace(evidenciasJSON, "]", "")
    
    Dim anexosArray, anexo, i
    anexosArray = Split(evidenciasJSON, "},{")
    
    For i = 0 To UBound(anexosArray)
        anexo = anexosArray(i)
        
        Dim nome, caminho
        nome = ExtrairValorJSON(anexo, "name")
        caminho = ExtrairValorJSON(anexo, "uploadedName")
        
        SQL = "INSERT INTO SolicitacaoAlteracaoPedidoEvidencia "
        SQL = SQL & "(IdSolicitacaoAlteracaoPedido, Nome, Caminho, DataCriacao) VALUES ("
        SQL = SQL & idSolicitacao & ", "
        SQL = SQL & "'" & nome & "', "
        SQL = SQL & "'" & caminho & "', "
        SQL = SQL & "GETDATE()"
        SQL = SQL & ")"
        
        On Error Resume Next
        SET PIR_HELPDESK = Conecta("PIR-HELPDESK", SQL)
        
        If Err.Number = 0 AND IsObject(PIR_HELPDESK) Then
            PIR_HELPDESK.Close()
            SET PIR_HELPDESK = NOTHING
        End If
        On Error Goto 0
    Next
End Sub
%>