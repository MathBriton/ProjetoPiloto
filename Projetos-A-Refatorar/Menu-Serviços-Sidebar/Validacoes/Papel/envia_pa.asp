<%Response.Buffer=True%>
<%
    SESSION.LCID=1046
    Server.ScriptTimeout=2000
%>
<!--#include file="../../../../includes/variaveis.asp" -->
'<!--#include file="../../../../includes/conecta_banco_siga.inc" --> 'DADOSADV
'<!--#include file="../../../../cnn/DADOSADV.ASP" -->
<!--#include file="../../../../cnn/cnn.asp" -->
<html>
<head>
<title>Envia todos os pedidos liberados</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">

.barra_menu {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; color: #FFFFFF; text-decoration: none}
.barra_menu:hover {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; color: #99CCFF; text-decoration: none}

.asp {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-style: normal; font-weight: normal; text-transform: uppercase; text-decoration: none; line-height: normal; font-variant: normal; color: #000000}
</style>
</head>
<%

Dados.Open("EXEC SP_PEDIDOS_PAPEL_LIBERADOS")

vTotPed = 0
vTotPedEnv = 0

do until Dados.Eof

   vTotPed = vTotPed + 1
   '********************************************************************'
   '*  Verificando se os pedidos liberados possuem alguma incosist�ncia	'
   '********************************************************************'
    vzf_numero  = Dados("ZF_NUMERO")
    vzf_produto = Dados("ZF_PRODUTO")
    vzf_emissao = Dados("ZF_EMISSAO")
    vzf_ent1    = Dados("ZF_ENT1")
    vzf_IdSF    = Trim(Dados("ZF_ID_SF"))

    If  Trim(Dados("ZF_ENT2")) <> "" then
        vzf_ent2    = Dados("ZF_ENT2")
    else
        vzf_ent2 = " "
    end if
   
   SQL = "EXEC SP_ENVIA_PEDIDO_PAPEL '"&vzf_numero&"',"
   SQL = SQL & "'" &vzf_produto&"',"
   SQL = SQL & "'" &vzf_emissao&"',"
   SQL = SQL & "'" &vzf_ent1&"',"
   SQL = SQL & "'" &vzf_ent2&"'"

   Dados2.open(SQL)

   'response.write Dados2("TESTE")
   '************************************************************'
   ' Se estiver OK, envio os dados do pedido para o SC5 e SC6   '
   '************************************************************'	

   if Trim(Dados2("SITUACAO")) = "" then
       
       '**********************'  
	   '*****    S C 5 *******'
	   '**********************'	

      vc5_Filial  = "  "
      vc5_tipo    = "N"
      vzf_cliente = Dados("ZF_CLIENTE")
      vzf_lojacli = Dados("ZF_LOJACLI") 
      vc5_tipocli = "R"
      vzf_tpfrete = Dados("ZF_TPFRETE")
      vzf_condpag = Dados("ZF_CONDPAG")
      vzf_vencpag = Dados("ZF_DATA5")
      vc5_tabela  = "1"
      vzf_emissao = Dados("ZF_EMISSAO")
      vc5_moeda   = Dados("ZF_MOEDA")
      vc5_moedasa = Dados("ZF_MOEDASA")
      vzf_desc    = Dados("ZF_DESC")
       

      vzf_destino = Dados("ZF_DESTINO")
      if vzf_destino = "" or vzf_destino = " " then
         vzf_destino = vzf_cliente
      end if

      vzf_lojades = Dados("ZF_LOJADES")

      if vzf_lojades = ""  or vzf_lojades = " " then
         vzf_lojades = vzf_lojacli
      end if

      vzf_remessa = Dados("ZF_REMESSA")
      va1_end     = Dados("A1_END")
      va1_endent  = Dados("A1_ENDENT")
      va1_mun     = Dados("A1_MUN")
      va1_mune    = Dados("A1_MUNE")
      va1_est     = Dados("A1_EST")
      va1_este    = Dados("A1_ESTE")
      va1_nome    = Dados("A1_NOME")
      va1_bairro  = Dados("A1_BAIRRO")
      va1_bairroE  = Dados("A1_BAIRROE")

       '*****************************'
       '*****  Se for remessa   *****'
       '*****************************'
      if vzf_remessa = "F" then

         vc5_destino = vzf_destino
         vc5_lojdest = vzf_lojades 

         VarCliDest = "Select A1_ENDENT,A1_END,A1_MUN,A1_MUNE,A1_EST,A1_ESTE,A1_NOME"
         VarCliDest = VarCliDest & " from SA1010 where A1_COD = '"&vzf_destino&"' and A1_LOJA = '"&vzf_lojades&"' and " 
         VarCliDest = VarCliDest & " D_E_L_E_T_ <> '*'"

         Dados3.Open(VarCliDest)

         if va1_endent = "" then

            vc5_end = Dados3("a1_end")
            vc5_mun = Dados3("a1_mun")
            vc5_est = Dados3("a1_est")

         else

            vc5_end = Dados3("a1_endent")
            vc5_mun = Dados3("a1_mune")	
            vc5_est = Dados3("a1_este")

         end if 			 	

         va1_nome    = Dados3("a1_nome")

      else

         vc5_destino = vzf_cliente
         vc5_lojdest = vzf_lojacli 

         if va1_endent = "" then

            vc5_end = va1_end & " " & va1_bairro
            vc5_mun = va1_mun
            vc5_est = va1_est

         else

            vc5_end = va1_endent & " " & va1_bairroe
            vc5_mun = va1_mune	
            vc5_est = va1_este

         end if 			 	
         '** Como a NF � normal, o nome j� foi pego antes ! **'
      end if

      if vc5_end = "" then
         vc5_end = " "
      end if

      if vc5_mun = "" then
         vc5_mun = " "
      end if

      if vc5_est = "" then
         vc5_est = " "
      end if

	   '.........................................................................................
	   'AGORA PEGA O PEDIDO DO SZF QUE FOI SELECIONADO PELO USU�RIO - N�O PEGA MAIS DO ZF PEGA DO A1
	   va1_vend     = Dados("A1_VEND")
	   'va1_vend     = Dados("ZF_VEND")  
	   '.........................................................................................  

	   va1_comis    = replace(Dados("A1_COMIS"),"," , ".")

      vc5_status   = " "
      vc5_envrio   = " "
      vzf_obsvend  = Dados("ZF_OBSVEND")   

      if vzf_obsvend = "" then
         vzf_obsvend = " "
      end if

      vzf_digitad  = Dados("ZF_DIGITAD")

      if vzf_digitad = "" then
         vzf_digitad = " "
      end if
       
	   '*********************************************'	
       '  			Data e hora do servidor          '
   	   '*********************************************'	
		'DataHora = Now
		'Espaco = InStr(1,DataHora, " ")
		'HW = Mid(DataHora,1,Espaco-1)
		'if Len(HW) = 10 then
		'    Hoje = Mid(HW,7,4) & Mid(HW,4,2) & Mid(HW,1,2) 
		'else
		'    Hoje = "20" & Mid(HW,7,2) & Mid(HW,4,2) & Mid(HW,1,2) 
		'end if

		'DoisPontos = InStr(1,DataHora,":")
		'Hora = Mid(DataHora,DoisPontos-2,2) & Mid(DataHora,DoisPontos,6) 

        'CHAMA DATA DA FUN��O INVERTEDATASIGA AAAAMMDD E HORA DO TIME LCID 1046
        Hoje = Trim(InverteDataSiga(Date()))
        Hora = Trim(Time())

	   vdatacad     = Hoje
	   vhoracad     = Hora
	   vzf_quemlib  = Dados("ZF_QUEMLIB")
       if vzf_quemlib = "" then
          vzf_quemlib = " "
       end if
       
	   vzf_datalib  = Dados("ZF_DATALIB") 
       if vzf_datalib = "" then
          vzf_datalib = " "
       end if

	   vzf_horalib  = Dados("ZF_HORALIB") 
       if vzf_horalib = "" then
          vzf_horalib = " "
       end if

       vzf_fretuni = Dados("ZF_FRETUNI")
       vzf_fretuni = replace(vzf_fretuni,",",".")

       '*****************************************'
	   '****             S C 6              *****'
       '*****************************************'
	   vc6_filial   = "  "
	   vc6_item     = "01"

       vc6_descri = Dados("B1_DESC")
       vzf_comp = Dados("ZF_COMPAP")
       vzf_grampap = Dados("ZF_GRAMPAP")
       vzf_formpap = Dados("ZF_FORMPAP")
       vzf_diampap = Dados("ZF_DIAMPAP")
       vzf_tubepap = Dados("ZF_TUBEPAP")
       vzf_transpa = Dados("ZF_TRANSPA")
      
       if vzf_grampap = "" then
          vzf_grampap = " "
       end if    

       if vzf_formpap = "" then
          vzf_formpap = " "
       end if    

       if vzf_diampap = "" then
          vzf_diampap = " "
       end if    

       if vzf_tubepap = "" then
          vzf_tubepap = " "
       end if    

       if vzf_transpa = "" then
          vzf_transpa = " "
       end if    

        '**** vc6_Newdescri = Trim(vc6_descri) &" "&  Trim(vzf_comp) ******'
		vc6_Newdescri = Trim(vc6_descri)&","& Trim(vzf_grampap)&","
		vc6_Newdescri = vc6_Newdescri & Trim(vzf_formpap)

       '**** &","& Trim(vzf_diampap)&","& Trim(vzf_tubepap) ****'
		vc6_newdescri = mid(vc6_newdescri,1,30)

	   vzf_tes     = Dados("ZF_TES")
       vf4_cf      = Dados("F4_CF") 
	   
	   If Trim(va1_est) = "MG" then
	   		vc6_cf = vf4_cf
       End If
	   If Trim(va1_est) = "EX" Then
	   		vc6_cf = "7" & mid(vf4_cf,2,3)
	   End If
	   If Trim(va1_est) <> "EX" And Trim(va1_est) <> "MG" Then
            vc6_cf = "6" & mid(vf4_cf,2,3)
       end if	

       ' No caso de papel�o , preciso considerar a toler�ncia '
	   vzf_quant   = Dados("ZF_QUANT")
       vc6_qtdorig = vzf_quant

	   vzf_prcven  = Dados("ZF_PRECO")
	   vzf_valor   = (vzf_quant * 1.1) * vzf_prcven

	   if vzf_ent2 = "" or vzf_ent2 = "00000000" then
			vzf_ent2 = vzf_ent1
	   end if

	   vzf_pedcli   = Dados("ZF_PEDCLI")  	
       if vzf_pedcli = "" then
          vzf_pedcli = " "
       end if

	   vzf_larg     = Dados("ZF_LARG") 
       if vzf_larg = "" then
          vzf_larg = 0
       end if

	   vzf_comp     = Dados("ZF_COMP") 
       if vzf_comp = "" then
          vzf_comp = 0
       end if

	   vzf_nfpale   = Dados("ZF_NFPALE")
       vzf_local = "01"
	   vb1_um       = Dados("B1_UM") 
	   vb1_segum    = Dados("B1_SEGUM") 
	   vzf_vinco1   = Dados("ZF_VINCO1")   

       if vzf_vinco1 = "" then
          vzf_vinco1 = 0
       end if

	   vzf_vinco2   = Dados("ZF_VINCO2") 
       if vzf_vinco2 = "" then
          vzf_vinco2 = 0
       end if

	   vzf_vinco3   = Dados("ZF_VINCO3") 
       if vzf_vinco3 = "" then
          vzf_vinco3 = 0
       end if

	   vzf_vinco4   = Dados("ZF_VINCO4") 
       if vzf_vinco4 = "" then
          vzf_vinco4 = 0
       end if

	   vzf_vinco5   = Dados("ZF_VINCO5") 
       if vzf_vinco5 = "" then
          vzf_vinco5 = 0
       end if

	   vb1_difsus   = Dados("B1_DIFSUS") 
       if vb1_difsus = "" then
          vzf_obsvend = " "
       end if

	   vf4_clasfis  = Dados("F4_CLASFIS") 
       if vf4_clasfis = "" or vf4_clasfis = "000" then
          vf4_clasfis = " "
       end if

       vzf_DescCom = Dados("ZF_DESCCOM")
       
       If Trim(Dados("ZF_MENNOTA")) <> "" then
            ZF_MENNOTA = Trim(Dados("ZF_MENNOTA"))
       else
            ZF_MENNOTA = " "
       end if

       If Trim(Dados("ZF_MENNOT2")) <> "" then
            ZF_MENNOT2 = Trim(Dados("ZF_MENNOT2"))
       else
            ZF_MENNOT2 = " "
       end if

       If Trim(Dados("ZF_MENNOT3")) <> "" then
            ZF_MENNOT3 = Trim(Dados("ZF_MENNOT3"))
       else
            ZF_MENNOT3 = " "
       end if

       If Trim(Dados("ZF_MENNOT4")) <> "" then
            ZF_MENNOT4 = Trim(Dados("ZF_MENNOT4"))
       else
            ZF_MENNOT4 = " "
       end if

        If Trim(Dados("ZF_NUMPCOM")) <> "" then
            ZF_NUMPCOM = Trim(Dados("ZF_NUMPCOM"))
       else
            ZF_NUMPCOM = " "
       end if      

        If Trim(Dados("ZF_ITEMPC")) <> "" then
            ZF_ITEMPC = Trim(Dados("ZF_ITEMPC"))
       else
            ZF_ITEMPC = "0"
       end if     

       VarSC = "EXEC SP_GRAVA_PEDIDO_PAPEL [" & vc5_Filial & "],"
       VarSC = VarSC & "[" & vzf_numero & "],"
       VarSC = VarSC & "[" & vc5_tipo & "],"
       VarSC = VarSC & "[" & vzf_cliente & "],"
       VarSC = VarSC & "[" & vzf_lojacli & "],"
       VarSC = VarSC & "[" & vc5_tipocli & "],"
       VarSC = VarSC & "[" & vzf_tpfrete & "],"
       VarSC = VarSC & "[" & vzf_condpag & "],"
	    VarSC = VarSC & "[" & vzf_vencpag & "],"
       VarSC = VarSC & "[" & vc5_tabela & "],"
       VarSC = VarSC & "[" & vzf_emissao & "],"
       VarSC = VarSC & "[" & vc5_moeda & "],"
       VarSC = VarSC & "[" & vc5_moedasa & "],"
       VarSC = VarSC & "[" & vzf_desc & "],"
       VarSC = VarSC & "[" & vzf_remessa & "],"
       VarSC = VarSC & "[" & vc5_destino & "],"
       VarSC = VarSC & "[" & vc5_lojdest & "],"
       VarSC = VarSC & "[" & vc5_end & "],"
       VarSC = VarSC & "[" & vc5_mun & "],"
       VarSC = VarSC & "[" & vc5_est & "],"
       VarSC = VarSC & "[" & va1_nome & "],"
       VarSC = VarSC & "[" & va1_vend & "],"
       VarSC = VarSC & "[" & va1_comis & "],"
       VarSC = VarSC & "[" & vc5_status & "],"
       VarSC = VarSC & "[" & vc5_envrio & "],"
       VarSC = VarSC & "[" & vzf_obsvend & "],"
       VarSC = VarSC & "[" & vzf_digitad & "],"
       VarSC = VarSC & "[" & vdatacad & "],"
       VarSC = VarSC & "[" & vhoracad & "],"
       VarSC = VarSC & "[" & vzf_quemlib & "],"
       VarSC = VarSC & "[" & vzf_datalib & "],"
       VarSC = VarSC & "[" & vzf_horalib & "],"

	   '***************    S  C  6      ***********'	
       VarSC = VarSC & "[" &vc6_filial& "],"
       VarSC = VarSC & "[" &vc6_item & "],"
       VarSC = VarSC & "[" &vzf_produto& "],"
       VarSC = VarSC & "[" &vc6_Newdescri& "],"
       VarSC = VarSC & "[" &vzf_grampap& "],"
       VarSC = VarSC & "[" &vzf_formpap& "],"
       VarSC = VarSC & "[" &vzf_diampap& "],"
       VarSC = VarSC & "[" &vzf_tubepap& "],"
       VarSC = VarSC & "[" &vzf_transpa& "],"
       VarSC = VarSC & "[" &vzf_tes& "],"
       VarSC = VarSC & "[" &vc6_cf & "],"
       VarSC = VarSC & "[" &vzf_quant& "],"

       vzf_prcven = replace(vzf_prcven,",",".")

       VarSC = VarSC & "[" &vzf_prcven& "],"

       vzf_valor = replace(vzf_valor,",",".")

       VarSC = VarSC & "[" &vzf_valor& "],"
       VarSC = VarSC & "[" &vzf_ent1& "],"
       VarSC = VarSC & "[" &vzf_ent2& "],"
       VarSC = VarSC & "[" &vzf_pedcli & "],"
       VarSC = VarSC & "[" &vzf_larg & "],"
       VarSC = VarSC & "[" &vzf_comp & "],"
       VarSC = VarSC & "[" &vzf_nfpale& "],"
       VarSC = VarSC & "[" &vzf_local & "],"
       VarSC = VarSC & "[" &vb1_um & "],"
       VarSC = VarSC & "[" &vb1_segum& "],"
       VarSC = VarSC & "[" &vzf_vinco1& "],"
       VarSC = VarSC & "[" &vzf_vinco2& "],"
       VarSC = VarSC & "[" &vzf_vinco3& "],"
       VarSC = VarSC & "[" &vzf_vinco4& "],"
       VarSC = VarSC & "[" &vzf_vinco5& "],"
       VarSC = VarSC & "[" &vb1_difsus& "],"
       VarSC = VarSC & "[" &vf4_clasfis& "],"
       VarSC = VarSC & "[" &vzf_DescCom& "],"

       VarSC = VarSC & "[" &ZF_MENNOTA& "],"
       VarSC = VarSC & "[" &ZF_MENNOT2& "],"
       VarSC = VarSC & "[" &ZF_MENNOT3& "],"
       VarSC = VarSC & "[" &ZF_MENNOT4& "],"

        VarSC = VarSC & "[" &vzf_fretuni& "],"
        VarSC = VarSC & "['" &ZF_NUMPCOM& "'],"
        VarSC = VarSC & "[" &ZF_ITEMPC& "]"
       
        Dados4.Open(VarSC)

   else
        if vzf_IdSF <> "" then
            Dados3.Open("UPDATE SZF010 SET ZF_BLOQ = '"&Mid(Trim(Dados2("SITUACAO")),1,2)&"', ZF_DESCBL = '"&Mid((Dados2("SITUACAO")),4,len((Dados2("SITUACAO"))))&"', D_E_L_E_T_ = '*' WHERE ZF_NUMERO = '"&vzf_numero&"' ")
        else
            Dados3.Open("UPDATE SZF010 SET ZF_BLOQ = '"&Mid(Trim(Dados2("SITUACAO")),1,2)&"', ZF_DESCBL = '"&Mid((Dados2("SITUACAO")),4,len((Dados2("SITUACAO"))))&"' WHERE ZF_NUMERO = '"&vzf_numero&"' ")
        end if
   end if 
    
   Dados.MoveNext()
   Dados2.Close()

loop

Dados.Close()

RESPONSE.REDIRECT "pa.asp"
%>
<body bgcolor="#FFFFFF" leftmargin="25" topmargin="5" marginwidth="5" marginheight="5">
<font face="Verdana, Arial, Helvetica, sans-serif" color=#ffffff size=1><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
  </font></font> 
</body>
</html>
<%
Dim Ano, Mes, Dia
Function InverteDataSiga(InverteData)
  Ano = Year(InverteData)
  Mes = Month(InverteData)
  Dia = Day(InverteData)
  If Len(Ano) > 4 Then
    Ano = "20" & Ano
  End If		
  If Len(Mes) =1 Then 
    Mes = "0" & Mes
  End If
  If Len(Dia) = 1 Then
    Dia = "0" & Dia
  End If
  InverteDataSiga = Ano & Mes & Dia
End Function
%>