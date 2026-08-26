 <!-- #include file="../../cnn2/cnn2.asp" -->
 <!--#Include File = "../../includes/variaveis.asp"-->
 <!-- #include file="EmailExtratoPedido.asp" -->
 
<%
session.lcid = 1046

id   = Request.Form("id")
'IdUsuario = 1   'session("CODIGO")
'Representante = SESSION("NIVEL") 
TES = replace(request("Tes"),"'","") 
CodCliente =  replace(request("codCliente"),"'","") 
CodLoja = request("codLoja")
'TipoItem =  request.form("hidTipoItem")
'TipoProdutoItem = replace(request.form("cboTipoItem"),"'","")
Tipofrete = replace(request("cboTipoFrete"),"'","") 
clienteRemessa = replace(request("codClienteRemessa"),"'","")
CondicaoPagamento = replace(request("cboCondicaoFaturamento"),"'","")
tolerancia = request.form("hidTolMaior")
validarLimiteCredito = request.form("hidValidaLimiteCredito")
flagAtualizarLotePiloto = iif(Request.Form("hidAtualizaLotePiloto")="1",true,false)
AtualizarLotePilotoProduto = request.form("hidAtualizaLotePilotoProduto")
DataPrevistaEntrega = request.form("dtpDataPrevistaEntrega")

MensagemLimiteCredito = ""
MensagemCargaMaquina = ""
Planta = request.form("hidFilial")
NomePlanta = fNomePlanta(Planta)

if Nivel <> 0 then
    GeradoRepresentante = "S"
else    
    GeradoRepresentante = "N"
end if 

listaDePedido= Request.Form("listaPedido")
if trim(listaDePedido)="" then
   response.write "LISTA DE PEDIDOS VAZIA"
   response.redirect("Pedidos.asp")
   response.end
end if

dim arraymensagem()
dim arrayPedidosGravados()
ValorICMSDIF = 0
PercentualDIF = 100
            
if Planta = "05" then 
        listaDePedido=replace(listaDePedido,"|,","|")
        arrayPedido=split(listaDePedido,"|")

        linha=""
        DataDesejadaDisponivel = 1
        SaldoDisponivel = 0
        MensagemLimiteCredito = ""
        MensagemDataDisponivel = ""
        GerouArrayMensagens = false
        ContadorArrayMensagens = 1
        ContadorArrayPedido = 1
        FlagGravou=false
        nroPedido = 0
        nroPedidoBento = ""
        Representante = "''"
        TabelaComissao = "''"
        ZonaVenda = ""
        GeradoRepresentante = ""
        CST_IPI ="''"
        CST_ICMS = "''"
        ContaContabil="''"
        CodigoEnquadramentoIPI="''"
        DepositoOriginal="''"
        CentroCusto="''"
            
        for contaPorPedido = 0 to UBound(arrayPedido)
            linhaAtual = arrayPedido(contaPorPedido)
       
            if trim(linhaAtual) = ""  then 
              exit for
            end if

            Estado = ""
            ValorPISUnitario = 0
            ValorCOFINSUnitario = 0
            ValorICMSUnitario = 0
            ValorBaseIcmsUnitario = 0
            axvaloricms = 0
            ValorICMSUnitario = 0
            PercentualCOFINS = 0
            PercentualPIS = 0 
            CST_IPI = 0
            CST_ICMS = 0
            PercentualBaseCOFINS = 0
            PercentualCOFINS= 0
            ValorCOFINSUnitario= 0
            PercentualPIS= 0
            PercentualBasePIS= 0
            PercentualBaseICMS = 0
            PercentualBaseIPI  = 0
            PercentualBaseICMSBanco = 0
            PercentualBaseIPIBanco = 0
            BaseTributacaoIPITes = 0
            BaseTributacaoIPIProduto = 0
            ValorIPIUnitario = 0
            PermiteAlterarIPI = "N"
            CodigoEnquadramentoIPI = 0
            Peso = 0

            ColunaAtual = split(linhaAtual,",")
            TipoItem = left(replace(ColunaAtual(75),"'" ,""),2)
            CodProduto = ColunaAtual(20) 
            larguraChapa = ColunaAtual(22) 
            comprimentoChapa = ColunaAtual(23) 
            estruturaChapa = ColunaAtual(69)

            quantidade = replace(ColunaAtual(32),",",".") 
            valorunitario = replace(ColunaAtual(33),".",",") 
            
            DataEntrega = ColunaAtual(4) 
            DataEntregaOriginal = mid(DataEntrega,8,10)
            
            'Busca o Representante no Cadastro do Cliente
            SQL = "SELECT COD_VENDEDOR, ESTADO_CLIENTE FROM VW_CLIENTES_PEDIDO WHERE COD_CLIENTE = '" & CodCliente & "'"
            
            SET Consulta = Conecta( "BEN-DADOSADV", SQL )
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if

            If not Consulta.eof then
               Representante = Consulta("COD_VENDEDOR") 
               Estado =  Consulta("ESTADO_CLIENTE") 
            end if 
            Consulta.close()
    
            'Tabela TES 
            SQL = "SELECT CALC_IPI, CALC_ICM, PIS_COF, ISENCAO_FECP, DIFERE_ICMS, RED_BASE_ICMS, PERC_DIF FROM VW_BI_TES WHERE CODIGO  = '" & TES & "'" 
            SET Consulta = Conecta( "BEN-DADOSADV", SQL )
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if

            CalculaIPI = "N"
            CalculaICMS = "N"
            CalculaPISCofins = "N"
            IsentoFECP = "N"
            DifereICMS = "N"
            
            If not Consulta.eof then
               PercentualBaseICMS = iif(Consulta("RED_BASE_ICMS")=0, 100,replace(Consulta("RED_BASE_ICMS"),",","."))
               CalculaIPI  = Consulta("CALC_IPI")
               CalculaICMS = Consulta("CALC_ICM")
               DifereICMS =  Consulta("DIFERE_ICMS")

               if Consulta("PIS_COF") = "NENHUM" then 
                   CalculaPISCofins = "N"
               end if 

               if Consulta("PIS_COF") = "PIS" then 
                   CalculaPISCofins = "P"
               end if 

               if Consulta("PIS_COF") = "COFINS" then 
                  CalculaPISCofins = "C"
               end if 

               if Consulta("PIS_COF") = "AMBOS" then 
                  CalculaPISCofins = "A"
               end if 

               PercentualDIF = Consulta("PERC_DIF")
               IsentoFECP = Consulta("ISENCAO_FECP")
            end if 

            Consulta.close()
            
            if TipoItem = "CJ" then 
                 SQL = "SELECT TOP 1 ALIQUOTA_IPI, ESTRUTURA, PESO FROM VW_BI_PRODUTOS_PEDIDO WHERE CODIGO_CONJUNTO = " & CodProduto & " ORDER BY CODIGO_ITEM "
            else    
                if TipoItem = "PA" then 
                    SQL = "SELECT ALIQUOTA_IPI, ESTRUTURA, PESO FROM VW_PAPEL_PEDIDO WHERE CODIGO_ITEM = " & CodProduto 
                else
                    SQL = "SELECT TOP 1 ALIQUOTA_IPI, ESTRUTURA, PESO FROM VW_BI_PRODUTOS_PEDIDO WHERE CODIGO_CONJUNTO = " & CodProduto 
                end if
            end if  

            if TipoItem = "PA" then 
                SET Consulta = Conecta( "BEN-DADOSADV", SQL )
            else 
                SET Consulta = Conecta( "BEN-TRIMBOX", SQL )
            end if 

            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if

            PercentualIPI=0
            Estrutura =""
            if not Consulta.eof then 
                PercentualIPI = Consulta("ALIQUOTA_IPI")  
                Estrutura = Consulta("ESTRUTURA") 
                Peso = replace(Consulta("PESO"),",",".")
            end if  
            Consulta.close()

            'Recalcula o Peso no caso de CHAPA 
            if TipoItem = "CH" then 
                if larguraChapa > 0  and comprimentoChapa > 0 then 
                    SQL = " SELECT PESO FROM dbo.FUNC_CALCULA_AREA_PESO(" & larguraChapa & ", " & comprimentoChapa & ", '" & Estrutura & "')"
                    SET BEN_CORPORATIVO = Conecta( "BEN-CORPORATIVO", SQL )
                    if not isobject(BEN_CORPORATIVO) or (BEN_CORPORATIVO IS NOTHING) then
                        response.end
                    end if

                    if not BEN_CORPORATIVO.eof then
                        Peso =  replace(BEN_CORPORATIVO("PESO"),",",".")
                    end if 
                end if 
            end if 

            SQL = "SELECT dbo.FUNC_PERCENTUAL_ICMS('BENTO','" & Estado & "') as AliquotaICMS" 
            SET Consulta = Conecta( "BEN-CORPORATIVO", SQL )
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if

            PercentualICMS = 0 
            if not Consulta.eof then 
                PercentualICMS =  Consulta("AliquotaICMS")  
                If Estado = "RJ" and IsentoFECP = "N" then
                   PercentualICMS = PercentualICMS + 2
                end if 
            end if  
            Consulta.close()

            'Calculo dos Impostos
            if CalculaIPI = "S" then 
               ValorBaseIpiUnitario = replace(valorunitario,",",".")
               ValorIPIUnitario =  replace(valorunitario * (replace(PercentualIPI,".",",")/100),",",".")
            end if 

            if CalculaICMS = "S" AND DifereICMS = "N"  then 
                if PercentualBaseICMS = 100 then 
                   ValorBaseIcmsUnitario = valorunitario
                   ValorICMSUnitario = replace(ValorBaseIcmsUnitario * (replace(PercentualICMS,".",",")/100),",",".")
                else 
                   ValorBaseIcmsUnitario = valorunitario * (replace(PercentualBaseICMS,".",",")/100)
                   ValorICMSUnitario = replace(ValorBaseIcmsUnitario * (replace(PercentualICMS,".",",")/100),",",".")
               end if   
               ValorBaseIcmsUnitario  = replace(ValorBaseIcmsUnitario,",",".")
            end if 

            if CalculaICMS = "S" AND DifereICMS = "S"  then 
                    if PercentualDIF < 100 then 
                        ValorBaseIcmsUnitario = valorunitario
                        ValorICMSUnitario = ValorBaseIcmsUnitario * (replace(PercentualICMS,".",",")/100)
                        ValorICMSDIF = ValorICMSUnitario * (replace(PercentualDIF,".",",")/100)
                        ValorDiferido = ValorICMSUnitario -  ValorICMSDIF
                        ValorICMSUnitario = replace(ValorDiferido,",",".")
                    end if 
                    ValorBaseIcmsUnitario  = replace(ValorBaseIcmsUnitario,",",".")
            end if 

            if CalculaPISCofins <> "N" then
               PercentualPIS = 1.65
               PercentualCOFINS = 7.60
               if CalculaPISCofins = "P"  or CalculaPISCofins = "A" then 
                    ValorPISUnitario = replace((valorunitario - replace(ValorICMSUnitario,".",",")) * (replace(PercentualPIS,".",",")/100),",",".")
               end if 

               if  CalculaPISCofins = "C" or CalculaPISCofins = "A" then 
                   ValorCOFINSUnitario = replace((valorunitario - replace(ValorICMSUnitario,".",",")) * (replace(PercentualCOFINS,".",",")/100),",",".")
               end if 

               PercentualPIS = replace(PercentualPIS,",",".")
               PercentualCOFINS = replace(PercentualCOFINS,",",".")
            end if 

            if validarLimiteCredito = "S" then 
                ChamaLimiteCreditoCliente

                if tolerancia > 0 then 
                    quantidade = quantidade + (quantidade * (tolerancia /100))
                end if     

                SaldoDisponivel = SaldoDisponivel - ((quantidade * valorunitario) +  (quantidade * replace(ValorIPIUnitario,".",",")))
                if SaldoDisponivel < 0 then
                        GerouArrayMensagens = True
                        redim preserve arraymensagem(ContadorArrayMensagens)
                        arraymensagem(ContadorArrayMensagens - 1) =  "Cliente não possui Limite suficiente para inclusão do Pedido para o Item " & codproduto      '& " Saldo " & SaldoDisponivel & " "
                        ContadorArrayMensagens = uBound(arraymensagem)
                end if
            else 
                SaldoDisponivel=9999
            end if 
            
            if DataDesejadaDisponivel = 1 and SaldoDisponivel >= 0 then 
                    UltimoPedidoBento()
                    SaldoDisponivel = 0
                    FlagGravou=true
                
                    SaldoDisponivel = 0
                    FlagGravou=true
                   
                    axtotalpedido = quantidade * valorunitario
                   
                    linhaAtual = replace(linhaAtual,"nroPedido",nroPedidoBento) 
                    linhaatual = replace(linhaAtual,"Representante",Representante)
                    linhaatual = replace(linhaAtual,"Estrutura",Estrutura)
                    linhaatual = replace(linhaAtual,"TabelaComissao",TabelaComissao)
                    linhaatual = replace(linhaAtual,"ZonaVenda",ZonaVenda)
                    linhaatual = replace(linhaAtual,"Gerado",GeradoRepresentante)
                    linhaatual = replace(linhaAtual,"CST_IPI",CST_IPI)
                    linhaatual = replace(linhaAtual,"Peso",Peso)
                    linhaatual = replace(linhaAtual,"ValorIPIUnitario",ValorIPIUnitario)
                    linhaatual = replace(linhaAtual,"CST_ICMS",CST_ICMS)
                    linhaatual = replace(linhaAtual,"ValorBaseICMSUnitario",ValorBaseIcmsUnitario)
                    linhaatual = replace(linhaAtual,"PercentualBaseICMS",PercentualBaseICMS)
                    linhaatual = replace(linhaAtual,"ValorICMSUnitario",ValorICMSUnitario)
                    linhaatual = replace(linhaAtual,"PercentualICMS",PercentualICMS) 
                    linhaatual = replace(linhaAtual,"PercentualBasePIS",PercentualBasePIS) 
                    linhaatual = replace(linhaAtual,"ValorPISUnitario",ValorPISUnitario)
                    linhaatual = replace(linhaAtual,"PercentualPIS",PercentualPIS)
                    linhaatual = replace(linhaAtual,"ValorCOFINSUnitario",ValorCOFINSUnitario)
                    linhaatual = replace(linhaAtual,"PercentualCOFINS",PercentualCOFINS) 
                    linhaatual = replace(linhaAtual,"PercentualBaseCOFINS",PercentualBaseCOFINS)  
                    linhaatual = replace(linhaAtual,"MarcarLotePiloto",MarcarLotePiloto)   
                    linhaatual = replace(linhaAtual,"CodigoEnquadramentoIPI",CodigoEnquadramentoIPI) 
                                
                    linha = linha + linhaatual + ";"

                    redim preserve arrayPedidosGravados(ContadorArrayPedido)
                    arrayPedidosGravados(ContadorArrayPedido - 1) =  nroPedidoBento
                    ContadorArrayPedido = ContadorArrayPedido + 1
                        
                    SET Consulta = Conecta( "BEN-CORPORATIVO", cstr(linhaatual) )
                    if not isobject(Consulta) or (Consulta IS NOTHING) then
                        response.end
                    end if

                    linha = ""       
                    if TipoItem <> "PA" then 
                            'Valida Pedido por Orçamento
                            SQL = "EXEC SP_VALIDA_PEDIDO_COM_ORCAMENTO '" & Planta & "'," & nroPedidoBento
                            SET PIR_CORPORATIVO = Conecta( "PIR-CORPORATIVO", SQL )
                            if not isobject(PIR_CORPORATIVO) or (PIR_CORPORATIVO IS NOTHING) then
                                response.end
                            end if
                            set PIR_CORPORATIVO = nothing 

                            'Valida Pedido por Orçamento
                            SQL = "SELECT QuemValidou from Pedido Where Pedido = " & nroPedidoBento  & " and QuemValidou = '999999' and DataValidacao is not null" 
                            SET BEN_CORPORATIVO = Conecta( "BEN-CORPORATIVO", SQL )
                            if not isobject(BEN_CORPORATIVO) or (BEN_CORPORATIVO IS NOTHING) then
                                response.end
                            end if

                            if not BEN_CORPORATIVO.eof then 
                                SQL = "EXEC SP_INCLUI_PEDIDO_TRIMBOX_BENTO " & nroPedidoBento & ",'1'"  
                                SET BEN_CORPORATIVO = Conecta( "BEN-CORPORATIVO", SQL )
                                if not isobject(BEN_CORPORATIVO) or (BEN_CORPORATIVO IS NOTHING) then
                                    response.end
                                end if
                            end if 
                            set BEN_CORPORATIVO = nothing 
                    end if 

                   CodProdutoLotePiloto = trim(replace(CodProduto,"'",""))
                   if AtualizarLotePilotoProduto = CodProdutoLotePiloto then 
                        if not AtualizarLotePiloto() then 
                           'response.write "Lote Piloto não foi atualizado"
                        end if 
                   end if     
            end if
        next
end if 

if Planta = "06" then  '************ Fortaleza

        listaDePedido=replace(listaDePedido,"|,","|")
        arrayPedido=split(listaDePedido,"|")

        linha=""
        DataDesejadaDisponivel = 1
        SaldoDisponivel = 0
        MensagemLimiteCredito = ""
        MensagemDataDisponivel = ""
        GerouArrayMensagens = false
        ContadorArrayMensagens = 1
        ContadorArrayPedido = 1
        FlagGravou=false
        nroPedido = 0
        nroPedidoFortaleza = ""
        Representante = "''"
        TabelaComissao = "''"
        ZonaVenda = ""
        GeradoRepresentante = ""
        CST_IPI ="''"
        CST_ICMS = "''"
        ContaContabil="''"
        CodigoEnquadramentoIPI="''"
        DepositoOriginal="''"
        CentroCusto="''"

            
        for contaPorPedido = 0 to UBound(arrayPedido)

            linhaAtual = arrayPedido(contaPorPedido)

            if trim(linhaAtual) = ""  then 
                exit for
            end if

            Estado = ""
            ValorPISUnitario = 0
            ValorCOFINSUnitario = 0
            ValorICMSUnitario = 0
            ValorBaseIcmsUnitario = 0
            axvaloricms = 0
            ValorICMSUnitario = 0

            PercentualCOFINS = 0
            PercentualPIS = 0 
            CST_IPI = 0
            CST_ICMS = 0
            PercentualBaseCOFINS = 0
            PercentualCOFINS= 0
            ValorCOFINSUnitario= 0
            PercentualPIS= 0
            PercentualBasePIS= 0
            PercentualBaseICMS = 0
            PercentualBaseIPI  = 0
            PercentualBaseICMSBanco = 0
            PercentualBaseIPIBanco = 0
            BaseTributacaoIPITes = 0
            BaseTributacaoIPIProduto = 0
            ValorIPIUnitario = 0
            PermiteAlterarIPI = "N"
            CodigoEnquadramentoIPI = 0

            ColunaAtual = split(linhaAtual,",")
            TipoItem = left(replace(ColunaAtual(75),"'" ,""),2)
            CodProduto = ColunaAtual(20) 
            quantidade = replace(ColunaAtual(32),",",".") 
            valorunitario = replace(ColunaAtual(33),".",",") 
            
            DataEntrega = ColunaAtual(4) 
            DataEntregaOriginal = mid(DataEntrega,8,10)

            'Busca o Representante no Cadastro do Cliente
            SQL = "SELECT COD_VENDEDOR, ESTADO_CLIENTE FROM VW_CLIENTES_PEDIDO WHERE COD_CLIENTE = '" & CodCliente & "'"
            SET Consulta = Conecta( "FOR-DADOSADV", SQL )
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if
            If not Consulta.eof then
               Representante = Consulta("COD_VENDEDOR") 
               Estado =  Consulta("ESTADO_CLIENTE") 
            end if 
            Consulta.close()

            'Tabela TES 
            SQL = "SELECT CALC_IPI, CALC_ICM, PIS_COF, ISENCAO_FECP, DIFERE_ICMS, RED_BASE_ICMS, PERC_DIF FROM VW_BI_TES WHERE CODIGO  = '" & TES & "'" 
            SET Consulta = Conecta( "FOR-DADOSADV", SQL )
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if
            CalculaIPI = "N"
            CalculaICMS = "N"
            CalculaPISCofins = "N"
            IsentoFECP = "N"
            DifereICMS = "N"
            If not Consulta.eof then
                    PercentualBaseICMS = iif(Consulta("RED_BASE_ICMS")=0, 100,replace(Consulta("RED_BASE_ICMS"),",","."))
                    CalculaIPI  = Consulta("CALC_IPI")
                    CalculaICMS = Consulta("CALC_ICM")
                    DifereICMS =  Consulta("DIFERE_ICMS")
                    if Consulta("PIS_COF") = "NENHUM" then 
                        CalculaPISCofins = "N"
                    end if 
                    if Consulta("PIS_COF") = "PIS" then 
                        CalculaPISCofins = "P"
                    end if 
                    if Consulta("PIS_COF") = "COFINS" then 
                        CalculaPISCofins = "C"
                    end if 
                    if Consulta("PIS_COF") = "AMBOS" then 
                        CalculaPISCofins = "A"
                    end if 
                    PercentualDIF = Consulta("PERC_DIF")
                    IsentoFECP = Consulta("ISENCAO_FECP")
            end if 
            Consulta.close()
            
            if TipoItem = "CJ" then 
                SQL = "SELECT TOP 1 ALIQUOTA_IPI, ESTRUTURA, PESO FROM VW_BI_PRODUTOS_PEDIDO WHERE CODIGO_CONJUNTO = " & CodProduto & " ORDER BY CODIGO_ITEM "
            else     
                SQL = "SELECT TOP 1 ALIQUOTA_IPI, ESTRUTURA, PESO FROM VW_BI_PRODUTOS_PEDIDO WHERE CODIGO_ITEM = " & CodProduto 
            end if  
 
            SET Consulta = Conecta( "FOR-TRIMBOX", SQL )
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if
            PercentualIPI=0
            Peso = 0
            Estrutura =""

            if not Consulta.eof then 
                PercentualIPI = Consulta("ALIQUOTA_IPI")  
                Estrutura = Consulta("ESTRUTURA") 
                Peso= replace(Consulta("PESO"),",",".")
            end if  
            Consulta.close()

            SQL = "SELECT dbo.FUNC_PERCENTUAL_ICMS('FORTALEZA','" & Estado & "') as AliquotaICMS" 
            SET Consulta = Conecta( "FOR-CORPORATIVO", SQL )
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if
            PercentualICMS = 0 
            if not Consulta.eof then 
                PercentualICMS =  Consulta("AliquotaICMS")  
                If Estado = "RJ" and IsentoFECP = "N" then
                PercentualICMS = PercentualICMS + 2
                end if 
            end if  
            Consulta.close()

            'Calculo dos Impostos
            if CalculaIPI = "S" then 
                    ValorBaseIpiUnitario = replace(valorunitario,",",".")
                    ValorIPIUnitario =  replace(valorunitario * (replace(PercentualIPI,".",",")/100),",",".")
            end if 
            if CalculaICMS = "S" AND DifereICMS = "N" then 
                    ValorBaseIcmsUnitario = replace(valorunitario,",",".")
                    ValorICMSUnitario = replace(valorunitario * (replace(PercentualICMS,".",",")/100),",",".")
            end if 
            if CalculaICMS = "S" AND DifereICMS = "S"  then 
                    if PercentualDIF < 100 then 
                        ValorBaseIcmsUnitario = valorunitario
                        
                        ValorICMSUnitario = ValorBaseIcmsUnitario * (replace(PercentualICMS,".",",")/100)
                        ValorICMSDIF = ValorICMSUnitario * (replace(PercentualDIF,".",",")/100)
                        ValorDiferido = ValorICMSUnitario -  ValorICMSDIF
                        ValorICMSUnitario = replace(ValorDiferido,",",".")

                    end if 
                    ValorBaseIcmsUnitario  = replace(ValorBaseIcmsUnitario,",",".")
            end if 
            if CalculaPISCofins <> "N" then
                    PercentualPIS = 1.65
                    PercentualCOFINS = 7.60
                    if CalculaPISCofins = "P"  or CalculaPISCofins = "A" then 
                            ValorPISUnitario = replace((valorunitario - replace(ValorICMSUnitario,".",",")) * (replace(PercentualPIS,".",",")/100),",",".")
                    end if 
                    if  CalculaPISCofins = "C" or CalculaPISCofins = "A" then 
                        ValorCOFINSUnitario = replace((valorunitario - replace(ValorICMSUnitario,".",",")) * (replace(PercentualCOFINS,".",",")/100),",",".")
                    end if 
                    PercentualPIS = replace(PercentualPIS,",",".")
                    PercentualCOFINS = replace(PercentualCOFINS,",",".")

            end if 
            
            if validarLimiteCredito = "S" then 
                ChamaLimiteCreditoCliente 

                if tolerancia > 0 then 
                    quantidade = quantidade + (quantidade * (tolerancia /100))
                end if     

                'SaldoDisponivel = SaldoDisponivel - (quantidade * valorunitario)
                 SaldoDisponivel = SaldoDisponivel - ((quantidade * replace(valorunitario,".",",")) + (quantidade * replace(ValorIPIUnitario,".",",")))

                 'response.write "PercentualPIS  "  & PercentualPIS   &  "PercentualCOFINS  " &  PercentualCOFINS   &  "   ValorICMSUnitario " & ValorICMSUnitario & "   ValorIPIUnitario " & ValorIPIUnitario
                if SaldoDisponivel < 0 then
                    GerouArrayMensagens = True
                    redim preserve arraymensagem(ContadorArrayMensagens)
                    arraymensagem(ContadorArrayMensagens - 1) =  "Cliente não possui Limite suficiente para inclusão do Pedido para o Item " & codproduto           '& " Saldo " & SaldoDisponivel & " "
                    ContadorArrayMensagens = uBound(arraymensagem)
                end if    
            else 
                SaldoDisponivel=9999
            end if
            
            'SOLICITAÇÃO PARA NÃO VALIDAR CARGA MÁQUINA CHAMADO I2403-641
            'CargaMaquinaSaquarema codproduto,largura,comprimento,estrutura,nulo,nulo,Tipofrete,CodCliente ,clienteRemessa,quantidade,DataEntregaOriginal
            'if CargaMaquinaSaquarema = 0 then
            '        GerouArrayMensagens = True
            '        redim preserve arraymensagem(ContadorArrayMensagens)
            '        arraymensagem(ContadorArrayMensagens - 1) =  "Data de Entrega Indisponivel para " & DataEntregaOriginal  & "  Produto " & CodProduto
            '        ContadorArrayMensagens = uBound(arraymensagem)
            'end if
            'response.write "data disponivel " & DataDesejadaDisponivel  & "   SaldoDisponivel " &  SaldoDisponivel & " Data Original " & DataEntregaOriginal
            'response.end

        'DataDesejadaDisponivel=1
        'SaldoDisponivel=1
        if DataDesejadaDisponivel = 1 and SaldoDisponivel >= 0 then         
            nroPedidoFortaleza = UltimoPedidoFortaleza()
            
            'SOLICITAÇÃO PARA NÃO VALIDAR CARGA MÁQUINA CHAMADO I2403-641
            'DataDesejadaDisponivel = 0
            SaldoDisponivel = 0
            FlagGravou=true
        
            axtotalpedido = quantidade * valorunitario
        
            linhaAtual = replace(linhaAtual,"nroPedido",nroPedidoFortaleza) 
            linhaatual = replace(linhaAtual,"Representante",Representante)
            linhaatual = replace(linhaAtual,"Estrutura",Estrutura)
            linhaatual = replace(linhaAtual,"TabelaComissao",TabelaComissao)
            linhaatual = replace(linhaAtual,"ZonaVenda",ZonaVenda)
            linhaatual = replace(linhaAtual,"Gerado",GeradoRepresentante)
            linhaatual = replace(linhaAtual,"CST_IPI",CST_IPI)
            linhaatual = replace(linhaAtual,"Peso",Peso)
            linhaatual = replace(linhaAtual,"ValorIPIUnitario",ValorIPIUnitario)
            linhaatual = replace(linhaAtual,"CST_ICMS",CST_ICMS)
            linhaatual = replace(linhaAtual,"ValorBaseICMSUnitario",ValorBaseIcmsUnitario)
            linhaatual = replace(linhaAtual,"PercentualBaseICMS",PercentualBaseICMSBanco)
            linhaatual = replace(linhaAtual,"ValorICMSUnitario",ValorICMSUnitario)
            linhaatual = replace(linhaAtual,"PercentualICMS",PercentualICMS) 
            linhaatual = replace(linhaAtual,"PercentualBasePIS",PercentualBasePIS) 
            linhaatual = replace(linhaAtual,"ValorPISUnitario",ValorPISUnitario)
            linhaatual = replace(linhaAtual,"PercentualPIS",PercentualPIS)
            linhaatual = replace(linhaAtual,"ValorCOFINSUnitario",ValorCOFINSUnitario)
            linhaatual = replace(linhaAtual,"PercentualCOFINS",PercentualCOFINS) 
            linhaatual = replace(linhaAtual,"PercentualBaseCOFINS",PercentualBaseCOFINS)  
            linhaatual = replace(linhaAtual,"MarcarLotePiloto",MarcarLotePiloto)   
            linhaatual = replace(linhaAtual,"CodigoEnquadramentoIPI",CodigoEnquadramentoIPI) 
            linhaAtual = replace(linhaAtual,"nroPedido",nroPedido) 
                        
            linha = linha + linhaatual + ";"
          
            'response.write linhaatual
            
            redim preserve arrayPedidosGravados(ContadorArrayPedido)
            arrayPedidosGravados(ContadorArrayPedido - 1) =  nroPedidoFortaleza
            ContadorArrayPedido = ContadorArrayPedido + 1
            
            SET Consulta = Conecta( "FOR-CORPORATIVO", cstr(linhaatual) )
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if
            linha = ""

            'Valida Pedido por Orçamento
            SQL = "EXEC SP_VALIDA_PEDIDO_COM_ORCAMENTO '" & Planta & "'," & nroPedidoFortaleza 
            SET PIR_CORPORATIVO = Conecta( "PIR-CORPORATIVO", SQL )
            if not isobject(PIR_CORPORATIVO) or (PIR_CORPORATIVO IS NOTHING) then
                response.end
            end if
            set PIR_CORPORATIVO = nothing 

            'Valida Pedido por Orçamento
             SQL = "SELECT QuemValidou from Pedido Where Pedido = " & nroPedidoFortaleza  & " and QuemValidou = '999999' and DataValidacao is not null" 
             SET FOR_CORPORATIVO = Conecta( "FOR-CORPORATIVO", SQL )
             if not isobject(FOR_CORPORATIVO) or (FOR_CORPORATIVO IS NOTHING) then
                 response.end
             end if
             if not FOR_CORPORATIVO.eof then 
                SQL = "EXEC SP_INCLUI_PEDIDO_TRIMBOX_MARANGUAPE " & nroPedidoFortaleza & ",'1'"  
                SET FOR_CORPORATIVO = Conecta( "FOR-CORPORATIVO", SQL )
                if not isobject(FOR_CORPORATIVO) or (FOR_CORPORATIVO IS NOTHING) then
                    response.end
                end if
             end if 
             set FOR_CORPORATIVO = nothing 
             
             CodProdutoLotePiloto = trim(replace(CodProduto,"'",""))
             if AtualizarLotePilotoProduto = CodProdutoLotePiloto then 
                if not AtualizarLotePiloto() then 
                        'response.write "Lote Piloto não foi atualizado"
                end if 
             end if    
        end if
    next
end if

if Planta = "07" then  '************ Saquarema        

        listaDePedido=replace(listaDePedido,"|,","|")
        arrayPedido=split(listaDePedido,"|")

        linha=""
        DataDesejadaDisponivel = 1
        SaldoDisponivel = 0
        MensagemLimiteCredito = ""
        MensagemDataDisponivel = ""
        GerouArrayMensagens = false
        ContadorArrayMensagens = 1
        ContadorArrayPedido = 1
        FlagGravou=false
        nroPedido = 0
        nroPedidoSaquarema = ""
        Representante = "''"
        TabelaComissao = "''"
        ZonaVenda = ""
        GeradoRepresentante = ""
        CST_IPI ="''"
        CST_ICMS = "''"
        ContaContabil="''"
        CodigoEnquadramentoIPI="''"
        DepositoOriginal="''"
        CentroCusto="''"

            
        for contaPorPedido = 0 to UBound(arrayPedido)

            linhaAtual = arrayPedido(contaPorPedido)
       
            if trim(linhaAtual) = ""  then 
                 exit for
            end if

            Estado = ""
            ValorPISUnitario = 0
            ValorCOFINSUnitario = 0
            ValorICMSUnitario = 0
            ValorBaseIcmsUnitario = 0
            axvaloricms = 0
            ValorICMSUnitario = 0

            PercentualCOFINS = 0
            PercentualPIS = 0 
            CST_IPI = 0
            CST_ICMS = 0
            PercentualBaseCOFINS = 0
            PercentualCOFINS= 0
            ValorCOFINSUnitario= 0
            PercentualPIS= 0
            PercentualBasePIS= 0
            PercentualBaseICMS = 0
            PercentualBaseIPI  = 0
            PercentualBaseICMSBanco = 0
            PercentualBaseIPIBanco = 0
            BaseTributacaoIPITes = 0
            BaseTributacaoIPIProduto = 0
            ValorIPIUnitario = 0
            PermiteAlterarIPI = "N"
            CodigoEnquadramentoIPI = 0

            ColunaAtual = split(linhaAtual,",")
            TipoItem = left(replace(ColunaAtual(75),"'" ,""),2)
            CodProduto = ColunaAtual(20) 
            quantidade = replace(ColunaAtual(32),",",".") 
            valorunitario = replace(ColunaAtual(33),".",",") 
            larguraChapa = ColunaAtual(22) 
            comprimentoChapa = ColunaAtual(23) 
            estruturaChapa = ColunaAtual(69)
            
            DataEntrega = ColunaAtual(4) 
            DataEntregaOriginal = mid(DataEntrega,8,10)
            'Busca o Representante no Cadastro do Cliente
            SQL = "SELECT COD_VENDEDOR, ESTADO_CLIENTE FROM VW_BI_CLIENTES WHERE COD_CLIENTE = '" & CodCliente & "' AND LOJA_CLIENTE = '" & CodLoja & "'"
            'Dados161.Source = SQL
            'Dados161.open()
            
            SET Consulta = Conecta( "SAQ-DADOSADV", SQL )
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if
            If not Consulta.eof then
               Representante = Consulta("COD_VENDEDOR") 
               Estado =  Consulta("ESTADO_CLIENTE") 
            end if 
            Consulta.close()
    
            'Tabela TES 
            SQL = "SELECT CALC_IPI, CALC_ICM, PIS_COF, ISENCAO_FECP, DIFERE_ICMS, RED_BASE_ICMS, PERC_DIF FROM VW_BI_TES WHERE CODIGO  = '" & TES & "'" 
            SET Consulta = Conecta( "SAQ-DADOSADV", SQL )
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if
            CalculaIPI = "N"
            CalculaICMS = "N"
            CalculaPISCofins = "N"
            IsentoFECP = "N"
            DifereICMS = "N"
            If not Consulta.eof then
               PercentualBaseICMS = iif(Consulta("RED_BASE_ICMS")=0, 100,replace(Consulta("RED_BASE_ICMS"),",","."))
               CalculaIPI  = Consulta("CALC_IPI")
               CalculaICMS = Consulta("CALC_ICM")
               DifereICMS =  Consulta("DIFERE_ICMS")
               if Consulta("PIS_COF") = "NENHUM" then 
                   CalculaPISCofins = "N"
               end if 
               if Consulta("PIS_COF") = "PIS" then 
                   CalculaPISCofins = "P"
               end if 
               if Consulta("PIS_COF") = "COFINS" then 
                  CalculaPISCofins = "C"
               end if 
               if Consulta("PIS_COF") = "AMBOS" then 
                  CalculaPISCofins = "A"
               end if 
               PercentualDIF = Consulta("PERC_DIF")
               IsentoFECP = Consulta("ISENCAO_FECP")
            end if 
            Consulta.close()
            
            if TipoItem = "CJ" then 
                 SQL = "SELECT TOP 1 ALIQUOTA_IPI, ESTRUTURA, PESO FROM VW_BI_PRODUTOS_PEDIDO WHERE CODIGO_CONJUNTO = " & CodProduto & " ORDER BY CODIGO_ITEM "
            else     
                 SQL = "SELECT TOP 1 ALIQUOTA_IPI, ESTRUTURA, PESO FROM VW_BI_PRODUTOS_PEDIDO WHERE CODIGO_ITEM = " & CodProduto 
            end if  

            SET Consulta = Conecta( "SAQ-TRIMBOX", SQL )
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if
            PercentualIPI=0
            Peso = 0
            Estrutura =""
            if not Consulta.eof then 
                PercentualIPI = Consulta("ALIQUOTA_IPI")  
                Estrutura = Consulta("ESTRUTURA") 
                Peso= replace(Consulta("PESO"),",",".")
            end if  
            Consulta.close()

            'Recalcula o Peso no caso de CHAPA 
            if TipoItem = "CH" then 
                    if larguraChapa > 0  and comprimentoChapa > 0 then 
                        SQL = " SELECT PESO FROM dbo.FUNC_CALCULA_AREA_PESO(" & larguraChapa & ", " & comprimentoChapa & ", '" & Estrutura & "')"
                        SET SAQ_CORPORATIVO = Conecta( "SAQ-CORPORATIVO", SQL )
                        if not isobject(SAQ_CORPORATIVO) or (SAQ_CORPORATIVO IS NOTHING) then
                            response.end
                        end if
                        if not SAQ_CORPORATIVO.eof then
                            Peso =  replace(SAQ_CORPORATIVO("PESO"),",",".")
                        end if 
                    end if 
            end if 

            SQL = "SELECT dbo.FUNC_PERCENTUAL_ICMS('SAQUAREMA','" & Estado & "') as AliquotaICMS" 
            SET Consulta = Conecta( "SAQ-CORPORATIVO", SQL )
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if
            PercentualICMS = 0 
            if not Consulta.eof then 
                PercentualICMS =  Consulta("AliquotaICMS")  
                If Estado = "RJ" and IsentoFECP = "N" then
                   PercentualICMS = PercentualICMS + 2
                end if 
            end if  
            Consulta.close()

            'Calculo dos Impostos
            if CalculaIPI = "S" then 
               ValorBaseIpiUnitario = replace(valorunitario,",",".")
               ValorIPIUnitario =  replace(valorunitario * (replace(PercentualIPI,".",",")/100),",",".")
            end if 
            if CalculaICMS = "S" AND DifereICMS = "N" then 
                ValorBaseIcmsUnitario = replace(valorunitario,",",".")
                ValorICMSUnitario = replace(valorunitario * (replace(PercentualICMS,".",",")/100),",",".")
            end if 
            if CalculaICMS = "S" AND DifereICMS = "N"  then 
               if PercentualBaseICMS = 100 then 
                    ValorBaseIcmsUnitario = valorunitario
                    
                    ValorICMSUnitario = replace(ValorBaseIcmsUnitario * (replace(PercentualICMS,".",",")/100),",",".")
                else 
                    ValorBaseIcmsUnitario = valorunitario * (replace(PercentualBaseICMS,".",",")/100)
                    
                    ValorICMSUnitario = replace(ValorBaseIcmsUnitario * (replace(PercentualICMS,".",",")/100),",",".")
                end if   
                ValorBaseIcmsUnitario  = replace(ValorBaseIcmsUnitario,",",".")
            end if 


            if CalculaPISCofins <> "N" then
               PercentualPIS = 1.65
               PercentualCOFINS = 7.60
               if CalculaPISCofins = "P"  or CalculaPISCofins = "A" then 
                    ValorPISUnitario = replace((valorunitario - replace(ValorICMSUnitario,".",",")) * (replace(PercentualPIS,".",",")/100),",",".")
               end if 
               if  CalculaPISCofins = "C" or CalculaPISCofins = "A" then 
                   ValorCOFINSUnitario = replace((valorunitario - replace(ValorICMSUnitario,".",",")) * (replace(PercentualCOFINS,".",",")/100),",",".")
               end if 
               PercentualPIS = replace(PercentualPIS,",",".")
               PercentualCOFINS = replace(PercentualCOFINS,",",".")

            end if 

            if validarLimiteCredito = "S" then 
               ChamaLimiteCreditoCliente

               if tolerancia > 0 then 
                   quantidade = quantidade + (quantidade * (tolerancia /100))
               end if     
 
               'SaldoDisponivel = SaldoDisponivel - (quantidade * valorunitario)
                SaldoDisponivel = SaldoDisponivel - ((quantidade * valorunitario) + (quantidade * replace(ValorIPIUnitario,".",",")))

                'response.write "PercentualPIS  "  & PercentualPIS   &  "PercentualCOFINS  " &  PercentualCOFINS   &  "   ValorICMSUnitario " & ValorICMSUnitario & "   ValorIPIUnitario " & ValorIPIUnitario
                if SaldoDisponivel < 0 then
                    GerouArrayMensagens = True
                    redim preserve arraymensagem(ContadorArrayMensagens)
                    arraymensagem(ContadorArrayMensagens - 1) =  "Cliente não possui Limite suficiente para inclusão do Pedido para o Item " & codproduto      '& " Saldo " & SaldoDisponivel & " "
                    ContadorArrayMensagens = uBound(arraymensagem)
                end if
            else 
                SaldoDisponivel=9999
            end if     
            
            'SOLICITAÇÃO PARA NÃO VALIDAR CARGA MÁQUINA CHAMADO I2403-641
            'CargaMaquinaSaquarema codproduto,largura,comprimento,estrutura,nulo,nulo,Tipofrete,CodCliente ,clienteRemessa,quantidade,DataEntregaOriginal
            'if CargaMaquinaSaquarema = 0 then
            '        GerouArrayMensagens = True
            '        redim preserve arraymensagem(ContadorArrayMensagens)
            '        arraymensagem(ContadorArrayMensagens - 1) =  "Data de Entrega Indisponivel para " & DataEntregaOriginal  & "  Produto " & CodProduto
            '        ContadorArrayMensagens = uBound(arraymensagem)
            'end if
            'response.write "data disponivel " & DataDesejadaDisponivel  & "   SaldoDisponivel " &  SaldoDisponivel & " Data Original " & DataEntregaOriginal
            'response.end

        'DataDesejadaDisponivel=1
        'SaldoDisponivel=1
        if DataDesejadaDisponivel = 1 and SaldoDisponivel >= 0 then 
           
            UltimoPedidoSaquarema()
            
            'SOLICITAÇÃO PARA NÃO VALIDAR CARGA MÁQUINA CHAMADO I2403-641
            'DataDesejadaDisponivel = 0
            SaldoDisponivel = 0
            FlagGravou=true
           
            axtotalpedido = quantidade * valorunitario
           
            linhaAtual = replace(linhaAtual,"nroPedido",nroPedidoSaquarema) 
            linhaatual = replace(linhaAtual,"Representante",Representante)
            linhaatual = replace(linhaAtual,"Estrutura",Estrutura)
            linhaatual = replace(linhaAtual,"TabelaComissao",TabelaComissao)
            linhaatual = replace(linhaAtual,"ZonaVenda",ZonaVenda)
            linhaatual = replace(linhaAtual,"Gerado",GeradoRepresentante)
            linhaatual = replace(linhaAtual,"CST_IPI",CST_IPI)
            linhaatual = replace(linhaAtual,"Peso",Peso)
            linhaatual = replace(linhaAtual,"ValorIPIUnitario",ValorIPIUnitario)
            linhaatual = replace(linhaAtual,"CST_ICMS",CST_ICMS)
            linhaatual = replace(linhaAtual,"ValorBaseICMSUnitario",ValorBaseIcmsUnitario)
            linhaatual = replace(linhaAtual,"PercentualBaseICMS",PercentualBaseICMSBanco)
            linhaatual = replace(linhaAtual,"ValorICMSUnitario",ValorICMSUnitario)
            linhaatual = replace(linhaAtual,"PercentualICMS",PercentualICMS) 
            linhaatual = replace(linhaAtual,"PercentualBasePIS",PercentualBasePIS) 
            linhaatual = replace(linhaAtual,"ValorPISUnitario",ValorPISUnitario)
            linhaatual = replace(linhaAtual,"PercentualPIS",PercentualPIS)
            linhaatual = replace(linhaAtual,"ValorCOFINSUnitario",ValorCOFINSUnitario)
            linhaatual = replace(linhaAtual,"PercentualCOFINS",PercentualCOFINS) 
            linhaatual = replace(linhaAtual,"PercentualBaseCOFINS",PercentualBaseCOFINS)  
            linhaatual = replace(linhaAtual,"MarcarLotePiloto",MarcarLotePiloto)   
           'linhaatual = replace(linhaAtual,"CodigoOpereracaoContaContabil",ContaContabil) 
            linhaatual = replace(linhaAtual,"CodigoEnquadramentoIPI",CodigoEnquadramentoIPI) 
           'linhaatual = replace(linhaAtual,"DepositoOriginal",DepositoOriginal) 
           'linhaatual = replace(linhaAtual,"CentroCusto",CentroCusto) 
                        
            linha = linha + linhaatual + ";"
            
            redim preserve arrayPedidosGravados(ContadorArrayPedido)
            arrayPedidosGravados(ContadorArrayPedido - 1) =  nroPedidoSaquarema
            ContadorArrayPedido = ContadorArrayPedido + 1

            SET Consulta = Conecta( "SAQ-CORPORATIVO", cstr(linhaatual) )
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if
            linha = ""

            'Valida Pedido por Orçamento
            SQL = "EXEC SP_VALIDA_PEDIDO_COM_ORCAMENTO '" & Planta & "'," & nroPedidoSaquarema 
            SET PIR_CORPORATIVO = Conecta( "PIR-CORPORATIVO", SQL )
            if not isobject(PIR_CORPORATIVO) or (PIR_CORPORATIVO IS NOTHING) then
                response.end
            end if
            set PIR_CORPORATIVO = nothing 

            'Valida Pedido por Orçamento
            SQL = "SELECT QuemValidou from Pedido Where Pedido = " & nroPedidoSaquarema  & " and QuemValidou = '999999' and DataValidacao is not null" 
            SET SAQ_CORPORATIVO = Conecta( "SAQ-CORPORATIVO", SQL )
            if not isobject(SAQ_CORPORATIVO) or (SAQ_CORPORATIVO IS NOTHING) then
                 response.end
            end if
            if not SAQ_CORPORATIVO.eof then 
                SQL = "EXEC SP_INCLUI_PEDIDO_TRIMBOX_SAQUAREMA " & nroPedidoSaquarema & ",'1'"  
                SET SAQ_CORPORATIVO = Conecta( "SAQ-CORPORATIVO", SQL )
                if not isobject(SAQ_CORPORATIVO) or (SAQ_CORPORATIVO IS NOTHING) then
                    response.end
                end if
            end if 
            set SAQ_CORPORATIVO = nothing 
           
            CodProdutoLotePiloto = trim(replace(CodProduto,"'",""))
            if AtualizarLotePilotoProduto = CodProdutoLotePiloto then 
                if not AtualizarLotePiloto() then 
                   'response.write "Lote Piloto não foi atualizado"
                end if 
            end if   
        end if
    next
end if

if FlagGravou  then
    'SET Consulta = Conecta( "BEN-CORPORATIVO", cstr(linha) )
    'if not isobject(Consulta) then
    '     response.end
    'end if
  
    if isarray(arrayPedidosGravados) then
       if ubound(arrayPedidosGravados) > 0 then
          for i = 0 to ubound(arrayPedidosGravados)-1
              call ExtratoPedido(Planta, arrayPedidosGravados(i), DataPrevistaEntrega)
          next        
        end if
    end if
 end if

if GerouArrayMensagens then
    MensagemAlerta = ""
    if isarray(arraymensagem) then
        if ubound(arraymensagem) > 0 then
            for i = 0 to ubound(arraymensagem)-1
                MensagemAlerta =  MensagemAlerta & "- " & replace(arraymensagem(i), "'","") & "\n"
            next
        end if
    end if     
    if trim(MensagemAlerta) <> "" then
            response.write "<script type=""text/javascript"">alert('" + MensagemAlerta + "')</script>"
            response.end()            
    end if
end if
 
response.write "<script type=""text/javascript"" charset=""utf-8"">"
response.write "alert('Pedido incluido com sucesso');"
response.write "location.href = window.location.protocol + '//' + window.location.hostname +'/corporativo/comercial/Brasil/pedido_digitado/Pedidodigitado.asp?Planta="+Planta+"'"
response.write "</script>"
response.end()

'Ultimo Pedido Bento
Function UltimoPedidoBento()
        
        SQL = "SELECT dbo.FUNC_NOVO_NUM_PEDIDO_LETRAS() as PedidoBento "
        SET Consulta = Conecta( "BEN-DADOSADV", SQL)
        if not isobject(Consulta) or (Consulta IS NOTHING) then
             response.end
        end if
        nroPedidoBento = 0
        if not Consulta.eof then 
           nroPedidoBento = "'" & Consulta("PedidoBento") & "'" 
        end if 
        Consulta.close 
       
end function

'Ultimo Pedido Saquarema
Function UltimoPedidoSaquarema()
    SQL = "SELECT dbo.FUNC_NOVO_NUM_PEDIDO_LETRAS() as PedidoSaquarema"
    SET Consulta = Conecta( "SAQ-DADOSADV", SQL)
    
    if not isobject(Consulta) or (Consulta IS NOTHING) then
        response.end
    end if

    if not Consulta.eof then 
        nroPedidoSaquarema = "'" & Consulta("PedidoSaquarema")  & "'"
    end if   

    Consulta.close() 
end function

'Ultimo Pedido Fortaleza
Function UltimoPedidoFortaleza()
    SQL = "SELECT dbo.FUNC_NOVO_NUM_PEDIDO_LETRAS() as PedidoFortaleza "
    SET Consulta = Conecta( "FOR-DADOSADV", SQL)
    if not isobject(Consulta) or (Consulta IS NOTHING) then
        response.end
    end if
    
    if not Consulta.eof then 
        UltimoPedidoFortaleza =  "'" & Consulta("PedidoFortaleza") & "'"
    else
        UltimoPedidoFortaleza = 0
    end if  

    Consulta.close() 
end function

'// Verifica Credito do Cliente
function LimiteCreditoCliente(byPassaLimite)
         if byPassaLimite <> 1 then
           
            if Planta = "05" or  Planta = "06" then  
               'SQL = "exec [spLimiteCredito] '" & CodCliente & "'" 
                SQL = "SELECT dbo.FUNC_CALCULA_LIMITE_CREDITO('"& CodCliente &"') as Saldo"
                if Planta = "05" then 
                    Set Consulta = Conecta( "BEN-DADOSADV", SQL )
                else 
                    Set Consulta = Conecta( "FOR-DADOSADV", SQL )
                end if     
            else
                SQL = "SELECT dbo.FUNC_CALCULA_LIMITE_CREDITO('"& CodCliente &"', '"& CodLoja &"') as Saldo"
                Set Consulta = Conecta( "SAQ-DADOSADV", SQL )
            end if 
            if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
            end if
            
            if not Consulta.eof then
                SaldoDisponivel = FormatNumber(Consulta("Saldo"),2)
            Else
                SaldoDisponivel = 0
            end if
            Consulta.close
        else
           SaldoDisponivel = 9999999
        end if
end function

Function ChamaLimiteCreditoCliente()
    byPass = 0
   'matrizCliente = Array("76728","76946","75200","76927","76962","76527","76800","76930","75600","73910","74796","76722","76984","71242","2987","77020","77009","73693","76791","75353","77031")
    matrizCliente = Array("76728","00000607")
    for each a in matrizCliente
        if a = CodCliente then 
            byPass = 1
        end if
    next

    LimiteCreditoCliente byPass
End function

'//Carga Maquina
function CargaMaquina(produto, largura, comprimento, estrutura, resinaExterna, resinaInterna, tipoFrete, CodCliente, clienteRemessa, quantidade, dataEntrega)
     estrutura="NULL"
     comprimento="NULL"
     largura="NULL"
     'if resinaExterna="" then
        resinaExterna="NULL"
    'end if   

    'if resinaInterna="" then
        resinaInterna="NULL"
    'end if 
      
         SQL = "Select DataDesejadaDisponivel, DataPossivel FROM dbo.fq_PreviaPedidos_InfoDataPossivel_Paema(" 
         SQL = SQL & Trim(produto) 
         SQL = SQL & "," & Trim(largura) 
         SQL = SQL & "," & Trim(comprimento)
         if Trim(estrutura) = "NULL" then
            SQL = SQL & "," & Trim(estrutura) 
         else
            SQL = SQL & ",'" & Trim(estrutura)&"' "
         end if
         SQL = SQL & "," & Trim(resinaExterna)
         SQL = SQL & "," & Trim(resinaInterna)
         SQL = SQL & "," & Trim(tipoFrete)
         SQL = SQL & "," & Trim(CodCliente)
         SQL = SQL & "," & Trim(clienteRemessa)
         SQL = SQL & "," & Trim(quantidade)
         SQL = SQL & ",'" & Trim(dataEntrega) + "'"
         SQL = SQL & ")"
          
         SET Consulta = Conecta( "BEN-TRIMBOX", SQL )
         if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
         end if
         if Consulta.eof then
             DataDesejadaDisponivel =  0
         else     
            if not Consulta("DataDesejadaDisponivel")  then
                DataDesejadaDisponivel =  0
            else
                DataDesejadaDisponivel = 1
            end if
         end if
         Consulta.close
end function

'//Carga Maquina
function CargaMaquinaSaquarema(produto, largura, comprimento, estrutura, resinaExterna, resinaInterna, tipoFrete, CodCliente, clienteRemessa, quantidade, dataEntrega)
         estrutura="NULL"
         comprimento="NULL"
         largura="NULL"
    
      
         SQL = "Select DataDesejadaDisponivel, DataPossivel FROM dbo.fq_PreviaPedidos_InfoDataPossivel_PaperBox(" 
         SQL = SQL & Trim(produto) 
         SQL = SQL & "," & Trim(largura) 
         SQL = SQL & "," & Trim(comprimento)
         if Trim(estrutura) = "NULL" then
            SQL = SQL & "," & Trim(estrutura) 
         else
            SQL = SQL & ",'" & Trim(estrutura)&"' "
         end if
         SQL = SQL & "," & Trim(resinaExterna)
         SQL = SQL & "," & Trim(resinaInterna)
         SQL = SQL & "," & Trim(tipoFrete)
         SQL = SQL & ",'" & Trim(CodCliente)&CodLoja & "'"
         SQL = SQL & ",'" & Trim(clienteRemessa) & "'"
         SQL = SQL & "," & Trim(quantidade)
         SQL = SQL & ",'" & Trim(dataEntrega) + "'"
         SQL = SQL & ")"
          
         SET Consulta = Conecta( "SAQ-TRIMBOX", SQL )
         if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
         end if
         if Consulta.eof then
             DataDesejadaDisponivel =  0
         else     
            if not Consulta("DataDesejadaDisponivel")  then
                DataDesejadaDisponivel =  0
            else
                DataDesejadaDisponivel = 1
            end if
         end if
         Consulta.close
end function

'//Carga Maquina Maranguape 
function CargaMaquinaFortaleza(produto, largura, comprimento, estrutura, resinaExterna, resinaInterna, tipoFrete, CodCliente, clienteRemessa, quantidade, dataEntrega)
         estrutura="NULL"
         comprimento="NULL"
         largura="NULL"
      
         SQL = "Select DataDesejadaDisponivel, DataPossivel FROM dbo.fq_PreviaPedidos_InfoDataPossivel_SKMaranguape(" 
         SQL = SQL & Trim(produto) 
         SQL = SQL & "," & Trim(largura) 
         SQL = SQL & "," & Trim(comprimento)
         if Trim(estrutura) = "NULL" then
            SQL = SQL & "," & Trim(estrutura) 
         else
            SQL = SQL & ",'" & Trim(estrutura)&"' "
         end if
         SQL = SQL & "," & Trim(resinaExterna)
         SQL = SQL & "," & Trim(resinaInterna)
         SQL = SQL & "," & Trim(tipoFrete)
         SQL = SQL & ",'" & Trim(CodCliente)&CodLoja & "'"
         SQL = SQL & ",'" & Trim(clienteRemessa) & "'"
         SQL = SQL & "," & Trim(quantidade)
         SQL = SQL & ",'" & Trim(dataEntrega) + "'"
         SQL = SQL & ")"

         SET Consulta = Conecta( "FOR-TRIMBOX", SQL )
         if not isobject(Consulta) or (Consulta IS NOTHING) then
                response.end
         end if
         if Consulta.eof then
             DataDesejadaDisponivel =  0
         else     
            if not Consulta("DataDesejadaDisponivel")  then
                DataDesejadaDisponivel =  0
            else
                DataDesejadaDisponivel = 1
            end if
         end if
         Consulta.close
end function

function AtualizarLotePiloto()


        'Busca o status atual do lote piloto
        SQL = "SELECT TOP 1 ID, CASE STATUS "
        SQL = SQL & "   WHEN '1' THEN 'Aguardando Pedido LP' "
        SQL = SQL & "   WHEN '2' THEN 'Convertido em Venda'  "
        SQL = SQL & "	WHEN '3' THEN 'Aprovado sem Pedido de Conversao'  "
        SQL = SQL & "	WHEN '4' THEN 'Em Teste Cliente'  "
        SQL = SQL & "	WHEN '5' THEN 'Reprovado Pelo Cliente'  "
        SQL = SQL & "	WHEN '6' THEN 'Reprovado Internamente/Retrabalho'  " 
        SQL = SQL & "	WHEN '7' THEN 'Aguardando Producao'  "
        SQL = SQL & "	WHEN '8' THEN 'Descontinuado'  "
        SQL = SQL & "	WHEN '9' THEN 'Sem Retorno Vendas/Cliente'  "
        SQL = SQL & "	WHEN '10' THEN 'Processo segue em Novo Produto/Ficha Técnica' "
        SQL = SQL & "	WHEN '11' THEN 'Validacao Maquina Nova/Homologaca' END STATUS  "
        SQL = SQL & " FROM LOTE_PILOTO  "
        SQL = SQL & " WHERE CLIENTE = '" & CodCliente & "'"
        SQL = SQL & "   AND PLANTA =  '" & Planta & "'"
        SQL = SQL & "   AND APROVACAO = 0  "
        SQL = SQL & "   AND QUEM_EXCLUIU IS NULL " 
        SQL = SQL & "   AND PRODUTO = " & codproduto  
        SQL = SQL & " ORDER BY DATEDIFF(DAY,DATA_HORA,GETDATE()) " 

        SET rstTrazerLotePiloto = Conecta( "PIR-CORPORATIVO", SQL )
        if not isobject(rstTrazerLotePiloto) or (rstTrazerLotePiloto IS NOTHING) then
            response.end
        end if
        
        if rstTrazerLotePiloto.eof then 
            AtualizarLotePiloto = false 
            exit function
        end if    

        statusAnterior = rstTrazerLotePiloto("STATUS")

        'Atualiza o status do lote piloto para 7 - Aguardando Produção
        SQL = Request.Form("hidAtualizaLotePilotoSQL")
        if SQL = "" then 
            AtualizarLotePiloto = false 
            exit function
        end if

        SET rstAtualizaLotePiloto = Conecta( "PIR-CORPORATIVO", SQL )
        if not isobject(rstAtualizaLotePiloto) or (rstAtualizaLotePiloto IS NOTHING) then
            response.end
        end if
        set rstAtualizaLotePiloto = nothing
      
        ' Log o status na tabela de historico
        InsertHistoricoStatus = "INSERT INTO LOTE_PILOTO_HISTORICO_STATUS (IDLOTEPILOTO, STATUS_ANTIGO, STATUS_NOVO, ALTERADO_POR, DATA_INCLUSAO, USUARIO) "
        InsertHistoricoStatus = InsertHistoricoStatus & " VALUES ("
        InsertHistoricoStatus = InsertHistoricoStatus & rstTrazerLotePiloto("ID")
        InsertHistoricoStatus = InsertHistoricoStatus & ", '" & statusAnterior & "'"
        InsertHistoricoStatus = InsertHistoricoStatus & ", 'Aguardando Producao'"
        InsertHistoricoStatus = InsertHistoricoStatus & ",'PEDIDO'"
        InsertHistoricoStatus = InsertHistoricoStatus & ", getdate() "
        InsertHistoricoStatus = InsertHistoricoStatus & ", '" & Session("IdUsuario") & "')"
    
        SET rstHistoricoStatus = Conecta( "PIR-CORPORATIVO", InsertHistoricoStatus )
        if not isobject(rstHistoricoStatus) or (rstHistoricoStatus IS NOTHING) then
           response.end
        end if

        set rstHistoricoStatus=nothing
        AtualizarLotePiloto = true
end function 
%>