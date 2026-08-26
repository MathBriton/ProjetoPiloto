<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../../../cnn2/cnn2.asp" -->
<!--#include file="../../../../cnn2/conexoes.asp" -->
<!--#include file="../../../../funcoes/Funcoes.asp" -->
<!--#include file="../../../../includes/variaveis.asp" -->
<%
' =================================================================================================
' alterapedido04.asp  — Processador de alteração de pedido
' Opção A: ASP Clássico puro. Sem JSON/AJAX.
' Refatorado em 2025: guards consolidados, SanitizaStr centralizada,
' corpo do e-mail extraído para MontaMsgEmail(), sem duplicações.
' =================================================================================================

' -------------------------------------------------------------------------------------------------
' GUARDS DE SESSÃO (variaveis.asp já verifica e faz response.end; esta linha é defesa extra)
' -------------------------------------------------------------------------------------------------
if Trim(USR) = "" then
    response.end
end if

server.ScriptTimeout = 6000

' -------------------------------------------------------------------------------------------------
' HELPERS INLINE
' -------------------------------------------------------------------------------------------------
function SanitizaStr(s)
    SanitizaStr = Replace(Trim(s), "'", "")
end function

function Req(campo)
    Req = Trim(Request(campo))
end function

' -------------------------------------------------------------------------------------------------
' PARÂMETROS GLOBAIS DA REQUISIÇÃO
' -------------------------------------------------------------------------------------------------
Dim idSolicitacao, nomeCliente
Dim tipoOp, pedido, clienteCod, lojaCod
Dim de, para, novoTotal, valorFrete
Dim dataAlteracao, horaAlteracao

idSolicitacao = Req("IdSolicitacao")
pedido        = Req("Pedido")
clienteCod    = Req("Cliente")
lojaCod       = Req("Loja")
tipoOp        = Req("TipoOperacao")
de            = Req("De")
para          = Req("Para")
novoTotal     = Req("NovoTotal")
dataAlteracao = Req("DataAlteracao")
horaAlteracao = Req("HoraAlteracao")
valorFrete    = CDbl( iif(Req("valorfrete") = "", "0", Req("valorfrete")) )

nomeCliente   = Req("NomeCliente")
if Len(nomeCliente) > 40 then nomeCliente = Left(nomeCliente, 40)

' -------------------------------------------------------------------------------------------------
' TRANSAÇÃO SQL
' -------------------------------------------------------------------------------------------------
Dim IniciaTransacao, FinalizaTransacao
IniciaTransacao   = "SET XACT_ABORT ON; BEGIN TRANSACTION TRANS; "
FinalizaTransacao = "IF @@ERROR <> 0 ROLLBACK TRANSACTION TRANS ELSE COMMIT TRANSACTION TRANS; "

' -------------------------------------------------------------------------------------------------
' DESTINATÁRIOS PADRÃO
' -------------------------------------------------------------------------------------------------
Dim emailRep, emailDe, gerente, copia
emailRep = ""
emailDe  = "noreply.brasil@smurfitwestrock.com.br"

gerente = "comercial-rj@smurfitwestrock.com.br"
copia   = "andreia.genari@smurfitwestrock.com.br;" & _
          "jose.lacerda@smurfitwestrock.com.br;" & _
          "thiago.andrade@smurfitwestrock.com.br;" & _
          "sergiane.duarte@smurfitwestrock.com.br"

' -------------------------------------------------------------------------------------------------
' CONSULTA DADOS DO PEDIDO (representante, quantidade vendida, produto)
' -------------------------------------------------------------------------------------------------
Dim QuantVendida, Produto
QuantVendida = 0
Produto      = ""

Dim SQL_DADOS
SQL_DADOS = "" & _
    "SELECT " & _
        "A3_COD, A3_NREDUZ, " & _
        "A3_EMAIL, " & _
        "C6_QTDVEN, " & _
        "C6_PRODUTO " & _
    "FROM SC5010 SC5 " & _
        "INNER JOIN SC6010 SC6 ON " & _
                    "C5_NUM = C6_NUM " & _
                "AND ISNULL(SC6.D_E_L_E_T_,'') = '' " & _
        "INNER JOIN SA1010 SA1 ON " & _
                    "C5_CLIENTE = A1_COD " & _
                "AND C5_LOJACLI = A1_LOJA " & _
                "AND ISNULL(SA1.D_E_L_E_T_,'') = '' " & _
                "AND ISNULL(SC5.D_E_L_E_T_,'') = '' " & _
        "INNER JOIN SA3010 SA3 ON " & _
                    "SA1.A1_VEND = SA3.A3_COD " & _
                "AND ISNULL(SA3.D_E_L_E_T_,'') = '' " & _
        "LEFT JOIN HELPDESK..CRM_DTI AS USUARIO ON " & _
                    "SA3.A3_COD = USUARIO.CRM_COD COLLATE SQL_Latin1_General_CP437_BIN " & _
                "AND ISNULL(CRM_D_E_L_E_T,'') = '' " & _
    "WHERE C5_NUM = '" & pedido & "' "

Dim DADOS
SET DADOS = Conecta("PIR-DADOSADV", SQL_DADOS)
if not isObject(DADOS) or (DADOS IS NOTHING) then response.end
if not DADOS.EOF then
    QuantVendida = DADOS("C6_QTDVEN")
    Produto      = Trim(DADOS("C6_PRODUTO"))
    emailRep     = DADOS("A3_EMAIL")
end if
DADOS.Close() : SET DADOS = NOTHING

' -------------------------------------------------------------------------------------------------
' LIMITE DE CRÉDITO ATUAL DO CLIENTE
' -------------------------------------------------------------------------------------------------
Dim LimiteDe
LimiteDe = 0

SET DADOS = Conecta("PIR-DADOSADV", "SELECT dbo.fn_SaldoCliente('" & clienteCod & "') AS SALDO_CREDITO")
if not isObject(DADOS) or (DADOS IS NOTHING) then response.end
if not DADOS.EOF then LimiteDe = CDbl(DADOS("SALDO_CREDITO"))
DADOS.Close() : SET DADOS = NOTHING

' =================================================================================================
' HELPER: monta body HTML do e-mail de notificação
' =================================================================================================
function MontaMsgEmail(linhasExtras)
    Dim corpo
    corpo = "<html>" & _
                "<body>" & _
                    "<font face='verdana, arial' size='2'>" & _
                        "Informamos que o pedido <b>" & pedido & "</b> do cliente " & _
                        clienteCod & "-" & lojaCod & " - " & nomeCliente & _
                        ", foi alterado. Segue abaixo os dados de alteração:<br><br>" & _
                        "Tipo de alteração:" & Req("NomeTipoOperacao") & "<br>" & _
                        linhasExtras & _
                        "Usuário:" & Usr & "<br>" & _
                        "Data da alteração:" & AcertaData(dataAlteracao) & " " & horaAlteracao & "<br>" & _
                        MsgPubOp & _
                    "</font>" & _
                "</body>" & _
            "</html>"
    MontaMsgEmail = corpo
end function

' =================================================================================================
' PROCESSAMENTO POR TIPO DE OPERAÇÃO
' =================================================================================================
Dim SQL, SQL1, SQL2, SQL3, SQL4
Dim Consulta, Mensagem, Assunto, Para_Email, Gravou
Mensagem = ""
Gravou   = 0

Consulta = IniciaTransacao

select case CInt(tipoOp)

' -------------------------------------------------------------------------------------------------
case 1 ' TIPO DE FRETE
' -------------------------------------------------------------------------------------------------
    if Left(Produto, 3) <> "700" then

        SQL = "" & _
            "UPDATE SC5010 SET " & _
                "C5_TPFRETE = '" & para & "', " & _
                "C5_FRETUNI = 0, " & _
                "C5_FRETE = 0 " & _
            "WHERE C5_NUM = '" & pedido & "'"

    elseif para = "C" then

        SQL = "" & _
            "UPDATE SC5010 SET " & _
                "C5_TPFRETE = '" & para & "', " & _
                "C5_FRETUNI = " & Replace(Replace(valorFrete, ".", ""), ",", ".") & ", " & _
                "C5_FRETE = "   & Replace(Replace(QuantVendida * valorFrete, ".", ""), ",", ".") & " " & _
            "WHERE C5_NUM = '" & pedido & "'"

    else

        SQL = "" & _
            "UPDATE SC5010 SET " & _
                "C5_TPFRETE = '" & para & "', " & _
                "C5_FRETUNI = 0, " & _
                "C5_FRETE = 0 " & _
            "WHERE C5_NUM = '" & pedido & "'"

    end if

    SQL1 = GravaHistorico()
    Consulta = Consulta & SQL & " " & SQL1 & " " & FinalizaTransacao

    SET DADOS = Conecta("PIR-DADOSADV", Consulta)
    if not isObject(DADOS) or (DADOS IS NOTHING) then response.end

    if GravouInHistorico() = 1 then
        Para_Email = "transporte.pira@smurfitwestrock.com.br"
        Assunto    = "Alteração de tipo de frete para o pedido " & pedido
        copia      = copia & ";sac@smurfitwestrock.com.br;" & gerente

        Mensagem = MontaMsgEmail( _
            "De:" & de & "<br>" & _
            "Para:" & para & "<br>" & _
            "Frete Unitário:" & FormatNumber(valorFrete, 5) & "<br>")

        call EnviaEmail(Para_Email, copia, "", "", emailDe, Mensagem, Assunto)
    end if

' -------------------------------------------------------------------------------------------------
case 2 ' PREÇO
' -------------------------------------------------------------------------------------------------
    if CDbl(novoTotal) <= 0 then
        Aviso("O novo preço deve ser maior do que zero.")
        response.end
    end if

    if EVAL(CDbl(novoTotal) - CDbl(Req("Total"))) > EVAL(LimiteDe) then
        Aviso("O limite de crédito do cliente foi ultrapassado.")
        response.end
    end if

    SQL = "" & _
        "UPDATE SC6010 SET " & _
            "C6_PRCVEN = " & Replace(para, ",", ".") & ", " & _
            "C6_PRUNIT = " & Replace(para, ",", ".") & ", " & _
            "C6_VALOR  = " & Replace(novoTotal, ",", ".") & " " & _
        "WHERE C6_NUM = '" & pedido & "'; "

    SQL1 = GravaHistorico()
    Consulta = Consulta & SQL & " " & SQL1 & " " & FinalizaTransacao

    SET DADOS = Conecta("PIR-DADOSADV", Consulta)
    if not isObject(DADOS) or (DADOS IS NOTHING) then response.end

    if GravouInHistorico() = 1 then
        Para_Email = "comercial-rj@smurfitwestrock.com.br"
        Assunto    = "Alteração do pedido " & pedido
        copia      = copia & ";sac@smurfitwestrock.com.br;" & gerente & ";luiz.souza@smurfitwestrock.com.br"

        Mensagem = MontaMsgEmail( _
            "De:" & de & "<br>" & _
            "Para:" & para & "<br>")

        call EnviaEmail(Para_Email, copia, "", "", emailDe, Mensagem, Assunto)
    end if

' -------------------------------------------------------------------------------------------------
case 3 ' QUANTIDADE
' -------------------------------------------------------------------------------------------------
    if CLng(para) <= 0 then
        Aviso("A nova quantidade deve ser maior do que zero.")
        response.end
    end if

    if CDbl(novoTotal) - CDbl(Req("Total")) > CDbl(LimiteDe) then
        Aviso("O limite de crédito do cliente foi ultrapassado.")
        response.end
    end if

    SQL = "" & _
        "UPDATE SC6010 SET " & _
            "C6_QTDVEN = " & Replace(para, ",", ".") & ", " & _
            "C6_VALOR  = " & Replace(para, ",", ".") & " * C6_PRCVEN " & _
        "WHERE C6_NUM = '" & pedido & "'"

    SQL1 = GravaHistorico()
    Consulta = Consulta & SQL & " " & SQL1 & " " & FinalizaTransacao

    SET DADOS = Conecta("PIR-DADOSADV", Consulta)
    if not isObject(DADOS) or (DADOS IS NOTHING) then response.end

    if GravouInHistorico() = 1 then
        Para_Email = "sara.nascimento@smurfitwestrock.com.br;" & _
                     "flavia.pereira@smurfitwestrock.com.br;" & _
                     "luiz.junior@smurfitwestrock.com.br;" & _
                     "marcos.abreu@smurfitwestrock.com.br;" & _
                     "fransciele.dias@smurfitwestrock.com.br"
        Assunto = "Alteração do pedido " & pedido
        copia   = copia & ";sac@smurfitwestrock.com.br;" & gerente & ";comercial-rj@smurfitwestrock.com.br"

        Mensagem = MontaMsgEmail( _
            "De:" & de & "<br>" & _
            "Para:" & para & "<br>")

        call EnviaEmail(Para_Email, copia, "", "", emailDe, Mensagem, Assunto)
    end if

' -------------------------------------------------------------------------------------------------
case 4 ' QUANTIDADE PRONTA
' -------------------------------------------------------------------------------------------------
    SQL = "" & _
        "UPDATE SC6010 SET " & _
            "C6_PRONTO = " & Replace(para, ",", ".") & " " & _
        "WHERE C6_NUM = '" & pedido & "' "

    SQL1 = GravaHistorico()
    Consulta = Consulta & SQL & " " & SQL1 & " " & FinalizaTransacao

    SET DADOS = Conecta("PIR-DADOSADV", Consulta)
    if not isObject(DADOS) or (DADOS IS NOTHING) then response.end

' -------------------------------------------------------------------------------------------------
case 5 ' DATA DE ENTREGA
' -------------------------------------------------------------------------------------------------
    SQL = "" & _
        "UPDATE SC6010 SET " & _
            "C6_ENTREG  = '" & InverteData(Req("NovaEntrega1")) & "', " & _
            "C6_ENTREG2 = '" & InverteData(Req("NovaEntrega2")) & "' " & _
        "WHERE C6_NUM = '" & pedido & "'"

    SQL1 = GravaHistorico()
    Consulta = Consulta & SQL & " " & SQL1 & " " & FinalizaTransacao

    SET DADOS = Conecta("PIR-DADOSADV", Consulta)
    if not isObject(DADOS) or (DADOS IS NOTHING) then response.end

    if GravouInHistorico() = 1 then
        Para_Email = "sara.nascimento@smurfitwestrock.com.br;" & _
                     "flavia.pereira@smurfitwestrock.com.br;" & _
                     "luiz.junior@smurfitwestrock.com.br;" & _
                     "fransciele.dias@smurfitwestrock.com.br;" & _
                     "martimira.oliveira@smurfitwestrock.com.br"
        Assunto = "Alteração do pedido " & pedido
        copia   = copia & ";sac@smurfitwestrock.com.br;" & gerente & ";comercial-rj@smurfitwestrock.com.br"

        Mensagem = MontaMsgEmail( _
            "De:" & de & "<br>" & _
            "Para:" & para & "<br>")

        call EnviaEmail(Para_Email, copia, "", "", emailDe, Mensagem, Assunto)
    end if

' -------------------------------------------------------------------------------------------------
case 6 ' CONDIÇÕES DE PAGAMENTO
' -------------------------------------------------------------------------------------------------
    SQL = "" & _
        "UPDATE SC5010 SET " & _
            "C5_CONDPAG = '" & Req("AlteraCondPag") & "' " & _
        "WHERE C5_NUM = '" & pedido & "'"

    SQL1 = GravaHistorico()
    Consulta = Consulta & SQL & " " & SQL1 & " " & FinalizaTransacao

    SET DADOS = Conecta("PIR-DADOSADV", Consulta)
    if not isObject(DADOS) or (DADOS IS NOTHING) then response.end

    if GravouInHistorico() = 1 then
        Para_Email = "comercial-rj@smurfitwestrock.com.br"
        Assunto    = "Alteração do pedido " & pedido
        copia      = copia & ";sac@smurfitwestrock.com.br;" & gerente & ";gilberto.diniz@smurfitwestrock.com.br"

        Mensagem = MontaMsgEmail( _
            "De:" & de & "<br>" & _
            "Para:" & para & "<br>")

        call EnviaEmail(Para_Email, copia, "", "", emailDe, Mensagem, Assunto)
    end if

' -------------------------------------------------------------------------------------------------
case 7 ' CANCELAMENTO
' -------------------------------------------------------------------------------------------------
    SQL  = "UPDATE SC5010 SET C5_ESTATUS = '" & para & "' WHERE C5_NUM = '" & pedido & "'"
    SQL1 = "UPDATE SC6010 SET C6_BLQ = 'R' WHERE C6_NUM = '" & pedido & "'"
    SQL2 = GravaHistorico()

    ' Determina filial para atualizar Trimbox
    Dim filialCancelamento
    filialCancelamento = ObtemFilial(pedido)
    SQL3 = SqlTrimbox(filialCancelamento, pedido, 4, "<> 4")

    Consulta = Consulta & SQL & " " & SQL1 & " " & SQL3 & " " & SQL2 & " " & FinalizaTransacao

    SET DADOS = Conecta("PIR-DADOSADV", Consulta)
    if not isObject(DADOS) or (DADOS IS NOTHING) then response.end

    if GravouInHistorico() = 1 then
        Para_Email = "sara.nascimento@smurfitwestrock.com.br;" & _
                     "flavia.pereira@smurfitwestrock.com.br;" & _
                     "luiz.junior@smurfitwestrock.com.br;" & _
                     "marcos.abreu@smurfitwestrock.com.br;" & _
                     "fransciele.dias@smurfitwestrock.com.br;" & _
                     "martimira.oliveira@smurfitwestrock.com.br"
        Assunto = "Alteração do pedido " & pedido
        copia   = copia & ";sac@smurfitwestrock.com.br;" & gerente

        Mensagem = MontaMsgEmail( _
            "De:" & de & "<br>" & _
            "Para:" & para & "<br>")

        call EnviaEmail(Para_Email, copia, "", "", emailDe, Mensagem, Assunto)
    end if

' -------------------------------------------------------------------------------------------------
case 8 ' PEDIDO DO CLIENTE
' -------------------------------------------------------------------------------------------------
    Dim sqlPedCli
    sqlPedCli = "REPLACE(REPLACE(Ltrim(Rtrim('" & para & "')), char(9), ''),' ','')"

    SQL = "" & _
        "UPDATE SC6010 SET " & _
            "C6_PEDCLI  = " & sqlPedCli & ", " & _
            "C6_NUMPCOM = " & sqlPedCli & " " & _
        "WHERE C6_NUM = '" & pedido & "'"

    SQL1 = "" & _
        "UPDATE SC6010 SET " & _
            "C6_PEDCLI  = REPLACE(REPLACE(Ltrim(Rtrim(C6_PEDCLI)),  char(9), ''),' ',''), " & _
            "C6_NUMPCOM = REPLACE(REPLACE(Ltrim(Rtrim(C6_NUMPCOM)), char(9), ''),' ','') " & _
        "WHERE C6_NUM = '" & pedido & "'"

    SQL2 = GravaHistorico()
    Consulta = Consulta & SQL & " " & SQL1 & " " & SQL2 & " " & FinalizaTransacao

    SET DADOS = Conecta("PIR-DADOSADV", Consulta)
    if not isObject(DADOS) or (DADOS IS NOTHING) then response.end

' -------------------------------------------------------------------------------------------------
case 9 ' STATUS DO PEDIDO (exceto cancelamento — case 7)
' -------------------------------------------------------------------------------------------------
    ' Fluxo: saindo de L ou C para status ativo → verifica limite de crédito
    if (de = "L" or de = "C") and (para <> "L" and para <> "C") then

        Dim Total9
        SET DADOS = Conecta("PIR-DADOSADV", "" & _
            "SELECT (C6_QTDVEN - C6_QTDENT) * C6_PRCVEN AS TOTAL " & _
            "FROM SC6010 WITH(NOLOCK) " & _
            "WHERE C6_NUM = '" & pedido & "' " & _
              "AND D_E_L_E_T_ <> '*' " & _
              "AND SUBSTRING(C6_PRODUTO,1,3) <> '120' ")

        if not isObject(DADOS) or (DADOS IS NOTHING) then response.end
        Total9 = iif(not DADOS.EOF, CDbl(DADOS("TOTAL")), 0)
        DADOS.Close() : SET DADOS = NOTHING

        if (LimiteDe - Total9) < 0 then
            Aviso("O limite de crédito do cliente foi ultrapassado. O saldo após alteração seria de: " & FormatNumber(LimiteDe - Total9, 2))
            response.end
        end if

        SQL  = "UPDATE SC5010 SET C5_ESTATUS = '" & para & "', C5_DTLIQUI = ' ' WHERE C5_NUM = '" & pedido & "'"
        SQL1 = "UPDATE SC6010 SET C6_BLQ = ' ' WHERE C6_NUM = '" & pedido & "'"

        Dim filial9 : filial9 = ObtemFilial(pedido)
        SQL2 = SqlTrimbox(filial9, pedido, 1, "= 4")
        SQL3 = GravaHistorico()

        Consulta = Consulta & SQL & " " & SQL1 & " " & SQL2 & " " & SQL3 & " " & FinalizaTransacao

        SET DADOS = Conecta("PIR-DADOSADV", Consulta)
        if not isObject(DADOS) or (DADOS IS NOTHING) then response.end
        Gravou = GravouInHistorico()

    else ' para = "L" ou para = "C"

        if para = "L" then

            SQL  = "UPDATE SC5010 SET C5_ESTATUS = '" & para & "' WHERE C5_NUM = '" & pedido & "'"

            Dim DataLiq, HoraLiq
            DataLiq = InverteData(Date())
            HoraLiq = Time()

            SQL1 = "UPDATE SC5010 SET " & _
                       "C5_DTLIQUI = '" & AcertaData(DataLiq) & "-" & HoraLiq & _
                       Replace(Left(Usr, 9), ".", "") & "-INPANET' " & _
                   "WHERE C5_NUM = '" & pedido & "'"

            SQL2 = "UPDATE SC6010 SET C6_BLQ = 'R' WHERE C6_NUM = '" & pedido & "'"

            Dim filialL : filialL = ObtemFilial(pedido)
            SQL3 = SqlTrimbox(filialL, pedido, 4, "<> 4")
            SQL4 = GravaHistorico()

            Consulta = Consulta & SQL & " " & SQL1 & " " & SQL2 & " " & SQL3 & " " & SQL4 & " " & FinalizaTransacao

            SET DADOS = Conecta("PIR-DADOSADV", Consulta)
            if not isObject(DADOS) or (DADOS IS NOTHING) then response.end
            Gravou = GravouInHistorico()

        end if

        if para = "C" then

            SQL  = "UPDATE SC5010 SET C5_ESTATUS = '" & para & "' WHERE C5_NUM = '" & pedido & "'"
            SQL1 = "UPDATE SC6010 SET C6_BLQ = 'R' WHERE C6_NUM = '" & pedido & "'"

            Dim filialC : filialC = ObtemFilial(pedido)
            SQL2 = SqlTrimbox(filialC, pedido, 4, "<> 4")
            SQL3 = GravaHistorico()

            Consulta = Consulta & SQL & " " & SQL1 & " " & SQL2 & " " & SQL3 & " " & FinalizaTransacao

            SET DADOS = Conecta("PIR-DADOSADV", Consulta)
            if not isObject(DADOS) or (DADOS IS NOTHING) then response.end
            Gravou = GravouInHistorico()

        end if

        if Gravou = 1 then
            Para_Email = "sara.nascimento@smurfitwestrock.com.br;" & _
                         "flavia.pereira@smurfitwestrock.com.br;" & _
                         "luiz.junior@smurfitwestrock.com.br;" & _
                         "fransciele.dias@smurfitwestrock.com.br;" & _
                         "martimira.oliveira@smurfitwestrock.com.br"
            Assunto  = "Alteração do pedido " & pedido
            copia    = copia & ";sac@smurfitwestrock.com.br;" & gerente

            Mensagem = MontaMsgEmail( _
                "De:" & de & "<br>" & _
                "Para:" & para & "<br>")

            call EnviaEmail(Para_Email, copia, "", "", emailDe, Mensagem, Assunto)
        end if

    end if

' -------------------------------------------------------------------------------------------------
case 10 ' FILIAL DO PEDIDO
' -------------------------------------------------------------------------------------------------
    if (de = "I" or de = "U") and (para = "I" or para = "U") then

        SQL  = "UPDATE SC5010 SET C5_INPUBE = '" & para & "' WHERE C5_NUM = '" & pedido & "' AND D_E_L_E_T_ <> '*'"
        SQL1 = GravaHistorico()
        Consulta = Consulta & SQL & " " & SQL1 & " " & FinalizaTransacao

        SET DADOS = Conecta("PIR-DADOSADV", Consulta)
        if not isObject(DADOS) or (DADOS IS NOTHING) then response.end

        if GravouInHistorico() = 1 then
            Para_Email = "sara.nascimento@smurfitwestrock.com.br;" & _
                         "flavia.pereira@smurfitwestrock.com.br;" & _
                         "luiz.junior@smurfitwestrock.com.br;" & _
                         "fransciele.dias@smurfitwestrock.com.br;" & _
                         "martimira.oliveira@smurfitwestrock.com.br"
            Assunto = "Alteração do pedido " & pedido
            copia   = copia & ";sac@smurfitwestrock.com.br;" & gerente & ";comercial-rj@smurfitwestrock.com.br"

            Dim labelDe10, labelPara10
            labelDe10   = iif(de = "U",   "UBERABA", "PIRAPETINGA")
            labelPara10 = iif(para = "U", "UBERABA", "PIRAPETINGA")

            Mensagem = MontaMsgEmail( _
                "De: "  & labelDe10   & "<br>" & _
                "Para: " & labelPara10 & "<br>")

            call EnviaEmail(Para_Email, copia, "", "", emailDe, Mensagem, Assunto)
        end if

    end if

' -------------------------------------------------------------------------------------------------
case 11 ' MOTIVO ATRASO
' -------------------------------------------------------------------------------------------------
    SQL  = "UPDATE SC5010 SET C5_OBSATR = '" & Req("Para") & "' WHERE C5_NUM = '" & pedido & "'"
    SQL1 = GravaHistorico()
    Consulta = Consulta & SQL & " " & SQL1 & " " & FinalizaTransacao

    SET DADOS = Conecta("PIR-DADOSADV", Consulta)
    if not isObject(DADOS) or (DADOS IS NOTHING) then response.end

    if GravouInHistorico() = 1 then
        Para_Email = "antonio.golfeto@smurfitwestrock.com.br"
        Assunto    = "Alteração do pedido " & pedido
        copia      = copia & ";sac@smurfitwestrock.com.br;" & gerente

        ' Motivo atraso mostra QuemLiberou, não Usr — mantido fiel ao original
        Dim msgAtraso
        msgAtraso = "<html><body><font face='verdana, arial' size='2'>" & _
                        "Informamos que o pedido <b>" & pedido & "</b> do cliente " & _
                        clienteCod & "-" & lojaCod & " - " & nomeCliente & _
                        ", foi alterado. Segue abaixo os dados de alteração:<br><br>" & _
                        "Tipo de alteração:" & Req("NomeTipoOperacao") & "<br>" & _
                        "De:" & de & "<br>" & _
                        "Para:" & para & "<br>" & _
                        "Usuário:" & Req("QuemLiberou") & "<br>" & _
                        "Data da alteração:" & dataAlteracao & " " & horaAlteracao & "<br>" & _
                        MsgPubOp & _
                    "</font></body></html>"

        call EnviaEmail(Para_Email, copia, "", "", emailDe, msgAtraso, Assunto)
    end if

' -------------------------------------------------------------------------------------------------
case 14 ' FRETE UNITÁRIO (somente Papel — produto iniciado em 700)
' -------------------------------------------------------------------------------------------------
    if Left(Req("Produto"), 3) <> "700" then
        Response.Write "É permitida alteração de frete, apenas para Papel."
        response.end
    end if

    SQL  = "UPDATE SC5010 SET C5_FRETUNI = '" & para & "' WHERE C5_NUM = '" & pedido & "'"
    SQL1 = GravaHistorico()
    Consulta = Consulta & SQL & " " & SQL1 & " " & FinalizaTransacao

    SET DADOS = Conecta("PIR-DADOSADV", Consulta)
    if not isObject(DADOS) or (DADOS IS NOTHING) then response.end

    if GravouInHistorico() = 1 then
        Para_Email = "luiz.junior@smurfitwestrock.com.br"
        Assunto    = "Alteração de frete unitário do pedido " & pedido
        copia      = copia & ";igor.muniz@smurfitwestrock.com.br;Gilberto.Diniz@smurfitwestrock.com.br;" & gerente

        Mensagem = MontaMsgEmail( _
            "De:" & de & "<br>" & _
            "Para:" & para & "<br>")

        call EnviaEmail(Para_Email, copia, "", "", emailDe, Mensagem, Assunto)
    end if

end select

' -------------------------------------------------------------------------------------------------
' E-MAIL PARA O REPRESENTANTE (complementar, mesmo corpo do último case executado)
' -------------------------------------------------------------------------------------------------
if emailRep <> "" and Mensagem <> "" then
    call EnviaEmail(emailRep, "", "", "", emailDe, Mensagem, Assunto)
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>SWnet - Confirmação da alteração do pedido</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />

    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" />
    <link rel="stylesheet"
          href="https://use.fontawesome.com/releases/v5.7.0/css/all.css"
          integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ"
          crossorigin="anonymous" />

    <style type="text/css">
        #loader {
            position: absolute;
            left: 50%; top: 50%;
            z-index: 1;
            width: 120px; height: 120px;
            margin: -76px 0 0 -76px;
            border: 16px solid #f3f3f3;
            border-radius: 50%;
            border-top: 16px solid #3498db;
            -webkit-animation: spin 2s linear infinite;
            animation: spin 2s linear infinite;
        }
        @-webkit-keyframes spin { 0% { -webkit-transform: rotate(0deg); } 100% { -webkit-transform: rotate(360deg); } }
        @keyframes spin          { 0% { transform: rotate(0deg); }        100% { transform: rotate(360deg); } }

        .animate-bottom {
            position: relative;
            -webkit-animation-name: animatebottom; -webkit-animation-duration: 1s;
            animation-name: animatebottom;          animation-duration: 1s;
        }
        @-webkit-keyframes animatebottom { from { bottom:-100px; opacity:0 } to { bottom:0px; opacity:1 } }
        @keyframes animatebottom         { from { bottom:-100px; opacity:0 } to { bottom:0;   opacity:1 } }
    </style>

    <script>
        function navegarParaGrade() { location.href = 'SolicitaAlteracao.asp'; }
    </script>
</head>
<body>
<%
if GravouInHistorico() = 1 then

    response.write "<script>HideLoader();</script>"
    AvisoSucesso("Pedido Alterado com Sucesso.")

    response.write "" & _
        "<div class='row'>" & _
            "<div class='col d-flex justify-content-center'>" & _
                "<button class='btn btn-danger' onclick='navegarParaGrade();'>Voltar</button>" & _
            "</div>" & _
        "</div>"

else

    Aviso("Houve um problema ao alterar este pedido. Tente novamente.")
    response.end

end if
%>
</body>
</html>
<%
' =================================================================================================
' FUNÇÕES AUXILIARES
' =================================================================================================

' -------------------------------------------------------------------------------------------------
' Retorna o SQL de UPDATE no Trimbox correto (PIR ou UBE) conforme filial do pedido
' statusNovo  : valor inteiro a gravar em StatusPedido
' condicaoWhere : cláusula WHERE parcial, ex: "<> 4" ou "= 4"
' -------------------------------------------------------------------------------------------------
function SqlTrimbox(filial, numPedido, statusNovo, condicaoWhere)
    Dim srv, db
    if filial <> "U" then
        srv = cnx(seekCnx("PIR-TRIMBOX")).getServer
        db  = cnx(seekCnx("PIR-TRIMBOX")).getDataBase
    else
        srv = cnx(seekCnx("UBE-TRIMBOX")).getServer
        db  = cnx(seekCnx("UBE-TRIMBOX")).getDataBase
    end if
    SqlTrimbox = "" & _
        "UPDATE " & srv & "." & db & ".dbo.Pedidos SET " & _
            "StatusPedido = " & statusNovo & " " & _
        "WHERE StatusPedido " & condicaoWhere & " " & _
          "AND Pedido = '" & numPedido & "' "
end function

' -------------------------------------------------------------------------------------------------
' Busca a filial (C5_INPUBE) do pedido no SC5010
' -------------------------------------------------------------------------------------------------
function ObtemFilial(numPedido)
    Dim rs, filialRet
    filialRet = "0"
    SET rs = Conecta("PIR-DADOSADV", "" & _
        "SELECT C5_INPUBE FROM SC5010 WITH(NOLOCK) " & _
        "WHERE C5_NUM = '" & numPedido & "' AND D_E_L_E_T_ <> '*' ")
    if isObject(rs) and not (rs IS NOTHING) then
        if not rs.EOF then filialRet = Trim(rs("C5_INPUBE"))
        rs.Close() : SET rs = NOTHING
    end if
    ObtemFilial = filialRet
end function

' -------------------------------------------------------------------------------------------------
' Monta o INSERT no histórico + UPDATE da solicitação (quando idSolicitacao preenchido).
' Usa as variáveis globais da requisição — sem repetir 31 parâmetros em cada case.
' -------------------------------------------------------------------------------------------------
function GravaHistorico()

    Dim sqlH
    sqlH = "" & _
        "INSERT INTO HELPDESK..InHistoricoAlteracoesPedidos (" & _
            "Pedido, Quantidade, QuantidadePronto, Preco, Total, " & _
            "Entrega1, Entrega2, Emissao, TipoFrete, CondPag, " & _
            "Estatus, SiglaEstatusTrim, " & _
            "QuemCadastrou, DataCadastro, HoraCadastro, " & _
            "QuemValidou, DataValidacao, HoraValidacao, " & _
            "CodCliente, LojaCliente, NomeCliente, " & _
            "Produto, DescriProduto, " & _
            "Operacao, De, Para, NovoTotal, " & _
            "[Login], DataAlteracao, HoraAlteracao, " & _
            "LimiteCreditoDe, LimiteCreditoPara, ValorFrete " & _
        ") VALUES( " & _
            "'" & pedido & "', " & _
            Replace(Replace(Req("Quantidade"),        ".", ""), ",", ".") & ", " & _
            Replace(Replace(Req("QuantidadePronta"),  ".", ""), ",", ".") & ", " & _
            Replace(Replace(FormatNumber(CDbl(Req("Preco")),5), ".", ""), ",", ".") & ", " & _
            Replace(Replace(Req("Total"),             ".", ""), ",", ".") & ", '" & _
            Req("Entrega1")        & "', '" & _
            Req("Entrega2")        & "', '" & _
            Req("Emissao")         & "', '" & _
            Req("TipoFrete")       & "', '" & _
            Req("NomeCondPag")     & "', '" & _
            Req("Estatus")         & "', '" & _
            Req("SiglaEstatusTrim") & "', '" & _
            Req("QuemCadastrou")   & "', '" & _
            Req("DataCadastro")    & "', '" & _
            Req("HoraCadastro")    & "', '" & _
            Req("QuemLiberou")     & "', '" & _
            Req("DataLiberado")    & "', '" & _
            Req("HoraLiberado")    & "', '" & _
            clienteCod             & "', '" & _
            lojaCod                & "', '" & _
            SanitizaStr(nomeCliente) & "', '" & _
            Req("Produto")         & "', '" & _
            SanitizaStr(Req("DescricaoProduto")) & "', " & _
            tipoOp                 & ", '" & _
            de                     & "', '" & _
            para                   & "', " & _
            Replace(Replace(novoTotal, ".", ""), ",", ".") & ", '" & _
            Usr                    & "', '" & _
            dataAlteracao          & "', '" & _
            horaAlteracao          & "', " & _
            Replace(Replace(FormatNumber(CDbl(LimiteDe), 5),      ".", ""), ",", ".") & ", " & _
            Replace(Replace(FormatNumber(CDbl(LimitePara(clienteCod)), 5), ".", ""), ",", ".") & ", " & _
            Replace(Replace(FormatNumber(CDbl(valorFrete),  5),   ".", ""), ",", ".") & _
        "); "

    if idSolicitacao <> "" then
        sqlH = sqlH & _
            "UPDATE HELPDESK..InHistoricoAlteracoesPedidosSolicitacao SET " & _
                "Realizado = 1, " & _
                "QuemRealizou = '" & Usr & "', " & _
                "DtRealizou = convert(varchar(8),getdate(),112) " & _
            "WHERE Pedido = '" & pedido & "' " & _
              "AND Tipo = '" & tipoOp & "' " & _
              "AND Id = " & idSolicitacao & " "
    end if

    GravaHistorico = sqlH

end function

' -------------------------------------------------------------------------------------------------
' Verifica se o registro foi gravado no histórico
' -------------------------------------------------------------------------------------------------
function GravouInHistorico()
    Dim rs, saved
    saved = 0
    SET rs = Conecta("PIR-HELPDESK", "" & _
        "SELECT COUNT(*) AS EXISTE " & _
        "FROM InHistoricoAlteracoesPedidos " & _
        "WHERE Pedido   = '" & pedido  & "' " & _
          "AND Para      = '" & para    & "' " & _
          "AND Operacao  = '" & tipoOp  & "' " & _
          "AND ISNULL(D_E_L_E_T_,'') = ''")
    if isObject(rs) and not (rs IS NOTHING) then
        saved = iif(not rs.EOF and CInt(Trim(rs("EXISTE"))) > 0, 1, 0)
        rs.Close() : SET rs = NOTHING
    end if
    GravouInHistorico = saved
end function

' -------------------------------------------------------------------------------------------------
' Consulta o saldo de crédito do cliente após a operação
' -------------------------------------------------------------------------------------------------
function LimitePara(codCli)
    Dim rs2
    SET rs2 = Conecta("PIR-DADOSADV", "SELECT dbo.fn_SaldoCliente('" & Trim(codCli) & "') AS SALDO_CREDITO")
    if not isObject(rs2) or (rs2 IS NOTHING) then response.end
    LimitePara = iif(not rs2.EOF, rs2("SALDO_CREDITO"), 0)
    rs2.Close() : SET rs2 = NOTHING
end function

' -------------------------------------------------------------------------------------------------
' AcertaData: AAAAMMDD → DD/MM/AAAA
' -------------------------------------------------------------------------------------------------
function AcertaData(d)
    AcertaData = Right(d, 2) & "/" & Mid(d, 5, 2) & "/" & Left(d, 4)
end function

' -------------------------------------------------------------------------------------------------
' InverteData: recebe Date() do VBScript e devolve AAAAMMDD
' -------------------------------------------------------------------------------------------------
function InverteData(d)
    Dim ano, mes, dia
    ano = Year(d)
    mes = Right("0" & Month(d), 2)
    dia = Right("0" & Day(d),   2)
    InverteData = ano & mes & dia
end function

' -------------------------------------------------------------------------------------------------
' Aviso: exibe alerta de erro Bootstrap 4
' -------------------------------------------------------------------------------------------------
function Aviso(msg)
    response.write "" & _
        "<br><br>" & _
        "<div class='alert alert-danger alert-dismissible w-50 mx-auto alert-vertical-center' role='alert'>" & _
            "<a href='#' class='close' data-dismiss='alert' aria-label='close' onclick='window.history.back();'>&times;</a>" & _
            "<br><h5>A V I S O !!!</h5><hr>" & _
            "<p>" & msg & "</p>" & _
        "</div>"
end function

' -------------------------------------------------------------------------------------------------
' AvisoSucesso: exibe alerta de sucesso Bootstrap 4
' -------------------------------------------------------------------------------------------------
function AvisoSucesso(msg)
    response.write "" & _
        "<br><br>" & _
        "<div class='alert alert-success alert-dismissible w-50 mx-auto alert-vertical-center' role='alert'>" & _
            "<br><h6 class='text-center'>" & msg & "</h6><br>" & _
        "</div>"
end function
%>
