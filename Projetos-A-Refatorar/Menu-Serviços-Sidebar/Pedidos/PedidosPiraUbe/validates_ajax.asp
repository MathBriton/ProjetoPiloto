<%session.LCID = 1046 %>
<!--#include file="../includes/ValidaTES.asp" -->
<!--#Include File = "../cnn/DADOSADV.asp"-->
<!--#Include File = "../funcoes/funcoes.asp"-->
<%
    if trim(request("cliente")) <> "" and trim(request("loja")) <> "" and trim(request("produto")) <> "" and trim(request("tes")) <> "" then
        if trim(request("codsiga")) = "" and trim(request("nivel")) <> "" then
            response.write "ErroLogof|"&trim(request("tipo"))
            response.End
        end if
        TesDescricao = ""

        SQL = "SELECT F4_CODIGO,F4_DESC FROM SF4010 WHERE F4_CODIGO = '"&trim(request("tes"))&"' AND D_E_L_E_T_ = '' "
        'response.write SQL
        Dados.Source = SQL
        Dados.open()
        if Not Dados.Eof then
            TesDescricao = Trim(Dados("F4_DESC"))
        end if
        Dados.Close()

        set vt = new ValidaTES
        Dim validou 
        validou = False
        if TesDescricao <> "" then
            if trim(request("eshop")) = "1" then
                response.write "|"&trim(request("tipo"))&"|True|"&trim(request("tes"))&"|"&TesDescricao
                response.end
            end if
            SQL = "SELECT A1_GERTES, A1_TESEXCE,A1_DTVLTES FROM SA1010 WHERE A1_COD = '"&trim(request("cliente"))&"' AND A1_LOJA = '"&trim(request("loja"))&"' AND D_E_L_E_T_ = '' AND A1_MSBLQL = '2'"        
            'RESPONSE.WRITE SQL
            Dados.Source = SQL
            Dados.Open()
            if not Dados.Eof then
                Tes = split(Dados("A1_GERTES"),",")
                for i = 0 to Ubound(Tes)
                    if Tes(i) = trim(request("tes")) then
                        validou = True
                    end if
                next
                if validou then
                    response.write "|"&trim(request("tipo"))&"|True|"&trim(request("tes"))&"|"&TesDescricao
                    response.end
                else
                    if Dados("A1_DTVLTES") = InverteData(Date()) then
                        if Trim(Dados("A1_TESEXCE")) = trim(request("tes")) then
                            response.write "|"&trim(request("tipo"))&"|True|"&trim(request("tes"))&"|"&TesDescricao
                            response.End
                        'else
                            'response.write "|"&trim(request("tipo"))&"|False|"&trim(request("tes"))&"|"&TesDescricao
                        end if
                    'else
                        'response.write "|"&trim(request("tipo"))&"|False|"&trim(request("tes"))&"|"&TesDescricao
                    end if
                end if
            'else
                'validou = vt.Validar(trim(request("tes")),trim(request("cliente")),trim(request("loja")),trim(request("cliente2")),trim(request("loja2")),trim(request("tiponf")),trim(request("produto")),trim(request("fnc")))
                response.write "|"&trim(request("tipo"))&"|false|"&trim(request("tes"))&"|"&TesDescricao
                response.end
            else
                response.write "ErroNEncontrado|"&trim(request("tipo"))
            end if
        Else
            response.write "ErroNEncontrado|"&trim(request("tipo"))
        end if
    else
        response.Write "ErroTes|"&trim(request("tipo"))
    end if
%>
