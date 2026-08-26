<!--#Include File = "../../../funcoes/funcoes.asp"-->
<!--#include file = "../../../cnn2/cnn2.asp" -->
<%
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

Function VerificaSessaoExpirada()
         if Trim(Session("CODUSER")) = "" then
            Call SesseaoExpirada()
         end if
End function   

Function ShowMsg(Mensagem)
         response.Write "<script>alert('" & Mensagem & "')</script>"
End Function

Function rp(v)
  v = Trim(v)
  rp = Replace(v,"'","")
End Function

Function rpPonto(v)
  v = Trim(v)
  rpPonto = Replace(Replace(v,".",""),",",".")
End Function

Function Hoje()
        Dia = Day(Date())
        Mes = Month(Date())
        Ano = Year(Date())
        if Len(Dia) = 1 then
            Dia = "0" & Dia
        end if

        if Len(Mes) = 1 then
            Mes = "0" & Mes
        end if
        
        Hoje = Ano & Mes & Dia
End Function

Function Agora()
        Hora = Hour(Time())
        Minuto = Minute(Time())
        Segundo = Second(Time())
        if Len(cStr(Hora)) = 1 then
            Hora = "0" & Hora
        end if

        if Len(cStr(Minuto)) = 1 then
            Minuto = "0" & Minuto
        end if

        if Len(cStr(Segundo)) = 1 then
            Segundo = "0" & Segundo
        end if

        Agora = Hora & ":" & Minuto & ":" & Segundo
End Function

Function fnumber(v)
        if v = "" then
            fnumber = ""
        Else
            fnumber = Replace(Replace(FormatNumber(v,2),".",""),",",".")
        end if
End Function

Function TestaValor(argValor)
    If IsNull(argValor) Or Trim(argValor) = "" Then
        TestaValor = ""
    Else
        TestaValor = Replace(argValor, "'", "''") 
    End If
End Function

function DataInvertidaSemSeparador(argValor)
         if argValor="" then
            DataInvertidaSemSeparador=""
            exit function
         end if
         
         dim arrayData
         arrayData=split(argValor,"/")
         DataInvertidaSemSeparador=arrayData(2)+arrayData(1)+arrayData(0)
end function

Function EhZonaFrancaDeManaus(argCidade, argEstado)
    Dim cidadesZFM, i
    Dim cidadeZFM
    Dim resultado

    ' Lista dos municípios com incentivos da Zona Franca de Manaus
    cidadesZFM = Array( _
        "Manaus", _
        "Itacoatiara", _
        "Manacapuru", _
        "Iranduba", _
        "Rio Preto da Eva", _
        "Presidente Figueiredo", _
        "Tabatinga" _
    )

    resultado = False
    If UCase(Trim(argEstado)) = "AM" Then
        For i = 0 To UBound(cidadesZFM)
            cidadeZFM = cidadesZFM(i)
            If LCase(Trim(argCidade)) = LCase(cidadeZFM) Then
                resultado = True
                Exit For
            End If
        Next
    End If

    EhZonaFrancaDeManaus = resultado
End Function

Function LimparCNPJ(cnpj)
    Dim i, ch, resultado
    resultado = ""

    For i = 1 To Len(cnpj)
        ch = Mid(cnpj, i, 1)
        If ch >= "0" And ch <= "9" Then
            resultado = resultado & ch
        End If
    Next

    LimparCNPJ = resultado
End Function
%>