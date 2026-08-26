<!--#include file = "conecta.asp"--> 
<!--#include file = "../funcoes/funcoes.asp"-->
<%
On Error GoTo 0

On Error Resume Next
'Ajusta o tempo da sessao em minutos
Session.TimeOut = 60

'Dim Conexao
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.CommandTimeout = 300
Conexao.Open connect(0)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO PROTHEUS" )
else
'----------------------------------------------
' CONEXÃO PROTHEUS
'----------------------------------------------

  Dim Dados
  Set Dados = Server.CreateObject("ADODB.Recordset")
  Dados.ActiveConnection = Conexao

  Dim Dados1
  Set Dados1 = Server.CreateObject("ADODB.Recordset")
  Dados1.ActiveConnection = Conexao

  Dim Dados2
  Set Dados2 = Server.CreateObject("ADODB.Recordset")
  Dados2.ActiveConnection = Conexao

  Dim Dados3
  Set Dados3 = Server.CreateObject("ADODB.Recordset")
  Dados3.ActiveConnection = Conexao

  Dim Dados4
  Set Dados4 = Server.CreateObject("ADODB.Recordset")
  Dados4.ActiveConnection = Conexao

  Dim Dados5
  Set Dados5 = Server.CreateObject("ADODB.Recordset")
  Dados5.ActiveConnection = Conexao

  Dim Dados6
  Set Dados6 = Server.CreateObject("ADODB.Recordset")
  Dados6.ActiveConnection = Conexao

  Dim Dados7
  Set Dados7 = Server.CreateObject("ADODB.Recordset")
  Dados7.ActiveConnection = Conexao

  Dim Dados01
  Set Dados01 = Server.CreateObject("ADODB.Recordset")
  Dados01.ActiveConnection = Conexao

  Dim Dados02
  Set Dados02 = Server.CreateObject("ADODB.Recordset")
  Dados02.ActiveConnection = Conexao

  Dim Dados03
  Set Dados03 = Server.CreateObject("ADODB.Recordset")
  Dados3.ActiveConnection = Conexao

  Dim Dados04
  Set Dados14 = Server.CreateObject("ADODB.Recordset")
  Dados4.ActiveConnection = Conexao

  Dim Dados05
  Set Dados05 = Server.CreateObject("ADODB.Recordset")
  Dados05.ActiveConnection = Conexao

  Dim Dados06
  Set Dados06 = Server.CreateObject("ADODB.Recordset")
  Dados06.ActiveConnection = Conexao

  Dim Dados07
  Set Dados07 = Server.CreateObject("ADODB.Recordset")
  Dados07.ActiveConnection = Conexao  

end if

on error goto 0

'----------------------------------------------
' CONEXÃO CORPORATIVO
'----------------------------------------------

On Error Resume Next

'Dim Conexao1
Set Conexao1 = Server.CreateObject("ADODB.Connection")
Conexao1.CommandTimeout = 180
Conexao1.Open connect(1)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO CORPORATIVO" )
else
  
  Session.TimeOut = 60

  Dim Dados11
  Set Dados11 = Server.CreateObject("ADODB.Recordset")
  Dados11.ActiveConnection = Conexao1

  Dim Dados12
  Set Dados12 = Server.CreateObject("ADODB.Recordset")
  Dados12.ActiveConnection = Conexao1

  Dim Dados13
  Set Dados13 = Server.CreateObject("ADODB.Recordset")
  Dados13.ActiveConnection = Conexao1

  Dim Dados15
  Set Dados15 = Server.CreateObject("ADODB.Recordset")
  Dados15.ActiveConnection = Conexao1

end if
'----------------------------------------------
' CONEXÃO HELPDESK
'----------------------------------------------
on error goto 0
On Error Resume Next

'Dim Conexao2
Set Conexao2 = Server.CreateObject("ADODB.Connection")
Conexao2.CommandTimeout = 180
Conexao2.Open connect(2)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO HELPDESK" )
else
  
  Session.TimeOut = 60

  Dim Dados21
  Set Dados21 = Server.CreateObject("ADODB.Recordset")
  Dados21.ActiveConnection = Conexao2

  Dim Dados22
  Set Dados22 = Server.CreateObject("ADODB.Recordset")
  Dados22.ActiveConnection = Conexao2

  Dim Dados23
  Set Dados23 = Server.CreateObject("ADODB.Recordset")
  Dados23.ActiveConnection = Conexao2

  Dim Dados24
  Set Dados24 = Server.CreateObject("ADODB.Recordset")
  Dados24.ActiveConnection = Conexao2

end if
'--------------------------------------------------------------
'               CONEXÃO TRIMBOX PIRAPETINGA
'--------------------------------------------------------------
on error goto 0
On Error Resume Next

'Dim Conexao5
Set Conexao5 = Server.CreateObject("ADODB.Connection")
Conexao5.CommandTimeout = 180
Conexao5.Open connect(5)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO TRIMBOX PIRAPETINGA" )
else

  Dim Dados51
  Set Dados51 = Server.CreateObject("ADODB.Recordset")
  Dados51.ActiveConnection = Conexao5

  Dim Dados52
  Set Dados52 = Server.CreateObject("ADODB.Recordset")
  Dados52.ActiveConnection = Conexao5

  Dim Dados53
  Set Dados53 = Server.CreateObject("ADODB.Recordset")
  Dados53.ActiveConnection = Conexao5

end if
'--------------------------------------------------------------
' CONEXÃO TRIMBOX UBERABA
'--------------------------------------------------------------
on error goto 0
On Error Resume Next

'Dim Conexao6
Set Conexao6 = Server.CreateObject("ADODB.Connection")
Conexao6.CommandTimeout = 180
Conexao6.Open connect(6)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO TRIMBOX UBERABA" )
else

  Dim Dados61
  Set Dados61 = Server.CreateObject("ADODB.Recordset")
  Dados61.ActiveConnection = Conexao6

  Dim Dados62
  Set Dados62 = Server.CreateObject("ADODB.Recordset")
  Dados62.ActiveConnection = Conexao6

  Dim Dados63
  Set Dados63 = Server.CreateObject("ADODB.Recordset")
  Dados63.ActiveConnection = Conexao6

end if
'--------------------------------------------------------------
' CONEXÃO TRIMPAPEL UBERABA
'--------------------------------------------------------------
on error goto 0
On Error Resume Next

'Dim Conexao8
Set Conexao8 = Server.CreateObject("ADODB.Connection")
Conexao8.CommandTimeout = 180
Conexao8.Open connect(8)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO TRIMPAPEL UBERABA" )
else
  
  Dim Dados81
  Set Dados81 = Server.CreateObject("ADODB.Recordset")
  Dados81.ActiveConnection = Conexao8

  Dim Dados82
  Set Dados82 = Server.CreateObject("ADODB.Recordset")
  Dados82.ActiveConnection = Conexao8

  Dim Dados83
  Set Dados83 = Server.CreateObject("ADODB.Recordset")
  Dados83.ActiveConnection = Conexao8 

end if
'----------------------------------------------
' CONEXÃO TRIMBOX BENTO
'----------------------------------------------
on error goto 0
On Error Resume Next

'Dim Conexao9
Set Conexao9 = Server.CreateObject("ADODB.Connection")
Conexao9.CommandTimeout = 180
Conexao9.Open connect(9)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO TRIMBOX BENTO" )
else

  Session.TimeOut = 60

  Dim Dados91
  Set Dados91 = Server.CreateObject("ADODB.Recordset")
  Dados91.ActiveConnection = Conexao9
  
  Dim Dados92
  Set Dados92 = Server.CreateObject("ADODB.Recordset")
  Dados92.ActiveConnection = Conexao9
  
  Dim Dados93
  Set Dados93 = Server.CreateObject("ADODB.Recordset")
  Dados93.ActiveConnection = Conexao9  

end if
'--------------------------------------------------------------
' CONEXÃO SOLPEN BENTO
'--------------------------------------------------------------
on error goto 0
On Error Resume Next

'Dim Conexao10
Set Conexao10 = Server.CreateObject("ADODB.Connection")
Conexao10.CommandTimeout = 180
Conexao10.Open connect(10)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO SOLPEN BENTO" )
else

  Dim Dados101
  Set Dados101 = Server.CreateObject("ADODB.Recordset")
  Dados101.ActiveConnection = Conexao10

  Dim Dados102
  Set Dados102 = Server.CreateObject("ADODB.Recordset")
  Dados102.ActiveConnection = Conexao10

  Dim Dados103
  Set Dados103 = Server.CreateObject("ADODB.Recordset")
  Dados103.ActiveConnection = Conexao10

  Dim Dados104
  Set Dados104 = Server.CreateObject("ADODB.Recordset")
  Dados104.ActiveConnection = Conexao10

end if
'--------------------------------------------------------------
' CONEXÃO SIADNE FORTALEZA
'--------------------------------------------------------------
on error goto 0
On Error Resume Next

'Dim Conexao11
Set Conexao11 = Server.CreateObject("ADODB.Connection")
Conexao11.CommandTimeout = 180
Conexao11.Open connect(11)

if err.number <> 0 then

  call HandleError( err, GetFileName(), "CONEXÃO SIADNE FORTALEZA" )

else

  Dim Dados111
  Set Dados111 = Server.CreateObject("ADODB.Recordset")
  Dados111.ActiveConnection = Conexao11
  
end if

on error goto 0

'----------------------------------------------
' CONEXÃO TRIMBOX FORTALEZA
'----------------------------------------------
On Error Resume Next

'Dim Conexao12
Set Conexao12 = Server.CreateObject("ADODB.Connection")
Conexao12.CommandTimeout = 180
Conexao12.Open connect(12)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO TRIMBOX FORTALEZA" )
else

  Session.TimeOut = 60

  Dim Dados121
  Set Dados121 = Server.CreateObject("ADODB.Recordset")
  Dados121.ActiveConnection = Conexao12

  Dim Dados122
  Set Dados122 = Server.CreateObject("ADODB.Recordset")
  Dados122.ActiveConnection = Conexao12

  Dim Dados123
  Set Dados123 = Server.CreateObject("ADODB.Recordset")
  Dados123.ActiveConnection = Conexao12

end if
'----------------------------------------------
' CONEXÃO CORPORATIVO FORTALEZA
'----------------------------------------------
On Error Resume Next

Set Conexao20 = Server.CreateObject("ADODB.Connection")
Conexao20.CommandTimeout = 180
Conexao20.Open connect(19)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO CORPORATIVO FORTALEZA" )
else

  Session.TimeOut = 60

  Dim Dados201
  Set Dados201 = Server.CreateObject("ADODB.Recordset")
  Dados201.ActiveConnection = Conexao20

  Dim Dados202
  Set Dados202 = Server.CreateObject("ADODB.Recordset")
  Dados202.ActiveConnection = Conexao20
end if

'----------------------------------------------
' CONEXÃO DADOSADV FORTALEZA
'----------------------------------------------
On Error Resume Next

Set Conexao22 = Server.CreateObject("ADODB.Connection")
Conexao22.CommandTimeout = 180
Conexao22.Open connect(21)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO DADOSADV FORTALEZA" )
else

  Session.TimeOut = 60

  Dim Dados221
  Set Dados221 = Server.CreateObject("ADODB.Recordset")
  Dados221.ActiveConnection = Conexao22

  Dim Dados222
  Set Dados222 = Server.CreateObject("ADODB.Recordset")
  Dados222.ActiveConnection = Conexao22
end if
'----------------------------------------------
' CONEXÃO CORPORATIVO BENTO
'----------------------------------------------
on error goto 0
On Error Resume Next

'Dim Conexao13
Set Conexao13 = Server.CreateObject("ADODB.Connection")
Conexao13.CommandTimeout = 180
Conexao13.Open connect(13)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO CORPORATIVO BENTO" )
else

  Session.TimeOut = 60

  Dim Dados131
  Set Dados131 = Server.CreateObject("ADODB.Recordset")
  Dados131.ActiveConnection = Conexao13

  Dim Dados132
  Set Dados132 = Server.CreateObject("ADODB.Recordset")
  Dados132.ActiveConnection = Conexao13

  Dim Dados133
  Set Dados133 = Server.CreateObject("ADODB.Recordset")
  Dados133.ActiveConnection = Conexao13

  Dim Dados134
  Set Dados134 = Server.CreateObject("ADODB.Recordset")
  Dados134.ActiveConnection = Conexao13

  Dim Dados135
  Set Dados135 = Server.CreateObject("ADODB.Recordset")
  Dados135.ActiveConnection = Conexao13  

end if
'----------------------------------------------
' CONEXÃO TRIMPAPEL BENTO
'----------------------------------------------
on error goto 0
On Error Resume Next

'Dim Conexao14
Set Conexao14 = Server.CreateObject("ADODB.Connection")
Conexao14.CommandTimeout = 180
Conexao14.Open connect(14)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO TRIMPAPEL BENTO" )
else

  Session.TimeOut = 60

  Dim Dados141
  Set Dados141 = Server.CreateObject("ADODB.Recordset")
  Dados141.ActiveConnection = Conexao14
  
  Dim Dados142
  Set Dados142 = Server.CreateObject("ADODB.Recordset")
  Dados142.ActiveConnection = Conexao14
  
  Dim Dados143
  Set Dados143 = Server.CreateObject("ADODB.Recordset")
  Dados143.ActiveConnection = Conexao14

end if
'----------------------------------------------
' CONEXÃO TRIMPAPEL PIRAPETINGA
'----------------------------------------------
on error goto 0
On Error Resume Next

'Dim Conexao15
Set Conexao15 = Server.CreateObject("ADODB.Connection")
Conexao15.CommandTimeout = 180
Conexao15.Open connect(7)
'Conexão alterada de 15 para 7 porque foi corrigido o arquivo conexao.txt em 21/12/2023

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO TRIMPAPEL PIRAPETINGA" )
else
  
  Session.TimeOut = 60
 
  Dim Dados151
  Set Dados151 = Server.CreateObject("ADODB.Recordset")
  Dados151.ActiveConnection = Conexao15
  
  Dim Dados152
  Set Dados152 = Server.CreateObject("ADODB.Recordset")
  Dados152.ActiveConnection = Conexao15
  
  Dim Dados153
  Set Dados153 = Server.CreateObject("ADODB.Recordset")
  Dados153.ActiveConnection = Conexao15

end if
'----------------------------------------------
' CONEXÃO DADOSADV SAQUAREMA
'----------------------------------------------
on error goto 0
On Error Resume Next

Set Conexao16 = Server.CreateObject("ADODB.Connection")
Conexao16.CommandTimeout = 180
Conexao16.Open connect(15)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO DADOSADV SAQUAREMA" )
else
  
 Session.TimeOut = 60
 
  Dim Dados161
  Set Dados161 = Server.CreateObject("ADODB.Recordset")
  Dados161.ActiveConnection = Conexao16
  
  Dim Dados162
  Set Dados162 = Server.CreateObject("ADODB.Recordset")
  Dados162.ActiveConnection = Conexao16
 
end if
'----------------------------------------------
' CONEXÃO CORPORATIVO SAQUAREMA
'----------------------------------------------
on error goto 0
On Error Resume Next

Set Conexao18 = Server.CreateObject("ADODB.Connection")
Conexao18.CommandTimeout = 180
Conexao18.Open connect(17)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO CORPORATIVO SAQUAREMA" )
else
  
  Session.TimeOut = 60
 
  Dim Dados181
  Set Dados181 = Server.CreateObject("ADODB.Recordset")
  Dados181.ActiveConnection = Conexao18
  
  Dim Dados182
  Set Dados182 = Server.CreateObject("ADODB.Recordset")
  Dados182.ActiveConnection = Conexao18
  
end if
'----------------------------------------------
' CONEXÃO TRIMBOX SAQUAREMA
'----------------------------------------------
on error goto 0
On Error Resume Next

Set Conexao19 = Server.CreateObject("ADODB.Connection")
Conexao19.CommandTimeout = 180
Conexao19.Open connect(18)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO TRIMBOX SAQUAREMA" )
else
  
  Session.TimeOut = 60
 
  Dim Dados191
  Set Dados191 = Server.CreateObject("ADODB.Recordset")
  Dados191.ActiveConnection = Conexao19
  
  Dim Dados192
  Set Dados192 = Server.CreateObject("ADODB.Recordset")
  Dados192.ActiveConnection = Conexao19
  
end if

'----------------------------------------------
' CONEXÃO CORPORATIVO SAO PAULO
'----------------------------------------------
on error goto 0
On Error Resume Next

Set Conexao17 = Server.CreateObject("ADODB.Connection")
Conexao17.CommandTimeout = 180
Conexao17.Open connect(16)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO CORPORATIVO SAO PAULO" )
else
  
  Session.TimeOut = 60
 
  Dim Dados171
  Set Dados171 = Server.CreateObject("ADODB.Recordset")
  Dados171.ActiveConnection = Conexao17
    
end if

'----------------------------------------------
' CONEXÃO GUARDIANTI BENTO GONCALVES
'----------------------------------------------
on error goto 0
On Error Resume Next

Dim Conexao21
Set Conexao21 = Server.CreateObject("ADODB.Connection")
Conexao21.CommandTimeout = 180
Conexao21.Open connect(20)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO GUARDIANTI BENTO GONCALVES" )
else
  
  Session.TimeOut = 60
 
  Dim Dados211
  Set Dados211 = Server.CreateObject("ADODB.Recordset")
  Dados211.ActiveConnection = Conexao21

  Dim Dados212
  Set Dados212 = Server.CreateObject("ADODB.Recordset")
  Dados212.ActiveConnection = Conexao21

  Dim Dados213
  Set Dados213 = Server.CreateObject("ADODB.Recordset")
  Dados213.ActiveConnection = Conexao21

end if


'----------------------------------------------
' CONEXÃO CORPORATIVO UBERABA
'----------------------------------------------
on error goto 0
On Error Resume Next

Set Conexao23 = Server.CreateObject("ADODB.Connection")
Conexao23.CommandTimeout = 180
Conexao23.Open connect(23)

if err.number <> 0 then
  call HandleError( err, GetFileName(), "CONEXÃO CORPORATIVO UBERABA" )
else
  
  Session.TimeOut = 60
 
  Dim Dados231
  Set Dados231 = Server.CreateObject("ADODB.Recordset")
  Dados231.ActiveConnection = Conexao23
  
  Dim Dados232
  Set Dados232 = Server.CreateObject("ADODB.Recordset")
  Dados232.ActiveConnection = Conexao23
  
end if

on error goto 0
%>
