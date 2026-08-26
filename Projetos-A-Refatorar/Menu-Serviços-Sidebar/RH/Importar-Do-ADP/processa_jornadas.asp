<!--#include file="../../cnn2/cnn2.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->

<%
'Option Explicit
'Response.ContentType = "text/plain"

Function GetVariableType(var)
  Dim vType
  vType = VarType(var)
  Select Case vType
  Case vbEmpty    GetVariableType = "Empty"
  Case vbNull     GetVariableType = "Null"
  Case vbInteger  GetVariableType = "Integer"
  Case vbLong     GetVariableType = "Long"
  Case vbSingle   GetVariableType = "Single"
  Case vbDouble   GetVariableType = "Double"
  Case vbCurrency GetVariableType = "Currency"
  Case vbString   GetVariableType = "String"
  Case vbObject   GetVariableType = "Object"
  Case vbBoolean  GetVariableType = "Boolean"
  Case vbVariant  GetVariableType = "Variant"
  Case vbDate     GetVariableType = "Date"
  Case Else       GetVariableType = "Unknown Type"
  End Select
End Function

Dim tipo, csvdata
tipo = LCase(trim(Request.form("tipo")))

Dim totalLinhas, inseridos
totalLinhas = 0
inseridos = 0

SQL = "SELECT DISTINCT Credencial FROM FuncionariosAcessos"

SET TABACESSOS = Conecta("BEN-MDACESSO", SQL)

if NOT isObject(TABACESSOS) then
  response.end
end if

while NOT TABACESSOS.EOF

  totalLinhas = totalLinhas + 1

  SQLmontajornada = "" &_
    "SELECT DISTINCT TOP 1 " &_
      "Credencial," &_
      "HoraInicioPrimeiraEntrada AS HoraInicioPrimeira," &_
      "HoraFimSegundaSaida AS HoraFimPrimeira," &_
      "(SELECT DISTINCT TOP 1 " &_
          "HoraInicioPrimeiraEntrada " &_
        "FROM FuncionariosAcessos " &_
        "WHERE HoraInicioPrimeiraEntrada <> A.HoraInicioPrimeiraEntrada " &_
          "and HoraFimSegundaSaida <> A.HoraFimSegundaSaida " &_
          "and Credencial = '" & trim(TABACESSOS("Credencial")) & "' " &_
      ") AS HoraInicioSegunda, " &_
      "(SELECT DISTINCT TOP 1 " &_
          "HoraFimSegundaSaida " &_
        "FROM FuncionariosAcessos " &_
        "WHERE HoraFimSegundaSaida <> A.HoraFimSegundaSaida " &_
          "and HoraInicioPrimeiraEntrada <> A.HoraInicioPrimeiraEntrada " &_
          "and Credencial = '" & trim(TABACESSOS("Credencial")) & "'" &_
      ") AS HoraFimSegunda " &_
    "FROM FuncionariosAcessos A " &_
    "WHERE A.HoraInicioPrimeiraEntrada <> '00:00:00.00000' " &_
        "and A.HoraFimSegundaSaida <> '00:00:00.00000' " &_
        "and A.Credencial = '" & trim(TABACESSOS("Credencial")) & "' "

  SET MONTAJORNADA = Conecta( "BEN-MDACESSO", SQLmontajornada )

  if NOT isObject( MONTAJORNADA ) then
    response.end
  end if

  ctJornada = 0

  SQLjornada = "BEGIN TRANSACTION "

  while NOT MONTAJORNADA.EOF

    inseridos = inseridos + 1

    Dim vHoraInicioPrimeira
    vHoraInicioPrimeira = MONTAJORNADA("HoraInicioPrimeira")

    ' Convertendo o valor TIME para string
    Dim vHoraInicioPrimeiraStr
    vHoraInicioPrimeiraStr = vHoraInicioPrimeira(timeValue, vbLongTime) 

    response.write MONTAJORNADA("Credencial")
    response.write "|"
    response.write vHoraInicioPrimeiraStr
    response.end

    response.write "|"
    response.write cstr(MONTAJORNADA("HoraFimPrimeira"))
    response.write "|"
    response.write cstr(MONTAJORNADA("HoraInicioSegunda"))
    response.write "|"
    response.write cstr(MONTAJORNADA("HoraFimSegunda"))
    response.write "|<br>"

    response.end


    ' ifCredencial         = iif( isnull( MONTAJORNADA("Credencial") )        , " NULL, ", " '" & trim( MONTAJORNADA("Credencial") )         & "', ")
    ' ifHoraInicioPrimeira = iif( isnull( MONTAJORNADA("HoraInicioPrimeira") ), " NULL, ", " '" & trim( MONTAJORNADA("HoraInicioPrimeira") ) & "', ")
    ' ifHoraFimPrimeira    = iif( isnull( MONTAJORNADA("HoraFimPrimeira") )   , " NULL, ", " '" & trim( MONTAJORNADA("HoraFimPrimeira") )    & "', ")
    ' ifHoraInicioSegunda  = iif( isnull( MONTAJORNADA("HoraInicioSegunda") ) , " NULL, ", " '" & trim( MONTAJORNADA("HoraInicioSegunda") )  & "', ")
    ' ifHoraFimSegunda     = iif( isnull( MONTAJORNADA("HoraFimSegunda") )    , " NULL, ", " '" & trim( MONTAJORNADA("HoraFimSegunda") )     & "', ")

    ' if not isnull( MONTAJORNADA("Credencial") ) then

    '   SQLjornada = SQLjornada &_
    '     "INSERT INTO FuncionariosJornadas( " &_
    '       "Credencial," &_
    '       "HoraInicioPrimeira," &_
    '       "HoraFimPrimeira," &_
    '       "HoraInicioSegunda," &_
    '       "HoraFimSegunda" &_
    '     ") VALUES (" &_
    '       ifCredencial &_
    '       ifHoraInicioPrimeira &_
    '       ifHoraFimPrimeira &_
    '       ifHoraInicioSegunda &_
    '       ifHoraFimSegunda &_
    '     "); "

    '   ctJornada = ctJornada + 1

    ' end if

    MONTAJORNADA.MoveNext()

  wend

  if (ctJornada > 0 ) then

    SQLjornada = SQLjornada &_
      "IF @@ERROR <> 0 " &_
        "ROLLBACK TRANSACTION " &_
      "ELSE " &_
        "COMMIT TRANSACTION "

    'RESPONSE.WRITE SQLjornada
    'RESPONSE.END

    SET TABJORNADA = Conecta("BEN-MDACESSO", SQLjornada)
    SET TABJORNADA = NOTHING

  end if

  MONTAJORNADA.Close()
  SET MONTAJORNADA = NOTHING

  TABACESSOS.MoveNext()

wend

TABACESSOS.Close()
SET TABACESSOS = NOTHING

Response.Write "Importação tipo '" & tipo & "' concluída!" & vbCrLf & _
               "Linhas processadas: " & totalLinhas & vbCrLf & _
               "Novas inserções: " & inseridos
%>