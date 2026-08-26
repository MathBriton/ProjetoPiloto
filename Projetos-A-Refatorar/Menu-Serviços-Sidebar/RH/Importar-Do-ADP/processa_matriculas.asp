<!--#include file="../../cnn2/cnn2.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->

<%
'Option Explicit
Response.ContentType = "text/plain"

Dim tipo, csvdata, dataEhora
tipo = LCase(trim(Request.form("tipo")))
csvdata = Request.form("csvdata")
dataEhora = Request.form("dataehora")
aPlanta = Request.form("planta")

'response.write "<br>" & tipo & ", " & dataEhora & ", " & aPlanta & "<br>"

if tipo <> "funcionarios" And tipo <> "acessos" then
  Response.Write "Erro: parâmetro 'tipo' inválido."
  Response.End
end if

if Len(trim(csvdata)) = 0 then
  Response.Write "Erro: nenhum dado CSV recebido."
  Response.End
end if

' Separar linhas CSV por CRLF ou LF
Dim linhas
linhas = Split(csvdata, vbCrLf)
if UBound(linhas) = 0 then
  linhas = Split(csvdata, vbLf)
end if

Dim totalLinhas, inseridos, existentes
totalLinhas = 0
inseridos = 0
existentes = 0

'--------------------------------------------------------------------------------------------------'
function EscapaSQL(texto)
  EscapaSQL = Replace(texto, "'", "''")
end function
'--------------------------------------------------------------------------------------------------'
function MatriculaExiste(matricula, tabela)

  SQL = "" &_
    "SELECT " &_
      "COUNT(*) AS Total " &_
    "FROM " & tabela & " " &_
    "WHERE " & iif( tabela = "Funcionarios", "Matricula", "Credencial" ) & " = '" & EscapaSQL(matricula) & "'"

  SET BEN_MDACESSO = Conecta( "BEN-MDACESSO", SQL )

  if not isObject( BEN_MDACESSO ) or ( BEN_MDACESSO IS NOTHING ) then
    response.end
  end if

  MatriculaExiste = False

  if (tabela = "Funcionarios") and (not BEN_MDACESSO.EOF) and (BEN_MDACESSO("Total") > 0) then 
    
      MatriculaExiste = True

  elseif (tabela = "FuncionariosAcessos") then

      MatriculaExiste = false

  end if

  BEN_MDACESSO.Close
  SET BEN_MDACESSO = NOTHING

end function
'--------------------------------------------------------------------------------------------------'
Sub InserirRegistro(campos, tabela, aPlanta)

  Dim sql, i, colunas, valores

  if tabela = "Funcionarios" then

    colunas = "Matricula" &_
              ",Nome" &_
              ",CPF" &_
              ",PIS" &_
              ",Credencial" &_
              ",Afastado" &_
              ",DataInicioAfastamento" &_
              ",DataFimAfastamento" &_
              ",BatePonto" &_
              ",Filial" &_
              ",CancelaCarros" &_
              ",DataAtualizacao"

  else

    colunas = "Credencial" &_
              ",DataLiberacao" &_
              ",HoraInicioPrimeiraEntrada" &_
              ",HoraFimPrimeiraEntrada" &_
              ",HoraInicioPrimeiraSaida" &_
              ",HoraFimPrimeiraSaida" &_
              ",HoraInicioSegundaEntrada" &_
              ",HoraFimSegundaEntrada" &_
              ",HoraInicioSegundaSaida" &_
              ",HoraFimSegundaSaida" &_
              ",ViraTurno" &_
              ",DataAtualizacao"

  end if


  valores = "'" & EscapaSQL(campos(0)) & "'"

  for i = 1 to UBound(campos)

    if tabela = "Funcionarios" then
      if i = 6 and campos(5) = "0" then
        valores = valores & ", NULL "
      elseif i = 7 and campos(5) = "0" then
        valores = valores & ", NULL "
      else  
        valores = valores & ", '" & EscapaSQL(campos(i)) & "'"
      end if
    else
        valores = valores & ", '" & EscapaSQL(campos(i)) & "'"
    end if

  next

  valores = valores & ", '" & dataEhora & "' "
            
  SQL = "BEGIN TRANSACTION " &_
        "INSERT INTO " & tabela & " (" & colunas & ") VALUES (" & valores & ");" &_
        "IF @@ERROR <> 0 " &_
          "ROLLBACK TRANSACTION " &_
        "ELSE " &_
          "COMMIT TRANSACTION "

  'response.write SQL & "<br>"

  select case aPlanta
  case "00"
    SET PIR_CORPORATIVO = Conecta( "PIR-CORPORATIVO", SQL )
    SET PIR_CORPORATIVO = NOTHING
  case "04"
    SET UBE_DMPACESSO = Conecta( "UBE-DMPACESSO", SQL )
    SET UBE_DMPACESSO = NOTHING
  case "05"
    SET BEN_MDACESSO = Conecta( "BEN-MDACESSO", SQL )
    SET BEN_MDACESSO = NOTHING
  end select

End Sub
'--------------------------------------------------------------------------------------------------'
Dim i, linhaRaw, camposArray, matricula, tabelaDestino

if tipo = "funcionarios" then

  tabelaDestino = "Funcionarios"

elseif tipo = "acessos" then

  tabelaDestino = "FuncionariosAcessos"

end if

for i = 0 to UBound(linhas)

  linhaRaw = trim(linhas(i))

  if linhaRaw <> "" then

    camposArray = Split(linhaRaw, ";")

    if UBound(camposArray) >= 0 then

      totalLinhas = totalLinhas + 1
      matricula = trim(camposArray(0))

      if matricula <> "" then

''        if MatriculaExiste(matricula, tabelaDestino) then
''          existentes = existentes + 1
''        else

          On Error Resume next
          InserirRegistro camposArray, tabelaDestino, aPlanta

          if Err.Number = 0 then
            inseridos = inseridos + 1
          end if

          On Error GoTo 0

''        end if

      end if

    end if

  end if

next

response.write "<br>Importação tipo '" & tipo & "' concluída!" & "<br>" & _
               "Linhas processadas: " & totalLinhas & "<br>" & _
               "Novas inserções: " & inseridos & "<br>" & _
               "Já existiam: " & existentes
%>