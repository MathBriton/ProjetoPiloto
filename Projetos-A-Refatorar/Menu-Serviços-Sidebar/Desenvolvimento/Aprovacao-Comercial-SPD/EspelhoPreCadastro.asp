<!DOCTYPE html>
<%SESSION.LCID = 1046%>
<html lang="pt-br" xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta charset="utf-8" />
    <title>Espelho Pre-Cadastro de Produto</title>
    <link rel="stylesheet" href="../../bootstrap/3.3.5/css/bootstrap.min.css" />
    <script src="../../ajax/libs/jquery/1.11.3/jquery.min.js"></script>
    <script src="../../bootstrap/3.3.5/js/bootstrap.min.js"></script>
    <!--#Include File = "../cnn/cnn.asp"-->
    <!--#Include File = "../funcoes/funcoes.asp"-->
    <script charset="utf-8" src="../funcoes/funcoesJAVASCRIPT.js"></script>
    <script type="text/javascript">
    </script>
    <style type="text/css">
        body {
            font-size:90%;
            margin:0px;       
        }
        img {
            width:147px;
            height:46px;
        }
        h4 {
            display:inline-block;
        }
        table {
            margin-top: 0px;
            margin-bottom: 0px;
            padding:0;
        }
        div{
            margin-top: 0px;
            margin-bottom: 0px;
            padding:0;
        }
        .BordaInferior {
            border-bottom: 1px solid #ddd;
        }
        .BordaDireita {
            border-right: 1px solid #ddd;
        }
        .BordaEsquerda {
            border-left: 1px solid #ddd;
        }
        .BordaSuperior {
            border-top: 1px solid #ddd;
        }
        .MargemBotton {
            margin-bottom:10px;
        }
        .border {
            border:1px solid black;
        }
        .CabecalhoDestaque {
            background-color:#ccc;
        }
        .MargemEsquerda {
            margin-left:15px;
        }
        .MargemDireita {
            margin-right:30px;
        }
        .auto-style1 {
            height: 22px;
        }
        .auto-style2 {
            text-align: center;
            height: 22px;
        }
    </style>
    <style type="text/css" media="print">
        @media print {
            @page {
                margin: 0.5cm;
                }
            body {
                font-size:98.5%;
                margin:0px;       
            }
            img {
                margin-right:42px;
            }
            #printable {
                visibility:collapse;
            }
        }
    </style>
  </head>
<%

select case trim(request("Planta"))
      case "00" set Consulta = Dados11: set Consulta1 = Dados11
      case "04" set Consulta = Dados11: set Consulta1 = Dados11
      case "05" set Consulta = Dados91: set Consulta1 = Dados131
      case "06" set Consulta = Dados121: set Consulta1 = Dados201
      case "07" set Consulta = Dados191: set Consulta1 = Dados181   
end select

mascaraFT = 1

if Request("Planta") = "05" Or Request("Planta") = "06" Or Request("Planta") = "07" then
  mascaraFT = 0

  SQL = "" &_
    "SELECT " &_
      iif(Request("Planta") = "06", "FT1_FTNOVAFOR", "FT1_FTNOVA") & " AS FT1_FTNOVA, " & _
      "FT1_B5TPPALE, " &_
      "FT1_USU " &_
    "FROM SKFT1 " &_
    "WHERE ISNULL(FT1_DEL,0) = 0 " &_
      "AND FT1_ID = "&Trim(request("Id"))
  Consulta1.Open(SQL)

  if Not Consulta1.Eof then
    FT1_FTNOVA = Consulta1("FT1_FTNOVA")
    FT1_B5TPPALE = Consulta1("FT1_B5TPPALE")
    FT1_USU = Consulta1("FT1_USU")
  end if

  Consulta1.Close()
  SQL = "" &_
    "SELECT " &_
      "Descricao " &_
    "FROM vw_ConsultaPaletesPelaInpaNet " &_
    "WHERE Desativado = 0 " &_
      "AND isnull(CodigoERP,'') = '" & FT1_B5TPPALE & "'"
  Consulta.Open(SQL)
  Descricao = iif( Not Consulta.EOF,Consulta("Descricao"), "" )
  Consulta.Close()

  SQL = "" &_
    "SELECT " &_
      "isnull(FitasLargPalete,0) FitasLarguraPalete " &_
      ",isnull(FitasCompPalete,0) FitasComprimentoPalete " &_
      ",isnull(FitilhosLargPacote,0) FitilhosLargura " &_
      ",isnull(FitilhosCompPacote,0) FitilhosComprimento " &_
      ",Referencia " &_
      ",TolMais " &_
      ",TolMenos " &_                 
      ",ComprimentoInterno " &_
      ",LarguraInterna " &_
      ",AlturaInterna " &_
      ",TipoABNT " &_
      ",Compressao " &_
      ",Composicao " &_            
      ",Fechamento" &_
      ",Amarrado" &_
      ",PecasPorPacote " &_
      ",Paletizado " &_
      ",TipoPalete " &_
      ",FacesOpostas " &_
      ",Cantoneiras " &_
      ",Espelho " &_
      ",Filme " &_
      ",Gramatura " &_
      ",f.CodCliente " &_
      ",c.RazaoSocial " &_
      ",ColunaMinimo " &_
      ",CobbExtMaximo " &_
      ",CobbIntMaximo " &_
      ",Compressao " &_
      ",MullenMinimo " &_
      ",Espessura " &_
      ",ColunaMinimo " &_
      ",Cor1 " &_
      ",Cor2 " &_
      ",Cor3 " &_
      ",Cor4 " &_
      ",PecasPorConjunto " &_
      ",Arranjo " &_
      ",PecasPorPalete " &_
      ",Largura " &_
      ",Comprimento " &_
      ",MultLarg " &_
      ",MultComp " &_
      ",AreaBrutaChapa " &_
      ",AreaLiquidaChapa " &_
      ",VincoLarg1 " &_
      ",VincoLarg2 " &_
      ",VincoLarg3 " &_
      ",VincoLarg4 " &_
      ",VincoLarg5 " &_
      ",VincoLarg6 " &_
      ",VincoLarg7 " &_
      ",VincoLarg1 " &_
      ",VincoLarg2 " &_
      ",VincoLarg3 " &_
      ",VincoLarg4 " &_
      ",VincoLarg5 " &_
      ",VincoLarg6 " &_
      ",VincoLarg7 " &_
      ",Maquina1 " &_
      ",Maquina2 " &_
      ",Maquina3 " &_
      ",Maquina4 " &_
      ",Maquina5 " &_
      ",Faca " &_
      ",Obs1 " &_ 
      ",Obs2 " &_
      ",Obs3 " &_
      ",Representante " &_
      ",ViscosidadeCor1 " &_
      ",ViscosidadeCor2 " &_ 
      ",ViscosidadeCor3 " &_   
      ",DataCriacao " &_
    "FROM vw_ConsultaFTs f " &_
    "LEFT JOIN vw_RazaoSocialCliente c ON f.CodCliente = c.CodCliente " &_
    "WHERE Item = '" & FT1_FTNOVA & "'"
   
    Consulta.Open(SQL)

  if not Consulta.EOF then
    B5_REF     = Trim(Consulta("Referencia"))
    B5_TOLMA   = Trim(Consulta("TolMais"))
    B5_TOLME   = Trim(Consulta("TolMenos"))                       
    B5_COMPR   = Trim(Consulta("ComprimentoInterno"))
    B5_LARG    = Trim(Consulta("LarguraInterna"))
    B5_ALTU    = Trim(Consulta("AlturaInterna"))
    B5_ESTILO  = Trim(Consulta("TipoABNT"))
    B5_COMPRE  = Trim(Consulta("Compressao"))      
    B5_ESTRUTU = Trim(Consulta("Composicao"))              
    B5_FECHAM  = Trim(Consulta("Fechamento"))   
    B5_ESPES2  = ""
    B1_ALIMENT = ""
    B5_AMARRAD = Trim(Consulta("Amarrado"))
    B5_QTDPCT  = Trim(Consulta("PecasPorPacote"))
    B5_FITLAR  = Trim(Consulta("FitilhosLargura"))
    B5_FITCOM  = Trim(Consulta("FitilhosComprimento"))
    B5_PALETE  = Trim(Consulta("Paletizado"))
    B5_TPPALE  = Trim(Consulta("TipoPalete"))
    B5_DESCPAL = ""
    B5_FACOP   = Trim(Consulta("FacesOpostas"))  
    B5_CANTON  = Trim(Consulta("Cantoneiras"))   
    B5_ESPELHO = Trim(Consulta("Espelho"))    
    B5_TAMPA   = ""
    B5_FILME   = Trim(Consulta("Filme"))  
    B5_GRAM    = Trim(Consulta("Gramatura")) 
    B5_CLI     = Trim(Consulta("CodCliente")) 
    B5_NOMECLI = Trim(Consulta("RazaoSocial"))    
    B5_COLUNA  = Trim(Consulta("ColunaMinimo"))   
    B5_COBEXT  = Trim(Consulta("CobbExtMaximo"))
    B5_COBINT  = Trim(Consulta("CobbIntMaximo"))   
    B5_ESMAG   = Trim(Consulta("Compressao"))
    Z1_MULLEN  = Trim(Consulta("MullenMinimo"))
    Z1_ESPESS  = Trim(Consulta("Espessura"))
    Z1_COLUNA  = Trim(Consulta("ColunaMinimo"))  
    B5_COR1    = Trim(Consulta("Cor1"))
    B5_COR2    = Trim(Consulta("Cor2"))
    B5_COR3    = Trim(Consulta("Cor3"))
    B5_COR4    = Trim(Consulta("Cor4"))
    B5_ESPESS  = Z1_ESPESS
    B5_PCSCON  = Trim(Consulta("PecasPorConjunto"))  
    B5_PLLAST  = "0"
    B5_ARRPA   = Trim(Consulta("Arranjo")) 
    B5_PCSPAL  = Trim(Consulta("PecasPorPalete"))  
    B5_FITASL  = Trim(Consulta("FitasLarguraPalete")) 
    B5_FICASC  = Trim(Consulta("FitasComprimentoPalete")) 
    B5_LARGCH  = Trim(Consulta("Largura")) 
    B5_COMPCH  = Trim(Consulta("Comprimento")) 
    B5_ARRL    = Trim(Consulta("MultLarg")) 
    B5_ARRC    = Trim(Consulta("MultComp")) 
    B5_BLANKL  = "0"   
    B5_BLANKC  = "0"     
    B5_AREA    = Trim(Consulta("AreaBrutaChapa")) 
    B5_AREALI  = Trim(Consulta("AreaLiquidaChapa"))   
    B5_PESOB   = "0"
    B5_PESOL   = "0"
    B5_VINCO1  = Trim(Consulta("VincoLarg1"))
    B5_VINCO2  = Trim(Consulta("VincoLarg2"))
    B5_VINCO3  = Trim(Consulta("VincoLarg3"))
    B5_VINCO4  = Trim(Consulta("VincoLarg4"))
    B5_VINCO5  = Trim(Consulta("VincoLarg5"))
    B5_VINCO6  = Trim(Consulta("VincoLarg6"))   
    B5_VINCO7  = Trim(Consulta("VincoLarg7"))

    if IsNull(Trim(Consulta("VincoLarg1"))) then
      B5_VINCO1  = 0
    end if

    if IsNull(Trim(Consulta("VincoLarg2"))) then
      B5_VINCO2  = 0
    end if

    if IsNull(Trim(Consulta("VincoLarg3"))) then
      B5_VINCO3  = 0
    end if

    if IsNull(Trim(Consulta("VincoLarg4"))) then
      B5_VINCO4  = 0
    end if

    if IsNull(Trim(Consulta("VincoLarg5"))) then
      B5_VINCO5  = 0
    end if

    if IsNull(Trim(Consulta("VincoLarg6"))) then
      B5_VINCO6  = 0
    end if

    if IsNull(Trim(Consulta("VincoLarg7"))) then
      B5_VINCO7  = 0
    end if

    B5_COD     = FT1_FTNOVA
    B5_MAQ1    = Trim(Consulta("Maquina1"))
    B5_MAQ2    = Trim(Consulta("Maquina2"))
    B5_MAQ3    = Trim(Consulta("Maquina3"))
    B5_MAQ4    = Trim(Consulta("Maquina4"))
    B5_MAQ5    = Trim(Consulta("Maquina5"))
    B5_COLUNA  = B5_COLUNA
    B5_MULLEN  = Z1_MULLEN
    B5_VINCI1  = "0"  
    B5_VINCI2  = "0"  
    B5_VINCI3  = "0"  
    B5_VINCI4  = "0"  
    B5_VINCI5  = "0"  
    B5_NUMCLI  = ""
    B5_NUMFACA = Trim(Consulta("Faca"))   
    B1_CODPCP  = ""
    B5_VINCACE = ""
    B5_CORTACE = ""
    B5_OBSVEN  = ""
    B5_OBSDDP  = ""
    B5_OBS1    = Trim(Consulta("Obs1"))  
    B5_OBS2    = Trim(Consulta("Obs2"))  
    B5_OBS3    = Trim(Consulta("Obs3"))  
    B5_NOMVEN  = Trim(Consulta("Representante"))
    B5_USUARIO = FT1_USU
    B5_RESPONS = FT1_USU
    B5_VISCOR1 = Trim(Consulta("ViscosidadeCor1"))    
    B5_VISCOR2 = Trim(Consulta("ViscosidadeCor2"))    
    B5_VISCOR3 = Trim(Consulta("ViscosidadeCor3"))    
    B1_INPUBE  = "05"
    B5_DTCADFT = Trim(Consulta("DataCriacao"))    
  else
    B5_REF     = ""
    B5_TOLMA   = "0"
    B5_TOLME   = "0"
    B5_COMPR   = "0"
    B5_LARG    = "0"
    B5_ALTU    = "0"
    B5_ESTILO  = ""
    B5_COMPRE  = ""
    B5_ESTRUTU = ""
    B5_FECHAM  = ""
    B5_ESPES2  = ""
    B1_ALIMENT = ""
    B5_AMARRAD = ""
    B5_QTDPCT  = "0"
    B5_FITLAR  = "0"
    B5_FITCOM  = "0"
    B5_PALETE  = ""
    B5_TPPALE  = ""
    B5_DESCPAL = ""
    B5_FACOP   = ""
    B5_CANTON  = ""
    B5_ESPELHO = ""
    B5_TAMPA   = ""
    B5_FILME   = ""
    B5_GRAM    = ""
    B5_CLI     = ""
    B5_NOMECLI = ""
    B5_COLUNA  = ""
    B5_COBEXT  = ""
    B5_COBINT  = ""
    B5_ESMAG   = ""
    Z1_MULLEN  = ""
    Z1_ESPESS  = ""
    Z1_COLUNA  = ""
    B5_COR1    = ""
    B5_COR2    = ""
    B5_COR3    = ""
    B5_COR4    = ""
    B5_ESPESS  = "0"
    B5_PCSCON  = "0"
    B5_PLLAST  = "0"
    B5_ARRPA   = "0"
    B5_PCSPAL  = "0"
    B5_FITASL  = "0"
    B5_FICASC  = "0"
    B5_LARGCH  = "0"
    B5_COMPCH  = "0"
    B5_ARRL    = "0"
    B5_ARRC    = "0"
    B5_BLANKL  = "0"
    B5_BLANKC  = "0"
    B5_AREA    = "0"
    B5_AREALI  = "0"
    B5_PESOB   = "0"
    B5_PESOL   = "0"
    B5_VINCO1  = "0"
    B5_VINCO2  = "0"
    B5_VINCO3  = "0"
    B5_VINCO4  = "0"
    B5_VINCO5  = "0"
    B5_VINCO6  = "0"
    B5_VINCO7  = "0"
    B5_COD     = ""
    B5_MAQ1    = ""
    B5_MAQ2    = ""
    B5_MAQ3    = ""
    B5_MAQ4    = ""
    B5_MAQ5    = ""
    B5_COLUNA  = ""
    B5_MULLEN  = ""
    B5_VINCI1  = "0"
    B5_VINCI2  = "0"
    B5_VINCI3  = "0"
    B5_VINCI4  = "0"
    B5_VINCI5  = "0"
    B5_NUMCLI  = ""
    B5_NUMFACA = ""
    B1_CODPCP  = ""
    B5_VINCACE = ""
    B5_CORTACE = ""
    B5_OBSVEN  = ""
    B5_OBSDDP  = ""
    B5_OBS1    = ""
    B5_OBS2    = ""
    B5_OBS3    = ""
    B5_NOMVEN  = ""
    B5_USUARIO = ""
    B5_RESPONS = ""
    B5_VISCOR1 = "0"
    B5_VISCOR2 = "0"
    B5_VISCOR3 = "0"
    B1_INPUBE  = ""
    B5_DTCADFT = "" 
  end if

  Consulta.Close()
  SQL = "" &_
    "SELECT " &_
      "FT1_QuemEnviouAmostra, " &_
      "FT1_DataEnvioAmostra, " &_
      "CASE '" & B5_REF     & "' WHEN '' then FT1_B5REF      ELSE '" & B5_REF       & "' END AS FT1_B5REF, " &_
      "CASE  " & B5_TOLMA   & "  WHEN 0  then FT1_B5TOLMA    ELSE  " & B5_TOLMA     & "  END AS FT1_B5TOLMA, " &_
      "CASE  " & B5_TOLME   & "  WHEN 0  then FT1_B5TOLME    ELSE  " & B5_TOLME     & "  END AS FT1_B5TOLME, " &_
      "CASE  " & B5_COMPR   & "  WHEN 0  then FT1_B5COMPR    ELSE  " & B5_COMPR     & "  END AS FT1_B5COMPR, " &_
      "CASE  " & B5_LARG    & "  WHEN 0  then FT1_B5LARG     ELSE  " & B5_LARG      & "  END AS FT1_B5LARG, " &_
      "CASE  " & B5_ALTU    & "  WHEN 0  then FT1_B5ALTU     ELSE  " & B5_ALTU      & "  END AS FT1_B5ALTU, " &_
      "CASE '" & B5_ESTILO  & "' WHEN '' then FT1_B5ESTILO   ELSE '" & B5_ESTILO    & "' END AS FT1_B5ESTILO, " &_
      "CASE '" & B5_COMPRE  & "' WHEN 0  then FT1_B5COMPRE   ELSE '" & B5_COMPRE    & "' END AS FT1_B5COMPRE, " &_
      "CASE '" & B5_ESTRUTU & "' WHEN '' then FT1_B5ESTRUTU  ELSE '" & B5_ESTRUTU   & "' END AS FT1_B5ESTRUTU, " &_
      "CASE '" & B5_FECHAM  & "' WHEN '' then FT1_B5FECHAM   ELSE '" & B5_FECHAM    & "' END FT1_B5FECHAM, " &_
      "CASE '" & B5_ESPES2  & "' WHEN '' then FT1_B5ESPES2   ELSE '" & B5_ESPES2    & "' END AS FT1_B5ESPES2, " &_
      "CASE '" & B1_ALIMENT & "' WHEN '' then FT1_B1ALIMENT  ELSE '" & B1_ALIMENT   & "' END AS FT1_B1ALIMENT, " &_
      "CASE '" & B5_AMARRAD & "' WHEN '' then FT1_B5AMARRAD  ELSE '" & B5_AMARRAD   & "' END AS FT1_B5AMARRAD, " &_
      "CASE  " & B5_QTDPCT  & "  WHEN 0  then FT1_B5QTDPCT   ELSE  " & B5_QTDPCT    & "  END AS FT1_B5QTDPCT, " &_
      "CASE  " & B5_FITLAR  & "  WHEN 0  then FT1_B5FITLAR   ELSE  " & B5_FITLAR    & "  END AS FT1_B5FITLAR, " &_
      "CASE  " & B5_FITCOM  & "  WHEN 0  then FT1_B5FITCOM   ELSE  " & B5_FITCOM    & "  END AS FT1_B5FITCOM, " &_
      "CASE '" & B5_PALETE  & "' WHEN '' then FT1_B5PALETE   ELSE '" & B5_PALETE    & "' END AS FT1_B5PALETE, " &_
      "CASE '" & B5_TPPALE  & "' WHEN '' then FT1_B5TPPALE   ELSE '" & B5_TPPALE    & "' END AS FT1_B5TPPALE, " &_
      "CASE '" & B5_DESCPAL & "' WHEN '' then '"& Descricao &"' ELSE '"& B5_DESCPAL & "' END AS B5_DESCPAL, " &_
      "CASE '" & B5_FACOP   & "' WHEN '' then FT1_B5FACOP    ELSE '" & B5_FACOP     & "' END AS FT1_B5FACOP, " &_
      "CASE '" & B5_CANTON  & "' WHEN '' then FT1_B5CANTON   ELSE '" & B5_CANTON    & "' END AS FT1_B5CANTON, " &_
      "CASE '" & B5_ESPELHO & "' WHEN '' then FT1_B5ESPELHO  ELSE '" & B5_ESPELHO   & "' END AS FT1_B5ESPELHO, " &_
      "CASE '" & B5_TAMPA   & "' WHEN '' then FT1_B5TAMPA    ELSE '" & B5_TAMPA     & "' END AS FT1_B5TAMPA, " &_
      "CASE '" & B5_FILME   & "' WHEN '' then FT1_B5FILME    ELSE '" & B5_FILME     & "' END AS FT1_B5FILME, " &_
      "CASE '" & B5_GRAM    & "' WHEN '' then FT1_B5GRAM     ELSE '" & B5_GRAM      & "' END FT1_B5GRAM, " &_
      "CASE '" & B5_CLI     & "' WHEN '' then FT1_B5CLI      ELSE '" & B5_CLI       & "' END FT1_B5CLI, " &_
      "CASE '" & B5_NOMECLI & "' WHEN '' then FT1_B5NOMECLI  ELSE '" & B5_NOMECLI   & "' END FT1_B5NOMECLI, " &_
      "CASE '" & B5_COLUNA  & "' WHEN '' then FT1_B5COLUNA   ELSE '" & B5_COLUNA    & "' END FT1_B5COLUNA, " &_
      "CASE '" & B5_COBEXT  & "' WHEN '' then FT1_B5COBEXT   ELSE '" & B5_COBEXT    & "' END FT1_B5COBEXT, " &_
      "CASE '" & B5_COBINT  & "' WHEN '' then FT1_B5COBINT   ELSE '" & B5_COBINT    & "' END FT1_B5COBINT, " &_
      "CASE '" & B5_ESMAG   & "' WHEN '' then FT1_B5ESMAG    ELSE '" & B5_ESMAG     & "' END FT1_B5ESMAG, " &_
      "isnull(FT1_B5AREA,0) FT1_B5AREA, " &_
      "isnull(FT1_B5GRAM,0) FT1_B5GRAM, " &_
      "FT1_LOJA, " &_
      "FT1_ID, " &_
      "FT1_TIPO, " &_
      "FT1_FT, " &_
      "FT1_RACK, " &_
      "FT1_OVERHANG, " &_
      "FT1_UNIREL, " &_
      "FT1_TEMPOEST, " &_
      "FT1_CONSMES, " &_
      "FT1_QTDPRI, " &_
      "FT1_PESOPRI, " &_
      "FT1_B5CANTON, " &_
      "FT1_COMPRIMENTOCHAPA, " &_
      "FT1_LARGURACHAPA, " &_
      "FT1_OBSP, " &_
      "FT1_INKJET, " &_
      "FT1_LINAUTO, " &_
      "FT1_TIPOPRD, " &_
      "FT1_OBSERV, " &_
      "FT1_B5COBINT, " &_
      "FT1_ALTPAL, " &_
      "FT1_SUPORTEPRIMARIO, " &_
      "FT1_DATACORTEAMOSTRA, " &_
      "FT1_Margem, " &_
      "FT1_Resultado, " &_
      "FT1_QTDAMO, " &_
      "FT1_CON, " &_
      "FT1_END, " &_
      "FT1_BAI, " &_
      "FT1_MUN, " &_
      "FT1_CEP, " &_
      "FT1_EMAIL, " &_
      "FT1_TEL, " &_
      "FT1_MULLEN, " &_
      "FT1_USU, " &_
      "FT1_DATA, " &_
      "FT1_TOLPESO, " &_
      "FT1_A3NREDUZ, " &_
      "FT1_A1VEND, " &_
      "FT1_DataGeracaoFT, " &_
      "FT1_VALORCLICHE, " &_
      "FT1_VALORMATRIZ, " &_
      "FT1_TECNICO, " &_
      "FT1_Processo, " &_
      "FT1_DataStatusAmostra, " &_
      "FT1_HoraStatusAmostra, " &_
      "FT1_UsuStatusAmostra, " &_
      "FT1_FormaEnvioAmostra, " &_
      "FT1_TESTECOLUNA, " &_
      "FT1_TESTEGRAMATURA, " &_
      "FT1_TESTECOBINTERNO, " &_
      "FT1_TESTECOBEXTERNO, " &_
      "FT1_TESTEMULLENINTERNOCAPA, " &_
      "FT1_TESTEMULLENINTERNOMIOLO, " &_
      "FT1_STATUS, " &_
      "FT1_TESTEMULLENINTERNOCAPAINTERMEDIARIA, " &_
      "FT1_TESTEMULLENEXTERNOCAPA, " &_
      "FT1_TESTEMULLENEXTERNOMIOLO, " &_
      "FT1_COMPRESSAO1, " &_
      "FT1_COMPRESSAO2, " &_
      "FT1_COMPRESSAO3, " &_
      "(ISNULL(FT1_COMPRESSAO1,0) + ISNULL(FT1_COMPRESSAO2,0) + ISNULL(FT1_COMPRESSAO3,0) / 3) AS MEDIA_COMPRESSAO, " &_
      "'" & B5_COBEXT & "' Z1_COBBE, " &_
      "'" & B5_COBINT & "' Z1_COBBI, " &_
      "'" & B5_ESMAG  & "' Z1_ESMAG, " &_
      "'" & Z1_MULLEN & "' Z1_MULLEN, " &_
      "'" & B5_GRAM   & "' Z1_GRAM, " &_
      "'" & Z1_ESPESS & "' Z1_ESPESS, " &_
      "'" & Z1_COLUNA & "' Z1_COLUNA, " &_
      "'" & B5_COBEXT & "' Z1_COBBE, " &_
      "'" & B5_COBINT & "' Z1_COBBI, " &_
      "'" & B5_ESMAG  & "' Z1_ESMAG, " &_
      "'" & Z1_MULLEN & "' Z1_MULLEN, " &_
      "'" & B5_GRAM   & "' Z1_GRAM, " &_
      "'" & Z1_ESPESS & "' Z1_ESPESS, " &_
      "'' B5_COR1, " &_
      "'' B5_COR2, " &_
      "'' B5_COR3, " &_
      "'' B5_COR4, " &_
      "'" & Z1_ESPESS & "' B5_ESPESS, " &_
      "'" & B5_COBEXT & "' B5_COBEXT, " &_
      B5_PCSCON & " B5_PCSCON, " &_
      "'" & B5_ESMAG & "' B5_ESMAG, " &_
      "'" & B5_GRAM  & "' B5_GRAM, " &_
      B5_PLLAST & " B5_PLLAST, " &_
      B5_ARRPA  & " B5_ARRPA, " &_
      B5_PCSPAL & " B5_PCSPAL, " &_
      B5_FITASL & " B5_FITASL, " &_
      B5_FICASC & " B5_FICASC, " &_
      B5_LARGCH & " B5_LARGCH, " &_
      B5_COMPCH & " B5_COMPCH, " &_
      B5_ARRL   & " B5_ARRL, " &_
      B5_ARRC   & " B5_ARRC, " &_
      B5_BLANKL & " B5_BLANKL, " &_
      B5_BLANKC & " B5_BLANKC, " &_
      B5_AREA   & " B5_AREA, " &_
      B5_AREALI & " B5_AREALI, " &_
      B5_PESOB  & " B5_PESOB, " &_
      B5_PESOL  & " B5_PESOL, " &_
      B5_VINCO1 & " B5_VINCO1, " &_
      B5_VINCO2 & " B5_VINCO2, " &_
      B5_VINCO3 & " B5_VINCO3, " &_
      B5_VINCO4 & " B5_VINCO4, " &_
      B5_VINCO5 & " B5_VINCO5, " &_
      "0 B5_VINCO6, " &_
      "0 B5_VINCO7, " &_
      "'" & B5_COD  & "' B5_COD, " &_
      "'" & B5_MAQ1 & "' B5_MAQ1, " &_
      "'" & B5_MAQ2 & "' B5_MAQ2, " &_
      "'" & B5_MAQ3 & "' B5_MAQ3, " &_
      "'" & B5_MAQ4 & "' B5_MAQ4, " &_
      "'" & B5_MAQ5 & "' B5_MAQ5, " &_
      "'" & B5_COLUNA & "' B5_COLUNA, " &_
      "'" & B5_MULLEN & "' B5_MULLEN, " &_
      "'" & B5_COBINT & "' B5_COBINT, " &_
      "0 B5_VINCI1, " &_
      "0 B5_VINCI2, " &_
      "0 B5_VINCI3, " &_
      "0 B5_VINCI4, " &_
      "0 B5_VINCI5, " &_
      "'' B5_NUMCLI, " &_
      "'" & B5_NUMFACA & "' B5_NUMFACA, " &_
      "'' B1_CODPCP, " &_
      "'' B5_VINCACE, " &_
      "'' B5_CORTACE, " &_
      "'' B5_OBSVEN, " &_
      "'' B5_OBSDDP, " &_
      "'" & B5_OBS1 & "' B5_OBS1, " &_
      "'" & B5_OBS2 & "' B5_OBS2, " &_
      "'" & B5_OBS3 & "' B5_OBS3, " &_
      "'" & B5_NOMVEN & "' B5_NOMVEN, " &_
      "'" & B5_USUARIO & "' B5_USUARIO, " &_
      "'" & B5_USUARIO & "' B5_RESPONS, " &_
      "'" & B5_VISCOR1 & "' B5_VISCOR1, " &_
      "'" & B5_VISCOR2 & "' B5_VISCOR2, " &_
      "'" & B5_VISCOR3 & "' B5_VISCOR3, " &_
      " FT1_B1INPUBE, " &_
      "'" & B5_DTCADFT & "'  B5_DTCADFT " &_
    "FROM SKFT1 " &_
    "WHERE ISNULL(FT1_DEL,0) = 0 " &_
      "AND FT1_ID = "&Trim(request("Id"))

else ' not (Request("Planta") = "05" or Request("Planta") = "06" )'
  SQL = "" &_
    "SELECT " &_
      "FT1_QuemEnviouAmostra, " &_
      "FT1_DataEnvioAmostra, " &_
      "FT1_B5REF, " &_
      "FT1_B5TOLMA, " &_
      "FT1_B5TOLME, " &_
      "FT1_B5COMPR, " &_
      "FT1_B5LARG, " &_
      "FT1_B5ALTU, " &_
      "FT1_B5ESTILO, " &_
      "FT1_B5COMPRE, " &_
      "FT1_B5ESTRUTU, " &_
      "FT1_B5FECHAM, " &_
      "FT1_B5ESPES2, " &_
      "FT1_B1ALIMENT, " &_
      "FT1_B5AMARRAD, " &_
      "FT1_B5QTDPCT, " &_
      "FT1_B5FITLAR, " &_
      "FT1_B5FITCOM, " &_
      "FT1_B5PALETE, " &_
      "FT1_B5TPPALE, " &_
      "FT1_B5FACOP, " &_
      "FT1_B5CANTON, " &_
      "FT1_B5ESPELHO, " &_
      "FT1_B5TAMPA, " &_
      "FT1_B5FILME, " &_
      "FT1_B5GRAM, " &_
      "FT1_B5CLI, " &_
      "FT1_B5NOMECLI, " &_
      "FT1_B5COLUNA, " &_
      "FT1_B5COBEXT, " &_
      "FT1_B5COBINT, " &_
      "FT1_B5ESMAG, " &_
      "FT1_B5AREA, " &_
      "FT1_B5GRAM, " &_
      "FT1_LOJA, " &_
      "FT1_ID, " &_
      "FT1_TIPO, " &_
      "FT1_FT, " &_
      "FT1_RACK, " &_
      "FT1_OVERHANG, " &_
      "FT1_UNIREL, " &_
      "FT1_TEMPOEST, " &_
      "FT1_CONSMES, " &_
      "FT1_QTDPRI, " &_
      "FT1_PESOPRI, " &_
      "FT1_B5CANTON, " &_
      "FT1_COMPRIMENTOCHAPA, " &_
      "FT1_LARGURACHAPA, " &_
      "FT1_OBSP, " &_
      "FT1_INKJET, " &_
      "FT1_LINAUTO, " &_
      "FT1_TIPOPRD, " &_
      "FT1_OBSERV, " &_
      "FT1_B5COBINT, " &_
      "FT1_ALTPAL, " &_
      "FT1_SUPORTEPRIMARIO, " &_
      "FT1_DATACORTEAMOSTRA, " &_
      "FT1_Margem, " &_
      "FT1_Resultado, " &_
      "FT1_QTDAMO, " &_
      "FT1_CON, " &_
      "FT1_END, " &_
      "FT1_BAI, " &_
      "FT1_MUN, " &_
      "FT1_CEP, " &_
      "FT1_EMAIL, " &_
      "FT1_TEL, " &_
      "FT1_MULLEN, " &_
      "FT1_USU, " &_
      "FT1_DATA, " &_
      "FT1_TOLPESO, " &_
      "FT1_A3NREDUZ, " &_
      "FT1_A1VEND, " &_
      "FT1_DataGeracaoFT, " &_
      "FT1_VALORCLICHE, " &_
      "FT1_VALORMATRIZ, " &_
      "FT1_TECNICO, " &_
      "FT1_Processo, " &_
      "FT1_DataStatusAmostra, " &_
      "FT1_HoraStatusAmostra, " &_
      "FT1_UsuStatusAmostra, " &_
      "FT1_FormaEnvioAmostra, " &_
      "FT1_TESTECOLUNA, " &_
      "FT1_TESTEGRAMATURA, " &_
      "FT1_TESTECOBINTERNO, " &_
      "FT1_TESTECOBEXTERNO, " &_
      "FT1_TESTEMULLENINTERNOCAPA, " &_
      "FT1_TESTEMULLENINTERNOMIOLO, " &_
      "FT1_STATUS, " &_
      "FT1_TESTEMULLENINTERNOCAPAINTERMEDIARIA, " &_
      "FT1_TESTEMULLENEXTERNOCAPA, " &_
      "FT1_TESTEMULLENEXTERNOMIOLO, " &_
      "FT1_COMPRESSAO1, " &_
      "FT1_COMPRESSAO2, " &_
      "FT1_COMPRESSAO3, " &_
      "(ISNULL(FT1_COMPRESSAO1,0) + ISNULL(FT1_COMPRESSAO2,0) + ISNULL(FT1_COMPRESSAO3,0) / 3) MEDIA_COMPRESSAO, " &_
      "FT1_B1INPUBE " &_
    "FROM SKFT1 " &_
    "WHERE ISNULL(FT1_DEL,0) = 0 " &_
      "AND FT1_ID = "&Trim(request("Id"))
end if 
Consulta1.Open(SQL)
%>
<body>
  <div class="container-fluid body-content margem">
    <div class="row">
      <div class="col-md-10">
        <table>
          <tr>
            <td>
              <table class="table table-condensed" style="margin-bottom:0px">
                <thead>
                  <tr class="CabecalhoDestaque">
                    <th colspan="4" class="text-center" style="border-top:2px solid #ddd">DADOS DO PRODUTO</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td class="BordaEsquerda" style="border-top:1px solid black"><strong>Cliente:</strong></td>
                    <td colspan="3" class="BordaDireita"><%=Consulta1("FT1_B5NOMECLI")%>&nbsp;&nbsp;<%=Consulta1("FT1_B5CLI") %>&nbsp;<%=Consulta1("FT1_LOJA") %></td>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda"><strong>Referência:</strong></td>
                    <td colspan="3" class="BordaDireita" ><%=Consulta1("FT1_B5REF")%></td>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda"><strong>Representante:</strong></td>
                    <td colspan="3" class="BordaDireita" ><%=Consulta1("FT1_A1VEND")%>-<%=Consulta1("FT1_A3NREDUZ")%></td>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda"><strong>Cod. FEFCO:</strong></td>
                    <td><%=Consulta1("FT1_B5ESTILO")%></td>                                        
<%
                  select case Consulta1("FT1_OBSP")
                  case "F" observacaoProduto = "Furos"
                  case "C" observacaoProduto = "Chanfros"       
                  case "E" observacaoProduto = "Com furos/chanfros"
                  case "D" observacaoProduto = "Com furo alça"
                  case "A" observacaoProduto = "Caixa com display"
                  end select
%>
                    <td ><strong>Observações:</strong></td>
                    <td colspan="3" class="BordaDireita"><%=observacaoProduto%></td>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda"><strong>Medida Int.(CxLxA):</strong></td>
                    <td><%=Consulta1("FT1_B5COMPR")%></td>
                    <td><%=Consulta1("FT1_B5LARG")%></td>
                    <td class="BordaDireita"><%=Consulta1("FT1_B5ALTU")%></td>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda"><strong>Comp Chapa (m):</strong></td>
                    <td><%=Consulta1("FT1_COMPRIMENTOCHAPA") / 1000 %></td>
                    <td><strong>Larg. Chapa (m)</strong></td>
                    <td class="BordaDireita"><%=Consulta1("FT1_LARGURACHAPA") / 1000 %></td>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda"><strong>Área Blank (m<sup>2</sup>)</strong></td>
                    <td><%=(Consulta1("FT1_COMPRIMENTOCHAPA") / 1000) * (Consulta1("FT1_LARGURACHAPA") / 1000)%></td>
                    <td><strong>Peso (g)</strong></td>     
<% 
                    if trim(Consulta1("FT1_B5AREA")) <> "" AND trim(Consulta1("FT1_B5GRAM")) <> "" then 
%>
                    <td class="BordaDireita"><%=CDbl(Consulta1("FT1_B5AREA")) * CDbl(Consulta1("FT1_B5GRAM")) %></td> 
<% 
                    else 
%>
                    <td class="BordaDireita">0</td> 
<% 
                    end if 
%>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda"><strong>Tol % (+/-):</strong></td>
                    <td><%=Consulta1("FT1_B5TOLMA")%>/<%=Consulta1("FT1_B5TOLME")%></td>
                    <td><strong>Tolerância de Peso (± %):</strong></td>
                    <td class="BordaDireita"><%=Consulta1("FT1_TOLPESO")%></td>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda"><strong>Acompanhamento:</strong></td>
                    <td></td>
                    <td><strong>Planta Fabricação:</strong></td>
                    <td class="BordaDireita">
<%
                  select case Consulta1("FT1_B1INPUBE")
                         case "I" 
%>
                              Pirapetinga
<%
                         case "U" 
%>
                              Uberaba                                
<%
                         case "F" 
%>                         
                              Fortaleza
<%                              
                         case "B"
%>                         
                              Bento Gonçalves
<%                         
                         case "S"
%>                         
                               Saquarema
<%                               
                  end select
%>
                    </td>
                  </tr>
                </tbody>
              </table>
              <table class="table table-condensed" style="margin-bottom:0px">
                <thead>
                  <tr class="CabecalhoDestaque">
                    <th colspan="4" class="text-center" style="border-top:2px solid #ddd">Especificações Estruturais Smurfit Westrock / Cliente</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td class="BordaEsquerda" ><strong>Estrutura SW:</strong></td>
                    <td colspan="3" class="BordaDireita "><strong><%=Consulta1("FT1_B5ESTRUTU") %></strong></td>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda"><strong>ECT/Coluna(kgf/cm):</strong></td>
                    <td class=" "><%=Trim(Consulta1("FT1_B5COLUNA"))%> <%'=Trim(Consulta1("Z1_COLUNA"))%></td>
                    <td ><strong>Gramatura(g/m²):</strong></td>
                    <td  class="BordaDireita "><%= Consulta1("FT1_B5GRAM")%> <%'=Consulta1("Z1_GRAM") %></td>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda" ><strong>Esmag.(kgf/cm²):</strong></td>
                    <td class=""><%=Consulta1("FT1_B5ESMAG")%> <%'=Consulta1("Z1_ESMAG") %></td>
                    <td ><strong>Espessura(mm):</strong></td>
                    <td class="BordaDireita "><%=Consulta1("FT1_B5ESPES2")%> <%'=Consulta1("Z1_ESPESS") %></td>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda" ><strong>BCT/Compressão(kgf):</strong></td>
                    <td class=""><%=Consulta1("FT1_B5COMPRE")%></td>
                    <td ><strong>COBB Int (g/M²H<sub>2</sub>O):</strong></td>
                    <td class="BordaDireita "><%=Consulta1("FT1_B5COBINT")%> <%'=Consulta1("Z1_COBBI")%></td>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda" ><strong>Mull Refer(kgf/cm²):</strong></td>
                    <td><%=Consulta1("FT1_MULLEN")%> <%'=Consulta1("Z1_MULLEN") %></td>
                    <td ><strong>COBB Ext (g/M²H<sub>2</sub>O):</strong></td>
                    <td class="BordaDireita"><%=Consulta1("FT1_B5COBEXT")%> <%'=Consulta1("Z1_COBBE") %></td>
                  </tr>
                </tbody>
              </table>
<%
          SQL = "" &_
            "SELECT " &_
              "FT2_ID, " &_
              "FT2_ITEM, " &_
              "FT2_ACEQTD, " &_
              "FT2_ACEFEFCO, " &_
              "FT2_ACECORTE, " &_
              "FT2_ACELARG, " &_
              "FT2_ACECOMP, " &_
              "FT2_ACEESTR " &_
            "FROM SKFT2 " &_
            "WHERE FT2_ID = " & Consulta1("FT1_ID") & " " &_
            "ORDER BY FT2_ITEM "

          Dados12.Open(SQL)

          x = 0
          if Not Dados12.EOF then
%>
              <table class="table table-condensed" style="margin-bottom:0px">
                <thead>
                  <tr class="CabecalhoDestaque">
                    <th colspan="4" class="text-center" style="border-top:2px solid #ddd">Especificações Acessório</th>
                  </tr>
                </thead>
                <tbody>
<%
            while not Dados12.EOF
              x = x + 1
%>
                  <tr>
                    <td class="BordaEsquerda" ><strong>Acessório Estilo:</strong></td>
                    <td><%=Dados12("FT2_ACEFEFCO") %></td>
                    <td><strong>Acessório Estilo:</strong></td>
                    <td class="BordaDireita"><%=Dados12("FT2_ACECOMP") %> x <%=Dados12("FT2_ACELARG") %></td>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda" ><strong>Estrutura:</strong></td>
                    <td><strong></strong><%=Dados12("FT2_ACEESTR") %></td>
                    <td><strong>Peças Conjunto:</strong></td>
                    <td class="BordaDireita"><%=Dados12("FT2_ACEQTD") %></td>
                  </tr>
<%
              Dados12.MoveNext()
            wend
%>
                </tbody>
              </table>
<%
          end if
          Dados12.Close()
%>
              <table class="table table-condensed" style="margin-bottom:0px">
                <thead>
                  <tr class="CabecalhoDestaque">
                    <th colspan="4" class="text-center" style="border-top:2px solid #ddd">Testes Físicos da Amostra</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td class="BordaEsquerda text-center"><strong>ECT/Coluna(kgf/cm):</strong></td>
                    <td class="BordaEsquerda text-center"><strong>Gramatura (g/m<sup>2</sup>)</strong></td>
                    <td class="BordaEsquerda text-center"><strong>Cobb Interno (g/M²H<sub>2</sub>O)</strong></td>
                    <td class="BordaEsquerda BordaDireita text-center"><strong>Cobb Externo (g/M²H<sub>2</sub>O)</strong></td>
                  </tr>
                  <tr>
                    <td class="BordaEsquerda text-center"><%=Consulta1("FT1_TESTECOLUNA") %></td>
                    <td class="BordaEsquerda text-center"><%=Consulta1("FT1_TESTEGRAMATURA") %></td>
                    <td class="BordaEsquerda text-center"><%=Consulta1("FT1_TESTECOBINTERNO") %></td>
                    <td class="BordaEsquerda BordaDireita text-center text-nowrap"><%=Consulta1("FT1_TESTECOBEXTERNO") %><strong></td>
                  </tr>
                  <tr>
                    <td rowspan="5" class="BordaEsquerda BordaInferior BordaDireita text-center" style="vertical-align:central"><strong>Descolamento de Capas (g/m<sup>2</sup>)</strong></td>
                    <td><strong>Capa Interna</strong></td>
                    <td class="BordaDireita"><%=Consulta1("FT1_TESTEMULLENINTERNOCAPA") %></td>
                    <td class="BordaDireita"><strong>BCT/Compressão (kgf)</strong></td>
                  </tr>
                  <tr>
                    <td><strong>Miolo 1</strong></td>
                    <td class="BordaDireita"><%=Consulta1("FT1_TESTEMULLENINTERNOMIOLO") %></td>
                    <td class="BordaDireita"><%=Consulta1("FT1_COMPRESSAO1") %></td>
                  </tr>
                  <tr>
                    <td><strong>Miolo 2</strong></td>
                    <td class="BordaDireita"><%=Consulta1("FT1_TESTEMULLENINTERNOCAPAINTERMEDIARIA") %></td>
                    <td class="BordaDireita"><%=Consulta1("FT1_COMPRESSAO2") %></td>
                  </tr>
                  <tr>
                    <td><strong>Miolo 3</strong></td>
                    <td class="BordaDireita"><%=Consulta1("FT1_TESTEMULLENEXTERNOMIOLO") %></td>
                    <td class="BordaDireita"><%=Consulta1("FT1_COMPRESSAO3") %></td>
                  </tr>
                  <tr>
                    <td class="BordaInferior"><strong>Capa Externa</strong></td>
                    <td class="BordaInferior BordaDireita"><%=Consulta1("FT1_TESTEMULLENEXTERNOCAPA") %></td>
                    <td class="BordaInferior BordaDireita">Média (<%=Consulta1("MEDIA_COMPRESSAO") %>)</td>
                  </tr>
                </tbody>
              </table>
            </td>
            <td>
              <div class="col-md-10">
                <table class="table table-condensed" style="margin-bottom:0px">
                  <tbody>
                    <tr>
                      <td class="BordaEsquerda"><h4>SAP: <%=Consulta1("FT1_ID") %></h4></td>
                      <td class="BordaDireita"><h4>FT: <%=FT1_FTNOVA %></h4></td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div class="col-md-10">
                <table class="table table-condensed" style="margin-bottom:0px">
                  <thead>
                    <tr class="CabecalhoDestaque">
                      <th colspan="4" class="text-center" style="border-top:2px solid #ddd">Dados da Solicitação</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
<%
                  select case Consulta1("FT1_TIPO") 
                         case "1" tipoPreCadastro = "Projetar/alterar exatamente conforme desenho ou amostra enviado ao Desenvolvimento"
                         case "2" tipoPreCadastro = "Projetar/alterar exatamente conforme especificações identificadas nesta Solicitação"
                         case "3" tipoPreCadastro = "Alterar ou Cancelar uma Ficha Técnica "&Trim(Consulta1("FT1_FT"))
                         case "4" tipoPreCadastro = "Inovação"
                  end select

                  SQL = "SELECT CRM_NOME FROM HELPDESK..CRM_DTI WHERE CRM_COD = '"&Consulta1("FT1_TECNICO")&"' AND ISNULL(ValidaPreCadastroProduto,0) = 1 "
                  Dados3.Open(SQL)        
                  tecnicoResponsavel = iif( (NOT Dados3.EOF), Trim(Dados3("CRM_NOME")), "" )
                  Dados3.Close()
%>
                      <td class="BordaEsquerda BordaDireita text-center" colspan="4"><%=tipoPreCadastro %></td>
                    </tr>
                    <tr>
                      <td class="BordaEsquerda" ><strong>Técnico/Desenvolvedor:</strong></td>
                      <td><%=tecnicoResponsavel %></td>
                      <td><strong>Processo:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_Processo") %></td>
                    </tr>
                    <tr>
                      <td class="BordaEsquerda" ><strong>Andamento processo:</strong></td>
                      <td><%=Consulta1("FT1_DataStatusAmostra") %> / <%=Consulta1("FT1_HoraStatusAmostra") %></td>
                      <td><strong>Status Amostra:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_UsuStatusAmostra") %></td>
                    </tr>
                  </tbody>
                </table>
              </div>
<%
          if Cint(Consulta1("FT1_QTDAMO")) > 0 then
%>
              <div class="col-md-10">
                <table class="table table-condensed" style="margin-bottom:0px">
                  <thead>
                    <tr class="CabecalhoDestaque">
                      <th colspan="4" class="text-center" style="border-top:2px solid #ddd">Amostras</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td class="BordaEsquerda" ><strong>Quant:</strong></td>
                      <td class=" BordaDireita"><%=Consulta1("FT1_QTDAMO") %></td>
                      <td><strong>Aos Cuidados:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_CON") %></td>
                    </tr>
                    <tr>
                      <td class="BordaEsquerda" ><strong>Endereço:</strong></td>
                      <td class=" BordaDireita"><%=Consulta1("FT1_END") %></td>
                      <td ><strong>Tel:</strong></td>
                      <td class="BordaDireita "><%=Consulta1("FT1_TEL") %></td>
                    </tr>
                    <tr>
                      <td class="BordaEsquerda"><strong>Bairro:</strong></td>
                      <td class=" BordaDireita"><%=Consulta1("FT1_BAI") %></td>
                      <td ><strong>CNPJ/CPF:</strong></td>
                      <td class="BordaDireita "></td>
                    </tr>
                    <tr>
                      <td class="BordaEsquerda" ><strong>Município:</strong></td>
                      <td  class=" BordaDireita"><%=Consulta1("FT1_MUN") %></td>
                      <td ><strong>Modelo de Envio:</strong></td>
                      <td class="BordaDireita "><%=Consulta1("FT1_FormaEnvioAmostra") %></td>
                    </tr>
                    <tr>
                      <td class="BordaEsquerda "><strong>CEP:</strong></td>
                      <td class="BordaDireita "><%=Consulta1("FT1_CEP") %></td>
                      <td ><strong>Código de rastreio/placa:</strong></td>
                      <td class="BordaDireita "></td>
                    </tr>
<%
                If len(Consulta1("FT1_EMAIL")) > 80 then
                  emailCliente = Mid(Consulta1("FT1_EMAIL"),1,80)&"..."
                else
                  emailCliente = Consulta1("FT1_EMAIL")
                end if
%>
                      <td><%=emailCliente%></td>
                      <td ><strong>Tel:</strong></td>
                      <td class="BordaDireita "><%=Consulta1("FT1_TEL") %></td>
                    </tr>
                  </tbody>
                </table>
              </div>
<%
              end if 
%>
              <div class="col-md-10">
                <table class="table table-condensed" style="margin-bottom:0px">
                  <thead>
                    <tr class="CabecalhoDestaque">
                      <th colspan="4" class="text-center" style="border-top:2px solid #ddd">Dados de utilização da embalagem</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
<%
                SQL = "SELECT COD, DSC FROM SKALI WHERE ATV = 1 AND DEL = 0 AND COD = '"&Consulta1("FT1_B1ALIMENT")&"' ORDER BY COD"
                Dados12.Open(SQL)
                conteudoEmbalagem = iif( Not Dados12.EOF, Trim(Dados12("DSC")), "" )
                Dados12.Close()

                linhaProducao = iif( Consulta1("FT1_LINAUTO") = "M", "Manual", "Automatica" )
%>
                      <td class="BordaEsquerda" ><strong>Conteúdo:</strong></td>
                      <td class="BordaDireita"><%=conteudoEmbalagem %></td>
                      <td><strong>Consumo Mensal:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_CONSMES") %></td>
                    </tr>
                    <tr>
                      <td class="BordaEsquerda" ><strong>Linha prod. Cliente:</strong></td>
                      <td><%=linhaProducao %></td>
                      <td class="BordaEsquerda" ><strong>Quantidade de Cores:</strong></td>
                      <td class="BordaDireita">?</td>
                    </tr>
                    <tr>
<%
                  SQL = "SELECT COD, DSC FROM SKTIP WHERE ATV = 1 AND DEL = 0 AND COD = '"&Consulta1("FT1_TIPOPRD")&"' ORDER BY COD"
                  Dados12.Open(SQL)
                  if (not Dados12.EOF) then
                    tipoProdutoPrimario = Trim(Dados12("DSC"))
                  else
                    tipoProdutoPrimario = ""
                  end if
                  Dados12.Close()
%>
                      <td class="BordaEsquerda"><strong>Tipo prod. Primário:</strong></td>
                      <td class="BordaDireita"><%=tipoProdutoPrimario %></td>
                      <td><strong>Suportação Primário:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_SUPORTEPRIMARIO")%></td>
                    </tr>
<%
                  if Consulta1("FT1_QTDPRI") <> "" or  Consulta1("FT1_PESOPRI") <> "" then
%>
                    <tr>
                      <td class="BordaEsquerda" ><strong>Quant. Primário:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_QTDPRI") %></td>
                      <td><strong>Peso Primário(kg):</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_PESOPRI") %></td>
                    </tr>
<%
                  end if 

                  if Consulta1("FT1_TEMPOEST") <> "" or  Consulta1("FT1_UNIREL") <> "" then
%>
                    <tr>
                      <td class="BordaEsquerda" ><strong>Estocagem(meses):</strong></td>
                      <td  class="BordaDireita"><%=Consulta1("FT1_TEMPOEST") %></td>
<% 
                    select case Consulta1("FT1_UNIREL")
                           case "0"  umidadeRelativa = "0 a 10"
                           case "10" umidadeRelativa = "11 a 20"
                           case "20" umidadeRelativa = "21 a 30"
                           case "30" umidadeRelativa = "31 a 40"
                           case "40" umidadeRelativa = "41 a 50"
                           case "50" umidadeRelativa = "51 a 60"
                           case "60" umidadeRelativa = "61 a 70"
                           case "70" umidadeRelativa = "71 a 80"   
                           case "80" umidadeRelativa = "81 a 90" 
                           case "90" umidadeRelativa = "91 a 100"    
                    end select                                      
%>
                      <td><strong>Umid. Relat(%):</strong></td>
                      <td class="BordaDireita"><%=umidadeRelativa %></td>
                    </tr>
<%
                  end if 
%>
                    <tr>
                      <td class="BordaEsquerda"><strong>MC (%):</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_Margem") %></td>
                      <td><strong>RO (%):</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_Resultado") %></td>
                    </tr>
                    <tr>
                      <td class="BordaEsquerda"><strong>Valor Estimado Clichê:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_VALORCLICHE") %></td>
                      <td><strong>Valor Estimado Matriz:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_VALORMATRIZ") %></td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div class="col-md-10">
                <table class="table table-condensed" style="margin-bottom:0px">
                  <thead>
                    <tr class="CabecalhoDestaque">
                      <th colspan="4" class="text-center" style="border-top:2px solid #ddd">Acabamento</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td class="BordaEsquerda"><strong>Fechamento</strong></td>
                      <td><%=Consulta1("FT1_B5AMARRAD") %></td> 
                      <td class="BordaEsquerda" ><strong>Quant. Pacote:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_B5QTDPCT") %></td>
                    </tr>

                    <tr>
                      <td class="BordaEsquerda" ><strong>Fitilho(LxC):</strong></td>
                      <td><%=Consulta1("FT1_B5FITLAR") %>x<%=Consulta1("FT1_B5FITCOM") %></td>
                      <td class="BordaEsquerda"><strong>Paletizado:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_B5PALETE") %></td>
                    </tr>

                    <tr>
                      <td class="BordaEsquerda" ><strong>Palete:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_B5TPPALE") %> - <%'=Consulta1("FT1_B5DESCPAL")%></td>
<%
                    usaRack = iif( Consulta1("FT1_RACK") = "S", "Sim", "Não" )
%>
                      <td class="BordaEsquerda" ><strong>Usa Rack:</strong></td>
                      <td class="BordaDireita"><%=usaRack %></td>
                    </tr>
                    <tr>
                  <%
                  overhang = iif( Consulta1("FT1_OVERHANG") = "S", "Sim", "Não" )
                  %>                                   
                      <td  class="BordaEsquerda"><strong>Overhang</strong></td>
                      <td class="BordaDireita"><%=overhang%></td>
                      <td class="BordaEsquerda" ><strong>Alt Max Palete(mm):</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_ALTPAL") %></td>
                    </tr>
                    <tr>
                      <td class="BordaEsquerda"><strong>Cantoneira:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_B5CANTON")%></td>
                      <td class="BordaEsquerda" ><strong>Face oposta:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_B5FACOP") %></td>
                    </tr>
                    <tr>
                      <td class="BordaEsquerda" ><strong>Cantinho:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_B5CANTON")%></td>
                      <td><strong>Tampa Mad:</strong></td>
                      <td class="BordaDireita"><%=Consulta1("FT1_B5TAMPA")%></td>
                    </tr>

                    <tr>
                      <td class="BordaInferior BordaEsquerda" ><strong>Espelho:</strong></td>
                      <td class="BordaInferior BordaDireita"><%=Consulta1("FT1_B5ESPELHO") %></td>
                      <td class="BordaInferior"><strong>Filme Strech:</strong></td>
                      <td class="BordaInferior BordaDireita"><%=Consulta1("FT1_B5FILME")%></td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          </td>
        </tr>
        <tr>
          <td colspan="2">
<%
        if Trim(Consulta1("FT1_OBSERV")) <> "" then
%>
            <table class="table table-condensed" style="margin-bottom:0px">
              <tbody>
<%
            if len(pedidosPendentes) > 0 then 
%>
                <tr class="alert-danger">
                  <th colspan="8" class="text-center" style="border-top:2px solid #ddd">Peddos em aberto</th>
                </tr>
                <tr>
                  <td colspan="8"><strong><%=pedidosPendentes %></strong></td>
                </tr>
<%
            end if 
%>
                <tr class="CabecalhoDestaque">
                  <th colspan="8" class="text-center" style="border-top:2px solid #ddd">Informações Adicionais (tem que prevalecer a da abertura da SAP quando aprovado para FT)</th>
                </tr>
                <tr>
                  <td colspan="8" class="BordaEsquerda BordaDireita" ><%=Consulta1("FT1_OBSERV") %></td>
                </tr>
              </tbody>
            </table>
<%
        end if 
%>
            <table class="table table-condensed" style="margin-bottom:0px">
              <thead>
                <tr class="CabecalhoDestaque">
                  <th colspan="9" class="text-center" style="border-top:2px solid #ddd">Observações</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td class="BordaEsquerda text-center" ><strong>Emissão SAP</strong></td>
                  <td class="text-center"><strong>Data da Emissão</strong></td>
                  <td class="text-center"><strong>Expira Em</strong></td>
                  <td class="text-center"><strong>Status</strong></td>
                  <td class="text-center"><strong>Aprovou SAP</strong></td>
                  <td class="text-center"><strong>Data da aprovação</strong></td>
                  <td class="text-center"><strong>Data de Corte Amostra</strong></td>
                  <td class="text-center"><strong>Data de Envio Amostra</strong></td>
                  <td class="text-center BordaDireita"><strong>Quem Enviou Amostra</strong></td>
                </tr>
                <tr>
                  <td class="BordaEsquerda BordaInferior text-center" ><%=Consulta1("FT1_USU") %></td>
                  <td class="text-center BordaInferior"><%=AcertaData2(Consulta1("FT1_DATA")) %></td>
                  <td class="text-center BordaInferior"><%if Trim(Consulta1("FT1_DATA")) <> "" then%> <%=Dateadd("d",90,AcertaData2(Consulta1("FT1_DATA")))%> <%end if %></td>
                  <td class="text-center BordaInferior"><%=Consulta1("FT1_STATUS") %></td>
                  <td class="text-center BordaInferior"><%=Consulta1("FT1_USU") %></td>
                  <td class="text-center BordaInferior"><%=Consulta1("FT1_DataGeracaoFT")%></td>
                  <td class="text-center BordaInferior"><%=Consulta1("FT1_DATACORTEAMOSTRA") %></td>
                  <td class="text-center BordaInferior"><%=AcertaData2(FormataData(Consulta1("FT1_DataEnvioAmostra"))) %></td>
                  <td class="text-center BordaDireita BordaInferior"><%=Consulta1("FT1_QuemEnviouAmostra") %></td>
                </tr>
              </tbody>
              <tfoot>
                <tr>
                  <td style="border:none !important;" colspan="7">&nbsp;</td>
                  <td style="border:none !important;" colspan="2">Data de Edição: 03/09/2013</td>
                </tr>
                <tr>
                  <td style="border:none !important;" colspan="7">&nbsp;</td>
                  <td style="border:none !important;" colspan="2">Data de Revisão: 27/08/2024</td>
                </tr>
                <tr>
                  <td style="border:none !important;" colspan="7">&nbsp;</td>
                  <td style="border:none !important;" colspan="2">RGQ305rev05</td>
                </tr>
              </tfoot>
            </table>
              </td>
            </tr>
          </table>
        </div>
      </div>
      <div>
        <div class="col-md-10 text-center" id="printable">
            
        </div>
      </div>
    </div>
  </body>
</html>
<%
Response.end
Response.Write Session("usr") & "-----"
For Each Key in Session.Contents
  Response.Write Key & "|" & Session.Contents(Key) & "<br>"
Next
Response.Write "----"

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
' -------------------------------------------------------------------------------------------------
'DATA FORMATO NORMAL
function AcertaData2(data)

  if Trim(data) <> "" then
    Hora = ""
    if len(data) = 8 then
      'FORMATO AAAAMMDD
      dia = mid(data,7,2)
      if len(dia) = 1 then
        dia = "0"&dia
      end if
      mes = mid(data,5,2)
      if len(mes) = 1 then
        mes = "0"&mes
      end if
      ano = mid(data,1,4)
    else
      if len(data) > 10 then
        'FORMATO AAAAMMDD HH:MM:SS
        dia = mid(data,9,2)
        if len(dia) = 1 then
          dia = "0"&dia
        end if
        mes = mid(data,6,2)
        if len(mes) = 1 then
          mes = "0"&mes
        end if
        ano = mid(data,1,4)
        Hora = mid(data,12,8)
      else
        'FORMATO AAAA/MM/DD
        dia = Trim(mid(data,9,2))
        if len(dia) = 1 then
          dia = "0"&dia
        end if
        mes = Trim(mid(data,6,2))
        if len(mes) = 1 then
          mes = "0"&mes
        end if
        ano = mid(data,1,4)
      end if
    end if
    AcertaData2 = dia&"/"&mes&"/"&ano&" "&hora
  else
    AcertaData2 = ""
  end if
end function
%>