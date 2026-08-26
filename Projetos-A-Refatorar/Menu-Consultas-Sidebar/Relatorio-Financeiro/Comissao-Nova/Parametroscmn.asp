<!DOCTYPE html>

<html lang="pt-br" xmlns="http://www.w3.org/1999/xhtml">

<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Relatório de comissões</title>
  <link rel="stylesheet" type="text/css" href="../../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.css" />
  <link rel="stylesheet" type="text/css"
    href="../../../dhtmlx/dhtmlxCalendar/codebase/skins/dhtmlxcalendar_dhx_web.css" />
  <script type="text/javascript" src="../../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.js"></script>
  <script type="text/javascript" src="../../corporativo/funcoes/funcoesJAVASCRIPT.js"></script>
  <!--#Include File="../../corporativo/cnn/cnn.asp"-->
  <!--#Include File="../../corporativo/funcoes/funcoes.asp"-->
  <!--#Include File="../../corporativo/includes/variaveis.asp"-->
  <!--BOOTSTRAP-->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet"
    integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
  <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css"
    integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />
   
</script>
  <script type="text/javascript">
    var myCalendar;
    function doOnLoad() {
      myCalendar = new dhtmlXCalendarObject(["DataIni", "DataFim"]);
      myCalendar.setDate(myCalendar.getDate());
      myCalendar.setDateFormat("%d.%M.%Y");
      myCalendar.loadUserLanguage("br");
    }
    // ------------------------------------------------------------------------------------------
    function mascara(o, f) {
      document.getElementById("DataIni").value = "04/2023"
      v_obj = o;
      v_fun = f;
      setTimeout("execmascara()", 1);
    }
    // ------------------------------------------------------------------------------------------
    function execmascara() {
      v_obj.value = v_fun(v_obj.value);
    }
    function mascaraDataIni() {
      DataFormatada = document.getElementById("DataIni").value;
      if (DataFormatada.length == 10) {
        document.getElementById("DataIni").value = DataFormatada.substring(3, 10);
      }
      else if (DataFormatada.length == 7) {
        document.getElementById("DataIni").value = DataFormatada;
      }
      else {
        alert("Parametro Invalido");
      }
    }
    function mascaraDataFim() {
      DataFormatada = document.getElementById("DataFim").value;
      if (DataFormatada.length == 10) {
        document.getElementById("DataFim").value = document.getElementById("DataFim").value.substring(3, 10);
      }
      else if (DataFormatada.length == 7) {
        document.getElementById("DataFim").value = DataFormatada;
      }
      else {
        alert("Parametro Invalido");
      }
    }
    function DefineParametrosPlanta() {
      let Plantas = "";
      
      if (document.getElementById('BPRC').checked == false)
        Plantas = Plantas.replace(document.getElementById('BPRC').value + ',', '');
      else
        Plantas += document.getElementById('BPRC').value + ',';

      if (document.getElementById('BPCE').checked == false)
        Plantas = Plantas.replace(document.getElementById('BPCE').value + ',', '');
      else
        Plantas += document.getElementById('BPCE').value + ',';  

      if (document.getElementById('BPIC').checked == false)
        Plantas = Plantas.replace(document.getElementById('BPIC').value + ',', '');
      else
        Plantas += document.getElementById('BPIC').value + ','; 

      if (document.getElementById('BSQC').checked == false)
        Plantas = Plantas.replace(document.getElementById('BSQC').value + ',', '');
      else
        Plantas += document.getElementById('BSQC').value + ','; 
        
      if (document.getElementById('BUBC').checked == false)
        Plantas = Plantas.replace(document.getElementById('BUBC').value + ',', '');
      else
        Plantas += document.getElementById('BUBC').value + ',';    

      document.getElementById('hidParametros').value=Plantas;
    }
  </script>
  <style type="text/css">
    .fundo {
      background-color: lightgray;
      width: 100% !important;
      padding: 15px;
    }
  </style>
</head>
<% 'HELPDESK
' call ValidaPermissaoRelatorios("../../intranet/comissao/parametros",SESSION("CODUSER"),SESSION("FILIAL")) %>

  <body onload="doOnLoad()">
   
    <div class="container-fluid table-responsive" style="margin-top: 20px;">
      <form action="relatoriocmn.asp" method="post" class="form-group">
        <input type="hidden" id="hidParametros" name="hidParametros">
        <div class="row">
          <div class="col-4 offset-4">
            <table class="table fundo">
              <tr>
                <td colspan="2" class="text-center">
                  <h4><strong>Comissões</strong></h4>
                </td>
              </tr>
              <tr>
                <td>Filial</td>
                <td>
                  <input type="radio" id="BPRC" name="Planta" value="BPRC|BPRM" onclick="DefineParametrosPlanta()">
                  <label for="BPRC">BENTO</label><br>
                 
                  <input type="radio" id="BPCE" name="Planta" value="BPCE" onclick="DefineParametrosPlanta()">
                  <label for="BPCE">FORTALEZA</label><br>
                  
                  <input type="radio" id="BPIC" name="Planta" value="BPIC|BPIM|BPIW" onclick="DefineParametrosPlanta()">
                  <label for="BPIC">PIRAPETINGA</label><br>

                  <input type="radio" id="BSQC" name="Planta" value="BSQC" onclick="DefineParametrosPlanta()">
                  <label for="BSQC">SAQUAREMA</label><br>

                  <input type="radio" id="BUBC" name="Planta" value="BUBC|BUBM" onclick="DefineParametrosPlanta()">                  
                  <label for="BUBC">UBERABA</label><br>

                </td>
              </tr>
              <tr>
                <td>Emissão De</td>
                <td><input name="DataIni" id="DataIni" type="text" onblur="mascaraDataIni()" maxlength="10"
                    class="form-control" required="required" /> </td>
              </tr>
              <tr>
                <td>Emissão Até</td>
                <td><input name="DataFim" id="DataFim" type="text" onblur="mascaraDataFim()" maxlength="10"
                    class="form-control" required="required" /></td>
              </tr>
              <tr>
                <td>Cliente</td>
                <td><input name="ClieIni" type="text" maxlength="6" class="form-control" /></td>
              </tr>

              <%If Nivel=0 Then%>
                <tr>
                  <td>Vendedor</td>
                  <td><input name="vendedor" type="text" maxlength="6" class="form-control" required="required" /></td>
                </tr>
                <%else %>
                  <tr>
                    <td>Assinatura</td>
                    <td><input name="Assinatura" type="password" maxlength="20" class="form-control"
                        required="required" /></td>
                  </tr>
                  <%End If%>
                    <tr>
                      <td colspan="2" class="text-center">
                        <input type="submit" name="Submit" value="Gerar relatório" class="btn btn-primary" />
                      </td>
                    </tr>
            </table>
          </div>
        </div>
      </form>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js"
      integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM"
      crossorigin="anonymous"></script>
  </body>

</html>