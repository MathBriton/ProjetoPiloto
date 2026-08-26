<!DOCTYPE html>
<html lang="pt-br" xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <!--BOOTSTRAP-->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js"></script>
    <script src="/corporativo/LotePiloto/jsLotePiloto/FuncoesDoLotePiloto.js"></script>
    <title></title>
    <!-- #include file="../cnn/cnn.asp" -->
    <!-- #include file="../includes/variaveis.asp" -->
    <!-- #include file="../funcoes/funcoes.asp"-->
  </head>
  <body>
    <div class="container container-fluid">
        <input type="hidden" name="Planta" id="Planta" />
        <table class="table-hover" style="margin-top:6em; width: 24%; margin-left: 38%; margin-right: 38%; font-size: 16px; background-color: rgb(235, 235, 235); box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.2), 0 6px 20px 0 rgba(0, 0, 0, 0.19); letter-spacing: 2px;" border="0" cellpadding="20">
          <thead>
            <tr style="background-color: rgb(70, 70, 100); color: whitesmoke;">
              <th style="text-align: center;"><b>Selecione a Planta:</b></th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><label><input type="radio" name="selecione" value="00" onclick="VaiPara('00')"/>&nbsp;Pirapetinga</label></td>
            </tr>
            <tr>
              <td><label><input type="radio" name="selecione" value="04" onclick="VaiPara('04')"/>&nbsp;Uberaba</label></td>
            </tr>
            <tr>
              <td><label><input type="radio" name="selecione" value="05" onclick="VaiPara('05')"/>&nbsp;Bento Gonçalves</label></td>
            </tr>
            <tr>
              <td><label><input type="radio" name="selecione" value="06" onclick="VaiPara('06')"/>&nbsp;Maranguape</label></td>
            </tr>
            <tr>
              <td><label><input type="radio" name="selecione" value="07" onclick="VaiPara('07')"/>&nbsp;Saquarema</label></td>
            </tr>
            <tr>
              <td><label><input type="radio" name="selecione" value="08" onclick="VaiPara('08')"/>&nbsp;Araçatuba</label></td>
            </tr>
                        <tr>
              <td><label><input type="radio" name="selecione" value="09" onclick="VaiPara('09')"/>&nbsp;Blumenau</label></td>
            </tr>
            <tr>
              <td><label><input type="radio" name="selecione" value="10" onclick="VaiPara('10')"/>&nbsp;Pacajus</label></td>
            </tr>
            <tr>
              <td><label><input type="radio" name="selecione" value="11" onclick="VaiPara('11')"/>&nbsp;Porto Feliz</label></td>
            </tr>
            <tr>
              <td><label><input type="radio" name="selecione" value="12" onclick="VaiPara('12')"/>&nbsp;Tres Barras</label></td>
            </tr>
            <tr>
              <td><label><input type="radio" name="selecione" value="13" onclick="VaiPara('13')"/>&nbsp;Florestal</label></td>
            </tr>
          </tbody>
        </table>
    </div> 
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>  
  </body>
</html>