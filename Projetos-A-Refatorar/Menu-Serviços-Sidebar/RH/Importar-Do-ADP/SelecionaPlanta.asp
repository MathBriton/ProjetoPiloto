<!DOCTYPE html>
<!--#Include File="../../funcoes/funcoes.asp"-->
<!--#include file="../../includes/variaveis.asp" -->

<html lang="pt-br" xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <!--BOOTSTRAP-->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />
    <title></title>
    
    <style type="text/css">
      .objeto-centralizado{
        margin-left:auto !important;
        margin-right:auto !important;
        font-size: 80%;
        color:whitesmoke;
        border-color:#d6d8db;
      }
    </style>
  </head>
  <body>

  <div class="text-center">
    <br>
    <h3 style="color: gray">Importa Arquivos</h3>
    <br>
    <p style="color: gray">Funcionários e Acessos de Funcionários</p>
  </div> 

  <div class="container"> 
    
    <form name="form" method="post">
      <input type="hidden" name="Planta" id="Planta" />
      
      <table class="table-hover" style="margin-top:6em; min-width: 30%; margin-left: 35%; margin-right: 35%; font-size: 16px; background-color: rgb(235, 235, 235); box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.2), 0 6px 20px 0 rgba(0, 0, 0, 0.19); letter-spacing: 2px;" border="0" cellpadding="20">
        <thead>
            <tr style="background-color: rgb(70, 70, 100); color: whitesmoke;">
              <th style="text-align: center;"><b>Selecione a Planta</b></th>
            </tr>
        </thead>
        <tbody>
          <tr>
            <td>
<%
if UNIDADE_FILIAL = "00" then
%>            
              <label><input type="radio" name="selecione" value="00" onclick="Planta.value = this.value; location.href= 'importADPv3.asp?planta='+Planta.value;" />&nbsp;Pirapetinga</label>
<%
end if

if UNIDADE_FILIAL = "04" then
%>              
              <label><input type="radio" name="selecione" value="04" onclick="Planta.value = this.value; location.href= 'importADPv3.asp?planta='+Planta.value;" />&nbsp;Uberaba</label>
<%
end if

if UNIDADE_FILIAL = "05" then
%>              
              <label><input type="radio" name="selecione" value="05" onclick="Planta.value = this.value; location.href= 'importADPv3.asp?planta='+Planta.value;" />&nbsp;Bento Gonçalves</label>
<%
end if
%>              

<!--              
              <br>
              <br>
              <label><input type="radio" name="selecione" value="04" onclick="Planta.value = this.value; location.href= 'ImportaADP.asp?Planta='+Planta.value;" />&nbsp;Uberaba</label>
-->              
<!--              
              <br>
              <br>
              <label><input type="radio" name="selecione" value="06" onclick="Planta.value = this.value; location.href= 'ImportaADP.asp?Planta='+Planta.value;" />&nbsp;Maranguape</label>
              <br>
              <br>
              <label><input type="radio" name="selecione" value="07" onclick="Planta.value = this.value; location.href= 'ImportaADP.asp?Planta='+Planta.value;" />&nbsp;Saquarema</label>
-->              
            </td>
          </tr>
        </tbody>
      </table>
    </form>       
  </div> 
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>  
</body>
</html>