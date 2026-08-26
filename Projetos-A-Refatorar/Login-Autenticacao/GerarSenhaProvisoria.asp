<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta charset="utf-8" />
    <title>SWnet - Recuperação de senha</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />

    <script type="text/javascript">
      function Envia() {
        form.submit();
      }
    </script>

    <style>
      .MargemAbaixo {
        margin-bottom:30px;
      }
      .MargemAbaixoMetade {
        margin-bottom:15px;
      }
    </style>

  </head>
  <body class="body">
    <div class="container-fluid body-content margem">
      <h3><strong>SWnet - Recupere sua senha.</strong></h3>
      <hr/>
      <form name="form" action="ConfirmaGeracao.asp" class="form-horizontal" method="post" role="form">
        <div class="jumbotron col-sm-offset-3 col-sm-6">
          <div class="row">
            <h4>Digite seu e-mail.</h4>
            <hr/>
          </div>
          <div class="row MargemAbaixoMetade">
            <div class="form-group-sm">
              <div class="col-sm-10">
                <div class="input-group">
                  <span class="input-group-addon"><i class="glyphicon glyphicon-envelope"></i></span>
                  <input type="email" name="email" id="email" value="" class="form-control" placeholder="E-mail" required="required"/>
                </div>
              </div>  
            </div>
          </div>
          <div class="row MargemAbaixo">
            <div class="col-sm-10">
              <button class="btn btn-success btn-sm " onclick="Envia()">
                <span class="glyphicon glyphicon-thumbs-up"></span> Enviar
              </button>
            </div>
          </div>
        </div>
      </form>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
  </body>
</html>