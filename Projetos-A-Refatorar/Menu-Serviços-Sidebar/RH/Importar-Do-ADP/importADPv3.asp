<!DOCTYPE html>
<html lang="pt-br">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <!--BOOTSTRAP-->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.7.0/css/all.css" integrity="sha384-lZN37f5QGtY3VHgisS14W3ExzMWZxybE1SJSEsQp9S+oqd12jhcu+A56Ebc1zFSJ" crossorigin="anonymous" />
    <title>SWnet - Importa Arquivos do ADP</title>
  </head>

  <body>

<%
select case trim(request("planta"))
case "00" : subTitulo = "Pirapetinga"
case "04" : subTitulo = "Uberaba"
case "05" : subTitulo = "Bento Gonçalves"
end select
%>
    <div class="text-center">
      <br>
      <h3 style="color: gray">Importa Arquivos</h3>
      <br>
      <h4 style="color: gray">Funcionários e Acessos de Funcionários</h4>
      <br>
      <h4 style="color: gray"><%=subTitulo%></h4>
    </div>

    <div class="container"> 

      <div class="row">     
        <div class="col">
        
          <label for="funcionariosFile" class="form-label">Selecione o arquivo Funcionarios.csv:</label>
          <input type="file" id="funcionariosFile" accept=".csv" class="form-control"/>

        </div>
      </div>

      <br>

      <div class="row">     
        <div class="col">

          <label for="acessosFile" class="form-label">Selecione o arquivo AcessosFuncionarios.csv:</label>
          <input type="file" id="acessosFile" accept=".csv" class="form-control"/>

        </div>
      </div>

      <br>
      <br>

      <div class="row">     
        <div class="col">

          <button id="btnEnviar" class="btn btn-primary" disabled>Enviar para Importação</button>

        </div>
      </div>

      <br>
      <br>

      <div class="row">     
        <div class="col">

          <div id="resultado"></div>

        </div>
      </div>

      <script type="text/javascript">

        const funcionariosInput = document.getElementById('funcionariosFile');
        const acessosInput = document.getElementById('acessosFile');
        const btnEnviar = document.getElementById('btnEnviar');
        const resultadoDiv = document.getElementById('resultado');

        let funcionariosContent = "";
        let acessosContent = "";
        // ----------------------------------------------------------------------------------------
        function validaArquivos() {
          if (!funcionariosInput.files.length || !acessosInput.files.length) {
            btnEnviar.disabled = true;
            return;
          }
          if (!funcionariosInput.files[0].name.toLowerCase().endsWith("funcionarios.csv") ||
              !acessosInput.files[0].name.toLowerCase().endsWith("acessosfuncionarios.csv")) {
            alert('Os arquivos devem ser "Funcionarios.csv" e "AcessosFuncionarios.csv" respectivamente.');
            btnEnviar.disabled = true;
            return;
          }
          btnEnviar.disabled = false;
        }
        // ----------------------------------------------------------------------------------------
        funcionariosInput.addEventListener('change', () => {
          validaArquivos();
          funcionariosContent = "";
          resultadoDiv.textContent = '';
        });
        // ----------------------------------------------------------------------------------------
        acessosInput.addEventListener('change', () => {
          validaArquivos();
          acessosContent = "";
          resultadoDiv.textContent = '';
        });
        // ----------------------------------------------------------------------------------------
        async function lerArquivo(inputFile) {
          return new Promise((resolve, reject) => {
            const file = inputFile.files[0];
            if (!file) resolve("");
            const reader = new FileReader();
            reader.onload = e => resolve(e.target.result);
            reader.onerror = e => reject(e.target.error);
            reader.readAsText(file, "UTF-8");
          });
        }
        // ----------------------------------------------------------------------------------------
        btnEnviar.addEventListener('click', async () => {
          btnEnviar.disabled = true;
          resultadoDiv.textContent = "Lendo arquivos...";

          DataEhora = formatarDataHora();

          try {
            funcionariosContent = await lerArquivo(funcionariosInput);
            acessosContent = await lerArquivo(acessosInput);

            if (!funcionariosContent || !acessosContent) {
              resultadoDiv.textContent = "Erro ao ler um ou ambos os arquivos.";
              btnEnviar.disabled = false;
              return;
            }

            resultadoDiv.textContent = "Enviando dados dos Funcionários...";

            // enviarDados - funcionarios ---------------------------------------------------------
            let respostaFunc = await enviarDados( 'funcionarios', funcionariosContent, DataEhora, adicionarZerosAEsquerda( <%=trim(request("planta"))%>, 2 ) );
            resultadoDiv.innerHTML = "Funcionários importados:<br>" + respostaFunc + "<br><br>Enviando dados dos Acessos...";

            // enviarDados - acessos --------------------------------------------------------------
            let respostaAcessos = await enviarDados( 'acessos', acessosContent, DataEhora, adicionarZerosAEsquerda( <%=trim(request("planta"))%>, 2 ) );
            resultadoDiv.innerHTML = `Funcionários:<br>${respostaFunc}<br><br>Acessos de Funcionários:<br>${respostaAcessos}`;

            // execIntegraAdpMadis ----------------------------------------------------------------
            let respostaIntegra = await execIntegraAdpMadis( adicionarZerosAEsquerda( <%=trim(request("planta"))%>, 2 )  );
            resultadoDiv.innerHTML = `Funcionários:<br>${respostaFunc}<br><br>Acessos de Funcionários:<br>${respostaAcessos}<br><br>Integração ADP Madis Bento<br>${respostaIntegra}`;

            // let respostaJornadas = await montaJornadas('jornadas');
            // resultadoDiv.textContent = `Funcionários:\n${respostaFunc}\n\nAcessos de Funcionários:\n${respostaAcessos}\n\nJornadas de Funcionários:\n${respostaJornadas}`;

          } catch (err) {
            resultadoDiv.textContent = "Erro: " + err.message;
          }

          btnEnviar.disabled = false;
        });
        // ----------------------------------------------------------------------------------------
        async function enviarDados(tipo, conteudoCSV, aDataEhora, aPlanta) {

          const corpo = 'tipo=' + encodeURIComponent(tipo) + 
                        '&csvdata=' + encodeURIComponent(conteudoCSV) +
                        '&dataehora=' + encodeURIComponent(aDataEhora) + 
                        '&planta=' + encodeURIComponent(aPlanta);

          const resposta = await fetch('processa_matriculas.asp', {
                                    method: 'POST',
                                    headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
                                    body: corpo
                                  });

          if(!resposta.ok) {
            throw new Error('Erro servidor: ' + resposta.status);
          }

          return await resposta.text();

        }
        // ----------------------------------------------------------------------------------------
        async function montaJornadas(tipo) {

          const corpo = 'tipo=' + encodeURIComponent(tipo);
          const resposta = await fetch('processa_jornadas.asp', {
                                    method: 'POST',
                                    headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
                                    body: corpo
                                  });

          if(!resposta.ok) {
            throw new Error('Erro servidor: ' + resposta.status);
          }

          return await resposta.text();

        }
        // ----------------------------------------------------------------------------------------
        async function execIntegraAdpMadis(aPlanta) {

          const corpo = 'planta=' + encodeURIComponent(aPlanta);
          const resposta = await fetch('processa_IntegraAdpMadis.asp', {
                                    method: 'POST',
                                    headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
                                    body: corpo
                                  });

          if(!resposta.ok) {
            throw new Error('Erro servidor: ' + resposta.status);
          }

          return await resposta.text();

        }
        // ----------------------------------------------------------------------------------------
        function formatarDataHora() {

          const dataHoraAtual = new Date();

          const ano       = dataHoraAtual.getFullYear();
          const mes       = String(dataHoraAtual.getMonth() + 1).padStart(2, '0'); // Meses começam em 0
          const dia       = String(dataHoraAtual.getDate()).padStart(2, '0');
          const horas     = String(dataHoraAtual.getHours()).padStart(2, '0');
          const minutos   = String(dataHoraAtual.getMinutes()).padStart(2, '0');
          const segundos  = String(dataHoraAtual.getSeconds()).padStart(2, '0');
          //const milissegundos = String(dataHoraAtual.getMilliseconds()).padStart(3, '0');

          return `${ano}-${mes}-${dia} ${horas}:${minutos}:${segundos}`;

        }
        // ----------------------------------------------------------------------------------------
        function adicionarZerosAEsquerda(numero, totalDeDigitos) {
          return String(numero).padStart(totalDeDigitos, '0');
        }

      </script>

    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>  

  </body>
</html>
