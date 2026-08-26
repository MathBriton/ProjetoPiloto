fetch("/corporativo/CriacaoUsuario/data/header.html")//SEMPRE LEMBRAR DE POR ESSE CAMINHO AQUI POIS SE TRATA DENTRO DO SERVIDOR SEM ISSO N VAI CARREGAR... fetch("/corporativo/CriacaoUsuario/data/header.html")

  .then(res => res.text())
  .then(html => {
    document.getElementById("header-container").innerHTML = html;
  })
  .catch(err => console.error("Erro ao carregar header:", err));