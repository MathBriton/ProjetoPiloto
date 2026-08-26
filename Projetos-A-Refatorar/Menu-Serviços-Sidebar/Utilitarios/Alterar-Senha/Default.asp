<% @LANGUAGE="VBSCRIPT" TRANSACTION="REQUIRED" %>
<%
response.Buffer=true
response.AddHeader "cache-control", "private"
response.AddHeader "pragma", "no-cache"
response.ExpiresAbsolute = #January 1, 1990 00:00:01#
response.Expires = Now()-1
response.AddHeader "Cache-Control", "must-revalidate"
response.AddHeader "Cache-Control", "no-cache"
response.Addheader "Last-Modified:", Now()
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css" />
    <link rel="icon" type="image/png" sizes="32x32"   href="corporativo/images/favicon/favicon-32x32.png">
    <link rel="icon" type="image/png" sizes="16x16"   href="corporativo/images/favicon/favicon-16x16.png">
    <style>
        body {
            min-height: 100vh;
            background-color: #f5f5f5;
        }
        
        .logo-img {
            max-width: 180px;
            max-height: 52px;
            width: auto;
            height: auto;
            object-fit: contain;
        }

        .auth-nav {
            background-color: #f1f3f7;
            border-radius: 11px;
            padding: 4px;
            gap: 3px;
        }

        .auth-nav .nav-link {
            border-radius: 8px;
            font-size: .8rem;
            font-weight: 600;
            color: #6c757d;
            padding: .48rem .7rem;
            border: none;
            white-space: nowrap;
        }

        .auth-nav .nav-link.active {
            background: #fff;
            color: #0d6efd;
            box-shadow: 0 1px 5px rgba(0,0,0,.10), 0 0 0 1px #e2e8f0;
        }

        .auth-nav .nav-link:hover:not(.active) {
            background: rgba(255,255,255,.55);
            color: #495057;
        }

        .form-label {
            font-size: .8rem;
            font-weight: 600;
            color: #495057;
            margin-bottom: .35rem;
        }
    </style>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    
    <script type="text/javascript">
        $(document).ready(function () {
			verificarInatividade();
            anexarEventos();
        });

		/*
		 Função para recarregar pagina ao ficar inativo por mais de 1 minuto
		 Fazemos isso pois perdemos a referencia da variavel AUTH_USER do asp após um tempo e é necessario recarregar para atualizar 
		*/
		function verificarInatividade() {
			const TEMPO_LIMITE = 1 * 60 * 1000; 
			let timerInatividade;

			function resetarTimer() {
				clearTimeout(timerInatividade);

				timerInatividade = setTimeout(function () {
					location.reload(); 
				}, TEMPO_LIMITE);
			}

			const eventos = ["mousemove","keydown","mousedown","touchstart","scroll","click"];
			eventos.forEach(function (evento) {
				$(document).on(evento, resetarTimer);
			});

			resetarTimer();
		}

        function anexarEventos() {
            // Ao trocar de aba esconde alertas e limpa formulários
            $('#authTabs button').on('shown.bs.tab', function () {
                $('.alert[id$="-alert"]').addClass('d-none');

                $('form').each(function () { 
                    this.classList.remove('was-validated'); 
                });
            });

            $('#sistema-login').on('input', function () { 
                $(this).removeClass('is-invalid'); 
            });

            // Formulário Windows
            $('#form-windows').on('submit', function (e) {
                e.preventDefault();
                esconderAlerta('#win-alert');

                enviarFormulario('WINDOWS');
            });

            // Formulário Usuário e Senha
            $('#form-sistema').on('submit', function (e) {
                e.preventDefault();
                var $form = $(this);

                esconderAlerta('#up-alert');
                $form[0].classList.add('was-validated');
                if (!$form[0].checkValidity()) return;

                enviarFormulario('SISTEMA');
            });

            // Formulário AD
            $('#form-ad').on('submit', function (e) {
                e.preventDefault();
                var $form = $(this);

                esconderAlerta('#ad-alert');
                $form[0].classList.add('was-validated');
                if (!$form[0].checkValidity()) return;

                enviarFormulario('AD');
            });
        }
        
        function mostrarAlerta(sel, msg) { 
            $(sel).text(msg).removeClass('d-none'); 
        }

        function esconderAlerta(sel){ 
            $(sel).addClass('d-none'); 
        }

        function enviarFormulario(metodoAutenticacao) {   
            const $btn = $('.btn-enter');  
            var label = $btn.text().trim();  

            $btn.prop('disabled', true).html(
            	'<span class="spinner-border spinner-border-sm me-2"></span>Aguarde...'
            );
            
            let dadosFormulario = {
                metodoAutenticacao: metodoAutenticacao
            };

            switch (metodoAutenticacao) {
                case 'WINDOWS':
                    dadosFormulario.usuario = $('#windows-login').val();
                    break;
                case 'SISTEMA':
                    dadosFormulario.usuario = $('#sistema-login').val();
                    dadosFormulario.senha = $('#sistema-password').val();
                    break;
                case 'AD':
                    dadosFormulario.dominio = $('#ad-dominio').val();
                    dadosFormulario.usuario = $('#ad-login').val();
                    dadosFormulario.senha = $('#ad-password').val();
                    break;
                default:
                    Toast.error('Método de autenticação inválido.');
                    $btn.prop('disabled', false).text(label);
                    return;
            }

            $.ajax({
                url: 'Autenticacao.asp',
                type: 'POST',
                dataType: 'json',
                data: dadosFormulario,
                success: function(response) {
                    if (!response.sucesso) {
                        if (response.usuarioDesativado) {
                            var msgDesativado = response.mensagem || 'Seu acesso foi desativado. Entre em contato com o suporte.';
                            Toast.warning('Atenção: ' + msgDesativado);
                        } else {
                            Toast.error(response.mensagem);
                        }
                        $btn.prop('disabled', false).text('Entrar');
                        return;
                    }

                    var url = '';
                    if (response.flagTrocarSenha) {
                        Toast.warning(response.mensagem);
                        url = '/corporativo/usuario/TrocarSenha.asp'; 
                    } else {
                        Toast.success("Login realizado com sucesso!");
                        url = '/corporativo/frameset/principal.asp';
                        /* Sinaliza ao principal.asp que é um primeiro acesso
                           pós-login, para que a sidebar inicie expandida
                           independente do que estava salvo no localStorage. */
                        try { sessionStorage.setItem('swnet.fromLogin', '1'); } catch(e) {}
                    }

                    setTimeout(function () {
                        $btn.prop('disabled', false).text('Entrar');
                        window.location.href = url;
                    }, 2000);
                },
                error: function(xhr, status, error) {
                    Toast.error('Erro ao salvar formulário: ' + error);
                    $btn.prop('disabled', false).text('Entrar');
                },
            });

        }
    </script>
</head>
<body class="d-flex align-items-center justify-content-center min-vh-100 py-4">
    <div class="w-100 px-3" style="max-width:520px;">
        <div class="card">
            <div class="card-header bg-white border-bottom border-light text-center py-2">
                <img src="/corporativo/images/SWlogo500w.png" alt="Logo da empresa" class="logo-img" />
            </div>

            <div class="card-body px-4 pt-4 pb-3">
                <h5 class="fw-bold text-dark mb-1" style="font-size:1.2rem; letter-spacing:-.2px;">
                    Acesso ao sistema
                </h5>
                <p class="text-secondary mb-4" style="font-size:.85rem;">
                    Selecione o método de autenticação e faça login.
                </p>

                <!-- Abas de autenticação -->
                <ul class="nav auth-nav d-flex flex-nowrap mb-4" id="authTabs">
                    <li class="nav-item flex-fill">
                        <button class="nav-link active w-100 text-center" id="tab-windows" data-bs-toggle="pill" data-bs-target="#panel-windows" type="button">
                            <i class="bi bi-windows me-2"></i>Windows
                        </button>
                    </li>
                    <li class="nav-item flex-fill">
                        <button class="nav-link w-100 text-center" id="tab-userpass" data-bs-toggle="pill" data-bs-target="#panel-userpass" type="button">
                            <i class="bi bi-person me-2"></i>Usuário / Senha
                        </button>
                    </li>
                    <li class="nav-item flex-fill">
                        <button class="nav-link w-100 text-center" id="tab-ad"data-bs-toggle="pill" data-bs-target="#panel-ad" type="button">
                            <i class="bi bi-shield me-2"></i>Active Directory
                        </button>
                    </li>
                </ul>

                <div class="tab-content" id="authTabContent">
                    <!-- Windows -->
                    <div class="tab-pane fade show active" id="panel-windows">
                        <form id="form-windows" novalidate>
                            <div class="mb-4">
                                <label class="form-label" for="windows-login">Usuário Windows</label>
                                <div class="input-group">
                                    <span class="input-group-text py-0 px-2">
                                        <i class="bi bi-person-fill text-secondary"></i>
                                    </span>
                                    <input disabled type="text" id="windows-login" class="form-control form-control-sm" value="<%=request.ServerVariables("AUTH_USER")%>" />
                                </div>
                                <div class="form-text mt-1" style="font-size:.75rem;">
                                    <i class="bi bi-info-circle me-1"></i>Detectado automaticamente pela sessão Windows ativa.
                                </div>
                            </div>

                            <div id="win-alert" class="alert alert-danger border-0 rounded-3 py-2 px-3 d-none mb-3"></div>

                            <button type="submit" class="btn btn-primary btn-enter w-100">
                                <i class="bi bi-box-arrow-in-right me-2"></i>Entrar com Windows
                            </button>
                        </form>
                    </div>

                    <!-- Usuário e Senha -->
                    <div class="tab-pane fade" id="panel-userpass">
                        <form id="form-sistema" novalidate>
                            <div class="mb-3">
                                <label class="form-label" for="sistema-login">Usuário Sistema</label>
                                <div class="input-group">
                                    <span class="input-group-text py-0 px-2">
                                        <i class="bi bi-person-fill text-secondary"></i>
                                    </span>
                                    <input type="text" id="sistema-login" class="form-control form-control-sm" placeholder="Usuario" autocomplete="username" required />
                                    <div class="invalid-feedback" style="font-size:.76rem;">Informe seu login.</div>
                                </div>
                            </div>

                            <div class="mb-1">
                                <label class="form-label mb-0" for="sistema-password">Senha</label>
                                <input type="password" id="sistema-password" class="form-control form-control-sm" placeholder="••••••••" autocomplete="current-password" required />
                                <div class="invalid-feedback" style="font-size:.76rem;">Informe sua senha.</div>
                            </div>

                            <div class="d-flex align-items-center justify-content-end mb-1">
                                <a href="/corporativo/usuario/GerarSenhaProvisoria.asp" id="forgot-link"
                                    class="text-primary text-decoration-none fw-semibold"
                                    style="font-size:.78rem;">
                                    <i class="bi bi-question-circle me-1"></i>Esqueceu a senha?
                                </a>
                            </div>

                            <div id="up-alert" class="alert alert-danger border-0 rounded-3 py-2 px-3 d-none mt-3"></div>

                            <button type="submit" class="btn btn-primary btn-enter w-100 mt-4">
                                <i class="bi bi-box-arrow-in-right me-2"></i>Entrar
                            </button>
                        </form>
                    </div>

                    <!-- AD -->
                    <div class="tab-pane fade" id="panel-ad">
                        <form id="form-ad" novalidate>
                            <div class="mb-2">
                                <label for="ad-dominio" class="form-label">Domínio</label>
                                <select required class="form-select form-select-sm select2" id="ad-dominio" name="ad-dominio" data-placeholder="Todas as filiais">
                                    <option selected value="GROUP">GROUP</option>
                                    <option value="WESTROCK">WESTROCK</option>
                                </select>
                            </div>
                                
                            <div class="mb-3">
                                <label class="form-label" for="ad-login">Usuário AD</label>
                                <div class="input-group">
                                    <span class="input-group-text py-0 px-2">
                                        <i class="bi bi-person-fill text-secondary"></i>
                                    </span>
                                    <input type="text" id="ad-login" class="form-control form-control-sm" placeholder="Usuario" autocomplete="username" required />
                                    <div class="invalid-feedback" style="font-size:.76rem;">Informe seu login.</div>
                                </div>
                            </div>

                            <div class="mb-1">
                                <div class="d-flex align-items-center justify-content-between mb-1">
                                    <label class="form-label mb-0" for="ad-password">Senha</label>
                                </div>
                                <input type="password" id="ad-password" class="form-control form-control-sm" placeholder="••••••••" autocomplete="current-password" required />
                                <div class="invalid-feedback" style="font-size:.76rem;">Informe sua senha.</div>
                            </div>

                            <div id="up-alert" class="alert alert-danger border-0 rounded-3 py-2 px-3 d-none mt-3"></div>

                            <button type="submit" class="btn btn-primary btn-enter w-100 mt-4">
                                <i class="bi bi-box-arrow-in-right me-2"></i>Entrar
                            </button>
                        </form>
                    </div>

                </div>
            </div>
        </div>
    </div>
</body>
</html>