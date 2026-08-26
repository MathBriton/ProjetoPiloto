<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Sistema de Chamados — Detalhes</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet" />
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
    <link href="https://cdn.jsdelivr.net/npm/quill@2.0.2/dist/quill.snow.css" rel="stylesheet" />
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css" rel="stylesheet">
    <link rel="stylesheet" href="../Compartilhado/estilos.css?v=<%=Now()%>">

    <style>
        :root {
            --paper: #f5f1e8;
            --paper-light: #f5f5f5;
            --ink: #1a1612;
            --ink-soft: #4a453e;
            --ink-mute: #8b857c;
            --line: rgba(26, 22, 18, 0.1);
            --line-soft: rgba(26, 22, 18, 0.06);
            --accent: #b64628;
            --amber: #c49a2f;
            --amber-soft: #f4ecdc;
            --amber-paper: #fbf1d7;
            --rose: #b53a37;
            --rose-soft: #f4dcdb;
            --sage: #5a7a4e;
            --blue-soft: #e0e1f9;
            --blue: #5a43d6;
            --sage-deep: #3a5530;
            --sage-soft: #e2eadc;
        }
        body {
            background-color: #f5f5f5;
            color: var(--ink);
        }
        .form-label {
            font-weight: 600;
            font-size: .8rem;
        }

        .btnVoltar {
            color: #000 !important;
            background-color: #edebeb !important;
            border-color: #c6c7c8 !important;
        }
        
        a { text-decoration: none; color: inherit; }
        button { background: transparent; border: 0; padding: 0; color: inherit; font: inherit; cursor: pointer; }
        .font-mono { font-family: monospace; }
        .soft-shadow { box-shadow: 0 1px 0 rgba(26,22,18,0.04), 0 8px 24px -12px rgba(26,22,18,0.10); }

        /* Tipografia compacta */
        .text-10 { font-size: 10px; } .text-11 { font-size: 11px; } .text-12 { font-size: 12px; }
        .text-13 { font-size: 13px; } .text-14 { font-size: 14px; }
        .text-28 { font-size: 28px; line-height: 1.15; }
        .text-34 { font-size: 34px; line-height: 1.1; }
        .text-mute { color: var(--ink-mute); }
        .text-soft { color: var(--ink-soft); }
        .text-accent { color: var(--accent); }
        .text-sage { color: var(--sage-deep); }
        .uppercase-tight { text-transform: uppercase; letter-spacing: 0.18em; }
        .border-line { border-color: var(--line) !important; }
        .hover-bg-paper:hover { background: var(--paper); }
        .hover-ink:hover { color: var(--ink) !important; }
        .hover-accent:hover { color: var(--accent) !important; }
        .hover-underline:hover { text-decoration: underline; }
        .cursor-pointer { cursor: pointer; }
        .min-w-0 { min-width: 0; }
        .flex-1 { flex: 1; }

        /* Pills */
        .pill { display: inline-flex; align-items: center; gap: 6px; border-radius: 9999px; padding: 4px 10px; font-size: 11px; font-weight: 500; }
        .pill .dot { width: 6px; height: 6px; border-radius: 9999px; }
        .pill-blue { background: var(--blue-soft); color: var(--blue); } .pill-blue .dot { background: var(--blue); }
        .pill-amber { background: var(--amber-soft); color: #7a5a12; } .pill-amber .dot { background: var(--amber); }
        .pill-rose  { background: var(--rose-soft);  color: #8e2825; } .pill-rose  .dot { background: var(--rose); }
        .pill-sage  { background: var(--sage-soft);  color: var(--sage-deep); } .pill-sage  .dot { background: var(--sage); }
        .pill-slate { background: #dde4ed; color: #3e4a63; } .pill-slate .dot { background: #6b7aa1; }
        .pill-stone { background: #e8e2d6; color: #5a4f3a; } .pill-stone .dot { background: #8a7a5e; }
        .pill-rose-strong { background: var(--rose); color: #fff; } .pill-rose-strong .dot { background: #fff; }
        .pill-neutral { background: #eee8da; color: var(--ink-soft); } .pill-neutral .dot { background: var(--ink-mute); }
        .pill-cream { background: var(--paper); color: var(--ink); box-shadow: inset 0 0 0 1px var(--line); }
        .pill-sm { padding: 2px 8px; }

        /* Avatar */
        .avatar { display: inline-flex; align-items: center; justify-content: center; border-radius: 9999px; font-family: monospace; font-weight: 500; flex-shrink: 0; }
        .av-ink    { background: var(--ink); color: var(--paper); }
        .av-sand { background: var(--accent); color: var(--paper); }
        .av-sage   { background: var(--sage-deep); color: var(--sage-soft); }
        .av-sand   { background: #e8ddc4; color: #4a3d1e; }
        .av-18 { width: 18px; height: 18px; font-size: 8px; }
        .av-34 { width: 34px; height: 34px; font-size: 12px; }
        .av-38 { width: 38px; height: 38px; font-size: 13px; }
        .av-42 { width: 42px; height: 42px; font-size: 15px; }
        .av-44 { width: 44px; height: 44px; font-size: 16px; }

        .section-label { font-size: 14px;  margin: 0; }

        /* Card */
        .card-custom { border-radius: 12px; border: 1px solid var(--line); background: #fff; padding: 1.25rem; }

        /* Mensagens — visual unificado */
        .msg { border-radius: 12px; border: 1px solid rgba(26,22,18,0.08); padding: 1.25rem; background: #fff; }
        .msg-body { font-size: 14px; line-height: 1.65; color: var(--ink); padding-left: 50px; }
        .msg-body p + p { margin-top: 0.5rem; }
        .msg-body ol, .msg-body ul { padding-left: 1.25rem; margin: 0.5rem 0; }
        .msg-body li + li { margin-top: 0.25rem; }
        .msg .more-btn { opacity: 0; transition: opacity .15s; }
        .msg:hover .more-btn { opacity: 1; }

        /* Nota interna */
        .note { border-radius: 12px; background: var(--amber-paper); border: 1px solid rgba(196,154,47,0.25); padding: 1.25rem; }

        .ticket-origin {
            border-radius: 12px;
            border: 1px solid rgba(26,22,18,0.08);
            background: #fff;
            padding: 1.25rem;
        }
        .ticket-origin-head { display: flex; align-items: flex-start; gap: 12px; margin-bottom: 12px; }
        .ticket-origin-head .origin-name { font-weight: 500; font-size: 14px; }
        .ticket-origin-head .origin-when { font-size: 11px; color: var(--ink-mute); font-family: monospace; }
        .ticket-origin-body { padding-left: 50px; }
        .ticket-origin-body .desc-text { font-size: 14px; line-height: 1.65; color: var(--ink); }
        .ticket-origin-body .desc-text p + p { margin-top: 0.5rem; }
        .ticket-origin-body .desc-text ol, .ticket-origin-body .desc-text ul { padding-left: 1.25rem; margin: 0.5rem 0; }

        /* Mecanismo de fade + mostrar mais/menos (reutilizável) */
        .fade-wrap { position: relative; overflow: hidden; transition: max-height .3s ease; }
        .fade-wrap.is-collapsed::after {
            content: ""; position: absolute; left: 0; right: 0; bottom: 0; height: 100%;
            background: linear-gradient(to bottom, rgba(250,246,237,0), #fff 92%);
            pointer-events: none;
        }
        .fade-toggle {
            display: inline-flex; align-items: center; gap: 6px;
            margin-top: 10px; font-size: 12px; font-weight: 500; color: var(--accent);
            padding: 4px 0; transition: opacity .15s;
        }
        .fade-toggle:hover { opacity: 0.7; }
        .fade-toggle i { font-size: 11px; transition: transform .2s; }

        #detalhamento-chamado:not(:empty) { margin-top: 1rem; }

        /* Abas do detalhamento (textual, sem ícones, alinhado à tela) */
        .det-tabs { display: flex; gap: 2px; flex-wrap: wrap; border-bottom: 1px solid var(--line); margin-bottom: 14px; }
        .det-tab {
            border: 0; background: transparent; cursor: pointer;
            padding: 8px 12px; font-size: 12px; color: var(--ink-mute);
            border-bottom: 2px solid transparent; margin-bottom: -1px;
            transition: color .12s, border-color .12s;
        }
        .det-tab:hover { color: var(--ink-soft); }
        .det-tab.active { color: var(--ink); border-bottom-color: var(--ink); font-weight: 500; }
        .det-tabpanel { display: none; }
        .det-tabpanel.active { display: block; }
        .det-empty { font-size: 13px; color: var(--ink-mute); padding: 4px 0; }


        /* Evento de historico inline */
        .activity-event { display: flex; align-items: center; gap: 12px; padding: 6px 16px; font-size: 12px; color: var(--ink-mute); }
        .activity-event .activity-ico { width: 22px; height: 22px; border-radius: 9999px; background: var(--paper); border: 1px solid var(--line); display: inline-flex; align-items: center; justify-content: center; color: var(--ink-soft); font-size: 11px; flex-shrink: 0; }
        .activity-event .activity-text { flex: 1; line-height: 1.4; }
        .activity-event .activity-text strong { color: var(--ink); font-weight: 500; }
        .activity-event .activity-when { font-family: monospace; font-size: 10px; }

        /* Filtro da timeline */
        .timeline-filter { display: inline-flex; align-items: center; gap: 4px; background: #fff; border: 1px solid var(--line); border-radius: 8px; padding: 3px; }
        .timeline-filter button { display: inline-flex; align-items: center; gap: 6px; padding: 5px 10px; border-radius: 5px; font-size: 12px; color: var(--ink-soft); transition: all .15s; }
        .timeline-filter button:hover { color: var(--ink); }
        .timeline-filter button.active { background: var(--ink); color: var(--paper); }
        .timeline-filter .count-badge { font-family: monospace; font-size: 10px; padding: 1px 5px; border-radius: 4px; background: rgba(26,22,18,0.06); }
        .timeline-filter button.active .count-badge { background: rgba(255,255,255,0.12); color: var(--paper); }

        /* Property row */
        .prop-row { display: grid; grid-template-columns: 110px 1fr; gap: 12px; align-items: center; padding: 8px 0; border-bottom: 1px solid var(--line-soft); }
        .prop-row:last-child { border-bottom: 0; }
        .prop-label {  font-size: 12px; letter-spacing: 0.05em; color: var(--ink-mute); }
        .prop-trigger { display: inline-flex; align-items: center; gap: 8px; font-size: 12px; padding: 4px 8px; border-radius: 4px; margin-left: -8px; transition: background .15s; }
        .prop-trigger:hover { background: var(--paper); }
        .prop-static { font-size: 12px; color: var(--ink); padding: 4px 0; }

        /* Botões custom */
        .btn-c { display: inline-flex; align-items: center; gap: 8px; padding: 8px 14px; border-radius: 6px; font-size: 12px; font-weight: 500; transition: all .15s; border: 0; cursor: pointer; line-height: 1; }
        .btn-c-primary { background: var(--ink); color: var(--paper); } .btn-c-primary:hover { background: #2a2620; color: var(--paper); }
        .btn-c-success { background: var(--sage-deep); color: var(--sage-soft); } .btn-c-success:hover { background: #2c4124; color: var(--sage-soft); }
        .btn-c-amber { background: var(--amber); color: white; } .btn-c-amber:hover { background: #b58a1f; color: white; }
        .btn-c-ghost { color: var(--ink-soft); padding: 8px 12px; } .btn-c-ghost:hover { background: var(--paper); color: var(--ink); }

        /* Composer */
        .composer { border-radius: 12px; border: 1px solid var(--line); background: #fff; overflow: hidden; }
        .mode-btn { padding: 6px 12px; border-radius: 6px; font-size: 12px; transition: all .15s; color: var(--ink-soft); }
        .mode-btn:hover { background: var(--paper); }
        .mode-btn.active.publica { background: var(--ink); color: var(--paper); font-weight: 500; }
        .mode-btn.active.nota    { background: var(--amber); color: white; font-weight: 500; }
        .composer-body { margin: 8px 12px; border-radius: 8px; padding: 12px 16px; background: var(--paper-light); }
        .composer-body.nota { background: var(--amber-paper); }
        .composer-toolbar { padding: 10px 12px; margin-top: 5px; border-top: 1px solid var(--line-soft); display: flex; align-items: center; justify-content: space-between; gap: 8px; flex-wrap: wrap; }
        .tool-ico-btn { color: var(--ink-mute); padding: 6px; border-radius: 4px; transition: all .15s; }
        .tool-ico-btn:hover { color: var(--ink); background: var(--paper); }

        /* Quill */
        .composer .ql-toolbar.ql-snow, .composer .ql-container.ql-snow { border: 0; }
        .composer .ql-toolbar.ql-snow { padding: 6px 12px; border-bottom: 1px solid var(--line-soft); background: transparent; }
        .composer .ql-container.ql-snow { font-size: 14px; }
        .composer .ql-editor { min-height: 100px; padding: 0; line-height: 1.6; color: var(--ink); }
        .composer .ql-editor.ql-blank::before { font-style: normal; color: var(--ink-mute); left: 0; right: 0; }
        .composer .ql-snow .ql-stroke { stroke: var(--ink-mute); }
        .composer .ql-snow .ql-fill { fill: var(--ink-mute); }
        .composer .ql-snow button:hover .ql-stroke, .composer .ql-snow button.ql-active .ql-stroke { stroke: var(--ink); }
        .composer .ql-snow button:hover .ql-fill, .composer .ql-snow button.ql-active .ql-fill { fill: var(--ink); }
        .composer .ql-snow button:hover, .composer .ql-snow button.ql-active { background: var(--paper); border-radius: 4px; }

        /* Anexos  */
        .uploads-zone { padding: 8px 12px 0; display: flex; flex-wrap: wrap; gap: 8px; }
        .uploads-zone:empty { display: none; }
        .upload-chip {
            display: inline-flex; align-items: center; gap: 8px;
            padding: 6px 10px 6px 8px;
            border-radius: 6px;
            border: 1px solid var(--line);
            background: #fff;
            font-size: 12px;
            max-width: 280px;
        }
        .upload-chip.uploading { background: var(--paper-light); border-style: dashed; }
        .upload-chip.error     { background: rgba(181,58,55,0.05); border-color: rgba(181,58,55,0.3); }
        .upload-chip .uc-ico {
            width: 24px; height: 24px; border-radius: 4px;
            display: inline-flex; align-items: center; justify-content: center;
            background: var(--paper); color: var(--ink-soft); flex-shrink: 0; font-size: 12px;
        }
        .upload-chip.uploading .uc-ico { color: var(--ink-mute); }
        .upload-chip.error     .uc-ico { background: rgba(181,58,55,0.1); color: var(--rose); }
        .upload-chip .uc-info { display: flex; flex-direction: column; min-width: 0; line-height: 1.2; }
        .upload-chip .uc-name { color: var(--ink); text-overflow: ellipsis; overflow: hidden; white-space: nowrap; }
        .upload-chip .uc-status { font-size: 10px; color: var(--ink-mute); font-family: monospace; }
        .upload-chip.error .uc-status { color: var(--rose); }
        .upload-chip .uc-actions { display: inline-flex; align-items: center; gap: 2px; margin-left: 4px; }
        .upload-chip .uc-btn {
        width: 20px; height: 20px; border-radius: 4px;
        color: var(--ink-mute); display: inline-flex; align-items: center; justify-content: center;
        transition: all .15s; font-size: 11px;
        }
        .upload-chip .uc-btn:hover { background: var(--paper); color: var(--ink); }
        .upload-chip.error .uc-btn:hover { background: rgba(181,58,55,0.1); color: var(--rose); }

        /* Spinner */
        .uc-spinner {
            width: 12px; height: 12px;
            border: 2px solid var(--line); border-top-color: var(--accent);
            border-radius: 9999px;
            animation: spin 0.7s linear infinite;
        }
        @keyframes spin { to { transform: rotate(360deg); } }

        /* Attachment chip */
        .att { display: flex; align-items: center; gap: 10px; border-radius: 8px; border: 1px solid var(--line); background: var(--paper-light); padding: 8px 12px; cursor: pointer; transition: background .15s; }
        .att:hover { background: var(--paper); }
        .att .att-ico { width: 32px; height: 32px; border-radius: 6px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
        .att .att-ico.image { background: rgba(58,85,48,0.1); color: var(--sage-deep); }
        .att .att-ico.doc   { background: rgba(26,22,18,0.08); color: var(--ink-soft); }

        /* SLA card */
        .sla-card { border-radius: 12px; background: var(--ink); color: var(--paper); padding: 1rem; }
        .sla-bar { height: 4px; background: rgba(255,255,255,0.08); border-radius: 9999px; overflow: hidden; }
        .sla-bar > div { height: 100%; border-radius: 9999px; }

        /* Dropdowns */
        .prop-trigger.dropdown-toggle::after { display: none; }   /* esconde a caret padrão */
        .dropdown-menu.pd-menu {
            min-width: 240px;
            border: 1px solid var(--line);
            border-radius: 10px;
            box-shadow: 0 1px 0 rgba(26,22,18,0.04), 0 12px 32px -10px rgba(26,22,18,0.18);
            padding: 6px;
            margin-top: 6px !important;
        }
        .pd-menu-search { display: flex; align-items: center; gap: 6px; padding: 6px 10px; border-bottom: 1px solid var(--line-soft); margin: -6px -6px 6px; }
        .pd-menu-search input { background: transparent; font-size: 12px; border: 0; outline: 0; flex: 1; color: var(--ink); }
        .pd-item { display: flex; align-items: center; gap: 10px; width: 100%; padding: 8px 10px; border-radius: 6px; text-align: left; font-size: 13px; color: var(--ink); background: transparent; border: 0; }
        .pd-item:hover, .pd-item:focus { background: var(--paper); color: var(--ink); outline: 0; }
        .pd-item.is-selected { background: transparent; }
        .pd-item .pd-check { margin-left: auto; color: var(--accent); opacity: 0; font-size: 12px; }
        .pd-item.is-selected .pd-check { opacity: 1; }
        .pd-item .pd-sub { font-size: 11px; color: var(--ink-mute); }

        /* Containers */
        .frame-wide   { max-width: 1500px; }
        .input-search { padding: 6px 12px 6px 32px; font-size: 12px; background: var(--paper); border-radius: 6px; border: 1px solid transparent; width: 290px; }
        .input-search:focus { outline: 0; border-color: var(--line); background: #fff; }
        .select-c { font-size: 11px; background: transparent; color: var(--ink-soft); border: 1px solid var(--line); border-radius: 4px; padding: 6px 8px; }

        /* Mensagem de erro */
        .error-message {
            background-color: #f8d7da;
            border: 1px solid #dc3545;
            color: #842029;
            border-radius: 8px;
            text-align: center;
            margin: 40px auto;
            max-width: 600px;
        }
        .error-message i {
            font-size: 3rem;
            margin-bottom: 15px;
        }
        .error-message h4 {
            margin-bottom: 10px;
        }
    </style>
</head>
<body>

<div id="loadingOverlay" class="position-fixed top-0 start-0 w-100 h-100 d-none justify-content-center align-items-center" style="background: rgba(255,255,255,0.8); z-index: 9999;">
    <div class="spinner-border text-primary" role="status">
        <span class="visually-hidden">Carregando...</span>
    </div>
</div>

<!-- container montado dinamicamente pelas funções de renderização -->
<div id="app"></div>

<script src="https://cdn.jsdelivr.net/npm/quill@2.0.2/dist/quill.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jquery@3.7.1/dist/jquery.min.js"></script>

<script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
<script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
<script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

<script>
    window.FormularioChamado = (function () {
        var registroPendente = null;
        function registrar(modulo) { registroPendente = modulo; }
        function obterRegistro() {
            var m = registroPendente;
            registroPendente = null;
            return m;
        }
        return { registrar: registrar, obterRegistro: obterRegistro };
    })();

    var cacheTipos = {};

    function carregarModuloTipo(caminho) {
        if (cacheTipos[caminho]) return Promise.resolve(cacheTipos[caminho]);
        return $.getScript(caminho).then(function () {
            var modulo = FormularioChamado.obterRegistro();
            if (!modulo || typeof modulo.secoesView !== 'function') {
                throw new Error('Tipo sem secoesView válido: ' + caminho);
            }
            cacheTipos[caminho] = modulo;
            return modulo;
        });
    }

    // monta a estrutura de abas (uma por seção) a partir da lista declarada pelo tipo
    function montarAbasDetalhamento(secoes) {
        if (!secoes || !secoes.length) return '';

        var navs = secoes.map(function (s, i) {
            return '<button type="button" class="det-tab' + (i === 0 ? ' active' : '') +
                   '" data-det-tab="' + i + '">' + s.titulo + '</button>';
        }).join('');

        var panes = secoes.map(function (s, i) {
            return '<div class="det-tabpanel' + (i === 0 ? ' active' : '') +
                   '" data-det-pane="' + i + '">' + s.html + '</div>';
        }).join('');

        return '<div class="det-tabs">' + navs + '</div><div class="det-tabcontent">' + panes + '</div>';
    }

    // alterna a aba ativa do detalhamento
    function ativarAbaDetalhamento($host, idx) {
        $host.find('.det-tab').removeClass('active');
        $host.find('.det-tabpanel').removeClass('active');
        $host.find('.det-tab[data-det-tab="' + idx + '"]').addClass('active');
        $host.find('.det-tabpanel[data-det-pane="' + idx + '"]').addClass('active');
    }

    function montarDetalhamento(chamado) {
        var $host = $('#detalhamento-chamado');
        if (!$host.length || !chamado || !chamado.detalhamento) return;

        var caminho = chamado.tipoCaminho;
        if (!caminho) return;

        var dados;
        try {
            dados = (typeof chamado.detalhamento === 'string')
                ? JSON.parse(chamado.detalhamento)
                : chamado.detalhamento;
        } catch (e) {
            console.error('Detalhamento inválido:', e);
            return;
        }

        carregarModuloTipo(caminho)
            .then(function (modulo) {
                var secoes = modulo.secoesView(dados) || [];
                $host.html(montarAbasDetalhamento(secoes));

                // troca de aba
                $host.off('click.detTab').on('click.detTab', '.det-tab', function () {
                    ativarAbaDetalhamento($host, $(this).attr('data-det-tab'));
                });

                // recalcula o fade, pois a altura mudou após a injeção
                if (typeof aplicarFade === 'function') aplicarFade();
            })
            .catch(function (err) {
                console.error(err);
            });
    }
</script>

<script>
    let catalogos = {
        // Lista de itens fixos
        status: [
            { id: 'ABERTO', label: 'Em aberto', sub: 'Chamado enviado', pillClass: 'pill-blue', dotColor: 'var(--blue)' },
            { id: 'EM_ANALISE', label: 'Em análise', sub: 'Atendente trabalhando no chamado', pillClass: 'pill-amber', dotColor: 'var(--amber)' },
            { id: 'AGUARDANDO_SOLICITANTE', label: 'Aguardando solicitante',  sub: 'Aguardando resposta do solicitante', pillClass: 'pill-slate', dotColor: '#6b7aa1' },
            { id: 'AGUARDANDO_TERCEIROS', label: 'Aguardando terceiros', sub: 'Bloqueado por área externa', pillClass: 'pill-stone', dotColor: '#8a7a5e' },
            { id: 'FINALIZADO', label: 'Finalizado', sub: 'Solução entregue', pillClass: 'pill-sage', dotColor: 'var(--sage)' }
        ],
        prioridade: [
            { id: 'BAIXA',   label: 'Baixa Prioridade',  sub: 'Sem impacto operacional',         icon: 'bi-thermometer-low',  iconColor: 'var(--ink-mute)', pillClass: 'pill-neutral' },
            { id: 'MEDIA',   label: 'Média Prioridade',  sub: 'Afeta um usuário',                icon: 'bi-thermometer-half', iconColor: 'var(--amber)',    pillClass: 'pill-amber' },
            { id: 'ALTA',    label: 'Alta Prioridade',   sub: 'Afeta uma equipe ou processo',    icon: 'bi-fire',             iconColor: 'var(--rose)',     pillClass: 'pill-rose' },
            { id: 'URGENTE', label: 'Urgente', label: 'Urgente',            sub: 'Sistema fora ou solicitante parado',  icon: 'bi-exclamation-octagon-fill', iconColor: 'var(--rose)', pillClass: 'pill-rose-strong' }
        ],

        // Listas dinâmicas (populados via banco de dados)
        atendentes: [],
        categorias: [],
        departamentos: [],
    };

    let chamado = null;
    let usuario = null;

    const composerState = { modo: 'publica' };

    const Upload = {
        ENDPOINT: '/corporativo/components/upload/upload.ajax.asp',
        FILE_CODE: 'chamado',
        TAMANHO_MAX_MB: 25,

        pending: {
            requester: [],
            agent: []
        },

        gerarLocalId() {
            return 'u_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7);
        },

        renderizarZona(zona) {
            const $zona = $(`.uploads-zone[data-uploads="${zona}"]`);
            $zona.html(this.pending[zona].map(Templates.uploadChip).join(''));
        },

        enviar(zona, file) {
            if (file.size > this.TAMANHO_MAX_MB * 1024 * 1024) {
                mostrarToast(`Arquivo <strong>${esc(file.name)}</strong> excede ${this.TAMANHO_MAX_MB}MB`);
                return;
            }

            const localId = this.gerarLocalId();
            const upload = { localId, fileName: file.name, file, status: 'uploading' };
            const hash = `${new Date().getTime()}_${Math.random().toString(36).substring(2, 8)}.${file.name.split('.').pop()}`;
            this.pending[zona].push(upload);
            this.renderizarZona(zona);

            const formData = new FormData();
            formData.append('fileCode', this.FILE_CODE);
            formData.append('file', file, hash);

            const xhr = new XMLHttpRequest();
            xhr.open('POST', this.ENDPOINT);
            xhr.onload = () => {
                let response;
                try {
                    response = JSON.parse(xhr.responseText); 
                } catch { 
                    response = { sucesso: false, mensagem: 'Resposta inválida do servidor' }; 
                }
                
                if (response.sucesso && response.arquivos) {
                    this.marcarSucesso(zona, localId, file.name, response.arquivos[0]);
                } else {
                    this.marcarErro(zona, localId, response.mensagem || 'Falha no upload');
                }
            };
            xhr.onerror = () => this.marcarErro(zona, localId, 'Erro de rede');
            xhr.send(formData);

            upload.xhr = xhr;
        },

        marcarSucesso(zona, localId, nomeOriginal, arquivo) {
            const upload = this.pending[zona].find(u => u.localId === localId);
            if (!upload) return;
            upload.status = 'done';
            upload.fileName = nomeOriginal;
            chamado.anexosPendentes.push({ name: nomeOriginal, uploadedName: arquivo.nome, localId });
            this.renderizarZona(zona);
        },

        marcarErro(zona, localId, mensagem) {
            const upload = this.pending[zona].find(u => u.localId === localId);
            if (!upload) return;
            upload.status = 'error';
            upload.mensagemErro = mensagem;
            this.renderizarZona(zona);
        },

        remover(zona, localId) {
            const idx = this.pending[zona].findIndex(u => u.localId === localId);
            if (idx === -1) return;
            const upload = this.pending[zona][idx];
            if (upload.status === 'uploading' && upload.xhr) upload.xhr.abort();
            this.pending[zona].splice(idx, 1);
            chamado.anexosPendentes = chamado.anexosPendentes.filter(a => a.localId !== localId);
            this.renderizarZona(zona);
        },

        retentar(zona, localId) {
            const upload = this.pending[zona].find(u => u.localId === localId);
            if (!upload || upload.status !== 'error') return;
            upload.status = 'uploading';
            delete upload.mensagemErro;
            this.renderizarZona(zona);
            this.pending[zona] = this.pending[zona].filter(u => u.localId !== localId);
            this.enviar(zona, upload.file);
        },

        limparZona(zona) {
            this.pending[zona] = [];
            this.renderizarZona(zona);
        }
    };

    // Helpers
    const esc = s => $('<div>').text(String(s)).html();

    function obterIniciais(nome) {
        if (!nome) return '?';
        const partes = String(nome).trim().split(/\s+/).filter(Boolean);
        if (partes.length === 0) return '?';
        if (partes.length === 1) return partes[0].slice(0, 2).toUpperCase();
        const primeira = partes[0][0];
        const ultima   = partes[partes.length - 1][0];
        return (primeira + ultima).toUpperCase();
    }

    const Catalogo = {
        status:       id => catalogos.status.find(o => o.id === id),
        prioridade:   id => catalogos.prioridade.find(o => o.id === id),
        atendente:    id => catalogos.atendentes.find(o => o.id === id),
        categoria:    id => catalogos.categorias.find(o => o.id === id),
        departamento: id => catalogos.departamentos.find(o => o.id === id)
    };

    const Permissoes = {
        podeVerNotasInternas:    usuario => usuario.tipo === 'ATENDENTE',
        podeEditarPropriedades:  usuario => usuario.tipo === 'ATENDENTE',
        podeEscrevernotaInterna: usuario => usuario.tipo === 'ATENDENTE',
        podeVerSlaDetalhado:     usuario => usuario.tipo === 'ATENDENTE',
        podeVerHistoricoSolicitante: usuario => usuario.tipo === 'ATENDENTE',
    };

    // Templates (recebem dados, retornam HTML)
    const Templates = {
        avatar(initials, cls, size) {
            return `<span class="avatar ${cls || 'av-ink'} av-${size || 38}">${esc(initials)}</span>`;
        },

        pill(label, cls, withDot) {
            return `<span class="pill ${cls}">${withDot ? '<span class="dot"></span>' : ''}${esc(label)}</span>`;
        },

        mensagem(m) {
            const anexos = (m.anexos || []).length === 0 ? '' : `
            <div class="row g-2 mt-3" style="padding-left:50px;">
                ${m.anexos.map(a => `
                <div class="col-md-6">
                    <a class="att" href="${`/Clientes_Docs/${a.caminho}`}" download="${a.nome}" title="Baixar ${a.nome}">
                        <div class="att-ico ${a.tipo}">
                            <i class="bi bi-${a.tipo === 'image' ? 'image' : 'file-earmark-text'}"></i>
                        </div>
                        <div class="flex-1 min-w-0">
                            <div class="text-12 fw-medium text-truncate">${esc(a.nome)}</div>
                        </div>
                        <i class="bi bi-download download text-mute" style="font-size:14px;"></i>
                    </a>
                </div>`).join('')}
            </div>`;

            return `
            <article class="msg soft-shadow" data-event="conversa">
                <header class="d-flex align-items-start gap-3 mb-3">
                    ${Templates.avatar(obterIniciais(m.nomeAutor), m.tipoAutor == 'ATENDENTE' ? 'av-ink' : 'av-sand', 38)}
                    <div class="flex-1 min-w-0">
                        <div class="d-flex align-items-baseline gap-2 flex-wrap">
                            <span class="fw-medium text-14">${esc(m.nomeAutor)}</span>
                            <span class="text-11 text-mute">· ${esc(m.tipoAutor)}</span>
                        </div>
                        <div class="d-flex align-items-center gap-2 flex-wrap" style="margin-top:2px;">
                            <span class="text-11 text-mute font-mono">${formatarParaDataBR(m.dataCriacao)}</span>
                        </div>
                    </div>
                </header>
                <div class="msg-body">${m.descricao}</div>
                ${anexos}
            </article>`;
        },

        nota(n) {
            const anexos = (n.anexos || []).length === 0 ? '' : `
            <div class="row g-2 mt-3" style="padding-left:50px;">
                ${n.anexos.map(a => `
                <div class="col-md-6">
                    <a class="att" href="${`/Clientes_Docs/${a.caminho}`}" download="${a.nome}" title="Baixar ${a.nome}">
                        <div class="att-ico ${a.tipo}">
                            <i class="bi bi-${a.tipo === 'image' ? 'image' : 'file-earmark-text'}"></i>
                        </div>
                        <div class="flex-1 min-w-0">
                            <div class="text-12 fw-medium text-truncate">${esc(a.nome)}</div>
                        </div>
                        <i class="bi bi-download download text-mute" style="font-size:14px;"></i>
                    </a>
                </div>`).join('')}
            </div>`;

            return `
            <article class="note soft-shadow" data-event="nota">
                <header class="d-flex align-items-start gap-3 mb-2">
                ${Templates.avatar(obterIniciais(n.nomeAutor), 'av-sand', 34)}
                <div class="flex-1">
                    <div class="d-flex align-items-baseline gap-2 flex-wrap">
                        <span class="fw-medium text-13">${esc(n.nomeAutor)}</span>
                        <span class="text-11" style="color:#7a5a12;">· ${esc(n.tipoAutor)}</span>
                    </div>
                    <div class="d-flex align-items-center gap-2" style="margin-top:2px;">
                        <span class="text-11 font-mono" style="color:#7a5a12;opacity:0.8;">${formatarParaDataBR(n.dataCriacao)}</span>
                        <span class="text-10 d-inline-flex align-items-center gap-1 px-2 py-1 rounded" style="background:var(--amber-soft);color:#7a5a12;">
                            <i class="bi bi-lock"></i> apenas equipe
                        </span>
                    </div>
                </div>
                </header>
                <div class="msg-body">${n.descricao}</div>
                ${anexos}
            </article>`;
        },

        historico(a) {
            return `
            <div class="activity-event" data-event="historico">
                <span class="activity-ico"><i class="bi bi-gear"></i></span>
                <span class="activity-text">${a.descricao}</span>
                <span class="activity-when text-mute">${formatarParaDataBR(a.dataCriacao)}</span>
            </div>`;
        },

        itemHistorico(h) {
            return `
            <a href="ChamadoDetalhes.asp?id=${h.Id}" target="_blank" class="d-flex align-items-center gap-3 py-2 px-2 rounded cursor-pointer hover-bg-paper" style="margin-left:-8px;margin-right:-8px;">
                <span style="width:6px;height:6px;border-radius:9999px;background:var(--sage);flex-shrink:0;"></span>
                <div class="flex-1 min-w-0">
                    <div class="text-12 text-truncate">${esc(h.Titulo)}</div>
                    <div class="text-10 text-mute font-mono">${esc(h.Id)} · ${formatarParaDataBR(h.DataCriacao)}</div>
                </div>
            </a>`;
        },

        uploadChip(upload) {
            const ehImagem = /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(upload.fileName);
            const iconeArquivo = ehImagem ? 'bi-image' : 'bi-file-earmark-text';

            let iconeStatus, textoStatus, btnAcoes;
            if (upload.status === 'uploading') {
                iconeStatus = '<div class="uc-spinner"></div>';
                textoStatus = 'Enviando...';
                btnAcoes    = `<button class="uc-btn" data-cancel="${esc(upload.localId)}" title="Cancelar"><i class="bi bi-x"></i></button>`;
            } else if (upload.status === 'error') {
                iconeStatus = `<i class="bi ${iconeArquivo}"></i>`;
                textoStatus = esc(upload.mensagemErro || 'Falha no envio');
                btnAcoes    = `
                    <button class="uc-btn" data-retry="${esc(upload.localId)}" title="Tentar novamente"><i class="bi bi-arrow-clockwise"></i></button>
                    <button class="uc-btn" data-remove="${esc(upload.localId)}" title="Remover"><i class="bi bi-x"></i></button>`;
            } else {
                iconeStatus = `<i class="bi ${iconeArquivo}"></i>`;
                textoStatus = 'Pronto';
                btnAcoes    = `<button class="uc-btn" data-remove="${esc(upload.localId)}" title="Remover"><i class="bi bi-x"></i></button>`;
            }

            return `
                <div class="upload-chip ${upload.status}" data-upload-id="${esc(upload.localId)}">
                    <span class="uc-ico">${iconeStatus}</span>
                    <span class="uc-info">
                    <span class="uc-name">${esc(upload.fileName)}</span>
                    <span class="uc-status">${textoStatus}</span>
                    </span>
                    <span class="uc-actions">${btnAcoes}</span>
                </div>`;
        }
    };

    function construirTimeline(chamado, usuario) {
        const mensagens = chamado.mensagens.filter(item => item.notaInterna == '0').map(m => {
            return {  ...m, tipo: 'mensagem' };
        })

        const notas = chamado.mensagens.filter(item => item.notaInterna == '1').map(n => {
            return {  ...n, tipo: 'nota' };
        })

        const historico = chamado.historico.map(h => {
            return {  ...h, tipo: 'historico' };
        })

        const eventos = [
            ...mensagens,
            ...(Permissoes.podeVerNotasInternas(usuario) ? notas : []),
            ...historico
        ];

        return eventos.slice().sort((a, b) => new Date(a.dataCriacao) - new Date(b.dataCriacao));
    }

    function renderizarEvento(e) {
        if (e.tipo === 'mensagem')  return Templates.mensagem(e);
        if (e.tipo === 'nota')      return Templates.nota(e);
        if (e.tipo === 'historico') return Templates.historico(e);
        return '';
    }

    function renderizarCabecalho(chamado, usuario) {
        const status     = Catalogo.status(chamado.status);
        const prioridade = Catalogo.prioridade(chamado.prioridade);
        const categoria  = Catalogo.categoria(chamado.idCategoria);
        const ehAtendente = usuario.tipo === 'ATENDENTE';
        const titulo = `<h1 class="font-display text-28" style="letter-spacing:-0.01em;">${esc(chamado.titulo)} #${chamado.id}</h1>`;
        const botaoVoltar = ehAtendente ? `
            <button type="button" onclick="navegarParaGrade()" class="btn btn-sm btnVoltar me-2 fs-5">
                <i class="bi bi-arrow-left"></i>
            </button>
            ` : ''

        const meta = `
            <div class="mt-3 d-flex align-items-center gap-3 text-12 text-mute flex-wrap">
                <span class="d-inline-flex align-items-center gap-1"><i class="bi bi-calendar3"></i> Aberto em ${esc(formatarParaDataBR(chamado.dataCriacao))}</span>
                <span class="d-inline-flex align-items-center gap-1"><i class="bi bi-clock"></i> Última atualização em ${esc(formatarParaDataBR(chamado.dataUltimaAtualizacao))}</span>
                <span class="d-inline-flex align-items-center gap-1"><i class="bi bi-chat-square-text"></i> ${chamado.mensagens.length} mensagens</span>
            </div>` ;

        const toolbar = chamado.status == 'FINALIZADO' ? 
            `<div class="d-flex gap-2 mt-4">
                <button onclick="confirmacaoReabertura()" class="w-fit btn-c w-fit btn-c-primary" data-action="resolve"><i class="bi bi-arrow-clockwise"></i> Reabrir</button>
                ${chamado.avaliado == '0' && !ehAtendente ? `<button onclick="navegarParaAvaliacao()" class="w-fit btn-c w-fit btn-c-success" data-action="resolve"><i class="bi bi-star"></i> Avaliar Chamado</button>` : ''}
            </div>` 
            : '';

        return `
            <div class="d-flex align-items-start justify-content-between gap-4 pb-4 border-bottom border-line flex-wrap">
                <div class="flex-1 min-w-0">
                    <div class="d-flex align-items-baseline">
                        ${botaoVoltar}
                        ${titulo}
                    </div>
                    ${meta}
                    ${toolbar}
                </div>
            </div>`;
    }

    function renderizarFiltroTimeline(chamado, usuario) {
        if (usuario.tipo !== 'ATENDENTE') return '';

        const eventos = construirTimeline(chamado, usuario);
        const conversa  = eventos.filter(e => e.tipo === 'mensagem').length;
        const nota      = eventos.filter(e => e.tipo === 'nota').length;
        const historico = eventos.filter(e => e.tipo === 'historico').length;

        return `
            <div class="timeline-filter mb-3">
                <button class="active" data-filter="all">Tudo <span class="count-badge">${eventos.length}</span></button>
                <button data-filter="conversa"><i class="bi bi-chat-square"></i> Conversa <span class="count-badge">${conversa}</span></button>
                <button data-filter="nota"><i class="bi bi-sticky"></i> Notas <span class="count-badge">${nota}</span></button>
                <button data-filter="historico"><i class="bi bi-activity"></i> Atividade <span class="count-badge">${historico}</span></button>
            </div>
        `;
    }

    function renderizarDescricao(chamado, usuario) {
        const anexos = (chamado && (chamado.anexos || []).length) ? `
            <div class="row g-2 mt-3">
            ${chamado.anexos.map(a => `
                <div class="col-md-6">
                     <a class="att" href="${`/Clientes_Docs/${a.caminho}`}" download="${a.nome}" title="Baixar ${a.nome}">
                        <div class="att-ico ${a.tipo}">
                            <i class="bi bi-${a.tipo === 'image' ? 'image' : 'file-earmark-text'}"></i>
                        </div>
                        <div class="flex-1 min-w-0">
                            <div class="text-12 fw-medium text-truncate">${esc(a.nome)}</div>
                        </div>
                        <i class="bi bi-download download text-mute" style="font-size:14px;"></i>
                    </a>
                </div>`).join('')}
            </div>` : '';

        // O detalhamento é injetado de forma assíncrona em #detalhamento-chamado
        // (após carregar o .js do tipo) por montarDetalhamento(), chamado em renderizarTela.
        const detalhamento = '<div id="detalhamento-chamado"></div>';

        return `
            <div class="ticket-origin mb-4">
                <div class="ticket-origin-head">
                    ${Templates.avatar(obterIniciais(chamado.solicitante.nome), 'av-sand', 38)}
                    <div class="flex-1 min-w-0">
                        <div class="d-flex align-items-baseline gap-2 flex-wrap">
                            <span class="origin-name">${esc(chamado.solicitante.nome)}</span>
                        </div>
                        <div class="origin-when" style="margin-top:2px;">abriu o chamado · ${esc(formatarParaDataBR(chamado.dataCriacao))}</div>
                    </div>
                </div>
                <div class="ticket-origin-body">
                    <div class="fade-wrap" data-fade data-fade-max="180">
                        <div class="desc-text">${chamado.descricao}</div>
                        ${anexos}
                        ${detalhamento}
                    </div>
                </div>
            </div>`;
    }


    function renderizarTimeline(chamado, usuario) {
        const eventos = construirTimeline(chamado, usuario);
        return `<div id="timeline" class="d-flex flex-column gap-3">${eventos.map(renderizarEvento).join('')}</div>`;
    }

    function renderizarAreaResposta(chamado, usuario) {
        if (usuario.tipo === 'SOLICITANTE') {
            return `
            <div class="composer soft-shadow mt-4">
                <div class="px-4 py-2 border-bottom border-line d-flex align-items-center justify-content-between" style="background:var(--paper-light);">
                    <div class="d-flex align-items-center gap-2">
                        <i class="bi bi-reply text-mute" style="font-size:14px;"></i>
                        <span class="text-12 fw-medium">Responder Atendente</span>
                    </div>
                    <span class="text-11 text-mute">Resposta visível para todos os participantes</span>
                </div>
                <div class="p-0">
                    <div id="requester-toolbar" style="padding:6px 12px;border-bottom:1px solid var(--line-soft);">
                        <span class="ql-formats"><button class="ql-bold"></button><button class="ql-italic"></button><button class="ql-underline"></button></span>
                        <span class="ql-formats"><button class="ql-list" value="bullet"></button><button class="ql-list" value="ordered"></button></span>
                        <span class="ql-formats"><button class="ql-link"></button><button class="ql-clean"></button></span>
                    </div>
                    <div id="requester-editor" style="padding:16px 20px;min-height:100px;"></div>
                </div>
                <div class="uploads-zone" data-uploads="requester"></div>
                <div class="composer-toolbar">
                    <div class="d-flex align-items-center gap-1">
                        <input type="file" class="d-none" data-file-input="requester" multiple />
                        <button class="tool-ico-btn" data-trigger-upload="requester" title="Anexar arquivo">
                            <i class="bi bi-paperclip"></i>
                        </button>
                    </div>
                    <button id="send-btn-requester" class="btn-c btn-c-primary"><i class="bi bi-send"></i> Enviar resposta</button>
                </div>
            </div>`;
        }

        return `
            <div id="agent-composer" class="composer soft-shadow mt-4">
                <div class="d-flex align-items-center gap-1 px-3 pt-3 flex-wrap">
                    <button id="mode-publica" class="mode-btn active publica"><i class="bi bi-reply"></i> Resposta pública</button>
                    ${Permissoes.podeEscrevernotaInterna(usuario) ? '<button id="mode-nota" class="mode-btn nota"><i class="bi bi-sticky"></i> Nota interna</button>' : ''}
                </div>
                <div id="composer-body" class="composer-body">
                    <div id="composer-toolbar">
                        <span class="ql-formats"><button class="ql-bold"></button><button class="ql-italic"></button><button class="ql-underline"></button></span>
                        <span class="ql-formats"><button class="ql-list" value="bullet"></button><button class="ql-list" value="ordered"></button><button class="ql-blockquote"></button></span>
                        <span class="ql-formats"><button class="ql-link"></button><button class="ql-code-block"></button><button class="ql-clean"></button></span>
                    </div>
                    <div id="composer-editor"></div>
                </div>

                <div class="uploads-zone" data-uploads="agent"></div>

                <div class="composer-toolbar">
                    <div class="d-flex align-items-center gap-1 flex-wrap">
                        <input type="file" class="d-none" data-file-input="agent" multiple />
                        <button class="tool-ico-btn" data-trigger-upload="agent" title="Anexar arquivo">
                            <i class="bi bi-paperclip"></i>
                        </button>
                    </div>
                    <div class="d-flex align-items-center gap-2">
                        <select id="acaoAdicional" class="select-c">
                            <option selected value="manterAberto">Manter aberto</option>
                            <option value="finalizar">Finalizar após enviar</option>
                            <option value="aguardandoSolicitante">Aguardando solicitante</option>
                        </select>
                        <button id="send-btn" class="btn-c btn-c-primary"><i class="bi bi-send"></i> <span id="send-label">Enviar resposta</span></button>
                    </div>
                </div>
            </div>`;
        }

    function calcularStatusPrazo(dataInicio, dataFim) {
        const inicio = new Date(dataInicio);
        const fim = new Date(dataFim);
        const agora = new Date();

        if (fim - inicio <= 0) throw new Error('A data de fim deve ser posterior à data de início');

        const totalMs = fim - inicio;
        const decorridoMs = agora - inicio;
        const percentual = Math.round((decorridoMs / totalMs) * 100);

        const dentroPrazo = agora <= fim;

        const getColor = (p) => {
            if (p <= 50) return '#22c55e';
            if (p <= 75) return '#eab308';
            if (p <= 90) return '#f97316';
            return '#ef4444';
        };

        const getTempoRestante = () => {
            const restanteMs = fim - agora;
            if (restanteMs <= 0) return 'Prazo encerrado';

            const totalMinutos = Math.floor(restanteMs / 1000 / 60);
            const horas = Math.floor(totalMinutos / 60);
            const minutos = totalMinutos % 60;

            if (horas === 0) return `${minutos} min`;
            if (minutos === 0) return `${horas} h`;
            return `${horas} h e ${minutos} min`;
        };

        return {
            dentroPrazo,
            percentual: Math.min(Math.max(percentual, 0), 100),
            color: getColor(Math.max(percentual, 0)),
            tempoRestante: getTempoRestante(),
        };
    }

    function renderizarCardSlaOuStatus(chamado, usuario) {
        if (Permissoes.podeVerSlaDetalhado(usuario)) {
            const statusPrazo = calcularStatusPrazo(chamado.dataCriacao, chamado.dataPrevista);

            const badgeSLA = statusPrazo.dentroPrazo ? `
                <span class="pill pill-sage" style="background:var(--sage-deep); color:var(--sage-soft);">
                    <i class="bi bi-shield-check"></i> 
                    no prazo
                </span>
            ` : `
                <span class="pill pill-sage" style="background:#ef4444; color:var(--sage-soft);">
                    <i class="bi bi-shield-slash"></i> 
                    fora do prazo
                </span>
            `

            return `
            <div class="sla-card">
                <div class="d-flex align-items-center justify-content-between mb-3">
                    <h3 class="section-label" style="color:var(--ink-mute);">SLA</h3>
                    ${badgeSLA}
                </div>
                <div class="d-flex flex-column gap-3">
                <div>
                    <div class="d-flex justify-content-between text-11 mb-1">
                        <span class="text-mute">Tempo Restante</span>
                        <span class="font-mono">${statusPrazo.tempoRestante}</span>
                    </div>
                    <div class="sla-bar"><div style="background: ${statusPrazo.color}; width:${statusPrazo.percentual}%;"></div></div>
                </div>
                </div>
                <div class="mt-3 pt-3" style="border-top: 1px solid rgba(255,255,255,0.08);">
                    <div class="d-flex justify-content-between text-11">
                        <span class="text-mute">Prazo final</span>
                        <span class="font-mono">${esc(formatarParaDataBR(chamado.dataPrevista))}</span>
                    </div>
                </div>
            </div>`;
        }

        const status = Catalogo.status(chamado.status);

        return `
            <div class="card-custom soft-shadow">
                <h3 class="section-label mb-3">Status do chamado</h3>
                <div class="d-flex align-items-center gap-3 mb-4">
                    <div class="d-flex align-items-center justify-content-center" style="width:40px;height:40px;border-radius:9999px;background:var(--amber-soft);color:var(--amber);">
                        <i class="bi bi-stopwatch"></i>
                    </div>
                    <div>
                        <div class="text-14 fw-medium">${esc(status.label)}</div>
                        <div class="text-11 text-mute">${esc(status.sub)}</div>
                    </div>
                </div>
                <div class="d-flex flex-column gap-2 text-12">
                    <div class="d-flex justify-content-between">
                        <span class="text-mute">Aberto em</span>
                        <span class="font-mono text-11">${esc(formatarParaDataBR(chamado.dataCriacao))}</span>
                    </div>
                    <div class="d-flex justify-content-between">
                        <span class="text-mute">Resolução prevista</span>
                        <span class="font-mono text-11">${esc(formatarParaDataBR(chamado.dataPrevista))}</span>
                    </div>
                </div>
            </div>`;
    }

    function renderizarCardAtendente(chamado, usuario) {
        if (usuario.tipo !== 'SOLICITANTE' || !chamado.idAtendente) return '';

        return `
            <div class="card-custom soft-shadow">
                <h3 class="section-label mb-3">Quem está te atendendo</h3>
                <div class="d-flex align-items-center gap-3">
                    ${Templates.avatar(obterIniciais(chamado.nomeAtendente), 'av-ink', 44)}
                    <div class="flex-1">
                        <div class="text-14 fw-medium">${esc(chamado.nomeAtendente)}</div>
                    </div>
                </div>
            </div>`;
    }

    function renderizarCardPropriedades(chamado, usuario) {
        if (!Permissoes.podeEditarPropriedades(usuario)) return '';

        const status     = Catalogo.status(chamado.status);
        const prioridade = Catalogo.prioridade(chamado.prioridade);
        const atendente  = Catalogo.atendente(chamado.idAtendente);
        const categoria  = Catalogo.categoria(chamado.idCategoria);
        const depto      = Catalogo.departamento(chamado.idDepartamento);

        return `
            <div class="card-custom soft-shadow" style="padding:1rem;">
                <div class="d-flex align-items-center justify-content-between mb-2">
                    <h3 class="section-label">Propriedades</h3>
                </div>
                <div>
                    <div class="prop-row">
                        <span class="prop-label">Status</span>
                        <div class="dropdown">
                            <button id="trigger-status" class="prop-trigger dropdown-toggle" data-bs-toggle="dropdown" data-bs-offset="-8,4" aria-expanded="false">
                                <span style="width:6px;height:6px;border-radius:9999px;background:${status.dotColor};display:inline-block;"></span>
                                <span class="pd-text">${esc(status.label)}</span>
                                <i class="bi bi-chevron-down text-mute" style="font-size:10px;"></i>
                            </button>
                            <ul class="dropdown-menu pd-menu" data-pd="status"></ul>
                        </div>
                    </div>
                    <div class="prop-row">
                        <span class="prop-label">Prioridade</span>
                        <div class="dropdown">
                            <button id="trigger-priority" class="prop-trigger dropdown-toggle" data-bs-toggle="dropdown" data-bs-offset="-8,4" aria-expanded="false">
                                <i class="bi ${prioridade.icon} pd-icon" style="color:${prioridade.iconColor};"></i>
                                <span class="pd-text">${esc(prioridade.label)}</span>
                                <i class="bi bi-chevron-down text-mute" style="font-size:10px;"></i>
                            </button>
                            <ul class="dropdown-menu pd-menu" data-pd="priority"></ul>
                        </div>
                    </div>
                    <div class="prop-row">
                        <span class="prop-label">Departamento</span>
                        <div class="dropdown">
                            <button id="trigger-department" class="prop-trigger dropdown-toggle" data-bs-toggle="dropdown" data-bs-offset="-8,4" aria-expanded="false">
                                <span class="pd-text">${esc(depto.nome)}</span>
                                <i class="bi bi-chevron-down text-mute" style="font-size:10px;"></i>
                            </button>
                            <ul class="dropdown-menu pd-menu" data-pd="department"></ul>
                        </div>
                    </div>
                    <div class="prop-row">
                        <span class="prop-label">Atendente</span>
                        <div class="dropdown">
                            <button id="trigger-agent" class="prop-trigger dropdown-toggle" data-bs-toggle="dropdown" data-bs-offset="-8,4" data-bs-auto-close="outside" aria-expanded="false">
                                ${atendente ? 
                                    `${Templates.avatar(obterIniciais(atendente.nome), atendente.avCls, 18)} <span class="pd-text">${esc(atendente.nome)}</span>`
                                    : '<span class="pd-text">Nenhum atendente atribuído</span>'
                                }
                                <i class="bi bi-chevron-down text-mute" style="font-size:10px;"></i>
                            </button>
                            <ul class="dropdown-menu pd-menu" data-pd="agent" data-search="true"></ul>
                        </div>
                    </div>
                    <div class="prop-row">
                        <span class="prop-label">Categoria</span>
                        <div class="dropdown">
                            <button id="trigger-category" class="prop-trigger dropdown-toggle" data-bs-toggle="dropdown" data-bs-offset="-8,4" data-bs-auto-close="outside" aria-expanded="false">
                                ${categoria ? 
                                    `<span class="pd-text">${esc(categoria.nome)}</span>`
                                    : '<span class="pd-text">Nenhuma categoria definida</span>'
                                }
                                <i class="bi bi-chevron-down text-mute" style="font-size:10px;"></i>
                            </button>
                            <ul class="dropdown-menu pd-menu" data-pd="category" data-search="true"></ul>
                        </div>
                    </div>
                     <div class="prop-row">
                        <span class="prop-label">Centro de custo</span>
                        <div class="dropdown">
                            <div class="prop-trigger">
                                ${chamado.centroCusto ? 
                                    `<span class="pd-text">${esc(chamado.centroCusto)}</span>`
                                    : '<span class="pd-text">Nenhuma centro de custo definido</span>'
                                }
                            </div>
                        </div>
                    </div>
                </div>
            </div>`;
    }

    function renderizarCardSolicitante(chamado, usuario) {
        const s = chamado.solicitante;

        return `
            <div class="card-custom soft-shadow" style="padding:1rem;">
                <h3 class="section-label mb-3">Solicitante</h3>
                <div class="d-flex align-items-center gap-3">
                    ${Templates.avatar(obterIniciais(s.nome), 'av-sand', 42)}
                    <div class="flex-1 min-w-0">
                        <div class="text-14 fw-medium text-truncate">${esc(s.nome)}</div>
                    </div>
                </div>
                <div class="mt-3 pt-3 d-flex flex-column gap-1 text-12" style="border-top:1px solid var(--line-soft);">
                    <a class="d-flex align-items-center gap-2 text-soft hover-ink py-1 cursor-pointer">
                        <i class="bi bi-envelope text-mute" style="font-size:12px;"></i>${esc(s.email)}
                    </a>
                    <a class="d-flex align-items-center gap-2 text-soft hover-ink py-1 cursor-pointer">
                        <i class="bi bi-telephone text-mute" style="font-size:12px;"></i>${esc(s.telefone)}
                    </a>
                </div>
            </div>`;
    }

    function renderizarCardHistorico(chamado, usuario) {
        if (!Permissoes.podeVerHistoricoSolicitante(usuario)) return '';
        const temHistorico = chamado.solicitante.historico.length > 0;

        return `
            <div class="card-custom soft-shadow" style="padding:1rem;">
                <div class="d-flex align-items-center justify-content-between mb-2">
                    <h3 class="section-label">Histórico do solicitante</h3>
                    <a href="../Chamado/ChamadoGradeAtendente.asp?filtroSolicitante=${chamado.solicitante.email}" target="_blank" class="text-12 text-accent hover-underline cursor-pointer">Ver todos</a>
                </div>
                ${temHistorico ? 
                    `<div>${chamado.solicitante.historico.map(Templates.itemHistorico).join('')} </div>`  :
                    '<div class="text-12 text-mute">Nenhum histórico encontrado</div>'
                }
            </div>`;
    }

    function renderizarSidebar(chamado, usuario) {
        let cards;
        if (usuario.tipo === 'SOLICITANTE') {
            cards = [
                renderizarCardSlaOuStatus(chamado, usuario),
                renderizarCardAtendente(chamado, usuario),
            ];
        } else {
            cards = [
                renderizarCardSlaOuStatus(chamado, usuario),
                renderizarCardPropriedades(chamado, usuario),
                renderizarCardSolicitante(chamado, usuario),
                renderizarCardHistorico(chamado, usuario),
            ];
        }
        return `<aside class="col-lg-4 d-flex flex-column gap-3">${cards.join('')}</aside>`;
    }

    let agentQuill = null, requesterQuill = null;
    function inicializarEditores(usuario) {
        if (usuario.tipo === 'SOLICITANTE' && document.getElementById('requester-editor')) {
            requesterQuill = new Quill('#requester-editor', {
                theme: 'snow',
                modules: { toolbar: '#requester-toolbar' }
            });
        }
        if (usuario.tipo === 'ATENDENTE' && document.getElementById('composer-editor')) {
            agentQuill = new Quill('#composer-editor', {
                theme: 'snow',
                modules: { toolbar: '#composer-toolbar' }
            });
        }
    }


    function inicializarUpload(usuario) {
        const zonaAtiva = usuario.tipo === 'ATENDENTE' ? 'agent' : 'requester';
        Upload.renderizarZona(zonaAtiva);

        if (inicializarUpload._registrado) return;
        inicializarUpload._registrado = true;

        $(document).on('click', '[data-trigger-upload]', function () {
            const zona = $(this).data('trigger-upload');
            $(`[data-file-input="${zona}"]`).trigger('click');
        });

        $(document).on('change', '[data-file-input]', function () {
            const zona = $(this).data('file-input');
            const files = Array.from(this.files || []);
            files.forEach(f => Upload.enviar(zona, f));
            this.value = '';  
        });

        $(document).on('click', '.upload-chip [data-remove]', function () {
            const zona    = $(this).closest('.uploads-zone').data('uploads');
            const localId = $(this).data('remove');
            Upload.remover(zona, localId);
        });
        $(document).on('click', '.upload-chip [data-cancel]', function () {
            const zona    = $(this).closest('.uploads-zone').data('uploads');
            const localId = $(this).data('cancel');
            Upload.remover(zona, localId);
        });
        $(document).on('click', '.upload-chip [data-retry]', function () {
            const zona    = $(this).closest('.uploads-zone').data('uploads');
            const localId = $(this).data('retry');
            Upload.retentar(zona, localId);
        });
    }


    function inicializarDropdownsPropriedades(chamado, usuario) {
        if (!Permissoes.podeEditarPropriedades(usuario)) return;

        const renderItens = {
            status: (o, sel) => `
                <li>
                    <button class="dropdown-item pd-item ${sel ? 'is-selected' : ''}" type="button" data-id="${o.id}">
                        <span style="width:8px;height:8px;border-radius:9999px;background:${o.dotColor};flex-shrink:0;"></span>
                        <div class="flex-1"><div>${esc(o.label)}</div><div class="pd-sub">${esc(o.sub)}</div></div>
                        <i class="bi bi-check2 pd-check"></i>
                    </button>
                </li>`,
            prioridade: (o, sel) => `
                <li>
                    <button class="dropdown-item pd-item ${sel ? 'is-selected' : ''}" type="button" data-id="${o.id}">
                        <i class="bi ${o.icon}" style="color:${o.iconColor};flex-shrink:0;"></i>
                        <div class="flex-1"><div>${esc(o.label)}</div><div class="pd-sub">${esc(o.sub)}</div></div>
                        <i class="bi bi-check2 pd-check"></i>
                    </button>
                </li>`,
            atendente: (o, sel) => `
                <li>
                    <button class="dropdown-item pd-item ${sel ? 'is-selected' : ''}" type="button" data-id="${o.id}">
                        <div class="flex-1"><div>${esc(o.nome)}</div></div>
                        <i class="bi bi-check2 pd-check"></i>
                    </button>
                </li>`,
            departamento: (o, sel) => `
                <li>
                    <button class="dropdown-item pd-item ${sel ? 'is-selected' : ''}" type="button" data-id="${o.id}">
                        <i class="bi ${o.icon}" style="color:var(--ink-soft);flex-shrink:0;"></i>
                        <div class="flex-1"><div>${esc(o.nome)}</div></div>
                        <i class="bi bi-check2 pd-check"></i>
                    </button>
                </li>`,
            categoria: (o, sel) => `
                <li>
                    <button class="dropdown-item pd-item ${sel ? 'is-selected' : ''}" type="button" data-id="${o.id}">
                        <i class="bi ${o.icon}" style="color:var(--ink-soft);flex-shrink:0;"></i>
                        <div class="flex-1"><div>${esc(o.nome)}</div></div>
                        <i class="bi bi-check2 pd-check"></i>
                    </button>
                </li>`
        };

        const menus = {
            status:     { items: catalogos.status,        render: renderItens.status,        campo: 'statusId',       label: 'Status' },
            priority:   { items: catalogos.prioridade,    render: renderItens.prioridade,    campo: 'prioridadeId',   label: 'Prioridade' },
            agent:      { items: catalogos.atendentes,    render: renderItens.atendente,     campo: 'atendenteId',    label: 'Atendente' },
            category:   { items: catalogos.categorias,    render: renderItens.categoria,     campo: 'categoriaId',    label: 'Categoria' },
            department: { items: catalogos.departamentos, render: renderItens.departamento,  campo: 'departamentoId', label: 'Departamento' }
        };

        // Popula os menus (já estão no DOM, vindos do template)
        function popularMenu(key, filtro) {
            const m = menus[key];
            const $menu = $(`.dropdown-menu[data-pd="${key}"]`);
            
            const itensFiltrados = filtro
                ? m.items.filter(o => (o.label || o.nome).toLowerCase().includes(filtro.toLowerCase()))
                : m.items;

            const items = itensFiltrados.length
                ? itensFiltrados.map(o => m.render(o, o.id === chamado[m.campo])).join('')
                : '<li class="px-2 py-3 text-12 text-mute text-center">Nada encontrado</li>';

            $menu.html(items);
            if (filtro !== undefined) $menu.find('input').focus().val(filtro);
        }

        // Render inicial de todos os menus
        Object.keys(menus).forEach(key => popularMenu(key));

        // Evento ao selecionar item
        $(document).on('click', '.dropdown-menu[data-pd] .pd-item', function () {
            const key = $(this).closest('.dropdown-menu').data('pd');
            const menu  = menus[key];
            const novoItem = menu.items.find(o => o.id == $(this).data('id'));

            if (!novoItem || novoItem.id === chamado[menu.campo]) return;

            switch (key) {
                case 'status':
                    if (novoItem.id == 'FINALIZADO') {
                        confirmacaoFinalizacao(novoItem.id);
                    } else {
                        alterarStatus(novoItem.id);
                    }
                    break;
                case 'priority':
                    alterarPrioridade(novoItem.id);
                    break;
                case 'department':
                    alterarDepartamento(novoItem.id);
                    break;
                case 'agent':
                    alterarAtendente(novoItem.id);
                    break;
                case 'category':
                    alterarCategoria(novoItem.id, novoItem.nome);
                    break;
                default:
                    console.error('Propriedade inválida');
                    break;
            }
        });
    }

    function inicializarAreaResposta(usuario) {
        if (usuario.tipo === 'ATENDENTE') {
            composerState.modo = 'publica';

            function setModo(modo) {
                composerState.modo = modo;

                $('#mode-publica').toggleClass('active', modo === 'publica');
                $('#mode-nota').toggleClass('active', modo === 'nota');
                $('#composer-body').toggleClass('nota', modo === 'nota');

                const $btn = $('#send-btn');

                if (modo === 'nota') {
                    $btn.removeClass('btn-c-primary').addClass('btn-c-amber');
                    $('#send-label').text('Salvar nota');
                } else {
                    $btn.removeClass('btn-c-amber').addClass('btn-c-primary');
                    $('#send-label').text('Enviar resposta');
                }
            }

            $('#mode-publica').on('click', () => setModo('publica'));
            $('#mode-nota').on('click',    () => setModo('nota'));
            $('#send-btn').on('click', () => enviarMensagem());
        }

        if (usuario.tipo === 'SOLICITANTE') {
            $('#send-btn-requester').on('click', () => enviarMensagem());
        }
    }


    function inicializarFiltroTimeline(chamado, usuario) {
        $('.timeline-filter button').on('click', function () {
            $('.timeline-filter button').removeClass('active');
            $(this).addClass('active');
            const filtro = $(this).data('filter');
            const eventos = construirTimeline(chamado, usuario).filter(e => {
                if (filtro === 'all') return true;
                if (filtro === 'conversa')  return e.tipo === 'mensagem';
                if (filtro === 'nota')      return e.tipo === 'nota';
                if (filtro === 'historico') return e.tipo === 'historico';
                return true;
            });
            $('#timeline').html(eventos.map(renderizarEvento).join(''));
        });
    }

    /**
     * Monta a tela inteira a partir de chamado + usuario.
     * É só essa função que conhece o layout do "miolo".
     */
    function renderizarTela(chamado, usuario) {
        if (!chamado) {
            $('#app').html(`
                <div class="error-message">
                    <i class="bi bi-exclamation-triangle"></i>
                    <h4>Não autorizado</h4>
                    <p>Voce nao tem permissao para visualizar esse chamado</p>
                </div>
            `)

            return;
        }
        
        var p1 = carregarDepartamentos(chamado.codPlanta);
        var p2 = carregarCategorias(chamado.idTipo, chamado.codPlanta);
        var p3 = carregarAtendentes(chamado.codPlanta, chamado.idDepartamento);

        $.when(p1, p2, p3).then(function() {
            const corpo = `
                <div class="container-fluid frame-wide mx-auto px-4 py-4">
                    ${renderizarCabecalho(chamado, usuario)}
                    <div class="row g-4 mt-2">
                        <div class="col-lg-8">
                            ${renderizarDescricao(chamado, usuario)}
                            ${renderizarFiltroTimeline(chamado, usuario)}
                            ${renderizarTimeline(chamado, usuario)}
                            ${renderizarAreaResposta(chamado, usuario)}
                        </div>
                        ${renderizarSidebar(chamado, usuario)}
                    </div>
                </div>`;

            $('#app').html(corpo);

            inicializarEditores(usuario);
            inicializarDropdownsPropriedades(chamado, usuario);
            inicializarAreaResposta(usuario);
            inicializarFiltroTimeline(chamado, usuario);
            inicializarUpload(usuario);
            aplicarFade();

            // carrega o .js do tipo e injeta o detalhamento (read-only)
            montarDetalhamento(chamado);
        });
    }

    function aplicarFade() {
        $('[data-fade]').each(function () {
            const $wrap = $(this);
            const limite = parseInt($wrap.data('fade-max'), 10) || 200;

            $wrap.next('.fade-toggle').remove();
            $wrap.css('max-height', 'none');
            const alturaReal = $wrap.prop('scrollHeight');

            if (alturaReal <= limite + 24) {
                $wrap.removeClass('is-collapsed').css('max-height', 'none');
                return;
            }

            $wrap.addClass('is-collapsed').css('max-height', limite + 'px');
            const $btn = $(`
            <button class="fade-toggle" type="button">
                <i class="bi bi-chevron-down"></i><span>Mostrar mais</span>
            </button>`);

            $wrap.after($btn);

            $btn.on('click', function () {
                const colapsado = $wrap.hasClass('is-collapsed');
                if (colapsado) {
                    $wrap.removeClass('is-collapsed').css('max-height', $wrap.prop('scrollHeight') + 'px');
                    $btn.html('<i class="bi bi-chevron-up"></i><span>Mostrar menos</span>');
                    
                    setTimeout(() => { 
                        if (!$wrap.hasClass('is-collapsed')) $wrap.css('max-height', 'none'); 
                    }, 320);
                } else {
                    $wrap.css('max-height', $wrap.prop('scrollHeight') + 'px');  
                    void $wrap[0].offsetHeight;                                  
                    $wrap.addClass('is-collapsed').css('max-height', limite + 'px');
                    $btn.html('<i class="bi bi-chevron-down"></i><span>Mostrar mais</span>');
                }
            });
        });
    }


    function exibirLoading(exibir) {  
        if (exibir) {
            $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
        } else {
            $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
        }
    }

    function obterChamado() {            
        const params = {
            acao: 'obterChamado',
            idChamado: id,
            token: token
        };

        exibirLoading(true);

        $.ajax({
            url: 'ChamadoRepositorio.asp?v=<%=Now()%>',
            type: 'GET',
            data: params,
            dataType: 'json',
            success: function(response) {
                if (!response.sucesso) {
                    chamado = null
                    usuario = null
                    renderizarTela(chamado, usuario);
                    return;
                }

                chamado = {...response.dados.chamado, anexosPendentes: []}
                usuario = {...response.dados.usuario}
                renderizarTela(chamado, usuario);
            },
            error: function(xhr, status, error) {
                console.error('Erro ao pesquisar:', error);
                Toast.error('Erro ao realizar a pesquisa');
            }, complete: function() {
                exibirLoading(false);
            }
        });
    }

    function carregarAtendentes(codPlanta) {
        const params = {
            acao: 'listar',
            codPlanta: codPlanta
        };

        return $.ajax({
            url: '../Atendente/AtendenteRepositorio.asp?v=<%=Now()%>',
            type: 'GET',
            data: params,
            dataType: 'json',
            success: function(response) {
                if (!response.sucesso) {
                    console.error("Erro ao carregar atendentes: " + response.mensagem);
                    return;
                }

                catalogos.atendentes = response.dados
            },
            error: function(xhr, status, error) {
                console.error('Erro ao carregar atendentes:', error);
            }
        });
    }

    function carregarDepartamentos(codPlanta) {
        const params = {
            acao: 'listar',
            codPlanta: codPlanta,
        };

        return $.ajax({
            url: '../Departamento/DepartamentoRepositorio.asp?v=<%=Now()%>',
            type: 'GET',
            data: params,
            dataType: 'json',
            success: function(response) {
                if (!response.sucesso) {
                    console.error("Erro ao carregar departamentos: " + response.mensagem);
                    return;
                }

                catalogos.departamentos = response.dados
            },
            error: function(xhr, status, error) {
                console.error('Erro ao carregar departamentos:', error);
            }
        });
    }

    function carregarCategorias(idTipo, codPlanta) {
        const params = {
            acao: 'listar',
            idTipo: idTipo,
            codPlanta: codPlanta
        };

        return $.ajax({
            url: '../Categoria/CategoriaRepositorio.asp?v=<%=Now()%>',
            type: 'GET',
            data: params,
            dataType: 'json',
            success: function(response) {
                if (!response.sucesso) {
                    console.error("Erro ao carregar categorias: " + response.mensagem);
                    return;
                }

                catalogos.categorias = response.dados
            },
            error: function(xhr, status, error) {
                console.error('Erro ao carregar categorias:', error);
            }
        });
    }
    
    function enviarMensagem() {
        const ehAtendente = usuario.tipo === 'ATENDENTE';
        const zona  = ehAtendente ? 'agent' : 'requester';
        const quill = ehAtendente ? agentQuill : requesterQuill;

        const modo = ehAtendente ? composerState.modo : 'publica';

        const html  = quill ? quill.root.innerHTML.trim() : '';
        const texto = quill ? quill.getText().trim() : '';
        const temAnexo = chamado.anexosPendentes.length > 0;

        if (!texto && !temAnexo) {
            Toast.error('Escreva uma mensagem ou anexe um arquivo antes de enviar.');
            return;
        }

        const enviando = Upload.pending[zona].some(u => u.status === 'uploading');
        if (enviando) {
            Toast.error('Aguarde o término do envio dos anexos.');
            return;
        }

        const dadosFormulario = {
            acao: 'enviarMensagem',
            idChamado: chamado.id,
            token: token,
            ehNotaInterna: modo === 'nota' ? 1 : 0,
            mensagem: html,
            acaoAdicional: $('#acaoAdicional').val(),
            anexos: JSON.stringify(chamado.anexosPendentes)
        };

        exibirLoading(true)

        $.ajax({
            url: 'ChamadoRepositorio.asp',
            type: 'POST',
            data: dadosFormulario,
            dataType: 'json',
            success: function(response) {
                if (!response.sucesso) {
                    Toast.error(response.mensagem);
                    return;
                }
                
                Toast.success(response.mensagem);
                setTimeout(() => { window.location.reload() }, 1000);
            },
            error: function(xhr, status, error) {
                Toast.error('Erro ao salvar formulário');
            }, 
            complete: function() {
                exibirLoading(false)
            }
        });
    }

    function alterarStatus(status) {
        const dadosFormulario = {
            acao: 'alterarStatus',
            idChamado: chamado.id,
            status: status
        };

        exibirLoading(true)

        $.ajax({
            url: 'ChamadoRepositorio.asp',
            type: 'POST',
            data: dadosFormulario,
            dataType: 'json',
            success: function(response) {
                if (!response.sucesso) {
                    Toast.error(response.mensagem);
                    return;
                }
                
                Toast.success(response.mensagem);
                setTimeout(() => { window.location.reload() }, 1000);
            },
            error: function(xhr, status, error) {
                Toast.error('Erro ao salvar formulário');
            }, 
            complete: function() {
                exibirLoading(false)
            }
        });
    }
    
    function alterarPrioridade(prioridade) {
        const dadosFormulario = {
            acao: 'alterarPrioridade',
            idChamado: chamado.id,
            prioridade: prioridade
        };

        exibirLoading(true)

        $.ajax({
            url: 'ChamadoRepositorio.asp',
            type: 'POST',
            data: dadosFormulario,
            dataType: 'json',
            success: function(response) {
                if (!response.sucesso) {
                    Toast.error(response.mensagem);
                    return;
                }
                
                Toast.success(response.mensagem);
                setTimeout(() => { window.location.reload() }, 1000);
            },
            error: function(xhr, status, error) {
                Toast.error('Erro ao salvar formulário');
            }, 
            complete: function() {
                exibirLoading(false)
            }
        });
    }

    function alterarDepartamento(idDepartamento) {
        const dadosFormulario = {
            acao: 'alterarDepartamento',
            idChamado: chamado.id,
            idDepartamento: idDepartamento
        };

        exibirLoading(true)

        $.ajax({
            url: 'ChamadoRepositorio.asp',
            type: 'POST',
            data: dadosFormulario,
            dataType: 'json',
            success: function(response) {
                if (!response.sucesso) {
                    Toast.error(response.mensagem);
                    return;
                }
                
                Toast.success(response.mensagem);
                setTimeout(() => { window.location.reload() }, 1000);
            },
            error: function(xhr, status, error) {
                Toast.error('Erro ao salvar formulário');
            }, 
            complete: function() {
                exibirLoading(false)
            }
        });
    }

    function alterarAtendente(idAtendente) {
        const dadosFormulario = {
            acao: 'atribuirAtendente',
            idChamado: chamado.id,
            idAtendente: idAtendente
        };

        exibirLoading(true)

        $.ajax({
            url: 'ChamadoRepositorio.asp',
            type: 'POST',
            data: dadosFormulario,
            dataType: 'json',
            success: function(response) {
                if (!response.sucesso) {
                    Toast.error(response.mensagem);
                    return;
                }
                
                Toast.success(response.mensagem);
                setTimeout(() => { window.location.reload() }, 1000);
            },
            error: function(xhr, status, error) {
                Toast.error('Erro ao salvar formulário');
            }, 
            complete: function() {
                exibirLoading(false)
            }
        });
    }

    function alterarCategoria(idCategoria, nomeCategoria) {
        const dadosFormulario = {
            acao: 'alterarCategoria',
            idChamado: chamado.id,
            idCategoria: idCategoria,
            nomeCategoria: nomeCategoria
        };

        exibirLoading(true)

        $.ajax({
            url: 'ChamadoRepositorio.asp',
            type: 'POST',
            data: dadosFormulario,
            dataType: 'json',
            success: function(response) {
                if (!response.sucesso) {
                    Toast.error(response.mensagem);
                    return;
                }
                
                Toast.success(response.mensagem);
                setTimeout(() => { window.location.reload() }, 1000);
            },
            error: function(xhr, status, error) {
                Toast.error('Erro ao salvar formulário');
            }, 
            complete: function() {
                exibirLoading(false)
            }
        });
    }

    function confirmacaoFinalizacao() {
        ConfirmationModal.warning({
            message: 'Deseja realmente finalizar esse chamado?',
            confirmText: 'Confirmar',
            cancelText: 'Cancelar',
            onConfirm: () => {
                finalizarChamado();
            },
            onCancel: () => {}
        });
    }
        
    function finalizarChamado() {      
        const params = {
            acao: 'finalizarChamado',
            idChamado: chamado.id
        };

        exibirLoading(true)

        $.ajax({
            url: 'ChamadoRepositorio.asp?v=<%=Now()%>',
            type: 'GET',
            data: params,
            dataType: 'json',
            success: function(response) {
                if (!response.sucesso) {
                    Toast.error(response.mensagem || 'Erro ao finalizar chamado');
                    return;
                }

                Toast.success(response.mensagem || 'Chamado Finalizado com sucesso!');
                setTimeout(() => { window.location.reload() }, 1000);
            },
            error: function(xhr, status, error) {
                console.error('Erro ao finalizar chamado:', error);
                Toast.error('Erro ao finalizar chamado');
            }, 
            complete: function() {
                exibirLoading(false)
            }
        });
    }

    function navegarParaGrade() {
        window.location.href = 'ChamadoFilaAtendimento.asp';
    }

    function navegarParaAvaliacao() {
        window.location.href = 'ChamadoAvaliacao.asp?token=' + chamado.tokenAcesso;
    }
        
    function confirmacaoReabertura() {
        ConfirmationModal.warning({
            message: 'Deseja realmente reabrir esse chamado?',
            confirmText: 'Confirmar',
            cancelText: 'Cancelar',
            onConfirm: () => {
                reabrirChamado();
            },
            onCancel: () => {}
        });
    }
        
    function reabrirChamado() {            
        const params = {
            acao: 'reabrirChamado',
            idChamado: chamado.id
        };

        exibirLoading(true)

        $.ajax({
            url: 'ChamadoRepositorio.asp?v=<%=Now()%>',
            type: 'GET',
            data: params,
            dataType: 'json',
            success: function(response) {
                if (!response.sucesso) {
                    Toast.error(response.mensagem || 'Erro ao reabrir chamado');
                    return;
                }

                Toast.success(response.mensagem || 'Chamado Reaberto com sucesso!');
                setTimeout(() => { window.location.reload() }, 1000);
            },
            error: function(xhr, status, error) {
                console.error('Erro ao reabrir chamado:', error);
                Toast.error('Erro ao reabrir chamado');
            }, 
            complete: function() {
                exibirLoading(false)
            }
        });
    }

    $(document).ready(function() {
        const urlParams = new URLSearchParams(window.location.search);
        id = urlParams.get('id') || '';
        token = urlParams.get('token') || '';
        
        obterChamado();
    });
</script>

</body>
</html>
