<!doctype html>
<html lang="pt-BR">
<head>
	<meta charset="UTF-8" />
	<meta name="viewport" content="width=device-width, initial-scale=1" />
	<title>Avaliação de atendimento — Ática Helpdesk</title>

	<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet" />
	<link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet" />
	<link rel="preconnect" href="https://fonts.googleapis.com" />
	<link
		href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=IBM+Plex+Sans:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap"
		rel="stylesheet"
	/>

	<style>
		:root {
			--paper: #f5f5f5;
			--ink: #1a1612;
			--ink-soft: #4a453e;
			--ink-mute: #8b857c;
			--line: rgba(26, 22, 18, 0.1);
			--line-soft: rgba(26, 22, 18, 0.06);
			--accent: #2d7bdf;
			--amber: #e9bd4b;
			--amber-soft: #f4ecdc;
			--rose: #b53a37;
			--sage: #5a7a4e;
			--sage-deep: #3a5530;
			--sage-soft: #e2eadc;
		}

		body {
			background-color: #f5f5f5;
            color: var(--ink);
			min-height: 100vh;
			background-size: 22px 22px, 13px 13px;
			background-position: 0 0, 7px 7px;
		}

		a { text-decoration: none; color: inherit; }
		button { background: transparent; border: 0; padding: 0; color: inherit; font: inherit; cursor: pointer; }
		::selection { background: rgba(196,80,47,0.22); }

		/* Tipografia */
		.text-mute { color: var(--ink-mute); }
		.text-soft { color: var(--ink-soft); }
		.text-accent { color: var(--accent); }
		.text-sage { color: var(--sage-deep); }
		.text-rose { color: var(--rose); }
		.text-amber { color: #7a5a12; }
		.uppercase-tight { text-transform: uppercase; letter-spacing: 0.18em; }
		.border-line { border-color: var(--line) !important; }
		.soft-shadow { box-shadow: 0 1px 0 rgba(26,22,18,0.04), 0 24px 48px -16px rgba(26,22,18,0.12); }

		/* Card principal */
		.survey-card {
            background: #fff;
            border-radius: 16px;
            border: 1px solid var(--line);
            overflow: hidden;
		}
		.survey-header {
            background: var(--ink);
            color: var(--paper);
            padding: 1.5rem 2rem;
            background-size: 14px 14px;
		}

		/* Estrelas CSAT */
		.star-row { display: flex; justify-content: center; gap: 12px; }
		.star-btn {
            background: transparent; border: 0; padding: 6px; cursor: pointer;
            font-size: 36px; color: rgba(26,22,18,0.2); transition: all .15s;
            line-height: 1;
		}
		.star-btn:hover { transform: scale(1.15); }
		.star-btn.active { color: var(--amber); }
		.star-btn.hover-preview { color: var(--amber); opacity: 0.5; }

		/* NPS 0-10 */
		.nps-grid { display: grid; grid-template-columns: repeat(11, 1fr); gap: 6px; }
		.nps-btn {
            aspect-ratio: 1;
            border-radius: 8px;
            border: 1px solid var(--line);
            background: #fff;
            font-weight: 600;
            font-size: 14px;
            color: var(--ink-soft);
            transition: all .15s;
            cursor: pointer;
            display: flex; align-items: center; justify-content: center;
		}
		.nps-btn:hover { border-color: var(--ink); transform: translateY(-2px); }
		.nps-btn.active {
            background: var(--ink); color: var(--paper); border-color: var(--ink);
            transform: translateY(-2px);
            box-shadow: 0 4px 12px -4px rgba(26,22,18,0.3);
		}
		.nps-scale-labels {
            display: flex; justify-content: space-between;
            font-size: 11px; color: var(--ink-mute); margin-top: 8px;
            letter-spacing: 0.04em;
		}

		/* Chips de motivos */
		.chip {
            display: inline-flex; align-items: center; gap: 6px;
            padding: 8px 14px;
            border-radius: 9999px;
            border: 1px solid var(--line);
            background: #fff;
            font-size: 13px;
            color: var(--ink-soft);
            cursor: pointer;
            transition: all .15s;
            user-select: none;
		}
		.chip:hover { border-color: var(--ink); color: var(--ink); }
		.chip.active {
		    background: var(--ink); color: var(--paper); border-color: var(--ink);
		}
		.chip i { color: var(--ink-mute); transition: color .15s; }
		.chip.active i { color: var(--paper); }

		/* Botões custom */
		.btn-c { display: inline-flex; align-items: center; gap: 8px; padding: 10px 16px; border-radius: 6px; font-size: 12px; font-weight: 500; transition: all .15s; border: 0; cursor: pointer; line-height: 1; }
		.btn-primary-c { background: var(--ink); color: var(--paper); }
		.btn-primary-c:hover { background: #2a2620; color: var(--paper); }
		.btn-primary-c:disabled { background: var(--line); color: var(--ink-mute); cursor: not-allowed; }
		.btn-ghost-c { background: transparent; color: var(--ink-mute); }
		.btn-ghost-c:hover { color: var(--ink); }
		.btn-outline-c {
            background: transparent; color: var(--ink);
            border: 1px solid var(--line);
		}
		.btn-outline-c:hover { background: var(--paper); }

		/* Feedback dinâmico */
		.rating-feedback {
            text-align: center;
            font-size: 18px;
            color: var(--ink);
            min-height: 28px;
            margin-top: 12px;
            opacity: 0;
            transition: opacity .25s;
		}
		.rating-feedback.show { opacity: 1; }

		/* Progress bar */
		.progress-track { height: 3px; background: var(--line); border-radius: 9999px; overflow: hidden; }
		.progress-fill { height: 100%; background: var(--accent); transition: width .35s cubic-bezier(.2,.7,.3,1); }

		/* Step transition */
		.step { display: none; }
		.step.active { display: block; animation: stepIn .3s ease-out; }
		@keyframes stepIn {
            from { opacity: 0; transform: translateY(8px); }
            to   { opacity: 1; transform: translateY(0); }
		}

		/* Mensagens de erro inline */
		.field-error {
            color: var(--rose); font-size: 12px; margin-top: 8px;
            display: none;
		}
		.field-error.show { display: block; }

		/* Tela de agradecimento */
		.thanks-icon {
            width: 88px; height: 88px;
            border-radius: 50%; background: var(--sage-deep); color: #fff;
            display: inline-flex; align-items: center; justify-content: center;
            font-size: 44px; margin-bottom: 1.5rem;
            animation: pop .5s cubic-bezier(.4,1.6,.5,1);
		}
		@keyframes pop {
            0% { transform: scale(0); }
            60% { transform: scale(1.1); }
            100% { transform: scale(1); }
		}

		/* Recap */
		.recap-card {
            background: var(--paper);
            border: 1px solid var(--line);
            border-radius: 12px;
            padding: 1.25rem;
		}
		.recap-row {
            display: flex; justify-content: space-between; align-items: center;
            padding: 8px 0;
		}
		.recap-row + .recap-row { border-top: 1px solid var(--line); }

		.avatar {
            display: inline-flex; align-items: center; justify-content: center;
            border-radius: 9999px;
            font-weight: 500; flex-shrink: 0;
		}
		.av-accent { background: var(--accent); color: var(--paper); }
		.av-32 { width: 32px; height: 32px; font-size: 12px; }

		.pill {
            display: inline-flex; align-items: center; gap: 6px;
            border-radius: 9999px; padding: 4px 10px;
            font-size: 11px; font-weight: 500;
		}
		.pill-sage { background: var(--sage-soft); color: var(--sage-deep); }

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

<div class="container py-4 py-md-5" style="max-width: 720px;">
    <div id="app"></div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jquery@3.7.1/dist/jquery.min.js"></script>
<script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
<script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
<script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

<script>
    const catalogos = {
        // Feedback exibido conforme nota 
        mensagensNota: {
            1: { texto: 'Sentimos muito que tenha sido uma experiência ruim.',  tom: 'negativo' },
            2: { texto: 'Algo não saiu como esperado. Vamos entender melhor.',  tom: 'negativo' },
            3: { texto: 'Obrigado — vamos buscar melhorar.',                    tom: 'neutro' },
            4: { texto: 'Ótimo, fico feliz que tenha sido uma boa experiência!', tom: 'positivo' },
            5: { texto: 'Excelente! Adoramos saber disso.',                     tom: 'positivo' }
        },

        // Motivos exibidos quando a nota é ≥ 4
        motivosPositivos: [
            { id: 'rapidez',       icon: 'bi-lightning-fill', label: 'Rapidez na resposta' },
            { id: 'cordialidade',  icon: 'bi-emoji-smile',    label: 'Atendimento cordial' },
            { id: 'resolvido',     icon: 'bi-check2-circle',  label: 'Problema resolvido' },
            { id: 'comunicacao',   icon: 'bi-chat-dots',      label: 'Comunicação clara' },
            { id: 'solucao',       icon: 'bi-lightbulb',      label: 'Solução criativa' },
            { id: 'tecnico',       icon: 'bi-tools',          label: 'Conhecimento técnico' }
        ],

        // Motivos exibidos quando a nota é ≤ 3
        motivosNegativos: [
            { id: 'demora',          icon: 'bi-hourglass-split', label: 'Demorou para responder' },
            { id: 'incompreendido',  icon: 'bi-question-circle', label: 'Não entendi a solução' },
            { id: 'nao_resolvido',   icon: 'bi-x-circle',        label: 'Problema não resolvido' },
            { id: 'repetir',         icon: 'bi-arrow-repeat',    label: 'Tive que repetir o problema' },
            { id: 'pouco_cordial',   icon: 'bi-emoji-frown',     label: 'Atendimento pouco cordial' },
            { id: 'solucao_lenta',   icon: 'bi-clock-history',   label: 'Solução demorou muito' }
        ],

        // Configuração de cada passo do questionário
        passos: [
            { id: 1, nome: 'Sua nota',      progresso: 33 },
            { id: 2, nome: 'Detalhes',      progresso: 66 },
            { id: 3, nome: 'Recomendação',  progresso: 100 }
        ],

        // Mensagens da tela final, condicionadas à nota
        mensagensAgradecimento: {
            positivo: 'Sua avaliação foi enviada para {{atendente}} e para o time de qualidade. Continuamos trabalhando para melhorar ainda mais.',
            neutro:   'Sua avaliação foi recebida. Vamos analisar com cuidado para que sua próxima experiência seja melhor.',
            negativo: 'Sua avaliação foi recebida. Nosso time de qualidade vai analisar com cuidado para que sua próxima experiência seja melhor.',
            pulado:   'Tudo bem! Se mudar de ideia, o link da avaliação fica disponível por 7 dias.'
        }
    };

    const avaliacao = {
        chamado: {},

        estado: {
            passo: 1,
            nota: 0,              
            motivos: [],         
            comentario: '',
            nps: null,        
            pulado: false
        }
    };

    const esc = s => $('<div>').text(String(s)).html();

    function obterIniciais(nome) {
        if (!nome) return '?';
        const partes = String(nome).trim().split(/\s+/).filter(Boolean);
        if (partes.length === 0) return '?';
        if (partes.length === 1) return partes[0].slice(0, 2).toUpperCase();
        return (partes[0][0] + partes[partes.length - 1][0]).toUpperCase();
    }

    function interpolar(template, valores) {
        return String(template).replace(/\{\{(\w+)\}\}/g, (_, k) => esc(valores[k] || ''));
    }

    function motivosPorNota(nota) {
        return nota >= 4 ? catalogos.motivosPositivos : catalogos.motivosNegativos;
    }

    function tomDaNota(nota) {
        if (!nota) return null;
        return catalogos.mensagensNota[nota].tom;
    }

    /**
     * Encontra a configuração de um passo pelo id.
     */
    function buscarPasso(id) {
        return catalogos.passos.find(p => p.id === id);
    }

    const Templates = {
        avatar(nome, cls, size) {
            return `<span class="avatar ${cls || 'av-accent'} av-${size || 32}">${esc(obterIniciais(nome))}</span>`;
        },

        pill(label, cls, icon) {
            const ic = icon ? `<i class="bi ${icon}"></i>` : '';
            return `<span class="pill ${cls}">${ic} ${esc(label)}</span>`;
        },

        estrela(rating, ativa) {
            return `<button type="button" class="star-btn ${ativa ? 'active' : ''}" data-rating="${rating}" aria-label="${rating} estrelas">
            <i class="bi bi-star-fill"></i>
            </button>`;
        },

        chip(motivo, ativo) {
            return `<button type="button" class="chip ${ativo ? 'active' : ''}" data-motivo="${esc(motivo.id)}">
            <i class="bi ${motivo.icon}"></i> ${esc(motivo.label)}
            </button>`;
        },

        npsBtn(valor, ativo) {
            return `<button type="button" class="nps-btn ${ativo ? 'active' : ''}" data-nps="${valor}">${valor}</button>`;
        },

        recapRow(label, valor, mute) {
            return `
            <div class="recap-row">
                <span class="text-soft" style="font-size:14px;">${esc(label)}</span>
                <span class="${mute ? 'text-mute' : ''}" style="font-weight:500;font-size:14px;text-align:right;">${valor}</span>
            </div>`;
        }
    };


    function renderizarCabecalho(avaliacao) {
        const c = avaliacao.chamado;

        return `
            <div class="survey-header">
                <div class="d-flex justify-content-between align-items-start gap-3 flex-wrap">
                    <div>
                        <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.18em;color:var(--ink-mute);">
                            Avaliação · Chamado <span class="font-mono">#${esc(c.id)}</span>
                        </div>
                        <h1 class="font-display mt-2 mb-0" style="font-size:22px;line-height:1.3;">
                            ${esc(c.titulo)}
                        </h1>
                    </div>
                    <span class="pill" style="background:rgba(90,122,78,0.2);color:#a8c89a;">
                        <i class="bi bi-check-circle-fill"></i> ${esc(c.status)}
                    </span>
                </div>
                <div class="mt-3 d-flex align-items-center gap-3 flex-wrap" style="font-size:12px;color:var(--ink-mute);">
                    <span class="d-inline-flex align-items-center gap-1">
                        <i class="bi bi-person-circle"></i>
                        Atendido por <strong style="color:var(--paper);font-weight:500;">${esc(c.nomeAtendente)}</strong>
                    </span>
                    <span style="opacity:0.4;">·</span>
                    <span class="d-inline-flex align-items-center gap-1">
                        <i class="bi bi-clock"></i>
                        Resolvido em <strong style="color:var(--paper);font-weight:500;">${esc(formatarParaDataBR(c.dataFinalizacao))}</strong>
                    </span>
                    <span style="opacity:0.4;">·</span>
                </div>
            </div>`;
    }

    function renderizarProgresso(estado) {
        if (estado.passo > 3) return '';  
        const passo = buscarPasso(estado.passo);

        return `
            <div class="px-4 pt-3">
                <div class="progress-track">
                    <div id="progress-fill" class="progress-fill" style="width: ${passo.progresso}%;"></div>
                </div>
                <div class="d-flex justify-content-between mt-2" style="font-size:11px;color:var(--ink-mute);font-weight:500;letter-spacing:0.04em;">
                    <span>Passo <span>${passo.id}</span> de ${catalogos.passos.length}</span>
                    <span>${esc(passo.nome)}</span>
                </div>
            </div>`;
    }

    function renderizarPasso1(estado) {
        const estrelas = [1, 2, 3, 4, 5]
            .map(n => Templates.estrela(n, n <= estado.nota))
            .join('');

        const feedback = estado.nota
            ? `<div class="rating-feedback show">${esc(catalogos.mensagensNota[estado.nota].texto)}</div>`
            : `<div class="rating-feedback">&nbsp;</div>`;

        return `
            <div id="step-1" class="step active">
                <h2 class="font-display fw-bold mb-2" style="font-size:28px;">Como foi sua experiência?</h2>
                <p class="text-soft mb-4" style="font-size:15px;">
                    Sua opinião nos ajuda a melhorar o suporte. Leva menos de 1 minuto.
                </p>

                <div class="my-4">
                    <div class="star-row" id="star-row">${estrelas}</div>
                        ${feedback}
                    <div id="error-rating" class="field-error text-center">
                        Por favor, selecione uma nota antes de continuar.
                    </div>
                </div>

                <div class="d-flex justify-content-between align-items-center mt-5">
                    <button class="btn-c btn-ghost-c" id="skip-link">Pular avaliação</button>
                    <button class="btn-c btn-primary-c" id="btn-next-1">
                        Continuar <i class="bi bi-arrow-right"></i>
                    </button>
                </div>
            </div>`;
    }

    function renderizarPasso2(estado) {
        const ehPositivo = estado.nota >= 4;
        const tituloVar = ehPositivo ? 'mais contribuiu' : 'precisamos melhorar';

        const chips = motivosPorNota(estado.nota)
            .map(m => Templates.chip(m, estado.motivos.includes(m.id)))
            .join('');

        return `
            <div id="step-2" class="step active">
                <h2 class="font-display fw-bold mb-2" style="font-size:26px;">
                    O que <span>${esc(tituloVar)}</span> para essa nota?
                </h2>
                <p class="text-soft mb-4" style="font-size:14px;">
                    Escolha um ou mais motivos. Você pode selecionar quantos quiser.
                </p>

                <div class="d-flex flex-wrap gap-2 mb-4" id="reasons">${chips}</div>

                <label class="d-block mb-2" style="font-size:14px;font-weight:500;color:var(--ink-soft);">
                    Tem algum comentário adicional? <span class="text-mute" style="font-weight:400;">(opcional)</span>
                </label>

                <textarea
                    class="form-control"
                    id="comment"
                    rows="4"
                    placeholder="Conte com mais detalhes o que achou do atendimento..."
                    maxlength="500"
                >${esc(estado.comentario)}</textarea>
                
                <div class="d-flex justify-content-end" style="font-size:11px;color:var(--ink-mute);margin-top:6px;">
                    <span id="comment-count">${estado.comentario.length}</span>/500
                </div>

                <div class="d-flex justify-content-between align-items-center mt-4">
                    <button class="btn-c btn-ghost-c" id="btn-back-2">
                        <i class="bi bi-arrow-left"></i> Voltar
                    </button>
                    <button class="btn-c btn-primary-c" id="btn-next-2">
                        Continuar <i class="bi bi-arrow-right"></i>
                    </button>
                </div>
            </div>`;
    }

    function renderizarPasso3(estado) {
        const nps = Array.from({ length: 11 }, (_, i) =>
            Templates.npsBtn(i, estado.nps === i)
        ).join('');

        // Monta o recap
        const ratingRecap = estado.nota
            ? '<span class="text-amber">' + '★'.repeat(estado.nota) + '</span>'
            + '<span class="text-mute">' + '★'.repeat(5 - estado.nota) + '</span>'
            + ' <span class="text-mute"> · ' + estado.nota + '/5</span>'
            : '—';

        const motivosLabels = estado.motivos
            .map(id => motivosPorNota(estado.nota).find(m => m.id === id))
            .filter(Boolean)
            .map(m => m.label);
        const motivosRecap = motivosLabels.length ? motivosLabels.join(', ') : 'Nenhum';

        const comentarioRecap = estado.comentario.trim()
            ? '"' + esc(estado.comentario.trim().slice(0, 60)) + (estado.comentario.length > 60 ? '...' : '') + '"'
            : 'Nenhum';

        return `
            <div id="step-3" class="step active">
                <h2 class="font-display fw-bold mb-2" style="font-size:26px;">
                    O quanto você recomendaria nosso suporte?
                </h2>
                <p class="text-soft mb-4" style="font-size:14px;">
                    Considerando essa experiência, daria nota de 0 a 10 para indicar nosso suporte a outra pessoa?
                </p>

                <div class="nps-grid mb-1" id="nps-grid">${nps}</div>
                <div class="nps-scale-labels">
                    <span>Não recomendaria</span>
                    <span>Recomendaria muito</span>
                </div>
                <div id="error-nps" class="field-error">
                    Por favor, selecione uma nota antes de enviar.
                </div>

                <div class="recap-card mt-4">
                    <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.15em;color:var(--ink-mute);font-weight:600;margin-bottom:8px;">
                        Resumo da sua avaliação
                    </div>
                    ${Templates.recapRow('Nota geral',           ratingRecap)}
                    ${Templates.recapRow('Motivos selecionados', esc(motivosRecap), motivosLabels.length === 0)}
                    ${Templates.recapRow('Comentário',           comentarioRecap, !estado.comentario.trim())}
                </div>

                <div class="d-flex justify-content-between align-items-center mt-4">
                    <button class="btn-c btn-ghost-c" id="btn-back-3">
                        <i class="bi bi-arrow-left"></i> Voltar
                    </button>
                    <button class="btn-c btn-primary-c" id="btn-submit">
                        <i class="bi bi-send-fill"></i> Enviar avaliação
                    </button>
                </div>
            </div>`;
    }

    function renderizarPasso4(avaliacao) {
        const estado = avaliacao.estado;

        let chaveMensagem;
        if (estado.pulado) chaveMensagem = 'pulado';
        else if (!estado.nota) chaveMensagem = 'neutro';
        else chaveMensagem = tomDaNota(estado.nota);

        const mensagem = interpolar(
            catalogos.mensagensAgradecimento[chaveMensagem],
            { atendente: avaliacao.chamado.nomeAtendente }
        );

        return `
            <div id="step-4" class="step active text-center py-3">
                <div class="thanks-icon"><i class="bi bi-check-lg"></i></div>
                <h2 class="font-display fw-bold mb-2" style="font-size:30px;">
                    Obrigado pelo feedback, ${esc(avaliacao.chamado.solicitante.nome.split(' ')[0])}!
                </h2>
                <p class="text-soft mb-4" style="font-size:15px;max-width:480px;margin:0 auto;">
                    ${mensagem}
                </p>

                <div class="d-flex justify-content-center gap-2 flex-wrap mt-4">
                    <button class="btn-c btn-outline-c" id="btn-reopen">
                        <i class="bi bi-arrow-counterclockwise"></i> Reabrir este chamado
                    </button>
                </div>
            </div>`;
    }

    function renderizarConteudo(avaliacao) {
        const estado = avaliacao.estado;
        if (estado.passo === 1) return renderizarPasso1(estado);
        if (estado.passo === 2) return renderizarPasso2(estado);
        if (estado.passo === 3) return renderizarPasso3(estado);
        if (estado.passo === 4) return renderizarPasso4(avaliacao);
        return '';
    }

    function renderizarRodape() {
        return `
            <p class="text-center mt-4" style="font-size:12px;color:var(--ink-mute);">
                Suas respostas são confidenciais e usadas apenas para melhorar nosso atendimento.
            </p>`;
    }


    function inicializarPasso1() {
        $('#star-row').on('mouseover', '.star-btn', function () {
            const r = parseInt($(this).data('rating'), 10);

            $('.star-btn').each(function () {
                const v = parseInt($(this).data('rating'), 10);
                $(this).toggleClass('hover-preview', v <= r && !$(this).hasClass('active'));
            });
        });

        $('#star-row').on('mouseleave', function () {
            $('.star-btn').removeClass('hover-preview');
        });

        // Clique em estrela
        $('#star-row').on('click', '.star-btn', function () {
            const r = parseInt($(this).data('rating'), 10);
            avaliacao.estado.nota = r;
            $('#error-rating').removeClass('show');
            rerenderizar();  // re-renderiza pra atualizar visual e mensagem
        });

        $('#btn-next-1').on('click', () => {
            if (!avaliacao.estado.nota) {
                $('#error-rating').addClass('show');
                return;
            }
            avaliacao.estado.passo = 2;
            avaliacao.estado.motivos = []; // reset ao avançar (motivos dependem da nota)
            rerenderizar();
        });

        $('#skip-link').on('click', e => {
            e.preventDefault();
            Object.assign(avaliacao.estado, {
                pulado: true, nota: 0, motivos: [], comentario: '', nps: null, passo: 4
            });
            rerenderizar();
        });
    }

    function inicializarPasso2() {
        $('#reasons').on('click', '.chip', function () {
            const id = $(this).data('motivo');
            const ativo = $(this).hasClass('active');
            if (ativo) {
                avaliacao.estado.motivos = avaliacao.estado.motivos.filter(x => x !== id);
            } else {
                avaliacao.estado.motivos.push(id);
            }

            $(this).toggleClass('active');
        });

        $('#comment').on('input', function () {
            avaliacao.estado.comentario = $(this).val();
            $('#comment-count').text(avaliacao.estado.comentario.length);
        });

        $('#btn-back-2').on('click', () => {
            avaliacao.estado.passo = 1;
            rerenderizar();
        });

        $('#btn-next-2').on('click', () => {
            avaliacao.estado.passo = 3;
            rerenderizar();
        });
    }

    function inicializarPasso3() {
        $('#nps-grid').on('click', '.nps-btn', function () {
            avaliacao.estado.nps = parseInt($(this).data('nps'), 10);
            $('.nps-btn').removeClass('active');
            $(this).addClass('active');
            $('#error-nps').removeClass('show');
        });

        $('#btn-back-3').on('click', () => {
            avaliacao.estado.passo = 2;
            rerenderizar();
        });

        $('#btn-submit').on('click', () => {
            if (avaliacao.estado.nps === null) {
                $('#error-nps').addClass('show');
                return;
            }

            avaliarChamado();
        });
    }

    function avaliarChamado() {
        const dadosFormulario = {
            acao: 'avaliarChamado',
            idChamado: avaliacao.chamado.id,
            token: token,
            notaAtendimento: avaliacao.estado.nota,
            notaRecomendacao: avaliacao.estado.nps,
            motivos: avaliacao.estado.motivos.join(','),
            comentario: avaliacao.estado.comentario,
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
                
                avaliacao.estado.passo = 4;
                rerenderizar();
            },
            error: function(xhr, status, error) {
                Toast.error('Erro ao salvar formulário');
                avaliacao.estado.passo = 4;
                rerenderizar();
            }, 
            complete: function() {
                exibirLoading(false)
            }
        });
    }

    function inicializarPasso4() {
        $('#btn-reopen').on('click', () => {
           confirmacaoReabertura();
        });
    }

    function inicializarPasso(passo) {
        if (passo === 1) inicializarPasso1();
        if (passo === 2) inicializarPasso2();
        if (passo === 3) inicializarPasso3();
        if (passo === 4) inicializarPasso4();
    }

    function renderizarTela(avaliacao) {
         if (!avaliacao) {
            $('#app').html(`
                <div class="error-message">
                    <i class="bi bi-exclamation-triangle"></i>
                    <h4>Não autorizado</h4>
                    <p>Voce nao tem permissao para visualizar esse chamado</p>
                </div>
            `)

            return;
        }

        const html = `
            <div class="survey-card soft-shadow">
                ${renderizarCabecalho(avaliacao)}
                ${renderizarProgresso(avaliacao.estado)}
                <div class="p-4 p-md-5">
                    ${renderizarConteudo(avaliacao)}
                </div>
            </div>
            ${renderizarRodape()}
        `;

        $('#app').html(html);
        inicializarPasso(avaliacao.estado.passo);
    }

    function rerenderizar() {
        renderizarTela(avaliacao);
        $('html, body').animate({ scrollTop: 0 }, 200);
    }

    
    function exibirLoading(exibir) {  
        if (exibir) {
            $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
        } else {
            $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
        }
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
            idChamado: avaliacao.chamado.id
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
                setTimeout(() => { window.location.href= '../Chamado/ChamadoDetalhes.asp?token=' + token }, 1000);
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
    
    function obterChamado() {            
        const params = {
            acao: 'obterChamado',
            idChamado: id,
            token: token
        };

        exibirLoading(true);

        $.ajax({
            url: '../Chamado/ChamadoRepositorio.asp?v=<%=Now()%>',
            type: 'GET',
            data: params,
            dataType: 'json',
            success: function(response) {
                if (!response.sucesso) {
                    renderizarTela(avaliacao);
                    return;
                }

                avaliacao.chamado = response.dados.chamado
                renderizarTela(avaliacao);
            },
            error: function(xhr, status, error) {
                console.error('Erro ao pesquisar:', error);
                Toast.error('Erro ao realizar a pesquisa');
            }, complete: function() {
                exibirLoading(false);
            }
        });
    }

    $(document).ready(function() {
        const urlParams = new URLSearchParams(window.location.search);
        id = urlParams.get('id') || '';
        token = urlParams.get('token') || '';

        obterChamado()
        renderizarTela(avaliacao);
    });
</script>

</body>
</html>
