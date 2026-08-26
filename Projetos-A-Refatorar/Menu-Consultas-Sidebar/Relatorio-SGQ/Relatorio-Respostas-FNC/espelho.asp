<!DOCTYPE html>
<html lang="pt-br" xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Ficha de Ocorrências/Não Conformidades</title>
    <link rel="stylesheet" type="text/css" href="../../css/novo.css" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" />
    <script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
    <script src="../../funcoes/funcoesJAVASCRIPT.js"></script>

    <style type="text/css">
        :root {
            --cor-texto: #1f2933;
            --cor-borda: #4b5563;
            --cor-borda-suave: #9ca3af;
            --cor-cabecalho: #111827;
            --cor-secao: #e5e7eb;
            --cor-secao-texto: #111827;
            --cor-label: #374151;
            --cor-fundo: #f3f4f6;
            --cor-papel: #ffffff;
            --cor-botao: #198754;
        }

        * { box-sizing: border-box; }

        body {
            margin: 0;
            background: var(--cor-fundo);
            color: var(--cor-texto);
            font-family: Roboto;
        }

        .acoes-relatorio {
            width: 210mm;
            max-width: calc(100% - 24px);
            margin: 16px auto 8px auto;
            display: flex;
            justify-content: flex-end;
            gap: 8px;
            align-items: center;
        }

        #conteudo .acoes-relatorio {
            width: 100%;
            max-width: 100%;
            margin: 0 0 8px 0;
        }

        .btnImprimir {
            background-color: var(--cor-botao);
            color: white;
            border: none;
            padding: 8px 12px;
            border-radius: 5px;
            cursor: pointer;
            font-size: 13px;
            font-weight: 700;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }

        .btnImprimir:hover { filter: brightness(0.95); }

        .mensagem-carregando,
        .mensagem-erro,
        .mensagem-vazio {
            width: 210mm;
            max-width: calc(100% - 24px);
            margin: 24px auto;
            padding: 24px;
            text-align: center;
            background: #fff;
            border: 1px solid #ddd;
            border-radius: 6px;
        }

        .mensagem-erro { color: #dc3545; }

        #conteudo {
            min-height: 297mm;
            margin: 0 auto 20px auto;
            padding: 10mm;
            background: var(--cor-papel);
            box-shadow: 0 2px 12px rgba(15, 23, 42, .14);
        }

        .documento-topo {
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            margin-bottom: 6px;
            border: 1.5px solid var(--cor-borda);
        }

        .documento-topo td {
            border: 1px solid var(--cor-borda);
            padding: 7px 8px;
            vertical-align: middle;
        }

        .topo-logo { width: 28%; text-align: center; }
        .topo-titulo { width: 50%; text-align: center; }
        .topo-numero { width: 22%; text-align: center; }

        .logo_relatorio {
            max-width: 170px;
            max-height: 42px;
            width: auto;
            height: auto;
        }

        .titulo-documento {
            margin: 0;
            color: var(--cor-cabecalho);
            font-size: 15px;
            font-weight: 800;
            letter-spacing: .3px;
            
        }

        .subtitulo-documento {
            margin-top: 3px;
            font-size: 9px;
            color: #4b5563;
            letter-spacing: .2px;
        }

        .numero-documento-label {
            display: block;
            color: var(--cor-label);
            font-size: 9px;
            font-weight: 700;
            margin-bottom: 2px;
        }

        .numero-documento-valor {
            display: block;
            color: #000;
            font-size: 18px;
            font-weight: 800;
        }

        table.tabela-relatorio {
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            border: 1.5px solid var(--cor-borda);
        }

        .tabela-relatorio td {
            padding: 5px 6px;
            vertical-align: top;
            word-break: break-word;
            overflow-wrap: anywhere;
        }

        .table_cabecalho_destaque,
        .secao-titulo {
            background: var(--cor-secao) !important;
            color: var(--cor-secao-texto) !important;
            border: 1.5px solid var(--cor-borda) !important;
            padding: 5px 7px !important;
            font-weight: 800;
            letter-spacing: .25px;
            font-size: 12px
        }

        .borda { background: #fff; }

        .campo-label {
            display: block;
            color: var(--cor-label);
            font-size: 11px;
            font-weight: 800;
            letter-spacing: .15px;
            margin-bottom: 2px;
        }

        .campo-valor {
           display: block;
            min-height: 15px;
            color: #111827;
            font-size: 12px;
            font-weight: 500;
        }

        .campo-valor.texto-longo {
            min-height: 42px;
            white-space: pre-line;
        }


        .linha-subcabecalho td,
        td.subcabecalho {
            background: #f3f4f6 !important;
            color: #111827;
            font-weight: 800;
            
            font-size: 9px;
            text-align: center;
        }

        .valor-pre-line { white-space: pre-line; }
        .text-center { text-align: center; }
        .text-bold { font-weight: 800; }

        @media print {
            @page { size: A4 portrait; margin: 6mm; }

            html, body {
                width: 210mm;
                min-height: 297mm;
                background: #fff !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }

            body { font-size: 10.5px; }
            .no-print { display: none !important; }

            #conteudo {
                display: block !important;
                width: auto;
                min-height: auto;
                margin: 0;
                padding: 0;
                box-shadow: none;
                background: #fff;
            }

            .documento-topo,
            .tabela-relatorio { page-break-inside: auto; }

            tr, td { page-break-inside: avoid; }
            .table_cabecalho_destaque, .secao-titulo { page-break-after: avoid; }
        }
    </style>
</head>
<%
    Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
    Response.CodePage = 65001
    Response.CharSet = "UTF-8"
    Dim id
    id = Request.QueryString("id")
    If Trim(id & "") = "" Then id = "0"
%>
<body>
    <div id="mensagemCarregando" class="mensagem-carregando">Carregando dados da FNC...</div>
    <div id="mensagemErro" class="mensagem-erro" style="display:none;"></div>

    <div id="conteudo" style="display:none;">
        <div id="barraAcoes" class="acoes-relatorio no-print">
            <button onclick="window.print()" type="button" class="btnImprimir" id="btnExportarPDF">
                <i class="bi bi-file-earmark-pdf"></i> Imprimir / PDF
            </button>
        </div>

        <table class="documento-topo">
            <tr>
                <td class="topo-logo">
                    <img src="../../images/SW_LOGO_HORZ_BLACK_250x44.png" class="logo_relatorio" alt="Logo" />
                </td>
                <td class="topo-titulo">
                    <h1 class="titulo-documento">Ficha de Ocorrências/Não Conformidades</h1>
                    <div class="subtitulo-documento">Documento formal de registro e tratativa</div>
                </td>
                <td class="topo-numero">
                    <span class="numero-documento-label">Nº</span>
                    <span class="numero-documento-valor" data-campo="id"></span>
                </td>
            </tr>
        </table>

        <table class="tabela-relatorio">
            <tbody id="corpoRelatorio"></tbody>
        </table>
    </div>

    <script type="text/javascript">
        var id = '<%=Server.HTMLEncode(id)%>';
        var mostrarTodasRespostas = new URLSearchParams(window.location.search).get('mostrarTodasRespostas') === '1';

        $(document).ready(function() {
            carregarDadosFNC();
        });

        function carregarDadosFNC() {
            $.ajax({
                url: 'Principal/FNCRepositorio.asp',
                type: 'GET',
                data: {
                    acao: 'obterPorId',
                    id: id
                },
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        exibirErro(response.mensagem || 'Não existem informações para gerar este relatório.');
                        if (window.Toast && Toast.error) Toast.error(response.mensagem);
                        return;
                    }
                    preencherFormulario(response.dados || {});
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar dados para edição:', error);
                    exibirErro('Erro ao carregar dados para edição');
                    if (window.Toast && Toast.error) Toast.error('Erro ao carregar dados para edição');
                }
            });
        }

        function preencherFormulario(dados) {
            $('[data-campo="id"]').text(valor(dados.id));
            renderizarRelatorio(dados);
            $('#mensagemCarregando, #mensagemErro').hide();
            $('#conteudo').show();
        }

        function renderizarRelatorio(d) {
            var html = '';

            html += secao('Identificação');
            html += linha([
                celula('Data do Registro', formatarData(d.dataInclusao)),
                celula('Data de Ocorrência', formatarData(d.dataOcorrencia)),
                celula('Conclusão Prevista', formatarData(d.dataConclusaoPrevista)),
                celula('Conclusão Real', formatarData(d.dataEncerramento)),
                celula('Status', descricaoSituacao(d.situacao, d.status), 2),
                celula('Aprovação', descricaoAprovacao(d.aprovada))
            ]);
            html += linha([
                celula('Originador', d.nomeUsuarioInclusao || '', 3),
                celula('Responsável', d.nomeResponsavel || d.responsavel || '', 3),
                celula('Tipo', descricaoTipoFNC(d.tipo || d.tipoFNC))
            ]);

            html += secao('Dados do Cliente / Pedido');
            html += linha([
                celula('Cliente', juntarCodigoDescricao(d.codCliente, d.nomeCliente), 3),
                celula('Loja', d.codLoja || d.loja),
                celula('Município', d.cidadeEstado || d.municipio),
                celula('Pedido', listarPedidos(d.notas)),
                celula('Nota Fiscal', listarNotas(d.notas))
            ]);
            html += linha([celula('Contato', d.contato || '-', 7)]);

            html += renderizarAprovadores(d.aprovacoes || []);

            html += secao('Descrição Detalhada');
            html += linha([celulaValor(d.descricao || d.descricaoDefeito || '', 7, true)]);

            html += secao('Valores');
            html += linha([
                celula('Quantidade', d.quantidadeReclamada || d.quantidade, 3),
                celula('Valor Reclamado', formatarMoeda(d.valorReclamado), 4)
            ]);

            html += secao('Disposição');
            html += linha([celulaValor(d.clienteDisposicao || d.descricaoDisposicao || '', 7, true)]);

            html += secao('Verificação da Eficácia / Observação');
            html += linha([celulaValor(d.eficacia || '-', 7, true)]);

            html += secao('Análise');
            html += linha([
                celula('Disposição', juntarCodigoDescricao(d.disposicao), 3),
                celula('Cód. Origem', juntarCodigoDescricao(d.codigoOrigem, d.codigoOrigem), 4)
            ]);
            html += linha([celula('Efeito', d.descricaoTipoDefeito || d.descricaoDefeito || d.defeitoIshikawa, 7)]);
            html += linha([celula('Categoria FNC', d.descricaoTipoDefeito || d.descricaoDefeito || d.assunto5W2H, 7)]);

            html += renderizarCausasIshikawa(d.causasIshikawa || []);
            html += renderizarAcoes5W2H(d.acoes5W2H || []);
            html += renderizarRespostas(d.respostas || []);
            html += renderizarObservacoes(d.observacoes || []);

            $('#corpoRelatorio').html(html);
        }

        function renderizarAprovadores(aprovacoes) {
            var validas = $.grep(aprovacoes, function(a) { return String(a.invalido || '0') !== '1'; });
            if (!validas.length) return '';

            var html = secao('Aprovadores');
            html += linha([
                celulaTexto('<span class="campo-label text-center">Aprovador</span>', 2, 'subcabecalho'),
                celulaTexto('<span class="campo-label text-center">Ação</span>', 1, 'subcabecalho'),
                celulaTexto('<span class="campo-label text-center">Motivo</span>', 2, 'subcabecalho'),
                celulaTexto('<span class="campo-label text-center">Nível</span>', 1, 'subcabecalho'),
                celulaTexto('<span class="campo-label text-center">Data</span>', 1, 'subcabecalho')
            ]);
            $.each(validas, function(_, a) {
                html += linha([
                    celulaValor(a.nomeUsuarioAvaliacao || '-', 2),
                    celulaValor(descricaoAcaoAprovacao(a.avaliacao), 1),
                    celulaValor(limitarTexto(a.motivo || '-', 80), 2),
                    celulaValor(a.nivel, 1),
                    celulaValor(a.dataAvaliacao|| '-', 1)
                ]);
            });
            return html;
        }

        function renderizarCausasIshikawa(causas) {
            if (!causas.length) return '';
            var html = secao('Ishikawa');
            html += linha([
                celulaTexto('<span class="campo-label text-center">Categoria</span>', 2, 'subcabecalho'),
                celulaTexto('<span class="campo-label text-center">Causa</span>', 4, 'subcabecalho'),
                celulaTexto('<span class="campo-label text-center">Causa Raiz</span>', 1, 'subcabecalho')
            ]);
            $.each(causas, function(_, c) {
                html += linha([
                    celulaValor(c.categoria, 2),
                    celulaValor(c.descricao, 4, true),
                    celulaValor(descricaoSimNao(c.causaRaiz), 1)
                ]);
            });
            return html;
        }

        function renderizarAcoes5W2H(acoes) {
            if (!acoes.length) return '';
            var html = secao('Plano de Ação - 5W2H');
            $.each(acoes, function(indice, a) {
                html += linha([celulaTexto('Ação ' + (indice + 1), 7, 'secao-titulo')]);
                html += linha([celula('O que', a.oque || '-', 7)]);
                html += linha([
                    celula('Quem', a.quem || '-', 3),
                    celula('Quando', a.quando || '-', 2),
                    celula('Quanto', a.quanto || '-', 2)
                ]);
                html += linha([celula('Por que', a.porque || '-', 7)]);
                html += linha([celula('Onde', a.onde || '-', 7)]);
                html += linha([celula('Como', a.como || '-', 7)]);
            });
            return html;
        }

        function renderizarRespostas(respostas) {
            var finalizadas = $.grep(respostas, function(r) {
                return ehVerdadeiro(r.finalizada);
            });

            if (!mostrarTodasRespostas) {
                finalizadas = finalizadas.filter(item => String(item.avaliacao) === '1');
                finalizadas.splice(1)
            }

            if (!finalizadas.length) return '';

            var html = secao('Respostas / Plano de Ação');
            $.each(finalizadas, function(_, r) {
                html += linha([
                    celula('Usuário', r.nomeUsuarioFinalizacao || r.nomeUsuarioInclusao, 3),
                    celula('Data', r.dataInclusao || r.dataFinalizacao, 4)
                ]);
                html += linha([celula('Causa', r.descricao, 7)]);
                html += linha([celula('Plano de Ação', r.justificativa, 7)]);
            });
            return html;
        }

        function renderizarObservacoes(observacoes) {
            if (!observacoes.length) return '';

            var html = secao('Observações');
            $.each(observacoes, function(_, o) {
                html += linha([
                    celula('Usuário', o.loginUsuarioInclusao || o.nomeUsuarioInclusao || o.usuario, 3),
                    celula('Data', o.dataCriacao || o.dataInclusao, 4)
                ]);
                html += linha([celulaValor(o.descricao, 7, true)]);
            });
            return html;
        }

        function secao(titulo) {
            return '<tr><td colspan="7" class="table_cabecalho_destaque">' + escapar(titulo) + '</td></tr>';
        }

        function linha(celulas) { return '<tr>' + celulas.join('') + '</tr>'; }

        function celula(rotulo, conteudo, colspan) {
            return celulaTexto(
                '<span class="campo-label">' + escapar(rotulo) + '</span>' +
                '<span class="campo-valor">' + escapar(valor(conteudo)) + '&nbsp;</span>',
                colspan || 1,
                'borda'
            );
        }

        function celulaValor(conteudo, colspan, textoLongo) {
            return celulaTexto(
                '<span class="campo-valor valor-pre-line' + (textoLongo ? ' texto-longo' : '') + '">' + escapar(valor(conteudo)) + '&nbsp;</span>',
                colspan || 1,
                'borda'
            );
        }

        function celulaTexto(conteudo, colspan, classe, style) {
            return '<td colspan="' + (colspan || 1) + '" class="' + (classe || 'borda') + '"' + (style ? ' style="' + style + '"' : '') + '>' + conteudo + '</td>';
        }

        function valor(v) { return (v === null || v === undefined) ? '' : String(v); }
        function temValor(v) { return $.trim(valor(v)) !== ''; }

        function escapar(v) {
            return valor(v)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#039;');
        }

        function formatarData(data) {
            if (!temValor(data)) return '';
            var s = valor(data);
            if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
                var partes = s.substring(0, 10).split('-');
                return partes[2] + '/' + partes[1] + '/' + partes[0];
            }
            return s;
        }

        function descricaoSituacao(situacao, status) {
            switch (String(situacao || status || '')) {
                case '1': return 'Aberta';
                case '2': return 'Respondida';
                case '3': return 'Encerrada';
                default: return '';
            }
        }

        function descricaoAprovacao(aprovada) {
            if (String(aprovada) === '1') return 'Aprovada';
            if (String(aprovada) === '0' || aprovada === null || aprovada === undefined || aprovada === '') return 'Não Aprovada';
            return aprovada;
        }

        function descricaoSimNao(valorEntrada) {
            if (String(valorEntrada) === '1' || String(valorEntrada).toLowerCase() === 'true' || String(valorEntrada).toLowerCase() === 'sim') return 'Sim';
            if (String(valorEntrada) === '0' || String(valorEntrada).toLowerCase() === 'false' || String(valorEntrada).toLowerCase() === 'nao') return 'Não';
            return valor(valorEntrada);
        }

        function descricaoTipoFNC(tipo) {
            return 'Não conformidade existente';
        }

        function descricaoAcaoAprovacao(avaliacao) {
            if (String(avaliacao) === '1' || String(avaliacao).toLowerCase() === 'sim') return 'Aprovação';
            if (String(avaliacao) === '0' || String(avaliacao).toLowerCase() === 'nao') return 'Rejeição';
            return valor(avaliacao);
        }

        function ehVerdadeiro(v) {
            return String(v).toLowerCase() === 'true' || String(v) === '1' || String(v).toLowerCase() === 'sim';
        }

        function listarPedidos(notas) {
            return $.map(notas || [], function(n) { return n.numeroPedido; }).filter(Boolean).join(', ');
        }

        function listarNotas(notas) {
            return $.map(notas || [], function(n) { return n.notaFiscal; }).filter(Boolean).join(', ');
        }

        function juntarCodigoDescricao(codigo, descricao) {
            codigo = valor(codigo);
            descricao = valor(descricao);
            if (codigo && descricao && codigo !== descricao) return codigo + ' - ' + descricao;
            return codigo || descricao;
        }

        function limitarTexto(texto, tamanho) {
            texto = valor(texto);
            return texto.length > tamanho ? texto.substring(0, tamanho) : texto;
        }

        function exibirErro(mensagem) {
            $('#mensagemCarregando').hide();
            $('#conteudo').hide();
            $('#mensagemErro').text(mensagem || 'Não existem informações para gerar este relatório.').show();
        }
    </script>
</body>
</html>
