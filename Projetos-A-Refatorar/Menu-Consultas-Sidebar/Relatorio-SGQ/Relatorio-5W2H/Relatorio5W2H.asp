<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Smurfit Kappa - Relatório de Planos de Ação (5W2H)</title>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous" />
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css" />
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <link href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" rel="stylesheet" />
    
    <style>
        body {
            background-color: #f5f5f5;
            font-size: .8rem;
            padding-bottom: 20px;
        }

        .form-container {
            background: #fff;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        .form-label {
            font-weight: 600;
        }

        .required-field {
            color: #dc3545;
        }

        .filtros-card {
            background-color: white;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }

        .card {
            background: white;
            border-radius: 8px;
            padding: 0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            border: none;
        }

        .table-wrapper {
            overflow: auto;
            max-height: 60vh;
        }

        
        #gradeFNC {
            width: max-content;
            min-width: 100%;
        }

        thead {
            vertical-align: middle !important;
            position: sticky; 
            inset-block-start: 0; 
            background: #f5f5f5;
        }

        td, th {
            padding: 8px 12px !important;
        }

        th {
            font-weight: 600;
            color: #495057;
        }

        tr:hover {
            background: #f5f5f5;
        }

        .no-records {
            padding: 40px;
            text-align: center;
            background: white;
            color: #6c757d;
        }

         .truncate-text {
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            max-width: 360px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: normal;
            word-break: break-word;
        }
        .texto-detalhe-cell {
            min-width: 260px;
            vertical-align: top;
        }
        .texto-detalhe-link {
            cursor: pointer;
            transition: color .15s ease;
        }
        .texto-detalhe-link:hover,
        .texto-detalhe-link:focus {
            color: #0d6efd;
            outline: none;
            text-decoration: underline;
        }
        .detalhe-conteudo {
            background: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 8px;
            max-height: 60vh;
            overflow: auto;
            padding: 14px;
            white-space: pre-wrap;
            word-break: break-word;
        }

        .z-index-9999 {
            z-index: 9999 !important;
        }

        .loading-overlay {
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(255,255,255,0.8);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 1000;
        }

        .badge-aberta {
            background-color: #d1fae5;
            color: #065f46;
        }

        .badge-respondida {
            background-color: #dbeafe;
            color: #1e40af;
        }

        .badge-encerrada {
            background-color: #fce7f3;
            color: #9f1239;
        }

        .container {
            max-width: 100%;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js"></script>

    <script>
        let dadosFNC = [];
        let carregando = false;

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                dropdownCssClass: "z-index-9999",
                allowClear: true,
                placeholder: function() {
                    return $(this).data('placeholder');
                }
            });
        }

        function carregarDepartamentos() {
            const $select = $('#filtroDepartamento');
            $select.empty().append('<option value="">Todos os departamentos</option>');

            const params = {
                acao: 'listarDepartamentos'
            };

            $.ajax({
                url: 'Relatorio5W2HRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar departamentos: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar departamentos:', error);
                }
            });
        }

        function carregarDefeitos() {
            const $select = $('#filtroDefeito');
            $select.empty().append('<option value="">Todos os defeitos</option>');

            const params = {
                acao: 'listarDefeitos'
            };

            $.ajax({
                url: 'Relatorio5W2HRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar defeitos: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar defeitos:', error);
                }
            });
        }

        function pesquisarCliente() {
            const codPlanta = $('#filtroFilial').val();
            const baseUrl = `${location.protocol}//${location.host}`;
            const url = `${baseUrl}/corporativo/buscaCLIENTE.asp?id1=filtroCodCliente&id2=filtroCodLoja&id3=nomeCliente&id6=${codPlanta}`;
            window.open(url, '', "width=650,height=400");
        }

        function validarFiltros() {
            if (!$('#filtroFilial').val()) {
                alert("Selecione uma filial");
                return false;
            }

            return true;
        }

        function exibirLoading(exibir) {
            carregando = exibir;
            if (exibir) {
                $('#loadingOverlay').show();
            } else {
                $('#loadingOverlay').hide();
            }
        }

        function pesquisar() {
            if (!validarFiltros()) return;
            
            exibirLoading(true);

            const params = {
                acao: 'listarAcoes',
                filtroFilial: $('#filtroFilial').val(),
                filtroDepartamento: $('#filtroDepartamento').val(),
                filtroIdFNC: $('#filtroIdFNC').val(),
                filtroSituacao: $('#filtroSituacao').val(),
                filtroDataInclusaoDe: $('#filtroDataInclusaoDe').val(),
                filtroDataInclusaoAte: $('#filtroDataInclusaoAte').val(),
                filtroDefeito: $('#filtroDefeito').val(),
                filtroCodCliente: $('#filtroCodCliente').val(),
                filtroCodLoja: $('#filtroCodLoja').val()
            };

            $.ajax({
                url: 'Relatorio5W2HRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        alert('Erro ao pesquisar: ' + response.mensagem);
                        return;
                    }

                    dadosFNC = response.dados || [];
                    renderizarGrade();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao pesquisar:', error);
                    alert('Erro ao realizar a pesquisa');
                }
            });
        }

        function renderizarGrade() {
            const $tbody = $('#gradeFNC tbody');
            
            $tbody.empty();
            
            if (dadosFNC.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="29" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }

            dadosFNC.forEach(function(item) {
                const situacaoBadge = obterBadgeSituacao(item.situacao);
                const baseUrl = `${location.protocol}//${location.host}`;
                const urlFNC = `${baseUrl}/corporativo/sgq/fnc/espelho.asp?id=${item.idFNC}&filial=${item.filial}`;
                
                $tbody.append(`
                    <tr>
                        <td>
                            <a href="${urlFNC}" target="_blank">
                                <i class="bi bi-file-earmark-text"></i>
                                ${item.idFNC}
                            </a>
                        </td>
                        <td>${item.dataCriacao || '-'}</td>
                        <td>${(item.codigoCliente + ' - ' + item.nomeCliente)}</td>
                        <td>${item.setor || '-'}</td>
                        <td>${item.descricaoTipoDefeito}</td>
                        <td>${situacaoBadge}</td>
                        ${criarCelulaDetalhe(item.oque, 'O quê')}
                        ${criarCelulaDetalhe(item.porque, 'Porquê')}
                        ${criarCelulaDetalhe(item.onde, 'Onde')}
                        ${criarCelulaDetalhe(item.quando, 'Quando')}
                        ${criarCelulaDetalhe(item.quem, 'Quem')}
                        ${criarCelulaDetalhe(item.quanto, 'Quanto')}
                        ${criarCelulaDetalhe(item.como, 'Como')}
                    </tr>
                `);
            });
        }

        function obterDescricaoSituacao(situacao) {
            switch(situacao) {
                case '1':
                    return 'ABERTA';
                case '2':
                    return 'RESPONDIDA';
                case '3':
                    return 'ENCERRADA';
                default:
                    return '';
            }
        }

        function obterBadgeSituacao(situacao) {
            switch(situacao) {
                case '1':
                    return '<span class="badge badge-aberta">ABERTA</span>';
                case '2':
                    return '<span class="badge badge-respondida">RESPONDIDA</span>';
                case '3':
                    return '<span class="badge badge-encerrada">ENCERRADA</span>';
                default:
                    return '';
            }
        }

        function limparFiltros() {
            $('#filtroFilial').val('').trigger('change');
            $('#filtroDepartamento').val('').trigger('change');
            $('#filtroIdFNC').val('');
            $('#filtroSituacao').val('').trigger('change');
            $('#filtroDataInclusaoDe').val('');
            $('#filtroDataInclusaoAte').val('');
            $('#filtroDefeito').val('').trigger('change');
            $('#filtroCodCliente').val('');
            $('#filtroCodLoja').val('');
            
            dadosFNC = [];
            renderizarGrade();
        }

        function exportarPDF() {
            if (dadosFNC.length === 0) {
                alert('Não há dados para exportar');
                return;
            }

            const { jsPDF } = window.jspdf;
            const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a2' });
            const dataAtual = new Date().toLocaleDateString('pt-BR');

            const colunas = ['FNC', 'Data Criação', 'Cliente', 'Setor', 'Defeito', 'O quê', 'Por quê', 'Onde', 'Quando', 'Quem', 'Quanto', 'Como'];
            const dados = dadosFNC.map(item => [
                item.idFNC || '-',
                item.dataCriacao || '-',
                (item.codigoCliente + ' - ' + item.nomeCliente),
                item.setor || '-',
                item.descricaoTipoDefeito || '-',
                item.oque || '-',
                item.porque || '-',
                item.onde || '-',
                item.quando || '-',
                item.quem || '-',
                item.quanto || '-',
                item.como || '-',
            ]);

            doc.setFontSize(16);
            doc.text('Relatório de Respostas de FNC', 14, 15);
            
            doc.setFontSize(10);
            doc.text(`Data: ${dataAtual} | ${dados.length} registro(s) encontrado(s)`, 14, 22);

            doc.autoTable({
                head: [colunas],
                body: dados,
                startY: 28,
                theme: 'striped',
                styles: { 
                    fontSize: 8,
                    cellPadding: 2
                },
                headStyles: { 
                    fillColor: [211, 211, 212],
                    textColor: [0, 0, 0],
                    fontStyle: 'bold'
                },
                columnStyles: {
                    0: { cellWidth: 25 },
                    1: { cellWidth: 'auto' },
                    2: { cellWidth: 50 },
                    3: { cellWidth: 50 }
                },
                margin: { top: 28 }
            });
                            
            window.open(doc.output('bloburl'), '_blank');
        }

        function exportarExcel() {
            if (dadosFNC.length === 0) {
                alert('Não há dados para exportar');
                return;
            }

            const nomeArquivo = 'Relatorio_Respostas_FNC.csv';
            const colunas = ['FNC', 'Data Criação', 'Cliente', 'Setor', 'Defeito', 'O quê', 'Por quê', 'Onde', 'Quando', 'Quem', 'Quanto', 'Como'];

            let csvContent = '\uFEFF'; 
            csvContent += colunas.join(';') + '\r\n';

            dadosFNC.forEach(item => {
                const linha = [
                    item.idFNC || '-',
                    item.dataCriacao || '-',
                    (item.codigoCliente + ' - ' + item.nomeCliente),
                    item.setor || '-',
                    item.descricaoTipoDefeito || '-',
                    item.oque || '-',
                    item.porque || '-',
                    item.onde || '-',
                    item.quando || '-',
                    item.quem || '-',
                    item.quanto || '-',
                    item.como || '-',
                ];
                csvContent += linha.join(';') + '\r\n';
            });

            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            
            if (navigator.msSaveBlob) { 
                navigator.msSaveBlob(blob, nomeArquivo);
            } else {
                const url = URL.createObjectURL(blob);
                link.href = url;
                link.download = nomeArquivo;
                link.style.display = 'none';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
            }
        }

           function escapeHtml(valor) {
            return String(valor || '')
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;');
        }

        function normalizarTextoDetalhe(valor) {
            const texto = String(valor || '').trim();
            return texto || '-';
        }

        function criarCelulaDetalhe(valor, titulo) {
            const texto = normalizarTextoDetalhe(valor);
            const textoSeguro = escapeHtml(texto);
            const tituloSeguro = escapeHtml(titulo || 'Detalhes');

            if (texto === '-') {
                return '<td class="texto-detalhe-cell text-muted">-</td>';
            }

            return `
                <td class="texto-detalhe-cell">
                    <div class="truncate-text texto-detalhe-link"
                         role="button"
                         tabindex="0"
                         onclick="abrirModalDetalhe(this)"
                         data-titulo="${tituloSeguro}"
                         data-detalhe="${textoSeguro}">
                        ${textoSeguro}
                    </div>
                </td>
            `;
        }

        function abrirModalDetalhe(botao) {
            const titulo = $(botao).attr('data-titulo') || 'Detalhes';
            const detalhe = $(botao).attr('data-detalhe') || '-';

            $('#modalDetalheTituloTexto').text(titulo);
            $('#modalDetalheConteudo').text(detalhe);

            const modal = new bootstrap.Modal(document.getElementById('modalDetalheTexto'));
            modal.show();
        }

        function escapeCSV(valor) {
            if (!valor) return '""';
            valor = valor.replace(/[\r\n]+/g, ' ').trim();
            valor = valor.replace(/"/g, '""');
            return '"' + valor + '"';
        }
    

        function anexarEventos() {
            $('#btnPesquisar').on('click', pesquisar);
            $('#btnLimpar').on('click', limparFiltros);
            $('#btnExportarPDF').on('click', exportarPDF);
            $('#btnExportarExcel').on('click', exportarExcel);
        }

        $(document).ready(function() {
            inicializarSelect2();
            carregarDepartamentos();
            carregarDefeitos();
            anexarEventos();
            renderizarGrade();
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <h3 class="mb-4">Relatório de Planos de Ação (5W2H)</h3>

        <div class="filtros-card">
            <div class="row g-3">
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroFilial" class="form-label">Filial</label>
                    <select class="form-select form-select-sm select2" id="filtroFilial" name="filtroFilial" data-placeholder="Todas as filiais">
                        <option value="">Todas as filiais</option>
                        <option value="00">Pirapetinga</option>
                        <option value="04">Uberaba</option>
                        <option value="05">Bento</option>
                        <option value="06">Maranguape</option>
                        <option value="07">Saquarema</option>
                        <option value="08">Araçatuba</option>
                        <option value="09">Blumenau</option>
                        <option value="10">Pacajus</option>
                        <option value="11">Porto Feliz</option>
                        <option value="12">Três Barras</option>
                        <option value="13">Florestal</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroCodCliente" class="form-label">Cód. Cliente</label>
                    <div class="input-group input-group-sm">
                        <button type="button" class="btn btn-sm btn-secondary" onclick="pesquisarCliente()">
                            Buscar <i class="bi bi-box-arrow-up-right"></i>
                        </button>
                        <input type="text" class="form-control form-control-sm" id="filtroCodCliente" name="filtroCodCliente" placeholder="Código do cliente">
                    </div>
                    <small class="text-muted">Clique no botão para buscar ou digite o código</small>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroCodLoja" class="form-label">Cód. Loja</label>
                    <input type="text" class="form-control form-control-sm" id="filtroCodLoja" name="filtroCodLoja" placeholder="Código da loja">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroDepartamento" class="form-label">Departamento</label>
                    <select class="form-select form-select-sm select2" id="filtroDepartamento" name="filtroDepartamento" data-placeholder="Todos os departamentos">
                        <option value="">Todos os departamentos</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroIdFNC" class="form-label">Id FNC</label>
                    <input type="number" class="form-control form-control-sm" id="filtroIdFNC" name="filtroIdFNC" placeholder="Número da FNC">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroSituacao" class="form-label">Situação</label>
                    <select class="form-select form-select-sm select2" id="filtroSituacao" name="filtroSituacao" data-placeholder="Todas">
                        <option value="">Todas</option>
                        <option value="1">ABERTA</option>
                        <option value="2">RESPONDIDA</option>
                        <option value="3">ENCERRADA</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroDataInclusaoDe" class="form-label">Data Inclusão De</label>
                    <input type="date" class="form-control form-control-sm" id="filtroDataInclusaoDe" name="filtroDataInclusaoDe">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroDataInclusaoAte" class="form-label">Data Inclusão Até</label>
                    <input type="date" class="form-control form-control-sm" id="filtroDataInclusaoAte" name="filtroDataInclusaoAte">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroDefeito" class="form-label">Defeito</label>
                    <select class="form-select form-select-sm select2" id="filtroDefeito" name="filtroDefeito" data-placeholder="Todos os defeitos">
                        <option value="">Todos os defeitos</option>
                    </select>
                </div>

                <div class="col-12 d-flex justify-content-end align-items-end gap-2">
                    <button type="button" class="btn btn-sm btn-outline-secondary" id="btnLimpar">
                        <i class="bi bi-x-lg me-1"></i> Limpar
                    </button>
                    <button type="button" class="btn btn-sm btn-primary" id="btnPesquisar">
                        <i class="bi bi-search me-1"></i> Pesquisar
                    </button>
                </div>
            </div>
        </div>

        <div class="card">
            <div class="d-flex justify-content-between align-items-center p-3 border-bottom">
                <div class="d-flex gap-2">
                    <button type="button" class="btn btn-success btn-sm" id="btnExportarExcel">
                        <i class="bi bi-file-earmark-excel me-1"></i> Exportar Excel
                    </button>
                    <button type="button" class="btn btn-danger btn-sm" id="btnExportarPDF">
                        <i class="bi bi-file-earmark-pdf me-1"></i> Exportar PDF
                    </button>
                </div>
            </div>

            <div class="table-wrapper">
                <div id="loadingOverlay" class="loading-overlay" style="display: none;">
                    <div class="spinner-border text-primary" role="status">
                        <span class="visually-hidden">Carregando...</span>
                    </div>
                </div>

                <table class="table table-hover mb-0" id="gradeFNC">
                    <thead>
                        <tr>
                            <th>FNC</th>
                            <th>Data Criação</th>
                            <th>Cliente</th>
                            <th>Setor</th>
                            <th>Defeito</th>
                            <th>Situação</th>
                            <th>O quê</th>
                            <th>Por quê</th>
                            <th>Onde</th>
                            <th>Quando</th>
                            <th>Quem</th>
                            <th>Quanto</th>
                            <th>Como</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td colspan="4" class="text-center text-muted py-4">
                                Utilize os filtros acima para pesquisar
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <div class="modal fade" id="modalDetalheTexto" tabindex="-1" aria-labelledby="modalDetalheTituloTexto" aria-hidden="true">
        <div class="modal-dialog modal-lg modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalDetalheTituloTexto">
                        Detalhes
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div id="modalDetalheConteudo" class="detalhe-conteudo"></div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Fechar</button>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
