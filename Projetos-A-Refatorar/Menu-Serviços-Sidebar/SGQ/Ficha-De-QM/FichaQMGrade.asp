<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Smurfit Kappa - Ficha QM</title>
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

        .container {
            max-width: 90%;
        }
        
        .form-label {
            font-weight: 600;
        }

        .table-wrapper {
            overflow: auto;
            max-height: 70vh;
        }

        #gradeFicha {
            width: max-content;
            min-width: 100%;
        }

        #gradeFicha th,
        #gradeFicha td {
            white-space: nowrap;
            vertical-align: top;
        }

        #gradeFicha th:last-child,
        #gradeFicha td:last-child {
            position: sticky;
            right: 0;
            z-index: 5;
            white-space: nowrap;
        }
        
        #gradeFicha td:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px),#fff;
            background-origin: border-box;
            background-clip: padding-box;
        }
        
        #gradeFicha thead th:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px), #f5f5f5;
            background-origin: border-box;
            background-clip: padding-box;
            z-index: 10;
        }

        thead {
            vertical-align: middle !important;
            position: sticky; 
            inset-block-start: 0; 
            background: #f5f5f5;
            z-index: 10;
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
            max-width: 250px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            cursor: help;
        }

        .z-index-9999 {
            z-index: 9999 !important;
        }
        .badge {
            font-size: .7rem;
        }
        .badge-yellow {
            background-color: #fef3c7;
            color: #92400e;
        }
        .badge-green {
            background-color: #d1fae5;
            color: #065f46;
        }
        .badge-red {
            background-color: #fee2e2;
            color: #991b1b;
        }
        .badge-blue {
            background-color: #cffafe;
            color: #155e75;
        }
        .badge-purple {
            background-color: #e0e7ff;
            color: #3730a3;
        }
        .badge-gray {
            background-color: #e2e2e4;
            color: #374151;
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
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/dropdownMenu.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js"></script>

    <script>
        let id = 0;
        let dadosUsuario = null
        let listaFichas = []

        const listaPlantas = [
            { codigo: '00', descricao: 'Pirapetinga' },
            { codigo: '04', descricao: 'Uberaba' },
            { codigo: '05', descricao: 'Bento' },
            { codigo: '06', descricao: 'Maranguape' },
            { codigo: '07', descricao: 'Saquarema' },
            { codigo: '08', descricao: 'Araçatuba' },
            { codigo: '09', descricao: 'Blumenau' },
            { codigo: '10', descricao: 'Pacajus' },
            { codigo: '11', descricao: 'Porto Feliz' },
            { codigo: '12', descricao: 'Três Barras' },
            { codigo: '13', descricao: 'Florestal' }
        ]

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').show();
            } else {
                $('#loadingOverlay').hide();
            }
        }

        function carregarFichas() {            
            exibirLoading(true);

            const params = {
                acao: 'listar',
                id: $('#id').val(),
                idFNC: $('#idFNC').val(),
                codPlanta: $('#codPlanta').val(),
                dataCriacaoDe: $('#dataCriacaoDe').val(),
                dataCriacaoAte: $('#dataCriacaoAte').val(),
                status: $('#status').val(),
                NFDevolucao: $('#NFDevolucao').val(),
                ordemProducao: $('#ordemProducao').val(),
                nomeCliente: $('#nomeCliente').val(),
            };

            $.ajax({
                url: 'FichaQMRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                
                    listaFichas = response.dados || [];
                    renderizarGrade();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar fichas:', error);
                    Toast.error('Erro ao carregar fichas');
                },
                complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function renderizarGrade() {
            const $tbody = $('#gradeFicha tbody');
            
            $tbody.empty();
            
            if (listaFichas.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }

            listaFichas.forEach(function(item) {  
                let htmlAcoes = obterAcoes(item)
                
                $tbody.append(`
                    <tr>
                        <td>${item.id || '-'}</td>
                        <td>${obterNomePlanta(item.codPlanta)}</td>
                        <td>${item.nomeCliente || '-'}</td>
                        <td>${parseInt(item.idFNC) || '-'}</td>
                        <td>${obterBadgeStatus(item.status)}</td>
                        <td>${formatarMoeda(item.quantidadePaletes || 0, 0)}</td>
                        <td>${formatarMoeda(item.quantidadeRecebida || 0, 0)}</td>
                        <td>${formatarMoeda(item.quantidadeRefugada || 0, 0)}</td>
                        <td>${formatarMoeda(item.valor || 0)}</td>
                        <td>${formatarMoeda(item.valorUnitarioMaterial || 0)}</td>
                        <td>${item.NFDevolucao || '-'}</td>
                        <td>${item.codigoVA || '-'}</td>
                        <td>${item.codigoQM || '-'}</td>
                        <td>${item.modoDefeito || '-'}</td>
                        <td>${item.area || '-'}</td>
                        <td>${item.ordemProducao || '-'}</td>
                        <td>${item.dataEntradaDevolucao || '-'}</td>
                        <td>${item.dataConclusao || '-'}</td>
                        <td>${item.dataCriacao || '-'}</td>
                        ${htmlAcoes}
                    </tr>
                `);
            });
        }

        function obterAcoes(item) {
            const itens = [];

            if(dadosUsuario.nivel == 'RESPONSAVEL' && ['EXPEDICAO', 'CUSTOMER SERVICE'].includes(dadosUsuario.setor)) {
                itens.push(
                    `<a class="dropdown-item" onclick="editarFicha('${item.id}')"><i class="bi bi-pencil"></i> Editar</a>`,
                    `<a class="dropdown-item" onclick="excluirFicha('${item.id}')"><i class="bi bi bi-trash"></i> Excluir</a>`,
                )
            }

            return `
                <td style="text-align: center">
                    <dropdown-menu label="Ações">
                        ${itens.length > 0 ? itens.join('') : '<a class="dropdown-item"> Nenhuma ação disponível </a>'}
                    </dropdow-menu>
                </td>
            `;
        }
   
        function novaFicha() {
            window.location.href = `FichaQMFormulario.asp`;
        }

        function editarFicha(id) {
            window.location.href = `FichaQMFormulario.asp?id=${id}`;
        }

        function excluirFicha(id) {
            ConfirmationModal.danger({
                message: 'Deseja realmente excluir essa ficha?',
                type: 'danger',
                confirmText: 'Confirmar',
                cancelText: 'Cancelar',
                onConfirm: () => {
                    confirmarExclusao(id);
                },
                onCancel: () => {}
            });
        }

        function confirmarExclusao(id) {            
            exibirLoading(true);

            const params = {
                acao: 'excluir',
                id: id,
            };

            $.ajax({
                url: 'FichaQMRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    Toast.success(response.mensagem); 
                    carregarFichas();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao excluir:', error);
                    Toast.error('Erro ao excluir');
                }
            });
        }


        function limparFiltros() {
            $('#codPlanta').val('').trigger('change');
            $('#status').val('').trigger('change');
            $('#nomeCliente').val('');
            $('#dataCriacaoAte').val('');
            $('#dataCriacaoDe').val('');
            $('#NFDevolucao').val('');
            $('#ordemProducao').val('');
            $('#idFNC').val('');
            $('#id').val('');

            listaFichas = [];
            renderizarGrade();
        }

        function inicializarSelect2() {
            $('.select2').select2({ theme: 'bootstrap-5', width: '100%', dropdownCssClass: "z-index-9999", allowClear: true });
        }
 
        function obterDadosUsuario() {
            return new Promise((resolve, reject) => {
                $.ajax({
                    url: 'FichaQMRepositorio.asp?v=<%=Now()%>',
                    type: 'GET',
                    data: {
                        acao: 'obterDadosUsuario',
                    },
                    dataType: 'json',
                    success: function(response) {
                        exibirLoading(false);
                        resolve(response); 
                    },
                    error: function(xhr, status, error) {
                        exibirLoading(false);
                        reject(error); 
                    }
                });
            });
        }

        function popularPlantas(plantasString) {
            const plantasUsuario = plantasString.split(',').map(planta => planta.trim());
            
            const $select = $('#codPlanta');
            $select.empty().append('<option value="">Escolha uma planta</option>');

            const plantasVisiveis = listaPlantas.filter(planta => plantasUsuario.includes(planta.codigo));
            plantasVisiveis.forEach(function(item) {
                $select.append($('<option>', { value: item.codigo, text: item.descricao }));
            });
        }

        function obterDescricaoStatus(status) {
            const statusMap = {
                AGUARDANDO_RETRABALHO: 'Aguardando Retrabalho',
                AGUARDANDO_FATURAMENTO: 'Aguardando Faturamento',
                REFATURADO_TOTAL: 'Refaturado Total',
                REFUGADO_TOTAL: 'Refugado Total',
                RETRABALHO_FINALIZADO: 'Retrabalho Finalizado',
                SIMBOLICO_REFATURADO: 'Simbólico Refaturado',
                SIMBOLICO_REFUGADO: 'Simbólico Refugado'
            };
            return statusMap[status] || status || '-';
        }

        function obterBadgeStatus(status) {
            const badgeMap = {
                AGUARDANDO_RETRABALHO: 'badge-yellow',
                AGUARDANDO_FATURAMENTO: 'badge-yellow',
                REFATURADO_TOTAL: 'badge-green',
                REFUGADO_TOTAL: 'badge-red',
                RETRABALHO_FINALIZADO: 'badge-blue',
                SIMBOLICO_REFATURADO: 'badge-purple',
                SIMBOLICO_REFUGADO: 'badge-gray'
            };

            return `<span class="badge ${badgeMap[status] || 'badge-gray'}">${obterDescricaoStatus(status)}</span>`;
        }

        function exportarPDF() {
            if (listaFichas.length === 0) {
                alert('Não há dados para exportar');
                return;
            }

            const { jsPDF } = window.jspdf;
            var doc = new jsPDF({
                orientation: 'landscape', 
                unit: 'mm',
                format: 'a2'
            });

            const dataAtual = new Date().toLocaleDateString('pt-BR');

            const colunas = [
                '#', 'Planta', 'Cliente', 'Status', 'FNC', 'Quantidade Paletes', 'Quantidade Recebida', 
                'Quantidade Refugada', 'Valor', 'Valor Unit. Material', 'NF Devolução', 'Código VA', 
                'Código QM', 'Modo Defeito', 'Area', 
                'Ordem Produção', 'Data Entrada Devolução', 'Data Conclusão', 'Data Criação'
            ];

            const dados = listaFichas.map(item => [
                item.id || '-',
                obterNomePlanta(item.codPlanta),
                item.nomeCliente || '-',
                obterDescricaoStatus(item.status),
                parseInt(item.idFNC) || '-',
                formatarMoeda(item.quantidadePaletes || 0, 0),
                formatarMoeda(item.quantidadeRecebida || 0, 0),
                formatarMoeda(item.quantidadeRefugada || 0, 0),
                formatarMoeda(item.valor || 0),
                formatarMoeda(item.valorUnitarioMaterial || 0),
                item.NFDevolucao || '-',
                item.codigoVA || '-',
                item.codigoQM || '-',
                item.modoDefeito || '-',
                item.area || '-',
                item.ordemProducao || '-',
                item.dataEntradaDevolucao || '-',
                item.dataConclusao || '-',
                item.dataCriacao || '-'
            ]);

            doc.setFontSize(16);
            doc.text('Fichas de QM - Devoluções', 14, 15);
            
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
            if (listaFichas.length === 0) {
                alert('Não há dados para exportar');
                return;
            }

            const nomeArquivo = 'Fichas_QM.csv';

            const colunas = [
                '#', 'Planta', 'Cliente', 'Status', 'FNC', 'Quantidade Paletes', 'Quantidade Recebida', 
                'Quantidade Refugada', 'Valor', 'Valor Unit. Material', 'NF Devolução', 'Código VA', 
                'Código QM', 'Modo Defeito', 'Area', 
                'Ordem Produção', 'Data Entrada Devolução', 'Data Conclusão', 'Data Criação'
            ];

            let csvContent = '\uFEFF'; 
            csvContent += colunas.join(';') + '\r\n';

            listaFichas.forEach(item => {
                const linha = [
                    item.id || '-',
                    obterNomePlanta(item.codPlanta),
                    item.nomeCliente || '-',
                    obterDescricaoStatus(item.status),
                    parseInt(item.idFNC) || '-',
                    item.quantidadePaletes || 0,
                    item.quantidadeRecebida || 0,
                    item.quantidadeRefugada || 0,
                    item.valor || 0,
                    item.valorUnitarioMaterial || 0,
                    item.NFDevolucao || '-',
                    item.codigoVA || '-',
                    item.codigoQM || '-',
                    item.modoDefeito || '-',
                    item.area || '-',
                    item.ordemProducao || '-',
                    item.dataEntradaDevolucao || '-',
                    item.dataConclusao || '-',
                    item.dataCriacao || '-'
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

        function carregarStatus() {
            const $select = $('#status');

            const params = {
                acao: 'listarStatus'
            };

            $.ajax({
                url: 'FichaQMRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar status: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function (item) {
                        $select.append($('<option>', { value: item.codigo, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar status:', error);
                }
            });
        }
        
        $(document).ready(async function() {
            resposta = await obterDadosUsuario();
            if (!resposta.sucesso) {
                alert("Você não tem permissão para acessar esta página.");
                window.location.href = '/corporativo/frameset/principal.asp';
                return;
            }
        
            dadosUsuario = resposta.dados;
            if (!dadosUsuario || !dadosUsuario.nivel) {
                alert("Você não tem permissão para acessar esta página.");
                window.location.href = '/corporativo/frameset/principal.asp';
                return;
            }

            if (dadosUsuario.nivel == 'RESPONSAVEL' && dadosUsuario.setor == 'EXPEDICAO') {
                $('#btnNovaFicha').removeClass('d-none');
            }

            popularPlantas(dadosUsuario.plantas);
            renderizarGrade()
            carregarStatus();
            carregarFichas();
            inicializarSelect2();
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="d-flex my-4">
            <h3>Fichas de QM - Devoluções</h3>
        </div>

        <div class="card p-3 mb-4">
            <div class="row g-3">
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="codPlanta" class="form-label">Filial</label>
                    <select class="form-select form-select-sm select2" id="codPlanta" name="codPlanta" data-placeholder="Todas as filiais">
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
                    <label for="dataCriacaoDe" class="form-label">Data Criação De</label>
                    <input type="date" class="form-control form-control-sm" id="dataCriacaoDe" name="dataCriacaoDe">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="dataCriacaoAte" class="form-label">Data Criação Até</label>
                    <input type="date" class="form-control form-control-sm" id="dataCriacaoAte" name="dataCriacaoAte">
                </div>
                
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="nomeCliente" class="form-label">Cliente</label>
                    <input type="text" class="form-control form-control-sm" id="nomeCliente" name="nomeCliente" placeholder="Cliente">
                </div>
               
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="status" class="form-label">Status</label>
                    <select class="form-select form-select-sm select2" name="status" id="status" data-placeholder="Todos os status">
                        <option value="">Todas os status</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="id" class="form-label">Id Ficha</label>
                    <input type="number" class="form-control form-control-sm" id="id" name="id" placeholder="Id da Ficha">
                </div>
      
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="idFNC" class="form-label">Id FNC</label>
                    <input type="number" class="form-control form-control-sm" id="idFNC" name="idFNC" placeholder="Id da FNC">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="ordemProducao" class="form-label">Ordem de Produção</label>
                    <input type="text" class="form-control form-control-sm" id="ordemProducao" name="ordemProducao" placeholder="Ordem de produção">
                </div>
                
                <div class="col-12 col-md-4 col-lg-3">
                    <label for="NFDevolucao" class="form-label">NF Devolução</label>
                    <input type="text" class="form-control form-control-sm" id="NFDevolucao" name="NFDevolucao" placeholder="NF Devolução">
                </div>

                <div class="col-12 d-flex justify-content-end align-items-end gap-2 mt-4">
                    <button onclick="limparFiltros()" type="button" class="btn btn-sm btn-outline-secondary" id="btnLimpar">
                        <i class="bi bi-x-lg me-1"></i> Limpar
                    </button>
                    <button onclick="carregarFichas()" type="button" class="btn btn-sm btn-primary" id="btnPesquisar">
                        <i class="bi bi-search me-1"></i> Pesquisar
                    </button>
                </div>
            </div>
        </div>

        <div class="card">
            <div class="d-flex justify-content-between align-items-center p-3 border-bottom">
                <div class="d-flex gap-2">
                    <button id="btnNovaFicha" class="d-none btn btn-sm btn-primary" onclick="novaFicha()">
                        <i class="bi bi-plus-lg"></i>
                        Novo
                    </button>
                    <button type="button" class="btn btn-sm btn-success" onclick="exportarExcel()">
                        <i class="bi bi-file-earmark-excel me-1"></i> Exportar Excel
                    </button>
                    <button type="button" class="btn btn-sm btn-danger" onclick="exportarPDF()">
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

                <table class="table table-hover mb-0" id="gradeFicha">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Planta</th>
                            <th>Cliente</th>
                            <th>FNC</th>
                            <th>Status</th>
                            <th>Quantidade Paletes</th>
                            <th>Quantidade Recebida</th>
                            <th>Quantidade Refugada</th>
                            <th>Valor</th>
                            <th>Valor Unit. Material</th>
                            <th>NF Devolução</th>  
                            <th>Código VA</th> 
                            <th>Código QM</th>   
                            <th>Modo Defeito</th>
                            <th>Área</th>
                            <th>Ordem Produção</th>
                            <th>Data Entrada Devolução</th>    
                            <th>Data Conclusão</th>    
                            <th>Data Criação</th>      
                            <th style="width: 120px;" class="text-center">Ações</th>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td colspan="28" class="text-center text-muted py-4">
                                Nenhum registro encontrado
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</body>
</html>
