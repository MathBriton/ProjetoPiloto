<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Smurfit Kappa - Dashboard de FNC</title>
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
            max-width: 100%;
        }

        .form-label {
            font-weight: 600;
        }

        .chart-container {
            position: relative;
            height: 280px;
        }
        
        .icon-wrapper {
            width: 48px;
            height: 48px;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 22px;
        }
        
        .icon-wrapper.purple {
            background-color: #ede9fe;
            color: #8b5cf6;
        }
        
        .icon-wrapper.yellow {
            background-color: #fef9c3;
            color: #ca8a04;
        }
        
        .icon-wrapper.green {
            background-color: #d1fae5;
            color: #10b981;
        }
        
        .icon-wrapper.blue {
            background-color: #dbeafe;
            color: #3b82f6;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
    <script src="../../../components/toast.js?v=<%=Now()%>"></script>

    <script>
        let dadosDashboard = null;
        let carregando = false;
        let charts = {};

        const coresGrafico = [
            'rgba(13, 110, 253, 0.55)',      
            'rgba(25, 135, 84, 0.55)',       
            'rgba(253, 126, 20, 0.55)',      
            'rgba(220, 53, 69, 0.55)',       
            'rgba(111, 66, 193, 0.55)',     
            'rgba(32, 201, 151, 0.55)',      
            'rgba(13, 202, 240, 0.55)',     
            'rgba(255, 193, 7, 0.55)',      
            'rgba(226, 7, 255, 0.55)',      
            'rgba(214, 51, 132, 0.55)',     
            'rgba(73, 80, 87, 0.55)',        
            'rgba(173, 181, 189, 0.55)',     
            'rgba(255, 87, 51, 0.55)',      
            'rgba(156, 39, 176, 0.55)',     
            'rgba(0, 150, 136, 0.55)',      
            'rgba(63, 81, 181, 0.55)',    
            'rgba(255, 152, 0, 0.55)',      
            'rgba(121, 85, 72, 0.55)',     
            'rgba(96, 125, 139, 0.55)',     
            'rgba(233, 30, 99, 0.55)',   
            'rgba(76, 175, 80, 0.55)',   
            'rgba(255, 235, 59, 0.55)',    
            'rgba(103, 58, 183, 0.55)',    
            'rgba(0, 188, 212, 0.55)'      
        ];

        const bordasGrafico = [
            'rgba(10, 87, 202, 1.0)',      
            'rgba(20, 108, 67, 1.0)',      
            'rgba(202, 101, 16, 1.0)',      
            'rgba(176, 42, 55, 1.0)',      
            'rgba(89, 53, 154, 1.0)',      
            'rgba(26, 161, 121, 1.0)',      
            'rgba(10, 162, 192, 1.0)',      
            'rgba(204, 154, 6, 1.0)',      
            'rgba(226, 7, 255, 1.0)',      
            'rgba(171, 41, 106, 1.0)',      
            'rgba(58, 64, 70, 1.0)',        
            'rgba(138, 145, 151, 1.0)',    
            'rgba(204, 70, 41, 1.0)',      
            'rgba(125, 31, 141, 1.0)',      
            'rgba(0, 120, 109, 1.0)',      
            'rgba(50, 65, 145, 1.0)',       
            'rgba(204, 122, 0, 1.0)',      
            'rgba(97, 68, 58, 1.0)',      
            'rgba(77, 100, 111, 1.0)',    
            'rgba(186, 24, 79, 1.0)',     
            'rgba(61, 140, 64, 1.0)',    
            'rgba(204, 188, 47, 1.0)',     
            'rgba(82, 46, 146, 1.0)',     
            'rgba(0, 150, 170, 1.0)'     
        ];

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
       
        function carregarSetores() {
            const $select = $('#filtroSetor');
            $select.empty().append('<option value="">Todos os setores</option>');

            const params = {
                acao: 'listarSetores'
            };

            $.ajax({
                url: 'DashboardFNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar setores: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar setores:', error);
                }
            });
        }

        function carregarDefeitos() {
            const $select = $('#filtroDefeito');
            $select.empty().append('<option value="">Todos os defeitos</option>');

            const params = {
                acao: 'listarDefeitos',
                filtroTipoDefeito: $('#filtroTipoDefeito').val()
            };

            $.ajax({
                url: 'DashboardFNCRepositorio.asp?v=<%=Now()%>',
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
            if (
                (!$('#filtroDataInclusaoDe').val() || !$('#filtroDataInclusaoAte').val())
                && 
                (!$('#filtroDataEnvioClienteDe').val() || !$('#filtroDataEnvioClienteAte').val())) 
            {
                Toast.error("Preencha a data de inclusão ou envio ao cliente");
                return false;
            }
            return true;
        }


        function exibirLoading(exibir) {
            carregando = exibir;
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function pesquisar() {
            if (!validarFiltros()) return;
            
            exibirLoading(true);

            const params = {
                acao: 'obterDadosDashboard',
                filtroFilial: $('#filtroFilial').val(),
                filtroDataInclusaoDe: $('#filtroDataInclusaoDe').val(),
                filtroDataInclusaoAte: $('#filtroDataInclusaoAte').val(),
                filtroDataEnvioClienteDe: $('#filtroDataEnvioClienteDe').val(),
                filtroDataEnvioClienteAte: $('#filtroDataEnvioClienteAte').val(),
                filtroCodCliente: $('#filtroCodCliente').val(),
                filtroCodLoja: $('#filtroCodLoja').val(),
                filtroSetor: $('#filtroSetor').val()?.join(','),
                filtroSituacao: $('#filtroSituacao').val()?.join(','),
                filtroValidacao: $('#filtroValidacao').val()?.join(','),
                filtroDefeito: $('#filtroDefeito').val()?.join(',')
            };

            $.ajax({
                url: 'DashboardFNCRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }

                    dadosDashboard = response.dados || [];
                    controlarExibicaoAviso();
                    renderizarDashboard();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao pesquisar:', error);
                    Toast.error('Erro ao realizar a pesquisa');
                }
            });
        }

        function destruirCharts() {
            Object.keys(charts).forEach(key => {
                if (charts[key]) {
                    charts[key].destroy();
                    charts[key] = null;
                }
            });
        }

        function formatarDadosSimples(dados, horizontal) {
            const labels = dados.map(item => item.label);
            const totails = dados.map(item => item.total);

            datasets = [{
                label: 'Total',
                data: totails,
                backgroundColor: coresGrafico.slice(0, totails.length),
                borderColor: bordasGrafico.slice(0, totails.length),
                borderWidth: 0.55,
                borderRadius: 1,
                barThickness: horizontal ? 20 : 'flex',
                maxBarThickness: horizontal ? 24 : 50
            }];

            return { labels, datasets };
        }

        function formatarDadosAninhado(dados, horizontal) {
            const labels = dados.map(item => item.label);
            const seriesNomes = [...new Set(
                dados.flatMap(item => item.total.map(v => v.label))
            )];

            datasets = seriesNomes.map((serie, index) => ({
                label: serie,
                data: dados.map(item => {
                    const encontrado = item.total.find(v => v.label === serie);
                    return encontrado ? encontrado.total : 0;
                }),
                backgroundColor: coresGrafico[index % coresGrafico.length],
                borderColor: bordasGrafico[index % bordasGrafico.length],
                borderWidth: 0.55,
                borderRadius: 1,
                barThickness: horizontal ? 20 : 'flex',
                maxBarThickness: horizontal ? 24 : 50
            }));

            return { labels, datasets };
        }

        function criarGraficoBarras(canvasId, dados, horizontal = false, stacked = false) {
            const ctx = document.getElementById(canvasId);
            if (!ctx) return null;

            const dadosFormatados = stacked ? 
                formatarDadosAninhado(dados, horizontal) : formatarDadosSimples(dados, horizontal);

            const config = {
                type: 'bar',
                data: dadosFormatados,
                options: {
                    indexAxis: horizontal ? 'y' : 'x',
                    responsive: true,
                    maintainAspectRatio: false,
                    onClick: (event, activeElements) => {
                        if (activeElements.length > 0) {
                            const index = activeElements[0].index;
                            const itemValor = dados.find((item, i) => i === index).valor;
                            
                            navegarParaRelatorio(itemValor, canvasId); 
                        }
                    },
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            backgroundColor: '#212529',
                            titleFont: { size: 12 },
                            bodyFont: { size: 12 },
                            padding: 10,
                            cornerRadius: 2
                        }
                    },
                    scales: {
                        x: {
                            stacked: stacked,
                            grid: {
                                display: !horizontal,
                                color: '#e9ecef'
                            },
                            ticks: {
                                font: { size: 11 },
                                color: '#6c757d'
                            }
                        },
                        y: {
                            stacked: stacked,
                            grid: {
                                display: horizontal,
                                color: '#e9ecef'
                            },
                            ticks: {
                                font: { size: 11 },
                                color: '#6c757d'
                            },
                            beginAtZero: true
                        }
                    }
                }
            };

            return new Chart(ctx, config);
        }

        function criarGraficoPizza(canvasId, dados) {
            const ctx = document.getElementById(canvasId);
            if (!ctx) return null;

            const dadosFormatados = formatarDadosSimples(dados)
            const config = {
                type: 'doughnut',
                data: dadosFormatados,
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    onClick: (event, activeElements) => {
                        if (activeElements.length > 0) {
                            const index = activeElements[0].index;
                            const itemValor = dados.find((item, i) => i === index).valor;
                            
                            navegarParaRelatorio(itemValor, canvasId); 
                        }
                    },
                    plugins: {
                        legend: {
                            position: 'right',
                            labels: {
                                font: { size: 11 },
                                color: '#6c757d',
                                padding: 12,
                                usePointStyle: true,
                                pointStyle: 'circle'
                            }
                        },
                        tooltip: {
                            backgroundColor: '#212529',
                            titleFont: { size: 12 },
                            bodyFont: { size: 12 },
                            padding: 10,
                            cornerRadius: 4
                        }
                    },
                    cutout: '60%'
                }
            };

            return new Chart(ctx, config);
        }

        
        function navegarParaRelatorio(itemValor, canvasId) {
            // Ao clicar em um item o proprio item será enviado como filtro
            const codFiltro = {
                chartFilial: 'filtroFilial',
                chartCliente: 'filtroCodCliente',
                chartDefeito: 'filtroDefeito',
                chartSetor: 'filtroSetor'
            };
            
            const filtroClique = `${codFiltro[canvasId]}=${itemValor}`;

            let url = `../RelatorioFNC/RelatorioFNC.asp?${filtroClique}`;
      
            if ($('#filtroFilial').val()) url += `&filtroFilial=${$('#filtroFilial').val()}`;
            if ($('#filtroDataInclusaoDe').val()) url += `&filtroDataInclusaoDe=${$('#filtroDataInclusaoDe').val()}`;
            if ($('#filtroDataInclusaoAte').val()) url += `&filtroDataInclusaoAte=${$('#filtroDataInclusaoAte').val()}`;
            if ($('#filtroDataEnvioClienteDe').val()) url += `&filtroDataEnvioClienteDe=${$('#filtroDataEnvioClienteDe').val()}`;
            if ($('#filtroDataEnvioClienteAte').val()) url += `&filtroDataEnvioClienteAte=${$('#filtroDataEnvioClienteAte').val()}`;
            if ($('#filtroCodCliente').val()) url += `&filtroCodCliente=${$('#filtroCodCliente').val()}`;
            if ($('#filtroCodLoja').val()) url += `&filtroCodLoja=${$('#filtroCodLoja').val()}`;
            if ($('#filtroSetor').val()) url += `&filtroSetor=${$('#filtroSetor').val().join(',')}`;
            if ($('#filtroSituacao').val()) url += `&filtroSituacao=${$('#filtroSituacao').val().join(',')}`;
            if ($('#filtroValidacao').val()) url += `&filtroValidacao=${$('#filtroValidacao').val().join(',')}`;
            if ($('#filtroDefeito').val()) url += `&filtroDefeito=${$('#filtroDefeito').val().join(',')}`;

            window.open(url, '_blank');
        }

        function renderizarDashboard() {
            destruirCharts();

            if (!dadosDashboard) {
                $('#kpiTotal').text('0');
                $('#kpiAbertas').text('0');
                $('#kpiRespondidas').text('0');
                $('#kpiEncerradas').text('0');
                
                $('#dashboardContainer').html(`
                    <div class="col-12">
                        <div class="card shadow-sm border-0">
                            <div class="card-body text-center py-5 text-muted">
                                <i class="bi bi-bar-chart-line display-4 mb-3 opacity-50"></i>
                                <p class="mb-0">Utilize os filtros acima e clique em Pesquisar para visualizar o dashboard</p>
                            </div>
                        </div>
                    </div>
                `);
                return;
            }

            $('#kpiTotal').text(dadosDashboard.totalFNC.toLocaleString('pt-BR'));
            $('#kpiAbertas').text(dadosDashboard.totalAbertas.toLocaleString('pt-BR'));
            $('#kpiRespondidas').text(dadosDashboard.totalRespondidas.toLocaleString('pt-BR'));
            $('#kpiEncerradas').text(dadosDashboard.totalEncerradas.toLocaleString('pt-BR'));

            $('#dashboardContainer').html(`
                <div class="col-12 col-lg-6">
                    <div class="card shadow-sm border-0 h-100">
                        <div class="d-flex justify-content-between border-bottom p-3 border-color-secondary">
                            <span class="h6 mb-0">
                                Quantidade de FNC por Filial
                            </span>
                            <a onclick="abrirGraficoCompleto('fncPorFilial')" role="button" class="link-offset-2 link-offset-3-hover link-underline link-underline-opacity-0 link-underline-opacity-75-hover">
                                Ver Todos
                            </a>
                        </div>
                        <div class="card-body">
                            <div class="chart-container">
                                <canvas id="chartFilial"></canvas>
                            </div>
                        </div>
                    </div>
                </div>
                
                <div class="col-12 col-lg-6">
                    <div class="card shadow-sm border-0 h-100">
                        <div class="d-flex justify-content-between border-bottom p-3 border-color-secondary">
                            <span class="h6 mb-0">
                                Quantidade de FNC por Cliente
                            </span>
                            <a onclick="abrirGraficoCompleto('fncPorCliente')" role="button" class="link-offset-2 link-offset-3-hover link-underline link-underline-opacity-0 link-underline-opacity-75-hover">
                                Ver Todos
                            </a>
                        </div>
                        <div class="card-body">
                            <div class="chart-container">
                                <canvas id="chartCliente"></canvas>
                            </div>
                        </div>
                    </div>
                </div>
                
                <div class="col-12 col-lg-6">
                    <div class="card shadow-sm border-0 h-100">
                        <div class="d-flex justify-content-between border-bottom p-3 border-color-secondary">
                            <span class="h6 mb-0">
                                Quantidade de FNC por Defeito
                            </span>
                            <a onclick="abrirGraficoCompleto('fncPorDefeito')" role="button" class="link-offset-2 link-offset-3-hover link-underline link-underline-opacity-0 link-underline-opacity-75-hover">
                                Ver Todos
                            </a>
                        </div>
                        <div class="card-body">
                            <div class="chart-container">
                                <canvas id="chartDefeito"></canvas>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="col-12 col-lg-6">
                    <div class="card shadow-sm border-0 h-100">
                       <div class="d-flex justify-content-between border-bottom p-3 border-color-secondary">
                            <span class="h6 mb-0">
                                Quantidade de FNC por Setor                            
                            </span>
                            <a onclick="abrirGraficoCompleto('fncPorSetor')" role="button" class="link-offset-2 link-offset-3-hover link-underline link-underline-opacity-0 link-underline-opacity-75-hover">
                                Ver Todos
                            </a>
                        </div>
                        <div class="card-body">
                            <div class="chart-container">
                                <canvas id="chartSetor"></canvas>
                            </div>
                        </div>
                    </div>
                </div>
            `);


            charts.filial = criarGraficoBarras('chartFilial', dadosDashboard.fncPorFilial, true);
            charts.cliente = criarGraficoBarras('chartCliente', dadosDashboard.fncPorCliente, false);
            charts.defeito = criarGraficoPizza('chartDefeito', dadosDashboard.fncPorDefeito);
            charts.setor = criarGraficoPizza('chartSetor', dadosDashboard.fncPorSetor);
        }


        function abrirGraficoCompleto(tipoGrafico) {  
            let url = `DashboardFNCDetalhes.asp?tipoGrafico=${tipoGrafico}`;

            if ($('#filtroFilial').val()) url += `&filtroFilial=${$('#filtroFilial').val()}`;
            if ($('#filtroSetor').val()) url += `&filtroSetor=${$('#filtroSetor').val()}`;
            if ($('#filtroSituacao').val()) url += `&filtroSituacao=${$('#filtroSituacao').val()}`;
            if ($('#filtroDataInclusaoDe').val()) url += `&filtroDataInclusaoDe=${$('#filtroDataInclusaoDe').val()}`;
            if ($('#filtroDataInclusaoAte').val()) url += `&filtroDataInclusaoAte=${$('#filtroDataInclusaoAte').val()}`;
            if ($('#filtroDataEnvioClienteDe').val()) url += `&filtroDataEnvioClienteDe=${$('#filtroDataEnvioClienteDe').val()}`;
            if ($('#filtroDataEnvioClienteAte').val()) url += `&filtroDataEnvioClienteAte=${$('#filtroDataEnvioClienteAte').val()}`;
            if ($('#filtroDefeito').val()) url += `&filtroDefeito=${$('#filtroDefeito').val()}`;
            if ($('#filtroCodCliente').val()) url += `&filtroCodCliente=${$('#filtroCodCliente').val()}`;
            if ($('#filtroCodLoja').val()) url += `&filtroCodLoja=${$('#filtroCodLoja').val()}`;

            window.open(url, '_blank');
        }

        function limparFiltros() {
            $('#filtroFilial').val('').trigger('change');
            $('#filtroSetor').val('').trigger('change');
            $('#filtroSituacao').val('').trigger('change');
            $('#filtroDataInclusaoDe').val('');
            $('#filtroDataInclusaoAte').val('');
            $('#filtroDataEnvioClienteDe').val('');
            $('#filtroDataEnvioClienteAte').val('');
            $('#filtroDefeito').val('').trigger('change');
            $('#filtroCodCliente').val('');
            $('#filtroCodLoja').val('');
            
            dadosDashboard = null;
            renderizarDashboard();
        }

        function anexarEventos() {
            $('#btnPesquisar').on('click', pesquisar);
            $('#btnLimpar').on('click', limparFiltros);
        }

        function controlarExibicaoAviso() {    
            if (!$('#filtroFilial').val()) {
                $('#avisoFilial').removeClass('d-none');
            } else {
                $('#avisoFilial').addClass('d-none');
            }
        }

        $(document).ready(function() {
            inicializarSelect2();
            carregarSetores();
            carregarDefeitos();
            anexarEventos();
            renderizarDashboard();
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="d-flex my-4">
            <h3 class="mb-0">Dashboard de FNC<span id="idFNC"></span></h3>
        </div>

        <div class="card shadow-sm border-0 mb-4">
            <div class="card-body">
                <div class="row g-2">
                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroDataInclusaoDe" class="form-label">Data Inclusão De</label>
                        <input type="date" class="form-control form-control-sm" id="filtroDataInclusaoDe" name="filtroDataInclusaoDe">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroDataInclusaoAte" class="form-label">Data Inclusão Até</label>
                        <input type="date" class="form-control form-control-sm" id="filtroDataInclusaoAte" name="filtroDataInclusaoAte">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroDataEnvioClienteDe" class="form-label">Data Envio Cliente De</label>
                        <input type="date" class="form-control form-control-sm" id="filtroDataEnvioClienteDe" name="filtroDataEnvioClienteDe">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroDataEnvioClienteAte" class="form-label">Data Envio Cliente Até</label>
                        <input type="date" class="form-control form-control-sm" id="filtroDataEnvioClienteAte" name="filtroDataEnvioClienteAte">
                    </div>

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
                        <label for="filtroSetor" class="form-label">Setor</label>
                        <select multiple class="form-select form-select-sm select2" id="filtroSetor" name="filtroSetor" data-placeholder="Todos os setores">
                            <option value="">Todos os setores</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroSituacao" class="form-label">Situação</label>
                        <select multiple class="form-select form-select-sm select2" id="filtroSituacao" name="filtroSituacao" data-placeholder="Todas">
                            <option value="">Todas</option>
                            <option value="1">ABERTA</option>
                            <option value="2">RESPONDIDA</option>
                            <option value="3">ENCERRADA</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroValidacao" class="form-label">Validação</label>
                        <select multiple class="form-select form-select-sm select2" id="filtroValidacao" name="filtroValidacao" data-placeholder="Todas">
                            <option value="">Todas</option>
                            <option value="1">REGISTRADA</option>
                            <option value="2">EM ANÁLISE</option>
                            <option value="3">PROCEDE</option>
                            <option value="4">NÃO PROCEDE</option>
                            <option value="5">CANCELADA</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroDefeito" class="form-label">Defeito</label>
                        <select multiple class="form-select form-select-sm select2" id="filtroDefeito" name="filtroDefeito" data-placeholder="Todos os defeitos">
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
        </div>

        <div id="avisoFilial" class="d-none alert alert-warning mb-4">
            Atenção! Você está visualizando dados de todas as filiais. Para ver dados de uma filial específica, selecione-a no filtro acima.
        </div>

        <div class="row g-3 mb-4">
            <div class="col-6 col-lg-3">
                <div class="card shadow-sm border-0 h-100">
                    <div class="card-body d-flex justify-content-between align-items-center">
                        <div>
                            <div class="text-secondary mb-1">Total de FNCs</div>
                            <h2 id="kpiTotal">0</h2>
                        </div>
                        <div class="icon-wrapper purple">
                            <i class="bi bi-file-earmark-text"></i>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-6 col-lg-3">
                <div class="card shadow-sm border-0 h-100 border-start">
                    <div class="card-body d-flex justify-content-between align-items-center">
                        <div>
                            <div class="text-secondary mb-1">Abertas</div>
                            <h2 id="kpiAbertas">0</h2>
                        </div>
                        <div class="icon-wrapper yellow">
                            <i class="bi bi-clock"></i>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-6 col-lg-3">
                <div class="card shadow-sm border-0 h-100 border-start">
                    <div class="card-body d-flex justify-content-between align-items-center">
                        <div>
                            <div class="text-secondary mb-1">Respondidas</div>
                            <h2 id="kpiRespondidas">0</h2>
                        </div>
                        <div class="icon-wrapper blue">
                            <i class="bi bi-chat-text"></i>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-6 col-lg-3">
                <div class="card shadow-sm border-0 h-100 border-start">
                    <div class="card-body d-flex justify-content-between align-items-center">
                        <div>
                            <div class="text-secondary mb-1">Encerradas</div>
                            <h2 id="kpiEncerradas">0</h2>
                        </div>
                        <div class="icon-wrapper green">
                            <i class="bi bi-check-circle"></i>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="row g-3" id="dashboardContainer"></div>
    </div>

    <div id="loadingOverlay" class="position-fixed top-0 start-0 w-100 h-100 d-none justify-content-center align-items-center" style="background: rgba(255,255,255,0.8); z-index: 9999;">
        <div class="spinner-border text-primary" role="status">
            <span class="visually-hidden">Carregando...</span>
        </div>
    </div>
</body>
</html>