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

        .form-label {
            font-weight: 600;
        }

        .chart-container {
            position: relative;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
    <script src="../../../components/toast.js?v=<%=Now()%>"></script>

    <script>
        let dadosGrafico = null;
        let estiloGrafico = '';
        let carregando = false;
        let chart = {};

        let filtros = {
            tipoGrafico: '',
            filial: '',
            setor: '',
            situacao: '',
            validacao: '',
            dataInclusaoDe: '',
            dataInclusaoAte: '',
            dataEnvioClienteDe: '',
            dataEnvioClienteAte: '',
            defeito: '',
            codCliente: '',
            codLoja: ''
        }

        const chartConfig = {
            fncPorFilial: { 
                titulo: 'FNC por Filial',
                estiloGrafico: 'barra',
                codFiltro: 'filtroFilial'
            },
            fncPorCliente: {
                titulo: 'FNC por Cliente',
                estiloGrafico: 'barra',
                codFiltro: 'filtroCodCliente'
            },
            fncPorDefeito: {
                titulo: 'FNC por Defeito',
                estiloGrafico: 'pizza',
                codFiltro: 'filtroDefeito'
            },
            fncPorSetor: {
                titulo: 'FNC por Setor',
                estiloGrafico: 'pizza',
                codFiltro: 'filtroSetor'
            }
        };

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

        function exibirLoading(exibir) {
            carregando = exibir;
            if (exibir) {
                $('#loadingOverlay').removeClass('d-none').addClass('d-flex');
            } else {
                $('#loadingOverlay').removeClass('d-flex').addClass('d-none');
            }
        }

        function carregarDadosGrafico() {            
            exibirLoading(true);

            const params = {
                acao: 'obterDadosGrafico',
                tipoGrafico: filtros.tipoGrafico,
                filtroFilial: filtros.filial,
                filtroSetor: filtros.setor,
                filtroSituacao: filtros.situacao,
                filtroValidacao: filtros.validacao,
                filtroDataInclusaoDe: filtros.dataInclusaoDe,
                filtroDataInclusaoAte: filtros.dataInclusaoAte,
                filtroDataEnvioClienteDe: filtros.dataEnvioClienteDe,
                filtroDataEnvioClienteAte: filtros.dataEnvioClienteAte,
                filtroDefeito: filtros.defeito,
                filtroCodCliente: filtros.codCliente,
                filtroCodLoja: filtros.codLoja
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

                    dadosGrafico = response.dados || [];
                    renderizarDetalhes();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao pesquisar:', error);
                    Toast.error('Erro ao realizar a pesquisa');
                }
            });
        }

        function formatarDadosSimples(dados, horizontal) {
            const labels = dados.map(item => item.label);
            const totais = dados.map(item => item.total);

            datasets = [{
                label: 'Total',
                data: totais,
                backgroundColor: coresGrafico.slice(0, totais.length),
                borderColor: bordasGrafico.slice(0, totais.length),
                borderWidth: 0.5,
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
                borderWidth: 0.5,
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
                            
                            navegarParaRelatorio(itemValor); 
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
                            
                            navegarParaRelatorio(itemValor); 
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

        function navegarParaRelatorio(itemValor) {
            // Ao clicar em um item o proprio item será enviado como filtro
            const filtroClique = `${chartConfig[filtros.tipoGrafico].codFiltro}=${itemValor}`;

            let url = `../RelatorioFNC/RelatorioFNC.asp?${filtroClique}`;
      
            if (filtros.dataInclusaoDe) url += `&filtroDataInclusaoDe=${filtros.dataInclusaoDe}`;
            if (filtros.dataInclusaoAte) url += `&filtroDataInclusaoAte=${filtros.dataInclusaoAte}`;
            if (filtros.dataEnvioClienteDe) url += `&filtroDataEnvioClienteDe=${filtros.dataEnvioClienteDe}`;
            if (filtros.dataEnvioClienteAte) url += `&filtroDataEnvioClienteAte=${filtros.dataEnvioClienteAte}`;
            if (filtros.filial) url += `&filtroFilial=${filtros.filial}`;
            if (filtros.codCliente) url += `&filtroCodCliente=${filtros.codCliente}`;
            if (filtros.codLoja) url += `&filtroCodLoja=${filtros.codLoja}`;
            if (filtros.setor) url += `&filtroSetor=${filtros.setor}`;
            if (filtros.situacao) url += `&filtroSituacao=${filtros.situacao}`;
            if (filtros.validacao) url += `&filtroValidacao=${filtros.validacao}`;
            if (filtros.defeito) url += `&filtroDefeito=${filtros.defeito}`;

            window.open(url, '_blank', 'width=1200,height=800');
        }

        function renderizarDetalhes() {
            const chartHeight = estiloGrafico === 'pizza' ? '45vw' : `${(dadosGrafico.length * 32) + 15}px`;

            $('#chartContainer').html(`
                <div style="height: ${chartHeight}" class="chart-container">
                    <canvas id="chart"></canvas>
                </div>
            `);

            chart = estiloGrafico === 'barra' ? 
                criarGraficoBarras('chart', dadosGrafico, true) 
                : criarGraficoPizza('chart', dadosGrafico);
        }


        $(document).ready(function() {
            const urlParams = new URLSearchParams(window.location.search);
            
            if (urlParams.has('tipoGrafico')) filtros.tipoGrafico = urlParams.get('tipoGrafico');
            if (urlParams.has('filtroFilial')) filtros.filial = urlParams.get('filtroFilial');
            if (urlParams.has('filtroSetor')) filtros.setor = urlParams.get('filtroSetor');
            if (urlParams.has('filtroSituacao')) filtros.situacao = urlParams.get('filtroSituacao');
            if (urlParams.has('filtroValidacao')) filtros.validacao = urlParams.get('filtroValidacao');
            if (urlParams.has('filtroDataInclusaoDe')) filtros.dataInclusaoDe = urlParams.get('filtroDataInclusaoDe');
            if (urlParams.has('filtroDataInclusaoAte')) filtros.dataInclusaoAte = urlParams.get('filtroDataInclusaoAte');
            if (urlParams.has('filtroDataEnvioClienteDe')) filtros.dataEnvioClienteDe = urlParams.get('filtroDataEnvioClienteDe');
            if (urlParams.has('filtroDataEnvioClienteAte')) filtros.dataEnvioClienteAte = urlParams.get('filtroDataEnvioClienteAte');
            if (urlParams.has('filtroDefeito')) filtros.defeito = urlParams.get('filtroDefeito');
            if (urlParams.has('filtroCodCliente')) filtros.codCliente = urlParams.get('filtroCodCliente');
            if (urlParams.has('filtroCodLoja')) filtros.codLoja = urlParams.get('filtroCodLoja');


            $('#chartTitle').text(chartConfig[filtros.tipoGrafico].titulo || 'Gráfico');
            estiloGrafico = chartConfig[filtros.tipoGrafico].estiloGrafico || 'barra';

            carregarDadosGrafico();
        });
    </script>
</head>
<body>
    <div class="p-4">
        <h3 id="chartTitle" class="mb-4">Dashboard de FNC</h3>
        <div id="chartContainer"></div>

        <div id="loadingOverlay" class="position-fixed top-0 start-0 w-100 h-100 d-none justify-content-center align-items-center" style="background: rgba(255,255,255,0.8); z-index: 9999;">
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">Carregando...</span>
            </div>
        </div>
    </div>
</body>
</html>