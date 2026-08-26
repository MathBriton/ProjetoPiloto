<!DOCTYPE html>
<html lang="pt-br">
<head>
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

        #gradeChamado {
            width: max-content;
            min-width: 100%;
        }

        #gradeChamado th,
        #gradeChamado td {
            white-space: nowrap;
            vertical-align: top;
        }

        #gradeChamado th:last-child,
        #gradeChamado td:last-child {
            position: sticky;
            right: 0;
            z-index: 5;
            white-space: nowrap;
        }
        
        #gradeChamado td:last-child {
            background: linear-gradient(to right, rgba(0,0,0,0.1) 0%, transparent 1px),#fff;
            background-origin: border-box;
            background-clip: padding-box;
        }
        
        #gradeChamado thead th:last-child {
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
        
        .badge {
            font-size: .75rem;
            font-weight: 500;
        }
        .badge-red {
            background-color: #fee2e2;
            color: #991b1b;
        }
        .badge-green {
            background-color: #d1fae5;
            color: #065f46;
        }        
        .badge-purple {
            background-color: #e0e7ff;
            color: #3730a3;
        }
        .badge-pink {
            background-color: #fce7f3;
            color: #9f1239;
        }
        .badge-blue {
            background-color: #cffafe;
            color: #155e75;
        }
        .badge-yellow {
            background-color: #fef3c7;
            color: #92400e;
        }
        .badge-gray {
            background-color: #e5e5e5;
            color: #3e4a63;
        }
        .badge-stone {
            background-color: #eee8da;
            color: #4a453e;
        }
    </style>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js"></script>

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/confirmation.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>
    <script src="/corporativo/components/dropdownMenu.js?v=<%=Now()%>"></script>
    <script src="../Compartilhado/funcoes.js?v=<%=Now()%>"></script>

    <script>
        let listaChamados = [];

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').show();
            } else {
                $('#loadingOverlay').hide();
            }
        }

        function limparFiltros() {
            $('#filtroIdChamado').val('');
            $('#filtroDataCriacaoDe').val('');
            $('#filtroDataCriacaoAte').val('');
            $('#filtroFilial').val('').trigger('change');
            $('#filtroDepartamento').val('').trigger('change');
            $('#filtroCategoria').val('').trigger('change');
            $('#filtroStatus').val('').trigger('change');
            $('#filtroPrioridade').val('').trigger('change');
        }
        
        function carregarChamados() {  
            const params = {
                acao: 'listar',
                filtroIdChamado: $('#filtroIdChamado').val(),
                filtroFilial: $('#filtroFilial').val(),
                filtroDataCriacaoDe: $('#filtroDataCriacaoDe').val(),
                filtroDataCriacaoAte: $('#filtroDataCriacaoAte').val(),
                filtroDepartamento: $('#filtroDepartamento').val(),
                filtroCategoria: $('#filtroCategoria').val(),
                filtroStatus: $('#filtroStatus').val(),
                filtroPrioridade: $('#filtroPrioridade').val(),
                flagMeusChamados: 1
            };

            exibirLoading(true);

            $.ajax({
                url: 'ChamadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        Toast.error(response.mensagem);
                        return;
                    }
                
                    listaChamados = response.dados || [];
                    renderizarGrade();
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao pesquisar:', error);
                    Toast.error('Erro ao realizar a pesquisa');
                }, complete: function() {
                    exibirLoading(false);
                }
            });
        }

        function renderizarGrade() {
            const $tbody = $('#gradeChamado tbody');
            $tbody.empty();

            if (listaChamados.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Nenhum registro encontrado
                        </td>
                    </tr>
                `);
                return;
            }
            

            listaChamados.forEach(function(item) {    
                $tbody.append(`
                    <tr>
                        <td>${item.id || '-'}</td>
                        <td>${obterNomePlanta(item.codPlanta)}</td>
                        <td>${item.titulo || '-'}</td>
                        <td>${item.nomeAtendente || 'Nenhum atendente'}</td>
                        <td>${item.nomeDepartamento || '-'}</td>
                        <td>${item.nomeCategoria || '-'}</td>
                        <td>${obterBadgePrioridade(item.nivelPrioridade)}</td>
                        <td>${obterBadgeStatus(item.status)}</td>
                        <td>${item.dataCriacao || '-'}</td>
                        <td>${item.dataFinalizacao || '-'}</td>
                        <td style="text-align: center">
                            <dropdown-menu label="Ações">
                                <a class="dropdown-item" onclick="visualizarChamado('${item.tokenAcesso}')"><i class="bi bi-eye me-1"></i> Visualizar</a>,
                            </dropdow-menu>
                        </td>
                    </tr>
                `);
            });
        }

        function visualizarChamado(tokenAcesso) {
            window.location.href = `ChamadoDetalhes.asp?token=${tokenAcesso}`;
        }

        function inicializarSelect2() {
            $('.select2').select2({ theme: 'bootstrap-5', width: '100%', dropdownCssClass: "z-index-9999", allowClear: true });
        }

        function campoPlantaAlterado() {
            const codPlanta = $('#filtroFilial').val();
            carregarDepartamentos(codPlanta);
        }


        function carregarDepartamentos(codPlanta) {
            const $select = $('#filtroDepartamento');
            $select.empty().append('<option value="">Todos</option>');

            if(!codPlanta) {
                $select.prop('disabled', true);
                return;
            }

            $select.prop('disabled', false);

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

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar departamentos:', error);
                }
            });
        }

        function carregarCategorias() {
            const $select = $('#filtroCategoria');
            $select.empty().append('<option value="">Todas</option>').prop('disabled', false);

            const params = {
                acao: 'listar'
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

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar categorias:', error);
                }
            });
        }

        function anexarEventos() {
            $('#filtroFilial').on('change', function() {
                campoPlantaAlterado();
            });
        }

        $(document).ready(function() {
            inicializarSelect2();
            anexarEventos();

            const urlParams = new URLSearchParams(window.location.search);

            var p1 = carregarDepartamentos(urlParams.get('filtroFilial'));
            var p2 = carregarCategorias();

            $.when(p1, p2).then(function() {
                setTimeout(() => {
                    if (urlParams.has('filtroFilial')) 
                        $('#filtroFilial').val(urlParams.get('filtroFilial')).trigger('change.select2');

                    if (urlParams.has('filtroDataCriacaoDe')) 
                        $('#filtroDataCriacaoDe').val(urlParams.get('filtroDataCriacaoDe'));

                    if (urlParams.has('filtroDataCriacaoAte')) 
                        $('#filtroDataCriacaoAte').val(urlParams.get('filtroDataCriacaoAte'));

                    if (urlParams.has('filtroDepartamento')) 
                        $('#filtroDepartamento').val(urlParams.get('filtroDepartamento')).trigger('change.select2');

                    if (urlParams.has('filtroCategoria')) 
                        $('#filtroCategoria').val(urlParams.get('filtroCategoria')).trigger('change.select2');

                    if (urlParams.has('filtroStatus')) 
                        $('#filtroStatus').val(urlParams.get('filtroStatus')).trigger('change.select2');

                    if (urlParams.has('filtroPrioridade')) 
                        $('#filtroPrioridade').val(urlParams.get('filtroPrioridade')).trigger('change.select2');

                    if (urlParams.has('filtroIdChamado')) 
                        $('#filtroIdChamado').val(urlParams.get('filtroIdChamado'));

                    carregarChamados();
                }, 1000)
            });
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <div class="d-flex my-4">
            <h3>Meus Chamados</h3>
        </div>

        <div class="card shadow-sm border-0 mb-4">
            <div class="card-body">
                <div class="row g-2">
                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroDataCriacaoDe" class="form-label">Data Criação De</label>
                        <input type="date" class="form-control form-control-sm" id="filtroDataCriacaoDe" name="filtroDataCriacaoDe">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroDataCriacaoAte" class="form-label">Data Criação Até</label>
                        <input type="date" class="form-control form-control-sm" id="filtroDataCriacaoAte" name="filtroDataCriacaoAte">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroFilial" class="form-label">Filial</label>
                        <select class="form-select form-select-sm select2" id="filtroFilial" name="filtroFilial" data-placeholder="Todas as filiais">
                            <option value="">Todas as filiais</option>
                            <option value="00">Pirapetinga</option>
                            <option value="02">Corporativo</option>
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
                        <label for="filtroDepartamento" class="form-label">Departamento</label>
                        <select disabled class="form-select form-select-sm select2" id="filtroDepartamento" name="filtroDepartamento" data-placeholder="Todos">
                            <option value="">Todos</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroCategoria" class="form-label">Categoria</label>
                        <select class="form-select form-select-sm select2" id="filtroCategoria" name="filtroCategoria" data-placeholder="Todos">
                            <option value="">Todos</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroStatus" class="form-label">Status</label>
                        <select class="form-select form-select-sm select2" id="filtroStatus" name="filtroStatus" data-placeholder="Todos">
                            <option value="">Todos</option>
                            <option value="PENDENTE">Pendente</option>
                            <option value="ABERTO">Aberto</option>
                            <option value="EM_ANALISE">Em Análise</option>
                            <option value="AGUARDANDO_TERCEIROS">Aguardando Terceiros</option>
                            <option value="FINALIZADO">Finalizado</option>
                            <option value="CANCELADO">Cancelado</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroPrioridade" class="form-label">Prioridade</label>
                        <select class="form-select form-select-sm select2" id="filtroPrioridade" name="filtroPrioridade" data-placeholder="Todos">
                            <option value="">Todos</option>
                            <option value="BAIXA">Baixa</option>
                            <option value="MEDIA">Média</option>
                            <option value="ALTA">Alta</option>
                            <option value="URGENTE">Urgente</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroIdChamado" class="form-label">Id Chamado</label>
                        <input placeholder="Id" type="text" class="form-control form-control-sm" id="filtroIdChamado" name="filtroIdChamado">
                    </div>

                    <div class="col-12 d-flex justify-content-end align-items-end gap-2">
                        <button onclick="limparFiltros()" type="button" class="btn btn-sm btn-outline-secondary" id="btnLimpar">
                            <i class="bi bi-x-lg me-1"></i> Limpar
                        </button>
                        <button onclick="carregarChamados()" type="button" class="btn btn-sm btn-primary" id="btnPesquisar">
                            <i class="bi bi-search me-1"></i> Pesquisar
                        </button>
                    </div>
                </div>
            </div>
        </div>
       
        <div class="card">
            <div class="table-wrapper">
                <div id="loadingOverlay" class="loading-overlay" style="display: none;">
                    <div class="spinner-border text-primary" role="status">
                        <span class="visually-hidden">Carregando...</span>
                    </div>
                </div>

                <table class="table table-hover mb-0" id="gradeChamado">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Filial</th>
                            <th>Assunto</th>
                            <th>Atendente</th>
                            <th>Departamento</th>
                            <th>Categoria</th>
                            <th>Prioridade</th>
                            <th>Status</th>
                            <th>Data Criação</th>
                            <th>Data Finalização</th>
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
