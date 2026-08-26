<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Relatório de Permissões de Usuário</title>
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

        .required-field {
            color: #dc3545;
        }

        .table-wrapper {
            overflow: auto;
            max-height: 70vh;
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

        input[type="checkbox"] {
            width: 1rem;
            height: 1rem;
            cursor: pointer;
        }

        .badge {
            font-size: .7rem;
        }
    </style>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let dadosUsuarios = [];
        let usuariosSelecionados = [];
        let carregando = false;

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                allowClear: true,
                placeholder: function() {
                    return $(this).data('placeholder');
                }
            });
        }

        function carregarSubCategorias() {
            const $select = $('#filtroSubcategoria');
            $select.empty().append('<option value="">Todos os subcategorias</option>');

            const params = {
                acao: 'listarSubcategorias',
                filtroCategoria: $('#filtroCategoria').val()
            };

            $.ajax({
                url: 'RelatorioPermissaoUsuariosRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    if (!response.sucesso) {
                        console.error("Erro ao carregar subcategorias: " + response.mensagem);
                        return;
                    }

                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.descricao }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar subcategorias:', error);
                }
            });
        }

        function navegarParaPermissoesUsuario(matricula, filial, nome) { 
            window.location.href = `/corporativo/Usuario/AcessosUsuarios.asp?matricula=${matricula}&filial=${filial}&nome=${nome}`
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
            exibirLoading(true);

            const params = {
                acao: 'listarUsuarios',
                filtroFilial: $('#filtroFilial').val(),
                filtroCategoria: $('#filtroCategoria').val(),
                filtroSubcategoria: $('#filtroSubcategoria').val() ? $('#filtroSubcategoria').val().join(',') : '',
                filtroId: $('#filtroId').val(),
                filtroMatricula: $('#filtroMatricula').val(),
                filtroEmail: $('#filtroEmail').val(),
            };

            $.ajax({
                url: 'RelatorioPermissaoUsuariosRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error('Erro ao pesquisar: ' + response.mensagem);
                        return;
                    }

                    dadosUsuarios = response.dados || [];
                    usuariosSelecionados = [];
                    renderizarGrade();
                    $('#contadorSelecionados').text(usuariosSelecionados.length);

                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao pesquisar:', error);
                    Toast.error('Erro ao realizar a pesquisa');
                }
            });
        }

        function atualizarSelecao() {
            usuariosSelecionados = [];
            $('.rowCheckbox:checked').each(function() {
                const id = $(this).data('id');
                const usuario = dadosUsuarios.find(u => String(u.id) === String(id));
                if (usuario) {
                    usuariosSelecionados.push(usuario);
                }
            });

            $('#contadorSelecionados').text(usuariosSelecionados.length);
        }

        function renderizarGrade() {
            const $tbody = $('#gradeUsuarios tbody');
            
            $tbody.empty();
            $('#selectAll').prop('checked', false);
            
            if (dadosUsuarios.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Utilize os filtros acima para pesquisar
                        </td>
                    </tr>
                `);
                return;
            }

            dadosUsuarios.forEach(function(item) {
                $tbody.append(`
                    <tr>
                        <td><input type="checkbox" class="rowCheckbox" data-id="${item.id}" aria-label="Selecionar item"></td>
                        <td>${item.id}</td>
                        <td>${item.matricula || '-'}</td>
                        <td>${item.nomePlanta || '-'}</td>
                        <td>${item.funcao || '-'}</td>
                        <td>${item.login || '-'}</td>
                        <td>${item.nome || '-'}</td>
                        <td>${item.email || '-'}</td>
                        <td>${item.servicos.length}</td>
                        <td>${item.relatorios.length}</td>
                        <td>${item.dataCadastro || '-'}</td>
                        <td>
                           <div class="d-flex gap-2">
                                <button type="button" class="btn btn-sm btn-primary" title="Visualizar Acessos"
                                    onclick="navegarParaPermissoesUsuario('${item.matricula}', '${item.codPlanta}', '${item.nome}')">
                                    <i class="bi bi-box-arrow-up-right"></i>
                                </button>
                           </div>
                        </td>
                    </tr>
                `);
            });
        }

        function limparFiltros() {
            $('#filtroFilial').val('').trigger('change');
            $('#filtroCategoria').val('').trigger('change');
            $('#filtroSubcategoria').val('').prop('disabled', true).trigger('change');
            $('#filtroId').val('');
            $('#filtroMatricula').val('');
            $('#filtroEmail').val('');

            dadosUsuarios = [];
            usuariosSelecionados = [];
            renderizarGrade();
            $('#contadorSelecionados').text(usuariosSelecionados.length);
        }

        function coletarDadosSubcategorias(usuarios) {
            const subcategorias = {};

            usuarios.forEach(function(usuario) {
                if (usuario.servicos && usuario.servicos.length > 0) {
                    usuario.servicos.forEach(function(servico) {
                        const nomeAba = servico.NomeSubCategoria + ' (Serviços)';
                        if (!subcategorias[nomeAba]) {
                            subcategorias[nomeAba] = [];
                        }
                        subcategorias[nomeAba].push({
                            nomeUsuario: usuario.nome,
                            emailUsuario: usuario.email,
                            idUsuario: usuario.id,
                            codSubcategoria: servico.CodSubcategoria,
                            descricao: servico.Descricao,
                            categoria: 'Serviços'
                        });
                    });
                }

                if (usuario.relatorios && usuario.relatorios.length > 0) {
                    usuario.relatorios.forEach(function(relatorio) {
                        const nomeAba = relatorio.NomeSubCategoria + ' (Relatórios)';
                        if (!subcategorias[nomeAba]) {
                            subcategorias[nomeAba] = [];
                        }
                        subcategorias[nomeAba].push({
                            nomeUsuario: usuario.nome,
                            emailUsuario: usuario.email,
                            idUsuario: usuario.id,
                            codSubcategoria: relatorio.CodSubcategoria,
                            descricao: relatorio.Descricao,
                            categoria: 'Relatórios'
                        });
                    });
                }
            });

            Object.keys(subcategorias).forEach(function(key) {
                subcategorias[key].sort(function(a, b) {
                    return a.nomeUsuario.localeCompare(b.nomeUsuario);
                });
            });

            return subcategorias;
        }

        function sanitizarNomeAba(nome) {
            var sanitizado = nome.replace(/[\\\/\?\*\[\]]/g, '_');
            if (sanitizado.length > 31) {
                sanitizado = sanitizado.substring(0, 31);
            }
            return sanitizado;
        }

        function exportarExcel() {
            if (usuariosSelecionados.length === 0) {
                Toast.warning('Selecione ao menos um usuário para exportar');
                return;
            }

            const subcategorias = coletarDadosSubcategorias(usuariosSelecionados);
            const abas = Object.keys(subcategorias);

            if (abas.length === 0) {
                Toast.warning('Não há dados de acessos para exportar');
                return;
            }

            const wb = XLSX.utils.book_new();

            abas.sort().forEach(function(nomeAba) {
                const dados = subcategorias[nomeAba];
                const linhas = [['ID Usuário', 'Nome', 'Email', 'Cód. Subcategoria', 'Descrição', 'Categoria']];

                dados.forEach(function(item) {
                    linhas.push([
                        item.idUsuario,
                        item.nomeUsuario,
                        item.emailUsuario,
                        item.codSubcategoria,
                        item.descricao,
                        item.categoria
                    ]);
                });

                const ws = XLSX.utils.aoa_to_sheet(linhas);
                ws['!cols'] = [
                    { wch: 12 },
                    { wch: 30 },
                    { wch: 30 },
                    { wch: 18 },
                    { wch: 30 },
                    { wch: 15 }
                ];

                XLSX.utils.book_append_sheet(wb, ws, sanitizarNomeAba(nomeAba));
            });

            XLSX.writeFile(wb, 'Relatorio_Acessos_Usuarios.xlsx');
        }

        function exportarPDF() {
            if (usuariosSelecionados.length === 0) {
                Toast.warning('Selecione ao menos um usuário para exportar');
                return;
            }

            const subcategorias = coletarDadosSubcategorias(usuariosSelecionados);
            const abas = Object.keys(subcategorias);

            if (abas.length === 0) {
                Toast.warning('Não há dados de acessos para exportar');
                return;
            }

            const { jsPDF } = window.jspdf;
            const doc = new jsPDF('landscape');
            const dataAtual = new Date().toLocaleDateString('pt-BR');
            let primeiraPagina = true;

            abas.sort().forEach(function(nomeAba) {
                if (!primeiraPagina) {
                    doc.addPage();
                }
                primeiraPagina = false;

                const dados = subcategorias[nomeAba];

                doc.setFontSize(14);
                doc.text(nomeAba, 14, 15);

                doc.setFontSize(9);
                doc.text('Data: ' + dataAtual + ' | ' + dados.length + ' registro(s)', 14, 22);

                const colunas = ['ID Usuário', 'Nome', 'Email', 'Cód. Subcategoria', 'Descrição', 'Categoria'];

                const linhas = dados.map(function(item) {
                    return [
                        item.idUsuario,
                        item.nomeUsuario,
                        item.emailUsuario,
                        item.codSubcategoria,
                        item.descricao,
                        item.categoria
                    ];
                });

                doc.autoTable({
                    head: [colunas],
                    body: linhas,
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
                    margin: { top: 28 }
                });
            });

            window.open(doc.output('bloburl'), '_blank');
        }

        function anexarEventos() {
            $('#btnPesquisar').on('click', pesquisar);
            $('#btnLimpar').on('click', limparFiltros);
            $('#btnExportarExcel').on('click', exportarExcel);
            $('#btnExportarPDF').on('click', exportarPDF);

            $('#selectAll').on('click', function() {
                $('.rowCheckbox').prop('checked', this.checked);
                atualizarSelecao();
            });

            $('#filtroCategoria').on('change', function() {
                if (!this.value) {
                    $('#filtroSubcategoria').val('').prop('disabled', true).trigger('change');
                    return;
                }
                
                $('#filtroSubcategoria').prop('disabled', false).trigger('change');
                carregarSubCategorias();
            });

            $(document).on('click', '.rowCheckbox', function() {
                if (!this.checked) {
                    $('#selectAll').prop('checked', false);
                }

                if ($('.rowCheckbox:checked').length === $('.rowCheckbox').length && $('.rowCheckbox').length > 0) {
                    $('#selectAll').prop('checked', true);
                }

                atualizarSelecao();
            });
        }

        $(document).ready(function() {
            inicializarSelect2();
            anexarEventos();
            renderizarGrade();
            $('#contadorSelecionados').text(usuariosSelecionados.length);
        });
    </script>
</head>
<body>
    <div class="container my-4">
        <h3 class="mb-4">Relatório de Permissões de Usuário</h3>

        <div class="card p-3 mb-4">
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
                    <label for="filtroCategoria" class="form-label">Categoria</label>
                    <select class="form-select form-select-sm select2" id="filtroCategoria" name="filtroCategoria" data-placeholder="Todas">
                        <option value="">Todas</option>
                        <option value="1">Serviços</option>
                        <option value="2">Relatórios</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroSubcategoria" class="form-label">Subcategorias</label>
                    <select disabled multiple class="form-select form-select-sm select2" id="filtroSubcategoria" name="filtroSubcategoria" data-placeholder="Todos as subcategorias">
                        <option value="">Todos as subcategorias</option>
                    </select>
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroId" class="form-label">Id</label>
                    <input type="text" class="form-control form-control-sm" id="filtroId" name="filtroId">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroMatricula" class="form-label">Matricula</label>
                    <input type="text" class="form-control form-control-sm" id="filtroMatricula" name="filtroMatricula">
                </div>

                <div class="col-12 col-md-4 col-lg-3">
                    <label for="filtroEmail" class="form-label">Email</label>
                    <input type="text" class="form-control form-control-sm" id="filtroEmail" name="filtroEmail">
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
            <div id="divExportar" class="d-flex p-3 gap-2 align-items-center">
                <button type="button" class="btn btn-success btn-sm" id="btnExportarExcel">
                    <i class="bi bi-file-earmark-excel me-1"></i> Exportar Excel 
                </button>
                <button type="button" class="btn btn-danger btn-sm" id="btnExportarPDF">
                    <i class="bi bi-file-earmark-pdf me-1"></i> Exportar PDF 
                </button>
                <span class="badge bg-primary"><span id="contadorSelecionados">0</span> selecionado(s)</span>
            </div>

            <div class="table-wrapper">
                <div id="loadingOverlay" class="loading-overlay" style="display: none;">
                    <div class="spinner-border text-primary" role="status">
                        <span class="visually-hidden">Carregando...</span>
                    </div>
                </div>
                
                <table class="table table-hover mb-0" id="gradeUsuarios">
                    <thead>
                        <tr>
                            <th><input type="checkbox" id="selectAll" aria-label="Selecionar todos registros"></th>
                            <th>#</th>
                            <th>Matrícula</th>   
                            <th>Filial</th>
                            <th>Função</th>        
                            <th>Login</th>
                            <th>Nome</th>
                            <th>Email</th>  
                            <th>Serviços</th>  
                            <th>Relatórios</th>  
                            <th>Data de Cadastro</th>  
                            <th style="width: 70px;">Ações</th>  
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td colspan="29" class="text-center text-muted py-4">
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
