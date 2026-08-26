<%
Response.Buffer = true
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Response.AddHeader "Pragma", "no-cache"
%>
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Smurfit Kappa - Relatório de Faturamento Detalhado</title>
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

        h6 {
            font-size: 1.1rem !important;
        }
        
        .container {
            max-width: 100%;
        }

        .form-label {
            font-weight: 600;
        }

        .table-wrapper {
            overflow: auto;
            max-height: 70vh;
        }

        #gradeRelatorio {
            width: max-content;
            min-width: 100%;
        }

        #gradeRelatorio th,
        #gradeRelatorio td {
            white-space: nowrap;
            vertical-align: top;
            padding: 8px 12px !important;
        }

        thead {
            vertical-align: middle !important;
            position: sticky; 
            inset-block-start: 0; 
            background: #f5f5f5;
            z-index: 10;
        }

        tfoot {
            vertical-align: middle !important;
            position: sticky; 
            inset-block-start: 0; 
            background: #f5f5f5;
            z-index: 10;
            bottom: 0;
            font-weight: 600;
            color: #495057;
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
        .switch-item {
            display: flex;
            gap: 1rem;
        }
        .switch-item .form-check-input {
            width: 2.75rem;
            height: 1.45rem;
            cursor: pointer;
            flex-shrink: 0;
        }
        .switch-label {
            font-size: .92rem;
            font-weight: 500;
            margin-bottom: .1rem;
            cursor: pointer;
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

        .icon-wrapper.light-blue {
            background-color: #ddf4fb;
            color: #1ab1cb;
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

        .vertical-line {
            width: 1px;
            background-color: #d1d2d4;
        }

        .horizontal-line {
            height: 1px;
            background-color: #d1d2d4;
        }
    </style>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    
    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let dadosRelatorio = [];
        let carregando = false;

        // Ao entrar na página
        $(document).ready(function() {
            inicializarSelect2();
            anexarEventos();
            renderizarDados();

            const urlParams = new URLSearchParams(window.location.search);
            if (urlParams.size === 0) return;

            setTimeout(() => {
                if (urlParams.has('filtroFilial')) 
                    $('#filtroFilial').val(urlParams.get('filtroFilial')).trigger('change');

                if (urlParams.has('filtroDataInclusaoDe')) 
                    $('#filtroDataInclusaoDe').val(urlParams.get('filtroDataInclusaoDe'));

                if (urlParams.has('filtroDataInclusaoAte')) 
                    $('#filtroDataInclusaoAte').val(urlParams.get('filtroDataInclusaoAte'));

                if (urlParams.has('filtroCodCliente')) 
                    $('#filtroCodCliente').val(urlParams.get('filtroCodCliente'));
                
                if (urlParams.has('filtroCodLoja')) 
                    $('#filtroCodLoja').val(urlParams.get('filtroCodLoja'));

                if (urlParams.has('filtroRepresentante')) 
                    $('#filtroRepresentante').val(urlParams.get('filtroRepresentante')).trigger('change');

                if (urlParams.has('filtroClienteGrupo')) 
                    $('#filtroClienteGrupo').val(urlParams.get('filtroClienteGrupo')).trigger('change');

                if (urlParams.has('filtroGrupoProduto')) 
                    $('#filtroGrupoProduto').val(urlParams.get('filtroGrupoProduto')).trigger('change');

                if (urlParams.has('filtroTipoProduto')) 
                    $('#filtroTipoProduto').val(urlParams.get('filtroTipoProduto')).trigger('change');

                if (urlParams.has('filtroCodProduto')) {
                    $('#filtroCodProduto').val(urlParams.get('filtroCodProduto'));
                    if (urlParams.has('filtroDescProduto'))
                        $('#filtroDescProduto').val(urlParams.get('filtroDescProduto'));
                }

                if (urlParams.has('filtroTipoNotaFiscal')) 
                    $('#filtroTipoNotaFiscal').val(urlParams.get('filtroTipoNotaFiscal')).trigger('change');

                if (urlParams.has('filtroPedido')) 
                    $('#filtroPedido').val(urlParams.get('filtroPedido'));

                if (urlParams.has('filtroNotaFiscal')) 
                    $('#filtroNotaFiscal').val(urlParams.get('filtroNotaFiscal'));

                pesquisar();
            }, 1000)
        });

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

        function carregarRepresentantes() {
            const $select = $('#filtroRepresentante');
            $select.empty().append('<option value="">Todos os representantes</option>');

            const params = {
                acao: 'listarRepresentantes',
                filtroFilial: $('#filtroFilial').val()
            };

            $.ajax({
                url: 'RelatorioFaturamentoDetalhadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                success: function(data) {
                    var response;
                    try {
                        response = (typeof data === 'string') ? JSON.parse(data) : data;
                    } catch(e) {
                        console.error('Erro ao carregar representantes:', e);
                        return;
                    }
                    if (!response.sucesso) {
                        Toast.error("Erro ao carregar representantes: " + response.mensagem);
                        return;
                    }
                    response.dados.forEach(function(item) {
                        $select.append($('<option>', { value: item.id, text: item.nome }));
                    });
                },
                error: function(xhr, status, error) {
                    console.error('Erro ao carregar representantes:', error);
                }
            });
        }

        function obterCodPlanta(nomePlanta) {
            const codPlantaMap = {
                "PIRAPETINGA": "00",
                "ESHOP": "00",
                "UBERABA": "00",
                "BENTO": "05",
                "FORTALEZA": "06",
                "MARANGUAPE": "06",
                "SAQUAREMA": "07",
            }

            return codPlantaMap[nomePlanta] || "";
        }
        function pesquisarCliente() {
            const codPlanta = obterCodPlanta($('#filtroFilial').val());
            const baseUrl = `${location.protocol}//${location.host}`;
            const url = `${baseUrl}/corporativo/buscaCLIENTE.asp?id1=filtroCodCliente&id2=filtroCodLoja&id3=nomeCliente&id6=${codPlanta}`;
            window.open(url, '', "width=650,height=400");
        }

        function exibirLoading(exibir) {
            carregando = exibir;
            if (exibir) {
                $('#loadingOverlay').show();
            } else {
                $('#loadingOverlay').hide();
            }
        }

        function validarFiltros(e) {
            e.preventDefault();
            
            const form = $('#formularioFiltro')[0];
            if (!form.checkValidity()) {
                e.stopPropagation();
                $(form).addClass('was-validated');
                return false;
            }

            pesquisar();
        }

        function pesquisar() {            
            exibirLoading(true);

            const params = {
                acao: 'listar',
                filtroFilial: $('#filtroFilial').val(),
                filtroDataInclusaoDe: $('#filtroDataInclusaoDe').val(),
                filtroDataInclusaoAte: $('#filtroDataInclusaoAte').val(),
                filtroCodCliente: $('#filtroCodCliente').val(),
                filtroCodLoja: $('#filtroCodLoja').val(),
                filtroRepresentante: $('#filtroRepresentante').val(),
                filtroClienteGrupo: $('#filtroClienteGrupo').val(),
                filtroGrupoProduto: $('#filtroGrupoProduto').val(),
                filtroTipoProduto: $('#filtroTipoProduto').val(),
                filtroCodProduto: $('#filtroCodProduto').val(),
                filtroTipoNotaFiscal: $('#filtroTipoNotaFiscal').val(),
                filtroPedido: $('#filtroPedido').val(),
                filtroNotaFiscal: $('#filtroNotaFiscal').val(),
                flagTempoReal: $('#flagTempoReal').is(':checked') ? 1 : 0
            };

            $.ajax({
                url: 'RelatorioFaturamentoDetalhadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                success: function(data) {
                    exibirLoading(false);

                    var response;
                    try {
                        response = (typeof data === 'string') ? JSON.parse(data) : data;
                    } catch(e) {
                        Toast.error('Erro ao processar resposta do servidor.');
                        console.error('Parse error:', e, data);
                        return;
                    }

                    if (!response.sucesso) {
                        Toast.error('Erro ao pesquisar: ' + response.mensagem);
                        return;
                    }

                    dadosRelatorio = response.dados || [];
                    controlarAvisoTempoReal();
                    renderizarDados();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    var mensagem = 'Erro ao pesquisar.';
                    try {
                        var response = JSON.parse(xhr.responseText);
                        if (response && response.mensagem) mensagem = response.mensagem;
                    } catch(e) {}
                    Toast.error(mensagem);
                    console.error('Erro ao pesquisar:', error);
                }
            });
        }

        function renderizarDados() {
            const valorTotalText = dadosRelatorio.length <= 0 ? 0 : formatarMoeda(somarValores(dadosRelatorio, 'valorTotal', 2));
            const valorSemIpiText = dadosRelatorio.length <= 0 ? 0 : formatarMoeda(somarValores(dadosRelatorio, 'valorSemIpi', 2));
            const valorNetText = dadosRelatorio.length <= 0 ? 0 : formatarMoeda(somarValores(dadosRelatorio, 'valorNet', 2));
            const quantidadeText = dadosRelatorio.length <= 0 ? 0 : formatarMoeda(somarValores(dadosRelatorio, 'quantidade'), 2);
            const areaText = dadosRelatorio.length <= 0 ? 0 : formatarMoeda(somarValores(dadosRelatorio, 'area'), 2);
            const areaRefileText = dadosRelatorio.length <= 0 ? 0 : formatarMoeda(somarValores(dadosRelatorio, 'areaRefile'), 2);
            const pesoText = dadosRelatorio.length <= 0 ? 0 : formatarMoeda(somarValores(dadosRelatorio, 'peso'), 2);
            const pesoRefileText = dadosRelatorio.length <= 0 ? 0 : formatarMoeda(somarValores(dadosRelatorio, 'pesoRefile'), 2);
            
            const areaTotal = somarValores(dadosRelatorio, 'area');
            const areaRefileTotal = somarValores(dadosRelatorio, 'areaRefile');

            const pesoTotal = somarValores(dadosRelatorio, 'peso');
            const pesoRefileTotal = somarValores(dadosRelatorio, 'pesoRefile');

            const gramaturaText = (dadosRelatorio.length <= 0) ? 0 : formatarMoeda((pesoTotal / areaTotal) * 1000);
            const gramaturaRefileText = (dadosRelatorio.length <= 0) ? 0 : formatarMoeda((pesoRefileTotal / areaRefileTotal) * 1000);
            
            const valorNetTotal = somarValores(dadosRelatorio, 'valorNet');
            const precoKGText = (dadosRelatorio.length <= 0) ? 0 : formatarMoeda(valorNetTotal / pesoTotal);
            const precoM2Text = (dadosRelatorio.length <= 0) ? 0 : formatarMoeda(valorNetTotal / areaTotal);

            const $tbody = $('#gradeRelatorio tbody');
            const $tfoot = $('#gradeRelatorio tfoot');
            
            $tbody.empty();
            $tfoot.empty();
            
            if (dadosRelatorio.length === 0) {
                // Cartões
                $('#valorTotal').text('0');
                $('#valorSemIpi').text('0');
                $('#valorNet').text('0');
                $('#area').text('0');
                $('#areaRefile').text('0');
                $('#peso').text('0');
                $('#pesoRefile').text('0');
                $('#gramatura').text('0');
                $('#gramaturaRefile').text('0');
                $('#precoKG').text('0');
                $('#precoM2').text('0');

                // Tabela
                $tbody.html(`
                    <tr>
                        <td colspan="33" class="text-center text-muted py-4">
                            Utilize os filtros acima para pesquisar
                        </td>
                    </tr>
                `);
                return;
            }

            // Cartões
            const deveMostrarCardsPapelao = $('#filtroGrupoProduto').val() == 'PAPELAO';
            if(deveMostrarCardsPapelao) {
                $('#cardValores').removeClass().addClass('col-12 col-lg-4 col-md-6');
                $('#cardArea, #cardPeso, #cardGramatura, #cardPreco').removeClass().addClass('col-12 col-lg-2 col-md-6');
            } else {
                $('#cardValores, #cardArea, #cardPeso').removeClass().addClass('col-12 col-lg-4 col-md-6');
                $('#cardGramatura, #cardPreco').removeClass().addClass('d-none');
            }

            $('#valorTotal').text(valorTotalText);
            $('#valorSemIpi').text(valorSemIpiText);
            $('#valorNet').text(valorNetText);
            $('#area').text(areaText);
            $('#areaRefile').text(areaRefileText);
            $('#peso').text(pesoText);
            $('#pesoRefile').text(pesoRefileText);
            $('#gramatura').text(gramaturaText);
            $('#gramaturaRefile').text(gramaturaRefileText);
            $('#precoKG').text(precoKGText);
            $('#precoM2').text(precoM2Text);

            dadosRelatorio.forEach(function(item) {
                $tbody.append(`
                    <tr>
                        <td>${item.dataEmissao || '-'}</td>
                        <td>${item.pedidoMae || '-'}</td>
                        <td>${item.numeroPedido || '-'}</td>
                        <td>${item.notaFiscal || '-'}</td>
                        <td>${item.itemNF || '-'}</td>
                        <td>${item.tes || '-'}</td>
                        <td>${item.romaneio || '-'}</td>
                        <td>${item.nomeTransportadora || '-'}</td>
                        <td>${item.placaVeiculo || '-'}</td>
                        <td>${item.tipoFrete || '-'}</td>
                        <td>${item.codCliente || '-'}</td>
                        <td>${item.codLoja || '-'}</td>
                        <td>${item.razaoSocial || '-'}</td>
                        <td>${item.cidadeCliente + '-' + item.estadoCliente}</td>
                        <td>${item.numeroPedidoCliente + ' - ' + item.itemPedidoCliente}</td>
                        <td>${item.codProduto || '-'}</td>
                        <td>${item.descricaoProduto || '-'}</td>
                        <td>${item.estruturaProduto || '-'}</td>
                        <td>${item.maquina || '-'}</td>
                        <td>${item.grupoProduto || '-'}</td>
                        <td>${item.tipoProduto || '-'}</td>
                        <td>${item.codVendedor + ' - ' + item.nomeVendedor}</td>
                        <td>${formatarMoeda(item.quantidade, 0)}</td>
                        <td>${formatarMoeda(item.valorUnitario, 5)}</td>
                        <td>${formatarMoeda(item.area, 5)}</td>
                        <td>${formatarMoeda(item.areaRefile, 5)}</td>
                        <td>${formatarMoeda(item.peso, 5)}</td>
                        <td>${formatarMoeda(item.pesoRefile, 5)}</td>
                        <td>${formatarMoeda(item.valorTotal, 5)}</td>
                        <td>${formatarMoeda(item.valorSemIpi, 5)}</td>
                        <td>${formatarMoeda(item.valorNet, 5)}</td>
                        <td>${formatarMoeda(item.precoKg, 5)}</td>
                        <td>${formatarMoeda(item.precoM2, 5)}</td>
                    </tr>
                `);
            });

            $tfoot.html(`
                <tr>
                    <td>Totais</td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td>${quantidadeText}</td>
                    <td></td>
                    <td>${areaText}</td>
                    <td>${areaRefileText}</td>
                    <td>${pesoText}</td>
                    <td>${pesoRefileText}</td>
                    <td>${valorTotalText}</td>
                    <td>${valorSemIpiText}</td>
                    <td>${valorNetText}</td>
                    <td></td>
                    <td></td>
                </tr>
            `)
        }

        function limparFiltros() {
            $('#filtroFilial').val('').trigger('change');
            $('#filtroDataInclusaoDe').val('');
            $('#filtroDataInclusaoAte').val('');
            $('#filtroCodCliente').val('');
            $('#nomeCliente').val('').trigger('change');
            $('#filtroCodLoja').val('');
            $('#filtroRepresentante').val('').trigger('change');
            $('#filtroClienteGrupo').val('').trigger('change');
            $('#filtroGrupoProduto').val('').trigger('change');
            $('#filtroTipoProduto').val('').trigger('change');
            $('#filtroCodProduto').val('');
            $('#filtroDescProduto').val('');
            $('#filtroTipoNotaFiscal').val('').trigger('change');
            $('#filtroPedido').val('');
            $('#filtroNotaFiscal').val('');
            
            dadosRelatorio = [];
            renderizarDados();
        }

        function exportarPDF() {
            if (dadosRelatorio.length === 0) {
                Toast.error('Não há dados para exportar');
                return;
            }

            const { jsPDF } = window.jspdf;

            var doc = new jsPDF({
                orientation: 'landscape', 
                unit: 'mm',
                format: 'a0'
            });

            const now = new Date();

            const options = {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit',
                hour12: false 
            };

            const dataAtual = new Intl.DateTimeFormat('pt-BR', options).format(now);

            const colunas = [
                'Data de Emissão', 
                'Pedido Mãe',
                'Nº Pedido',
                'Nota Fiscal',
                'Item NF',
                'TES',
                'Romaneio',
                'Nome Transportadora',
                'Placa Veículo',
                'Tipo Frete',
                'Cód. Cliente',
                'Cód. Loja',
                'Razão Social',
                'Cidade/Estado',
                'Ped. Cliente',
                'Cód. Produto',
                'Descrição',
                'Estrutura',
                'Máquina Roteiro',
                'Grupo',
                'Tipo',
                'Vendedor',
                'Quantidade',
                'Valor Unit.',
                'Área',
                'Área Refile',
                'Peso',
                'Peso Refile',
                'Valor Total',
                'Valor Sem IPI',
                'Valor NET',
                'Preço KG',
                'Preço M2'
            ];

            const dados = dadosRelatorio.map(item => [
                item.dataEmissao || '-',
                item.pedidoMae || '-',
                item.numeroPedido || '-',
                item.notaFiscal || '-',
                item.itemNF || '-',
                item.tes || '-',
                item.romaneio || '-',
                item.nomeTransportadora || '-',
                item.placaVeiculo || '-',
                item.tipoFrete || '-',
                item.codCliente || '-',
                item.codLoja || '-',
                item.razaoSocial || '-',
                item.cidadeCliente + '-' + item.estadoCliente,
                item.numeroPedidoCliente + '-' + item.itemPedidoCliente,
                item.codProduto || '-',
                item.descricaoProduto || '-',
                item.estruturaProduto || '-',
                item.maquina || '-',
                item.grupoProduto || '-',
                item.tipoProduto || '-',
                item.codVendedor + ' - ' + item.nomeVendedor,
                formatarMoeda(item.quantidade, 0),
                formatarMoeda(item.valorUnitario, 5),
                formatarMoeda(item.area, 5),
                formatarMoeda(item.areaRefile, 5),
                formatarMoeda(item.peso, 5),
                formatarMoeda(item.pesoRefile, 5),
                formatarMoeda(item.valorTotal, 5),
                formatarMoeda(item.valorSemIpi, 5),
                formatarMoeda(item.valorNet, 5),
                formatarMoeda(item.precoKg, 5),
                formatarMoeda(item.precoM2, 5)
            ]);

            var yPosition = 15;
            let filtrosAplicados = [];
                
            const filtroDataInclusaoDe = $('#filtroDataInclusaoDe').val();
            const filtroDataInclusaoAte = $('#filtroDataInclusaoAte').val();
            const filtroFilial = $('#filtroFilial').val() && $('#filtroFilial option:selected').text();
            const filtroCodCliente = $('#filtroCodCliente').val();
            const filtroCodLoja = $('#filtroCodLoja').val();
            const filtroRepresentante = $('#filtroRepresentante').val() && $('#filtroRepresentante option:selected').text();
            const filtroClienteGrupo = $('#filtroClienteGrupo').val() && $('#filtroClienteGrupo option:selected').text();
            const filtroGrupoProduto = $('#filtroGrupoProduto').val() && $('#filtroGrupoProduto option:selected').text();
            const filtroTipoProduto = $('#filtroTipoProduto').val() && $('#filtroTipoProduto option:selected').text();
            const filtroCodProduto = $('#filtroCodProduto').val();
            const filtroDescProduto = $('#filtroDescProduto').val();
            const filtroTipoNotaFiscal = $('#filtroTipoNotaFiscal').val() && $('#filtroTipoNotaFiscal option:selected').text();
            const filtroPedido = $('#filtroPedido').val();
            const filtroNotaFiscal = $('#filtroNotaFiscal').val();

            if (filtroDataInclusaoDe) 
                filtrosAplicados.push('Data Inclusão (De): ' + new Date(`${filtroDataInclusaoDe} 03:00`).toLocaleDateString('pt-BR'));

            if (filtroDataInclusaoAte) 
                filtrosAplicados.push('Data Inclusão (Até): ' + new Date(`${filtroDataInclusaoAte} 03:00`).toLocaleDateString('pt-BR') );

            if (filtroFilial) 
                filtrosAplicados.push('Filial: ' + filtroFilial);

            if (filtroCodCliente) 
                filtrosAplicados.push('Cód. Cliente: ' + filtroCodCliente);

            if (filtroCodLoja) 
                filtrosAplicados.push('Cód. Loja: ' + filtroCodLoja);

            if (filtroRepresentante) 
                filtrosAplicados.push('Representante: ' + filtroRepresentante);

            if (filtroClienteGrupo) 
                filtrosAplicados.push('Cliente do Grupo: ' + filtroClienteGrupo);

            if (filtroGrupoProduto) 
                filtrosAplicados.push('Grupo Produto: ' + filtroGrupoProduto);

            if (filtroTipoProduto) 
                filtrosAplicados.push('Tipo Produto: ' + filtroTipoProduto);

            if (filtroCodProduto) {
                const labelProduto = filtroDescProduto ? filtroCodProduto + ' - ' + filtroDescProduto : filtroCodProduto;
                filtrosAplicados.push('Produto: ' + labelProduto);
            }

            if (filtroTipoNotaFiscal) 
                filtrosAplicados.push('Tipo Nota Fiscal: ' + filtroTipoNotaFiscal);

            if (filtroPedido) 
                filtrosAplicados.push('Pedido: ' + filtroPedido);

            if (filtroNotaFiscal) 
                filtrosAplicados.push('Nota Fiscal: ' + filtroNotaFiscal);

            if (filtrosAplicados.length > 0) {
                doc.text('Filtros Aplicados:', 14, yPosition);
                yPosition += 7;
                doc.setFont(undefined, 'normal');

                for (var i = 0; i < filtrosAplicados.length; i++) {
                    doc.text('  • ' + filtrosAplicados[i], 14, yPosition);
                    yPosition += 7;
                }

                yPosition += 2;
            }

            doc.setFontSize(16);
            doc.text('Relatório de Faturamento Detalhado', 14, yPosition);
            yPosition += 7;

            doc.setFontSize(10);
            doc.text(`Data: ${dataAtual} | ${dados.length} registro(s) encontrado(s)`, 14, yPosition);
            yPosition += 7;

            // Linha de totais para o rodapé do PDF
            const totalQtd    = somarValores(dadosRelatorio, 'quantidade');
            const totalArea   = somarValores(dadosRelatorio, 'area');
            const totalARefi  = somarValores(dadosRelatorio, 'areaRefile');
            const totalPeso   = somarValores(dadosRelatorio, 'peso');
            const totalPRefi  = somarValores(dadosRelatorio, 'pesoRefile');
            const totalVTotal = somarValores(dadosRelatorio, 'valorTotal');
            const totalVIpi   = somarValores(dadosRelatorio, 'valorSemIpi');
            const totalVNet   = somarValores(dadosRelatorio, 'valorNet');

            const linhaTotais = [
                'TOTAIS', '', '', '', '', '', '', '', '', '', '', '',
                '', '', '', '', '', '', '', '', '', '',
                formatarMoeda(totalQtd, 0),
                '',
                formatarMoeda(totalArea, 2),
                formatarMoeda(totalARefi, 2),
                formatarMoeda(totalPeso, 2),
                formatarMoeda(totalPRefi, 2),
                formatarMoeda(totalVTotal, 2),
                formatarMoeda(totalVIpi, 2),
                formatarMoeda(totalVNet, 2),
                '', ''
            ];

            doc.autoTable({
                head: [colunas],
                body: dados,
                foot: [linhaTotais],
                startY: yPosition,
                theme: 'grid',
                styles: { 
                    fontSize: 8,
                    cellPadding: 2
                },
                headStyles: { 
                    fillColor: [211, 211, 212],
                    textColor: [0, 0, 0],
                    fontStyle: 'bold'
                },
                footStyles: {
                    fillColor: [74, 74, 74],
                    textColor: [255, 255, 255],
                    fontStyle: 'bold'
                },
                showFoot: 'lastPage',
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
            if (dadosRelatorio.length === 0) {
                Toast.error('Não há dados para exportar');
                return;
            }

            const nomeArquivo = 'Faturamento_Detalhado.csv';
            const colunas = [
                'Data de Emissão', 
                'Pedido Mãe',
                'Nº Pedido',
                'Nota Fiscal',
                'Item NF',
                'TES',
                'Romaneio',
                'Nome Transportadora',
                'Placa Veículo',
                'Tipo Frete',
                'Cód. Cliente',
                'Cód. Loja',
                'Razão Social',
                'Cidade/Estado',
                'Ped. Cliente',
                'Cód. Produto',
                'Descrição',
                'Estrutura',
                'Máquina Roteiro',
                'Grupo',
                'Tipo',
                'Vendedor',
                'Quantidade',
                'Valor Unit.',
                'Área',
                'Área Refile',
                'Peso',
                'Peso Refile',
                'Valor Total',
                'Valor Sem IPI',
                'Valor NET',
                'Preço KG',
                'Preço M2'
            ];

            let csvContent = '\uFEFF'; 
            csvContent += colunas.join(';') + '\r\n';

            dadosRelatorio.forEach(item => {
                const linha = [
                    escapeCSV(item.dataEmissao || '-'),
                    escapeCSV(item.pedidoMae || '-'),
                    escapeCSV(item.numeroPedido || '-'),
                    escapeCSV(item.notaFiscal || '-'),
                    escapeCSV(item.itemNF || '-'),
                    escapeCSV(item.tes || '-'),
                    escapeCSV(item.romaneio || '-'),
                    escapeCSV(item.nomeTransportadora || '-'),
                    escapeCSV(item.placaVeiculo || '-'),
                    escapeCSV(item.tipoFrete || '-'),
                    escapeCSV(item.codCliente || '-'),
                    escapeCSV(item.codLoja || '-'),
                    escapeCSV(item.razaoSocial || '-'),
                    escapeCSV(item.cidadeCliente + '-' + item.estadoCliente),
                    escapeCSV(item.numeroPedidoCliente + '-' + item.itemPedidoCliente),
                    escapeCSV(item.codProduto || '-'),
                    escapeCSV(item.descricaoProduto || '-'),
                    escapeCSV(item.estruturaProduto || '-'),
                    escapeCSV(item.maquina || '-'),
                    escapeCSV(item.grupoProduto || '-'),
                    escapeCSV(item.tipo || '-'),
                    escapeCSV(item.codVendedor + ' - ' + item.nomeVendedor),
                    escapeCSV(formatarMoeda(item.quantidade, 0)),
                    escapeCSV(formatarMoeda(item.valorUnitario, 5)),
                    escapeCSV(formatarMoeda(item.area, 5)),
                    escapeCSV(formatarMoeda(item.areaRefile, 5)),
                    escapeCSV(formatarMoeda(item.peso, 5)),
                    escapeCSV(formatarMoeda(item.pesoRefile, 5)),
                    escapeCSV(formatarMoeda(item.valorTotal, 5)),
                    escapeCSV(formatarMoeda(item.valorSemIpi, 5)),
                    escapeCSV(formatarMoeda(item.valorNet, 5)),
                    escapeCSV(formatarMoeda(item.precoKg, 5)),
                    escapeCSV(formatarMoeda(item.precoM2, 5))
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

        function abrirModalProduto() {
            $('#inputBuscaModalProduto').val('');
            $('#tabelaModalProduto').html('<tr><td colspan="3" class="text-center text-muted py-3">Digite para buscar</td></tr>');
            var modal = new bootstrap.Modal(document.getElementById('modalBuscaProduto'));
            modal.show();
            setTimeout(function() { $('#inputBuscaModalProduto').focus(); }, 400);
        }

        function buscarProdutosModal(termo) {
            const $tbody = $('#tabelaModalProduto');
            const $btn   = $('#btnBuscarModalProduto');

            if (!termo || termo.length < 1) {
                $tbody.html('<tr><td colspan="3" class="text-center text-muted py-3">Digite para buscar</td></tr>');
                return;
            }

            $btn.prop('disabled', true).html('<span class="spinner-border spinner-border-sm me-1" role="status"></span> Buscando...');
            $tbody.html('<tr><td colspan="3" class="text-center py-3"><div class="spinner-border spinner-border-sm text-primary" role="status"></div></td></tr>');

            $.ajax({
                url: 'RelatorioFaturamentoDetalhadoRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: {
                    acao: 'buscarProdutos',
                    filtroCodProduto: termo,
                    filtroFilial: $('#filtroFilial').val()
                },
                success: function(data) {
                    $btn.prop('disabled', false).html('<i class="bi bi-search me-1"></i> Buscar');
                    $tbody.empty();

                    var response = (typeof data === 'string') ? JSON.parse(data) : data;

                    if (!response.sucesso || !response.dados || response.dados.length === 0) {
                        $tbody.html('<tr><td colspan="3" class="text-center text-muted py-3">Nenhum produto encontrado</td></tr>');
                        return;
                    }
                    response.dados.forEach(function(item) {
                        var $tr = $('<tr>').html(
                            '<td class="text-center align-middle">' +
                                '<button type="button" class="btn btn-success btn-sm d-flex align-items-center justify-content-center p-0" ' +
                                'style="width:28px;height:28px;border-radius:6px;" title="Selecionar produto">' +
                                    '<i class="bi bi-check-lg text-white"></i>' +
                                '</button>' +
                            '</td>' +
                            '<td class="align-middle text-center">' + (item.codProduto || '-') + '</td>' +
                            '<td class="align-middle text-center">' + (item.descricaoProduto || '-') + '</td>'
                        );
                        $tr.find('button').on('click', function() {
                            selecionarProduto(item.codProduto, item.descricaoProduto);
                        });
                        $tbody.append($tr);
                    });
                },
                error: function(xhr) {
                    $btn.prop('disabled', false).html('<i class="bi bi-search me-1"></i> Buscar');
                    try {
                        var response = JSON.parse(xhr.responseText);
                        if (response.sucesso && response.dados && response.dados.length > 0) {
                            $tbody.empty();
                            response.dados.forEach(function(item) {
                                var $tr = $('<tr>').html(
                                    '<td class="text-center align-middle">' +
                                        '<button type="button" class="btn btn-success btn-sm d-flex align-items-center justify-content-center p-0" ' +
                                        'style="width:28px;height:28px;border-radius:6px;" title="Selecionar produto">' +
                                            '<i class="bi bi-check-lg text-white"></i>' +
                                        '</button>' +
                                    '</td>' +
                                    '<td class="align-middle text-center">' + (item.codProduto || '-') + '</td>' +
                                    '<td class="align-middle text-center">' + (item.descricaoProduto || '-') + '</td>'
                                );
                                $tr.find('button').on('click', function() {
                                    selecionarProduto(item.codProduto, item.descricaoProduto);
                                });
                                $tbody.append($tr);
                            });
                            return;
                        }
                    } catch(e) {}
                    $tbody.html('<tr><td colspan="3" class="text-center text-danger py-3">Erro ao buscar produtos</td></tr>');
                }
            });
        }

        function selecionarProduto(cod, desc) {
            $('#filtroCodProduto').val(cod);
            $('#filtroDescProduto').val(desc);
            bootstrap.Modal.getInstance(document.getElementById('modalBuscaProduto')).hide();
        }

        function limparSelecaoProduto() {
            $('#filtroCodProduto').val('');
            $('#filtroDescProduto').val('');
        }

        function anexarEventos() {
            $('#formularioFiltro').on('submit', validarFiltros);
            $('#btnLimpar').on('click', limparFiltros);
            $('#btnExportarExcel').on('click', exportarExcel);
            $('#btnExportarPDF').on('click', exportarPDF);
            $('#filtroFilial').on('change', carregarRepresentantes);

            $('#filtroCodProduto').on('input', function() {
                $('#filtroDescProduto').val('');
            });

            $(document).on('keydown', '#inputBuscaModalProduto', function(e) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    buscarProdutosModal($(this).val().trim());
                }
            });

            // Observa o campo hidden #nomeCliente preenchido pelo buscaCLIENTE.asp
            // e exibe a razão social no <small> abaixo do campo de código
            var nomeClienteEl = document.getElementById('nomeCliente');
            if (nomeClienteEl) {
                new MutationObserver(function() {
                    var nome = nomeClienteEl.value;
                    var $exibir = $('#nomeClienteExibir');
                    if (nome) {
                        $exibir.text(nome).removeClass('text-muted').addClass('text-success fw-semibold');
                    } else {
                        $exibir.text('Clique no botão para buscar ou digite o código').removeClass('text-success fw-semibold').addClass('text-muted');
                    }
                }).observe(nomeClienteEl, { attributes: true, attributeFilter: ['value'] });

                // jQuery change também captura setters via .val()
                $('#nomeCliente').on('change input', function() {
                    var nome = $(this).val();
                    var $exibir = $('#nomeClienteExibir');
                    if (nome) {
                        $exibir.text(nome).removeClass('text-muted').addClass('text-success fw-semibold');
                    } else {
                        $exibir.text('Clique no botão para buscar ou digite o código').removeClass('text-success fw-semibold').addClass('text-muted');
                    }
                });
            }
        }

        function controlarAvisoTempoReal() {
            const flagTempoReal = $('#flagTempoReal').is(':checked');

            if (!flagTempoReal) {
                $('#avisoTempoReal').hide();
                return;
            }
            
            $('#avisoTempoReal').html(`
                <div class="alert alert-warning alert-dismissible py-2 px-3 mb-0 d-flex align-items-center" role="alert" style="font-size: 0.85rem;">
                    <i class="bi bi-info-circle me-2"></i>
                    <span>Buscando dados em tempo real. Apenas dados do mês atual disponíveis para este tipo de busca.</span>
                </div>
            `).show();
        }
    </script>
</head>
<body>
    <div class="container my-4">
        <h3 class="mb-4">Relatório de Faturamento Detalhado</h3>

        <form name="formularioFiltro" id="formularioFiltro" method="post">                              
            <div class="card p-3 mb-4">
                <div class="row g-3">
                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroFilial" class="form-label">Filial</label>
                        <select required class="form-select form-select-sm select2" id="filtroFilial" name="filtroFilial" data-placeholder="Todas as filiais">
                            <option value="">Todas as filiais</option>
                            <option value="PIRAPETINGA">Pirapetinga</option>
                            <option value="UBERABA">Uberaba</option>
                            <option value="BENTO">Bento</option>
                            <option value="MARANGUAPE">Maranguape</option>
                            <option value="FORTALEZA">Fortaleza</option>
                            <option value="SAQUAREMA">Saquarema</option>
                            <option value="ESHOP">Eshop</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroDataInclusaoDe" class="form-label">Data Inclusão De</label>
                        <input required type="date" class="form-control form-control-sm" id="filtroDataInclusaoDe" name="filtroDataInclusaoDe">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroDataInclusaoAte" class="form-label">Data Inclusão Até</label>
                        <input required type="date" class="form-control form-control-sm" id="filtroDataInclusaoAte" name="filtroDataInclusaoAte">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroCodCliente" class="form-label">Cód. Cliente</label>
                        <div class="input-group input-group-sm">
                            <button type="button" class="btn btn-sm btn-secondary" onclick="pesquisarCliente()">
                                Buscar &nbsp;<i class="bi bi-box-arrow-up-right"></i>
                            </button>
                            <input type="text" class="form-control form-control-sm" id="filtroCodCliente" name="filtroCodCliente" placeholder="Código do cliente">
                        </div>
                        <input type="hidden" id="nomeCliente">
                        <small class="text-muted" id="nomeClienteExibir" style="display:block; min-height:1.2em;">Clique no botão para buscar ou digite o código</small>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroCodLoja" class="form-label">Cód. Loja</label>
                        <input type="text" class="form-control form-control-sm" id="filtroCodLoja" name="filtroCodLoja" placeholder="Código da loja">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroRepresentante" class="form-label">Representante</label>
                        <select class="form-select form-select-sm select2" id="filtroRepresentante" name="filtroRepresentante" data-placeholder="Todos os representantes">
                            <option value="">Todos os representantes</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroClienteGrupo" class="form-label">Cliente do Grupo</label>
                        <select class="form-select form-select-sm select2" id="filtroClienteGrupo" name="filtroClienteGrupo" data-placeholder="Todos os tipos">
                            <option value="">Todos</option>
                            <option value="S">Sim</option>
                            <option value="N">Não</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroGrupoProduto" class="form-label">Grupo de Produto</label>
                        <select class="form-select form-select-sm select2" id="filtroGrupoProduto" name="filtroGrupoProduto" data-placeholder="Todos os grupos">
                            <option value="">Todos</option>
                            <option value="PAPEL">Papel</option>
                            <option value="PAPELAO">Papelão</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroTipoProduto" class="form-label">Tipo de Produto</label>
                        <select class="form-select form-select-sm select2" id="filtroTipoProduto" name="filtroTipoProduto" data-placeholder="Todos os tipos">
                            <option value="">Todos</option>
                            <option value="CAIXA">Caixa</option>
                            <option value="CHAPA">Chapa</option>
                            <option value="CORTE VINCO">Corte Vinco</option>
                            <option value="MALETA">Maleta</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroCodProduto" class="form-label">Buscar Produto</label>
                        <div class="input-group input-group-sm">
                            <button type="button" class="btn btn-sm btn-secondary" onclick="abrirModalProduto()" title="Buscar produto">
                                Buscar &nbsp;<i class="bi bi-box-arrow-up-right"></i>
                            </button>
                            <input type="text" class="form-control form-control-sm" id="filtroCodProduto" name="filtroCodProduto" placeholder="Cód. ou descrição">
                            <button type="button" class="btn btn-sm btn-outline-secondary" onclick="limparSelecaoProduto()" title="Limpar produto"><i class="bi bi-x"></i></button>
                        </div>
                        <small class="text-muted" id="filtroDescProduto" style="display:block; min-height:1.2em;"></small>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroTipoNotaFiscal" class="form-label">Tipo de Nota Fiscal</label>
                        <select class="form-select form-select-sm select2" id="filtroTipoNotaFiscal" name="filtroTipoNotaFiscal" data-placeholder="Todos os tipos">
                            <option value="">Todos</option>
                            <option value="VENDA">Sem Devolução</option>
                            <option value="DEVOLUCAO">Somente Devolução</option>
                            <option value="REMESSA">Somente Remessa</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroPedido" class="form-label">Pedido</label>
                        <input class="form-control form-control-sm" id="filtroPedido" name="filtroPedido" placeholder="Pedido">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroNotaFiscal" class="form-label">Nota Fiscal</label>
                        <input class="form-control form-control-sm" id="filtroNotaFiscal" name="filtroNotaFiscal" placeholder="Nota Fiscal">
                    </div> 

                    <div class="col-12 d-flex justify-content-between align-items-end mt-4">
                        <div class="switch-item">
                            <div class="">
                                <label class="switch-label" for="flagTempoReal">Mostrar dados em tempo real</label>
                                <div class="text-muted">Inclui apenas dados do mês atual</div>
                            </div>
                            <div class="form-check form-switch mb-0">
                                <input class="form-check-input" type="checkbox" id="flagTempoReal"/>
                            </div>
                        </div>

                        <div class="d-flex gap-2">
                            <button type="button" class="btn btn-sm btn-outline-secondary" id="btnLimpar">
                                <i class="bi bi-x-lg me-1"></i> Limpar
                            </button>
                            <button type="submit" type="button" class="btn btn-sm btn-primary" id="btnPesquisar">
                                <i class="bi bi-search me-1"></i> Pesquisar
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </form>

        <div class="row g-3 mb-4">
            <div id="cardValores" class="col-12 col-lg-4 col-md-6">
                <div class="card shadow-sm border-0 h-100">
                    <div class="card-body d-flex justify-content-between align-items-start">
                        <div class="w-100 ">                    
                            <div class="d-flex gap-3 justify-content-between">
                                <div>
                                    <div class="text-secondary mb-1" style="font-size:0.75rem;">Valor Net</div>
                                    <h6 id="valorNet" class="fw-bold mb-0 text-body-secondary">0</h6>
                                </div>

                                <div class="icon-wrapper light-blue">
                                    <i class="bi bi-bar-chart-line"></i>
                                </div>
                            </div>

                            <div class="horizontal-line my-2"></div>
                    
                            <div class="d-flex gap-3 align-items-stretch">
                                <div class="flex-fill">
                                    <div class="text-secondary mb-1" style="font-size:0.75rem;">Valor Total</div>
                                    <h6 id="valorTotal" class="fw-bold mb-0 text-body-secondary">0</h6>
                                </div>

                                <div class="vertical-line"></div>

                                <div class="flex-fill">
                                    <div class="text-secondary mb-1" style="font-size:0.75rem;">Valor Sem IPI</div>
                                    <h6  id="valorSemIpi" class="fw-bold mb-0 text-body-secondary">0</h6>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <div id="cardArea" class="col-12 col-lg-4 col-md-6">
                <div class="card shadow-sm border-0 h-100">
                    <div class="card-body d-flex justify-content-between align-items-start">
                        <div class="w-100 ">         
                            <div class="d-flex gap-3 justify-content-between">
                                <div>
                                    <div class="text-secondary mb-1" style="font-size:0.75rem;">Área</div>
                                    <h6 id="area" class="fw-bold mb-0 text-body-secondary">0</h6>
                                </div>
                                <div class="icon-wrapper blue">
                                    <i class="bi bi-arrows-fullscreen"></i>
                                </div>
                            </div>

                            <div class="horizontal-line my-2"></div>

                            <div class="flex-fill">
                                <div class="text-secondary mb-1" style="font-size:0.75rem;">Área Refile</div>
                                <h6 id="areaRefile" class="fw-bold mb-0 text-body-secondary">0</h6>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <div id="cardPeso" class="col-12 col-lg-4 col-md-6">
                <div class="card shadow-sm border-0 h-100">
                    <div class="card-body d-flex justify-content-between align-items-start">
                        <div class="w-100 ">         
                            <div class="d-flex gap-3 justify-content-between">
                                <div>
                                    <div class="text-secondary mb-1" style="font-size:0.75rem;">Peso</div>
                                    <h6 id="peso" class="fw-bold mb-0 text-body-secondary">0</h6>
                                </div>
                                <div class="icon-wrapper yellow">
                                    <i class="bi bi-box-seam"></i>
                                </div>
                            </div>

                            <div class="horizontal-line my-2"></div>

                            <div class="flex-fill">
                                <div class="text-secondary mb-1" style="font-size:0.75rem;">Peso Refile</div>
                                <h6 id="pesoRefile" class="fw-bold mb-0 text-body-secondary">0</h6>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            
            <div id="cardGramatura" class="d-none col-12 col-lg-4 col-md-6">
                <div class="card shadow-sm border-0 h-100">
                    <div class="card-body d-flex justify-content-between align-items-start">
                        <div class="w-100 ">         
                            <div class="d-flex gap-3 justify-content-between">
                                <div>
                                    <div class="text-secondary mb-1" style="font-size:0.75rem;">Gramatura Média</div>
                                    <h6 id="gramatura" class="fw-bold mb-0 text-body-secondary">0</h6>
                                </div>
                                <div class="icon-wrapper purple">
                                    <i class="bi bi-layers"></i>
                                </div>
                            </div>

                            <div class="horizontal-line my-2"></div>

                            <div class="flex-fill">
                                <div class="text-secondary mb-1" style="font-size:0.75rem;">Gramatura Média Refile</div>
                                <h6 id="gramaturaRefile" class="fw-bold mb-0 text-body-secondary">0</h6>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <div id="cardPreco" class="d-none col-12 col-lg-4 col-md-6">
                <div class="card shadow-sm border-0 h-100">
                    <div class="card-body d-flex justify-content-between align-items-start">
                        <div class="w-100 ">         
                            <div class="d-flex gap-3 justify-content-between">
                                <div>
                                    <div class="text-secondary mb-1" style="font-size:0.75rem;">Preço KG</div>
                                    <h6 id="precoKG" class="fw-bold mb-0 text-body-secondary">0</h6>
                                </div>
                                <div class="icon-wrapper green">
                                    <i class="bi bi-cash"></i>
                                </div>
                            </div>

                            <div class="horizontal-line my-2"></div>

                            <div class="flex-fill">
                                <div class="text-secondary mb-1" style="font-size:0.75rem;">Preço M2</div>
                                <h6 id="precoM2" class="fw-bold mb-0 text-body-secondary">0</h6>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div id="avisoTempoReal" class="w-100 mb-4" style="display: none;"></div>

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

                <table class="table table-hover mb-0" id="gradeRelatorio">
                    <thead>
                        <tr>
                            <th>Data de Emissão</th> 
                            <th>Pedido Mãe</th>
                            <th>Nº Pedido</th>
                            <th>Nota Fiscal</th>
                            <th>Item NF</th>
                            <th>TES</th>
                            <th>Romaneio</th>
                            <th>Nome Transportadora</th>
                            <th>Placa Veículo</th>
                            <th>Tipo Frete</th>
                            <th>Cód. Cliente</th>
                            <th>Cód. Loja</th>
                            <th>Razão Social</th>
                            <th>Cidade/Estado</th>
                            <th>Ped. Cliente</th>
                            <th>Cód. Produto</th>
                            <th>Descrição</th>
                            <th>Estrutura</th>
                            <th>Máquina Roteiro</th>
                            <th>Grupo</th>
                            <th>Tipo</th>
                            <th>Vendedor</th>
                            <th>Quantidade</th>
                            <th>Valor Unit.</th>
                            <th>Área</th>
                            <th>Área Refile</th>
                            <th>Peso</th>
                            <th>Peso Refile</th>
                            <th>Valor Total</th>
                            <th>Valor Sem IPI</th>
                            <th>Valor NET</th>
                            <th>Preço KG</th>
                            <th>Preço M2</th>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td colspan="33" class="text-center text-muted py-4">
                                Nenhum registro encontrado
                            </td>
                        </tr>
                    </tbody>
                    <tfoot>
                    </tfoot>
                </table>
            </div>
        </div>
    </div>
    <!-- Modal Busca Produto -->
    <div class="modal fade" id="modalBuscaProduto" tabindex="-1" aria-labelledby="modalBuscaProdutoLabel" aria-hidden="true">
        <div class="modal-dialog modal-lg modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="modalBuscaProdutoLabel">Buscar Produto</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Fechar"></button>
                </div>
                <div class="modal-body">
                    <div class="mb-3">
                        <div class="input-group input-group-sm">
                            <input type="text" class="form-control form-control-sm" id="inputBuscaModalProduto"
                                   placeholder="Digite o código ou descrição do produto...">
                            <button type="button" class="btn btn-primary" id="btnBuscarModalProduto" onclick="buscarProdutosModal($('#inputBuscaModalProduto').val().trim())">
                                <i class="bi bi-search me-1"></i> Buscar
                            </button>
                        </div>
                    </div>
                    <div class="table-responsive" style="max-height: 400px; overflow-y: auto;">
                        <table class="table table-hover table-striped table-sm mb-0">
                            <thead style="position: sticky; top: 0; background: #fff; z-index: 1;">
                                <tr>
                                    <th style="width: 48px;">Ação</th>
                                    <th class="text-center" style="width: 150px;">Cód. Produto</th>
                                    <th class="text-center">Descrição</th>
                                </tr>
                            </thead>
                            <tbody id="tabelaModalProduto">
                                <tr>
                                    <td colspan="3" class="text-center text-muted py-3">Digite para buscar</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-sm btn-secondary" data-bs-dismiss="modal">Fechar</button>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
