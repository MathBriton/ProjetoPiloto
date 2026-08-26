<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Smurfit Kappa - Resumo de Vendas</title>
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
    </style>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

    <script src="/corporativo/components/toast.js?v=<%=Now()%>"></script>
    <script src="/corporativo/funcoes/funcoesJAVASCRIPT.js?v=<%=Now()%>"></script>

    <script>
        let dadosRelatorio = [];
        let mesesSelecionados = [];
        let carregando = false;

        // Ao entrar na página
        $(document).ready(function() {
            inicializarSelect2();
            anexarEventos();
            renderizarGrade();

            const urlParams = new URLSearchParams(window.location.search);
            if (urlParams.size === 0) return;

            setTimeout(() => {
                if (urlParams.has('filtroFilial')) 
                    $('#filtroFilial').val(urlParams.get('filtroFilial')).trigger('change');

                if (urlParams.has('filtroCodGrupo'))  
                    $('#codGrupo').text(urlParams.get('filtroCodGrupo'));

                if (urlParams.has('filtroDataInicial')) 
                    $('#filtroDataInicial').val(urlParams.get('filtroDataInicial'));

                if (urlParams.has('filtroDataFinal')) 
                    $('#filtroDataFinal').val(urlParams.get('filtroDataFinal'));

                if (urlParams.has('filtroRepresentante')) 
                    $('#filtroRepresentante').val(urlParams.get('filtroRepresentante'));

                if (urlParams.has('filtroTipoProduto')) 
                    $('#filtroTipoProduto').val(urlParams.get('filtroTipoProduto')).trigger('change');

                if (urlParams.has('filtroTipoPapelao')) 
                    $('#filtroTipoPapelao').val(urlParams.get('filtroTipoPapelao')).trigger('change');

                if (urlParams.has('filtroValores')) 
                    $('#filtroValores').val(urlParams.get('filtroValores')).trigger('change');

                if (urlParams.has('filtroImpostos')) 
                    $('#filtroImpostos').val(urlParams.get('filtroImpostos')).trigger('change');
                pesquisar();
            }, 1000)
        });

        /**
         * Calculo prc 
         * Filtro Valores (Todos ou Peso) => Valor/Peso
         * Filtro Tipo Produto = Papel => Valor/Peso
         * Filtro Valores (M2) => Valor/Area
         */
        function calcularValorPrc(peso, area, valor) {
            const valores = $('#filtroValores').val();
            const ehTipoPapel = $('#filtroTipoProduto').val() == 1;

            if (valores == 1 || valores == 0 || ehTipoPapel) {
                if (valor == 0 || peso == 0) return 0
                return (valor / peso);
            } else {
                if (valor == 0 || area == 0) return 0
                return (valor / area);
            } 
        }

        function obterMesesPeriodo() {
            const meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                            'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

            let inicio = new Date($('#filtroDataInicial').val() + 'T00:00:00');
            const fim  = new Date($('#filtroDataFinal').val()   + 'T00:00:00');
            const resultado = [];

            while (inicio <= fim) {
                resultado.push({
                    label: `${meses[inicio.getMonth()]} ${inicio.getFullYear()}`,
                    key: `${inicio.getFullYear()}${inicio.getMonth() + 1}`
                });
                inicio.setMonth(inicio.getMonth() + 1);
            }

            return resultado;
        }

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                placeholder: function() {
                    return $(this).data('placeholder');
                }
            });
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

        function agruparPorLoja(dados) {
            let dadosAgrupados = []
            
            dados.forEach(item => {
                const itemAgrupado = dadosAgrupados.find(d => (d.codigoCliente == item.codigoCliente) && (d.nomeVendedor == item.nomeVendedor));

                if (itemAgrupado) {
                    itemAgrupado.meses.push({ano: item.ano, peso: parseFloat(item.peso), area: parseFloat(item.area), valor: parseFloat(item.valor)});
                } else {
                    dadosAgrupados.push({
                        nomeVendedor: item.nomeVendedor,
                        codigoCliente: item.codigoCliente,
                        loja: item.loja,
                        meses: [{ano: item.ano, peso: parseFloat(item.peso), area: parseFloat(item.area), valor: parseFloat(item.valor)}]
                    });
                }
            })

            return dadosAgrupados
        }

        function pesquisar() { 
            dadosRelatorio = []; 
            exibirLoading(true);

            const params = {
                acao: 'listarVendasLojas',
                filtroCodGrupo: $('#codGrupo').text(),
                filtroFilial: $('#filtroFilial').val(),
                filtroDataInicial: $('#filtroDataInicial').val(),
                filtroDataFinal: $('#filtroDataFinal').val(),
                filtroRepresentante: $('#filtroRepresentante').val(),
                filtroTipoProduto: $('#filtroTipoProduto').val(),
                filtroTipoPapelao: $('#filtroTipoPapelao').val(),
                filtroValores: $('#filtroValores').val(),
                filtroImpostos: $('#filtroImpostos').val()
            };

            $.ajax({
                url: 'RelatorioVendasRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error('Erro ao pesquisar: ' + response.mensagem);
                        return;
                    }
                    
                    // Backend só retorna dados do cliente nos meses com vendas
                    // Se não tiver venda em um dos meses precisamos tratar adicionando 0 nos valores
                    mesesSelecionados = obterMesesPeriodo();
                    
                    const mesVazio = { peso: 0, area: 0, valor: 0, prc: 0 };
                    const montarMes = (mesItem) => ({
                        peso: mesItem.peso,
                        area: mesItem.area,
                        valor: mesItem.valor,
                        prc: calcularValorPrc(mesItem.peso, mesItem.area, mesItem.valor),
                    });

                    const dadosAgrupados = agruparPorLoja(response.dados);
                    dadosAgrupados.forEach(item => {
                        const meses = mesesSelecionados.reduce((acc, { key }) => {
                            const mesEncontrado = item.meses.find(m => m.ano == key);
                            acc[key] = mesEncontrado ? montarMes(mesEncontrado) : mesVazio;
                            return acc;
                        }, {})

                        const pesoTotal = item.meses.reduce((acc, mes) => acc + mes.peso, 0)
                        const areaTotal = item.meses.reduce((acc, mes) => acc + mes.area, 0)
                        const valorTotal = item.meses.reduce((acc, mes) => acc + mes.valor, 0)
                        const prcTotal = calcularValorPrc(pesoTotal, areaTotal, valorTotal)

                        dadosRelatorio.push({
                            ...item,
                            meses,
                            pesoTotal,
                            areaTotal,
                            valorTotal,
                            prcTotal
                        })
                    })

                    renderizarGrade();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao pesquisar:', error);
                }
            });
        }

         const COLUNAS_RELATORIO = {
            peso: { key: 'peso', label: 'Peso', casasDecimais: 3 },
            area: { key: 'area', label: 'Área', casasDecimais: 3 },
            valor: { key: 'valor', label: 'Valor', casasDecimais: 2 },
            prc: { key: 'prc', label: 'PRC', casasDecimais: 2 },
        }

        function obterColunasRelatorio() {
            const ehTipoPapel = $('#filtroTipoProduto').val() == 1;
            const soPeso = $('#filtroValores').val() == 1;
            const soArea = $('#filtroValores').val() == 2;

            if (ehTipoPapel || soPeso) {
                return [COLUNAS_RELATORIO.peso, COLUNAS_RELATORIO.valor, COLUNAS_RELATORIO.prc]
            }

            if (soArea) {
                return [COLUNAS_RELATORIO.area, COLUNAS_RELATORIO.valor, COLUNAS_RELATORIO.prc]
            }

            return [COLUNAS_RELATORIO.peso, COLUNAS_RELATORIO.area, COLUNAS_RELATORIO.valor, COLUNAS_RELATORIO.prc]
        }

        function formatarValorColuna(valor, coluna) {
            return formatarMoeda(parseFloat(valor) || 0, coluna.casasDecimais)
        }

        function obterLabelsColunas(colunas) {
            return colunas.map(coluna => coluna.label)
        }

        function somarCampo(valores, campo) {
            return valores.reduce((soma, item) => soma + (parseFloat(item[campo]) || 0), 0)
        }

        function montarTotalValores(valores) {
            const peso = somarCampo(valores, 'peso');
            const area = somarCampo(valores, 'area');
            const valor = somarCampo(valores, 'valor');

            return { peso: peso, area: area, valor: valor, prc: calcularValorPrc(peso, area, valor) }
        }

        function montarTotalGeral() {
            const peso = dadosRelatorio.reduce((s, item) => s + (parseFloat(item.pesoTotal) || 0), 0)
            const area = dadosRelatorio.reduce((s, item) => s + (parseFloat(item.areaTotal) || 0), 0)
            const valor = dadosRelatorio.reduce((s, item) => s + (parseFloat(item.valorTotal) || 0), 0)

            return { peso: peso, area: area, valor: valor, prc: calcularValorPrc(peso, area, valor) }
        }

        function obterValorTotalItem(item, coluna) {
            return item[`${coluna.key}Total`]
        }

        function gerarCelulasValores(valores, colunas, tag = 'td') {
            return colunas.map(coluna =>
                `<${tag} style="text-align: right">${formatarValorColuna(valores[coluna.key], coluna)}</${tag}>`
            ).join('')
        }

        function montarLinhaExportacao(item, colunas) {
            return [
                item.nomeVendedor,
                item.codigoCliente,
                item.loja,
                ...mesesSelecionados.flatMap(({ key }) =>
                    colunas.map(coluna => formatarValorColuna(item.meses[key][coluna.key], coluna))
                ),
                ...colunas.map(coluna => formatarValorColuna(obterValorTotalItem(item, coluna), coluna)),
            ]
        }

        function renderizarGrade() {
            const $thead = $('#gradeRelatorio thead');
            const $tbody = $('#gradeRelatorio tbody');
            const $tfoot = $('#gradeRelatorio tfoot');

            $thead.empty(); 
            $tbody.empty(); 
            $tfoot.empty();

            if (dadosRelatorio.length === 0) {
                $tbody.html(`
                    <tr>
                        <td colspan="28" class="text-center text-muted py-4">
                            Utilize os filtros acima para pesquisar
                        </td>
                    </tr>
                `);
                return;
            }

            const colunas = obterColunasRelatorio();
            const colSpan = colunas.length;
            const subHeaders = obterLabelsColunas(colunas).map(label => `<th>${label}</th>`).join('');
            
            // Monta Cabeçalho
            $thead.append(`
                <tr>
                    <th class="text-center align-middle" rowspan="2">Representante</th>
                    <th class="text-center align-middle" rowspan="2">Código</th>
                    <th class="text-center align-middle" rowspan="2">Loja</th>
                    ${mesesSelecionados.map(mes => `<th class="text-center" colspan="${colSpan}">${mes.label}</th>`).join('')}
                    <th class="text-center" colspan="${colSpan}">Total</th>
                </tr>
                <tr>
                    ${mesesSelecionados.map(() => subHeaders).join('')}
                    ${subHeaders}
                </tr>
            `);

            // Monta Corpo
            dadosRelatorio.forEach(item => {
                const mesesValores = mesesSelecionados.map(({ key }) => item.meses[key]);
                const totalLinha = montarTotalValores(mesesValores);

                $tbody.append(`
                    <tr>
                        <td>${item.nomeVendedor}</td>
                        <td>${item.codigoCliente}</td>
                        <td>${item.loja}</td>
                        ${mesesValores.map(mes => gerarCelulasValores(mes, colunas)).join('')}
                        ${gerarCelulasValores(totalLinha, colunas, 'th')}
                    </tr>
                `);
            });

            // Monta Rodapé
            const mesesTotais = mesesSelecionados.map(({ key }) => {
                const valoresMes = dadosRelatorio.map(item => item.meses[key]);
                return gerarCelulasValores(montarTotalValores(valoresMes), colunas);
            }).join('');

            $tfoot.html(`
                <tr>
                    <td>Totais</td><td></td>
                    <td style="text-align: right">${dadosRelatorio.length}</td>
                    ${mesesTotais}
                    ${gerarCelulasValores(montarTotalGeral(), colunas)}
                </tr>
            `);
        }

        function limparFiltros() {
            $('#filtroDataInicial').val('');
            $('#filtroDataFinal').val('');
            $('#filtroRepresentante').val('');
            $('#filtroTipoProduto').val('0').trigger('change');
            $('#filtroTipoPapelao').val('0').trigger('change');
            $('#filtroValores').val('0').trigger('change');
            $('#filtroImpostos').val('0').trigger('change');
            
            dadosRelatorio = [];
            mesesSelecionados = [];
            renderizarGrade();
        }

        function exportarPDF() {
            if (dadosRelatorio.length === 0) { 
                Toast.error('Não há dados para exportar'); 
                return; 
            }

            const colunas = obterColunasRelatorio();
            const subLabels = obterLabelsColunas(colunas);

            const linhaHeaders = [
                { content: 'Representante', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
                { content: 'Código', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
                { content: 'Loja', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
                ...mesesSelecionados.map(mes => ({ content: mes.label, colSpan: colunas.length, styles: { halign: 'center' } })),
                { content: 'Total', colSpan: colunas.length, styles: { halign: 'center' } },
            ];

            const linhaSubheaders = [
                ...mesesSelecionados.flatMap(() => subLabels),
                ...subLabels,
            ]

            const dados = dadosRelatorio.map(item => montarLinhaExportacao(item, colunas));
            const totalGeral = montarTotalGeral();
            const totalMeses = mesesSelecionados.flatMap(({ key }) => {
                const valoresMes = dadosRelatorio.map(item => item.meses[key]);
                const totalMes = montarTotalValores(valoresMes);
                return colunas.map(coluna => formatarValorColuna(totalMes[coluna.key], coluna));
            });
            const linhaTotal = [
                'Totais',
                '',
                dadosRelatorio.length,
                ...totalMeses,
                ...colunas.map(coluna => formatarValorColuna(totalGeral[coluna.key], coluna)),
            ];

            const dataAtual = new Intl.DateTimeFormat('pt-BR', {
                day: '2-digit', month: '2-digit', year: 'numeric',
                hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
            }).format(new Date());

            const { jsPDF } = window.jspdf;
            const formatoPapel = obterFormatoPapelPDF(dados[0].length);
            const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: formatoPapel });

            let yPosition = 15;
            let filtrosAplicados = [];
            
            const filtroFilial = $('#filtroFilial').val() && $('#filtroFilial option:selected').text();
            const filtroDataInicial = $('#filtroDataInicial').val();
            const filtroDataFinal = $('#filtroDataFinal').val();
            const filtroRepresentante = $('#filtroRepresentante').val();
            const filtroTipoProduto = $('#filtroTipoProduto').val() && $('#filtroTipoProduto option:selected').text();
            const filtroTipoPapelao = $('#filtroTipoPapelao').val() && $('#filtroTipoPapelao option:selected').text();
            const filtroValores = $('#filtroValores').val() && $('#filtroValores option:selected').text();
            const filtroImpostos = $('#filtroImpostos').val() && $('#filtroImpostos option:selected').text();

            if (filtroFilial) 
                filtrosAplicados.push('Filial: ' + filtroFilial);
            
            if (filtroDataInicial) 
                filtrosAplicados.push('Data Início: ' + new Date(`${filtroDataInicial} 03:00`).toLocaleDateString('pt-BR'));

            if (filtroDataFinal) 
                filtrosAplicados.push('Data Fim: ' + new Date(`${filtroDataFinal} 03:00`).toLocaleDateString('pt-BR') );

            if (filtroRepresentante) 
                filtrosAplicados.push('Representante: ' + filtroRepresentante);

            if (filtroTipoProduto) 
                filtrosAplicados.push('Tipo de Produto: ' + filtroTipoProduto);

            if (filtroTipoPapelao) 
                filtrosAplicados.push('Tipo de Papelão: ' + filtroTipoPapelao);

            if (filtroValores) 
                filtrosAplicados.push('Valores: ' + filtroValores);

            if (filtroImpostos) 
                filtrosAplicados.push('Impostos: ' + filtroImpostos);

            if (filtrosAplicados.length > 0) {
                doc.text('Filtros Aplicados:', 14, yPosition);
                yPosition += 7;
                doc.setFont(undefined, 'normal');
                doc.setFontSize(10);

                for (var i = 0; i < filtrosAplicados.length; i++) {
                    doc.text('  • ' + filtrosAplicados[i], 14, yPosition);
                    yPosition += 4;
                }

                yPosition += 2;
            }

            yPosition += 7;
            doc.setFontSize(16);
            doc.text(`Detalhes de Venda para Grupo #${$('#codGrupo').text()}`, 14, yPosition);

            yPosition += 7;
            doc.setFontSize(10);
            doc.text(`Data: ${dataAtual} | ${dados.length} registro(s) encontrado(s)`, 14, yPosition);
            yPosition += 7;

            doc.autoTable({
                head: [linhaHeaders, linhaSubheaders],
                body: dados,
                foot: [linhaTotal],
                startY: yPosition,
                theme: 'grid',
                styles: {
                    fontSize: 8,
                    cellPadding: 2,
                    lineWidth: 0.1, 
                    lineColor: [0, 0, 0, 0.4],
                },
                headStyles: {
                    fillColor: [211, 211, 212],
                    textColor: [0, 0, 0],
                    fontStyle: 'bold',
                    halign: 'center',
                    lineWidth: 0.1, 
                    lineColor: [0, 0, 0, 0.4],
                },
                footStyles: {
                    fillColor: [211, 211, 212],
                    textColor: [0, 0, 0],
                    fontStyle: 'bold',
                    lineWidth: 0.1, 
                    lineColor: [0, 0, 0, 0.4],
                },
                columnStyles: {
                    0: { cellWidth: 40 },  
                    1: { cellWidth: 25 },  
                    2: { cellWidth: 40 }, 
                },
                margin: { top: 28 },
            });

            window.open(doc.output('bloburl'), '_blank');
        }

        function exportarXLSX() {
            if (dadosRelatorio.length === 0) { 
                Toast.error('Não há dados para exportar'); 
                return; 
            }

            const colunas = obterColunasRelatorio();
            const quantidadeColunas = colunas.length;
            const subLabels = obterLabelsColunas(colunas);

            const linhaHeaders    = ['Representante', 'Código', 'Loja'];
            const linhaSubHeaders = ['', '', ''];
            const mesclagens = [
                { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }, // Representante
                { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } }, // Código
                { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } }, // Loja
            ];

            mesesSelecionados.forEach((mes, i) => {
                const colunaInicio = 3 + i * quantidadeColunas;
                
                linhaHeaders.push(mes.label, ...Array(quantidadeColunas - 1).fill(''));
                linhaSubHeaders.push(...subLabels);
                mesclagens.push({ s: { r: 0, c: colunaInicio }, e: { r: 0, c: colunaInicio + quantidadeColunas - 1 } });
            });

            linhaHeaders.push('Total', ...Array(quantidadeColunas - 1).fill(''));
            linhaSubHeaders.push(...subLabels);
            mesclagens.push({ s: { r: 0, c: 3 + mesesSelecionados.length * quantidadeColunas }, e: { r: 0, c: 3 + ((mesesSelecionados.length + 1) * quantidadeColunas) - 1 } });

            const dados = dadosRelatorio.map(item => montarLinhaExportacao(item, colunas));

            const planilha = XLSX.utils.aoa_to_sheet([linhaHeaders, linhaSubHeaders, ...dados]);
            planilha['!merges'] = mesclagens;
            planilha['!cols']   = [{ wch: 24 }, { wch: 12 }, ...Array(linhaSubHeaders.length - 2).fill({ wch: 16 })];

            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, planilha, 'Detalhes de Vendas');
            XLSX.writeFile(workbook, 'detalhes_vendas_grupo.xlsx');
        }

        function anexarEventos() {
            $('#filtroTipoProduto').on('change', filtroTipoProdutoAlterado)
            $('#formularioFiltro').on('submit', validarFiltros);
            $('#btnLimpar').on('click', limparFiltros);
            $('#btnExportarExcel').on('click', exportarXLSX);
            $('#btnExportarPDF').on('click', exportarPDF);
        }

        function filtroTipoProdutoAlterado() {
            const tipoProduto = $('#filtroTipoProduto').val();
            
            if (tipoProduto == 1) {
                $('#filtroTipoPapelao').prop('disabled', true).trigger('change');
                $('#filtroTipoPapelao').val('0').trigger('change');
            } else {
                $('#filtroTipoPapelao').prop('disabled', false).trigger('change');
            }
        }
    </script>
</head>
<body>
    <div class="container my-4">
        <h3 class="mb-4">Detalhes de Vendas Para Grupo #<span id="codGrupo"></span></h3>

        <form name="formularioFiltro" id="formularioFiltro" method="post">                              
            <div class="card p-3 mb-4">
                <div class="row g-3">
                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroFilial" class="form-label">Filial</label>
                        <select disabled required class="form-select form-select-sm select2" id="filtroFilial" name="filtroFilial" data-placeholder="Todas as filiais">
                            <option value="">Todas as filiais</option>
                            <option value="00">Pirapetinga</option>
                            <option value="04">Uberaba</option>
                            <option value="05">Bento</option>
                            <option value="06">Maranguape</option>
                            <option value="07">Saquarema</option>
                            <option value="10">Eshop</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroDataInicial" class="form-label">Data Inicial</label>
                        <input required type="date" class="form-control form-control-sm" id="filtroDataInicial" name="filtroDataInicial">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroDataFinal" class="form-label">Data Final</label>
                        <input required type="date" class="form-control form-control-sm" id="filtroDataFinal" name="filtroDataFinal">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroRepresentante" class="form-label">Representante</label>
                        <input type="text" class="form-control form-control-sm" id="filtroRepresentante" name="filtroRepresentante" placeholder="Representante">
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroTipoProduto" class="form-label">Tipo de Produto</label>
                        <select class="form-select form-select-sm select2" id="filtroTipoProduto" name="filtroTipoProduto" data-placeholder="Todos os tipos">
                            <option value="0">Ambos</option>
                            <option value="1">Papel</option>
                            <option value="2">Papelão</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroTipoPapelao" class="form-label">Tipo de Papelão</label>
                        <select class="form-select form-select-sm select2" id="filtroTipoPapelao" name="filtroTipoPapelao" data-placeholder="Todos os tipos">
                            <option value="0">Ambos</option>
                            <option value="1">Caixa</option>
                            <option value="2">Chapa</option>
                            <option value="3">Corte Vinco</option>
                            <option value="4">Maleta</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroValores" class="form-label">Valores</label>
                        <select class="form-select form-select-sm select2" id="filtroValores" name="filtroValores" data-placeholder="Todos os tipos">
                            <option value="0">Todos</option>
                            <option value="1">Só Peso</option>
                            <option value="2">Só M2</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroImpostos" class="form-label">Impostos</label>
                        <select class="form-select form-select-sm select2" id="filtroImpostos" name="filtroImpostos" data-placeholder="Todos os tipos">
                            <option value="0">COM IPI</option>
                            <option value="1">SEM IPI</option>
                            <option value="2">Sem Imposto (ICM, PIS, COF)</option>
                        </select>
                    </div>

                    <div class="col-12 d-flex justify-content-end align-items-end gap-2">
                        <button type="button" class="btn btn-sm btn-outline-secondary" id="btnLimpar">
                            <i class="bi bi-x-lg me-1"></i> Limpar
                        </button>
                        <button type="submit" type="button" class="btn btn-sm btn-primary" id="btnPesquisar">
                            <i class="bi bi-search me-1"></i> Pesquisar
                        </button>
                    </div>
                </div>
            </div>
        </form>

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

                <table class="table table-bordered table-hover mb-0" id="gradeRelatorio">
                    <thead>  
                    </thead>

                    <tbody>

                    </tbody>

                    <tfoot>
                    </tfoot>
                </table>
            </div>
        </div>
    </div>
</body>
</html>
