<!DOCTYPE html>
<html lang="pt-br">
<head>
    <title>Smurfit Kappa - Resumo de Vendas Carteira</title>
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
        // Auxiliar para funções
        const meses = [
            { label: 'Janeiro',   key: 'JAN' },
            { label: 'Fevereiro', key: 'FEV' },
            { label: 'Março',     key: 'MAR' },
            { label: 'Abril',     key: 'ABR' },
            { label: 'Maio',      key: 'MAI' },
            { label: 'Junho',     key: 'JUN' },
            { label: 'Julho',     key: 'JUL' },
            { label: 'Agosto',    key: 'AGO' },
            { label: 'Setembro',  key: 'SET' },
            { label: 'Outubro',   key: 'OUT' },
            { label: 'Novembro',  key: 'NOV' },
            { label: 'Dezembro',  key: 'DEZ' },
        ];

        let dadosRelatorio = [];

        // Ao entrar na página
        $(document).ready(function() {
            inicializarSelect2();
            carregarAnos();
            anexarEventos();
            renderizarGrade();

            $('#filtroMes').val((new Date().getMonth() + 1).toString().padStart(2, '0')).trigger('change');
            $('#filtroAno').val(new Date().getFullYear()).trigger('change');
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

        function inicializarSelect2() {
            $('.select2').select2({
                theme: 'bootstrap-5',
                width: '100%',
                placeholder: function() {
                    return $(this).data('placeholder');
                }
            });
        }

        // Obtem ano atual e os 8 anteriores
        function carregarAnos() {
            const $select = $('#filtroAno');
            $select.empty().append('<option value="">Todos os anos</option>');
            const anoAtual = new Date().getFullYear();
            for (let ano = anoAtual; ano >= anoAtual - 8; ano--) {
                $select.append($('<option>', { value: ano, text: ano }));
            }
        }

        function pesquisarCliente() {
            const codPlanta = $('#filtroFilial').val();
            const baseUrl = `${location.protocol}//${location.host}`;
            const url = `${baseUrl}/corporativo/buscaCLIENTE.asp?id1=filtroCodCliente&id2=filtroCodLoja&id3=nomeCliente&id6=${codPlanta}`;
            window.open(url, '', "width=650,height=400");
        }

        function exibirLoading(exibir) {
            if (exibir) {
                $('#loadingOverlay').show();
            } else {
                $('#loadingOverlay').hide();
            }
        }

        function obterUrlDetalhesCliente(grupoCliente) {
            let urlDetalhes = `RelatorioVendasCarteiraClienteDetalhes.asp?filtroCodGrupo=${grupoCliente}` 
            
            if($('#filtroFilial').val())
                urlDetalhes += `&filtroFilial=${$('#filtroFilial').val()}`

            if($('#filtroAno').val())
                urlDetalhes += `&filtroAno=${$('#filtroAno').val()}`

            if($('#filtroMes').val())
                urlDetalhes += `&filtroMes=${$('#filtroMes').val()}`

            if($('#filtroClienteGrupo').val())
                urlDetalhes += `&filtroClienteGrupo=${$('#filtroClienteGrupo').val()}`

            if($('#filtroRepresentante').val())
                urlDetalhes += `&filtroRepresentante=${$('#filtroRepresentante').val()}`

            if($('#filtroTipoProduto').val())
                urlDetalhes += `&filtroTipoProduto=${$('#filtroTipoProduto').val()}`

            if($('#filtroTipoPapelao').val())
                urlDetalhes += `&filtroTipoPapelao=${$('#filtroTipoPapelao').val()}`

            if($('#filtroValores').val())
                urlDetalhes += `&filtroValores=${$('#filtroValores').val()}`

            if($('#filtroImpostos').val())
                urlDetalhes += `&filtroImpostos=${$('#filtroImpostos').val()}`

            return urlDetalhes
        }

        function obterMesesPeriodo() {
            const anoSelecionado = parseInt($('#filtroAno').val());
            const mesSelecionado = parseInt($('#filtroMes').val());
            const mesAtual = new Date().getMonth() + 1;
            const anoAtual = new Date().getFullYear();

            const resultado = [];
            let ano = anoSelecionado;
            let mes = mesSelecionado;

            while (ano < anoAtual || (ano === anoAtual && mes <= mesAtual)) {
                const { label, key } = meses[mes - 1];
                resultado.push({ label: `${label} ${ano}`, key: `${key}${ano}`, prefixo: key, mes: mes, ano: ano });

                mes++;
                if (mes > 12) {
                    mes = 1;
                    ano++;
                }
            }

            return resultado;
        }

        function obterProximoMes() {
            const hoje = new Date();
            let mes = hoje.getMonth() + 1;
            let ano = hoje.getFullYear();

            mes++;
            if (mes > 12) {
                mes = 1;
                ano++;
            }

            const { label, key } = meses[mes - 1];
            return { label: `${label} ${ano}`, key: `${key}`, prefixo: key, mes: mes, ano: ano };
        }
        
        // Função auxiliar para normalizar chaves de itens
        // Backend está retornando chaves com espaço no meio, substituimos por _
        function renameKeys(obj) {
            return Object.fromEntries(
                Object.entries(obj).map(function([key, value]) {
                    return [key.replaceAll(' ', '_'), value];
                })
            );
        }

        function formatarDadosRelatorio(dados) {
            const anoSelecionado = $('#filtroAno').val();
            const mesAtual = mesesSelecionados[mesesSelecionados.length - 1];
            const proximoMes = obterProximoMes(mesAtual);
            const dadosTratados = dados.map(renameKeys);
            const resultado = [];
            
            dadosTratados.forEach(item => {
                const meses = mesesSelecionados.reduce((acc, { key }) => {
                    acc[key] = {
                        mes: key,
                        ano: anoSelecionado,
                        peso: item[`PESO_${key}`],
                        area: item[`AREA_${key}`],
                        valor: item[`VALOR_${key}`],
                        prc: item[`PRC_${key}`],
                    }
                    return acc;
                }, {})

                const pesoTotal = Object.values(meses).reduce((acc, mes) => acc + mes.peso, 0)
                const areaTotal = Object.values(meses).reduce((acc, mes) => acc + mes.area, 0)
                const valorTotal = Object.values(meses).reduce((acc, mes) => acc + mes.valor, 0)
                const prcTotal = calcularValorPrc(pesoTotal, areaTotal, valorTotal)
                
                resultado.push({
                    // Dados gerais
                    nomeVendedor: item.NOME_VENDEDOR,
                    grupoCliente: item.GRUPO_DO_CLIENTE,
                    nomeCliente: item.NOME_CLIENTE,

                    // Dados de cada mês do período
                    meses,

                    // Totais da linha
                    pesoTotal,
                    areaTotal,
                    valorTotal,
                    prcTotal,

                    // Carteira Atual
                    carteiraMesAtualPeso: item.TOTAL_CARTEIRA_PESO,
                    carteiraMesAtualArea: item.TOTAL_CARTEIRA_AREA,
                    carteiraMesAtualValor: item.TOTAL_CARTEIRA_VALOR,
                    carteiraMesAtualPrc: item.TOTAL_CARTEIRA_PRC,
                    
                    // Totais do mês atual
                    mesAtualPeso: item[`TOTAL_${mesAtual.prefixo}_${mesAtual.ano}_PESO`],
                    mesAtualArea: item[`TOTAL_${mesAtual.prefixo}_${mesAtual.ano}_AREA`],
                    mesAtualValor: item[`TOTAL_${mesAtual.prefixo}_${mesAtual.ano}_VALOR`],
                    mesAtualPrc: item[`TOTAL_${mesAtual.prefixo}_${mesAtual.ano}_PRC`],

                    // Carteira proximo mês
                    carteiraProximoMesPeso: item[`TOTAL_${proximoMes.prefixo}_${proximoMes.ano}_PESO`],
                    carteiraProximoMesArea: item[`TOTAL_${proximoMes.prefixo}_${proximoMes.ano}_AREA`],
                    carteiraProximoMesValor: item[`TOTAL_${proximoMes.prefixo}_${proximoMes.ano}_VALOR`],
                    carteiraProximoMesPrc: item[`TOTAL_${proximoMes.prefixo}_${proximoMes.ano}_PRC`],
                })
            })

            return resultado
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
            dadosRelatorio = []
            exibirLoading(true);

            const params = {
                acao: 'listar',
                filtroFilial: $('#filtroFilial').val(),
                filtroCodCliente: $('#filtroCodCliente').val(),
                filtroCodLoja: $('#filtroCodLoja').val(),
                filtroMes: $('#filtroMes').val(),
                filtroAno: $('#filtroAno').val(),
                filtroClienteGrupo: $('#filtroClienteGrupo').val(),
                filtroRepresentante: $('#filtroRepresentante').val(),
                filtroTipoProduto: $('#filtroTipoProduto').val(),
                filtroTipoPapelao: $('#filtroTipoPapelao').val(),
                filtroValores: $('#filtroValores').val(),
                filtroImpostos: $('#filtroImpostos').val(),
                filtroRetiraHref: '0'
            };

            $.ajax({
                url: 'RelatorioVendasCarteiraRepositorio.asp?v=<%=Now()%>',
                type: 'GET',
                data: params,
                dataType: 'json',
                success: function(response) {
                    exibirLoading(false);

                    if (!response.sucesso) {
                        Toast.error('Erro ao pesquisar: ' + response.mensagem);
                        return;
                    }

                    mesesSelecionados = obterMesesPeriodo();
                    dadosRelatorio = formatarDadosRelatorio(response.dados);
                    renderizarGrade();
                },
                error: function(xhr, status, error) {
                    exibirLoading(false);
                    console.error('Erro ao pesquisar:', error);
                }
            });
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
            
            const mesAtual = mesesSelecionados[mesesSelecionados.length - 1];
            const proximoMes = obterProximoMes(mesAtual);
            const ehTipoPapel = $('#filtroTipoProduto').val() == 1;
            const soPeso = $('#filtroValores').val() == 1;
            const soArea = $('#filtroValores').val() == 2;

            let colunas = []
            if (ehTipoPapel || soPeso) {
                colunas = [{ key: 'peso', label: 'Peso' }, { key: 'valor', label: 'Valor' }, { key: 'prc', label: 'PRC' }]
            } else if (soArea) {
                colunas = [{ key: 'area', label: 'Área' }, { key: 'valor', label: 'Valor' }, { key: 'prc', label: 'PRC' }]
            } else {
                colunas = [{ key: 'peso', label: 'Peso' }, { key: 'area', label: 'Área' }, { key: 'valor', label: 'Valor' }, { key: 'prc', label: 'PRC' }]
            }

            const colSpan = colunas.length;
            const gerarCelulas = (valores, tag = 'td') =>
                colunas.map(({ key }) =>
                    `<${tag} style="text-align: right">${formatarMoeda(valores[key], 3)}</${tag}>`
                ).join('');
            
            // Monta Cabeçalho
            const subHeaders = colunas.map(({ label }) => `<th>${label}</th>`).join('');
            $thead.append(`
                <tr>
                    <th class="text-center align-middle" rowspan="2">Representante</th>
                    <th class="text-center align-middle" rowspan="2">Grupo</th>
                    <th class="text-center align-middle" rowspan="2">Cliente</th>
                    ${mesesSelecionados.map(mes => `<th class="text-center" colspan="${colSpan}">${mes.label}</th>`).join('')}
                    <th class="text-center" colspan="${colSpan}">Total</th>
                    <th class="text-center" colspan="${colSpan}">Carteira</th>
                    <th class="text-center" colspan="${colSpan}">Total Mês ${mesAtual.label}</th>
                    <th class="text-center" colspan="${colSpan}">Carteira ${proximoMes.label}</th>
                </tr>
                <tr>
                    ${mesesSelecionados.map(() => subHeaders).join('')}
                    ${subHeaders}
                    ${subHeaders}
                    ${subHeaders}
                    ${subHeaders}
                </tr>
            `);

            // Monta Corpo
            dadosRelatorio.forEach(item => {
                const mesesValores = mesesSelecionados.map(({ key }) => item.meses[key]);

                const peso = mesesValores.reduce((s, m) => s + m.peso, 0);
                const area = mesesValores.reduce((s, m) => s + m.area, 0);
                const valor = mesesValores.reduce((s, m) => s + m.valor, 0);
                
                const totalLinha = { 
                    peso: peso, 
                    area: area, 
                    valor: valor, prc: calcularValorPrc(peso, area, valor) 
                };

                const carteiraMesAtual = { 
                    peso: item.carteiraMesAtualPeso, 
                    area: item.carteiraMesAtualArea, 
                    valor: item.carteiraMesAtualValor, 
                    prc: item.carteiraMesAtualPrc 
                };

                const totalMesAtual = {
                    peso: item.mesAtualPeso, 
                    area: item.mesAtualArea, 
                    valor: item.mesAtualValor, 
                    prc: item.mesAtualPrc
                }

                const carteiraProximoMes = {
                    peso: item.carteiraProximoMesPeso, 
                    area: item.carteiraProximoMesArea, 
                    valor: item.carteiraProximoMesValor, 
                    prc: item.carteiraProximoMesPrc
                }

                const urlDetalhes = obterUrlDetalhesCliente(item.grupoCliente);
                $tbody.append(`
                    <tr>
                        <td>${item.nomeVendedor}</td>
                        <td>${item.grupoCliente}</td>
                        <td>
                            <a href="${urlDetalhes}" target="_blank">
                                ${item.nomeCliente}
                                <i class="bi bi-box-arrow-up-right"></i>
                            </a>
                        </td>
                        ${mesesValores.map(mes => gerarCelulas(mes)).join('')}
                        ${gerarCelulas(totalLinha)}
                        ${gerarCelulas(carteiraMesAtual)}
                        ${gerarCelulas(totalMesAtual)}
                        ${gerarCelulas(carteiraProximoMes)}
                    </tr>
                `);
            });

            // Monta Rodapé
            const mesesTotais = mesesSelecionados.map(({ key }) => {
                const peso = dadosRelatorio.reduce((s, item) => s + parseFloat(item.meses[key].peso), 0);
                const area = dadosRelatorio.reduce((s, item) => s + parseFloat(item.meses[key].area), 0);
                const valor = dadosRelatorio.reduce((s, item) => s + parseFloat(item.meses[key].valor), 0);
                return gerarCelulas({ peso: peso, area: area, valor: valor, prc: calcularValorPrc(peso, area, valor) });
            }).join('');

            const pesoTotal = dadosRelatorio.reduce((s, item) => s + parseFloat(item.pesoTotal), 0)
            const areaTotal = dadosRelatorio.reduce((s, item) => s + parseFloat(item.areaTotal), 0)
            const valorTotal = dadosRelatorio.reduce((s, item) => s + parseFloat(item.valorTotal), 0)
            const totalGeral = { 
                peso: pesoTotal, 
                area: areaTotal, 
                valor: valorTotal, 
                prc: calcularValorPrc(pesoTotal, areaTotal, valorTotal) 
            };

            const pesoCarteiraMesAtualTotal = dadosRelatorio.reduce((s, item) => s + parseFloat(item.carteiraMesAtualPeso), 0)
            const areaCarteiraMesAtualTotal = dadosRelatorio.reduce((s, item) => s + parseFloat(item.carteiraMesAtualArea), 0)
            const valorCarteiraMesAtualTotal = dadosRelatorio.reduce((s, item) => s + parseFloat(item.carteiraMesAtualValor), 0)
            const totalGeralCarteiraMesAtual = { 
                peso: pesoCarteiraMesAtualTotal, 
                area: areaCarteiraMesAtualTotal, 
                valor: valorCarteiraMesAtualTotal, 
                prc: calcularValorPrc(pesoCarteiraMesAtualTotal, areaCarteiraMesAtualTotal, valorCarteiraMesAtualTotal)
            };

            const mesAtualPesoTotal = dadosRelatorio.reduce((s, item) => s + parseFloat(item.mesAtualPeso), 0)
            const mesAtualAreaTotal = dadosRelatorio.reduce((s, item) => s + parseFloat(item.mesAtualArea), 0)
            const mesAtualValorTotal = dadosRelatorio.reduce((s, item) => s + parseFloat(item.mesAtualValor), 0)
            const totalGeralMesAtual = {
                peso: mesAtualPesoTotal, 
                area: mesAtualAreaTotal, 
                valor: mesAtualValorTotal, 
                prc: calcularValorPrc(mesAtualPesoTotal, mesAtualAreaTotal, mesAtualValorTotal)
            }

            const carteiraProximoMesPesoTotal = dadosRelatorio.reduce((s, item) => s + parseFloat(item.carteiraProximoMesPeso), 0)
            const carteiraProximoMesAreaTotal = dadosRelatorio.reduce((s, item) => s + parseFloat(item.carteiraProximoMesArea), 0)
            const carteiraProximoMesValorTotal = dadosRelatorio.reduce((s, item) => s + parseFloat(item.carteiraProximoMesValor), 0)
            const totalGeralcarteiraProximoMes = {
                peso: carteiraProximoMesPesoTotal, 
                area: carteiraProximoMesAreaTotal, 
                valor: carteiraProximoMesValorTotal, 
                prc: calcularValorPrc(carteiraProximoMesPesoTotal, carteiraProximoMesAreaTotal, carteiraProximoMesValorTotal)
            }

            $tfoot.html(`
                <tr>
                    <td>Totais</td><td></td>
                    <td style="text-align: right">${dadosRelatorio.length}</td>
                    ${mesesTotais}
                    ${gerarCelulas(totalGeral)}
                    ${gerarCelulas(totalGeralCarteiraMesAtual)}
                    ${gerarCelulas(totalGeralMesAtual)}
                    ${gerarCelulas(totalGeralcarteiraProximoMes)}
                </tr>
            `);
        }

        function limparFiltros() {
            $('#filtroFilial').val('').trigger('change');
            $('#filtroCodCliente').val('');
            $('#filtroCodLoja').val('');
            $('#filtroMes').val('').trigger('change');
            $('#filtroAno').val('').trigger('change');
            $('#filtroClienteGrupo').val('T').trigger('change');
            $('#filtroRepresentante').val('');
            $('#filtroTipoProduto').val('0').trigger('change');
            $('#filtroTipoPapelao').val('0').trigger('change');
            $('#filtroValores').val('0').trigger('change');
            $('#filtroImpostos').val('0').trigger('change');
            
            dadosRelatorio = [];
            renderizarGrade();
        }

        function exportarPDF() {
            console.log(dadosRelatorio)

            if (dadosRelatorio.length === 0) { 
                Toast.error('Não há dados para exportar'); 
                return; 
            }

            const mesAtual = mesesSelecionados[mesesSelecionados.length - 1];
            const proximoMes = obterProximoMes(mesAtual);
            const ehTipoPapel = $('#filtroTipoProduto').val() == 1;
            const soPeso = $('#filtroValores').val() == 1;
            const soArea = $('#filtroValores').val() == 2;

            let colunas = []
            if (ehTipoPapel || soPeso) {
                colunas = [{ key: 'peso', label: 'Peso' }, { key: 'valor', label: 'Valor' }, { key: 'prc', label: 'PRC' }]
            } else if (soArea) {
                colunas = [{ key: 'area', label: 'Área' }, { key: 'valor', label: 'Valor' }, { key: 'prc', label: 'PRC' }]
            } else {
                colunas = [{ key: 'peso', label: 'Peso' }, { key: 'area', label: 'Área' }, { key: 'valor', label: 'Valor' }, { key: 'prc', label: 'PRC' }]
            }

            const subLabels = colunas.map(c => c.label);

            const linhaHeaders = [
                { content: 'Representante', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
                { content: 'Grupo', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
                { content: 'Cliente', rowSpan: 2, styles: { halign: 'center', valign: 'middle' } },
                ...mesesSelecionados.map(mes => ({ content: mes.label, colSpan: colunas.length, styles: { halign: 'center' } })),
                { content: 'Total', colSpan: colunas.length, styles: { halign: 'center' } },
                { content: 'Carteira', colSpan: colunas.length, styles: { halign: 'center' } },
                { content: `Total Mês ${mesAtual.label}`, colSpan: colunas.length, styles: { halign: 'center' } },
                { content: `Carteira ${proximoMes.label}`, colSpan: colunas.length, styles: { halign: 'center' } },
            ];

            const linhaSubheaders = [
                ...mesesSelecionados.flatMap(() => subLabels),
                ...colunas.flatMap(c => subLabels),
                ...colunas.flatMap(c => subLabels),
                ...colunas.flatMap(c => subLabels),
                ...colunas.flatMap(c => subLabels),
            ]

            const dados = dadosRelatorio.map(item => [
                item.nomeVendedor,
                item.grupoCliente,
                item.nomeCliente,
                ...mesesSelecionados.flatMap(({ key }) =>
                    colunas.map(col => formatarMoeda(item.meses[key][col.key], 3))
                ),
                // Total
                ...colunas.map(col => formatarMoeda(item[`${col.key}Total`], 3)),
                // Carteira
                ...colunas.map(col => formatarMoeda(item[`carteiraMesAtual${capitalize(col.key)}`], 3)),
                // Mês Atual
                ...colunas.map(col => formatarMoeda(item[`mesAtual${capitalize(col.key)}`], 3)),
                // Próximo Mês
                ...colunas.map(col => formatarMoeda(item[`carteiraProximoMes${capitalize(col.key)}`], 3)),
            ]);

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
            const filtroCodCliente = $('#filtroCodCliente').val();
            const filtroMes = $('#filtroMes').val();
            const filtroAno = $('#filtroAno').val();
            const filtroClienteGrupo = $('#filtroClienteGrupo').val() && $('#filtroClienteGrupo option:selected').text();
            const filtroRepresentante = $('#filtroRepresentante').val();
            const filtroTipoProduto = $('#filtroTipoProduto').val() && $('#filtroTipoProduto option:selected').text();
            const filtroTipoPapelao = $('#filtroTipoPapelao').val() && $('#filtroTipoPapelao option:selected').text();
            const filtroValores = $('#filtroValores').val() && $('#filtroValores option:selected').text();
            const filtroImpostos = $('#filtroImpostos').val() && $('#filtroImpostos option:selected').text();

            if (filtroFilial) 
                filtrosAplicados.push('Filial: ' + filtroFilial);

            if (filtroCodCliente) 
                filtrosAplicados.push('Cód. Cliente: ' + filtroCodCliente);
            
            if (filtroMes) 
                filtrosAplicados.push('Mês: ' + filtroMes);

            if (filtroAno) 
                filtrosAplicados.push('Ano: ' + filtroAno);

            if (filtroClienteGrupo) 
                filtrosAplicados.push('Grupo: ' + filtroClienteGrupo);

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
            doc.text('Relatório de Vendas Carteira', 14, yPosition);

            yPosition += 7;
            doc.setFontSize(10);
            doc.text(`Data: ${dataAtual} | ${dados.length} registro(s) encontrado(s)`, 14, yPosition);
            yPosition += 7;

            doc.autoTable({
                head: [linhaHeaders, linhaSubheaders],
                body: dados,
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
                columnStyles: {
                    0: { cellWidth: 40 },  
                    1: { cellWidth: 40 },  
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

            const mesAtual = mesesSelecionados[mesesSelecionados.length - 1];
            const proximoMes = obterProximoMes(mesAtual);
            const ehTipoPapel = $('#filtroTipoProduto').val() == 1;
            const soPeso = $('#filtroValores').val() == 1;
            const soArea = $('#filtroValores').val() == 2;

            let colunas = []
            if (ehTipoPapel || soPeso) {
                colunas = [{ key: 'peso', label: 'Peso' }, { key: 'valor', label: 'Valor' }, { key: 'prc', label: 'PRC' }]
            } else if (soArea) {
                colunas = [{ key: 'area', label: 'Área' }, { key: 'valor', label: 'Valor' }, { key: 'prc', label: 'PRC' }]
            } else {
                colunas = [{ key: 'peso', label: 'Peso' }, { key: 'area', label: 'Área' }, { key: 'valor', label: 'Valor' }, { key: 'prc', label: 'PRC' }]
            }

            const quantidadeColunas = colunas.length;
            const subLabels = colunas.map(c => c.label);

            const linhaHeaders    = ['Representante', 'Grupo', 'Cliente'];
            const linhaSubHeaders = ['', '', ''];
            const mesclagens = [
                { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }, // Representante
                { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } }, // Grupo
                { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } }, // Cliente
            ];

            mesesSelecionados.forEach((mes, i) => {
                const colunaInicio = 3 + i * quantidadeColunas;
                
                linhaHeaders.push(mes.label, ...Array(quantidadeColunas - 1).fill(''));
                linhaSubHeaders.push(...subLabels);
                mesclagens.push({ s: { r: 0, c: colunaInicio }, e: { r: 0, c: colunaInicio + quantidadeColunas - 1 } });
            });

            linhaHeaders.push(
                'Total', ...Array(quantidadeColunas - 1).fill(''),
                'Carteira', ...Array(quantidadeColunas - 1).fill(''),
                `Total Mês ${mesAtual.label}`, ...Array(quantidadeColunas - 1).fill(''),
                `Carteira ${proximoMes.label}`, ...Array(quantidadeColunas - 1).fill('')
            );

            linhaSubHeaders.push(
                ...subLabels,
                ...subLabels,
                ...subLabels,
                ...subLabels
            );

            const dados = dadosRelatorio.map(item => [
                item.nomeVendedor,
                item.grupoCliente,
                item.nomeCliente,
                ...mesesSelecionados.flatMap(({ key }) =>
                    colunas.map(col => item.meses[key][col.key])
                ),
                // Total
                ...colunas.map(col => item[`${col.key}Total`]),
                // Carteira
                ...colunas.map(col => item[`carteiraMesAtual${capitalize(col.key)}`]),
                // Mês Atual
                ...colunas.map(col => item[`mesAtual${capitalize(col.key)}`]),
                // PróximoMês
                ...colunas.map(col => item[`carteiraProximoMes${capitalize(col.key)}`]),
            ]);

            const planilha = XLSX.utils.aoa_to_sheet([linhaHeaders, linhaSubHeaders, ...dados]);
            planilha['!merges'] = mesclagens;
            planilha['!cols']   = [{ wch: 24 }, ...Array(linhaSubHeaders.length - 1).fill({ wch: 16 })];

            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, planilha, 'Relatório de Vendas Carteira');
            XLSX.writeFile(workbook, 'relatorio_vendas_carteira.xlsx');
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
        <h3 class="mb-4">Relatório de Vendas Carteira</h3>

        <form name="formularioFiltro" id="formularioFiltro" method="post">                              
            <div class="card p-3 mb-4">
                <div class="row g-3">
                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroFilial" class="form-label">Filial</label>
                        <select required class="form-select form-select-sm select2" id="filtroFilial" name="filtroFilial" data-placeholder="Todas as filiais">
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
                        <label for="filtroAno" class="form-label">Ano</label>
                        <select required class="form-select form-select-sm select2" id="filtroAno" name="filtroAno" data-placeholder="Todos">
                            <option value="">Todos</option>
                        </select>
                    </div>

                    <div class="col-12 col-md-4 col-lg-3">
                        <label for="filtroMes" class="form-label">Mês</label>
                        <select required class="form-select form-select-sm select2" id="filtroMes" name="filtroMes" data-placeholder="Todos os meses">
                            <option value="01">Janeiro</option>
                            <option value="02">Fevereiro</option>
                            <option value="03">Março</option>
                            <option value="04">Abril</option>
                            <option value="05">Maio</option>
                            <option value="06">Junho</option>
                            <option value="07">Julho</option>
                            <option value="08">Agosto</option>
                            <option value="09">Setembro</option>
                            <option value="10">Outubro</option>
                            <option value="11">Novembro</option>
                            <option value="12">Dezembro</option>
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
                        <label for="filtroClienteGrupo" class="form-label">Clientes Grupo</label>
                        <select class="form-select form-select-sm select2" id="filtroClienteGrupo" name="filtroClienteGrupo" data-placeholder="Todos os grupos">
                            <option value="T">Todos</option>
                            <option value="S">Somente Grupo</option>
                            <option value="N">Somente Clientes</option>
                        </select>
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
