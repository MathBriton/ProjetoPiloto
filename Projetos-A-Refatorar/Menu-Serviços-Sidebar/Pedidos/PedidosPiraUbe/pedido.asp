<!DOCTYPE html>
<!--#Include File="../../cnn/DADOSADV.asp"-->
<!--#include file="../../includes/variaveis.asp" -->
<!--#include file="../../funcoes/funcoes.asp" -->
<% 
'response.Expires = -1
'response.ExpiresAbsolute = Now() -1 
response.Buffer  = true
Response.AddHeader "Pragma", "no-store"
Response.AddHeader "Cache-Control", "no-store, no-cache, must-revalidate"
Response.AddHeader "Expires", "0"
session.LCID = 1046

select case CATCODIGO
    case 1: Nivel = 0
    case 2: Nivel = 2
    case 3: Nivel = 3
    case 4: Nivel = 1
end select 
'ATENÇÃO: FOI INSERIDO O FORM SELECIONAPLANTA.ASP, QUE INDICA QUAL PLANTA ESTÁ EMITINDO O PEDIDOS
'NÃO MAIS A UNIDADE PRODUTORA COMO ERA ANTES. ESSA ALTERAÇÃO TERÁ IMPACTO EM VARIOS PONTOS NO PROCESSO DE INCLUSAO DO PEDIDO
'EXEMPLO: CARGA MAQUINA, PROC DE INCLUSÃO DO PEDIDO
Planta = request("paramPlanta")
descricaoPlanta = fnomePlanta(Planta)
%>
<html lang="pt-br">
  <head>
    <meta charset="UTF-8" />
    <meta http-equiv="PRAGMA" content="no-cache" />
    <meta http-equiv="cache-control" content="no-cache" />
    <title>Emissão de Pedido de Caixa</title>

    <style>
    .msg-lote-piloto{
        background: #fff3cd;
        border: 1px solid #ffe69c;
        color: #856404;
        padding: 8px 12px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 600;
        margin-top: 5px;
        margin-bottom: 5px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }

    /* === CÉLULAS DE PESO / ÁREA / PALETES === */
    #msgPeso {
        padding: 4px;
    }

    .info-cells-wrapper {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 5px;
        width: 100%;
        padding: 0;
        box-sizing: border-box;
    }

    .info-cell {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        min-width: 0;
        min-height: 44px;
        padding: 6px 8px;
        box-sizing: border-box;
        background: #ffffff;
        border: 1px solid #6BA539;
        border-radius: 3px;
        line-height: 1.25;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
    }

    .info-cell .cell-label {
        margin-bottom: 3px;
        color: #111111;
        font-size: 10px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.02em;
        white-space: nowrap;
    }

    .info-cell .cell-value {
        color: #5B982F;
        font-size: 13px;
        font-weight: 700;
        white-space: nowrap;
    }

    .info-cell.info-cell-full {
        grid-column: 1 / -1;
        align-items: flex-start;
        min-height: 38px;
    }

    .info-cell.info-cell-full .cell-value {
        white-space: normal;
    }

    @media (max-width: 620px) {
        .info-cells-wrapper {
            grid-template-columns: 1fr;
        }

        .info-cell.info-cell-full {
            grid-column: auto;
        }
    }
    </style>
    <script type="text/javascript">

        document.addEventListener("DOMContentLoaded", function () {

            const selectFrete = document.getElementById("tipo_frete");

            if (!selectFrete) return;

            let valorAnterior = selectFrete.value;

            selectFrete.addEventListener("change", function () {

                const valorAtual = this.value;

                tpFrete(valorAtual);

                if (valorAtual !== valorAnterior) {

                    OrcamentoPreco('<%=Planta%>', valorAtual, this);

                    valorAnterior = valorAtual;
                }

            });

        });

        function unitario(valor, campo) {
            debugger;
            var val = valor.replace(",", '.')
            val = parseFloat(val)
            if (!isNaN(val)) {
                if (val > 100) {
                    alert("ATENÇÃO!\nPreço unitário muito alto!");
                }
                if (campo.value != val) {
                   campo.value = val;
                }
            } else {
                if (valor != "") {
                    alert('Este campo é exclusivamente numérico')
                    campo.value = "0"
                } else {
                    campo.value = "0"
                }
            }
        }

        function verificaVencimento(value) {
            if (value == "090") {
                document.getElementById("data5").disabled = false;
                document.getElementById("data5").focus();
            } else {
                document.getElementById("data5").disabled = true;
            }
            document.getElementById("CondicaoPagamento").value = value;
        }
        /*validando comprimento m�ximo e m�nimo da chapa e removendo . ou , */
        function comprimentoch(valor) {
            real = parseInt(valor);
            var comprimentoMin = parseInt(document.getElementById("hidComprimentoMin").value);
            var comprimentoMax = parseInt(document.getElementById("hidComprimentoMax").value);

            if ((real > comprimentoMax || real < comprimentoMin) && (document.getElementById('condicao').value != 'M2')) {
                alert(`VALORES INVALIDOS.\nMáximo: ${comprimentoMax} mm\nMínimo= ${comprimentoMin} mm\nPor favor, clique em OK e corrija!`);
                document.getElementById('comp').value = "";
                document.getElementById('comp').focus();
            } else {
                if ((!isNaN(real)) && (valor != "") && (document.getElementById('condicao').value != 'M2')) {
                    document.getElementById('comp').value = real;
                }
            }
        }

        function Proposta() {

            if ((document.getElementById('codcli2').readOnly) && (document.getElementById('codcli2').value == "") && (document.getElementById('condicao').value == "CX")) {
                var cliente = document.getElementById('codcli').value;
                var loja = document.getElementById('lojacli').value;
                var produto = document.getElementById('produto').value;
                var entrega = document.getElementById('data1').value;
                var quantidade = parseInt(document.getElementById('quantidade').value);
                PropostaAjaxLoad('../../ajax/Proposta_ajax.asp?', 'Pro', produto, cliente, loja, entrega, quantidade)
            } else {
                document.getElementById('preco_unit').readOnly = false;
                document.getElementById('cond_pgto').disabled = false;
                document.getElementById('tipo_frete').disabled = false;
            }

        }

        /*validando largura maxima e  minima da chapa e removendo . ou , */
        function largurach(valor) {
            real = parseInt(valor)
            var larguraMin = parseInt(document.getElementById("hidLarguraMin").value);
            var larguraMax = parseInt(document.getElementById("hidLarguraMax").value);
            if ((real > larguraMax || real < larguraMin) && (document.getElementById('condicao').value != 'M2')) {
                alert(`VALORES INVÁLIDOS\nMáxima: ${larguraMax} mm\nMínima= ${larguraMin} mm\nPor favor, clique em OK e corrija.`);
                document.getElementById('larg').value = "";
                document.getElementById('larg').focus();
            } else {
                if ((!isNaN(real)) && (valor != "") && (document.getElementById('condicao').value != 'M2')) {
                    document.getElementById('larg').value = real;
                }
            }
        }

        function clienteremessa(valor1, valor2, valor3, valor4) {
            if (document.getElementById('remessa').checked) {
                if (Trim(document.getElementById('codcli').value) != '' && Trim(document.getElementById('lojacli').value) != '') {
                    if (Trim(document.getElementById('codcli').value) + '' + Trim(document.getElementById('lojacli').value) != Trim(document.getElementById('codcli2').value) + '' + Trim(document.getElementById('lojacli2').value)) {
                        if (Trim(document.getElementById('codcli2').value) != '' && Trim(document.getElementById('lojacli2').value) != '') {
                            ObjAjax('../../ajax/validaclienteremessa_ajax.asp?cliente=' + valor3 + '&loja=' + valor4 + '&clienteremessa=' + valor1 + '&lojaremessa=' + valor2 + '&codsiga=<%=CODUSER%>&nivel=<%=Nivel%>&tipo=Remessa')
                        } else {
                            alert('Cliente e Loja de remessa devem estar preenchidos.');
                            document.getElementById('codcli2').value = ''
                            document.getElementById('lojacli2').value = ''
                            document.getElementById('razao2').value = ''
                        }
                    } else {
                        alert('O Cliente ou loja de Fatura deve ser diferente do Cliente ou loja de remessa!')
                        document.getElementById('codcli2').value = ''
                        document.getElementById('lojacli2').value = ''
                        document.getElementById('razao2').value = ''
                        document.getElementById('codcli2').focus();
                    }
                } else {
                    alert('Preencha o campo Cliente, bem como a loja para faturamento.');
                    document.getElementById('codcli2').value = ''
                    document.getElementById('lojacli2').value = ''
                    document.getElementById('razao2').value = ''
                    document.getElementById('fatura').checked = true;
                    document.getElementById('codcli2').readOnly = true;
                    document.getElementById('lojacli2').readOnly = true;
                    document.getElementById('imgPesquisaClienteRemessa').style.visibility = "hidden";
                    document.getElementById('codcli').focus();
                }
            }
        }

        function ajaxProduto(valor) {
            if (Trim(document.getElementById('condicao').value) == '') {
                document.getElementById('produto').value = '';
                document.getElementById('descricao').value = '';
                document.getElementById('estrutura').value = '';
                document.getElementById('filial').value = '';
                document.getElementById('larg').value = '';
                document.getElementById('comp').value = '';
                document.getElementById('vinco_01').value = '';
                document.getElementById('vinco_02').value = '';
                document.getElementById('vinco_03').value = '';
                document.getElementById('vinco_04').value = '';
                document.getElementById('vinco_05').value = '';
                document.getElementById('pallet').value = '';
                document.getElementById('descpallet').value = '';
                alert('Selecione o tipo de pedido');
                return false
            }
            if (Trim(document.getElementById('codcli').value) != '' && Trim(document.getElementById('lojacli').value) != '') {
                if (Trim(valor) != '') {
                    ObjAjax('../../ajax/produto_ajax.asp?produto=' + Trim(valor) + '&cliente=' + Trim(document.getElementById('codcli').value) + '&codsiga=<%=CODUSER%>&nivel=<%=Nivel%>&tipo=Produto&tppedido=' + Trim(document.getElementById('condicao').value + '&eshop=' + document.getElementById('eshop').value))
                } else {
                    document.getElementById('produto').value = ''
                    document.getElementById('descricao').value = ''
                    document.getElementById('estrutura').value = ''
                    document.getElementById('filial').value = ''
                    document.getElementById('larg').value = ''
                    document.getElementById('comp').value = ''
                    document.getElementById('vinco_01').value = ''
                    document.getElementById('vinco_02').value = ''
                    document.getElementById('vinco_03').value = ''
                    document.getElementById('vinco_04').value = ''
                    document.getElementById('vinco_05').value = ''
                    document.getElementById('pallet').value = ''
                    document.getElementById('descpallet').value = ''
                }
            } else {
                alert('Preencha o campo do codigo do Cliente e a loja.');
                document.getElementById('produto').value = ''
                document.getElementById('descricao').value = ''
                document.getElementById('estrutura').value = ''
                document.getElementById('filial').value = ''
                document.getElementById('larg').value = ''
                document.getElementById('comp').value = ''
                document.getElementById('vinco_01').value = ''
                document.getElementById('vinco_02').value = ''
                document.getElementById('vinco_03').value = ''
                document.getElementById('vinco_04').value = ''
                document.getElementById('vinco_05').value = ''
                document.getElementById('pallet').value = ''
                document.getElementById('descpallet').value = ''
                document.getElementById('codcli').focus();
            }
        }
        function vazio(valor, campo) {
            if (Trim(valor) == "") {
                if (campo == "quantidade") {

                    document.getElementById('PagamentoDescricao').style.borderRadius = "8px"
                    document.getElementById('PagamentoDescricao').style.verticalAlign = "middle"
                    document.getElementById('PagamentoDescricao').style.textAlign = "center"
                    document.getElementById('PagamentoDescricao').innerHTML = "";
                    document.getElementById('CondicaoPagamento').value = '';

                    document.getElementById('TipoFreteDescricao').style.borderRadius = "8px"
                    document.getElementById('TipoFreteDescricao').style.verticalAlign = "middle"
                    document.getElementById('TipoFreteDescricao').style.textAlign = "center"
                    document.getElementById('TipoFreteDescricao').innerHTML = "";

                    document.getElementById('msgCargaMaquina').style.borderRadius = "none"
                    document.getElementById('msgCargaMaquina').style.verticalAlign = "none"
                    document.getElementById('msgCargaMaquina').style.textAlign = "none"
                    document.getElementById('msgCargaMaquina').style.color = "none";
                    document.getElementById('msgCargaMaquina').style.border = "none"
                    document.getElementById('msgCargaMaquina').style.background = "none"
                    document.getElementById('msgCargaMaquina').innerHTML = "";

                    document.getElementById('tipoFrete').value = 'C';
                    document.getElementById('data1').value = ''
                    document.getElementById('preco_unit').value = ''
                } else {
                    if (campo != 'produto') {
                        document.getElementById('lojacli').value = ''
                        document.getElementById('razao').value = ''
                        document.getElementById("fatura").checked = true;
                        document.getElementById('enable').value = 'N'
                        document.getElementById("remessa").checked = false;
                        //document.getElementById('imgPesquisaClienteRemessa').style.visibility = "hidden";
                        document.getElementById('codcli2').value = ''
                        document.getElementById('lojacli2').value = ''
                        document.getElementById('razao2').value = ''
                        document.getElementById('codcli2').readOnly = true;
                        document.getElementById('lojacli2').readOnly = true;

                        vazio('', 'produto')
                    }

                    document.getElementById('produto').value = ''
                    document.getElementById('descricao').value = ''
                    document.getElementById('estrutura').value = ''
                    document.getElementById('filial').value = ''
                    document.getElementById('larg').value = ''
                    document.getElementById('comp').value = ''
                    document.getElementById('vinco_01').value = ''
                    document.getElementById('vinco_02').value = ''
                    document.getElementById('vinco_03').value = ''
                    document.getElementById('vinco_04').value = ''
                    document.getElementById('vinco_05').value = ''

                    document.getElementById('larg').readOnly = false;
                    document.getElementById('comp').readOnly = false;
                    document.getElementById('vinco_01').readOnly = false;
                    document.getElementById('vinco_02').readOnly = false;
                    document.getElementById('vinco_03').readOnly = false;
                    document.getElementById('vinco_04').readOnly = false;
                    document.getElementById('vinco_05').readOnly = false;

                    document.getElementById('pallet').value = ''
                    document.getElementById('descpallet').value = ''
                    document.getElementById('quantidade').value = 0
                    document.getElementById("LotePiloto").checked=false;
                    document.getElementById("msgLotePiloto").value="";
                    document.getElementById("msgLotePiloto").style.display = "none";
                    document.getElementById('msgPeso').innerHTML = ''
                    document.getElementById('preco_unit').value = ''
                    document.getElementById('data1').value = ''
                    document.getElementById('ped_cliente').value = ''
                    document.getElementById('itempedcli').value = ''
                    document.getElementById('numfnc').value = ''
                    document.getElementById('descFnc').value = ''
                    document.getElementById('tes').value = ''
                    document.getElementById('descTes').value = ''

                    document.getElementById('tipoFrete').value = 'C';

                    document.getElementById('msgCargaMaquina').style.borderRadius = "none"
                    document.getElementById('msgCargaMaquina').style.verticalAlign = "none"
                    document.getElementById('msgCargaMaquina').style.textAlign = "none"
                    document.getElementById('msgCargaMaquina').style.color = "none";
                    document.getElementById('msgCargaMaquina').style.border = "none"
                    document.getElementById('msgCargaMaquina').style.background = "none"
                    document.getElementById('msgCargaMaquina').innerHTML = "";

                    document.getElementById('msgPeso').style.borderRadius = "none"
                    document.getElementById('msgPeso').style.verticalAlign = "none"
                    document.getElementById('msgPeso').style.textAlign = "none"
                    document.getElementById('msgPeso').style.color = "none";
                    document.getElementById('msgPeso').style.border = "none"
                    document.getElementById('msgPeso').style.background = "none"

                    document.getElementById('PagamentoDescricao').style.borderRadius = "8px"
                    document.getElementById('PagamentoDescricao').style.verticalAlign = "middle"
                    document.getElementById('PagamentoDescricao').style.textAlign = "center"
                    document.getElementById('PagamentoDescricao').innerHTML = "";
                    document.getElementById('CondicaoPagamento').value = '';
                    document.getElementById("cond_pgto").options[0].selected = "true";
                    document.getElementById('cond_pgto').disabled = false;

                    document.getElementById('TipoFreteDescricao').style.borderRadius = "8px"
                    document.getElementById('TipoFreteDescricao').style.verticalAlign = "middle"
                    document.getElementById('TipoFreteDescricao').style.textAlign = "center"
                    document.getElementById('TipoFreteDescricao').innerHTML = "";
                    //document.getElementById('lotepiloto').checked = false;
                    //document.getElementById("obs_lotepiloto").value = "";
                }
            }
        }

        function validaProduto(valor) {
            var produto = valor.value;
            var tipoPedido = document.getElementById('condicao');
            var tipoProduto = produto.substring(0, 3);
            tipoProduto = parseInt(tipoProduto);
            if ((Trim(produto) != "") && (Trim(tipoPedido.value) != "")) {
                switch (tipoPedido.value) {
                    case "CX":
                        if ((tipoProduto > 709) && (tipoProduto != 720) && (tipoProduto < 800)) {
                            ajaxProduto(produto)
                        } else {
                            alert('Produto incompatível com o tipo de pedido.');
                            valor.value = "";
                        }
                        break
                    case "CH":
                        if (tipoProduto = 720) {
                            ajaxProduto(produto)
                        } else {
                            alert('Produto incompatível com o tipo de pedido.')
                            valor.value = "";
                        }
                        break
                    case "PA":
                    case "M2":
                        if ((tipoProduto >= 700) && (tipoProduto <= 709)) {
                            ajaxProduto(produto)
                        } else {
                            alert('Produto incompatível com o tipo de pedido.')
                            valor.value = "";
                        }
                        break
                }
            }
        }
        function LimpaDadosProduto(){
            debugger;
            document.getElementById('quantidade').value = 0;
            document.getElementById("LotePiloto").checked=false;
            document.getElementById("msgLotePiloto").value="";
            document.getElementById("msgLotePiloto").style.display = "none";

            document.getElementById('msgPeso').style.borderRadius = "none"
            document.getElementById('msgPeso').style.verticalAlign = "none"
            document.getElementById('msgPeso').style.textAlign = "none"
            document.getElementById('msgPeso').style.color = "none";
            document.getElementById('msgPeso').style.border = "none"
            document.getElementById('msgPeso').style.background = "none"
            document.getElementById('msgPeso').innerHTML = '';
        }
        // ----------------------------------------------------------------------------------------
        function tppedido(valor) {
            vazio('', 'produto')
            switch (valor) {
            case "CX":
                document.getElementById('descpallet').disabled = false;
                document.getElementById('larg').disabled = true;
                document.getElementById('comp').disabled = true;
                document.getElementById('vinco_01').disabled = true;
                document.getElementById('vinco_02').disabled = true;
                document.getElementById('vinco_03').disabled = true;
                document.getElementById('vinco_04').disabled = true;
                document.getElementById('vinco_05').disabled = true;
                //document.getElementById('formato').disabled = true;
                document.getElementById('gramatura').disabled = true;
                document.getElementById('diametro').disabled = true;
                document.getElementById('diametro').style.width = "80px"
                document.getElementById('tubete').disabled = true;
                document.getElementById('data2').disabled = true;
                document.getElementById('data2').style.width = "80px"
                document.getElementById('preco_unit').readOnly = false;
                document.getElementById('cond_pgto').disabled = false;
                document.getElementById('tipo_frete').disabled = false;
                document.getElementById('data5').disabled = true;
                break;
            case "CH":
                document.getElementById('descpallet').disabled = false;
                document.getElementById('larg').disabled = false;
                document.getElementById('comp').disabled = false;
                document.getElementById('vinco_01').disabled = false;
                document.getElementById('vinco_02').disabled = false;
                document.getElementById('vinco_03').disabled = false;
                document.getElementById('vinco_04').disabled = false;
                document.getElementById('vinco_05').disabled = false;
                document.getElementById('gramatura').disabled = true;
                //document.getElementById('formato').disabled = true;
                document.getElementById('diametro').disabled = true;
                document.getElementById('diametro').style.width = "80px"
                document.getElementById('tubete').disabled = true;
                document.getElementById('data2').disabled = true;
                document.getElementById('data2').style.width = "80px"
                document.getElementById('preco_unit').readOnly = false;
                document.getElementById('cond_pgto').disabled = false;
                document.getElementById('tipo_frete').disabled = false;
                document.getElementById('data5').disabled = false;
                break;
            case "PA":
                document.getElementById('descpallet').disabled = true;
                document.getElementById('larg').disabled = true;
                document.getElementById('comp').disabled = true;
                document.getElementById('vinco_01').disabled = true;
                document.getElementById('vinco_02').disabled = true;
                document.getElementById('vinco_03').disabled = true;
                document.getElementById('vinco_04').disabled = true;
                document.getElementById('vinco_05').disabled = true;
                document.getElementById('gramatura').disabled = false;
                //document.getElementById('formato').disabled = false;
                document.getElementById('diametro').disabled = false;
                document.getElementById('diametro').style.width = "80px"
                document.getElementById('tubete').disabled = false;
                document.getElementById('data2').disabled = false;

                document.getElementById('preco_unit').readOnly = false;
                document.getElementById('cond_pgto').disabled = false;
                document.getElementById('tipo_frete').disabled = false;
                document.getElementById('data5').disabled = false;
                Gramatura(0)
                /*
                document.getElementById('tipo_frete').options[0].selected = "true";
                document.getElementById('FreteUnitario').value = 0;                        
                document.getElementById('FreteUnitario').disabled = false; // desabilita    
                */

                break;
            case "M2":
                document.getElementById('descpallet').disabled = true;
                document.getElementById('larg').disabled = false;
                document.getElementById('comp').disabled = false;

                document.getElementById('vinco_01').disabled = true;
                document.getElementById('vinco_02').disabled = true;
                document.getElementById('vinco_03').disabled = true;
                document.getElementById('vinco_04').disabled = true;
                document.getElementById('vinco_05').disabled = true;
                document.getElementById('gramatura').disabled = false;
                //document.getElementById('formato').disabled = false;
                document.getElementById('diametro').disabled = false;
                document.getElementById('diametro').style.width = "80px"
                document.getElementById('tubete').disabled = false;
                document.getElementById('data2').disabled = true;
                document.getElementById('preco_unit').readOnly = false;
                document.getElementById('cond_pgto').disabled = false;
                document.getElementById('tipo_frete').disabled = false;
                document.getElementById('data5').disabled = false;
                Gramatura(1)

                break;
            }

            var eShopSel = document.getElementById('eshop').value;

            console.log(eShopSel);

            if ( eShopSel != null) {
                SelecionaEShop( eShopSel );
            }
        }        
        // ----------------------------------------------------------------------------------------
        function SelecionaEShop(obj) {

            var tipoPedido = document.getElementById('condicao').value;

            console.log( 'tipoPedido: ' + tipoPedido );
            console.log( 'eShop: ' + obj );

            if (tipoPedido == "CX" || tipoPedido == "CH") {

                document.getElementById('tipo_frete').options[0].selected = "true";
                document.getElementById('tipo_frete').disabled = false;

                document.getElementById('FreteUnitario').value = 0;                        
                document.getElementById('FreteUnitario').disabled = true; // desabilita    

            } else { // PAPEL

                document.getElementById('tipo_frete').options[0].selected = "true";
                document.getElementById('tipo_frete').disabled = false;

                if (obj == "0") { // eShop NAO
                    document.getElementById('FreteUnitario').value = 0;                        
                    document.getElementById('FreteUnitario').disabled = false; // desabilita    
                } else { // eShop SIM
                    document.getElementById('FreteUnitario').value = 0;                        
                    document.getElementById('FreteUnitario').disabled = true; // habilita    
                }

            }
        }
        // ----------------------------------------------------------------------------------------
        // COMENTEI EM 03/05/2018 POIS AGORA OS PEDIDOS DE PAPEL TERAO O FRETE TRATADO DA MESMA FORA QUE PO
        function tpFrete(valor) {

            var tipoPedido = document.getElementById('condicao').value;
            var eShopSel = document.getElementById('eshop').value;
            var tipoFrete = valor;            
            var ativa = false;

            if (tipoPedido == "PA") { // PAPEL ----------------------------------------------------

                if (eShopSel == "0") {  // eShop NÃO ----------------------------------------------

                    if (tipoFrete == "C") { // CIF
                        document.getElementById('FreteUnitario').disabled = false;
                        document.getElementById('FreteUnitario').style.width = "80px"
                        document.getElementById('FreteUnitario').value = 0;
                    } else { // FOB ---------------------------------------------------------------
                        document.getElementById('FreteUnitario').disabled = true;
                        document.getElementById('FreteUnitario').style.width = "80px"
                        document.getElementById('FreteUnitario').value = 0;
                    }
                } else { // eShop SIM -------------------------------------------------------------
                    if (tipoFrete == "C") { // CIF
                        document.getElementById('FreteUnitario').disabled = true;
                        document.getElementById('FreteUnitario').style.width = "80px"
                        document.getElementById('FreteUnitario').value = 0;
                    } else { // FOB ---------------------------------------------------------------
                        document.getElementById('FreteUnitario').disabled = false;
                        document.getElementById('FreteUnitario').style.width = "80px"
                        document.getElementById('FreteUnitario').value = 0;
                    }
                }

            } else {

                if (eShopSel == "1") { // eShop SIM -----------------------------------------------

                    //document.getElementById('tipo_frete').value = "F";
                    document.getElementById('tipo_frete').disabled = false;

                    if (valor == "F") {
                        ativa = false;
                    } else {
                        ativa = true;
                    }

                    document.getElementById('FreteUnitario').disabled = ativa;
                    document.getElementById('FreteUnitario').style.width = "80px"
                    document.getElementById('FreteUnitario').value = 0;
                }
            }
        }
        // ----------------------------------------------------------------------------------------
        function ApagaData(valor) {
            document.getElementById('msgCargaMaquina').style.borderRadius = "none";
            document.getElementById('msgCargaMaquina').style.verticalAlign = "none";
            document.getElementById('msgCargaMaquina').style.textAlign = "none";
            document.getElementById('msgCargaMaquina').style.color = "none";
            document.getElementById('msgCargaMaquina').style.border = "none";
            document.getElementById('msgCargaMaquina').style.background = "none";
            document.getElementById('msgCargaMaquina').innerHTML = "";
            document.getElementById(valor).value = ""
        }
        function Gramatura(Tipo) {
            var comboGramatura = document.getElementById("gramatura");
            while (comboGramatura.length) {
                comboGramatura.remove(0);
            }
            if (Tipo == 1) {
                var opt0 = document.createElement("option");
                opt0.value = "";
                opt0.text = "";
                comboGramatura.add(opt0, comboGramatura.options[0]);

                var opt1 = document.createElement("option");
                opt1.value = "38";
                opt1.text = "38";
                comboGramatura.add(opt1, comboGramatura.options[1]);

                var opt2 = document.createElement("option");
                opt2.value = "45";
                opt2.text = "45";
                comboGramatura.add(opt2, comboGramatura.options[2]);

                var opt3 = document.createElement("option");
                opt3.value = "50";
                opt3.text = "50";
                comboGramatura.add(opt3, comboGramatura.options[3]);

                var opt4 = document.createElement("option");
                opt4.value = "55";
                opt4.text = "55";
                comboGramatura.add(opt4, comboGramatura.options[4]);

                var opt5 = document.createElement("option");
                opt5.value = "60";
                opt5.text = "60";
                comboGramatura.add(opt5, comboGramatura.options[5]);

                var opt6 = document.createElement("option");
                opt6.value = "65";
                opt6.text = "65";
                comboGramatura.add(opt6, comboGramatura.options[6]);

                var opt7 = document.createElement("option");
                opt7.value = "70";
                opt7.text = "70";
                comboGramatura.add(opt7, comboGramatura.options[7]);

            } else {

                var opt0 = document.createElement("option");
                opt0.value = "";
                opt0.text = "";
                comboGramatura.add(opt0, comboGramatura.options[0]);

                var opt1 = document.createElement("option");
                opt1.value = "100";
                opt1.text = "100";
                comboGramatura.add(opt1, comboGramatura.options[1]);

                var opt2 = document.createElement("option");
                opt2.value = "105";
                opt2.text = "105";
                comboGramatura.add(opt2, comboGramatura.options[2]);

                var opt3 = document.createElement("option");
                opt3.value = "110";
                opt3.text = "110";
                comboGramatura.add(opt3, comboGramatura.options[3]);

                var opt4 = document.createElement("option");
                opt4.value = "115";
                opt4.text = "115";
                comboGramatura.add(opt4, comboGramatura.options[4]);

                var opt5 = document.createElement("option");
                opt5.value = "120";
                opt5.text = "120";
                comboGramatura.add(opt5, comboGramatura.options[5]);

                var opt6 = document.createElement("option");
                opt6.value = "125";
                opt6.text = "125";
                comboGramatura.add(opt6, comboGramatura.options[6]);

                var opt7 = document.createElement("option");
                opt7.value = "130";
                opt7.text = "130";
                comboGramatura.add(opt7, comboGramatura.options[7]);

                var opt8 = document.createElement("option");
                opt8.value = "135";
                opt8.text = "135";
                comboGramatura.add(opt8, comboGramatura.options[8]);

                var opt9 = document.createElement("option");
                opt9.value = "140";
                opt9.text = "140";
                comboGramatura.add(opt9, comboGramatura.options[9]);

                var opt10 = document.createElement("option");
                opt10.value = "145";
                opt10.text = "145";
                comboGramatura.add(opt10, comboGramatura.options[10]);

                var opt11 = document.createElement("option");
                opt11.value = "150";
                opt11.text = "150";
                comboGramatura.add(opt11, comboGramatura.options[11]);

                var opt12 = document.createElement("option");
                opt12.value = "155";
                opt12.text = "155";
                comboGramatura.add(opt12, comboGramatura.options[12]);

                var opt13 = document.createElement("option");
                opt13.value = "160";
                opt13.text = "160";
                comboGramatura.add(opt13, comboGramatura.options[13]);

                var opt14 = document.createElement("option");
                opt14.value = "165";
                opt14.text = "165";
                comboGramatura.add(opt14, comboGramatura.options[14]);

                var opt15 = document.createElement("option");
                opt15.value = "170";
                opt15.text = "170";
                comboGramatura.add(opt15, comboGramatura.options[15]);

                var opt16 = document.createElement("option");
                opt16.value = "175";
                opt16.text = "175";
                comboGramatura.add(opt16, comboGramatura.options[16]);

                var opt17 = document.createElement("option");
                opt17.value = "180";
                opt17.text = "180";
                comboGramatura.add(opt17, comboGramatura.options[17]);

                var opt18 = document.createElement("option");
                opt18.value = "185";
                opt18.text = "185";
                comboGramatura.add(opt18, comboGramatura.options[18]);

                var opt19 = document.createElement("option");
                opt19.value = "190";
                opt19.text = "190";
                comboGramatura.add(opt19, comboGramatura.options[19]);

                var opt20 = document.createElement("option");
                opt20.value = "195";
                opt20.text = "195";
                comboGramatura.add(opt20, comboGramatura.options[20]);

                var opt21 = document.createElement("option");
                opt21.value = "200";
                opt21.text = "200";
                comboGramatura.add(opt21, comboGramatura.options[21]);

                var opt22 = document.createElement("option");
                opt22.value = "205";
                opt22.text = "205";
                comboGramatura.add(opt22, comboGramatura.options[22]);

                var opt23 = document.createElement("option");
                opt23.value = "210";
                opt23.text = "210";
                comboGramatura.add(opt23, comboGramatura.options[23]);

                var opt24 = document.createElement("option");
                opt24.value = "215";
                opt24.text = "215";
                comboGramatura.add(opt24, comboGramatura.options[24]);

                var opt25 = document.createElement("option");
                opt25.value = "220";
                opt25.text = "220";
                comboGramatura.add(opt25, comboGramatura.options[25]);

                var opt26 = document.createElement("option");
                opt26.value = "230";
                opt26.text = "230";
                comboGramatura.add(opt26, comboGramatura.options[26]);

                var opt27 = document.createElement("option");
                opt27.value = "240";
                opt27.text = "240";
                comboGramatura.add(opt27, comboGramatura.options[27]);

                var opt28 = document.createElement("option");
                opt28.value = "250";
                opt28.text = "250";
                comboGramatura.add(opt28, comboGramatura.options[28]);
            }
        }
        function ValidaLotePiloto(){
            console.log("Dentro da validalotepiloto");
            document.getElementById("msgLotePiloto").style.display = "none";
            debugger;
            if (window.XMLHttpRequest) {// code for IE7+, Firefox, Chrome, Opera, Safari
                xmlhttp = new XMLHttpRequest();
            } else {// code for IE6, IE5
                xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
            }
            
            let resposta=false;
            xmlhttp.onreadystatechange = function () {
                if (this.readyState == 4 && this.status == 200){
                    console.log('xmlhttp.responseText',xmlhttp.responseText);
                    var Mensagem = '';
                    document.getElementById("LotePiloto").checked=false;
                    document.getElementById("obs_lotepiloto").value = ""
                    if(xmlhttp.responseText!=''){
                        
                         var dados = xmlhttp.responseText.split('|');
                         if(dados.length>0){
                            console.log('dados 0', dados[0])
                           if(dados[0] == '0')
                              resposta=false;
                           else{
                              resposta=true;
                              if(dados[1] == '0'){
                                 Mensagem =  dados[11];
                              }else{
                                 console.log('dados 2', dados[2])
                                 if(dados[2] == '0'){
                                    Mensagem = 'PEDIDO BLOQUEADO! QUANTIDADE DE PALETES SOLICITADA MAIOR QUE A PERMITIDA PARA O LOTE PILOTO!';
                                 }   
                              }
                           } 
                         }
                    }
                    if(resposta && Mensagem == ''){
                      console.log('Marcando Lote Piloto')
                      document.getElementById("LotePiloto").checked=true;
                    
                      document.getElementById("obs_lotepiloto").value = ""
                      document.getElementById("obs_lotepiloto").value += 'ACOMPANHAR ';
                      if(dados[3] == '1'){ // VALIDA_ONDULADEIRA
                          document.getElementById("obs_lotepiloto").value += '- ONDULADEIRA ';
                        } 
        
                        if(dados[4] == '1'){ // VALIDA_IMPRESSORA
                          document.getElementById("obs_lotepiloto").value += ' - IMPRESSORA ';
                        }
        
                        if(dados[5] == '1'){ // VALIDA_EXPEDIÇÃO
                          document.getElementById("obs_lotepiloto").value += ' - EXPEDICAO ';
                        }
                        
                        if(dados[6] == '1'){ // VALIDA_CLIENTE
                          document.getElementById("obs_lotepiloto").value += ' - CLIENTE ';
                        } 
                        
                        if(dados[7] == '1'){ // VALIDA_DDP
                           document.getElementById("obs_lotepiloto").value += ' - DDP ';
                        } 

                        if(dados[8] == '1'){ // VALIDA_CLICHERIA
                           document.getElementById("obs_lotepiloto").value += ' - CLICHERIA ';
                        } 

                   }
                    console.log('OBS LOTE PILOTO', document.getElementById("obs_lotepiloto").value)
        
                   if(resposta && Mensagem != ''){
                     alert(Mensagem);
                     return;
                  }

                  //console.log('VerificaOrcamento')
                  
                  /*VerificaOrcamento(document.getElementById("filial").value
                                , document.getElementById("codcli").value
                                , document.getElementById("lojacli").value
                                , document.getElementById("produto").value
                                , Trim(document.getElementById("tipo_frete").value)
                                , Trim(document.getElementById("cond_pgto").value)
                                , Trim(document.getElementById("preco_unit").value));*/

                    // CHAMA O AJAX SOMENTE APÓS VALIDAR O LOTE PILOTO

                    console.log("validapeso_ajax",document.getElementById('hidPlanta').value)
                    ajaxload('../../ajax/validapeso_ajax.asp?',
                        document.getElementById('codcli').value,
                        document.getElementById('lojacli').value,
                        '',
                        '',
                        document.getElementById('produto').value,
                        '',
                        '',
                        '',
                        '',
                        document.getElementById('hidPlanta').value,
                        'quantidade',
                        'msgPeso',
                        'Pes',
                        document.getElementById('quantidade').value,
                        document.getElementById('larg').value,
                        document.getElementById('comp').value,
                        document.getElementById('condicao').value
                    );

                    
                    
              }
            }
            
            var urlLotePiloto = `PedidoLotePilotoRepositorio.asp?paramCliente=${document.getElementById("codcli").value.trim()}&paramLoja=${document.getElementById("lojacli").value.trim()}&paramProduto=${document.getElementById("produto").value.trim()}&paramPlanta=${'<%=Planta%>'}&paramQuantidade=${document.getElementById("quantidade").value.trim()}`
            console.log('urlLotePiloto', urlLotePiloto);
            xmlhttp.open("GET", urlLotePiloto);
            xmlhttp.send()
        }
    </script>

    <link rel="stylesheet" type="text/css" href="../../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.css" />
    <link rel="stylesheet" type="text/css" href="../../../dhtmlx/dhtmlxCalendar/codebase/skins/dhtmlxcalendar_dhx_web.css" />
    <script src="../../../dhtmlx/dhtmlxCalendar/codebase/dhtmlxcalendar.js"></script>
    <script src="../../funcoes/funcoesJAVASCRIPT.js?<%=time() %>"></script>
    <script src="includes/js/PedidoFuncoes.js?v=<%=Now()%>"></script>
    <link href='http://fonts.googleapis.com/css?family=Hammersmith+One' rel='stylesheet' type='text/css' />
    <link href='http://fonts.googleapis.com/css?family=Aldrich' rel='stylesheet' type='text/css' />
    <link rel="stylesheet" type="text/css" href="../../css/padrao.css" />

    <script type="text/javascript">
        function DiametroPapel(valor, campo) {
            real(valor, campo)
            if (Trim(campo.value) != 0) {
                var diametro = campo.value;
                maximo = 1.45
                minimo = 0.20
                //if ((document.getElementById('formato').value < 2500) || (document.getElementById('formato').value > 300)) {
                //    if (diametro > 1.45 || diametro < 0.20) {
                //        alert("Diâmetro permitido ultrapassado\nDiâmetro Máximo: 1,45\nDiâmetro Mínimo: 0,20");
                //        campo.value = 0;
                //        campo.focus();
                //    }
                //}
            }
        }

        var myCalendar;

        function doOnLoad() {

            myCalendar = new dhtmlXCalendarObject(["data1", "data5", "data2", "dataDescarga"]);
            myCalendar.setDate(myCalendar.getDate());
            myCalendar.loadUserLanguage("br");

            if (Trim(document.getElementById('codusu').value) == '') {
                document.getElementById('codusu').value = '<%=CODUSER%>'
            }

            ajaxUsuario(document.getElementById('codusu').value, 'Usr')

            var tipo = ""
            if (document.getElementById("condicao").options[1].selected) {
                tipo = "Caixa"
            }

            if (document.getElementById("condicao").options[2].selected) {
                tipo = "Chapa"
            }

            if (document.getElementById("condicao").options[3].selected) {
                tipo = "Papel"
            }

            continua = false;

            if (tipo != "") {
                continua = confirm("Deseja continuar emitindo pedido de " + tipo + "?")
            }

            if (!continua) {
                document.getElementById('eshop').value = '0'
                document.getElementById('lojacli').value = ''
                document.getElementById('razao').value = ''
                document.getElementById("fatura").checked = true;
                document.getElementById('enable').value = 'N'
                document.getElementById("remessa").checked = false;
                //document.getElementById('imgPesquisaClienteRemessa').style.visibility = "hidden";
                document.getElementById('codcli2').value = ''
                document.getElementById('lojacli2').value = ''
                document.getElementById('razao2').value = ''
                document.getElementById('codcli2').readOnly = true;
                document.getElementById('lojacli2').readOnly = true;

                document.getElementById('produto').value = ''
                document.getElementById('descricao').value = ''
                document.getElementById('estrutura').value = ''
                document.getElementById('filial').value = ''
                document.getElementById('larg').value = ''
                document.getElementById('comp').value = ''
                document.getElementById('vinco_01').value = ''
                document.getElementById('vinco_02').value = ''
                document.getElementById('vinco_03').value = ''
                document.getElementById('vinco_04').value = ''
                document.getElementById('vinco_05').value = ''

                document.getElementById('pallet').value = ''
                document.getElementById('descpallet').value = ''
                document.getElementById('quantidade').value = 0
                document.getElementById("LotePiloto").checked=false;
            document.getElementById("msgLotePiloto").value="";
            document.getElementById("msgLotePiloto").style.display = "none";
                document.getElementById('msgPeso').innerHTML = ''
                document.getElementById('preco_unit').value = 0
                document.getElementById('data1').value = ''
                document.getElementById('msgCargaMaquina').innerHTML = ''
                document.getElementById('ped_cliente').value = ''
                document.getElementById('itempedcli').value = ''
                document.getElementById('numfnc').value = ''
                document.getElementById('descFnc').value = ''
                document.getElementById('tes').value = ''
                document.getElementById('descTes').value = ''
                document.getElementById('gramatura').value = ''
                document.getElementById('diametro').value = ''
                document.getElementById('tubete').value = ''
                document.getElementById('data2').value = ''

                document.getElementById('descpallet').disabled = false;
                document.getElementById('larg').disabled = false;
                document.getElementById('comp').disabled = false;
                document.getElementById('vinco_01').disabled = false;
                document.getElementById('vinco_02').disabled = false;
                document.getElementById('vinco_03').disabled = false;
                document.getElementById('vinco_04').disabled = false;
                document.getElementById('vinco_05').disabled = false;
                document.getElementById('larg').readOnly = true;
                document.getElementById('comp').readOnly = true;
                document.getElementById('vinco_01').readOnly = true;
                document.getElementById('vinco_02').readOnly = true;
                document.getElementById('vinco_03').readOnly = true;
                document.getElementById('vinco_04').readOnly = true;
                document.getElementById('vinco_05').readOnly = true;
                //document.getElementById('formato').disabled = true;
                document.getElementById('gramatura').disabled = true;
                document.getElementById('diametro').disabled = true;
                document.getElementById('tubete').disabled = true;
                document.getElementById('data2').disabled = true;

                document.getElementById('preco_unit').readOnly = false;
                document.getElementById('cond_pgto').disabled = false;
                document.getElementById('tipo_frete').disabled = false;
                document.getElementById('data5').disabled = true;

                document.getElementById('msgCargaMaquina').style.borderRadius = "none"
                document.getElementById('msgCargaMaquina').style.verticalAlign = "none"
                document.getElementById('msgCargaMaquina').style.textAlign = "none"
                document.getElementById('msgCargaMaquina').style.color = "none";
                document.getElementById('msgCargaMaquina').style.border = "none"
                document.getElementById('msgCargaMaquina').style.background = "none"

                document.getElementById('msgPeso').style.borderRadius = "none"
                document.getElementById('msgPeso').style.verticalAlign = "none"
                document.getElementById('msgPeso').style.textAlign = "none"
                document.getElementById('msgPeso').style.color = "none";
                document.getElementById('msgPeso').style.border = "none"
                document.getElementById('msgPeso').style.background = "none"
                document.getElementById("condicao").options[0].selected = "true";
                document.getElementById('CondicaoPagamento').value = '';
            } else {

                validaProduto(document.getElementById('produto'))

                if (tipo == "Caixa") {
                    document.getElementById('preco_unit').value = 0;
                    document.getElementById('preco_unit').readOnly = false;

                    document.getElementById("cond_pgto").options[0].selected = "true";
                    document.getElementById('cond_pgto').disabled = false;
                    document.getElementById('CondicaoPagamento').value = '';
                    document.getElementById('PagamentoDescricao').innerHTML = '';
                    document.getElementById('data5').disabled = true;

                    document.getElementById('tipo_frete').disabled = false;

                    document.getElementById('data1').value = '';
                    document.getElementById('msgCargaMaquina').innerHTML = ''
                    document.getElementById('msgCargaMaquina').style.borderRadius = "none"
                    document.getElementById('msgCargaMaquina').style.verticalAlign = "none"
                    document.getElementById('msgCargaMaquina').style.textAlign = "none"
                    document.getElementById('msgCargaMaquina').style.color = "none";
                    document.getElementById('msgCargaMaquina').style.border = "none"
                    document.getElementById('msgCargaMaquina').style.background = "none"
                }

                if (tipo == "Papel") {
                    //document.getElementById('formato').disabled = false;
                    document.getElementById('gramatura').disabled = false;
                    document.getElementById('diametro').disabled = false;
                    document.getElementById('tubete').disabled = false;
                    document.getElementById('data2').disabled = false;
                    Gramatura(0);
                }
                if (tipo == "Papel M2") {
                    document.getElementById('larg').readOnly = true;
                    document.getElementById('comp').readOnly = false;
                    //document.getElementById('formato').disabled = false;
                    document.getElementById('gramatura').disabled = false;
                    document.getElementById('diametro').disabled = false;
                    document.getElementById('tubete').disabled = false;
                    document.getElementById('data2').disabled = true;
                    document.getElementById('quant_pedido').disabled = true;
                    Gramatura(1)
                }
            }
        }
        function QuantidadeMaxima(quantidade) {
            if (quantidade > 909090) {
                document.getElementById("quantidade").value = 0;
                alert("Quantidade máxima permitida, 909090");
                return false;
            }
        }
        //function SelecaoEShop(objeto) {
        //    vazio('', 'produto');
        //    if (objeto.value != '1') {
        //        document.getElementById('CondicaoPagamento').value = '';
        //        document.getElementById('PagamentoDescricao').style.borderRadius = "8px"
        //        document.getElementById('PagamentoDescricao').style.verticalAlign = "middle"
        //        document.getElementById('PagamentoDescricao').style.textAlign = "center"
        //        document.getElementById('PagamentoDescricao').innerHTML = '';
        //        document.getElementById("cond_pgto").options[0].selected = "true";
        //        document.getElementById('cond_pgto').disabled = true;
        //    }
        //}
    </script>
  </head>

  <body class="fundo_formulario" onload="doOnLoad();ParametrosFiliais('<%=Planta%>')">

    <form name="form" id="form" method="post" action="../validacao/ValidaPedidoVenda.asp">

      <input type="hidden" name="pallet" id="pallet" />
      <input type="hidden" name="enable" id="enable" value="N" />
      <input type="hidden" name="codusu" id="codusu"/>
      <input type="hidden" name="area" id="area"/>
      <input type="hidden" name="obs_lotepiloto" id="obs_lotepiloto" value=""/>
      <input type="hidden" name="hidProdutoEstoque" id="hidProdutoEstoque" value="" />
      <input type="hidden" name="hidDiasTransporte" id="hidDiasTransporte" value="" />
      <input type="hidden" name="hidCargaMaquina" id="hidCargaMaquina" value="" />
      <input type="hidden" name="hidLimiteDeCredito" id="hidLimiteDeCredito" value="" />
      <input type="hidden" name="hidFlagDiasUtil" id="hidFlagDiasUtil" value="" />
      <input type="hidden" name="hidLarguraMin" id="hidLarguraMin" value="" />
      <input type="hidden" name="hidLarguraMax" id="hidLarguraMax" value="" />
      <input type="hidden" name="hidComprimentoMin" id="hidComprimentoMin" value="" />
      <input type="hidden" name="hidComprimentoMax" id="hidComprimentoMax" value="" />
      <input type="hidden" name="hidDiasAcrescimoPA" id="hidDiasAcrescimoPA" value="" />
      <input type="hidden" name="hidDiasAcrescimoCX" id="hidDiasAcrescimoCX" value="" />
      <input type="hidden" name="hidDiasAcrescimoCH" id="hidDiasAcrescimoCH" value="" />
      <input type="hidden" name="hidFlagDiasViagem" id="hidFlagDiasViagem" value="" />
      <input type="hidden" name="hidFlagDiaUtil" id="hidFlagDiaUtil" value="" />
      <input type="hidden" name="hidBloqueiaMaquinaCaixa" id="hidBloqueiaMaquinaCaixa" value="" />
      <input type="hidden" name="hidBloqueiaMaquinaChapa" id="hidBloqueiaMaquinaChapa" value="" />
      <input type="hidden" name="hidBloqueiaMaquinaCorteVinco" id="hidBloqueiaMaquinaCorteVinco" value="" />
      <input type="hidden" name="hidBloqueiaMaletaDE" id="hidBloqueiaMaletaDE" value="" />
      <input type="hidden" name="hidBloqueiaMaletaATE" id="hidBloqueiaMaletaATE" value="" />
      <input type="hidden" name="hidBloqueiaCVDE" id="hidBloqueiaCVDE" value="" />
      <input type="hidden" name="hidBloqueiaCVATE" id="hidBloqueiaCVATE" value="" />
      <input type="hidden" name="hidBloqueiaCHDE" id="hidBloqueiaCHDE" value="" />
      <input type="hidden" name="hidBloqueiaCHATE" id="hidBloqueiaCHATE" value="" />
      <input type="hidden" name="hidMtsCVmin" id="hidMtsCVmin" value="" />
      <input type="hidden" name="hidMtsCVmax" id="hidMtsCVmax" value="" />
      <input type="hidden" name="hidMtsCXmin" id="hidMtsCXmin" value="" />
      <input type="hidden" name="hidMtsCXmax" id="hidMtsCXmax" value="" />
      <input type="hidden" name="hidMtsCHmin" id="hidMtsCHmin" value="" />
      <input type="hidden" name="hidMtsCHmax" id="hidMtsCHmax" value="" />
      <input type="hidden" name="hidPrecoKgCorteVinco" id="hidPrecoKgCorteVinco" value="" />
      <input type="hidden" name="hidPrecoKgMaleta" id="hidPrecoKgMaleta" value="" />
      <input type="hidden" name="hidPrecoKgChapa" id="hidPrecoKgChapa" value="" />
      <input type="hidden" name="hidPeso" id="hidPeso" value="0" />
      <input type="hidden" name="hidPesoMN" id="hidPesoMN" value="" />
      <input type="hidden" name="hidPesoCV" id="hidPesoCV" value="" />
      <input type="hidden" name="hidPesoCH" id="hidPesoCH" value="" />
      <input type="hidden" name="hidTipoProduto" id="hidTipoProduto" value="" />
      <input type="hidden" name="hidComprimentoProduto" id="hidComprimentoProduto" value="0" />
      <input type="hidden" name="hidLarguraProduto" id="hidLarguraProduto" value="0" />
      <input type="hidden" name="hidCargaMaquinaValida" id="hidCargaMaquinaValida" value="0" />
      <input type="hidden" name="hidPlanta" id="hidPlanta" value="<%=Planta%>" />
      <input type="hidden" name="hidValorOrcamento" id="hidValorOrcamento" value="" />
<%
if nivel <> "0" then
%>
      <input type="hidden" name="eshop" id="eshop" value="0"/>
<%
end if 
%>
      <div id="formulario" style="width:55%">
        <header class="cabecalho2">Emissão de Pedido - <%=descricaoPlanta%></header>
        <hr>
        <table>
<!--        
          <tr>
            <th colspan="5" id="msgUsuario"></th>
          </tr>
-->          
          <tr>
            <th><label for="condicao">Tipo de Pedido:</label></th>
            <td colspan="4">
              <select name="condicao" id="condicao" onchange="tppedido(this.value)" tabindex="1">
                <option value="" <%if trim(request("condicao")) = "" then %> selected="selected"<%end if %>>Selecione o tipo de pedido</option>
                <option value="CX" <%if trim(request("condicao")) = "CX" then %> selected="selected"<%end if %>>Caixa</option>
                <option value="CH" <%if trim(request("condicao")) = "CH" then %> selected="selected"<%end if %>>Chapa</option>
                <option value="PA" <%if trim(request("condicao")) = "PA" then %> selected="selected"<%end if %>>Papel</option>
              </select>
            </td>
          </tr>
<%
if nivel = "0" then
%>
          <tr>
            <th><label for="condicao">eShop:</label></th>
            <td>
              <select name="eshop" id="eshop" tabindex="2" onchange="SelecionaEShop(this.value)">
                <option value="0" <%if trim(request("eshop")) = "0" then %> selected="selected" <%end if %>>Não</option>
                <option value="1" <%if trim(request("eshop")) = "1" then %> selected="selected" <%end if %>>Sim</option>
              </select>
            </td>
          </tr>
<%
end if 
%>
          <tr>
            <th><label for="codcli">Cliente:</label></th>
            <td colspan="2">
              <input name="codcli" id="codcli" type="text"  required="required" onblur="vazio(this.value,'codcli')" onchange="vazio('','codcli')" tabindex="2" style="width:150px"/>
              &nbsp;
              <input name="lojacli" id="lojacli" placeholder="Loja" type="text" onblur="if(Trim(document.getElementById('codcli').value) != '' && Trim(this.value) != ''){ObjAjax('../../ajax/validacliente_ajax.asp?planta=00&cliente='+form.codcli.value+'&loja='+this.value+'&codsiga=<%=CODUSER%>&nivel=<%=Nivel%>&tipo=Fatura')}else{document.getElementById('codcli').value = ''; this.value = '';document.getElementById('razao').value = '' };vazio(this.value,'lojacli')" tabindex="3" />
            </td>
            <td>
              <img id="imgPesquisaClienteFatura" src="../../images/pesquisa_verde.jpg" width="20" height="20" alt="pesquisar" onclick="pesquisaCLiente()" tabindex="4" />
            </td>
            <td >
              <input type="text" name="razao" id="razao" disabled="disabled" />
            </td>
          </tr>
          <tr> 
            <th ><label for="fatura">Tipo NF:</label></th>
            <td colspan="4" > 
              <input type="radio" name="tiponf" id="fatura" value="N" <%if trim(Request("tiponf")) = "N" or trim(Request("tiponf")) = "" then %>checked="checked" <%end if %> onclick="tpNF(this.value, 'codcli2', 'lojacli2', 'razao2', 'imgPesquisaClienteRemessa', 'enable')" tabindex="5" />Normal 
              <input type="radio" name="tiponf" id="remessa" value="F"  <%if trim(Request("tiponf")) = "F" then %>checked="checked" <%end if %> onclick="tpNF(this.value, 'codcli2', 'lojacli2', 'razao2', 'imgPesquisaClienteRemessa', 'enable')" tabindex="6" />Faturamento
            </td>
          </tr>
          <tr> 
            <th><label for="codcli2">Cliente Remessa:</label></th>
            <td colspan="2">
              <input name="codcli2" id="codcli2" type="text" readonly="readonly" tabindex="7" style="width:150px"/>
              &nbsp;
              <input name="lojacli2" id="lojacli2" type="text" placeholder="Loja" onblur="clienteremessa(document.getElementById('codcli2').value,this.value,document.getElementById('codcli').value,document.getElementById('lojacli').value)" readonly="readonly" tabindex="8" />
            </td>
            <td>
              <img id="imgPesquisaClienteFatura" src="../../images/pesquisa_verde.jpg" width="20" height="20" alt="pesquisar" onclick="abrir_('Remessa','',document.getElementById('codcli').value,document.getElementById('lojacli').value,document.getElementById('codcli2').value,document.getElementById('lojacli2').value,'','',document.getElementById('enable').value,'codcli2','lojacli2','razao2','','','','','','','','','','','','codcli','','quantidade','data1','msgPeso','msgCargaMaquina','preco_unit',document.getElementById('condicao').value,document.getElementById('eshop').value)" />
            </td>
            <td >
              <input type="text" name="razao2" id="razao2" disabled="disabled" />
            </td>
          </tr> 
          <tr> 
            <th><label for="produto">Produto:</label></th>
            <td colspan="2"> 
              <input name="produto" id="produto" type="text" oninput="LimpaDadosProduto()" onblur="validaProduto(this);vazio(this.value,'produto');" tabindex="9" style="width:250px" />
            </td>
            <td>
              <img id="imgPesquisaProduto" 
                    src="../../images/pesquisa_verde.jpg" 
                    width="20" height="20" 
                    alt="pesquisar produto" 
                    onclick="abrir_('Produto','',document.getElementById('codcli').value,document.getElementById('lojacli').value,document.getElementById('codcli2').value,document.getElementById('lojacli2').value,'','',document.getElementById('enable').value,'produto','descricao','estrutura','filial','larg','comp','vinco_01','vinco_02','vinco_03','vinco_04','vinco_05',document.getElementById('codcli').value,'pallet','descpallet',document.getElementById('hidPlanta').value,'codcli','produto','quantidade','data1','msgPeso','msgCargaMaquina','preco_unit',document.getElementById('condicao').value,document.getElementById('eshop').value);" tabindex="18" />
            </td>
            <td colspan="2">
              <input name="descricao" id="descricao" type="text" disabled="disabled" />
            </td>
          </tr>
          <tr>
            <th><label for="estrutura">Estrutura:</label></th>
            <td colspan="4">                        
              <input name="estrutura" id="estrutura" type="text" readonly="readonly" tabindex="10" />
            </td>
          </tr>
          <tr>
            <th><label for="filial">Unidade Produtora:</label></th>
            <td colspan="4">
              <input name="filial" id="filial" type="text" readonly="readonly" tabindex="11" />
            </td>
          </tr>
          <tr> 
            <th><label for="descpallet">Pallet:</label></th>
            <td colspan="4"> 
              <input type="text" name="descpallet" id="descpallet" readonly="readonly" title="Campo prenchido automaticamente pelo sistema" tabindex="12" />
            </td>
          </tr>
          <tr> 
            <th><label for="larg">Largura:</label> </th>
            <td colspan="4">
              <input name="larg" id="larg" type="text"  onblur="largurach(this.value);" tabindex="13" />
            </td>
          </tr>
          <tr> 
            <th><label for="comp">Comprimento:</label></th>
            <td colspan="4">
              <input name="comp" id="comp" type="text" onblur="comprimentoch(this.value);" tabindex="14" />
            </td>
          </tr>
          <tr> 
            <th><label for="vinco_01">Vinco:</label></th>
            <td colspan="4">
              <input name="vinco_01" id="vinco_01" type="text" onblur="inteiro(this.value,'vinco_01');" placeholder="Vinco" tabindex="15" />
              <input name="vinco_02" id="vinco_02" type="text" onblur="inteiro(this.value,'vinco_02');" placeholder="Vinco" tabindex="16" />
              <input name="vinco_03" id="vinco_03" type="text" onblur="inteiro(this.value,'vinco_03');" placeholder="Vinco" tabindex="17" />
              <input name="vinco_04" id="vinco_04" type="text" onblur="inteiro(this.value,'vinco_04');" placeholder="Vinco" tabindex="18" />
              <input name="vinco_05" id="vinco_05" type="text" onblur="inteiro(this.value,'vinco_05');" placeholder="Vinco" tabindex="19" />
            </td>
          </tr>
          <tr>
            <th><label for="gramatura">Gramatura</label></th>
            <td colspan="4">
              <select name="gramatura" id="gramatura" disabled="disabled" tabindex="20">
              </select>
            </td>
          </tr>
<!--
          <tr>
            <th><label for="formato">Formato</label></th>
            <td colspan="4">
              <input type="text" name="formato" id="formato" title="Em milimteros" placeholder="milimetros" onblur="real(this.value,this)" maxlength="10" disabled="disabled" tabindex="21" />
            </td>
          </tr>
              -->
          <tr>
            <th><label for="diametro">Diâmetro</label></th>
            <td colspan="4">
              <input type="text" name="formato2" id="diametro" onblur="DiametroPapel(this.value,this)" disabled="disabled" tabindex="22" />
            </td>
          </tr>
          <tr>
            <th><label for="tubete">Tubete</label></th>
            <td>
              <select name="tubete" id="tubete" disabled="disabled" tabindex="23">
              <option value="" selected="selected"></option>
              <option value="3">3</option>
              <option value="4">4</option>
              <option value="9">9</option>
              </select>
            </td>
          </tr>
          <tr> 
            <th><label for="quantidade">Quantidade:</label></th>
            <td>
              <input name="quantidade" id="quantidade" type="number" onchange="vazio('','quantidade');inteiro(this.value,this);QuantidadeMaxima(this.value)" 
                  onblur="ValidaLotePiloto()"
                  max="99999" min="0" tabindex="24" /> 
            </td>
            <td colspan="4" id="msgPeso"></td>
          </tr>
          <tr> 
            <th title="Prazo de entrega com uma flutuação de 5 dias."><label for="data1">Data de Entrega:</label></th>
            <td title="Prazo de entrega com uma flutuação de 5 dias."  > 
              <input name="data1" id="data1" type="text" onkeypress="formatar(this, '##/##/####');" 
                      onblur="executaCargaMaquina(this)" maxlength="10" title="Prazo de entrega com uma flutuação de 5 dias." tabindex="25" />
              &nbsp;
              <img src="../../images/clean.png" alt="apagar data" title="Apagar Data" onclick="ApagaData('data1')" />
            </td>
            <td id="msgCargaMaquina" colspan="3"></td>
          </tr>
          <tr> 
            <th><label for="data2">2ª Data de Entrega:</label></th>
            <td colspan="4"  > 
              <input name="data2" id="data2"  type="text" onkeypress="formatar(this, '##/##/####');" onblur="validaData(this);" maxlength="10" disabled="disabled" tabindex="26" />
              &nbsp;
              <img src="../../images/clean.png" alt="apagar data" title="Apagar Data" onclick="ApagaData('data2')" />
            </td>
          </tr>
          <tr>
            <th><label for="tes">Tipo Venda:</label></th>
            <td colspan="2">
              <input name="descFnc" id="descFnc" type="hidden" disabled="disabled" />
              <input name="numfnc" id="numfnc" type="hidden" />
              <input type="text" name="tes" id="tes" onblur="ajaxload('../../ajax/validates_ajax.asp?eshop='+document.getElementById('eshop').value+'&',document.getElementById('codcli').value,document.getElementById('lojacli').value,document.getElementById('codcli2').value,document.getElementById('lojacli2').value,document.getElementById('produto').value,document.getElementById('tes').value,document.getElementById('numfnc').value,document.getElementById('enable').value,document.getElementById('data1').value,document.getElementById('filial').value,'tes','descTes','Tes','','','','CH')" tabindex="27" style="width:250px" />
            </td>
            <td>
              <img id="imgPesquisaTes" src="../../images/pesquisa_verde.jpg" width="20" height="20" alt="pesquisar" 
                  onclick="abrir_('TES','',document.getElementById('codcli').value,document.getElementById('lojacli').value,document.getElementById('codcli2').value,document.getElementById('lojacli2').value,document.getElementById('produto').value,document.getElementById('numfnc').value,document.getElementById('enable').value,'tes','descTes','estrutura','filial','larg','comp','vinco_01','vinco_02','vinco_03','vinco_04','vinco_05',document.getElementById('codcli').value,'pallet','descpallet','codcli','produto','quantidade','data1','msgPeso','msgCargaMaquina','preco_unit',document.getElementById('condicao').value,document.getElementById('eshop').value)" tabindex="42" />
            </td>
            <td colspan="2">
              <input name="descTes" id="descTes" type="text" disabled="disabled" />
            </td>
          </tr>
          <tr> 
            <th><label for="cond_pgto">Cond Pgto:</label></th>
            <td colspan="4"> 
              <select name="cond_pgto"  id="cond_pgto" onchange="verificaVencimento(this.value); OrcamentoPreco('<%=Planta%>',this.value,this)" tabindex="28">
                <option value="">Selecione</option>
<%
              SQL = "Execute SP_CONDPAG"
              Dados2.Source = SQL
              Dados2.Open()
              While not Dados2.EOF
%>
                <option value="<%=Trim(Dados2("E4_CODIGO"))%>"> <%=Dados2("E4_DESCRI")%> </option>
<%
                Dados2.MoveNext()
              Wend 
              Dados2.Close()
%> 
              </select>
              <input type="hidden" name="CondicaoPagamento" id="CondicaoPagamento" />&nbsp;<span id="PagamentoDescricao"></span>
            </td>
          </tr>
           <tr> 
            <th><label for="tipo_frete">Tipo Frete:</label></th>
            <td colspan="4">
              <select name="tipo_frete" id="tipo_frete" tabindex="29">
                 <option value="C" selected>CIF</option>
                 <option value="F">FOB</option>
              </select>&nbsp;
              <input name="FreteUnitario" id="FreteUnitario" type="text" value="0" tabindex="30" disabled="disabled" onblur="unitario(this.value,this)"/>
              <input type="hidden" name="tipoFrete" id="tipoFrete" value="C" />
              <span id="TipoFreteDescricao"></span>
            </td>
          </tr>
          <tr>
            <th><label for="data5">Data Vencimento:</label></th>
            <td colspan="4">
              <input name="data5" type="text" id="data5" onkeypress="formatar(this, '##/##/####');" onblur="validaData(this);" maxlength="10" tabindex="31" />
            </td>
          </tr>
          <tr> 
            <th><label for="preco_unit">Preço Unitário:</label></th>
            <td> 
              <input name="preco_unit" id="preco_unit" type="text" onclick="PrecoClick()", document.getElementById('produto').value, 'preco_unit');" onblur="unitario(this.value,this);validaPrecoKG(this.value)" tabindex="32" /> 
            </td>
            <td colspan="4" id="msgPrecoUnitario"></td>
          </tr>
          <tr>
            <th><label for="preco_unit">Se EXPORTAÇÃO, informe a moeda:</label></th>
            <td colspan="4">
              <select name="MoedaExportacao" id="MoedaExportacao" tabindex="33">
                <option value="1" selected="selected"></option>
                <option value="2">$ - Dolar</option>
                <option value="3">€ - Euro</option>
              </select>
            </td>
          </tr>
          <tr> 
            <th><label for="desconto">Desconto:</label></th>
            <td colspan="4"> 
              <input name="desconto" id="desconto" type="text" value="0" onblur="real(this.value,this);form.obs_venda.focus();" tabindex="34" />
            </td>
          </tr>
         
          <tr>
            <th><label for="obs_venda">OBS:</label></th>
            <td colspan="4"> 
              <input name="obs_venda" id="obs_venda" type="text" maxlength="150" tabindex="35"/>
            </td>
          </tr>
          <tr> 
            <th><label for="nf_pallet">NF Pallet:</label></th>
            <td colspan="4">
              <select name="nf_pallet" id="nf_pallet" tabindex="36">
                <option value="N" selected="selected">Não</option>
                <option value="S">Sim</option>
              </select> 
            </td>
          </tr>
          <tr> 
            <th><label for="ped_cliente">Pedido Cliente:</label></th>
            <td colspan="4"> 
              <input name="ped_cliente" id="ped_cliente" type="text" maxlength="15" tabindex="37"  />
              &nbsp;
              <input name="itempedcli" id="itempedcli" type="number" min="1" max="999999" tabindex="38" onblur="verificaPedidoCompra('<%=Planta%>')"  /> 
            </td>
          </tr>
          <tr>
            <th>
              <label for="dataDescarga">Descarga: Data, Hora e senha</label>
            </th>
            <td colspan="4">                        
              <input name="dataDescarga" id="dataDescarga"  type="text" onkeypress="formatar(this, '##/##/####');" onblur="validaData(this);checarDatas(this,document.getElementById('data1'))" maxlength="10" tabindex="39" />
              &nbsp;
              <img src="../../images/clean.png" alt="apagar data" title="Apagar Data" onclick="ApagaData('dataDescarga')" tabindex="40" />
              &nbsp;
              <input type="text" value="" id="horaDescarga" name="horaDescarga" onkeypress="formatar(this, '##:##');" maxlength="5" onblur="validaHora(this)" tabindex="41" />
              &nbsp;
              <img src="../../images/clean.png" alt="apagar hora" title="Apagar Hora" onclick="ApagaData('horaDescarga')" tabindex="42" />
              &nbsp;
              <input type="text" value="" id="descarregamento" name="descarregamento" tabindex="43" />
            </td>
          </tr>
          <tr>
            <th title="Obs. para incluir na NF - Max. de 130 caracteres">
              <label for="MsgNota">1ª Obs para NF</label>
            </th>
            <td colspan="4" title="Obs. para incluir na NF - Max. de 60 caracteres">
              <input type="text" name="MsgNota" id="MsgNota" maxlength="60" class="Observacao" tabindex="44"/>
            </td>
          </tr>
          <tr>
            <th title="Obs. para incluir na NF - Max. de 130 caracteres">
              <label for="MsgNota">2ª Obs para NF</label>
            </th>
            <td colspan="4" title="Obs. para incluir na NF - Max. de 130 caracteres">
              <input type="text" name="MsgNota2" id="MsgNota2" maxlength="130" class="Observacao" disabled="disabled" tabindex="45"/>
            </td>
          </tr>
          <tr>
            <th title="Obs. para incluir na NF - Max. de 130 caracteres">
              <label for="MsgNota">3ª Obs para NF</label>
            </th>
            <td colspan="4" title="Obs. para incluir na NF - Max. de 130 caracteres">
              <input type="text" name="MsgNota3" id="MsgNota3" maxlength="130" class="Observacao" disabled="disabled" tabindex="46"/>
            </td>
          </tr>
          <tr>
            <th title="Obs. para incluir na NF">
              <label for="MsgNota">4ª Obs para NF</label>
            </th>
            <td colspan="4" title="Obs. para incluir na NF">
              <input type="text" name="MsgNota4" id="MsgNota4" maxlength="130" class="Observacao" disabled="disabled" tabindex="47"/>
            </td>
          </tr>
          <tr>
            <th class="pading-botton" colspan="5">
              <input type="checkbox" name="LotePiloto" id="LotePiloto"  onclick="return false;" tabindex="48" style="pointer-events: none;"/>
              &nbsp;
              <label for="LotePiloto">Marque se for Lote Piloto.</label>
            </th>
          </tr>
          <tr>
            <th class="pading-botton" colspan="5">
              <input type="checkbox" name="checkRefaturamento" id="checkRefaturamento" tabindex="48" />
              &nbsp;
              <label for="checkRefaturamento">Marque se for Refaturamento.</label>
            </th>
          </tr>
         <tr>
                <td colspan="5">
                    <div id="msgLotePiloto" class="msg-lote-piloto" style="display:none;"></div>
                </td>
          </tr>
          <tr>
            <td class="pading-botton" colspan="5"></td>
          </tr>
          <tr>
            <td class="pading-top" colspan="5"></td>
          </tr>
          <tr> 
            <td colspan="5" class="input_text_align_right">
              <input name="btnGrava" id="btnGrava" type="button" onclick="verify();" value="Gravar Pedido" class="button" tabindex="49" style="width:150px"/>
              &nbsp;
              <input name="limpar" id="limpar" type="reset" value="Limpar Campos" class="button" tabindex="48" />
            </td>
          </tr>
        </table>
      </div>
    </form>

    <script type="text/javascript">
        /* VALIDANDO O FORMUL�RIO */
        function verify() {
            var parametroPreco = 0
            var parametroPeso = 0
            var themessage = "Campo de preenchimento obrigatório:\n........................................................";

            //Verificando se o tipo de tes � 580 ou 581 para obrigar a preencher
            //o n�mero da fnc no campo observa��es
            var prod = document.getElementById('produto').value
            prod = prod.substring(0, 3)
            prod2 = prod.substring(0, 2)

            if (Trim(document.getElementById('tes').value) == "") {
                themessage = themessage + "\nInforme a TES. ";
            }

            /*Validando c�digo do cliente em branco*/
            if (document.getElementById('codcli').value == "") {
                themessage = themessage + "\nCódigo do Cliente. ";
            }

            if (document.getElementById('ped_cliente').value != "") {
                if (document.getElementById('itempedcli').value == "" || isNaN(document.getElementById('itempedcli').value)) {
                    themessage = themessage + "\nItem pedido do Cliente é NUMERICO e obrigatório. ";
                    document.form.lojacli.focus()
                }
            }

            /*Validando c�digo da loja em branco*/
            if (document.getElementById('lojacli').value == "") {
                themessage = themessage + "\nCódigo da Loja do Cliente. ";
                document.getElementById('lojacli').focus()
            }

            /*Validando c�digo do cliente de destino em branco*/
            if (document.getElementById('enable').value == "F" && document.getElementById('codcli2').value == '') {
                themessage = themessage + "\nCódigo do cliente para remessa. ";
            }

            /*Validando c�digo da loja de destino em branco*/
            if (document.getElementById('enable').value == "F" && document.getElementById('lojacli2').value == '') {
                themessage = themessage + "\nCódigo da loja do cliente para remessa. ";
            }

            /*Validando c�digo do produto em branco*/
            if (document.getElementById('produto').value == "") {
                themessage = themessage + "\nProduto. ";
            }

            /*Validando quantidade em branco ou zerada*/
            if (document.getElementById('quantidade').value == "" || document.getElementById('quantidade').value == "0") {
                themessage = themessage + "\nQuantidade. ";
            }

            /*Validando pre�o unit�rio em branco*/
            var valor = document.getElementById('preco_unit');
            var val = valor.value.replace(",", '.')
            val = parseFloat(val)
            if (document.getElementById('preco_unit').value == "") {
                themessage = themessage + "\nPreço unitário. ";
            } else {
                if (!isNaN(val)) {
                    if (val <= 0) {
                        themessage = themessage + "\nPreço unitário, menor ou igual a zero. ";
                        document.getElementById('preco_unit').value = "0";
                    }
                } else {
                    if (document.getElementById('preco_unit').value != "") {
                        themessage = themessage + "\nPreço unitário é exclusivamente numerico.";
                        document.getElementById('preco_unit').value = "0"
                    } else {
                        themessage = themessage + "\nPreço unitário não pode ser vazio.";
                        document.getElementById('preco_unit').value = "0"
                    }
                }
            }

            /*Validando data de entrega em branco*/
            if (document.getElementById('data1').value == "") {
                themessage = themessage + "\nPrimeira data de entrega. ";
            }

            if (document.getElementById('condicao').value == "PA" && document.getElementById('data2').value == "") {
                themessage = themessage + "\nSegunda data de entrega. ";
            }
        
            /*if (!ValidaDataEmissao()){
              themessage = themessage + "\nData de Entrega Indisponivel. Data mínima do Cliente: " + dataMinima. ";
            }*/

            dias = document.getElementById("hidDiasTransporte").value;
            if (dias == '')
               dias = 0

            console.log('dias', dias)
            console.log('hidFlagDiasViagem',document.getElementById("hidFlagDiasViagem").value)
            
            dataMinima = new Date().toLocaleDateString('pt-BR', {
                    day: '2-digit',
                    month: '2-digit',
                    year: 'numeric'
            });
            console.log('dataMinima',dataMinima)
            if (document.getElementById("hidFlagDiasViagem").value == "1") {
               dataMinima = AdicionarDias(dataMinima, parseInt(dias));
               console.log('regra document.getElementById("hidFlagDiasViagem").value', dataMinima);
            }
                        
            var dataEmissaoValida = VerificaDataEmissao(document.getElementById("data1").value,dataMinima)
            console.log('dataEmissaoValida',dataEmissaoValida)
            if (dataEmissaoValida != true){
               themessage = themessage + `\Data mínima do Cliente: ${dataMinima}`;
            }

            var diasAcrescimo = 0
                if (document.getElementById("condicao").value == "CX") {
                    parametroM2 = TrataValores(document.getElementById("hidMtsCXmin").value)
                    var diasAcrescimo = document.getElementById("hidDiasAcrescimoCX").value
                    console.log('diasAcrescimoCX', diasAcrescimo)
                } else if (document.getElementById("condicao").value == "CH") {
                    parametroM2 = TrataValores(document.getElementById("hidMtsCHmin").value)
                    var diasAcrescimo = document.getElementById("hidDiasAcrescimoCH").value
                    console.log('diasAcrescimoCH', diasAcrescimo)
                } else if (document.getElementById("condicao").value == "PA") {
                    var diasAcrescimo = document.getElementById("hidDiasAcrescimoPA").value
                    console.log('hidDiasAcrescimoPA', diasAcrescimo)
                }

                if (parseInt(diasAcrescimo) > 0) {
                    dataMinima = AdicionarDias(dataMinima, parseInt(diasAcrescimo));
                    console.log('dataMinima adicionado de dias de acrescimo por tipo de Produto: ', dataMinima)
                }

                if (document.getElementById("data1").value !== "") {
                    var dataEmissaoValida = VerificaDataEmissao(document.getElementById("data1").value,dataMinima)
                    console.log('dataEmissaoValida',dataEmissaoValida)
                    if (dataEmissaoValida != true){
                        themessage = themessage + `\Data mínima do Cliente: ${dataMinima}`;
                    }
                }
            
        
            if (document.getElementById("condicao").value == "CH") {
                parametroPreco = TrataValores(document.getElementById("hidPrecoKgChapa").value)
                parametroPeso = TrataValores(document.getElementById("hidPesoCH").value)
                
                
                /*Validando largura da chapa em branco*/
                var larguraMin = parseInt(document.getElementById("hidLarguraMin").value);
                var larguraMax = parseInt(document.getElementById("hidLarguraMax").value);
                console.log('larguraMax', larguraMax)
                if (document.getElementById('larg').value == ""){
                    themessage = themessage + "\nLargura da Chapa. ";
                }
                
                if (document.getElementById('larg').value < larguraMin){
                    themessage = themessage + `\nLargura mínima ${larguraMin} mm`;
                }
                if (document.getElementById('larg').value > larguraMax) {
                    themessage = themessage + `\nLargura maxima ${larguraMax} mm`;
                }
              
                /*Validando comprimento Maximo e Minimo da chapa*/
                var comprimentoMin = parseInt(document.getElementById("hidComprimentoMin").value);
                var comprimentoMax = parseInt(document.getElementById("hidComprimentoMax").value);
                if (document.getElementById('comp').value == ""){
                    themessage = themessage + "\nComprimento da Chapa.";
                }
                
                if ((document.getElementById('comp').value < comprimentoMin) && (document.getElementById('condicao').value == "CH")) {
                    themessage = themessage + `\nComprimento mínimo ${comprimentoMin} mm.`;
                }
                if ((document.getElementById('comp').value > comprimentoMax) && (document.getElementById('condicao').value == "CH")) {
                    themessage = themessage + `\nComprimento maximo ${comprimentoMax} mm.`;
                }
            }   
            
            if (document.getElementById("condicao").value != "PA") {
              quantidade = TrataValores(document.getElementById("quantidade").value)
              console.log('LarguraProduto', document.getElementById("larg").value)
              console.log('ComprimentoProduto', document.getElementById("comp").value)
              TotalM2 = parseFloat(document.getElementById("larg").value) / 1000 * parseFloat(document.getElementById("comp").value) / 1000
              TotalM2 = TotalM2 * quantidade
              console.log('TotalM2', TotalM2)
              console.log('parametroM2', parametroM2)
              if (!ConsisteM2(TotalM2, parametroM2))
                   themessage = themessage + `M2 Total menor que o metragem mínima de ${parametroM2} M2`;
            }

            // Verificando a Condição de pagamento
            if (document.getElementById("CondicaoPagamento").value == "") {
                themessage = themessage + "\nInforme a Condição de Pagamento";
            }

            /*Validando vincos*/
            if (Trim(document.getElementById('vinco_01').value) == "") {
                vinco1 = 0
            } else {
                vinco1 = parseInt(document.getElementById('vinco_01').value)
            }

            if (Trim(document.getElementById('vinco_02').value) == "") {
                vinco2 = 0
            } else {
                vinco2 = parseInt(document.getElementById('vinco_02').value)
            }

            if (Trim(document.getElementById('vinco_03').value) == "") {
                vinco3 = 0
            } else {
                vinco3 = parseInt(document.getElementById('vinco_03').value)
            }

            if (Trim(document.getElementById('vinco_04').value) == "") {
                vinco4 = 0
            } else {
                vinco4 = parseInt(document.getElementById('vinco_04').value)
            }

            if (Trim(document.getElementById('vinco_05').value) == "") {
                vinco5 = 0
            } else {
                vinco5 = parseInt(document.getElementById('vinco_05').value)
            }

            vinco = vinco1 + vinco2 + vinco3 + vinco4 + vinco5

            if ((vinco > 0 && vinco != document.getElementById('larg').value) && (document.getElementById('condicao').value == "CH")) {
                themessage = themessage + "\nA soma dos vincos tem que ser igual a largura da chapa ";
                document.getElementById('vinco_01').focus();
            }

            //if (((Trim(document.getElementById('formato').value) == "0") || (Trim(document.getElementById('formato').value) == "")) && (document.getElementById('condicao').value == "PA")) {
            //    themessage = themessage + "\nFavor informar o Formato do Papel. ";
            //}

            if (((Trim(document.getElementById('diametro').value) == "0") || (Trim(document.getElementById('diametro').value) == "")) && (document.getElementById('condicao').value == "PA")) {
                themessage = themessage + "\nFavor informar o Diâmetro do Papel. ";
            }

            /*Validando c�digo cliente e da loja de destino e c�digo cliente e loja iguais*/
            if (document.getElementById('codcli').value == document.getElementById('codcli2').value && document.getElementById('lojacli').value == document.getElementById('lojacli2').value) {
                themessage = themessage + "\nNão é permitido emitir uma nota de faturamento para o mesmo cliente e loja.\nSelecione Normal. ";
            }

            if ((document.getElementById('eshop').value == "1") && ((document.getElementById('condicao').value != "CX") && (document.getElementById('condicao').value != "CH"))) {
                themessage = themessage + "\nNão é permitido emitir pedidos eShop que não sejam para caixa ou chapa.\nSelecione Normal. ";
            }
            console.log('hidCargaMaquinaValida',document.getElementById('hidCargaMaquinaValida').value)
            if(document.getElementById('hidCargaMaquinaValida').value=="0" && document.getElementById("hidCargaMaquina").value == "1"){
              themessage = themessage + "\nCarga maquina fechada.";
            } 
            //Carga maquina fechada

            /*Chama o alert se achar algum campo obrigatorio vazio ou fora dos padr�es*/
             if (themessage == "Campo de preenchimento obrigatório:\n........................................................") {
                  //ValidaLotePiloto()
                   VerificaOrcamento(document.getElementById("filial").value
                                , document.getElementById("codcli").value
                                , document.getElementById("lojacli").value
                                , document.getElementById("produto").value
                                , Trim(document.getElementById("tipo_frete").value)
                                , Trim(document.getElementById("cond_pgto").value)
                                , Trim(document.getElementById("preco_unit").value));
                  
             } else {
                 alert(themessage);
                 return false;
             }
        }
        // ----------------------------------------------------------------------------------------
        function objReturn(obj) {
            if (obj.erro == 'ErroLogof') {
                alert('Favor se conectar novamente ao SKNET.\nSeu logon EXPIROU.');
                if (obj.tipo == 'Fatura') {
                    document.getElementById('codcli').value = ''
                    document.getElementById('lojacli').value = ''
                    document.getElementById('razao').value = ''
                }/*FIM VALIDA��O RETORNO STRING ErroLogof CASO TIPO = Fatura*/

                if (obj.tipo == 'Remessa') {
                    document.getElementById('codcli2').value = ''
                    document.getElementById('lojacli2').value = ''
                    document.getElementById('razao2').value = ''
                }/*FIM VALIDA��O RETORNO STRING ErroLogof CASO TIPO = Remessa*/

                if (obj.tipo == 'Produto') {
                    document.getElementById('produto').value = ''
                    document.getElementById('descricao').value = ''
                } /*FIM VALIDA��O RETORNO STRING ErroLogof CASO TIPO = Produto*/

                if (obj.tipo == 'Tes') {
                    document.getElementById('Tes').value = ''
                }/*FIM VALIDA��O RETORNO STRING ErroLogof CASO TIPO = Tes*/

                if (obj.tipo == 'Fnc') {
                    document.getElementById('numfnc').value = ''
                }/*FIM VALIDA��O RETORNO STRING ErroLogof CASO TIPO = Fnc*/

                if (obj.tipo == 'Maq') {
                    document.getElementById('msgCargaMaquina').innerHTML = '';
                    document.getElementById('data1').value = ''
                    document.getElementById('msgCargaMaquina').style.borderRadius = "none"
                    document.getElementById('msgCargaMaquina').style.verticalAlign = "none"
                    document.getElementById('msgCargaMaquina').style.textAlign = "none"
                    document.getElementById('msgCargaMaquina').style.color = "none";
                    document.getElementById('msgCargaMaquina').style.border = "none"
                    document.getElementById('msgCargaMaquina').style.background = "none"
                    document.getElementById('msgCargaMaquina').innerHTML = '';
                }/*FIM VALIDA��O RETORNO STRING ErroLogof CASO TIPO = Maq*/

                if (obj.tipo == 'Pes') {
                    document.getElementById('msgPeso').innerHTML = '';
                    document.getElementById('quantidade').value = 0
                    document.getElementById('msgPeso').style.borderRadius = "none"
                    document.getElementById('msgPeso').style.verticalAlign = "none"
                    document.getElementById('msgPeso').style.textAlign = "none"
                    document.getElementById('msgPeso').style.color = "none";
                    document.getElementById('msgPeso').style.border = "none"
                    document.getElementById('msgPeso').style.background = "none"
                    document.getElementById('msgPeso').innerHTML = '';
                }/*FIM VALIDA��O RETORNO STRING ErroLogof CASO TIPO = Maq*/

                if (obj.tipo == 'Prc') {
                    document.getElementById('').value = 0
                }
                return false;

            }/*FIM VALIDA��O RETORNO STRING ErroLogof*/
            if (obj.erro == 'ErroNEncontrado') {
                if (obj.tipo == 'Fatura') {
                    document.getElementById('codcli').value = ''
                    document.getElementById('lojacli').value = ''
                    document.getElementById('razao').value = ''
                    alert('CLIENTE NÃO ENCONTRADO.\n\nPossíveis causas:\nCliente Bloqueado.\nCliente não pertence a sua carteira.\nCliente não Cadastrado.');
                }/*FIM VALIDA��O RETORNO STRING ErroNEncontrado CASO TIPO = FATURA*/

                if (obj.tipo == 'Remessa') {
                    document.getElementById('codcli2').value = ''
                    document.getElementById('lojacli2').value = ''
                    document.getElementById('razao2').value = ''
                    alert('CLIENTE REMESSA NÃO ENCONTRADO.\n\nPossíveis causas:\nCliente não cadastrado para envio de remessa.\nFaça a solicitação do cadastro atraves do menu "Solicita Cliente Remessa" e aguarde aprovação.');
                }/*FIM VALIDA��O RETORNO STRING ErroNEncontrado CASO TIPO = REMESSA*/

                if (obj.tipo == 'Produto') {
                    document.getElementById('produto').value = ''

                    /*CAMPOS DESABILITADOS*/
                    document.getElementById('descricao').value = ''
                    document.getElementById('estrutura').value = ''
                    document.getElementById('filial').value = ''

                    /*CAMPOS READONLY*/
                    document.getElementById('larg').value = ''
                    document.getElementById('comp').value = ''
                    document.getElementById('vinco_01').value = ''
                    document.getElementById('vinco_02').value = ''
                    document.getElementById('vinco_03').value = ''
                    document.getElementById('vinco_04').value = ''
                    document.getElementById('vinco_05').value = ''
                    document.getElementById('pallet').value = ''
                    document.getElementById('descpallet').value = ''

                    /*CAMPOS READONLY*/
                    document.getElementById('larg').readOnly = true;
                    document.getElementById('comp').readOnly = true;
                    document.getElementById('vinco_01').readOnly = true;
                    document.getElementById('vinco_02').readOnly = true;
                    document.getElementById('vinco_03').readOnly = true;
                    document.getElementById('vinco_04').readOnly = true;
                    document.getElementById('vinco_05').readOnly = true;
                    alert('PRODUTO NÃO ENCONTRADO.\n\nPossíveis causas:\nProduto Bloqueado.\nProduto não pertence a sua carteira.\nProduto não cadastrado.');
                }/*FIM VALIDA��O RETORNO STRING ErroNEncontrado CASO TIPO = PRODUTO*/

                if (obj.tipo == 'Pro') {
                    document.getElementById('preco_unit').readOnly = false;
                    //if (document.getElementById('eshop').value == '0') {
                    document.getElementById('cond_pgto').disabled = false;
                    document.getElementById('PagamentoDescricao').style.borderRadius = "8px"
                    document.getElementById('PagamentoDescricao').style.verticalAlign = "middle"
                    document.getElementById('PagamentoDescricao').style.textAlign = "center"
                    document.getElementById('PagamentoDescricao').innerHTML = "";
                    //}
                    document.getElementById('tipo_frete').disabled = false;
                    document.getElementById('data5').disabled = false;

                    document.getElementById('TipoFreteDescricao').style.borderRadius = "8px"
                    document.getElementById('TipoFreteDescricao').style.verticalAlign = "middle"
                    document.getElementById('TipoFreteDescricao').style.textAlign = "center"
                    document.getElementById('TipoFreteDescricao').innerHTML = "";
                    alert('Não foi encontrada uma proposta para este pedido.\nOs campos Condição de Pagamento, Preço Unitário e Tipo de Frete estão sendo liberados para preenchimeto.');
                }/*FIM VALIDA��O RETORNO STRING ErroNEncontrado CASO TIPO = PRODUTO*/

                if (obj.tipo == 'Tes') {
                    document.getElementById('tes').value = ''
                    document.getElementById('descTes').value = ''
                    alert('TES NÃO ENCONTRADA.');
                }/*FIM VALIDA��O RETORNO STRING ErroNEncontrado CASO TIPO = TES*/

                if (obj.tipo == 'Fnc') {
                    document.getElementById('numfnc').value = ''
                    document.getElementById('descFnc').value = ''
                    alert('FNC NÃO ENCONTRADA.');
                }/*FIM VALIDA��O RETORNO STRING ErroNEncontrado CASO TIPO = Fnc*/

                if (obj.tipo == 'Maq') {
                    document.getElementById('msgCargaMaquina').innerHTML = '';
                    document.getElementById('data1').value = ''
                    document.getElementById('msgCargaMaquina').style.borderRadius = "none"
                    document.getElementById('msgCargaMaquina').style.verticalAlign = "none"
                    document.getElementById('msgCargaMaquina').style.textAlign = "none"
                    document.getElementById('msgCargaMaquina').style.color = "none";
                    document.getElementById('msgCargaMaquina').style.border = "none"
                    document.getElementById('msgCargaMaquina').style.background = "none"
                }/*FIM VALIDA��O RETORNO STRING ErroNEncontrado CASO TIPO = Maq*/

                if (obj.tipo == 'Pes') {
                    document.getElementById('msgPeso').innerHTML = '';
                    document.getElementById('quantidade').value = 0
                    document.getElementById('msgPeso').style.borderRadius = "none"
                    document.getElementById('msgPeso').style.verticalAlign = "none"
                    document.getElementById('msgPeso').style.textAlign = "none"
                    document.getElementById('msgPeso').style.color = "none";
                    document.getElementById('msgPeso').style.border = "none"
                    document.getElementById('msgPeso').style.background = "none"
                }/*FIM VALIDA��O RETORNO STRING ErroNEncontrado CASO TIPO = Pes*/

                if (obj.tipo == 'Prc') {
                    document.getElementById('preco_unit').value = 0
                }

                return false;
            }/*FIM RETORNO STRING ErroNEncontrado*/

            if (obj.erro == 'ErroCliente') {
                alert('Favor preencher os campos Cliente e Loja.');
                if (obj.tipo == 'Fatura') {
                    document.getElementById('codcli').value = ''
                    document.getElementById('lojacli').value = ''
                    document.getElementById('razao').value = ''
                }
                if (obj.tipo == 'Remessa') {
                    document.getElementById('codcli2').value = ''
                    document.getElementById('lojacli2').value = ''
                    document.getElementById('razao2').value = ''
                }
                return false;
            }/*FIM RETORNO STRING - ErroCliente*/
            if (obj.erro == 'ErroPrc') {
                document.getElementById('preco_unit').value = 0
                return false;
            }

            if (obj.erro == 'ErroParametro') {
                if (obj.tipo == 'Pes') {
                    alert('Faltam parametros para calcular o peso.')
                    document.getElementById('msgPeso').style.borderRadius = "none"
                    document.getElementById('msgPeso').style.verticalAlign = "none"
                    document.getElementById('msgPeso').style.textAlign = "none"
                    document.getElementById('msgPeso').style.color = "none";
                    document.getElementById('msgPeso').style.border = "none"
                    document.getElementById('msgPeso').style.background = "none"
                    document.getElementById('msgPeso').innerHTML = '';
                    document.getElementById('quantidade').value = 0;
                }/*FIM PESO*/
            }/*FIM ERRO PARAMETRO*/

            if (obj.tipo == 'Fatura') {
                document.getElementById('razao').value = obj.val1;
            }/*FIM REMESSA*/

            if (obj.tipo == 'Remessa') {
                document.getElementById('razao2').value = obj.val1;
            }/*FIM REMESSA*/

            if (obj.tipo == 'Produto') {
                document.getElementById('produto').value = Trim(document.getElementById('produto').value);

                /*CAMPOS DISABLED*/
                document.getElementById('descricao').value = obj.val1
                document.getElementById('estrutura').value = obj.val2
                document.getElementById('filial').value = obj.val3

                /*CAMPO HIDDEN*/
                document.getElementById('pallet').value = obj.val11

                /*CAMPOS DISABLED*/
                document.getElementById('descpallet').value = obj.val12

                /*SETANDO OS CAMPOS COMO READONLY*/
                document.getElementById('larg').readOnly = true;
                document.getElementById('comp').readOnly = true;
                document.getElementById('vinco_01').readOnly = true;
                document.getElementById('vinco_02').readOnly = true;
                document.getElementById('vinco_03').readOnly = true;
                document.getElementById('vinco_04').readOnly = true;
                document.getElementById('vinco_05').readOnly = true;


                /*VERIFICANDO SE A LARGURA EST� ZERADO PARA SETAR O CAMPO PARA ALTERA��O*/
                if ((Trim(obj.val4) == '0') && (obj.val13 != "CX") && (obj.val13 != "PA")) {
                    document.getElementById('larg').readOnly = false;
                    document.getElementById('comp').readOnly = false;
                    document.getElementById('vinco_01').readOnly = false;
                    document.getElementById('vinco_02').readOnly = false;
                    document.getElementById('vinco_03').readOnly = false;
                    document.getElementById('vinco_04').readOnly = false;
                    document.getElementById('vinco_05').readOnly = false;
                } else {
                    if (Trim(obj.val4) != '0') {
                        document.getElementById('larg').value = obj.val4
                        document.getElementById('comp').value = obj.val5
                        document.getElementById('vinco_01').value = obj.val6
                        document.getElementById('vinco_02').value = obj.val7
                        document.getElementById('vinco_03').value = obj.val8
                        document.getElementById('vinco_04').value = obj.val9
                        document.getElementById('vinco_05').value = obj.val10
                    }
                }

                //BUSCANDO O PRE�O DO PRODUTO
                if (obj.val13 == 'CX') {
                    ajaxPreco(document.getElementById('codcli').value, document.getElementById('produto').value, 'preco_unit');
                } else {
                    document.getElementById('preco_unit').value = 0
                }

            }/*FIM PRODUTO*/

            if (obj.tipo == 'Tes') {
                if ((obj.val1 != 'True') && (obj.val1 != 'Verdadeiro')) {
                    document.getElementById('tes').value = '';
                    document.getElementById('descTes').value = '';
                    alert('Verifique se a TES ja esta no cadastro do Cliente.\nDesde o dia 01/12/2016 esta validação esta sendo\nefeitivada apenas pelo cadastro do Cliente!');
                } else {
                    document.getElementById('descTes').value = obj.val3;
                }
            }/*FIM TES*/

            if (obj.tipo == 'Fnc') {
                document.getElementById('numfnc').value = obj.val1;
                document.getElementById('descFnc').value = obj.val2;
            } /*FIM FNC*/
debugger;
console.log('msg do carga maquina obj.erro',Trim(obj.erro))
            if (obj.tipo == 'Maq') {
                if (Trim(obj.val1) == '1') {
                    document.getElementById('msgCargaMaquina').style.color = "#6BA539";
                    document.getElementById('msgCargaMaquina').style.border = "1px dashed #6BA539"
                    document.getElementById('msgCargaMaquina').style.background = "#d2f8a1"
                    document.getElementById('hidCargaMaquinaValida').value = "1"
                } else {
                    document.getElementById('msgCargaMaquina').style.color = "red";
                    document.getElementById('msgCargaMaquina').style.border = "1px dashed red"
                    document.getElementById('msgCargaMaquina').style.background = "#fea6a0"
                    document.getElementById('hidCargaMaquinaValida').value = "0"
                }
                document.getElementById('msgCargaMaquina').style.borderRadius = "8px"
                document.getElementById('msgCargaMaquina').style.verticalAlign = "middle"
                document.getElementById('msgCargaMaquina').style.textAlign = "center"
                document.getElementById('msgCargaMaquina').innerHTML = '<b>' + obj.val2 + '</b>';
                Proposta();
            } /*FIM VALIDA��O RETORNO STRING ErroNEncontrado CASO TIPO = Maq*/

            if (obj.tipo == 'Pes') {
                var msgPesoEl = document.getElementById('msgPeso');
                var isOk = (parseInt(Trim(obj.val1)) == 1);

                // Limpa estilos inline anteriores (agora controlados por CSS de classe)
                msgPesoEl.removeAttribute('style');

                if (isOk) {
                    document.getElementById('btnGrava').removeAttribute('disabled','');
                } else {
                    if (document.getElementById("LotePiloto").checked){
                        document.getElementById('btnGrava').removeAttribute('disabled','');
                        const msgLote = document.getElementById("msgLotePiloto");
                        msgLote.style.display = "block";
                        msgLote.innerHTML = "PEDIDO COM LOTE PILOTO, PODE SER EMITIDO COM PESO E ÁREA ABAIXO DOS PARÂMETROS";
                    } else {
                        document.getElementById('btnGrava').setAttribute('disabled','disabled');
                    }
                }

                // Tenta extrair Peso, Área e Paletes do texto retornado pelo servidor
                var rawText = obj.val2 || '';
                var mPeso  = rawText.match(/Peso\s+Calc[^:]*:\s*([\d.,]+(?:\s*[A-Za-zÀ-ÿ.]+)?)/i);
                var mArea  = rawText.match(/[ÁA]rea\s+Calc[^:]*:\s*([\d.,]+(?:\s*[A-Za-zÀ-ÿ.]+)?)/i);
                var mPalet = rawText.match(/Qtd[^:]*Palet[^:]*:\s*(\d+)/i);

                if (mPeso && mArea && mPalet) {
                    var pesoFormatado = Trim(mPeso[1]).replace(/\s*Kg\.?/gi, ' KG.');
                    var areaFormatada = Trim(mArea[1]).replace(/\s*Mts\.?/gi, ' M²');

                    // Renderiza em células brancas, com rótulo preto e valor verde.
                    msgPesoEl.innerHTML =
                        '<div class="info-cells-wrapper ' + (isOk ? 'ok' : 'err') + '">' +
                            '<div class="info-cell"><span class="cell-label">Peso Calc.</span><span class="cell-value">' + pesoFormatado + '</span></div>' +
                            '<div class="info-cell"><span class="cell-label">Área Calc.</span><span class="cell-value">' + areaFormatada + '</span></div>' +
                            '<div class="info-cell"><span class="cell-label">Qtd. Paletes</span><span class="cell-value">' + Trim(mPalet[1]) + '</span></div>' +
                        '</div>';
                } else {
                    // Fallback sem borda pontilhada e sem fundo colorido.
                    rawText = rawText
                        .replace(/\s*Kg\.?/gi, ' KG.')
                        .replace(/\s*Mts\.?/gi, ' M²');

                    msgPesoEl.innerHTML =
                        '<div class="info-cells-wrapper ' + (isOk ? 'ok' : 'err') + '">' +
                            '<div class="info-cell info-cell-full"><span class="cell-value">' + rawText + '</span></div>' +
                        '</div>';
                }
            }/*FIM VALIDA��O RETORNO STRING ErroNEncontrado CASO TIPO = Pes*/

            if (obj.tipo == 'Prc') {
                document.getElementById('preco_unit').value = obj.val2
            }

            if (obj.tipo == 'Pro') {

              document.getElementById('CondicaoPagamento').value = obj.val1;
              document.getElementById('tipoFrete').value = obj.val2;
              document.getElementById('preco_unit').value = obj.val3;
              
              unitario(document.getElementById('preco_unit').value, document.getElementById('preco_unit'))

              document.form.PagamentoDescricao.style.borderRadius = "8px"
              document.form.PagamentoDescricao.style.verticalAlign = "middle"
              document.form.PagamentoDescricao.style.textAlign = "center"
              document.form.PagamentoDescricao.innerHTML = '<b>' + obj.val4 + '</b>';

              document.form.TipoFreteDescricao.style.borderRadius = "8px"
              document.form.TipoFreteDescricao.style.verticalAlign = "middle"
              document.form.TipoFreteDescricao.style.textAlign = "center"
              document.form.TipoFreteDescricao.innerHTML = '<b>' + obj.val5 + '</b>';

            }

            if (obj.tipo == 'Usr') {

              if (obj.erro != '') {
                document.form.msgUsuario.style.color = "#006298";
                document.form.msgUsuario.style.border = "1px dashed #006298";
                document.form.msgUsuario.style.background = "#CBD9E4";
                document.form.msgUsuario.style.borderRadius = "8px";
                document.form.msgUsuario.style.verticalAlign = "middle";
                document.form.msgUsuario.style.textAlign = "center";
                document.form.msgUsuario.style.paddingBottom = "30px";
                document.form.codusu.value = '<%=CODUSER%>';
                document.form.msgUsuario.innerHTML = '<b><%=NOME%></b>';
              } else {
                if (document.form.msgUsuario) {
                  document.form.msgUsuario.style.color = "#006298";
                  document.form.msgUsuario.style.border = "1px solid #006298";
                  document.form.msgUsuario.style.background = "#CBD9E4";
                  document.form.msgUsuario.style.borderRadius = "8px";
                  document.form.msgUsuario.style.verticalAlign = "middle";
                  document.form.msgUsuario.style.textAlign = "center";
                  document.form.msgUsuario.innerHTML = '<b>' + obj.val2 + ' ' + obj.val1 + '</b>';
                }
              }

            }
        }    
        // ----------------------------------------------------------------------------------------
      function pesquisaCLiente() {

        // Obtém a altura da tela
        const screenHeight = window.screen.height;
        // Define a altura máxima desejada (por exemplo, 800 pixels)
        const maxHeight = 800;
        // Calcula a altura a ser usada
        const height = screenHeight; //Math.min(screenHeight, maxHeight);
        // Abre uma nova janela com a altura calculada
        window.open ('../../buscaCLIENTE.asp?id1=codcli&id2=lojacli&id3=razao&id6=00&id8=hidDiasTransporte','Consulta Cliente',`left,width=900,height=${height}`);

      }

    </script>

  </body>
</html>