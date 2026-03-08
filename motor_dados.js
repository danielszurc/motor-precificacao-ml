/**
 * MÓDULO 1: CAMADA DE DADOS E ESTRUTURAÇÃO DO BLOCO VIRTUAL
 * Responsabilidade: Ler o Google Sheets de forma otimizada (In-Memory)
 * e preparar a estrutura de dados (O Bloco Virtual) para o motor matemático.
 */

// 1. CARREGAMENTO DO BANCO DE DADOS PARA A MEMÓRIA
function carregarBancoDeDados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1.1. Lendo as Configurações Globais (Aba 0)
  var abaConfig = ss.getSheetByName("CONFIG_SELLER");
  var dadosConfig = abaConfig.getDataRange().getValues();
  // Linha 0 é o cabeçalho, Linha 1 são os dados.
  var config = {
    reputacao: dadosConfig[1][0],
    pisCofins: dadosConfig[1][1],
    irpj: dadosConfig[1][2],
    csll: dadosConfig[1][3]
  };

  // 1.2. Mapeando a TGFPRO (Catálogo e Logística Master)
  var abaPro = ss.getSheetByName("TGFPRO");
  var dadosPro = abaPro.getDataRange().getValues();
  var mapPro = {}; // Usaremos um objeto para busca em O(1) pelo SKU
  
  for (var i = 1; i < dadosPro.length; i++) {
    var sku = dadosPro[i][0];
    if (!sku) continue; // Pula linhas vazias
    
    mapPro[sku] = {
      tipoProduto: dadosPro[i][1], // Simples ou Kit
      origemProduto: dadosPro[i][5],
      custoAquisicao: parseFloat(dadosPro[i][6]) || 0,
      pesoKg: parseFloat(dadosPro[i][7]) || 0,
      comprimento: parseFloat(dadosPro[i][8]) || 0,
      largura: parseFloat(dadosPro[i][9]) || 0,
      altura: parseFloat(dadosPro[i][10]) || 0,
      margemPadrao: parseFloat(dadosPro[i][11]) || 0
    };
  }

  // 1.3. Mapeando a TGFKIT (A Receita de Bolo)
  var abaKit = ss.getSheetByName("TGFKIT");
  var dadosKit = abaKit.getDataRange().getValues();
  var mapKit = {}; // Chave será o SKU_KIT, o valor será um Array de Componentes
  
  for (var j = 1; j < dadosKit.length; j++) {
    var skuKit = dadosKit[j][0];
    if (!skuKit) continue;
    
    // Se o kit ainda não existe no mapa, cria um array vazio para ele
    if (!mapKit[skuKit]) {
      mapKit[skuKit] = [];
    }
    
    mapKit[skuKit].push({
      skuComponente: dadosKit[j][1],
      qtdComponente: parseFloat(dadosKit[j][2]) || 1,
      margemKit: dadosKit[j][3] !== "" ? parseFloat(dadosKit[j][3]) : null // Pode ser nula se não houver tática
    });
  }

  return { config: config, produtos: mapPro, kits: mapKit };
}

// 2. O ENGENHEIRO DO BLOCO VIRTUAL (Aglutinador Top-Down)
function construirBlocoVirtual(skuAnunciado, qtdNoAnuncio, tipoMargem, margemCustomizada, db) {
  var prodMaster = db.produtos[skuAnunciado];
  if (!prodMaster) return null; // Trava de segurança: SKU não existe na TGFPRO

  var bloco = {
    custoTotal: 0,
    margemPonderada: 0,
    pesoFisicoMaster: prodMaster.pesoKg * qtdNoAnuncio,
    cubagemMaster: ((prodMaster.comprimento * prodMaster.largura * prodMaster.altura) / 6000) * qtdNoAnuncio,
    origemICMSArray: [] // Guardará {pesoFinanceiro, aliquotaICMS} para o liquidificador fiscal
  };

  // 2.1. CÁLCULO SE FOR PRODUTO SIMPLES
  if (prodMaster.tipoProduto === "Simples") {
    bloco.custoTotal = prodMaster.custoAquisicao * qtdNoAnuncio;
    
    // Definição da Margem Simples
    if (tipoMargem === "Do anúncio") {
      bloco.margemPonderada = margemCustomizada;
    } else {
      bloco.margemPonderada = prodMaster.margemPadrao;
    }

    // Regra da Alíquota de Origem (Nacional vs Importado)
    var aliquotaOrigem = (prodMaster.origemProduto === 1 || prodMaster.origemProduto === 2 || prodMaster.origemProduto === 3 || prodMaster.origemProduto === 8) ? 0.04 : 0.12;
    
    // Como é simples, o peso financeiro dele é 100% (1.0)
    bloco.origemICMSArray.push({ pesoFinanceiro: 1.0, aliquota: aliquotaOrigem });
  } 
  
  // 2.2. CÁLCULO SE FOR UM KIT (O liquidificador algébrico)
  else if (prodMaster.tipoProduto === "Kit") {
    var componentes = db.kits[skuAnunciado];
    if (!componentes) return null; // Kit sem receita

    var lucroAbsolutoTotal = 0;

    // Loop pelas partes para descobrir o custo e o lucro esperado (em R$) de cada uma
    for (var k = 0; k < componentes.length; k++) {
      var comp = componentes[k];
      var dadosComp = db.produtos[comp.skuComponente];
      
      var custoParte = (dadosComp.custoAquisicao * comp.qtdComponente) * qtdNoAnuncio;
      bloco.custoTotal += custoParte;

      var margemDestaParte = 0;
      // Árvore de decisão da margem do kit
      if (tipoMargem === "Do anúncio") {
        margemDestaParte = margemCustomizada;
      } else if (tipoMargem === "Do kit" && comp.margemKit !== null) {
        margemDestaParte = comp.margemKit;
      } else {
        margemDestaParte = dadosComp.margemPadrao; // Fallback para "Do produto"
      }

      var lucroParte = custoParte * margemDestaParte;
      lucroAbsolutoTotal += lucroParte;

      // Descobrindo a alíquota de origem desta peça específica
      var aliquotaParte = (dadosComp.origemProduto === 1 || dadosComp.origemProduto === 2 || dadosComp.origemProduto === 3 || dadosComp.origemProduto === 8) ? 0.04 : 0.12;
      
      // Empurramos o Valor Alvo (Custo + Lucro) e a alíquota para o array para ponderar depois
      bloco.origemICMSArray.push({
        valorAlvoAbsoluto: custoParte + lucroParte,
        aliquota: aliquotaParte
      });
    }

    // Calculando a Margem Ponderada Final do Kit inteiro
    bloco.margemPonderada = lucroAbsolutoTotal / bloco.custoTotal;

    // Calculando a Carga Tributária Ponderada (ICMS Sintético)
    var valorAlvoTotalDoBloco = bloco.custoTotal + lucroAbsolutoTotal;
    var aliquotaSinteticaAcumulada = 0;

    for (var m = 0; m < bloco.origemICMSArray.length; m++) {
      var itemICMS = bloco.origemICMSArray[m];
      var pesoProporcional = itemICMS.valorAlvoAbsoluto / valorAlvoTotalDoBloco;
      aliquotaSinteticaAcumulada += (itemICMS.aliquota * pesoProporcional);
    }

    // Sobrescrevemos o array original deixando apenas o resultado sintético de 100% de peso
    bloco.origemICMSArray = [{ pesoFinanceiro: 1.0, aliquota: aliquotaSinteticaAcumulada }];
  }

  return bloco;
}



/**
 * MÓDULO 2: O MOTOR FINANCEIRO (CORE PRICING)
 * Responsabilidade: Receber o Bloco Virtual, aplicar a carga tributária,
 * cruzar com a matriz de fretes do ML e encontrar o Preço de Venda Final.
 */

function calcularPrecoMLB(blocoVirtual, config, taxaCategoriaML) {
  if (!blocoVirtual) return "ERRO: Bloco Vazio";

  // --- 1. CARGA TRIBUTÁRIA E DIVISOR (A Tese do Século) ---
  // Puxamos a alíquota que o Módulo 1 já ponderou perfeitamente
  var cargaIcmsTotal = blocoVirtual.origemICMSArray[0].aliquota;

  // Fator Federal Ajustado (PIS/COFINS sobre base sem ICMS)
  var fatorFederaisAjustado = config.pisCofins * (1 - cargaIcmsTotal) + config.irpj + config.csll;

  // O Divisor Mágico: 1 - (Comissão ML + Margem + Impostos)
  var divisor = 1 - (taxaCategoriaML + blocoVirtual.margemPonderada + cargaIcmsTotal + fatorFederaisAjustado);

  if (divisor <= 0) return "ERRO: Divisor Negativo";

  // --- 2. CONFIGURAÇÃO DAS FAIXAS (TIERS DO MERCADO LIVRE) ---
  // No JS, em vez de vários Arrays soltos, usamos um Array de Objetos para organizar as regras
  var faixas = [
    { min: 0.01,   max: 12.50,  taxaFixa: -1 },   // 0: Regra dos 50%
    { min: 12.51,  max: 28.99,  taxaFixa: 6.25 }, // 1: Taxa Fixa
    { min: 29.00,  max: 49.99,  taxaFixa: 6.50 }, // 2: Taxa Fixa
    { min: 50.00,  max: 78.99,  taxaFixa: 6.75 }, // 3: Taxa Fixa
    { min: 79.00,  max: 99.99,  taxaFixa: 0 },    // 4: Frete Grátis
    { min: 100.00, max: 119.99, taxaFixa: 0 },    // 5: Frete Grátis
    { min: 120.00, max: 149.99, taxaFixa: 0 },    // 6: Frete Grátis
    { min: 150.00, max: 199.99, taxaFixa: 0 },    // 7: Frete Grátis
    { min: 200.00, max: 999999, taxaFixa: 0 }     // 8: Frete Grátis
  ];

  // --- 3. PESAGEM: BALANÇA VS TRENA ---
  // Math.max pega automaticamente o maior valor entre os dois informados
  var pesoCobrado = Math.max(blocoVirtual.pesoFisicoMaster, blocoVirtual.cubagemMaster);

  // --- 4. MOTOR DE BUSCA DO MELHOR PREÇO (O Loop Top-Down) ---
  var melhorPreco = 999999;
  var precoCalculado = 0;

  for (var i = 0; i < faixas.length; i++) {
    var tier = faixas[i];
    var custoFrete = 0;
    var custoFixo = 0;

    if (i === 0) {
      // Regra Especial: Baixíssimo Ticket
      if ((divisor - 0.5) > 0) {
        precoCalculado = blocoVirtual.custoTotal / (divisor - 0.5);
      } else {
        precoCalculado = 999999;
      }
    } else if (i <= 3) {
      // Regras de Taxa Fixa (Abaixo de R$ 79)
      custoFixo = tier.taxaFixa;
      precoCalculado = (blocoVirtual.custoTotal + custoFixo) / divisor;
    } else {
      // Regras de Frete Grátis Subsidiado (Acima de R$ 79)
      custoFrete = calcularFreteMatriz(pesoCobrado, i, config.reputacao);
      precoCalculado = (blocoVirtual.custoTotal + custoFrete) / divisor;
    }

    // Blindagem de Ponto Flutuante: Arredonda para 2 casas decimais antes de testar
    var precoArredondado = Math.round(precoCalculado * 100) / 100;

    // Validação: O preço encontrado "cabe" dentro da regra em que foi testado?
    if (precoArredondado >= tier.min && precoArredondado <= tier.max) {
      if (precoCalculado < melhorPreco) {
        melhorPreco = precoCalculado;
      }
    }
  }

  // --- 5. ZONA MORTA (Fallback de Segurança) ---
  if (melhorPreco === 999999) {
    var freteFallback = calcularFreteMatriz(pesoCobrado, 8, config.reputacao);
    melhorPreco = (blocoVirtual.custoTotal + freteFallback) / divisor;
  }

  return Math.round(melhorPreco * 100) / 100;
}

// --- FUNÇÃO AUXILIAR: MATRIZ DE FRETES ---
function calcularFreteMatriz(peso, indiceFaixa, reputacao) {
  var valorBase = 0;

  // A cascata de peso (Idêntica ao seu VBA)
  if (peso <= 0.3) valorBase = 39.90;
  else if (peso <= 0.5) valorBase = 42.90;
  else if (peso <= 1.0) valorBase = 44.90;
  else if (peso <= 2.0) valorBase = 46.90;
  else if (peso <= 3.0) valorBase = 49.90;
  else if (peso <= 4.0) valorBase = 53.90;
  else if (peso <= 5.0) valorBase = 56.90;
  else if (peso <= 9.0) valorBase = 88.90;
  else if (peso <= 13.0) valorBase = 131.90;
  else if (peso <= 17.0) valorBase = 146.90;
  else if (peso <= 23.0) valorBase = 171.90;
  else if (peso <= 30.0) valorBase = 197.90;
  else if (peso <= 40.0) valorBase = 203.90;
  else if (peso <= 50.0) valorBase = 210.90;
  else if (peso <= 60.0) valorBase = 224.90;
  else if (peso <= 70.0) valorBase = 240.90;
  else if (peso <= 80.0) valorBase = 251.90;
  else if (peso <= 90.0) valorBase = 279.90;
  else if (peso <= 100.0) valorBase = 319.90;
  else if (peso <= 125.0) valorBase = 357.90;
  else if (peso <= 150.0) valorBase = 379.90;
  else valorBase = 498.90;

  // Limpeza de String: Remove espaços e coloca em maiúsculo
  var repFormatada = String(reputacao).trim().toUpperCase();
  var percentualDesconto = 0;

  if (repFormatada === "LÍDER" || repFormatada === "LIDER" || repFormatada === "VERDE" || repFormatada === "CINZA") {
    switch (indiceFaixa) {
      case 4: percentualDesconto = 0.70; break;
      case 5: percentualDesconto = 0.65; break;
      case 6: percentualDesconto = 0.60; break;
      case 7: percentualDesconto = 0.55; break;
      case 8: percentualDesconto = 0.50; break;
    }
  } else if (repFormatada === "AMARELA") {
    switch (indiceFaixa) {
      case 4: percentualDesconto = 0.60; break;
      case 5: percentualDesconto = 0.55; break;
      case 6: percentualDesconto = 0.50; break;
      case 7: percentualDesconto = 0.45; break;
      case 8: percentualDesconto = 0.40; break;
    }
  }

  return valorBase * (1 - percentualDesconto);
}



/**
 * MÓDULO 3: O GATILHO DE EXECUÇÃO (CONTROLLER)
 * Responsabilidade: Varrer a aba TGFADS, orquestrar os cálculos
 * e gravar os resultados em lote (batch) para máxima performance.
 */

function processarPrecificacaoEmMassa() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaAds = ss.getSheetByName("TGFADS");
  
  // 1. Inicia o cronômetro mental e carrega o banco de dados
  var db = carregarBancoDeDados();
  
  // 2. Lê todos os anúncios de uma vez só
  // Assumindo que os dados começam na linha 2 e vão até a última linha preenchida
  var ultimaLinha = abaAds.getLastRow();
  if (ultimaLinha < 2) return; // Se só tiver cabeçalho, aborta.
  
  // getRange(linhaInicial, colunaInicial, qtdLinhas, qtdColunas)
  var rangeAds = abaAds.getRange(2, 1, ultimaLinha - 1, abaAds.getLastColumn());
  var dadosAds = rangeAds.getValues();
  
  // 3. Prepara um Array vazio para guardar os preços calculados
  // Gravar célula por célula é lento. Vamos guardar tudo aqui e cuspir de uma vez.
  var resultadosPrecoFinal = [];
  
  // 4. O Loop de Orquestração
  for (var i = 0; i < dadosAds.length; i++) {
    var linha = dadosAds[i];
    
    // Mapeamento das colunas da TGFADS (Lembrando: Índice começa no ZERO)
    var skuAnunciado = linha[1];            // Coluna B
    var qtdNoAnuncio = parseFloat(linha[2]) || 1; // Coluna C
    var taxaCategoriaML = parseFloat(linha[4]) || 0; // Coluna E
    var tipoMargem = linha[5];              // Coluna F
    var margemCustomizada = parseFloat(linha[6]) || 0; // Coluna G
    
    // Ignora linhas vazias
    if (!skuAnunciado) {
      resultadosPrecoFinal.push([""]); // Empurra uma célula vazia para manter o alinhamento
      continue;
    }
    
    // 5. Aciona o Engenheiro do Bloco Virtual (Módulo 1)
    var bloco = construirBlocoVirtual(skuAnunciado, qtdNoAnuncio, tipoMargem, margemCustomizada, db);
    
    if (!bloco) {
      resultadosPrecoFinal.push(["ERRO: SKU não encontrado"]);
      continue;
    }
    
    // 6. Aciona o Motor Financeiro (Módulo 2)
    var precoFinal = calcularPrecoMLB(bloco, db.config, taxaCategoriaML);
    
    // Empurra o resultado encapsulado em um array (exigência do Sheets para colunas)
    resultadosPrecoFinal.push([precoFinal]);
  }
  
  // 7. A Injeção Final em Lote (Batch Write)
  // Coluna M é a 13ª coluna. Vamos injetar os dados a partir da linha 2.
  var rangeSaida = abaAds.getRange(2, 13, resultadosPrecoFinal.length, 1);
  rangeSaida.setValues(resultadosPrecoFinal);
}



/**
 * MÓDULO 4: INTERFACE DE USUÁRIO (UI)
 * Responsabilidade: Criar o menu nativo no Google Sheets para o operador.
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // Cria o menu principal com o nome da sua consultoria
  ui.createMenu('360 Gestão')
    .addItem('⚡ Recalcular Preços', 'processarPrecificacaoEmMassa')
    .addSeparator() // Linha divisória visual
    .addItem('ℹ️ Sobre o Motor', 'exibirSobre')
    .addToUi();
}

// Função auxiliar apenas para dar um feedback no botão 'Sobre'
function exibirSobre() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    'Motor de Precificação Dinâmica',
    'Versão 2.0 (Top-Down / Bloco Virtual)\nDesenvolvido para operações de alta performance no Mercado Livre.',
    ui.ButtonSet.OK
  );
}