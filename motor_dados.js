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

/*
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
*/

/**
 * MÓDULO 2 (V2.0): O MOTOR FINANCEIRO (CORE PRICING - FRETE PADRÃO ML)
 * Responsabilidade: Aplicar a carga tributária, varrer as matrizes tridimensionais
 * de envio do ML (Peso x Preço x Reputação) e encontrar o Preço de Venda Final.
 */

function calcularPrecoMLB(blocoVirtual, config, taxaCategoriaML, forcarFreteRapidoSub79) {
  if (!blocoVirtual) return "ERRO: Bloco Vazio";

  // --- 1. CARGA TRIBUTÁRIA E DIVISOR (A Tese do Século) ---
  var cargaIcmsTotal = blocoVirtual.origemICMSArray[0].aliquota;
  var fatorFederaisAjustado = config.pisCofins * (1 - cargaIcmsTotal) + config.irpj + config.csll;
  var divisor = 1 - (taxaCategoriaML + blocoVirtual.margemPonderada + cargaIcmsTotal + fatorFederaisAjustado);

  if (divisor <= 0) return "ERRO: Divisor Negativo/Margem Excessiva";

  // --- 2. CONFIGURAÇÃO DAS FAIXAS DE PREÇO E COLUNAS DA MATRIZ ---
  // A propriedade 'col' indica qual coluna da Tabela Cheia o algoritmo deve ler
  var faixasDePreco = [
    { min: 0.01,   max: 18.99,  col: 0 }, // Regra especial: Frete máx 50% do preço
    { min: 19.00,  max: 48.99,  col: 1 },
    { min: 49.00,  max: 78.99,  col: 2 },
    { min: 79.00,  max: 99.99,  col: 3 }, // Barreira do Frete Rápido
    { min: 100.00, max: 119.99, col: 4 },
    { min: 120.00, max: 149.99, col: 5 },
    { min: 150.00, max: 199.99, col: 6 },
    { min: 200.00, max: 999999, col: 7 }
  ];

  // --- 3. PESAGEM E BUSCA DO FRETE BASE (TABELA CHEIA VERMELHA) ---
  var pesoCobrado = Math.max(blocoVirtual.pesoFisicoMaster, blocoVirtual.cubagemMaster);
  
  // Função auxiliar que lê a matriz baseada no peso e na coluna
  var buscarFreteTabelaCheia = function(peso, colunaIndex) {
    // Array com os limites de peso em kg (Cada posição corresponde a uma linha da matriz abaixo)
    var limitesPeso = [0.3, 0.5, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 11.0, 13.0, 15.0, 17.0, 20.0, 25.0, 30.0, 40.0, 50.0, 60.0, 70.0, 80.0, 90.0, 100.0, 125.0, 150.0, 9999];
    
    // A Tabela Cheia (Validada no Excel)
    var matrizFrete = [
      [8.07, 9.36, 11.07, 24.70, 28.70, 32.90, 36.90, 41.90], // <= 0.3
      [8.50, 9.50, 11.21, 26.50, 30.90, 35.30, 39.70, 45.10], // <= 0.5
      [8.64, 9.64, 11.36, 27.70, 32.30, 36.90, 41.50, 47.30], // <= 1.0
      [8.79, 9.79, 11.50, 28.30, 32.90, 37.70, 42.30, 49.30], // <= 1.5
      [8.93, 9.93, 11.64, 28.90, 33.70, 38.50, 43.30, 49.30], // <= 2.0
      [9.07, 11.36, 12.21, 31.50, 36.70, 42.10, 47.30, 52.50], // <= 3.0
      [9.21, 11.64, 12.79, 34.10, 39.70, 45.30, 51.10, 56.70], // <= 4.0
      [9.36, 11.93, 13.93, 36.90, 43.10, 49.30, 55.50, 61.50], // <= 5.0
      [9.50, 12.21, 14.21, 50.90, 57.10, 65.30, 71.50, 79.50], // <= 6.0
      [9.64, 12.50, 14.50, 54.10, 62.10, 72.10, 80.10, 88.10], // <= 7.0
      [9.79, 12.79, 14.79, 57.70, 67.30, 76.90, 86.50, 96.10], // <= 8.0
      [9.93, 13.07, 15.07, 59.30, 69.10, 79.10, 88.90, 98.70], // <= 9.0
      [10.07, 13.64, 15.64, 82.50, 96.10, 109.90, 123.50, 137.30], // <= 11.0
      [10.21, 14.21, 16.21, 84.30, 98.50, 112.50, 126.50, 140.50], // <= 13.0
      [10.36, 14.50, 16.50, 90.10, 104.90, 119.90, 134.90, 149.90], // <= 15.0
      [10.50, 14.79, 16.79, 97.10, 112.10, 127.10, 141.50, 157.30], // <= 17.0
      [10.64, 15.07, 17.07, 109.50, 127.70, 145.90, 164.10, 182.30], // <= 20.0
      [10.93, 15.64, 17.36, 128.10, 150.10, 169.50, 190.70, 211.90], // <= 25.0
      [11.07, 15.93, 17.64, 131.90, 150.90, 171.10, 192.50, 213.90], // <= 30.0
      [11.21, 16.21, 17.93, 135.50, 157.90, 177.90, 198.30, 214.10], // <= 40.0
      [11.36, 16.50, 18.21, 140.50, 162.10, 184.10, 205.10, 221.50], // <= 50.0
      [11.50, 16.79, 18.50, 149.90, 172.90, 196.30, 218.70, 236.30], // <= 60.0
      [11.64, 17.07, 18.79, 160.50, 185.90, 210.10, 234.30, 253.10], // <= 70.0
      [11.79, 17.36, 19.07, 167.90, 194.10, 219.70, 244.90, 264.50], // <= 80.0
      [11.93, 17.64, 19.36, 186.50, 214.90, 244.10, 272.10, 293.90], // <= 90.0
      [12.07, 17.93, 19.64, 213.10, 247.90, 279.10, 311.10, 335.90], // <= 100.0
      [12.21, 18.21, 19.93, 238.50, 276.10, 312.10, 347.90, 375.90], // <= 125.0
      [12.36, 18.21, 20.21, 253.10, 292.30, 331.30, 369.30, 398.90], // <= 150.0
      [12.50, 18.21, 20.50, 332.30, 384.90, 435.10, 485.10, 523.90]  // > 150.0
    ];

    // Encontra a linha correspondente ao peso
    var linhaIndex = limitesPeso.length - 1; // Fallback para o mais pesado
    for (var p = 0; p < limitesPeso.length; p++) {
      if (peso <= limitesPeso[p]) {
        linhaIndex = p;
        break;
      }
    }
    return matrizFrete[linhaIndex][colunaIndex];
  };

  // --- 4. APLICAÇÃO DOS DESCONTOS AUDITADOS (VERDE VS AMARELA) ---
  var repFormatada = String(config.reputacao).trim().toUpperCase();
  var isVerdeOuLider = (repFormatada === "VERDE" || repFormatada === "LÍDER" || repFormatada === "LIDER" || repFormatada === "CINZA");
  var isAmarela = (repFormatada === "AMARELA");

  // --- 5. MOTOR DE BUSCA DO MELHOR PREÇO (O Loop de Tiers) ---
  var melhorPreco = 999999;

  for (var i = 0; i < faixasDePreco.length; i++) {
    var tier = faixasDePreco[i];
    var precoCalculado = 0;

    // 5.1. Busca o frete na Tabela Cheia para a coluna que estamos testando
    var freteCheio = buscarFreteTabelaCheia(pesoCobrado, tier.col);
    
    // 5.2. Aplica o multiplicador de desconto conforme a faixa (Abaixo ou Acima de R$ 79)
    var isAbaixo79 = (tier.col <= 2);
    var desconto = 0;

    if (isAbaixo79 && forcarFreteRapidoSub79) {
      // O seller quer rankear melhor! Pegamos o custo da Coluna 3 (Tabela de R$ 79)
      freteCheio = buscarFreteTabelaCheia(pesoCobrado, 3);
      // E aplicamos os descontos premium
      if (isVerdeOuLider) desconto = 0.50;
      else if (isAmarela) desconto = 0.40;
    }
    else if (isAbaixo79) {
      // Fluxo Padrão Normal (< 79)
      if (isVerdeOuLider) desconto = 0.30;
      else if (isAmarela) desconto = 0.20;
    }
    else {
      // Fluxo Padrão Normal (>= 79)
      if (isVerdeOuLider) desconto = 0.50;
      else if (isAmarela) desconto = 0.40;
    }

    var freteFinalSendoTestado = freteCheio * (1 - desconto);

    // 5.3. A Álgebra do Preço
    if (tier.col === 0) {
      // REGRA ESPECIAL: Anúncios até 18.99 pagam no máx 50% do valor do produto em frete.
      var precoSemTrava = (blocoVirtual.custoTotal + freteFinalSendoTestado) / divisor;
      
      // Se o frete calculado representa mais de 50% do preço de venda, a trava do ML é acionada.
      if (freteFinalSendoTestado > (precoSemTrava * 0.5)) {
        // Recalculando o preço isolando a variável Custo: Preço = Custo / (Divisor - 0.5)
        if ((divisor - 0.5) > 0) {
          precoCalculado = blocoVirtual.custoTotal / (divisor - 0.5);
        } else {
          precoCalculado = 999999; // Margem inviável para itens muito baratos
        }
      } else {
        precoCalculado = precoSemTrava; // Passou liso, o frete não atingiu 50%
      }
    } else {
      // REGRA PADRÃO PARA AS DEMAIS FAIXAS
      precoCalculado = (blocoVirtual.custoTotal + freteFinalSendoTestado) / divisor;
    }

    // 5.4. Blindagem Decimal e Teste de Validação
    var precoArredondado = Math.round(precoCalculado * 100) / 100;

    // Se o preço calculado "couber" matematicamente dentro da faixa que ditou o custo do frete, achamos o candidato perfeito!
    if (precoArredondado >= tier.min && precoArredondado <= tier.max) {
      if (precoCalculado < melhorPreco) {
        melhorPreco = precoCalculado;
      }
    }
  }

  // --- 6. ZONA MORTA (Fallback de Segurança para preços astronômicos) ---
  if (melhorPreco === 999999) {
    var freteCheioFallback = buscarFreteTabelaCheia(pesoCobrado, 7); // Última coluna (A partir de 200)
    var descFallback = isVerdeOuLider ? 0.50 : (isAmarela ? 0.40 : 0);
    var freteFallbackFinal = freteCheioFallback * (1 - descFallback);
    melhorPreco = (blocoVirtual.custoTotal + freteFallbackFinal) / divisor;
  }

  return Math.round(melhorPreco * 100) / 100;
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

    // Força Frete Grátis Rápido mesmo se o preço do anúncio for menor do que R$79
    var forcarFreteRapido = (String(linha[10]).trim().toUpperCase() === "SIM");
    
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
    var precoFinal = calcularPrecoMLB(bloco, db.config, taxaCategoriaML, forcarFreteRapido);
    
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