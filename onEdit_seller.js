/**
 * MÓDULO 5: AUTOMAÇÕES DE INTERFACE E TRAVAS FISCAIS (EVENT LISTENERS)
 * Responsabilidade: Escutar alterações na interface do usuário, 
 * preencher alíquotas automáticas e bloquear erros de preenchimento tributário.
 */

function onEdit(e) {
  if (!e) return; // Trava de segurança caso seja rodado manualmente
  
  var aba = e.source.getActiveSheet();
  var nomeAba = aba.getName();
  var linhaEditada = e.range.getRow();
  var colunaEditada = e.range.getColumn();
  
// =========================================================================
  // 1. AUTOMAÇÃO DA ABA CONFIG_SELLER (Regimes e UF Padrão)
  // =========================================================================
  if (nomeAba === "CONFIG_SELLER") {
    
    // A. Automação de Regime Tributário (Célula E2 - Coluna 5)
    if (linhaEditada === 2 && colunaEditada === 5) {
      var regime = e.value; 
      
      var celulaPisCofins = aba.getRange("B2");
      var celulaIrpj = aba.getRange("C2");
      var celulaCsll = aba.getRange("D2");
      var celulaCredPisCofins = aba.getRange("F2");
      var celulaCredBCPisCofins = aba.getRange("G2");
      
      if (regime === "Lucro Real") {
        celulaPisCofins.setValue(0.0925); 
        celulaIrpj.setValue(0.00);        
        celulaCsll.setValue(0.00);        
      } 
      else if (regime === "Lucro Presumido") {
        celulaPisCofins.setValue(0.0365);         
        celulaIrpj.setValue(0.0120);              
        celulaCsll.setValue(0.0108);              
        celulaCredPisCofins.setValue("Não");      
        celulaCredBCPisCofins.setValue("Nenhum"); 
      }
    }

    // B. NOVA FEATURE: Propagação de UF de Destino (Célula J2 - Coluna 10)
    if (linhaEditada === 2 && colunaEditada === 10) {
      var siglaUF = e.value ? String(e.value).trim().toUpperCase() : "";
      if (!siglaUF) return; // Se apagarem a célula, não faz nada

      var ss = e.source;
      var abaIcms = ss.getSheetByName("TGFICMS");
      
      if (!abaIcms) {
        ss.toast("Aba TGFICMS não encontrada para buscar as alíquotas.", "⚠️ Erro de Arquitetura", 5);
        return;
      }

      var dadosIcms = abaIcms.getDataRange().getValues();
      var aliqIcms = null;
      var aliqFecop = null;

      // 1. Varre a TGFICMS buscando a sigla (Coluna B é o índice 1)
      for (var i = 1; i < dadosIcms.length; i++) {
        if (dadosIcms[i][1] === siglaUF) { 
          aliqIcms = parseFloat(dadosIcms[i][3]) || 0; // Coluna D
          aliqFecop = parseFloat(dadosIcms[i][4]) || 0; // Coluna E
          break;
        }
      }

      if (aliqIcms === null) {
        ss.toast("Sigla " + siglaUF + " não encontrada na tabela TGFICMS.", "⚠️ UF Inválida", 5);
        return;
      }

      // 2. O Lançamento em Massa na TGFADS (Batch Update)
      var abaAds = ss.getSheetByName("TGFADS");
      var ultimaLinhaAds = abaAds.getLastRow();

      if (ultimaLinhaAds > 1) {
        var numLinhas = ultimaLinhaAds - 1;
        var matrizAtualizacao = [];

        // Montamos uma matriz virtual com o tamanho exato da tabela de anúncios
        for (var r = 0; r < numLinhas; r++) {
          matrizAtualizacao.push([aliqIcms, aliqFecop]);
        }

        // Injetamos a matriz de uma só vez nas Colunas H (8) e I (9)
        abaAds.getRange(2, 8, numLinhas, 2).setValues(matrizAtualizacao);

        // Feedback de sucesso em tempo real para o operador
        ss.toast(
          "O motor atualizou todos os anúncios para o destino " + siglaUF + " (" + (aliqIcms * 100).toFixed(2) + "% ICMS | " + (aliqFecop * 100).toFixed(2) + "% FECOP).", 
          "🚀 Propagação Geográfica Concluída", 
          8
        );
      }
    }
  }

  // =========================================================================
  // 2. TRAVA FISCAL DA ABA TGFPRO (Convênio ICMS 52/91 vs Resolução 13/2012)
  // =========================================================================
  else if (nomeAba === "TGFPRO") {
    if (linhaEditada < 2) return; // Ignora o cabeçalho
    
    // Coluna F (Origem) é o índice 6. Coluna O (Redução BC) é o índice 15.
    if (colunaEditada === 6 || colunaEditada === 15) {
      
      var celulaOrigem = aba.getRange(linhaEditada, 6);
      var celulaReducao = aba.getRange(linhaEditada, 15);
      
      var origem = String(celulaOrigem.getValue()).trim();0
      
      // Avalia se a origem pertence ao grupo de importados com alíquota de 4%
      if (origem === "1" || origem === "2" || origem === "3" || origem === "8") {
        
        var reducaoAtual = parseFloat(celulaReducao.getValue()) || 0;
        
        // Se houver uma redução digitada, a trava é acionada
        if (reducaoAtual > 0) {
          celulaReducao.setValue(0); // Zera o campo imediatamente
          
          // Exibe um alerta educacional para o operador no canto inferior direito
          e.source.toast(
            "Origens 1, 2, 3 e 8 já possuem alíquota interestadual de 4%. Como este valor é menor que o piso do benefício (5,14%), a redução foi zerada para manter a carga mínima.",
            "🛡️ Trava Fiscal 360 Ativada",
            10 // Tempo de exibição em segundos
          );
        }
      }
    }
  }
}