/**
 * MÓDULO 5: AUTOMAÇÕES DE INTERFACE (EVENT LISTENERS)
 * Responsabilidade: Escutar alterações na aba CONFIG_SELLER e 
 * preencher automaticamente as alíquotas conforme o Regime Tributário.
 */

function onEdit(e) {
  if (!e) return; // Trava de segurança caso seja rodado manualmente
  
  var aba = e.source.getActiveSheet();
  
  // Só executa se a alteração for na aba de configurações
  if (aba.getName() !== "CONFIG_SELLER") return;
  
  var linhaEditada = e.range.getRow();
  var colunaEditada = e.range.getColumn();
  
  // Verifica se o usuário alterou a célula E2 (Regime Tributário)
  if (linhaEditada === 2 && colunaEditada === 5) {
    var regime = e.value; // O valor que o usuário acabou de selecionar
    
    // Mapeamento das células: B2 (PIS/COFINS), C2 (IRPJ), D2 (CSLL)
    var celulaPisCofins = aba.getRange("B2");
    var celulaIrpj = aba.getRange("C2");
    var celulaCsll = aba.getRange("D2");
    var celulaCredPisCofins = aba.getRange("F2");
    var celulaCredBCPisCofins = aba.getRange("G2");
    
    if (regime === "Lucro Real") {
      celulaPisCofins.setValue(0.0925); // 9,25%
      celulaIrpj.setValue(0.00);        // Zera IRPJ
      celulaCsll.setValue(0.00);        // Zera CSLL
    } 
    else if (regime === "Lucro Presumido") {
      celulaPisCofins.setValue(0.0365);         // 3,65%
      celulaIrpj.setValue(0.0120);              // 1,20% -> percentual mínimo de IRPJ para Presumido (geralmente é maior com o adicional de 10%)
      celulaCsll.setValue(0.0108);              // 1,08%
      celulaCredPisCofins.setValue("Não");      // Lucro Presumido não toma crédito de PIS/COFINS
      celulaCredBCPisCofins.setValue("Nenhum"); // Limpa a base de créditos de PIS/COFINS
    }
  }
}