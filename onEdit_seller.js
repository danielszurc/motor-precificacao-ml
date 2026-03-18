/**
 * FRONT-END: INTEGRAÇÃO 360 GESTÃO IND & AUT
 * Responsabilidade: Desenhar a interface do usuário (Menu) e 
 * delegar os comandos para o motor central (Biblioteca Motor360).
 */

// =========================================================
// 1. INTERFACE DE USUÁRIO (MENU)
// =========================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('360 Gestão')
    .addItem('⚡ Recalcular Preços', 'acionarMotorRemoto')
    .addSeparator() 
    .addItem('ℹ️ Sobre o Motor', 'exibirSobre')
    .addToUi();
}

function exibirSobre() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    'Motor de Precificação Dinâmica',
    'Versão 2.1 (SaaS Edition)\nArquitetura Fiscal Top-Down com Deflação de IPI e Tese do Século.\nDesenvolvido pela 360 Gestão Ind & Aut.',
    ui.ButtonSet.OK
  );
}

// =========================================================
// 2. GATILHOS DE EXECUÇÃO (DELEGATORS)
// =========================================================

// Função casca que chama a biblioteca invisível
function acionarMotorRemoto() {
  // O Motor360 é a nossa biblioteca importada
  Motor360.processarPrecificacaoEmMassa();
}

// A Ponte de Eventos para a Trava Fiscal
function onEdit(e) {
  if (!e) return;
  // Despacha a edição para o policial fiscal no back-end
  Motor360.onEdit(e);
}