/**
 * INTEGRAÇÃO 360 GESTÃO IND & AUT
 * Este script é apenas um gatilho para a biblioteca central.
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('360 Gestão')
    .addItem('⚡ Recalcular Preços', 'acionarMotorRemoto')
    .addToUi();
}

// Função casca que chama a biblioteca invisível
function acionarMotorRemoto() {
  // Motor360 é o identificador que você deu no Passo 3
  Motor360.processarPrecificacaoEmMassa();
}