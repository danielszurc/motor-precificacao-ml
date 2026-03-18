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

// Esta função escuta a edição local e joga o "pacote de dados" (e) para o motor.
function onEdit(e) {
  // Verificação de segurança primária
  if (!e) return;
  
  // Despacha o evento diretamente para a função onEdit que está dentro do Motor360
  Motor360.onEdit(e);
}