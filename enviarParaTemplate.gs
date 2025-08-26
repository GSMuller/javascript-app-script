function enviarParaTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var support = ss.getSheetByName('Support');
  var teste = ss.getSheetByName('Teste');

  // Pega dados de F2:M20 (ignora cabeçalho)
  var dados = support.getRange('F2:M20').getValues();

  // Função para formatar número
  function formatarNumero(num) {
    if (typeof num === 'number') {
      return Utilities.formatString('%s', num.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}));
    }
    return num;
  }

  // Mescla OPx e OPx.Num com formatação
  var resultado = dados.map(function(linha) {
    return [
      linha[0] + '\n' + formatarNumero(linha[1]), // OP1 + OP1.Num
      linha[2] + '\n' + formatarNumero(linha[3]), // OP2 + OP2.Num
      linha[4] + '\n' + formatarNumero(linha[5]), // OP3 + OP3.Num
      linha[6] + '\n' + formatarNumero(linha[7])  // OP4 + OP4.Num
    ];
  });

  // Escreve na planilha Teste, de H8:K26
  teste.getRange(8, 8, resultado.length, resultado[0].length).setValues(resultado);
}
