/**
 * COMO USAR:
 * 1. Crie uma planilha no Google Drive chamada, por exemplo, "Inscricoes CONFRA2026".
 * 2. Na primeira linha (linha 1), crie estes cabeçalhos, um por coluna, EXATAMENTE nesta ordem:
 *    A: Timestamp
 *    B: Nome Responsavel
 *    C: RG Responsavel
 *    D: Telefone
 *    E: Total Minimo
 *    F: Total Planejado
 *    G: Familiares (JSON)
 *    H: Valor Ago      I: Confirmado Ago
 *    J: Valor Set      K: Confirmado Set
 *    L: Valor Out      M: Confirmado Out
 *    N: Valor Nov      O: Confirmado Nov
 *    P: Valor Dez      Q: Confirmado Dez
 *    R: Vai de Onibus (Sim / Nao)
 * 3. A coluna "Confirmado <Mês>" é de texto livre: a Secretaria escreve QUALQUER coisa
 *    (ex: "OK", "PAGO", a data, um "x") para marcar como confirmado. Deixar a célula em
 *    branco significa que a parcela ainda está pendente. Não precisa ser um valor específico.
 * 4. Nessa planilha, vá em Extensões > Apps Script, apague o conteúdo padrão e cole este arquivo.
 * 5. Implantar > Nova implantação > tipo "App da Web" (Executar como: Eu / Acesso: Qualquer pessoa).
 * 6. Copie a URL (termina em /exec) e cole na landing page, na constante APPS_SCRIPT_URL.
 *
 * COMO A SECRETARIA CONFIRMA UM PAGAMENTO:
 * Ao conferir o Pix no extrato do banco, ela encontra a linha da pessoa (pelo nome/telefone)
 * e marca a caixinha "Confirmado <Mês>" correspondente. Pronto — não precisa mexer em mais nada.
 * A partir daí, esse mês fica protegido: mesmo que a pessoa reenvie o formulário para ajustar
 * outros meses, o valor e a confirmação desse mês específico NUNCA são sobrescritos.
 *
 * COMPORTAMENTO: se o telefone já existir numa linha, essa linha é ATUALIZADA em vez de criar
 * uma nova. Meses já confirmados na planilha são preservados e ignoram o que vier do formulário.
 */

var MESES_COLS = {
  Ago: { valor: 8,  confirmado: 9  },
  Set: { valor: 10, confirmado: 11 },
  Out: { valor: 12, confirmado: 13 },
  Nov: { valor: 14, confirmado: 15 },
  Dez: { valor: 16, confirmado: 17 }
};

function estaConfirmado(valorCelula) {
  return valorCelula !== '' && valorCelula !== null && valorCelula !== undefined;
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  var telefoneNormalizado = (data.respTel || '').replace(/\D/g, '');
  var parcelasEnviadas = JSON.parse(data.parcelas || '[]'); // [{mes:'Ago', valor: 50}, ...]

  var linhas = sheet.getDataRange().getValues();
  var linhaExistente = -1;
  for (var i = 1; i < linhas.length; i++) {
    var telCelula = (linhas[i][3] || '').toString().replace(/\D/g, '');
    if (telCelula && telCelula === telefoneNormalizado) {
      linhaExistente = i + 1; // 1-indexado para o Sheets
      break;
    }
  }

  var novaLinha = new Array(18);
  novaLinha[0] = new Date();
  novaLinha[1] = data.respNome || '';
  novaLinha[2] = data.respRG || '';
  novaLinha[3] = data.respTel || '';
  novaLinha[4] = data.totalMinimo || '';
  novaLinha[5] = data.totalDigitado || '';
  novaLinha[6] = data.familiares || '[]';
  novaLinha[17] = data.vaiOnibus || 'Sim';

  for (var mes in MESES_COLS) {
    var col = MESES_COLS[mes];
    var confirmadoAtual = false;

    if (linhaExistente > 0) {
      confirmadoAtual = estaConfirmado(sheet.getRange(linhaExistente, col.confirmado).getValue());
    }

    if (confirmadoAtual) {
      // Mês já confirmado (célula preenchida com qualquer texto): preserva valor e confirmação.
      novaLinha[col.valor - 1] = sheet.getRange(linhaExistente, col.valor).getValue();
      novaLinha[col.confirmado - 1] = sheet.getRange(linhaExistente, col.confirmado).getValue();
    } else {
      var item = parcelasEnviadas.filter(function (p) { return p.mes === mes; })[0];
      novaLinha[col.valor - 1] = item ? item.valor : 0;
      novaLinha[col.confirmado - 1] = '';
    }
  }

  if (linhaExistente > 0) {
    sheet.getRange(linhaExistente, 1, 1, novaLinha.length).setValues([novaLinha]);
  } else {
    sheet.appendRow(novaLinha);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e.parameter.telefone) {
    return buscarPorTelefone(e);
  }
  return ContentService.createTextOutput('O script está no ar. Use POST para enviar dados.');
}

function buscarPorTelefone(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var telefoneBusca = (e.parameter.telefone || '').replace(/\D/g, '');
  var callback = e.parameter.callback || 'callback';
  var linhas = sheet.getDataRange().getValues();
  var resultado = { encontrado: false };

  for (var i = 1; i < linhas.length; i++) {
    var telCelula = (linhas[i][3] || '').toString().replace(/\D/g, '');
    if (telCelula && telCelula === telefoneBusca) {
      var parcelas = {};
      for (var mes in MESES_COLS) {
        var col = MESES_COLS[mes];
        parcelas[mes] = {
          valor: linhas[i][col.valor - 1] || 0,
          confirmado: estaConfirmado(linhas[i][col.confirmado - 1])
        };
      }
      resultado = {
        encontrado: true,
        respNome: linhas[i][1] || '',
        respRG: linhas[i][2] || '',
        familiares: linhas[i][6] || '[]',
        vaiOnibus: linhas[i][17] || 'Sim',
        parcelas: parcelas
      };
      break;
    }
  }

  var saida = callback + '(' + JSON.stringify(resultado) + ');';
  return ContentService.createTextOutput(saida).setMimeType(ContentService.MimeType.JAVASCRIPT);
}
