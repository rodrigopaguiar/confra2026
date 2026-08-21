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
 *    S: Email
 *
 * ENVIO AUTOMATICO DE EMAIL: ao confirmar um cadastro NOVO (nao em edicoes), o script envia
 * automaticamente um email de confirmacao em HTML (com o visual da landing page) para o
 * endereco informado, usando a propria conta Google desta planilha (MailApp) — gratuito, sem
 * necessidade de configurar nada a mais. O template completo (com as variantes de onibus/carro
 * e o placeholder de parcelas) esta embutido no topo deste arquivo, nas variaveis
 * TEMPLATE_EMAIL_CONFIRMACAO, BLOCO_TRANSPORTE_ONIBUS e BLOCO_TRANSPORTE_CARRO.
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

  var novaLinha = new Array(19);
  novaLinha[0] = new Date();
  novaLinha[1] = data.respNome || '';
  novaLinha[2] = data.respRG || '';
  novaLinha[3] = data.respTel || '';
  novaLinha[4] = data.totalMinimo || '';
  novaLinha[5] = data.totalDigitado || '';
  novaLinha[6] = data.familiares || '[]';
  novaLinha[17] = data.vaiOnibus || 'Sim';
  novaLinha[18] = data.respEmail || '';

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

  var eraCadastroNovo = linhaExistente <= 0;

  if (linhaExistente > 0) {
    sheet.getRange(linhaExistente, 1, 1, novaLinha.length).setValues([novaLinha]);
  } else {
    sheet.appendRow(novaLinha);
  }

  if (eraCadastroNovo) {
    Logger.log('Cadastro NOVO detectado (telefone: ' + telefoneNormalizado + '). Disparando e-mail para: ' + data.respEmail);
    enviarEmailConfirmacao(data.respNome, data.respEmail, data.totalDigitado, data.vaiOnibus, parcelasEnviadas);
  } else {
    Logger.log('Cadastro EXISTENTE (edição, telefone: ' + telefoneNormalizado + '). E-mail não enviado por design.');
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

var TEMPLATE_EMAIL_CONFIRMACAO = `
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Inscrição confirmada — CONFRA2026</title>
</head>
<body style="margin:0; padding:0; background-color:#EDE9DD; font-family: Arial, Helvetica, sans-serif;">

<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#EDE9DD;">
<tr>
<td align="center" style="padding: 32px 16px;">

  <table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px; width:100%; background-color:#FFFFFF; border-radius:16px; overflow:hidden; box-shadow:0 4px 16px rgba(0,0,0,0.08);">

    <!-- Header: logo CONFRA2026 -->
    <tr>
      <td align="center" style="background-color:#008E82; padding:34px 24px 30px;">
        <img src="https://rodrigopaguiar.github.io/confra2026/CONFRA2026%20LOGO.png" width="150" alt="CONFRA2026 - Que dia Feliz!" style="display:block; margin:0 auto;">
      </td>
    </tr>

    <!-- Corpo -->
    <tr>
      <td style="padding:36px 32px 8px;">
        <div style="font-size:32px; text-align:center; margin-bottom:14px;">🎉</div>
        <p style="font-family: Arial, Helvetica, sans-serif; font-size:20px; font-weight:bold; color:#008E82; margin:0 0 14px; text-align:center;">
          Olá, Irmão(ã) {{PRIMEIRO_NOME}}! Graça e paz!
        </p>
        <p style="font-size:15px; line-height:1.7; color:#3F3B33; margin:0 0 14px; text-align:center;">
          Sua inscrição para o <strong>CONFRA2026</strong> da IBFO está confirmada! Não perca os avisos nos cultos e continue orando, porque esse será um dia incrível.
        </p>
        <p style="font-family: Arial, Helvetica, sans-serif; font-size:17px; font-weight:bold; color:#D94916; margin:0 0 16px; text-align:center;">
          Que dia feliz! #CONFRA2026
        </p>
        <p style="font-size:13px; font-style:italic; line-height:1.6; color:#5B5644; margin:0 0 28px; text-align:center;">
          "Este é o dia que fez o Senhor; regozijemo-nos e alegremo-nos nele."<br>
          <strong style="color:#D94916; font-style:normal;">Salmos 118:24</strong>
        </p>
      </td>
    </tr>

    <!-- Card de detalhes do evento -->
    <tr>
      <td style="padding:0 32px 8px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#F3F8F7; border:1px solid #CFE9E5; border-radius:12px;">
          <tr>
            <td style="padding:20px 22px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td style="padding:6px 0; font-size:13px; color:#5B5644;">📅&nbsp; <strong>Data</strong></td>
                  <td style="padding:6px 0; font-size:13px; color:#3F3B33; text-align:right;">28/11/2026</td>
                </tr>
                <tr>
                  <td style="padding:6px 0; font-size:13px; color:#5B5644; border-top:1px solid #DCEEEB;">🕖&nbsp; <strong>Horário</strong></td>
                  <td style="padding:6px 0; font-size:13px; color:#3F3B33; text-align:right; border-top:1px solid #DCEEEB;">07:00 às 18:00</td>
                </tr>
                <tr>
                  <td style="padding:6px 0; font-size:13px; color:#5B5644; border-top:1px solid #DCEEEB;">📍&nbsp; <strong>Local</strong></td>
                  <td style="padding:6px 0; font-size:13px; color:#3F3B33; text-align:right; border-top:1px solid #DCEEEB;">Portal do Sol - Mairinque</td>
                </tr>
                <tr>
                  <td style="padding:6px 0; font-size:13px; color:#5B5644; border-top:1px solid #DCEEEB;">💰&nbsp; <strong>Total planejado</strong></td>
                  <td style="padding:6px 0; font-size:13px; color:#3F3B33; text-align:right; border-top:1px solid #DCEEEB;">R$ {{TOTAL_PLANEJADO}}</td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>

        <!-- {{BLOCO_TRANSPORTE}} -->

    <!-- O que levar -->
    <tr>
      <td style="padding:20px 32px 8px;">
        <p style="font-size:13px; font-weight:bold; color:#3F3B33; margin:0 0 6px;">🎒 O que levar no dia:</p>
        <p style="font-size:13px; line-height:1.7; color:#5B5644; margin:0;">
          Roupa confortável, traje de banho apropriado (caso deseje usar a piscina), toalha, protetor solar, repelente e sua Bíblia para o momento devocional.
        </p>
      </td>
    </tr>

    <!-- Historico de parcelas: {{TABELA_PARCELAS}} sera montada pelo Apps Script, uma linha por mes com valor > 0 -->
    <tr>
      <td style="padding:20px 32px 8px;">
        <p style="font-size:13px; font-weight:bold; color:#3F3B33; margin:0 0 10px;">🧾 Como você planejou sua contribuição:</p>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #EDE9DD; border-radius:10px; overflow:hidden;">
          <tr style="background-color:#F7F5EE;">
            <td style="padding:8px 14px; font-size:11px; font-weight:bold; text-transform:uppercase; color:#5B5644;">Mês</td>
            <td style="padding:8px 14px; font-size:11px; font-weight:bold; text-transform:uppercase; color:#5B5644; text-align:right;">Valor</td>
          </tr>
          <!-- {{LINHAS_PARCELAS}} -->
          <tr style="background-color:#F3F8F7;">
            <td style="padding:8px 14px; font-size:12.5px; font-weight:bold; color:#008E82; border-top:1px solid #CFE9E5;">Total</td>
            <td style="padding:8px 14px; font-size:12.5px; font-weight:bold; color:#008E82; text-align:right; border-top:1px solid #CFE9E5;">R$ {{TOTAL_PLANEJADO}}</td>
          </tr>
        </table>
        <p style="font-size:11px; color:#9B9585; margin:8px 0 0;">Mudou de ideia? É só voltar na página com o mesmo telefone para reorganizar as parcelas futuras — o que já foi confirmado como pago fica preservado.</p>
      </td>
    </tr>

    <!-- Botão -->
    <tr>
      <td align="center" style="padding:28px 32px 8px;">
        <table role="presentation" cellpadding="0" cellspacing="0">
          <tr>
            <td style="background-color:#F15A24; border-radius:9px;">
              <a href="https://rodrigopaguiar.github.io/confra2026/" target="_blank" style="display:inline-block; padding:13px 30px; font-family: Arial, Helvetica, sans-serif; font-size:14px; font-weight:bold; color:#FFFFFF; text-decoration:none;">
                Ver a página do evento
              </a>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <tr>
      <td align="center" style="padding:0 32px;">
        <p style="text-align:center; font-size:11px; color:#9B9585; margin:10px 0 0;">
          Precisa ajustar suas parcelas? É só voltar na página com o mesmo telefone que você usou na inscrição.
        </p>
      </td>
    </tr>

    <!-- Rodapé -->
    <tr>
      <td align="center" style="background-color:#008E82; padding:28px 32px 32px;">
        <img src="https://rodrigopaguiar.github.io/confra2026/fotos/ibfo-logo-branco.png" width="70" alt="IBFO" style="display:block; margin:0 auto 12px;">
        <p style="font-size:14px; font-weight:bold; color:#FFFFFF; margin:0 0 4px;">
          Igreja Batista da Família em Osasco
        </p>
        <p style="font-size:11px; color:#CFEEEA; margin:0 0 10px;">
          Av. Sete de Setembro, 763 - Cipava - Osasco - SP
        </p>
        <p style="font-size:10px; color:#A8DDD6; margin:0; line-height:1.5;">
          Você recebeu este e-mail porque se inscreveu no CONFRA2026 através da nossa página oficial.
        </p>
      </td>
    </tr>

  </table>

</td>
</tr>
</table>

</body>
</html>

`;

var BLOCO_TRANSPORTE_ONIBUS = `
    <tr>
      <td style="padding:16px 32px 8px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#FFF4EC; border:1px solid #FBD9C3; border-radius:12px;">
          <tr>
            <td style="padding:16px 20px; text-align:center;">
              <div style="font-size:22px; margin-bottom:4px;">🚌</div>
              <p style="font-size:14px; font-weight:bold; color:#D94916; margin:0 0 4px;">Sua vaga no ônibus fretado está garantida!</p>
              <p style="font-size:12.5px; color:#5B5644; margin:0;">Embarque às 07:00, saindo da IBFO. Chegue com alguns minutos de antecedência.</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
`;

var BLOCO_TRANSPORTE_CARRO = `
    <tr>
      <td style="padding:16px 32px 8px;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#F3F8F7; border:1px solid #CFE9E5; border-radius:12px;">
          <tr>
            <td style="padding:16px 20px; text-align:center;">
              <div style="font-size:22px; margin-bottom:4px;">🚗</div>
              <p style="font-size:14px; font-weight:bold; color:#008E82; margin:0 0 4px;">Você vai com veículo próprio — nos vemos por lá!</p>
              <p style="font-size:12.5px; color:#5B5644; margin:0;">O local tem estacionamento disponível. Chegada prevista a partir das 07:00.</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
`;

var NOMES_MESES_COMPLETO = { Ago: 'Agosto', Set: 'Setembro', Out: 'Outubro', Nov: 'Novembro', Dez: 'Dezembro' };

function substituirTodas(str, busca, valor) {
  return str.split(busca).join(valor);
}

function enviarEmailConfirmacao(nome, email, totalPlanejado, vaiOnibus, parcelas) {
  if (!email) return;

  var primeiroNome = (nome || '').trim().split(' ')[0] || '';
  var totalFormatado = Number(totalPlanejado || 0).toFixed(2).replace('.', ',');

  var blocoTransporte = (vaiOnibus === 'Não') ? BLOCO_TRANSPORTE_CARRO : BLOCO_TRANSPORTE_ONIBUS;

  var linhasParcelas = '';
  (parcelas || []).forEach(function (p) {
    if (p.valor && p.valor > 0) {
      var valorFmt = Number(p.valor).toFixed(2).replace('.', ',');
      var nomeMes = NOMES_MESES_COMPLETO[p.mes] || p.mes;
      linhasParcelas +=
        '<tr><td style="padding:8px 14px; font-size:13px; color:#3F3B33; border-top:1px solid #EDE9DD;">' +
        nomeMes + '/26</td><td style="padding:8px 14px; font-size:13px; color:#3F3B33; text-align:right; border-top:1px solid #EDE9DD;">R$ ' +
        valorFmt + '</td></tr>';
    }
  });

  var html = TEMPLATE_EMAIL_CONFIRMACAO;
  html = substituirTodas(html, '{{PRIMEIRO_NOME}}', primeiroNome);
  html = substituirTodas(html, '{{TOTAL_PLANEJADO}}', totalFormatado);
  html = substituirTodas(html, '<!-- {{BLOCO_TRANSPORTE}} -->', blocoTransporte);
  html = substituirTodas(html, '<!-- {{LINHAS_PARCELAS}} -->', linhasParcelas);

  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Inscrição confirmada — CONFRA2026',
      htmlBody: html
    });
    Logger.log('E-mail de confirmação enviado com sucesso para: ' + email);
  } catch (err) {
    // Se o envio falhar por qualquer motivo, o cadastro já foi gravado normalmente,
    // então não interrompemos o fluxo por causa disso — mas registramos o erro para
    // conseguirmos diagnosticar depois, em Execuções.
    Logger.log('FALHA ao enviar e-mail para ' + email + ': ' + err.message);
  }
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
        respEmail: linhas[i][18] || '',
        parcelas: parcelas
      };
      break;
    }
  }

  var saida = callback + '(' + JSON.stringify(resultado) + ');';
  return ContentService.createTextOutput(saida).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/**
 * FUNÇÃO DE TESTE MANUAL — não é usada pelo formulário, é só para você diagnosticar.
 * Troque o e-mail abaixo pelo seu, selecione "testeEnviarEmail" no menu de funções lá em
 * cima do editor (ao lado do botão "Executar"), e clique em Executar. O resultado (sucesso
 * ou o erro exato) aparece no painel "Registro de execução" na parte de baixo do editor.
 */
function testeEnviarEmail() {
  var emailTeste = 'COLOQUE_SEU_EMAIL_AQUI@gmail.com';
  var parcelasTeste = [
    { mes: 'Ago', valor: 50 },
    { mes: 'Set', valor: 50 }
  ];
  enviarEmailConfirmacao('Rodrigo Teste', emailTeste, 100, 'Sim', parcelasTeste);
  Logger.log('Função de teste concluída — confira as mensagens acima.');
}
