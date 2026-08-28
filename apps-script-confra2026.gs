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
 *    T: Indice Centavo (NAO PREENCHER MANUALMENTE - atribuido automaticamente pelo script,
 *       usado para identificar quem pagou no extrato do Pix: o valor final cobrado da pessoa
 *       é o valor da parcela + R$0,26 + esse indice em centavos, ex: indice 0 = ,26; indice 1 = ,27)
 *    U: Quitação Confirmada (PREENCHIMENTO MANUAL E DELIBERADO PELA SECRETARIA — só escreva
 *       algo aqui, tipo "OK" ou a data, depois de conferir a linha inteira e decidir que
 *       aquela família realmente concluiu as parcelas. Esse campo existe separado dos
 *       "Confirmado <Mês>" de propósito: evita que um clique errado num mês dispare
 *       sozinho o email de conclusão — só dispara quando ISSO aqui E todos os meses
 *       batem ao mesmo tempo.)
 *    V: Email Conclusao Enviado (NAO PREENCHER MANUALMENTE - marcado com a data/hora quando
 *       o email de "contribuição concluída" já foi disparado para essa família, pra nunca
 *       reenviar duas vezes)
 *
 * IDENTIFICACAO DO PIX POR CENTAVO: cada cadastro NOVO recebe um indice sequencial unico
 * (0, 1, 2...) na coluna T, atribuido de forma segura mesmo se dois cadastros chegarem ao
 * mesmo tempo (via LockService, em obterProximoIndiceCentavo()). Esse indice e somado a uma
 * base fixa de 26 centavos (BASE_CENTAVOS_PIX, o ano da campanha) para gerar o valor final
 * que a landing page usa ao montar o QR Code Pix de cada parcela — assim, quando a Tesouraria
 * olhar o extrato, cada valor "quebrado" (ex: R$ 50,26 / R$ 50,27) aponta para uma familia
 * especifica, sem depender so do nome do remetente. O indice e atribuido UMA UNICA VEZ por
 * cadastro (nunca muda em edicoes futuras), garantindo que o mesmo valor sempre identifique
 * a mesma familia ao longo de toda a campanha.
 *
 * EMAIL DE CONCLUSAO DAS CONTRIBUICOES: dispara quando DUAS coisas sao verdade ao mesmo
 * tempo — (1) a Secretaria escreveu algo na coluna U ("Quitação Confirmada"), de forma
 * deliberada, depois de revisar a linha, e (2) o script confere sozinho que todos os meses
 * com valor > 0 ja estao de fato marcados como confirmados. Essa dupla checagem existe pra
 * que um erro de digitação num unico mês nao dispare o email sozinho — precisa das duas
 * confirmações batendo juntas. Isso funciona via um gatilho de EDICAO NA PLANILHA (nao no
 * formulario), entao exige UM PASSO DE CONFIGURACAO MANUAL, UMA UNICA VEZ (o Apps Script
 * nao pode se autoconceder essa permissao sozinho):
 *   a) No editor do Apps Script, no menu lateral esquerdo, clique no icone de relogio
 *      ("Acionadores" / "Triggers").
 *   b) Clique em "+ Adicionar acionador" (canto inferior direito).
 *   c) Configure: Função a ser executada = aoEditarPlanilha | Evento = Do tipo de planilha
 *      | Tipo de evento = Ao editar. Salve e autorize as permissoes pedidas (sua conta Google).
 * Sem esse passo, a coluna T e o cadastro continuam funcionando normalmente — só o email de
 * conclusão não vai disparar.
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
 *
 * LISTA PARA O ÔNIBUS: na própria planilha, use o menu "CONFRA2026 > Gerar lista para o
 * ônibus (WhatsApp)" (aparece sozinho ao abrir a planilha — recarregue a aba se não aparecer
 * logo após colar este código pela primeira vez). Gera a lista de todo mundo que marcou
 * "Vai de Onibus = Sim", já formatada pra colar direto numa conversa do WhatsApp.
 */

var MESES_COLS = {
  Ago: { valor: 8,  confirmado: 9  },
  Set: { valor: 10, confirmado: 11 },
  Out: { valor: 12, confirmado: 13 },
  Nov: { valor: 14, confirmado: 15 },
  Dez: { valor: 16, confirmado: 17 }
};

var COL_INDICE_CENTAVO = 20; // Coluna T
var COL_QUITACAO_CONFIRMADA = 21; // Coluna U (manual, preenchida pela Secretaria)
var COL_EMAIL_CONCLUSAO_ENVIADO = 22; // Coluna V (automatica)
var BASE_CENTAVOS_PIX = 26; // Ano da campanha (2026) - usado tambem no index.html

function estaConfirmado(valorCelula) {
  return valorCelula !== '' && valorCelula !== null && valorCelula !== undefined;
}

/**
 * MENU PERSONALIZADO NA PLANILHA — "CONFRA2026 > Gerar lista para o ônibus (WhatsApp)".
 * Aparece sozinho quando a planilha é aberta (pode ser preciso recarregar a aba depois
 * de colar este código pela primeira vez). Gera uma lista de todos os inscritos que
 * marcaram "Vai de Onibus" = Sim — um por linha, responsável e cada dependente — com
 * nome, documento e, só para dependentes, o nome do responsável. Termina com o total
 * de assentos necessários. Sai em texto monoespaçado (envolto em ``` ```), formato que
 * o WhatsApp reconhece e exibe com fonte de largura fixa, preservando o alinhamento
 * das colunas ao colar direto numa conversa.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('CONFRA2026')
    .addItem('Gerar lista para o ônibus (WhatsApp)', 'mostrarListaOnibus')
    .addToUi();
}

function montarListaOnibus() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var linhas = sheet.getDataRange().getValues();
  var pessoas = []; // { nome, documento, responsavel }

  for (var i = 1; i < linhas.length; i++) {
    var linha = linhas[i];
    if (!linha[1]) continue; // linha vazia, sem Nome Responsavel

    var vaiDeOnibus = String(linha[17] || '').trim().toLowerCase() === 'sim';
    if (!vaiDeOnibus) continue;

    pessoas.push({ nome: String(linha[1] || ''), documento: String(linha[2] || '—'), responsavel: '' });

    var familiares = [];
    try { familiares = JSON.parse(linha[6] || '[]'); } catch (err) { familiares = []; }

    familiares.forEach(function (f) {
      pessoas.push({
        nome: String((f && f.nome) || ''),
        documento: String((f && f.rg) || '—'),
        responsavel: String(linha[1] || '')
      });
    });
  }

  return pessoas;
}

function formatarListaOnibusParaWhatsapp(pessoas) {
  var tituloNome = 'NOME';
  var tituloDoc = 'DOCUMENTO';
  var tituloResp = 'RESPONSÁVEL (se dependente)';

  var larguraNome = tituloNome.length;
  var larguraDoc = tituloDoc.length;
  pessoas.forEach(function (p) {
    if (p.nome.length > larguraNome) larguraNome = p.nome.length;
    if (p.documento.length > larguraDoc) larguraDoc = p.documento.length;
  });

  function pad(str, largura) {
    str = String(str);
    while (str.length < largura) str += ' ';
    return str;
  }

  var linhasTexto = [];
  linhasTexto.push(pad(tituloNome, larguraNome) + '  ' + pad(tituloDoc, larguraDoc) + '  ' + tituloResp);
  linhasTexto.push(pad('', larguraNome).replace(/ /g, '-') + '  ' + pad('', larguraDoc).replace(/ /g, '-') + '  ' + '---------------------------');

  pessoas.forEach(function (p) {
    linhasTexto.push(pad(p.nome, larguraNome) + '  ' + pad(p.documento, larguraDoc) + '  ' + p.responsavel);
  });

  var texto = '```\n' + linhasTexto.join('\n') + '\n```';
  texto += '\n\n*Total de assentos necessários: ' + pessoas.length + '*';
  return texto;
}

function mostrarListaOnibus() {
  var pessoas = montarListaOnibus();

  if (pessoas.length === 0) {
    SpreadsheetApp.getUi().alert('Nenhum inscrito marcou "Vai de Ônibus = Sim" até agora.');
    return;
  }

  var texto = formatarListaOnibusParaWhatsapp(pessoas);
  var textoEscapadoParaHtml = texto.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  var html =
    '<textarea id="txt" style="width:100%; height:360px; font-family:monospace; font-size:12px; box-sizing:border-box;" readonly>' +
    textoEscapadoParaHtml +
    '</textarea>' +
    '<div style="margin-top:10px; font-family:Arial, sans-serif;">' +
    '<button onclick="document.getElementById(\'txt\').select(); document.execCommand(\'copy\');" style="padding:8px 16px; font-weight:bold;">Copiar tudo</button>' +
    '<p style="font-size:12px; color:#666; margin-top:10px;">Se o botão não copiar automaticamente (alguns navegadores bloqueiam), clique dentro do texto, aperte Ctrl+A (ou Cmd+A no Mac) e depois Ctrl+C (ou Cmd+C) para copiar manualmente.<br>' +
    'Cole direto numa conversa do WhatsApp — os três crases (```) no início e no fim fazem o WhatsApp exibir em fonte de largura fixa, mantendo as colunas alinhadas.</p>' +
    '</div>';

  var output = HtmlService.createHtmlOutput(html).setWidth(700).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(output, 'Lista para o Ônibus — ' + pessoas.length + ' assento(s)');
}

/**
 * Retorna o proximo indice sequencial disponivel (0, 1, 2...) para identificacao via
 * centavo do Pix, de forma segura contra dois cadastros simultaneos (LockService).
 * Guardado em PropertiesService (nao numa celula), entao nao depende de nenhuma linha
 * especifica da planilha e nao corre risco de leitura suja entre execucoes concorrentes.
 */
function obterProximoIndiceCentavo() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // espera ate 30s se outro cadastro estiver sendo gravado ao mesmo tempo
  try {
    var props = PropertiesService.getScriptProperties();
    var atual = Number(props.getProperty('ULTIMO_INDICE_CENTAVO') || '-1');
    var proximo = atual + 1;
    props.setProperty('ULTIMO_INDICE_CENTAVO', String(proximo));
    return proximo;
  } finally {
    lock.releaseLock();
  }
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

  var novaLinha = new Array(20);
  novaLinha[0] = new Date();
  novaLinha[1] = data.respNome || '';
  novaLinha[2] = data.respRG || '';
  novaLinha[3] = data.respTel || '';
  novaLinha[4] = data.totalMinimo || '';
  novaLinha[5] = data.totalDigitado || '';
  novaLinha[6] = data.familiares || '[]';
  novaLinha[17] = data.vaiOnibus || 'Sim';
  novaLinha[18] = data.respEmail || '';

  // Indice de identificacao via centavo: atribuido UMA VEZ por cadastro e preservado em
  // qualquer edicao futura, para o valor continuar apontando pra mesma familia sempre.
  var indiceCentavo;
  if (linhaExistente > 0) {
    var indiceExistente = sheet.getRange(linhaExistente, COL_INDICE_CENTAVO).getValue();
    var jaTemIndiceValido = indiceExistente !== '' && indiceExistente !== null && !isNaN(Number(indiceExistente));
    indiceCentavo = jaTemIndiceValido ? Number(indiceExistente) : obterProximoIndiceCentavo();
  } else {
    indiceCentavo = obterProximoIndiceCentavo();
  }
  novaLinha[COL_INDICE_CENTAVO - 1] = indiceCentavo;

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

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok', indiceCentavo: indiceCentavo }))
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

/**
 * GATILHO DE EDICAO NA PLANILHA — precisa ser configurado manualmente como acionador
 * instalável (veja instruções no topo do arquivo). Monitora tanto as colunas "Confirmado
 * <Mês>" quanto a coluna U ("Quitação Confirmada"), já que o email só deve sair quando as
 * duas condições baterem juntas — a edição que completa essa condição pode vir de qualquer
 * uma das duas frentes.
 */
function aoEditarPlanilha(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();

    var colunasMonitoradas = Object.keys(MESES_COLS).map(function (m) { return MESES_COLS[m].confirmado; });
    colunasMonitoradas.push(COL_QUITACAO_CONFIRMADA);

    var colInicio = e.range.getColumn();
    var colFim = colInicio + e.range.getNumColumns() - 1;
    var tocouColunaRelevante = colunasMonitoradas.some(function (c) { return c >= colInicio && c <= colFim; });
    if (!tocouColunaRelevante) return;

    var linhaInicio = Math.max(e.range.getRow(), 2); // pula o cabeçalho
    var linhaFim = e.range.getRow() + e.range.getNumRows() - 1;

    for (var row = linhaInicio; row <= linhaFim; row++) {
      verificarConclusaoEEnviarEmail(sheet, row);
    }
  } catch (err) {
    Logger.log('Erro em aoEditarPlanilha: ' + err.message);
  }
}

/**
 * Só envia o email de conclusão quando DUAS condições baterem ao mesmo tempo:
 * (1) a Secretaria já escreveu algo na coluna U (Quitação Confirmada) — confirmação
 *     deliberada, feita depois de revisar a linha inteira — e
 * (2) todas as parcelas com valor > 0 da linha estão de fato marcadas como confirmadas.
 * Se sim, e o email ainda não foi enviado para essa família, dispara o email e marca
 * a coluna V com a data/hora do envio (para nunca reenviar duas vezes).
 */
function verificarConclusaoEEnviarEmail(sheet, row) {
  var jaEnviado = sheet.getRange(row, COL_EMAIL_CONCLUSAO_ENVIADO).getValue();
  if (jaEnviado) return;

  var quitacaoConfirmadaPelaSecretaria = sheet.getRange(row, COL_QUITACAO_CONFIRMADA).getValue();
  if (!estaConfirmado(quitacaoConfirmadaPelaSecretaria)) return; // ainda nao foi revisada/confirmada manualmente

  var ultimaColuna = Math.max(COL_EMAIL_CONCLUSAO_ENVIADO, sheet.getLastColumn());
  var linha = sheet.getRange(row, 1, 1, ultimaColuna).getValues()[0];

  var teveAlgumaParcela = false;
  var todasConfirmadas = true;
  for (var mes in MESES_COLS) {
    var col = MESES_COLS[mes];
    var valor = Number(linha[col.valor - 1]) || 0;
    if (valor > 0) {
      teveAlgumaParcela = true;
      if (!estaConfirmado(linha[col.confirmado - 1])) {
        todasConfirmadas = false;
        break;
      }
    }
  }

  if (teveAlgumaParcela && todasConfirmadas) {
    var nome = linha[1];
    var email = linha[18];
    var total = linha[5];
    var enviouComSucesso = enviarEmailConclusao(nome, email, total);
    if (enviouComSucesso) {
      sheet.getRange(row, COL_EMAIL_CONCLUSAO_ENVIADO).setValue(new Date());
    } else {
      // Não marca a coluna V: assim, na próxima edição relevante dessa linha
      // (ex: a Secretaria mexer em qualquer célula monitorada), o script tenta
      // enviar de novo automaticamente, em vez de ficar "travado" achando que
      // já enviou quando na verdade falhou (email vazio ou erro no envio).
      Logger.log('Email de conclusão NÃO enviado para a linha ' + row + ' (email vazio ou falha no envio) — coluna V deixada em branco de propósito.');
    }
  }
}

var TEMPLATE_EMAIL_CONCLUSAO = `
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Contribuição concluída — CONFRA2026</title>
</head>
<body style="margin:0; padding:0; background-color:#EDE9DD; font-family: Arial, Helvetica, sans-serif;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#EDE9DD;">
<tr>
<td align="center" style="padding: 32px 16px;">
  <table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px; width:100%; background-color:#FFFFFF; border-radius:16px; overflow:hidden; box-shadow:0 4px 16px rgba(0,0,0,0.08);">
    <tr>
      <td align="center" style="background-color:#008E82; padding:34px 24px 30px;">
        <img src="https://rodrigopaguiar.github.io/confra2026/CONFRA2026%20LOGO.png" width="150" alt="CONFRA2026 - Que dia Feliz!" style="display:block; margin:0 auto;">
      </td>
    </tr>
    <tr>
      <td style="padding:36px 32px 8px;">
        <div style="font-size:32px; text-align:center; margin-bottom:14px;">✅</div>
        <p style="font-size:20px; font-weight:bold; color:#008E82; margin:0 0 14px; text-align:center;">
          Olá, Irmão(ã) {{PRIMEIRO_NOME}}! Graça e paz!
        </p>
        <p style="font-size:15px; line-height:1.7; color:#3F3B33; margin:0 0 14px; text-align:center;">
          Boa notícia: identificamos que todas as suas parcelas do <strong>CONFRA2026</strong> foram recebidas — sua contribuição de <strong>R$ {{TOTAL_PLANEJADO}}</strong> está completa. Muito obrigado por caminhar junto com a gente até aqui!
        </p>
        <p style="font-size:17px; font-weight:bold; color:#D94916; margin:0 0 8px; text-align:center;">
          Nos vemos em 28/11! #CONFRA2026
        </p>
      </td>
    </tr>
    <tr>
      <td align="center" style="background-color:#008E82; padding:28px 32px 32px;">
        <img src="https://rodrigopaguiar.github.io/confra2026/fotos/ibfo-logo-branco.png" width="70" alt="IBFO" style="display:block; margin:0 auto 12px;">
        <p style="font-size:14px; font-weight:bold; color:#FFFFFF; margin:0 0 4px;">
          Igreja Batista da Família em Osasco
        </p>
        <p style="font-size:11px; color:#CFEEEA; margin:0;">
          Av. Sete de Setembro, 763 - Cipava - Osasco - SP
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

/**
 * Retorna true se o email foi de fato disparado com sucesso, false caso contrário
 * (email vazio na planilha, ou qualquer erro no envio). O chamador usa esse retorno
 * pra decidir se marca a coluna V — nunca marca "enviado" sem confirmação real.
 */
function enviarEmailConclusao(nome, email, totalPlanejado) {
  if (!email) {
    Logger.log('Email de conclusão não enviado: coluna Email está vazia para "' + nome + '".');
    return false;
  }

  var primeiroNome = (nome || '').trim().split(' ')[0] || '';
  var totalFormatado = Number(totalPlanejado || 0).toFixed(2).replace('.', ',');

  var html = TEMPLATE_EMAIL_CONCLUSAO;
  html = substituirTodas(html, '{{PRIMEIRO_NOME}}', primeiroNome);
  html = substituirTodas(html, '{{TOTAL_PLANEJADO}}', totalFormatado);

  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Contribuição concluída — CONFRA2026 🎉',
      htmlBody: html
    });
    Logger.log('E-mail de conclusão enviado com sucesso para: ' + email);
    return true;
  } catch (err) {
    Logger.log('FALHA ao enviar e-mail de conclusão para ' + email + ': ' + err.message);
    return false;
  }
}

/**
 * SUBSTITUA a função doGet(e) existente por esta versão abaixo, e adicione a nova
 * função gerarAgregadosDashboard() em algum lugar do mesmo arquivo (não precisa
 * mexer em mais nada — doPost, buscarPorTelefone, envio de e-mail, etc. continuam
 * exatamente iguais).
 *
 * O que muda: o doGet passa a reconhecer um terceiro caso, "?dashboard=1", além dos
 * dois que já existiam (busca por telefone e a resposta padrão "O script está no ar").
 *
 * URL a colocar no dashboard.html: a MESMA URL que a landing page já usa
 * (APPS_SCRIPT_URL) + "?dashboard=1" no final.
 */

const META_FINANCEIRA = 11500;
const CAPACIDADE_ONIBUS = 45;
const ANO_EVENTO = 2026;
const MES_NUMERO = { Ago: 8, Set: 9, Out: 10, Nov: 11, Dez: 12 };

function doGet(e) {
  if (e.parameter.telefone) {
    return buscarPorTelefone(e);
  }
  if (e.parameter.dashboard) {
    return gerarAgregadosDashboard();
  }
  return ContentService.createTextOutput('O script está no ar. Use POST para enviar dados.');
}

function gerarAgregadosDashboard() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var linhas = sheet.getDataRange().getValues();

  var totalPessoas = 0, familias = 0;
  var criancasGratis = 0, meia = 0, inteira = 0;
  var onibusConfirmados = 0, carroProprio = 0;
  var confirmadoTotal = 0, aConfirmarTotal = 0;
  var parcelasAtrasadasQtd = 0, parcelasAtrasadasValor = 0;

  var porMes = {};
  for (var m in MESES_COLS) porMes[m] = { mes: m, confirmado: 0, pendente: 0 };

  var now = new Date();
  var mesAtual = now.getMonth() + 1;
  var anoAtual = now.getFullYear();

  for (var i = 1; i < linhas.length; i++) {
    var linha = linhas[i];
    if (!linha[1]) continue; // sem Nome Responsavel, pula linha vazia
    familias++;

    var familiares = [];
    try { familiares = JSON.parse(linha[6] || '[]'); } catch (err) { familiares = []; }

    totalPessoas += 1 + familiares.length;
    inteira += 1; // responsável = adulto

    familiares.forEach(function (f) {
      var v = String(f.faixaValor);
      if (v === '0') criancasGratis++;
      else if (v === '72') meia++;
      else inteira++;
    });

    var vaiDeOnibus = String(linha[17] || '').trim().toLowerCase() === 'sim';
    var pessoasFamilia = 1 + familiares.length;
    if (vaiDeOnibus) onibusConfirmados += pessoasFamilia;
    else carroProprio += pessoasFamilia;

    for (var mes in MESES_COLS) {
      var col = MESES_COLS[mes];
      var valor = Number(linha[col.valor - 1]) || 0;
      var confirmado = estaConfirmado(linha[col.confirmado - 1]);

      if (valor > 0) {
        if (confirmado) {
          porMes[mes].confirmado += valor;
          confirmadoTotal += valor;
        } else {
          porMes[mes].pendente += valor;
          aConfirmarTotal += valor;
          var mesNum = MES_NUMERO[mes];
          var jaPassouOuEhAtual = anoAtual > ANO_EVENTO || (anoAtual === ANO_EVENTO && mesNum <= mesAtual);
          if (jaPassouOuEhAtual) {
            parcelasAtrasadasQtd++;
            parcelasAtrasadasValor += valor;
          }
        }
      }
    }
  }

  var resultado = {
    atualizado_em: Utilities.formatDate(now, 'America/Sao_Paulo', 'yyyy-MM-dd'),
    prazos: { inscricao: '2026-11-21', evento: '2026-11-28' },
    pessoas: {
      total_inscritos: totalPessoas,
      familias: familias,
      criancas_gratis: criancasGratis,
      meia_contribuicao: meia,
      contribuicao_inteira: inteira
    },
    onibus: {
      capacidade: CAPACIDADE_ONIBUS,
      confirmados: onibusConfirmados,
      carro_proprio: carroProprio
    },
    financeiro: {
      meta: META_FINANCEIRA,
      confirmado_total: confirmadoTotal,
      a_confirmar_total: aConfirmarTotal,
      por_mes: Object.keys(porMes).map(function (k) { return porMes[k]; }),
      parcelas_atrasadas_qtd: parcelasAtrasadasQtd,
      parcelas_atrasadas_valor: parcelasAtrasadasValor
    },
    apoio_nominal: {
      familias_cobertas: 0,
      sem_cobertura_abaixo_minimo: 0,
      obs: 'campo ainda nao capturado no formulario'
    }
  };

  return ContentService.createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
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
      var indiceCentavoCelula = linhas[i][COL_INDICE_CENTAVO - 1];
      resultado = {
        encontrado: true,
        respNome: linhas[i][1] || '',
        respRG: linhas[i][2] || '',
        familiares: linhas[i][6] || '[]',
        vaiOnibus: linhas[i][17] || 'Sim',
        respEmail: linhas[i][18] || '',
        parcelas: parcelas,
        indiceCentavo: (indiceCentavoCelula === '' || indiceCentavoCelula === null || isNaN(Number(indiceCentavoCelula)))
          ? null
          : Number(indiceCentavoCelula)
      };
      break;
    }
  }

  var saida = callback + '(' + JSON.stringify(resultado) + ');';
  return ContentService.createTextOutput(saida).setMimeType(ContentService.MimeType.JAVASCRIPT);
}
