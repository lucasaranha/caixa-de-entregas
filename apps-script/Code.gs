/**
 * Caixa de Entregas · servidor
 * A página fica no GitHub Pages. Este script recebe os arquivos e guarda tudo
 * no seu Google Drive: uma pasta por série e uma planilha com cada entrega.
 *
 * Primeira vez:
 *   1. No topo do editor, escolha a função "configurar" e clique em Executar.
 *      Autorize o acesso ao Drive e às Planilhas quando o Google pedir.
 *   2. Implantar > Nova implantação > App da Web
 *      Executar como: Eu · Quem pode acessar: Qualquer pessoa
 *   3. Copie o link que termina em /exec e cole em config.js, no repositório.
 */

const FUSO = 'America/Sao_Paulo';
const NOME_PASTA = 'Caixa de Entregas';
const NOME_PLANILHA = 'Caixa de Entregas · registro';
const NOME_ABA = 'Entregas';
const CABECALHO = ['Entregue em', 'Protocolo', 'Aluno(a)', 'Série', 'Arquivos', 'Corrigida'];
const LIMITE_MB = 10;
const MAX_ARQUIVOS = 5;

// As chaves precisam ser iguais aos "value" da lista de séries no Index.html.
const SERIES = {
  '1ano': '1º ano',
  '2ano': '2º ano'
};

const TIPOS = {
  pdf: 'application/pdf',
  doc: 'application/msword',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  odt: 'application/vnd.oasis.opendocument.text',
  rtf: 'application/rtf',
  txt: 'text/plain',
  ppt: 'application/vnd.ms-powerpoint',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  odp: 'application/vnd.oasis.opendocument.presentation',
  xls: 'application/vnd.ms-excel',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  png: 'image/png',
  heic: 'image/heic',
  heif: 'image/heif',
  webp: 'image/webp',
  gif: 'image/gif'
};

/** Abrir o link /exec no navegador mostra se o servidor está no ar. */
function doGet() {
  return json_({ ok: true, servico: 'Caixa de Entregas' });
}

/** A página manda cada pedido como texto JSON: { acao, dados }. */
function doPost(e) {
  let resposta;
  try {
    const pedido = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (pedido.acao === 'enviarArquivo') {
      resposta = { ok: true, dados: enviarArquivo_(pedido.dados || {}) };
    } else if (pedido.acao === 'registrarEntrega') {
      resposta = { ok: true, dados: registrarEntrega_(pedido.dados || {}) };
    } else {
      falha_('Pedido desconhecido. Recarregue a página e tente de novo.');
    }
  } catch (err) {
    const m = String((err && err.message) || err);
    const i = m.indexOf('CE|');
    if (i < 0) console.error(err);
    resposta = { ok: false, erro: i > -1 ? m.slice(i + 3) : null };
  }
  return json_(resposta);
}

/**
 * Rode pelo editor na primeira vez: cria a pasta, as subpastas por série e a planilha.
 * Rodar de novo é seguro: nada é duplicado, e os nomes das pastas, da planilha e
 * do cabeçalho voltam a ser os definidos aqui (útil depois de mudar as séries).
 */
function configurar() {
  const raiz = pastaRaiz_();
  if (raiz.getName() !== NOME_PASTA) raiz.setName(NOME_PASTA);
  Object.keys(SERIES).forEach(function (chave) {
    const pasta = pastaDaSerie_(chave);
    if (pasta.getName() !== SERIES[chave]) pasta.setName(SERIES[chave]);
  });
  const planilha = planilha_();
  if (planilha.getName() !== NOME_PLANILHA) planilha.rename(NOME_PLANILHA);
  abaEntregas_().getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO]);
  console.log('Pasta das entregas: ' + raiz.getUrl());
  console.log('Planilha de registro: ' + planilha.getUrl());
}

/** Recebe um arquivo por vez. */
function enviarArquivo_(envio) {
  const dados = validar_(envio);
  const original = limparArquivo_(envio.arquivo);
  const ext = extensao_(original);
  if (!TIPOS[ext]) {
    falha_('“' + original + '” não é um formato aceito. Envie PDF, Word, PowerPoint ou imagem.');
  }
  const bytes = Utilities.base64Decode(String(envio.conteudo || ''));
  if (!bytes.length) falha_('“' + original + '” está vazio. Escolha o arquivo de novo.');
  if (bytes.length > LIMITE_MB * 1024 * 1024) {
    falha_('“' + original + '” passa de ' + LIMITE_MB + ' MB. Envie um arquivo menor.');
  }

  const quando = Utilities.formatDate(new Date(), FUSO, "yyyy-MM-dd HH'h'mm");
  const titulo = quando + ' · ' + dados.nome + ' · ' + original;
  const pasta = comTrava_(function () { return pastaDaSerie_(dados.serie); });
  const arquivo = pasta.createFile(Utilities.newBlob(bytes, TIPOS[ext], titulo));
  arquivo.setDescription(marca_(dados.protocolo));
  return { id: arquivo.getId() };
}

/** Depois que todos os arquivos chegam, anota a entrega na planilha. */
function registrarEntrega_(entrega) {
  const dados = validar_(entrega);
  const ids = Array.isArray(entrega.ids) ? entrega.ids.slice(0, MAX_ARQUIVOS) : [];
  if (!ids.length) falha_('Nenhum arquivo chegou. Tente entregar de novo.');

  const arquivos = ids.map(function (id) {
    const f = DriveApp.getFileById(String(id));
    if (f.getDescription() !== marca_(dados.protocolo)) {
      falha_('Um dos arquivos não confere com esta entrega. Tente entregar de novo.');
    }
    const titulo = f.getName();
    return { nome: titulo.slice(titulo.lastIndexOf(' · ') + 3), url: f.getUrl() };
  });

  comTrava_(function () {
    const aba = abaEntregas_();
    aba.appendRow([new Date(), dados.protocolo, dados.nome, SERIES[dados.serie], '', false]);
    const linha = aba.getLastRow();

    const texto = arquivos.map(function (a) { return a.nome; }).join('\n');
    const rico = SpreadsheetApp.newRichTextValue().setText(texto);
    let pos = 0;
    arquivos.forEach(function (a) {
      rico.setLinkUrl(pos, pos + a.nome.length, a.url);
      pos += a.nome.length + 1;
    });
    aba.getRange(linha, 5).setRichTextValue(rico.build());
    aba.getRange(linha, 6).insertCheckboxes();
  });

  return { protocolo: dados.protocolo };
}

/* ---------- internos ---------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function validar_(x) {
  x = x || {};
  const nome = String(x.nome || '').replace(/[·\s]+/g, ' ').trim().slice(0, 80);
  if (nome.split(' ').length < 2) falha_('Escreva nome e sobrenome, como aparece na chamada.');
  const serie = String(x.serie || '');
  if (!Object.prototype.hasOwnProperty.call(SERIES, serie)) falha_('Escolha sua série na lista.');
  const protocolo = String(x.protocolo || '');
  if (!/^ENT-\d{6}-[A-Z0-9]{4}$/.test(protocolo)) falha_('Recarregue a página e tente entregar de novo.');
  return { nome: nome, serie: serie, protocolo: protocolo };
}

// O prefixo "CE|" avisa a página que a mensagem pode ser mostrada ao aluno.
function falha_(mensagem) {
  throw new Error('CE|' + mensagem);
}

function marca_(protocolo) {
  return 'Caixa de Entregas · protocolo ' + protocolo;
}

function extensao_(nome) {
  const i = nome.lastIndexOf('.');
  return i > -1 ? nome.slice(i + 1).toLowerCase() : '';
}

function limparArquivo_(nome) {
  const limpo = String(nome || '').replace(/[\\/]/g, '-').replace(/\s+/g, ' ').trim();
  return limpo.length > 120 ? limpo.slice(-120) : limpo;
}

function comTrava_(fn) {
  const trava = LockService.getScriptLock();
  trava.waitLock(20000);
  try {
    return fn();
  } finally {
    trava.releaseLock();
  }
}

function pastaRaiz_() {
  const props = PropertiesService.getScriptProperties();
  const existente = pastaPorId_(props.getProperty('PASTA_RAIZ'));
  if (existente) return existente;
  const nova = DriveApp.createFolder(NOME_PASTA);
  props.setProperty('PASTA_RAIZ', nova.getId());
  return nova;
}

function pastaDaSerie_(chave) {
  const props = PropertiesService.getScriptProperties();
  const existente = pastaPorId_(props.getProperty('PASTA_' + chave));
  if (existente) return existente;
  const nova = pastaRaiz_().createFolder(SERIES[chave]);
  props.setProperty('PASTA_' + chave, nova.getId());
  return nova;
}

function pastaPorId_(id) {
  if (!id) return null;
  try {
    const pasta = DriveApp.getFolderById(id);
    return pasta.isTrashed() ? null : pasta;
  } catch (e) {
    return null;
  }
}

function planilha_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PLANILHA');
  if (id) {
    try {
      if (!DriveApp.getFileById(id).isTrashed()) return SpreadsheetApp.openById(id);
    } catch (e) {
      // a planilha foi apagada: cria outra abaixo
    }
  }
  const ss = SpreadsheetApp.create(NOME_PLANILHA);
  ss.setSpreadsheetTimeZone(FUSO);
  const aba = ss.getSheets()[0];
  aba.setName(NOME_ABA);
  aba.appendRow(CABECALHO);
  aba.setFrozenRows(1);
  aba.getRange(1, 1, 1, CABECALHO.length).setFontWeight('bold');
  aba.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm');
  aba.getRange('E:E').setWrap(true);
  aba.setColumnWidth(1, 140);
  aba.setColumnWidth(2, 150);
  aba.setColumnWidth(3, 220);
  aba.setColumnWidth(4, 120);
  aba.setColumnWidth(5, 320);
  aba.setColumnWidth(6, 90);
  DriveApp.getFileById(ss.getId()).moveTo(pastaRaiz_());
  props.setProperty('PLANILHA', ss.getId());
  return ss;
}

function abaEntregas_() {
  const ss = planilha_();
  return ss.getSheetByName(NOME_ABA) || ss.getSheets()[0];
}
