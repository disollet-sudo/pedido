var SHEET_ID  = '1gKLHEX1B9tKsrNC87PGrQgwNuMnm4ww4dQtteuYtykg';
var FOLDER_ID = '1RUUqIUmk2b8-n6d1lC5RYunI5v4Z9BrT';

function testarPermissoes() {
  console.log('--- Iniciando teste de permissões ---');
  try { var quota = MailApp.getRemainingDailyQuota(); console.log('MailApp OK. Cota: ' + quota); } catch (e) { console.log('ERRO MailApp: ' + e.message); }
  try { var ss = SpreadsheetApp.openById(SHEET_ID); console.log('SpreadsheetApp OK. Planilha: ' + ss.getName()); } catch (e) { console.log('ERRO SpreadsheetApp: ' + e.message); }
  try { var folder = DriveApp.getFolderById(FOLDER_ID); console.log('DriveApp OK. Pasta: ' + folder.getName()); } catch (e) { console.log('ERRO DriveApp: ' + e.message); }
  console.log('--- Teste finalizado ---');
}

function doGet(e) {
  var forcarAtualizacao = (e && e.parameter && e.parameter.atualizar === 'true');
  var produtos = getProdutos(forcarAtualizacao);
  return ContentService.createTextOutput(JSON.stringify(produtos)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    var cache = CacheService.getScriptCache();

    // Salvar Novo Cliente com a nova estrutura estendida (10 colunas)
    if (dados.acao === 'salvar_cliente') {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var abaClientes = ss.getSheetByName('Clientes');
      if (!abaClientes) {
        abaClientes = ss.insertSheet('Clientes');
        abaClientes.appendRow(['CNPJ', 'RAZÃO SOCIAL', 'NOME FANTASIA', 'TELEFONE', 'ENDEREÇO', 'ESTADO', 'BAIRRO', 'MUNICIPIO', 'NUMERO', 'CEP']);
      }
      
      var c = dados.cliente;
      var dadosExistentes = abaClientes.getDataRange().getValues();
      var existe = false;
      
      // Validação de segurança no Backend contra CNPJ duplicado
      var cnpjLimpoNovo = String(c.cnpj).replace(/\D/g, '').trim();
      for (var i = 1; i < dadosExistentes.length; i++) {
        var cnpjLimpoExistente = String(dadosExistentes[i][0]).replace(/\D/g, '').trim();
        if (cnpjLimpoExistente === cnpjLimpoNovo && cnpjLimpoNovo !== '') {
          existe = true; 
          break;
        }
      }
      
      if (existe) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Bloqueado! Este CNPJ já possui cadastro na base.' })).setMimeType(ContentService.MimeType.JSON);
      } else {
        abaClientes.appendRow([c.cnpj, c.razao, c.fantasia, c.telefone, c.endereco, c.estado, c.bairro, c.municipio, c.numero, c.cep]);
        // Invalida o cache de clientes para forçar reatualização na próxima abertura
        cache.remove('cat_clientes_info');
        return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'Cliente cadastrado com sucesso!' })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Upload de Pedido Manual (PDF)
    if (dados.acao === 'upload_pdf_manual') {
      var FOLDER_PEDIDOS_ID = '1Jdxb2j_dOWTJtgABUNVOcM6kYLVkmPTH';
      var dataBytes = Utilities.base64Decode(dados.fileBase64.split(',')[1]);
      var blob = Utilities.newBlob(dataBytes, dados.fileMimeType, dados.fileName);
      var file = DriveApp.getFolderById(FOLDER_PEDIDOS_ID).createFile(blob);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', fileId: file.getId() })).setMimeType(ContentService.MimeType.JSON);
    }

    if (dados.acao === 'pdf') {
      var result = gerarPdfPedido(dados.dadosPdf);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', base64: result.base64, emailStatus: result.emailStatus, nomeArquivo: result.nomeArquivo })).setMimeType(ContentService.MimeType.JSON);
    }

    registrarPedido(dados);
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function normalizarReferencia(val) {
  if (val === undefined || val === null) return '';
  var s = String(val).trim();
  if (s.endsWith('.0')) s = s.slice(0, -2);
  return s.toLowerCase();
}

function obterBase64DriveFile(fileId) {
  if (!fileId) return '';
  try {
    var file  = DriveApp.getFileById(fileId);
    var mime  = file.getMimeType();
    var bytes = file.getBlob().getBytes();
    return 'data:' + mime + ';base64,' + Utilities.base64Encode(bytes);
  } catch (e) {
    return '';
  }
}

function getProdutos(forceUpdate) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var cache = CacheService.getScriptCache();
  var clientes = [];

  // --- OTIMIZAÇÃO: Tenta recuperar Clientes do Cache primeiro ---
  if (!forceUpdate) {
    try {
      var cacheCliInfo = cache.get('cat_clientes_info');
      if (cacheCliInfo) {
        var infoCli = JSON.parse(cacheCliInfo);
        var strCliCache = '';
        var validCli = true;
        for (var i = 0; i < infoCli.chunks; i++) {
          var chunkCli = cache.get('cat_cli_chk_' + i);
          if (!chunkCli) { validCli = false; break; }
          strCliCache += chunkCli;
        }
        if (validCli) {
          clientes = JSON.parse(strCliCache);
        }
      }
    } catch (e) { console.log("Erro ao processar cache de clientes: " + e.message); }
  }

  // Se não houver cache de clientes ou forçado, lê direto da planilha
  if (clientes.length === 0) {
    try {
      var abaClientes = ss.getSheetByName('Clientes');
      if (!abaClientes) {
        abaClientes = ss.insertSheet('Clientes');
        abaClientes.appendRow(['CNPJ', 'RAZÃO SOCIAL', 'NOME FANTASIA', 'TELEFONE', 'ENDEREÇO', 'ESTADO', 'BAIRRO', 'MUNICIPIO', 'NUMERO', 'CEP']);
      } else {
        var dadosClientes = abaClientes.getDataRange().getValues();
        for (var f = 1; f < dadosClientes.length; f++) {
          if(dadosClientes[f][0]) {
            clientes.push({
              cnpj: String(dadosClientes[f][0]).trim(),
              razao: String(dadosClientes[f][1]).trim(),
              fantasia: String(dadosClientes[f][2]).trim(),
              telefone: String(dadosClientes[f][3]).trim(),
              endereco: String(dadosClientes[f][4]).trim(),
              estado: String(dadosClientes[f][5]).trim(),
              bairro: String(dadosClientes[f][6] || '').trim(),
              municipio: String(dadosClientes[f][7] || '').trim(),
              numero: String(dadosClientes[f][8] || '').trim(),
              cep: String(dadosClientes[f][9] || '').trim()
            });
          }
        }
      }
      // Salva a lista de clientes estruturada no cache em Chunks
      var jsonCliStr = JSON.stringify(clientes);
      var chunkCliSize = 90000, chunksCli = Math.ceil(jsonCliStr.length / chunkCliSize);
      for (var c = 0; c < chunksCli; c++) {
        cache.put('cat_cli_chk_' + c, jsonCliStr.substring(c * chunkCliSize, (c + 1) * chunkCliSize), 14400);
      }
      cache.put('cat_clientes_info', JSON.stringify({ chunks: chunksCli }), 14400);
    } catch(e) { console.log("Erro Clientes Planilha: " + e.message); }
  }

  // --- Tenta recuperar Produtos do Cache ---
  if (!forceUpdate) {
    try {
      var cacheInfo = cache.get('cat_info');
      if (cacheInfo) {
        var info = JSON.parse(cacheInfo), strCache = '', valid = true;
        for (var i = 0; i < info.chunks; i++) {
          var chunk = cache.get('cat_chk_' + i);
          if (!chunk) { valid = false; break; }
          strCache += chunk;
        }
        if (valid) {
          var resCache = JSON.parse(strCache);
          resCache.clientes = clientes; 
          return resCache;
        }
      }
    } catch (e) {}
  }

  // Fallback: Reconstrução total do cache caso expire ou forceUpdate seja true
  var estados = [], freteRegras = {};
  try {
    var abaFrete = ss.getSheetByName('Frete');
    if (abaFrete) {
      var dadosFrete = abaFrete.getDataRange().getValues();
      var estadosSet = {};
      for (var f = 1; f < dadosFrete.length; f++) {
        var estadoVal = String(dadosFrete[f][0]).trim().toUpperCase();
        if (estadoVal && estadoVal !== 'ESTADO' && estadoVal.length === 2) {
          estadosSet[estadoVal] = true;
          freteRegras[estadoVal] = {
            gratis: parseFloat(String(dadosFrete[f][1]).replace(',', '.')) || 0,
            pedidoMinimo: parseFloat(String(dadosFrete[f][2]).replace(',', '.')) || 0,
            intervalo: parseFloat(String(dadosFrete[f][3]).replace(',', '.')) || 0
          };
        }
      }
      estados = Object.keys(estadosSet).sort();
    }
  } catch (e) {}

  if (!estados || estados.length === 0) {
    estados = ['AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA','PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO'];
    for(var i=0; i<estados.length; i++){ freteRegras[estados[i]] = { pedidoMinimo: 0, gratis: 0, intervalo: 0 }; }
  }

  var sheet = ss.getSheetByName('Tabela') || ss.getSheets()[0];
  var data  = sheet.getDataRange().getValues();
  var produtosMap = {}, mapaImagens = {}, logoId = '';

  try {
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var files  = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var nomeLower = file.getName().toLowerCase();
      if (nomeLower.indexOf('logo') === 0) { logoId = file.getId(); continue; }
      mapaImagens[nomeLower.split('.')[0]] = file.getId();
    }
  } catch (e) {}

  var keysImagens = Object.keys(mapaImagens);
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var tbPreco = String(row[0]).trim();
    if (!tbPreco || tbPreco.toLowerCase() === 'tb_preco') continue;

    var cod = String(row[1]).trim();
    if (!cod) continue;
    var codNorm = normalizarReferencia(row[1]);
    var precoVal = parseFloat(String(row[3]).replace(',', '.')) || 0;
    var precoPromo = parseFloat(String(row[14]).replace(',', '.')) || 0; 
    var promoAtiva = String(row[15]).toLowerCase().trim() === 'true';   

    if (!produtosMap[codNorm]) {
      var fileId = mapaImagens[codNorm];
      if (!fileId) {
        for (var k = 0; k < keysImagens.length; k++) {
          if (codNorm.startsWith(keysImagens[k]) || keysImagens[k].startsWith(codNorm)) {
            fileId = mapaImagens[keysImagens[k]]; mapaImagens[codNorm] = fileId; break;
          }
        }
      }
      produtosMap[codNorm] = {
        codigo: cod, descricao: String(row[2]).trim(), ncm: String(row[4]).trim(),
        ipi: parseFloat(String(row[5]).replace(',', '.')) || 0,
        qtdEmbalagem: parseInt(row[6]) || 1, codigoEan: String(row[7]).trim(),
        precos: {}, precosPromo: {}, emPromocao: false, fileId: fileId || ''
      };
    }
    produtosMap[codNorm].precos[tbPreco] = precoVal;
    produtosMap[codNorm].precosPromo[tbPreco] = precoPromo;
    if (promoAtiva) produtosMap[codNorm].emPromocao = true;
  }

  var produtos = Object.values(produtosMap).sort(function (a, b) {
    if (a.emPromocao && !b.emPromocao) return -1;
    if (!a.emPromocao && b.emPromocao) return 1;
    return 0;
  });

  var resultadoParaCache = { produtos: produtos, logoUrl : logoId ? 'https://drive.google.com/thumbnail?id=' + logoId + '&sz=w400' : '', estados : estados, freteRegras: freteRegras };
  
  try {
    var jsonStr = JSON.stringify(resultadoParaCache);
    var chunkSize = 90000, chunks = Math.ceil(jsonStr.length / chunkSize);
    for (var c = 0; c < chunks; c++) {
      cache.put('cat_chk_' + c, jsonStr.substring(c * chunkSize, (c + 1) * chunkSize), 14400);
    }
    cache.put('cat_info', JSON.stringify({ chunks: chunks }), 14400);
  } catch (e) {}

  resultadoParaCache.clientes = clientes; 
  return resultadoParaCache;
}

function registrarPedido(dados) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var abaPedidos = ss.getSheetByName('Pedidos');
  if (!abaPedidos) {
    abaPedidos = ss.insertSheet('Pedidos');
    abaPedidos.appendRow(['Data', 'Representante', 'Quantidade (Caixas)', 'Subtotal Produtos', 'IPI', 'Descontos (Prazo)', 'Prazo Pagamento', 'Total Líquido', 'Dados do Cliente', 'Itens do Pedido']);
  }
  var dataStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  abaPedidos.appendRow([ dataStr, dados.codigoRepre || 'Repre', dados.qtd, dados.subtotalProdutos, dados.totalIpi, dados.totalDescontos, dados.prazo, dados.total, dados.clienteInfo, dados.itens ]);
}

function fmtBRL(v) {
  var n = parseFloat(v) || 0;
  return 'R$ ' + n.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function gerarPdfPedido(dados) {
  var dataAtual = new Date();
  var dataFormatada = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  // ── cabeçalho da página ──────────────────────────────────────────────────
  var html = "<!DOCTYPE html><html><head><meta charset='UTF-8'><style>" +
    "* { box-sizing:border-box; margin:0; padding:0; }" +
    "body { font-family:Arial,sans-serif; font-size:10px; color:#222; padding:14px; }" +
    "h1 { font-size:18px; font-weight:bold; color:#1D9E75; }" +
    "h2 { font-size:13px; font-weight:bold; color:#1D9E75; margin-bottom:6px; }" +
    ".sub { font-size:11px; color:#555; margin-bottom:12px; }" +
    ".info-box { border:1px solid #ddd; border-radius:4px; padding:10px 12px; margin-bottom:12px; font-size:10px; line-height:1.7; }" +
    ".info-box b { color:#0F6E56; }" +
    "table { width:100%; border-collapse:collapse; margin-bottom:14px; font-size:9.5px; }" +
    "thead th { background:#f4f4f4; border:1px solid #ccc; padding:5px 4px; text-align:center; font-weight:bold; color:#333; }" +
    "tbody td { border:1px solid #ddd; padding:5px 4px; vertical-align:middle; }" +
    "tbody tr:nth-child(even) td { background:#fafafa; }" +
    ".td-foto { text-align:center; width:72px; }" +
    ".td-foto img { max-width:68px; max-height:68px; object-fit:contain; }" +
    ".td-ref { width:80px; font-weight:bold; color:#0F6E56; font-size:8px; }" +
    ".td-desc { font-size:8px; }" +
    ".td-num { text-align:center; width:38px; }" +
    ".td-val { text-align:right; width:80px; }" +
    ".td-ipi-pct { text-align:center; width:44px; }" +
    ".totals-wrap { margin-left:auto; width:340px; font-size:10px; }" +
    ".tot-row { display:flex; justify-content:space-between; padding:3px 0; border-bottom:1px solid #eee; }" +
    ".tot-row:last-child { border-bottom:none; }" +
    ".tot-label { color:#555; }" +
    ".tot-val { font-weight:bold; color:#333; }" +
    ".tot-desc { color:#c0392b; }" +
    ".tot-ipi { color:#2980b9; }" +
    ".tot-final { font-size:13px; font-weight:bold; color:#1D9E75; border-top:2px solid #1D9E75; padding-top:5px; margin-top:4px; display:flex; justify-content:space-between; }" +
    ".obs-box { margin-top:14px; font-size:9px; color:#555; border-top:1px solid #eee; padding-top:8px; }" +
    "</style></head><body>";

  // ── título ───────────────────────────────────────────────────────────────
  html += "<h1>DI SOLLE - PEDIDO DE COMPRA</h1>" +
    "<div class='sub'>Data: " + dataFormatada + "</div>" +
    "<h2>REPRESENTANTE: " + (dados.codigoRepre || '-') + "</h2>";

  // ── dados do cliente ──────────────────────────────────────────────────────
  if (dados.clienteInfo) {
    html += "<div class='info-box'><b>DADOS DO CLIENTE / PEDIDO:</b><br>" +
      dados.clienteInfo.replace(/\n/g, '<br>');
    if (dados.observacoes) {
      html += "<br><b>Observações:</b><br>" + dados.observacoes;
    }
    // estado e prazo na mesma linha
    html += "<br><b>Estado Destino (UF):</b> " + (dados.estado || '-') +
      " &nbsp;|&nbsp; <b>Prazo Selecionado:</b> " + (dados.prazo || '-') +
      "</div>";
  } else {
    html += "<div class='info-box'>" +
      "<b>Estado Destino (UF):</b> " + (dados.estado || '-') +
      " &nbsp;|&nbsp; <b>Prazo Selecionado:</b> " + (dados.prazo || '-') +
      "</div>";
  }

  // ── tabela de itens ────────────────────────────────────────────────────────
  html += "<table><thead><tr>" +
    "<th class='td-foto'>Foto</th>" +
    "<th class='td-ref'>Referência</th>" +
    "<th class='td-desc'>Descrição</th>" +
    "<th class='td-num'>Qtd</th>" +
    "<th class='td-val'>Valor c/<br>Desconto</th>" +
    "<th class='td-ipi-pct'>Valor<br>IPI</th>" +
    "<th class='td-val'>Preço +<br>IPI</th>" +
    "<th class='td-val'>Total<br>Líquido</th>" +
    "</tr></thead><tbody>";

  var subtotalItensPdf = 0; // acumulador: valorComDesconto × qtd de cada item

  for (var i = 0; i < dados.itens.length; i++) {
    var it = dados.itens[i];
    var imgTag = '';
    if (it.fileId) {
      var imgData = obterBase64DriveFile(it.fileId);
      if (imgData) imgTag = "<img src='" + imgData + "'>";
    }
    var ipiPct = parseFloat(it.ipi) || 0;
    var ipiLabel = ipiPct.toFixed(2).replace('.', ',') + '%';
    var totalIpiItem = parseFloat(it.valorIpiCada) * parseInt(it.qtd);
    // Garante que valorTotalItem é sempre número (evita problema se vier como string pt-BR)
    var valorTotalItemNum = typeof it.valorTotalItem === 'string'
      ? parseFloat(it.valorTotalItem.replace(/\./g, '').replace(',', '.'))
      : parseFloat(it.valorTotalItem) || 0;
    // Acumula subtotal: coluna 5 (valorComDesconto) × coluna 4 (qtd)
    subtotalItensPdf += parseFloat(it.valorComDesconto) * parseInt(it.qtd);

    html += "<tr>" +
      "<td class='td-foto'>" + imgTag + "</td>" +
      "<td class='td-ref'>" + it.codigo + "</td>" +
      "<td class='td-desc'>" + it.descricao + "</td>" +
      "<td class='td-num'>" + it.qtd + "</td>" +
      "<td class='td-val'>" + fmtBRL(it.valorComDesconto) + "</td>" +
      "<td class='td-ipi-pct'>" + fmtBRL(totalIpiItem) + "<br><span style='font-size:8px;color:#888'>(" + ipiLabel + ")</span></td>" +
      "<td class='td-val'>" + fmtBRL(it.valorComIpi) + "</td>" +
      "<td class='td-val'><b>" + fmtBRL(valorTotalItemNum) + "</b></td>" +
      "</tr>";
  }

  html += "</tbody></table>";

  // ── bloco de totais ────────────────────────────────────────────────────────
  var c = dados.contas;
  var freteLabel = (c.valorFrete && c.valorFrete > 0) ? fmtBRL(c.valorFrete) : 'R$ 0,00';

  html += "<div class='totals-wrap'>" +
    "<div class='tot-row'><span class='tot-label'>Subtotal Itens:</span><span class='tot-val'>" + fmtBRL(subtotalItensPdf) + "</span></div>" +
    "<div class='tot-row'><span class='tot-label'>IPI da Indústria (+):</span><span class='tot-val tot-ipi'>+ " + fmtBRL(c.totalIpi) + "</span></div>" +
    "<div class='tot-row'><span class='tot-label'>Frete:</span><span class='tot-val'>" + freteLabel + "</span></div>" +
    "<div class='tot-final'><span>TOTAL LÍQUIDO FINAL:</span><span>" + fmtBRL(c.liquido) + "</span></div>" +
    "</div>";

  if (dados.observacoes && dados.clienteInfo) {
    // já inserido acima; não duplicar
  } else if (dados.observacoes) {
    html += "<div class='obs-box'><b>Observações:</b> " + dados.observacoes + "</div>";
  }

  html += "</body></html>";

  // ── gerar PDF ──────────────────────────────────────────────────────────────
  var pdfBlob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
  // Nome do arquivo: Representante_ddMMyyyy_NomeCliente.pdf
  var nomeCliente = '';
  if (dados.cliente && dados.cliente.razao) {
    nomeCliente = dados.cliente.razao;
  } else if (dados.clienteInfo) {
    // Tenta extrair Razão Social do texto clienteInfo
    var matchRazao = dados.clienteInfo.match(/Razão Social:\s*(.+)/);
    if (matchRazao) nomeCliente = matchRazao[1].trim();
  }
  // Remove caracteres inválidos para nome de arquivo
  var nomeClienteLimpo = nomeCliente.replace(/[\/\\:*?"<>|]/g, '').trim() || 'SemCliente';
  var nomeArquivoPdf = (dados.codigoRepre || 'Repre') + '_' +
    Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "ddMMyyyy") + '_' +
    nomeClienteLimpo + '.pdf';
  pdfBlob.setName(nomeArquivoPdf);

  // ← CORREÇÃO CRÍTICA: inclui prefixo data URI para download funcionar no browser
  var base64 = 'data:application/pdf;base64,' + Utilities.base64Encode(pdfBlob.getBytes());

  var emailStatus = "Não aplicável";
  var enviarEmail = false;
  var folderDestinoId = '';

  if (dados.tipoAcao === 'enviar') {
    folderDestinoId = '1g_3tAV4rY765KuTh5XuhSt4mpfYVBjPW';
    enviarEmail = true;
  } else if (dados.tipoAcao === 'enviar_disolle') {
    folderDestinoId = '1Jdxb2j_dOWTJtgABUNVOcM6kYLVkmPTH';
    enviarEmail = false;
  }

  if (folderDestinoId !== '') {
    try { DriveApp.getFolderById(folderDestinoId).createFile(pdfBlob); } catch(e) {}
  }

  if (enviarEmail) {
    try {
      if (MailApp.getRemainingDailyQuota() > 0) {
        MailApp.sendEmail({
          to: 'roana@disolle.com.br',
          subject: 'Novo Pedido - Representante: ' + dados.codigoRepre,
          body: 'O representante ' + dados.codigoRepre + ' enviou um novo pedido.\nData: ' + dataFormatada + '\nTotal: ' + fmtBRL(c.liquido),
          attachments: [pdfBlob]
        });
        emailStatus = "Enviado com sucesso";
      } else { emailStatus = "Cota de e-mail diária excedida"; }
    } catch (e) { emailStatus = "Erro ao enviar e-mail: " + e.message; }
  }

  return { base64: base64, emailStatus: emailStatus, nomeArquivo: nomeArquivoPdf };
}