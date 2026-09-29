/* =============================================
   Di Solle — 11. CLIENTES — BUSCA E CADASTRO
   Modal de novo cliente, busca geral de clientes cadastrados e
   preenchimento automático dos dados ao digitar um CNPJ conhecido.
   Depende de: 01-estado-global.js, 05-utils-busca.js,
               03-persistencia-local.js,
               06-precos-kne825-millenium.js (ativarClienteKNE825),
               10-carrinho.js (calcularTudo)
   ============================================= */

function verificarNovoClienteExistente(cnpj) {
  if (!cnpj) return;
  let cLimpo = cnpj.replace(/\D/g, '').trim();
  let c = CLIENTES.find(x => x.cnpj.replace(/\D/g, '') === cLimpo);
  if (c) {
    alert("⚠️ ALERTA IMPEDITIVO: Este CNPJ já existe cadastrado na planilha! Não é permitido criar duplicados.");
    ['razao','fantasia','telefone','endereco','estado','bairro','municipio','numero','cep','email'].forEach(f => {
      document.getElementById('nc-' + f).value = c[f] || '';
    });
    BLOQUEIA_SALVAMENTO_CNPJ = true;
    document.getElementById('btn-salvar-nc').disabled = true;
  } else {
    BLOQUEIA_SALVAMENTO_CNPJ = false;
    document.getElementById('btn-salvar-nc').disabled = false;
  }
}

function buscarClienteAoDigitar(cnpj) {
  if (!cnpj) return;
  let cLimpo = cnpj.replace(/\D/g, '').trim();
  let c = CLIENTES.find(x => x.cnpj.replace(/\D/g, '') === cLimpo);
  if (c) {
    ['razao','fantasia','telefone','endereco','estado','bairro','municipio','numero','cep','email'].forEach(f => {
      document.getElementById('cli-' + f).value = c[f] || '';
    });
    salvarClienteLocal();
    showToast("✅ Dados do cliente preenchidos automaticamente.");
    ativarClienteKNE825(cnpj);
    atualizarBotaoMix();
  } else {
    atualizarBotaoMix();
    if (confirm("❌ Cliente não localizado! Deseja abrir a tela de cadastro para este CNPJ agora?")) {
      fecharModalCliente();
      setTimeout(() => {
        abrirModalNovoCliente();
        document.getElementById('nc-cnpj').value = cnpj;
      }, 350);
    }
  }
}

function salvarNovoCliente() {
  let cCnpj = document.getElementById('nc-cnpj').value;
  if (BLOQUEIA_SALVAMENTO_CNPJ || CLIENTES.find(x => x.cnpj.replace(/\D/g, '') === cCnpj.replace(/\D/g, ''))) {
    alert("❌ Operação abortada! CNPJ duplicado na base de dados.");
    return;
  }

  let c = {
    cnpj: cCnpj,
    razao: document.getElementById('nc-razao').value,
    fantasia: document.getElementById('nc-fantasia').value,
    telefone: document.getElementById('nc-telefone').value,
    endereco: document.getElementById('nc-endereco').value,
    estado: document.getElementById('nc-estado').value,
    bairro: document.getElementById('nc-bairro').value,
    municipio: document.getElementById('nc-municipio').value,
    numero: document.getElementById('nc-numero').value,
    cep: document.getElementById('nc-cep').value,
    email: document.getElementById('nc-email').value.trim()
  };

  if (!c.cnpj || !c.razao) { alert("Preencha obrigatoriamente CNPJ e Razão Social."); return; }
  if (!c.email) { alert("Preencha obrigatoriamente o E-mail."); return; }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(c.email)) { alert("E-mail inválido."); return; }

  document.getElementById('loading-modal').style.display = 'flex';
  document.getElementById('loading-modal').classList.add('open');

  fetch(URL_GOOGLE_SCRIPT, { method: 'POST', body: JSON.stringify({ acao: 'salvar_cliente', cliente: c }) })
    .then(r => r.json()).then(res => {
      document.getElementById('loading-modal').classList.remove('open');
      document.getElementById('loading-modal').style.display = 'none';
      if (res.status === 'success') {
        CLIENTES.push(c);
        showToast("✅ Cliente salvo com sucesso!");
        fecharModalNovoCliente();
        if (Object.keys(SELECIONADOS).length > 0) {
          document.getElementById('modal-cliente').style.display = 'flex';
          document.getElementById('modal-cliente').classList.add('open');
          ['cnpj','razao','fantasia','telefone','endereco','estado','bairro','municipio','numero','cep','email'].forEach(f => {
            document.getElementById('cli-' + f).value = c[f] || '';
          });
          salvarClienteLocal();
        }
      } else { alert(res.message); }
    }).catch(() => {
      document.getElementById('loading-modal').classList.remove('open');
      document.getElementById('loading-modal').style.display = 'none';
      alert("Erro de conexão.");
    });
}

function abrirModalBuscarCliente() {
  document.getElementById('input-busca-cliente').value = '';
  document.getElementById('lista-busca-clientes').innerHTML = '<div class="vazio">Digite CNPJ, Fantasia ou Cidade para pesquisar...</div>';
  document.getElementById('modal-buscar-cliente').style.display = 'flex';
  document.getElementById('modal-buscar-cliente').classList.add('open');
}

function fecharModalBuscarCliente() {
  document.getElementById('modal-buscar-cliente').classList.remove('open');
  setTimeout(() => document.getElementById('modal-buscar-cliente').style.display = 'none', 300);
}

function fecharModalDetalhesCliente() {
  document.getElementById('modal-detalhes-cliente').classList.remove('open');
  setTimeout(() => document.getElementById('modal-detalhes-cliente').style.display = 'none', 300);
}

function executarBuscaCliente() {
  let v = document.getElementById('input-busca-cliente').value.trim();
  if (!v) { alert("Digite algum parâmetro para pesquisar."); return; }

  document.getElementById('loading-modal').style.display = 'flex';
  document.getElementById('loading-modal').classList.add('open');

  setTimeout(() => {
    filtrarClientesBusca();
    document.getElementById('loading-modal').classList.remove('open');
    document.getElementById('loading-modal').style.display = 'none';
  }, 300);
}

function filtrarClientesBusca() {
  let v = document.getElementById('input-busca-cliente').value.trim();
  let container = document.getElementById('lista-busca-clientes');
  container.innerHTML = '';

  if (!v) { container.innerHTML = '<div class="vazio">Digite CNPJ, nome, fantasia ou cidade para pesquisar...</div>'; return; }

  const esc = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const soDigitos = s => String(s || '').replace(/\D/g, '');
  const fmtDoc = s => {
    let d = soDigitos(s);
    if (d.length === 14) return d.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
    if (d.length === 11) return d.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, '$1.$2.$3-$4');
    return String(s || '');
  };

  let termoNorm = normalizarTexto(v);
  let termoDigitos = soDigitos(v);
  let buscaNumerica = termoDigitos.length >= 3 && !/[a-zA-Z]/.test(v);
  let tokens = termoNorm.split(/\s+/).filter(Boolean);

  // 1) Busca exata: TODAS as palavras precisam aparecer (parte do nome, fantasia, razão, cidade, UF, bairro, e-mail)
  //    Números batem com CNPJ/CPF ou telefone (com ou sem pontuação).
  let filtrados = CLIENTES.filter(c => {
    let cnpjDig = soDigitos(c.cnpj);
    if (buscaNumerica) {
      return cnpjDig.includes(termoDigitos) || soDigitos(c.telefone).includes(termoDigitos);
    }
    let txt = normalizarTexto([c.fantasia, c.razao, c.municipio, c.estado, c.bairro, c.email].join(' '));
    return tokens.every(t => {
      if (txt.includes(t)) return true;
      let dt = soDigitos(t);
      return !/[a-zA-Z]/.test(t) && dt.length >= 3 && cnpjDig.includes(dt);
    });
  });

  // 2) Se nada foi encontrado, tenta aproximado (tolera erro de digitação) só em nome/fantasia/cidade
  let aproximado = false;
  if (filtrados.length === 0 && !buscaNumerica) {
    filtrados = CLIENTES.filter(c => buscaInteligente([c.fantasia, c.razao, c.municipio], v));
    aproximado = true;
  }

  if (filtrados.length === 0) { container.innerHTML = '<div class="vazio">Nenhum cliente localizado na base.</div>'; return; }

  // Ordena: nome idêntico > contém a frase inteira > começa com a 1ª palavra > demais
  let pontos = c => {
    let f = normalizarTexto(c.fantasia), r = normalizarTexto(c.razao);
    if (f === termoNorm || r === termoNorm) return 0;
    if (f.includes(termoNorm) || r.includes(termoNorm)) return 1;
    if (tokens.length && (f.startsWith(tokens[0]) || r.startsWith(tokens[0]))) return 2;
    return 3;
  };
  filtrados.sort((a, b) => pontos(a) - pontos(b));

  const LIMITE = 50;
  let total = filtrados.length;
  filtrados = filtrados.slice(0, LIMITE);

  let cab = document.createElement('div');
  cab.style.cssText = 'padding:8px 10px;font-size:11px;color:var(--sub);background:#f9f9f9;border-bottom:1px solid var(--borda);';
  cab.innerText = (aproximado ? 'Sem resultado exato — mostrando aproximados: ' : '') +
    total + ' cliente(s) encontrado(s)' + (total > LIMITE ? ' (mostrando os ' + LIMITE + ' primeiros — refine a busca)' : '');
  container.appendChild(cab);

  filtrados.forEach(c => {
    let d = document.createElement('div');
    d.className = 'sel-row';
    d.style.cursor = 'pointer';
    d.style.padding = '10px';
    d.onclick = () => mostrarFichaCompletaCliente(c);

    let nomePrincipal = c.fantasia || c.razao || '-';
    let linhas = [];
    if (c.razao && c.razao !== nomePrincipal) linhas.push('<b>Razão:</b> ' + esc(c.razao));
    linhas.push('<b>CNPJ/CPF:</b> ' + esc(fmtDoc(c.cnpj)) + (c.telefone ? ' &nbsp;|&nbsp; <b>Tel:</b> ' + esc(c.telefone) : ''));
    let ender = [c.endereco, c.numero].filter(Boolean).join(', ');
    if (c.bairro) ender += (ender ? ' — ' : '') + c.bairro;
    if (ender) linhas.push('<b>Endereço:</b> ' + esc(ender));
    let cidade = [c.municipio, c.estado].filter(Boolean).join('/');
    if (cidade || c.cep) linhas.push('<b>Cidade:</b> ' + esc(cidade || '-') + (c.cep ? ' &nbsp;|&nbsp; <b>CEP:</b> ' + esc(c.cep) : ''));
    if (c.email) linhas.push('<b>E-mail:</b> ' + esc(c.email));

    d.innerHTML = `<div style="display:flex;flex-direction:column;width:100%;gap:2px;">
      <span style="font-weight:bold;color:var(--verde-dk);">${esc(nomePrincipal)}</span>
      <span style="font-size:11px;color:var(--sub);line-height:1.5;">${linhas.join('<br>')}</span>
    </div>`;
    container.appendChild(d);
  });
}

function mostrarFichaCompletaCliente(c) {
  document.getElementById('conteudo-detalhes-cliente').innerHTML = `
    <div style="margin-bottom:6px;"><b>CNPJ:</b> ${c.cnpj || '-'}</div>
    <div style="margin-bottom:6px;"><b>RAZÃO SOCIAL:</b> ${c.razao || '-'}</div>
    <div style="margin-bottom:6px;"><b>NOME FANTASIA:</b> ${c.fantasia || '-'}</div>
    <div style="margin-bottom:6px;"><b>TELEFONE:</b> ${c.telefone || '-'}</div>
    <div style="margin-bottom:6px;"><b>ENDEREÇO:</b> ${c.endereco || '-'}</div>
    <div style="margin-bottom:6px;"><b>ESTADO:</b> ${c.estado || '-'}</div>
    <div style="margin-bottom:6px;"><b>BAIRRO:</b> ${c.bairro || '-'}</div>
    <div style="margin-bottom:6px;"><b>MUNICÍPIO:</b> ${c.municipio || '-'}</div>
    <div style="margin-bottom:6px;"><b>NÚMERO:</b> ${c.numero || '-'}</div>
    <div style="margin-bottom:6px;"><b>CEP:</b> ${c.cep || '-'}</div>
    <div style="margin-bottom:6px;"><b>E-MAIL:</b> ${c.email || '-'}</div>
  `;

  let btnUsar = document.getElementById('btn-selecionar-cliente-busca');
  btnUsar.style.display = 'block';
  
  btnUsar.onclick = () => {
    ['cnpj','razao','fantasia','telefone','endereco','estado','bairro','municipio','numero','cep','email'].forEach(f => {
      let input = document.getElementById('cli-' + f);
      if (input) input.value = c[f] || '';
    });
    salvarClienteLocal();
    atualizarBotaoMix();

    if(c.estado) {
      let estadoUpper = c.estado.toUpperCase().trim();
      let optD = document.querySelector(`#uf-d option[value="${estadoUpper}"]`);
      if(optD) {
        document.getElementById('uf-d').value = estadoUpper;
        document.getElementById('uf-m').value = estadoUpper;
        calcularTudo();
      }
    }

    ativarClienteKNE825(c.cnpj || '');
    fecharModalDetalhesCliente();
    fecharModalBuscarCliente();
    showToast("✅ Cliente vinculado! Adicione os itens e finalize.");
  };

  document.getElementById('modal-detalhes-cliente').style.display = 'flex';
  document.getElementById('modal-detalhes-cliente').classList.add('open');
}
