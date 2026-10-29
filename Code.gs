/***********************************************************************
 * RHITMO ATENDE — app do RH para os funcionários
 * Ponto, atendimento com o RH, recibos de pagamento e informe de rendimentos.
 *
 * Roda de graça no Google Apps Script, na conta Google do RH.
 * Login próprio: o funcionário entra com MATRÍCULA e SENHA;
 * o RH entra com o usuário "rh".
 *
 * Os dados ficam numa planilha e numa pasta do Google Drive criadas
 * pela função configurar(). NÃO compartilhe essa planilha nem a pasta.
 ***********************************************************************/

const TZ = 'America/Sao_Paulo';
const SESSAO_SEG = 6 * 60 * 60;          // sessão dura 6 h sem uso
const MAX_PDF = 10 * 1024 * 1024;        // 10 MB por PDF
const MAX_TEXTO = 3000;                  // tamanho máximo de uma mensagem

const ABAS = {
  Funcionarios: ['matricula', 'nome', 'setor', 'cargo', 'hash', 'salt', 'ativo', 'trocarSenha', 'criadoEm', 'empresa', 'cpf', 'horario', 'usuario', 'pis'],
  Chamados: ['id', 'matricula', 'assunto', 'status', 'criadoEm', 'atualizadoEm', 'naoLidoRH', 'naoLidoFunc'],
  Mensagens: ['id', 'chamadoId', 'de', 'texto', 'em'],
  Documentos: ['id', 'matricula', 'tipo', 'titulo', 'competencia', 'ano', 'fileId', 'arquivo', 'tamanho', 'enviadoEm'],
  BancoHoras: ['matricula', 'saldo', 'prazoMeses', 'abertos', 'lanc', 'comp', 'atualizadoEm'],
  Respostas: ['id', 'palavras', 'resposta', 'ativo', 'criadoEm', 'avisar'],
  Horarios: ['codigo', 'descricao', 'ent', 'si', 'ri', 'sai', 'dias', 'carga'],
  Ferias: ['id', 'matricula', 'inicio', 'dias', 'abono', 'obs', 'status', 'criadoEm', 'respostaRH', 'respondidoEm'],
  Fotos: ['matricula', 'foto', 'atualizadoEm'],
  Aparelhos: ['hash', 'papel', 'matricula', 'salt', 'criadoEm', 'usadoEm']   // "manter conectado": um registro por celular/computador
};
const APARELHO_DIAS = 120;   // o login salvo vale 120 dias sem uso; trocar a senha derruba todos os aparelhos
const CAB_PONTO = ['id', 'matricula', 'data', 'hora', 'em', 'nsr', 'origem'];   // uma aba por mês: "Ponto 2026-10"; nsr = nº sequencial do AFD; origem: app (vazio), dixi ou rhitmo

function doGet() {
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  const html = HtmlService.createHtmlOutputFromFile('Index').getContent().replace('__URL_APP__', url);
  return HtmlService.createHtmlOutput(html)
    .setTitle('Rhitmo Atende')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);   // permite colocar o app dentro do Google Sites
}

/* =====================================================================
 * SEGUNDO CAMINHO DE CONEXÃO
 * Alguns celulares (principalmente iPhone/Safari) bloqueiam o caminho padrão do
 * Google ("NetworkError ... HTTP 0"). Nesses casos o app chama as mesmas funções
 * por aqui, com um POST simples no link do app. Só as funções da lista abaixo podem
 * ser chamadas, e cada uma continua conferindo o login (token) como sempre.
 * ===================================================================== */
function apiPublica_() {
  return {
    login: login,
    entrarAparelho: entrarAparelho,
    sair: sair,
    trocarSenha: trocarSenha,
    inicio: inicio,
    baterPonto: baterPonto,
    abrirChamado: abrirChamado,
    enviarMensagem: enviarMensagem,
    marcarLido: marcarLido,
    baixarDoc: baixarDoc,
    rhPainel: rhPainel,
    rhFuncionario: rhFuncionario,
    rhSalvarFuncionario: rhSalvarFuncionario,
    rhResetSenha: rhResetSenha,
    rhAtivar: rhAtivar,
    rhEncerrar: rhEncerrar,
    rhEnviarDoc: rhEnviarDoc,
    rhExcluirDoc: rhExcluirDoc,
    rhConfig: rhConfig,
    rhImportar: rhImportar,
    rhExportarAFD: rhExportarAFD,
    rhSalvarBH: rhSalvarBH,
    funcChat: funcChat,
    funcFerias: funcFerias,
    salvarFoto: salvarFoto,
    rhRespostas: rhRespostas,
    rhSalvarResposta: rhSalvarResposta,
    rhRespostasPadrao: rhRespostasPadrao,
    rhFerias: rhFerias,
    rhResponderFerias: rhResponderFerias,
    rhIA: rhIA,
    rhSalvarIA: rhSalvarIA,
    rhTestarIA: rhTestarIA,
    rhAgendaAFD: rhAgendaAFD,
    rhSalvarAgendaAFD: rhSalvarAgendaAFD,
    rhGerarAFDsAgora: rhGerarAFDsAgora,
    funcSalvarUsuario: funcSalvarUsuario,
    rhImportarHorarios: rhImportarHorarios,
    funcEspelho: funcEspelho,
    rhEspelho: rhEspelho,
    rhModeloEspelho: rhModeloEspelho,
    rhImportarMarcacoes: rhImportarMarcacoes,
    rhCpfsFuncionarios: rhCpfsFuncionarios
  };
}
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e && e.postData ? e.postData.contents : '{}');
    const fn = apiPublica_()[req.fn];
    if (!fn) throw new Error('Função não permitida.');
    const r = fn.apply(null, Array.isArray(req.args) ? req.args : []);
    out = { ok: true, r: r === undefined ? null : r };
  } catch (err) {
    out = { ok: false, erro: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* =====================================================================
 * CONFIGURAÇÃO — rode "configurar" UMA vez no editor (botão Executar).
 * A senha provisória do RH aparece no "Registro de execução".
 * ===================================================================== */
function configurar() {
  const p = props_();
  const novo = !p.getProperty('SS_ID');
  let ss;
  if (novo) {
    ss = SpreadsheetApp.create('Rhitmo Atende - Dados (NÃO compartilhar)');
    p.setProperty('SS_ID', ss.getId());
  } else {
    ss = SpreadsheetApp.openById(p.getProperty('SS_ID'));
  }
  SS_ = ss;
  Object.keys(ABAS).forEach(function (n) { aba_(n, ABAS[n]); });
  pontoAba_(agora_().mes);
  if (novo) {
    ss.getSheets().forEach(function (sh) {
      const n = sh.getName();
      if (!ABAS[n] && n.indexOf('Ponto ') !== 0) ss.deleteSheet(sh);
    });
  }
  if (!p.getProperty('PASTA_ID')) {
    p.setProperty('PASTA_ID', DriveApp.createFolder('Rhitmo Atende - Documentos (NÃO compartilhar)').getId());
  }
  if (!p.getProperty('EMPRESA')) p.setProperty('EMPRESA', 'MEPCO');
  if (!p.getProperty('ATENDIMENTO')) p.setProperty('ATENDIMENTO', 'Segunda a sexta, das 8h às 17h');

  if (!p.getProperty('RH_HASH')) {
    const senha = gerarSenha_();
    const salt = gerarSalt_();
    p.setProperties({ RH_SALT: salt, RH_HASH: hash_(senha, salt), RH_TROCAR: '1' });
    Logger.log('PRONTO! Entre no app com  usuário: rh   senha provisória: ' + senha);
  } else {
    Logger.log('O app já estava configurado. Se esqueceu a senha do RH, rode "redefinirSenhaRH".');
  }
  Logger.log('Planilha de dados: ' + ss.getUrl());
}

/** Rode no editor se esquecer a senha do RH. A nova senha aparece no registro. */
function redefinirSenhaRH() {
  const senha = gerarSenha_();
  const salt = gerarSalt_();
  props_().setProperties({ RH_SALT: salt, RH_HASH: hash_(senha, salt), RH_TROCAR: '1' });
  Logger.log('Nova senha provisória do RH (usuário rh): ' + senha);
}

/* =====================================================================
 * LOGIN E SENHAS
 * ===================================================================== */
function login(usuario, senha, prefixo, lembrar) {
  usuario = String(usuario || '').trim();
  prefixo = String(prefixo || '');
  senha = String(senha || '').trim();   // ignora espaço copiado sem querer
  if (!usuario || !senha) throw new Error('Digite a matrícula e a senha.');
  const cache = CacheService.getScriptCache();
  const kf = 'falha_' + usuario.toLowerCase();
  const falhas = Number(cache.get(kf) || 0);
  if (falhas >= 5) throw new Error('Muitas tentativas erradas. Aguarde 15 minutos ou fale com o RH.');

  const p = props_();
  let s = null, nome = '', trocar = false;
  if (usuario.toLowerCase() === 'rh') {
    if (p.getProperty('RH_HASH') && hash_(senha, p.getProperty('RH_SALT')) === p.getProperty('RH_HASH')) {
      s = { papel: 'rh', matricula: 'rh', salt: p.getProperty('RH_SALT') };
      nome = 'RH';
      trocar = p.getProperty('RH_TROCAR') === '1';
    }
  } else {
    const f = acharFuncionario_(usuario, prefixo);
    if (f && f.ativo === '1' && hash_(senha, f.salt) === f.hash) {
      s = { papel: 'func', matricula: f.matricula, salt: f.salt };
      nome = f.nome;
      trocar = f.trocarSenha === '1';
    }
  }
  if (!s) {
    cache.put(kf, String(falhas + 1), 900);
    throw new Error('Matrícula, usuário, CPF ou senha incorretos.');
  }
  cache.remove(kf);
  return { token: criarSessao_(s), papel: s.papel, nome: nome, trocar: trocar, aparelho: lembrar ? criarAparelho_(s) : null };
}

/* ---------- manter conectado neste aparelho ---------- */
function hashAparelho_(tk) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'aparelho|' + tk, Utilities.Charset.UTF_8)); }
function criarAparelho_(s) {
  const tk = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const t = agora_().iso;
  comLock_(function () {
    const lim = new Date(Date.now() - APARELHO_DIAS * 864e5).toISOString();
    const velhos = lerT_('Aparelhos').filter(function (a) { return a.usadoEm < lim; }).map(function (a) { return a._row; });
    const sh = aba_('Aparelhos', ABAS.Aparelhos);
    velhos.sort(function (a, b) { return b - a; }).slice(0, 50).forEach(function (r) { sh.deleteRow(r); });
    if (velhos.length) sujo_('Aparelhos');
    inserirT_('Aparelhos', { hash: hashAparelho_(tk), papel: s.papel, matricula: s.matricula, salt: s.salt, criadoEm: t, usadoEm: t });
  });
  return tk;
}
/** Entra sem digitar a senha, usando o login salvo neste aparelho. */
function entrarAparelho(tk) {
  tk = String(tk || '');
  if (!/^[a-f0-9]{96}$/.test(tk)) throw new Error('APARELHO_INVALIDO');
  const h = hashAparelho_(tk);
  const a = lerT_('Aparelhos').filter(function (x) { return x.hash === h; })[0];
  if (!a || a.usadoEm < new Date(Date.now() - APARELHO_DIAS * 864e5).toISOString()) throw new Error('APARELHO_INVALIDO');
  let s = null, nome = '', trocar = false;
  if (a.papel === 'rh') {
    if (a.salt === props_().getProperty('RH_SALT')) { s = { papel: 'rh', matricula: 'rh', salt: a.salt }; nome = 'RH'; trocar = props_().getProperty('RH_TROCAR') === '1'; }
  } else {
    const f = lerT_('Funcionarios').filter(function (x) { return x.matricula === a.matricula; })[0];
    if (f && f.ativo === '1' && f.salt === a.salt) { s = { papel: 'func', matricula: f.matricula, salt: f.salt }; nome = f.nome; trocar = f.trocarSenha === '1'; }
  }
  if (!s) {   // senha trocada/redefinida ou funcionário desativado: o login salvo deixa de valer
    comLock_(function () { apagarLinha_('Aparelhos', a._row); });
    throw new Error('APARELHO_INVALIDO');
  }
  if (a.usadoEm < new Date(Date.now() - 864e5).toISOString()) comLock_(function () { atualizarT_('Aparelhos', a._row, { usadoEm: agora_().iso }); });
  return { token: criarSessao_(s), papel: s.papel, nome: nome, trocar: trocar };
}
function esquecerAparelho_(tk) {
  if (!tk) return;
  const h = hashAparelho_(String(tk));
  comLock_(function () {
    const a = lerT_('Aparelhos').filter(function (x) { return x.hash === h; })[0];
    if (a) apagarLinha_('Aparelhos', a._row);
  });
}

/** Acha o funcionário pelo CPF, pelo usuário ou pela matrícula (com o prefixo da empresa escolhida). */
function acharFuncionario_(u, prefixo) {
  const lista = lerT_('Funcionarios');
  const dig = u.replace(/\D/g, '');
  if (dig.length === 11 && /^[\d.\-\s]+$/.test(u)) {
    const porCpf = lista.filter(function (x) { return x.cpf === dig; });
    if (porCpf.length) return porCpf.filter(function (x) { return x.ativo === '1'; })[0] || porCpf[0];
  }
  const ul = u.toLowerCase();
  const porUsuario = lista.filter(function (x) { return x.usuario && x.usuario.toLowerCase() === ul; })[0];
  if (porUsuario) return porUsuario;
  const tentativas = [];
  if (prefixo && u.toUpperCase().indexOf(prefixo.toUpperCase()) !== 0) tentativas.push(prefixo + u);
  tentativas.push(u);
  for (let i = 0; i < tentativas.length; i++) {
    const f = lista.filter(function (x) { return x.matricula === tentativas[i]; })[0];
    if (f) return f;
  }
  return null;
}
function gerarUsuario_(nome, usados) {
  const p = normTxt_(nome).split(' ').filter(function (w) { return w && !/^(de|da|do|das|dos|e)$/.test(w); });
  const base = (p[0] || 'func') + (p.length > 1 ? '.' + p[p.length - 1] : '');
  let u = base, n = 2;
  while (usados[u]) u = base + n++;
  usados[u] = 1;
  return u;
}
/** O funcionário troca o próprio usuário de login. */
function funcSalvarUsuario(token, u) {
  const s = sessao_(token, 'func');
  u = String(u || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,30}$/.test(u)) throw new Error('Use de 3 a 30 letras sem acento, números, ponto, hífen ou sublinhado.');
  if (/^\d+$/.test(u)) throw new Error('O usuário precisa ter pelo menos uma letra.');
  return comLock_(function () {
    const lista = lerT_('Funcionarios');
    if (lista.some(function (x) { return x.usuario === u && x.matricula !== s.matricula; })) throw new Error('Esse usuário já está em uso. Escolha outro.');
    const f = lista.filter(function (x) { return x.matricula === s.matricula; })[0];
    atualizarT_('Funcionarios', f._row, { usuario: u });
    return u;
  });
}

function sair(token, aparelho) {
  if (token) CacheService.getScriptCache().remove('s_' + token);
  try { esquecerAparelho_(aparelho); } catch (e) {}
  return true;
}

function trocarSenha(token, atual, nova, aparelho) {
  const s = sessao_(token);
  nova = String(nova || '').trim();
  atual = String(atual || '').trim();
  if (nova.length < 6) throw new Error('A nova senha precisa ter pelo menos 6 caracteres.');
  if (nova === String(atual)) throw new Error('A nova senha precisa ser diferente da atual.');
  const salt = gerarSalt_();
  if (s.papel === 'rh') {
    const p = props_();
    if (hash_(String(atual || ''), p.getProperty('RH_SALT')) !== p.getProperty('RH_HASH')) throw new Error('A senha atual está errada.');
    p.setProperties({ RH_SALT: salt, RH_HASH: hash_(nova, salt), RH_TROCAR: '0' });
  } else {
    if (hash_(String(atual || ''), s.f.salt) !== s.f.hash) throw new Error('A senha atual está errada.');
    comLock_(function () {
      atualizarT_('Funcionarios', s.f._row, { hash: hash_(nova, salt), salt: salt, trocarSenha: '0' });
    });
  }
  CacheService.getScriptCache().remove('s_' + token);
  if (aparelho) {   // este aparelho continua conectado; os outros precisam entrar de novo
    const ha = hashAparelho_(String(aparelho));
    comLock_(function () {
      const a = lerT_('Aparelhos').filter(function (x) { return x.hash === ha && x.matricula === s.matricula; })[0];
      if (a) atualizarT_('Aparelhos', a._row, { salt: salt, usadoEm: agora_().iso });
    });
  }
  return { token: criarSessao_({ papel: s.papel, matricula: s.matricula, salt: salt }) };
}

/* =====================================================================
 * FUNCIONÁRIO
 * ===================================================================== */
function inicio(token, leve) {
  const s = sessao_(token);
  const cfg = config_();
  const t = agora_();
  if (s.papel === 'rh') return { papel: 'rh', cfg: cfg, agora: t, trocar: props_().getProperty('RH_TROCAR') === '1' };
  const f = s.f;
  return {
    papel: 'func', cfg: cfg, agora: t, trocar: f.trocarSenha === '1',
    perfil: { matricula: f.matricula, nome: f.nome, setor: f.setor, cargo: f.cargo, empresa: f.empresa, usuario: f.usuario, temCpf: !!f.cpf, horario: horarioTexto_(f.horario) },
    ponto: pontoDe_(f.matricula, t.mes),
    chamados: chamadosDe_(f.matricula),
    docs: docsDe_(f.matricula),
    bh: bhDe_(f.matricula),
    ferias: feriasDe_(f.matricula),
    foto: leve ? null : fotoDe_(f.matricula)
  };
}

function baterPonto(token) {
  const s = sessao_(token, 'func');
  return comLock_(function () {
    const t = agora_();
    const hoje = pontoDe_(s.matricula, t.mes).filter(function (p) { return p.data === t.data; });
    const ult = hoje[hoje.length - 1];
    if (ult && t.ms - new Date(ult.em).getTime() < 60000) {
      throw new Error('Você acabou de registrar. Aguarde 1 minuto para um novo registro.');
    }
    const sh = pontoAba_(t.mes);
    escreverLinha_(sh, sh.getLastRow() + 1, [id_(), s.matricula, t.data, t.hora, t.iso, proximoNSR_(s.f.empresa), 'app']);
    sujo_('Ponto ' + t.mes);
    return { agora: t, ponto: pontoDe_(s.matricula, t.mes) };
  });
}

function abrirChamado(token, assunto, texto) {
  const s = sessao_(token, 'func');
  texto = limparTexto_(texto);
  assunto = String(assunto || 'Outro assunto').slice(0, 80);
  return comLock_(function () {
    const t = agora_().iso;
    const id = id_();
    inserirT_('Chamados', { id: id, matricula: s.matricula, assunto: assunto, status: 'aberto', criadoEm: t, atualizadoEm: t, naoLidoRH: '1', naoLidoFunc: '0' });
    inserirT_('Mensagens', { id: id_(), chamadoId: id, de: 'func', texto: texto, em: t });
    return { id: id, chamados: chamadosDe_(s.matricula) };
  });
}

/** Funcionário ou RH responde num pedido. */
function enviarMensagem(token, chamadoId, texto) {
  const s = sessao_(token);
  texto = limparTexto_(texto);
  return comLock_(function () {
    const c = chamado_(chamadoId, s);
    if (s.papel === 'func' && c.status === 'encerrado') throw new Error('Este pedido foi encerrado. Abra um novo pedido.');
    const t = agora_().iso;
    inserirT_('Mensagens', { id: id_(), chamadoId: c.id, de: s.papel, texto: texto, em: t });
    const patch = { atualizadoEm: t };
    if (s.papel === 'rh') { patch.status = 'respondido'; patch.naoLidoFunc = '1'; patch.naoLidoRH = '0'; }
    else { patch.status = 'aberto'; patch.naoLidoRH = '1'; }
    atualizarT_('Chamados', c._row, patch);
    return chamadosDe_(c.matricula);
  });
}

/* =====================================================================
 * CONVERSA COM O RH (com respostas automáticas)
 * O funcionário tem uma conversa aberta por vez. Cada mensagem é comparada
 * com as respostas automáticas cadastradas pelo RH; sem resposta, vai para o RH.
 * humano = true: o funcionário pediu para falar com uma pessoa do RH.
 * ===================================================================== */
function funcChat(token, texto, humano) {
  const s = sessao_(token, 'func');
  texto = limparTexto_(texto);
  const cfg = config_();
  // 1) guarda a mensagem do funcionário
  const c = comLock_(function () {
    const c0 = conversaAberta_(s.matricula);
    inserirT_('Mensagens', { id: id_(), chamadoId: c0.id, de: 'func', texto: texto, em: new Date().toISOString() });
    return c0;
  });
  // 2) resposta cadastrada ou IA (fora do bloqueio, porque a IA leva alguns segundos)
  const regra = humano ? null : regraPara_(texto);
  let ia = null;
  if (!humano && !regra && cfg.iaAtiva) {
    try { ia = iaResponder_(s.f, texto, c.id, cfg); } catch (e) { console.warn('IA: ' + e.message); ia = null; }
  }
  return comLock_(function () {
    const c1 = lerT_('Chamados').filter(function (x) { return x.id === c.id; })[0];
    const t = new Date();
    const iso = function (ms) { return new Date(t.getTime() + ms).toISOString(); };
    const patch = { atualizadoEm: iso(1), status: 'aberto' };
    if (regra) {
      inserirT_('Mensagens', { id: id_(), chamadoId: c.id, de: 'bot', texto: regra.resposta, em: iso(1) });
      patch.naoLidoFunc = '1';
      if (regra.avisar === '1') patch.naoLidoRH = '1';
    } else if (ia && ia.resposta) {
      inserirT_('Mensagens', { id: id_(), chamadoId: c.id, de: 'bot', texto: String(ia.resposta).slice(0, 2000), em: iso(1) });
      patch.naoLidoFunc = '1';
      if (ia.encaminhar) patch.naoLidoRH = '1';
    } else {
      patch.naoLidoRH = '1';
      const msgs = lerT_('Mensagens').filter(function (m) { return m.chamadoId === c.id && m.de === 'bot' && m.texto.indexOf('encaminhei') >= 0; });
      const ult = msgs.length ? new Date(msgs[msgs.length - 1].em).getTime() : 0;
      if (t.getTime() - ult > 30 * 60000) {
        const at = config_().atendimento;
        inserirT_('Mensagens', { id: id_(), chamadoId: c.id, de: 'bot', texto: 'Recebi sua mensagem. Vou verificar e te respondo aqui mesmo' + (at ? ' (atendimento: ' + at + ').' : '.'), em: iso(1) });
        patch.naoLidoFunc = '1';
      }
    }
    atualizarT_('Chamados', c1._row, patch);
    return chamadosDe_(s.matricula);
  });
}

/* =====================================================================
 * ASSISTENTE COM IA (opcional) — Google Gemini, chave gratuita do Google AI Studio
 * Responde como o RH, só com as informações abaixo; encaminha ao RH o que pede análise.
 * ===================================================================== */
function iaResponder_(f, texto, chamadoId, cfg) {
  const hoje = agora_();
  const dias = ['domingo', 'segunda', 'terça', 'quarta', 'quinta', 'sexta', 'sábado'];
  const dSem = function (iso) { const p = iso.split('-').map(Number); return dias[new Date(Date.UTC(p[0], p[1] - 1, p[2], 12)).getUTCDay()]; };
  const br = function (iso) { const p = String(iso).slice(0, 10).split('-'); return p[2] + '/' + p[1] + '/' + p[0]; };
  const ano = Number(hoje.data.slice(0, 4));
  const fer = feriados_(ano).concat(feriados_(ano + 1)).filter(function (x) { return x.data >= hoje.data && x.data <= (ano + 1) + hoje.data.slice(4); });
  const bh = bhDe_(f.matricula);
  const ferias = feriasDe_(f.matricula).slice(0, 3);
  const docs = docsDe_(f.matricula).slice(0, 8);
  const regras = lerT_('Respostas').filter(function (r) { return r.ativo === '1'; });
  const hist = lerT_('Mensagens').filter(function (m) { return m.chamadoId === chamadoId; }).slice(-8);
  const nomeRH = cfg.rhNome || 'RH';
  const sistema = [
    'Você é ' + nomeRH + ', do RH da empresa ' + cfg.empresa + '. Você responde os funcionários pelo app Rhitmo Atende, como se fosse você mesmo escrevendo.',
    'Escreva em português do Brasil, em primeira pessoa, com tom cordial, simples e direto. No máximo 5 frases curtas. Pode cumprimentar pelo primeiro nome. Não use markdown, só texto.',
    'Use APENAS as informações fornecidas abaixo e conhecimento geral seguro (calendário, CLT em termos gerais). Nunca invente valores, datas de pagamento, saldos, regras ou benefícios da empresa que não estejam nas informações.',
    'Marque encaminhar=true (e diga que vai verificar e responder em breve) quando: a pergunta precisar de análise individual do RH (erro ou dúvida sobre valor do salário, descontos, ponto errado ou esquecido, atestado, advertência, rescisão, aumento, promoção, empréstimo, mudança de horário), envolver saúde, problema pessoal, assédio, conflito ou reclamação, pedir documento que não está no app, ou quando você não tiver certeza da resposta.',
    'Marque encaminhar=false quando a resposta estiver clara nas informações (ex.: feriados, datas de pagamento informadas, como usar o app, regras cadastradas).',
    'Nunca peça CPF, senha ou dados bancários. Para recibos, espelhos, banco de horas e férias, lembre que estão no próprio app (Documentos, Ponto, Perfil).',
    'Responda em JSON: {"resposta": texto, "encaminhar": true/false}.'
  ].join('\n');
  const ctx = [
    'HOJE: ' + br(hoje.data) + ' (' + dSem(hoje.data) + '), ' + hoje.hora.slice(0, 5) + ' em Brasília.',
    'HORÁRIO DE ATENDIMENTO DO RH: ' + (cfg.atendimento || 'não informado'),
    'FUNCIONÁRIO: ' + String(f.nome || '').split(' ')[0] + ' · setor ' + (f.setor || '-') + ' · cargo ' + (f.cargo || '-') + ' · empresa ' + (f.empresa || cfg.empresa),
    'HORÁRIO DO FUNCIONÁRIO: ' + (horarioTexto_(f.horario) || 'não cadastrado') + (props_().getProperty('SEXTA_1H') !== '0' ? '; às sextas sai 1 hora mais cedo' : ''),
    'BANCO DE HORAS: ' + (bh ? (bh.saldo >= 0 ? '+' : '-') + Math.floor(Math.abs(bh.saldo) / 60) + 'h' + ('0' + Math.abs(bh.saldo) % 60).slice(-2) + ' (referência ' + (bh.comp || '-') + ')' : 'não enviado'),
    'PEDIDOS DE FÉRIAS: ' + (ferias.length ? ferias.map(function (x) { return x.dias + ' dias a partir de ' + br(x.inicio) + ' (' + x.status + ')'; }).join('; ') : 'nenhum'),
    'DOCUMENTOS NO APP: ' + (docs.length ? docs.map(function (d) { return d.titulo; }).join('; ') : 'nenhum'),
    'FERIADOS DOS PRÓXIMOS 12 MESES: ' + fer.map(function (x) { return br(x.data) + ' ' + dSem(x.data) + ' - ' + x.nome; }).join('; '),
    'REGRAS E RESPOSTAS DO RH: ' + (regras.length ? regras.map(function (r) { return '[' + r.palavras + '] ' + r.resposta; }).join(' | ') : 'nenhuma'),
    'INFORMAÇÕES DA EMPRESA ESCRITAS PELO RH: ' + (props_().getProperty('IA_INFO') || 'nenhuma'),
    'CONVERSA RECENTE:\n' + hist.map(function (m) { return (m.de === 'func' ? 'Funcionário' : nomeRH) + ': ' + m.texto; }).join('\n')
  ].join('\n');
  return gemini_(sistema, ctx + '\n\nNOVA MENSAGEM DO FUNCIONÁRIO: ' + texto);
}
const IA_MODELO_PADRAO = 'gemini-3.8-flash';
/** Modelo em uso: os da linha 2.x saíram do ar para chaves novas (out/2026), então caem no padrão. */
function iaModelo_() {
  const m = String(props_().getProperty('IA_MODELO') || '').trim();
  return !m || /^gemini-(1|2)\./.test(m) ? IA_MODELO_PADRAO : m;
}
function gemini_(sistema, usuario) {
  const p = props_();
  const chave = p.getProperty('IA_CHAVE');
  if (!chave) throw new Error('A chave da IA não foi configurada.');
  const pedir = function (modelo, pensar) {
    const gc = {
      temperature: 0.3, maxOutputTokens: 2048, responseMimeType: 'application/json',
      responseSchema: { type: 'OBJECT', properties: { resposta: { type: 'STRING' }, encaminhar: { type: 'BOOLEAN' } }, required: ['resposta', 'encaminhar'] }
    };
    if (pensar) gc.thinkingConfig = /^gemini-2\./.test(modelo) ? { thinkingBudget: 0 } : { thinkingLevel: 'low' };   // pensa pouco: resposta mais rápida
    return UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(modelo) + ':generateContent', {
      method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': chave },
      payload: JSON.stringify({ systemInstruction: { parts: [{ text: sistema }] }, contents: [{ role: 'user', parts: [{ text: usuario }] }], generationConfig: gc }),
      muteHttpExceptions: true
    });
  };
  let modelo = iaModelo_();
  let r = pedir(modelo, true);
  if (r.getResponseCode() === 400 && /thinking/i.test(r.getContentText())) r = pedir(modelo, false);          // modelo não aceita o ajuste de raciocínio
  if (r.getResponseCode() === 404 && modelo !== 'gemini-flash-latest') { modelo = 'gemini-flash-latest'; r = pedir(modelo, true); }   // modelo saiu do ar
  if (r.getResponseCode() === 400 && /thinking/i.test(r.getContentText())) r = pedir(modelo, false);
  const cod = r.getResponseCode();
  if (cod !== 200) {
    const t = r.getContentText();
    if (cod === 400 && /API key not valid|API_KEY_INVALID/i.test(t)) throw new Error('A chave da IA não é válida. Copie de novo a chave no AI Studio (começa com AIza) e cole aqui.');
    if (cod === 403) throw new Error('A chave da IA não tem permissão. Confira se ela foi criada no AI Studio com a conta certa.');
    if (cod === 429) throw new Error('A cota gratuita da IA acabou por agora. Ela volta sozinha em alguns minutos (ou no dia seguinte).');
    throw new Error('A IA respondeu com erro ' + cod + ': ' + t.slice(0, 300));
  }
  const j = JSON.parse(r.getContentText());
  const partes = (((j.candidates || [])[0] || {}).content || {}).parts || [];
  const txt = partes.filter(function (x) { return !x.thought; }).map(function (x) { return x.text || ''; }).join('');
  const o = JSON.parse(txt);
  return { resposta: String(o.resposta || '').trim(), encaminhar: !!o.encaminhar };
}
/** Feriados usados pela IA: os mesmos do Rhitmo (nacionais, Sexta-feira Santa e 20/11) + extras do RH. */
function feriados_(y) {
  const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3),
    h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, mm = Math.floor((a + 11 * h + 22 * l) / 451),
    mes = Math.floor((h + l - 7 * mm + 114) / 31), dia = ((h + l - 7 * mm + 114) % 31) + 1;
  const pascoa = new Date(Date.UTC(y, mes - 1, dia));
  const iso = function (dt) { return dt.toISOString().slice(0, 10); };
  const desloca = function (n) { return iso(new Date(pascoa.getTime() + n * 86400000)); };
  const fx = function (m, d) { return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2); };
  const l2 = [[fx(1, 1), 'Confraternização Universal'], [desloca(-2), 'Sexta-feira Santa'], [fx(4, 21), 'Tiradentes'], [fx(5, 1), 'Dia do Trabalho'],
    [fx(9, 7), 'Independência'], [fx(10, 12), 'Nossa Senhora Aparecida'], [fx(11, 2), 'Finados'], [fx(11, 15), 'Proclamação da República'],
    [fx(11, 20), 'Dia da Consciência Negra'], [fx(12, 25), 'Natal']];
  String(props_().getProperty('FERIADOS_EXTRAS') || '').split('\n').forEach(function (lin) {
    const m = lin.trim().match(/^(\d{1,2})\/(\d{1,2})\s*[-–]?\s*(.*)$/);
    if (m) l2.push([fx(Number(m[2]), Number(m[1])), (m[3] || 'Feriado').trim()]);
  });
  return l2.map(function (x) { return { data: x[0], nome: x[1] }; }).sort(function (p1, p2) { return p1.data < p2.data ? -1 : 1; });
}
function rhIA(token) {
  sessao_(token, 'rh');
  const p = props_();
  return { ativa: p.getProperty('IA_ATIVA') === '1', temChave: !!p.getProperty('IA_CHAVE'), modelo: iaModelo_(),
    info: p.getProperty('IA_INFO') || '', extras: p.getProperty('FERIADOS_EXTRAS') || '', rhNome: p.getProperty('RH_NOME') || 'Fabiano Lourenço', avisoAuto: p.getProperty('AVISO_AUTO') !== '0' };
}
function rhSalvarIA(token, d) {
  sessao_(token, 'rh');
  const p = props_();
  const v = {
    IA_ATIVA: d.ativa ? '1' : '0', IA_MODELO: String(d.modelo || IA_MODELO_PADRAO).trim().slice(0, 60),
    IA_INFO: String(d.info || '').slice(0, 8000), FERIADOS_EXTRAS: String(d.extras || '').slice(0, 2000),
    RH_NOME: String(d.rhNome || '').trim().slice(0, 60) || 'RH', AVISO_AUTO: d.avisoAuto ? '1' : '0'
  };
  if (d.chave) v.IA_CHAVE = String(d.chave).trim();
  if (d.ativa && !v.IA_CHAVE && !p.getProperty('IA_CHAVE')) throw new Error('Cole a chave da IA antes de ativar.');
  p.setProperties(v);
  return rhIA(token);
}
/** Testa a IA com uma pergunta, como se fosse um funcionário (usa dados fictícios). */
function rhTestarIA(token, pergunta) {
  sessao_(token, 'rh');
  const f = { matricula: '__teste__', nome: 'Funcionário Teste', setor: 'Produção', cargo: 'Auxiliar', empresa: config_().empresa };
  return iaResponder_(f, limparTexto_(pergunta), '__teste__', config_());
}
function conversaAberta_(mat) {
  const abertas = lerT_('Chamados').filter(function (c) { return c.matricula === mat && c.status !== 'encerrado'; })
    .sort(function (a, b) { return a.atualizadoEm < b.atualizadoEm ? 1 : -1; });
  if (abertas.length) return abertas[0];
  const t = agora_().iso;
  const novo = { id: id_(), matricula: mat, assunto: 'Conversa', status: 'aberto', criadoEm: t, atualizadoEm: t, naoLidoRH: '0', naoLidoFunc: '0' };
  inserirT_('Chamados', novo);
  return lerT_('Chamados').filter(function (c) { return c.id === novo.id; })[0];
}
function normTxt_(t) { return String(t || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim(); }
function regraPara_(texto) {
  const t = ' ' + normTxt_(texto) + ' ';
  let melhor = null, pontos = 0;
  lerT_('Respostas').forEach(function (r) {
    if (r.ativo !== '1') return;
    const n = String(r.palavras || '').split(/[,;\n]+/).map(normTxt_).filter(Boolean)
      .filter(function (p) { return t.indexOf(' ' + p) >= 0; }).length;
    if (n > pontos) { pontos = n; melhor = r; }
  });
  return melhor;
}
/** Mensagem automática (do sistema) dentro da conversa do funcionário. */
function avisoNaConversa_(mat, texto, de, paraRH) {
  const c = conversaAberta_(mat);
  const t = agora_().iso;
  inserirT_('Mensagens', { id: id_(), chamadoId: c.id, de: de || 'bot', texto: texto, em: t });
  const patch = { atualizadoEm: t, naoLidoFunc: '1' };
  if (paraRH) { patch.naoLidoRH = '1'; patch.status = 'aberto'; }
  atualizarT_('Chamados', c._row, patch);
}

/* ----- férias ----- */
function funcFerias(token, d) {
  const s = sessao_(token, 'func');
  const ini = String(d && d.inicio || '');
  const dias = Number(d && d.dias);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ini)) throw new Error('Escolha a data de início.');
  if (ini <= agora_().data) throw new Error('A data de início precisa ser no futuro.');
  if (!(dias >= 5 && dias <= 30)) throw new Error('As férias precisam ter entre 5 e 30 dias.');
  const abono = d.abono ? '1' : '0';
  if (abono === '1' && dias > 20) throw new Error('Para vender 10 dias, tire no máximo 20 dias de descanso.');
  const obs = String(d.obs || '').trim().slice(0, 500);
  return comLock_(function () {
    const pend = lerT_('Ferias').filter(function (f) { return f.matricula === s.matricula && f.status === 'pendente'; });
    if (pend.length >= 3) throw new Error('Você já tem 3 pedidos aguardando o RH. Aguarde a resposta.');
    inserirT_('Ferias', { id: id_(), matricula: s.matricula, inicio: ini, dias: String(dias), abono: abono, obs: obs, status: 'pendente', criadoEm: agora_().iso });
    const p = ini.split('-');
    avisoNaConversa_(s.matricula, 'Recebi seu pedido de férias: ' + dias + ' dias a partir de ' + p[2] + '/' + p[1] + '/' + p[0] + (abono === '1' ? ', vendendo 10 dias' : '') + '. Vou analisar e te respondo aqui.', 'bot', true);
    return { ferias: feriasDe_(s.matricula), chamados: chamadosDe_(s.matricula) };
  });
}
function feriasDe_(mat) {
  return lerT_('Ferias').filter(function (f) { return f.matricula === mat; }).map(function (f) {
    return { id: f.id, inicio: f.inicio, dias: Number(f.dias), abono: f.abono === '1', obs: f.obs, status: f.status, criadoEm: f.criadoEm, respostaRH: f.respostaRH, respondidoEm: f.respondidoEm };
  }).sort(function (a, b) { return a.criadoEm < b.criadoEm ? 1 : -1; });
}

/* ----- foto de perfil ----- */
function salvarFoto(token, b64) {
  const s = sessao_(token, 'func');
  b64 = String(b64 || '');
  if (b64 && (b64.length > 48000 || !/^[A-Za-z0-9+/=]+$/.test(b64))) throw new Error('Foto grande demais. Tente outra imagem.');
  return comLock_(function () {
    const f = lerT_('Fotos').filter(function (x) { return x.matricula === s.matricula; })[0];
    if (f) { if (b64) atualizarT_('Fotos', f._row, { foto: b64, atualizadoEm: agora_().iso }); else apagarLinha_('Fotos', f._row); }
    else if (b64) inserirT_('Fotos', { matricula: s.matricula, foto: b64, atualizadoEm: agora_().iso });
    return fotoDe_(s.matricula);
  });
}
function fotoDe_(mat) {
  const sh = ss_().getSheetByName('Fotos');
  if (!sh || sh.getLastRow() < 2) return '';
  const cel = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(String(mat)).matchEntireCell(true).findNext();
  return cel ? sh.getRange(cel.getRow(), 2).getDisplayValue() : '';
}

function marcarLido(token, chamadoId) {
  const s = sessao_(token);
  return comLock_(function () {
    const c = chamado_(chamadoId, s);
    atualizarT_('Chamados', c._row, s.papel === 'rh' ? { naoLidoRH: '0' } : { naoLidoFunc: '0' });
    return true;
  });
}

function baixarDoc(token, docId) {
  const s = sessao_(token);
  const d = lerT_('Documentos').filter(function (x) { return x.id === docId; })[0];
  if (!d || (s.papel === 'func' && d.matricula !== s.matricula)) throw new Error('Documento não encontrado.');
  const blob = DriveApp.getFileById(d.fileId).getBlob();
  return { nome: nomeArquivo_(d.titulo) + '.pdf', b64: Utilities.base64Encode(blob.getBytes()) };
}

/* =====================================================================
 * RH
 * ===================================================================== */
function rhPainel(token) {
  sessao_(token, 'rh');
  const t = agora_();
  const pend = {};
  lerT_('Chamados').forEach(function (c) { if (c.naoLidoRH === '1') pend[c.matricula] = (pend[c.matricula] || 0) + 1; });
  const ultimo = {};
  lerPonto_(t.mes).forEach(function (p) { if (p.data === t.data) ultimo[p.matricula] = p.hora; });
  const ferPend = {};
  lerT_('Ferias').forEach(function (f) { if (f.status === 'pendente') ferPend[f.matricula] = (ferPend[f.matricula] || 0) + 1; });
  const funcs = lerT_('Funcionarios').map(function (f) {
    return {
      matricula: f.matricula, nome: f.nome, setor: f.setor, cargo: f.cargo, empresa: f.empresa, usuario: f.usuario, horario: f.horario, cpf: f.cpf ? '***.' + f.cpf.slice(3, 6) + '.***-' + f.cpf.slice(9) : '', temPis: !!f.pis,
      ativo: f.ativo === '1', trocar: f.trocarSenha === '1',
      pendentes: pend[f.matricula] || 0, pontoHoje: ultimo[f.matricula] || '', feriasPend: ferPend[f.matricula] || 0
    };
  });
  return { cfg: config_(), agora: t, funcs: funcs, horarios: Object.keys(horarios_()), modelos: MODELOS_ESPELHO };
}

function rhFuncionario(token, matricula, mes) {
  sessao_(token, 'rh');
  if (!/^\d{4}-\d{2}$/.test(String(mes))) mes = agora_().mes;
  return { ponto: pontoDe_(matricula, mes), chamados: chamadosDe_(matricula), docs: docsDe_(matricula), bh: bhDe_(matricula), ferias: feriasDe_(matricula), foto: fotoDe_(matricula) };
}

/** Cria (novo=true) ou edita um funcionário. Ao criar, devolve a senha provisória. */
function rhSalvarFuncionario(token, d, novo) {
  sessao_(token, 'rh');
  const mat = String(d.matricula || '').trim();
  const nome = String(d.nome || '').trim().slice(0, 120);
  if (!/^[A-Za-z0-9._-]{1,20}$/.test(mat)) throw new Error('Matrícula inválida. Use só letras, números, ponto, hífen ou sublinhado.');
  if (mat.toLowerCase() === 'rh') throw new Error('"rh" é reservado para o login do RH.');
  if (nome.length < 3) throw new Error('Digite o nome completo.');
  return comLock_(function () {
    const lista = lerT_('Funcionarios');
    const atual = lista.filter(function (x) { return x.matricula === mat; })[0];
    const dados = { nome: nome, setor: String(d.setor || '').trim().slice(0, 60), cargo: String(d.cargo || '').trim().slice(0, 60) };
    if (d.horario != null) dados.horario = String(d.horario).trim().slice(0, 30);
    if (d.cpf != null) dados.cpf = cpfNum_(d.cpf);
    if (d.pis != null) dados.pis = pisNum_(d.pis);
    const usados = {};
    lista.forEach(function (x) { if (x.usuario && x.matricula !== mat) usados[x.usuario] = 1; });
    if (d.usuario) {
      const u = String(d.usuario).trim().toLowerCase();
      if (!/^[a-z0-9._-]{3,30}$/.test(u) || /^\d+$/.test(u)) throw new Error('Usuário inválido: use letras sem acento, números, ponto, hífen ou sublinhado.');
      if (usados[u]) throw new Error('Esse usuário já está em uso.');
      dados.usuario = u;
    }
    if (novo) {
      if (atual) throw new Error('Já existe um funcionário com a matrícula ' + mat + '.');
      const senha = gerarSenha_();
      const salt = gerarSalt_();
      const usuario = dados.usuario || gerarUsuario_(nome, usados);
      inserirT_('Funcionarios', Object.assign({}, dados, {
        matricula: mat, hash: hash_(senha, salt), salt: salt, ativo: '1', trocarSenha: '1', criadoEm: agora_().iso,
        empresa: String(d.empresa || '').trim().slice(0, 60), usuario: usuario
      }));
      return { senha: senha, usuario: usuario };
    }
    if (!atual) throw new Error('Funcionário não encontrado.');
    if (!atual.usuario && !dados.usuario) dados.usuario = gerarUsuario_(nome, usados);
    atualizarT_('Funcionarios', atual._row, dados);
    return { ok: true };
  });
}

function rhResetSenha(token, matricula) {
  sessao_(token, 'rh');
  return comLock_(function () {
    const f = lerT_('Funcionarios').filter(function (x) { return x.matricula === matricula; })[0];
    if (!f) throw new Error('Funcionário não encontrado.');
    const senha = gerarSenha_();
    const salt = gerarSalt_();
    atualizarT_('Funcionarios', f._row, { hash: hash_(senha, salt), salt: salt, trocarSenha: '1' });
    CacheService.getScriptCache().remove('falha_' + String(matricula).toLowerCase());
    return { senha: senha, usuario: f.usuario };
  });
}

function rhAtivar(token, matricula, ativo) {
  sessao_(token, 'rh');
  return comLock_(function () {
    const f = lerT_('Funcionarios').filter(function (x) { return x.matricula === matricula; })[0];
    if (!f) throw new Error('Funcionário não encontrado.');
    atualizarT_('Funcionarios', f._row, { ativo: ativo ? '1' : '0' });
    return true;
  });
}

function rhEncerrar(token, chamadoId) {
  const s = sessao_(token, 'rh');
  return comLock_(function () {
    const c = chamado_(chamadoId, s);
    atualizarT_('Chamados', c._row, { status: 'encerrado', naoLidoRH: '0', atualizadoEm: agora_().iso });
    return chamadosDe_(c.matricula);
  });
}

/**
 * Importa a lista de funcionários do Rhitmo (planilha 01_Funcionarios.xlsx).
 * itens = [{matricula, nome, setor, cargo, empresa, ativo}]
 * Cria quem é novo e está ativo (com senha provisória), atualiza nome/setor/função
 * de quem já existe e desativa quem tem data de demissão. Senhas existentes não mudam.
 */
function rhImportar(token, itens) {
  sessao_(token, 'rh');
  if (!Array.isArray(itens) || !itens.length) throw new Error('A lista está vazia.');
  if (itens.length > 2000) throw new Error('Lista grande demais. Importe no máximo 2000 por vez.');
  return comLock_(function () {
    const cab = ABAS.Funcionarios;
    const sh = aba_('Funcionarios', cab);
    const lista = lerT_('Funcionarios');
    const porMat = {};
    lista.forEach(function (f) { porMat[f.matricula] = f; });
    const novos = [], senhas = [];
    const usados = {};
    lista.forEach(function (x) { if (x.usuario) usados[x.usuario] = 1; });
    let atualizados = 0, desativados = 0, reativados = 0, ignorados = 0, cpfs = 0;
    itens.forEach(function (it) {
      const mat = String(it.matricula || '').trim();
      const nome = String(it.nome || '').trim().slice(0, 120);
      if (!/^[A-Za-z0-9._-]{1,20}$/.test(mat) || mat.toLowerCase() === 'rh' || nome.length < 3) { ignorados++; return; }
      const dados = {
        nome: nome, setor: String(it.setor || '').trim().slice(0, 60),
        cargo: String(it.cargo || '').trim().slice(0, 60), empresa: String(it.empresa || '').trim().slice(0, 60),
        cpf: cpfNum_(it.cpf), horario: String(it.horario || '').trim().slice(0, 30), pis: pisNum_(it.pis)
      };
      if (!dados.cpf && porMat[mat]) dados.cpf = porMat[mat].cpf || '';
      if (!dados.pis && porMat[mat]) dados.pis = porMat[mat].pis || '';
      if (dados.cpf && (!porMat[mat] || porMat[mat].cpf !== dados.cpf)) cpfs++;
      if (!dados.horario && porMat[mat]) dados.horario = porMat[mat].horario || '';
      const f = porMat[mat];
      if (f) {
        const mudou = ['nome', 'setor', 'cargo', 'empresa', 'cpf', 'horario', 'pis'].some(function (k) { return f[k] !== dados[k]; });
        Object.assign(f, dados);
        if (!f.usuario) f.usuario = gerarUsuario_(f.nome, usados);
        if (!it.ativo && f.ativo === '1') { f.ativo = '0'; desativados++; }
        else if (it.ativo && f.ativo !== '1' && it.reativar) { f.ativo = '1'; reativados++; }
        else if (mudou) atualizados++;
        return;
      }
      if (!it.ativo) { ignorados++; return; }
      const senha = gerarSenha_();
      const salt = gerarSalt_();
      const novo = Object.assign({ matricula: mat, hash: hash_(senha, salt), salt: salt, ativo: '1', trocarSenha: '1', criadoEm: agora_().iso, usuario: gerarUsuario_(nome, usados) }, dados);
      porMat[mat] = novo;
      novos.push(novo);
      senhas.push({ matricula: mat, nome: nome, empresa: dados.empresa, senha: senha, usuario: novo.usuario });
    });
    if (lista.length) {
      const r = sh.getRange(2, 1, lista.length, cab.length);
      r.setNumberFormat('@');
      r.setValues(lista.map(function (f) { return cab.map(function (c) { return f[c] == null ? '' : String(f[c]); }); }));
    }
    if (novos.length) {
      const r2 = sh.getRange(lista.length + 2, 1, novos.length, cab.length);
      r2.setNumberFormat('@');
      r2.setValues(novos.map(function (f) { return cab.map(function (c) { return f[c] == null ? '' : String(f[c]); }); }));
    }
    sujo_('Funcionarios');
    const todos = lista.concat(novos).filter(function (f) { return f.ativo === '1'; });
    return { criados: novos.length, atualizados: atualizados, desativados: desativados, reativados: reativados, ignorados: ignorados, senhas: senhas,
      cpfs: cpfs, semCpf: todos.filter(function (f) { return !f.cpf; }).map(function (f) { return f.nome; }) };
  });
}

/* =====================================================================
 * EXPORTAÇÃO AFD (Portaria MTP 671/2021, Anexo V) — para importar no Dixi
 * Registros: 1 (cabeçalho), 3 (marcação), 9 (trailer) e linha de assinatura.
 * ===================================================================== */
function proximoNSR_(empresa) {
  const k = 'NSR_' + empresaChave_(empresa);
  const p = props_();
  const n = Number(p.fresco(k) || 0) + 1;
  p.setProperty(k, String(n));
  return n;
}
function empresaChave_(e) { return String(e || 'MEPCO').toUpperCase().replace(/[^A-Z0-9]/g, '') || 'MEPCO'; }
function crc16Kermit_(str) {
  let crc = 0;
  for (let i = 0; i < str.length; i++) {
    crc ^= str.charCodeAt(i) & 0xFF;
    for (let b = 0; b < 8; b++) crc = (crc & 1) ? ((crc >>> 1) ^ 0x8408) : (crc >>> 1);
  }
  return ('0000' + crc.toString(16).toUpperCase()).slice(-4);
}
function afdTxt_(v, n) { v = String(v == null ? '' : v); return (v + ' '.repeat(n)).slice(0, n); }
function afdNum_(v, n) { v = String(v == null ? '' : v).replace(/\D/g, ''); return ('0'.repeat(n) + v).slice(-n); }
function afdAscii_(v) { return String(v || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\x20-\x7E]/g, ' ').toUpperCase(); }

/**
 * Gera o AFD de uma empresa num período.
 * emp = {nome:'MEPCO', cnpj:'08182147000151', razao:'...'}; ini/fim = 'AAAA-MM-DD'
 */
function rhExportarAFD(token, emp, ini, fim) {
  sessao_(token, 'rh');
  return gerarAFD_(emp, ini, fim);
}
function gerarAFD_(emp, ini, fim) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ini) || !/^\d{4}-\d{2}-\d{2}$/.test(fim) || ini > fim) throw new Error('Período inválido.');
  const cnpj = String(emp && emp.cnpj || '').replace(/\D/g, '');
  if (cnpj.length !== 14) throw new Error('CNPJ da empresa inválido.');
  const chave = empresaChave_(emp.nome);
  return comLock_(function () {
    const funcs = {};
    lerT_('Funcionarios').forEach(function (f) { if (empresaChave_(f.empresa) === chave) funcs[f.matricula] = f; });
    const meses = [];
    let m = ini.slice(0, 7);
    while (m <= fim.slice(0, 7)) {
      meses.push(m);
      const y = Number(m.slice(0, 4)), mm = Number(m.slice(5, 7));
      m = mm === 12 ? (y + 1) + '-01' : y + '-' + ('0' + (mm + 1)).slice(-2);
    }
    const marc = [], semCpf = {};
    meses.forEach(function (mes) {
      lerPonto_(mes).forEach(function (p) {
        const f = funcs[p.matricula];
        if (!f || p.data < ini || p.data > fim) return;
        if (p.origem && p.origem !== 'app') return;   // marcações que vieram do Dixi/Rhitmo já estão lá
        if (!p.nsr) {   // batidas antigas, de antes do AFD: recebem número agora e guardam
          p.nsr = String(proximoNSR_(f.empresa));
          atualizarPonto_(mes, p._row, p.nsr);
        }
        const cpf = String(f.cpf || '').replace(/\D/g, '');
        if (cpf.length !== 11) { semCpf[f.matricula] = f.nome; return; }
        marc.push({ nsr: Number(p.nsr), dh: p.data + 'T' + p.hora.slice(0, 5) + ':00-0300', cpf: cpf });
      });
    });
    marc.sort(function (a, b) { return a.nsr - b.nsr; });
    const agora = new Date();
    const gerado = Utilities.formatDate(agora, TZ, 'yyyy-MM-dd') + 'T' + Utilities.formatDate(agora, TZ, 'HH:mm') + ':00-0300';
    let cab = afdNum_(0, 9) + '1' + '1' + afdTxt_(cnpj, 14) + ' '.repeat(14) + afdTxt_(afdAscii_(emp.razao || emp.nome), 150) +
      '9'.repeat(17) + ini + fim + gerado + '004' + '1' + afdTxt_(cnpj, 14) + ' '.repeat(30);
    cab += crc16Kermit_(cab);
    const linhas = [cab];
    marc.forEach(function (x) {
      let l = afdNum_(x.nsr, 9) + '3' + x.dh + afdNum_(x.cpf, 12);
      linhas.push(l + crc16Kermit_(l));
    });
    linhas.push('9'.repeat(9) + afdNum_(0, 9) + afdNum_(marc.length, 9) + afdNum_(0, 9) + afdNum_(0, 9) + afdNum_(0, 9) + afdNum_(0, 9) + '9');
    linhas.push(afdTxt_('ASSINATURA_DIGITAL_EM_ARQUIVO_P7S', 100));
    const faltam = Object.keys(semCpf).map(function (k) { return { matricula: k, nome: semCpf[k] }; });
    return { conteudo: linhas.join('\r\n') + '\r\n', marcacoes: marc.length, semCpf: faltam, nome: 'AFD' + cnpj + 'REP_A.txt' };
  });
}
function atualizarPonto_(mes, row, nsr) {
  const sh = pontoAba_(mes);
  const r = sh.getRange(row, CAB_PONTO.indexOf('nsr') + 1);
  r.setNumberFormat('@');
  r.setValue(String(nsr));
  sujo_('Ponto ' + mes);
}

/* =====================================================================
 * IMPORTAÇÃO DE MARCAÇÕES (AFD do Dixi ou 03_Ponto.xlsx do Rhitmo)
 * itens = [{id:'CPF ou PIS', mat:'matrícula', data:'AAAA-MM-DD', hora:'HH:MM'}]
 * Marcações repetidas (mesmo funcionário, dia e minuto) são ignoradas, então
 * dá para importar o mesmo arquivo de novo sem duplicar nada.
 * ===================================================================== */
function cpfNum_(v) { let d = String(v == null ? '' : v).replace(/\.0+$/, '').replace(/\D/g, ''); if (d.length > 11 && /^0+$/.test(d.slice(0, d.length - 11))) d = d.slice(-11); return d.length >= 9 && d.length <= 11 ? ('00' + d).slice(-11) : ''; }
function pisNum_(v) { const d = String(v == null ? '' : v).replace(/\.0+$/, '').replace(/\D/g, ''); return d.length >= 10 && d.length <= 12 ? ('00' + d).slice(-11) : ''; }
function rhImportarMarcacoes(token, itens, origem, simular) {
  sessao_(token, 'rh');
  if (!Array.isArray(itens) || !itens.length) throw new Error('Nenhuma marcação no arquivo.');
  if (itens.length > 60000) throw new Error('Arquivo grande demais. Escolha um período menor.');
  origem = origem === 'rhitmo' ? 'rhitmo' : 'dixi';
  const run = function () {
    const porCpf = {}, porPis = {}, porMat = {};
    lerT_('Funcionarios').forEach(function (f) {
      porMat[f.matricula] = f;
      if (f.cpf) porCpf[f.cpf] = porCpf[f.cpf] && porCpf[f.cpf].ativo === '1' ? porCpf[f.cpf] : f;
      if (f.pis) porPis[f.pis] = f;
    });
    const existe = {}, novasPorMes = {}, funcs = {}, naoAchados = {};
    let novas = 0, repetidas = 0, invalidas = 0, ini = '', fim = '';
    const chaves = function (mes) {
      if (!existe[mes]) {
        existe[mes] = {};
        lerPonto_(mes).forEach(function (p) { existe[mes][p.matricula + '|' + p.data + '|' + String(p.hora).slice(0, 5)] = 1; });
      }
      return existe[mes];
    };
    itens.forEach(function (it) {
      const data = String(it.data || ''), hora = String(it.hora || '').slice(0, 5);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(data) || !/^([01]\d|2[0-3]):[0-5]\d$/.test(hora)) { invalidas++; return; }
      let f = null, ref = '';
      if (it.mat) { f = porMat[String(it.mat)]; ref = 'Matrícula ' + String(it.mat).replace(/^L2-/, ''); }
      else {
        const d = String(it.id || '').replace(/\D/g, '');
        const c = cpfNum_(d), pi = pisNum_(d);
        f = (it.tipo !== 'pis' && porCpf[c]) || porPis[pi] || porCpf[c] || null;
        ref = (it.tipo === 'pis' ? 'PIS ' : 'CPF ') + (it.tipo === 'pis' ? pi : c.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4'));
      }
      if (!f) { naoAchados[ref] = (naoAchados[ref] || 0) + 1; return; }
      const mes = data.slice(0, 7), k = f.matricula + '|' + data + '|' + hora, ch = chaves(mes);
      if (ch[k]) { repetidas++; return; }
      ch[k] = 1;
      novas++;
      funcs[f.matricula] = 1;
      if (!ini || data < ini) ini = data;
      if (!fim || data > fim) fim = data;
      const em = new Date(data + 'T' + hora + ':00-03:00').toISOString();
      (novasPorMes[mes] = novasPorMes[mes] || []).push([id_(), f.matricula, data, hora + ':00', em, '', origem]);
    });
    if (!simular) {
      Object.keys(novasPorMes).forEach(function (mes) {
        const l = novasPorMes[mes], sh = pontoAba_(mes);
        const r = sh.getRange(sh.getLastRow() + 1, 1, l.length, CAB_PONTO.length);
        r.setNumberFormat('@');
        r.setValues(l);
        sujo_('Ponto ' + mes);
      });
    }
    return {
      novas: novas, repetidas: repetidas, invalidas: invalidas, funcionarios: Object.keys(funcs).length, ini: ini, fim: fim,
      naoAchados: Object.keys(naoAchados).sort().map(function (k) { return { ref: k, n: naoAchados[k] }; })
    };
  };
  return simular ? run() : comLock_(run);
}

/* =====================================================================
 * HORÁRIOS DE TRABALHO (mesmos códigos do Rhitmo) E ESPELHO DE PONTO DO APP
 * ===================================================================== */
const HORARIOS_PADRAO = [
  ['0630-1630/11', '06:30 às 16:30, intervalo 11:00–12:00', '06:30', '11:00', '12:00', '16:30', 'Seg a Sex', '44'],
  ['07-17/11', '07:00 às 17:00, intervalo 11:00–12:00', '07:00', '11:00', '12:00', '17:00', 'Seg a Sex', '44'],
  ['07-17/1130', '07:00 às 17:00, intervalo 11:30–12:30', '07:00', '11:30', '12:30', '17:00', 'Seg a Sex', '44'],
  ['07-17/12', '07:00 às 17:00, intervalo 12:00–13:00', '07:00', '12:00', '13:00', '17:00', 'Seg a Sex', '44'],
  ['08-18/11', '08:00 às 18:00, intervalo 11:00–12:00', '08:00', '11:00', '12:00', '18:00', 'Seg a Sex', '44'],
  ['08-18/1130', '08:00 às 18:00, intervalo 11:30–12:30', '08:00', '11:30', '12:30', '18:00', 'Seg a Sex', '44'],
  ['08-18/12', '08:00 às 18:00, intervalo 12:00–13:00', '08:00', '12:00', '13:00', '18:00', 'Seg a Sex', '44'],
  ['6H 07-13', 'Jornada de 6h 07:00 às 13:00 (aprendiz/estágio)', '07:00', '', '', '13:00', 'Seg a Sex', '30'],
  ['6H 12-18', 'Jornada de 6h 12:00 às 18:00 (aprendiz/estágio)', '12:00', '', '', '18:00', 'Seg a Sex', '30'],
  ['HE 07-17', 'Motorista/ajudante 07:00 às 17:00 (com horas extras)', '07:00', '', '', '17:00', 'Seg a Sex', '44'],
  ['AFASTADO', 'Afastado pelo INSS (sem jornada)', '', '', '', '', '—', '0'],
  ['13-18/15', '13:00 às 18:00, pausa 15:00–15:15 (estágio)', '13:00', '15:00', '15:15', '18:00', 'Seg a Sex', '23.75']
];
function horarios_() {
  const l = lerT_('Horarios');
  const fonte = l.length ? l : HORARIOS_PADRAO.map(function (x) { return { codigo: x[0], descricao: x[1], ent: x[2], si: x[3], ri: x[4], sai: x[5], dias: x[6], carga: x[7] }; });
  const m = {};
  fonte.forEach(function (h) { m[h.codigo] = h; });
  return m;
}
function horarioTexto_(cod) {
  if (!cod) return '';
  const h = horarios_()[cod];
  return h ? h.descricao + (h.dias && h.dias !== '—' ? ' (' + h.dias + ')' : '') : cod;
}
function rhImportarHorarios(token, itens) {
  sessao_(token, 'rh');
  return comLock_(function () {
    const cab = ABAS.Horarios, sh = aba_('Horarios', cab);
    const l = (itens || []).filter(function (x) { return x && x.codigo; }).map(function (x) { return cab.map(function (c) { return String(x[c] == null ? '' : x[c]).slice(0, 80); }); });
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, cab.length).clearContent();
    if (l.length) { const r = sh.getRange(2, 1, l.length, cab.length); r.setNumberFormat('@'); r.setValues(l); }
    sujo_('Horarios');
    return l.length;
  });
}
const MODELOS_ESPELHO = [['1', 'Clássico'], ['2', 'Compacto'], ['3', 'Moderno'], ['4', 'Calendário'], ['5', 'Semanal']];
function minHM_(t) { if (!t) return null; const p = String(t).split(':').map(Number); return p[0] * 60 + (p[1] || 0); }
function hmTxt_(m, sinal) { if (m == null) return ''; const neg = m < 0; m = Math.abs(Math.round(m)); return (sinal ? (neg ? '-' : m ? '+' : '') : (neg ? '-' : '')) + Math.floor(m / 60) + ':' + ('0' + (m % 60)).slice(-2); }
function esc_(t) { return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
/** Calcula o espelho do mês a partir das batidas feitas no app. */
function espelhoDados_(f, mes) {
  const hoje = agora_().data;
  const y = Number(mes.slice(0, 4)), m = Number(mes.slice(5, 7));
  const ult = new Date(Date.UTC(y, m, 0)).getUTCDate();
  const h = horarios_()[f.horario] || null;
  const sexta1h = props_().getProperty('SEXTA_1H') !== '0';
  const tol = 10;
  const dSem = String(h && h.dias || 'Seg a Sex').toLowerCase();
  const diasTrab = { 1: 1, 2: 1, 3: 1, 4: 1, 5: 1 };
  if (/s[aá]b/.test(dSem)) diasTrab[6] = 1;
  if (/dom/.test(dSem)) diasTrab[0] = 1;
  const fer = {};
  feriados_(y).forEach(function (x) { fer[x.data] = x.nome; });
  const batidas = {};
  pontoDe_(f.matricula, mes).forEach(function (p) { (batidas[p.data] = batidas[p.data] || []).push(p.hora.slice(0, 5)); });
  const NOMES = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sáb'];
  const dias = [];
  const tot = { trab: 0, prev: 0, extra: 0, falta: 0, faltas: 0, diasTrab: 0, impares: 0 };
  for (let d = 1; d <= ult; d++) {
    const iso = mes + '-' + ('0' + d).slice(-2);
    const dow = new Date(Date.UTC(y, m - 1, d, 12)).getUTCDay();
    const b = (batidas[iso] || []).slice().sort();
    let prev = 0;
    if (h && h.ent && h.sai && diasTrab[dow] && !fer[iso]) {
      prev = minHM_(h.sai) - minHM_(h.ent) - (h.si && h.ri ? minHM_(h.ri) - minHM_(h.si) : 0);
      if (dow === 5 && sexta1h && h.si && !/^HE/i.test(h.codigo)) prev -= 60;
    }
    let trab = 0;
    for (let i = 0; i + 1 < b.length; i += 2) trab += minHM_(b[i + 1]) - minHM_(b[i]);
    const passado = iso < hoje, futuro = iso > hoje;
    let saldo = null, oc = '';
    if (fer[iso]) oc = 'Feriado: ' + fer[iso];
    else if (!diasTrab[dow]) oc = dow === 6 ? 'Sábado' : 'Domingo';
    const emAndamento = iso === hoje && (b.length < 4 || b.length % 2);   // o dia de hoje só fecha com as 4 batidas
    if (!futuro && !emAndamento) {
      saldo = trab - prev;
      if (Math.abs(saldo) <= tol) saldo = 0;
    }
    if (emAndamento) oc = oc || 'Em andamento';
    if (b.length % 2 && passado) { oc = (oc ? oc + ' · ' : '') + 'Marcação ímpar'; tot.impares++; }
    else if (!b.length && prev && passado) { oc = 'Falta'; tot.faltas++; }
    else if (saldo > 0) oc = (oc ? oc + ' · ' : '') + 'Hora extra';
    else if (saldo < 0 && b.length) oc = (oc ? oc + ' · ' : '') + 'Atraso / saída antecipada';
    if (b.length) tot.diasTrab++;
    tot.trab += trab;
    if (saldo != null) tot.prev += prev;
    if (saldo > 0) tot.extra += saldo; else if (saldo < 0) tot.falta += -saldo;
    dias.push({ iso: iso, d: d, dow: dow, sem: NOMES[dow], batidas: b, trab: trab, prev: prev, saldo: saldo, oc: oc, folga: !prev && !b.length });
  }
  tot.saldo = tot.extra - tot.falta;
  const emp = EMPRESAS_AFD.filter(function (e) { return empresaChave_(e.nome) === empresaChave_(f.empresa); })[0] || EMPRESAS_AFD[0];
  return { mes: mes, dias: dias, tot: tot, horario: h, emp: emp, rhNome: config_().rhNome,
    func: { nome: f.nome, mat: String(f.matricula).replace(/^L2-/, ''), cpf: f.cpf ? f.cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4') : '', cargo: f.cargo, setor: f.setor },
    gerado: Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') };
}
function espelhoHtml_(D, modelo) {
  const MES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
  const comp = MES[Number(D.mes.slice(5, 7)) - 1] + ' de ' + D.mes.slice(0, 4);
  const hTxt = D.horario ? D.horario.descricao : 'Horário não cadastrado';
  const b = function (x, i) { return x.batidas[i] || ''; };
  const ext = function (x) { return x.batidas.slice(4).join(' '); };
  const sal = function (x) { return x.saldo == null ? '' : x.saldo === 0 ? '0:00' : hmTxt_(x.saldo, true); };
  const corSal = function (x) { return x.saldo > 0 ? '#15803d' : x.saldo < 0 ? '#b42318' : '#111'; };
  const T = D.tot;
  const rodape = '<p style="font-size:9px;color:#555;margin:10px 0 0">Espelho gerado pelo Rhitmo Atende em ' + D.gerado + ' a partir das marcações registradas no aplicativo. Tolerância de 10 minutos no dia (CLT art. 58 §1º). Às sextas a jornada tem 1 hora a menos.</p>' +
    '<table style="width:100%;margin-top:34px;border-collapse:collapse"><tr><td style="width:46%;border-top:1px solid #000;text-align:center;font-size:10px;padding-top:4px">' + esc_(D.func.nome) + '<br>Funcionário</td><td style="width:8%"></td><td style="width:46%;border-top:1px solid #000;text-align:center;font-size:10px;padding-top:4px">' + esc_(D.rhNome || 'RH') + '<br>Recursos Humanos</td></tr></table>';
  const cab = function (titulo) {
    return '<table style="width:100%;border-collapse:collapse;font-size:11px"><tr><td><b style="font-size:13px">' + esc_(D.emp.razao) + '</b><br>CNPJ ' + esc_(D.emp.cnpj.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5')) + (D.emp.endereco ? ' · ' + esc_(D.emp.endereco) : '') + '</td><td style="text-align:right"><b style="font-size:15px">' + titulo + '</b><br>Competência: ' + comp + '</td></tr></table>' +
      '<table style="width:100%;border-collapse:collapse;font-size:11px;margin-top:8px;border:1px solid #000"><tr><td style="padding:4px 6px"><b>Funcionário:</b> ' + esc_(D.func.nome) + '</td><td style="padding:4px 6px"><b>Matrícula:</b> ' + esc_(D.func.mat) + '</td><td style="padding:4px 6px"><b>CPF:</b> ' + esc_(D.func.cpf) + '</td></tr>' +
      '<tr><td style="padding:4px 6px"><b>Cargo:</b> ' + esc_(D.func.cargo) + '</td><td style="padding:4px 6px" colspan="2"><b>Horário:</b> ' + esc_(hTxt) + '</td></tr></table>';
  };
  const resumoLinha = '<table style="width:100%;border-collapse:collapse;font-size:11px;margin-top:8px;border:1px solid #000"><tr>' +
    [['Dias trabalhados', T.diasTrab], ['Horas trabalhadas', hmTxt_(T.trab)], ['Horas previstas', hmTxt_(T.prev)], ['Horas extras', hmTxt_(T.extra)], ['Atrasos / saídas', hmTxt_(T.falta)], ['Saldo do mês', hmTxt_(T.saldo, true)], ['Faltas', T.faltas]]
      .map(function (x) { return '<td style="border:1px solid #000;padding:4px;text-align:center">' + x[0] + '<br><b>' + x[1] + '</b></td>'; }).join('') + '</tr></table>';
  const corpoBase = 'font-family:Arial,Helvetica,sans-serif;color:#111;';
  let html = '';
  if (modelo === '2') { // Compacto
    html = cab('Espelho de ponto') + '<table style="width:100%;border-collapse:collapse;font-size:10.5px;margin-top:8px">' +
      '<tr style="background:#111;color:#fff"><th style="padding:4px;text-align:left">Dia</th><th style="padding:4px;text-align:left">Marcações</th><th style="padding:4px">Trab.</th><th style="padding:4px">Saldo</th><th style="padding:4px;text-align:left">Ocorrência</th></tr>' +
      D.dias.map(function (x) { return '<tr style="border-bottom:1px solid #ddd' + (x.folga ? ';color:#888' : '') + '"><td style="padding:3px 4px">' + ('0' + x.d).slice(-2) + ' ' + x.sem + '</td><td style="padding:3px 4px;font-family:Courier New,monospace">' + x.batidas.join(' · ') + '</td><td style="padding:3px 4px;text-align:center">' + (x.trab ? hmTxt_(x.trab) : '') + '</td><td style="padding:3px 4px;text-align:center;color:' + corSal(x) + '">' + sal(x) + '</td><td style="padding:3px 4px">' + esc_(x.oc) + '</td></tr>'; }).join('') +
      '</table>' + resumoLinha + rodape;
  } else if (modelo === '3') { // Moderno
    const caixa = function (r, v, cor) { return '<td style="background:#f5f5f2;padding:10px;text-align:center;width:25%"><div style="font-size:10px;color:#555">' + r + '</div><div style="font-size:18px;font-weight:bold;color:' + (cor || '#111') + '">' + v + '</div></td>'; };
    html = '<table style="width:100%;border-collapse:collapse"><tr><td style="background:#ffcc00;padding:14px 16px"><b style="font-size:20px">Espelho de ponto</b><br><span style="font-size:12px">' + comp + ' · ' + esc_(D.emp.razao) + '</span></td><td style="background:#111;color:#ffcc00;padding:14px 16px;text-align:right;font-size:12px"><b style="font-size:14px;color:#fff">' + esc_(D.func.nome) + '</b><br>Matrícula ' + esc_(D.func.mat) + ' · CPF ' + esc_(D.func.cpf) + '<br>' + esc_(hTxt) + '</td></tr></table>' +
      '<table style="width:100%;border-collapse:separate;border-spacing:6px;margin-top:6px"><tr>' + caixa('Horas trabalhadas', hmTxt_(T.trab)) + caixa('Horas previstas', hmTxt_(T.prev)) + caixa('Saldo do mês', hmTxt_(T.saldo, true), T.saldo > 0 ? '#15803d' : T.saldo < 0 ? '#b42318' : '#111') + caixa('Faltas', T.faltas, T.faltas ? '#b42318' : '#111') + '</tr></table>' +
      '<table style="width:100%;border-collapse:collapse;font-size:10.5px;margin-top:4px"><tr style="border-bottom:2px solid #ffcc00"><th style="padding:5px;text-align:left">Data</th><th style="padding:5px">Entrada</th><th style="padding:5px">Intervalo</th><th style="padding:5px">Volta</th><th style="padding:5px">Saída</th><th style="padding:5px">Trab.</th><th style="padding:5px">Saldo</th><th style="padding:5px;text-align:left">Ocorrência</th></tr>' +
      D.dias.map(function (x, i) { return '<tr style="background:' + (x.folga ? '#f0efe9' : i % 2 ? '#fafaf7' : '#fff') + '"><td style="padding:4px 5px">' + ('0' + x.d).slice(-2) + '/' + D.mes.slice(5, 7) + ' ' + x.sem + '</td>' + [0, 1, 2, 3].map(function (k) { return '<td style="padding:4px 5px;text-align:center;font-family:Courier New,monospace">' + b(x, k) + '</td>'; }).join('') + '<td style="padding:4px 5px;text-align:center">' + (x.trab ? hmTxt_(x.trab) : '') + '</td><td style="padding:4px 5px;text-align:center;font-weight:bold;color:' + corSal(x) + '">' + sal(x) + '</td><td style="padding:4px 5px">' + esc_(x.oc) + '</td></tr>'; }).join('') +
      '</table>' + rodape;
  } else if (modelo === '4') { // Calendário
    const prim = D.dias[0].dow;
    let cel = [];
    for (let i = 0; i < prim; i++) cel.push('<td style="border:1px solid #ccc;background:#f5f5f2"></td>');
    D.dias.forEach(function (x) {
      const fundo = x.oc.indexOf('Falta') >= 0 ? '#fde2df' : x.folga ? '#f0efe9' : x.saldo > 0 ? '#e8f5ec' : '#fff';
      cel.push('<td style="border:1px solid #ccc;vertical-align:top;padding:4px;height:62px;background:' + fundo + '"><b style="font-size:12px">' + x.d + '</b>' +
        (x.batidas.length ? '<div style="font-family:Courier New,monospace;font-size:9.5px">' + x.batidas[0] + ' – ' + x.batidas[x.batidas.length - 1] + '</div><div style="font-size:10px">' + hmTxt_(x.trab) + (x.saldo ? ' <span style="color:' + corSal(x) + '">(' + sal(x) + ')</span>' : '') + '</div>' : '') +
        (x.oc ? '<div style="font-size:9px;color:#555">' + esc_(x.oc.replace('Feriado: ', '')) + '</div>' : '') + '</td>');
    });
    while (cel.length % 7) cel.push('<td style="border:1px solid #ccc;background:#f5f5f2"></td>');
    let linhas = '';
    for (let i = 0; i < cel.length; i += 7) linhas += '<tr>' + cel.slice(i, i + 7).join('') + '</tr>';
    html = cab('Espelho de ponto — calendário') + '<table style="width:100%;border-collapse:collapse;margin-top:8px;table-layout:fixed"><tr>' + ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'].map(function (n) { return '<th style="background:#111;color:#ffcc00;padding:4px;font-size:11px">' + n + '</th>'; }).join('') + '</tr>' + linhas + '</table>' + resumoLinha + rodape;
  } else if (modelo === '5') { // Semanal
    let semana = [], blocos = '', acum = 0, n = 1;
    const fecha = function () {
      if (!semana.length) return;
      const t = semana.reduce(function (a, x) { a.trab += x.trab; a.prev += x.saldo != null ? x.prev : 0; a.saldo += x.saldo || 0; return a; }, { trab: 0, prev: 0, saldo: 0 });
      acum += t.saldo;
      blocos += '<tr><td colspan="8" style="background:#111;color:#ffcc00;padding:4px 6px;font-weight:bold">Semana ' + n++ + ' · ' + ('0' + semana[0].d).slice(-2) + ' a ' + ('0' + semana[semana.length - 1].d).slice(-2) + '/' + D.mes.slice(5, 7) + '</td></tr>' +
        semana.map(function (x) { return '<tr style="border-bottom:1px solid #e5e5e5' + (x.folga ? ';color:#888' : '') + '"><td style="padding:3px 5px">' + ('0' + x.d).slice(-2) + ' ' + x.sem + '</td>' + [0, 1, 2, 3].map(function (k) { return '<td style="padding:3px 5px;text-align:center;font-family:Courier New,monospace">' + b(x, k) + '</td>'; }).join('') + '<td style="padding:3px 5px;text-align:center">' + (x.trab ? hmTxt_(x.trab) : '') + '</td><td style="padding:3px 5px;text-align:center;color:' + corSal(x) + '">' + sal(x) + '</td><td style="padding:3px 5px">' + esc_(x.oc) + '</td></tr>'; }).join('') +
        '<tr style="background:#fff3c2;font-weight:bold"><td style="padding:4px 5px" colspan="5">Total da semana · previsto ' + hmTxt_(t.prev) + '</td><td style="padding:4px 5px;text-align:center">' + hmTxt_(t.trab) + '</td><td style="padding:4px 5px;text-align:center">' + hmTxt_(t.saldo, true) + '</td><td style="padding:4px 5px">Acumulado: ' + hmTxt_(acum, true) + '</td></tr>';
      semana = [];
    };
    D.dias.forEach(function (x) { semana.push(x); if (x.dow === 6) fecha(); });
    fecha();
    html = cab('Espelho de ponto semanal') + '<table style="width:100%;border-collapse:collapse;font-size:10.5px;margin-top:8px"><tr><th style="padding:4px;text-align:left">Dia</th><th style="padding:4px">Entrada</th><th style="padding:4px">Intervalo</th><th style="padding:4px">Volta</th><th style="padding:4px">Saída</th><th style="padding:4px">Trab.</th><th style="padding:4px">Saldo</th><th style="padding:4px;text-align:left">Ocorrência</th></tr>' + blocos + '</table>' + resumoLinha + rodape;
  } else { // 1 Clássico
    html = cab('Espelho de ponto') + '<table style="width:100%;border-collapse:collapse;font-size:10.5px;margin-top:8px">' +
      '<tr style="background:#e5e5e5">' + ['Data', 'Dia', 'Entrada', 'Saída', 'Entrada', 'Saída', 'Extras', 'Trab.', 'Previsto', 'Saldo', 'Ocorrência'].map(function (t) { return '<th style="border:1px solid #000;padding:3px">' + t + '</th>'; }).join('') + '</tr>' +
      D.dias.map(function (x) { return '<tr' + (x.folga ? ' style="background:#f2f2f2"' : '') + '><td style="border:1px solid #000;padding:2px 4px">' + ('0' + x.d).slice(-2) + '/' + D.mes.slice(5, 7) + '</td><td style="border:1px solid #000;padding:2px 4px">' + x.sem + '</td>' + [0, 1, 2, 3].map(function (k) { return '<td style="border:1px solid #000;padding:2px 4px;text-align:center">' + b(x, k) + '</td>'; }).join('') + '<td style="border:1px solid #000;padding:2px 4px;text-align:center">' + ext(x) + '</td><td style="border:1px solid #000;padding:2px 4px;text-align:center">' + (x.trab ? hmTxt_(x.trab) : '') + '</td><td style="border:1px solid #000;padding:2px 4px;text-align:center">' + (x.prev ? hmTxt_(x.prev) : '') + '</td><td style="border:1px solid #000;padding:2px 4px;text-align:center;color:' + corSal(x) + '">' + sal(x) + '</td><td style="border:1px solid #000;padding:2px 4px">' + esc_(x.oc) + '</td></tr>'; }).join('') +
      '</table>' + resumoLinha + rodape;
  }
  return '<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head><body style="' + corpoBase + 'margin:14px">' + html + '</body></html>';
}
function espelhoSaida_(f, mes, modelo, pdf) {
  if (!/^\d{4}-\d{2}$/.test(String(mes))) throw new Error('Mês inválido.');
  modelo = String(modelo || props_().getProperty('ESPELHO_MODELO') || '1');
  const html = espelhoHtml_(espelhoDados_(f, mes), modelo);
  const nome = 'Espelho de ponto ' + mes + ' ' + String(f.nome).split(' ')[0];
  if (!pdf) return { html: html, nome: nome };
  const blob = Utilities.newBlob(html, 'text/html', 'espelho.html').getAs('application/pdf');
  return { b64: Utilities.base64Encode(blob.getBytes()), nome: nomeArquivo_(nome) + '.pdf' };
}
function funcEspelho(token, mes, pdf) {
  const s = sessao_(token, 'func');
  return espelhoSaida_(s.f, mes, null, pdf);
}
function rhEspelho(token, matricula, mes, modelo, pdf) {
  sessao_(token, 'rh');
  const f = lerT_('Funcionarios').filter(function (x) { return x.matricula === matricula; })[0];
  if (!f) throw new Error('Funcionário não encontrado.');
  return espelhoSaida_(f, mes, modelo, pdf);
}
function rhModeloEspelho(token, modelo) {
  sessao_(token, 'rh');
  if (!MODELOS_ESPELHO.some(function (m) { return m[0] === String(modelo); })) throw new Error('Modelo inválido.');
  props_().setProperty('ESPELHO_MODELO', String(modelo));
  return String(modelo);
}

/* ----- AFD automático nos horários escolhidos pelo RH ----- */
const EMPRESAS_AFD = [
  { nome: 'MEPCO', cnpj: '08182147000151', razao: 'MEPCO DO BRASIL INDUSTRIAL DE ADITIVOS QUIMICOS LTDA', endereco: 'Avenida Calombé, 1373 · Duque de Caxias/RJ' },
  { nome: 'L2 Comercial', cnpj: '63009938000192', razao: 'L2 COMERCIAL LTDA' }
];
function rhAgendaAFD(token) {
  sessao_(token, 'rh');
  return agendaStatus_();
}
function agendaStatus_() {
  const p = props_();
  let ag = {}; try { ag = JSON.parse(p.getProperty('AFD_AGENDA') || '{}'); } catch (e) {}
  let log = []; try { log = JSON.parse(p.getProperty('AFD_LOG') || '[]'); } catch (e) {}
  let email = ''; try { email = Session.getEffectiveUser().getEmail(); } catch (e) {}
  return { ativa: !!ag.ativa, horarios: ag.horarios || ['18:00'], dias: ag.dias || [1, 2, 3, 4, 5], email: ag.email || email, log: log };
}
function rhSalvarAgendaAFD(token, ag) {
  sessao_(token, 'rh');
  const horarios = (ag.horarios || []).map(String).filter(function (x) { return /^\d{2}:\d{2}$/.test(x); }).slice(0, 6);
  const dias = (ag.dias || []).map(Number).filter(function (x) { return x >= 0 && x <= 6; });
  const email = String(ag.email || '').trim();
  if (ag.ativa && !horarios.length) throw new Error('Escolha pelo menos um horário.');
  if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('E-mail inválido.');
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'afdAgendado') ScriptApp.deleteTrigger(t); });
  if (ag.ativa) horarios.forEach(function (hm) {
    const h = Number(hm.slice(0, 2)), m = Number(hm.slice(3, 5));
    ScriptApp.newTrigger('afdAgendado').timeBased().everyDays(1).atHour(h).nearMinute(m).inTimezone(TZ).create();
  });
  props_().setProperty('AFD_AGENDA', JSON.stringify({ ativa: !!ag.ativa, horarios: horarios, dias: dias, email: email }));
  return agendaStatus_();
}
/** Chamado pelos acionadores de horário. Também dá para rodar no editor para testar. */
function afdAgendado() {
  const st = agendaStatus_();
  const t = agora_();
  const dow = new Date(Date.UTC(Number(t.data.slice(0, 4)), Number(t.data.slice(5, 7)) - 1, Number(t.data.slice(8, 10)), 12)).getUTCDay();
  if (st.dias.indexOf(dow) < 0) return;
  const r = gerarAFDsAgora_(st.email);
  const p = props_();
  const log = [{ em: t.iso, resumo: r.resumo }].concat(st.log).slice(0, 15);
  p.setProperty('AFD_LOG', JSON.stringify(log));
}
function rhGerarAFDsAgora(token) {
  sessao_(token, 'rh');
  const st = agendaStatus_();
  const r = gerarAFDsAgora_(st.email);
  const log = [{ em: agora_().iso, resumo: r.resumo + ' (manual)' }].concat(st.log).slice(0, 15);
  props_().setProperty('AFD_LOG', JSON.stringify(log));
  return agendaStatus_();
}
/** Gera o AFD de cada empresa (mês atual; nos 5 primeiros dias também o mês anterior), salva no Drive e manda por e-mail. */
function gerarAFDsAgora_(email) {
  const t = agora_();
  const y = Number(t.data.slice(0, 4)), m = Number(t.data.slice(5, 7)), d = Number(t.data.slice(8, 10));
  let ini = t.data.slice(0, 8) + '01';
  if (d <= 5) ini = (m === 1 ? (y - 1) + '-12' : y + '-' + ('0' + (m - 1)).slice(-2)) + '-01';
  const p = props_();
  let pasta;
  if (p.getProperty('AFD_PASTA_ID')) { try { pasta = DriveApp.getFolderById(p.getProperty('AFD_PASTA_ID')); } catch (e) {} }
  if (!pasta) { pasta = DriveApp.createFolder('Rhitmo Atende - AFD para o Dixi'); p.setProperty('AFD_PASTA_ID', pasta.getId()); }
  const anexos = [], partes = [];
  EMPRESAS_AFD.forEach(function (emp) {
    const r = gerarAFD_(emp, ini, t.data);
    partes.push(emp.nome + ': ' + r.marcacoes + ' marcações' + (r.semCpf.length ? ' (' + r.semCpf.length + ' sem CPF)' : ''));
    if (!r.marcacoes) return;
    const bytes = [];
    for (let i = 0; i < r.conteudo.length; i++) { const c = r.conteudo.charCodeAt(i); bytes.push(c < 256 ? (c > 127 ? c - 256 : c) : 63); }
    const carimbo = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd_HHmm');
    pasta.createFile(Utilities.newBlob(bytes, 'text/plain', r.nome.replace('.txt', '_' + carimbo + '.txt')));
    anexos.push(Utilities.newBlob(bytes, 'text/plain', r.nome));
  });
  const resumo = 'Período ' + ini.split('-').reverse().join('/') + ' a ' + t.data.split('-').reverse().join('/') + ' · ' + partes.join(' · ');
  if (email && anexos.length) {
    MailApp.sendEmail(email, 'Rhitmo Atende · AFD para o Dixi (' + t.data.split('-').reverse().join('/') + ' ' + t.hora.slice(0, 5) + ')',
      'Seguem os arquivos AFD gerados pelo Rhitmo Atende para importar no Dixi.\n\n' + resumo + '\n\nOs mesmos arquivos ficam na pasta "Rhitmo Atende - AFD para o Dixi" do Google Drive.', { attachments: anexos });
  }
  return { resumo: resumo, arquivos: anexos.length };
}
/** Rode uma vez no editor depois de atualizar o código: libera e-mail, acionadores e acesso à IA. */
function autorizar() {
  ScriptApp.getProjectTriggers();
  Logger.log('E-mails restantes hoje: ' + MailApp.getRemainingDailyQuota());
  UrlFetchApp.fetch('https://generativelanguage.googleapis.com/', { muteHttpExceptions: true });
  DriveApp.getRootFolder();
  Logger.log('Pronto! Permissões liberadas para ' + Session.getEffectiveUser().getEmail());
}

/* =====================================================================
 * BANCO DE HORAS — saldos vindos do Rhitmo (arquivo bancodehoras.json)
 * itens = [{matricula, saldo (min), prazoMeses, abertos:[{data,rest,venc}], lanc:[{data,tipo,min,motivo}]}]
 * ===================================================================== */
function rhSalvarBH(token, itens, comp) {
  sessao_(token, 'rh');
  if (!Array.isArray(itens)) throw new Error('Lista de banco de horas inválida.');
  return comLock_(function () {
    const cab = ABAS.BancoHoras;
    const sh = aba_('BancoHoras', cab);
    const lista = lerT_('BancoHoras');
    const porMat = {};
    lista.forEach(function (b) { porMat[b.matricula] = b; });
    const novos = [];
    const agora = agora_().iso;
    itens.forEach(function (it) {
      const mat = String(it.matricula || '').trim();
      if (!mat) return;
      const reg = {
        matricula: mat, saldo: String(Math.round(Number(it.saldo) || 0)), prazoMeses: String(Number(it.prazoMeses) || ''),
        abertos: JSON.stringify((it.abertos || []).slice(0, 50)), lanc: JSON.stringify((it.lanc || []).slice(0, 15)),
        comp: String(comp || ''), atualizadoEm: agora
      };
      if (porMat[mat]) Object.assign(porMat[mat], reg); else { porMat[mat] = reg; novos.push(reg); }
    });
    const todos = lista.concat(novos);
    if (todos.length) {
      const r = sh.getRange(2, 1, todos.length, cab.length);
      r.setNumberFormat('@');
      r.setValues(todos.map(function (b) { return cab.map(function (c) { return b[c] == null ? '' : String(b[c]); }); }));
    }
    sujo_('BancoHoras');
    return { atualizados: itens.length };
  });
}
function bhDe_(matricula) {
  const b = lerT_('BancoHoras').filter(function (x) { return x.matricula === matricula; })[0];
  if (!b) return null;
  let abertos = [], lanc = [];
  try { abertos = JSON.parse(b.abertos || '[]'); } catch (e) {}
  try { lanc = JSON.parse(b.lanc || '[]'); } catch (e) {}
  return { saldo: Number(b.saldo) || 0, prazoMeses: Number(b.prazoMeses) || null, abertos: abertos, lanc: lanc, comp: b.comp, atualizadoEm: b.atualizadoEm };
}

/** d = {matricula, tipo, titulo, competencia, ano, nomeArquivo, b64} */
function rhEnviarDoc(token, d) {
  sessao_(token, 'rh');
  const mat = String(d.matricula || '');
  const f = lerT_('Funcionarios').filter(function (x) { return x.matricula === mat; })[0];
  if (!f) throw new Error('Funcionário não encontrado: ' + mat);
  if (['recibo', 'informe', 'espelho', 'outro'].indexOf(d.tipo) < 0) throw new Error('Tipo de documento inválido.');
  const titulo = String(d.titulo || '').trim().slice(0, 120);
  if (!titulo) throw new Error('Documento sem título.');
  const bytes = Utilities.base64Decode(String(d.b64 || ''));
  if (!bytes.length) throw new Error('O arquivo está vazio.');
  if (bytes.length > MAX_PDF) throw new Error('O PDF tem mais de 10 MB.');
  if (String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]) !== '%PDF') throw new Error('O arquivo não é um PDF válido.');
  const pasta = DriveApp.getFolderById(props_().getProperty('PASTA_ID'));
  const arq = pasta.createFile(Utilities.newBlob(bytes, 'application/pdf', mat + ' - ' + nomeArquivo_(titulo) + '.pdf'));
  return comLock_(function () {
    const reg = {
      matricula: mat, tipo: d.tipo, titulo: titulo,
      competencia: String(d.competencia || ''), ano: String(Number(d.ano) || ''),
      fileId: arq.getId(), arquivo: String(d.nomeArquivo || '').slice(0, 150),
      tamanho: String(bytes.length), enviadoEm: agora_().iso
    };
    // mesmo funcionário, tipo e título (ex.: recibo de março enviado de novo): substitui em vez de duplicar
    const antigo = lerT_('Documentos').filter(function (x) { return x.matricula === mat && x.tipo === d.tipo && x.titulo === titulo; })[0];
    if (antigo) {
      try { DriveApp.getFileById(antigo.fileId).setTrashed(true); } catch (e) {}
      atualizarT_('Documentos', antigo._row, reg);
    } else {
      inserirT_('Documentos', Object.assign({ id: id_() }, reg));
    }
    return d.leve ? true : docsDe_(mat);
  });
}

/* ---------- correções de documentos (rodar no editor) ---------- */
/** Apaga do app (e manda para a lixeira do Drive) as linhas de Documentos que a função "manter" recusar. Regrava a aba de uma vez só. */
function filtrarDocumentos_(manter) {
  return comLock_(function () {
    const cab = ABAS.Documentos, sh = aba_('Documentos', cab);
    const todos = lerT_('Documentos');
    const fica = [], sai = [];
    todos.forEach(function (d, i) { (manter(d, i, todos) ? fica : sai).push(d); });
    if (!sai.length) return 0;
    sai.forEach(function (d) { try { DriveApp.getFileById(d.fileId).setTrashed(true); } catch (e) {} });
    if (todos.length) sh.getRange(2, 1, todos.length, cab.length).clearContent();
    if (fica.length) {
      const r = sh.getRange(2, 1, fica.length, cab.length);
      r.setNumberFormat('@');
      r.setValues(fica.map(function (d) { return cab.map(function (c) { return d[c] == null ? '' : String(d[c]); }); }));
    }
    sujo_('Documentos');
    return sai.length;
  });
}
/** Mesmo funcionário, tipo e título mais de uma vez: fica só o envio mais recente. */
function limparDocumentosRepetidos() {
  const novo = {};
  lerT_('Documentos').forEach(function (d) { const k = d.matricula + '|' + d.tipo + '|' + d.titulo; if (!novo[k] || d.enviadoEm > novo[k]) novo[k] = d.enviadoEm; });
  const n = filtrarDocumentos_(function (d) { return d.enviadoEm === novo[d.matricula + '|' + d.tipo + '|' + d.titulo]; });
  Logger.log(n + ' documento(s) repetido(s) apagado(s).');
  return n;
}
/**
 * CORREÇÃO ÚNICA (29/10/2026). Os recibos enviados em 21/10 eram de julho, agosto, setembro e outubro,
 * mas entraram todos com o nome "outubro de 2026". Os meses certos já foram enviados de novo depois
 * (janeiro a setembro em 29/10 e outubro em 22/10), então os de 21/10 são apagados.
 * Rode uma vez: selecione "corrigirRecibosDe21deOutubro" no editor e clique em Executar.
 */
function corrigirRecibosDe21deOutubro() {
  const n = filtrarDocumentos_(function (d) { return !(d.tipo === 'recibo' && d.enviadoEm < '2026-10-22' && /outubro de 2026$/.test(d.titulo)); });
  Logger.log(n + ' recibo(s) com o mês errado apagado(s).');
  limparDocumentosRepetidos();
}

/** Para separar o PDF de informes da contabilidade: diz de quem é cada CPF encontrado. */
function rhCpfsFuncionarios(token, cpfs) {
  sessao_(token, 'rh');
  const quer = {};
  (cpfs || []).slice(0, 5000).forEach(function (c) { const n = cpfNum_(c); if (n) quer[n] = 1; });
  return lerT_('Funcionarios').filter(function (f) { return f.cpf && quer[f.cpf]; })
    .map(function (f) { return { cpf: f.cpf, matricula: f.matricula, nome: f.nome, ativo: f.ativo === '1', empresa: f.empresa }; });
}

function rhExcluirDoc(token, docId) {
  sessao_(token, 'rh');
  return comLock_(function () {
    const d = lerT_('Documentos').filter(function (x) { return x.id === docId; })[0];
    if (!d) throw new Error('Documento não encontrado.');
    try { DriveApp.getFileById(d.fileId).setTrashed(true); } catch (e) { /* arquivo já removido */ }
    apagarLinha_('Documentos', d._row);
    return docsDe_(d.matricula);
  });
}

/* ----- RH: respostas automáticas ----- */
function rhRespostas(token) {
  sessao_(token, 'rh');
  if (props_().getProperty('RESP_SEMEADAS') !== '1') {
    comLock_(function () { semearRespostas_(); props_().setProperty('RESP_SEMEADAS', '1'); });
  }
  return rhRespostasSemSessao_();
}
function rhSalvarResposta(token, r) {
  sessao_(token, 'rh');
  const palavras = String(r && r.palavras || '').trim().slice(0, 300);
  const resposta = String(r && r.resposta || '').trim().slice(0, 1500);
  return comLock_(function () {
    const atual = r && r.id ? lerT_('Respostas').filter(function (x) { return x.id === r.id; })[0] : null;
    if (r && r.excluir) { if (atual) apagarLinha_('Respostas', atual._row); return rhRespostasSemSessao_(); }
    if (!palavras || resposta.length < 2) throw new Error('Preencha as palavras-chave e a resposta.');
    const dados = { palavras: palavras, resposta: resposta, ativo: r.ativo === false ? '0' : '1', avisar: r.avisar ? '1' : '0' };
    if (atual) atualizarT_('Respostas', atual._row, dados);
    else inserirT_('Respostas', Object.assign({ id: id_(), criadoEm: agora_().iso }, dados));
    return rhRespostasSemSessao_();
  });
}
function rhRespostasSemSessao_() {
  return lerT_('Respostas').map(function (r) { return { id: r.id, palavras: r.palavras, resposta: r.resposta, ativo: r.ativo === '1', avisar: r.avisar === '1' }; });
}
/** Volta a criar os modelos que foram apagados (não mexe nos que existem). */
function rhRespostasPadrao(token) {
  sessao_(token, 'rh');
  return comLock_(function () { semearRespostas_(); return rhRespostasSemSessao_(); });
}
/* Respostas prontas, com as regras da empresa tiradas do Rhitmo. [palavras, resposta, avisar o RH?] */
const RESPOSTAS_PADRAO = [
  ['adiantamento, vale do dia 20, adiantamento salarial, dia 20', 'O adiantamento é pago todo dia 20 e corresponde a 40% do salário base. Se o dia 20 cair em fim de semana ou feriado, ele sai no dia útil anterior.', 0],
  ['salario, pagamento, quando cai, quando recebo, quando vou receber, dia 5, cai o salario', 'O salário é pago no dia 5. Se cair em fim de semana ou feriado, sai no dia útil anterior. O recibo fica aqui no app, em Documentos.', 0],
  ['decimo terceiro, 13o, 13 salario, decimo, parcela do 13', 'O 13º é pago em 3 parcelas e o INSS é descontado na última. A primeira parcela sai até 30/11. Os recibos de cada parcela aparecem em Documentos.', 0],
  ['comprar tinta, compra de tinta, desconto funcionario, comprar produto, compra na fabrica, tinta', 'Você pode comprar produtos da fábrica com 15% de desconto sobre a tabela de preço. O valor é descontado na folha, no adiantamento do dia 20. Me manda aqui o que você precisa que eu encaminho o pedido.', 1],
  ['horario, expediente, que horas entro, que horas saio, escala, jornada', 'Seu horário de trabalho aparece no app, em Perfil. Às sextas todo mundo sai 1 hora mais cedo. Lembre de bater o ponto em todas as entradas e saídas, inclusive no intervalo.', 0],
  ['atraso, atrasado, tolerancia, cheguei tarde, cheguei atrasado', 'A tolerância é de até 5 minutos por marcação e 10 minutos no dia (CLT). Passando disso, a diferença aparece como atraso no espelho de ponto e pode ser descontada ou compensada no banco de horas.', 0],
  ['hora extra, horas extras, fazer extra, ficar ate mais tarde, trabalhar sabado', 'Hora extra só pode ser feita com autorização do seu encarregado. As horas vão para o banco de horas e precisam ser compensadas em até 6 meses. Seu saldo aparece em Ponto.', 0],
  ['esqueci de bater, esqueci o ponto, nao bati, bati errado, ponto errado, ajustar ponto, ajuste de ponto', 'Sem problema. Me diga aqui o dia e os horários certos (entrada, saída e volta do intervalo e saída) que eu confiro com seu encarregado e faço o ajuste.', 1],
  ['atestado, medico, consulta, doente, exame medico, estou mal', 'Melhoras! Entregue o atestado original no RH em até 48 horas. Se quiser adiantar, mande uma foto dele aqui.', 1],
  ['falta, faltar, vou faltar, nao vou poder ir, nao vou trabalhar, imprevisto', 'Obrigado por avisar. Avise também seu encarregado e me conte aqui o motivo. Falta sem justificativa é descontada do salário e do descanso semanal (DSR) da semana.', 1],
  ['casamento, vou casar, casei, licenca gala', 'Parabéns! Pela CLT você tem 3 dias seguidos de licença a partir do casamento. Me mande a certidão de casamento aqui ou entregue no RH.', 1],
  ['nascimento, filho nasceu, meu filho nasceu, bebe, licenca paternidade, licenca maternidade, gravida, gestante', 'Parabéns! A licença-paternidade é de 5 dias e a licença-maternidade é de 120 dias. Me mande a certidão de nascimento para eu registrar e incluir o dependente.', 1],
  ['falecimento, faleceu, luto, morreu, velorio, enterro', 'Meus sentimentos. Pela CLT você tem 2 dias de licença no falecimento de pai, mãe, cônjuge, filho, irmão ou dependente. Me avise aqui e depois entregue a certidão de óbito.', 1],
  ['vale transporte, passagem, vt, onibus, riocard, cartao de transporte', 'O vale-transporte tem desconto de 6% do salário base. Para incluir, mudar ou cancelar, me mande aqui seu trajeto e as linhas que você usa.', 1],
  ['uniforme, blusa, calca, bota, jaleco, epi, luva, oculos', 'Blusa, calça e polo são trocadas a cada 6 meses; bota, jaleco e blusa social a cada 12 meses. Se algo estragou antes, me diga aqui o item e o tamanho.', 1],
  ['declaracao, comprovante de vinculo, declaracao de trabalho, carta de trabalho, comprovante de renda', 'Faço sim. Me diga para que é (banco, escola, aluguel ou outro) que eu preparo e te aviso aqui quando estiver pronta.', 1],
  ['mudei de endereco, endereco novo, telefone novo, celular novo, trocar conta, conta bancaria, pix, mudar conta', 'Me mande aqui o dado novo que eu atualizo seu cadastro. Para conta bancária, envie também um comprovante no seu nome.', 1],
  ['fgts, saque, extrato do fgts, saque aniversario', 'O FGTS é depositado todo mês: 8% do seu salário. Você vê o saldo e o extrato no aplicativo FGTS, da Caixa, entrando com seu CPF.', 0],
  ['carteira de trabalho, ctps, carteira digital', 'Sua carteira de trabalho é digital. Baixe o app Carteira de Trabalho Digital e entre com sua conta gov.br para ver seus registros.', 0],
  ['salario familia, salario-familia', 'O salário-família é de R$ 67,54 por filho de até 14 anos para quem ganha até R$ 1.980,38 (valores de 2026). Para receber, me mande a certidão de nascimento e a carteira de vacinação do filho.', 1],
  ['emprestimo, consignado, credito do trabalhador, adiantamento extra', 'O empréstimo consignado (Crédito do Trabalhador) é contratado no app Carteira de Trabalho Digital ou no seu banco, e as parcelas são descontadas na folha. Se tiver dúvida sobre algum desconto, me avise aqui.', 1],
  ['aumento, reajuste, dissidio, data base, data-base', 'O reajuste da categoria acontece na data-base, em maio. O último foi de 5,5%. Se quiser conversar sobre o seu salário, me conte aqui que eu levo para a direção.', 1],
  ['aso, exame periodico, exame admissional, exame demissional', 'O exame periódico (ASO) é feito uma vez por ano. Quando estiver perto do seu, eu aviso a data e o local aqui.', 0],
  ['rescisao, pedir demissao, pedir as contas, sair da empresa, aviso previo, acordo', 'Entendi. Esse assunto eu prefiro tratar com você pessoalmente. Me diga um horário bom que eu te chamo no RH.', 1],
  ['sexta, sexta-feira, sair mais cedo', 'Às sextas todo mundo sai 1 hora mais cedo. Bata o ponto normalmente na saída.', 0]
];
function semearRespostas_() {
  const ja = {};
  lerT_('Respostas').forEach(function (r) { ja[normTxt_(r.palavras)] = 1; });
  RESPOSTAS_PADRAO.forEach(function (x) {
    if (ja[normTxt_(x[0])]) return;
    inserirT_('Respostas', { id: id_(), palavras: x[0], resposta: x[1], ativo: '1', criadoEm: agora_().iso, avisar: x[2] ? '1' : '0' });
  });
}

/* ----- RH: pedidos de férias ----- */
function rhFerias(token) {
  sessao_(token, 'rh');
  const nomes = {};
  lerT_('Funcionarios').forEach(function (f) { nomes[f.matricula] = f.nome; });
  return lerT_('Ferias').map(function (f) {
    return { id: f.id, matricula: f.matricula, nome: nomes[f.matricula] || f.matricula, inicio: f.inicio, dias: Number(f.dias), abono: f.abono === '1', obs: f.obs, status: f.status, criadoEm: f.criadoEm, respostaRH: f.respostaRH };
  }).sort(function (a, b) { return (a.status === 'pendente' ? 0 : 1) - (b.status === 'pendente' ? 0 : 1) || (a.criadoEm < b.criadoEm ? 1 : -1); });
}
function rhResponderFerias(token, id, aprovar, msg) {
  sessao_(token, 'rh');
  msg = String(msg || '').trim().slice(0, 500);
  comLock_(function () {
    const f = lerT_('Ferias').filter(function (x) { return x.id === id; })[0];
    if (!f) throw new Error('Pedido não encontrado.');
    atualizarT_('Ferias', f._row, { status: aprovar ? 'aprovada' : 'recusada', respostaRH: msg, respondidoEm: agora_().iso });
    const p = f.inicio.split('-');
    avisoNaConversa_(f.matricula, (aprovar ? '✅ Suas férias foram aprovadas: ' : '❌ Seu pedido de férias não foi aprovado: ') + f.dias + ' dias a partir de ' + p[2] + '/' + p[1] + '/' + p[0] + '.' + (msg ? ' ' + msg : ''), 'rh', false);
  });
  return rhFerias(token);
}

function rhConfig(token, empresa, atendimento) {
  sessao_(token, 'rh');
  props_().setProperties({
    EMPRESA: String(empresa || 'MEPCO').trim().slice(0, 60) || 'MEPCO',
    ATENDIMENTO: String(atendimento || '').trim().slice(0, 120) || 'Segunda a sexta, das 8h às 17h'
  });
  return config_();
}

/* =====================================================================
 * APOIO (funções internas)
 * ===================================================================== */
let SS_ = null;
/* ---------- desempenho ----------
 * Ler a planilha é o que deixa o app lento (cada consulta ao Google leva um tempo).
 * Por isso cada aba lida fica guardada no cache do Apps Script por alguns minutos e
 * cada gravação apaga só o cache daquela aba. As configurações são lidas de uma vez só.
 * Se você editar a planilha à mão, o app pode levar até 10 minutos para ver a mudança
 * (ou rode "limparCache" no editor). */
const CACHE_SEG = 600;
const SEM_CACHE = { Fotos: 1 };      // fotos são grandes e só são lidas ao abrir o app
let PROPS_ = null, MEMO_ = {}, ABAS_ = {};
function props_() {
  const real = PropertiesService.getScriptProperties();
  const tudo = function () { if (!PROPS_) PROPS_ = real.getProperties(); return PROPS_; };
  return {
    getProperty: function (k) { const t = tudo(); return Object.prototype.hasOwnProperty.call(t, k) ? t[k] : null; },
    setProperty: function (k, v) { real.setProperty(k, String(v)); if (PROPS_) PROPS_[k] = String(v); return this; },
    setProperties: function (o) { real.setProperties(o); if (PROPS_) Object.keys(o).forEach(function (k) { PROPS_[k] = String(o[k]); }); return this; },
    deleteProperty: function (k) { real.deleteProperty(k); if (PROPS_) delete PROPS_[k]; return this; },
    getProperties: function () { return Object.assign({}, tudo()); },
    fresco: function (k) { return real.getProperty(k); }   // leitura sem memória (contadores)
  };
}
function ss_() {
  if (!SS_) {
    const id = props_().getProperty('SS_ID');
    if (!id) throw new Error('O app ainda não foi configurado. No editor do Apps Script, rode a função "configurar".');
    SS_ = SpreadsheetApp.openById(id);
  }
  return SS_;
}
function config_() {
  const p = props_();
  return {
    empresa: p.getProperty('EMPRESA') || 'MEPCO', atendimento: p.getProperty('ATENDIMENTO') || '',
    rhNome: p.getProperty('RH_NOME') || 'Fabiano Lourenço', iaAtiva: p.getProperty('IA_ATIVA') === '1',
    avisoAuto: p.getProperty('AVISO_AUTO') !== '0', modeloEspelho: p.getProperty('ESPELHO_MODELO') || '1'
  };
}
function aba_(nome, cab) {
  if (ABAS_[nome]) return ABAS_[nome];
  const ss = ss_();
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    escreverLinha_(sh, 1, cab);
    sh.setFrozenRows(1);
    sujo_(nome);
  } else if (sh.getLastColumn() < cab.length) {
    escreverLinha_(sh, 1, cab);   // versão nova do app acrescentou colunas
  }
  return ABAS_[nome] = sh;
}
function escreverLinha_(sh, row, vals) {
  const r = sh.getRange(row, 1, 1, vals.length);
  r.setNumberFormat('@');   // texto puro: preserva zeros à esquerda da matrícula
  r.setValues([vals.map(function (v) { return v == null ? '' : String(v); })]);
}
function chaveCache_(nome) { return 't_' + nome.replace(/[^A-Za-z0-9]/g, '_'); }
/** Marca que a aba mudou: apaga o cache dela (troca a "versão", então nada antigo é aproveitado). */
function sujo_(nome) {
  delete MEMO_[nome];
  if (SEM_CACHE[nome]) return;
  try { CacheService.getScriptCache().put(chaveCache_(nome) + '_v', id_(), 21600); } catch (e) {}
}
/** Linhas da aba (sem o cabeçalho), da memória, do cache ou da planilha. */
function linhas_(nome, cab) {
  if (MEMO_[nome]) return MEMO_[nome];
  const usaCache = !SEM_CACHE[nome];
  const c = CacheService.getScriptCache(), k = chaveCache_(nome);
  let ver = '';
  if (usaCache) {
    try {
      const g = c.getAll([k + '_v', k + '_n', k + '_0', k + '_1', k + '_2']);   // quase sempre cabe tudo numa consulta só
      ver = g[k + '_v'] || '';
      const n = Number(g[k + '_n'] || 0);
      if (ver && n) {
        const ks = []; for (let i = 0; i < n; i++) ks.push(k + '_' + i);
        const partes = n <= 3 ? g : c.getAll(ks);
        let txt = '', ok = true;
        for (let i = 0; i < n && ok; i++) { if (partes[ks[i]] == null) ok = false; else txt += partes[ks[i]]; }
        if (ok) { const o = JSON.parse(txt); if (o.v === ver) return MEMO_[nome] = o.r; }
      }
      if (!ver) { ver = id_(); c.put(k + '_v', ver, 21600); }
    } catch (e) { ver = ''; }
  }
  let rows = [];
  const ss = ss_();
  let sh = ABAS_[nome] || ss.getSheetByName(nome);
  if (!sh && nome.indexOf('Ponto ') !== 0) sh = aba_(nome, cab);
  if (sh) {
    ABAS_[nome] = sh;
    const v = sh.getDataRange().getDisplayValues();
    if (v.length && v[0].filter(String).length < cab.length) escreverLinha_(sh, 1, cab);   // versão nova do app acrescentou colunas
    rows = v.slice(1).map(function (r) { const o = []; for (let j = 0; j < cab.length; j++) o.push(r[j] == null ? '' : String(r[j])); return o; });
  }
  if (usaCache && ver) {
    try {
      const txt = JSON.stringify({ v: ver, r: rows }), TAM = 40000, put = {};
      let n = 0;
      for (let i = 0; i < txt.length; i += TAM) put[k + '_' + (n++)] = txt.slice(i, i + TAM);
      if (n <= 200) { put[k + '_n'] = String(n); c.putAll(put, CACHE_SEG); }
    } catch (e) {}
  }
  return MEMO_[nome] = rows;
}
function ler_(nome, cab) {
  const out = [];
  linhas_(nome, cab).forEach(function (r, i) {
    if (!r.some(function (v) { return v !== ''; })) return;   // linha em branco (ex.: depois de uma limpeza)
    const o = { _row: i + 2 };
    cab.forEach(function (c, j) { o[c] = r[j]; });
    out.push(o);
  });
  return out;
}
function lerT_(nome) { return ler_(nome, ABAS[nome]); }
function inserirT_(nome, obj) {
  const cab = ABAS[nome];
  const sh = aba_(nome, cab);
  escreverLinha_(sh, sh.getLastRow() + 1, cab.map(function (c) { return obj[c]; }));
  sujo_(nome);
}
function atualizarT_(nome, row, patch) {
  const cab = ABAS[nome];
  const sh = aba_(nome, cab);
  const cols = Object.keys(patch).map(function (k) { return cab.indexOf(k); }).filter(function (j) { return j >= 0; }).sort(function (a, b) { return a - b; });
  if (!cols.length) return;
  const ini = cols[0], fim = cols[cols.length - 1];
  const atual = fim - ini + 1 > cols.length ? (linhas_(nome, cab)[row - 2] || []).slice() : [];
  const vals = [];
  for (let j = ini; j <= fim; j++) { const k = cab[j]; vals.push(Object.prototype.hasOwnProperty.call(patch, k) ? String(patch[k]) : (atual[j] == null ? '' : atual[j])); }
  const r = sh.getRange(row, ini + 1, 1, vals.length);   // grava as colunas de uma vez só
  r.setNumberFormat('@');
  r.setValues([vals]);
  const memo = MEMO_[nome];
  sujo_(nome);
  if (memo && memo[row - 2]) { for (let j = ini; j <= fim; j++) memo[row - 2][j] = vals[j - ini]; MEMO_[nome] = memo; }   // a própria execução continua vendo o valor novo
}
function apagarLinha_(nome, row) {
  aba_(nome, ABAS[nome] || CAB_PONTO).deleteRow(row);
  sujo_(nome);
}
/** Rode no editor se editou a planilha à mão e quer que o app veja na hora. */
function limparCache() {
  Object.keys(ABAS).forEach(sujo_);
  const ss = ss_();
  ss.getSheets().forEach(function (sh) { if (sh.getName().indexOf('Ponto ') === 0) sujo_(sh.getName()); });
  Logger.log('Cache limpo.');
}
function pontoAba_(mes) { return aba_('Ponto ' + mes, CAB_PONTO); }
function lerPonto_(mes) { return ler_('Ponto ' + mes, CAB_PONTO); }
function pontoDe_(matricula, mes) {
  return lerPonto_(mes)
    .filter(function (p) { return p.matricula === matricula; })
    .map(function (p) { return { data: p.data, hora: p.hora, em: p.em, origem: p.origem || 'app' }; })
    .sort(function (a, b) { return a.em < b.em ? -1 : 1; });
}
function chamadosDe_(matricula) {
  const ch = lerT_('Chamados').filter(function (c) { return c.matricula === matricula; });
  const msgs = {};
  ch.forEach(function (c) { msgs[c.id] = []; });
  lerT_('Mensagens').forEach(function (m) {
    if (msgs[m.chamadoId]) msgs[m.chamadoId].push({ de: m.de, texto: m.texto, em: m.em });
  });
  return ch.map(function (c) {
    return {
      id: c.id, assunto: c.assunto, status: c.status, criadoEm: c.criadoEm, atualizadoEm: c.atualizadoEm,
      naoLidoRH: c.naoLidoRH === '1', naoLidoFunc: c.naoLidoFunc === '1',
      mensagens: msgs[c.id].sort(function (a, b) { return a.em < b.em ? -1 : 1; })
    };
  }).sort(function (a, b) { return a.atualizadoEm < b.atualizadoEm ? 1 : -1; });
}
function docsDe_(matricula) {
  const ordem = function (d) { return /^Demonstrativo/i.test(d.titulo) ? 1 : 0; };   // no mesmo mês: recibo primeiro, depois o demonstrativo
  return lerT_('Documentos').filter(function (d) { return d.matricula === matricula; }).map(function (d) {
    return { id: d.id, tipo: d.tipo, titulo: d.titulo, competencia: d.competencia, ano: Number(d.ano), tamanho: Number(d.tamanho), enviadoEm: d.enviadoEm };
  }).sort(function (a, b) {
    return (b.competencia || '').localeCompare(a.competencia || '') || ordem(a) - ordem(b) || (a.enviadoEm < b.enviadoEm ? 1 : -1);
  });
}
function chamado_(id, s) {
  const c = lerT_('Chamados').filter(function (x) { return x.id === id; })[0];
  if (!c || (s.papel === 'func' && c.matricula !== s.matricula)) throw new Error('Pedido não encontrado.');
  return c;
}
function limparTexto_(t) {
  t = String(t || '').trim();
  if (t.length < 2) throw new Error('Escreva a mensagem antes de enviar.');
  if (t.length > MAX_TEXTO) throw new Error('A mensagem passou de ' + MAX_TEXTO + ' caracteres. Divida em duas.');
  return t;
}
function agora_() {
  const d = new Date();
  return {
    iso: d.toISOString(), ms: d.getTime(),
    data: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'),
    mes: Utilities.formatDate(d, TZ, 'yyyy-MM'),
    hora: Utilities.formatDate(d, TZ, 'HH:mm:ss')
  };
}
function id_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); }
function comLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); }
  finally { SpreadsheetApp.flush(); lock.releaseLock(); }
}
function nomeArquivo_(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w.\- ]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 100) || 'documento';
}

/* ---------- segurança ---------- */
function gerarSalt_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); }
function gerarSenha_() {
  const al = 'abcdefghjkmnpqrstuvwxyz23456789';
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Date.now());
  let s = '';
  for (let i = 0; i < 8; i++) s += al.charAt((b[i] + 256) % al.length);
  return s;
}
function hash_(senha, salt) {
  let h = salt + '|' + senha;
  for (let i = 0; i < 300; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  }
  return h;
}
function criarSessao_(s) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('s_' + token, JSON.stringify(s), SESSAO_SEG);
  return token;
}
/** Valida a sessão. Senha trocada, senha redefinida ou funcionário desativado derrubam a sessão. */
function sessao_(token, papel) {
  const cache = CacheService.getScriptCache();
  const raw = token ? cache.get('s_' + token) : null;
  if (!raw) throw new Error('SESSAO_EXPIRADA');
  const s = JSON.parse(raw);
  if (papel && s.papel !== papel) throw new Error('Acesso não permitido.');
  if (s.papel === 'rh') {
    if (s.salt !== props_().getProperty('RH_SALT')) { cache.remove('s_' + token); throw new Error('SESSAO_EXPIRADA'); }
  } else {
    const f = lerT_('Funcionarios').filter(function (x) { return x.matricula === s.matricula; })[0];
    if (!f || f.ativo !== '1' || f.salt !== s.salt) { cache.remove('s_' + token); throw new Error('SESSAO_EXPIRADA'); }
    s.f = f;
  }
  cache.put('s_' + token, raw, SESSAO_SEG);
  return s;
}
