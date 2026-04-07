// ====================================================================================
// 🧠 ARQUIVO: Codigo.gs (SGCW v1.0 - Sistema de Gestão de Comunicação Web)
// ====================================================================================

const PLANILHA_ID = "1wcgYDTH7C9vRuu2CE43WMB1Rh0h9xf3JplTJWqQtsZA";
const ID_PLANILHA_TECNICOS = "1yrYwyE0iy4aYKHEthMxPF3rOUiygfQkSLqXBzIGtK4I";
const EMAIL_REMETENTE = "rastreamento@zenseguro.com";

const MAPA_COLUNAS = {
  DATA: 1, NOME: 2, PLACA: 3, CHASSI: 4, FIPE: 5, EMAIL: 6, TELEFONE: 7,
  CHECK_EMAIL: 8, DATA_EMAIL: 9, RESPONDEU_EMAIL: 10, CHECK_WHATS: 11, 
  DATA_WHATS: 12, RESPONDEU_WHATS: 13, RESPONSAVEL: 14, FIPE_BAIXA: 15, 
  TECNICO_INDISPONIVEL: 16, ESTADO: 17
};

const SGA_CONFIG = {
  URL_AUTH: "https://api.hinova.com.br/api/sga/v2/usuario/autenticar",
  URL_CONSULTA_BASE: "https://api.hinova.com.br/api/sga/v2/veiculo/buscar/",
  TOKEN_ASSOCIACAO: "041e6d561c08d16fce2a5beead2ca02fa4f4ee113d51a6f56e9c9e7c89694e4187b271c6cf997f4ece7bf4180e13ee750cb67ffbaa15d5bc82260c3b57a27cfc114322ab6e29a6dc368f1c2b4ea59702456d6c9df528c97f3386aee5978276f6",
  USUARIO: "victor rodrigues", SENHA: "ZEN0102"
};

const MAPA_SITUACAO_SGA = { "1": "Ativo", "2": "Inativo", "3": "Pendente", "4": "Inadimplente", "5": "Negado", "6": "Cancelado", "7": "Evento", "8": "Indenizado", "11": "Cancelado com rastreador", "12": "Inativos com rastreador", "13": "Inativos sem rastreador", "14": "Ativo com adesivo", "17": "Cancelamento pendente", "18": "Envio de termos", "19": "Desligado do corpo associativo", "22": "Aguardando indenização" };

function Z_AUTORIZAR_SCRIPT() {
  const usuario = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({ to: usuario, subject: "SGCW - Autorização de BI e Slides", body: "Permissões de E-mail concedidas!" });
  try { SlidesApp.create("SGCW_Auth").setTrashed(true); } catch (e) { }
  console.log("✅ Permissões de envio de E-mail e Google Slides concedidas com sucesso.");
}

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.urlApp = ScriptApp.getService().getUrl();
  t.viewParam = e.parameter.view || '';
  return t.evaluate().setTitle('SGCW - Portal Operacional').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function solicitarDadosWeb(tela, parametro) {
  try {
    if (tela === 'dashboard') return web_obterDadosDashboard();
    if (tela === 'filaGeral') return web_obterFilaGeral();
    if (tela === 'logs') return web_obterDadosLogs(parametro);
    if (tela === 'config') return web_obterConfiguracoes();
    if (tela === 'tecnicos') return web_obterTecnicos();
    if (tela === 'statusAPI') return web_testarAPIs(); 
    return null;
  } catch (erro) { return { erro: "Erro Backend: " + erro.message }; }
}

function web_formatarDataSegura(v) { return (v && v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(v || "").trim(); }
function web_converterBooleano(v) { return (v === true || v === "TRUE" || v === 1) ? "✔️" : (v === false || v === "FALSE" || v === 0) ? "❌" : String(v || ""); }
function web_extrairDataParaCalculo(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return new Date(valor.getTime());
  var str = String(valor).trim().split(" ")[0];
  var p = str.split("/");
  if (p.length === 3) return new Date(p[2], p[1] - 1, p[0]);
  return null;
}

function getTemplatesDict(ss) {
  const aba = ss.getSheetByName("⚙️ Configurações");
  const dict = {};
  if (aba) {
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0]) dict[String(dados[i][0]).trim()] = String(dados[i][1] || "");
    }
  }
  return dict;
}

// ====================================================================================
// NOVO MOTOR DE DASHBOARD (BI COMPLETO)
// ====================================================================================
function web_obterDadosDashboard() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();

  const stats = {
    kpis: {
      pendentesTotal: 0, pendentesEtapa: { '1': 0, '2': 0, '3': 0 },
      conclusaoTotal: 0, conclusaoEtapa: { '1': 0, '2': 0, '3': 0 },
      emails: { hoje: 0, semana: 0, mes: 0, total: 0, etapa: { '1': 0, '2': 0, '3': 0 } },
      whats: { hoje: 0, semana: 0, mes: 0, total: 0, etapa: { '1': 0, '2': 0, '3': 0 } }
    },
    graficos: { historicoMensal: {}, equipeMes: {} }
  };

  ["1 -", "2 -", "3 -"].forEach(nomeFrag => {
    const aba = ss.getSheets().find(s => s.getName().includes(nomeFrag));
    if (aba) {
      const etapaStr = nomeFrag.includes("1") ? '1' : nomeFrag.includes("2") ? '2' : '3';
      const d = aba.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) {
        if (!d[i][MAPA_COLUNAS.PLACA] && !d[i][MAPA_COLUNAS.CHASSI]) continue;
        const eE = d[i][MAPA_COLUNAS.CHECK_EMAIL] === true || d[i][MAPA_COLUNAS.CHECK_EMAIL] === "TRUE" || d[i][MAPA_COLUNAS.CHECK_EMAIL] === 1;
        const eW = d[i][MAPA_COLUNAS.CHECK_WHATS] === true || d[i][MAPA_COLUNAS.CHECK_WHATS] === "TRUE" || d[i][MAPA_COLUNAS.CHECK_WHATS] === 1;
        if (!eE && !eW) { stats.kpis.pendentesTotal++; stats.kpis.pendentesEtapa[etapaStr]++; }
      }
    }
  });

  const aud = ss.getSheetByName("4 -Registro - NÃO ALTERAR");
  if (aud) {
    const dAud = aud.getDataRange().getValues();
    if (dAud.length > 0) {
      const cab = dAud[0].map(c => String(c).trim());
      const mapasEnvio = [
        { canal: 'emails', etapa: '1', idxDat: cab.findIndex(c => c.includes("1- Enviado e-mail")), idxResp: cab.findIndex(c => c.includes("1-Responsável")) },
        { canal: 'whats', etapa: '1', idxDat: cab.findIndex(c => c.includes("1 -Enviado whats")), idxResp: -1 },
        { canal: 'emails', etapa: '2', idxDat: cab.findIndex(c => c.includes("2- Enviado e-mail")), idxResp: cab.findIndex(c => c.includes("2-Responsável")) },
        { canal: 'whats', etapa: '2', idxDat: cab.findIndex(c => c.includes("2 -Enviado whats")), idxResp: -1 },
        { canal: 'emails', etapa: '3', idxDat: cab.findIndex(c => c.includes("3- Enviado e-mail")), idxResp: cab.findIndex(c => c.includes("3-Responsável")) },
        { canal: 'whats', etapa: '3', idxDat: cab.findIndex(c => c.includes("3 -Enviado whats")), idxResp: -1 }
      ];

      for (let i = 1; i < dAud.length; i++) {
        mapasEnvio.forEach(m => {
          if (m.idxDat > -1 && dAud[i][m.idxDat]) {
            const dataCalc = web_extrairDataParaCalculo(dAud[i][m.idxDat]);
            if (dataCalc) {
              stats.kpis[m.canal].total++;
              stats.kpis[m.canal].etapa[m.etapa]++;
              const diffDias = Math.floor((hoje.getTime() - dataCalc.getTime()) / 86400000);
              if (diffDias === 0) stats.kpis[m.canal].hoje++;
              if (diffDias >= 0 && diffDias <= 7) stats.kpis[m.canal].semana++;

              if (dataCalc.getMonth() === mesAtual && dataCalc.getFullYear() === anoAtual) {
                stats.kpis[m.canal].mes++;
                if (m.canal === 'emails' && m.idxResp > -1) {
                  const resp = String(dAud[i][m.idxResp]).trim();
                  if (resp && resp !== "Sistema" && !resp.includes("Gatilho")) {
                    stats.graficos.equipeMes[resp] = (stats.graficos.equipeMes[resp] || 0) + 1;
                  }
                }
              }

              const chaveMes = ("0" + (dataCalc.getMonth() + 1)).slice(-2) + "/" + dataCalc.getFullYear();
              if (!stats.graficos.historicoMensal[chaveMes]) {
                stats.graficos.historicoMensal[chaveMes] = { email: 0, whats: 0, sortKey: (dataCalc.getFullYear() * 100) + dataCalc.getMonth() };
              }
              if (m.canal === 'emails') stats.graficos.historicoMensal[chaveMes].email++;
              else stats.graficos.historicoMensal[chaveMes].whats++;
            }
          }
        });
      }
    }
  }

  const conc = ss.getSheetByName("Log Concluídos");
  if (conc) {
    const dConc = conc.getDataRange().getValues();
    for (let i = 1; i < dConc.length; i++) {
      stats.kpis.conclusaoTotal++;
      const abaOrigem = String(dConc[i][7] || "");
      if (abaOrigem.includes("1 -")) stats.kpis.conclusaoEtapa['1']++;
      else if (abaOrigem.includes("2 -")) stats.kpis.conclusaoEtapa['2']++;
      else if (abaOrigem.includes("3 -")) stats.kpis.conclusaoEtapa['3']++;
    }
  }
  return JSON.parse(JSON.stringify(stats));
}

// ====================================================================================
// EXPORTADOR PARA GOOGLE SLIDES
// ====================================================================================
function exportarDashboardParaSlidesWeb(graficos, statsObj) {
  try {
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    const presentation = SlidesApp.create("Relatório Executivo SGCW - " + timestamp);
    const slideTitulo = presentation.getSlides()[0];

    slideTitulo.insertTextBox("Relatório Executivo - BI\nSGCW OPERACIONAL", 50, 150, 600, 100).getText().getTextStyle().setFontSize(32).setBold(true).setForegroundColor("#4f46e5");
    slideTitulo.insertTextBox("Gerado automaticamente em: " + timestamp, 50, 260, 600, 40).getText().getTextStyle().setFontSize(14).setForegroundColor("#64748b");

    if (statsObj && statsObj.kpis) {
      const slideKpi = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      slideKpi.insertTextBox("Resumo Operacional", 50, 30, 620, 50).getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#1e293b");

      let textoKpi = "📊 NÚMEROS GERAIS DA OPERAÇÃO:\n\n";
      textoKpi += "• Veículos Pendentes de Ação: " + statsObj.kpis.pendentesTotal + "\n";
      textoKpi += "   (Etapa 1: " + statsObj.kpis.pendentesEtapa['1'] + " | Etapa 2: " + statsObj.kpis.pendentesEtapa['2'] + " | Etapa 3: " + statsObj.kpis.pendentesEtapa['3'] + ")\n\n";
      textoKpi += "• Total de E-mails Enviados: " + statsObj.kpis.emails.total + "\n";
      textoKpi += "   (Hoje: " + statsObj.kpis.emails.hoje + " | Semana: " + statsObj.kpis.emails.semana + " | Mês Atual: " + statsObj.kpis.emails.mes + ")\n\n";
      textoKpi += "• Total de WhatsApps Marcados: " + statsObj.kpis.whats.total + "\n";
      textoKpi += "   (Hoje: " + statsObj.kpis.whats.hoje + " | Semana: " + statsObj.kpis.whats.semana + " | Mês Atual: " + statsObj.kpis.whats.mes + ")\n\n";
      textoKpi += "• Conclusões (Instalações Finalizadas): " + statsObj.kpis.conclusaoTotal + "\n";
      textoKpi += "   (Etapa 1: " + statsObj.kpis.conclusaoEtapa['1'] + " | Etapa 2: " + statsObj.kpis.conclusaoEtapa['2'] + " | Etapa 3: " + statsObj.kpis.conclusaoEtapa['3'] + ")\n";

      slideKpi.insertTextBox(textoKpi, 50, 100, 620, 350).getText().getTextStyle().setFontSize(14).setForegroundColor("#334155");
    }

    graficos.forEach(graf => {
      const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      slide.insertTextBox(graf.titulo, 50, 30, 620, 50).getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#1e293b");
      const base64Data = graf.base64.split(',')[1];
      const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', graf.titulo + ".png");
      slide.insertImage(blob, 50, 100, 600, 300);
    });

    return { url: presentation.getUrl(), erro: null };
  } catch (e) { return { url: null, erro: e.message }; }
}

// ====================================================================================
// MOTOR GRAMATICAL E SLA
// ====================================================================================
function aplicarTemplate(dict, chave, nomeCliente, identificadorVeiculo, isPlural) {
  let txt = dict[chave] || "⚠️ Erro: Template não encontrado.";
  let textoFinal = txt.replace(/{{NOME}}/g, nomeCliente).replace(/{{VEICULO}}/g, identificadorVeiculo);

  if (isPlural) {
    const mapaPlural = [
      [/do seu veículo/gi, "dos seus veículos"], [/o seu veículo/gi, "os seus veículos"],
      [/em seu veículo/gi, "em seus veículos"], [/seu veículo/gi, "seus veículos"],
      [/a instalação do rastreador/gi, "a instalação dos rastreadores"], [/do rastreador/gi, "dos rastreadores"],
      [/do equipamento/gi, "dos equipamentos"], [/o equipamento não for instalado/gi, "os equipamentos não forem instalados"],
      [/o rastreador ainda não/gi, "os rastreadores ainda não"], [/esteja instalado/gi, "estejam instalados"],
      [/um rastreador instalado/gi, "rastreadores instalados"], [/o veículo não contará/gi, "os veículos não contarão"],
      [/o veículo não estará/gi, "os veículos não estarão"], [/o veículo não se encontra/gi, "os veículos não se encontram"],
      [/permaneça assegurado/gi, "permaneçam assegurados"], [/o veículo/gi, "os veículos"]
    ];
    mapaPlural.forEach(p => textoFinal = textoFinal.replace(p[0], p[1]));
  }
  return textoFinal;
}

function obterFeriadosDoAno(ano) {
  const feriados = [];
  const calcularPascoa = (year) => {
    const a = year % 19, b = Math.floor(year / 100), c = year % 100;
    const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
    const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
    const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
    const m = Math.floor((a + 11 * h + 22 * l) / 451);
    const mes = Math.floor((h + l - 7 * m + 114) / 31);
    const dia = ((h + l - 7 * m + 114) % 31) + 1;
    return new Date(year, mes - 1, dia);
  };
  const pascoa = calcularPascoa(ano);
  const addDias = (data, dias) => { const nd = new Date(data.getTime()); nd.setDate(nd.getDate() + dias); return nd; };
  
  feriados.push(formatar(addDias(pascoa, -48)));
  feriados.push(formatar(addDias(pascoa, -47)));
  feriados.push(formatar(addDias(pascoa, -2)));
  feriados.push(formatar(addDias(pascoa, 60)));
  
  function formatar(d) { 
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0'); 
    return `${dd}/${mm}/${ano}`; 
  }
  
  const fixos = [ `01/01/${ano}`, `21/04/${ano}`, `23/04/${ano}`, `01/05/${ano}`, `24/06/${ano}`, `07/09/${ano}`, `12/10/${ano}`, `02/11/${ano}`, `15/11/${ano}`, `20/11/${ano}`, `22/11/${ano}`, `25/12/${ano}` ];
  return [...feriados, ...fixos];
}

function calcularDiasUteis(dataInicial, dataFinal, arrayFeriadosPersonalizadosTime) {
  let diasUteis = 0;
  let dataAtual = new Date(dataInicial.getTime());
  dataAtual.setHours(0, 0, 0, 0);
  let dFinal = new Date(dataFinal.getTime());
  dFinal.setHours(0, 0, 0, 0);
  const cacheFeriados = {};

  while (dataAtual < dFinal) {
    dataAtual.setDate(dataAtual.getDate() + 1);
    let diaSemana = dataAtual.getDay();
    if (diaSemana !== 0 && diaSemana !== 6) {
      const anoAtual = dataAtual.getFullYear();
      if (!cacheFeriados[anoAtual]) cacheFeriados[anoAtual] = obterFeriadosDoAno(anoAtual);
      const dd = String(dataAtual.getDate()).padStart(2, '0');
      const mm = String(dataAtual.getMonth() + 1).padStart(2, '0');
      const strDataAtual = `${dd}/${mm}/${anoAtual}`;
      
      let ehFeriado = cacheFeriados[anoAtual].includes(strDataAtual);
      if (!ehFeriado && arrayFeriadosPersonalizadosTime && arrayFeriadosPersonalizadosTime.includes(dataAtual.getTime())) ehFeriado = true;
      if (!ehFeriado) diasUteis++;
    }
  }
  return diasUteis;
}

// ====================================================================================
// MOTOR DE IMPORTAÇÃO E FILA GERAL
// ====================================================================================
function cadastrarLoteWeb(loteDeClientes) {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    const aba1 = ss.getSheets().find(s => s.getName().includes("1 -"));
    const aba2 = ss.getSheets().find(s => s.getName().includes("2 -"));
    const aba3 = ss.getSheets().find(s => s.getName().includes("3 -"));
    
    if (!aba1 || !aba2 || !aba3) return "❌ Erro: Abas de operação não encontradas.";

    let feriadosTime = [];
    const abaFeriados = ss.getSheetByName("Feriados");
    if (abaFeriados) {
      feriadosTime = abaFeriados.getRange("A2:A").getValues().map(r => r[0] instanceof Date ? r[0].getTime() : null).filter(r => r);
    }

    const chassisNoSistema = new Set();
    const placasNoSistema = new Set();

    ss.getSheets().filter(s => s.getName().includes("1 -") || s.getName().includes("2 -") || s.getName().includes("3 -") || s.getName().includes("4 -")).forEach(aba => {
      const dados = aba.getDataRange().getValues();
      for (let i = 1; i < dados.length; i++) {
        if (dados[i][MAPA_COLUNAS.CHASSI]) chassisNoSistema.add(String(dados[i][MAPA_COLUNAS.CHASSI]).trim().toUpperCase());
        if (dados[i][MAPA_COLUNAS.PLACA]) placasNoSistema.add(String(dados[i][MAPA_COLUNAS.PLACA]).trim().toUpperCase());
      }
    });

    const token = autenticarHINOVA();
    if (!token) return "❌ Erro: Falha na autenticação com a Hinova.";

    const requests = loteDeClientes.map(cli => {
      const vb = cli.chassi || cli.placa;
      const pb = cli.chassi ? "chassi" : "placa";
      return { url: `${SGA_CONFIG.URL_CONSULTA_BASE}${encodeURIComponent(vb)}/${pb}`, method: "get", headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true };
    });

    const responses = UrlFetchApp.fetchAll(requests);
    const qtdColunasParaInserir = Math.max(aba1.getLastColumn(), 20) - 1;
    const dtHoje = new Date();
    const dtHojeStr = Utilities.formatDate(dtHoje, Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    let contInseridos = 0, contDuplicados = 0, contIgnoradosStatus = 0;
    const lotesPorAba = { 1: [], 2: [], 3: [] };

    loteDeClientes.forEach((cliente, index) => {
      const chassiCli = String(cliente.chassi || "").trim().toUpperCase();
      const placaCli = String(cliente.placa || "").trim().toUpperCase();
      if ((chassiCli && chassisNoSistema.has(chassiCli)) || (placaCli && placasNoSistema.has(placaCli))) { contDuplicados++; return; }

      let classificacaoSGA = null;
      try {
        if (responses[index].getResponseCode() === 200) {
          const j = JSON.parse(responses[index].getContentText());
          if (j && j.length > 0) classificacaoSGA = String(j[0].codigo_classificacao);
        }
      } catch (e) { }

      if (classificacaoSGA !== "14") { contIgnoradosStatus++; return; }

      let dataCriacao = web_extrairDataParaCalculo(cliente.data) || dtHoje;
      let diasUteis = calcularDiasUteis(dataCriacao, dtHoje, feriadosTime);
      let etapaAlvo = (diasUteis >= 10) ? 3 : (diasUteis >= 5) ? 2 : 1;

      const novaLinha = new Array(qtdColunasParaInserir).fill("");
      novaLinha[0] = cliente.data || dtHojeStr;
      novaLinha[MAPA_COLUNAS.NOME - 1] = String(cliente.nome || "").trim().toUpperCase();
      novaLinha[MAPA_COLUNAS.PLACA - 1] = placaCli;
      novaLinha[MAPA_COLUNAS.CHASSI - 1] = chassiCli;
      novaLinha[MAPA_COLUNAS.FIPE - 1] = String(cliente.fipe || "").trim();
      novaLinha[MAPA_COLUNAS.EMAIL - 1] = String(cliente.email || "").trim().toLowerCase();
      novaLinha[MAPA_COLUNAS.TELEFONE - 1] = String(cliente.telefone || "").trim();
      
      lotesPorAba[etapaAlvo].push(novaLinha);
      contInseridos++;
      
      if (chassiCli) chassisNoSistema.add(chassiCli);
      if (placaCli) placasNoSistema.add(placaCli);
    });

    const inserirNaAba = (aba, matriz) => {
      if (matriz.length === 0) return;
      const nomes = aba.getRange("C1:C").getValues();
      let ultimaLinhaReal = 1;
      for (let j = nomes.length - 1; j >= 0; j--) {
        if (String(nomes[j][0]).trim() !== "") { ultimaLinhaReal = j + 1; break; }
      }
      aba.getRange(ultimaLinhaReal + 1, 2, matriz.length, qtdColunasParaInserir).setValues(matriz);
    };

    inserirNaAba(aba1, lotesPorAba[1]); 
    inserirNaAba(aba2, lotesPorAba[2]); 
    inserirNaAba(aba3, lotesPorAba[3]);

    let msg = `✅ Inteligência Processou!\n📥 ${contInseridos} roteados.`;
    if (contDuplicados > 0) msg += `\n⚠️ ${contDuplicados} duplicados ignorados.`;
    if (contIgnoradosStatus > 0) msg += `\n🚫 ${contIgnoradosStatus} barrados (Não Pendentes).`;
    return msg;

  } catch (e) { 
    return "❌ Erro Crítico no Motor: " + e.message;
  }
}

function web_obterFilaGeral() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const abas = ss.getSheets().filter(s => s.getName().includes("1 -") || s.getName().includes("2 -") || s.getName().includes("3 -"));

  const fila = [];
  const templatesDict = getTemplatesDict(ss);

  abas.forEach(aba => {
    const nomeAba = aba.getName();
    let numEtapa = nomeAba.includes("2 -") ? 2 : nomeAba.includes("3 -") ? 3 : 1;
    const ultimaLinha = aba.getLastRow();
    const ultimaColuna = aba.getLastColumn();
    if (ultimaLinha < 2 || ultimaColuna < 1) return;

    const range = aba.getRange(1, 1, ultimaLinha, ultimaColuna);
    const dados = range.getValues();
    const notas = range.getNotes();

    for (let i = 1; i < dados.length; i++) {
      const l = dados[i];
      const nome = l[MAPA_COLUNAS.NOME] ? String(l[MAPA_COLUNAS.NOME]).trim() : "";
      const placa = l[MAPA_COLUNAS.PLACA] ? String(l[MAPA_COLUNAS.PLACA]).trim() : "";
      const chassi = l[MAPA_COLUNAS.CHASSI] ? String(l[MAPA_COLUNAS.CHASSI]).trim() : "";
      
      if (!placa && !chassi && !nome) continue;

      const notaNome = (MAPA_COLUNAS.NOME < ultimaColuna && notas[i][MAPA_COLUNAS.NOME]) ? String(notas[i][MAPA_COLUNAS.NOME]) : "";
      const notaPlaca = (MAPA_COLUNAS.PLACA < ultimaColuna && notas[i][MAPA_COLUNAS.PLACA]) ? String(notas[i][MAPA_COLUNAS.PLACA]).toUpperCase() : "";
      const notaEmail = (MAPA_COLUNAS.EMAIL < ultimaColuna && notas[i][MAPA_COLUNAS.EMAIL]) ? String(notas[i][MAPA_COLUNAS.EMAIL]) : "";
      const notaEstado = (MAPA_COLUNAS.ESTADO < ultimaColuna && notas[i][MAPA_COLUNAS.ESTADO]) ? String(notas[i][MAPA_COLUNAS.ESTADO]) : "";

      let cidade = "", bairro = "";
      let tecnicoDisp = "", tecnicoDist = "", tecnicoTempo = "", tecnicoTipo = "Volante"; 

      if (notaEstado.includes("Cidade:")) {
        const parts = notaEstado.split("\n");
        cidade = parts[0] ? parts[0].replace("📍 Cidade:", "").trim() : "";
        bairro = parts[1] ? parts[1].replace("🏘️ Bairro:", "").trim() : "";
      }

      if (notaEstado.includes("🛰️ LOGÍSTICA")) {
        let logMatchNovo = notaEstado.match(/Atendimento: \[(.*?)\] "(.*?)" - (.*?) \/ (.*?) de distância/);
        if (logMatchNovo) {
          tecnicoTipo = logMatchNovo[1];
          tecnicoDisp = logMatchNovo[2];
          tecnicoDist = logMatchNovo[3];
          tecnicoTempo = logMatchNovo[4];
        } else {
          let logMatch = notaEstado.match(/Técnico Disponível: "(.*?)" - (.*?) \/ (.*?) de distância/);
          if (logMatch) {
            tecnicoDisp = logMatch[1];
            tecnicoDist = logMatch[2];
            tecnicoTempo = logMatch[3];
          }
        }
      }

      const telefone = l[MAPA_COLUNAS.TELEFONE] ? String(l[MAPA_COLUNAS.TELEFONE]).trim() : "";
      let msgWhats = "";
      if (telefone) {
        const idVeic = placa || chassi;
        const isPlural = String(idVeic).includes(",") || String(idVeic).includes(" e ");
        let chaveCorpo = numEtapa === 1 ? (l[MAPA_COLUNAS.FIPE_BAIXA] === true ? "BOAS_VINDAS_FIPE_BAIXA" : "BOAS_VINDAS_NORMAL") : numEtapa === 2 ? "LEMBRETE_5_DIAS" : "PRAZO_EXPIRADO";

        let txtCorpo = aplicarTemplate(templatesDict, chaveCorpo, nome || "Cliente", idVeic, isPlural);
        let disclaimer = aplicarTemplate(templatesDict, "WHATSAPP_DISCLAIMER", nome || "Cliente", idVeic, false);
        msgWhats = (disclaimer && !disclaimer.includes("⚠️")) ? disclaimer + "\n\n" + txtCorpo : txtCorpo;
      }

      fila.push({
        idUnico: nomeAba + "-" + (i + 1), 
        etapaNum: numEtapa, linhaOriginal: i + 1, abaNome: nomeAba, nome: nome, placa: placa, chassi: chassi,
        fipe: l[MAPA_COLUNAS.FIPE] ? String(l[MAPA_COLUNAS.FIPE]).trim() : "",
        email: l[MAPA_COLUNAS.EMAIL] ? String(l[MAPA_COLUNAS.EMAIL]).trim() : "",
        telefone: telefone,
        estado: l[MAPA_COLUNAS.ESTADO] ? String(l[MAPA_COLUNAS.ESTADO]).trim() : "",
        cidade: cidade, bairro: bairro,
        tecnicoDisp: tecnicoDisp, tecnicoDist: tecnicoDist, tecnicoTempo: tecnicoTempo, 
        tecnicoTipo: tecnicoTipo, 
        dataPlanilha: (l[MAPA_COLUNAS.DATA] instanceof Date) ? Utilities.formatDate(l[MAPA_COLUNAS.DATA], Session.getScriptTimeZone(), "dd/MM/yyyy") : String(l[MAPA_COLUNAS.DATA] || "").split(" ")[0],
        dataEnvio: web_formatarDataSegura(l[MAPA_COLUNAS.DATA_EMAIL] || l[MAPA_COLUNAS.DATA_WHATS]),
        isEnviado: (l[MAPA_COLUNAS.CHECK_EMAIL] === true || l[MAPA_COLUNAS.CHECK_EMAIL] === "TRUE" || l[MAPA_COLUNAS.CHECK_EMAIL] === 1),
        isRespondeuEmail: (l[MAPA_COLUNAS.RESPONDEU_EMAIL] === true || l[MAPA_COLUNAS.RESPONDEU_EMAIL] === "TRUE" || l[MAPA_COLUNAS.RESPONDEU_EMAIL] === 1),
        isRespondeuWhats: (l[MAPA_COLUNAS.RESPONDEU_WHATS] === true || l[MAPA_COLUNAS.RESPONDEU_WHATS] === "TRUE" || l[MAPA_COLUNAS.RESPONDEU_WHATS] === 1),
        isFipeBaixa: (l[MAPA_COLUNAS.FIPE_BAIXA] === true || l[MAPA_COLUNAS.FIPE_BAIXA] === "TRUE" || l[MAPA_COLUNAS.FIPE_BAIXA] === 1),
        isTecnicoIndisp: (l[MAPA_COLUNAS.TECNICO_INDISPONIVEL] === true || l[MAPA_COLUNAS.TECNICO_INDISPONIVEL] === "TRUE" || l[MAPA_COLUNAS.TECNICO_INDISPONIVEL] === 1),
        isMoto: notaPlaca.includes("MOTO"),
        isInativo: notaNome.includes("Situação SGA"),
        isErroEmail: notaEmail.includes("Erro:"),
        notaNome: notaNome, notaEmail: notaEmail, mensagemWhatsApp: msgWhats
      });
    }
  });
  return fila;
}

// ====================================================================================
// AUTOMAÇÕES DE LOTE (COM LOGÍSTICA ATUALIZADA)
// ====================================================================================
function processarItemLoteWeb(cli, comando) {
  const token = autenticarHINOVA();
  if (!token) return { status: 'erro', msg: 'Falha na autenticação Hinova.' };
  
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const aba = ss.getSheetByName(cli.abaNome);
  if (!aba) return { status: 'erro', msg: 'Aba não encontrada.' };
  
  const vb = cli.chassi || cli.placa;
  const pb = cli.chassi ? "chassi" : "placa";
  const linha = cli.linhaOriginal;
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const codMoto = ["3", "126", "115", "100", "105", "127", "116", "32", "33", "34", "35", "95", "96", "97"];

  try {
    if (comando === "logistica") {
      const est = String(aba.getRange(linha, MAPA_COLUNAS.ESTADO + 1).getValue()).trim().toUpperCase();
      const celulaEstado = aba.getRange(linha, MAPA_COLUNAS.ESTADO + 1);
      const notaEndereço = celulaEstado.getNote();
        
      if (notaEndereço && notaEndereço.indexOf("Cidade:") !== -1) {
        const ssTec = SpreadsheetApp.openById(ID_PLANILHA_TECNICOS);
        const listaTecnicos = ssTec.getSheets()[0].getDataRange().getValues().slice(1).filter(t => t[0]); 
          
        const cidadeMatch = notaEndereço.match(/Cidade:\s*(.*)/);
        const bairroMatch = notaEndereço.match(/Bairro:\s*(.*)/);
        const cidadeCli = cidadeMatch ? cidadeMatch[1].split('\n')[0].trim() : "";
        const bairroCli = bairroMatch ? bairroMatch[1].split('\n')[0].trim() : "";

        const enderecoDestino = `${bairroCli}, ${cidadeCli} - ${est}, Brasil`;

        let melhorDistancia = Infinity;
        let melhorTecnico = null;
        let melhorTempoSeg = 0;
        let melhorTipo = "Volante"; 

        listaTecnicos.forEach(tecnico => {
          try {
            const logradouro = tecnico[1] || "";
            const numero = tecnico[2] || "";
            const bairro = tecnico[3] || "";
            const cidade = tecnico[4] || "";
            const estado = tecnico[5] || "";
            const cep = tecnico[6] || "";
            const tipo = tecnico[8] ? String(tecnico[8]).trim() : "Volante"; 
            
            const origemCompleta = `${logradouro}, ${numero} - ${bairro}, ${cidade} - ${estado}, ${cep}, Brasil`;

            const direcoes = Maps.newDirectionFinder()
              .setOrigin(origemCompleta)
              .setDestination(enderecoDestino)
              .setMode(Maps.DirectionFinder.Mode.DRIVING)
              .getDirections();

            if (direcoes.routes && direcoes.routes.length > 0) {
              const rota = direcoes.routes[0].legs[0];
              if (rota.distance.value < melhorDistancia) {
                melhorDistancia = rota.distance.value;
                melhorTecnico = tecnico[0];
                melhorTempoSeg = rota.duration.value;
                melhorTipo = tipo; 
              }
            }
          } catch (e) {}
        });

        if (melhorTecnico) {
          const distKm = (melhorDistancia / 1000).toFixed(1);
          const h = Math.floor(melhorTempoSeg / 3600);
          const m = Math.floor((melhorTempoSeg % 3600) / 60);
          const tempoFormatado = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;

          const notaLimpa = notaEndereço.split("\n\n--- 🛰️ LOGÍSTICA ---")[0];
          celulaEstado.setNote(notaLimpa + "\n\n--- 🛰️ LOGÍSTICA ---\n" + `Atendimento: [${melhorTipo}] "${melhorTecnico}" - ${distKm} Km / ${tempoFormatado} de distância de carro`);

          return { status: 'ok', msg: 'Rota calculada' };
        }
      }
      return { status: 'ignorado', msg: 'Endereço inválido para roteirização' };
    }

    const resp = UrlFetchApp.fetch(`${SGA_CONFIG.URL_CONSULTA_BASE}${encodeURIComponent(vb)}/${pb}`, { "method": "get", "headers": { "Authorization": "Bearer " + token }, "muteHttpExceptions": true });

    if (resp.getResponseCode() === 200) {
      const j = JSON.parse(resp.getContentText());

      if (j && j.length > 0) {
        const v = j[0];
        let alterado = false;

        if (comando === "motos") {
          if (codMoto.includes(String(v.codigo_tipo_veiculo))) {
            aba.getRange(linha, 1, 1, aba.getLastColumn()).setBackground("#d1fae5");
            aba.getRange(linha, MAPA_COLUNAS.PLACA + 1).setNote("🏍️ MOTO (SGA)");
            alterado = true;
          } else {
            aba.getRange(linha, 1, 1, aba.getLastColumn()).setBackground(null);
            aba.getRange(linha, MAPA_COLUNAS.PLACA + 1).clearNote();
          }
        } else if (comando === "inativos" && v.codigo_situacao != null) {
          const cSit = String(v.codigo_situacao);
          const cNome = aba.getRange(linha, MAPA_COLUNAS.NOME + 1);
          if (cSit === "1" || cSit === "14") {
            cNome.setFontColor("#000000").setFontWeight("normal").clearNote();
          } else {
            cNome.setFontColor("#9C27B0").setFontWeight("bold").setNote(`⚠️ Situação SGA: ${MAPA_SITUACAO_SGA[cSit] || "Desconhecida"}\nVerificado: ${dt}`);
            alterado = true;
          }
        } else if (comando === "fipe") {
          const fTexto = v.valor_fipe || "";
          const f = parseFloat(String(fTexto).replace(/\./g, '').replace(',', '.')) || 0;
          if (fTexto) aba.getRange(linha, MAPA_COLUNAS.FIPE + 1).setValue(fTexto);

          if ((codMoto.includes(String(v.codigo_tipo_veiculo)) && f > 0 && f < 20000) || (!codMoto.includes(String(v.codigo_tipo_veiculo)) && f > 0 && f < 30000)) {
            aba.getRange(linha, MAPA_COLUNAS.FIPE_BAIXA + 1).setValue(true);
            alterado = true;
          } else {
            aba.getRange(linha, MAPA_COLUNAS.FIPE_BAIXA + 1).setValue(false);
            alterado = true;
          }
        } else if (comando === "completar") {
          const dadosLinha = aba.getRange(linha, 1, 1, aba.getLastColumn()).getValues()[0];
          if (!dadosLinha[MAPA_COLUNAS.NOME - 1] && v.nome) { aba.getRange(linha, MAPA_COLUNAS.NOME + 1).setValue(String(v.nome).trim()); alterado = true; }
          if (!dadosLinha[MAPA_COLUNAS.EMAIL - 1] && v.email) { aba.getRange(linha, MAPA_COLUNAS.EMAIL + 1).setValue(String(v.email).toLowerCase().trim()); alterado = true; }
          if (!dadosLinha[MAPA_COLUNAS.TELEFONE - 1] && v.telefone_celular) { aba.getRange(linha, MAPA_COLUNAS.TELEFONE + 1).setValue(`(${v.ddd_celular || ""}) ${v.telefone_celular}`); alterado = true; }
          if (!dadosLinha[MAPA_COLUNAS.FIPE - 1] && v.valor_fipe) { aba.getRange(linha, MAPA_COLUNAS.FIPE + 1).setValue(v.valor_fipe); alterado = true; }
        
        } else if (comando === "estados") {
          let est = v.estado ? String(v.estado).trim().toUpperCase() : "";
          let cid = v.cidade ? String(v.cidade).trim() : "";
          let bai = v.bairro ? String(v.bairro).trim() : "";

          if ((!est || est === "N/A" || !cid) && v.codigo_associado) {
            const rA = UrlFetchApp.fetch(`https://api.hinova.com.br/api/sga/v2/associado/buscar/${v.codigo_associado}/codigo`, { "method": "get", "headers": { "Authorization": "Bearer " + token }, "muteHttpExceptions": true });

            if (rA.getResponseCode() === 200) {
              const jA = JSON.parse(rA.getContentText());
              const aD = (Array.isArray(jA) && jA.length > 0) ? jA[0] : jA;
              
              est = aD.estado ? String(aD.estado).trim().toUpperCase() : "N/A";
              cid = aD.cidade ? String(aD.cidade).trim() : "";
              bai = aD.bairro ? String(aD.bairro).trim() : "";
            }
          }

          if (!est) est = "N/A";

          aba.getRange(linha, MAPA_COLUNAS.ESTADO + 1).setValue(est).setNote(`📍 Cidade: ${cid}\n🏘️ Bairro: ${bai}`);
          
          if (est !== "RJ" && est !== "N/A") {
            aba.getRange(linha, MAPA_COLUNAS.TECNICO_INDISPONIVEL + 1).setValue(true);
            alterado = true;
          } else {
             alterado = true;
          }
        }
        return { status: alterado ? 'ok' : 'ignorado', msg: 'Processado' };
      }
    }
    return { status: 'ignorado', msg: 'Nenhum dado na consulta' };

  } catch (e) { 
    return { status: 'erro', msg: e.message };
  }
}

// ====================================================================================
// EDIÇÃO MANUAL DO CLIENTE (VIA FILA)
// ====================================================================================
function web_atualizarDadosCliente(abaNome, linha, dados) {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    const aba = ss.getSheetByName(abaNome);
    if (!aba) return "❌ Erro: Aba não encontrada no banco.";

    if (dados.nome) aba.getRange(linha, MAPA_COLUNAS.NOME + 1).setValue(String(dados.nome).toUpperCase());
    if (dados.placa) aba.getRange(linha, MAPA_COLUNAS.PLACA + 1).setValue(String(dados.placa).toUpperCase());
    if (dados.chassi) aba.getRange(linha, MAPA_COLUNAS.CHASSI + 1).setValue(String(dados.chassi).toUpperCase());
    if (dados.email) aba.getRange(linha, MAPA_COLUNAS.EMAIL + 1).setValue(String(dados.email).toLowerCase());
    if (dados.telefone) aba.getRange(linha, MAPA_COLUNAS.TELEFONE + 1).setValue(String(dados.telefone));
    if (dados.fipe) aba.getRange(linha, MAPA_COLUNAS.FIPE + 1).setValue(String(dados.fipe));

    return "✅ Dados cadastrais atualizados com sucesso!";
  } catch (e) {
    return "❌ Erro ao atualizar: " + e.message;
  }
}

// ====================================================================================
// OBTENÇÃO DE PRÉVIA E DISPARO CANAL A CANAL
// ====================================================================================
function obterPreviewDisparoAgrupadoWeb(grupos) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const templatesDict = getTemplatesDict(ss);

  return grupos.map(g => {
    let ass = "", txt = "", tituloHeader = "";
    const isPlural = g.veiculos.length > 1;
    const lblVeic = isPlural ? "Veículos" : "Veículo";

    if (g.etapaNum === 1) {
      tituloHeader = "BEM-VINDO À ZEN SEGUROS";
      ass = `Bem-vindo à ZEN Seguros - Orientações - ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, g.isFipeBaixa ? "BOAS_VINDAS_FIPE_BAIXA" : "BOAS_VINDAS_NORMAL", g.nome, g.veiculosStr, isPlural);
    } else if (g.etapaNum === 2) {
      tituloHeader = "LEMBRETE: INSTALAÇÃO PENDENTE";
      ass = `Lembrete: Instalação Pendente - ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, "LEMBRETE_5_DIAS", g.nome, g.veiculosStr, isPlural);
    } else {
      tituloHeader = "URGENTE: PRAZO EXPIRADO";
      ass = `[URGENTE] Prazo Expirado! ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, "PRAZO_EXPIRADO", g.nome, g.veiculosStr, isPlural);
    }

    const htmlBodyFormatado = formatarComoEmail(txt, tituloHeader);
    let cabecalhoWhatsApp = templatesDict["WHATSAPP_DISCLAIMER"] || "";
    let msgWhats = "";

    if (cabecalhoWhatsApp && cabecalhoWhatsApp.indexOf("⚠️ Erro:") === -1) {
      msgWhats = cabecalhoWhatsApp + "\n\n" + txt;
    } else {
      msgWhats = "> *MENSAGEM AUTOMÁTICA*\n> _Esse WhatsApp é utilizado apenas para envio de recados_\n> _Nossos contatos estarão disponíveis no final da mensagem_\n\n" + txt;
    }

    let telefoneBase = g.linhas[0].telefone || "";
    let numeroLimpo = telefoneBase.toString().replace(/\D/g, "");

    if (numeroLimpo.length >= 10 && !numeroLimpo.startsWith("55")) numeroLimpo = "55" + numeroLimpo;

    return { 
      email: g.email, nome: g.nome, veiculosStr: g.veiculosStr, 
      etapaNum: g.etapaNum, assunto: ass, emailHtml: htmlBodyFormatado, 
      whatsText: msgWhats, telefoneLimpo: numeroLimpo,
      isErroEmail: g.isErroEmail, isEnviado: g.isEnviado, isInativo: g.isInativo 
    };
  });
}

function dispararEmailAgrupadoWeb(grupos, responsavel) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const templatesDict = getTemplatesDict(ss);
  let errosCriticos = [];

  grupos.forEach(g => {
    let ass = "", txt = "", ac = "", tituloHeader = "";
    const isPlural = g.veiculos.length > 1;
    const lblVeic = isPlural ? "Veículos" : "Veículo";

    if (g.etapaNum === 1) {
      tituloHeader = "BEM-VINDO À ZEN SEGUROS";
      ass = `Bem-vindo à ZEN Seguros - Orientações - ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, g.isFipeBaixa ? "BOAS_VINDAS_FIPE_BAIXA" : "BOAS_VINDAS_NORMAL", g.nome, g.veiculosStr, isPlural);
      ac = "1_EMAIL";
    } else if (g.etapaNum === 2) {
      tituloHeader = "LEMBRETE: INSTALAÇÃO PENDENTE";
      ass = `Lembrete: Instalação Pendente - ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, "LEMBRETE_5_DIAS", g.nome, g.veiculosStr, isPlural);
      ac = "2_EMAIL";
    } else {
      tituloHeader = "URGENTE: PRAZO EXPIRADO";
      ass = `[URGENTE] Prazo Expirado! ${lblVeic}: ${g.veiculosStr}`;
      txt = aplicarTemplate(templatesDict, "PRAZO_EXPIRADO", g.nome, g.veiculosStr, isPlural);
      ac = "3_EMAIL";
    }

    try {
      const htmlBodyFormatado = formatarComoEmail(txt, tituloHeader);
      MailApp.sendEmail({
        to: g.email, subject: ass,
        body: txt + "\n\nAtenciosamente,\nSetor de Rastreamento\nZEN Seguros",
        htmlBody: htmlBodyFormatado, name: "Setor de Rastreamento - ZEN Seguros"
      });

      g.linhas.forEach(cli => {
        const aba = ss.getSheetByName(cli.abaNome);
        if (aba) {
          aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.CHECK_EMAIL + 1).setValue(true);
          aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.DATA_EMAIL + 1).setValue(dt);
          aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.RESPONSAVEL + 1).setValue(responsavel);
          registrarAuditoriaExata(cli.nome, cli.placa, cli.chassi, cli.email, cli.telefone, ac, dt, responsavel);
        }
      });

    } catch (e) {
      errosCriticos.push(`Destino [${g.email}]\nErro Final: ${e.message}`);
      g.linhas.forEach(cli => {
        const aba = ss.getSheetByName(cli.abaNome);
        if (aba) sinalizarErroEmail(aba, cli.linhaOriginal, "Falha Envio: " + e.message, dt);
      });
    }
  });

  if (errosCriticos.length > 0) throw new Error("\n" + errosCriticos.join("\n\n"));
  return `E-mail enviado com sucesso!`;
}

function marcarWhatsAgrupadoWeb(grupos, responsavel) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  grupos.forEach(g => {
    let ac = g.etapaNum === 1 ? "1_WHATS" : g.etapaNum === 2 ? "2_WHATS" : "3_WHATS";
    g.linhas.forEach(cli => {
      const aba = ss.getSheetByName(cli.abaNome);
      if (aba) {
        aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.CHECK_WHATS + 1).setValue(true);
        aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.DATA_WHATS + 1).setValue(dt);
        aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.RESPONSAVEL + 1).setValue(responsavel);
        registrarAuditoriaExata(cli.nome, cli.placa, cli.chassi, cli.email, cli.telefone, ac, dt, responsavel);
      }
    });
  });
  return "WhatsApp marcado na Planilha e Auditoria!";
}

function formatarComoEmail(textoHtmlOriginal, tituloEmail) {
  var textoHTML = textoHtmlOriginal.replace(/\n/g, '<br>');
  var htmlFinal = `
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #333333; max-width: 600px; margin: 0; line-height: 1.6;">
      <h3 style="color: #333333; margin-bottom: 20px; font-weight: bold; text-transform: uppercase;">${tituloEmail}</h3>
      <div style="margin-bottom: 20px;">${textoHTML}</div>
      <div style="border-top: 1px solid #dddddd; padding-top: 15px; margin-top: 20px; font-size: 13px; line-height: 1.5; color: #666666;">
        <img src="https://www.zensegurosbr.com/uploads/images/configuracoes/redimencionar-230-78-logo.png" width="160" alt="ZEN Seguros" style="display: block; margin-bottom: 8px; border: none; outline: none; text-decoration: none;">
        Atenciosamente,<br><strong style="color: #444444; font-size: 14px;">Setor de Rastreamento</strong><br>ZEN Seguros
      </div>
      <div style="display:none; color:transparent; font-size:1px;">Anti-Spam ID: ${new Date().getTime()}</div>
    </div>
  `;
  return htmlFinal;
}

// ====================================================================================
// ROTINAS GLOBAIS E UTILITÁRIOS
// ====================================================================================
function executarFerramentaWebGlobal(comando) {
  try {
    if (comando === "auditar_aba4") { auditarVisualAba4(); return "✅ Auditoria Aba 4 concluída!"; }
    if (comando === "sincronizar_erros") { conciliarErrosMailerDaemon(); return "✅ Caixa do Gmail mapeada!"; }
    if (comando === "varrer_concluidos") { return varrerConcluidosGlobalWeb(); }
    return "⚠️ Comando não reconhecido.";
  } catch (e) { return "❌ Erro API: " + e.message; }
}

function autenticarHINOVA() {
  const cache = CacheService.getScriptCache();
  const tokenCache = cache.get("HINOVA_TOKEN");
  if (tokenCache) return tokenCache;

  try {
    const options = {
      "method": "post",
      "headers": { "Authorization": "Bearer " + SGA_CONFIG.TOKEN_ASSOCIACAO, "Content-Type": "application/json" },
      "payload": JSON.stringify({ "usuario": SGA_CONFIG.USUARIO, "senha": SGA_CONFIG.SENHA }),
      "muteHttpExceptions": true
    };

    const resp = UrlFetchApp.fetch(SGA_CONFIG.URL_AUTH, options);
    if (resp.getResponseCode() !== 200) return null;
    
    const token = JSON.parse(resp.getContentText()).token_usuario || null;
    if (token) cache.put("HINOVA_TOKEN", token, 3600);
    return token;
  } catch (e) { return null; }
}

function varrerConcluidosGlobalWeb() {
  const token = autenticarHINOVA();
  if (!token) return "❌ Falha na autenticação com a Hinova.";

  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  let contConcluidos = 0;

  let logSheet = ss.getSheetByName("Log Concluídos") || ss.insertSheet("Log Concluídos");
  const todosClientes = [];

  ["1 -", "2 -", "3 -"].forEach(nomeFrag => {
    const aba = ss.getSheets().find(s => s.getName().includes(nomeFrag));
    if (!aba) return;
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      const placa = String(dados[i][MAPA_COLUNAS.PLACA] || "").trim().toUpperCase();
      const chassi = String(dados[i][MAPA_COLUNAS.CHASSI] || "").trim().toUpperCase();
      if (placa || chassi) {
        todosClientes.push({
          abaNome: aba.getName(), linha: i + 1, placa: placa, chassi: chassi,
          nome: dados[i][MAPA_COLUNAS.NOME], fipe: dados[i][MAPA_COLUNAS.FIPE],
          email: dados[i][MAPA_COLUNAS.EMAIL], telefone: dados[i][MAPA_COLUNAS.TELEFONE],
          identificador: chassi || placa, tipoBusca: chassi ? "chassi" : "placa"
        });
      }
    }
  });

  if (todosClientes.length === 0) return "Nenhum cliente na fila para varrer.";

  const requests = todosClientes.map(cli => ({
    url: `${SGA_CONFIG.URL_CONSULTA_BASE}${encodeURIComponent(cli.identificador)}/${cli.tipoBusca}`, method: "get", headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true
  }));

  const responses = UrlFetchApp.fetchAll(requests);
  const linhasParaDeletar = {};
  const dadosParaLog = [];

  todosClientes.forEach((cli, index) => {
    try {
      if (responses[index].getResponseCode() === 200) {
        const j = JSON.parse(responses[index].getContentText());
        if (j && j.length > 0 && String(j[0].codigo_classificacao) === "1") {
          if (!linhasParaDeletar[cli.abaNome]) linhasParaDeletar[cli.abaNome] = [];
          linhasParaDeletar[cli.abaNome].push(cli.linha);
          dadosParaLog.push([dt, cli.nome, cli.placa, cli.chassi, cli.fipe, cli.email, cli.telefone, cli.abaNome]);
          contConcluidos++;
        }
      }
    } catch (e) { }
  });

  if (dadosParaLog.length > 0) {
    const ultimaLinhaLog = logSheet.getLastRow() || 1;
    logSheet.getRange(ultimaLinhaLog + 1, 1, dadosParaLog.length, dadosParaLog[0].length).setValues(dadosParaLog);
  }
  
  for (const abaNome in linhasParaDeletar) {
    const aba = ss.getSheetByName(abaNome);
    if (aba) {
      linhasParaDeletar[abaNome].sort((a, b) => b - a);
      linhasParaDeletar[abaNome].forEach(linha => aba.deleteRow(linha));
    }
  }
  return `✅ Varredura Completa na API!\n${contConcluidos} veículos "Concluídos" encontrados e movidos para o Log.`;
}

function atualizarMarcacaoWeb(abaNome, linha, campo, valorBooleano) {
  try {
    const aba = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(abaNome);
    if (!aba) return "Aba não encontrada.";
    let col = campo === 'fipeBaixa' ? MAPA_COLUNAS.FIPE_BAIXA + 1 : campo === 'tecnicoIndisp' ? MAPA_COLUNAS.TECNICO_INDISPONIVEL + 1 : campo === 'respEmail' ? MAPA_COLUNAS.RESPONDEU_EMAIL + 1 : campo === 'respWhats' ? MAPA_COLUNAS.RESPONDEU_WHATS + 1 : 0;
    if (col > 0) { 
      aba.getRange(linha, col).setValue(valorBooleano);
      return "✅ Salvo na planilha!"; 
    }
    return "Campo inválido.";
  } catch (e) { return "❌ Erro: " + e.message; }
}

function marcarComoEnviadoWeb(clientesSelecionados, responsavel) {
  if (!clientesSelecionados || clientesSelecionados.length === 0) return "Vazio.";
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  let cont = 0;

  clientesSelecionados.forEach(cli => {
    const aba = ss.getSheetByName(cli.abaNome);
    if (!aba) return;
    aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.CHECK_EMAIL + 1).setValue(true);
    aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.DATA_EMAIL + 1).setValue(dt);
    aba.getRange(cli.linhaOriginal, MAPA_COLUNAS.RESPONSAVEL + 1).setValue(responsavel);
    const chaveAcao = cli.etapaNum === 1 ? "1_EMAIL" : cli.etapaNum === 2 ? "2_EMAIL" : "3_EMAIL";
    registrarAuditoriaExata(cli.nome, cli.placa, cli.chassi, cli.email, cli.telefone, chaveAcao, dt, responsavel);
    cont++;
  });
  return `✅ ${cont} marcados na planilha!`;
}

function web_obterConfiguracoes() {
  const aba = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName("⚙️ Configurações");
  if (!aba) return [];
  const dados = aba.getDataRange().getValues();
  const configs = [];
  for (let i = 1; i < dados.length; i++) { 
    if (dados[i][0]) configs.push({ linhaOriginal: i + 1, chave: String(dados[i][0]), texto: dados[i][1] ? String(dados[i][1]) : "" });
  }
  return JSON.parse(JSON.stringify(configs));
}

function salvarConfiguracaoWeb(linha, novoTexto) {
  try { 
    SpreadsheetApp.openById(PLANILHA_ID).getSheetByName("⚙️ Configurações").getRange(linha, 2).setValue(novoTexto); 
    return "✅ Template atualizado!";
  } catch (e) { return "❌ Erro: " + e.message; }
}

// ====================================================================================
// AUDITORIA E LOGS
// ====================================================================================
function sinalizarErroEmail(aba, numeroLinha, motivo, dataHora) {
  var ss = SpreadsheetApp.openById(PLANILHA_ID);
  var abaErro = ss.getSheetByName("Erro") || ss.insertSheet("Erro");
  var dados = aba.getRange(numeroLinha, 1, 1, aba.getLastColumn()).getValues()[0];
  var emailAtual = dados[MAPA_COLUNAS.EMAIL] ? String(dados[MAPA_COLUNAS.EMAIL]).toLowerCase().trim() : "";
  
  if (emailAtual !== "") {
    var linhaParaErro = dados.slice(); 
    linhaParaErro.push("FALHA: " + motivo); 
    linhaParaErro.push(dataHora);
    abaErro.appendRow(linhaParaErro);
  }
  aba.getRange(numeroLinha, MAPA_COLUNAS.EMAIL + 1).setFontColor("#FF0000").setFontWeight("bold").setNote("⚠️ Erro: " + motivo);
}

function registrarAuditoriaExata(nome, placa, chassi, email, telefone, chaveAcao, dataHora, responsavel) {
  var abaAuditoria = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName("4 -Registro - NÃO ALTERAR");
  if (!abaAuditoria) return;

  const MAPA_AUDITORIA = { 
    "1_EMAIL": { check: "1- E-mail (boas vindas)", data: "1- Enviado e-mail em:", resp: "1-Responsável (boas vindas)" }, 
    "1_WHATS": { check: "1-Whatsapp (boas vindas)", data: "1 -Enviado whats em:", resp: "1-Responsável (boas vindas)" }, 
    "2_EMAIL": { check: "2- E-mail (5 dias)", data: "2- Enviado e-mail em:", resp: "2-Responsável (5 dias)" }, 
    "2_WHATS": { check: "2-Whatsapp (5 dias)", data: "2 -Enviado whats em:", resp: "2-Responsável (5 dias)" }, 
    "3_EMAIL": { check: "3- E-mail (prazo)", data: "3- Enviado e-mail em:", resp: "3-Responsável (prazo)" }, 
    "3_WHATS": { check: "3-Whatsapp (prazo)", data: "3 -Enviado whats em:", resp: "3-Responsável (prazo)" } 
  };

  if (!MAPA_AUDITORIA[chaveAcao]) return;

  var cabecalho = abaAuditoria.getRange(1, 1, 1, abaAuditoria.getLastColumn()).getValues()[0].map(x => x ? String(x).trim() : "");
  var cChk = cabecalho.indexOf(MAPA_AUDITORIA[chaveAcao].check.trim()), 
      cDat = cabecalho.indexOf(MAPA_AUDITORIA[chaveAcao].data.trim()), 
      cRes = cabecalho.indexOf(MAPA_AUDITORIA[chaveAcao].resp.trim());
      
  var cNom = cabecalho.indexOf("Nome"), 
      cPla = cabecalho.indexOf("Placa"), 
      cCha = cabecalho.indexOf("Chassi"), 
      cEma = cabecalho.indexOf("E-mail") > -1 ? cabecalho.indexOf("E-mail") : cabecalho.indexOf("Email"), 
      cTel = cabecalho.indexOf("Telefone");
      
  var uLinha = abaAuditoria.getLastRow(), 
      dados = uLinha > 1 ? abaAuditoria.getRange(2, 1, uLinha - 1, abaAuditoria.getLastColumn()).getValues() : [];
  var lAlvo = -1;

  for (var i = 0; i < dados.length; i++) { 
    if ((placa && String(dados[i][cPla]).trim() === placa) || (chassi && String(dados[i][cCha]).trim() === chassi)) { 
      lAlvo = i + 2;
      break; 
    } 
  }

  if (lAlvo === -1) { 
    lAlvo = uLinha + 1;
    if (cNom > -1) abaAuditoria.getRange(lAlvo, cNom + 1).setValue(nome); 
    if (cPla > -1) abaAuditoria.getRange(lAlvo, cPla + 1).setValue(placa);
    if (cCha > -1) abaAuditoria.getRange(lAlvo, cCha + 1).setValue(chassi); 
    if (cEma > -1) abaAuditoria.getRange(lAlvo, cEma + 1).setValue(email);
    if (cTel > -1) abaAuditoria.getRange(lAlvo, cTel + 1).setValue(telefone); 
  }

  if (cChk > -1) abaAuditoria.getRange(lAlvo, cChk + 1).setValue(true);
  if (cDat > -1) abaAuditoria.getRange(lAlvo, cDat + 1).setValue(dataHora); 
  if (cRes > -1 && responsavel) abaAuditoria.getRange(lAlvo, cRes + 1).setValue(responsavel);
}

function conciliarErrosMailerDaemon() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID); 
  const abaErro = ss.getSheetByName("Erro") || ss.insertSheet("Erro"); 
  let threads;

  try { 
    threads = GmailApp.search('(from:mailer-daemon OR from:postmaster) subject:("Delivery" OR "Failure" OR "Falha" OR "Undeliverable" OR "Returned" OR "Undelivered") newer_than:2d', 0, 50);
  } catch (e) { return; }
  if (threads.length === 0) return;
  const errosDet = {};

  threads.forEach(t => { 
    t.getMessages().forEach(m => { 
      const c = m.getPlainBody(); 
      const match = c.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/g); 
      if (match) { 
        match.forEach(em => { 
          let e = em.toLowerCase().trim(); 
          if (!e.includes("postmaster") && !e.includes("mailer-daemon")) errosDet[e] = "Falha na entrega"; 
        }); 
      } 
    }); 
  });

  const abas = ss.getSheets();
  for (let a = 0; a < abas.length; a++) {
    if (abas[a].getName().includes("1 -") || abas[a].getName().includes("2 -") || abas[a].getName().includes("3 -")) {
      const d = abas[a].getDataRange().getValues();
      for (let j = d.length - 1; j >= 1; j--) { 
        const em = d[j][MAPA_COLUNAS.EMAIL] ? String(d[j][MAPA_COLUNAS.EMAIL]).toLowerCase().trim() : "";
        if (em && errosDet[em]) { 
          abas[a].getRange(j + 1, MAPA_COLUNAS.EMAIL + 1).setFontColor("#FF0000").setFontWeight("bold").setNote("⚠️ Erro: " + errosDet[em]);
        } 
      }
    }
  }
}

function auditarVisualAba4() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID), a4 = ss.getSheetByName("4 -Registro - NÃO ALTERAR");
  if (!a4) return;
  const aConc = ss.getSheetByName("Log Concluídos"), aSit = ss.getSheetByName("6 -Situação"), aErr = ss.getSheetByName("Erro");
  const sC = new Set(), sI = new Set(), sE = new Set();
  
  if (aConc) { 
    const d = aConc.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) { 
      if (d[i][2]) sC.add(String(d[i][2]).trim().toUpperCase()); 
      if (d[i][3]) sC.add(String(d[i][3]).trim().toUpperCase());
    } 
  }
  if (aSit) { 
    const d = aSit.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) { 
      const c = String(d[i][6]).trim();
      if (c && c !== "1" && c !== "14") { 
        if (d[i][2]) sI.add(String(d[i][2]).trim().toUpperCase()); 
        if (d[i][3]) sI.add(String(d[i][3]).trim().toUpperCase());
      } 
    } 
  }
  if (aErr) { 
    const d = aErr.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) { 
      const e = String(d[i][MAPA_COLUNAS.EMAIL]).trim().toLowerCase(); 
      if (e) sE.add(e);
    } 
  }
  
  const d4 = a4.getDataRange().getValues(); 
  if (d4.length < 2) return;
  const cab = d4[0].map(c => String(c).trim()), iN = cab.indexOf("Nome"), iP = cab.indexOf("Placa"), iC = cab.indexOf("Chassi"), iE = cab.findIndex(c => c === "E-mail" || c === "Email");
  if (iN === -1) return;
  
  const cF = [], pF = [];
  for (let i = 1; i < d4.length; i++) {
    const p = iP > -1 && d4[i][iP] ? String(d4[i][iP]).trim().toUpperCase() : "";
    const c = iC > -1 && d4[i][iC] ? String(d4[i][iC]).trim().toUpperCase() : "";
    const e = iE > -1 && d4[i][iE] ? String(d4[i][iE]).trim().toLowerCase() : "";

    let cr = "#000000", ps = "normal";
    if ((p && sC.has(p)) || (c && sC.has(c))) { 
      cr = "#2E7D32"; ps = "bold"; 
    } else if ((p && sI.has(p)) || (c && sI.has(c))) { 
      cr = "#9C27B0"; ps = "bold";
    } else if (e && sE.has(e)) { 
      cr = "#FF0000"; ps = "bold"; 
    }
    cF.push([cr]); pF.push([ps]);
  }
  if (cF.length > 0) { 
    const rN = a4.getRange(2, iN + 1, cF.length, 1); 
    rN.setFontColors(cF); 
    rN.setFontWeights(pF);
  }
}

function web_obterDadosLogs(nomeDaAba) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID); 
  const aba = ss.getSheetByName(nomeDaAba);
  if (!aba) return { colunas: [], linhas: [] };
  const dados = aba.getDataRange().getValues();
  if (dados.length < 2) return { colunas: [], linhas: [] };

  if (nomeDaAba === "4 -Registro - NÃO ALTERAR") {
    const aConc = ss.getSheetByName("Log Concluídos"), aSit = ss.getSheetByName("6 -Situação"), aErr = ss.getSheetByName("Erro");
    const sC = new Set(), sI = new Set(), sE = new Set();
    
    if (aConc) { 
      const d = aConc.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) { 
        if (d[i][2]) sC.add(String(d[i][2]).trim().toUpperCase()); 
        if (d[i][3]) sC.add(String(d[i][3]).trim().toUpperCase());
      } 
    }
    if (aSit) { 
      const d = aSit.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) { 
        const c = String(d[i][6]).trim();
        if (c && c !== "1" && c !== "14") { 
          if (d[i][2]) sI.add(String(d[i][2]).trim().toUpperCase()); 
          if (d[i][3]) sI.add(String(d[i][3]).trim().toUpperCase());
        } 
      } 
    }
    if (aErr) { 
      const d = aErr.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) { 
        const e = String(d[i][MAPA_COLUNAS.EMAIL]).trim().toLowerCase(); 
        if (e) sE.add(e);
      } 
    }
    
    const cab = dados[0].map(c => String(c).trim()), iN = cab.indexOf("Nome"), iP = cab.indexOf("Placa"), iC = cab.indexOf("Chassi"), iE = cab.findIndex(c => c === "E-mail" || c === "Email");
    const colunas = ["Status Principal", "Cliente", "Identificação", "Envios Processados", "Responsável pelo Envio"], linhasMatriz = [];
    let inicio = dados.length > 1001 ? dados.length - 1000 : 1;

    for (let i = dados.length - 1; i >= inicio; i--) {
      const p = iP > -1 && dados[i][iP] ? String(dados[i][iP]).trim().toUpperCase() : "";
      const c = iC > -1 && dados[i][iC] ? String(dados[i][iC]).trim().toUpperCase() : "";
      const e = iE > -1 && dados[i][iE] ? String(dados[i][iE]).trim().toLowerCase() : "";
      const nome = iN > -1 ? dados[i][iN] : "";
      const ident = (p || "---") + " / " + (c || "---");

      let envios = [], resps = new Set(), highestStage = 0;
      const mapEtapas = [ 
        { emj: "🔰", chkE: "1- E-mail (boas vindas)", datE: "1- Enviado e-mail em:", chkW: "1-Whatsapp (boas vindas)", datW: "1 -Enviado whats em:", resp: "1-Responsável (boas vindas)" }, 
        { emj: "⚠️", chkE: "2- E-mail (5 dias)", datE: "2- Enviado e-mail em:", chkW: "2-Whatsapp (5 dias)", datW: "2 -Enviado whats em:", resp: "2-Responsável (5 dias)" }, 
        { emj: "⛔", chkE: "3- E-mail (prazo)", datE: "3- Enviado e-mail em:", chkW: "3-Whatsapp (prazo)", datW: "3 -Enviado whats em:", resp: "3-Responsável (prazo)" } 
      ];

      mapEtapas.forEach((etp, index) => {
        let idxCE = cab.indexOf(etp.chkE), idxDE = cab.indexOf(etp.datE), idxR = cab.indexOf(etp.resp), idxCW = cab.indexOf(etp.chkW), idxDW = cab.indexOf(etp.datW);
        let dE = idxDE > -1 && dados[i][idxDE] ? web_formatarDataSegura(dados[i][idxDE]).split(" ")[0] : "";
        let dW = idxDW > -1 && dados[i][idxDW] ? web_formatarDataSegura(dados[i][idxDW]).split(" ")[0] : "";
        let responsavel = idxR > -1 && dados[i][idxR] ? String(dados[i][idxR]).trim() : "";
        let hasEnvio = false;
        
        if ((idxCE > -1 && dados[i][idxCE] === true) || dE) { 
          envios.push(`<div class="mb-1.5 text-[12px] flex items-center gap-1.5"><span class="text-base">${etp.emj}</span> <span class="text-indigo-600 dark:text-indigo-400 font-bold bg-indigo-50 dark:bg-indigo-900/30 px-1.5 py-0.5 rounded border border-indigo-100 dark:border-indigo-800/50">E-mail</span> <span class="text-slate-500">${dE || '✔️'}</span></div>`); 
          if (responsavel && responsavel !== "Sistema") resps.add(`<div class="mb-1.5 text-[12px] flex items-center gap-1.5"><span class="text-base">${etp.emj}</span> <span class="font-black text-slate-700 dark:text-slate-300">${responsavel}</span></div>`);
          hasEnvio = true; 
        }
        if ((idxCW > -1 && dados[i][idxCW] === true) || dW) { 
          envios.push(`<div class="mb-1.5 text-[12px] flex items-center gap-1.5"><span class="text-base">${etp.emj}</span> <span class="text-emerald-600 dark:text-emerald-400 font-bold bg-emerald-50 dark:bg-emerald-900/30 px-1.5 py-0.5 rounded border border-emerald-100 dark:border-emerald-800/50">Whats</span> <span class="text-slate-500">${dW || '✔️'}</span></div>`);
          if (responsavel && responsavel !== "Sistema") resps.add(`<div class="mb-1.5 text-[12px] flex items-center gap-1.5"><span class="text-base">${etp.emj}</span> <span class="font-black text-slate-700 dark:text-slate-300">${responsavel}</span></div>`); 
          hasEnvio = true;
        }
        if (hasEnvio) highestStage = index + 1;
      });

      let status = `<span class="bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400 px-2.5 py-1 rounded-md text-[10px] font-black border border-slate-200 dark:border-slate-700 tracking-wider shadow-sm">⏳ PENDENTE</span>`;
      if (highestStage === 1) status = `<span class="bg-indigo-100 dark:bg-indigo-900/40 text-indigo-800 dark:text-indigo-400 px-2.5 py-1 rounded-md text-[10px] font-black border border-indigo-300 dark:border-indigo-800/50 tracking-wider shadow-sm">🔰 BOAS VINDAS</span>`;
      else if (highestStage === 2) status = `<span class="bg-amber-100 dark:bg-amber-900/40 text-amber-800 dark:text-amber-400 px-2.5 py-1 rounded-md text-[10px] font-black border border-amber-300 dark:border-amber-800/50 tracking-wider shadow-sm">⚠️ ALERTA 5D</span>`;
      else if (highestStage === 3) status = `<span class="bg-rose-100 dark:bg-rose-900/40 text-rose-800 dark:text-rose-400 px-2.5 py-1 rounded-md text-[10px] font-black border border-rose-300 dark:border-rose-800/50 tracking-wider shadow-sm">⛔ PRAZO EXP</span>`;

      if ((p && sC.has(p)) || (c && sC.has(c))) status = `<span class="bg-emerald-100 dark:bg-emerald-900/40 text-emerald-800 dark:text-emerald-400 px-2.5 py-1 rounded-md text-[10px] font-black border border-emerald-300 dark:border-emerald-800/50 tracking-wider shadow-sm">✅ CONCLUÍDO</span>`;
      else if ((p && sI.has(p)) || (c && sI.has(c))) status = `<span class="bg-purple-100 dark:bg-purple-900/40 text-purple-800 dark:text-purple-400 px-2.5 py-1 rounded-md text-[10px] font-black border border-purple-300 dark:border-purple-800/50 tracking-wider shadow-sm">🟣 INATIVO</span>`;
      else if (e && sE.has(e)) status = `<span class="bg-rose-100 dark:bg-rose-900/40 text-rose-800 dark:text-rose-400 px-2.5 py-1 rounded-md text-[10px] font-black border border-rose-300 dark:border-rose-800/50 tracking-wider shadow-sm">❌ ERRO</span>`;

      linhasMatriz.push([ 
        status, 
        `<div class="font-black text-slate-800 dark:text-white text-sm mb-1">${nome}</div><div class="text-[11px] text-slate-500 font-medium">${e}</div>`, 
        `<div class="font-mono text-slate-600 dark:text-slate-400 font-bold bg-slate-100 dark:bg-slate-800 px-2 py-1 rounded w-max border border-slate-200 dark:border-slate-700">${ident}</div>`, 
        envios.length > 0 ? envios.join("") : `<span class="text-slate-400 dark:text-slate-500 italic text-xs">Aguardando operação...</span>`, 
        resps.size > 0 ? Array.from(resps).join("") : `<span class="text-slate-400 dark:text-slate-500">-</span>` 
      ]);
    }
    return JSON.parse(JSON.stringify({ colunas, linhas: linhasMatriz, isHtml: true }));
  }

  const cabecalho = dados[0].map(c => String(c)), idxRegistroUnico = cabecalho.indexOf("Registro único"); 
  if (idxRegistroUnico > -1) cabecalho.splice(idxRegistroUnico, 1);
  const linhasMatriz = []; 
  let inicio = dados.length > 1001 ? dados.length - 1000 : 1;

  for (let i = dados.length - 1; i >= inicio; i--) { 
    let row = dados[i];
    if (idxRegistroUnico > -1) { row = [...row]; row.splice(idxRegistroUnico, 1); } 
    linhasMatriz.push(row.map(val => web_converterBooleano(val)));
  }
  return JSON.parse(JSON.stringify({ colunas: cabecalho, linhas: linhasMatriz, isHtml: false }));
}

// ====================================================================================
// GESTÃO DE TÉCNICOS (CRUD PLANILHA TÉCNICOS LENDO ATÉ COLUNA I)
// ====================================================================================
function web_obterTecnicos() {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_TECNICOS);
    const aba = ss.getSheets()[0];
    const dados = aba.getDataRange().getValues();
    const tecnicos = [];

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0]) {
        tecnicos.push({
          linha: i + 1,
          nome: String(dados[i][0] || "").trim(),
          endereco: String(dados[i][1] || "").trim(),
          numero: String(dados[i][2] || "").trim(),
          bairro: String(dados[i][3] || "").trim(),
          cidade: String(dados[i][4] || "").trim(),
          estado: String(dados[i][5] || "").trim(),
          cep: String(dados[i][6] || "").trim(),
          telefone: String(dados[i][7] || "").trim(),
          tipo: String(dados[i][8] || "Volante").trim() // Lê a coluna I (Índice 8)
        });
      }
    }
    return tecnicos;
  } catch (e) { return { erro: "Erro ao ler a planilha de técnicos: " + e.message }; }
}

function web_adicionarTecnico(dadosObj) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_TECNICOS);
    const aba = ss.getSheets()[0];
    // Grava as 9 colunas
    aba.appendRow([ dadosObj.nome, dadosObj.endereco, dadosObj.numero, dadosObj.bairro, dadosObj.cidade, dadosObj.estado, dadosObj.cep, dadosObj.telefone, dadosObj.tipo || 'Volante' ]);
    return "✅ Técnico cadastrado com sucesso!";
  } catch (e) { return "❌ Erro ao salvar técnico: " + e.message; }
}

function web_atualizarTecnico(linha, dadosObj) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_TECNICOS);
    const aba = ss.getSheets()[0];
    // Atualiza as 9 colunas
    aba.getRange(linha, 1, 1, 9).setValues([[ dadosObj.nome, dadosObj.endereco, dadosObj.numero, dadosObj.bairro, dadosObj.cidade, dadosObj.estado, dadosObj.cep, dadosObj.telefone, dadosObj.tipo || 'Volante' ]]);
    return "✅ Dados do técnico atualizados com sucesso!";
  } catch (e) { return "❌ Erro ao atualizar técnico: " + e.message; }
}

function web_removerTecnico(linha) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA_TECNICOS);
    const aba = ss.getSheets()[0];
    aba.deleteRow(linha);
    return "✅ Técnico removido com sucesso!";
  } catch (e) { return "❌ Erro ao remover técnico: " + e.message; }
}

// ====================================================================================
// HEALTH CHECK (TESTE DE APIS)
// ====================================================================================
function web_testarAPIs() {
  const status = { hinova: false, tempo: 0, erro: null };
  const inicio = new Date().getTime();
  
  try {
    const options = {
      "method": "post",
      "headers": { 
        "Authorization": "Bearer " + SGA_CONFIG.TOKEN_ASSOCIACAO, 
        "Content-Type": "application/json" 
      },
      "payload": JSON.stringify({ 
        "usuario": SGA_CONFIG.USUARIO, 
        "senha": SGA_CONFIG.SENHA 
      }),
      "muteHttpExceptions": true
    };
    
    const resp = UrlFetchApp.fetch(SGA_CONFIG.URL_AUTH, options);
    status.tempo = new Date().getTime() - inicio;

    if (resp.getResponseCode() === 200) {
      const json = JSON.parse(resp.getContentText());
      if (json.token_usuario) {
        status.hinova = true;
      } else {
        status.erro = "A API Hinova respondeu, mas o token de autenticação veio vazio.";
      }
    } else {
      status.erro = `Falha de Comunicação (HTTP ${resp.getResponseCode()})`;
    }
  } catch (e) {
    status.erro = "Erro interno do servidor Google: " + e.message;
  }
  
  return status;
}
