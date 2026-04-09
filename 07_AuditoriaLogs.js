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
      cr = "#2E7D32";
      ps = "bold"; 
    } else if ((p && sI.has(p)) || (c && sI.has(c))) { 
      cr = "#9C27B0";
      ps = "bold";
    } else if (e && sE.has(e)) { 
      cr = "#FF0000";
      ps = "bold"; 
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