// ====================================================================================
// 🧠 ARQUIVO: 09_Auth.js (Gestão de Identidade e Sessões)
// ====================================================================================

function web_validarLoginInterno(login, senha) {
  try {
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    const aba = ss.getSheetByName("🔐 Usuários");
    if (!aba) return { erro: "Aba '🔐 Usuários' não configurada no banco." };

    const dados = aba.getDataRange().getValues();
    
    // Varre a planilha procurando o match perfeito de login e senha
    for (let i = 1; i < dados.length; i++) {
      const loginBanco = String(dados[i][1] || "").trim();
      const senhaBanco = String(dados[i][2] || "").trim();
      
      if (loginBanco === login && senhaBanco === senha) {
        return {
          sucesso: true,
          nome: String(dados[i][0]).trim(),
          nivel: String(dados[i][5] || "Operador").trim().toUpperCase()
        };
      }
    }
    
    return { erro: "Login ou senha incorretos." };
  } catch (e) {
    return { erro: "Erro de servidor: " + e.message };
  }
}