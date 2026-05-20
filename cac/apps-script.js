/**
 * ══════════════════════════════════════════════════════════════════════════
 *  IBNET TELECOM — Google Apps Script
 *  Sincroniza instalações do painel Operações → planilha Controle de Vendas
 * ══════════════════════════════════════════════════════════════════════════
 *
 *  COMO IMPLANTAR:
 *  1. Acesse: https://script.google.com
 *  2. Crie um novo projeto ("Novo projeto")
 *  3. Cole TODO o conteúdo deste arquivo na área de código
 *  4. Clique em "Implantar" → "Nova implantação"
 *  5. Tipo: "App da Web"
 *  6. Executar como: "Eu (seu email)"
 *  7. Quem tem acesso: "Qualquer pessoa" (Anyone)
 *  8. Clique "Implantar" e copie a URL gerada
 *  9. Cole a URL na constante APPS_SCRIPT_URL em cac/ativacao.html
 *
 *  ⚠️ Não compartilhe a URL publicamente — quem tiver ela pode gravar na planilha.
 * ══════════════════════════════════════════════════════════════════════════
 */

// ID da planilha Controle de Vendas (IBNET)
const PLANILHA_ID = '1Tw_1VOAC3lzm_cAIcPx9q8ekoJU4r7ZzJG9Qavov_Xo';

// Mapeamento: mês ISO (YYYY-MM) → nome da aba na planilha
const MES_PARA_ABA = {
  '2026-01': 'JAN2026',
  '2026-02': 'FEV2026',
  '2026-03': 'MAR2026',
  '2026-04': 'ABRI2026',
  '2026-05': 'MAIO2026',
  '2026-06': 'JUN2026',
  '2026-07': 'JUL2026',
  '2026-08': 'AGO2026',
  '2026-09': 'SET2026',
  '2026-10': 'OUT2026',
  '2026-11': 'NOV2026',
  '2026-12': 'DEZ2026',
};

// ── ENDPOINT PRINCIPAL ────────────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const { cliente, pppoe, contrato, sn, tipo, data, tecnico } = payload;

    if (!cliente) {
      return resposta({ ok: false, erro: 'Nome do cliente não informado' });
    }

    const ss = SpreadsheetApp.openById(PLANILHA_ID);

    // Determinar qual aba usar pelo mês do campo 'data' (YYYY-MM-DD)
    const mesChave = data ? data.slice(0, 7) : '';
    const nomeAba  = MES_PARA_ABA[mesChave] || abaDoMesAtual(ss);

    let sheet = ss.getSheetByName(nomeAba);
    if (!sheet) {
      // Fallback: última aba que não seja utilitária
      const ignorar = ['Base_Dados', 'INADIMPLENTES', 'Log', 'DASHBOARD'];
      sheet = ss.getSheets().reverse().find(s => !ignorar.includes(s.getName()));
    }
    if (!sheet) {
      return resposta({ ok: false, erro: 'Aba não encontrada: ' + nomeAba });
    }

    // Ler todos os dados da aba para localizar colunas pelo cabeçalho
    const dados       = sheet.getDataRange().getValues();
    const cabecalhoLn = encontrarCabecalho(dados);
    const cab         = dados[cabecalhoLn] || [];

    // Índices das colunas relevantes (busca pelo nome do cabeçalho)
    const colNome = encontrarColuna(cab, ['nome do cliente', 'nome', 'cliente']);
    const colSN   = encontrarColuna(cab, ['sn do equipamento', 'serial', 'sn', 'equipamento sn']);
    const colMes  = encontrarColuna(cab, ['mês', 'mes', 'month']);
    const colPPP  = encontrarColuna(cab, ['pppoe', 'login', 'usuário']);
    const colTec  = encontrarColuna(cab, ['técnico', 'tecnico', 'instalador']);

    // Usar posições padrão se cabeçalho não encontrado
    const iNome = colNome >= 0 ? colNome : 0;   // col A
    const iSN   = colSN   >= 0 ? colSN   : 18;  // col S (índice 18)
    const iMes  = colMes  >= 0 ? colMes  : -1;
    const iPPP  = colPPP  >= 0 ? colPPP  : -1;
    const iTec  = colTec  >= 0 ? colTec  : -1;

    // Procurar linha do cliente (pela coluna nome, após o cabeçalho)
    const clienteNorm = normalizar(cliente);
    let linhaCliente  = -1;

    for (let r = cabecalhoLn + 1; r < dados.length; r++) {
      if (normalizar(String(dados[r][iNome] || '')) === clienteNorm) {
        linhaCliente = r;
        break;
      }
    }

    if (linhaCliente >= 0) {
      // ── Cliente encontrado: atualizar S/N e dados do técnico ──────────
      const linha1Based = linhaCliente + 1; // Sheets é 1-based
      if (sn) sheet.getRange(linha1Based, iSN + 1).setValue(sn);
      if (iMes >= 0 && data) {
        const [ano, mes, dia] = data.split('-');
        sheet.getRange(linha1Based, iMes + 1).setValue(`${dia}/${mes}/${ano}`);
      }
      if (iTec >= 0 && tecnico) sheet.getRange(linha1Based, iTec + 1).setValue(tecnico);

      return resposta({
        ok: true,
        aba: sheet.getName(),
        acao: 'atualizado',
        linha: linha1Based,
        clienteEncontrado: true,
      });

    } else {
      // ── Cliente não encontrado: adicionar nova linha ───────────────────
      const totalCols = Math.max(iSN + 1, iNome + 1, 19); // mínimo col S
      const novaLinha = new Array(totalCols).fill('');

      novaLinha[iNome] = cliente;
      if (sn)      novaLinha[iSN]   = sn;
      if (pppoe && iPPP >= 0)  novaLinha[iPPP]  = pppoe;
      if (tecnico && iTec >= 0) novaLinha[iTec] = tecnico;
      if (data && iMes >= 0) {
        const [ano, mes, dia] = data.split('-');
        novaLinha[iMes] = `${dia}/${mes}/${ano}`;
      }

      sheet.appendRow(novaLinha);

      return resposta({
        ok: true,
        aba: sheet.getName(),
        acao: 'adicionado',
        clienteEncontrado: false,
      });
    }

  } catch (err) {
    return resposta({ ok: false, erro: err.toString() });
  }
}

// ── HELPERS ───────────────────────────────────────────────────────────────

function resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Normaliza string para comparação: remove acentos, maiúscula, espaços extras
function normalizar(str) {
  return String(str || '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ');
}

// Encontra a linha de cabeçalho (primeira linha com "Nome" ou similar)
function encontrarCabecalho(dados) {
  for (let r = 0; r < Math.min(5, dados.length); r++) {
    const linha = dados[r].map(c => normalizar(String(c)));
    if (linha.some(c => c.includes('NOME') || c.includes('CLIENTE') || c.includes('SN'))) {
      return r;
    }
  }
  return 0; // padrão: primeira linha
}

// Encontra o índice de uma coluna dado uma lista de nomes possíveis
function encontrarColuna(cabecalho, nomesPossiveis) {
  const normCab = cabecalho.map(c => normalizar(String(c)));
  for (const nome of nomesPossiveis) {
    const idx = normCab.findIndex(c => c.includes(normalizar(nome)));
    if (idx >= 0) return idx;
  }
  return -1;
}

// Descobre o nome da aba do mês atual
function abaDoMesAtual(ss) {
  const agora  = new Date();
  const ano    = agora.getFullYear();
  const mesIdx = agora.getMonth(); // 0-based
  const chave  = `${ano}-${String(mesIdx + 1).padStart(2, '0')}`;
  return MES_PARA_ABA[chave] || '';
}

// ── TESTE (execute manualmente no Apps Script para validar) ───────────────
function testar() {
  const e = {
    postData: {
      contents: JSON.stringify({
        cliente:  'TESTE IBNET',
        pppoe:    'teste@ibnet',
        contrato: '99999',
        sn:       'ABC123456789',
        tipo:     'instalacao',
        data:     new Date().toISOString().slice(0, 10),
        tecnico:  'Carlos',
      })
    }
  };
  Logger.log(doPost(e).getContent());
}
