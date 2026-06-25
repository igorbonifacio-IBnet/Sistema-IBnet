/**
 * ════════════════════════════════════════════════════════════════════════════
 *  IBNET TELECOM — Automação Google Sheets (v4)
 *  Funções:
 *    1. Cálculo automático de "Dias em atraso"
 *    2. Alerta diário de inadimplência por e-mail
 *    3. Relatório semanal de KPIs por e-mail
 *    4. Score de risco de churn por cliente
 *    5. Log automático de todas as alterações
 *    6. Backup semanal automático (sexta-feira)
 *    7. Endpoint doPost — sincroniza instalações do painel Operações
 *    8. [NOVO] Proxy doGet — busca fotos de OS do ERP SGP
 * ════════════════════════════════════════════════════════════════════════════
 *
 *  ATUALIZAÇÃO (substitua o script antigo por este):
 *  1. Extensões → Apps Script → seleciona tudo → cola este arquivo
 *  2. Salva (Ctrl+S)
 *  3. Seleciona "configurarGatilhos" e clica ▶ (uma única vez)
 *  4. Autorize as permissões
 *
 *  PARA ATIVAR O ENDPOINT DE INSTALAÇÕES (função 7):
 *  1. Clique em "Implantar" → "Nova implantação"  (ou "Gerenciar implantações")
 *  2. Tipo: "App da Web"
 *  3. Executar como: "Eu (seu email)"
 *  4. Quem tem acesso: "Qualquer pessoa" (Anyone)
 *  5. Clique "Implantar" e copie a URL gerada
 *  6. Cole a URL na constante APPS_SCRIPT_URL em cac/ativacao.html
 *
 *  ⚠️  Não compartilhe a URL publicamente — quem tiver ela pode gravar na planilha.
 * ════════════════════════════════════════════════════════════════════════════
 */

// ── CONFIGURAÇÃO GERAL ────────────────────────────────────────────────────

const ABAS_MONITORADAS = ['JAN2026', 'FEV2026', 'MAR2026', 'ABRI2026', 'MAIO2026', 'JUN2026'];
const STATUS_IGNORAR   = ['Pago', 'Isento', 'Cancelado'];
const DIAS_PARA_ATRASO = 3;    // dias em atraso para mudar status de Pendente → Em atraso
const DIAS_MIN_ALERTA  = 5;    // só alerta clientes com mais de X dias em atraso

// E-mails que receberão alertas e relatórios
const EMAILS_GESTAO = [
  'igorbonifacio23@gmail.com',
  'matheushenrique161013@gmail.com'
];

// Nome da coluna de score (será criada automaticamente se não existir)
const COL_SCORE = 'Score Churn';

// ID da planilha Controle de Vendas (usada pelo endpoint de instalações)
const PLANILHA_VENDAS_ID = '1Tw_1VOAC3lzm_cAIcPx9q8ekoJU4r7ZzJG9Qavov_Xo';

// Mapeamento mês ISO (YYYY-MM) → nome da aba na planilha Controle de Vendas
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


// ════════════════════════════════════════════════════════════════════════════
//  1. CÁLCULO AUTOMÁTICO DE DIAS EM ATRASO
// ════════════════════════════════════════════════════════════════════════════

function calcularDiasEmAtraso() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  let totalAtualizados = 0;

  ABAS_MONITORADAS.forEach(nomeAba => {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) return;

    const dados    = aba.getDataRange().getValues();
    const cab      = dados[0];
    const iVenc    = cab.findIndex(h => normaliza(h) === 'data de vencimento');
    const iStatus  = cab.findIndex(h => normaliza(h) === 'status do pagamento');
    const iDias    = cab.findIndex(h => normaliza(h) === 'dias em atraso');

    if (iVenc === -1 || iStatus === -1 || iDias === -1) {
      Logger.log(`⚠️  Aba "${nomeAba}": colunas não encontradas.`);
      return;
    }

    for (let i = 1; i < dados.length; i++) {
      const linha  = dados[i];
      const status = (linha[iStatus] || '').toString().trim();
      const vencRaw = linha[iVenc];
      if (!vencRaw) continue;

      if (STATUS_IGNORAR.includes(status)) {
        if (linha[iDias] !== 0 && linha[iDias] !== '') {
          aba.getRange(i+1, iDias+1).setValue(0);
          totalAtualizados++;
        }
        continue;
      }

      const dataVenc = parsearData(vencRaw);
      if (!dataVenc) continue;

      const dias = Math.max(0, Math.floor((hoje - dataVenc) / 86400000));

      if (linha[iDias] !== dias) {
        aba.getRange(i+1, iDias+1).setValue(dias);
        totalAtualizados++;
      }

      if (dias >= DIAS_PARA_ATRASO && status === 'Pendente') {
        aba.getRange(i+1, iStatus+1).setValue('Em atraso');
        totalAtualizados++;
      }
    }
  });

  Logger.log(`✅ Dias em atraso: ${totalAtualizados} célula(s) atualizada(s).`);
  return totalAtualizados;
}


// ════════════════════════════════════════════════════════════════════════════
//  2. ALERTA DIÁRIO DE INADIMPLÊNCIA POR E-MAIL
// ════════════════════════════════════════════════════════════════════════════

function enviarAlertaInadimplencia() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const hoje      = new Date(); hoje.setHours(0,0,0,0);
  const inadimplentes = [];

  ABAS_MONITORADAS.forEach(nomeAba => {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) return;

    const dados   = aba.getDataRange().getValues();
    const cab     = dados[0];
    const iNome   = cab.findIndex(h => normaliza(h) === 'nome do cliente');
    const iStatus = cab.findIndex(h => normaliza(h) === 'status do pagamento');
    const iDias   = cab.findIndex(h => normaliza(h) === 'dias em atraso');
    const iPlano  = cab.findIndex(h => normaliza(h) === 'plano contratado');
    const iValor  = cab.findIndex(h => normaliza(h).includes('valor mensal'));
    const iPOP    = cab.findIndex(h => normaliza(h) === 'pop');

    if (iNome === -1 || iStatus === -1) return;

    for (let i = 1; i < dados.length; i++) {
      const linha  = dados[i];
      const status = (linha[iStatus] || '').toString().trim();
      const dias   = parseInt(linha[iDias]) || 0;
      const nome   = (linha[iNome] || '').toString().trim();

      if (!nome) continue;
      if (status !== 'Em atraso' && status !== 'Pendente') continue;
      if (dias < DIAS_MIN_ALERTA) continue;

      inadimplentes.push({
        mes:    nomeAba,
        nome:   nome,
        status: status,
        dias:   dias,
        plano:  iPlano  >= 0 ? (linha[iPlano]  || '—') : '—',
        valor:  iValor  >= 0 ? (linha[iValor]  || '—') : '—',
        pop:    iPOP    >= 0 ? (linha[iPOP]    || '—') : '—',
      });
    }
  });

  if (inadimplentes.length === 0) {
    Logger.log('✅ Nenhum inadimplente acima do limite — e-mail não enviado.');
    return;
  }

  // Ordena por dias em atraso (maior primeiro)
  inadimplentes.sort((a, b) => b.dias - a.dias);

  const dataHoje = hoje.toLocaleDateString('pt-BR');
  const assunto  = `⚠️ IBnet — ${inadimplentes.length} cliente(s) em atraso · ${dataHoje}`;

  // Monta tabela HTML
  const linhasTabela = inadimplentes.map(c => `
    <tr>
      <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9">${c.nome}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;color:#64748b">${c.mes}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9">${c.plano}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9">${c.pop}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9">${c.valor}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;
          color:${c.dias >= 15 ? '#dc2626' : c.dias >= 7 ? '#d97706' : '#64748b'};
          font-weight:700">${c.dias} dias</td>
      <td style="padding:8px 12px;border-bottom:1px solid #f1f5f9">
        <span style="background:${c.status==='Em atraso'?'#fee2e2':'#fef9c3'};
          color:${c.status==='Em atraso'?'#b91c1c':'#854d0e'};
          padding:2px 8px;border-radius:6px;font-size:12px;font-weight:600">${c.status}</span>
      </td>
    </tr>`).join('');

  const html = `
  <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:800px;margin:0 auto;background:#f8fafc;padding:24px">
    <div style="background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e2e8f0">

      <!-- Header -->
      <div style="background:linear-gradient(135deg,#CC2200,#FF5500);padding:24px 28px">
        <h1 style="margin:0;color:#fff;font-size:20px;font-weight:700">⚠️ Alerta de Inadimplência</h1>
        <p style="margin:6px 0 0;color:rgba(255,255,255,.85);font-size:13px">${dataHoje} · IBnet Telecom</p>
      </div>

      <!-- Resumo -->
      <div style="padding:20px 28px;background:#fff7f5;border-bottom:1px solid #ffe4de">
        <p style="margin:0;font-size:14px;color:#7c2d12">
          <strong>${inadimplentes.length} cliente(s)</strong> com mais de ${DIAS_MIN_ALERTA} dias em atraso requerem atenção.
        </p>
      </div>

      <!-- Tabela -->
      <div style="padding:20px 28px">
        <table style="width:100%;border-collapse:collapse;font-size:13px">
          <thead>
            <tr style="background:#f8fafc">
              <th style="padding:10px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#64748b;border-bottom:2px solid #e2e8f0">Cliente</th>
              <th style="padding:10px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#64748b;border-bottom:2px solid #e2e8f0">Mês</th>
              <th style="padding:10px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#64748b;border-bottom:2px solid #e2e8f0">Plano</th>
              <th style="padding:10px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#64748b;border-bottom:2px solid #e2e8f0">POP</th>
              <th style="padding:10px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#64748b;border-bottom:2px solid #e2e8f0">Valor</th>
              <th style="padding:10px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#64748b;border-bottom:2px solid #e2e8f0">Atraso</th>
              <th style="padding:10px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#64748b;border-bottom:2px solid #e2e8f0">Status</th>
            </tr>
          </thead>
          <tbody>${linhasTabela}</tbody>
        </table>
      </div>

      <!-- Footer -->
      <div style="padding:16px 28px;background:#f8fafc;border-top:1px solid #e2e8f0">
        <p style="margin:0;font-size:12px;color:#94a3b8">Enviado automaticamente pelo Sistema IBnet · ${dataHoje}</p>
      </div>
    </div>
  </div>`;

  EMAILS_GESTAO.forEach(email => {
    MailApp.sendEmail({ to: email, subject: assunto, htmlBody: html });
  });

  Logger.log(`✅ Alerta enviado para ${EMAILS_GESTAO.join(', ')} · ${inadimplentes.length} inadimplentes`);
}


// ════════════════════════════════════════════════════════════════════════════
//  3. RELATÓRIO SEMANAL DE KPIs POR E-MAIL
// ════════════════════════════════════════════════════════════════════════════

function enviarRelatorioSemanal() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let total = 0, pagos = 0, pendentes = 0, emAtraso = 0,
      cancelados = 0, isentos = 0, mrr = 0, churns = 0;

  ABAS_MONITORADAS.forEach(nomeAba => {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) return;

    const dados   = aba.getDataRange().getValues();
    const cab     = dados[0];
    const iNome   = cab.findIndex(h => normaliza(h) === 'nome do cliente');
    const iStatus = cab.findIndex(h => normaliza(h) === 'status do pagamento');
    const iValor  = cab.findIndex(h => normaliza(h).includes('valor mensal'));
    const iChurn  = cab.findIndex(h => normaliza(h) === 'churn');

    if (iNome === -1) return;

    for (let i = 1; i < dados.length; i++) {
      const nome = (dados[i][iNome] || '').toString().trim();
      if (!nome) continue;

      const status = iStatus >= 0 ? (dados[i][iStatus] || '').toString().trim() : '';
      const valor  = iValor  >= 0 ? parsearValor(dados[i][iValor]) : 0;
      const churn  = iChurn  >= 0 ? (dados[i][iChurn] || '').toString().toUpperCase().trim() : '';

      total++;
      if (status === 'Pago')       { pagos++;      mrr += valor; }
      if (status === 'Pendente')   { pendentes++;  mrr += valor; }
      if (status === 'Em atraso')  { emAtraso++;   mrr += valor; }
      if (status === 'Cancelado')    cancelados++;
      if (status === 'Isento')     { isentos++;    }
      if (churn === 'SIM')           churns++;
    }
  });

  const ativos      = total - cancelados;
  const taxaPag     = ativos > 0 ? ((pagos / ativos) * 100).toFixed(1) : 0;
  const inadimpl    = ativos > 0 ? (((pendentes + emAtraso) / ativos) * 100).toFixed(1) : 0;
  const taxaChurn   = total  > 0 ? ((cancelados / total) * 100).toFixed(1) : 0;
  const ticketMedio = ativos > 0 ? (mrr / ativos).toFixed(2) : 0;
  const dataHoje    = new Date().toLocaleDateString('pt-BR');

  const assunto = `📊 IBnet — Relatório Semanal · ${dataHoje}`;

  const kpiCard = (emoji, label, value, color) => `
    <div style="background:#f8fafc;border-radius:10px;padding:16px;border-left:4px solid ${color};flex:1;min-width:140px">
      <div style="font-size:22px;margin-bottom:6px">${emoji}</div>
      <div style="font-size:22px;font-weight:700;color:${color};line-height:1">${value}</div>
      <div style="font-size:11px;color:#64748b;margin-top:4px;text-transform:uppercase;letter-spacing:.5px">${label}</div>
    </div>`;

  const html = `
  <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:700px;margin:0 auto;background:#f8fafc;padding:24px">
    <div style="background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e2e8f0">

      <!-- Header -->
      <div style="background:linear-gradient(135deg,#CC2200,#FF5500);padding:24px 28px">
        <h1 style="margin:0;color:#fff;font-size:20px;font-weight:700">📊 Relatório Semanal</h1>
        <p style="margin:6px 0 0;color:rgba(255,255,255,.85);font-size:13px">${dataHoje} · IBnet Telecom</p>
      </div>

      <!-- KPIs principais -->
      <div style="padding:24px 28px">
        <h2 style="margin:0 0 16px;font-size:13px;text-transform:uppercase;letter-spacing:.6px;color:#64748b">Visão Geral</h2>
        <div style="display:flex;gap:12px;flex-wrap:wrap">
          ${kpiCard('👥', 'Total Clientes', total, '#3b82f6')}
          ${kpiCard('💰', 'MRR Total', 'R$ ' + mrr.toLocaleString('pt-BR', {minimumFractionDigits:2}), '#22c55e')}
          ${kpiCard('🎯', 'Ticket Médio', 'R$ ' + parseFloat(ticketMedio).toLocaleString('pt-BR', {minimumFractionDigits:2}), '#6366f1')}
        </div>
      </div>

      <!-- Status -->
      <div style="padding:0 28px 24px">
        <h2 style="margin:0 0 16px;font-size:13px;text-transform:uppercase;letter-spacing:.6px;color:#64748b">Status de Pagamento</h2>
        <div style="display:flex;gap:12px;flex-wrap:wrap">
          ${kpiCard('✅', 'Pagos', pagos, '#22c55e')}
          ${kpiCard('⏳', 'Pendentes', pendentes, '#eab308')}
          ${kpiCard('⚠️', 'Em Atraso', emAtraso, '#ef4444')}
          ${kpiCard('❌', 'Cancelados', cancelados, '#64748b')}
        </div>
      </div>

      <!-- Taxas -->
      <div style="padding:0 28px 24px">
        <h2 style="margin:0 0 16px;font-size:13px;text-transform:uppercase;letter-spacing:.6px;color:#64748b">Indicadores</h2>
        <table style="width:100%;border-collapse:collapse;font-size:13px">
          ${linhaIndicador('Taxa de Pagamento', taxaPag + '%', taxaPag >= 80 ? '#22c55e' : taxaPag >= 60 ? '#eab308' : '#ef4444')}
          ${linhaIndicador('Inadimplência', inadimpl + '%', inadimpl <= 10 ? '#22c55e' : inadimpl <= 25 ? '#eab308' : '#ef4444')}
          ${linhaIndicador('Taxa de Churn', taxaChurn + '%', taxaChurn <= 5 ? '#22c55e' : taxaChurn <= 10 ? '#eab308' : '#ef4444')}
          ${linhaIndicador('Clientes com Churn', churns, churns === 0 ? '#22c55e' : '#ef4444')}
        </table>
      </div>

      <!-- Footer -->
      <div style="padding:16px 28px;background:#f8fafc;border-top:1px solid #e2e8f0">
        <p style="margin:0;font-size:12px;color:#94a3b8">Relatório automático semanal · Sistema IBnet · ${dataHoje}</p>
      </div>
    </div>
  </div>`;

  EMAILS_GESTAO.forEach(email => {
    MailApp.sendEmail({ to: email, subject: assunto, htmlBody: html });
  });

  Logger.log(`✅ Relatório semanal enviado para ${EMAILS_GESTAO.join(', ')}`);
}

function linhaIndicador(label, valor, cor) {
  return `
    <tr>
      <td style="padding:10px 0;border-bottom:1px solid #f1f5f9;color:#374151">${label}</td>
      <td style="padding:10px 0;border-bottom:1px solid #f1f5f9;text-align:right;font-weight:700;color:${cor}">${valor}</td>
    </tr>`;
}


// ════════════════════════════════════════════════════════════════════════════
//  4. SCORE DE RISCO DE CHURN
// ════════════════════════════════════════════════════════════════════════════

/**
 * Calcula um score de 0 a 100 para cada cliente em todas as abas.
 * Quanto maior, maior o risco de cancelar.
 * Cria a coluna "Score Churn" automaticamente se não existir.
 *
 * Fatores:
 *   Status Em atraso     → +40 pts
 *   Status Pendente      → +15 pts
 *   Dias 1–7 em atraso   → +10 pts adicionais
 *   Dias 8–15 em atraso  → +20 pts adicionais
 *   Dias 16–30 em atraso → +30 pts adicionais
 *   Dias > 30 em atraso  → +45 pts adicionais
 *   Churn = SIM          → 100 (já cancelou)
 *   Cancelado            → 100
 */
function calcularScoreChurn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalAtualizados = 0;

  ABAS_MONITORADAS.forEach(nomeAba => {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) return;

    const dados   = aba.getDataRange().getValues();
    const cab     = dados[0];
    const iNome   = cab.findIndex(h => normaliza(h) === 'nome do cliente');
    const iStatus = cab.findIndex(h => normaliza(h) === 'status do pagamento');
    const iDias   = cab.findIndex(h => normaliza(h) === 'dias em atraso');
    const iChurn  = cab.findIndex(h => normaliza(h) === 'churn');

    if (iNome === -1) return;

    // Garante que a coluna Score existe
    let iScore = cab.findIndex(h => normaliza(h) === normaliza(COL_SCORE));
    if (iScore === -1) {
      iScore = cab.length;
      aba.getRange(1, iScore + 1).setValue(COL_SCORE);
    }

    for (let i = 1; i < dados.length; i++) {
      const nome = (dados[i][iNome] || '').toString().trim();
      if (!nome) continue;

      const status = iStatus >= 0 ? (dados[i][iStatus] || '').toString().trim() : '';
      const dias   = iDias   >= 0 ? (parseInt(dados[i][iDias]) || 0) : 0;
      const churn  = iChurn  >= 0 ? (dados[i][iChurn] || '').toString().toUpperCase().trim() : '';

      let score = 0;

      // Já cancelou ou churn confirmado → risco máximo
      if (status === 'Cancelado' || churn === 'SIM') {
        score = 100;
      } else {
        // Fator status
        if (status === 'Em atraso') score += 40;
        if (status === 'Pendente')  score += 15;

        // Fator dias em atraso
        if      (dias > 30) score += 45;
        else if (dias > 15) score += 30;
        else if (dias > 7)  score += 20;
        else if (dias > 0)  score += 10;

        score = Math.min(score, 99); // nunca 100 se não for churn/cancelado
      }

      if (dados[i][iScore] !== score) {
        aba.getRange(i + 1, iScore + 1).setValue(score);
        totalAtualizados++;
      }
    }
  });

  Logger.log(`✅ Score Churn: ${totalAtualizados} célula(s) atualizada(s).`);
  return totalAtualizados;
}


// ════════════════════════════════════════════════════════════════════════════
//  5. LOG AUTOMÁTICO DE ALTERAÇÕES
// ════════════════════════════════════════════════════════════════════════════

/**
 * Registra cada edição na aba "📋 Log".
 * Acionado por gatilho instalável (não é o onEdit simples).
 * Captura: data/hora, usuário, aba, célula, valor anterior, novo valor.
 */
function registrarEdicao(e) {
  if (!e || !e.range) return;

  const ss    = e.source || SpreadsheetApp.getActiveSpreadsheet();
  const range = e.range;
  const nomeAba = range.getSheet().getName();

  // Ignora edições dentro do próprio Log para não criar loop
  if (nomeAba === '📋 Log') return;

  // Cria a aba de Log se ainda não existir
  let logSheet = ss.getSheetByName('📋 Log');
  if (!logSheet) {
    logSheet = ss.insertSheet('📋 Log');
    // Cabeçalho formatado
    const cabecalho = [['Data/Hora', 'Usuário', 'Aba', 'Célula', 'Valor Anterior', 'Novo Valor']];
    logSheet.getRange(1, 1, 1, 6).setValues(cabecalho);
    logSheet.getRange(1, 1, 1, 6)
      .setBackground('#CC2200')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(11);
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidth(1, 160); // Data/Hora
    logSheet.setColumnWidth(2, 220); // Usuário
    logSheet.setColumnWidth(3, 120); // Aba
    logSheet.setColumnWidth(4, 80);  // Célula
    logSheet.setColumnWidth(5, 180); // Valor Anterior
    logSheet.setColumnWidth(6, 180); // Novo Valor
  }

  const agora         = new Date();
  const usuario       = e.user ? e.user.getEmail() : Session.getActiveUser().getEmail() || 'Desconhecido';
  const celula        = range.getA1Notation();
  const valorAnterior = e.oldValue !== undefined ? e.oldValue : '';
  const novoValor     = e.value    !== undefined ? e.value    : range.getValue();

  logSheet.appendRow([agora, usuario, nomeAba, celula, valorAnterior, novoValor]);

  // Cor da linha por tipo de alteração
  const ultimaLinha = logSheet.getLastRow();
  const corLinha = valorAnterior === '' ? '#f0fdf4' :   // novo valor
                   novoValor === ''     ? '#fff7f5' :   // valor apagado
                   '#fffbeb';                            // valor modificado
  logSheet.getRange(ultimaLinha, 1, 1, 6).setBackground(corLinha);

  Logger.log(`📝 Log: ${usuario} editou ${nomeAba}!${celula} — "${valorAnterior}" → "${novoValor}"`);
}


// ════════════════════════════════════════════════════════════════════════════
//  6. BACKUP SEMANAL AUTOMÁTICO
// ════════════════════════════════════════════════════════════════════════════

/**
 * Cria uma cópia de todas as abas monitoradas em uma nova planilha no Drive.
 * Nome do arquivo: "IBnet Backup DD-MM-AAAA"
 * Roda toda sexta-feira às 18h (configurado em configurarGatilhos).
 * Envia e-mail com o link direto para o backup.
 */
function backupSemanal() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const hoje    = new Date();
  const dataStr = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  const nomeBackup = `IBnet Backup ${dataStr}`;

  // Cria nova planilha no Drive
  const backup = SpreadsheetApp.create(nomeBackup);
  let abaCopiadasNomes = [];

  ABAS_MONITORADAS.forEach(nomeAba => {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) return;
    aba.copyTo(backup).setName(nomeAba);
    abaCopiadasNomes.push(nomeAba);
  });

  // Copia também a aba de Log se existir
  const logSheet = ss.getSheetByName('📋 Log');
  if (logSheet) {
    logSheet.copyTo(backup).setName('📋 Log');
    abaCopiadasNomes.push('📋 Log');
  }

  // Remove a aba padrão vazia criada automaticamente pelo Google
  ['Planilha1', 'Sheet1', 'Plan1'].forEach(nome => {
    const aba = backup.getSheetByName(nome);
    if (aba) try { backup.deleteSheet(aba); } catch(err) {}
  });

  const linkBackup = `https://docs.google.com/spreadsheets/d/${backup.getId()}`;
  Logger.log(`✅ Backup criado: "${nomeBackup}" · Link: ${linkBackup}`);

  // Envia e-mail com link do backup
  const dataFormatada = hoje.toLocaleDateString('pt-BR');
  const assunto = `💾 IBnet — Backup semanal realizado · ${dataFormatada}`;
  const html = `
  <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:600px;margin:0 auto;background:#f8fafc;padding:24px">
    <div style="background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e2e8f0">
      <div style="background:linear-gradient(135deg,#CC2200,#FF5500);padding:24px 28px">
        <h1 style="margin:0;color:#fff;font-size:20px;font-weight:700">💾 Backup Semanal Concluído</h1>
        <p style="margin:6px 0 0;color:rgba(255,255,255,.85);font-size:13px">${dataFormatada} · IBnet Telecom</p>
      </div>
      <div style="padding:24px 28px">
        <p style="margin:0 0 16px;color:#374151">O backup semanal foi criado com sucesso no Google Drive.</p>
        <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:20px">
          <tr style="background:#f8fafc">
            <td style="padding:10px 14px;border-bottom:1px solid #e2e8f0;color:#64748b;font-weight:600">Arquivo</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e2e8f0;color:#1e293b">${nomeBackup}</td>
          </tr>
          <tr>
            <td style="padding:10px 14px;border-bottom:1px solid #e2e8f0;color:#64748b;font-weight:600">Abas incluídas</td>
            <td style="padding:10px 14px;border-bottom:1px solid #e2e8f0;color:#1e293b">${abaCopiadasNomes.join(', ')}</td>
          </tr>
          <tr style="background:#f8fafc">
            <td style="padding:10px 14px;color:#64748b;font-weight:600">Data</td>
            <td style="padding:10px 14px;color:#1e293b">${dataFormatada}</td>
          </tr>
        </table>
        <a href="${linkBackup}"
           style="display:inline-block;background:#CC2200;color:#fff;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px">
          📂 Abrir Backup no Drive
        </a>
      </div>
      <div style="padding:16px 28px;background:#f8fafc;border-top:1px solid #e2e8f0">
        <p style="margin:0;font-size:12px;color:#94a3b8">Backup automático semanal · Sistema IBnet · ${dataFormatada}</p>
      </div>
    </div>
  </div>`;

  EMAILS_GESTAO.forEach(email => {
    MailApp.sendEmail({ to: email, subject: assunto, htmlBody: html });
  });

  return linkBackup;
}


// ════════════════════════════════════════════════════════════════════════════
//  7. ENDPOINT WEB — SINCRONIZA INSTALAÇÕES DO PAINEL OPERAÇÕES
// ════════════════════════════════════════════════════════════════════════════

/**
 * Recebe um POST do painel Operações (cac/ativacao.html) e grava
 * os dados do cliente + S/N na planilha "Controle de Vendas".
 *
 * Payload esperado (JSON):
 *   { cliente, pppoe, contrato, sn, tipo, data, tecnico }
 *
 * Para ativar:
 *   Implantar → Nova implantação → App da Web → Qualquer pessoa → Implantar
 *   Copiar a URL gerada e colar em APPS_SCRIPT_URL (cac/ativacao.html)
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const { cliente, pppoe, contrato, sn, tipo, data, tecnico } = payload;

    if (!cliente) {
      return _resposta({ ok: false, erro: 'Nome do cliente não informado' });
    }

    const ss = SpreadsheetApp.openById(PLANILHA_VENDAS_ID);

    // Determinar qual aba usar pelo mês do campo 'data' (YYYY-MM-DD)
    const mesChave = data ? data.slice(0, 7) : '';
    const nomeAba  = MES_PARA_ABA[mesChave] || _abaDoMesAtual();

    let sheet = ss.getSheetByName(nomeAba);
    if (!sheet) {
      // Fallback: última aba que não seja utilitária
      const ignorar = ['Base_Dados', 'INADIMPLENTES', 'Log', 'DASHBOARD'];
      sheet = ss.getSheets().reverse().find(s => !ignorar.includes(s.getName()));
    }
    if (!sheet) {
      return _resposta({ ok: false, erro: 'Aba não encontrada: ' + nomeAba });
    }

    // Ler todos os dados da aba para localizar colunas pelo cabeçalho
    const dados       = sheet.getDataRange().getValues();
    const cabecalhoLn = _encontrarCabecalho(dados);
    const cab         = dados[cabecalhoLn] || [];

    // Índices das colunas relevantes (busca pelo nome do cabeçalho)
    const colNome = _encontrarColuna(cab, ['nome do cliente', 'nome', 'cliente']);
    const colSN   = _encontrarColuna(cab, ['sn do equipamento', 'serial', 'sn', 'equipamento sn']);
    const colMes  = _encontrarColuna(cab, ['mes', 'month']);
    const colPPP  = _encontrarColuna(cab, ['pppoe', 'login', 'usuario']);
    const colTec  = _encontrarColuna(cab, ['tecnico', 'instalador']);

    // Usar posições padrão se cabeçalho não encontrado
    const iNome = colNome >= 0 ? colNome : 0;   // col A
    const iSN   = colSN   >= 0 ? colSN   : 18;  // col S (índice 18)
    const iMes  = colMes  >= 0 ? colMes  : -1;
    const iPPP  = colPPP  >= 0 ? colPPP  : -1;
    const iTec  = colTec  >= 0 ? colTec  : -1;

    // Procurar linha do cliente (pela coluna nome, após o cabeçalho)
    const clienteNorm = normaliza(cliente);
    let linhaCliente  = -1;

    for (let r = cabecalhoLn + 1; r < dados.length; r++) {
      if (normaliza(String(dados[r][iNome] || '')) === clienteNorm) {
        linhaCliente = r;
        break;
      }
    }

    if (linhaCliente >= 0) {
      // ── Cliente encontrado: atualizar S/N e dados do técnico ─────────
      const linha1Based = linhaCliente + 1; // Sheets é 1-based
      if (sn) sheet.getRange(linha1Based, iSN + 1).setValue(sn);
      if (iMes >= 0 && data) {
        const [ano, mes, dia] = data.split('-');
        sheet.getRange(linha1Based, iMes + 1).setValue(`${dia}/${mes}/${ano}`);
      }
      if (iTec >= 0 && tecnico) sheet.getRange(linha1Based, iTec + 1).setValue(tecnico);

      return _resposta({
        ok: true,
        aba: sheet.getName(),
        acao: 'atualizado',
        linha: linha1Based,
        clienteEncontrado: true,
      });

    } else {
      // ── Cliente não encontrado: adicionar nova linha ──────────────────
      const totalCols = Math.max(iSN + 1, iNome + 1, 19); // mínimo col S
      const novaLinha = new Array(totalCols).fill('');

      novaLinha[iNome] = cliente;
      if (sn)                        novaLinha[iSN]  = sn;
      if (pppoe  && iPPP >= 0)       novaLinha[iPPP] = pppoe;
      if (tecnico && iTec >= 0)      novaLinha[iTec] = tecnico;
      if (data   && iMes >= 0) {
        const [ano, mes, dia] = data.split('-');
        novaLinha[iMes] = `${dia}/${mes}/${ano}`;
      }

      sheet.appendRow(novaLinha);

      return _resposta({
        ok: true,
        aba: sheet.getName(),
        acao: 'adicionado',
        clienteEncontrado: false,
      });
    }

  } catch (err) {
    return _resposta({ ok: false, erro: err.toString() });
  }
}

// ── Helpers internos do endpoint (prefixo _ para não conflitar) ──────────

function _resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Encontra a linha de cabeçalho (primeira linha com "Nome" ou similar)
function _encontrarCabecalho(dados) {
  for (let r = 0; r < Math.min(5, dados.length); r++) {
    const linha = dados[r].map(c => normaliza(String(c)));
    if (linha.some(c => c.includes('nome') || c.includes('cliente') || c.includes('sn'))) {
      return r;
    }
  }
  return 0;
}

// Encontra o índice de uma coluna dado uma lista de nomes possíveis
function _encontrarColuna(cabecalho, nomesPossiveis) {
  const normCab = cabecalho.map(c => normaliza(String(c)));
  for (const nome of nomesPossiveis) {
    const idx = normCab.findIndex(c => c.includes(normaliza(nome)));
    if (idx >= 0) return idx;
  }
  return -1;
}

// Descobre o nome da aba do mês atual na Controle de Vendas
function _abaDoMesAtual() {
  const agora  = new Date();
  const ano    = agora.getFullYear();
  const mesIdx = agora.getMonth(); // 0-based
  const chave  = `${ano}-${String(mesIdx + 1).padStart(2, '0')}`;
  return MES_PARA_ABA[chave] || '';
}

// ── Teste manual do endpoint (execute no Apps Script para validar) ────────
function testarEndpoint() {
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


// ════════════════════════════════════════════════════════════════════════════
//  8. PROXY SGP — BUSCA FOTOS DO ERP (doGet)
// ════════════════════════════════════════════════════════════════════════════

/**
 * Endpoint GET: o dashboard chama este proxy para buscar as fotos de uma OS
 * direto do ERP SGP (evita CORS — o Apps Script faz a requisição server-side).
 *
 * URL: <APPS_SCRIPT_URL>?action=sgp_fotos&ocorrencia_id=XXXX
 *
 * Retorna JSON:
 *   { ok: true, fotos: [{id, data_hora, por, content_type, base64}], total: N }
 *   { ok: false, erro: "..." }
 *
 * ⚠️  Trocar SGP_USER/SGP_PASS quando o SGP criar o usuário dedicado.
 */

const SGP_BASE = 'https://sbsginfo.sgp.tsmx.com.br';

// ── CREDENCIAIS (NÃO ficam no código) ──────────────────────────────────────
// Este arquivo é público (GitHub Pages), então segredos NUNCA podem estar aqui.
// Configure uma única vez em:
//   Apps Script → ⚙ Configurações do projeto → Propriedades do script → Adicionar:
//     SGP_USER       → usuário do SGP
//     SGP_PASS       → senha do SGP
//     ANTHROPIC_KEY  → chave da API Anthropic (sk-ant-...)
// Os valores são lidos com segurança em tempo de execução, sem aparecer no fonte.
const _SCRIPT_PROPS = PropertiesService.getScriptProperties();
const SGP_USER      = _SCRIPT_PROPS.getProperty('SGP_USER')      || '';
const SGP_PASS      = _SCRIPT_PROPS.getProperty('SGP_PASS')      || '';
const ANTHROPIC_KEY = _SCRIPT_PROPS.getProperty('ANTHROPIC_KEY') || '';

/**
 * Verifica (sem expor valores) quais credenciais já estão configuradas.
 * Rode esta função no editor do Apps Script após cadastrar as propriedades.
 */
function verificarCredenciais() {
  ['SGP_USER', 'SGP_PASS', 'ANTHROPIC_KEY'].forEach(nome => {
    const v = _SCRIPT_PROPS.getProperty(nome);
    Logger.log(`${nome}: ${v ? '✅ configurado (' + v.length + ' chars)' : '❌ FALTANDO'}`);
  });
}

/**
 * Testa o login no SGP com as credenciais atuais (SGP_USER/SGP_PASS).
 * NÃO expõe a senha — apenas informa se a sessão foi aberta com sucesso.
 * Rode no editor do Apps Script e veja o resultado em "Execução / Logs".
 */
function sgpTestarLogin() {
  if (!SGP_USER || !SGP_PASS) {
    Logger.log('❌ Configure SGP_USER e SGP_PASS em Propriedades do script antes de testar.');
    return;
  }
  Logger.log(`Testando login SGP com usuário "${SGP_USER}" (senha: ${SGP_PASS.length} chars)…`);
  try {
    const cookie = sgpGetSession();
    Logger.log(cookie.includes('sessionid')
      ? '✅ Login OK — sessão autenticada (sessionid presente). Acesso do SGP válido.'
      : '⚠️ Resposta sem sessionid — verifique se o usuário tem permissão de acesso.');
  } catch (err) {
    Logger.log('❌ Falha no login SGP: ' + err.message);
  }
}

/**
 * DEBUG (Fase 3): inspeciona o relatório de Ativações do SGP.
 *   1) abre sessão  2) lê o formulário  3) lista os nomes dos campos
 *   4) consulta o período e loga um trecho do HTML da tabela de resultados.
 * Rode no editor: sgpInspecionarAtivacoes('01/06/2026','30/06/2026')
 * e cole o Log aqui pra eu mapear as colunas e montar o parser.
 */
function sgpInspecionarAtivacoes(dataIni, dataFim) {
  dataIni = dataIni || '01/06/2026';
  dataFim = dataFim || '30/06/2026';
  const RELAT  = SGP_BASE + '/admin/relatorios/contrato/ativacoes/';
  const cookie = sgpGetSession();

  // O relatório é consultado via GET com querystring (POST devolve 405).
  const url = RELAT
    + '?data_inicial=' + encodeURIComponent(dataIni)
    + '&data_final='   + encodeURIComponent(dataFim);
  const resp = UrlFetchApp.fetch(url, {
    method: 'get', followRedirects: true, muteHttpExceptions: true,
    headers: { Cookie: cookie, Referer: RELAT },
  });
  const out = resp.getContentText();
  Logger.log(`GET (${dataIni}–${dataFim}) → HTTP ${resp.getResponseCode()} (${out.length} chars)`);
  Logger.log(`Resultados na página? Total Receita: ${out.includes('Total Receita')} | Total de Contratos: ${out.includes('Total de Contratos')}`);
  Logger.log(`Contagem → <table>: ${(out.match(/<table/gi)||[]).length} | <tr>: ${(out.match(/<tr/gi)||[]).length} | <th>: ${(out.match(/<th/gi)||[]).length} | <td>: ${(out.match(/<td/gi)||[]).length}`);

  // loga as tabelas que parecem conter os resultados (cliente/contrato/plano/status)
  const tabelas = out.match(/<table[\s\S]*?<\/table>/gi) || [];
  let achou = false;
  tabelas.forEach((t, i) => {
    if (/contrato|cliente|nome|plano|status/i.test(t)) {
      achou = true;
      Logger.log(`TABELA ${i} (${t.length} chars):\n${t.substr(0, 2200)}`);
    }
  });
  if (!achou) {
    const tIdx = out.search(/<table/i);
    Logger.log('Nenhuma tabela de resultados clara. 1ª <table>:\n' + (tIdx >= 0 ? out.substr(tIdx, 2000) : '(nenhuma <table>)'));
  }
}

// ── Parser do relatório de Ativações ───────────────────────────────────────
function _sgpStripTags(s) {
  return (s || '').replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&')
    .replace(/\s+/g, ' ').trim();
}
function _sgpCell(raw) {
  let c = (raw || '').replace(/<\/td>[\s\S]*$/i, ''); // corta o fechamento e o que vier depois
  c = c.replace(/^[^>]*>/, '');                         // remove o resto do <td ...>
  return _sgpStripTags(c);
}

/**
 * Consulta o relatório de Ativações e devolve os registros estruturados.
 * statusAtual (opcional) filtra pelo status atual do contrato.
 */
function sgpBuscarAtivacoes(dataIni, dataFim, statusAtual) {
  const RELAT  = SGP_BASE + '/admin/relatorios/contrato/ativacoes/';
  const cookie = sgpGetSession();
  let url = RELAT
    + '?data_inicial=' + encodeURIComponent(dataIni)
    + '&data_final='   + encodeURIComponent(dataFim);
  if (statusAtual) url += '&status_atual=' + encodeURIComponent(statusAtual);

  const html = UrlFetchApp.fetch(url, {
    method: 'get', followRedirects: true, muteHttpExceptions: true,
    headers: { Cookie: cookie, Referer: RELAT },
  }).getContentText();

  const tabela = (html.match(/<table[^>]*class="tablelist[^"]*"[\s\S]*?<\/table>/i) || [])[0] || '';
  const tbody  = (tabela.match(/<tbody[\s\S]*?<\/tbody>/i) || [])[0] || tabela;

  const linhas = tbody.split(/<tr[\s>]/i).slice(1);
  const registros = [];
  linhas.forEach(linha => {
    if (!/\/admin\/cliente\//i.test(linha)) return;               // só linhas de contrato
    const cel = linha.split(/<td[\s>]/i).slice(1);
    if (cel.length < 9) return;

    // col 0 — Contrato: <a href="/admin/cliente/ID/contratos/">NUM - NOME</a>
    const mCli      = cel[0].match(/\/admin\/cliente\/(\d+)\/contratos\/"?>([^<]*)</i);
    const clienteId = mCli ? mCli[1] : '';
    const txtLink   = mCli ? mCli[2].trim() : _sgpCell(cel[0]);
    const mNum      = txtLink.match(/^(\d+)\s*-\s*([\s\S]+)$/);
    const contrato  = mNum ? mNum[1] : '';
    let   nome      = (mNum ? mNum[2] : txtLink).trim();
    nome = nome.replace(/^[\d.\-\/]+\s+/, '').trim();   // remove CPF/CNPJ que às vezes precede o nome

    const pop      = _sgpCell(cel[1]);
    const planoRaw = _sgpCell(cel[2]);
    const valores  = planoRaw.match(/R\$\s*([\d.]+,\d{2})/g) || [];
    const valor    = valores.length ? valores[valores.length - 1].replace(/R\$\s*/, '') : '';
    const plano    = planoRaw.replace(/^\d+\s*-\s*/, '').split(/\s*\/\//)[0].trim();

    const tecnico  = _sgpCell(cel[4]);
    const dataFin  = _sgpCell(cel[5]);
    const motivo   = _sgpCell(cel[6]);
    const usuario  = _sgpCell(cel[7]);
    const dataAtiv = (_sgpCell(cel[8]).match(/\d{2}\/\d{2}\/\d{4}/) || [''])[0];

    registros.push({
      contrato, clienteId, nome, pop, plano, valor,
      dataAtivacao: dataAtiv, tecnico, usuario,
      dataFinalizacao: dataFin, motivo, statusAtual: statusAtual || '',
    });
  });
  return registros;
}

/** DEBUG: roda o parser de junho + lista as opções do filtro status_atual. */
function sgpDebugAtivacoes() {
  const recs = sgpBuscarAtivacoes('01/06/2026', '30/06/2026');
  Logger.log('Total parseado: ' + recs.length);
  Logger.log('Amostra (3):\n' + JSON.stringify(recs.slice(0, 3), null, 2));

  const cookie = sgpGetSession();
  const html   = UrlFetchApp.fetch(SGP_BASE + '/admin/relatorios/contrato/ativacoes/', {
    headers: { Cookie: cookie }, muteHttpExceptions: true,
  }).getContentText();
  const sel  = (html.match(/<select[^>]*name="status_atual"[\s\S]*?<\/select>/i) || [''])[0];
  const opts = [...sel.matchAll(/<option[^>]*value="([^"]*)"[^>]*>([^<]*)</gi)]
    .map(m => `"${m[1]}" = ${m[2].trim()}`);
  Logger.log('Opções status_atual (' + opts.length + '):\n' + opts.join('\n'));
}

// ── Sincronização SGP → Dashboard (comercial/vendas) ────────────────────────
const SGP_STATUS = {
  '1': 'Ativo', '7': 'Ativo V. Reduzida', '4': 'Suspenso',
  '3': 'Cancelado', '2': 'Inativo', '6': 'Novo', '5': 'Inviabilidade Técnica',
};
const _MES_ABBR = ['jan.','fev.','mar.','abr.','mai.','jun.','jul.','ago.','set.','out.','nov.','dez.'];
function _mesLabelBR(br) {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec((br || '').trim());
  return m ? `${_MES_ABBR[+m[2] - 1]}/${m[3]}` : '';
}
function _normNome(s) {
  return (s || '').toString().toUpperCase().normalize('NFD')
    .replace(/[̀-ͯ]/g, '').replace(/[^A-Z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ').trim();
}

/** Coleta as ativações do período já com o status atual de cada contrato. */
function sgpColetarAtivacoesComStatus(dataIni, dataFim) {
  const mapa = {}; // contrato -> registro
  Object.keys(SGP_STATUS).forEach(code => {
    sgpBuscarAtivacoes(dataIni, dataFim, code).forEach(r => {
      r.statusContrato = SGP_STATUS[code];
      mapa[r.contrato] = r;
    });
  });
  return Object.values(mapa);
}

/**
 * Sincroniza as vendas do SGP com o dashboard (comercial/vendas no Firebase).
 *   gravar = false  → DRY-RUN: só relata, não grava nada (padrão)
 *   gravar = true   → aplica via PATCH (casa por nome, preserva campos manuais)
 * Ex.: sgpSincronizarComercial('01/06/2026','30/06/2026')          // dry-run junho
 *      sgpSincronizarComercial('01/01/2026','30/06/2026', true)    // aplica jan–jun
 */
function sgpSincronizarComercial(dataIni, dataFim, gravar) {
  dataIni = dataIni || '01/01/2026';
  dataFim = dataFim || '30/06/2026';

  const sgpRecs = sgpColetarAtivacoesComStatus(dataIni, dataFim);
  Logger.log(`SGP: ${sgpRecs.length} contratos no período ${dataIni}–${dataFim}`);

  const atualRaw = UrlFetchApp.fetch(`${FIREBASE_RTDB_URL}/comercial/vendas.json`, {
    muteHttpExceptions: true,
  }).getContentText();
  const atual  = JSON.parse(atualRaw || '{}') || {};
  // índices: por contrato (já sincronizado antes) e por nome (lista p/ consumir 1-a-1)
  const porContrato = {}, porNome = {};
  Object.entries(atual).forEach(([key, v]) => {
    if (v.contrato) porContrato[v.contrato] = { key, v };
    const n = _normNome(v.nome);
    (porNome[n] = porNome[n] || []).push({ key, v });
  });

  const updates = {};
  let nCasados = 0, nNovos = 0;
  const novos = [], casados = [];

  sgpRecs.forEach(r => {
    const mesLabel  = _mesLabelBR(r.dataAtivacao);
    const churn     = r.statusContrato === 'Cancelado' ? 'Sim' : null;

    // casa 1º por número de contrato; depois por nome (consumindo da lista)
    let match = porContrato[r.contrato] || null;
    if (!match) {
      const lista = porNome[_normNome(r.nome)];
      if (lista && lista.length) match = lista.shift();
    }

    if (match) {
      nCasados++;
      const v = match.v;
      updates[match.key] = Object.assign({}, v, {
        id:    v.id || r.contrato,
        contrato: r.contrato,
        pop:   r.pop   || v.pop,
        plano: r.plano || v.plano,
        valor: r.valor || v.valor,
        dataVenda: r.dataAtivacao || v.dataVenda,
        mes:   v.mes || mesLabel,
        statusContrato: r.statusContrato,
        churn: churn || v.churn || 'Não',
      });
      casados.push(`${r.contrato} ↔ ${r.nome} (${r.statusContrato})`);
    } else {
      nNovos++;
      updates['sgp_' + r.contrato] = {
        id: r.contrato, contrato: r.contrato, nome: r.nome,
        dataVenda: r.dataAtivacao, pop: r.pop, bairro: '',
        plano: r.plano, valor: r.valor, vencimento: '',
        status: '', formaPgto: '', responsavel: '',
        churn: churn || 'Não', equip: '', motivo: '', obs: '',
        diasAtraso: '', mes: mesLabel, scoreChurn: '',
        statusContrato: r.statusContrato,
      };
      novos.push(`${r.contrato} - ${r.nome} (${r.statusContrato})`);
    }
  });

  Logger.log(`Casados por nome (atualizar): ${nCasados} | Novos: ${nNovos}`);
  if (novos.length)   Logger.log('NOVOS:\n'   + novos.slice(0, 60).join('\n'));
  if (casados.length) Logger.log('CASADOS:\n' + casados.slice(0, 60).join('\n'));

  if (!gravar) {
    Logger.log('🔎 DRY-RUN — nada foi gravado. Para aplicar: sgpSincronizarComercial(ini, fim, true)');
    return { ok: true, dryRun: true, total: sgpRecs.length, atualizados: nCasados, novos: nNovos };
  }
  const resp = UrlFetchApp.fetch(`${FIREBASE_RTDB_URL}/comercial/vendas.json`, {
    method: 'patch', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify(updates),
  });
  const http = resp.getResponseCode();
  Logger.log(`✅ PATCH Firebase → HTTP ${http} | ${nCasados} atualizados, ${nNovos} novos`);
  return { ok: http >= 200 && http < 300, http, total: sgpRecs.length, atualizados: nCasados, novos: nNovos };
}

/** Período padrão do sync: 01/01 do ano corrente até hoje (datas dinâmicas). */
function _periodoSyncBR() {
  const tz = 'America/Sao_Paulo';
  const hoje = new Date();
  const fim = Utilities.formatDate(hoje, tz, 'dd/MM/yyyy');
  const ini = '01/01/' + Utilities.formatDate(hoje, tz, 'yyyy');
  return { ini, fim };
}

/**
 * ATALHO / handler do trigger: roda o sync do ano corrente GRAVANDO no Firebase.
 * Selecione no dropdown e clique ▶ Executar para aplicar manualmente.
 * Também é a função chamada pelo trigger diário e pelo endpoint sgp_sync.
 */
function sgpAplicarSincronizacao() {
  const p = _periodoSyncBR();
  return sgpSincronizarComercial(p.ini, p.fim, true);
}

/**
 * Cria (ou recria) o trigger diário que roda sgpAplicarSincronizacao todo dia.
 * Rode UMA vez no editor. Remove triggers antigos do mesmo handler antes de criar.
 * Ajuste a hora em atHour() se quiser (formato 24h, fuso do projeto).
 */
function criarTriggerDiarioSGP() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sgpAplicarSincronizacao') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sgpAplicarSincronizacao')
    .timeBased().everyDays(1).atHour(6).create();
  Logger.log('✅ Trigger diário criado: sgpAplicarSincronizacao roda todo dia entre 6h–7h.');
}

/** Remove o trigger diário do sync (caso queira desativar a automação). */
function removerTriggersSGP() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sgpAplicarSincronizacao') { ScriptApp.deleteTrigger(t); n++; }
  });
  Logger.log(`🗑️ ${n} trigger(s) removido(s).`);
}

/**
 * DIAGNÓSTICO: quebra as ativações do período por MÊS de ativação e por status.
 * Use para conferir contra os números que o SGP mostra mês a mês
 * (ex.: junho = quantos contratos realmente ativaram em junho).
 * Selecione no dropdown e clique ▶ Executar (não grava nada).
 */
function sgpAtivacoesPorMes() {
  const p = _periodoSyncBR();
  const dataIni = p.ini, dataFim = p.fim;
  const recs = sgpColetarAtivacoesComStatus(dataIni, dataFim);
  Logger.log(`SGP: ${recs.length} contratos no período ${dataIni}–${dataFim}\n`);

  const porMes = {};   // 'mm/aaaa' -> { total, status:{} }
  let semData = 0;
  recs.forEach(r => {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec((r.dataAtivacao || '').trim());
    if (!m) { semData++; return; }
    const chave = `${m[2].padStart(2, '0')}/${m[3]}`;
    const bloco = porMes[chave] = porMes[chave] || { total: 0, status: {} };
    bloco.total++;
    const st = r.statusContrato || '—';
    bloco.status[st] = (bloco.status[st] || 0) + 1;
  });

  Object.keys(porMes).sort().forEach(mes => {
    const b = porMes[mes];
    const det = Object.entries(b.status).map(([s, n]) => `${s}:${n}`).join('  ');
    Logger.log(`${mes}  →  ${b.total} ativações   (${det})`);
  });
  if (semData) Logger.log(`(sem data de ativação: ${semData})`);
}

function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();
  if (action === 'sgp_fotos') {
    const ocId = (e.parameter.ocorrencia_id || '').trim();
    if (!ocId) return _resposta({ ok: false, erro: 'Informe ocorrencia_id' });
    return _resposta(sgpBuscarFotos(ocId));
  }
  if (action === 'sgp_sync') {
    try {
      const r = sgpAplicarSincronizacao();
      return _resposta(Object.assign({ ok: true }, r));
    } catch (err) {
      return _resposta({ ok: false, erro: err.toString() });
    }
  }
  return _resposta({ ok: false, erro: 'Use ?action=sgp_fotos&ocorrencia_id=XXXX ou ?action=sgp_sync' });
}

/**
 * Abre sessão no SGP via login form e retorna a string de cookies autenticada.
 *
 * FIXES aplicados:
 *  - URL com barra final (/accounts/login/) para evitar redirect GET que perde o POST
 *  - followRedirects: false no POST para capturar o Set-Cookie da resposta 302
 *  - CSRF extraído do HTML do formulário (mais confiável que do cookie)
 *  - Payload como string URL-encoded explícita
 */
function sgpGetSession() {
  // 1. GET na página de login para capturar CSRF (via HTML + cookie)
  const loginGet  = UrlFetchApp.fetch(SGP_BASE + '/accounts/login/', {
    followRedirects: true,
    muteHttpExceptions: true,
  });
  const loginHtml = loginGet.getContentText();

  // Extrai csrfmiddlewaretoken do HTML do formulário (mais robusto)
  const csrfFromHtml = (
    loginHtml.match(/name="csrfmiddlewaretoken"\s+value="([^"]+)"/i) ||
    loginHtml.match(/csrfmiddlewaretoken[^>]*value="([^"]+)"/i) ||
    []
  )[1] || '';

  // Também pega dos cookies (fallback)
  const rawCookies = loginGet.getAllHeaders()['Set-Cookie'] || '';
  const cookieArr  = Array.isArray(rawCookies) ? rawCookies : [rawCookies];
  const cookieMap  = {};
  cookieArr.forEach(c => {
    const part = c.split(';')[0].trim();
    const eq   = part.indexOf('=');
    if (eq > 0) cookieMap[part.slice(0, eq).trim()] = part.slice(eq + 1).trim();
  });

  const csrfToken = csrfFromHtml || cookieMap['csrftoken'] || '';
  const cookieJar = Object.entries(cookieMap).map(([k, v]) => `${k}=${v}`).join('; ');

  // 2. POST com credenciais — followRedirects: FALSE para capturar Set-Cookie da 302
  //    (quando followRedirects:true, o Apps Script descarta os cookies do redirect)
  const postPayload = [
    'csrfmiddlewaretoken=' + encodeURIComponent(csrfToken),
    'username='            + encodeURIComponent(SGP_USER),
    'password='            + encodeURIComponent(SGP_PASS),
    'next='                + encodeURIComponent('/admin/'),
  ].join('&');

  const loginPost = UrlFetchApp.fetch(SGP_BASE + '/accounts/login/', {
    method: 'post',
    followRedirects: false,          // ← CRÍTICO: captura cookies da 302
    muteHttpExceptions: true,
    payload: postPayload,
    headers: {
      Cookie:         cookieJar,
      Referer:        SGP_BASE + '/accounts/login/',
      'Content-Type': 'application/x-www-form-urlencoded',
    },
  });

  const postCode    = loginPost.getResponseCode();
  const postCookies = loginPost.getAllHeaders()['Set-Cookie'] || '';
  const postArr     = Array.isArray(postCookies) ? postCookies : [postCookies];

  // Merge: atualiza cookieMap com sessionid vindo da 302
  postArr.forEach(c => {
    const part = c.split(';')[0].trim();
    const eq   = part.indexOf('=');
    if (eq > 0) cookieMap[part.slice(0, eq).trim()] = part.slice(eq + 1).trim();
  });

  const finalCookie = Object.entries(cookieMap).map(([k, v]) => `${k}=${v}`).join('; ');

  // Login bem-sucedido = 302 (redirect para /admin/) ou sessionid presente
  if (postCode !== 302 && !finalCookie.includes('sessionid')) {
    throw new Error(`Login SGP falhou — HTTP ${postCode}. Verifique usuário/senha em SGP_USER/SGP_PASS.`);
  }

  return finalCookie;
}

/**
 * Busca até 10 fotos da ocorrência no SGP e retorna como base64.
 */
function sgpBuscarFotos(ocorrenciaId) {
  try {
    const cookieStr = sgpGetSession();

    // Listar anexos da ocorrência
    const listUrl  = `${SGP_BASE}/admin/atendimento/ocorrencia/${ocorrenciaId}/anexo/list/`;
    const listResp = UrlFetchApp.fetch(listUrl, {
      method: 'get',
      muteHttpExceptions: true,
      headers: { Cookie: cookieStr },
    });

    if (listResp.getResponseCode() !== 200) {
      return { ok: false, erro: `SGP retornou HTTP ${listResp.getResponseCode()} em ${listUrl}` };
    }

    const html = listResp.getContentText();

    // Extrai id de cada anexo — padrões conhecidos do SGP:
    // href="/admin/atendimento/ocorrencia/anexo/1852/get/"
    const fotos = [];
    const seenIds = new Set();

    // Padrão 1 (galeria com id + data + por)
    const re1 = /href="\/admin\/atendimento\/ocorrencia\/anexo\/(\d+)\/get\/"[^>]*>[\s\S]{0,600}?<p[^>]*>Data:\s*([^<]+)<\/p>[\s\S]{0,200}?<p[^>]*>Por:\s*([^<]+)<\/p>/g;
    let m;
    while ((m = re1.exec(html)) !== null) {
      if (!seenIds.has(m[1])) {
        seenIds.add(m[1]);
        fotos.push({ id: m[1], data_hora: m[2].trim(), por: m[3].trim() });
      }
    }

    // Padrão 2 (só href, sem metadados)
    const re2 = /href="\/admin\/atendimento\/ocorrencia\/anexo\/(\d+)\/get\/"/g;
    while ((m = re2.exec(html)) !== null) {
      if (!seenIds.has(m[1])) {
        seenIds.add(m[1]);
        fotos.push({ id: m[1], data_hora: '', por: '' });
      }
    }

    if (!fotos.length) {
      return { ok: true, fotos: [], total: 0, msg: 'Nenhum anexo encontrado nesta OS.' };
    }

    // Baixa até 10 imagens como base64 (evita timeout do Apps Script)
    const resultado = fotos.slice(0, 10).map(f => {
      try {
        const imgUrl  = `${SGP_BASE}/admin/atendimento/ocorrencia/anexo/${f.id}/get/?noattachment=1`;
        const imgResp = UrlFetchApp.fetch(imgUrl, {
          method: 'get',
          muteHttpExceptions: true,
          headers: { Cookie: cookieStr },
        });
        const ct  = (imgResp.getHeaders()['Content-Type'] || 'image/jpeg').split(';')[0].trim();
        const b64 = Utilities.base64Encode(imgResp.getBlob().getBytes());
        return {
          id:           f.id,
          data_hora:    f.data_hora,
          por:          f.por,
          content_type: ct,
          base64:       b64,
        };
      } catch (imgErr) {
        return { id: f.id, data_hora: f.data_hora, por: f.por, erro: imgErr.toString() };
      }
    });

    return { ok: true, fotos: resultado, total: fotos.length };

  } catch (err) {
    return { ok: false, erro: err.toString() };
  }
}

/** Teste manual — execute no Apps Script Editor para validar */
function testarSGPFotos() {
  // Substitua pelo número real de uma OS com fotos no SGP
  const resultado = sgpBuscarFotos('1');
  Logger.log(JSON.stringify({ ok: resultado.ok, total: resultado.total, erro: resultado.erro }));
  if (resultado.fotos && resultado.fotos.length) {
    Logger.log(`Primeira foto: ${resultado.fotos[0].data_hora} por ${resultado.fotos[0].por}`);
  }
}


// ════════════════════════════════════════════════════════════════════════════
//  9. SINCRONIZAÇÃO AUTOMÁTICA SGP → IBNET OPERAÇÕES
// ════════════════════════════════════════════════════════════════════════════

/**
 * Roda automaticamente a cada 30 minutos (gatilho configurado em configurarGatilhos).
 *
 * Fluxo:
 *   1. Loga no SGP
 *   2. Lista OSes encerradas (página de lista do admin, ordem: mais recentes primeiro)
 *   3. Para cada OS ainda não importada:
 *      a. Baixa a página de detalhe e extrai: tipo, técnico, contrato, insumos
 *      b. Cria registro em Firebase cac/ativacoes/sgp_{pk}
 *   4. Registra PKs importados para não duplicar
 *
 * Técnico não precisa mexer no IBnet — só preenche a OS no SGP.
 */

const FIREBASE_RTDB_URL = 'https://ibnet-platform-default-rtdb.firebaseio.com';

// URL da lista de OS no SGP (descoberta via debugListaSGP)
// Filtros obrigatórios: data_cadastro_inicial / data_cadastro_final (DD/MM/YYYY HH:MM:SS)
// Link de cada OS na lista: /admin/atendimento/ocorrencia/os/{pk}/edit/
const SGP_OS_LIST_URL = `${SGP_BASE}/admin/atendimento/relatorios/ocorrencia/os/`;

function sincronizarSGP() {
  _sincronizarSGPComDatas(null, null, false);
}

/**
 * Limpa TODAS as ativações do Firebase (cac/ativacoes) e reimporta
 * apenas as OSes encerradas de Maio/2026 com todos os campos:
 * drop, S/N ONT via IA, fotos, técnico, cliente, contrato.
 */
function reimportarMaio2026() {
  Logger.log('🗑️ Limpando cac/ativacoes no Firebase…');
  // 1. Busca todas as chaves existentes (shallow=true retorna só os IDs, sem dados)
  const keysResp = UrlFetchApp.fetch(`${FIREBASE_RTDB_URL}/cac/ativacoes.json?shallow=true`, {
    method: 'get', muteHttpExceptions: true,
  });
  if (keysResp.getResponseCode() === 200) {
    const keysData = keysResp.getContentText();
    if (keysData && keysData !== 'null') {
      const keys = Object.keys(JSON.parse(keysData));
      Logger.log(`   Encontradas ${keys.length} ativações para remover…`);
      // 2. PATCH com todos os IDs → null (apaga cada filho sem mexer na raiz)
      const nullMap = {};
      keys.forEach(k => nullMap[k] = null);
      UrlFetchApp.fetch(`${FIREBASE_RTDB_URL}/cac/ativacoes.json`, {
        method: 'patch',
        contentType: 'application/json',
        payload: JSON.stringify(nullMap),
        muteHttpExceptions: true,
      });
      Logger.log('   ✅ Todas as ativações removidas.');
    } else {
      Logger.log('   ℹ️ Nenhuma ativação existente para remover.');
    }
  } else {
    Logger.log(`   ⚠️ Não foi possível listar ativações: HTTP ${keysResp.getResponseCode()}`);
  }

  // Limpa cache de PKs sincronizados
  PropertiesService.getScriptProperties().deleteProperty('SGP_PKS_SYNC');
  Logger.log('🗑️ Cache de PKs limpo.');

  // Importa somente Maio/2026
  Logger.log('📅 Reimportando OSes de Maio/2026…');
  _sincronizarSGPComDatas('01/05/2026 00:00:00', '31/05/2026 23:59:59', true);
}

/**
 * Núcleo do sync — aceita datas opcionais (formato DD/MM/YYYY HH:MM:SS).
 * Se iniStr/fimStr forem null, usa os últimos 60 dias (comportamento padrão).
 * forcarReimport=true ignora o cache e reimporta tudo no intervalo.
 */
function _sincronizarSGPComDatas(iniStr, fimStr, forcarReimport) {
  const props  = PropertiesService.getScriptProperties();
  const jaSync = forcarReimport
    ? new Set()
    : new Set(JSON.parse(props.getProperty('SGP_PKS_SYNC') || '[]'));

  Logger.log('🔄 Iniciando sincronização SGP → IBnet Operações…');

  let cookieStr;
  try {
    cookieStr = sgpGetSession();
  } catch (err) {
    Logger.log('❌ Falha no login SGP: ' + err);
    return;
  }

  const osList = _sgpListarOS(cookieStr, iniStr, fimStr);
  if (!osList.length) {
    Logger.log('ℹ️ Nenhuma OS encerrada encontrada no período.');
    return;
  }

  const novas = osList.filter(os => os.encerrada && !jaSync.has(String(os.pk)));
  Logger.log(`📋 ${osList.length} OS(es) total · ${osList.filter(o=>o.encerrada).length} encerradas · ${novas.length} nova(s) para importar`);

  let importadas = 0;
  novas.forEach(os => {
    try {
      const ativ = _sgpExtrairOS(cookieStr, os.pk, os);
      if (!ativ) return;

      const id = `sgp_${os.pk}`;
      _firebasePut(`cac/ativacoes/${id}`, ativ);

      jaSync.add(String(os.pk));
      importadas++;
      Logger.log(`✅ OS ${os.pk} importada → ${id}`);

      Utilities.sleep(400);
    } catch (err) {
      Logger.log(`⚠️ Erro OS pk=${os.pk}: ${err}`);
    }
  });

  props.setProperty('SGP_PKS_SYNC', JSON.stringify([...jaSync].slice(-2000)));
  Logger.log(`✅ Sync encerradas concluído · ${importadas} OS(es) importada(s).`);

  // ── Sync pendentes (OSes não encerradas) ──────────────────────────────────
  const pendentes = osList.filter(os => !os.encerrada);
  Logger.log(`⏳ Atualizando ${pendentes.length} OS(es) pendente(s) no Firebase…`);

  // 1. Remove do nó /pendentes/ qualquer OS que agora está encerrada
  const encerradasPks = new Set(osList.filter(o => o.encerrada).map(o => String(o.pk)));
  try {
    const kResp = UrlFetchApp.fetch(`${FIREBASE_RTDB_URL}/cac/pendentes.json?shallow=true`, {
      method: 'get', muteHttpExceptions: true,
    });
    if (kResp.getResponseCode() === 200) {
      const kd = kResp.getContentText();
      if (kd && kd !== 'null') {
        const toRemove = {};
        Object.keys(JSON.parse(kd)).forEach(k => {
          if (encerradasPks.has(k.replace('sgp_', ''))) toRemove[k] = null;
        });
        if (Object.keys(toRemove).length) {
          UrlFetchApp.fetch(`${FIREBASE_RTDB_URL}/cac/pendentes.json`, {
            method: 'patch', contentType: 'application/json',
            payload: JSON.stringify(toRemove), muteHttpExceptions: true,
          });
          Logger.log(`   🗑 ${Object.keys(toRemove).length} OS(es) movida(s) de pendentes → encerradas.`);
        }
      }
    }
  } catch(e) {
    Logger.log(`⚠️ Erro ao limpar pendentes: ${e}`);
  }

  // 2. Grava/atualiza as pendentes atuais
  pendentes.forEach(os => {
    try {
      _firebasePut(`cac/pendentes/sgp_${os.pk}`, {
        id:           `sgp_${os.pk}`,
        sgp_pk:       os.pk,
        cliente:      os.clienteNome,
        contrato:     os.contrato,
        motivo:       os.motivo,
        tecnico:      os.tecnico,
        data:         os.dataISO,
        sgpStatus:    os.sgpStatus || 'Aberta',
        origem:       'sgp_auto',
        atualizadoEm: new Date().toISOString(),
      });
      Utilities.sleep(150);
    } catch(err) {
      Logger.log(`⚠️ Erro ao salvar pendente OS ${os.pk}: ${err}`);
    }
  });
  Logger.log(`⏳ Pendentes atualizadas: ${pendentes.length}`);
}

/**
 * Lista OSes no SGP dentro de um período.
 * iniStr / fimStr: "DD/MM/YYYY HH:MM:SS" — se null usa últimos 60 dias.
 * Retorna [{pk, encerrada, contrato, clienteNome, motivo, tecnico, dataISO}]
 */
function _sgpListarOS(cookieStr, iniStr, fimStr) {
  const hoje = new Date();
  const fmt  = d => `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
  const ini  = new Date(hoje); ini.setDate(ini.getDate() - 60);

  const dataIni = iniStr || (fmt(ini)  + ' 00:00:00');
  const dataFim = fimStr || (fmt(hoje) + ' 23:59:59');

  Logger.log(`📅 Buscando OSes de ${dataIni} até ${dataFim}`);

  const listUrl = SGP_OS_LIST_URL
    + '?data_cadastro_inicial=' + encodeURIComponent(dataIni)
    + '&data_cadastro_final='   + encodeURIComponent(dataFim);

  const resp = UrlFetchApp.fetch(listUrl, {
    method: 'get', muteHttpExceptions: true,
    headers: { Cookie: cookieStr }, followRedirects: true,
  });

  const code = resp.getResponseCode();
  if (code !== 200) {
    Logger.log('⚠️ Erro ao listar OSes: HTTP ' + code);
    return [];
  }

  const html   = resp.getContentText();
  const osList = [];
  const seen   = new Set();

  // Percorre blocos <tr> (cada OS tem um checkbox name="ordemservico[]")
  const rows = html.split('<tr');
  for (const row of rows) {
    const pkMatch = row.match(/name="ordemservico\[\]"[^>]*value="(\d+)"/);
    if (!pkMatch) continue;
    const pk = pkMatch[1];
    if (seen.has(pk)) continue;
    seen.add(pk);

    // Detecção de encerrada: método original confiável (red_bold ou texto simples)
    const encerrada = /class="red_bold"[^>]*>\s*Encerrada/i.test(row)
                   || />Encerrada\s*</.test(row);

    // Status SGP para display (texto simples, sem afetar a lógica de sync)
    let sgpStatus = encerrada ? 'Encerrada' : 'Aberta';
    if (!encerrada) {
      if (/Em\s+Andamento/i.test(row))  sgpStatus = 'Em Andamento';
      else if (/Agendada/i.test(row))   sgpStatus = 'Agendada';
      else if (/Pendente/i.test(row))   sgpStatus = 'Pendente';
      else if (/Cancelada/i.test(row))  sgpStatus = 'Cancelada';
    }

    // Cliente/Contrato: "296 - SHERLEY SOARES BELE"
    const cliMatch  = row.match(/href="\/admin\/cliente\/\d+\/contratos\/">\s*(\d+)\s*-\s*([^<]+)</i);
    const contrato  = cliMatch ? cliMatch[1].trim() : '';
    const clienteNome = cliMatch ? cliMatch[2].trim() : '';

    // Motivo: texto simples em <td>
    const motivoMatch = row.match(/>\s*(Instala[çc][ãa]o[^<]{0,20}|Remo[çc][ãa]o[^<]{0,20}|Preventiva|Corretiva|Financeiro|Mudan[çc]a[^<]{0,20})\s*<\/td>/i);
    const motivo = motivoMatch ? motivoMatch[1].trim() : '';

    // Data encerramento: usa o ÚLTIMO sort span (14 dígitos) da linha
    // O SGP exibe: [data_cadastro] ... [data_encerramento] — o último é o encerramento
    const sortAll = [...row.matchAll(/<span class="sort">(\d{14})<\/span>/g)].map(m => m[1]);
    const sortUsar = sortAll.length ? sortAll[sortAll.length - 1] : null;
    const dataISO = sortUsar
      ? `${sortUsar.slice(0,4)}-${sortUsar.slice(4,6)}-${sortUsar.slice(6,8)}`
      : new Date().toISOString().slice(0, 10);

    // Técnico: padrão "tecnico.nome" ou primeiro nome após as datas
    const tecMatch = row.match(/>\s*(tecnico\.\w+)\s*</i);
    const tecnico  = tecMatch ? tecMatch[1].trim() : '';

    osList.push({ pk, encerrada, sgpStatus, contrato, clienteNome, motivo, tecnico, dataISO });
  }

  Logger.log(`📋 _sgpListarOS: ${osList.length} OS(es) · ${osList.filter(o=>o.encerrada).length} encerradas`);
  return osList;
}

/**
 * Visita a página de edição da OS para extrair o Serviço Prestado
 * (ONT, drop, conectores) e contar fotos.
 * Monta o objeto Firebase-ready com dados do list + edit page.
 *
 * URL de edição: /admin/atendimento/ocorrencia/os/{pk}/edit/
 */
function _sgpExtrairOS(cookieStr, pk, osInfo) {
  const editUrl = `${SGP_BASE}/admin/atendimento/ocorrencia/os/${pk}/edit/`;
  const resp    = UrlFetchApp.fetch(editUrl, {
    method: 'get', muteHttpExceptions: true,
    headers: { Cookie: cookieStr }, followRedirects: true,
  });

  let servicoTxt  = '';
  let qtdFotos    = 0;
  let ocorrenciaPk = pk; // fallback: usa o próprio pk da OS

  if (resp.getResponseCode() === 200) {
    const html = resp.getContentText();

    // ── Ocorrência pk real (diferente do pk da OS!) ───────────────────────
    // A página de edição contém: href="/admin/atendimento/ocorrencia/XXXX/anexo/list/"
    // onde XXXX é o pk da Ocorrencia mãe (modelo diferente de OrdemServico)
    const ocMatch = html.match(/href="\/admin\/atendimento\/ocorrencia\/(\d+)\/anexo\/list\//);
    if (ocMatch) ocorrenciaPk = ocMatch[1];

    // Serviço Prestado (textarea na página de edição)
    const mServ = html.match(/id="id_servico_prestado"[^>]*>([\s\S]*?)<\/textarea>/i)
               || html.match(/[Ss]ervi[çc]o\s*[Pp]restado[\s\S]{0,400}?<textarea[^>]*>([\s\S]*?)<\/textarea>/i);
    servicoTxt = mServ ? mServ[1].replace(/&amp;/g,'&').replace(/&#\d+;/g,'').trim() : '';
  } else {
    Logger.log(`⚠️ OS ${pk} edit page retornou HTTP ${resp.getResponseCode()}`);
  }

  // Conta fotos via URL de anexos usando o ocorrencia pk correto
  try {
    const anexoUrl  = `${SGP_BASE}/admin/atendimento/ocorrencia/${ocorrenciaPk}/anexo/list/`;
    const anexoResp = UrlFetchApp.fetch(anexoUrl, {
      method: 'get', muteHttpExceptions: true,
      headers: { Cookie: cookieStr }, followRedirects: true,
    });
    if (anexoResp.getResponseCode() === 200) {
      const aHtml = anexoResp.getContentText();
      // Conta URLs únicas de imagens: /anexo/XXXX/get/
      const matches = aHtml.match(/\/atendimento\/ocorrencia\/anexo\/(\d+)\/get\//g) || [];
      const uniq    = new Set(matches.map(m => m.match(/\/(\d+)\//)[1]));
      qtdFotos = uniq.size;
    }
  } catch(_) { /* silencioso */ }

  const insumos    = _sgpParseServico(servicoTxt);
  const motivoLow  = (osInfo.motivo || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'');
  let tipo = 'manutencao';
  if      (motivoLow.includes('instala')) tipo = 'instalacao';
  else if (motivoLow.includes('migra'))   tipo = 'migracao';
  else if (motivoLow.includes('infra'))   tipo = 'infra';
  else if (motivoLow.includes('remoca'))  tipo = 'manutencao';

  // Tenta ler S/N da ONT nas fotos via IA (só se tem ONT e fotos)
  let snOnt = null;
  if (insumos.ont > 0 && qtdFotos > 0) {
    snOnt = _sgpExtrairSNdaFoto(cookieStr, ocorrenciaPk);
  }

  Logger.log(`   OS ${pk} → ocorrencia_pk=${ocorrenciaPk} | fotos=${qtdFotos} | drop=${insumos.drop}m | ont=${insumos.ont} | sn=${snOnt||'-'} | servico="${servicoTxt.slice(0,50)}"`);

  const tecnicoNome = osInfo.tecnico || '';
  const agora       = new Date().toISOString();

  return {
    id:        `sgp_${pk}`,
    tipo,
    cliente:   osInfo.clienteNome || (osInfo.contrato ? `Contrato ${osInfo.contrato}` : `OS ${pk}`),
    pppoe:     '',
    contrato:  osInfo.contrato  || '',
    tecnico:   tecnicoNome,
    tecnicoId: `sgp_${tecnicoNome.replace(/\s+/g,'_').toLowerCase()}`,
    data:      osInfo.dataISO   || agora.slice(0, 10),
    criadoEm:  agora,
    sgp_pk:            pk,
    sgp_os:            pk,
    sgp_ocorrencia_pk: ocorrenciaPk,
    sgp_motivo:        osInfo.motivo   || '',
    sgp_servico:       servicoTxt      || '',   // texto do campo "Serviço Prestado" da OS
    origem:    'sgp_auto',
    ont:       { qtd: insumos.ont, ...(snOnt ? { sn: snOnt } : {}) },
    drop:      { metros: insumos.drop },
    conector:  { qtd: insumos.conector },
    ...(qtdFotos > 0 ? { sgpFotos: qtdFotos } : {}),
  };
}

/**
 * Extrai quantidades de insumos do campo "Serviço Prestado".
 * Exemplos reconhecidos:
 *   "2 conectores"  "1 ont"  "179 metros drop"  "179 mts drop"
 *   "78 metros"  "120 metros" (standalone — técnico omite palavra "drop")
 *   "drop 85m"  "3x conector"  "ONT: 1"
 *   "1004-845=159 mts drop"  "302-188= 114 mts drop"  (cálculo com resultado)
 */
function _sgpParseServico(texto) {
  if (!texto) return { ont: 0, drop: 0, conector: 0 };
  const ont      = parseInt((texto.match(/(\d+)\s*(?:x\s*)?ont/i)      || [])[1] || 0);
  const conector = parseInt((texto.match(/(\d+)\s*(?:x\s*)?conect/i)   || [])[1] || 0);

  // Drop — por ordem de prioridade:
  // 1. Cálculo com resultado: "1004-845=159 mts drop" ou "= 159 metros"
  // 2. Número antes de drop: "179 metros drop" / "114 mts drop"
  // 3. Drop antes de número: "drop: 85m"
  // 4. Fallback: "78 metros" standalone (técnicos geralmente omitem "drop")
  const dropRaw = (
    texto.match(/=\s*([\d.,]+)\s*(?:metros?|mts?)\s*(?:de\s*)?drop/i)
    || texto.match(/([\d.,]+)\s*(?:metros?|mts?)\s*(?:de\s*)?drop/i)
    || texto.match(/drop\s*[:\-]?\s*([\d.,]+)\s*(?:metros?|mts?|m\b)/i)
    || texto.match(/^\s*([\d.,]+)\s*(?:metros?|mts?)\s*$/im)
    || texto.match(/^\s*([\d.,]+)\s*(?:metros?|mts?)\b/im)
    || []
  )[1] || '0';
  const drop = parseFloat(dropRaw.replace(',', '.'));

  return { ont, drop, conector };
}

/**
 * Tenta ler o S/N da ONT nas fotos do SGP usando Claude Haiku (visão).
 * Baixa até MAX_FOTOS_SN fotos e retorna o primeiro S/N encontrado, ou null.
 * Só é chamado quando a OS tem ONT (qtd > 0) e fotos.
 */
const MAX_FOTOS_SN = 4; // máximo de fotos para tentar ler o S/N

function _sgpExtrairSNdaFoto(cookieStr, ocorrenciaPk) {
  try {
    // Lista anexos
    const aUrl  = `${SGP_BASE}/admin/atendimento/ocorrencia/${ocorrenciaPk}/anexo/list/`;
    const aResp = UrlFetchApp.fetch(aUrl, {
      method: 'get', muteHttpExceptions: true,
      headers: { Cookie: cookieStr }, followRedirects: true,
    });
    if (aResp.getResponseCode() !== 200) return null;

    const aHtml = aResp.getContentText();
    const ids   = [...new Set(
      (aHtml.match(/\/atendimento\/ocorrencia\/anexo\/(\d+)\/get\//g) || [])
        .map(m => m.match(/\/(\d+)\//)[1])
    )];
    if (!ids.length) return null;

    // Tenta cada foto até achar o S/N
    for (const id of ids.slice(0, MAX_FOTOS_SN)) {
      try {
        const imgUrl  = `${SGP_BASE}/admin/atendimento/ocorrencia/anexo/${id}/get/?noattachment=1`;
        const imgResp = UrlFetchApp.fetch(imgUrl, {
          method: 'get', muteHttpExceptions: true,
          headers: { Cookie: cookieStr },
        });
        if (imgResp.getResponseCode() !== 200) continue;

        const ct = (imgResp.getHeaders()['Content-Type'] || 'image/jpeg').split(';')[0].trim();
        if (!ct.startsWith('image/')) continue;

        const b64 = Utilities.base64Encode(imgResp.getBlob().getBytes());

        const aiResp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
          method: 'post', muteHttpExceptions: true,
          headers: {
            'x-api-key':         ANTHROPIC_KEY,
            'anthropic-version': '2023-06-01',
            'content-type':      'application/json',
          },
          payload: JSON.stringify({
            model:      'claude-haiku-4-5',
            max_tokens: 80,
            messages: [{
              role: 'user',
              content: [
                { type: 'image', source: { type: 'base64', media_type: ct, data: b64 } },
                { type: 'text',  text:
                  'Esta é a etiqueta traseira de um roteador/ONT de fibra óptica. ' +
                  'Encontre o número de série principal — indicado como "S/N", "SN", "Serial Number" ou "Serial No". ' +
                  'Responda APENAS com o número/código, sem texto adicional. ' +
                  'Se não encontrar claramente, responda: NAO_ENCONTRADO'
                }
              ]
            }]
          }),
        });

        if (aiResp.getResponseCode() !== 200) continue;
        const aiData = JSON.parse(aiResp.getContentText());
        const sn     = (aiData.content?.[0]?.text || '').trim().replace(/\s+/g,'').toUpperCase();

        if (sn && sn !== 'NAO_ENCONTRADO' && sn.length >= 4 && sn.length <= 30) {
          return sn;
        }
      } catch(_) { continue; }
    }
    return null;
  } catch(_) { return null; }
}

/**
 * Grava um objeto no Firebase RTDB via REST API (sem SDK).
 * Funciona porque as regras do banco permitem escrita não autenticada
 * (mesma configuração usada pelo painel no browser).
 */
function _firebasePut(path, data) {
  const url  = `${FIREBASE_RTDB_URL}/${path}.json`;
  const resp = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    payload: JSON.stringify(data),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) {
    Logger.log(`⚠️ Firebase PUT falhou [${resp.getResponseCode()}]: ${resp.getContentText().slice(0,200)}`);
  }
  return resp;
}

/** Teste manual — execute para forçar uma sincronização imediata */
function testarSyncSGP() {
  sincronizarSGP();
}

/**
 * DEBUG DE FOTOS — execute para ver como o SGP serve os anexos de uma OS.
 * Muda o PK para uma OS real que tenha fotos cadastradas no SGP.
 * ATENÇÃO: use o sgp_ocorrencia_pk (ex: 1586), NÃO o sgp_pk da OS (ex: 1538)!
 */
function debugFotosSGP() {
  const PK_TESTE = '1586'; // ← OCORRÊNCIA pk (não OS pk!) — veja sgp_ocorrencia_pk no Firebase

  Logger.log('=== DEBUG FOTOS SGP — OS ' + PK_TESTE + ' ===');

  let cookieStr;
  try {
    cookieStr = sgpGetSession();
    Logger.log('✅ Login OK');
  } catch(e) {
    Logger.log('❌ Login falhou: ' + e);
    return;
  }

  // ── 1. Página de edição (onde tentamos contar gallery-item) ──────────────
  const editUrl = `${SGP_BASE}/admin/atendimento/ocorrencia/os/${PK_TESTE}/edit/`;
  const editResp = UrlFetchApp.fetch(editUrl, {
    method: 'get', muteHttpExceptions: true,
    headers: { Cookie: cookieStr }, followRedirects: true,
  });
  const editHtml = editResp.getContentText();
  Logger.log(`\n── Página de edição (${editUrl})`);
  Logger.log(`HTTP: ${editResp.getResponseCode()} | Tamanho: ${editHtml.length} chars`);
  Logger.log(`gallery-item encontrados: ${(editHtml.match(/class="gallery-item/g)||[]).length}`);
  Logger.log(`Trechos com "foto"|"anex"|"imag": ` + JSON.stringify(
    (editHtml.match(/.{0,60}(foto|anex|imag).{0,60}/gi) || []).slice(0,5)
  ));
  // Primeiros 3000 chars (header) + últimos 3000 chars (onde ficam as fotos)
  Logger.log('Início da página (3000):\n' + editHtml.slice(0, 3000));
  Logger.log('Final da página (4000):\n' + editHtml.slice(-4000));

  // ── 2. URL de anexos (usada pelo proxy sgpBuscarFotos) ───────────────────
  const anexoUrl = `${SGP_BASE}/admin/atendimento/ocorrencia/${PK_TESTE}/anexo/list/`;
  const anexoResp = UrlFetchApp.fetch(anexoUrl, {
    method: 'get', muteHttpExceptions: true,
    headers: { Cookie: cookieStr }, followRedirects: true,
  });
  const anexoHtml = anexoResp.getContentText();
  Logger.log(`\n── URL de anexos (${anexoUrl})`);
  Logger.log(`HTTP: ${anexoResp.getResponseCode()} | Tamanho: ${anexoHtml.length} chars`);
  Logger.log('Conteúdo completo:\n' + anexoHtml.slice(0, 5000));

  // ── 3. URLs alternativas de anexo ────────────────────────────────────────
  const alts = [
    `${SGP_BASE}/admin/atendimento/ocorrencia/os/${PK_TESTE}/anexo/list/`,
    `${SGP_BASE}/admin/atendimento/ocorrencia/${PK_TESTE}/fotos/`,
    `${SGP_BASE}/admin/atendimento/ocorrencia/${PK_TESTE}/galeria/`,
  ];
  for (const url of alts) {
    const r = UrlFetchApp.fetch(url, { method:'get', muteHttpExceptions:true, headers:{Cookie:cookieStr}, followRedirects:true });
    Logger.log(`\n── Alt: ${url} → HTTP ${r.getResponseCode()} (${r.getContentText().length} chars)`);
  }
}

/**
 * DEBUG — executa isso para diagnosticar o login e a estrutura da lista de OS no SGP.
 * Versão aprimorada: mostra o HTTP code do POST de login + cookies obtidos.
 */
function debugListaSGP() {
  // ── Passo 1: diagnóstico detalhado do login ────────────────────────────
  Logger.log('=== DIAGNÓSTICO DE LOGIN SGP ===');

  // GET para pegar CSRF
  const loginGet  = UrlFetchApp.fetch(SGP_BASE + '/accounts/login/', {
    followRedirects: true, muteHttpExceptions: true,
  });
  const loginHtml = loginGet.getContentText();
  const csrfFromHtml = (
    loginHtml.match(/name="csrfmiddlewaretoken"\s+value="([^"]+)"/i) ||
    loginHtml.match(/csrfmiddlewaretoken[^>]*value="([^"]+)"/i) ||
    []
  )[1] || '';
  Logger.log('CSRF do HTML: ' + (csrfFromHtml ? csrfFromHtml.slice(0,12) + '…' : '❌ NÃO ENCONTRADO'));

  const rawCookies = loginGet.getAllHeaders()['Set-Cookie'] || '';
  const cookieArr  = Array.isArray(rawCookies) ? rawCookies : [rawCookies];
  const cookieMap  = {};
  cookieArr.forEach(c => {
    const part = c.split(';')[0].trim();
    const eq   = part.indexOf('=');
    if (eq > 0) cookieMap[part.slice(0, eq).trim()] = part.slice(eq + 1).trim();
  });
  const csrfToken = csrfFromHtml || cookieMap['csrftoken'] || '';
  const cookieJar = Object.entries(cookieMap).map(([k, v]) => `${k}=${v}`).join('; ');

  // POST de login SEM followRedirects
  const postPayload = [
    'csrfmiddlewaretoken=' + encodeURIComponent(csrfToken),
    'username='            + encodeURIComponent(SGP_USER),
    'password='            + encodeURIComponent(SGP_PASS),
    'next='                + encodeURIComponent('/admin/'),
  ].join('&');

  const loginPost = UrlFetchApp.fetch(SGP_BASE + '/accounts/login/', {
    method: 'post',
    followRedirects: false,
    muteHttpExceptions: true,
    payload: postPayload,
    headers: {
      Cookie: cookieJar,
      Referer: SGP_BASE + '/accounts/login/',
      'Content-Type': 'application/x-www-form-urlencoded',
    },
  });

  const postCode = loginPost.getResponseCode();
  Logger.log('HTTP POST login: ' + postCode + (postCode === 302 ? ' ✅ (redirect = login OK)' : ' ❌ (esperava 302)'));

  const postCookieRaw = loginPost.getAllHeaders()['Set-Cookie'] || '';
  const postCookieArr = Array.isArray(postCookieRaw) ? postCookieRaw : [postCookieRaw];
  postCookieArr.forEach(c => {
    const part = c.split(';')[0].trim();
    const eq   = part.indexOf('=');
    if (eq > 0) cookieMap[part.slice(0, eq).trim()] = part.slice(eq + 1).trim();
  });

  const finalCookie = Object.entries(cookieMap).map(([k, v]) => `${k}=${v}`).join('; ');
  Logger.log('sessionid presente: ' + (finalCookie.includes('sessionid') ? '✅ SIM' : '❌ NÃO'));
  Logger.log('Cookies finais: ' + finalCookie.slice(0, 200));

  if (postCode !== 302 && !finalCookie.includes('sessionid')) {
    Logger.log('❌ Login falhou — verifique SGP_USER e SGP_PASS');
    Logger.log('Resposta do POST (500 chars):\n' + loginPost.getContentText().slice(0, 500));
    return;
  }
  Logger.log('✅ Login SGP OK — testando acesso à lista de OS…\n');

  // ── Passo 2: testa as URLs de lista de OS identificadas no menu ──────
  Logger.log('=== TESTANDO URLs DE LISTA DE OS ===');

  // URLs de lista de OS identificadas no menu do SGP
  const urlsCandidatas = [
    `${SGP_BASE}/admin/atendimento/relatorios/ocorrencia/os/`,
    `${SGP_BASE}/admin/atendimento/agenda/ocorrencia/view/`,
    `${SGP_BASE}/admin/atendimento/relatorios/ocorrencia/`,
  ];

  for (const url of urlsCandidatas) {
    Logger.log(`\n── Testando: ${url}`);
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: { Cookie: finalCookie },
      followRedirects: true,
    });
    const code = resp.getResponseCode();
    const html = resp.getContentText();
    Logger.log(`   HTTP: ${code}`);

    if (code === 200 && !html.includes('id="login-form"')) {

      // Os nomes reais dos campos de data são:
      //   data_cadastro_inicial  (DD/MM/AAAA HH:MM:SS)
      //   data_cadastro_final    (DD/MM/AAAA HH:MM:SS)
      // (não data_inicio/data_fim como eu tentava antes!)
      const hoje = new Date();
      const dd   = String(hoje.getDate()).padStart(2,'0');
      const mm   = String(hoje.getMonth()+1).padStart(2,'0');
      const yyyy = hoje.getFullYear();
      const ini  = new Date(hoje); ini.setDate(ini.getDate()-60);
      const ddI  = String(ini.getDate()).padStart(2,'0');
      const mmI  = String(ini.getMonth()+1).padStart(2,'0');
      const yyyyI= ini.getFullYear();

      const filtUrl = url
        + '?data_cadastro_inicial=' + encodeURIComponent(`${ddI}/${mmI}/${yyyyI} 00:00:00`)
        + '&data_cadastro_final='   + encodeURIComponent(`${dd}/${mm}/${yyyy} 23:59:59`);

      Logger.log('Testando com campos corretos: ' + filtUrl);
      const filtResp = UrlFetchApp.fetch(filtUrl, {
        method: 'get', muteHttpExceptions: true,
        headers: { Cookie: finalCookie }, followRedirects: true,
      });
      const filtHtml = filtResp.getContentText();
      Logger.log('Tamanho resposta: ' + filtHtml.length + ' (sem filtro era 136295)');

      // Mostra os últimos 8000 chars (onde fica o tbody)
      Logger.log('\n── ÚLTIMOS 8000 chars (tbody) ──\n' + filtHtml.slice(-8000));

      // Conta <tr> no tbody
      const trCount = (filtHtml.match(/<tr[^>]*>/gi) || []).length;
      Logger.log('Total de <tr> no HTML: ' + trCount);

      // Busca qualquer link com número
      const nums = filtHtml.match(/href="[^"]*\/\d+[^"]*"/g) || [];
      Logger.log(`Links com número (${nums.length}):`);
      [...new Set(nums)].slice(0,30).forEach(l => Logger.log('  ' + l));

      break;
    } else if (html.includes('id="login-form"')) {
      Logger.log('   ↳ Retornou página de login');
    } else {
      Logger.log('   ↳ HTTP ' + code);
    }
  }
}

/** Limpa o cache de PKs importados (re-importa tudo na próxima rodada) */
function limparCacheSyncSGP() {
  PropertiesService.getScriptProperties().deleteProperty('SGP_PKS_SYNC');
  Logger.log('🗑️ Cache de sync SGP apagado. Próxima execução reimportará todas as OSes.');
  SpreadsheetApp.getUi().alert('Cache limpo! Na próxima sincronização automática todas as OSes serão reimportadas.');
}


// ════════════════════════════════════════════════════════════════════════════
//  GATILHOS E AGENDAMENTO
// ════════════════════════════════════════════════════════════════════════════

function configurarGatilhos() {
  // Remove todos os gatilhos antigos para evitar duplicatas
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Dias em atraso + Score → 07h todo dia
  ScriptApp.newTrigger('rodarCalculosDiarios')
    .timeBased().everyDays(1).atHour(7).create();

  // Alerta de inadimplência → 08h todo dia
  ScriptApp.newTrigger('enviarAlertaInadimplencia')
    .timeBased().everyDays(1).atHour(8).create();

  // Relatório semanal → segunda-feira às 07h
  ScriptApp.newTrigger('enviarRelatorioSemanal')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();

  // Backup semanal → sexta-feira às 18h
  ScriptApp.newTrigger('backupSemanal')
    .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(18).create();

  // Log de alterações → a cada edição na planilha (gatilho instalável)
  ScriptApp.newTrigger('registrarEdicao')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  // ── NOVO: Sincronização automática SGP → IBnet a cada 30 minutos ──────
  ScriptApp.newTrigger('sincronizarSGP')
    .timeBased().everyMinutes(30).create();

  Logger.log('✅ Todos os gatilhos configurados!');
  SpreadsheetApp.getUi().alert(
    '✅ Automação v4 completa!\n\n' +
    '• 07h00 todo dia     → Dias em atraso + Score de Churn\n' +
    '• 08h00 todo dia     → Alerta de inadimplência (e-mail)\n' +
    '• Segunda às 07h     → Relatório semanal de KPIs (e-mail)\n' +
    '• Sexta às 18h       → Backup semanal no Drive (e-mail)\n' +
    '• A cada edição      → Registro na aba 📋 Log\n' +
    '• A cada 30 minutos  → Sync automático SGP → Operações IBnet\n\n' +
    'Endpoint de instalações: implante como App da Web para ativar a sincronização automática.\n\n' +
    'E-mails: ' + EMAILS_GESTAO.join(' · ')
  );
}

/** Agrupa cálculos diários numa função só */
function rodarCalculosDiarios() {
  calcularDiasEmAtraso();
  calcularScoreChurn();
}

function removerGatilhos() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('🗑️  Todos os gatilhos removidos.');
  SpreadsheetApp.getUi().alert('Automação desativada. Execute "configurarGatilhos" para reativar.');
}


// ════════════════════════════════════════════════════════════════════════════
//  MENU E UTILITÁRIOS
// ════════════════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ IBnet')
    .addItem('🔄 Calcular dias em atraso agora',          'calcularDiasEmAtraso')
    .addItem('📊 Calcular score de churn agora',          'calcularScoreChurn')
    .addItem('⚠️  Enviar alerta inadimplência agora',     'enviarAlertaInadimplencia')
    .addItem('📧 Enviar relatório semanal agora',         'enviarRelatorioSemanal')
    .addItem('💾 Fazer backup agora',                     'backupSemanal')
    .addSeparator()
    .addItem('🔌 Testar endpoint de instalações',         'testarEndpoint')
    .addItem('📷 Testar proxy de fotos SGP',              'testarSGPFotos')
    .addSeparator()
    .addItem('🔄 Sincronizar SGP agora',                  'testarSyncSGP')
    .addItem('🗑️  Limpar cache de sync SGP',              'limparCacheSyncSGP')
    .addItem('📅 Reimportar Maio/2026 (limpa tudo)',      'reimportarMaio2026')
    .addSeparator()
    .addItem('⚙️  Configurar automação (fazer uma vez)',   'configurarGatilhos')
    .addItem('🗑️  Remover automação',                     'removerGatilhos')
    .addToUi();
}

// Normaliza string para comparação (remove acentos, minúscula, espaços extras)
function normaliza(str) {
  return (str || '').toString().trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function parsearData(valor) {
  if (!valor) return null;
  if (valor instanceof Date) { const d = new Date(valor); d.setHours(0,0,0,0); return d; }
  const s = valor.toString().trim();
  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m1) return new Date(+m1[3], +m1[2]-1, +m1[1]);
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2) return new Date(+m2[1], +m2[2]-1, +m2[3]);
  return null;
}

function parsearValor(str) {
  if (!str) return 0;
  return parseFloat(
    str.toString().replace(/R\$\s?/g,'').replace(/\./g,'').replace(',','.')
  ) || 0;
}
