/**
 * ============================================================
 *  IBNET TELECOM — Automação Google Sheets (v3)
 *  Funções:
 *    1. Cálculo automático de "Dias em atraso"
 *    2. Alerta diário de inadimplência por e-mail
 *    3. Relatório semanal de KPIs por e-mail
 *    4. Score de risco de churn por cliente
 *    5. Log automático de todas as alterações
 *    6. Backup semanal automático (sexta-feira)
 * ============================================================
 *
 *  ATUALIZAÇÃO (substitua o script antigo por este):
 *  1. Extensões → Apps Script → seleciona tudo → cola este arquivo
 *  2. Salva (Ctrl+S)
 *  3. Seleciona "configurarGatilhos" e clica ▶ (uma única vez)
 *  4. Autorize as permissões
 * ============================================================
 */

// ── CONFIGURAÇÃO GERAL ────────────────────────────────────────────────────

const ABAS_MONITORADAS = ['JAN2026', 'FEV2026', 'MAR2026', 'ABRI2026'];
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

  Logger.log('✅ Todos os gatilhos configurados!');
  SpreadsheetApp.getUi().alert(
    '✅ Automação v3 completa!\n\n' +
    '• 07h00 todo dia     → Dias em atraso + Score de Churn\n' +
    '• 08h00 todo dia     → Alerta de inadimplência (e-mail)\n' +
    '• Segunda às 07h     → Relatório semanal de KPIs (e-mail)\n' +
    '• Sexta às 18h       → Backup semanal no Drive (e-mail)\n' +
    '• A cada edição      → Registro na aba 📋 Log\n\n' +
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
    .addItem('🔄 Calcular dias em atraso agora',      'calcularDiasEmAtraso')
    .addItem('📊 Calcular score de churn agora',      'calcularScoreChurn')
    .addItem('⚠️  Enviar alerta inadimplência agora', 'enviarAlertaInadimplencia')
    .addItem('📧 Enviar relatório semanal agora',     'enviarRelatorioSemanal')
    .addItem('💾 Fazer backup agora',                 'backupSemanal')
    .addSeparator()
    .addItem('⚙️  Configurar automação (fazer uma vez)', 'configurarGatilhos')
    .addItem('🗑️  Remover automação',                 'removerGatilhos')
    .addToUi();
}

function normaliza(str) {
  return (str || '').toString().trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
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
