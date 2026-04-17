/**
 * ============================================================
 *  IBNET TELECOM — Automação Google Sheets
 *  Script: Cálculo automático de "Dias em atraso"
 * ============================================================
 *
 *  INSTALAÇÃO (faça uma vez):
 *  1. Abra a planilha no Google Sheets
 *  2. Menu: Extensões → Apps Script
 *  3. Apague o conteúdo padrão e cole TODO este arquivo
 *  4. Salve (Ctrl+S) e dê um nome ao projeto (ex: "IBnet Automação")
 *  5. Rode configurarGatilhos() UMA VEZ (botão ▶ com essa função selecionada)
 *     → isso cria o agendamento diário automático
 *  6. Autorize as permissões quando solicitado
 *
 *  COLUNAS ESPERADAS EM CADA ABA MENSAL:
 *  - "Data de Vencimento"   → data no formato DD/MM/AAAA
 *  - "Status do Pagamento"  → Pago | Pendente | Em atraso | Isento | Cancelado
 *  - "Dias em atraso"       → coluna que será PREENCHIDA automaticamente
 * ============================================================
 */

// ── CONFIGURAÇÃO ──────────────────────────────────────────────────────────

const ABAS_MONITORADAS = ['JAN2026', 'FEV2026', 'MAR2026', 'ABRI2026'];

// Status que NÃO devem ter dias em atraso calculados (zera o campo)
const STATUS_IGNORAR = ['Pago', 'Isento', 'Cancelado'];

// Quantos dias de atraso para mudar status automaticamente para "Em atraso"
const DIAS_PARA_ATRASO = 3;


// ── FUNÇÃO PRINCIPAL ──────────────────────────────────────────────────────

/**
 * Percorre todas as abas monitoradas e atualiza "Dias em atraso"
 * com base em HOJE() - "Data de Vencimento".
 * Também atualiza o Status para "Em atraso" se ultrapassar o limite.
 */
function calcularDiasEmAtraso() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoje  = new Date();
  hoje.setHours(0, 0, 0, 0); // normaliza para meia-noite

  let totalAtualizados = 0;
  const resumo = [];

  ABAS_MONITORADAS.forEach(nomeAba => {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) {
      Logger.log(`⚠️  Aba "${nomeAba}" não encontrada — pulando.`);
      return;
    }

    const dados     = aba.getDataRange().getValues();
    const cabecalho = dados[0];

    // Localiza as colunas pelo nome do cabeçalho (tolerante a espaços extras)
    const iVenc    = cabecalho.findIndex(h => normaliza(h) === 'data de vencimento');
    const iStatus  = cabecalho.findIndex(h => normaliza(h) === 'status do pagamento');
    const iDias    = cabecalho.findIndex(h => normaliza(h) === 'dias em atraso');

    if (iVenc === -1 || iStatus === -1 || iDias === -1) {
      Logger.log(`⚠️  Aba "${nomeAba}": colunas não encontradas.`);
      Logger.log(`   Cabeçalho detectado: ${JSON.stringify(cabecalho)}`);
      return;
    }

    let atualizadosAba = 0;
    let atrasadosAba   = 0;

    for (let i = 1; i < dados.length; i++) {
      const linha  = dados[i];
      const status = (linha[iStatus] || '').toString().trim();
      const vencRaw = linha[iVenc];

      // Pula linhas sem data de vencimento
      if (!vencRaw) continue;

      // Se status é Pago/Isento/Cancelado → zera dias em atraso
      if (STATUS_IGNORAR.includes(status)) {
        if (linha[iDias] !== 0 && linha[iDias] !== '') {
          aba.getRange(i + 1, iDias + 1).setValue(0);
          atualizadosAba++;
        }
        continue;
      }

      // Converte data de vencimento
      const dataVenc = parsearData(vencRaw);
      if (!dataVenc) continue;

      // Calcula dias em atraso
      const diffMs  = hoje.getTime() - dataVenc.getTime();
      const dias    = Math.floor(diffMs / (1000 * 60 * 60 * 24));
      const diasFinal = dias > 0 ? dias : 0;

      // Atualiza "Dias em atraso" se mudou
      if (linha[iDias] !== diasFinal) {
        aba.getRange(i + 1, iDias + 1).setValue(diasFinal);
        atualizadosAba++;
      }

      // Atualiza status para "Em atraso" se necessário
      if (diasFinal >= DIAS_PARA_ATRASO && status === 'Pendente') {
        aba.getRange(i + 1, iStatus + 1).setValue('Em atraso');
        atualizadosAba++;
      }

      if (diasFinal > 0) atrasadosAba++;
    }

    totalAtualizados += atualizadosAba;
    resumo.push(`${nomeAba}: ${atrasadosAba} cliente(s) em atraso, ${atualizadosAba} célula(s) atualizada(s)`);
    Logger.log(`✅ ${nomeAba}: ${atrasadosAba} em atraso | ${atualizadosAba} atualizados`);
  });

  Logger.log(`\n📊 RESUMO — ${new Date().toLocaleString('pt-BR')}`);
  resumo.forEach(r => Logger.log(`   ${r}`));
  Logger.log(`   Total de atualizações: ${totalAtualizados}`);

  return { totalAtualizados, resumo };
}


// ── GATILHOS (AGENDAMENTO) ────────────────────────────────────────────────

/**
 * Cria os gatilhos automáticos.
 * Execute ESTA FUNÇÃO uma única vez para configurar o agendamento.
 */
function configurarGatilhos() {
  // Remove gatilhos antigos deste script para evitar duplicatas
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'calcularDiasEmAtraso') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Gatilho 1: Todo dia às 07:00
  ScriptApp.newTrigger('calcularDiasEmAtraso')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  // Gatilho 2: Toda vez que a planilha é aberta
  ScriptApp.newTrigger('calcularDiasEmAtraso')
    .spreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();

  Logger.log('✅ Gatilhos configurados com sucesso!');
  Logger.log('   → Execução automática: diariamente às 07h + toda vez que abrir a planilha');

  SpreadsheetApp.getUi().alert(
    '✅ Automação configurada!\n\n' +
    '• Cálculo automático todo dia às 07h00\n' +
    '• Cálculo automático ao abrir a planilha\n\n' +
    'A coluna "Dias em atraso" será mantida sempre atualizada.'
  );
}

/**
 * Remove todos os gatilhos (use se quiser desativar a automação).
 */
function removerGatilhos() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('🗑️  Todos os gatilhos removidos.');
}


// ── UTILITÁRIOS ───────────────────────────────────────────────────────────

/** Normaliza string para comparação de cabeçalhos */
function normaliza(str) {
  return (str || '').toString().trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, ''); // remove acentos
}

/**
 * Converte vários formatos de data para objeto Date.
 * Suporta: objetos Date nativos do Sheets, strings DD/MM/AAAA, AAAA-MM-DD.
 */
function parsearData(valor) {
  if (!valor) return null;

  // Já é um Date (formato nativo do Sheets)
  if (valor instanceof Date) {
    const d = new Date(valor);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  const str = valor.toString().trim();

  // DD/MM/AAAA
  const m1 = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m1) {
    return new Date(parseInt(m1[3]), parseInt(m1[2]) - 1, parseInt(m1[1]));
  }

  // AAAA-MM-DD
  const m2 = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2) {
    return new Date(parseInt(m2[1]), parseInt(m2[2]) - 1, parseInt(m2[3]));
  }

  return null;
}


// ── MENU PERSONALIZADO ────────────────────────────────────────────────────

/**
 * Adiciona menu "IBnet" na barra do Sheets para rodar manualmente.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ IBnet')
    .addItem('Calcular dias em atraso agora', 'calcularDiasEmAtraso')
    .addSeparator()
    .addItem('Configurar automação (fazer uma vez)', 'configurarGatilhos')
    .addItem('Remover automação', 'removerGatilhos')
    .addToUi();
}
