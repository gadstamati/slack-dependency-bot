/**
 * ==============================================================================
 * PEDIDOSYA - AUTOMATIZACIÓN DE AUDITORÍA DE INICIATIVAS & DEPENDENCIAS
 * Bot de Slack para Notificación de Faltantes de Documentación (DoD y PRD)
 * ==============================================================================
 */

// ------------------------------------------------------------------------------
// CONFIGURACIÓN GLOBAL
// ------------------------------------------------------------------------------
const CONFIG = {
  SPREADSHEET_ID: '10idj90beOyWMzI2AIwATImwUl3aI0-rggzvvoeimv4s',
  ORGANIGRAMA_SHEET_NAME: 'Organigrama + Impacto',
  
  // Pestañas que NUNCA deben auditarse como dominios
  EXCLUDED_SHEETS: [
    'Notas Q1 27',
    'Status Iniciativas',
    'Proceso',
    'Organigrama + Impacto',
    'Dependencias Dinámico'
  ],

  // Mapeo de columnas en las pestañas de dominios (Base 0 para arrays)
  COLUMNS: {
    TRIBU: 0,        // Col A
    DOMINIO: 1,      // Col B
    SQUAD: 2,        // Col C
    INICIATIVA: 3,   // Col D
    TIPO: 4,         // Col E (Iniciativa propia / Dependencia)
    DOD: 6,          // Col G (Descripcion corta + Definition of Done)
    LINKS: 7         // Col H (Links PRD/PPT)
  },

  HEADER_ROW: 3,     // Fila de encabezados
  DATA_START_ROW: 4, // Primera fila con datos

  /**
   * MODO DE PRUEBA (DRY RUN)
   * true: No envía DMs reales a los PMs. Imprime los mensajes en el Registro de Ejecución (Logs).
   * false: Envía los DMs a cada PM en Slack usando chat.postMessage.
   */
  DRY_RUN: true,
  TEST_SLACK_EMAIL: 'gad.stamati@pedidosya.com', // Recibe prueba si DRY_RUN está activo y hay token

  // Overrides de email en caso de excepciones o nombres sin chip
  EMAIL_OVERRIDES: {
    'gad stamati': 'gad.stamati@pedidosya.com',
    'nicolás de león': 'nicolas.deleon@pedidosya.com',
    'nicolas de leon': 'nicolas.deleon@pedidosya.com',
    'julián izrael': 'julian.izrael@pedidosya.com',
    'julian izrael': 'julian.izrael@pedidosya.com',
    'manuel pintos': 'manuel.pintos@pedidosya.com',
    'ionatan stamati': 'ionatan.stamati@pedidosya.com',
    'carolina iölster': 'carolina.iolster@pedidosya.com'
  }
};

// ------------------------------------------------------------------------------
// FUNCIÓN PRINCIPAL DE AUDITORÍA
// ------------------------------------------------------------------------------
function runDocumentationAudit() {
  console.log('--- Iniciando Auditoría de Documentación de Iniciativas ---');
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const slackToken = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  
  if (!slackToken) {
    if (!CONFIG.DRY_RUN) {
      throw new Error('Error: No se encontró SLACK_BOT_TOKEN en las Propiedades de la secuencia de comandos.');
    } else {
      console.warn('[MODO SIMULACIÓN] Sin token de Slack. Los mensajes se imprimirán únicamente en los Logs.');
    }
  }

  // 1. Obtener la matriz de Squad -> PM (con emails nativos de Person Chips)
  const squadToPmMap = loadSquadToPmMap(ss);
  console.log(`Matriz cargada: ${Object.keys(squadToPmMap).length} squads mapeados a PMs.`);

  // 2. Primera pasada: recolectar todas las filas con iniciativa y detectar candidatos a Smart Chip
  const sheets = ss.getSheets();
  const rawInitiatives = []; // Lista plana de iniciativas a evaluar
  const smartChipCandidates = []; // Referencias de celdas que podrían ser Smart Chips

  sheets.forEach(sheet => {
    const sheetName = sheet.getName().trim();
    if (CONFIG.EXCLUDED_SHEETS.includes(sheetName) || sheet.isSheetHidden()) {
      return;
    }

    // Verificar si la hoja tiene el encabezado de iniciativas en la celda D3
    const headerCheck = sheet.getRange(CONFIG.HEADER_ROW, CONFIG.COLUMNS.INICIATIVA + 1).getValue().toString().trim();
    if (!headerCheck.toLowerCase().includes('iniciativa')) {
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.DATA_START_ROW) return;

    const rowCount = lastRow - CONFIG.DATA_START_ROW + 1;
    const range = sheet.getRange(CONFIG.DATA_START_ROW, 1, rowCount, 8);
    
    const values = range.getValues();
    const richTexts = range.getRichTextValues();
    const formulas = range.getFormulas();

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const iniciativa = String(row[CONFIG.COLUMNS.INICIATIVA] || '').trim();
      if (!iniciativa) continue;

      const currentRowNum = CONFIG.DATA_START_ROW + i;
      const squadRaw = String(row[CONFIG.COLUMNS.SQUAD] || '').trim();
      const tribu = String(row[CONFIG.COLUMNS.TRIBU] || '').trim();
      const dominio = String(row[CONFIG.COLUMNS.DOMINIO] || '').trim();

      const dodText = String(row[CONFIG.COLUMNS.DOD] || '').trim();
      const linkCell = row[CONFIG.COLUMNS.LINKS];
      const richTextCell = richTexts[i][CONFIG.COLUMNS.LINKS];
      const formulaCell = formulas[i][CONFIG.COLUMNS.LINKS];

      // Verificación directa (URL plana, hipervínculo tradicional o fórmula =HYPERLINK)
      const directLink = extractDirectLink(linkCell, richTextCell, formulaCell);

      const item = {
        sheetName: sheetName,
        rowNum: currentRowNum,
        iniciativa: iniciativa,
        squad: squadRaw,
        tribu: tribu,
        dominio: dominio,
        dod: dodText,
        link: directLink,
        needsSmartChipCheck: false
      };

      // Si no tiene enlace directo pero la celda no está vacía, puede ser un Smart Chip de Google Drive/Docs
      if (!directLink && linkCell && String(linkCell).trim().length > 0) {
        item.needsSmartChipCheck = true;
        smartChipCandidates.push({
          item: item,
          sheetName: sheetName,
          rowNum: currentRowNum
        });
      }

      rawInitiatives.push(item);
    }
  });

  // 3. Resolución nativa de Smart Chips (utilizando fórmulas internas de Google Sheets, 100% libre de APIs de Cloud)
  if (smartChipCandidates.length > 0) {
    resolveSmartChipsNatively(ss, smartChipCandidates);
  }

  // 4. Agrupar iniciativas con faltantes por PM
  const issuesByPm = {};

  rawInitiatives.forEach(item => {
    const hasValidDod = isValidDoD(item.dod);
    const hasValidLink = Boolean(item.link && (item.link.startsWith('http://') || item.link.startsWith('https://')));

    if (!hasValidDod || !hasValidLink) {
      let missingReason = '';
      if (!hasValidDod && !hasValidLink) {
        missingReason = ':red_circle: Falta Link a PRD/PPT y falta Descripción/DoD';
      } else if (!hasValidLink) {
        missingReason = ':red_circle: Falta Link a PRD/PPT (la celda está vacía o sin URL válida)';
      } else {
        missingReason = ':red_circle: Falta Descripción/DoD';
      }

      const squadKey = normalizeText(item.squad);
      const pmInfo = squadToPmMap[squadKey] || {
        pmName: 'Sin Asignar (TBD)',
        pmEmail: null,
        tribu: item.tribu,
        dominio: item.dominio
      };

      const pmGroupKey = pmInfo.pmName.toLowerCase();
      if (!issuesByPm[pmGroupKey]) {
        issuesByPm[pmGroupKey] = {
          pmName: pmInfo.pmName,
          pmEmail: pmInfo.pmEmail,
          tribu: item.tribu || pmInfo.tribu || 'Verticals',
          dominio: item.dominio || pmInfo.dominio || item.sheetName,
          issues: []
        };
      }

      issuesByPm[pmGroupKey].issues.push({
        iniciativa: item.iniciativa,
        squad: item.squad || pmInfo.squad || 'Sin Squad',
        missingReason: missingReason
      });
    }
  });

  // 5. Notificar a cada PM por Slack (o imprimir en Logs)
  const pmKeys = Object.keys(issuesByPm);
  console.log(`\n==================================================================`);
  console.log(`RESUMEN DE AUDITORÍA: Se detectaron ${pmKeys.length} PM(s) con iniciativas con faltantes.`);
  console.log(`==================================================================\n`);

  if (pmKeys.length === 0) {
    console.log('🎉 ¡Excelente! No se detectaron iniciativas con faltantes en ninguna pestaña.');
    return;
  }

  const emailToSlackIdCache = {};

  pmKeys.forEach(pmKey => {
    const pmData = issuesByPm[pmKey];
    const message = buildSlackMessage(pmData);

    console.log(`------------------------------------------------------------------`);
    console.log(`📢 MENSAJE PARA: ${pmData.pmName} (${pmData.pmEmail || 'Sin email'})`);
    console.log(`Tribu: ${pmData.tribu} | Dominio: ${pmData.dominio} | Faltantes: ${pmData.issues.length}`);
    console.log(`------------------------------------------------------------------`);
    console.log(message);
    console.log(`\n`);

    // Envío a Slack si no es DRY_RUN y hay token
    if (!CONFIG.DRY_RUN && slackToken && pmData.pmEmail) {
      sendSlackDMByEmail(pmData.pmEmail, message, slackToken, emailToSlackIdCache);
      Utilities.sleep(600); // Pausa de cortesía para respetar Rate Limits de Slack
    }
  });

  console.log('=== Fin de la ejecución ===');
}

// ------------------------------------------------------------------------------
// CONSTRUCCIÓN DEL MENSAJE DE SLACK
// ------------------------------------------------------------------------------
function buildSlackMessage(pmData) {
  const count = pmData.issues.length;
  let text = `:warning: *Auditoría de Documentación: Iniciativas con faltantes*\n`;
  text += `*Tribu:* ${pmData.tribu} | *Dominio:* ${pmData.dominio}\n\n`;
  text += `Hola! Se detectaron ${count} iniciativa(s) propia(s) o asignadas que aún no cuentan con link a PRD/Documentación o descripción completa:\n\n`;

  pmData.issues.forEach(issue => {
    text += `• *${issue.iniciativa}* (Squad: ${issue.squad}) ${issue.missingReason}\n`;
  });

  text += `\n:point_right: *Acción requerida:* Por favor actualicen los enlaces en la columna *Links (PRD/PPT)* y las descripciones correspondientes para asegurar el freeze.\n`;
  text += `<https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/edit|Ver Planilla de Dependencias>`;

  return text;
}

// ------------------------------------------------------------------------------
// ENVÍO DE SLACK DM VIA API (chat.postMessage + users.lookupByEmail)
// ------------------------------------------------------------------------------
function sendSlackDMByEmail(email, text, slackToken, cache) {
  if (!slackToken) return;

  try {
    let slackUserId = cache[email];
    
    // Si no está en caché, buscarlo en Slack
    if (!slackUserId) {
      const lookupUrl = `https://slack.com/api/users.lookupByEmail?email=${encodeURIComponent(email)}`;
      const lookupResponse = UrlFetchApp.fetch(lookupUrl, {
        method: 'get',
        headers: { 'Authorization': `Bearer ${slackToken}` },
        muteHttpExceptions: true
      });
      
      const lookupData = JSON.parse(lookupResponse.getContentText());
      if (!lookupData.ok) {
        console.error(`Slack Lookup Error para ${email}: ${lookupData.error}`);
        return;
      }
      slackUserId = lookupData.user.id;
      cache[email] = slackUserId;
    }

    // Enviar el mensaje al DM del usuario (channel = slackUserId)
    const postUrl = 'https://slack.com/api/chat.postMessage';
    const payload = {
      channel: slackUserId,
      text: text,
      mrkdwn: true
    };

    const postResponse = UrlFetchApp.fetch(postUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': `Bearer ${slackToken}` },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const postData = JSON.parse(postResponse.getContentText());
    if (!postData.ok) {
      console.error(`Slack Post Error para ${email}: ${postData.error}`);
    }
  } catch (err) {
    console.error(`Excepción enviando DM a ${email}: ${err.message}`);
  }
}

// ------------------------------------------------------------------------------
// CARGA DE LA MATRIZ SQUAD -> PM (100% NATIVA)
// ------------------------------------------------------------------------------
function loadSquadToPmMap(ss) {
  const sheet = ss.getSheetByName(CONFIG.ORGANIGRAMA_SHEET_NAME);
  if (!sheet) {
    throw new Error(`No se encontró la pestaña "${CONFIG.ORGANIGRAMA_SHEET_NAME}"`);
  }

  const lastRow = sheet.getLastRow();
  const rowCount = lastRow - 3;
  // Leer columnas B (Tribu), C (Dominio), D (Squad), E (PM)
  const range = sheet.getRange(4, 2, rowCount, 4);
  const data = range.getValues();

  // Extraer emails nativamente de los Person chips usando fórmulas dot notation en memoria
  const chipEmails = extractPersonChipEmailsNatively(ss, rowCount);

  const squadMap = {};
  let currentTribu = '';
  let currentDominio = '';

  data.forEach((row, idx) => {
    const tribu = String(row[0] || '').trim();
    const dominio = String(row[1] || '').trim();
    const squad = String(row[2] || '').trim();
    const pm = String(row[3] || '').trim();

    if (tribu) currentTribu = tribu;
    if (dominio) currentDominio = dominio;

    if (squad && squad !== 'Domain Leads' && squad !== 'Tribe Leads') {
      const normalizedSquad = normalizeText(squad);
      const emailFromChip = chipEmails[idx] || null;

      squadMap[normalizedSquad] = {
        squad: squad,
        tribu: currentTribu,
        dominio: currentDominio,
        pmName: pm,
        pmEmail: emailFromChip || resolvePmEmail(pm)
      };
    }
  });

  return squadMap;
}

// ------------------------------------------------------------------------------
// RESOLUCIÓN NATIVA DE SMART CHIPS & PERSON CHIPS (SIN GOOGLE CLOUD APIS)
// ------------------------------------------------------------------------------

/**
 * Resuelve las URLs de los Smart Chips de Drive/Docs utilizando la sintaxis de punto nativa (=H*.url)
 * mediante una hoja temporal de cálculo interna que se crea y destruye en <0.5 seg.
 */
function resolveSmartChipsNatively(ss, candidates) {
  if (!candidates || candidates.length === 0) return;

  const tempSheetName = '__audit_eval_temp__';
  let tempSheet = ss.getSheetByName(tempSheetName);
  if (tempSheet) ss.deleteSheet(tempSheet);

  tempSheet = ss.insertSheet(tempSheetName);
  tempSheet.hideSheet();

  try {
    const formulas = candidates.map(c => [`=IFERROR('${c.sheetName}'!H${c.rowNum}.url, "")`]);
    tempSheet.getRange(1, 1, formulas.length, 1).setFormulas(formulas);
    SpreadsheetApp.flush();

    const results = tempSheet.getRange(1, 1, formulas.length, 1).getValues();

    candidates.forEach((c, idx) => {
      const resolvedUrl = results[idx] ? String(results[idx][0] || '').trim() : '';
      if (resolvedUrl && (resolvedUrl.startsWith('http://') || resolvedUrl.startsWith('https://'))) {
        c.item.link = resolvedUrl;
      }
    });
  } catch (err) {
    console.warn(`[SmartChip] Error en cálculo de Smart Chips: ${err.message}`);
  } finally {
    try {
      ss.deleteSheet(tempSheet);
    } catch (e) {}
  }
}

/**
 * Extrae los emails de los Person chips en Organigrama usando dot notation (=E*.email)
 */
function extractPersonChipEmailsNatively(ss, rowCount) {
  const tempSheetName = '__email_eval_temp__';
  let tempSheet = ss.getSheetByName(tempSheetName);
  if (tempSheet) ss.deleteSheet(tempSheet);

  tempSheet = ss.insertSheet(tempSheetName);
  tempSheet.hideSheet();

  const emails = [];
  try {
    const formulas = [];
    for (let i = 0; i < rowCount; i++) {
      const rowNum = 4 + i;
      formulas.push([`=IFERROR('${CONFIG.ORGANIGRAMA_SHEET_NAME}'!E${rowNum}.email, "")`]);
    }
    tempSheet.getRange(1, 1, formulas.length, 1).setFormulas(formulas);
    SpreadsheetApp.flush();

    const results = tempSheet.getRange(1, 1, formulas.length, 1).getValues();
    results.forEach(r => {
      const email = r ? String(r[0] || '').trim() : '';
      emails.push(email || null);
    });
  } catch (err) {
    console.warn(`[PersonChip] Error resolviendo emails: ${err.message}`);
  } finally {
    try {
      ss.deleteSheet(tempSheet);
    } catch (e) {}
  }

  return emails;
}

// ------------------------------------------------------------------------------
// RESOLUCIÓN DE EMAIL DE PM (FALLBACK)
// ------------------------------------------------------------------------------
function resolvePmEmail(pmName) {
  if (!pmName || pmName.includes('TBD') || pmName === '-') return null;

  const cleanName = pmName.toLowerCase().trim();
  if (CONFIG.EMAIL_OVERRIDES[cleanName]) {
    return CONFIG.EMAIL_OVERRIDES[cleanName];
  }

  if (pmName.includes('/')) {
    const firstPm = pmName.split('/')[0].trim();
    return resolvePmEmail(firstPm);
  }

  const normalized = cleanName
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  const parts = normalized.split(' ');
  if (parts.length >= 2) {
    const firstName = parts[0];
    const lastName = parts[parts.length - 1];
    return `${firstName}.${lastName}@pedidosya.com`;
  }

  return `${normalized}@pedidosya.com`;
}

// ------------------------------------------------------------------------------
// HELPERS DE VALIDACIÓN
// ------------------------------------------------------------------------------
function isValidDoD(text) {
  if (!text) return false;
  const clean = text.replace(/[\s\-_.]/g, '');
  return clean.length >= 5; // Mínimo 5 caracteres significativos
}

function extractDirectLink(cellValue, richTextCell, formula) {
  // 1. Verificar fórmula =HYPERLINK("url", ...)
  if (formula && typeof formula === 'string' && formula.toUpperCase().includes('HYPERLINK')) {
    const match = formula.match(/HYPERLINK\(\s*["']([^"']+)["']/i);
    if (match && (match[1].startsWith('http://') || match[1].startsWith('https://'))) {
      return match[1].trim();
    }
  }

  // 2. Verificar RichText (enlace tradicional incrustado)
  if (richTextCell && typeof richTextCell.getLinkUrl === 'function') {
    const url = richTextCell.getLinkUrl();
    if (url && (url.startsWith('http://') || url.startsWith('https://'))) {
      return url.trim();
    }
  }

  // 3. Verificar texto plano que sea una URL directa
  if (cellValue && typeof cellValue === 'string') {
    const trimmed = cellValue.trim();
    if (trimmed.startsWith('http://') || trimmed.startsWith('https://')) {
      return trimmed;
    }
  }

  return null;
}

function normalizeText(text) {
  return String(text || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// ------------------------------------------------------------------------------
// GESTIÓN DE TRIGGERS AUTOMÁTICOS (LUNES Y JUEVES 9:00 AM)
// ------------------------------------------------------------------------------
function installWeeklyTriggers() {
  console.log('Instalando disparadores programados para Lunes y Jueves a las 9:00 AM...');
  
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runDocumentationAudit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('runDocumentationAudit')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .nearMinute(0)
    .create();

  ScriptApp.newTrigger('runDocumentationAudit')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.THURSDAY)
    .atHour(9)
    .nearMinute(0)
    .create();

  console.log('Disparadores programados correctamente para Lunes y Jueves 9:00 AM.');
}
