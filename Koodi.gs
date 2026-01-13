/***** CONFIG *****/
const DEFAULT_LANG = 'en';

/***** INSTALL TRIGGER *****/
function onInstall(e) {
  onOpen(e); // Simply calls the existing menu-creation logic
}

/***** EDITOR ADD-ON MENU *****/
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Open sidebar', 'openSidebar')
    .addToUi();
}

/***** THE OPEN FUNCTION *****/
function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Wikidata Reconciler')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/***** SETUP COLUMNS (INSERTION MODE) *****/
function setupColumns(config) {
  const ctx = getActiveContext();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(ctx.sheetName);
  
  // 1. Determine headers based on config
  const headers = ['Wikidata QID'];
  if (config.includeLabel) {
    config.langs.forEach(lang => headers.push(`Label (${lang})`));
  }
  if (config.includeDesc) {
    config.langs.forEach(lang => headers.push(`Description (${lang})`));
  }

  // 2. Insert the columns to prevent overwriting
  // This pushes everything from (ctx.col + 1) to the right
  sheet.insertColumnsAfter(ctx.col, headers.length);
  
// 3. Apply Headers to the newly created space
  const headerRange = sheet.getRange(1, ctx.col + 1, 1, headers.length);
  
  headerRange.setValues([headers])
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setBorder(true, true, true, true, true, true, '#ffffff', SpreadsheetApp.BorderStyle.SOLID);
             
  // Optional: Auto-resize the new columns so they aren't squished
  for (let i = 0; i < headers.length; i++) {
    sheet.autoResizeColumn(ctx.col + 1 + i);
  }

  return { success: true };
}

/***** GET ACTIVE CONTEXT *****/
function getActiveContext() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  if (!range) throw new Error('No active cell selected');
  
  return {
    sheetName: sheet.getName(),
    row: range.getRow(),
    col: range.getColumn(),
    value: range.getValue() // Direct fetch is fine for local sidebar use
  };
}

/***** SEARCH WIKIDATA *****/
function searchWikidata(query, langArray) {
  if (!query) return [];
  
  // Robust check: Use DEFAULT_LANG if langArray is missing or empty
  const primaryLang = (langArray && langArray.length > 0) ? langArray[0] : DEFAULT_LANG;
  
  const url = `https://www.wikidata.org/w/api.php?action=wbsearchentities&format=json&search=${encodeURIComponent(query)}&language=${primaryLang}&limit=10&type=item`;

  try {
    const res = UrlFetchApp.fetch(url, { headers: { 'User-Agent': 'GoogleSheets-Wikidata-Reconciler/1.1' } });
    const data = JSON.parse(res.getContentText());
    if ((!data.search || data.search.length === 0) && primaryLang !== 'en') {
      const fallbackUrl = url.replace(`language=${primaryLang}`, `language=en`) + `&uselang=${primaryLang}`;
      return JSON.parse(UrlFetchApp.fetch(fallbackUrl).getContentText()).search || [];
    }
    return data.search || [];
  } catch (e) {
    throw new Error("Search Failed: " + e.message);
  }
}

/***** APPLY ENTITY *****/
function applyEntity(qid, scope, config, ctx) {
  if (!ctx) ctx = getActiveContext();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(ctx.sheetName);
  
  // Use config.langs if it exists and has items, otherwise use DEFAULT_LANG in an array
  const langs = (config && config.langs && config.langs.length > 0) ? config.langs : [DEFAULT_LANG];
  const props = [];
  if (config.includeLabel) props.push('labels');
  if (config.includeDesc) props.push('descriptions');
  
  const entity = fetchEntity(qid, props, langs);

  let rowData = [qid];
  if (config.includeLabel) {
    langs.forEach(lang => rowData.push(entity.labels[lang] || ''));
  }
  if (config.includeDesc) {
    langs.forEach(lang => rowData.push(entity.descriptions[lang] || ''));
  }

  // --- 2. PRESERVE ORIGINAL WRITING LOGIC (ctx.col + 1) ---
  const rowsToUpdate = getTargetRows(ctx, scope);
  rowsToUpdate.forEach(rowNum => {
    // Writes next to the search term
    sheet.getRange(rowNum, ctx.col + 1, 1, rowData.length).setValues([rowData]);
  });

  // --- 3. ADDED: SMART JUMP LOGIC ---
  // We only jump if doing a single-cell match to maintain workflow focus
  if (scope === 'SINGLE_CELL') {
    const lastRow = sheet.getLastRow();
    // We get the column of data to see what is already filled
    const dataRange = sheet.getRange(1, ctx.col, lastRow, 2).getValues(); 
    
    let nextRow = ctx.row + 1;
    let found = false;

    for (let i = nextRow - 1; i < lastRow; i++) {
      const sourceVal = dataRange[i][0]; // Original search term
      const qidVal = dataRange[i][1];    // The QID column (ctx.col + 1)

      if (sourceVal && !qidVal) {
        nextRow = i + 1;
        found = true;
        break;
      }
    }

    if (found) {
      sheet.getRange(nextRow, ctx.col).activate();
    }
  }
  return { success: true, rowsUpdated: rowsToUpdate.length };
}

/***** FETCH ENTITY DETAILS *****/
function fetchEntity(qid, props, langs) {
  const langParam = encodeURIComponent(langs.join('|'));
  const propsParam = encodeURIComponent(props.join('|'));
  const url = `https://www.wikidata.org/w/api.php?action=wbgetentities&format=json&formatversion=2&ids=${encodeURIComponent(qid)}&props=${propsParam}&languages=${langParam}`;
  const res = UrlFetchApp.fetch(url);
  const data = JSON.parse(res.getContentText());
  const e = data.entities[qid];
  const entity = { labels: {}, descriptions: {} };
  langs.forEach(lang => {
    if (props.includes('labels')) entity.labels[lang] = e.labels?.[lang]?.value || '';
    if (props.includes('descriptions')) entity.descriptions[lang] = e.descriptions?.[lang]?.value || '';
  });
  return entity;
}

/***** UTILITIES *****/
function normalize(value) {
  return String(value || '').trim().toLowerCase();
}

function getTargetRows(ctx, scope) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(ctx.sheetName);
  if (scope === 'SINGLE_CELL') return [ctx.row];
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  const values = sheet.getRange(1, ctx.col, lastRow).getValues();
  const target = normalize(ctx.value);
  return values.map((v, i) => ({ value: v[0], row: i + 1 })).filter(v => normalize(v.value) === target).map(v => v.row);
}
