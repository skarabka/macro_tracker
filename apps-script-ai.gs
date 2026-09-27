const SPREADSHEET_ID = '1-o6k-9UKY5vxI4Swmt6K4Ugj1ifE28-KWEU-DfdXQKc';
const PHOTO_FOLDER_ID = '1Pj5dm-XAkxxyXABj0ZdD4Qvu_5PoS5xv';

function doGet() {
  return json_({ok:true, service:'bzhv-tracker', version:2});
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = body.action || '';
    if (action === 'analyze_photo') return analyzePhoto_(body.imageDataUrl || '', body.mode || 'photo');
    if (action === 'analyze_manual') return analyzeManual_(body.meal_name || '', body.weight_g || 0);
    if (action === 'add_food') return addFood_(body.entry || {});
    if (action === 'save_settings') return saveSettings_(body.settings || {});
    if (action === 'list_food') return json_({ok:true, entries:listFood_()});
    if (action === 'get_settings') return json_({ok:true, settings:getSettings_()});
    if (action === 'bootstrap') return json_({ok:true, entries:listFood_(), settings:getSettings_()});
    return json_({ok:false,error:'Unknown action'});
  } catch (err) {
    return json_({ok:false,error:String(err && err.message ? err.message : err)});
  }
}

function addFood_(x) {
  if (!x.id) x.id = Utilities.getUuid();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName('food_log');
    if (!sh) throw new Error('Sheet food_log not found');

    // Idempotency: repeated offline sync must not duplicate a meal.
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const ids = sh.getRange(2,1,lastRow-1,1).getDisplayValues().flat();
      const found = ids.indexOf(String(x.id));
      if (found >= 0) {
        return json_({ok:true, duplicate:true, image_url:sh.getRange(found+2,10).getDisplayValue() || ''});
      }
    }

    let imageUrl = '';
    if (x.imageDataUrl) imageUrl = saveImage_(x.imageDataUrl, x.id);
    const dt = x.datetime ? new Date(x.datetime) : new Date();
    sh.appendRow([
      String(x.id), dt, String(x.meal_name || ''), String(x.input_type || 'manual'),
      Number(x.weight_g)||0, Number(x.kcal)||0, Number(x.protein)||0, Number(x.fat)||0, Number(x.carbs)||0,
      imageUrl, String(x.comment || '')
    ]);
    rebuildSummary_();
    return json_({ok:true,id:String(x.id),image_url:imageUrl});
  } finally { lock.releaseLock(); }
}

function saveSettings_(s) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('settings');
    if (!sh) throw new Error('Sheet settings not found');
    const map = {kcal_target:Number(s.kcal)||0,protein_target:Number(s.protein)||0,fat_target:Number(s.fat)||0,carbs_target:Number(s.carbs)||0};
    const rows = Math.max(sh.getLastRow()-1,4);
    const values = sh.getRange(2,1,rows,2).getValues();
    values.forEach((row,i)=>{ if(Object.prototype.hasOwnProperty.call(map,row[0])) sh.getRange(i+2,2).setValue(map[row[0]]); });
    return json_({ok:true,settings:getSettings_()});
  } finally { lock.releaseLock(); }
}

function getSettings_() {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('settings');
  if (!sh || sh.getLastRow()<2) return {kcal:0,protein:0,fat:0,carbs:0};
  const vals=sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
  const raw=Object.fromEntries(vals.filter(r=>r[0]).map(r=>[String(r[0]),Number(r[1])||0]));
  return {kcal:raw.kcal_target||0,protein:raw.protein_target||0,fat:raw.fat_target||0,carbs:raw.carbs_target||0};
}

function listFood_() {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('food_log');
  if (!sh || sh.getLastRow()<2) return [];
  const values = sh.getRange(2,1,sh.getLastRow()-1,11).getValues();
  const tz = Session.getScriptTimeZone();
  return values.filter(r=>r[0]).map(r=>({
    id:String(r[0]),
    datetime:r[1] instanceof Date ? r[1].toISOString() : String(r[1]||''),
    date:r[1] ? Utilities.formatDate(new Date(r[1]),tz,'yyyy-MM-dd') : '',
    meal_name:String(r[2]||''), input_type:String(r[3]||'manual'), weight_g:Number(r[4])||0,
    kcal:Number(r[5])||0, protein:Number(r[6])||0, fat:Number(r[7])||0, carbs:Number(r[8])||0,
    image_url:String(r[9]||''), comment:String(r[10]||''), sync_status:'synced'
  }));
}

function rebuildSummary_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const src = ss.getSheetByName('food_log');
  const dst = ss.getSheetByName('daily_summary');
  const rows = src.getLastRow()>1 ? src.getRange(2,1,src.getLastRow()-1,11).getValues() : [];
  const out = {};
  rows.forEach(r=>{
    if(!r[1]) return;
    const d = Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!out[d]) out[d] = [d,0,0,0,0];
    out[d][1]+=Number(r[5])||0; out[d][2]+=Number(r[6])||0; out[d][3]+=Number(r[7])||0; out[d][4]+=Number(r[8])||0;
  });
  if(dst.getLastRow()>1) dst.getRange(2,1,dst.getLastRow()-1,5).clearContent();
  const vals=Object.values(out).sort((a,b)=>String(a[0]).localeCompare(String(b[0])));
  if(vals.length) dst.getRange(2,1,vals.length,5).setValues(vals);
}

function saveImage_(dataUrl, id) {
  const m = String(dataUrl).match(/^data:(image\/[^;]+);base64,(.+)$/);
  if(!m) return '';
  const bytes=Utilities.base64Decode(m[2]);
  if(bytes.length > 8*1024*1024) throw new Error('Image is too large after compression');
  const ext = m[1].split('/')[1].replace('jpeg','jpg');
  const blob = Utilities.newBlob(bytes, m[1], `${id}.${ext}`);
  return DriveApp.getFolderById(PHOTO_FOLDER_ID).createFile(blob).getUrl();
}

function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


function analyzePhoto_(imageDataUrl, mode) {
  const key = PropertiesService.getScriptProperties().getProperty('MACRO_TRACKER_API_KEY');
  if (!key) throw new Error('MACRO_TRACKER_API_KEY is not configured in Script Properties');

  imageDataUrl = String(imageDataUrl || '');
  mode = String(mode || 'photo');
  if (!/^data:image\/(jpeg|jpg|png|webp);base64,/i.test(imageDataUrl)) {
    throw new Error('Invalid image format');
  }
  if (imageDataUrl.length > 8000000) {
    throw new Error('Image is too large for AI analysis');
  }

  enforceAiDailyLimit_();

  const photoInstruction = mode === 'label'
    ? 'Read the nutrition label carefully. If values and serving size are clearly visible, return nutrition for one stated serving. If only per-100-g values are clear, set weight_g to 100 and return those values. Use the visible product name when possible. Do not invent unreadable numbers.'
    : 'Analyze the visible meal. Estimate the total visible portion weight and nutrition for the whole visible portion. Account for likely cooking oil or sauce only when visually plausible. Do not pretend exact precision and do not invent ingredients that are not reasonably supported by the image.';

  const schema = {
    type: 'object',
    additionalProperties: false,
    properties: {
      meal_name: { type: 'string' },
      weight_g: { type: 'number', minimum: 0 },
      kcal: { type: 'number', minimum: 0 },
      protein_g: { type: 'number', minimum: 0 },
      fat_g: { type: 'number', minimum: 0 },
      carbs_g: { type: 'number', minimum: 0 },
      confidence: { type: 'number', minimum: 0, maximum: 1 },
      assumptions: { type: 'string' }
    },
    required: ['meal_name','weight_g','kcal','protein_g','fat_g','carbs_g','confidence','assumptions']
  };

  const payload = {
    model: 'gpt-5.6-terra',
    store: false,
    input: [
      {
        role: 'developer',
        content: [{
          type: 'input_text',
          text: 'You estimate food nutrition from images for a personal macro tracker. Return a practical estimate, not a medical claim. Use Ukrainian for meal_name and assumptions. Keep assumptions concise. Calories and macros must describe the same portion and be internally plausible.'
        }]
      },
      {
        role: 'user',
        content: [
          { type: 'input_text', text: photoInstruction },
          { type: 'input_image', image_url: imageDataUrl, detail: 'high' }
        ]
      }
    ],
    text: {
      format: {
        type: 'json_schema',
        name: 'nutrition_estimate',
        strict: true,
        schema: schema
      }
    }
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const raw = response.getContentText();
  if (status < 200 || status >= 300) {
    let detail = raw;
    try {
      const parsedError = JSON.parse(raw);
      detail = parsedError && parsedError.error && parsedError.error.message ? parsedError.error.message : raw;
    } catch (_) {}
    throw new Error('OpenAI API: ' + status + ' — ' + detail);
  }

  const parsed = JSON.parse(raw);
  let outputText = parsed.output_text || '';
  if (!outputText && Array.isArray(parsed.output)) {
    parsed.output.forEach(item => {
      if (!item || !Array.isArray(item.content)) return;
      item.content.forEach(part => {
        if (part && part.type === 'output_text' && part.text) outputText += part.text;
      });
    });
  }
  if (!outputText) throw new Error('OpenAI returned no structured result');

  const analysis = JSON.parse(outputText);
  return json_({ ok: true, analysis: analysis, model: parsed.model || 'gpt-5.6-terra' });
}

function enforceAiDailyLimit_() {
  const props = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const dateKey = 'AI_USAGE_DATE';
  const countKey = 'AI_USAGE_COUNT';
  const limit = 60;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const savedDate = props.getProperty(dateKey);
    let count = Number(props.getProperty(countKey) || 0);
    if (savedDate !== today) {
      count = 0;
      props.setProperty(dateKey, today);
    }
    if (count >= limit) throw new Error('Daily AI analysis limit reached');
    props.setProperty(countKey, String(count + 1));
  } finally {
    lock.releaseLock();
  }
}


function analyzeManual_(mealName, weightG) {
  const key = PropertiesService.getScriptProperties().getProperty('MACRO_TRACKER_API_KEY');
  if (!key) throw new Error('MACRO_TRACKER_API_KEY is not configured in Script Properties');

  mealName = String(mealName || '').trim();
  weightG = Number(weightG);
  if (!mealName) throw new Error('Meal name is required');
  if (!isFinite(weightG) || weightG <= 0 || weightG > 5000) throw new Error('Weight must be between 1 and 5000 g');

  enforceAiDailyLimit_();

  const schema = {
    type: 'object',
    additionalProperties: false,
    properties: {
      meal_name: { type: 'string' },
      weight_g: { type: 'number', minimum: 0 },
      kcal: { type: 'number', minimum: 0 },
      protein_g: { type: 'number', minimum: 0 },
      fat_g: { type: 'number', minimum: 0 },
      carbs_g: { type: 'number', minimum: 0 },
      confidence: { type: 'number', minimum: 0, maximum: 1 },
      assumptions: { type: 'string' }
    },
    required: ['meal_name','weight_g','kcal','protein_g','fat_g','carbs_g','confidence','assumptions']
  };

  const payload = {
    model: 'gpt-5.6-terra',
    store: false,
    reasoning: { effort: 'low' },
    input: [
      {
        role: 'developer',
        content: [{
          type: 'input_text',
          text: 'You estimate nutrition for a personal macro tracker. Use typical nutritional values for the named food or dish. Return a practical estimate, not a medical claim. Use Ukrainian for meal_name and assumptions. Calories and macros must describe the same requested portion. If preparation or brand is unspecified, use a common average preparation and state the assumption briefly.'
        }]
      },
      {
        role: 'user',
        content: [{
          type: 'input_text',
          text: 'Estimate nutrition for exactly ' + weightG + ' g of: ' + mealName + '. Keep weight_g equal to ' + weightG + '.'
        }]
      }
    ],
    text: {
      format: {
        type: 'json_schema',
        name: 'nutrition_estimate',
        strict: true,
        schema: schema
      }
    }
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const raw = response.getContentText();
  if (status < 200 || status >= 300) {
    let detail = raw;
    try {
      const parsedError = JSON.parse(raw);
      detail = parsedError && parsedError.error && parsedError.error.message ? parsedError.error.message : raw;
    } catch (_) {}
    throw new Error('OpenAI API: ' + status + ' — ' + detail);
  }

  const parsed = JSON.parse(raw);
  let outputText = parsed.output_text || '';
  if (!outputText && Array.isArray(parsed.output)) {
    parsed.output.forEach(item => {
      if (!item || !Array.isArray(item.content)) return;
      item.content.forEach(part => {
        if (part && part.type === 'output_text' && part.text) outputText += part.text;
      });
    });
  }
  if (!outputText) throw new Error('OpenAI returned no structured result');

  const analysis = JSON.parse(outputText);
  analysis.weight_g = weightG;
  return json_({ ok: true, analysis: analysis, model: parsed.model || 'gpt-5.6-terra' });
}
