const $ = (id) => document.getElementById(id);
const STORE = 'bzhv_tracker_entries_v2';
const SETTINGS = 'bzhv_tracker_settings_v2';
const PENDING = 'bzhv_tracker_pending_v1';
const defaultSettings = { kcal: 0, protein: 0, fat: 0, carbs: 0 };
let photoDataUrl = '';
let syncing = false;

function loadJson(key, fallback){
  try { return JSON.parse(localStorage.getItem(key) || JSON.stringify(fallback)); }
  catch { return fallback; }
}
function saveJson(key, value){ localStorage.setItem(key, JSON.stringify(value)); }
function loadEntries(){ return loadJson(STORE, []); }
function saveEntries(v){ saveJson(STORE, v); }
function loadSettings(){ return {...defaultSettings, ...loadJson(SETTINGS, {})}; }
function saveSettings(v){ saveJson(SETTINGS, v); }
function loadPending(){ return loadJson(PENDING, []); }
function savePending(v){ saveJson(PENDING, v); }
function dayKey(d=new Date()){ const y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,'0'),day=String(d.getDate()).padStart(2,'0'); return `${y}-${m}-${day}`; }
function fmtTime(iso){ return new Date(iso).toLocaleTimeString('uk-UA',{hour:'2-digit',minute:'2-digit'}); }
function fmtDate(iso){ return new Date(iso).toLocaleDateString('uk-UA',{day:'2-digit',month:'2-digit',year:'numeric'}); }
function sum(items,key){ return items.reduce((a,x)=>a+(Number(x[key])||0),0); }
function pct(v,t){ return t>0?Math.min(100,Math.round(v/t*100)):0; }
function esc(s=''){ return String(s).replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c])); }
function hasBackend(){ return Boolean(window.APP_CONFIG?.appsScriptUrl); }
function normalizeEntry(x){
  const dt = x.datetime ? new Date(x.datetime) : new Date();
  return {
    id:String(x.id || ''),
    datetime:Number.isNaN(dt.getTime()) ? new Date().toISOString() : dt.toISOString(),
    date:x.date || dayKey(dt),
    meal_name:String(x.meal_name || ''),
    input_type:String(x.input_type || 'manual'),
    weight_g:Number(x.weight_g)||0,
    kcal:Number(x.kcal)||0,
    protein:Number(x.protein ?? x.protein_g)||0,
    fat:Number(x.fat ?? x.fat_g)||0,
    carbs:Number(x.carbs ?? x.carbs_g)||0,
    image_url:String(x.image_url || ''),
    comment:String(x.comment || ''),
    sync_status:String(x.sync_status || 'synced')
  };
}

async function api(action,payload={}){
  const url = window.APP_CONFIG?.appsScriptUrl;
  if(!url) return {ok:false,offline:true,error:'Backend not configured'};
  const controller = new AbortController();
  const timeout = setTimeout(()=>controller.abort(), 45000);
  try{
    const res = await fetch(url,{
      method:'POST',
      headers:{'Content-Type':'text/plain;charset=utf-8'},
      body:JSON.stringify({action,...payload}),
      signal:controller.signal,
      cache:'no-store'
    });
    if(!res.ok) throw new Error(`HTTP ${res.status}`);
    return await res.json();
  } finally { clearTimeout(timeout); }
}

function setSyncState(text, kind='muted'){
  const el=$('syncStatus'); if(!el) return;
  el.textContent=text; el.dataset.kind=kind;
}

async function bootstrapFromServer(){
  if(!hasBackend()) { setSyncState('Локальний режим'); return; }
  setSyncState('Синхронізація…');
  try{
    const r = await api('bootstrap');
    if(!r.ok) throw new Error(r.error || 'Bootstrap failed');
    const serverEntries=(r.entries||[]).map(normalizeEntry);
    const pendingIds=new Set(loadPending().map(x=>x.entry?.id).filter(Boolean));
    const localPending=loadEntries().filter(x=>pendingIds.has(x.id));
    const merged=new Map();
    [...serverEntries,...localPending].forEach(x=>merged.set(x.id,x));
    saveEntries([...merged.values()]);
    if(r.settings) saveSettings({
      kcal:Number(r.settings.kcal)||0,
      protein:Number(r.settings.protein)||0,
      fat:Number(r.settings.fat)||0,
      carbs:Number(r.settings.carbs)||0
    });
    setSyncState('Google Sheets синхронізовано','ok');
    render();
    await flushPending();
  }catch(e){
    console.warn(e); setSyncState('Офлайн — дані збережені на телефоні','warn');
  }
}

function queueAction(action,payload){
  const q=loadPending();
  q.push({id:crypto.randomUUID?crypto.randomUUID():String(Date.now()+Math.random()),action,payload,created_at:new Date().toISOString()});
  savePending(q);
}

async function flushPending(){
  if(syncing || !hasBackend() || !navigator.onLine) return;
  syncing=true;
  const q=loadPending();
  if(!q.length){ syncing=false; return; }
  setSyncState(`Синхронізація: ${q.length}…`);
  const remain=[];
  for(const item of q){
    try{
      const r=await api(item.action,item.payload);
      if(!r.ok) throw new Error(r.error||'Sync failed');
      if(item.action==='add_food' && item.payload?.entry?.id){
        const entries=loadEntries();
        const i=entries.findIndex(x=>x.id===item.payload.entry.id);
        if(i>=0){ entries[i].sync_status='synced'; if(r.image_url) entries[i].image_url=r.image_url; saveEntries(entries); }
      }
    }catch(e){ remain.push(item); }
  }
  savePending(remain);
  setSyncState(remain.length?`Очікує синхронізації: ${remain.length}`:'Google Sheets синхронізовано',remain.length?'warn':'ok');
  syncing=false; render();
}

async function compressImage(file, maxSide=1600, quality=.78){
  const img=await new Promise((resolve,reject)=>{ const i=new Image(); i.onload=()=>resolve(i); i.onerror=reject; i.src=URL.createObjectURL(file); });
  const scale=Math.min(1,maxSide/Math.max(img.width,img.height));
  const w=Math.max(1,Math.round(img.width*scale)), h=Math.max(1,Math.round(img.height*scale));
  const canvas=document.createElement('canvas'); canvas.width=w; canvas.height=h;
  canvas.getContext('2d').drawImage(img,0,0,w,h);
  URL.revokeObjectURL(img.src);
  return canvas.toDataURL('image/jpeg',quality);
}

async function analyzeCurrentPhoto(){
  if(!photoDataUrl || !hasBackend()) return;
  const note=$('aiNote');
  note.textContent='AI аналізує фото…';
  try{
    const r=await api('analyze_photo',{imageDataUrl:photoDataUrl,mode:$('inputType').value});
    if(!r.ok) throw new Error(r.error||'AI analysis failed');
    const a=r.analysis||{};
    if(a.meal_name) $('mealName').value=a.meal_name;
    if(Number(a.weight_g)>0) $('weightG').value=Math.round(Number(a.weight_g));
    if(Number(a.kcal)>=0) $('kcal').value=Math.round(Number(a.kcal));
    if(Number(a.protein_g)>=0) $('protein').value=Math.round(Number(a.protein_g)*10)/10;
    if(Number(a.fat_g)>=0) $('fat').value=Math.round(Number(a.fat_g)*10)/10;
    if(Number(a.carbs_g)>=0) $('carbs').value=Math.round(Number(a.carbs_g)*10)/10;
    if(a.assumptions && !$('comment').value.trim()) $('comment').value=a.assumptions;
    const confidence=Number(a.confidence);
    const confidenceText=Number.isFinite(confidence)?` · впевненість ${Math.round(confidence*100)}%`:'';
    note.textContent=`AI заповнив оцінку${confidenceText}. Перевір значення перед збереженням.`;
  }catch(err){
    console.warn(err);
    note.textContent=`AI-аналіз не вдався: ${err.message || err}. Можна заповнити БЖВ вручну.`;
  }
}

function render(){
  const entries = loadEntries().map(normalizeEntry).sort((a,b)=>new Date(b.datetime)-new Date(a.datetime));
  const s = loadSettings();
  const today = entries.filter(x=>x.date===dayKey());
  const totals = {kcal:sum(today,'kcal'),protein:sum(today,'protein'),fat:sum(today,'fat'),carbs:sum(today,'carbs')};

  $('todayLabel').textContent = new Date().toLocaleDateString('uk-UA',{weekday:'long',day:'numeric',month:'long'});
  ['kcal','protein','fat','carbs'].forEach(k=>{
    $(k+'Now').textContent = Math.round(totals[k]*10)/10;
    $(k+'Target').textContent = s[k] || '—';
    const p = pct(totals[k],s[k]);
    $(k+'Bar').style.width = p+'%';
    if(k==='kcal'){ $('kcalPct').textContent=s.kcal?p+'%':'—'; $('kcalRing').style.background=`conic-gradient(var(--accent) ${p*3.6}deg,#2a2d33 0deg)`; }
  });
  $('todayCount').textContent = `${today.length} запис${today.length===1?'':today.length<5?'и':'ів'}`;
  $('todayList').className='food-list'+(today.length?'':' empty-state');
  $('todayList').innerHTML = today.length? today.map(foodCard).join(''):'Поки нічого не додано.';

  $('historyList').className='food-list'+(entries.length?'':' empty-state');
  $('historyList').innerHTML = entries.length? entries.map(foodCard).join(''):'Немає записів.';
  renderWeek(entries,s);

  $('setKcal').value=s.kcal||''; $('setProtein').value=s.protein||''; $('setFat').value=s.fat||''; $('setCarbs').value=s.carbs||'';
  $('syncNote').textContent = hasBackend() ? 'Google Sheets: синхронізація увімкнена' : 'Google Sheets: потрібен URL Apps Script Web App';
  if(!hasBackend()) setSyncState('Локальний режим');
}

function foodCard(x){
  const pending=x.sync_status==='pending'?'<span class="sync-chip">очікує синхронізації</span>':'';
  return `<article class="food-item">
    <div><div class="name">${esc(x.meal_name)}</div><div class="meta">${fmtDate(x.datetime)} · ${fmtTime(x.datetime)} · ${esc(x.input_type)} ${pending}</div></div>
    <div class="kcal">${Math.round(x.kcal)} kcal</div>
    <div class="macros"><span class="pill">Б ${x.protein} г</span><span class="pill">Ж ${x.fat} г</span><span class="pill">В ${x.carbs} г</span>${x.weight_g?`<span class="pill">${x.weight_g} г</span>`:''}</div>
  </article>`;
}

function renderWeek(entries,s){
  const days=[]; const now=new Date();
  for(let i=6;i>=0;i--){ const d=new Date(now); d.setHours(0,0,0,0); d.setDate(now.getDate()-i); const k=dayKey(d); const items=entries.filter(x=>x.date===k); days.push({date:d,kcal:sum(items,'kcal'),protein:sum(items,'protein'),fat:sum(items,'fat'),carbs:sum(items,'carbs')}); }
  const activeDays=days.filter(d=>d.kcal>0);
  const divisor=activeDays.length||7;
  const avg=(key)=>Math.round(days.reduce((a,d)=>a+d[key],0)/divisor);
  $('avgKcal').textContent=avg('kcal'); $('avgProtein').textContent=avg('protein'); $('avgFat').textContent=avg('fat'); $('avgCarbs').textContent=avg('carbs');
  $('weekAvg').textContent=s.kcal?`норма ${s.kcal} kcal`:'денну норму не задано';
  const max=Math.max(s.kcal||0,...days.map(d=>d.kcal),1);
  $('weekChart').innerHTML=days.map(d=>`<div class="bar-wrap"><div class="bar" style="height:${Math.max(2,d.kcal/max*160)}px" title="${Math.round(d.kcal)} kcal"></div><small>${d.date.toLocaleDateString('uk-UA',{weekday:'short'}).replace('.','')}</small></div>`).join('');
}

function setMode(mode){
  $('inputType').value=mode;
  document.querySelectorAll('.mode-tab').forEach(b=>b.classList.toggle('active',b.dataset.mode===mode));
  const photo = mode!=='manual';
  $('photoBlock').style.display=photo?'block':'none';
  $('photoPrompt').textContent = mode==='label'?'Сфотографувати етикетку':'Сфотографувати страву';
}

document.querySelectorAll('.nav-btn').forEach(btn=>btn.addEventListener('click',()=>{
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active')); btn.classList.add('active');
  document.querySelectorAll('.screen').forEach(s=>s.classList.toggle('active',s.dataset.screen===btn.dataset.target));
}));
document.querySelectorAll('.mode-tab').forEach(btn=>btn.addEventListener('click',()=>setMode(btn.dataset.mode)));
$('openAdd').addEventListener('click',()=>{ $('foodForm').reset(); photoDataUrl=''; $('photoPreview').hidden=true; $('aiNote').textContent='Після фото AI запропонує назву, вагу, калорії та БЖВ. Перед збереженням усе можна змінити.'; setMode('photo'); $('addDialog').showModal(); });
$('openSettings').addEventListener('click',()=>$('settingsDialog').showModal());
$('closeAdd').addEventListener('click',()=>$('addDialog').close());
$('closeSettings').addEventListener('click',()=>$('settingsDialog').close());
$('retrySync').addEventListener('click',()=>bootstrapFromServer());

$('photoInput').addEventListener('change',async e=>{
  const file=e.target.files?.[0]; if(!file)return;
  $('photoPrompt').textContent='Обробляю фото…';
  try{
    photoDataUrl=await compressImage(file);
    $('photoPreview').src=photoDataUrl; $('photoPreview').hidden=false;
    $('photoPrompt').textContent=$('inputType').value==='label'?'Замінити фото етикетки':'Замінити фото страви';
    await analyzeCurrentPhoto();
  }catch(err){ alert('Не вдалося обробити фото. Спробуй інше.'); }
});

$('foodForm').addEventListener('submit',async e=>{
  e.preventDefault();
  const now=new Date();
  const entry={
    id:crypto.randomUUID?crypto.randomUUID():String(Date.now()), datetime:now.toISOString(), date:dayKey(now),
    meal_name:$('mealName').value.trim(), input_type:$('inputType').value, weight_g:Number($('weightG').value)||0,
    kcal:Number($('kcal').value)||0, protein:Number($('protein').value)||0, fat:Number($('fat').value)||0, carbs:Number($('carbs').value)||0,
    comment:$('comment').value.trim(), imageDataUrl:photoDataUrl||'', sync_status:hasBackend()?'pending':'local'
  };
  const entries=loadEntries(); entries.push({...entry,imageDataUrl:''}); saveEntries(entries);
  if(hasBackend()) queueAction('add_food',{entry});
  render(); $('addDialog').close();
  await flushPending();
});

$('settingsForm').addEventListener('submit',async e=>{
  e.preventDefault();
  const s={kcal:Number($('setKcal').value),protein:Number($('setProtein').value),fat:Number($('setFat').value),carbs:Number($('setCarbs').value)};
  saveSettings(s); render(); $('settingsDialog').close();
  if(hasBackend()){ queueAction('save_settings',{settings:s}); await flushPending(); }
});
$('clearDemo').addEventListener('click',()=>{ if(confirm('Очистити локальну копію на цьому пристрої? Дані в Google Sheets не видаляються.')){localStorage.removeItem(STORE);render();} });

window.addEventListener('online',()=>flushPending());
if('serviceWorker' in navigator) window.addEventListener('load',()=>navigator.serviceWorker.register('sw.js'));
render();
bootstrapFromServer();
