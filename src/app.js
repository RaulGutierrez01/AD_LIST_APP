/*
 * QA MANTO - Aplicacion web de consulta aeronáutica
 * Stack: HTML5 + CSS3 + JavaScript ES6+ + SheetJS
 * Persistencia de usuarios/sesion: LocalStorage
 * Datos: Excel local (.xlsx/.xls/.csv) o URL directa con CORS permitido.
 */

const USERS_KEY = 'qaMantoUsers';
const SESSION_KEY = 'currentUser';
const DATA_KEY = 'qaMantoData';
const DATA_META_KEY = 'qaMantoDataMeta';
const DEFAULT_FILE = 'data/EA_1.xlsx';
const HISTORY_KEY = 'qaMantoHistory';
const FC_LIMIT = 2000;
const CONTROL_SHEET = 'HISTORICO';
const CONTROL_NAME = 'CONTROL_AD.xlsx';
// Nombres de columna del Excel (sin espacios/guiones, en mayúsculas) donde se buscan los datos.
// Si tus columnas se llaman distinto, agrega aquí el nombre.
const COL_FC_PREV = ['CICLOANT'];   // CICLO_ANT: ciclo anterior
const COL_FC_CURR = ['CICLOACT'];   // CICLO_ACT: ciclo actual
const COL_NUM = ['ADNUMBER','SBNUMBER','ADNO','SBNO','NUMEROAD','NUMEROSB','ADNUM','SBNUM','ADSB','ADSBNUMBER'];

const state = {
  view: 'login',
  rows: [],
  sourceName: '',
  sourceType: '',
  isLoaded: false,
  oneDriveUrl: '',
  currentUser: null,
  lastResults: [],
  lastSearch: null,
  selected: null,
  controlHandle: null,
  findingRows: [], findingFile: '', findingQuery: '', findingFilters: {risk:'', auditor:'', source:'', category:''}, findingSort: 'Finding ID', findingAsc: true, findingSelected: null
};

const $ = (selector, root = document) => root.querySelector(selector);

function getUsers() {
  try { return JSON.parse(localStorage.getItem(USERS_KEY) || '[]'); }
  catch { return []; }
}

function setUsers(users) { localStorage.setItem(USERS_KEY, JSON.stringify(users)); }
function getSession() {
  try { return JSON.parse(localStorage.getItem(SESSION_KEY) || 'null'); }
  catch { return null; }
}
function setSession(user) { localStorage.setItem(SESSION_KEY, JSON.stringify(user)); }
function clearSession() { localStorage.removeItem(SESSION_KEY); }

function escapeHTML(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;').replaceAll("'", '&#039;');
}

function toast(message, type = 'info') {
  const container = $('#toast-container');
  const item = document.createElement('div');
  item.className = `toast ${type}`;
  item.textContent = message;
  container.appendChild(item);
  setTimeout(() => item.remove(), 3600);
}

function normalize(value) {
  return String(value ?? '').trim().toUpperCase();
}

function formatDate(value) {
  if (value === null || value === undefined || value === '') return '—';
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toLocaleString('es-PA', { dateStyle: 'short', timeStyle: 'short' });
  }
  // SheetJS may return Excel serial dates depending on parsing mode.
  if (typeof value === 'number' && value > 20000 && value < 80000) {
    const date = XLSX.SSF.parse_date_code(value);
    if (date) {
      const d = new Date(date.y, date.m - 1, date.d, date.H || 0, date.M || 0, date.S || 0);
      return d.toLocaleString('es-PA', { dateStyle: 'short', timeStyle: 'short' });
    }
  }
  const d = new Date(value);
  if (!Number.isNaN(d.getTime()) && /\d/.test(String(value))) {
    return d.toLocaleString('es-PA', { dateStyle: 'short', timeStyle: 'short' });
  }
  return String(value);
}

function extractEA(description) {
  const text = String(description ?? '').trim();
  // Prioridad: patrón explícito "EA 2000-01516". Permite letras/sufijos posteriores.
  const match = text.match(/\bEA\s*([A-Z0-9]+-[A-Z0-9]+)\b/i);
  if (match) return match[1];
  // Fallback para descripciones donde el código está en posiciones similares.
  const fallback = text.match(/\b(\d{3,5}-[A-Z0-9]{3,8})\b/i);
  return fallback ? fallback[1] : '';
}

function validateRequiredHeaders(rows) {
  const required = ['AIRCRAFT','TASK_DESCRIPTION','TASK_STATUS','ORIGINATOR','TASK_BARCODE','COMPLETION_DT','SCHEDULED_START_DT'];
  const first = rows[0] || {};
  return required.filter(key => !Object.prototype.hasOwnProperty.call(first, key));
}

function parseWorkbook(arrayBuffer, sourceName, sourceType = 'Excel') {
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error('El archivo no contiene hojas de cálculo.');
  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
  if (!rows.length) throw new Error('El archivo no contiene registros.');

  const missing = validateRequiredHeaders(rows);
  if (missing.length) {
    throw new Error(`Faltan columnas requeridas: ${missing.join(', ')}`);
  }

  state.rows = rows;
  state.sourceName = sourceName;
  state.sourceType = sourceType;
  state.isLoaded = true;
  state.lastResults = [];

  // Se guarda una copia para conservar el último dataset entre recargas.
  try {
    localStorage.setItem(DATA_KEY, JSON.stringify(rows));
    localStorage.setItem(DATA_META_KEY, JSON.stringify({ sourceName, sourceType, loadedAt: Date.now() }));
  } catch (error) {
    console.warn('No fue posible guardar el dataset en LocalStorage:', error);
  }

  toast(`Archivo cargado: ${sourceName} (${rows.length.toLocaleString()} registros)`, 'success');
  render();
}

async function loadLocalFile(file) {
  if (!file) return;
  const allowed = /\.(xlsx|xls|csv)$/i.test(file.name);
  if (!allowed) {
    toast('Selecciona un archivo .xlsx, .xls o .csv.', 'error');
    return;
  }
  try {
    const buffer = await file.arrayBuffer();
    parseWorkbook(buffer, file.name, 'Archivo local');
  } catch (error) {
    console.error(error);
    toast(error.message || 'No se pudo leer el archivo.', 'error');
  }
}

async function loadDefaultFile() {
  try {
    const response = await fetch(DEFAULT_FILE);
    if (!response.ok) throw new Error('No se encontró data/EA_1.xlsx.');
    const buffer = await response.arrayBuffer();
    parseWorkbook(buffer, 'EA_1.xlsx', 'Dataset del proyecto');
  } catch (error) {
    console.warn('Carga automática del Excel:', error.message);
    restoreSavedDataset();
  }
}

function restoreSavedDataset() {
  try {
    const rows = JSON.parse(localStorage.getItem(DATA_KEY) || '[]');
    const meta = JSON.parse(localStorage.getItem(DATA_META_KEY) || '{}');
    if (rows.length) {
      state.rows = rows;
      state.sourceName = meta.sourceName || 'Dataset guardado localmente';
      state.sourceType = meta.sourceType || 'LocalStorage';
      state.isLoaded = true;
      return true;
    }
  } catch (error) {
    console.warn(error);
  }
  return false;
}

async function loadOneDriveUrl(url) {
  if (!url) return;
  // OneDrive debe proporcionar un enlace que devuelva el archivo o una URL de descarga.
  // El servidor de destino debe permitir CORS. Para entornos empresariales, sustituir
  // esta función por el OneDrive File Picker SDK con clientId/redirectUri configurados.
  try {
    const response = await fetch(url, { mode: 'cors' });
    if (!response.ok) throw new Error(`OneDrive respondió HTTP ${response.status}.`);
    const buffer = await response.arrayBuffer();
    const name = url.split('/').pop()?.split('?')[0] || 'archivo-OneDrive.xlsx';
    parseWorkbook(buffer, name, 'OneDrive / URL');
  } catch (error) {
    console.error(error);
    toast('No se pudo leer la URL de OneDrive. Verifica que sea un enlace de descarga y que permita CORS.', 'error');
  }
}


function findingField(row, field) {
  const aliases = {
    'Finding ID':['Finding ID','FindingID','ID','Finding Id'],
    'Title':['Title','Finding Title'],
    'Description':['Description','Finding Description'],
    'Risk':['Risk','Risk Level'],
    'Auditor':['Auditor','Audit By'],
    'Source':['Source','Audit Source'],
    'Category':['Category','Finding Category']
  };
  const keys=Object.keys(row||{});
  const wanted=aliases[field]||[field];
  const key=keys.find(k=>wanted.some(w=>normalize(k).replace(/[^A-Z0-9]/g,'')===normalize(w).replace(/[^A-Z0-9]/g,'')));
  return key ? String(row[key]??'').trim() : '';
}
function findingsOptions(key) {
  const vals=[...new Set(state.findingRows.map(r=>findingField(r,key==='risk'?'Risk':key==='auditor'?'Auditor':key==='source'?'Source':'Category')).filter(Boolean))].sort();
  return `<option value="">Todos</option>${vals.map(v=>`<option value="${escapeHTML(v)}" ${state.findingFilters[key]===v?'selected':''}>${escapeHTML(v)}</option>`).join('')}`;
}
async function loadFindingsFile(file) {
  if(!file) return;
  if(!/\.(xlsx|xls|csv)$/i.test(file.name)){toast('Carga un archivo Excel o CSV válido.','error');return;}
  try {
    const wb=XLSX.read(await file.arrayBuffer(),{type:'array'});
    const sheet=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(sheet,{defval:'',raw:false});
    if(!rows.length) throw new Error('El archivo no contiene registros.');
    const needed=['Finding ID','Title','Description','Risk','Auditor','Source','Category'];
    const first=rows[0], missing=needed.filter(f=>!Object.keys(first).some(k=>normalize(k).replace(/[^A-Z0-9]/g,'')===normalize(f).replace(/[^A-Z0-9]/g,'')));
    if(missing.length) throw new Error('No se encontraron las columnas: '+missing.join(', '));
    state.findingRows=rows; state.findingFile=file.name; state.findingQuery=''; state.findingFilters={risk:'',auditor:'',source:'',category:''};
    render(); toast(`Excel cargado: ${rows.filter(r=>findingField(r,'Finding ID').trim()!=='').length.toLocaleString()} hallazgos con Finding ID`, 'success');
  }catch(e){toast(e.message||'No se pudo leer el Excel.','error');}
}
function renderFindings(){
 const rows=state.findingRows;
 const count=rows.filter(r=>findingField(r,'Finding ID').trim()!=='').length;
 const riskCounts={}; const categoryCounts={}; const auditorCounts={}; const sourceCounts={};
 rows.forEach(r=>{[['Risk',riskCounts],['Category',categoryCounts],['Auditor',auditorCounts],['Source',sourceCounts]].forEach(([k,obj])=>{const v=findingField(r,k)||'Sin dato';obj[v]=(obj[v]||0)+1;});});
 const stats=(label,obj,mode='records')=>{const total=mode==='unique'?Object.keys(obj).length:Object.values(obj).reduce((a,b)=>a+b,0);return `<div class="finding-stat"><strong>${total.toLocaleString()}</strong><span>${label}</span></div>`;};
 return `<section class="screen">${headerHTML()}<div class="page">
 <div class="topbar"><h2>FINDINGS</h2><div class="topbar-actions"><button id="back-dashboard" class="btn btn-secondary">← Dashboard</button></div></div>
 <div class="hero"><h2>Gestión de Hallazgos</h2><p>Consulta, filtra y analiza los hallazgos de auditoría desde tu archivo.</p></div>
 <section class="data-source"><div class="source-head"><div class="source-title"><span class="status-dot ${count?'active':''}"></span><div><div class="source-name">${count?escapeHTML(state.findingFile):'No hay Excel cargado'}</div><div class="source-meta">${state.findingRows.length?count.toLocaleString()+' hallazgos con Finding ID válido':'Carga el archivo para comenzar'}</div></div></div>
 <div class="source-actions"><label for="findings-file" class="btn btn-primary">Cargar Excel</label><input id="findings-file" class="hidden" type="file" accept=".xlsx,.xls,.csv"></div></div></section>
 ${count?`<div class="finding-stats">${stats('Hallazgos totales',{'Total':count})}${stats('Niveles de riesgo',riskCounts,'unique')}${stats('Categorías',categoryCounts,'unique')}${stats('Auditores únicos',auditorCounts,'unique')}${stats('Fuentes',sourceCounts,'unique')}</div>`:''}
 <section class="search-card"><h3>Búsqueda y filtros</h3><div class="finding-search"><input id="findings-query" value="${escapeHTML(state.findingQuery)}" placeholder="Busca en lenguaje natural o por palabra clave…"><div class="field-help">Ejemplos: high risk, training, auditor, safety, equipment inspection.</div></div>
 <div class="finding-filters"><div><label>Risk</label><select id="filter-risk">${findingsOptions('risk')}</select></div><div><label>Auditor</label><select id="filter-auditor">${findingsOptions('auditor')}</select></div><div><label>Source</label><select id="filter-source">${findingsOptions('source')}</select></div><div><label>Category</label><select id="filter-category">${findingsOptions('category')}</select></div><button id="findings-clear" class="btn btn-secondary">Limpiar filtros</button></div></section>
 <section class="results-card"><div class="results-head"><h3>Hallazgos</h3><span id="findings-count" class="result-count">${count?count.toLocaleString()+' registros cargados':'Sin datos'}</span></div><div id="findings-table-area">${renderFindingsTable(rows)}</div></section>
 </div></section>`;
}
function getFilteredFindings(){
 let rows=state.findingRows;
 const q=normalize(state.findingQuery);
 const filters=state.findingFilters;
 if(q){
   const riskWords={HIGH:['HIGH','ALTO','CRITICAL','SEVERE'],MEDIUM:['MEDIUM','MODERATE','MEDIO'],LOW:['LOW','MINOR','BAJO']};
   let riskWanted='';
   if(/\b(HIGH|ALTO|CRITICAL|SEVERE)\b/.test(q))riskWanted='high';
   else if(/\b(MEDIUM|MODERATE|MEDIO)\b/.test(q))riskWanted='medium';
   else if(/\b(LOW|MINOR|BAJO)\b/.test(q))riskWanted='low';
   const cleaned=q.replace(/\b(FINDINGS?|HALLAZGOS?|RELATED TO|RELACIONADOS CON|FOUND BY|ENCONTRADOS POR|CATEGORY|CATEGORIA|RISK|RIESGO|HIGH|ALTO|MEDIUM|MEDIO|LOW|BAJO|CRITICAL|SEVERE|MODERATE|MINOR)\b/g,' ').trim();
   rows=rows.filter(r=>{
    const risk=normalize(findingField(r,'Risk'));
    const all=Object.values(r).map(v=>String(v??'')).join(' ').toUpperCase();
    const riskMatch=!riskWanted || (riskWanted==='high'?riskWords.high.some(w=>risk.includes(w)):riskWanted==='medium'?riskWords.medium.some(w=>risk.includes(w)):riskWords.low.some(w=>risk.includes(w)));
    return riskMatch && (!cleaned || all.includes(cleaned) || cleaned.split(/\s+/).filter(Boolean).every(word=>all.includes(word)));
   });
 }
 rows=rows.filter(r=>['risk','auditor','source','category'].every(k=>!filters[k]||findingField(r,k==='risk'?'Risk':k==='auditor'?'Auditor':k==='source'?'Source':'Category')===filters[k]));
 return rows.sort((a,b)=>{const x=findingField(a,state.findingSort),y=findingField(b,state.findingSort);return x.localeCompare(y,undefined,{numeric:true,sensitivity:'base'})*(state.findingAsc?1:-1);});
}
function renderFindingsTable(rows){
 if(!state.findingRows.length)return `<div class="empty-results"><strong>Sin Excel cargado</strong><span>Selecciona un archivo para mostrar los hallazgos.</span></div>`;
 if(!rows.length)return `<div class="empty-results"><strong>No se encontraron resultados</strong><span>Prueba otra búsqueda o elimina alguno de los filtros.</span></div>`;
 const shown=rows.slice(0,1000);
 return `<div class="table-wrap"><table class="findings-table"><thead><tr>${['Finding ID','Title','Risk','Auditor','Source','Category'].map(k=>`<th><button class="sort-finding" data-sort-finding="${k}">${k} ↕</button></th>`).join('')}<th>Detalle</th></tr></thead><tbody>${shown.map(r=>{const originalIndex=state.findingRows.indexOf(r);return `<tr><td>${escapeHTML(findingField(r,'Finding ID'))}</td><td>${escapeHTML(findingField(r,'Title'))}</td><td><span class="risk-chip">${escapeHTML(findingField(r,'Risk')||'—')}</span></td><td>${escapeHTML(findingField(r,'Auditor'))}</td><td>${escapeHTML(findingField(r,'Source'))}</td><td>${escapeHTML(findingField(r,'Category'))}</td><td><button class="btn btn-outline btn-sm" data-finding-detail="${originalIndex}">Ver</button></td></tr>`}).join('')}</tbody></table>${rows.length>1000?`<div class="field-help table-limit">Se muestran los primeros 1,000 de ${rows.length.toLocaleString()} resultados para mantener el rendimiento.</div>`:''}</div>`;
}
function renderFindingsResults(){
 const filtered=getFilteredFindings();
 const area=$('#findings-table-area'); if(area) area.innerHTML=renderFindingsTable(filtered);
 const count=$('#findings-count'); if(count) count.textContent=`${filtered.length.toLocaleString()} coincidencias`;
 bindFindingsTableEvents();
}
function bindFindingsTableEvents(){
 document.querySelectorAll('[data-sort-finding]').forEach(btn=>btn.addEventListener('click',()=>{const key=btn.dataset.sortFinding;state.findingAsc=state.findingSort===key?!state.findingAsc:true;state.findingSort=key;renderFindingsResults();}));
 document.querySelectorAll('[data-finding-detail]').forEach(btn=>btn.addEventListener('click',()=>{state.findingSelected=state.findingRows[Number(btn.dataset.findingDetail)];renderFindingDetail();}));
}
function closeFindingDetail(event){
 if(event){event.preventDefault();event.stopPropagation();}
 const modal=document.getElementById('finding-detail-modal');
 if(modal) modal.remove();
 state.findingSelected=null;
 // La vista FINDINGS permanece montada; solo se actualiza el listado.
 if(state.view==='findings') renderFindingsResults();
}
window.closeFindingDetail=closeFindingDetail;

function renderFindingDetail(){
 const r=state.findingSelected; if(!r)return;
 const root=$('#view-root');
 // Evita acumular overlays si se abre más de un detalle.
 document.getElementById('finding-detail-modal')?.remove();
 root.insertAdjacentHTML('beforeend',`<div class="finding-overlay" id="finding-detail-modal"><div class="finding-modal" role="dialog" aria-modal="true" aria-labelledby="finding-detail-title"><div class="topbar"><h2 id="finding-detail-title">Detalle del hallazgo</h2><button id="finding-detail-close" type="button" class="btn btn-secondary" onclick="window.closeFindingDetail(event)">Cerrar</button></div>${['Finding ID','Title','Description','Risk','Auditor','Source','Category'].map(k=>`<div class="finding-detail-field"><label>${k}</label><p>${escapeHTML(findingField(r,k)||'—')}</p></div>`).join('')}</div></div>`);
}


// Navegación pública para módulos integrados que no comparten el scope interno.
 window.QAMantoNavigateDashboard = function(){
   state.view = 'dashboard';
   state.lastResults = [];
   state.lastSearch = null;
   render();
 };

 function render() {
  const session = getSession();
  state.currentUser = session;
  if (!session && state.view !== 'register') state.view = 'login';
  if (session && ['login','register'].includes(state.view)) state.view = 'dashboard';

  const root = $('#view-root');
  const views = {
    login: renderLogin,
    register: renderRegister,
    dashboard: renderDashboard,
    elt: () => renderUnavailable('ELT'),
    findings: renderFindings,
    case: () => window.ADControlModule.render(),
    ad: renderMonitoreoEA,
    detail: renderDetail,
    proveedor: () => window.ProveedorModule ? window.ProveedorModule.render() : renderUnavailable('Proveedor')
  };
  root.innerHTML = views[state.view]();
  bindViewEvents();

  if (state.view === 'ad') { restoreSearchForm(); updateAircraftSuggestions(); }
  if (state.view === 'detail') updateCycles();
  if (state.view === 'findings') bindFindingsTableEvents();
  if (state.view === 'case') window.ADControlModule?.init();
  if (state.view === 'proveedor') window.ProveedorModule?.init();
  window.scrollTo(0, 0);
}

function restoreSearchForm() {
  const q = state.lastSearch;
  if (!q) return;
  $('#aircraft').value = q.aircraft; $('#ea').value = q.ea;
  $('#status').value = q.status; $('#originator').value = q.originator;
  updateDynamicNumber();
  if ($('#dynamic-number')) $('#dynamic-number').value = q.dynamicNumber;
}

function renderLogin() {
  return `
    <section class="screen auth-screen">
      <div class="auth-card">
        <div class="brand">
          <div class="brand-mark">QA</div>
          <div><h1>QA MANTO</h1><p>Gestión de información aeronáutica</p></div>
        </div>
        <form id="login-form" novalidate>
          <div class="form-group"><label for="login-user">Usuario / Email</label><input id="login-user" autocomplete="username" required></div>
          <div class="form-group"><label for="login-password">Contraseña</label><input id="login-password" type="password" autocomplete="current-password" required></div>
          <div id="login-error" class="error-text hidden"></div>
          <button class="btn btn-primary btn-full" type="submit">Iniciar Sesión</button>
        </form>
        <p class="auth-switch">¿No tienes una cuenta? <button id="go-register" class="link-btn" type="button">Registrarse</button></p>
      </div>
    </section>`;
}

function renderRegister() {
  return `
    <section class="screen auth-screen">
      <div class="auth-card">
        <div class="brand">
          <div class="brand-mark">QA</div>
          <div><h1>Crear cuenta</h1><p>Registro de usuario QA MANTO</p></div>
        </div>
        <form id="register-form" novalidate>
          <div class="form-group"><label for="reg-name">Nombre completo</label><input id="reg-name" required></div>
          <div class="form-group"><label for="reg-email">Correo electrónico</label><input id="reg-email" type="email" required></div>
          <div class="form-group"><label for="reg-user">Nombre de usuario</label><input id="reg-user" autocomplete="username" required></div>
          <div class="form-group"><label for="reg-pass">Contraseña</label><input id="reg-pass" type="password" autocomplete="new-password" required></div>
          <div class="form-group"><label for="reg-confirm">Confirmar contraseña</label><input id="reg-confirm" type="password" autocomplete="new-password" required></div>
          <div id="register-error" class="error-text hidden"></div>
          <button class="btn btn-primary btn-full" type="submit">Guardar</button>
        </form>
        <p class="auth-switch">¿Ya tienes cuenta? <button id="go-login" class="link-btn" type="button">Iniciar sesión</button></p>
      </div>
    </section>`;
}

function renderDashboard() {
  const user = state.currentUser || {};
  return `
    <section class="screen">
      ${headerHTML()}
      <div class="page">
        <div class="hero"><h2>Panel de control</h2><p>Consulta y gestión de información de mantenimiento aeronáutico.</p></div>
        <div class="module-grid">
          <button class="module-card" data-module="findings" type="button"><div class="module-icon">F</div><h3>FINDINGS</h3><p>Gestión, búsqueda y análisis de hallazgos de auditoría.</p></button>
          <button class="module-card" data-module="ad" type="button"><div class="module-icon">EA</div><h3>Monitoreo EA</h3><p>Consulta y filtrado de Engineering Authorizations desde Excel.</p></button>
          <button class="module-card" data-module="case" type="button"><div class="module-icon">AD</div><h3>AD Control</h3><p>Registro, seguimiento y revisión de directivas de aeronavegabilidad.</p></button>
          <button class="module-card" data-module="proveedor" type="button"><div class="module-icon">P</div><h3>Proveedor</h3><p>Sincronización de registros usando el enlace compartido de OneDrive (sin iniciar sesión).</p></button>
        </div>
      </div>
    </section>`;
}

function headerHTML() {
  const name = escapeHTML(state.currentUser?.name || 'Usuario');
  return `<header class="app-header"><div class="header-brand"><div class="brand-mark">QA</div><strong>QA MANTO</strong></div><div class="user-area"><span>Hola, <span class="user-name">${name}</span></span><button id="logout" class="btn btn-secondary" type="button">Cerrar Sesión</button></div></header>`;
}

window.QAMantoHeaderHTML = headerHTML;
window.QAMantoToast = toast;

function renderUnavailable(moduleName) {
  return `
    <section class="screen">
      ${headerHTML()}
      <div class="page"><div class="unavailable">
        <div class="unavailable-icon">✈</div><h2>${escapeHTML(moduleName)}</h2>
        <p>Por el momento no se cuenta con estos datos.</p>
        <button id="back-dashboard" class="btn btn-primary" type="button">Volver al Dashboard</button>
      </div></div>
    </section>`;
}

function renderMonitoreoEA() {
  const sourceLabel = state.isLoaded ? state.sourceName : 'Sin archivo activo';
  const sourceType = state.isLoaded ? state.sourceType : 'Selecciona un Excel o carga el dataset del proyecto';
  const hasResults = state.lastResults.length > 0;
  return `
    <section class="screen">
      ${headerHTML()}
      <div class="page">
        <div class="topbar"><h2>Monitoreo EA</h2><div class="topbar-actions"><button id="back-dashboard" class="btn btn-secondary" type="button">← Dashboard</button></div></div>

        <section class="data-source">
          <div class="source-head">
            <div class="source-title"><span class="status-dot ${state.isLoaded ? 'active' : ''}"></span><div><div class="source-name">Archivo activo: ${escapeHTML(sourceLabel)}</div><div class="source-meta">${escapeHTML(sourceType)}${state.isLoaded ? ` · ${state.rows.length.toLocaleString()} registros` : ''}</div></div></div>
            <div class="source-actions">
              <label class="btn btn-primary" for="excel-file">Cargar desde la PC</label>
              <input id="excel-file" class="hidden" type="file" accept=".xlsx,.xls,.csv">
            
            </div>
          </div>
        </section>

        <section class="search-card">
          <h3>Monitoreo EA · Búsqueda</h3>
          <form id="ad-search-form" novalidate>
            <div class="form-grid">
              <div class="form-group"><label for="aircraft">Aircraft</label><input id="aircraft" list="aircraft-list" autocomplete="off" placeholder="Ej. HP-1371CMP"><datalist id="aircraft-list"></datalist><div class="field-help">Sugiere valores únicos de AIRCRAFT.</div></div>
              <div class="form-group"><label for="ea">N° EA</label><input id="ea" placeholder="Ej. 2400-01516" autocomplete="off"><div class="field-help">Se compara contra el EA extraído de TASK_DESCRIPTION.</div></div>
              <div class="form-group"><label for="status">Status *</label><select id="status" required><option value="">Seleccionar...</option><option value="COMPLETE">Historical only</option><option value="ACTV">Active only</option></select></div>
              <div class="form-group"><label for="originator">Originator *</label><select id="originator" required><option value="">Seleccionar...</option><option value="AD">AD (Directiva de Aeronavegabilidad)</option><option value="SB">SB (Service Bulletin)</option></select></div>
              <div id="dynamic-number-wrap" class="form-group hidden"><label id="dynamic-number-label" for="dynamic-number">N° AD</label><input id="dynamic-number" inputmode="numeric" pattern="[0-9]*" placeholder="Opcional"><div class="field-help">Solo números. Campo opcional.</div></div>
              <div class="field-span-2" style="display:flex; gap:10px; align-items:end; flex-wrap:wrap;"><button class="btn btn-primary" type="submit">Buscar</button><button id="clear-all" class="btn btn-secondary" type="button">Limpiar todo</button></div>
            </div>
          </form>
        </section>

        <section class="results-card" aria-live="polite">
          <div class="results-head"><h3>Resultados</h3><span class="result-count">${hasResults ? `${state.lastResults.length.toLocaleString()} coincidencia(s)` : 'Sin resultados'}</span></div>
          ${hasResults ? resultsTableHTML() : `<div class="empty-results"><strong>Sin resultados para mostrar</strong><span>Ejecuta una búsqueda para consultar el Excel.</span></div>`}
        </section>
      </div>
    </section>`;
}

function resultsTableHTML() {
  const q = state.lastSearch || {};
  const aircraftFilled = Boolean(q.aircraft);
  const headers = ['Barcode', aircraftFilled ? 'N° EA' : 'Aircraft', 'Status', 'Fecha de Cumplimiento / Programación', 'Acción'];
  const rows = state.lastResults.map((row, i) => {
    const date = q.status === 'ACTV' ? row.SCHEDULED_START_DT : row.COMPLETION_DT;
    const first = aircraftFilled ? extractEA(row.TASK_DESCRIPTION) : row.AIRCRAFT;
    return `<tr><td>${escapeHTML(row.TASK_BARCODE)}</td><td>${escapeHTML(first)}</td><td>${statusBadge(row.TASK_STATUS)}</td><td>${escapeHTML(formatDate(date))}</td><td><button class="btn btn-primary btn-sm" type="button" data-detail="${i}">Ver detalle</button></td></tr>`;
  }).join('');
  const cols = headers.map(h => `<th>${h}</th>`).join('');
  return `<div class="table-wrap"><table><thead><tr>${cols}</tr></thead><tbody>${rows}</tbody></table></div>`;
}

function statusBadge(status) {
  const s = normalize(status);
  const cls = s === 'ACTV' ? 'badge-active' : 'badge-complete';
  return `<span class="badge ${cls}">${escapeHTML(status)}</span>`;
}

function bindViewEvents() {
  $('#go-register')?.addEventListener('click', () => { state.view = 'register'; render(); });
  $('#go-login')?.addEventListener('click', () => { state.view = 'login'; render(); });
  $('#logout')?.addEventListener('click', () => { clearSession(); state.view = 'login'; state.currentUser = null; render(); toast('Sesión cerrada.', 'info'); });
  $('#back-dashboard')?.addEventListener('click', () => { state.view = 'dashboard'; state.lastResults = []; state.lastSearch = null; render(); });

  $('#login-form')?.addEventListener('submit', handleLogin);
  $('#register-form')?.addEventListener('submit', handleRegister);

  document.querySelectorAll('[data-module]').forEach(card => card.addEventListener('click', () => {
    state.view = card.dataset.module === 'case' ? 'case' : card.dataset.module;
    render();
    if (state.view === 'ad' && !state.isLoaded) loadDefaultFile();
  }));

  $('#excel-file')?.addEventListener('change', event => loadLocalFile(event.target.files[0]));
  $('#findings-file')?.addEventListener('change', event => loadFindingsFile(event.target.files[0]));
  $('#findings-query')?.addEventListener('input', event => { state.findingQuery = event.target.value; renderFindingsResults(); });
  ['risk','auditor','source','category'].forEach(key => $('#filter-'+key)?.addEventListener('change', event => { state.findingFilters[key] = event.target.value; renderFindingsResults(); }));
  $('#findings-clear')?.addEventListener('click', () => { state.findingQuery=''; state.findingFilters={risk:'',auditor:'',source:'',category:''}; render(); });
  document.querySelectorAll('[data-sort-finding]').forEach(btn => btn.addEventListener('click', () => { const key=btn.dataset.sortFinding; state.findingAsc = state.findingSort===key ? !state.findingAsc : true; state.findingSort=key; renderFindingsResults(); }));
  document.querySelectorAll('[data-finding-detail]').forEach(btn => btn.addEventListener('click', () => { state.findingSelected=state.findingRows[Number(btn.dataset.findingDetail)]; renderFindingDetail(); }));
  $('#load-default')?.addEventListener('click', loadDefaultFile);
  $('#open-onedrive')?.addEventListener('click', openOneDriveDialog);
  $('#originator')?.addEventListener('change', updateDynamicNumber);
  $('#dynamic-number')?.addEventListener('input', event => {
    event.target.value = event.target.value.replace(/\D/g, '');
  });
  $('#aircraft')?.addEventListener('input', updateAircraftSuggestions);
  $('#ad-search-form')?.addEventListener('submit', handleSearch);
  $('#clear-all')?.addEventListener('click', clearAll);
  document.querySelectorAll('[data-detail]').forEach(btn => btn.addEventListener('click', () => {
    state.selected = state.lastResults[Number(btn.dataset.detail)];
    state.view = 'detail';
    render();
  }));
  $('#detail-save')?.addEventListener('click', saveRecord);
  $('#detail-exit')?.addEventListener('click', () => { state.view = 'ad'; state.selected = null; render(); });
}

function handleLogin(event) {
  event.preventDefault();
  const identifier = $('#login-user').value.trim().toLowerCase();
  const password = $('#login-password').value;
  const error = $('#login-error');
  error.classList.add('hidden');

  if (!identifier || !password) {
    error.textContent = 'Completa usuario/email y contraseña.';
    error.classList.remove('hidden');
    return;
  }
  const user = getUsers().find(u => (u.username.toLowerCase() === identifier || u.email.toLowerCase() === identifier) && u.password === password);
  if (!user) {
    error.textContent = 'Usuario o contraseña incorrectos';
    error.classList.remove('hidden');
    return;
  }
  const safeUser = { name: user.name, username: user.username, email: user.email };
  setSession(safeUser);
  state.currentUser = safeUser;
  state.view = 'dashboard';
  render();
}

function handleRegister(event) {
  event.preventDefault();
  const name = $('#reg-name').value.trim();
  const email = $('#reg-email').value.trim().toLowerCase();
  const username = $('#reg-user').value.trim();
  const password = $('#reg-pass').value;
  const confirm = $('#reg-confirm').value;
  const error = $('#register-error');
  error.classList.add('hidden');

  if (!name || !email || !username || !password || !confirm) return showRegisterError('Todos los campos son obligatorios.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return showRegisterError('Ingresa un correo electrónico válido.');
  if (password !== confirm) return showRegisterError('Las contraseñas no coinciden.');
  const users = getUsers();
  if (users.some(u => u.username.toLowerCase() === username.toLowerCase())) return showRegisterError('El nombre de usuario ya existe.');
  if (users.some(u => u.email.toLowerCase() === email)) return showRegisterError('El correo electrónico ya existe.');

  users.push({ name, email, username, password, createdAt: new Date().toISOString() });
  setUsers(users);
  toast('Registro exitoso. Ahora puedes iniciar sesión.', 'success');
  setTimeout(() => { state.view = 'login'; render(); }, 700);
}

function showRegisterError(message) {
  const error = $('#register-error');
  error.textContent = message;
  error.classList.remove('hidden');
}

function updateDynamicNumber() {
  const value = $('#originator')?.value;
  const wrap = $('#dynamic-number-wrap');
  if (!wrap) return;
  if (!value) { wrap.classList.add('hidden'); return; }
  wrap.classList.remove('hidden');
  $('#dynamic-number-label').textContent = value === 'AD' ? 'N° AD' : 'N° SB';
  $('#dynamic-number').value = '';
}

function updateAircraftSuggestions() {
  const list = $('#aircraft-list');
  if (!list) return;
  const current = normalize($('#aircraft')?.value);
  const values = [...new Set(state.rows.map(r => String(r.AIRCRAFT ?? '').trim()).filter(Boolean))]
    .filter(value => !current || normalize(value).includes(current))
    .sort((a,b) => a.localeCompare(b));
  list.innerHTML = values.slice(0, 100).map(value => `<option value="${escapeHTML(value)}"></option>`).join('');
}

function handleSearch(event) {
  event.preventDefault();
  if (!state.isLoaded || !state.rows.length) return toast('Debes cargar o conectar un archivo Excel antes de buscar.', 'error');

  const aircraft = $('#aircraft').value.trim();
  const ea = $('#ea').value.trim();
  const status = $('#status').value;
  const originator = $('#originator').value;
  const dynamicNumber = $('#dynamic-number').value.trim();

  if (!aircraft && !ea) return toast('Completa Aircraft, N° EA o ambos.', 'error');
  if (!status) return toast('Selecciona Status.', 'error');
  if (!originator) return toast('Selecciona Originator.', 'error');
  if (dynamicNumber && !/^\d+$/.test(dynamicNumber)) return toast('El número AD/SB solo puede contener números.', 'error');

  state.lastSearch = { aircraft, ea, status, originator, dynamicNumber };
  const aircraftNorm = normalize(aircraft);
  const eaNorm = normalize(ea).replace(/^EA\s*/, '');

  state.lastResults = state.rows.filter(row => {
    const rowAircraft = normalize(row.AIRCRAFT);
    const rowStatus = normalize(row.TASK_STATUS);
    const rowOriginator = normalize(row.ORIGINATOR);
    const rowEA = normalize(extractEA(row.TASK_DESCRIPTION));

    const aircraftMatch = !aircraftNorm || rowAircraft === aircraftNorm;
    const eaMatch = !eaNorm || rowEA === eaNorm || rowEA.includes(eaNorm);
    const statusMatch = rowStatus === status;
    const originatorMatch = rowOriginator === originator;

    // N° AD/SB es opcional. En este dataset no existe una columna dedicada al número
    // AD/SB; cuando se ingresa, se intenta encontrar el número dentro de TASK_DESCRIPTION.
    let numberMatch = true;
    if (dynamicNumber) numberMatch = normalize(row.TASK_DESCRIPTION).includes(dynamicNumber);

    return aircraftMatch && eaMatch && statusMatch && originatorMatch && numberMatch;
  });

  render();
  if (!state.lastResults.length) toast('No se encontraron registros que coincidan con la búsqueda.', 'info');
}

function clearAll() {
  state.lastResults = [];
  state.lastSearch = null;
  render();
  toast('Formulario restablecido.', 'info');
}

/* ---------- Detalle, ciclos y guardado ---------- */
const keyOf = k => String(k).toUpperCase().replace(/[^A-Z0-9]/g, '');

function findField(row, candidates) {
  const keys = Object.keys(row);
  for (const c of candidates) {
    const k = keys.find(key => keyOf(key) === c && row[key] !== '' && row[key] != null);
    if (k) return { key: k, value: row[k] };
  }
  return null;
}

function toNumber(v) {
  if (v === '' || v == null) return NaN;
  return Number(String(v).replace(/[,\s]/g, ''));
}

function getADSBNumber(row) {
  const col = findField(row, COL_NUM);
  if (col) return { value: String(col.value), source: `Excel (${col.key})` };
  const m = String(row.TASK_DESCRIPTION ?? '').match(/\b(AD|SB)\s*(?:NO\.?|N°|#|:)?\s*([0-9][A-Z0-9.\-\/]*)/i);
  if (m) return { value: m[2], source: 'Excel (TASK_DESCRIPTION)' };
  const typed = state.lastSearch?.dynamicNumber;
  if (typed) return { value: typed, source: 'Número usado en la búsqueda' };
  return { value: '', source: 'No disponible en el Excel' };
}

function renderDetail() {
  const row = state.selected || {};
  const q = state.lastSearch || {};
  const isActive = normalize(row.TASK_STATUS) === 'ACTV';
  const date = isActive ? row.SCHEDULED_START_DT : row.COMPLETION_DT;
  const prev = findField(row, COL_FC_PREV), curr = findField(row, COL_FC_CURR);
  const ad = getADSBNumber(row);
  const label = q.originator === 'SB' ? 'N° SB' : q.originator === 'AD' ? 'N° AD' : 'N° AD / SB';
  const item = (k, v) => `<div class="detail-item"><span>${k}</span><strong>${escapeHTML(v || '—')}</strong></div>`;
  const used = new Set(['AIRCRAFT','TASK_DESCRIPTION','TASK_STATUS','ORIGINATOR','TASK_BARCODE','COMPLETION_DT','SCHEDULED_START_DT']);
  const extra = Object.keys(row).filter(k => !used.has(k) && row[k] !== '' && row[k] != null)
    .map(k => item(k, /_DT$|DATE/i.test(k) ? formatDate(row[k]) : String(row[k]))).join('');
  const fcNote = prev && curr ? 'Valores tomados del Excel: CICLO_ANT y CICLO_ACT.'
    : 'Este registro no tiene valores en las columnas CICLO_ANT / CICLO_ACT del Excel.';
  const fmtFC = f => f ? toNumber(f.value).toLocaleString('es-PA') : '—';
  return `
    <section class="screen">
      ${headerHTML()}
      <div class="page">
        <div class="topbar"><h2>Monitoreo EA · Detalle</h2></div>

        <section class="search-card">
          <h3>Datos de la búsqueda</h3>
          <div class="detail-grid">
            ${item('Aircraft', row.AIRCRAFT)}${item('N° EA', extractEA(row.TASK_DESCRIPTION))}
            ${item('Status', isActive ? 'ACTV (Active)' : 'COMPLETE (Historical)')}${item('Originator', row.ORIGINATOR)}
            ${item('Barcode', row.TASK_BARCODE)}${item(isActive ? 'Fecha programada' : 'Fecha de cumplimiento', formatDate(date))}
            ${item('Descripción', row.TASK_DESCRIPTION)}
          </div>
        </section>

        ${extra ? `<section class="search-card"><h3>Información adicional del Excel</h3><div class="detail-grid">${extra}</div></section>` : ''}

        <section class="search-card">
          <h3>Ciclos de vuelo (FC)</h3>
          <p class="field-help">Un ciclo = despegue + aterrizaje. ${fcNote}</p>
          <div class="form-grid">
            <div class="form-group"><label>Ciclo anterior (CICLO_ANT)</label><div id="fc-prev" class="fc-box">${fmtFC(prev)}</div></div>
            <div class="form-group"><label>Ciclo actual (CICLO_ACT)</label><div id="fc-curr" class="fc-box">${fmtFC(curr)}</div></div>
            <div class="form-group"><label>Diferencia (CICLO_ACT − CICLO_ANT)</label><div id="fc-diff" class="fc-box">—</div></div>
            <div class="form-group"><label>Límite: ${FC_LIMIT.toLocaleString('es-PA')} FC</label><div id="fc-result" class="fc-box">Pendiente</div></div>
          </div>
        </section>

        <section class="search-card">
          <h3>Información adicional</h3>
          <div class="detail-grid">${item(label, ad.value)}${item('Origen del dato', ad.source)}</div>
          <div class="form-group" style="margin-top:16px"><label for="comments">Comentarios</label><textarea id="comments" rows="4" placeholder="Observaciones o notas de la revisión..."></textarea></div>
          <div class="detail-actions"><button id="detail-save" class="btn btn-primary" type="button">SAVE</button><button id="detail-exit" class="btn btn-secondary" type="button">EXIT</button></div>
        </section>
      </div>
    </section>`;
}

function calcCycles() {
  const row = state.selected || {};
  const prev = toNumber(findField(row, COL_FC_PREV)?.value);
  const curr = toNumber(findField(row, COL_FC_CURR)?.value);
  if (Number.isNaN(prev) || Number.isNaN(curr)) return { valid: false };
  const diff = curr - prev;           // Diferencia = CICLO_ACT - CICLO_ANT
  return { valid: diff >= 0, prev, curr, diff, ok: diff >= 0 && diff <= FC_LIMIT };
}

function updateCycles() {
  const r = calcCycles();
  const diffEl = $('#fc-diff'), resEl = $('#fc-result');
  if (!diffEl) return;
  if (r.prev === undefined) { diffEl.textContent = '—'; resEl.textContent = 'Pendiente'; resEl.className = 'fc-box'; return; }
  diffEl.textContent = `${(r.curr - r.prev).toLocaleString('es-PA')} FC`;
  if (!r.valid) { resEl.textContent = 'Revisar datos: CICLO_ACT menor que CICLO_ANT'; resEl.className = 'fc-box fc-warn'; return; }
  resEl.textContent = r.ok ? 'CUMPLE' : 'NO CUMPLE';
  resEl.className = `fc-box ${r.ok ? 'fc-ok' : 'fc-bad'}`;
}

function getHistory() {
  try { return JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]'); } catch { return []; }
}

async function writeControlFile(record) {
  // Chrome/Edge: se abre/crea el Excel de control en la PC y se agrega una fila al histórico.
  if (!state.controlHandle) {
    state.controlHandle = await window.showSaveFilePicker({
      suggestedName: CONTROL_NAME,
      types: [{ description: 'Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
    });
  }
  const handle = state.controlHandle;
  if ((await handle.requestPermission({ mode: 'readwrite' })) !== 'granted') throw new Error('Permiso denegado.');
  const file = await handle.getFile();
  let wb, existing = [];
  if (file.size > 0) {
    wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    if (wb.Sheets[CONTROL_SHEET]) existing = XLSX.utils.sheet_to_json(wb.Sheets[CONTROL_SHEET], { defval: '' });
  } else {
    wb = XLSX.utils.book_new();
  }
  const ws = XLSX.utils.json_to_sheet([...existing, record]);
  ws['!cols'] = Object.keys(record).map(k => ({ wch: Math.max(14, k.length + 2) }));
  ws['!autofilter'] = { ref: ws['!ref'] };
  if (wb.Sheets[CONTROL_SHEET]) wb.Sheets[CONTROL_SHEET] = ws; else XLSX.utils.book_append_sheet(wb, ws, CONTROL_SHEET);
  const writable = await handle.createWritable();
  await writable.write(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
  await writable.close();
}

function downloadHistory(history) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(history), CONTROL_SHEET);
  XLSX.writeFile(wb, CONTROL_NAME);
}

async function saveRecord() {
  const r = calcCycles();
  if (r.prev === undefined) return toast('Este registro no tiene CICLO_ANT / CICLO_ACT en el Excel.', 'error');
  if (!r.valid) return toast('CICLO_ACT es menor que CICLO_ANT: revisa los datos del Excel.', 'error');
  const row = state.selected, q = state.lastSearch || {};
  const record = {
    'Fecha y hora de registro': new Date().toLocaleString('es-PA'),
    'Usuario': state.currentUser?.username || state.currentUser?.name || '',
    'Aircraft': row.AIRCRAFT, 'N° EA': extractEA(row.TASK_DESCRIPTION),
    'Status': row.TASK_STATUS, 'Originator': row.ORIGINATOR, 'Barcode': row.TASK_BARCODE,
    'N° AD/SB': getADSBNumber(row).value,
    'Ciclo anterior': r.prev, 'Ciclo actual': r.curr, 'Diferencia (FC)': r.diff, 'Límite (FC)': FC_LIMIT,
    'Resultado': r.ok ? 'Cumple' : 'No cumple',
    'Comentarios': $('#comments').value.trim()
  };
  const history = [...getHistory(), record];
  try { localStorage.setItem(HISTORY_KEY, JSON.stringify(history)); } catch (e) { console.warn(e); }

  if ('showSaveFilePicker' in window) {
    try {
      await writeControlFile(record);
      return toast('Registro guardado en el Excel de control.', 'success');
    } catch (error) {
      if (error.name === 'AbortError') return toast('Guardado cancelado: no se eligió el archivo de control.', 'info');
      console.error(error);
      toast('No se pudo escribir en el archivo de control. Se descargará el histórico.', 'error');
    }
  }
  downloadHistory(history);
  toast('Este navegador no permite editar el archivo directamente: se descargó el histórico actualizado.', 'info');
}

function openOneDriveDialog() {
  const overlay = document.createElement('div');
  overlay.className = 'overlay';
  overlay.innerHTML = `<div class="modal">
    <h3>Conectar archivo de OneDrive</h3>
    <p>Ingresa una URL directa de descarga del Excel. Para un OneDrive empresarial, el enlace debe permitir acceso y CORS. Para integrar el selector oficial de OneDrive, configura el File Picker SDK con tu Client ID.</p>
    <div class="form-group"><label for="onedrive-url">URL del archivo</label><input id="onedrive-url" type="url" placeholder="https://.../archivo.xlsx"></div>
    <div class="modal-actions"><button id="cancel-onedrive" class="btn btn-secondary">Cancelar</button><button id="confirm-onedrive" class="btn btn-primary">Conectar</button></div>
  </div>`;
  document.body.appendChild(overlay);
  $('#cancel-onedrive', overlay).onclick = () => overlay.remove();
  $('#confirm-onedrive', overlay).onclick = async () => {
    const url = $('#onedrive-url', overlay).value.trim();
    if (!url) return toast('Ingresa una URL.', 'error');
    overlay.remove();
    await loadOneDriveUrl(url);
  };
}

function protectRoute() {
  const session = getSession();
  if (!session && !['login','register'].includes(state.view)) state.view = 'login';
}

// Inicio de aplicación.
(function init() {
  state.currentUser = getSession();
  state.view = state.currentUser ? 'dashboard' : 'login';
  restoreSavedDataset();
  render();
})();
