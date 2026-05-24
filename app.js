/* ═══════════════════════════════════════════════════════════
   app.js  —  Capstone Studio v2.3
   Updated: schema section now renders as formatted tables
            in the combined .docx download
═══════════════════════════════════════════════════════════ */

/* ── CONSTANTS ──────────────────────────────────────────── */
const USERS = { admin: 'admin123', student: 'capstone2024' };

const SECTORS = [
  ['Healthcare','🏥','Core Industries'],
  ['Finance & Banking','🏦','Core Industries'],
  ['Education & EdTech','🎓','Core Industries'],
  ['Retail & E-Commerce','🛍️','Core Industries'],
  ['Manufacturing','🏭','Core Industries'],
  ['Transportation & Logistics','🚚','Core Industries'],
  ['Energy & Utilities','⚡','Core Industries'],
  ['Telecom','📡','Core Industries'],
  ['Government & Public Services','🏛️','Public Sector'],
  ['Agriculture','🌾','Public Sector'],
  ['Real Estate','🏘️','Public Sector'],
  ['Insurance','🛡️','Public Sector'],
  ['Travel & Hospitality','✈️','Lifestyle'],
  ['Food & Beverage','🍽️','Lifestyle'],
  ['Sports & Fitness','⚽','Lifestyle'],
  ['Food Delivery & Online Ordering','🍕','Lifestyle'],
  ['Media & Entertainment','🎬','Technology'],
  ['Cybersecurity','🔐','Technology'],
  ['IoT & Smart Devices','📱','Technology'],
  ['Social Media & Analytics','📊','Technology'],
  ['Gaming & Esports','🎮','Technology'],
  ['Music & Streaming','🎵','Technology'],
  ['Digital Marketing','📣','Technology'],
  ['Space Research & Satellites','🛸','Science'],
  ['Healthcare Research & Genomics','🧬','Science'],
  ['Weather & Climate Analytics','🌦️','Science'],
  ['Environmental & Sustainability','🌱','Science'],
  ['Pharmaceuticals','💊','Science'],
  ['Automotive','🚗','Engineering'],
  ['Aviation & Aerospace','🛩️','Engineering'],
  ['Mining & Natural Resources','⛏️','Engineering'],
  ['Construction & Infrastructure','🏗️','Engineering'],
  ['Legal & Compliance','⚖️','Professional'],
  ['Human Resources','👥','Professional'],
  ['Supply Chain & Procurement','📦','Professional'],
  ['Event Management & Ticketing','🎫','Professional'],
  ['Charity & Non-Profit','❤️','Social'],
  ['Education Research','📚','Social'],
  ['Cyber Risk Intelligence','🕵️','Social'],
  ['Waste Management & Recycling','♻️','Social'],
];

const ENTITY_MAP = {
  'Healthcare':                      ['Patients','Admissions','Medical_Staff','Billing_Records'],
  'Finance & Banking':               ['Customers','Transactions','Loan_Accounts','Branch_Operations'],
  'Education & EdTech':              ['Students','Courses','Faculty','Enrollment_Records'],
  'Retail & E-Commerce':             ['Products','Orders','Customers','Inventory'],
  'Manufacturing':                   ['Products','Production_Batches','Suppliers','Quality_Checks'],
  'Transportation & Logistics':      ['Shipments','Vehicles','Drivers','Routes'],
  'Energy & Utilities':              ['Meters','Consumption_Records','Grid_Assets','Maintenance_Logs'],
  'Telecom':                         ['Subscribers','Call_Records','Network_Nodes','Service_Plans'],
  'Government & Public Services':    ['Citizens','Service_Requests','Departments','Budget_Allocations'],
  'Agriculture':                     ['Farms','Crop_Yields','Farmers','Weather_Stations'],
  'Insurance':                       ['Policyholders','Claims','Agents','Premium_Records'],
  'Travel & Hospitality':            ['Guests','Bookings','Properties','Reviews'],
  'Food & Beverage':                 ['Products','Suppliers','Sales_Transactions','Inventory'],
  'Automotive':                      ['Vehicles','Customers','Service_Records','Dealers'],
  'Cybersecurity':                   ['Incidents','Assets','Users','Threat_Logs'],
  'Gaming & Esports':                ['Players','Matches','Tournaments','Leaderboards'],
  'Pharmaceuticals':                 ['Drugs','Clinical_Trials','Patients','Regulatory_Records'],
  'Real Estate':                     ['Properties','Agents','Transactions','Listings'],
  'Legal & Compliance':              ['Cases','Clients','Attorneys','Court_Records'],
  'Human Resources':                 ['Employees','Departments','Payroll','Performance_Reviews'],
};

const FILE_FORMATS  = ['CSV','JSON','Excel (.xlsx)','XML / Parquet'];
const FORMATS_SHORT = ['csv','json','xlsx','xml'];
const DOT_CLASSES   = ['dot-csv','dot-json','dot-xlsx','dot-xml'];
const FMT_CLASSES   = ['fmt-csv','fmt-json','fmt-xlsx','fmt-xml'];
const FMT_ICONS     = ['📊','📋','📗','🔷'];

const NON_INDIAN_NAMES = [
  'Marcus Ellsworth','Claire Henderson','Thomas Ridley','Sofia Marchetti',
  'James Whitfield','Emma Larsson','Oliver Brentwood','Laura Fischer',
  'Nathan Calloway','Grace Hartman','Leo Fitzgerald','Anna Svensson',
  'Daniel Forsythe','Victoria Pemberton','Sebastian Mercer','Alice Thornton',
];

const COMPANIES = [
  'NorthBridge Solutions','Apex DataWorks','Vantage Analytics Group',
  'Meridian Tech Partners','Crestline Digital','Summit DataOps',
  'Horizon Infosystems','Pinnacle Analytics','Redwood Data Labs',
];

/* ── STATE ──────────────────────────────────────────────── */
let currentUser    = null;
let selectedSector = null;
let selectedEmoji  = '';
let charLimit      = 8000;
let complexity     = 'medium';
let currentPrompt  = '';
let currentReq     = '';
let currentSol     = '';
let currentSchema  = '';   // now stores OOXML string
let generating     = false;
let sourceMode     = 'demo';
let uploadedFiles  = {};
let demoDataCache  = {};

/* ── HELPERS ────────────────────────────────────────────── */
function rand(arr)       { return arr[Math.floor(Math.random() * arr.length)]; }
function escH(t)         { return t.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function getEntities(s)  { return ENTITY_MAP[s] || ['Entity_A','Entity_B','Entity_C','Entity_D']; }
function randInt(a,b)    { return Math.floor(Math.random()*(b-a+1))+a; }
function randFloat(a,b)  { return +(Math.random()*(b-a)+a).toFixed(2); }

function escXml(str){
  return String(str)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&apos;');
}

function shuffle(arr) {
  const a=[...arr];
  for(let i=a.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[a[i],a[j]]=[a[j],a[i]];}
  return a;
}

/* ══════════════════════════════════════════════════════════
   PURE-JS ZIP / DOCX BUILDER  (no CDN needed)
══════════════════════════════════════════════════════════ */
function makeZipBlob(files){
  const enc = new TextEncoder();

  function crc32(buf){
    const t=new Uint32Array(256);
    for(let i=0;i<256;i++){
      let c=i;
      for(let j=0;j<8;j++) c=c&1?0xEDB88320^(c>>>1):c>>>1;
      t[i]=c;
    }
    let c=0xFFFFFFFF;
    for(let i=0;i<buf.length;i++) c=t[(c^buf[i])&0xFF]^(c>>>8);
    return (c^0xFFFFFFFF)>>>0;
  }

  function w16(v,o,x){ v.setUint16(o,x,true); }
  function w32(v,o,x){ v.setUint32(o,x,true); }

  const locals=[], centrals=[];
  let offset=0;

  for(const [name, data] of files){
    const nb = enc.encode(name);
    const crc = crc32(data);
    const sz = data.length;

    const lh = new Uint8Array(30+nb.length+sz);
    const lv = new DataView(lh.buffer);
    w32(lv,0,0x504B0304); w16(lv,4,20); w16(lv,6,0); w16(lv,8,0);
    w16(lv,10,0); w16(lv,12,0);
    w32(lv,14,crc); w32(lv,18,sz); w32(lv,22,sz);
    w16(lv,26,nb.length); w16(lv,28,0);
    lh.set(nb,30); lh.set(data,30+nb.length);
    locals.push(lh);

    const cd = new Uint8Array(46+nb.length);
    const cv = new DataView(cd.buffer);
    w32(cv,0,0x504B0102); w16(cv,4,20); w16(cv,6,20);
    w16(cv,8,0); w16(cv,10,0); w16(cv,12,0); w16(cv,14,0);
    w32(cv,16,crc); w32(cv,20,sz); w32(cv,24,sz);
    w16(cv,28,nb.length); w16(cv,30,0); w16(cv,32,0);
    w16(cv,34,0); w16(cv,36,0); w16(cv,38,0); w32(cv,42,offset);
    cd.set(nb,46);
    centrals.push(cd);
    offset += lh.length;
  }

  const cdSize = centrals.reduce((s,c)=>s+c.length,0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  w32(ev,0,0x504B0506); w16(ev,4,0); w16(ev,6,0);
  w16(ev,8,files.length); w16(ev,10,files.length);
  w32(ev,12,cdSize); w32(ev,16,offset); w16(ev,20,0);

  const all=[...locals,...centrals,eocd];
  const total=all.reduce((s,c)=>s+c.length,0);
  const out=new Uint8Array(total); let pos=0;
  for(const c of all){ out.set(c,pos); pos+=c.length; }
  return new Blob([out],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}

/* Build a .docx from plain text — each line = one paragraph */
function makeTextDocx(text){
  const enc = new TextEncoder();

  const paras = text.split('\n').map(line=>{
    const x = escXml(line);
    return `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">${x}</w:t></w:r></w:p>`;
  }).join('');

  const docXml =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>${paras}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body>
</w:document>`;

  const relsXml =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const typesXml =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

  const wordRels =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;

  return makeZipBlob([
    ['[Content_Types].xml',           enc.encode(typesXml)],
    ['_rels/.rels',                   enc.encode(relsXml)],
    ['word/document.xml',             enc.encode(docXml)],
    ['word/_rels/document.xml.rels',  enc.encode(wordRels)],
  ]);
}

/* Build a .docx from rich OOXML body string + optional styles */
function makeRichDocx(bodyXml, stylesXml){
  const enc = new TextEncoder();

  const docXml =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>${bodyXml}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080"/></w:sectPr></w:body>
</w:document>`;

  const rootRels =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const wordRels =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

  const typesXml =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

  return makeZipBlob([
    ['[Content_Types].xml',           enc.encode(typesXml)],
    ['_rels/.rels',                   enc.encode(rootRels)],
    ['word/document.xml',             enc.encode(docXml)],
    ['word/styles.xml',               enc.encode(stylesXml)],
    ['word/_rels/document.xml.rels',  enc.encode(wordRels)],
  ]);
}

/* Trigger browser download of a Blob */
function triggerDownload(blob, filename){
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
  showToast('success','✓ Downloading ' + filename);
}

/* ── AUTH ───────────────────────────────────────────────── */
function doLogin() {
  const u=document.getElementById('login-user').value.trim();
  const p=document.getElementById('login-pass').value;
  const errEl=document.getElementById('login-error');
  if(USERS[u]&&USERS[u]===p){
    currentUser=u;
    errEl.classList.remove('show');
    document.getElementById('login-screen').classList.remove('active');
    document.getElementById('app-screen').classList.add('active');
    document.getElementById('user-label').textContent=u;
    buildSectorList();
  }else{
    errEl.classList.add('show');
    document.getElementById('login-pass').value='';
  }
}

document.addEventListener('keydown',e=>{
  if(e.key==='Enter'&&document.getElementById('login-screen').classList.contains('active')) doLogin();
});

function doLogout(){
  currentUser=null; selectedSector=null;
  document.getElementById('app-screen').classList.remove('active');
  document.getElementById('login-screen').classList.add('active');
  document.getElementById('login-user').value='';
  document.getElementById('login-pass').value='';
}

/* ── SOURCE MODE ────────────────────────────────────────── */
function setSourceMode(mode){
  sourceMode=mode;
  document.getElementById('tab-demo').classList.toggle('selected',mode==='demo');
  document.getElementById('tab-upload').classList.toggle('selected',mode==='upload');
  document.getElementById('demo-panel').style.display=mode==='demo'?'block':'none';
  document.getElementById('upload-panel').style.display=mode==='upload'?'block':'none';
  if(selectedSector){
    if(mode==='demo') buildDemoFilesGrid();
    else buildUploadGrid();
  }
  document.getElementById('source-status').textContent=mode==='demo'?'Random data':'Upload files';
  showToast('success',mode==='demo'?'Demo mode — random data will be generated':'Upload mode — add your data files');
}

/* ── SECTOR LIST ────────────────────────────────────────── */
function buildSectorList(filter=''){
  const list=document.getElementById('sector-list');
  list.innerHTML='';
  const cats={};
  SECTORS.forEach(([name,emoji,cat])=>{
    if(filter&&!name.toLowerCase().includes(filter.toLowerCase())) return;
    if(!cats[cat]) cats[cat]=[];
    cats[cat].push([name,emoji]);
  });
  Object.entries(cats).forEach(([cat,items])=>{
    const lbl=document.createElement('div');
    lbl.className='cat-label'; lbl.textContent=cat;
    list.appendChild(lbl);
    items.forEach(([name,emoji])=>{
      const item=document.createElement('div');
      item.className='sector-item'+(selectedSector===name?' selected':'');
      item.innerHTML=`<span class="sector-emoji">${emoji}</span><span>${name}</span>`;
      item.onclick=()=>selectSector(name,emoji);
      list.appendChild(item);
    });
  });
}

function filterSectors(){ buildSectorList(document.getElementById('sector-search').value); }

function selectSector(name,emoji){
  selectedSector=name; selectedEmoji=emoji;
  uploadedFiles={}; demoDataCache={};
  buildSectorList(document.getElementById('sector-search').value);

  const entities=getEntities(name);
  document.getElementById('sector-display').innerHTML=`<span class="sector-badge">${emoji} ${name}</span>`;

  const er=document.getElementById('entity-row');
  er.style.display='flex';
  er.innerHTML=entities.map((e,i)=>
    `<div class="entity-chip">
       <div class="entity-chip-dot ${DOT_CLASSES[i]}"></div>
       ${e} <span style="color:var(--text3);font-size:10px;">.${FORMATS_SHORT[i]}</span>
     </div>`
  ).join('');

  document.getElementById('snum2').classList.add('active');
  document.getElementById('gen-prompt-btn').disabled=false;
  document.getElementById('prompt-preview-wrap').style.display='none';
  document.getElementById('gen-btn').disabled=true;
  document.getElementById('step-download').style.display='none';
  document.getElementById('gen-progress').style.display='none';
  document.getElementById('snum3').className='step-num';
  document.getElementById('char-status').textContent='—';
  document.getElementById('gen-status').textContent='—';
  currentPrompt='';

  if(sourceMode==='demo') buildDemoFilesGrid();
  else buildUploadGrid();
}

/* ══════════════════════════════════════════════════════════
   DEMO DATA GENERATION
══════════════════════════════════════════════════════════ */
const STATUSES=['ACTIVE','INACTIVE','PENDING','ERROR','ARCHIVED'];
const REGIONS=['North','South','East','West','Central','Northeast','Southwest'];
const CATEGORIES=['Premium','Standard','Basic','Enterprise','Starter','Pro','Lite'];
const PRIORITIES=['LOW','MED','HIGH'];
const SEVERITIES=['LOW','MED','HIGH','CRITICAL'];
const EVENT_TYPES=['STATUS_CHANGE','UPDATE','DELETE','CREATE','AUDIT','ALERT'];

function randomName(){ return rand(NON_INDIAN_NAMES.filter(n=>n)); }

function generateCSVData(entityName, n=20){
  const headers=['id','name','status','created_date','region','score','is_active'];
  const rows=[headers.join(',')];
  const nullSlots=shuffle([...Array(n).keys()]).slice(0,3);
  for(let i=0;i<n;i++){
    const hasNull=nullSlots.includes(i);
    const id=1000+i;
    const name=hasNull&&i%3===0?'':randomName();
    const status=hasNull&&i%3===1?'null':rand(STATUSES);
    const date=`2024-${String(randInt(1,12)).padStart(2,'0')}-${String(randInt(1,28)).padStart(2,'0')}`;
    const region=rand(REGIONS);
    const score=hasNull&&i%3===2?'N/A':randFloat(0,100);
    const active=Math.random()>0.3?'true':'false';
    rows.push(`${id},"${name}",${status},${date},${region},${score},${active}`);
  }
  return rows.join('\n');
}

function generateJSONData(entityName, n=20){
  const nullSlots=shuffle([...Array(n).keys()]).slice(0,3);
  const records=[];
  for(let i=0;i<n;i++){
    const hasNull=nullSlots.includes(i);
    records.push({
      record_id:5000+i,
      entity_ref:1000+randInt(0,n-1),
      category:hasNull?null:rand(CATEGORIES),
      value:hasNull?'':+randFloat(100,50000),
      last_updated:`2024-${String(randInt(1,12)).padStart(2,'0')}-${String(randInt(1,28)).padStart(2,'0')}T${String(randInt(0,23)).padStart(2,'0')}:00:00`,
      notes:Math.random()>0.6?'Sample note':'',
    });
  }
  return JSON.stringify(records,null,2);
}

function generateXLSXData(entityName, n=20){
  const headers=['seq_id','description','assigned_to','priority','completion_pct'];
  const rows=[headers.join(',')];
  const nullSlots=shuffle([...Array(n).keys()]).slice(0,3);
  for(let i=0;i<n;i++){
    const hasNull=nullSlots.includes(i);
    rows.push([
      200+i,
      `"Task ${i+1}: ${rand(['Review','Update','Analyse','Validate','Process'])} ${entityName}"`,
      hasNull?'':randomName(),
      hasNull?'':rand(PRIORITIES),
      hasNull?'N/A':randFloat(0,100),
    ].join(','));
  }
  return rows.join('\n');
}

function generateXMLData(entityName, n=20){
  const nullSlots=shuffle([...Array(n).keys()]).slice(0,3);
  let xml=`<?xml version="1.0" encoding="UTF-8"?>\n<records>\n`;
  for(let i=0;i<n;i++){
    const hasNull=nullSlots.includes(i);
    const ts=`2024-${String(randInt(1,12)).padStart(2,'0')}-${String(randInt(1,28)).padStart(2,'0')} ${String(randInt(0,23)).padStart(2,'0')}:${String(randInt(0,59)).padStart(2,'0')}:00`;
    xml+=`  <record>
    <log_id>${9000+i}</log_id>
    <source_ref>${1000+randInt(0,n-1)}</source_ref>
    <event_type>${hasNull?'':rand(EVENT_TYPES)}</event_type>
    <timestamp>${ts}</timestamp>
    <severity>${hasNull?'':rand(SEVERITIES)}</severity>
  </record>\n`;
  }
  xml+=`</records>`;
  return xml;
}

function generateEntityData(entityName, formatIndex){
  if(formatIndex===0) return { content:generateCSVData(entityName), ext:'csv', mime:'text/csv' };
  if(formatIndex===1) return { content:generateJSONData(entityName), ext:'json', mime:'application/json' };
  if(formatIndex===2) return { content:generateXLSXData(entityName), ext:'csv', mime:'text/csv' };
  return { content:generateXMLData(entityName), ext:'xml', mime:'application/xml' };
}

/* ── BUILD DEMO FILES GRID ─────────────────────────────── */
function buildDemoFilesGrid(){
  if(!selectedSector) return;
  const entities=getEntities(selectedSector);
  const grid=document.getElementById('demo-files-grid');
  grid.style.display='grid';
  demoDataCache={};
  grid.innerHTML=entities.map((eName,i)=>{
    const {content,ext}=generateEntityData(eName,i);
    demoDataCache[eName]={content,ext,formatIndex:i};
    const safeName=eName.toLowerCase().replace(/_/g,'-');
    const previewLine=content.split('\n').slice(0,3).join('\n');
    const records=content.split('\n').length-1;
    return `<div class="demo-file-card">
      <div class="demo-file-header">
        <span class="demo-file-icon">${FMT_ICONS[i]}</span>
        <span class="demo-file-name">${safeName}.${ext}</span>
        <span class="demo-file-format ${FMT_CLASSES[i]}">${FORMATS_SHORT[i].toUpperCase()}</span>
      </div>
      <div class="demo-file-meta">~${records} sample records · ${content.length.toLocaleString()} chars</div>
      <div class="demo-file-preview">${escH(previewLine)}</div>
      <button class="btn-dl-demo" onclick="downloadDemoFile('${eName}')">⬇ Download ${ext.toUpperCase()}</button>
    </div>`;
  }).join('');
}

function downloadDemoFile(entityName){
  const d=demoDataCache[entityName];
  if(!d){showToast('error','Generate demo files first');return;}
  const safeName=entityName.toLowerCase().replace(/_/g,'-');
  const blob=new Blob([d.content],{type:'text/plain;charset=utf-8'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download=`${safeName}_sample.${d.ext}`;
  a.click();
  URL.revokeObjectURL(a.href);
  showToast('success',`⬇ Downloading ${safeName}_sample.${d.ext}`);
}

/* ── BUILD UPLOAD GRID ──────────────────────────────────── */
function buildUploadGrid(){
  if(!selectedSector) return;
  const entities=getEntities(selectedSector);
  const grid=document.getElementById('upload-grid');
  grid.innerHTML=entities.map((eName,i)=>{
    const accept=i===0?'.csv':i===1?'.json':i===2?'.xlsx,.xls,.csv':'.xml,.parquet,.csv';
    return `<div class="upload-slot" id="slot-${eName}">
      <div class="upload-slot-header">
        <span class="entity-chip-dot ${DOT_CLASSES[i]}" style="width:8px;height:8px;border-radius:50%;flex-shrink:0;"></span>
        <span class="upload-slot-entity">${eName}</span>
        <span class="upload-slot-format ${FMT_CLASSES[i]}">${FORMATS_SHORT[i].toUpperCase()}</span>
      </div>
      <div class="upload-drop-area" onclick="document.getElementById('file-input-${eName}').click()">
        <input type="file" id="file-input-${eName}" accept="${accept}" onchange="handleFileUpload('${eName}',this)">
        <div class="upload-drop-text">📁 <span>Browse file</span> or drop here<br>Accepts: ${accept}</div>
      </div>
      <div class="upload-file-info" id="file-info-${eName}">
        <div class="upload-filename" id="fname-${eName}">—</div>
        <div class="upload-filesize" id="fsize-${eName}">—</div>
        <div class="upload-file-actions">
          <button class="btn-upload-action" onclick="previewUploadedFile('${eName}')">👁 Preview</button>
          <button class="btn-upload-action danger" onclick="removeUploadedFile('${eName}')">✕ Remove</button>
        </div>
      </div>
    </div>`;
  }).join('');
}

function handleFileUpload(entityName, input){
  const file=input.files[0];
  if(!file) return;
  uploadedFiles[entityName]=file;
  const slot=document.getElementById(`slot-${entityName}`);
  slot.classList.add('has-file');
  document.getElementById(`fname-${entityName}`).textContent=file.name;
  document.getElementById(`fsize-${entityName}`).textContent=`${(file.size/1024).toFixed(1)} KB · ${file.type||'unknown type'}`;
  showToast('success',`✓ ${file.name} uploaded for ${entityName}`);
  checkUploadComplete();
}

function removeUploadedFile(entityName){
  delete uploadedFiles[entityName];
  const slot=document.getElementById(`slot-${entityName}`);
  slot.classList.remove('has-file');
  const inp=document.getElementById(`file-input-${entityName}`);
  if(inp) inp.value='';
  showToast('success',`Removed file for ${entityName}`);
  checkUploadComplete();
}

function previewUploadedFile(entityName){
  const file=uploadedFiles[entityName];
  if(!file){showToast('error','No file uploaded');return;}
  const reader=new FileReader();
  reader.onload=e=>{
    const preview=e.target.result.substring(0,500);
    alert(`Preview: ${file.name}\n\n${preview}${e.target.result.length>500?'\n\n[truncated...]':''}`);
  };
  reader.readAsText(file);
}

function checkUploadComplete(){
  if(!selectedSector) return;
  const entities=getEntities(selectedSector);
  const count=entities.filter(e=>uploadedFiles[e]).length;
  document.getElementById('source-status').textContent=`${count}/${entities.length} files`;
}

/* ── COMPLEXITY ─────────────────────────────────────────── */
function selectComplexity(level, el){
  complexity=level;
  document.querySelectorAll('.complexity-btn').forEach(b=>b.classList.remove('selected'));
  el.classList.add('selected');
  if(currentPrompt) generatePrompt();
}

/* ── CHAR LIMIT ─────────────────────────────────────────── */
function selectCharLimit(n,el){
  charLimit=n;
  document.querySelectorAll('.char-btn').forEach(b=>b.classList.remove('selected'));
  el.classList.add('selected');
  if(currentPrompt) generatePrompt();
}

/* ══════════════════════════════════════════════════════════
   COMPLEXITY MODIFIERS
══════════════════════════════════════════════════════════ */
const COMPLEXITY_META={
  easy:{
    label:'Easy',pillClass:'pill-easy',icon:'🟢',
    desc:'Introductory-level. Basic read/write, simple SELECT queries, single-file scripts.',
    modifier:`
COMPLEXITY DIRECTIVE: EASY
Generate ONLY introductory-level user stories. Requirements should be:
• Python: single-file scripts using csv / json modules, basic read/write, simple loops
• SQL: simple SELECT with WHERE, INSERT, UPDATE, DELETE — no joins, no subqueries
• MongoDB: insertOne, findOne, find() with simple filter — no aggregation pipelines
• PySpark: load CSV, printSchema(), basic filter() and show() — no transformations
• Power BI: simple bar chart, pie chart, card visual; 2-3 measures (COUNT, SUM, AVERAGE)
• Unix: grep, wc -l, find, cut — single-command operations
Label stories: [EASY] at start of each US-xxx entry.
`,
  },
  medium:{
    label:'Medium',pillClass:'pill-medium',icon:'🟡',
    desc:'Moderate level. Simple joins, basic groupBy, pandas cleaning, basic DAX.',
    modifier:`
COMPLEXITY DIRECTIVE: MEDIUM
Generate moderate-difficulty user stories — basic to slightly complex only:
• Python: pandas read_csv, dropna, fillna, drop_duplicates, groupby with simple agg
• SQL: 2-table JOIN, GROUP BY, ORDER BY, simple stored procedure with one parameter
• MongoDB: find() with multiple filters, simple $group aggregation, updateMany()
• PySpark: read CSV/JSON, select(), filter(), withColumn(), simple groupBy().count()
• Power BI: 4-5 DAX measures (COUNT, SUM, AVERAGE, CALCULATE with one filter, DIVIDE)
• Unix: grep with pipe to wc, find with -name, cut + sort + head
Label stories: [MEDIUM] at start of each US-xxx entry.
`,
  },
  hard:{
    label:'Hard',pillClass:'pill-hard',icon:'🔴',
    desc:'Moderate-advanced. Multi-table joins, groupBy with agg, $lookup, window basics.',
    modifier:`
COMPLEXITY DIRECTIVE: HARD (Moderate-Advanced — NOT extreme)
Generate practical, moderately advanced user stories — keep them achievable for students:
• Python: pandas multi-step cleaning, merge two DataFrames, groupby + agg, write to CSV/JSON
• SQL: 2-3 table JOINs with GROUP BY and HAVING, stored procedure with parameter + error handling
• MongoDB: $group + $sort + $project pipeline, $lookup for simple join, updateMany with condition
• PySpark: DataFrame groupBy().agg(), join two DataFrames, withColumn() with when/otherwise
• Power BI: 5-6 DAX measures including CALCULATE with filter, RANKX on one column, DIVIDE
• Unix: grep + cut + sort pipeline, find with -mtime, wc -l with tee
Label stories: [HARD] at start of each US-xxx entry.
`,
  },
  mixed:{
    label:'Mixed',pillClass:'pill-mixed',icon:'🔵',
    desc:'Progressive: starts very basic, ends at moderate level — no extreme complexity.',
    modifier:`
COMPLEXITY DIRECTIVE: MIXED (Basic to Moderate — progressive)
Generate a progressive curve from easy to moderate only — NO advanced topics:
• Data Ingestion: [EASY] basic file copy and read; [MEDIUM] add validation and logging
• Unix Commands: [EASY] grep, wc -l, find; [MEDIUM] pipe combinations (grep + cut + sort)
• Python Scripting: [EASY] read CSV, print rows; [MEDIUM] groupby, write cleaned output
• Python File Handling: [EASY] open/read/write; [MEDIUM] try-except, DictReader/DictWriter
• PySpark Cleaning: [EASY] load and show schema; [MEDIUM] filter nulls, drop duplicates
• MongoDB CRUD: [EASY] insertOne, find(); [MEDIUM] updateMany, simple $group pipeline
• PySpark Analysis: [EASY] basic filter + count; [MEDIUM] groupBy + agg, simple join
• Advanced SQL: [EASY] SELECT + WHERE; [MEDIUM] JOIN + GROUP BY + stored procedure
• Power BI: [EASY] bar/card visuals; [MEDIUM] CALCULATE + DIVIDE measures
Label every story: [EASY] or [MEDIUM] only — do NOT use [HARD].
`,
  },
};

/* ══════════════════════════════════════════════════════════
   PROMPT BUILDER
══════════════════════════════════════════════════════════ */
function buildPrompt(sector, entities, limit){
  const names=shuffle(NON_INDIAN_NAMES).slice(0,3);
  const company=rand(COMPANIES);
  const now=new Date().toISOString().slice(0,16).replace('T',' ');
  const e0=entities[0],e1=entities[1],e2=entities[2],e3=entities[3];
  const cmeta=COMPLEXITY_META[complexity];

  const header=
`=== CAPSTONE PROJECT REQUIREMENT GENERATOR ===
Sector      : ${sector}
Company     : ${company}
Generated   : ${now}
Char Limit  : ${limit.toLocaleString()}
Complexity  : ${cmeta.label.toUpperCase()} ${cmeta.icon}

BUSINESS CONTEXT
${company} is a mid-sized technology consulting firm engaged with a leading
${sector} organisation to modernise their fragmented data infrastructure.
Stakeholders: ${names[0]} (Chief Data Officer), ${names[1]} (VP Operations),
${names[2]} (Lead Data Engineer). Project spans 6 Agile sprints.
${cmeta.desc}

ENTITIES & FORMATS
1. ${e0.padEnd(32)} → CSV
2. ${e1.padEnd(32)} → JSON
3. ${e2.padEnd(32)} → Excel (.xlsx)
4. ${e3.padEnd(32)} → XML / Parquet

DATA SPECS
• Each dataset: 80,000 – 1,00,000 records
• Each dataset: 15–20 intentional inconsistencies
• All person names: non-Indian (Western / European only)
• No real registered company or organisation names
${cmeta.modifier}
`;

  const sections=[

`━━━ DATA INGESTION AND SHELL SCRIPTING ━━━
US-ING-01: As a data engineer, I want to ingest data from multiple formats so that data from different sources can be unified into a single processing pipeline.
US-ING-02: As a system administrator, I want to automate ingestion workflows so that data pipelines run without manual intervention and produce repeatable results.
US-ING-03: As a DevOps engineer, I want to schedule periodic data loading so that the system processes new data at defined intervals without human trigger.
US-ING-04: As a data architect, I want to store ingested data in a staging
  layer with a source_file and load_timestamp column so that raw records
  can be audited before entering the cleansed layer.
US-ING-05: As a developer, I want to validate file structure and schema
  during ingestion so that malformed files are rejected
  early and an error report is written to ingest_errors.log.

`,

`━━━ UNIX COMMANDS — 15 Story-Based Requirements ━━━
US-UNX-01: As a data analyst, I want to extract all rows from ${e0.toLowerCase()}.csv where status matches ERROR or PENDING so
  that problematic records are isolated before the ETL pipeline runs.
US-UNX-02: As a data engineer, I want to compute the average score per region from ${e0.toLowerCase()}.csv so that a regional KPI summary is
  produced without loading data into any database.
US-UNX-03: As a system admin, I want to replace all N/A, null, and -- values with 0 in numeric fields of ${e1.toLowerCase()}.txt so that
  downstream processes do not fail on non-numeric strings.
US-UNX-04: As a reporting analyst, I want to rank the top 20 category values by frequency from ${e2.toLowerCase()}.csv so that they are ready for filter configuration.
US-UNX-05: As a DevOps engineer, I want to build a daily manifest.txt of recently changed files so that the backup script processes only modified data files.
US-UNX-06: As a data steward, I want record counts for all 4 entity files to be printed to the console and saved to size_report.txt simultaneously.
US-UNX-07: As a data engineer, I want to compare ${e0.toLowerCase()}_v1.csv
  and ${e0.toLowerCase()}_v2.csv so that records changed between two pipeline
  runs are identified and flagged automatically.
US-UNX-08: As a pipeline engineer, I want to emit a unified schema report from all 4 files so that schema drift is detected before any transformation step.
US-UNX-09: As a data analyst, I want to count distinct values in the last column of
  ${e1.toLowerCase()}.txt so that field cardinality is documented per entity.
US-UNX-10: As a QA engineer, I want to compute null density as a percentage per entity file so that a quality threshold check can gate the pipeline run automatically.
US-UNX-11: As a systems engineer, I want to list files exceeding 10 MB and redirect to large_files.txt so that storage alerts are triggered before disk quotas are breached.
US-UNX-12: As a data analyst, I want to extract top-scoring records from ${e0.toLowerCase()}.csv to a flat file without any database dependency.
US-UNX-13: As a DevOps engineer, I want to normalise all category strings in ${e1.toLowerCase()}.txt
  to lowercase so that case-inconsistent values are unified before loading.
US-UNX-14: As a pipeline engineer, I want to prepend a source_entity column to every record from all 4 files so that combined_raw.csv carries its originating file label.
US-UNX-15: As a DevOps engineer, I want to run inconsistency detection in parallel across all 4 entity files so that it completes in under 30 seconds on large datasets.

`,

`━━━ SHELL SCRIPTING — 10 Story-Based Requirements ━━━
US-SH-01: As a data steward, I want a script that iterates all 4 entity files, counts total rows, null rows, and duplicate rows, and writes a formatted audit_report.txt so that data quality KPIs are tracked daily.
US-SH-02: As an operations analyst, I want a script that reads ${e1.toLowerCase()}.txt and computes sum, avg, max of value per category, printing an ASCII KPI table so that metrics are reviewed without any BI tool.
US-SH-03: As a data engineer, I want a script that flags rows where score falls outside 0–100 in ${e0.toLowerCase()}.csv and writes them to anomalies.csv with an appended reason column so that outliers are quarantined.
US-SH-04: As a pipeline engineer, I want a script that merges all 4 entity files into combined_dataset.csv with a prepended source_entity column so that a unified flat file is ready for batch processing.
US-SH-05: As a reporting analyst, I want a script that generates a daily HTML summary of record counts, top 5 categories, and null percentages so that stakeholders receive an email-ready overview each morning.
US-SH-06: As a data engineer, I want a script that accepts a date argument and filters ${e0.toLowerCase()}.csv by created_date, writing to dated_extract_YYYYMMDD.csv so that point-in-time slices are produced on demand.
US-SH-07: As a QA analyst, I want a script that compares row counts of raw vs cleaned files and outputs a cleaning efficiency report showing records_removed and pct_cleaned per entity so that data-loss is audited.
US-SH-08: As a pipeline engineer, I want a script that monitors a landing directory and on new file arrival triggers the ingestion pipeline and logs the event with a timestamp so that the pipeline is fully event-driven.
US-SH-09: As a data engineer, I want a script that removes rows with empty mandatory fields from ${e0.toLowerCase()}.csv, writes cleaned output to _clean.csv, and logs the removed count so that ETL receives only complete records. (Data-cleaning story 1)
US-SH-10: As a DevOps engineer, I want a script that fixes delimiter inconsistencies in ${e2.toLowerCase()}.csv, normalising all separators to commas, so that the file conforms to CSV specification before downstream processing. (Data-cleaning story 2)

`,

`━━━ FILE HANDLING ━━━
US-PY-01: As a Python developer, I want to implement file handling to read
  and write ${e0.toLowerCase()}.csv so that pre-processing and deduplication
  can be performed efficiently.
  INPUT: data/${e0.toLowerCase()}.csv  OUTPUT: data/${e0.toLowerCase()}_clean.csv
  SCOPE: Remove duplicates by id, fix encoding, drop null mandatory fields.
US-PY-02: As a data engineer, I want to process ${e1.toLowerCase()}.json
  (80k–1L records) so that nested structures are flattened and saved as
  ${e1.toLowerCase()}_flat.csv for database import.
  INPUT: data/${e1.toLowerCase()}.json  OUTPUT: data/${e1.toLowerCase()}_flat.csv
US-PY-03: As a QA tester, I want to handle file-related exceptions such as
  missing or corrupted ${e2.toLowerCase()}.xlsx files using try-except so
  that the pipeline logs a descriptive error and skips rather than crashing.
US-PY-04: As a developer, I want to generate cleaned output files for all
  3 datasets in their required formats (CSV, JSON, XLSX) so that each
  cleaned file is immediately usable by downstream PySpark and SQL modules.
US-PY-05: As a data engineer, I want a Python script that randomly selects
  a sector and entity at runtime, reads the corresponding dataset, applies
  cleaning rules, and generates sector-specific user stories so that each
  execution produces unique, non-repeating requirements.

`,

`━━━ DATA CLEANING AND INCONSISTENCIES ━━━
US-PS-01: As a data quality analyst, I want to use PySpark to identify null,
  missing, and duplicate values across all 4 entity datasets so that overall
  data quality improves before analysis begins.
US-PS-02: As a data engineer, I want to standardise inconsistent date, text,
  and numerical formats in ${e0.toLowerCase()} so that data becomes uniform across all regions.
US-PS-03: As a compliance analyst, I want to detect and resolve 10 to 15
  inconsistencies in each dataset so that reporting remains accurate.
US-PS-04: As a data engineer, I want to apply business-rule transformation
  logic — normalising status enums, capping out-of-range scores, and
  forward-filling sparse columns — so that downstream analysis operates on validated data.
US-PS-05: As a data analyst, I want cleaned PySpark DataFrames to be
  validated with row-count assertions and null-count checks before writing
  to the cleansed layer so that unreliable data never reaches the BI layer.
US-PS-06: As a data engineer, I want to verify that the cleansed
  ${e0.toLowerCase()} DataFrame can produce syntactically valid MySQL INSERT
  statements via a PySpark UDF so that database loading is confirmed before
  the full batch is submitted.

`,

`━━━ CRUD OPERATIONS — 10+ Stories ━━━
US-MG-01: As a database administrator, I want to implement create operations
  using insertMany() with a compound index on status and region for
  ${e0.toLowerCase()} so that bulk loads complete and filtered reads stay under 100 ms.
US-MG-02: As a backend developer, I want to implement read operations using
  find() with projection and sort on ${e1.toLowerCase()} so that only
  required fields are returned and network payload is minimised.
US-MG-03: As an operations manager, I want to update records dynamically
  using updateMany() to set status='ARCHIVED' on ${e0.toLowerCase()} records
  older than 2 years so that business data remains current.
US-MG-04: As a DBA, I want to implement deleteMany() to permanently remove
  low-severity log entries from ${e3.toLowerCase()} older than 3 years so
  that storage costs are bounded without manual purge scripts.
US-MG-05: As a backend developer, I want to perform aggregation using
  $group, $sort, $project on ${e1.toLowerCase()} to compute sum, avg, and
  count per category so that KPI transformations are applied before the BI layer.
US-MG-06: As a data engineer, I want to use $lookup to join
  ${e0.toLowerCase()} with ${e1.toLowerCase()} on id/entity_ref and write
  the result to a combined_view collection so that analytics have a
  denormalised view without repeated joins.
US-MG-07: As a data engineer, I want to use $unwind on details array in
  ${e3.toLowerCase()} followed by $group to count distinct event sub-types
  so that event taxonomy is documented automatically.
US-MG-08: As a reporting analyst, I want to use $facet to produce three
  sub-aggregations in a single query so that dashboard APIs use one round-trip.
US-MG-09: As a data steward, I want a MongoDB aggregation that groups
  ${e0.toLowerCase()} on name and region where count > 1 and writes
  results to review_duplicates so that deduplication candidates are isolated.
US-MG-10: As a DBA, I want to create a TTL index on the timestamp field of
  ${e3.toLowerCase()} that auto-deletes documents older than 180 days so
  that log collection size is self-managing.
US-MG-11: As a backend developer, I want to use $lookup with $unwind and
  $replaceRoot to flatten nested references in ${e2.toLowerCase()} so that
  a denormalised document suitable for export is produced in one pipeline.

`,

`━━━ DATA PROCESSING AND ANALYSIS ━━━
[RDD]
US-RDD-01: As a data engineer, I want to load ${e0.toLowerCase()}.csv and
  remove null mandatory fields so that only valid records enter the
  distributed processing stage.
US-RDD-02: As a data engineer, I want to count records by status in
  ${e0.toLowerCase()} so that status distribution is computed efficiently.
US-RDD-03: As a data analyst, I want to compute per-partition statistics
  for ${e1.toLowerCase()} so that large datasets are summarised efficiently
  in a distributed manner.
US-RDD-04: As a data engineer, I want to combine multiple entity datasets
  so that cross-entity analysis can be performed in a single distributed pass.
US-RDD-05: As a business analyst, I want to compute sum and count per
  category from ${e1.toLowerCase()} so that weighted averages are derived
  without loading into SQL.

[SQL]
US-SQL-01: As a data analyst, I want to run a multi-table JOIN computing
  avg score, total value, and event count per region across all 4 entity
  views so that cross-entity insights are generated efficiently.
US-SQL-02: As a data analyst, I want to use window functions RANK() and
  LAG() partitioned by region to identify month-over-month trend changes
  in ${e0.toLowerCase()} so that time-series analysis is performed declaratively.
US-SQL-03: As a business analyst, I want to produce subtotal and grand-total
  aggregations from ${e1.toLowerCase()} so that hierarchical summaries are
  generated in one statement.
US-SQL-04: As a data engineer, I want to classify ${e0.toLowerCase()}
  records into score buckets so that downstream reports consume pre-labelled data.
US-SQL-05: As a data analyst, I want to query all 4 entity views with CTE
  expressions so that meaningful cross-entity insights are generated efficiently.

[DATAFRAME]
US-DF-01: As a data engineer, I want to join ${e0.toLowerCase()} with
  ${e1.toLowerCase()} on id and apply window functions partitioned by region
  so that enriched records are written to the cleansed layer.
US-DF-02: As a data engineer, I want to apply transformations on
  ${e2.toLowerCase()} so that performance is optimised and a clean schema
  is enforced.
US-DF-03: As a data analyst, I want to compute multiple statistics on
  ${e1.toLowerCase()} in a single pass so that KPI tables are generated
  without multiple scans.
US-DF-04: As a data engineer, I want to combine all 4 entity datasets
  after adding a source_entity column so that a single wide dataset is
  available for cross-entity analysis.
US-DF-05: As a business analyst, I want summarised insights written as
  partitioned files so that reporting layers consume pre-aggregated results
  and render dashboards quickly.

`,

`━━━ ADVANCED SQL — 15 Story-Based Requirements ━━━
US-ASQL-01: As a database developer, I want to create stored procedure
  sp_GetRegionalSummary(p_region VARCHAR) returning total records, avg score,
  active count, and top 3 categories via a 3-table JOIN so that the reporting
  layer calls a single reusable procedure per region.
US-ASQL-02: As a database developer, I want to create stored function
  fn_CalculateRiskScore(p_id INT) computing a weighted composite score from
  3 entity tables so that complex calculations are simplified to a single callable function.
US-ASQL-03: As a data analyst, I want stored procedure sp_MonthlyKPIReport
  (p_year INT, p_month INT) using RANK(), LAG(), LEAD() across all 4 entities
  so that month-over-month trends are generated consistently.
US-ASQL-04: As a database developer, I want stored procedure sp_CleanDuplicates
  (p_table VARCHAR) using CTE with ROW_NUMBER() PARTITION BY name, region to
  delete all but the newest duplicate so that deduplication is auditable.
US-ASQL-05: As a database developer, I want stored function fn_GetStatusLabel
  (p_score FLOAT) returning EXCELLENT / GOOD / AVERAGE / POOR / CRITICAL via
  a CASE expression so that application code avoids duplicating business label logic.
US-ASQL-06: As a data analyst, I want stored procedure sp_TopNByRegion
  (p_region VARCHAR, p_n INT) using DENSE_RANK() to return top-N records
  by score for a region so that leaderboard queries are parameterised.
US-ASQL-07: As a DBA, I want stored procedure sp_ArchiveOldRecords
  (p_cutoff_date DATE) wrapped in a transaction with ROLLBACK on error so
  that archival is safe and auditable.
US-ASQL-08: As a data engineer, I want stored function fn_ComputeGrowthPct
  (p_entity_id INT, p_period VARCHAR) querying current and previous period
  values from ${e1.toLowerCase()} so that growth percentage is encapsulated.
US-ASQL-09: As a reporting analyst, I want stored procedure
  sp_GenerateCrossEntityReport() using multi-CTE joins across all 4 tables
  with PIVOT-style CASE so that a wide summary is produced consistently.
US-ASQL-10: As a backend developer, I want stored procedure
  sp_GetDuplicateSummary() returning duplicate counts per entity table
  using information_schema so that data quality status is queryable at any time.
US-ASQL-11: As a data analyst, I want stored function fn_Normalize
  (p_val FLOAT, p_min FLOAT, p_max FLOAT) computing min-max normalisation
  so that score fields across entities are compared on a 0–1 scale.
US-ASQL-12: As a DBA, I want stored procedure sp_RebuildIndexes() that uses
  dynamic SQL to ANALYZE TABLE across all 4 entity tables so that optimiser
  statistics are refreshed in one call.
US-ASQL-13: As a pipeline engineer, I want stored procedure sp_ValidateSchema
  (p_table VARCHAR) checking null counts in NOT NULL columns and out-of-range
  values using INFORMATION_SCHEMA so that pre-load validation is fully automated.
US-ASQL-14: As a reporting analyst, I want stored procedure sp_HeatmapData
  (p_dim1 VARCHAR, p_dim2 VARCHAR) grouping ${e0.toLowerCase()} by two
  dimensions using a CASE-based pivot so that Power BI heatmaps receive
  pre-aggregated data.
US-ASQL-15: As a database developer, I want stored function fn_FormatCurrency
  (p_val DECIMAL) returning a locale-formatted VARCHAR so that all financial
  display values in reports are consistently formatted by the database layer.

`,

`━━━ POWER BI TRANSFORMATIONS AND VISUALIZATION ━━━
US-PBI-01: As a BI developer, I want to perform data transformation using
  Power Query on all 4 entity datasets applying 15 steps so that a clean
  star schema loads automatically on every refresh.
US-PBI-02: As a data analyst, I want to create 10 calculated measures using
  DAX — Total_Records, Avg_Score, Running_Total, MoM_Growth_Pct (DATEADD),
  Rank_By_Region (RANKX), YTD_Value (TOTALYTD), Pct_Share (DIVIDE+ALL),
  Rolling_30d_Avg (DATESINPERIOD), KPI_Status (SWITCH+TRUE()),
  Top_N_Dynamic (TOPN+RANKX) — so that all dashboard visuals share a
  consistent, centralised calculation layer.
US-PBI-03: As a business stakeholder, I want interactive dashboards with
  5 pages — Executive KPI cards, Regional bar+line combo with drill-through,
  Category decomposition tree, Detailed matrix with conditional formatting,
  and a Time-series trend with forecast — so that KPIs can be monitored
  at the right level of granularity.
US-PBI-04: As a reporting analyst, I want to structure ${e0.toLowerCase()}
  and ${e1.toLowerCase()} data by pre-aggregating in Power Query and building
  a star schema with a Calendar dimension so that reports are clear and performant.

`,

`━━━ AUTOMATION AND PROJECT GENERATION LOGIC ━━━
US-AUT-01: As a developer, I want the system to dynamically generate
  datasets, sectors, and story-based requirements at runtime so that each
  execution produces unique, non-repeating outputs.
US-AUT-02: As a developer, I want requirements and solutions to be generated
  automatically so that the process is efficient and consistent.
US-AUT-03: As a developer, I want the generated document to include project
  title, business scenario, stakeholder names, entity schemas, requirements
  with solutions, and an AGILE sprint sheet so that the deliverable is
  complete and submission-ready.

CONSTRAINT: All requirements must be USER STORIES:
"As a [role], I want to [action] so that [benefit]."
Character limit for this prompt: ${limit.toLocaleString()}
`,
  ];

  let prompt=header;
  for(const section of sections){
    if((prompt+section).length<=limit-80) prompt+=section;
  }
  const padding=`\n\nNOTE: All requirements above are sector-specific to "${sector}" and ` +
    `reference the 4 entities listed. Complexity level: ${COMPLEXITY_META[complexity].label.toUpperCase()}. ` +
    `Solutions must include working code (bash, Python, MongoDB shell, MySQL, PySpark). ` +
    `Ensure each user story demonstrates measurable business value at the specified complexity level.`;
  if((prompt+padding).length<=limit) prompt+=padding;
  return prompt.slice(0,limit);
}

/* ── GENERATE PROMPT (Step 2) ───────────────────────────── */
function generatePrompt(){
  if(!selectedSector) return;
  const entities=getEntities(selectedSector);
  currentPrompt=buildPrompt(selectedSector,entities,charLimit);

  const cmeta=COMPLEXITY_META[complexity];
  document.getElementById('complexity-badge-display').innerHTML=
    `<span class="complexity-pill ${cmeta.pillClass}">${cmeta.icon} ${cmeta.label}</span>
     <span class="complexity-badge-label">${cmeta.desc}</span>`;

  document.getElementById('prompt-preview').textContent=currentPrompt;
  document.getElementById('prompt-preview-wrap').style.display='block';

  const len=currentPrompt.length;
  const pct=Math.round(len/charLimit*100);
  const countEl=document.getElementById('char-count-display');
  countEl.textContent=`${len.toLocaleString()} / ${charLimit.toLocaleString()} chars (${pct}%)`;
  countEl.className='char-count '+(len<=charLimit?'ok':'over');
  document.getElementById('char-status').textContent=`${len.toLocaleString()} chars`;

  document.getElementById('gen-btn').disabled=false;
  document.getElementById('snum3').classList.add('active');
  showToast('success',`✦ Prompt ready — ${len.toLocaleString()} chars · ${cmeta.label} complexity`);
}

function copyPrompt(){
  if(!currentPrompt) return;
  navigator.clipboard.writeText(currentPrompt)
    .then(()=>showToast('success','Prompt copied to clipboard'));
}

/* ── PROGRESS BAR ────────────────────────────────────────── */
const PROGRESS_STEPS=[
  'Connecting to Claude API…',
  'Sending sector + entity context…',
  'Generating business scenario…',
  'Building data ingestion user stories…',
  'Writing Unix & shell scripting stories…',
  'Generating Python file handling stories…',
  'Building PySpark cleaning & analysis stories…',
  'Writing MongoDB CRUD stories…',
  'Generating Advanced SQL stored procedures…',
  'Building Power BI transformation stories…',
  'Compiling technical solutions with code…',
  'Finalising output documents…',
];

let progressInterval=null,progressStep=0;

function startProgress(){
  const stepsEl=document.getElementById('progress-steps');
  stepsEl.innerHTML=PROGRESS_STEPS.map((s,i)=>
    `<div class="progress-step" id="pstep-${i}"><div class="step-dot"></div><span>${s}</span></div>`
  ).join('');
  progressStep=0;
  document.getElementById('progress-fill').style.width='0%';
  progressInterval=setInterval(()=>{
    if(progressStep>0){
      const prev=document.getElementById(`pstep-${progressStep-1}`);
      if(prev){prev.classList.remove('active');prev.classList.add('done');}
    }
    if(progressStep<PROGRESS_STEPS.length){
      const cur=document.getElementById(`pstep-${progressStep}`);
      if(cur) cur.classList.add('active');
      document.getElementById('progress-label').textContent=PROGRESS_STEPS[progressStep];
      document.getElementById('progress-fill').style.width=
        ((progressStep+1)/PROGRESS_STEPS.length*90)+'%';
      progressStep++;
    }
  },650);
}

function stopProgress(){
  if(progressInterval) clearInterval(progressInterval);
  document.getElementById('progress-fill').style.width='100%';
  document.querySelectorAll('.progress-step').forEach(s=>{s.classList.remove('active');s.classList.add('done');});
}

/* ── TOAST ───────────────────────────────────────────────── */
function showToast(type,msg){
  const t=document.getElementById('toast');
  t.textContent=msg; t.className=`toast ${type} show`;
  setTimeout(()=>t.classList.remove('show'),3500);
}

/* ══════════════════════════════════════════════════════════
   GENERATE — main action
══════════════════════════════════════════════════════════ */
async function generate(){
  if(!selectedSector||generating||!currentPrompt) return;
  generating=true;
  const entities=getEntities(selectedSector);
  const apiKey=document.getElementById('api-key-input').value.trim();
  const cmeta=COMPLEXITY_META[complexity];

  document.getElementById('gen-btn').disabled=true;
  document.getElementById('gen-btn-inner').innerHTML='<div class="spinner"></div> Generating…';
  document.getElementById('gen-progress').style.display='block';
  document.getElementById('step-download').style.display='none';
  startProgress();

  try{
    let reqText='',solText='';

    if(apiKey){
      const systemPrompt=
`You are an expert capstone project document generator for data engineering students.
Generate a complete, professional capstone project document that exactly matches this structure and style.
Respond ONLY with valid JSON — no markdown fences, no preamble.
JSON must have exactly two keys: "requirements" and "solutions".

DOCUMENT STRUCTURE (follow this exact order in "requirements"):
1. Project Title
2. Business Scenario (plain narrative paragraphs + bullet lists)
3. Dataset Source Overview (pipe-table format)
4. Schema Details for each entity (pipe-table, 3 columns: SL No | Column Name | Description)
5. Implementation sections in this order:
   - Implement requirements using Unix Commands.  (Data Source: entity.csv)
   - Implement requirements using Unix Shell Scripts.  (Data Source: entity.csv)
   - Implement requirements using MongoDB.  (Data Source: entity.json)
   - Implement requirements using Python.  (Data Source: entity.xlsx)
   - Implement requirements using PySpark Core Program.  (Data Source: entity.csv)
   - Implement requirements using PySpark Core, DataFrame and SQL.
   - Implement requirements using Database Programming in MySQL.
   - Implement requirements using PowerBI.

FORMAT RULES (strictly follow):
- Each implementation section starts with "Implement requirements using X." on its own line
- Each task is: a number followed by period, then bold title, then italic scenario description on the next lines
- Example format for a task:
  1. Task Title Here
  The scenario team explains the business need here in italics. They want specific output.
- Use ONLY these job roles: Data Engineer, Junior Data Engineer, Python Developer, QA Analyst, Reporting Analyst, Data Analyst, ETL Developer, System Administrator, Operations Analyst, BI Developer, Database Developer, Support Engineer
- Do NOT use: Chief Architect, Director, VP, Senior Solution Architect, Enterprise Architect, Distinguished Engineer, Principal Engineer
- Keep difficulty beginner-to-intermediate (academic capstone level)
- Use realistic, practical scenario descriptions (not abstract)
- Generate 6-8 tasks per section minimum
- Solutions must include working code snippets for every task

COMPLEXITY LEVEL: ${cmeta.label.toUpperCase()}
${cmeta.modifier}`;

      const resp=await fetch('https://api.anthropic.com/v1/messages',{
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body:JSON.stringify({
          model:'claude-sonnet-4-20250514',
          max_tokens:4000,
          system:systemPrompt,
          messages:[{role:'user',content:currentPrompt+
            '\n\nRespond ONLY with a JSON object with keys "requirements" and "solutions".'}],
        }),
      });

      const data=await resp.json();
      if(data.error) throw new Error(data.error.message);

      const raw=(data.content||[]).filter(b=>b.type==='text').map(b=>b.text).join('');
      const clean=raw.trim().replace(/^```json\s*/,'').replace(/^```\s*/,'').replace(/\s*```$/,'').trim();
      const parsed=JSON.parse(clean);
      reqText=parsed.requirements||raw;
      solText=parsed.solutions||'';

    }else{
      await new Promise(r=>setTimeout(r,7500));
      const company=rand(COMPANIES);
      const names=shuffle(NON_INDIAN_NAMES).slice(0,2);
      reqText=buildDemoReq(selectedSector,entities,company,names);
      solText=buildDemoSol(selectedSector,entities);
    }

    stopProgress();
    await new Promise(r=>setTimeout(r,350));

    currentReq=reqText; currentSol=solText;
    // Build schema OOXML once and store it
    currentSchema=buildSchemaOoxml(selectedSector,getEntities(selectedSector));
    document.getElementById('gen-progress').style.display='none';
    buildDownloadSection(selectedSector,entities);
    document.getElementById('step-download').style.display='block';
    document.getElementById('step-download').classList.add('fadein');
    document.getElementById('snum3').className='step-num done';
    document.getElementById('gen-status').textContent='Done ✓';
    showToast('success',`✓ Output ready · ${cmeta.label} · ${selectedSector}`);

  }catch(err){
    stopProgress();
    document.getElementById('gen-progress').style.display='none';
    showToast('error','Error: '+err.message.slice(0,80));
    document.getElementById('snum3').classList.remove('active');
  }

  generating=false;
  document.getElementById('gen-btn').disabled=false;
  document.getElementById('gen-btn-inner').textContent='⚡ Generate Output Files';
}

/* ── BUILD DOWNLOAD SECTION ─────────────────────────────── */
function buildDownloadSection(sector,entities){
  const safe=sector.replace(/ & /g,'_').replace(/ /g,'_');
  const ts=new Date().toISOString().slice(0,10);
  const fCombined=`${safe}_capstone_${ts}.docx`;

  document.getElementById('files-grid').innerHTML=`
    <div class="file-card fadein" style="max-width:600px;margin:0 auto;">
      <div class="file-card-header">
        <div class="file-icon">📦</div>
        <div><div class="file-name">${fCombined}</div>
          <div style="font-size:11px;color:var(--text3);margin-top:2px;">Requirements · Solutions · Entity Schema (tables)</div>
        </div>
      </div>
      <div class="file-size">${(currentReq.length+currentSol.length).toLocaleString()} characters · ${entities.length} entities · all sections combined</div>
      <div class="file-preview">${escH(currentReq.slice(0,320))}</div>
      <button class="btn-action btn-download file-dl-btn" style="margin-top:14px;" onclick="downloadCombined()">⬇ Download Combined .docx</button>
    </div>`;
}

/* ══════════════════════════════════════════════════════════
   SCHEMA AS OOXML TABLES  ← the only section that changed
══════════════════════════════════════════════════════════ */
function buildSchemaOoxml(sector, entities){
  const e0=entities[0], e1=entities[1], e2=entities[2], e3=entities[3];

  const schemas=[
    {
      name:e0, format:'CSV',
      fields:[
        {field:'id',           datatype:'INT',       description:'Unique numeric identifier for each record'},
        {field:'name',         datatype:'VARCHAR',   description:'Full name of the entity (non-Indian, Western/European)'},
        {field:'status',       datatype:'VARCHAR',   description:'Current status: ACTIVE, INACTIVE, PENDING, ERROR, ARCHIVED'},
        {field:'created_date', datatype:'DATE',      description:'Date the record was created (YYYY-MM-DD)'},
        {field:'region',       datatype:'VARCHAR',   description:'Geographic region associated with the record'},
        {field:'score',        datatype:'FLOAT',     description:'Numeric performance or quality score (0–100)'},
        {field:'is_active',    datatype:'BOOLEAN',   description:'Flag indicating whether the record is currently active'},
      ]
    },
    {
      name:e1, format:'JSON',
      fields:[
        {field:'record_id',    datatype:'INT',        description:'Unique identifier for the JSON record'},
        {field:'entity_ref',   datatype:'INT (FK)',   description:`Foreign key referencing ${e0}(id)`},
        {field:'category',     datatype:'VARCHAR',    description:'Business category label for grouping records'},
        {field:'value',        datatype:'DECIMAL',    description:'Monetary or numeric value associated with the record'},
        {field:'last_updated', datatype:'TIMESTAMP',  description:'ISO 8601 timestamp of the most recent update'},
        {field:'notes',        datatype:'TEXT',       description:'Optional free-text notes or remarks'},
      ]
    },
    {
      name:e2, format:'Excel (.xlsx)',
      fields:[
        {field:'seq_id',         datatype:'INT',     description:'Sequential auto-increment identifier'},
        {field:'description',    datatype:'TEXT',    description:'Human-readable task or item description'},
        {field:'assigned_to',    datatype:'VARCHAR', description:'Name of the person assigned to this item'},
        {field:'priority',       datatype:'ENUM',    description:'Priority level: LOW, MED, or HIGH'},
        {field:'completion_pct', datatype:'FLOAT',   description:'Percentage of task completion (0.00–100.00)'},
      ]
    },
    {
      name:e3, format:'XML / Parquet',
      fields:[
        {field:'log_id',     datatype:'INT',       description:'Unique identifier for each log entry'},
        {field:'source_ref', datatype:'INT (FK)',  description:`Foreign key referencing ${e0}(id)`},
        {field:'event_type', datatype:'VARCHAR',   description:'Event type: STATUS_CHANGE, UPDATE, DELETE, CREATE, AUDIT, ALERT'},
        {field:'details',    datatype:'JSON',      description:'Structured JSON payload with event-specific metadata'},
        {field:'timestamp',  datatype:'DATETIME',  description:'Date and time the event was logged'},
        {field:'severity',   datatype:'VARCHAR',   description:'Severity: LOW, MED, HIGH, or CRITICAL'},
      ]
    },
  ];

  /* ── shared border definition ── */
  const bdr = `
    <w:top    w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
    <w:left   w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
    <w:right  w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>`;

  /* ── cell builder ── */
  function tc(text, width, opts={}){
    const fill  = opts.fill  ? `<w:shading w:val="clear" w:color="auto" w:fill="${opts.fill}"/>` : '';
    const bold  = opts.bold  ? '<w:b/>' : '';
    const color = opts.color ? `<w:color w:val="${opts.color}"/>` : '<w:color w:val="333333"/>';
    const font  = opts.mono
      ? '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>'
      : '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>';
    return `<w:tc>
      <w:tcPr>
        <w:tcW w:w="${width}" w:type="dxa"/>
        <w:tcBorders>${bdr}</w:tcBorders>
        <w:tcMar>
          <w:top    w:w="100" w:type="dxa"/>
          <w:left   w:w="120" w:type="dxa"/>
          <w:bottom w:w="100" w:type="dxa"/>
          <w:right  w:w="120" w:type="dxa"/>
        </w:tcMar>
        ${fill}
      </w:tcPr>
      <w:p>
        <w:r>
          <w:rPr>${bold}${color}${font}<w:sz w:val="18"/></w:rPr>
          <w:t xml:space="preserve">${escXml(text)}</w:t>
        </w:r>
      </w:p>
    </w:tc>`;
  }

  /* ── column widths: Field=2100, DataType=1700, Description=5460 → total=9260 ── */
  const W_FIELD=2100, W_TYPE=1700, W_DESC=5460;

  /* ── section heading ── */
  const sectionHeading =
    `<w:p>
       <w:pPr><w:spacing w:before="400" w:after="160"/></w:pPr>
       <w:r>
         <w:rPr>
           <w:b/>
           <w:sz w:val="30"/>
           <w:color w:val="1F4E79"/>
           <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
         </w:rPr>
         <w:t>Entity Schema Reference — ${escXml(sector)}</w:t>
       </w:r>
     </w:p>
     <w:p>
       <w:pPr>
         <w:pBdr>
           <w:bottom w:val="single" w:sz="6" w:space="1" w:color="2E75B6"/>
         </w:pBdr>
         <w:spacing w:after="200"/>
       </w:pPr>
     </w:p>`;

  /* ── one table per entity ── */
  const tablesXml = schemas.map(entity => {
    const entityHeading =
      `<w:p>
         <w:pPr><w:spacing w:before="280" w:after="80"/></w:pPr>
         <w:r>
           <w:rPr>
             <w:b/>
             <w:sz w:val="22"/>
             <w:color w:val="2E75B6"/>
             <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
           </w:rPr>
           <w:t>${escXml(entity.name)}</w:t>
         </w:r>
         <w:r>
           <w:rPr>
             <w:sz w:val="20"/>
             <w:color w:val="888888"/>
             <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
           </w:rPr>
           <w:t xml:space="preserve">  (${escXml(entity.format)})</w:t>
         </w:r>
       </w:p>`;

    /* header row — dark blue background, white bold text */
    const headerRow =
      `<w:tr>
        <w:trPr><w:trHeight w:val="380"/></w:trPr>
        ${tc('Field',       W_FIELD, {bold:true, color:'FFFFFF', fill:'2E75B6'})}
        ${tc('Data Type',   W_TYPE,  {bold:true, color:'FFFFFF', fill:'2E75B6'})}
        ${tc('Description', W_DESC,  {bold:true, color:'FFFFFF', fill:'2E75B6'})}
      </w:tr>`;

    /* data rows — alternating row tint */
    const dataRows = entity.fields.map((f, i) => {
      const bg = i % 2 === 0 ? 'EEF4FB' : 'FFFFFF';
      return `<w:tr>
        <w:trPr><w:trHeight w:val="340"/></w:trPr>
        ${tc(f.field,       W_FIELD, {mono:true, color:'1A1A1A', fill:bg})}
        ${tc(f.datatype,    W_TYPE,  {color:'1F6391',            fill:bg})}
        ${tc(f.description, W_DESC,  {color:'333333',            fill:bg})}
      </w:tr>`;
    }).join('');

    const table =
      `<w:tbl>
        <w:tblPr>
          <w:tblW w:w="${W_FIELD+W_TYPE+W_DESC}" w:type="dxa"/>
          <w:tblBorders>
            <w:top    w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
            <w:left   w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
            <w:right  w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
          </w:tblBorders>
          <w:tblCellMar>
            <w:top    w:w="0" w:type="dxa"/>
            <w:left   w:w="0" w:type="dxa"/>
            <w:bottom w:w="0" w:type="dxa"/>
            <w:right  w:w="0" w:type="dxa"/>
          </w:tblCellMar>
        </w:tblPr>
        <w:tblGrid>
          <w:gridCol w:w="${W_FIELD}"/>
          <w:gridCol w:w="${W_TYPE}"/>
          <w:gridCol w:w="${W_DESC}"/>
        </w:tblGrid>
        ${headerRow}
        ${dataRows}
      </w:tbl>
      <w:p><w:pPr><w:spacing w:after="240"/></w:pPr></w:p>`;

    return entityHeading + table;
  }).join('');

  return sectionHeading + tablesXml;
}

/* Keep old name as alias so nothing else breaks */
function buildSchemaText(sector, entities){
  return buildSchemaOoxml(sector, entities);
}

/* ══════════════════════════════════════════════════════════
   PDF-STYLE DOCUMENT GENERATOR
   Produces a professionally formatted .docx that mirrors
   the reference PDF: bold labels, italic scenario text,
   numbered tasks, coloured section dividers, schema tables.
══════════════════════════════════════════════════════════ */

/* ── SHARED STYLES XML ───────────────────────────────────── */
const RICH_STYLES_XML =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
      <w:sz w:val="20"/><w:color w:val="333333"/>
    </w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="numbering" w:styleId="ListNumber"><w:name w:val="List Number"/></w:style>
  <w:style w:type="numbering" w:styleId="ListBullet"><w:name w:val="List Bullet"/></w:style>
</w:styles>`;

/* ── NUMBERING XML (bullets + decimal lists) ─────────────── */
const NUMBERING_XML =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="2">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2022;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="3">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2013;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1080" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
  <w:num w:numId="3"><w:abstractNumId w:val="3"/></w:num>
</w:numbering>`;

/* ── makeRichDocx2: includes numbering.xml ───────────────── */
function makeRichDocx2(bodyXml){
  const enc = new TextEncoder();

  const docXml =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>${bodyXml}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body>
</w:document>`;

  const rootRels =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const wordRels =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>`;

  const typesXml =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`;

  return makeZipBlob([
    ['[Content_Types].xml',          enc.encode(typesXml)],
    ['_rels/.rels',                  enc.encode(rootRels)],
    ['word/document.xml',            enc.encode(docXml)],
    ['word/styles.xml',              enc.encode(RICH_STYLES_XML)],
    ['word/numbering.xml',           enc.encode(NUMBERING_XML)],
    ['word/_rels/document.xml.rels', enc.encode(wordRels)],
  ]);
}

/* ══════════════════════════════════════════════════════════
   OOXML PRIMITIVE BUILDERS
══════════════════════════════════════════════════════════ */

/* bold label like "Project Title:" */
function oxLabel(text){ return (
  `<w:p><w:pPr><w:spacing w:before="200" w:after="60"/></w:pPr>
   <w:r><w:rPr><w:b/><w:sz w:val="22"/><w:color w:val="1F3864"/>
     <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
   <w:t xml:space="preserve">${escXml(text)}</w:t></w:r></w:p>`);
}

/* bold+italic section sub-label e.g. "Data Source: patients2.txt" */
function oxDataSource(label, value){ return (
  `<w:p><w:pPr><w:spacing w:before="120" w:after="80"/></w:pPr>
   <w:r><w:rPr><w:b/><w:sz w:val="21"/>
     <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
   <w:t xml:space="preserve">${escXml(label)}</w:t></w:r>
   <w:r><w:rPr><w:sz w:val="21"/>
     <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
   <w:t xml:space="preserve">${escXml(value)}</w:t></w:r></w:p>`);
}

/* section heading with blue bottom border */
function oxSectionHeading(text){ return (
  `<w:p><w:pPr>
     <w:spacing w:before="300" w:after="120"/>
     <w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="2E74B5"/></w:pBdr>
   </w:pPr>
   <w:r><w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="1F3864"/>
     <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
   <w:t xml:space="preserve">${escXml(text)}</w:t></w:r></w:p>`);
}

/* plain body paragraph */
function oxBody(text){ return (
  `<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>
   <w:r><w:rPr><w:sz w:val="20"/>
     <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
   <w:t xml:space="preserve">${escXml(text)}</w:t></w:r></w:p>`);
}

/* italic indented scenario description */
function oxItalic(text){ return (
  `<w:p><w:pPr><w:spacing w:before="30" w:after="100"/>
     <w:ind w:left="720"/></w:pPr>
   <w:r><w:rPr><w:i/><w:sz w:val="20"/>
     <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
   <w:t xml:space="preserve">${escXml(text)}</w:t></w:r></w:p>`);
}

/* bullet item (numId=2) */
function oxBullet(runs){ return (
  `<w:p><w:pPr>
     <w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>
     <w:spacing w:before="60" w:after="60"/>
   </w:pPr>${runs}</w:p>`);
}

/* sub-bullet (numId=3, dash) */
function oxSubBullet(runs){ return (
  `<w:p><w:pPr>
     <w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr>
     <w:spacing w:before="30" w:after="30"/>
   </w:pPr>${runs}</w:p>`);
}

/* bold run */
function oxB(text){ return (
  `<w:r><w:rPr><w:b/><w:sz w:val="20"/>
     <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
   <w:t xml:space="preserve">${escXml(text)}</w:t></w:r>`);
}

/* italic run */
function oxI(text){ return (
  `<w:r><w:rPr><w:i/><w:sz w:val="20"/>
     <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
   <w:t xml:space="preserve">${escXml(text)}</w:t></w:r>`);
}

/* normal run */
function oxT(text){ return (
  `<w:r><w:rPr><w:sz w:val="20"/>
     <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
   <w:t xml:space="preserve">${escXml(text)}</w:t></w:r>`);
}

/* bold+italic run */
function oxBI(text){ return (
  `<w:r><w:rPr><w:b/><w:i/><w:sz w:val="20"/>
     <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
   <w:t xml:space="preserve">${escXml(text)}</w:t></w:r>`);
}

/* numbered item title (bold) — uses numId=1 */
function oxNumTitle(text, numRef){ return (
  `<w:p><w:pPr>
     <w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numRef||1}"/></w:numPr>
     <w:spacing w:before="120" w:after="30"/>
   </w:pPr>
   <w:r><w:rPr><w:b/><w:sz w:val="21"/>
     <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
   <w:t xml:space="preserve">${escXml(text)}</w:t></w:r></w:p>`);
}

/* empty spacer */
function oxSp(before,after){ return (
  `<w:p><w:pPr><w:spacing w:before="${before||60}" w:after="${after||40}"/></w:pPr></w:p>`);
}

/* page break */
function oxPageBreak(){ return (
  `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`);
}

/* ── 2-col Dataset Source table ─────────────────────────── */
function oxDatasetTable(rows){
  const bdr = `<w:top w:val="single" w:sz="4" w:color="2E74B5"/>
               <w:left w:val="single" w:sz="4" w:color="2E74B5"/>
               <w:bottom w:val="single" w:sz="4" w:color="2E74B5"/>
               <w:right w:val="single" w:sz="4" w:color="2E74B5"/>`;
  function hdr(t,w){ return `<w:tc>
    <w:tcPr><w:tcW w:w="${w}" w:type="dxa"/><w:tcBorders>${bdr}</w:tcBorders>
      <w:shading w:val="clear" w:fill="2E74B5"/>
      <w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>
               <w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>
    </w:tcPr>
    <w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
      <w:t>${escXml(t)}</w:t></w:r></w:p></w:tc>`; }
  function cell(t,w,shade){ return `<w:tc>
    <w:tcPr><w:tcW w:w="${w}" w:type="dxa"/><w:tcBorders>${bdr}</w:tcBorders>
      <w:shading w:val="clear" w:fill="${shade?'D6E4F0':'FFFFFF'}"/>
      <w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>
               <w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>
    </w:tcPr>
    <w:p><w:r><w:rPr><w:sz w:val="20"/>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
      <w:t>${escXml(t)}</w:t></w:r></w:p></w:tc>`; }
  const headerRow = `<w:tr><w:trPr><w:trHeight w:val="380"/></w:trPr>
    ${hdr('Source Type',2800)}${hdr('File Name',6560)}</w:tr>`;
  const dataRows = rows.map((r,i)=>
    `<w:tr><w:trPr><w:trHeight w:val="340"/></w:trPr>
      ${cell(r[0],2800,i%2===0)}${cell(r[1],6560,i%2===0)}</w:tr>`
  ).join('');
  return `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="9360" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:color="2E74B5"/>
        <w:left w:val="single" w:sz="4" w:color="2E74B5"/>
        <w:bottom w:val="single" w:sz="4" w:color="2E74B5"/>
        <w:right w:val="single" w:sz="4" w:color="2E74B5"/>
        <w:insideH w:val="single" w:sz="4" w:color="2E74B5"/>
        <w:insideV w:val="single" w:sz="4" w:color="2E74B5"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblGrid><w:gridCol w:w="2800"/><w:gridCol w:w="6560"/></w:tblGrid>
    ${headerRow}${dataRows}
  </w:tbl><w:p><w:pPr><w:spacing w:after="200"/></w:pPr></w:p>`;
}

/* ── 3-col Schema table (SL No / Column Name / Description) */
function oxSchemaTable3(title, rows){
  const bdr = `<w:top w:val="single" w:sz="4" w:color="2E74B5"/>
               <w:left w:val="single" w:sz="4" w:color="2E74B5"/>
               <w:bottom w:val="single" w:sz="4" w:color="2E74B5"/>
               <w:right w:val="single" w:sz="4" w:color="2E74B5"/>`;
  function hdr(t,w){ return `<w:tc>
    <w:tcPr><w:tcW w:w="${w}" w:type="dxa"/><w:tcBorders>${bdr}</w:tcBorders>
      <w:shading w:val="clear" w:fill="2E74B5"/>
      <w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>
               <w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>
    </w:tcPr>
    <w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
      <w:t>${escXml(t)}</w:t></w:r></w:p></w:tc>`; }
  function cell(t,w,shade,bold){ return `<w:tc>
    <w:tcPr><w:tcW w:w="${w}" w:type="dxa"/><w:tcBorders>${bdr}</w:tcBorders>
      <w:shading w:val="clear" w:fill="${shade?'D6E4F0':'FFFFFF'}"/>
      <w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>
               <w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>
    </w:tcPr>
    <w:p><w:r><w:rPr>${bold?'<w:b/>':''}<w:sz w:val="20"/>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
      <w:t>${escXml(t)}</w:t></w:r></w:p></w:tc>`; }
  const headerRow = `<w:tr><w:trPr><w:trHeight w:val="380"/></w:trPr>
    ${hdr('SL No',800)}${hdr('Column Name',2600)}${hdr('Description',5960)}</w:tr>`;
  const dataRows = rows.map((r,i)=>
    `<w:tr><w:trPr><w:trHeight w:val="340"/></w:trPr>
      ${cell(String(r[0]),800,i%2===0)}
      ${cell(r[1],2600,i%2===0,true)}
      ${cell(r[2],5960,i%2===0)}</w:tr>`
  ).join('');
  const labelXml = title
    ? `<w:p><w:pPr><w:spacing w:before="200" w:after="60"/></w:pPr>
       <w:r><w:rPr><w:b/><w:sz w:val="22"/><w:color w:val="1F3864"/>
         <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
       <w:t>${escXml(title)}</w:t></w:r></w:p>` : '';
  return labelXml + `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="9360" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:color="2E74B5"/>
        <w:left w:val="single" w:sz="4" w:color="2E74B5"/>
        <w:bottom w:val="single" w:sz="4" w:color="2E74B5"/>
        <w:right w:val="single" w:sz="4" w:color="2E74B5"/>
        <w:insideH w:val="single" w:sz="4" w:color="2E74B5"/>
        <w:insideV w:val="single" w:sz="4" w:color="2E74B5"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblGrid><w:gridCol w:w="800"/><w:gridCol w:w="2600"/><w:gridCol w:w="5960"/></w:tblGrid>
    ${headerRow}${dataRows}
  </w:tbl><w:p><w:pPr><w:spacing w:after="200"/></w:pPr></w:p>`;
}

/* ══════════════════════════════════════════════════════════
   PARSE PLAIN TEXT → STRUCTURED OOXML
   Converts structured text content into PDF-style layout
══════════════════════════════════════════════════════════ */
function plainToStructuredOoxml(text, sector, entities){
  let xml = '';
  const lines = text.split('\n');
  let i = 0;

  function isDivider(l){ return /^[═━─]{4,}/.test(l.trim()); }
  function isStoryId(l){ return /^US-[A-Z]+-\d+/.test(l.trim()); }
  function isNumberedLine(l){ return /^\d+\.\s/.test(l.trim()); }
  function isPipeTable(l){ return l.includes('|') && l.trim().startsWith(/\d|SL/i.exec(l.trim())||l.trim()[0]); }

  while(i < lines.length){
    const raw = lines[i];
    const t   = raw.trim();

    /* blank lines — minimal space */
    if(!t){ i++; continue; }

    /* ── Page break markers ── */
    if(t === '---PAGE---'){ xml += oxPageBreak(); i++; continue; }

    /* ── Section dividers (═══ ━━━) → skip, next line is heading ── */
    if(isDivider(t)){ i++; continue; }

    /* ── "Project Title:" label ── */
    if(/^Project Title:/i.test(t)){
      xml += oxLabel('Project Title:');
      i++;
      /* next non-empty line is the title itself */
      while(i < lines.length && !lines[i].trim()) i++;
      if(i < lines.length){
        xml += `<w:p><w:pPr><w:spacing w:before="60" w:after="200"/></w:pPr>
          <w:r><w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="1F3864"/>
            <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
          <w:t xml:space="preserve">${escXml(lines[i].trim())}</w:t></w:r></w:p>`;
        i++;
      }
      continue;
    }

    /* ── "Business Scenario:" heading ── */
    if(/^Business Scenario:/i.test(t)){
      xml += oxLabel('Business Scenario:');
      i++;
      while(i < lines.length){
        const next = lines[i].trim();
        if(!next){ i++; continue; }
        /* stop at next major section */
        if(/^Dataset Source|^Schema Details|^Implementation:|^Implement requirements/i.test(next)) break;
        if(/^•/.test(next)){
          xml += oxBullet(oxT(next.replace(/^•\s*/,'')));
        } else {
          xml += oxBody(next);
        }
        i++;
      }
      continue;
    }

    /* ── "Dataset Source Overview:" ── */
    if(/^Dataset Source Overview:/i.test(t)){
      xml += oxPageBreak();
      xml += oxSectionHeading('Dataset Source Overview:');
      i++;
      /* skip blank lines before table */
      while(i < lines.length && !lines[i].trim()) i++;
      /* collect pipe-table rows (skip header row "Source Type | File Name") */
      const tableRows = [];
      while(i < lines.length && (lines[i].includes('|') || !lines[i].trim())){
        const l = lines[i].trim();
        if(!l){ i++; continue; }
        /* stop if we hit a non-table line */
        if(!l.includes('|')) break;
        const parts = l.split('|').map(p=>p.trim()).filter(Boolean);
        if(parts.length >= 2 && !/^Source Type/i.test(parts[0]) && !/^-+/.test(parts[0])){
          tableRows.push([parts[0], parts[1]]);
        }
        i++;
      }
      if(tableRows.length) xml += oxDatasetTable(tableRows);
      continue;
    }

    /* ── "Schema Details for X:" ── */
    if(/^Schema Details for /i.test(t)){
      xml += oxLabel(t);
      i++;
      /* skip blank lines before table */
      while(i < lines.length && !lines[i].trim()) i++;
      /* skip header row "SL No | Column Name | Description" */
      if(i < lines.length && /^SL No/i.test(lines[i].trim())) i++;
      /* skip separator line "---" if present */
      while(i < lines.length && /^[-|]+$/.test(lines[i].trim())) i++;
      /* collect data rows */
      const schemaRows = [];
      while(i < lines.length){
        const sl = lines[i].trim();
        if(!sl){ i++; continue; }           /* skip blank lines inside table */
        if(!sl.includes('|')) break;         /* stop at non-table line */
        if(/^Schema Details|^Dataset|^Implementation|^Implement/i.test(sl)) break;
        const parts = sl.split('|').map(p=>p.trim());
        if(parts.length >= 3 && /^\d+$/.test(parts[0])){
          schemaRows.push([parseInt(parts[0]), parts[1], parts[2]]);
        }
        i++;
      }
      if(schemaRows.length) xml += oxSchemaTable3('', schemaRows);
      continue;
    }

    /* ── "Implementation:" standalone label ── */
    if(/^Implementation:$/i.test(t)){
      xml += oxPageBreak();
      xml += oxSectionHeading('Implementation:');
      i++; continue;
    }

    /* ── "Implement requirements using X." heading ── */
    if(/^Implement requirements using/i.test(t)){
      xml += `<w:p><w:pPr><w:spacing w:before="200" w:after="80"/></w:pPr>
        <w:r><w:rPr><w:b/><w:sz w:val="22"/>
          <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
        <w:t xml:space="preserve">${escXml(t)}</w:t></w:r></w:p>`;
      i++; continue;
    }

    /* ── "Data Source: xxx" ── */
    if(/^Data Source:/i.test(t)){
      const col = t.indexOf(':');
      xml += oxDataSource(t.slice(0,col+1)+' ', t.slice(col+1).trim());
      i++; continue;
    }

    /* ── "Note:" lines ── */
    if(/^Note:/i.test(t)){
      xml += `<w:p><w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr>
        ${oxB('Note: ')}${oxI(t.slice(5).trim())}</w:p>`;
      i++; continue;
    }

    /* ── US-xxx story ID lines ── */
    if(isStoryId(t)){
      const colon = t.indexOf(':');
      const id   = colon>0 ? t.slice(0,colon) : t;
      const rest = colon>0 ? t.slice(colon+1).trim() : '';
      xml += `<w:p><w:pPr><w:spacing w:before="120" w:after="30"/></w:pPr>
        ${oxB(id+(rest?': ':''))}${rest?oxT(rest):''}</w:p>`;
      i++;
      while(i < lines.length && lines[i].trim() &&
            !isStoryId(lines[i].trim()) && !isDivider(lines[i].trim()) &&
            !isNumberedLine(lines[i].trim()) &&
            !/^Implement requirements|^Data Source:|^US-/i.test(lines[i].trim())){
        const cont = lines[i].trim();
        if(/^(INPUT|OUTPUT|SCOPE):/i.test(cont)){
          xml += `<w:p><w:pPr><w:spacing w:before="20" w:after="20"/><w:ind w:left="720"/></w:pPr>
            ${oxB(cont.split(':')[0]+': ')}${oxI(cont.slice(cont.indexOf(':')+1).trim())}</w:p>`;
        } else if(/^[•\-\*]/.test(cont)){
          xml += oxSubBullet(oxI(cont.replace(/^[•\-\*]\s*/,'')));
        } else {
          xml += oxItalic(cont);
        }
        i++;
      }
      continue;
    }

    /* ── Numbered task items: "1. Title" ── */
    if(isNumberedLine(t)){
      const dot   = t.indexOf('.');
      const title = t.slice(dot+1).trim();
      xml += `<w:p><w:pPr>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
        <w:spacing w:before="140" w:after="30"/>
      </w:pPr><w:r><w:rPr><w:b/><w:sz w:val="21"/>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
        <w:t xml:space="preserve">${escXml(title)}</w:t></w:r></w:p>`;
      i++;
      /* skip single blank line between title and scenario */
      while(i < lines.length && !lines[i].trim()) i++;
      /* collect scenario lines until next numbered item / heading */
      while(i < lines.length){
        const next = lines[i].trim();
        /* stop conditions */
        if(isNumberedLine(next) || isDivider(next) || isStoryId(next)) break;
        if(/^Implement requirements|^Data Source:|^Schema Details|^Implementation:|^Dataset Source/i.test(next)) break;
        /* blank line = end of this task's scenario block */
        if(!next){
          /* peek ahead: if next non-blank line is another numbered item or heading, stop */
          let j = i+1;
          while(j < lines.length && !lines[j].trim()) j++;
          if(j >= lines.length) break;
          const peek = lines[j].trim();
          if(isNumberedLine(peek) || isDivider(peek) ||
             /^Implement requirements|^Data Source:|^Schema Details|^Implementation:/i.test(peek)) break;
          /* otherwise it's a blank separator within the scenario — skip it */
          i++;
          continue;
        }
        if(/^[•\-\*]/.test(next)){
          xml += oxSubBullet(oxI(next.replace(/^[•\-\*]\s*/,'')));
        } else if(/^(INPUT|OUTPUT|SCOPE|NOTE):/i.test(next)){
          xml += `<w:p><w:pPr><w:spacing w:before="20" w:after="20"/><w:ind w:left="720"/></w:pPr>
            ${oxB(next.split(':')[0]+': ')}${oxI(next.slice(next.indexOf(':')+1).trim())}</w:p>`;
        } else {
          xml += oxItalic(next);
        }
        i++;
      }
      continue;
    }

    /* ── Bullet lines (•, *, -) ── */
    if(/^[•\-\*]/.test(t)){
      xml += oxBullet(oxT(t.replace(/^[•\-\*]\s*/,'')));
      i++; continue;
    }

    /* ── Code-like lines (indented or starts with $, #, keywords) ── */
    if(/^\s{2,}/.test(raw) ||
       /^(\$|#|\/\/|use |db\.|spark\.|from |import |def |class |if |for |SELECT|CREATE|INSERT|UPDATE|DELETE|CALL|DELIMITER|awk|sed|grep|bash)/.test(t)){
      xml += `<w:p><w:pPr><w:spacing w:before="20" w:after="20"/><w:ind w:left="720"/></w:pPr>
        <w:r><w:rPr><w:sz w:val="18"/><w:color w:val="1a1a1a"/>
          <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr>
          <w:t xml:space="preserve">${escXml(t)}</w:t></w:r></w:p>`;
      i++; continue;
    }

    /* ── ALL-CAPS short lines (section sub-labels) ── */
    if(/^[A-Z][A-Z &,\/]{3,}:?\s*$/.test(t) && t.length < 60){
      xml += oxLabel(t);
      i++; continue;
    }

    /* ── "Solutions & Code — Sector" heading in solution text ── */
    if(/^Solutions & Code|^Solutions —/i.test(t)){
      xml += oxSectionHeading(t);
      i++; continue;
    }

    /* ── "Generated:" line — skip ── */
    if(/^Generated:/i.test(t)){ i++; continue; }

    /* ── X Solutions (sub-section labels like "Unix Commands — Solutions") ── */
    if(/\bSolutions\b/.test(t) && t.length < 80){
      xml += `<w:p><w:pPr><w:spacing w:before="180" w:after="80"/>
        <w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="BBBBBB"/></w:pBdr>
      </w:pPr><w:r><w:rPr><w:b/><w:sz w:val="22"/><w:color w:val="2E74B5"/>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
        <w:t xml:space="preserve">${escXml(t)}</w:t></w:r></w:p>`;
      i++; continue;
    }

    /* ── Pipe-delimited table row (schema etc) — already handled above, skip ── */
    if(t.includes('|') && /^\d+\s*\|/.test(t)){
      i++; continue;
    }

    /* ── Default: plain body ── */
    xml += oxBody(t);
    i++;
  }
  return xml;
}

/* ══════════════════════════════════════════════════════════
   BUILD FULL PDF-STYLE DOCUMENT BODY
══════════════════════════════════════════════════════════ */
function buildPdfStyleDocx(sector, entities, reqText, solText){
  const e0=entities[0]||'Entity_A', e1=entities[1]||'Entity_B',
        e2=entities[2]||'Entity_C', e3=entities[3]||'Entity_D';
  let xml = '';

  /* Parse the full reqText which now contains everything in structured form */
  xml += plainToStructuredOoxml(reqText, sector, entities);

  /* Solutions section */
  if(solText){
    xml += oxPageBreak();
    xml += oxSectionHeading('Solutions & Code');
    xml += plainToStructuredOoxml(solText, sector, entities);
  }

  return xml;
}

/* ── DOWNLOAD COMBINED .docx (PDF-style) ────────────────── */
function downloadCombined(){
  if(!currentReq && !currentSol){ showToast('error','No content to download'); return; }
  const entities = getEntities(selectedSector||'');
  const bodyXml = buildPdfStyleDocx(selectedSector||'Capstone', entities, currentReq, currentSol);
  const blob = makeRichDocx2(bodyXml);
  const safe=(selectedSector||'output').replace(/ & /g,'_').replace(/ /g,'_');
  const ts=new Date().toISOString().slice(0,10);
  triggerDownload(blob, `${safe}_capstone_${ts}.docx`);
}

/* Keep for backward compatibility */
function downloadFile(type){ downloadCombined(); }
function downloadBoth(){ downloadCombined(); }

/* ── SCHEMA WORD DOC DOWNLOAD (kept for backward compat) ── */
function downloadSchemaDoc(){ downloadCombined(); }

/* ══════════════════════════════════════════════════════════
   DEMO CONTENT BUILDERS  — PDF-matching format
   Structure mirrors the reference document exactly:
   numbered tasks + italic scenario descriptions per section
══════════════════════════════════════════════════════════ */

/* ── Allowed job titles (beginner–intermediate only) ─── */
const JOB_ROLES = [
  'Data Engineer','Junior Data Engineer','Python Developer',
  'QA Analyst','Reporting Analyst','Data Analyst','ETL Developer',
  'System Administrator','Operations Analyst','BI Developer',
  'Database Developer','Support Engineer',
];
function role(){ return rand(JOB_ROLES); }

function buildDemoReq(sector, entities, company, names){
  const e0=entities[0], e1=entities[1], e2=entities[2], e3=entities[3];
  const E0=e0.toLowerCase(), E1=e1.toLowerCase(), E2=e2.toLowerCase(), E3=e3.toLowerCase();

  return (
`Project Title:

"End-to-End ${sector} Data Engineering Pipeline for Scalable Analytics and Reporting"

Business Scenario:

A mid-sized ${sector.toLowerCase()} organisation is facing challenges in managing ${E0} records, ${E1} data, and ${E2} operations due to outdated legacy systems. These systems are unable to handle large volumes of data, lack flexibility, and do not support advanced analytics or real-time reporting.

The organisation wants to modernise its data infrastructure using:

• Unix for file management and automation
• MongoDB for flexible document-based storage
• Python for data preprocessing and file handling
• PySpark for scalable data cleansing and analytics
• Advanced MySQL for relational operations and stored procedures
• Power BI for interactive dashboards and business insights

The data is fragmented across multiple formats and files, with inconsistencies such as:

• Missing values
• Repeated records
• Irregular date formats
• Trailing/leading spaces
• Inconsistent column names
• And so on...

The goal is to build a robust, end-to-end data pipeline that:

• Cleans and integrates data from multiple sources
• Performs complex transformations and analytics
• Enables real-time reporting and decision-making

Dataset Source Overview:

Source Type | File Name
CSV         | ${E0}.csv
JSON        | ${E1}.json
Excel       | ${E2}.xlsx
XML/Parquet | ${E3}.xml

Schema Details for ${e0}:

SL No | Column Name   | Description
1     | id            | Unique numeric identifier for each record
2     | name          | Full name of the entity (Western/European)
3     | status        | Current status: ACTIVE, INACTIVE, PENDING, ERROR
4     | created_date  | Date the record was created (YYYY-MM-DD)
5     | region        | Geographic region associated with the record
6     | score         | Numeric performance or quality score (0–100)
7     | is_active     | Flag indicating whether the record is currently active

Schema Details for ${e1}:

SL No | Column Name   | Description
1     | record_id     | Unique identifier for the JSON record
2     | entity_ref    | Foreign key referencing ${e0}(id)
3     | category      | Business category label for grouping records
4     | value         | Monetary or numeric value associated with the record
5     | last_updated  | ISO 8601 timestamp of the most recent update
6     | notes         | Optional free-text notes or remarks

Schema Details for ${e2}:

SL No | Column Name    | Description
1     | seq_id         | Sequential auto-increment identifier
2     | description    | Human-readable task or item description
3     | assigned_to    | Name of the person assigned to this item
4     | priority       | Priority level: LOW, MED, or HIGH
5     | completion_pct | Percentage of task completion (0.00–100.00)

Schema Details for ${e3}:

SL No | Column Name | Description
1     | log_id      | Unique identifier for each log entry
2     | source_ref  | Foreign key referencing ${e0}(id)
3     | event_type  | Event type: STATUS_CHANGE, UPDATE, DELETE, CREATE, AUDIT
4     | details     | Structured JSON payload with event metadata
5     | timestamp   | Date and time the event was logged
6     | severity    | Severity: LOW, MED, HIGH, or CRITICAL

Implementation:

Implement requirements using Unix Commands.

Data Source: ${E0}.csv

1. Status-Based Record Extraction
The operations team wants to isolate problematic records before the ETL pipeline starts. They ask the Unix team to extract all rows from ${E0}.csv where the status is ERROR or PENDING, regardless of case sensitivity.

2. High-Value Record Audit
The finance team wants to audit high-value entries. They request a list of all records from ${E0}.csv where the score exceeds 80, ensuring only valid numeric entries are considered.

3. Top Records by Region
The reporting team wants to understand regional distribution. They need the top five records with the highest score from each region, sorted by score in descending order.

4. Data Quality Check
The IT team suspects incomplete records in ${E0}.csv. They need to count how many records have missing information in any field.

5. Duplicate Detection
The data steward wants to check for repeated entries. They ask for a count of duplicate rows in ${E0}.csv based on the id field.

6. Record Count Summary
The operations team wants a daily record count summary. They need the total number of records in all four entity files printed to the console and saved to size_report.txt.

7. Date Range Filter
The compliance team needs records created in a specific month. They ask for all rows from ${E0}.csv where created_date falls within the last 30 days.

Implement requirements using Unix Shell Scripts.

Data Source: ${E0}.csv

1. Daily Audit Report
The data governance team needs a daily quality audit. They ask the IT team to count total rows, null rows, and duplicate rows across all four entity files and save the result in a file named audit_report.txt for tracking.

2. Anomaly Extraction Script
The QA team needs to flag records with scores outside the valid range of 0 to 100 in ${E0}.csv. They want these written to anomalies.csv with a reason column appended for investigation.

3. Average Score Calculation
The finance department wants to analyse performance metrics. They ask the IT team to calculate the average score of all records in ${E0}.csv, rounded to two decimal places, ignoring any missing or invalid entries. The result should be printed to the console.

4. Data Cleaning and Standardisation
The data governance team is preparing for a compliance audit. They need a cleaned version of ${E0}.csv where names have no leading or trailing spaces, missing categories are replaced with Unknown, missing values are set to 0, and all dates are converted from DD-MM-YYYY to YYYY-MM-DD format. This cleaned file should be saved as ${E0}_cleaned.csv inside the InputFiles folder.

Implement requirements using MongoDB.

Data Source: ${E1}.json

1. Update Value for High-Priority Records
The finance team has identified that high-priority records require a value adjustment. They request the database team to update the value field to 25000 for all records in ${E1} whose category is Premium.

2. Add Default Category for Uncategorised Records
The operations team has partnered with a default category assignment. They ask the database team to set the category to General for all records in ${E1} that currently have no category listed.

3. Update Missing Timestamps
The administration wants accurate timestamp records for compliance. They instruct the team to update the last_updated field to today's date for all records in ${E1} where this field is missing.

4. Find High-Risk Records
The analysis team is examining high-value entries. They need a list of all records in ${E1} where the value is greater than 15000 for further study.

5. Count Records by Category
The operations team is monitoring category distribution. They ask for the count of records in ${E1} grouped by category, regardless of case sensitivity.

6. Export Updated Records
The data migration team needs all updated records for integration with the reporting system. They request exporting the complete updated dataset into a file named ${E1}_updated.csv and saving it in the InputFiles folder.

Implement requirements using Python.

Data Source: ${E2}.xlsx

1. The analytics rollout is blocked because ${E2}.xlsx is messy and inconsistent. The team lead has asked you to write one Python program that will:

• Read the file safely and trim leading/trailing spaces in every field
• Normalise the priority field to proper capitalisation for consistent reporting
• Convert all date fields to the standard YYYY-MM-DD format
• Fill missing assigned_to values with Unassigned and missing completion_pct with 0
• Skip any row that has fewer than 5 fields to avoid schema issues
• Remove exact duplicate rows caused by repeated uploads
• Save the fully cleansed dataset to InputFiles/${E2}_cleaned.xlsx so the PySpark jobs can run without manual fixes

Deliver this as a single, production-ready Python script that completes the end-to-end cleansing in one pass.

Implement requirements using PySpark Core Program.

Data Source: ${E0}.csv

1. Remove Empty Lines for Data Integrity
The ETL team noticed blank rows in ${E0}.csv that cause schema mismatches during Spark ingestion. They have asked you to ensure all empty lines are removed before processing begins.

2. Clean Up Name and Description Fields
The analytics dashboard is showing inconsistent names and descriptions because of extra spaces. The data team requests that you trim all leading and trailing spaces from the name and description columns to maintain uniformity.

3. Normalise Status for Consistent Grouping
Reports are failing to group records correctly because status values appear in mixed case. The operations team wants you to standardise status by converting it to uppercase.

4. Drop Unnecessary Column
The ingestion pipeline includes an ExtraColumn that is not part of the schema and is causing confusion in downstream joins. The data engineering team asks you to drop this column entirely from the dataset.

5. Remove Duplicate Records
Duplicate rows have been identified due to multiple uploads. The compliance team insists on removing all exact duplicates to prevent inflated counts in reports.

6. Handle Missing Score and Category Values
The finance department needs accurate data for reconciliation. They have asked you to replace missing score values with 0 and missing category values with Unknown so that aggregations do not fail.

Implement requirements using PySpark Core, DataFrame and SQL.

Data Source: ${E0}.csv (cleansed), ${E1}.json, ${E2}.xlsx, ${E3}.xml

Note: Before solving any requirement that uses entity data, first combine all entity files into a unified dataset.

1. Region-wise Record Count for Capacity Planning
The operations team wants to understand data load across regions. They have asked you to combine all entity files into a unified dataset and count how many records belong to each region. This will inform capacity planning.

2. Top Entities by Average Value for Revenue Insights
Finance leadership is reviewing entity performance. They want you to calculate the average value per category and show the top 3 categories with the highest average value.

3. Status Distribution Analysis
The QA team is assessing data quality across all entities. They ask you to count the number of records per status value across the unified dataset and present the results sorted by count descending.

4. Category-wise Aggregates for Business Analysis
Business analysts want to understand financial and demographic patterns. They request a report that groups records by category and computes total value and average score per category across all entity sources.

5. Seasonal Records Filter
The analytics team is examining trends in a specific period. They need you to filter records created between January and March of the current year and count the number of records per region during this window.

6. Top Scoring Records for Incentive Programs
HR and Finance are designing incentive schemes. They ask you to compute total score per assigned_to by aggregating records, and show the top 5 names with the highest total score.

7. Active Senior Records Snapshot
The operations team wants a snapshot of active, high-scoring senior records. Filter records where is_active is true and score is greater than 75, grouped by region and category. Show the top 5 combinations by count.

Implement requirements using Database Programming in MySQL.

Data Source: ${E0}.csv

1. Score Category Function for Finance Review
The finance team needs a way to classify records by score for tiered analysis. They have asked you to create a stored function named GetScoreCategory that accepts a score value and returns one of three categories based on strict rules:

• Low if score is less than 40
• Medium if score is between 40 and 75
• High if score is greater than 75

This function will be used in multiple reports and procedures to standardise score segmentation.

2. Status Label Function for Operations
The operations team needs a reusable function to map status codes to human-readable labels. Create a stored function named GetStatusLabel that accepts a status value and returns:

• Active if status is ACTIVE
• Inactive if status is INACTIVE
• Pending Review if status is PENDING
• Error if status is ERROR

This will help standardise status display across all reports.

3. Procedure to List Records by Score Category
Auditors often request lists of records by tier for compliance checks. Build a stored procedure named ListRecordsByScoreCategory that:

• Accepts a score category as input (Low, Medium, High)
• Internally uses the GetScoreCategory function to filter records
• Displays id, name, score, and ScoreCategory for all matching records

This ensures consistent logic and easy retrieval for audits.

4. Summary by Region and Score Category
Executives want a combined view of regional performance and score distribution. Create a stored procedure named SummaryByRegionAndScore that:

• Accepts no input parameters
• Uses both functions (GetStatusLabel and GetScoreCategory)
• Groups records by region and ScoreCategory
• Displays Region, ScoreCategory, TotalRecords, and AverageScore

This summary will be used for strategic planning and budgeting.

5. Billing by Status
Finance and HR need to identify which status types generate the most activity. Develop a stored procedure named RecordsByStatus that:

• Calculates total count and average score per status
• Displays Status and TotalRecords and AverageScore
• Sorts the result by TotalRecords in descending order

This will support resource allocation and performance benchmarking.

6. Dynamic Region Statistics
The COO wants a dynamic report for every region. Create a stored procedure named RegionStats that:

• Iterates through all distinct regions in the ${E0} table
• For each region, calculates total number of records and average score
• Displays results in the format: Region, RecordCount, AvgScore

This must handle multiple regions dynamically for scalability.

Implement requirements using PowerBI.

Data Source: ${E0}_powerbi.csv

Schema Details:

SL No | Column Name   | Description
1     | id            | Unique identifier for each record
2     | name          | Full name of the entity
3     | status        | Current status of the record
4     | created_date  | Date the record was created
5     | region        | Geographic region
6     | score         | Numeric performance score
7     | is_active     | Active flag (true/false)
8     | category      | Business category label
9     | value         | Monetary value associated with the record
10    | assigned_to   | Name of the assigned person

Transformations:

1. Clean Up Text Fields for Consistency
The reporting team noticed inconsistent spacing in names, regions, and category fields, which breaks slicers and filters. They have asked you to remove leading and trailing spaces from all text fields like name, region, and category.

2. Handle Missing Category Values
Dashboards show blank entries for category, making analysis unreliable. Replace all blank or null values in category with NA to maintain data integrity.

3. Remove Duplicate Records
The data quality team found duplicate rows due to multiple uploads. Remove all duplicate rows from the dataset to ensure accurate counts and value totals.

4. Standardise Status Values
The operations team wants uniform status data. If status is empty or null, replace it with Unknown to reflect the default assignment.

5. Convert Value to Numeric
Some value entries are stored as text like Ten Thousand, causing aggregation errors. Convert these to numeric values and set the value column data type to Decimal Number.

6. Fix Non-Numeric Score Entries
The score column contains text values like Thirty Four or Twenty Seven. Replace these with their numeric equivalents and set the column data type to Whole Number.

7. Flag Invalid Records (DAX)
Compliance checks require identifying records where score is outside the valid range of 0 to 100. Create a new DAX column named IsInvalidScore that flags records as TRUE where score is less than 0 or greater than 100, otherwise FALSE.

8. Calculate Average Score Excluding Zero (DAX)
The analytics team wants a KPI for average score but excluding zero-value cases. Use DAX to calculate AverageScoreExcludingZero where score is greater than 0.

9. Validate Status Entries (DAX)
Data validation checks revealed entries outside the allowed status values. Create a DAX column named IsValidStatus that marks records as TRUE if status is ACTIVE, INACTIVE, PENDING, or ERROR, and FALSE otherwise.

Visualisations:

Operations and Capacity Planning Dashboard:
• Count of records by region where score is greater than 75
• Count of records by created_date (monthly trend)
• Count of records by status (Treemap)
• Count of records by region (Bar Chart)
• Average score by category

Finance and Revenue Insights Dashboard:
• Average value by category
• Sum of value by region
• Total value (Card Visual)
• Sum of value by status (Treemap)
• Sum of value by assigned_to and category

Quality Overview Dashboard:
• Count of records by category where score is greater than 50
• Count of invalid score entries (Card)
• Average score by region
• Count of records by is_active (Donut Chart)
• Count of records by status (Horizontal Bar Chart)
`);
}

/* ── DEMO SOLUTIONS ──────────────────────────────────────── */
function buildDemoSol(sector, entities){
  const e0=entities[0], e1=entities[1], e2=entities[2], e3=entities[3];
  const E0=e0.toLowerCase(), E1=e1.toLowerCase(), E2=e2.toLowerCase(), E3=e3.toLowerCase();
  const db=sector.replace(/ & /g,'_').replace(/ /g,'_').toLowerCase()+'_db';

  return (
`Solutions & Code — ${sector}
Generated: ${new Date().toLocaleString()}

Unix Commands — Solutions

Data Source: ${E0}.csv

1. Status-Based Record Extraction
$ grep -iE ',ERROR,|,PENDING,' ${E0}.csv > flagged_records.csv
$ echo "Flagged records: $(wc -l < flagged_records.csv)"

2. High-Value Record Audit
$ awk -F',' 'NR>1 && $6~/^[0-9.]+$/ && $6+0>80 {print}' ${E0}.csv > high_score.csv

3. Top Records by Region
$ awk -F',' 'NR>1 {print $5","$6","$0}' ${E0}.csv | sort -t',' -k2 -rn | head -5

4. Data Quality Check
$ awk -F',' '{for(i=1;i<=NF;i++) if($i=="") missing++} END{print "Missing fields:", missing+0}' ${E0}.csv

5. Duplicate Detection
$ awk -F',' 'NR>1{print $1}' ${E0}.csv | sort | uniq -d | wc -l

6. Record Count Summary
$ for f in ${E0}.csv ${E1}.json ${E2}.xlsx ${E3}.xml; do echo "$f: $(wc -l < $f)"; done | tee size_report.txt

7. Date Range Filter
$ awk -F',' 'NR==1 || ($4 >= "'$(date -d '30 days ago' +%Y-%m-%d)'")' ${E0}.csv > recent_records.csv

Unix Shell Scripts — Solutions

Data Source: ${E0}.csv

1. Daily Audit Report
#!/bin/bash
REPORT="reports/audit_report_$(date +%Y%m%d).txt"
mkdir -p reports
echo "=== DATA QUALITY AUDIT $(date) ===" > "$REPORT"
printf "%-35s %8s %8s %8s\n" "FILE" "TOTAL" "NULLS" "DUPS" >> "$REPORT"
for FILE in ${E0}.csv ${E1}.json ${E2}.xlsx ${E3}.xml; do
  [ -f "$FILE" ] || continue
  TOTAL=$(tail -n +2 "$FILE" | wc -l)
  NULLS=$(awk -F',' '{for(i=1;i<=NF;i++) if($i==""||$i=="null"||$i=="N/A") n++} END{print n+0}' "$FILE")
  DUPS=$(tail -n +2 "$FILE" | sort | uniq -d | wc -l)
  printf "%-35s %8d %8d %8d\n" "$(basename $FILE)" "$TOTAL" "$NULLS" "$DUPS" >> "$REPORT"
done
echo "Audit saved to $REPORT"

2. Anomaly Extraction Script
#!/bin/bash
INPUT="${E0}.csv"
OUTPUT="anomalies.csv"
head -1 "$INPUT" > "$OUTPUT"
awk -F',' 'NR>1 {
  score=$6
  if (score !~ /^[0-9.]+$/ || score+0 < 0 || score+0 > 100)
    print $0 ",INVALID_SCORE"
}' "$INPUT" >> "$OUTPUT"
echo "Anomalies written to $OUTPUT"

3. Average Score Calculation
#!/bin/bash
AVG=$(awk -F',' 'NR>1 && $6~/^[0-9.]+$/ {sum+=$6; cnt++} END{printf "%.2f", sum/cnt}' ${E0}.csv)
echo "Average Score: $AVG"

4. Data Cleaning and Standardisation
#!/bin/bash
INPUT="${E0}.csv"; OUTPUT="InputFiles/${E0}_cleaned.csv"
mkdir -p InputFiles
awk -F',' 'BEGIN{OFS=","} NR==1{print; next} {
  gsub(/^[ \t]+|[ \t]+$/, "", $2)
  if ($3 == "") $3 = "Unknown"
  if ($6 == "" || $6 !~ /^[0-9.]/) $6 = "0"
  if ($4 ~ /[0-9]{2}-[0-9]{2}-[0-9]{4}/) {
    split($4, d, "-"); $4 = d[3]"-"d[2]"-"d[1]
  }
  print
}' "$INPUT" > "$OUTPUT"
echo "Cleaned file saved to $OUTPUT"

MongoDB Solutions

Data Source: ${E1}.json

1. Update Value for High-Priority Records
use ${db}
db.${E1}.updateMany(
  { category: "Premium" },
  { $set: { value: 25000 } }
)

2. Add Default Category for Uncategorised Records
db.${E1}.updateMany(
  { $or: [{ category: { $exists: false } }, { category: null }, { category: "" }] },
  { $set: { category: "General" } }
)

3. Update Missing Timestamps
db.${E1}.updateMany(
  { $or: [{ last_updated: { $exists: false } }, { last_updated: null }] },
  { $set: { last_updated: new Date().toISOString().slice(0, 10) } }
)

4. Find High-Risk Records
db.${E1}.find(
  { value: { $gt: 15000 } },
  { _id: 0, record_id: 1, category: 1, value: 1 }
).sort({ value: -1 })

5. Count Records by Category
db.${E1}.aggregate([
  { $group: { _id: { $toLower: "$category" }, count: { $sum: 1 } } },
  { $sort: { count: -1 } }
])

6. Export Updated Records
db.${E1}.find({}).forEach(function(doc) {
  // Run from mongo shell with --eval or export via mongoexport:
  // mongoexport --db ${db} --collection ${E1} --type=csv --fields record_id,category,value,last_updated --out InputFiles/${E1}_updated.csv
})

Python Solutions

Data Source: ${E2}.xlsx

1. End-to-End Data Cleaning Script
import pandas as pd

INPUT  = "${E2}.xlsx"
OUTPUT = "InputFiles/${E2}_cleaned.xlsx"

df = pd.read_excel(INPUT, dtype=str)

# Trim all text fields
df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

# Normalise priority capitalisation
if 'priority' in df.columns:
    df['priority'] = df['priority'].str.capitalize()

# Standardise date columns
for col in df.select_dtypes(include='object').columns:
    try:
        df[col] = pd.to_datetime(df[col], dayfirst=True, errors='ignore').dt.strftime('%Y-%m-%d')
    except Exception:
        pass

# Fill missing values
if 'assigned_to'    in df.columns: df['assigned_to']    = df['assigned_to'].fillna('Unassigned')
if 'completion_pct' in df.columns: df['completion_pct'] = df['completion_pct'].fillna('0')

# Drop rows with fewer than 5 non-null fields
df = df.dropna(thresh=5)

# Remove exact duplicates
df = df.drop_duplicates()

df.to_excel(OUTPUT, index=False)
print(f"Cleaned file saved to {OUTPUT} — {len(df)} records")

PySpark Core Solutions

Data Source: ${E0}.csv

from pyspark.sql import SparkSession
from pyspark.sql.functions import col, trim, upper, when

spark = SparkSession.builder.appName("${sector}_Cleaning").getOrCreate()

# 1. Load and remove empty lines
df = spark.read.option("header", True).csv("${E0}.csv").na.drop("all")

# 2. Trim name and description fields
df = df.withColumn("name", trim(col("name")))
if "description" in df.columns:
    df = df.withColumn("description", trim(col("description")))

# 3. Normalise status to uppercase
df = df.withColumn("status", upper(trim(col("status"))))

# 4. Drop ExtraColumn if present
if "ExtraColumn" in df.columns:
    df = df.drop("ExtraColumn")

# 5. Remove duplicate records
df = df.dropDuplicates()

# 6. Fill missing score and category
df = df.fillna({"score": "0", "category": "Unknown"})

df.write.mode("overwrite").option("header", True).csv("InputFiles/${E0}_cleaned")
print("PySpark cleaning complete.")

PySpark Core, DataFrame and SQL Solutions

from pyspark.sql import SparkSession
from pyspark.sql.functions import col, count, avg, sum as spark_sum, desc

spark = SparkSession.builder.appName("${sector}_Analysis").getOrCreate()

# Load all entity files into a unified dataset
df0 = spark.read.option("header", True).csv("InputFiles/${E0}_cleaned")
df1 = spark.read.option("header", True).json("${E1}.json")
df2 = spark.read.option("header", True).option("header", True).csv("${E2}.xlsx")
df3 = spark.read.option("header", True).xml("${E3}.xml")

# 1. Region-wise record count
df0.groupBy("region").count().orderBy(desc("count")).show()

# 2. Top categories by average value
df1.groupBy("category").agg(avg("value").alias("avg_value")) \
   .orderBy(desc("avg_value")).limit(3).show()

# 3. Status distribution
df0.groupBy("status").count().orderBy(desc("count")).show()

# 4. Category-wise aggregates
df1.groupBy("category").agg(
    spark_sum("value").alias("total_value"),
    avg("score").alias("avg_score")
).show()

# 5. Seasonal records filter (Jan–Mar)
from pyspark.sql.functions import month, to_date
df_dates = df0.withColumn("created_date", to_date(col("created_date")))
df_dates.filter(month(col("created_date")).between(1, 3)) \
        .groupBy("region").count().show()

# 6. Top scoring records
df0.groupBy("assigned_to").agg(spark_sum("score").alias("total_score")) \
   .orderBy(desc("total_score")).limit(5).show()

# 7. Active high-scoring records snapshot
df0.filter((col("is_active") == "true") & (col("score").cast("float") > 75)) \
   .groupBy("region", "category").count().orderBy(desc("count")).limit(5).show()

Advanced MySQL Solutions

Data Source: ${E0}.csv

1. GetScoreCategory Function
DELIMITER $$
CREATE FUNCTION GetScoreCategory(p_score FLOAT) RETURNS VARCHAR(10) DETERMINISTIC
BEGIN
  RETURN CASE
    WHEN p_score > 75 THEN 'High'
    WHEN p_score >= 40 THEN 'Medium'
    ELSE 'Low'
  END;
END $$
DELIMITER ;

2. GetStatusLabel Function
DELIMITER $$
CREATE FUNCTION GetStatusLabel(p_status VARCHAR(20)) RETURNS VARCHAR(30) DETERMINISTIC
BEGIN
  RETURN CASE p_status
    WHEN 'ACTIVE'   THEN 'Active'
    WHEN 'INACTIVE' THEN 'Inactive'
    WHEN 'PENDING'  THEN 'Pending Review'
    WHEN 'ERROR'    THEN 'Error'
    ELSE 'Unknown'
  END;
END $$
DELIMITER ;

3. ListRecordsByScoreCategory Procedure
DELIMITER $$
CREATE PROCEDURE ListRecordsByScoreCategory(IN p_category VARCHAR(10))
BEGIN
  SELECT id, name, score, GetScoreCategory(score) AS ScoreCategory
  FROM ${E0}
  WHERE GetScoreCategory(score) = p_category
  ORDER BY score DESC;
END $$
DELIMITER ;

CALL ListRecordsByScoreCategory('High');

4. SummaryByRegionAndScore Procedure
DELIMITER $$
CREATE PROCEDURE SummaryByRegionAndScore()
BEGIN
  SELECT region,
         GetScoreCategory(score)  AS ScoreCategory,
         COUNT(*)                 AS TotalRecords,
         ROUND(AVG(score), 2)     AS AverageScore
  FROM ${E0}
  GROUP BY region, GetScoreCategory(score)
  ORDER BY region, TotalRecords DESC;
END $$
DELIMITER ;

CALL SummaryByRegionAndScore();

5. RecordsByStatus Procedure
DELIMITER $$
CREATE PROCEDURE RecordsByStatus()
BEGIN
  SELECT status,
         COUNT(*)            AS TotalRecords,
         ROUND(AVG(score),2) AS AverageScore
  FROM ${E0}
  GROUP BY status
  ORDER BY TotalRecords DESC;
END $$
DELIMITER ;

CALL RecordsByStatus();

6. RegionStats Procedure
DELIMITER $$
CREATE PROCEDURE RegionStats()
BEGIN
  DECLARE done    INT DEFAULT 0;
  DECLARE v_region VARCHAR(50);
  DECLARE cur CURSOR FOR SELECT DISTINCT region FROM ${E0};
  DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = 1;

  CREATE TEMPORARY TABLE IF NOT EXISTS tmp_stats (
    Region VARCHAR(50), RecordCount INT, AvgScore DECIMAL(6,2)
  );

  OPEN cur;
  read_loop: LOOP
    FETCH cur INTO v_region;
    IF done THEN LEAVE read_loop; END IF;
    INSERT INTO tmp_stats
      SELECT region, COUNT(*), ROUND(AVG(score),2)
      FROM ${E0} WHERE region = v_region GROUP BY region;
  END LOOP;
  CLOSE cur;

  SELECT * FROM tmp_stats ORDER BY RecordCount DESC;
  DROP TEMPORARY TABLE tmp_stats;
END $$
DELIMITER ;

CALL RegionStats();

Power BI Solutions

Data Source: ${E0}_powerbi.csv

Power Query Transformations (M Language):

1. Clean Up Text Fields
= Table.TransformColumns(Source, {
    {"name",    Text.Trim},
    {"region",  Text.Trim},
    {"category",Text.Trim}
  })

2. Handle Missing Category Values
= Table.ReplaceValue(#"Trimmed", null, "NA", Replacer.ReplaceValue, {"category"})

3. Remove Duplicate Records
= Table.Distinct(#"Replaced NA")

4. Standardise Status Values
= Table.ReplaceValue(#"Distinct", null, "Unknown", Replacer.ReplaceValue, {"status"})

5. Convert Value to Numeric
= Table.TransformColumnTypes(#"Replaced Status", {{"value", type number}})

6. Fix Non-Numeric Score Entries
let wordMap = [#"Thirty Four"=34, #"Twenty Seven"=27, #"Sixty"=60, #"Eighty"=80]
in Table.TransformColumns(#"Numeric Value", {{"score", each
    if Value.Is(_, type number) then _ else Record.Field(wordMap, _)}})

DAX Measures and Columns:

7. Flag Invalid Records (DAX Column)
IsInvalidScore = IF([score] < 0 || [score] > 100, TRUE(), FALSE())

8. Average Score Excluding Zero (DAX Measure)
AverageScoreExcludingZero = CALCULATE(AVERAGE('${E0}'[score]), '${E0}'[score] > 0)

9. Validate Status Entries (DAX Column)
IsValidStatus = IF([status] IN {"ACTIVE","INACTIVE","PENDING","ERROR"}, TRUE(), FALSE())

Additional DAX Measures:
Total_Records    = COUNTROWS('${E0}')
Avg_Score        = AVERAGE('${E0}'[score])
Total_Value      = SUM('${E0}'[value])
Active_Count     = CALCULATE(COUNTROWS('${E0}'), '${E0}'[is_active] = "true")
High_Score_Count = CALCULATE(COUNTROWS('${E0}'), '${E0}'[score] > 75)
`);
}