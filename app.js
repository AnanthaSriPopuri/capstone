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
`You are an expert capstone project requirement generator for data engineering students.
Generate detailed, realistic capstone project documents.
Respond ONLY with valid JSON — no markdown fences, no preamble.
JSON must have exactly two string keys: "requirements" and "solutions".

COMPLEXITY LEVEL: ${cmeta.label.toUpperCase()}
${cmeta.modifier}

Rules:
- All requirements MUST be user stories: "As a [role], I want to [action] so that [benefit]."
- Follow these 7 categories strictly in "requirements"
- "solutions" must include a US-xxx reference for EVERY story and provide complete working code
- Vary the stories each run — never repeat the same wording
- No Indian names. No real/registered company names.`;

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

/* ── SHARED STYLES XML for makeRichDocx ─────────────────── */
const RICH_STYLES_XML =
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
        <w:sz w:val="20"/>
        <w:color w:val="333333"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>`;

/* ── DOWNLOAD COMBINED .docx ────────────────────────────── */
/* Requirements + Solutions rendered as plain paragraphs;
   Schema rendered as formatted OOXML tables.              */
function downloadCombined(){
  if(!currentReq && !currentSol){ showToast('error','No content to download'); return; }

  /* Convert a plain-text block to OOXML paragraphs */
  function plainToOoxml(text){
    return text.split('\n').map(line => {
      /* Section divider lines (═══ or ━━━) → styled heading paragraph */
      if(/^[═━]{4,}/.test(line)){
        return `<w:p>
          <w:pPr><w:spacing w:before="280" w:after="80"/>
            <w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="DDDDDD"/></w:pBdr>
          </w:pPr>
          <w:r>
            <w:rPr><w:b/><w:sz w:val="20"/><w:color w:val="2E75B6"/>
              <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
            <w:t xml:space="preserve">${escXml(line)}</w:t>
          </w:r>
        </w:p>`;
      }
      /* US-xxx story IDs get slight emphasis */
      if(/^US-[A-Z]+-\d+/.test(line.trim())){
        return `<w:p>
          <w:pPr><w:spacing w:before="120" w:after="40"/></w:pPr>
          <w:r>
            <w:rPr><w:b/><w:sz w:val="19"/><w:color w:val="1F4E79"/>
              <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
            <w:t xml:space="preserve">${escXml(line)}</w:t>
          </w:r>
        </w:p>`;
      }
      /* Normal line */
      return `<w:p>
        <w:pPr><w:spacing w:after="40"/></w:pPr>
        <w:r>
          <w:rPr><w:sz w:val="18"/><w:color w:val="333333"/>
            <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>
          <w:t xml:space="preserve">${escXml(line)}</w:t>
        </w:r>
      </w:p>`;
    }).join('');
  }

  /* Section banner heading */
  function sectionBanner(label){
    return `<w:p>
      <w:pPr>
        <w:spacing w:before="480" w:after="160"/>
        <w:pBdr>
          <w:bottom w:val="single" w:sz="8" w:space="1" w:color="2E75B6"/>
        </w:pBdr>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/><w:sz w:val="30"/><w:color w:val="1F4E79"/>
          <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
        </w:rPr>
        <w:t>${escXml(label)}</w:t>
      </w:r>
    </w:p>`;
  }

  /* Assemble body */
  let bodyXml = '';

  if(currentReq){
    bodyXml += sectionBanner('REQUIREMENTS');
    bodyXml += plainToOoxml(currentReq);
  }

  if(currentSol){
    bodyXml += sectionBanner('SOLUTIONS');
    bodyXml += plainToOoxml(currentSol);
  }

  /* Schema comes last, already OOXML with tables */
  if(currentSchema){
    bodyXml += `<w:p><w:pPr><w:spacing w:before="480"/></w:pPr></w:p>`;
    bodyXml += currentSchema;
  }

  const blob = makeRichDocx(bodyXml, RICH_STYLES_XML);
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
   DEMO CONTENT BUILDERS
══════════════════════════════════════════════════════════ */
function buildDemoReq(sector,entities,company,names){
  const e0=entities[0],e1=entities[1],e2=entities[2],e3=entities[3];
  return `BUSINESS SCENARIO
${company} has been contracted by a leading ${sector} organisation to modernise
their fragmented data infrastructure. Led by ${names[0]} (Chief Data Officer)
and ${names[1]} (VP of Operations) — 6 Agile sprints.

ENTITY SCHEMAS
─────────────────────────────────────────────────────────────────────────
1. ${e0} (CSV)
   id INT | name VARCHAR | status VARCHAR | created_date DATE | region VARCHAR | score FLOAT | is_active BOOLEAN

2. ${e1} (JSON)
   record_id INT | entity_ref INT FK→${e0}(id) | category VARCHAR | value DECIMAL | last_updated TIMESTAMP

3. ${e2} (Excel .xlsx)
   seq_id INT | description TEXT | assigned_to VARCHAR | priority ENUM | completion_pct FLOAT

4. ${e3} (XML / Parquet)
   log_id INT | source_ref INT FK→${e0}(id) | event_type VARCHAR | details JSON | timestamp DATETIME | severity VARCHAR

══════════════════════════════════════════════════════════════════════════
DATA INGESTION AND SHELL SCRIPTING
══════════════════════════════════════════════════════════════════════════

US-ING-01: As a data engineer, I want to ingest data from multiple formats so that data from different sources can be unified into a single processing pipeline.

US-ING-02: As a system administrator, I want to automate ingestion workflows so that data pipelines run without manual intervention and produce repeatable, auditable results.

US-ING-03: As a DevOps engineer, I want to schedule periodic data loading so that the system processes new data at defined intervals without any human trigger.

US-ING-04: As a data architect, I want to store ingested data in a staging
layer with source_file and load_timestamp columns so that raw records can
be audited before entering the cleansed layer.

US-ING-05: As a developer, I want to validate file structure and schema
during ingestion so that malformed files are rejected
early and an error report is written to ingest_errors.log.

══════════════════════════════════════════════════════════════════════════
UNIX COMMANDS — 15 Story-Based Requirements
══════════════════════════════════════════════════════════════════════════

US-UNX-01: As a data analyst, I want to extract all rows from ${e0.toLowerCase()}.csv where status matches ERROR or PENDING so
that problematic records are isolated before the ETL pipeline runs.

US-UNX-02: As a data engineer, I want to compute the average score per region from ${e0.toLowerCase()}.csv so that a regional KPI summary is produced
without loading data into any database.

US-UNX-03: As a system admin, I want to replace all N/A, null, and -- values with 0 in numeric fields of ${e1.toLowerCase()}.txt so that
downstream processes do not fail on non-numeric strings.

US-UNX-04: As a reporting analyst, I want to rank the top 20 category values by frequency from ${e2.toLowerCase()}.csv so that they are ready for filter configuration.

US-UNX-05: As a DevOps engineer, I want to build a daily manifest.txt of recently changed files so that the backup script processes only modified data files.

US-UNX-06: As a data steward, I want record counts for all 4 entity files to be printed to the console and saved to size_report.txt.

US-UNX-07: As a data engineer, I want to compare ${e0.toLowerCase()}_v1.csv
and ${e0.toLowerCase()}_v2.csv so that records changed between pipeline runs
are identified and flagged automatically.

US-UNX-08: As a pipeline engineer, I want to emit a unified schema report from all 4 files so that schema drift is detected before any transformation step.

US-UNX-09: As a data analyst, I want to count distinct values in the last column of ${e1.toLowerCase()}.txt
so that field cardinality is documented per entity.

US-UNX-10: As a QA engineer, I want to compute null density as a percentage per entity file so that a quality threshold check can gate the pipeline run.

US-UNX-11: As a systems engineer, I want to list files exceeding 10 MB and redirect to large_files.txt so that storage alerts are triggered before disk quotas are breached.

US-UNX-12: As a data analyst, I want to extract top-scoring records from ${e0.toLowerCase()}.csv to a flat file without any database dependency.

US-UNX-13: As a DevOps engineer, I want to normalise all category strings in ${e1.toLowerCase()}.txt to
lowercase so that case-inconsistent values are unified before loading.

US-UNX-14: As a pipeline engineer, I want to prepend a source_entity column to every record from all 4 files so that combined_raw.csv carries its originating file label.

US-UNX-15: As a DevOps engineer, I want to run inconsistency detection in parallel across all 4 entity files so that it completes in under 30 seconds on large datasets.

══════════════════════════════════════════════════════════════════════════
SHELL SCRIPTING — 10 Story-Based Requirements (8 analytics · 2 cleaning)
══════════════════════════════════════════════════════════════════════════

US-SH-01: As a data steward, I want a script that iterates all 4 entity files, counts total rows, null rows, and duplicate rows, and writes a formatted audit_report.txt so that data quality KPIs are tracked daily.

US-SH-02: As an operations analyst, I want a script that reads ${e1.toLowerCase()}.txt and computes sum, avg, max of value per category, printing an ASCII KPI table so that metrics are reviewed without any BI tool.

US-SH-03: As a data engineer, I want a script that flags rows where score falls outside 0–100 in ${e0.toLowerCase()}.csv and writes them to anomalies.csv with an appended reason column so that outliers are quarantined before loading.

US-SH-04: As a pipeline engineer, I want a script that merges all 4 entity files into combined_dataset.csv with a prepended source_entity column so that a unified flat file is ready for batch processing.

US-SH-05: As a reporting analyst, I want a script that generates a daily HTML summary of record counts, top 5 categories, and null percentages so that stakeholders receive an email-ready overview each morning.

US-SH-06: As a data engineer, I want a script that accepts a date argument and filters ${e0.toLowerCase()}.csv by created_date, writing results to dated_extract_YYYYMMDD.csv so that point-in-time slices are produced on demand.

US-SH-07: As a QA analyst, I want a script that compares row counts of raw vs cleaned files and outputs a cleaning efficiency report showing records_removed and pct_cleaned per entity so that data-loss is audited.

US-SH-08: As a pipeline engineer, I want a script that monitors the landing directory and triggers the ingestion pipeline on new file arrival, logging each event with timestamp, so that the pipeline is fully event-driven.

US-SH-09: As a data engineer, I want a script that removes rows with empty mandatory fields from ${e0.toLowerCase()}.csv, writes cleaned output to _clean.csv, and logs the removed count so that ETL receives only complete records. (Data-cleaning story 1)

US-SH-10: As a DevOps engineer, I want a script that fixes delimiter inconsistencies in ${e2.toLowerCase()}.csv, normalising all separators to commas, so that the file conforms to CSV specification before downstream processing. (Data-cleaning story 2)

══════════════════════════════════════════════════════════════════════════
FILE HANDLING
══════════════════════════════════════════════════════════════════════════

US-PY-01: As a Python developer, I want to implement file handling to read and
write ${e0.toLowerCase()}.csv so that pre-processing and deduplication can be
performed efficiently.
INPUT:  data/${e0.toLowerCase()}.csv
OUTPUT: data/${e0.toLowerCase()}_clean.csv
SCOPE:  Remove duplicates by id, fix encoding, drop rows with null mandatory fields.

US-PY-02: As a data engineer, I want to process ${e1.toLowerCase()}.json
(80k–1L records) so that nested structures are flattened and saved as ${e1.toLowerCase()}_flat.csv for database import.

US-PY-03: As a QA tester, I want to handle file-related exceptions such as
missing or corrupted ${e2.toLowerCase()}.xlsx files using try-except so that
the pipeline logs a descriptive error and skips rather than crashing.

US-PY-04: As a developer, I want to generate cleaned output files for all 3 datasets in their required formats so that each cleaned file is immediately usable by downstream modules.

US-PY-05: As a data engineer, I want a script that randomly selects a sector and entity at runtime, reads the corresponding dataset, applies cleaning rules, and generates sector-specific user stories so that each execution produces unique, non-repeating requirements.

══════════════════════════════════════════════════════════════════════════
DATA CLEANING AND INCONSISTENCIES
══════════════════════════════════════════════════════════════════════════

US-PS-01: As a data quality analyst, I want to use PySpark to identify null,
missing, and duplicate values across all 4 entity datasets so that overall data
quality improves before analysis begins.

US-PS-02: As a data engineer, I want to standardise inconsistent date, text,
and numerical formats in ${e0.toLowerCase()} so that data becomes uniform across all regions.

US-PS-03: As a compliance analyst, I want to detect and resolve 10 to 15
inconsistencies in each dataset so that reporting remains accurate and compliant with data governance policies.

US-PS-04: As a data engineer, I want to apply business-rule transformation logic — normalising status enums, capping out-of-range scores, and
forward-filling sparse columns — so that downstream analysis operates on validated data.

US-PS-05: As a data analyst, I want cleaned PySpark DataFrames to be validated
with row-count assertions and null-count checks before writing to the cleansed
layer so that unreliable data never reaches the BI layer.

US-PS-06: As a data engineer, I want to verify that the cleansed
${e0.toLowerCase()} DataFrame can produce syntactically valid MySQL INSERT
statements via a PySpark UDF so that database loading is confirmed before the
full batch is submitted.

══════════════════════════════════════════════════════════════════════════
CRUD OPERATIONS — 10+ Stories
══════════════════════════════════════════════════════════════════════════

US-MG-01 through US-MG-11: [Full CRUD stories for ${sector} entities]
See solutions file for implementation code.

══════════════════════════════════════════════════════════════════════════
DATA PROCESSING AND ANALYSIS (RDD · SQL · DataFrame)
══════════════════════════════════════════════════════════════════════════

US-RDD-01 through US-RDD-05: Distributed RDD operations on ${e0.toLowerCase()}
US-SQL-01 through US-SQL-05: SQL analysis across all 4 entity views
US-DF-01 through US-DF-05: DataFrame transformations and joins

══════════════════════════════════════════════════════════════════════════
ADVANCED SQL — 15 Stories (Stored Procedures & Functions)
══════════════════════════════════════════════════════════════════════════

US-ASQL-01 through US-ASQL-15: Stored procedures and functions for ${sector}

══════════════════════════════════════════════════════════════════════════
POWER BI TRANSFORMATIONS AND VISUALIZATION
══════════════════════════════════════════════════════════════════════════

US-PBI-01: Power Query 15-step transformation pipeline for all 4 entities
US-PBI-02: 10 DAX measures including RANKX, DATEADD, TOTALYTD, SWITCH+TRUE()
US-PBI-03: 5-page interactive dashboard with drill-through and forecast
US-PBI-04: Star schema with Calendar dimension for time-intelligence

══════════════════════════════════════════════════════════════════════════
AUTOMATION AND PROJECT GENERATION LOGIC
══════════════════════════════════════════════════════════════════════════

US-AUT-01: Dynamic runtime generation of datasets and requirements
US-AUT-02: Automated requirement authoring without manual intervention
US-AUT-03: Complete Word document generation with AGILE sprint sheet
`;
}

function buildDemoSol(sector,entities){
  const e0=entities[0],e1=entities[1],e2=entities[2],e3=entities[3];
  const db=sector.replace(/ & /g,'_').replace(/ /g,'_').toLowerCase()+'_db';
  return `SOLUTIONS & ANSWERS — ${sector}
Generated: ${new Date().toLocaleString()}

══════════════════════════════════════════════════════════════════════════
DATA INGESTION — SOLUTIONS
══════════════════════════════════════════════════════════════════════════

US-ING-01 — Ingest multiple formats
#!/bin/bash
mkdir -p staging
for f in landing/*.csv; do cp "$f" staging/; echo "CSV ingested: $f"; done
for f in landing/*.json; do cp "$f" staging/; echo "JSON ingested: $f"; done
python3 -c "
import pandas as pd, glob, os
for f in glob.glob('landing/*.xlsx'):
    df = pd.read_excel(f)
    out = 'staging/' + os.path.basename(f).replace('.xlsx','.csv')
    df.to_csv(out, index=False)
    print('XLSX ingested:', f)
"

US-ING-05 — Validate schema
#!/bin/bash
FILE="staging/${e0.toLowerCase()}.csv"
EXPECTED_COLS=7
ACTUAL=$(head -1 "$FILE" | tr ',' '\n' | wc -l)
if [ "$ACTUAL" -ne "$EXPECTED_COLS" ]; then
  echo "SCHEMA ERROR: $FILE" >> logs/ingest_errors.log; exit 1
fi
echo "Schema OK: $FILE"

══════════════════════════════════════════════════════════════════════════
UNIX SOLUTIONS
══════════════════════════════════════════════════════════════════════════

US-UNX-01 — grep ERROR|PENDING
$ grep -iE ',ERROR,|,PENDING,' ${e0.toLowerCase()}.csv > flagged_records.csv

US-UNX-02 — awk average score per region
$ awk -F',' 'NR>1 {sum[$5]+=$6; cnt[$5]++}
  END {for(r in sum) printf "%-25s Avg: %.2f\\n", r, sum[r]/cnt[r]}' \\
  ${e0.toLowerCase()}.csv | sort -t: -k2 -rn

US-UNX-03 — sed normalise nulls
$ sed -E -i "s/(^|,)(N\\/A|null|--)(,|$)/\\10\\3/g" ${e1.toLowerCase()}.txt

US-UNX-14 — prepend source column
$ for f in ${e0.toLowerCase()}.csv ${e1.toLowerCase()}.csv ${e2.toLowerCase()}.csv ${e3.toLowerCase()}.csv; do
    name=$(basename $f .csv)
    tail -n +2 "data/$f" | awk -v src="$name" -F',' '{print src","$0}'
  done > combined_raw.csv

══════════════════════════════════════════════════════════════════════════
SHELL SCRIPTING SOLUTIONS
══════════════════════════════════════════════════════════════════════════

US-SH-01 — Daily audit report
#!/bin/bash
REPORT="reports/audit_report_$(date +%Y%m%d).txt"
mkdir -p reports
echo "=== DATA QUALITY AUDIT $(date) ===" > "$REPORT"
printf "%-40s %8s %8s %8s\\n" "FILE" "TOTAL" "NULLS" "DUPS" >> "$REPORT"
for FILE in data/${e0.toLowerCase()}.csv data/${e1.toLowerCase()}.csv \\
            data/${e2.toLowerCase()}.csv data/${e3.toLowerCase()}.csv; do
  [ -f "$FILE" ] || continue
  TOTAL=$(tail -n +2 "$FILE" | wc -l)
  NULLS=$(awk -F',' '{for(i=1;i<=NF;i++) if($i==""||/null/||/N\\/A/) n++} END{print n+0}' "$FILE")
  DUPS=$(tail -n +2 "$FILE" | sort | uniq -d | wc -l)
  printf "%-40s %8d %8d %8d\\n" "$(basename $FILE)" "$TOTAL" "$NULLS" "$DUPS" >> "$REPORT"
done

US-SH-09 — Remove empty mandatory fields (data-cleaning 1)
#!/bin/bash
INPUT="data/${e0.toLowerCase()}.csv"; OUTPUT="data/${e0.toLowerCase()}_clean.csv"
REMOVED=0; KEPT=0
head -1 "$INPUT" > "$OUTPUT"
while IFS=',' read -r id name status rest; do
  if [[ -n "$id" && "$id" != "null" && -n "$name" ]]; then
    echo "$id,$name,$status,$rest" >> "$OUTPUT"; ((KEPT++))
  else ((REMOVED++)); fi
done < <(tail -n +2 "$INPUT")
echo "Kept:$KEPT Removed:$REMOVED"

US-SH-10 — Fix delimiter inconsistencies (data-cleaning 2)
#!/bin/bash
sed 's/;/,/g; s/\t/,/g' "data/${e2.toLowerCase()}.csv" > "data/${e2.toLowerCase()}_fixed.csv"

══════════════════════════════════════════════════════════════════════════
PYTHON FILE HANDLING SOLUTIONS
══════════════════════════════════════════════════════════════════════════

US-PY-01 — Clean ${e0} CSV
import csv, logging
from pathlib import Path
logging.basicConfig(level=logging.INFO)
INPUT = Path('data/${e0.toLowerCase()}.csv')
OUTPUT = Path('data/${e0.toLowerCase()}_clean.csv')
seen = set(); kept = 0; removed = 0
with open(INPUT, encoding='utf-8', errors='replace') as fin, \\
     open(OUTPUT, 'w', newline='', encoding='utf-8') as fout:
    reader = csv.DictReader(fin)
    writer = csv.DictWriter(fout, fieldnames=reader.fieldnames)
    writer.writeheader()
    for row in reader:
        if not row.get('id') or not row.get('name') or not row.get('status'):
            removed += 1; continue
        if row['id'] in seen: removed += 1; continue
        seen.add(row['id']); writer.writerow(row); kept += 1
logging.info(f"Kept:{kept:,} Removed:{removed:,}")

US-PY-02 — Flatten ${e1} JSON
import json, pandas as pd
with open('data/${e1.toLowerCase()}.json') as f:
    records = json.load(f)
df = pd.json_normalize(records)
df.fillna('', inplace=True)
df.to_csv('data/${e1.toLowerCase()}_flat.csv', index=False)
print(f"Flattened {len(df):,} records")

══════════════════════════════════════════════════════════════════════════
PYSPARK CLEANING SOLUTIONS
══════════════════════════════════════════════════════════════════════════

US-PS-01 — Null and duplicate detection
from pyspark.sql import SparkSession
from pyspark.sql.functions import col, count, when, isnan, isnull
spark = SparkSession.builder.appName("${sector}_QA").getOrCreate()
df = spark.read.csv("data/${e0.toLowerCase()}_clean.csv", header=True, inferSchema=True)
null_counts = df.select([count(when(isnull(c)|isnan(c), c)).alias(c) for c in df.columns])
null_counts.show()
dup_count = df.count() - df.dropDuplicates(['id']).count()
print(f"Duplicates by id: {dup_count:,}")

══════════════════════════════════════════════════════════════════════════
MONGODB SOLUTIONS
══════════════════════════════════════════════════════════════════════════

US-MG-01 — insertMany with index
use ${db}
db.${e0.toLowerCase()}.createIndex({ status:1, region:1 })
db.${e0.toLowerCase()}.insertMany(data_array, { ordered:false })

US-MG-05 — Aggregation KPI
db.${e1.toLowerCase()}.aggregate([
  { $match: { value: { $gt:0 } } },
  { $group: { _id:"$category", total:{$sum:"$value"}, avg:{$avg:"$value"}, count:{$sum:1} } },
  { $sort: { total:-1 } }
])

US-MG-06 — $lookup join
db.${e0.toLowerCase()}.aggregate([
  { $lookup:{ from:"${e1.toLowerCase()}", localField:"id", foreignField:"entity_ref", as:"linked" }},
  { $project:{ name:1, status:1, region:1, linked_count:{$size:"$linked"} }},
  { $out:"combined_${e0.toLowerCase()}" }
])

══════════════════════════════════════════════════════════════════════════
ADVANCED SQL SOLUTIONS
══════════════════════════════════════════════════════════════════════════

US-ASQL-01 — sp_GetRegionalSummary
DELIMITER $$
CREATE PROCEDURE sp_GetRegionalSummary(IN p_region VARCHAR(50))
BEGIN
  SELECT e.region, COUNT(*) total, ROUND(AVG(e.score),2) avg_score,
         SUM(CASE WHEN e.status='ACTIVE' THEN 1 ELSE 0 END) active_count
  FROM ${e0.toLowerCase()} e WHERE e.region=p_region GROUP BY e.region;
END $$
DELIMITER ;

US-ASQL-02 — fn_CalculateRiskScore
DELIMITER $$
CREATE FUNCTION fn_CalculateRiskScore(p_id INT) RETURNS DECIMAL(5,2) DETERMINISTIC
BEGIN
  DECLARE v_base FLOAT DEFAULT 0; DECLARE v_cnt INT DEFAULT 0;
  SELECT COALESCE(score,0) INTO v_base FROM ${e0.toLowerCase()} WHERE id=p_id;
  SELECT COUNT(*) INTO v_cnt FROM ${e1.toLowerCase()} WHERE entity_ref=p_id;
  RETURN ROUND(LEAST((v_base*0.5)+(v_cnt*0.3),100),2);
END $$
DELIMITER ;

US-ASQL-05 — fn_GetStatusLabel
DELIMITER $$
CREATE FUNCTION fn_GetStatusLabel(p_score FLOAT) RETURNS VARCHAR(20) DETERMINISTIC
BEGIN
  RETURN CASE
    WHEN p_score>=90 THEN 'EXCELLENT'
    WHEN p_score>=75 THEN 'GOOD'
    WHEN p_score>=50 THEN 'AVERAGE'
    WHEN p_score>=25 THEN 'POOR'
    ELSE 'CRITICAL' END;
END $$
DELIMITER ;

══════════════════════════════════════════════════════════════════════════
PYSPARK ANALYSIS SOLUTIONS
══════════════════════════════════════════════════════════════════════════

US-RDD-01 — RDD filter + map
sc = SparkSession.builder.appName("${sector}").getOrCreate().sparkContext
raw = sc.textFile("data/${e0.toLowerCase()}_clean.csv")
hdr = raw.first()
data = raw.filter(lambda x:x!=hdr).map(lambda x:x.split(',')) \\
          .filter(lambda f:len(f)>=3 and f[0].strip())
status_dist = data.map(lambda r:(r[2].strip().upper(),1)).reduceByKey(lambda a,b:a+b)
for s,c in sorted(status_dist.collect(),key=lambda x:-x[1]):
    print(f"  {s:<20} {c:>8,}")

US-SQL-01 — PySpark SQL JOIN
spark = SparkSession.builder.appName("${sector}_SQL").getOrCreate()
e0df = spark.read.csv("data/${e0.toLowerCase()}_clean.csv",header=True,inferSchema=True)
e1df = spark.read.json("data/${e1.toLowerCase()}_flat.json")
e0df.createOrReplaceTempView("${e0.toLowerCase()}")
e1df.createOrReplaceTempView("${e1.toLowerCase()}")
spark.sql("""
  SELECT e.region, COUNT(e.id) total, ROUND(AVG(e.score),2) avg_score,
         ROUND(SUM(r.value),2) total_value
  FROM ${e0.toLowerCase()} e JOIN ${e1.toLowerCase()} r ON r.entity_ref=e.id
  GROUP BY e.region ORDER BY total_value DESC""").show()

══════════════════════════════════════════════════════════════════════════
POWER BI DAX SOLUTIONS
══════════════════════════════════════════════════════════════════════════

Total_Records   = COUNTROWS('${e0.toLowerCase()}')
Avg_Score       = AVERAGE('${e0.toLowerCase()}'[score])
MoM_Growth_Pct  = DIVIDE([Total_Records]
                    - CALCULATE([Total_Records],DATEADD('Calendar'[Date],-1,MONTH)),
                    CALCULATE([Total_Records],DATEADD('Calendar'[Date],-1,MONTH)),0)
Rank_By_Region  = RANKX(ALL('${e0.toLowerCase()}'[region]),[Total_Records],,DESC,DENSE)
YTD_Value       = TOTALYTD(SUM('${e1.toLowerCase()}'[value]),'Calendar'[Date])
Pct_Share       = DIVIDE([Total_Records],CALCULATE([Total_Records],ALL('${e0.toLowerCase()}')))
KPI_Status      = SWITCH(TRUE(),[Avg_Score]>=80,"EXCELLENT",[Avg_Score]>=60,"GOOD",
                    [Avg_Score]>=40,"AVERAGE","NEEDS ATTENTION")
Rolling_30d_Avg = CALCULATE(AVERAGE('${e0.toLowerCase()}'[score]),
                    DATESINPERIOD('Calendar'[Date],LASTDATE('Calendar'[Date]),-30,DAY))
`;
}