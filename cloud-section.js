/* ===== Google Drive: shared library, automatic sync, conflict handling ===== */
const CLOUD_META_KEY='service-slides-sync';
const GSLIDES='application/vnd.google-apps.presentation';
const CPRES_SCHEMA=1;
const SYNC_DEBOUNCE_MS=5000;   // wait this long after the last edit before uploading
const SYNC_POLL_MS=45000;      // check Drive for remote changes this often
const LIBRARY_STALE_MS=20000;  // don't re-list the folder more often than this when the Services page opens
const sleep=ms=>new Promise(r=>setTimeout(r,ms));

let cloudUi={connected:false,email:'',folderId:'',folderName:'',files:[],error:'',busy:false,msg:'',online:true,authLost:false,lastSyncAt:0,lastAttemptAt:0};
let syncBusy=false,syncAgain=false,pushTimer=null,chipTimer=null;
const syncRt={};     // serviceId -> {state:'syncing'|'offline'|'error', msg, update?:true}
const seenHash={};   // serviceId -> content hash first seen this session (so merely opening an old service never uploads it)

/* ---------- styles for the new UI ---------- */
(function injectCloudStyles(){
  const s=document.createElement('style');
  s.textContent=`
.sync-chip{display:inline-flex;align-items:center;gap:2px;padding:4px 10px;border-radius:999px;font-size:12px;font-weight:600}
.sync-chip[hidden]{display:none}
.sync-chip i,.sync-line i{display:inline-block;width:8px;height:8px;border-radius:50%;background:#8a97a5;margin-right:6px;flex:none}
.sync-line{display:block;margin-top:2px}
.s-synced i{background:#2e9e6b}.s-pending i,.s-readonly i{background:#d89b1a}.s-syncing i{background:#3b82f6;animation:syncpulse 1s infinite}
.s-offline i,.s-local i,.s-paused i{background:#8a97a5}.s-update i{background:#7c5cff}.s-conflict i,.s-error i,.s-lost i{background:#d33}
@keyframes syncpulse{50%{opacity:.3}}
.outline-add{display:flex;gap:6px;margin-top:10px}
.outline-add .pri{flex:1}
.strip-add{display:grid;gap:8px;margin:14px 0 8px}
.strip-add button{width:100%;padding:10px 12px;font-weight:600}
#home-cloud-list{margin-top:22px}
.home-sub{margin:0 0 10px;font-size:13px;text-transform:uppercase;letter-spacing:.05em;color:var(--mu)}
.conflict-actions{display:grid;gap:8px;margin-top:14px}
.cloud-warn{color:#b26a00;font-size:12px;margin-top:4px}
`;
  document.head.appendChild(s);
})();

/* ---------- link metadata (never contains tokens) ---------- */
function loadCloudMeta(){try{const m=JSON.parse(localStorage.getItem(CLOUD_META_KEY)||'{}');m.services=m.services||{};return m}catch{return {services:{}}}}
function saveCloudMeta(meta){try{localStorage.setItem(CLOUD_META_KEY,JSON.stringify(meta))}catch(err){console.warn('cloud meta',err)}}
function getLink(svc){return svc&&svc.id?(loadCloudMeta().services[svc.id]||null):null}
function patchLink(svc,patch){const m=loadCloudMeta();const l=svc&&m.services[svc.id];if(!l)return;Object.assign(l,patch);saveCloudMeta(m)}
function recordLink(svc,patch){const m=loadCloudMeta();const old=m.services[svc.id]||{};m.services[svc.id]={...old,...patch,conflict:false,lost:null,syncedAt:Date.now()};saveCloudMeta(m)}
function linkedServices(){return services.filter(s=>getLink(s))}
function pruneLinks(){
  const ids=new Set(services.map(s=>s.id).filter(Boolean));
  const m=loadCloudMeta();let ch=false;
  Object.keys(m.services).forEach(k=>{if(!ids.has(k)){delete m.services[k];ch=true}});
  if(ch)saveCloudMeta(m);
}

/* ---------- small helpers ---------- */
function bytesToB64(bytes){let s='';const c=0x8000;for(let i=0;i<bytes.length;i+=c)s+=String.fromCharCode.apply(null,bytes.subarray(i,i+c));return btoa(s)}
function b64ToBytes(b64){const bin=atob(b64);const out=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)out[i]=bin.charCodeAt(i);return out}
function hashStr(s){let h=2166136261;for(let i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619)}return (h>>>0).toString(16)+':'+s.length}
function ensureSvcId(svc){if(!svc.id)svc.id='svc-'+Date.now().toString(36)+Math.random().toString(36).slice(2,7);return svc.id}
function fmtDriveDate(x){if(!x)return 'unknown';try{return new Date(x).toLocaleString(undefined,{day:'numeric',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'})}catch{return String(x)}}
function remoteRev(r){return (r&&(r.headRevisionId||r.version))||''}
/** Did the cloud CONTENT change since our last sync? (a rename alone does not change the revision) */
function remoteChangedVs(link,r){
  const rev=remoteRev(r);
  if(rev&&link.remoteRev)return rev!==link.remoteRev;
  return (r.modifiedTime||'')!==(link.remoteModified||'');
}
function driveKind(f){
  const n=(f.name||'').toLowerCase(),m=f.mimeType||'';
  if(m==='application/vnd.google-apps.folder')return 'folder';
  if(n.endsWith('.cpres'))return 'cpres';
  if(m===GSLIDES||/\.(pptx?|pdf)$/.test(n)||m==='application/pdf')return 'slides';
  return 'other';
}
function ownerName(){
  const p=((cloudUi.email||'').split('@')[0]||'').split(/[._-]/)[0]||'My';
  return p.charAt(0).toUpperCase()+p.slice(1);
}

/** Make sure every non-song slide has its elements array so hashing is stable and idempotent. */
function normalizeSlides(slides){
  (slides||[]).forEach(it=>{
    if(!it||it.k==='Song')return;
    (it.parts||[]).forEach(p=>{try{ensureElements(p,it)}catch{}});
  });
}
function serviceSyncBody(i){
  const svc=services[i];
  const slides=i===activeService?S:(svc.slides||[]);
  normalizeSlides(slides);
  return {name:svc.name||'',date:svc.date||'',tag:svc.tag||'',team:cl(svc.team||defaultTeam()),slides:sanitizeSlidesForSave(slides)};
}
function hashService(i){return hashStr(JSON.stringify(serviceSyncBody(i)))}
function cpresFileName(svc){
  const base=(formatServiceTitle(svc)+' '+(svc.date||'')).replace(/[\\/:*?"<>|]+/g,'-').trim()||'service';
  return base+'.cpres';
}
function missingMedia(slides){
  const out=[];
  const has=id=>id&&mediaURLs[id];
  (slides||[]).forEach(it=>{
    if(it.bgMedia&&it.bgMedia.id&&!has(it.bgMedia.id))out.push(it.bgMedia.name||'background');
    (it.parts||[]).forEach(p=>{
      if(p.bgMedia&&p.bgMedia.id&&!has(p.bgMedia.id))out.push(p.bgMedia.name||'background');
      if(p.videoMedia&&p.videoMedia.id&&!has(p.videoMedia.id))out.push(p.videoMedia.name||'video');
      (p.im||[]).forEach(m=>{if(m.mediaId&&!has(m.mediaId))out.push('image')});
      (p.elements||[]).forEach(el=>{if(el.kind==='image'&&el.mediaId&&!has(el.mediaId))out.push('image')});
    });
  });
  return out;
}

/* ---------- talking to Rust, with classification, friendly errors and bounded retry ---------- */
async function cloudInvoke(cmd,args){
  try{
    if(!(window.__TAURI_INTERNALS__||window.__TAURI__))return {ok:false,error:'Desktop app only'};
    const {invoke}=await import('@tauri-apps/api/core');
    const res=await invoke(cmd,args||{});
    if(res&&typeof res==='object')return res;
    return {ok:true,value:res};
  }catch(err){
    const msg=(err&&(err.message||err))||String(err);
    return {ok:false,error:typeof msg==='string'?msg:JSON.stringify(msg)};
  }
}
function classifyErr(msg){
  const m=String(msg||'');
  const code=c=>new RegExp('\\('+c+'\\)|"code":\\s*'+c+'\\b').test(m);
  if(/Desktop app only/i.test(m))return 'desktop';
  if(code(401)||/invalid_grant|unauthorized|Not connected to Google|expired or revoked|Missing Google Client ID/i.test(m))return 'auth';
  if(code(404)||/notFound|not found/i.test(m))return 'notfound';
  if(code(403)&&/rateLimit|quota/i.test(m))return 'quota';
  if(code(403)||/forbidden|insufficientPermissions|access denied/i.test(m))return 'denied';
  if(code(429)||code(500)||code(502)||code(503)||code(504)||/backendError|rateLimitExceeded/i.test(m))return 'transient';
  if(/could not be resolved|unable to connect|no such host|name resolution|network|timed out|timeout|connection|socket|ssl|tls/i.test(m))return 'offline';
  return 'error';
}
const ERR_TEXT={
  auth:'Google sign-in expired or was revoked. Reconnect Google to resume syncing. Your local work is safe.',
  denied:'Google Drive denied access to that file or folder.',
  notfound:'That file or folder no longer exists, or you no longer have access to it.',
  offline:'Google Drive can’t be reached. Changes are saved on this computer and will sync when you’re back online.',
  quota:'Google API quota exceeded. Sync will retry later.',
  transient:'Google Drive is busy. Sync will retry shortly.',
  desktop:'Google Drive sync only works in the desktop app.'
};
function cloudErr(raw){
  const kind=classifyErr(raw);
  return Object.assign(new Error(ERR_TEXT[kind]||String(raw||'Unknown error').replace(/^HTTP_ERROR:/,'').slice(0,300)),{kind,raw:String(raw||'').slice(0,500)});
}
const kindErr=(kind,message)=>Object.assign(new Error(message),{kind});
/** Retries only transient failures (429 / 5xx) with exponential backoff; permanent errors fail at once. */
async function cloudCall(cmd,args,{retries=3}={}){
  let delay=1000,last='';
  for(let a=0;a<=retries;a++){
    const res=await cloudInvoke(cmd,args);
    if(res&&res.ok!==false)return res;
    last=(res&&res.error)||'Unknown error';
    if(classifyErr(last)!=='transient')break;
    if(a<retries){await sleep(delay);delay*=2}
  }
  throw cloudErr(last);
}

/* ---------- per-service sync state ---------- */
function setRt(svc,state,msg){if(svc&&svc.id)syncRt[svc.id]={state,msg:msg||'',at:Date.now(),update:(syncRt[svc.id]||{}).update}}
function clearRtState(svc){if(svc&&svc.id&&syncRt[svc.id]){const u=syncRt[svc.id].update;if(u)syncRt[svc.id]={update:true};else delete syncRt[svc.id]}}
function svcSyncStatus(i){
  const svc=services[i];
  if(!svc)return {key:'local',label:'Local only'};
  const link=getLink(svc);
  if(!link)return {key:'local',label:'Local only'};
  if(live)return {key:'paused',label:'Sync paused while presenting'};
  const rt=syncRt[svc.id]||{};
  if(rt.state==='syncing')return {key:'syncing',label:'Syncing…'};
  if(link.conflict)return {key:'conflict',label:'Conflict — needs your decision'};
  if(link.lost)return {key:'lost',label:link.lost==='deleted'?'Deleted in Drive (local copy kept)':'Access lost (local copy kept)'};
  const dirty=hashService(i)!==link.localHash;
  if(!cloudUi.connected)return {key:'offline',label:dirty?'Saved locally · Drive not connected':'Drive not connected'};
  if(link.canEdit===false)return dirty?{key:'readonly',label:'View-only file · changes kept locally'}:{key:'synced',label:'Synced (view-only)'};
  if(rt.state==='offline'||cloudUi.online===false)return {key:'offline',label:dirty?'Offline · changes pending':'Offline'};
  if(rt.state==='error')return {key:'error',label:'Sync error',detail:rt.msg};
  if(rt.update)return {key:'update',label:'Update available from Drive'};
  if(dirty)return {key:'pending',label:'Sync pending'};
  return {key:'synced',label:'Synced'};
}
function ensureSyncChip(){
  if($('#sync-chip'))return;
  const t=$('#service-title-edit');
  if(t)t.insertAdjacentHTML('afterend','<button type="button" id="sync-chip" class="sync-chip" hidden title="Google Drive sync status"></button>');
}
function renderSyncChip(){
  ensureSyncChip();
  const chip=$('#sync-chip');if(!chip)return;
  const svc=services[activeService];
  if(!svc||document.body.classList.contains('home-mode')){chip.hidden=true;return}
  const st=svcSyncStatus(activeService);
  if(st.key==='local'&&!cloudUi.connected){chip.hidden=true;return}
  const link=getLink(svc);
  chip.hidden=false;
  chip.className='sync-chip s-'+st.key;
  chip.innerHTML='<i></i>'+e(st.label);
  chip.title=(st.detail?st.detail+'\n':'')+(link&&link.syncedAt?'Last synced '+fmtDriveDate(link.syncedAt):'Click to open Google Drive options');
}
function scheduleChip(){clearTimeout(chipTimer);chipTimer=setTimeout(()=>{try{renderSyncChip()}catch{}},300)}
function refreshSyncUi(){
  try{renderSyncChip()}catch{}
  try{renderCloudModal()}catch{}
  try{if(document.body.classList.contains('home-mode'))rHome()}catch{}
}
function cloudReady(){return !!(cloudUi.connected&&cloudUi.folderId&&!cloudUi.authLost)}

function handleCloudError(err,svc){
  const kind=err.kind||classifyErr(err.message||err);
  console.warn('[cloud]',kind,String(err.raw||err.message||err).slice(0,300));   // never contains tokens
  if(kind==='conflict'||kind==='cancelled'||kind==='readonly')return;
  cloudUi.error=err.message||String(err);
  if(kind==='connect')return;
  if(kind==='auth'){cloudUi.authLost=true;linkedServices().forEach(s=>setRt(s,'error',cloudUi.error))}
  else if(kind==='offline'){cloudUi.online=false;linkedServices().forEach(s=>setRt(s,'offline',cloudUi.error))}
  else if(svc)setRt(svc,'error',cloudUi.error);
}

/** One cloud job at a time. quiet = background work (no disabled buttons, no toasts). */
async function withCloudLock(label,fn,{quiet=false}={}){
  if(syncBusy){syncAgain=true;if(!quiet)toast('Sync in progress — try again in a moment.');return false}
  syncBusy=true;
  if(!quiet){cloudUi.busy=true;cloudUi.msg=label;cloudUi.error=''}
  refreshSyncUi();
  let ok=false;
  try{await fn();ok=true}
  catch(err){
    if(!(err&&err.cancelled)){handleCloudError(err);if(!quiet&&err.kind!=='conflict')toast(err.message||'Google Drive error.')}
  }
  finally{
    syncBusy=false;cloudUi.busy=false;cloudUi.msg='';
    refreshSyncUi();
    if(syncAgain){syncAgain=false;setTimeout(()=>cloudSyncAll({quiet:true}),800)}
  }
  return ok;
}

/* ---------- status / listing ---------- */
async function refreshCloudStatus(){
  const st=await cloudInvoke('gdrive_auth_status');
  if(st&&st.ok){
    cloudUi.connected=!!st.connected;cloudUi.email=st.email||'';
    cloudUi.folderId=st.folderId||loadCloudMeta().rootFolderId||'';
    cloudUi.folderName=st.folderName||'';
    if(cloudUi.connected)cloudUi.authLost=false;
  }else{cloudUi.connected=false;cloudUi.email=''}
  refreshSyncUi();
}
async function fetchCloudFiles(){
  if(!cloudUi.connected||!cloudUi.folderId)return;
  try{
    const res=await cloudCall('gdrive_list_presentations',{folderId:cloudUi.folderId});
    cloudUi.files=Array.isArray(res.files)?res.files:[];
    cloudUi.online=true;
  }catch(err){
    if(err.kind==='notfound'||err.kind==='denied'){
      throw Object.assign(new Error('The Drive folder is missing, was moved, or you no longer have access. Open Google Drive options and choose the folder again.'),{kind:err.kind,raw:err.raw});
    }
    throw err;
  }
}

/* ---------- reading / writing the .cpres format ---------- */
async function downloadDriveFile(f){
  const res=await cloudCall('gdrive_download_file',{fileId:f.id,mimeType:f.mimeType||''});
  if(!res.dataBase64)throw cloudErr('Download failed');
  let name=f.name||'file';
  if(f.mimeType===GSLIDES&&!/\.pptx$/i.test(name))name+='.pptx';
  return new File([b64ToBytes(res.dataBase64)],name);
}
async function readCpres(file){
  let data;
  try{data=JSON.parse(await file.text())}catch{throw kindErr('corrupt','This cloud file is corrupt or is not a Church Presenter service.')}
  if(!data||(data.format!=='church-presenter-service'&&!data.service&&!Array.isArray(data.slides)))throw kindErr('corrupt','This cloud file is not a Church Presenter service.');
  if((data.version||1)>CPRES_SCHEMA)throw kindErr('schema','This file was saved by a newer version of Church Presenter. Update the app to open it.');
  const s=data.service||data;
  return {name:(s.name||'').trim(),date:s.date||new Date().toISOString().slice(0,10),tag:s.tag||'',team:s.team?cl(s.team):defaultTeam(),slides:Array.isArray(s.slides)?cl(s.slides):[]};
}
function buildCpresJson(i,body){
  body=body||serviceSyncBody(i);
  return JSON.stringify({format:'church-presenter-service',version:CPRES_SCHEMA,exportedAt:new Date().toISOString(),service:{id:services[i].id,...body}},null,2);
}

/* ---------- local backups (so cloud replacements are always recoverable) ---------- */
async function backupService(i){
  const svc=services[i];ensureSvcId(svc);
  const id='backup-'+svc.id+'-'+Date.now().toString(36);
  await mPutFile(id,new Blob([buildCpresJson(i)],{type:'application/json'}));
  const m=loadCloudMeta();const l=m.services[svc.id];
  if(!l)return;
  l.backups=(l.backups||[]).concat([{id,at:Date.now()}]);
  while(l.backups.length>3){const o=l.backups.shift();try{await mDeleteFile(o.id)}catch{}}
  saveCloudMeta(m);
}
function applyServiceData(i,s){
  const svc=services[i];
  if(i===activeService)pushUndo();
  svc.name=s.name||'';svc.date=s.date||svc.date;svc.tag=s.tag||'';
  svc.team=s.team?cl(s.team):defaultTeam();ensureTeam(svc);
  svc.slides=s.slides.length?cl(s.slides):[mkS('Welcome',['Welcome'],0,6.5)];
  if(i===activeService){S=cl(svc.slides);sel=Math.min(sel,S.length-1);si=0;A=S[sel]||null;aj=0;elSel=null;rAll()}
  else saveState();
}
async function cloudRestoreBackup(i){
  const link=getLink(services[i]);
  const b=link&&link.backups&&link.backups[link.backups.length-1];
  if(!b)return toast('No local backup available.');
  if(!confirm('Replace this service with your local backup from '+fmtDriveDate(b.at)+'?\n\nThe current version is backed up first.'))return;
  const blob=await mGetFile(b.id);
  if(!blob)return toast('The backup file is missing.');
  const data=await readCpres(new File([blob],'backup.cpres'));
  await backupService(i);
  applyServiceData(i,data);
  toast('Restored local backup.');
}

/* ---------- pull / push ---------- */
async function cloudPull(i,fileId){
  const svc=services[i];
  const info=await cloudCall('gdrive_file_info',{fileId});          // read the revision FIRST so a newer one is never recorded as ours
  const file=await downloadDriveFile({id:fileId,name:'service.cpres',mimeType:'application/json'});
  const data=await readCpres(file);                                   // validates before anything local changes
  await backupService(i);
  applyServiceData(i,data);
  recordLink(svc,{fileId,remoteRev:remoteRev(info),remoteModified:info.modifiedTime,localHash:hashService(i),canEdit:info.canEdit!==false,webViewLink:info.webViewLink||''});
  if(syncRt[svc.id])delete syncRt[svc.id];
  toast('Loaded the latest version from Drive.');
}

async function cloudPush(i,{force=false,asCopy=false}={}){
  const svc=services[i];ensureSvcId(svc);
  let link=getLink(svc);
  let existingId=asCopy?null:((link&&link.fileId)||null);
  try{
    if(asCopy){
      const old=svc.id;
      svc.name=svc.name?(svc.name+' - '+ownerName()+'’s Copy'):(ownerName()+'’s Copy');
      svc.id=null;ensureSvcId(svc);
      const m=loadCloudMeta();delete m.services[old];saveCloudMeta(m);
      link=null;
    }
    // Idempotent create: if an earlier (interrupted) upload of this service exists, adopt it instead of making a duplicate
    if(!existingId&&!asCopy){
      await fetchCloudFiles();
      const adopt=cloudUi.files.find(f=>f.cpresId&&f.cpresId===svc.id);
      if(adopt){
        const remote=await readCpres(await downloadDriveFile(adopt));
        const mine=hashService(i);
        const same=hashStr(JSON.stringify({name:remote.name,date:remote.date,tag:remote.tag,team:remote.team,slides:remote.slides}))===mine;
        recordLink(svc,{fileId:adopt.id,remoteRev:remoteRev(adopt),remoteModified:adopt.modifiedTime,localHash:same?mine:'',canEdit:adopt.canEdit!==false,webViewLink:adopt.webViewLink||''});
        if(same){toast('Linked to the existing Drive copy.');return}
        markConflict(i,adopt,false);throw kindErr('conflict','Conflict');
      }
    }
    if(existingId&&!force){
      const info=await cloudCall('gdrive_file_info',{fileId:existingId});
      if(info.trashed)throw kindErr('notfound','The cloud copy was deleted.');
      if(link&&remoteChangedVs(link,info)){markConflict(i,info,false);throw kindErr('conflict','Conflict')}
      if(info.canEdit===false){patchLink(svc,{canEdit:false});throw kindErr('readonly','You only have view access to this Drive file, so changes stay on this computer. Use “Save as copy” to upload them.')}
    }
    syncRt[svc.id]={state:'syncing',msg:'',at:Date.now(),update:(syncRt[svc.id]||{}).update};
    refreshSyncUi();
    const body=serviceSyncBody(i);
    const hash=hashStr(JSON.stringify(body));     // hash of exactly what we upload; edits made during the upload stay "pending"
    const json=buildCpresJson(i,body);
    const res=await cloudCall('gdrive_upload_file',{
      name:cpresFileName(svc),
      dataBase64:bytesToB64(new TextEncoder().encode(json)),
      mimeType:'application/json',
      folderId:existingId?null:cloudUi.folderId,
      existingFileId:existingId,
      cpresId:svc.id
    });
    recordLink(svc,{fileId:res.id,remoteRev:remoteRev(res),remoteModified:res.modifiedTime,localHash:hash,canEdit:true,webViewLink:res.webViewLink||(link&&link.webViewLink)||''});
    delete syncRt[svc.id];
    if(!force&&!asCopy&&existingId){/* silent auto-save */}else toast('Saved to Drive: “'+cpresFileName(svc)+'”.');
  }catch(err){
    if(existingId&&!asCopy){
      if(err.kind==='notfound')markLost(svc,'deleted');
      else if(err.kind==='denied')markLost(svc,'access');
    }
    if(syncRt[svc.id]&&syncRt[svc.id].state==='syncing')delete syncRt[svc.id];
    if(err.kind!=='conflict'&&err.kind!=='readonly')handleCloudError(err,svc);
    throw err;
  }
}

/* ---------- reconcile one linked service with its cloud file ---------- */
function markLost(svc,kind){
  patchLink(svc,{lost:kind});
  delete syncRt[svc.id];
  toast(kind==='deleted'?'“'+formatServiceTitle(svc)+'” was deleted from Drive. Your local copy is kept.':'You no longer have access to “'+formatServiceTitle(svc)+'” on Drive. Your local copy is kept.');
}
function markConflict(i,remote,open){
  const svc=services[i];
  patchLink(svc,{conflict:true,conflictRemote:{rev:remoteRev(remote),modifiedTime:remote.modifiedTime||''}});
  delete syncRt[svc.id];
  toast('Conflict: “'+formatServiceTitle(svc)+'” was also changed on another computer. Your local changes are kept.');
  refreshSyncUi();
  if(open)showConflictDialog(i);
}
async function reconcileService(i,{manual=false}={}){
  const svc=services[i],link=getLink(svc);
  if(!link||link.conflict)return;
  if(link.lost&&!manual)return;
  let remote=cloudUi.files.find(f=>f.id===link.fileId);
  if(!remote){          // not in the folder listing: deleted, moved elsewhere, or access lost
    try{
      remote=await cloudCall('gdrive_file_info',{fileId:link.fileId});
      if(remote.trashed)throw kindErr('notfound','trashed');
    }catch(err){
      if(err.kind==='notfound'){markLost(svc,'deleted');return}
      if(err.kind==='denied'){markLost(svc,'access');return}
      throw err;
    }
  }
  const canEdit=remote.canEdit!==false;
  patchLink(svc,{lost:null,canEdit,webViewLink:remote.webViewLink||link.webViewLink||''});
  const changed=remoteChangedVs(link,remote);
  const dirty=hashService(i)!==link.localHash;
  if(!changed){
    patchLink(svc,{remoteRev:remoteRev(remote)||link.remoteRev,remoteModified:remote.modifiedTime||link.remoteModified});
    if(syncRt[svc.id])delete syncRt[svc.id].update;
    if(dirty&&canEdit)await cloudPush(i);
    return;
  }
  if(!dirty){
    const editing=i===activeService&&!document.body.classList.contains('home-mode');
    if(editing&&!manual){
      // Never swap content under someone who is editing/presenting — flag it and let them choose
      if(!(syncRt[svc.id]&&syncRt[svc.id].update))toast('A newer version of this service is on Drive. Click the sync indicator to load it.');
      syncRt[svc.id]={...(syncRt[svc.id]||{}),update:true};
      return;
    }
    await cloudPull(i,remote.id);
    return;
  }
  markConflict(i,remote,manual);
}

async function cloudSyncAll({manual=false,quiet=false}={}){
  if(live||document.body.classList.contains('presentation-mode'))return false;
  if(!cloudReady()){
    if(manual){toast(cloudUi.connected?'Choose a Drive folder first.':'Connect Google Drive first.');openCloudModal()}
    return false;
  }
  cloudUi.lastAttemptAt=Date.now();
  return withCloudLock('Syncing with Drive…',async()=>{
    await fetchCloudFiles();
    pruneLinks();
    for(let i=0;i<services.length;i++){
      if(live)break;
      const svc=services[i];
      if(!getLink(svc))continue;
      try{await reconcileService(i,{manual})}
      catch(err){if(err.kind==='offline'||err.kind==='auth')throw err;handleCloudError(err,svc)}
    }
    linkedServices().forEach(s=>{const r=syncRt[s.id];if(r&&(r.state==='offline'||r.state==='error'))clearRtState(s)});
    cloudUi.error='';cloudUi.online=true;cloudUi.lastSyncAt=Date.now();
  },{quiet:quiet&&!manual});
}

/* ---------- automatic saving to Drive ---------- */
async function cloudAutoPush(){
  if(live||!cloudReady()||cloudUi.online===false)return;
  const i=activeService,svc=services[i];if(!svc)return;
  ensureSvcId(svc);
  const link=getLink(svc);
  if(link){
    if(link.conflict||link.lost||link.canEdit===false)return;
    if(hashService(i)===link.localHash)return;           // nothing changed — never re-upload unchanged content
  }else{
    if(loadCloudMeta().autoUpload===false)return;
    if(seenHash[svc.id]===hashService(i))return;         // opened but untouched this session
  }
  await withCloudLock('Saving to Drive…',()=>cloudPush(i),{quiet:true});
}
function cloudAfterSave(){
  if(IS_AUDIENCE||live)return;
  const svc=services[activeService];if(!svc)return;
  if(!svc.id){ensureSvcId(svc);try{_saveStateOrig()}catch{}}
  if(!(svc.id in seenHash))seenHash[svc.id]=hashService(activeService);
  scheduleChip();
  clearTimeout(pushTimer);
  pushTimer=setTimeout(cloudAutoPush,SYNC_DEBOUNCE_MS);
}
const _saveStateOrig=saveState;
saveState=function(){const r=_saveStateOrig.apply(this,arguments);try{cloudAfterSave()}catch(err){console.warn('cloudAfterSave',err)}return r};

/* ---------- conflict dialog ---------- */
function ensureConflictModal(){
  if($('#conflict-modal'))return;
  document.body.insertAdjacentHTML('beforeend',`<div id="conflict-modal" class="app-modal" hidden><div class="app-dialog" role="dialog" aria-modal="true" aria-labelledby="conflict-title" style="width:min(520px,calc(100vw - 32px))"><header><h2 id="conflict-title">Conflict detected</h2><button type="button" id="conflict-close">Resolve later</button></header><div id="conflict-body"></div></div></div>`);
}
function showConflictDialog(i){
  if(live)return;
  const svc=services[i],link=getLink(svc);if(!svc||!link)return;
  ensureConflictModal();
  const m=$('#conflict-modal');m.dataset.svc=svc.id;
  $('#conflict-body').innerHTML=`<p style="margin:0 0 8px">“${e(formatServiceTitle(svc))}” has been changed on another computer.</p>
    <p style="margin:0 0 8px">Your local version was last modified at <b>${e(fmtDriveDate(svc.lastEdited))}</b>.<br>The cloud version was modified at <b>${e(fmtDriveDate(link.conflictRemote&&link.conflictRemote.modifiedTime))}</b>.</p>
    <p class="note">Your local changes have been preserved.</p>
    <div class="conflict-actions">
      <button type="button" data-conflict="cloud">Use cloud version <small style="display:block;color:var(--mu);font-weight:400">Replaces this computer’s version (a backup is kept)</small></button>
      <button type="button" data-conflict="mine">Keep my version <small style="display:block;color:var(--mu);font-weight:400">Overwrites the cloud file with yours (Drive keeps its version history)</small></button>
      <button type="button" class="pri" data-conflict="copy">Save my version as a copy <small style="display:block;opacity:.85;font-weight:400">Uploads yours as a new file; the other version stays untouched</small></button>
      <button type="button" data-conflict="later">Resolve later</button>
    </div>`;
  m.hidden=false;
}
function closeConflict(){const m=$('#conflict-modal');if(m)m.hidden=true}
async function resolveConflict(action){
  const m=$('#conflict-modal');if(!m)return;
  const i=services.findIndex(s=>s.id===m.dataset.svc);
  if(i<0||action==='later')return closeConflict();
  const svc=services[i],link=getLink(svc);if(!link)return closeConflict();
  if(action==='cloud'){
    if(!confirm('Replace this computer’s version with the cloud version?\n\nYour unsynced local changes will be replaced. A local backup is kept and can be restored from the Google Drive menu.'))return;
    closeConflict();
    await withCloudLock('Downloading the cloud version…',()=>cloudPull(i,link.fileId));
  }else if(action==='mine'){
    if(!confirm('Overwrite the cloud version with yours?\n\nThe other computer’s changes will be replaced in Drive (they stay in Drive’s version history).'))return;
    closeConflict();
    await withCloudLock('Uploading your version…',()=>cloudPush(i,{force:true}));
  }else if(action==='copy'){
    closeConflict();
    await withCloudLock('Saving a copy to Drive…',()=>cloudPush(i,{asCopy:true}));
  }
}

/* ---------- Google Drive modal ---------- */
function openCloudModal(){
  const m=$('#cloud-modal');if(!m)return;
  m.hidden=false;
  renderCloudModal();
  refreshCloudStatus().then(()=>{if(cloudReady())cloudSyncAll({quiet:true})});
}
function closeCloudModal(){const m=$('#cloud-modal');if(m)m.hidden=true}
function cloudStatusLine(){
  if(cloudUi.busy)return cloudUi.msg||'Working…';
  if(!cloudUi.connected)return 'Connect a Google account to share services between computers through a Drive folder. Everything stays usable offline.';
  let t='Connected as '+cloudUi.email;
  if(cloudUi.folderName)t+=' · Folder: '+cloudUi.folderName;
  if(cloudUi.online===false)t+=' · Offline';
  if(cloudUi.lastSyncAt)t+=' · Last checked '+fmtDriveDate(cloudUi.lastSyncAt);
  return t;
}
function thisServiceCard(dis){
  const i=activeService,svc=services[i];if(!svc)return '';
  const link=getLink(svc),st=svcSyncStatus(i);
  const btn=(act,label,pri)=>`<button type="button"${pri?' class="pri"':''} data-cloud-act="${act}"${dis}>${label}</button>`;
  let b='';
  if(!link)b+=btn('upload','Upload to Drive',true);
  else if(link.conflict)b+=btn('conflict','Resolve conflict…',true);
  else if(link.lost)b+=btn('unlink','Keep local only')+btn('reupload','Upload as new file');
  else if(st.key==='update')b+=btn('pull','Load update from Drive',true);
  else if(link.canEdit===false)b+=btn('copy','Save as copy');
  else b+=btn('sync','Sync now',true);
  if(link&&link.webViewLink)b+=btn('open','Open in Drive')+btn('copylink','Copy link');
  if(link&&link.backups&&link.backups.length)b+=btn('restore','Restore local backup');
  const miss=link?missingMedia(i===activeService?S:svc.slides):[];
  return `<div class="service-card" style="margin-bottom:14px;flex-wrap:wrap"><div class="info"><h3 style="font-size:14px">This service: ${e(formatServiceTitle(svc))}</h3>
    <small class="sync-line s-${st.key}"><i></i>${e(st.label)}${link&&link.syncedAt?' · last synced '+e(fmtDriveDate(link.syncedAt)):''}</small>
    ${st.detail?`<small style="color:#c22">${e(st.detail)}</small>`:''}
    ${miss.length?`<div class="cloud-warn">⚠ ${miss.length} media item${miss.length===1?'':'s'} (${e([...new Set(miss)].slice(0,4).join(', '))}) not on this computer. Media files are not synced through Drive.</div>`:''}
    </div><div style="display:flex;gap:6px;flex-wrap:wrap">${b}</div></div>`;
}
function renderCloudModal(){
  const m=$('#cloud-modal');if(!m||m.hidden)return;
  const status=$('#cloud-status'),acct=$('#cloud-account-row'),folder=$('#cloud-folder-row'),lib=$('#cloud-library');
  if(!status||!acct)return;
  const dis=cloudUi.busy?' disabled':'';
  status.textContent=cloudStatusLine();
  if(cloudUi.authLost)status.innerHTML+='<br><span style="color:#c22">Google sign-in expired or was revoked. Use Reconnect — your local services are safe.</span>';
  else if(cloudUi.error)status.innerHTML+='<br><span style="color:#c22">'+e(cloudUi.error)+'</span>';
  acct.innerHTML=cloudUi.connected
    ?(cloudUi.authLost?`<button type="button" id="cloud-connect" class="pri"${dis}>Reconnect Google</button>`:`<button type="button" id="cloud-sync-all" class="pri"${dis}>Sync everything now</button>`)
      +`<button type="button" id="cloud-disconnect"${dis}>Disconnect</button><button type="button" id="cloud-switch"${dis}>Switch account</button>`
    :`<button type="button" id="cloud-connect" class="pri"${dis}>Connect Google</button>`;
  const auto=loadCloudMeta().autoUpload!==false;
  folder.innerHTML=cloudUi.connected
    ?`<button type="button" id="cloud-pick-folder"${dis}>${cloudUi.folderId?'Change folder…':'Set presentation folder…'}</button>
      <p class="note" style="margin:8px 0 0">Paste the folder link or ID from Google Drive. Services are saved there as <b>.cpres</b> files and sync automatically.</p>
      <label style="display:flex;align-items:center;gap:8px;margin-top:8px"><input type="checkbox" id="cloud-auto-upload" ${auto?'checked':''}> Automatically save edited services to Drive</label>`:'';
  if(!lib)return;
  if(!cloudUi.connected){lib.innerHTML='';return}
  let html=thisServiceCard(dis);
  if(!cloudUi.folderId)html+='<p class="note">Choose a presentation folder to get started.</p>';
  else html+='<p class="note">Services stored in your Drive folder appear on the <b>Services</b> page. Use the ↻ button there to sync.</p>';
  const slides=cloudUi.files.filter(f=>driveKind(f)==='slides');
  if(slides.length){
    html+=`<details style="margin-top:12px"><summary style="cursor:pointer;color:var(--mu)">Import a PowerPoint / PDF / Google Slides file from Drive (${slides.length})</summary>`
      +slides.map(f=>`<div class="service-card" style="margin-top:8px"><div class="info"><h3 style="font-size:14px">${e(f.name)}</h3><small>${e(fmtDriveDate(f.modifiedTime))}</small></div><button type="button" class="pri" data-cloud-import="${e(f.id)}"${dis}>Import</button></div>`).join('')
      +'</details>';
  }
  lib.innerHTML=html;
}

/* ---------- connect / disconnect / folder ---------- */
async function cloudConnect(){
  await withCloudLock('Waiting for Google sign-in in your browser…',async()=>{
    const res=await cloudInvoke('gdrive_connect');
    if(!(res&&res.ok!==false&&res.connected!==false))throw kindErr('connect',(res&&res.error)||'Connect failed.');
    const email=res.email||'';
    const meta=loadCloudMeta();
    if(meta.account&&email&&meta.account!==email){
      meta.services={};
      toast('Signed in with a different Google account. Services were unlinked from the previous account (they remain on this computer).');
    }
    meta.account=email||meta.account||'';saveCloudMeta(meta);
    cloudUi.connected=true;cloudUi.email=email;cloudUi.authLost=false;cloudUi.online=true;
    cloudUi.folderId=res.folderId||cloudUi.folderId;cloudUi.folderName=res.folderName||cloudUi.folderName;
  });
  if(cloudReady())cloudSyncAll({quiet:true});
}
async function cloudDisconnect(){
  await withCloudLock('Disconnecting…',async()=>{
    await cloudInvoke('gdrive_disconnect');
    cloudUi.connected=false;cloudUi.email='';cloudUi.files=[];cloudUi.authLost=false;
  });
  toast('Disconnected. All local services were kept.');
}
async function cloudPickFolder(){
  const input=prompt('Paste the Google Drive folder link or ID:',cloudUi.folderId||'');
  if(input==null)return;
  const mm=input.match(/folders\/([\w-]+)/);
  const id=(mm?mm[1]:input).trim();
  if(!id)return toast('Folder ID required');
  await withCloudLock('Checking folder…',async()=>{
    const res=await cloudCall('gdrive_pick_folder',{folderId:id,folderName:null});
    cloudUi.folderId=res.folderId||id;cloudUi.folderName=res.folderName||'Presentation folder';
    const meta=loadCloudMeta();meta.rootFolderId=cloudUi.folderId;saveCloudMeta(meta);
    await fetchCloudFiles();
  });
  if(cloudReady())cloudSyncAll({quiet:true});
}

/* ---------- actions from the modal / Services page ---------- */
async function cloudImportSlides(id){
  const f=cloudUi.files.find(x=>x.id===id);if(!f)return;
  await withCloudLock('Downloading '+f.name+'…',async()=>{
    const file=await downloadDriveFile(f);
    closeCloudModal();
    await importPowerPoint(file);
  });
}
async function cloudOpenAsService(fileId){
  const f=cloudUi.files.find(x=>x.id===fileId);if(!f)return;
  await withCloudLock('Downloading '+f.name+'…',async()=>{
    if(services.some(s=>{const l=getLink(s);return l&&l.fileId===fileId})){toast('That service is already on this computer.');return}
    const s=await readCpres(await downloadDriveFile(f));
    const svc={id:f.cpresId||null,name:s.name,date:s.date,tag:s.tag,team:s.team,slides:s.slides.length?s.slides:[mkS('Welcome',['Welcome'],0,6.5)],lastEdited:Date.now()};
    if(svc.id&&services.some(x=>x.id===svc.id))svc.id=null;
    ensureSvcId(svc);ensureTeam(svc);
    services.push(svc);
    const i=services.length-1;
    recordLink(svc,{fileId:f.id,remoteRev:remoteRev(f),remoteModified:f.modifiedTime,localHash:hashService(i),canEdit:f.canEdit!==false,webViewLink:f.webViewLink||''});
    saveState();
    toast('Downloaded “'+formatServiceTitle(svc)+'”. It now works offline.');
  });
}
async function cloudAct(act){
  const i=activeService,svc=services[i];if(!svc)return;
  const link=getLink(svc);
  if(act==='upload'){await withCloudLock('Uploading…',()=>cloudPush(i));}
  else if(act==='sync'||act==='pull'){
    await withCloudLock('Syncing…',async()=>{await fetchCloudFiles();await reconcileService(i,{manual:true})});
  }
  else if(act==='conflict')showConflictDialog(i);
  else if(act==='copy'){await withCloudLock('Saving a copy…',()=>cloudPush(i,{asCopy:true}));}
  else if(act==='unlink'){
    if(!confirm('Keep this service on this computer only? The file in Drive is not deleted.'))return;
    const m=loadCloudMeta();delete m.services[svc.id];saveCloudMeta(m);refreshSyncUi();
  }
  else if(act==='reupload'){
    const m=loadCloudMeta();delete m.services[svc.id];saveCloudMeta(m);
    await withCloudLock('Uploading…',()=>cloudPush(i));
  }
  else if(act==='open'&&link&&link.webViewLink){try{await openExternalUrl(link.webViewLink)}catch(err){toast(err.message||'Could not open the link.')}}
  else if(act==='copylink'&&link&&link.webViewLink){
    navigator.clipboard.writeText(link.webViewLink).then(()=>toast('Drive link copied. Recipients still need access in Google Drive.')).catch(()=>toast('Could not copy.'));
  }
  else if(act==='restore'){try{await cloudRestoreBackup(i)}catch(err){toast(err.message||'Restore failed.')}}
}

/* ---------- Services page (home): show Drive contents + sync button ---------- */
function ensureHomeChrome(){
  if($('#home-cloud-refresh'))return;
  const tools=document.querySelector('.service-tools');
  if(tools)tools.insertAdjacentHTML('beforeend','<span class="note" id="home-cloud-status" style="margin-left:auto"></span><button type="button" id="home-cloud-refresh">↻ Sync with Drive</button>');
  const head=document.querySelector('.home-head > div:last-child');
  if(head)head.insertAdjacentHTML('afterbegin','<button type="button" id="home-cloud-btn">Google Drive</button>');
  const list=$('#service-list');
  if(list)list.insertAdjacentHTML('afterend','<div id="home-cloud-list"></div>');
}
function renderHomeCloud(){
  ensureHomeChrome();
  const btn=$('#home-cloud-refresh'),st=$('#home-cloud-status');
  if(btn){btn.disabled=syncBusy||!cloudUi.connected;btn.textContent=syncBusy?'Syncing…':'↻ Sync with Drive';btn.title='Send local changes to Drive and receive changes made on other computers'}
  if(st)st.textContent=cloudUi.connected?(cloudUi.online===false?'Offline':(cloudUi.lastSyncAt?'Last synced '+fmtDriveDate(cloudUi.lastSyncAt):'')):'';
  document.querySelectorAll('#service-list [data-open-service]').forEach(b=>{
    const i=+b.dataset.openService;
    const info=b.closest('.service-card')&&b.closest('.service-card').querySelector('.info');
    if(!info||!services[i])return;
    const old=info.querySelector('.sync-line');if(old)old.remove();
    const s=svcSyncStatus(i);
    if(s.key==='local'&&!cloudUi.connected)return;
    info.insertAdjacentHTML('beforeend','<small class="sync-line s-'+s.key+'"><i></i>'+e(s.label)+'</small>');
  });
  const box=$('#home-cloud-list');if(!box)return;
  if(!cloudUi.connected){
    box.innerHTML='<p class="note">Google Drive isn’t connected, so services are saved on this computer only. <button type="button" data-open-cloud-modal="1">Connect Google Drive</button></p>';return;
  }
  if(!cloudUi.folderId){
    box.innerHTML='<p class="note">Choose a Drive folder to share services between computers. <button type="button" data-open-cloud-modal="1">Choose folder</button></p>';return;
  }
  const linked=new Set(services.map(s=>{const l=getLink(s);return l&&l.fileId}).filter(Boolean));
  const only=cloudUi.files.filter(f=>driveKind(f)==='cpres'&&!linked.has(f.id));
  let html='';
  if(cloudUi.error)html+='<p style="color:#c22;margin:0 0 10px">'+e(cloudUi.error)+'</p>';
  if(only.length){
    html+='<h3 class="home-sub">On Google Drive — not yet on this computer</h3><div class="service-list">'
      +only.map(f=>`<div class="service-card"><div class="info"><h3>${e(f.name.replace(/\.cpres$/i,''))}</h3><small>On Google Drive · ${e(fmtDriveDate(f.modifiedTime))}</small></div><button type="button" class="pri" data-home-cloud-open="${e(f.id)}">Download</button></div>`).join('')+'</div>';
  }else if(cloudUi.lastSyncAt){
    html+='<p class="note">Every service in your Drive folder is on this computer.</p>';
  }
  box.innerHTML=html;
}
const _rHomeOrig=rHome;
rHome=function(){
  _rHomeOrig();
  try{renderHomeCloud()}catch(err){console.warn('renderHomeCloud',err)}
  if(cloudReady()&&!syncBusy&&Date.now()-cloudUi.lastAttemptAt>LIBRARY_STALE_MS)setTimeout(()=>cloudSyncAll({quiet:true}),50);
};

/* ---------- editor layout: Add section / Add slide, Add tab removed ---------- */
function nextSectionName(){
  const used=new Set();
  S.forEach(it=>{const m=/^Section (\d+)$/.exec(it.t||'');if(m)used.add(+m[1])});
  let n=1;while(used.has(n))n++;
  return 'Section '+n;
}
function addSectionAuto(){
  const name=nextSectionName();
  if(!S.length){
    pushUndo();
    const it=mkS(name,['Type your text here'],0,6.5);
    S.push(it);sel=0;si=0;A=it;aj=0;tab='Structure';rAll();return;
  }
  addBlank(name);
}
function stripAddSlide(){
  if(!A||!S.length)return;
  if(isSongItem(A)){toast('Songs: use “Add” in the structure builder to add a section.');openStruct();return}
  addNewSlideInItem(S.indexOf(A),aj);
}
const _rOutlineOrig=rOutline;
rOutline=function(){
  _rOutlineOrig();
  const o=$('#outline');if(!o)return;
  o.insertAdjacentHTML('beforeend','<div class="outline-add"><button type="button" id="outline-add-section" class="pri">+ Add a section</button><button type="button" id="outline-add-more" title="Other section types" aria-label="Other section types">▾</button></div>');
};
const _rStripOrig=rStrip;
rStrip=function(){
  _rStripOrig();
  const st=$('#strip');if(!st||!S.length)return;
  st.insertAdjacentHTML('beforeend','<div class="strip-add"><button type="button" id="strip-add-slide-btn" class="pri">+ Add slide</button><button type="button" id="strip-add-section-btn">+ Add section</button></div>');
};
const _rPanelOrig=rPanel;
rPanel=function(){
  if(tab==='Add')tab='Structure';       // service-level "Add" moved to the section/slide columns
  _rPanelOrig();
  const b=document.querySelector('#panel [data-tab="Add"]');if(b)b.remove();
};
function ensureAddMenu(){
  if(!$('#add-menu')){
    document.body.insertAdjacentHTML('beforeend',`<div id="add-menu" class="outline-menu" role="menu" hidden>
      <button type="button" data-add-menu="song">♪ Song…</button>
      <button type="button" data-add-menu="Welcome">Welcome</button>
      <button type="button" data-add-menu="Announcements">Announcements</button>
      <button type="button" data-add-menu="Bible Reading">Bible Reading</button>
      <button type="button" data-add-menu="Prayer">Prayer</button>
      <button type="button" data-add-menu="Offering">Offering</button>
      <button type="button" data-add-menu="Benediction">Benediction</button>
      <button type="button" data-add-menu="video">🎬 Video…</button>
      <button type="button" data-add-menu="import">⬆ Import slides…</button>
    </div>`);
  }
  if(!$('#add-video-slide-file'))document.body.insertAdjacentHTML('beforeend','<input type="file" id="add-video-slide-file" accept="video/*" hidden>');
}
function hideAddMenu(){const m=$('#add-menu');if(m){m.hidden=true;m.classList.remove('open')}}
function toggleAddMenu(anchor){
  ensureAddMenu();
  const m=$('#add-menu');
  if(m.classList.contains('open')){hideAddMenu();return}
  m.hidden=false;m.classList.add('open');
  const r=anchor.getBoundingClientRect(),w=m.offsetWidth||180,h=m.offsetHeight||300;
  m.style.left=Math.max(6,Math.min(r.left,window.innerWidth-w-8))+'px';
  m.style.top=Math.max(6,Math.min(r.top-h-6,window.innerHeight-h-8))+'px';
}

/* ---------- one click handler for everything above ---------- */
document.addEventListener('click',ev=>{
  const t=ev.target;
  if(!t||!t.closest)return;
  if(!t.closest('#add-menu')&&!t.closest('#outline-add-more'))hideAddMenu();
  if(t.closest('#cloud-btn')||t.closest('#home-cloud-btn')||t.closest('[data-open-cloud-modal]')){openCloudModal();return}
  if(t.closest('#sync-chip')){
    const l=getLink(services[activeService]);
    if(l&&l.conflict)showConflictDialog(activeService);else openCloudModal();
    return;
  }
  if(t.id==='cloud-close'||t.id==='cloud-modal'){closeCloudModal();return}
  if(t.id==='conflict-close'){closeConflict();return}
  const b=t.closest('button');if(!b)return;
  const d=b.dataset;
  if(d.conflict){resolveConflict(d.conflict);return}
  if(b.id==='home-cloud-refresh'){
    cloudSyncAll({manual:true}).then(ok=>{if(ok&&!cloudUi.error)toast('Synced with Google Drive.')});return;
  }
  if(d.homeCloudOpen){cloudOpenAsService(d.homeCloudOpen);return}
  if(b.id==='outline-add-section'||b.id==='strip-add-section-btn'){hideAddMenu();addSectionAuto();return}
  if(b.id==='strip-add-slide-btn'){stripAddSlide();return}
  if(b.id==='outline-add-more'){toggleAddMenu(b);return}
  if(d.addMenu){
    const k=d.addMenu;hideAddMenu();
    if(k==='song')openLibraryModal();
    else if(k==='video')$('#add-video-slide-file').click();
    else if(k==='import'){const f=$('#import-slides-file');if(f)f.click()}
    else addBlank(k);
    return;
  }
  const modal=b.closest('#cloud-modal');if(!modal)return;
  if(b.id==='cloud-connect')cloudConnect();
  else if(b.id==='cloud-disconnect')cloudDisconnect();
  else if(b.id==='cloud-switch')cloudDisconnect().then(()=>cloudConnect());
  else if(b.id==='cloud-pick-folder')cloudPickFolder();
  else if(b.id==='cloud-sync-all')cloudSyncAll({manual:true});
  else if(d.cloudAct)cloudAct(d.cloudAct);
  else if(d.cloudImport)cloudImportSlides(d.cloudImport);
});
document.addEventListener('change',ev=>{
  if(ev.target&&ev.target.id==='cloud-auto-upload'){const m=loadCloudMeta();m.autoUpload=!!ev.target.checked;saveCloudMeta(m)}
});
document.addEventListener('keydown',ev=>{
  if(ev.key!=='Escape')return;
  const c=$('#conflict-modal');if(c&&!c.hidden){ev.preventDefault();closeConflict();return}
  const m=$('#cloud-modal');if(m&&!m.hidden){ev.preventDefault();closeCloudModal()}
  hideAddMenu();
});

/* ---------- start-up: status check, background polling, offline/online handling ---------- */
if(!IS_AUDIENCE){
  ensureSyncChip();
  ensureAddMenu();
  services.forEach(ensureSvcId);
  window.addEventListener('online',()=>{cloudUi.online=true;if(cloudReady())cloudSyncAll({quiet:true})});
  window.addEventListener('offline',()=>{cloudUi.online=false;linkedServices().forEach(s=>setRt(s,'offline',''));refreshSyncUi()});
  window.addEventListener('focus',()=>{if(cloudReady()&&!live&&Date.now()-cloudUi.lastAttemptAt>LIBRARY_STALE_MS)cloudSyncAll({quiet:true})});
  setInterval(()=>{if(cloudReady()&&!live&&!syncBusy)cloudSyncAll({quiet:true})},SYNC_POLL_MS);
  setTimeout(async()=>{
    await refreshCloudStatus();
    if(cloudReady())cloudSyncAll({quiet:true});
  },1500);
}

