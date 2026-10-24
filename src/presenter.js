import { loadPresentation, getSlides } from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';
import { isTauri, invoke, convertFileSrc } from '@tauri-apps/api/core';
import { appDataDir, join as joinPath } from '@tauri-apps/api/path';
import { writeFile, remove as removeFile } from '@tauri-apps/plugin-fs';
import { getVersion as getAppVersion } from '@tauri-apps/api/app';
import { availableMonitors, currentMonitor, getCurrentWindow } from '@tauri-apps/api/window';
import { check as checkAppUpdate } from '@tauri-apps/plugin-updater';
import { relaunch } from '@tauri-apps/plugin-process';
import { configureSongSearch, searchSongs, getLyrics } from './data/lyricsSearch.js';

/** True when this window is the audience output opened by Present (Tauri second window). */
const IS_AUDIENCE = new URLSearchParams(location.search).get('audience') === '1';
const MIME={png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',webp:'image/webp',svg:'image/svg+xml'};
const mimeOf=n=>MIME[(String(n).split('.').pop()||'').toLowerCase()]||'application/octet-stream';
async function checkForAppUpdate(){
  if(!isTauri()||IS_AUDIENCE)return;
  const platform=navigator.userAgentData?.platform||navigator.platform||navigator.userAgent;
  const target=/mac/i.test(platform)?'macos-universal':/win/i.test(platform)?'windows-x86_64':null;
  if(!target)return;
  try{
    const update=await checkAppUpdate({target});
    if(!update)return;
    const notes=update.body?`\n\n${update.body}`:'';
    const install=await appConfirm(`Version ${update.version} is available.${notes}\n\nDownload and install it now?`,{title:'Update available',confirmLabel:'Update now'});
    if(!install)return;
    toast('Downloading update…');
    await update.downloadAndInstall();
    await relaunch();
  }catch(error){
    console.warn('[updater] Could not check or install update:',error);
  }
}
if(!IS_AUDIENCE)setTimeout(checkForAppUpdate,8000);
if(IS_AUDIENCE){
  // Hide chrome until audience boot replaces the body (avoids a flash of the editor UI)
  const s=document.createElement('style');
  s.textContent='html,body{background:#000!important}body>*{visibility:hidden!important}';
  document.documentElement.appendChild(s);
}

const $=s=>document.querySelector(s),e=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
let U=0;const uid=()=>'p'+(++U),cl=o=>JSON.parse(JSON.stringify(o));
const BG=[{n:'Deep blue',c:'linear-gradient(160deg,#16305c,#0a1428)'},{n:'Sunrise',c:'linear-gradient(160deg,#d9822b,#6b2a4a)'},{n:'Forest',c:'linear-gradient(160deg,#1f5a48,#0b1f1a)'},{n:'Plain black',c:'#000'},{n:'Waves (video)',v:1},{n:'Colour picker',pick:1}];
const VA={top:'flex-start',middle:'center',bottom:'flex-end'},TYPES=['Verse','Chorus','Bridge','Pre-chorus','Tag','Intro','Outro'];

/** Normalize online lyric text: [Verse 1], (Chorus), VERSE:, etc. → standard headers. */
function isFillerLyricLine(line){
  const t=String(line||'').trim();
  if(!t)return true;
  if(/^ccli\b/i.test(t))return true;
  // Repeat markers, instrumental cues, decorative separators
  if(/^(x\d+|\(\s*x\s*\d+\s*\)|\d+x)$/i.test(t))return true;
  if(/^(repeat|repeats?|instrumental|interlude|spoken|n\.?c\.?|n\/a)$/i.test(t))return true;
  if(/^[\-\—\–\*\_\.=]{2,}$/.test(t))return true;
  if(/^\(.*\)$/.test(t)&&t.length<40&&/\b(repeat|x\d|instrumental|spoken|whisper)\b/i.test(t))return true;
  return false;
}

function normalizeLyricStructure(raw){
  let t=String(raw||'').replace(/\r\n/g,'\n').replace(/\r/g,'\n');
  // Strip common copyright / contributor footers that create junk "verses"
  t=t.replace(/\n+\s*(source:|written by|lyrics by|©|copyright|\(?c\)?)\b[\s\S]*$/i,'');
  // Bracket / paren section markers
  t=t.replace(/^\s*[\[\(]\s*(verse|chorus|bridge|pre-?chorus|prechorus|tag|outro|intro|ending|refrain|hook|v|c|b)\s*(\d*)\s*[\]\)]\s*:?\s*$/gim,
    (_,k,n)=>{
      const key=k.toLowerCase().replace(/\s+/g,'');
      const map={v:'Verse',c:'Chorus',b:'Bridge',prechorus:'Pre-Chorus','pre-chorus':'Pre-Chorus',ending:'Outro',refrain:'Chorus',hook:'Chorus'};
      const name=map[key]||(k[0].toUpperCase()+k.slice(1).toLowerCase());
      return '\n\n'+name+(n?' '+n:'')+'\n';
    });
  t=t.replace(/^\s*(verse|chorus|bridge|pre-?chorus|prechorus|tag|outro|intro|ending|refrain|hook)\s*(\d*)\s*:\s*$/gim,
    (_,k,n)=>{
      const key=k.toLowerCase().replace(/\s+/g,'');
      const map={prechorus:'Pre-Chorus','pre-chorus':'Pre-Chorus',ending:'Outro',refrain:'Chorus',hook:'Chorus'};
      const name=map[key]||(k[0].toUpperCase()+k.slice(1).toLowerCase());
      return '\n\n'+name+(n?' '+n:'')+'\n';
    });
  return t.replace(/\n{3,}/g,'\n\n').trim();
}

/** Score how well-structured a lyric string is (higher = better Verse/Chorus labels). */
function lyricStructureScore(text){
  const t=String(text||'');
  const hdr=/\b(verse|chorus|bridge|pre-?chorus|tag|intro|outro)\b/gi;
  const m=t.match(hdr);
  if(!m)return 0;
  const kinds=new Set(m.map(x=>x.toLowerCase().replace(/\s+/g,'')));
  let score=m.length*2+kinds.size*5;
  if(kinds.has('chorus')||kinds.has('verse'))score+=10;
  if(kinds.has('bridge'))score+=3;
  return score;
}

/** Strip ad-libs / filler so near-identical sections match. */
function coreLyricText(lines){
  return (lines||[]).map(l=>String(l||'').toLowerCase()
    .replace(/["""'']/g,'')
    .replace(/[^\w\s]/g,' ')
    // common vocalist ad-libs
    .replace(/\b(oh+|ooh+|ah+|yeah+|woah+|whoa+|hey+|mmm+|yes|no|sing|repeat)\b/g,' ')
    .replace(/\s+/g,' ').trim()
  ).filter(Boolean).join(' ').replace(/\s+/g,' ').trim();
}
function lyricWordSet(text){
  return new Set(String(text||'').split(/\s+/).filter(w=>w.length>1));
}
/** 0–1 similarity: word Jaccard + containment (handles extra/missing lines). */
function lyricSimilarity(aLines,bLines){
  const a=coreLyricText(aLines),b=coreLyricText(bLines);
  if(!a||!b)return 0;
  if(a===b)return 1;
  // substring containment for short tags
  if(a.length>=8&&b.length>=8){
    if(a.includes(b)||b.includes(a)){
      const ratio=Math.min(a.length,b.length)/Math.max(a.length,b.length);
      if(ratio>=0.45)return 0.85+0.15*ratio;
    }
  }
  const sa=lyricWordSet(a),sb=lyricWordSet(b);
  if(!sa.size||!sb.size)return 0;
  let inter=0;
  sa.forEach(w=>{if(sb.has(w))inter++});
  const union=sa.size+sb.size-inter;
  const jaccard=union?inter/union:0;
  const contain=inter/Math.min(sa.size,sb.size);
  return Math.max(jaccard,contain*0.92);
}
const LYRIC_MATCH_THRESHOLD=0.62;

/** Merge tiny unnamed stanzas and label structure so we don't get 17 "Verse" slides. */
function consolidateLyricBlocks(blocks){
  if(!blocks.length)return blocks;
  // Drop pure filler blocks
  let list=blocks.map(b=>({
    name:b.name||'',
    lines:(b.lines||[]).filter(l=>!isFillerLyricLine(l)),
    piece:b.piece||0
  })).filter(b=>b.lines.length);

  // Merge consecutive 1-line unnamed stanzas into the previous block (or pool them)
  const merged=[];
  for(const b of list){
    const prev=merged[merged.length-1];
    if(prev&&!b.name&&b.lines.length<=1&&!prev.name){
      prev.lines.push(...b.lines);
      continue;
    }
    if(prev&&!b.name&&b.lines.length<=2&&prev.lines.length<6&&!b.name){
      if(!prev.name&&prev.lines.length<=4){prev.lines.push(...b.lines);continue}
    }
    merged.push({name:b.name,lines:[...b.lines],piece:b.piece});
  }
  list=merged;

  const named=list.filter(b=>b.name).length;
  if(named===0&&list.length){
    // Target ~4–8 lines per verse when no labels: regroup flat lines if too fragmented
    if(list.length>8||list.every(b=>b.lines.length<=2)){
      const all=list.flatMap(b=>b.lines);
      const regrouped=[];
      const target=4;
      for(let i=0;i<all.length;i+=target){
        regrouped.push({name:'',lines:all.slice(i,i+target),piece:0});
      }
      list=regrouped;
    }

    // Cluster similar blocks (fuzzy) so slight vocalist variations still match
    const clusterOf=new Array(list.length).fill(-1);
    let nextCluster=0;
    for(let i=0;i<list.length;i++){
      if(clusterOf[i]>=0)continue;
      clusterOf[i]=nextCluster;
      for(let j=i+1;j<list.length;j++){
        if(clusterOf[j]>=0)continue;
        if(lyricSimilarity(list[i].lines,list[j].lines)>=LYRIC_MATCH_THRESHOLD){
          clusterOf[j]=nextCluster;
        }
      }
      nextCluster++;
    }
    const clusterSize=c=>clusterOf.filter(x=>x===c).length;
    const clusterLines=c=>{
      // Prefer the longest member as representative length
      let best=0;
      list.forEach((b,i)=>{if(clusterOf[i]===c)best=Math.max(best,b.lines.length)});
      return best;
    };

    // Chorus = largest cluster with enough content (prefer longer sections)
    let chorusCluster=-1,chorusScore=-1;
    for(let c=0;c<nextCluster;c++){
      const n=clusterSize(c);
      if(n<2)continue;
      const len=clusterLines(c);
      const score=n*10+len; // prefer frequent + reasonably long
      if(score>chorusScore){chorusScore=score;chorusCluster=c}
    }
    // If the biggest cluster is tiny (1–2 lines) and another long cluster also repeats, prefer the long one for Chorus
    if(chorusCluster>=0&&clusterLines(chorusCluster)<=2){
      let alt=-1,altScore=-1;
      for(let c=0;c<nextCluster;c++){
        if(clusterSize(c)<2||clusterLines(c)<=2)continue;
        const score=clusterSize(c)*10+clusterLines(c);
        if(score>altScore){altScore=score;alt=c}
      }
      if(alt>=0)chorusCluster=alt;
    }

    list.forEach((b,i)=>{
      if(chorusCluster>=0&&clusterOf[i]===chorusCluster)b.name='Chorus';
    });

    // Tags: short clusters that repeat (or appear more than once), not the chorus — anywhere in the song
    const tagClusters=new Set();
    for(let c=0;c<nextCluster;c++){
      if(c===chorusCluster)continue;
      const n=clusterSize(c);
      const len=clusterLines(c);
      if(n>=2&&len<=3)tagClusters.add(c);
      // Single short section that is very similar to another short one already tagged — handled by cluster
    }
    // Also: single short sections that fuzzy-match a known tag cluster member
    list.forEach((b,i)=>{
      if(b.name)return;
      if(tagClusters.has(clusterOf[i]))b.name='Tag';
    });

    // Bridge: unique (not in a multi-member cluster), between choruses, after ≥2 other non-chorus
    const nonChorusIdx=[];
    list.forEach((b,i)=>{if(b.name!=='Chorus')nonChorusIdx.push(i)});
    let bridgeIdx=-1;
    if(chorusCluster>=0){
      for(const i of nonChorusIdx){
        if(list[i].name)continue; // already Tag
        const alone=clusterSize(clusterOf[i])===1;
        const chorusBefore=list.slice(0,i).some(b=>b.name==='Chorus');
        const chorusAfter=list.slice(i+1).some(b=>b.name==='Chorus');
        const nonChorusBefore=nonChorusIdx.filter(j=>j<i).length;
        if(
          alone&&chorusBefore&&chorusAfter&&
          list[i].lines.length>=3&&
          nonChorusBefore>=2&&
          bridgeIdx<0
        ){
          bridgeIdx=i;
          break;
        }
      }
    }

    let v=0;
    list.forEach((b,i)=>{
      if(b.name==='Chorus'||b.name==='Tag')return;
      if(i===bridgeIdx){b.name='Bridge';return}
      // Trailing short unique leftover → Outro (not Tag unless already matched)
      if(i===list.length-1&&b.lines.length<=2&&v>=1&&clusterSize(clusterOf[i])===1){
        b.name='Outro';
        return;
      }
      b.name='Verse '+(++v);
    });
    // First very short unique → Intro
    if(list[0]&&list[0].name&&list[0].name.startsWith('Verse')&&list[0].lines.length<=2&&list.length>2){
      list[0].name='Intro';
      let vv=0;
      list.forEach(b=>{if(/^Verse/.test(b.name))b.name='Verse '+(++vv)});
    }
  }else{
    // Has some names: still merge tiny unnamed gaps into previous
    const out=[];
    for(const b of list){
      const prev=out[out.length-1];
      if(prev&&!b.name&&b.lines.length<=2){prev.lines.push(...b.lines);continue}
      out.push(b);
    }
    list=out;
    // Unnamed leftovers → Verse n / Tag / Outro
    let v=list.filter(b=>/^Verse/i.test(b.name)).length;
    list.forEach((b,i)=>{
      if(b.name)return;
      if(i===list.length-1&&b.lines.length<=3){b.name=b.lines.length<=2?'Tag':'Outro';return}
      b.name='Verse '+(++v);
    });
  }
  return list;
}

function parse(x, linesPerSlide = 0){
  const P=[], O=[], seen={};let vn=0;
  const lps=parseInt(linesPerSlide,10)||0;
  const normalized=normalizeLyricStructure(x);
  const HDR=/^(verse|chorus|bridge|pre-chorus|prechorus|tag|intro|outro|ending|refrain|hook)\b\s*(\d*)\s*(?:\((\d+)\/(\d+)\))?\s*:?\s*$/i;
  const blocks=[];
  const rawLines=normalized.split('\n');
  let cur={name:'',lines:[],piece:0};
  const flush=()=>{
    const lines=cur.lines.map(l=>l.trim()).filter(l=>l&&!isFillerLyricLine(l));
    if(!lines.length){cur={name:'',lines:[],piece:0};return}
    const prev=blocks[blocks.length-1];
    if(cur.piece>1&&prev&&prev.name===cur.name&&prev.piece>=1){
      prev.lines.push(...lines);prev.piece=cur.piece;
    }else blocks.push({name:cur.name,lines,piece:cur.piece});
    cur={name:'',lines:[],piece:0};
  };
  for(const line of rawLines){
    const trimmed=line.trim();
    if(!trimmed){flush();continue}
    const m=trimmed.match(HDR);
    if(m){
      flush();
      let label=m[1].toLowerCase().replace(/\s+/g,'');
      if(label==='prechorus')label='pre-chorus';
      if(label==='ending')label='outro';
      if(label==='refrain'||label==='hook')label='chorus';
      const nice=label==='pre-chorus'?'Pre-Chorus':(label[0].toUpperCase()+label.slice(1));
      cur={name:nice+(m[2]?' '+m[2]:''),lines:[],piece:m[3]?+m[3]:0};
      continue;
    }
    if(!isFillerLyricLine(trimmed))cur.lines.push(trimmed);
  }
  flush();

  const consolidated=consolidateLyricBlocks(blocks);

  consolidated.forEach(b=>{
    const base=b.name||('Verse '+(++vn));
    const chunks=[];
    if(lps>0&&b.lines.length>lps)for(let i=0;i<b.lines.length;i+=lps)chunks.push(b.lines.slice(i,i+lps).join('\n'));
    else chunks.push(b.lines.join('\n'));
    chunks.forEach((tx,k)=>{
      const nm=chunks.length>1?`${base} (${k+1}/${chunks.length})`:base,key=nm+'|'+tx;
      if(!seen[key]){const p={id:uid(),n:nm,tx};P.push(p);seen[key]=p.id}
      O.push(seen[key]);
    });
  });
  return { parts: P, ord: O };
}

const LIB=[];
const emptyContact=()=>({name:'',email:''});
const defaultTeam=()=>({leader:emptyContact(),band:[emptyContact(),emptyContact(),emptyContact()],sound:emptyContact()});
function ensureTeam(service){const t=service.team||(service.team=defaultTeam());t.leader={...emptyContact(),...(t.leader||{})};t.band=Array.isArray(t.band)?t.band.map(x=>({...emptyContact(),...x})):[emptyContact(),emptyContact(),emptyContact()];t.sound={...emptyContact(),...(t.sound||{})};return t}

function defaultTextColor(bg,customBg){
  const b=BG[bg];
  if(b&&b.c==='#000')return '#ffffff';
  const value=String((b&&b.pick?customBg:'')||'').trim();
  const m=/^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(value);
  if(m){let h=m[1];if(h.length===3)h=h.split('').map(c=>c+c).join('');if(parseInt(h,16)>=0xf0f0f0)return '#000000';if(parseInt(h,16)<=0x101010)return '#ffffff'}
  return '#ffffff';
}
const mk=(t,k,parts,ord,bg,fs)=>({t,k,parts,ord,bg,dim:30,f:{fs:fs||5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:defaultTextColor(bg)}});
const mkS=(t,arr,bg,fs)=>{const ps=arr.map((tx,i)=>({id:uid(),n:'Slide '+(i+1),tx}));return mk(t,'Slide',ps,ps.map(p=>p.id),bg,fs)};
const fromLib=(x,bg)=>{const c=cl(x);const it=mk(c.t,'Song',c.parts,c.ord,bg,3.5);if(c.lps!=null)it.lps=c.lps;return it};

let S=[];
let sel=0,si=0,clr=null,live=false,tab='Style',tt,dg=null,ed=false,A=null,aj=0,editingLibIndex=null;
let churchLogos=[],defaultLogoId='';
try{churchLogos=JSON.parse(localStorage.getItem('church-presenter-logos')||'[]')||[];defaultLogoId=localStorage.getItem('church-presenter-logo-default')||''}catch{churchLogos=[]}
let churchLogo=localStorage.getItem('church-presenter-logo')||'';
if(!churchLogos.length&&churchLogo){churchLogos=[{id:'logo-1',name:'Logo 1',src:churchLogo}];defaultLogoId='logo-1';try{localStorage.setItem('church-presenter-logos',JSON.stringify(churchLogos));localStorage.setItem('church-presenter-logo-default',defaultLogoId)}catch{}}
let churchLogoBg=(()=>{try{return localStorage.getItem('church-presenter-logo-bg')||'#0a1428'}catch{return'#0a1428'}})();
let churchLogoSize=(()=>{try{return Math.max(10,Math.min(100,parseInt(localStorage.getItem('church-presenter-logo-size')||'80',10)||80))}catch{return 80}})();
let churchLogoVPos=(()=>{try{return localStorage.getItem('church-presenter-logo-vpos')||'center'}catch{return'center'}})();
let churchLogoHPos=(()=>{try{return localStorage.getItem('church-presenter-logo-hpos')||'center'}catch{return'center'}})();
function logoBgCss(){return churchLogoBg||'#0a1428'}
function logoAlignCss(){return `align-items:${churchLogoVPos};justify-content:${churchLogoHPos}`}
function logoImageCss(){return `max-width:${churchLogoSize}%;max-height:${churchLogoSize}%;object-fit:contain`}
function logoPlaceholderCss(){return `font-size:${Math.max(1,churchLogoSize/50)}cqw;opacity:.85;color:#fff;text-align:center`}
function logoMarkup(){return churchLogo?`<img src="${e(churchLogo)}" alt="Church Logo" style="${logoImageCss()}">`:`<div style="${logoPlaceholderCss()}"><b>Church Logo</b><br><small style="font-size:.6em">Upload logo in Settings</small></div>`}
let services=[],activeService=0,serviceTagFilter='',serviceSort='edited',displayMonitors=[],imageSel=-1,imageDrag=null,selectedDisplay='default',elSel=null,canvasClip=null,slideClip=null,elDrag=null,stripFocus='slide',pptHandoffActive=false,pptHandoffKey='';

/* Consolidated keys & state for multi-select, templates, and members */
const TPL_KEY='service-slides-templates';
const TOMB_KEY='service-slides-template-deleted';
const TPL_DEFAULT_KEY='service-slides-default-template';
const MEMBERS_KEY='service-slides-church-members';

const MS=window.__CP_MS=window.__CP_MS||{slides:new Set(),els:new Set(),slideAnchor:null,elAnchor:null};
function sk(i,j){return String(i)+':'+String(j)}
function pk(k){const p=String(k).split(':');return{i:+p[0],j:+p[1]}}
function paintSlides(){
  document.querySelectorAll('#strip .th.ms-on').forEach(n=>n.classList.remove('ms-on'));
  MS.slides.forEach(k=>{
    const p=pk(k);
    const th=document.querySelector('#strip .th[data-i="'+p.i+'"][data-j="'+p.j+'"]');
    if(th)th.classList.add('ms-on');
  });
}
function paintEls(){
  document.querySelectorAll('#stage .el.ms-on,.mid .el.ms-on').forEach(n=>n.classList.remove('ms-on'));
  MS.els.forEach(id=>{
    const n=document.querySelector('#stage [data-el="'+id+'"],.mid [data-el="'+id+'"]');
    if(n)n.classList.add('ms-on');
  });
}
function multiParts(){
  const out=[];
  MS.slides.forEach(k=>{
    const p=pk(k);
    try{
      const it=S[p.i];if(!it||isSongItem(it))return;
      const part=it.parts.find(x=>x.id===it.ord[p.j]);
      if(part)out.push({i:p.i,j:p.j,part});
    }catch{}
  });
  return out.sort((a,b)=>a.i-b.i||a.j-b.j);
}

// Duplicate UI utility functions removed – original definitions earlier in the file are kept
function localISODate(d){const m=d.getMonth()+1,dd=d.getDate();return d.getFullYear()+'-'+(m<10?'0':'')+m+'-'+(dd<10?'0':'')+dd}
function upcomingSundayISO(){const d=new Date();d.setHours(12,0,0,0);d.setDate(d.getDate()+(7-d.getDay())%7);return localISODate(d)}
function relDayLabel(d){const t=new Date(),k=localISODate(d);if(k===localISODate(t))return 'Today';if(k===localISODate(new Date(t.getFullYear(),t.getMonth(),t.getDate()-1)))return 'Yesterday';return ''}
function relDayLabelISO(iso){const m=/^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso||''));return m?relDayLabel(new Date(+m[1],+m[2]-1,+m[3],12)):''}

function defaultTplId(){
  try{
    const id=localStorage.getItem(TPL_DEFAULT_KEY)||'';
    if(!id)return '';
    const list=JSON.parse(localStorage.getItem(TPL_KEY)||'[]');
    return list.some(t=>t.id===id)?id:'';
  }catch{return ''}
}
function isDefaultTpl(id){return !!id&&defaultTplId()===id}
function loadTemplates(){try{const a=JSON.parse(localStorage.getItem(TPL_KEY)||'[]');return Array.isArray(a)?a:[]}catch{return []}}
function saveTemplates(list){try{localStorage.setItem(TPL_KEY,JSON.stringify(list));return true}catch{toast('Could not save the template: storage is full.');return false}}
function loadTombs(){try{const a=JSON.parse(localStorage.getItem(TOMB_KEY)||'[]');return Array.isArray(a)?a:[]}catch{return []}}
function saveTombs(a){try{localStorage.setItem(TOMB_KEY,JSON.stringify(a))}catch{}}
function newTplId(){return 'tpl-'+Date.now().toString(36)+Math.random().toString(36).slice(2,6)}

try{
  const saved=JSON.parse(localStorage.getItem('service-slides-state')||'null');
  if(Array.isArray(saved?.services)){
    services=saved.services;
    if(!services.length){
      activeService=0;S=[];
    }else{
      activeService=Math.min(Math.max(0,saved.activeService||0),services.length-1);
      S=cl(services[activeService].slides||[]);
    }
    if(Array.isArray(saved.library)){
      const removedDefaultSongs=new Set(['goodnessofgod','amazinggrace','holyholyholy']);
      const keptLibrary=saved.library.filter(song=>{
        const title=String(song&&song.t||song&&song.title||'').toLowerCase().replace(/[^a-z0-9]/g,'');
        const id=String(song&&song.id||'').toLowerCase();
        return ![...removedDefaultSongs].some(removed=>title.startsWith(removed))&&!['song-goodness-of-god','song-amazing-grace','song-holy-holy-holy'].includes(id);
      });
      LIB.splice(0,LIB.length,...keptLibrary);
      if(keptLibrary.length!==saved.library.length){
        saved.library=keptLibrary;
        try{localStorage.setItem('service-slides-state',JSON.stringify(saved))}catch{}
      }
    }
    document.documentElement.dataset.theme=saved.theme==='system'?'':saved.theme;
    $('#theme-setting').value=saved.theme||'system';
    selectedDisplay=saved.display||'default';
    $('#display-setting').value=selectedDisplay;
  }
}catch{}
services.forEach(s=>{
  ensureTeam(s);
  if(!s.date)s.date=new Date().toISOString().slice(0,10);
  if(s.tag==null)s.tag='';
});

// Song search always uses the free online lyric database (no AI providers).
const lyricsSettings={provider:'free',localEndpoint:'',localModel:'',hostedEndpoint:'',hostedModel:'',hostedApiKey:''};
try{configureSongSearch(lyricsSettings)}catch{}

let songResults=[],searchController=null,lyricsController=null,searchDebounce=null,addingSongId=null;
let songSettings={blankSlide:'none',credits:false,creditsPos:'middle',creditFields:{title:true,artist:true,writers:true,year:true,ccli:true}};
try{
  const ss=JSON.parse(localStorage.getItem('service-slides-song-settings')||'null');
  if(ss&&typeof ss==='object')songSettings={...songSettings,...ss};
}catch{}
function saveSongSettings(){
  try{localStorage.setItem('service-slides-song-settings',JSON.stringify(songSettings))}catch{}
}
function closeSettingsAndRefresh(){
  const panel=$('#settings');if(!panel?.classList.contains('open'))return;
  panel.classList.remove('open');
  try{rAll()}catch(err){console.warn('[song credits] refresh after settings',err)}
}
let appDialogActive=null;
function appDialog({title='Church Presenter',message='',value,confirmLabel='Continue',danger=false}={}){
  if(appDialogActive)return Promise.resolve(null);
  return new Promise(resolve=>{
    const promptMode=value!==undefined,layer=document.createElement('div');
    layer.className='app-modal';layer.style.cssText='position:fixed;inset:0;z-index:2147483646;background:#0009;display:grid;place-items:center;padding:18px';
    layer.innerHTML=`<div class="app-dialog" role="alertdialog" aria-modal="true" style="width:min(440px,100%);padding:22px"><h2 style="margin:0 0 12px;font-size:calc(18px * var(--ui-scale))"></h2><p data-message style="white-space:pre-wrap;margin:0 0 18px;color:var(--mu)"></p>${promptMode?'<input data-dialog-input type="text" style="width:100%;margin:0 0 18px">':''}<div style="display:flex;justify-content:flex-end;gap:8px"><button type="button" data-dialog-cancel>Cancel</button><button type="button" data-dialog-confirm class="${danger?'danger-action':'pri'}"></button></div></div>`;
    layer.querySelector('h2').textContent=title;layer.querySelector('[data-message]').textContent=message;layer.querySelector('[data-dialog-confirm]').textContent=confirmLabel;
    const input=layer.querySelector('[data-dialog-input]');if(input)input.value=String(value??'');
    let settled=false;
    const finish=result=>{if(settled)return;settled=true;window.removeEventListener('keydown',onKey,true);layer.remove();appDialogActive=null;resolve(result)};
    const onKey=ev=>{if(ev.key==='Escape'){ev.preventDefault();ev.stopImmediatePropagation();finish(null)}else if(ev.key==='Enter'||ev.key===' '){ev.preventDefault();ev.stopImmediatePropagation();finish(promptMode?input.value:true)}};
    layer.addEventListener('click',ev=>{if(ev.target===layer||ev.target.closest('[data-dialog-cancel]'))finish(null);else if(ev.target.closest('[data-dialog-confirm]'))finish(promptMode?input.value:true)});
    window.addEventListener('keydown',onKey,true);appDialogActive=layer;document.body.append(layer);(input||layer.querySelector('[data-dialog-confirm]')).focus();input?.select();
  });
}
const appConfirm=(message,opts={})=>appDialog({...opts,message});
const appPrompt=(title,message,value='')=>appDialog({title,message,value,confirmLabel:'Done'});
window.churchPresenterConfirm=appConfirm;
const sls=it=>it.ord.map(id=>it.parts.find(p=>p.id==id));
const mediaURLs={};
/* ---- import progress + PPT state (these were missing) ---- */
let importInProgress=false,importCancelled=false,pptCompletedKeys=new Set(),pptOpenedThisPresent=new Set();
function setImportProgress(pct,msg){
  const p=Math.max(0,Math.min(100,Math.round(pct)));
  const bar=$('#import-progress-bar'),lab=$('#import-progress-label'),txt=$('#import-status-text');
  if(bar)bar.style.width=p+'%';
  if(lab)lab.textContent=p+'%';
  if(msg&&txt)txt.textContent=msg;
}
function showImportModal(){
  importCancelled=false;importInProgress=true;
  setImportProgress(0,'Preparing…');
  const m=$('#import-modal');if(m){m.style.zIndex='2147483647';m.hidden=false;}
}
function hideImportModal(){
  importInProgress=false;
  const m=$('#import-modal');if(m)m.hidden=true;
}
function throwIfImportCancelled(){
  if(importCancelled){const err=new Error('Import cancelled');err.cancelled=true;throw err;}
}



const isSongItem=it=>it&&it.k==='Song';
function ensureElements(part,it){
  if(!part||isSongItem(it))return null;
  if(!Array.isArray(part.elements)){
    part.elements=[];
    if((part.tx||'').trim()){
      const f={...(it.f||{fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'})};
      part.elements.push({id:uid(),kind:'text',tx:part.tx,x:8,y:18,w:84,h:50,rot:0,z:1,f});
    }
    // migrate free-position images (non-full)
    (part.im||[]).forEach((m,i)=>{
      if(m.full)return;
      part.elements.push({
        id:uid(),kind:'image',
        src:m.src||mediaURLs[m.mediaId]||'',mediaId:m.mediaId||'',
        x:m.x??50,y:m.y??50,w:m.w??30,h:m.w??30,rot:0,z:2+i
      });
    });
  }
  return part.elements;
}
function getSelectedEl(){
  const part=sls(A)[aj];if(!part||!elSel)return null;
  return (part.elements||[]).find(el=>el.id===elSel)||null;
}
/** Text for presenter list mode: freeform text boxes (foreground→background by z), else lyrics/tx. */
function slideDisplayText(it,p){
  if(!p)return '(empty)';
  if(p.logo)return '(Church logo)';
  if(isSongItem(it)){
    const t=(p.tx||'').trim();
    return t||'(empty)';
  }
  if(typeof ensureElements==='function')ensureElements(p,it);
  const els=[...(p.elements||[])].filter(el=>el&&el.kind==='text'&&String(el.tx||'').trim());
  // Foreground first (higher z), background last (lower z)
  els.sort((a,b)=>(b.z||0)-(a.z||0));
  if(els.length)return els.map(el=>String(el.tx||'').trim()).join('\n\n');
  const t=(p.tx||'').trim();
  return t||'(empty)';
}
function styleForTextEl(el,f){
  const ff=f.ff==='serif'?"'Lora',Georgia,'Times New Roman',serif":"'Figtree','Segoe UI',system-ui,sans-serif";
  const tt=f.tc==='upper'?'uppercase':f.tc==='title'?'capitalize':'none';
  const ta=f.al||'center';
  const jc=f.al==='left'?'flex-start':f.al==='right'?'flex-end':'center';
  const ai=VA[f.va]||'center';
  return `font-size:${f.fs||5.2}cqw;font-family:${ff};text-transform:${tt};text-align:${ta};justify-content:${jc};align-items:${ai};font-weight:${f.b?700:400};color:${f.color||'#ffffff'};${f.sh?'text-shadow:0 .1em .35em rgba(0,0,0,0.85)':''}`;
}
function renderEl(el,editable,selected,stage){
  if(el.kind==='video')return renderVideoEl(el,editable,selected,stage);
  const rot=el.rot||0;
  const base=`left:${el.x}%;top:${el.y}%;width:${el.w}%;height:${el.h}%;transform:rotate(${rot}deg);z-index:${el.z||1}`;
  const sel=selected?' selected':'';
  const handles=editable&&selected?elHandlesHtml(el):'';
  if(el.kind==='text'){
    const f=el.f||{};
    return `<div class="el text-el${sel}" data-el="${e(el.id)}" style="${base};${styleForTextEl(el,f)};pointer-events:${editable?'auto':'none'};cursor:${editable?'move':'default'}">${handles}<span class="el-tx" style="pointer-events:none">${e(el.tx||'')}</span></div>`;
  }
  const src=el.src||mediaURLs[el.mediaId]||'';
  const fit=el.fit==='cover'?'cover':(el.fit==='fill'?'fill':'contain');
  return `<div class="el img-el${sel}" data-el="${e(el.id)}" style="${base};pointer-events:${editable?'auto':'none'}">${handles}<img src="${e(src)}" alt="" draggable="false" style="width:100%;height:100%;object-fit:${fit};pointer-events:none;display:block"></div>`;
}

// Shared click-drag animation: the picked item shrinks into a small pill (or a mini slide picture),
// follows the pointer, auto-scrolls near the edges of o.scrollEl, and drops via o.onDrop.
// o: {el, label, sub, thumbOf(), scrollEl (element or fn), scrollPadX, onMove(x,y), onDrop(x,y), onTap(ev), onEnd()}
function pillDrag(ev,o){
  if(!ev||ev.button!==0||!o||!o.el)return;
  const pid=ev.pointerId;
  const thumbEl=o.thumbOf?o.thumbOf():null;
  const srcRect=(thumbEl||o.el).getBoundingClientRect();
  const st={x:ev.clientX,y:ev.clientY,sx:ev.clientX,sy:ev.clientY,moved:false,ghost:null,t:0,raf:0,sraf:0,
    gx:ev.clientX-srcRect.left,gy:ev.clientY-srcRect.top,sw:srcRect.width||1,sh:srcRect.height||1,tw:0,th:0,r0:8,r1:16};
  function frame(){
    const g=st.ghost;if(!g)return;
    const k=st.t*st.t*(3-2*st.t);
    const w=st.sw+(st.tw-st.sw)*k,h=st.sh+(st.th-st.sh)*k;
    const ox=st.gx+(st.tw/2-st.gx)*k,oy=st.gy+(st.th/2-st.gy)*k;
    g.style.width=w+'px';g.style.height=h+'px';
    g.style.borderRadius=(st.r0+(st.r1-st.r0)*k)+'px';
    g.style.transform='translate('+(st.x-ox)+'px,'+(st.y-oy)+'px) rotate('+(2*k)+'deg)';
    const sm=g.querySelector('small');
    if(sm){sm.style.opacity=String(1-k);sm.style.display=k>=1?'none':'block'}
  }
  function scrollTick(){
    st.sraf=0;
    if(!st.ghost)return;
    const sc=typeof o.scrollEl==='function'?o.scrollEl():o.scrollEl;
    if(sc){
      const r=sc.getBoundingClientRect();
      const zone=Math.max(30,Math.min(80,r.height/3));
      const padX=o.scrollPadX!=null?o.scrollPadX:140;
      let v=0;
      if(st.x>=r.left-padX&&st.x<=r.right+padX){
        if(st.y<r.top+zone)v=-Math.min(1.8,(r.top+zone-st.y)/zone);
        else if(st.y>r.bottom-zone)v=Math.min(1.8,(st.y-(r.bottom-zone))/zone);
      }
      if(v){
        const before=sc.scrollTop;
        sc.scrollTop+=Math.sign(v)*Math.max(2,Math.round(Math.abs(v)*18));
        if(sc.scrollTop!==before&&o.onMove)o.onMove(st.x,st.y);
      }
    }
    st.sraf=requestAnimationFrame(scrollTick);
  }
  function begin(){
    st.moved=true;
    const g=document.createElement('div');
    g.className='pk-ghost';
    if(thumbEl){
      g.classList.add('pk-thumb');
      const c=thumbEl.cloneNode(true);
      c.querySelectorAll('video').forEach(v=>v.remove());
      c.style.cssText='width:100%;height:100%;margin:0;pointer-events:none';
      g.appendChild(c);
      st.tw=132;st.th=Math.round(132*st.sh/st.sw);st.r0=4;st.r1=8;
    }else{
      g.innerHTML='<b></b><small></small>';
      g.children[0].textContent=o.label||'';
      g.children[1].textContent=o.sub||'';
      const m=document.createElement('span');
      m.style.cssText='position:fixed;left:-9999px;top:0;font:700 13px Figtree,system-ui,sans-serif;white-space:nowrap';
      m.textContent=o.label||'';document.body.appendChild(m);
      st.tw=m.offsetWidth+32;m.remove();st.th=32;
    }
    document.body.appendChild(g);
    st.ghost=g;
    o.el.classList.add('pk-src');
    document.body.dataset.pointerDragging='1';
    try{getSelection().removeAllRanges()}catch{}
    const t0=performance.now();
    const tick=now=>{
      if(st.ghost!==g)return;
      st.t=Math.min(1,(now-t0)/200);frame();
      if(st.t<1)st.raf=requestAnimationFrame(tick);
    };
    st.raf=requestAnimationFrame(tick);
    st.sraf=requestAnimationFrame(scrollTick);
  }
  function move(e){
    if(e.pointerId!==pid)return;
    st.x=e.clientX;st.y=e.clientY;
    if(!st.moved){
      if(Math.hypot(st.x-st.sx,st.y-st.sy)<6)return;
      begin();
    }
    frame();
    if(o.onMove)o.onMove(st.x,st.y);
    e.preventDefault();
  }
  function finish(drop,e){
    document.removeEventListener('pointermove',move,true);
    document.removeEventListener('pointerup',up,true);
    document.removeEventListener('pointercancel',cancel,true);
    cancelAnimationFrame(st.raf);cancelAnimationFrame(st.sraf);
    const g=st.ghost;st.ghost=null;
    o.el.classList.remove('pk-src');
    document.body.dataset.pointerDragging='';
    if(g){g.style.transition='opacity .12s';g.style.opacity='0';setTimeout(()=>g.remove(),140)}
    if(st.moved){
      const sup=x=>{x.stopImmediatePropagation();x.preventDefault()};
      document.addEventListener('click',sup,true);
      setTimeout(()=>document.removeEventListener('click',sup,true),80);
    }
    try{
      if(drop&&st.moved&&o.onDrop)o.onDrop(st.x,st.y);
      else if(drop&&!st.moved&&o.onTap)o.onTap(e);
    }finally{
      if(st.moved&&o.onEnd)o.onEnd();
    }
  }
  function up(e){if(e.pointerId===pid)finish(true,e)}
  function cancel(e){if(e.pointerId===pid)finish(false,e)}
  document.addEventListener('pointermove',move,true);
  document.addEventListener('pointerup',up,true);
  document.addEventListener('pointercancel',cancel,true);
}
document.addEventListener('pointerup',()=>document.querySelectorAll('#stage .ratio-snapped').forEach(n=>n.classList.remove('ratio-snapped')));
document.addEventListener('pointercancel',()=>document.querySelectorAll('#stage .ratio-snapped').forEach(n=>n.classList.remove('ratio-snapped')));
// Paste handling is consolidated in the unified clipboard handler

function elHandlesHtml(el){
  let h='<span class="handle nw" data-h="nw"></span><span class="handle ne" data-h="ne"></span><span class="handle sw" data-h="sw"></span><span class="handle se" data-h="se"></span><span class="handle rot" data-h="rot"></span>';
  if(el&&el.kind!=='text'){
    const locked=!!el.lock;
    h+='<span class="lock-btn'+(locked?' locked':'')+'" title="'+(locked?'Aspect ratio LOCKED (click to unlock)':'Aspect ratio FREE - stretches (click to lock)')+'"><svg viewBox="0 0 24 24"><g class="lk-sh"><path d="M7 11V7a5 5 0 0 1 10 0v4"/></g><rect x="3" y="11" width="18" height="11" rx="2"/></svg></span>';
  }
  return h;
}
function mediaNatRatio(node){
  const img=node.querySelector('img'),vid=node.querySelector('video');
  if(img&&img.naturalWidth&&img.naturalHeight)return img.naturalWidth/img.naturalHeight;
  if(vid&&vid.videoWidth&&vid.videoHeight)return vid.videoWidth/vid.videoHeight;
  return 0;
}
// Make the box hug the visible picture, then switch to fill so resizing really stretches.
// ------------------------------------------------------------
// Helper: add an image blob to the current slide (or a new slide)
// ------------------------------------------------------------
function addImageFromBlob(blob) {
  const url = URL.createObjectURL(blob);
  const targetSlide = (A && aj >= 0) ? A : mkS('Untitled', [], 0, 6.5);
  if (!targetSlide.im) targetSlide.im = [];
  // Add required fields for the renderer
  targetSlide.im.push({
    src: url,
    full: true,
    kind: 'image',
    x: 0,
    y: 0,
    w: 30, // percent width
    h: 30, // percent height
    fit: 'fill',
  });
  rStrip();
  // Primary UI refresh; fall back to drawStage if renderStage is not defined
  if (typeof renderStage === 'function') renderStage();
  else if (typeof drawStage === 'function') drawStage();
  // Revoke the object URL after a short delay to avoid memory leaks
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

function normalizeMediaBox(el,node,rectW,rectH){
  const nr=mediaNatRatio(node);
  if(nr&&!el.nr)el.nr=nr;
  if(el.fit==='fill'||el.fit==='cover')return;
  if(nr){
    const bw=el.w/100*rectW,bh=el.h/100*rectH;
    let dw=bw,dh=bh;
    if(bw/bh>nr)dw=bh*nr;else dh=bw/nr;
    const cx=el.x+el.w/2,cy=el.y+el.h/2;
    el.w=dw/rectW*100;el.h=dh/rectH*100;el.x=cx-el.w/2;el.y=cy-el.h/2;
    node.style.left=el.x+'%';node.style.top=el.y+'%';node.style.width=el.w+'%';node.style.height=el.h+'%';
  }
  el.fit='fill';
  node.querySelectorAll('img,video').forEach(m=>{m.style.objectFit='fill'});
}
// Stretch by default. Shift or the padlock keeps the ratio. Unlocked: gentle snap to natural ratio (Alt disables).
function resizeMediaEl(el,mode,orig,dx,dy,ev,part,rectW,rectH){
  const keep=ev.shiftKey||!!el.lock;
  const clamp=(v,a,b)=>Math.max(a,Math.min(b,v));
  const nd=document.querySelector('#stage [data-el="'+el.id+'"]');
  let w,h;
  if(mode==='se'){w=orig.w+dx;h=orig.h+dy}
  else if(mode==='nw'){w=orig.w-dx;h=orig.h-dy}
  else if(mode==='ne'){w=orig.w+dx;h=orig.h-dy}
  else{w=orig.w-dx;h=orig.h+dy}
  let ratioSnapped=false;
  if(keep){
    const sw=w/orig.w,sh=h/orig.h;
    let s=Math.abs(sw-1)>=Math.abs(sh-1)?sw:sh;
    s=clamp(s,Math.max(4/orig.w,4/orig.h),Math.min(100/orig.w,100/orig.h));
    w=orig.w*s;h=orig.h*s;
  }else{
    w=clamp(w,4,100);h=clamp(h,4,100);
    if(el.nr&&!ev.altKey){
      const rp=el.nr*rectH/rectW;
      const was=!!(nd&&nd.classList.contains('ratio-snapped'));
      // sticky zone: snaps in within ~9% of the original ratio, lets go beyond ~14%
      if(Math.abs((w/h)/rp-1)<(was?0.14:0.09)){
        const cw=Math.abs(w/orig.w-1),ch=Math.abs(h/orig.h-1);
        if(cw>=ch)h=w/rp;else w=h*rp;
        w=clamp(w,4,100);h=clamp(h,4,100);ratioSnapped=true;
      }
    }
  }
  if(mode==='se'){el.x=orig.x;el.y=orig.y}
  else if(mode==='nw'){el.x=orig.x+orig.w-w;el.y=orig.y+orig.h-h}
  else if(mode==='ne'){el.x=orig.x;el.y=orig.y+orig.h-h}
  else{el.x=orig.x+orig.w-w;el.y=orig.y}
  el.w=w;el.h=h;
  if(nd)nd.classList.toggle('ratio-snapped',ratioSnapped);
  const g=(!keep&&!ratioSnapped&&!ev.altKey)?applyResizeSnap(el,mode,part):{guideXs:[],guideYs:[]};
  g.ratioSnapped=ratioSnapped;
  return g;
}

function bgStyleFor(it,p){
  const own=!!(p&&p.ownBg);
  const b=BG[own&&p.bg!=null?p.bg:it.bg]||BG[0];
  if(b&&b.pick)return (own&&p.customBg)||it.customBg||'#16213e';
  if(b&&b.c)return b.c;
  return '#000';
}
function sl(it,p,x,stage=false,slideIndex=-1,audience=false){
  
  const own=!!p.ownBg;
  const b=BG[own&&p.bg!=null?p.bg:it.bg]||BG[0],f=it.f||{fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'};
  const media=own?p.bgMedia:(p.bgMedia||it.bgMedia),video=media?.kind==='video',src=(media?.id&&mediaURLs[media.id])||media?.path||media?.source||media?.src||'';
  const vkey=video?(media.id||src):'';
  const background=media&&src?(video?(stage?`<div class="vslot" data-vkey="${e(vkey)}" data-vsrc="${e(src)}"></div>`:`<video src="${e(src)}" muted playsinline preload="metadata"></video>`):`<img src="${e(src)}" alt="">`):'';

  const fontFamily=f.ff==='serif'?"'Lora',Georgia,'Times New Roman',serif":"'Figtree','Segoe UI',system-ui,sans-serif";
  const textTransform=f.tc==='upper'?'uppercase':f.tc==='title'?'capitalize':'none';
  const justifyContent=f.al==='left'?'flex-start':f.al==='right'?'flex-end':'center';
  const textAlign=f.al||'center';
  const alignItems=VA[f.va]||'center';

  // Always scale text from slide width (cqw) so editor, previews, and output match.
  // .sl must have container-type:inline-size and a consistent 16:9 box.
  const fontSize=`${f.fs}cqw`;

  const els=ensureElements(p,it);
  const canEditEls=!!(els&&stage&&!audience&&!live&&!isSongItem(it));

  if((it&&it.k==='Logo')||(p&&p.logo)){
    return `<div class="sl" style="background:${logoBgCss()};display:flex;${logoAlignCss()};overflow:hidden">${logoMarkup()}</div>`;
  }
  let elHtml='';
  if(els&&els.length){
    const sorted=[...els].sort((a,b)=>(a.z||0)-(b.z||0));
    elHtml=sorted.map(el=>renderEl(el,canEditEls,canEditEls&&el.id===elSel,stage)).join('');
  }else if(isSongItem(it)){
    // Credits only on the last lyric slide of the arrangement (by position, not shared part id)
    let cred='',cpos=songSettings.creditsPos||it.creditsPos||'middle';
    if(isSongItem(it)&&songSettings.credits&&slideIndex>=0){
      let last=-1;
      for(let i=0;i<(it.ord||[]).length;i++){
        const pp=it.parts.find(x=>x.id===it.ord[i]);
        if(pp&&typeof isAutoBlankPart==='function'&&isAutoBlankPart(pp))continue;
        if(pp&&String(pp.tx||'').trim())last=i;
      }
      if(slideIndex===last){
        const meta=it.credits||{};
        cred=formatSongCredits(meta,it.t)||String(it.creditLine||p.credits||'').trim();
        cpos=songSettings.creditsPos||it.creditsPos||p.creditsPos||'middle';
      }
    }
    const credAlign=cpos==='left'?'left':cpos==='right'?'right':'center';
    const credHtml=cred?`<div class="song-credits" style="position:absolute;left:4%;right:4%;bottom:3%;z-index:6;font-size:1.35cqw;line-height:1.25;opacity:${(100-(songSettings.creditsTransparency!=null?songSettings.creditsTransparency:22))/100};color:${f.color||'#ffffff'};text-align:${credAlign};font-family:${fontFamily};font-weight:400;pointer-events:none;text-shadow:0 .05em .2em rgba(0,0,0,0.75)">${e(cred)}</div>`:'';
    const lyricMargin=Math.max(0,Math.min(25,Number(f.mb??8)));
    elHtml=`<div class="tx" style="padding-left:${lyricMargin}%;padding-right:${lyricMargin}%;font-size:${fontSize};font-family:${fontFamily};text-transform:${textTransform};text-align:${textAlign};justify-content:${justifyContent};align-items:${alignItems};font-weight:${f.b?700:400};color:${f.color||'#ffffff'};pointer-events:${x?'auto':'none'};${f.sh?'text-shadow:0 .1em .35em rgba(0,0,0,0.85)':''}">
      <span style="font-family:inherit;text-transform:inherit;text-align:${textAlign};text-wrap:pretty;"${x?' id="ce"':''}>${e(p.tx)}</span>
    </div>${credHtml}`;
  }else{ elHtml=''; }
  // Full-bleed images still from p.im
  const fullImgs=(p.im||[]).filter(m=>m.full).map((m,n)=>`<img class="stage-image" alt="" src="${m.src||mediaURLs[m.mediaId]||''}" draggable="false" style="position:absolute;inset:0;width:100%;height:100%;object-fit:contain;pointer-events:none">`).join('');

    return `<div class="sl">
    <div class="bgl ${!media&&b.v?'wave':''}" style="${!media&&!b.v?'background:'+bgStyleFor(it,p):''}">${background}</div>
    <div class="dim" style="opacity:${(own&&p.dim!=null?p.dim:it.dim)/100}"></div>
    ${fullImgs}
    ${p.videoMedia&&mediaURLs[p.videoMedia.id]?(()=>{const vs=p.videoMedia.settings||{};const muted=vs.muted===true;const vol=vs.volume!=null?vs.volume:1;return `<video class="slide-video${stage?'':' thumbnail-video'}" src="${e(mediaURLs[p.videoMedia.id])}" ${stage&&vs.autoplay?'autoplay':''} ${stage&&vs.loop?'loop':''} ${muted?'muted':''} data-vol="${vol}" playsinline ${stage?'controls':''} preload="metadata" ></video>`})():''}
    ${elHtml}
  </div>`;
}

/** Refresh only the current slide thumbnail in the strip (content edits). */
function updateCurrentThumb(){
  const i=S.indexOf(A);
  if(i<0){rStrip();return}
  const j=aj;
  const wrap=document.querySelector(`#strip .thumb-wrap[data-i="${i}"][data-j="${j}"]`);
  if(!wrap){rStrip();return}
  const it=A;
  const p=sls(it)[j];
  if(!p)return;
  const on=it===A&&j===aj;
  const itemOn=i===sel;
  const btn=wrap.querySelector('.th');
  if(!btn){rStrip();return}
  btn.className=`th ${on?'on':''} ${itemOn&&j===si?'es':''}`;
  btn.setAttribute('aria-label',`${it.t}, ${p.n}`);
  btn.innerHTML=`${sl(it,p,false,false,j)}<span class="pn">${e(p.n)}</span>`;
}

/* ===== Background targets: whole section / single slide / multi-selection ===== */
let _bgUndoAt=0;
function bgMultiList(){
  const list=multiParts();
  if(list.length<2)return [];
  return list.some(x=>x.i===sel&&x.j===si)?list:[];
}
function makeOwnBg(it,part){
  if(part.ownBg)return;
  part.ownBg=true;part.bg=it.bg;part.dim=it.dim;
  if(it.customBg)part.customBg=it.customBg;
  if(!part.bgMedia&&it.bgMedia)part.bgMedia=cl(it.bgMedia);
}
function clearOwnBg(part){['ownBg','bg','customBg','dim','bgMedia'].forEach(k=>delete part[k])}
function recolorDefaultText(obj,from,to){
  if(!obj||from===to)return;
  if(obj.f&&obj.f.color===from)obj.f.color=to;
  (obj.elements||[]).forEach(el=>{if(el.f&&el.f.color===from)el.f.color=to});
}
function bgTargets(){
  const it=S[sel];if(!it)return null;
  if(isSongItem(it))return {section:it,parts:[]};
  const multi=bgMultiList();
  if(multi.length>1)return {parts:multi.map(x=>({it:S[x.i],part:x.part}))};
  if(it.bgAll!==true){
    const p=sls(it)[Math.min(si,it.ord.length-1)];
    if(p)return {parts:[{it,part:p}]};
  }
  return {section:it,parts:[]};
}
function bgApply(fn,{quiet=false,keepOwn=false}={}){
  const tg=bgTargets();if(!tg)return;
  if(quiet){if(Date.now()-_bgUndoAt>1500)pushUndo();_bgUndoAt=Date.now()}
  else pushUndo();
  if(tg.section){
    const from=defaultTextColor(tg.section.bg,tg.section.customBg);
    fn(tg.section);
    const to=defaultTextColor(tg.section.bg,tg.section.customBg);
    recolorDefaultText(tg.section,from,to);
    (tg.section.parts||[]).forEach(p=>{if(keepOwn){if(p.ownBg){const was=defaultTextColor(p.bg,p.customBg);fn(p);recolorDefaultText(p,was,defaultTextColor(p.bg,p.customBg))}}else{clearOwnBg(p);recolorDefaultText(p,from,to)}});
  }else tg.parts.forEach(({it,part})=>{makeOwnBg(it,part);const was=defaultTextColor(part.bg,part.customBg);fn(part);recolorDefaultText(part,was,defaultTextColor(part.bg,part.customBg))});
  if(quiet){rStrip();rStage();saveSoon()}else rAll();
  setTimeout(paintSlides,0);
}
function applyCustomBg(t){
  let v=(t.value||'').trim();
  if(v&&v[0]!=='#')v='#'+v;
  if(!/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(v))return;
  if(v.length===4)v='#'+v[1]+v[1]+v[2]+v[2]+v[3]+v[3];
  const pickIdx=BG.findIndex(b=>b.pick);
  bgApply(o=>{o.customBg=v;if(pickIdx>=0)o.bg=pickIdx;o.dim=0;delete o.bgMedia},{quiet:true});
  const dim=$('input[type="range"][data-k="dim"]');if(dim)dim.value='0';
  const hex=$('#custom-bg-hex'),col=$('#custom-bg-color');
  if(hex&&document.activeElement!==hex)hex.value=v;
  if(col&&document.activeElement!==col)col.value=v;
}

/** Only move the "on" / "es" highlight in the strip (navigation). */
function updateStripSelection(){
  const strip=$('#strip');if(!strip)return;
  try{placeStripAdd()}catch{}
  strip.querySelectorAll('.th').forEach(btn=>{
    const i=+btn.dataset.i,j=+btn.dataset.j;
    const on=S[i]===A&&j===aj;
    const itemOn=i===sel;
    btn.classList.toggle('on',on);
    btn.classList.toggle('es',itemOn&&j===si);
  });
  strip.querySelectorAll('.hb').forEach(hb=>{
    const i=+hb.dataset.dh;
    hb.classList.toggle('on',i===sel);
    const hd=hb.querySelector('.hd');
    if(hd)hd.classList.toggle('on',i===sel);
  });
}

function rStrip(){
  const stripEl=$('#strip');
  const prevScrollY=stripEl?stripEl.scrollTop:0;
  const prevScrollX=stripEl?stripEl.scrollLeft:0;
  if(!S.length){
    $('#strip').innerHTML=`<div style="padding:14px 10px">
      <p class="note" style="margin:0 0 12px">No sections yet.</p>
      <button type="button" id="strip-add-slide" class="pri" style="width:100%">Add Slide</button>
    </div>`;
    if(stripEl){stripEl.scrollTop=prevScrollY;stripEl.scrollLeft=prevScrollX;}
    return;
  }
  $('#strip').innerHTML=S.map((it,i)=>{
    const isFocused=stripFocus!=='none';
    const on=i==sel&&isFocused,ss=sls(it);
    return `<div class="hb ${on?'on':''}" draggable="true" data-dh="${i}">
      <button class="hd ${on?'on':''}" data-hs="${i}">
        <b>${e(it.t)}</b>
        <small>${it.k}, ${ss.length} slide${ss.length>1?'s':''}</small>
      </button>
      
    </div>`+ss.map((p,j)=>`<div class="thumb-wrap" data-i="${i}" data-j="${j}">
      <button class="th ${it===A&&j==aj&&isFocused?'on':''} ${on&&j==si?'es':''}" data-i="${i}" data-j="${j}" aria-label="${e(it.t)}, ${e(p.n)}">
        ${sl(it,p,false,false,j)}
        <span class="pn">${e(p.n)}</span>
      </button>
      ${(isSongItem(it)||isPptHandoffItem(it))?'':`<button type="button" class="th-add" data-add-after-i="${i}" data-add-after-j="${j}" title="Add slide after this one">+ Add slide</button>`}
    </div>`).join('');
  }).join('');

  if(stripEl){stripEl.scrollTop=prevScrollY;stripEl.scrollLeft=prevScrollX;}
  paintSlides();
}

function rStage(){
  if(elDrag)return; // never rebuild DOM while dragging
  if(!A||!S.length){
    $('#stage').innerHTML='<div class="cv" style="display:grid;place-items:center;color:var(--mu);padding:24px"><div style="text-align:center"><b>No slides yet</b><br><small>Add a section with Add Slide in the left column</small></div></div>';
    paintEls();
    return;
  }
  const it=A;
  // Clear layer sits ABOVE the whole slide (outside .sl) so text boxes / elements cannot show through
  let ovHtml='';
  if(clr==='black') ovHtml='<div class="ov black" style="position:absolute;inset:0;z-index:10000;background:#000"></div>';
  else if(clr==='white') ovHtml='<div class="ov white" style="position:absolute;inset:0;z-index:10000;background:#fff"></div>';
  else if(clr==='logo'){
    ovHtml=`<div class="ov logo" style="position:absolute;inset:0;z-index:10000;background:${logoBgCss()};display:flex;${logoAlignCss()};overflow:hidden">${logoMarkup()}</div>`;
  }

  const inPresenter=document.body.classList.contains('presenter-mode');
  if(live&&inPresenter){
    // Presenter preview: clear covers the card completely
    $('#stage').innerHTML=`<div class="cv lv" style="position:relative">${stillHtml(it,sls(it)[aj])}${ovHtml}<div class="bd r">Live</div></div>`;
    Object.values(vpool).forEach(v=>{try{v.pause()}catch{}});
  }else{
    $('#stage').innerHTML=`<div class="cv ${live?'lv':''}" style="position:relative">${sl(it,sls(it)[aj],ed,true,aj)}${ovHtml}<div class="bd ${live?'r':''}">${ed?'Editing '+e(sls(it)[aj].n)+', Esc when done':(live?'Live':'Preview')+(clr?', screen cleared. Esc to show':'')}</div></div>`;
    attachVideos();
    // Apply mute/volume for per-slide videos
    document.querySelectorAll('#stage video.slide-video,#stage video.el-video').forEach(v=>{
      v.playbackRate=Number(sls(it)[aj]?.videoMedia?.settings?.playbackRate)||1;
      const vol=parseFloat(v.dataset.vol);
      if(!isNaN(vol))v.volume=Math.max(0,Math.min(1,vol));
      // muted attribute already set in HTML; if not muted, ensure audio can play
      if(!v.hasAttribute('muted')){v.muted=false}
    });
  }
  document.querySelectorAll('[data-c]').forEach(b=>b.classList.toggle('on',b.dataset.c==clr));
  const eb=$('#edb');if(eb){eb.textContent=ed?'Done editing':'Edit text';eb.classList.toggle('on',ed)}
  if(ed){const c=$('#ce');if(c){makeEditable(c);if(editFocus){editFocus=false;c.focus();caretToEnd(c)}}}
  if(live){
    if(document.body.classList.contains('presenter-mode'))renderPresenterView();
    pushAudienceStage();
  }
  paintEls();
}

let editFocus=false,saveTimer=null;
const saveSoon=()=>{clearTimeout(saveTimer);saveTimer=setTimeout(saveState,400)};
function makeEditable(c){
  c.setAttribute('contenteditable','plaintext-only');
  if(c.contentEditable!=='plaintext-only')c.setAttribute('contenteditable','true');
  c.setAttribute('role','textbox');c.setAttribute('aria-label','Slide text');
}
function caretToEnd(c){
  try{const r=document.createRange();r.selectNodeContents(c);r.collapse(false);const s=getSelection();s.removeAllRanges();s.addRange(r)}catch{}
}
function startEdit(){
  if(live||document.body.classList.contains('presentation-mode'))return;
  if(clr){toast('Press Esc to show the slide first, then edit its text.');return}
  if(A.k==='PowerPoint'||A.k==='PDF'){
    toast('Double-click a text box to edit it, or add one from Edit.');
    return;
  }
  if(!isSongItem(A)){
    const part=sls(A)[aj];
    if(part&&typeof ensureElements==='function')ensureElements(part,A);
    const selEl=getSelectedEl();
    if(selEl&&selEl.kind==='text'){
      const node=document.querySelector('#stage [data-el="'+selEl.id+'"]');
      openTextElEditor(selEl,node);
      return;
    }
    const els=(part&&part.elements||[]).filter(x=>x.kind==='text');
    if(els.length===1){
      const node=document.querySelector('#stage [data-el="'+els[0].id+'"]');
      openTextElEditor(els[0],node);
      return;
    }
    if(els.length>1){
      toast('Double-click a text box to edit it.');
      return;
    }
  }
  ed=true;editFocus=true;rStage();
}
function stopEdit(){
  if(!ed)return;
  const part=sls(A)[aj];if(part)part.tx=part.tx.replace(/\n+$/,'');
  if(isSongItem(A)){try{refitSong(A)}catch(err){console.warn(err)}}
  ed=false;clearTimeout(saveTimer);saveState();updateCurrentThumb();rStage();rPanel();rStrip();
}

function rLib(){
  const q=($('#q')?.value||'').toLowerCase();
  const filtered=LIB.filter(x=>(q||!x.churchBank)&&x.t.toLowerCase().includes(q));
  $('#ll').innerHTML=filtered.map((x,i)=>{
    const origIndex=LIB.indexOf(x);
    return `<div class="li">
      <div><b>${e(x.t)}</b> <small style="display:block;color:var(--mu)">${x.ord.length} slides${x.churchBank?' · Church song bank':''}${x.credits&&x.credits.ccli?' · CCLI '+e(x.credits.ccli):''}</small></div>
      <div style="display:flex;gap:6px">
        ${x.churchBank?'':'<button data-edit-lib="'+origIndex+'">Edit</button>'}
        <button data-add="${origIndex}" class="pri">Add</button>
      </div>
    </div>`;
  }).join('');
}

function seg(k,a,v){
  return `<div class="seg">${a.map(x=>`<button data-${k}="${x}" class="${v==x?'on':''}">${x[0].toUpperCase()+x.slice(1)}</button>`).join('')}</div>`;
}

function textStyleControls(f,{song=false}={}){
  const col=f.color||'#ffffff';
  return `<h3>Typography & Text Style</h3>
  <label>Font Family
    <select data-k="ff">
      <option value="sans" ${f.ff==='sans'||!f.ff?'selected':''}>Sans-Serif (Figtree)</option>
      <option value="serif" ${f.ff==='serif'?'selected':''}>Serif (Lora)</option>
    </select>
  </label>
  <label>Text Casing
    <select data-k="tc">
      <option value="none" ${f.tc==='none'||!f.tc?'selected':''}>Normal Case</option>
      <option value="upper" ${f.tc==='upper'?'selected':''}>ALL CAPS (Uppercase)</option>
      <option value="title" ${f.tc==='title'?'selected':''}>Title Case</option>
    </select>
  </label>
  <label>Font colour
    <span style="display:flex;align-items:center;gap:8px;margin-top:5px">
      <input type="color" data-k="color" value="${col}" style="width:48px;height:34px;padding:2px;cursor:pointer">
      <input type="text" data-k="color-hex" value="${col}" placeholder="#ffffff" style="flex:1;font-family:ui-monospace,monospace">
    </span>
  </label>
  <label>Font Size <span style="color:var(--mu);font-variant-numeric:tabular-nums" id="fs-val">${(f.fs!=null?+f.fs:5.2).toFixed(1)}</span>
    <input type="range" data-k="fs" min="1" max="10" step="0.1" value="${Math.max(1,Math.min(10,f.fs!=null?+f.fs:5.2))}"></label>
  <div style="display:flex;align-items:center;gap:18px;margin:8px 0 14px">
    <label style="margin:0;display:inline-flex;align-items:center;gap:6px"><input type="checkbox" data-k="b" ${f.b?'checked':''}> Bold</label>
    <label style="margin:0;display:inline-flex;align-items:center;gap:6px"><input type="checkbox" data-k="sh" ${f.sh?'checked':''}> Shadow</label>
  </div>
  <label>Horizontal Alignment</label>
  ${song?`<label>Left/right margin <span style="color:var(--mu);font-variant-numeric:tabular-nums" id="lyric-margin-val">${Number(f.mb??8)}%</span><input type="range" data-k="mb" min="0" max="25" step="0.5" value="${Math.max(0,Math.min(25,Number(f.mb??8)))}" aria-label="Lyric left and right margin"></label>`:''}
  ${seg('al',['left','center','right'],f.al||'center')}
  <label>Vertical Position</label>
  ${seg('va',['top','middle','bottom'],f.va||'middle')}`;
}

function openLibraryModal(insertAt=null){
  pendingSectionInsert=insertAt;
  const root=$('#library-content');
  root.innerHTML=`<form id="song-search-form" style="display:flex;gap:6px;margin:0 0 10px">
    <input type="text" id="q" placeholder="Search for song" autocomplete="off" style="flex:1;min-width:0;margin:0">
    <button class="pri" type="submit">Search</button>
  </form>
  <h3 style="margin:0 0 6px;font-size:calc(12px * var(--ui-scale));text-transform:uppercase;letter-spacing:.05em;color:var(--mu)">Your library</h3>
  <div id="ll"></div>
  <h3 style="margin:16px 0 6px;font-size:calc(12px * var(--ui-scale));text-transform:uppercase;letter-spacing:.05em;color:var(--mu)">Online results</h3>
  <div id="song-search-state" class="song-search-state" role="status" aria-live="polite">Type a title and press Search for online lyrics.</div>
  <div id="song-search-results" class="song-search-results" role="listbox" aria-label="Song search results"></div>
  <h3 style="margin-top:20px">Add Custom Song / Paste Lyrics</h3>
  <p class="note">Type or paste full lyrics with section headers (Verse 1, Chorus, etc.).</p>
  <input type="text" id="nt" placeholder="Song title (e.g. 10,000 Reasons)">
  <textarea id="nl" rows="7" placeholder="Paste lyrics here. Start sections with Verse 1, Chorus, Bridge, or leave blank lines between sections."></textarea>
  <button id="ns" class="pri" style="width:100%">Add Custom Song to Library & Service</button>`;
  $('#library-modal').hidden=false;
  songResults=[];
  rLib();
  renderSongResults();
  $('#q').focus();
}

function openEditSongModal(index){
  editingLibIndex=index;
  pendingOnlineSong=null;
  const song=LIB[index];
  if(!song) return;
  const c=song.credits||{};
  const rawText=songToText(song,true);
  const root=$('#edit-song-content');
  $('#edit-song-title').textContent='Edit Song & Sections';
  $('#edit-song-header-actions')?.remove();
  $('#edit-song-title').insertAdjacentHTML('afterend','<div id="edit-song-header-actions" style="display:flex;gap:8px;align-items:center;margin-left:auto"><button type="button" id="delete-song-lib-top" style="color:#c22">Delete from Library</button><button type="button" id="cancel-song-edit-top">Cancel</button><button type="button" id="save-song-edit-top" class="pri">Done</button></div>');
  root.innerHTML=`<label>Song Title<input type="text" id="edit-song-title-input" value="${e(song.t)}"></label>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:10px">
    <label style="margin:0">Artist / Author<input type="text" id="edit-song-artist" value="${e(c.artist||c.authors||c.author||'')}"></label>
    <label style="margin:0">Writers<input type="text" id="edit-song-writers" value="${e(c.writers||c.writer||'')}"></label>
    <label style="margin:0">Release year<input type="text" id="edit-song-year" value="${e(c.year||'')}"></label>
    <label style="margin:0">CCLI song #<span style="display:flex;gap:6px"><input type="text" id="edit-song-ccli" value="${e(c.ccli||c.CCLI||'')}" placeholder="e.g. 7117038" style="margin:0"><button type="button" data-ccli-lookup="1" title="Search SongSelect for this song">Find</button></span></label>
  </div>
  <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;margin:12px 0;background:var(--bg);padding:10px 12px;border-radius:8px;border:1px solid var(--ln)">
    <div>
      <b>Lines per slide</b>
      <small style="display:block;color:var(--mu)">Automatically split long sections into slides with this line count</small>
    </div>
    <div style="display:flex;align-items:center;gap:6px">
      <select id="edit-song-lines-select" style="margin:0;width:auto;min-width:140px">
        <option value="1">1 line / slide</option>
        <option value="2">2 lines / slide</option>
        <option value="3">3 lines / slide</option>
        <option value="4" selected>4 lines / slide</option>
        <option value="6">6 lines / slide</option>
        <option value="0">Full section (no split)</option>
      </select>
      <button id="reformat-lines-btn" type="button">Re-split</button>
    </div>
  </div>
  <label>Lyrics & Section Breakdown<textarea id="edit-song-lyrics-input" rows="20" style="font-family:monospace;font-size:calc(13px * var(--ui-scale));min-height:min(48vh,420px);height:min(48vh,420px)">${e(rawText)}</textarea></label>
  `;
  $('#edit-song-modal').hidden=false;
}

function closeLibraryModal(){
  clearTimeout(searchDebounce);
  searchController?.abort();
  lyricsController?.abort();
  $('#library-modal').hidden=true;
}

function closeAllWindows(){
  pendingOnlineSong=null;
  try{if(!$('#edit-song-modal').hidden)closeEditSongModal()}catch{}
  try{if(!$('#struct-modal').hidden)closeStruct()}catch{}
  try{closeMediaLibrary()}catch{}
  try{closeBackgroundPicker()}catch{}
  try{if(!$('#library-modal').hidden)closeLibraryModal()}catch{}
  try{closeShareModal()}catch{}
  try{closeTeamModal()}catch{}
  try{closeCloudModal()}catch{}
  try{closeConflict()}catch{}
  try{closeEmailCompose()}catch{}
  const lc=$('#logo-controls-modal');if(lc)lc.hidden=true;
  const sc=$('#shortcuts-modal');if(sc)sc.hidden=true;
  if($('#settings')?.classList.contains('open'))closeSettingsAndRefresh();
}
let _bdDown=null;
document.addEventListener('pointerdown',ev=>{_bdDown=ev.target},true);
document.addEventListener('click',ev=>{
  const t=ev.target;
  if(!t||!t.classList)return;
  if(!(t.classList.contains('app-modal')||t.id==='settings'))return;
  if(_bdDown!==t)return;
  if(t.id==='import-modal'||t.id==='close-confirm-modal')return;
  closeAllWindows();
},true);

function closeEditSongModal(){
  $('#edit-song-modal').hidden=true;
  editingLibIndex=null;
  // keep pendingOnlineSong only if still reviewing — cancel clears it separately
}

function reformatEditedSongLines(){
  const lyricsInput=$('#edit-song-lyrics-input');
  const linesSelect=$('#edit-song-lines-select');
  if(!lyricsInput||!linesSelect) return;
  const lps=parseInt(linesSelect.value,10);
  const parsed=parse(lyricsInput.value, lps);
  lyricsInput.value=songToText(parsed,false);
  toast(`Re-formatted into ${parsed.parts.length} slides (${lps?lps+' lines per slide':'full sections'}).`);
}

function saveEditedSong(){
  if(editingLibIndex==null) return;
  const title=$('#edit-song-title-input').value.trim()||'Untitled Song';
  const lyrics=$('#edit-song-lyrics-input').value.trim();
  const lps=parseInt($('#edit-song-lines-select')?.value||'0', 10);
  if(!lyrics) return toast('Please enter some lyrics.');
  const parsed=parse(lyrics, lps);
  const prev=(LIB[editingLibIndex]&&LIB[editingLibIndex].credits)||{};
  const credits={
    ...prev,
    artist:($('#edit-song-artist')?.value||'').trim(),
    authors:($('#edit-song-artist')?.value||'').trim(),
    writers:($('#edit-song-writers')?.value||'').trim(),
    year:($('#edit-song-year')?.value||'').trim(),
    ccli:($('#edit-song-ccli')?.value||'').trim(),
  };
  LIB[editingLibIndex]={t:title,...parsed,credits,lps};
  saveState();
  closeEditSongModal();
  rLib();
}

async function deleteSongFromLib(){
  if(editingLibIndex==null) return;
  if(await appConfirm('Are you sure you want to delete this song from your library?',{title:'Delete song',confirmLabel:'Delete',danger:true})){
    LIB.splice(editingLibIndex,1);
    saveState();
    closeEditSongModal();
    rLib();
    toast('Song removed from library.');
  }
}

function rPanel(){
  const songStructureAction=$('#song-structure-action');
  if(songStructureAction)songStructureAction.hidden=!isSongItem(S[sel]);
  if(!S.length||!S[sel]){
    $('#panel').innerHTML='<div class="panel-body"><p class="note">No sections in this service yet. Use <b>Add Slide</b> in the slide column.</p></div>';
    return;
  }
  const it=S[sel],f=it.f||{fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'};
  const structureLabel='Edit';
  const _bm=(tab=='Style'&&!isSongItem(it))?bgMultiList():[];
  const stickyNote=tab=='Style'
  ?(_bm.length>1?`<p class="note">Background for the <b>${_bm.length} selected slides</b>.</p>`
   :(!isSongItem(it)&&it.bgAll!==true)?`<p class="note">Background for <b>${e(it.t)}</b>, this slide only.</p>`
   :`<p class="note">Background for <b>${e(it.t)}</b> (${it.ord.length} slide${it.ord.length>1?'s':''}).</p>`)
  :'';
  const curPartForTabs=sls(it)[si]||sls(it)[0];
  const hasVideoTab=!!(it.k==='Video'||(curPartForTabs&&curPartForTabs.videoMedia));
  const hasSlidesTab=it.k==='PowerPoint';
  const tabList=['Style','Structure','Add'];
  if(hasVideoTab) tabList.push('Video');
  if(hasSlidesTab) tabList.push('Slides');
  if(tab==='Team') tab='Structure';
  if(tab==='Video'&&!hasVideoTab) tab='Structure';
  if(tab==='Slides'&&!hasSlidesTab) tab='Structure';
  const tabLabel=n=>n==='Structure'?structureLabel:n==='Add'?'Add':n==='Style'?'Background':n==='Video'?'Video Settings':n==='Slides'?'Slide Settings':n;
  let h=`<div class="panel-sticky"><div class="tabs">${tabList.map(n=>`<button data-tab="${n}" class="${n==tab?'on':''}">${tabLabel(n)}</button>`).join('')}</div>${stickyNote}</div><div class="panel-body">`;

    if(tab=='Style'){
    const song=isSongItem(it);
    const multi=song?[]:bgMultiList();
    const bgPart=sls(it)[Math.min(si,it.ord.length-1)]||sls(it)[0];
    const own=!song&&!!(bgPart&&bgPart.ownBg);
    const bgIdx=own&&bgPart.bg!=null?bgPart.bg:it.bg;
    const customV=(own&&bgPart.customBg)||it.customBg||'#16213e';
    const mediaV=own?bgPart.bgMedia:(bgPart.bgMedia||it.bgMedia);
    const dimV=own&&bgPart.dim!=null?bgPart.dim:it.dim;
    const perSlide=!song&&(multi.length>1||it.bgAll!==true);
    h+=`
    <h3>Background</h3>
    ${song?'':(multi.length>1
      ?`<p class="note" style="margin:0 0 10px">Applies to the ${multi.length} selected slides.</p>`
      :`<label style="display:flex;align-items:center;gap:8px;margin:0 0 10px"><input type="checkbox" id="bg-all-slides" ${it.bgAll===true?'checked':''}> Use for all slides in the section</label>`)}
    <div class="sw">${BG.map((b,i)=>{
      const on=i==bgIdx;
      if(b.pick)return `<button class="swb ${on?'on':''} swb-colour-picker" data-bg="${i}">${b.n}</button>`;
      return `<button class="swb ${b.v?'wave':''} ${on?'on':''}" data-bg="${i}" style="${b.v?'':'background:'+b.c}">${b.n}</button>`;
    }).join('')}</div>
    ${(BG[bgIdx]&&BG[bgIdx].pick)?`<label style="margin-top:10px">Solid colour
      <span style="display:flex;align-items:center;gap:8px;margin-top:5px">
        <input type="color" id="custom-bg-color" value="${e(customV)}" style="width:48px;height:34px;padding:2px;cursor:pointer">
        <input type="text" id="custom-bg-hex" value="${e(customV)}" placeholder="#16213e" style="flex:1;font-family:ui-monospace,monospace">
      </span>
    </label>`:''}
    <button id="add-background-btn" type="button" style="width:100%;margin-top:8px">Add a background</button>
    ${(()=>{
      const chosenName=mediaV?(mediaV.name||'Custom background'):((BG[bgIdx]&&BG[bgIdx].pick)?'Custom colour':((BG[bgIdx]&&BG[bgIdx].n)||'Background'));
      const previewPart=bgPart||{n:'Preview',tx:''};
      return `<div class="background-current"><div class="background-current-head"><strong title="${e(chosenName)}">${e(chosenName)}</strong><button id="bg-clear" type="button">Remove</button></div><div class="background-current-preview">${sl(it,previewPart,false,false,0)}</div><p class="note" style="margin:7px 0 0">Preview of the current slide with the selected background.</p></div>`;
    })()}
    <label>Darken background<input type="range" data-k="dim" min="0" max="80" value="${dimV}"></label>
    ${mediaV?.kind==='video'?(()=>{const speed=Number(mediaV.playbackRate)||1,sliderValue=speed<=1?Math.round(((speed-.1)/.9)*50):Math.round(50+((speed-1)/4)*50);return `<label>Background video playback speed <span style="color:var(--mu)" id="background-video-speed-val">${speed.toFixed(1).replace(/\.0$/ , '')}x</span><input type="range" id="background-video-playback-speed" min="0" max="100" step="1" value="${sliderValue}" aria-label="Background video playback speed"></label>`})():''}`;
  }
  else if(tab=='Structure'){
    const cur=sls(it)[si],song=it.k==='Song';
    if(song){
      h+=`<button id="sb-open" class="pri" style="width:100%">Change song structure</button>
      <label>Lines per slide<select id="lps-select">${lpsOptions(curLps(it))}</select></label>
      <h3>Song lyrics: ${e(cur.n)}</h3>
      <textarea id="pt" rows="7">${e(cur.tx)}</textarea>
      ${textStyleControls(f,{song:true})}`;
    }else if(it.k==='PowerPoint'||it.k==='PDF'){
      const cur=sls(it)[si];
      const els=ensureElements(cur,it)||[];
      const selEl=getSelectedEl();
      const textEl=selEl&&selEl.kind==='text'?selEl:null;
      const mediaEl=selEl&&selEl.kind!=='text'?selEl:null;
      const styleTarget=textEl?(textEl.f||(textEl.f={fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'})):f;
      const sortedEls=[...els].sort((a,b)=>(b.z||0)-(a.z||0));
      h+=`<p class="note">Imported slides are pictures. You can still add text boxes, images, or video on top.</p>
      <h3>Add to slide</h3>
      <div style="display:flex;gap:8px;margin:0 0 12px;flex-wrap:wrap">
        <button type="button" id="add-text-box">+ Text box</button>
        <button type="button" id="add-canvas-image">+ Image</button>
        <button type="button" id="add-slide-video">+ Video</button>
      </div>
      <input type="file" id="slide-video-file" accept="video/*" hidden>
      ${mediaEl?'':textEl?`<label>Selected text box<textarea id="el-tx" rows="5">${e(textEl.tx||'')}</textarea></label>
      <p class="note">Editing style of the <b>selected text box</b>.</p>`
        :`<p class="note">${els.length?'Select a text box on the slide to edit its words and style.':'Add a text box, image, or video on top of this imported slide.'}</p>`}
      ${mediaEl?'':textStyleControls(styleTarget)}
      ${els.length?`<h3>Elements on this slide</h3>
      <div class="el-list-label">↑ Foreground</div>
      <div id="el-list" class="el-list">
      ${sortedEls.map(el=>`<div class="el-row ${el.id===elSel?'cur':''}" data-pick-el="${e(el.id)}" data-el-id="${e(el.id)}">
        <span class="el-drag" data-el-drag="${e(el.id)}" title="Drag to reorder">⋮⋮</span>
        <span class="el-label">${el.kind==='text'?e((el.tx||'').split('\n')[0].slice(0,36)||'Text box'):'🖼 Image'}</span>
        <button type="button" data-dup-el="${e(el.id)}" title="Duplicate">⧉</button>
        <button type="button" data-del-el="${e(el.id)}" title="Remove" style="color:#c22">✕</button>
      </div>`).join('')}
      </div>
      <div class="el-list-label">↓ Background</div>`:''}`;
    }else{
      const els=ensureElements(cur,it)||[];
      const selEl=getSelectedEl();
      const textEl=selEl&&selEl.kind==='text'?selEl:null;
      const mediaEl=selEl&&selEl.kind!=='text'?selEl:null;
      const styleTarget=textEl?(textEl.f||(textEl.f={fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'})):f;
      const sortedEls=[...els].sort((a,b)=>(b.z||0)-(a.z||0));
      h+=`<h3>Add to slide</h3>
      <div style="display:flex;gap:8px;margin:0 0 12px;flex-wrap:wrap">
        <button type="button" id="add-text-box">+ Text box</button>
        <button type="button" id="add-canvas-image">+ Image</button>
        <button type="button" id="add-slide-video">+ Video</button>
      </div>
      <input type="file" id="slide-video-file" accept="video/*" hidden>
      ${mediaEl?'':textEl?`<label>Selected text box<textarea id="el-tx" rows="5">${e(textEl.tx||'')}</textarea></label>
      <p class="note">Editing style of the <b>selected text box</b>.</p>`
        :`<p class="note">${els.length?'Select a text box on the slide to edit its words and style.':'Add a text box, image, or video to this slide.'}</p>`}
      ${mediaEl?'':textStyleControls(styleTarget)}
      ${els.length?`<h3>Elements on this slide</h3>
      <div class="el-list-label">↑ Foreground</div>
      <div id="el-list" class="el-list">
      ${sortedEls.map(el=>`<div class="el-row ${el.id===elSel?'cur':''}" data-pick-el="${e(el.id)}" data-el-id="${e(el.id)}">
        <span class="el-drag" data-el-drag="${e(el.id)}" title="Drag to reorder">⋮⋮</span>
        <span class="el-label">${el.kind==='text'?e((el.tx||'').split('\\n')[0].slice(0,36)||'Text box'):'🖼 Image'}</span>
        <button type="button" data-dup-el="${e(el.id)}" title="Duplicate">⧉</button>
        <button type="button" data-del-el="${e(el.id)}" title="Remove" style="color:#c22">✕</button>
      </div>`).join('')}
      </div>
      <div class="el-list-label">↓ Background</div>`:''}
`;
    }
  }
  else if(tab=='Video'){
    const cur=sls(it)[si]||{};
    const vm=cur.videoMedia||(it.k==='Video'?(cur.videoMedia||{}):null);
    if(!vm){
      h+=`<p class="note">No video on this slide. Add one from Edit → + Video, or add a Video section.</p>`;
    }else{
      const vs=vm.settings||{};
      const meta=mediaMetaFor(vm.id)||{};
      h+=`<h3>Video on this slide</h3>
      <div class="rw"><div style="flex:1"><b>🎬 ${e(meta.name||vm.name||'Video')}</b></div>
      <button id="remove-slide-video" type="button" style="color:#c22">Remove</button></div>
      <label style="display:flex;align-items:center;gap:8px;margin-top:12px"><input type="checkbox" id="video-autoplay" ${vs.autoplay?'checked':''}> Auto play</label>
      <label style="display:flex;align-items:center;gap:8px"><input type="checkbox" id="video-loop" ${vs.loop?'checked':''}> Loop</label>
      <label style="display:flex;align-items:center;gap:8px"><input type="checkbox" id="video-muted" ${vs.muted===true?'checked':''}> Mute video</label>
      <p class="note">When unmuted, audio plays during presentation (audience and presenter). Background videos stay muted so they do not clash with live sound.</p>
      <label>Volume <span style="color:var(--mu)" id="video-vol-val">${Math.round((vs.volume!=null?vs.volume:1)*100)}</span>%
        <input type="range" id="video-volume" min="0" max="100" value="${Math.round((vs.volume!=null?vs.volume:1)*100)}"></label>
      <p class="note">Volume applies when Mute is off.</p>`;
    }
  }
  else if(tab=='Slides'){
    const mode=it.viewMode||(it.k==='PowerPoint'?'powerpoint':'church');
    h+=`<h3>Slides Settings</h3>
    <p class="note" style="margin-bottom:10px">View presentation with:</p>
    <label style="display:flex;align-items:flex-start;gap:10px;margin:8px 0;padding:10px;border:1px solid var(--ln);border-radius:8px;background:var(--bg);cursor:pointer">
      <input type="radio" name="ppt-view-mode" value="church" ${mode==='church'?'checked':''} style="margin-top:3px">
      <span><b>Church Presenter</b><small style="display:block;color:var(--mu);margin-top:2px">Show slides inside this app.</small></span>
    </label>
    <label style="display:flex;align-items:flex-start;gap:10px;margin:8px 0;padding:10px;border:1px solid var(--ln);border-radius:8px;background:var(--bg);cursor:pointer">
      <input type="radio" name="ppt-view-mode" value="powerpoint" ${mode==='powerpoint'?'checked':''} style="margin-top:3px">
      <span><b>Microsoft PowerPoint</b> <small>(Windows)</small><small style="display:block;color:var(--mu);margin-top:2px">Open in PowerPoint slideshow when this section plays.</small></span>
    </label>
    <p class="note">${it.sourceFile?('File: <b>'+e(it.sourceFile)+'</b>'+(it.originalFileId?' · saved':' · re-import to enable hand-off')):'Import a .pptx to enable PowerPoint mode.'}</p>
    <button type="button" id="ppt-test-open" style="margin-top:8px;width:100%">Test open in PowerPoint</button>`;
  }
  else if(tab=='Add'){
    h+=`<h3>Add slide</h3>
    <p class="note">Add a new slide inside the current service item.</p>
    <button id="add-slide-to-item" class="pri" type="button" style="width:100%">+ Add slide</button>
    <h3 style="margin-top:18px">New section</h3>
    <div class="chips" style="margin:8px 0"><button id="add-song-from-panel" class="pri" type="button">+ Add song</button>
      ${['Welcome','Announcements','Bible Reading','Prayer','Offering','Benediction'].map(n=>`<button data-qs="${n}">+ ${n}</button>`).join('')}</div>
    <input type="text" id="sn" placeholder="Or type custom title">
    <button id="sadd" class="pri" style="width:100%">Add blank section</button>
    <h3 style="margin-top:18px">Add logo slide</h3>
    <p class="note">Add a slide showing the church logo, like pressing L while presenting.</p>
    <button id="add-logo-slide" type="button" class="pri" style="width:100%">＋ Add logo slide</button>
    <h3 style="margin-top:18px">Add Video</h3>
    <button id="add-video-slide" class="pri" type="button" style="width:100%">🎬 Add Video</button>
    <input type="file" id="add-video-slide-file" accept="video/*" hidden>
    <h3 style="margin-top:18px">Import slides</h3>
    <input type="file" id="fi" accept=".pptx,.ppt,.pdf,.cpres,application/json" style="margin-top:6px">
    <div id="ppt-import-status" class="note" style="margin-top:8px"></div>`;
  }
  else if(tab=='Team'){
    /* Team panel moved to header button — keep handler for compatibility */
    openTeamModal();
    tab='Structure';
    return rPanel();
  }

  h+='</div>';
  $('#panel').innerHTML=h;
  if(tab=='Structure')document.querySelectorAll('[data-im="y"]').forEach(x=>x.value=100-(+x.value));
}

const lyricPreviewCache=new Map();
function renderSongResults(){
  const root=$('#song-search-results');
  if(!root)return;
  root.innerHTML=songResults.map((r,i)=>{
    return `<div class="song-result" data-search-result="${i}">
      <div class="song-result-main" tabindex="0">
        <b>${e(r.title)}</b>
        <small>${e(r.authors||'Worship Song')}</small>
        <div class="song-preview" data-preview="${i}" hidden>Hover to preview lyrics…</div>
      </div>
      <div class="song-result-actions">
        <button data-search-add="${i}" class="pri">Add</button>
      </div>
    </div>`;
  }).join('');
}
async function showSongPreview(index){
  const r=songResults[index],box=document.querySelector('[data-preview="'+index+'"]');
  if(!r||!box)return;
  box.hidden=false;
  if(lyricPreviewCache.has(r.id)){box.innerHTML=lyricPreviewCache.get(r.id);return}
  box.textContent='Loading lyrics preview…';
  try{
    const text=await getLyrics(r.id);
    const preview=String(text||'').trim().slice(0,420);
    const html=e(preview)+(text&&String(text).length>420?'…':'');
    lyricPreviewCache.set(r.id,html);
    if(document.querySelector('[data-preview="'+index+'"]'))document.querySelector('[data-preview="'+index+'"]').innerHTML=html;
  }catch(error){box.textContent='Could not load a lyrics preview.'}
}
function hideSongPreview(index){
  const box=document.querySelector('[data-preview="'+index+'"]');if(box)box.hidden=true;
}

function setSearchState(message,type=''){
  const el=$('#song-search-state');if(!el)return;
  el.className='song-search-state'+(type?' '+type:'');
  el.textContent=message;
}

async function runSongSearch(query=$('#q')?.value||$('#song-search-query')?.value||''){
  const q=query.trim();clearTimeout(searchDebounce);searchController?.abort();
  if(!q){songResults=[];renderSongResults();return setSearchState('Enter a song title to search online.')}
  const controller=new AbortController();searchController=controller;songResults=[];renderSongResults();
  setSearchState('Searching free online database...');
  try{
    const results=await searchSongs(q,{signal:controller.signal,onResults:partial=>{if(searchController!==controller)return;songResults=partial;renderSongResults();setSearchState('Found '+partial.length+' so far, still searching...')}});
    if(searchController!==controller)return;
    // Prefer results that already expose section labels (Verse/Chorus/…)
    songResults=[...results].sort((a,b)=>{
      const sa=lyricStructureScore((a&&(a.snippet||a.lyrics||a.preview||a.title))||'')+(a.hasStructure?20:0);
      const sb=lyricStructureScore((b&&(b.snippet||b.lyrics||b.preview||b.title))||'')+(b.hasStructure?20:0);
      return sb-sa;
    });
    renderSongResults();
    if(results.length) setSearchState(`Found ${results.length} match${results.length===1?'':'es'}. Click Add to insert.`);
    else setSearchState('No matches found. Try another song title or artist.');
  }catch(error){
    if(error.name==='AbortError')return;
    setSearchState(error.message||'Search failed. Check your connection.','error');
  }
}

let pendingOnlineSong=null; // {song, resultTitle} while reviewing online search add

async function addSearchResult(index){
  const result=songResults[index];
  if(!result||addingSongId)return;
  addingSongId=result.id;
  const button=document.querySelector('[data-search-add="'+index+'"]');
  if(button){button.disabled=true;button.textContent='Loading…'}
  try{
    const lyricsText=await getLyrics(result.id);
    const song={t:result.title,...parse(lyricsText, 4)};
    song.credits={
      authors:result.authors||result.author||'',
      artist:result.artist||result.artists||'',
      writers:result.writers||result.writer||'',
      writer:result.writer||'',
      composer:result.composer||result.composers||'',
      lyricist:result.lyricist||result.lyricists||'',
      songwriter:result.songwriter||result.songwriters||'',
      publisher:result.publisher||result.publishing||'',
      copyright:result.copyright||result.rights||'',
      license:result.license||'',
      year:result.year||result.releaseDate||result.released||result.published||result.publicationDate||result.date||'',
      album:result.album||'',
      ccli:result.ccli||result.CCLI||'',
      source:result.source||result.provider||''
    };
    // Pull a year out of copyright/lyrics if the API didn't provide one
    if(!song.credits.year){
      const blob=[lyricsText,song.credits.copyright,song.credits.publisher,result.title].join(' ');
      const ym=String(blob).match(/\b(?:©|\(c\)|copyright)?\s*((?:19|20)\d{2})\b/i);
      if(ym)song.credits.year=ym[1];
    }
    if(!song.credits.ccli){const cm=String(lyricsText||'').match(/CCLI\s*(?:Song\s*)?(?:#|No\.?|Number)?\s*:?\s*(\d{4,9})/i);if(cm)song.credits.ccli=cm[1]}
    if(!song.ord.length||!song.parts.some(p=>p.tx&&p.tx.trim())) throw new Error('No sections found in lyrics.');
    pendingOnlineSong={song,resultTitle:result.title};
    openReviewOnlineSongModal();
  }catch(error){
    toast(error.message||'Could not load lyrics. Try pasting directly.');
  }finally{
    if(button){button.disabled=false;button.textContent='Add'}
    addingSongId=null;
  }
}

/** Review/edit an online song (title + lyrics) before adding to library & service. */
function openReviewOnlineSongModal(){
  if(!pendingOnlineSong||!pendingOnlineSong.song)return;
  editingLibIndex=null;
  const song=pendingOnlineSong.song;
  const c=song.credits||{};
  const rawText=songToText(song,true);
  const root=$('#edit-song-content');
  $('#edit-song-header-actions')?.remove();
  $('#edit-song-title').textContent='Review song before adding';
  root.innerHTML=`<p class="note" style="margin:0 0 10px">Check title, credits, and section labels, then add to your service.</p>
  <label>Song Title<input type="text" id="edit-song-title-input" value="${e(song.t)}"></label>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:10px">
    <label style="margin:0">Artist / Author<input type="text" id="edit-song-artist" value="${e(c.artist||c.authors||c.author||'')}" placeholder="e.g. Bethel Music"></label>
    <label style="margin:0">Writers<input type="text" id="edit-song-writers" value="${e(c.writers||c.writer||c.songwriter||'')}" placeholder="Songwriters"></label>
    <label style="margin:0">Release year<input type="text" id="edit-song-year" value="${e(c.year||'')}" placeholder="e.g. 2013"></label>
    <label style="margin:0">CCLI song #<span style="display:flex;gap:6px"><input type="text" id="edit-song-ccli" value="${e(c.ccli||c.CCLI||'')}" placeholder="e.g. 7117038" style="margin:0"><button type="button" data-ccli-lookup="1" title="Search SongSelect for this song">Find</button></span></label>
  </div>
  <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;margin:12px 0;background:var(--bg);padding:10px 12px;border-radius:8px;border:1px solid var(--ln)">
    <div>
      <b>Lines per slide</b>
      <small style="display:block;color:var(--mu)">Split long sections into slides with this line count</small>
    </div>
    <div style="display:flex;align-items:center;gap:6px">
      <select id="edit-song-lines-select" style="margin:0;width:auto;min-width:140px">
        <option value="1">1 line / slide</option>
        <option value="2">2 lines / slide</option>
        <option value="3">3 lines / slide</option>
        <option value="4" selected>4 lines / slide</option>
        <option value="6">6 lines / slide</option>
        <option value="0">Full section (no split)</option>
      </select>
      <button id="reformat-lines-btn" type="button">Re-split</button>
    </div>
  </div>
  <label>Lyrics & Section Breakdown<textarea id="edit-song-lyrics-input" rows="22" style="font-family:monospace;font-size:calc(13px * var(--ui-scale));min-height:min(52vh,480px);height:min(52vh,480px)">${e(rawText)}</textarea></label>
  <footer style="display:flex;justify-content:flex-end;gap:8px;margin-top:14px">
    <button id="cancel-song-edit-btn">Cancel</button>
    <button id="confirm-online-song-btn" class="pri">Add to service</button>
  </footer>`;
  $('#edit-song-modal').hidden=false;
}

async function confirmOnlineSongAdd(){
  if(!pendingOnlineSong)return;
  const title=($('#edit-song-title-input')?.value||'').trim()||pendingOnlineSong.resultTitle||'Untitled Song';
  const lyrics=($('#edit-song-lyrics-input')?.value||'').trim();
  const lps=parseInt($('#edit-song-lines-select')?.value||'4',10);
  if(!lyrics)return toast('Please enter some lyrics.');
  const parsed=parse(lyrics,lps);
  const prev=pendingOnlineSong.song.credits||{};
  const credits={
    ...prev,
    artist:($('#edit-song-artist')?.value||'').trim()||prev.artist||'',
    authors:($('#edit-song-artist')?.value||'').trim()||prev.authors||prev.author||'',
    writers:($('#edit-song-writers')?.value||'').trim()||prev.writers||'',
    year:($('#edit-song-year')?.value||'').trim()||prev.year||'',
    ccli:($('#edit-song-ccli')?.value||'').trim()||prev.ccli||prev.CCLI||'',
  };
  const song={t:title,...parsed,credits,lps};
  if(!song.ord.length)return toast('No sections found in lyrics.');
  const existing=LIB.findIndex(x=>x.t.trim().toLowerCase()===song.t.trim().toLowerCase());
  if(existing>=0){
    const replace=await appConfirm('A song named "'+song.t+'" is already in your Song Library.\n\nReplace the existing song with this version?',{title:'Replace song',confirmLabel:'Replace'});
    if(!replace){toast('Song was not added.');return}
    LIB[existing]=song;
  }else LIB.push(song);
  saveState();
  pendingOnlineSong=null;
  closeEditSongModal();
  closeLibraryModal();
  addSong(song,'Added '+title+' to the service and library.');
}

function updateLogoPreview(){
  const preview=$('#settings-logo-preview'),full=$('#logo-full-preview'),removeBtn=$('#remove-logo-btn'),bgInput=$('#logo-bg-color'),sizeInput=$('#logo-size'),sizeLabel=$('#logo-size-value'),vInput=$('#logo-vertical-position'),hInput=$('#logo-horizontal-position');
  const render=el=>{if(!el)return;el.style.background=logoBgCss();el.style.display='flex';el.style.alignItems=churchLogoVPos;el.style.justifyContent=churchLogoHPos;el.style.overflow='hidden';el.innerHTML=logoMarkup()};
  render(preview);render(full);
  const configure=$('#open-logo-controls');if(configure)configure.hidden=!churchLogo;
  if(preview){preview.style.borderRadius='6px';preview.style.border='1px solid var(--ln)'}
  if(removeBtn)removeBtn.hidden=!churchLogo;
  if(bgInput&&bgInput.value!==logoBgCss())bgInput.value=logoBgCss();
  if(sizeInput)sizeInput.value=String(churchLogoSize);
  if(sizeLabel)sizeLabel.textContent=churchLogoSize+'%';
  if(vInput)vInput.value=churchLogoVPos;
  if(hInput)hInput.value=churchLogoHPos;
  document.querySelectorAll('[data-logo-vpos]').forEach(btn=>{const active=btn.dataset.logoVpos===churchLogoVPos;btn.classList.toggle('active',active);btn.setAttribute('aria-pressed',String(active))});
  document.querySelectorAll('[data-logo-hpos]').forEach(btn=>{const active=btn.dataset.logoHpos===churchLogoHPos;btn.classList.toggle('active',active);btn.setAttribute('aria-pressed',String(active))});
}
function saveLogoLayout(){
  try{localStorage.setItem('church-presenter-logo-size',String(churchLogoSize));localStorage.setItem('church-presenter-logo-vpos',churchLogoVPos);localStorage.setItem('church-presenter-logo-hpos',churchLogoHPos)}catch{}
  updateLogoPreview();rStage();try{pushAudienceClear()}catch{}
}

function uploadLogoFile(file){
  if(!file)return;
  const reader=new FileReader();
  reader.onload=()=>{
    churchLogo=reader.result;
    localStorage.setItem('church-presenter-logo',churchLogo);
    updateLogoPreview();
    rStage();
    toast('Church logo saved.');
  };
  reader.readAsDataURL(file);
}


/** Strip non-persistent blob/data URLs so localStorage stays small and reloads rehydrate from IndexedDB. */
function sanitizeSlidesForSave(slides){
  const out=cl(slides||[]);
  (out||[]).forEach(it=>{
    // Never persist present-only PPT placeholders
    if(it._pptBackup){
      it.parts=it._pptBackup.parts;
      it.ord=it._pptBackup.ord;
      if(it._pptBgBackup!=null)it.bg=it._pptBgBackup;
      delete it._pptBackup;delete it._pptBgBackup;
    }
    it.parts=(it.parts||[]).filter(p=>!p._pptPlaceholder);
    if(it._pptBackup)delete it._pptBackup;
    if(it.bgMedia&&typeof it.bgMedia==='object'){
      // keep id/kind/name only
      it.bgMedia={id:it.bgMedia.id,kind:it.bgMedia.kind,name:it.bgMedia.name};
    }
    // transient sourcePath is machine-specific — drop it; originalFileId is enough
    if(it.sourcePath)delete it.sourcePath;
    (it.parts||[]).forEach(p=>{
      if(p.bgMedia)p.bgMedia={id:p.bgMedia.id,kind:p.bgMedia.kind,name:p.bgMedia.name};
      if(p.videoMedia){
        p.videoMedia={id:p.videoMedia.id,kind:p.videoMedia.kind,name:p.videoMedia.name,settings:p.videoMedia.settings||{}};
      }
      (p.im||[]).forEach(im=>{
        if(im.mediaId){
          // never persist blob: or huge data: URLs
          if(!im.src||String(im.src).startsWith('blob:')||String(im.src).startsWith('data:')){
            im.src='';
          }
        }
      });
      (p.elements||[]).forEach(el=>{
        if(el.kind==='image'&&el.mediaId){
          if(!el.src||String(el.src).startsWith('blob:')||String(el.src).startsWith('data:'))el.src='';
        }
      });
    });
  });
  return out;
}

function saveState(){
  // Never persist settings-injected present blanks into saved services
  if(!live){
    try{(S||[]).forEach(stripPresentBlanks)}catch{}
  }
  if(!services.length){
    try{
      localStorage.setItem('service-slides-state',JSON.stringify({
        services:[],
        activeService:0,
        library:LIB,
        theme:$('#theme-setting')?.value||'system',
        display:selectedDisplay||$('#display-setting')?.value||'default'
      }));
    }catch(err){console.warn('saveState empty failed',err)}
    return;
  }
  if(activeService<0||activeService>=services.length)activeService=0;
  // Persist a sanitized copy (media lives in IndexedDB; only ids in localStorage)
  try{
    const clean=sanitizeSlidesForSave(S);
    if(!live){try{clean.forEach(stripPresentBlanks)}catch{}}
    services[activeService].slides=clean;
    touchServiceEdited();
  }catch(err){
    console.warn('sanitize save failed, fallback clone',err);
    services[activeService].slides=cl(S);
    touchServiceEdited();
  }
  try{
    localStorage.setItem('service-slides-state',JSON.stringify({
      services,
      activeService,
      library:LIB,
      theme:$('#theme-setting')?.value||'system',
      display:selectedDisplay||$('#display-setting')?.value||'default'
    }));
  }catch(err){
    console.warn('saveState localStorage failed (quota?)',err);
    // Retry without library songs if quota exceeded
    try{
      localStorage.setItem('service-slides-state',JSON.stringify({
        services,
        activeService,
        library:[],
        theme:$('#theme-setting')?.value||'system',
        display:selectedDisplay||$('#display-setting')?.value||'default'
      }));
    }catch(err2){console.error('saveState retry failed',err2)}
  }
  try{saveMediaMeta()}catch{}
}

function rAll(){
  ed=false;
  if(!S.length){A=null;sel=0;si=0;aj=0}
  else if(!A||!S.includes(A)){A=S[Math.min(sel,S.length-1)]||S[0];aj=0;sel=Math.max(0,S.indexOf(A))}
  if(A){
    aj=Math.max(0,Math.min(aj,(A?.ord?.length||1)-1));
    si=Math.max(0,Math.min(si,(S[sel]?.ord?.length||1)-1));
  }else{aj=0;si=0}
  try{syncEffectiveLogo()}catch{}
  const _svcTitle=services[activeService]||{name:'Service'};
  $('.svc').textContent=(_svcTitle.name||'').trim();try{renderTitleBar(_svcTitle)}catch{}
  const penEl=document.querySelector('#service-title-edit .pen');
  if(penEl&&_svcTitle.lastEdited){
    penEl.setAttribute('title',formatLastEdited(_svcTitle.lastEdited));
  }
  rStrip();
  rStage();
  rPanel();
  rOutline();
  try{syncSpeakerNotesUI()}catch{}
  if(live&&document.body.classList.contains('presenter-mode')){
    renderPresenterView();
    pushAudienceStage();
  }else if(live){
    pushAudienceStage();
  }
  saveState();
}

let _lastGoAt=0;
function go(d){
  // Prevent double-advance from capture+bubble key handlers (or click+key)
  const now=Date.now();
  if(now-_lastGoAt<150)return;
  _lastGoAt=now;
  imageSel=-1;
  if(!S.length||!A)return;
  // Ignore navigation events while Present is still opening the audience display.
  // Startup must remain on the first service item and first slide.
  if(live&&presentStarting)return;
  // While PPT hand-off is actively waiting, still allow ending the presentation
  if(pptHandoffActive&&d!==0){
    // allow user to force-end with further navigation past end only via endPresent
  }
  if(pptHandoffActive&&d>0){
    // cancel stuck hand-off so we are not frozen
    pptHandoffActive=false;
    pptHandoffKey='';
  }
  const a=S.flatMap(it=>it.ord.map((_,j)=>[it,j]));
  if(!a.length)return;
  const k=a.findIndex(x=>x[0]==A&&x[1]==aj);
  if(k<0)return;
  const target=k+d;

  // Past the last slide while presenting → end immediately
  if(d>0&&target>=a.length){
    if(live){
      endPresent();
      return;
    }
    return;
  }

  // Going back from end-black restores the last slide
  if(d<0&&live&&clr==='black'&&k===a.length-1){
    clr=null;
    rStage();
    pushAudienceClear();
    if(document.body.classList.contains('presenter-mode'))renderPresenterView();
    return;
  }

  if(target<0||target>=a.length)return;
  const n=a[target];
  if(n){
    const hadClear=!!clr;
    const _prevA=A;
    A=n[0];aj=n[1];sel=S.indexOf(A);si=aj;clr=null;
    if(live&&_prevA&&_prevA!==A)pptLeaveItem(_prevA);
    if(live){
      // Fast path: push audience immediately, light UI updates only
      pushAudienceStage();
      if(hadClear)pushAudienceClear();
      // Update hidden stage + presenter chrome without full rAll/saveState
      try{
        if(document.body.classList.contains('presenter-mode')){
          renderPresenterView();
        }else{
          rStage();
        }
      }catch(err){console.error(err)}
      try{updateStripSelection()}catch{}
      try{
        const outline=$('#outline');
        if(outline)outline.querySelectorAll('.oi').forEach((el,i)=>el.classList.toggle('on',i===sel));
        const pvOutline=$('#pv-outline');
        if(pvOutline)pvOutline.querySelectorAll('.pv-oi').forEach((el,i)=>el.classList.toggle('on',i===sel));
      }catch{}
      try{openNativePowerpointIfNeeded(A)}catch{}
      // Do not saveState on every advance — it stalls the UI
    }else{
      rAll();
      setTimeout(scrollCurrentSlideIntoView,0);
    }
  }
}

function isTypingTarget(el){
  if(!el||!(el instanceof Element))return false;
  if(el.closest('[contenteditable="true"],[contenteditable="plaintext-only"],[contenteditable=""]'))return true;
  const tag=(el.tagName||'').toLowerCase();
  if(tag==='input'||tag==='textarea'||tag==='select')return true;
  if(el.isContentEditable)return true;
  return false;
}
/** True when any modal/overlay menu is open — editor hotkeys must not run. */
function isEditorOverlayOpen(){
  try{
    const ids=['media-modal','library-modal','edit-song-modal','struct-modal','team-modal','share-modal','import-modal','close-confirm-modal'];
    for(const id of ids){
      const m=document.getElementById(id);
      if(m&&!m.hidden)return true;
    }
    const settings=document.getElementById('settings');
    if(settings&&settings.classList.contains('open'))return true;
    const songSearch=document.getElementById('song-search-panel');
    if(songSearch&&!songSearch.hidden)return true;
    // Generic: any visible .app-modal
    const openModal=document.querySelector('.app-modal:not([hidden])');
    if(openModal)return true;
  }catch{}
  return false;
}

function scrollElementIntoParent(el,parent,{behavior='smooth',pad=12}={}){
  if(!el||!parent)return;
  const a=parent.getBoundingClientRect(),b=el.getBoundingClientRect();
  // Only scroll when the element is outside the visible area (avoids clunky jump on forward)
  const fullyVisible=b.top>=a.top+pad&&b.bottom<=a.bottom-pad;
  if(fullyVisible)return;
  let nextTop=parent.scrollTop;
  if(b.top<a.top+pad)nextTop=parent.scrollTop+(b.top-a.top)-pad;
  else if(b.bottom>a.bottom-pad)nextTop=parent.scrollTop+(b.bottom-a.bottom)+pad;
  parent.scrollTo({top:Math.max(0,nextTop),behavior});
}


/** Scroll the strip so the section at index i sits near the top of the column. */
function scrollNewItemToTop(i){
  const st=$('#strip');
  if(!st)return;
  const el=st.querySelector('[data-dh="'+i+'"]')||st.querySelector('.thumb-wrap[data-i="'+i+'"]');
  if(!el)return;
  const a=st.getBoundingClientRect(),b=el.getBoundingClientRect();
  const target=Math.max(0,b.top-a.top+st.scrollTop-8);
  st.scrollTo({top:target,left:0,behavior:'smooth'});
  const outline=$('#outline');
  if(outline){
    const oi=outline.querySelector('[data-oi="'+i+'"]');
    if(oi){
      const oa=outline.getBoundingClientRect(),ob=oi.getBoundingClientRect();
      outline.scrollTo({top:Math.max(0,ob.top-oa.top+outline.scrollTop-8),behavior:'smooth'});
    }
  }
}

/* ===== Undo / Redo (editor mutations) ===== */
const UNDO_LIMIT=40;
let _undoStack=[],_redoStack=[],_undoQuiet=false;
function snapshotEditor(){
  try{
    return {
      S:cl(S),
      sel,si,aj,
      AIndex:Math.max(0,S.indexOf(A)),
      elSel:elSel||null
    };
  }catch(err){console.warn('snapshotEditor',err);return null}
}
function applySnapshot(snap){
  if(!snap)return;
  _undoQuiet=true;
  try{
    S=cl(snap.S)||[];
    sel=Math.min(Math.max(0,snap.sel||0),Math.max(0,S.length-1));
    si=snap.si||0;aj=snap.aj||0;
    A=S[snap.AIndex]||S[sel]||S[0]||null;
    if(A){
      aj=Math.max(0,Math.min(aj,(A.ord||[]).length-1));
      si=Math.max(0,Math.min(si,(A.ord||[]).length-1));
    }
    elSel=snap.elSel||null;
    if(services[activeService])services[activeService].slides=sanitizeSlidesForSave(S);
    rAll();
  }finally{_undoQuiet=false}
}
function pushUndo(){
  if(_undoQuiet||live)return;
  const snap=snapshotEditor();
  if(!snap)return;
  _undoStack.push(snap);
  if(_undoStack.length>UNDO_LIMIT)_undoStack.shift();
  _redoStack.length=0;
}
function undoEditor(){
  if(live||!_undoStack.length){toast(_undoStack.length?'':'Nothing to undo');return}
  const cur=snapshotEditor();
  const prev=_undoStack.pop();
  if(cur)_redoStack.push(cur);
  applySnapshot(prev);
  toast('Undo');
}
function redoEditor(){
  if(live||!_redoStack.length){toast(_redoStack.length?'':'Nothing to redo');return}
  const cur=snapshotEditor();
  const next=_redoStack.pop();
  if(cur)_undoStack.push(cur);
  applySnapshot(next);
  toast('Redo');
}
function scrollCurrentSlideIntoView(){
  const strip=$('#strip');
  if(strip){
    const thumb=strip.querySelector('.th.on')||strip.querySelector(`[data-i="${sel}"][data-j="${aj}"]`)||strip.querySelector(`[data-dh="${sel}"]`);
    if(thumb)scrollElementIntoParent(thumb,strip,{behavior:'smooth',pad:16});
  }
  const outline=$('#outline');
  if(outline){
    const oi=outline.querySelector(`.oi.on`)||outline.querySelector(`[data-oi="${sel}"]`);
    if(oi)scrollElementIntoParent(oi,outline,{behavior:'smooth',pad:12});
  }
}

function toast(m){
  const t=$('#toast');t.textContent=m;t.classList.add('on');
  clearTimeout(tt);tt=setTimeout(()=>t.classList.remove('on'),3200);
}

function mvOrd(it,f,t){
  const[x]=it.ord.splice(f,1);
  it.ord.splice(t,0,x);
  if(it===A){if(aj==f)aj=t;else if(f<aj&&t>=aj)aj--;else if(f>aj&&t<=aj)aj++}
}
function insOrd(it,p,id){it.ord.splice(p,0,id);if(it===A&&p<=aj)aj++}
function rmOrd(it,r){it.ord.splice(r,1);if(it===A){if(r<aj)aj--;aj=Math.min(aj,it.ord.length-1)}}
/** Move a slide within an item, or across section boundaries when to is out of range. */
function moveSlideInItem(itemIndex,from,to){
  const it=S[itemIndex];if(!it)return;
  if(from<0||from>=it.ord.length)return;
  if(from===to)return;
  pushUndo();

  // Within the same section
  if(to>=0&&to<it.ord.length){
    mvOrd(it,from,to);
    if(itemIndex===sel){si=to;if(it===A)aj=to}
    rAll();
    return;
  }

  if(isSongItem(it)){
    toast('Use Edit to rearrange song sections.');
    return;
  }

  const partId=it.ord[from];
  const src=it.parts.find(p=>p.id===partId);if(!src)return;
  const shared=it.ord.filter(id=>id===partId).length>1;
  const moving=shared?clonePart(src):src;

  // Detach from source
  it.ord.splice(from,1);
  if(!shared)it.parts=it.parts.filter(p=>p.id!==partId);

  const direction=to>=it.ord.length?'down':'up';

  if(direction==='down'){
    // Prefer next freeform section; skip songs; else create new section after
    let destIdx=itemIndex+1;
    while(destIdx<S.length&&isSongItem(S[destIdx]))destIdx++;
    if(destIdx<S.length&&!isSongItem(S[destIdx])){
      const dest=S[destIdx];
      if(!dest.parts.includes(moving)&&!dest.parts.find(p=>p.id===moving.id))dest.parts.push(moving);
      dest.ord.unshift(moving.id);
    }else{
      const fresh=mkS(it.t||'Slide',[''],0,it.f?.fs||5.2);
      fresh.parts=[moving];
      fresh.ord=[moving.id];
      fresh.bg=it.bg;fresh.dim=it.dim;fresh.f=cl(it.f||{});
      if(it.bgMedia)fresh.bgMedia=cl(it.bgMedia);
      S.splice(itemIndex+1,0,fresh);
    }
  }else{
    // Prefer previous freeform section; skip songs; else create new section before
    let destIdx=itemIndex-1;
    while(destIdx>=0&&isSongItem(S[destIdx]))destIdx--;
    if(destIdx>=0&&!isSongItem(S[destIdx])){
      const dest=S[destIdx];
      if(!dest.parts.find(p=>p.id===moving.id))dest.parts.push(moving);
      dest.ord.push(moving.id);
    }else{
      const fresh=mkS(it.t||'Slide',[''],0,it.f?.fs||5.2);
      fresh.parts=[moving];
      fresh.ord=[moving.id];
      fresh.bg=it.bg;fresh.dim=it.dim;fresh.f=cl(it.f||{});
      if(it.bgMedia)fresh.bgMedia=cl(it.bgMedia);
      S.splice(itemIndex,0,fresh);
      itemIndex++; // source shifted right
    }
  }

  // Drop empty source section
  if(!it.ord.length){
    const emptyIdx=S.indexOf(it);
    if(emptyIdx>=0)S.splice(emptyIdx,1);
  }

  // Focus the moved slide
  let found=false;
  for(let i=0;i<S.length;i++){
    const j=S[i].ord.indexOf(moving.id);
    if(j>=0){sel=i;si=j;A=S[i];aj=j;found=true;break}
  }
  if(!found){sel=Math.min(sel,S.length-1);A=S[sel];aj=0;si=0}
  elSel=null;stripFocus='slide';
  rAll();
  setTimeout(scrollCurrentSlideIntoView,0);
}

/** Duplicate ord entry at index (same section id → same name, like structure builder). */
function clonePart(src){
  const copy=cl(src);copy.id=uid();
  if(Array.isArray(copy.elements))copy.elements=copy.elements.map(el=>{const e2=cl(el);e2.id=uid();return e2});
  return copy;
}
function uniqueifySlideParts(it){
  if(!it||isSongItem(it))return;
  const seen=new Set();
  for(let i=0;i<it.ord.length;i++){
    const id=it.ord[i];
    if(!seen.has(id)){seen.add(id);continue}
    const src=it.parts.find(p=>p.id===id);if(!src)continue;
    const copy=clonePart(src);it.parts.push(copy);it.ord[i]=copy.id;
  }
}
function slideClipFor(i,j){
  const it=S[i];
  if(!it||j<0||j>=it.ord.length)return null;
  if(isSongItem(it)){
    const p=it.parts.find(x=>x.id===it.ord[j]);
    return {kind:'song-ord',id:it.ord[j],name:(p&&p.n)||'Section'};
  }
  const src=it.parts.find(p=>p.id===it.ord[j]);
  return src?{kind:'part',part:clonePart(src)}:null;
}
function writeSlideClipToClipboard(){
  try{
    if(!slideClip)return;
    const txt='<!--church-presenter-slide-->'+JSON.stringify(slideClip);
    if(navigator.clipboard&&navigator.clipboard.writeText){
      navigator.clipboard.writeText(txt).catch(()=>{});
    }
  }catch(err){}
}
function copySlideAt(i,j,quiet){
  const c=slideClipFor(i,j);
  if(!c)return false;
  slideClip=c;
  writeSlideClipToClipboard();
  if(!quiet)toast('Slide copied');
  return true;
}
function cutSlideAt(i,j){
  if(live)return;
  if(!copySlideAt(i,j,true))return;
  removeSlideInItem(i,j);
  toast('Slide cut');
}
function pasteSlideAfter(i,j){
  if(live)return;
  if(!slideClip){toast('Nothing copied yet.');return}
  const it=S[i];if(!it)return;
  pushUndo();
  if(slideClip.kind==='song-ord'){
    if(!isSongItem(it)){toast('Song sections can only be pasted into a song.');return}
    insOrd(it,j+1,slideClip.id);
  }else{
    if(isSongItem(it)){toast('Use Edit Song to manage song sections.');return}
    const copy=clonePart(slideClip.part);
    it.parts.push(copy);
    insOrd(it,j+1,copy.id);
  }
  sel=i;A=it;si=j+1;aj=j+1;
  rAll();
  setTimeout(scrollCurrentSlideIntoView,0);
  toast('Slide pasted');
}

function duplicateSlideInItem(itemIndex,j){
  const it=S[itemIndex];if(!it)return;
  if(j<0||j>=it.ord.length)return;
  const srcId=it.ord[j];
  if(isSongItem(it)){insOrd(it,j+1,srcId)}
  else{
    const src=it.parts.find(p=>p.id===srcId);if(!src)return;
    const copy=clonePart(src);it.parts.push(copy);insOrd(it,j+1,copy.id);
  }
  if(itemIndex===sel){si=j+1;if(it===A)aj=j+1}
  rAll();
}

/** Copy the currently selected strip slide (or multi-selected slides). */
function copySelectedSlide(){
  if(live)return;
  const items=multiParts();
  if(items.length>1){
    window.__multiSlideClip=items.map(x=>clonePart(x.part));
    slideClip={kind:'part',part:clonePart(items[0].part)};
    writeSlideClipToClipboard();
    toast(items.length+' slides copied');
    return;
  }
  const it=A;if(!it)return;
  const j=aj;
  if(j<0||j>=it.ord.length)return;
  if(isSongItem(it)){
    slideClip={kind:'song-ord',id:it.ord[j],name:(sls(it)[j]&&sls(it)[j].n)||'Section'};
  }else{
    const src=it.parts.find(p=>p.id===it.ord[j]);if(!src)return;
    slideClip={kind:'part',part:clonePart(src)};
  }
  writeSlideClipToClipboard();
  toast('Slide copied');
}
/** Paste the copied slide directly below the current selection (or multi-slide clip). */
function pasteSlideBelow(){
  if(live)return;
  const clips=window.__multiSlideClip;
  if(clips&&clips.length>1){
    const it=A;if(!it||isSongItem(it)){toast('Paste into a normal section.');return}
    pushUndo();
    let at=aj+1;
    clips.forEach(src=>{
      const copy=clonePart(src);
      copy.id=uid();
      it.parts.push(copy);
      insOrd(it,at,copy.id);
      at++;
    });
    aj=at-1;si=aj;rAll();
    setTimeout(scrollCurrentSlideIntoView,0);
    toast(clips.length+' slides pasted');
    return;
  }
  if(!slideClip)return;
  const it=A;if(!it)return;
  const j=aj;
  if(slideClip.kind==='song-ord'){
    if(!isSongItem(it)){toast('Song sections can only be pasted into a song.');return}
    pushUndo();
    insOrd(it,j+1,slideClip.id);
    si=j+1;aj=j+1;
  }else if(slideClip.kind==='part'){
    if(isSongItem(it)){toast('Use Edit Song to manage song sections.');return}
    pushUndo();
    const copy=clonePart(slideClip.part);
    it.parts.push(copy);
    insOrd(it,j+1,copy.id);
    si=j+1;aj=j+1;
  }else return;
  rAll();
  setTimeout(scrollCurrentSlideIntoView,0);
  toast('Slide pasted');
}
function deleteSelectedSlide(){
  if(live)return;
  const items=multiParts();
  if(items.length>1){
    pushUndo();
    items.slice().reverse().forEach(x=>removeSlideInItem(x.i,x.j));
    MS.slides.clear();
    rAll();
    toast(items.length+' slides deleted');
    return;
  }
  pushUndo();
  const it=A;if(!it)return;
  removeSlideInItem(S.indexOf(it),aj);
}
function addNewSlideInItem(itemIndex,afterJ){
  const it=S[itemIndex];if(!it)return;
  if(isSongItem(it)){toast('Use Edit Song to add song sections.');return}
  pushUndo();
  const f={...(it.f||{fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'})};
  const prevId=(afterJ!=null&&afterJ>=0&&it.ord.length)?it.ord[Math.min(afterJ,it.ord.length-1)]:null;
  const prev=prevId?it.parts.find(p=>p.id===prevId):null;
  const part={id:uid(),n:'Slide '+(it.parts.length+1),tx:'',elements:[
    {id:uid(),kind:'text',tx:'Type your text here',x:8,y:18,w:84,h:50,rot:0,z:1,f:cl(f)}
  ]};
  if(prev&&prev.ownBg){
    part.ownBg=true;part.bg=prev.bg;part.dim=prev.dim;
    if(prev.customBg)part.customBg=prev.customBg;
    if(prev.bgMedia)part.bgMedia=cl(prev.bgMedia);
  }
  it.parts.push(part);
  const at=afterJ==null?it.ord.length:Math.min(it.ord.length,Math.max(0,afterJ+1));
  insOrd(it,at,part.id);
  sel=itemIndex;si=at;A=it;aj=at;elSel=part.elements[0].id;tab='Structure';
  rAll();setTimeout(()=>{scrollNewItemToTop(itemIndex);scrollCurrentSlideIntoView()},0);toast('Added new slide.');
}

function removeServiceItem(itemIndex,{silent=false}={}){if(!silent)pushUndo();
  if(itemIndex<0||itemIndex>=S.length)return;
  const name=S[itemIndex]?.t||'item';
  S.splice(itemIndex,1);
  elSel=null;stripFocus='item';
  if(!S.length){
    sel=0;si=0;A=null;aj=0;
  }else{
    sel=Math.min(Math.max(0,itemIndex),S.length-1);
    si=0;A=S[sel];aj=0;
  }
  rAll();
}
function removeSlideInItem(itemIndex,j){pushUndo();
  const it=S[itemIndex];if(!it)return;
  // Last slide in the section → remove the whole section
  if(it.ord.length<=1){
    removeServiceItem(itemIndex);
    return;
  }
  rmOrd(it,j);
  if(itemIndex===sel){si=Math.min(si,it.ord.length-1);if(it===A)aj=Math.min(aj,it.ord.length-1)}
  rAll();
}

function formatSongCredits(meta,title){
  if(!meta&&!title)return '';
  const m=meta||{};
  const bits=[];
  const selected=songSettings.creditFields||{title:true,artist:true,writers:true,year:true,ccli:true};
  const push=(label,val)=>{
    if(val==null)return;
    let s=Array.isArray(val)?val.filter(Boolean).join(', '):String(val).trim();
    if(!s)return;
    // Avoid duplicates (case-insensitive)
    const low=s.toLowerCase();
    if(bits.some(b=>b.toLowerCase()===low||b.toLowerCase().includes(low)||low.includes(b.toLowerCase())))return;
    bits.push(label?`${label} ${s}`:s);
  };
  // Title
  if(selected.title&&title)push('',title);
  // People — keep full strings (not first name only)
  const artists=[m.artist,m.author,m.artists];
  if(selected.artist)artists.forEach(p=>{
    if(!p)return;
    if(Array.isArray(p))push('',p.join(', '));
    else push('',p);
  });
  if(selected.writers)push('Writers',m.writers||m.writer||m.composers||m.composer||m.lyricists||m.lyricist||m.songwriters||m.songwriter||m.authors||'');
  // Dates — always try to surface a published year
  let yearRaw=m.year||m.releaseDate||m.released||m.published||m.publicationDate||m.date||m.copyright||'';
  if(!yearRaw){
    const blob=Object.values(m).filter(v=>v!=null&&(typeof v==='string'||typeof v==='number')).join(' ');
    const ym=String(blob).match(/\b((?:19|20)\d{2})\b/);
    if(ym)yearRaw=ym[1];
  }
  if(selected.year&&yearRaw){
    const y=String(yearRaw).trim();
    const ym=y.match(/\b((?:19|20)\d{2})\b/);
    const year=ym?ym[1]:y;
    if(year&&!bits.some(b=>String(b).includes(String(year)))){
      bits.push('Published '+year);
    }
  }
  // Album / CCLI / source when useful
  if(selected.ccli&&(m.ccli||m.CCLI))push('CCLI',m.ccli||m.CCLI);
  return bits.filter(Boolean).join('  ·  ');
}

/** Auto-inserted blank slides are tagged so we can add/remove them when the setting changes. */
/** Settings-injected blank slides (present-only). Not the same as user-made blank slides. */
function isAutoBlankPart(p){
  if(!p)return false;
  // Only explicit present-mode blanks. Never treat real song sections as blanks
  // even if they are named "Blank" or have empty lyrics — that was deleting lyric slides.
  if(!(p.autoBlank||p.presentBlank))return false;
  // Safety: never treat a part with real lyrics as an auto-blank
  if(String(p.tx||'').trim())return false;
  return true;
}
function makePresentBlankPart(){
  return {id:uid(),n:'Blank',tx:'',autoBlank:true,presentBlank:true};
}

/** Remove all settings-injected blank slides from a song (editor should never show them). */
function stripPresentBlanks(item){
  if(!item||!isSongItem(item)||!Array.isArray(item.ord))return item;
  const drop=new Set();
  (item.parts||[]).forEach(p=>{
    if(!isAutoBlankPart(p))return;
    // Hard guard: never remove a part that still has lyric text
    if(String(p.tx||'').trim())return;
    drop.add(p.id);
  });
  if(!drop.size)return item;
  item.ord=item.ord.filter(id=>!drop.has(id));
  item.parts=(item.parts||[]).filter(p=>!drop.has(p.id));
  return item;
}

/** Inject settings blank slides for live presentation only. */
function injectPresentBlanks(){
  const blank=songSettings.blankSlide||'none';
  (S||[]).forEach(item=>{
    if(!isSongItem(item))return;
    stripPresentBlanks(item);
    if(blank==='none')return;
    item.parts=item.parts||[];
    if(blank==='start'||blank==='both'){
      const p=makePresentBlankPart();
      item.parts.push(p);
      item.ord.unshift(p.id);
    }
    if(blank==='end'||blank==='both'){
      const p=makePresentBlankPart();
      item.parts.push(p);
      item.ord.push(p.id);
    }
  });
}

/** Remove present blanks from every song in the active service (and keep selection valid). */
function removePresentBlanksFromService(){
  (S||[]).forEach(stripPresentBlanks);
  if(A&&A.ord&&A.ord.length){
    aj=Math.max(0,Math.min(aj,A.ord.length-1));
    si=Math.max(0,Math.min(si,(S[sel]&&S[sel].ord?S[sel].ord.length:1)-1));
  }else if(S.length){
    A=S[Math.min(sel,S.length-1)]||S[0];
    aj=0;si=0;
  }
}

/** While presenting a Microsoft PowerPoint hand-off section, collapse its slides
 *  to a single black placeholder (strip + presenter + audience). Restored on End. */
function injectPptPresentPlaceholders(){
  (S||[]).forEach(it=>{
    if(typeof isPptHandoffItem!=='function'||!isPptHandoffItem(it))return;
    if(it._pptBackup)return;
    it._pptBackup={parts:cl(it.parts||[]),ord:cl(it.ord||[])};
    const black={id:uid(),n:'PowerPoint',tx:'',_pptPlaceholder:true};
    it.parts=[black];
    it.ord=[black.id];
    it._pptBgBackup=it.bg;
    it.bg=3;
  });
  if(A&&A.ord&&A.ord.length){
    aj=Math.max(0,Math.min(aj,A.ord.length-1));
    si=aj;
  }
}
function restorePptPresentPlaceholders(){
  (S||[]).forEach(it=>{
    if(!it||!it._pptBackup)return;
    it.parts=it._pptBackup.parts;
    it.ord=it._pptBackup.ord;
    if(it._pptBgBackup!=null){it.bg=it._pptBgBackup;delete it._pptBgBackup}
    delete it._pptBackup;
  });
  if(A&&A.ord&&A.ord.length){
    aj=Math.max(0,Math.min(aj,A.ord.length-1));
    si=aj;
  }else if(S.length){
    A=S[Math.min(sel,S.length-1)]||S[0];
    aj=0;si=0;
  }
}

/** Strip any leftover settings blanks from all services (editor cleanup). */
function stripAllPresentBlanksEverywhere(){
  (S||[]).forEach(stripPresentBlanks);
  (services||[]).forEach(svc=>{
    if(Array.isArray(svc.slides))svc.slides.forEach(stripPresentBlanks);
    if(svc===services[activeService])svc.slides=S;
  });
}

/** Apply song credits according to Song settings. Blanks are present-only — not applied here. */
function applySongSettingsToItem(item,sourceSong){
  if(!item||!isSongItem(item))return item;
  // Ensure editor never carries present blanks on add
  stripPresentBlanks(item);
  // Clear any per-part credits (shared sections caused duplicates)
  (item.parts||[]).forEach(p=>{delete p.credits;delete p.creditsPos});
  if(songSettings.credits){
    const meta=(sourceSong&&sourceSong.credits)||item.credits||{};
    const line=formatSongCredits(meta,item.t);
    if(line){
      item.credits=meta;
      item.creditLine=line;
      item.creditsPos=songSettings.creditsPos||'middle';
    }
  }else{
    delete item.creditLine;
  }
  return item;
}

function addSong(x,m){
  pushUndo();
  const item=fromLib(x,0);
  if(x&&x.credits)item.songCredits=cl(x.credits);
  applySongSettingsToItem(item,x);
  try{refitSong(item)}catch(err){console.warn('fit',err)}
  const at=Math.max(0,Math.min(pendingSectionInsert!=null?pendingSectionInsert:sel+1,S.length));
  pendingSectionInsert=null;
  S.splice(at,0,item);
  sel=at;si=0;A=item;aj=0;tab='Style';
  rAll();
  setTimeout(()=>scrollNewItemToTop(sel),0);
  toast(m||'Song added with all slides.');
}
function addBlank(n){pushUndo();const it=mkS(n,['Type your text here'],0,6.5);S.splice(sel+1,0,it);sel++;si=0;A=it;aj=0;tab='Structure';rAll();setTimeout(()=>scrollNewItemToTop(sel),0);toast('Added slide segment.')}

document.addEventListener('click',ev=>{
  const pickRow=ev.target.closest('[data-pick-el]');
  if(pickRow&&!ev.target.closest('[data-del-el],[data-dup-el],[data-el-drag]')){selectEl(pickRow.dataset.pickEl);return}
  const t=ev.target.closest('button');if(!t)return;
  const d=t.dataset,it=S[sel];

  if(t.id==='open-logo-controls'){
    updateLogoPreview();$('#logo-controls-modal').hidden=false;
  }
  else if(t.id==='logo-controls-close'||t.id==='logo-controls-close-bottom')$('#logo-controls-modal').hidden=true;
  else if(t.dataset.logoVpos){churchLogoVPos=t.dataset.logoVpos;saveLogoLayout()}
  else if(t.dataset.logoHpos){churchLogoHPos=t.dataset.logoHpos;saveLogoLayout()}
  else if(t.id=='upload-logo-settings-btn'){
    $('#church-logo-file').click();
  }
  else if(t.id=='remove-logo-btn'){removeSelectedLogo();return}
  else if(t.id=='__legacy-remove-logo'){
    churchLogo='';
    localStorage.removeItem('church-presenter-logo');
    updateLogoPreview();
    rStage();
    toast('Church logo removed.');
  }
  else if(t.id=='reformat-lines-btn'){
    reformatEditedSongLines();
  }
  else if(d.editLib!=null){
    openEditSongModal(+d.editLib);
  }
  else if(t.id=='edit-song-close'||t.id=='cancel-song-edit-btn'||t.id=='cancel-song-edit-top'){
    pendingOnlineSong=null;
    closeEditSongModal();
  }
  else if(t.id=='confirm-online-song-btn'){
    confirmOnlineSongAdd();
  }
  else if(t.id=='save-song-edit-btn'||t.id=='save-song-edit-top'){
    saveEditedSong();
  }
  else if(t.id=='delete-song-lib-btn'||t.id=='delete-song-lib-top'){
    deleteSongFromLib();
  }
  else if(d.openService!=null){
    saveState();activeService=+d.openService;S=cl(services[activeService].slides);sel=0;si=0;A=S[0];aj=0;document.body.classList.remove('home-mode');rAll();
  }
  else if(t.id=='services-btn'){saveState();homeShown=5;document.body.classList.add('home-mode');rHome()}
  else if(t.id=='settings-btn'||t.id==='home-settings-btn'){
    if(elSel){elSel=null;rStage()}
    $('#settings').classList.add('open');
    updateLogoPreview();
    const blank=$('#song-blank-slide');if(blank)blank.value=songSettings.blankSlide||'none';
    const cred=$('#song-credits-enabled');if(cred)cred.checked=!!songSettings.credits;
    const fields=$('#song-credit-fields');if(fields)fields.hidden=!songSettings.credits;
    const defaults={title:true,artist:true,writers:true,year:true,ccli:true},configured={...defaults,...(songSettings.creditFields||{})};
    Object.keys(defaults).forEach(k=>{const el=$('#credit-'+k);if(el)el.checked=configured[k]!==false});
    const pos=$('#song-credits-pos');if(pos)pos.value=songSettings.creditsPos||'middle';
    const wrap=$('#song-credits-pos-wrap');if(wrap)wrap.hidden=!songSettings.credits;syncCreditsTranspUi();
  }
  else if(t.id=='close-settings')closeSettingsAndRefresh()
  else if(t.id=='library-open'||t.id=='add-song-from-panel')openLibraryModal()
  else if(t.id=='library-close')closeLibraryModal()
  else if(t.id=='song-search-open'){$('#song-search-panel').hidden=false;$('#song-search-query').focus()}
  else if(t.id=='song-search-close'){$('#song-search-panel').hidden=true}
  else if(d.searchAdd!=null)addSearchResult(+d.searchAdd)
  else if(d.add!=null){closeLibraryModal();addSong(LIB[+d.add])}
  else if(d.c){if(!live&&!document.body.classList.contains('presenter-mode')&&!document.body.classList.contains('presentation-mode'))return;clr=clr==d.c?null:d.c;rStage();if(live)pushAudienceClear()}
  else if(d.bg!=null){
    const idx=+d.bg;
    bgApply(o=>{o.bg=idx;delete o.bgMedia;if(BG[idx]&&BG[idx].pick&&!o.customBg)o.customBg='#16213e'});
  }
  else if(d.al){const el=(!isSongItem(it)&&getSelectedEl())||null;const target=(el&&el.kind==='text')?(el.f||(el.f={fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'})):(it.f||(it.f={fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'}));target.al=d.al;if(el){saveSoon();rStage();rPanel()}else rAll()}
  else if(d.va){const el=(!isSongItem(it)&&getSelectedEl())||null;const target=(el&&el.kind==='text')?(el.f||(el.f={fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'})):(it.f||(it.f={fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'}));target.va=d.va;if(el){saveSoon();rStage();rPanel()}else rAll()}
  else if(d.tab){tab=d.tab;rPanel()}
  else if(d.mi){const n=sel+ +d.mi;if(n>=0&&n<S.length){[S[sel],S[n]]=[S[n],S[sel]];sel=n;rAll();setTimeout(()=>scrollServiceItemIntoView(n),0)}}
  else if(d.rmi){removeServiceItem(sel)}
  else if(d.renameItem!=null){renameServiceItem(+d.renameItem)}
  else if(d.newSlide!=null){
    const idx=+d.newSlide;
    if(S[idx]&&S[idx].k==='Song'){toast('Use Edit Song to add song sections.');return}
    addNewSlideInItem(idx,S[idx]?S[idx].ord.length-1:-1);
  }
  else if(d.sm!=null||d.ms!=null){
    const j=+(d.j!=null?d.j:0);
    const itemI=d.i!=null?+d.i:sel;
    const delta=+(d.sm!=null?d.sm:d.ms);
    const n=j+delta;
    moveSlideInItem(itemI,j,n);
  }
  else if(d.rm!=null){if(it.ord.length<2)return toast('Keep at least one section.');rmOrd(it,+d.rm);rAll()}
  else if(d.ap){insOrd(it,si+1,d.ap);si++;rAll()}
  else if(d.np){const c=it.parts.filter(p=>p.n.startsWith(d.np)).length,p={id:uid(),n:c?`${d.np} ${c+1}`:(d.np=='Verse'?'Verse 1':d.np),tx:''};it.parts.push(p);insOrd(it,si+1,p.id);si++;rAll()}
  else if(d.sv){const x={t:it.t,parts:cl(it.parts),ord:[...it.ord]},k=LIB.findIndex(y=>y.t==it.t);if(k<0)LIB.push(x);else LIB[k]=x;toast('Saved to song library.')}
  else if(d.hs!=null){imageSel=-1;elSel=null;stripFocus='item';sel=+d.hs;si=0;A=S[sel];aj=0;rAll()}
  else if(d.i!=null){imageSel=-1;elSel=null;stripFocus='slide';A=S[+d.i];aj=+d.j;sel=+d.i;si=aj;clr=null;rAll()}
  else if(d.es!=null){si=+d.es;rAll()}
  else if(t.id=='edb'){ed?stopEdit():startEdit()}
  else if(t.id=='pv')go(-1)
  else if(t.id=='nx')go(1)
  else if(t.id=='strip-add-slide'||t.id=='strip-add-slide-empty'){
    const part={id:uid(),n:'Slide 1',tx:'Insert text here'};
    const item=mk('Section 1','Slide',[part],[part.id],0,5.2);
    S.push(item);sel=S.length-1;si=0;A=item;aj=0;tab='Structure';rAll();
  }
  else if(t.id=='live'){if(typeof isPresentingUI==='function'?isPresentingUI():(live||document.body.classList.contains('presenter-mode')||document.body.classList.contains('presentation-mode')))endPresent();else{presentFrom=null;startPresent().catch(err=>console.error(err))}}
  else if(t.id=='add-text-box'){addTextBox()}
  else if(t.id=='add-canvas-image'){
    const inp=document.createElement('input');inp.type='file';inp.accept='image/*';
    inp.onchange=async()=>{
      const file=inp.files&&inp.files[0];if(!file)return;
      if(isSongItem(A)||live)return;
      const id='m'+Date.now().toString(36)+Math.random().toString(36).slice(2,8);
      await mPut(id,file);mediaURLs[id]=URL.createObjectURL(file);
      mediaMeta.push({id,name:file.name,kind:'image',size:file.size,createdAt:Date.now(),source:'canvas'});saveMediaMeta();
      const part=sls(A)[aj];if(!part)return;
      const els=ensureElements(part,A);
      const el={id:uid(),kind:'image',src:mediaURLs[id],mediaId:id,x:20,y:20,w:40,h:40,rot:0,z:(els.reduce((m,e)=>Math.max(m,e.z||1),0)+1)};
      els.push(el);elSel=el.id;saveSoon();rStage();rPanel();
    };
    inp.click();
  }
  else if(d.delEl){
    const part=sls(A)[aj];if(!part)return;
    const els=part.elements||[];
    const ix=els.findIndex(x=>x.id===d.delEl);
    if(ix>=0){els.splice(ix,1);if(elSel===d.delEl)elSel=null;saveSoon();rStage();rPanel()}
  }
  else if(d.dupEl){
    const part=sls(A)[aj];if(!part)return;
    const els=ensureElements(part,A)||[];
    const src=els.find(x=>x.id===d.dupEl);if(!src)return;
    const copy=cl(src);copy.id=uid();
    copy.x=Math.min(90,(src.x||0)+3);copy.y=Math.min(90,(src.y||0)+3);
    copy.z=(els.reduce((m,e)=>Math.max(m,e.z||1),0)+1);
    els.push(copy);elSel=copy.id;saveSoon();rStage();rPanel();
  }
  else if(t.id=='ns'){const s=$('#nl').value.trim();if(!s)return toast('Paste lyrics first.');const c={t:$('#nt').value.trim()||'New Song',...parse(s, 4),lps:4};LIB.push(c);closeLibraryModal();addSong(c,'Added to service and library.')}
  else if(t.id=='add-slide-to-item'){
    if(isSongItem(S[sel])){toast('Use Edit Song to add song sections.');return}
    addNewSlideInItem(sel,S[sel]?S[sel].ord.length-1:-1);
  }
  else if(t.id==='add-logo-slide')addChurchLogoSection();
  else if(d.qs)addBlank(d.qs)
  else if(t.id=='sadd')addBlank($('#sn').value.trim()||'New Slide Segment')
  else if(t.id=='send'){sendServiceTeamEmail()}
  else if(t.id=='add-band-member'){
    const team=ensureTeam(services[activeService]);
    team.band.push(emptyContact());setTimeout(openTeamModal,0);
    saveState();rPanel();
  }
  else if(t.closest&&t.closest('[data-remove-band]')){
    const btn=t.closest('[data-remove-band]');
    const idx=+btn.dataset.removeBand;
    const team=ensureTeam(services[activeService]);
    if(Number.isInteger(idx)&&idx>=0&&idx<team.band.length){
      team.band.splice(idx,1);setTimeout(openTeamModal,0);
      saveState();rPanel();
    }
  }
});

document.addEventListener('click',ev=>{if(ev.target&&ev.target.id==='logo-controls-modal')ev.target.hidden=true});
document.addEventListener('input',ev=>{
  const t=ev.target;
  if(t.id==='logo-size'){churchLogoSize=Math.max(10,Math.min(100,parseInt(t.value,10)||80));saveLogoLayout();return;}
  if(t.id==='logo-vertical-position'){churchLogoVPos=t.value||'center';saveLogoLayout();return;}
  if(t.id==='logo-horizontal-position'){churchLogoHPos=t.value||'center';saveLogoLayout();return;}
});

document.addEventListener('change',ev=>{
  const t=ev.target;
  if(t.id&&t.id.startsWith('credit-')){const key=t.id.slice(7);songSettings.creditFields={...(songSettings.creditFields||{}),[key]:!!t.checked};saveSongSettings();return}
  if(t.id==='song-blank-slide'){songSettings.blankSlide=t.value||'none';saveSongSettings();toast('Blank slides will be applied when you Present.');return}
  if(t.id==='song-credits-enabled'){
    songSettings.credits=!!t.checked;saveSongSettings();
    const fields=$('#song-credit-fields');if(fields)fields.hidden=!songSettings.credits;
    const wrap=$('#song-credits-pos-wrap');if(wrap)wrap.hidden=!songSettings.credits;syncCreditsTranspUi();
    return;
  }
  if(t.id==='song-credits-pos'){songSettings.creditsPos=t.value||'middle';saveSongSettings();return}
  if(t.id==='church-logo-file'){uploadLogoFile(t.files[0]);t.value=''}
  if(t.id==='logo-bg-color'){
    churchLogoBg=t.value||'#0a1428';
    try{localStorage.setItem('church-presenter-logo-bg',churchLogoBg)}catch{}
    updateLogoPreview();
    if(clr==='logo')rStage();
    try{pushAudienceClear()}catch{}
    toast('Logo background colour saved.');
    return;
  }
  if(t.id==='custom-bg-color'||t.id==='custom-bg-hex'){applyCustomBg(t);return}
  else if(t.dataset.k==='ff'||t.dataset.k==='tc'||t.dataset.k==='color'){
    const it=S[sel];
    const el=(!isSongItem(it)&&getSelectedEl())||null;
    const target=(el&&el.kind==='text')?(el.f||(el.f={fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'})):(it.f||(it.f={fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'}));
    if(t.dataset.k==='color'){
      let v=t.value.trim();
      if(v&&v[0]!=='#')v='#'+v;
      target.color=v;
      const hex=document.querySelector('[data-k="color-hex"]');
      if(hex)hex.value=target.color;
    }else target[t.dataset.k]=t.value;
    if(el&&el.kind==='text'){saveSoon();rStage()}else{if(isSongItem(it)){try{refitSong(it)}catch(err){console.warn(err)}}rAll()}
  }
});

document.addEventListener('change',ev=>{
  const t=ev.target;
  if(t&&t.id==='bg-all-slides'){
    const it=S[sel];if(!it||isSongItem(it))return;
    pushUndo();
    it.bgAll=t.checked===true;
    saveState();
    rPanel();rStrip();
  }
});

document.addEventListener('pointerdown',ev=>{
  if(ev.target.matches && ev.target.matches('input[type=range][data-k="dim"],input[type=range][data-k="mb"]'))ev.stopPropagation();
},true);
document.addEventListener('wheel',ev=>{
  if(ev.target.matches && ev.target.matches('input[type=range][data-k="dim"],input[type=range][data-k="mb"]'))ev.stopPropagation();
},true);

document.addEventListener('input',ev=>{
  const t=ev.target,k=t.dataset.k;
  if(t.id==='logo-bg-color'){
    churchLogoBg=t.value||'#0a1428';
    updateLogoPreview();
    if(clr==='logo')rStage();
    return;
  }
  if(t.id==='custom-bg-color'||t.id==='custom-bg-hex'){applyCustomBg(t);return}
  if(k){
    const it=S[sel];
    if(!it.f) it.f={fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'};
    const el=(!isSongItem(it)&&typeof getSelectedEl==='function'&&getSelectedEl())||null;
    const target=(el&&el.kind==='text')?(el.f||(el.f={fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'})):it.f;
    if(k=='dim'){const v=+t.value;bgApply(o=>{o.dim=v},{quiet:true,keepOwn:true});return}
    else if(k=='mb'){
      target.mb=Math.max(0,Math.min(25,+t.value));
      const mv=document.getElementById('lyric-margin-val');
      if(mv)mv.textContent=target.mb+'%';
    }
    else if(k=='fs'){
      target.fs=+t.value;
      const fv=document.getElementById('fs-val');
      if(fv)fv.textContent=(+t.value).toFixed(1);
    }
    else if(k=='b'||k=='sh')target[k]=t.checked?1:0;
    else if(k=='color'){
      let v=t.value.trim();
      if(v&&v[0]!=='#')v='#'+v;
      if(/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(v)){
        target.color=v.length===4?('#'+v[1]+v[1]+v[2]+v[2]+v[3]+v[3]):v;
        const hex=document.querySelector('[data-k="color-hex"]');
        const pick=document.querySelector('input[type=color][data-k=color]');
        if(hex&&document.activeElement!==hex)hex.value=target.color;
        if(pick&&document.activeElement!==pick)pick.value=target.color;
      }
    }
    else if(k=='color-hex'){
      let v=t.value.trim();
      if(v&&v[0]!=='#')v='#'+v;
      if(/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(v)){
        target.color=v.length===4?('#'+v[1]+v[1]+v[2]+v[2]+v[3]+v[3]):v;
        const pick=document.querySelector('input[type=color][data-k=color]');
        if(pick)pick.value=target.color;
      }
    }
    if(el&&el.kind==='text'){updateCurrentThumb();rStage();saveSoon()}
    else{if((k=='fs'||k=='b'||k=='mb')&&isSongItem(it)){scheduleSongFit(it)}else{rStrip();rStage();saveSoon()}}
  }
  else if(t.id==='ce'){
    const part=sls(A)[aj];
    if(part){
      part.tx=(t.innerText??t.textContent).replace(/\u00a0/g,' ');syncChunkEdit(A,part);
      const pt=$('#pt');if(pt&&S[sel]===A&&sls(A)[si]===part&&document.activeElement!==pt)pt.value=part.tx;
      updateCurrentThumb();saveSoon();
    }
  }
  else if(t.id==='song-search-query'){
    clearTimeout(searchDebounce);
    if(t.value.trim().length<3){searchController?.abort();songResults=[];renderSongResults();setSearchState('Keep typing, or press Search.');return}
    searchDebounce=setTimeout(()=>runSongSearch(t.value),450);
  }
  else if(t.id=='q'){rLib();clearTimeout(searchDebounce);const v=t.value.trim();if(v.length>=3)searchDebounce=setTimeout(()=>runSongSearch(v),450);else{songResults=[];renderSongResults();setSearchState(v?'Keep typing for online results…':'Type a title and press Search for online lyrics.')}}
  else if(t.id=='pt'){const _p=sls(S[sel])[si];_p.tx=t.value;syncChunkEdit(S[sel],_p);updateCurrentThumb();rStage()}
  else if(t.id=='el-tx'){
    const el=getSelectedEl();
    if(el&&el.kind==='text'){el.tx=t.value;saveSoon();rStage();updateCurrentThumb()}
  }
  else if(t.dataset&&t.dataset.teamField){
    const team=ensureTeam(services[activeService]);
    const key=t.dataset.teamKey, field=t.dataset.teamField;
    if(key==='band'){
      const idx=+(t.dataset.bandIndex||0);
      if(!team.band[idx])team.band[idx]=emptyContact();
      team.band[idx][field]=t.value;
    }else if(team[key]){
      team[key][field]=t.value;
    }
    saveSoon();
  }
});

document.addEventListener('mouseenter',ev=>{
  const main=ev.target.closest?.('.song-result-main');
  if(main){const row=main.closest('[data-search-result]');if(row)showSongPreview(+row.dataset.searchResult)}
},true);
document.addEventListener('mouseleave',ev=>{
  const main=ev.target.closest?.('.song-result-main');
  if(main){const row=main.closest('[data-search-result]');if(row)hideSongPreview(+row.dataset.searchResult)}
},true);

document.addEventListener('submit',ev=>{
  if(ev.target.id==='song-search-form'){
    ev.preventDefault();
    runSongSearch();
  }
});

function playCurrentVideo(){
  const v=document.querySelector('#stage .slide-video');if(v){v.play().catch(()=>{});return true}return false;
}

document.addEventListener('click',ev=>{
  const v=ev.target.closest?.('#stage .slide-video[data-video-clickseq="1"]');
  if(v){ev.stopPropagation();if(v.paused)v.play().catch(()=>{});else v.pause();}
});


let audienceWin=null,audienceIsTauri=false,presentTimerStart=0,presentTimerId=null,screenDetailsCache=null;
const AUDIENCE_LABEL='audience';

function formatElapsed(ms){
  const s=Math.floor(ms/1000);
  const m=Math.floor(s/60);
  const h=Math.floor(m/60);
  const ss=String(s%60).padStart(2,'0');
  const mm=String(m%60).padStart(2,'0');
  return h>0?(`${h}:${mm}:${ss}`):(`${mm}:${ss}`);
}
function startPresentTimer(){
  presentTimerStart=Date.now();
  const tick=()=>{
    const el=$('#pv-timer');
    if(el)el.textContent=formatElapsed(Date.now()-presentTimerStart);
  };
  tick();
  clearInterval(presentTimerId);
  presentTimerId=setInterval(tick,250);
}
function stopPresentTimer(){
  clearInterval(presentTimerId);presentTimerId=null;
  const el=$('#pv-timer');if(el)el.textContent='00:00';
}

function getFlatSlides(){
  return S.flatMap(it=>it.ord.map((_,j)=>({it,j})));
}
function getNextSlideRef(){
  const a=getFlatSlides();
  const k=a.findIndex(x=>x.it===A&&x.j===aj);
  return a[k+1]||null;
}

let pvStripScale=55;
try{const v=+localStorage.getItem('pv-strip-scale');if(!Number.isNaN(v)&&v>=0&&v<=100)pvStripScale=v}catch{}

function applyPvStripScale(){
  const root=$('#presenter-view');if(!root)return;
  const v=Math.max(0,Math.min(100,pvStripScale));
  // Grid cell min width: ~110px (small) → ~300px (large) so 16:9 height fits fully
  const cell=Math.round(110+(v/100)*190);
  root.style.setProperty('--pv-cell',cell+'px');
  root.dataset.stripMode=v<22?'text':'preview';
  const strip=$('#pv-strip');
  if(strip){
    strip.classList.toggle('pv-list',v<22);
    strip.classList.toggle('pv-grid',v>=22);
  }
  const slider=$('#pv-strip-scale');
  if(slider&&document.activeElement!==slider)slider.value=String(Math.round(v));
}

let _pvLastSel=-1,_pvLastAj=-1,_pvLastScale=-1,_pvLastClr=null;
function renderPresenterView(forceFull){
  const root=$('#presenter-view');if(!root)return;
  root.hidden=false;
  applyPvStripScale();
  $('#pv-service-title').textContent=formatServiceTitle(services[activeService]);

  // Outline: only toggle .on unless forced
  const outline=$('#pv-outline');
  if(forceFull||!outline||!outline.querySelector('.pv-oi')||outline.dataset.n!==String(S.length)){
    if(outline)outline.dataset.n=String(S.length);
    $('#pv-outline').innerHTML='<h3>Order of Sections</h3>'+S.map((it,i)=>`
      <div class="pv-oi ${i===sel?'on':''}" data-pv-oi="${i}">
        <b>${i+1}. ${e(it.t)}</b>
        <small>${e(it.k)} · ${sls(it).length} slide${sls(it).length===1?'':'s'}</small>
      </div>`).join('');
  }else{
    outline.querySelectorAll('.pv-oi').forEach((el,i)=>el.classList.toggle('on',i===sel));
  }

  const it=S[sel];
  const ss=it?sls(it):[];
  const textMode=pvStripScale<22;
  const strip=$('#pv-strip');
  if(strip){
    strip.classList.toggle('pv-list',textMode);
    strip.classList.toggle('pv-grid',!textMode);
  }

  // Only rebuild the middle strip when the service item or scale mode changes
  // (NOT on every lyric slide advance — that was causing the black flash cascade)
  const needStripRebuild=forceFull||_pvLastSel!==sel||_pvLastScale!==pvStripScale||!strip||!strip.querySelector('.pv-thumb');
  if(needStripRebuild){
    _pvLastSel=sel;
    _pvLastScale=pvStripScale;
    $('#pv-strip').innerHTML='<div class="pv-strip-heading">Slides</div>'+(ss.map((p,j)=>{
      const on=it===A&&j===aj?'on':'';
      if(textMode){
        const fullText=slideDisplayText(it,p);
        return `<div class="pv-thumb pv-text-mode ${on}" data-pv-i="${sel}" data-pv-j="${j}">
          <span class="pv-text-title">${e(p.n)}</span>
          <span class="pv-text-body">${e(fullText)}</span>
        </div>`;
      }
      // Stills only — never embed live video in presenter thumbs
      return `<div class="pv-thumb ${on}" data-pv-i="${sel}" data-pv-j="${j}">
        <div class="pv-slide-frame">${stillHtml(it,p)}</div>
        <span class="pn">${e(p.n)}</span>
      </div>`;
    }).join('')||'<p class="note" style="color:#8b949e">No slides</p>');
  }else{
    // Light path: just move the "on" highlight
    strip.querySelectorAll('.pv-thumb').forEach(el=>{
      const j=+el.dataset.pvJ;
      el.classList.toggle('on',it===A&&j===aj);
    });
  }

  // Current + next previews (stills only — no video decode in presenter UI)
  const cur=sls(A)[aj];
  $('#pv-current').innerHTML=cur?stillHtml(A,cur,aj):'';
  const nxt=getNextSlideRef();
  $('#pv-next-slide').innerHTML=nxt?stillHtml(nxt.it,sls(nxt.it)[nxt.j],nxt.j):'<div style="color:#8b949e;padding:16px;text-align:center;font-size:calc(13px * var(--ui-scale))">End of service</div>';
  document.querySelectorAll('#presenter-view [data-c]').forEach(b=>b.classList.toggle('on',b.dataset.c===clr));
  try{updatePresenterNotes()}catch{}
  _pvLastAj=aj;_pvLastClr=clr;

  requestAnimationFrame(()=>{
    const strip=$('#pv-strip');
    const curEl=strip&&strip.querySelector('.pv-thumb.on');
    if(curEl&&strip){
      try{curEl.scrollIntoView({block:'nearest',inline:'nearest',behavior:'auto'})}catch{
        const top=curEl.offsetTop-strip.clientHeight/2+curEl.clientHeight/2;
        strip.scrollTop=Math.max(0,top);
      }
    }
  });
}


/** Critical CSS so audience stage renders even without relying on main-app layout rules */
const AUDIENCE_STAGE_CSS=`
html,body{margin:0;background:#000;height:100%;overflow:hidden}
#stage{position:fixed;inset:0;background:#000;display:flex;align-items:center;justify-content:center}
#stage .cv{position:relative;overflow:hidden;background:#000;width:min(100vw,calc(100vh * 16 / 9));height:min(100vh,calc(100vw * 9 / 16))}
#stage .sl{position:relative;width:100%;height:100%;overflow:hidden;background:#000;aspect-ratio:16/9;container-type:inline-size}
#stage .bgl,#stage .dim,#stage .tx,#stage .ov{position:absolute;inset:0}
#stage .dim{background:#000;pointer-events:none}
#stage .bgl>img,#stage .bgl>video,#stage .vslot,#stage .vslot video,#stage .bg-video{width:100%;height:100%;object-fit:cover;display:block}
#stage .tx{display:flex;padding:6% 8%;color:#fff;white-space:pre-line;line-height:1.28;box-sizing:border-box}
#stage .ov.black{background:#000}#stage .ov.white{background:#fff}
#stage .ov.logo{background:#0a1428;display:flex;align-items:center;justify-content:center}
#stage .ov.logo img{max-width:80%;max-height:80%;object-fit:contain}
#stage .slide-video{position:absolute;inset:0;width:100%;height:100%;object-fit:contain;background:#000;z-index:4}
#stage .wave{background:linear-gradient(120deg,#0b3b6b,#1b7fa8,#0b3b6b,#143d5f);background-size:300% 300%}
#stage .el{position:absolute;box-sizing:border-box;transform-origin:center center;pointer-events:none;user-select:none}
#stage .el.text-el{display:flex;padding:2%;color:#fff;white-space:pre-wrap;line-height:1.28;overflow:hidden;word-break:normal;overflow-wrap:normal}
#stage .el.img-el img{width:100%;height:100%;object-fit:fill;pointer-events:none;display:block}
#stage .el-tx{width:100%;display:block}
#stage .handle{display:none}
`;

const AUDIENCE_BC_NAME='church-presenter-audience';
let audienceBc=null;
function getAudienceBc(){
  if(audienceBc)return audienceBc;
  try{audienceBc=new BroadcastChannel(AUDIENCE_BC_NAME)}catch{audienceBc=null}
  return audienceBc;
}

/** Audience-only boot: second Tauri window loads this same page with ?audience=1 */
async function bootAudienceMode(){
  document.documentElement.style.background='#000';
  document.body.className='presentation-mode';
  document.body.style.cssText='margin:0;background:#000;overflow:hidden;min-height:100vh';
  document.body.innerHTML=`<style id="audience-stage-css">${AUDIENCE_STAGE_CSS}

.sl .el{position:absolute;box-sizing:border-box;transform-origin:center center;cursor:move;user-select:none}
.sl .el.text-el{display:flex;padding:2%;color:#fff;white-space:pre-wrap;line-height:1.28;overflow:hidden}
.sl .el.img-el img{width:100%;height:100%;object-fit:fill;pointer-events:none;display:block}
.sl .el.selected{outline:2px solid #f4c542;outline-offset:2px;z-index:50!important}
.sl .el .handle{position:absolute;width:10px;height:10px;background:#f4c542;border:1px solid #000;border-radius:2px;z-index:51}
.sl .el .handle.nw{left:-5px;top:-5px;cursor:nwse-resize}
.sl .el .handle.ne{right:-5px;top:-5px;cursor:nesw-resize}
.sl .el .handle.sw{left:-5px;bottom:-5px;cursor:nesw-resize}
.sl .el .handle.se{right:-5px;bottom:-5px;cursor:nwse-resize}
.sl .el .handle.rot{left:50%;top:-22px;transform:translateX(-50%);width:12px;height:12px;border-radius:50%;cursor:grab}
.ctx-menu{position:fixed;z-index:500;min-width:160px;background:var(--pn);border:1px solid var(--ln);border-radius:8px;box-shadow:0 10px 28px #0005;padding:4px;display:none}
.ctx-menu.open{display:block}
.ctx-menu button{display:block;width:100%;text-align:left;border:0;background:none;padding:8px 12px;border-radius:6px;cursor:pointer;font:inherit;color:inherit;position:relative}
.ctx-menu button:hover{background:var(--hl)}
.ctx-menu .sub{display:none;position:absolute;left:100%;top:0;min-width:160px;background:var(--pn);border:1px solid var(--ln);border-radius:8px;box-shadow:0 10px 28px #0005;padding:4px}
.ctx-menu .has-sub:hover>.sub{display:block}
.ctx-menu button.danger{color:#c22}

</style><div id="stage"></div>`;
  document.addEventListener('keydown',ev=>{
    const isF5=ev.key==='F5'||ev.code==='F5';
    const isReload=isF5||((ev.ctrlKey||ev.metaKey)&&ev.key.toLowerCase()==='r');
    if(!isReload)return;
    ev.preventDefault();ev.stopImmediatePropagation();
    if(isF5){
      const msg={type:'control',action:'toggle-clear',key:'logo',t:Date.now()};
      try{getAudienceBc()?.postMessage(msg)}catch{}
      try{localStorage.setItem('audience-control',JSON.stringify(msg))}catch{}
    }
  },true);
  const applyHtml=(html)=>{
    const stage=document.getElementById('stage');
    if(stage)stage.innerHTML=html||'';
  };
  const applyAudienceClear=(mode,logoBg)=>{
    let layer=document.getElementById('audience-clear-layer');
    if(!mode){
      if(layer)layer.remove();
      return;
    }
    if(!layer){
      layer=document.createElement('div');
      layer.id='audience-clear-layer';
      layer.style.cssText='position:fixed;inset:0;z-index:10000;pointer-events:none';
      document.body.appendChild(layer);
    }
    layer.innerHTML='';
    if(mode==='black'){layer.style.background='#000';return}
    if(mode==='white'){layer.style.background='#fff';return}
    if(mode==='logo'){
      let bg=logoBg;
      try{if(!bg)bg=localStorage.getItem('church-presenter-logo-bg')||'#0a1428'}catch{bg='#0a1428'}
      layer.style.background=bg||'#0a1428';
      layer.style.display='flex';
      layer.style.alignItems='center';
      layer.style.justifyContent='center';
      try{
        const logo=localStorage.getItem('church-presenter-logo')||'';
        if(logo){
          const img=document.createElement('img');
          img.src=logo;img.alt='Logo';
          img.style.cssText='max-width:80%;max-height:80%;object-fit:contain';
          layer.appendChild(img);
        }
      }catch{}
      return;
    }
  };
  // Restore clear if present was already active
  try{
    const mode=localStorage.getItem('audience-clr')||'';
    if(mode)applyAudienceClear(mode);
  }catch{}
  // 1) BroadcastChannel — most reliable same-origin channel between Tauri windows
  try{
    const bc=getAudienceBc();
    if(bc){
      bc.onmessage=(ev)=>{
        const d=ev?.data;
        if(!d)return;
        if(d.type==='stage')applyHtml(d.html);
        if(d.type==='clear')applyAudienceClear(d.mode||'',d.logoBg);
        if(d.type==='ping')bc.postMessage({type:'ready'});
      };
      bc.postMessage({type:'ready'});
    }
  }catch(err){console.warn('BroadcastChannel failed',err)}

  // 2) localStorage bridge (storage events fire in *other* windows)
  try{
    const existing=localStorage.getItem('audience-stage-html');
    if(existing)applyHtml(existing);
    window.addEventListener('storage',ev=>{
      if(ev.key==='audience-stage-html'&&ev.newValue!=null)applyHtml(ev.newValue);
    });
  }catch{}

  // 3) Tauri events
  try{
    const {listen,emit}=await import('@tauri-apps/api/event');
    await listen('audience-stage',ev=>{
      const payload=ev?.payload;
      const html=typeof payload==='string'?payload:payload?.html;
      applyHtml(html);
    });
    try{await emit('audience-ready',{ok:true})}catch{}
  }catch(err){
    console.warn('Tauri event listen failed',err);
  }

  window.addEventListener('message',ev=>{
    if(ev.data&&ev.data.type==='audience-stage')applyHtml(ev.data.html);
  });
  try{window.opener&&window.opener.postMessage({type:'audience-ready'},'*')}catch{}
}

async function closeAudienceWindow(){
  if(audienceIsTauri){
    try{
      const {WebviewWindow}=await import('@tauri-apps/api/webviewWindow');
      const w=await WebviewWindow.getByLabel(AUDIENCE_LABEL);
      if(w)await w.close();
    }catch(err){console.warn('close audience',err)}
  }else{
    try{if(audienceWin&&!audienceWin.closed)audienceWin.close()}catch{}
  }
  audienceWin=null;
  audienceIsTauri=false;
}

function buildAudienceDocument(){
  // Browser fallback only — build audience HTML without a literal script end tag.
  const css=[
    'html,body{margin:0;background:#000;height:100%;overflow:hidden}',
    '#stage{position:fixed;inset:0}',
    '.sl{position:relative;width:100vw;height:100vh;overflow:hidden;background:#000;container-type:size}',
    '.bgl,.dim,.tx,.ov{position:absolute;inset:0}',
    '.bgl>img,.bgl>video,.vslot,.vslot video{width:100%;height:100%;object-fit:cover;display:block}',
    '.tx{display:flex;padding:6% 8%;color:#fff;white-space:pre-line;line-height:1.28;box-sizing:border-box}',
    '.ov.black{background:#000}.ov.white{background:#fff}',
    '.ov.logo{background:#0a1428;display:flex;align-items:center;justify-content:center}',
    '.ov.logo img{max-width:80%;max-height:80%;object-fit:contain}',
    '.slide-video{position:absolute;inset:0;width:100%;height:100%;object-fit:contain;background:#000}',
    '.el{position:absolute;box-sizing:border-box;transform-origin:center center}',
    '.el.text-el{display:flex;padding:2%;color:#fff;white-space:pre-wrap;line-height:1.28;overflow:hidden}',
    '.el.img-el img{width:100%;height:100%;object-fit:fill;display:block}'
  ].join('');
  const js=[
    "window.addEventListener('message',function(ev){",
    "if(!ev.data||ev.data.type!=='audience-stage')return;",
    "document.getElementById('stage').innerHTML=ev.data.html||'';",
    "});",
    "window.opener&&window.opener.postMessage({type:'audience-ready'},'*');"
  ].join('');
  return '<!doctype html><html><head><meta charset="utf-8"><title>Audience</title><style>'+css+'</style></head><body><div id="stage"></div><script>'+js+'</scr'+'ipt></body></html>';
}

const audienceSrcCache=new Map();
const songBgPreload=new Map(); // key -> resolved data URL, warmed at Present

async function resolveAudienceSrc(src){
  if(!src)return '';
  if(src.startsWith('data:')||src.startsWith('http://')||src.startsWith('https://')||src.startsWith('asset:')||src.startsWith('tauri:'))return src;
  if(audienceSrcCache.has(src))return audienceSrcCache.get(src);
  if(songBgPreload.has(src))return songBgPreload.get(src);
  try{
    const res=await fetch(src);
    const blob=await res.blob();
    const dataUrl=await new Promise((resolve,reject)=>{
      const r=new FileReader();
      r.onload=()=>resolve(r.result);
      r.onerror=()=>reject(r.error||new Error('read failed'));
      r.readAsDataURL(blob);
    });
    audienceSrcCache.set(src,dataUrl);
    return dataUrl;
  }catch(err){console.warn('audience src',src,err);return src}
}

/** Preload every unique background video (songs + freeform) so switches stay seamless. */
async function preloadSongBackgroundVideos(){
  const jobs=[];
  const seen=new Set();
  const collect=(media)=>{
    if(!media||media.kind!=='video')return;
    const src=(media.id&&mediaURLs[media.id])||media.path||media.source||media.src||'';
    if(!src||seen.has(src))return;
    seen.add(src);
    jobs.push((async()=>{
      try{
        const resolved=await resolveAudienceSrc(src);
        songBgPreload.set(src,resolved);
        // Warm the browser decoder with a hidden video element
        const v=document.createElement('video');
        v.muted=true;v.preload='auto';v.playsInline=true;v.src=resolved;
        v.load();
        await new Promise(r=>{
          const done=()=>{v.onloadeddata=null;v.onerror=null;r()};
          v.onloadeddata=done;v.onerror=done;
          setTimeout(done,6000);
        });
      }catch(err){console.warn('preload bg',err)}
    })());
  };
  for(const it of S){
    collect(it.bgMedia);
    for(const p of (it.parts||[])){
      collect(p.bgMedia);
      collect(p.videoMedia);
    }
  }
  await Promise.all(jobs);
  // Tell audience to warm the same sources
  try{
    const bc=getAudienceBc();
    if(bc){
      const list=[...songBgPreload.entries()].map(([src,url])=>({key:src,url}));
      bc.postMessage({type:'preload-bgs',list});
    }
  }catch{}
}

const audienceHtmlCache=new Map(); // key -> html
function audienceSlideKey(it,j){
  if(!it)return '';
  const p=sls(it)[j];
  const id=p&&p.id||j;
  const media=p&&p.bgMedia||it.bgMedia;
  const mid=media&&(media.id||media.src||'')||'';
  return (it.t||'')+'|'+id+'|'+mid+'|'+(p&&p.tx||'')+'|'+(p&&p.credits||'');
}
function resolveAudienceSrcSync(src){
  if(!src)return '';
  if(src.startsWith('data:')||src.startsWith('http://')||src.startsWith('https://')||src.startsWith('asset:')||src.startsWith('tauri:'))return src;
  if(audienceSrcCache.has(src))return audienceSrcCache.get(src);
  if(songBgPreload.has(src))return songBgPreload.get(src);
  return src; // leave blob; audience may still paint if same origin, else next preload fills cache
}
async function buildAudienceStageHtml(){
  const it=A;const part=sls(it)[aj];
  // During PPT hand-off, show a seamless black frame on the audience display only
  if(live&&typeof isPptHandoffItem==='function'&&isPptHandoffItem(it)){
    return '<div class="cv" style="position:relative;width:100%;height:100%;background:#000"><div class="sl" style="background:#000"></div></div>';
  }
  if(!it||!part)return '<div style="color:#fff;padding:40px;font:24px system-ui">No slide</div>';
  const cacheKey=audienceSlideKey(it,aj)+'|audience';
  if(audienceHtmlCache.has(cacheKey))return audienceHtmlCache.get(cacheKey);
  if(typeof ensureElements==='function'&&!isSongItem(it))ensureElements(part,it);
  let html=`<div class="cv">${sl(it,part,false,true,aj,true)}</div>`;
  const vslots=[...html.matchAll(/<div class="vslot"[^>]*>\s*<\/div>/g)];
  for(const m of vslots){
    const full=m[0];
    const sm=full.match(/data-vsrc="([^"]*)"/);
    const km=full.match(/data-vkey="([^"]*)"/);
    if(!sm)continue;
    // Prefer sync cache to avoid await lag on every advance
    let resolved=resolveAudienceSrcSync(sm[1]);
    if(resolved===sm[1]&&(sm[1].startsWith('blob:')||!sm[1].startsWith('data:'))){
      try{resolved=await resolveAudienceSrc(sm[1])}catch{}
    }
    const key=km?km[1]:sm[1];
    html=html.replace(full,
      `<video class="bg-video" data-vkey="${key}" src="${resolved}" autoplay loop muted playsinline style="width:100%;height:100%;object-fit:cover;display:block"></video>`);
  }
  const srcs=[...html.matchAll(/\bsrc="(blob:[^"]+)"/g)].map(m=>m[1]);
  for(const s of [...new Set(srcs)]){
    let resolved=resolveAudienceSrcSync(s);
    if(resolved===s){
      try{resolved=await resolveAudienceSrc(s)}catch{}
    }
    if(resolved&&resolved!==s)html=html.split(s).join(resolved);
  }
  // Audience markup must never carry editor-only selection chrome, even if a
  // stale cached slide or another renderer path included it.
  html=html.replace(/\sselected(?=[\s"])/g,'')
    .replace(/<span class="handle\b[^\"]*"[^>]*><\/span>/g,'')
    .replace(/<span class="lock-btn\b[\s\S]*?<\/span>/g,'');
  if(audienceHtmlCache.size>80){
    const first=audienceHtmlCache.keys().next().value;
    audienceHtmlCache.delete(first);
  }
  audienceHtmlCache.set(cacheKey,html);
  return html;
}

function pushAudienceClear(){
  const mode=clr||'';
  const logoBg=logoBgCss();
  try{
    localStorage.setItem('audience-clr',mode);
    localStorage.setItem('church-presenter-logo-bg',logoBg);
  }catch{}
  try{
    const bc=getAudienceBc();
    if(bc)bc.postMessage({type:'clear',mode,logoBg});
  }catch{}
  try{
    if(audienceWin&&!audienceWin.closed&&typeof audienceWin.postMessage==='function'){
      audienceWin.postMessage({type:'audience-clear',mode,logoBg},'*');
    }
  }catch{}
  // Reflect clear chrome on presenter buttons
  try{
    document.querySelectorAll('#presenter-view [data-c], [data-c]').forEach(b=>{
      b.classList.toggle('on',!!mode&&b.dataset.c===mode);
    });
  }catch{}
}

let _audiencePushTimer=null;
let _lastAudienceHtml='';
let _lastAudienceBgKey='';
let _audiencePushInflight=false;
let _audiencePushPending=false;

function extractAudienceBgKey(html){
  const m=html&&html.match(/data-vkey="([^"]+)"/);
  return m?m[1]:'';
}

async function pushAudienceStage(){
  // Queue another push if one is already running (never drop the latest slide)
  if(_audiencePushInflight){
    _audiencePushPending=true;
    return;
  }
  if(_audiencePushTimer){clearTimeout(_audiencePushTimer);_audiencePushTimer=null}
  _audiencePushInflight=true;
  try{
    const html=await buildAudienceStageHtml();
    if(html===_lastAudienceHtml)return;
    const bgKey=extractAudienceBgKey(html);
    const sameBg=!!(bgKey&&bgKey===_lastAudienceBgKey);
    _lastAudienceHtml=html;
    _lastAudienceBgKey=bgKey||'';
    // Broadcast first (fast path for audience)
    try{const bc=getAudienceBc();if(bc)bc.postMessage({type:'stage',html,sameBg,bgKey})}catch(err){console.warn(err)}
    try{if(audienceWin&&!audienceWin.closed&&typeof audienceWin.postMessage==='function')audienceWin.postMessage({type:'audience-stage',html,sameBg,bgKey},'*')}catch{}
    // localStorage is backup only — can be large/slow; skip when huge
    try{
      if(html.length<900000){
        localStorage.setItem('audience-stage-html',html);
        if(bgKey)localStorage.setItem('audience-bg-key',bgKey);
        localStorage.setItem('audience-stage-ping',String(Date.now()));
      }
    }catch{}
    if(clr)pushAudienceClear();
  }finally{
    _audiencePushInflight=false;
    if(_audiencePushPending){
      _audiencePushPending=false;
      pushAudienceStage();
    }
  }
}


/** Handle remote controls from the audience window (keys / click). Dedupes BC + localStorage. */
let lastAudienceControlT=0;
function handleAudienceControl(msg){
  if(!live||!msg)return;
  // Same event arrives via BroadcastChannel and localStorage — only handle once
  const t=Number(msg.t)||0;
  if(t){
    if(t<=lastAudienceControlT)return;
    lastAudienceControlT=t;
  }else{
    // fallback debounce if no timestamp
    const now=Date.now();
    if(now-lastAudienceControlT<150)return;
    lastAudienceControlT=now;
  }
  const action=msg.action;
  if(action==='next'){go(1);return}
  if(action==='prev'){go(-1);return}
  if(action==='end'){endPresent();return}
  if(action==='toggle-clear'){
    const mode=msg.key; // black | white | logo
    if(!mode)return;
    clr=clr===mode?null:mode;
    rStage();
    pushAudienceClear();
    if(document.body.classList.contains('presenter-mode'))renderPresenterView();
    return;
  }
}

async function detectTauriRuntime(){
  try{
    if(typeof isTauri==='function'){
      const v=await Promise.resolve(isTauri());
      if(v)return true;
    }
  }catch{}
  return !!(
    window.__TAURI_INTERNALS__
    ||window.__TAURI__
    ||window.__TAURI_METADATA__
    ||window.isTauri
    ||(window.__TAURI_OS_PLUGIN_INTERNALS__)
  );
}

async function openAudienceOnDisplay(displayId){
  await closeAudienceWindow();

  // Prefer Tauri WebviewWindow whenever the API is importable — do not rely only on isTauri().
  let WebviewWindowMod=null;
  try{
    WebviewWindowMod=await import('@tauri-apps/api/webviewWindow');
  }catch(err){
    console.warn('webviewWindow import failed',err);
  }
  const WebviewWindow=WebviewWindowMod?.WebviewWindow||WebviewWindowMod?.default||null;
  const looksLikeTauri=await detectTauriRuntime();

  if(WebviewWindow&&looksLikeTauri){
    let monitors=[],current=null;
    try{
      monitors=await availableMonitors();
      current=await currentMonitor();
    }catch(err){console.warn('monitor query failed',err)}

    // Prefer selected label, else a monitor that is not the current one, else first
    let target=null;
    if(displayId&&displayId!=='default'&&displayId!=='this'&&displayId!=='external'){
      target=monitors.find(m=>String(m.name)===String(displayId))
        ||monitors.find((m,i)=>String(i)===String(displayId));
    }
    if(!target&&monitors.length>1&&current){
      target=monitors.find(m=>m.name!==current.name)||monitors.find(m=>{
        try{return m.position.x!==current.position.x||m.position.y!==current.position.y}catch{return false}
      });
    }
    if(!target)target=monitors[monitors.length>1?1:0]||null;

    // Minimal standalone page — no app module, polls localStorage for slides
    const audienceUrl=(()=>{
      try{
        const u=new URL('audience.html',location.href);
        return u.toString();
      }catch{
        return 'audience.html';
      }
    })();

    // Write slide HTML BEFORE opening the window so the audience page
    // can read it immediately on load (and via its 80ms poll).
    try{
      const html=await buildAudienceStageHtml();
      localStorage.setItem('audience-stage-html',html);
      localStorage.setItem('audience-stage-ping',String(Date.now()));
      const bc=getAudienceBc();
      if(bc)bc.postMessage({type:'stage',html});
    }catch(err){console.warn('pre-write audience stage failed',err)}

    const opts={
      url:audienceUrl,
      title:'Audience',
      decorations:false,
      resizable:false,
      // Set fullscreen after creation and display placement. Some platforms ignore
      // fullscreen in the initial WebviewWindow options when explicit geometry is set.
      fullscreen:false,
      focus:false,
      visible:true,
      skipTaskbar:true,
      alwaysOnTop:false,
    };
    if(target){
      try{
        const pos=typeof target.position?.toLogical==='function'
          ?target.position.toLogical(target.scaleFactor)
          :{x:(target.position?.x||0)/(target.scaleFactor||1),y:(target.position?.y||0)/(target.scaleFactor||1)};
        const size=typeof target.size?.toLogical==='function'
          ?target.size.toLogical(target.scaleFactor)
          :{width:(target.size?.width||1920)/(target.scaleFactor||1),height:(target.size?.height||1080)/(target.scaleFactor||1)};
        opts.x=pos.x;
        opts.y=pos.y;
        opts.width=size.width;
        opts.height=size.height;
      }catch(err){console.warn('Could not map monitor geometry',err)}
    }

    // Close any leftover window with this label first
    try{
      const existing=await WebviewWindow.getByLabel(AUDIENCE_LABEL);
      if(existing)await existing.close();
    }catch{}

    const audience=new WebviewWindow(AUDIENCE_LABEL,opts);
    await new Promise((resolve,reject)=>{
      let settled=false;
      const done=(fn,arg)=>{if(settled)return;settled=true;fn(arg)};
      audience.once('tauri://created',()=>done(resolve));
      audience.once('tauri://error',e=>{
        const msg=(typeof e?.payload==='string'?e.payload:null)
          ||(e?.payload&&typeof e.payload==='object'?JSON.stringify(e.payload):null)
          ||'Failed to create audience window. Put audience.html next to index.html, and ensure capabilities include windows: ["main","audience"].';
        done(reject,new Error(msg));
      });
      setTimeout(()=>done(resolve),2500);
    });

    audienceIsTauri=true;
    audienceWin=audience;
    // Ensure control channel listener is attached (in case startup IIFE ran too early)
    try{
      const bc=getAudienceBc();
      if(bc&&!bc.__controlBound){
        bc.__controlBound=true;
        bc.addEventListener('message',(ev)=>{
          const d=ev?.data;
          if(!d)return;
          if(d.type==='ready'){pushAudienceStage();refocusPresenterWindow().catch(()=>{})}
          if(d.type==='control')handleAudienceControl(d);
        });
      }
    }catch{}
    // Keep keyboard focus on the presenter (main) window — audience is output-only
    const ensureAudienceFullscreen=async()=>{if(pptAudienceClosed)return;
      if(target){try{
          const pos=typeof target.position?.toLogical==='function'?target.position.toLogical(target.scaleFactor):{x:(target.position?.x||0)/(target.scaleFactor||1),y:(target.position?.y||0)/(target.scaleFactor||1)};
          const size=typeof target.size?.toLogical==='function'?target.size.toLogical(target.scaleFactor):{width:(target.size?.width||1920)/(target.scaleFactor||1),height:(target.size?.height||1080)/(target.scaleFactor||1)};
          await audience.setPosition(pos);
          await audience.setSize(size);
      }catch(err){console.warn('Could not place audience window on selected display',err)}}
      try{if(typeof audience.setFullscreen==='function')await audience.setFullscreen(true)}catch(err){console.warn('Could not fullscreen audience window',err)}
    };
    await new Promise(r=>setTimeout(r,100));
    await ensureAudienceFullscreen();
    await pushAudienceStage();
    // Re-push stage as the audience page finishes booting (no setFocus — that stole control)
    for(const ms of [300,800,1500]){
      setTimeout(()=>{
        if(!live)return;
        ensureAudienceFullscreen();
        pushAudienceStage();
      },ms);
    }
    return;
  }

  // Browser-only fallback (popup — blocked inside Tauri webviews)
  if(looksLikeTauri){
    throw new Error(
      'Tauri is running but the audience WebviewWindow could not be opened. '+
      'Check that index.html is the updated build, capabilities list "audience" in windows, '+
      'and core:webview:allow-create-webview-window is granted. Then fully restart the app.'
    );
  }

  const features='popup=yes,width=1280,height=720';
  audienceWin=window.open('audience.html', 'church-presenter-audience', features);
  if(!audienceWin){
    throw new Error('Could not open the audience window. Allow pop-ups for this site, or use Settings → Presentation display → This screen.');
  }
  audienceIsTauri=false;
  audienceWin.document.write(buildAudienceDocument());
  audienceWin.document.close();
  try{
    if(screenDetailsCache&&displayId!=null&&displayId!=='default'&&displayId!=='this'){
      const scr=screenDetailsCache.screens?.find(s=>String(s.label||s.id)===String(displayId))||screenDetailsCache.screens?.[+displayId];
      if(scr&&typeof audienceWin.moveTo==='function'){
        audienceWin.moveTo(scr.availLeft||scr.left||0, scr.availTop||scr.top||0);
        audienceWin.resizeTo(scr.availWidth||scr.width||1280, scr.availHeight||scr.height||720);
      }
    }
  }catch(err){console.warn('Could not position audience window',err)}
  try{
    if(audienceWin.document.documentElement.requestFullscreen){
      await audienceWin.document.documentElement.requestFullscreen();
    }
  }catch{}
}

let presentLocalOnly=false; // true = single display, no presenter chrome / no audience window

/** True when we should present on this screen only (no external output). */
async function shouldPresentLocally(mode){
  const m=String(mode||'').toLowerCase();
  // Explicit this-screen choices
  if(m==='this'||m==='default'||m===''||m==='0'||m==='primary')return true;

  let monitors=[];
  try{monitors=await availableMonitors()}catch{monitors=[]}
  // 0–1 monitors → always local fullscreen presentation (no separate audience window)
  if(!monitors||monitors.length<=1){
    try{
      if(window.getScreenDetails){
        const details=await window.getScreenDetails();
        const screens=details.screens||[];
        if(screens.length<=1)return true;
      }else return true;
    }catch{return true}
  }
  if(monitors&&monitors.length>1){
    try{
      const cur=await currentMonitor();
      if(cur){
        const curName=String(cur.name||'');
        if(m&&m===curName.toLowerCase())return true;
      }
    }catch{}
    // Selected something that looks like another display → dual mode
    return false;
  }
  try{
    if(window.getScreenDetails){
      const details=await window.getScreenDetails();
      const screens=details.screens||[];
      if(screens.length<=1)return true;
      const cur=details.currentScreen;
      const match=screens.find(s=>String(s.label||s.id)===String(mode));
      if(match&&cur&&match===cur)return true;
      if(!mode||m==='this'||m==='default'||m==='0')return true;
      return false;
    }
  }catch{}
  return true;
}

function hidePresenterView(){
  const pv=document.getElementById('presenter-view');
  if(!pv)return;
  pv.hidden=true;
  pv.style.removeProperty('display');
  pv.style.removeProperty('z-index');
  pv.style.removeProperty('visibility');
  pv.style.removeProperty('opacity');
  pv.style.removeProperty('pointer-events');
  pv.style.removeProperty('position');
  pv.style.removeProperty('inset');
}

/** Dual-monitor: presenter chrome + separate audience window. */
function showPresenterShell(){
  presentLocalOnly=false;
  document.body.classList.remove('home-mode');
  document.body.classList.add('presentation-mode','presenter-mode');
  const pv=document.getElementById('presenter-view');
  if(pv){
    pv.hidden=false;
    pv.removeAttribute('hidden');
    pv.style.setProperty('display','grid','important');
    pv.style.setProperty('z-index','200','important');
    pv.style.setProperty('visibility','visible','important');
    pv.style.setProperty('opacity','1','important');
    pv.style.setProperty('pointer-events','auto','important');
    pv.style.setProperty('position','fixed','important');
    pv.style.setProperty('inset','0','important');
  }
  try{renderPresenterView()}catch(err){console.error('renderPresenterView',err)}
}

/** Enter true OS-level fullscreen on the main window (Tauri) or browser fullscreen API. */
async function enterLocalDisplayFullscreen(){
  // Tauri: native window fullscreen covers the whole monitor (not just CSS layout)
  try{
    const {getCurrentWindow}=await import('@tauri-apps/api/window');
    const w=getCurrentWindow();
    if(w){
      if(typeof w.setFullscreen==='function'){
        await w.setFullscreen(true);
        return true;
      }
      if(typeof w.maximize==='function'){
        try{await w.maximize()}catch{}
      }
    }
  }catch(err){console.warn('Tauri setFullscreen failed',err)}
  // Browser fallback
  try{
    const el=document.documentElement;
    if(!document.fullscreenElement){
      const req=el.requestFullscreen||el.webkitRequestFullscreen||el.msRequestFullscreen;
      if(req){
        await Promise.resolve(req.call(el));
        return true;
      }
    }
  }catch(err){console.warn('requestFullscreen failed',err)}
  return false;
}

async function exitLocalDisplayFullscreen(){
  try{
    const {getCurrentWindow}=await import('@tauri-apps/api/window');
    const w=getCurrentWindow();
    if(w&&typeof w.setFullscreen==='function'){
      await w.setFullscreen(false);
    }
  }catch{}
  try{
    if(document.fullscreenElement){
      const exit=document.exitFullscreen||document.webkitExitFullscreen||document.msExitFullscreen;
      if(exit)await Promise.resolve(exit.call(document));
    }
  }catch{}
}

/** Single-display: true display fullscreen, stage only (no presenter view). Esc / B / W / L still work. */
function showLocalPresentShell(){
  presentLocalOnly=true;
  document.body.classList.remove('home-mode','presenter-mode');
  document.body.classList.add('presentation-mode');
  hidePresenterView();
  try{rStage()}catch(err){console.error(err)}
  // Fire-and-forget OS fullscreen so the click gesture is still valid
  enterLocalDisplayFullscreen().catch(()=>{});
  try{document.body.focus()}catch{}
}


let presentStarting=false,presentFrom=null,pptShowActive=false,pptAudienceClosed=false;
let presentEndedAt=0;
let pptLaunchArmed=false;
function jumpToServiceStart(){
  // Hard reset: first section, first slide/part — used whenever Present starts
  elSel=null;imageSel=-1;stripFocus='slide';clr=null;
  if(!S||!S.length){A=null;sel=0;si=0;aj=0;return;}
  if(presentFrom&&S.includes(presentFrom.item)){
    const it=presentFrom.item;let jj=-1,nn=0;
    for(let x=0;x<it.ord.length;x++){if(it.ord[x]===presentFrom.id){if(nn===presentFrom.rank){jj=x;break}nn++}}
    A=it;sel=S.indexOf(it);si=Math.max(0,jj);aj=si;
    return;
  }
  sel=0;si=0;A=S[0];
  // First item in presentation order (ord), not last-edited part
  if(A&&Array.isArray(A.ord)&&A.ord.length){aj=0;si=0;}
  else{aj=0;si=0;}
}
async function startPresent(){
  // Recover from a stuck "starting" flag (e.g. previous attempt never finished)
  if(presentStarting&&Date.now()-presentEndedAt>800){
    presentStarting=false;
  }
  if(presentStarting){
    if(presentLocalOnly)showLocalPresentShell();else showPresenterShell();
    return;
  }
  // If already live, don't no-op silently — re-show shell (recovers stuck UI)
  if(live){
    if(presentLocalOnly)showLocalPresentShell();else showPresenterShell();
    const liveBtn=document.getElementById('live');
    if(liveBtn)liveBtn.textContent='End presentation';
    return;
  }
  presentStarting=true;
  pptLaunchArmed=false;
  // Presentation output must not include editor selection or resize handles.
  elSel=null;
  MS.els.clear();
  MS.elAnchor=null;
  try{
    // ALWAYS begin at section 0 / slide 0 BEFORE live work (PPT handoff, audience, etc.)
    jumpToServiceStart();
    live=true;
    const liveBtn=document.getElementById('live');
    if(liveBtn)liveBtn.textContent='End presentation';
    clr=null;
    pptHandoffActive=false;
    pptHandoffKey='';
    pptCompletedKeys=new Set();
    try{pptOpenedThisPresent.clear()}catch{pptOpenedThisPresent=new Set();}
    try{audienceHtmlCache.clear()}catch{}
    _lastAudienceHtml='';
    _lastAudienceBgKey='';
    _audiencePushPending=false;
    // Inject settings blank slides for this presentation only
    try{
      injectPresentBlanks();
    }catch(err){console.warn('injectPresentBlanks',err)}
    try{
      injectPptPresentPlaceholders();
    }catch(err){console.warn('injectPptPresentPlaceholders',err)}
    // Re-assert first slide after blanks/placeholders (must stay sequential from start)
    jumpToServiceStart();
    try{rStrip();rOutline();rStage();updateSpeakerNotes()}catch{}
    try{pushAudienceStage()}catch{}
    // Warm PowerPoint files in the background (write to disk only — open when section is reached)
    preloadPowerpointsForService().catch(()=>{})
    startPresentTimer();

    const mode=($('#display-setting')?.value)||selectedDisplay||'';
    selectedDisplay=mode;

    const local=await shouldPresentLocally(mode);
    if(local){
      // Single display (or this screen): fullscreen stage on this window — audience = main stage
      presentLocalOnly=true;
      showLocalPresentShell();
      try{rStage()}catch{}
      try{await enterLocalDisplayFullscreen()}catch(err){console.warn('fullscreen',err)}
      requestAnimationFrame(()=>{ if(live){showLocalPresentShell();try{rStage()}catch{}} });
      setTimeout(()=>{ if(live&&presentLocalOnly){try{enterLocalDisplayFullscreen()}catch{}} },200);
      preloadSongBackgroundVideos().catch(()=>{});
    }else{
      // Dual display: presenter chrome + separate audience window
      showPresenterShell();
      try{rStage()}catch{}
      try{renderPresenterView()}catch{}
      try{
        await openAudienceOnDisplay(mode);
        await pushAudienceStage();
        showPresenterShell();
        await refocusPresenterWindow();
        preloadSongBackgroundVideos().catch(()=>{});
        for(const ms of [120,350,700,1400]){
          setTimeout(()=>{
            if(!live||presentLocalOnly)return;
            showPresenterShell();
            pushAudienceStage().catch(()=>{});
            refocusPresenterWindow().catch(()=>{});
          },ms);
        }
      }catch(err){
        console.error('Audience open failed, presenting locally',err);
        try{await closeAudienceWindow()}catch{}
        presentLocalOnly=true;
        showLocalPresentShell();
        try{rStage()}catch{}
        try{await enterLocalDisplayFullscreen()}catch{}
      }
    }
    // Do not allow navigation-triggered hand-offs until startup has finished. This
    // prevents focus/key events during audience-window setup from skipping ahead.
    pptLaunchArmed=true;
    jumpToServiceStart();
    try{rStrip();rOutline();rStage();updateSpeakerNotes();pushAudienceStage()}catch{}
    // Only hand off if PowerPoint is truly the first service item.
    if(live&&A&&(presentFrom?true:A===S[0])&&isPptHandoffItem(A))openNativePowerpointIfNeeded(A);
  }catch(err){
    console.error('startPresent failed',err);
    live=true;
    presentLocalOnly=true;
    showLocalPresentShell();
    try{rStage()}catch{}
    try{await enterLocalDisplayFullscreen()}catch{}
  }finally{
    presentStarting=false;
  }
}

async function refocusPresenterWindow(){if(pptShowActive)return;
  try{
    // Prefer Tauri main window focus
    const {getCurrentWindow}=await import('@tauri-apps/api/window');
    const w=getCurrentWindow();
    if(w&&typeof w.setFocus==='function')await w.setFocus();
  }catch{}
  try{window.focus()}catch{}
  try{document.body&&document.body.focus()}catch{}
}

function isPresentingUI(){
  const btn=document.getElementById('live');
  return !!(live
    || document.body.classList.contains('presentation-mode')
    || document.body.classList.contains('presenter-mode')
    || (btn&&/end presentation/i.test(btn.textContent||'')));
}

document.addEventListener('keydown',function _escEndPresent(ev){
  if(isPresentingUI()){
    const isF5=ev.key==='F5'||ev.code==='F5';
    const isReload=isF5||((ev.ctrlKey||ev.metaKey)&&ev.key.toLowerCase()==='r');
    if(isReload){
      ev.preventDefault();ev.stopImmediatePropagation();
      if(isF5){
        clr=clr==='logo'?null:'logo';
        rStage();pushAudienceClear();
        if(document.body.classList.contains('presenter-mode'))renderPresenterView();
      }
      return;
    }
  }
  if(ev.key!=='Escape'&&ev.code!=='Escape')return;
  if(closeConfirmOpen)return;
  if(document.getElementById('el-inline-edit'))return;
  // Always end present if UI looks like presenting — Esc must never get stuck
  if(isPresentingUI()){
    ev.preventDefault();
    ev.stopImmediatePropagation();
    pptHandoffActive=false;
    endPresent();
  }
}, true);

window.addEventListener('keydown',function _escEndPresentWin(ev){
  if(ev.key!=='Escape'&&ev.code!=='Escape')return;
  if(closeConfirmOpen)return;
  if(isPresentingUI()){
    ev.preventDefault();
    pptHandoffActive=false;
    endPresent();
  }
}, true);

function endPresent(){presentFrom=null;pptShowActive=false;pptAudienceClosed=false;
  presentStarting=false;
  presentEndedAt=Date.now();
  live=false;
  pptLaunchArmed=false;
  presentLocalOnly=false;
  ed=false;
  clr=null;
  pptHandoffActive=false;
  pptHandoffKey='';
  pptCompletedKeys=new Set();
  try{stopPresentTimer()}catch{}
  document.body.classList.remove('presentation-mode','presenter-mode');
  // Force layout back to editor (remove any leftover fullscreen styles)
  document.body.style.removeProperty('overflow');
  const liveBtn=document.getElementById('live');
  if(liveBtn)liveBtn.textContent='Present';
  const pv=document.getElementById('presenter-view');
  if(pv){
    pv.hidden=true;
    pv.setAttribute('hidden','');
    ['display','z-index','visibility','opacity','pointer-events','position','inset'].forEach(p=>pv.style.removeProperty(p));
  }
  // Unhide editor chrome in case CSS was overridden
  ['.top','#strip','#outline','#panel','.bar','.mid'].forEach(sel=>{
    try{
      document.querySelectorAll(sel).forEach(el=>{
        el.style.removeProperty('display');
      });
    }catch{}
  });
  _lastAudienceHtml='';
  _lastAudienceBgKey='';
  _audiencePushPending=false;
  _audiencePushInflight=false;
  if(_audiencePushTimer){clearTimeout(_audiencePushTimer);_audiencePushTimer=null}
  try{
    localStorage.setItem('audience-clr','');
    localStorage.setItem('audience-stage-html','');
    localStorage.setItem('audience-bg-key','');
    const bc=getAudienceBc();
    if(bc)bc.postMessage({type:'clear',mode:''});
  }catch{}
  try{closeAudienceWindow()}catch{}
  try{stopPptSlideshow()}catch{}
  // Exit fullscreen sync-ish
  try{exitLocalDisplayFullscreen()}catch{}
  (async()=>{
    try{
      const {getCurrentWindow}=await import('@tauri-apps/api/window');
      const w=getCurrentWindow();
      if(w&&typeof w.setFullscreen==='function')await w.setFullscreen(false);
      if(w&&typeof w.unmaximize==='function'){/* leave maximized alone */}
    }catch{}
    try{
      if(document.fullscreenElement){
        const exit=document.exitFullscreen||document.webkitExitFullscreen||document.msExitFullscreen;
        if(exit)await Promise.resolve(exit.call(document));
      }
    }catch{}
  })();
  try{restorePptPresentPlaceholders()}catch{}
  try{removePresentBlanksFromService()}catch{}
  try{rAll()}catch{}
  try{rStage()}catch{}
  // Second pass after fullscreen exits (layout reflow)
  setTimeout(()=>{
    live=false;
    document.body.classList.remove('presentation-mode','presenter-mode');
    if(liveBtn)liveBtn.textContent='Present';
    try{rAll()}catch{}
  },80);
  setTimeout(()=>{
    live=false;
    document.body.classList.remove('presentation-mode','presenter-mode');
    try{rStage()}catch{}
  },250);
}

// One Esc exits present: browser may exit fullscreen first — mirror that into endPresent
document.addEventListener('fullscreenchange',()=>{
  // If OS/browser left fullscreen but we still look like presenting, fully end
  if(!document.fullscreenElement&&isPresentingUI()){
    endPresent();
  }
});
window.addEventListener('message',ev=>{
  if(ev.data&&ev.data.type==='audience-ready')pushAudienceStage();
});

// When the audience window finishes booting, push the current slide
(async()=>{
  try{
    const {listen}=await import('@tauri-apps/api/event');
    await listen('audience-ready',()=>{pushAudienceStage()});
  }catch{}
  try{
    const bc=getAudienceBc();
    if(bc){
      const onBc=(ev)=>{
        const d=ev?.data;
        if(!d)return;
        if(d.type==='ready'&&(audienceIsTauri||(audienceWin&&!audienceWin.closed))){
          pushAudienceStage();
          refocusPresenterWindow().catch(()=>{});
        }
        if(d.type==='control')handleAudienceControl(d);
      };
      bc.addEventListener('message',onBc);
      // also keep onmessage for environments that only fire that
      bc.onmessage=onBc;
    }
  }catch{}
})();

// Audience remote control via localStorage (backup channel; deduped in handleAudienceControl)
setInterval(()=>{
  if(!live)return;
  try{
    const raw=localStorage.getItem('audience-control');
    if(!raw)return;
    const msg=JSON.parse(raw);
    if(!msg||!msg.t)return;
    handleAudienceControl(msg);
  }catch{}
},60);

// Presenter strip size slider
document.addEventListener('input',ev=>{
  if(ev.target?.id==='pv-strip-scale'){
    pvStripScale=+ev.target.value;
    try{localStorage.setItem('pv-strip-scale',String(pvStripScale))}catch{}
    if(document.body.classList.contains('presenter-mode'))renderPresenterView();
  }
});



document.addEventListener('keydown',ev=>{
  const typing=isTypingTarget(ev.target);

  // Presentation mode — never treat stage output as a text field
  if(document.body.classList.contains('presentation-mode')||document.body.classList.contains('presenter-mode')){
    if(ev.key==='Escape'){ev.preventDefault();endPresent();return}
    if(typing)return; // safety
    const k=ev.key.toLowerCase();
    if(k==='b'||k==='w'||k==='l'){
      ev.preventDefault();
      const mode=k==='b'?'black':k==='w'?'white':'logo';
      clr=clr===mode?null:mode;
      rStage();
      pushAudienceClear();
      if(document.body.classList.contains('presenter-mode')){
        try{renderPresenterView()}catch{}
      }
      return;
    }
    if(k==='arrowright'||k==='arrowdown'||k==='pagedown'||k==='enter'||ev.key===' '||ev.code==='Space'||ev.code==='NumpadEnter'){
      ev.preventDefault();go(1);return;
    }
    if(k==='arrowleft'||k==='arrowup'||k==='pageup'||k==='backspace'){
      ev.preventDefault();go(-1);return;
    }
    return;
  }

  // Modal / overlay open: only Escape (and typing inside the modal) — never slide/order hotkeys
  if(isEditorOverlayOpen()){
    if(ev.key==='Escape'&&!$('#logo-controls-modal')?.hidden){ev.preventDefault();ev.stopPropagation();$('#logo-controls-modal').hidden=true;return}
    if(typing)return;
    if(ev.key==='Escape'){
      // Handled below in typing branch pattern — close topmost modal
      if(!$('#edit-song-modal')?.hidden){ev.preventDefault();closeEditSongModal();return}
      if(!$('#struct-modal')?.hidden){ev.preventDefault();closeStruct();return}
      if(!$('#media-modal')?.hidden){ev.preventDefault();closeMediaLibrary();return}
      if(!$('#library-modal')?.hidden){ev.preventDefault();closeLibraryModal();return}
      if(!$('#share-modal')?.hidden){ev.preventDefault();closeShareModal();return}
      if(!$('#team-modal')?.hidden){ev.preventDefault();closeTeamModal();return}
      if(!$('#import-modal')?.hidden){ev.preventDefault();try{hideImportModal()}catch{};return}
      if($('#settings')?.classList.contains('open')){ev.preventDefault();closeSettingsAndRefresh();return}
      if($('#song-search-panel')&&!$('#song-search-panel').hidden){ev.preventDefault();$('#song-search-panel').hidden=true;return}
      if(!$('#close-confirm-modal')?.hidden){ev.preventDefault();try{finishCloseConfirm(false)}catch{};return}
    }
    return; // block Ctrl+C/V, Backspace, arrows that move slides, etc.
  }

  // While editing slide text on the canvas
  if(ev.target.closest&&ev.target.closest('[contenteditable]')){
    if(ev.key==='Escape'){ev.preventDefault();stopEdit()}
    return; // full keyboard available for typing
  }

  // Any input / textarea / select — no global shortcuts
  if(typing){
    if(ev.key==='Escape'){
      // Topmost modal first (edit-song sits above library)
      if(!$('#edit-song-modal').hidden){closeEditSongModal();return}
      if(!$('#struct-modal').hidden){closeStruct();return}
      if(!$('#media-modal').hidden){closeMediaLibrary();return}
      if(!$('#library-modal').hidden){closeLibraryModal();return}
      if($('#settings').classList.contains('open')){closeSettingsAndRefresh();return}
      if($('#service-form').classList.contains('open')){$('#service-form').classList.remove('open');return}
      if(!$('#email-modal')?.hidden){closeEmailCompose();return}
    }
    return;
  }

  // Canvas element OR selected strip slide/section: delete / copy / paste / duplicate
  if(!live&&(ev.key==='Backspace'||ev.key==='Delete')){
    if(!isSongItem(A)&&elSel){ev.preventDefault();deleteSelectedEl();return}
    if(!elSel){
      ev.preventDefault();
      // Section header selected → remove whole service item
      if(stripFocus==='item'){removeServiceItem(sel);return}
      deleteSelectedSlide();
      return;
    }
  }
  if(!live&&(ev.ctrlKey||ev.metaKey)&&ev.key.toLowerCase()==='z'&&!ev.shiftKey){
    if(!isTypingTarget(ev.target)){ev.preventDefault();undoEditor();return}
  }
  if(!live&&(ev.ctrlKey||ev.metaKey)&&((ev.key.toLowerCase()==='y')||(ev.key.toLowerCase()==='z'&&ev.shiftKey))){
    if(!isTypingTarget(ev.target)){ev.preventDefault();redoEditor();return}
  }
  if(!live&&(ev.ctrlKey||ev.metaKey)&&ev.key.toLowerCase()==='c'){
    if(!isSongItem(A)&&elSel){ev.preventDefault();copySelectedEl();return}
    if(!elSel){ev.preventDefault();copySelectedSlide();return}
  }
  if(!live&&(ev.ctrlKey||ev.metaKey)&&ev.key.toLowerCase()==='v'){
    /* Handled by document paste event to support screenshots and system clipboard */
  }
  if(!live&&(ev.ctrlKey||ev.metaKey)&&ev.key.toLowerCase()==='d'){
    if(!isSongItem(A)&&elSel){ev.preventDefault();duplicateSelectedEl();return}
    if(!elSel){
      ev.preventDefault();
      const i=S.indexOf(A);
      if(i>=0)duplicateSlideInItem(i,aj);
      return;
    }
  }

  // Global editor shortcuts (only when not typing)
  // B / W / L clear overlays are presentation-only (handled above in present modes)
  const k=ev.key.toLowerCase();
  if(ev.key==='Escape'){
    // Topmost modal first (edit-song sits above library)
    if(!$('#edit-song-modal').hidden){closeEditSongModal();return}
    if(!$('#struct-modal').hidden){closeStruct();return}
    if(!$('#media-modal').hidden){closeMediaLibrary();return}
    if(!$('#library-modal').hidden){closeLibraryModal();return}
    if($('#settings').classList.contains('open')){closeSettingsAndRefresh();return}
    if($('#service-form').classList.contains('open')){$('#service-form').classList.remove('open');return}
    if(!$('#email-modal').hidden){closeEmailCompose();return}
    clr=null;rStage();return;
  }
  // Navigation: space / down / right = next; up / left = previous
  if(k==='arrowright'||k==='arrowdown'||ev.key===' '){ev.preventDefault();go(1);return}
  if(k==='arrowleft'||k==='arrowup'){ev.preventDefault();go(-1);return}
});


function openTextElEditor(el, elNode){
  if(!el||el.kind!=='text'||live)return;
  // Prefer editing inside the actual text box on the slide
  let node=elNode;
  if(!node||!node.isConnected){
    node=document.querySelector('#stage [data-el="'+el.id+'"]');
  }
  if(node&&node.isConnected){
    beginInPlaceTextEdit(el, node);
    return;
  }
  // Fallback floating editor if the node is missing
  elSel=el.id;
  const existing=document.getElementById('el-inline-edit');
  if(existing)existing.remove();
  const ta=document.createElement('textarea');
  ta.id='el-inline-edit';
  ta.value=el.tx||'';
  Object.assign(ta.style,{
    position:'fixed',left:'120px',top:'120px',width:'320px',height:'120px',
    zIndex:'10000',font:'16px/1.35 Figtree,system-ui,sans-serif',padding:'10px',
    border:'2px solid var(--ac)',borderRadius:'8px',background:'#0f141a',color:'#fff',resize:'both',
    boxShadow:'0 12px 32px rgba(0,0,0,.55)'
  });
  document.body.appendChild(ta);
  requestAnimationFrame(()=>{ta.focus();try{ta.setSelectionRange(ta.value.length,ta.value.length)}catch{}});
  let done=false;
  const commit=()=>{
    if(done)return;done=true;
    const next=ta.value;
    if(next!==(el.tx||'')){pushUndo();el.tx=next;}
    if(ta.isConnected)ta.remove();
    saveSoon();try{updateCurrentThumb()}catch{}try{rPanel()}catch{}try{rStage()}catch{}
  };
  ta.addEventListener('blur',()=>setTimeout(commit,50));
  ta.addEventListener('keydown',ke=>{
    if(ke.key==='Escape'){ke.preventDefault();ke.stopPropagation();ta.blur()}
    if(ke.key==='Enter'&&(ke.ctrlKey||ke.metaKey)){ke.preventDefault();ta.blur()}
  });
}
/** Edit text directly inside the slide text box (contenteditable). */
function beginInPlaceTextEdit(el, elNode){
  if(!el||el.kind!=='text'||live||!elNode)return;
  // Cancel any active drag so it cannot steal the edit
  elDrag=null;
  elSel=el.id;
  // Ensure selection chrome without destroying the node mid-edit
  document.querySelectorAll('#stage .el').forEach(n=>{
    n.classList.toggle('selected',n.dataset.el===el.id);
  });
  const span=elNode.querySelector('.el-tx')||elNode;
  // Already editing?
  if(span.isContentEditable&&span.getAttribute('data-editing')==='1'){
    span.focus();
    return;
  }
  span.setAttribute('data-editing','1');
  elNode.setAttribute('data-editing','1');
  span.setAttribute('contenteditable','plaintext-only');
  if(span.contentEditable!=='plaintext-only')span.setAttribute('contenteditable','true');
  span.style.outline='2px dashed rgba(255,255,255,0.7)';
  span.style.outlineOffset='4px';
  span.style.cursor='text';
  span.style.pointerEvents='auto';
  elNode.style.cursor='text';
  elNode.style.overflow='visible';
  // Place caret at end
  const focusEdit=()=>{
    span.focus();
    try{
      const range=document.createRange();
      range.selectNodeContents(span);
      range.collapse(false);
      const sel=window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }catch{}
  };
  requestAnimationFrame(focusEdit);
  setTimeout(focusEdit,30);
  let done=false;
  const commit=()=>{
    if(done)return;done=true;
    span.removeAttribute('data-editing');
    elNode.removeAttribute('data-editing');
    span.removeAttribute('contenteditable');
    span.style.outline='';
    span.style.outlineOffset='';
    span.style.pointerEvents='none';
    const next=(span.innerText||span.textContent||'').replace(/\r\n/g,'\n').replace(/\r/g,'\n');
    // Keep a single trailing newline behaviour consistent with lyric editor
    const cleaned=next.replace(/\n+$/,'');
    if(cleaned!==(el.tx||'')){
      pushUndo();
      el.tx=cleaned;
      // Keep legacy tx field in sync for anything still reading it
      const part=sls(A)[aj];
      if(part&&part.elements&&part.elements.length===1&&part.elements[0].id===el.id){
        part.tx=cleaned;
      }
    }else{
      // restore visual text in case browser altered whitespace
      span.textContent=el.tx||'';
    }
    saveSoon();
    try{updateCurrentThumb()}catch{}
    try{rPanel()}catch{}
    try{rStage()}catch{}
  };
  const onKey=(ke)=>{
    if(ke.key==='Escape'){ke.preventDefault();ke.stopPropagation();span.blur()}
    // Ctrl/Cmd+Enter commits
    if(ke.key==='Enter'&&(ke.ctrlKey||ke.metaKey)){ke.preventDefault();span.blur()}
  };
  const onBlur=()=>setTimeout(commit,80);
  span.addEventListener('keydown',onKey);
  span.addEventListener('blur',onBlur,{once:true});
}

// Double-click for text edit is handled in stage pointerup only (not here —
// a click listener would double-count the same physical click after pointerup).
let _elClickT=0,_elClickId='';

document.addEventListener('dblclick',ev=>{
  if(live||document.body.classList.contains('presentation-mode')||document.body.classList.contains('presenter-mode'))return;
  if(ev.target.closest('#el-inline-edit'))return;
  const stage=$('#stage');
  if(!stage||!stage.contains(ev.target))return;
  if(A&&isSongItem(A)){
    if(ed)return;
    startEdit();
    return;
  }
  if(!A||isSongItem(A))return;
  const part=sls(A)[aj];
  if(part&&typeof ensureElements==='function')ensureElements(part,A);
  const elNode=ev.target.closest('#stage [data-el]');
  if(elNode){
    const id=elNode.dataset.el;
    const el=(part&&part.elements||[]).find(x=>x.id===id);
    if(el&&el.kind==='text'){
      ev.preventDefault();
      openTextElEditor(el, elNode);
      return;
    }
  }
  const texts=(part&&part.elements||[]).filter(x=>x&&x.kind==='text');
  if(texts.length){
    const primary=texts.slice().sort((a,b)=>(b.z||0)-(a.z||0))[0];
    const node=document.querySelector('#stage [data-el="'+primary.id+'"]');
    ev.preventDefault();
    openTextElEditor(primary,node);
  }
}, true);

$('#stage').addEventListener('mousedown',ev=>{
  if(!ed)return;
  const c=$('#ce');
  if(c&&!ev.target.closest('#ce')&&ev.target.closest('.tx')){ev.preventDefault();c.focus();caretToEnd(c)}
});

/* =====================================================================
   Added features
   ===================================================================== */

/* ===== theme ===== */
const themeMQ=matchMedia('(prefers-color-scheme: dark)');
function applyTheme(mode){
  const dark=mode==='dark'||(mode!=='light'&&themeMQ.matches);
  document.documentElement.dataset.theme=dark?'dark':'light';
  document.documentElement.style.colorScheme=dark?'dark':'light';
}
themeMQ.addEventListener('change',()=>{if(($('#theme-setting').value||'system')==='system')applyTheme('system')});
$('#theme-setting').addEventListener('change',ev=>{applyTheme(ev.target.value);saveState()});
applyTheme($('#theme-setting').value||'system');

/* ===== persistent media (IndexedDB) ===== */
const mdb=new Promise(res=>{try{const r=indexedDB.open('service-slides-media',2);r.onupgradeneeded=()=>{const db=r.result;if(!db.objectStoreNames.contains('m'))db.createObjectStore('m');if(!db.objectStoreNames.contains('files'))db.createObjectStore('files')};r.onsuccess=()=>res(r.result);r.onerror=()=>res(null)}catch{res(null)}});
const thumbDb=new Promise(res=>{try{const r=indexedDB.open('service-slides-thumbnails',1);r.onupgradeneeded=()=>{if(!r.result.objectStoreNames.contains('thumbs'))r.result.createObjectStore('thumbs')};r.onsuccess=()=>res(r.result);r.onerror=()=>res(null)}catch{res(null)}});
async function mPutFile(id,blob){const d=await mdb;if(!d)return;return new Promise(r=>{const t=d.transaction('files','readwrite');t.objectStore('files').put(blob,id);t.oncomplete=r;t.onerror=r})}
async function mGetFile(id){const d=await mdb;if(!d)return null;return new Promise(r=>{const q=d.transaction('files').objectStore('files').get(id);q.onsuccess=()=>r(q.result||null);q.onerror=()=>r(null)})}
async function mDeleteFile(id){const d=await mdb;if(!d)return;return new Promise(r=>{const t=d.transaction('files','readwrite');t.objectStore('files').delete(id);t.oncomplete=r;t.onerror=r})}
async function mPut(id,blob){const d=await mdb;if(!d)return;return new Promise(r=>{const t=d.transaction('m','readwrite');t.objectStore('m').put(blob,id);t.oncomplete=r;t.onerror=r})}
async function mGet(id){const d=await mdb;if(!d)return null;return new Promise(r=>{const q=d.transaction('m').objectStore('m').get(id);q.onsuccess=()=>r(q.result||null);q.onerror=()=>r(null)})}
async function mKeys(){const d=await mdb;if(!d)return [];return new Promise(r=>{const q=d.transaction('m').objectStore('m').getAllKeys();q.onsuccess=()=>r(q.result||[]);q.onerror=()=>r([])})}
async function mDelete(id){const d=await mdb;if(!d)return;return new Promise(r=>{const t=d.transaction('m','readwrite');t.objectStore('m').delete(id);t.oncomplete=r;t.onerror=r})}
async function mPutThumb(id,blob){const d=await thumbDb;if(!d)return;return new Promise(r=>{const t=d.transaction('thumbs','readwrite');t.objectStore('thumbs').put(blob,id);t.oncomplete=r;t.onerror=r})}
async function mGetThumb(id){const d=await thumbDb;if(!d)return null;return new Promise(r=>{const q=d.transaction('thumbs').objectStore('thumbs').get(id);q.onsuccess=()=>r(q.result||null);q.onerror=()=>r(null)})}
async function mKeysThumb(){const d=await thumbDb;if(!d)return [];return new Promise(r=>{const q=d.transaction('thumbs').objectStore('thumbs').getAllKeys();q.onsuccess=()=>r(q.result||[]);q.onerror=()=>r([])})}
let mediaMeta=[];
try{mediaMeta=JSON.parse(localStorage.getItem('service-slides-media-meta')||'[]')}catch{mediaMeta=[]}
function saveMediaMeta(){localStorage.setItem('service-slides-media-meta',JSON.stringify(mediaMeta))}
function mediaMetaFor(id){
  let m=mediaMeta.find(x=>x.id===id);
  if(!m)m={id,name:'Background',kind:'image',createdAt:Date.now(),size:0},mediaMeta.push(m);
  return m;
}
function mediaReferences(id){
  const uses=[];
  const scan=(label,slides)=>(slides||[]).forEach(it=>{
    let hit=it.bgMedia&&it.bgMedia.id===id;
    (it.parts||[]).forEach(p=>{
      if((p.bgMedia&&p.bgMedia.id===id)||(p.videoMedia&&p.videoMedia.id===id))hit=true;
      (p.im||[]).forEach(im=>{if(im.mediaId===id)hit=true});
      (p.elements||[]).forEach(el=>{if(el.mediaId===id)hit=true});
    });
    if(hit)uses.push(label+' - '+it.t);
  });
  services.forEach((s,i)=>scan(formatServiceTitle(s),i===activeService?S:s.slides));
  return uses;
}
function purgeMedia(slides,id){
  (slides||[]).forEach(it=>{
    if(it.bgMedia&&it.bgMedia.id===id)delete it.bgMedia;
    (it.parts||[]).forEach(p=>{
      if(p.bgMedia&&p.bgMedia.id===id)delete p.bgMedia;
      if(p.videoMedia&&p.videoMedia.id===id)delete p.videoMedia;
      if(Array.isArray(p.im))p.im=p.im.filter(im=>im.mediaId!==id);
      if(Array.isArray(p.elements))p.elements=p.elements.filter(el=>el.mediaId!==id);
    });
  });
}
async function deleteMedia(id){
  const m=mediaMeta.find(x=>x.id===id);
  const name=(m&&m.name)||'this media';
  const uses=mediaReferences(id);
  const msg=uses.length
    ?'Delete "'+name+'"? It is used in:\n- '+uses.join('\n- ')+'\n\nIt will be removed from those slides.'
    :'Delete "'+name+'" from the media library?';
  if(!await appConfirm(msg,{title:'Delete media',confirmLabel:'Delete',danger:true}))return;
  services.forEach(s=>purgeMedia(s.slides,id));
  purgeMedia(S,id);
  try{if(vpool[id]){vpool[id].pause();delete vpool[id]}}catch{}
  if(mediaURLs[id]){try{URL.revokeObjectURL(mediaURLs[id])}catch{}delete mediaURLs[id]}
  try{await mDelete(id)}catch(err){console.warn('mDelete',err)}
  mediaMeta=mediaMeta.filter(x=>x.id!==id);
  saveMediaMeta();
  try{renderMediaLibrary()}catch{}
  rAll();
  try{renderBackgroundPicker()}catch{}
  toast('Deleted.');
}
function mediaUsage(id){
  const uses=[];
  services.forEach(s=>(s.slides||[]).forEach(it=>{if(it.bgMedia?.id===id)uses.push(s.name+' — '+it.t)}));
  return uses;
}
async function loadMedia(){
  const backgroundThumbnailIds=new Set(mediaMeta.filter(m=>m.source==='church-library'&&m.thumbnailId).map(m=>m.thumbnailId));
  const allKeys=await mKeys();
  // One-time migration from the old shared media store. Thumbnail blobs live in
  // their own IndexedDB store from here on, so they never become media inventory.
  for(const id of allKeys.filter(key=>String(key).startsWith('thumb-'))){
    const blob=await mGet(id);
    if(blob)await mPutThumb(id,blob);
    await mDelete(id);
  }
  for(const id of await mKeysThumb()){
    const blob=await mGetThumb(id);
    if(blob&&!mediaURLs[id])mediaURLs[id]=URL.createObjectURL(blob);
  }
  const ids=new Set(allKeys.filter(id=>!String(id).startsWith('thumb-')));
  services.forEach(s=>(s.slides||[]).forEach(it=>{if(it.bgMedia?.id)ids.add(it.bgMedia.id);(it.parts||[]).forEach(p=>{if(p.bgMedia?.id)ids.add(p.bgMedia.id);if(p.videoMedia?.id)ids.add(p.videoMedia.id);(p.im||[]).forEach(im=>{if(im.mediaId)ids.add(im.mediaId)})})}));
  S.forEach(it=>{if(it.bgMedia?.id)ids.add(it.bgMedia.id);(it.parts||[]).forEach(p=>{if(p.bgMedia?.id)ids.add(p.bgMedia.id);if(p.videoMedia?.id)ids.add(p.videoMedia.id);(p.im||[]).forEach(im=>{if(im.mediaId)ids.add(im.mediaId)})})});
  for(const id of ids){
    const b=await mGet(id);
    if(b&&!mediaURLs[id])mediaURLs[id]=URL.createObjectURL(b);
    const applySrc=(it)=>{
      (it.parts||[]).forEach(p=>{
        (p.im||[]).forEach(im=>{if(im.mediaId===id)im.src=mediaURLs[id]||im.src||''});
        (p.elements||[]).forEach(el=>{if(el.kind==='image'&&el.mediaId===id)el.src=mediaURLs[id]||el.src||''});
        if(p.bgMedia?.id===id){/* src resolved at render via mediaURLs */}
      });
      if(it.bgMedia?.id===id){/* resolved at render */}
    };
    services.forEach(s=>(s.slides||[]).forEach(applySrc));
    S.forEach(applySrc);
    if(!backgroundThumbnailIds.has(id))mediaMetaFor(id);
  }
  mediaMeta=mediaMeta.filter(m=>ids.has(m.id)&&!backgroundThumbnailIds.has(m.id)&&!String(m.id).startsWith('thumb-'));
  saveMediaMeta();
  saveState();
  if(!document.body.classList.contains('home-mode')){
    rStrip();rStage();
  }
}
function openBackgroundPicker(){
  const modal=$('#background-picker-modal');
  if(!modal)return;
  modal.hidden=false;
  renderBackgroundPicker();
}
function closeBackgroundPicker(){
  const modal=$('#background-picker-modal');
  if(modal)modal.hidden=true;
}
function renderBackgroundPicker(){
  const root=$('#background-picker-content');
  if(!root)return;
  const active=root.dataset.backgroundTab||'images';
  const driveItems=mediaMeta.filter(m=>!m.libraryRemoved&&m.source==='church-library'&&m.libraryType==='church-background').map(m=>({...m,downloadedItem:mediaMeta.find(d=>d.id===m.downloadedItemId&&!d.libraryRemoved)||mediaMeta.find(d=>!d.libraryRemoved&&d.source==='downloaded'&&(d.driveFileId===m.driveFileId||d.downloadedFromDriveId===m.driveFileId))})).sort((a,b)=>a.name.localeCompare(b.name));
  const announcementDriveIds=new Set(mediaMeta.filter(m=>m.libraryType==='church-announcement'&&m.driveFileId).map(m=>m.driveFileId));
  const items=mediaMeta
    .filter(m=>!m.libraryRemoved&&m.source!=='powerpoint'&&m.source!=='church-library'&&m.kind===active.slice(0,-1)&&m.libraryType!=='church-announcement'&&m.libraryType!=='church-logo'&&!(m.downloadedFromDriveId&&announcementDriveIds.has(m.downloadedFromDriveId)))
    .sort((a,b)=>(b.createdAt||0)-(a.createdAt||0));
  root.innerHTML=`
    <div class="background-picker-tabs">
      <button type="button" data-background-tab="images" class="${active==='images'?'on':''}">Images</button>
      <button type="button" data-background-tab="videos" class="${active==='videos'?'on':''}">Videos</button>
      <button type="button" data-background-tab="drive" class="${active==='drive'?'on':''}">From drive</button>
      <button type="button" id="background-upload-btn">Upload</button>
      <button type="button" data-background-sync="1" style="margin-left:auto">↻ Sync backgrounds</button>
    </div>
    <input type="file" id="background-upload-file" accept="image/*,video/*" multiple hidden>
    <p class="note" style="margin:0 0 12px">Upload images or videos here, or drag and drop them into this window. Uploads stay on this computer.</p>
    <div class="background-picker-grid">
      ${(active==='drive'?driveItems:items).length?(active==='drive'?driveItems:items).map(m=>`
        <div class="background-picker-item">
          ${m.kind==='video'
            ?`<div style="width:100%;height:100%;display:grid;place-items:center;position:relative"><img src="${e(mediaURLs[m.thumbnailId]||'')}" alt="" style="width:100%;height:100%;object-fit:cover"><span style="position:absolute;padding:4px 8px;background:#0009;color:white;border-radius:6px">▶ Video</span></div>`
            :`<img src="${e(mediaURLs[m.thumbnailId]||mediaURLs[m.id]||'')}" alt="">`}
          <div class="background-picker-name" title="${e(m.name)}">${e(m.name)}</div>
          <div class="background-picker-select">${active==='drive'?(m.downloadedItem?`<button type="button" data-background-select="${e(m.downloadedItem.id)}">Select</button>`:`<button type="button" data-background-download="${e(m.id)}">Download</button>`):`<button type="button" data-background-select="${e(m.id)}">Select</button><button type="button" class="background-picker-delete" data-background-delete="${e(m.id)}" title="Delete from media library" aria-label="Delete from media library">Delete</button>`}</div>
        </div>`).join('')
        :`<div class="media-empty" style="grid-column:1/-1">No ${active==='drive'?'backgrounds in the Drive library':active+' uploaded yet.'}</div>`}
    </div>`;
}
async function addBackgroundFiles(files){
  const accepted=[...files].filter(file=>/^(image|video)\//i.test(file.type||''));
  if(!accepted.length)return toast('Choose image or video files.');
  for(const file of accepted){
    const id='m'+Date.now().toString(36)+Math.random().toString(36).slice(2,8);
    await mPut(id,file);mediaURLs[id]=URL.createObjectURL(file);
    mediaMeta.push({id,name:file.name,kind:file.type.startsWith('video/')?'video':'image',size:file.size,createdAt:Date.now(),source:'local-background'});
  }
  saveMediaMeta();renderBackgroundPicker();toast('Added '+accepted.length+' local background file'+(accepted.length===1?'':'s')+'.');
}
function selectSongBackground(id){
  const item=mediaMeta.find(m=>m.id===id);
  if(!item||!S[sel])return;
  closeBackgroundPicker();
  const tg=bgTargets();
  bgApply(o=>{o.bgMedia={id:item.id,kind:item.kind,name:item.name}});
  toast(tg&&tg.section?'Background applied to all slides in this section.':'Background applied to the selected slide'+(tg&&tg.parts.length>1?'s':'')+'.');
}
function openMediaLibrary(){
  const m=$('#media-modal');
  if(!m)return;
  m.hidden=false;
  try{renderMediaLibrary()}catch(err){console.error(err)}
}
function closeMediaLibrary(){
  const m=$('#media-modal');
  if(m)m.hidden=true;
}
function renderMediaLibrary(){
  const root=$('#media-content');if(!root)return;
  const active=root.dataset.mediaTab||'images';
  const groups=new Map();
  mediaMeta.forEach(m=>{
    if(m.libraryRemoved)return;
    if(m.presentationId){
      const key=m.presentationId||m.id;
      if(!groups.has(key))groups.set(key,{name:m.presentationName||m.name.split(' — Slide ')[0],items:[]});
      groups.get(key).items.push(m);
    }
  });
  const pptGroups=[...groups.values()].sort((a,b)=>a.name.localeCompare(b.name));
  const backgroundIds=new Set();
  const collectBackgrounds=slides=>(slides||[]).forEach(it=>{if(it.bgMedia?.id)backgroundIds.add(it.bgMedia.id);(it.parts||[]).forEach(p=>{if(p.bgMedia?.id)backgroundIds.add(p.bgMedia.id)})});
  services.forEach(s=>collectBackgrounds(s.slides));collectBackgrounds(S);
  let taggedBackground=false;
  backgroundIds.forEach(id=>{const meta=mediaMeta.find(m=>m.id===id);if(meta&&!meta.source){meta.source='background';taggedBackground=true}});
  if(taggedBackground)saveMediaMeta();
  const images=mediaMeta.filter(m=>!m.libraryRemoved&&m.kind==='image'&&!m.presentationId&&m.source!=='powerpoint'&&m.source!=='pdf'&&m.source!=='church-library'&&m.source!=='background'&&m.source!=='local-background'&&m.source!=='downloaded'&&!backgroundIds.has(m.id)&&m.libraryType!=='church-announcement'&&!(m.downloadedFromDriveId&&mediaMeta.some(a=>a.libraryType==='church-announcement'&&a.driveFileId===m.downloadedFromDriveId))).sort((a,b)=>(b.createdAt||0)-(a.createdAt||0));
  const announcements=mediaMeta.filter(m=>!m.libraryRemoved&&m.libraryType==='church-announcement'&&m.source==='church-library').map(m=>({...m,downloadedItemId:m.downloadedItemId||mediaMeta.find(d=>d.downloadedFromDriveId===m.driveFileId)?.id})).sort((a,b)=>a.name.localeCompare(b.name));
  const videos=mediaMeta.filter(m=>!m.libraryRemoved&&m.kind==='video'&&m.source!=='church-library'&&m.source!=='background'&&m.source!=='local-background'&&m.source!=='downloaded'&&!backgroundIds.has(m.id)).sort((a,b)=>(b.createdAt||0)-(a.createdAt||0));
  const uploadRow=active==='powerpoints'
    ?`<div style="display:flex;gap:8px;align-items:center;margin-bottom:12px;flex:none"><button id="media-upload-slides-btn" class="pri">Upload slides</button><input id="media-slides-file" type="file" accept=".pptx,.ppt,.pdf,application/pdf" hidden></div>`
    :`<div style="display:flex;gap:8px;align-items:center;margin-bottom:12px;flex:none"><button id="media-upload-btn" class="pri">Upload image or video</button><input id="media-file" type="file" accept="image/*,video/*" hidden></div>`;
  const tabs=`<p class="note" style="margin:0 0 10px;flex:none">Images, videos, and imported slide decks used across your services.</p>
  <div class="seg" style="margin-bottom:12px;flex:none"><button data-media-tab="images" class="${active==='images'?'on':''}">Images</button><button data-media-tab="videos" class="${active==='videos'?'on':''}">Videos</button><button data-media-tab="powerpoints" class="${active==='powerpoints'?'on':''}">PowerPoints and PDFs</button><button data-media-tab="announcements" class="${active==='announcements'?'on':''}">Church Announcements</button></div>${active==='announcements'?`<div style="display:flex;align-items:center;gap:10px"><p class="note" style="flex:1">Approved announcement slides from the church Drive library.</p>${announcementsHaveSynced&&announcements.length?'<button type="button" class="pri" data-announcement-download-all="1">Download all</button>':''}<button type="button" data-manual-sync="1">↻ Sync</button></div>`:uploadRow}`;
  const actions=(btns)=>`<div class="media-actions">${btns}</div>`;
  let body='';
  if(active==='images'){
    body=images.length?images.map(m=>`<div class="media-item">
      <img class="media-thumb" src="${e(mediaURLs[m.id]||'')}" alt="">
      <div class="media-info"><b title="${e(m.name)}">${e(m.name)}</b></div>
      ${actions(`<button data-media-view="${e(m.id)}">View</button><button data-media-rename="${e(m.id)}">Rename</button><button data-media-add-slide="${e(m.id)}" class="pri">Add</button><button data-media-delete="${e(m.id)}" class="media-del" title="Delete" aria-label="Delete">✕</button>`)}
    </div>`).join(''):'<div class="media-empty">No images.</div>';
  }else if(active==='videos'){
    body=videos.length?videos.map(m=>`<div class="media-item">
      <video class="media-thumb" src="${e(mediaURLs[m.id]||'')}" muted preload="metadata"></video>
      <div class="media-info"><b title="${e(m.name)}">${e(m.name)}</b></div>
      ${actions(`<button data-media-view="${e(m.id)}">View</button><button data-media-rename="${e(m.id)}">Rename</button><button data-media-add-slide="${e(m.id)}" class="pri">Add</button><button data-media-delete="${e(m.id)}" class="media-del" title="Delete" aria-label="Delete">✕</button>`)}
    </div>`).join(''):'<div class="media-empty">No videos.</div>';
  }else if(active==='announcements'){
    body=announcements.length?announcements.map(m=>`<div class="media-item"><img class="media-thumb announcement-thumb" src="${e(mediaURLs[m.thumbnailId]||mediaURLs[m.id]||'')}" alt=""><div class="media-info"><b title="${e(m.name)}">${e(m.name)}</b></div>${actions(`<button data-media-view="${e(m.id)}">View</button>${m.downloadedItemId?`<button data-media-add-announcement="${e(m.downloadedItemId)}" class="pri">Add slide</button>`:`<button data-announcement-download="${e(m.id)}" class="pri">Download</button>`}`)}</div>`).join(''):'<div class="media-empty">No church announcements are available. Sync with Drive to refresh this library.</div>';
  }else{
    body=pptGroups.length?pptGroups.map(g=>{
      const first=g.items[0];
      const thumb=mediaURLs[first.id]||'';
      return `<div class="media-item">
        ${thumb?`<img class="media-thumb" src="${e(thumb)}" alt="">`:`<div class="media-thumb" style="display:grid;place-items:center;font-size:calc(28px * var(--ui-scale));color:#fff">▦</div>`}
        <div class="media-info"><b title="${e(g.name)}">${e(g.name)}</b><small>${g.items.length} slide${g.items.length===1?'':'s'}</small></div>
        ${actions(`<button data-media-view-ppt="${e(first.presentationId||first.id)}">View</button><button data-media-rename="${e(first.id)}">Rename</button><button data-media-add-ppt="${e(first.presentationId||first.id)}" class="pri">Add</button><button data-media-delete-ppt="${e(first.presentationId||first.id)}" class="media-del" title="Delete" aria-label="Delete">✕</button>`)}
      </div>`;
    }).join(''):'<div class="media-empty">No imported slide decks.</div>';
  }
  root.innerHTML=tabs+`<div class="media-grid">${body}</div>`;
}

function openMediaPreview(id, presentationId=null){
  const item=mediaMeta.find(m=>m.id===id);if(!item)return;
  const slides=presentationId?mediaMeta.filter(m=>m.presentationId===presentationId).sort((a,b)=>(a.slideNumber||0)-(b.slideNumber||0)):mediaMeta.filter(m=>m.kind===item.kind&&!m.presentationId&&m.libraryType===item.libraryType&&m.source===item.source&&!m.libraryRemoved).sort((a,b)=>a.name.localeCompare(b.name));
  let index=Math.max(0,slides.findIndex(m=>m.id===id));
  const layer=document.createElement('div');layer.className='media-preview-overlay';
  layer.style.cssText='position:fixed;inset:0;z-index:2147483646;background:#000d;display:flex;align-items:center;justify-content:center;padding:24px';
  layer.innerHTML=`<section role="dialog" aria-modal="true" aria-label="Media preview" style="width:min(1200px,96vw);height:min(90vh,900px);background:var(--pn,#fff);color:var(--ink,#111);border-radius:12px;display:flex;flex-direction:column;overflow:hidden"><header style="display:flex;align-items:center;gap:12px;padding:12px 16px;border-bottom:1px solid var(--ln,#ddd)"><b data-preview-title style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"></b><button type="button" data-preview-prev>Previous</button><span data-preview-count></span><button type="button" data-preview-next>Next</button><button type="button" data-preview-close>Close</button><button type="button" data-preview-add class="pri">Add</button></header><div data-preview-stage style="flex:1;min-height:0;display:grid;place-items:center;padding:18px;background:#161616"></div></section>`;
  document.body.appendChild(layer);
  const stage=layer.querySelector('[data-preview-stage]');
  function show(){const current=slides[index],downloaded=mediaMeta.find(m=>current.libraryType==='church-announcement'&&m.downloadedFromDriveId===current.driveFileId),url=mediaURLs[downloaded?.id||current.id]||mediaURLs[current.thumbnailId]||'';layer.querySelector('[data-preview-title]').textContent=current.presentationName||current.name;layer.querySelector('[data-preview-count]').textContent=slides.length>1?`${index+1} / ${slides.length}`:'';layer.querySelector('[data-preview-prev]').disabled=index===0;layer.querySelector('[data-preview-next]').disabled=index===slides.length-1;stage.innerHTML='';if(current.kind==='video'){const v=document.createElement('video');v.src=url;v.controls=true;v.autoplay=false;v.style.cssText='max-width:100%;max-height:100%;object-fit:contain';stage.appendChild(v)}else if(url){const img=document.createElement('img');img.src=url;img.alt=current.presentationName||current.name;img.style.cssText='max-width:100%;max-height:100%;object-fit:contain';stage.appendChild(img)}else stage.textContent='Preview unavailable';}
  function close(){layer.remove();document.removeEventListener('keydown',onKey)}
  function onKey(ev){if(ev.key==='Escape')close();else if(ev.key==='ArrowRight'&&index<slides.length-1){index++;show()}else if(ev.key==='ArrowLeft'&&index>0){index--;show()}}
  layer.addEventListener('click',async ev=>{const b=ev.target.closest('button');if(!b)return;if(b.hasAttribute('data-preview-close'))close();else if(b.hasAttribute('data-preview-add')){const current=slides[index];let downloaded=mediaMeta.find(m=>current.libraryType==='church-announcement'&&m.downloadedFromDriveId===current.driveFileId);if(presentationId)addPowerPointGroup(presentationId);else if(current.libraryType==='church-announcement'){if(!downloaded){await downloadAnnouncement(current.id);downloaded=mediaMeta.find(m=>m.downloadedFromDriveId===current.driveFileId)}if(downloaded)addAnnouncementSlide(downloaded.id)}else if(item.kind==='video'||item.kind==='image')addMediaAsSection(current.id);else useMedia(current.id)}else if(b.hasAttribute('data-preview-prev')&&index>0){index--;show()}else if(b.hasAttribute('data-preview-next')&&index<slides.length-1){index++;show()}});
  layer.addEventListener('click',ev=>{if(ev.target===layer)close()});document.addEventListener('keydown',onKey);show();
 }


async function renameMedia(id){
  const m=mediaMetaFor(id);if(!m)return;
  const name=await appPrompt('Rename media','Enter a new name for this media item.',m.name);if(name==null)return;
  const clean=name.trim();if(!clean)return;
  if(m.source==='powerpoint'&&m.presentationId){mediaMeta.filter(x=>x.presentationId===m.presentationId).forEach(x=>{x.presentationName=clean;x.name=`${clean} — Slide ${x.slideNumber||1}`});}
  else m.name=clean;
  saveMediaMeta();renderMediaLibrary();saveState();toast('Renamed.');
}
async function deletePowerPointGroup(key){
  const items=mediaMeta.filter(m=>m.presentationId===key);if(!items.length)return;
  if(!await appConfirm(`Delete the PowerPoint presentation "${items[0].presentationName||items[0].name}" from the media library? This will also remove its imported slides from the service.`,{title:'Delete presentation',confirmLabel:'Delete',danger:true}))return;
  const ids=new Set(items.map(m=>m.id));
  services.forEach(svc=>(svc.slides||[]).forEach(it=>{it.parts=(it.parts||[]).filter(p=>!(p.im||[]).some(im=>ids.has(im.mediaId)));it.ord=(it.ord||[]).filter(id=>it.parts.some(p=>p.id===id))}));
  for(const m of items){if(mediaURLs[m.id]){URL.revokeObjectURL(mediaURLs[m.id]);delete mediaURLs[m.id]}await mDelete(m.id)}
  mediaMeta=mediaMeta.filter(m=>m.presentationId!==key);saveMediaMeta();S=cl(services[activeService].slides);sel=Math.min(sel,S.length-1);A=S[sel];aj=0;renderMediaLibrary();rAll();
}

function addPowerPointGroup(key){
  const items=mediaMeta.filter(m=>m.presentationId===key).sort((a,b)=>(a.slideNumber||0)-(b.slideNumber||0));
  if(!items.length)return toast('PowerPoint presentation not found.');
  const name=items[0].presentationName||items[0].name.split(' — Slide ')[0];
  const parts=items.map((m,i)=>({id:uid(),n:`Slide ${i+1}`,tx:'',im:[{full:true,src:mediaURLs[m.id]||'',mediaId:m.id}]}));
  const item=mk(name,items[0].presentationKind||'PowerPoint',parts,parts.map(p=>p.id),0,5.2);
  item.sourceFile=name;item.presentationId=key;item.importedAt=Date.now();
  item.originalFileId=items[0].originalFileId||'';
  S.splice(sel+1,0,item);sel++;si=0;A=item;aj=0;tab='Structure';rAll();
  toast(`Added ${name} to the service.`);
}

function useMedia(id){
  addMediaAsSection(id);
}
function driveMediaFingerprint(blob){return `${blob.size}:${blob.type||''}:${blob.lastModified||0}`}
async function rootBackgroundDownload(item){
  if(!item?.driveFileId)return toast('Drive details are missing. Sync the From drive list again.');
  const downloaded=await withCloudLock('Downloading '+item.name+'...',async()=>{
    const file=await downloadDriveFile({id:item.driveFileId,name:item.name,mimeType:item.mimeType||''});
    const blob=new Blob([await file.arrayBuffer()],{type:file.type||item.mimeType||mimeOf(item.name)});
    const id='download-'+uid();await mPut(id,blob);const sourceItem=mediaMeta.find(m=>m.id===item.id);if(sourceItem)sourceItem.downloadedItemId=id;mediaMeta.push({...item,id,source:'downloaded',downloadedFromDriveId:item.driveFileId,libraryType:'church-background',size:blob.size,driveMediaFingerprint:driveMediaFingerprint(blob),createdAt:Date.now()});mediaURLs[id]=URL.createObjectURL(blob);saveMediaMeta();
  },{progress:true});
  if(downloaded){renderBackgroundPicker();toast('Background downloaded for local use.');}
}
async function addAnnouncementSlide(id){
  const meta=mediaMeta.find(x=>x.id===id),src=mediaURLs[id];if(!meta||!src)return toast('Announcement image is not available. Sync with Drive and try again.');
  let colour='#101820';
  try{const img=await new Promise((resolve,reject)=>{const x=new Image();x.onload=()=>resolve(x);x.onerror=reject;x.src=src});const c=document.createElement('canvas');c.width=8;c.height=8;const ctx=c.getContext('2d',{willReadFrequently:true});ctx.drawImage(img,0,0,8,8);const d=ctx.getImageData(0,0,8,8).data;let r=0,g=0,b=0,n=0;for(const [x,y] of [[0,0],[7,0],[0,7],[7,7],[4,0],[4,7],[0,4],[7,4]]){const p=(y*8+x)*4;r+=d[p];g+=d[p+1];b+=d[p+2];n++}colour='#'+[r,g,b].map(v=>Math.round(v/n).toString(16).padStart(2,'0')).join('')}catch{}
  pushUndo();const part={id:uid(),n:'Announcement',tx:'',im:[{full:true,fit:'contain',src,mediaId:id}]};
  const target=S[sel];
  if(target&&!isSongItem(target)){
    target.parts.push(part);target.ord.push(part.id);A=target;aj=target.ord.length-1;
  }else{
    const item=mk('Announcement: '+meta.name.replace(/\.[^.]+$/,''),'Slide',[part],[part.id],5,5.2);item.customBg=colour;item.dim=0;
    S.splice(Math.max(-1,sel)+1,0,item);sel=S.indexOf(item);A=item;aj=0;
  }
  si=0;tab='Structure';rAll();toast('Added announcement slide to the current section.');
}
async function downloadAnnouncement(id){
  const item=mediaMeta.find(x=>x.id===id);if(!item?.driveFileId)return toast('Drive details are missing. Sync announcements again.');
  await withCloudLock('Downloading '+item.name+'…',async()=>{await downloadAnnouncementOriginal(item);renderMediaLibrary()},{progress:true});
}
async function downloadAnnouncementOriginal(item){
  // Fetch the Drive file bytes (alt=media), never its thumbnail, so the local
  // copy retains the original image resolution and quality.
  const file=await downloadDriveFile({id:item.driveFileId,name:item.name,mimeType:item.mimeType||''});
  const blob=new Blob([await file.arrayBuffer()],{type:file.type||item.mimeType||mimeOf(item.name)}),localId='download-'+uid();
  await mPut(localId,blob);mediaURLs[localId]=URL.createObjectURL(blob);mediaMeta.push({...item,id:localId,source:'downloaded',downloadedFromDriveId:item.driveFileId,thumbnailId:undefined,size:blob.size,driveMediaFingerprint:driveMediaFingerprint(blob),createdAt:Date.now()});item.downloadedItemId=localId;saveMediaMeta();
}
async function downloadAllAnnouncements(){
  const items=mediaMeta.filter(m=>!m.libraryRemoved&&m.libraryType==='church-announcement'&&m.source==='church-library'&&!mediaMeta.some(d=>d.source==='downloaded'&&d.downloadedFromDriveId===m.driveFileId));
  if(!items.length)return toast('All announcement slides are already downloaded.');
  await withCloudLock('Downloading announcement slides…',async()=>{
    let completed=0;
    for(const item of items){try{await downloadAnnouncementOriginal(item);completed++}catch(err){console.warn('[announcement download]',item.name,err)}}
    renderMediaLibrary();toast(`Downloaded ${completed} of ${items.length} announcement slide${items.length===1?'':'s'}.`);
  },{progress:true});
}
/** Add a new service section where the media fills the slide (scale to fit, never stretch). */
function addMediaAsSection(id){
  const m=mediaMetaFor(id);if(!m)return toast('Media not found.');
  pushUndo();
  const src=mediaURLs[m.id]||'';
  const title=(m.name||'Media').replace(/\.[^.]+$/,'');
  let part;
  if(m.kind==='video'){
    part={id:uid(),n:'Slide 1',tx:'',videoMedia:{id:m.id,kind:'video',name:m.name,settings:{autoplay:true,loop:true,muted:true,volume:1}}};
  }else{
    part={id:uid(),n:'Slide 1',tx:'',im:[{full:true,src,mediaId:m.id}]};
  }
  const item=mk(title,m.kind==='video'?'Video':'Slide',[part],[part.id],3,5.2);
  item.bg=3;
  S.splice(sel+1,0,item);sel++;si=0;A=item;aj=0;tab='Structure';
  closeMediaLibrary();rAll();
  setTimeout(()=>scrollNewItemToTop(sel),0);
  toast('Added “'+title+'? as a new section.');
}
/** Add media as a freeform element on the current slide (move/resize/rotate). */
function addMediaToCurrentSlide(id){
  const m=mediaMetaFor(id);if(!m)return toast('Media not found.');
  if(!A||!S.length)return toast('Open a service section first.');
  if(isSongItem(A))return toast('Cannot add freeform media onto a song lyric slide.');
  const part=sls(A)[aj];if(!part)return;
  pushUndo();
  if(m.kind==='video'){
    addVideoElToCurrentSlide(m.id,m.name);closeMediaLibrary();return;
    closeMediaLibrary();rAll();toast('Video added to this slide.');
    return;
  }
  const els=ensureElements(part,A);
  const src=mediaURLs[m.id]||'';
  const el={
    id:uid(),kind:'image',src,mediaId:m.id,
    x:20,y:20,w:40,h:40,rot:0,
    z:(els.reduce((mx,e)=>Math.max(mx,e.z||1),0)+1),
    fit:'contain'
  };
  els.push(el);elSel=el.id;tab='Structure';
  closeMediaLibrary();saveSoon();rStage();updateCurrentThumb();rPanel();
  toast('Image added to this slide — drag to move, handles to resize.');
}

/* ===== seamless background video ===== */
const vpool={};
function stillHtml(it,p,slideIndex=-1){
  let h=sl(it,p,false,false,slideIndex);
  h=h.replace(/<video\b[^>]*>[\s\S]*?<\/video>/gi,'');
  h=h.replace(/<div class="vslot"[^>]*>\s*<\/div>/gi,'<div style="position:absolute;inset:0;background:#15202b"></div>');
  return h;
}
function attachVideos(){
  if(live&&document.body.classList.contains('presenter-mode')){
    Object.values(vpool).forEach(v=>{try{v.pause()}catch{}});
    return;
  }

  const slot=document.querySelector('#stage .vslot');
  Object.entries(vpool).forEach(([k,v])=>{if(!slot||slot.dataset.vkey!==k)v.pause()});
  if(!slot||!slot.dataset.vsrc)return;
  const k=slot.dataset.vkey,src=slot.dataset.vsrc;
  const currentPart=(sls(A)||[])[aj],currentBackground=currentPart?.ownBg?currentPart.bgMedia:(currentPart?.bgMedia||A.bgMedia),playbackRate=Number(currentBackground?.playbackRate)||1;
  let v=vpool[k];
  if(!v){v=document.createElement('video');v.muted=true;v.loop=true;v.playsInline=true;v.autoplay=true;v.preload='auto';vpool[k]=v}
  if(v.dataset.src!==src){v.dataset.src=src;v.src=src}
  v.playbackRate=playbackRate;
  slot.appendChild(v);
  if(v.paused)v.play().catch(()=>{});
}

/* ===== song text helpers ===== */
function songToText(song,merge){
  const out=[];
  song.ord.forEach(id=>{
    const p=song.parts.find(x=>x.id==id);if(!p)return;
    const m=p.n.match(/^(.*?)\s*\((\d+)\/(\d+)\)$/),last=out[out.length-1];
    if(merge&&m&&+m[2]>1&&last&&last.base===m[1])last.lines.push(p.tx);
    else out.push({base:merge&&m?m[1]:p.n,lines:[p.tx]});
  });
  return out.map(o=>o.base+'\n'+o.lines.join('\n')).join('\n\n');
}
function curLps(it){
  if(it.lps!=null)return it.lps;
  const bp=it.parts.filter(p=>!p.auto);return bp.some(p=>/\(\d+\/\d+\)$/.test(p.n))?Math.max(...bp.map(p=>p.tx.split('\n').length)):0;
}
const lpsOptions=v=>[[1,'1 line per slide'],[2,'2 lines per slide'],[3,'3 lines per slide'],[4,'4 lines per slide'],[6,'6 lines per slide'],[8,'8 lines per slide'],[0,'Whole section on one slide']]
  .map(([n,l])=>`<option value="${n}" ${n==v?'selected':''}>${l}</option>`).join('');
function reflowItem(it,lps,expand=true){collapseFit(it);
  const parsed=reflowParts(it,lps);
  it.parts=parsed.parts;it.ord=parsed.ord;it.lps=lps;
  si=0;if(it===A)aj=0;if(expand)expandFit(it);
}
document.addEventListener('change',ev=>{
  const t=ev.target;
  if(['video-autoplay','video-loop','video-muted'].includes(t.id)){const p=sls(A)[aj];if(p?.videoMedia){p.videoMedia.settings={...(p.videoMedia.settings||{}),autoplay:$('#video-autoplay')?.checked||false,loop:$('#video-loop')?.checked||false,muted:!!$('#video-muted')?.checked};saveSoon();rStage();updateCurrentThumb()}return}
  if(t.id==='video-volume'){const p=sls(A)[aj];if(p?.videoMedia){const v=(+$('#video-volume').value)/100;p.videoMedia.settings={...(p.videoMedia.settings||{}),volume:v};const lab=$('#video-vol-val');if(lab)lab.textContent=Math.round(v*100);saveSoon();rStage()}return}
  if(t.id==='background-video-playback-speed'){const value=+t.value,speed=value<=50?.1+(value/50)*.9:1+((value-50)/50)*4,it=S[sel],p=sls(it)[si],own=!isSongItem(it)&&!!p?.ownBg,bg=own?p.bgMedia:(p?.bgMedia||it.bgMedia);if(bg?.kind==='video'){bg.playbackRate=speed;const lab=$('#background-video-speed-val');if(lab)lab.textContent=speed.toFixed(1).replace(/\.0$/,'')+'x';saveSoon();rStage()}return}
  if(t.id==='video-playback-speed'){const value=+t.value,speed=value<=50?.1+(value/50)*.9:1+((value-50)/50)*4,p=sls(A).find(x=>x.videoMedia);if(p?.videoMedia){p.videoMedia.settings={...(p.videoMedia.settings||{}),playbackRate:speed};const lab=$('#video-speed-val');if(lab)lab.textContent=speed.toFixed(1).replace(/\.0$/,'')+'x';saveSoon();rStage()}return}
  if(t.name==='ppt-view-mode'){const it=S[sel];if(it){it.viewMode=t.value;if(services[activeService])services[activeService].slides=sanitizeSlidesForSave(S);saveSoon()}return}
  /* __ppt_test_removed__ */
  if(t.id==='lps-select'){const it=S[sel];reflowItem(it,+ev.target.value);toast('Re-split into '+it.ord.length+' slides.');rAll()}
});

/* ===== media library / backgrounds ===== */
async function importMediaFile(file){const id='m'+Date.now().toString(36)+Math.random().toString(36).slice(2,8);await mPut(id,file);mediaURLs[id]=URL.createObjectURL(file);mediaMeta.push({id,name:file.name,kind:file.type.startsWith('video')?'video':'image',size:file.size,createdAt:Date.now()});saveMediaMeta();renderMediaLibrary();toast('Media added to the library.')}

document.addEventListener('change',async ev=>{
  const t=ev.target;
  if(t.id==='background-upload-file'){const files=[...t.files];t.value='';await addBackgroundFiles(files);return}
  if(t.id==='import-slides-file'||t.id==='media-slides-file'){
    const f=t.files[0];t.value='';
    if(f){
      const isCpres=/\.cpres$/i.test(f.name||'')||((f.type||'').includes('json')&&/\.cpres$/i.test(f.name||''));
      if(/^(image|video)\//.test(f.type||'')){
        await importMediaFile(f);
        const root=$('#media-content');if(root){root.dataset.mediaTab=f.type.startsWith('video')?'videos':'images';openMediaLibrary()}
      }else if(isCpres||/\.cpres$/i.test(f.name||'')){
        if(t.id==='import-slides-file')await importCpresMergeIntoCurrent(f);
        else await importCpresFile(f);
      }else{
        await importPowerPoint(f,{addToService:t.id!=='media-slides-file'});
        if(t.id==='media-slides-file')renderMediaLibrary();
      }
    }
    return;
  }
  if(t.id==='add-video-slide-file'){const f=t.files[0];t.value='';if(f)await addVideoSlideItem(f);return}
  if(t.id==='slide-video-file'){const f=t.files[0];t.value='';if(f)await addVideoToCurrentSlide(f);return}
  if(t.id!=='media-file'||!t.files[0])return;
  const f=t.files[0];
  if(!/^(image|video)\//.test(f.type)){t.value='';return toast('Choose an image or video file.')}
  const id='m'+Date.now().toString(36)+Math.random().toString(36).slice(2,8);
  await mPut(id,f);
  mediaURLs[id]=URL.createObjectURL(f);
  mediaMeta.push({id,name:f.name,kind:f.type.startsWith('video')?'video':'image',size:f.size,createdAt:Date.now()});
  saveMediaMeta();t.value='';renderMediaLibrary();toast('Media added to the library.');
});
document.addEventListener('dragover',ev=>{const modal=$('#background-picker-modal');if(modal&&!modal.hidden&&modal.contains(ev.target)){ev.preventDefault();modal.classList.add('is-drop-target')}});
document.addEventListener('dragleave',ev=>{const modal=$('#background-picker-modal');if(modal&&(!ev.relatedTarget||!modal.contains(ev.relatedTarget)))modal.classList.remove('is-drop-target')});
document.addEventListener('drop',async ev=>{const modal=$('#background-picker-modal');if(!modal||modal.hidden||!modal.contains(ev.target))return;ev.preventDefault();modal.classList.remove('is-drop-target');await addBackgroundFiles(ev.dataTransfer?.files||[])});
document.addEventListener('click',ev=>{
  const t=ev.target.closest('button');if(!t)return;
  if(t.id==='import-slides'){$('#import-slides-file')?.click()}
  else if(t.id==='add-background-btn'){openBackgroundPicker()}
  else if(t.id==='background-upload-btn'){$('#background-upload-file')?.click()}
  else if(t.dataset.mediaView){openMediaPreview(t.dataset.mediaView)}
  else if(t.dataset.mediaViewPpt){const first=mediaMeta.find(m=>m.presentationId===t.dataset.mediaViewPpt)||mediaMeta.find(m=>m.id===t.dataset.mediaViewPpt);if(first)openMediaPreview(first.id,first.presentationId||null)}
  else if(t.dataset.mediaAddAnnouncement){addAnnouncementSlide(t.dataset.mediaAddAnnouncement)}
  else if(t.dataset.announcementDownloadAll){downloadAllAnnouncements()}
  else if(t.dataset.announcementDownload){downloadAnnouncement(t.dataset.announcementDownload)}
  else if(t.dataset.mediaAddPpt){addPowerPointGroup(t.dataset.mediaAddPpt)}
  else if(t.id==='bg-clear'){bgApply(o=>{if(o.bgMedia)delete o.bgMedia;else{o.bg=0;delete o.customBg}})}
  else if(t.id==='media-open'){openMediaLibrary()}
  else if(t.id==='ppt-test-open'){testOpenPowerpoint(t)}
  else if(t.id==='media-close'){closeMediaLibrary()}
  else if(t.id==='background-picker-close'){closeBackgroundPicker()}
  else if(t.id==='logo-library-sync'){syncDriveLibraryKind('logos');}
  else if(t.id==='library-sync'){cloudSyncAll({manual:true});}
  else if(t.dataset.manualSync){syncDriveLibraryKind('announcements');}
  else if(t.dataset.backgroundSync){syncDriveBackgrounds();}
  else if(t.dataset.backgroundDownload!=null){const m=mediaMeta.find(x=>x.id===t.dataset.backgroundDownload);if(m){rootBackgroundDownload(m);}}
  else if(t.dataset.backgroundSelect!=null){selectSongBackground(t.dataset.backgroundSelect)}
  else if(t.dataset.backgroundDelete!=null){deleteMedia(t.dataset.backgroundDelete)}
  else if(t.dataset.backgroundTab){const root=$('#background-picker-content');if(root){root.dataset.backgroundTab=t.dataset.backgroundTab;renderBackgroundPicker()}}
  else if(t.id==='media-upload-btn'){$('#media-file').click()}
  else if(t.id==='media-upload-slides-btn'){$('#media-slides-file')?.click()||$('#import-slides-file')?.click()}
  else if(t.dataset.mediaAddSection!=null){addMediaAsSection(t.dataset.mediaAddSection)}
  else if(t.dataset.mediaAddSlide!=null){addMediaToCurrentSlide(t.dataset.mediaAddSlide)}
  else if(t.dataset.mediaUse!=null){addMediaAsSection(t.dataset.mediaUse)}
  else if(t.dataset.mediaDelete!=null){deleteMedia(t.dataset.mediaDelete)}
  else if(t.dataset.mediaRename!=null){renameMedia(t.dataset.mediaRename)}
  else if(t.dataset.mediaDeletePpt!=null){deletePowerPointGroup(t.dataset.mediaDeletePpt)}
  else if(t.dataset.mediaTab){const root=$('#media-content');if(root){root.dataset.mediaTab=t.dataset.mediaTab;renderMediaLibrary()}}
  else if(t.id==='sb-open'||t.id==='sb-open-secondary')openStruct();
  else if(t.id==='add-slide-video')$('#slide-video-file')?.click();
  else if(t.id==='add-video-slide')$('#add-video-slide-file')?.click();
  else if(t.id==='remove-slide-video'){const p=sls(A)[aj];if(p){delete p.videoMedia;rAll()}}
});


/* ===== service overview column ===== */
let odrag=null;
function rOutline(){
  if(!S.length){
    $('#outline').innerHTML='<h3>Order of Sections</h3><p class="note" style="padding:8px">Empty — add a slide from the strip.</p>';
    return;
  }
  let html='<h3>Order of Sections</h3><div class="oi-insert" data-insert="0"></div>';
  S.forEach((it,i)=>{
    html+=`<div class="oi ${i==sel?'on':''}" data-oi="${i}" title="Click to select · drag to reorder · right-click for more" style="cursor:grab">
      <span class="og" data-drag-handle="1" title="Drag to reorder">⋮⋮</span>
      <span class="ot">${i+1}. ${e(it.t)}</span>
      <small>${e(it.k)}</small>
    </div>
    <div class="oi-insert" data-insert="${i+1}"></div>`;
  });
  $('#outline').innerHTML=html;
}
function jumpToItem(i){
  imageSel=-1;sel=i;si=0;A=S[i];aj=0;clr=null;rAll();
  const st=$('#strip'),el=st.querySelector('[data-dh="'+i+'"]');
  if(el){const a=st.getBoundingClientRect(),b=el.getBoundingClientRect();
    st.scrollTo({top:b.top-a.top+st.scrollTop-6,left:b.left-a.left+st.scrollLeft-6,behavior:'smooth'})}
}

function scrollServiceItemIntoView(i){
  const st=$('#strip'),el=st.querySelector('[data-dh="'+i+'"]');
  if(!st||!el)return;
  const a=st.getBoundingClientRect(),b=el.getBoundingClientRect();
  st.scrollTo({top:b.top-a.top+st.scrollTop-6,left:b.left-a.left+st.scrollLeft-6,behavior:'smooth'});
}
{
  const ol=$('#outline');
  const menu=$('#outline-menu');
  let outlineMenuIndex=-1;
  let drag=null; // {from, pointerId}

  function hideOutlineMenu(){
    if(!menu)return;
    menu.hidden=true;
    menu.classList.remove('open');
    outlineMenuIndex=-1;
  }
  function showOutlineMenu(x,y,index){
    if(!menu)return;
    outlineMenuIndex=index;
    menu.hidden=false;
    menu.classList.add('open');
    const pad=6;
    menu.style.left='0px';menu.style.top='0px';
    const w=menu.offsetWidth||160,h=menu.offsetHeight||90;
    menu.style.left=Math.min(x,window.innerWidth-w-pad)+'px';
    menu.style.top=Math.min(y,window.innerHeight-h-pad)+'px';
  }
  function clearDragUI(){
    ol.querySelectorAll('.dragging,.oi-insert.on').forEach(x=>x.classList.remove('dragging','on'));
    document.body.dataset.pointerDragging='';
  }
  function insertIndexFromPoint(clientY){
    const items=[...ol.querySelectorAll('[data-oi]')];
    if(!items.length)return 0;
    for(let i=0;i<items.length;i++){
      const r=items[i].getBoundingClientRect();
      const mid=r.top+r.height/2;
      if(clientY<mid)return i;
    }
    return items.length; // after last
  }
  function paintInsertMarker(slot){
    ol.querySelectorAll('.oi-insert').forEach(x=>x.classList.remove('on'));
    const markers=[...ol.querySelectorAll('.oi-insert')];
    if(markers[slot])markers[slot].classList.add('on');
  }
  function moveOutlineItem(from,toSlot){
    if(from<0||from>=S.length)return;
    pushUndo();
    let dest=toSlot;
    if(dest>from)dest-=1; // account for removal
    if(dest===from||dest<0||dest>=S.length) {
      // allow move to end
      if(!(toSlot===S.length && from!==S.length-1)) {
        if(dest===from)return;
      }
    }
    if(toSlot===S.length) dest=S.length-1;
    // recompute properly:
    // remove from `from`, insert so final index is `toSlot` if toSlot>from then toSlot-1 else toSlot
    dest = toSlot > from ? toSlot - 1 : toSlot;
    if(dest===from)return;
    const cur=S[sel];
    const [m]=S.splice(from,1);
    S.splice(Math.max(0,Math.min(dest,S.length)),0,m);
    let newSel=S.indexOf(cur);
    if(newSel<0)newSel=Math.min(dest,S.length-1);
    sel=newSel;si=0;A=S[sel];aj=0;
    rAll();
    setTimeout(scrollCurrentSlideIntoView,0);
  }

  // Outline HTML (with insert markers) is produced by rOutline()

  ol.addEventListener('click',ev=>{
    if(menu&&!menu.hidden&&!ev.target.closest('#outline-menu'))hideOutlineMenu();
    if(drag)return;
    if(ev.target.closest('[data-drag-handle]'))return;
    const o=ev.target.closest('[data-oi]');
    if(o)jumpToItem(+o.dataset.oi);
  });

  ol.addEventListener('contextmenu',ev=>{
    const o=ev.target.closest('[data-oi]');
    if(!o)return;
    ev.preventDefault();
    showOutlineMenu(ev.clientX,ev.clientY,+o.dataset.oi);
  });

  document.addEventListener('click',async ev=>{
    if(menu&&!menu.hidden&&!ev.target.closest('#outline-menu'))hideOutlineMenu();
  });

  menu?.addEventListener('click',ev=>{
    const btn=ev.target.closest('[data-outline-action]');
    if(!btn||outlineMenuIndex<0)return;
    const action=btn.dataset.outlineAction;
    const i=outlineMenuIndex;
    hideOutlineMenu();
    if(action==='rename')renameServiceItem(i);
    else if(action==='duplicate'){
      if(i<0||i>=S.length)return;
      const copy=cl(S[i]);
      if(Array.isArray(copy.parts)){
        const map={};
        copy.parts=copy.parts.map(p=>{const nid=uid();map[p.id]=nid;return {...p,id:nid}});
        copy.ord=(copy.ord||[]).map(id=>map[id]||id);copy.parts.forEach(p=>{if(p.of&&map[p.of])p.of=map[p.of]});
      }
      /* keep original title */
      pushUndo();
      S.splice(i+1,0,copy);
      sel=i+1;si=0;A=S[sel];aj=0;
      rAll();setTimeout(()=>{scrollNewItemToTop(sel);scrollCurrentSlideIntoView()},0);
      toast('Duplicated “'+copy.t+'?.');
    }else if(action==='remove'){
      removeServiceItem(i);
    }
  });

  // Pointer drag from anywhere on the section row (shared pill animation)
  ol.addEventListener('pointerdown',ev=>{
    if(ev.button!==0)return;
    if(ev.target.closest('button,a,input,textarea,select'))return;
    const o=ev.target.closest('[data-oi]');
    if(!o||!ol.contains(o))return;
    const from=+o.dataset.oi;
    const onHandle=!!ev.target.closest('[data-drag-handle]');
    let toSlot=null;
    pillDrag(ev,{
      el:o,label:(S[from]&&S[from].t)||'Section',sub:(S[from]&&S[from].k)||'',
      scrollEl:ol,
      onTap:()=>{if(onHandle)jumpToItem(from)},
      onMove:(x,y)=>{
        const r=ol.getBoundingClientRect();
        if(x>=r.left-60&&x<=r.right+60){toSlot=insertIndexFromPoint(y);paintInsertMarker(toSlot)}
        else{toSlot=null;ol.querySelectorAll('.oi-insert').forEach(m=>m.classList.remove('on'))}
      },
      onDrop:()=>{if(toSlot!=null)moveOutlineItem(from,toSlot)},
      onEnd:()=>clearDragUI()
    });
  });
}


/* ===== slide strip context menu ===== */
{
  const strip=$('#strip');
  const menu=$('#slide-menu');
  let slideMenuTarget=null; // {i,j}

  function hideSlideMenu(){
    if(!menu)return;
    menu.hidden=true;
    menu.classList.remove('open');
    slideMenuTarget=null;
  }
  function addSectionAfter(i){
    pushUndo();
    const it=mkS(nextSectionName(),['Type your text here'],0,6.5);
    S.splice(i+1,0,it);
    sel=i+1;si=0;A=it;aj=0;tab='Structure';
    rAll();
    setTimeout(()=>scrollNewItemToTop(sel),0);
    toast('Added new section.');
  }
  function showSlideMenu(x,y,i,j){
    if(!menu)return;
    slideMenuTarget={i,j};
    const it=S[i];
    const isSong=it&&it.k==='Song';
    menu.innerHTML=`
      <div class="has-sub">
        <button type="button">New ▸</button>
        <div class="sub">
          ${isSong?'':`<button type="button" data-slide-action="new">New slide</button>`}
          <button type="button" data-slide-action="section">New section</button>
        </div>
      </div>
      <button type="button" data-slide-action="copy">Copy</button>
      <button type="button" data-slide-action="paste" ${slideClip?'':'disabled'}>Paste</button>
      <button type="button" data-slide-action="cut">Cut</button>
      <button type="button" data-slide-action="duplicate">Duplicate</button>
      <button type="button" data-slide-action="remove" class="danger">Remove</button>`;
    menu.hidden=false;menu.classList.add('open');
    const pad=6;
    menu.style.left='0px';menu.style.top='0px';
    const w=menu.offsetWidth||160,h=menu.offsetHeight||180;
    menu.style.left=Math.min(x,window.innerWidth-w-pad)+'px';
    menu.style.top=Math.min(y,window.innerHeight-h-pad)+'px';
  }
  

  // Section header context menu (right-click on strip item header)
  const itemMenu=$('#item-menu');
  let itemMenuIndex=-1;
  function hideItemMenu(){
    if(!itemMenu)return;
    itemMenu.hidden=true;
    itemMenu.classList.remove('open');
    itemMenuIndex=-1;
  }
  function showItemMenu(x,y,i){
    if(!itemMenu)return;
    itemMenuIndex=i;
    const it=S[i];
    const isSong=it&&it.k==='Song';
    itemMenu.innerHTML=`
      <button type="button" data-item-action="up" ${i===0?'disabled':''}>↑ Up</button>
      <button type="button" data-item-action="down" ${i===S.length-1?'disabled':''}>↓ Down</button>
      <button type="button" data-item-action="rename">✎ Rename</button>
      <button type="button" data-item-action="duplicate">⧉ Duplicate</button>
      ${isSong?'':`<button type="button" data-item-action="new">＋ New Slide</button>`}
      <button type="button" data-item-action="remove" class="danger">✕ Remove</button>`;
    itemMenu.hidden=false;
    itemMenu.classList.add('open');
    const pad=6;
    itemMenu.style.left='0px';itemMenu.style.top='0px';
    const w=itemMenu.offsetWidth||160,h=itemMenu.offsetHeight||200;
    itemMenu.style.left=Math.min(x,window.innerWidth-w-pad)+'px';
    itemMenu.style.top=Math.min(y,window.innerHeight-h-pad)+'px';
  }
  strip?.addEventListener('contextmenu',ev=>{
    // Prefer slide thumb menu when right-clicking a thumb
    if(ev.target.closest('.thumb-wrap[data-i][data-j]'))return;
    const hb=ev.target.closest('.hb[data-dh]');
    if(!hb)return;
    ev.preventDefault();
    showItemMenu(ev.clientX,ev.clientY,+hb.dataset.dh);
  });
  document.addEventListener('click',ev=>{
    if(itemMenu&&!itemMenu.hidden&&!ev.target.closest('#item-menu'))hideItemMenu();
  });
  itemMenu?.addEventListener('click',ev=>{
    const btn=ev.target.closest('[data-item-action]');
    if(!btn||itemMenuIndex<0)return;
    const action=btn.dataset.itemAction;
    const i=itemMenuIndex;
    hideItemMenu();
    const it=S[i];if(!it)return;
    if(action==='up'){
      if(i<=0)return;
      [S[i-1],S[i]]=[S[i],S[i-1]];
      sel=i-1;si=0;A=S[sel];aj=0;stripFocus='item';
      rAll();setTimeout(()=>scrollServiceItemIntoView(sel),0);
    }else if(action==='down'){
      if(i>=S.length-1)return;
      [S[i],S[i+1]]=[S[i+1],S[i]];
      sel=i+1;si=0;A=S[sel];aj=0;stripFocus='item';
      rAll();setTimeout(()=>scrollServiceItemIntoView(sel),0);
    }else if(action==='rename')renameServiceItem(i);
    else if(action==='duplicate'){
      const copy=cl(it);
      if(Array.isArray(copy.parts)){
        const map={};
        copy.parts=copy.parts.map(p=>{const nid=uid();map[p.id]=nid;const p2={...p,id:nid};if(Array.isArray(p2.elements))p2.elements=p2.elements.map(el=>({...cl(el),id:uid()}));return p2});
        copy.ord=(copy.ord||[]).map(id=>map[id]||id);copy.parts.forEach(p=>{if(p.of&&map[p.of])p.of=map[p.of]});
      }
      /* keep original title */
      pushUndo();
      S.splice(i+1,0,copy);
      sel=i+1;si=0;A=S[sel];aj=0;stripFocus='item';
      rAll();setTimeout(()=>{scrollNewItemToTop(sel);scrollCurrentSlideIntoView()},0);
      toast('Duplicated “'+copy.t+'?.');
    }else if(action==='new'){
      if(it.k==='Song'){toast('Use Edit to add song sections.');return}
      addNewSlideInItem(i,it.ord.length-1);
    }else if(action==='remove'){
      removeServiceItem(i);
    }
  });

  strip?.addEventListener('contextmenu',ev=>{
    const wrap=ev.target.closest('.thumb-wrap[data-i][data-j]');
    if(!wrap)return;
    ev.preventDefault();
    showSlideMenu(ev.clientX,ev.clientY,+wrap.dataset.i,+wrap.dataset.j);
  });
  document.addEventListener('click',ev=>{
    if(menu&&!menu.hidden&&!ev.target.closest('#slide-menu'))hideSlideMenu();
  });
  menu?.addEventListener('click',ev=>{
    const btn=ev.target.closest('[data-slide-action]');
    if(!btn||!slideMenuTarget)return;
    const {i,j}=slideMenuTarget;
    const action=btn.dataset.slideAction;
    hideSlideMenu();
    const it=S[i];if(!it)return;
    if(action==='up')moveSlideInItem(i,j,j-1);
    else if(action==='copy')copySlideAt(i,j);
    else if(action==='cut')cutSlideAt(i,j);
    else if(action==='paste')pasteSlideAfter(i,j);
    else if(action==='new')addNewSlideInItem(i,j);
    else if(action==='section')addSectionAfter(i);
    else if(action==='duplicate')duplicateSlideInItem(i,j);
    else if(action==='remove')removeSlideInItem(i,j);
  });
}


// ===== Strip: drag slides (shared pill animation as a mini slide picture, auto-scroll) =====
{
  const strip=$('#strip');
  function clearStripDragUI(){
    strip?.querySelectorAll('.thumb-wrap.drag-over-above,.thumb-wrap.drag-over-below').forEach(x=>x.classList.remove('drag-over-above','drag-over-below'));
  }
  function stripTargetFromY(clientY){
    if(!strip)return null;
    const wraps=[...strip.querySelectorAll('.thumb-wrap[data-i][data-j]')];
    if(!wraps.length)return null;
    let idx=-1;
    for(let k=0;k<wraps.length;k++){
      const r=wraps[k].getBoundingClientRect();
      if(clientY<r.top+r.height/2){idx=k;break}
    }
    let el,before;
    if(idx<0){el=wraps[wraps.length-1];before=false}
    else{el=wraps[idx];before=true}
    // Gap just above a song's first slide: treat it as "end of the section above" when that is a normal section
    if(before&&idx>0&&+el.dataset.j===0&&S[+el.dataset.i]&&isSongItem(S[+el.dataset.i])){
      const prev=wraps[idx-1];
      if(S[+prev.dataset.i]&&!isSongItem(S[+prev.dataset.i])){el=prev;before=false}
    }
    return {el,before,i:+el.dataset.i,j:+el.dataset.j};
  }
  function dropSlide(fromI,fromJ,t){
    const fromItem=S[fromI];
    if(!fromItem||!Array.isArray(fromItem.ord))return;
    if(fromI===t.i){
      let toJ=t.before?t.j:t.j+1;
      if(toJ>fromJ)toJ--;
      if(toJ===fromJ||toJ<0||toJ>=fromItem.ord.length)return;
      moveSlideInItem(fromI,fromJ,toJ);
      return;
    }
    const toItem=S[t.i];
    if(!toItem||toItem.k==='Song'){toast('Cannot move slides into a song section.');return}
    if(isSongItem(fromItem)||fromItem.k==='Song'){toast('Song slides cannot be moved outside of their song.');return}
    const id=fromItem.ord[fromJ];
    if(!id)return;
    const part=fromItem.parts.find(p=>p.id===id);
    if(!part)return;
    pushUndo();
    if(toItem.bgAll!==true)makeOwnBg(fromItem,part);
    else clearOwnBg(part);
    fromItem.ord.splice(fromJ,1);
    if(!fromItem.ord.length){
      const blank={id:uid(),n:'Slide 1',tx:''};
      fromItem.parts.push(blank);
      fromItem.ord.push(blank.id);
    }
    if(!toItem.parts.find(p=>p.id===part.id))toItem.parts.push(cl(part));
    let insertAt=t.j+(t.before?0:1);
    insertAt=Math.max(0,Math.min(insertAt,toItem.ord.length));
    toItem.ord.splice(insertAt,0,part.id);
    sel=t.i;si=insertAt;A=toItem;aj=insertAt;stripFocus='slide';
    rAll();
    setTimeout(scrollCurrentSlideIntoView,0);
  }
  strip?.addEventListener('pointerdown',ev=>{
    if(ev.button!==0||live)return;
    if(ev.target.closest('.hd,.th-add,#outline-menu,#slide-menu,#item-menu'))return;
    const wrap=ev.target.closest('.thumb-wrap[data-i][data-j]');
    if(!wrap||!strip.contains(wrap))return;
    const fromI=+wrap.dataset.i,fromJ=+wrap.dataset.j;
    if(!S[fromI])return;
    let tgt=null;
    pillDrag(ev,{
      el:wrap,
      thumbOf:()=>wrap.querySelector('.th .sl')||wrap.querySelector('.sl'),
      scrollEl:strip,
      onMove:(x,y)=>{
        clearStripDragUI();
        const t=stripTargetFromY(y);
        const fromItem=S[fromI];
        const isFromSong=fromItem&&(isSongItem(fromItem)||fromItem.k==='Song');
        if(t&&S[t.i]&&(!isSongItem(S[t.i])||t.i===fromI)&&(!isFromSong||t.i===fromI)){
          tgt=t;t.el.classList.add(t.before?'drag-over-above':'drag-over-below')
        }else tgt=null;
      },
      onDrop:()=>{if(tgt)dropSlide(fromI,fromJ,tgt)},
      onEnd:clearStripDragUI
    });
  });
  strip?.addEventListener('dragstart',ev=>{
    if(ev.target.closest('.thumb-wrap')){ev.preventDefault()}
  });
  // keep the "New >" submenu on screen
  $('#slide-menu')?.addEventListener('mouseover',ev=>{
    const hs=ev.target.closest('.has-sub');if(!hs)return;
    const sub=hs.querySelector('.sub');if(!sub)return;
    hs.classList.remove('flip');
    if(sub.getBoundingClientRect().right>window.innerWidth-8)hs.classList.add('flip');
  });
}

/* ===== song structure builder ===== */

let sbItem=null,sbSel=0,sbDrag=null,sbEditing=true,sbLps=0,sbNeedReflow=false,homeShown=5,olderServicesVisible=false;
const snip=t=>{const s=String(t||'').replace(/\s+/g,' ').trim().slice(0,60);return e(s||'(empty)')};
function sbPartCard(it,p){
  const n=it.ord.filter(x=>x==p.id).length;
  return `<div class="sb-card" draggable="true" data-sb-part="${p.id}"><div class="sb-main"><b>${e(p.n)}</b><small>${snip(p.tx)}</small></div><span class="sb-actions"><span class="sb-count">×${n}</span><button data-sb-add="${p.id}" title="Add to end of arrangement">+</button>${n?'':`<button data-sb-delpart="${p.id}" title="Delete unused section">✕</button>`}</span></div>`;
}
function sbOrdCard(it,id,j){
  const p=it.parts.find(x=>x.id==id)||{n:'?',tx:''};
  return `<div class="sb-card ${j==sbSel?'sel':''}" draggable="true" data-sb-ord="${j}" data-sb-slot="${j}"><span class="sb-n">${j+1}</span><div class="sb-main" data-sb-pick="${j}"><b>${e(p.n)}</b><small>${snip(p.tx)}</small></div><span class="sb-actions"><button data-sb-dup="${j}" title="Repeat this section after it">⧉</button><button data-sb-rm="${j}" title="Remove">✕</button></span></div>`;
}
function rStruct(){
  const it=sbItem;if(!it)return;
  // Preserve scroll so clicking a row does not jump the order list
  const contentEl=$('#struct-content');
  const ordListEl=$('#sb-ord-list');
  const prevContentScroll=contentEl?contentEl.scrollTop:0;
  const prevOrdScroll=ordListEl?ordListEl.scrollTop:0;
  // Skip auto blank slides in the structure builder (controlled by Song settings)
  const lyricOrdIdxs=it.ord.map((id,j)=>({id,j})).filter(({id})=>{
    const p=it.parts.find(x=>x.id===id);
    return p&&!isAutoBlankPart(p);
  });
  if(!lyricOrdIdxs.length){
    sbSel=0;
  }else if(!lyricOrdIdxs.some(x=>x.j===sbSel)){
    const near=lyricOrdIdxs.find(x=>x.j>=sbSel)||lyricOrdIdxs[lyricOrdIdxs.length-1];
    sbSel=near.j;
  }
  const curPart=it.parts.find(p=>p.id==it.ord[sbSel]);
  const cur=(curPart&&!isAutoBlankPart(curPart)&&sbEditing)?curPart:null;
  const lyricParts=(it.parts||[]).filter(p=>!isAutoBlankPart(p));
  $('#struct-title').textContent='Song structure — '+it.t;
  $('#struct-content').innerHTML=`
  <div class="sb-top" style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:12px">
    <p class="note" style="margin:0;flex:1;min-width:200px">Arrange sections for this song. Blank slides from Settings stay outside this editor.</p>
    <label style="margin:0;display:flex;align-items:center;gap:8px">Lines per slide
      <select id="sb-lps" style="width:auto;margin:0">${lpsOptions(sbLps)}</select>
    </label>
  </div>
  <div class="sb-layout">
    <section class="sb-edit-pane">
      <h3 class="sb-edit-heading">Edit selected section</h3>
      ${cur?`
        <label style="font-size:calc(13px * var(--ui-scale));font-weight:600;margin:0 0 4px">Section name</label>
        <input id="sb-name" type="text" value="${e(cur.n)}" style="font-size:calc(16px * var(--ui-scale));font-weight:600;padding:10px 12px;margin-bottom:10px">
        ${cur.blank?'<p class="note" style="margin:6px 0">Blank slide: shows the song background with no lyrics.</p>':'<label style="font-size:calc(13px * var(--ui-scale));font-weight:600;margin:0 0 4px">Lyrics</label>'}
        ${cur.blank?'':`<textarea id="sb-text" rows="14" placeholder="Type lyrics here" style="font-size:calc(15px * var(--ui-scale));line-height:1.45;padding:12px;min-height:220px">${e(cur.tx)}</textarea>`}
        <div style="display:flex;align-items:center;gap:12px;margin-top:10px"><button type="button" id="sb-done" class="sb-done">Done</button><span class="note" style="margin:0">Edits apply everywhere this section is used in the order.</span></div>
      `:'<p class="note">Select a section in the order list to edit its name and lyrics.</p>'}
    </section>
    <section class="sb-col sb-palette-pane">
      <h3>Available sections</h3>
      <p class="note">Building blocks. Drag a section into the song order.</p>
      <div class="sb-palette">
        ${lyricParts.map(p=>`<div class="sb-section-row" data-sb-part-row="${p.id}"><span class="sb-grip" title="Drag into the song order">⋮⋮</span><div class="sb-main"><b>${e(p.n)}</b><small>${it.ord.filter(x=>x===p.id).length}× in order</small></div><button data-sb-delpart="${p.id}" title="Delete section">Delete</button></div>`).join('')||'<p class="note">No sections yet.</p>'}
      </div>
      <h4 style="margin:14px 0 6px;font-size:calc(12px * var(--ui-scale));text-transform:uppercase;letter-spacing:.04em;color:var(--mu)">Create</h4>
      <div class="chips">${TYPES.map(n=>`<button data-sb-new="${n}">+ ${n}</button>`).join('')}<button data-sb-blank="1">+ Blank slide</button></div>
    </section>
    <section class="sb-col sb-order-pane">
      <h3>Song order (${lyricOrdIdxs.length})</h3>
      <p class="note">Drag ⋮⋮ to reorder. ⧉ repeats a section.</p>
      <div class="sb-ord-list" id="sb-ord-list">
        <div class="sb-insert-line" data-sb-insert="0"></div>
        ${it.ord.map((id,j)=>{
          const p=it.parts.find(x=>x.id===id);
          if(p&&isAutoBlankPart(p))return '';
          return sbSimpleOrdCard(it,id,j)+`<div class="sb-insert-line" data-sb-insert="${j+1}"></div>`;
        }).join('')||'<p class="note">Add a section from Available sections.</p>'}
      </div>
    </section>
  </div>`;

  // Restore scroll positions after full re-render
  requestAnimationFrame(()=>{
    const c=$('#struct-content');
    const o=$('#sb-ord-list');
    if(c)c.scrollTop=prevContentScroll;
    if(o)o.scrollTop=prevOrdScroll;
    // Keep selected row visible without jumping the whole list to top
    const selRow=o&&o.querySelector('.sb-arr-row.sel');
    if(selRow&&o){
      const top=selRow.offsetTop;
      const bottom=top+selRow.offsetHeight;
      if(top<o.scrollTop)o.scrollTop=Math.max(0,top-8);
      else if(bottom>o.scrollTop+o.clientHeight)o.scrollTop=bottom-o.clientHeight+8;
    }
  });
}

function sbSimpleOrdCard(it,id,j){
  const p=it.parts.find(x=>x.id===id)||{n:'?',tx:''};
  return `<div class="sb-arr-row ${j===sbSel?'sel':''}" data-sb-ord="${j}" data-sb-pick="${j}">
    <span class="og" data-sb-drag="1" title="Drag to reorder" style="cursor:grab;padding:2px 4px;color:var(--mu)">⋮⋮</span>
    <span style="font-weight:700;min-width:22px">${j+1}.</span>
    <div class="sb-main"><b>${e(p.n)}</b><small>${snip(p.tx)}</small></div>
    <button data-sb-dup="${j}" title="Duplicate section">⧉</button>
    <button data-sb-rm="${j}" title="Remove">✕</button>
  </div>`;
}
function openStruct(){
  sbItem=S[sel];if(!sbItem)return;
  let keepName='';
  sbNeedReflow=false;
  if(isSongItem(sbItem)){
    collapseFit(sbItem);
    const cp=sbItem.parts.find(x=>x.id===sbItem.ord[Math.min(si,sbItem.ord.length-1)]);
    keepName=cp?String(cp.n).replace(/\s*\(\d+\/\d+\)$/,''):'';
    sbNeedReflow=sbItem.parts.some(p=>/\(\d+\/\d+\)$/.test(p.n));
    sbLps=+curLps(sbItem)||0;
    if(sbNeedReflow){reflowItem(sbItem,0,false);sbItem.lps=sbLps}
  }
  sbSel=0;
  if(keepName){
    const k=sbItem.ord.findIndex(id=>{const q=sbItem.parts.find(x=>x.id===id);return q&&q.n===keepName});
    if(k>=0)sbSel=k;
  }
  if(sbItem.ord.length){
    const p=sbItem.parts.find(x=>x.id===sbItem.ord[sbSel]);
    if(p&&isAutoBlankPart(p)){
      const idx=sbItem.ord.findIndex(id=>{const q=sbItem.parts.find(x=>x.id===id);return q&&!isAutoBlankPart(q)});
      if(idx>=0)sbSel=idx;
    }
  }
  sbEditing=true;$('#struct-modal').hidden=false;rStruct();
}
function closeStruct(){
  const _fitIt=sbItem;
  if(_fitIt&&isSongItem(_fitIt)){
    try{stripPresentBlanks(_fitIt)}catch{}
    if(sbNeedReflow&&sbLps>0){try{reflowItem(_fitIt,sbLps,false)}catch(err){console.warn('reflow',err)}}
  }
  $('#struct-modal').hidden=true;sbItem=null;
  if(_fitIt&&isSongItem(_fitIt)&&S.includes(_fitIt)){try{expandFit(_fitIt)}catch(err){console.warn('fit',err)}}
  rAll();
}
const sbChanged=()=>{rAll();rStruct()};
{
  const M=$('#struct-modal');
  $('#struct-close').addEventListener('click',closeStruct);
  document.addEventListener('keydown',ev=>{if(ev.key==='Escape'&&!M.hidden)closeStruct()});
  M.addEventListener('click',async ev=>{
    const it=sbItem;if(!it)return;
    const b=ev.target.closest('[data-sb-add],[data-sb-delpart],[data-sb-new],[data-sb-dup],[data-sb-rm],[data-sb-pick],[data-sb-move],[data-sb-blank]');
    if(!b)return;const d=b.dataset;sbEditing=true;
    if(d.sbAdd){
      // Insert before trailing auto-blanks
      let at=it.ord.length;
      while(at>0){
        const p=it.parts.find(x=>x.id===it.ord[at-1]);
        if(!p||!isAutoBlankPart(p))break;
        at--;
      }
      insOrd(it,at,d.sbAdd);sbSel=at;
    }
    else if(d.sbDelpart){
      const id=d.sbDelpart,n=it.ord.filter(x=>x===id).length;
      if(!await appConfirm(n?`Delete ${it.parts.find(p=>p.id===id)?.n||'this section'} and remove its ${n} occurrence(s) from the song?`:`Delete ${it.parts.find(p=>p.id===id)?.n||'this section'}?`,{title:'Delete song section',confirmLabel:'Delete',danger:true}))return;
      it.parts=it.parts.filter(p=>p.id!==id);it.ord=it.ord.filter(x=>x!==id);sbSel=Math.max(0,Math.min(sbSel,it.ord.length-1));
    }
    else if(d.sbNew){
      const c=it.parts.filter(p=>p.n===d.sbNew||p.n.startsWith(d.sbNew+' ')).length;
      const p={id:uid(),n:c?`${d.sbNew} ${c+1}`:(d.sbNew==='Verse'?'Verse 1':d.sbNew),tx:''};
      it.parts.push(p);
      let at=it.ord.length;
      while(at>0){
        const q=it.parts.find(x=>x.id===it.ord[at-1]);
        if(!q||!isAutoBlankPart(q))break;
        at--;
      }
      insOrd(it,at,p.id);sbSel=at;
    }
    else if(d.sbBlank){
      const c=it.parts.filter(p=>p.blank).length;
      const p={id:uid(),n:c?`Blank ${c+1}`:'Blank',tx:'',blank:true};
      it.parts.push(p);
      const at=Math.min(it.ord.length,sbSel+1);
      insOrd(it,at,p.id);sbSel=at;
    }
    else if(d.sbMove){
      const j=+d.sbIndex,n=d.sbMove==='up'?j-1:j+1;if(n>=0&&n<it.ord.length){mvOrd(it,j,n);sbSel=n}
    }
    else if(d.sbDup!=null){const j=+d.sbDup;insOrd(it,j+1,it.ord[j]);sbSel=j+1}
    else if(d.sbRm!=null){if(it.ord.length<2)return toast('Keep at least one section.');rmOrd(it,+d.sbRm);sbSel=Math.min(sbSel,it.ord.length-1)}
    else if(d.sbPick!=null){sbSel=+d.sbPick;rStruct();return}
    sbChanged();
  });
  M.addEventListener('input',ev=>{
    const it=sbItem;if(!it)return;const p=it.parts.find(x=>x.id===it.ord[sbSel]);if(!p)return;
    if(ev.target.id==='sb-name')p.n=ev.target.value;
    else if(ev.target.id==='sb-text'){p.tx=ev.target.value;if(sbLps>0)sbNeedReflow=true}
    else return;
    rStrip();rStage();saveSoon();
  });
  
  // Smooth pointer drag-and-drop for song order
  let sbDrag=null;
  function clearSbDragUI(){
    M.querySelectorAll('.sb-arr-row.dragging,.sb-insert-line.on').forEach(x=>{
      x.classList.remove('dragging','on');
    });
    document.body.dataset.pointerDragging='';
  }
  function sbInsertSlotFromY(clientY){
    const rows=[...M.querySelectorAll('.sb-arr-row[data-sb-ord]')];
    if(!rows.length)return 0;
    for(let i=0;i<rows.length;i++){
      const r=rows[i].getBoundingClientRect();
      if(clientY < r.top + r.height/2) return i;
    }
    return rows.length;
  }
  function paintSbInsert(slot){
    M.querySelectorAll('.sb-insert-line').forEach(x=>x.classList.remove('on'));
    const line=M.querySelector(`.sb-insert-line[data-sb-insert="${slot}"]`);
    if(line)line.classList.add('on');
  }
  M.addEventListener('pointerdown',ev=>{
    if(ev.button!==0)return;
    const handle=null; /* replaced by whole-row drag below */
    if(!handle)return;
    const row=handle.closest('[data-sb-ord]');
    if(!row||!sbItem)return;
    ev.preventDefault();
    ev.stopPropagation();
    const from=+row.dataset.sbOrd;
    sbDrag={from,pointerId:ev.pointerId,moved:false,toSlot:from};
    row.classList.add('dragging');
    document.body.dataset.pointerDragging='1';
    try{M.setPointerCapture(ev.pointerId)}catch{}
    paintSbInsert(from);
  });
  M.addEventListener('pointermove',ev=>{
    if(!sbDrag||ev.pointerId!==sbDrag.pointerId)return;
    sbDrag.moved=true;
    const slot=sbInsertSlotFromY(ev.clientY);
    sbDrag.toSlot=slot;
    paintSbInsert(slot);
  });
  function endSbDrag(ev){
    if(!sbDrag||(ev&&ev.pointerId!==sbDrag.pointerId))return;
    const {from,toSlot,moved}=sbDrag;
    sbDrag=null;
    clearSbDragUI();
    if(!moved||!sbItem||toSlot==null)return;
    let dest=toSlot>from?toSlot-1:toSlot;
    if(dest===from)return;
    if(dest<0||dest>=sbItem.ord.length)return;
    mvOrd(sbItem,from,dest);
    sbSel=dest;
    sbChanged();
  }
  M.addEventListener('pointerup',endSbDrag);
  M.addEventListener('pointercancel',endSbDrag);
  // ---- drag & drop in the song structure builder (shared pill animation) ----
  function sbSlotToAt(slot){
    const rows=[...M.querySelectorAll('.sb-arr-row[data-sb-ord]')];
    if(slot<rows.length)return +rows[slot].dataset.sbOrd;
    return rows.length?+rows[rows.length-1].dataset.sbOrd+1:(sbItem?sbItem.ord.length:0);
  }
  function sbDragStart(ev,row,kind){
    if(ev.button!==0||!sbItem)return;
    if(ev.target.closest('button,input,textarea,select,[contenteditable]'))return;
    const it=sbItem;
    let pid=null,from=-1,part=null;
    if(kind==='palette'){pid=row.dataset.sbPartRow;part=it.parts.find(x=>x.id===pid)}
    else{from=+row.dataset.sbOrd;part=it.parts.find(x=>x.id===it.ord[from])}
    if(!part)return;
    let slot=null;
    const clearUi=()=>{
      M.querySelectorAll('.sb-drop-hot').forEach(x=>x.classList.remove('sb-drop-hot'));
      M.querySelectorAll('.sb-insert-line.on').forEach(x=>x.classList.remove('on'));
    };
    pillDrag(ev,{
      el:row,label:part.n,sub:String(part.tx||'').replace(/\s+/g,' ').trim().slice(0,50),
      scrollEl:()=>M.querySelector('#sb-ord-list'),
      onMove:(x,y)=>{
        const pane=M.querySelector('.sb-order-pane');
        const rr=pane?pane.getBoundingClientRect():null;
        const over=!!rr&&x>=rr.left&&x<=rr.right&&y>=rr.top&&y<=rr.bottom;
        clearUi();
        if(over){slot=sbInsertSlotFromY(y);paintSbInsert(slot);if(kind==='palette'&&pane)pane.classList.add('sb-drop-hot')}
        else slot=null;
      },
      onDrop:()=>{
        if(slot==null||!sbItem)return;
        const at=sbSlotToAt(slot);
        if(kind==='palette'){insOrd(it,at,pid);sbSel=at;sbChanged()}
        else{
          const dest=at>from?at-1:at;
          if(dest===from)return;
          mvOrd(it,from,dest);sbSel=dest;sbChanged();
        }
      },
      onEnd:clearUi
    });
  }
  M.addEventListener('pointerdown',ev=>{
    const prow=ev.target.closest('.sb-section-row[data-sb-part-row]');
    if(prow){sbDragStart(ev,prow,'palette');return}
    const orow=ev.target.closest('.sb-arr-row[data-sb-ord]');
    if(orow)sbDragStart(ev,orow,'order');
  });


  M.addEventListener('change',ev=>{
    if(ev.target.id==='sb-lps'&&sbItem){
      sbLps=+ev.target.value;sbNeedReflow=true;sbItem.lps=sbLps;
      toast(sbLps?('Sections will split into '+sbLps+' line'+(sbLps>1?'s':'')+' per slide when you close this window.'):'Sections will stay on one slide each.');
    }
  });
}



/* ===== Service Team modal (moved from side panel) ===== */
function openTeamModal(){
  const team=ensureTeam(services[activeService]);
  const row=(label,key,value,index)=>{
    const canRemove=key==='band'&&index!=null;
    const removeBtn=canRemove?`<button type="button" class="team-remove" data-remove-band="${index}" title="Remove band member" aria-label="Remove band member">×</button>`:'';
    return `<tr>
      <td>${e(label)}</td>
      <td><input type="text" data-team-key="${key}" ${index==null?'':'data-band-index="'+index+'"'} data-team-field="name" value="${e(value.name)}" placeholder="Name"></td>
      <td><input type="email" data-team-key="${key}" ${index==null?'':'data-band-index="'+index+'"'} data-team-field="email" value="${e(value.email)}" placeholder="Email"></td>
      <td class="team-remove-cell">${removeBtn}</td>
    </tr>`;
  };
  const root=$('#team-content');
  if(!root)return;
  root.innerHTML=`<p class="note">Service roster team contacts.</p>
    <div class="band-heading"><h3>Team</h3><button id="add-band-member" type="button">+ Band member</button></div>
    <table class="team-table">
      <thead><tr><th>Role</th><th>Name</th><th>Email</th><th></th></tr></thead>
      <tbody>
        ${row('Service leader','leader',team.leader)}
        ${team.band.map((person,i)=>row('Band member '+(i+1),'band',person,i)).join('')}
        ${row('Sound desk','sound',team.sound)}
      </tbody>
    </table>
    <button id="send" class="pri" style="margin-top:10px;width:100%">Send Service Details via Email</button>`;
  $('#team-modal').hidden=false;
}
function closeTeamModal(){const m=$('#team-modal');if(m)m.hidden=true}

/* ===== Speaker notes ===== */
function syncSpeakerNotesUI(){
  const ta=$('#speaker-notes');
  if(!ta)return;
  const part=sls(A)[aj];
  ta.value=(part&&part.notes)||'';
}
function updatePresenterNotes(){
  const body=$("#pv-notes-body");
  if(!body)return;
  const part=sls(A)[aj];
  const notes=(part&&part.notes||'').trim();
  if(notes){
    body.classList.remove('pv-notes-empty');
    body.textContent=notes;
  }else{
    body.classList.add('pv-notes-empty');
    body.textContent='No notes for this slide';
  }
}

/* ===== UI scale (GUI only) ===== */

/* ===== Column resize (Order of Sections + slide strip) ===== */
(function initColResize(){
  const MIN_OUT=140, MAX_OUT=420, MIN_STRIP=150, MAX_STRIP=480;
  function loadW(key, def){
    try{const v=parseInt(localStorage.getItem(key)||'',10);if(v&&v>=100)return v}catch{}
    return def;
  }
  function applyCols(){
    const o=loadW('cp-outline-w',240);
    const s=loadW('cp-strip-w',260);
    document.documentElement.style.setProperty('--outline-w',o+'px');
    document.documentElement.style.setProperty('--strip-w',s+'px');
  }
  applyCols();
  function bind(handleId, cssVar, storageKey, minW, maxW){
    const el=document.getElementById(handleId);
    if(!el)return;
    let dragging=false,startX=0,startW=0;
    el.addEventListener('pointerdown',ev=>{
      if(ev.button!==0)return;
      dragging=true;
      el.classList.add('dragging');
      startX=ev.clientX;
      startW=parseInt(getComputedStyle(document.documentElement).getPropertyValue(cssVar),10)||minW;
      try{el.setPointerCapture(ev.pointerId)}catch{}
      ev.preventDefault();
    });
    el.addEventListener('pointermove',ev=>{
      if(!dragging)return;
      const dx=ev.clientX-startX;
      let w=Math.round(startW+dx);
      w=Math.max(minW,Math.min(maxW,w));
      document.documentElement.style.setProperty(cssVar,w+'px');
    });
    function end(ev){
      if(!dragging)return;
      dragging=false;
      el.classList.remove('dragging');
      const w=parseInt(getComputedStyle(document.documentElement).getPropertyValue(cssVar),10);
      try{localStorage.setItem(storageKey,String(w))}catch{}
    }
    el.addEventListener('pointerup',end);
    el.addEventListener('pointercancel',end);
  }
  bind('resize-outline','--outline-w','cp-outline-w',MIN_OUT,MAX_OUT);
  bind('resize-strip','--strip-w','cp-strip-w',MIN_STRIP,MAX_STRIP);
})();


/* ===== Google Drive: shared library, automatic sync, conflict handling ===== */
const CLOUD_META_KEY='service-slides-sync';
const GSLIDES='application/vnd.google-apps.presentation';
const CPRES_SCHEMA=2;
const SYNC_DEBOUNCE_MS=5000;   // wait this long after the last edit before uploading
const SYNC_POLL_MS=45000;      // check Drive for remote changes this often
const LIBRARY_STALE_MS=20000;  // don't re-list the folder more often than this when the Services page opens
const sleep=ms=>new Promise(r=>setTimeout(r,ms));

let cloudUi={connected:false,email:'',folderId:'',folderName:'',serviceFolderId:'',files:[],error:'',busy:false,msg:'',online:true,authLost:false,lastSyncAt:0,lastAttemptAt:0};
let syncBusy=false,syncAgain=false,pushTimer=null,chipTimer=null,cloudLockWaiters=[];
const syncRt={};     // serviceId -> {state:'syncing'|'offline'|'error', msg, update?:true}
const seenHash={};   // serviceId -> content hash first seen this session (so merely opening an old service never uploads it)

/* ---------- styles for the new UI ---------- */
(function injectCloudStyles(){
  const s=document.createElement('style');
  s.textContent=`
.sync-chip{display:inline-flex;align-items:center;gap:2px;padding:4px 10px;border-radius:999px;font-size:calc(12px * var(--ui-scale));font-weight:600}
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
.home-sub{margin:0 0 10px;font-size:calc(13px * var(--ui-scale));text-transform:uppercase;letter-spacing:.05em;color:var(--mu)}
.conflict-actions{display:grid;gap:8px;margin-top:14px}
.cloud-warn{color:#b26a00;font-size:calc(12px * var(--ui-scale));margin-top:4px}
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
function fmtDriveDate(x){if(!x)return 'unknown';try{const d=new Date(x);const rel=relDayLabel(d);if(rel)return rel+', '+d.toLocaleTimeString(undefined,{hour:'2-digit',minute:'2-digit'});return d.toLocaleString(undefined,{day:'numeric',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'})}catch{return String(x)}}
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
  // A transient empty editor state must not replace the last saved service
  // contents during a background Drive sync.
  const liveSlides=i===activeService?S:null;
  const slides=liveSlides&&liveSlides.length?liveSlides:(svc.slides||[]);
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
      (p.elements||[]).forEach(el=>{if((el.kind==='image'||el.kind==='video')&&el.mediaId&&!has(el.mediaId))out.push(el.kind)});
    });
  });
  return out;
}

/* ---------- talking to Rust, with classification, friendly errors and bounded retry ---------- */
async function cloudInvoke(cmd,args,{timeoutMs=30000}={}){
  try{
    if(!(window.__TAURI_INTERNALS__||window.__TAURI__))return {ok:false,error:'Desktop app only'};
    const {invoke}=await import('@tauri-apps/api/core');
    let timeoutId;
    const res=await Promise.race([
      invoke(cmd,args||{}),
      new Promise((_,reject)=>{timeoutId=setTimeout(()=>reject(new Error(`Google Drive request timed out after ${Math.round(timeoutMs/1000)} seconds. Check your connection and try again.`)),timeoutMs)})
    ]).finally(()=>clearTimeout(timeoutId));
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
async function cloudCall(cmd,args,{retries=3,timeoutMs=30000}={}){
  let delay=1000,last='';
  for(let a=0;a<=retries;a++){
    const res=await cloudInvoke(cmd,args,{timeoutMs});
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
let driveUploadHideTimer=0;
let driveUploadCompleted=false;
function ensureDriveUploadProgress(){
  const html='<div class="drive-upload-progress" hidden role="status" aria-live="polite"><div class="drive-upload-progress-text"><span class="drive-upload-progress-label">Uploading to Drive</span><b>0%</b></div><div class="drive-upload-progress-track"><i></i></div></div>';
  const editorCloud=document.querySelector('#cloud-btn');
  if(editorCloud&&!editorCloud.nextElementSibling?.classList.contains('drive-upload-progress'))editorCloud.insertAdjacentHTML('afterend',html);
  const homeCloud=document.querySelector('#home-cloud-btn'),homeHead=homeCloud?.closest('.home-head');
  if(homeHead&&!document.querySelector('.home-drive-upload-progress'))homeHead.insertAdjacentHTML('afterend',html.replace('class="drive-upload-progress"','class="drive-upload-progress home-drive-upload-progress"'));
}
function setDriveUploadProgress({active=true,bytesSent=0,totalBytes=0,label='Uploading to Drive'}={}){
  ensureDriveUploadProgress();clearTimeout(driveUploadHideTimer);
  const percent=totalBytes?Math.max(0,Math.min(100,Math.round(bytesSent/totalBytes*100))):0;
  document.querySelectorAll('.drive-upload-progress').forEach(root=>{
    root.hidden=!active;
    root.style.setProperty('--upload-progress',percent+'%');
    root.querySelector('.drive-upload-progress-label').textContent=label;
    root.querySelector('b').textContent=percent+'%';
    root.title=label+' · '+percent+'%';
    root.setAttribute('aria-label',label+' '+percent+'%');
  });
}
async function listenForDriveUploadProgress(){
  try{const {listen}=await import('@tauri-apps/api/event');await listen('gdrive-upload-progress',event=>{if(!driveUploadCompleted)setDriveUploadProgress({bytesSent:event.payload?.bytesSent||0,totalBytes:event.payload?.totalBytes||0,label:'Uploading to Drive'})})}catch(err){console.debug('[cloud] upload progress events unavailable',err)}
}
ensureDriveUploadProgress();void listenForDriveUploadProgress();
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
async function withCloudLock(label,fn,{quiet=false,progress=false}={}){
  if(syncBusy){
    if(quiet){syncAgain=true;return false}
    // User-triggered sync actions wait behind any automatic sync instead of
    // appearing to do nothing when the background job currently owns the lock.
    await new Promise(resolve=>cloudLockWaiters.push(resolve));
    return withCloudLock(label,fn,{quiet,progress});
  }
  syncBusy=true;
  let progressEl=null;
  if(progress&&!document.getElementById('manual-sync-progress')){progressEl=document.createElement('div');progressEl.id='manual-sync-progress';progressEl.setAttribute('role','dialog');progressEl.setAttribute('aria-modal','true');progressEl.style.cssText='position:fixed;inset:0;z-index:20000;background:#0009;display:grid;place-items:center';progressEl.innerHTML='<style>@keyframes gd-progress{from{transform:translateX(-120%)}to{transform:translateX(260%)}}</style><div style="width:min(400px,90vw);padding:24px;border-radius:12px;background:var(--pn,#202938);color:var(--ink,#fff);box-shadow:0 12px 50px #0008"><b id="manual-sync-label"></b><div style="height:8px;background:var(--ln,#455);border-radius:8px;overflow:hidden;margin-top:16px"><div style="height:100%;width:35%;background:var(--ac,#4788ff);animation:gd-progress 1s ease-in-out infinite alternate"></div></div></div>';document.body.append(progressEl);$('#manual-sync-label').textContent=label}
  if(!quiet){cloudUi.busy=true;cloudUi.msg=label;cloudUi.error=''}
  refreshSyncUi();
  let ok=false;
  try{await fn();ok=true}
  catch(err){
    if(!(err&&err.cancelled)){handleCloudError(err);if(!quiet&&err.kind!=='conflict')toast(err.message||'Google Drive error.')}
  }
  finally{
    syncBusy=false;cloudUi.busy=false;cloudUi.msg='';progressEl?.remove();
    refreshSyncUi();
    cloudLockWaiters.splice(0).forEach(resolve=>resolve());
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
async function fetchCloudFiles({syncAssets=false}={}){
  if(!cloudUi.connected||!cloudUi.folderId)return;
  try{
    const root=await cloudCall('gdrive_list_presentations',{folderId:cloudUi.folderId});
    const rootFiles=Array.isArray(root.files)?root.files:[];
    let serviceFolder=rootFiles.find(f=>f.mimeType==='application/vnd.google-apps.folder'&&String(f.name||'').toLowerCase()==='services');
    if(!serviceFolder){const made=await cloudCall('gdrive_ensure_subfolder',{parentId:cloudUi.folderId,name:'services'});if(made&&made.id)serviceFolder=made}
    cloudUi.serviceFolderId=serviceFolder&&serviceFolder.id||cloudUi.folderId;
    const sr=cloudUi.serviceFolderId===cloudUi.folderId?root:await cloudCall('gdrive_list_presentations',{folderId:cloudUi.serviceFolderId});
    cloudUi.files=[...rootFiles.filter(f=>/\.cpres$/i.test(f.name||'')),...(Array.isArray(sr.files)?sr.files:[])].filter((f,i,a)=>a.findIndex(x=>x.id===f.id)===i);
    // Keep service sync responsive. Refresh the church song bank in the
    // background at most once per library-stale interval instead of making
    // every Services button click wait for song downloads.
    if(syncAssets)await syncChurchLibraries(rootFiles,{syncAssets});
    else if(Date.now()-(cloudUi.lastSongSyncAt||0)>LIBRARY_STALE_MS){
      cloudUi.lastSongSyncAt=Date.now();
      setTimeout(()=>syncChurchLibraries(rootFiles,{syncAssets:false}).catch(err=>console.warn('[church songs]',err)),0);
    }
    cloudUi.online=true;
    return rootFiles;
  }catch(err){
    if(err.kind==='notfound'||err.kind==='denied'){
      throw Object.assign(new Error('The Drive folder is missing, was moved, or you no longer have access. Open Google Drive options and choose the folder again.'),{kind:err.kind,raw:err.raw});
    }
    throw err;
  }
}

async function syncChurchLibraries(rootFiles,{syncAssets=false}={}){
  const folders={},remoteAssetIds=new Set();
  for(const kind of (syncAssets?['songs','logos']:['songs'])){
    const folder=(rootFiles||[]).find(f=>f.mimeType==='application/vnd.google-apps.folder'&&String(f.name||'').toLowerCase()===kind);
    folders[kind]=folder?((await cloudCall('gdrive_list_presentations',{folderId:folder.id})).files||[]):[];
  }
  LIB.splice(0,LIB.length,...LIB.filter(song=>!song.churchBank));
  const songKey='church-song-bank:'+cloudUi.folderId;let songCache={};try{songCache=JSON.parse(localStorage.getItem(songKey)||'{}')}catch{}
  const currentSongIds=new Set(folders.songs.filter(x=>/\.txt$/i.test(x.name||'')||/^text\/plain$/i.test(x.mimeType||'')).map(x=>x.id));
  for(const f of folders.songs.filter(x=>/\.txt$/i.test(x.name||'')||/^text\/plain$/i.test(x.mimeType||''))){
    try{
      const cached=songCache[f.id];if(cached&&cached.modified===(f.modifiedTime||'')&&cached.song){LIB.push(cached.song);continue}
      const text=await (await downloadDriveFile(f)).text(),lines=text.replace(/\r/g,'').split('\n'),meta={},start=lines.findIndex(x=>!x.trim());
      for(const line of (start<0?lines:lines.slice(0,start))){const m=line.match(/^\s*([^:]+):\s*(.*?)\s*$/);if(m)meta[m[1].toLowerCase().replace(/[^a-z]/g,'')]=m[2]}
      const title=meta.songtitle||String(f.name).replace(/\.txt$/i,'').replace(/[_-]+/g,' ').trim(),lyrics=(start<0?'':lines.slice(start+1).join('\n')).trim();
      if(!title||!lyrics)continue;
      const parsed=parse(lyrics),song={t:title,...parsed,churchBank:true,churchDriveId:f.id,credits:{artist:meta.artistauthor||meta.artist||meta.author||'',authors:meta.writers||'',writers:meta.writers||'',year:meta.releaseyear||'',ccli:meta.cclisong||meta.cclisongnumber||''}};LIB.push(song);songCache[f.id]={modified:f.modifiedTime||'',song};
    }catch(err){console.warn('[church songs] Could not read '+f.name,err)}
  }
  Object.keys(songCache).forEach(id=>{if(!currentSongIds.has(id))delete songCache[id]});try{localStorage.setItem(songKey,JSON.stringify(songCache))}catch{}
  if(!syncAssets){return}
  const logoDriveIds=new Set();
  // Startup/full-service sync may refresh logos, but backgrounds and announcement
  // slides are only listed via their thumbnail sync buttons and downloaded on demand.
  for(const kind of ['logos'])for(const f of folders[kind]){
    const isImage=isDriveImageFile(f),isVideo=/\.(mp4|webm|mov|m4v|ogv|avi|mkv)$/i.test(f.name||'')||/^video\//i.test(f.mimeType||'');
    if(!isImage&&!(kind==='backgrounds'&&isVideo))continue;
    const id='church-'+f.id,meta=mediaMeta.find(x=>x.id===id),category=kind==='logos'?'church-logo':kind==='backgrounds'?'church-background':'church-announcement';
    remoteAssetIds.add(id);
    if(kind==='logos')logoDriveIds.add(f.id);
    if(!mediaURLs[id]||!meta||meta.driveModified!==f.modifiedTime){
      try{const file=await downloadDriveFile(f),blob=new Blob([await file.arrayBuffer()],{type:file.type||mimeOf(f.name)});await mPut(id,blob);if(mediaURLs[id])URL.revokeObjectURL(mediaURLs[id]);mediaURLs[id]=URL.createObjectURL(blob);
      const record={id,name:f.name,kind:isVideo?'video':'image',size:blob.size,createdAt:Date.now(),source:'church-library',libraryType:category,driveModified:f.modifiedTime||''};
        const at=mediaMeta.findIndex(x=>x.id===id);if(at>=0)mediaMeta[at]=record;else mediaMeta.push(record);
      }catch(err){console.warn('[church library] Could not download '+f.name,err)}
    }else{meta.libraryType=category;meta.name=f.name;meta.libraryRemoved=false}
    if(kind==='logos'&&mediaURLs[id]){
      const logoId='church-logo-'+f.id;
      const blob=await mGet(id);if(blob){const data=await new Promise(resolve=>{const rd=new FileReader();rd.onload=()=>resolve(rd.result);rd.onerror=()=>resolve('');rd.readAsDataURL(blob)});if(data){const logo={id:logoId,name:f.name.replace(/\.[^.]+$/,''),src:data,churchLibrary:true,driveId:f.id};const at=churchLogos.findIndex(x=>x.id===logoId);if(at>=0)churchLogos[at]=logo;else churchLogos.push(logo)}}
    }
  }
  mediaMeta.filter(x=>x.source==='church-library'&&x.libraryType==='church-logo').forEach(x=>{x.libraryRemoved=!remoteAssetIds.has(x.id)});
  churchLogos=churchLogos.filter(x=>!x.churchLibrary||logoDriveIds.has(x.driveId));
  if(defaultLogoId&&!churchLogos.some(x=>x.id===defaultLogoId))defaultLogoId=churchLogos[0]?.id||'';
  saveLogos();syncEffectiveLogo();try{renderLogoPicker();updateLogoPreview()}catch{}
  saveMediaMeta();
}

/* ---------- reading / writing the .cpres format ---------- */
async function downloadDriveFile(f){
  // Drive media downloads can take longer than metadata sync, especially for
  // larger backgrounds. Keep the JS timeout aligned with Rust's 120 second HTTP timeout.
  const res=await cloudCall('gdrive_download_file',{fileId:f.id,mimeType:f.mimeType||''},{timeoutMs:125000});
  if(!res.dataBase64)throw cloudErr('Download failed');
  let name=f.name||'file';
  if(f.mimeType===GSLIDES&&!/\.pptx$/i.test(name))name+='.pptx';
  return new File([b64ToBytes(res.dataBase64)],name);
}
/* .cpres ZIP v1: service.json plus media/<stable media id> entries. Stored entries keep
   the package dependency-free and make browser/Drive downloads deterministic. */
const CPRES_ZIP_MIME='application/vnd.church-presenter.service+zip';
function servicePastRetention(svc){
  const m=String(svc&&svc.date||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!m)return false;
  const day=Date.UTC(+m[1],+m[2]-1,+m[3]),now=new Date(),today=Date.UTC(now.getFullYear(),now.getMonth(),now.getDate());
  return today-day>=14*86400000;
}
function serviceCleanupComplete(svc){const link=getLink(svc);return !!(link&&link.cleanupDone)}
const ZIP_CRC_TABLE=(()=>{const t=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=(c&1)?(0xedb88320^(c>>>1)):(c>>>1);t[n]=c>>>0}return t})();
function zipCrc32(bytes){let c=0xffffffff;for(const b of bytes)c=ZIP_CRC_TABLE[(c^b)&255]^(c>>>8);return (c^0xffffffff)>>>0}
function zip16(v){return [v&255,(v>>>8)&255]}
function zip32(v){return [v&255,(v>>>8)&255,(v>>>16)&255,(v>>>24)&255]}
async function makeStoredZip(entries){
  const enc=new TextEncoder(),locals=[],centrals=[];let offset=0;
  for(const entry of entries){
    const name=enc.encode(entry.name),data=entry.bytes instanceof Uint8Array?entry.bytes:new Uint8Array(await entry.bytes.arrayBuffer()),crc=zipCrc32(data);
    const header=new Uint8Array([...zip32(0x04034b50),...zip16(20),...zip16(0x0800),...zip16(0),...zip16(0),...zip16(0),...zip32(crc),...zip32(data.length),...zip32(data.length),...zip16(name.length),...zip16(0),...name]);
    const local=new Uint8Array(header.length+data.length);local.set(header);local.set(data,header.length);
    locals.push(local);
    centrals.push(new Uint8Array([...zip32(0x02014b50),...zip16(20),...zip16(20),...zip16(0x0800),...zip16(0),...zip16(0),...zip16(0),...zip32(crc),...zip32(data.length),...zip32(data.length),...zip16(name.length),...zip16(0),...zip16(0),...zip16(0),...zip16(0),...zip32(0),...zip32(offset),...name]));
    offset+=local.length;
  }
  const centralSize=centrals.reduce((n,x)=>n+x.length,0),count=entries.length;
  const end=new Uint8Array([...zip32(0x06054b50),...zip16(0),...zip16(0),...zip16(count),...zip16(count),...zip32(centralSize),...zip32(offset),...zip16(0)]);
  return new Blob([...locals,...centrals,end],{type:CPRES_ZIP_MIME});
}
function blobChunkStream(blob,chunkSize=4*1024*1024,onChunk=()=>{}){
  let offset=0;
  return new ReadableStream({async pull(controller){
    if(offset>=blob.size){controller.close();return}
    const end=Math.min(offset+chunkSize,blob.size);
    controller.enqueue(new Uint8Array(await blob.slice(offset,end).arrayBuffer()));
    offset=end;
    onChunk(offset);
  }});
}
async function readStoredZip(file){
  const b=new Uint8Array(await file.arrayBuffer()),v=new DataView(b.buffer,b.byteOffset,b.byteLength),out=new Map();let end=-1;
  for(let p=Math.max(0,b.length-65557);p<=b.length-22;p++)if(v.getUint32(p,true)===0x06054b50)end=p;
  if(end<0)throw new Error('This .cpres ZIP is missing its directory.');
  const count=v.getUint16(end+10,true);let p=v.getUint32(end+16,true);const dec=new TextDecoder();
  for(let n=0;n<count;n++){
    if(v.getUint32(p,true)!==0x02014b50)throw new Error('This .cpres ZIP has an invalid directory.');
    const method=v.getUint16(p+10,true),size=v.getUint32(p+24,true),nl=v.getUint16(p+28,true),el=v.getUint16(p+30,true),cln=v.getUint16(p+32,true),lo=v.getUint32(p+42,true),name=dec.decode(b.subarray(p+46,p+46+nl));
    if(name.startsWith('/')||name.split('/').includes('..'))throw new Error('Unsafe path in .cpres archive.');
    const lnl=v.getUint16(lo+26,true),lel=v.getUint16(lo+28,true),start=lo+30+lnl+lel,raw=b.slice(start,start+size);
    if(method===0)out.set(name,raw);else if(method===8&&typeof DecompressionStream!=='undefined'){
      const ds=new DecompressionStream('deflate-raw');out.set(name,new Uint8Array(await new Response(new Blob([raw]).stream().pipeThrough(ds)).arrayBuffer()));
    }else throw new Error('Unsupported compression in .cpres archive.');
    p+=46+nl+el+cln;
  }
  return out;
}
async function buildCpresPackageParts(i,body,applyCloudRetention=false,remoteMedia=null){
  const svc=services[i],slides=body?body.slides:(i===activeService?sanitizeSlidesForSave(S):sanitizeSlidesForSave(svc.slides||[]));
  const expired=applyCloudRetention&&servicePastRetention(svc),data={format:'church-presenter-service',version:2,packaged:true,exportedAt:new Date().toISOString(),service:{id:svc.id,...(body||serviceSyncBody(i))}};
  const mediaFiles=[],ids=new Set(),manifest=[];
  const add=id=>{if(id)ids.add(id)};
  (slides||[]).forEach(it=>{add(it.originalFileId);(it.parts||[]).forEach(p=>{add(p.videoMedia&&p.videoMedia.id);add(p.bgMedia&&p.bgMedia.id);(p.im||[]).forEach(m=>add(m.mediaId));(p.elements||[]).forEach(el=>add(el.mediaId))});add(it.bgMedia&&it.bgMedia.id)});
  const removed=[];
  for(const id of ids){const meta=mediaMetaFor(id);const blob=await mGetFile(id)||await mGet(id);if(blob){if(expired){removed.push(id);continue}const remote=remoteMedia&&remoteMedia[id];manifest.push({id,name:meta.name||id,type:blob.type||mimeOf(meta.name||id),size:blob.size,...(remote?{driveFileId:remote.driveFileId,contentHash:remote.contentHash}: {})});if(!remoteMedia)mediaFiles.push({id,bytes:blob})}}
  data.cleanup={mediaRemoved:removed.length>0,largeMediaRemoved:removed.length>0,removedCount:removed.length,retentionDays:14};
  if(applyCloudRetention)svc.largeMediaRemoved=removed.length>0;
  return {serviceJson:JSON.stringify(data,null,2),manifestJson:JSON.stringify(manifest),mediaFiles};
}
async function buildCpresPackage(i,body,applyCloudRetention=false){
  const parts=await buildCpresPackageParts(i,body,applyCloudRetention),enc=new TextEncoder();
  return makeStoredZip([
    {name:'service.json',bytes:enc.encode(parts.serviceJson)},
    ...parts.mediaFiles.map(m=>({name:'media/'+encodeURIComponent(m.id),bytes:m.bytes})),
    {name:'media-manifest.json',bytes:enc.encode(parts.manifestJson)}
  ]);
}
async function readCpres(file){
  let data,archive=null;
  try{
    if(file.size>=4&&new DataView(await file.slice(0,4).arrayBuffer()).getUint32(0,true)===0x04034b50){archive=await readStoredZip(file);const meta=archive.get('service.json');if(!meta)throw new Error('The archive has no service.json.');data=JSON.parse(new TextDecoder().decode(meta));}
    else data=JSON.parse(await file.text());
  }catch(err){throw kindErr('corrupt','This cloud file is corrupt or is not a Church Presenter service. '+(err&&err.message||''))}
  if(!data||(data.format!=='church-presenter-service'&&!data.service&&!Array.isArray(data.slides)))throw kindErr('corrupt','This cloud file is not a Church Presenter service.');
  if((data.version||1)>CPRES_SCHEMA)throw kindErr('schema','This file was saved by a newer version of Church Presenter. Update the app to open it.');
  const s=data.service||data;
  const slides=Array.isArray(s.slides)?cl(s.slides):[];
  if(archive){
    let manifest=[];try{manifest=JSON.parse(new TextDecoder().decode(archive.get('media-manifest.json')||new Uint8Array()))}catch{}
    const mm=new Map(manifest.map(x=>[x.id,x]));
    for(const [path,bytes] of archive){if(!path.startsWith('media/'))continue;const id=decodeURIComponent(path.slice(6)),meta=mm.get(id)||{};const blob=new Blob([bytes],{type:meta.type||'application/octet-stream'});await mPut(id,blob);mediaURLs[id]=URL.createObjectURL(blob);if(!mediaMeta.some(x=>x.id===id))mediaMeta.push({id,name:meta.name||id,kind:(meta.type||'').startsWith('video/')?'video':'image',size:blob.size,createdAt:Date.now(),source:'service-package'})}
    for(const item of manifest){if(!item.driveFileId||!item.id)continue;let blob=await mGetFile(item.id)||await mGet(item.id);if(!blob){blob=await downloadDriveFile({id:item.driveFileId,name:item.name||item.id,mimeType:item.type||'application/octet-stream'});await mPut(item.id,blob);mediaURLs[item.id]=URL.createObjectURL(blob)}let meta=mediaMetaFor(item.id);Object.assign(meta,{name:item.name||meta.name,driveFileId:item.driveFileId,driveContentHash:item.contentHash||'',driveMediaFingerprint:`${blob.size}:${blob.type||''}:${blob.lastModified||0}`})}
    saveMediaMeta();
  }
  return {id:typeof s.id==='string'?s.id:'',name:(s.name||'').trim(),date:s.date||new Date().toISOString().slice(0,10),tag:s.tag||'',team:s.team?cl(s.team):defaultTeam(),slides,cleanup:data.cleanup||null};
}
function buildCpresJson(i,body){
  body=body||serviceSyncBody(i);
  return JSON.stringify({format:'church-presenter-service',version:CPRES_SCHEMA,exportedAt:new Date().toISOString(),service:{id:services[i].id,...body}},null,2);
}

/* ---------- local backups (so cloud replacements are always recoverable) ---------- */
async function backupService(i){
  const svc=services[i];ensureSvcId(svc);
  const id='backup-'+svc.id+'-'+Date.now().toString(36);
  await mPutFile(id,await buildCpresPackage(i));
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
  svc.largeMediaRemoved=!!(s.cleanup&&(s.cleanup.mediaRemoved||s.cleanup.largeMediaRemoved));
  svc.slides=s.slides.length?cl(s.slides):[mkS('Welcome',['Welcome'],0,6.5)];
  if(i===activeService){S=cl(svc.slides);sel=Math.min(sel,S.length-1);si=0;A=S[sel]||null;aj=0;elSel=null;rAll()}
  else saveState();
}
async function cloudRestoreBackup(i){
  const link=getLink(services[i]);
  const b=link&&link.backups&&link.backups[link.backups.length-1];
  if(!b)return toast('No local backup available.');
  if(!await appConfirm('Replace this service with your local backup from '+fmtDriveDate(b.at)+'?\n\nThe current version is backed up first.',{title:'Restore local backup',confirmLabel:'Restore'}))return;
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
  recordLink(svc,{fileId,remoteRev:remoteRev(info),remoteModified:info.modifiedTime,localHash:hashService(i),canEdit:info.canEdit!==false,webViewLink:info.webViewLink||'',cleanupDone:!!(data.cleanup&&(data.cleanup.mediaRemoved||data.cleanup.largeMediaRemoved))});
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
      if(info.canEdit===false){patchLink(svc,{canEdit:false});throw kindErr('readonly','You only have view access to this Drive file, so changes stay on this computer. Use “Save as copy? to upload them.')}
    }
    syncRt[svc.id]={state:'syncing',msg:'',at:Date.now(),update:(syncRt[svc.id]||{}).update};
    refreshSyncUi();
    const body=serviceSyncBody(i);
    if(!body.slides.length)throw kindErr('emptyservice','This service has no slides, so it was not uploaded. Open the working copy on this computer and sync it again.');
    const hash=hashStr(JSON.stringify(body));     // hash of exactly what we upload; edits made during the upload stay "pending"
    let pkg=await buildCpresPackageParts(i,body,true);
    const tempPaths=[],mediaFiles=[];
    let res;
    try{
      const tempDir=await joinPath(await appDataDir(),'gdrive-tmp');
      const remoteMedia={};
      const fingerprint=m=>driveMediaFingerprint(m.bytes);
      const toSync=pkg.mediaFiles.filter(m=>{const meta=mediaMetaFor(m.id),fp=fingerprint(m);if(meta.driveFileId&&meta.driveMediaFingerprint===fp){remoteMedia[m.id]={driveFileId:meta.driveFileId,contentHash:meta.driveContentHash||''};return false}return true});
      const transferSize=toSync.reduce((size,m)=>size+m.bytes.size,pkg.serviceJson.length+pkg.manifestJson.length);
      let stagedBytes=0;
      for(const media of toSync){
        const tempName='cpres-upload-'+crypto.randomUUID()+'.bin';
        const tempPath=await joinPath(tempDir,tempName);
        tempPaths.push(tempPath);
        setDriveUploadProgress({bytesSent:stagedBytes,totalBytes:transferSize,label:'Preparing service upload'});
        await writeFile(tempPath,blobChunkStream(media.bytes,16*1024*1024,chunkBytes=>setDriveUploadProgress({bytesSent:stagedBytes+chunkBytes,totalBytes:transferSize,label:'Preparing service upload'})),{createNew:true});
        stagedBytes+=media.bytes.size;
        const meta=mediaMetaFor(media.id),name=meta.name||media.id;
        mediaFiles.push({id:media.id,tempName,name,mimeType:media.bytes.type||mimeOf(name)});
      }
      const mediaFolder=await cloudCall('gdrive_ensure_subfolder',{parentId:cloudUi.folderId,name:'media'});
      const synced=mediaFiles.length?await cloudCall('gdrive_sync_service_media',{folderId:mediaFolder.id,mediaFiles},{retries:0,timeoutMs:Math.max(125000,Math.ceil(transferSize/(512*1024))*1000+60000)}):{files:[]};
      for(const item of synced.files||[]){remoteMedia[item.id]={driveFileId:item.driveFileId,contentHash:item.contentHash};const source=pkg.mediaFiles.find(m=>m.id===item.id),meta=mediaMetaFor(item.id);if(source){meta.driveFileId=item.driveFileId;meta.driveContentHash=item.contentHash;meta.driveMediaFingerprint=fingerprint(source)}}
      saveMediaMeta();
      pkg=await buildCpresPackageParts(i,body,true,remoteMedia);
      const packageSize=pkg.serviceJson.length+pkg.manifestJson.length;
      setDriveUploadProgress({bytesSent:0,totalBytes:transferSize,label:'Uploading to Drive'});
      driveUploadCompleted=false;
      res=await cloudCall('gdrive_upload_service_package',{
        name:cpresFileName(svc),
        serviceJson:pkg.serviceJson,
        manifestJson:pkg.manifestJson,
        mediaFiles:[],
        folderId:existingId?null:(cloudUi.serviceFolderId||cloudUi.folderId),
        existingFileId:existingId,
        cpresId:svc.id
      },{retries:0,timeoutMs:Math.max(125000,Math.ceil(transferSize/(512*1024))*1000+60000)});
      driveUploadCompleted=true;
      setDriveUploadProgress({bytesSent:transferSize+packageSize,totalBytes:transferSize+packageSize,label:'Upload complete'});
      driveUploadHideTimer=setTimeout(()=>setDriveUploadProgress({active:false}),1200);
    }catch(err){driveUploadCompleted=true;setDriveUploadProgress({active:false});throw err
    }finally{for(const path of tempPaths)try{await removeFile(path)}catch{}}
    recordLink(svc,{fileId:res.id,remoteRev:remoteRev(res),remoteModified:res.modifiedTime,localHash:hash,canEdit:true,webViewLink:res.webViewLink||(link&&link.webViewLink)||'',cleanupDone:servicePastRetention(svc)});
    delete syncRt[svc.id];
    if(!force&&!asCopy&&existingId){/* silent auto-save */}else toast('Saved to Drive: “'+cpresFileName(svc)+'?.');
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
  toast(kind==='deleted'?'“'+formatServiceTitle(svc)+'? was deleted from Drive. Your local copy is kept.':'You no longer have access to “'+formatServiceTitle(svc)+'? on Drive. Your local copy is kept.');
}
function markConflict(i,remote,open){
  const svc=services[i];
  patchLink(svc,{conflict:true,conflictRemote:{rev:remoteRev(remote),modifiedTime:remote.modifiedTime||''}});
  delete syncRt[svc.id];
  toast('Conflict: “'+formatServiceTitle(svc)+'? was also changed on another computer. Your local changes are kept.');
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
    if(servicePastRetention(svc)&&!link.cleanupDone){await cloudPush(i);return}
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

async function cloudSyncAll({manual=false,quiet=false,syncAssets=false}={}){
  if(live||document.body.classList.contains('presentation-mode'))return false;
  if(!cloudReady()){
    if(manual){toast(cloudUi.connected?'Choose a Drive folder first.':'Connect Google Drive first.');openCloudModal()}
    return false;
  }
  cloudUi.lastAttemptAt=Date.now();
  return withCloudLock('Syncing with Drive…',async()=>{
    await fetchCloudFiles({syncAssets});
    pruneLinks();
    // Download unseen service packages so Drive services are available as
    // ready-to-open local services with all packaged media installed.
    const linkedFiles=new Set(services.map(s=>{const l=getLink(s);return l&&l.fileId}).filter(Boolean));
    for(const f of cloudUi.files.filter(file=>driveKind(file)==='cpres'&&!linkedFiles.has(file.id))){
      if(live)break;
      try{await importCloudService(f,{quiet:true});linkedFiles.add(f.id)}
      catch(err){handleCloudError(err);console.warn('[service sync] Could not import '+(f.name||f.id),err)}
    }
    for(let i=0;i<services.length;i++){
      if(live)break;
      const svc=services[i];
      if(!getLink(svc))continue;
      try{await reconcileService(i,{manual})}
      catch(err){if(err.kind==='offline'||err.kind==='auth')throw err;handleCloudError(err,svc)}
    }
    linkedServices().forEach(s=>{const r=syncRt[s.id];if(r&&(r.state==='offline'||r.state==='error'))clearRtState(s)});
    cloudUi.error='';cloudUi.online=true;cloudUi.lastSyncAt=Date.now();
  },{quiet:quiet&&!manual,progress:manual});
}

const SHARED_BACKGROUNDS_FOLDER_ID='1PFk-CRkNtseZLh7B7cDAY-UaIrmpapZ0';
function isDriveImageFile(file){
  return /^image\//i.test(file.mimeType||'')||/\.(png|jpe?g|gif|webp|bmp|svg|avif|tiff?|heic|heif|ico)$/i.test(file.name||'');
}
const announcementThumbnailJobs=new Set();
let announcementsHaveSynced=false;
async function syncDriveLibraryKind(kind,folderOverride='',{quiet=false,skipSongSync=false}={}){
  if(!cloudReady()){if(!quiet){toast('Connect Google Drive and choose its shared folder first.');openCloudModal()}return}
  await withCloudLock('Syncing '+kind+' from Drive…',async()=>{
    let folderId=folderOverride,rootResult=null;
    if(!folderId){
      const root=await cloudCall('gdrive_list_presentations',{folderId:cloudUi.folderId});rootResult=root;
      const folder=(root.files||[]).find(f=>f.mimeType==='application/vnd.google-apps.folder'&&String(f.name||'').trim().toLowerCase()===kind);
      if(!folder){if(kind==='announcements'){if(!quiet)toast('No announcements folder is available in the selected Drive folder.');return}throw new Error('Could not find a “'+kind+'” folder in the selected Drive folder.')}
      folderId=folder.id;
    }
    const result=folderId===cloudUi.folderId&&rootResult?rootResult:await cloudCall('gdrive_list_presentations',{folderId}),files=Array.isArray(result.files)?result.files:[];
    if(kind==='announcements'){
      const announcements=files.filter(f=>/\.(png|jpe?g|gif|webp|bmp)$/i.test(f.name||'')||/^image\//i.test(f.mimeType||''));
      announcementsHaveSynced=true;
      const needingThumbs=[];
      for(const f of announcements){
        const id='church-'+f.id,entry={id,name:f.name,kind:'image',source:'church-library',libraryType:'church-announcement',driveFileId:f.id,mimeType:f.mimeType||'',driveModified:f.modifiedTime||'',thumbnailId:'thumb-'+f.id,libraryRemoved:false},old=mediaMeta.find(x=>x.id===id);
        if(old&&old.driveModified===(f.modifiedTime||'')){old.libraryRemoved=false;if(mediaURLs[old.thumbnailId||id]||announcementThumbnailJobs.has(entry.thumbnailId))continue}
        if(old)Object.assign(old,entry);else mediaMeta.push(entry);
        needingThumbs.push({file:f,entry});
      }
      // Thumbnail requests used to run one at a time; on a folder with many
      // announcements that meant one full network round trip per image.
      saveMediaMeta();renderMediaLibrary();
      // Fetch announcement thumbnails in batches so the library is ready when
      // startup finishes, while rendering each completed batch progressively.
      await (async()=>{
        for(let offset=0;offset<needingThumbs.length;offset+=8){
          const batch=needingThumbs.slice(offset,offset+8);
          batch.forEach(({entry})=>announcementThumbnailJobs.add(entry.thumbnailId));
          await Promise.all(batch.map(async({file,entry})=>{
            try{
              const thumb=await cloudCall('gdrive_download_thumbnail',{fileId:file.id});
              if(thumb?.dataBase64){const blob=new Blob([b64ToBytes(thumb.dataBase64)],{type:'image/jpeg'});await mPutThumb(entry.thumbnailId,blob);if(mediaURLs[entry.thumbnailId])URL.revokeObjectURL(mediaURLs[entry.thumbnailId]);mediaURLs[entry.thumbnailId]=URL.createObjectURL(blob)}
            }catch(err){console.warn('[announcement thumbnail]',file.name,err)}
            finally{announcementThumbnailJobs.delete(entry.thumbnailId)}
          }));
          renderMediaLibrary();
        }
      })();
      if(!quiet)toast('Found '+announcements.length+' announcement slide'+(announcements.length===1?'':'s')+'.');return;
    }
    if(kind==='backgrounds'){
      const backgrounds=files.filter(f=>/\.(png|jpe?g|gif|webp|bmp|mp4|webm|mov|m4v|ogv|avi|mkv)$/i.test(f.name||'')||/^(image|video)\//i.test(f.mimeType||''));
      for(let offset=0;offset<backgrounds.length;offset+=6){
        await Promise.all(backgrounds.slice(offset,offset+6).map(async f=>{
          const video=/\.(mp4|webm|mov|m4v|ogv|avi|mkv)$/i.test(f.name||'')||/^video\//i.test(f.mimeType||''),id='church-'+f.id,existing=mediaMeta.find(x=>x.id===id);
          if(existing&&existing.driveModified===(f.modifiedTime||'')&&existing.thumbnailId&&mediaURLs[existing.thumbnailId]){existing.libraryRemoved=false;return}
          const entry={id,name:f.name,kind:video?'video':'image',source:'church-library',libraryType:'church-background',driveFileId:f.id,mimeType:f.mimeType||'',driveModified:f.modifiedTime||'',thumbnailId:'thumb-'+f.id,libraryRemoved:false};
          try{const thumb=await cloudCall('gdrive_download_thumbnail',{fileId:f.id});if(thumb?.dataBase64){const blob=new Blob([b64ToBytes(thumb.dataBase64)],{type:'image/jpeg'});await mPutThumb(entry.thumbnailId,blob);if(mediaURLs[entry.thumbnailId])URL.revokeObjectURL(mediaURLs[entry.thumbnailId]);mediaURLs[entry.thumbnailId]=URL.createObjectURL(blob)}}catch(err){console.warn('[background thumbnail]',f.name,err)}
          if(existing)Object.assign(existing,entry);else mediaMeta.push(entry);
        }));
      }
      saveMediaMeta();const picker=$('#background-picker-content');if(picker)picker.dataset.backgroundTab='drive';renderBackgroundPicker();
      toast('Found '+backgrounds.length+' background file'+(backgrounds.length===1?'':'s')+'.');return;
    }
    let downloaded=0,listed=0;
    for(const f of files){
      const image=isDriveImageFile(f);
      const video=/\.(mp4|webm|mov|m4v|ogv|avi|mkv)$/i.test(f.name||'')||/^video\//i.test(f.mimeType||'');
      if(!image&&!(kind==='backgrounds'&&video))continue;
      const id='church-'+f.id,existing=mediaMeta.find(x=>x.id===id);
      if(existing&&existing.driveModified===(f.modifiedTime||'')&&mediaURLs[id]){existing.libraryRemoved=false;existing.libraryType=kind==='logos'?'church-logo':existing.libraryType;continue}
      let blob;
      try{
        const file=await downloadDriveFile(f);blob=new Blob([await file.arrayBuffer()],{type:file.type||f.mimeType||mimeOf(f.name)});
        await mPut(id,blob);if(mediaURLs[id])URL.revokeObjectURL(mediaURLs[id]);mediaURLs[id]=URL.createObjectURL(blob);
      }catch(err){console.warn('[Drive '+kind+'] Could not download '+f.name,err);continue}
      const entry={id,name:f.name,kind:video?'video':'image',size:blob.size,createdAt:existing?.createdAt||Date.now(),source:'church-library',libraryType:kind==='backgrounds'?'church-background':kind==='announcements'?'church-announcement':'church-logo',driveModified:f.modifiedTime||'',libraryRemoved:false};
      const at=mediaMeta.findIndex(x=>x.id===id);if(at<0)mediaMeta.push(entry);else mediaMeta[at]=entry;downloaded++;
    }
    saveMediaMeta();
    if(kind==='backgrounds'){const picker=$('#background-picker-content');if(picker)picker.dataset.backgroundTab='drive';renderBackgroundPicker();}
    if(kind==='announcements')renderMediaLibrary();
    if(kind==='logos'){
      const logoDriveIds=new Set();
      for(const f of files){
        const image=isDriveImageFile(f);
        if(!image)continue;
        const id='church-'+f.id;
        const meta=mediaMeta.find(x=>x.id===id);
        const logoId='church-logo-'+f.id;
        let blob=await mGet(id);
        // A failed or interrupted media download must not be reported as a
        // synced logo. Retry here so the settings picker only receives usable
        // image data, even when the general media sync path failed earlier.
        if(!blob){
          try{
            const file=await downloadDriveFile(f);
            blob=new Blob([await file.arrayBuffer()],{type:file.type||f.mimeType||mimeOf(f.name)});
            await mPut(id,blob);
            if(mediaURLs[id])URL.revokeObjectURL(mediaURLs[id]);
            mediaURLs[id]=URL.createObjectURL(blob);
            const entry={id,name:f.name,kind:'image',size:blob.size,createdAt:meta?.createdAt||Date.now(),source:'church-library',libraryType:'church-logo',driveModified:f.modifiedTime||'',libraryRemoved:false};
            if(meta)Object.assign(meta,entry);else mediaMeta.push(entry);
          }catch(err){console.warn('[Drive logos] Could not download '+f.name,err);continue}
        }
        if(blob){
          const data=await new Promise(resolve=>{const rd=new FileReader();rd.onload=()=>resolve(rd.result);rd.onerror=()=>resolve('');rd.readAsDataURL(blob)});
          if(data){
            const logo={id:logoId,name:f.name.replace(/\.[^.]+$/,''),src:data,churchLibrary:true,driveId:f.id};
            const existingLogo=churchLogos.findIndex(x=>x.id===logoId);
            if(existingLogo>=0)churchLogos[existingLogo]=logo;else churchLogos.push(logo);
            logoDriveIds.add(f.id);
            if(meta)meta.libraryRemoved=false;
          }
        }
      }
      saveMediaMeta();
      churchLogos=churchLogos.filter(x=>!x.churchLibrary||logoDriveIds.has(x.driveId));
      if(defaultLogoId&&!churchLogos.some(x=>x.id===defaultLogoId))defaultLogoId=churchLogos[0]?.id||'';
      saveLogos();syncEffectiveLogo();renderLogoPicker();updateLogoPreview();
      // The Settings logo sync is also the explicit refresh action for the
      // church song bank, so bypass the normal stale-library timer here.
      if(!skipSongSync)await syncChurchLibraries(rootResult?.files||[]);
      cloudUi.lastSongSyncAt=Date.now();
      toast('Synced '+logoDriveIds.size+' image'+(logoDriveIds.size===1?'':'s')+' from the Drive logos folder.');
      return;
    }
    toast(downloaded?'Downloaded '+downloaded+' '+kind+' item'+(downloaded===1?'':'s')+'.':'No new '+kind+' files to download.');
  },{progress:!quiet});
}
async function syncDriveBackgrounds(){
  while(syncBusy)await new Promise(resolve=>cloudLockWaiters.push(resolve));
  return syncDriveLibraryKind('backgrounds',SHARED_BACKGROUNDS_FOLDER_ID);
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
  $('#conflict-body').innerHTML=`<p style="margin:0 0 8px">“${e(formatServiceTitle(svc))}? has been changed on another computer.</p>
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
    if(!await appConfirm('Replace this computer’s version with the cloud version?\n\nYour unsynced local changes will be replaced. A local backup is kept and can be restored from the Google Drive menu.',{title:'Use cloud version',confirmLabel:'Replace'}))return;
    closeConflict();
    await withCloudLock('Downloading the cloud version…',()=>cloudPull(i,link.fileId));
  }else if(action==='mine'){
    if(!await appConfirm('Overwrite the cloud version with yours?\n\nThe other computer’s changes will be replaced in Drive (they stay in Drive’s version history).', {title:'Keep this version',confirmLabel:'Overwrite'}))return;
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
  return `<div class="service-card" style="margin-bottom:14px;flex-wrap:wrap"><div class="info"><h3 style="font-size:calc(14px * var(--ui-scale))">This service: ${e(formatServiceTitle(svc))}</h3>
    <small class="sync-line s-${st.key}"><i></i>${e(st.label)}${link&&link.syncedAt?' · last synced '+e(fmtDriveDate(link.syncedAt)):''}</small>
    ${st.detail?`<small style="color:#c22">${e(st.detail)}</small>`:''}
    ${miss.length?`<div class="cloud-warn">⚠ ${miss.length} media item${miss.length===1?'':'s'} (${e([...new Set(miss)].slice(0,4).join(', '))}) not on this computer yet. Sync the service package to download its included media.</div>`:''}
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
      <p class="note" style="margin:8px 0 0">Paste the church Drive folder link or ID. Sync reads the shared songs, logos, backgrounds, and announcements folders and stores service packages in <b>services</b>. Service media stays in the package for 14 days after the service date, then expires from Drive copies.</p>
      <label style="display:flex;align-items:center;gap:8px;margin-top:8px"><input type="checkbox" id="cloud-auto-upload" ${auto?'checked':''}> Automatically save edited services to Drive</label>`:'';
  if(!lib)return;
  if(!cloudUi.connected){lib.innerHTML='';return}
  let html=thisServiceCard(dis);
  if(!cloudUi.folderId)html+='<p class="note">Choose a presentation folder to get started.</p>';
  else html+='<p class="note">Services stored in your Drive folder appear on the <b>Services</b> page. Use the ↻ button there to sync.</p>';
  const slides=cloudUi.files.filter(f=>driveKind(f)==='slides');
  if(slides.length){
    html+=`<details style="margin-top:12px"><summary style="cursor:pointer;color:var(--mu)">Import a PowerPoint / PDF / Google Slides file from Drive (${slides.length})</summary>`
      +slides.map(f=>`<div class="service-card" style="margin-top:8px"><div class="info"><h3 style="font-size:calc(14px * var(--ui-scale))">${e(f.name)}</h3><small>${e(fmtDriveDate(f.modifiedTime))}</small></div><button type="button" class="pri" data-cloud-import="${e(f.id)}"${dis}>Import</button></div>`).join('')
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
function textEntryPopup({title,label,value='',placeholder=''}){
  return new Promise(resolve=>{
    const layer=document.createElement('div');layer.className='app-modal';layer.style.cssText='position:fixed;inset:0;z-index:100000;display:grid;place-items:center;background:#0009;padding:16px';
    layer.innerHTML=`<div class="app-dialog" role="dialog" aria-modal="true" style="width:min(480px,100%);padding:20px"><h2 style="margin:0 0 14px">${e(title)}</h2><label style="display:block;margin-bottom:16px">${e(label)}<input data-popup-value type="text" value="${e(value)}" placeholder="${e(placeholder)}" style="display:block;width:100%;margin-top:7px"></label><div style="display:flex;justify-content:flex-end;gap:8px"><button type="button" data-popup-cancel>Cancel</button><button type="button" class="pri" data-popup-done>Done</button></div></div>`;
    const input=layer.querySelector('[data-popup-value]');let settled=false;
    const finish=result=>{if(settled)return;settled=true;layer.remove();resolve(result)};
    layer.addEventListener('click',ev=>{if(ev.target===layer||ev.target.closest('[data-popup-cancel]'))finish(null);else if(ev.target.closest('[data-popup-done]'))finish(input.value)});
    input.addEventListener('keydown',ev=>{if(ev.key==='Enter'){ev.preventDefault();finish(input.value)}if(ev.key==='Escape'){ev.preventDefault();finish(null)}});
  document.body.appendChild(layer);input.focus();input.select();
  });
}
async function cloudPickFolder(){
  const input=await textEntryPopup({title:'Connect church Drive folder',label:'Paste the Google Drive folder link or ID',value:cloudUi.folderId||'',placeholder:'https://drive.google.com/drive/folders/…'});
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
    const s=await parseCpresFile(await downloadDriveFile(f));
    const svc=serviceFromCpres(s,{name:f.name,id:f.cpresId||s.id||null});
    if(svc.id&&services.some(x=>x.id===svc.id))svc.id=null;
    ensureSvcId(svc);ensureTeam(svc);
    services.push(svc);
    const i=services.length-1;
    recordLink(svc,{fileId:f.id,remoteRev:remoteRev(f),remoteModified:f.modifiedTime,localHash:hashService(i),canEdit:f.canEdit!==false,webViewLink:f.webViewLink||''});
    saveState();
    toast('Downloaded “'+formatServiceTitle(svc)+'?. It now works offline.');
  });
}
async function importCloudService(f,{quiet=false}={}){
  if(services.some(s=>{const l=getLink(s);return l&&l.fileId===f.id}))return null;
  // Use the same parser and service defaults as manual .cpres import.
  const s=await parseCpresFile(await downloadDriveFile(f));
  const svc=serviceFromCpres(s,{name:f.name,id:f.cpresId||s.id||null});
  if(svc.id&&services.some(x=>x.id===svc.id))svc.id=null;
  ensureSvcId(svc);ensureTeam(svc);
  services.push(svc);
  const i=services.length-1;
  recordLink(svc,{fileId:f.id,remoteRev:remoteRev(f),remoteModified:f.modifiedTime,localHash:hashService(i),canEdit:f.canEdit!==false,webViewLink:f.webViewLink||'',cleanupDone:!!(s.cleanup&&(s.cleanup.mediaRemoved||s.cleanup.largeMediaRemoved))});
  saveState();
  if(!quiet)toast('Downloaded '+formatServiceTitle(svc)+'. It now works offline.');
  return svc;
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
    if(!await appConfirm('Keep this service on this computer only? The file in Drive is not deleted.',{title:'Unlink service',confirmLabel:'Keep local'}))return;
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
  if(head)head.insertAdjacentHTML('afterbegin','<button type="button" id="home-cloud-btn" class="icon-btn" title="Google Drive" aria-label="Google Drive"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M7 18a4.5 4.5 0 0 1 0-9 5.5 5.5 0 0 1 10.2-1.8A4 4 0 1 1 18 18H7z"/></svg></button>');
  ensureDriveUploadProgress();
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

/* Add a section using the Church Logo layout configured in Settings. */
function addChurchLogoSection(){
  pushUndo();
  const cur=S[sel];
  const part={id:uid(),n:'Church Logo',tx:'',logo:true};
  if(cur&&!isSongItem(cur)&&!isPptHandoffItem(cur)){
    cur.parts.push(part);
    const at=Math.min(cur.ord.length,(cur===A?aj:cur.ord.length-1)+1);
    insOrd(cur,at,part.id);
    A=cur;sel=S.indexOf(cur);si=at;aj=at;
    toast('Added church logo slide.');
  }else{
    const it=mkS(nextSectionName(),[''],0,6.5);
    it.parts=[part];it.ord=[part.id];
    S.splice(Math.max(0,sel)+1,0,it);
    sel=S.indexOf(it);A=it;si=0;aj=0;
    toast('Added new section with a church logo slide.');
  }
  clr=null;elSel=null;tab='Structure';
  if(services[activeService])services[activeService].slides=sanitizeSlidesForSave(S);
  saveState();rAll();
  setTimeout(()=>{scrollNewItemToTop(sel);scrollCurrentSlideIntoView()},0);
}

/* ---------- editor layout: Add section / Add slide, Add tab removed ---------- */
function nextSectionName(){
  const used=new Set();
  S.forEach(it=>{const m=/^Section (\d+)$/.exec(it.t||'');if(m)used.add(+m[1])});
  let n=1;while(used.has(n))n++;
  return 'Section '+n;
}


function addSectionAtEnd() {
  pushUndo();

  const it = mkS(
    nextSectionName(),
    ['Type your text here'],
    0,
    6.5
  );

  S.push(it);

  sel = S.length - 1;
  A = it;
  si = 0;
  aj = 0;
  tab = 'Structure';

  if (services[activeService]) {
    services[activeService].slides =
      sanitizeSlidesForSave(S);
  }

  saveState();
  rAll();

  requestAnimationFrame(() => {
    if (typeof scrollNewItemToTop === 'function') {
      scrollNewItemToTop(sel);
    }
  });

  if (typeof toast === 'function') {
    toast('Added section to end of service');
  }
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
  if(isSongItem(A)){toast('Songs: use “Add? in the structure builder to add a section.');openStruct();return}
  addNewSlideInItem(S.indexOf(A),aj);
}
const _rOutlineOrig=rOutline;
rOutline=function(){
  _rOutlineOrig();
  const o=$('#outline');if(!o)return;
  o.insertAdjacentHTML('beforeend','<div class="outline-add"><button type="button" id="outline-add-section" class="pri">+ Add section</button><button type="button" id="outline-add-more" title="Other section types" aria-label="Other section types">▾</button></div>');
};
const _rStripOrig=rStrip;
rStrip=function(){
  _rStripOrig();
  const st=$('#strip');if(!st||!S.length)return;
  st.insertAdjacentHTML('beforeend','<div class="strip-add"></div>');
  placeStripAdd();
};
const _rPanelOrig=rPanel;
rPanel=function(){
  if(tab==='Add')tab='Structure';       // service-level "Add" moved to the section/slide columns
  _rPanelOrig();
  const b=document.querySelector('#panel [data-tab="Add"]');if(b)b.remove();
  try{decorateVideoPanel()}catch(err){console.warn(err)}
};
function ensureAddMenu(){
  
if (!$('#add-menu')) {
  document.body.insertAdjacentHTML('beforeend', `
    <div id="add-menu" class="outline-menu" role="menu" hidden>
      <button type="button" data-add-menu="section">New Section</button>
      <button type="button" data-add-menu="song">Song</button>
      <button type="button" data-add-menu="import">Import slides</button>
      <button type="button" data-add-menu="logo">Church logo</button>
    </div>
  `);
}
  if(!$('#add-video-slide-file'))document.body.insertAdjacentHTML('beforeend','<input type="file" id="add-video-slide-file" accept="video/*" hidden>');
}
function hideAddMenu(){const m=$('#add-menu');if(m){m.hidden=true;m.classList.remove('open')}}
function toggleAddMenu(anchor){
  ensureAddMenu();
  const m=$('#add-menu');
  if(m.classList.contains('open')){hideAddMenu();return}
  m.querySelectorAll('[data-add-menu="logo"]').forEach(x=>x.remove());
  if(anchor&&anchor.id==='strip-add-more')m.insertAdjacentHTML('beforeend','<button type="button" data-add-menu="logo">Church logo</button>');
  m.hidden=false;m.classList.add('open');
  const r=anchor.getBoundingClientRect(),w=m.offsetWidth||180,h=m.offsetHeight||300;
  m.style.left=Math.max(6,Math.min(r.left,window.innerWidth-w-8))+'px';
  const below=r.bottom+6;m.style.top=(below+h<window.innerHeight-8?below:Math.max(6,r.top-h-6))+'px';
}

/* ---------- one click handler for everything above ---------- */
document.addEventListener('click',ev=>{
  const t=ev.target;
  if(!t||!t.closest)return;
  if(!t.closest('#add-menu')&&!t.closest('#outline-add-more')&&!t.closest('#strip-add-more'))hideAddMenu();
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
  if(b.id==='strip-edit-structure'){hideAddMenu();if(A){sel=Math.max(0,S.indexOf(A));si=aj}openStruct();return}
  if(b.id==='strip-add-slide-btn'){stripAddSlide();return}
  if(b.id==='outline-add-more'||b.id==='strip-add-more'){toggleAddMenu(b);return}
  
  if(d.addMenu){
    const k=d.addMenu;
    hideAddMenu();

    if(k==='section')addSectionAtEnd();
    else if(k==='song')openLibraryModal();
    else if(k==='video')$('#add-video-slide-file').click();
    else if(k==='import'){
      const f=$('#import-slides-file');
      if(f)f.click();
    }
    else if(k==='logo')addChurchLogoSection();
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
  window.__cpStartupSyncDone=false;
  window.__cpStartupSyncPromise=(async()=>{
    try{
      await refreshCloudStatus();
      if(!cloudReady())return;
      await cloudSyncAll({quiet:true});
      if(!cloudReady()||cloudUi.online===false)return;
      await syncDriveLibraryKind('logos','',{quiet:true,skipSongSync:true});
    }catch(err){console.warn('[startup sync] Drive sync did not finish:',err)}
    finally{window.__cpStartupSyncDone=true}
  })();
}


/* ===== .cpres export / import ===== */
function openShareModal(){
  const m=$('#share-modal');
  if(!m)return;
  m.hidden=false;
  try{$('#share-export-btn')?.focus()}catch{}
}
function closeShareModal(){
  const m=$('#share-modal');
  if(m)m.hidden=true;
}
async function exportCurrentServiceCpres(){
  try{
    const svc=services[activeService];
    if(!svc){toast('No service to export.');return}
    svc.slides=sanitizeSlidesForSave(S);
    const blob=await buildCpresPackage(activeService,{name:svc.name||'',date:svc.date||'',tag:svc.tag||'',team:cl(svc.team||defaultTeam()),slides:sanitizeSlidesForSave(svc.slides||[])});
    const cleanPart=value=>String(value||'').trim().replace(/[\\/:*?"<>|]/g,'-').replace(/[. ]+$/,'');
    const tag=cleanPart(svc.tag),name=cleanPart(svc.name);
    const base=(tag&&name?tag+' - '+name:tag||name||'Service').slice(0,120);
    const filename=base+'.cpres';
    const a=document.createElement('a');
    a.href=URL.createObjectURL(blob);
    a.download=filename;
    document.body.appendChild(a);a.click();a.remove();
    setTimeout(()=>URL.revokeObjectURL(a.href),2000);
    toast('Exported '+filename);
    closeShareModal();
  }catch(err){
    console.error(err);
    toast('Export failed: '+((err&&err.message)||err));
  }
}
async function parseCpresFile(file){
  const src=await readCpres(file);
  return {...src,name:(src.name||'Imported service').trim()||'Imported service'};
}
function serviceFromCpres(src,{name=src.name,id=null}={}){
  const slides=src.slides.length?src.slides:[mkS('Welcome',['Welcome'],0,6.5)];
  const fallbackName=String(name||'').replace(/\.cpres$/i,'').trim()||'Imported service';
  const svc={id,name:src.name&&src.name!=='Imported service'?src.name:fallbackName,date:src.date,tag:src.tag,slides,team:src.team||defaultTeam(),lastEdited:Date.now()};
  ensureTeam(svc);
  return svc;
}
/** Remap part ids so merged sections never clash with existing ones. */
function remapServiceItemIds(item){
  const copy=cl(item);
  if(Array.isArray(copy.parts)){
    const map={};
    copy.parts=copy.parts.map(p=>{
      const nid=uid();
      map[p.id]=nid;
      const p2={...p,id:nid};
      if(Array.isArray(p2.elements))p2.elements=p2.elements.map(el=>({...cl(el),id:uid()}));
      return p2;
    });
    copy.ord=(copy.ord||[]).map(id=>map[id]||id);copy.parts.forEach(p=>{if(p.of&&map[p.of])p.of=map[p.of]});
  }
  return copy;
}
async function importCpresFile(file){
  if(!file)return;
  try{
    const src=await parseCpresFile(file);
    const svc=serviceFromCpres(src);
    services.push(svc);
    activeService=services.length-1;
    S=cl(svc.slides);sel=0;si=0;A=S[0]||null;aj=0;
    saveState();
    try{closeSettingsAndRefresh()}catch{}
    document.body.classList.remove('home-mode');
    rAll();
    try{closeShareModal()}catch{}
    toast('Imported “'+formatServiceTitle(svc)+'?.');
  }catch(err){
    console.error(err);
    toast('Import failed: '+((err&&err.message)||err));
  }
}
async function importCpresPath(path){
  const started=Date.now();
  const overlay=showCpresOpenProgress();
  try{
    const encoded=await invoke('read_cpres_file',{path});
    const binary=atob(encoded),bytes=new Uint8Array(binary.length);
    for(let i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i);
    await importCpresFile(new File([bytes],String(path).split(/[\\/]/).pop()||'service.cpres'));
  }catch(err){console.error('[open cpres]',err);toast('Could not open service file: '+((err&&err.message)||err))}
  finally{await new Promise(resolve=>setTimeout(resolve,Math.max(0,3000-(Date.now()-started))));overlay.remove()}
}
function showCpresOpenProgress(){
  if(!document.getElementById('cpres-open-styles')){
    const style=document.createElement('style');style.id='cpres-open-styles';
    style.textContent='#cpres-open-progress,#cpres-present-confirm{position:fixed;inset:0;z-index:2147483000;display:flex;align-items:center;justify-content:center;background:rgba(3,8,18,.72);backdrop-filter:blur(5px);font-family:inherit}#cpres-open-progress .cpres-open-card,#cpres-present-confirm .cpres-present-card{width:min(440px,calc(100vw - 40px));padding:28px;border:1px solid #334155;border-radius:16px;background:#111c30;color:#f5f7fb;box-shadow:0 24px 80px #0009}#cpres-open-progress .cpres-open-card>b{display:block;font-size:20px;margin-bottom:20px}.cpres-open-track{height:8px;border-radius:8px;background:#27364e;overflow:hidden}.cpres-open-track i{display:block;height:100%;width:0;background:#65c989;border-radius:8px;animation:cpres-opening 3s ease-out forwards}@keyframes cpres-opening{to{width:100%}}#cpres-present-confirm .cpres-present-card h2{margin:0 0 10px;font-size:22px}#cpres-present-confirm .cpres-present-card p{margin:0;color:#bdc8d8;line-height:1.5}.cpres-present-actions{display:flex;justify-content:flex-end;gap:10px;margin-top:24px}.cpres-present-actions button{border:1px solid #43516a;border-radius:8px;padding:10px 16px;color:#f5f7fb;background:#202d42;font:inherit;cursor:pointer}.cpres-present-actions button.cpres-present-open{background:#35a765;border-color:#35a765;color:#06150b;font-weight:700}.cpres-present-actions button:hover{filter:brightness(1.12)}';
    document.head.appendChild(style);
  }
  const old=document.getElementById('cpres-open-progress');if(old)old.remove();
  const box=document.createElement('div');box.id='cpres-open-progress';
  box.innerHTML='<div class="cpres-open-card"><b>Importing Service</b><div class="cpres-open-track"><i></i></div></div>';
  document.body.appendChild(box);return box;
}
async function waitForStartupSequence(){
  // Cold-start file associations arrive before the boot overlay has cleared.
  // Wait for it to finish before showing the service import progress overlay.
  while(document.getElementById('cp-boot'))await new Promise(resolve=>setTimeout(resolve,100));
}
function showCpresPresentConfirm(onOpen,onCancel){
  const old=document.getElementById('cpres-present-confirm');if(old)old.remove();
  const box=document.createElement('div');box.id='cpres-present-confirm';
  box.innerHTML='<div class="cpres-present-card"><h2>Open another service?</h2><p>A service is being presented. Opening this file will end the current presentation.</p><div class="cpres-present-actions"><button type="button" data-cpres-cancel>Cancel</button><button type="button" class="cpres-present-open" data-cpres-open>Open service</button></div></div>';
  box.addEventListener('click',ev=>{if(ev.target===box||ev.target.closest('[data-cpres-cancel]')){box.remove();onCancel?.()}else if(ev.target.closest('[data-cpres-open]')){box.remove();onOpen()}});
  document.body.appendChild(box);
}
if(isTauri()&&!IS_AUDIENCE){
  (async()=>{
    const {listen}=await import('@tauri-apps/api/event');
    const seen=new Set();
    const openPaths=(paths,{waitForBoot=false}={})=>{for(let path of paths||[]){try{if(/^file:/i.test(path)){path=decodeURIComponent(new URL(path).pathname);if(/^\/[A-Za-z]:\//.test(path))path=path.slice(1)}}catch{};if(!String(path).toLowerCase().endsWith('.cpres')||seen.has(path))continue;
      if(isPresentingUI()){
        seen.add(path);
        showCpresPresentConfirm(()=>{endPresent();void importCpresPath(path).finally(()=>seen.delete(path))},()=>seen.delete(path));
      }else{seen.add(path);void (async()=>{if(waitForBoot)await waitForStartupSequence();await importCpresPath(path)})().finally(()=>seen.delete(path))}
    }};
    await listen('open-cpres-files',event=>openPaths(event.payload));
    const deepLink=await import('@tauri-apps/plugin-deep-link');
    await deepLink.onOpenUrl(openPaths);
    openPaths(await deepLink.getCurrent()||[]);
    const paths=await invoke('take_pending_cpres_files');
    setTimeout(()=>openPaths(paths,{waitForBoot:true}),500);
  })().catch(err=>console.warn('[open cpres] setup failed',err));
}
/** Merge sections from a .cpres into the currently open service (join presentations). */
async function importCpresMergeIntoCurrent(file){
  if(!file)return;
  if(!services.length||activeService<0){
    // No open service — fall back to creating a new one
    return importCpresFile(file);
  }
  try{
    const src=await parseCpresFile(file);
    const incoming=src.slides||[];
    if(!incoming.length)throw new Error('That .cpres file has no sections.');
    pushUndo();
    const at=Math.max(0,Math.min(pendingSectionInsert!=null?pendingSectionInsert:S.length,S.length));
    pendingSectionInsert=null;
    incoming.forEach((it,n)=>{S.splice(at+n,0,remapServiceItemIds(it))});
    sel=at;si=0;A=S[sel]||S[S.length-1];aj=0;
    if(services[activeService]){
      services[activeService].slides=sanitizeSlidesForSave(S);
      touchServiceEdited();
    }
    saveState();
    rAll();
    setTimeout(()=>scrollNewItemToTop(sel),0);
    toast('Added '+incoming.length+' section'+(incoming.length===1?'':'s')+' from “'+src.name+'?.');
  }catch(err){
    console.error(err);
    toast('Import failed: '+((err&&err.message)||err));
  }
}
document.addEventListener('click',ev=>{
  const t=ev.target;
  if(!t)return;
  if(t.id==='export-cpres-btn'||(t.closest&&t.closest('#export-cpres-btn'))){
    openShareModal();
    return;
  }
  if(t.id==='share-close'||t.id==='share-modal'){closeShareModal();return}
  if(t.id==='share-team-btn'||(t.closest&&t.closest('#share-team-btn'))){
    closeShareModal();
    openTeamModal();
    return;
  }
  if(t.id==='share-export-btn'||(t.closest&&t.closest('#share-export-btn'))){
    exportCurrentServiceCpres();
    return;
  }
  if(t.id==='share-import-btn'||(t.closest&&t.closest('#share-import-btn'))){
    const inp=$('#import-cpres-file');
    if(inp){inp.value='';inp.click()}
    return;
  }
});
document.addEventListener('change',ev=>{
  if(ev.target&&ev.target.id==='import-cpres-file'){
    const f=ev.target.files&&ev.target.files[0];
    ev.target.value='';
    importCpresFile(f);
  }
});
document.addEventListener('keydown',ev=>{
  if(ev.key==='Escape'){
    const m=$('#share-modal');
    if(m&&!m.hidden){ev.preventDefault();closeShareModal()}
  }
});
/* ===== Close app confirmation ===== */
let closeConfirmOpen=false;
let closeConfirmResolver=null;
function showCloseConfirm(){
  return new Promise(resolve=>{
    const m=$('#close-confirm-modal');
    if(!m){resolve(true);return}
    closeConfirmOpen=true;
    closeConfirmResolver=resolve;
    m.hidden=false;
    m.style.zIndex='2147483647';
    try{import('@tauri-apps/api/window').then(({getCurrentWindow})=>getCurrentWindow().setAlwaysOnTop(true).catch(()=>{}))}catch{}
    try{$('#close-confirm-yes')?.focus()}catch{}
  });
}
function finishCloseConfirm(yes){
  const m=$('#close-confirm-modal');
  if(m)m.hidden=true;
  try{import('@tauri-apps/api/window').then(({getCurrentWindow})=>getCurrentWindow().setAlwaysOnTop(false).catch(()=>{}))}catch{}
  closeConfirmOpen=false;
  const r=closeConfirmResolver;
  closeConfirmResolver=null;
  if(r)r(!!yes);
}
document.addEventListener('click',ev=>{
  if(ev.target?.id==='close-confirm-yes')finishCloseConfirm(true);
  if(ev.target?.id==='close-confirm-no')finishCloseConfirm(false);
});
document.addEventListener('keydown',ev=>{
  if(!closeConfirmOpen)return;
  if(ev.key==='Enter'||ev.key===' '){ev.preventDefault();finishCloseConfirm(true)}
  if(ev.key==='Escape'){ev.preventDefault();finishCloseConfirm(false)}
}, true);

(async function bindCloseConfirm(){
  const quit=async()=>{
    try{await invoke('quit_app');return}catch(err){console.warn('quit_app',err)}
    try{const {getCurrentWindow}=await import('@tauri-apps/api/window');await getCurrentWindow().destroy();return}catch(err){console.warn('destroy',err)}
    try{const {getCurrentWindow}=await import('@tauri-apps/api/window');await getCurrentWindow().close()}catch(err){console.warn('close',err)}
  };
  const ask=async()=>{
    if(closeConfirmOpen)return;
    const ok=await showCloseConfirm();
    if(ok)await quit();
  };
  // Preferred: Rust holds the close and emits "confirm-close"
  try{
    const {listen}=await import('@tauri-apps/api/event');
    await listen('confirm-close',()=>{ask()});
    await invoke('close_guard_ready');
    return;
  }catch(err){console.warn('Rust close guard unavailable, falling back to JS handler',err)}
  // Fallback: JS-side close handler
  try{
    const {getCurrentWindow}=await import('@tauri-apps/api/window');
    const w=getCurrentWindow();
    if(w&&typeof w.onCloseRequested==='function'){
      await w.onCloseRequested(async(event)=>{event.preventDefault();await ask()});
    }
  }catch(err){console.warn('close confirm bind',err)}
})();

function applyUiScale(scale){
  const s=parseFloat(scale)||1;
  document.documentElement.style.setProperty('--ui-scale',String(s));
  try{localStorage.setItem('service-slides-ui-scale',String(s))}catch{}
  const sel=$('#ui-scale-setting');
  if(sel&&sel.value!==String(s))sel.value=String(s);
}
(function initUiScale(){
  let s=1;
  try{s=parseFloat(localStorage.getItem('service-slides-ui-scale')||'1')||1}catch{}
  if(![0.75,1,1.5].includes(s)){if(s===0.5)s=0.75;else s=1;}
  applyUiScale(s);
})();

function loadEmailFromPrefs(){
  try{return JSON.parse(localStorage.getItem('service-slides-email-from')||'{}')}catch{return {}}
}
function saveEmailFromPrefs(from,fromName){
  localStorage.setItem('service-slides-email-from',JSON.stringify({fromEmail:from||'',fromName:fromName||''}));
}
function buildServiceEmailContent(){
  const team=ensureTeam(services[activeService]);
  const svc=services[activeService];
  const people=[
    {role:'Leader', ...(team.leader||{})},
    ...(team.band||[]).map((p,i)=>({role:'Band '+(i+1),...p})),
    {role:'Sound', ...(team.sound||{})}
  ].filter(p=>(p.email||'').trim()||(p.name||'').trim());
  const itemLines=S.map((it,i)=>`${i+1}. ${it.t}`).join('\n');
  const dateText=(()=>{
    const match=String(svc.date||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if(!match)return '';
    const date=new Date(+match[1],+match[2]-1,+match[3]);
    if(date.getFullYear()!==+match[1]||date.getMonth()!==+match[2]-1||date.getDate()!==+match[3])return '';
    const day=date.getDate(),lastTwo=day%100;
    const suffix=lastTwo>=11&&lastTwo<=13?'th':day%10===1?'st':day%10===2?'nd':day%10===3?'rd':'th';
    return `${day}${suffix} ${date.toLocaleString('en-AU',{month:'long'})}`;
  })();
  const leader=(team.leader?.name||'').trim();
  const tag=(svc.tag||'').trim();
  const serviceLabel=tag?tag+' service':'service';
  const dateLabel=dateText?' on the '+dateText:'';
  const subject=`${svc.name||'Service'} — ${svc.date||'upcoming'}`;
  const body=`Hi team,\n\nHere are the details for the ${serviceLabel} coming up${dateLabel}.\n\nOrder of Sections:\n${itemLines}\n\nTeam:\nService Leader: ${leader}\n\nPlease reply if you cannot make it.\n\nThanks!\nAustin`;
  return {people,subject,body,svc};
}
function openEmailCompose(){
  const prefs=loadEmailFromPrefs();
  const {people,subject,body}=buildServiceEmailContent();
  const withEmail=people.filter(p=>(p.email||'').trim());
  if(!withEmail.length){
    toast('Add at least one team email address first.');
    return;
  }
  $('#email-from').value=prefs.fromEmail||'';
  $('#email-from-name').value=prefs.fromName||'';
  $('#email-subject').value=subject;
  $('#email-body').value=body;
  $('#email-recipients').innerHTML=`<table class="email-team-table"><thead><tr><th></th><th>Role</th><th>Name</th><th>Email</th></tr></thead><tbody>${
    withEmail.map((p,i)=>`<tr>
      <td style="text-align:center"><input type="checkbox" data-email-rcpt="${i}" checked></td>
      <td>${e(p.role)}</td>
      <td>${e(p.name||'—')}</td>
      <td>${e(p.email.trim())}</td>
    </tr>`).join('')
  }</tbody></table>`;
  $('#email-recipients').dataset.emails=JSON.stringify(withEmail.map(p=>p.email.trim()));
  $('#email-status').textContent='';
  $('#email-modal').hidden=false;
  $('#email-from').focus();
}
function closeEmailCompose(){$('#email-modal').hidden=true}
function sendEmailFromCompose(){
  const from=($('#email-from').value||'').trim();
  const fromName=($('#email-from-name').value||'').trim();
  const subject=($('#email-subject').value||'').trim();
  const message=$('#email-body').value||'';
  const client=$('#email-client')?.value||'gmail';
  const all=JSON.parse($('#email-recipients').dataset.emails||'[]');
  const checked=[...document.querySelectorAll('#email-recipients input[type=checkbox]:checked')].map(cb=>all[+cb.dataset.emailRcpt]).filter(Boolean);
  if(!checked.length){toast('Select at least one recipient.');return}
  saveEmailFromPrefs(from,fromName);
  const to=checked.join(',');
  const su=encodeURIComponent(subject);
  const body=encodeURIComponent(message);
  let url='';
  if(client==='gmail'){
    // Opens Gmail compose in the browser (does not depend on OS mail-handler settings)
    url=`https://mail.google.com/mail/?view=cm&fs=1&tf=1&cc=${encodeURIComponent(to)}&su=${su}&body=${body}`;
  }else if(client==='outlook'){
    url=`https://outlook.live.com/mail/0/deeplink/compose?cc=${encodeURIComponent(to)}&subject=${su}&body=${body}`;
  }else{
    url=`mailto:?cc=${checked.map(encodeURIComponent).join(',')}&subject=${su}&body=${body}`;
  }
  void openExternalUrl(url).then(()=>{
    toast(client==='gmail'?'Opening Gmail…':client==='outlook'?'Opening Outlook…':'Opening your mail app…');
    closeEmailCompose();
  }).catch(err=>{
    console.error(err);
    toast(err?.message||'Could not open the email link.');
  });
}

/** Open http(s)/mailto links outside the WebView (Tauri opener) or in a new browser tab. */
async function openExternalUrl(url){
  const inTauri=typeof isTauri==='function'?await isTauri():!!(window.__TAURI_INTERNALS__||window.__TAURI__);
  if(inTauri){
    // This project already registers tauri_plugin_opener in Rust.
    try{
      const opener=await import(/* @vite-ignore */ '@tauri-apps/plugin-opener');
      if(typeof opener.openUrl==='function'){await opener.openUrl(url);return}
      if(typeof opener.open==='function'){await opener.open(url);return}
    }catch(err){console.warn('plugin-opener import failed',err)}
    try{
      // Direct plugin command (v2)
      await invoke('plugin:opener|open_url',{url});
      return;
    }catch(err){console.warn('plugin:opener|open_url failed',err)}
    throw new Error('Could not open the link in Tauri. Install @tauri-apps/plugin-opener, enable opener permissions, and match the Rust crate version.');
  }

  if(String(url).startsWith('mailto:')){
    window.location.href=url;
    return;
  }
  const w=window.open(url,'_blank','noopener,noreferrer');
  if(!w) window.location.href=url;
}

function copyEmailCompose(){
  const subject=$('#email-subject').value||'';
  const message=$('#email-body').value||'';
  const all=JSON.parse($('#email-recipients').dataset.emails||'[]');
  const checked=[...document.querySelectorAll('#email-recipients input[type=checkbox]:checked')].map(cb=>all[+cb.dataset.emailRcpt]).filter(Boolean);
  const text=`To: ${checked.join(', ')}\nSubject: ${subject}\n\n${message}`;
  navigator.clipboard.writeText(text).then(()=>toast('Message copied to clipboard.')).catch(()=>toast('Could not copy.'));
}
async function renameServiceItem(index){
  if(index<0||index>=S.length)return;
  const cur=S[index];
  const name=await appPrompt('Rename service item','Enter a new name for this service item.',cur.t||'');
  if(name==null)return;
  const clean=name.trim();
  cur.t=clean;
  rAll();
  setTimeout(scrollCurrentSlideIntoView,0);
  toast('Renamed to “'+clean+'?.');
}
function sendServiceTeamEmail(){openEmailCompose()}




/* ===== Slide import: PowerPoint → PDF (browser) → images, or PDF directly ===== */
function ensurePdfJs(){
  const lib=window.pdfjsLib||window['pdfjs-dist/build/pdf']||null;
  if(!lib)throw new Error('PDF.js failed to load. Check your network connection.');
  if(!lib.GlobalWorkerOptions.workerSrc){
    lib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  }
  return lib;
}

/** Rasterize every PDF page to PNG blobs via pdf.js. */
async function pdfBlobToSlideImages(pdfBlob,{scale=2,onProgress}={}){
  const pdfjs=ensurePdfJs();
  const data=new Uint8Array(await pdfBlob.arrayBuffer());
  const pdf=await pdfjs.getDocument({data}).promise;
  const images=[];
  for(let i=1;i<=pdf.numPages;i++){
    if(onProgress)onProgress(i,pdf.numPages);
    const page=await pdf.getPage(i);
    const viewport=page.getViewport({scale});
    const canvas=document.createElement('canvas');
    canvas.width=Math.floor(viewport.width);
    canvas.height=Math.floor(viewport.height);
    const ctx=canvas.getContext('2d',{alpha:false});
    if(!ctx)throw new Error('Could not create canvas for PDF rendering.');
    ctx.fillStyle='#ffffff';
    ctx.fillRect(0,0,canvas.width,canvas.height);
    await page.render({canvasContext:ctx,viewport}).promise;
    const blob=await new Promise((resolve,reject)=>canvas.toBlob(b=>b?resolve(b):reject(new Error('Canvas toBlob failed')),'image/png'));
    images.push(blob);
  }
  return images;
}

/** Turn an SVG string (or Blob) into a PNG Blob at the given scale. */
async function svgToPngBlob(svgInput,scale=2){
  let svgText;
  if(svgInput instanceof Blob)svgText=await svgInput.text();
  else if(typeof svgInput==='string')svgText=svgInput;
  else if(svgInput&&typeof svgInput==='object'&&typeof svgInput.svg==='string')svgText=svgInput.svg;
  else throw new Error('Unexpected SVG output from PowerPoint renderer.');

  // Ensure xmlns so the browser can paint it as an image
  if(!/xmlns=/.test(svgText)){
    svgText=svgText.replace(/<svg\b/,'<svg xmlns="http://www.w3.org/2000/svg"');
  }

  // Read intrinsic size
  let w=1920,h=1080;
  const wm=svgText.match(/\bwidth=["']([\d.]+)/);
  const hm=svgText.match(/\bheight=["']([\d.]+)/);
  const vb=svgText.match(/\bviewBox=["']\s*([\d.-]+)\s+([\d.-]+)\s+([\d.]+)\s+([\d.]+)/);
  if(wm)w=parseFloat(wm[1])||w;
  if(hm)h=parseFloat(hm[1])||h;
  if(vb){w=parseFloat(vb[3])||w;h=parseFloat(vb[4])||h;}

  const blob=new Blob([svgText],{type:'image/svg+xml;charset=utf-8'});
  const url=URL.createObjectURL(blob);
  try{
    const img=await new Promise((resolve,reject)=>{
      const i=new Image();
      i.onload=()=>resolve(i);
      i.onerror=()=>reject(new Error('Could not rasterize a PowerPoint slide SVG.'));
      i.src=url;
    });
    const canvas=document.createElement('canvas');
    canvas.width=Math.max(1,Math.round((img.naturalWidth||w)*scale));
    canvas.height=Math.max(1,Math.round((img.naturalHeight||h)*scale));
    const ctx=canvas.getContext('2d',{alpha:false});
    ctx.fillStyle='#ffffff';
    ctx.fillRect(0,0,canvas.width,canvas.height);
    ctx.drawImage(img,0,0,canvas.width,canvas.height);
    return await new Promise((resolve,reject)=>canvas.toBlob(b=>b?resolve(b):reject(new Error('PNG encode failed')),'image/png'));
  }finally{
    URL.revokeObjectURL(url);
  }
}

/** Load @web-ppt/core and return its exports. */
async function loadWebPpt(){
  if(typeof window.loadPowerPointRenderer==='function'){
    try{return await window.loadPowerPointRenderer();}catch(err){console.warn('loadPowerPointRenderer failed',err)}
  }
  return import('https://cdn.jsdelivr.net/npm/@web-ppt/core@0.4.5/+esm');
}

/**
 * Render one slide from a parsed web-ppt presentation to SVG text/Blob.
 * Tries several API shapes because package exports have varied across versions.
 */
async function renderWebPptSlide(mod,pres,slide,index){
  const candidates=[
    'slideToSvgFile','renderSlideToSvg','slideToSvg','renderSlideSvg','toSvg'
  ];
  for(const name of candidates){
    if(typeof mod[name]==='function'){
      try{
        const out=await mod[name](pres,slide);
        if(out)return out;
      }catch(err){
        try{
          const out=await mod[name](pres,index);
          if(out)return out;
        }catch{}
        console.warn('web-ppt',name,'failed',err);
      }
    }
  }
  // Some builds hang helpers off default export
  const d=mod.default||{};
  for(const name of candidates){
    if(typeof d[name]==='function'){
      try{
        const out=await d[name](pres,slide);
        if(out)return out;
      }catch{}
    }
  }
  throw new Error('This browser build of the PowerPoint engine has no SVG export function.');
}

/**
 * Convert a PowerPoint file to a PDF Blob entirely in the browser:
 * PPTX → SVG slides (@web-ppt) → PNG → multi-page PDF (jsPDF).
 */
async function powerpointToPdfBlob(file,{onProgress}={}){
  const mod=await loadWebPpt();
  const parse=mod.parse||mod.default?.parse;
  if(typeof parse!=='function')throw new Error('PowerPoint parser is not available.');

  if(onProgress)onProgress(0,0,'Parsing PowerPoint…');
  const pres=await parse(file);
  const slides=pres?.slides||pres?.Slide||[];
  if(!Array.isArray(slides)||!slides.length)throw new Error('No slides were found in this PowerPoint file.');

  const pngs=[];
  for(let i=0;i<slides.length;i++){
    if(onProgress)onProgress(i+1,slides.length,`Rendering slide ${i+1} of ${slides.length}…`);
    // Prefer native PNG if the library offers it
    let pngBlob=null;
    if(typeof mod.slideToPng==='function'){
      try{
        const out=await mod.slideToPng(pres,slides[i],2);
        if(out instanceof Blob)pngBlob=out;
        else if(out instanceof ArrayBuffer)pngBlob=new Blob([out],{type:'image/png'});
        else if(out?.buffer)pngBlob=new Blob([out],{type:'image/png'});
      }catch(err){console.warn('slideToPng failed, falling back to SVG',err)}
    }
    if(!pngBlob){
      const svg=await renderWebPptSlide(mod,pres,slides[i],i);
      pngBlob=await svgToPngBlob(svg,2);
    }
    pngs.push(pngBlob);
  }

  if(onProgress)onProgress(slides.length,slides.length,'Building PDF…');

  // Build a PDF from the PNG pages so the rest of the pipeline is PDF-only.
  let jsPdfMod;
  try{
    jsPdfMod=await import('https://cdn.jsdelivr.net/npm/jspdf@2.5.2/+esm');
  }catch(err){
    // If jsPDF cannot load, skip the PDF step and return images via a synthetic path:
    // caller will detect non-PDF and use images directly — but we still try.
    throw new Error('Could not load the PDF builder (jsPDF). Check your network connection.');
  }
  const JsPDF=jsPdfMod.jsPDF||jsPdfMod.default?.jsPDF||jsPdfMod.default;
  if(!JsPDF)throw new Error('jsPDF did not export a constructor.');

  // Probe first image size for page orientation
  const firstUrl=URL.createObjectURL(pngs[0]);
  let pageW=1920,pageH=1080;
  try{
    const img=await new Promise((resolve,reject)=>{
      const i=new Image();
      i.onload=()=>resolve(i);
      i.onerror=()=>reject(new Error('Could not read rendered slide size.'));
      i.src=firstUrl;
    });
    pageW=img.naturalWidth||pageW;
    pageH=img.naturalHeight||pageH;
  }finally{URL.revokeObjectURL(firstUrl)}

  // jsPDF uses mm; convert px→mm assuming 96dpi
  const pxToMm=px=>px*25.4/96;
  const wMm=pxToMm(pageW);
  const hMm=pxToMm(pageH);
  const orient=wMm>=hMm?'landscape':'portrait';
  const doc=new JsPDF({orientation:orient,unit:'mm',format:[wMm,hMm],compress:true});

  for(let i=0;i<pngs.length;i++){
    if(i>0)doc.addPage([wMm,hMm],orient);
    const dataUrl=await new Promise((resolve,reject)=>{
      const reader=new FileReader();
      reader.onload=()=>resolve(reader.result);
      reader.onerror=()=>reject(new Error('Could not read slide image.'));
      reader.readAsDataURL(pngs[i]);
    });
    doc.addImage(dataUrl,'PNG',0,0,wMm,hMm);
  }

  const pdfArray=doc.output('arraybuffer');
  return new Blob([pdfArray],{type:'application/pdf'});
}

async function importSlidesFromImages(fileName,imageBlobs,{kind='PowerPoint',source='powerpoint',addToService=true}={}){
  const parts=[];
  const presentationId='ppt-'+Date.now().toString(36)+'-'+Math.random().toString(36).slice(2,7);
  const baseName=fileName.replace(/\.(pptx?|pdf)$/i,'');
  for(let i=0;i<imageBlobs.length;i++){
    const blob=imageBlobs[i];
    const id='m'+Date.now().toString(36)+Math.random().toString(36).slice(2,8);
    await mPut(id,blob);
    const objectUrl=URL.createObjectURL(blob);
    mediaURLs[id]=objectUrl;
    let src=objectUrl;
    if(blob.type==='image/svg+xml'){
      try{src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(await blob.text())}catch{}
    }
    mediaMeta.push({
      id,
      name:`${fileName} — Slide ${i+1}`,
      kind:'image',
      size:blob.size,
      createdAt:Date.now(),
      source,
      presentationId,
      presentationName:fileName,
      presentationKind:kind,
      slideNumber:i+1
    });
    parts.push({id:uid(),n:`Slide ${i+1}`,tx:'',im:[{full:true,src,mediaId:id}]});
  }
  saveMediaMeta();
  const item=mk(baseName,kind,parts,parts.map(p=>p.id),0,5.2);
  item.sourceFile=fileName;item.presentationId=presentationId;item.importedAt=Date.now();
  if(addToService){
    const at=Math.max(0,Math.min(pendingSectionInsert!=null?pendingSectionInsert:sel+1,S.length));
    pendingSectionInsert=null;
    S.splice(at,0,item);sel=at;si=0;A=item;aj=0;tab='Structure';
    rAll();
    setTimeout(()=>scrollNewItemToTop(sel),0);
  }
  return {count:parts.length,presentationId};
}



/* =====================================================================
   Native PowerPoint (fresh implementation)
   Plan:
   1. On import, always store the original .pptx in IndexedDB (originalFileId).
   2. User picks "Microsoft PowerPoint" on Slides Settings (viewMode=powerpoint).
   3. During Present, the first time we land on that service item, write the
      file to app data and launch: POWERPNT.EXE /S "path" (non-blocking).
   4. Never re-open while advancing slides of the same item (opened set).
   5. Never block Next/Esc; Church Presenter keeps running underneath.
   ===================================================================== */
// pptOpenedThisPresent declared above; tracks keys opened during current present session

async function ensurePptFileOnDisk(it){
  if(!it)return '';
  if(it.sourcePath)return it.sourcePath;
  if(!it.originalFileId)return '';
  const blob=await mGetFile(it.originalFileId);
  if(!blob)return '';
  const buf=new Uint8Array(await blob.arrayBuffer());
  let binary='';
  const chunk=0x8000;
  for(let i=0;i<buf.length;i+=chunk)binary+=String.fromCharCode.apply(null,buf.subarray(i,i+chunk));
  const ext=((it.sourceFile||'pptx').split('.').pop()||'pptx').toLowerCase().replace(/[^a-z0-9]/g,'')||'pptx';
  const safeId=String(it.presentationId||it.originalFileId).replace(/[^a-zA-Z0-9_-]/g,'_').slice(0,80);
  const filePath=await invoke('store_presentation_file',{
    id:safeId,
    fileName:it.sourceFile||('presentation.'+ext),
    extension:ext,
    dataBase64:btoa(binary)
  });
  if(filePath){it.sourcePath=filePath;saveSoon()}
  return filePath||'';
}

async function launchPowerpointSlideshow(filePath){
  if(!filePath)throw new Error('No path');
  let firstErr=null;
  try{
    await invoke('launch_ppt_slideshow',{path:filePath});
    return 'launch_ppt_slideshow';
  }catch(e1){firstErr=e1;console.warn('[PPT] launch_ppt_slideshow',e1)}
  try{
    await invoke('open_presentation_slideshow',{path:filePath});
    return 'open_presentation_slideshow';
  }catch(e2){console.warn('[PPT] open_presentation_slideshow',e2)}
  // Deliberately no edit-mode fallback: it would open PowerPoint without starting the slideshow.
  throw new Error(typeof firstErr==='string'?firstErr:((firstErr&&firstErr.message)||'Could not launch PowerPoint'));
}

function pptKeyFor(it){return String(it.presentationId||it.originalFileId||it.sourceFile||it.t||'')}
function isPptHandoffItem(it){return !!it&&it.k==='PowerPoint'&&(it.viewMode||'church')==='powerpoint'}
function inTauriRuntime(){return !!(window.__TAURI_INTERNALS__||window.__TAURI__)}
function stopPptSlideshow(){
  try{if(inTauriRuntime())invoke('stop_ppt_slideshow').catch(()=>{})}catch{}
}
/** Called when presentation navigation leaves a PowerPoint item. */
function pptLeaveItem(it){if(isPptHandoffItem(it)){pptShowActive=false;restoreAudienceAfterPpt()}
  if(!isPptHandoffItem(it))return;
  pptOpenedThisPresent.delete(pptKeyFor(it));
  stopPptSlideshow();
  try{_lastAudienceHtml='';pushAudienceStage()}catch{}
}
/** PowerPoint finished: jump to the first slide of the NEXT service item. */
function pptAdvanceAfterShow(it){pptShowActive=false;restoreAudienceAfterPpt();
  if(!live||A!==it)return; // user already moved on, or presentation ended
  const i=S.indexOf(it);
  if(i<0)return;
  pptOpenedThisPresent.delete(pptKeyFor(it));
  if(i+1>=S.length){endPresent();return}
  A=S[i+1];sel=i+1;si=0;aj=0;clr=null;
  try{_lastAudienceHtml=''}catch{}
  try{rStage()}catch(err){console.error(err)}
  try{updateStripSelection()}catch{}
  try{
    const outline=$('#outline');
    if(outline)outline.querySelectorAll('.oi').forEach((el,k)=>el.classList.toggle('on',k===sel));
    const pvOutline=$('#pv-outline');
    if(pvOutline)pvOutline.querySelectorAll('.pv-oi').forEach((el,k)=>el.classList.toggle('on',k===sel));
  }catch{}
  refocusPresenterWindow().catch(()=>{});
}
async function waitForPptEnd(it){
  try{
    const res=await invoke('wait_ppt_slideshow_end',{startTimeoutSecs:40});
    if(res==='ended')pptAdvanceAfterShow(it);
    else{console.warn('[PPT] slideshow never started:',res);pptShowActive=false;restoreAudienceAfterPpt()}
  }catch(err){console.warn('[PPT] cannot track slideshow end (rebuild the Rust side?)',err)}
}

function openNativePowerpointIfNeeded(it){
  if(!live||!pptLaunchArmed||!isPptHandoffItem(it))return;
  const key=pptKeyFor(it);
  if(!key||pptOpenedThisPresent.has(key))return;
  pptOpenedThisPresent.add(key); // set BEFORE async work to prevent loops
  try{_lastAudienceHtml='';pushAudienceStage()}catch{}
  (async()=>{
    try{
      if(!inTauriRuntime())throw new Error('only works in the desktop app');
      const path=await ensurePptFileOnDisk(it);
      if(!path)throw new Error('original .pptx not stored - re-import it');
      await releaseAudienceForPpt();
      pptShowActive=true;
      await launchPowerpointSlideshow(path);
      try{await invoke('focus_ppt_slideshow')}catch(fe){console.warn('[PPT] focus',fe)}
      waitForPptEnd(it); // not awaited: runs in the background
    }catch(err){
      console.error('[PPT] open failed',err);
      pptOpenedThisPresent.delete(key);
      pptShowActive=false;restoreAudienceAfterPpt();
      toast('Could not open PowerPoint: '+((err&&err.message)||err));
      try{_lastAudienceHtml='';pushAudienceStage()}catch{}
    }
  })();
}

async function testOpenPowerpoint(btn){
  const it=S[sel];if(!it)return;
  const prev=btn.textContent;btn.disabled=true;btn.textContent='Opening…';
  const done=m=>{btn.textContent=m;setTimeout(()=>{btn.textContent=prev;btn.disabled=false},3500)};
  try{
    if(!(window.__TAURI_INTERNALS__||window.__TAURI__))return done('Desktop app only');
    const path=await ensurePptFileOnDisk(it);
    if(!path)return done('No file - re-import');
    done('OK: '+await launchPowerpointSlideshow(path));
  }catch(err){console.error('[PPT test]',err);done('Failed: '+((err&&err.message)||err))}
}

/** No-op preload: only ensure files are written when Present starts (optional warm). */
async function preloadPowerpointsForService(){
  try{
    if(!(window.__TAURI_INTERNALS__||window.__TAURI__))return;
    for(const it of S){
      if(!(it.k==='PowerPoint'||it.k==='PDF'))continue;
      if((it.viewMode||'church')!=='powerpoint')continue;
      try{await ensurePptFileOnDisk(it)}catch(err){console.warn('[PPT] preload store',err)}
    }
  }catch(err){console.warn(err)}
}

async function tryPowerpointHandoff(it){openNativePowerpointIfNeeded(it);return true}
async function finishPowerpointHandoff(){}

async function importPowerPoint(file,{addToService=true}={}){
  const status=$('#ppt-import-status');
  if(!file)return;
  importInProgress=false; // never block a fresh file pick
  if(!file)return;
  const name=file.name||'';
  const isPdf=/\.pdf$/i.test(name)||file.type==='application/pdf';
  const isPpt=/\.pptx?$/i.test(name);
  if(!isPdf&&!isPpt)return;
  showImportModal();
  let importedPresentationId='';
  try{
    let pdfBlob;
    if(isPdf){
      throwIfImportCancelled();
      setImportProgress(5,'Opening PDF…');
      pdfBlob=file;
    }else{
      throwIfImportCancelled();
      setImportProgress(5,'Converting PowerPoint…');
      pdfBlob=await powerpointToPdfBlob(file,{
        onProgress:(i,n,msg)=>{
          throwIfImportCancelled();
          const pct=n?5+Math.round((i/n)*45):10;
          setImportProgress(pct,msg||(n?`Converting slide ${i} of ${n}…`:'Converting…'));
        }
      });
      throwIfImportCancelled();
      setImportProgress(50,'PDF ready. Rendering pages…');
    }
    const images=await pdfBlobToSlideImages(pdfBlob,{
      scale:2,
      onProgress:(i,n)=>{
        throwIfImportCancelled();
        const pct=50+Math.round((i/Math.max(1,n))*40);
        setImportProgress(pct,`Rendering page ${i} of ${n}…`);
      }
    });
    throwIfImportCancelled();
    if(!images.length)throw new Error('No pages were found.');
    setImportProgress(92,'Saving slides…');
    const imported=await importSlidesFromImages(name,images,{
      kind:isPdf?'PDF':'PowerPoint',
      source:'powerpoint',
      addToService
    });
    importedPresentationId=imported.presentationId;
    throwIfImportCancelled();
    try{
      const fileId='orig-'+Date.now().toString(36)+'-'+Math.random().toString(36).slice(2,8);
      await mPutFile(fileId,file);
      const item=addToService?S[sel]:null;
      const presentationId=imported.presentationId;
      if(item&&(item.k==='PowerPoint'||item.k==='PDF')){
        item.originalFileId=fileId;
        mediaMeta.filter(m=>m.presentationId===item.presentationId).forEach(m=>{m.originalFileId=fileId});
        item.sourceFile=name;
        item.sourcePath='';
        item.viewMode=item.viewMode||(item.k==='PowerPoint'?'powerpoint':'church');
        if(services[activeService])services[activeService].slides=sanitizeSlidesForSave(S);
        saveState();
        saveMediaMeta();
      }else if(presentationId){
        mediaMeta.filter(m=>m.presentationId===presentationId).forEach(m=>{m.originalFileId=fileId});
        saveMediaMeta();
      }
    }catch(storeErr){console.warn('Could not store original presentation file',storeErr)}
    try{
      if(services[activeService])services[activeService].slides=sanitizeSlidesForSave(S);
      saveState();
    }catch{}
    setImportProgress(100,'Done');
    if(status)status.textContent=`Imported ${imported.count} slide${imported.count===1?'':'s'} into the media library${addToService?' and service':''}.`;
  }catch(err){
    if(err&&err.cancelled){
      // Roll back any partial item added at end of importSlidesFromImages — if cancelled mid-way, nothing added yet
      // If cancelled after importSlidesFromImages, remove last item if it matches this import
      try{
        if(!addToService&&importedPresentationId){
          const removed=mediaMeta.filter(m=>m.presentationId===importedPresentationId);
          for(const media of removed){if(mediaURLs[media.id])URL.revokeObjectURL(mediaURLs[media.id]);delete mediaURLs[media.id];await mDelete(media.id)}
          mediaMeta=mediaMeta.filter(m=>m.presentationId!==importedPresentationId);saveMediaMeta();
        }
        const last=S[sel];
        if(addToService&&last&&(last.k==='PowerPoint'||last.k==='PDF')&&last.sourceFile===name){
          S.splice(sel,1);
          if(!S.length){A=null;sel=0;aj=0}
          else{sel=Math.min(sel,S.length-1);A=S[sel];aj=0}
          if(services[activeService])services[activeService].slides=sanitizeSlidesForSave(S);
          saveState();
          rAll();
        }
      }catch{}
      if(status)status.textContent='Import cancelled.';
    }else{
      console.error(err);
      if(status)status.textContent='Import failed.';
      toast('Import failed: '+((err&&err.message)||err));
    }
  }finally{
    pendingSectionInsert=null;
    hideImportModal();
    if(!addToService){const root=$('#media-content');if(root){root.dataset.mediaTab='powerpoints';renderMediaLibrary()}}
  }
}

document.addEventListener('change',ev=>{
  if(ev.target.id==='fi'){
    const f=ev.target.files[0];ev.target.value='';
    if(!f)return;
    if(/\.cpres$/i.test(f.name||''))importCpresMergeIntoCurrent(f);
    else importPowerPoint(f);
  }
});




/* ===== Per-slide video ===== */
async function addVideoSlideItem(file){
  if(!file||!file.type.startsWith('video/'))return toast('Choose a video file.');
  const id='m'+Date.now().toString(36)+Math.random().toString(36).slice(2,8);
  await mPut(id,file);mediaURLs[id]=URL.createObjectURL(file);
  mediaMeta.push({id,name:file.name,kind:'video',size:file.size,createdAt:Date.now(),source:'slide-video'});saveMediaMeta();
  const part={id:uid(),n:file.name,tx:'',videoMedia:{id,kind:'video',name:file.name,settings:{autoplay:true,loop:false,muted:false,volume:1}}};
  const item=mk(file.name.replace(/\.[^.]+$/,''),'Video',[part],[part.id],0,5.2);
  item.sourceFile=file.name;
  S.splice(sel+1,0,item);sel++;si=0;A=item;aj=0;tab='Structure';rAll();setTimeout(()=>scrollNewItemToTop(sel),0);toast('Added video slide.');
}

async function addVideoToCurrentSlide(file){
  if(!file||!file.type.startsWith('video/'))return toast('Choose a video file.');
  const part=sls(A)[aj];if(!part)return;
  const id='m'+Date.now().toString(36)+Math.random().toString(36).slice(2,8);
  await mPut(id,file);mediaURLs[id]=URL.createObjectURL(file);
  mediaMeta.push({id,name:file.name,kind:'video',size:file.size,createdAt:Date.now(),source:'slide-video'});saveMediaMeta();
  pushUndo();addVideoElToCurrentSlide(id,file.name);
}


function formatLastEdited(ts){
  if(ts==null||ts==='')return '';
  let t=typeof ts==='number'?ts:Date.parse(ts);
  if(!t||isNaN(t))return '';
  const sec=Math.max(0,Math.round((Date.now()-t)/1000));
  if(sec<45)return 'Edited just now';
  if(sec<90)return 'Edited 1 minute ago';
  if(sec<3600)return 'Edited '+Math.round(sec/60)+' minutes ago';
  if(sec<5400)return 'Edited 1 hour ago';
  if(sec<86400)return 'Edited '+Math.round(sec/3600)+' hours ago';
  if(sec<172800)return 'Edited yesterday';
  try{
    const d=new Date(t);
    return 'Edited '+d.toLocaleDateString(undefined,{day:'numeric',month:'short',year:'numeric'});
  }catch{return ''}
}
const edHash=new WeakMap();
function touchServiceEdited(){
  if(!services.length)return;
  if(activeService<0||activeService>=services.length)return;
  const svc=services[activeService];if(!svc)return;
  let h;
  try{h=hashService(activeService)}catch{svc.lastEdited=Date.now();return}
  const prev=edHash.get(svc);
  edHash.set(svc,h);
  if(prev===undefined)return;
  if(prev!==h){
    svc.lastEdited=Date.now();
  }
}
function formatServiceDate(iso){
  if(!iso)return 'No date';{const rl=relDayLabelISO(iso);if(rl)return rl}
  const m=String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m)return m[3]+'/'+m[2]+'/'+m[1].slice(2);
  // already formatted or other
  try{
    const d=new Date(iso);
    if(!isNaN(d)){
      const dd=String(d.getDate()).padStart(2,'0');
      const mm=String(d.getMonth()+1).padStart(2,'0');
      const yy=String(d.getFullYear()).slice(2);
      return dd+'/'+mm+'/'+yy;
    }
  }catch{}
  return String(iso);
}
function formatServiceTitle(s){
  if(!s)return 'Service';
  const tag=(s.tag||'').trim();
  const name=(s.name||'').trim();
  if(tag&&name)return tag+' - '+name;
  if(tag)return tag;
  return name||'Service';
}

/* ===== Services home page ===== */
function rHome(){
  $('#service-tag-filter').value=serviceTagFilter;
  $('#service-sort').value=serviceSort;
  let list=services.map((s,i)=>({s,i}));
  if(serviceTagFilter) list=list.filter(x=>x.s.tag===serviceTagFilter);
  if(serviceSort==='tag') list.sort((a,b)=>(a.s.tag||'').localeCompare(b.s.tag||''));
  else if(serviceSort==='date') list.sort((a,b)=>(b.s.date||'').localeCompare(a.s.date||''));
  else if(serviceSort==='edited') list.sort((a,b)=>(b.s.lastEdited||0)-(a.s.lastEdited||0));
  if(!list.length){
    $('#service-list').innerHTML=`<div style="padding:28px 12px;text-align:center">
      <p class="note" style="margin:0 0 16px">No services yet.</p>
      <button type="button" id="make-new-service" class="pri" style="padding:12px 22px;font-size:calc(15px * var(--ui-scale))">Make a new service</button>
    </div>`;
  }else{
    $('#service-list').innerHTML=list.map(({s,i})=>`
    <div class="service-card">
      <div class="info">
        ${svcTitleHtml(s,i)}
        <small>${e(formatServiceDate(s.date))}${s.lastEdited?' · '+e(formatLastEdited(s.lastEdited)):''}</small>
      </div>
      <button class="pri" data-open-service="${i}">Open</button>
      <button data-del-service="${i}" style="color:#c22">Delete</button>
    </div>`).join('');
  }
}

$('#service-tag-filter').addEventListener('change',ev=>{serviceTagFilter=ev.target.value;homeShown=5;rHome()});
$('#service-sort').addEventListener('change',ev=>{serviceSort=ev.target.value;homeShown=5;rHome()});

function openNewServiceForm(){
  $('#service-date').value=upcomingSundayISO();
  $('#service-form').classList.add('open');
  $('#service-tag').focus();
}
$('#new-service').addEventListener('click',openNewServiceForm);
document.addEventListener('click',ev=>{
  if(ev.target&&ev.target.id==='home-import-cpres'){
    const inp=$('#home-import-cpres-file')||$('#import-cpres-file');
    if(inp){inp.value='';inp.click()}
  }
});
document.addEventListener('change',ev=>{
  if(ev.target&&ev.target.id==='home-import-cpres-file'){
    const f=ev.target.files&&ev.target.files[0];
    ev.target.value='';
    if(typeof importCpresFile==='function')importCpresFile(f);
  }
});
document.addEventListener('click',ev=>{
  if(ev.target&&ev.target.id==='make-new-service')openNewServiceForm();
});
$('#cancel-service').addEventListener('click',()=>$('#service-form').classList.remove('open'));

$('#service-form').addEventListener('submit',ev=>{
  ev.preventDefault();
  const tag=($('#service-tag').value||'').trim();
  if(!tag){toast('Choose a service tag.');$('#service-tag').focus();return}
  const name=($('#service-name').value||'').trim();
  const date=$('#service-date').value;
  if(!date){toast('Choose a date.');$('#service-date').focus();return}
  const svc={
    name:name,
    date,
    tag,
    slides:[mkS('Slide',[''],0,6.5)],
    team:defaultTeam(),
    lastEdited:Date.now()
  };
  services.push(svc);
  activeService=services.length-1;
  S=cl(svc.slides);sel=0;si=0;A=S[0];aj=0;
  saveState();
  $('#service-form').classList.remove('open');
  $('#service-form').reset();
  document.body.classList.remove('home-mode');
  rAll();
});

document.addEventListener('click',async ev=>{
  const b=ev.target.closest('[data-rename-service]');
  if(!b)return;
  ev.preventDefault();ev.stopPropagation();
  const i=+b.dataset.renameService;
  const svc=services[i];if(!svc)return;
  const current=svc.name||'';
  const n=await textEntryPopup({title:'Rename service',label:'Service name',value:current,placeholder:'Enter service name'});
  if(n==null)return;
  const clean=n.trim();
  svc.name=clean;svc.lastEdited=Date.now();
  saveState();rHome();
});

document.addEventListener('click',async ev=>{
  const b=ev.target.closest('[data-del-service]');
  if(!b)return;
  const di=+b.dataset.delService;
  const svc=services[di];
  if(!svc)return;
  const title=formatServiceTitle(svc);
  if(!await appConfirm('Delete "'+title+'"?',{title:'Delete service',confirmLabel:'Delete',danger:true}))return;

  const link=getLink(svc);
  const fileId=link?link.fileId:null;

  if(svc.id){
    try{
      const meta=loadCloudMeta();
      delete meta.services[svc.id];
      saveCloudMeta(meta);
    }catch(e){}
  }

  services.splice(di,1);
  if(!services.length){
    activeService=0;
    S=[];sel=0;si=0;A=null;aj=0;
  }else{
    activeService=Math.min(activeService,services.length-1);
    if(activeService<0)activeService=0;
    S=cl(services[activeService].slides||[]);sel=0;si=0;A=S[0]||null;aj=0;
  }

  if(fileId&&cloudUi&&Array.isArray(cloudUi.files)){
    cloudUi.files=cloudUi.files.filter(f=>f.id!==fileId);
  }

  saveState();
  rHome();
  toast('Service deleted.');

  if(fileId&&cloudReady()){
    cloudInvoke('gdrive_delete_file',{fileId}).then(res=>{
      console.log('[cloud] background delete ok',res);
    }).catch(err=>{
      console.warn('[cloud] background delete error',err);
    });
  }
});

/* Rename current service by clicking its title */
$('#service-title-edit').addEventListener('click',async()=>{
  const svc=services[activeService];if(!svc)return;
  const n=await textEntryPopup({title:'Rename service',label:'Service name',value:svc.name||'',placeholder:'Enter service name'});
  if(n!=null){svc.name=n.trim();svc.lastEdited=Date.now();saveState();rAll()}
});
window.addEventListener('church-presenter-manual-sync',()=>cloudSyncAll({manual:true}));

/* Settings */
$('#display-setting').addEventListener('change',ev=>{selectedDisplay=ev.target.value;saveState()});
$('#close-settings').addEventListener('click',closeSettingsAndRefresh);
$('#settings').addEventListener('click',ev=>{if(ev.target===$('#settings'))closeSettingsAndRefresh()});

if(!IS_AUDIENCE) rAll();


/* ===== team header + notes + ui scale ===== */
$('#team-close')?.addEventListener('click',closeTeamModal);
$('#team-modal')?.addEventListener('click',ev=>{if(ev.target.id==='team-modal')closeTeamModal()});
document.addEventListener('input',ev=>{
  if(ev.target?.id==='speaker-notes'){
    const part=sls(A)[aj];
    if(part){part.notes=ev.target.value;saveSoon()}
    return;
  }
});
$('#ui-scale-setting')?.addEventListener('change',ev=>{
  applyUiScale(ev.target.value);
  saveState();
});
// Apply volume/mute on stage videos after render
const _origAttach=typeof attachVideos==='function'?attachVideos:null;

/* ===== email compose UI ===== */
$('#import-cancel-btn')?.addEventListener('click',()=>{importCancelled=true;setImportProgress(0,'Cancelling…')});
$('#email-close')?.addEventListener('click',closeEmailCompose);
$('#email-send')?.addEventListener('click',sendEmailFromCompose);
$('#email-copy')?.addEventListener('click',copyEmailCompose);
$('#email-modal')?.addEventListener('click',ev=>{if(ev.target.id==='email-modal')closeEmailCompose()});

/* ===== connected displays ===== */
/** Turn raw OS monitor names (e.g. \\\\.\\DISPLAY1) into readable labels. */
function friendlyMonitorName(raw,index,isPrimary,isCurrent,width,height){
  let name=(raw||'').toString().trim();
  // Windows device paths: \\.\DISPLAY1 → Display 1
  const win=name.match(/DISPLAY\s*(\d+)/i);
  if(win)name=`Display ${win[1]}`;
  else if(!name||name==='null'||name==='undefined')name=`Display ${index+1}`;

  const parts=[name];
  if(isPrimary)parts.push('Primary');
  if(isCurrent)parts.push('this window');
  if(width&&height)parts.push(`${width}×${height}`);
  return parts[0]+(parts.length>1?' · '+parts.slice(1).join(' · '):'');
}

async function refreshDisplayOptions(){
  const sel=$('#display-setting');if(!sel)return;
  const previous=sel.value||selectedDisplay||'';
  const opts=[]; // {value,label,isCurrent}

  try{
    const monitors=await availableMonitors();
    let currentMon=null;
    try{currentMon=await currentMonitor()}catch{}

    // Primary ≈ top-left-most monitor
    let primaryIdx=0;
    let best=Infinity;
    (monitors||[]).forEach((m,i)=>{
      const x=m.position?.x??0,y=m.position?.y??0;
      const score=y*100000+x;
      if(score<best){best=score;primaryIdx=i}
    });

    let externalCount=0;
    (monitors||[]).forEach((m,i)=>{
      const isCurrent=!!(currentMon&&(
        (m.name&&currentMon.name&&m.name===currentMon.name)||
        (m.position?.x===currentMon.position?.x&&m.position?.y===currentMon.position?.y)
      ));
      const isPrimary=i===primaryIdx;
      if(!isPrimary)externalCount++;
      const w=m.size?.width;
      const h=m.size?.height;
      let base=(m.name||'').toString().trim();
      const win=base.match(/DISPLAY\s*(\d+)/i);
      if(win){
        base=isPrimary?'Primary display':(externalCount>1?`External display ${externalCount}`:'External display');
      }else if(!base||base==='null'||base==='undefined'){
        base=isPrimary?'Primary display':(externalCount>1?`External display ${externalCount}`:'External display');
      }
      // If OS gave a real product name, keep it; still mark primary/current
      opts.push({
        value:String(m.name||i),
        label:friendlyMonitorName(base,i,isPrimary&&!/primary/i.test(base),isCurrent,w,h),
        isCurrent
      });
    });
  }catch(err){console.warn('Tauri availableMonitors failed',err)}

  if(!opts.length){
    try{
      if(window.getScreenDetails){
        screenDetailsCache=await window.getScreenDetails();
        const screens=screenDetailsCache.screens||[];
        screens.forEach((s,i)=>{
          const isCurrent=screenDetailsCache.currentScreen===s;
          const isPrimary=!!s.isPrimary;
          const w=s.width||s.availWidth;
          const h=s.height||s.availHeight;
          const raw=s.label||s.id||`Display ${i+1}`;
          opts.push({
            value:String(s.label||i),
            label:friendlyMonitorName(raw,i,isPrimary,isCurrent,w,h),
            isCurrent:!!isCurrent
          });
        });
      }
    }catch(err){console.warn('getScreenDetails unavailable',err)}
  }

  if(!opts.length){
    opts.push({value:'0',label:'Primary display',isCurrent:true});
  }

  sel.innerHTML=opts.map(o=>`<option value="${e(o.value)}">${e(o.label)}</option>`).join('');

  // Fingerprint of available monitors — only auto-pick external when this changes
  // (app launch or a display plugged in/out), not every time Settings is opened.
  const fingerprint=opts.map(o=>o.value).slice().sort().join('|');
  const prevFingerprint=(()=>{try{return localStorage.getItem('display-fingerprint')||''}catch{return ''}})();
  const monitorsChanged=fingerprint!==prevFingerprint;
  try{localStorage.setItem('display-fingerprint',fingerprint)}catch{}

  const external=opts.find(o=>!o.isCurrent);
  const prevStillValid=previous&&opts.some(o=>o.value===previous);

  if(monitorsChanged&&external){
    // New monitor set: prefer external output
    sel.value=external.value;
  }else if(prevStillValid){
    sel.value=previous;
  }else if(external){
    sel.value=external.value;
  }else{
    sel.selectedIndex=0;
  }
  selectedDisplay=sel.value;
  // Persist only when monitors changed (auto-pick) — user picks are saved on change event
  if(monitorsChanged)saveState();
}

$('#display-setting')?.addEventListener('change',ev=>{
  selectedDisplay=ev.target.value;saveState();
});
$('#settings-btn')?.addEventListener('click',()=>{refreshDisplayOptions();updateLogoPreview()});
// initial attempt (may prompt permission on some browsers)
refreshDisplayOptions().catch(()=>{});


document.addEventListener('focusin',ev=>{
  const t=ev.target;
  if(!t||!t.matches)return;
  // Color / range inputs in the style panel must not scroll the left strip
  if(t.matches('input[type=color],input[data-k="color"],input[data-k="color-hex"],input[data-k="fs"],input[data-k="mb"]')){
    const strip=$('#strip');
    if(strip){
      const y=strip.scrollTop;
      requestAnimationFrame(()=>{strip.scrollTop=y});
      setTimeout(()=>{strip.scrollTop=y},0);
    }
  }
},true);


/* ===== Freeform canvas (non-song slides) ===== */
function closeCanvasCtx(){
  const m=$('#canvas-ctx');if(m){m.classList.remove('open');m.hidden=true}
}
function openCanvasCtx(x,y){
  const m=$('#canvas-ctx');if(!m)return;
  m.hidden=false;m.classList.add('open');
  const w=m.offsetWidth||160,h=m.offsetHeight||180;
  m.style.left=Math.min(x,window.innerWidth-w-8)+'px';
  m.style.top=Math.min(y,window.innerHeight-h-8)+'px';
}
function selectEl(id,{refresh=true}={}){
  const changed=elSel!==(id||null);
  elSel=id||null;
  if(elSel&&!isSongItem(A))tab='Structure';
  if(!refresh)return;
  if(changed){rStage();rPanel()}
  else{
    // already selected — only refresh selection chrome
    document.querySelectorAll('#stage .el').forEach(n=>{
      n.classList.toggle('selected',n.dataset.el===elSel);
    });
  }
}
function addTextBox(){
  if(isSongItem(A)||live)return;
  pushUndo();
  const part=sls(A)[aj];if(!part)return;
  const els=ensureElements(part,A);
  const f={...(A.f||{fs:5.2,ff:'sans',tc:'none',al:'center',va:'middle',b:1,sh:0,color:'#ffffff'})};
  const el={id:uid(),kind:'text',tx:'New text',x:20,y:30,w:40,h:20,rot:0,z:(els.reduce((m,e)=>Math.max(m,e.z||1),0)+1),f};
  els.push(el);elSel=el.id;tab='Structure';saveSoon();rStage();updateCurrentThumb();rPanel();
}
function deleteSelectedEl(){
  const part=sls(A)[aj];if(!part||!elSel)return;
  pushUndo();
  const els=part.elements||[];
  const i=els.findIndex(e=>e.id===elSel);
  if(i>=0){els.splice(i,1);elSel=null;saveSoon();rStage();updateCurrentThumb();rPanel()}
}
function duplicateSelectedEl(){
  const part=sls(A)[aj];if(!part||!elSel)return;
  pushUndo();
  const els=ensureElements(part,A);
  const src=els.find(e=>e.id===elSel);if(!src)return;
  const copy=cl(src);copy.id=uid();copy.x=Math.min(90,(src.x||0)+4);copy.y=Math.min(90,(src.y||0)+4);
  copy.z=(els.reduce((m,e)=>Math.max(m,e.z||1),0)+1);
  els.push(copy);elSel=copy.id;saveSoon();rStage();updateCurrentThumb();rPanel();
}
const EL_CLIP_MARK='[church-presenter-element]';
function writeElementClip(){
  const clip=canvasClip;if(!clip)return;
  const plain=clip.kind==='text'?String(clip.tx||''):'Church Presenter element';
  let handled=false;
  const prevFocus=document.activeElement;
  const ta=document.createElement('textarea');
  ta.value=plain||' ';
  ta.style.cssText='position:fixed;left:-9999px;top:0;opacity:0';
  document.body.appendChild(ta);
  ta.select();
  const h=function(ev){
    ev.preventDefault();
    ev.clipboardData.setData('text/plain',plain||' ');
    ev.clipboardData.setData('text/html','<span data-church-presenter-el="1">'+e(plain)+'</span>');
    handled=true;
  };
  document.addEventListener('copy',h,true);
  try{document.execCommand('copy')}catch{}
  document.removeEventListener('copy',h,true);
  ta.remove();
  try{if(prevFocus&&prevFocus!==document.body&&prevFocus.focus)prevFocus.focus()}catch{}
  if(!handled&&navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(EL_CLIP_MARK).catch(()=>{})}
}
function copySelectedEl(){
  const el=getSelectedEl();if(!el)return;
  canvasClip=cl(el);
  try{writeElementClip()}catch(err){console.warn('[clipboard]',err)}
  toast('Copied');
}
function pasteCanvasEl(){
  if(isSongItem(A)||live||!canvasClip)return;
  const part=sls(A)[aj];if(!part)return;
  const els=ensureElements(part,A);
  const copy=cl(canvasClip);copy.id=uid();
  copy.x=Math.min(85,(copy.x||10)+5);copy.y=Math.min(85,(copy.y||10)+5);
  copy.z=(els.reduce((m,e)=>Math.max(m,e.z||1),0)+1);
  els.push(copy);elSel=copy.id;saveSoon();rStage();updateCurrentThumb();rPanel();
}
document.addEventListener('paste',function(ev){
  if(live||!A||isSongItem(A)||!canvasClip)return;
  if(isTypingTarget(ev.target))return;
  const cd=ev.clipboardData;if(!cd)return;
  let htmlData='',txt='';
  try{htmlData=cd.getData('text/html')||'';txt=cd.getData('text/plain')||''}catch{}
  if(htmlData.indexOf('data-church-presenter-el')>=0||txt===EL_CLIP_MARK){
    ev.preventDefault();
    ev.stopImmediatePropagation();
    pasteCanvasEl();
  }
},true);
function layerSelected(cmd){
  const part=sls(A)[aj];if(!part||!elSel)return;
  const els=part.elements||[];
  const sorted=[...els].sort((a,b)=>(a.z||0)-(b.z||0));
  const idx=sorted.findIndex(e=>e.id===elSel);if(idx<0)return;
  if(cmd==='front'){sorted.forEach((e,i)=>e.z=i+1);sorted[idx].z=sorted.length+1}
  else if(cmd==='back'){sorted.forEach((e,i)=>e.z=i+2);sorted[idx].z=1}
  else if(cmd==='forward'&&idx<sorted.length-1){
    const z=sorted[idx].z;sorted[idx].z=sorted[idx+1].z;sorted[idx+1].z=z;
  }else if(cmd==='backward'&&idx>0){
    const z=sorted[idx].z;sorted[idx].z=sorted[idx-1].z;sorted[idx-1].z=z;
  }
  // normalize
  [...els].sort((a,b)=>(a.z||0)-(b.z||0)).forEach((e,i)=>e.z=i+1);
  saveSoon();rStage();
}


/** Snap threshold in % of slide size */
const SNAP_THRESH=1.15;
function clearSnapGuides(){
  document.querySelectorAll('#stage .snap-guide').forEach(n=>n.remove());
}
function showSnapGuides(slEl,xs,ys){
  clearSnapGuides();
  if(!slEl)return;
  (xs||[]).forEach(x=>{
    const g=document.createElement('div');
    g.className='snap-guide v';
    g.style.left=x+'%';
    slEl.appendChild(g);
  });
  (ys||[]).forEach(y=>{
    const g=document.createElement('div');
    g.className='snap-guide h';
    g.style.top=y+'%';
    slEl.appendChild(g);
  });
}
/** Snap a value to nearest target within threshold; returns {v, snapped}. */
function snapTo(val,targets,thresh){
  let best=val,d=thresh;
  for(const t of targets){
    const dd=Math.abs(val-t);
    if(dd<=d){d=dd;best=t}
  }
  return {v:best,snapped:best!==val||(targets.includes(val)&&d<=thresh)};
}
function collectSnapTargets(part,excludeId){
  const xs=[0,50,100],ys=[0,50,100];
  (part&&part.elements||[]).forEach(o=>{
    if(!o||o.id===excludeId)return;
    const L=+o.x||0,T=+o.y||0,W=+o.w||0,H=+o.h||0;
    xs.push(L,L+W/2,L+W);
    ys.push(T,T+H/2,T+H);
  });
  return {xs,ys};
}
function applyMoveSnap(el,part){
  const {xs,ys}=collectSnapTargets(part,el.id);
  const L=el.x,T=el.y,W=el.w,H=el.h;
  const R=L+W,B=T+H,CX=L+W/2,CY=T+H/2;
  // Prefer snapping left, center, or right edges
  const candX=[
    {edge:'L',val:L,shift:v=>L+(v-L)},
    {edge:'C',val:CX,shift:v=>L+(v-CX)},
    {edge:'R',val:R,shift:v=>L+(v-R)}
  ];
  let bestX=null,bestXd=SNAP_THRESH;
  const guideXs=[];
  for(const c of candX){
    for(const t of xs){
      const dd=Math.abs(c.val-t);
      if(dd<bestXd){bestXd=dd;bestX={t,shift:c.shift(t)}}
    }
  }
  if(bestX!=null&&bestXd<SNAP_THRESH){el.x=bestX.shift;guideXs.push(bestX.t)}
  const candY=[
    {edge:'T',val:T,shift:v=>T+(v-T)},
    {edge:'C',val:CY,shift:v=>T+(v-CY)},
    {edge:'B',val:B,shift:v=>T+(v-B)}
  ];
  let bestY=null,bestYd=SNAP_THRESH;
  const guideYs=[];
  for(const c of candY){
    for(const t of ys){
      const dd=Math.abs(c.val-t);
      if(dd<bestYd){bestYd=dd;bestY={t,shift:c.shift(t)}}
    }
  }
  if(bestY!=null&&bestYd<SNAP_THRESH){el.y=bestY.shift;guideYs.push(bestY.t)}
  el.x=Math.max(-5,Math.min(95,el.x));
  el.y=Math.max(-5,Math.min(95,el.y));
  return {guideXs,guideYs};
}
function applyResizeSnap(el,mode,part){
  const {xs,ys}=collectSnapTargets(part,el.id);
  const guideXs=[],guideYs=[];
  // Snap the moving edges
  if(mode==='se'||mode==='ne'){
    const r=el.x+el.w;
    const s=snapTo(r,xs,SNAP_THRESH);
    if(s.snapped){el.w=Math.max(4,s.v-el.x);guideXs.push(s.v)}
  }
  if(mode==='sw'||mode==='nw'){
    const L=el.x,R=el.x+el.w;
    const s=snapTo(L,xs,SNAP_THRESH);
    if(s.snapped){el.w=Math.max(4,R-s.v);el.x=s.v;guideXs.push(s.v)}
  }
  if(mode==='se'||mode==='sw'){
    const b=el.y+el.h;
    const s=snapTo(b,ys,SNAP_THRESH);
    if(s.snapped){el.h=Math.max(4,s.v-el.y);guideYs.push(s.v)}
  }
  if(mode==='ne'||mode==='nw'){
    const T=el.y,B=el.y+el.h;
    const s=snapTo(T,ys,SNAP_THRESH);
    if(s.snapped){el.h=Math.max(4,B-s.v);el.y=s.v;guideYs.push(s.v)}
  }
  return {guideXs,guideYs};
}

// Pointer interactions on stage canvas — drag only after threshold; double-click edits text in place
let _stagePtr = null; // {id, pointerId, startX, startY, el, node, handle, moved}
document.addEventListener('pointerdown',ev=>{
  if(live||isSongItem(A))return;
  if(ev.target.closest&&ev.target.closest('#el-inline-edit,[contenteditable="true"],[contenteditable="plaintext-only"],[data-editing="1"]'))return;
  const stage=$('#stage');if(!stage||!stage.contains(ev.target))return;
  // If currently editing text, don't start a drag
  if(document.querySelector('#stage [data-editing="1"]'))return;
  const lockBtn=ev.target.closest&&ev.target.closest('.lock-btn');
  if(lockBtn){
    ev.preventDefault();ev.stopPropagation();
    const lockNode=lockBtn.closest('[data-el]');
    const lockPart=sls(A)[aj];
    const lockEl=lockNode&&((lockPart&&lockPart.elements)||[]).find(x=>x.id===lockNode.dataset.el);
    if(lockEl){pushUndo();lockEl.lock=!lockEl.lock;saveSoon();lockBtn.classList.toggle('locked',!!lockEl.lock);lockBtn.title=lockEl.lock?'Aspect ratio LOCKED (click to unlock)':'Aspect ratio FREE - stretches (click to lock)';lockBtn.classList.remove('pop');void lockBtn.offsetWidth;lockBtn.classList.add('pop');toast(lockEl.lock?'Aspect ratio locked':'Aspect ratio unlocked: stretches freely')}
    return;
  }

  const handle=ev.target.closest('.handle');
  const node=ev.target.closest('[data-el]');
  if(node){
    const id=node.dataset.el;
    const part=sls(A)[aj];
    if(part&&typeof ensureElements==='function')ensureElements(part,A);
    const el=(part&&part.elements||[]).find(e=>e.id===id);
    if(!el)return;
    // Select without rebuilding the whole stage (rStage would kill double-click)
    elSel=id;
    if(!isSongItem(A))tab='Structure';
    document.querySelectorAll('#stage .el').forEach(n=>{
      const on=n.dataset.el===id;
      n.classList.toggle('selected',on);
      // Ensure handles exist when selected
      if(on&&!n.querySelector('.handle')){
        n.insertAdjacentHTML('beforeend',elHandlesHtml(el));
      }
      if(!on){n.querySelectorAll('.handle,.lock-btn').forEach(h=>h.remove())}
    });
    // Light panel refresh for selection list only (no stage rebuild)
    try{
      const list=$('#el-list');
      if(list)list.querySelectorAll('.el-row').forEach(r=>r.classList.toggle('cur',r.dataset.elId===id||r.dataset.pickEl===id));
    }catch{}
    _stagePtr={
      id, pointerId:ev.pointerId, startX:ev.clientX, startY:ev.clientY,
      el, node, handle:handle?handle.dataset.h:null, moved:false, t:Date.now()
    };
    // Resize/rotate handles: capture immediately
    if(handle){
      const slEl=node.closest('.sl');if(!slEl)return;
      const rect=slEl.getBoundingClientRect();
      if(el.kind!=='text'&&handle.dataset.h!=='rot')normalizeMediaBox(el,node,rect.width||1,rect.height||1);
      elDrag={
        el, node, mode:handle.dataset.h,
        startX:ev.clientX,startY:ev.clientY,
        orig:{x:+el.x||0,y:+el.y||0,w:+el.w||10,h:+el.h||10,rot:+el.rot||0},
        rectW:rect.width||1,rectH:rect.height||1,
        moved:false, pointerId:ev.pointerId, deferCapture:false
      };
      try{node.setPointerCapture(ev.pointerId)}catch{}
      ev.preventDefault();
      ev.stopPropagation();
      _stagePtr=null;
    }
    // For move: wait for threshold so double-click can fire
    return;
  }else{
    const t=ev.target;
    const isEl=t&&t.closest&&(t.closest('.el')||t.closest('.handle')||t.closest('.lock-btn')||t.closest('#panel')||t.closest('.outline-menu')||t.closest('.ctx-menu'));
    if(!isEl&&elSel){
      elSel=null;
      document.querySelectorAll('#stage .el.selected').forEach(n=>{
        n.classList.remove('selected');
        n.querySelectorAll('.handle,.lock-btn').forEach(h=>h.remove());
      });
      try{rPanel()}catch{}
    }
    if(t&&(t.id==='strip'||(t.closest&&t.closest('.mid')&&!t.closest('.cv')&&!t.closest('.sl')))){
      if(stripFocus!=='none'){
        stripFocus='none';
        document.querySelectorAll('#strip .th.on').forEach(n=>n.classList.remove('on'));
        document.querySelectorAll('#strip .hb.on').forEach(n=>n.classList.remove('on'));
        document.querySelectorAll('#strip .th.ms-on').forEach(n=>n.classList.remove('ms-on'));
      }
    }
  }
});
document.addEventListener('pointermove',ev=>{
  // Promote delayed stage pointer into a real drag after threshold
  if(_stagePtr&&ev.pointerId===_stagePtr.pointerId&&!elDrag){
    const dist=Math.hypot(ev.clientX-_stagePtr.startX,ev.clientY-_stagePtr.startY);
    if(dist>=6){
      _stagePtr.moved=true;
      const node=_stagePtr.node;
      const el=_stagePtr.el;
      const slEl=node&&node.closest('.sl');
      if(slEl&&el){
        const rect=slEl.getBoundingClientRect();
        elDrag={
          el, node, mode:'move',
          startX:_stagePtr.startX,startY:_stagePtr.startY,
          orig:{x:+el.x||0,y:+el.y||0,w:+el.w||10,h:+el.h||10,rot:+el.rot||0},
          rectW:rect.width||1,rectH:rect.height||1,
          moved:true, pointerId:ev.pointerId, deferCapture:false
        };
        try{node.setPointerCapture(ev.pointerId)}catch{}
      }
      _stagePtr=null;
    }
  }
  if(!elDrag)return;
  const {el,node,mode,orig,rectW,rectH,startX,startY}=elDrag;
  if(!el)return;
  const dx=((ev.clientX-startX)/rectW)*100;
  const dy=((ev.clientY-startY)/rectH)*100;
  if(Math.abs(dx)>0.05||Math.abs(dy)>0.05){
    if(!elDrag.moved)pushUndo();
    elDrag.moved=true;
    if(elDrag.deferCapture){
      elDrag.deferCapture=false;
      try{node&&node.setPointerCapture(elDrag.pointerId)}catch{}
    }
  }
  let guides={guideXs:[],guideYs:[]};
  const part=sls(A)[aj];
  const slEl=node&&node.closest('.sl');
  if(mode==='move'){
    el.x=Math.max(-5,Math.min(95,orig.x+dx));
    el.y=Math.max(-5,Math.min(95,orig.y+dy));
    // Hold Alt to disable snapping
    if(!ev.altKey)guides=applyMoveSnap(el,part);
  }else if(el.kind!=='text'&&(mode==='se'||mode==='ne'||mode==='sw'||mode==='nw')){
    guides=resizeMediaEl(el,mode,orig,dx,dy,ev,part,rectW,rectH);
  }else if(mode==='se'){
    el.w=Math.max(4,Math.min(100,orig.w+dx));
    el.h=Math.max(4,Math.min(100,orig.h+dy));
    if(!ev.altKey)guides=applyResizeSnap(el,mode,part);
  }else if(mode==='nw'){
    const nw=Math.max(4,Math.min(100,orig.w-dx));
    const nh=Math.max(4,Math.min(100,orig.h-dy));
    el.x=orig.x+(orig.w-nw);el.y=orig.y+(orig.h-nh);
    el.w=nw;el.h=nh;
    if(!ev.altKey)guides=applyResizeSnap(el,mode,part);
  }else if(mode==='ne'){
    const nw=Math.max(4,Math.min(100,orig.w+dx));
    const nh=Math.max(4,Math.min(100,orig.h-dy));
    el.y=orig.y+(orig.h-nh);
    el.w=nw;el.h=nh;
    if(!ev.altKey)guides=applyResizeSnap(el,mode,part);
  }else if(mode==='sw'){
    const nw=Math.max(4,Math.min(100,orig.w-dx));
    const nh=Math.max(4,Math.min(100,orig.h+dy));
    el.x=orig.x+(orig.w-nw);
    el.w=nw;el.h=nh;
    if(!ev.altKey)guides=applyResizeSnap(el,mode,part);
  }else if(mode==='rot'){
    const slEl2=node&&node.closest('.sl');
    const r=slEl2?slEl2.getBoundingClientRect():{left:0,top:0};
    const midX=r.left+rectW*((el.x+el.w/2)/100);
    const midY=r.top+rectH*((el.y+el.h/2)/100);
    let deg=Math.atan2(ev.clientY-midY,ev.clientX-midX)*180/Math.PI+90;
    // Normalize to -180..180 first so "near horizontal/vertical" is simple
    deg=((deg+180)%360+360)%360-180;
    // Only snap when close to axis-aligned angles (smooth everywhere else)
    if(!ev.altKey){
      const SNAP=10; // degrees — must be nearer than this to stick
      // Horizontal: 0° or ±180° ; Vertical: ±90°
      const targets=[0,90,-90,180,-180];
      let nearest=deg, nearestDist=SNAP+1;
      for(const t of targets){
        let d=Math.abs(deg-t);
        if(d>180)d=360-d;
        if(d<nearestDist){nearestDist=d;nearest=t}
      }
      if(nearestDist<=SNAP){
        deg=nearest===-180?180:nearest;
      }
    }
    el.rot=Math.round(deg*10)/10;
  }
  if(node&&node.isConnected){
    node.style.left=el.x+'%';
    node.style.top=el.y+'%';
    node.style.width=el.w+'%';
    node.style.height=el.h+'%';
    node.style.transform='rotate('+(el.rot||0)+'deg)';
  }
  if(mode!=='rot'&&!ev.altKey)showSnapGuides(slEl,guides.guideXs,guides.guideYs);
  else clearSnapGuides();
  ev.preventDefault();
});
document.addEventListener('pointerup',ev=>{
  // Double-click detection for text boxes (works even when native dblclick is eaten)
  if(_stagePtr&&ev.pointerId===_stagePtr.pointerId){
    const ptr=_stagePtr;_stagePtr=null;
    if(!ptr.moved){
      const now=Date.now();
      if(_elClickId===ptr.id&&(now-_elClickT)<500){
        _elClickT=0;_elClickId='';
        const part=sls(A)[aj];
        const el=(part&&part.elements||[]).find(x=>x.id===ptr.id);
        if(el&&el.kind==='text'){
          const node=document.querySelector('#stage [data-el="'+ptr.id+'"]')||ptr.node;
          openTextElEditor(el, node);
          return;
        }
      }
      _elClickT=now;_elClickId=ptr.id;
      // Single click: refresh panel for selection
      try{rPanel()}catch{}
      return;
    }
  }
  if(!elDrag)return;
  if(ev.pointerId!=null&&elDrag.pointerId!=null&&ev.pointerId!==elDrag.pointerId)return;
  const wasMove=elDrag.moved;
  elDrag=null;
  clearSnapGuides();
  saveSoon();
  rPanel();
  if(wasMove)updateCurrentThumb();
  else{
    // keep selection handles without full rebuild when possible
    try{updateCurrentThumb()}catch{}
  }
});
document.addEventListener('pointercancel',()=>{_stagePtr=null;elDrag=null;clearSnapGuides()});
document.addEventListener('contextmenu',ev=>{
  if(live||isSongItem(A))return;
  const node=ev.target.closest('#stage [data-el]');
  if(!node){closeCanvasCtx();return}
  ev.preventDefault();
  selectEl(node.dataset.el);
  openCanvasCtx(ev.clientX,ev.clientY);
});
document.addEventListener('click',ev=>{
  if(ev.target.closest&&ev.target.closest('#canvas-ctx')){
    const btn=ev.target.closest('[data-ctx]');if(!btn)return;
    const c=btn.dataset.ctx;
    if(c==='copy')copySelectedEl();
    else if(c==='paste')pasteCanvasEl();
    else if(c==='duplicate')duplicateSelectedEl();
    else if(c==='delete')deleteSelectedEl();
    else if(c==='front')layerSelected('front');
    else if(c==='back')layerSelected('back');
    else if(c==='forward')layerSelected('forward');
    else if(c==='backward')layerSelected('backward');
    if(c!=='layer'&&c!=='layer-label')closeCanvasCtx();
    return;
  }
  if(!ev.target.closest||!ev.target.closest('#canvas-ctx'))closeCanvasCtx();
});

// Unified clipboard paste: screenshots/images, canvas elements, and slides
document.addEventListener('paste',async ev=>{
  if(live)return;
  if(isTypingTarget(ev.target))return;
  const cd=ev.clipboardData;if(!cd)return;
  const items=cd.items;
  if(items&&items.length){
    for(const it of items){
      if(it.type&&it.type.startsWith('image/')){
        const file=it.getAsFile();if(!file)continue;
        ev.preventDefault();
        if(isSongItem(A)){
          toast('Images cannot be pasted onto song slides.');
          return;
        }
        const id='m'+Date.now().toString(36)+Math.random().toString(36).slice(2,8);
        await mPut(id,file);mediaURLs[id]=URL.createObjectURL(file);
        mediaMeta.push({id,name:file.name||'Pasted screenshot',kind:'image',size:file.size,createdAt:Date.now(),source:'paste'});saveMediaMeta();
        const part=sls(A)[aj];if(!part)return;
        const els=ensureElements(part,A);
        const el={id:uid(),kind:'image',src:mediaURLs[id],mediaId:id,x:25,y:25,w:40,h:40,rot:0,z:(els.reduce((m,e)=>Math.max(m,e.z||1),0)+1)};
        els.push(el);elSel=el.id;saveSoon();rStage();rPanel();
        toast('Image pasted');
        return;
      }
    }
  }
  // paste copied element
  let htmlData='',txt='';
  try{htmlData=cd.getData('text/html')||'';txt=cd.getData('text/plain')||''}catch{}
  if((htmlData.indexOf('data-church-presenter-el')>=0||txt===EL_CLIP_MARK)&&canvasClip&&!isSongItem(A)){
    ev.preventDefault();
    pasteCanvasEl();
    return;
  }
  // slide paste
  if(txt.startsWith('<!--church-presenter-slide-->')){
    try{slideClip=JSON.parse(txt.slice('<!--church-presenter-slide-->'.length))}catch{}
  }
  if(!elSel&&slideClip){
    ev.preventDefault();
    pasteSlideBelow();
    return;
  }
  if(canvasClip&&!isSongItem(A)){
    pasteCanvasEl();
    ev.preventDefault();
    return;
  }
});



// ===== Element list layer reorder (shared pill animation) =====
document.addEventListener('pointerdown',ev=>{
  const row0=ev.target.closest&&ev.target.closest('#el-list .el-row');const handle=(row0&&!ev.target.closest('button'))?row0.querySelector('[data-el-drag]'):null;
  if(!handle||live||isSongItem(A))return;
  const row=handle.closest('.el-row');
  if(!row)return;
  const id=handle.dataset.elDrag;
  const part=sls(A)[aj];
  if(!part||!(part.elements||[]).find(e=>e.id===id))return;
  ev.preventDefault();
  ev.stopPropagation();
  let to=null;
  const clearUi=()=>document.querySelectorAll('#el-list .el-row').forEach(r=>r.classList.remove('drag-over','drag-over-below'));
  pillDrag(ev,{
    el:row,label:(row.querySelector('.el-label')?.textContent||'Layer').trim(),
    scrollEl:()=>$('#panel'),
    onMove:(x,y)=>{
      clearUi();to=null;
      const list=$('#el-list');if(!list)return;
      const rows=[...list.querySelectorAll('.el-row')];
      const lr=list.getBoundingClientRect();
      if(!rows.length||x<lr.left-80||x>lr.right+80)return;
      let k=rows.length;
      for(let i=0;i<rows.length;i++){
        const r=rows[i].getBoundingClientRect();
        if(y<r.top+r.height/2){k=i;break}
      }
      to=k;
      if(k<rows.length)rows[k].classList.add('drag-over');
      else rows[rows.length-1].classList.add('drag-over-below');
    },
    onDrop:()=>{
      if(to==null)return;
      const p=sls(A)[aj];if(!p||!p.elements)return;
      const list=$('#el-list');if(!list)return;
      const orderedIds=[...list.querySelectorAll('.el-row')].map(r=>r.dataset.elId||r.dataset.pickEl).filter(Boolean);
      const from=orderedIds.indexOf(id);
      if(from<0)return;
      let dest=to;
      if(dest>from)dest--;
      dest=Math.max(0,Math.min(orderedIds.length-1,dest));
      if(dest===from)return;
      orderedIds.splice(from,1);
      orderedIds.splice(dest,0,id);
      const n=orderedIds.length;
      orderedIds.forEach((eid,i)=>{const el=p.elements.find(e=>e.id===eid);if(el)el.z=n-i});
      elSel=id;
      saveSoon();rStage();rPanel();updateCurrentThumb();
    },
    onEnd:clearUi
  });
});

/* Presenter view controls */
$('#pv-end')?.addEventListener('click',()=>endPresent());
$('#pv-prev')?.addEventListener('click',()=>go(-1));
$('#pv-next')?.addEventListener('click',()=>go(1));
document.getElementById('presenter-view')?.addEventListener('click',ev=>{
  const oi=ev.target.closest('[data-pv-oi]');
  if(oi){jumpToItem(+oi.dataset.pvOi);return}
  const th=ev.target.closest('[data-pv-i][data-pv-j]');
  if(th){
    imageSel=-1;A=S[+th.dataset.pvI];aj=+th.dataset.pvJ;sel=+th.dataset.pvI;si=aj;clr=null;rAll();
  }
  // Clear buttons (B/W/L) are handled by the global click handler only —
  // handling them here too would toggle twice and cancel the clear.
});

/* ===== Cloud button status, Drive window layout, text context menu ===== */
if(!IS_AUDIENCE){

/* ---- 1. One cloud button: cloud icon, status badge, status dot ---- */
const SYNC_COLORS={synced:'#2e9e6b',pending:'#d89b1a',readonly:'#d89b1a',syncing:'#3b82f6',offline:'#8a97a5',local:'#8a97a5',paused:'#8a97a5',update:'#7c5cff',conflict:'#d33',error:'#d33',lost:'#d33'};
const SYNC_BADGE={
  synced:{c:'#2e9e6b',svg:'<path d="M2.8 6.3l2.2 2.2 4.2-4.6"/>'},
  syncing:{c:'#3b82f6',cls:'cb-spin',svg:'<path d="M9.5 6a3.5 3.5 0 1 1-1.2-2.6"/><path d="M9.6 2.2v2.2H7.4"/>'},
  pending:{c:'#d89b1a',svg:'<path d="M6 9.5V2.8M3.2 5.4L6 2.6l2.8 2.8"/>'},
  readonly:{c:'#d89b1a',svg:'<path d="M6 9.5V2.8M3.2 5.4L6 2.6l2.8 2.8"/>'},
  update:{c:'#7c5cff',svg:'<path d="M6 2.5v6.7M3.2 6.6L6 9.4l2.8-2.8"/>'},
  conflict:{c:'#d33',svg:'<path d="M6 2.6v4M6 9.2v.1"/>'},
  error:{c:'#d33',svg:'<path d="M6 2.6v4M6 9.2v.1"/>'},
  lost:{c:'#d33',svg:'<path d="M6 2.6v4M6 9.2v.1"/>'}
};
const CLOUD_PATH='M7 18a4.5 4.5 0 0 1 0-9 5.5 5.5 0 0 1 10.2-1.8A4 4 0 1 1 18 18H7z';
const cssA=document.createElement('style');
cssA.textContent=`
#sync-chip{display:none!important}
#cloud-btn{width:auto;min-width:calc(36px * var(--ui-scale));padding:0 11px;gap:9px}
#cloud-btn .cb-icon{position:relative;display:inline-flex;width:22px;height:22px}
#cloud-btn .cb-icon>svg{width:22px;height:22px}
#cloud-btn .cb-badge{position:absolute;right:-6px;bottom:-5px;width:13px;height:13px;border-radius:50%;display:grid;place-items:center;box-shadow:0 0 0 2px var(--pn)}
#cloud-btn .cb-badge svg{width:9px;height:9px;fill:none;stroke:#fff;stroke-width:1.7;stroke-linecap:round;stroke-linejoin:round}
#cloud-btn .cb-spin{animation:cbspin 1s linear infinite}
@keyframes cbspin{to{transform:rotate(360deg)}}
#cloud-btn .cb-dot{width:9px;height:9px;border-radius:50%;background:var(--cb,#8a97a5);flex:none}
#cloud-btn.s-syncing .cb-dot{animation:syncpulse 1s infinite}
@media(prefers-reduced-motion:reduce){#cloud-btn .cb-spin,#cloud-btn.s-syncing .cb-dot{animation:none}}

/* Drive window */
#cloud-modal .app-dialog{width:min(560px,calc(100vw - 32px))!important}
.gd-card{border:1px solid var(--ln);border-radius:10px;padding:14px;margin:0 0 12px;background:var(--bg)}
.gd-row{display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.gd-grow{flex:1;min-width:0}
.gd-label{font-size:calc(11px * var(--ui-scale));text-transform:uppercase;letter-spacing:.06em;color:var(--mu);margin-bottom:3px}
.gd-val{font-weight:600;overflow-wrap:anywhere;display:flex;align-items:center;gap:8px}
.gd-sub{font-size:calc(12px * var(--ui-scale));color:var(--mu);margin:6px 0 0}
.gd-actions{display:flex;gap:6px;flex-wrap:wrap}
.gd-dot{width:9px;height:9px;border-radius:50%;background:#8a97a5;flex:none}
.gd-pill{display:inline-flex;align-items:center;gap:6px;font-size:calc(12px * var(--ui-scale));font-weight:600;padding:3px 10px;border-radius:999px;border:1px solid var(--ln);background:var(--pn)}
.gd-alert{border:1px solid #d33;background:#d3333314;color:#c22;border-radius:8px;padding:10px 12px;margin:0 0 12px;font-size:calc(13px * var(--ui-scale))}
.gd-busy{display:flex;align-items:center;gap:8px;border:1px solid var(--ln);border-radius:8px;padding:9px 12px;margin:0 0 12px;font-size:calc(13px * var(--ui-scale));color:var(--mu)}
.gd-spin{width:12px;height:12px;border:2px solid var(--ln);border-top-color:var(--ac);border-radius:50%;animation:cbspin .8s linear infinite}
.gd-toggle{display:flex;align-items:center;gap:8px;margin:12px 0 0;padding-top:12px;border-top:1px solid var(--ln);font-size:calc(13px * var(--ui-scale))}
.gd-toggle input{width:auto;margin:0}
.gd-divider{border-top:1px solid var(--ln);margin:12px 0 0;padding-top:12px}
.gd-footer{display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;margin-top:4px}
.gd-hero{text-align:center;padding:22px 16px}
.gd-hero h3{margin:0 0 6px;font-size:calc(16px * var(--ui-scale))}
.gd-hero p{margin:0 0 14px;color:var(--mu);font-size:calc(13px * var(--ui-scale))}
`;
document.head.appendChild(cssA);

ensureSyncChip=function(){};
document.getElementById('sync-chip')?.remove();

renderSyncChip=function(){
  const btn=document.getElementById('cloud-btn');if(!btn)return;
  const st=svcSyncStatus(activeService);
  const key=st.key;
  if(btn.dataset.state!==key){
    btn.dataset.state=key;
    const badge=SYNC_BADGE[key];
    const slash=badge?'':'<line x1="3.5" y1="3.5" x2="20.5" y2="20.5" style="stroke:var(--pn)" stroke-width="5"/><line x1="3.5" y1="3.5" x2="20.5" y2="20.5"/>';
    btn.innerHTML=
      '<span class="cb-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="'+CLOUD_PATH+'"/>'+slash+'</svg>'+
      (badge?'<span class="cb-badge" style="background:'+badge.c+'"><svg viewBox="0 0 12 12" class="'+(badge.cls||'')+'" aria-hidden="true">'+badge.svg+'</svg></span>':'');
    btn.className='icon-btn s-'+key;
    btn.style.setProperty('--cb',SYNC_COLORS[key]||'#8a97a5');
  }
  const link=getLink(services[activeService]);
  const label=cloudUi.connected?st.label:'Google Drive not connected';
  btn.title='Google Drive: '+label+(link&&link.syncedAt?'\nLast synced '+fmtDriveDate(link.syncedAt):'');
  btn.setAttribute('aria-label','Google Drive: '+label);
};
renderSyncChip();

/* ---- 2. Redesigned Google Drive window ---- */
function gdServiceCard(dis){
  const i=activeService,svc=services[i];if(!svc)return '';
  const link=getLink(svc),st=svcSyncStatus(i);
  const btn=(act,label,pri)=>`<button type="button"${pri?' class="pri"':''} data-cloud-act="${act}"${dis}>${label}</button>`;
  let main='',more='';
  if(!link)main=btn('upload','Upload to Drive',true);
  else if(link.conflict)main=btn('conflict','Resolve conflict…',true);
  else if(link.lost)main=btn('reupload','Upload as new file',true)+btn('unlink','Keep local only');
  else if(st.key==='update')main=btn('pull','Load update from Drive',true);
  else if(link.canEdit===false)main=btn('copy','Save as copy');
  else main=btn('sync','Sync now',true);
  if(link&&link.webViewLink)more+=btn('open','Open in Drive')+btn('copylink','Copy link');
  if(link&&link.backups&&link.backups.length)more+=btn('restore','Restore local backup');
  const miss=link?missingMedia(i===activeService?S:svc.slides):[];
  return `<section class="gd-card">
    <div class="gd-row">
      <div class="gd-grow"><div class="gd-label">This service</div><div class="gd-val">${e(formatServiceTitle(svc))}</div></div>
      <span class="gd-pill"><span class="gd-dot" style="background:${SYNC_COLORS[st.key]||'#8a97a5'}"></span>${e(st.label)}</span>
    </div>
    ${link&&link.syncedAt?`<p class="gd-sub">Last synced ${e(fmtDriveDate(link.syncedAt))}</p>`:''}
    ${st.detail?`<p class="gd-sub" style="color:#c22">${e(st.detail)}</p>`:''}
    ${miss.length?`<div class="cloud-warn">⚠ ${miss.length} media item${miss.length===1?'':'s'} (${e([...new Set(miss)].slice(0,4).join(', '))}) not on this computer. Sync the service package to download its included media.</div>`:''}
    <div class="gd-actions" style="margin-top:12px">${main}</div>
    ${more?`<div class="gd-actions gd-divider">${more}</div>`:''}
  </section>`;
}
renderCloudModal=function(){
  const m=$('#cloud-modal');if(!m||m.hidden)return;
  const body=$('#cloud-body');if(!body)return;
  const dis=cloudUi.busy?' disabled':'';
  const auto=loadCloudMeta().autoUpload!==false;
  let h='';
  if(cloudUi.busy)h+=`<div class="gd-busy"><span class="gd-spin"></span>${e(cloudUi.msg||'Working…')}</div>`;
  if(cloudUi.authLost)h+='<div class="gd-alert">Google sign-in expired or was revoked. Reconnect to resume syncing. Your local services are safe.</div>';
  else if(cloudUi.error)h+=`<div class="gd-alert">${e(cloudUi.error)}</div>`;

  if(!cloudUi.connected){
    h+=`<section class="gd-card gd-hero"><h3>Share services between computers</h3>
      <p>Connect a Google account to keep your services in a Drive folder. Everything keeps working offline and syncs when you are back online.</p>
      <button type="button" id="cloud-connect" class="pri"${dis}>Connect Google</button></section>`;
    body.innerHTML=h;return;
  }

  h+=`<section class="gd-card"><div class="gd-row">
      <div class="gd-grow"><div class="gd-label">Account</div>
        <div class="gd-val"><span class="gd-dot" style="background:${cloudUi.authLost?'#d33':'#2e9e6b'}"></span>${e(cloudUi.email||'Connected')}</div></div>
      <div class="gd-actions">
        ${cloudUi.authLost?`<button type="button" id="cloud-connect" class="pri"${dis}>Reconnect</button>`:''}
        <button type="button" id="cloud-switch"${dis}>Switch account</button>
        <button type="button" id="cloud-disconnect"${dis}>Disconnect</button>
      </div></div></section>`;

  h+=`<section class="gd-card"><div class="gd-row">
      <div class="gd-grow"><div class="gd-label">Drive folder</div>
        <div class="gd-val">${cloudUi.folderId?e(cloudUi.folderName||'Presentation folder'):'<span style="color:var(--mu);font-weight:400">Not set yet</span>'}</div></div>
      <button type="button" id="cloud-pick-folder"${cloudUi.folderId?'':' class="pri"'}${dis}>${cloudUi.folderId?'Change folder…':'Set folder…'}</button>
    </div>
    <p class="gd-sub">Paste a folder link or ID from Google Drive. Services are saved there as .cpres files.</p>
    <label class="gd-toggle"><input type="checkbox" id="cloud-auto-upload" ${auto?'checked':''}> Automatically save edited services to Drive</label>
  </section>`;

  if(cloudUi.folderId)h+=gdServiceCard(dis);

  if(cloudUi.folderId){
    h+=`<div class="gd-footer"><span class="gd-sub" style="margin:0">${cloudUi.online===false?'Offline. ':''}${cloudUi.lastSyncAt?'Last checked '+e(fmtDriveDate(cloudUi.lastSyncAt)):'Not checked yet'}</span></div>`;
  }

  const slides=cloudUi.files.filter(f=>driveKind(f)==='slides');
  if(slides.length){
    h+=`<details style="margin-top:14px"><summary style="cursor:pointer;color:var(--mu)">Import a PowerPoint, PDF or Google Slides file (${slides.length})</summary>`
      +slides.map(f=>`<div class="service-card" style="margin-top:8px"><div class="info"><h3 style="font-size:calc(14px * var(--ui-scale))">${e(f.name)}</h3><small>${e(fmtDriveDate(f.modifiedTime))}</small></div><button type="button" class="pri" data-cloud-import="${e(f.id)}"${dis}>Import</button></div>`).join('')
      +'</details>';
  }
  body.innerHTML=h;
};

/* ---- 3. Replace the default right click menu with a text menu ---- */
const cssB=document.createElement('style');
cssB.textContent=`
#text-ctx{position:fixed;z-index:1200;min-width:140px;background:var(--pn);border:1px solid var(--ln);border-radius:8px;box-shadow:0 10px 28px #0004;padding:4px;display:none}
#text-ctx.open{display:block}
#text-ctx button{display:block;width:100%;text-align:left;border:0;background:none;padding:8px 12px;border-radius:6px;cursor:pointer;font:inherit;color:inherit}
#text-ctx button:hover:not(:disabled){background:var(--hl)}
#text-ctx button:disabled{opacity:.4;cursor:default}
`;
document.head.appendChild(cssB);
const textMenu=document.createElement('div');
textMenu.id='text-ctx';textMenu.setAttribute('role','menu');
textMenu.innerHTML='<button type="button" data-text-ctx="copy">Copy</button><button type="button" data-text-ctx="paste">Paste</button>';
document.body.appendChild(textMenu);
let textCtxState=null;

const hideTextMenu=()=>{textMenu.classList.remove('open');textCtxState=null};
const fieldOf=t=>{
  const f=t&&t.closest?t.closest('textarea,input,[contenteditable]'):null;
  if(!f||f.disabled||f.readOnly)return null;
  if(f.tagName==='INPUT')return /^(text|email|search|url|tel|password|)$/i.test(f.type||'')?f:null;
  if(f.tagName==='TEXTAREA'||f.isContentEditable)return f;
  return null;
};
const selectedTextOf=t=>{
  const f=t&&t.closest?t.closest('textarea,input'):null;
  if(f){
    try{if(f.selectionStart!==f.selectionEnd)return f.value.slice(f.selectionStart,f.selectionEnd)}catch{}
    return '';
  }
  return String(window.getSelection?window.getSelection().toString():'');
};
async function copyToClipboard(text){
  try{await navigator.clipboard.writeText(text);return true}catch{}
  try{
    const ta=document.createElement('textarea');ta.value=text;ta.style.cssText='position:fixed;opacity:0';
    document.body.appendChild(ta);ta.select();const ok=document.execCommand('copy');ta.remove();return ok;
  }catch{return false}
}
function pasteInto(f,text){
  f.focus();
  if(f.tagName==='TEXTAREA'||f.tagName==='INPUT'){
    const s=f.selectionStart==null?f.value.length:f.selectionStart;
    const end=f.selectionEnd==null?s:f.selectionEnd;
    try{f.setRangeText(text,s,end,'end')}catch{f.value=f.value.slice(0,s)+text+f.value.slice(end)}
    f.dispatchEvent(new Event('input',{bubbles:true}));
  }else document.execCommand('insertText',false,text);
}

/* Registered last, so any existing custom menu has already handled its own right click.
   If one did (defaultPrevented), this does nothing and never fights with it. */
document.addEventListener('contextmenu',ev=>{
  if(ev.defaultPrevented)return;
  ev.preventDefault();
  const t=ev.target;
  if(t.closest&&t.closest('.outline-menu,.ctx-menu,#text-ctx,#service-list .service-card'))return;
  document.querySelectorAll('.outline-menu.open,.ctx-menu.open').forEach(m=>{m.classList.remove('open');m.hidden=true});
  if(live||document.body.classList.contains('presentation-mode')){hideTextMenu();return}
  const text=selectedTextOf(t),field=fieldOf(t);
  if(!text&&!field){hideTextMenu();return}
  textCtxState={text,field};
  textMenu.querySelector('[data-text-ctx="copy"]').disabled=!text;
  textMenu.querySelector('[data-text-ctx="paste"]').disabled=!field;
  textMenu.classList.add('open');
  const w=textMenu.offsetWidth||140,h=textMenu.offsetHeight||80;
  textMenu.style.left=Math.max(4,Math.min(ev.clientX,window.innerWidth-w-8))+'px';
  textMenu.style.top=Math.max(4,Math.min(ev.clientY,window.innerHeight-h-8))+'px';
});
textMenu.addEventListener('mousedown',ev=>ev.preventDefault());   // keep focus and selection in the text
textMenu.addEventListener('click',async ev=>{
  ev.stopPropagation();
  const b=ev.target.closest('[data-text-ctx]');
  if(!b||b.disabled||!textCtxState)return;
  const {text,field}=textCtxState,act=b.dataset.textCtx;
  hideTextMenu();
  if(act==='copy'){toast(await copyToClipboard(text)?'Copied':'Could not copy.')}
  else if(act==='paste'){
    let clip='';
    try{clip=await navigator.clipboard.readText()}catch{toast('Clipboard access was blocked.');return}
    if(clip&&field)pasteInto(field,clip);
  }
});
document.addEventListener('click',ev=>{if(!ev.target.closest||!ev.target.closest('#text-ctx'))hideTextMenu()},true);
document.addEventListener('keydown',ev=>{if(ev.key==='Escape')hideTextMenu()},true);
document.addEventListener('scroll',hideTextMenu,true);
window.addEventListener('resize',hideTextMenu);
window.addEventListener('blur',hideTextMenu);

}


/* ===== Services page: one mixed list (local + Drive), delete from Drive ===== */
/* ---- Service tags + templates list: module scope ----
   These used to sit inside an if(!IS_AUDIENCE){...} block. ES modules are strict, so
   block-level functions are not visible outside the block and every caller outside it
   threw ReferenceError (which stopped the rest of the file from loading). ---- */
function serviceTagList(){
  return ['AM','PM','Youth','Other'];
}
function svcTitleHtml(s,i){
  const tag=(s.tag||'').trim(),name=(s.name||'').trim();
  const tagBtn='<button type="button" class="svc-tag-btn" data-tag-service="'+i+'" title="Change tag" aria-label="Change service tag: '+e(tag||'none')+'">'+e(tag||'No tag')+'<span class="svc-caret" aria-hidden="true">\u25be</span></button>';
  const nameBtn='<button type="button" class="service-name-edit" data-rename-service="'+i+'" title="Rename service" aria-label="Rename service'+(name?': '+e(name):'')+'">'+(name?'<span class="service-name-text">'+e(name)+'</span>':'')+'<span class="service-name-pen" aria-hidden="true">\u270e</span></button>';
  return '<div class="svc-title-row">'+tagBtn+nameBtn+'</div>';
}
function renderTitleBar(svc){
  let tb=document.getElementById('service-tag-btn');
  if(!tb){
    // The editor header only had the title button; add the tag button just before it.
    const te=document.getElementById('service-title-edit');
    if(!te||!te.parentNode||document.body.classList.contains('home-mode'))return;
    tb=document.createElement('button');
    tb.type='button';tb.id='service-tag-btn';tb.className='svc-tag-btn';
    tb.title='Change tag';tb.setAttribute('aria-haspopup','menu');
    te.parentNode.insertBefore(tb,te);
  }
  const tag=((svc&&svc.tag)||'').trim();
  tb.innerHTML=e(tag||'Select tag')+'<span class="svc-caret" aria-hidden="true">\u25be</span>';
  tb.setAttribute('aria-label','Service tag: '+(tag||'none')+'. Change tag');
  const sep=document.getElementById('service-title-sep');
  if(sep)sep.hidden=!tag;
}
function hideTagMenu(){const m=document.getElementById('tag-menu');if(m){m.classList.remove('open');m.hidden=true}}
/* The "only one menu open" pointerdown handler closes #tag-menu before the click lands, so remember
   whether the menu was open on the button the person pressed. Then a second press just closes it. */
let _tagMenuWasOpen=false;
window.addEventListener('pointerdown',ev=>{
  const m=document.getElementById('tag-menu');
  _tagMenuWasOpen=!!(m&&m.classList.contains('open')&&ev.target&&ev.target.closest&&ev.target.closest('[data-tag-service],#service-tag-btn'));
},true);
function showTagMenu(anchor,idx){
  const svc=services[idx];if(!svc)return;
  if(_tagMenuWasOpen){_tagMenuWasOpen=false;hideTagMenu();return}
  let m=document.getElementById('tag-menu');
  if(!m){m=document.createElement('div');m.id='tag-menu';m.className='outline-menu';m.setAttribute('role','menu');document.body.appendChild(m)}
  if(m.classList.contains('open')&&m.dataset.idx===String(idx)&&m._anchor===anchor){hideTagMenu();return}
  const cur=(svc.tag||'').trim();
  m.innerHTML=serviceTagList().map(t=>'<button type="button" role="menuitem" data-tag-pick="'+e(t)+'">'+(t===cur?'\u2713 ':'')+e(t)+'</button>').join('');
  m.dataset.idx=String(idx);m._anchor=anchor;m.hidden=false;m.classList.add('open');
  const r=anchor.getBoundingClientRect(),w=m.offsetWidth||140,h=m.offsetHeight||120;
  m.style.left=Math.max(6,Math.min(r.left,window.innerWidth-w-8))+'px';
  const below=r.bottom+4;
  m.style.top=(below+h<window.innerHeight-8?below:Math.max(6,r.top-h-4))+'px';
}
function pickServiceTag(idx,tag){
  const svc=services[idx];if(!svc||(svc.tag||'')===tag)return;
  svc.tag=tag;svc.lastEdited=Date.now();
  if(document.body.classList.contains('home-mode')){saveState();rHome()}else{saveSoon();rAll()}
}
function fmtDate(ts){try{const d=new Date(ts);const r=relDayLabel(d);return r||d.toLocaleDateString(undefined,{day:'numeric',month:'short',year:'numeric'})}catch{return ''}}
function renderTemplatesHome(){
  if(!document.body.classList.contains('home-mode'))return;
  let box=document.getElementById('home-templates');
  if(!box){
    const anchor=document.getElementById('home-cloud-list')||document.getElementById('service-list');
    if(!anchor)return;
    box=document.createElement('div');box.id='home-templates';
    anchor.insertAdjacentElement('afterend',box);
  }
  const defId=defaultTplId();
  const list=loadTemplates().slice().sort((a,b)=>{
    if(a.id===defId)return -1;
    if(b.id===defId)return 1;
    return String(a.name||'').localeCompare(String(b.name||''));
  });
  let h='<h3 class="tpl-head">Templates</h3>';
  if(!list.length){
    h+='<p class="note tpl-sub">Templates are reusable starting points for new services. Right-click a service and choose \u201cSave as template\u201d, or use Share \u2192 Save service as template.</p>';
    box.innerHTML=h;return;
  }
  h+='<p class="note tpl-sub">Pick \u201cUse\u201d to start a new service from a template.</p><div class="service-list">';
  h+=list.map(t=>{
    const n=(t.slides||[]).length;
    const where=t.driveFileId?(t.dirty?'Waiting to sync':'On Google Drive'):(t.localOnly?'This computer only':(cloudReady()?'Waiting to sync':'This computer only'));
    const isDef=t.id===defId;
    return `<div class="service-card" data-tpl-card="${e(t.id)}">
      <div class="info">
        <div class="tpl-name">
          <button type="button" class="service-name-edit" data-tpl-rename="${e(t.id)}" title="Rename template" aria-label="Rename template: ${e(t.name)}">
            <span class="service-name-text">${e(t.name)}</span>
            <span class="service-name-pen" aria-hidden="true">✎</span>
          </button>
        </div>
        <small>${n} section${n===1?'':'s'} · ${e(where)} · Updated ${e(fmtDate(t.updatedAt))}</small>
      </div>
      <button type="button" class="tpl-default-badge${isDef?' is-default':''}" data-tpl-default="${e(t.id)}">Default</button>
      <button type="button" class="pri" data-tpl-use="${e(t.id)}">Use</button>
      <button type="button" style="color:#c22" data-tpl-del="${e(t.id)}">Delete</button>
    </div>`;
  }).join('');
  h+='</div>';
  box.innerHTML=h;
}

/* ---- "Start from" list on the New service form (one copy, shared by every template block) ---- */
function updateTplNote(){
  const sel=document.getElementById('service-template'),note=document.getElementById('service-template-note');
  if(!sel||!note)return;
  const t=loadTemplates().find(x=>x.id===sel.value);
  if(!t){note.textContent='';return}
  const n=(t.slides||[]).length;
  note.textContent=n+' section'+(n===1?'':'s');
}
function refreshTemplateSelect(pre,applyDefault){
  const sel=document.getElementById('service-template');if(!sel)return;
  const defId=defaultTplId();
  const list=loadTemplates().slice().sort((a,b)=>{
    if(a.id===defId)return -1;
    if(b.id===defId)return 1;
    return String(a.name||'').localeCompare(String(b.name||''));
  });
  const key=list.map(t=>t.id+'|'+t.name).join('\n');
  // Rebuild when the list changed, or when the options were wiped by something else.
  if(sel.dataset.key!==key||sel.options.length!==list.length+1){
    const cur=sel.value;
    sel.dataset.key=key;
    sel.innerHTML='<option value="">Blank service</option>'+list.map(t=>'<option value="'+e(t.id)+'">'+e(t.name)+'</option>').join('');
    sel.value=list.some(t=>t.id===cur)?cur:'';
  }
  if(pre&&list.some(t=>t.id===pre))sel.value=pre;
  else if(applyDefault)sel.value=(defId&&list.some(t=>t.id===defId))?defId:'';   // '' = Blank service
  updateTplNote();
}

/* ---- Default template: toggle it, re-render, and slide the cards into their new places ---- */
function animateTemplateCards(before){
  const box=document.getElementById('home-templates');
  if(!box||!before||!before.size)return;
  if(window.matchMedia&&window.matchMedia('(prefers-reduced-motion: reduce)').matches)return;
  const moved=[];
  box.querySelectorAll('.service-card').forEach(c=>{
    const b=c.querySelector('[data-tpl-default]');if(!b)return;
    const from=before.get(b.dataset.tplDefault);if(!from)return;
    const to=c.getBoundingClientRect();
    const dx=from.left-to.left,dy=from.top-to.top;
    if(Math.abs(dx)<1&&Math.abs(dy)<1)return;
    c.style.transition='none';
    c.style.transform='translate('+dx+'px,'+dy+'px)';
    c.style.position='relative';
    c.style.zIndex=b.classList.contains('is-default')?'2':'1';
    moved.push(c);
  });
  if(!moved.length)return;
  void box.offsetHeight;   // commit the starting positions before animating
  requestAnimationFrame(()=>{
    moved.forEach(c=>{c.style.transition='transform .45s cubic-bezier(.2,.8,.2,1)';c.style.transform=''});
    setTimeout(()=>moved.forEach(c=>{c.style.transition='';c.style.transform='';c.style.position='';c.style.zIndex=''}),520);
  });
}
function setDefaultTemplate(id){
  if(!id)return;
  const wasDefault=isDefaultTpl(id);
  try{
    if(wasDefault)localStorage.removeItem(TPL_DEFAULT_KEY);
    else localStorage.setItem(TPL_DEFAULT_KEY,id);
  }catch{toast('Could not save the default template.');return}
  const before=new Map();
  const box=document.getElementById('home-templates');
  if(box)box.querySelectorAll('.service-card').forEach(c=>{
    const b=c.querySelector('[data-tpl-default]');
    if(b)before.set(b.dataset.tplDefault,c.getBoundingClientRect());
  });
  try{renderTemplatesHome()}catch(err){console.warn('[templates]',err)}
  try{refreshTemplateSelect()}catch{}
  try{animateTemplateCards(before)}catch{}
  toast(wasDefault?'Default template removed.':'New services will start from this template.');
}


if(!IS_AUDIENCE){

const HC_BADGE={
  synced:{c:'#2e9e6b',svg:'<path d="M2.8 6.3l2.2 2.2 4.2-4.6"/>'},
  syncing:{c:'#3b82f6',cls:'hc-spin',svg:'<path d="M9.5 6a3.5 3.5 0 1 1-1.2-2.6"/><path d="M9.6 2.2v2.2H7.4"/>'},
  pending:{c:'#d89b1a',svg:'<path d="M6 9.5V2.8M3.2 5.4L6 2.6l2.8 2.8"/>'},
  update:{c:'#7c5cff',svg:'<path d="M6 2.5v6.7M3.2 6.6L6 9.4l2.8-2.8"/>'},
  cloud:{c:'#3b82f6',svg:'<path d="M6 2.5v6.7M3.2 6.6L6 9.4l2.8-2.8"/>'},
  alert:{c:'#d33',svg:'<path d="M6 2.6v4M6 9.2v.1"/>'}
};
const HC_MAP={synced:'synced',syncing:'syncing',pending:'pending',readonly:'pending',update:'update',conflict:'alert',error:'alert',lost:'alert'};

const hcCss=document.createElement('style');
hcCss.textContent=`
.hc-cloud{position:relative;display:inline-flex;width:20px;height:20px;flex:none;color:var(--mu);margin-right:4px}
.hc-cloud>svg{width:20px;height:20px}
.service-card>.hc-cloud{margin:0 4px 0 2px}
.hc-badge{position:absolute;right:-5px;bottom:-4px;width:12px;height:12px;border-radius:50%;display:grid;place-items:center;box-shadow:0 0 0 2px var(--pn)}
.hc-badge svg{width:8px;height:8px;fill:none;stroke:#fff;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
.hc-spin{animation:hcspin 1s linear infinite}
@keyframes hcspin{to{transform:rotate(360deg)}}
@media(prefers-reduced-motion:reduce){.hc-spin{animation:none}}
.hc-title{margin:0;padding:3px 6px;font-size:calc(16px * var(--ui-scale));font-weight:600;overflow-wrap:anywhere}
.service-date-edit{appearance:none;border:1px solid transparent;border-radius:5px;background:transparent;color:inherit;font:inherit;padding:2px 5px;margin-left:-5px;cursor:pointer;transition:border-color .12s,background .12s}
.service-date-edit:hover,.service-date-edit:focus-visible{border-color:var(--ln);background:var(--bg);outline:none}
.service-date-edit:focus-visible{box-shadow:0 0 0 2px var(--hl)}
`;
document.head.appendChild(hcCss);

document.addEventListener('click',ev=>{
  const button=ev.target.closest&&ev.target.closest('[data-edit-service-date]');if(!button)return;
  const picker=document.querySelector('[data-service-date-picker="'+button.dataset.editServiceDate+'"]');
  if(picker){picker.value=services[+button.dataset.editServiceDate]?.date||'';try{picker.showPicker()}catch{picker.click()}}
});
document.addEventListener('change',async ev=>{
  const picker=ev.target;if(!picker.matches||!picker.matches('[data-service-date-picker]'))return;
  const index=Number(picker.dataset.serviceDatePicker),svc=services[index];if(!svc||!/^\d{4}-\d{2}-\d{2}$/.test(picker.value))return;
  svc.date=picker.value;svc.lastEdited=Date.now();
  try{
    const state=JSON.parse(localStorage.getItem('service-slides-state')||'{}');
    if(Array.isArray(state.services)&&state.services[index]){state.services[index].date=svc.date;state.services[index].lastEdited=svc.lastEdited;localStorage.setItem('service-slides-state',JSON.stringify(state))}
  }catch(err){console.warn('[service-date] Could not save date:',err);toast('Could not save the service date.');return}
  const link=getLink(svc);
  if(link){link.cleanupDone=false;recordLink(svc,{...link,cleanupDone:false,localHash:hashService(index)})}
  rHome();
  if(link&&cloudUi.connected&&cloudUi.online!==false){try{await cloudPush(index,{force:true})}catch(err){console.warn('[service-date] Could not update Drive package:',err);toast('Date saved. Drive media expiry will update at the next sync.')}}
});

function hcIcon(kind,title){
  const b=HC_BADGE[kind];
  const slash=b?'':'<line x1="3.5" y1="3.5" x2="20.5" y2="20.5" style="stroke:var(--pn)" stroke-width="5"/><line x1="3.5" y1="3.5" x2="20.5" y2="20.5"/>';
  return '<span class="hc-cloud" title="'+e(title)+'" role="img" aria-label="'+e(title)+'">'+
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M7 18a4.5 4.5 0 0 1 0-9 5.5 5.5 0 0 1 10.2-1.8A4 4 0 1 1 18 18H7z"/>'+slash+'</svg>'+
    (b?'<span class="hc-badge" style="background:'+b.c+'"><svg viewBox="0 0 12 12" class="'+(b.cls||'')+'" aria-hidden="true">'+b.svg+'</svg></span>':'')+
    '</span>';
}
function hcLocalIcon(i){
  if(!getLink(services[i]))return hcIcon('none','Only on this computer');
  const st=svcSyncStatus(i);
  return hcIcon(HC_MAP[st.key]||'none',st.label);
}

rHome=function(){
  ensureHomeChrome();
  const tagSel=$('#service-tag-filter'),sortSel=$('#service-sort');
  if(tagSel)tagSel.value=serviceTagFilter;
  if(sortSel)sortSel.value=serviceSort;

  const showCloud=cloudUi.connected&&!!cloudUi.folderId;
  const linkedIds=new Set(services.map(s=>{const l=getLink(s);return l&&l.fileId}).filter(Boolean));
  let items=services.map((s,i)=>({type:'local',s,i}));
  if(serviceTagFilter)items=items.filter(x=>x.s.tag===serviceTagFilter);
  else if(showCloud){
    cloudUi.files.filter(f=>driveKind(f)==='cpres'&&!linkedIds.has(f.id)).forEach(f=>items.push({type:'cloud',f}));
  }
  const olderCount=items.filter(x=>x.type==='local'&&servicePastRetention(x.s)&&serviceCleanupComplete(x.s)).length;
  if(!olderServicesVisible)items=items.filter(x=>x.type!=='local'||!servicePastRetention(x.s)||!serviceCleanupComplete(x.s));
  const dateOf=x=>x.type==='local'?(x.s.date||''):String(x.f.modifiedTime||'').slice(0,10);
  const editedOf=x=>x.type==='local'?(x.s.lastEdited||0):(Date.parse(x.f.modifiedTime)||0);
  const tagOf=x=>x.type==='local'?(x.s.tag||''):'\uffff';
  if(serviceSort==='tag')items.sort((a,b)=>tagOf(a).localeCompare(tagOf(b)));
  else if(serviceSort==='date')items.sort((a,b)=>dateOf(b).localeCompare(dateOf(a)));
  else if(serviceSort==='edited')items.sort((a,b)=>editedOf(b)-editedOf(a));
  const totalItems=items.length;items=items.slice(0,homeShown);

  const list=$('#service-list');
  if(!items.length){
    list.innerHTML=`<div style="padding:28px 12px;text-align:center">
      <p class="note" style="margin:0 0 16px">No services yet.</p>
      <button type="button" id="make-new-service" class="pri" style="padding:12px 22px;font-size:calc(15px * var(--ui-scale))">Make a new service</button>
    </div>`;
  }else{
    list.innerHTML=items.map(x=>{
      if(x.type==='local'){
        const s=x.s,i=x.i;
        return `<div class="service-card">
          ${hcLocalIcon(i)}
          <div class="info">
            ${svcTitleHtml(s,i)}
            <small><button type="button" class="service-date-edit" data-edit-service-date="${i}" aria-label="Change service date, currently ${e(formatServiceDate(s.date))}" title="Change service date">${e(formatServiceDate(s.date))}</button><input type="date" class="service-date-picker" data-service-date-picker="${i}" value="${e(s.date||'')}" aria-label="New service date" tabindex="-1" hidden>${s.lastEdited?' · '+e(formatLastEdited(s.lastEdited)):''}${s.largeMediaRemoved?' · Drive media expired after 14 days':''}</small>
          </div>
          <button class="pri" data-open-service="${i}">Open</button>
          <button data-del-service="${i}" style="color:#c22">Delete</button>
        </div>`;
      }
      const f=x.f;
      return `<div class="service-card">
        ${hcIcon('cloud','In Google Drive only, not downloaded to this computer')}
        <div class="info">
          <div style="display:flex;align-items:center;gap:7px;flex-wrap:wrap"><h3 class="hc-title">${e(f.name.replace(/\.cpres$/i,''))}</h3></div>
          <small>Google Drive only · ${e(fmtDriveDate(f.modifiedTime))}</small>
        </div>
        <button type="button" class="pri" data-home-cloud-open="${e(f.id)}">Download</button>
        <button type="button" style="color:#c22" data-home-cloud-delete="${e(f.id)}">Delete from Drive</button>
      </div>`;
    }).join('');
  }

  const btn=$('#home-cloud-refresh'),st=$('#home-cloud-status');
  (function(){
    const old=document.getElementById('home-see-more-wrap');if(old)old.remove();
    if(totalItems>items.length)list.insertAdjacentHTML('afterend','<div id="home-see-more-wrap" style="text-align:center;margin:14px 0 0"><button type="button" id="home-see-more">See more</button></div>');
  })();
  if(btn){btn.disabled=syncBusy||!cloudUi.connected;btn.textContent=syncBusy?'Syncing…':'↻ Sync'}
  if(st)st.textContent=cloudUi.connected?(cloudUi.online===false?'Offline':(cloudUi.lastSyncAt?'Last synced '+fmtDriveDate(cloudUi.lastSyncAt):'')):'';

  const box=$('#home-cloud-list');
  if(box){
    let h='';
    if(!cloudUi.connected)h='<p class="note">Google Drive isn’t connected, so services are saved on this computer only. <button type="button" data-open-cloud-modal="1">Connect Google Drive</button></p>';
    else if(!cloudUi.folderId)h='<p class="note">Choose a Drive folder to share services between computers. <button type="button" data-open-cloud-modal="1">Choose folder</button></p>';
    else if(cloudUi.error)h='<p style="color:#c22;margin:0">'+e(cloudUi.error)+'</p>';
    box.innerHTML=h;
  }

  try{renderTemplatesHome()}catch(err){console.warn('[templates]',err)}
  try{renderOlderServicesToggle(olderCount)}catch(err){console.warn('[older services]',err)}
  try{refreshTemplateSelect()}catch{}

  if(cloudReady()&&!syncBusy&&Date.now()-cloudUi.lastAttemptAt>LIBRARY_STALE_MS)setTimeout(()=>cloudSyncAll({quiet:true}),50);
};
if(document.body.classList.contains('home-mode'))rHome();

function renderOlderServicesToggle(count){
  let btn=document.getElementById('home-see-older');
  if(!count){if(btn)btn.remove();return}
  const tpl=document.getElementById('home-templates');if(!tpl)return;
  if(!btn){btn=document.createElement('button');btn.type='button';btn.id='home-see-older';btn.style.cssText='display:block;margin:18px auto 0';}
  if(btn.nextElementSibling!==tpl)tpl.insertAdjacentElement('beforebegin',btn);
  btn.textContent=olderServicesVisible?'Hide older services':`Show older services (${count})`;
  btn.setAttribute('aria-expanded',olderServicesVisible?'true':'false');
}

/* ---------- delete from Drive ---------- */
function driveReason(raw){
  const s=String(raw||'');
  const j=s.indexOf('{');
  if(j>=0){
    try{
      const o=JSON.parse(s.slice(j));const er=o.error||{};
      const reason=(er.errors&&er.errors[0]&&er.errors[0].reason)||er.status||'';
      return (er.message||'Drive error')+(reason?' ('+reason+')':'');
    }catch{}
  }
  return s.replace(/^HTTP_ERROR:/,'').slice(0,300);
}
async function cloudDeleteFile(fileId){
  if(!fileId)return;
  const f=(cloudUi&&cloudUi.files||[]).find(x=>x.id===fileId);
  const name=(f&&f.name?f.name:fileId).replace(/\.cpres$/i,'');
  if(!await appConfirm('Permanently delete "'+name+'" from Google Drive?\n\nThis cannot be undone, and other computers will no longer see it.',{title:'Delete from Google Drive',confirmLabel:'Delete',danger:true}))return;
  if(cloudUi&&Array.isArray(cloudUi.files)){
    cloudUi.files=cloudUi.files.filter(x=>x.id!==fileId);
  }
  rHome();
  toast('Deleted "'+name+'" from Drive.');
  try{
    const res=await cloudInvoke('gdrive_delete_file',{fileId});
    if(!res||res.ok===false){
      const raw=(res&&res.error)||'Unknown error';
      console.warn('[cloud] delete failed',raw);
    }
  }catch(err){
    console.warn('[cloud] delete error',err);
  }
}
document.addEventListener('click',ev=>{
  const b=ev.target.closest&&ev.target.closest('[data-home-cloud-delete]');
  if(b){
    ev.preventDefault();
    ev.stopPropagation();
    cloudDeleteFile(b.dataset.homeCloudDelete||b.getAttribute('data-home-cloud-delete'));
  }
});

}


/* ===== Only one context menu open at a time ===== */
if(!IS_AUDIENCE){
  const MENU_SEL='#outline-menu,#item-menu,#slide-menu,#canvas-ctx,#add-menu,#svc-ctx,#tag-menu,#inline-add-popover';
  function closeAllMenus(){
    document.querySelectorAll(MENU_SEL).forEach(m=>{m.classList.remove('open');m.hidden=true});
    const tm=document.getElementById('text-ctx');
    if(tm)tm.classList.remove('open');
  }
  // Capture phase: runs before any menu's own handler, so a right click
  // closes the old menu first and then the new one opens as normal.
  document.addEventListener('pointerdown',ev=>{
    if(ev.button!==0)return;
    const t=ev.target;
    if(!t||!t.closest)return;
    if(t.closest(MENU_SEL+',#text-ctx'))return;   // clicking inside a menu is fine
    if(t.closest('#outline-add-more'))return;     // that button toggles its own menu
    closeAllMenus();
  },true);
  document.addEventListener('contextmenu',closeAllMenus,true);
}


/* =====================================================================
   UPDATE BLOCK
   ===================================================================== */
/* ---- Inline Add slide button at the end of the strip ---- */
(function injectInlineAddStyles(){
  const s=document.createElement('style');
  s.textContent=`
.strip-inline-add{display:flex;justify-content:center;padding:12px 8px 18px;margin-top:8px;}
.strip-inline-add-btn{display:inline-flex;align-items:center;gap:6px;padding:6px 16px;border-radius:999px;font-size:12px;font-weight:600;background:var(--pn);border:1px dashed var(--ln);color:var(--mu);cursor:pointer;transition:all .15s ease;}
.strip-inline-add-btn:hover{border-color:var(--ac);color:var(--ac);background:var(--hl);border-style:solid;}
#strip .strip-add{display:none!important;}
#strip .thumb-wrap{position:relative}
#strip .th-add{position:absolute;right:8px;bottom:26px;z-index:5;display:none;padding:3px 9px;font-size:11px;font-weight:600;border-radius:999px;background:var(--pn);border:1px solid var(--ln);color:var(--ink);box-shadow:0 2px 8px rgba(0,0,0,.3);cursor:pointer}
#strip .thumb-wrap:hover .th-add,#strip .th-add.th-add-open,#strip .th-add:focus-visible{display:block}
#strip .th-add:hover{border-color:var(--ac);color:var(--ac)}
body[data-pointer-dragging="1"] #strip .th-add{display:none!important}
#inline-add-popover{min-width:160px;z-index:1500;}
`;
  document.head.appendChild(s);
})();

function placeStripAdd(){
  const st=$('#strip');if(!st||!S.length)return;
  const oldBar=st.querySelector('.strip-add');if(oldBar)oldBar.remove();

  let inlineWrap=st.querySelector('.strip-inline-add');
  if(!inlineWrap){
    inlineWrap=document.createElement('div');
    inlineWrap.className='strip-inline-add';
    st.appendChild(inlineWrap);
  }
  inlineWrap.innerHTML='<button type="button" id="strip-inline-add-btn" class="strip-inline-add-btn" title="Add slide" aria-label="Add slide">'
    +'<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>'
    +'<span>Add slide</span></button>';
  st.appendChild(inlineWrap);
}

let inlineAddCtx={mode:'end'},inlineAddAnchor=null,pendingSectionInsert=null;

/** Where does a new slide go? {i,after} = into section i after slide `after`; {newSectionAt} = make a new section. */
function resolveAddTarget(ctx){
  if(ctx&&ctx.mode==='after'){
    const it=S[ctx.i];
    if(it&&!isSongItem(it)&&!isPptHandoffItem(it))return {i:ctx.i,after:ctx.j};
    return {newSectionAt:Math.min(S.length,ctx.i+1)};
  }
  if(!S.length)return {newSectionAt:0};
  const last=S[S.length-1];
  if(isSongItem(last)||isPptHandoffItem(last))return {newSectionAt:S.length};
  return {i:S.length-1,after:last.ord.length-1};
}
/** Index for things that are always their own section (songs, PowerPoints). */
function sectionInsertIndex(ctx){return ctx&&ctx.mode==='after'?Math.min(S.length,ctx.i+1):S.length}

function addBlankSlideAt(ctx){
  const tg=resolveAddTarget(ctx);
  if(tg.newSectionAt!=null){
    pushUndo();
    const it=mkS(nextSectionName(),['Type your text here'],0,6.5);
    S.splice(tg.newSectionAt,0,it);
    sel=tg.newSectionAt;A=it;si=0;aj=0;tab='Structure';
    rAll();setTimeout(()=>scrollNewItemToTop(sel),0);toast('Added new section.');
    return;
  }
  addNewSlideInItem(tg.i,tg.after);
}

function renderInlineAddMain(pop){
  let html='<button type="button" data-inline-add="blank">Blank slide</button><button type="button" data-inline-add="other">Other slide\u2026</button>';
  if(inlineAddCtx.mode==='end'&&isSongItem(A))html+='<button type="button" data-inline-add="struct">Edit song structure</button>';
  pop.innerHTML=html;
}
function showInlineAddPopover(btn,ctx){
  let pop=document.getElementById('inline-add-popover');
  if(!pop){
    pop=document.createElement('div');
    pop.id='inline-add-popover';pop.className='outline-menu';pop.setAttribute('role','menu');
    document.body.appendChild(pop);
  }
  if(pop.classList.contains('open')&&inlineAddAnchor===btn){hideInlineAddPopover();return}
  hideInlineAddPopover();
  inlineAddCtx=ctx||{mode:'end'};inlineAddAnchor=btn;
  btn.classList.add('th-add-open');
  renderInlineAddMain(pop);
  pop.hidden=false;pop.classList.add('open');
  const r=btn.getBoundingClientRect(),w=pop.offsetWidth||160,h=pop.offsetHeight||90;
  pop.style.left=Math.max(6,Math.min(r.left,window.innerWidth-w-8))+'px';
  const below=r.bottom+4;
  pop.style.top=(below+h<window.innerHeight-8?below:Math.max(6,r.top-h-4))+'px';
}
function hideInlineAddPopover(){
  const pop=document.getElementById('inline-add-popover');
  if(pop){pop.hidden=true;pop.classList.remove('open')}
  if(inlineAddAnchor)inlineAddAnchor.classList.remove('th-add-open');
  inlineAddAnchor=null;
}

document.addEventListener('click',ev=>{
  const t=ev.target;if(!t||!t.closest)return;
  const hoverBtn=t.closest('.th-add');
  if(hoverBtn){
    ev.preventDefault();ev.stopPropagation();
    showInlineAddPopover(hoverBtn,{mode:'after',i:+hoverBtn.dataset.addAfterI,j:+hoverBtn.dataset.addAfterJ});
    return;
  }
  const addBtn=t.closest('#strip-inline-add-btn');
  if(addBtn){
    ev.preventDefault();ev.stopPropagation();
    showInlineAddPopover(addBtn,{mode:'end'});
    return;
  }
  const popItem=t.closest('#inline-add-popover [data-inline-add]');
  if(popItem){
    const act=popItem.dataset.inlineAdd,ctx=inlineAddCtx,pop=document.getElementById('inline-add-popover');
    if(act==='other'){
      pop.innerHTML='<div style="padding:6px 10px;font-size:11px;font-weight:700;color:var(--mu);text-transform:uppercase">Add slide</div>'
        +'<button type="button" data-inline-add="logo">Church Logo</button>'
        +'<button type="button" data-inline-add="song">Song</button>'
        +'<button type="button" data-inline-add="ppt">PowerPoint</button>'
        +'<button type="button" data-inline-add="back" style="color:var(--mu);border-top:1px solid var(--ln);margin-top:4px">\u2190 Back</button>';
      return;
    }
    if(act==='back'){renderInlineAddMain(pop);return}
    hideInlineAddPopover();
    if(act==='blank'){addBlankSlideAt(ctx);return}
    if(act==='logo'){addChurchLogoSection(resolveAddTarget(ctx));return}
    if(act==='song'){openLibraryModal(sectionInsertIndex(ctx));return}
    if(act==='ppt'){
      const f=$('#import-slides-file');
      if(f){
        pendingSectionInsert=sectionInsertIndex(ctx);
        f.value='';
        f.addEventListener('cancel',()=>{pendingSectionInsert=null},{once:true});
        f.click();
      }
      return;
    }
    if(act==='struct'){openStruct();return}
    return;
  }
  if(!t.closest('#inline-add-popover'))hideInlineAddPopover();
});

/* ---- Videos behave like images: free-form element with move / resize / rotate ---- */
function renderVideoEl(el,editable,selected,stage){
  const base=`left:${el.x}%;top:${el.y}%;width:${el.w}%;height:${el.h}%;transform:rotate(${el.rot||0}deg);z-index:${el.z||1}`;
  const handles=editable&&selected?elHandlesHtml(el):'';
  const vs=el.settings||{};
  const play=!!stage&&!editable;
  const raw=mediaURLs[el.mediaId]||'';
  // Editor: seek to first frame (#t=0.1). Present: full src so play works.
  const src=raw?(play?raw:raw+'#t=0.1'):'';
  const vol=vs.volume!=null?vs.volume:1;
  // Native controls when editing (selected) or presenting without autoplay so the user can press play
  const showControls=(editable&&selected)||(play&&vs.autoplay===false);
  const peVideo=showControls?'auto':'none';
  return `<div class="el vid-el${selected?' selected':''}" data-el="${e(el.id)}" style="${base};pointer-events:${editable||showControls?'auto':'none'}">${handles}<video class="el-video" src="${e(src)}" ${play&&vs.autoplay!==false?'autoplay':''} ${play&&vs.loop?'loop':''} ${vs.muted===true?'muted':''} data-vol="${vol}" playsinline ${showControls?'controls':''} preload="metadata" style="width:100%;height:100%;object-fit:${el.fit==='fill'?'fill':'contain'};pointer-events:${peVideo};display:block;background:#000"></video></div>`;
}
function addVideoElToCurrentSlide(mediaId,name){
  if(!A||isSongItem(A)||live){toast('Videos can only be added to regular (non-song) slides.');return}
  const part=sls(A)[aj];if(!part)return;
  const els=ensureElements(part,A);
  const el={id:uid(),kind:'video',mediaId,name:name||'Video',x:25,y:25,w:50,h:50,rot:0,
    z:els.reduce((m,x)=>Math.max(m,x.z||1),0)+1,
    settings:{autoplay:true,loop:false,muted:false,volume:1}};
  els.push(el);elSel=el.id;tab='Structure';
  saveSoon();rStage();updateCurrentThumb();rPanel();
  toast('Video added. Drag to move, corners to resize, top handle to rotate.');
}
function decorateVideoPanel(){
  if(!A)return;
  const body=document.querySelector('#panel .panel-body');if(!body)return;
  const part=sls(A)[aj];
  document.querySelectorAll('#panel .el-row').forEach(r=>{
    const el=((part&&part.elements)||[]).find(x=>x.id===r.dataset.elId);
    if(el&&el.kind==='video'){const lab=r.querySelector('.el-label');if(lab)lab.textContent='\uD83C\uDFAC '+(el.name||'Video')}
  });
  const el=getSelectedEl();
  if(!el||el.kind!=='video'||tab!=='Structure')return;
  const vs=el.settings||{};
  const vol=Math.round((vs.volume!=null?vs.volume:1)*100);
  body.insertAdjacentHTML('afterbegin',`<h3>Video: ${e(el.name||'Video')}</h3>
    <label style="display:flex;align-items:center;gap:8px"><input type="checkbox" id="elv-autoplay" ${vs.autoplay!==false?'checked':''}> Auto play</label>
    <label style="display:flex;align-items:center;gap:8px"><input type="checkbox" id="elv-loop" ${vs.loop?'checked':''}> Loop</label>
    <label style="display:flex;align-items:center;gap:8px"><input type="checkbox" id="elv-muted" ${vs.muted===true?'checked':''}> Mute video</label>
    <label>Volume <span id="elv-vol-val" style="color:var(--mu)">${vol}</span>%<input type="range" id="elv-volume" min="0" max="100" value="${vol}"></label>`);
}
/* ---- Drag & drop files onto the slide (same result as copy / paste) ---- */
function dropHasFiles(ev){const t=ev.dataTransfer&&ev.dataTransfer.types;return !!t&&[...t].includes('Files')}
async function addDroppedMediaToSlide(files,pt){
  if(live||!A||!S.length)return;
  if(isSongItem(A)){toast('Song slides cannot hold loose images or videos. Use "Add a background", or drop onto a regular slide.');return}
  const part=sls(A)[aj];if(!part)return;
  const els=ensureElements(part,A);
  let n=0;
  pushUndo();
  for(const file of files){
    const isVid=/^video\//.test(file.type),isImg=/^image\//.test(file.type);
    if(!isVid&&!isImg)continue;
    const id='m'+Date.now().toString(36)+Math.random().toString(36).slice(2,8);
    await mPut(id,file);
    mediaURLs[id]=URL.createObjectURL(file);
    mediaMeta.push({id,name:file.name||(isVid?'Dropped video':'Dropped image'),kind:isVid?'video':'image',size:file.size,createdAt:Date.now(),source:'drop'});
    const w=isVid?50:40,h=isVid?50:40,off=n*3;
    const x=pt?Math.max(0,Math.min(100-w,pt.x-w/2+off)):25+off;
    const y=pt?Math.max(0,Math.min(100-h,pt.y-h/2+off)):25+off;
    const z=els.reduce((m,x2)=>Math.max(m,x2.z||1),0)+1;
    const el=isVid
      ?{id:uid(),kind:'video',mediaId:id,name:file.name||'Video',x,y,w,h,rot:0,z,settings:{autoplay:true,loop:false,muted:false,volume:1}}
      :{id:uid(),kind:'image',src:mediaURLs[id],mediaId:id,x,y,w,h,rot:0,z,fit:'contain'};
    els.push(el);elSel=el.id;n++;
  }
  if(!n){toast('Drop an image or video file.');return}
  saveMediaMeta();saveSoon();tab='Structure';
  rStage();updateCurrentThumb();rPanel();
  toast(n===1?'Added to slide.':'Added '+n+' items to slide.');
}
/* ---- Multiple church logos (per-service choice) ---- */
function effectiveLogoId(){
  const svc=services[activeService];
  const ids=churchLogos.map(l=>l.id);
  if(svc&&svc.logoId&&ids.includes(svc.logoId))return svc.logoId;
  if(defaultLogoId&&ids.includes(defaultLogoId))return defaultLogoId;
  return ids[0]||'';
}
function syncEffectiveLogo(){
  const l=churchLogos.find(x=>x.id===effectiveLogoId());
  const src=l?l.src:'';
  if(src!==churchLogo){
    churchLogo=src;
    try{src?localStorage.setItem('church-presenter-logo',src):localStorage.removeItem('church-presenter-logo')}catch{}
  }
}
function saveLogos(){
  try{
    localStorage.setItem('church-presenter-logos',JSON.stringify(churchLogos));
    localStorage.setItem('church-presenter-logo-default',defaultLogoId||'');
  }catch(err){toast('That logo is too large to store. Try a smaller image.')}
}
function renderLogoPicker(){
  const box=$('#logo-picker');if(!box)return;
  if(!churchLogos.length){
    box.innerHTML='<p class="note" style="margin:12px 0 0">No logos yet. Upload one below. Upload more than one to pick a different logo per service.</p>';
    return;
  }
  const cur=effectiveLogoId();
  box.innerHTML=`<label style="margin-top:14px">Logo for this service
    <select id="logo-select">${churchLogos.map(l=>`<option value="${e(l.id)}" ${l.id===cur?'selected':''}>${e(l.name)}${l.id===defaultLogoId?' (default)':''}</option>`).join('')}</select></label>
    <div class="logo-picker-actions"><button type="button" id="logo-rename">Rename</button><button type="button" id="logo-set-default">Set as default</button></div>
    <p class="note" style="margin:6px 0 0">Services without their own choice use the default logo.</p>`;
}
async function removeSelectedLogo(){
  const id=effectiveLogoId();
  const l=churchLogos.find(x=>x.id===id);if(!l)return;
  if(!await appConfirm('Remove the logo "'+l.name+'"? Services using it will fall back to the default logo.',{title:'Remove logo',confirmLabel:'Remove',danger:true}))return;
  churchLogos=churchLogos.filter(x=>x.id!==id);
  if(defaultLogoId===id)defaultLogoId=churchLogos[0]?churchLogos[0].id:'';
  services.forEach(s=>{if(s.logoId===id)delete s.logoId});
  saveLogos();saveState();syncEffectiveLogo();updateLogoPreview();rStage();
  try{pushAudienceClear()}catch{}
  toast('Logo removed.');
}
uploadLogoFile=function(file){
  if(!file)return;
  const reader=new FileReader();
  reader.onload=()=>{
    const name=(file.name||'').replace(/\.[^.]+$/,'')||('Logo '+(churchLogos.length+1));
    const id='logo-'+Date.now().toString(36);
    churchLogos.push({id,name,src:reader.result});
    if(!defaultLogoId)defaultLogoId=id;
    const svc=services[activeService];if(svc)svc.logoId=id;
    saveLogos();saveState();syncEffectiveLogo();updateLogoPreview();rStage();
    try{pushAudienceClear()}catch{}
    toast('Logo "'+name+'" added and selected for this service.');
  };
  reader.readAsDataURL(file);
};
{
  const _ulp=updateLogoPreview;
  updateLogoPreview=function(){_ulp.apply(this,arguments);try{renderLogoPicker()}catch{}};
}
/* ---- Credits transparency ---- */
function syncCreditsTranspUi(){
  const w=$('#song-credits-transp-wrap');if(w)w.hidden=!songSettings.credits;
  const v=songSettings.creditsTransparency!=null?songSettings.creditsTransparency:22;
  const s=$('#song-credits-transp');if(s)s.value=String(v);
  const l=$('#song-credits-transp-val');if(l)l.textContent=v+'%';
}
/* ---- Auto-fit: enlarged lyrics that wrap past the line limit spill onto extra slides ----
   The "base" structure (sections + lines per slide) is never lost. Extra slides are derived
   chunks flagged auto:1 that point back at their base part (of). Shrinking the font removes them. */
let _fitProbe=null,_fitRaf=0,_fitFontsP=null;
function ensureFitFonts(){
  if(_fitFontsP)return _fitFontsP;
  try{_fitFontsP=Promise.all(['400 20px Figtree','700 20px Figtree','400 20px Lora','700 20px Lora'].map(f=>document.fonts.load(f))).catch(()=>{})}
  catch{_fitFontsP=Promise.resolve()}
  return _fitFontsP;
}
/** Display-line count of every source line when rendered with the slide's real text box geometry. */
function fitMeasure(f,lines){
  const W=1000;
  if(!_fitProbe){
    _fitProbe=document.createElement('div');
    _fitProbe.setAttribute('aria-hidden','true');
    _fitProbe.style.cssText='position:fixed;left:-99999px;top:0;visibility:hidden;pointer-events:none;box-sizing:border-box;white-space:pre-line;line-height:1.28;margin:0;padding:0';
    document.body.appendChild(_fitProbe);
  }
  const fsPx=(+f.fs||5.2)*W/100;
  const ff=f.ff==='serif'?"'Lora',Georgia,'Times New Roman',serif":"'Figtree','Segoe UI',system-ui,sans-serif";
  const st=_fitProbe.style;
  const margin=Math.max(0,Math.min(25,Number(f.mb??8)));
  st.width=(W*(1-(2*margin/100)))+'px';
  st.fontSize=fsPx+'px';
  st.fontFamily=ff;
  st.fontWeight=f.b?'700':'400';
  st.textTransform=f.tc==='upper'?'uppercase':f.tc==='title'?'capitalize':'none';
  _fitProbe.innerHTML=lines.map(l=>'<div>'+(e(l)||'&nbsp;')+'</div>').join('');
  const lh=fsPx*1.28;
  return [..._fitProbe.children].map(d=>Math.max(1,Math.round(d.getBoundingClientRect().height/lh)));
}
/** Max display lines per slide: the lines-per-slide setting, never more than physically fits vertically. */
function fitLimit(it){
  const fs=+((it.f||{}).fs)||5.2;
  const cap=Math.max(1,Math.floor(44.25/(fs*1.28)));   // (56.25 - 2*6) / line height, in slide-width units
  const n=+curLps(it)||0;
  return Math.max(1,n>0?Math.min(n,cap):cap);
}
function fitSplit(lines,counts,limit){
  const chunks=[];let cur=[],used=0;
  lines.forEach((ln,i)=>{
    const c=counts[i]||1;
    if(cur.length&&used+c>limit){chunks.push(cur);cur=[];used=0}
    cur.push(ln);used+=c;
  });
  if(cur.length)chunks.push(cur);
  return chunks.map(c=>c.join('\n'));
}
function fitPosKey(it,j){
  let occ=-1,start=0,prevOf=null,prevK=-1;
  const lim=Math.min(j,it.ord.length-1);
  for(let x=0;x<=lim;x++){
    const p=it.parts.find(q=>q.id==it.ord[x]);
    if(p&&p.auto){
      if(p.of!==prevOf||p.k<=prevK){occ++;start=x}
      prevOf=p.of;prevK=p.k;
    }else{occ++;start=x;prevOf=null;prevK=-1}
  }
  return {occ:Math.max(0,occ),off:Math.max(0,j-start)};
}
/** Back to the pure base structure (what the structure builder edits). */
function collapseFit(it){
  if(!it||!Array.isArray(it.ord)||!it.parts.some(p=>p.auto))return;
  const keyAj=it===A?fitPosKey(it,aj):null;
  const keySi=it===S[sel]?fitPosKey(it,si):null;
  const ids=new Set(it.parts.map(p=>p.id));
  const groups=new Map();
  it.parts.filter(p=>p.auto).sort((a,b)=>(a.k||0)-(b.k||0)).forEach(p=>{
    if(!groups.has(p.of))groups.set(p.of,[]);
    groups.get(p.of).push(p);
  });
  groups.forEach((cs,of)=>{          // rebuild a base that went missing (e.g. after a copy)
    if(ids.has(of))return;
    const c0=cs[0];
    const nb={...c0,id:of,n:String(c0.n).replace(/\s*\[\d+\/\d+\]$/,''),tx:cs.map(c=>c.tx).join('\n')};
    delete nb.auto;delete nb.of;delete nb.k;delete nb.fitHid;
    it.parts.push(nb);
  });
  const out=[];let prevOf=null,prevK=-1;
  for(const id of it.ord){
    const p=it.parts.find(x=>x.id==id);
    if(p&&p.auto){
      if(p.of!==prevOf||p.k<=prevK)out.push(p.of);
      prevOf=p.of;prevK=p.k;
    }else{out.push(id);prevOf=null;prevK=-1}
  }
  it.parts=it.parts.filter(p=>!p.auto);
  it.parts.forEach(p=>{delete p.fitHid});
  it.ord=out;
  if(keyAj)aj=Math.max(0,Math.min(keyAj.occ,out.length-1));
  if(keySi)si=Math.max(0,Math.min(keySi.occ,out.length-1));
  it._fitPos={aj:keyAj,si:keySi};
}
/** Split base parts into extra slides wherever they exceed the display-line limit at the current font. */
function expandFit(it){
  if(!it||!isSongItem(it)||!Array.isArray(it.ord))return;
  const pos=it._fitPos;delete it._fitPos;
  const f=it.f||{};
  const limit=fitLimit(it);
  const bases=[...new Set(it.ord)].map(id=>it.parts.find(p=>p.id==id))
    .filter(p=>p&&!p.blank&&!isAutoBlankPart(p)&&String(p.tx||'').trim());
  const chunkIds=new Map(),newParts=[];
  bases.forEach(p=>{
    const lines=String(p.tx).split('\n');
    const counts=fitMeasure(f,lines);
    if(counts.reduce((a,b)=>a+b,0)<=limit)return;
    const texts=fitSplit(lines,counts,limit);
    if(texts.length<2)return;
    const idsOut=texts.map((tx,k)=>{
      const c={...p,id:uid(),n:`${p.n} [${k+1}/${texts.length}]`,tx,auto:1,of:p.id,k};
      newParts.push(c);return c.id;
    });
    chunkIds.set(p.id,idsOut);
  });
  const out=[],starts=[],lens=[];
  it.ord.forEach(id=>{
    starts.push(out.length);
    const c=chunkIds.get(id);
    if(c){out.push(...c);lens.push(c.length)}else{out.push(id);lens.push(1)}
  });
  chunkIds.forEach((_,id)=>{const b=it.parts.find(p=>p.id==id);if(b)b.fitHid=1});
  it.parts.push(...newParts);
  it.ord=out;
  const map=occ=>{
    if(!starts.length)return 0;
    const o=Math.max(0,Math.min(occ,starts.length-1));
    return starts[o];
  };
  const mapOff=(occ,off)=>{
    if(!starts.length)return 0;
    const o=Math.max(0,Math.min(occ,starts.length-1));
    return starts[o]+Math.max(0,Math.min(off||0,lens[o]-1));
  };
  if(it===A)aj=mapOff(aj,pos&&pos.aj?pos.aj.off:0);
  if(it===S[sel])si=mapOff(si,pos&&pos.si?pos.si.off:0);
  void map;
}
function refitSong(it){
  if(!it||!isSongItem(it))return;
  collapseFit(it);
  expandFit(it);
}
/** Edits made on an auto-split slide are written back to its base part. */
function syncChunkEdit(it,part){
  if(!it||!part||!part.auto)return;
  const base=it.parts.find(p=>p.id===part.of);if(!base)return;
  base.tx=it.parts.filter(p=>p.auto&&p.of===part.of).sort((a,b)=>(a.k||0)-(b.k||0)).map(c=>c.tx).join('\n');
}
/** Reflow lines-per-slide while keeping blank / empty sections in place. */
function reflowParts(it,lps){
  const parts=[],ord=[],seen={};
  const use=parsed=>{
    parsed.ord.forEach(id=>{
      const p=parsed.parts.find(x=>x.id==id);if(!p)return;
      const key=p.n+'|'+p.tx;
      if(!seen[key]){const np={...p,id:uid()};parts.push(np);seen[key]=np.id}
      ord.push(seen[key]);
    });
  };
  let run=[];
  const flush=()=>{
    if(!run.length)return;
    use(parse(songToText({ord:run,parts:it.parts},true),lps));
    run=[];
  };
  it.ord.forEach(id=>{
    const p=it.parts.find(x=>x.id==id);if(!p)return;
    if(p.blank||isAutoBlankPart(p)||!String(p.tx||'').trim()){
      flush();
      if(!seen['\0'+p.id]){parts.push(p);seen['\0'+p.id]=1}
      ord.push(p.id);
    }else run.push(id);
  });
  flush();
  return {parts,ord};
}
function scheduleSongFit(it){
  cancelAnimationFrame(_fitRaf);
  _fitRaf=requestAnimationFrame(async()=>{
    try{await ensureFitFonts()}catch{}
    if(!S.includes(it))return;
    refitSong(it);
    rStrip();rStage();saveSoon();
    const ta=$('#pt');
    if(ta&&S[sel]===it){
      const cp=sls(it)[si];
      if(cp){
        if(document.activeElement!==ta)ta.value=cp.tx;
        const hd=ta.previousElementSibling;
        if(hd&&hd.tagName==='H3')hd.textContent='Song lyrics: '+cp.n;
      }
    }
  });
}
/* ---- PowerPoint + external display: free the external screen while the slideshow runs ---- */
async function releaseAudienceForPpt(){
  if(!live||presentLocalOnly)return;
  if(!(audienceIsTauri||(audienceWin&&!audienceWin.closed)))return;
  pptAudienceClosed=true;
  try{await closeAudienceWindow()}catch(err){console.warn('[PPT] close audience',err)}
  await new Promise(r=>setTimeout(r,200));
}
async function restoreAudienceAfterPpt(){
  if(!pptAudienceClosed)return;
  pptAudienceClosed=false;
  if(!live||presentLocalOnly)return;
  try{
    await openAudienceOnDisplay(selectedDisplay||($('#display-setting')&&$('#display-setting').value)||'');
    await pushAudienceStage();
    await refocusPresenterWindow();
  }catch(err){console.warn('[PPT] restore audience',err)}
}
/* ---- listeners ---- */
if(!IS_AUDIENCE){
  ensureFitFonts();
  // video element settings
  document.addEventListener('change',ev=>{
    const t=ev.target;
    if(!t||!t.id||t.id.indexOf('elv-')!==0)return;
    const el=getSelectedEl();if(!el||el.kind!=='video')return;
    const vs=el.settings||(el.settings={});
    if(t.id==='elv-autoplay')vs.autoplay=!!t.checked;
    else if(t.id==='elv-loop')vs.loop=!!t.checked;
    else if(t.id==='elv-muted')vs.muted=!!t.checked;
    else if(t.id==='elv-volume')vs.volume=(+t.value)/100;
    saveSoon();rStage();updateCurrentThumb();
  });
  document.addEventListener('input',ev=>{
    const t=ev.target;
    if(t&&t.id==='elv-volume'){
      const el=getSelectedEl();if(!el||el.kind!=='video')return;
      (el.settings||(el.settings={})).volume=(+t.value)/100;
      const lab=$('#elv-vol-val');if(lab)lab.textContent=t.value;
      saveSoon();
    }
    if(t&&t.id==='song-credits-transp'){
      songSettings.creditsTransparency=+t.value;saveSongSettings();syncCreditsTranspUi();
      try{rStrip();rStage()}catch{}
    }
  });
  // drag & drop (browser HTML5 + Tauri native — OS drops never reach the webview in Tauri)
  const setHot=on=>{const s=$('#stage');if(s)s.classList.toggle('drop-hot',!!on)};
  function stageDropPoint(clientX,clientY){
    const sle=document.querySelector('#stage .sl');
    if(!sle)return null;
    const r=sle.getBoundingClientRect();
    return {x:(clientX-r.left)/r.width*100,y:(clientY-r.top)/r.height*100};
  }
  function canDropOnStage(){
    return !live&&!document.body.classList.contains('home-mode')&&!!A&&!isSongItem(A);
  }
  ['dragenter','dragover'].forEach(type=>document.addEventListener(type,ev=>{
    if(!dropHasFiles(ev))return;
    ev.preventDefault();
    const over=!!(ev.target.closest&&ev.target.closest('#stage'))&&canDropOnStage();
    try{ev.dataTransfer.dropEffect=over?'copy':'none'}catch{}
    setHot(over);
  }));
  document.addEventListener('dragleave',ev=>{if(!ev.relatedTarget)setHot(false)});
  document.addEventListener('drop',ev=>{
    if(!dropHasFiles(ev))return;
    ev.preventDefault();setHot(false);
    if(!canDropOnStage()){
      if(isSongItem(A))toast('Song slides cannot hold loose images or videos. Drop onto a regular slide.');
      return;
    }
    if(!(ev.target.closest&&ev.target.closest('#stage'))){toast('Drop the file onto the slide in the middle of the screen.');return}
    addDroppedMediaToSlide([...ev.dataTransfer.files],stageDropPoint(ev.clientX,ev.clientY));
  });
  // Tauri: native OS file drops (HTML5 drop never fires inside the webview)
  (async()=>{
    if(!isTauri||IS_AUDIENCE)return;
    try{
      let unlisten=null;
      // Prefer webview API (Tauri 2); fall back to window API (Tauri 1)
      try{
        const {getCurrentWebviewWindow}=await import('@tauri-apps/api/webviewWindow');
        const w=getCurrentWebviewWindow();
        unlisten=await w.onDragDropEvent(async ev=>{
          const p=ev.payload||{};
          if(p.type==='enter'||p.type==='over'){setHot(canDropOnStage());return}
          if(p.type==='leave'){setHot(false);return}
          if(p.type!=='drop')return;
          setHot(false);
          if(!canDropOnStage()){toast(isSongItem(A)?'Song slides cannot hold loose media. Drop onto a regular slide.':'Nothing to drop onto.');return}
          const paths=p.paths||[];
          if(!paths.length)return;
          const files=[];
          for(const path of paths){
            try{
              const url=convertFileSrc(path);
              const res=await fetch(url);
              if(!res.ok)throw new Error('Could not read '+path);
              const blob=await res.blob();
              const name=String(path).split(/[/\\\\]/).pop()||'file';
              // Infer type from extension when blob.type is empty
              let type=blob.type;
              if(!type){
                const ext=(name.split('.').pop()||'').toLowerCase();
                const map={png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',webp:'image/webp',svg:'image/svg+xml',mp4:'video/mp4',webm:'video/webm',mov:'video/quicktime',m4v:'video/mp4',avi:'video/x-msvideo'};
                type=map[ext]||'';
              }
              files.push(new File([blob],name,{type}));
            }catch(err){console.warn('[drop]',path,err);toast('Could not read dropped file: '+(String(path).split(/[/\\\\]/).pop()||'file'))}
          }
          if(files.length)addDroppedMediaToSlide(files,null);
        });
      }catch(e1){
        // Tauri 1 style
        try{
          const win=getCurrentWindow();
          if(win&&win.onFileDropEvent){
            unlisten=await win.onFileDropEvent(async ev=>{
              if(ev.type==='hover'){setHot(canDropOnStage());return}
              if(ev.type==='cancel'){setHot(false);return}
              if(ev.type!=='drop')return;
              setHot(false);
              if(!canDropOnStage())return;
              const paths=ev.paths||[];
              const files=[];
              for(const path of paths){
                try{
                  const url=convertFileSrc(path);
                  const res=await fetch(url);
                  const blob=await res.blob();
                  const name=String(path).split(/[/\\\\]/).pop()||'file';
                  let type=blob.type;
                  if(!type){
                    const ext=(name.split('.').pop()||'').toLowerCase();
                    const map={png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',webp:'image/webp',mp4:'video/mp4',webm:'video/webm',mov:'video/quicktime'};
                    type=map[ext]||'';
                  }
                  files.push(new File([blob],name,{type}));
                }catch(err){console.warn('[drop]',path,err)}
              }
              if(files.length)addDroppedMediaToSlide(files,null);
            });
          }
        }catch(e2){console.warn('[drop] Tauri file-drop not available',e1,e2)}
      }
    }catch(err){console.warn('[drop] setup',err)}
  })();
  // logos
  document.addEventListener('change',ev=>{
    const t=ev.target;
    if(!t||t.id!=='logo-select')return;
    const svc=services[activeService];if(svc)svc.logoId=t.value;
    saveState();syncEffectiveLogo();updateLogoPreview();rStage();
    try{pushAudienceClear()}catch{}
  });
  document.addEventListener('click',async ev=>{
    const t=ev.target;if(!t||!t.closest)return;
    if(t.closest('#logo-rename')){
      const l=churchLogos.find(x=>x.id===effectiveLogoId());if(!l)return;
      const n=await appPrompt('Rename logo','Enter a name for this logo.',l.name);if(n==null||!n.trim())return;
      l.name=n.trim();saveLogos();renderLogoPicker();
    }else if(t.closest('#logo-set-default')){
      defaultLogoId=effectiveLogoId();saveLogos();renderLogoPicker();toast('Default logo updated.');
    }
  });
  // shortcuts window
  document.addEventListener('click',ev=>{
    const t=ev.target;if(!t||!t.closest)return;
    const m=$('#shortcuts-modal');if(!m)return;
    if(t.closest('#shortcuts-btn')){m.hidden=!m.hidden;return}
    if(t.id==='shortcuts-close'||t.id==='shortcuts-modal')m.hidden=true;
  });
  document.addEventListener('keydown',ev=>{
    const m=$('#shortcuts-modal');if(!m)return;
    if(!m.hidden&&ev.key==='Escape'){ev.preventDefault();m.hidden=true;return}
    if(ev.key==='?'&&m.hidden&&!isTypingTarget(ev.target)&&!isPresentingUI()&&!isEditorOverlayOpen()&&!document.body.classList.contains('home-mode')){
      ev.preventDefault();m.hidden=false;
    }
  });
  // structure builder: green Done
  document.addEventListener('click',ev=>{
    const b=ev.target.closest&&ev.target.closest('#sb-done');if(!b)return;
    sbEditing=false;
    try{if(document.activeElement)document.activeElement.blur()}catch{}
    rStrip();rStage();saveSoon();rStruct();
    toast('Section saved.');
  });
  // font family / lines-per-slide edits via the lyrics textarea: refit when the box loses focus
  document.addEventListener('change',ev=>{
    if(ev.target&&ev.target.id==='pt'){
      const it=S[sel];if(it&&isSongItem(it))scheduleSongFit(it);
    }
  });
  // CCLI lookup
  document.addEventListener('click',ev=>{
    const b=ev.target.closest&&ev.target.closest('[data-ccli-lookup]');if(!b)return;
    const title=(($('#edit-song-title-input')&&$('#edit-song-title-input').value)||'').trim();
    const artist=(($('#edit-song-artist-input')&&$('#edit-song-artist-input').value)||'').trim();
    const query=(title+(artist?' '+artist:'')).trim();
    openExternalUrl('https://songselect.ccli.com/Search/Results?SearchText='+encodeURIComponent(query))
      .catch(err=>toast((err&&err.message)||'Could not open SongSelect.'));
    toast('Opening SongSelect. Copy the CCLI Song # into the box.');
  });
  // Present from the current slide
  document.addEventListener('click',ev=>{
    const b=ev.target.closest&&ev.target.closest('#live-from-current');if(!b)return;
    if(isPresentingUI())return;
    if(!A||!S.length){toast('Nothing to present.');return}
    const id=A.ord[aj];let rank=0;
    for(let x=0;x<aj;x++)if(A.ord[x]===id)rank++;
    presentFrom={item:A,id,rank};
    startPresent().catch(err=>console.error(err));
  });
}
/* ===== Sync images and slide decks (as PDF) through Drive ===== */
if(!IS_AUDIENCE){

const MEDIA_MAX=60*1024*1024;            // per file; bigger ones are skipped with a message
const MISS_TTL=5*60*1000;                // don't re-check something missing from both sides for 5 minutes
const KNOWN_KEY='service-slides-media-synced';
const mediaMiss={};
let mediaFolder={parent:'',id:''};
const safeName=n=>String(n||'image').replace(/[\\/:*?"<>|]+/g,'-');
function loadKnown(){try{return new Set(JSON.parse(localStorage.getItem(KNOWN_KEY)||'[]'))}catch{return new Set()}}
function saveKnown(s){try{localStorage.setItem(KNOWN_KEY,JSON.stringify([...s]))}catch{}}

/* ---- what a service needs ---- */
function collectServiceMedia(slides){
  const images=new Set(),decks=new Map();
  (slides||[]).forEach(it=>{
    if(!it)return;
    const deck=(it.k==='PowerPoint'||it.k==='PDF')&&it.presentationId;
    if(deck&&!decks.has(it.presentationId))decks.set(it.presentationId,{pid:it.presentationId,item:it});
    const img=id=>{if(id)images.add(id)};
    if(it.bgMedia&&it.bgMedia.kind==='image')img(it.bgMedia.id);
    (it.parts||[]).forEach(p=>{
      if(p.bgMedia&&p.bgMedia.kind==='image')img(p.bgMedia.id);
      if(!deck)(p.im||[]).forEach(m=>img(m.mediaId));      // deck pages travel inside the PDF
      (p.elements||[]).forEach(el=>{if(el.kind==='image')img(el.mediaId)});
    });
  });
  return {images,decks};
}
function deckPageIds(d){
  const ids=[];
  (d.item.parts||[]).forEach(p=>(p.im||[]).forEach(m=>{if(m.mediaId)ids.push(m.mediaId)}));
  return ids;
}
function deckLocal(d){const ids=deckPageIds(d);return ids.length>0&&ids.every(id=>mediaURLs[id])}

/* ---- Drive media folder ---- */
async function ensureMediaFolder(){
  if(mediaFolder.parent===cloudUi.folderId&&mediaFolder.id)return mediaFolder.id;
  const res=await cloudCall('gdrive_ensure_subfolder',{parentId:cloudUi.folderId,name:'media'});
  if(!res.id)throw cloudErr('Could not create the media folder.');
  mediaFolder={parent:cloudUi.folderId,id:res.id};
  return res.id;
}
async function listMediaFolder(){
  const fid=await ensureMediaFolder();
  const res=await cloudCall('gdrive_list_presentations',{folderId:fid});
  const byKey=new Map();
  (res.files||[]).forEach(f=>{
    const n=f.name||'';
    byKey.set(n.startsWith('deck-')?n:n.split('__')[0],f);
  });
  return {fid,byKey};
}

/* ---- single images ---- */
async function uploadImage(id,fid){
  const blob=await mGet(id);if(!blob)return false;
  if(blob.size>MEDIA_MAX){toast('An image is over 60 MB and was not synced to Drive.');return false}
  const m=mediaMeta.find(x=>x.id===id);
  let nm=safeName(m&&m.name||'image');
  if(!/\.[a-z0-9]{2,5}$/i.test(nm))nm+='.'+(blob.type.split('/')[1]||'png').replace('jpeg','jpg').replace('svg+xml','svg');
  await cloudCall('gdrive_upload_file',{
    name:id+'__'+nm,
    dataBase64:bytesToB64(new Uint8Array(await blob.arrayBuffer())),
    mimeType:blob.type||mimeOf(nm),
    folderId:fid,
    existingFileId:null
  });
  return true;
}
async function downloadImage(id,f){
  const file=await downloadDriveFile(f);
  const nm=f.name.slice(f.name.indexOf('__')+2)||'image';
  const blob=new Blob([await file.arrayBuffer()],{type:mimeOf(nm)});
  await mPut(id,blob);
  mediaURLs[id]=URL.createObjectURL(blob);
  if(!mediaMeta.some(x=>x.id===id))mediaMeta.push({id,name:nm,kind:'image',size:blob.size,createdAt:Date.now(),source:'drive'});
  saveMediaMeta();
}

/* ---- decks as PDF ---- */
async function blobToJpeg(b){
  const url=URL.createObjectURL(b);
  try{
    const img=await new Promise((res,rej)=>{const i=new Image();i.onload=()=>res(i);i.onerror=()=>rej(new Error('Could not read a slide image.'));i.src=url});
    const c=document.createElement('canvas');c.width=img.naturalWidth;c.height=img.naturalHeight;
    const x=c.getContext('2d',{alpha:false});x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);x.drawImage(img,0,0);
    return {url:c.toDataURL('image/jpeg',0.92),w:c.width,h:c.height};
  }finally{URL.revokeObjectURL(url)}
}
async function imagesToPdfBlob(blobs){
  const mod=await import('https://cdn.jsdelivr.net/npm/jspdf@2.5.2/+esm');
  const JsPDF=mod.jsPDF||mod.default?.jsPDF||mod.default;
  if(!JsPDF)throw new Error('Could not load the PDF builder.');
  let doc=null;
  for(const b of blobs){
    const p=await blobToJpeg(b);
    const w=p.w*25.4/96,h=p.h*25.4/96,o=w>=h?'landscape':'portrait';
    if(!doc)doc=new JsPDF({orientation:o,unit:'mm',format:[w,h],compress:true});
    else doc.addPage([w,h],o);
    doc.addImage(p.url,'JPEG',0,0,w,h);
  }
  return new Blob([doc.output('arraybuffer')],{type:'application/pdf'});
}
async function uploadDeck(d,fid){
  const metas=mediaMeta.filter(m=>m.presentationId===d.pid).sort((a,b)=>(a.slideNumber||0)-(b.slideNumber||0));
  if(!metas.length)return false;
  const blobs=[];
  for(const m of metas){const b=await mGet(m.id);if(!b)return false;blobs.push(b)}
  const pdf=await imagesToPdfBlob(blobs);
  if(pdf.size>MEDIA_MAX){toast('A slide deck is over 60 MB as a PDF and was not synced to Drive.');return false}
  const manifest={
    presentationId:d.pid,
    name:metas[0].presentationName||d.item.sourceFile||d.item.t,
    source:metas[0].source||'powerpoint',
    pages:metas.map(m=>({id:m.id,n:m.slideNumber||0}))
  };
  const key='deck-'+d.pid;
  await cloudCall('gdrive_upload_file',{name:key+'.pdf',dataBase64:bytesToB64(new Uint8Array(await pdf.arrayBuffer())),mimeType:'application/pdf',folderId:fid,existingFileId:null});
  // the manifest goes last, so its presence means the PDF is complete
  await cloudCall('gdrive_upload_file',{name:key+'.json',dataBase64:bytesToB64(new TextEncoder().encode(JSON.stringify(manifest))),mimeType:'application/json',folderId:fid,existingFileId:null});
  return true;
}
async function downloadDeck(d,pdfFile,manFile){
  const man=JSON.parse(await (await downloadDriveFile(manFile)).text());
  const pdf=await downloadDriveFile(pdfFile);
  const images=await pdfBlobToSlideImages(pdf,{scale:2});
  for(let i=0;i<man.pages.length;i++){
    const pg=man.pages[i],blob=images[i];
    if(!blob||mediaURLs[pg.id])continue;
    await mPut(pg.id,blob);
    mediaURLs[pg.id]=URL.createObjectURL(blob);
    if(!mediaMeta.some(x=>x.id===pg.id))mediaMeta.push({
      id:pg.id,name:man.name+' — Slide '+(pg.n||i+1),kind:'image',size:blob.size,createdAt:Date.now(),
      source:man.source||'powerpoint',presentationId:d.pid,presentationName:man.name,slideNumber:pg.n||i+1
    });
  }
  saveMediaMeta();
}

/* ---- the sync itself ---- */
async function syncServiceMedia(i){
  if(!cloudReady()||live)return;
  const svc=services[i];if(!svc)return;
  const need=collectServiceMedia(i===activeService?S:(svc.slides||[]));
  const known=loadKnown(),fk=cloudUi.folderId+':',now=Date.now();
  const fresh=k=>!(mediaMiss[k]&&now-mediaMiss[k]<MISS_TTL);
  const imgs=[...need.images].filter(id=>!(mediaURLs[id]&&known.has(fk+id))&&fresh(id));
  const decks=[...need.decks.values()].filter(d=>!(deckLocal(d)&&known.has(fk+'deck-'+d.pid))&&fresh('deck-'+d.pid));
  if(!imgs.length&&!decks.length)return;

  const {fid,byKey}=await listMediaFolder();
  let changed=false;
  for(const id of imgs){
    const f=byKey.get(id);
    if(mediaURLs[id]){
      if(f){known.add(fk+id);continue}
      cloudUi.msg='Uploading images…';refreshSyncUi();
      if(await uploadImage(id,fid))known.add(fk+id);else mediaMiss[id]=Date.now();
    }else if(f){
      cloudUi.msg='Downloading images…';refreshSyncUi();
      await downloadImage(id,f);known.add(fk+id);changed=true;
    }else mediaMiss[id]=Date.now();
  }
  for(const d of decks){
    const key='deck-'+d.pid,pdf=byKey.get(key+'.pdf'),man=byKey.get(key+'.json');
    if(deckLocal(d)){
      if(pdf&&man){known.add(fk+key);continue}
      cloudUi.msg='Uploading slides…';refreshSyncUi();
      if(await uploadDeck(d,fid))known.add(fk+key);else mediaMiss[key]=Date.now();
    }else if(pdf&&man){
      cloudUi.msg='Downloading slides…';refreshSyncUi();
      await downloadDeck(d,pdf,man);changed=true;known.add(fk+key);
    }else mediaMiss[key]=Date.now();
  }
  saveKnown(known);
  if(changed&&!document.body.classList.contains('home-mode')){rStrip();rStage()}
}
async function safeMediaSync(i){
  // Service media now travels inside the .cpres ZIP package; never upload it to Drive's legacy media folder.
  return;
}

/* ---- hook into the existing sync ---- */
const _cloudPush=cloudPush;
cloudPush=async function(i,opts){
  await safeMediaSync(i);                 // media first, so another computer finds it when it sees the service
  return _cloudPush(i,opts);
};
const _cloudPull=cloudPull;
cloudPull=async function(i,fileId){
  await _cloudPull(i,fileId);
  await safeMediaSync(i);
};
const _reconcileService=reconcileService;
reconcileService=async function(i,o){
  await _reconcileService(i,o);
  const l=getLink(services[i]);
  if(l&&!l.conflict&&!l.lost)await safeMediaSync(i);
};
const _cloudOpenAsService=cloudOpenAsService;
cloudOpenAsService=async function(fileId){
  await _cloudOpenAsService(fileId);
  const i=services.findIndex(s=>{const l=getLink(s);return l&&l.fileId===fileId});
  if(i>=0)await withCloudLock('Downloading media…',()=>safeMediaSync(i));
};

}


if(IS_AUDIENCE){
  bootAudienceMode();
}else{
  try{stripAllPresentBlanksEverywhere()}catch{}
  try{updateLogoPreview()}catch{}
  try{applyUiScale(localStorage.getItem('service-slides-ui-scale')||'1')}catch{}
  // Render services list immediately from loaded state
  try{rHome()}catch(err){console.error('rHome',err)}
  loadMedia().then(()=>{
    // Re-render home after media URLs are restored so cards stay visible
    if(document.body.classList.contains('home-mode')){
      try{rHome()}catch{}
    }
  }).catch(()=>{
    if(document.body.classList.contains('home-mode'))try{rHome()}catch{}
  });
}


/* =====================================================================
   Church Presenter Consolidated Extensions
   Settings tabs, templates, church members, team autocomplete,
   drag-and-drop stack/groups, and UI enhancements.
   ===================================================================== */
if(!IS_AUDIENCE){(function(){
  'use strict';

  /* ---------- 1. Consolidated Styles ---------- */
  const css=document.createElement('style');
  css.id='church-presenter-consolidated-css';
  css.textContent=[
    '/* Startup Boot Loader */',
    '#cp-boot{position:fixed;inset:0;z-index:999999;display:grid;place-items:center;background:var(--bg,#0f141a);color:var(--ink,#e8eef5);font-family:Figtree,system-ui,sans-serif}',
    '#cp-boot.hid{opacity:0;pointer-events:none;transition:opacity .35s}',
    '.cp-bi{width:min(340px,88vw);text-align:center}',
    '.cp-bt{font-weight:700;font-size:20px;margin:0 0 18px}',
    '.cp-bb{height:10px;border-radius:999px;background:rgba(128,128,128,.25);overflow:hidden}',
    '.cp-bf{height:100%;width:0;border-radius:999px;background:var(--ac,#2e9e6b);transition:width .2s linear}',
    '.cp-bm{margin-top:12px;font-size:13px;color:var(--mu,#8a97a5)}',
    '',
    '/* Clean Thumbnail Borders */',
    '#strip .th{border:none!important;box-shadow:none!important;background:transparent!important;outline:none!important;padding:0!important}',
    '#strip .th .sl{outline:none!important;border:2px solid #9ca3af!important;border-radius:4px;box-sizing:border-box}',
    ':root[data-theme="dark"] #strip .th .sl{border-color:#8b949e!important}',
    '#strip .th.on .sl,#strip .th.on.es .sl{border-color:#ffd400!important;border-width:3px!important}',
    '#strip .th.ms-on .sl{border:3px solid #3b82f6!important}',
    '#strip .th.ms-on{background:rgba(59,130,246,.14)!important}',
    '#strip .th .pn{display:block;padding:4px 2px 2px;border:none!important;outline:none!important;background:transparent!important}',
    '#stage .el.ms-on,.mid .el.ms-on{outline:2px solid #f4c542!important;outline-offset:2px}',
    '.svc-sep,#service-title-sep{display:none!important}',
    '',
    '/* 5-Column Grid Layout for Home Services and Templates */',
    '#service-list .service-card,#home-templates .service-card{display:grid!important;grid-template-columns:28px minmax(0,1fr) auto 108px 100px!important;align-items:center;column-gap:8px;padding:10px 12px!important}',
    '#home-templates .service-card::before{content:"";grid-column:1;width:28px}',
    '#service-list .service-card>.info,#home-templates .service-card>.info{grid-column:2;min-width:0}',
    '#home-templates .service-card>.tpl-default-badge,button.tpl-default-badge{grid-column:3!important;justify-self:end;margin:0!important;display:inline-flex!important;align-items:center;padding:4px 10px;border-radius:999px;font-size:11px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;cursor:pointer;border:1px solid var(--ln);background:var(--bg);color:var(--mu);opacity:.85;white-space:nowrap}',
    '#home-templates .tpl-default-badge.is-default{background:var(--ac)!important;color:var(--ai,#fff)!important;border-color:var(--ac)!important;opacity:1}',
    '#service-list .service-card>button.pri,#service-list .service-card>button[data-open-service],#service-list .service-card>button[data-home-cloud-open],#home-templates .service-card>button.pri,#home-templates .service-card>button[data-tpl-use]{grid-column:4!important;width:108px!important;padding:7px 10px!important;border-radius:8px!important;font-weight:600;text-align:center}',
    '#service-list .service-card>button[data-del-service],#service-list .service-card>button[data-home-cloud-delete],#home-templates .service-card>button[data-tpl-del]{grid-column:5!important;width:100px!important;padding:7px 8px!important;border-radius:8px!important;text-align:center;background:transparent!important;border:1px solid transparent!important;color:#c22!important}',
    '#service-list .service-card .info>small{display:block;margin:3px 0 0;padding-left:8px}',
    '#home-templates .service-card .info>small{display:block;margin:3px 0 0;padding-left:6px}',
    '#home-templates .service-card{transition:transform .32s cubic-bezier(.2,.8,.2,1)}',
    '#home-templates{margin-top:26px}',
    '.tpl-head{margin:0 0 4px;font-size:calc(13px * var(--ui-scale));text-transform:uppercase;letter-spacing:.05em;color:var(--mu)}',
    '.tpl-sub{margin:0 0 10px}',
    '.tpl-name{font-size:calc(16px * var(--ui-scale));font-weight:600;overflow-wrap:anywhere;padding:3px 0}',
    '.service-title-wrap{display:inline-flex;align-items:center;gap:2px;min-width:0}',
    '.svc-tag-btn{display:inline-flex;align-items:center;gap:4px;padding:4px 8px;border:1px solid transparent;border-radius:6px;background:transparent;font-weight:700;color:var(--ink)}',
    '.svc-tag-btn:hover,.svc-tag-btn:focus-visible{background:var(--bg);border-color:var(--ln)}',
    '.svc-caret{font-size:calc(10px * var(--ui-scale));color:var(--mu)}',
    '.svc-title-row{display:flex;align-items:center;gap:2px;flex-wrap:wrap}',
    '#tag-menu{z-index:1400}',
    '#svc-ctx{z-index:1300}',
    '',
    '/* Drag Over Hints & Stack Ghost */',
    '#strip .thumb-wrap.drag-over-above{box-shadow:0 -3px 0 0 #3b82f6!important}',
    '#strip .thumb-wrap.drag-over-below{box-shadow:0 3px 0 0 #3b82f6!important}',
    '.pk-ghost.pk-stack-slides{position:fixed;z-index:4000;pointer-events:none;width:120px;height:72px}',
    '.pk-ghost.pk-stack-slides .pk-layer{position:absolute;inset:0;border-radius:6px;border:2px solid #3b82f6;overflow:hidden;box-shadow:0 10px 24px rgba(0,0,0,.35);background:var(--bg,#111)}',
    '.pk-ghost.pk-stack-slides .pk-badge{position:absolute;top:-8px;right:-8px;z-index:10;background:#3b82f6;color:#fff;font-size:12px;font-weight:700;padding:2px 7px;border-radius:999px;box-shadow:0 2px 6px rgba(0,0,0,.3)}',
    '',
    '/* Settings Tabbed Dialog */',
    '.settings.ai-settings .dialog.settings-dialog{width:min(880px,100%);height:min(640px,90vh);max-height:90vh;display:flex;flex-direction:column;overflow:hidden;padding:0}',
    '.settings-dialog .st-head{padding:18px 26px 14px;border-bottom:1px solid var(--ln)}',
    '.settings-dialog .st-head h2{margin:0;font-size:calc(20px * var(--ui-scale))}',
    '.st-body{flex:1;min-height:0;display:grid;grid-template-columns:210px minmax(0,1fr)}',
    '.st-nav{padding:12px;border-right:1px solid var(--ln);background:var(--bg);display:flex;flex-direction:column;gap:4px;overflow:auto}',
    '.st-tab{display:flex;align-items:center;gap:10px;width:100%;text-align:left;border:0;background:none;padding:10px 12px;border-radius:8px;font-weight:600;color:var(--mu);cursor:pointer}',
    '.st-tab:hover{background:var(--hl);color:var(--ink)}',
    '.st-tab.active{background:var(--ac);color:var(--ai)}',
    '.st-tab svg{flex:none}',
    '.st-pages{padding:22px 28px;overflow:auto;min-width:0}',
    '.st-page{display:none}.st-page.active{display:block}',
    '.st-page>h3{margin:0 0 4px;font-size:calc(19px * var(--ui-scale))}',
    '.st-lead{margin:0 0 18px!important}',
    '.st-page label{margin:14px 0 4px}',
    '.settings-dialog .st-foot{padding:12px 26px;border-top:1px solid var(--ln);display:flex;justify-content:flex-end;margin:0}',
    '#home-see-more{min-width:140px;padding:9px 18px}',
    '@media(max-width:640px){.st-body{grid-template-columns:1fr;grid-template-rows:auto 1fr}.st-nav{flex-direction:row;border-right:0;border-bottom:1px solid var(--ln);overflow-x:auto}.st-tab{width:auto;white-space:nowrap}}',
    '',
    '/* Church Members & Suggest Box */',
    '#cm-suggest{position:fixed;z-index:2200;background:var(--pn);border:1px solid var(--ln);border-radius:8px;box-shadow:0 12px 30px rgba(0,0,0,.35);max-height:220px;overflow-y:auto;display:none}',
    '#cm-suggest.open{display:block}',
    '#cm-suggest .cm-item{padding:8px 12px;cursor:pointer;border-bottom:1px solid var(--ln);display:flex;flex-direction:column;gap:2px}',
    '#cm-suggest .cm-item:last-child{border-bottom:none}',
    '#cm-suggest .cm-item:hover,#cm-suggest .cm-item.active{background:var(--hl)}',
    '#cm-suggest .cm-item-name{font-weight:600;color:var(--ink)}',
    '#cm-suggest .cm-item-email{font-size:12px;color:var(--mu)}',
    '.cm-table{width:100%;border-collapse:collapse;margin-top:12px}',
    '.cm-table th,.cm-table td{padding:8px 10px;text-align:left;border-bottom:1px solid var(--ln);font-size:calc(13px * var(--ui-scale))}',
    '.cm-table th{color:var(--mu);font-weight:600;font-size:calc(12px * var(--ui-scale))}',
    '.cm-table tr:hover td{background:var(--hl)}',
    '.cm-bar{display:flex;gap:8px;align-items:center;margin:12px 0;flex-wrap:wrap}',
    '.cm-bar input[type=search]{flex:1;min-width:160px}',
    '',
    '/* Save as template modal */',
    '#tpl-save-modal label.tpl-radio{display:flex;align-items:flex-start;gap:8px;margin:8px 0;font-size:calc(13px * var(--ui-scale));cursor:pointer}',
    '#tpl-save-modal label.tpl-radio input{width:auto;margin:3px 0 0}',
    '#tpl-save-modal label.tpl-radio small{display:block;color:var(--mu)}',
    '#tpl-save-modal fieldset{border:1px solid var(--ln);border-radius:8px;margin:14px 0 0;padding:6px 12px 8px}',
    '#tpl-save-modal legend{font-size:calc(12px * var(--ui-scale));color:var(--mu);padding:0 6px}'
  ].join('\n');
  document.head.appendChild(css);

  /* ---------- 2. Startup Boot Overlay ---------- */
  (function initBoot(){
    if(document.getElementById('cp-boot'))return;
    const w=document.createElement('div');w.id='cp-boot';
    w.innerHTML='<div class="cp-bi"><div class="cp-bt"><img src="/logo.png" alt="Church Presenter"></div><div class="cp-bb"><div class="cp-bf" id="cp-bf"></div></div><div class="cp-bm" id="cp-bm">Starting…</div></div>';
    const mount=()=>{if(!document.body){setTimeout(mount,20);return}document.body.classList.add('cp-boot-ready');if(!document.getElementById('cp-boot'))document.body.appendChild(w)};
    mount();
    const t0=Date.now(),MIN=4000;
    const msgs=[[10,'Starting…'],[30,'Loading interface…'],[50,'Preparing services…'],[70,'Syncing Drive libraries…'],[90,'Almost there!']];
    const iv=setInterval(()=>{
      const el=Date.now()-t0,p=Math.min(98,(el/MIN)*98);
      let msg='Starting…';
      for(let i=0;i<msgs.length;i++)if(p>=msgs[i][0])msg=msgs[i][1];
      const f=document.getElementById('cp-bf'),m=document.getElementById('cp-bm');
      if(f)f.style.width=p+'%';if(m)m.textContent=msg;
      // Drive sync runs in the background; never hold the app behind a slow
      // network request once the normal opening animation has completed.
      if(el>=MIN){clearInterval(iv);if(f)f.style.width='100%';if(m)m.textContent='Ready';setTimeout(()=>{const b=document.getElementById('cp-boot');if(b){b.classList.add('hid');setTimeout(()=>{try{b.remove()}catch{}},400)}},200)}
    },100);
    if(!window.__cpStartupSyncPromise)window.__cpStartupSyncPromise=Promise.resolve();
    window.__cpStartupSyncPromise.finally(()=>{window.__cpStartupSyncDone=true});
  })();

  /* ---------- 3. Service Team Top Button & Back Button ---------- */
  function ensureTeamButtons(){
    try{
      if(!document.getElementById('team-top-btn')){
        const share=document.getElementById('export-cpres-btn'),cloud=document.getElementById('cloud-btn');
        if(share||cloud){
          const btn=document.createElement('button');
          btn.type='button';btn.id='team-top-btn';btn.className='icon-btn';
          btn.setAttribute('aria-label','Service team');btn.title='Service team';
          btn.innerHTML='<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>';
          if(share&&share.parentNode)share.parentNode.insertBefore(btn,share);
          else if(cloud&&cloud.parentNode)cloud.parentNode.insertBefore(btn,cloud.nextSibling);
          btn.addEventListener('click',ev=>{ev.preventDefault();try{openTeamModal()}catch{const m=document.getElementById('team-modal');if(m)m.hidden=false}});
        }
      }
      const tb=document.getElementById('share-team-btn'),eb=document.getElementById('share-export-btn');
      if(tb&&eb){
        tb.classList.add('pri');eb.classList.remove('pri');
        const ts=tb.querySelector('small'),es=eb.querySelector('small');
        if(ts)ts.style.cssText='opacity:.85;font-weight:400';
        if(es)es.style.cssText='color:var(--mu);font-weight:400';
        if(!document.getElementById('share-template-btn')){
          eb.insertAdjacentHTML('afterend','<div id="share-template-sep" style="height:1px;background:var(--ln);margin:4px 0"></div><button type="button" id="share-template-btn" style="width:100%;padding:12px 14px;text-align:left;display:flex;flex-direction:column;gap:2px"><span style="font-weight:700">Save service as template</span><small style="color:var(--mu);font-weight:400">Reuse this service as a starting point for new ones</small></button>');
        }
      }
    }catch(err){console.warn('[team-buttons]',err)}
  }

  /* ---------- 4. Settings Dialog Tabs Navigation ---------- */
  function showSettingsPage(p){
    document.querySelectorAll('.st-tab').forEach(x=>x.classList.toggle('active',x.dataset.settingsPage===p));
    document.querySelectorAll('.st-page').forEach(x=>x.classList.toggle('active',x.dataset.page===p));
  }
  function ensureSettingsTabs(){
    const dia=document.querySelector('#settings .settings-dialog');
    if(!dia||dia.querySelector('.st-nav'))return;
    dia.classList.add('ai-settings');
  }
  async function ensureUpdatesPage(){
    const nav=document.querySelector('.st-nav'),pages=document.querySelector('.st-pages');
    if(!nav||!pages)return;
    let tab=nav.querySelector('[data-settings-page="updates"]');
    if(!tab){tab=document.createElement('button');tab.type='button';tab.className='st-tab';tab.dataset.settingsPage='updates';tab.innerHTML='<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M20 7v5h-5M4 17v-5h5"/><path d="M5.6 9a7 7 0 0 1 11.6-2L20 12M4 12l2.8 5a7 7 0 0 0 11.6-2"/></svg> Updates';}
    const membersTab=nav.querySelector('[data-settings-page="members"]'),helpTab=nav.querySelector('[data-settings-page="help"]');
    if(membersTab)nav.insertBefore(tab,membersTab.nextSibling);
    else if(helpTab)nav.insertBefore(tab,helpTab);
    else nav.appendChild(tab);
    let page=pages.querySelector('[data-page="updates"]');
    if(!page){page=document.createElement('section');page.className='st-page';page.dataset.page='updates';page.innerHTML='<h3>Updates</h3><p class="note st-lead">Check GitHub for the latest Church Presenter release.</p><div style="border:1px solid var(--ln);border-radius:10px;padding:16px;margin:12px 0"><div style="font-weight:700">Church Presenter</div><p class="note" style="margin:6px 0 14px">Current version: <strong id="updates-current-version">Loading…</strong></p><p class="note" id="updates-status" aria-live="polite"></p><button type="button" class="pri" id="check-for-updates">Check for updates</button></div>';pages.appendChild(page);}
    try{const version=await getAppVersion();const label=document.getElementById('updates-current-version');if(label)label.textContent=version}catch{const label=document.getElementById('updates-current-version');if(label)label.textContent='Unavailable outside the desktop app'}
  }
  document.addEventListener('click',async ev=>{
    const button=ev.target.closest&&ev.target.closest('#check-for-updates');if(!button)return;
    const status=document.getElementById('updates-status');
    if(!isTauri()){if(status)status.textContent='Updates can only be checked in the installed desktop app.';return}
    button.disabled=true;if(status)status.textContent='Checking GitHub for updates…';
    try{
      const platform=navigator.userAgentData?.platform||navigator.platform||navigator.userAgent;
      const target=/mac/i.test(platform)?'macos-universal':/win/i.test(platform)?'windows-x86_64':null;
      if(!target)throw new Error('This platform is not supported by the automatic updater yet.');
      const update=await checkAppUpdate({target});
      if(!update){if(status)status.textContent='You’re up to date.';return}
      if(status)status.textContent=`Version ${update.version} is available${update.body?`. ${update.body}`:''}`;
      const install=await appConfirm(`Version ${update.version} is available.\n\nDownload and install it now?`,{title:'Update available',confirmLabel:'Update now'});
      if(!install)return;
      if(status)status.textContent='Downloading and installing update…';
      await update.downloadAndInstall();await relaunch();
    }catch(error){console.warn('[updater] Could not check or install update:',error);if(status)status.textContent=`Could not check for updates: ${error?.message||error}`}
    finally{button.disabled=false}
  });

  /* ---------- 5. Church Members & Service Team Autocomplete ---------- */
  function loadMembers(){
    try{const a=JSON.parse(localStorage.getItem(MEMBERS_KEY)||'[]');return Array.isArray(a)?a:[]}catch{return []}
  }
  function saveMembers(list){
    try{localStorage.setItem(MEMBERS_KEY,JSON.stringify(list));return true}catch{toast('Could not save members (storage full).');return false}
  }
  function isAddressLike(s){
    const t=String(s||'').trim();if(!t)return false;
    if(/^address\s*[:：]/i.test(t)||/^addr\s*[:：]/i.test(t))return true;
    if(/^\d{1,4}\s*[/\\-]/.test(t))return true;
    if(/^\d{1,5}\s+[A-Za-z]/.test(t)&&/\b(rd|road|st|street|ave|avenue|dr|drive|cl|close|ct|court|pl|place|cres|crescent|way|ln|lane|pde|parade|hwy|highway|tce|terrace|blvd|boulevard|crt)\b/i.test(t))return true;
    if(/\b(rd|road|st|street|ave|avenue|dr|drive|cl|close|ct|court|pl|place|cres|crescent|way|ln|lane|pde|parade|hwy|highway|tce|terrace)\b/i.test(t)&&/\d/.test(t))return true;
    if(/\b(1|2)\d{3}\b/.test(t)&&/\b(nsw|vic|qld|sa|wa|tas|act|nt)\b/i.test(t))return true;
    if(/\b(1|2)\d{3}\b/.test(t)&&t.length<60&&!/@/.test(t)&&(t.match(/[A-Za-z]/g)||[]).length>3){
      if(/^\d/.test(t)||/\b\d{4}\b/.test(t))return true;
    }
    return false;
  }
  function improvedParseContactsTxt(text){
    const lines=String(text||'').replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n');
    const out=[];let cur=null,detailsStarted=false;
    const isEmailLike=s=>/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s||'').trim());
    const isPhoneLike=s=>{const t=String(s||'').trim();if(!t)return false;const digits=t.replace(/[\s\-\.\(\)\+\/]/g,'');return /^\d{6,15}$/.test(digits)&&((t.match(/[A-Za-z]/g)||[]).length)<2};
    const flush=()=>{
      if(!cur)return;
      let name=String(cur.name||'').trim(),email=String(cur.email||'').trim();
      if(isPhoneLike(name)||isAddressLike(name))name='';
      if(isEmailLike(name)){if(!email)email=name;name=''}
      if(/^contact directory$/i.test(name)||/^email\s*:/i.test(name))name='';
      if(name||email)out.push({name,email});
      cur=null;detailsStarted=false;
    };
    lines.forEach(raw=>{
      const line=String(raw||'').trim();if(!line||/^contact directory$/i.test(line))return;
      if(/^no contact details/i.test(line)){if(cur)detailsStarted=true;return}
      const mEmail=/^(?:e-?mail)\s*[:：]\s*(.+)$/i.exec(line);
      if(mEmail){const em=mEmail[1].trim();if(!cur)cur={name:'',email:''};if(isEmailLike(em)&&!cur.email)cur.email=em;detailsStarted=true;return}
      if(/^(?:phone|mobile|tel|cell)\s*[:：]/i.test(line)){if(!cur)cur={name:'',email:''};detailsStarted=true;return}
      if(/^(?:address|addr)\s*[:：]/i.test(line)||isAddressLike(line)){if(cur)detailsStarted=true;return}
      if(isEmailLike(line)){if(!cur)cur={name:'',email:''};if(!cur.email)cur.email=line;detailsStarted=true;return}
      if(cur&&cur.name&&!detailsStarted)return;
      flush();
      if(isAddressLike(line)){cur=null;detailsStarted=false;return}
      cur={name:line,email:''};detailsStarted=false;
    });
    flush();
    return out;
  }
  function cleanupAddressNames(){
    try{
      const list=loadMembers();if(!list.length)return;
      let changed=false;
      list.forEach(m=>{if(m&&isAddressLike(m.name)){m.name='';changed=true}});
      const next=list.filter(m=>(m.name&&m.name.trim())||(m.email&&m.email.trim()));
      if(changed||next.length!==list.length)saveMembers(next);
    }catch{}
  }
  cleanupAddressNames();

  /* Suggestion popover for Service Team modal */
  let cmSuggestEl=null,cmActiveInput=null,cmSuggestIndex=-1;
  function getSuggestEl(){
    if(!cmSuggestEl){
      cmSuggestEl=document.createElement('div');cmSuggestEl.id='cm-suggest';
      document.body.appendChild(cmSuggestEl);
    }
    return cmSuggestEl;
  }
  function hideSuggest(){if(cmSuggestEl){cmSuggestEl.classList.remove('open');cmSuggestEl.innerHTML=''};cmActiveInput=null;cmSuggestIndex=-1}
  function showSuggest(input){
    const val=(input.value||'').trim().toLowerCase();
    if(!val){hideSuggest();return}
    const list=loadMembers();
    const hits=list.filter(m=>(m.name&&m.name.toLowerCase().includes(val))||(m.email&&m.email.toLowerCase().includes(val))).slice(0,8);
    if(!hits.length){hideSuggest();return}
    const box=getSuggestEl();cmActiveInput=input;cmSuggestIndex=-1;
    box.innerHTML=hits.map((m,idx)=>`<div class="cm-item" data-idx="${idx}" data-name="${e(m.name)}" data-email="${e(m.email)}">
      <div class="cm-item-name">${e(m.name||'No name')}</div>
      ${m.email?`<div class="cm-item-email">${e(m.email)}</div>`:''}
    </div>`).join('');
    const r=input.getBoundingClientRect();
    box.style.left=Math.max(8,r.left)+'px';
    box.style.width=Math.max(220,r.width)+'px';
    box.style.top=(r.bottom+4)+'px';
    box.classList.add('open');
  }
  function applySuggest(item){
    if(!cmActiveInput||!item)return;
    const name=item.dataset.name||'',email=item.dataset.email||'';
    const row=cmActiveInput.closest('.team-row')||cmActiveInput.parentElement;
    if(row){
      const nameInp=row.querySelector('input[data-team-field="name"]');
      const emailInp=row.querySelector('input[data-team-field="email"]');
      if(nameInp)nameInp.value=name;
      if(emailInp)emailInp.value=email;
      if(nameInp)nameInp.dispatchEvent(new Event('input',{bubbles:true}));
      if(emailInp)emailInp.dispatchEvent(new Event('input',{bubbles:true}));
    }else{
      cmActiveInput.value=name||email;
      cmActiveInput.dispatchEvent(new Event('input',{bubbles:true}));
    }
    hideSuggest();
  }

  document.addEventListener('input',ev=>{
    const t=ev.target;
    if(t&&t.dataset&&t.dataset.teamField&&t.closest('#team-modal'))showSuggest(t);
  });
  document.addEventListener('focusin',ev=>{
    const t=ev.target;
    if(t&&t.dataset&&t.dataset.teamField&&t.closest('#team-modal'))showSuggest(t);
  });
  document.addEventListener('keydown',ev=>{
    if(!cmSuggestEl||!cmSuggestEl.classList.contains('open'))return;
    const items=cmSuggestEl.querySelectorAll('.cm-item');if(!items.length)return;
    if(ev.key==='ArrowDown'){ev.preventDefault();cmSuggestIndex=Math.min(cmSuggestIndex+1,items.length-1);items.forEach((it,i)=>it.classList.toggle('active',i===cmSuggestIndex));items[cmSuggestIndex].scrollIntoView({block:'nearest'})}
    else if(ev.key==='ArrowUp'){ev.preventDefault();cmSuggestIndex=Math.max(cmSuggestIndex-1,0);items.forEach((it,i)=>it.classList.toggle('active',i===cmSuggestIndex));items[cmSuggestIndex].scrollIntoView({block:'nearest'})}
    else if(ev.key==='Enter'&&cmSuggestIndex>=0){ev.preventDefault();applySuggest(items[cmSuggestIndex])}
    else if(ev.key==='Escape'){hideSuggest()}
  });
  document.addEventListener('click',ev=>{
    const it=ev.target.closest&&ev.target.closest('#cm-suggest .cm-item');
    if(it){applySuggest(it);return}
    if(!ev.target.closest('#cm-suggest')&&!ev.target.closest('#team-modal'))hideSuggest();
  });

  /* ---------- 6. Templates System & Dialogs ---------- */
  if(!document.getElementById('tpl-save-modal')){
    document.body.insertAdjacentHTML('beforeend',`
      <div id="tpl-save-modal" class="app-modal" hidden>
        <div class="app-dialog" role="dialog" aria-modal="true" aria-labelledby="tpl-save-title" style="width:min(500px,calc(100vw - 32px))">
          <header><h2 id="tpl-save-title">Save as template</h2></header>
          <p class="note" id="tpl-save-src" style="margin:0 0 6px"></p>
          <label>Save as<select id="tpl-save-target"></select></label>
          <label>Template name<input type="text" id="tpl-save-name" placeholder="e.g. Sunday morning" autocomplete="off"></label>
          <fieldset id="tpl-songs-box"><legend>Songs</legend>
            <label class="tpl-radio"><input type="radio" name="tpl-songs" value="strip" checked><span>Leave a placeholder for each song<small>Best for a reusable order of service. Lyrics are not saved.</small></span></label>
            <label class="tpl-radio"><input type="radio" name="tpl-songs" value="keep"><span>Keep the songs as they are<small>Use this when the same songs repeat.</small></span></label>
          </fieldset>
          <footer style="display:flex;justify-content:flex-end;gap:8px;margin-top:16px">
            <button type="button" id="tpl-save-cancel">Cancel</button>
            <button type="button" id="tpl-save-confirm" class="pri">Save template</button>
          </footer>
        </div>
      </div>
      <div id="svc-ctx" class="outline-menu" role="menu" hidden></div>
    `);
  }
  // Start from on New Service Form
  (function initTemplateSelect(){
    const fields=document.querySelector('#service-form .fields');
    if(fields&&!document.getElementById('service-template')){
      fields.insertAdjacentHTML('afterbegin','<label style="grid-column:1/-1">Start from<select id="service-template"><option value="">Blank service</option></select><small id="service-template-note" class="note" style="display:block;margin-top:4px"></small></label>');
    }
  })();

  let tplSaveIdx=-1,tplSyncing=false,tplLastSync=0;
  const TPL_SYNC_MS=30000;

  function songPlaceholder(it){
    const ns=mkS('Song',['Song title'],0,6.5);
    ns.bg=it.bg;ns.dim=it.dim;
    if(it.customBg)ns.customBg=it.customBg;
    if(it.bgMedia)ns.bgMedia=cl(it.bgMedia);
    return ns;
  }
  function defaultName(svc){return ((svc.name||'').trim()||(svc.tag||'').trim()||'Service')+' template'}
  function openTplSave(i){
    const svc=services[i];if(!svc){toast('No service to save.');return}
    tplSaveIdx=i;
    const body=serviceSyncBody(i);
    const hasSongs=body.slides.some(it=>it&&it.k==='Song');
    const list=loadTemplates().slice().sort((a,b)=>String(a.name||'').localeCompare(String(b.name||'')));
    const sel=document.getElementById('tpl-save-target');
    sel.innerHTML='<option value="new">A new template</option>'+list.map(t=>`<option value="${e(t.id)}">Update \u201c${e(t.name)}\u201d</option>`).join('');
    sel.value='new';
    document.getElementById('tpl-save-name').value=defaultName(svc);
    document.getElementById('tpl-save-src').textContent='From: '+formatServiceTitle(svc);
    document.getElementById('tpl-songs-box').hidden=!hasSongs;
    const r=document.querySelector('#tpl-save-modal input[name="tpl-songs"][value="strip"]');if(r)r.checked=true;
    document.getElementById('tpl-save-modal').hidden=false;
    setTimeout(()=>{const n=document.getElementById('tpl-save-name');if(n){n.focus();n.select()}},30);
  }
  function closeTplSave(){document.getElementById('tpl-save-modal').hidden=true;tplSaveIdx=-1}
  async function confirmSaveTemplate(){
    const i=tplSaveIdx,svc=services[i];if(!svc){closeTplSave();return}
    const name=document.getElementById('tpl-save-name').value.trim();
    if(!name){toast('Please enter a template name.');return}
    let target=document.getElementById('tpl-save-target').value;
    const list=loadTemplates();
    const dup=list.find(t=>t.name.trim().toLowerCase()===name.toLowerCase()&&t.id!==target);
    if(dup){
      if(!await appConfirm('A template named "'+dup.name+'" already exists.\n\nReplace it with this service?',{title:'Replace template',confirmLabel:'Replace'}))return;
      target=dup.id;
    }
    const body=serviceSyncBody(i);
    const strip=!document.getElementById('tpl-songs-box').hidden&&(document.querySelector('#tpl-save-modal input[name="tpl-songs"]:checked')||{}).value==='strip';
    if(strip)body.slides=body.slides.map(it=>it&&it.k==='Song'?songPlaceholder(it):it);
    const now=Date.now();
    let t=target!=='new'?list.find(x=>x.id===target):null;
    if(t){t.name=name;t.team=body.team;t.slides=body.slides;t.updatedAt=now;t.dirty=true;t.localOnly=false}
    else{t={id:newTplId(),name,team:body.team,slides:body.slides,updatedAt:now,dirty:true,localOnly:false,driveFileId:''};list.push(t)}
    if(!saveTemplates(list))return;
    closeTplSave();
    renderTemplatesHome();refreshTemplateSelect();
    toast(cloudReady()?'Saved template "'+name+'". Syncing to Drive…':'Saved template "'+name+'" on this computer.');
    syncTemplates({force:true});
  }

  function useTemplate(id){
    openNewServiceForm();
    refreshTemplateSelect(id);
  }
  function createFromTemplate(tpl,form){
    const slides=cl(tpl.slides||[]).map(it=>remapServiceItemIds(it));
    const svc={name:form.name,date:form.date,tag:form.tag,slides:slides.length?slides:[mkS('Welcome',['Welcome'],0,6.5)],team:tpl.team?cl(tpl.team):defaultTeam(),lastEdited:Date.now()};
    ensureTeam(svc);
    services.push(svc);activeService=services.length-1;
    S=cl(svc.slides);sel=0;si=0;A=S[0]||null;aj=0;
    saveState();
    const f=document.getElementById('service-form');f.classList.remove('open');f.reset();
    document.body.classList.remove('home-mode');
    rAll();
    toast('Created "'+formatServiceTitle(svc)+'" from the "'+tpl.name+'" template.');
  }

  // Intercept New Service form submission when a template is selected
  document.addEventListener('submit',ev=>{
    const f=ev.target;if(!f||f.id!=='service-form')return;
    const sel=document.getElementById('service-template');
    if(!sel||!sel.value)return;
    const tpl=loadTemplates().find(t=>t.id===sel.value);
    if(!tpl)return;
    ev.preventDefault();ev.stopImmediatePropagation();
    const tag=(document.getElementById('service-tag').value||'').trim();
    if(!tag){toast('Choose a service tag.');document.getElementById('service-tag').focus();return}
    const name=(document.getElementById('service-name').value||'').trim();
    const date=document.getElementById('service-date').value;
    if(!date){toast('Choose a date.');document.getElementById('service-date').focus();return}
    createFromTemplate(tpl,{tag,name,date});
  },true);

  // Template sync with Google Drive
  async function syncTemplates(opts={}){
    if(!cloudReady()||live||syncBusy||tplSyncing)return;
    if(!opts.force&&Date.now()-tplLastSync<TPL_SYNC_MS)return;
    tplSyncing=true;tplLastSync=Date.now();
    try{
      await withCloudLock('Syncing templates…',async()=>{
        const folder=await cloudCall('gdrive_ensure_subfolder',{parentId:cloudUi.folderId,name:'templates'});
        if(!folder.id)throw cloudErr('Could not open the templates folder.');
        const tombs=loadTombs(),keep=[];
        for(let ti=0;ti<tombs.length;ti++){
          const dr=await cloudInvoke('gdrive_delete_file',{fileId:tombs[ti]});
          if(dr&&dr.ok===false&&classifyErr(dr.error)!=='notfound')keep.push(tombs[ti]);
        }
        saveTombs(keep);
        const listing=await cloudCall('gdrive_list_presentations',{folderId:folder.id});
        const files=(listing.files||[]).filter(f=>/\.cpres$/i.test(f.name||'')&&!keep.includes(f.id));
        const byId={};files.forEach(f=>{byId[f.id]=f});
        const list=loadTemplates();let changed=false;
        for(let i=0;i<list.length;i++){
          const t=list[i];
          if(t.driveFileId&&!byId[t.driveFileId]&&!t.dirty){t.driveFileId='';t.localOnly=true;changed=true}
        }
        for(let i=0;i<list.length;i++){
          const t=list[i];
          if(t.dirty||(!t.driveFileId&&!t.localOnly)){
            const json=JSON.stringify({format:'church-presenter-template',version:CPRES_SCHEMA,exportedAt:new Date().toISOString(),service:{name:t.name,date:'',tag:'',team:t.team,slides:t.slides}},null,2);
            const data=bytesToB64(new TextEncoder().encode(json));
            const safeName=String(t.name||'template').replace(/[\\/:*?"<>|]+/g,'-');
            const res=await cloudCall('gdrive_upload_file',{name:safeName+'.cpres',dataBase64:data,mimeType:'application/json',folderId:t.driveFileId?null:folder.id,existingFileId:t.driveFileId||null});
            t.driveFileId=res.id;t.driveModified=res.modifiedTime||'';t.dirty=false;t.localOnly=false;
            changed=true;
          }
        }
        if(changed)saveTemplates(list);
      },{quiet:opts.quiet!==false});
    }catch(err){console.warn('[templates] sync',err)}finally{tplSyncing=false}
    try{renderTemplatesHome();refreshTemplateSelect()}catch{}
  }

  // FLIP reorder animation when setting a template as default
  function flipReorderDefaultFirst(targetCard){
    try{
      const box=document.getElementById('home-templates');if(!box)return;
      const list=box.querySelector('.service-list')||box;
      const cards=Array.from(list.querySelectorAll(':scope > .service-card'));
      if(cards.length<2||cards[0]===targetCard)return;
      const prev=cards.map(c=>({c,r:c.getBoundingClientRect()}));
      list.insertBefore(targetCard,cards[0]);
      prev.forEach(item=>{
        const n=item.c.getBoundingClientRect(),dx=item.r.left-n.left,dy=item.r.top-n.top;
        if(!dx&&!dy)return;
        item.c.style.transition='none';item.c.style.transform=`translate(${dx}px,${dy}px)`;
      });
      requestAnimationFrame(()=>{requestAnimationFrame(()=>{
        prev.forEach(item=>{item.c.style.transition='transform .32s cubic-bezier(.2,.8,.2,1)';item.c.style.transform=''});
        setTimeout(()=>{prev.forEach(item=>{item.c.style.transition=''})},350);
      })});
    }catch(err){console.warn('[flip]',err)}
  }

  /* ---------- 7. Multi-Slide & Element Drag Systems ---------- */
  function buildStackGhost(wraps){
    const g=document.createElement('div');
    g.className='pk-ghost pk-stack-slides';
    const n=Math.min(wraps.length,4);
    for(let i=n-1;i>=0;i--){
      const layer=document.createElement('div');
      layer.className='pk-layer';
      const off=(n-1-i)*5;
      layer.style.transform=`translate(${off}px,${off}px) rotate(${off*0.4}deg)`;
      layer.style.zIndex=String(i+1);
      const thumb=wraps[i].querySelector('.th .sl, .sl');
      if(thumb){
        const clone=thumb.cloneNode(true);
        clone.style.width='100%';clone.style.height='100%';
        layer.appendChild(clone);
      }else{layer.style.background='var(--bg,#222)'}
      g.appendChild(layer);
    }
    const badge=document.createElement('div');
    badge.className='pk-badge';
    badge.textContent=String(wraps.length);
    g.appendChild(badge);
    return g;
  }

  // Pointer listener on strip for multi-slide stack drag
  document.addEventListener('pointerdown',ev=>{
    if(ev.button!==0||live)return;
    const wrap=ev.target&&ev.target.closest&&ev.target.closest('#strip .thumb-wrap[data-i][data-j]');
    if(!wrap)return;
    const key=sk(wrap.dataset.i,wrap.dataset.j);
    if(!MS.slides||MS.slides.size<2||!MS.slides.has(key))return;
    const keys=Array.from(MS.slides).map(pk);
    const itemIndex=keys[0].i;
    if(!keys.every(p=>p.i===itemIndex))return;
    const fromJs=keys.map(p=>p.j).sort((a,b)=>a-b);
    const wraps=fromJs.map(j=>document.querySelector('#strip .thumb-wrap[data-i="'+itemIndex+'"][data-j="'+j+'"]')).filter(Boolean);
    if(wraps.length<2)return;
    // The regular single-slide pointer handler is also registered on the document.
    // Keep it from starting a competing drag for this selected stack.
    ev.stopImmediatePropagation();
    ev.preventDefault();
    ev.stopPropagation();
    const startX=ev.clientX,startY=ev.clientY;
    let moved=false,ghost=null,tgt=null;
    const stripEl=document.getElementById('strip');
    const clearHints=()=>{document.querySelectorAll('#strip .thumb-wrap.drag-over-above,#strip .thumb-wrap.drag-over-below').forEach(x=>x.classList.remove('drag-over-above','drag-over-below'))};
    const onMove=e=>{
      if(!moved&&Math.hypot(e.clientX-startX,e.clientY-startY)<6)return;
      if(!moved){
        moved=true;ghost=buildStackGhost(wraps);document.body.appendChild(ghost);
        try{pushUndo()}catch{}
      }
      if(ghost){ghost.style.left='0px';ghost.style.top='0px';ghost.style.transform=`translate(${e.clientX-40}px,${e.clientY-24}px)`}
      clearHints();tgt=null;
      const el=document.elementFromPoint(e.clientX,e.clientY);
      const tw=el&&el.closest&&el.closest('#strip .thumb-wrap[data-i][data-j]');
      if(tw&&+tw.dataset.i===itemIndex){
        const r=tw.getBoundingClientRect(),before=e.clientY<r.top+r.height/2;
        tw.classList.add(before?'drag-over-above':'drag-over-below');
        tgt={j:+tw.dataset.j,before};
      }
      if(stripEl){
        const sr=stripEl.getBoundingClientRect();
        if(e.clientY<sr.top+40)stripEl.scrollTop-=12;
        else if(e.clientY>sr.bottom-40)stripEl.scrollTop+=12;
      }
    };
    const onUp=()=>{
      document.removeEventListener('pointermove',onMove,true);
      document.removeEventListener('pointerup',onUp,true);
      if(ghost)ghost.remove();
      clearHints();
      if(!moved||!tgt)return;
      try{
        const it=S[itemIndex];if(!it)return;
        const ids=fromJs.map(j=>it.ord[j]);
        fromJs.slice().reverse().forEach(j=>it.ord.splice(j,1));
        let toJ=tgt.j+(tgt.before?0:1);
        const removedBefore=fromJs.filter(j=>j<toJ).length;
        toJ=Math.max(0,toJ-removedBefore);
        ids.forEach((id,n)=>it.ord.splice(toJ+n,0,id));
        aj=toJ;si=toJ;sel=itemIndex;A=it;
        MS.slides=new Set(ids.map((_,n)=>sk(itemIndex,toJ+n)));
        saveSoon();rAll();
        setTimeout(paintSlides,0);
      }catch(err){console.warn(err)}
    };
    document.addEventListener('pointermove',onMove,true);
    document.addEventListener('pointerup',onUp,true);
  },true);

  // Stage multi-element group drag
  let gDrag=null;
  document.addEventListener('pointerdown',ev=>{
    if(ev.button!==0)return;
    const node=ev.target&&ev.target.closest&&ev.target.closest('#stage [data-el],.mid [data-el]');
    if(!node||(ev.target.closest&&ev.target.closest('.handle')))return;
    const id=node.getAttribute('data-el');
    if(!id||MS.els.size<2||!MS.els.has(id))return;
    try{
      const part=sls(A)[aj];
      const els=((part&&part.elements)||[]).filter(el=>MS.els.has(el.id));
      if(els.length<2)return;
      const stage=document.querySelector('#stage .sl')||document.querySelector('.mid .sl');
      const rect=stage?stage.getBoundingClientRect():{width:1,height:1};
      gDrag={pid:ev.pointerId,sx:ev.clientX,sy:ev.clientY,moved:false,w:rect.width,h:rect.height,origins:els.map(el=>({el,x:+el.x,y:+el.y}))};
      try{elDrag=null}catch{}
    }catch{gDrag=null}
  },true);
  document.addEventListener('pointermove',ev=>{
    if(!gDrag||ev.pointerId!==gDrag.pid)return;
    const dx=ev.clientX-gDrag.sx,dy=ev.clientY-gDrag.sy;
    if(!gDrag.moved&&Math.hypot(dx,dy)<4)return;
    if(!gDrag.moved){gDrag.moved=true;try{pushUndo()}catch{}try{elDrag=null}catch{}}
    const pdx=(dx/gDrag.w)*100,pdy=(dy/gDrag.h)*100;
    gDrag.origins.forEach(o=>{
      o.el.x=Math.max(0,Math.min(95,o.x+pdx));
      o.el.y=Math.max(0,Math.min(95,o.y+pdy));
      const n=document.querySelector('#stage [data-el="'+o.el.id+'"],.mid [data-el="'+o.el.id+'"]');
      if(n){n.style.left=o.el.x+'%';n.style.top=o.el.y+'%'}
    });
    ev.preventDefault();
  },true);
  document.addEventListener('pointerup',ev=>{
    if(!gDrag||ev.pointerId!==gDrag.pid)return;
    const moved=gDrag.moved;gDrag=null;
    if(moved){
      saveSoon();updateCurrentThumb();
      setTimeout(paintEls,0);
    }
  },true);

  // Multi-slide strip click selection (Shift for range, Ctrl/Cmd for toggle)
  document.addEventListener('click',ev=>{
    const th=ev.target&&ev.target.closest&&ev.target.closest('#strip .th[data-i][data-j]');
    if(!th||(ev.target.closest&&ev.target.closest('.slide-actions,#slide-menu')))return;
    const key=sk(th.dataset.i,th.dataset.j);
    const keys=Array.from(document.querySelectorAll('#strip .th[data-i][data-j]')).map(t=>sk(t.dataset.i,t.dataset.j));
    if(ev.shiftKey){
      ev.preventDefault();ev.stopImmediatePropagation();
      ev.preventDefault();ev.stopImmediatePropagation();
      if(MS.slideAnchor==null)MS.slideAnchor=key;
      let a=keys.indexOf(MS.slideAnchor),b=keys.indexOf(key);
      if(a<0)a=b;if(a>b){const tmp=a;a=b;b=tmp}
      MS.slides=new Set(keys.slice(a,b+1));
      try{const p=pk(key);A=S[p.i];sel=p.i;aj=p.j;si=p.j;elSel=null}catch{}
      rStage();rPanel();setTimeout(paintSlides,0);
      return;
    }
    if(ev.ctrlKey||ev.metaKey){
      ev.preventDefault();ev.stopImmediatePropagation();
      ev.preventDefault();ev.stopImmediatePropagation();
      if(MS.slides.has(key))MS.slides.delete(key);else MS.slides.add(key);
      MS.slideAnchor=key;
      try{const p=pk(key);A=S[p.i];sel=p.i;aj=p.j;si=p.j;elSel=null}catch{}
      rStage();setTimeout(paintSlides,0);
      return;
    }
    MS.slides=new Set([key]);MS.slideAnchor=key;
    setTimeout(paintSlides,0);
  },true);

  // Stage element multi-selection click
  document.addEventListener('pointerdown',ev=>{
    if(ev.button!==0)return;
    const node=ev.target&&ev.target.closest&&ev.target.closest('#stage [data-el],.mid [data-el]');
    if(!node){
      if(ev.target.closest&&ev.target.closest('#stage .sl,.mid .sl')&&MS.els.size&&!ev.target.closest('[data-el]')){
        MS.els.clear();MS.elAnchor=null;setTimeout(paintEls,0);
      }
      return;
    }
    if(ev.target.closest&&ev.target.closest('.handle'))return;
    try{if(A&&isSongItem(A))return}catch{}
    const id=node.getAttribute('data-el');if(!id)return;
    if(ev.shiftKey){
      ev.preventDefault();ev.stopImmediatePropagation();
      try{
        const part=sls(A)[aj];
        const ids=((part&&part.elements)||[]).map(x=>x.id);
        if(MS.elAnchor!=null&&ids.includes(MS.elAnchor)&&ids.includes(id)){
          let a=ids.indexOf(MS.elAnchor),b=ids.indexOf(id);if(a>b){const tmp=a;a=b;b=tmp}
          MS.els=new Set(ids.slice(a,b+1));
        }else MS.els.add(id);
        MS.elAnchor=id;
      }catch{MS.els.add(id);MS.elAnchor=id}
      elSel=id;setTimeout(()=>{rPanel();paintEls()},0);
      return;
    }
    if(ev.ctrlKey||ev.metaKey){
      ev.preventDefault();ev.stopImmediatePropagation();
      if(MS.els.has(id))MS.els.delete(id);else MS.els.add(id);
      MS.elAnchor=id;elSel=id;setTimeout(()=>{rPanel();paintEls()},0);
      return;
    }
    if(MS.els.size>1&&MS.els.has(id)){elSel=id;setTimeout(paintEls,0);return}
    MS.els=new Set([id]);MS.elAnchor=id;setTimeout(paintEls,0);
  },true);

  /* ---------- 8. Church Logo Slide Section ---------- */
  let logoMediaCache={};
  try{logoMediaCache=JSON.parse(localStorage.getItem('church-presenter-logo-media')||'{}')||{}}catch{logoMediaCache={}}
  async function ensureLogoMedia(src){
    const key=hashStr(src);let id=logoMediaCache[key];
    if(id&&mediaURLs[id])return id;
    let blob;
    if(src.indexOf('data:')===0){
      const comma=src.indexOf(','),head=src.slice(5,comma),body=src.slice(comma+1);
      const b64=/;base64$/i.test(head),mime=head.replace(/;base64$/i,'').split(';')[0]||'application/octet-stream';
      if(b64){const bin=atob(body),arr=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);blob=new Blob([arr],{type:mime})}
      else blob=new Blob([decodeURIComponent(body)],{type:mime});
    }else{blob=await (await fetch(src)).blob()}
    id='m'+Date.now().toString(36)+Math.random().toString(36).slice(2,8);
    await mPut(id,blob);
    mediaURLs[id]=URL.createObjectURL(blob);
    const l=churchLogos.find(x=>x.id===effectiveLogoId());
    mediaMeta.push({id,name:'Church logo'+(l?' - '+l.name:''),kind:'image',size:blob.size,createdAt:Date.now(),source:'logo'});
    saveMediaMeta();
    logoMediaCache[key]=id;
    try{localStorage.setItem('church-presenter-logo-media',JSON.stringify(logoMediaCache))}catch{}
    return id;
  }
  function imgSize(src){
    return new Promise(res=>{const im=new Image();im.onload=()=>res({w:im.naturalWidth||0,h:im.naturalHeight||0});im.onerror=()=>res({w:0,h:0});im.src=src});
  }
  window.addChurchLogoSection=async function(){
    try{syncEffectiveLogo()}catch{}
    const src=churchLogo||'',size=churchLogoSize||80,els=[];
    if(src){
      const mid=await ensureLogoMedia(src);
      const d=await imgSize(src);
      const nr=d.w&&d.h?d.w/d.h:0;
      let w,h,fit='fill';
      if(nr){
        const kw=size/100*16,kh=size/100*9;
        let dw=kw,dh=kw/nr;
        if(dh>kh){dh=kh;dw=kh*nr}
        w=dw/16*100;h=dh/9*100;
      }else{w=size;h=size;fit='contain'}
      const x=churchLogoHPos==='flex-start'?0:churchLogoHPos==='flex-end'?100-w:(100-w)/2;
      const y=churchLogoVPos==='flex-start'?0:churchLogoVPos==='flex-end'?100-h:(100-h)/2;
      const el={id:uid(),kind:'image',src:mediaURLs[mid],mediaId:mid,x,y,w,h,rot:0,z:1,fit};
      if(nr)el.nr=nr;
      els.push(el);
    }else{toast('No logo uploaded yet. Add one in Settings, Church logo.')}
    pushUndo();
    const part={id:uid(),n:'Slide 1',tx:'',elements:els};
    const pickIdx=BG.findIndex(b=>b.pick);
    const item=mk('Church logo','Slide',[part],[part.id],pickIdx>=0?pickIdx:3,5.2);
    item.customBg=logoBgCss();item.dim=0;
    S.splice(Math.max(0,sel)+1,0,item);
    sel=S.indexOf(item);A=item;si=0;aj=0;clr=null;elSel=null;tab='Structure';
    if(services[activeService])services[activeService].slides=sanitizeSlidesForSave(S);
    saveState();rAll();
    setTimeout(()=>{scrollNewItemToTop(sel);scrollCurrentSlideIntoView()},0);
    if(src)toast('Added logo slide.');
  };

  /* ---------- 9. Auto-Upload for Brand New Services ---------- */
  const knownSvc=new WeakSet();
  services.forEach(s=>knownSvc.add(s));
  function queueNewServiceUpload(svc,n=0){
    setTimeout(()=>{
      const i=services.indexOf(svc);
      if(i<0||getLink(svc)||loadCloudMeta().autoUpload===false||!cloudReady())return;
      if(live||syncBusy){if(n<15)queueNewServiceUpload(svc,n+1);return}
      withCloudLock('Saving to Drive…',()=>cloudPush(i),{quiet:true});
    },n?2000:1200);
  }
  const _origSaveState=saveState;
  saveState=function(){
    const r=_origSaveState.apply(this,arguments);
    try{
      services.forEach(s=>{
        if(knownSvc.has(s))return;
        knownSvc.add(s);
        if(getLink(s))return;
        ensureSvcId(s);
        seenHash[s.id]='__new__';
        queueNewServiceUpload(s,0);
      });
    }catch(err){console.warn('[auto-upload]',err)}
    return r;
  };

  /* ---------- 10. Global Event Delegation ---------- */
  document.addEventListener('click',ev=>{
    const t=ev.target;if(!t||!t.closest)return;
    const pg=t.closest('[data-settings-page]');
    if(pg){showSettingsPage(pg.dataset.settingsPage);return}
    if(t.closest('#home-see-more')){homeShown+=5;rHome();return}
    if(t.closest('#home-see-older')){olderServicesVisible=!olderServicesVisible;homeShown=olderServicesVisible?Number.MAX_SAFE_INTEGER:5;rHome();return}
    let b;
    /* Template Use / Rename / Delete / Save-dialog buttons are handled once, in the block further down. */
    if((b=t.closest('[data-tpl-default]'))){setDefaultTemplate(b.dataset.tplDefault);return}
    if((b=t.closest('[data-tag-service]'))){ev.preventDefault();ev.stopPropagation();showTagMenu(b,+b.dataset.tagService);return}
    if((b=t.closest('#service-tag-btn'))){showTagMenu(b,activeService);return}
    const pick=t.closest('#tag-menu [data-tag-pick]');
    if(pick){
      const m=document.getElementById('tag-menu');
      const idx=+m.dataset.idx;hideTagMenu();
      pickServiceTag(idx,pick.dataset.tagPick);return;
    }
    if(!t.closest('#tag-menu'))hideTagMenu();
    if(t.closest('#new-service')||t.closest('#make-new-service')){setTimeout(()=>refreshTemplateSelect(null,true),0)}
  });

  document.addEventListener('change',ev=>{
    const t=ev.target;if(!t)return;
    if(t.id==='service-template')updateTplNote();
    if(t.id==='tpl-save-target'){
      const svc=services[tplSaveIdx];
      if(t.value==='new')document.getElementById('tpl-save-name').value=svc?defaultName(svc):'';
      else{const tp=loadTemplates().find(x=>x.id===t.value);if(tp)document.getElementById('tpl-save-name').value=tp.name}
    }
  });

  document.addEventListener('keydown',ev=>{
    if(ev.key==='Escape'){
      hideTagMenu();
      const sv=document.getElementById('tpl-save-modal');if(sv&&!sv.hidden){ev.preventDefault();closeTplSave()}
      return;
    }
  });

  // Boot extensions
  ensureTeamButtons();
  ensureSettingsTabs();
  renderTitleBar(services[activeService]);
  refreshTemplateSelect();
  setTimeout(()=>{syncTemplates({force:true})},4000);
  setInterval(()=>{syncTemplates()},60000);

  console.info('[church-presenter] consolidated extensions active');
})();}


/* mjs-patch-v4 */
if(!IS_AUDIENCE){(function(){
  var TPL_KEY='service-slides-templates',TOMB_KEY='service-slides-template-deleted',TPL_SYNC_MS=30000;
  var tplLastSync=0,tplSyncing=false,tplSaveIdx=-1;

  /* ---------------- styles + dialogs ---------------- */
  var css=document.createElement('style');
  css.textContent=[
    '#home-templates{margin-top:26px}',
    '.tpl-head{margin:0 0 4px;font-size:calc(13px * var(--ui-scale));text-transform:uppercase;letter-spacing:.05em;color:var(--mu)}',
    '.tpl-sub{margin:0 0 10px}',
    '.tpl-name{font-size:calc(16px * var(--ui-scale));font-weight:600;overflow-wrap:anywhere;padding:3px 0}',
    '#tpl-save-modal label.tpl-radio{display:flex;align-items:flex-start;gap:8px;margin:8px 0;font-size:calc(13px * var(--ui-scale));cursor:pointer}',
    '#tpl-save-modal label.tpl-radio input{width:auto;margin:3px 0 0}',
    '#tpl-save-modal label.tpl-radio small{display:block;color:var(--mu)}',
    '#tpl-save-modal fieldset{border:1px solid var(--ln);border-radius:8px;margin:14px 0 0;padding:6px 12px 8px}',
    '#tpl-save-modal legend{font-size:calc(12px * var(--ui-scale));color:var(--mu);padding:0 6px}',
    '#svc-ctx{z-index:1300}'
  ].join('\n');
  document.head.appendChild(css);

  document.body.insertAdjacentHTML('beforeend',[
    '<div id="tpl-save-modal" class="app-modal" hidden><div class="app-dialog" role="dialog" aria-modal="true" aria-labelledby="tpl-save-title" style="width:min(500px,calc(100vw - 32px))">',
    '<header><h2 id="tpl-save-title">Save as template</h2></header>',
    '<p class="note" id="tpl-save-src" style="margin:0 0 6px"></p>',
    '<label>Save as<select id="tpl-save-target"></select></label>',
    '<label>Template name<input type="text" id="tpl-save-name" placeholder="e.g. Sunday morning" autocomplete="off"></label>',
    '<fieldset id="tpl-songs-box"><legend>Songs</legend>',
    '<label class="tpl-radio"><input type="radio" name="tpl-songs" value="strip" checked><span>Leave a placeholder for each song<small>Best for a reusable order of service. Lyrics are not saved.</small></span></label>',
    '<label class="tpl-radio"><input type="radio" name="tpl-songs" value="keep"><span>Keep the songs as they are<small>Use this when the same songs repeat.</small></span></label>',
    '</fieldset>',
    '<footer style="display:flex;justify-content:flex-end;gap:8px;margin-top:16px"><button type="button" id="tpl-save-cancel">Cancel</button><button type="button" id="tpl-save-confirm" class="pri">Save template</button></footer>',
    '</div></div>',
    '<div id="svc-ctx" class="outline-menu" role="menu" hidden></div>'
  ].join(''));

  /* "Start from" on the New service form (first row) */
  (function(){
    var fields=document.querySelector('#service-form .fields');
    if(!fields||document.getElementById('service-template'))return;
    fields.insertAdjacentHTML('afterbegin','<label style="grid-column:1/-1">Start from<select id="service-template"><option value="">Blank service</option></select><small id="service-template-note" class="note" style="display:block;margin-top:4px"></small></label>');
  })();

  /* Share window: reword the template button */
  (function(){
    var b=document.getElementById('share-template-btn');
    var s=b&&b.querySelector('small');
    if(s)s.textContent='Reuse this service as a starting point for new ones';
  })();

  /* ---------------- storage ---------------- */
  function loadTemplates(){try{var a=JSON.parse(localStorage.getItem(TPL_KEY)||'[]');return Array.isArray(a)?a:[]}catch(err){return []}}
  function saveTemplates(list){try{localStorage.setItem(TPL_KEY,JSON.stringify(list));return true}catch(err){toast('Could not save the template: storage is full.');return false}}
  function loadTombs(){try{var a=JSON.parse(localStorage.getItem(TOMB_KEY)||'[]');return Array.isArray(a)?a:[]}catch(err){return []}}
  function saveTombs(a){try{localStorage.setItem(TOMB_KEY,JSON.stringify(a))}catch(err){}}
  function newTplId(){return 'tpl-'+Date.now().toString(36)+Math.random().toString(36).slice(2,6)}
  function byName(a,b){return String(a.name||'').localeCompare(String(b.name||''))}
  function fmtDate(ts){try{var d=new Date(ts);var r=relDayLabel(d);return r||d.toLocaleDateString(undefined,{day:'numeric',month:'short',year:'numeric'})}catch(err){return ''}}
  function longDate(iso){try{var d=new Date(iso+'T12:00:00');if(isNaN(d))return iso;return d.toLocaleDateString(undefined,{weekday:'long',day:'numeric',month:'long',year:'numeric'})}catch(err){return iso}}
  function safeName(n){return String(n||'template').replace(/[\\/:*?"<>|]+/g,'-')}

  /* ---------------- save as template ---------------- */
  function defaultName(svc){return ((svc.name||'').trim()||(svc.tag||'').trim()||'Service')+' template'}
  function songPlaceholder(it){
    var ns=mkS('Song',['Song title'],0,6.5);
    ns.bg=it.bg;ns.dim=it.dim;
    if(it.customBg)ns.customBg=it.customBg;
    if(it.bgMedia)ns.bgMedia=cl(it.bgMedia);
    return ns;
  }
  function openTplSave(i){
    var svc=services[i];if(!svc){toast('No service to save.');return}
    tplSaveIdx=i;
    var body=serviceSyncBody(i);
    var hasSongs=body.slides.some(function(it){return it&&it.k==='Song'});
    var list=loadTemplates().slice().sort(byName);
    var sel=document.getElementById('tpl-save-target');
    sel.innerHTML='<option value="new">A new template</option>'+list.map(function(t){return '<option value="'+e(t.id)+'">Update \u201c'+e(t.name)+'\u201d</option>'}).join('');
    sel.value='new';
    document.getElementById('tpl-save-name').value=defaultName(svc);
    document.getElementById('tpl-save-src').textContent='From: '+formatServiceTitle(svc);
    document.getElementById('tpl-songs-box').hidden=!hasSongs;
    var r=document.querySelector('#tpl-save-modal input[name="tpl-songs"][value="strip"]');if(r)r.checked=true;
    document.getElementById('tpl-save-modal').hidden=false;
    setTimeout(function(){var n=document.getElementById('tpl-save-name');n.focus();n.select()},30);
  }
  function closeTplSave(){document.getElementById('tpl-save-modal').hidden=true;tplSaveIdx=-1}
  async function confirmSaveTemplate(){
    var i=tplSaveIdx,svc=services[i];if(!svc){closeTplSave();return}
    var name=document.getElementById('tpl-save-name').value.trim();
    if(!name){toast('Please enter a template name.');return}
    var target=document.getElementById('tpl-save-target').value;
    var list=loadTemplates();
    var dup=list.filter(function(t){return t.name.trim().toLowerCase()===name.toLowerCase()&&t.id!==target})[0];
    if(dup){
      if(!await appConfirm('A template named "'+dup.name+'" already exists.\n\nReplace it with this service?',{title:'Replace template',confirmLabel:'Replace'}))return;
      target=dup.id;
    }
    var body=serviceSyncBody(i);
    var strip=!document.getElementById('tpl-songs-box').hidden&&(document.querySelector('#tpl-save-modal input[name="tpl-songs"]:checked')||{}).value==='strip';
    if(strip)body.slides=body.slides.map(function(it){return it&&it.k==='Song'?songPlaceholder(it):it});
    var now=Date.now(),t=null;
    if(target!=='new')t=list.filter(function(x){return x.id===target})[0]||null;
    if(t){t.name=name;t.team=body.team;t.slides=body.slides;t.updatedAt=now;t.dirty=true;t.localOnly=false}
    else{t={id:newTplId(),name:name,team:body.team,slides:body.slides,updatedAt:now,dirty:true,localOnly:false,driveFileId:''};list.push(t)}
    if(!saveTemplates(list))return;
    closeTplSave();
    renderTemplatesHome();refreshTemplateSelect();
    toast(cloudReady()?'Saved template "'+name+'". Syncing to Drive\u2026':'Saved template "'+name+'" on this computer. Connect Google Drive to share it.');
    syncTemplates({force:true});
  }

  /* ---------------- Drive sync ---------------- */
  function tplRemoteChanged(t,f){
    var r=remoteRev(f);
    if(r&&t.driveRev)return r!==t.driveRev;
    return (f.modifiedTime||'')!==(t.driveModified||'');
  }
  async function uploadTemplate(t,folderId){
    var json=JSON.stringify({format:'church-presenter-template',version:CPRES_SCHEMA,exportedAt:new Date().toISOString(),service:{name:t.name,date:'',tag:'',team:t.team,slides:t.slides}},null,2);
    var data=bytesToB64(new TextEncoder().encode(json));
    function args(existing){return {name:safeName(t.name)+'.cpres',dataBase64:data,mimeType:'application/json',folderId:existing?null:folderId,existingFileId:existing||null}}
    var res;
    try{res=await cloudCall('gdrive_upload_file',args(t.driveFileId))}
    catch(err){
      if(t.driveFileId&&err.kind==='notfound')res=await cloudCall('gdrive_upload_file',args(null));
      else throw err;
    }
    t.driveFileId=res.id;t.driveModified=res.modifiedTime||'';t.driveRev=remoteRev(res);t.dirty=false;t.localOnly=false;
  }
  async function syncTemplates(opts){
    opts=opts||{};
    if(!cloudReady()||live||syncBusy||tplSyncing)return;
    if(!opts.force&&Date.now()-tplLastSync<TPL_SYNC_MS)return;
    tplSyncing=true;tplLastSync=Date.now();
    try{
      await withCloudLock('Syncing templates\u2026',async function(){
        var folder=await cloudCall('gdrive_ensure_subfolder',{parentId:cloudUi.folderId,name:'templates'});
        if(!folder.id)throw cloudErr('Could not open the templates folder.');
        /* finish deletions made while offline */
        var tombs=loadTombs(),keep=[];
        for(var ti=0;ti<tombs.length;ti++){
          var dr=await cloudInvoke('gdrive_delete_file',{fileId:tombs[ti]});
          if(dr&&dr.ok===false&&classifyErr(dr.error)!=='notfound')keep.push(tombs[ti]);
        }
        saveTombs(keep);
        var listing=await cloudCall('gdrive_list_presentations',{folderId:folder.id});
        var files=(listing.files||[]).filter(function(f){return /\.cpres$/i.test(f.name||'')&&keep.indexOf(f.id)<0});
        var byId={};files.forEach(function(f){byId[f.id]=f});
        var list=loadTemplates(),changed=false,i,t;
        /* removed from Drive elsewhere: keep the local copy, just unlink it */
        for(i=0;i<list.length;i++){
          t=list[i];
          if(t.driveFileId&&!byId[t.driveFileId]&&!t.dirty){t.driveFileId='';t.localOnly=true;changed=true}
        }
        for(i=0;i<list.length;i++){
          t=list[i];
          if(t.dirty||(!t.driveFileId&&!t.localOnly)){await uploadTemplate(t,folder.id);changed=true}
          else if(t.driveFileId&&byId[t.driveFileId]&&tplRemoteChanged(t,byId[t.driveFileId])){
            var f=byId[t.driveFileId];
            try{
              var d=await readCpres(await downloadDriveFile({id:f.id,name:f.name,mimeType:f.mimeType||''}));
              t.name=d.name||t.name;t.team=d.team;t.slides=d.slides;
              t.updatedAt=Date.parse(f.modifiedTime)||Date.now();t.driveModified=f.modifiedTime||'';t.driveRev=remoteRev(f);changed=true;
            }catch(err){console.warn('[templates] could not update',t.name,err&&err.message)}
          }
        }
        var have={};list.forEach(function(x){if(x.driveFileId)have[x.driveFileId]=1});
        for(var k=0;k<files.length;k++){
          var rf=files[k];if(have[rf.id])continue;
          try{
            var rd=await readCpres(await downloadDriveFile({id:rf.id,name:rf.name,mimeType:rf.mimeType||''}));
            list.push({id:newTplId(),name:rd.name||rf.name.replace(/\.cpres$/i,''),team:rd.team,slides:rd.slides,updatedAt:Date.parse(rf.modifiedTime)||Date.now(),driveFileId:rf.id,driveModified:rf.modifiedTime||'',driveRev:remoteRev(rf),dirty:false,localOnly:false});
            changed=true;
          }catch(err){console.warn('[templates] skipped',rf.name,err&&err.message)}
        }
        if(changed)saveTemplates(list);
      },{quiet:opts.quiet!==false});
    }finally{tplSyncing=false}
    try{renderTemplatesHome();refreshTemplateSelect()}catch(err){}
  }

  /* ---------------- rename / delete ---------------- */
  async function renameTemplate(id){
    var list=loadTemplates(),t=list.filter(function(x){return x.id===id})[0];if(!t)return;
    var n=await appPrompt('Rename template','Enter a name for this template.',t.name);if(n==null)return;
    n=n.trim();if(!n||n===t.name)return;
    t.name=n;t.updatedAt=Date.now();t.dirty=true;t.localOnly=false;
    if(!saveTemplates(list))return;
    renderTemplatesHome();refreshTemplateSelect();
    syncTemplates({force:true});
  }
  async function deleteTemplate(id){
    var list=loadTemplates(),t=list.filter(function(x){return x.id===id})[0];if(!t)return;
    if(!await appConfirm('Delete the template "'+t.name+'"?'+(t.driveFileId?'\n\nIt will also be removed from Google Drive and from your other computers.':''),{title:'Delete template',confirmLabel:'Delete',danger:true}))return;
    if(!saveTemplates(list.filter(function(x){return x.id!==id})))return;
    if(t.driveFileId){var tb=loadTombs();tb.push(t.driveFileId);saveTombs(tb)}
    renderTemplatesHome();refreshTemplateSelect();
    toast('Template deleted.');
    syncTemplates({force:true});
  }

  /* ---------------- Services page: Templates section ---------------- */
  /* renderTemplatesHome lives at module level (default template first). */

  /* ---------------- New service form: Start from ---------------- */
  /* refreshTemplateSelect / updateTplNote live at module level (shared). */
  function useTemplate(id){
    openNewServiceForm();
    refreshTemplateSelect(id);
  }

  function createFromTemplate(tpl,form){
    var slides=cl(tpl.slides||[]).map(function(it){return remapServiceItemIds(it)});
    var svc={name:form.name,date:form.date,tag:form.tag,slides:slides.length?slides:[mkS('Welcome',['Welcome'],0,6.5)],team:tpl.team?cl(tpl.team):defaultTeam(),lastEdited:Date.now()};
    ensureTeam(svc);
    services.push(svc);activeService=services.length-1;
    S=cl(svc.slides);sel=0;si=0;A=S[0]||null;aj=0;
    saveState();
    var f=document.getElementById('service-form');f.classList.remove('open');f.reset();
    document.body.classList.remove('home-mode');
    rAll();
    toast('Created "'+formatServiceTitle(svc)+'" from the "'+tpl.name+'" template.');
  }
  /* Intercept the New service form when a template is chosen */
  document.addEventListener('submit',function(ev){
    var f=ev.target;if(!f||f.id!=='service-form')return;
    var sel=document.getElementById('service-template');
    if(!sel||!sel.value)return;
    var tpl=loadTemplates().filter(function(t){return t.id===sel.value})[0];
    if(!tpl)return;
    ev.preventDefault();ev.stopImmediatePropagation();
    var tag=(document.getElementById('service-tag').value||'').trim();
    if(!tag){toast('Choose a service tag.');document.getElementById('service-tag').focus();return}
    var name=(document.getElementById('service-name').value||'').trim();
    var date=document.getElementById('service-date').value;
    if(!date){toast('Choose a date.');document.getElementById('service-date').focus();return}
    var form={tag:tag,name:name,date:date};
    createFromTemplate(tpl,form);
  },true);

  /* ---------------- right-click menu on services ---------------- */
  function hideSvcMenu(){var m=document.getElementById('svc-ctx');if(m){m.hidden=true;m.classList.remove('open')}}
  function showSvcMenu(x,y,idx,home){
    var m=document.getElementById('svc-ctx');
    if(!m){
      m=document.createElement('div');
      m.id='svc-ctx';
      m.className='outline-menu';
      m.setAttribute('role','menu');
      document.body.appendChild(m);
    }
    m.innerHTML=(home?'<button type="button" data-svc-action="open">Open</button>':'')
      +'<button type="button" data-svc-action="save">Make a template</button>'
      +(home?'<button type="button" data-svc-action="delete" class="danger" style="color:#c22">Delete</button>':'');
    m.dataset.idx=String(idx);
    m.hidden=false;m.classList.add('open');
    var w=m.offsetWidth||180,h=m.offsetHeight||100;
    m.style.left=Math.max(4,Math.min(x,window.innerWidth-w-8))+'px';
    m.style.top=Math.max(4,Math.min(y,window.innerHeight-h-8))+'px';
  }
  document.addEventListener('contextmenu',function(ev){
    var t=ev.target;if(!t||!t.closest)return;
    if(live||document.body.classList.contains('presentation-mode'))return;
    var idx=-1,home=false;
    var card=t.closest('#service-list .service-card');
    if(card){
      var ob=card.querySelector('[data-open-service]');
      if(!ob)return;
      idx=+ob.dataset.openService;home=true;
    }else if(t.closest('#service-title-edit'))idx=activeService;
    else return;
    if(!services[idx])return;
    ev.preventDefault();
    ev.stopPropagation();
    showSvcMenu(ev.clientX,ev.clientY,idx,home);
  },true);

  /* ---------------- clicks, changes, keys ---------------- */
  document.addEventListener('click',function(ev){
    var t=ev.target;if(!t||!t.closest)return;
    if(t.closest('#share-template-btn')){ev.preventDefault();ev.stopImmediatePropagation();closeShareModal();openTplSave(activeService)}
  },true);

  document.addEventListener('click',function(ev){
    var t=ev.target;if(!t||!t.closest)return;
    var menu=document.getElementById('svc-ctx');
    var mb=t.closest('#svc-ctx [data-svc-action]');
    if(mb){
      var act=mb.dataset.svcAction,idx=+menu.dataset.idx;
      hideSvcMenu();
      if(act==='save')openTplSave(idx);
      else if(act==='open'){var ob=document.querySelector('#service-list [data-open-service="'+idx+'"]');if(ob)ob.click()}
      else if(act==='delete'){var db=document.querySelector('#service-list [data-del-service="'+idx+'"]');if(db)db.click()}
      return;
    }
    if(menu&&!menu.hidden&&!t.closest('#svc-ctx'))hideSvcMenu();
    var b;
    if((b=t.closest('[data-tpl-use]'))){useTemplate(b.dataset.tplUse);return}
    if((b=t.closest('[data-tpl-rename]'))){renameTemplate(b.dataset.tplRename);return}
    if((b=t.closest('[data-tpl-del]'))){deleteTemplate(b.dataset.tplDel);return}
    if(t.closest('#tpl-save-confirm')){confirmSaveTemplate();return}
    if(t.closest('#tpl-save-cancel')||t.closest('#tpl-save-close')||t.id==='tpl-save-modal'){closeTplSave();return}
    if(t.closest('#new-service')||t.closest('#make-new-service')){setTimeout(function(){refreshTemplateSelect(null,true)},0)}
  });

  document.addEventListener('change',function(ev){
    var t=ev.target;if(!t)return;
    if(t.id==='service-template'){updateTplNote();return}
    if(t.id==='tpl-save-target'){
      var svc=services[tplSaveIdx];
      if(t.value==='new')document.getElementById('tpl-save-name').value=svc?defaultName(svc):'';
      else{
        var tp=loadTemplates().filter(function(x){return x.id===t.value})[0];
        if(tp)document.getElementById('tpl-save-name').value=tp.name;
      }
    }
  });

  document.addEventListener('keydown',function(ev){
    if(ev.key==='Escape'){
      var sm=document.getElementById('svc-ctx');
      if(sm&&!sm.hidden){hideSvcMenu();return}
      var sv=document.getElementById('tpl-save-modal');
      if(sv&&!sv.hidden){ev.preventDefault();closeTplSave();return}
      return;
    }
    if(ev.key==='Enter'){
      if(ev.target&&ev.target.id==='tpl-save-name'){ev.preventDefault();confirmSaveTemplate()}
    }
  });
  window.addEventListener('blur',hideSvcMenu);

  /* ---------------- hook the Services page ---------------- */
  var _rHomeV4=rHome;
  rHome=function(){
    _rHomeV4.apply(this,arguments);
    try{renderTemplatesHome()}catch(err){console.warn('[templates]',err)}
    try{refreshTemplateSelect()}catch(err){}
    if(cloudReady()&&!syncBusy&&!tplSyncing&&Date.now()-tplLastSync>TPL_SYNC_MS)setTimeout(function(){syncTemplates()},120);
  };
  refreshTemplateSelect();
  if(document.body.classList.contains('home-mode'))rHome();

  setTimeout(function(){syncTemplates({force:true})},4000);
  setInterval(function(){syncTemplates()},60000);
})();}

/* mjs-patch-v5 */
if(!IS_AUDIENCE){(function(){
  var css=document.createElement('style');
  css.textContent='#share-template-sep{border-top:1px solid var(--ln);margin:18px 0 14px;height:0}';
  document.head.appendChild(css);
  var imp=document.getElementById('share-import-btn'),tpl=document.getElementById('share-template-btn');
  if(!imp||!tpl)return;
  var sep=document.createElement('div');
  sep.id='share-template-sep';
  sep.setAttribute('role','separator');
  imp.insertAdjacentElement('afterend',sep);
  sep.insertAdjacentElement('afterend',tpl);
})();}


// Duplicate UI utilities block removed – originals retained earlier.

// Duplicated UI/date/template utilities removed – originals retained earlier.



// Duplicate UI utilities and template helpers removed – originals are defined earlier in the file.

// Duplicate relDayLabelISO removed – original defined earlier.

// Duplicate UI utilities block removed – originals retained earlier.
// Removed duplicate declaration – using the original const at the top of the file
// Duplicate defaultTplId & isDefaultTpl removed – originals retained earlier.

if(!IS_AUDIENCE){(function(){
  /* ---------- styles ---------- */
  var st=document.createElement('style');
  st.textContent=[
    '.service-title-wrap{display:inline-flex;align-items:center;gap:2px;min-width:0}',
    '.svc-tag-btn{display:inline-flex;align-items:center;gap:4px;padding:4px 8px;border:1px solid transparent;border-radius:6px;background:transparent;font-weight:700;color:var(--ink)}',
    '.svc-tag-btn:hover,.svc-tag-btn:focus-visible{background:var(--bg);border-color:var(--ln)}',
    '.svc-caret{font-size:calc(10px * var(--ui-scale));color:var(--mu)}',
    '.svc-sep{color:var(--mu);padding:0 2px}.svc-sep[hidden]{display:none}',
    '.service-title .svc:empty{display:none}',
    '.service-title{gap:0}',
    '.svc-title-row{display:flex;align-items:center;gap:2px;flex-wrap:wrap}',
    '.tpl-name{display:flex;align-items:center;flex-wrap:wrap}',
    '.tpl-default-badge{margin-left:8px;font-size:calc(11px * var(--ui-scale));font-weight:700;padding:2px 8px;border-radius:999px;background:var(--hl);color:var(--ink);border:1px solid var(--ln)}',
    '#tag-menu{z-index:1400}',
    'kbd{font-size:calc(12px * var(--ui-scale))}'
  ].join('\n');
  document.head.appendChild(st);

  /* ---------- "last edited" only changes when content really changes ---------- */
  var edHash=new WeakMap();
  var _origTouch=touchServiceEdited;
  touchServiceEdited=function(){
    var svc=services[activeService];if(!svc)return;
    var h;
    try{h=hashService(activeService)}catch(err){return _origTouch()}
    var prev=edHash.get(svc);
    edHash.set(svc,h);
    if(prev===undefined)return;
    if(prev!==h)_origTouch();
  };
  try{var s0=services[activeService];if(s0)edHash.set(s0,hashService(activeService))}catch(err){}

  /* ---------- tags: keep to the global list; auto upload brand new services ---------- */
  function normTags(){
    services.forEach(function(s){
      var t=(s.tag||'').trim();
      if(t&&serviceTagList().indexOf(t)<0)s.tag='Other';
    });
  }
  var knownSvc=new WeakSet();
  services.forEach(function(s){knownSvc.add(s)});
  function queueNewServiceUpload(svc,n){
    setTimeout(function(){
      var i=services.indexOf(svc);
      if(i<0||getLink(svc))return;
      if(loadCloudMeta().autoUpload===false)return;
      if(!cloudReady())return;
      if(live||syncBusy){if(n<15)queueNewServiceUpload(svc,n+1);return}
      withCloudLock('Saving to Drive\u2026',function(){return cloudPush(i)},{quiet:true});
    },n?2000:1200);
  }
  var _ss6=saveState;
  saveState=function(){
    try{normTags()}catch(err){}
    var r=_ss6.apply(this,arguments);
    try{
      services.forEach(function(s){
        if(knownSvc.has(s))return;
        knownSvc.add(s);
        if(getLink(s))return;
        ensureSvcId(s);
        seenHash[s.id]='__new__';
        queueNewServiceUpload(s,0);
      });
    }catch(err){console.warn('[v6] new service upload',err)}
    return r;
  };
  normTags();

  /* ---------- church logo slide: becomes a normal editable slide ---------- */
  var logoMediaCache={};
  try{logoMediaCache=JSON.parse(localStorage.getItem('church-presenter-logo-media')||'{}')||{}}catch(err){logoMediaCache={}}
  function dataUrlToBlob(u){
    var comma=u.indexOf(','),head=u.slice(5,comma),body=u.slice(comma+1);
    var b64=/;base64$/i.test(head);
    var mime=head.replace(/;base64$/i,'').split(';')[0]||'application/octet-stream';
    if(b64){var bin=atob(body),arr=new Uint8Array(bin.length);for(var i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);return new Blob([arr],{type:mime})}
    return new Blob([decodeURIComponent(body)],{type:mime});
  }
  function imgSize(src){
    return new Promise(function(res){
      var im=new Image();
      im.onload=function(){res({w:im.naturalWidth||0,h:im.naturalHeight||0})};
      im.onerror=function(){res({w:0,h:0})};
      im.src=src;
    });
  }
  async function ensureLogoMedia(src){
    var key=hashStr(src),id=logoMediaCache[key];
    if(id&&mediaURLs[id])return id;
    var blob=src.indexOf('data:')===0?dataUrlToBlob(src):await (await fetch(src)).blob();
    id='m'+Date.now().toString(36)+Math.random().toString(36).slice(2,8);
    await mPut(id,blob);
    mediaURLs[id]=URL.createObjectURL(blob);
    var l=churchLogos.find(function(x){return x.id===effectiveLogoId()});
    mediaMeta.push({id:id,name:'Church logo'+(l?' - '+l.name:''),kind:'image',size:blob.size,createdAt:Date.now(),source:'logo'});
    saveMediaMeta();
    logoMediaCache[key]=id;
    try{localStorage.setItem('church-presenter-logo-media',JSON.stringify(logoMediaCache))}catch(err){}
    return id;
  }
  addChurchLogoSection=async function(target){
    try{syncEffectiveLogo()}catch(err){}
    var src=churchLogo||'',size=churchLogoSize||80,els=[];
    if(src){
      var mid=await ensureLogoMedia(src);
      var d=await imgSize(src);
      var nr=d.w&&d.h?d.w/d.h:0,w,h,fit='fill';
      if(nr){
        var kw=size/100*16,kh=size/100*9,dw=kw,dh=kw/nr;
        if(dh>kh){dh=kh;dw=kh*nr}
        w=dw/16*100;h=dh/9*100;
      }else{w=size;h=size;fit='contain'}
      var x=churchLogoHPos==='flex-start'?0:churchLogoHPos==='flex-end'?100-w:(100-w)/2;
      var y=churchLogoVPos==='flex-start'?0:churchLogoVPos==='flex-end'?100-h:(100-h)/2;
      var el={id:uid(),kind:'image',src:mediaURLs[mid],mediaId:mid,x:x,y:y,w:w,h:h,rot:0,z:1,fit:fit};
      if(nr)el.nr=nr;
      els.push(el);
    }else toast('No logo uploaded yet. Add one in Settings, Church logo.');
    pushUndo();
    var pickIdx=BG.findIndex(function(b){return b.pick});
    var tgtIt=(target&&typeof target.i==='number')?S[target.i]:null;
    if(tgtIt&&!isSongItem(tgtIt)){
      // slide inside the existing section, with its own logo-coloured background
      var part2={id:uid(),n:'Slide '+(tgtIt.parts.length+1),tx:'',elements:els,
        ownBg:true,bg:pickIdx>=0?pickIdx:3,customBg:logoBgCss(),dim:0};
      tgtIt.parts.push(part2);
      var at=Math.min(tgtIt.ord.length,Math.max(0,target.after+1));
      insOrd(tgtIt,at,part2.id);
      sel=target.i;A=tgtIt;si=at;aj=at;
    }else{
      var part={id:uid(),n:'Slide 1',tx:'',elements:els};
      var item=mk('Church logo','Slide',[part],[part.id],pickIdx>=0?pickIdx:3,5.2);
      item.customBg=logoBgCss();item.dim=0;
      var insAt=(target&&target.newSectionAt!=null)?Math.min(target.newSectionAt,S.length):Math.max(0,sel)+1;
      S.splice(insAt,0,item);
      sel=S.indexOf(item);A=item;si=0;aj=0;
    }
    clr=null;elSel=null;tab='Structure';
    if(services[activeService])services[activeService].slides=sanitizeSlidesForSave(S);
    saveState();rAll();
    setTimeout(function(){scrollNewItemToTop(sel);scrollCurrentSlideIntoView()},0);
    if(src)toast('Added logo slide. It is a normal slide now, so you can edit it freely.');
  };

  /* ---------- copy / paste: always paste what was copied last ---------- */
  var EL_CLIP_MARK='[church-presenter-element]';
  function writeElementClip(){
    var clip=canvasClip;if(!clip)return;
    var plain=clip.kind==='text'?String(clip.tx||''):'Church Presenter element';
    var handled=false;
    var prevFocus=document.activeElement;
    var ta=document.createElement('textarea');
    ta.value=plain||' ';
    ta.style.cssText='position:fixed;left:-9999px;top:0;opacity:0';
    document.body.appendChild(ta);
    ta.select();
    var h=function(ev){
      ev.preventDefault();
      ev.clipboardData.setData('text/plain',plain||' ');
      ev.clipboardData.setData('text/html','<span data-church-presenter-el="1">'+e(plain)+'</span>');
      handled=true;
    };
    document.addEventListener('copy',h,true);
    try{document.execCommand('copy')}catch(err){}
    document.removeEventListener('copy',h,true);
    ta.remove();
    try{if(prevFocus&&prevFocus!==document.body&&prevFocus.focus)prevFocus.focus()}catch(err){}
    if(!handled&&navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(EL_CLIP_MARK).catch(function(){})}
  }
  var _copySel=copySelectedEl;
  copySelectedEl=function(){
    _copySel.apply(this,arguments);
    try{if(canvasClip)writeElementClip()}catch(err){console.warn('[v6] clipboard',err)}
  };
  document.addEventListener('paste',function(ev){
    if(live||!A||isSongItem(A)||!canvasClip)return;
    if(isTypingTarget(ev.target))return;
    var cd=ev.clipboardData;if(!cd)return;
    var htmlData='',txt='';
    try{htmlData=cd.getData('text/html')||'';txt=cd.getData('text/plain')||''}catch(err){}
    if(htmlData.indexOf('data-church-presenter-el')>=0||txt===EL_CLIP_MARK){
      ev.preventDefault();
      ev.stopImmediatePropagation();
      pasteCanvasEl();
    }
  },true);

  /* ---------- tag menu keys ---------- */
  /* The click handling for the tag menu lives in the global handler (section 10).
     A second copy of it here made every click open the menu and immediately close it again. */
  document.addEventListener('keydown',function(ev){if(ev.key==='Escape')hideTagMenu()});
  window.addEventListener('blur',hideTagMenu);


  document.addEventListener('click',function(ev){
    var t=ev.target;if(!t||!t.closest)return;
    if(!(t.closest('#new-service')||t.closest('#make-new-service')))return;
    setTimeout(function(){
      var id=defaultTplId(),sel2=document.getElementById('service-template');
      if(!id||!sel2)return;
      if([].some.call(sel2.options,function(o){return o.value===id})){
        sel2.value=id;
        sel2.dispatchEvent(new Event('change',{bubbles:true}));
      }
    },60);
  });

  try{renderTitleBar(services[activeService])}catch(err){}
  try{if(document.body.classList.contains('home-mode'))rHome()}catch(err){}
})();}


/* mjs-patch-church-presenter-safe */
if(typeof IS_AUDIENCE !== 'undefined' && !IS_AUDIENCE){(function(){
  'use strict';
  try {

  /* ---- CSS: single border on slide preview only ---- */
  var old = document.getElementById('patch-ui-fix-css');
  if(old) old.remove();
  var old2 = document.getElementById('patch-cpa-css');
  if(old2) old2.remove();
  var css = document.createElement('style');
  css.id = 'patch-safe-css';
  css.textContent = [
    '#strip .th{',
    '  border:none!important;',
    '  box-shadow:none!important;',
    '  background:transparent!important;',
    '  outline:none!important;',
    '  padding:0!important;',
    '}',
    '#strip .th .sl{',
    '  outline:none!important;',
    '  border:2px solid #9ca3af!important;',
    '  border-radius:4px;',
    '  box-sizing:border-box;',
    '}',
    ':root[data-theme="dark"] #strip .th .sl{border-color:#8b949e!important;}',
    '#strip .th.on .sl, #strip .th.on.es .sl{',
    '  border-color:#ffd400!important;',
    '  border-width:3px!important;',
    '}',
    '#strip .th .pn{',
    '  display:block;',
    '  padding:4px 2px 2px;',
    '  border:none!important;',
    '  outline:none!important;',
    '  background:transparent!important;',
    '}',
    '.svc-sep, #service-title-sep{display:none!important;}'
  ].join('\n');
  document.head.appendChild(css);

  /* ---- Import contacts label (no MutationObserver — that froze the app) ---- */
  function forceImportContactsLabel(){
    try{
      ['cm-import-txt','cm-import-csv','cm-import-vcf'].forEach(function(id){
        var btn = document.getElementById(id);
        if(!btn) return;
        if(btn.id !== 'cm-import-txt') btn.id = 'cm-import-txt';
        if(btn.textContent !== 'Import contacts') btn.textContent = 'Import contacts';
        btn.title = 'Import a plain-text contact directory';
      });
      var file = document.getElementById('cm-txt-file')
              || document.getElementById('cm-csv-file')
              || document.getElementById('cm-vcf-file');
      if(file){
        if(file.id !== 'cm-txt-file') file.id = 'cm-txt-file';
        file.accept = '.txt,text/plain';
      }
    }catch(err){}
  }

  document.addEventListener('click', async function(ev){
    var t = ev.target;
    if(!t) return;
    try{
      if(t.id === 'cm-import-txt' || t.id === 'cm-import-csv' || t.id === 'cm-import-vcf'){
        setTimeout(forceImportContactsLabel, 0);
      }
      if(t.closest && t.closest('[data-settings-page="members"]')){
        setTimeout(forceImportContactsLabel, 0);
      }
      // Team top button click is bound below
    }catch(err){}
  }, true);

  document.addEventListener('change', function(ev){
    var t = ev.target;
    if(!t) return;
    if(t.id === 'cm-txt-file' || t.id === 'cm-csv-file' || t.id === 'cm-vcf-file'){
      setTimeout(forceImportContactsLabel, 50);
    }
  }, true);

  /* ---- Service Team top-bar button ---- */
  function ensureTeamTopBtn(){
    try{
      if(document.getElementById('team-top-btn')) return;
      var share = document.getElementById('export-cpres-btn');
      var cloud = document.getElementById('cloud-btn');
      if(!share && !cloud) return;

      var btn = document.createElement('button');
      btn.type = 'button';
      btn.id = 'team-top-btn';
      btn.className = 'icon-btn';
      btn.setAttribute('aria-label', 'Service team');
      btn.title = 'Service team';
      btn.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>';

      if(share && share.parentNode) share.parentNode.insertBefore(btn, share);
      else if(cloud && cloud.parentNode) cloud.parentNode.insertBefore(btn, cloud.nextSibling);

      btn.addEventListener('click', function(ev){
        ev.preventDefault();
        try{
          if(typeof openTeamModal === 'function') openTeamModal();
          else {
            var m = document.getElementById('team-modal');
            if(m) m.hidden = false;
          }
        }catch(err){ console.warn(err); }
      });
    }catch(err){ console.warn('[safe] team btn', err); }
  }

  function boot(){
    try{ forceImportContactsLabel(); }catch(e){}
    try{ ensureTeamTopBtn(); }catch(e){}
  }
  if(document.readyState === 'loading'){
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
  // Gentle retry once (no tight loops)
  setTimeout(boot, 800);
  setTimeout(boot, 2500);

  console.info('[church-presenter] safe UI patch loaded');
  } catch(err){
    console.error('[church-presenter] safe patch failed', err);
  }
})();}


/* mjs-patch-church-members */
if(typeof IS_AUDIENCE !== 'undefined' && !IS_AUDIENCE){(function(){
  'use strict';
  try {

  var MEMBERS_KEY = 'service-slides-church-members';
  var cmSelected = new Set();
  var cmAnchorId = null;

  function loadMembers(){
    try{
      var a = JSON.parse(localStorage.getItem(MEMBERS_KEY) || '[]');
      return Array.isArray(a) ? a : [];
    }catch(e){ return []; }
  }
  function saveMembers(list){
    try{ localStorage.setItem(MEMBERS_KEY, JSON.stringify(list)); return true; }
    catch(e){ try{ toast('Could not save members (storage full).'); }catch(x){} return false; }
  }
  function newMemberId(){
    return 'mbr-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  }
  function isEmailLike(s){
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '').trim());
  }
  function isPhoneLike(s){
    var t = String(s || '').trim();
    if(!t) return false;
    var digits = t.replace(/[\s\-\.\(\)\+\/]/g, '');
    if(!/^\d{6,15}$/.test(digits)) return false;
    return ((t.match(/[A-Za-z]/g) || []).length) < 2;
  }
  function esc(s){
    if(typeof e === 'function') return e(s);
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  /** Plain-text contact directory: Name, then Phone:/Email: lines. */
  function parseContactsTxt(text){
    var lines = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    var out = [];
    var cur = null;
    var detailsStarted = false;

    function flush(){
      if(!cur) return;
      var name = String(cur.name || '').trim();
      var email = String(cur.email || '').trim();
      if(isPhoneLike(name)) name = '';
      if(isEmailLike(name)){
        if(!email) email = name;
        name = '';
      }
      if(/^contact directory$/i.test(name)) name = '';
      if(/^email\s*:/i.test(name)) name = '';
      if(name || email) out.push({name: name, email: email});
      cur = null;
      detailsStarted = false;
    }

    lines.forEach(function(raw){
      var line = String(raw || '').trim();
      if(!line) return;

      if(/^contact directory$/i.test(line)) return;
      if(/^no contact details/i.test(line)){
        if(cur) detailsStarted = true;
        return;
      }

      var mEmail = /^(?:e-?mail)\s*[:：]\s*(.+)$/i.exec(line);
      if(mEmail){
        var em = mEmail[1].trim();
        if(!cur) cur = {name: '', email: ''};
        if(isEmailLike(em) && !cur.email) cur.email = em;
        detailsStarted = true;
        return;
      }

      if(/^(?:phone|mobile|tel|cell)\s*[:：]/i.test(line)){
        if(!cur) cur = {name: '', email: ''};
        detailsStarted = true;
        return;
      }

      if(isEmailLike(line)){
        if(!cur) cur = {name: '', email: ''};
        if(!cur.email) cur.email = line;
        detailsStarted = true;
        return;
      }

      // Org / note under current person before details
      if(cur && cur.name && !detailsStarted) return;

      flush();
      cur = {name: line, email: ''};
      detailsStarted = false;
    });
    flush();
    return out;
  }

  function importMembersFromTxt(file){
    return new Promise(function(resolve, reject){
      var reader = new FileReader();
      reader.onload = function(){
        try{ resolve(parseContactsTxt(String(reader.result || ''))); }
        catch(err){ reject(err); }
      };
      reader.onerror = function(){ reject(new Error('Could not read file')); };
      reader.readAsText(file);
    });
  }

  function ensureMembersCss(){
    if(document.getElementById('patch-members-css')) return;
    var css = document.createElement('style');
    css.id = 'patch-members-css';
    css.textContent = [
      '.cm-list{display:flex;flex-direction:column;gap:8px;margin:12px 0}',
      '.cm-row{display:grid;grid-template-columns:auto 1fr 1fr auto;gap:8px;align-items:center;padding:8px;border:1px solid var(--ln);border-radius:8px;background:var(--bg)}',
      '.cm-row.cm-sel{border-color:var(--ac);background:var(--hl);box-shadow:inset 3px 0 0 var(--ac)}',
      '.cm-row input[type=checkbox]{width:auto;margin:0;cursor:pointer}',
      '.cm-row input[type=text],.cm-row input[type=email]{margin:0}',
      '.cm-actions{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0;align-items:center}',
      '.cm-actions .cm-danger{color:#c22}',
      '.cm-suggest{position:absolute;z-index:80;min-width:220px;max-width:min(360px,90vw);max-height:200px;overflow:auto;background:var(--pn);border:1px solid var(--ln);border-radius:8px;box-shadow:0 10px 28px #0004;padding:4px}',
      '.cm-suggest button{display:block;width:100%;text-align:left;border:0;background:none;padding:8px 12px;border-radius:6px;cursor:pointer}',
      '.cm-suggest button:hover,.cm-suggest button.active{background:var(--hl)}',
      '.cm-suggest small{display:block;color:var(--mu);font-weight:400}'
    ].join('\n');
    document.head.appendChild(css);
  }

  function ensureMembersPage(){
    ensureMembersCss();
    var nav = document.querySelector('.st-nav');
    var pages = document.querySelector('.st-pages');
    if(!nav || !pages) return false;

    if(!document.querySelector('[data-settings-page="members"]')){
      var tab = document.createElement('button');
      tab.type = 'button';
      tab.className = 'st-tab';
      tab.dataset.settingsPage = 'members';
      tab.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg> Church Members';
      nav.appendChild(tab);
    }

    var page = document.querySelector('.st-page[data-page="members"]');
    if(!page){
      page = document.createElement('section');
      page.className = 'st-page';
      page.dataset.page = 'members';
      pages.appendChild(page);
    }

    // Only rebuild chrome if missing the import button
    if(!document.getElementById('cm-import-txt') || !document.getElementById('cm-list')){
      page.innerHTML = [
        '<h3>Church Members</h3>',
        '<p class="note st-lead">People available for autocomplete in the Service team window. Import a <b>.txt</b> contact directory (Name, then Phone/Email lines). Only <b>name</b> and <b>email</b> are kept.</p>',
        '<div class="cm-actions">',
        '  <button type="button" id="cm-add" class="pri">Add member</button>',
        '  <button type="button" id="cm-import-txt">Import contacts</button>',
        '  <button type="button" id="cm-delete-selected" class="cm-danger" hidden>Delete selected</button>',
        '  <button type="button" id="cm-delete-all" class="cm-danger">Delete all</button>',
        '  <input type="file" id="cm-txt-file" accept=".txt,text/plain" hidden>',
        '</div>',
        '<div id="cm-list" class="cm-list"></div>',
        '<p class="note" id="cm-count"></p>'
      ].join('');
    }
    return true;
  }

  function showMembersPage(){
    ensureMembersPage();
    document.querySelectorAll('.st-tab').forEach(function(x){
      x.classList.toggle('active', x.dataset.settingsPage === 'members');
    });
    document.querySelectorAll('.st-page').forEach(function(x){
      x.classList.toggle('active', x.dataset.page === 'members');
    });
    renderMembersPage();
  }

  function renderMembersPage(){
    ensureMembersPage();
    var list = loadMembers().slice().sort(function(a, b){
      return String(a.name || a.email || '').localeCompare(String(b.name || b.email || ''));
    });
    var alive = new Set(list.map(function(m){ return m.id; }));
    cmSelected.forEach(function(id){ if(!alive.has(id)) cmSelected.delete(id); });

    var box = document.getElementById('cm-list');
    var count = document.getElementById('cm-count');
    var delSel = document.getElementById('cm-delete-selected');
    if(!box) return;

    if(!list.length){
      box.innerHTML = '<p class="note">No members yet. Import a .txt contact directory or click Add member.</p>';
    } else {
      box.innerHTML = list.map(function(m){
        var sel = cmSelected.has(m.id);
        return '<div class="cm-row' + (sel ? ' cm-sel' : '') + '" data-cm-id="' + esc(m.id) + '">' +
          '<input type="checkbox" data-cm-check="' + esc(m.id) + '" ' + (sel ? 'checked' : '') + ' aria-label="Select">' +
          '<input type="text" data-cm-field="name" value="' + esc(m.name || '') + '" placeholder="Name" autocomplete="off">' +
          '<input type="email" data-cm-field="email" value="' + esc(m.email || '') + '" placeholder="Email" autocomplete="off">' +
          '<button type="button" data-cm-del="' + esc(m.id) + '" style="color:#c22" title="Delete">\u2715</button>' +
          '</div>';
      }).join('');
    }
    if(count){
      var n = list.length, s = cmSelected.size;
      count.textContent = n + ' member' + (n === 1 ? '' : 's') + (s ? ' \u00b7 ' + s + ' selected' : '');
    }
    if(delSel) delSel.hidden = cmSelected.size === 0;

    // Keep label stable
    var btn = document.getElementById('cm-import-txt') || document.getElementById('cm-import-vcf') || document.getElementById('cm-import-csv');
    if(btn){
      btn.id = 'cm-import-txt';
      btn.textContent = 'Import contacts';
    }
  }

  function selectCmRange(fromId, toId){
    var list = loadMembers().slice().sort(function(a, b){
      return String(a.name || a.email || '').localeCompare(String(b.name || b.email || ''));
    });
    var a = list.findIndex(function(m){ return m.id === fromId; });
    var b = list.findIndex(function(m){ return m.id === toId; });
    if(a < 0 || b < 0){ cmSelected.add(toId); renderMembersPage(); return; }
    if(a > b){ var t = a; a = b; b = t; }
    for(var i = a; i <= b; i++) cmSelected.add(list[i].id);
    renderMembersPage();
  }

  function doImport(file){
    importMembersFromTxt(file).then(function(parsed){
      parsed = (parsed || []).filter(function(p){
        return (p.name && p.name.trim()) || (p.email && p.email.trim());
      });
      if(!parsed.length){
        try{ toast('No contacts found. Expected: Name then Email: someone@\u2026'); }catch(e){}
        return;
      }
      var list = loadMembers();
      var byEmail = {}, byName = {};
      list.forEach(function(m){
        if(m.email) byEmail[m.email.toLowerCase()] = m;
        if(m.name) byName[m.name.toLowerCase()] = m;
      });
      var added = 0, updated = 0;
      parsed.forEach(function(p){
        var key = (p.email || '').toLowerCase();
        var nkey = (p.name || '').toLowerCase();
        var cur = (key && byEmail[key]) || (nkey && byName[nkey]) || null;
        if(cur){
          var changed = false;
          if(p.email && !cur.email){ cur.email = p.email; changed = true; }
          if(p.name && (!cur.name || isPhoneLike(cur.name))){ cur.name = p.name; changed = true; }
          if(changed) updated++;
        } else {
          var m = {id: newMemberId(), name: p.name || '', email: p.email || ''};
          list.push(m);
          if(key) byEmail[key] = m;
          if(nkey) byName[nkey] = m;
          added++;
        }
      });
      saveMembers(list);
      renderMembersPage();
      try{
        toast('Imported ' + added + ' new contact' + (added === 1 ? '' : 's') +
          (updated ? ', updated ' + updated : '') + '.');
      }catch(e){}
    }).catch(function(err){
      try{ toast('Import failed: ' + ((err && err.message) || err)); }catch(e){}
    });
  }

  document.addEventListener('click', async function(ev){
    var t = ev.target; if(!t || !t.closest) return;

    // Open members tab
    if(t.closest('[data-settings-page="members"]')){
      setTimeout(showMembersPage, 0);
      return;
    }
    if(t.id === 'settings-btn' || t.closest('#settings-btn')){
      setTimeout(function(){ try{ ensureMembersPage(); }catch(e){} }, 0);
      return;
    }

    if(t.id === 'cm-add'){
      var list = loadMembers();
      list.push({id: newMemberId(), name: '', email: ''});
      saveMembers(list);
      renderMembersPage();
      setTimeout(function(){
        var rows = document.querySelectorAll('#cm-list .cm-row');
        var last = rows[rows.length - 1];
        if(last){ var inp = last.querySelector('[data-cm-field="name"]'); if(inp) inp.focus(); }
      }, 30);
      return;
    }

    if(t.id === 'cm-import-txt' || t.id === 'cm-import-csv' || t.id === 'cm-import-vcf'){
      ev.preventDefault();
      ev.stopImmediatePropagation();
      var f = document.getElementById('cm-txt-file');
      if(!f){
        ensureMembersPage();
        f = document.getElementById('cm-txt-file');
      }
      if(f) f.click();
      return;
    }

    if(t.id === 'cm-delete-all'){
      var n = loadMembers().length;
      if(!n){ try{ toast('No members to delete.'); }catch(e){} return; }
      if(!await appConfirm('Delete all ' + n + ' church member' + (n === 1 ? '' : 's') + '?\n\nThis cannot be undone.',{title:'Delete all members',confirmLabel:'Delete all',danger:true})) return;
      saveMembers([]);
      cmSelected.clear();
      cmAnchorId = null;
      renderMembersPage();
      try{ toast('All members deleted.'); }catch(e){}
      return;
    }

    if(t.id === 'cm-delete-selected'){
      var s = cmSelected.size;
      if(!s) return;
      if(!await appConfirm('Delete ' + s + ' selected member' + (s === 1 ? '' : 's') + '?',{title:'Delete selected members',confirmLabel:'Delete',danger:true})) return;
      saveMembers(loadMembers().filter(function(m){ return !cmSelected.has(m.id); }));
      cmSelected.clear();
      cmAnchorId = null;
      renderMembersPage();
      try{ toast('Deleted ' + s + ' member' + (s === 1 ? '' : 's') + '.'); }catch(e){}
      return;
    }

    var del = t.closest('[data-cm-del]');
    if(del){
      var id = del.dataset.cmDel;
      if(!await appConfirm('Remove this member?',{title:'Remove member',confirmLabel:'Remove',danger:true})) return;
      saveMembers(loadMembers().filter(function(m){ return m.id !== id; }));
      cmSelected.delete(id);
      renderMembersPage();
      return;
    }

    var check = t.closest('[data-cm-check]');
    if(check){
      var cid = check.dataset.cmCheck;
      ev.preventDefault();
      if(ev.shiftKey && cmAnchorId){
        selectCmRange(cmAnchorId, cid);
      } else if(ev.metaKey || ev.ctrlKey){
        if(cmSelected.has(cid)) cmSelected.delete(cid); else cmSelected.add(cid);
        cmAnchorId = cid;
        renderMembersPage();
      } else {
        var only = cmSelected.size === 1 && cmSelected.has(cid);
        cmSelected.clear();
        if(!only) cmSelected.add(cid);
        cmAnchorId = cid;
        renderMembersPage();
      }
      return;
    }
  }, true);

  document.addEventListener('change', function(ev){
    var t = ev.target;
    if(!t) return;
    if(t.id === 'cm-txt-file' || t.id === 'cm-csv-file' || t.id === 'cm-vcf-file'){
      var file = t.files && t.files[0];
      t.value = '';
      if(file) doImport(file);
    }
  }, true);

  document.addEventListener('input', function(ev){
    var row = ev.target && ev.target.closest && ev.target.closest('.cm-row[data-cm-id]');
    if(!row) return;
    var field = ev.target.dataset.cmField;
    if(!field) return;
    var list = loadMembers();
    var m = list.find(function(x){ return x.id === row.dataset.cmId; });
    if(!m) return;
    m[field] = ev.target.value;
    saveMembers(list);
  });

  /* Team modal autocomplete */
  var suggestBox = null;
  function hideSuggest(){ if(suggestBox){ suggestBox.remove(); suggestBox = null; } }
  function filterMembers(q){
    q = String(q || '').trim().toLowerCase();
    if(!q) return [];
    return loadMembers().filter(function(m){
      return (m.name || '').toLowerCase().indexOf(q) >= 0 ||
             (m.email || '').toLowerCase().indexOf(q) >= 0;
    }).sort(function(a, b){ return String(a.name || '').localeCompare(String(b.name || '')); });
  }
  function applyMemberToInput(input, m){
    input.value = m.name || '';
    input.dispatchEvent(new Event('input', {bubbles: true}));
    input.dispatchEvent(new Event('change', {bubbles: true}));
    var row = input.closest('tr, .team-person, label, div');
    var emailInp = row && row.querySelector('input[type=email], input[placeholder*="email" i]');
    if(!emailInp){
      var all = Array.from(document.querySelectorAll('#team-content input, #team-modal input'));
      var ix = all.indexOf(input);
      for(var i = ix + 1; i < all.length && i < ix + 4; i++){
        if(/email/i.test(all[i].type + all[i].placeholder + all[i].id)){ emailInp = all[i]; break; }
      }
    }
    if(emailInp && m.email){
      emailInp.value = m.email;
      emailInp.dispatchEvent(new Event('input', {bubbles: true}));
      emailInp.dispatchEvent(new Event('change', {bubbles: true}));
    }
    try{ if(typeof saveState === 'function') saveState(); }catch(e){}
  }
  function showSuggest(input, matches){
    hideSuggest();
    if(!matches.length) return;
    suggestBox = document.createElement('div');
    suggestBox.className = 'cm-suggest';
    matches.slice(0, 12).forEach(function(m, i){
      var b = document.createElement('button');
      b.type = 'button';
      if(i === 0) b.classList.add('active');
      b.innerHTML = '<span>' + esc(m.name || '(no name)') + '</span>' +
        (m.email ? '<small>' + esc(m.email) + '</small>' : '');
      b.addEventListener('mousedown', function(ev){
        ev.preventDefault();
        applyMemberToInput(input, m);
        hideSuggest();
      });
      suggestBox.appendChild(b);
    });
    document.body.appendChild(suggestBox);
    var r = input.getBoundingClientRect();
    suggestBox.style.left = Math.max(6, r.left) + 'px';
    suggestBox.style.top = (r.bottom + 4) + 'px';
    suggestBox.style.width = Math.max(r.width, 220) + 'px';
  }
  document.addEventListener('input', function(ev){
    var t = ev.target;
    if(!t || !t.closest || t.tagName !== 'INPUT' || t.type === 'email') return;
    if(!t.closest('#team-content, #team-modal')) return;
    showSuggest(t, filterMembers(t.value));
  });
  document.addEventListener('focusin', function(ev){
    var t = ev.target;
    if(!t || !t.closest || t.tagName !== 'INPUT' || t.type === 'email') return;
    if(!t.closest('#team-content, #team-modal')) return;
    if(t.value) showSuggest(t, filterMembers(t.value));
  });
  document.addEventListener('focusout', function(){ setTimeout(hideSuggest, 150); });

  function boot(){
    try{ ensureMembersPage(); }catch(e){ console.warn('[members] boot', e); }
    try{ void ensureUpdatesPage(); }catch(e){ console.warn('[updates] boot', e); }
  }
  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
  setTimeout(boot, 600);
  setTimeout(boot, 2000);

  console.info('[church-presenter] church-members patch loaded');
  } catch(err){
    console.error('[church-presenter] church-members failed', err);
  }
})();}


/* mjs-patch-polish-import-ccli */
if(typeof IS_AUDIENCE !== 'undefined' && !IS_AUDIENCE){(function(){
  'use strict';
  try {

  /* ============================================================
   * Contact import: skip addresses / street lines as names
   * ============================================================ */
  function isAddressLike(s){
    var t = String(s || '').trim();
    if(!t) return false;
    if(/^address\s*[:：]/i.test(t)) return true;
    if(/^addr\s*[:：]/i.test(t)) return true;
    // Starts with unit/street number patterns: 11/5, 320/62-80, 15, 3 Claremont…
    if(/^\d{1,4}\s*[/\\-]/.test(t)) return true;
    if(/^\d{1,5}\s+[A-Za-z]/.test(t) && /\b(rd|road|st|street|ave|avenue|dr|drive|cl|close|ct|court|pl|place|cres|crescent|way|ln|lane|pde|parade|hwy|highway|tce|terrace|blvd|boulevard|crt)\b/i.test(t)) return true;
    if(/\b(rd|road|st|street|ave|avenue|dr|drive|cl|close|ct|court|pl|place|cres|crescent|way|ln|lane|pde|parade|hwy|highway|tce|terrace)\b/i.test(t) && /\d/.test(t)) return true;
    // Postcode-ish AU: 2119, 2126 + suburb
    if(/\b(1|2)\d{3}\b/.test(t) && /\b(nsw|vic|qld|sa|wa|tas|act|nt)\b/i.test(t)) return true;
    if(/\b(1|2)\d{3}\b/.test(t) && t.length < 60 && !/@/.test(t) && (t.match(/[A-Za-z]/g)||[]).length > 3){
      // "2119 Beecroft" style without street type
      if(/^\d/.test(t) || /\b\d{4}\b/.test(t)) return true;
    }
    return false;
  }

  // Hook FileReader import path used by church-members patch by intercepting parse
  // We redefine a global helper if members patch exposes nothing — patch doImport via change handler priority.
  // Safer: monkey-patch after members load by wrapping parse if we can find it.
  // Instead: install our own capture-phase change handler that re-parses with stricter rules
  // when a txt is chosen, by filtering loadMembers after members patch runs.
  // Best: provide improved parse and run a cleanup pass on existing bad members + fix import.

  function improvedParseContactsTxt(text){
    var lines = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    var out = [];
    var cur = null;
    var detailsStarted = false;

    function isEmailLike(s){
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '').trim());
    }
    function isPhoneLike(s){
      var t = String(s || '').trim();
      if(!t) return false;
      var digits = t.replace(/[\s\-\.\(\)\+\/]/g, '');
      if(!/^\d{6,15}$/.test(digits)) return false;
      return ((t.match(/[A-Za-z]/g) || []).length) < 2;
    }
    function flush(){
      if(!cur) return;
      var name = String(cur.name || '').trim();
      var email = String(cur.email || '').trim();
      if(isPhoneLike(name) || isAddressLike(name)) name = '';
      if(isEmailLike(name)){
        if(!email) email = name;
        name = '';
      }
      if(/^contact directory$/i.test(name)) name = '';
      if(/^email\s*:/i.test(name)) name = '';
      if(name || email) out.push({name: name, email: email});
      cur = null;
      detailsStarted = false;
    }

    lines.forEach(function(raw){
      var line = String(raw || '').trim();
      if(!line) return;
      if(/^contact directory$/i.test(line)) return;
      if(/^no contact details/i.test(line)){ if(cur) detailsStarted = true; return; }

      var mEmail = /^(?:e-?mail)\s*[:：]\s*(.+)$/i.exec(line);
      if(mEmail){
        var em = mEmail[1].trim();
        if(!cur) cur = {name: '', email: ''};
        if(isEmailLike(em) && !cur.email) cur.email = em;
        detailsStarted = true;
        return;
      }
      if(/^(?:phone|mobile|tel|cell)\s*[:：]/i.test(line)){
        if(!cur) cur = {name: '', email: ''};
        detailsStarted = true;
        return;
      }
      if(/^(?:address|addr)\s*[:：]/i.test(line) || isAddressLike(line)){
        // Never a name — skip
        if(cur) detailsStarted = true;
        return;
      }
      if(isEmailLike(line)){
        if(!cur) cur = {name: '', email: ''};
        if(!cur.email) cur.email = line;
        detailsStarted = true;
        return;
      }
      // Org/note under person before details
      if(cur && cur.name && !detailsStarted) return;

      flush();
      // Don't start a person whose "name" is an address
      if(isAddressLike(line)){ cur = null; detailsStarted = false; return; }
      cur = {name: line, email: ''};
      detailsStarted = false;
    });
    flush();
    return out;
  }

  // Clean already-imported bad rows (address-as-name)
  function cleanupAddressNames(){
    try{
      var key = 'service-slides-church-members';
      var list = JSON.parse(localStorage.getItem(key) || '[]');
      if(!Array.isArray(list) || !list.length) return;
      var changed = false;
      list.forEach(function(m){
        if(m && isAddressLike(m.name)){
          m.name = '';
          changed = true;
        }
      });
      // Drop empty entries
      var next = list.filter(function(m){ return (m.name && m.name.trim()) || (m.email && m.email.trim()); });
      if(changed || next.length !== list.length){
        localStorage.setItem(key, JSON.stringify(next));
        console.info('[polish] cleaned address-names from members');
      }
    }catch(e){}
  }
  cleanupAddressNames();

  // Intercept txt file imports: run improved parser (capture phase, before other handlers finish)
  document.addEventListener('change', function(ev){
    var t = ev.target;
    if(!t || (t.id !== 'cm-txt-file' && t.id !== 'cm-csv-file' && t.id !== 'cm-vcf-file')) return;
    var file = t.files && t.files[0];
    if(!file) return;
    // Let the members patch also run; then after a tick, re-import with improved parse replacing list merges
    var reader = new FileReader();
    reader.onload = function(){
      try{
        var parsed = improvedParseContactsTxt(String(reader.result || ''));
        parsed = parsed.filter(function(p){ return (p.name && p.name.trim()) || (p.email && p.email.trim()); });
        if(!parsed.length) return;
        var key = 'service-slides-church-members';
        var list = [];
        try{ list = JSON.parse(localStorage.getItem(key) || '[]'); if(!Array.isArray(list)) list = []; }catch(e){ list = []; }
        var byEmail = {}, byName = {};
        list.forEach(function(m){
          if(m.email) byEmail[String(m.email).toLowerCase()] = m;
          if(m.name) byName[String(m.name).toLowerCase()] = m;
        });
        parsed.forEach(function(p){
          var keyE = (p.email || '').toLowerCase();
          var keyN = (p.name || '').toLowerCase();
          var cur = (keyE && byEmail[keyE]) || (keyN && byName[keyN]) || null;
          if(cur){
            if(p.email && !cur.email) cur.email = p.email;
            if(p.name && (!cur.name || isAddressLike(cur.name))) cur.name = p.name;
            if(isAddressLike(cur.name)) cur.name = p.name || '';
          } else {
            var m = {
              id: 'mbr-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
              name: isAddressLike(p.name) ? '' : (p.name || ''),
              email: p.email || ''
            };
            list.push(m);
            if(keyE) byEmail[keyE] = m;
            if(keyN) byName[keyN] = m;
          }
        });
        // Final sweep
        list.forEach(function(m){ if(isAddressLike(m.name)) m.name = ''; });
        list = list.filter(function(m){ return (m.name && m.name.trim()) || (m.email && m.email.trim()); });
        localStorage.setItem(key, JSON.stringify(list));
        // Refresh UI if members page is open
        setTimeout(function(){
          try{
            var tab = document.querySelector('[data-settings-page="members"]');
            if(tab && tab.classList.contains('active')) tab.click();
            else {
              var box = document.getElementById('cm-list');
              if(box){
                // trigger render by dispatching a custom event members patches may ignore —
                // force rebuild by clicking members tab if settings open
              }
            }
          }catch(e){}
        }, 100);
      }catch(err){ console.warn('[polish] import refine', err); }
    };
    // Read a copy — file input may be cleared by other handler
    try{ reader.readAsText(file); }catch(e){}
  }, true);

  /* ============================================================
   * CCLI Find: search with real song title (+ artist)
   * ============================================================ */
  document.addEventListener('click', function(ev){
    var b = ev.target && ev.target.closest && ev.target.closest('[data-ccli-lookup]');
    if(!b) return;
    ev.preventDefault();
    ev.stopImmediatePropagation();

    var title = '';
    var artist = '';
    var titleInp = document.getElementById('edit-song-title-input');
    var artistInp = document.getElementById('edit-song-artist');
    if(titleInp) title = String(titleInp.value || '').trim();
    if(artistInp) artist = String(artistInp.value || '').trim();

    // Fall back to selected song item
    try{
      if(!title && typeof S !== 'undefined' && typeof sel === 'number' && S[sel] && S[sel].t){
        title = String(S[sel].t || '').trim();
      }
      if(!artist && typeof S !== 'undefined' && S[sel] && S[sel].credits){
        var c = S[sel].credits;
        artist = String(c.artist || c.authors || c.author || '').trim();
      }
    }catch(e){}

    if(!title){
      try{ toast('Enter a song title first, then click Find.'); }catch(e){}
      if(titleInp) titleInp.focus();
      return;
    }

    var q = title;
    if(artist) q = title + ' ' + artist;
    var url = 'https://songselect.ccli.com/search/results?SearchText=' + encodeURIComponent(q);

    if(typeof openExternalUrl === 'function'){
      openExternalUrl(url).catch(function(err){
        try{ toast((err && err.message) || 'Could not open SongSelect.'); }catch(e){}
      });
    } else {
      window.open(url, '_blank');
    }
    try{ toast('Searching SongSelect for “' + title + '?…'); }catch(e){}
  }, true);

  /* ============================================================
   * Homepage polish
   * ============================================================ */
  var homeCss = document.createElement('style');
  homeCss.id = 'patch-home-polish-css';
  homeCss.textContent = [
    /* Layout */
    'body.home-mode{',
    '  background: linear-gradient(165deg, var(--bg) 0%, var(--pn) 48%, var(--bg) 100%);',
    '  min-height: 100vh;',
    '}',
    '.home{',
    '  max-width: 960px !important;',
    '  margin: 4vh auto 6vh !important;',
    '  padding: 28px 28px 40px !important;',
    '  background: var(--pn);',
    '  border: 1px solid var(--ln);',
    '  border-radius: 18px;',
    '  box-shadow: 0 18px 50px rgba(0,0,0,.08), 0 2px 8px rgba(0,0,0,.04);',
    '}',
    ':root[data-theme="dark"] .home{',
    '  box-shadow: 0 18px 50px rgba(0,0,0,.35), 0 2px 8px rgba(0,0,0,.2);',
    '}',
    '.home-head{',
    '  margin-bottom: 22px !important;',
    '  padding-bottom: 18px;',
    '  border-bottom: 1px solid var(--ln);',
    '  align-items: flex-start !important;',
    '}',
    '.home-head h1{',
    '  font-size: calc(28px * var(--ui-scale, 1));',
    '  font-weight: 700;',
    '  letter-spacing: -0.02em;',
    '  margin: 0 0 4px !important;',
    '}',
    '.home-head .note{',
    '  margin: 0;',
    '  color: var(--mu);',
    '  font-size: calc(14px * var(--ui-scale, 1));',
    '}',
    '.home-head > div:last-child{',
    '  gap: 10px !important;',
    '}',
    '#new-service.pri, #home-import-cpres{',
    '  border-radius: 10px;',
    '  padding: 10px 16px;',
    '  font-weight: 600;',
    '  min-height: 40px;',
    '  box-sizing: border-box;',
    '  display: inline-flex;',
    '  align-items: center;',
    '  justify-content: center;',
    '}',
    '#home-settings-btn,#home-cloud-btn{height:42px!important;min-height:42px!important;box-sizing:border-box;border-radius:10px;display:inline-flex;align-items:center;justify-content:center}',
    '#home-settings-btn{width:44px;min-width:44px;padding:0}',
    '#home-cloud-btn{width:44px;min-width:44px;padding:0}',
    '#home-settings-btn svg{display:block;width:18px;height:18px;flex:none}',
    /* Tools row */
    '.service-tools{',
    '  display: flex;',
    '  flex-wrap: wrap;',
    '  gap: 14px;',
    '  align-items: flex-end;',
    '  margin: 8px 0 16px;',
    '  padding: 12px 14px;',
    '  background: var(--bg);',
    '  border: 1px solid var(--ln);',
    '  border-radius: 12px;',
    '}',
    '.service-tools label{',
    '  display: flex;',
    '  flex-direction: column;',
    '  gap: 4px;',
    '  margin: 0;',
    '  font-size: calc(12px * var(--ui-scale, 1));',
    '  font-weight: 600;',
    '  color: var(--mu);',
    '  text-align: left;',
    '  text-transform: uppercase;',
    '  letter-spacing: 0.04em;',
    '}',
    '.service-tools select{',
    '  min-width: 160px;',
    '  margin: 0;',
    '  border-radius: 8px;',
    '  font-weight: 500;',
    '  text-transform: none;',
    '  letter-spacing: 0;',
    '  color: var(--ink);',
    '  font-size: calc(14px * var(--ui-scale, 1));',
    '}',
    /* Cards */
    '.service-list{',
    '  display: grid !important;',
    '  gap: 12px !important;',
    '}',
    '.service-card{',
    '  display: grid !important;',
    '  grid-template-columns: 1fr auto auto;',
    '  align-items: center;',
    '  gap: 12px 14px;',
    '  background: var(--bg) !important;',
    '  border: 1px solid var(--ln) !important;',
    '  border-radius: 14px !important;',
    '  padding: 16px 18px !important;',
    '  transition: border-color .15s, box-shadow .15s, transform .12s;',
    '  cursor: default;',
    '}',
    '.service-card:hover{',
    '  border-color: var(--ac) !important;',
    '  box-shadow: 0 8px 24px rgba(0,0,0,.06);',
    '  transform: translateY(-1px);',
    '}',
    ':root[data-theme="dark"] .service-card:hover{',
    '  box-shadow: 0 8px 24px rgba(0,0,0,.25);',
    '}',
    '.service-card .info{',
    '  min-width: 0;',
    '}',
    '.service-card .info h3,',
    '.service-card .service-name-text{',
    '  font-size: calc(17px * var(--ui-scale, 1)) !important;',
    '  font-weight: 650;',
    '  letter-spacing: -0.01em;',
    '}',
    '.service-card small{',
    '  display: block;',
    '  margin-top: 4px;',
    '  color: var(--mu);',
    '  font-size: calc(13px * var(--ui-scale, 1));',
    '}',
    '.service-card .pri{',
    '  border-radius: 9px;',
    '  padding: 9px 16px;',
    '  font-weight: 600;',
    '}',
    '.service-card [data-del-service]{',
    '  border-radius: 9px;',
    '  padding: 9px 12px;',
    '  opacity: 0.85;',
    '}',
    '.service-card [data-del-service]:hover{',
    '  opacity: 1;',
    '  background: rgba(200,40,40,.08);',
    '}',
    /* Tag chip on cards */
    '.service-card .svc-tag-chip{',
    '  display: inline-block;',
    '  margin-right: 8px;',
    '  padding: 2px 9px;',
    '  border-radius: 999px;',
    '  font-size: calc(11px * var(--ui-scale, 1));',
    '  font-weight: 700;',
    '  letter-spacing: 0.04em;',
    '  text-transform: uppercase;',
    '  background: var(--hl);',
    '  color: var(--ac);',
    '  vertical-align: middle;',
    '}',
    /* New service form */
    '.service-form{',
    '  border-radius: 14px !important;',
    '  border: 1px solid var(--ln);',
    '  padding: 18px !important;',
    '  margin-bottom: 18px;',
    '  background: var(--bg);',
    '}',
    '.service-form h2{',
    '  margin: 0 0 6px;',
    '  font-size: calc(18px * var(--ui-scale, 1));',
    '}',
    '.service-form .fields{',
    '  gap: 12px;',
    '}',
    '.service-form footer{',
    '  margin-top: 14px;',
    '  display: flex;',
    '  gap: 8px;',
    '  justify-content: flex-end;',
    '}',
    /* Empty state */
    '.home-empty{',
    '  padding: 48px 20px !important;',
    '  text-align: center;',
    '  border: 1px dashed var(--ln);',
    '  border-radius: 14px;',
    '  background: var(--bg);',
    '}',
    '.home-empty p{',
    '  margin: 0 0 18px;',
    '  color: var(--mu);',
    '  font-size: calc(15px * var(--ui-scale, 1));',
    '}',
    '@media (max-width: 640px){',
    '  .home{ margin: 2vh 12px !important; padding: 18px !important; }',
    '  .service-card{ grid-template-columns: 1fr; }',
    '  .service-card .pri, .service-card [data-del-service]{ width: 100%; }',
    '}'
  ].join('\n');
  document.head.appendChild(homeCss);

  // Enhance cards after each rHome
  function polishHomeCards(){
    try{
      document.querySelectorAll('.service-card').forEach(function(card){
        // Tag chip from title area if tag text present
        var info = card.querySelector('.info');
        if(!info) return;
        // Soften delete button label
        var del = card.querySelector('[data-del-service]');
        if(del && !del.dataset.polished){
          del.dataset.polished = '1';
          if(!del.textContent.trim() || del.textContent.trim() === 'Delete'){
            del.textContent = 'Delete';
            del.title = 'Delete service';
          }
        }
        var open = card.querySelector('[data-open-service]');
        if(open && !open.dataset.polished){
          open.dataset.polished = '1';
          open.textContent = 'Open';
        }
      });
      // Empty state class
      var list = document.getElementById('service-list');
      if(list && list.querySelector('#make-new-service')){
        var wrap = list.firstElementChild;
        if(wrap) wrap.classList.add('home-empty');
      }
    }catch(e){}
  }

  if(typeof rHome === 'function'){
    var _rHomePolish = rHome;
    rHome = function(){
      _rHomePolish.apply(this, arguments);
      try{ polishHomeCards(); }catch(e){}
    };
  }

  // One-time head copy tweak
  function polishHomeHead(){
    try{
      var note = document.querySelector('.home-head .note');
      if(note && !note.dataset.polished){
        note.dataset.polished = '1';
        note.textContent = 'Pick a service to edit, or create a new one for this week.';
      }
    }catch(e){}
  }
  if(document.readyState === 'loading'){
    document.addEventListener('DOMContentLoaded', function(){ polishHomeHead(); polishHomeCards(); });
  } else {
    polishHomeHead();
    polishHomeCards();
  }
  setTimeout(function(){ polishHomeHead(); polishHomeCards(); }, 400);

  console.info('[church-presenter] polish/import/ccli patch loaded');
  } catch(err){
    console.error('[church-presenter] polish patch failed', err);
  }
})();}


/* mjs-patch-final-clean */
if(typeof IS_AUDIENCE==='undefined'||!IS_AUDIENCE){(function(){
'use strict';
try{

/* ---- 12s startup loader ---- */
(function(){
  if(document.getElementById('cp-boot'))return;
  var MIN=12000,t0=Date.now();
  var css=document.createElement('style');
  css.textContent='#cp-boot{position:fixed;inset:0;z-index:999999;display:grid;place-items:center;background:var(--bg,#0f141a);color:var(--ink,#e8eef5);font-family:Figtree,system-ui,sans-serif}'
    +'#cp-boot.hid{opacity:0;pointer-events:none;transition:opacity .35s}'
    +'.cp-bi{width:min(340px,88vw);text-align:center}'
    +'.cp-bt{font-weight:700;font-size:20px;margin:0 0 18px}'
    +'.cp-bb{height:10px;border-radius:999px;background:rgba(128,128,128,.25);overflow:hidden}'
    +'.cp-bf{height:100%;width:0;border-radius:999px;background:var(--ac,#2e9e6b);transition:width .2s linear}'
    +'.cp-bm{margin-top:12px;font-size:13px;color:var(--mu,#8a97a5)}';
  (document.head||document.documentElement).appendChild(css);
  var w=document.createElement('div');w.id='cp-boot';
  w.innerHTML='<div class="cp-bi"><div class="cp-bt"><img src="/logo.png" alt="Church Presenter"></div><div class="cp-bb"><div class="cp-bf" id="cp-bf"></div></div><div class="cp-bm" id="cp-bm">Starting…</div></div>';
  function mount(){if(!document.body){setTimeout(mount,20);return;}if(!document.getElementById('cp-boot'))document.body.appendChild(w);}
  mount();
  var msgs=[[10,'Starting…'],[30,'Loading interface…'],[50,'Preparing services…'],[70,'Loading media…'],[90,'Almost ready…']];
  var iv=setInterval(function(){
    var el=Date.now()-t0,p=Math.min(98,(el/MIN)*98),msg='Starting…';
    for(var i=0;i<msgs.length;i++)if(p>=msgs[i][0])msg=msgs[i][1];
    var f=document.getElementById('cp-bf'),m=document.getElementById('cp-bm');
    if(f)f.style.width=p+'%';if(m)m.textContent=msg;
    if(el>=MIN){
      clearInterval(iv);if(f)f.style.width='100%';if(m)m.textContent='Ready';
      setTimeout(function(){var el2=document.getElementById('cp-boot');if(el2){el2.classList.add('hid');setTimeout(function(){try{el2.remove()}catch(e){}},400);}},250);
    }
  },100);
  setTimeout(function(){try{clearInterval(iv)}catch(e){}var el2=document.getElementById('cp-boot');if(el2){el2.classList.add('hid');setTimeout(function(){try{el2.remove()}catch(e){}},400);}},18000);
})();

/* ---- Layout CSS ---- */
var st=document.createElement('style');st.id='patch-final-clean-css';
st.textContent=[
  '#service-list .service-card,#home-templates .service-card{',
  '  display:grid!important;grid-template-columns:28px minmax(0,1fr) auto 108px 100px!important;',
  '  align-items:center;column-gap:8px;padding:10px 12px!important;}',
  '#home-templates .service-card::before{content:"";grid-column:1;width:28px;}',
  '#service-list .service-card>.info,#home-templates .service-card>.info{grid-column:2;min-width:0;}',
  '#home-templates .service-card>.tpl-default-badge,button.tpl-default-badge{',
  '  grid-column:3!important;justify-self:end;margin:0!important;',
  '  display:inline-flex!important;align-items:center;padding:4px 10px;border-radius:999px;',
  '  font-size:11px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;',
  '  cursor:pointer;border:1px solid var(--ln);background:var(--bg);color:var(--mu);opacity:.85;white-space:nowrap;}',
  '#home-templates .tpl-default-badge.is-default{background:var(--ac)!important;color:var(--ai,#fff)!important;border-color:var(--ac)!important;opacity:1;}',
  '#home-templates .service-card>button[data-tpl-default]:not(.tpl-default-badge){display:none!important;}',
  '#service-list .service-card>button.pri,#service-list .service-card>button[data-open-service],',
  '#service-list .service-card>button[data-home-cloud-open],',
  '#home-templates .service-card>button.pri,#home-templates .service-card>button[data-tpl-use]{',
  '  grid-column:4!important;width:108px!important;padding:7px 10px!important;border-radius:8px!important;font-weight:600;text-align:center;}',
  '#service-list .service-card>button[data-del-service],#service-list .service-card>button[data-home-cloud-delete],',
  '#home-templates .service-card>button[data-tpl-del]{',
  '  grid-column:5!important;width:100px!important;padding:7px 8px!important;border-radius:8px!important;text-align:center;',
  '  background:transparent!important;border:1px solid transparent!important;color:#c22!important;}',
  '#service-list .service-card .info>small{display:block;margin:3px 0 0;padding-left:8px;}',
  '#home-templates .service-card .info>small{display:block;margin:3px 0 0;padding-left:6px;}',
  '#strip .th.ms-on .sl{border:3px solid #3b82f6!important;}',
  '#strip .th.ms-on{background:rgba(59,130,246,.14)!important;}',
  '#stage .el.ms-on,.mid .el.ms-on{outline:2px solid #f4c542!important;outline-offset:2px;}',
  '#home-templates .service-card{transition:transform .32s cubic-bezier(.2,.8,.2,1);}'
].join('');
document.head.appendChild(st);

/* Ensure pill is direct child of card (left of Use) after any render */
function fixPills(){
  try{
    var box=document.getElementById('home-templates');if(!box)return;
    box.querySelectorAll('.service-card').forEach(function(card){
      card.querySelectorAll('button[data-tpl-default]:not(.tpl-default-badge)').forEach(function(b){b.style.display='none';});
      var use=card.querySelector('[data-tpl-use]');
      var pill=card.querySelector('.tpl-default-badge');
      if(!pill||!use)return;
      if(pill.parentElement!==card)card.insertBefore(pill,use);
      else if(pill.nextElementSibling!==use)card.insertBefore(pill,use);
      // remove duplicates
      card.querySelectorAll('.tpl-default-badge').forEach(function(p,i){if(i>0)p.remove();});
    });
  }catch(e){}
}
if(typeof rHome==='function'){
  var _rh=rHome;
  rHome=function(){_rh.apply(this,arguments);setTimeout(fixPills,0);};
}
document.addEventListener('click',function(ev){
  if(!(ev.target&&ev.target.closest&&ev.target.closest('.tpl-default-badge[data-tpl-default]')))return;
  setTimeout(function(){
    fixPills();
    var def=document.querySelector('#home-templates .tpl-default-badge.is-default');
    if(!def)return;
    var card=def.closest('.service-card'),list=card&&card.parentElement;
    if(!list||list.firstElementChild===card)return;
    var cards=Array.from(list.querySelectorAll(':scope>.service-card'));
    var prev=cards.map(function(c){return{c:c,r:c.getBoundingClientRect()};});
    list.insertBefore(card,cards[0]);
    prev.forEach(function(item){
      var n=item.c.getBoundingClientRect(),dx=item.r.left-n.left,dy=item.r.top-n.top;
      if(!dx&&!dy)return;
      item.c.style.transition='none';item.c.style.transform='translate('+dx+'px,'+dy+'px)';
    });
    requestAnimationFrame(function(){requestAnimationFrame(function(){
      prev.forEach(function(item){item.c.style.transition='transform .32s cubic-bezier(.2,.8,.2,1)';item.c.style.transform='';});
      setTimeout(function(){prev.forEach(function(item){item.c.style.transition='';});},350);
    });});
  },40);
});
setTimeout(fixPills,500);

/* ---- Multi-select slides (single handler, no rAll) ---- */
var MS=window.__CP_MS=window.__CP_MS||{slides:new Set(),els:new Set(),slideAnchor:null,elAnchor:null};
function sk(i,j){return i+':'+j;}
function pk(k){var p=String(k).split(':');return{i:+p[0],j:+p[1]};}
function paintSlides(){
  document.querySelectorAll('#strip .th.ms-on').forEach(function(n){n.classList.remove('ms-on');});
  MS.slides.forEach(function(k){
    var p=pk(k);
    var th=document.querySelector('#strip .th[data-i="'+p.i+'"][data-j="'+p.j+'"]');
    if(th)th.classList.add('ms-on');
  });
}
function paintEls(){
  document.querySelectorAll('#stage .el.ms-on,.mid .el.ms-on').forEach(function(n){n.classList.remove('ms-on');});
  MS.els.forEach(function(id){
    var n=document.querySelector('#stage [data-el="'+id+'"],.mid [data-el="'+id+'"]');
    if(n)n.classList.add('ms-on');
  });
}
if(typeof rStrip==='function'){var _rs=rStrip;rStrip=function(){_rs.apply(this,arguments);setTimeout(paintSlides,0);};}
if(typeof rStage==='function'){var _rst=rStage;rStage=function(){_rst.apply(this,arguments);setTimeout(paintEls,0);};}

document.addEventListener('click',function(ev){
  var th=ev.target&&ev.target.closest&&ev.target.closest('#strip .th[data-i][data-j]');
  if(!th)return;
  if(ev.target.closest&&ev.target.closest('.slide-actions,#slide-menu'))return;
  var key=sk(th.dataset.i,th.dataset.j);
  var keys=Array.from(document.querySelectorAll('#strip .th[data-i][data-j]')).map(function(t){return sk(t.dataset.i,t.dataset.j);});

  if(ev.shiftKey){
    ev.preventDefault();ev.stopImmediatePropagation();
    if(MS.slideAnchor==null)MS.slideAnchor=key;
    var a=keys.indexOf(MS.slideAnchor),b=keys.indexOf(key);
    if(a<0)a=b;if(a>b){var t=a;a=b;b=t;}
    MS.slides=new Set(keys.slice(a,b+1));
    try{var p=pk(key);A=S[p.i];sel=p.i;aj=p.j;si=p.j;elSel=null;}catch(e){}
    // Soft navigate without wiping selection
    try{if(typeof rStage==='function')rStage();}catch(e){}
    try{if(typeof rPanel==='function')rPanel();}catch(e){}
    setTimeout(paintSlides,0);
    return;
  }
  if(ev.ctrlKey||ev.metaKey){
    ev.preventDefault();ev.stopImmediatePropagation();
    if(MS.slides.has(key))MS.slides.delete(key);else MS.slides.add(key);
    MS.slideAnchor=key;
    try{var p2=pk(key);A=S[p2.i];sel=p2.i;aj=p2.j;si=p2.j;elSel=null;}catch(e){}
    try{if(typeof rStage==='function')rStage();}catch(e){}
    setTimeout(paintSlides,0);
    return;
  }
  MS.slides=new Set([key]);MS.slideAnchor=key;
  setTimeout(paintSlides,0);
},true);

function multiParts(){
  var out=[];
  MS.slides.forEach(function(k){
    var p=pk(k);
    try{
      var it=S[p.i];if(!it||(typeof isSongItem==='function'&&isSongItem(it)))return;
      var part=it.parts.find(function(x){return x.id===it.ord[p.j];});
      if(part)out.push({i:p.i,j:p.j,part:part});
    }catch(e){}
  });
  return out.sort(function(a,b){return a.i-b.i||a.j-b.j;});
}
if(typeof copySelectedSlide==='function'){
  var _c=copySelectedSlide;
  copySelectedSlide=function(){
    var items=multiParts();
    if(items.length>1){
      window.__multiSlideClip=items.map(function(x){return typeof clonePart==='function'?clonePart(x.part):JSON.parse(JSON.stringify(x.part));});
      try{toast(items.length+' slides copied');}catch(e){}
      return;
    }
    return _c.apply(this,arguments);
  };
}
if(typeof pasteSlideBelow==='function'){
  var _p=pasteSlideBelow;
  pasteSlideBelow=function(){
    var clips=window.__multiSlideClip;
    if(clips&&clips.length>1){
      try{
        var it=A;if(!it||(typeof isSongItem==='function'&&isSongItem(it))){toast('Paste into a normal section.');return;}
        if(typeof pushUndo==='function')pushUndo();
        var at=aj+1;
        clips.forEach(function(src){
          var copy=typeof clonePart==='function'?clonePart(src):JSON.parse(JSON.stringify(src));
          if(typeof uid==='function')copy.id=uid();
          it.parts.push(copy);
          if(typeof insOrd==='function')insOrd(it,at,copy.id);else it.ord.splice(at,0,copy.id);
          at++;
        });
        aj=at-1;si=aj;rAll();toast(clips.length+' slides pasted');
      }catch(e){}
      return;
    }
    return _p.apply(this,arguments);
  };
}
if(typeof deleteSelectedSlide==='function'){
  var _d=deleteSelectedSlide;
  deleteSelectedSlide=function(){
    var items=multiParts();
    if(items.length>1){
      try{
        if(typeof pushUndo==='function')pushUndo();
        items.slice().reverse().forEach(function(x){if(typeof removeSlideInItem==='function')removeSlideInItem(x.i,x.j);});
        MS.slides.clear();rAll();toast(items.length+' slides deleted');
      }catch(e){}
      return;
    }
    return _d.apply(this,arguments);
  };
}
document.addEventListener('keydown',function(ev){
  if(typeof live!=='undefined'&&live)return;
  var t=ev.target;if(t&&(t.tagName==='INPUT'||t.tagName==='TEXTAREA'||t.isContentEditable))return;
  if(multiParts().length<2)return;
  var mod=ev.ctrlKey||ev.metaKey;
  if(mod&&ev.key.toLowerCase()==='c'){ev.preventDefault();copySelectedSlide();}
  else if(mod&&ev.key.toLowerCase()==='v'){ev.preventDefault();pasteSlideBelow();}
  else if(mod&&ev.key.toLowerCase()==='x'){ev.preventDefault();copySelectedSlide();deleteSelectedSlide();}
  else if(ev.key==='Delete'||ev.key==='Backspace'){ev.preventDefault();deleteSelectedSlide();}
},true);

/* ---- Element multi-select + group move on STAGE ---- */
document.addEventListener('pointerdown',function(ev){
  if(ev.button!==0)return;
  var node=ev.target&&ev.target.closest&&ev.target.closest('#stage [data-el],.mid [data-el]');
  if(!node){
    if(ev.target.closest&&ev.target.closest('#stage .sl,.mid .sl')&&MS.els.size&&!ev.target.closest('[data-el]')){
      MS.els.clear();MS.elAnchor=null;setTimeout(paintEls,0);
    }
    return;
  }
  if(ev.target.closest&&ev.target.closest('.handle'))return;
  try{if(typeof isSongItem==='function'&&A&&isSongItem(A))return;}catch(e){}
  var id=node.getAttribute('data-el');if(!id)return;

  if(ev.shiftKey){
    try{
      var part=(typeof sls==='function'?sls(A):A.parts)[aj];
      var ids=((part&&part.elements)||[]).map(function(x){return x.id;});
      if(MS.elAnchor!=null&&ids.indexOf(MS.elAnchor)>=0&&ids.indexOf(id)>=0){
        var a=ids.indexOf(MS.elAnchor),b=ids.indexOf(id);if(a>b){var t=a;a=b;b=t;}
        MS.els=new Set(ids.slice(a,b+1));
      }else MS.els.add(id);
      MS.elAnchor=id;
    }catch(e){MS.els.add(id);MS.elAnchor=id;}
    try{elSel=id;}catch(e){}
    setTimeout(function(){try{rPanel();}catch(e){}paintEls();},0);
    return;
  }
  if(ev.ctrlKey||ev.metaKey){
    if(MS.els.has(id))MS.els.delete(id);else MS.els.add(id);
    MS.elAnchor=id;try{elSel=id;}catch(e){}
    setTimeout(function(){try{rPanel();}catch(e){}paintEls();},0);
    return;
  }
  if(MS.els.size>1&&MS.els.has(id)){try{elSel=id;}catch(e){}setTimeout(paintEls,0);return;}
  MS.els=new Set([id]);MS.elAnchor=id;setTimeout(paintEls,0);
},true);

var gDrag=null;
document.addEventListener('pointerdown',function(ev){
  if(ev.button!==0)return;
  var node=ev.target&&ev.target.closest&&ev.target.closest('#stage [data-el],.mid [data-el]');
  if(!node||(ev.target.closest&&ev.target.closest('.handle')))return;
  var id=node.getAttribute('data-el');
  if(!id||MS.els.size<2||!MS.els.has(id))return;
  try{
    var part=(typeof sls==='function'?sls(A):A.parts)[aj];
    var els=((part&&part.elements)||[]).filter(function(el){return MS.els.has(el.id);});
    if(els.length<2)return;
    var stage=document.querySelector('#stage .sl')||document.querySelector('.mid .sl');
    var rect=stage?stage.getBoundingClientRect():{width:1,height:1};
    gDrag={pid:ev.pointerId,sx:ev.clientX,sy:ev.clientY,moved:false,w:rect.width,h:rect.height,
      origins:els.map(function(el){return{el:el,x:+el.x,y:+el.y};})};
    // Stop the app's single-element drag from fighting us
    try{elDrag=null;}catch(e){}
  }catch(e){gDrag=null;}
},true);
document.addEventListener('pointermove',function(ev){
  if(!gDrag||ev.pointerId!==gDrag.pid)return;
  var dx=ev.clientX-gDrag.sx,dy=ev.clientY-gDrag.sy;
  if(!gDrag.moved&&Math.hypot(dx,dy)<4)return;
  if(!gDrag.moved){gDrag.moved=true;try{if(typeof pushUndo==='function')pushUndo();}catch(e){}try{elDrag=null;}catch(e){}}
  var pdx=(dx/gDrag.w)*100,pdy=(dy/gDrag.h)*100;
  gDrag.origins.forEach(function(o){
    o.el.x=Math.max(0,Math.min(95,o.x+pdx));
    o.el.y=Math.max(0,Math.min(95,o.y+pdy));
    // STAGE only — not strip thumbs
    var n=document.querySelector('#stage [data-el="'+o.el.id+'"],.mid [data-el="'+o.el.id+'"]');
    if(n){n.style.left=o.el.x+'%';n.style.top=o.el.y+'%';}
  });
  ev.preventDefault();
},true);
document.addEventListener('pointerup',function(ev){
  if(!gDrag||ev.pointerId!==gDrag.pid)return;
  var moved=gDrag.moved;gDrag=null;
  if(moved){
    try{if(typeof saveSoon==='function')saveSoon();}catch(e){}
    try{if(typeof updateCurrentThumb==='function')updateCurrentThumb();}catch(e){}
    setTimeout(paintEls,0);
  }
},true);

console.info('[church-presenter] final-clean loaded');
}catch(err){console.error('[church-presenter] final-clean failed',err);}
})();}


/* mjs-patch-default-drag-delete */
if(typeof IS_AUDIENCE==='undefined'||!IS_AUDIENCE){(function(){
'use strict';
try{

/* ---- CSS: stack ghost ---- */
var st=document.createElement('style');
st.id='patch-default-drag-delete-css';
st.textContent=[
  '.pk-ghost.pk-stack-slides{',
  '  position:fixed;z-index:4000;pointer-events:none;',
  '  width:120px;height:72px;',
  '}',
  '.pk-ghost.pk-stack-slides .pk-layer{',
  '  position:absolute;left:0;top:0;width:100%;height:100%;',
  '  border-radius:8px;overflow:hidden;',
  '  border:1px solid var(--ln,#333);background:var(--pn,#111);',
  '  box-shadow:0 4px 14px rgba(0,0,0,.35);',
  '}',
  '.pk-ghost.pk-stack-slides .pk-layer img,',
  '.pk-ghost.pk-stack-slides .pk-layer .sl{',
  '  width:100%!important;height:100%!important;transform:none!important;',
  '}',
  '.pk-ghost.pk-stack-slides .pk-badge{',
  '  position:absolute;right:6px;top:6px;z-index:5;',
  '  background:#3b82f6;color:#fff;border-radius:999px;',
  '  font:700 12px/1 Figtree,system-ui,sans-serif;padding:4px 8px;',
  '}'
].join('');
document.head.appendChild(st);

/* ---- Default template always first ---- */
function sortDefaultFirst(list){
  var defId='';
  try{ defId=(typeof defaultTplId==='function'?defaultTplId():'')||localStorage.getItem('service-slides-default-template')||''; }catch(e){}
  return list.slice().sort(function(a,b){
    var ad=(defId&&a.id===defId)?0:1;
    var bd=(defId&&b.id===defId)?0:1;
    if(ad!==bd) return ad-bd;
    return String(a.name||'').localeCompare(String(b.name||''));
  });
}
// Hook any loadTemplates().slice().sort(byName) results by wrapping after home render
function reorderTemplateCards(){
  try{
    var box=document.getElementById('home-templates');
    if(!box) return;
    var list=box.querySelector('.service-list')||box;
    var cards=Array.from(list.querySelectorAll(':scope > .service-card'));
    if(cards.length<2) return;
    cards.sort(function(a,b){
      var ad=a.querySelector('.tpl-default-badge.is-default')?0:1;
      var bd=b.querySelector('.tpl-default-badge.is-default')?0:1;
      return ad-bd;
    });
    cards.forEach(function(c){ list.appendChild(c); });
  }catch(e){}
}
if(typeof rHome==='function'){
  var _rHome=rHome;
  rHome=function(){
    _rHome.apply(this,arguments);
    setTimeout(reorderTemplateCards,0);
  };
}
// Also patch template list building if loadTemplates is used with sort
if(typeof loadTemplates==='function'){
  var _lt=loadTemplates;
  // sort is applied by callers: list=loadTemplates().slice().sort(byName)
  // Override byName comparison via wrapping render — reorderTemplateCards is enough
}

/* ---- Delete from Drive: always works ---- */
if(typeof cloudDeleteFile==='function'){
  cloudDeleteFile=async function(fileId){
    if(!fileId) return;
    var f=(cloudUi&&cloudUi.files||[]).find(function(x){return x.id===fileId;});
    var name=(f&&f.name?f.name:fileId).replace(/\.cpres$/i,'');
    if(!await appConfirm('Permanently delete "'+name+'" from Google Drive?\n\nThis cannot be undone, and other computers will no longer see it.',{title:'Delete from Google Drive',confirmLabel:'Delete',danger:true})) return;
    if(cloudUi&&Array.isArray(cloudUi.files)){
      cloudUi.files=cloudUi.files.filter(function(x){return x.id!==fileId;});
    }
    try{ if(typeof rHome==='function') rHome(); }catch(e){}
    try{ toast('Deleted "'+name+'" from Drive.'); }catch(e){}
    try{
      var res=await cloudInvoke('gdrive_delete_file',{fileId:fileId});
      if(!res||res.ok===false){
        var raw=(res&&res.error)||'Unknown error';
        console.warn('[cloud] delete failed',raw);
      }
    }catch(err){
      console.warn('[cloud] delete error',err);
    }
  };
}
// Never leave the button disabled just because sync is busy
document.addEventListener('click',function(ev){
  var b=ev.target&&ev.target.closest&&ev.target.closest('[data-home-cloud-delete]');
  if(!b) return;
  // Ensure our handler runs even if disabled attribute was set
  if(b.disabled){ b.disabled=false; }
  var id=b.getAttribute('data-home-cloud-delete')||b.dataset.homeCloudDelete;
  if(id&&typeof cloudDeleteFile==='function'){
    ev.preventDefault();
    ev.stopPropagation();
    cloudDeleteFile(id);
  }
},true);

// Remove disabled from cloud delete/download on each home paint
if(typeof rHome==='function'){
  var _rHome2=rHome;
  rHome=function(){
    _rHome2.apply(this,arguments);
    setTimeout(function(){
      document.querySelectorAll('[data-home-cloud-delete],[data-home-cloud-open]').forEach(function(btn){
        btn.disabled=false;
        btn.removeAttribute('disabled');
      });
      reorderTemplateCards();
    },0);
  };
}

/* ---- Multi-slide stack drag ---- */
var MS=window.__CP_MS=window.__CP_MS||{slides:new Set(),els:new Set(),slideAnchor:null,elAnchor:null};
function sk(i,j){return String(i)+':'+String(j);}
function pk(k){var p=String(k).split(':');return{i:+p[0],j:+p[1]};}

function selectedSlideKeys(){
  return Array.from(MS.slides||[]);
}

function buildStackGhost(wraps){
  var g=document.createElement('div');
  g.className='pk-ghost pk-stack-slides';
  var n=Math.min(wraps.length,4);
  for(var i=n-1;i>=0;i--){
    var layer=document.createElement('div');
    layer.className='pk-layer';
    var off=(n-1-i)*5;
    layer.style.transform='translate('+off+'px,'+off+'px) rotate('+(off*0.4)+'deg)';
    layer.style.zIndex=String(i+1);
    var thumb=wraps[i].querySelector('.th .sl, .sl');
    if(thumb){
      var clone=thumb.cloneNode(true);
      clone.style.width='100%';
      clone.style.height='100%';
      layer.appendChild(clone);
    }else{
      layer.style.background='var(--bg,#222)';
    }
    g.appendChild(layer);
  }
  var badge=document.createElement('div');
  badge.className='pk-badge';
  badge.textContent=String(wraps.length);
  g.appendChild(badge);
  return g;
}

document.addEventListener('pointerdown',function(ev){
  if(ev.button!==0||(typeof live!=='undefined'&&live)) return;
  var wrap=ev.target&&ev.target.closest&&ev.target.closest('#strip .thumb-wrap[data-i][data-j]');
  if(!wrap) return;
  var key=sk(wrap.dataset.i,wrap.dataset.j);
  if(!MS.slides||MS.slides.size<2||!MS.slides.has(key)) return;

  // All selected must be in the same service item for reorder
  var keys=selectedSlideKeys().map(pk);
  var itemIndex=keys[0].i;
  if(!keys.every(function(p){return p.i===itemIndex;})) return;

  // Collect wraps in order
  var fromJs=keys.map(function(p){return p.j;}).sort(function(a,b){return a-b;});
  var wraps=fromJs.map(function(j){
    return document.querySelector('#strip .thumb-wrap[data-i="'+itemIndex+'"][data-j="'+j+'"]');
  }).filter(Boolean);
  if(wraps.length<2) return;

  // Take over from single-slide pillDrag
  ev.stopPropagation();

  var startX=ev.clientX,startY=ev.clientY,moved=false,ghost=null,tgt=null;
  var stripEl=document.getElementById('strip');

  function clearHints(){
    document.querySelectorAll('#strip .thumb-wrap.drag-over-above,#strip .thumb-wrap.drag-over-below')
      .forEach(function(x){x.classList.remove('drag-over-above','drag-over-below');});
  }

  function onMove(e){
    if(!moved&&Math.hypot(e.clientX-startX,e.clientY-startY)<6) return;
    if(!moved){
      moved=true;
      ghost=buildStackGhost(wraps);
      document.body.appendChild(ghost);
      try{if(typeof pushUndo==='function')pushUndo();}catch(err){}
    }
    if(ghost){
      ghost.style.left='0px';
      ghost.style.top='0px';
      ghost.style.transform='translate('+(e.clientX-40)+'px,'+(e.clientY-24)+'px)';
    }
    clearHints();
    tgt=null;
    var el=document.elementFromPoint(e.clientX,e.clientY);
    var tw=el&&el.closest&&el.closest('#strip .thumb-wrap[data-i][data-j]');
    if(tw&&+tw.dataset.i===itemIndex){
      var r=tw.getBoundingClientRect();
      var before=e.clientY<r.top+r.height/2;
      tw.classList.add(before?'drag-over-above':'drag-over-below');
      tgt={j:+tw.dataset.j, before:before};
    }
    // auto-scroll strip
    if(stripEl){
      var sr=stripEl.getBoundingClientRect();
      if(e.clientY<sr.top+40) stripEl.scrollTop-=12;
      else if(e.clientY>sr.bottom-40) stripEl.scrollTop+=12;
    }
  }

  function onUp(e){
    document.removeEventListener('pointermove',onMove,true);
    document.removeEventListener('pointerup',onUp,true);
    if(ghost) ghost.remove();
    clearHints();
    if(!moved||!tgt) return;
    try{
      var it=S[itemIndex];
      if(!it) return;
      var ids=fromJs.map(function(j){return it.ord[j];});
      // Remove block
      fromJs.slice().reverse().forEach(function(j){ it.ord.splice(j,1); });
      var toJ=tgt.j+(tgt.before?0:1);
      var removedBefore=fromJs.filter(function(j){return j<toJ;}).length;
      toJ=Math.max(0,toJ-removedBefore);
      ids.forEach(function(id,n){ it.ord.splice(toJ+n,0,id); });
      aj=toJ; si=toJ; sel=itemIndex; A=it;
      MS.slides=new Set(ids.map(function(_,n){return sk(itemIndex,toJ+n);}));
      if(typeof saveSoon==='function') saveSoon();
      if(typeof rAll==='function') rAll();
      setTimeout(function(){
        try{
          document.querySelectorAll('#strip .th.ms-on').forEach(function(n){n.classList.remove('ms-on');});
          MS.slides.forEach(function(k){
            var p=pk(k);
            var th=document.querySelector('#strip .th[data-i="'+p.i+'"][data-j="'+p.j+'"]');
            if(th) th.classList.add('ms-on');
          });
        }catch(err){}
      },0);
    }catch(err){ console.warn(err); }
  }

  document.addEventListener('pointermove',onMove,true);
  document.addEventListener('pointerup',onUp,true);
},true);

console.info('[church-presenter] default-drag-delete loaded');
}catch(err){console.error('[church-presenter] default-drag-delete failed',err);}
})();}


/* mjs-patch-undo-default-anim */
if(typeof IS_AUDIENCE==='undefined'||!IS_AUDIENCE){(function(){
'use strict';
try{

/* ---- Slide rearrange always on the undo stack ---- */
if(typeof moveSlideInItem==='function'){
  var _moveSlideInItem=moveSlideInItem;
  moveSlideInItem=function(itemIndex,from,to){
    if(from===to) return;
    try{ if(typeof pushUndo==='function') pushUndo(); }catch(e){}
    return _moveSlideInItem.apply(this,arguments);
  };
}
if(typeof dropSlide==='function'){
  var _dropSlide=dropSlide;
  dropSlide=function(fromI,fromJ,t){
    // dropSlide already pushUndo for cross-item; ensure same-item path also undoes via moveSlideInItem wrapper
    return _dropSlide.apply(this,arguments);
  };
}
// Multi-slide block moves in default-drag-delete already call pushUndo; reinforce if present
document.addEventListener('pointerup',function(){
  // no-op: multi drag handles its own undo
},true);

/* ---- Default template: always first + FLIP animation when it moves ---- */
var _lastDefaultId=null;
try{ _lastDefaultId=(typeof defaultTplId==='function'?defaultTplId():'')||localStorage.getItem('service-slides-default-template')||''; }catch(e){}

function flipReorderDefaultFirst(){
  try{
    var box=document.getElementById('home-templates');
    if(!box) return;
    var list=box.querySelector('.service-list')||box;
    var cards=Array.from(list.querySelectorAll(':scope > .service-card'));
    if(cards.length<2) return;

    var defCard=cards.find(function(c){ return c.querySelector('.tpl-default-badge.is-default'); });
    if(!defCard){
      // no default — leave alphabetical/current order
      return;
    }
    if(cards[0]===defCard) return; // already on top

    // FLIP: record positions
    var prev=cards.map(function(c){ return {c:c, r:c.getBoundingClientRect()}; });
    list.insertBefore(defCard, cards[0]);

    prev.forEach(function(item){
      var now=item.c.getBoundingClientRect();
      var dx=item.r.left-now.left;
      var dy=item.r.top-now.top;
      if(!dx&&!dy) return;
      item.c.style.transition='none';
      item.c.style.transform='translate('+dx+'px,'+dy+'px)';
    });
    requestAnimationFrame(function(){
      requestAnimationFrame(function(){
        prev.forEach(function(item){
          item.c.style.transition='transform .32s cubic-bezier(.2,.8,.2,1)';
          item.c.style.transform='';
        });
        setTimeout(function(){
          prev.forEach(function(item){
            item.c.style.transition='';
            item.c.style.transform='';
          });
        }, 350);
      });
    });
  }catch(e){}
}

function quietPutDefaultFirst(){
  // Instant order without animation (first paint / home open)
  try{
    var box=document.getElementById('home-templates');
    if(!box) return;
    var list=box.querySelector('.service-list')||box;
    var cards=Array.from(list.querySelectorAll(':scope > .service-card'));
    var defCard=cards.find(function(c){ return c.querySelector('.tpl-default-badge.is-default'); });
    if(defCard&&cards[0]!==defCard) list.insertBefore(defCard, cards[0]);
  }catch(e){}
}

// On home render: quiet sort (no animation spam)
if(typeof rHome==='function'){
  var _rHome=rHome;
  rHome=function(){
    _rHome.apply(this,arguments);
    setTimeout(quietPutDefaultFirst,0);
  };
}

// (The Default pill click is handled by setDefaultTemplate, which animates the move.)

console.info('[church-presenter] undo-default-anim loaded');
}catch(err){console.error('[church-presenter] undo-default-anim failed',err);}
})();}

/* ---------- Guided first-time tour and Settings > Help ---------- */
if(!IS_AUDIENCE){(function(){
  'use strict';
  const TOUR_KEY='church-presenter-tour-v2';
  const steps=[
    {title:'Services home',text:'Find your services here. Filter, sort, import a service file, or create a new service.',target:'#home .home-head'},
    {title:'Open the service',text:'Open a service to view its sections and slides.',target:()=>{const i=services.indexOf(tourDemo);return i>=0?'#service-list [data-open-service="'+i+'"]':'#service-list'},safe:()=>!!tourDemo&&!!document.querySelector('#service-list [data-open-service="'+services.indexOf(tourDemo)+'"]')},
    {title:'Service editor',text:'Order of Sections shows the service running order. The slide list shows the slides in the selected section.',target:'#outline'},
    {title:'Add slide menu',text:'Add slide opens options for adding a blank slide, song, video, logo or imported slide.',target:()=>document.querySelector('#inline-add-popover.open')?'#inline-add-popover':'#strip-inline-add-btn'},
    {title:'Arrange the service',text:'Use section insertion points and menus to place new content in the running order.',target:'#outline'},
    {title:'Edit a slide',text:'Select a slide to preview it. Use the editing panel to change its text, appearance and background.',target:'#stage'},
    {title:'Slide actions',text:'Right-click a thumbnail to duplicate, remove, copy or add a slide. Drag thumbnails to reorder them.',target:()=>document.querySelector('#slide-menu.open')?'#slide-menu':'#strip .th'},
    {title:'Find Songs',text:'Songs is in the top bar. Open it to browse your song library or search for songs.',target:'#library-open'},
    {title:'Song library',text:'Your library and online song search are available here.',target:'#library-modal'},
    {title:'Search for songs',text:'Enter a title or keyword, then search. Matching online songs appear in the results.',target:'#song-search-form',safe:()=>true},
    {title:'Add a song',text:'Choose Add beside a result or a song in Your library to include it in the service.',target:'#ll [data-add]',safe:()=>!!document.querySelector('#ll [data-add]')},
    {title:'Edit song lyrics and structure',text:'In the song library, choose Edit to open the text editor. Put headings such as Verse 1, Chorus and Bridge on their own lines, followed by the matching lyrics.',target:'#edit-song-lyrics-input'},
    {title:'Song structure builder',text:'The left panel edits the selected section. Available sections lets you create Verse, Chorus and other sections, then drag them into the song.',target:'#struct-content .sb-edit-pane'},
    {title:'Arrange song sections',text:'The Song order panel controls the running order. Drag sections to reorder them, use the repeat button to add another instance, or remove a section from the order.',target:'#struct-content .sb-order-pane'},
    {title:'Find Media',text:'Media is in the top bar. Open it to browse reusable images, videos and presentation files.',target:'#media-open'},
    {title:'Media library',text:'Browse your media and choose an item to add it to the service.',target:'#media-content'},
    {title:'Find Service Team',text:'Open the people icon in the top bar to view and edit the service roster.',target:'#team-top-btn'},
    {title:'Service team',text:'Add the service leader, band members and sound desk contacts here.',target:'#team-modal'},
    {title:'Find Share and Export',text:'Use the share icon in the top bar to share or export the service.',target:'#export-cpres-btn'},
    {title:'Share and export',text:'Export a .cpres backup or import a .cpres file from another device.',target:'#share-modal'},
    {title:'Find Import',text:'Import is in the top bar. Use it to add PowerPoint, PDF or .cpres files to the open service.',target:'#import-slides'},
    {title:'Import a PowerPoint',text:'Choose a .pptx or PDF to import its slides into this service.',target:'#import-slides',action:'import',safe:()=>true},
    {title:'Open in Microsoft PowerPoint',text:'For an imported PowerPoint section, choose Microsoft PowerPoint in Slides Settings to present the deck in PowerPoint.',target:'input[name="ppt-view-mode"][value="powerpoint"]',safe:()=>true},
    {title:'Present',text:'Present sends the selected slide to the audience display.',target:'#live'},
    {title:'Audience display',text:'Choose the Audience Display in settings inside the Presentation tab. The audience window displays the live service',target:'#display-setting'},
    {title:'Presenter View',text:'Presenter View shows the current and next slides, running order, notes and presentation controls.',target:'#pv-outline'},
    {title:'Navigate Presenter View',text:'Use Previous and Next to move through the service, or select a thumbnail to jump to a slide.',target:'#pv-next',safe:()=>true},
    {title:'Back to Services',text:'Return to Services to open another service or create a new one. Changes to services save automatically; Google Drive sync is available when connected.',target:'#home .home-head'},
    {title:'Tour complete',text:'You have explored the service editor, songs, media, team tools, sharing, importing and Presenter View.',complete:true}
  ];
  let overlay=null,mask=null,spot=null,card=null,active=false,idx=0,focusEl=null,priorFocus=null;
  let tourActive=false,tourPresenterPreview=false,tourDemo=null,tourSnapshot=null,tourSaveOriginal=null,tourGoOriginal=null,tourLibraryOriginal=null,tourSong=null,welcomeFromHelp=false;
  const esc=s=>String(s||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  function tourState(){try{return JSON.parse(localStorage.getItem(TOUR_KEY)||'null')}catch{return null}}
  function saveTour(status,step){try{localStorage.setItem(TOUR_KEY,JSON.stringify({version:2,status,step}))}catch{}}
  function ensureLayer(){
    if(overlay)return;
    const css=document.createElement('style');css.id='cp-tour-css';css.textContent=`
      body.tour-home-over-presenter #presenter-view{z-index:0!important;pointer-events:none!important}
      body.tour-home-over-presenter #home{position:relative;z-index:210;background:var(--bg)}
      #cp-tour-layer{position:fixed;inset:0;z-index:50000;font-family:inherit;color:var(--ink,#e8eef5);pointer-events:none}
      #cp-tour-mask{position:fixed;inset:0;pointer-events:none}.cp-tour-shade{position:fixed;background:rgba(3,8,16,.72);pointer-events:auto}
      #cp-tour-block{position:fixed;background:transparent;pointer-events:auto}
      #cp-tour-spot{position:fixed;border:2px solid #f4c542;border-radius:8px;box-shadow:0 0 0 3px rgba(244,197,66,.25);pointer-events:none;z-index:2;transition:all .15s ease}
      #cp-tour-card{position:fixed;width:min(400px,calc(100vw - 28px));max-height:min(72vh,560px);overflow:auto;background:var(--pn,#202833);border:1px solid var(--ln,#526070);border-radius:14px;padding:18px;box-shadow:0 16px 48px #0009;z-index:4;pointer-events:auto}
      #cp-tour-card.cp-first-welcome{width:min(620px,calc(100vw - 28px));max-height:min(82vh,720px)}
      #cp-tour-card .cp-welcome-section{margin:18px 0}
      #cp-tour-card .cp-welcome-section h3{font-size:15px;margin:0 0 6px}
      #cp-tour-card .cp-welcome-section p{margin:0;line-height:1.55}
      #cp-tour-card h2{font-size:18px;margin:0 32px 8px 0}#cp-tour-card p{line-height:1.5;margin:0 0 12px}
      #cp-tour-progress{height:4px;background:var(--ln,#526070);border-radius:4px;margin:0 0 14px;overflow:hidden}#cp-tour-progress span{display:block;height:100%;background:#f4c542}
      .cp-tour-count{color:var(--mu,#aab4c0);font-size:12px;margin-bottom:6px}.cp-tour-actions{display:flex;align-items:center;gap:7px;flex-wrap:wrap;margin-top:14px}.cp-tour-actions .spacer{flex:1}
      #cp-tour-close{position:absolute;right:10px;top:10px;width:30px;height:30px;padding:0;font-size:18px}
      #cp-tour-card:after{content:'';position:absolute;width:12px;height:12px;background:var(--pn,#202833);border-left:1px solid var(--ln,#526070);border-top:1px solid var(--ln,#526070);transform:rotate(45deg);display:none}
      #cp-tour-card[data-arrow="up"]:after{display:block;top:-7px;left:24px}#cp-tour-card[data-arrow="down"]:after{display:block;bottom:-7px;left:24px;transform:rotate(225deg)}
      @media(max-width:600px){#cp-tour-card{max-height:58vh;padding:15px}}
    `;document.head.appendChild(css);
    overlay=document.createElement('div');overlay.id='cp-tour-layer';overlay.innerHTML='<div id="cp-tour-mask"></div><div id="cp-tour-block" hidden></div><div id="cp-tour-spot" hidden></div><section id="cp-tour-card" role="dialog" aria-modal="true" aria-live="polite"></section>';
    document.body.appendChild(overlay);mask=overlay.querySelector('#cp-tour-mask');spot=overlay.querySelector('#cp-tour-spot');card=overlay.querySelector('#cp-tour-card');
    mask.addEventListener('click',ev=>{ev.preventDefault();ev.stopPropagation()},true);card.addEventListener('click',ev=>ev.stopPropagation());
  }
  function visible(el){if(!el||!el.isConnected)return false;const r=el.getBoundingClientRect(),s=getComputedStyle(el);return r.width>0&&r.height>0&&s.display!=='none'&&s.visibility!=='hidden'}
  function getTarget(step){try{const raw=typeof step.target==='function'?step.target():step.target,el=raw&&document.querySelector(raw);return visible(el)?el:null}catch{return null}}
  function setShade(r){
    mask.innerHTML='';const W=innerWidth,H=innerHeight;
    const panes=r?[{left:0,top:0,width:W,height:r.top},{left:0,top:r.bottom,width:W,height:H-r.bottom},{left:0,top:r.top,width:r.left,height:r.height},{left:r.right,top:r.top,width:W-r.right,height:r.height}]:[{left:0,top:0,width:W,height:H}];
    panes.filter(p=>p.width>0&&p.height>0).forEach(p=>{const d=document.createElement('div');d.className='cp-tour-shade';d.style.left=p.left+'px';d.style.top=p.top+'px';d.style.width=p.width+'px';d.style.height=p.height+'px';mask.appendChild(d)});
  }
  function positionTour(){
    if(!active||!overlay)return;const step=steps[idx],el=getTarget(step),block=overlay.querySelector('#cp-tour-block');
    if(el){focusEl=el;const r=el.getBoundingClientRect(),pad=5,rr={left:Math.max(0,r.left-pad),top:Math.max(0,r.top-pad),right:Math.min(innerWidth,r.right+pad),bottom:Math.min(innerHeight,r.bottom+pad)};rr.width=rr.right-rr.left;rr.height=rr.bottom-rr.top;setShade(rr);spot.hidden=false;Object.assign(spot.style,{left:rr.left+'px',top:rr.top+'px',width:rr.width+'px',height:rr.height+'px'});
      const safe=typeof step.safe==='function'&&step.safe();block.hidden=!!safe;if(!safe)Object.assign(block.style,{left:r.left+'px',top:r.top+'px',width:r.width+'px',height:r.height+'px'});
      const vw=innerWidth,vh=innerHeight,cw=Math.min(400,vw-28),ch=card.offsetHeight||250;const rightFits=r.right+cw+16<vw;let left=rightFits?r.right+14:r.left-cw-14;left=Math.max(14,Math.min(vw-cw-14,left));let top=r.bottom+ch+18<vh?r.bottom+14:(r.top-ch-14>8?r.top-ch-14:Math.max(12,(vh-ch)/2));card.dataset.arrow=top>r.bottom?'up':(top+ch<r.top?'down':'none');card.style.left=left+'px';card.style.top=top+'px';
    }else{focusEl=null;setShade(null);spot.hidden=true;block.hidden=true;card.dataset.arrow='none';const cw=Math.min(400,innerWidth-28),ch=card.offsetHeight||250;card.style.left=Math.max(14,(innerWidth-cw)/2)+'px';card.style.top=Math.max(14,(innerHeight-ch)/2)+'px'}
  }
  function targetForHomeDemo(){const n=services.indexOf(tourDemo);return n>=0?'#service-list [data-open-service="'+n+'"]':'#service-list'}
  function renderTourWelcome(){
    ensureLayer();active=true;card.classList.remove('cp-first-welcome');card.dataset.arrow='none';setShade(null);spot.hidden=true;overlay.querySelector('#cp-tour-block').hidden=true;
    card.innerHTML='<button type="button" id="cp-tour-close" data-tour="close" aria-label="Close tour" title="Close tour">×</button><div class="cp-tour-count">Welcome</div><h2>Welcome to Church Presenter</h2><p>This guided tour will take you from the Services home page through a blank practice service, the song and media libraries, importing slides and Presenter View. It will not send content to an audience display.</p><div class="cp-tour-actions"><button type="button" class="pri" data-tour="start">Start tour</button><span class="spacer"></span><button type="button" data-tour="skip">Skip tour</button></div>';
    card.style.left=Math.max(14,(innerWidth-Math.min(400,innerWidth-28))/2)+'px';card.style.top=Math.max(14,(innerHeight-(card.offsetHeight||250))/2)+'px';
  }
  function renderWelcome(){
    ensureLayer();active=true;card.classList.add('cp-first-welcome');card.dataset.arrow='none';setShade(null);spot.hidden=true;overlay.querySelector('#cp-tour-block').hidden=true;
    card.innerHTML='<button type="button" id="cp-tour-close" data-tour="close" aria-label="Close welcome page" title="Close welcome page">×</button><div class="cp-tour-count">A quick introduction</div><h2>Welcome to Church Presenter</h2><p>A familiar way to plan and present your church service.</p><section class="cp-welcome-section"><h3>What can it do?</h3><p>Bring songs, Bible readings, announcements and the rest of your service together in one place. Church Presenter is designed around the way a church service flows.</p></section><section class="cp-welcome-section"><h3>Why not PowerPoint or Google Slides? What about ProPresenter?</h3><p>PowerPoint and Google Slides are great presentation tools, but they were not built for the specific needs of a church service. We wanted the familiar feel of a slide editor, with capabilities that are genuinely useful in a church environment.</p><p style="margin-top:10px">ProPresenter can do all of that and much more. For some churches, its broad production capabilities are exactly what they need. For others, it can feel like more software and complexity than their services require.</p><p style="margin-top:10px">Church Presenter aims for the middle ground: church focused tools in a familiar, approachable experience.</p></section><div class="cp-tour-actions"><button type="button" class="pri" data-tour="continue">Tour</button><span class="spacer"></span><button type="button" data-tour="skip">Explore on my own</button></div>';
    const cw=Math.min(620,innerWidth-28);card.style.left=Math.max(14,(innerWidth-cw)/2)+'px';card.style.top=Math.max(14,(innerHeight-(card.offsetHeight||420))/2)+'px';
  }
  function makeDemo(){
    if(tourDemo)return;
    tourSaveOriginal=saveState;tourGoOriginal=go;tourLibraryOriginal=LIB.slice();
    tourSong={t:'Tour Example Song',...parse('Verse 1\nExample verse lyrics\n\nChorus\nExample chorus lyrics\n\nBridge\nExample bridge lyrics')};LIB.push(tourSong);
    tourSnapshot={services,activeService,S,sel,si,A,aj,tab,serviceTagFilter,serviceSort,bodyClass:document.body.className};
    tourActive=true;
    saveState=function(){if(tourActive)return;return tourSaveOriginal.apply(this,arguments)};
    go=function(d){tourGoOriginal(d);if(tourPresenterPreview)renderPresenterView(true)};
    const slide=mkS('Welcome',['Type your text here'],0,6.5);
    tourDemo={name:'Tour Practice Service',date:new Date().toISOString().slice(0,10),tag:'',slides:[slide],team:defaultTeam(),lastEdited:Date.now(),_tourDemo:true};
    services=services.concat(tourDemo);activeService=services.length-1;S=cl(tourDemo.slides);sel=0;si=0;A=S[0];aj=0;serviceTagFilter='';
    document.body.classList.add('home-mode');document.body.classList.remove('presenter-mode','presentation-mode');rHome();
  }
  function openDemo(){if(!tourDemo)return;activeService=services.indexOf(tourDemo);S=cl(tourDemo.slides);sel=0;si=0;A=S[0]||null;aj=0;tab='Structure';document.body.classList.remove('home-mode');rAll()}
  function ensureTourSongInService(){
    if(!tourDemo||!tourSong)return;
    tourSong=LIB.find(x=>x.t==='Tour Example Song')||tourSong;
    let item=S.find(x=>x.k==='Song'&&x.t===tourSong.t);
    if(!item){item=fromLib(tourSong,0);S.push(item)}
    tourDemo.slides=cl(S);activeService=services.indexOf(tourDemo);sel=S.indexOf(item);si=0;A=item;aj=0;tab='Structure';rAll();
  }
  function injectDemoPowerpoint(){
    if(!tourActive||!tourDemo||S.some(it=>it._tourPowerpoint))return;
    const item=mk('PowerPoint example','PowerPoint',[{id:uid(),n:'Slide 1',tx:'Example PowerPoint slide'},{id:uid(),n:'Slide 2',tx:'A second example slide'}].map(p=>({...p,elements:[{id:uid(),kind:'text',tx:p.tx,x:8,y:18,w:84,h:50,rot:0,z:1,f:{...S[0].f}}]})),[],0,5.2);
    item.ord=item.parts.map(p=>p.id);item._tourPowerpoint=true;item.viewMode='church';S.push(item);tourDemo.slides=cl(S);activeService=services.indexOf(tourDemo);sel=S.length-1;si=0;A=S[sel];aj=0;tab='Slides';rAll();
  }
  function closeTourWindows(){
    try{if(typeof closeLibraryModal==='function')closeLibraryModal()}catch{}try{if(typeof closeMediaLibrary==='function')closeMediaLibrary()}catch{}try{if(typeof closeEditSongModal==='function')closeEditSongModal()}catch{}try{if(typeof closeStruct==='function'&&!$('#struct-modal')?.hidden)closeStruct()}catch{}try{if(typeof closeTeamModal==='function')closeTeamModal()}catch{}try{if(typeof closeShareModal==='function')closeShareModal()}catch{}
    ['#settings','#shortcuts-modal','#import-modal','#struct-modal','#background-picker-modal','#team-modal','#share-modal'].forEach(s=>{const el=document.querySelector(s);if(el){if(el.id==='settings')closeSettingsAndRefresh();else el.hidden=true}});
    document.querySelectorAll('.outline-menu.open,.ctx-menu.open').forEach(m=>{m.classList.remove('open');m.hidden=true});
    try{if(typeof hideInlineAddPopover==='function')hideInlineAddPopover()}catch{}try{if(typeof hideAddMenu==='function')hideAddMenu()}catch{}
  }
  function leavePresenterPreview(){if(!tourPresenterPreview)return;tourPresenterPreview=false;document.body.classList.remove('presentation-mode','presenter-mode','tour-home-over-presenter');const end=document.getElementById('pv-end');if(end)end.textContent='End presentation';const pv=$('#presenter-view');if(pv){pv.hidden=true;pv.setAttribute('hidden','');['display','z-index','visibility','opacity','pointer-events','position','inset'].forEach(p=>pv.style.removeProperty(p))}}
  async function removeTourPracticeFromDrive(){
    try{
      if(!cloudReady())return;
      await fetchCloudFiles();
      const matches=(cloudUi.files||[]).filter(f=>/^Tour Practice Service(?:\s|\.cpres|$)/i.test(String(f.name||''))&&/\.cpres$/i.test(String(f.name||'')));
      for(const f of matches){
        try{await cloudInvoke('gdrive_delete_file',{fileId:f.id})}catch(err){console.warn('[tour] Drive cleanup failed',err)}
      }
      if(matches.length){cloudUi.files=cloudUi.files.filter(f=>!matches.some(m=>m.id===f.id));}
    }catch(err){console.warn('[tour] Drive cleanup failed',err)}
  }
  function cleanupSession(toHome){
    if(!tourActive)return;closeTourWindows();leavePresenterPreview();tourActive=false;
    if(tourSaveOriginal)saveState=tourSaveOriginal;if(tourGoOriginal)go=tourGoOriginal;
    if(tourSnapshot){services=tourSnapshot.services;activeService=tourSnapshot.activeService;S=tourSnapshot.S;sel=tourSnapshot.sel;si=tourSnapshot.si;A=tourSnapshot.A;aj=tourSnapshot.aj;tab=tourSnapshot.tab;serviceTagFilter=tourSnapshot.serviceTagFilter;serviceSort=tourSnapshot.serviceSort;}
    if(tourLibraryOriginal)LIB.splice(0,LIB.length,...tourLibraryOriginal);
    const oldCount=services.length;services=services.filter(s=>String(s.name||'').trim().toLowerCase()!=='tour practice service');
    if(services.length!==oldCount){activeService=Math.max(0,Math.min(activeService,services.length-1));try{saveState()}catch(err){console.warn('[tour] local cleanup failed',err)}}
    void removeTourPracticeFromDrive();
    tourDemo=null;tourSnapshot=null;tourSaveOriginal=null;tourGoOriginal=null;tourLibraryOriginal=null;
    if(toHome){document.body.classList.remove('presenter-mode','presentation-mode');document.body.classList.add('home-mode');try{rHome()}catch{}}
  }
  function enterStep(n){
    idx=Math.max(0,Math.min(n,steps.length-1));const st=steps[idx];closeTourWindows();
    if(idx===0||idx===1){document.body.classList.add('home-mode');rHome()}
    if(idx===3){const b=$('#strip-inline-add-btn');if(b)b.click()}
    if(idx===4){const row=document.querySelector('#outline [data-oi]');if(row)row.dispatchEvent(new MouseEvent('contextmenu',{bubbles:true,cancelable:true,clientX:row.getBoundingClientRect().left+40,clientY:row.getBoundingClientRect().top+30,view:window}))}
    if(idx===6){const th=document.querySelector('#strip .th');if(th)th.dispatchEvent(new MouseEvent('contextmenu',{bubbles:true,cancelable:true,clientX:th.getBoundingClientRect().left+30,clientY:th.getBoundingClientRect().top+20,view:window}))}
    if(idx===8||idx===9||idx===10){openLibraryModal();if(idx===9){const q=$('#q');if(q)q.value='';}}
    if(idx===11){openLibraryModal();openEditSongModal(LIB.indexOf(tourSong))}
    if(idx===12||idx===13){ensureTourSongInService();openStruct()}
    if(idx===15)openMediaLibrary();
    if(idx===17){closeTeamModal();openTeamModal()}
    if(idx===19)openShareModal();
    if(idx===22)injectDemoPowerpoint();
    if(idx===24){leavePresenterPreview();document.body.classList.remove('home-mode');$('#settings')?.classList.add('open');const tabBtn=document.querySelector('[data-settings-page="display"]');if(tabBtn)tabBtn.click()}
    if(idx===25||idx===27){if(!tourPresenterPreview){document.body.classList.remove('home-mode');showPresenterShell();tourPresenterPreview=true;live=false;const end=document.getElementById('pv-end');if(end)end.textContent='Close preview';renderPresenterView(true)}}
    if(idx===27){document.body.classList.remove('presentation-mode');document.body.classList.add('presenter-mode','home-mode','tour-home-over-presenter');rHome()}
    if(st.complete){cleanupSession(true);saveTour('complete',idx)}else saveTour('running',idx);
    renderTour();
  }
  function renderTour(){
    if(!active)return;ensureLayer();const st=steps[idx],n=idx+1,complete=!!st.complete;
    card.innerHTML='<button type="button" id="cp-tour-close" data-tour="close" aria-label="Close tour" title="Close tour">×</button><div class="cp-tour-count">Step '+n+' of '+steps.length+'</div><div id="cp-tour-progress"><span style="width:'+Math.round(n/steps.length*100)+'%"></span></div><h2>'+esc(st.title)+'</h2><p>'+esc(st.text)+'</p><div class="cp-tour-actions">'+(complete?'<button type="button" data-tour="restart">Restart tour</button><span class="spacer"></span><button type="button" class="pri" data-tour="explore">Explore the app</button>':'<button type="button" data-tour="back" '+(idx===0?'disabled':'')+'>Back</button><button type="button" class="pri" data-tour="next">Next</button><span class="spacer"></span><button type="button" data-tour="skip">Skip tour</button>')+'</div>'+(st.action==="import"?'<p class="cp-tour-actions"><button type="button" data-tour-action="import">Choose a PowerPoint or PDF (optional)</button></p>':'');
    active=true;positionTour();
  }
  function positionTour(){
    if(!active||!overlay)return;const st=steps[idx],el=getTarget(st),block=overlay.querySelector('#cp-tour-block');
    if(el){focusEl=el;const r=el.getBoundingClientRect(),p=5,q={left:Math.max(0,r.left-p),top:Math.max(0,r.top-p),right:Math.min(innerWidth,r.right+p),bottom:Math.min(innerHeight,r.bottom+p)};q.width=q.right-q.left;q.height=q.bottom-q.top;setShade(q);spot.hidden=false;Object.assign(spot.style,{left:q.left+'px',top:q.top+'px',width:q.width+'px',height:q.height+'px'});const safe=typeof st.safe==='function'&&st.safe();block.hidden=!!safe;if(!safe)Object.assign(block.style,{left:r.left+'px',top:r.top+'px',width:r.width+'px',height:r.height+'px'});const cw=Math.min(400,innerWidth-28),ch=card.offsetHeight||260;let left=r.right+cw+18<innerWidth?r.right+14:r.left-cw-14;left=Math.max(14,Math.min(innerWidth-cw-14,left));let top=r.bottom+ch+18<innerHeight?r.bottom+14:(r.top-ch-14>8?r.top-ch-14:Math.max(12,(innerHeight-ch)/2));card.dataset.arrow=top>r.bottom?'up':(top+ch<r.top?'down':'none');card.style.left=left+'px';card.style.top=top+'px'}
    else{focusEl=null;setShade(null);spot.hidden=true;block.hidden=true;card.dataset.arrow='none';const cw=Math.min(400,innerWidth-28),ch=card.offsetHeight||260;card.style.left=Math.max(14,(innerWidth-cw)/2)+'px';card.style.top=Math.max(14,(innerHeight-ch)/2)+'px'}
  }
  function beginTour(step){
    welcomeFromHelp=false;priorFocus=document.activeElement;if(!tourActive){try{saveState()}catch{};makeDemo()}
    active=true;ensureLayer();
    const resume=step!=null&&step>0;if(resume&&step>1)openDemo();if(resume&&step>=22)injectDemoPowerpoint();enterStep(step||0);window.addEventListener('resize',positionTour);window.addEventListener('scroll',positionTour,true);document.addEventListener('keydown',tourKey,true);
  }
  function endTour(status){saveTour(status,idx);cleanupSession(true);removeLayer()}
  function removeLayer(){if(!overlay)return;window.removeEventListener('resize',positionTour);window.removeEventListener('scroll',positionTour,true);document.removeEventListener('keydown',tourKey,true);overlay.remove();overlay=null;mask=spot=card=null;active=false;focusEl=null;if(priorFocus&&priorFocus.isConnected)try{priorFocus.focus()}catch{}priorFocus=null}
  function tourKey(ev){if(!active)return;if(ev.key==='Escape'){ev.preventDefault();ev.stopImmediatePropagation();if(welcomeFromHelp){welcomeFromHelp=false;removeLayer()}else endTour('dismissed')}}
  function setTourStep(n){if(n<0)return;const current=idx;if(current===1&&n>1)openDemo();if((current===25||current===26||current===27)&&n<25)leavePresenterPreview();enterStep(n)}
  document.addEventListener('click',ev=>{
    const b=ev.target.closest&&ev.target.closest('[data-tour]');if(b){ev.preventDefault();ev.stopPropagation();ev.stopImmediatePropagation();switch(b.dataset.tour){case'continue':renderTourWelcome();break;case'start':beginTour(0);break;case'next':if(idx===steps.length-2)enterStep(steps.length-1);else setTourStep(idx+1);break;case'back':setTourStep(idx-1);break;case'skip':case'close':if(welcomeFromHelp){welcomeFromHelp=false;removeLayer()}else endTour('dismissed');break;case'explore':endTour('complete');break;case'restart':cleanupSession(true);removeLayer();beginTour(0);break}return}
    const act=ev.target.closest&&ev.target.closest('[data-tour-action="import"]');if(act&&active){ev.preventDefault();ev.stopPropagation();ev.stopImmediatePropagation();$('#import-slides')?.click();return}
  },true);
  document.addEventListener('click',ev=>{if(!active||!focusEl)return;const st=steps[idx];if(!(typeof st.safe==='function'&&st.safe()))return;if(!focusEl.contains(ev.target))return;setTimeout(positionTour,220)},true);
  document.addEventListener('click',ev=>{
    const b=ev.target.closest&&ev.target.closest('#tour-start-from-help,#tour-why-church-presenter,#tour-help-guidance,#tour-help-shortcuts');if(!b)return;
    ev.preventDefault();ev.stopPropagation();ev.stopImmediatePropagation();
    if(b.id==='tour-help-guidance'){const box=$('#tour-guidance');box.hidden=!box.hidden;return}
    if(b.id==='tour-help-shortcuts'){const m=$('#shortcuts-modal');if(m)m.hidden=false;return}
    if(b.id==='tour-why-church-presenter'){welcomeFromHelp=true;ensureLayer();renderWelcome();document.addEventListener('keydown',tourKey,true);return}
    beginTour(0);
  },true);
  const helpTab=document.querySelector('[data-settings-page="help"]');if(helpTab&&helpTab.parentElement){helpTab.innerHTML='<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="10"/><path d="M9.1 9a3 3 0 0 1 5.8 1c0 2-3 3-3 3M12 17h.01"/></svg> Help';helpTab.parentElement.appendChild(helpTab);}
  const membersTab=document.querySelector('[data-settings-page="members"]'),updatesTab=document.querySelector('[data-settings-page="updates"]');
  if(membersTab&&updatesTab&&membersTab.parentElement===updatesTab.parentElement)membersTab.parentElement.insertBefore(updatesTab,membersTab.nextSibling);
  function showWelcome(){if(tourState())return;ensureLayer();renderWelcome();document.addEventListener('keydown',tourKey,true)}
  const saved=tourState();if(saved&&saved.status==='running')setTimeout(()=>beginTour(Number.isInteger(saved.step)?saved.step:0),700);else if(!saved)setTimeout(showWelcome,900);
})();}
// Keep the desktop window compact on Services, and maximize the editor window
// without entering native fullscreen mode.
(()=>{
  let resizing=false, requestedHome=null;
  async function resizeServicesWindow(home){
    requestedHome=home;
    if(resizing)return;
    resizing=true;
    try{
      const [{getCurrentWindow},{LogicalSize}]=await Promise.all([
        import('@tauri-apps/api/window'),import('@tauri-apps/api/dpi')
      ]);
      const win=getCurrentWindow();
      while(requestedHome!==null){
        const next=requestedHome;requestedHome=null;
        if(next){
          try{await win.unmaximize()}catch{}
          await win.setMinSize(new LogicalSize(1100,760));
          await win.setMaxSize(new LogicalSize(1100,760));
          await win.setSize(new LogicalSize(1100,760));
        }else{
          await win.setMaxSize(null);
          await win.setMinSize(new LogicalSize(1120,700));
          await win.maximize();
        }
      }
    }catch(err){console.warn('[window] Could not resize Services window',err)}
    finally{resizing=false}
  }
  const sync=()=>resizeServicesWindow(document.body.classList.contains('home-mode'));
  new MutationObserver(sync).observe(document.body,{attributes:true,attributeFilter:['class']});
  sync();
})();
