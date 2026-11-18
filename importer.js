/* Lector inteligente de cronogramas — Mi Tablero
   Entiende Excel (.xlsx/.xls/.ods), CSV, Word (.docx), PDF y texto.
   Devuelve tareas con: title, subject, date (AAAA-MM-DD), time (HH:MM), status ('done'|'prog'|''), desc. */
(function(){
'use strict';
const pad = n => String(n).padStart(2,'0');
const LIBS = { xlsx:'lib/xlsx-0.20.3.full.min.js', mammoth:'lib/mammoth.browser.min.js', pdf:'lib/pdfjs-4.10.38/cargar.js' };
const MAX_ARCHIVO = 20 * 1048576; // 20 MB por archivo
const loaded = {};
function loadLib(k){
  return new Promise((res, rej)=>{
    if(loaded[k]) return res();
    const s = document.createElement('script'); s.src = LIBS[k];
    s.onload = ()=>{ loaded[k]=true; res(); };
    s.onerror = ()=>rej(new Error('No se pudo cargar el lector de este tipo de archivo. Abre la app con internet una vez para descargarlo.'));
    document.head.appendChild(s);
  });
}

/* ---------- Texto normalizado (misma longitud que el original) ---------- */
function normSameLen(s){ let o=''; for(const ch of String(s)){ o += ch.length===1 ? (ch.normalize('NFD')[0]||ch).toLowerCase() : ch; } return o; }
const norm = s => String(s??'').normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase().trim();

/* ---------- Fechas ---------- */
const MES = {ene:1,enero:1,jan:1,january:1,feb:2,febrero:2,february:2,mar:3,marzo:3,march:3,abr:4,abril:4,apr:4,april:4,may:5,mayo:5,jun:6,junio:6,june:6,jul:7,julio:7,july:7,ago:8,agosto:8,aug:8,august:8,sep:9,sept:9,set:9,septiembre:9,setiembre:9,september:9,oct:10,octubre:10,october:10,nov:11,noviembre:11,november:11,dic:12,diciembre:12,dec:12,december:12};
const MES_RE = '(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre|january|february|march|april|june|july|august|september|october|november|december|ene|feb|mar|abr|may|jun|jul|ago|sept|sep|set|oct|nov|dic|jan|apr|aug|dec)\\.?';
const DIA_RE = '(?:(?:lunes|martes|miercoles|jueves|viernes|sabado|domingo|lun|mar|mie|jue|vie|sab|dom|monday|tuesday|wednesday|thursday|friday|saturday|sunday|mon|tue|wed|thu|fri|sat|sun)\\.?,?\\s+)?';
const PATTERNS = [
  // 2026-11-20 / 2026/11/20
  {re:/\b(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})\b/, f:m=>({y:+m[1], mo:+m[2], d:+m[3]})},
  // 17 al 21 de noviembre (de 2026)  → se toma el último día
  {re:new RegExp('\\b(\\d{1,2})\\s*(?:al|a|-|–|y)\\s*(\\d{1,2})\\s*(?:de\\s+)?'+MES_RE+'(?:\\s*(?:de|del|,)?\\s*(20\\d{2}))?','i'), f:m=>({y:m[4]?+m[4]:null, mo:MES[m[3]], d:+m[2], start:+m[1]})},
  // (lunes) 20 de noviembre (de 2026) / 20 nov 2026 / 20-nov-26
  {re:new RegExp('\\b'+DIA_RE+'(\\d{1,2})\\s*(?:de\\s+|-|\\/|\\s)?\\s*'+MES_RE+'(?:\\s*(?:de|del|,|-|\\/)?\\s*(20\\d{2}|\\d{2})(?!\\d|:))?','i'), f:m=>({y:m[3]?(m[3].length===2?2000+(+m[3]):+m[3]):null, mo:MES[m[2]], d:+m[1]})},
  // November 20, 2026 / nov 20
  {re:new RegExp('\\b'+MES_RE+'\\s+(\\d{1,2})(?:st|nd|rd|th)?(?:,?\\s*(20\\d{2}))?\\b','i'), f:m=>({y:m[3]?+m[3]:null, mo:MES[m[1]], d:+m[2]})},
  // 20/11/2026, 20/11, 20-11-2026, 20.11.2026 (con guion o punto se exige el año para no confundir "capítulos 3-4")
  {re:/(?<![\d:\/.\-])(\d{1,2})(?:\/(\d{1,2})(?:\/(20\d{2}|\d{2}))?|[.\-](\d{1,2})[.\-](20\d{2}|\d{2}))(?![\d:\/])/, f:m=>{
    if(m[4]){ m = [m[0], m[1], m[4], m[5]]; }
    let a=+m[1], b=+m[2], d, mo;
    if(a>12 && b<=12){ d=a; mo=b; } else if(b>12 && a<=12){ d=b; mo=a; } else { d=a; mo=b; } // formato latino día/mes por defecto
    return {y:m[3]?(m[3].length===2?2000+(+m[3]):+m[3]):null, mo, d};
  }},
];
function validDate(y,mo,d){ if(!(mo>=1&&mo<=12&&d>=1&&d<=31)) return false; const t=new Date(y,mo-1,d); return t.getMonth()===mo-1; }
function inferYear(mo, d, ctx){
  if(ctx && ctx.year) return ctx.year;
  const t = new Date(); let y = t.getFullYear();
  if(t - new Date(y, mo-1, d) > 75*86400000) y++;
  return y;
}
function toISO(y,mo,d){ return y+'-'+pad(mo)+'-'+pad(d); }
// Busca una fecha en el texto. Devuelve {date, index, length, start?}
function findDate(text, ctx){
  const n = normSameLen(text);
  for(const p of PATTERNS){
    const m = n.match(p.re);
    if(!m) continue;
    const r = p.f(m); if(!r.mo) continue;
    const y = r.y || inferYear(r.mo, r.d, ctx);
    if(!validDate(y, r.mo, r.d)) continue;
    const out = {date: toISO(y,r.mo,r.d), index:m.index, length:m[0].length};
    if(r.start && r.start < r.d) out.start = toISO(y, r.mo, r.start);
    return out;
  }
  return null;
}
function dateFromCell(v, ctx){
  if(v instanceof Date && !isNaN(v)){
    // SheetJS crea fechas en hora local con desfase; usamos la parte de fecha
    const d = new Date(v.getTime() + 12*3600000);
    return {date: toISO(d.getFullYear(), d.getMonth()+1, d.getDate()), time: (v.getHours()||v.getMinutes()) ? pad(v.getHours())+':'+pad(v.getMinutes()) : ''};
  }
  if(typeof v === 'number' && v > 30000 && v < 80000){ // número de serie de Excel
    const d = new Date(Math.round((v - 25569) * 86400000)); const frac = v % 1;
    const time = frac ? pad(Math.floor(frac*24))+':'+pad(Math.round((frac*24%1)*60)%60) : '';
    return {date: toISO(d.getUTCFullYear(), d.getUTCMonth()+1, d.getUTCDate()), time};
  }
  const s = String(v??'').trim(); if(!s) return null;
  const f = findDate(s, ctx); if(!f) return null;
  const t = findTime(s.slice(0,f.index)+' '+s.slice(f.index+f.length));
  return {date:f.date, time: t ? t.time : ''};
}
/* ---------- Hora ---------- */
function findTime(text){
  const n = normSameLen(text);
  let m = n.match(/(?<![\d\/])(\d{1,2})[:h](\d{2})\s*(a\.?\s?m\.?|p\.?\s?m\.?|hrs?|horas)?(?![\d\/])/i)
       || n.match(/(?<![\d\/])(\d{1,2})()\s*(a\.?\s?m\.?|p\.?\s?m\.?)(?![a-z])/i);
  if(!m) return null;
  let h = +m[1], mi = m[2] ? +m[2] : 0; const ap = (m[3]||'').replace(/[\.\s]/g,'');
  if(ap==='pm' && h<12) h+=12; if(ap==='am' && h===12) h=0;
  if(h>23 || mi>59) return null;
  return {time: pad(h)+':'+pad(mi), index:m.index, length:m[0].length};
}
/* ---------- Estado ---------- */
const DONE_RE = /(\b(completad[oa]s?|hech[oa]s?|entregad[oa]s?|list[oa]s?|finalizad[oa]s?|terminad[oa]s?|realizad[oa]s?|aprobad[oa]s?|done|completed?|finished)\b|✔|✓|☑|✅)/i;
const PROG_RE = /\b(en progreso|en curso|en proceso|haciendo|iniciad[oa]|empezad[oa]|in progress|doing|realizando)\b/i;
const TODO_RE = /\b(pendiente|por hacer|sin empezar|no iniciad[oa]|to ?do)\b/i;
function statusOf(s){ const t = norm(s); if(!t) return ''; if(PROG_RE.test(t)) return 'prog'; if(TODO_RE.test(t)) return ''; if(DONE_RE.test(s) || DONE_RE.test(t)) return 'done'; return ''; }

/* ---------- Limpieza de títulos ---------- */
function clean(s){
  return String(s??'')
    .replace(/^[\s\-–—•*·▪◦●○■□➢►▶→>|:;,]+/, '')
    .replace(/^\(?\d{1,2}[\.\)]\s+/, '')
    .replace(/[\s\-–—|:;,]+$/, '')
    .replace(/\s{2,}/g,' ')
    .trim()
    .replace(/^(?:el|la|para el|para)\s+(?=\S)/i,'')
    .replace(/(?:\s+(?:el|la|los|las|para|del?|hasta|antes|a|las?|en|y|que|vence|entrega(?:r)?\s+el))+\s*$/i,'')
    .replace(/[\s\-–—|:;,(]+$/,'')
    .replace(/^\p{Ll}/u, ch=>ch.toUpperCase());
}
// Materia a partir del título de un documento: "Plan de estudio — Biología" → "Biología"
const GENERIC = /^(cronograma|calendario|plan(?: de (?:estudio|trabajo|evaluacion))?|horario|programa(?: del curso)?|planificacion|actividades|tareas|evaluaciones|agenda|semestre|cuatrimestre|trimestre|ciclo|i+ (?:cuatrimestre|semestre|ciclo)|curso|asignaciones|entregas)(?:\s+\S+){0,2}$/;
function docTitleSubject(line){
  const parts = String(line).replace(/^#+\s*/,'').split(/\s+[—–\-|:]\s+|:\s+/).map(x=>x.trim()).filter(Boolean);
  const good = parts.filter(p=>p.length<=40 && !GENERIC.test(norm(p)) && !/\b20\d{2}\b/.test(p));
  return good.length ? clean(good[good.length-1]) : '';
}
function stripPieces(text, pieces){
  // quita fragmentos (de mayor a menor posición) y devuelve el resto
  let s = text; pieces.filter(Boolean).sort((a,b)=>b.index-a.index).forEach(p=>{ s = s.slice(0,p.index)+' '+s.slice(p.index+p.length); });
  return s;
}
const WD = {domingo:0,dom:0,lunes:1,lun:1,martes:2,miercoles:3,mie:3,jueves:4,jue:4,viernes:5,vie:5,sabado:6,sab:6,sunday:0,monday:1,tuesday:2,wednesday:3,thursday:4,friday:5,saturday:6};
function weekdayIn(line){ const m = norm(line).match(/\b(domingo|lunes|martes|miercoles|jueves|viernes|sabado|monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/); return m ? WD[m[1]] : null; }
function dateForWeekday(range, wd){
  const [y,mo,d] = range.start.split('-').map(Number); const st = new Date(y,mo-1,d);
  for(let i=0;i<7;i++){ const x = new Date(st.getFullYear(), st.getMonth(), st.getDate()+i); if(x.getDay()===wd){ const iso = toISO(x.getFullYear(), x.getMonth()+1, x.getDate()); return iso<=range.end || !range.end ? iso : null; } }
  return null;
}
const WEEKDAY_ONLY = /^(?:el\s+)?(lunes|martes|miercoles|jueves|viernes|sabado|domingo|hoy|manana|semana\s*\d*|sem\.?\s*\d*|dia|fecha|entrega|para el|hasta el|vence|antes del?|a las|hora|hrs?)$/;

/* ---------- Tablas ---------- */
const HEAD = {
  date:   /\b(fecha|fechas|date|dia|dias|entrega|vence|vencimiento|deadline|cuando|limite|plazo|due)\b/,
  time:   /\b(hora|horario|time)\b/,
  title:  /\b(tarea|tareas|actividad|actividades|tema|temas|descripcion|asignacion|evaluacion|evaluaciones|contenido|titulo|nombre|pendiente|trabajo|task|detalle|que|item|asunto|entregable|rubro)\b/,
  subject:/\b(materia|materias|curso|cursos|asignatura|clase|subject|course|catedra|modulo|area|profesor|docente)\b/,
  status: /\b(estado|status|avance|progreso|completado|hecho|listo|check)\b/,
  start:  /\b(inicio|desde|start|apertura)\b/,
  session:/\b(sesion|sesiones|semana|semanas|week|leccion|clase n|n[°º] sesion)\b/,
};
function isHeaderRow(row, ctx){
  // una fila de encabezados no tiene fechas y sus celdas son cortas
  if(row.some(c=>c instanceof Date || (typeof c==='number' && c>30000) || (typeof c==='string' && findDate(c, ctx||{})))) return false;
  if(row.some(c=>/^\s*\d+\s*$/.test(String(c??'')))) return false; // números de sesión = fila de datos
  let hits = 0, short = 0; const used = new Set(); const filled = row.filter(c=>norm(c));
  row.forEach(c=>{ const t = norm(c); if(!t || t.length>30) return; if(t.length<=22) short++; for(const k in HEAD){ if(!used.has(k) && HEAD[k].test(t)){ hits++; used.add(k); break; } } });
  return (hits >= 2 && short >= filled.length-1) || (hits>=1 && filled.length<=2 && short===filled.length);
}
function mapHeader(row){
  const map = {};
  row.forEach((c,i)=>{
    const t = norm(c); if(!t) return;
    for(const k of ['start','date','time','subject','status','title']){
      if(map[k]==null && HEAD[k].test(t)){
        // "fecha de inicio" cuenta como inicio, no como entrega
        if(k==='date' && HEAD.start.test(t) && !/entrega|vence|limite/.test(t)) continue;
        map[k]=i; return;
      }
    }
  });
  return map;
}
function inferColumns(rows, ctx){
  const width = Math.max(...rows.map(r=>r.length));
  const stats = [];
  for(let i=0;i<width;i++){
    let dates=0, text=0, len=0, vals=new Map(), nonEmpty=0;
    for(const r of rows.slice(0,200)){
      const v = r[i]; const s = String(v??'').trim(); if(!s && !(v instanceof Date)) continue;
      nonEmpty++;
      if(dateFromCell(v, ctx)) dates++;
      else { text++; len += s.length; vals.set(norm(s),(vals.get(norm(s))||0)+1); }
    }
    stats.push({i, dates, text, avg: text? len/text : 0, distinct: vals.size, nonEmpty, statusish: [...vals.keys()].filter(k=>statusOf(k)||TODO_RE.test(k)).length});
  }
  const map = {};
  const byDates = [...stats].sort((a,b)=>b.dates-a.dates);
  if(byDates[0] && byDates[0].dates>0) map.date = byDates[0].i;
  const texty = stats.filter(s=>s.i!==map.date && s.text>0);
  const st = texty.find(s=>s.statusish >= Math.max(1, s.text*0.6)); if(st) map.status = st.i;
  const rest = texty.filter(s=>s.i!==map.status);
  const byLen = [...rest].sort((a,b)=>b.avg-a.avg);
  if(byLen[0]) map.title = byLen[0].i;
  const subj = rest.filter(s=>s.i!==map.title && s.avg<40 && s.distinct < Math.max(2, s.text*0.8)).sort((a,b)=>b.text-a.text)[0];
  if(subj) map.subject = subj.i;
  return map;
}
function parseTable(rows, ctx, defaultSubject){
  rows = rows.map(r=>Array.isArray(r)?r:[r]).filter(r=>r.some(v=>String(v??'').trim() || v instanceof Date));
  if(!rows.length) return [];
  let hIdx = -1;
  for(let i=0;i<Math.min(10,rows.length);i++){ if(isHeaderRow(rows[i], ctx)){ hIdx=i; break; } }
  let map = hIdx>=0 ? mapHeader(rows[hIdx]) : {};
  const headers = hIdx>=0 ? rows[hIdx].map(c=>String(c??'').trim()) : [];
  const data = rows.slice(hIdx+1);
  const inf = inferColumns(data, ctx);
  for(const k of ['date','title','subject','status']) if(map[k]==null && inf[k]!=null && !Object.values(map).includes(inf[k])) map[k]=inf[k];
  // ¿Tabla organizada por sesiones/semanas y sin fechas?
  if(hIdx>=0){
    const sIdx = rows[hIdx].findIndex(c=>{ const t=norm(c); return t && t.length<=15 && SESSION_HEAD.test(t); });
    const anyDate = data.some(r=>r.some(v=>dateFromCell(v, ctx)));
    if(sIdx>=0 && !anyDate){
      const label = /semana|week/.test(norm(rows[hIdx][sIdx])) ? 'Semana' : 'Sesión';
      const others = rows[hIdx].map((c,i)=>i).filter(i=>i!==sIdx);
      const temaIdx = others.find(i=>/tema|contenido|unidad|modulo/.test(norm(rows[hIdx][i]))) ?? others[0];
      const actIdx = others.find(i=>i!==temaIdx && /actividad|tarea|evaluacion|entrega|asignacion/.test(norm(rows[hIdx][i]))) ?? others.find(i=>i!==temaIdx);
      const srows = []; let cur = null;
      for(const r of data){
        const n = parseInt(String(r[sIdx]??'').trim(),10);
        const tema = String(r[temaIdx]??''), acts = actIdx!=null ? String(r[actIdx]??'') : '';
        if(!isNaN(n)){ cur = {n, label, tema, acts}; srows.push(cur); }
        else if(cur){ cur.tema += '\n'+tema; cur.acts += '\n'+acts; }
      }
      if(srows.length>=2) return expandSessions(srows, defaultSubject||'');
    }
  }
  const out = []; let lastDate=null, lastSubject=defaultSubject||'';
  for(const r of data){
    // fila que es sólo un título de sección (p. ej. nombre de materia o "Semana 3")
    const filled = r.map((v,i)=>({v,i})).filter(o=>String(o.v??'').trim() || o.v instanceof Date);
    if(filled.length===1 && map.title!=null){
      const only = filled[0].v; const d = dateFromCell(only, ctx);
      if(d){ lastDate = d; continue; }
      const t = String(only).trim();
      if(t.length<=40 && filled[0].i!==map.title){ lastSubject = clean(t); continue; }
    }
    let d = map.date!=null ? dateFromCell(r[map.date], ctx) : null;
    if(!d){ for(let i=0;i<r.length && !d;i++){ if(i!==map.title) d = dateFromCell(r[i], ctx); } }
    if(!d && map.date!=null && !String(r[map.date]??'').trim()) d = lastDate; // celdas combinadas
    if(d) lastDate = d;
    let time = d && d.time || '';
    if(map.time!=null){ const tt = findTime(String(r[map.time]??'')); if(tt) time = tt.time; }
    let title = map.title!=null ? String(r[map.title]??'').trim() : '';
    if(!title){ title = r.filter((v,i)=>i!==map.date && i!==map.subject && i!==map.status && i!==map.time && typeof v==='string' && v.trim() && !dateFromCell(v,ctx)).join(' · '); }
    if(map.title!=null && map.date==null){ const inT = findDate(title, ctx); if(inT && !d){ d = {date:inT.date}; title = stripPieces(title,[inT]); } }
    title = clean(title);
    if(!title || title.length<2) continue;
    let subject = map.subject!=null ? clean(r[map.subject]) : '';
    if(!subject) subject = lastSubject; else lastSubject = subject;
    let status = map.status!=null ? statusOf(String(r[map.status]??'')) : '';
    if(!status && map.status==null) status = statusOf(title)==='done' && /^[✔✓☑✅]/.test(String(r[map.title]??'').trim()) ? 'done' : '';
    const notes = [];
    r.forEach((v,i)=>{ if([map.date,map.title,map.subject,map.status,map.time].includes(i)) return; const s = v instanceof Date ? '' : String(v??'').trim(); if(s && s.length<300) notes.push((headers[i] ? headers[i]+': ' : '') + s); });
    out.push({title, subject, date: d ? d.date : '', time, status, desc: notes.join('\n')});
  }
  return out;
}

/* ---------- Texto libre (PDF, Word, pegado) ---------- */
function looksLikeHeading(line){
  const t = line.trim();
  if(!t || t.length>50) return false;
  if(/[:：]$/.test(t)) return true;
  if(/^#{1,4}\s/.test(t)) return true;
  const letters = t.replace(/[^A-Za-zÁÉÍÓÚÑáéíóúñ]/g,'');
  return letters.length>=3 && letters === letters.toUpperCase();
}
function splitTitleSubject(rest, knownSubjects){
  const parts = rest.split(/\s*[|;\t]\s*|\s+[-–—]\s+|\s*:\s+(?=\S)/).filter(p=>!WEEKDAY_ONLY.test(norm(clean(p.replace(/^[\s•*\-–—]+/,'')).replace(/^\s+|\s+$/g,''))) && !/^\s*[•*\-–—]?\s*(a|a las|al|el|la|para|para el|de|del|hasta|antes de|a la)\s*$/i.test(p)).map(clean).filter(p=>p && p.length>=2);
  if(!parts.length) return {title:'', subject:''};
  if(parts.length===1) return {title:parts[0], subject:''};
  // si alguna parte es una materia conocida, esa es la materia
  const ki = parts.findIndex(p=>knownSubjects.has(norm(p)));
  if(ki>=0){ const subject = parts[ki]; parts.splice(ki,1); return {title:parts.join(' - '), subject}; }
  // "Cálculo: Examen parcial" → materia corta antes de los dos puntos
  const colon = rest.match(/^\s*([^:]{2,30}):\s+(.+)$/);
  if(colon && colon[1].split(/\s+/).length<=4){ return {title:clean(colon[2]), subject:clean(colon[1])}; }
  const [a,b,...more] = parts;
  // la parte más corta (si es breve) suele ser la materia
  if(b.length<=30 && a.length>b.length) return {title:a, subject:b, extra:more.join(' · ')};
  if(a.length<=30 && b.length>a.length*1.5) return {title:b, subject:a, extra:more.join(' · ')};
  return {title:a, subject:b, extra:more.join(' · ')};
}
function parseText(text, ctx, defaultSubject){
  const rawLines = String(text).replace(/\r/g,'').split('\n');
  // bloques con tabulaciones o barras se tratan como tabla
  const tabRows = rawLines.filter(l=>/\t|\s\|\s|^\|/.test(l));
  if(tabRows.length >= 3 && tabRows.length >= rawLines.filter(l=>l.trim()).length*0.5){
    const first = rawLines.find(l=>l.trim()) || '';
    const titleSubj = !/\t|\|/.test(first) && !findDate(first, ctx) ? docTitleSubject(first) : '';
    const rows = rawLines.filter(l=>l.trim()).map(l=>l.replace(/^\||\|$/g,'').split(/\t|\s*\|\s*/));
    const t = parseTable(rows.filter(r=>!r.every(c=>/^[-:\s]*$/.test(c))), ctx, defaultSubject || titleSubj);
    if(t.length) return t;
  }
  const known = new Set((ctx.knownSubjects||[]).map(norm));
  const out = []; let ctxDate = null, ctxRange = null, ctxSubject = defaultSubject || '';
  const firstIdx = rawLines.findIndex(l=>l.trim());
  for(let i=0;i<rawLines.length;i++){
    const line = rawLines[i].trim(); if(!line) continue;
    // primera línea del documento = título → posible materia
    if(i===firstIdx && line.length<=70 && !findDate(line, ctx)){ const sj = docTitleSubject(line); if(sj && !ctxSubject){ ctxSubject = sj; known.add(norm(sj)); } if(!/[|\t]/.test(line)) continue; }
    const d = findDate(line, ctx);
    const rest0 = d ? stripPieces(line,[d]) : line;
    const t = findTime(rest0);
    let rest = t ? stripPieces(rest0,[t]) : rest0;
    const restClean = clean(rest);
    // línea que es sólo una fecha → encabezado de fecha para las siguientes
    if(d && (restClean.length<3 || WEEKDAY_ONLY.test(norm(restClean)) || /^(semana|sem\.?)\s*\d*/i.test(norm(restClean)))){ ctxDate = {date:d.date, time: t?t.time:''}; ctxRange = d.start ? {start:d.start, end:d.date} : null; continue; }
    if(!d){
      if(looksLikeHeading(line) && !statusOf(line)){ ctxSubject = clean(line.replace(/^#+\s*/,'').replace(/[:：]$/,'')); known.add(norm(ctxSubject)); ctxDate = null; continue; }
      if(!ctxDate) continue;                         // texto suelto sin fecha: se ignora
      if(line.length>140) continue;                  // párrafos largos: no son tareas
    }
    const status = statusOf(line);
    const cleaned = rest.replace(DONE_RE,' ').replace(PROG_RE,' ').replace(/\(\s*\)/g,' ');
    const ts = splitTitleSubject(cleaned, known);
    if(!ts.title || ts.title.length<2) continue;
    let dd = d ? {date:d.date} : ctxDate;
    if(!d && ctxRange){ const wd = weekdayIn(line); const x = wd!=null ? dateForWeekday(ctxRange, wd) : null; if(x) dd = {date:x}; }
    out.push({title: ts.title, subject: ts.subject || ctxSubject, date: dd ? dd.date : '', time: t ? t.time : (dd && dd.time) || '', status, desc: [ts.extra, d && d.start ? 'Desde el '+d.start.split('-').reverse().join('/') : ''].filter(Boolean).join('\n')});
  }
  return out;
}

/* ---------- Cronogramas por SESIONES / SEMANAS (sin fechas) ---------- */
const NOT_GRADED = /entrega de (notas|calificaciones|promedios)|cierre de curso|bienvenida/i;
const GRADED_RE = /(valor\s*:?\s*\d+\s*%|\d+\s*%|\bprueba\b|\bparcial\b|\bexamen\b|\bquiz\b|\bentrega\b|\bautoevaluaci[oó]n\b|\bevaluaci[oó]n\b|\bproyecto final\b|\bdefensa\b)/i;
// une renglones que el PDF/Word cortó a mitad de frase
const GRADED = { test: t => GRADED_RE.test(t) && !NOT_GRADED.test(t) };
function joinWrapped(text){
  const lines = String(text||'').split('\n').map(l=>l.trim()).filter(Boolean);
  const out = [];
  for(const l of lines){
    const bullet = /^[•·\-–*▪◦●]/.test(l);
    const t = l.replace(/^[•·\-–*▪◦●]\s*/,'');
    const prev = out[out.length-1];
    const startsLower = /^[a-záéíóúñ(]/.test(t);
    const isLabel = /^m[oó]dulo\s+\d+\s*\.?/i.test(t);
    if(prev && !bullet && !isLabel && (startsLower || !/[.%:;!?)]$/.test(prev.t)) && !/^m[oó]dulo\s+\d+\s*\.\s*\S+$/i.test(prev.t) && !(GRADED.test(t) && /^[A-ZÁÉÍÓÚ]/.test(t))){
      prev.t = (prev.t + ' ' + t).replace(/\s+/g,' ');
    } else out.push({t, bullet});
  }
  return out;
}
function prettyCase(sj){
  const t = String(sj||'').trim();
  if(!t || t !== t.toUpperCase() || !/[A-ZÁÉÍÓÚ]{4}/.test(t)) return t;
  return t.toLowerCase().replace(/(^|\s|\()(\p{L})/gu, (m,a,b)=>a+b.toUpperCase()).replace(/\b(i{1,3}|iv|v|vi{1,3}|ix|x)\b/gi, x=>x.toUpperCase()).replace(/\b(De|Del|La|Las|Los|El|Y|E|En|Para|Por|A)\b/g, x=>x.toLowerCase()).replace(/^\p{L}/u, x=>x.toUpperCase());
}
function cleanGraded(t){
  let x = t.replace(/\s*[:.]\s*(sin puntaje|no tiene puntaje)\.?\s*$/i,'');
  const v = x.match(/\.?\s*valor\s*:?\s*(\d+\s*%)/i);
  if(v) x = x.replace(v[0],'').trim().replace(/[.:,]$/,'') + ' ('+v[1].replace(/\s/g,'')+')';
  return clean(x);
}
// rows: [{n, tema, acts}] → tareas con sesión (las fechas se calculan después)
function expandSessions(rows, subject){
  const out = [];
  for(const r of rows){
    const tema = joinWrapped(r.tema), acts = joinWrapped(r.acts);
    const seen = new Set(), graded = [], todo = [];
    const keyOf = t => norm(t).replace(/[:.].*$/,'').replace(/\s+/g,' ').trim();
    for(const it of acts.concat(tema)){
      const t = it.t;
      if(/^m[oó]dulo\s+\d+\s*\.\s*(lectura|actividad calificable)\s*$/i.test(t)) continue; // etiquetas, no tareas
      if(GRADED.test(t)){ const k = keyOf(t); if(seen.has(k)) continue; seen.add(k); graded.push(cleanGraded(t)); }
    }
    for(const it of acts){ if(!GRADED.test(it.t)) todo.push(clean(it.t)); }
    const topics = tema.filter(x=>!GRADED.test(x.t) && !/^m[oó]dulo\s+\d+\s*\.\s*(lectura|actividad calificable)\s*$/i.test(x.t)).map(x=>clean(x.t));
    const modLabel = (tema.find(x=>/^m[oó]dulo\s+\d+/i.test(x.t))||{}).t || '';
    const main = (tema.find(x=>x.bullet && !GRADED.test(x.t)) || {}).t || topics[0] || todo[0] || modLabel || '';
    const label = r.label || 'Sesión';
    if(todo.length || topics.length || !graded.length){
      out.push({session:r.n, offset:0, kind:'session', title: label+' '+r.n+' · '+clean(main||'Clase'), subject,
        desc: [modLabel ? clean(modLabel) : '', topics.length ? 'Temas:\n'+topics.map(t=>'• '+t).join('\n') : ''].filter(Boolean).join('\n\n'),
        checklist: todo.filter(t=>t && norm(t)!==norm(main)), date:'', time:'', status:''});
    }
    for(const g of graded){
      out.push({session:r.n, offset:'due', kind:'graded', title:g, subject, desc: label+' '+r.n + (modLabel ? ' · '+clean(modLabel) : ''), checklist:[], date:'', time:'', status:''});
    }
  }
  return out;
}
const SESSION_HEAD = /\b(sesion|sesiones|semana|semanas|clase|week|leccion|unidad|n[°º.]?|#)\b/;

/* ---------- Año predominante del documento ---------- */
function detectYear(text){
  const c = {}; (String(text).match(/\b20\d{2}\b/g)||[]).forEach(y=>c[y]=(c[y]||0)+1);
  const best = Object.entries(c).sort((a,b)=>b[1]-a[1])[0];
  const y = best ? +best[0] : null; const now = new Date().getFullYear();
  return y && Math.abs(y-now)<=1 ? y : null;
}

/* ---------- Lectores por tipo de archivo ---------- */
function decodeText(buf){
  try{ return new TextDecoder('utf-8', {fatal:true}).decode(buf).replace(/^\uFEFF/,''); }
  catch(e){ return new TextDecoder('windows-1252').decode(buf); }
}
async function readSpreadsheet(file, ctx){
  await loadLib('xlsx');
  const buf = await file.arrayBuffer();
  const isText = /\.(csv|tsv)$/i.test(file.name);
  const wb = isText ? XLSX.read(decodeText(buf), {type:'string', cellDates:true, raw:false})
                    : XLSX.read(buf, {type:'array', cellDates:true, dateNF:'yyyy-mm-dd'});
  let out = [];
  const multi = wb.SheetNames.length > 1;
  const allText = wb.SheetNames.map(n=>XLSX.utils.sheet_to_csv(wb.Sheets[n])).join('\n');
  ctx.year = ctx.year || detectYear(allText);
  for(const name of wb.SheetNames){
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], {header:1, raw:true, defval:''});
    const generic = /^(hoja|sheet|tabla|table)\s*\d*$/i.test(name.trim());
    out = out.concat(parseTable(rows, ctx, multi && !generic ? name.trim() : ''));
  }
  return out;
}
async function readDocx(file, ctx){
  await loadLib('mammoth');
  const res = await mammoth.convertToHtml({arrayBuffer: await file.arrayBuffer()});
  const doc = new DOMParser().parseFromString(res.value, 'text/html');
  ctx.year = ctx.year || detectYear(doc.body.textContent);
  let out = [];
  // tablas
  doc.querySelectorAll('table').forEach(tb=>{
    const rows = [...tb.querySelectorAll('tr')].map(tr=>[...tr.children].map(td=>td.textContent.trim()));
    let h = tb.previousElementSibling, subj = '';
    while(h && !subj){ if(/^H\d$/.test(h.tagName) || (h.tagName==='P' && h.textContent.trim().length<=60 && looksLikeHeading(h.textContent))) subj = docTitleSubject(h.textContent); h = h.previousElementSibling; }
    out = out.concat(parseTable(rows, ctx, subj));
    tb.remove();
  });
  // resto del texto (párrafos, listas, encabezados)
  const lines = [...doc.body.querySelectorAll('h1,h2,h3,h4,p,li')].map(el=>(/^H\d$/.test(el.tagName)?'# ':'')+el.textContent.trim());
  return out.concat(parseText(lines.join('\n'), ctx, ''));
}
function mul(m, n){ return [m[0]*n[0]+m[2]*n[1], m[1]*n[0]+m[3]*n[1], m[0]*n[2]+m[2]*n[3], m[1]*n[2]+m[3]*n[3], m[0]*n[4]+m[2]*n[5]+m[4], m[1]*n[4]+m[3]*n[5]+m[5]]; }
function apply(m, x, y){ return [m[0]*x+m[2]*y+m[4], m[1]*x+m[3]*y+m[5]]; }
// Extrae las líneas (bordes de tabla) dibujadas en una página
async function pageSegments(page){
  const OPS = pdfjsLib.OPS; const ol = await page.getOperatorList();
  let ctm = [1,0,0,1,0,0]; const stack = []; const H = [], V = [];
  const seg = (a,b)=>{ const [x1,y1]=apply(ctm,a[0],a[1]), [x2,y2]=apply(ctm,b[0],b[1]);
    if(Math.abs(y1-y2)<1.5 && Math.abs(x2-x1)>25) H.push({y:(y1+y2)/2, x1:Math.min(x1,x2), x2:Math.max(x1,x2)});
    else if(Math.abs(x1-x2)<1.5 && Math.abs(y2-y1)>8) V.push({x:(x1+x2)/2, y1:Math.min(y1,y2), y2:Math.max(y1,y2)}); };
  for(let i=0;i<ol.fnArray.length;i++){
    const fn = ol.fnArray[i], a = ol.argsArray[i];
    if(fn===OPS.save) stack.push(ctm.slice());
    else if(fn===OPS.restore) ctm = stack.pop() || [1,0,0,1,0,0];
    else if(fn===OPS.transform) ctm = mul(ctm, a);
    else if(fn===OPS.constructPath){
      const ops = a[0], args = a[1]; let k = 0, cur = null, start = null;
      for(const op of ops){
        if(op===OPS.rectangle){ const [x,y,w,h] = args.slice(k,k+4); k+=4;
          const p1=[x,y],p2=[x+w,y],p3=[x+w,y+h],p4=[x,y+h];
          if(Math.abs(h)<3){ seg([x,y+h/2],[x+w,y+h/2]); } else if(Math.abs(w)<3){ seg([x+w/2,y],[x+w/2,y+h]); } else { seg(p1,p2); seg(p2,p3); seg(p3,p4); seg(p4,p1); } }
        else if(op===OPS.moveTo){ cur = [args[k],args[k+1]]; start = cur; k+=2; }
        else if(op===OPS.lineTo){ const nx=[args[k],args[k+1]]; k+=2; if(cur) seg(cur,nx); cur = nx; }
        else if(op===OPS.curveTo){ k+=6; } else if(op===OPS.curveTo2 || op===OPS.curveTo3){ k+=4; }
        else if(op===OPS.closePath){ if(cur && start) seg(cur,start); cur = start; }
      }
    }
  }
  return {H, V};
}
function clusterVals(vals, tol){ vals = vals.slice().sort((a,b)=>a-b); const out=[]; for(const v of vals){ if(out.length && v-out[out.length-1].last<=tol){ const c=out[out.length-1]; c.sum+=v; c.n++; c.last=v; } else out.push({sum:v,n:1,last:v}); } return out.map(c=>c.sum/c.n); }
// Lee una tabla con bordes: devuelve filas de celdas (texto con saltos de línea)
async function pdfGridRows(pdf){
  const all = [];
  for(let p=1;p<=pdf.numPages;p++){
    const page = await pdf.getPage(p);
    const seg = await pageSegments(page);
    const vp = page.getViewport({scale:1}); const PW = vp.width, PH = vp.height;
    // descarta bordes de la página y líneas de adornos: se quedan las líneas largas de la tabla
    const V = seg.V.filter(v=> v.x>4 && v.x<PW-4 && (v.y2-v.y1)<PH*0.95);
    const H = seg.H.filter(h=> h.y>4 && h.y<PH-4 && (h.x2-h.x1)<PW*0.98);
    if(H.length < 3 || V.length < 2) return null;
    const groupLen = (arr, key, len) => { const g = new Map(); for(const a of arr){ const k = Math.round(a[key]/3); g.set(k, (g.get(k)||0) + len(a)); } return g; };
    const vg = groupLen(V, 'x', v=>v.y2-v.y1); const vmax = Math.max(...vg.values());
    const goodV = V.filter(v=> vg.get(Math.round(v.x/3)) >= vmax*0.08);
    const tx1 = Math.min(...goodV.map(v=>v.x)), tx2 = Math.max(...goodV.map(v=>v.x));
    const ty1 = Math.min(...goodV.map(v=>v.y1)), ty2 = Math.max(...goodV.map(v=>v.y2));
    const inH = H.filter(h=> h.y>=ty1-2 && h.y<=ty2+2 && h.x2>tx1+5 && h.x1<tx2-5);
    const hg = groupLen(inH, 'y', h=>Math.min(h.x2,tx2)-Math.max(h.x1,tx1));
    const goodH = inH.filter(h=> hg.get(Math.round(h.y/3)) >= (tx2-tx1)*0.5);
    const xs = clusterVals(goodV.map(v=>v.x), 3), ys = clusterVals(goodH.map(h=>h.y), 3).sort((a,b)=>b-a);
    if(xs.length < 3 || ys.length < 3) return null;
    const tc = await page.getTextContent();
    const cells = new Map();
    for(const it of tc.items){
      if(!it.str || !it.str.trim()) continue;
      const x = it.transform[4] + 1, y = it.transform[5] + 2;
      const ci = xs.findIndex((v,i)=> i<xs.length-1 && x>=v && x<xs[i+1]);
      const ri = ys.findIndex((v,i)=> i<ys.length-1 && y<=v && y>ys[i+1]);
      if(ci<0 || ri<0) continue;
      const key = ri+'|'+ci; if(!cells.has(key)) cells.set(key, []);
      cells.get(key).push({x:it.transform[4], y:it.transform[5], w:it.width, s:it.str});
    }
    for(let ri=0; ri<ys.length-1; ri++){
      const row = [];
      for(let ci=0; ci<xs.length-1; ci++){
        const items = (cells.get(ri+'|'+ci)||[]).sort((a,b)=> Math.abs(a.y-b.y)<3 ? a.x-b.x : b.y-a.y);
        let txt = '', ly = null, end = null;
        for(const it of items){
          if(ly!==null && Math.abs(it.y-ly)>=3){ txt += '\n'; end = null; }
          else if(end!==null){ const gap = it.x - end; txt += gap > 1.5 ? ' ' : ''; }
          txt += it.s; ly = it.y; end = it.x + it.w;
        }
        row.push(txt.replace(/[ \t]+/g,' ').trim());
      }
      if(row.some(c=>c)) all.push({page:p, cells:row});
    }
  }
  return all.length >= 3 ? all : null;
}
async function readPdf(file, ctx){
  await loadLib('pdf');
  await window.cargarPdfjs();
  const pdf = await pdfjsLib.getDocument({data: new Uint8Array(await file.arrayBuffer()), isEvalSupported: false}).promise;
  // 1) Si la tabla tiene bordes dibujados, se lee celda por celda
  try{
    const grid = await pdfGridRows(pdf);
    if(grid){
      const firstPage = await (await pdf.getPage(1)).getTextContent();
      const topText = firstPage.items.map(i=>i.str).join(' ');
      ctx.year = ctx.year || detectYear(topText);
      // título del documento (texto fuera de la tabla, arriba) → materia
      const titleLine = firstPage.items.filter(i=>i.str.trim()).sort((a,b)=>b.transform[5]-a.transform[5]);
      const topY = titleLine.length ? titleLine[0].transform[5] : 0;
      const title = titleLine.filter(i=>Math.abs(i.transform[5]-topY)<3).sort((a,b)=>a.transform[4]-b.transform[4]).map(i=>i.str).join(' ').replace(/\s*-\s*/g,'-').replace(/\s+/g,' ');
      let subject = prettyCase(docTitleSubject(title.replace(/^([A-Z]{1,4}\d+[-–]\d+)\s*:?\s*/,'')) || '');
      // filas: quita encabezados repetidos y une filas cortadas entre páginas
      const rows = []; let header = null;
      for(const g of grid){
        if(isHeaderRow(g.cells, ctx) || (header && g.cells.join('|')===header.join('|'))){ header = header || g.cells; continue; }
        if(rows.length && !g.cells[0] && g.page !== rows[rows.length-1].page){ const prev = rows[rows.length-1]; prev.cells = prev.cells.map((c,i)=> (c + '\n' + (g.cells[i]||'')).trim()); continue; }
        rows.push({page:g.page, cells:g.cells});
      }
      if(header && rows.length){
        const res = parseTable([header].concat(rows.map(r=>r.cells)), ctx, subject);
        if(res.length) return res;
      }
    }
  }catch(e){ /* si falla, se usa la lectura por renglones */ }
  const lines = [];
  for(let p=1; p<=pdf.numPages; p++){
    const page = await pdf.getPage(p);
    const tc = await page.getTextContent();
    // agrupa los trozos de texto por renglón (coordenada y) y separa columnas con tabulador
    const rows = [];
    for(const it of tc.items){
      if(!it.str || !it.str.trim()) continue;
      const x = it.transform[4], y = it.transform[5];
      let row = rows.find(r=>Math.abs(r.y - y) < 3);
      if(!row){ row = {y, items:[]}; rows.push(row); }
      row.items.push({x, w: it.width, s: it.str});
    }
    rows.sort((a,b)=>b.y-a.y);
    for(const r of rows){
      r.items.sort((a,b)=>a.x-b.x);
      let s = '', end = null;
      for(const it of r.items){
        if(end!==null){ const gap = it.x - end; s += gap > 14 ? '\t' : (gap > 1.5 ? ' ' : ''); }
        s += it.s; end = it.x + it.w;
      }
      lines.push(s);
    }
  }
  const text = lines.join('\n');
  ctx.year = ctx.year || detectYear(text);
  // si hay muchas líneas con columnas, se lee como tabla; si no, como texto
  return parseText(text, ctx, '');
}
async function readFile(file, ctx){
  ctx = ctx || {};
  if(file.size > MAX_ARCHIVO) throw new Error('El archivo pesa más de 20 MB. Un cronograma suele pesar mucho menos: revisa que sea el archivo correcto.');
  const name = file.name.toLowerCase();
  if(/\.(xlsx|xlsm|xls|ods|csv|tsv)$/.test(name)) return readSpreadsheet(file, ctx);
  if(/\.docx$/.test(name)) return readDocx(file, ctx);
  if(/\.pdf$/.test(name)) return readPdf(file, ctx);
  if(/\.(txt|md|text)$/.test(name) || file.type.startsWith('text/')){ const t = decodeText(await file.arrayBuffer()); ctx.year = ctx.year || detectYear(t); return parseText(t, ctx, ''); }
  if(/\.doc$/.test(name)) throw new Error('Los archivos .doc antiguos no se pueden leer. En Word usa “Guardar como” → .docx o PDF.');
  if(/\.(png|jpe?g|heic|webp|gif)$/.test(name) || file.type.startsWith('image/')) throw new Error('Las fotos todavía no se pueden leer. Copia el texto del cronograma y pégalo en el cuadro, o súbelo en PDF, Word o Excel.');
  throw new Error('Formato no reconocido. Usa Excel, CSV, Word (.docx), PDF o texto.');
}
function readPasted(text, ctx){ ctx = ctx || {}; ctx.year = ctx.year || detectYear(text); return parseText(text, ctx, ''); }

window.SmartImport = { readFile, readPasted, parseText, parseTable, findDate, findTime, statusOf };
})();
