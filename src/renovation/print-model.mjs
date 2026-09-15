import {departmentTone} from './colors.mjs';
export const PRINT_DAYS=['จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์','อาทิตย์'];
export const FIELDS={subject_name:'ชื่อวิชา',subject_short:'ชื่อวิชา (ย่อ)',subject_code:'รหัสวิชา',teacher_name:'ชื่อครูและครูร่วม',teacher_fname:'ชื่อครู (ชื่อต้น)',teacher_code:'รหัสครู',room_name:'ห้องเรียน',special_room:'ห้องพิเศษ',period_time:'เวลา',period_num:'คาบที่',custom_text:'ข้อความเอง'};
export const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const num=(n,a,b,d)=>Number.isFinite(+n)?Math.min(b,Math.max(a,+n)):d;
const color=(s,d)=>/^#[a-f0-9]{6}$/i.test(s||'')?s:d;
export function makeTemplate(level,config){
 const columns=[];
 const addBreak=pid=>(config.breaks||[]).filter(b=>b.afterPid===pid).forEach((b,i)=>columns.push({id:`b${pid}-${i}`,type:pid===0?'assembly':'break',label:pid===0?'กิจกรรมเช้า':'พัก',time:b.label,width:45,show:pid!==0,text:'',days:[]}));
 addBreak(0);config.periods.forEach(p=>{columns.push({id:'p'+p.id,type:'period',periodId:p.id,label:'คาบ '+p.id,time:p.time,width:100,show:true});addBreak(p.id)});
 return {name:'แบบพิมพ์ '+(level?.name||'ครู'),columns:level?columns:columns.map(c=>c.type==='period'?{...c,time:''}:c),days:PRINT_DAYS.slice(0,5),fields:level?['subject_short','teacher_name','room_name']:['subject_short','room_name','period_time'],paper:'A4',orientation:'landscape',margin:10,fontSize:12,rowHeight:78,colored:true,title:'',subtitle:'',footer:'',showLogo:true};
}
export function migrateLegacy(raw,base){return {...base,columns:(raw.columns||[]).filter(c=>c.type!=='day').map((c,i)=>({...c,id:c.id||'old'+i,time:c.timeLabel||'',text:c.label||'',days:[]})),fields:(raw.cellRows||[]).map(r=>r.field).filter(f=>FIELDS[f]),paper:raw.paperSize||'A4',orientation:raw.orientation||'landscape',margin:raw.marginMm??10,fontSize:12,title:raw.titleText||'',subtitle:raw.subtitleText||'',footer:[raw.footerLeft,raw.footerRight].filter(Boolean).join('     ')};}
export function entriesAt(S,type,id,day,pid){
 const all=[];
 Object.entries(S.schedule||{}).forEach(([key,entries])=>{const parts=key.split('_'),p=parts.pop(),d=parts.pop(),rid=parts.join('_');if(d!==day||String(pid)!==p||(type==='room'&&rid!==id))return;
 (Array.isArray(entries)?entries:[]).forEach(e=>{const co=e.coTeacherIds?.length?e.coTeacherIds:e.coTeacherId?[e.coTeacherId]:[];if(type==='teacher'&&![e.teacherId,...co].includes(id))return;all.push({...e,roomId:rid,periodId:pid,coTeacherIds:co})});});return all;
}
export function templateIssues(t,S,type,id){
 const issues=[];const cols=t.columns.filter(c=>c.show),periods=cols.filter(c=>c.type==='period').map(c=>String(c.periodId));
 if(!t.days.length)issues.push('เลือกวันอย่างน้อยหนึ่งวัน');if(!cols.length)issues.push('เปิดแสดงอย่างน้อยหนึ่งคอลัมน์');if(!t.fields.length)issues.push('เลือกข้อมูลในช่องอย่างน้อยหนึ่งรายการ');
 if(new Set(periods).size!==periods.length)issues.push('มีคาบซ้ำในแบบพิมพ์ กรุณาตรวจหมายเลขคาบ');
 if(periods.some(p=>!Number.isInteger(+p)||+p<1||+p>20))issues.push('หมายเลขคาบต้องเป็นจำนวนเต็ม 1–20');
 for(const day of PRINT_DAYS)for(let p=1;p<=20;p++)if(entriesAt(S,type,id,day,p).length&&(!t.days.includes(day)||!periods.includes(String(p)))){issues.push('มีคาบที่จัดไว้ถูกซ่อนจากแบบพิมพ์ กรุณาเพิ่มวันหรือคอลัมน์ให้ครบ');return issues;}
 return issues;
}
function fieldValue(f,e,S,c){const sub=S.subjects.find(s=>s.id===e.subjectId),ts=[e.teacherId,...e.coTeacherIds].map(id=>S.teachers.find(t=>t.id===id)).filter(Boolean);
 return {subject_name:sub?.name,subject_short:sub?.shortName||sub?.name,subject_code:sub?.code,teacher_name:ts.map(t=>`${t.firstName||''} ${t.lastName||''}`.trim()).join(' / '),teacher_fname:ts.map(t=>t.firstName).join(' / '),teacher_code:ts.map(t=>t.teacherCode).filter(Boolean).join(' / '),room_name:S.rooms.find(r=>r.id===e.roomId)?.name,special_room:S.specialRooms?.find(r=>r.id===e.specialRoomId)?.name,period_time:S.printPeriodConfigs?.[e.roomId]?.periods.find(p=>String(p.id)===String(e.periodId))?.time||c.time,period_num:'คาบ '+e.periodId,custom_text:c.text||''}[f]||'';
}
export function printDocument(pages,S,ay={},sh={}){
 const styles=[],sheets=pages.map(({template:t,type,id},index)=>{
 const paper=['A4','A3','Letter'].includes(t.paper)?t.paper:'A4',land=t.orientation!=='portrait',dims=paper==='A3'?[297,420]:paper==='Letter'?[216,279]:[210,297],w=land?dims[1]:dims[0],h=land?dims[0]:dims[1],margin=num(t.margin,4,30,10);
 styles.push(`@page sheet${index}{size:${paper} ${land?'landscape':'portrait'};margin:${margin}mm}.sheet${index}{page:sheet${index};width:${w-margin*2}mm;min-height:${h-margin*2}mm;font-size:${num(t.fontSize,8,24,12)}px}.sheet${index} tbody td{height:${num(t.rowHeight,30,180,78)}px}`);
 const target=type==='room'?S.rooms.find(r=>r.id===id):S.teachers.find(r=>r.id===id),level=S.levels.find(l=>l.id===target?.levelId),name=type==='room'?target?.name:`${target?.firstName||''} ${target?.lastName||''}`;
 const fill=text=>String(text||'').replaceAll('{ห้อง}',type==='room'?name:'').replaceAll('{ครู}',type==='teacher'?name:'').replaceAll('{ระดับชั้น}',level?.name||'').replaceAll('{ปี}',String(ay.year||''));
 const cols=t.columns.filter(c=>c.show),head=cols.map(c=>`<th>${esc(c.label)}<small>${esc(c.time)}</small></th>`).join('');
 const rows=t.days.map(day=>`<tr><th scope="row">${esc(day)}</th>${cols.map(c=>{if(c.type!=='period'){const show=!c.days?.length||c.days.includes(day);const val=c.type==='assembly'?(level?.assemblyDay===day?'เข้าหอประชุม':c.text||'กิจกรรมเช้า'):c.text||c.label;return `<td class="activity">${show?esc(val):''}</td>`;}const entries=entriesAt(S,type,id,day,c.periodId);return `<td>${entries.map(e=>{const sub=S.subjects.find(s=>s.id===e.subjectId),tone=departmentTone(S.depts.find(d=>d.id===sub?.departmentId));return `<div class="entry" style="${t.colored?`background:${tone.bg};border-left:3px solid ${tone.ink}`:''}">${t.fields.map((f,i)=>`<div class="${i===0?'first':''}">${esc(f==='custom_text'?t.customText:fieldValue(f,e,S,c))}</div>`).join('')}</div>`}).join('')}</td>`}).join('')}</tr>`).join('');
 const logo=t.showLogo&&/^data:image\/(png|jpeg|webp);base64,[a-z0-9+/=]+$/i.test(sh.logo||'')?`<img alt="ตราโรงเรียน" src="${sh.logo}">`:'';
 return `<section class="sheet sheet${index}"><header>${logo}<div><h1>${esc(fill(t.title)||((type==='room'?'ตารางเรียน ':'ตารางสอน ')+name))}</h1><p>${esc(fill(t.subtitle)||[sh.schoolName||'โรงเรียนดาราวิทยาลัย',`ปีการศึกษา ${ay.year||sh.year||''}`,(ay.semester||ay.term)?`ภาคเรียนที่ ${ay.semester||ay.term}`:''].filter(Boolean).join(' · '))}</p></div></header><table><colgroup><col style="width:65px">${cols.map(c=>`<col style="width:${num(c.width,25,250,100)}px">`).join('')}</colgroup><thead><tr><th>วัน / เวลา</th>${head}</tr></thead><tbody>${rows}</tbody></table>${t.footer?`<footer>${esc(fill(t.footer))}</footer>`:''}</section>`;
 }).join('');
 return `<!doctype html><html lang="th"><head><meta charset="utf-8"><title>ตารางสอน DARA</title><style>*{box-sizing:border-box}body{margin:0;color:#253348;font-family:Sarabun,Tahoma,sans-serif;background:#e8ebf0}.sheet{background:white;padding:0;margin:20px auto;break-after:page}.sheet:last-child{break-after:auto}header{display:flex;align-items:center;justify-content:center;gap:16px;text-align:center;margin-bottom:18px}header img{width:52px;height:52px;object-fit:contain}h1{font-size:1.55em;margin:0 0 6px}p{margin:0}table{border-collapse:collapse;width:100%;table-layout:fixed}th,td{border:1px solid #b7c2cf;padding:5px;overflow-wrap:anywhere}th{background:#f0f2f6;font-weight:600}small{display:block;font-size:.8em;margin-top:5px}td{vertical-align:middle;text-align:center}.entry{padding:5px;margin:2px 0;line-height:1.5;break-inside:avoid}.entry+.entry{border-top:1px dashed #aeb9c8}.first{font-weight:700}.activity{background:#f7f7f8;color:#5b6472}footer{white-space:pre-wrap;margin-top:24px;line-height:2}thead{display:table-header-group}tr{break-inside:avoid}${styles.join('')}@media print{body{background:white}.sheet{margin:0;min-height:0}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}}</style></head><body>${sheets}</body></html>`;
}
