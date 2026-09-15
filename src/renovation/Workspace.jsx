import {departmentTone} from './colors.mjs';
import Mascot from "./Mascot.jsx";
import React, {useState} from 'react';
import {progress,teacherLoad,DAYS} from './model.mjs';
import './workspace.css';

export function Glyph({name='grid',size=20}) {
 const paths={grid:<><rect x="3" y="3" width="7" height="7" rx="2"/><rect x="14" y="3" width="7" height="7" rx="2"/><rect x="3" y="14" width="7" height="7" rx="2"/><rect x="14" y="14" width="7" height="7" rx="2"/></>,calendar:<><rect x="3" y="5" width="18" height="16" rx="3"/><path d="M7 3v4m10-4v4M3 11h18m-13 5h2m4 0h2"/></>,users:<><circle cx="9" cy="8" r="3"/><path d="M3 21v-3a6 6 0 0 1 12 0v3m1-17a3 3 0 0 1 0 6m3 11v-3a6 6 0 0 0-3-5"/></>,book:<><path d="M12 6c-3-3-7-3-10-2v15c3-1 7-1 10 2 3-3 7-3 10-2V4c-3-1-7-1-10 2Zm0 0v15"/></>,check:<><circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/></>,arrow:<path d="M5 12h14m-5-5 5 5-5 5"/>,search:<><circle cx="10" cy="10" r="6"/><path d="m15 15 5 5"/></>,menu:<path d="M4 6h16M4 12h16M4 18h16"/>,settings:<><circle cx="12" cy="12" r="4"/><path d="M12 2v3m0 14v3M2 12h3m14 0h3M5 5l2 2m10 10 2 2M5 19l2-2M17 7l2-2"/></>,plus:<path d="M12 5v14M5 12h14"/>,alert:<><path d="m12 3 10 18H2L12 3Z"/><path d="M12 9v5m0 3v1"/></>,download:<><path d="M12 3v12m-4-4 4 4 4-4M4 16v5h16v-5"/></>,clock:<><circle cx="12" cy="12" r="9"/><path d="M12 6v6l4 2"/></>,close:<path d="m6 6 12 12M6 18 18 6"/>,edit:<><path d="m15 4 5 5L9 20H4v-5L15 4Z"/><path d="m12 7 5 5"/></>};
 return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{paths[name]||paths.grid}</svg>;
}
const groups=[['ภาพรวม',[['dashboard','grid','แดชบอร์ด']]],['เตรียมข้อมูล',[['levels','grid','ระดับชั้น / ห้องเรียน'],['plans','book','แผนการเรียน'],['departments','users','กลุ่มสาระ'],['teachers','users','จัดการครู'],['subjects','book','รายวิชา'],['specialrooms','grid','ห้องพิเศษ']]],['ตารางสอน',[['assignments','edit','มอบหมายงานครู'],['homeroom','users','ครูประจำชั้น'],['meetings','clock','คาบล็อก / ประชุม'],['scheduler','calendar','จัดตารางสอน'],['swap','calendar','แลกคาบ / สอนแทน'],['reports','check','ตรวจสอบ / รายงาน']]],['ระบบ',[['settings','settings','ตั้งค่าปีการศึกษา']]]];
export function Workspace({page,setPage,div,divisions,switchDivision,ay,children,syncing,demo,user,onLogout,onAdmin,canVisit=()=>true}) {
 const [open,setOpen]=useState(false);
 const title=groups.flatMap(g=>g[1]).find(n=>n[0]===page)?.[2];
 const go=id=>{if(canVisit(id)){setPage(id);setOpen(false);}};
 return <div className="dara-workspace">
  {open&&<button className="drawer-backdrop" aria-label="ปิดเมนู" onClick={()=>setOpen(false)}/>}
  <aside className={'workspace-sidebar '+(open?'is-open':'')}>
   <a className="brand" href="#dashboard" onClick={e=>{e.preventDefault();go('dashboard')}}><span className="brand-mark">ด</span><span><strong>DARA</strong><small>ระบบจัดตารางสอน</small></span><span className="brand-version">04</span></a>
   <label className="division-label">ระดับการศึกษา<select value={div.id} onChange={e=>switchDivision(e.target.value)}>{divisions.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}</select></label>
   <nav aria-label="เมนูหลัก">{groups.map(([label,items])=><section key={label}><p className="nav-label">{label}</p>{items.map(([id,icon,label])=><button key={id} disabled={!canVisit(id)} className={'nav-item '+(page===id?'active':'')} onClick={()=>go(id)} aria-current={page===id?'page':undefined}><Glyph name={icon}/><span>{label}</span>{page===id&&<i/>}</button>)}</section>)}</nav>
   <div className="sidebar-account"><span className="avatar">{(user?.displayName||'ด')[0]}</span><div><strong>{user?.displayName||'พื้นที่ทดลองใช้งาน'}</strong><small>{demo?'ข้อมูลตัวอย่าง':user?.email}</small></div>{!demo&&<button title="ออกจากระบบ" onClick={onLogout}><Glyph name="arrow"/></button>}</div>
   {!demo&&onAdmin&&<button className="admin-link" onClick={onAdmin}>ผู้ดูแลระบบ</button>}
  </aside>
  <div className="workspace-body"><header className="workspace-topbar"><button className="mobile-menu" onClick={()=>setOpen(!open)} aria-label="เปิดเมนู"><Glyph name="menu"/></button><span className="breadcrumb">งานวิชาการ <span>/</span> <strong>{title}</strong></span><div className="topbar-right"><span className="term">ภาคเรียน {ay.semester} / {ay.year}</span><span className={'save-state '+(demo?'demo':'')}>{demo?'● ทดลอง · บันทึกในเครื่อง':syncing?'กำลังบันทึก…':'เชื่อมต่อระบบ'}</span></div></header>
  <main className={'workspace-main page-'+page} id="main-content">{demo&&<div className="preview-notice"><span><strong>เวอร์ชันทดลอง</strong> ใช้ข้อมูลสมมติ · การแก้ไขเก็บเฉพาะเบราว์เซอร์นี้</span><span>ยังไม่เชื่อมข้อมูลโรงเรียน</span></div>}{children}</main>
  </div>
 </div>;
}
export function PageHeading({eyebrow,title,description,children}) {return <div className="page-heading"><div>{eyebrow&&<p className="eyebrow">{eyebrow}</p>}<h1>{title}</h1>{description&&<p>{description}</p>}</div><div className="heading-actions">{children}</div></div>}
export function Dashboard({S,setPage,ay}) {
 const p=progress(S);
 const overloaded=S.teachers.filter(t=>teacherLoad(S,t.id)>(Number(t.totalPeriods)||0));
 const unassigned=S.teachers.filter(t=>!S.assigns.some(a=>a.teacherId===t.id));
 const stats=[['users','ครูทั้งหมด',S.teachers.length,'คน','teachers'],['grid','ห้องเรียน',S.rooms.length,'ห้อง','levels'],['calendar','ลงตารางแล้ว',p.placed,'คาบ','scheduler'],['clock','รอจัดลงตาราง',p.remaining,'คาบ','scheduler']];
 const steps=[['ข้อมูลพื้นฐาน','ครู วิชา และห้องเรียน','teachers',S.teachers.length>0&&S.rooms.length>0&&S.subjects.length>0],['มอบหมายงาน','กำหนดผู้สอนและจำนวนคาบ','assignments',S.assigns.length>0],['จัดตารางสอน','ลงคาบและตรวจเวลาว่าง','scheduler',p.total>0&&p.remaining===0],['ตรวจสอบ / ส่งออก','ตรวจคาบชนก่อนนำไปใช้','reports',false]];
 const deptColors=['#9c2638','#386b9b','#5b6e4a','#aa7531','#705c9b','#328078'];
 return <div className="modern-dashboard"><PageHeading eyebrow="ภาพรวมงานวิชาการ" title="จัดการตารางสอน" description={`ภาคเรียนที่ ${ay.semester} ปีการศึกษา ${ay.year} · เริ่มต่อจากงานที่ยังค้างอยู่`}><button className="primary" onClick={()=>setPage('scheduler')}><Glyph name="calendar"/> เปิดตารางสอน <Glyph name="arrow" size={17}/></button></PageHeading>
  <Mascot/>
  <div className="metric-grid">{stats.map(([icon,label,value,unit,page],i)=><button className={'metric-card metric-'+i} key={label} onClick={()=>setPage(page)}><div className="metric-top"><span>{label}</span><Glyph name={icon}/></div><div className="metric-value">{value}<small>{unit}</small></div><span className="metric-link">{i===3?'ไปจัดคาบที่เหลือ':'ดูรายละเอียด'} <Glyph name="arrow" size={15}/></span></button>)}</div>
  <div className="dashboard-columns"><section className="surface progress-panel"><div className="section-title"><h2>ความคืบหน้าตารางสอน</h2><span className="tag">ฉบับร่าง</span></div><div className="progress-layout"><div className="progress-ring" style={{'--progress':p.percent+'%'}}><div><strong>{p.percent}<small>%</small></strong><span>ลงคาบแล้ว</span></div></div><div className="progress-copy"><h3>{p.remaining?`อีก ${p.remaining} คาบให้จัดต่อ`:p.total?'ลงคาบตามงานมอบหมายครบแล้ว':'เริ่มจากเตรียมข้อมูล'}</h3><p>ลงตาราง {p.placed} จาก {p.total} คาบที่มอบหมาย</p><p className="muted">ตรวจคาบชนและเงื่อนไขอีกครั้งก่อนใช้งาน</p><button className="text-button" onClick={()=>setPage('scheduler')}>จัดตารางต่อ <Glyph name="arrow" size={16}/></button></div></div><div className="workflow">{steps.map(([title,desc,page,done],i)=><button key={title} onClick={()=>setPage(page)}><span className={'step-number '+(done?'done':'')}>{done?'✓':String(i+1).padStart(2,'0')}</span><strong>{title}</strong><small>{desc}</small></button>)}</div></section>
  <section className="surface attention-panel"><div className="section-title"><h2>งานที่ต้องตรวจสอบ</h2><Glyph name="alert"/></div><button className="attention-item" onClick={()=>setPage('scheduler')}><span className="attention-icon amber"><Glyph name="clock"/></span><span><strong>คาบที่ยังไม่ได้จัด</strong><small>จัดเพิ่มให้ครบตามงานมอบหมาย</small></span><b>{p.remaining}</b></button><button className="attention-item" onClick={()=>setPage('assignments')}><span className="attention-icon blue"><Glyph name="users"/></span><span><strong>ครูที่ยังไม่มีงานมอบหมาย</strong><small>ตรวจรายวิชาและห้องที่รับผิดชอบ</small></span><b>{unassigned.length}</b></button><button className="attention-item" onClick={()=>setPage('teachers')}><span className="attention-icon red"><Glyph name="alert"/></span><span><strong>ครูที่ลงคาบเกินภาระ</strong><small>ตรวจจำนวนคาบต่อสัปดาห์</small></span><b>{overloaded.length}</b></button><button className="secondary full" onClick={()=>setPage('reports')}>เปิดรายงานตรวจสอบ <Glyph name="arrow" size={16}/></button></section></div>
  <div className="dashboard-columns lower"><section className="surface"><div className="section-title"><div><h2>ภาพรวมรายวัน</h2><p>จำนวนรายการสอนที่ลงในตาราง</p></div><button className="text-button" onClick={()=>setPage('scheduler')}>ดูตาราง <Glyph name="arrow" size={16}/></button></div><div className="day-bars">{DAYS.map((day,i)=>{const n=Object.entries(S.schedule).filter(([k])=>k.split('_')[1]===day).reduce((n,[,v])=>n+v.length,0);const max=Math.max(1,...DAYS.map(d=>Object.entries(S.schedule).filter(([k])=>k.split('_')[1]===d).reduce((n,[,v])=>n+v.length,0)));return <div className="day-bar" key={day}><b>{n}</b><div className="bar-track"><div style={{height:Math.max(3,n/max*100)+'%',background:i===0?'#9c2638':'#e5bdc4'}}/></div><span>{day}</span></div>})}</div></section><section className="surface"><div className="section-title"><div><h2>ครูตามกลุ่มสาระ</h2><p>{S.depts.length} กลุ่มสาระ · {S.teachers.length} คน</p></div><button className="text-button" onClick={()=>setPage('teachers')}>ดูทั้งหมด <Glyph name="arrow" size={16}/></button></div><div className="department-list">{S.depts.slice(0,6).map((d,i)=><button key={d.id} onClick={()=>setPage('teachers')}><i style={{background:departmentTone(d).ink}}/><span>{d.name}</span><strong>{S.teachers.filter(t=>t.departmentId===d.id).length}</strong><small>คน</small></button>)}</div></section></div>
 </div>;
}

import './cards.css';

import './management.css';

import './scheduling.css';

import './finish.css';
