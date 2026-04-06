import { useState, useCallback, useEffect, useRef, useMemo } from "react";
import * as XLSX from 'xlsx';

const DAYS = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์"];
const PERIODS = [
  { id: 1, time: "08.30-09.20" }, { id: 2, time: "09.20-10.10" },
  { id: 3, time: "10.25-11.15" }, { id: 4, time: "11.15-12.05" },
  { id: 5, time: "13.00-13.50" }, { id: 6, time: "14.00-14.50" },
  { id: 7, time: "14.50-15.40" },
];
const DC = [
  { bg:"#DC2626",lt:"#FEE2E2",tx:"#991B1B",bd:"#FECACA" },
  { bg:"#2563EB",lt:"#DBEAFE",tx:"#1E40AF",bd:"#BFDBFE" },
  { bg:"#059669",lt:"#D1FAE5",tx:"#065F46",bd:"#A7F3D0" },
  { bg:"#D97706",lt:"#FEF3C7",tx:"#92400E",bd:"#FDE68A" },
  { bg:"#7C3AED",lt:"#EDE9FE",tx:"#5B21B6",bd:"#DDD6FE" },
  { bg:"#DB2777",lt:"#FCE7F3",tx:"#9D174D",bd:"#FBCFE8" },
  { bg:"#0891B2",lt:"#CFFAFE",tx:"#155E75",bd:"#A5F3FC" },
  { bg:"#65A30D",lt:"#ECFCCB",tx:"#3F6212",bd:"#D9F99D" },
  { bg:"#EA580C",lt:"#FFEDD5",tx:"#9A3412",bd:"#FED7AA" },
  { bg:"#4F46E5",lt:"#E0E7FF",tx:"#3730A3",bd:"#C7D2FE" },
];
const SROLES = [
  { id:"academic",name:"ฝ่ายวิชาการ",blocked:[{day:"พฤหัสบดี",periods:[5,6,7]}] },
  { id:"discipline",name:"ฝ่ายพัฒนาวินัย",blocked:[{day:"ศุกร์",periods:[5,6,7]}] },
];
const gid = () => Math.random().toString(36).substr(2,9);

// ===== CONFIG — ใส่ URL ของ GAS Web App ที่นี่ =====
const GAS_URL = "https://script.google.com/macros/s/AKfycbwWym1QWA-wumRvYKpRexd44eR3FrWw6fwjXA2-shsEjGOqNq5UXLThQvpiPGICs83ZKQ/exec";
// ====================================================

// localStorage helpers (ใช้เป็น cache offline)
const saveLS = (key, data) => { try { localStorage.setItem(`dara_${key}`, JSON.stringify(data)); } catch(e) {} };
const loadLS = (key, fb) => { try { const d = localStorage.getItem(`dara_${key}`); return d ? JSON.parse(d) : fb; } catch(e) { return fb; } };

// GAS helpers
const gasGet = async () => {
  const res = await fetch(GAS_URL);
  const json = await res.json();
  return json.ok ? json.data : null;
};
const gasPost = async (data) => {
  // GAS ไม่รับ CORS preflight → ใช้ no-cors + mode fetch trick
  await fetch(GAS_URL, {
    method: "POST",
    mode: "no-cors",
    headers: { "Content-Type": "text/plain" },
    body: JSON.stringify({ action: "save", data }),
  });
};

// Excel Export helper (SheetJS)
const exportExcel = (headers, rows, filename, sheetName = "Sheet1") => {
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols'] = headers.map((h, i) => ({ wch: Math.max(String(h).length * 2, ...rows.map(r => String(r[i] || "").length * 1.5), 14) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
};

// Excel Multi-sheet Export
const exportExcelMulti = (sheets, filename) => {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, headers, rows }) => {
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = headers.map(() => ({ wch: 25 }));
    XLSX.utils.book_append_sheet(wb, ws, name.substring(0, 31));
  });
  XLSX.writeFile(wb, filename);
};

// Excel Import helper
const readExcelFile = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array" });
      resolve(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" }));
    } catch (err) { reject(err); }
  };
  reader.onerror = reject;
  reader.readAsArrayBuffer(file);
});

// CSV Export (fallback)
const exportCSV = (headers, rows, filename) => {
  const bom = "\uFEFF";
  const csv = bom + [headers.join(","), ...rows.map(r => r.map(c => `"${String(c||"").replace(/"/g,'""')}"`).join(","))].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
};

// CSV Import (fallback)
const parseCSV = (text) => {
  const lines = text.split("\n").filter(l => l.trim());
  if (lines.length < 2) return [];
  const headers = lines[0].split(",").map(h => h.replace(/"/g,"").trim());
  return lines.slice(1).map(line => {
    const vals = line.match(/(".*?"|[^,]*)/g) || [];
    const obj = {};
    headers.forEach((h, i) => { obj[h] = (vals[i] || "").replace(/^"|"$/g, "").trim(); });
    return obj;
  });
};

const Icon = ({ name, size=18 }) => {
  const paths = {
    plus:"M12 5v14M5 12h14", trash:"M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2",
    lock:"M3 11h18v11H3zM7 11V7a5 5 0 0110 0v4", unlock:"M3 11h18v11H3zM7 11V7a5 5 0 019.9-1",
    users:"M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2M9 11a4 4 0 100-8 4 4 0 000 8z",
    check:"M20 6L9 17l-5-5", alert:"M12 2a10 10 0 100 20 10 10 0 000-20zM12 8v4M12 16h.01",
    download:"M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3",
    search:"M11 3a8 8 0 100 16 8 8 0 000-16zM21 21l-4.35-4.35",
    grid:"M3 3h7v7H3zM14 3h7v7h-7zM3 14h7v7H3zM14 14h7v7h-7z",
    upload:"M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12",
    x:"M18 6L6 18M6 6l12 12", menu:"M3 12h18M3 6h18M3 18h18",
    book:"M4 19.5A2.5 2.5 0 016.5 17H20M6.5 2H20v20H6.5A2.5 2.5 0 014 19.5v-15A2.5 2.5 0 016.5 2z",
    clock:"M12 2a10 10 0 100 20 10 10 0 000-20zM12 6v6l4 2",
    home:"M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z",
    edit:"M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7M18.5 2.5a2.12 2.12 0 013 3L12 15l-4 1 1-4z",
    file:"M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8zM14 2v6h6",
    layers:"M12 2L2 7l10 5 10-5zM2 17l10 5 10-5M2 12l10 5 10-5",
  };
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d={paths[name]||""}/></svg>;
};

const Modal = ({ open, onClose, title, children, wide }) => {
  if(!open) return null;
  return <div style={{position:"fixed",inset:0,zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.5)"}} onClick={onClose}>
    <div onClick={e=>e.stopPropagation()} style={{background:"#fff",borderRadius:16,boxShadow:"0 25px 50px rgba(0,0,0,0.25)",width:wide?"92%":"min(580px,92%)",maxHeight:"88vh",display:"flex",flexDirection:"column",overflow:"hidden"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"18px 24px",borderBottom:"1px solid #E5E7EB"}}>
        <h3 style={{margin:0,fontSize:17,fontWeight:700}}>{title}</h3>
        <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",color:"#9CA3AF",padding:4}}><Icon name="x"/></button>
      </div>
      <div style={{padding:24,overflowY:"auto",flex:1}}>{children}</div>
    </div>
  </div>;
};

const IS={width:"100%",padding:"10px 14px",border:"1.5px solid #D1D5DB",borderRadius:10,fontSize:14,outline:"none",fontFamily:"inherit",boxSizing:"border-box"};
const BS=(c="#DC2626")=>({padding:"10px 20px",background:c,color:"#fff",border:"none",borderRadius:10,fontSize:14,fontWeight:600,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6,fontFamily:"inherit"});
const BO=(c="#DC2626")=>({padding:"10px 20px",background:"#fff",color:c,border:`2px solid ${c}`,borderRadius:10,fontSize:14,fontWeight:600,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6,fontFamily:"inherit"});
const LS={display:"block",fontSize:13,fontWeight:600,color:"#374151",marginBottom:6};

const Toast=({message,type="success",onClose})=>{useEffect(()=>{const t=setTimeout(onClose,3000);return()=>clearTimeout(t)},[onClose]);return<div style={{position:"fixed",top:24,right:24,zIndex:9999,background:type==="error"?"#DC2626":type==="warning"?"#D97706":"#059669",color:"#fff",padding:"14px 24px",borderRadius:12,fontSize:14,fontWeight:600,boxShadow:"0 10px 30px rgba(0,0,0,0.2)",display:"flex",alignItems:"center",gap:8,animation:"slideIn 0.3s ease"}}><Icon name={type==="error"?"alert":"check"} size={16}/>{message}</div>};

export default function App() {
  const [page,setPage]=useState("dashboard");
  const [side,setSide]=useState(true);
  const [toast,setToast]=useState(null);
  const [syncing,setSyncing]=useState(false);
  const [gasReady,setGasReady]=useState(false);

  const [levels,setLevels]=useState(()=>loadLS("levels",[{id:gid(),name:"ม.4"},{id:gid(),name:"ม.5"},{id:gid(),name:"ม.6"}]));
  const [plans,setPlans]=useState(()=>loadLS("plans",[]));
  const [depts,setDepts]=useState(()=>loadLS("depts",[]));
  const [teachers,setTeachers]=useState(()=>loadLS("teachers",[]));
  const [subjects,setSubjects]=useState(()=>loadLS("subjects",[]));
  const [rooms,setRooms]=useState(()=>loadLS("rooms",[]));
  const [assigns,setAssigns]=useState(()=>loadLS("assigns",[]));
  const [meetings,setMeetings]=useState(()=>loadLS("meetings",[]));
  const [schedule,setSchedule]=useState(()=>loadLS("schedule",{}));
  const [locks,setLocks]=useState(()=>loadLS("locks",{}));

  const [academicYear,setAcademicYear]=useState(()=>loadLS("academicYear",{year:"2568",semester:"1"}));
  const [schoolHeader,setSchoolHeader]=useState(()=>loadLS("schoolHeader",{name:"โรงเรียนดาราวิทยาลัย",logo:""}));

  useEffect(()=>saveLS("academicYear",academicYear),[academicYear]);
  useEffect(()=>saveLS("schoolHeader",schoolHeader),[schoolHeader]);
  const stateRef=useRef({});
  useEffect(()=>{stateRef.current={levels,plans,depts,teachers,subjects,rooms,assigns,meetings,schedule,locks}},[levels,plans,depts,teachers,subjects,rooms,assigns,meetings,schedule,locks]);

  const saveTimer=useRef(null);
  const syncToGas=useCallback(()=>{
    if(!GAS_URL||GAS_URL.includes("YOUR_DEPLOYMENT_ID"))return;
    clearTimeout(saveTimer.current);
    saveTimer.current=setTimeout(()=>{
      setSyncing(true);
      gasPost(stateRef.current).catch(()=>{}).finally(()=>setSyncing(false));
    },1500);
  },[]);

  // โหลดจาก GAS ตอนเริ่ม
  useEffect(()=>{
    if(!GAS_URL||GAS_URL.includes("YOUR_DEPLOYMENT_ID"))return;
    setSyncing(true);
    gasGet().then(d=>{
      if(d){
        if(d.levels)   setLevels(d.levels);
        if(d.plans)    setPlans(d.plans);
        if(d.depts)    setDepts(d.depts);
        if(d.teachers) setTeachers(d.teachers);
        if(d.subjects) setSubjects(d.subjects);
        if(d.rooms)    setRooms(d.rooms);
        if(d.assigns)  setAssigns(d.assigns);
        if(d.meetings) setMeetings(d.meetings);
        if(d.schedule) setSchedule(d.schedule);
        if(d.locks)    setLocks(d.locks);
        setGasReady(true);
      } else { setGasReady(true); }
    }).catch(()=>{ setGasReady(true); }).finally(()=>setSyncing(false));
  },[]);

  // Auto-save ไป localStorage + GAS เมื่อข้อมูลเปลี่ยน
  useEffect(()=>{ saveLS("levels",levels);   if(gasReady) syncToGas(); },[levels,gasReady]);
  useEffect(()=>{ saveLS("plans",plans);     if(gasReady) syncToGas(); },[plans,gasReady]);
  useEffect(()=>{ saveLS("depts",depts);     if(gasReady) syncToGas(); },[depts,gasReady]);
  useEffect(()=>{ saveLS("teachers",teachers); if(gasReady) syncToGas(); },[teachers,gasReady]);
  useEffect(()=>{ saveLS("subjects",subjects); if(gasReady) syncToGas(); },[subjects,gasReady]);
  useEffect(()=>{ saveLS("rooms",rooms);     if(gasReady) syncToGas(); },[rooms,gasReady]);
  useEffect(()=>{ saveLS("assigns",assigns); if(gasReady) syncToGas(); },[assigns,gasReady]);
  useEffect(()=>{ saveLS("meetings",meetings); if(gasReady) syncToGas(); },[meetings,gasReady]);
  useEffect(()=>{ saveLS("schedule",schedule); if(gasReady) syncToGas(); },[schedule,gasReady]);
  useEffect(()=>{ saveLS("locks",locks);     if(gasReady) syncToGas(); },[locks,gasReady]);

  const st=(m,t="success")=>setToast({message:m,type:t});
  const gc=did=>{const i=depts.findIndex(d=>d.id===did);return DC[i%DC.length]||DC[0]};

  const nav=[
    {id:"dashboard",icon:"home",label:"แดชบอร์ด"},
    {id:"levels",icon:"grid",label:"ระดับชั้น / ห้องเรียน"},
    {id:"plans",icon:"layers",label:"แผนการเรียน"},
    {id:"departments",icon:"users",label:"กลุ่มสาระ"},
    {id:"teachers",icon:"users",label:"จัดการครู"},
    {id:"subjects",icon:"book",label:"จัดการวิชา"},
    {id:"assignments",icon:"edit",label:"มอบหมายงานครู"},
    {id:"meetings",icon:"clock",label:"คาบล็อค / ประชุม"},
    {id:"scheduler",icon:"grid",label:"จัดตารางสอน"},
    {id:"reports",icon:"download",label:"รายงาน / Export"},
    {id:"settings",icon:"file",label:"ตั้งค่า / ปีการศึกษา"},
  ];
  const S={levels,plans,depts,teachers,subjects,rooms,assigns,meetings,schedule,locks};
  const U={setLevels,setPlans,setDepts,setTeachers,setSubjects,setRooms,setAssigns,setMeetings,setSchedule,setLocks};

  return <div style={{display:"flex",height:"100vh",fontFamily:"'Sarabun','Noto Sans Thai',sans-serif",background:"#F3F4F6",overflow:"hidden"}}>
    <style>{`@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700;800&display=swap');*{box-sizing:border-box;margin:0;padding:0}::-webkit-scrollbar{width:6px}::-webkit-scrollbar-thumb{background:#CBD5E1;border-radius:3px}@keyframes slideIn{from{transform:translateX(100px);opacity:0}to{transform:translateX(0);opacity:1}}@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}.ni:hover{background:rgba(255,255,255,0.15)!important}.ni.a{background:rgba(255,255,255,0.2)!important}input:focus,select:focus{border-color:#DC2626!important;box-shadow:0 0 0 3px rgba(220,38,38,0.1)!important}.drag-card{cursor:grab;user-select:none}.drag-card:active{cursor:grabbing}.dz{transition:background 0.2s}.dz.over{background:#FEE2E2!important;outline:2px dashed #DC2626}button:hover{opacity:0.85}select{appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%236B7280' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 12px center;padding-right:36px!important}`}</style>

    <div style={{width:side?260:0,background:"linear-gradient(180deg,#991B1B,#7F1D1D)",transition:"width 0.3s",overflow:"hidden",flexShrink:0,display:"flex",flexDirection:"column"}}>
      <div style={{padding:"24px 20px",borderBottom:"1px solid rgba(255,255,255,0.1)"}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          {schoolHeader.logo
            ?<img src={schoolHeader.logo} alt="logo" style={{width:42,height:42,borderRadius:12,objectFit:"cover",flexShrink:0}}/>
            :<div style={{width:42,height:42,borderRadius:12,background:"rgba(255,255,255,0.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,fontWeight:800,color:"#fff"}}>ดว</div>
          }
          <div><div style={{color:"#fff",fontSize:15,fontWeight:700}}>{schoolHeader.name||"ดาราวิทยาลัย"}</div><div style={{color:"rgba(255,255,255,0.6)",fontSize:11}}>ระบบจัดตารางสอน v3</div></div>
        </div>
      </div>
      <nav style={{flex:1,padding:"12px 10px",overflowY:"auto"}}>
        {nav.map(n=><div key={n.id} className={`ni ${page===n.id?"a":""}`} onClick={()=>setPage(n.id)} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 14px",borderRadius:10,cursor:"pointer",color:page===n.id?"#fff":"rgba(255,255,255,0.7)",fontSize:14,fontWeight:page===n.id?700:400,marginBottom:2}}><Icon name={n.icon} size={18}/>{n.label}</div>)}
      </nav>
      <div style={{padding:"16px 20px",borderTop:"1px solid rgba(255,255,255,0.1)"}}>
        <div style={{color:"rgba(255,255,255,0.4)",fontSize:11}}>ผู้พัฒนา</div>
        <div style={{color:"rgba(255,255,255,0.7)",fontSize:12,fontWeight:600,marginTop:2}}>พนิต เกิดมงคล</div>
      </div>
    </div>

    <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
      <header style={{height:60,background:"#fff",borderBottom:"1px solid #E5E7EB",display:"flex",alignItems:"center",padding:"0 24px",gap:16,flexShrink:0}}>
        <button onClick={()=>setSide(!side)} style={{background:"none",border:"none",cursor:"pointer",color:"#6B7280",padding:4}}><Icon name="menu" size={22}/></button>
        <h2 style={{fontSize:18,fontWeight:700}}>{nav.find(n=>n.id===page)?.label}</h2>
        <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:8}}>
          {GAS_URL&&!GAS_URL.includes("YOUR_DEPLOYMENT_ID")
            ?syncing
              ?<span style={{fontSize:12,color:"#D97706",display:"flex",alignItems:"center",gap:4}}>⏳ กำลัง sync...</span>
              :gasReady
                ?<span style={{fontSize:12,color:"#059669",display:"flex",alignItems:"center",gap:4}}>☁️ sync แล้ว</span>
                :null
            :<span style={{fontSize:12,color:"#9CA3AF",display:"flex",alignItems:"center",gap:4}}>💾 local only</span>
          }
        </div>
      </header>
      <main style={{flex:1,overflow:"auto",padding:24}}>
        {page==="dashboard"&&<Dash S={S} setPage={setPage}/>}
        {page==="levels"&&<Levels S={S} U={U} st={st}/>}
        {page==="plans"&&<Plans S={S} U={U} st={st}/>}
        {page==="departments"&&<Depts S={S} U={U} st={st} gc={gc}/>}
        {page==="teachers"&&<Teachers S={S} U={U} st={st} gc={gc}/>}
        {page==="subjects"&&<Subjects S={S} U={U} st={st} gc={gc}/>}
        {page==="assignments"&&<Assigns S={S} U={U} st={st} gc={gc}/>}
        {page==="meetings"&&<Meetings S={S} U={U} st={st} gc={gc}/>}
        {page==="scheduler"&&<Scheduler S={S} U={U} st={st} gc={gc}/>}
        {page==="reports"&&<Reports S={S} st={st} gc={gc} ay={academicYear} sh={schoolHeader}/>}
        {page==="settings"&&<Settings S={S} U={U} st={st} ay={academicYear} setAY={setAcademicYear} sh={schoolHeader} setSH={setSchoolHeader}/>}
      </main>
    </div>
    {toast&&<Toast {...toast} onClose={()=>setToast(null)}/>}
  </div>;
}

/* ===== DASHBOARD ===== */
function Dash({S,setPage}){
  const stats=[{l:"ระดับชั้น",v:S.levels.length,c:"#DC2626"},{l:"แผนการเรียน",v:S.plans.length,c:"#7C3AED"},{l:"กลุ่มสาระ",v:S.depts.length,c:"#2563EB"},{l:"ครู",v:S.teachers.length,c:"#059669"},{l:"วิชา",v:S.subjects.length,c:"#D97706"},{l:"ห้อง",v:S.rooms.length,c:"#DB2777"}];
  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:16,marginBottom:32}}>
      {stats.map((s,i)=><div key={i} style={{background:"#fff",borderRadius:14,padding:20,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}><div style={{fontSize:28,fontWeight:800}}>{s.v}</div><div style={{fontSize:13,color:"#6B7280",marginTop:2}}>{s.l}</div><div style={{height:4,background:s.c,borderRadius:2,marginTop:12,width:"40%"}}/></div>)}
    </div>
    <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
      <h3 style={{fontSize:16,fontWeight:700,marginBottom:16}}>ขั้นตอนการใช้งาน</h3>
      {[{s:1,t:"สร้างระดับชั้นและห้องเรียน",p:"levels"},{s:2,t:"สร้างแผนการเรียน (ใช้ร่วมข้ามระดับได้)",p:"plans"},{s:3,t:"สร้างกลุ่มสาระการเรียนรู้",p:"departments"},{s:4,t:"เพิ่มครู + กำหนดคาบที่ได้รับ",p:"teachers"},{s:5,t:"สร้างวิชา + ระบุระดับชั้น",p:"subjects"},{s:6,t:"มอบหมายวิชาและห้องให้ครู",p:"assignments"},{s:7,t:"ตั้งคาบล็อค/ประชุม",p:"meetings"},{s:8,t:"จัดตารางสอน (Drag & Drop)",p:"scheduler"},{s:9,t:"ตรวจสอบและ Export CSV",p:"reports"}].map(s=><div key={s.s} onClick={()=>setPage(s.p)} style={{display:"flex",alignItems:"center",gap:14,padding:"12px 16px",borderRadius:10,cursor:"pointer",background:"#F9FAFB",marginBottom:6}} onMouseEnter={e=>e.currentTarget.style.background="#FEE2E2"} onMouseLeave={e=>e.currentTarget.style.background="#F9FAFB"}><div style={{width:30,height:30,borderRadius:"50%",background:"#DC2626",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,flexShrink:0}}>{s.s}</div><span style={{fontSize:14}}>{s.t}</span></div>)}
    </div>
  </div>;
}

/* ===== LEVELS & ROOMS (+ import/export) ===== */
function Levels({S,U,st}){
  const [rm,setRm]=useState(false);
  const [rf,setRf]=useState({levelId:"",planId:"",name:""});
  const fileRefLv=useRef(null);
  const fileRefRm=useRef(null);

  const addLv=()=>{const n=prompt("ชื่อระดับชั้น:");if(n){U.setLevels(p=>[...p,{id:gid(),name:n}]);st("เพิ่มสำเร็จ")}};
  const editLv=(lv)=>{const n=prompt("แก้ไขชื่อระดับชั้น:",lv.name);if(n){U.setLevels(p=>p.map(l=>l.id===lv.id?{...l,name:n}:l));st("แก้ไขสำเร็จ")}};

  const importLevels=async(e)=>{const f=e.target.files?.[0];if(!f)return;
    const rows=f.name.endsWith('.csv')?parseCSV(await f.text()):await readExcelFile(f);
    const newL=rows.map(r=>({id:gid(),name:String(r["ชื่อระดับชั้น"]||"").trim()})).filter(x=>x.name);
    U.setLevels(p=>[...p,...newL]);st(`นำเข้า ${newL.length} ระดับชั้น`);e.target.value=""};
  const exportLevels=()=>{exportExcel(["ชื่อระดับชั้น"],S.levels.map(l=>[l.name]),"ระดับชั้น.xlsx","ระดับชั้น");st("Export สำเร็จ")};
  const templateLevels=()=>{exportExcel(["ชื่อระดับชั้น"],[["ม.4"],["ม.5"],["ม.6"]],"Template_ระดับชั้น.xlsx","Template");st("ดาวน์โหลด Template")};

  const importRooms=async(e)=>{const f=e.target.files?.[0];if(!f)return;
    const rows=f.name.endsWith('.csv')?parseCSV(await f.text()):await readExcelFile(f);
    const newR=rows.map(r=>{const lv=S.levels.find(l=>l.name===String(r["ระดับชั้น"]||"").trim());const pl=S.plans.find(p=>p.name===String(r["แผนการเรียน"]||"").trim());
      return{id:gid(),name:String(r["ชื่อห้อง"]||"").trim(),levelId:lv?.id||"",planId:pl?.id||""}}).filter(x=>x.name&&x.levelId);
    U.setRooms(p=>[...p,...newR]);st(`นำเข้า ${newR.length} ห้อง`);e.target.value=""};
  const exportRooms=()=>{exportExcel(["ชื่อห้อง","ระดับชั้น","แผนการเรียน"],S.rooms.map(r=>[r.name,S.levels.find(l=>l.id===r.levelId)?.name||"",S.plans.find(p=>p.id===r.planId)?.name||""]),"ห้องเรียน.xlsx","ห้อง");st("Export สำเร็จ")};
  const templateRooms=()=>{exportExcel(["ชื่อห้อง","ระดับชั้น","แผนการเรียน"],[["ม.4/1","ม.4","วิทย์-คณิต"],["ม.4/2","ม.4","ศิลป์-ภาษา"]],"Template_ห้องเรียน.xlsx","Template");st("ดาวน์โหลด Template")};

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap"}}>
      <button onClick={addLv} style={BS()}><Icon name="plus" size={16}/>เพิ่มระดับชั้น</button>
      <button onClick={()=>fileRefLv.current?.click()} style={BS("#2563EB")}><Icon name="upload" size={16}/>Import ระดับชั้น</button>
      <button onClick={templateLevels} style={BO("#2563EB")}><Icon name="file" size={16}/>Template</button>
      <button onClick={exportLevels} style={BO("#059669")}><Icon name="download" size={16}/>Export</button>
      <input ref={fileRefLv} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={importLevels}/>
    </div>
    <div style={{display:"flex",gap:10,marginBottom:24,flexWrap:"wrap"}}>
      <button onClick={()=>setRm(true)} style={BS("#7C3AED")}><Icon name="plus" size={16}/>เพิ่มห้องเรียน</button>
      <button onClick={()=>fileRefRm.current?.click()} style={BS("#0891B2")}><Icon name="upload" size={16}/>Import ห้อง</button>
      <button onClick={templateRooms} style={BO("#0891B2")}><Icon name="file" size={16}/>Template ห้อง</button>
      <button onClick={exportRooms} style={BO("#059669")}><Icon name="download" size={16}/>Export ห้อง</button>
      <input ref={fileRefRm} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={importRooms}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(320px,1fr))",gap:20}}>
      {S.levels.map(lv=><div key={lv.id} style={{background:"#fff",borderRadius:14,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",overflow:"hidden"}}>
        <div style={{background:"linear-gradient(135deg,#991B1B,#DC2626)",padding:"16px 20px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <h3 style={{color:"#fff",fontSize:18,fontWeight:700}}>{lv.name}</h3>
          <div style={{display:"flex",gap:6}}>
            <button onClick={()=>editLv(lv)} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:6,padding:6,color:"#fff",cursor:"pointer"}}><Icon name="edit" size={14}/></button>
            <button onClick={()=>{U.setLevels(p=>p.filter(l=>l.id!==lv.id));st("ลบแล้ว","warning")}} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:6,padding:6,color:"#fff",cursor:"pointer"}}><Icon name="trash" size={14}/></button>
          </div>
        </div>
        <div style={{padding:16}}>
          <div style={{fontSize:12,fontWeight:600,color:"#9CA3AF",marginBottom:6}}>ห้องเรียน:</div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {S.rooms.filter(r=>r.levelId===lv.id).map(rm=>{const plan=S.plans.find(p=>p.id===rm.planId);return<span key={rm.id} style={{background:"#DBEAFE",color:"#1E40AF",fontSize:12,padding:"4px 12px",borderRadius:20,fontWeight:600,display:"inline-flex",alignItems:"center",gap:4}}>
              {rm.name}{plan?" ("+plan.name+(plan.subPlans?.length?" \u2014 "+plan.subPlans.join(", "):"")+")":""}              <button onClick={()=>{const n=prompt("แก้ไขชื่อห้อง:",rm.name);if(n){U.setRooms(p=>p.map(r=>r.id===rm.id?{...r,name:n}:r));st("แก้ไขสำเร็จ")}}} style={{background:"none",border:"none",cursor:"pointer",color:"#1E40AF",padding:0}}><Icon name="edit" size={10}/></button>
              <button onClick={()=>U.setRooms(p=>p.filter(r=>r.id!==rm.id))} style={{background:"none",border:"none",cursor:"pointer",color:"#1E40AF",padding:0}}><Icon name="x" size={10}/></button>
            </span>})}
            {!S.rooms.filter(r=>r.levelId===lv.id).length&&<span style={{fontSize:12,color:"#9CA3AF"}}>ยังไม่มี</span>}
          </div>
        </div>
      </div>)}
    </div>
    <Modal open={rm} onClose={()=>setRm(false)} title="เพิ่มห้องเรียน">
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <div><label style={LS}>ระดับชั้น</label><select style={IS} value={rf.levelId} onChange={e=>setRf(p=>({...p,levelId:e.target.value}))}><option value="">--</option>{S.levels.map(l=><option key={l.id} value={l.id}>{l.name}</option>)}</select></div>
        <div><label style={LS}>แผนการเรียน (ถ้ามี)</label><select style={IS} value={rf.planId} onChange={e=>setRf(p=>({...p,planId:e.target.value}))}><option value="">--</option>{S.plans.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></div>
        <div><label style={LS}>ชื่อห้อง</label><input style={IS} value={rf.name} onChange={e=>setRf(p=>({...p,name:e.target.value}))} placeholder="ม.4/1"/></div>
        <button onClick={()=>{if(!rf.name||!rf.levelId)return;U.setRooms(p=>[...p,{id:gid(),...rf}]);setRf({levelId:"",planId:"",name:""});setRm(false);st("เพิ่มสำเร็จ")}} style={BS()}>บันทึก</button>
      </div>
    </Modal>
  </div>;
}

/* ===== PLANS (+ import/export) ===== */
function Plans({S,U,st}){
  const [modal,setModal]=useState(false);
  const [form,setForm]=useState({name:"",subPlans:"",levelIds:[]});
  const [editId,setEditId]=useState(null);
  const fileRef=useRef(null);

  const save=()=>{
    if(!form.name){st("กรุณาใส่ชื่อ","error");return}
    const subs=form.subPlans?form.subPlans.split(",").map(s=>s.trim()).filter(Boolean):[];
    if(editId){U.setPlans(p=>p.map(x=>x.id===editId?{...x,name:form.name,subPlans:subs,levelIds:form.levelIds}:x));st("แก้ไขสำเร็จ")}
    else{U.setPlans(p=>[...p,{id:gid(),name:form.name,subPlans:subs,levelIds:form.levelIds}]);st("เพิ่มสำเร็จ")}
    setForm({name:"",subPlans:"",levelIds:[]});setModal(false);setEditId(null);
  };
  const openEdit=(plan)=>{setEditId(plan.id);setForm({name:plan.name,subPlans:(plan.subPlans||[]).join(", "),levelIds:plan.levelIds||[]});setModal(true)};
  const toggleLv=(lid)=>setForm(p=>({...p,levelIds:p.levelIds.includes(lid)?p.levelIds.filter(x=>x!==lid):[...p.levelIds,lid]}));

  const importPlans=async(e)=>{const f=e.target.files?.[0];if(!f)return;
    const rows=f.name.endsWith('.csv')?parseCSV(await f.text()):await readExcelFile(f);
    const newP=rows.map(r=>{const lvNames=String(r["ระดับชั้น"]||"").split(",").map(s=>s.trim()).filter(Boolean);const lvIds=lvNames.map(n=>S.levels.find(l=>l.name===n)?.id).filter(Boolean);
      return{id:gid(),name:String(r["ชื่อแผน"]||"").trim(),subPlans:String(r["สายรอง"]||"").split(",").map(s=>s.trim()).filter(Boolean),levelIds:lvIds}}).filter(x=>x.name);
    U.setPlans(p=>[...p,...newP]);st(`นำเข้า ${newP.length} แผน`);e.target.value=""};
  const exportPlans=()=>{exportExcel(["ชื่อแผน","สายรอง","ระดับชั้น"],S.plans.map(p=>[p.name,(p.subPlans||[]).join(","),(p.levelIds||[]).map(lid=>S.levels.find(l=>l.id===lid)?.name).filter(Boolean).join(",")]),"แผนการเรียน.xlsx","แผน");st("Export สำเร็จ")};
  const templatePlans=()=>{exportExcel(["ชื่อแผน","สายรอง","ระดับชั้น"],[["วิทย์-คณิต","วิทย์สุขภาพ,วิศวะ","ม.4,ม.5,ม.6"],["ศิลป์-ภาษา","","ม.4,ม.5"]],"Template_แผนการเรียน.xlsx","Template");st("ดาวน์โหลด Template")};

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"flex",gap:10,marginBottom:24,flexWrap:"wrap"}}>
      <button onClick={()=>{setEditId(null);setForm({name:"",subPlans:"",levelIds:[]});setModal(true)}} style={BS()}><Icon name="plus" size={16}/>เพิ่มแผนการเรียน</button>
      <button onClick={()=>fileRef.current?.click()} style={BS("#2563EB")}><Icon name="upload" size={16}/>Import Excel</button>
      <button onClick={templatePlans} style={BO("#2563EB")}><Icon name="file" size={16}/>Template</button>
      <button onClick={exportPlans} style={BO("#059669")}><Icon name="download" size={16}/>Export</button>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={importPlans}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:16}}>
      {S.plans.map(plan=><div key={plan.id} style={{background:"#fff",borderRadius:14,padding:20,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{display:"flex",justifyContent:"space-between"}}>
          <h4 style={{fontSize:16,fontWeight:700}}>{plan.name}</h4>
          <div style={{display:"flex",gap:6}}>
            <button onClick={()=>openEdit(plan)} style={{background:"none",border:"none",cursor:"pointer",color:"#2563EB"}}><Icon name="edit" size={16}/></button>
            <button onClick={()=>{U.setPlans(p=>p.filter(x=>x.id!==plan.id));st("ลบแล้ว","warning")}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444"}}><Icon name="trash" size={16}/></button>
          </div>
        </div>
        {plan.subPlans?.length>0&&<div style={{marginTop:8,display:"flex",gap:6,flexWrap:"wrap"}}>{plan.subPlans.map((sp,i)=><span key={i} style={{background:"#FEE2E2",color:"#991B1B",fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{sp}</span>)}</div>}
        <div style={{marginTop:8,display:"flex",gap:6,flexWrap:"wrap"}}>{(plan.levelIds||[]).map(lid=>{const lv=S.levels.find(l=>l.id===lid);return lv?<span key={lid} style={{background:"#DBEAFE",color:"#1E40AF",fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{lv.name}</span>:null})}</div>
      </div>)}
    </div>
    <Modal open={modal} onClose={()=>{setModal(false);setEditId(null)}} title={editId?"แก้ไขแผนการเรียน":"เพิ่มแผนการเรียน"}>
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <div><label style={LS}>ชื่อแผน</label><input style={IS} value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))} placeholder="วิทย์-คณิต"/></div>
        <div><label style={LS}>สายรอง (คอมม่า)</label><input style={IS} value={form.subPlans} onChange={e=>setForm(p=>({...p,subPlans:e.target.value}))} placeholder="วิทย์สุขภาพ, วิศวะ"/></div>
        <div><label style={LS}>ใช้กับระดับชั้น</label>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{S.levels.map(lv=><button key={lv.id} onClick={()=>toggleLv(lv.id)} style={{padding:"8px 16px",borderRadius:10,border:`2px solid ${form.levelIds.includes(lv.id)?"#DC2626":"#D1D5DB"}`,background:form.levelIds.includes(lv.id)?"#FEE2E2":"#fff",color:form.levelIds.includes(lv.id)?"#991B1B":"#374151",fontSize:13,fontWeight:600,cursor:"pointer"}}>{form.levelIds.includes(lv.id)?"✓ ":""}{lv.name}</button>)}</div>
        </div>
        <button onClick={save} style={BS()}>{editId?"บันทึกการแก้ไข":"เพิ่ม"}</button>
      </div>
    </Modal>
  </div>;
}

/* ===== DEPARTMENTS (+ import/export) ===== */
function Depts({S,U,st,gc}){
  const [name,setName]=useState("");
  const fileRef=useRef(null);

  const importDepts=async(e)=>{const f=e.target.files?.[0];if(!f)return;
    const rows=f.name.endsWith('.csv')?parseCSV(await f.text()):await readExcelFile(f);
    const newD=rows.map(r=>({id:gid(),name:String(r["ชื่อกลุ่มสาระ"]||"").trim()})).filter(x=>x.name);
    U.setDepts(p=>[...p,...newD]);st(`นำเข้า ${newD.length} กลุ่มสาระ`);e.target.value=""};
  const exportDepts=()=>{exportExcel(["ชื่อกลุ่มสาระ"],S.depts.map(d=>[d.name]),"กลุ่มสาระ.xlsx","กลุ่มสาระ");st("Export สำเร็จ")};
  const templateDepts=()=>{exportExcel(["ชื่อกลุ่มสาระ"],[["วิทยาศาสตร์และเทคโนโลยี"],["คณิตศาสตร์"],["ภาษาไทย"],["ภาษาต่างประเทศ"],["สังคมศึกษา"],["สุขศึกษาและพลศึกษา"],["ศิลปะ"],["การงานอาชีพ"]],"Template_กลุ่มสาระ.xlsx","Template");st("ดาวน์โหลด Template")};

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"flex",gap:10,marginBottom:24,flexWrap:"wrap",alignItems:"center"}}>
      <input style={{...IS,maxWidth:300}} value={name} onChange={e=>setName(e.target.value)} placeholder="ชื่อกลุ่มสาระ" onKeyDown={e=>{if(e.key==="Enter"&&name){U.setDepts(p=>[...p,{id:gid(),name}]);setName("");st("เพิ่มสำเร็จ")}}}/>
      <button onClick={()=>{if(!name)return;U.setDepts(p=>[...p,{id:gid(),name}]);setName("");st("เพิ่มสำเร็จ")}} style={{...BS(),flexShrink:0}}><Icon name="plus" size={16}/>เพิ่ม</button>
      <button onClick={()=>fileRef.current?.click()} style={BS("#2563EB")}><Icon name="upload" size={16}/>Import</button>
      <button onClick={templateDepts} style={BO("#2563EB")}><Icon name="file" size={16}/>Template</button>
      <button onClick={exportDepts} style={BO("#059669")}><Icon name="download" size={16}/>Export</button>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={importDepts}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:16}}>
      {S.depts.map(d=>{const c=gc(d.id);return<div key={d.id} style={{background:"#fff",borderRadius:14,overflow:"hidden",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{height:6,background:c.bg}}/><div style={{padding:20}}>
          <div style={{display:"flex",justifyContent:"space-between"}}><h4 style={{fontSize:16,fontWeight:700}}>{d.name}</h4>
            <div style={{display:"flex",gap:6}}>
              <button onClick={()=>{const n=prompt("แก้ไข:",d.name);if(n){U.setDepts(p=>p.map(x=>x.id===d.id?{...x,name:n}:x));st("แก้ไขสำเร็จ")}}} style={{background:"none",border:"none",cursor:"pointer",color:"#2563EB"}}><Icon name="edit" size={16}/></button>
              <button onClick={()=>{U.setDepts(p=>p.filter(x=>x.id!==d.id));st("ลบแล้ว","warning")}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444"}}><Icon name="trash" size={16}/></button>
            </div>
          </div>
          <div style={{display:"flex",gap:12,marginTop:8}}><span style={{fontSize:12,color:"#6B7280"}}>ครู {S.teachers.filter(t=>t.departmentId===d.id).length}</span><span style={{fontSize:12,color:"#6B7280"}}>วิชา {S.subjects.filter(s=>s.departmentId===d.id).length}</span></div>
        </div>
      </div>})}
    </div>
  </div>;
}

/* ===== TEACHERS (fix#2,#5: edit + คาบที่ได้รับ) ===== */
function Teachers({S,U,st,gc}){
  const [modal,setModal]=useState(false);
  const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({prefix:"",firstName:"",lastName:"",departmentId:"",specialRoles:[],totalPeriods:0});
  const [search,setSearch]=useState("");
  const fileRef=useRef(null);

  const save=()=>{
    if(!form.firstName||!form.departmentId){st("กรุณากรอกให้ครบ","error");return}
    if(editId){
      U.setTeachers(p=>p.map(t=>t.id===editId?{...t,...form}:t));st("แก้ไขสำเร็จ");
    } else {
      U.setTeachers(p=>[...p,{id:gid(),...form}]);st("เพิ่มครูสำเร็จ");
    }
    setForm({prefix:"",firstName:"",lastName:"",departmentId:"",specialRoles:[],totalPeriods:0});setModal(false);setEditId(null);
  };

  const openEdit=(t)=>{setEditId(t.id);setForm({prefix:t.prefix,firstName:t.firstName,lastName:t.lastName,departmentId:t.departmentId,specialRoles:t.specialRoles||[],totalPeriods:t.totalPeriods||0});setModal(true)};
  const toggleRole=(rid)=>setForm(p=>({...p,specialRoles:p.specialRoles.includes(rid)?p.specialRoles.filter(r=>r!==rid):[...p.specialRoles,rid]}));

  // Import Excel/CSV
  const handleFile=async(e)=>{const f=e.target.files?.[0];if(!f)return;
    let rows;
    if(f.name.endsWith('.csv')){const txt=await f.text();rows=parseCSV(txt)}
    else{rows=await readExcelFile(f)}
    if(!rows||!rows.length){st("ไม่พบข้อมูล","error");return}
    const newT=rows.map(r=>{const dept=S.depts.find(d=>d.name===String(r["กลุ่มสาระ"]||"").trim());const roles=[];const rs=String(r["หน้าที่พิเศษ"]||"");if(rs.includes("วิชาการ"))roles.push("academic");if(rs.includes("วินัย"))roles.push("discipline");
      return{id:gid(),prefix:String(r["คำนำหน้า"]||"").trim(),firstName:String(r["ชื่อ"]||"").trim(),lastName:String(r["นามสกุล"]||"").trim(),departmentId:dept?.id||"",specialRoles:roles,totalPeriods:parseInt(r["คาบที่ได้รับ"])||0}
    }).filter(t=>t.firstName);
    U.setTeachers(p=>[...p,...newT]);st(`นำเข้า ${newT.length} คน`);e.target.value="";
  };

  const exportT=()=>{exportExcel(["คำนำหน้า","ชื่อ","นามสกุล","กลุ่มสาระ","หน้าที่พิเศษ","คาบที่ได้รับ"],S.teachers.map(t=>[t.prefix,t.firstName,t.lastName,S.depts.find(d=>d.id===t.departmentId)?.name||"",(t.specialRoles||[]).map(r=>SROLES.find(sr=>sr.id===r)?.name).filter(Boolean).join("/")||"ครูทั่วไป",t.totalPeriods||0]),"รายชื่อครู_ดาราวิทยาลัย.xlsx","ครู");st("Export สำเร็จ")};

  const downloadTemplate=()=>{exportExcel(["คำนำหน้า","ชื่อ","นามสกุล","กลุ่มสาระ","หน้าที่พิเศษ","คาบที่ได้รับ"],[["นาย","สมชาย","ใจดี","วิทยาศาสตร์","ฝ่ายวิชาการ",18],["นางสาว","สมหญิง","รักเรียน","คณิตศาสตร์","ครูทั่วไป",20]],"Template_ครู.xlsx","Template");st("ดาวน์โหลด Template")};

  const usedPeriods=(tid)=>{let u=0;S.assigns.filter(a=>a.teacherId===tid).forEach(a=>{u+=a.totalPeriods||0});return u};

  const filtered=S.teachers.filter(t=>`${t.prefix}${t.firstName} ${t.lastName}`.includes(search)||S.depts.find(d=>d.id===t.departmentId)?.name?.includes(search));

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"flex",gap:10,marginBottom:20,flexWrap:"wrap",alignItems:"center"}}>
      <button onClick={()=>{setEditId(null);setForm({prefix:"",firstName:"",lastName:"",departmentId:"",specialRoles:[],totalPeriods:0});setModal(true)}} style={BS()}><Icon name="plus" size={16}/>เพิ่มครู</button>
      <button onClick={()=>fileRef.current?.click()} style={BS("#2563EB")}><Icon name="upload" size={16}/>Import Excel</button>
      <button onClick={downloadTemplate} style={BO("#2563EB")}><Icon name="file" size={16}/>Template</button>
      <button onClick={exportT} style={BO("#059669")}><Icon name="download" size={16}/>Export Excel</button>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleFile}/>
      <div style={{position:"relative",flex:"1 1 200px",maxWidth:350}}><input style={{...IS,paddingLeft:38}} value={search} onChange={e=>setSearch(e.target.value)} placeholder="ค้นหาครู..."/><div style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:"#9CA3AF"}}><Icon name="search" size={16}/></div></div>
    </div>

    <div style={{background:"#fff",borderRadius:14,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",overflow:"auto"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:14}}>
        <thead><tr style={{background:"#F9FAFB"}}>{["#","ชื่อ-สกุล","กลุ่มสาระ","คาบได้รับ","มอบหมาย","เหลือ","หน้าที่พิเศษ","จัดการ"].map(h=><th key={h} style={{padding:"12px 14px",textAlign:"left",fontWeight:600,color:"#6B7280",fontSize:12}}>{h}</th>)}</tr></thead>
        <tbody>{filtered.map((t,i)=>{const dept=S.depts.find(d=>d.id===t.departmentId);const c=dept?gc(dept.id):{bg:"#6B7280",lt:"#F3F4F6",tx:"#374151"};const used=usedPeriods(t.id);const rem=(t.totalPeriods||0)-used;return<tr key={t.id} style={{borderTop:"1px solid #F3F4F6"}}>
          <td style={{padding:"12px 14px",color:"#9CA3AF"}}>{i+1}</td>
          <td style={{padding:"12px 14px",fontWeight:600}}>{t.prefix}{t.firstName} {t.lastName}</td>
          <td style={{padding:"12px 14px"}}>{dept?<span style={{background:c.lt,color:c.tx,padding:"3px 12px",borderRadius:20,fontSize:12,fontWeight:600}}>{dept.name}</span>:<span style={{color:"#EF4444",fontSize:12}}>ไม่พบ</span>}</td>
          <td style={{padding:"12px 14px",fontWeight:700}}>{t.totalPeriods||0}</td>
          <td style={{padding:"12px 14px"}}>{used}</td>
          <td style={{padding:"12px 14px",fontWeight:700,color:rem>0?"#D97706":rem===0?"#059669":"#DC2626"}}>{rem}</td>
          <td style={{padding:"12px 14px"}}><div style={{display:"flex",gap:4,flexWrap:"wrap"}}>{t.specialRoles?.length?t.specialRoles.map(r=><span key={r} style={{background:"#FEF3C7",color:"#92400E",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>{SROLES.find(sr=>sr.id===r)?.name}</span>):<span style={{color:"#9CA3AF",fontSize:12}}>ครูทั่วไป</span>}</div></td>
          <td style={{padding:"12px 14px"}}><div style={{display:"flex",gap:6}}>
            <button onClick={()=>openEdit(t)} style={{background:"none",border:"none",cursor:"pointer",color:"#2563EB"}}><Icon name="edit" size={16}/></button>
            <button onClick={()=>{U.setTeachers(p=>p.filter(x=>x.id!==t.id));st("ลบแล้ว","warning")}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444"}}><Icon name="trash" size={16}/></button>
          </div></td>
        </tr>})}</tbody>
      </table>
      {!filtered.length&&<div style={{padding:40,textAlign:"center",color:"#9CA3AF"}}>ยังไม่มีข้อมูลครู</div>}
    </div>

    <Modal open={modal} onClose={()=>{setModal(false);setEditId(null)}} title={editId?"แก้ไขครู":"เพิ่มครู"}>
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <div style={{display:"grid",gridTemplateColumns:"100px 1fr 1fr",gap:12}}>
          <div><label style={LS}>คำนำหน้า</label><select style={IS} value={form.prefix} onChange={e=>setForm(p=>({...p,prefix:e.target.value}))}><option value="">--</option><option>นาย</option><option>นาง</option><option>นางสาว</option></select></div>
          <div><label style={LS}>ชื่อ</label><input style={IS} value={form.firstName} onChange={e=>setForm(p=>({...p,firstName:e.target.value}))}/></div>
          <div><label style={LS}>นามสกุล</label><input style={IS} value={form.lastName} onChange={e=>setForm(p=>({...p,lastName:e.target.value}))}/></div>
        </div>
        <div><label style={LS}>กลุ่มสาระ</label><select style={IS} value={form.departmentId} onChange={e=>setForm(p=>({...p,departmentId:e.target.value}))}><option value="">--</option>{S.depts.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}</select></div>
        <div><label style={LS}>คาบที่ได้รับ (ต่อสัปดาห์)</label><input type="number" min="0" style={IS} value={form.totalPeriods} onChange={e=>setForm(p=>({...p,totalPeriods:parseInt(e.target.value)||0}))}/></div>
        <div><label style={LS}>หน้าที่พิเศษ</label><div style={{display:"flex",gap:8}}>{SROLES.map(r=><button key={r.id} onClick={()=>toggleRole(r.id)} style={{padding:"8px 16px",borderRadius:10,border:`2px solid ${form.specialRoles.includes(r.id)?"#DC2626":"#D1D5DB"}`,background:form.specialRoles.includes(r.id)?"#FEE2E2":"#fff",fontSize:13,fontWeight:600,cursor:"pointer"}}>{form.specialRoles.includes(r.id)?"✓ ":""}{r.name}</button>)}</div></div>
        <button onClick={save} style={BS()}>{editId?"บันทึก":"เพิ่มครู"}</button>
      </div>
    </Modal>
  </div>;
}

/* ===== SUBJECTS (fix#3: + level) ===== */
function Subjects({S,U,st,gc}){
  const [modal,setModal]=useState(false);
  const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({code:"",name:"",credits:1,periodsPerWeek:1,departmentId:"",levelId:""});
  const fileRef=useRef(null);

  const save=()=>{if(!form.name||!form.departmentId||!form.levelId){st("กรอกให้ครบ","error");return}
    if(editId){U.setSubjects(p=>p.map(s=>s.id===editId?{...s,...form}:s));st("แก้ไขสำเร็จ")}
    else{U.setSubjects(p=>[...p,{id:gid(),...form}]);st("เพิ่มวิชาสำเร็จ")}
    setForm({code:"",name:"",credits:1,periodsPerWeek:1,departmentId:"",levelId:""});setModal(false);setEditId(null);
  };
  const openEdit=(s)=>{setEditId(s.id);setForm({code:s.code,name:s.name,credits:s.credits,periodsPerWeek:s.periodsPerWeek,departmentId:s.departmentId,levelId:s.levelId||""});setModal(true)};

  const handleFile=async(e)=>{const f=e.target.files?.[0];if(!f)return;
    let rows;
    if(f.name.endsWith('.csv')){const txt=await f.text();rows=parseCSV(txt)}
    else{rows=await readExcelFile(f)}
    if(!rows||!rows.length){st("ไม่พบข้อมูล","error");return}
    const ns=rows.map(r=>{const dept=S.depts.find(d=>d.name===String(r["กลุ่มสาระ"]||"").trim());const lv=S.levels.find(l=>l.name===String(r["ระดับชั้น"]||"").trim());
      return{id:gid(),code:String(r["รหัสวิชา"]||"").trim(),name:String(r["ชื่อวิชา"]||"").trim(),credits:parseFloat(r["หน่วยกิต"])||1,periodsPerWeek:parseInt(r["คาบ/สัปดาห์"])||1,departmentId:dept?.id||"",levelId:lv?.id||""}
    }).filter(s=>s.name);
    U.setSubjects(p=>[...p,...ns]);st(`นำเข้า ${ns.length} วิชา`);e.target.value=""};

  const exportS=()=>{exportExcel(["รหัสวิชา","ชื่อวิชา","หน่วยกิต","คาบ/สัปดาห์","กลุ่มสาระ","ระดับชั้น"],S.subjects.map(s=>[s.code,s.name,s.credits,s.periodsPerWeek,S.depts.find(d=>d.id===s.departmentId)?.name||"",S.levels.find(l=>l.id===s.levelId)?.name||""]),"รายวิชา_ดาราวิทยาลัย.xlsx","วิชา");st("Export สำเร็จ")};

  const downloadTemplate=()=>{exportExcel(["รหัสวิชา","ชื่อวิชา","หน่วยกิต","คาบ/สัปดาห์","กลุ่มสาระ","ระดับชั้น"],[["ว33201","ฟิสิกส์ 3",1.5,3,"วิทยาศาสตร์","ม.6"]],"Template_วิชา.xlsx","Template");st("ดาวน์โหลด Template")};

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"flex",gap:10,marginBottom:20,flexWrap:"wrap"}}>
      <button onClick={()=>{setEditId(null);setForm({code:"",name:"",credits:1,periodsPerWeek:1,departmentId:"",levelId:""});setModal(true)}} style={BS()}><Icon name="plus" size={16}/>เพิ่มวิชา</button>
      <button onClick={()=>fileRef.current?.click()} style={BS("#2563EB")}><Icon name="upload" size={16}/>Import Excel</button>
      <button onClick={downloadTemplate} style={BO("#2563EB")}><Icon name="file" size={16}/>Template</button>
      <button onClick={exportS} style={BO("#059669")}><Icon name="download" size={16}/>Export Excel</button>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleFile}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:16}}>
      {S.subjects.map(sub=>{const dept=S.depts.find(d=>d.id===sub.departmentId);const c=dept?gc(dept.id):{bg:"#6B7280",lt:"#F3F4F6",tx:"#374151"};const lv=S.levels.find(l=>l.id===sub.levelId);return<div key={sub.id} style={{background:"#fff",borderRadius:14,overflow:"hidden",boxShadow:"0 1px 3px rgba(0,0,0,0.06)",borderLeft:`4px solid ${c.bg}`}}>
        <div style={{padding:16}}>
          <div style={{display:"flex",justifyContent:"space-between"}}><div><div style={{fontSize:11,color:"#9CA3AF",fontWeight:600}}>{sub.code}</div><h4 style={{fontSize:15,fontWeight:700,marginTop:2}}>{sub.name}</h4></div>
            <div style={{display:"flex",gap:6}}><button onClick={()=>openEdit(sub)} style={{background:"none",border:"none",cursor:"pointer",color:"#2563EB"}}><Icon name="edit" size={14}/></button><button onClick={()=>{U.setSubjects(p=>p.filter(x=>x.id!==sub.id));st("ลบแล้ว","warning")}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444"}}><Icon name="trash" size={14}/></button></div>
          </div>
          <div style={{display:"flex",gap:6,marginTop:10,flexWrap:"wrap"}}>
            <span style={{background:c.lt,color:c.tx,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>{dept?.name||"?"}</span>
            {lv&&<span style={{background:"#DBEAFE",color:"#1E40AF",padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>{lv.name}</span>}
            <span style={{background:"#F3F4F6",padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>{sub.credits} หน่วยกิต</span>
            <span style={{background:"#F3F4F6",padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>{sub.periodsPerWeek} คาบ/สัปดาห์</span>
          </div>
        </div>
      </div>})}
    </div>
    <Modal open={modal} onClose={()=>{setModal(false);setEditId(null)}} title={editId?"แก้ไขวิชา":"เพิ่มวิชา"}>
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <div><label style={LS}>รหัสวิชา</label><input style={IS} value={form.code} onChange={e=>setForm(p=>({...p,code:e.target.value}))} placeholder="ว33202"/></div>
        <div><label style={LS}>ชื่อวิชา</label><input style={IS} value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))} placeholder="ฟิสิกส์ 4"/></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          <div><label style={LS}>หน่วยกิต</label><input type="number" min="0.5" step="0.5" style={IS} value={form.credits} onChange={e=>setForm(p=>({...p,credits:parseFloat(e.target.value)||0}))}/></div>
          <div><label style={LS}>คาบ/สัปดาห์</label><input type="number" min="1" style={IS} value={form.periodsPerWeek} onChange={e=>setForm(p=>({...p,periodsPerWeek:parseInt(e.target.value)||1}))}/></div>
        </div>
        <div><label style={LS}>ระดับชั้น</label><select style={IS} value={form.levelId} onChange={e=>setForm(p=>({...p,levelId:e.target.value}))}><option value="">--</option>{S.levels.map(l=><option key={l.id} value={l.id}>{l.name}</option>)}</select></div>
        <div><label style={LS}>กลุ่มสาระ</label><select style={IS} value={form.departmentId} onChange={e=>setForm(p=>({...p,departmentId:e.target.value}))}><option value="">--</option>{S.depts.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}</select></div>
        <button onClick={save} style={BS()}>{editId?"บันทึก":"เพิ่มวิชา"}</button>
      </div>
    </Modal>
  </div>;
}

/* ===== ASSIGNMENTS (fix#3,#4,#5: level-filter rooms, show code+name, countdown) ===== */
function Assigns({S,U,st,gc}){
  const [selDept,setSelDept]=useState("");
  const [sel,setSel]=useState("");
  const [modal,setModal]=useState(false);
  const [form,setForm]=useState({subjectId:"",roomIds:[],totalPeriods:0});
  const deptTeachers=selDept?S.teachers.filter(t=>t.departmentId===selDept):[];
  const teacher=S.teachers.find(t=>t.id===sel);
  const asgns=S.assigns.filter(a=>a.teacherId===sel);
  const totalUsed=asgns.reduce((s,a)=>s+a.totalPeriods,0);
  const teacherQuota=teacher?.totalPeriods||0;
  const remaining=teacherQuota-totalUsed;

  // fix#3: when subject selected, show rooms of that level only
  const selSub=S.subjects.find(s=>s.id===form.subjectId);
  const filteredRooms=selSub?S.rooms.filter(r=>r.levelId===selSub.levelId):S.rooms;

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"flex",gap:12,marginBottom:24,alignItems:"center",flexWrap:"wrap"}}>
      <select style={{...IS,maxWidth:280}} value={selDept} onChange={e=>{setSelDept(e.target.value);setSel("")}}><option value="">-- เลือกกลุ่มสาระก่อน --</option>{S.depts.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}</select>
      {selDept&&<select style={{...IS,maxWidth:350}} value={sel} onChange={e=>setSel(e.target.value)}><option value="">-- เลือกครู --</option>{deptTeachers.map(t=><option key={t.id} value={t.id}>{t.prefix}{t.firstName} {t.lastName}</option>)}</select>}
      {sel&&<button onClick={()=>{setForm({subjectId:"",roomIds:[],totalPeriods:0});setModal(true)}} style={BS()}><Icon name="plus" size={16}/>เพิ่มวิชา</button>}
    </div>
    {teacher&&<div>
      <div style={{background:"#fff",borderRadius:14,padding:20,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",marginBottom:20,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
        <div><h3 style={{fontSize:18,fontWeight:700}}>{teacher.prefix}{teacher.firstName} {teacher.lastName}</h3><div style={{fontSize:13,color:"#6B7280",marginTop:4}}>{S.depts.find(d=>d.id===teacher.departmentId)?.name}</div></div>
        <div style={{display:"flex",gap:12}}>
          <div style={{background:"#DBEAFE",color:"#1E40AF",padding:"8px 20px",borderRadius:10,fontWeight:700}}>คาบได้รับ: {teacherQuota}</div>
          <div style={{background:"#FEF3C7",color:"#92400E",padding:"8px 20px",borderRadius:10,fontWeight:700}}>มอบหมาย: {totalUsed}</div>
          <div style={{background:remaining>=0?"#D1FAE5":"#FEE2E2",color:remaining>=0?"#065F46":"#991B1B",padding:"8px 20px",borderRadius:10,fontWeight:700}}>เหลือ: {remaining}</div>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:16}}>
        {asgns.map(a=>{const sub=S.subjects.find(s=>s.id===a.subjectId);const dept=S.depts.find(d=>d.id===sub?.departmentId);const c=dept?gc(dept.id):{bg:"#6B7280",lt:"#F3F4F6",tx:"#374151"};return<div key={a.id} style={{background:"#fff",borderRadius:14,borderLeft:`4px solid ${c.bg}`,padding:16,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <div style={{display:"flex",justifyContent:"space-between"}}><div><h4 style={{fontSize:15,fontWeight:700}}>{sub?.code} — {sub?.name}</h4><div style={{fontSize:12,color:"#6B7280",marginTop:4}}>{a.totalPeriods} คาบ/สัปดาห์</div></div>
            <div style={{display:"flex",gap:6}}>
              <button onClick={()=>{const n=prompt("แก้ไขจำนวนคาบ:",a.totalPeriods);if(n!==null){U.setAssigns(p=>p.map(x=>x.id===a.id?{...x,totalPeriods:parseInt(n)||1}:x));st("แก้ไขสำเร็จ")}}} style={{background:"none",border:"none",cursor:"pointer",color:"#2563EB"}}><Icon name="edit" size={14}/></button>
              <button onClick={()=>{U.setAssigns(p=>p.filter(x=>x.id!==a.id));st("ลบแล้ว","warning")}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444"}}><Icon name="trash" size={14}/></button>
            </div>
          </div>
          <div style={{display:"flex",gap:6,marginTop:10,flexWrap:"wrap"}}>{a.roomIds.map(rid=><span key={rid} style={{background:"#DBEAFE",color:"#1E40AF",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>{S.rooms.find(r=>r.id===rid)?.name}</span>)}</div>
        </div>})}
      </div>
    </div>}
    <Modal open={modal} onClose={()=>setModal(false)} title="มอบหมายวิชา">
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <div><label style={LS}>วิชา (รหัส — ชื่อ)</label><select style={IS} value={form.subjectId} onChange={e=>{setForm(p=>({...p,subjectId:e.target.value,roomIds:[]}))}}><option value="">--</option>{S.subjects.map(s=><option key={s.id} value={s.id}>{s.code} — {s.name} ({S.levels.find(l=>l.id===s.levelId)?.name})</option>)}</select></div>
        <div><label style={LS}>ห้อง (เฉพาะระดับของวิชา)</label><div style={{display:"flex",gap:8,flexWrap:"wrap",maxHeight:200,overflowY:"auto"}}>{filteredRooms.map(rm=><button key={rm.id} onClick={()=>setForm(p=>({...p,roomIds:p.roomIds.includes(rm.id)?p.roomIds.filter(r=>r!==rm.id):[...p.roomIds,rm.id]}))} style={{padding:"6px 14px",borderRadius:8,border:`2px solid ${form.roomIds.includes(rm.id)?"#DC2626":"#D1D5DB"}`,background:form.roomIds.includes(rm.id)?"#FEE2E2":"#fff",fontSize:13,fontWeight:600,cursor:"pointer"}}>{form.roomIds.includes(rm.id)?"✓ ":""}{rm.name}</button>)}</div></div>
        <div><label style={LS}>คาบรวม (0=อัตโนมัติ)</label><input type="number" min="0" style={IS} value={form.totalPeriods} onChange={e=>setForm(p=>({...p,totalPeriods:parseInt(e.target.value)||0}))}/></div>
        {remaining<=0&&<div style={{padding:12,background:"#FEE2E2",borderRadius:10,color:"#991B1B",fontSize:13,fontWeight:600}}>⚠️ คาบที่ได้รับหมดแล้ว!</div>}
        <button onClick={()=>{if(!form.subjectId||!form.roomIds.length){st("เลือกวิชาและห้อง","error");return}const sub=S.subjects.find(s=>s.id===form.subjectId);const tp=form.totalPeriods||(sub?.periodsPerWeek||1)*form.roomIds.length;U.setAssigns(p=>[...p,{id:gid(),teacherId:sel,subjectId:form.subjectId,roomIds:form.roomIds,totalPeriods:tp}]);setForm({subjectId:"",roomIds:[],totalPeriods:0});setModal(false);st("มอบหมายสำเร็จ")}} style={BS()}>บันทึก</button>
      </div>
    </Modal>
  </div>;
}

/* ===== MEETINGS ===== */
function Meetings({S,U,st,gc}){
  const [form,setForm]=useState({departmentId:"",day:"",periods:[]});
  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",marginBottom:24,maxWidth:600}}>
      <h3 style={{fontSize:16,fontWeight:700,marginBottom:16}}>เพิ่มคาบล็อค</h3>
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <div><label style={LS}>กลุ่มสาระ</label><select style={IS} value={form.departmentId} onChange={e=>setForm(p=>({...p,departmentId:e.target.value}))}><option value="">--</option>{S.depts.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}</select></div>
        <div><label style={LS}>วัน</label><select style={IS} value={form.day} onChange={e=>setForm(p=>({...p,day:e.target.value}))}><option value="">--</option>{DAYS.map(d=><option key={d}>{d}</option>)}</select></div>
        <div><label style={LS}>คาบ</label><div style={{display:"flex",gap:8}}>{PERIODS.map(p=><button key={p.id} onClick={()=>setForm(prev=>({...prev,periods:prev.periods.includes(p.id)?prev.periods.filter(x=>x!==p.id):[...prev.periods,p.id]}))} style={{width:48,height:48,borderRadius:10,border:`2px solid ${form.periods.includes(p.id)?"#DC2626":"#D1D5DB"}`,background:form.periods.includes(p.id)?"#DC2626":"#fff",color:form.periods.includes(p.id)?"#fff":"#374151",fontSize:16,fontWeight:700,cursor:"pointer"}}>{p.id}</button>)}</div></div>
        <button onClick={()=>{if(!form.departmentId||!form.day||!form.periods.length){st("กรอกให้ครบ","error");return}U.setMeetings(p=>[...p,{id:gid(),...form}]);setForm({departmentId:"",day:"",periods:[]});st("เพิ่มสำเร็จ")}} style={BS()}>เพิ่มคาบล็อค</button>
      </div>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:16}}>
      {S.meetings.map(m=>{const dept=S.depts.find(d=>d.id===m.departmentId);const c=dept?gc(dept.id):{bg:"#6B7280"};return<div key={m.id} style={{background:"#fff",borderRadius:14,borderLeft:`4px solid ${c.bg}`,padding:16,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{display:"flex",justifyContent:"space-between"}}><div><h4 style={{fontSize:15,fontWeight:700}}>{dept?.name}</h4><div style={{fontSize:13,color:"#6B7280",marginTop:4}}>วัน{m.day} — คาบ {m.periods.sort().join(", ")}</div></div>
          <button onClick={()=>{U.setMeetings(p=>p.filter(x=>x.id!==m.id));st("ลบแล้ว","warning")}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444"}}><Icon name="trash" size={14}/></button>
        </div>
      </div>})}
    </div>
  </div>;
}

/* ===== SCHEDULER (fix#6,#7,#8,#9) ===== */
function Scheduler({S,U,st,gc}){
  const [selDept,setSelDept]=useState("");
  const [selT,setSelT]=useState("");
  const [drag,setDrag]=useState(null);
  const [coM,setCoM]=useState(null);
  const [coS,setCoS]=useState("");
  const teacher=S.teachers.find(t=>t.id===selT);
  const asgns=S.assigns.filter(a=>a.teacherId===selT);
  const tRooms=[...new Set(asgns.flatMap(a=>a.roomIds))];

  const blocked=useCallback(tid=>{const t=S.teachers.find(x=>x.id===tid);if(!t)return[];const b=[];(t.specialRoles||[]).forEach(rid=>{const r=SROLES.find(x=>x.id===rid);r?.blocked?.forEach(bl=>bl.periods.forEach(p=>b.push({day:bl.day,period:p,reason:r.name})))});S.meetings.filter(m=>m.departmentId===t.departmentId).forEach(m=>m.periods.forEach(p=>b.push({day:m.day,period:p,reason:"ประชุม"})));return b},[S.teachers,S.meetings]);

  const isBlk=(tid,day,p)=>blocked(tid).some(b=>b.day===day&&b.period===p);
  const sk=(rid,day,p)=>`${rid}_${day}_${p}`;

  // fix#6: count how many periods of THIS subject+room are already placed
  const countSubjectInRoom=(assignId,roomId)=>{let c=0;Object.entries(S.schedule).forEach(([k,en])=>{if(!k.startsWith(roomId+"_"))return;en?.forEach(e=>{if(e.assignmentId===assignId)c++})});return c};

  // Get assignment's periodsPerWeek per room
  const getPerRoomLimit=(assignId)=>{const a=S.assigns.find(x=>x.id===assignId);if(!a)return 999;const sub=S.subjects.find(s=>s.id===a.subjectId);return sub?.periodsPerWeek||999};

  // aUsed counts periods used including co-teaching assignments from schedule
  const aUsed=(aid)=>{let c=0;Object.entries(S.schedule).forEach(([k,en])=>{en?.forEach(e=>{if(e.assignmentId===aid)c++})});return c};

  // Count ALL scheduled periods for a specific teacher (as main OR co-teacher)
  const teacherScheduledTotal=(tid)=>{let c=0;Object.values(S.schedule).forEach(en=>{en?.forEach(e=>{if(e.teacherId===tid||e.coTeacherId===tid)c++})});return c};

  const fTeachers=selDept?S.teachers.filter(t=>t.departmentId===selDept):S.teachers;

  // Fix#3: handle drop for BOTH new cards from sidebar AND re-dragging existing entries
  // Check if teacher is already scheduled in ANY room at this day+period
  const teacherBusy=(tid,day,period,excludeKey)=>{
    let busy=false;
    Object.entries(S.schedule).forEach(([k,en])=>{
      if(k===excludeKey)return;
      if(!k.endsWith(`_${day}_${period}`))return;
      en?.forEach(e=>{if(e.teacherId===tid||e.coTeacherId===tid)busy=true});
    });
    return busy;
  };

  const handleDrop=(rid,day,p)=>{
    const key=sk(rid,day,p);
    if(S.locks[key]){st("ล็อคแล้ว","error");return}
    if((S.schedule[key]||[]).length>=3){st("ครบ 3 วิชาแล้ว","error");return}

    // If re-dragging an existing entry from another cell
    if(drag?.fromKey){
      if(drag.fromKey===key)return;
      const entry=drag.entry;
      if(isBlk(entry.teacherId,day,p)){st("ครูถูกล็อคคาบนี้","error");return}
      // Check teacher conflict (exclude the source cell since we're moving FROM there)
      if(teacherBusy(entry.teacherId,day,p,drag.fromKey)){st("ครูคนนี้สอนคาบนี้อยู่แล้ว (ห้องอื่น)","error");return}
      const room=S.rooms.find(r=>r.id===rid);
      const sub=S.subjects.find(s=>s.id===entry.subjectId);
      if(room&&sub&&room.levelId!==sub.levelId){st("ระดับชั้นไม่ตรงกัน!","error");return}
      U.setSchedule(prev=>{
        const updated={...prev};
        updated[drag.fromKey]=(updated[drag.fromKey]||[]).filter(e=>e.id!==entry.id);
        updated[key]=[...(updated[key]||[]),entry];
        return updated;
      });
      setDrag(null);
      return;
    }

    // Normal new card from sidebar
    if(!drag)return;
    if(isBlk(drag.teacherId,day,p)){st("ครูถูกล็อคคาบนี้","error");return}
    // Check teacher conflict
    if(teacherBusy(drag.teacherId,day,p,null)){st("ครูคนนี้สอนคาบนี้อยู่แล้ว (ห้องอื่น)","error");return}
    const room=S.rooms.find(r=>r.id===rid);
    const sub=S.subjects.find(s=>s.id===drag.subjectId);
    if(room&&sub&&room.levelId!==sub.levelId){st("ระดับชั้นไม่ตรงกัน!","error");return}
    const placed=countSubjectInRoom(drag.assignmentId,rid);
    const limit=getPerRoomLimit(drag.assignmentId);
    if(placed>=limit){st(`ห้องนี้ลงครบ ${limit} คาบแล้ว`,"error");return}
    U.setSchedule(prev=>({...prev,[key]:[...(prev[key]||[]),{id:gid(),teacherId:drag.teacherId,subjectId:drag.subjectId,assignmentId:drag.assignmentId,coTeacherId:null}]}));
    setDrag(null);
  };

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"flex",gap:12,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
      <select style={{...IS,maxWidth:250}} value={selDept} onChange={e=>{setSelDept(e.target.value);setSelT("")}}><option value="">-- ทุกกลุ่มสาระ --</option>{S.depts.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}</select>
      <select style={{...IS,maxWidth:300}} value={selT} onChange={e=>setSelT(e.target.value)}><option value="">-- เลือกครู --</option>{fTeachers.map(t=>{const sched=teacherScheduledTotal(t.id);const rem=(t.totalPeriods||0)-sched;return<option key={t.id} value={t.id}>{t.prefix}{t.firstName} {t.lastName} (เหลือ {rem})</option>})}</select>
    </div>

    {/* Teacher summary bar */}
    {teacher&&<div style={{background:"#fff",borderRadius:12,padding:"12px 20px",marginBottom:16,display:"flex",gap:16,alignItems:"center",flexWrap:"wrap",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
      <div style={{fontSize:16,fontWeight:700}}>{teacher.prefix}{teacher.firstName} {teacher.lastName}</div>
      <div style={{fontSize:13,color:"#6B7280"}}>{S.depts.find(d=>d.id===teacher.departmentId)?.name}</div>
      <div style={{marginLeft:"auto",display:"flex",gap:12}}>
        <div style={{background:"#DBEAFE",color:"#1E40AF",padding:"6px 16px",borderRadius:8,fontWeight:700,fontSize:14}}>ได้รับ {teacher.totalPeriods||0}</div>
        <div style={{background:"#FEF3C7",color:"#92400E",padding:"6px 16px",borderRadius:8,fontWeight:700,fontSize:14}}>จัดแล้ว {teacherScheduledTotal(teacher.id)}</div>
        <div style={{background:(teacher.totalPeriods||0)-teacherScheduledTotal(teacher.id)>0?"#D1FAE5":"#FEE2E2",color:(teacher.totalPeriods||0)-teacherScheduledTotal(teacher.id)>0?"#065F46":"#991B1B",padding:"6px 16px",borderRadius:8,fontWeight:700,fontSize:14}}>เหลือ {(teacher.totalPeriods||0)-teacherScheduledTotal(teacher.id)}</div>
      </div>
    </div>}

    {teacher?<div style={{display:"flex",gap:16}}>
      {/* Sidebar - wider, bigger text */}
      <div style={{width:300,flexShrink:0,position:"sticky",top:0,alignSelf:"flex-start",maxHeight:"calc(100vh - 200px)",overflowY:"auto"}}>
        <h4 style={{fontSize:15,fontWeight:700,color:"#374151",marginBottom:12}}>วิชา — ลากวาง</h4>
        {asgns.map(a=>{const sub=S.subjects.find(s=>s.id===a.subjectId);const dept=S.depts.find(d=>d.id===sub?.departmentId);const c=dept?gc(dept.id):{bg:"#6B7280",lt:"#F3F4F6",tx:"#374151",bd:"#D1D5DB"};const u=aUsed(a.id);const rem=a.totalPeriods-u;return<div key={a.id} className="drag-card" draggable={rem>0} onDragStart={()=>setDrag({teacherId:selT,subjectId:a.subjectId,assignmentId:a.id})} onDragEnd={()=>setDrag(null)} style={{background:c.lt,border:`2px solid ${c.bd}`,borderRadius:12,padding:16,opacity:rem<=0?0.4:1,marginBottom:10}}>
          <div style={{fontSize:15,fontWeight:700,color:c.tx}}>{sub?.code} — {sub?.name}</div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:10}}><div style={{display:"flex",gap:4,flexWrap:"wrap"}}>{a.roomIds.map(rid=><span key={rid} style={{background:"rgba(0,0,0,0.1)",padding:"2px 10px",borderRadius:10,fontSize:12,fontWeight:600}}>{S.rooms.find(r=>r.id===rid)?.name}</span>)}</div><span style={{background:rem>0?c.bg:"#9CA3AF",color:"#fff",padding:"4px 14px",borderRadius:20,fontSize:13,fontWeight:700}}>{rem}/{a.totalPeriods}</span></div>
        </div>})}
      </div>
      {/* Table area - bigger */}
      <div style={{flex:1,overflowX:"auto"}}>
        {tRooms.map(rid=>{const rm=S.rooms.find(r=>r.id===rid);return<div key={rid} style={{marginBottom:24}}>
          <h4 style={{fontSize:15,fontWeight:700,marginBottom:10}}><span style={{background:"#DC2626",color:"#fff",padding:"3px 12px",borderRadius:8,fontSize:12}}>{rm?.name}</span></h4>
          <div style={{background:"#fff",borderRadius:14,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",overflow:"hidden"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,minWidth:900}}>
              <thead><tr><th style={{padding:"12px 14px",background:"#F9FAFB",fontWeight:700,color:"#6B7280",width:90,textAlign:"left",fontSize:14}}>วัน</th>{PERIODS.map(p=><th key={p.id} style={{padding:"10px 8px",background:"#F9FAFB",fontWeight:600,color:"#6B7280",textAlign:"center",minWidth:120}}><div style={{fontSize:13}}>คาบ {p.id}</div><div style={{fontSize:11,fontWeight:400}}>{p.time}</div></th>)}</tr></thead>
              <tbody>{DAYS.map(day=><tr key={day}><td style={{padding:"10px 14px",fontWeight:700,borderTop:"1px solid #E5E7EB",fontSize:14}}>{day}</td>
                {PERIODS.map(p=>{const key=sk(rid,day,p.id);const en=S.schedule[key]||[];const lk=S.locks[key];const bl=isBlk(selT,day,p.id);return<td key={p.id} className="dz" onDragOver={e=>{e.preventDefault();e.currentTarget.classList.add("over")}} onDragLeave={e=>e.currentTarget.classList.remove("over")} onDrop={e=>{e.preventDefault();e.currentTarget.classList.remove("over");handleDrop(rid,day,p.id)}} style={{padding:4,borderTop:"1px solid #F3F4F6",verticalAlign:"top",background:bl?"#FEF3C7":lk?"#F0FDF4":"#fff",minHeight:60}}>
                  {bl&&!en.length&&<div style={{fontSize:10,color:"#92400E",textAlign:"center",padding:4}}>🔒 {blocked(selT).find(b=>b.day===day&&b.period===p.id)?.reason}</div>}
                  {en.map(entry=>{const sub=S.subjects.find(s=>s.id===entry.subjectId);const dept=S.depts.find(d=>d.id===sub?.departmentId);const c=dept?gc(dept.id):{bg:"#6B7280",lt:"#F3F4F6",tx:"#374151",bd:"#D1D5DB"};const et=S.teachers.find(t=>t.id===entry.teacherId);const ct=entry.coTeacherId?S.teachers.find(t=>t.id===entry.coTeacherId):null;
                    /* Fix#3: make placed cards draggable to move them */
                    return<div key={entry.id} draggable={!lk} onDragStart={(e)=>{e.stopPropagation();setDrag({fromKey:key,entry})}} onDragEnd={()=>setDrag(null)} style={{background:c.lt,border:`1.5px solid ${c.bd}`,borderRadius:8,padding:"6px 8px",marginBottom:3,fontSize:12,position:"relative",cursor:lk?"default":"grab"}}>
                      <div style={{fontWeight:700,color:c.tx,fontSize:12}}>{sub?.code}</div>
                      <div style={{fontWeight:600,color:c.tx,fontSize:11}}>{sub?.name}</div>
                      <div style={{color:c.tx,opacity:0.7,fontSize:11}}>{et?.firstName}{ct?` + ${ct.firstName}`:""}</div>
                      {!lk&&<div style={{display:"flex",gap:2,marginTop:2}}>
                        <button onClick={()=>U.setSchedule(prev=>({...prev,[key]:(prev[key]||[]).filter(e=>e.id!==entry.id)}))} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",padding:0}}><Icon name="x" size={10}/></button>
                        <button onClick={()=>setCoM({key,entryId:entry.id})} style={{background:"none",border:"none",cursor:"pointer",color:"#2563EB",padding:0}}><Icon name="users" size={10}/></button>
                        <button onClick={()=>{U.setLocks(prev=>({...prev,[key]:true}));st("ล็อคแล้ว")}} style={{background:"none",border:"none",cursor:"pointer",color:"#059669",padding:0}}><Icon name="lock" size={10}/></button>
                      </div>}
                      {lk&&<div style={{position:"absolute",top:2,right:4}}><button onClick={()=>{U.setLocks(prev=>({...prev,[key]:false}));st("ปลดล็อค")}} style={{background:"none",border:"none",cursor:"pointer",color:"#059669",padding:0}}><Icon name="unlock" size={10}/></button></div>}
                    </div>})}
                </td>})}</tr>)}</tbody>
            </table>
          </div>
        </div>})}
        {!tRooms.length&&<div style={{padding:40,textAlign:"center",color:"#9CA3AF"}}>เลือกครูที่มอบหมายวิชาแล้ว</div>}
      </div>
    </div>
    :<div style={{background:"#fff",borderRadius:14,padding:60,textAlign:"center"}}><div style={{fontSize:48,marginBottom:16}}>📋</div><h3 style={{fontSize:18,fontWeight:700,marginBottom:8}}>เลือกครูเพื่อจัดตาราง</h3></div>}

    <Modal open={!!coM} onClose={()=>{setCoM(null);setCoS("")}} title="เพิ่มครูสอนร่วม">
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <div><label style={LS}>ค้นหาและเลือกครูสอนร่วม</label>
        <select style={IS} value={coS} onChange={e=>setCoS(e.target.value)}><option value="">-- เลือกครู --</option>{S.teachers.filter(t=>t.id!==selT).map(t=>{
          const sched=teacherScheduledTotal(t.id);const rem=(t.totalPeriods||0)-sched;
          return<option key={t.id} value={t.id}>{t.prefix}{t.firstName} {t.lastName} — เหลือ {rem} คาบ</option>
        })}</select></div>
        {coS&&(()=>{
          // Extract day+period from coM.key
          const parts=coM?.key?.split("_")||[];
          const cDay=parts[1];const cPer=parseInt(parts[2]);
          const isBusy=teacherBusy(coS,cDay,cPer,null);
          const coTeacher=S.teachers.find(t=>t.id===coS);
          const sched=teacherScheduledTotal(coS);const rem=(coTeacher?.totalPeriods||0)-sched;
          return <div>
            {isBusy&&<div style={{padding:12,background:"#FEE2E2",borderRadius:10,color:"#991B1B",fontSize:13,fontWeight:600,marginBottom:8}}>⚠️ ครูท่านนี้สอนคาบนี้อยู่แล้ว (ห้องอื่น)</div>}
            {rem<=0&&<div style={{padding:12,background:"#FEF3C7",borderRadius:10,color:"#92400E",fontSize:13,fontWeight:600,marginBottom:8}}>⚠️ ครูท่านนี้คาบเต็มแล้ว ({coTeacher?.totalPeriods} คาบ)</div>}
            <div style={{fontSize:13,color:"#6B7280"}}>คาบที่ใช้: {sched}/{coTeacher?.totalPeriods||0} | เหลือ: {rem}</div>
          </div>
        })()}
        <button onClick={()=>{
          if(!coS||!coM)return;
          // Validate co-teacher not busy
          const parts=coM.key.split("_");const cDay=parts[1];const cPer=parseInt(parts[2]);
          if(teacherBusy(coS,cDay,cPer,null)){st("ครูท่านนี้สอนคาบนี้อยู่แล้ว","error");return}
          U.setSchedule(prev=>({...prev,[coM.key]:(prev[coM.key]||[]).map(e=>e.id===coM.entryId?{...e,coTeacherId:coS}:e)}));
          setCoM(null);setCoS("");st("เพิ่มครูร่วมสำเร็จ")
        }} style={BS()}>เพิ่มครูสอนร่วม</button>
      </div>
    </Modal>
  </div>;
}

/* ===== REPORTS (fix#10: working page) ===== */
function Reports({S,st,gc,ay,sh}){
  const roomSt=S.rooms.map(rm=>{let f=0;DAYS.forEach(d=>PERIODS.forEach(p=>{const k=`${rm.id}_${d}_${p.id}`;if(S.schedule[k]?.length)f++}));const total=DAYS.length*PERIODS.length;return{room:rm,filled:f,total,pct:Math.round(f/total*100)}});
  const teacherSt=S.teachers.map(t=>{const tot=t.totalPeriods||0;let u=0;Object.values(S.schedule).forEach(en=>{en?.forEach(e=>{if(e.teacherId===t.id||e.coTeacherId===t.id)u++})});return{teacher:t,tot,used:u,rem:tot-u}});

  const exportRoomXL=(rm)=>{const h=["วัน",...PERIODS.map(p=>`คาบ${p.id}(${p.time})`)];const d=DAYS.map(day=>[day,...PERIODS.map(p=>{const en=S.schedule[`${rm.id}_${day}_${p.id}`]||[];return en.map(e=>{const sub=S.subjects.find(s=>s.id===e.subjectId);const t=S.teachers.find(x=>x.id===e.teacherId);return`${sub?.code||""} ${sub?.name||""} (${t?.firstName||""})`}).join(" / ")})]);exportExcel(h,d,`ตารางเรียน_${rm.name}.xlsx`,rm.name);st(`Export ${rm.name}`)};

  const exportTeacherXL=(t)=>{const h=["วัน",...PERIODS.map(p=>`คาบ${p.id}(${p.time})`)];const d=DAYS.map(day=>[day,...PERIODS.map(p=>{let parts=[];Object.entries(S.schedule).forEach(([k,en])=>{if(!k.endsWith(`_${day}_${p.id}`))return;en?.forEach(e=>{if(e.teacherId===t.id||e.coTeacherId===t.id){const sub=S.subjects.find(s=>s.id===e.subjectId);const rid=k.split("_")[0];const rm=S.rooms.find(r=>r.id===rid);parts.push(`${sub?.code||""} ${sub?.name||""} (${rm?.name||""})`)}})});return parts.join(" / ")})]);exportExcel(h,d,`ตารางสอน_${t.prefix}${t.firstName}.xlsx`,"ตารางสอน");st(`Export ${t.firstName}`)};

  const exportAllRooms=()=>{exportExcelMulti(S.rooms.map(rm=>({name:rm.name,headers:["วัน",...PERIODS.map(p=>`คาบ${p.id}(${p.time})`)],rows:DAYS.map(day=>[day,...PERIODS.map(p=>{const en=S.schedule[`${rm.id}_${day}_${p.id}`]||[];return en.map(e=>{const sub=S.subjects.find(s=>s.id===e.subjectId);const t=S.teachers.find(x=>x.id===e.teacherId);return`${sub?.code||""} ${sub?.name||""} (${t?.firstName||""})`}).join(" / ")})])})),"ตารางเรียนทุกห้อง.xlsx");st("Export ทุกห้อง")};

  const exportAllTeachers=()=>{exportExcelMulti(S.teachers.map(t=>({name:`${t.firstName} ${t.lastName}`,headers:["วัน",...PERIODS.map(p=>`คาบ${p.id}(${p.time})`)],rows:DAYS.map(day=>[day,...PERIODS.map(p=>{let parts=[];Object.entries(S.schedule).forEach(([k,en])=>{if(!k.endsWith(`_${day}_${p.id}`))return;en?.forEach(e=>{if(e.teacherId===t.id||e.coTeacherId===t.id){const sub=S.subjects.find(s=>s.id===e.subjectId);const rid=k.split("_")[0];const rm=S.rooms.find(r=>r.id===rid);parts.push(`${sub?.code||""} ${sub?.name||""} (${rm?.name||""})`)}})});return parts.join(" / ")})])})),"ตารางสอนทุกคน.xlsx");st("Export ทุกคน")};

  const exportStatus=()=>{
    const sheets=[{name:"ห้องเรียน",headers:["ห้อง","จัดแล้ว","ทั้งหมด","%"],rows:roomSt.map(r=>[r.room.name,r.filled,r.total,`${r.pct}%`])},{name:"ครู",headers:["ชื่อ","คาบได้รับ","จัดแล้ว","เหลือ","สถานะ"],rows:teacherSt.filter(t=>t.tot>0).map(t=>[`${t.teacher.prefix}${t.teacher.firstName} ${t.teacher.lastName}`,t.tot,t.used,t.rem,t.rem===0?"ครบ":"เหลือ "+t.rem])}];
    exportExcelMulti(sheets,"รายงานสถานะ.xlsx");st("Export สำเร็จ");
  };

  // PDF print for teacher
  // PDF: ตารางสอนครู — แสดง วิชา + ห้อง (ไม่มีครูร่วม)
  const printTeacherPDF=(t)=>{
    const w=window.open('','_blank');
    let html=pdfPage(
      "ตารางสอน "+(t.prefix||"")+(t.firstName||"")+" "+(t.lastName||""),
      "ภาคเรียนที่ "+(ay?.semester||"1")+"/"+(ay?.year||"2568")+" "+(sh?.name||"โรงเรียนดาราวิทยาลัย"),
      DAYS.map(day=>({day,cells:PERIODS.map(p=>{
        let parts=[];
        Object.entries(S.schedule).forEach(([k,en])=>{
          const keySuffix="_"+day+"_"+p.id;
          if(!k.endsWith(keySuffix))return;
          en?.forEach(e=>{
            if(e.teacherId===t.id||e.coTeacherId===t.id){
              const sub=S.subjects.find(s=>s.id===e.subjectId);
              const rid=k.split("_")[0];
              const rm=S.rooms.find(r=>r.id===rid);
              parts.push({sub:sub?.name||"",room:rm?.name||"",room2:""});
            }
          });
        });
        return parts;
      })})),
      "",
      sh?.logo||null
    );
    w.document.write(html);w.document.close();setTimeout(()=>w.print(),600);
  };

  // helper: สร้าง pages สำหรับห้องหนึ่ง
  // จำนวนใบ = จำนวน entry สูงสุดในคาบที่ซ้อนมากที่สุด (ปกติ 1, ถ้ามีซ้อน 2→2ใบ, 3→3ใบ)
  // ใบที่ i: แต่ละคาบเลือก entry[i] ถ้ามี, ถ้าไม่มีใช้ entry[0] (คาบปกติแสดงเหมือนกันทุกใบ)
  const buildRoomPages=(room)=>{
    const subtitle="ภาคเรียนที่ "+(ay?.semester||"1")+"/"+(ay?.year||"2568")+" "+(sh?.name||"โรงเรียนดาราวิทยาลัย");
    // หา maxEntries ของห้องนี้
    let maxEntries=1;
    DAYS.forEach(day=>PERIODS.forEach(p=>{
      const len=(S.schedule[room.id+"_"+day+"_"+p.id]||[]).length;
      if(len>maxEntries) maxEntries=len;
    }));
    // สร้าง maxEntries ใบ
    return Array.from({length:maxEntries},(_,sheetIdx)=>({
      title:"ตารางเรียน "+room.name+(maxEntries>1?" (ฉบับที่ "+(sheetIdx+1)+"/"+maxEntries+")":""),
      subtitle:subtitle,
      dayRows:DAYS.map(day=>({day,cells:PERIODS.map(p=>{
        const en=S.schedule[room.id+"_"+day+"_"+p.id]||[];
        if(!en.length) return [];
        // คาบปกติ (1 entry) → ทุกใบแสดงเหมือนกัน
        // คาบซ้อน (>1 entry) → ใบที่ i แสดง entry[i], ถ้า i เกิน → entry[0]
        const e=en[sheetIdx]||en[0];
        const sub=S.subjects.find(s=>s.id===e.subjectId);
        const t2=S.teachers.find(x=>x.id===e.teacherId);
        return[{sub:sub?.name||"",room:(t2?.prefix||"")+(t2?.firstName||""),room2:""}];
      })}))
    }));
  };

  // PDF: ตารางเรียนห้อง — แยกใบตามครู
  const printRoomPDF=(room)=>{
    const pages=buildRoomPages(room);
    if(!pages.length){st("ยังไม่มีตารางในห้องนี้","error");return}
    const w=window.open('','_blank');
    w.document.write(pdfMultiPage(pages,sh?.logo||null));
    w.document.close();setTimeout(()=>w.print(),600);
  };

  // PDF: พิมพ์ตารางสอนครูทั้งหมด (2 คน/หน้า A4 แนวตั้ง)
  const printAllTeachersPDF=()=>{
    const teachers=S.teachers.filter(t=>t.totalPeriods>0);
    if(!teachers.length){st("ไม่มีครูที่กำหนดคาบ","error");return}
    const pages=teachers.map(t=>({
      title:"ตารางสอน "+(t.prefix||"")+(t.firstName||"")+" "+(t.lastName||""),
      subtitle:"ภาคเรียนที่ "+(ay?.semester||"1")+"/"+(ay?.year||"2568")+" "+(sh?.name||"โรงเรียนดาราวิทยาลัย"),
      dayRows:DAYS.map(day=>({day,cells:PERIODS.map(p=>{
        const keySuffix="_"+day+"_"+p.id;
        let parts=[];
        Object.entries(S.schedule).forEach(([k,en])=>{
          if(!k.endsWith(keySuffix))return;
          en?.forEach(e=>{
            if(e.teacherId===t.id||e.coTeacherId===t.id){
              const sub=S.subjects.find(s=>s.id===e.subjectId);
              const rid=k.split("_")[0];
              const rm=S.rooms.find(r=>r.id===rid);
              parts.push({sub:sub?.name||"",room:rm?.name||"",room2:""});
            }
          });
        });
        return parts;
      })}))
    }));
    const w=window.open('','_blank');
    w.document.write(pdfMultiPage(pages,sh?.logo||null));
    w.document.close();setTimeout(()=>w.print(),800);
    st("กำลังพิมพ์ตารางสอน "+teachers.length+" คน ("+Math.ceil(teachers.length/2)+" หน้า)");
  };

  // PDF: พิมพ์ตารางเรียนทุกห้อง — เรียงระดับชั้น ม.4→ม.5→ม.6 แล้วเรียงห้อง, แยกใบตามวิชาซ้อน
  const printAllRoomsPDF=()=>{
    if(!S.rooms.length){st("ไม่มีห้องเรียน","error");return}
    const sorted=[...S.rooms].sort((a,b)=>{
      const la=S.levels.find(l=>l.id===a.levelId)?.name||"";
      const lb=S.levels.find(l=>l.id===b.levelId)?.name||"";
      if(la!==lb) return la.localeCompare(lb,"th");
      return a.name.localeCompare(b.name,"th");
    });
    const pages=sorted.flatMap(room=>buildRoomPages(room));
    if(!pages.length){st("ยังไม่มีตารางในระบบ","error");return}
    const w=window.open('','_blank');
    w.document.write(pdfMultiPage(pages,sh?.logo||null));
    w.document.close();setTimeout(()=>w.print(),800);
    st("กำลังพิมพ์ตารางเรียน "+sorted.length+" ห้อง ("+pages.length+" ใบ)");
  };

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"flex",gap:10,marginBottom:24,flexWrap:"wrap"}}>
      <button onClick={exportAllRooms} style={BS("#2563EB")}><Icon name="download" size={16}/>ตารางทุกห้อง (.xlsx)</button>
      <button onClick={exportAllTeachers} style={BS("#7C3AED")}><Icon name="download" size={16}/>ตารางสอนทุกคน (.xlsx)</button>
      <button onClick={exportStatus} style={BS("#059669")}><Icon name="download" size={16}/>รายงานสถานะ (.xlsx)</button>
      <div style={{width:"100%",height:0,borderTop:"1px solid #E5E7EB",margin:"4px 0"}}/>
      <button onClick={printAllTeachersPDF} style={BS("#DC2626")}><Icon name="file" size={16}/>พิมพ์ตารางสอนทุกคน (PDF)</button>
      <button onClick={printAllRoomsPDF} style={BS("#DB2777")}><Icon name="file" size={16}/>พิมพ์ตารางเรียนทุกห้อง (PDF)</button>
    </div>

    <h3 style={{fontSize:18,fontWeight:700,marginBottom:20}}>สถานะห้องเรียน</h3>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:12,marginBottom:24}}>
      {roomSt.map(({room,filled,total,pct})=><div key={room.id} style={{padding:14,borderRadius:10,background:pct===100?"#F0FDF4":pct>0?"#FFFBEB":"#FEF2F2",border:`1px solid ${pct===100?"#BBF7D0":pct>0?"#FDE68A":"#FECACA"}`,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{display:"flex",justifyContent:"space-between"}}><span style={{fontWeight:700,fontSize:14}}>{room.name}</span><span style={{fontSize:12,fontWeight:700,color:pct===100?"#059669":pct>0?"#D97706":"#DC2626"}}>{pct}%</span></div>
        <div style={{height:6,background:"rgba(0,0,0,0.08)",borderRadius:3,marginTop:8,overflow:"hidden"}}><div style={{width:`${pct}%`,height:"100%",background:pct===100?"#059669":pct>0?"#D97706":"#DC2626",borderRadius:3}}/></div>
        <div style={{display:"flex",gap:6,marginTop:8}}>
          <button onClick={()=>exportRoomXL(room)} style={{background:"none",border:"1.5px solid #2563EB",borderRadius:6,padding:"3px 10px",color:"#2563EB",fontSize:11,fontWeight:600,cursor:"pointer"}}>Excel</button>
          <button onClick={()=>printRoomPDF(room)} style={{background:"none",border:"1.5px solid #DC2626",borderRadius:6,padding:"3px 10px",color:"#DC2626",fontSize:11,fontWeight:600,cursor:"pointer"}}>PDF</button>
        </div>
      </div>)}
      {!roomSt.length&&<div style={{padding:20,color:"#9CA3AF"}}>ยังไม่มีห้องเรียน</div>}
    </div>

    <h3 style={{fontSize:18,fontWeight:700,marginBottom:16}}>สถานะครู</h3>
    <div style={{background:"#fff",borderRadius:14,padding:20,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
        <thead><tr style={{background:"#F9FAFB"}}>{["ชื่อ","คาบได้รับ","จัดแล้ว","เหลือ","สถานะ","Export"].map(h=><th key={h} style={{padding:"10px 14px",textAlign:"left",fontWeight:600,color:"#6B7280"}}>{h}</th>)}</tr></thead>
        <tbody>
          {teacherSt.filter(t=>t.tot>0).map(ts=><tr key={ts.teacher.id} style={{borderTop:"1px solid #F3F4F6"}}>
            <td style={{padding:"10px 14px",fontWeight:600}}>{ts.teacher.prefix}{ts.teacher.firstName} {ts.teacher.lastName}</td>
            <td style={{padding:"10px 14px"}}>{ts.tot}</td>
            <td style={{padding:"10px 14px"}}>{ts.used}</td>
            <td style={{padding:"10px 14px",fontWeight:700,color:ts.rem>0?"#D97706":"#059669"}}>{ts.rem}</td>
            <td style={{padding:"10px 14px"}}>{ts.rem===0?<span style={{background:"#D1FAE5",color:"#065F46",padding:"3px 12px",borderRadius:20,fontSize:11,fontWeight:700}}>ครบ</span>:<span style={{background:"#FEF3C7",color:"#92400E",padding:"3px 12px",borderRadius:20,fontSize:11,fontWeight:700}}>เหลือ {ts.rem}</span>}</td>
            <td style={{padding:"10px 14px"}}><div style={{display:"flex",gap:4}}>
              <button onClick={()=>exportTeacherXL(ts.teacher)} style={{background:"none",border:"1.5px solid #2563EB",borderRadius:6,padding:"3px 10px",color:"#2563EB",fontSize:11,fontWeight:600,cursor:"pointer"}}>Excel</button>
              <button onClick={()=>printTeacherPDF(ts.teacher)} style={{background:"none",border:"1.5px solid #DC2626",borderRadius:6,padding:"3px 10px",color:"#DC2626",fontSize:11,fontWeight:600,cursor:"pointer"}}>PDF</button>
            </div></td>
          </tr>)}
          {!teacherSt.filter(t=>t.tot>0).length&&<tr><td colSpan={6} style={{padding:30,textAlign:"center",color:"#9CA3AF"}}>ยังไม่มีครูที่กำหนดคาบ</td></tr>}
        </tbody>
      </table>
    </div>
  </div>;
}
/* ===== SETTINGS (Fix#4: academic year, school header, reset) ===== */
function Settings({S,U,st,ay,setAY,sh,setSH}){
  const logoRef=useRef(null);
  const resetAll=()=>{
    if(!confirm("⚠️ คุณแน่ใจหรือไม่ว่าต้องการลบข้อมูลทั้งหมด?\nข้อมูลที่จัดตารางไว้จะหายทั้งหมด!"))return;
    if(!confirm("ยืนยันอีกครั้ง — ลบข้อมูลทั้งหมดและเริ่มต้นใหม่?"))return;
    U.setLevels([{id:gid(),name:"ม.4"},{id:gid(),name:"ม.5"},{id:gid(),name:"ม.6"}]);
    U.setPlans([]);U.setDepts([]);U.setTeachers([]);U.setSubjects([]);
    U.setRooms([]);U.setAssigns([]);U.setMeetings([]);U.setSchedule({});U.setLocks({});
    st("รีเซ็ทข้อมูลทั้งหมดแล้ว","warning");
  };
  const resetScheduleOnly=()=>{
    if(!confirm("ลบเฉพาะข้อมูลตารางสอน (ข้อมูลครู/วิชา/ห้องยังอยู่)?"))return;
    U.setSchedule({});U.setLocks({});st("ล้างตารางสอนแล้ว","warning");
  };
  const handleLogo=(e)=>{const f=e.target.files?.[0];if(!f)return;const reader=new FileReader();reader.onload=ev=>{setSH(p=>({...p,logo:ev.target.result}));st("อัพโหลดโลโก้สำเร็จ")};reader.readAsDataURL(f);e.target.value=""};

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(400px,1fr))",gap:24}}>
      {/* Academic Year */}
      <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <h3 style={{fontSize:16,fontWeight:700,marginBottom:20}}>ปีการศึกษา</h3>
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          <div><label style={LS}>ปีการศึกษา (พ.ศ.)</label><input style={IS} value={ay.year} onChange={e=>setAY(p=>({...p,year:e.target.value}))} placeholder="2568"/></div>
          <div><label style={LS}>ภาคเรียนที่</label><select style={IS} value={ay.semester} onChange={e=>setAY(p=>({...p,semester:e.target.value}))}><option value="1">1</option><option value="2">2</option></select></div>
        </div>
      </div>

      {/* School Header + Logo */}
      <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <h3 style={{fontSize:16,fontWeight:700,marginBottom:20}}>หัวเอกสาร (สำหรับ PDF)</h3>
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          <div><label style={LS}>ชื่อโรงเรียน</label><input style={IS} value={sh.name} onChange={e=>setSH(p=>({...p,name:e.target.value}))} placeholder="โรงเรียนดาราวิทยาลัย"/></div>
          <div>
            <label style={LS}>โลโก้โรงเรียน (จะแสดงในตาราง PDF)</label>
            <div style={{display:"flex",alignItems:"center",gap:14,marginTop:8}}>
              {sh.logo
                ?<img src={sh.logo} alt="logo" style={{width:56,height:56,borderRadius:"50%",objectFit:"cover",border:"2px solid #E5E7EB"}}/>
                :<div style={{width:56,height:56,borderRadius:"50%",background:"#F3F4F6",border:"2px dashed #D1D5DB",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:"#9CA3AF"}}>LOGO</div>
              }
              <div style={{display:"flex",gap:8,flexDirection:"column"}}>
                <button onClick={()=>logoRef.current?.click()} style={BS("#2563EB")}><Icon name="upload" size={14}/>อัพโหลดโลโก้</button>
                {sh.logo&&<button onClick={()=>{setSH(p=>({...p,logo:""}));st("ลบโลโก้แล้ว","warning")}} style={BO("#DC2626")}><Icon name="trash" size={14}/>ลบโลโก้</button>}
              </div>
              <input ref={logoRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleLogo}/>
            </div>
            <div style={{fontSize:12,color:"#6B7280",marginTop:8}}>รองรับ PNG, JPG, SVG ขนาดแนะนำ 200×200px ขึ้นไป</div>
          </div>
        </div>
      </div>

      {/* Reset */}
      <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <h3 style={{fontSize:16,fontWeight:700,marginBottom:20,color:"#DC2626"}}>รีเซ็ทข้อมูล</h3>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <button onClick={resetScheduleOnly} style={BO("#D97706")}><Icon name="trash" size={16}/>ล้างเฉพาะตารางสอน</button>
          <div style={{fontSize:12,color:"#6B7280"}}>ลบข้อมูลตารางสอนที่จัดไว้ แต่ข้อมูลครู วิชา ห้อง ยังอยู่</div>
          <div style={{borderTop:"1px solid #E5E7EB",paddingTop:12,marginTop:4}}/>
          <button onClick={resetAll} style={BS("#DC2626")}><Icon name="trash" size={16}/>รีเซ็ทข้อมูลทั้งหมด</button>
          <div style={{fontSize:12,color:"#DC2626"}}>⚠️ ลบข้อมูลทุกอย่าง — ไม่สามารถกู้คืนได้</div>
        </div>
      </div>

      {/* Summary */}
      <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <h3 style={{fontSize:16,fontWeight:700,marginBottom:20}}>สรุปข้อมูลในระบบ</h3>
        <div style={{display:"flex",flexDirection:"column",gap:8,fontSize:14}}>
          <div>ระดับชั้น: <b>{S.levels.length}</b></div>
          <div>แผนการเรียน: <b>{S.plans.length}</b></div>
          <div>กลุ่มสาระ: <b>{S.depts.length}</b></div>
          <div>ครู: <b>{S.teachers.length}</b></div>
          <div>วิชา: <b>{S.subjects.length}</b></div>
          <div>ห้องเรียน: <b>{S.rooms.length}</b></div>
          <div>คาบที่จัดแล้ว: <b>{Object.values(S.schedule).reduce((s,en)=>s+(en?.length||0),0)}</b></div>
          <div>คาบที่ล็อค: <b>{Object.values(S.locks).filter(Boolean).length}</b></div>
        </div>
      </div>
    </div>
  </div>;
}


/* ===== PDF HELPER — ตามแบบฟอร์มดาราวิทยาลัย (A4 แนวตั้ง) ===== */
function pdfPage(title, subtitle, dayRows, footerText, logoBase64) {
  const PLIST = [
    { id: 1, time: "08.30-09.20" }, { id: 2, time: "09.20-10.10" },
    { id: 3, time: "10.25-11.15" }, { id: 4, time: "11.15-12.05" },
    { id: 5, time: "13.00-13.50" }, { id: 6, time: "14.00-14.50" },
    { id: 7, time: "14.50-15.40" },
  ];

  const thNums = PLIST.map(p => '<th class="period-num">' + p.id + '</th>').join("");
  const thTimes = PLIST.map(p => '<th class="period-time">' + p.time + '</th>').join("");

  const bodyRows = dayRows.map(function(r) {
    const dayCells = r.cells.map(function(entries) {
      if (!entries || !entries.length) return '<td class="slot"></td>';
      const inner = entries.map(function(e) {
        let h = '<div class="ent"><div class="ent-sub">' + e.sub + '</div><div class="ent-room">' + e.room + '</div>';
        if (e.room2) h += '<div class="ent-room2">' + e.room2 + '</div>';
        h += '</div>';
        return h;
      }).join("");
      return '<td class="slot">' + inner + '</td>';
    }).join("");
    return '<tr><td class="day-cell">' + r.day + '</td>' + dayCells + '</tr>';
  }).join("\n");

  const logoHtml = logoBase64
    ? '<img src="' + logoBase64 + '" style="width:48px;height:48px;border-radius:50%;object-fit:cover;flex-shrink:0"/>'
    : '<div class="logo">LOGO</div>';

  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    "@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');" +
    '@page{size:A4 portrait;margin:10mm 8mm}' +
    '*{margin:0;padding:0;box-sizing:border-box}' +
    "body{font-family:'Sarabun','Noto Sans Thai',sans-serif;font-size:11px;color:#000}" +
    '.page{width:100%;position:relative}' +
    '.header-row{display:flex;align-items:center;margin-bottom:6px;gap:12px}' +
    '.logo{width:48px;height:48px;border:1.5px solid #999;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:8px;color:#666;flex-shrink:0}' +
    '.title-block{flex:1}' +
    '.title-main{font-size:14px;font-weight:700}' +
    '.title-sub{font-size:11px;color:#444;margin-top:2px}' +
    'table{width:100%;border-collapse:collapse;table-layout:fixed;margin-top:4px}' +
    'th,td{border:1px solid #000;text-align:center;vertical-align:middle}' +
    'th{padding:3px 1px;font-weight:700}' +
    'th.period-num{font-size:13px;height:24px}' +
    'th.period-time{font-size:9px;height:18px;font-weight:400}' +
    'th.day-col{width:52px;font-size:11px;font-weight:700}' +
    'td.day-cell{font-weight:700;font-size:12px;padding:4px 2px;width:52px}' +
    'td.slot{padding:3px 2px;vertical-align:top;height:68px}' +
    '.ent{margin-bottom:2px}' +
    '.ent-sub{font-weight:700;font-size:11px;line-height:1.3}' +
    '.ent-room{font-size:10px;color:#111;line-height:1.25}' +
    '.ent-room2{font-size:9px;color:#333;line-height:1.2}' +
    '.sig-area{margin-top:16px;font-size:11px}' +
    '.sig-flex{display:flex;justify-content:space-between;padding:0 20px}' +
    '.sig-box{text-align:center}' +
    '.sig-line{display:inline-block;width:160px;border-bottom:1px dotted #000;margin-bottom:3px}' +
    '@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}' +
    '</style></head><body>' +
    '<div class="page">' +
    '<div class="header-row">' +
    logoHtml +
    '<div class="title-block"><div class="title-main">' + title + '</div><div class="title-sub">' + subtitle + '</div></div>' +
    '</div>' +
    '<table><thead>' +
    '<tr><th class="day-col" rowspan="2">วัน<br/><span style="font-size:9px;font-weight:400">คาบ/เวลา</span></th>' + thNums + '</tr>' +
    '<tr>' + thTimes + '</tr>' +
    '</thead><tbody>' +
    bodyRows +
    '</tbody></table>' +
    '<div class="sig-area"><div class="sig-flex">' +
    '<div class="sig-box">ลงชื่อ<div class="sig-line"></div><br/>รองฯฝ่ายวิชาการ</div>' +
    '<div class="sig-box">ลงชื่อ<div class="sig-line"></div><br/>ผู้อำนวยการ</div>' +
    '</div></div>' +
    '</div></body></html>';
}

/* ===== PDF: พิมพ์หลายตาราง 2 ต่อ 1 หน้า A4 แนวตั้ง ===== */
function pdfMultiPage(pages, logoBase64) {
  const PLIST = [
    { id: 1, time: "08.30-09.20" }, { id: 2, time: "09.20-10.10" },
    { id: 3, time: "10.25-11.15" }, { id: 4, time: "11.15-12.05" },
    { id: 5, time: "13.00-13.50" }, { id: 6, time: "14.00-14.50" },
    { id: 7, time: "14.50-15.40" },
  ];
  const thNums = PLIST.map(p => '<th class="period-num">' + p.id + '</th>').join("");
  const thTimes = PLIST.map(p => '<th class="period-time">' + p.time + '</th>').join("");

  const logoHtml = logoBase64
    ? '<img src="' + logoBase64 + '" style="width:36px;height:36px;border-radius:50%;object-fit:cover;flex-shrink:0"/>'
    : '<div class="logo">LOGO</div>';

  const buildBlock = (pg) => {
    const bodyRows = pg.dayRows.map(function(r) {
      const dayCells = r.cells.map(function(entries) {
        if (!entries || !entries.length) return '<td class="slot"></td>';
        const inner = entries.map(function(e) {
          let h = '<div class="ent"><div class="ent-sub">' + e.sub + '</div><div class="ent-room">' + e.room + '</div>';
          if (e.room2) h += '<div class="ent-room2">' + e.room2 + '</div>';
          h += '</div>';
          return h;
        }).join("");
        return '<td class="slot">' + inner + '</td>';
      }).join("");
      return '<tr><td class="day-cell">' + r.day + '</td>' + dayCells + '</tr>';
    }).join("\n");

    return '<div class="block">' +
      '<div class="header-row">' + logoHtml +
      '<div class="title-block"><div class="title-main">' + pg.title + '</div><div class="title-sub">' + pg.subtitle + '</div></div>' +
      '</div>' +
      '<table><thead>' +
      '<tr><th class="day-col" rowspan="2">วัน<br/><span style="font-size:8px;font-weight:400">คาบ/เวลา</span></th>' + thNums + '</tr>' +
      '<tr>' + thTimes + '</tr>' +
      '</thead><tbody>' + bodyRows + '</tbody></table>' +
      '<div class="sig-area"><div class="sig-flex">' +
      '<div class="sig-box">ลงชื่อ<div class="sig-line"></div><br/>รองฯฝ่ายวิชาการ</div>' +
      '<div class="sig-box">ลงชื่อ<div class="sig-line"></div><br/>ผู้อำนวยการ</div>' +
      '</div></div>' +
      '</div>';
  };

  let pagesHtml = "";
  for (let i = 0; i < pages.length; i += 2) {
    const a = pages[i];
    const b = pages[i + 1];
    pagesHtml += '<div class="sheet">' + buildBlock(a) + (b ? '<hr class="divider"/>' + buildBlock(b) : '') + '</div>';
  }

  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    "@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');" +
    '@page{size:A4 portrait;margin:8mm 7mm}' +
    '*{margin:0;padding:0;box-sizing:border-box}' +
    "body{font-family:'Sarabun','Noto Sans Thai',sans-serif;font-size:10px;color:#000}" +
    '.sheet{page-break-after:always}' +
    '.sheet:last-child{page-break-after:avoid}' +
    '.block{}' +
    '.header-row{display:flex;align-items:center;margin-bottom:4px;gap:8px}' +
    '.logo{width:36px;height:36px;border:1px solid #999;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:7px;color:#666;flex-shrink:0}' +
    '.title-block{flex:1}' +
    '.title-main{font-size:12px;font-weight:700}' +
    '.title-sub{font-size:10px;color:#444;margin-top:1px}' +
    'table{width:100%;border-collapse:collapse;table-layout:fixed;margin-top:3px}' +
    'th,td{border:1px solid #000;text-align:center;vertical-align:middle}' +
    'th{padding:2px 1px;font-weight:700}' +
    'th.period-num{font-size:12px;height:20px}' +
    'th.period-time{font-size:8px;height:15px;font-weight:400}' +
    'th.day-col{width:46px;font-size:10px;font-weight:700}' +
    'td.day-cell{font-weight:700;font-size:11px;padding:2px;width:46px}' +
    'td.slot{padding:2px 1px;vertical-align:top;height:56px}' +
    '.ent{margin-bottom:1px}' +
    '.ent-sub{font-weight:700;font-size:10px;line-height:1.25}' +
    '.ent-room{font-size:9px;color:#111;line-height:1.2}' +
    '.ent-room2{font-size:8px;color:#333;line-height:1.15}' +
    '.sig-area{margin-top:4px;font-size:9px}' +
    '.sig-flex{display:flex;justify-content:space-between;padding:0 15px}' +
    '.sig-box{text-align:center}' +
    '.sig-line{display:inline-block;width:130px;border-bottom:1px dotted #000;margin-bottom:2px}' +
    '.divider{border:none;border-top:1.5px dashed #aaa;margin:5px 0}' +
    '@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}' +
    '</style></head><body>' +
    pagesHtml +
    '</body></html>';
}
