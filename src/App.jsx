import {schoolLogo} from './renovation/school-logo.mjs';
import AssignmentHeader from './renovation/AssignmentHeader.jsx';
import {uiConfirm,uiPrompt,uiAlert} from './renovation/AppQuestion.jsx';
import UnifiedDialog from './renovation/UnifiedDialog.jsx';
import TeacherMini from './renovation/TeacherMini.jsx';
import SwapWorkbench from './renovation/SwapWorkbench.jsx';
import PrintStudio from './renovation/PrintStudio.jsx';
import ColorBadge from './renovation/ColorBadge.jsx';
import {departmentTone,levelTone} from './renovation/colors.mjs';
import {ManagementToolbar,RecordList,LevelsManager,FileActions,EditDialog} from "./renovation/Management.jsx";
import ReportLauncher from "./renovation/ReportLauncher.jsx";
import TeacherDirectory from "./renovation/TeacherDirectory.jsx";
import {Workspace, Dashboard, PageHeading, Glyph} from "./renovation/Workspace.jsx";
import {progress} from "./renovation/model.mjs";
import {readLiveConfig,schoolAccount,effectivePermissions,canonical} from './renovation/live-config.mjs';
import AccessAdmin from './renovation/AccessAdmin.jsx';
const LIVE=readLiveConfig(import.meta.env);
const PREVIEW_MODE=LIVE.preview;
import { useState, useCallback, useEffect, useRef, useMemo } from "react";
import * as XLSX from 'xlsx';
import { initializeApp } from "firebase/app";
import { getAuth, GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged } from "firebase/auth";
import { initializeFirestore, runTransaction, doc, getDoc, setDoc, collection, getDocs, onSnapshot } from "firebase/firestore";

// ===== FIREBASE CONFIG — ใส่ค่าจาก Firebase Console =====
const FIREBASE_CONFIG = LIVE.firebase;
// =========================================================



// Firebase instances (lazy init เพื่อกัน crash ถ้ายังไม่ได้ตั้งค่า)
let _fbApp=null, _auth=null, _db=null;
const getFB=()=>{
  if(!LIVE.ready) return {auth:null,db:null};
  if(!_fbApp&&!FIREBASE_CONFIG.apiKey.includes("YOUR")){
    _fbApp=initializeApp(FIREBASE_CONFIG);
    _auth=getAuth(_fbApp);
    // autoDetectLongPolling แก้ปัญหา WebChannel 400 error บน GitHub Pages
    _db=initializeFirestore(_fbApp,{
      experimentalAutoDetectLongPolling:true,

    });
    // Memory cache only: live data is not retained between shared-computer sessions.
  }
  return{auth:_auth,db:_db};
};

// Firestore helpers
const fsGetPermissions=async(uid)=>{
  const {db}=getFB();if(!db)return null;
  const snap=await getDoc(doc(db,"permissions",uid));
  return snap.exists()?snap.data():null;
};
const fsSetPermissions=async(uid,data)=>{
  const {db}=getFB();if(!db)return;
  await setDoc(doc(db,"permissions",uid),data,{merge:true});
};

// ===== FIRESTORE TIMETABLE HELPERS (Realtime) =====
const DATA_FIELDS = ["levels","plans","depts","teachers","subjects","rooms","specialRooms","assigns","meetings","schedule","locks"];
// ตรวจ environment: localhost = dev, github.io = production
const IS_DEV = typeof window!=="undefined" && (window.location.hostname==="localhost"||window.location.hostname==="127.0.0.1");
const FS_COLLECTION = LIVE.collection;

// Save ข้อมูลทั้งหมดไป Firestore (merge เพื่อไม่ทับ _init)
const fsSaveTimetable = async (divId, data, expected) => {
  const {db} = getFB(); if(!db) return;
  const payload = {};
  DATA_FIELDS.forEach(f => { if(data[f] !== undefined) payload[f] = data[f]; });
  if(data.schoolHeader) payload.schoolHeader = data.schoolHeader;
  if(data.academicYear) payload.academicYear = data.academicYear;
  // ใช้ setDoc ไม่ merge เพื่อให้ schedule ถูก replace ทั้งก้อน (กัน entries เก่าค้าง)
  const target=doc(db,FS_COLLECTION,divId);
  await runTransaction(db,async tx=>{
    const snap=await tx.get(target);
    if(canonical(snap.exists()?snap.data():{})!==expected) throw Error('ตารางถูกแก้ไขจากเครื่องอื่น กรุณาโหลดข้อมูลล่าสุดก่อนบันทึก');
    tx.set(target,payload);
  });
  return canonical(payload);
};

// Subscribe realtime — returns unsubscribe function
const fsSubscribeTimetable = (divId, onData, onError) => {
  const {db} = getFB(); if(!db) return ()=>{};
  return onSnapshot(doc(db,FS_COLLECTION,divId), {includeMetadataChanges:true}, (snap) => {
    // ถ้า document ไม่มี → ส่ง {} เพื่อให้ระบบ init state ว่างได้ (ไม่ค้าง syncing)
    if(!snap.metadata.fromCache && !snap.metadata.hasPendingWrites) onData(snap.exists() ? snap.data() : {});
  }, (err) => { onError(err); });
};

// ===== LOGIN SCREEN =====
function LoginScreen({onLogin}){
  const [loading,setLoading]=useState(false);
  const [err,setErr]=useState("");

  const handleGoogle=async()=>{
    const {auth}=getFB();
    if(!auth){setErr("ยังไม่ได้ตั้งค่าการเชื่อมต่อระบบ");return;}
    setLoading(true);setErr("");
    try{
      const provider=new GoogleAuthProvider();
      provider.setCustomParameters({hd:"web1.dara.ac.th"}); // จำกัดเฉพาะ domain โรงเรียน
      const result=await signInWithPopup(auth,provider);
      onLogin(result.user);
    } catch(e){
      setErr(e.code==="auth/popup-closed-by-user"?"ปิด popup ก่อนเลือกบัญชี":e.message);
    }
    setLoading(false);
  };

  return(
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#991B1B,#7F1D1D)"}}>
      <div data-ui-surface="true" style={{background:"#fff",borderRadius:20,padding:"48px 40px",width:400,textAlign:"center",boxShadow:"0 25px 60px rgba(0,0,0,0.3)"}}>
        <div style={{fontSize:48,marginBottom:16}}>📋</div>
        <h1 style={{fontSize:22,fontWeight:700,marginBottom:4}}>ระบบจัดตารางสอน</h1>
        <p style={{color:"#6B7280",fontSize:13,marginBottom:32}}>โรงเรียนดาราวิทยาลัย</p>
        <button data-ui-control="true"
          onClick={handleGoogle}
          disabled={loading}
          style={{width:"100%",padding:"13px 0",background:loading?"#F3F4F6":"#fff",border:"1.5px solid #D1D5DB",borderRadius:12,fontSize:14,fontWeight:600,cursor:loading?"not-allowed":"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:10,marginBottom:16}}
        >
          <svg width="20" height="20" viewBox="0 0 48 48"><path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/><path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/><path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/><path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.18 1.48-4.97 2.35-8.16 2.35-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/></svg>
          {loading?"กำลังเข้าสู่ระบบ...":"เข้าสู่ระบบด้วย Google โรงเรียน"}
        </button>
        {err&&<div style={{padding:10,background:"#FEE2E2",borderRadius:8,color:"#991B1B",fontSize:12,marginBottom:8}}>{err}</div>}
        <p style={{color:"#9CA3AF",fontSize:11}}>ใช้บัญชี @web1.dara.ac.th เท่านั้น</p>
      </div>
    </div>
  );
}

// ===== ADMIN PANEL =====
const DAYS = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์"];
const PERIODS = [
  { id: 1, time: "08.30-09.20" }, { id: 2, time: "09.20-10.10" },
  { id: 3, time: "10.25-11.15" }, { id: 4, time: "11.15-12.05" },
  { id: 5, time: "13.00-13.50" }, { id: 6, time: "14.00-14.50" },
  { id: 7, time: "14.50-15.40" },
];
// ชื่อวิชาย่อ: ใช้ shortName ถ้ามี ไม่งั้นใช้ name เต็ม
const subDisplayName = (sub) => sub?.shortName||sub?.name||"";


// ===== Design tokens (Dara red scheme) =====
const CRED="#9C2638";      // แดงดารา หลัก
const CBGW="#FFFFFF";       // white card
const IS={width:"100%",padding:"10px 14px",border:"1.5px solid #E5E7EB",borderRadius:12,fontSize:14,outline:"none",fontFamily:"inherit",boxSizing:"border-box",background:"#fff",color:"#1A1A1A"};
const BS=(c=CRED)=>({padding:'9px 15px',minHeight:40,background:c===CRED?'#9C2638':'#FFFFFF',color:c===CRED?'#FFFFFF':'#526174',border:c===CRED?'1px solid #9C2638':'1px solid #DDE3EB',borderRadius:8,fontSize:14,fontWeight:600,cursor:'pointer',display:'inline-flex',alignItems:'center',gap:7,fontFamily:'inherit'});
const BO=(c=CRED)=>({padding:'9px 15px',minHeight:40,background:'#FFFFFF',color:'#526174',border:'1px solid #DDE3EB',borderRadius:8,fontSize:14,fontWeight:500,cursor:'pointer',display:'inline-flex',alignItems:'center',gap:7,fontFamily:'inherit'});
const LS={display:"block",fontSize:13,fontWeight:600,color:"#374151",marginBottom:6};

// ===== SearchSelect — Searchable Dropdown =====
function SearchSelect({value, onChange, options, placeholder="-- เลือก --", style={}, disabled=false}){
  const [open,setOpen]=useState(false);
  const [q,setQ]=useState("");
  const ref=useRef(null);
  const inputRef=useRef(null);
  const selected=options.find(o=>o.value===value);

  // ปิด dropdown เมื่อคลิกนอก
  useEffect(()=>{
    const handler=(e)=>{
      if(ref.current&&!ref.current.contains(e.target)){
        setOpen(false);
        setQ("");
      }
    };
    document.addEventListener("mousedown",handler);
    return()=>document.removeEventListener("mousedown",handler);
  },[]);

  const filtered=q.trim()
    ?options.filter(o=>o.label.toLowerCase().includes(q.toLowerCase()))
    :options;

  const displayText = open ? q : (selected ? selected.label : "");

  return(
    <div ref={ref} style={{position:"relative",width:"100%",...style}}>
      {/* Input เป็น trigger หลัก — คลิกแล้วพิมพ์ได้เลย */}
      <div style={{position:"relative"}}>
        <input data-ui-control="true"
          ref={inputRef}
          value={displayText}
          readOnly={disabled}
          placeholder={open ? "พิมพ์เพื่อค้นหา..." : placeholder}
          onClick={()=>{ if(!disabled){ setOpen(true); setQ(""); } }}
          onChange={e=>{ setOpen(true); setQ(e.target.value); }}
          onKeyDown={e=>{
            if(e.key==="Enter"&&filtered.length>0){ onChange(filtered[0].value); setOpen(false); setQ(""); inputRef.current?.blur(); }
            if(e.key==="Escape"){ setOpen(false); setQ(""); inputRef.current?.blur(); }
            if(e.key==="ArrowDown"){ setOpen(true); }
          }}
          style={{
            ...IS,
            cursor:disabled?"default":"text",
            background:disabled?"#F3F4F6":open?"#fff":"#F9FAFB",
            paddingRight:36,
            color: open ? "#111" : (selected ? "#111" : "#9CA3AF"),
            borderColor: open ? "#991B1B" : undefined,
          }}
        />
        <span
          style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",color:"#9CA3AF",fontSize:10,pointerEvents:"none",userSelect:"none"}}>
          {open?"▲":"▼"}
        </span>
      </div>

      {/* Dropdown list */}
      {open&&!disabled&&(
        <div data-ui-surface="true"
          onMouseDown={e=>e.preventDefault()} // ป้องกัน input blur เมื่อคลิกใน list
          style={{
            position:"absolute",top:"calc(100% + 2px)",left:0,right:0,
            background:"#fff",border:"1.5px solid #E5E7EB",borderRadius:10,
            boxShadow:"0 8px 24px rgba(0,0,0,0.13)",zIndex:9999,
            maxHeight:260,display:"flex",flexDirection:"column",overflow:"hidden",
          }}>
          <div style={{overflowY:"auto",maxHeight:260}}>
            {filtered.length===0
              ?<div style={{padding:"10px 12px",color:"#9CA3AF",fontSize:13}}>ไม่พบผลลัพธ์</div>
              :filtered.map(o=>(
                <div key={o.value}
                  onMouseDown={e=>{
                    e.preventDefault();
                    onChange(o.value);
                    setOpen(false);
                    setQ("");
                  }}
                  style={{
                    padding:"9px 12px",cursor:"pointer",fontSize:13,
                    background:o.value===value?"#FEF2F2":"transparent",
                    color:o.value===value?CRED:"#111",
                    fontWeight:o.value===value?700:400,
                  }}
                  onMouseEnter={e=>e.currentTarget.style.background=o.value===value?"#FEF2F2":"#F9FAFB"}
                  onMouseLeave={e=>e.currentTarget.style.background=o.value===value?"#FEF2F2":"transparent"}
                >{o.label}</div>
              ))
            }
          </div>
        </div>
      )}
    </div>
  );
}

const DC = [
  { bg:"#DC2626",lt:"#FEE2E2",tx:"#991B1B",bd:"#FECACA" }, // แดง
  { bg:"#2563EB",lt:"#DBEAFE",tx:"#1E40AF",bd:"#BFDBFE" }, // น้ำเงิน
  { bg:"#059669",lt:"#D1FAE5",tx:"#065F46",bd:"#A7F3D0" }, // เขียว
  { bg:"#D97706",lt:"#FEF3C7",tx:"#92400E",bd:"#FDE68A" }, // เหลืองส้ม
  { bg:"#7C3AED",lt:"#EDE9FE",tx:"#5B21B6",bd:"#DDD6FE" }, // ม่วง
  { bg:"#DB2777",lt:"#FCE7F3",tx:"#9D174D",bd:"#FBCFE8" }, // ชมพู
  { bg:"#0E7490",lt:"#CFFAFE",tx:"#164E63",bd:"#A5F3FC" }, // ฟ้าเข้ม
  { bg:"#4D7C0F",lt:"#ECFCCB",tx:"#1A2E05",bd:"#BEF264" }, // เขียวเข้ม
  { bg:"#C2410C",lt:"#FFEDD5",tx:"#7C2D12",bd:"#FDBA74" }, // ส้มเข้ม
  { bg:"#0F766E",lt:"#CCFBF1",tx:"#134E4A",bd:"#5EEAD4" }, // เขียวน้ำทะเล
  { bg:"#6D28D9",lt:"#F5F3FF",tx:"#4C1D95",bd:"#C4B5FD" }, // ม่วงเข้ม
  { bg:"#B45309",lt:"#FEF3C7",tx:"#78350F",bd:"#FCD34D" }, // น้ำตาลทอง
];
const SROLES = [
  { id:"academic",name:"ฝ่ายวิชาการ",blocked:[{day:"พฤหัสบดี",periods:[5,6,7]}] },
  { id:"discipline",name:"ฝ่ายพัฒนาวินัย",blocked:[{day:"ศุกร์",periods:[5,6,7]}] },
];
const gid = () => Math.random().toString(36).substr(2,9);

// ===== GAS BACKUP URL (ไม่ได้ใช้แล้ว — ย้ายมา Firestore) =====
// const GAS_URL = "https://script.google.com/macros/s/AKfycbwWym1QWA-...";

// ===== LOCAL STORAGE HELPERS (ใช้เป็น offline cache) =====
const saveLS = (key, data) => { try { if(PREVIEW_MODE) localStorage.setItem(`dara_preview_${key}`, JSON.stringify(data)); } catch(e) {} };
const loadLS = (key, fb) => { if(!PREVIEW_MODE)return fb; try { const d = localStorage.getItem(`dara_preview_${key}`); return d ? JSON.parse(d) : fb; } catch(e) { return fb; } };

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

const Modal = UnifiedDialog;

// ===== TOAST NOTIFICATION =====
const Toast=({message,type="success",onClose})=>{useEffect(()=>{const t=setTimeout(onClose,3000);return()=>clearTimeout(t)},[onClose]);return<div style={{position:"fixed",top:24,right:24,zIndex:9999,background:type==="error"?"#DC2626":type==="warning"?"#D97706":"#059669",color:"#fff",padding:"14px 24px",borderRadius:12,fontSize:14,fontWeight:600,boxShadow:"0 10px 30px rgba(0,0,0,0.2)",display:"flex",alignItems:"center",gap:8,animation:"slideIn 0.3s ease"}}><Icon name={type==="error"?"alert":"check"} size={16}/>{message}</div>};

const DIVISIONS=[
  {id:"p1",name:"ประถมศึกษาตอนต้น",short:"ประถมต้น",defaultLevels:["ป.1","ป.2","ป.3"]},
  {id:"p2",name:"ประถมศึกษาตอนปลาย",short:"ประถมปลาย",defaultLevels:["ป.4","ป.5","ป.6"]},
  {id:"m1",name:"มัธยมศึกษาตอนต้น",short:"มัธยมต้น",defaultLevels:["ม.1","ม.2","ม.3"]},
  {id:"m2",name:"มัธยมศึกษาตอนปลาย",short:"มัธยมปลาย",defaultLevels:["ม.4","ม.5","ม.6"]},
];

// ===== PERIOD CONFIG ตามระดับ =====
// คาบ 1-5 เหมือนกันทุกระดับ
// ประถมต้น (p1): คาบ6=13.50-14.40, พักหลังคาบ6 (14.40-14.50), คาบ7=14.50-15.40
// ระดับอื่น:     พักหลังคาบ5 (13.50-14.00), คาบ6=14.00-14.50, คาบ7=14.50-15.40
const PERIOD_BASE=[
  {id:1,time:"08.30-09.20"},{id:2,time:"09.20-10.10"},
  {id:3,time:"10.25-11.15"},{id:4,time:"11.15-12.05"},
  {id:5,time:"13.00-13.50"},
];
// break ก่อนคาบ = [{afterPeriod, label, key}]
// afterPeriod: หลังคาบไหน / key: "brk0"=08.00-08.30, "brk1"=10.10-10.25, "brk2"=12.05-13.00
const PERIOD_CONFIG={
  // ประถมต้น: ไม่มีพักหลัง p5, คาบ6=13.50-14.40, พักหลังคาบ6, คาบ7=14.50-15.40
  p1:{
    periods:[
      ...PERIOD_BASE,
      {id:6,time:"13.50-14.40"},
      {id:7,time:"14.50-15.40"},
    ],
    // break columns: [ก่อนคาบ1, ก่อนคาบ3, ก่อนคาบ5, ก่อนคาบ7]
    breaks:[
      {label:"08.00-08.30",afterPid:0},   // ก่อนคาบ 1
      {label:"10.10-10.25",afterPid:2},   // หลังคาบ 2
      {label:"12.05-13.00",afterPid:4},   // หลังคาบ 4
      {label:"14.40-14.50",afterPid:6},   // หลังคาบ 6 ← ต่างจากระดับอื่น
    ],
  },
  // ระดับอื่น (p2, m1, m2): พักหลังคาบ5, คาบ6=14.00-14.50, คาบ7=14.50-15.40
  default:{
    periods:[
      ...PERIOD_BASE,
      {id:6,time:"14.00-14.50"},
      {id:7,time:"14.50-15.40"},
    ],
    breaks:[
      {label:"08.00-08.30",afterPid:0},   // ก่อนคาบ 1
      {label:"10.10-10.25",afterPid:2},   // หลังคาบ 2
      {label:"12.05-13.00",afterPid:4},   // หลังคาบ 4
      {label:"13.50-14.00",afterPid:5},   // หลังคาบ 5 ← ต่างจาก p1
    ],
  },
};
// helper: ได้ config ตาม divisionId
function getPeriodCfg(divisionId){
  return PERIOD_CONFIG[divisionId]||PERIOD_CONFIG.default;
}
// helper: ได้ divisionId จาก levelId
// ใช้ level.divisionId ที่บันทึกไว้ (migrate อัตโนมัติเมื่อเปิด LevelsPage)
// fallback: guess จากชื่อ level
// ── single source of truth: guess division จากชื่อ level ──
function guessDivisionFromName(name){
  const n=(name||"").trim();
  for(const div of DIVISIONS){
    if((div.defaultLevels||[]).some(dl=>n===dl||n.startsWith(dl+"/")||n.startsWith(dl+" "))) return div.id;
  }
  if(/ป\.?\s*[1-3]\b/.test(n)) return "p1";
  if(/ป\.?\s*[4-6]\b/.test(n)) return "p2";
  if(/ม\.?\s*[1-3]\b/.test(n)) return "m1";
  if(/ม\.?\s*[4-6]\b/.test(n)) return "m2";
  if(n.includes("ประถมต้น")||n.includes("ป.ต้น")) return "p1";
  if(n.includes("ประถมปลาย")||n.includes("ป.ปลาย")) return "p2";
  if(n.includes("มัธยมต้น")||n.includes("ม.ต้น")) return "m1";
  if(n.includes("มัธยมปลาย")||n.includes("ม.ปลาย")) return "m2";
  return "m2";
}
function getDivisionForLevel(levelId, levels){
  const lv = levels?.find(l => l.id === levelId);
  if(!lv) return "m2";
  // ใช้ค่าที่บันทึกไว้ก่อน ถ้าไม่มีค่อย guess จากชื่อ
  return lv.divisionId || guessDivisionFromName(lv.name);
}
// helper: ได้ divisionId จาก roomId ผ่าน S.rooms+S.levels
function getDivisionForRoom(room,S){
  return getDivisionForLevel(room?.levelId,S?.levels);
}


// helper: หา divisionId หลักของครูจาก rooms ที่สอน
function getDivisionForTeacher(teacherId, S){
  for(const day of ["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"]){
    for(const pid of [1,2,3,4,5,6,7]){
      const entry=Object.entries(S.schedule||{}).find(([k,en])=>{
        if(!k.endsWith("_"+day+"_"+pid))return false;
        return(en||[]).some(e=>{
          const co=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
          return e.teacherId===teacherId||co.includes(teacherId);
        });
      });
      if(entry){
        const rid=entry[0].split("_")[0];
        const rm=S.rooms?.find(r=>r.id===rid);
        if(rm) return getDivisionForLevel(rm.levelId, S.levels);
      }
    }
  }
  return "m2";
}

// ===== MAIN APP COMPONENT =====
function groupEntries(entries) {
  if (!entries || !entries.length) return [];
  const map = {};
  entries.forEach(function(e) {
    const key = e.sub;
    if (!map[key]) {
      map[key] = { sub: e.sub, rooms: [], double: e.double };
    }
    if (e.room && !map[key].rooms.includes(e.room)) map[key].rooms.push(e.room);
    if (e.double) map[key].double = true;
  });
  return Object.values(map).map(function(g) {
    g.rooms.sort(function(a, b) {
      var na = parseInt((a.match(/(\d+)$/) || [0, 9999])[1]);
      var nb = parseInt((b.match(/(\d+)$/) || [0, 9999])[1]);
      return na !== nb ? na - nb : a.localeCompare(b, 'th');
    });
    // แต่ละห้องเป็น div แยกบรรทัด
    var roomHtml = g.rooms.map(function(r){ return '<div class="ent-room">' + r + '</div>'; }).join('');
    var roomHtmlTeacher = g.rooms.map(function(r){ return '<div class="ent-room">ครู' + r + '</div>'; }).join('');
    return { sub: g.sub, rooms: g.rooms, roomHtml: roomHtml, roomHtmlTeacher: roomHtmlTeacher, room2: '', double: g.double, roomCount: g.rooms.length };
  });
}


/* ===== REACT PRINT PREVIEW SYSTEM ===== */

const PLIST_PRINT=[
  {id:1,label:"คาบ 1",time:"08.30-09.20"},
  {id:2,label:"คาบ 2",time:"09.20-10.10"},
  {id:3,label:"คาบ 3",time:"10.25-11.15"},
  {id:4,label:"คาบ 4",time:"11.15-12.05"},
  {id:5,label:"คาบ 5",time:"13.00-13.50"},
  {id:6,label:"คาบ 6",time:"14.00-14.50"},
  {id:7,label:"คาบ 7",time:"14.50-15.40"},
];
const DAYS_PRINT=["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
const SD={"พฤหัสบดี":"พฤหัส"};
const sd=d=>SD[d]||d;

// CSS as object — no template literals, no backticks
function mkPrintStyle(ps){
  const P=ps||DEFAULT_PRINT_SETTINGS;
  const C=PRINT_COLORS[P.color]||PRINT_COLORS["แดง"];
  const f=P.fontSize/100;
  const r=P.rowHeight/100;
  const ff=P.fontFamily||"TH SarabunNew";
  const bdr=P.showBorder?"1px solid "+C.border:"1px solid #E5E7EB";
  const hBdr=P.showBorder?"1px solid "+C.border:"1px solid transparent";
  return [
    "@page{size:A4 portrait;margin:10mm 8mm}",
    // scope ทุก rule ไว้ใน .pt-root เพื่อไม่กระทบ UI ของ preview modal
    ".pt-root{font-family:'"+ff+"','Sarabun','Noto Sans Thai',sans-serif;font-size:"+Math.round(11*f)+"px;color:#000}",
    ".pt-root *{box-sizing:border-box;margin:0;padding:0}",
    ".pt-root table{width:100%;border-collapse:collapse;table-layout:fixed}",
    ".pt-root th{padding:2px 1px;font-weight:700;background:"+C.header+";color:"+C.headerText+";text-align:center;vertical-align:middle;border:"+hBdr+"}",
    ".pt-root td{text-align:center;vertical-align:middle;border:"+bdr+"}",
    P.showAltRow?".pt-root tbody tr:nth-child(even){background:"+C.rowAlt+"}":"",
    ".pt-wrap{width:100%;page-break-inside:avoid}",
    ".pt-hdr{display:flex;align-items:center;gap:10px;margin-bottom:6px}",
    ".pt-logo{width:48px;height:48px;border-radius:50%;object-fit:cover}",
    ".pt-logo-ph{width:48px;height:48px;border:1.5px solid #999;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:8px;color:#666}",
    ".pt-title{font-size:"+Math.round(14*f)+"px;font-weight:700;line-height:1.3}",
    ".pt-sub{font-size:"+Math.round(11*f)+"px;color:#555;margin-top:2px}",
    ".pt-dept{font-size:"+Math.round(10*f)+"px;color:#777;margin-top:1px}",
    ".th-num{font-size:"+Math.round(12*f)+"px;height:"+Math.round(22*r)+"px}",
    ".th-time{font-size:"+Math.round(8*f)+"px;height:"+Math.round(16*r)+"px;font-weight:400;white-space:nowrap}",
    ".td-day{font-weight:700;font-size:"+Math.round(12*f)+"px;background:#F3F4F6;padding:2px;width:50px}",
    ".td-slot{padding:2px;height:"+Math.round(72*r)+"px;vertical-align:middle}",
    ".td-slot-hi{background:#f0f0f0}",
    ".td-brk{background:#fffde7;padding:0;width:26px}",
    ".td-hm{padding:2px;background:#f5fff5;width:30px}",
    ".ent{margin-bottom:2px}",
    ".ent-sub{font-weight:700;font-size:"+Math.round(12*f)+"px;line-height:1.3}",
    ".ent-room{font-size:"+Math.round(11*f)+"px;color:#111;line-height:1.2}",
    ".vt{writing-mode:vertical-rl;transform:rotate(180deg);white-space:nowrap;font-weight:600;font-size:9px;letter-spacing:1px;display:flex;align-items:center;justify-content:center;height:100%}",
    ".sig{margin-top:14px;display:flex;justify-content:space-between;padding:0 20px;font-size:"+Math.round(11*f)+"px}",
    ".sig-box{text-align:center}",
    ".sig-line{display:inline-block;width:150px;border-bottom:1px dotted #000;margin-bottom:3px}",
    "@media print{.pt-root{-webkit-print-color-adjust:exact;print-color-adjust:exact}}",
  ].filter(Boolean).join("\n");
}

// Diagonal corner cell
function CornerTh(){
  return (
    <th rowSpan={2} style={{position:"relative",height:42,width:50,padding:0}}>
      <svg style={{position:"absolute",top:0,left:0,width:"100%",height:"100%"}} preserveAspectRatio="none">
        <line x1="0" y1="0" x2="100%" y2="100%" stroke="#aaa" strokeWidth="0.8"/>
      </svg>
      <span style={{position:"absolute",top:3,right:3,fontSize:"0.75em",fontWeight:700}}>คาบ</span>
      <span style={{position:"absolute",bottom:3,left:3,fontSize:"0.75em",fontWeight:700}}>วัน</span>
    </th>
  );
}

// Slot cell for format 1
function SlotCell1({entries,isRoom}){
  if(!entries||!entries.length)return <td className="td-slot"/>;
  // custom lock — แสดงสีส้มอ่อน
  if(entries[0]?.isCustomLock){
    return (
      <td className="td-slot" style={{background:"#FFF3E0"}}>
        <div className="ent">
          <div className="ent-sub" style={{color:"#E65100",fontSize:"8pt"}}>{entries[0].sub}</div>
        </div>
      </td>
    );
  }
  const grp=groupEntries(entries);
  const hi=grp.some(e=>e.double||e.roomCount>1);
  return (
    <td className={"td-slot"+(hi?" td-slot-hi":"")}>
      {grp.map((e,i)=>(
        <div key={i} className="ent">
          <div className="ent-sub">{e.sub}</div>
          <div className="ent-room">{isRoom?("ครู"+(e.rooms[0]||"")):e.rooms[0]||""}</div>
        </div>
      ))}
    </td>
  );
}

// Signature row
function SigRow(){
  return (
    <div className="sig">
      <div className="sig-box">ลงชื่อ<div className="sig-line"/><br/><span>รองฯฝ่ายวิชาการ</span></div>
      <div className="sig-box">ลงชื่อ<div className="sig-line"/><br/><span>ผู้อำนวยการ</span></div>
    </div>
  );
}

// Logo + title
function TblHdr({title,subtitle,dept,logo}){
  return (
    <div className="pt-hdr">
      {logo?<img src={logo} className="pt-logo" alt=""/>:<div className="pt-logo-ph">L</div>}
      <div>
        <div className="pt-title">{title}</div>
        <div className="pt-sub">{subtitle}</div>
        {dept&&<div className="pt-dept">{dept}</div>}
      </div>
    </div>
  );
}

// FORMAT 1: pdfPage replacement
function PrintF1({pages,logo,subtitle,ps,isRoom}){
  return <div className="pt-root" style={{padding:"10mm 8mm"}}>
    <style>{mkPrintStyle(ps)}</style>
    {pages.map((pg,pi)=>(
      <div key={pi} className="pt-wrap" style={pi>0?{pageBreakBefore:"always"}:{}}>
        <TblHdr title={pg.title} subtitle={subtitle} logo={logo}/>
        <table>
          <thead>
            <tr>
              <CornerTh/>
              {PLIST_PRINT.map(p=><th key={p.id} className="th-num">{p.id}</th>)}
            </tr>
            <tr>
              {PLIST_PRINT.map(p=><th key={p.id} className="th-time">{p.time}</th>)}
            </tr>
          </thead>
          <tbody>
            {DAYS_PRINT.map(day=>(
              <tr key={day}>
                <td className="td-day">{sd(day)}</td>
                {PLIST_PRINT.map(p=>(
                  <SlotCell1 key={p.id} entries={(pg.dayRows.find(r=>r.day===day)||{cells:[]}).cells[p.id-1]||[]} isRoom={isRoom}/>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
        <SigRow/>
      </div>
    ))}
  </div>;
}

// FORMAT 2: Teacher table (1 row per day, with break columns)
function PrintF2({teachers,S,ay,sh,ps}){
  const f=(ps||DEFAULT_PRINT_SETTINGS).fontSize/100;
  const C=PRINT_COLORS[(ps||DEFAULT_PRINT_SETTINGS).color]||PRINT_COLORS["แดง"];
  const rowH=Math.round(52*f);
  const thSt={border:"1px solid "+C.border,background:C.header,color:C.headerText,textAlign:"center",padding:"2px 1px",fontWeight:700};
  const brkSt={border:"1px solid #ddd",background:"#fffde7",padding:0,width:26};
  const pairs=[];for(let i=0;i<teachers.length;i+=2)pairs.push(teachers.slice(i,i+2));
  const NDAYS=DAYS_PRINT.length;
  function getEntries(t,day,pid){
    const out=[];
    Object.entries(S.schedule).forEach(([k,en])=>{
      if(!en?.length)return;const pts=k.split("_");
      if(pts[pts.length-2]!==day||parseInt(pts[pts.length-1])!==pid)return;
      en.forEach(e=>{
        const co=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
        if(e.teacherId!==t.id&&!co.includes(t.id))return;
        const sub=S.subjects.find(s=>s.id===e.subjectId);
        const rid=pts.slice(0,-2).join("_");const rm=S.rooms.find(r=>r.id===rid);
        out.push({sub:sub?.shortName||sub?.name||sub?.code||"—",room:rm?.name||"—"});
      });
    });
    return out;
  }
  function getHM(t,day){
    const m=(S.meetings||[]).find(m=>m.teacherId===t.id&&m.day===day&&(m.isAssembly||m.isHomeroom||(m.periods||[]).includes(0)));
    return m?(m.isAssembly?"หอประชุม":(m.label||"Homeroom")):"Homeroom";
  }
  function SlotTd({arr}){
    if(!arr.length)return <td style={{border:"1px solid #ddd",padding:0}}><div style={{height:rowH}}/></td>;
    return <td style={{border:"1px solid #ddd",padding:0}}>
      <div style={{height:rowH,overflow:"hidden",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",textAlign:"center",padding:"1px"}}>
        {arr.map((e,i)=><div key={i} style={{lineHeight:1.2}}>
          <div style={{fontWeight:700,fontSize:Math.round(8.5*f)+"pt"}}>{e.sub}</div>
          <div style={{fontSize:Math.round(7.5*f)+"pt",color:"#1a237e"}}>{e.room}</div>
        </div>)}
      </div>
    </td>;
  }
  const extraCss=mkPrintStyle(ps)+"\n.vt{writing-mode:vertical-rl;transform:rotate(180deg);white-space:nowrap;font-weight:600;font-size:9px;letter-spacing:1px;display:flex;align-items:center;justify-content:center;height:100%}";
  return <div className="pt-root" style={{padding:"8mm 8mm"}}>
    <style>{extraCss}</style>
    {pairs.map((pair,pi)=>(
      <div key={pi} style={pi>0?{pageBreakBefore:"always"}:{}}>
        {pair.map((t,ti)=>(
          <div key={t.id} style={ti>0?{borderTop:"1px dashed #bbb",marginTop:8,paddingTop:8}:{}}>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
              {sh?.logo?<img src={sh.logo} style={{width:36,height:36,borderRadius:"50%"}} alt=""/>
                       :<div style={{width:36,height:36,border:"1px solid #999",borderRadius:"50%"}}/>}
              <div>
                <div style={{fontSize:Math.round(12*f)+"pt",fontWeight:700}}>ตารางสอน {t.prefix||""}{t.firstName||""} {t.lastName||""}  ปีการศึกษา {ay?.year||"2568"}</div>
                <div style={{fontSize:Math.round(9*f)+"pt",color:"#555"}}>{(S.depts.find(d=>d.id===t.departmentId)||{}).name||""}</div>
              </div>
            </div>
            <table style={{width:"100%",borderCollapse:"collapse",tableLayout:"fixed"}}>
              <colgroup>
                <col style={{width:50}}/><col style={{width:26}}/><col/><col/>
                <col style={{width:26}}/><col/><col/><col style={{width:28}}/><col/><col/>
                <col style={{width:22}}/><col/>
              </colgroup>
              <thead>
                <tr>
                  <th style={{...thSt,position:"relative",height:40,padding:0}} rowSpan={2}>
                    <svg style={{position:"absolute",top:0,left:0,width:"100%",height:"100%"}} preserveAspectRatio="none"><line x1="0" y1="0" x2="100%" y2="100%" stroke="#aaa" strokeWidth="0.8"/></svg>
                    <span style={{position:"absolute",top:3,right:3,fontSize:Math.round(7*f)+"pt",fontWeight:700}}>เวลา</span>
                    <span style={{position:"absolute",bottom:3,left:3,fontSize:Math.round(7*f)+"pt",fontWeight:700}}>วัน</span>
                  </th>
                  <th style={{...brkSt,height:40,verticalAlign:"middle"}} rowSpan={2}><div className="vt">08.00-08.30</div></th>
                  <th style={thSt}>คาบ 1<br/><span style={{fontSize:Math.round(7*f)+"pt",fontWeight:400}}>08.30-09.20</span></th>
                  <th style={thSt}>คาบ 2<br/><span style={{fontSize:Math.round(7*f)+"pt",fontWeight:400}}>09.20-10.10</span></th>
                  <th style={{...brkSt,height:40,verticalAlign:"middle"}} rowSpan={2}><div className="vt">10.10-10.25</div></th>
                  <th style={thSt}>คาบ 3<br/><span style={{fontSize:Math.round(7*f)+"pt",fontWeight:400}}>10.25-11.15</span></th>
                  <th style={thSt}>คาบ 4<br/><span style={{fontSize:Math.round(7*f)+"pt",fontWeight:400}}>11.15-12.05</span></th>
                  <th style={{...brkSt,height:40,verticalAlign:"middle"}} rowSpan={2}><div className="vt">12.05-13.00</div></th>
                  <th style={thSt}>คาบ 5<br/><span style={{fontSize:Math.round(7*f)+"pt",fontWeight:400}}>13.00-13.50</span></th>
                  <th style={thSt}>คาบ 6<br/><span style={{fontSize:Math.round(7*f)+"pt",fontWeight:400}}>14.00-14.50</span></th>
                  <th style={{...brkSt,height:40,verticalAlign:"middle"}} rowSpan={2}><div className="vt">13.50-14.00</div></th>
                  <th style={thSt}>คาบ 7<br/><span style={{fontSize:Math.round(7*f)+"pt",fontWeight:400}}>14.50-15.40</span></th>
                </tr>
              </thead>
              <tbody>
                {DAYS_PRINT.map((day,di)=>{
                  const hm=getHM(t,day);
                  const hmBg=hm.includes("หอประชุม")?"#e8f5e9":"#f5fff5";
                  const slots=PLIST_PRINT.map(p=>getEntries(t,day,p.id));
                  const bg=di%2===1?{background:"#fafafa"}:{};
                  return <tr key={day} style={{height:rowH,...bg}}>
                    <td style={{border:"1px solid #888",fontWeight:700,fontSize:Math.round(9*f)+"pt",background:"#F3F4F6",textAlign:"center",padding:2,verticalAlign:"middle"}}>{sd(day)}</td>
                    <td style={{border:"1px solid #888",background:hmBg,padding:2,textAlign:"center",fontSize:Math.round(7.5*f)+"pt",fontWeight:600,verticalAlign:"middle",lineHeight:1.3}}>{hm}</td>
                    <SlotTd arr={slots[0]}/><SlotTd arr={slots[1]}/>
                    {di===0&&<td rowSpan={NDAYS} style={{background:"#fffde7",padding:0,width:26,border:"1px solid #ddd",verticalAlign:"middle"}}><div className="vt">พักน้อย 15 นาที</div></td>}
                    <SlotTd arr={slots[2]}/><SlotTd arr={slots[3]}/>
                    {di===0&&<td rowSpan={NDAYS} style={{background:"#fffde7",padding:0,width:28,border:"1px solid #ddd",verticalAlign:"middle"}}><div className="vt">พักกลางวัน 55 นาที</div></td>}
                    <SlotTd arr={slots[4]}/><SlotTd arr={slots[5]}/>
                    {di===0&&<td rowSpan={NDAYS} style={{background:"#fffde7",padding:0,width:22,border:"1px solid #ddd",verticalAlign:"middle"}}><div className="vt">พักน้อย 10 นาที</div></td>}
                    <SlotTd arr={slots[6]}/>
                  </tr>;
                })}
              </tbody>
            </table>
          </div>
        ))}
      </div>
    ))}
  </div>;
}

// FORMAT 3: room code + summary (3 per page)
function PrintF3({teachers,S,ay,sh}){
  const yr=ay?.year||"2568";
  const NDAYS=DAYS_PRINT.length;
  const PTIMES=["08.30-09.20","09.20-10.10","10.25-11.15","11.15-12.05","13.00-13.50","14.00-14.50","14.50-15.40"];
  const thS={border:"1px solid #888",fontSize:"7pt",fontWeight:700,textAlign:"center",background:"#f0f0f0",padding:"1px 2px"};
  const bS={border:"1px solid #ddd",background:"#fffde7",padding:0,width:20};
  const css=[
    "@page{size:A4 portrait;margin:8mm 6mm}",
    "*{box-sizing:border-box;margin:0;padding:0}",
    ".pt-root{font-family:'TH SarabunNew','Sarabun',sans-serif;font-size:9px}",
    ".pt-root table{border-collapse:collapse;table-layout:fixed;width:100%}",
    ".pt-root td,.pt-root th{overflow:hidden;text-align:center;vertical-align:middle}",
    ".f3-wrap{padding:2mm 3mm;page-break-inside:avoid;width:100%;box-sizing:border-box;}",
    ".f3-sep{border:none;border-top:1px dashed #bbb;margin:1mm 3mm}",
    ".f3-day{font-weight:700;font-size:9pt;background:#f5f5f5;text-align:center;padding:1px;border:1px solid #888}",
    ".f3-hm{font-size:8pt;font-weight:600;text-align:center;padding:1px;line-height:1.2;border:1px solid #888}",
    ".f3-cell{font-size:8.5pt;font-weight:700;padding:1px;border:1px solid #ddd}",
    ".vt{writing-mode:vertical-rl;transform:rotate(180deg);white-space:nowrap;font-weight:600;font-size:7pt;letter-spacing:1px;display:flex;align-items:center;justify-content:center;height:100%}",
    "@media print{.pt-root{-webkit-print-color-adjust:exact;print-color-adjust:exact}}"
  ].join("\n");
  function getCell(t,day,pid){
    const out=[];
    S.rooms.forEach(room=>{
      (S.schedule[room.id+"_"+day+"_"+pid]||[]).forEach(e=>{
        if(e.teacherId!==t.id&&!(e.coTeacherIds||[]).includes(t.id))return;
        out.push(room.name);
      });
    });
    (S.meetings||[]).forEach(m=>{if(m.teacherId===t.id&&m.day===day&&(m.periods||[]).includes(pid))out.push(m.label||"Lock");});
    return out;
  }
  function getHM(t,day){
    const m=(S.meetings||[]).find(m=>m.teacherId===t.id&&m.day===day&&(m.isAssembly||m.isHomeroom||(m.periods||[]).includes(0)));
    return m?(m.isAssembly?"หอประชุม":(m.label||"Homeroom")):"Homeroom";
  }
  return <div className="pt-root" style={{padding:"6mm 6mm"}}>
    <style>{css}</style>
    {teachers.map((t,ti)=>{
      const dept=(S.depts.find(d=>d.id===t.departmentId)||{}).name||"";
      const assigns=S.assigns.filter(a=>a.teacherId===t.id);
      let grand=0;
      const sumRows=assigns.map(a=>{
        const sub=S.subjects.find(s=>s.id===a.subjectId);if(!sub)return null;
        const rCount=(a.roomIds||[]).length;
        const ppr=sub.periodsPerWeek||Math.round((a.totalPeriods||0)/Math.max(rCount,1));
        const total=a.totalPeriods||(ppr*rCount);grand+=total;
        return{name:(sub.name||"")+" "+(sub.code?"("+sub.code+")":""),rCount,ppr,total};
      }).filter(Boolean);
      const pbAfter=(ti+1)%3===0&&ti<teachers.length-1;
      const showSep=!pbAfter&&ti<teachers.length-1;
      return <div key={t.id}>
        <div className="f3-wrap">
          <div style={{textAlign:"center",marginBottom:3}}>
            {sh?.logo&&<img src={sh.logo} style={{height:32,verticalAlign:"middle",marginRight:6}} alt=""/>}
            <b style={{fontSize:"10pt"}}>ตารางสอน ปีการศึกษา {yr}</b>
          </div>
          <table style={{width:"100%"}}>
            <colgroup>
              <col style={{width:"5%"}}/><col style={{width:"5%"}}/>
              <col style={{width:"11%"}}/><col style={{width:"11%"}}/><col style={{width:"2.5%"}}/>
              <col style={{width:"11%"}}/><col style={{width:"11%"}}/>
              <col style={{width:"2.5%"}}/><col style={{width:"11%"}}/><col style={{width:"11%"}}/>
              <col style={{width:"2.5%"}}/><col style={{width:"11.5%"}}/>
            </colgroup>
            <thead>
              <tr>
                <th style={{...thS,position:"relative",height:36}} rowSpan={2}>
                  <svg style={{position:"absolute",top:0,left:0,width:"100%",height:"100%"}} preserveAspectRatio="none">
                    <line x1="0" y1="0" x2="100%" y2="100%" stroke="#aaa" strokeWidth="0.8"/>
                  </svg>
                  <span style={{position:"absolute",top:2,right:2,fontSize:"6pt"}}>เวลา</span>
                  <span style={{position:"absolute",bottom:2,left:2,fontSize:"6pt"}}>วัน</span>
                </th>
                <th style={{...bS,height:36}} rowSpan={2}><div className="vt">08:00-08:30</div></th>
                <th style={thS}>คาบ 1</th><th style={thS}>คาบ 2</th>
                <th style={{...bS,height:36}} rowSpan={2}><div className="vt">10.10-10.25</div></th>
                <th style={thS}>คาบ 3</th><th style={thS}>คาบ 4</th>
                <th style={{...bS,height:36,width:22}} rowSpan={2}><div className="vt">12.05-13.00</div></th>
                <th style={thS}>คาบ 5</th><th style={thS}>คาบ 6</th>
                <th style={{...bS,height:36}} rowSpan={2}><div className="vt">13.50-14.00</div></th>
                <th style={thS}>คาบ 7</th>
              </tr>
              <tr>
                {PTIMES.map((tm,i)=><th key={i} style={{fontSize:"6pt",fontWeight:400,background:"#f0f0f0",border:"1px solid #aaa",padding:"1px",whiteSpace:"nowrap"}}>{tm}</th>)}
              </tr>
            </thead>
            <tbody>
              {DAYS_PRINT.map((day,di)=>{
                const hm=getHM(t,day);
                const hmBg=hm.includes("หอประชุม")?"#e8f5e9":"#fafff7";
                const cells=PLIST_PRINT.map(p=>getCell(t,day,p.id).join(", "));
                return <tr key={day} style={{height:22,background:di%2===1?"#fafafa":""}}>
                  <td className="f3-day">{sd(day)}</td>
                  <td className="f3-hm" style={{background:hmBg}}>{hm}</td>
                  <td className="f3-cell">{cells[0]}</td><td className="f3-cell">{cells[1]}</td>
                  {di===0&&<td rowSpan={NDAYS} style={{...bS,width:20,verticalAlign:"middle"}}><div className="vt">พักน้อย 15 นาที</div></td>}
                  <td className="f3-cell">{cells[2]}</td><td className="f3-cell">{cells[3]}</td>
                  {di===0&&<td rowSpan={NDAYS} style={{...bS,width:22,verticalAlign:"middle"}}><div className="vt">พักกลางวัน 55 นาที</div></td>}
                  <td className="f3-cell">{cells[4]}</td><td className="f3-cell">{cells[5]}</td>
                  {di===0&&<td rowSpan={NDAYS} style={{...bS,width:20,verticalAlign:"middle"}}><div className="vt">พักน้อย 10 นาที</div></td>}
                  <td className="f3-cell">{cells[6]}</td>
                </tr>;
              })}
            </tbody>
          </table>
          <table style={{width:"100%",fontSize:"8pt",borderCollapse:"collapse",marginTop:2}}>
            <tbody>
              <tr style={{verticalAlign:"top"}}>
                <td style={{width:"42%",paddingRight:8}}>
                  <div style={{color:"#1a237e",fontWeight:700,marginBottom:2}}>กลุ่มสาระ {dept}</div>
                  <div><b>อาจารย์ผู้สอน</b> {t.prefix||""}{t.firstName||""} {t.lastName||""}</div>
                  {assigns.map(a=>{const s=S.subjects.find(x=>x.id===a.subjectId);return s?<div key={a.id} style={{paddingLeft:4}}>{s.name||""} {s.code?"("+s.code+")":""}</div>:null;})}
                </td>
                <td style={{width:"58%"}}>
                  <table style={{width:"100%",fontSize:"8pt",borderCollapse:"collapse"}}>
                    <tbody>
                      {sumRows.map((r,i)=><tr key={i}>
                        <td style={{padding:"1px 3px"}}>{r.name}</td>
                        <td style={{padding:"1px 3px",textAlign:"right",whiteSpace:"nowrap"}}>{r.rCount} ห้อง × {r.ppr} คาบ = <b>{r.total}</b> คาบ</td>
                      </tr>)}
                      <tr style={{borderTop:"1px solid #999"}}>
                        <td style={{padding:"1px 3px",textAlign:"right"}}>รวม</td>
                        <td style={{padding:"1px 3px",textAlign:"right",fontWeight:700}}>= {grand} คาบ</td>
                      </tr>
                    </tbody>
                  </table>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
        {pbAfter&&<div style={{pageBreakAfter:"always"}}/>}
        {showSep&&<hr className="f3-sep"/>}
      </div>;
    })}
  </div>;
}

/* ===== HTML BUILDERS for iframe preview (F2 = 2คน/หน้า, F3 = รหัสห้อง+สรุป) ===== */

function buildF2Html(teachers, S, ay, sh, ps) {
  const P = ps || DEFAULT_PRINT_SETTINGS;
  const C = PRINT_COLORS[P.color] || PRINT_COLORS["แดง"];
  const f = P.fontSize / 100;
  const ff = P.fontFamily || "TH SarabunNew";
  const yr = ay?.year || "2568";
  const logo = sh?.logo
    ? '<img src="' + sh.logo + '" style="width:28px;height:28px;border-radius:50%;object-fit:cover;flex-shrink:0"/>'
    : '<div style="width:28px;height:28px;border:1px solid #999;border-radius:50%;flex-shrink:0"></div>';

  const DAYS2 = ["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
  const SD2 = {"พฤหัสบดี":"พฤหัส"};
  const sd2 = d => SD2[d] || d;
  const NDAYS = DAYS2.length;
  // A4 portrait @ 96dpi = 794px. Margin 8mm each = ~60px total → content 734px
  // 2 teachers/page. Each teacher: header ~32px, 5rows × ROW_H, sep ~10px
  // Total = 2 * (32 + 5*ROW_H) + 10 ≤ 734 → ROW_H ≤ ~65px → use 58px safe
  const ROW_H = Math.round(54 * f);
  const HDR_H = Math.round(28 * f);
  const BRK_W = 13;
  const DAY_W = 34;
  const HM_W  = 38;

  function getEntries(t, day, pid) {
    const out = [];
    Object.entries(S.schedule).forEach(([k, en]) => {
      if (!en?.length) return;
      const pts = k.split("_");
      if (pts[pts.length-2] !== day || parseInt(pts[pts.length-1]) !== pid) return;
      en.forEach(e => {
        const co = e.coTeacherIds?.length ? e.coTeacherIds : (e.coTeacherId ? [e.coTeacherId] : []);
        if (e.teacherId !== t.id && !co.includes(t.id)) return;
        const sub = S.subjects.find(s => s.id === e.subjectId);
        const rid = pts.slice(0,-2).join("_");
        const rm = S.rooms.find(r => r.id === rid);
        out.push({ sub: sub?.shortName || sub?.name || sub?.code || "—", room: rm?.name || "—" });
      });
    });
    return out;
  }
  function getHM2(t, day) {
    const m = (S.meetings||[]).find(m => m.teacherId===t.id && m.day===day && (m.isAssembly||m.isHomeroom||(m.periods||[]).includes(0)));
    return m ? (m.isAssembly ? "หอประชุม" : (m.label||"Homeroom")) : "Homeroom";
  }
  function slotTd2(arr) {
    if (!arr.length) return '<td class="slot"></td>';
    const inner = arr.map(e =>
      '<div class="ent"><div class="esub">' + e.sub + '</div><div class="eroom">' + e.room + '</div></div>'
    ).join("");
    return '<td class="slot">' + inner + '</td>';
  }

  const thBg = 'background:' + C.header + ';color:' + C.headerText;
  const vtH2 = HDR_H * 2;

  const pairs = [];
  for (let i = 0; i < teachers.length; i += 2) pairs.push(teachers.slice(i,i+2));

  let pagesHtml = "";
  pairs.forEach((pair, pi) => {
    let pairHtml = "";
    pair.forEach((t, ti) => {
      const dept = (S.depts.find(d => d.id === t.departmentId)||{}).name || "";
      let tbody = "";
      DAYS2.forEach((day, di) => {
        const hm = getHM2(t, day);
        const hmBg = hm.includes("หอ") ? "#e8f5e9" : "#f5fff5";
        const slots = [1,2,3,4,5,6,7].map(pid => getEntries(t, day, pid));
        const rowBg = di%2===1 ? "#fafafa" : "#fff";
        tbody +=
          '<tr style="height:' + ROW_H + 'px;background:' + rowBg + '">' +
          '<td class="dcell">' + sd2(day) + '</td>' +
          '<td class="hmcell" style="background:' + hmBg + '">' + hm + '</td>' +
          slotTd2(slots[0]) + slotTd2(slots[1]) +
          (di===0 ? '<td rowspan="'+NDAYS+'" class="brkc"><div class="vt2" style="height:'+(ROW_H*NDAYS)+'px">พักน้อย 15 นาที</div></td>' : '') +
          slotTd2(slots[2]) + slotTd2(slots[3]) +
          (di===0 ? '<td rowspan="'+NDAYS+'" class="brkc brkm"><div class="vt2" style="height:'+(ROW_H*NDAYS)+'px">พักกลางวัน 55 นาที</div></td>' : '') +
          slotTd2(slots[4]) + slotTd2(slots[5]) +
          (di===0 ? '<td rowspan="'+NDAYS+'" class="brkc"><div class="vt2" style="height:'+(ROW_H*NDAYS)+'px">พักน้อย 10 นาที</div></td>' : '') +
          slotTd2(slots[6]) +
          '</tr>';
      });

      const sep = ti > 0 ? ' style="border-top:1px dashed #ccc;margin-top:5px;padding-top:5px"' : '';

      // หา division ของครูจาก rooms ที่สอน (ใช้ first match)
      const tDivId=(()=>{
        for(const day of DAYS2){for(const pid of [1,2,3,4,5,6,7]){
          const entry=Object.entries(S.schedule).find(([k,en])=>{
            if(!k.endsWith("_"+day+"_"+pid))return false;
            return(en||[]).some(e=>{const co=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);return e.teacherId===t.id||co.includes(t.id);});
          });
          if(entry){const rid=entry[0].split("_")[0];const rm=S.rooms.find(r=>r.id===rid);if(rm){const lv=S.levels.find(l=>l.id===rm.levelId);if(lv?.divisionId)return lv.divisionId;}}
        }}
        return "m2";
      })();
      const tPcfg=getPeriodCfg(tDivId);
      const p6time=tPcfg.periods[5]?.time||"14.00-14.50";
      const brk4label=tDivId==="p1"?"14.40-14.50":"13.50-14.00";

      pairHtml +=
        '<div' + sep + '>' +
        '<div class="hdr2">' + logo +
        '<div><div class="ttl">' +
        'ตารางสอน ' + (t.prefix||"") + (t.firstName||"") + ' ' + (t.lastName||"") +
        ' \u00a0 ปีการศึกษา ' + yr + '</div>' +
        '<div class="dept2">' + dept + '</div></div></div>' +
        '<table><colgroup>' +
        '<col style="width:' + DAY_W + 'px"/>' +
        '<col style="width:' + HM_W + 'px"/>' +
        '<col/><col/>' +
        '<col style="width:' + BRK_W + 'px"/>' +
        '<col/><col/>' +
        '<col style="width:' + (BRK_W+2) + 'px"/>' +
        '<col/><col/>' +
        '<col style="width:' + BRK_W + 'px"/>' +
        '<col/>' +
        '</colgroup><thead>' +
        '<tr style="height:' + HDR_H + 'px">' +
        '<th class="cnr2" rowspan="2" style="' + thBg + '">' +
        '<svg style="position:absolute;top:0;left:0;width:100%;height:100%" preserveAspectRatio="none"><line x1="0" y1="0" x2="100%" y2="100%" stroke="rgba(255,255,255,0.5)" stroke-width="0.8"/></svg>' +
        '<span class="cr2r">เวลา</span><span class="cr2l">วัน</span></th>' +
        '<th class="brkh" rowspan="2"><div class="vt2" style="height:' + vtH2 + 'px">08.00-08.30</div></th>' +
        '<th class="ph2" style="' + thBg + '">คาบ 1<br/><span class="tl2">08.30-09.20</span></th>' +
        '<th class="ph2" style="' + thBg + '">คาบ 2<br/><span class="tl2">09.20-10.10</span></th>' +
        '<th class="brkh" rowspan="2"><div class="vt2" style="height:' + vtH2 + 'px">10.10-10.25</div></th>' +
        '<th class="ph2" style="' + thBg + '">คาบ 3<br/><span class="tl2">10.25-11.15</span></th>' +
        '<th class="ph2" style="' + thBg + '">คาบ 4<br/><span class="tl2">11.15-12.05</span></th>' +
        '<th class="brkh brkm" rowspan="2"><div class="vt2" style="height:' + vtH2 + 'px">12.05-13.00</div></th>' +
        '<th class="ph2" style="' + thBg + '">คาบ 5<br/><span class="tl2">13.00-13.50</span></th>' +
        (tDivId==="p1"
          ? '<th class="ph2" style="' + thBg + '">คาบ 6<br/><span class="tl2">' + p6time + '</span></th>' +
            '<th class="brkh" rowspan="2"><div class="vt2" style="height:' + vtH2 + 'px">' + brk4label + '</div></th>' +
            '<th class="ph2" style="' + thBg + '">คาบ 7<br/><span class="tl2">14.50-15.40</span></th>'
          : '<th class="ph2" style="' + thBg + '">คาบ 6<br/><span class="tl2">' + p6time + '</span></th>' +
            '<th class="brkh" rowspan="2"><div class="vt2" style="height:' + vtH2 + 'px">' + brk4label + '</div></th>' +
            '<th class="ph2" style="' + thBg + '">คาบ 7<br/><span class="tl2">14.50-15.40</span></th>'
        ) +
        '</tr><tr style="height:' + HDR_H + 'px"></tr>' +
        '</thead><tbody>' + tbody + '</tbody></table></div>';
    });
    pagesHtml += '<div class="pg">' + pairHtml + '</div>';
  });

  const css =
    "@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');\n" +
    "@page{size:A4 portrait;margin:0}\n" +
    "*{box-sizing:border-box;margin:0;padding:0}\n" +
    "html,body{width:794px}\n" +
    "body{font-family:'" + ff + "','Sarabun','Noto Sans Thai',sans-serif;font-size:" + Math.round(9*f) + "px;color:#000;background:#fff;padding:30px 32px}\n" +
    ".pg{width:730px;page-break-after:always;display:flex;flex-direction:column;gap:0}\n" +
    ".pg:last-child{page-break-after:avoid}\n" +
    ".hdr2{display:flex;align-items:center;gap:6px;margin-bottom:3px}\n" +
    ".ttl{font-size:" + Math.round(10*f) + "px;font-weight:700;line-height:1.3}\n" +
    ".dept2{font-size:" + Math.round(8*f) + "px;color:#555}\n" +
    "table{width:100%;border-collapse:collapse;table-layout:fixed}\n" +
    "th{text-align:center;vertical-align:middle;border:1px solid " + C.border + ";font-size:" + Math.round(7*f) + "px;padding:1px}\n" +
    "td{text-align:center;vertical-align:middle;border:1px solid #ddd;font-size:" + Math.round(8*f) + "px;padding:1px;overflow:hidden}\n" +
    ".cnr2{position:relative;padding:0;border:1px solid " + C.border + "}\n" +
    ".cr2r{position:absolute;top:2px;right:2px;font-size:" + Math.round(6*f) + "px;font-weight:700;color:" + C.headerText + "}\n" +
    ".cr2l{position:absolute;bottom:2px;left:2px;font-size:" + Math.round(6*f) + "px;font-weight:700;color:" + C.headerText + "}\n" +
    ".tl2{font-size:" + Math.round(6*f) + "px;font-weight:400}\n" +
    ".ph2{font-size:" + Math.round(7*f) + "px;font-weight:700}\n" +
    ".brkh{background:#fffde7;border:1px solid #ddd;padding:0;overflow:hidden}\n" +
    ".brkm{width:" + (BRK_W+2) + "px}\n" +
    ".brkc{background:#fffde7;border:1px solid #ddd;padding:0;overflow:hidden}\n" +
    ".dcell{font-weight:700;font-size:" + Math.round(8*f) + "px;background:#F3F4F6;border:1px solid #888}\n" +
    ".hmcell{font-size:" + Math.round(7*f) + "px;font-weight:600;line-height:1.2;border:1px solid #888}\n" +
    ".slot{padding:1px;overflow:hidden}\n" +
    ".ent{margin-bottom:1px;line-height:1.2}\n" +
    ".esub{font-weight:700;font-size:" + Math.round(8*f) + "px}\n" +
    ".eroom{font-size:" + Math.round(7*f) + "px;color:#1a237e}\n" +
    ".vt2{writing-mode:vertical-rl;transform:rotate(180deg);white-space:nowrap;font-weight:600;font-size:" + Math.round(7*f) + "px;display:flex;align-items:center;justify-content:center}\n" +
    "@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}\n";

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' + css + '</style></head><body>' + pagesHtml + '</body></html>';
}


function buildF3Html(teachers, S, ay, sh) {
  const yr = ay?.year || "2568";
  const ff = "TH SarabunNew";
  const DAYS3 = ["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
  const SD3 = {"พฤหัสบดี":"พฤหัส"};
  const PTIMES3_DEFAULT = ["08.30-09.20","09.20-10.10","10.25-11.15","11.15-12.05","13.00-13.50","14.00-14.50","14.50-15.40"];
  const PTIMES3_P1      = ["08.30-09.20","09.20-10.10","10.25-11.15","11.15-12.05","13.00-13.50","13.50-14.40","14.50-15.40"];
  const PIDS3 = [1,2,3,4,5,6,7];
  const NDAYS = DAYS3.length;
  const sd3 = d => SD3[d] || d;

  // A4 = 794×1123px. body padding 22px top+bottom = 1079px usable height
  // 3 blocks + 2 separators (5px) = 3B + 10
  // Each block = title(18) + thead_row1(26) + thead_row2(22) + 5×ROW_H + summary(40) + mb(6) = 112 + 5×ROW_H
  // 3×(112 + 5×ROW_H) + 10 ≤ 1079 → ROW_H ≤ (1069/3 - 112)/5 = 22.1 → ROW_H = 22
  const ROW_H  = 22;
  const HDR1_H = 26;  // thead row 1 (period numbers)
  const HDR2_H = 22;  // thead row 2 (times)
  const DAY_W  = 30;
  const HM_W   = 34;
  const BRK_W  = 12;

  function getCell3(t, day, pid) {
    const out = [];
    S.rooms.forEach(room => {
      (S.schedule[room.id+"_"+day+"_"+pid]||[]).forEach(e => {
        if (e.teacherId!==t.id && !(e.coTeacherIds||[]).includes(t.id)) return;
        out.push(room.name);
      });
    });
    (S.meetings||[]).forEach(m => {
      if (m.teacherId===t.id && m.day===day && (m.periods||[]).includes(pid)) out.push(m.label||"Lock");
    });
    return out;
  }
  function getHM3(t, day) {
    const m = (S.meetings||[]).find(m => m.teacherId===t.id && m.day===day && (m.isAssembly||m.isHomeroom||(m.periods||[]).includes(0)));
    return m ? (m.isAssembly ? "หอประชุม" : (m.label||"Homeroom")) : "Homeroom";
  }

  let bodyHtml = "";
  teachers.forEach((t, ti) => {
    const dept = (S.depts.find(d => d.id===t.departmentId)||{}).name || "";
    const assigns = S.assigns.filter(a => a.teacherId===t.id);
    let grand = 0;
    const sumRows = assigns.map(a => {
      const sub = S.subjects.find(s => s.id===a.subjectId); if(!sub) return null;
      const rCount = (a.roomIds||[]).length;
      const ppr = sub.periodsPerWeek || Math.round((a.totalPeriods||0)/Math.max(rCount,1));
      const total = a.totalPeriods||(ppr*rCount); grand+=total;
      return {name:(sub.name||"")+" "+(sub.code?"("+sub.code+")":""),rCount,ppr,total};
    }).filter(Boolean);

    let tbody = "";
    DAYS3.forEach((day, di) => {
      const hm = getHM3(t, day);
      const hmBg = hm.includes("หอ") ? "#e8f5e9" : "#fafff7";
      const cells = PIDS3.map(pid => getCell3(t,day,pid).join(", "));
      const rowBg = di%2===1 ? "#f7f7f7" : "#fff";
      tbody +=
        '<tr style="height:'+ROW_H+'px;background:'+rowBg+'">' +
        '<td class="f3d">'+sd3(day)+'</td>' +
        '<td class="f3h" style="background:'+hmBg+'">'+hm+'</td>' +
        '<td class="f3c">'+cells[0]+'</td>' +
        '<td class="f3c">'+cells[1]+'</td>' +
        (di===0?'<td rowspan="'+NDAYS+'" class="f3b" style="width:'+BRK_W+'px"><div class="vt3" style="height:'+(ROW_H*NDAYS)+'px">พักน้อย 15 นาที</div></td>':'') +
        '<td class="f3c">'+cells[2]+'</td>' +
        '<td class="f3c">'+cells[3]+'</td>' +
        (di===0?'<td rowspan="'+NDAYS+'" class="f3b" style="width:'+(BRK_W+2)+'px"><div class="vt3" style="height:'+(ROW_H*NDAYS)+'px">พักกลางวัน 55 นาที</div></td>':'') +
        '<td class="f3c">'+cells[4]+'</td>' +
        '<td class="f3c">'+cells[5]+'</td>' +
        (di===0?'<td rowspan="'+NDAYS+'" class="f3b" style="width:'+BRK_W+'px"><div class="vt3" style="height:'+(ROW_H*NDAYS)+'px">พักน้อย 10 นาที</div></td>':'') +
        '<td class="f3c">'+cells[6]+'</td>' +
        '</tr>';
    });

    const logoHtml = sh?.logo ? '<img src="'+sh.logo+'" style="height:14px;vertical-align:middle;margin-right:3px"/>' : '';
    const sumRowsHtml = sumRows.map(r =>
      '<tr>' +
      '<td class="sn3">'+r.name+'</td>' +
      '<td class="sv3">'+r.rCount+' ห้อง</td>' +
      '<td class="so3">×</td>' +
      '<td class="sv3">'+r.ppr+' คาบ</td>' +
      '<td class="so3">=</td>' +
      '<td class="sb3">'+r.total+'</td>' +
      '<td class="sn3">คาบ</td>' +
      '</tr>'
    ).join("") +
    '<tr class="stot3">'+
      '<td colspan="4" class="sn3" style="text-align:right">รวม</td>'+
      '<td class="so3">=</td>'+
      '<td class="sb3" style="border-top:2px double #333">'+grand+'</td>'+
      '<td class="sn3">คาบ</td>'+
    '</tr>';

    const tDivId3=(()=>{
      for(const day of DAYS3){for(const pid of PIDS3){
        const rm3=S.rooms.find(room=>(S.schedule[room.id+"_"+day+"_"+pid]||[]).some(e=>{const co=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);return e.teacherId===t.id||co.includes(t.id);}));
        if(rm3){const lv3=S.levels.find(l=>l.id===rm3.levelId);if(lv3?.divisionId)return lv3.divisionId;}
      }}return"m2";
    })();
    const PTIMES3=tDivId3==="p1"?PTIMES3_P1:PTIMES3_DEFAULT;
    const brk4lbl3=tDivId3==="p1"?"14.40-14.50":"13.50-14.00";
    const showSep  = !pbAfter && ti<teachers.length-1;

    bodyHtml +=
      '<div class="f3w">' +
      // title row
      '<div class="f3t">'+logoHtml+'<b>ตารางสอน ปีการศึกษา '+yr+'</b></div>' +
      // main timetable
      '<table><colgroup>' +
        '<col style="width:'+DAY_W+'px"/>' +
        '<col style="width:'+HM_W+'px"/>' +
        '<col/><col/>' +
        '<col style="width:'+BRK_W+'px"/>' +
        '<col/><col/>' +
        '<col style="width:'+(BRK_W+2)+'px"/>' +
        '<col/><col/>' +
        '<col style="width:'+BRK_W+'px"/>' +
        '<col/>' +
      '</colgroup><thead>' +
      '<tr style="height:'+HDR1_H+'px">' +
        '<th class="cnr3" rowspan="2">' +
          '<svg style="position:absolute;top:0;left:0;width:100%;height:100%" preserveAspectRatio="none"><line x1="0" y1="0" x2="100%" y2="100%" stroke="#aaa" stroke-width="0.8"/></svg>' +
          '<span class="cRr">เวลา</span><span class="cLl">วัน</span>' +
        '</th>' +
        '<th class="f3b" rowspan="2" style="width:'+HM_W+'px"><div class="vt3" style="height:'+(HDR1_H+HDR2_H)+'px">08:00-08:30</div></th>' +
        '<th class="f3h2">คาบ 1</th><th class="f3h2">คาบ 2</th>' +
        '<th class="f3b" rowspan="2" style="width:'+BRK_W+'px"><div class="vt3" style="height:'+(HDR1_H+HDR2_H)+'px">10.10-10.25</div></th>' +
        '<th class="f3h2">คาบ 3</th><th class="f3h2">คาบ 4</th>' +
        '<th class="f3b" rowspan="2" style="width:'+(BRK_W+2)+'px"><div class="vt3" style="height:'+(HDR1_H+HDR2_H)+'px">12.05-13.00</div></th>' +
        '<th class="f3h2">คาบ 5</th><th class="f3h2">คาบ 6</th>' +
        '<th class="f3b" rowspan="2" style="width:'+BRK_W+'px"><div class="vt3" style="height:'+(HDR1_H+HDR2_H)+'px">'+brk4lbl3+'</div></th>' +
        '<th class="f3h2">คาบ 7</th>' +
      '</tr>' +
      '<tr style="height:'+HDR2_H+'px">' +
        PTIMES3.map(tm=>'<th class="f3tm">'+tm+'</th>').join("") +
      '</tr>' +
      '</thead><tbody>'+tbody+'</tbody></table>' +
      // summary
      '<table class="stbl3"><tbody><tr valign="top">' +
        '<td class="sl3">' +
          '<div class="sdept3">กลุ่มสาระ '+dept+'</div>' +
          '<div><b>อาจารย์ผู้สอน</b> '+(t.prefix||"")+(t.firstName||"")+" "+(t.lastName||"")+'</div>' +
          assigns.map(a=>{const sub=S.subjects.find(s=>s.id===a.subjectId);return sub?'<div class="ssub3">'+(sub.name||"")+(sub.code?" ("+sub.code+")":"")+'</div>':"";}).join("") +
        '</td>' +
        '<td class="sr3"><table class="srt3"><tbody>'+sumRowsHtml+'</tbody></table></td>' +
      '</tr></tbody></table>' +
      '</div>' +
      (pbAfter ? '<div class="pb3"></div>' : '') +
      (showSep  ? '<hr class="sep3"/>' : '');
  });

  // ===== CSS =====
  const css = [
    "@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');",
    "@page{size:A4 portrait;margin:0}",
    "*{box-sizing:border-box;margin:0;padding:0}",
    "html,body{width:794px}",
    "body{font-family:'"+ff+"','Sarabun',sans-serif;font-size:9px;color:#000;background:#fff;padding:22px 24px}",
    // page break
    ".pb3{page-break-after:always}",
    ".sep3{border:none;border-top:1px dashed #ccc;margin:5px 0}",
    // block
    ".f3w{page-break-inside:avoid;width:100%;margin-bottom:6px}",
    ".f3t{text-align:center;margin-bottom:3px;font-size:9.5px}",
    // tables
    "table{width:100%;border-collapse:collapse;table-layout:fixed}",
    "th,td{overflow:hidden;vertical-align:middle;text-align:center}",
    // timetable header
    ".cnr3{position:relative;padding:0;background:#f0f0f0;border:1px solid #999}",
    ".cRr{position:absolute;top:2px;right:2px;font-size:6.5px;font-weight:700}",
    ".cLl{position:absolute;bottom:2px;left:2px;font-size:6.5px;font-weight:700}",
    ".f3h2{font-size:8px;font-weight:700;background:#f0f0f0;border:1px solid #999;padding:1px}",
    ".f3tm{font-size:6px;font-weight:400;background:#f5f5f5;border:1px solid #bbb;padding:0;white-space:nowrap}",
    ".f3b{background:#fffde7;border:1px solid #ddd;padding:0;overflow:hidden}",
    // timetable data cells
    ".f3d{font-weight:700;font-size:9px;background:#f5f5f5;border:1px solid #888;padding:1px}",
    ".f3h{font-size:7px;font-weight:600;line-height:1.2;border:1px solid #888;padding:0 1px}",
    ".f3c{font-size:8.5px;font-weight:700;border:1px solid #ddd;padding:1px}",
    // vertical text
    ".vt3{writing-mode:vertical-rl;transform:rotate(180deg);white-space:nowrap;font-weight:600;font-size:6.5px;display:flex;align-items:center;justify-content:center}",
    // summary table
    ".stbl3{margin-top:2px;font-size:7.5px}",
    ".sl3{width:42%;text-align:left;vertical-align:top;padding-right:6px}",
    ".sr3{width:58%;vertical-align:top}",
    ".sdept3{color:#1a237e;font-weight:700;margin-bottom:1px}",
    ".ssub3{padding-left:4px;font-size:7px}",
    ".srt3{font-size:7.5px;width:100%}",
    ".sn3{padding:0 3px;text-align:left}",
    ".sv3{padding:0 3px;text-align:right;white-space:nowrap}",
    ".so3{padding:0 2px;text-align:center}",
    ".sb3{padding:0 3px;text-align:right;font-weight:700}",
    ".stot3 td{border-top:1px solid #aaa}",
    "@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}",
  ].join("\n");

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'+css+'</style></head><body>'+bodyHtml+'</body></html>';
}


// PrintPreviewModal — iframe-based: รับ {html:string} แสดงผ่าน srcdoc
// ไม่มี CSS leak, ไม่มี scale เพี้ยน — browser จัดการ layout เองใน sandbox
function PrintPreviewModal({data,onClose}){
  if(!data)return null;
  const iframeRef=useRef(null);
  const handlePrint=()=>{
    const fr=iframeRef.current;
    if(!fr)return;
    try{ fr.contentWindow.focus(); fr.contentWindow.print(); }
    catch(e){
      const w=window.open("","_blank");
      if(w){w.document.write(data.html);w.document.close();setTimeout(()=>w.print(),400);}
    }
  };
  return (
    <div style={{position:"fixed",inset:0,zIndex:9000,display:"flex",flexDirection:"column",background:"rgba(0,0,0,0.85)"}}>
      <div style={{background:"#111827",padding:"10px 16px",display:"flex",alignItems:"center",gap:10,flexShrink:0}}>
        <span style={{color:"#fff",fontWeight:700,fontSize:15}}>🖨️ ตัวอย่างก่อนพิมพ์</span>
        <div style={{flex:1}}/>
        <button data-ui-control="true" onClick={handlePrint} style={{background:"#B91C1C",color:"#fff",border:"none",borderRadius:8,padding:"8px 20px",fontSize:14,fontWeight:700,cursor:"pointer"}}>🖨️ พิมพ์</button>
        <button data-ui-control="true" onClick={onClose} style={{background:"#374151",color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",fontSize:14,cursor:"pointer"}}>✕ ปิด</button>
      </div>
      <iframe
        ref={iframeRef}
        srcDoc={data.html}
        style={{flex:1,border:"none",background:"#525659"}}
        title="print-preview"
      />
    </div>
  );
}

/* ===== END REACT PRINT PREVIEW SYSTEM ===== */

export default function App() {
  const [page,setPage]=useState("dashboard");
  const [side,setSide]=useState(true);
  const [toast,setToast]=useState(null);
  const [syncing,setSyncing]=useState(false);
  const [gasReady,setGasReady]=useState(false);

  // ===== AUTH STATE =====
  const [authUser,setAuthUser]=useState(undefined);
  const [userPerms,setUserPerms]=useState(null);
  const [isAdmin,setIsAdmin]=useState(false);
  const [accessError,setAccessError]=useState('');
  const [showAdmin,setShowAdmin]=useState(false);
  const refreshPerms=async()=>{
    if(!authUser)return;
    const token=await authUser.getIdTokenResult(true);
    setIsAdmin(token.claims.admin===true);
    setUserPerms(effectivePermissions(await fsGetPermissions(authUser.uid),token.claims.admin===true));
  };
  useEffect(()=>{
    const {auth,db}=getFB();if(!auth){setAuthUser(null);return;}
    let unsubscribePermissions=()=>{},generation=0;
    const unsubscribe=onAuthStateChanged(auth,async user=>{
      const current=++generation;
      unsubscribePermissions();setUserPerms(null);setIsAdmin(false);setAccessError('');setAuthUser(user);
      if(!user)return;
      try{
        if(!schoolAccount(user))throw Error('กรุณาใช้บัญชีโรงเรียนที่ยืนยันอีเมลแล้ว');
        const token=await user.getIdTokenResult(true);
        const data=await fsGetPermissions(user.uid);
        if(current!==generation)return;
        if(!data)await fsSetPermissions(user.uid,{displayName:user.displayName||'',email:user.email});
        if(current!==generation)return;
        setIsAdmin(token.claims.admin===true);
        unsubscribePermissions=onSnapshot(doc(db,'permissions',user.uid),snap=>{
          if(current===generation)setUserPerms(effectivePermissions(snap.exists()?snap.data():null,token.claims.admin===true));
        },e=>{if(current===generation){setUserPerms(null);setAccessError('ตรวจสิทธิ์ไม่สำเร็จ: '+e.message);}});
      }catch(e){if(current===generation)setAccessError(e.message);}
    });
    return()=>{generation++;unsubscribe();unsubscribePermissions();};
  },[]);

  const handleLogout=async()=>{
    if(isSavingRef.current){await uiAlert('กำลังบันทึก กรุณารอสักครู่ก่อนออกจากระบบ');return;}
    if(LIVE.ready&&fsReadyRef.current&&canonical({...stateRef.current,...shRef.current})!==cleanPayload.current){await uiAlert('ยังมีงานที่ไม่ได้บันทึก กรุณารอให้บันทึกเสร็จก่อนออกจากระบบ');return;}
    const {auth}=getFB();
    if(auth)await signOut(auth);
  };

  // division state — persist ใน localStorage (ไม่ใช่ per-division key)
  const [divId,setDivId]=useState(()=>localStorage.getItem("dara_preview_division")||"m2");
  const div=DIVISIONS.find(d=>d.id===divId)||DIVISIONS[3];

  // helper โหลด/บันทึก per-division
  const loadD=(key,fb)=>loadLS(divId+"_"+key,fb);
  const saveD=(key,data)=>saveLS(divId+"_"+key,data);

  const [levels,setLevels]=useState(()=>loadLS(divId+"_levels",DIVISIONS.find(d=>d.id===divId)?.defaultLevels.map(n=>({id:gid(),name:n}))||[]));
  const [plans,setPlans]=useState(()=>loadLS(divId+"_plans",[]));
  const [depts,setDepts]=useState(()=>loadLS(divId+"_depts",[]));
  const [teachers,setTeachers]=useState(()=>loadLS(divId+"_teachers",[]));
  const [subjects,setSubjects]=useState(()=>loadLS(divId+"_subjects",[]));
  const [rooms,setRooms]=useState(()=>loadLS(divId+"_rooms",[]));
  const [specialRooms,setSpecialRooms]=useState(()=>loadLS(divId+"_specialRooms",[]));
  const [assigns,setAssigns]=useState(()=>loadLS(divId+"_assigns",[]));
  const [meetings,setMeetings]=useState(()=>loadLS(divId+"_meetings",[]));
  const [schedule,setSchedule]=useState(()=>loadLS(divId+"_schedule",{}));
  const [locks,setLocks]=useState(()=>loadLS(divId+"_locks",{}));

  const [academicYear,setAcademicYear]=useState(()=>loadLS("academicYear",{year:"2568",semester:"1"}));
  const [schoolHeader,setSchoolHeader]=useState(()=>loadLS("schoolHeader",{name:"โรงเรียนดาราวิทยาลัย",logo:""}));

  useEffect(()=>{saveLS("academicYear",academicYear);if(fsReadyRef.current)syncToFirestore();},[academicYear]);
  useEffect(()=>{
    saveLS("schoolHeader",schoolHeader);
    if(fsReadyRef.current) syncToFirestore();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[schoolHeader]);
  // บันทึก division ที่เลือกไว้
  useEffect(()=>{ localStorage.setItem("dara_preview_division",divId); },[divId]);

  const stateRef=useRef({});
  useEffect(()=>{stateRef.current={levels,plans,depts,teachers,subjects,rooms,specialRooms,assigns,meetings,schedule,locks}},[levels,plans,depts,teachers,subjects,rooms,specialRooms,assigns,meetings,schedule,locks]);
  // sync schoolHeader และ academicYear ผ่าน GAS ด้วย เพื่อให้ทุกเครื่องเห็นโลโก้และปีการศึกษาเดียวกัน
  const shRef=useRef({});
  useEffect(()=>{shRef.current={schoolHeader,academicYear};},[schoolHeader,academicYear]);

  // เมื่อ switch division → โหลดข้อมูลชุดใหม่
  const switchDivision=async(newDivId)=>{
    if(isSavingRef.current){await uiAlert('กำลังบันทึก กรุณารอสักครู่ก่อนเปลี่ยนช่วงชั้น');return;}
    if(LIVE.ready&&fsReadyRef.current&&canonical({...stateRef.current,...shRef.current})!==cleanPayload.current){await uiAlert('ยังมีงานที่ไม่ได้บันทึก กรุณารอให้บันทึกเสร็จก่อนเปลี่ยนช่วงชั้น');return;}
    const d=DIVISIONS.find(x=>x.id===newDivId);
    if(!d) return;
    clearTimeout(saveTimer.current);fsReadyRef.current=false;setSyncError('');
    setDivId(newDivId);
    setLevels(loadLS(newDivId+"_levels",d.defaultLevels.map(n=>({id:gid(),name:n}))));
    setPlans(loadLS(newDivId+"_plans",[]));
    setDepts(loadLS(newDivId+"_depts",[]));
    setTeachers(loadLS(newDivId+"_teachers",[]));
    setSubjects(loadLS(newDivId+"_subjects",[]));
    setRooms(loadLS(newDivId+"_rooms",[]));
    setSpecialRooms(loadLS(newDivId+"_specialRooms",[]));
    setAssigns(loadLS(newDivId+"_assigns",[]));
    setMeetings(loadLS(newDivId+"_meetings",[]));
    setSchedule(loadLS(newDivId+"_schedule",{}));
    setLocks(loadLS(newDivId+"_locks",{}));
    setGasReady(false);
    fsReadyRef.current=false;
    setPage("dashboard");
    st("เปลี่ยนเป็น "+d.name);
  };

  // Auto-switch ไปยัง division แรกที่มีสิทธิ์ ถ้า divId ปัจจุบันไม่มีสิทธิ์
  useEffect(()=>{
    if(!firebaseConfigured||!userPerms)return;
    const currentOk=Boolean(userPerms?.divisions?.[divId]);
    if(!currentOk){
      const firstAllowed=DIVISIONS.find(d=>Boolean(userPerms?.divisions?.[d.id]));
      if(firstAllowed&&firstAllowed.id!==divId){
        switchDivision(firstAllowed.id);
      }
    }
    // Auto-redirect ครู (isTeacher only) ไปหน้าแลกคาบทันทีที่โหลดสิทธิ์
    if(userPerms?.divisions?.isTeacher===true&&!(userPerms?.divisions?.canEdit)){
      setPage('swap');
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[userPerms]);

  // ===== FIRESTORE REALTIME SYNC =====
  const saveTimer=useRef(null);
  const fsReadyRef=useRef(false); // กัน loop: onSnapshot trigger → setState → save → onSnapshot
  const isSavingRef=useRef(false); // กัน onSnapshot overwrite ขณะ save

  const [syncError,setSyncError]=useState('');
  const remoteBaseline=useRef(null);
  const cleanPayload=useRef(null);
  const rightsRef=useRef(null);rightsRef.current={user:authUser,perms:userPerms,divId};
  const syncToFirestore=useCallback((immediate=false)=>{
    const rights=rightsRef.current;
    if(!LIVE.ready||!fsReadyRef.current||!rights?.user||!rights.perms?.divisions?.[divId]||!rights.perms?.divisions?.canEdit)return;
    clearTimeout(saveTimer.current);
    saveTimer.current=setTimeout(async()=>{
      if(!fsReadyRef.current||rightsRef.current?.user?.uid!==rights.user.uid||rightsRef.current?.divId!==divId||!rightsRef.current?.perms?.divisions?.canEdit||!rightsRef.current?.perms?.divisions?.[divId])return;
      const payload={...stateRef.current,schoolHeader:shRef.current.schoolHeader,academicYear:shRef.current.academicYear};
      if(canonical(payload)===cleanPayload.current)return;
      if(isSavingRef.current){syncToFirestore();return;}
      isSavingRef.current=true;setSyncing(true);
      try{remoteBaseline.current=await fsSaveTimetable(divId,payload,remoteBaseline.current);cleanPayload.current=canonical(payload);}
      catch(e){fsReadyRef.current=false;setSyncError('ยังไม่ได้บันทึก: '+e.message);}
      finally{isSavingRef.current=false;setSyncing(false);if(fsReadyRef.current)syncToFirestore();}
    },immediate?0:500);
  },[divId]);
  useEffect(()=>{
    if(!LIVE.ready||!authUser||!userPerms?.divisions?.[divId])return;
    let active=true;
    const loadingTimeout=setTimeout(()=>{if(active&&!fsReadyRef.current){setSyncing(false);setSyncError('ยังเชื่อมต่อฐานข้อมูลไม่ได้ กรุณาตรวจอินเทอร์เน็ตและลองใหม่');}},15000);
    fsReadyRef.current=false;setSyncing(true);setSyncError('');
    let snapshotSequence=0;
    const unsub=fsSubscribeTimetable(divId,d=>{
      const sequence=++snapshotSequence;
      const apply=()=>{
      if(!active||sequence!==snapshotSequence)return;
      if(isSavingRef.current){setTimeout(apply,80);return;}
      if(fsReadyRef.current&&canonical(d)===remoteBaseline.current)return;
      // Do not replace unsaved edits when another editor saves first.
      if(fsReadyRef.current&&saveTimer.current&&canonical({...stateRef.current,...shRef.current})!==cleanPayload.current){
        fsReadyRef.current=false;clearTimeout(saveTimer.current);setSyncing(false);setSyncError('มีข้อมูลใหม่จากเครื่องอื่น กรุณาสำรองงานที่ยังไม่บันทึก แล้วโหลดข้อมูลล่าสุด');return;
      }
      fsReadyRef.current=false;
      remoteBaseline.current=canonical(d);
      cleanPayload.current=canonical({...Object.fromEntries(DATA_FIELDS.map(k=>[k,d[k]||(k==='schedule'||k==='locks'?{}:[])])),schoolHeader:d.schoolHeader||{name:'โรงเรียนดาราวิทยาลัย',logo:''},academicYear:d.academicYear||{year:String(new Date().getFullYear()+543),semester:'1'}});
      setLevels(d.levels||[]);setPlans(d.plans||[]);setDepts(d.depts||[]);setTeachers(d.teachers||[]);
      setSubjects(d.subjects||[]);setRooms(d.rooms||[]);setSpecialRooms(d.specialRooms||[]);
      setAssigns(d.assigns||[]);setMeetings(d.meetings||[]);setSchedule(d.schedule||{});setLocks(d.locks||{});
      setSchoolHeader(d.schoolHeader||{name:'โรงเรียนดาราวิทยาลัย',logo:''});
      setAcademicYear(d.academicYear||{year:String(new Date().getFullYear()+543),semester:'1'});
      // Hydration is not a user edit. Enable writes after its React effects finish.
      setTimeout(()=>{if(active){fsReadyRef.current=true;setSyncing(false);setGasReady(true);}},0);
      };apply();
    },e=>{if(active){fsReadyRef.current=false;setSyncing(false);setSyncError('โหลดข้อมูลไม่สำเร็จ: '+e.message);}});
    return()=>{active=false;unsub();clearTimeout(loadingTimeout);clearTimeout(saveTimer.current);fsReadyRef.current=false;};
  },[divId,authUser,userPerms]);
  useEffect(()=>{const guard=e=>{if(isSavingRef.current||(LIVE.ready&&fsReadyRef.current&&canonical({...stateRef.current,...shRef.current})!==cleanPayload.current)){e.preventDefault();e.returnValue='';}};window.addEventListener('beforeunload',guard);return()=>window.removeEventListener('beforeunload',guard);},[]);

  // Auto-save ไป localStorage (cache offline) + Firestore เมื่อข้อมูลเปลี่ยน
  useEffect(()=>{ saveLS(divId+"_levels",levels);       if(fsReadyRef.current) syncToFirestore(); },[levels,divId]); // eslint-disable-line
  useEffect(()=>{ saveLS(divId+"_plans",plans);         if(fsReadyRef.current) syncToFirestore(); },[plans,divId]); // eslint-disable-line
  useEffect(()=>{ saveLS(divId+"_depts",depts);         if(fsReadyRef.current) syncToFirestore(); },[depts,divId]); // eslint-disable-line
  useEffect(()=>{ saveLS(divId+"_teachers",teachers);   if(fsReadyRef.current) syncToFirestore(); },[teachers,divId]); // eslint-disable-line
  useEffect(()=>{ saveLS(divId+"_subjects",subjects);   if(fsReadyRef.current) syncToFirestore(); },[subjects,divId]); // eslint-disable-line
  useEffect(()=>{ saveLS(divId+"_rooms",rooms);         if(fsReadyRef.current) syncToFirestore(); },[rooms,divId]); // eslint-disable-line
  useEffect(()=>{ saveLS(divId+"_specialRooms",specialRooms); if(fsReadyRef.current) syncToFirestore(); },[specialRooms,divId]); // eslint-disable-line
  useEffect(()=>{ saveLS(divId+"_assigns",assigns);     if(fsReadyRef.current) syncToFirestore(); },[assigns,divId]); // eslint-disable-line
  useEffect(()=>{ saveLS(divId+"_meetings",meetings);   if(fsReadyRef.current) syncToFirestore(); },[meetings,divId]); // eslint-disable-line
  useEffect(()=>{ saveLS(divId+"_schedule",schedule);   if(fsReadyRef.current) syncToFirestore(); },[schedule,divId]); // eslint-disable-line
  useEffect(()=>{ saveLS(divId+"_locks",locks);         if(fsReadyRef.current) syncToFirestore(); },[locks,divId]); // eslint-disable-line

  const st=(m,t="success")=>setToast({message:m,type:t});
  const gc=did=>{const i=depts.findIndex(d=>d.id===did);return DC[i%DC.length]||DC[0]};

  const nav=[
    {id:"dashboard",icon:"home",label:"แดชบอร์ด"},
    {id:"levels",icon:"grid",label:"ระดับชั้น / ห้องเรียน"},
    {id:"plans",icon:"layers",label:"แผนการเรียน"},
    {id:"departments",icon:"users",label:"กลุ่มสาระ"},
    {id:"teachers",icon:"users",label:"จัดการครู"},
    {id:"subjects",icon:"book",label:"จัดการวิชา"},
    {id:"specialrooms",icon:"home",label:"ห้องพิเศษ"},
    {id:"assignments",icon:"edit",label:"มอบหมายงานครู"},
    {id:"homeroom",icon:"users",label:"ครูประจำชั้น"},
    {id:"meetings",icon:"clock",label:"คาบล็อค / ประชุม"},
    {id:"scheduler",icon:"grid",label:"จัดตารางสอน"},
    {id:"swap",icon:"layers",label:"แลกคาบ"},
    {id:"reports",icon:"download",label:"รายงาน / Export"},
    {id:"settings",icon:"file",label:"ตั้งค่า / ปีการศึกษา"},
  ];
  const S={levels,plans,depts,teachers,subjects,rooms,specialRooms,assigns,meetings,schedule,locks};
  const U={setLevels,setPlans,setDepts,setTeachers,setSubjects,setRooms,setSpecialRooms,setAssigns,setMeetings,setSchedule,setLocks};

  // ===== AUTH GUARDS =====
  const firebaseConfigured=LIVE.ready;

  if(!PREVIEW_MODE&&!LIVE.ready)return <main className="content-card" style={{margin:40,padding:32}}><h1>ยังไม่ได้ตั้งค่าระบบ</h1><p>กรุณาให้ผู้ดูแลตั้งค่าการเชื่อมต่อก่อนเปิดใช้งาน ข้อมูลตารางยังไม่ถูกโหลดหรือบันทึก</p></main>;
  if(accessError)return <main className="content-card" style={{margin:40,padding:32}}><h1>ไม่สามารถเปิดข้อมูลได้</h1><p role="alert">{accessError}</p><button onClick={handleLogout}>ออกจากระบบ</button><button onClick={()=>window.location.reload()}>ลองใหม่</button></main>;
  if(firebaseConfigured&&authUser&&!userPerms)return <p role="status">กำลังตรวจสิทธิ์บัญชี...</p>;
  if(syncError)return <main className="content-card" style={{margin:40,padding:32}}><h1>หยุดการบันทึกชั่วคราว</h1><p role="alert">{syncError}</p><button onClick={()=>{const url=URL.createObjectURL(new Blob([JSON.stringify({...stateRef.current,...shRef.current},null,2)],{type:'application/json'}));const a=document.createElement('a');a.href=url;a.download='dara-unsaved-recovery.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}}>ดาวน์โหลดงานในหน้าจอนี้</button><button onClick={()=>window.location.reload()}>โหลดข้อมูลล่าสุด (ทิ้งงานที่ยังไม่บันทึก)</button></main>;
  // Loading
  if(firebaseConfigured&&authUser===undefined){
    return <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#991B1B,#7F1D1D)"}}>
      <div style={{color:"#fff",fontSize:16,fontWeight:600}}>⏳ กำลังโหลด...</div>
    </div>;
  }

  // Not logged in
  if(firebaseConfigured&&!authUser){
    return <LoginScreen onLogin={u=>setAuthUser(u)}/>;
  }

  // Admin panel
  if(showAdmin&&isAdmin){
    return <AccessAdmin db={getFB().db} user={authUser} onBack={()=>{setShowAdmin(false);refreshPerms();}} refreshPerms={()=>refreshPerms()}/>;
  }

  if(LIVE.ready&&userPerms?.divisions?.[divId]&&!fsReadyRef.current)return <p role="status">กำลังโหลดตารางจากโรงเรียน...</p>;

  // Filter division selector ตาม permissions
  const availDivs=firebaseConfigured
    ?DIVISIONS.filter(d=>Boolean(userPerms?.divisions?.[d.id]))
    :DIVISIONS;

  const divHasAccess=!firebaseConfigured||Boolean(userPerms?.divisions?.[divId]);

  const canVisit=(id)=>{
    if(PREVIEW_MODE) return true;
    if(userPerms?.divisions?.isTeacher&&!userPerms?.divisions?.canEdit) return id==='swap';
    return Boolean(userPerms?.divisions?.canEdit)||['dashboard','reports','swap'].includes(id);
  };
  const sp=progress(S);
  return <Workspace logo={schoolLogo(schoolHeader.logo)} page={page} setPage={setPage} div={div} divisions={availDivs} switchDivision={switchDivision} ay={academicYear} syncing={syncing} demo={PREVIEW_MODE} user={authUser} onLogout={handleLogout} onAdmin={isAdmin?()=>setShowAdmin(true):null} canVisit={canVisit}>
    {!divHasAccess ? <div className="empty-message">ไม่มีสิทธิ์เข้าระดับชั้นนี้ กรุณาติดต่อผู้ดูแลระบบ</div> : !canVisit(page) ? <div className="empty-message">ไม่มีสิทธิ์แก้ไขหน้านี้ <button data-ui-control="true" className="secondary" onClick={()=>setPage('reports')}>ดูรายงาน</button></div> : <>
    {page==='dashboard'&&<Dashboard S={S} setPage={setPage} ay={academicYear}/>}
    {page==='teachers'&&<><PageHeading eyebrow="เตรียมข้อมูล" title="จัดการครู" description="ค้นหาครู ตรวจภาระสอน และปรับข้อมูลในที่เดียว"/><Teachers S={S} U={U} st={st} gc={gc}/></>}
    {page==='scheduler'&&<><PageHeading eyebrow="ตารางสอน" title="พื้นที่จัดตารางสอน" description="เลือกรายครูหรือรายห้อง แล้วลากวิชาลงคาบที่ต้องการ"><button data-ui-control="true" className="secondary" onClick={()=>setPage('reports')}><Glyph name="check"/> ตรวจสอบ / รายงาน</button></PageHeading><div className="scheduler-summary"><span>ลงตารางแล้ว <strong>{sp.placed} คาบ</strong></span><span>รอจัด <strong>{sp.remaining} คาบ</strong></span><span>คาบล็อก <strong>{Object.values(locks).filter(Boolean).length}</strong></span></div><div className="scheduler-help">เริ่มด้วยการเลือกครูหรือห้อง · ลากวิชาจากรายการลงในตาราง · ใช้ “จัดตารางอัตโนมัติ” เพื่อช่วยลงคาบที่เหลือ</div><div className="legacy-screen scheduler-surface"><Scheduler S={S} U={U} st={st} gc={gc} isSavingRef={isSavingRef} fsReadyRef={fsReadyRef} fsSave={(s)=>fsSaveTimetable(divId,{...stateRef.current,schedule:s})}/></div></>}
    {!['dashboard','teachers','scheduler'].includes(page)&&<div className="legacy-screen"><PageHeading eyebrow="งานวิชาการ" title={nav.find(n=>n.id===page)?.label}/>
    {page==='levels'&&<Levels S={S} U={U} st={st}/>}
    {page==='plans'&&<Plans S={S} U={U} st={st}/>}
    {page==='departments'&&<Depts S={S} U={U} st={st} gc={gc}/>}
    {page==='subjects'&&<Subjects S={S} U={U} st={st} gc={gc}/>}
    {page==='specialrooms'&&<SpecialRooms S={S} U={U} st={st}/>}
    {page==='assignments'&&<Assigns S={S} U={U} st={st} gc={gc}/>}
    {page==='homeroom'&&<HomeroomSettings S={S} U={U} st={st}/>}
    {page==='meetings'&&<Meetings S={S} U={U} st={st} gc={gc}/>}
    {page==='swap'&&<SwapPage S={S} st={st} ay={academicYear} sh={{...schoolHeader,logo:schoolLogo(schoolHeader.logo)}}/>}
    {page==='reports'&&<Reports S={S} U={U} st={st} gc={gc} ay={academicYear} sh={{...schoolHeader,logo:schoolLogo(schoolHeader.logo)}}/>}
    {page==='settings'&&<Settings S={S} U={U} st={st} ay={academicYear} setAY={setAcademicYear} sh={{...schoolHeader,logo:schoolLogo(schoolHeader.logo)}} setSH={setSchoolHeader} div={div} setSyncing={setSyncing} stateRef={stateRef}/>}
    </div>}
    </>}
    {toast&&<Toast {...toast} onClose={()=>setToast(null)}/>}
  </Workspace>;
}

/* ===== DASHBOARD ===== */
function Dash({S,setPage}){
  const stats=[{l:"ระดับชั้น",v:S.levels.length,c:"#DC2626"},{l:"แผนการเรียน",v:S.plans.length,c:"#7C3AED"},{l:"กลุ่มสาระ",v:S.depts.length,c:"#2563EB"},{l:"ครู",v:S.teachers.length,c:"#059669"},{l:"วิชา",v:S.subjects.length,c:"#D97706"},{l:"ห้อง",v:S.rooms.length,c:"#DB2777"}];
  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:16,marginBottom:32}}>
      {stats.map((s,i)=><div data-ui-surface="true" key={i} style={{background:"#fff",borderRadius:14,padding:20,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}><div style={{fontSize:28,fontWeight:800}}>{s.v}</div><div style={{fontSize:13,color:"#6B7280",marginTop:2}}>{s.l}</div><div style={{height:4,background:s.c,borderRadius:2,marginTop:12,width:"40%"}}/></div>)}
    </div>
    <div data-ui-surface="true" className="content-card" style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
      <h3 style={{fontSize:16,fontWeight:700,marginBottom:16}}>ขั้นตอนการใช้งาน</h3>
      {[{s:1,t:"สร้างระดับชั้นและห้องเรียน",p:"levels"},{s:2,t:"สร้างแผนการเรียน (ใช้ร่วมข้ามระดับได้)",p:"plans"},{s:3,t:"สร้างกลุ่มสาระการเรียนรู้",p:"departments"},{s:4,t:"เพิ่มครู + กำหนดคาบที่ได้รับ",p:"teachers"},{s:5,t:"สร้างวิชา + ระบุระดับชั้น",p:"subjects"},{s:6,t:"มอบหมายวิชาและห้องให้ครู",p:"assignments"},{s:7,t:"ตั้งคาบล็อค/ประชุม",p:"meetings"},{s:8,t:"จัดตารางสอน (Drag & Drop)",p:"scheduler"},{s:9,t:"ตรวจสอบและ Export CSV",p:"reports"}].map(s=><div data-ui-surface="true" key={s.s} onClick={()=>setPage(s.p)} style={{display:"flex",alignItems:"center",gap:14,padding:"12px 16px",borderRadius:10,cursor:"pointer",background:"#F9FAFB",marginBottom:6}} onMouseEnter={e=>e.currentTarget.style.background="#FEE2E2"} onMouseLeave={e=>e.currentTarget.style.background="#F9FAFB"}><div style={{width:30,height:30,borderRadius:"50%",background:"#DC2626",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,flexShrink:0}}>{s.s}</div><span style={{fontSize:14}}>{s.t}</span></div>)}
    </div>
  </div>;
}

/* ===== LEVELS & ROOMS (+ import/export) ===== */
function Levels({S,U,st}){
  const [rm,setRm]=useState(false);
  const [rf,setRf]=useState({levelId:"",planId:"",name:""});
  // Auto-migrate: levels ที่ไม่มี divisionId → guess จากชื่อ
  useEffect(()=>{
    const needMigrate=S.levels.some(l=>!l.divisionId);
    if(needMigrate){
      U.setLevels(p=>p.map(l=>l.divisionId?l:{...l,divisionId:guessDivision(l.name)}));
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);

  const fileRefLv=useRef(null);
  const fileRefRm=useRef(null);

  // ใช้ guessDivisionFromName จาก constants แทน (single source of truth)
  const guessDivision=(name)=>guessDivisionFromName(name);
  const addLv=async ()=>{
    const n=await uiPrompt("ชื่อระดับชั้น:");
    if(n){U.setLevels(p=>[...p,{id:gid(),name:n,divisionId:guessDivision(n)}]);st("เพิ่มสำเร็จ")}
  };
  const editLv=async (lv)=>{
    const n=await uiPrompt("แก้ไขชื่อระดับชั้น:",lv.name);
    if(n){U.setLevels(p=>p.map(l=>l.id===lv.id?{...l,name:n,divisionId:l.divisionId||guessDivision(n)}:l));st("แก้ไขสำเร็จ")}
  };
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

  return <div className="management-view"><LevelsManager S={S} U={U} st={st} divisions={DIVISIONS} actions={[
    ['นำเข้าระดับชั้น',()=>fileRefLv.current?.click()],['แบบฟอร์มระดับชั้น',templateLevels],['ส่งออกระดับชั้น',exportLevels],['นำเข้าห้องเรียน',()=>fileRefRm.current?.click()],['แบบฟอร์มห้องเรียน',templateRooms],['ส่งออกห้องเรียน',exportRooms]
  ]}/><input data-ui-control="true" hidden ref={fileRefLv} type="file" accept=".xlsx,.xls,.csv" onChange={importLevels}/><input data-ui-control="true" hidden ref={fileRefRm} type="file" accept=".xlsx,.xls,.csv" onChange={importRooms}/></div>;
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

  return <div className="management-view">
    <ManagementToolbar label="เพิ่มแผนการเรียน" onAdd={()=>{setEditId(null);setForm({name:'',subPlans:'',levelIds:[]});setModal(true)}} actions={[
      ['นำเข้า Excel',()=>fileRef.current?.click()],['ดาวน์โหลดแบบฟอร์ม',templatePlans],['ส่งออก Excel',exportPlans]
    ]}><p className="context-note">แผนที่ใช้ร่วมกันได้หลายระดับชั้น</p></ManagementToolbar>
    <input data-ui-control="true" hidden ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={importPlans}/>
    <RecordList rows={S.plans} placeholder="ค้นหาแผนการเรียน…" columns={[
      {key:'name',label:'แผนการเรียน',render:r=><strong>{r.name}</strong>},
      {key:'levels',label:'ระดับชั้น',render:r=>r.levelIds?.length?r.levelIds.map(id=>S.levels.find(l=>l.id===id)?.name).filter(Boolean).join(', '):'ใช้ได้ทุกระดับ'},
      {key:'subPlans',label:'สายรอง',render:r=>(r.subPlans||[]).join(', ')||'—'},
      {key:'rooms',label:'ห้องที่ใช้',render:r=>S.rooms.filter(x=>x.planId===r.id).length}
    ]} onEdit={openEdit} onDelete={async r=>{if(S.rooms.some(x=>x.planId===r.id)){st('แผนนี้ยังมีห้องเรียนใช้อยู่ กรุณาเปลี่ยนแผนของห้องก่อนลบ','error');return}if(await uiConfirm('ลบแผน '+r.name+'?')){U.setPlans(p=>p.filter(x=>x.id!==r.id));st('ลบแล้ว','warning')}}}/>
    <Modal open={modal} onClose={()=>{setModal(false);setEditId(null)}} title={editId?"แก้ไขแผนการเรียน":"เพิ่มแผนการเรียน"}>
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <div><label style={LS}>ชื่อแผน</label><input data-ui-control="true" style={IS} value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))} placeholder="วิทย์-คณิต"/></div>
        <div><label style={LS}>สายรอง (คอมม่า)</label><input data-ui-control="true" style={IS} value={form.subPlans} onChange={e=>setForm(p=>({...p,subPlans:e.target.value}))} placeholder="วิทย์สุขภาพ, วิศวะ"/></div>
        <div><label style={LS}>ใช้กับระดับชั้น</label>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{S.levels.map(lv=><button data-ui-control="true" key={lv.id} onClick={()=>toggleLv(lv.id)} style={{padding:"8px 16px",borderRadius:10,border:`2px solid ${form.levelIds.includes(lv.id)?"#DC2626":"#D1D5DB"}`,background:form.levelIds.includes(lv.id)?"#FEE2E2":"#fff",color:form.levelIds.includes(lv.id)?"#991B1B":"#374151",fontSize:13,fontWeight:600,cursor:"pointer"}}>{form.levelIds.includes(lv.id)?"✓ ":""}{lv.name}</button>)}</div>
        </div>
        <button data-ui-control="true" onClick={save} style={BS()}>{editId?"บันทึกการแก้ไข":"เพิ่ม"}</button>
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

  return <div className="management-view">
    <ManagementToolbar actions={[
      ['นำเข้า Excel',()=>fileRef.current?.click()],['ดาวน์โหลดแบบฟอร์ม',templateDepts],['ส่งออก Excel',exportDepts]
    ]}><form className="inline-add" onSubmit={e=>{e.preventDefault();if(!name.trim())return;U.setDepts(p=>[...p,{id:gid(),name:name.trim()}]);setName('');st('เพิ่มแล้ว')}}><input data-ui-control="true" required aria-label="ชื่อกลุ่มสาระใหม่" placeholder="ชื่อกลุ่มสาระใหม่" value={name} onChange={e=>setName(e.target.value)}/><button data-ui-control="true" className="primary">เพิ่มกลุ่มสาระ</button></form></ManagementToolbar>
    <input data-ui-control="true" hidden ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={importDepts}/>
    <RecordList rows={S.depts} placeholder="ค้นหากลุ่มสาระ…" columns={[{key:'name',label:'กลุ่มสาระ',render:r=><ColorBadge item={r}/>},{key:'teachers',label:'จำนวนครู',render:r=>S.teachers.filter(t=>t.departmentId===r.id).length},{key:'subjects',label:'จำนวนวิชา',render:r=>S.subjects.filter(t=>t.departmentId===r.id).length}]} onEdit={async r=>{const value=await uiPrompt('ชื่อกลุ่มสาระ',r.name);if(value?.trim()){U.setDepts(p=>p.map(x=>x.id===r.id?{...x,name:value.trim()}:x));st('บันทึกแล้ว')}}} onDelete={async r=>{if(S.teachers.some(x=>x.departmentId===r.id)||S.subjects.some(x=>x.departmentId===r.id)){st('กลุ่มสาระนี้ยังมีครูหรือวิชา กรุณาย้ายข้อมูลก่อนลบ','error');return}if(await uiConfirm('ลบ '+r.name+'?')){U.setDepts(p=>p.filter(x=>x.id!==r.id));st('ลบแล้ว','warning')}}}/>
  </div>;
}

/* ===== TEACHERS ===== */
function Teachers({S,U,st,gc}){
  const [modal,setModal]=useState(false);
  const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({prefix:"",firstName:"",lastName:"",teacherCode:"",departmentId:"",specialRoles:[],totalPeriods:0});
  const resetForm=()=>setForm({prefix:"",firstName:"",lastName:"",teacherCode:"",departmentId:"",specialRoles:[],totalPeriods:0});
  const [search,setSearch]=useState("");
  const fileRef=useRef(null);

  const save=()=>{
    if(!form.firstName.trim()||!form.departmentId){st("กรุณากรอกให้ครบ","error");return}
    if(!Number.isInteger(form.totalPeriods)||form.totalPeriods<0){st('จำนวนคาบต้องเป็นจำนวนเต็มตั้งแต่ 0 ขึ้นไป','error');return;}
    if(form.teacherCode&&S.teachers.some(t=>t.id!==editId&&t.teacherCode?.trim().toLowerCase()===form.teacherCode.trim().toLowerCase())){st('รหัสครูซ้ำกับรายการที่มีอยู่','error');return;}
    if(editId){
      U.setTeachers(p=>p.map(t=>t.id===editId?{...t,...form}:t));st("แก้ไขสำเร็จ");
    } else {
      U.setTeachers(p=>[...p,{id:gid(),...form}]);st("เพิ่มครูสำเร็จ");
    }
    setForm({prefix:"",firstName:"",lastName:"",teacherCode:"",departmentId:"",specialRoles:[],totalPeriods:0});setModal(false);setEditId(null);
  };

  const openEdit=(t)=>{setEditId(t.id);setForm({prefix:t.prefix,firstName:t.firstName,lastName:t.lastName,teacherCode:t.teacherCode||"",departmentId:t.departmentId,specialRoles:t.specialRoles||[],totalPeriods:t.totalPeriods||0});setModal(true)};
  const toggleRole=(rid)=>setForm(p=>({...p,specialRoles:p.specialRoles.includes(rid)?p.specialRoles.filter(r=>r!==rid):[...p.specialRoles,rid]}));

  // Import Excel/CSV — อัพเดทครูที่มีชื่อซ้ำ แทนที่จะเพิ่มใหม่
  const handleFile=async(e)=>{
    const f=e.target.files?.[0]; if(!f)return;
    let rows;
    if(f.name.endsWith('.csv')){const txt=await f.text();rows=parseCSV(txt);}
    else{rows=await readExcelFile(f);}
    if(!rows?.length){st("ไม่พบข้อมูล","error");return;}

    let added=0, updated=0;
    const newTeachers=[...S.teachers];
    rows.forEach(r=>{
      const prefix=String(r["คำนำหน้า"]||"").trim();
      const firstName=String(r["ชื่อ"]||"").trim();
      const lastName=String(r["นามสกุล"]||"").trim();
      const teacherCode=String(r["รหัสครู"]||"").trim();
      if(!firstName) return;

      const dept=S.depts.find(d=>d.name===String(r["กลุ่มสาระ"]||"").trim());
      const roles=[];
      const rs=String(r["หน้าที่พิเศษ"]||"");
      if(rs.includes("วิชาการ"))roles.push("academic");
      if(rs.includes("วินัย"))roles.push("discipline");

      // ตรวจว่ามีชื่อซ้ำหรือไม่ (firstName + lastName)
      const existIdx=newTeachers.findIndex(t=>
        t.firstName===firstName && t.lastName===lastName
      );
      if(existIdx>=0){
        // อัพเดทข้อมูลที่มีอยู่ — เพิ่มรหัสครูเป็นหลัก
        newTeachers[existIdx]={
          ...newTeachers[existIdx],
          ...(teacherCode?{teacherCode}:{}),
          ...(dept?{departmentId:dept.id}:{}),
          ...(roles.length?{specialRoles:roles}:{}),
          ...(r["คาบที่ได้รับ"]?{totalPeriods:parseInt(r["คาบที่ได้รับ"])||newTeachers[existIdx].totalPeriods}:{}),
        };
        updated++;
      } else {
        newTeachers.push({id:gid(),prefix,firstName,lastName,teacherCode,departmentId:dept?.id||"",specialRoles:roles,totalPeriods:parseInt(r["คาบที่ได้รับ"])||0});
        added++;
      }
    });
    U.setTeachers(newTeachers);
    st(`นำเข้าสำเร็จ: เพิ่มใหม่ ${added} คน, อัพเดท ${updated} คน`);
    e.target.value="";
  };

  const exportT=()=>{
    exportExcel(
      ["รหัสครู","คำนำหน้า","ชื่อ","นามสกุล","กลุ่มสาระ","หน้าที่พิเศษ","คาบที่ได้รับ"],
      S.teachers.map(t=>[
        t.teacherCode||"",
        t.prefix,t.firstName,t.lastName,
        S.depts.find(d=>d.id===t.departmentId)?.name||"",
        (t.specialRoles||[]).map(r=>SROLES.find(sr=>sr.id===r)?.name).filter(Boolean).join("/")||"ครูทั่วไป",
        t.totalPeriods||0
      ]),
      "รายชื่อครู_ดาราวิทยาลัย.xlsx","ครู"
    );
    st("Export สำเร็จ");
  };

  const downloadTemplate=()=>{
    exportExcel(
      ["รหัสครู","คำนำหน้า","ชื่อ","นามสกุล","กลุ่มสาระ","หน้าที่พิเศษ","คาบที่ได้รับ"],
      [["T001","นาย","สมชาย","ใจดี","วิทยาศาสตร์","ฝ่ายวิชาการ",18],
       ["T002","นางสาว","สมหญิง","รักเรียน","คณิตศาสตร์","ครูทั่วไป",20]],
      "Template_ครู.xlsx","Template"
    );
    st("ดาวน์โหลด Template");
  };

  // นับคาบจากตารางจริง (รองรับ coTeacherIds array) เหมือน teacherScheduledTotal ใน Scheduler
  const usedPeriods=(tid)=>{
    const seen=new Set();
    let c=0;
    Object.entries(S.schedule).forEach(([k,en])=>{
      const pts=k.split("_");
      en?.forEach(e=>{
        const coIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
        if(e.teacherId!==tid&&!coIds.includes(tid))return;
        const sub=S.subjects.find(s=>s.id===e.subjectId);
        const ca=sub?.consecutiveAllowed||0;
        if(ca===-1||ca===-2){
          const npKey=e.subjectId+"_"+pts[1]+"_"+pts[2];
          if(!seen.has(npKey)){seen.add(npKey);c++;}
        } else {c++;}
      });
    });
    return c;
  };

  const deleteTeacher=async (t)=>{
    const referenced=S.assigns.some(a=>a.teacherId===t.id)||Object.values(S.schedule).flat().some(e=>e.teacherId===t.id||(e.coTeacherIds||[]).includes(t.id)||e.coTeacherId===t.id);
    if(referenced){st('ครูคนนี้มีงานมอบหมายหรือคาบสอน กรุณาย้ายงานก่อนลบ','error');return;}
    if(!await uiConfirm('ลบ '+t.firstName+' '+t.lastName+' ออกจากรายชื่อครู?'))return;
    U.setTeachers(p=>p.filter(x=>x.id!==t.id));st('ลบครูแล้ว','warning');
  };
  return <div>
    <TeacherDirectory S={S} onAdd={()=>{setEditId(null);resetForm();setModal(true)}} onEdit={openEdit} onDelete={deleteTeacher} onImport={()=>fileRef.current?.click()} onExport={exportT} onTemplate={downloadTemplate}/>
    <input data-ui-control="true" ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:'none'}} onChange={handleFile}/>
    <Modal open={modal} onClose={()=>{setModal(false);setEditId(null)}} title={editId?"แก้ไขครู":"เพิ่มครู"}>
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <div style={{display:"grid",gridTemplateColumns:"100px 1fr 1fr",gap:12}}>
          <div><label style={LS}>คำนำหน้า</label><select data-ui-control="true" style={IS} value={form.prefix} onChange={e=>setForm(p=>({...p,prefix:e.target.value}))}><option value="">--</option><option>นาย</option><option>นาง</option><option>นางสาว</option></select></div>
          <div><label style={LS}>ชื่อ</label><input data-ui-control="true" style={IS} value={form.firstName} onChange={e=>setForm(p=>({...p,firstName:e.target.value}))}/></div>
          <div><label style={LS}>นามสกุล</label><input data-ui-control="true" style={IS} value={form.lastName} onChange={e=>setForm(p=>({...p,lastName:e.target.value}))}/></div>
        </div>
        <div><label style={LS}>รหัสครู (Username)</label><input data-ui-control="true" style={IS} value={form.teacherCode||""} onChange={e=>setForm(p=>({...p,teacherCode:e.target.value}))} placeholder="เช่น T001, prachya@dara.ac.th"/></div>
        <div><label style={LS}>กลุ่มสาระ</label><SearchSelect value={form.departmentId} onChange={v=>setForm(p=>({...p,departmentId:v}))} options={[{value:"",label:"--"},...S.depts.map(d=>({value:d.id,label:d.name}))]} placeholder="-- เลือกกลุ่มสาระ --"/></div>
        <div><label style={LS}>คาบที่ได้รับ (ต่อสัปดาห์)</label><input data-ui-control="true" type="number" min="0" style={IS} value={form.totalPeriods} onChange={e=>setForm(p=>({...p,totalPeriods:parseInt(e.target.value)||0}))}/></div>
        <div><label style={LS}>หน้าที่พิเศษ</label><div style={{display:"flex",gap:8}}>{SROLES.map(r=><button data-ui-control="true" key={r.id} onClick={()=>toggleRole(r.id)} style={{padding:"8px 16px",borderRadius:10,border:`2px solid ${form.specialRoles.includes(r.id)?"#DC2626":"#D1D5DB"}`,background:form.specialRoles.includes(r.id)?"#FEE2E2":"#fff",fontSize:13,fontWeight:600,cursor:"pointer"}}>{form.specialRoles.includes(r.id)?"✓ ":""}{r.name}</button>)}</div></div>
        <button data-ui-control="true" onClick={save} style={BS()}>{editId?"บันทึก":"เพิ่มครู"}</button>
      </div>
    </Modal>
  </div>;
}

/* ===== SPECIAL ROOMS (ห้องพิเศษ) ===== */
function SpecialRooms({S,U,st}){
  const [modal,setModal]=useState(false);
  const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({name:"",capacity:0,note:""});

  const save=()=>{
    if(!form.name.trim()){st("กรอกชื่อห้อง","error");return}
    if(editId){
      U.setSpecialRooms(p=>p.map(r=>r.id===editId?{...r,...form}:r));st("แก้ไขสำเร็จ");
    } else {
      U.setSpecialRooms(p=>[...p,{id:gid(),...form}]);st("เพิ่มห้องพิเศษสำเร็จ");
    }
    setForm({name:"",capacity:0,note:""});setModal(false);setEditId(null);
  };
  const openEdit=(r)=>{setEditId(r.id);setForm({name:r.name,capacity:r.capacity||0,note:r.note||""});setModal(true)};

  // นับวิชาที่ใช้ห้องนี้
  const subCount=(srId)=>S.subjects.filter(s=>s.specialRoomId===srId).length;

  return <div className="management-view">
    <ManagementToolbar label="เพิ่มห้องพิเศษ" onAdd={()=>{setEditId(null);setForm({name:'',capacity:0,note:''});setModal(true)}}><p className="context-note">ห้องที่ต้องตรวจการใช้ซ้อน เช่น ห้องแล็บและห้องคอมพิวเตอร์</p></ManagementToolbar>
    <RecordList rows={S.specialRooms} placeholder="ค้นหาห้องพิเศษ…" columns={[{key:'name',label:'ห้อง',render:r=><strong>{r.name}</strong>},{key:'capacity',label:'ความจุ (คน)',render:r=>r.capacity||'ไม่ระบุ'},{key:'note',label:'หมายเหตุ'},{key:'subjects',label:'วิชาที่ใช้',render:r=>subCount(r.id)}]} onEdit={openEdit} onDelete={async r=>{if(subCount(r.id)>0){st('มีวิชาใช้ห้องนี้อยู่ ลบไม่ได้','error');return}if(await uiConfirm('ลบห้อง '+r.name+'?')){U.setSpecialRooms(p=>p.filter(x=>x.id!==r.id));st('ลบแล้ว','warning')}}}/>
    <Modal open={modal} onClose={()=>{setModal(false);setEditId(null)}} title={editId?"แก้ไขห้องพิเศษ":"เพิ่มห้องพิเศษ"}>
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <div><label style={LS}>ชื่อห้อง</label><input data-ui-control="true" style={IS} value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))} placeholder="เช่น ห้องคอมพิวเตอร์ 1, ห้องแลบวิทย์"/></div>
        <div><label style={LS}>ความจุ (คน) — ไม่บังคับ</label><input data-ui-control="true" type="number" min="0" style={IS} value={form.capacity} onChange={e=>setForm(p=>({...p,capacity:parseInt(e.target.value)||0}))}/></div>
        <div><label style={LS}>หมายเหตุ</label><input data-ui-control="true" style={IS} value={form.note} onChange={e=>setForm(p=>({...p,note:e.target.value}))} placeholder="รายละเอียดเพิ่มเติม"/></div>
        <button data-ui-control="true" onClick={save} style={BS()}>{editId?"บันทึก":"เพิ่มห้องพิเศษ"}</button>
      </div>
    </Modal>
  </div>;
}

/* ===== SUBJECTS ===== */
function Subjects({S,U,st,gc}){
  const [modal,setModal]=useState(false);
  const [editId,setEditId]=useState(null);
  const BLANK={code:"",name:"",shortName:"",credits:1,periodsPerWeek:1,departmentId:"",levelId:"",specialRoomId:"",consecutiveAllowed:0,allDepts:false};
  const [form,setForm]=useState(BLANK);
  const fileRef=useRef(null);
  const [filterLv,setFilterLv]=useState("");
  const [filterDept,setFilterDept]=useState("");
  const [search,setSearch]=useState("");

  const save=()=>{
    if(!form.name||!form.departmentId||!form.levelId){st("กรอกให้ครบ","error");return}
    if(editId){U.setSubjects(p=>p.map(s=>s.id===editId?{...s,...form}:s));st("แก้ไขสำเร็จ")}
    else{U.setSubjects(p=>[...p,{id:gid(),...form}]);st("เพิ่มวิชาสำเร็จ")}
    setForm(BLANK);setModal(false);setEditId(null);
  };
  const openEdit=(s)=>{
    setEditId(s.id);
    setForm({code:s.code||"",name:s.name||"",shortName:s.shortName||"",credits:s.credits||1,periodsPerWeek:s.periodsPerWeek||1,
      departmentId:s.departmentId||"",levelId:s.levelId||"",
      specialRoomId:s.specialRoomId||"",consecutiveAllowed:s.consecutiveAllowed||0});
    setModal(true);
  };

  const handleFile=async(e)=>{
    const f=e.target.files?.[0]; if(!f)return;
    let rows;
    if(f.name.endsWith('.csv')){const txt=await f.text();rows=parseCSV(txt);}
    else{rows=await readExcelFile(f);}
    if(!rows?.length){st("ไม่พบข้อมูล","error");return;}

    let added=0, updated=0;
    const newSubs=[...S.subjects];
    rows.forEach(r=>{
      const code=String(r["รหัสวิชา"]||"").trim();
      const name=String(r["ชื่อวิชา"]||"").trim();
      if(!name) return;
      const dept=S.depts.find(d=>d.name===String(r["กลุ่มสาระ"]||"").trim());
      const lv=S.levels.find(l=>l.name===String(r["ระดับชั้น"]||"").trim());
      const subData={code,name,shortName:String(r["ชื่อย่อ"]||"").trim(),credits:parseFloat(r["หน่วยกิต"])||1,periodsPerWeek:parseInt(r["คาบ/สัปดาห์"])||1,departmentId:dept?.id||"",levelId:lv?.id||"",specialRoomId:"",consecutiveAllowed:0};

      // ตรวจซ้ำด้วยชื่อ หรือรหัสวิชา
      const existIdx=newSubs.findIndex(s=>
        (code&&s.code===code)||(s.name===name&&s.levelId===(lv?.id||""))
      );
      if(existIdx>=0){
        newSubs[existIdx]={...newSubs[existIdx],...subData};
        updated++;
      } else {
        newSubs.push({id:gid(),...subData});
        added++;
      }
    });
    U.setSubjects(newSubs);
    st(`นำเข้าสำเร็จ: เพิ่มใหม่ ${added} วิชา, อัพเดท ${updated} วิชา`);
    e.target.value="";
  };

  const exportS=()=>{exportExcel(["รหัสวิชา","ชื่อวิชา","ชื่อย่อ","หน่วยกิต","คาบ/สัปดาห์","กลุ่มสาระ","ระดับชั้น"],S.subjects.map(s=>[s.code,s.name,s.shortName||"",s.credits,s.periodsPerWeek,S.depts.find(d=>d.id===s.departmentId)?.name||"",S.levels.find(l=>l.id===s.levelId)?.name||""]),"รายวิชา_ดาราวิทยาลัย.xlsx","วิชา");st("Export สำเร็จ")};
  const downloadTemplate=()=>{exportExcel(["รหัสวิชา","ชื่อวิชา","ชื่อย่อ","หน่วยกิต","คาบ/สัปดาห์","กลุ่มสาระ","ระดับชั้น"],[["ว33201","ฟิสิกส์ 3","ฟิสิกส์",1.5,3,"วิทยาศาสตร์","ม.6"],["ค33101","คณิตศาสตร์พื้นฐาน","คณิต",1,2,"คณิตศาสตร์","ม.6"]],"Template_วิชา.xlsx","Template");st("ดาวน์โหลด Template")};

  // กรอง + จัดกลุ่ม level → dept
  const filtered=S.subjects.filter(s=>{
    if(filterLv&&s.levelId!==filterLv)return false;
    if(filterDept&&s.departmentId!==filterDept)return false;
    if(search&&!s.name.includes(search)&&!s.code.includes(search))return false;
    return true;
  });
  // เรียงตาม level name → dept name
  const sortedLevels=[...S.levels].sort((a,b)=>a.name.localeCompare(b.name,"th"));
  const groups=sortedLevels.map(lv=>{
    const lvSubs=filtered.filter(s=>s.levelId===lv.id);
    if(!lvSubs.length)return null;
    const deptGroups=S.depts.map(dept=>{
      const ds=lvSubs.filter(s=>s.departmentId===dept.id);
      return ds.length?{dept,subs:ds}:null;
    }).filter(Boolean);
    // วิชาที่ไม่มีกลุ่มสาระ
    const noDept=lvSubs.filter(s=>!S.depts.find(d=>d.id===s.departmentId));
    if(noDept.length)deptGroups.push({dept:null,subs:noDept});
    return{lv,deptGroups};
  }).filter(Boolean);
  // วิชาที่ไม่มีระดับชั้น
  const noLevel=filtered.filter(s=>!S.levels.find(l=>l.id===s.levelId));

  const SubCard=({sub})=>{
    const dept=S.depts.find(d=>d.id===sub.departmentId);
    const c=dept?gc(dept.id):{bg:"#6B7280",lt:"#F3F4F6",tx:"#374151"};
    const sr=S.specialRooms.find(r=>r.id===sub.specialRoomId);
    return<div data-ui-surface="true" style={{background:"#fff",borderRadius:12,overflow:"hidden",boxShadow:"0 2px 12px rgba(0,0,0,0.06)",borderLeft:"3px solid "+c.bg}}>
      <div style={{padding:"12px 14px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:10,color:"#9CA3AF",fontWeight:600}}>{sub.code}</div>
            <h4 style={{fontSize:14,fontWeight:700,marginTop:1,wordBreak:"break-word"}}>{sub.name}</h4>
            {sub.shortName&&<div style={{fontSize:11,color:"#6B7280",marginTop:1}}>ชื่อย่อ: <strong>{sub.shortName}</strong></div>}
          </div>
          <div style={{display:"flex",gap:4,flexShrink:0,marginLeft:8}}>
            <button data-ui-control="true" onClick={()=>openEdit(sub)} style={{background:"none",border:"none",cursor:"pointer",color:"#2563EB"}}><Icon name="edit" size={13}/></button>
            <button data-ui-control="true" onClick={()=>{U.setSubjects(p=>p.filter(x=>x.id!==sub.id));st("ลบแล้ว","warning")}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444"}}><Icon name="trash" size={13}/></button>
          </div>
        </div>
        <div style={{display:"flex",gap:4,marginTop:8,flexWrap:"wrap"}}>
          <span style={{background:"#F3F4F6",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>{sub.credits} หน่วยกิต</span>
          <span style={{background:"#F3F4F6",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>{sub.periodsPerWeek} คาบ/สป.</span>
          {sr&&<span style={{background:"#EDE9FE",color:"#5B21B6",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>📍{sr.name}</span>}
          {sub.consecutiveAllowed>0&&<span style={{background:"#FEF3C7",color:"#92400E",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>⚡{sub.consecutiveAllowed}คาบติด</span>}
          {sub.consecutiveAllowed===-1&&<span style={{background:"#EFF6FF",color:"#1E40AF",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>🔀NP</span>}
          {sub.consecutiveAllowed===-2&&<span style={{background:"#FDF4FF",color:"#6B21A8",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>🏛️เศรษฐ-วิศวะ</span>}
          {sub.allDepts&&<span style={{background:"#FEF9C3",color:"#92400E",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:700}}>🏫 ทุกกลุ่มสาระ</span>}
        </div>
      </div>
    </div>;
  };

  return <div className="management-view">
    <ManagementToolbar label="เพิ่มวิชา" onAdd={()=>{setEditId(null);setForm(BLANK);setModal(true)}} actions={[
      ['นำเข้า Excel',()=>fileRef.current?.click()],['ดาวน์โหลดแบบฟอร์ม',downloadTemplate],['ส่งออก Excel',exportS]
    ]}><p className="context-note">ค้นหารหัสหรือชื่อวิชา แล้วกรองตามระดับชั้นและกลุ่มสาระ</p></ManagementToolbar>
    <input data-ui-control="true" hidden ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile}/>
    <RecordList rows={S.subjects} placeholder="ค้นหารหัสหรือชื่อวิชา…" searchText={r=>(r.code||'')+' '+r.name+' '+(r.shortName||'')} filters={[
      {key:'level',label:'ทุกระดับชั้น',options:S.levels.map(l=>({value:l.id,label:l.name})),match:(r,v)=>r.levelId===v},
      {key:'dept',label:'ทุกกลุ่มสาระ',options:S.depts.map(d=>({value:d.id,label:d.name})),match:(r,v)=>r.departmentId===v}
    ]} columns={[
      {key:'code',label:'รหัสวิชา'}, {key:'name',label:'รายวิชา',render:r=><strong style={{color:departmentTone(S.depts.find(d=>d.id===r.departmentId)).ink}}>{r.name}</strong>},
      {key:'level',label:'ระดับ',render:r=><ColorBadge item={S.levels.find(l=>l.id===r.levelId)} kind="level"/>},
      {key:'dept',label:'กลุ่มสาระ',render:r=><ColorBadge item={S.depts.find(d=>d.id===r.departmentId)}/>},
      {key:'periodsPerWeek',label:'คาบ / สัปดาห์'},
      {key:'condition',label:'เงื่อนไข',render:r=>[r.consecutiveAllowed>1?r.consecutiveAllowed+' คาบติด':r.consecutiveAllowed===-1?'NP':r.consecutiveAllowed===-2?'เศรษฐ–วิศวะ':'',r.specialRoomId?'ห้องพิเศษ':'',r.allDepts?'สอนร่วมทุกสาระ':''].filter(Boolean).join(' · ')||'ปกติ'}
    ]} onEdit={openEdit} onDelete={async r=>{if(S.assigns.some(a=>a.subjectId===r.id)||Object.values(S.schedule).flat().some(e=>e.subjectId===r.id)){st('วิชานี้มีงานมอบหมายหรือคาบสอน กรุณาย้ายข้อมูลก่อนลบ','error');return}if(await uiConfirm('ลบวิชา '+r.name+'?')){U.setSubjects(p=>p.filter(x=>x.id!==r.id));st('ลบแล้ว','warning')}}}/>
    <Modal open={modal} onClose={()=>{setModal(false);setEditId(null)}} title={editId?"แก้ไขวิชา":"เพิ่มวิชา"}>
      <div style={{display:"flex",flexDirection:"column",gap:14}}>
        <div><label style={LS}>รหัสวิชา</label><input data-ui-control="true" style={IS} value={form.code} onChange={e=>setForm(p=>({...p,code:e.target.value}))} placeholder="ว33202"/></div>
        <div><label style={LS}>ชื่อวิชาเต็ม</label><input data-ui-control="true" style={IS} value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))} placeholder="ฟิสิกส์ 4"/></div>
        <div><label style={LS}>ชื่อย่อ <span style={{fontWeight:400,color:"#9CA3AF"}}>(แสดงบนการ์ดและตารางพิมพ์)</span></label><input data-ui-control="true" style={IS} value={form.shortName||""} onChange={e=>setForm(p=>({...p,shortName:e.target.value}))} placeholder="ฟิสิกส์"/></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          <div><label style={LS}>หน่วยกิต</label><input data-ui-control="true" type="number" min="0.5" step="0.5" style={IS} value={form.credits} onChange={e=>setForm(p=>({...p,credits:parseFloat(e.target.value)||0}))}/></div>
          <div><label style={LS}>คาบ/สัปดาห์</label><input data-ui-control="true" type="number" min="1" style={IS} value={form.periodsPerWeek} onChange={e=>setForm(p=>({...p,periodsPerWeek:parseInt(e.target.value)||1}))}/></div>
        </div>
        <div><label style={LS}>ระดับชั้น</label><SearchSelect value={form.levelId} onChange={v=>setForm(p=>({...p,levelId:v}))} options={[{value:"",label:"--"},...S.levels.map(l=>({value:l.id,label:l.name}))]} placeholder="-- เลือกระดับชั้น --"/></div>
        <div><label style={LS}>กลุ่มสาระ</label><SearchSelect value={form.departmentId} onChange={v=>setForm(p=>({...p,departmentId:v}))} options={[{value:"",label:"--"},...S.depts.map(d=>({value:d.id,label:d.name}))]} placeholder="-- เลือกกลุ่มสาระ --"/></div>
        <div><label style={LS}>ห้องพิเศษ (ถ้าต้องใช้) — ตรวจ conflict ข้ามทุกห้อง</label>
          <select data-ui-control="true" style={IS} value={form.specialRoomId} onChange={e=>setForm(p=>({...p,specialRoomId:e.target.value}))}>
            <option value="">-- ไม่ใช้ห้องพิเศษ --</option>
            {S.specialRooms.map(r=><option key={r.id} value={r.id}>{r.name}</option>)}
          </select>
        </div>
        <div><label style={LS}>คาบติดต่อกัน / คาบพิเศษ</label>
          <select data-ui-control="true" style={IS} value={form.consecutiveAllowed} onChange={e=>setForm(p=>({...p,consecutiveAllowed:parseInt(e.target.value)||0}))}>
            <option value={0}>ปกติ — ห้ามซ้ำ 2 คาบ/วัน</option>
            <option value={2}>อนุญาต 2 คาบติด</option>
            <option value={3}>อนุญาต 3 คาบติด</option>
            <option value={4}>อนุญาต 4 คาบติด</option>
            <option value={-1}>NP — ลงคาบเดียวกันคนละห้องได้ (นับครู 1 คาบ)</option>
            <option value={-2}>ห้องเศรษฐศาสตร์วิศวกรรม — 2 ห้องพร้อมกัน 2 คาบติด ครูหลายคน</option>
          </select>
          {form.consecutiveAllowed===-1&&<div style={{marginTop:6,padding:"8px 12px",background:"#EFF6FF",border:"1px solid #BFDBFE",borderRadius:8,fontSize:12,color:"#1E40AF"}}>
            📌 วิชานี้สามารถวางในคาบเดียวกันได้หลายห้อง (เช่น ม.5/1, ม.5/5, ม.5/6 คาบเดียวกัน) และระบบจะนับเป็น <strong>1 คาบ</strong> สำหรับครูผู้สอน
          </div>}
          {form.consecutiveAllowed===-2&&<div style={{marginTop:6,padding:"8px 12px",background:"#FDF4FF",border:"1px solid #E9D5FF",borderRadius:8,fontSize:12,color:"#6B21A8"}}>
            📌 <strong>ห้องเศรษฐศาสตร์วิศวกรรม:</strong> 2 ห้องเรียนพร้อมกัน วางคาบเดียวกันคนละห้องได้ · ต้องวาง 2 คาบติดกัน · ครูทุกคนในการ์ดนับคาบตามนี้ · นับแต่ละคาบ 1 ครั้ง (ไม่ซ้ำข้ามห้อง)
          </div>}
        </div>

        {/* allDepts flag */}
        <label style={{display:"flex",alignItems:"flex-start",gap:12,padding:"12px 14px",borderRadius:12,border:`2px solid ${form.allDepts?"#D97706":"#E5E7EB"}`,background:form.allDepts?"#FFFBEB":"#F9FAFB",cursor:"pointer"}}>
          <input data-ui-control="true" type="checkbox" checked={!!form.allDepts} onChange={e=>setForm(p=>({...p,allDepts:e.target.checked}))} style={{marginTop:2,accentColor:"#D97706",flexShrink:0}}/>
          <div>
            <div style={{fontSize:13,fontWeight:700,color:form.allDepts?"#92400E":"#374151"}}>🏫 วิชาที่ทุกกลุ่มสาระสอนร่วมกัน</div>
            <div style={{fontSize:11,color:"#6B7280",marginTop:2}}>เช่น กิจกรรมพัฒนาผู้เรียน, ลูกเสือ — ครูต่างสาระสามารถ assign วิชานี้ได้ และระบบจะตรวจการชนของครูทุกคนที่สอนวิชานี้</div>
          </div>
        </label>

        <button data-ui-control="true" onClick={save} style={BS()}>{editId?"บันทึก":"เพิ่มวิชา"}</button>
      </div>
    </Modal>
  </div>;
}

/* ===== PERSONAL LOCK PANEL ===== */
function PersonalLockPanel({teacher,U,st,sel}){
  const [plDay,setPlDay]=useState("");
  const [plPeriods,setPlPeriods]=useState([]);
  const [plReason,setPlReason]=useState("");
  const personalLocks=teacher.personalLocks||[];

  const addLock=()=>{
    if(!plDay||!plPeriods.length){st("เลือกวันและคาบ","error");return;}
    U.setTeachers(prev=>prev.map(t=>{
      if(t.id!==sel)return t;
      const existing=t.personalLocks||[];
      const idx=existing.findIndex(l=>l.day===plDay&&(l.reason||"ส่วนตัว")===(plReason||"ส่วนตัว"));
      if(idx>=0){
        const merged=[...new Set([...existing[idx].periods,...plPeriods])].sort((a,b)=>a-b);
        const upd=[...existing];upd[idx]={...existing[idx],periods:merged};
        return{...t,personalLocks:upd};
      }
      return{...t,personalLocks:[...existing,{id:gid(),day:plDay,periods:[...plPeriods].sort((a,b)=>a-b),reason:plReason||"ส่วนตัว"}]};
    }));
    setPlDay("");setPlPeriods([]);setPlReason("");
    st("เพิ่มคาบล็อกสำเร็จ");
  };

  const removeLock=(id)=>{
    U.setTeachers(prev=>prev.map(t=>t.id!==sel?t:{...t,personalLocks:(t.personalLocks||[]).filter(l=>l.id!==id)}));
    st("ลบคาบล็อกแล้ว","warning");
  };

  return(
    <div data-ui-surface="true" data-work-panel="true" className="content-card" style={{background:"#fff",borderRadius:14,padding:20,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",marginBottom:20}}>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:16}}>
        <span style={{fontSize:20}}>🔒</span>
        <h3 style={{fontSize:15,fontWeight:700,margin:0}}>คาบล็อกส่วนตัว</h3>
        <span style={{fontSize:12,color:"#6B7280"}}>— {teacher.prefix}{teacher.firstName} {teacher.lastName}</span>
      </div>
      <div data-work-panel="true" style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end",marginBottom:16,padding:"14px 16px",background:"#FFF7ED",borderRadius:12,border:"1px solid #FED7AA"}}>
        <div style={{flex:"1 1 130px"}}>
          <label style={LS}>วัน</label>
          <select data-ui-control="true" style={IS} value={plDay} onChange={e=>setPlDay(e.target.value)}>
            <option value="">-- เลือกวัน --</option>
            {DAYS.map(d=><option key={d}>{d}</option>)}
          </select>
        </div>
        <div style={{flex:"2 1 300px"}}>
          <label style={LS}>คาบ (เลือกได้หลายคาบ)</label>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {PERIODS.map(p=>(
              <button data-ui-control="true" key={p.id}
                onClick={()=>setPlPeriods(prev=>prev.includes(p.id)?prev.filter(x=>x!==p.id):[...prev,p.id])}
                style={{width:44,height:44,borderRadius:8,border:`2px solid ${plPeriods.includes(p.id)?"#DC2626":"#D1D5DB"}`,background:plPeriods.includes(p.id)?"#DC2626":"#fff",color:plPeriods.includes(p.id)?"#fff":"#374151",fontSize:14,fontWeight:700,cursor:"pointer"}}>
                {p.id}
              </button>
            ))}
          </div>
        </div>
        <div style={{flex:"1 1 160px"}}>
          <label style={LS}>เหตุผล (ไม่บังคับ)</label>
          <input data-ui-control="true" style={IS} value={plReason} onChange={e=>setPlReason(e.target.value)} placeholder="ติดธุระ, อบรม ฯ" onKeyDown={e=>e.key==="Enter"&&addLock()}/>
        </div>
        <button data-ui-control="true" onClick={addLock} style={{...BS("#C2410C"),flexShrink:0}}>+ เพิ่มล็อก</button>
      </div>
      {personalLocks.length===0
        ?<div style={{textAlign:"center",color:"#9CA3AF",fontSize:13,padding:"12px 0"}}>ยังไม่มีคาบล็อกส่วนตัว</div>
        :<div style={{display:"flex",flexDirection:"column",gap:8}}>
          {[...personalLocks].sort((a,b)=>DAYS.indexOf(a.day)-DAYS.indexOf(b.day)).map(pl=>(
            <div data-work-panel="true" key={pl.id} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 14px",background:"#FFF7ED",borderRadius:10,border:"1px solid #FED7AA"}}>
              <span style={{fontSize:16}}>🔒</span>
              <div style={{flex:1}}>
                <span style={{fontWeight:700,color:"#C2410C",fontSize:13}}>วัน{pl.day}</span>
                <span style={{color:"#6B7280",fontSize:12,marginLeft:8}}>คาบ {(pl.periods||[]).join(", ")}</span>
                {pl.reason&&<span style={{marginLeft:8,fontSize:11,background:"#FFEDD5",color:"#9A3412",padding:"1px 8px",borderRadius:20,fontWeight:600}}>{pl.reason}</span>}
              </div>
              <button data-ui-control="true" onClick={()=>removeLock(pl.id)} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",padding:4}}><Icon name="trash" size={14}/></button>
            </div>
          ))}
        </div>
      }
    </div>
  );
}

/* ===== ASSIGNMENTS ===== */
function Assigns({S,U,st,gc}){
  const [selDept,setSelDept]=useState("");
  const [sel,setSel]=useState("");
  const [modal,setModal]=useState(false);
  const [form,setForm]=useState({subjectId:"",roomIds:[],totalPeriods:0});
  const [modalDeptFilter,setModalDeptFilter]=useState("");
  const [basket,setBasket]=useState([]); // [{subjectId, roomIds, totalPeriods}] รอบันทึก
  const fileRefA=useRef(null);
  const [editAssign,setEditAssign]=useState(null);
  const [editForm,setEditForm]=useState({roomIds:[],totalPeriods:0});
  const deptTeachers=selDept?S.teachers.filter(t=>t.departmentId===selDept):[];
  const teacher=S.teachers.find(t=>t.id===sel);
  const asgns=S.assigns.filter(a=>a.teacherId===sel);
  // วิชาครูร่วม: assignment ที่ครูนี้เป็น co-teacher (ผ่าน schedule entries)
  const coAsgnsIdsA = new Set(
    Object.entries(S.schedule).flatMap(([,en])=>
      (en||[]).filter(e=>{
        const coIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
        return coIds.includes(sel)&&e.teacherId!==sel;
      }).map(e=>e.assignmentId)
    ).filter(Boolean)
  );
  const coAsgnsA=S.assigns.filter(a=>coAsgnsIdsA.has(a.id)&&!asgns.find(x=>x.id===a.id));
  // นับคาบจริงจาก schedule (รองรับ NP/-2 deduplicate และ coTeacherIds)
  const scheduledUsed=(tid)=>{
    const seen=new Set(); let c=0;
    Object.entries(S.schedule).forEach(([k,en])=>{
      const pts=k.split("_");
      const day=pts[pts.length-2]; const per=pts[pts.length-1];
      en?.forEach(e=>{
        const coIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
        if(e.teacherId!==tid&&!coIds.includes(tid))return;
        const sub=S.subjects.find(s=>s.id===e.subjectId);
        const ca=sub?.consecutiveAllowed||0;
        if(ca===-1||ca===-2){const k2=e.subjectId+"_"+day+"_"+per;if(!seen.has(k2)){seen.add(k2);c++;}}
        else c++;
      });
    });
    return c;
  };
  const totalScheduled=scheduledUsed(sel);   // คาบที่ลงตารางแล้วจริง (รวมครูร่วม, deduplicate NP/-2)
  // totalAssigned = คาบที่ครูต้องสอน (นับตาม periodsPerWeek จริง ไม่ × จำนวนห้อง)
  const totalAssigned=(()=>{
    const seen=new Set(); let c=0;
    // นับจาก assignment ตัวเอง (deduplicate วิชา NP/-2 ด้วย subjectId_roomId)
    asgns.forEach(a=>{
      const sub=S.subjects.find(s=>s.id===a.subjectId);
      const ca=sub?.consecutiveAllowed||0;
      if(ca===-1||ca===-2){
        // NP/-2: นับแค่ periodsPerWeek ต่อวิชา (ไม่คูณห้อง)
        const k2="own_"+a.subjectId;
        if(!seen.has(k2)){seen.add(k2);c+=sub?.periodsPerWeek||a.totalPeriods;}
      } else {
        c+=a.totalPeriods;
      }
    });
    // นับคาบครูร่วม (deduplicate NP/-2 เหมือนกัน)
    const seenCo=new Set();
    Object.entries(S.schedule).forEach(([k,en])=>{
      const pts=k.split("_");
      en?.forEach(e=>{
        const coIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
        if(!coIds.includes(sel)||e.teacherId===sel)return;
        if(!coAsgnsIdsA.has(e.assignmentId))return;
        const sub=S.subjects.find(s=>s.id===e.subjectId);
        const ca=sub?.consecutiveAllowed||0;
        if(ca===-1||ca===-2){const k2=e.subjectId+"_"+pts[pts.length-2]+"_"+pts[pts.length-1];if(!seenCo.has(k2)){seenCo.add(k2);c++;}}
        else c++;
      });
    });
    return c;
  })();
  const totalUsed=totalScheduled;
  const teacherQuota=teacher?.totalPeriods||0;
  const remaining=teacherQuota-totalAssigned;
  const notScheduled=totalAssigned-totalScheduled; // มอบหมายแล้วแต่ยังไม่ลงตาราง

  // แสดงวิชาทุกสาระเรียงตามกลุ่มสาระ พร้อม label บอกสาระ
  const teacherDeptSubs = S.subjects.slice().sort((a,b)=>{
    const da = S.depts.find(d=>d.id===a.departmentId)?.name||"zzz";
    const db = S.depts.find(d=>d.id===b.departmentId)?.name||"zzz";
    if(da!==db) return da.localeCompare(db,"th");
    return (a.code||"").localeCompare(b.code||"");
  });

  // ข้อ 3: when subject selected, show rooms of that level only
  const selSub=S.subjects.find(s=>s.id===form.subjectId);
  const filteredRooms=selSub?S.rooms.filter(r=>r.levelId===selSub.levelId):S.rooms;

  // Export assignments ทุกคน → Excel
  const exportAssigns=()=>{
    const rows=[];
    S.assigns.forEach(a=>{
      const t=S.teachers.find(x=>x.id===a.teacherId);
      const sub=S.subjects.find(s=>s.id===a.subjectId);
      const rooms=a.roomIds.map(rid=>S.rooms.find(r=>r.id===rid)?.name||"").join(",");
      rows.push([
        t?`${t.prefix}${t.firstName} ${t.lastName}`:"",
        S.depts.find(d=>d.id===t?.departmentId)?.name||"",
        sub?.code||"",sub?.name||"",
        a.totalPeriods||0,
        rooms,
      ]);
    });
    exportExcel(["ครู","กลุ่มสาระ","รหัสวิชา","ชื่อวิชา","คาบที่มอบหมาย","ห้องเรียน"],rows,"มอบหมายงานครู.xlsx","มอบหมาย");
    st("Export สำเร็จ");
  };

  // Import assignments จาก Excel
  const importAssigns=async(e)=>{
    const f=e.target.files?.[0];if(!f)return;
    let rows;
    if(f.name.endsWith('.csv')){const txt=await f.text();rows=parseCSV(txt);}
    else{rows=await readExcelFile(f);}
    if(!rows?.length){st("ไม่พบข้อมูล","error");return;}
    const ns=[];
    const failLog=[];
    rows.forEach(r=>{
      const tName=String(r["ครู"]||"").trim();
      const subCode=String(r["รหัสวิชา"]||"").trim();
      const subName=String(r["ชื่อวิชา"]||"").trim();
      const roomNames=String(r["ห้องเรียน"]||"").split(",").map(x=>x.trim()).filter(Boolean);
      const periods=parseInt(r["คาบที่มอบหมาย"])||1;
      if(!tName||!subName||!roomNames.length)return;
      const normalize=(n)=>n.replace(/^ม\./,"").replace(/\s+/g,"");
      // ค้นหาครู
      const t=S.teachers.find(x=>{
        const full=`${x.prefix}${x.firstName} ${x.lastName}`.replace(/\s+/g," ");
        const noPrefix=`${x.firstName} ${x.lastName}`.replace(/\s+/g," ");
        const tn=tName.replace(/\s+/g," ");
        return full===tn||noPrefix===tn||x.firstName===tn;
      });
      if(!t){
        // หาครูที่ชื่อใกล้เคียง (firstName หรือ lastName มี substring)
        const hint=S.teachers.find(x=>tName.includes(x.firstName)||x.firstName.includes(tName.split(" ")[0]));
        failLog.push({row:`${tName} / ${subName}`, reason:`ไม่เจอครู "${tName}"`, hint:hint?`ในระบบมี: "${hint.prefix}${hint.firstName} ${hint.lastName}"`:""});
        return;
      }
      // ค้นหาวิชา
      const sub=S.subjects.find(s=>
        (subCode&&s.code===subCode)||s.name===subName||(s.shortName&&s.shortName===subName)
      );
      if(!sub){
        const hint=S.subjects.find(s=>s.name.includes(subName.substring(0,4))||subName.includes(s.name.substring(0,4)));
        failLog.push({row:`${tName} / ${subName}`, reason:`ไม่เจอวิชา "${subName}"`, hint:hint?`ในระบบมี: "${hint.name}"`:""});
        return;
      }
      // ค้นหาห้อง
      const roomIds=roomNames.map(n=>{
        const norm=normalize(n);
        return (S.rooms.find(rm=>rm.name===n)||S.rooms.find(rm=>normalize(rm.name)===norm))?.id;
      }).filter(Boolean);
      if(!roomIds.length){
        const sampleRooms=S.rooms.slice(0,5).map(rm=>rm.name).join(", ");
        failLog.push({row:`${tName} / ${subName}`, reason:`ไม่เจอห้อง "${roomNames[0]}"`, hint:`ตัวอย่างห้องในระบบ: ${sampleRooms}`});
        return;
      }
      const exists=S.assigns.find(a=>a.teacherId===t.id&&a.subjectId===sub.id&&JSON.stringify(a.roomIds.sort())===JSON.stringify(roomIds.sort()));
      if(!exists) ns.push({id:gid(),teacherId:t.id,subjectId:sub.id,roomIds,totalPeriods:periods});
    });
    if(ns.length) U.setAssigns(p=>[...p,...ns]);
    // แสดง diagnostic popup
    if(failLog.length>0){
      const lines=failLog.slice(0,8).map(f=>`• ${f.reason}${f.hint?" → "+f.hint:""}`).join("\n");
      const extra=failLog.length>8?`\n... และอีก ${failLog.length-8} รายการ`:"";
      await uiAlert(`${ns.length>0?`นำเข้าสำเร็จ ${ns.length} รายการ\n\n`:""}ข้าม ${failLog.length} รายการ:\n${lines}${extra}\n\n💡 วิธีแก้: กด Export ก่อน แล้วใช้ไฟล์นั้นเป็นแม่แบบ`);
    } else if(ns.length){
      st(`นำเข้า ${ns.length} รายการ`);
    }
    e.target.value="";
  };

  return <div className="assignment-page">
    <AssignmentHeader S={S} teacher={teacher} department={selDept} onDepartment={v=>{setSelDept(v);setSel("")}} onTeacher={setSel} onAdd={()=>{setForm({subjectId:"",roomIds:[],totalPeriods:0});setBasket([]);setModalDeptFilter(teacher?.departmentId||"");setModal(true)}} actions={[['ส่งออกงานมอบหมาย',exportAssigns],['นำเข้างานมอบหมาย',()=>fileRefA.current?.click()]]} quota={teacherQuota} assigned={totalAssigned} scheduled={totalScheduled} pending={notScheduled} remaining={remaining}/>
    <input ref={fileRefA} type="file" accept=".xlsx,.xls,.csv" hidden onChange={importAssigns}/>
    {teacher&&<div>
      <RecordList rows={asgns} placeholder="ค้นหาวิชาหรือห้องที่มอบหมาย…" searchText={a=>{const sub=S.subjects.find(s=>s.id===a.subjectId);return (sub?.code||'')+' '+(sub?.name||'')+' '+(a.roomIds||[]).map(id=>S.rooms.find(r=>r.id===id)?.name).join(' ')}} columns={[
  {key:'subject',label:'วิชา',render:a=>{const sub=S.subjects.find(s=>s.id===a.subjectId);return <><small className="row-subtitle">{sub?.code}</small><strong>{subDisplayName(sub)||'ไม่พบวิชา'}</strong></>}},
  {key:'rooms',label:'ห้องเรียน',render:a=>(a.roomIds||[]).map(id=>S.rooms.find(r=>r.id===id)?.name).join(', ')},
  {key:'assigned',label:'มอบหมาย (คาบ)',render:a=>{const sub=S.subjects.find(s=>s.id===a.subjectId);return [-1,-2].includes(sub?.consecutiveAllowed)?sub.periodsPerWeek||a.totalPeriods:a.totalPeriods}},
  {key:'scheduled',label:'ลงตารางแล้ว',render:a=>{const sub=S.subjects.find(s=>s.id===a.subjectId);const cells=Object.entries(S.schedule).flatMap(([k,es])=>es.filter(e=>e.assignmentId===a.id).map(()=>k.split('_').slice(-2).join('_')));return [-1,-2].includes(sub?.consecutiveAllowed)?new Set(cells).size:cells.length}}
]} onEdit={a=>{setEditAssign(a);setEditForm({roomIds:[...(a.roomIds||[])],totalPeriods:a.totalPeriods})}} onDelete={async a=>{if(!await uiConfirm('ลบงานมอบหมายนี้ รวมถึงคาบที่ลงตารางไว้ด้วย?'))return;U.setAssigns(p=>p.filter(x=>x.id!==a.id));U.setSchedule(prev=>Object.fromEntries(Object.entries(prev).map(([k,en])=>[k,en.filter(e=>e.assignmentId!==a.id)]).filter(([,en])=>en.length)));st('ลบงานมอบหมายและคาบที่เกี่ยวข้องแล้ว','warning')}}/>
    </div>}
    {editAssign&&(()=>{
      const eSub=S.subjects.find(s=>s.id===editAssign.subjectId);
      const eRooms=eSub?.levelId?S.rooms.filter(r=>r.levelId===eSub.levelId):S.rooms;
      const autoTP=(eSub?.periodsPerWeek||1)*Math.max(editForm.roomIds.length,1);
      return(
        <Modal open={!!editAssign} onClose={()=>setEditAssign(null)} title={"✏️ แก้ไข — "+(eSub?.code||"")+" "+(eSub?.name||"")}>
          <div style={{display:"flex",flexDirection:"column",gap:16}}>
            <div data-ui-surface="true" data-work-panel="true" style={{background:"#F9FAFB",borderRadius:10,padding:"10px 14px"}}>
              <div style={{fontSize:14,fontWeight:700}}>{eSub?.code} — {eSub?.name}</div>
            </div>
            <div>
              <label style={LS}>ห้องเรียน <span style={{fontSize:11,color:"#9CA3AF"}}>(กดเลือก/ยกเลิก)</span></label>
              <div style={{display:"flex",flexWrap:"wrap",gap:6,maxHeight:180,overflowY:"auto"}}>
                {eRooms.map(rm=>{
                  const on=editForm.roomIds.includes(rm.id);
                  return <button data-ui-control="true" key={rm.id} onClick={()=>setEditForm(p=>({...p,roomIds:on?p.roomIds.filter(r=>r!==rm.id):[...p.roomIds,rm.id]}))}
                    style={{padding:"5px 14px",borderRadius:20,border:"2px solid "+(on?"#DC2626":"#D1D5DB"),background:on?"#FEE2E2":"#fff",color:on?"#991B1B":"#374151",fontSize:12,fontWeight:on?700:400,cursor:"pointer"}}>{on?"✓ ":""}{rm.name}</button>;
                })}
              </div>
              <div style={{display:"flex",gap:6,marginTop:8}}>
                <button data-ui-control="true" onClick={()=>setEditForm(p=>({...p,roomIds:eRooms.map(r=>r.id)}))} style={{fontSize:11,color:"#DC2626",background:"none",border:"1px solid #FECACA",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>เลือกทั้งหมด</button>
                <button data-ui-control="true" onClick={()=>setEditForm(p=>({...p,roomIds:[]}))} style={{fontSize:11,color:"#6B7280",background:"none",border:"1px solid #E5E7EB",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>ล้าง</button>
                <span style={{fontSize:11,color:"#6B7280"}}>เลือก {editForm.roomIds.length} ห้อง</span>
              </div>
            </div>
            <div>
              <label style={LS}>จำนวนคาบ/สัปดาห์</label>
              <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                <input data-ui-control="true" type="number" min="0" style={{...IS,width:100}} value={editForm.totalPeriods} onChange={e=>setEditForm(p=>({...p,totalPeriods:parseInt(e.target.value)||0}))}/>
                {eSub?.periodsPerWeek&&editForm.roomIds.length>0&&<button data-ui-control="true" onClick={()=>setEditForm(p=>({...p,totalPeriods:autoTP}))} style={{fontSize:11,background:"#EFF6FF",color:"#1D4ED8",border:"1px solid #BFDBFE",borderRadius:8,padding:"4px 12px",cursor:"pointer"}}>อัตโนมัติ: {eSub.periodsPerWeek}×{editForm.roomIds.length}={autoTP}</button>}
              </div>
            </div>
            <div style={{display:"flex",gap:10}}>
              <button data-ui-control="true" onClick={()=>setEditAssign(null)} style={{...BO(),flex:1}}>ยกเลิก</button>
              <button data-ui-control="true" disabled={editForm.roomIds.length===0} onClick={()=>{
                const finalTP=editForm.totalPeriods||autoTP||1;
                U.setAssigns(p=>p.map(x=>x.id===editAssign.id?{...x,roomIds:editForm.roomIds,totalPeriods:finalTP}:x));
                setEditAssign(null);st("แก้ไขสำเร็จ ✓");
              }} style={{...BS(),flex:2,opacity:editForm.roomIds.length===0?0.4:1}}>💾 บันทึก</button>
            </div>
          </div>
        </Modal>
      );
    })()}
    {teacher&&<details className="advanced-options"><summary>เวลาที่ครูไม่ว่าง / คาบล็อกส่วนตัว</summary><PersonalLockPanel teacher={teacher} U={U} st={st} sel={sel}/></details>}
    <Modal open={modal} onClose={()=>{setModal(false);setBasket([]);}} title={`มอบหมายวิชา — ${teacher?.prefix||""}${teacher?.firstName||""}`}>
      <div style={{display:"flex",flexDirection:"column",gap:14}}>

        {/* ── ตะกร้าวิชาที่เพิ่มแล้ว ── */}
        {basket.length>0&&(
          <div data-work-panel="true" style={{background:"#F0FDF4",border:"1.5px solid #BBF7D0",borderRadius:12,padding:"10px 14px"}}>
            <div style={{fontSize:12,fontWeight:700,color:"#065F46",marginBottom:8}}>
              🛒 วิชาที่รอบันทึก ({basket.length} รายการ)
            </div>
            {basket.map((b,bi)=>{
              const bs=S.subjects.find(s=>s.id===b.subjectId);
              return(
                <div data-ui-surface="true" key={bi} style={{display:"flex",alignItems:"center",gap:8,marginBottom:4,background:"#fff",borderRadius:8,padding:"5px 10px",border:"1px solid #D1FAE5"}}>
                  <div style={{flex:1,fontSize:12}}>
                    <span style={{fontWeight:700,color:"#065F46"}}>{bs?.code}</span>
                    <span style={{color:"#374151",marginLeft:6}}>{bs?.name}</span>
                    <span style={{color:"#9CA3AF",marginLeft:6,fontSize:11}}>
                      {b.roomIds.map(rid=>S.rooms.find(r=>r.id===rid)?.name).join(", ")}
                      {b.totalPeriods>0?` · ${b.totalPeriods} คาบ`:""}
                    </span>
                  </div>
                  <button data-ui-control="true" onClick={()=>setBasket(p=>p.filter((_,i)=>i!==bi))}
                    style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",fontSize:14,padding:0,flexShrink:0}}>✕</button>
                </div>
              );
            })}
          </div>
        )}

        {/* ── ฟอร์มเพิ่มวิชาใหม่ ── */}
        <div data-ui-surface="true" data-work-panel="true" style={{background:"#F9FAFB",borderRadius:12,padding:"14px 16px",border:"1px solid #E5E7EB"}}>
          <div style={{fontSize:12,fontWeight:700,color:"#374151",marginBottom:10}}>➕ เพิ่มวิชา</div>

          {/* เลือกสาระ */}
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6}}>
            <label style={{...LS,marginBottom:0,fontSize:12}}>
              วิชา
              <span style={{fontSize:10,color:"#6B7280",fontWeight:400,marginLeft:5}}>
                {modalDeptFilter===teacher?.departmentId
                  ? `(${S.depts.find(d=>d.id===teacher?.departmentId)?.name||"สาระหลัก"})`
                  : modalDeptFilter ? `(${S.depts.find(d=>d.id===modalDeptFilter)?.name})`
                  : "(ทุกสาระ)"}
              </span>
            </label>
            {modalDeptFilter===teacher?.departmentId
              ? <button data-ui-control="true" onClick={()=>{setModalDeptFilter("");setForm(p=>({...p,subjectId:"",roomIds:[]}));}}
                  style={{fontSize:10,padding:"2px 10px",borderRadius:20,border:"1.5px solid #7C3AED",background:"#F5F3FF",color:"#5B21B6",cursor:"pointer",fontWeight:600}}>
                  📚 สาระอื่น
                </button>
              : <button data-ui-control="true" onClick={()=>{setModalDeptFilter(teacher?.departmentId||"");setForm(p=>({...p,subjectId:"",roomIds:[]}));}}
                  style={{fontSize:10,padding:"2px 10px",borderRadius:20,border:"1.5px solid #DC2626",background:"#FEF2F2",color:"#991B1B",cursor:"pointer",fontWeight:600}}>
                  ⭐ สาระหลัก
                </button>
            }
          </div>

          {/* pills กลุ่มสาระอื่น */}
          {modalDeptFilter!==teacher?.departmentId&&(
            <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:8}}>
              {[{id:"",name:"ทั้งหมด"},...S.depts.filter(d=>d.id!==teacher?.departmentId)].map(d=>(
                <button data-ui-control="true" key={d.id}
                  onClick={()=>{setModalDeptFilter(d.id);setForm(p=>({...p,subjectId:"",roomIds:[]}));}}
                  style={{fontSize:10,padding:"2px 9px",borderRadius:20,border:`1.5px solid ${modalDeptFilter===d.id?"#2563EB":"#E5E7EB"}`,background:modalDeptFilter===d.id?"#EFF6FF":"#fff",color:modalDeptFilter===d.id?"#1E40AF":"#6B7280",cursor:"pointer",fontWeight:modalDeptFilter===d.id?700:400}}>
                  {d.name}
                </button>
              ))}
            </div>
          )}

          <SearchSelect
            value={form.subjectId}
            onChange={v=>setForm(p=>({...p,subjectId:v,roomIds:[],totalPeriods:0}))}
            options={[{value:"",label:"-- เลือกวิชา --"},...teacherDeptSubs
              .filter(s=>{
                // ข้ามวิชาที่อยู่ใน basket แล้ว
                if(basket.some(b=>b.subjectId===s.id)) return false;
                if(modalDeptFilter===teacher?.departmentId) return s.departmentId===teacher?.departmentId;
                if(modalDeptFilter==="") return s.departmentId!==teacher?.departmentId;
                return s.departmentId===modalDeptFilter;
              })
              .map(s=>{
                const dname=S.depts.find(d=>d.id===s.departmentId)?.name||"";
                const lname=S.levels.find(l=>l.id===s.levelId)?.name||"";
                const isSame=s.departmentId===teacher?.departmentId;
                return{value:s.id,label:`${!isSame?"["+dname+"] ":""}${s.code} — ${s.name} (${lname})`};
              })
            ]}
            placeholder="-- เลือกวิชา --"
          />

          {/* ห้องเรียน */}
          {form.subjectId&&(
            <div style={{marginTop:10}}>
              <label style={{...LS,fontSize:12}}>ห้อง</label>
              <div style={{display:"flex",gap:6,flexWrap:"wrap",maxHeight:160,overflowY:"auto"}}>
                {filteredRooms.map(rm=>(
                  <button data-ui-control="true" key={rm.id}
                    onClick={()=>setForm(p=>({...p,roomIds:p.roomIds.includes(rm.id)?p.roomIds.filter(r=>r!==rm.id):[...p.roomIds,rm.id]}))}
                    style={{padding:"5px 12px",borderRadius:8,border:`2px solid ${form.roomIds.includes(rm.id)?"#DC2626":"#D1D5DB"}`,background:form.roomIds.includes(rm.id)?"#FEE2E2":"#fff",fontSize:12,fontWeight:600,cursor:"pointer"}}>
                    {form.roomIds.includes(rm.id)?"✓ ":""}{rm.name}
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* คาบรวม (optional) */}
          {form.subjectId&&form.roomIds.length>0&&(
            <div style={{marginTop:8,display:"flex",alignItems:"center",gap:8}}>
              <label style={{...LS,marginBottom:0,fontSize:12,flexShrink:0}}>คาบรวม (0=อัตโนมัติ)</label>
              <input data-ui-control="true" type="number" min="0" style={{...IS,width:90}} value={form.totalPeriods}
                onChange={e=>setForm(p=>({...p,totalPeriods:parseInt(e.target.value)||0}))}/>
            </div>
          )}

          {/* ปุ่ม + เพิ่มใส่ตะกร้า */}
          <button data-ui-control="true"
            disabled={!form.subjectId||!form.roomIds.length}
            onClick={()=>{
              if(!form.subjectId||!form.roomIds.length) return;
              setBasket(p=>[...p,{subjectId:form.subjectId,roomIds:form.roomIds,totalPeriods:form.totalPeriods}]);
              setForm({subjectId:"",roomIds:[],totalPeriods:0});
            }}
            style={{...BS("#059669"),marginTop:10,opacity:(!form.subjectId||!form.roomIds.length)?0.4:1,fontSize:13}}>
            + เพิ่มใส่รายการ
          </button>
        </div>

        {/* ── ปุ่มบันทึกทั้งหมด ── */}
        <div style={{display:"flex",gap:10}}>
          <button data-ui-control="true" onClick={()=>{setModal(false);setBasket([]);}} style={{...BO(),flex:1}}>ยกเลิก</button>
          <button data-ui-control="true"
            disabled={basket.length===0}
            onClick={()=>{
              if(!basket.length){st("ยังไม่มีวิชาในรายการ","error");return;}
              const newAssigns=basket.map(b=>{
                const sub=S.subjects.find(s=>s.id===b.subjectId);
                const tp=b.totalPeriods||(sub?.periodsPerWeek||1)*b.roomIds.length;
                return{id:gid(),teacherId:sel,subjectId:b.subjectId,roomIds:b.roomIds,totalPeriods:tp};
              });
              U.setAssigns(p=>[...p,...newAssigns]);
              setBasket([]);
              setForm({subjectId:"",roomIds:[],totalPeriods:0});
              setModal(false);
              st(`มอบหมาย ${newAssigns.length} วิชาสำเร็จ`);
            }}
            style={{...BS(),flex:2,opacity:basket.length===0?0.4:1}}>
            💾 บันทึก {basket.length>0?`(${basket.length} วิชา)`:""}
          </button>
        </div>

      </div>
    </Modal>
  </div>;
}

/* ===== HOMEROOM SETTINGS ===== */
function HomeroomSettings({S,U,st}){
  const [editId,setEditId]=useState(null); // roomId ที่กำลัง edit
  const [form,setForm]=useState({homeroom1:"",homeroom2:"",homeroomCo:""});
  const [filterLevel,setFilterLevel]=useState("");

  const openEdit=(rm)=>{
    setEditId(rm.id);
    setForm({homeroom1:rm.homeroom1||"",homeroom2:rm.homeroom2||"",homeroomCo:rm.homeroomCo||""});
  };
  const save=()=>{
    U.setRooms(p=>p.map(r=>r.id===editId?{...r,...form}:r));
    setEditId(null);
    st("บันทึกครูประจำชั้นแล้ว");
  };

  const filteredRooms=S.rooms.filter(r=>!filterLevel||r.levelId===filterLevel);
  // เรียงตามระดับชั้น → ชื่อห้อง
  const sorted=[...filteredRooms].sort((a,b)=>{
    const la=S.levels.find(l=>l.id===a.levelId)?.name||"";
    const lb=S.levels.find(l=>l.id===b.levelId)?.name||"";
    if(la!==lb) return la.localeCompare(lb,"th");
    return a.name.localeCompare(b.name,"th");
  });

  const teacherOptions=[{value:"",label:"-- ไม่ระบุ --"},...S.teachers.map(t=>({value:t.prefix+t.firstName+" "+t.lastName,label:t.prefix+t.firstName+" "+t.lastName}))];

  return <div className="management-view"><p className="context-note" style={{marginBottom:18}}>ค้นหาห้องเพื่อกำหนดครูประจำชั้นและครูร่วม</p><RecordList rows={S.rooms} placeholder="ค้นหาห้องหรือชื่อครูประจำชั้น…" filters={[
    {key:'level',label:'ทุกระดับชั้น',options:S.levels.map(l=>({value:l.id,label:l.name})),match:(r,v)=>r.levelId===v},
    {key:'status',label:'ทุกสถานะ',options:[{value:'missing',label:'ยังไม่มีครูประจำชั้น'},{value:'assigned',label:'กำหนดครูแล้ว'}],match:(r,v)=>v==='missing'?!r.homeroom1&&!r.homeroom2:!!r.homeroom1||!!r.homeroom2}
  ]} columns={[{key:'name',label:'ห้องเรียน',render:r=><strong>{r.name}</strong>},{key:'level',label:'ระดับชั้น',render:r=><ColorBadge item={S.levels.find(l=>l.id===r.levelId)} kind="level"/>},{key:'homeroom1',label:'ครูประจำชั้น 1'},{key:'homeroom2',label:'ครูประจำชั้น 2'},{key:'homeroomCo',label:'ครูร่วม'}]} onEdit={openEdit}/>
  {editId&&<EditDialog title={'ครูประจำชั้น '+(S.rooms.find(r=>r.id===editId)?.name||'')} onClose={()=>setEditId(null)} onSave={save}>{[['homeroom1','ครูประจำชั้น 1'],['homeroom2','ครูประจำชั้น 2'],['homeroomCo','ครูร่วม']].map(([key,label])=><label className="wide" key={key}>{label}<select data-ui-control="true" value={form[key]} onChange={e=>setForm({...form,[key]:e.target.value})}>{teacherOptions.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}</select></label>)}</EditDialog>}</div>;
}

/* ===== MEETINGS ===== */
function Meetings({S,U,st,gc}){
  const [tab,setTab]=useState("dept");   // "dept" | "custom"

  // ── ฟอร์ม: คาบล็อคกลุ่มสาระ (เดิม — 1 วัน หลายคาบ) ──
  const [deptForm,setDeptForm]=useState({departmentId:"",day:"",periods:[]});

  // ── ฟอร์ม: คาบล็อคแผนก (ใหม่ — หลายวัน หลายคาบ + ชื่อ) ──
  const BLANK_CUSTOM={departmentId:"",name:"",slots:[]}; // slots: [{day,period}]
  const [cusForm,setCusForm]=useState(BLANK_CUSTOM);

  const toggleSlot=(day,pid)=>{
    setCusForm(prev=>{
      const exists=prev.slots.find(s=>s.day===day&&s.period===pid);
      return{...prev,slots:exists
        ?prev.slots.filter(s=>!(s.day===day&&s.period===pid))
        :[...prev.slots,{day,period:pid}]};
    });
  };
  const slotActive=(day,pid)=>!!cusForm.slots.find(s=>s.day===day&&s.period===pid);

  const saveDept=()=>{
    if(!deptForm.departmentId||!deptForm.day||!deptForm.periods.length){st("กรอกให้ครบ","error");return;}
    U.setMeetings(p=>[...p,{id:gid(),...deptForm}]);
    setDeptForm({departmentId:"",day:"",periods:[]});
    st("เพิ่มสำเร็จ");
  };

  const saveCustom=()=>{
    if(!cusForm.name||!cusForm.slots.length){st("กรอกชื่อและเลือกคาบ","error");return;}
    // type:"custom" ไม่ผูกกับ departmentId → ล็อคทุกคน
    U.setMeetings(p=>[...p,{id:gid(),departmentId:"all",name:cusForm.name,slots:cusForm.slots,type:"custom"}]);
    setCusForm(BLANK_CUSTOM);
    st("เพิ่มคาบล็อคสำเร็จ");
  };

  // แยก meetings ตาม type
  const deptMeetings=S.meetings.filter(m=>!m.type||m.type==="dept");
  const customMeetings=S.meetings.filter(m=>m.type==="custom");

  const TAB_STYLE=(active)=>({
    padding:"9px 20px",fontWeight:700,fontSize:13,cursor:"pointer",border:"none",fontFamily:"inherit",
    background:active?CRED:"transparent",color:active?"#fff":"#6B7280",
    borderBottom:active?"2px solid "+CRED:"2px solid transparent",transition:"all 0.15s",
  });

  return <div style={{animation:"fadeIn 0.3s"}}>
    {/* Tab bar */}
    <div style={{display:"flex",borderBottom:"2px solid #F3F4F6",marginBottom:20}}>
      <button data-ui-control="true" style={TAB_STYLE(tab==="dept")} onClick={()=>setTab("dept")}>ประชุมกลุ่มสาระ</button>
      <button data-ui-control="true" style={TAB_STYLE(tab==="custom")} onClick={()=>setTab("custom")}>คาบล็อกส่วนกลาง</button>
    </div>

    {/* ── Tab 1: คาบล็อคกลุ่มสาระ เดิม ── */}
    {tab==="dept"&&<>
      <details className="advanced-options creation-panel"><summary>เพิ่มคาบประชุมกลุ่มสาระ</summary><div data-ui-surface="true" className="content-card" style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",marginBottom:24,maxWidth:600}}>
        <h3 style={{fontSize:16,fontWeight:700,marginBottom:16}}>เพิ่มคาบล็อคกลุ่มสาระ</h3>
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          <div><label style={LS}>กลุ่มสาระ</label>
            <SearchSelect value={deptForm.departmentId} onChange={v=>setDeptForm(p=>({...p,departmentId:v}))} options={[{value:"",label:"--"},...S.depts.map(d=>({value:d.id,label:d.name}))]} placeholder="-- เลือกกลุ่มสาระ --"/>
          </div>
          <div><label style={LS}>วัน</label>
            <select data-ui-control="true" style={IS} value={deptForm.day} onChange={e=>setDeptForm(p=>({...p,day:e.target.value}))}>
              <option value="">--</option>{DAYS.map(d=><option key={d}>{d}</option>)}
            </select>
          </div>
          <div><label style={LS}>คาบ</label>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              {PERIODS.map(p=><button data-ui-control="true" key={p.id}
                onClick={()=>setDeptForm(prev=>({...prev,periods:prev.periods.includes(p.id)?prev.periods.filter(x=>x!==p.id):[...prev.periods,p.id]}))}
                style={{width:48,height:48,borderRadius:10,border:`2px solid ${deptForm.periods.includes(p.id)?"#DC2626":"#D1D5DB"}`,background:deptForm.periods.includes(p.id)?"#DC2626":"#fff",color:deptForm.periods.includes(p.id)?"#fff":"#374151",fontSize:16,fontWeight:700,cursor:"pointer"}}>
                {p.id}
              </button>)}
            </div>
          </div>
          <button data-ui-control="true" onClick={saveDept} style={BS()}>เพิ่มคาบล็อค</button>
        </div>
      </div></details>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:16}}>
        {deptMeetings.map(m=>{
          const dept=S.depts.find(d=>d.id===m.departmentId);
          const c=dept?gc(dept.id):{bg:"#6B7280"};
          return<div data-ui-surface="true" key={m.id} style={{background:"#fff",borderRadius:14,borderLeft:`4px solid ${c.bg}`,padding:16,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
            <div style={{display:"flex",justifyContent:"space-between"}}>
              <div>
                <h4 style={{fontSize:15,fontWeight:700}}>{dept?.name}</h4>
                <div style={{fontSize:13,color:"#6B7280",marginTop:4}}>วัน{m.day} — คาบ {(m.periods||[]).slice().sort().join(", ")}</div>
              </div>
              <button data-ui-control="true" onClick={()=>{U.setMeetings(p=>p.filter(x=>x.id!==m.id));st("ลบแล้ว","warning")}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444"}}><Icon name="trash" size={14}/></button>
            </div>
          </div>;
        })}
      </div>
    </>}

    {/* ── Tab 2: คาบล็อคแผนก หลายวันหลายคาบ ── */}
    {tab==="custom"&&<>
      <details className="advanced-options creation-panel"><summary>เพิ่มคาบล็อกหลายวัน</summary><div data-ui-surface="true" className="content-card" style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",marginBottom:24}}>
        <h3 style={{fontSize:16,fontWeight:700,marginBottom:16}}>เพิ่มคาบล็อคแผนก</h3>
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          {/* ชื่อ */}
          <div>
            <label style={LS}>ชื่อคาบล็อค</label>
            <input data-ui-control="true" style={{...IS,maxWidth:400}} value={cusForm.name} onChange={e=>setCusForm(p=>({...p,name:e.target.value}))} placeholder="เช่น ประชุมวิชาการ, อบรม, สอบกลางภาค"/>
          </div>

          {/* ตาราง grid วัน × คาบ เลือกได้หลายช่อง */}
          <div>
            <label style={LS}>เลือกวัน × คาบ (คลิกเพื่อเลือก/ยกเลิก)</label>
            <div style={{overflowX:"auto"}}>
              <table data-ui-table="true" style={{borderCollapse:"collapse",minWidth:500}}>
                <thead>
                  <tr>
                    <th style={{padding:"8px 12px",background:"#F3F4F6",fontSize:12,fontWeight:700,color:"#374151",border:"1px solid #E5E7EB",minWidth:70}}>วัน \ คาบ</th>
                    {PERIODS.map(p=>(
                      <th key={p.id} style={{padding:"6px 8px",background:"#F3F4F6",fontSize:12,fontWeight:700,color:"#374151",border:"1px solid #E5E7EB",textAlign:"center",minWidth:52}}>
                        <div>{p.id}</div>
                        <div style={{fontSize:9,fontWeight:400,color:"#9CA3AF"}}>{p.time.split("-")[0]}</div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {DAYS.map((day,di)=>(
                    <tr key={day} style={{background:di%2===0?"#fff":"#FAFAFA"}}>
                      <td style={{padding:"8px 12px",fontWeight:700,fontSize:13,color:"#374151",border:"1px solid #E5E7EB",background:"#F9FAFB"}}>{day}</td>
                      {PERIODS.map(p=>{
                        const active=slotActive(day,p.id);
                        return(
                          <td key={p.id}
                            onClick={()=>toggleSlot(day,p.id)}
                            style={{padding:"6px 4px",border:"1px solid #E5E7EB",textAlign:"center",cursor:"pointer",
                              background:active?"#DC2626":"transparent",
                              transition:"background 0.1s"}}
                          >
                            {active&&<span style={{color:"#fff",fontSize:14,fontWeight:700}}>✓</span>}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {cusForm.slots.length>0&&(
              <div style={{marginTop:8,fontSize:12,color:"#6B7280"}}>
                เลือกแล้ว {cusForm.slots.length} ช่อง:&nbsp;
                {DAYS.filter(d=>cusForm.slots.some(s=>s.day===d)).map(d=>(
                  <span key={d} style={{marginRight:8}}>
                    <strong>{d}</strong> คาบ {cusForm.slots.filter(s=>s.day===d).map(s=>s.period).sort((a,b)=>a-b).join(",")}
                  </span>
                ))}
                <button data-ui-control="true" onClick={()=>setCusForm(p=>({...p,slots:[]}))} style={{marginLeft:8,fontSize:11,color:"#EF4444",background:"none",border:"none",cursor:"pointer"}}>ล้างทั้งหมด</button>
              </div>
            )}
          </div>
          <button data-ui-control="true" onClick={saveCustom} style={BS()}>เพิ่มคาบล็อค</button>
        </div>
      </div></details>

      {/* รายการ custom locks */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(320px,1fr))",gap:16}}>
        {customMeetings.map(m=>{
          const slotsByDay=DAYS.map(day=>{
            const ps=(m.slots||[]).filter(s=>s.day===day).map(s=>s.period).sort((a,b)=>a-b);
            return ps.length?{day,periods:ps}:null;
          }).filter(Boolean);
          return<div data-ui-surface="true" key={m.id} style={{background:"#fff",borderRadius:14,borderLeft:"4px solid #DC2626",padding:16,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div style={{flex:1}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                  <h4 style={{fontSize:15,fontWeight:700}}>{m.name}</h4>
                  <span style={{fontSize:11,background:"#FEE2E2",color:"#991B1B",padding:"1px 8px",borderRadius:20,fontWeight:600}}>🏫 ทุกกลุ่มสาระ</span>
                </div>
                <div style={{display:"flex",flexDirection:"column",gap:3}}>
                  {slotsByDay.map(({day,periods})=>(
                    <div key={day} style={{fontSize:12,color:"#374151"}}>
                      <span style={{fontWeight:700,color:"#6B7280",minWidth:60,display:"inline-block"}}>{day}</span>
                      <span>คาบ {periods.join(", ")}</span>
                    </div>
                  ))}
                </div>
                <div style={{marginTop:6,fontSize:11,color:"#9CA3AF"}}>{(m.slots||[]).length} ช่องรวม</div>
              </div>
              <button data-ui-control="true" onClick={()=>{U.setMeetings(p=>p.filter(x=>x.id!==m.id));st("ลบแล้ว","warning")}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",flexShrink:0}}><Icon name="trash" size={14}/></button>
            </div>
          </div>;
        })}
        {customMeetings.length===0&&<div style={{color:"#9CA3AF",fontSize:13,padding:"20px 0"}}>ยังไม่มีคาบล็อคแผนก</div>}
      </div>
    </>}
  </div>;
}

/* ===== EMPTY STATE HELPER ===== */
function EmptyState({icon,title}){
  return <div data-ui-surface="true" className="content-card" style={{background:"#fff",borderRadius:14,padding:60,textAlign:"center"}}>
    <div style={{fontSize:48,marginBottom:16}}>{icon}</div>
    <h3 style={{fontSize:18,fontWeight:700,color:"#374151"}}>{title}</h3>
  </div>;
}

/* ===== SCHEDULER ENTRY CARD (top-level เพื่อกัน React recreate) ===== */
function SchedulerEntryCard({entry,cellKey,lk,cellCount,selT,mode,S,U,gc,setDrag,setCoM}){
  const sub=S.subjects.find(s=>s.id===entry.subjectId),teacher=S.teachers.find(t=>t.id===entry.teacherId);
  const coIds=entry.coTeacherIds?.length?entry.coTeacherIds:(entry.coTeacherId?[entry.coTeacherId]:[]);
  const dimmed=mode==='teacher'&&!!selT&&entry.teacherId!==selT&&!coIds.includes(selT);
  const tone=departmentTone(S.depts.find(d=>d.id===sub?.departmentId));
  return <article className={'lesson-tile '+(dimmed?'other-teacher ':'')+(lk?'is-locked ':'')+(cellCount>1?'is-stacked':'')} style={{'--lesson-accent':tone.ink,'--lesson-bg':tone.bg}}
    draggable={!lk&&!dimmed}
    onDragStart={e=>{if(dimmed||lk){e.preventDefault();return;}e.stopPropagation();const parts=cellKey.split('_');setDrag({fromKey:cellKey,fromRoomId:parts.slice(0,-2).join('_'),entry});}}
    onDragEnd={()=>setDrag(null)}>
    <div className="lesson-code"><span>{sub?.code||'ไม่ระบุรหัส'}</span>{lk&&<span className="locked-label"><Icon name="lock" size={12}/>ล็อก</span>}</div>
    <h4 title={subDisplayName(sub)}>{subDisplayName(sub)||'ไม่พบรายวิชา'}</h4>
    <p className="lesson-teacher" title={[teacher?.firstName,...coIds.map(id=>S.teachers.find(t=>t.id===id)?.firstName)].filter(Boolean).join(', ')}>{teacher?.firstName||'ไม่พบครู'}{coIds.length>0?' + '+coIds.length+' ครูร่วม':''}</p>
    {dimmed?<span className="lesson-context">คาบของครูท่านอื่น</span>:<div className="lesson-actions">
      {!lk&&<><button data-ui-control="true" title="นำคาบนี้ออกจากตาราง" aria-label={'นำ '+(sub?.name||'วิชา')+' ออกจากตาราง'} onClick={()=>U.setSchedule(prev=>({...prev,[cellKey]:(prev[cellKey]||[]).filter(e=>e.id!==entry.id)}))}><Icon name="x" size={14}/></button><button data-ui-control="true" title="จัดการครูร่วม" aria-label="จัดการครูร่วม" onClick={()=>setCoM({key:cellKey,entryId:entry.id})}><Icon name="users" size={14}/></button></>}
      <button data-ui-control="true" title={lk?'ปลดล็อกคาบ':'ล็อกคาบนี้'} aria-label={lk?'ปลดล็อกคาบ':'ล็อกคาบนี้'} onClick={()=>U.setLocks(prev=>({...prev,[cellKey]:!lk}))}><Icon name={lk?'unlock':'lock'} size={14}/></button>
    </div>}
  </article>;
}

/* ===== SCHEDULER ===== */
function Scheduler({S,U,st,gc,isSavingRef,fsReadyRef,fsSave}){
  const [mode,setMode]=useState("teacher");
  const [selDept,setSelDept]=useState("");
  const [selT,setSelT]=useState(()=>S.teachers[0]?.id||"");
  const [showWeekly,setShowWeekly]=useState(false);
  const [selRoom,setSelRoom]=useState("");
  const [drag,setDrag]=useState(null);
  const dragRef=useRef(null);  // ref สำหรับอ่านใน handleDrop กัน stale/race condition
  const setDragBoth=(v)=>{setDrag(v);dragRef.current=v;};
  const [coM,setCoM]=useState(null);   // {key, entryId} — modal บนการ์ดที่วางแล้ว
  const [coS,setCoS]=useState("");
  const [coDept,setCoDept]=useState("");
  const [cardCoM,setCardCoM]=useState(null); // assignId — modal ครูร่วม (เดิม)
  const [showGearId,setShowGearId]=useState(null); // assignId — gear panel inline
  const [cardCoS,setCardCoS]=useState("");
  const [cardCoDept,setCardCoDept]=useState("");
  const [cardCoMap,setCardCoMap]=useState({}); // {assignId: [teacherId, ...]} สูงสุด 4 ครูร่วม
  const [bundleMap,setBundleMap]=useState({}); // {assignId: [{assignId,teacherId},...]} วิชาที่สอนคาบเดียวกัน
  const [showBundleM,setShowBundleM]=useState(null);
  const [bundleSelSub,setBundleSelSub]=useState("");
  const [bundleSelTeacher,setBundleSelTeacher]=useState("");
  const [autoRunning,setAutoRunning]=useState(false);
  const [proposal,setProposal]=useState(null);
  const [previousAuto,setPreviousAuto]=useState(null);
  const autoBaseRef=useRef(null);
  const [autoResult,setAutoResult]=useState(null); // {placed, skipped, details}
  const [showAutoModal, setShowAutoModal] = useState(false);
  const [autoOpts, setAutoOpts] = useState({
    mode:        "remaining",   // "remaining" | "full"
    allowNormal: true,          // วิชาปกติ (ไม่มี consecutive)
    allowConsec: false,         // วิชาคาบติด (consecutive ≥ 2)
    allowNP:     false,         // วิชา NP (−1)
    allowSR:     false,         // วิชาห้องพิเศษ
    spreadDay:   true,          // กระจายไม่ให้วิชาเดียวอยู่วันเดียวกัน 2 คาบ (default เปิด)
    noFirstLast: true,          // ไม่วางคาบ 1 + คาบ 7 วันเดียวกัน (วิชาเดิม)
    maxConsecTeacher: 0,        // 0 = ไม่จำกัด, 1/2/3/4 = ห้ามครูสอนติดกันเกิน N คาบ
    maxPerDayTeacher: false,    // true = ครูสอนไม่เกิน 1 คาบ/วัน
    noConsecTeacher:  false,    // true = ห้ามครูสอนติดกัน 2 คาบขึ้นไปเลย (= maxConsec 1)
    penalizeLunchGap: false,    // true = soft penalty: หลีกเลี่ยงครูว่างช่วงคาบ 4+5 > 2 วัน
    runs:        10,            // จำนวนรอบ (10 default)
  });
  const [autoProgress, setAutoProgress] = useState(null); // {run, total}

  const teacher  = S.teachers.find(t=>t.id===selT);
  // asgns: รวม assignment ที่ครูเป็นหลัก + assignment ที่ครูถูก assign เป็น coTeacher ใน cardCoMap
  const asgns    = S.assigns.filter(a=>a.teacherId===selT);
  // coAsgns: assignment ที่มี selT เป็น co-teacher (ผ่าน cardCoMap หรือ schedule entry)
  const coAsgnsIds = new Set(
    Object.entries(S.schedule).flatMap(([,en])=>
      (en||[]).filter(e=>{
        const coIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
        return coIds.includes(selT) && e.teacherId!==selT;
      }).map(e=>e.assignmentId)
    ).filter(Boolean)
  );
  const coAsgns  = S.assigns.filter(a=>coAsgnsIds.has(a.id));
  const allAsgns = [...asgns, ...coAsgns.filter(a=>!asgns.find(x=>x.id===a.id))];
  const fTeachers= selDept ? S.teachers.filter(t=>t.departmentId===selDept) : S.teachers;

  // sort helper: inline ใน useMemo เพื่อกัน stale closure
  const sortedRooms = useMemo(()=>{
    const key=(r)=>{
      const lvName=S.levels.find(l=>l.id===r.levelId)?.name||"";
      const lvNum=parseInt((lvName.match(/(\d+)/)||[0,999])[1]);
      const rmNum=parseInt((r.name.match(/(\d+)$/) || r.name.match(/(\d+)/) ||[0,0])[1]||0);
      return lvNum*10000+rmNum;
    };
    return [...S.rooms].sort((a,b)=>key(a)-key(b));
  },[S.rooms,S.levels]);
  // tRooms: ห้องของครูที่เลือก (รวมห้องที่เป็น co-teacher) เรียงตาม sortedRooms
  const tRoomsSet = new Set(allAsgns.flatMap(a=>a.roomIds));
  const tRooms = sortedRooms.filter(r=>tRoomsSet.has(r.id)).map(r=>r.id);

  /* ── helpers ── */
  const blocked=useCallback(tid=>{
    const t=S.teachers.find(x=>x.id===tid);
    if(!t)return[];
    const b=[];
    (t.specialRoles||[]).forEach(rid=>{
      const r=SROLES.find(x=>x.id===rid);
      r?.blocked?.forEach(bl=>bl.periods.forEach(p=>b.push({day:bl.day,period:p,reason:r.name})));
    });
    // คาบล็อคแผนก (custom) — ล็อคทุกคนในโรงเรียน
    S.meetings.filter(m=>m.type==="custom")
      .forEach(m=>(m.slots||[]).forEach(sl=>b.push({day:sl.day,period:sl.period,reason:m.name||"ล็อคแผนก"})));
    // คาบล็อคกลุ่มสาระ (เดิม) — ล็อคเฉพาะกลุ่มสาระ
    S.meetings.filter(m=>(!m.type||m.type==="dept")&&m.departmentId===t.departmentId)
      .forEach(m=>m.periods.forEach(p=>b.push({day:m.day,period:p,reason:"ประชุม"})));
    // คาบล็อกส่วนตัว
    (t.personalLocks||[]).forEach(pl=>
      (pl.periods||[]).forEach(p=>b.push({day:pl.day,period:p,reason:pl.reason||"ส่วนตัว"}))
    );
    return b;
  },[S.teachers,S.meetings]);

  const isBlk=(tid,day,p)=>blocked(tid).some(b=>b.day===day&&b.period===p);
  const sk=(rid,day,p)=>rid+"_"+day+"_"+p;

  const teacherBusy=(tid,day,period,excludeKey,newSubjectId=null)=>{
    for(const [k,en] of Object.entries(S.schedule)){
      if(k===excludeKey)continue;
      if(!k.endsWith("_"+day+"_"+period))continue;
      if(en?.some(e=>{
        const eCoIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
        if(e.teacherId!==tid&&!eCoIds.includes(tid))return false;
        // NP/-2 mode: ถ้าวิชาเดียวกัน → อนุญาตลงคนละห้องคาบเดียวกัน
        if(newSubjectId&&e.subjectId===newSubjectId){
          const sub=S.subjects.find(s=>s.id===e.subjectId);
          const ca=sub?.consecutiveAllowed||0;
          if(ca===-1||ca===-2)return false;
        }
        return true;
      }))return true;
    }
    return false;
  };

  const specialRoomBusy=(subjectId,day,period,excludeKey)=>{
    const srId=S.subjects.find(s=>s.id===subjectId)?.specialRoomId;
    if(!srId)return false;
    for(const [k,en] of Object.entries(S.schedule)){
      if(k===excludeKey)continue;
      if(!k.endsWith("_"+day+"_"+period))continue;
      if(en?.some(e=>S.subjects.find(s=>s.id===e.subjectId)?.specialRoomId===srId))return true;
    }
    return false;
  };

  const sameSubjectSameDay=(subjectId,roomId,day,excludeKey)=>{
    const allowed=S.subjects.find(s=>s.id===subjectId)?.consecutiveAllowed||0;
    // NP (-1): อนุญาตลงวันเดิมได้ไม่จำกัดคาบ (สอนหลายห้องพร้อมกัน นับครูแค่ 1 คาบ)
    if(allowed===-1) return false;
    // เศรษฐ-วิศวะ (-2): อนุญาต 2 คาบต่อห้องต่อวัน (2 คาบติด) แต่ห้ามเกิน 2
    if(allowed===-2){
      let c=0;
      for(const [k,en] of Object.entries(S.schedule)){
        if(k===excludeKey)continue;
        const pts=k.split("_");
        if(pts[0]!==roomId||pts[1]!==day)continue;
        en?.forEach(e=>{if(e.subjectId===subjectId)c++;});
      }
      return c>=2;
    }
    if(allowed>0)return false;
    let count=0;
    for(const [k,en] of Object.entries(S.schedule)){
      if(k===excludeKey)continue;
      const pts=k.split("_");
      if(pts[0]!==roomId||pts[1]!==day)continue;
      en?.forEach(e=>{if(e.subjectId===subjectId)count++;});
    }
    return count>=1;
  };

  const countSubjectInRoom=(assignId,roomId)=>{
    let c=0;
    Object.entries(S.schedule).forEach(([k,en])=>{
      if(!k.startsWith(roomId+"_"))return;
      en?.forEach(e=>{if(e.assignmentId===assignId)c++;});
    });
    return c;
  };

  const getPerRoomLimit=(assignId)=>{
    const a=S.assigns.find(x=>x.id===assignId);
    if(!a)return 999;
    return S.subjects.find(s=>s.id===a.subjectId)?.periodsPerWeek||999;
  };

  const aUsed=(aid)=>{
    const a=S.assigns.find(x=>x.id===aid);
    const sub=a?S.subjects.find(s=>s.id===a.subjectId):null;
    const ca=sub?.consecutiveAllowed||0;
    if(ca===-2){
      // -2 mode: นับ entries ทั้งหมด (ทุกห้อง ทุกคาบ) ของ subjectId นี้
      // เพื่อเทียบกับ periodsPerWeek × จำนวนห้อง
      const allAids=new Set(S.assigns.filter(x=>x.subjectId===a.subjectId).map(x=>x.id));
      let c=0;
      Object.values(S.schedule).forEach(en=>en?.forEach(e=>{
        if(allAids.has(e.assignmentId)) c++;
      }));
      return c; // จำนวน entries ทั้งหมด
    }
    let c=0;
    Object.values(S.schedule).forEach(en=>en?.forEach(e=>{if(e.assignmentId===aid)c++;}));
    return c;
  };

  const teacherScheduledTotal=(tid)=>{
    // NP mode: วิชาเดียวกัน วันเดียวกัน คาบเดียวกัน → นับแค่ 1 คาบ (ไม่ว่าจะลงกี่ห้อง)
    const seen=new Set();
    let c=0;
    Object.entries(S.schedule).forEach(([k,en])=>{
      const pts=k.split("_"); // [roomId, day, period]
      en?.forEach(e=>{
        const eCIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
        if(e.teacherId===tid||eCIds.includes(tid)){
          const sub=S.subjects.find(s=>s.id===e.subjectId);
          const ca=sub?.consecutiveAllowed||0;
          if(ca===-1||ca===-2){
            // NP/-2: deduplicate ด้วย subjectId_day_period (ไม่นับซ้ำข้ามห้อง)
            const npKey=e.subjectId+"_"+pts[pts.length-2]+"_"+pts[pts.length-1];
            if(!seen.has(npKey)){seen.add(npKey);c++;}
          } else {
            c++;
          }
        }
      });
    });
    return c;
  };


  /* ── Auto Schedule (multi-run) ── */
  const runAutoSchedule = () => setShowAutoModal(true);

  const executeAutoSchedule = (opts) => {
    autoBaseRef.current=JSON.stringify({schedule:S.schedule,locks:S.locks,assigns:S.assigns,teachers:S.teachers,subjects:S.subjects,rooms:S.rooms,meetings:S.meetings});
    setProposal(null);
    setShowAutoModal(false);
    setAutoRunning(true);
    setAutoResult(null);
    setAutoProgress({ run: 0, total: opts.runs });

    // รันแบบ async loop เพื่อให้ UI อัพเดท progress ได้
    let bestResult = null;

    const runOnce = (runIdx) => {
      setTimeout(() => {
        setAutoProgress({ run: runIdx + 1, total: opts.runs });

        // ── เริ่มต้น schedule ──
        // ถ้า full mode → เก็บเฉพาะคาบที่ล็อคไว้
        const newSchedule = {};
        if (opts.mode === "full") {
          Object.entries(S.schedule).forEach(([k, en]) => {
            if (S.locks[k]) newSchedule[k] = en; // เก็บคาบที่ล็อค
          });
        } else {
          Object.assign(newSchedule, S.schedule);
        }

        let placed = 0, skipped = 0;
        const skippedList = [];

        // ── helper functions ──
        const sk2 = (rid, day, p) => rid + "_" + day + "_" + p;

        const isBusy2 = (tid, day, p, excKey, subId = null) => {
          for (const [k, en] of Object.entries(newSchedule)) {
            if (k === excKey) continue;
            const pts = k.split("_");
            const kDay = pts[pts.length - 2];
            const kPer = parseInt(pts[pts.length - 1]);
            if (kDay !== day || kPer !== p) continue;
            for (const e of (en || [])) {
              const coIds = e.coTeacherIds?.length ? e.coTeacherIds : (e.coTeacherId ? [e.coTeacherId] : []);
              if (e.teacherId !== tid && !coIds.includes(tid)) continue;
              if (subId) {
                const sub = S.subjects.find(s => s.id === e.subjectId);
                const ca = sub?.consecutiveAllowed || 0;
                if ((ca === -1 || ca === -2) && e.subjectId === subId) return false;
              }
              return true;
            }
          }
          return false;
        };

        const isLocked2 = (key) => !!S.locks[key];
        const isBlk2 = (tid, day, p) => blocked(tid).some(b => b.day === day && b.period === p);

        const srBusy2 = (subId, day, p) => {
          const sub = S.subjects.find(s => s.id === subId);
          if (!sub?.specialRoomId) return false;
          for (const [k, en] of Object.entries(newSchedule)) {
            const pts = k.split("_");
            if (pts[pts.length - 2] !== day || parseInt(pts[pts.length - 1]) !== p) continue;
            if ((en || []).some(e => {
              const s2 = S.subjects.find(x => x.id === e.subjectId);
              return s2?.specialRoomId === sub.specialRoomId;
            })) return true;
          }
          return false;
        };

        const countInRoom2 = (aId, rId) => {
          let c = 0;
          for (const [k, en] of Object.entries(newSchedule)) {
            if (!k.startsWith(rId + "_")) continue;
            (en || []).forEach(e => { if (e.assignmentId === aId) c++; });
          }
          return c;
        };

        const sameSubDay2 = (subId, rId, day) => {
          const sub = S.subjects.find(s => s.id === subId);
          const ca = sub?.consecutiveAllowed || 0;
          if (ca === -1) return false; // NP: ลงวันเดิมได้ไม่จำกัด
          if (ca >= 2) return false;
          let c = 0;
          for (const [k, en] of Object.entries(newSchedule)) {
            const pts = k.split("_");
            if (pts.slice(0, -2).join("_") !== rId || pts[pts.length - 2] !== day) continue;
            (en || []).forEach(e => { if (e.subjectId === subId) c++; });
          }
          return c >= (ca === 0 ? 1 : ca);
        };

        // ── เงื่อนไขเพิ่มเติม ──

        // noFirstLast: ถ้าวิชานี้มีคาบ 1 อยู่แล้ว ห้ามวางคาบ 7 ในวันเดิม (และกลับกัน)
        const violatesFirstLast = (subId, rId, day, period) => {
          if (!opts.noFirstLast) return false;
          if (period !== 1 && period !== 7) return false;
          const counterPeriod = period === 1 ? 7 : 1;
          const counterKey = sk2(rId, day, counterPeriod);
          return (newSchedule[counterKey] || []).some(e => e.subjectId === subId);
        };

        // maxConsecTeacher: ครูสอนติดกันไม่เกิน N คาบ
        const teacherConsecCount = (tid, day, period) => {
          if (!opts.maxConsecTeacher) return false;
          let streak = 0;
          for (let p = period - 1; p >= 1; p--) {
            let found = false;
            Object.entries(newSchedule).forEach(([k, en]) => {
              const pts = k.split("_");
              if (pts[pts.length - 2] !== day || parseInt(pts[pts.length - 1]) !== p) return;
              if ((en || []).some(e => {
                const coIds = e.coTeacherIds?.length ? e.coTeacherIds : (e.coTeacherId ? [e.coTeacherId] : []);
                return e.teacherId === tid || coIds.includes(tid);
              })) found = true;
            });
            if (found) streak++;
            else break;
          }
          return streak >= opts.maxConsecTeacher;
        };

        // maxPerDayTeacher: ครูสอนไม่เกิน 1 คาบ/วัน
        const teacherAlreadyTaughtToday = (tid, day) => {
          if (!opts.maxPerDayTeacher) return false;
          for (const [k, en] of Object.entries(newSchedule)) {
            const pts = k.split("_");
            if (pts[pts.length - 2] !== day) continue;
            if ((en || []).some(e => {
              const coIds = e.coTeacherIds?.length ? e.coTeacherIds : (e.coTeacherId ? [e.coTeacherId] : []);
              return e.teacherId === tid || coIds.includes(tid);
            })) return true;
          }
          return false;
        };

        // ── สร้าง jobs ──
        const jobs = [];
        S.assigns.forEach(a => {
          const sub = S.subjects.find(s => s.id === a.subjectId);
          const ca = sub?.consecutiveAllowed || 0;

          // กรองตาม opts
          if (ca === -2) return; // เศรษฐ-วิศวะ ข้ามเสมอ (complex)
          if (ca === -1 && !opts.allowNP) return;
          if (ca >= 2 && !opts.allowConsec) return;
          if (sub?.specialRoomId && !opts.allowSR) return;
          if (ca === 0 && !sub?.specialRoomId && !opts.allowNormal) return;

          a.roomIds.forEach(rid => {
            const limit = sub?.periodsPerWeek || a.totalPeriods || 1;
            const placed2 = countInRoom2(a.id, rid);
            const remaining = limit - placed2;
            if (remaining <= 0) return;
            const coTids = cardCoMap[a.id] || [];
            const busyScore = teacherScheduledTotal(a.teacherId);
            const score = busyScore * 10 + (ca > 0 ? ca * 5 : 0) + (sub?.specialRoomId ? 8 : 0);
            for (let i = 0; i < remaining; i++) jobs.push({ a, rid, sub, ca, coTids, score });
          });
        });

        // เรียงจากยากไปง่าย
        jobs.sort((x, y) => y.score - x.score);

        const shuffled = (arr) => [...arr].sort(() => Math.random() - 0.5);

        // วาง jobs
        jobs.forEach(({ a, rid, sub, ca, coTids }) => {
          const subId = a.subjectId;
          const tid = a.teacherId;
          let foundSlot = false;

          const days = shuffled(DAYS);
          outer: for (const day of days) {
            const periods = shuffled(PERIODS);
            for (const p of periods) {
              const key = sk2(rid, day, p.id);

              if (isLocked2(key)) continue;
              if ((newSchedule[key] || []).length >= 3) continue;
              if (isBlk2(tid, day, p.id)) continue;
              if (isBusy2(tid, day, p.id, null, subId)) continue;
              if (srBusy2(subId, day, p.id)) continue;
              if (sameSubDay2(subId, rid, day)) continue;

              // เงื่อนไขเพิ่มเติม
              if (violatesFirstLast(subId, rid, day, p.id)) continue;
              if (teacherConsecCount(tid, day, p.id)) continue;
              if (teacherAlreadyTaughtToday(tid, day)) continue;
              // noConsecTeacher: ห้ามติดกันเลย — ตรวจคาบก่อนหน้าและถัดไป
              if (opts.noConsecTeacher) {
                const prevBusy = Object.entries(newSchedule).some(([k,en])=>{
                  const pts=k.split("_"); if(pts[pts.length-2]!==day||parseInt(pts[pts.length-1])!==p.id-1)return false;
                  return (en||[]).some(e=>{const c=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);return e.teacherId===tid||c.includes(tid);});
                });
                const nextBusy = Object.entries(newSchedule).some(([k,en])=>{
                  const pts=k.split("_"); if(pts[pts.length-2]!==day||parseInt(pts[pts.length-1])!==p.id+1)return false;
                  return (en||[]).some(e=>{const c=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);return e.teacherId===tid||c.includes(tid);});
                });
                if (prevBusy || nextBusy) continue;
              }

              // consecutive ≥ 2
              if (ca >= 2) {
                const hasPrev = (newSchedule[sk2(rid, day, p.id - 1)] || []).some(e => e.subjectId === subId);
                const hasNext = (newSchedule[sk2(rid, day, p.id + 1)] || []).some(e => e.subjectId === subId);
                const countSameDay = (() => {
                  let c = 0;
                  PERIODS.forEach(pp => {
                    (newSchedule[sk2(rid, day, pp.id)] || []).forEach(e => { if (e.subjectId === subId) c++; });
                  });
                  return c;
                })();
                if (!hasPrev && !hasNext && countSameDay === 0) {
                  const nextKey = sk2(rid, day, p.id + 1);
                  const nextFree = !isLocked2(nextKey)
                    && (newSchedule[nextKey] || []).length < 3
                    && !isBusy2(tid, day, p.id + 1, null, subId)
                    && !isBlk2(tid, day, p.id + 1);
                  if (!nextFree) continue;
                }
              }

              const entry = {
                id: gid(),
                teacherId: tid,
                subjectId: subId,
                assignmentId: a.id,
                coTeacherIds: coTids,
                coTeacherId: coTids[0] || null,
              };
              newSchedule[key] = [...(newSchedule[key] || []), entry];
              placed++;
              foundSlot = true;
              break outer;
            }
          }
          if (!foundSlot) {
            skipped++;
            skippedList.push(`${sub?.code || ""} ${subDisplayName(sub) || ""} — ${S.rooms.find(r => r.id === rid)?.name || ""}`);
          }
        });

        // penalizeLunchGap: นับครูที่ว่างคาบ 4+5 พร้อมกันมากกว่า 2 วัน (soft penalty)
        let lunchPenalty = 0;
        if (opts.penalizeLunchGap) {
          S.teachers.forEach(t => {
            let freeCount = 0;
            DAYS.forEach(day => {
              const free4 = !Object.entries(newSchedule).some(([k,en]) => {
                const pts=k.split("_"); if(pts[pts.length-2]!==day||parseInt(pts[pts.length-1])!==4)return false;
                return (en||[]).some(e=>{const c=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);return e.teacherId===t.id||c.includes(t.id);});
              });
              const free5 = !Object.entries(newSchedule).some(([k,en]) => {
                const pts=k.split("_"); if(pts[pts.length-2]!==day||parseInt(pts[pts.length-1])!==5)return false;
                return (en||[]).some(e=>{const c=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);return e.teacherId===t.id||c.includes(t.id);});
              });
              if (free4 && free5) freeCount++;
            });
            if (freeCount > 2) lunchPenalty += (freeCount - 2);
          });
        }

        const resultScore = placed * 100 - skipped * 10 - lunchPenalty;
        const result = { placed, skipped, details: skippedList, schedule: newSchedule, score: resultScore };

        if (!bestResult || resultScore > bestResult.score) {
          bestResult = result;
        }

        if (runIdx + 1 < opts.runs) {
          runOnce(runIdx + 1);
        } else {
          // จบครบ opts.runs รอบ — ใช้ bestResult
          setProposal({schedule:bestResult.schedule,base:autoBaseRef.current});
          setAutoResult({
            placed: bestResult.placed,
            skipped: bestResult.skipped,
            details: bestResult.details,
            runs: opts.runs,
          });
          setAutoRunning(false);
          setAutoProgress(null);
          st(`เตรียมผลเสนอแล้ว: ลงได้ ${bestResult.placed} คาบ กรุณาตรวจสอบก่อนนำไปใช้`, "success");
        }
      }, 80); // delay เล็กน้อยให้ UI re-render ได้
    };

    runOnce(0);
  };

  /* ── drop handler ── */
  const handleDrop=(rid,day,p)=>{
    const drag=dragRef.current;  // อ่านจาก ref กัน stale state
    const key=sk(rid,day,p);
    if(S.locks[key]){st("ล็อคแล้ว","error");return;}
    if((S.schedule[key]||[]).length>=3){st("ครบ 3 วิชาแล้ว","error");return;}

    // กรณี re-drag การ์ดที่วางอยู่แล้ว → ย้ายช่อง (ทำได้ทั้ง 2 mode)
    if(drag?.fromKey){
      if(drag.fromKey===key)return;
      // ข้อ 3: ห้ามลากข้ามห้อง — เปรียบเทียบ roomId โดยตรงจาก entry กับ target room
      // ตรวจ cross-room โดยใช้ fromRoomId ที่ฝังไว้ตั้งแต่ onDragStart
      if(drag.fromRoomId!==rid){st("ห้ามลากข้ามห้องเรียน!","error");setDragBoth(null);return;}
      const entry=drag.entry;
      const sub=S.subjects.find(s=>s.id===entry.subjectId);
      // ไม่ตรวจ sub.levelId เพราะวิชาอาจสอนหลายระดับ (NP/multi-room) — fromRoomId ตรวจแล้ว
      if(specialRoomBusy(entry.subjectId,day,p,drag.fromKey)){
        const sr=S.specialRooms.find(r=>r.id===sub?.specialRoomId);
        st("ห้องพิเศษ '"+(sr?.name||"")+"' ถูกใช้อยู่","error");return;
      }
      // ตรวจ teacher conflict เฉพาะ teacher-mode
      if(selT){
        if(isBlk(entry.teacherId,day,p)){st("ครูถูกล็อคคาบนี้","error");return;}
        if(teacherBusy(entry.teacherId,day,p,drag.fromKey,entry.subjectId)){st("ครูคนนี้สอนคาบนี้อยู่แล้ว","error");return;}
      }
      U.setSchedule(prev=>{
        const u={...prev};
        u[drag.fromKey]=(u[drag.fromKey]||[]).filter(e=>e.id!==entry.id);
        u[key]=[...(u[key]||[]),entry];
        return u;
      });
      setDragBoth(null);return;
    }

    // กรณีลากจาก sidebar (teacher-mode เท่านั้น)
    if(!drag?.teacherId)return;
    const sub=S.subjects.find(s=>s.id===drag.subjectId);
    const targetRoom=S.rooms.find(r=>r.id===rid);
    // ห้ามวางในห้องที่ไม่ได้อยู่ใน assignment
    const asgn=S.assigns.find(a=>a.id===drag.assignmentId);
    // mode -2: อนุญาตถ้า rid อยู่ใน assignment ใดก็ได้ที่มี subjectId เดียวกัน (2 ห้องพร้อมกัน)
    const subCa=S.subjects.find(s=>s.id===drag.subjectId)?.consecutiveAllowed||0;
    const roomAllowed = asgn?.roomIds?.includes(rid) ||
      (subCa===-2 && S.assigns.some(a=>a.subjectId===drag.subjectId&&a.roomIds?.includes(rid)));
    if(!roomAllowed){st("ห้องนี้ไม่ได้รับมอบหมายวิชานี้!","error");setDragBoth(null);return;}
    if(isBlk(drag.teacherId,day,p)){st("ครูถูกล็อคคาบนี้","error");return;}
    if(teacherBusy(drag.teacherId,day,p,null,drag.subjectId)){st("ครูคนนี้สอนคาบนี้อยู่แล้ว (ห้องอื่น)","error");return;}
    if(specialRoomBusy(drag.subjectId,day,p,null)){
      const sr=S.specialRooms.find(r=>r.id===sub?.specialRoomId);
      st("ห้องพิเศษ '"+(sr?.name||"")+"' ถูกใช้อยู่แล้วในคาบนี้","error");return;
    }
    if(targetRoom&&sub&&targetRoom.levelId!==sub.levelId){
      // อนุญาตถ้าห้องนี้อยู่ใน assignment roomIds แล้ว (วิชาสอนหลายระดับ เช่น NP)
      const assignHasRoom=S.assigns.some(a=>a.subjectId===drag.subjectId&&a.roomIds?.includes(rid));
      if(!assignHasRoom){st("ระดับชั้นไม่ตรงกัน!","error");return;}
    }
    if(sameSubjectSameDay(drag.subjectId,rid,day,null)){st("วิชานี้มีในวัน"+day+"แล้ว (ห้ามซ้ำ/วัน)","error");return;}
    // สำหรับ -2 mode: หา assignment ที่ตรงกับห้องปลายทาง (อาจต่างจาก drag.assignmentId)
    // -2 mode: หา assignment ที่ตรงกับ rid และ teacherId เดียวกัน ถ้าไม่มีค่อยหา assignment อื่นของ subjectId
    const effectiveAsgn=subCa===-2
      ? (S.assigns.find(a=>a.teacherId===drag.teacherId&&a.subjectId===drag.subjectId&&a.roomIds?.includes(rid))
         || S.assigns.find(a=>a.subjectId===drag.subjectId&&a.roomIds?.includes(rid))
         || asgn)
      : asgn;
    const effectiveAid=effectiveAsgn?.id||drag.assignmentId;
    const placed=countSubjectInRoom(effectiveAid,rid);
    const limit=getPerRoomLimit(effectiveAid);
    if(placed>=limit){st("ห้องนี้ลงครบ "+limit+" คาบแล้ว","error");return;}
    const coTids=cardCoMap[drag.assignmentId]||cardCoMap[effectiveAid]||[];
    const mainEntry={id:gid(),teacherId:drag.teacherId,subjectId:drag.subjectId,assignmentId:effectiveAid,coTeacherIds:coTids,coTeacherId:coTids[0]||null};
    const bundles=bundleMap[drag.assignmentId]||[];
    const bundleEntries=bundles.map(b=>{
      const ba=S.assigns.find(a=>a.id===b.assignId);if(!ba)return null;
      const bCoTids=cardCoMap[b.assignId]||[];
      return{id:gid(),teacherId:b.teacherId||ba.teacherId,subjectId:ba.subjectId,assignmentId:b.assignId,coTeacherIds:bCoTids,coTeacherId:bCoTids[0]||null};
    }).filter(Boolean);
    U.setSchedule(prev=>({...prev,[key]:[...(prev[key]||[]),mainEntry,...bundleEntries]}));
    setDragBoth(null);
  };

  /* ── co-teacher dept+teacher selector ── */
  const CoTeacherSelect=({coSVal,setCoSFn,coDeptVal,setCoDeptFn,excludeId})=>(
    <div style={{display:"flex",flexDirection:"column",gap:8}}>
      <SearchSelect value={coDeptVal} onChange={v=>{setCoDeptFn(v);setCoSFn("");}} options={[{value:"",label:"-- เลือกกลุ่มสาระก่อน --"},...S.depts.map(d=>({value:d.id,label:d.name}))]} placeholder="-- เลือกกลุ่มสาระก่อน --"/>
      {coDeptVal&&(
        <SearchSelect value={coSVal} onChange={v=>setCoSFn(v)} options={[{value:"",label:"-- เลือกครู --"},...S.teachers.filter(t=>t.departmentId===coDeptVal&&t.id!==excludeId).map(t=>{const rem=(t.totalPeriods||0)-teacherScheduledTotal(t.id);return{value:t.id,label:`${t.prefix}${t.firstName} ${t.lastName} — เหลือ ${rem} คาบ`}})]} placeholder="-- เลือกครู --"/>
      )}
    </div>
  );

  /* ── render timetable table ── */
  const LEVEL_COLORS=[
    {bg:"#FFF7ED",border:"#FED7AA",head:"#EA580C"},
    {bg:"#F0FDF4",border:"#BBF7D0",head:"#16A34A"},
    {bg:"#EFF6FF",border:"#BFDBFE",head:"#2563EB"},
    {bg:"#FDF4FF",border:"#E9D5FF",head:"#9333EA"},
    {bg:"#FFF1F2",border:"#FECDD3",head:"#E11D48"},
    {bg:"#F0FDFA",border:"#99F6E4",head:"#0D9488"},
  ];
  const renderTable=(roomIds)=>(
    <div style={{flex:1,overflowX:"auto"}}>
      {roomIds.map(rid=>{
        const rm=S.rooms.find(r=>r.id===rid);
        const rmPlan=S.plans.find(p=>p.id===rm?.planId);
        const rmLevel=S.levels.find(l=>l.id===rm?.levelId);
        const lvIdx=S.levels.findIndex(l=>l.id===rm?.levelId);
        const lc=LEVEL_COLORS[lvIdx>=0?lvIdx%LEVEL_COLORS.length:0];
        return (
          <section key={rid} className="room-board">
            <div className="room-board-heading">
              <span className="room-name"><Glyph name="grid" size={18}/>{rm?.name}</span>
              {rmPlan&&<span className="room-plan">{rmPlan.name}</span>}<ColorBadge item={S.levels.find(l=>l.id===rm?.levelId)} kind="level"/>
              {rmLevel&&<span style={{color:"#9CA3AF",fontSize:11}}>{rmLevel.name}</span>}
            </div>
            <div className="room-table-scroll">
              <table data-ui-table="true" className="timetable-grid">
                <thead>
                  <tr style={{borderBottom:`2px solid ${lc.head}`}}>
                    <th style={{padding:"10px 10px",background:lc.head,color:"#fff",width:62,textAlign:"left",fontSize:13,fontWeight:700,letterSpacing:"0.02em"}}>วัน</th>
              {/* header: ใช้เวลาคาบตาม division ของห้อง */}
              {(()=>{const rm=S.rooms.find(r=>r.id===rid);const divId=getDivisionForRoom(rm,S);const pList=getPeriodCfg(divId).periods;return pList.map(p=>(
                <th key={p.id} style={{padding:"6px 2px",background:lc.head,textAlign:"center",borderLeft:"1px solid rgba(255,255,255,0.2)"}}>
                  <div style={{fontSize:11,color:"#fff",fontWeight:700}}>คาบ {p.id}</div>
                  <div style={{fontSize:9,color:"rgba(255,255,255,0.7)",fontWeight:400}}>{p.time}</div>
                </th>
              ));})()}
                  </tr>
                </thead>
                <tbody>
                  {DAYS.map((day,di)=>{
                    const rowBg=di%2===0?"#FFFFFF":lc.border+"44";
                    return(
                    <tr key={day} style={{background:rowBg}}>
                      <td style={{padding:"8px 8px",fontWeight:700,fontSize:12,color:lc.head,borderRight:`2px solid ${lc.border}`,borderBottom:`1px solid ${lc.border}`,background:lc.bg}}>{day}</td>
                      {PERIODS.map(p=>{
                        const key=sk(rid,day,p.id);
                        const en=S.schedule[key]||[];
                        const lk=!!S.locks[key];
                        const bl=mode==="teacher"&&!!selT&&isBlk(selT,day,p.id);
                        // คาบล็อคแผนก (custom) — แสดงทุกตาราง
                        const customLock=(S.meetings||[]).find(m=>m.type==="custom"&&(m.slots||[]).some(s=>s.day===day&&s.period===p.id));
                        return (
                          <td key={p.id}
                            className={"dz "+(customLock||bl?"cell-blocked ":"")+(lk?"cell-locked":"")}
                            onDragOver={e=>{const d=dragRef.current;if(!d){e.currentTarget.classList.remove("over");return;}
if(d.fromRoomId&&d.fromRoomId!==rid){e.currentTarget.classList.remove("over");return;}
if(d.assignmentId){const a=S.assigns.find(x=>x.id===d.assignmentId);const sCa=S.subjects.find(s=>s.id===d.subjectId)?.consecutiveAllowed||0;const ok=a?.roomIds?.includes(rid)||(sCa===-2&&S.assigns.some(x=>x.subjectId===d.subjectId&&x.roomIds?.includes(rid)));if(!ok){e.currentTarget.classList.remove("over");return;}}
e.preventDefault();e.currentTarget.classList.add("over");}}
                            onDragLeave={e=>e.currentTarget.classList.remove("over")}
                            onDrop={e=>{e.preventDefault();e.currentTarget.classList.remove("over");handleDrop(rid,day,p.id);}}
                            
                          >
                            {!en.length&&!bl&&!customLock&&<span className="empty-slot" aria-hidden="true">·</span>}
                            {customLock&&(
                              <div style={{fontSize:9,color:"#E65100",textAlign:"center",padding:"2px 2px 0",fontWeight:700,lineHeight:1.2}}>
                                🏫 {customLock.name}
                              </div>
                            )}
                            {bl&&en.length===0&&(
                              <div style={{fontSize:9,color:"#92400E",textAlign:"center",padding:4}}>
                                🔒 {blocked(selT).find(b=>b.day===day&&b.period===p.id)?.reason}
                              </div>
                            )}
                            {en.map(entry=>(
                              <SchedulerEntryCard
                                key={entry.id}
                                entry={entry}
                                cellKey={key}
                                lk={lk}
                                cellCount={en.length}
                                selT={selT}
                                mode={mode}
                                S={S}
                                U={U}
                                gc={gc}
                                setDrag={setDragBoth}
                                setCoM={setCoM}
                              />
                            ))}
                          </td>
                        );
                      })}
                    </tr>
                  );})}
                </tbody>
              </table>
            </div>
          </section>
        );
      })}
    </div>
  );

  const renderTeacherWeeklySummary=()=> !selT||mode!=="teacher"?null:<TeacherMini S={S} teacherId={selT} periods={getPeriodCfg(getDivisionForTeacher(selT,S)).periods} open={showWeekly} onToggle={()=>setShowWeekly(v=>!v)} isBlocked={isBlk}/>;

  /* ── render ── */
  return (
    <div style={{animation:"fadeIn 0.3s"}}>

      {/* Mode + selector bar */}
      <div style={{display:"flex",gap:8,marginBottom:14,alignItems:"center",flexWrap:"wrap"}}>
        <div style={{display:"flex",borderRadius:10,overflow:"hidden",border:"1.5px solid "+CRED,boxShadow:"0 2px 8px rgba(185,28,28,0.15)"}}>
          <button data-ui-control="true" onClick={()=>{setMode("teacher");setSelRoom("");}} style={{padding:"8px 20px",background:mode==="teacher"?CRED:"#fff",color:mode==="teacher"?"#fff":CRED,border:"none",fontWeight:700,fontSize:13,cursor:"pointer",transition:"background 0.15s"}}>จัดรายครู</button>
          <button data-ui-control="true" onClick={()=>{setMode("room");setSelT("");setSelDept("");}} style={{padding:"8px 20px",background:mode==="room"?CRED:"#fff",color:mode==="room"?"#fff":CRED,border:"none",fontWeight:700,fontSize:13,cursor:"pointer",transition:"background 0.15s"}}>จัดรายห้อง</button>
        </div>

        {mode==="teacher"&&<>
          <SearchSelect value={selDept} onChange={v=>{setSelDept(v);setSelT("");}} options={[{value:"",label:"-- ทุกกลุ่มสาระ --"},...S.depts.map(d=>({value:d.id,label:d.name}))]} placeholder="-- ทุกกลุ่มสาระ --" style={{maxWidth:200}}/>
          <select data-ui-control="true" style={{...IS,maxWidth:280}} value={selT} onChange={e=>setSelT(e.target.value)}>
            <option value="">-- เลือกครู --</option>
            {fTeachers.map(t=>{
              const rem=(t.totalPeriods||0)-teacherScheduledTotal(t.id);
              return <option key={t.id} value={t.id}>{t.prefix}{t.firstName} {t.lastName} (เหลือ {rem})</option>;
            })}
          </select>
        </>}

        {mode==="room"&&(
          <select data-ui-control="true" style={{...IS,maxWidth:300}} value={selRoom} onChange={e=>setSelRoom(e.target.value)}>
            <option value="">-- เลือกห้องเรียน --</option>
            {sortedRooms.map(r=>{
              const lv=S.levels.find(l=>l.id===r.levelId);
              return <option key={r.id} value={r.id}>{lv?.name} — {r.name}</option>;
            })}
          </select>
        )}
        {/* Auto Schedule + ล้างคาบกำพร้า */}
        <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center",flexShrink:0}}>
          <button data-ui-control="true" onClick={async()=>{
            const validAssignIds=new Set(S.assigns.map(a=>a.id));
            const validSubjectIds=new Set(S.subjects.map(s=>s.id));
            const validTeacherIds=new Set(S.teachers.map(t=>t.id));
            const validTeacherSubs=new Map();
            S.assigns.forEach(a=>{
              if(!validTeacherSubs.has(a.teacherId)) validTeacherSubs.set(a.teacherId,[]);
              validTeacherSubs.get(a.teacherId).push(a.subjectId);
            });

            let removed=0;
            const next={};
            Object.entries(S.schedule).forEach(([k,en])=>{
              const filtered=(en||[]).filter(e=>{
                // มี assignmentId → ตรวจว่ายังมี assign อยู่ไหม
                if(e.assignmentId) return validAssignIds.has(e.assignmentId);
                // subjectId ถูกลบ → กำพร้า
                if(e.subjectId&&!validSubjectIds.has(e.subjectId)) return false;
                // teacherId ถูกลบ → กำพร้า
                if(e.teacherId&&!validTeacherIds.has(e.teacherId)) return false;
                // ไม่มี assignmentId → ตรวจ teacher+subject combo
                if(e.teacherId&&e.subjectId){
                  return (validTeacherSubs.get(e.teacherId)||[]).includes(e.subjectId);
                }
                return false;
              });
              removed+=(en||[]).length-filtered.length;
              if(filtered.length) next[k]=filtered;
            });

            if(removed===0){st("ไม่มีคาบกำพร้า ✓");return;}
            if(!await uiConfirm(`พบ ${removed} คาบกำพร้า\nลบออกทั้งหมดไหม?`))return;

            // Use the same guarded autosave as ordinary timetable edits.
            U.setSchedule(next);
            st(`ลบ ${removed} คาบกำพร้าแล้ว กำลังบันทึกตามปกติ`,"warning");
          }} style={{...BO("#DC2626"),fontSize:12,padding:"7px 12px",whiteSpace:"nowrap",flexShrink:0}}>
            🧹 ล้างคาบกำพร้า
          </button>
          <button data-ui-control="true" onClick={runAutoSchedule} disabled={autoRunning}
            style={{...BS("#059669"),opacity:autoRunning?0.6:1,position:"relative",minWidth:160}}>
            {autoRunning
              ? <span style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{display:"inline-block",width:14,height:14,border:"2px solid rgba(255,255,255,0.4)",borderTopColor:"#fff",borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/>
                  รอบ {autoProgress?.run||0}/{autoProgress?.total||10}...
                </span>
              : "จัดตารางอัตโนมัติ"
            }
          </button>
        </div>
      </div>

      {proposal&&<section className="proposal-panel" aria-label="ผลเสนอการจัดตาราง">
        <div><strong>ผลเสนอพร้อมตรวจสอบ — ตารางเดิมยังไม่เปลี่ยน</strong><p>ตรวจรายการคาบใหม่ด้านล่าง แล้วเลือกนำไปใช้หรือยกเลิก</p></div>
        <div className="proposal-actions"><button data-ui-control="true" className="secondary" onClick={()=>{setProposal(null);setAutoResult(null)}}>ยกเลิกผลเสนอ</button><button data-ui-control="true" className="primary" onClick={()=>{
          const now=JSON.stringify({schedule:S.schedule,locks:S.locks,assigns:S.assigns,teachers:S.teachers,subjects:S.subjects,rooms:S.rooms,meetings:S.meetings});
          if(now!==proposal.base){st('ข้อมูลเปลี่ยนระหว่างจัดตาราง กรุณาสร้างผลเสนอใหม่','error');setProposal(null);return;}
          setPreviousAuto({schedule:S.schedule,applied:JSON.stringify(proposal.schedule)});
          U.setSchedule(proposal.schedule);setProposal(null);setAutoResult(null);st('นำผลจัดตารางไปใช้แล้ว');
        }}>นำผลนี้ไปใช้</button></div>
        <details><summary>ดูคาบที่เปลี่ยนแปลง</summary><div className="proposal-list">{[...new Set([...Object.keys(S.schedule),...Object.keys(proposal.schedule)])].filter(k=>JSON.stringify(S.schedule[k]||[])!==JSON.stringify(proposal.schedule[k]||[])).map(k=>{const [rid,day,pid]=k.split('_');return <div key={k}><strong>{S.rooms.find(r=>r.id===rid)?.name} · {day} คาบ {pid}</strong><span>{(S.schedule[k]||[]).map(e=>S.subjects.find(x=>x.id===e.subjectId)?.name).join(', ')||'ว่าง'} → {(proposal.schedule[k]||[]).map(e=>S.subjects.find(x=>x.id===e.subjectId)?.name).join(', ')||'ว่าง'}</span></div>})}</div></details>
      </section>}
      {previousAuto&&<div className="scheduler-help">นำผลอัตโนมัติไปใช้แล้ว <button data-ui-control="true" className="text-button" onClick={()=>{
        if(JSON.stringify(S.schedule)!==previousAuto.applied){st('มีการแก้ไขต่อจากผลอัตโนมัติแล้ว จึงไม่ย้อนทับงานที่แก้เพิ่ม','error');return;}
        U.setSchedule(previousAuto.schedule);setPreviousAuto(null);st('กลับสู่ตารางก่อนจัดอัตโนมัติแล้ว');
      }}>ย้อนกลับการจัดอัตโนมัติ</button></div>}
      {/* Auto result panel */}
      {autoResult&&(
        <div style={{background:autoResult.skipped===0?"#F0FDF4":"#FFFBEB",border:`1.5px solid ${autoResult.skipped===0?"#86EFAC":"#FDE68A"}`,borderRadius:12,padding:"12px 16px",marginBottom:12,display:"flex",gap:16,alignItems:"flex-start",flexWrap:"wrap"}}>
          <div style={{fontSize:13,fontWeight:700,color:autoResult.skipped===0?"#065F46":"#92400E",display:"flex",gap:12,flexWrap:"wrap",alignItems:"center"}}>
            {autoResult.skipped===0?"✅":"⚠️"}
            <span>จัดด้วย <strong>{autoResult.runs} รอบ</strong> — เสนอให้ลง <strong>{autoResult.placed}</strong> คาบ</span>
            {autoResult.skipped>0&&<span style={{color:"#DC2626"}}>| ข้ามไม่ได้ <strong>{autoResult.skipped}</strong> คาบ</span>}
          </div>
          {autoResult.details.length>0&&(
            <div style={{fontSize:11,color:"#92400E",flex:1}}>
              ❌ ไม่สามารถจัดได้: {autoResult.details.slice(0,5).join(", ")}{autoResult.details.length>5?` และอีก ${autoResult.details.length-5} รายการ`:""}
            </div>
          )}
          <button data-ui-control="true" onClick={()=>setAutoResult(null)} style={{background:"none",border:"none",cursor:"pointer",color:"#9CA3AF",fontSize:16}}>✕</button>
        </div>
      )}

      {/* Teacher summary bar */}
      {mode==="teacher"&&teacher&&(
        <div className="teacher-focus-card">
          <div style={{fontSize:15,fontWeight:700}}>{teacher.prefix}{teacher.firstName} {teacher.lastName}</div>
          <div style={{fontSize:12,color:"#6B7280"}}>{S.depts.find(d=>d.id===teacher.departmentId)?.name}</div>
          <div style={{marginLeft:"auto",display:"flex",gap:8}}>
            {[
              {label:"ได้รับ",val:teacher.totalPeriods||0,bg:"#DBEAFE",tx:"#1E40AF"},
              {label:"จัดแล้ว",val:teacherScheduledTotal(teacher.id),bg:"#FEF3C7",tx:"#92400E"},
              {label:"เหลือ",val:(teacher.totalPeriods||0)-teacherScheduledTotal(teacher.id),bg:(teacher.totalPeriods||0)-teacherScheduledTotal(teacher.id)>0?"#D1FAE5":"#FEE2E2",tx:(teacher.totalPeriods||0)-teacherScheduledTotal(teacher.id)>0?"#065F46":"#991B1B"},
            ].map(({label,val,bg,tx})=>(
              <div key={label} style={{background:bg,color:tx,padding:"4px 12px",borderRadius:8,fontWeight:700,fontSize:13}}>{label} {val}</div>
            ))}
          </div>
        </div>
      )}

      {/* Teacher mode */}
      {mode==="teacher"&&(teacher
        ?<div style={{display:"flex",flexDirection:"column",gap:0}}><div className="scheduler-layout">
            {/* Sidebar */}
            <div className="assignment-rail">
              <div className="color-guide"><small>ขอบสี = กลุ่มสาระ · ป้ายห้อง = ระดับชั้น</small></div><div className="rail-heading"><h3>วิชารอจัด</h3><span>{allAsgns.length} รายการ</span><p>ลากการ์ดลงคาบว่างในตาราง</p></div>
              {allAsgns.map(a=>{
                const sub=S.subjects.find(s=>s.id===a.subjectId);
                const dept=S.depts.find(d=>d.id===sub?.departmentId);
                // สีตามระดับชั้นของห้องแรก
                const LEVEL_COLORS_CARD=[
                  {bg:"#FFF7ED",border:"#FED7AA",head:"#EA580C",tx:"#9A3412"},
                  {bg:"#F0FDF4",border:"#BBF7D0",head:"#16A34A",tx:"#14532D"},
                  {bg:"#EFF6FF",border:"#BFDBFE",head:"#2563EB",tx:"#1E3A8A"},
                  {bg:"#FDF4FF",border:"#E9D5FF",head:"#9333EA",tx:"#581C87"},
                  {bg:"#FFF1F2",border:"#FECDD3",head:"#E11D48",tx:"#881337"},
                  {bg:"#F0FDFA",border:"#99F6E4",head:"#0D9488",tx:"#134E4A"},
                ];
                const firstRoom=S.rooms.find(r=>a.roomIds.includes(r.id));
                const lvIdx=S.levels.findIndex(l=>l.id===firstRoom?.levelId);
                const lc=LEVEL_COLORS_CARD[lvIdx>=0?lvIdx%LEVEL_COLORS_CARD.length:0];
                const u=aUsed(a.id);
                const subCa2=sub?.consecutiveAllowed||0;
                const totalForCard=subCa2===-2
                  ? (sub?.periodsPerWeek||2) * S.assigns.filter(x=>x.subjectId===a.subjectId).reduce((s,x)=>s+x.roomIds.length,0)
                  : a.totalPeriods;
                const rem=totalForCard-u;
                const coIds2=Array.isArray(cardCoMap[a.id])?cardCoMap[a.id]:(cardCoMap[a.id]?[cardCoMap[a.id]]:[]);
                // รวม co-teacher จาก schedule entries จริง
                const coIdsFromSchedule=new Set();
                Object.values(S.schedule).forEach(en=>(en||[]).forEach(e=>{
                  if(e.assignmentId===a.id){
                    const ids=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
                    ids.forEach(id=>coIdsFromSchedule.add(id));
                  }
                }));
                // merge: cardCoMap (UI) + schedule entries
                const allCoIds=[...new Set([...coIds2,...coIdsFromSchedule])];
                const coTeachers2=allCoIds.map(id=>S.teachers.find(t=>t.id===id)).filter(Boolean);
                const buns=bundleMap[a.id]||[];
                return (
                  <div key={a.id} className={"assignment-tile "+(rem<=0?"assignment-done":"")} style={{"--assignment-accent":departmentTone(dept).ink}}>
                    {/* ปุ่ม ⚙️ settings มุมขวาบน */}
                    <button data-ui-control="true"
                      onClick={()=>setShowGearId(showGearId===a.id?null:a.id)}
                      className="assignment-settings" aria-label="ตั้งค่าครูร่วมและวิชาคู่" title="ครูร่วม / วิชาคู่"
                      style={{position:"absolute",top:6,right:6,background:"rgba(0,0,0,0.07)",border:"none",borderRadius:6,width:22,height:22,cursor:"pointer",fontSize:12,display:"flex",alignItems:"center",justifyContent:"center",color:rem<=0?"#9CA3AF":lc.tx}}>⚙</button>

                    {coAsgnsIds.has(a.id)&&<div style={{fontSize:9,color:"#7C3AED",fontWeight:700,marginBottom:3}}>👥 ครูร่วม ({S.teachers.find(t=>t.id===a.teacherId)?.firstName||""})</div>}

                    <div
                      className="drag-card"
                      draggable={rem>0&&!coAsgnsIds.has(a.id)}
                      onDragStart={()=>setDragBoth({teacherId:selT,subjectId:a.subjectId,assignmentId:a.id})}
                      onDragEnd={()=>setDragBoth(null)}
                      style={{cursor:rem>0&&!coAsgnsIds.has(a.id)?"grab":"default",paddingRight:20}}
                    >
                      <span className="assignment-code">{sub?.code}</span>
                      <h4 className="assignment-title">{subDisplayName(sub)||sub?.code}</h4>
                      <div className="assignment-rooms">{a.roomIds.map(rid=><span key={rid} style={{color:levelTone(S.levels.find(l=>l.id===S.rooms.find(r=>r.id===rid)?.levelId)).ink,background:levelTone(S.levels.find(l=>l.id===S.rooms.find(r=>r.id===rid)?.levelId)).bg}}>{S.rooms.find(r=>r.id===rid)?.name}</span>)}</div>
                      <div className="assignment-requirements">{sub?.consecutiveAllowed>=2&&<span>{sub.consecutiveAllowed} คาบต่อเนื่อง</span>}{sub?.consecutiveAllowed===-1&&<span>NP</span>}{sub?.consecutiveAllowed===-2&&<span>เศรษฐ–วิศวะ</span>}{sub?.specialRoomId&&<span>ใช้ห้องพิเศษ</span>}</div>
                      <div className="assignment-progress"><span>{rem>0?'รอจัดอีก':'จัดครบแล้ว'} <strong>{Math.max(0,rem)}</strong> คาบ</span><small>{u}/{totalForCard}</small></div>
                      <div className="assignment-meter"><i style={{width:Math.min(100,totalForCard?u/totalForCard*100:0)+'%'}}/></div>
                      {/* สรุป co-teacher/bundle ย่อ */}
                      {(coTeachers2.length>0||buns.length>0)&&(
                        <div style={{marginTop:5,display:"flex",gap:4,flexWrap:"wrap"}}>
                          {coTeachers2.map(ct=><span key={ct.id} style={{fontSize:9,background:"rgba(124,58,237,0.12)",color:"#5B21B6",padding:"1px 6px",borderRadius:10,fontWeight:600}}>👥{ct.firstName}</span>)}
                          {buns.map((b,bi)=>{const bS=S.subjects.find(s=>s.id===S.assigns.find(x=>x.id===b.assignId)?.subjectId);return<span key={bi} style={{fontSize:9,background:"rgba(5,150,105,0.12)",color:"#065F46",padding:"1px 6px",borderRadius:10,fontWeight:600}}>📎{bS?.code||"?"}</span>;})}
                        </div>
                      )}
                    </div>

                    {/* Mini panel ⚙️ — ครูร่วม + วิชาคู่ */}
                    {showGearId===a.id&&(
                      <div style={{marginTop:8,padding:"8px 10px",background:"rgba(0,0,0,0.04)",borderRadius:8,border:`1px solid ${lc.border}`}}>
                        {/* ครูร่วม */}
                        <div style={{fontSize:10,fontWeight:700,color:lc.tx,marginBottom:5}}>👥 ครูร่วม</div>
                        {coTeachers2.map((ct2)=>{
                          const isFromSchedule=coIdsFromSchedule.has(ct2.id);
                          return<div key={ct2.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3}}>
                            <span style={{fontSize:10,color:lc.tx}}>
                              {ct2.firstName} {ct2.lastName}
                              {isFromSchedule&&<span style={{fontSize:8,color:"#059669",marginLeft:3}}>📅ในตาราง</span>}
                            </span>
                            <button data-ui-control="true" onClick={()=>{
                              // ลบออกจาก cardCoMap
                              setCardCoMap(p=>({...p,[a.id]:coIds2.filter(id=>id!==ct2.id)}));
                              // ถ้าลงตารางแล้ว ลบออกจาก schedule entries ด้วย
                              if(isFromSchedule){
                                U.setSchedule(prev=>{
                                  const next={};
                                  Object.entries(prev).forEach(([k,en])=>{
                                    next[k]=(en||[]).map(e=>{
                                      if(e.assignmentId!==a.id) return e;
                                      const newCoIds=(e.coTeacherIds||[]).filter(id=>id!==ct2.id);
                                      return{...e,coTeacherIds:newCoIds};
                                    });
                                  });
                                  return next;
                                });
                              }
                              st(`ลบ ${ct2.firstName} ออกจากครูร่วมแล้ว`,"warning");
                            }} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",padding:0,fontSize:12}}>✕</button>
                          </div>;
                        })}

                        {coTeachers2.length<4&&(
                          <button data-ui-control="true" onClick={()=>{ setShowGearId(null); setCardCoM(a.id); }} style={{fontSize:10,color:lc.head,background:"rgba(0,0,0,0.06)",border:`1px solid ${lc.border}`,borderRadius:6,padding:"3px 8px",cursor:"pointer",width:"100%",textAlign:"left",marginBottom:6}}>
                            + เพิ่มครูร่วม ({coTeachers2.length}/4)
                          </button>
                        )}
                        {/* วิชาคู่ */}
                        <div style={{fontSize:10,fontWeight:700,color:"#065F46",marginTop:4,marginBottom:5}}>📎 วิชาคู่</div>
                        {buns.map((b,bi)=>{
                          const bA=S.assigns.find(x=>x.id===b.assignId);
                          const bS=S.subjects.find(s=>s.id===bA?.subjectId);
                          const bT=S.teachers.find(t=>t.id===b.teacherId);
                          return<div key={bi} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3,background:"rgba(5,150,105,0.07)",borderRadius:4,padding:"2px 6px"}}>
                            <span style={{fontSize:9,color:"#065F46"}}>{bS?.code||""}{bT?` (${bT.firstName})`:""}</span>
                            <button data-ui-control="true" onClick={()=>setBundleMap(p=>({...p,[a.id]:buns.filter((_,i)=>i!==bi)}))} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",padding:0,fontSize:10}}>✕</button>
                          </div>;
                        })}
                        <button data-ui-control="true" onClick={()=>{ setShowGearId(null); setShowBundleM(a.id); setBundleSelSub(""); setBundleSelTeacher(""); }} style={{fontSize:10,color:"#059669",background:"rgba(5,150,105,0.08)",border:"1px solid #BBF7D0",borderRadius:6,padding:"3px 8px",cursor:"pointer",width:"100%",textAlign:"left"}}>+ เพิ่มวิชาคู่</button>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
            {renderTable(tRooms)}
          </div>
        </div>
        :<EmptyState icon="📋" title="เลือกครูเพื่อจัดตาราง"/>
      )}
      {mode==="teacher"&&teacher&&renderTeacherWeeklySummary()}

      {/* Room mode */}
      {mode==="room"&&(selRoom
        ?<div>{renderTable([selRoom])}</div>
        :<EmptyState icon="🏫" title="เลือกห้องเรียนเพื่อจัดตาราง"/>
      )}

      {/* Modal: co-teacher บนการ์ดที่วางแล้ว */}
      <Modal open={!!coM} onClose={()=>{setCoM(null);setCoS("");setCoDept("");}} title="เพิ่มครูสอนร่วม">
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <CoTeacherSelect coSVal={coS} setCoSFn={setCoS} coDeptVal={coDept} setCoDeptFn={setCoDept} excludeId={selT}/>
          {coS&&(()=>{
            const pts=coM?.key?.split("_")||[];
            const cDay=pts[1];const cPer=parseInt(pts[2]);
            const isBusy=teacherBusy(coS,cDay,cPer,null);
            const ct=S.teachers.find(t=>t.id===coS);
            const rem=(ct?.totalPeriods||0)-teacherScheduledTotal(coS);
            return <div>
              {isBusy&&<div style={{padding:10,background:"#FEE2E2",borderRadius:8,color:"#991B1B",fontSize:12,fontWeight:600,marginBottom:6}}>⚠️ ครูท่านนี้สอนคาบนี้อยู่แล้ว</div>}
              {rem<=0&&<div style={{padding:10,background:"#FEF3C7",borderRadius:8,color:"#92400E",fontSize:12,fontWeight:600,marginBottom:6}}>⚠️ คาบเต็มแล้ว</div>}
              <div style={{fontSize:12,color:"#6B7280"}}>จัดแล้ว {teacherScheduledTotal(coS)}/{ct?.totalPeriods||0} | เหลือ {rem}</div>
            </div>;
          })()}
          <button data-ui-control="true"
            onClick={()=>{
              if(!coS||!coM)return;
              const pts=coM.key.split("_");const cDay=pts[1];const cPer=parseInt(pts[2]);
              if(teacherBusy(coS,cDay,cPer,null)){st("ครูท่านนี้สอนคาบนี้อยู่แล้ว","error");return;}
              U.setSchedule(prev=>({...prev,[coM.key]:(prev[coM.key]||[]).map(e=>{
                if(e.id!==coM.entryId)return e;
                const existing=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
                if(existing.includes(coS))return e;
                if(existing.length>=4){st("ครูร่วมได้สูงสุด 4 คน","error");return e;}
                const newIds=[...existing,coS];
                return{...e,coTeacherIds:newIds,coTeacherId:newIds[0]||null};
              })}));
              setCoM(null);setCoS("");setCoDept("");st("เพิ่มครูร่วมสำเร็จ");
            }}
            style={BS()}>ยืนยัน</button>
        </div>
      </Modal>

      {/* Modal: co-teacher บน sidebar card */}
      <Modal open={!!cardCoM} onClose={()=>{setCardCoM(null);setCardCoS("");setCardCoDept("");}} title="กำหนดครูสอนร่วม (ติดไปกับการ์ด)">
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <div style={{fontSize:12,color:"#6B7280"}}>ครูร่วมจะถูกกำหนดทุกครั้งที่ลากการ์ดนี้ลงตาราง</div>
          <CoTeacherSelect coSVal={cardCoS} setCoSFn={setCardCoS} coDeptVal={cardCoDept} setCoDeptFn={setCardCoDept} excludeId={selT}/>
          <button data-ui-control="true"
            onClick={()=>{
              if(!cardCoS)return;
              setCardCoMap(p=>{
                const existing=Array.isArray(p[cardCoM])?p[cardCoM]:[];
                if(existing.includes(cardCoS))return p;
                if(existing.length>=4){st("ครูร่วมได้สูงสุด 4 คน","error");return p;}
                return{...p,[cardCoM]:[...existing,cardCoS]};
              });
              setCardCoM(null);setCardCoS("");setCardCoDept("");st("กำหนดครูร่วมสำเร็จ");
            }}
            style={BS()}>ยืนยัน</button>
        </div>
      </Modal>

      {/* Modal: วิชาคู่ (bundle) */}
      <Modal open={!!showBundleM} onClose={()=>setShowBundleM(null)} title="📎 กำหนดวิชาที่สอนคาบเดียวกัน">
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <div style={{fontSize:12,color:"#6B7280",background:"#F0FDF4",padding:"8px 12px",borderRadius:8,border:"1px solid #BBF7D0"}}>เมื่อลากการ์ดนี้ลงตาราง ระบบจะวางวิชาเหล่านี้ลงช่องเดียวกันด้วยอัตโนมัติ</div>
          {(bundleMap[showBundleM]||[]).length>0&&(
            <div style={{display:"flex",flexDirection:"column",gap:6}}>
              {(bundleMap[showBundleM]||[]).map((b,bi)=>{
                const bA=S.assigns.find(x=>x.id===b.assignId);
                const bS=S.subjects.find(s=>s.id===bA?.subjectId);
                const bT=S.teachers.find(t=>t.id===b.teacherId);
                return<div key={bi} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 12px",background:"#F0FDF4",borderRadius:10,border:"1px solid #BBF7D0"}}>
                  <div>
                    <div style={{fontSize:13,fontWeight:700,color:"#065F46"}}>{bS?.code} — {subDisplayName(bS)}</div>
                    <div style={{fontSize:11,color:"#6B7280"}}>ครู: {bT?`${bT.prefix}${bT.firstName} ${bT.lastName}`:"(ครูหลัก)"}</div>
                  </div>
                  <button data-ui-control="true" onClick={()=>setBundleMap(p=>({...p,[showBundleM]:(p[showBundleM]||[]).filter((_,i)=>i!==bi)}))} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",fontSize:16}}>✕</button>
                </div>;
              })}
            </div>
          )}
          <div style={{borderTop:"1px solid #E5E7EB",paddingTop:12}}>
            <div style={{fontSize:12,fontWeight:700,color:"#374151",marginBottom:8}}>เพิ่มวิชาคู่ใหม่</div>
            <div style={{display:"flex",flexDirection:"column",gap:10}}>
              <div>
                <label style={LS}>เลือก assignment วิชาคู่</label>
                <SearchSelect
                  value={bundleSelSub}
                  onChange={v=>{setBundleSelSub(v);setBundleSelTeacher("");}}
                  options={[{value:"",label:"-- เลือกวิชา --"},...S.assigns
                    .filter(a=>a.id!==showBundleM&&!(bundleMap[showBundleM]||[]).find(b=>b.assignId===a.id))
                    .map(a=>{
                      const sub=S.subjects.find(s=>s.id===a.subjectId);
                      const tch=S.teachers.find(t=>t.id===a.teacherId);
                      return{value:a.id,label:`${sub?.code||""} ${subDisplayName(sub)||""} — ${tch?.firstName||""} (${a.roomIds.map(r=>S.rooms.find(x=>x.id===r)?.name||"").join(",")})`};
                    })
                  ]}
                  placeholder="-- เลือก assignment --"
                />
              </div>
              {bundleSelSub&&(()=>{
                const bA=S.assigns.find(a=>a.id===bundleSelSub);
                const eligibleTeachers=S.assigns.filter(a=>a.subjectId===bA?.subjectId).map(a=>S.teachers.find(t=>t.id===a.teacherId)).filter(Boolean);
                return eligibleTeachers.length>1?<div>
                  <label style={LS}>ครูผู้สอน</label>
                  <SearchSelect value={bundleSelTeacher} onChange={v=>setBundleSelTeacher(v)}
                    options={[{value:"",label:"-- ใช้ครูหลักของ assignment --"},...eligibleTeachers.map(t=>({value:t.id,label:`${t.prefix}${t.firstName} ${t.lastName}`}))]}
                    placeholder="-- ใช้ครูหลัก --"/>
                </div>:null;
              })()}
              <button data-ui-control="true"
                onClick={()=>{
                  if(!bundleSelSub)return;
                  const bA=S.assigns.find(a=>a.id===bundleSelSub);if(!bA)return;
                  setBundleMap(p=>({...p,[showBundleM]:[...(p[showBundleM]||[]),{assignId:bundleSelSub,teacherId:bundleSelTeacher||bA.teacherId}]}));
                  setBundleSelSub("");setBundleSelTeacher("");
                }}
                disabled={!bundleSelSub}
                style={{...BS("#059669"),opacity:bundleSelSub?1:0.4}}>+ เพิ่มวิชาคู่</button>
            </div>
          </div>
        </div>
      </Modal>

      {/* ── Auto Schedule Modal ── */}
      {showAutoModal && (
        <div style={{position:"fixed",inset:0,zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.55)"}}>
          <div data-ui-surface="true" style={{background:"#fff",borderRadius:20,boxShadow:"0 30px 60px rgba(0,0,0,0.25)",width:"min(520px,94%)",maxHeight:"90vh",display:"flex",flexDirection:"column",overflow:"hidden",fontFamily:"'Sarabun','Noto Sans Thai',sans-serif"}}>
            {/* Header */}
            <div style={{background:"linear-gradient(135deg,#991B1B,#B91C1C)",padding:"20px 24px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <div>
                <div style={{color:"#fff",fontSize:17,fontWeight:700}}>⚡ Auto จัดตารางสอน</div>
                <div style={{color:"rgba(255,255,255,0.7)",fontSize:12,marginTop:2}}>เลือกเงื่อนไขก่อนกด "เริ่มจัด"</div>
              </div>
              <button data-ui-control="true" onClick={()=>setShowAutoModal(false)} style={{background:"rgba(255,255,255,0.15)",border:"none",borderRadius:8,padding:"6px 10px",cursor:"pointer",color:"#fff",fontSize:16}}>✕</button>
            </div>
            <div style={{padding:"20px 24px",overflowY:"auto",flex:1,display:"flex",flexDirection:"column",gap:18}}>
              {/* Section 1: Mode */}
              <div>
                <div style={{fontSize:13,fontWeight:700,color:"#374151",marginBottom:10}}>📌 วิธีจัด</div>
                <div style={{display:"flex",flexDirection:"column",gap:8}}>
                  {[
                    {val:"remaining", label:"เติมเฉพาะคาบที่ยังไม่ได้ลง", sub:"ปลอดภัยที่สุด — ไม่แตะคาบที่วางไว้แล้ว", badge:null, safe:true},
                    {val:"full",      label:"รีเซ็ตแล้วจัดใหม่ทั้งหมด",   sub:"จะลบทุกคาบที่ไม่ได้ล็อค แล้วจัดใหม่ตั้งแต่ต้น", badge:"⚠️ อันตราย", safe:false},
                  ].map(o=>(
                    <label key={o.val} style={{display:"flex",alignItems:"flex-start",gap:12,padding:"12px 14px",borderRadius:12,border:`2px solid ${autoOpts.mode===o.val?(o.safe?"#059669":"#DC2626"):"#E5E7EB"}`,background:autoOpts.mode===o.val?(o.safe?"#F0FDF4":"#FEF2F2"):"#F9FAFB",cursor:"pointer"}}>
                      <input data-ui-control="true" type="radio" name="autoMode" value={o.val} checked={autoOpts.mode===o.val} onChange={()=>setAutoOpts(p=>({...p,mode:o.val}))} style={{marginTop:2,accentColor:o.safe?"#059669":"#DC2626",flexShrink:0}}/>
                      <div>
                        <div style={{display:"flex",alignItems:"center",gap:8}}>
                          <span style={{fontSize:14,fontWeight:700,color:autoOpts.mode===o.val?(o.safe?"#065F46":"#991B1B"):"#374151"}}>{o.label}</span>
                          {o.badge&&<span style={{fontSize:10,background:"#FEE2E2",color:"#991B1B",padding:"1px 8px",borderRadius:20,fontWeight:700}}>{o.badge}</span>}
                        </div>
                        <div style={{fontSize:11,color:"#6B7280",marginTop:2}}>{o.sub}</div>
                      </div>
                    </label>
                  ))}
                </div>
              </div>
              {/* Section 2: ประเภทวิชา */}
              <div>
                <div style={{fontSize:13,fontWeight:700,color:"#374151",marginBottom:4}}>📚 ประเภทวิชาที่ให้ระบบจัด</div>
                <div style={{fontSize:11,color:"#6B7280",marginBottom:10}}>วิชายากแนะนำให้ลงเอง — ติ๊กเฉพาะที่ต้องการให้ระบบช่วย</div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                  {[
                    {key:"allowNormal", label:"วิชาปกติ",        sub:"ไม่มี consecutive", emoji:"📖", recommended:true},
                    {key:"allowConsec", label:"วิชาคาบติด",       sub:"consecutive ≥ 2",    emoji:"⚡", recommended:false},
                    {key:"allowNP",     label:"วิชา NP",          sub:"สอนหลายห้องพร้อมกัน", emoji:"🔀", recommended:false},
                    {key:"allowSR",     label:"วิชาห้องพิเศษ",    sub:"แล็บ, พละ, ศิลปะ ฯ", emoji:"🏫", recommended:false},
                  ].map(o=>(
                    <label key={o.key} style={{display:"flex",alignItems:"flex-start",gap:10,padding:"10px 12px",borderRadius:12,border:`2px solid ${autoOpts[o.key]?"#2563EB":"#E5E7EB"}`,background:autoOpts[o.key]?"#EFF6FF":"#F9FAFB",cursor:"pointer"}}>
                      <input data-ui-control="true" type="checkbox" checked={!!autoOpts[o.key]} onChange={e=>setAutoOpts(p=>({...p,[o.key]:e.target.checked}))} style={{marginTop:2,accentColor:"#2563EB",flexShrink:0}}/>
                      <div>
                        <div style={{display:"flex",alignItems:"center",gap:5}}>
                          <span style={{fontSize:14}}>{o.emoji}</span>
                          <span style={{fontSize:13,fontWeight:700,color:autoOpts[o.key]?"#1E40AF":"#374151"}}>{o.label}</span>
                          {o.recommended&&<span style={{fontSize:9,background:"#D1FAE5",color:"#065F46",padding:"1px 6px",borderRadius:20,fontWeight:700}}>แนะนำ</span>}
                        </div>
                        <div style={{fontSize:10,color:"#6B7280"}}>{o.sub}</div>
                      </div>
                    </label>
                  ))}
                </div>
              </div>
              {/* Section 3: เงื่อนไขเพิ่มเติม */}
              <div>
                <div style={{fontSize:13,fontWeight:700,color:"#374151",marginBottom:10}}>🛡️ เงื่อนไขเพิ่มเติม</div>
                <div style={{display:"flex",flexDirection:"column",gap:8}}>
                  {[
                    {key:"spreadDay",        label:"กระจายวิชา — ไม่ซ้ำวันเดิม",          sub:"วิชาเดียวกันในห้องเดิม จะไม่ถูกวาง 2 คาบในวันเดียว"},
                    {key:"noFirstLast",      label:"ไม่วางคาบ 1 + คาบ 7 วันเดิม (วิชาเดิม)", sub:"ป้องกันวิชาหนักอยู่หัว-ท้ายวันพร้อมกัน"},
                    {key:"maxPerDayTeacher", label:"ครูสอน 1 คาบ/วัน (กระจายทั้งสัปดาห์)",  sub:"ครูแต่ละคนจะไม่ถูกวางมากกว่า 1 คาบในวันเดียวกัน"},
                    {key:"noConsecTeacher",  label:"ห้ามครูสอนติดกัน 2 คาบขึ้นไป",          sub:"ทุกคาบของครูต้องมีช่วงพักคั่น — เข้มงวดมาก ควรใช้กับ run มากๆ"},
                    {key:"penalizeLunchGap", label:"หลีกเลี่ยงครูว่างช่วงพัก (คาบ 4+5) > 2 วัน", sub:"Soft constraint — run ที่ครูว่างพักกลางวันน้อยกว่าจะถูกเลือก"},
                  ].map(o=>(
                    <label key={o.key} style={{display:"flex",alignItems:"flex-start",gap:12,padding:"10px 14px",borderRadius:12,border:`2px solid ${autoOpts[o.key]?"#7C3AED":"#E5E7EB"}`,background:autoOpts[o.key]?"#F5F3FF":"#F9FAFB",cursor:"pointer"}}>
                      <input data-ui-control="true" type="checkbox" checked={!!autoOpts[o.key]} onChange={e=>setAutoOpts(p=>({...p,[o.key]:e.target.checked}))} style={{marginTop:2,accentColor:"#7C3AED",flexShrink:0}}/>
                      <div>
                        <span style={{fontSize:13,fontWeight:600,color:autoOpts[o.key]?"#5B21B6":"#374151"}}>{o.label}</span>
                        <div style={{fontSize:11,color:"#6B7280",marginTop:1}}>{o.sub}</div>
                      </div>
                    </label>
                  ))}
                  <div style={{padding:"10px 14px",borderRadius:12,border:`2px solid ${autoOpts.maxConsecTeacher>0?"#D97706":"#E5E7EB"}`,background:autoOpts.maxConsecTeacher>0?"#FFFBEB":"#F9FAFB"}}>
                    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                      <div>
                        <span style={{fontSize:13,fontWeight:600,color:autoOpts.maxConsecTeacher>0?"#92400E":"#374151"}}>⏱ ครูสอนติดกันสูงสุด</span>
                        <div style={{fontSize:11,color:"#6B7280",marginTop:1}}>0 = ไม่จำกัด</div>
                      </div>
                      <select data-ui-control="true" value={autoOpts.maxConsecTeacher} onChange={e=>setAutoOpts(p=>({...p,maxConsecTeacher:parseInt(e.target.value)}))} style={{padding:"6px 28px 6px 10px",border:"1.5px solid #D97706",borderRadius:8,fontSize:13,fontWeight:700,color:"#92400E",background:"#fff",cursor:"pointer",outline:"none",fontFamily:"inherit"}}>
                        <option value={0}>ไม่จำกัด</option>
                        <option value={1}>สูงสุด 1 คาบ (ไม่ติดกันเลย)</option>
                        <option value={2}>สูงสุด 2 คาบติด</option>
                        <option value={3}>สูงสุด 3 คาบติด</option>
                        <option value={4}>สูงสุด 4 คาบติด</option>
                      </select>
                    </div>
                  </div>
                </div>
              </div>
              {/* Section 4: จำนวนรอบ */}
              <div>
                <div style={{fontSize:13,fontWeight:700,color:"#374151",marginBottom:10}}>🔁 จำนวนรอบ (เลือกผลดีสุด)</div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                  {[
                    {val:1,label:"1 รอบ",sub:"เร็ว"},
                    {val:5,label:"5 รอบ",sub:"แนะนำ"},
                    {val:10,label:"10 รอบ",sub:"ดีที่สุด",highlight:true},
                    {val:20,label:"20 รอบ",sub:"ช้ามาก"},
                  ].map(o=>(
                    <button data-ui-control="true" key={o.val} onClick={()=>setAutoOpts(p=>({...p,runs:o.val}))} style={{flex:"1 1 80px",padding:"10px 8px",borderRadius:12,border:`2px solid ${autoOpts.runs===o.val?"#059669":"#E5E7EB"}`,background:autoOpts.runs===o.val?"#F0FDF4":"#F9FAFB",cursor:"pointer",fontFamily:"inherit"}}>
                      <div style={{fontSize:16,fontWeight:800,color:autoOpts.runs===o.val?"#065F46":"#374151"}}>{o.label}</div>
                      <div style={{fontSize:10,color:autoOpts.runs===o.val?"#059669":"#9CA3AF"}}>{o.sub}</div>
                      {o.highlight&&<div style={{fontSize:9,background:"#D1FAE5",color:"#065F46",padding:"1px 6px",borderRadius:20,fontWeight:700,marginTop:3,display:"inline-block"}}>default</div>}
                    </button>
                  ))}
                </div>
              </div>
              {/* Summary */}
              <div style={{background:"#F8FAFF",border:"1.5px solid #BFDBFE",borderRadius:12,padding:"12px 16px"}}>
                <div style={{fontSize:12,fontWeight:700,color:"#1E40AF",marginBottom:6}}>📋 สรุปการตั้งค่า</div>
                <div style={{fontSize:12,color:"#374151",display:"flex",flexDirection:"column",gap:3}}>
                  <span>{autoOpts.mode==="remaining"?"✅ เติมเฉพาะคาบที่ยังขาด":"⚠️ รีเซ็ตแล้วจัดใหม่ทั้งหมด"}</span>
                  <span>📚 จัดวิชา: {[autoOpts.allowNormal&&"ปกติ",autoOpts.allowConsec&&"คาบติด",autoOpts.allowNP&&"NP",autoOpts.allowSR&&"ห้องพิเศษ"].filter(Boolean).join(", ")||"— ยังไม่ได้เลือก"}</span>
                  <span>🔁 {autoOpts.runs} รอบ — ใช้ผลที่ดีที่สุด</span>
                  {autoOpts.maxConsecTeacher>0&&<span>⏱ ครูสอนติดกันไม่เกิน {autoOpts.maxConsecTeacher} คาบ</span>}
                  {autoOpts.maxPerDayTeacher&&<span>📅 ครูสอน 1 คาบ/วัน</span>}
                  {autoOpts.noConsecTeacher&&<span>🚫 ห้ามครูสอนติดกันเลย</span>}
                  {autoOpts.penalizeLunchGap&&<span>🍱 หลีกเลี่ยงว่างช่วงพักกลางวัน</span>}
                </div>
              </div>
            </div>
            {/* Footer */}
            <div style={{padding:"16px 24px",borderTop:"1px solid #E5E7EB",display:"flex",gap:10,justifyContent:"flex-end",background:"#FAFAFA"}}>
              <button data-ui-control="true" onClick={()=>setShowAutoModal(false)} style={BO()}>ยกเลิก</button>
              <button data-ui-control="true"
                onClick={()=>executeAutoSchedule(autoOpts)}
                disabled={!autoOpts.allowNormal&&!autoOpts.allowConsec&&!autoOpts.allowNP&&!autoOpts.allowSR}
                style={{...BS("#059669"),opacity:(!autoOpts.allowNormal&&!autoOpts.allowConsec&&!autoOpts.allowNP&&!autoOpts.allowSR)?0.4:1,cursor:(!autoOpts.allowNormal&&!autoOpts.allowConsec&&!autoOpts.allowNP&&!autoOpts.allowSR)?"not-allowed":"pointer"}}
              >
                ⚡ เริ่มจัดตาราง ({autoOpts.runs} รอบ)
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


/* ===== PDF: ตารางสอนรวมแบบตาราง ครูเป็นแถว × วัน/คาบเป็นคอลัมน์ ===== */
/* mode: "dept" = แยกกลุ่มสาระ, "level" = กรองระดับชั้น (levelId) */
function buildMasterTableHTML(S, ay, sh, filterLevelId) {
  // ตารางสอนครูรวม: ครูเป็นแถว × วัน/คาบเป็นคอลัมน์ (ขาวดำ)
  const subtitle = "ภาคเรียนที่ "+(ay?.semester||"1")+"/"+(ay?.year||"2568")+" "+(sh?.name||"โรงเรียนดาราวิทยาลัย");
  const logoHtml = sh?.logo ? '<img src="'+sh.logo+'" style="width:34px;height:34px;border-radius:50%;object-fit:cover;flex-shrink:0"/>' : '';
  const lvName = filterLevelId ? (S.levels.find(l=>l.id===filterLevelId)?.name||'') : '';
  const title = "ตารางสอนครู ปีการศึกษา "+(ay?.year||"2568")+(lvName?' — '+lvName:'');

  const getRoomShort = (rmId) => {
    const rm = S.rooms.find(r=>r.id===rmId);
    if(!rm) return null;
    if(filterLevelId && rm.levelId!==filterLevelId) return null;
    const m = rm.name.match(/(\d+\/\d+|\d+)$/);
    return m ? m[1] : rm.name;
  };
  const getTeacherCells = (tid) => {
    const cells = {};
    DAYS.forEach(d=>{ cells[d]={}; PERIODS.forEach(p=>{ cells[d][p.id]=[]; }); });
    Object.entries(S.schedule).forEach(([k,en])=>{
      en?.forEach(e=>{
        const mCoIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
        if(e.teacherId!==tid && !mCoIds.includes(tid)) return;
        const pts=k.split("_"); const rmId=pts.slice(0,pts.length-2).join("_");
        const day=pts[pts.length-2]; const per=parseInt(pts[pts.length-1]);
        const short=getRoomShort(rmId);
        if(short && cells[day] && cells[day][per]!==undefined) cells[day][per].push(short);
      });
    });
    return cells;
  };
  const teacherGroups = S.depts.map(dept=>{
    let ts = S.teachers.filter(t=>t.departmentId===dept.id&&(t.totalPeriods||0)>0);
    if(filterLevelId) ts=ts.filter(t=>{ const c=getTeacherCells(t.id); return DAYS.some(d=>PERIODS.some(p=>(c[d][p.id]||[]).length>0)); });
    return {dept,teachers:ts};
  }).filter(g=>g.teachers.length>0);

  const P=PERIODS.length; const totalCols=DAYS.length*P;
  let headRow1='<th rowspan="2" style="width:68px;border:1px solid #000;background:#e8e8e8;font-size:8px;font-weight:700;padding:2px 3px;text-align:center">ครูผู้สอน</th>';
  DAYS.forEach((day,di)=>{
    const br=di<DAYS.length-1?'border-right:2.5px solid #000;':'';
    headRow1+='<th colspan="'+P+'" style="border:1px solid #000;'+br+'background:#333;color:#fff;font-size:8px;font-weight:700;padding:2px 1px;text-align:center">'+day+'</th>';
  });
  let headRow2='';
  DAYS.forEach((_,di)=>{ PERIODS.forEach((p,pi)=>{
    const br=(pi===P-1&&di<DAYS.length-1)?'border-right:2.5px solid #000;':'';
    headRow2+='<th style="border:1px solid #999;'+br+'background:#e8e8e8;font-size:7px;font-weight:700;padding:1px;text-align:center;width:'+(540/totalCols).toFixed(1)+'px">'+p.id+'</th>';
  }); });

  let bodyHTML='';
  teacherGroups.forEach(({dept,teachers})=>{
    bodyHTML+='<tr><td colspan="'+(totalCols+1)+'" style="background:#555;color:#fff;font-size:8px;font-weight:700;padding:2px 5px;border:1px solid #000">'+dept.name+'</td></tr>';
    teachers.forEach((t,ti)=>{
      const cells=getTeacherCells(t.id);
      const rowBg=ti%2===0?'#fff':'#f5f5f5';
      let row='<tr>';
      row+='<td style="background:'+rowBg+';font-size:7.5px;padding:2px 3px;border:1px solid #000;white-space:nowrap;font-weight:600;vertical-align:middle">'+(t.prefix||"")+(t.firstName||"")+'<br/><span style="font-weight:400;font-size:7px">'+(t.lastName||"")+'</span></td>';
      DAYS.forEach((_,di)=>{ PERIODS.forEach((p,pi)=>{
        const rooms=cells[DAYS[di]]?.[p.id]||[];
        const isBlocked=(S.meetings||[]).some(m=>m.departmentId===t.departmentId&&m.day===DAYS[di]&&m.periods.includes(p.id));
        const br=(pi===P-1&&di<DAYS.length-1)?'border-right:2.5px solid #000;':'';
        let cellTxt=''; let extra='background:'+rowBg+';';
        if(isBlocked&&rooms.length===0){ cellTxt='🏫 ประชุม'; extra='background:#ddd;color:#555;font-size:7px;'; }
        else if(rooms.length>0){ cellTxt=rooms.join('<br/>'); extra='background:'+rowBg+';font-weight:700;'; }
        // คาบล็อคแผนก (custom) — แสดงเพิ่มถ้ายังว่าง
        const custLock=(S.meetings||[]).find(m=>m.type==="custom"&&(m.slots||[]).some(s=>s.day===DAYS[di]&&s.period===p.id));
        if(custLock&&!cellTxt){ cellTxt='🏫 '+custLock.name; extra='background:#FFF3E0;color:#E65100;font-size:7px;'; }
        row+='<td style="border:1px solid #ccc;'+br+extra+'font-size:7.5px;padding:1px 2px;text-align:center;vertical-align:middle;line-height:1.2">'+cellTxt+'</td>';
      }); });
      row+='</tr>'; bodyHTML+=row;
    });
  });

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
    +"@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');"
    +'@page{size:A4 landscape;margin:6mm 5mm}'
    +'*{margin:0;padding:0;box-sizing:border-box}'
    +"body{font-family:'Sarabun','Noto Sans Thai',sans-serif;color:#000;background:#fff}"
    +'.hdr{display:flex;align-items:center;gap:8px;margin-bottom:4px}'
    +'table{width:100%;border-collapse:collapse;table-layout:fixed}'
    +'@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}'
    +'</style></head><body>'
    +'<div class="hdr">'+logoHtml+'<div><div style="font-size:12px;font-weight:700">'+title+'</div><div style="font-size:9px;color:#444;margin-top:1px">'+subtitle+'</div></div></div>'
    +'<table><thead><tr>'+headRow1+'</tr><tr>'+headRow2+'</tr></thead><tbody>'+bodyHTML+'</tbody></table>'
    +'</body></html>';
}

/* ===== PDF: ตารางเรียนรวมระดับชั้น ห้องเป็นแถว × วัน/คาบ ===== */
function buildLevelTableHTML(S, ay, sh, filterLevelId) {
  const subtitle = "ภาคเรียนที่ "+(ay?.semester||"1")+"/"+(ay?.year||"2568")+" "+(sh?.name||"โรงเรียนดาราวิทยาลัย");
  const logoHtml = sh?.logo ? '<img src="'+sh.logo+'" style="width:30px;height:30px;border-radius:50%;object-fit:cover;flex-shrink:0"/>' : '';
  const lvName = filterLevelId ? (S.levels.find(l=>l.id===filterLevelId)?.name||'') : 'ทุกระดับ';
  const title = "ตารางเรียน "+lvName+" ปีการศึกษา "+(ay?.year||"2568");

  const sortKey=(r)=>{ const lv=S.levels.find(l=>l.id===r.levelId)?.name||""; const lvN=parseInt((lv.match(/(\d+)/)||[0,99])[1]); const rmN=parseInt((r.name.match(/(\d+)$/)||[0,0])[1]); return lvN*10000+rmN; };
  let rooms = [...S.rooms].sort((a,b)=>sortKey(a)-sortKey(b));
  if(filterLevelId) rooms=rooms.filter(r=>r.levelId===filterLevelId);
  if(!rooms.length) return '<html><body>ไม่มีห้องเรียนในระดับนี้</body></html>';

  const P=PERIODS.length; const totalCols=DAYS.length*P;
  // คำนวณความกว้าง cell: A4 landscape ~257mm - margin 10mm - col ห้อง ~14mm = 243mm / 35 col ≈ 6.9mm
  const cellW = (243/totalCols).toFixed(1);

  let headRow1='<th rowspan="2" style="width:14mm;border:1px solid #000;background:#333;color:#fff;font-size:7px;font-weight:700;padding:2px;text-align:center">ห้อง</th>';
  DAYS.forEach((day,di)=>{
    const br=di<DAYS.length-1?'border-right:2px solid #000;':'';
    headRow1+='<th colspan="'+P+'" style="border:1px solid #000;'+br+'background:#333;color:#fff;font-size:7px;font-weight:700;padding:2px 1px;text-align:center">'+day+'</th>';
  });
  let headRow2='';
  DAYS.forEach((_,di)=>{ PERIODS.forEach((p,pi)=>{
    const br=(pi===P-1&&di<DAYS.length-1)?'border-right:2px solid #000;':'';
    headRow2+='<th style="border:1px solid #bbb;'+br+'background:#e0e0e0;font-size:6px;font-weight:700;padding:1px;text-align:center;width:'+cellW+'mm">'+p.id+'</th>';
  }); });

  let bodyHTML='';
  const levelIds=[...new Set(rooms.map(r=>r.levelId))];
  levelIds.forEach(lvId=>{
    const lvRooms=rooms.filter(r=>r.levelId===lvId);
    const lvNameStr=S.levels.find(l=>l.id===lvId)?.name||'';
    bodyHTML+='<tr><td colspan="'+(totalCols+1)+'" style="background:#555;color:#fff;font-size:7px;font-weight:700;padding:2px 5px;border:1px solid #000">'+lvNameStr+'</td></tr>';
    lvRooms.forEach((rm,ri)=>{
      const rowBg=ri%2===0?'#fff':'#fafafa';
      let row='<tr>';
      // ชื่อห้องย่อ เช่น "ม.5/1" → "5/1"
      const rmShort=rm.name.replace(/[ม\.ป\.]/g,'').replace(/\s/g,'');
      row+='<td style="background:#e8e8e8;font-size:8px;padding:2px;border:1px solid #999;font-weight:700;text-align:center;vertical-align:middle;color:#000">'+rm.name+'</td>';
      DAYS.forEach((_,di)=>{ PERIODS.forEach((p,pi)=>{
        const key=rm.id+"_"+DAYS[di]+"_"+p.id;
        const en=S.schedule[key]||[];
        const br=(pi===P-1&&di<DAYS.length-1)?'border-right:2px solid #000;':'';
        let cellTxt=''; let extra='background:'+rowBg+';';
        if(en.length>0){
          cellTxt=en.map(e=>{
            const sub=S.subjects.find(s=>s.id===e.subjectId);
            const t=S.teachers.find(x=>x.id===e.teacherId);
            const coIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
            const coTs=coIds.map(id=>S.teachers.find(x=>x.id===id)).filter(Boolean);
            const subName=(sub?.shortName||sub?.name||'');
            const teacherNames=[t,...coTs].filter(Boolean).map(x=>"ครู"+(x.firstName||'')).join('+');
            return '<span style="font-weight:700">'+subName+'</span><br/>'+teacherNames;
          }).join('<hr style="border:none;border-top:1px dashed #bbb;margin:0"/>');
          extra='background:'+rowBg+';';
        }
        row+='<td style="border:1px solid #ddd;'+br+extra+'font-size:6px;padding:1px;text-align:center;vertical-align:middle;line-height:1.3;overflow:hidden">'+cellTxt+'</td>';
      }); });
      row+='</tr>'; bodyHTML+=row;
    });
  });

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
    +"@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');"
    +'@page{size:A4 landscape;margin:5mm 5mm}'
    +'*{margin:0;padding:0;box-sizing:border-box}'
    +"body{font-family:'Sarabun','Noto Sans Thai',sans-serif;color:#000;background:#fff;font-size:6px}"
    +'.hdr{display:flex;align-items:center;gap:6px;margin-bottom:3px}'
    +'table{width:100%;border-collapse:collapse;table-layout:fixed}'
    +'td,th{overflow:hidden;word-break:break-all}'
    +'@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}'
    +'</style></head><body>'
    +'<div class="hdr">'+logoHtml+'<div><div style="font-size:11px;font-weight:700">'+title+'</div><div style="font-size:8px;color:#444;margin-top:1px">'+subtitle+'</div></div></div>'
    +'<table><thead><tr>'+headRow1+'</tr><tr>'+headRow2+'</tr></thead><tbody>'+bodyHTML+'</tbody></table>'
    +'</body></html>';
}

/* ===== SWAP PAGE ===== */

function SwapPage({S,st,ay,sh}){
 const periodConfigs=Object.fromEntries(S.rooms.map(r=>[r.id,getPeriodCfg(getDivisionForLevel(r.levelId,S.levels))]));
 const teacherPeriods=Object.fromEntries(S.teachers.map(t=>{const ids=new Set();Object.entries(S.schedule).forEach(([key,en])=>{if(en?.some(e=>e.teacherId===t.id||(e.coTeacherIds||[]).includes(t.id)))ids.add(key.split('_').slice(0,-2).join('_'))});return [t.id,ids.size?[...ids].map(id=>periodConfigs[id]?.periods||PERIOD_CONFIG.default.periods):[getPeriodCfg(getDivisionForTeacher(t.id,S)).periods]]}));
 return <SwapWorkbench S={{...S,periodConfigs,teacherPeriods,defaultPeriods:PERIOD_CONFIG.default.periods,roles:SROLES}} st={st} ay={ay} sh={sh}/>;
}
/* ===== TEACHER TABLE FORMAT 3 ===== */
function buildTeacherTableHTML3(teacher,S,ay,sh){
  const yr=ay?.year||"2568";
  const logo=sh?.logo?'<img src="'+sh.logo+'" style="height:44px;vertical-align:middle;margin-right:8px;"/>':"";
  const tName=(teacher.prefix||"")+(teacher.firstName||"")+" "+(teacher.lastName||"");
  const dept=S.depts.find(d=>d.id===teacher.departmentId)?.name||"";
  const DAYS_T=["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
  const PIDS_T=[1,2,3,4,5,6,7];
  // master table ใช้ default (m2/ม.ปลาย) เพราะครูอาจสอนหลาย division
  const PTIMES=PERIOD_CONFIG.default.periods.map(p=>p.time);
  const getCell=(day,pid)=>{
    const out=[];
    S.rooms.forEach(room=>{
      const key=room.id+"_"+day+"_"+pid;
      (S.schedule[key]||[]).forEach(e=>{
        if(e.teacherId!==teacher.id&&!(e.coTeacherIds||[]).includes(teacher.id))return;
        if(e.isLock)out.push({type:"lock",roomName:room.name});
        else out.push({type:"class",roomShort:room.name});
      });
    });
    (S.meetings||[]).forEach(m=>{if(m.teacherId===teacher.id&&m.day===day&&(m.periods||[]).includes(pid))out.push({type:"meeting",label:m.label||"Lock"});});
    return out;
  };
  const getHomeroom=(day)=>{
    const meets=(S.meetings||[]).find(m=>m.teacherId===teacher.id&&m.day===day&&(m.isAssembly||m.isHomeroom||(m.periods||[]).includes(0)));
    if(meets)return meets.isAssembly?"เข้าหอประชุม":(meets.label||"Homeroom");
    return"โฮมรูม";
  };
  const vert=(txt,fs="9pt")=>'<div style="writing-mode:vertical-rl;transform:rotate(180deg);white-space:nowrap;font-size:'+fs+';font-weight:600;letter-spacing:1px;text-align:center;">'+txt+'</div>';
  const thS="border:1px solid #555;text-align:center;vertical-align:middle;font-size:7pt;font-weight:bold;padding:1px;";
  const brkS="border:1px solid #555;background:#f9f9e8;padding:0;vertical-align:middle;text-align:center;width:22px;";  const hdr='<tr style="background:#f0f0f0;"><th rowspan="2" style="'+thS+'width:52px;position:relative;min-height:40px;"><svg style="position:absolute;top:0;left:0;width:100%;height:100%;" preserveAspectRatio="none"><line x1="0" y1="0" x2="100%" y2="100%" stroke="#888" stroke-width="0.8"/></svg><span style="position:absolute;top:3px;right:4px;font-size:8pt;font-weight:600;color:#333;">เวลา</span><span style="position:absolute;bottom:3px;left:4px;font-size:8pt;font-weight:600;color:#333;">วัน</span></th><th rowspan="2" style="'+thS+'width:60px;">08:00<br/>08:30</th><th style="'+thS+'height:20px;">คาบ 1</th><th style="'+thS+'height:20px;">คาบ 2</th><th rowspan="2" style="'+brkS+'">'+vert("10.10-10.25")+'</th><th style="'+thS+'height:20px;">คาบ 3</th><th style="'+thS+'height:20px;">คาบ 4</th><th rowspan="2" style="'+brkS+'">'+vert("12.05-13.00")+'</th><th style="'+thS+'height:20px;">คาบ 5</th><th style="'+thS+'height:20px;">คาบ 6</th><th rowspan="2" style="'+brkS+'">'+vert("13.50-14.00")+'</th><th style="'+thS+'height:20px;">คาบ 7</th></tr><tr style="background:#f0f0f0;height:14px;max-height:14px;">'+PTIMES.map(t=>'<td style="border:1px solid #888;font-size:7pt;text-align:center;padding:0px 1px;height:14px;">'+t+'</td>').join("")+'</tr>';
  const CELL_H="22px";
  const renderCell=(cells,multi)=>{
    const bg=multi?"background:#eeeeee;":"";
    if(!cells.length)return'<td style="border:1px solid #ccc;padding:0;'+bg+'"><div style="height:'+CELL_H+';"></div></td>';
    const inner=cells.map(c=>{
      if(c.type==="lock"||c.type==="meeting")return'<div style="color:#cc0000;font-size:7pt;font-weight:bold;">'+(c.label||"Lock")+'</div>';
      return'<div style="font-size:9pt;font-weight:bold;">'+c.roomShort+'</div>';
    }).join('');
    return'<td style="border:1px solid #ccc;padding:0;'+bg+'"><div style="height:'+CELL_H+';overflow:hidden;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;">'+inner+'</div></td>';
  };
  let body="";
  DAYS_T.forEach((day,di)=>{
    const hmTxt=getHomeroom(day);const hmBg="#fafff7";const bgRow=di%2===0?"":"background:#fafafa;";
    const cells=PIDS_T.map(pid=>getCell(day,pid));
    const multi=PIDS_T.map((pid,i)=>{let cnt=0;S.rooms.forEach(room=>{cnt+=(S.schedule[room.id+"_"+day+"_"+pid]||[]).filter(e=>e.teacherId===teacher.id||(e.coTeacherIds||[]).includes(teacher.id)).length;});return cnt>1;});
    const dDisp=day==="พฤหัสบดี"?"พฤหัส":day;
    body+='<tr style="'+bgRow+'"><td style="border:1px solid #888;padding:0;background:#f5f5f5;"><div style="height:'+CELL_H+';display:flex;align-items:center;justify-content:center;font-weight:bold;font-size:9pt;text-align:center;">'+dDisp+'</div></td><td style="border:1px solid #888;padding:0;background:'+hmBg+';"><div style="height:'+CELL_H+';display:flex;align-items:center;justify-content:center;text-align:center;font-size:8pt;font-weight:600;line-height:1.2;">'+hmTxt+'</div></td>'+renderCell(cells[0],multi[0])+renderCell(cells[1],multi[1])+(di===0?'<td rowspan="5" style="'+brkS+'">'+vert("พักน้อย 15 นาที")+'</td>':"")+renderCell(cells[2],multi[2])+renderCell(cells[3],multi[3])+(di===0?'<td rowspan="5" style="'+brkS+'">'+vert("พักกลางวัน 55 นาที")+'</td>':"")+renderCell(cells[4],multi[4])+renderCell(cells[5],multi[5])+(di===0?'<td rowspan="5" style="'+brkS+'">'+vert("พักน้อย 10 นาที")+'</td>':"")+renderCell(cells[6],multi[6])+'</tr>';
  });
  const assigns=S.assigns.filter(a=>a.teacherId===teacher.id);
  const specialMeets=[...new Set((S.meetings||[]).filter(m=>m.teacherId===teacher.id&&m.label&&!m.isAssembly&&!m.isHomeroom).map(m=>m.label))];
  let summaryRows="";let grandTotal=0;
  assigns.forEach(a=>{
    const sub=S.subjects.find(s=>s.id===a.subjectId);if(!sub)return;
    const rooms=(a.roomIds||[]).map(rid=>S.rooms.find(r=>r.id===rid)).filter(Boolean);
    const rCount=rooms.length;const pPerRoom=sub.periodsPerWeek||Math.round((a.totalPeriods||0)/Math.max(rCount,1));
    const total=a.totalPeriods||(pPerRoom*rCount);grandTotal+=total;
    const roomNames=rooms.map(r=>r.name).join(", ");
    const codes=sub.code?"("+sub.code+")":"";
    summaryRows+='<tr><td style="padding:1px 4px;font-size:9.5pt;">'+(sub.name||"")+" "+codes+'</td><td style="padding:1px 4px;font-size:9.5pt;text-align:right;">'+rCount+' ห้อง</td><td style="padding:1px 4px;font-size:9.5pt;text-align:center;">×</td><td style="padding:1px 4px;font-size:9.5pt;text-align:right;">'+pPerRoom+' คาบ</td><td style="padding:1px 4px;font-size:9.5pt;text-align:center;">=</td><td style="padding:1px 4px;font-size:9.5pt;text-align:right;font-weight:bold;">'+total+'</td><td style="padding:1px 4px;font-size:9.5pt;">คาบ</td></tr>';
  });
  specialMeets.forEach(l=>{summaryRows+='<tr><td colspan="8" style="padding:1px 4px;font-size:9.5pt;">'+l+'</td></tr>';});
  summaryRows+='<tr><td colspan="4" style="padding:1px 4px;font-size:9.5pt;text-align:right;border-top:1px solid #999;">รวม</td><td style="padding:1px 4px;border-top:1px solid #999;text-align:center;">=</td><td style="padding:1px 4px;font-size:9.5pt;text-align:right;font-weight:bold;border-top:2px double #333;">'+grandTotal+'</td><td style="padding:1px 4px;font-size:9.5pt;border-top:1px solid #999;">คาบ</td></tr>';
  return'<div style="font-family:\'TH SarabunNew\',\'Sarabun\',sans-serif;page-break-inside:avoid;">'
    +'<div style="text-align:center;margin-bottom:4px;">'+logo+'<span style="font-size:11pt;font-weight:bold;">ตารางสอน ปีการศึกษา '+yr+'</span></div>'
    +'<table style="width:100%;border-collapse:collapse;table-layout:fixed;margin-bottom:2px;"><colgroup><col style="width:44px;"><col style="width:52px;"><col><col><col style="width:22px;"><col><col><col style="width:22px;"><col><col><col style="width:22px;"><col></colgroup><thead>'+hdr+'</thead><tbody>'+body+'</tbody></table>'
    +'<table style="width:100%;border-collapse:collapse;margin-top:2px;font-size:8pt;"><tbody><tr valign="top">'
    +'<td style="width:40%;padding-right:8px;">'
      +'<div style="color:#1a237e;font-weight:bold;margin-bottom:2px;">กลุ่มสาระการเรียนรู้ '+dept+'</div>'
      +'<div><b>อาจารย์ผู้สอน</b> '+tName+'</div>'
      +assigns.map(a=>{const sub=S.subjects.find(s=>s.id===a.subjectId);if(!sub)return"";return'<div style="padding-left:4px;">'+(sub.name||"")+" "+(sub.code?"("+sub.code+")":"")+"</div>";}).join("")
      +specialMeets.map(l=>'<div style="padding-left:4px;">'+l+'</div>').join("")
    +'</td>'
    +'<td style="width:60%;">'
      +'<table style="font-size:8pt;border-collapse:collapse;width:100%;"><tbody>'+summaryRows+'</tbody></table>'
    +'</td>'
    +'</tr></tbody></table>'
    +'</div>';
}

/* ===== REPORTS ===== */


function Reports({S,U,st,gc,ay,sh}){
  const fileRefSched=useRef(null);
  const [printSettings,setPrintSettings]=useState(loadPrintSettings);
  const [showPrintSettings,setShowPrintSettings]=useState(false);
  const [showPrintDesigner,setShowPrintDesigner]=useState(false);
  const [printPreview,setPrintPreview]=useState(null);
  const [reportTab,setReportTab]=useState("print");
  const [selTeacherPDF,setSelTeacherPDF]=useState("");
  const [selRoomPDF,setSelRoomPDF]=useState("");
  const [selTeacherXL,setSelTeacherXL]=useState("");
  const [selRoomXL,setSelRoomXL]=useState("");
  const [showNewRoomPDF,setShowNewRoomPDF]=useState(false);
  const [newRoomPDFOpts,setNewRoomPDFOpts]=useState({selectedRooms:[],layout:"2portrait"});
  const [showNewTeacherPDF,setShowNewTeacherPDF]=useState(false);
  const [selectedTeachersPDF,setSelectedTeachersPDF]=useState([]);
  const [teacherSearchQ,setTeacherSearchQ]=useState("");
  const [showExcelModal,setShowExcelModal]=useState(false);
  const [excelSelectedRooms,setExcelSelectedRooms]=useState([]);
  const roomSt=S.rooms.map(rm=>{let f=0;DAYS.forEach(d=>PERIODS.forEach(p=>{const k=`${rm.id}_${d}_${p.id}`;if(S.schedule[k]?.length)f++}));const total=DAYS.length*PERIODS.length;return{room:rm,filled:f,total,pct:Math.round(f/total*100)}});
  const teacherSt=S.teachers.map(t=>{
    const tot=t.totalPeriods||0;
    const seen=new Set(); let u=0;
    Object.entries(S.schedule).forEach(([k,en])=>{
      const pts=k.split("_");
      en?.forEach(e=>{
        const rCoIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
        if(e.teacherId!==t.id&&!rCoIds.includes(t.id))return;
        const sub=S.subjects.find(s=>s.id===e.subjectId);
        const ca=sub?.consecutiveAllowed||0;
        if(ca===-1||ca===-2){const npKey=e.subjectId+"_"+pts[pts.length-2]+"_"+pts[pts.length-1];if(!seen.has(npKey)){seen.add(npKey);u++;}}
        else u++;
      });
    });
    return{teacher:t,tot,used:u,rem:tot-u};
  });

  // Export schedule → JSON file (เก็บทุก entry ครบถ้วน)
  // พิมพ์ตารางสอนครูแบบใหม่ (เหมือน room format, 2 คน/หน้า A4 แนวตั้ง)
  const printTeacherPDFNew=(teachers)=>{
    const list=Array.isArray(teachers)?teachers:[teachers];
    if(!list.length){st("ไม่มีครูที่เลือก","error");return;}
    setPrintPreview({html:buildF2Html(list,S,ay,sh,printSettings)});
  };
  // Export ตารางห้องเรียน ตาม format import_Schedule.xlsx
  const exportScheduleJSON=()=>{
    const data={version:1,exportedAt:new Date().toISOString(),schedule:S.schedule,locks:S.locks,assigns:S.assigns,teachers:S.teachers,subjects:S.subjects,rooms:S.rooms,levels:S.levels,plans:S.plans,depts:S.depts,meetings:S.meetings,specialRooms:S.specialRooms};
    const blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});
    const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`backup_timetable_${new Date().toISOString().slice(0,10)}.json`;a.click();
    st("Backup สำเร็จ");
  };

  const exportRoomScheduleXLSX=async(rooms)=>{
    const roomList=Array.isArray(rooms)?rooms:(rooms?[rooms]:S.rooms);
    if(!roomList.length){st("ไม่มีห้องเรียน","error");return;}
    st("กำลังโหลด library...","warning");

    // โหลด SheetJS — ลอง unpkg ถ้า cdnjs ไม่ผ่าน
    let XLib=window.XLSX;
    if(!XLib){
      for(const src of[
        "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
        "https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js",
      ]){
        try{
          await new Promise((res,rej)=>{
            const s=document.createElement("script");
            s.src=src; s.onload=res;
            s.onerror=()=>rej(new Error("fail"));
            document.head.appendChild(s);
          });
          XLib=window.XLSX;
          if(XLib) break;
        }catch(e){continue;}
      }
    }
    if(!XLib){st("โหลด library ไม่สำเร็จ กรุณาตรวจสอบ internet","error");return;}

    const DAYS_TH=["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];

    const wb=XLib.utils.book_new();
    roomList.forEach(room=>{
      // ใช้เวลาคาบตาม division ของห้องนั้น
      const roomDivId=getDivisionForRoom(room,S);
      const PERIOD_TIMES=getPeriodCfg(roomDivId).periods.map(p=>{
        const [start,end]=(p.time||"").split("-");
        return {id:p.id,start:start||"",end:end||""};
      });
      const rows=[["วัน","รหัสวิชา","เริ่มเวลา","หมดเวลา","รหัสผู้ใช้งาน(ครูผู้สอน)"]];
      DAYS_TH.forEach(day=>{
        PERIOD_TIMES.forEach(p=>{
          const entries=S.schedule[room.id+"_"+day+"_"+p.id]||[];
          if(!entries.length){
            rows.push([day,"",p.start,p.end,""]);
          }else{
            entries.forEach(e=>{
              const sub=S.subjects.find(s=>s.id===e.subjectId);
              const tch=S.teachers.find(t=>t.id===e.teacherId);
              rows.push([day,sub?.code||"",p.start,p.end,tch?.teacherCode||""]);
            });
          }
        });
      });
      const ws=XLib.utils.aoa_to_sheet(rows);
      ws["!cols"]=[{wch:12},{wch:15},{wch:10},{wch:10},{wch:25}];
      const sheetName=room.name.replace(/[:\\\/\?\*\[\]]/g,"").trim().slice(0,31)||"Room"+i;
      XLib.utils.book_append_sheet(wb,ws,sheetName);
    });
    const fname=roomList.length===1
      ?`ตารางสอน_${roomList[0].name}_${ay?.year||"2568"}.xlsx`
      :`ตารางสอนห้องเรียน_${ay?.year||"2568"}.xlsx`;
    XLib.writeFile(wb,fname);
    st(`✅ Export ${roomList.length} ห้อง สำเร็จ`);
  };

  // Import schedule จาก JSON
  const importScheduleJSON=async(e)=>{
    const f=e.target.files?.[0];if(!f)return;
    try{
      const txt=await f.text();
      const data=JSON.parse(txt);
      // ตรวจ format
      if(typeof data !== "object"||(!data.schedule&&!data.assigns)){
        st("ไฟล์ไม่ถูกต้อง — ต้องเป็น JSON ที่ backup จากระบบนี้","error");
        e.target.value="";return;
      }
      if(!await uiConfirm(`Restore ตารางสอน?\n\nไฟล์: ${f.name}\nบันทึกเมื่อ: ${data.exportedAt||"ไม่ทราบ"}\n\n⚠️ ข้อมูลตารางสอนปัจจุบันจะถูกทับ`))return;

      // restore ทีละ field — ใช้ set functions โดยตรงเพื่อ trigger Firebase sync
      if(data.schedule) U.setSchedule(data.schedule);
      if(data.locks)    U.setLocks(data.locks);
      if(data.assigns?.length){
        // merge: เก็บ assigns ปัจจุบันที่ไม่มีใน backup ไว้ + เอา backup มาทับ
        U.setAssigns(prev=>{
          const kept=prev.filter(a=>!data.assigns.find(x=>x.id===a.id));
          return [...kept,...data.assigns];
        });
      }
      // restore ข้อมูลอื่นๆ ถ้ามี (full backup)
      if(data.teachers?.length)     U.setTeachers(data.teachers);
      if(data.subjects?.length)     U.setSubjects(data.subjects);
      if(data.rooms?.length)        U.setRooms(data.rooms);
      if(data.levels?.length)       U.setLevels(data.levels);
      if(data.plans?.length)        U.setPlans(data.plans);
      if(data.depts?.length)        U.setDepts(data.depts);
      if(data.meetings?.length)     U.setMeetings(data.meetings);
      if(data.specialRooms?.length) U.setSpecialRooms(data.specialRooms);

      st(`✅ Restore สำเร็จ — ${f.name}`);
    }catch(err){
      st("อ่านไฟล์ไม่ได้: "+err.message,"error");
    }
    e.target.value="";
  };


  const exportRoomXL=(rm)=>{
    const pcfg=getPeriodCfg(getDivisionForRoom(rm,S));
    const h=["วัน",...pcfg.periods.map(p=>`คาบ${p.id}(${p.time})`)];
    const d=DAYS.map(day=>[day,...PERIODS.map(p=>{const en=S.schedule[`${rm.id}_${day}_${p.id}`]||[];return en.map(e=>{const sub=S.subjects.find(s=>s.id===e.subjectId);const t=S.teachers.find(x=>x.id===e.teacherId);return`${sub?.code||""} ${subDisplayName(sub)||""} (${t?.firstName||""})`}).join(" / ")})]);
    exportExcel(h,d,`ตารางเรียน_${rm.name}.xlsx`,rm.name);st(`Export ${rm.name}`);
  };

  const exportTeacherXL=(t)=>{
    const tDiv=getDivisionForTeacher(t.id,S);
    const tPcfg=getPeriodCfg(tDiv);
    const h=["วัน",...tPcfg.periods.map(p=>`คาบ${p.id}(${p.time})`)];const d=DAYS.map(day=>[day,...PERIODS.map(p=>{let parts=[];Object.entries(S.schedule).forEach(([k,en])=>{if(!k.endsWith(`_${day}_${p.id}`))return;en?.forEach(e=>{const xCoIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);if(e.teacherId===t.id||xCoIds.includes(t.id)){const sub=S.subjects.find(s=>s.id===e.subjectId);const rid=k.split("_")[0];const rm=S.rooms.find(r=>r.id===rid);parts.push(`${sub?.code||""} ${sub?.name||""} (${rm?.name||""})`)}})});return parts.join(" / ")})]);exportExcel(h,d,`ตารางสอน_${t.prefix}${t.firstName}.xlsx`,"ตารางสอน");st(`Export ${t.firstName}`)};

  const exportAllRooms=()=>{
    exportExcelMulti(S.rooms.map(rm=>{
      const pcfg=getPeriodCfg(getDivisionForRoom(rm,S));
      return {
        name:rm.name,
        headers:["วัน",...pcfg.periods.map(p=>`คาบ${p.id}(${p.time})`)],
        rows:DAYS.map(day=>[day,...PERIODS.map(p=>{const en=S.schedule[`${rm.id}_${day}_${p.id}`]||[];return en.map(e=>{const sub=S.subjects.find(s=>s.id===e.subjectId);const t=S.teachers.find(x=>x.id===e.teacherId);return`${sub?.code||""} ${subDisplayName(sub)||""} (${t?.firstName||""})`}).join(" / ")})])
      };
    }),"ตารางเรียนทุกห้อง.xlsx");
    st("Export ทุกห้อง");
  };

  const exportAllTeachers=()=>{
    const sheets=S.teachers.map(t=>{
      const tDivA=getDivisionForTeacher(t.id,S);
      const tPcfgA=getPeriodCfg(tDivA);
      const headers=["วัน",...tPcfgA.periods.map(p=>"คาบ"+p.id+"("+p.time+")")];
      const rows=DAYS.map(day=>[day,...PERIODS.map(p=>{
        let parts=[];
        Object.entries(S.schedule).forEach(([k,en])=>{
          if(!k.endsWith("_"+day+"_"+p.id))return;
          en?.forEach(e=>{
            if(e.teacherId===t.id||e.coTeacherId===t.id){
              const sub=S.subjects.find(s=>s.id===e.subjectId);
              const rid=k.split("_")[0];
              const rm=S.rooms.find(r=>r.id===rid);
              parts.push((sub?.code||"")+" "+(sub?.name||"")+" ("+(rm?.name||"")+")");
            }
          });
        });
        return parts.join(" / ");
      })]);
      return {name:t.firstName+" "+t.lastName,headers,rows};
    });
    exportExcelMulti(sheets,"ตารางสอนทุกคน.xlsx");
    st("Export ทุกคน");
  };

  const exportStatus=()=>{
    const sheets=[{name:"ห้องเรียน",headers:["ห้อง","จัดแล้ว","ทั้งหมด","%"],rows:roomSt.map(r=>[r.room.name,r.filled,r.total,`${r.pct}%`])},{name:"ครู",headers:["ชื่อ","คาบได้รับ","จัดแล้ว","เหลือ","สถานะ"],rows:teacherSt.filter(t=>t.tot>0).map(t=>[`${t.teacher.prefix}${t.teacher.firstName} ${t.teacher.lastName}`,t.tot,t.used,t.rem,t.rem===0?"ครบ":"เหลือ "+t.rem])}];
    exportExcelMulti(sheets,"รายงานสถานะ.xlsx");st("Export สำเร็จ");
  };

  // PDF print for teacher
  // PDF: ตารางสอนครู — แสดง วิชา + ห้อง (ไม่มีครูร่วม)
  const printTeacherPDF=(t)=>{
    const sortParts=(parts)=>parts.sort((a,b)=>{
      const numA=parseInt((a.room.match(/(\d+)$/)||[0,9999])[1]);
      const numB=parseInt((b.room.match(/(\d+)$/)||[0,9999])[1]);
      if(numA!==numB) return numA-numB;
      return a.room.localeCompare(b.room,"th");
    });
    const tDiv=getDivisionForTeacher(t.id,S);

    // helper: หา lock cell สำหรับครูคนนี้ในแต่ละ day/period
    const getLockCell=(day,pid)=>{
      // 1) คาบล็อคส่วนตัว (personalLocks)
      for(const pl of (t.personalLocks||[])){
        if(pl.day===day&&(pl.periods||[]).includes(pid))
          return [{sub:"🔒 "+(pl.reason||"ส่วนตัว"),room:"",room2:"",isLock:true,lockColor:"#FEF3C7",lockTextColor:"#92400E"}];
      }
      // 2) หน้าที่พิเศษ (specialRoles → วิชาการ/พัฒนาวินัย)
      for(const rid of (t.specialRoles||[])){
        const role=SROLES.find(r=>r.id===rid);
        const bl=(role?.blocked||[]).find(b=>b.day===day&&(b.periods||[]).includes(pid));
        if(bl) return [{sub:"📋 "+role.name.replace("ฝ่ายวิชาการ","วิชาการ").replace("ฝ่ายพัฒนาวินัย","พัฒนาวินัย"),room:"",room2:"",isLock:true,lockColor:"#EDE9FE",lockTextColor:"#5B21B6"}];
      }
      // 3) คาบล็อคกลุ่มสาระ (meetings type=dept)
      for(const m of (S.meetings||[])){
        if(m.type&&m.type!=="dept") continue;
        if(m.isAssembly||m.isHomeroom) continue;
        const isMyDept=!m.departmentId||m.departmentId===t.departmentId||m.departmentId==="all";
        if(isMyDept&&m.day===day&&(m.periods||[]).includes(pid))
          return [{sub:"🔒 LOCK",room:"",room2:"",isLock:true,lockColor:"#DBEAFE",lockTextColor:"#1D4ED8"}];
      }
      // 4) คาบ custom lock
      for(const m of (S.meetings||[])){
        if(m.type!=="custom") continue;
        const inSlot=(m.slots||[]).some(s=>s.day===day&&s.period===pid);
        if(inSlot) return [{sub:"🏫 "+(m.name||"Lock"),room:"",room2:"",isLock:true,lockColor:"#F3F4F6",lockTextColor:"#374151"}];
      }
      // 5) ชมรม/Assembly (meetings ที่มี isAssembly/isHomeroom และ teacherId ตรง)
      for(const m of (S.meetings||[])){
        if(m.teacherId===t.id&&(m.isAssembly||m.isHomeroom)&&m.day===day&&(m.periods||[]).includes(pid))
          return [{sub:m.isAssembly?"🎤 หอประชุม":"🏠 Homeroom",room:"",room2:"",isLock:true,lockColor:"#FEF9C3",lockTextColor:"#92400E"}];
      }
      return null;
    };

    const title="ตารางสอน "+(t.prefix||"")+(t.firstName||"")+" "+(t.lastName||"");
    const subtitle="ภาคเรียนที่ "+(ay?.semester||"1")+"/"+(ay?.year||"2568")+" "+(sh?.name||"โรงเรียนดาราวิทยาลัย");
    setPrintPreview({html:pdfPage(title,subtitle,buildTeacherDayRows(t),"",sh?.logo||null,printSettings,false,tDiv)});
  };

  // helper: สร้าง pages สำหรับห้องหนึ่ง
  // จำนวนใบ = maxEntries ในคาบที่ซ้อนมากที่สุด
  // ใบที่ i: คาบปกติ→เหมือนกันทุกใบ, คาบซ้อน→entry[i]
  const buildRoomPages=(room)=>{
    const subtitle="ภาคเรียนที่ "+(ay?.semester||"1")+"/"+(ay?.year||"2568")+" "+(sh?.name||"โรงเรียนดาราวิทยาลัย");
    let maxEntries=1;
    DAYS.forEach(day=>PERIODS.forEach(p=>{
      const len=(S.schedule[room.id+"_"+day+"_"+p.id]||[]).length;
      if(len>maxEntries) maxEntries=len;
    }));
    return Array.from({length:maxEntries},(_,sheetIdx)=>({
      title:"ตารางเรียน "+room.name+(maxEntries>1?" (ฉบับที่ "+(sheetIdx+1)+"/"+maxEntries+")":""),
      subtitle:subtitle,
      dayRows:DAYS.map(day=>({day,cells:PERIODS.map(p=>{
        const en=S.schedule[room.id+"_"+day+"_"+p.id]||[];
        if(!en.length){
          // คาบล็อคแผนก — แสดงแทนช่องว่าง
          const custLock=(S.meetings||[]).find(m=>m.type==="custom"&&(m.slots||[]).some(s=>s.day===day&&s.period===p.id));
          if(custLock) return[{sub:"🏫 "+custLock.name,room:"",room2:"",isCustomLock:true}];
          return [];
        }
        const isDouble=en.length>1;
        const e=en[sheetIdx]||en[0];
        const sub=S.subjects.find(s=>s.id===e.subjectId);
        const t2=S.teachers.find(x=>x.id===e.teacherId);
        return[{sub:(sub?.shortName||sub?.name||""),room:(t2?.prefix||"")+(t2?.firstName||""),room2:room.name,double:isDouble}];
      })}))
    }));
  };

  const printRoomPDF=(room)=>{
    const pages=buildRoomPages(room);
    if(!pages.length){st("ยังไม่มีตารางในห้องนี้","error");return;}
    const divId=getDivisionForRoom(room,S);
    setPrintPreview({html:pdfMultiPage(pages,sh?.logo||null,printSettings,true,divId)});
  };


  // PDF แบบใหม่ — 2 ห้อง/หน้า A4 แนวตั้ง, 3 แถว/คาบ, คาบพักแนวตั้ง, auto homeroom

  // สร้าง HTML ตารางเรียนแบบเดียวกับ Excel ต้นแบบ
  const buildRoomTableHTML=(room,opts={})=>{
    const lvl=S.levels.find(l=>l.id===room.levelId);
    const divId=getDivisionForRoom(room,S);
    const pcfg=getPeriodCfg(divId);
    const asmDay=lvl?.assemblyDay||"";
    const h1=room.homeroom1||""; const h2=room.homeroom2||""; const hco=room.homeroomCo||"";
    const yr=ay?.year||"2568";
    const logoImg=sh?.logo?`<img src="${sh.logo}" style="height:40px;vertical-align:middle;margin-right:8px;"/>` :"";

    const DAYS_TH=["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
    const PIDS=[1,2,3,4,5,6,7];

    // หา maxEntries
    let maxEntries=1;
    DAYS_TH.forEach(day=>PIDS.forEach(pid=>{
      const len=(S.schedule[room.id+"_"+day+"_"+pid]||[]).length;
      if(len>maxEntries) maxEntries=len;
    }));
    const totalCopies=maxEntries;

    const copies=[];
    for(let copyIdx=0;copyIdx<totalCopies;copyIdx++){
      const copyLabel=totalCopies>1?` (ฉบับที่ ${copyIdx+1}/${totalCopies})`:"";
      const title=opts.title?opts.title+copyLabel:("ตารางเรียน "+room.name+copyLabel);

      const getCells=(day,pid)=>{
        const key=room.id+"_"+day+"_"+pid;
        const all=S.schedule[key]||[];
        if(!all.length){
          const custLock=(S.meetings||[]).find(m=>m.type==="custom"&&(m.slots||[]).some(s=>s.day===day&&s.period===pid));
          if(custLock) return [{th:"🏫 "+custLock.name,en:"",tch:"",isCustomLock:true}];
          return [];
        }
        const e=all.length>1?(all[copyIdx]||all[0]):all[0];
        if(!e) return [];
        const sub=S.subjects.find(s=>s.id===e.subjectId);
        const t=S.teachers.find(t=>t.id===e.teacherId);
        const cos=(e.coTeacherIds||[]).map(id=>S.teachers.find(x=>x.id===id)).filter(Boolean);
        return[{th:sub?.name||sub?.code||"",en:sub?.shortName||"",tch:[t,...cos].filter(Boolean).map(x=>(x.prefix||"")+x.firstName).join(", ")}];
      };

      // colgroup — ปรับตาม break position
      // p1: break หลัง p6 → col[10] เป็น break, p6 ก่อน break
      // default: break หลัง p5 → col[8] เป็น break
      const colgroup=`<colgroup>
        <col style="width:5%;"><col style="width:2.5%;">
        <col style="width:11.9%;"><col style="width:11.9%;">
        <col style="width:3%;">
        <col style="width:11.9%;"><col style="width:11.9%;">
        <col style="width:3%;">
        <col style="width:11.9%;">
        ${divId==="p1"
          ? `<col style="width:11.9%;"><col style="width:3%;"><col style="width:12.1%;">`
          : `<col style="width:3%;"><col style="width:11.9%;"><col style="width:12.1%;">`
        }
      </colgroup>`;

      const vert=(txt,bg="#fffde7",fw="normal",fs="9pt")=>
        `<div style="writing-mode:vertical-rl;transform:rotate(180deg);white-space:nowrap;font-size:${fs};font-weight:${fw};letter-spacing:1px;text-align:center;">${txt}</div>`;

      const HDR=pcfg.periods.map(p=>({label:`คาบ ${p.id}`,time:p.time}));
      const BRK=[["08.00-","08.30"],["10.10-","10.25"],["12.05-","13.00"]];
      const BRK4=divId==="p1"?["14.40-","14.50"]:["13.50-","14.00"];
      const BRKall=[...BRK,BRK4];
      const vertBRK=(parts,fs="9pt")=>
        `<div style="writing-mode:vertical-rl;transform:rotate(180deg);font-size:${fs};font-weight:600;letter-spacing:1px;text-align:center;display:flex;flex-direction:column;align-items:center;">${parts.map(p=>'<span style="white-space:nowrap;">'+p+'</span>').join("")}</div>`;

      const brkTh=(i)=>`<th rowspan="2" style="border:1px solid #666;background:#fffde7;padding:0;height:38px;vertical-align:middle;text-align:center;">${vertBRK(BRKall[i])}</th>`;

      // Header แถว 1 — ปรับตาม divisionId
      const h1row=`<tr style="background:#f0f0f0;height:22px;max-height:22px;">
        <th rowspan="2" style="border:1px solid #666;padding:0;position:relative;vertical-align:middle;font-size:7pt;height:38px;">
          <div style="position:absolute;top:0;left:0;width:100%;height:100%;">
            <svg style="position:absolute;top:0;left:0;width:100%;height:100%;" preserveAspectRatio="none">
              <line x1="0" y1="0" x2="100%" y2="100%" stroke="#888" stroke-width="0.8"/>
            </svg>
            <div style="position:absolute;top:2px;right:2px;font-size:5.5pt;color:#555;">เวลา</div>
            <div style="position:absolute;bottom:2px;left:2px;font-size:5.5pt;color:#555;">วัน</div>
          </div>
        </th>
        ${brkTh(0)}
        <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[0].label}</th>
        <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[1].label}</th>
        ${brkTh(1)}
        <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[2].label}</th>
        <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[3].label}</th>
        ${brkTh(2)}
        <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[4].label}</th>
        ${divId==="p1"
          ? `<th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[5].label}</th>
             ${brkTh(3)}
             <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[6].label}</th>`
          : `${brkTh(3)}
             <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[5].label}</th>
             <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[6].label}</th>`
        }
      </tr>
      <tr style="background:#f0f0f0;height:16px;max-height:16px;">
        ${[0,1,2,3,4,5,6].map(i=>`<td style="border:1px solid #888;font-size:6.5pt;text-align:center;padding:1px;height:16px;">${HDR[i].time}</td>`).join("")}
      </tr>`;

      const DAYS_TH2=["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
      let body="";
      DAYS_TH2.forEach((day,di)=>{
        const isAsm=asmDay===day;
        const hmTxt=isAsm?"หอประชุม<br>Assembly":"Homeroom";
        const hmBg=isAsm?"#e8f5e9":"#fafff7";
        const D=[1,2,3,4,5,6,7].map(pid=>getCells(day,pid));
        const isMulti=[1,2,3,4,5,6,7].map(pid=>(S.schedule[room.id+"_"+day+"_"+pid]||[]).length>1);
        const MBG="#eeeeee";

        const cell=(arr,type,multi=false)=>{
          const isLock=arr[0]?.isCustomLock;
          const v=arr.map(c=>c[type]).filter(Boolean).join("<br>");
          const s=isLock?"font-size:8pt;font-weight:bold;color:#E65100;"
            :type==="th"?"font-size:8.5pt;font-weight:bold;"
            :type==="en"?"font-size:7.5pt;color:#444;"
            :"font-size:7.5pt;color:#1a237e;";
          const bg=isLock?"background:#FFF3E0;":multi?`background:${MBG};`:"";
          if(isLock&&type!=="th") return `<td style="border:1px solid #ddd;border-top:none;border-bottom:none;${bg}"></td>`;
          return`<td style="border:1px solid #ddd;border-top:none;border-bottom:none;text-align:center;vertical-align:middle;padding:2px;${bg}${s}">${v}</td>`;
        };
        const cellTop=(arr,type,multi=false)=>cell(arr,type,multi).replace("border-top:none;","border-top:1px solid #888;");
        const cellBot=(arr,type,multi=false)=>cell(arr,type,multi).replace("border-bottom:none;","border-bottom:1px solid #888;");

        const BKcell=(rows,vtext,bg="#fffde7")=>
          `<td rowspan="${rows}" style="border:1px solid #888;background:${bg};padding:0;vertical-align:middle;text-align:center;min-width:26px;">${vert(vtext,bg,"600","9pt")}</td>`;

        const bk=di===0;
        const TOTAL_ROWS=DAYS_TH2.length*3;
        const brkLbl4=divId==="p1"?"พัก 14.40-14.50":"พัก 13.50-14.00";

        body+=`
          <tr style="height:20px;max-height:20px;">
            <td rowspan="3" style="border:1px solid #888;text-align:center;font-weight:bold;font-size:10pt;vertical-align:middle;background:#f5f5f5;padding:2px;">${day==="พฤหัสบดี"?"พฤหัส":day}</td>
            <td rowspan="3" style="border:1px solid #888;background:${hmBg};padding:0;vertical-align:middle;text-align:center;min-width:26px;">${vert(hmTxt.replace('<br>','/').replace('<br/>','/'),hmBg,"600","9pt")}</td>
            ${cellTop(D[0],"th",isMulti[0])}${cellTop(D[1],"th",isMulti[1])}
            ${bk?BKcell(TOTAL_ROWS,"พัก 10.10-10.25"):""}
            ${cellTop(D[2],"th",isMulti[2])}${cellTop(D[3],"th",isMulti[3])}
            ${bk?BKcell(TOTAL_ROWS,"พัก 12.05-13.00"):""}
            ${cellTop(D[4],"th",isMulti[4])}
            ${divId==="p1"
              ? `${cellTop(D[5],"th",isMulti[5])}${bk?BKcell(TOTAL_ROWS,brkLbl4):""}${cellTop(D[6],"th",isMulti[6])}`
              : `${bk?BKcell(TOTAL_ROWS,brkLbl4):""}${cellTop(D[5],"th",isMulti[5])}${cellTop(D[6],"th",isMulti[6])}`
            }
          </tr>
          <tr style="height:17px;max-height:17px;">
            ${cell(D[0],"en",isMulti[0])}${cell(D[1],"en",isMulti[1])}
            ${cell(D[2],"en",isMulti[2])}${cell(D[3],"en",isMulti[3])}
            ${cell(D[4],"en",isMulti[4])}
            ${divId==="p1"
              ? `${cell(D[5],"en",isMulti[5])}${cell(D[6],"en",isMulti[6])}`
              : `${cell(D[5],"en",isMulti[5])}${cell(D[6],"en",isMulti[6])}`
            }
          </tr>
          <tr style="height:17px;max-height:17px;">
            ${cellBot(D[0],"tch",isMulti[0])}${cellBot(D[1],"tch",isMulti[1])}
            ${cellBot(D[2],"tch",isMulti[2])}${cellBot(D[3],"tch",isMulti[3])}
            ${cellBot(D[4],"tch",isMulti[4])}${cellBot(D[5],"tch",isMulti[5])}${cellBot(D[6],"tch",isMulti[6])}
          </tr>`;
    });

      const footer=(h1||h2||hco)?`
        <div style="margin-top:6px;font-size:10pt;font-family:'TH SarabunNew','Sarabun',sans-serif;text-align:right;line-height:2;">
          ${h1||h2?`<div><b>ครูประจำชั้นหลัก</b>&emsp;&emsp;${h1}${h2?"&emsp;&emsp;&emsp;&emsp;&emsp;"+h2:""}</div>`:""}
          ${hco?`<div><b>ครูประจำชั้นร่วม</b>&emsp;&emsp;${hco}</div>`:""}
        </div>`:"";

      copies.push(`
        <div style="text-align:center;margin-bottom:5px;font-family:'TH SarabunNew','Sarabun',sans-serif;">
          ${logoImg}<b style="font-size:13pt;">${title}&emsp;&emsp;ปีการศึกษา ${yr}</b>
        </div>
        <table style="width:100%;border-collapse:collapse;table-layout:fixed;overflow:hidden;">
          ${colgroup}
          <thead>${h1row}</thead>
          <tbody>${body}</tbody>
        </table>
        ${footer}`);
    } // end copyIdx loop

    return copies; // return array of HTML strings
  };

  const printRoomPDFNew=(rooms,opts={})=>{
    const roomList=Array.isArray(rooms)?rooms:[rooms];
    if(!roomList.length){st("ไม่มีห้องที่เลือก","error");return;}
    const w=window.open('','_blank');
    if(!w){st("Browser บล็อก popup","error");return;}

    const layout=opts.layout||"2portrait";

    // สร้าง pages: แต่ละ element คือ array ของ room copies
    // flatten: [room1copy1, room1copy2, room2copy1, ...]
    const allCopies=roomList.flatMap(rm=>buildRoomTableHTML(rm,{}));

    let pagesHTML="";
    if(layout==="1landscape"){
      pagesHTML=allCopies.map((html,pi)=>`
        <div style="page-break-after:${pi<allCopies.length-1?"always":"avoid"};padding:8mm 10mm;box-sizing:border-box;">
          ${html}
        </div>`).join("");
    } else {
      // 2 ต่อหน้า
      const pages=[];
      for(let i=0;i<allCopies.length;i+=2) pages.push(allCopies.slice(i,i+2));
      pagesHTML=pages.map((pair,pi)=>`
        <div style="page-break-after:${pi<pages.length-1?"always":"avoid"};padding:6mm 8mm;box-sizing:border-box;">
          ${pair.join(`<div style="border-top:1px dashed #ccc;margin:6px 0;"></div>`)}
        </div>`).join("");
    }

    const pageSize=layout==="1landscape"?"A4 landscape":"A4 portrait";
    const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/>
    <style>
      @page{size:${pageSize};margin:0}
      body{font-family:'TH SarabunNew','Sarabun','Arial',sans-serif;margin:0;padding:0;}
      td,th{word-wrap:break-word;overflow:hidden;line-height:1.15;max-height:20px;}
      @media print{body{margin:0;}}
    </style></head><body>${pagesHTML}</body></html>`;

    w.document.write(html);
    w.document.close();
    setTimeout(()=>w.print(),700);
  };
  // helper: build dayRows สำหรับครู 1 คน พร้อม lock cells
  const buildTeacherDayRows=(t)=>{    const sortParts=(parts)=>parts.sort((a,b)=>{      const numA=parseInt((a.room.match(/(\d+)$/)||[0,9999])[1]);      const numB=parseInt((b.room.match(/(\d+)$/)||[0,9999])[1]);      if(numA!==numB) return numA-numB;      return a.room.localeCompare(b.room,"th");    });    const getLock=(day,pid)=>{      for(const pl of (t.personalLocks||[])){        if(pl.day===day&&(pl.periods||[]).includes(pid))          return [{sub:"🔒 "+(pl.reason||"ส่วนตัว"),room:"",room2:"",isLock:true,lockColor:"#FEF3C7",lockTextColor:"#92400E"}];      }      for(const rid of (t.specialRoles||[])){        const role=SROLES.find(r=>r.id===rid);        const bl=(role?.blocked||[]).find(b=>b.day===day&&(b.periods||[]).includes(pid));        if(bl) return [{sub:"📋 "+role.name.replace("ฝ่ายวิชาการ","วิชาการ").replace("ฝ่ายพัฒนาวินัย","พัฒนาวินัย"),room:"",room2:"",isLock:true,lockColor:"#EDE9FE",lockTextColor:"#5B21B6"}];      }      for(const m of (S.meetings||[])){        if(m.type&&m.type!=="dept") continue;        if(m.isAssembly||m.isHomeroom) continue;        const isMyDept=!m.departmentId||m.departmentId===t.departmentId||m.departmentId==="all";        if(isMyDept&&m.day===day&&(m.periods||[]).includes(pid))          return [{sub:"🔒 LOCK",room:"",room2:"",isLock:true,lockColor:"#DBEAFE",lockTextColor:"#1D4ED8"}];      }      for(const m of (S.meetings||[])){        if(m.type!=="custom") continue;        if((m.slots||[]).some(s=>s.day===day&&s.period===pid))          return [{sub:"🏫 "+(m.name||"Lock"),room:"",room2:"",isLock:true,lockColor:"#F3F4F6",lockTextColor:"#374151"}];      }      for(const m of (S.meetings||[])){        if(m.teacherId===t.id&&(m.isAssembly||m.isHomeroom)&&m.day===day&&(m.periods||[]).includes(pid))          return [{sub:m.isAssembly?"🎤 หอประชุม":"🏠 Homeroom",room:"",room2:"",isLock:true,lockColor:"#FEF9C3",lockTextColor:"#92400E"}];      }      return null;    };    return DAYS.map(day=>({day,cells:PERIODS.map(p=>{      let parts=[];      Object.entries(S.schedule).forEach(([k,en])=>{        if(!k.endsWith("_"+day+"_"+p.id))return;        en?.forEach(e=>{          const pCoIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);          if(e.teacherId===t.id||pCoIds.includes(t.id)){            const sub=S.subjects.find(s=>s.id===e.subjectId);            const rid=k.split("_")[0];            const rm=S.rooms.find(r=>r.id===rid);            parts.push({sub:(sub?.shortName||sub?.name||""),room:rm?.name||"",room2:""});          }        });      });      if(parts.length) return sortParts(parts);      return getLock(day,p.id)||[];    })}));  };
  const printAllTeachersPDF=()=>{
    const teachers=S.teachers.filter(t=>t.totalPeriods>0);
    if(!teachers.length){st("ไม่มีครูที่กำหนดคาบ","error");return}
    const pages=teachers.map(t=>({
      title:"ตารางสอน "+(t.prefix||"")+(t.firstName||"")+" "+(t.lastName||""),
      subtitle:"ภาคเรียนที่ "+(ay?.semester||"1")+"/"+(ay?.year||"2568")+" "+(sh?.name||"โรงเรียนดาราวิทยาลัย"),
      dayRows:buildTeacherDayRows(t)
    }));
    const w=window.open('','_blank');
    w.document.write(pdfMultiPage(pages,sh?.logo||null,printSettings,false));
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
    if(!sorted.length){st("ยังไม่มีตารางในระบบ","error");return}
    // build HTML per-room เพื่อใช้เวลาคาบที่ถูกต้องของแต่ละห้อง
    // แล้ว concat body เข้ากัน (ใช้ pdfPage ทีละห้อง แล้ว open window เดียว)
    const w=window.open('','_blank');
    if(!w){st("Browser บล็อก popup","error");return;}
    // สร้าง HTML รวม: loop แต่ละห้อง สร้าง pdfPage แต่เอาแค่ body
    const allHtml=sorted.flatMap(room=>{
      const pages=buildRoomPages(room);
      if(!pages.length) return [];
      const divId=getDivisionForRoom(room,S);
      return pages.map((pg,pi)=>
        pdfMultiPage([pg],sh?.logo||null,printSettings,true,divId)
      );
    });
    if(!allHtml.length){st("ยังไม่มีตารางในระบบ","error");w.close();return;}
    // เอาแค่ HTML แรกเป็น wrapper แล้วต่อ body block ของที่เหลือ
    const combined=sorted.flatMap(room=>{
      const pages=buildRoomPages(room);
      const divId=getDivisionForRoom(room,S);
      return pages;
    });
    w.document.write(pdfMultiPage(combined,sh?.logo||null,printSettings,true,"m2"));
    w.document.close();setTimeout(()=>w.print(),600);
    st("กำลังพิมพ์ตารางเรียนทุกห้อง");
  };

  // PDF: ตารางสอนรวมกลุ่มสาระ (landscape)
  const printMasterByDept=()=>{
    const w=window.open('','_blank');
    w.document.write(buildMasterTableHTML(S,ay,sh,null));
    w.document.close();setTimeout(()=>w.print(),600);
    st("กำลังพิมพ์ตารางรวมกลุ่มสาระ");
  };

  // PDF: ตารางสอนรวมระดับชั้น
  const [masterLevel,setMasterLevel]=useState("");
  const printMasterByLevel=()=>{
    if(!masterLevel){st("เลือกระดับชั้นก่อน","error");return;}
    const w=window.open('','_blank');
    w.document.write(buildLevelTableHTML(S,ay,sh,masterLevel));
    w.document.close();setTimeout(()=>w.print(),600);
    const lvName=S.levels.find(l=>l.id===masterLevel)?.name||"";
    st("กำลังพิมพ์ตารางรวมห้องระดับ "+lvName);
  };

  return <div style={{animation:"fadeIn 0.3s",display:"flex",flexDirection:"column",gap:20}}>

    {/* ── PRINT DESIGNER ── */}
    <PrintDesignerModal open={showPrintDesigner} onClose={()=>setShowPrintDesigner(false)} S={S} ay={ay} sh={sh} />

    {/* ── REPORT TABS ── */}
    <div data-ui-surface="true" style={{background:"#fff",borderRadius:16,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.07)"}}>
      <div style={{display:"flex",gap:8,marginBottom:20,flexWrap:"wrap",alignItems:"center"}}>
        {[["print","พิมพ์ตาราง"],["excel","ส่งออก Excel"],["hours","ชั่วโมงสอน"],["conflicts","ตรวจคาบชน"],["backup","สำรอง / คืนค่า"]].map(([v,l])=>(
          <button data-ui-control="true" key={v} onClick={()=>setReportTab(v)} style={{padding:"8px 18px",borderRadius:10,fontWeight:700,fontSize:13,border:`2px solid ${reportTab===v?"#B91C1C":"#E5E7EB"}`,background:reportTab===v?"#B91C1C":"#fff",color:reportTab===v?"#fff":"#374151",cursor:"pointer"}}>{l}</button>
        ))}
        <button data-ui-control="true" className="secondary" onClick={()=>setShowPrintDesigner(true)}>ออกแบบตารางพิมพ์ · Print Studio</button><FileActions label="ตั้งค่ารายงาน" actions={[['ออกแบบรูปแบบพิมพ์',()=>setShowPrintDesigner(true)],['ตั้งค่ากระดาษและตัวอักษร',()=>setShowPrintSettings(true)]]}/>
      </div>
      <PrintSettingsPanel open={showPrintSettings} onClose={()=>setShowPrintSettings(false)} onApply={(s)=>{setPrintSettings(s);st("บันทึกการตั้งค่าแล้ว ✓");}} />
      <PrintPreviewModal data={printPreview} onClose={()=>setPrintPreview(null)} ps={printSettings}/>

      {reportTab==="hours" && <>
        {/* ── ปุ่ม Export Excel รายงานครู ── */}
        <div style={{marginBottom:16}}>
          <button data-ui-control="true" onClick={()=>{
            // สร้างแถว: รหัสครู | ชื่อ-สกุล | รหัสวิชา | ชื่อวิชา | ห้องเรียน | วัน | คาบ
            const DAYS_TH=["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
            const rows=[];
            S.teachers.forEach(t=>{
              const taught=new Set();
              DAYS_TH.forEach(day=>{
                PERIODS.forEach(p=>{
                  Object.entries(S.schedule||{}).forEach(([k,ens])=>{
                    const pts=k.split("_");
                    if(pts[pts.length-2]===day&&parseInt(pts[pts.length-1])===p.id){
                      ens?.forEach(e=>{
                        const allT=[e.teacherId,...(e.coTeacherIds||[])].filter(Boolean);
                        if(!allT.includes(t.id)) return;
                        const sub=S.subjects.find(s=>s.id===e.subjectId);
                        const roomId=pts[0];
                        const room=S.rooms.find(r=>r.id===roomId);
                        const key=`${e.subjectId}_${roomId}`;
                        if(!taught.has(key)){
                          taught.add(key);
                          rows.push([
                            t.teacherCode||"",
                            (t.prefix||"")+(t.firstName||"")+" "+(t.lastName||""),
                            S.depts.find(d=>d.id===t.departmentId)?.name||"",
                            sub?.code||"",
                            sub?.name||"",
                            room?.name||"",
                          ]);
                        }
                      });
                    }
                  });
                });
              });
            });
            // เรียงตามรหัสครู
            rows.sort((a,b)=>(a[0]||"").localeCompare(b[0]||"","th"));
            exportExcel(
              ["รหัสครู","ชื่อ-สกุลครู","กลุ่มสาระ","รหัสวิชา","ชื่อวิชา","ห้องเรียน"],
              rows,
              "รายงาน_ครู_วิชา_ห้องเรียน.xlsx",
              "ครู-วิชา-ห้อง"
            );
            st("Export Excel สำเร็จ ✓");
          }} style={{
            padding:"10px 20px",borderRadius:10,border:"none",fontFamily:"inherit",
            fontSize:13,fontWeight:600,cursor:"pointer",
            background:"linear-gradient(135deg,#059669,#047857)",
            color:"#fff",boxShadow:"0 4px 12px rgba(5,150,105,0.3)",
            display:"inline-flex",alignItems:"center",gap:8
          }}>
            📊 Export Excel — ครู / วิชา / ห้องเรียน
          </button>
        </div>
        <TeacherHoursSummary S={S} />
      </>}
      {reportTab==="conflicts" && <ConflictPanel S={S} />}
      <div style={{display:["print","excel","backup"].includes(reportTab)?"block":"none"}}>

      {['print','excel','backup'].includes(reportTab)&&<ReportLauncher key={reportTab} mode={reportTab} S={S} actions={{
        teacherPDFAll:printAllTeachersPDF,roomPDFAll:printAllRoomsPDF,deptPDF:printMasterByDept,
        teacherPDF:id=>{const t=S.teachers.find(x=>x.id===id);if(t)printTeacherPDF(t)},
        roomPDF:id=>{const r=S.rooms.find(x=>x.id===id);if(r)printRoomPDF(r)},
        levelPDF:id=>{const w=window.open('','_blank');if(!w){st('อนุญาตหน้าต่างพิมพ์ในเบราว์เซอร์ก่อน','error');return}w.document.write(buildLevelTableHTML(S,ay,sh,id));w.document.close();setTimeout(()=>w.print(),600)},
        teacherExcelAll:exportAllTeachers,roomExcelAll:exportAllRooms,status:exportStatus,
        teacherExcel:id=>{const t=S.teachers.find(x=>x.id===id);if(t)exportTeacherXL(t)},roomExcel:id=>{const r=S.rooms.find(x=>x.id===id);if(r)exportRoomXL(r)},
        teacherTwo:()=>{setSelectedTeachersPDF([]);setTeacherSearchQ('');setShowNewTeacherPDF(true)},roomLayout:()=>{setNewRoomPDFOpts({selectedRooms:[],layout:'2portrait'});setShowNewRoomPDF(true)},roomExcelChoose:()=>{setExcelSelectedRooms([]);setShowExcelModal(true)},
        backup:exportScheduleJSON,restore:()=>fileRefSched.current?.click()
      }}/>}
      {/* Modal: พิมพ์แบบใหม่ */}
      {showNewRoomPDF&&(
        <div style={{position:"fixed",inset:0,zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.5)"}}>
          <div data-ui-surface="true" style={{background:"#fff",borderRadius:16,boxShadow:"0 20px 60px rgba(0,0,0,0.3)",width:"min(560px,94%)",maxHeight:"90vh",overflowY:"auto",padding:24,fontFamily:"inherit"}}>
            <div style={{fontSize:16,fontWeight:800,marginBottom:4}}>🆕 พิมพ์ตารางเรียนแบบใหม่</div>
            <div style={{fontSize:11,color:"#6B7280",marginBottom:16}}>auto-อ่านครูประจำชั้นและวันหอประชุมจากระบบ</div>

            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              {/* Layout selector */}
              <div>
                <label style={LS}>รูปแบบการพิมพ์</label>
                <div style={{display:"flex",gap:10}}>
                  {[
                    {val:"2portrait",label:"2 ห้อง / หน้า",sub:"A4 แนวตั้ง",icon:"📄"},
                    {val:"1landscape",label:"1 ห้อง / หน้า",sub:"A4 แนวนอน (เต็มหน้า)",icon:"🖥️"},
                  ].map(opt=>{
                    const sel=newRoomPDFOpts.layout===opt.val;
                    return<button data-ui-control="true" key={opt.val} onClick={()=>setNewRoomPDFOpts(p=>({...p,layout:opt.val}))}
                      style={{flex:1,padding:"10px 8px",borderRadius:12,border:`2px solid ${sel?"#7C3AED":"#E5E7EB"}`,background:sel?"#F5F3FF":"#fff",cursor:"pointer",textAlign:"center"}}>
                      <div style={{fontSize:18}}>{opt.icon}</div>
                      <div style={{fontSize:12,fontWeight:700,color:sel?"#7C3AED":"#374151"}}>{opt.label}</div>
                      <div style={{fontSize:10,color:"#6B7280"}}>{opt.sub}</div>
                    </button>;
                  })}
                </div>
              </div>
              {/* เลือกห้อง */}
              <div>
                <label style={LS}>เลือกห้องที่ต้องการพิมพ์ (กดหลายห้องได้)</label>
                <div style={{display:"flex",gap:6,flexWrap:"wrap",maxHeight:200,overflowY:"auto",padding:4,border:"1px solid #E5E7EB",borderRadius:8}}>
                  {[...S.rooms].sort((a,b)=>{
                    const la=S.levels.find(l=>l.id===a.levelId)?.name||"";
                    const lb=S.levels.find(l=>l.id===b.levelId)?.name||"";
                    if(la!==lb)return la.localeCompare(lb,"th");
                    return a.name.localeCompare(b.name,"th");
                  }).map(r=>{
                    const sel=(newRoomPDFOpts.selectedRooms||[]).includes(r.id);
                    return<button data-ui-control="true" key={r.id}
                      onClick={()=>setNewRoomPDFOpts(p=>({...p,selectedRooms:sel?p.selectedRooms.filter(id=>id!==r.id):[...p.selectedRooms,r.id]}))}
                      style={{padding:"4px 12px",borderRadius:20,border:`2px solid ${sel?"#7C3AED":"#E5E7EB"}`,background:sel?"#7C3AED":"#fff",color:sel?"#fff":"#374151",fontSize:12,fontWeight:sel?700:400,cursor:"pointer"}}>
                      {r.name}
                    </button>;
                  })}
                </div>
                <div style={{display:"flex",gap:6,marginTop:6}}>
                  <button data-ui-control="true" onClick={()=>setNewRoomPDFOpts(p=>({...p,selectedRooms:S.rooms.map(r=>r.id)}))} style={{fontSize:11,color:"#7C3AED",background:"none",border:"1px solid #E5E7EB",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>เลือกทั้งหมด</button>
                  <button data-ui-control="true" onClick={()=>setNewRoomPDFOpts(p=>({...p,selectedRooms:[]}))} style={{fontSize:11,color:"#6B7280",background:"none",border:"1px solid #E5E7EB",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>ล้าง</button>
                  <span style={{fontSize:11,color:"#6B7280",marginLeft:4,alignSelf:"center"}}>เลือกแล้ว {newRoomPDFOpts.selectedRooms?.length||0} ห้อง → {Math.ceil((newRoomPDFOpts.selectedRooms?.length||0)/2)} หน้า</span>
                </div>
              </div>

              {/* ครูประจำชั้น — แสดงจากระบบ */}
              <div style={{padding:"10px 12px",background:"#F0F9FF",borderRadius:8,fontSize:11,color:"#0369A1"}}>
                💡 ครูประจำชั้นจะถูกอ่านจากข้อมูลในเมนู <b>ครูประจำชั้น</b> โดยอัตโนมัติ
                <div style={{marginTop:4,color:"#0284C7"}}>
                  ตัวอย่าง: {S.rooms.filter(r=>(newRoomPDFOpts.selectedRooms||[]).includes(r.id)&&r.homeroom1).slice(0,2).map(r=>`${r.name}: ${r.homeroom1}`).join(" · ")||"(เลือกห้องก่อน)"}
                </div>
              </div>
            </div>

            <div style={{display:"flex",gap:10,marginTop:20}}>
              <button data-ui-control="true" onClick={()=>setShowNewRoomPDF(false)} style={{...BO(),flex:1}}>ยกเลิก</button>
              <button data-ui-control="true"
                disabled={!newRoomPDFOpts.selectedRooms?.length}
                onClick={()=>{
                  const rooms=S.rooms.filter(r=>(newRoomPDFOpts.selectedRooms||[]).includes(r.id));
                  const sorted=[...rooms].sort((a,b)=>{
                    const la=S.levels.find(l=>l.id===a.levelId)?.name||"";
                    const lb=S.levels.find(l=>l.id===b.levelId)?.name||"";
                    if(la!==lb)return la.localeCompare(lb,"th");
                    return a.name.localeCompare(b.name,"th");
                  });
                  const w=window.open('','_blank');
                  if(!w){st("Browser บล็อก popup","error");return;}
                  const saved=w.setTimeout;w.setTimeout=()=>{};
                  printRoomPDFNew(sorted,{layout:newRoomPDFOpts.layout});
                  setTimeout(()=>{w.setTimeout=saved;},100);
                }}
                style={{...BO("#7C3AED"),flex:1,opacity:newRoomPDFOpts.selectedRooms?.length?1:0.4,fontSize:12}}>
                👁️ ดูตัวอย่าง
              </button>
              <button data-ui-control="true"
                disabled={!newRoomPDFOpts.selectedRooms?.length}
                onClick={()=>{
                  const rooms=S.rooms.filter(r=>(newRoomPDFOpts.selectedRooms||[]).includes(r.id));
                  const sorted=[...rooms].sort((a,b)=>{
                    const la=S.levels.find(l=>l.id===a.levelId)?.name||"";
                    const lb=S.levels.find(l=>l.id===b.levelId)?.name||"";
                    if(la!==lb)return la.localeCompare(lb,"th");
                    return a.name.localeCompare(b.name,"th");
                  });
                  printRoomPDFNew(sorted,{layout:newRoomPDFOpts.layout});
                  setShowNewRoomPDF(false);
                }}
                style={{...BS("#7C3AED"),flex:2,opacity:newRoomPDFOpts.selectedRooms?.length?1:0.4}}>
                🖨️ พิมพ์ ({newRoomPDFOpts.layout==="1landscape"
                  ?`${newRoomPDFOpts.selectedRooms?.length||0}+ หน้า`
                  :`${newRoomPDFOpts.selectedRooms?.length||0}+ หน้า`})
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Export Excel ตารางห้อง */}
      {showExcelModal&&(
        <div style={{position:"fixed",inset:0,zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.5)"}}>
          <div data-ui-surface="true" style={{background:"#fff",borderRadius:16,boxShadow:"0 20px 60px rgba(0,0,0,0.3)",width:"min(520px,94%)",maxHeight:"90vh",overflowY:"auto",padding:24,fontFamily:"inherit"}}>
            <div style={{fontSize:16,fontWeight:800,marginBottom:4}}>📊 Export Excel ตารางห้องเรียน</div>
            <div style={{fontSize:11,color:"#6B7280",marginBottom:16}}>แต่ละห้อง = 1 sheet · format: วัน/รหัสวิชา/เวลา/รหัสครู</div>

            <div>
              <label style={LS}>เลือกห้องที่ต้องการ export</label>
              <div style={{display:"flex",gap:6,flexWrap:"wrap",maxHeight:200,overflowY:"auto",padding:6,border:"1px solid #E5E7EB",borderRadius:8}}>
                {[...S.rooms].sort((a,b)=>{
                  const la=S.levels.find(l=>l.id===a.levelId)?.name||"";
                  const lb=S.levels.find(l=>l.id===b.levelId)?.name||"";
                  if(la!==lb)return la.localeCompare(lb,"th");
                  return a.name.localeCompare(b.name,"th");
                }).map(r=>{
                  const sel=excelSelectedRooms.includes(r.id);
                  return<button data-ui-control="true" key={r.id}
                    onClick={()=>setExcelSelectedRooms(p=>sel?p.filter(id=>id!==r.id):[...p,r.id])}
                    style={{padding:"4px 12px",borderRadius:20,border:`2px solid ${sel?"#059669":"#E5E7EB"}`,background:sel?"#059669":"#fff",color:sel?"#fff":"#374151",fontSize:12,fontWeight:sel?700:400,cursor:"pointer"}}>
                    {r.name}
                  </button>;
                })}
              </div>
              <div style={{display:"flex",gap:6,marginTop:6,alignItems:"center"}}>
                <button data-ui-control="true" onClick={()=>setExcelSelectedRooms(S.rooms.map(r=>r.id))} style={{fontSize:11,color:"#059669",background:"none",border:"1px solid #D1FAE5",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>เลือกทั้งหมด</button>
                <button data-ui-control="true" onClick={()=>setExcelSelectedRooms([])} style={{fontSize:11,color:"#6B7280",background:"none",border:"1px solid #E5E7EB",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>ล้าง</button>
                <span style={{fontSize:11,color:"#6B7280"}}>เลือกแล้ว {excelSelectedRooms.length} ห้อง → {excelSelectedRooms.length} sheets</span>
              </div>
            </div>

            <div style={{display:"flex",gap:10,marginTop:20}}>
              <button data-ui-control="true" onClick={()=>setShowExcelModal(false)} style={{...BO(),flex:1}}>ยกเลิก</button>
              <button data-ui-control="true"
                disabled={!excelSelectedRooms.length}
                onClick={async()=>{
                  const rooms=S.rooms.filter(r=>excelSelectedRooms.includes(r.id));
                  setShowExcelModal(false);
                  await exportRoomScheduleXLSX(rooms);
                }}
                style={{...BS("#059669"),flex:2,opacity:excelSelectedRooms.length?1:0.4}}>
                <Icon name="download" size={14}/>📊 Export ({excelSelectedRooms.length} ห้อง)
              </button>
            </div>
          </div>
        </div>
      )}
      {showNewTeacherPDF&&(
        <div style={{position:"fixed",inset:0,zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.5)"}}>
          <div data-ui-surface="true" style={{background:"#fff",borderRadius:16,boxShadow:"0 20px 60px rgba(0,0,0,0.3)",width:"min(560px,94%)",maxHeight:"90vh",overflowY:"auto",padding:24,fontFamily:"inherit"}}>
            <div style={{fontSize:16,fontWeight:800,marginBottom:4}}>🆕 พิมพ์ตารางสอนครูแบบใหม่</div>
            <div style={{fontSize:11,color:"#6B7280",marginBottom:16}}>A4 แนวตั้ง — 2 คนต่อหน้า · แสดงวิชา+ห้อง+ชื่ออังกฤษ</div>
            <div>
              <label style={LS}>เลือกครู (กดหลายคนได้)</label>
              <input data-ui-control="true"
                style={{...IS,marginBottom:8,fontSize:12}}
                placeholder="🔍 ค้นหาชื่อครู..."
                value={teacherSearchQ||""}
                onChange={e=>setTeacherSearchQ(e.target.value)}
              />
              <div style={{display:"flex",gap:6,flexWrap:"wrap",maxHeight:220,overflowY:"auto",padding:6,border:"1px solid #E5E7EB",borderRadius:8}}>
                {[...S.teachers].sort((a,b)=>{
                  const da=S.depts.find(d=>d.id===a.departmentId)?.name||"";
                  const db=S.depts.find(d=>d.id===b.departmentId)?.name||"";
                  if(da!==db)return da.localeCompare(db,"th");
                  return a.firstName.localeCompare(b.firstName,"th");
                }).filter(t=>{
                  const q=(teacherSearchQ||"").trim().toLowerCase();
                  if(!q) return true;
                  const full=(t.prefix+t.firstName+" "+t.lastName).toLowerCase();
                  const dept=(S.depts.find(d=>d.id===t.departmentId)?.name||"").toLowerCase();
                  return full.includes(q)||dept.includes(q);
                }).map(t=>{
                  const sel=selectedTeachersPDF.includes(t.id);
                  const dept=S.depts.find(d=>d.id===t.departmentId)?.name||"";
                  return<button data-ui-control="true" key={t.id}
                    onClick={()=>setSelectedTeachersPDF(p=>sel?p.filter(id=>id!==t.id):[...p,t.id])}
                    style={{padding:"4px 12px",borderRadius:20,border:`2px solid ${sel?"#7C3AED":"#E5E7EB"}`,background:sel?"#7C3AED":"#fff",color:sel?"#fff":"#374151",fontSize:12,fontWeight:sel?700:400,cursor:"pointer"}}>
                    {t.prefix}{t.firstName} {t.lastName}
                    {dept&&<span style={{fontSize:10,opacity:0.7,marginLeft:4}}>[{dept}]</span>}
                  </button>;
                })}
              </div>
              <div style={{display:"flex",gap:6,marginTop:6}}>
                <button data-ui-control="true" onClick={()=>setSelectedTeachersPDF(S.teachers.map(t=>t.id))} style={{fontSize:11,color:"#7C3AED",background:"none",border:"1px solid #E5E7EB",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>เลือกทั้งหมด</button>
                <button data-ui-control="true" onClick={()=>setSelectedTeachersPDF([])} style={{fontSize:11,color:"#6B7280",background:"none",border:"1px solid #E5E7EB",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>ล้าง</button>
                <span style={{fontSize:11,color:"#6B7280",alignSelf:"center"}}>เลือก {selectedTeachersPDF.length} คน → {Math.ceil(selectedTeachersPDF.length/2)} หน้า</span>
              </div>
            </div>
            <div style={{display:"flex",gap:10,marginTop:20,flexWrap:"wrap"}}>
              <button data-ui-control="true" onClick={()=>setShowNewTeacherPDF(false)} style={{...BO(),flex:1,minWidth:80}}>ยกเลิก</button>
              <button data-ui-control="true"
                onClick={()=>{
                  if(!selectedTeachersPDF.length){st("เลือกครูก่อน","error");return;}
                  printTeacherPDFNew(S.teachers.filter(t=>selectedTeachersPDF.includes(t.id)));
                }}
                style={{...BO("#7C3AED"),flex:1,minWidth:80,opacity:selectedTeachersPDF.length?1:0.4,fontSize:12}}>
                👁️ ดูตัวอย่าง
              </button>
              <button data-ui-control="true"
                onClick={()=>{
                  if(!selectedTeachersPDF.length){st("เลือกครูก่อน","error");return;}
                  const list2=S.teachers.filter(t=>selectedTeachersPDF.includes(t.id));
                  setShowNewTeacherPDF(false);
                  setTimeout(()=>setPrintPreview({html:buildF2Html(list2,S,ay,sh,printSettings)}),50);
                }}
                style={{...BS("#7C3AED"),flex:2,minWidth:120,opacity:selectedTeachersPDF.length?1:0.4}}>
                🖨️ แบบ 2 — 2คน/หน้า ({Math.ceil(selectedTeachersPDF.length/2)} หน้า)
              </button>
              <button data-ui-control="true"
                onClick={()=>{
                  if(!selectedTeachersPDF.length){st("เลือกครูก่อน","error");return;}
                  const list3=S.teachers.filter(t=>selectedTeachersPDF.includes(t.id));
                  setShowNewTeacherPDF(false);
                  setTimeout(()=>setPrintPreview({html:buildF3Html(list3,S,ay,sh)}),50);
                }}
                style={{...BS("#B91C1C"),flex:2,minWidth:120,opacity:selectedTeachersPDF.length?1:0.4}}>
                🖨️ แบบ 3 — รหัสห้อง+สรุปวิชา ({selectedTeachersPDF.length} หน้า)
              </button>
            </div>
          </div>
        </div>
      )}

      <input data-ui-control="true" ref={fileRefSched} type="file" accept=".json" hidden onChange={importScheduleJSON}/>
    </div>
  </div>
  </div>;
}

/* ===== CONFLICT CHECKER ===== */
function checkConflicts(S) {
  const conflicts = [];
  const teacherSlots = {}; // teacherId → set of "day_period"
  const roomSlots = {};    // roomId → set of "day_period"
  const srSlots = {};      // specialRoomId → set of "day_period"

  Object.entries(S.schedule || {}).forEach(([key, entries]) => {
    if (!entries?.length) return;
    const parts = key.split('_');
    const roomId = parts[0];
    const day = parts[1];
    const period = parts[2];
    const slotKey = `${day}_${period}`;

    // ตรวจห้องเรียนซ้ำ
    if (!roomSlots[roomId]) roomSlots[roomId] = {};
    const prevCount = roomSlots[roomId][slotKey] || 0;
    if (prevCount + entries.length > 1) {
      const room = S.rooms.find(r => r.id === roomId);
      conflicts.push({ type: 'room', key, day, period: parseInt(period),
        msg: `ห้อง ${room?.name || roomId} มี ${prevCount + entries.length} วิชาในคาบเดียวกัน` });
    }
    roomSlots[roomId][slotKey] = (prevCount || 0) + entries.length;

    entries.forEach(entry => {
      const allTeachers = [entry.teacherId, ...(entry.coTeacherIds || [])].filter(Boolean);
      allTeachers.forEach(tid => {
        if (!teacherSlots[tid]) teacherSlots[tid] = {};
        if (teacherSlots[tid][slotKey]) {
          const teacher = S.teachers.find(t => t.id === tid);
          conflicts.push({ type: 'teacher', key, day, period: parseInt(period),
            msg: `ครู${teacher?.firstName || tid} สอนซ้ำในคาบเดียวกัน` });
        }
        teacherSlots[tid][slotKey] = true;
      });

      // ตรวจห้องพิเศษซ้ำ
      if (entry.specialRoomId) {
        if (!srSlots[entry.specialRoomId]) srSlots[entry.specialRoomId] = {};
        if (srSlots[entry.specialRoomId][slotKey]) {
          const sr = S.specialRooms?.find(r => r.id === entry.specialRoomId);
          conflicts.push({ type: 'specialroom', key, day, period: parseInt(period),
            msg: `ห้องพิเศษ ${sr?.name || entry.specialRoomId} ถูกใช้ซ้ำ` });
        }
        srSlots[entry.specialRoomId][slotKey] = true;
      }
    });
  });

  return conflicts;
}

/* ===== TEACHER HOURS SUMMARY ===== */
function TeacherHoursSummary({ S }) {
  const DAYS_TH = ["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
  const PERIOD_IDS = [1,2,3,4,5,6,7];
  const [sortBy, setSortBy] = useState("name"); // name | used | rem
  const [filterDept, setFilterDept] = useState("");

  const rows = useMemo(() => {
    return S.teachers.map(t => {
      let used = 0;
      const dayMap = {}; // day → count
      DAYS_TH.forEach(d => { dayMap[d] = 0; });
      const seen = new Set();
      Object.entries(S.schedule || {}).forEach(([k, en]) => {
        const pts = k.split("_");
        const day = pts[pts.length-2];
        en?.forEach(e => {
          const allT = [e.teacherId, ...(e.coTeacherIds||[])].filter(Boolean);
          if (!allT.includes(t.id)) return;
          const sub = S.subjects.find(s => s.id === e.subjectId);
          const ca = sub?.consecutiveAllowed || 0;
          if (ca === -1 || ca === -2) {
            const npKey = e.subjectId+"_"+pts[pts.length-2]+"_"+pts[pts.length-1];
            if (!seen.has(npKey)) { seen.add(npKey); used++; if(dayMap[day]!==undefined) dayMap[day]++; }
          } else {
            used++;
            if(dayMap[day]!==undefined) dayMap[day]++;
          }
        });
      });
      const tot = t.totalPeriods || 0;
      const dept = S.depts.find(d => d.id === t.departmentId);
      return { teacher: t, used, tot, rem: tot - used, dayMap, dept };
    });
  }, [S]);

  const filtered = rows
    .filter(r => !filterDept || r.teacher.departmentId === filterDept)
    .sort((a, b) => {
      if (sortBy === "used") return b.used - a.used;
      if (sortBy === "rem") return b.rem - a.rem;
      return (a.teacher.firstName||"").localeCompare(b.teacher.firstName||"", "th");
    });

  const overloaded = filtered.filter(r => r.rem < 0);
  const underloaded = filtered.filter(r => r.tot > 0 && r.rem > 0 && r.used < r.tot * 0.5);

  const LS = {fontSize:13, fontWeight:600, display:"block", marginBottom:6};
  return (
    <div>
      {(overloaded.length > 0 || underloaded.length > 0) && (
        <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap"}}>
          {overloaded.length > 0 && (
            <div style={{background:"#FEF2F2",border:"1px solid #FECACA",borderRadius:10,padding:"10px 16px",flex:1,minWidth:200}}>
              <div style={{fontSize:12,fontWeight:700,color:"#B91C1C",marginBottom:4}}>⚠️ สอนเกินกำหนด ({overloaded.length} คน)</div>
              {overloaded.slice(0,3).map(r=><div key={r.teacher.id} style={{fontSize:12,color:"#7F1D1D"}}>{r.teacher.firstName} {r.teacher.lastName}: {r.used}/{r.tot} คาบ (+{Math.abs(r.rem)})</div>)}
              {overloaded.length > 3 && <div style={{fontSize:11,color:"#9CA3AF"}}>...และอีก {overloaded.length-3} คน</div>}
            </div>
          )}
          {underloaded.length > 0 && (
            <div style={{background:"#FFFBEB",border:"1px solid #FDE68A",borderRadius:10,padding:"10px 16px",flex:1,minWidth:200}}>
              <div style={{fontSize:12,fontWeight:700,color:"#D97706",marginBottom:4}}>💡 ยังสอนน้อยกว่า 50% ({underloaded.length} คน)</div>
              {underloaded.slice(0,3).map(r=><div key={r.teacher.id} style={{fontSize:12,color:"#78350F"}}>{r.teacher.firstName} {r.teacher.lastName}: {r.used}/{r.tot} คาบ</div>)}
              {underloaded.length > 3 && <div style={{fontSize:11,color:"#9CA3AF"}}>...และอีก {underloaded.length-3} คน</div>}
            </div>
          )}
        </div>
      )}

      <div style={{display:"flex",gap:10,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
        <select data-ui-control="true" value={filterDept} onChange={e=>setFilterDept(e.target.value)} style={{padding:"7px 10px",borderRadius:8,border:"1px solid #D1D5DB",fontSize:13,flex:1,minWidth:160}}>
          <option value="">กลุ่มสาระทั้งหมด</option>
          {S.depts.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}
        </select>
        <div style={{display:"flex",gap:6}}>
          {[["name","ชื่อ"],["used","คาบสอน"],["rem","คงเหลือ"]].map(([v,l])=>(
            <button data-ui-control="true" key={v} onClick={()=>setSortBy(v)} style={{padding:"6px 12px",borderRadius:8,fontSize:12,fontWeight:600,border:`2px solid ${sortBy===v?"#B91C1C":"#E5E7EB"}`,background:sortBy===v?"#FEE2E2":"#fff",color:sortBy===v?"#B91C1C":"#6B7280",cursor:"pointer"}}>เรียง{l}</button>
          ))}
        </div>
      </div>

      <div style={{overflowX:"auto"}}>
        <table data-ui-table="true" style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead>
            <tr style={{background:"#F9FAFB"}}>
              <th style={{padding:"8px 10px",textAlign:"left",border:"1px solid #E5E7EB",minWidth:120}}>ครู</th>
              <th style={{padding:"8px 6px",textAlign:"center",border:"1px solid #E5E7EB",minWidth:60}}>กำหนด</th>
              <th style={{padding:"8px 6px",textAlign:"center",border:"1px solid #E5E7EB",minWidth:60}}>สอนแล้ว</th>
              <th style={{padding:"8px 6px",textAlign:"center",border:"1px solid #E5E7EB",minWidth:60}}>คงเหลือ</th>
              {DAYS_TH.map(d=><th key={d} style={{padding:"6px 4px",textAlign:"center",border:"1px solid #E5E7EB",fontSize:11,minWidth:40}}>{d.slice(0,3)}</th>)}
            </tr>
          </thead>
          <tbody>
            {filtered.map((r,i)=>{
              const overload = r.rem < 0;
              const ok = r.tot > 0 && r.used === r.tot;
              return (
                <tr key={r.teacher.id} style={{background: overload?"#FEF2F2": ok?"#F0FDF4": i%2===1?"#F9FAFB":"#fff"}}>
                  <td style={{padding:"6px 10px",border:"1px solid #E5E7EB"}}>
                    <div style={{fontWeight:600,color:"#1F2937"}}>{r.teacher.firstName} {r.teacher.lastName}</div>
                    <div style={{fontSize:10,color:"#6B7280"}}>{r.dept?.name||""}</div>
                  </td>
                  <td style={{padding:"6px",textAlign:"center",border:"1px solid #E5E7EB",fontWeight:600}}>{r.tot||"—"}</td>
                  <td style={{padding:"6px",textAlign:"center",border:"1px solid #E5E7EB",fontWeight:700,color: overload?"#B91C1C": ok?"#059669":"#1F2937"}}>{r.used}</td>
                  <td style={{padding:"6px",textAlign:"center",border:"1px solid #E5E7EB",color: overload?"#B91C1C": r.rem>0?"#D97706":"#059669",fontWeight:600}}>{r.tot?r.rem:"—"}</td>
                  {DAYS_TH.map(d=>(
                    <td key={d} style={{padding:"4px",textAlign:"center",border:"1px solid #E5E7EB",fontSize:11,color:r.dayMap[d]>0?"#1F2937":"#D1D5DB"}}>
                      {r.dayMap[d]>0?r.dayMap[d]:"·"}
                    </td>
                  ))}
                </tr>
              );
            })}
            {!filtered.length && <tr><td colSpan={11} style={{padding:24,textAlign:"center",color:"#9CA3AF"}}>ไม่มีข้อมูล</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ===== CONFLICT PANEL ===== */
function ConflictPanel({ S }) {
  const conflicts = useMemo(() => checkConflicts(S), [S]);
  const DAYS_TH = ["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
  if (!conflicts.length) return (
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:48,color:"#059669"}}>
      <div style={{fontSize:48,marginBottom:12}}>✅</div>
      <div style={{fontSize:16,fontWeight:700}}>ไม่พบความขัดแย้ง!</div>
      <div style={{fontSize:13,color:"#6B7280",marginTop:4}}>ตารางสอนสมบูรณ์</div>
    </div>
  );
  return (
    <div>
      <div style={{background:"#FEF2F2",border:"1px solid #FECACA",borderRadius:10,padding:"12px 16px",marginBottom:16}}>
        <div style={{fontWeight:700,color:"#B91C1C"}}>⚠️ พบ {conflicts.length} ความขัดแย้ง</div>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {conflicts.map((c,i)=>{
          const typeIcon = c.type==="teacher"?"👨‍🏫": c.type==="room"?"🏫":"⭐";
          const typeLabel = c.type==="teacher"?"ครูสอนซ้ำ": c.type==="room"?"ห้องซ้ำ":"ห้องพิเศษซ้ำ";
          return (
            <div data-ui-surface="true" key={i} style={{background:"#fff",border:"1px solid #FCA5A5",borderRadius:10,padding:"10px 14px",display:"flex",gap:12,alignItems:"flex-start"}}>
              <div style={{fontSize:20}}>{typeIcon}</div>
              <div style={{flex:1}}>
                <div style={{fontWeight:600,fontSize:13,color:"#B91C1C"}}>{typeLabel}</div>
                <div style={{fontSize:13,color:"#374151",marginTop:2}}>{c.msg}</div>
                <div style={{fontSize:11,color:"#9CA3AF",marginTop:2}}>วัน{DAYS_TH.find(d=>c.day===d)||c.day} คาบ {c.period}</div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function PrintDesignerModal(props){return <PrintStudio {...props} getPeriodCfg={getPeriodCfg} S={{...props.S,printPeriodConfigs:Object.fromEntries(props.S.rooms.map(r=>[r.id,getPeriodCfg(getDivisionForLevel(r.levelId,props.S.levels))]))}}/>;}

/* ===== SETTINGS */
function Settings({S,U,st,ay,setAY,sh,setSH,div,setSyncing,stateRef}){
  const logoRef=useRef(null);
  // helper: ล้าง localStorage dara_preview_ keys และ force sync ไป GAS
  const clearLocalAndSync=async(newState)=>{
    // ล้าง localStorage ทุก key ของ division นี้
    Object.keys(localStorage)
      .filter(k=>k.startsWith("dara_preview_"+div?.id)||k==="dara_preview_division")
      .forEach(k=>localStorage.removeItem(k));
    // force sync ข้อมูลใหม่ไป GAS ทันที
    const {db}=getFB();
    if(db){ setSyncing(true); try{ await fsSaveTimetable(div?.id||"m2",newState); }catch(e){} setSyncing(false); }
  };

  const resetAll=async()=>{
    if(!PREVIEW_MODE){st('ต้องตั้งค่าสิทธิ์ผู้ดูแลก่อนรีเซ็ตข้อมูลจริง','error');return;}
    if(!await uiConfirm("⚠️ คุณแน่ใจหรือไม่ว่าต้องการลบข้อมูลทั้งหมด?\nข้อมูลที่จัดตารางไว้จะหายทั้งหมด!"))return;
    if(!await uiConfirm("ยืนยันอีกครั้ง — ลบข้อมูลทั้งหมดและเริ่มต้นใหม่?"))return;
    const newLevels=(div?.defaultLevels||["ระดับ 1","ระดับ 2","ระดับ 3"]).map(n=>({id:gid(),name:n}));
    const emptyState={levels:newLevels,plans:[],depts:[],teachers:[],subjects:[],rooms:[],specialRooms:[],assigns:[],meetings:[],schedule:{},locks:{}};
    U.setLevels(newLevels);
    U.setPlans([]);U.setDepts([]);U.setTeachers([]);U.setSubjects([]);
    U.setRooms([]);U.setSpecialRooms([]);U.setAssigns([]);U.setMeetings([]);U.setSchedule({});U.setLocks({});
    await clearLocalAndSync(emptyState);
    st("รีเซ็ทข้อมูลทั้งหมดแล้ว และ sync แล้ว","warning");
  };
  const resetScheduleOnly=async()=>{
    if(!PREVIEW_MODE){st('ต้องตั้งค่าสิทธิ์ผู้ดูแลก่อนรีเซ็ตข้อมูลจริง','error');return;}
    if(!await uiConfirm("ลบเฉพาะข้อมูลตารางสอน (ข้อมูลครู/วิชา/ห้องยังอยู่)?"))return;
    U.setSchedule({});U.setLocks({});
    ["schedule","locks"].forEach(k=>localStorage.removeItem("dara_preview_"+div?.id+"_"+k));
    const {db:db2}=getFB();
    if(db2){ setSyncing(true); try{ await fsSaveTimetable(div?.id||"m2",{...stateRef.current,schedule:{},locks:{}}); }catch(e){} setSyncing(false); }
    st("ล้างตารางสอนแล้ว และ sync แล้ว","warning");
  };
  const handleLogo=(e)=>{const f=e.target.files?.[0];if(!f)return;const reader=new FileReader();reader.onload=ev=>{setSH(p=>({...p,logo:ev.target.result}));st("อัพโหลดโลโก้สำเร็จ")};reader.readAsDataURL(f);e.target.value=""};

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div className="settings-stack"><section className="settings-section"><h3>ปีการศึกษา</h3>
        
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          <div><label style={LS}>ปีการศึกษา (พ.ศ.)</label><input data-ui-control="true" style={IS} value={ay.year} onChange={e=>{
            setAY(p=>({...p,year:e.target.value}));
          }} placeholder="2568"/></div>
          <div><label style={LS}>ภาคเรียนที่</label><select data-ui-control="true" style={IS} value={ay.semester} onChange={e=>setAY(p=>({...p,semester:e.target.value}))}><option value="1">1</option><option value="2">2</option></select></div>
          <button data-ui-control="true" onClick={async ()=>{
            if(!await uiConfirm(`เปลี่ยนปีการศึกษา → รีเซ็ตครูประจำชั้นทุกห้องด้วยไหม?\n(กด OK = รีเซ็ต, Cancel = ไม่รีเซ็ต)`))return;
            U.setRooms(p=>p.map(r=>({...r,homeroom1:"",homeroom2:"",homeroomCo:""})));
            st("รีเซ็ตครูประจำชั้นทุกห้องแล้ว","warning");
          }} style={{...BO("#D97706"),fontSize:12}}>🔄 รีเซ็ตครูประจำชั้นทุกห้อง (เมื่อเปลี่ยนปี)</button>
        </div>
      </section>
<details className="settings-section "><summary>หัวเอกสารและโลโก้</summary><div>
        
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          <div><label style={LS}>ชื่อโรงเรียน</label><input data-ui-control="true" style={IS} value={sh.name} onChange={e=>setSH(p=>({...p,name:e.target.value}))} placeholder="โรงเรียนดาราวิทยาลัย"/></div>
          <div>
            <label style={LS}>โลโก้โรงเรียน (จะแสดงในตาราง PDF)</label>
            <div style={{display:"flex",alignItems:"center",gap:14,marginTop:8}}>
              {sh.logo
                ?<img src={sh.logo} alt="logo" style={{width:56,height:56,borderRadius:"50%",objectFit:"cover",border:"2px solid #E5E7EB"}} onLoad={e=>{e.currentTarget.style.display='block'}} onError={e=>{e.currentTarget.style.display='none'}}/>
                :<div style={{width:56,height:56,borderRadius:"50%",background:"#F3F4F6",border:"2px dashed #D1D5DB",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:"#9CA3AF"}}>LOGO</div>
              }
              <div style={{flex:1,display:"flex",flexDirection:"column",gap:8}}>
                <input data-ui-control="true"
                  style={{...IS,fontSize:12}}
                  value={sh.logo||""}
                  onChange={e=>setSH(p=>({...p,logo:e.target.value}))}
                  placeholder="วาง URL รูปภาพ เช่น https://drive.google.com/uc?id=..."
                />
                <div style={{display:"flex",gap:6}}>
                  <button data-ui-control="true" onClick={()=>logoRef.current?.click()} style={{...BO("#2563EB"),fontSize:12,padding:"6px 12px"}}><Icon name="upload" size={13}/>Upload ไฟล์ (เครื่องนี้เท่านั้น)</button>
                  {sh.logo&&<button data-ui-control="true" onClick={()=>{setSH(p=>({...p,logo:""}));st("ลบโลโก้แล้ว","warning")}} style={{...BO("#DC2626"),fontSize:12,padding:"6px 12px"}}><Icon name="trash" size={13}/>ลบ</button>}
                </div>
              </div>
              <input data-ui-control="true" ref={logoRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleLogo}/>
            </div>
            <div style={{padding:"8px 12px",background:"#EFF6FF",borderRadius:8,marginTop:8,fontSize:12,color:"#1E40AF"}}>
              💡 <strong>แนะนำ:</strong> อัพโลโก้ขึ้น Google Drive → คลิกขวา → "Get link" → เปลี่ยน <code>drive.google.com/file/d/ID/view</code> เป็น <code>drive.google.com/uc?id=ID</code> แล้ววาง URL ด้านบน — ทุกเครื่องจะเห็นโลโก้เดียวกัน
            </div>
          </div>
        </div>
      </div></details>
<details className="settings-section danger-zone"><summary>ล้างข้อมูล / เริ่มต้นใหม่</summary><div>
        
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <button data-ui-control="true" onClick={resetScheduleOnly} style={BO("#D97706")}><Icon name="trash" size={16}/>ล้างเฉพาะตารางสอน</button>
          <button data-ui-control="true" onClick={()=>{
            // เว้น academicYear และ schoolHeader ไว้ ลบแค่ข้อมูลหลัก
            const keepKeys=["dara_preview_academicYear","dara_preview_schoolHeader","dara_preview_division"];
            Object.keys(localStorage)
              .filter(k=>k.startsWith("dara_preview_")&&!keepKeys.includes(k))
              .forEach(k=>localStorage.removeItem(k));
            st("ล้าง Cache แล้ว — กำลัง reload...","warning");
            setTimeout(()=>window.location.reload(),1000);
          }} style={BO("#6B7280")}><Icon name="x" size={16}/>ล้าง Cache (แก้ข้อมูลไม่ตรง)</button>
          <div style={{fontSize:12,color:"#6B7280"}}>ลบข้อมูลตารางสอนที่จัดไว้ แต่ข้อมูลครู วิชา ห้อง ยังอยู่</div>
          <div style={{borderTop:"1px solid #E5E7EB",paddingTop:12,marginTop:4}}/>
          <button data-ui-control="true" onClick={resetAll} style={BS("#DC2626")}><Icon name="trash" size={16}/>รีเซ็ทข้อมูลทั้งหมด</button>
          <div style={{fontSize:12,color:"#DC2626"}}>⚠️ ลบข้อมูลทุกอย่าง — ไม่สามารถกู้คืนได้</div>
        </div>
      </div></details>
<details className="settings-section "><summary>สรุปข้อมูลในระบบ</summary><div>
        
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
      </div></details></div>
  </div>;
}


/* ===== PDF HELPER — ตามแบบฟอร์มดาราวิทยาลัย (A4 แนวตั้ง) ===== */

// Merge entries ที่วิชาเดียวกันในคาบเดียว → แสดงชื่อวิชาแค่ครั้งเดียว เรียงห้องลงมา

/* ===== PRINT SETTINGS ===== */
const PRINT_COLORS={
  "แดง":{header:"#B91C1C",headerText:"#fff",rowAlt:"#FFF5F5",border:"#991B1B"},
  "ดำ":{header:"#1F2937",headerText:"#fff",rowAlt:"#F9FAFB",border:"#374151"},
  "เทา":{header:"#6B7280",headerText:"#fff",rowAlt:"#F9FAFB",border:"#9CA3AF"},
  "เหลือง":{header:"#D97706",headerText:"#fff",rowAlt:"#FFFBEB",border:"#F59E0B"},
  "ขาว":{header:"#F3F4F6",headerText:"#000",rowAlt:"#FAFAFA",border:"#D1D5DB"},
};
const PRINT_FONTS=["Sarabun","TH SarabunNew","Arial","Tahoma"];
const DEFAULT_PRINT_SETTINGS={fontFamily:"TH SarabunNew",fontSize:100,color:"แดง",rowHeight:100,showAltRow:true,showBorder:true};
const loadPrintSettings=()=>{try{const s=localStorage.getItem("dara_preview_printSettings");return s?{...DEFAULT_PRINT_SETTINGS,...JSON.parse(s)}:DEFAULT_PRINT_SETTINGS;}catch{return DEFAULT_PRINT_SETTINGS;}};
const savePrintSettings=(s)=>{try{localStorage.setItem("dara_preview_printSettings",JSON.stringify(s));}catch{}};
function PrintSettingsPanel({open,onClose,onApply}){
  const [s,setS]=useState(loadPrintSettings);
  if(!open)return null;
  const u=(k,v)=>setS(p=>({...p,[k]:v}));
  return(
    <div style={{position:"fixed",inset:0,zIndex:3000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.5)"}}>
      <div data-ui-surface="true" style={{background:"#fff",borderRadius:16,padding:28,width:"min(540px,95vw)",maxHeight:"90vh",overflowY:"auto",boxShadow:"0 20px 60px rgba(0,0,0,0.3)",fontFamily:"'Sarabun',sans-serif"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:18}}>
          <h2 style={{fontSize:18,fontWeight:700}}>⚙️ ตั้งค่าการพิมพ์</h2>
          <button data-ui-control="true" onClick={onClose} style={{background:"none",border:"none",fontSize:20,cursor:"pointer",color:"#6B7280"}}>✕</button>
        </div>
        <div style={{marginBottom:16}}>
          <label style={{fontSize:13,fontWeight:600,display:"block",marginBottom:8}}>🔤 ฟอนต์</label>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {PRINT_FONTS.map(f=><button data-ui-control="true" key={f} onClick={()=>u("fontFamily",f)} style={{padding:"6px 14px",borderRadius:8,border:"2px solid "+(s.fontFamily===f?"#B91C1C":"#D1D5DB"),background:s.fontFamily===f?"#FEE2E2":"#fff",fontFamily:f,fontSize:13,cursor:"pointer",fontWeight:s.fontFamily===f?700:400}}>{f}</button>)}
          </div>
        </div>
        <div style={{marginBottom:16}}>
          <label style={{fontSize:13,fontWeight:600,display:"block",marginBottom:8}}>📏 ขนาดตัวอักษร: <b style={{color:"#B91C1C"}}>{s.fontSize}%</b></label>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <button data-ui-control="true" onClick={()=>u("fontSize",Math.max(60,s.fontSize-10))} style={{width:32,height:32,borderRadius:8,border:"1px solid #D1D5DB",background:"#F9FAFB",fontSize:16,cursor:"pointer"}}>−</button>
            <input data-ui-control="true" type="range" min={60} max={160} value={s.fontSize} onChange={e=>u("fontSize",parseInt(e.target.value))} style={{flex:1,accentColor:"#B91C1C"}}/>
            <button data-ui-control="true" onClick={()=>u("fontSize",Math.min(160,s.fontSize+10))} style={{width:32,height:32,borderRadius:8,border:"1px solid #D1D5DB",background:"#F9FAFB",fontSize:16,cursor:"pointer"}}>+</button>
            <button data-ui-control="true" onClick={()=>u("fontSize",100)} style={{fontSize:11,color:"#6B7280",background:"none",border:"1px solid #D1D5DB",borderRadius:6,padding:"3px 8px",cursor:"pointer"}}>รีเซ็ต</button>
          </div>
        </div>
        <div style={{marginBottom:16}}>
          <label style={{fontSize:13,fontWeight:600,display:"block",marginBottom:8}}>↕️ ความสูงแถว: <b style={{color:"#B91C1C"}}>{s.rowHeight}%</b></label>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <button data-ui-control="true" onClick={()=>u("rowHeight",Math.max(60,s.rowHeight-10))} style={{width:32,height:32,borderRadius:8,border:"1px solid #D1D5DB",background:"#F9FAFB",fontSize:16,cursor:"pointer"}}>−</button>
            <input data-ui-control="true" type="range" min={60} max={160} value={s.rowHeight} onChange={e=>u("rowHeight",parseInt(e.target.value))} style={{flex:1,accentColor:"#B91C1C"}}/>
            <button data-ui-control="true" onClick={()=>u("rowHeight",Math.min(160,s.rowHeight+10))} style={{width:32,height:32,borderRadius:8,border:"1px solid #D1D5DB",background:"#F9FAFB",fontSize:16,cursor:"pointer"}}>+</button>
            <button data-ui-control="true" onClick={()=>u("rowHeight",100)} style={{fontSize:11,color:"#6B7280",background:"none",border:"1px solid #D1D5DB",borderRadius:6,padding:"3px 8px",cursor:"pointer"}}>รีเซ็ต</button>
          </div>
        </div>
        <div style={{marginBottom:16}}>
          <label style={{fontSize:13,fontWeight:600,display:"block",marginBottom:8}}>🎨 สีหัวตาราง</label>
          <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
            {Object.entries(PRINT_COLORS).map(([name,c])=>(
              <button data-ui-control="true" key={name} onClick={()=>u("color",name)} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:4,padding:"8px 12px",borderRadius:10,border:"2px solid "+(s.color===name?"#B91C1C":"#E5E7EB"),background:s.color===name?"#FEF2F2":"#fff",cursor:"pointer"}}>
                <div style={{width:36,height:20,borderRadius:5,background:c.header,display:"flex",alignItems:"center",justifyContent:"center"}}><span style={{color:c.headerText,fontSize:9,fontWeight:700}}>วัน</span></div>
                <div style={{width:36,height:10,borderRadius:4,background:c.rowAlt,border:"1px solid #eee"}}/>
                <span style={{fontSize:11,fontWeight:s.color===name?700:400,color:s.color===name?"#B91C1C":"#374151"}}>{name}</span>
              </button>
            ))}
          </div>
        </div>
        <div style={{marginBottom:18,display:"flex",gap:20}}>
          <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",fontSize:13}}>
            <input data-ui-control="true" type="checkbox" checked={s.showAltRow} onChange={e=>u("showAltRow",e.target.checked)} style={{width:16,height:16,accentColor:"#B91C1C"}}/>สลับสีแถว
          </label>
          <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",fontSize:13}}>
            <input data-ui-control="true" type="checkbox" checked={s.showBorder} onChange={e=>u("showBorder",e.target.checked)} style={{width:16,height:16,accentColor:"#B91C1C"}}/>เส้นขอบ
          </label>
        </div>
        <div data-ui-surface="true" style={{background:"#F9FAFB",borderRadius:10,padding:10,marginBottom:18}}>
          <div style={{fontSize:11,color:"#6B7280",marginBottom:5}}>ตัวอย่าง:</div>
          <table style={{width:"100%",borderCollapse:"collapse",fontFamily:s.fontFamily,fontSize:(11*s.fontSize/100)+"px"}}>
            <thead><tr>{["วัน","คาบ 1","คาบ 2","คาบ 3"].map(h=><th key={h} style={{background:PRINT_COLORS[s.color].header,color:PRINT_COLORS[s.color].headerText,padding:"4px 6px",border:s.showBorder?"1px solid "+PRINT_COLORS[s.color].border:"none",fontWeight:700}}>{h}</th>)}</tr></thead>
            <tbody>{[["จันทร์","คณิต ม.5/1","","ฟิสิกส์ ม.5/2"],["อังคาร","","ชีวะ ม.5/3",""]].map((row,ri)=>(
              <tr key={ri} style={{background:s.showAltRow&&ri%2===1?PRINT_COLORS[s.color].rowAlt:"#fff"}}>
                {row.map((cell,ci)=><td key={ci} style={{padding:(3*s.rowHeight/100)+"px 6px",border:s.showBorder?"1px solid #E5E7EB":"none",fontSize:(10*s.fontSize/100)+"px",textAlign:"center",height:(26*s.rowHeight/100)+"px"}}>{cell}</td>)}
              </tr>
            ))}</tbody>
          </table>
        </div>
        <div style={{display:"flex",gap:10}}>
          <button data-ui-control="true" onClick={()=>setS(DEFAULT_PRINT_SETTINGS)} style={{...BO(),flex:1}}>↩ ค่าเริ่มต้น</button>
          <button data-ui-control="true" onClick={()=>{savePrintSettings(s);onApply(s);onClose();}} style={{...BS(),flex:2}}>💾 บันทึก &amp; ใช้งาน</button>
        </div>
      </div>
    </div>
  );
}



function buildTeacherTableHTML(teacher, S, ay, sh, ps) {
  const P=ps||DEFAULT_PRINT_SETTINGS;
  const C=PRINT_COLORS[P.color]||PRINT_COLORS["แดง"];
  const fScale=P.fontSize/100;
  const rScale=P.rowHeight/100;
  const yr=ay?.year||"2568";
  const logoImg=sh?.logo?`<img src="${sh.logo}" style="height:40px;vertical-align:middle;margin-right:8px;"/>` :"";
  const title=`ตารางสอน ${teacher.prefix||""}${teacher.firstName} ${teacher.lastName}`;
  const dept=S.depts.find(d=>d.id===teacher.departmentId)?.name||"";

  const getCells=(day,pid)=>{
    const results=[];
    S.rooms.forEach(room=>{
      const key=room.id+"_"+day+"_"+pid;
      (S.schedule[key]||[]).forEach(e=>{
        if(e.teacherId!==teacher.id&&!(e.coTeacherIds||[]).includes(teacher.id)) return;
        const sub=S.subjects.find(s=>s.id===e.subjectId);
        results.push({th:sub?.name||sub?.code||"",en:sub?.shortName||"",room:room.name});
      });
    });
    return results;
  };

  const colgroup=`<colgroup>
    <col style="width:9mm;"><col style="width:7mm;">
    <col><col>
    <col style="width:6mm;">
    <col><col>
    <col style="width:7mm;">
    <col><col>
    <col style="width:6mm;">
    <col>
  </colgroup>`;

  const vert=(txt,bg="#fffde7",fw="600",fs="9pt")=>
    `<div style="writing-mode:vertical-rl;transform:rotate(180deg);white-space:nowrap;font-size:${fs};font-weight:${fw};letter-spacing:1px;text-align:center;">${txt}</div>`;

  const HDR=[
    {label:"คาบ 1",time:"08.30-09.20"},
    {label:"คาบ 2",time:"09.20-10.10"},
    {label:"คาบ 3",time:"10.25-11.15"},
    {label:"คาบ 4",time:"11.15-12.05"},
    {label:"คาบ 5",time:"13.00-13.50"},
    {label:"คาบ 6",time:"14.00-14.50"},
    {label:"คาบ 7",time:"14.50-15.40"},
  ];
  const BRK=[["08.00-","08.30"],["10.10-","10.25"],["12.05-","13.00"],["14.40-","14.50"]];
  const vertBRK=(parts,fs="9pt")=>
    `<div style="writing-mode:vertical-rl;transform:rotate(180deg);font-size:${fs};font-weight:600;letter-spacing:1px;text-align:center;">${parts.map(p=>'<span style="white-space:nowrap;">'+p+'</span>').join("")}</div>`;

  const thStyle="border:1px solid #666;font-size:9pt;font-weight:bold;text-align:center;padding:2px;";
  const brkStyle="border:1px solid #666;background:#fffde7;padding:0;height:40px;vertical-align:middle;text-align:center;";
  const hdrRow=`<tr style="background:#f0f0f0;">
    <th style="border:1px solid #666;padding:0;position:relative;height:40px;">
      <svg style="position:absolute;top:0;left:0;width:100%;height:100%;" preserveAspectRatio="none">
        <line x1="0" y1="0" x2="100%" y2="100%" stroke="#888" stroke-width="0.8"/>
      </svg>
      <div style="position:absolute;top:2px;right:2px;font-size:7pt;font-weight:600;">เวลา</div>
      <div style="position:absolute;bottom:2px;left:2px;font-size:7pt;font-weight:600;">วัน</div>
    </th>
    <th style="${brkStyle}">${vertBRK(BRK[0])}</th>
    ${HDR.slice(0,2).map(h=>`<th style="${thStyle}"><div>${h.label}</div><div style="font-size:7pt;font-weight:400;color:#555;">${h.time}</div></th>`).join("")}
    <th style="${brkStyle}">${vertBRK(BRK[1])}</th>
    ${HDR.slice(2,4).map(h=>`<th style="${thStyle}"><div>${h.label}</div><div style="font-size:7pt;font-weight:400;color:#555;">${h.time}</div></th>`).join("")}
    <th style="${brkStyle}">${vertBRK(BRK[2])}</th>
    ${HDR.slice(4,6).map(h=>`<th style="${thStyle}"><div>${h.label}</div><div style="font-size:7pt;font-weight:400;color:#555;">${h.time}</div></th>`).join("")}
    <th style="${brkStyle}">${vertBRK(BRK[3])}</th>
    <th style="${thStyle}"><div>${HDR[6].label}</div><div style="font-size:7pt;font-weight:400;color:#555;">${HDR[6].time}</div></th>
  </tr>`;


  // หา assemblyDay จากห้องที่ครูสอน (ใช้ level ของห้องแรก)
  const teacherRooms=[...new Set(
    Object.keys(S.schedule).flatMap(k=>
      (S.schedule[k]||[]).filter(e=>e.teacherId===teacher.id||(e.coTeacherIds||[]).includes(teacher.id))
        .map(()=>k.split("_")[0])
    )
  )].map(id=>S.rooms.find(r=>r.id===id)).filter(Boolean);

  const DAYS_TH=["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
  const asmLevel=teacherRooms[0]?S.levels.find(l=>l.id===teacherRooms[0].levelId):null;
  let body="";

  DAYS_TH.forEach((day,di)=>{
    const D=[1,2,3,4,5,6,7].map(pid=>getCells(day,pid));
    const isMulti=[1,2,3,4,5,6,7].map(pid=>getCells(day,pid).length>1);
    const MBG="#eeeeee";
    const bgRow=di%2===0?"":"background:#fafafa;";
    const isAsm=asmLevel?.assemblyDay===day;

    const hmTxt=isAsm?"หอประชุม/Assembly":"Homeroom";
    const hmBg=isAsm?"#e8f5e9":"#fafff7";

    const ROW_H=`${(48*fScale).toFixed(0)}px`;
    const cell=(arr,multi=false)=>{
      const bg=multi?`background:${MBG};`:"";
      if(!arr.length) return`<td style="border:1px solid #ddd;padding:0;${bg}"><div style="height:${ROW_H};"></div></td>`;
      const inner=arr.map(c=>`<div style="font-size:${(8.5*fScale).toFixed(1)}pt;font-weight:bold;line-height:1.25;">${c.th}</div><div style="font-size:${(7.5*fScale).toFixed(1)}pt;color:#1a237e;line-height:1.2;">${c.room}</div>`).join('<hr style="border:none;border-top:1px dotted #bbb;margin:1px 0;"/>');
      return`<td style="border:1px solid #ddd;padding:0;${bg}"><div style="height:${ROW_H};overflow:hidden;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:2px;">${inner}</div></td>`;
    };
    const BKcell=(rows,txt)=>
      `<td rowspan="${rows}" style="border:1px solid #888;background:#fffde7;padding:0;vertical-align:middle;text-align:center;">${vert(txt,"#fffde7","600","9pt")}</td>`;

    const bk=di===0;
    const TOTAL_ROWS=DAYS_TH.length;
    const hmDisplay=isAsm?"หอประชุม":"Home room";

    body+=`
      <tr style="${bgRow}">
        <td style="border:1px solid #888;padding:0;background:#f5f5f5;"><div style="height:${ROW_H};display:flex;align-items:center;justify-content:center;font-weight:bold;font-size:${(9*fScale).toFixed(1)}pt;text-align:center;">${day==="พฤหัสบดี"?"พฤหัส":day}</div></td>
        <td style="border:1px solid #888;padding:0;background:${hmBg};"><div style="height:${ROW_H};display:flex;align-items:center;justify-content:center;text-align:center;font-size:${(7.5*fScale).toFixed(1)}pt;font-weight:600;line-height:1.3;">${hmDisplay}</div></td>
        ${cell(D[0],isMulti[0])}${cell(D[1],isMulti[1])}
        ${bk?BKcell(TOTAL_ROWS,"พักน้อย 15 นาที"):""}
        ${cell(D[2],isMulti[2])}${cell(D[3],isMulti[3])}
        ${bk?BKcell(TOTAL_ROWS,"พักกลางวัน 55 นาที"):""}
        ${cell(D[4],isMulti[4])}${cell(D[5],isMulti[5])}
        ${bk?BKcell(TOTAL_ROWS,"พักน้อย 10 นาที"):""}
        ${cell(D[6],isMulti[6])}
      </tr>`;
  });

  return`
    <div style="text-align:center;margin-bottom:5px;font-family:'TH SarabunNew','Sarabun',sans-serif;">
      ${logoImg}<b style="font-size:13pt;">${title}&emsp;&emsp;ปีการศึกษา ${yr}</b>
      ${dept?`<div style="font-size:9pt;color:#555;">${dept}</div>`:""}
    </div>
    <table style="width:100%;border-collapse:collapse;table-layout:fixed;border-spacing:0;font-family:'TH SarabunNew','Sarabun',sans-serif;">
      ${colgroup}
      <thead>${hdrRow}</thead>
      <tbody>${body}</tbody>
    </table>`;
}

function pdfPage(title, subtitle, dayRows, footerText, logoBase64, ps, isRoom, divisionId) {
  const P=ps||DEFAULT_PRINT_SETTINGS;
  const C=PRINT_COLORS[P.color]||PRINT_COLORS["แดง"];
  const fScale=P.fontSize/100;
  const rScale=P.rowHeight/100;
  const ff=P.fontFamily||"TH SarabunNew";
  const pcfg=getPeriodCfg(divisionId||"m2");
  const PLIST = pcfg.periods;

  const thNums = PLIST.map(p => '<th class="period-num">' + p.id + '</th>').join("");
  const thTimes = PLIST.map(p => '<th class="period-time">' + p.time + '</th>').join("");

  const bodyRows = dayRows.map(function(r) {
    const dayCells = r.cells.map(function(rawEntries) {
      if (!rawEntries || !rawEntries.length) return '<td class="slot"></td>';
      // lock cell — แสดงสีพิเศษ
      if (rawEntries[0]?.isLock) {
        const lc = rawEntries[0].lockColor||"#FEF3C7";
        const lt = rawEntries[0].lockTextColor||"#92400E";
        return '<td class="slot" style="background:'+lc+';vertical-align:middle"><div class="ent"><div class="ent-sub" style="font-size:'+(11*fScale).toFixed(1)+'px;color:'+lt+'">'+rawEntries[0].sub+'</div></div></td>';
      }
      const entries = groupEntries(rawEntries);
      const isDouble = entries.some(function(e){ return e.double || e.roomCount > 1; });
      const inner = entries.map(function(e) {
        let h = '<div class="ent"><div class="ent-sub">' + e.sub + '</div>' + (isRoom ? (e.roomHtmlTeacher||e.roomHtml) : e.roomHtml) + '</div>';
        return h;
      }).join("");
      return '<td class="slot' + (isDouble ? ' slot-hi' : '') + '">' + inner + '</td>';
    }).join("");
    return '<tr><td class="day-cell">' + (r.day==="พฤหัสบดี"?"พฤหัส":r.day) + '</td>' + dayCells + '</tr>';
  }).join("\n");

  const logoHtml = logoBase64
    ? '<img src="' + logoBase64 + '" style="width:48px;height:48px;border-radius:50%;object-fit:cover;flex-shrink:0"/>'
    : '<div class="logo">LOGO</div>';

  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    "@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');" +
    '@page{size:A4 portrait;margin:10mm 8mm}' +
    '*{margin:0;padding:0;box-sizing:border-box}' +
    'html,body{width:210mm;overflow-x:hidden}' +
    "body{font-family:'"+ff+"','Sarabun','Noto Sans Thai',sans-serif;font-size:"+(11*fScale).toFixed(1)+"px;color:#000}" +
    '.page{width:100%;max-width:190mm;margin:0 auto;position:relative}' +
    '.header-row{display:flex;align-items:center;margin-bottom:6px;gap:12px}' +
    '.logo{width:48px;height:48px;border:1.5px solid #999;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:8px;color:#666;flex-shrink:0}' +
    '.title-block{flex:1}' +
    '.title-main{font-size:'+(14*fScale).toFixed(1)+'px;font-weight:700}' +
    '.title-sub{font-size:'+(11*fScale).toFixed(1)+'px;color:#444;margin-top:2px}' +
    'table{width:100%;border-collapse:collapse;table-layout:fixed;margin-top:4px}' +
    'th,td{border:1px solid #000;text-align:center;vertical-align:middle}' +
    'th{padding:3px 1px;font-weight:700;background:'+C.header+';color:'+C.headerText+'}' +
    (P.showBorder?'':'th,td{border:none}') +
    'th.period-num{font-size:'+(13*fScale).toFixed(1)+'px;height:'+(24*rScale).toFixed(0)+'px}' +
    'th.period-time{font-size:'+(9*fScale).toFixed(1)+'px;height:'+(18*rScale).toFixed(0)+'px;font-weight:400;white-space:nowrap}' +
    'th.day-col{width:54px;font-size:'+(12*fScale).toFixed(1)+'px;font-weight:700}' +
    'td.day-cell{font-weight:700;font-size:'+(13*fScale).toFixed(1)+'px;padding:4px 2px;width:54px;background:#F3F4F6}' +
    'td.slot{padding:3px 2px;vertical-align:middle;height:'+(76*rScale).toFixed(0)+'px;text-align:center}' +
    'td.slot-hi{background:#eeeeee}' +
    (P.showAltRow?'tbody tr:nth-child(even){background:'+C.rowAlt+'}':'') +
    '.ent{margin-bottom:3px;display:flex;flex-direction:column;align-items:center;justify-content:center}' +
    '.ent-sub{font-weight:700;font-size:'+(13*fScale).toFixed(1)+'px;line-height:1.3}' +
    '.ent-room{font-size:'+(12*fScale).toFixed(1)+'px;color:#111;line-height:1.25}' +
    '.ent-room2{font-size:'+(11*fScale).toFixed(1)+'px;color:#333;line-height:1.2}' +
    '.sig-area{margin-top:16px;font-size:'+(11*fScale).toFixed(1)+'px}' +
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
    '<tr><th class="day-col" rowspan="2" style="vertical-align:middle;text-align:center;">วัน<br/><span style="font-size:9px;font-weight:400;white-space:nowrap">คาบ/เวลา</span></th>' + thNums + '</tr>' +
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
function pdfMultiPage(pages, logoBase64, ps, isRoom, divisionId) {
  const P=ps||DEFAULT_PRINT_SETTINGS;
  const C=PRINT_COLORS[P.color]||PRINT_COLORS["แดง"];
  const fScale=P.fontSize/100;
  const rScale=P.rowHeight/100;
  const ff=P.fontFamily||"TH SarabunNew";
  const pcfg=getPeriodCfg(divisionId||"m2");
  const PLIST = pcfg.periods;
  const thNums = PLIST.map(p => '<th class="period-num">' + p.id + '</th>').join("");
  const thTimes = PLIST.map(p => '<th class="period-time">' + p.time + '</th>').join("");

  const logoHtml = logoBase64
    ? '<img src="' + logoBase64 + '" style="width:36px;height:36px;border-radius:50%;object-fit:cover;flex-shrink:0"/>'
    : '<div class="logo">LOGO</div>';

  const buildBlock = (pg) => {
    const bodyRows = pg.dayRows.map(function(r) {
      const dayCells = r.cells.map(function(rawEntries) {
        if (!rawEntries || !rawEntries.length) return '<td class="slot"></td>';
        const entries = groupEntries(rawEntries);
        const isDouble = entries.some(function(e){ return e.double || e.roomCount > 1; });
        const inner = entries.map(function(e) {
          let h = '<div class="ent"><div class="ent-sub">' + e.sub + '</div>' + (isRoom ? (e.roomHtmlTeacher||e.roomHtml) : e.roomHtml) + '</div>';
          return h;
        }).join("");
        return '<td class="slot' + (isDouble ? ' slot-hi' : '') + '">' + inner + '</td>';
      }).join("");
      return '<tr><td class="day-cell">' + (r.day==="พฤหัสบดี"?"พฤหัส":r.day) + '</td>' + dayCells + '</tr>';
    }).join("\n");

    return '<div class="block">' +
      '<div class="header-row">' + logoHtml +
      '<div class="title-block"><div class="title-main">' + pg.title + '</div><div class="title-sub">' + pg.subtitle + '</div></div>' +
      '</div>' +
      '<table><thead>' +
      '<tr><th class="day-col" rowspan="2" style="vertical-align:middle;text-align:center;">วัน<br/><span style="font-size:8px;font-weight:400;white-space:nowrap">คาบ/เวลา</span></th>' + thNums + '</tr>' +
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
    'html,body{width:210mm;overflow-x:hidden}' +
    "body{font-family:'"+ff+"','Sarabun','Noto Sans Thai',sans-serif;font-size:"+(11*fScale).toFixed(1)+"px;color:#000}" +
    '.sheet{page-break-after:always}' +
    '.sheet:last-child{page-break-after:avoid}' +
    '.block{}' +
    '.header-row{display:flex;align-items:center;margin-bottom:4px;gap:8px}' +
    '.logo{width:36px;height:36px;border:1px solid #999;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:7px;color:#666;flex-shrink:0}' +
    '.title-block{flex:1}' +
    '.title-main{font-size:'+(14*fScale).toFixed(1)+'px;font-weight:700}' +
    '.title-sub{font-size:'+(11*fScale).toFixed(1)+'px;color:#444;margin-top:1px}' +
    'table{width:100%;border-collapse:collapse;table-layout:fixed;margin-top:3px}' +
    'th,td{border:1px solid #000;text-align:center;vertical-align:middle}' +
    'th{padding:2px 1px;font-weight:700;background:'+C.header+';color:'+C.headerText+'}' +
    (P.showBorder?'':'th,td{border:none}') +
    'th.period-num{font-size:'+(13*fScale).toFixed(1)+'px;height:'+(22*rScale).toFixed(0)+'px}' +
    'th.period-time{font-size:'+(9*fScale).toFixed(1)+'px;height:'+(16*rScale).toFixed(0)+'px;font-weight:400;white-space:nowrap}' +
    'th.day-col{width:48px;font-size:'+(12*fScale).toFixed(1)+'px;font-weight:700}' +
    'td.day-cell{font-weight:700;font-size:'+(13*fScale).toFixed(1)+'px;padding:2px;width:48px;background:#F3F4F6}' +
    'td.slot{padding:2px 1px;vertical-align:middle;height:'+(62*rScale).toFixed(0)+'px;text-align:center}' +
    'td.slot-hi{background:#eeeeee}' +
    (P.showAltRow?'tbody tr:nth-child(even){background:'+C.rowAlt+'}':'') +
    '.ent{margin-bottom:2px;display:flex;flex-direction:column;align-items:center;justify-content:center}' +
    '.ent-sub{font-weight:700;font-size:'+(12*fScale).toFixed(1)+'px;line-height:1.3}' +
    '.ent-room{font-size:'+(11*fScale).toFixed(1)+'px;color:#111;line-height:1.2}' +
    '.ent-room2{font-size:'+(10*fScale).toFixed(1)+'px;color:#333;line-height:1.15}' +
    '.sig-area{margin-top:5px;font-size:'+(10*fScale).toFixed(1)+'px}' +
    '.sig-flex{display:flex;justify-content:space-between;padding:0 15px}' +
    '.sig-box{text-align:center}' +
    '.sig-line{display:inline-block;width:130px;border-bottom:1px dotted #000;margin-bottom:2px}' +
    '.divider{border:none;border-top:1.5px dashed #aaa;margin:5px 0}' +
    '@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}' +
    '</style></head><body>' +
    pagesHtml +
    '</body></html>';
}


