import { useState, useCallback, useEffect, useRef, useMemo } from "react";
import * as XLSX from 'xlsx';
import { initializeApp } from "firebase/app";
import { getAuth, GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged } from "firebase/auth";
import { getFirestore, doc, getDoc, setDoc, collection, getDocs, onSnapshot } from "firebase/firestore";

// ===== FIREBASE CONFIG — ใส่ค่าจาก Firebase Console =====
const FIREBASE_CONFIG = {
  apiKey:            "AIzaSyC_anUKRySlNxZSoM5euqWqaM3amgskUIk",
  authDomain:        "dara-timetable.firebaseapp.com",
  projectId:         "dara-timetable",
  storageBucket:     "dara-timetable.firebasestorage.app",
  messagingSenderId: "773925099624",
  appId:             "1:773925099624:web:8ff141bbf52db0030303dd",
};
// =========================================================

const ADMIN_PIN = "100625";

// Firebase instances (lazy init เพื่อกัน crash ถ้ายังไม่ได้ตั้งค่า)
let _fbApp=null, _auth=null, _db=null;
const getFB=()=>{
  if(!_fbApp&&!FIREBASE_CONFIG.apiKey.includes("YOUR")){
    _fbApp=initializeApp(FIREBASE_CONFIG);
    _auth=getAuth(_fbApp);
    _db=getFirestore(_fbApp);
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

// ===== LOGIN SCREEN =====
function LoginScreen({onLogin}){
  const [loading,setLoading]=useState(false);
  const [err,setErr]=useState("");

  const handleGoogle=async()=>{
    const {auth}=getFB();
    if(!auth){setErr("ยังไม่ได้ตั้งค่า Firebase Config ใน App.jsx");return;}
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
      <div style={{background:"#fff",borderRadius:20,padding:"48px 40px",width:400,textAlign:"center",boxShadow:"0 25px 60px rgba(0,0,0,0.3)"}}>
        <div style={{fontSize:48,marginBottom:16}}>📋</div>
        <h1 style={{fontSize:22,fontWeight:700,marginBottom:4}}>ระบบจัดตารางสอน</h1>
        <p style={{color:"#6B7280",fontSize:13,marginBottom:32}}>โรงเรียนดาราวิทยาลัย</p>
        <button
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
function AdminPanel({user,onBack,refreshPerms}){
  const [pin,setPin]=useState("");
  const [unlocked,setUnlocked]=useState(false);
  const [pinErr,setPinErr]=useState("");
  const [users,setUsers]=useState([]);
  const [loading,setLoading]=useState(false);
  const [search,setSearch]=useState("");
  const [saving,setSaving]=useState(false);
  const [toast,setToast]=useState(null);

  // เพิ่มอีเมลล่วงหน้า
  const [addEmail,setAddEmail]=useState("");
  const [addPerms,setAddPerms]=useState({p1:false,p2:false,m1:false,m2:false});
  const [addLoading,setAddLoading]=useState(false);

  // Edit existing user
  const [editUid,setEditUid]=useState(null);
  const [editPerms,setEditPerms]=useState({});

  const divNames={p1:"ประถมต้น",p2:"ประถมปลาย",m1:"มัธยมต้น",m2:"มัธยมปลาย"};

  const showToast=(msg,type="success")=>{setToast({msg,type});setTimeout(()=>setToast(null),3000);};

  const tryPin=()=>{
    if(pin===ADMIN_PIN){setUnlocked(true);loadUsers();}
    else{setPinErr("รหัสไม่ถูกต้อง");setPin("");}
  };

  const loadUsers=async()=>{
    setLoading(true);
    const {db}=getFB();if(!db){setLoading(false);return;}
    const snap=await getDocs(collection(db,"permissions"));
    setUsers(snap.docs.map(d=>({uid:d.id,...d.data()})));
    setLoading(false);
  };

  // เพิ่ม / อัปเดตผู้ใช้จากอีเมล (ใช้อีเมลเป็น uid placeholder)
  const makePreKey=(email)=>"pre_"+email.trim().toLowerCase().replace(/[@.]/g,"_");

  const handleAddEmail=async()=>{
    const email=addEmail.trim().toLowerCase();
    if(!email){showToast("กรุณากรอกอีเมล","error");return;}
    if(!email.includes("@")){showToast("รูปแบบอีเมลไม่ถูกต้อง","error");return;}
    setAddLoading(true);
    const {db}=getFB();
    if(!db){showToast("Firebase ไม่พร้อม","error");setAddLoading(false);return;}

    // ค้นหาว่ามี doc ที่มี email นี้อยู่แล้วไหม (จาก users ที่โหลดมา = login จริงแล้ว)
    const existing=users.find(u=>u.email===email&&!u.preAdded);
    const uid=existing?existing.uid:makePreKey(email);

    await setDoc(doc(db,"permissions",uid),{
      email,
      displayName:existing?.displayName||"",
      divisions:addPerms,
      preAdded:!existing,
      merged:false,
    },{merge:true});

    if(existing){
      setUsers(p=>p.map(u=>u.uid===uid?{...u,divisions:addPerms}:u));
    } else {
      // ลบ pre doc เก่าถ้ามี แล้วเพิ่มใหม่
      setUsers(p=>{
        const filtered=p.filter(u=>u.uid!==uid);
        return [...filtered,{uid,email,displayName:"",divisions:addPerms,preAdded:true}];
      });
    }
    setAddEmail("");
    setAddPerms({p1:false,p2:false,m1:false,m2:false});
    showToast("บันทึกสิทธิ์สำเร็จ: "+email);
    setAddLoading(false);
    if(refreshPerms)refreshPerms();
  };

  const savePerms=async()=>{
    if(!editUid)return;
    setSaving(true);
    await fsSetPermissions(editUid,{divisions:editPerms});
    setUsers(p=>p.map(u=>u.uid===editUid?{...u,divisions:editPerms}:u));
    setSaving(false);
    setEditUid(null);setEditPerms({});
    showToast("บันทึกสิทธิ์สำเร็จ");
    if(refreshPerms)refreshPerms();
  };

  const deleteUser=async(uid)=>{
    if(!confirm("ลบผู้ใช้นี้ออกจากระบบ?"))return;
    const {db}=getFB();if(!db)return;
    await setDoc(doc(db,"permissions",uid),{divisions:{p1:false,p2:false,m1:false,m2:false}},{merge:true});
    setUsers(p=>p.map(u=>u.uid===uid?{...u,divisions:{p1:false,p2:false,m1:false,m2:false}}:u));
    showToast("ถอนสิทธิ์แล้ว","warning");
  };

  if(!unlocked) return(
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"#F3F4F6"}}>
      <div style={{background:"#fff",borderRadius:16,padding:"40px 36px",width:360,textAlign:"center",boxShadow:"0 4px 20px rgba(0,0,0,0.1)"}}>
        <div style={{fontSize:36,marginBottom:12}}>🔐</div>
        <h2 style={{fontSize:18,fontWeight:700,marginBottom:4}}>Admin Panel</h2>
        <p style={{color:"#6B7280",fontSize:12,marginBottom:24}}>ใส่รหัสผู้ดูแลระบบ</p>
        <input
          type="password"
          style={{...IS,textAlign:"center",letterSpacing:6,fontSize:20,marginBottom:12}}
          value={pin}
          onChange={e=>setPin(e.target.value)}
          onKeyDown={e=>e.key==="Enter"&&tryPin()}
          placeholder="• • • • • •"
          maxLength={10}
        />
        {pinErr&&<div style={{color:"#DC2626",fontSize:12,marginBottom:8}}>{pinErr}</div>}
        <button onClick={tryPin} style={{...BS(),width:"100%",justifyContent:"center"}}>ยืนยัน</button>
        <button onClick={onBack} style={{marginTop:10,background:"none",border:"none",color:"#6B7280",cursor:"pointer",fontSize:13}}>← กลับ</button>
      </div>
    </div>
  );

  const filtered=users.filter(u=>(u.email||u.displayName||u.uid).toLowerCase().includes(search.toLowerCase()));

  return(
    <div style={{minHeight:"100vh",background:"#F3F4F6",padding:24,fontFamily:"'Sarabun','Noto Sans Thai',sans-serif"}}>
      {toast&&<div style={{position:"fixed",top:20,right:20,zIndex:9999,background:toast.type==="error"?"#DC2626":toast.type==="warning"?"#D97706":"#059669",color:"#fff",padding:"12px 20px",borderRadius:10,fontSize:14,fontWeight:600,boxShadow:"0 4px 20px rgba(0,0,0,0.2)"}}>{toast.msg}</div>}
      <div style={{maxWidth:960,margin:"0 auto"}}>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:24}}>
          <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:"#6B7280"}}><Icon name="x" size={20}/></button>
          <h1 style={{fontSize:20,fontWeight:700}}>Admin Panel — จัดการสิทธิ์</h1>
        </div>

        {/* ── เพิ่มอีเมลล่วงหน้า ── */}
        <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",marginBottom:20}}>
          <h2 style={{fontSize:15,fontWeight:700,marginBottom:4}}>➕ กำหนดสิทธิ์ล่วงหน้า (Admin พิมพ์อีเมลเอง)</h2>
          <p style={{fontSize:12,color:"#6B7280",marginBottom:16}}>เพิ่มอีเมลพร้อมสิทธิ์ได้เลย — เมื่อผู้ใช้ login ครั้งแรกระบบจะจำสิทธิ์ที่ตั้งไว้</p>
          <div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap"}}>
            <div style={{flex:"1 1 280px"}}>
              <label style={LS}>อีเมล</label>
              <input
                style={IS}
                value={addEmail}
                onChange={e=>setAddEmail(e.target.value)}
                onKeyDown={e=>e.key==="Enter"&&handleAddEmail()}
                placeholder="teacher@web1.dara.ac.th"
                type="email"
              />
            </div>
            <div style={{flex:"1 1 auto"}}>
              <label style={LS}>ระดับที่เข้าได้</label>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                {Object.entries(divNames).map(([k,name])=>(
                  <label key={k} style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer",padding:"8px 12px",borderRadius:8,border:`2px solid ${addPerms[k]?"#DC2626":"#D1D5DB"}`,background:addPerms[k]?"#FEE2E2":"#F9FAFB",userSelect:"none"}}>
                    <input type="checkbox" checked={!!addPerms[k]} onChange={e=>setAddPerms(p=>({...p,[k]:e.target.checked}))} style={{width:15,height:15,accentColor:"#DC2626"}}/>
                    <span style={{fontSize:13,fontWeight:addPerms[k]?700:400,color:addPerms[k]?"#991B1B":"#374151"}}>{name}</span>
                  </label>
                ))}
              </div>
            </div>
            <button
              onClick={handleAddEmail}
              disabled={addLoading}
              style={{...BS(),flexShrink:0,opacity:addLoading?0.6:1}}
            >
              {addLoading?"กำลังบันทึก...":"บันทึกสิทธิ์"}
            </button>
          </div>
        </div>

        {/* ── ค้นหา ── */}
        <div style={{background:"#fff",borderRadius:12,padding:16,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",marginBottom:16}}>
          <div style={{position:"relative"}}>
            <input style={{...IS,paddingLeft:36}} value={search} onChange={e=>setSearch(e.target.value)} placeholder="ค้นหาชื่อหรืออีเมล..."/>
            <div style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:"#9CA3AF"}}><Icon name="search" size={14}/></div>
          </div>
        </div>

        {loading&&<div style={{textAlign:"center",padding:40,color:"#6B7280"}}>กำลังโหลด...</div>}

        {/* ── ตารางผู้ใช้ ── */}
        <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:"#F9FAFB"}}>
                {["ชื่อ / อีเมล","สถานะ","ระดับที่เข้าได้","จัดการ"].map(h=>(
                  <th key={h} style={{padding:"12px 16px",textAlign:"left",fontWeight:600,color:"#6B7280",fontSize:12}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map(u=>(
                <tr key={u.uid} style={{borderTop:"1px solid #F3F4F6"}}>
                  <td style={{padding:"12px 16px"}}>
                    <div style={{fontWeight:600}}>{u.displayName||"—"}</div>
                    <div style={{fontSize:11,color:"#6B7280"}}>{u.email||u.uid}</div>
                  </td>
                  <td style={{padding:"12px 16px"}}>
                    {u.preAdded
                      ?<span style={{background:"#FEF3C7",color:"#92400E",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>⏳ รอ Login</span>
                      :<span style={{background:"#D1FAE5",color:"#065F46",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>✓ ใช้งานแล้ว</span>
                    }
                  </td>
                  <td style={{padding:"12px 16px"}}>
                    <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                      {Object.entries(u.divisions||{}).filter(([,v])=>v).map(([k])=>(
                        <span key={k} style={{background:"#FEE2E2",color:"#991B1B",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>{divNames[k]||k}</span>
                      ))}
                      {!Object.values(u.divisions||{}).some(Boolean)&&<span style={{color:"#9CA3AF",fontSize:12}}>ไม่มีสิทธิ์</span>}
                    </div>
                  </td>
                  <td style={{padding:"12px 16px"}}>
                    <div style={{display:"flex",gap:6}}>
                      <button
                        onClick={()=>{setEditUid(u.uid);setEditPerms(u.divisions||{p1:false,p2:false,m1:false,m2:false});}}
                        style={{background:"none",border:"1px solid #D1D5DB",borderRadius:8,padding:"4px 12px",cursor:"pointer",fontSize:12}}
                      >แก้ไขสิทธิ์</button>
                      <button
                        onClick={()=>deleteUser(u.uid)}
                        style={{background:"none",border:"1px solid #FECACA",borderRadius:8,padding:"4px 10px",cursor:"pointer",fontSize:12,color:"#DC2626"}}
                        title="ถอนสิทธิ์ทั้งหมด"
                      >✕</button>
                    </div>
                  </td>
                </tr>
              ))}
              {!loading&&!filtered.length&&(
                <tr><td colSpan={4} style={{padding:32,textAlign:"center",color:"#9CA3AF"}}>
                  {users.length===0?"ยังไม่มีผู้ใช้ — เพิ่มอีเมลได้ที่กล่องด้านบน":"ไม่พบผู้ใช้ที่ค้นหา"}
                </td></tr>
              )}
            </tbody>
          </table>
        </div>

        {/* ── Edit permissions modal ── */}
        {editUid&&(
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000}}>
            <div style={{background:"#fff",borderRadius:16,padding:28,width:420}}>
              <h3 style={{fontSize:16,fontWeight:700,marginBottom:4}}>แก้ไขสิทธิ์</h3>
              <p style={{fontSize:12,color:"#6B7280",marginBottom:16}}>{users.find(u=>u.uid===editUid)?.email||editUid}</p>
              <div style={{display:"flex",flexDirection:"column",gap:10,marginBottom:20}}>
                {Object.entries(divNames).map(([k,name])=>(
                  <label key={k} style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer",padding:"10px 14px",borderRadius:10,background:editPerms[k]?"#FEE2E2":"#F9FAFB",border:`1.5px solid ${editPerms[k]?"#DC2626":"#E5E7EB"}`}}>
                    <input type="checkbox" checked={!!editPerms[k]} onChange={e=>setEditPerms(p=>({...p,[k]:e.target.checked}))} style={{width:16,height:16,accentColor:"#DC2626"}}/>
                    <span style={{fontWeight:editPerms[k]?700:400,color:editPerms[k]?"#991B1B":"#374151",fontSize:14}}>{name}</span>
                  </label>
                ))}
              </div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={savePerms} disabled={saving} style={{...BS(),opacity:saving?0.6:1}}>{saving?"กำลังบันทึก...":"บันทึก"}</button>
                <button onClick={()=>{setEditUid(null);setEditPerms({});}} style={BO()}>ยกเลิก</button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

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
const gasGet = async (divId) => {
  const url = divId ? GAS_URL+"?division="+divId : GAS_URL;
  const res = await fetch(url);
  const json = await res.json();
  return json.ok ? json.data : null;
};
const gasPost = async (divId, data) => {
  // GAS ไม่รับ CORS preflight → ใช้ no-cors
  await fetch(GAS_URL, {
    method: "POST",
    mode: "no-cors",
    headers: { "Content-Type": "text/plain" },
    body: JSON.stringify({ action: "save", division: divId, data }),
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

const DIVISIONS=[
  {id:"p1",name:"ประถมศึกษาตอนต้น",short:"ประถมต้น",defaultLevels:["ป.1","ป.2","ป.3"]},
  {id:"p2",name:"ประถมศึกษาตอนปลาย",short:"ประถมปลาย",defaultLevels:["ป.4","ป.5","ป.6"]},
  {id:"m1",name:"มัธยมศึกษาตอนต้น",short:"มัธยมต้น",defaultLevels:["ม.1","ม.2","ม.3"]},
  {id:"m2",name:"มัธยมศึกษาตอนปลาย",short:"มัธยมปลาย",defaultLevels:["ม.4","ม.5","ม.6"]},
];

export default function App() {
  const [page,setPage]=useState("dashboard");
  const [side,setSide]=useState(true);
  const [toast,setToast]=useState(null);
  const [syncing,setSyncing]=useState(false);
  const [gasReady,setGasReady]=useState(false);

  // ===== AUTH STATE =====
  const [authUser,setAuthUser]=useState(undefined);
  const [userPerms,setUserPerms]=useState(null);
  const [showAdmin,setShowAdmin]=useState(false);

  // refresh permissions จาก Firestore (เรียกได้ทุกเวลา)
  const refreshPerms=async(u)=>{
    const user=u||authUser;
    if(!user)return;
    const perms=await fsGetPermissions(user.uid);
    if(perms)setUserPerms(perms);
  };

  // ฟัง Firebase auth state
  useEffect(()=>{
    const {auth}=getFB();
    if(!auth){setAuthUser(null);return;}
    let unsubPerms=null;
    const unsub=onAuthStateChanged(auth,async u=>{
      // ยกเลิก listener เก่า
      if(unsubPerms){unsubPerms();unsubPerms=null;}
      setAuthUser(u||null);
      if(u){
        const {db}=getFB();
        const makePreKey=(email)=>"pre_"+email.trim().toLowerCase().replace(/[@.]/g,"_");

        // โหลด permissions ครั้งแรก
        let perms=await fsGetPermissions(u.uid);

        if(!perms&&db){
          const emailKey=makePreKey(u.email);
          const preSnap=await getDoc(doc(db,"permissions",emailKey));
          if(preSnap.exists()&&!preSnap.data().merged){
            const preData=preSnap.data();
            const divs=preData.divisions||{p1:false,p2:false,m1:false,m2:false};
            await setDoc(doc(db,"permissions",u.uid),{
              displayName:u.displayName||"",email:u.email,divisions:divs,preAdded:false,
            });
            await setDoc(doc(db,"permissions",emailKey),{merged:true},{merge:true});
            perms={divisions:divs};
          }
        }

        if(!perms){
          const emptyDivs={p1:false,p2:false,m1:false,m2:false};
          await fsSetPermissions(u.uid,{displayName:u.displayName||"",email:u.email,divisions:emptyDivs});
          setUserPerms({divisions:emptyDivs});
        } else {
          await fsSetPermissions(u.uid,{displayName:u.displayName||"",email:u.email});
          setUserPerms(perms);
        }

        // Real-time listener — permissions อัปเดตทันทีเมื่อ admin แก้ไข
        if(db){
          unsubPerms=onSnapshot(doc(db,"permissions",u.uid),(snap)=>{
            if(snap.exists())setUserPerms(snap.data());
          });
        }
      } else {
        setUserPerms(null);
      }
    });
    return()=>{unsub();if(unsubPerms)unsubPerms();};
  },[]);

  const handleLogout=async()=>{
    const {auth}=getFB();
    if(auth)await signOut(auth);
  };

  // division state — persist ใน localStorage (ไม่ใช่ per-division key)
  const [divId,setDivId]=useState(()=>localStorage.getItem("dara_division")||"m2");
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

  useEffect(()=>saveLS("academicYear",academicYear),[academicYear]);
  useEffect(()=>saveLS("schoolHeader",schoolHeader),[schoolHeader]);
  // บันทึก division ที่เลือกไว้
  useEffect(()=>{ localStorage.setItem("dara_division",divId); },[divId]);

  const stateRef=useRef({});
  useEffect(()=>{stateRef.current={levels,plans,depts,teachers,subjects,rooms,specialRooms,assigns,meetings,schedule,locks}},[levels,plans,depts,teachers,subjects,rooms,specialRooms,assigns,meetings,schedule,locks]);

  // เมื่อ switch division → โหลดข้อมูลชุดใหม่
  const switchDivision=(newDivId)=>{
    const d=DIVISIONS.find(x=>x.id===newDivId);
    if(!d) return;
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
    setPage("dashboard");
    st("เปลี่ยนเป็น "+d.name);
  };

  const saveTimer=useRef(null);
  const syncToGas=useCallback(()=>{
    if(!GAS_URL||GAS_URL.includes("YOUR_DEPLOYMENT_ID"))return;
    clearTimeout(saveTimer.current);
    saveTimer.current=setTimeout(()=>{
      setSyncing(true);
      gasPost(divId, stateRef.current).catch(()=>{}).finally(()=>setSyncing(false));
    },1500);
  },[divId]);

  // โหลดจาก GAS ตอนเริ่ม — ทับ localStorage เฉพาะเมื่อ GAS มีข้อมูลจริง
  useEffect(()=>{
    if(!GAS_URL||GAS_URL.includes("YOUR_DEPLOYMENT_ID"))return;
    setSyncing(true);
    gasGet(divId).then(d=>{
      if(d){
        // ทับเฉพาะ field ที่มีข้อมูลจริง (array ไม่ว่าง หรือ object ไม่ว่าง)
        if(d.levels?.length)       setLevels(d.levels);
        if(d.plans?.length)        setPlans(d.plans);
        if(d.depts?.length)        setDepts(d.depts);
        if(d.teachers?.length)     setTeachers(d.teachers);
        if(d.subjects?.length)     setSubjects(d.subjects);
        if(d.rooms?.length)        setRooms(d.rooms);
        if(d.specialRooms?.length) setSpecialRooms(d.specialRooms);
        if(d.assigns?.length)      setAssigns(d.assigns);
        if(d.meetings?.length)     setMeetings(d.meetings);
        if(d.schedule&&Object.keys(d.schedule).length) setSchedule(d.schedule);
        if(d.locks&&Object.keys(d.locks).length)       setLocks(d.locks);
        setGasReady(true);
      } else { setGasReady(true); }
    }).catch(()=>{ setGasReady(true); }).finally(()=>setSyncing(false));
  },[divId]);

  // Auto-save ไป localStorage + GAS เมื่อข้อมูลเปลี่ยน (per-division key)
  useEffect(()=>{ saveLS(divId+"_levels",levels);   if(gasReady) syncToGas(); },[levels,gasReady]);
  useEffect(()=>{ saveLS(divId+"_plans",plans);     if(gasReady) syncToGas(); },[plans,gasReady]);
  useEffect(()=>{ saveLS(divId+"_depts",depts);     if(gasReady) syncToGas(); },[depts,gasReady]);
  useEffect(()=>{ saveLS(divId+"_teachers",teachers); if(gasReady) syncToGas(); },[teachers,gasReady]);
  useEffect(()=>{ saveLS(divId+"_subjects",subjects); if(gasReady) syncToGas(); },[subjects,gasReady]);
  useEffect(()=>{ saveLS(divId+"_rooms",rooms);     if(gasReady) syncToGas(); },[rooms,gasReady]);
  useEffect(()=>{ saveLS(divId+"_specialRooms",specialRooms); if(gasReady) syncToGas(); },[specialRooms,gasReady]);
  useEffect(()=>{ saveLS(divId+"_assigns",assigns); if(gasReady) syncToGas(); },[assigns,gasReady]);
  useEffect(()=>{ saveLS(divId+"_meetings",meetings); if(gasReady) syncToGas(); },[meetings,gasReady]);
  useEffect(()=>{ saveLS(divId+"_schedule",schedule); if(gasReady) syncToGas(); },[schedule,gasReady]);
  useEffect(()=>{ saveLS(divId+"_locks",locks);     if(gasReady) syncToGas(); },[locks,gasReady]);

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
    {id:"meetings",icon:"clock",label:"คาบล็อค / ประชุม"},
    {id:"scheduler",icon:"grid",label:"จัดตารางสอน"},
    {id:"reports",icon:"download",label:"รายงาน / Export"},
    {id:"settings",icon:"file",label:"ตั้งค่า / ปีการศึกษา"},
  ];
  const S={levels,plans,depts,teachers,subjects,rooms,specialRooms,assigns,meetings,schedule,locks};
  const U={setLevels,setPlans,setDepts,setTeachers,setSubjects,setRooms,setSpecialRooms,setAssigns,setMeetings,setSchedule,setLocks};

  // ===== AUTH GUARDS =====
  const firebaseConfigured=!FIREBASE_CONFIG.apiKey.includes("YOUR");

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
  if(showAdmin){
    return <AdminPanel user={authUser} onBack={()=>{setShowAdmin(false);refreshPerms();}} refreshPerms={()=>refreshPerms()}/>;
  }

  // Logged in but no permission for this division
  const allowedDivs=firebaseConfigured
    ?DIVISIONS.filter(d=>userPerms?.divisions?.[d.id]!==false) // ถ้าไม่มี key → อนุญาต (backward compat)
    :DIVISIONS;

  // Filter division selector ตาม permissions
  const availDivs=firebaseConfigured
    ?DIVISIONS.filter(d=>userPerms?.divisions?.[d.id]===true)
    :DIVISIONS;

  const divHasAccess=!firebaseConfigured||userPerms?.divisions?.[divId]===true;

  return <div style={{display:"flex",height:"100vh",fontFamily:"'Sarabun','Noto Sans Thai',sans-serif",background:"#F3F4F6",overflow:"hidden"}}>
    <style>{`@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700;800&display=swap');*{box-sizing:border-box;margin:0;padding:0}::-webkit-scrollbar{width:6px}::-webkit-scrollbar-thumb{background:#CBD5E1;border-radius:3px}@keyframes slideIn{from{transform:translateX(100px);opacity:0}to{transform:translateX(0);opacity:1}}@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}.ni:hover{background:rgba(255,255,255,0.15)!important}.ni.a{background:rgba(255,255,255,0.2)!important}input:focus,select:focus{border-color:#DC2626!important;box-shadow:0 0 0 3px rgba(220,38,38,0.1)!important}.drag-card{cursor:grab;user-select:none}.drag-card:active{cursor:grabbing}.dz{transition:background 0.2s}.dz.over{background:#FEE2E2!important;outline:2px dashed #DC2626}button:hover{opacity:0.85}select{appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%236B7280' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 12px center;padding-right:36px!important}.div-sel{appearance:none!important;background:#00000030!important;background-image:none!important;border:1px solid rgba(255,255,255,0.3)!important;border-radius:8px!important;color:#fff!important;font-size:13px!important;font-weight:700!important;font-family:inherit!important;padding:9px 32px 9px 12px!important;width:100%!important;cursor:pointer!important;outline:none!important}.div-sel:focus{box-shadow:0 0 0 2px rgba(255,255,255,0.3)!important;border-color:rgba(255,255,255,0.6)!important}.div-sel option{background:#7F1D1D;color:#fff}`}</style>

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
      {/* Division selector — dropdown */}
      <div style={{padding:"10px 12px",borderBottom:"1px solid rgba(255,255,255,0.1)"}}>
        <div style={{fontSize:10,color:"rgba(255,255,255,0.45)",marginBottom:5,paddingLeft:2}}>ระดับการศึกษา</div>
        <div style={{position:"relative"}}>
          <select className="div-sel" value={divId} onChange={e=>switchDivision(e.target.value)}>
            {(firebaseConfigured?availDivs:DIVISIONS).map(d=><option key={d.id} value={d.id}>{d.name}</option>)}
          </select>
          <div style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",pointerEvents:"none",color:"rgba(255,255,255,0.7)"}}>
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="6 9 12 15 18 9"/></svg>
          </div>
        </div>
      </div>
      <nav style={{flex:1,padding:"12px 10px",overflowY:"auto"}}>
        {nav.map(n=><div key={n.id} className={`ni ${page===n.id?"a":""}`} onClick={()=>setPage(n.id)} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 14px",borderRadius:10,cursor:"pointer",color:page===n.id?"#fff":"rgba(255,255,255,0.7)",fontSize:14,fontWeight:page===n.id?700:400,marginBottom:2}}><Icon name={n.icon} size={18}/>{n.label}</div>)}
      </nav>
      <div style={{padding:"14px 16px",borderTop:"1px solid rgba(255,255,255,0.1)"}}>
        {/* User info */}
        {firebaseConfigured&&authUser&&<div style={{marginBottom:10}}>
          <div style={{color:"rgba(255,255,255,0.85)",fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{authUser.displayName||authUser.email}</div>
          <div style={{color:"rgba(255,255,255,0.45)",fontSize:10,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{authUser.email}</div>
        </div>}
        <div style={{display:"flex",gap:6,marginBottom:8}}>
          {firebaseConfigured&&<button onClick={handleLogout} style={{flex:1,padding:"6px 0",background:"rgba(255,255,255,0.1)",border:"1px solid rgba(255,255,255,0.2)",borderRadius:8,color:"rgba(255,255,255,0.7)",fontSize:11,fontWeight:600,cursor:"pointer"}}>ออกจากระบบ</button>}
          {firebaseConfigured&&<button onClick={()=>setShowAdmin(true)} style={{flex:1,padding:"6px 0",background:"rgba(255,255,255,0.1)",border:"1px solid rgba(255,255,255,0.2)",borderRadius:8,color:"rgba(255,255,255,0.7)",fontSize:11,fontWeight:600,cursor:"pointer"}}>🔐 Admin</button>}
        </div>
        <div style={{color:"rgba(255,255,255,0.4)",fontSize:10}}>ผู้พัฒนา</div>
        <div style={{color:"rgba(255,255,255,0.65)",fontSize:11,fontWeight:600,marginTop:1}}>พนิต เกิดมงคล</div>
      </div>
    </div>

    <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
      <header style={{height:60,background:"#fff",borderBottom:"1px solid #E5E7EB",display:"flex",alignItems:"center",padding:"0 24px",gap:16,flexShrink:0}}>
        <button onClick={()=>setSide(!side)} style={{background:"none",border:"none",cursor:"pointer",color:"#6B7280",padding:4}}><Icon name="menu" size={22}/></button>
        <h2 style={{fontSize:18,fontWeight:700}}>{nav.find(n=>n.id===page)?.label}</h2>
        <span style={{fontSize:12,background:"#FEE2E2",color:"#991B1B",padding:"3px 10px",borderRadius:20,fontWeight:600}}>{div.short}</span>
        <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:10}}>
          {GAS_URL&&!GAS_URL.includes("YOUR_DEPLOYMENT_ID")
            ?syncing
              ?<span style={{fontSize:12,color:"#D97706",display:"flex",alignItems:"center",gap:4}}>⏳ กำลัง sync...</span>
              :gasReady
                ?<span style={{fontSize:12,color:"#059669",display:"flex",alignItems:"center",gap:4}}>☁️ sync แล้ว</span>
                :null
            :<span style={{fontSize:12,color:"#9CA3AF",display:"flex",alignItems:"center",gap:4}}>💾 local only</span>
          }
          {firebaseConfigured&&authUser?.photoURL&&(
            <img src={authUser.photoURL} alt="avatar" style={{width:32,height:32,borderRadius:"50%",objectFit:"cover",border:"2px solid #E5E7EB"}}/>
          )}
        </div>
      </header>
      <main style={{flex:1,overflow:"auto",padding:24}}>
        {/* No access guard */}
        {firebaseConfigured&&!divHasAccess
          ?<div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:"100%",gap:16}}>
              <div style={{fontSize:48}}>🔒</div>
              <h2 style={{fontSize:20,fontWeight:700,color:"#374151"}}>ไม่มีสิทธิ์เข้าระดับนี้</h2>
              <p style={{color:"#6B7280",fontSize:14}}>กรุณาติดต่อผู้ดูแลระบบเพื่อขอสิทธิ์ {div.name}</p>
            </div>
          :<>
            {page==="dashboard"&&<Dash S={S} setPage={setPage}/>}
            {page==="levels"&&<Levels S={S} U={U} st={st}/>}
            {page==="plans"&&<Plans S={S} U={U} st={st}/>}
            {page==="departments"&&<Depts S={S} U={U} st={st} gc={gc}/>}
            {page==="teachers"&&<Teachers S={S} U={U} st={st} gc={gc}/>}
            {page==="subjects"&&<Subjects S={S} U={U} st={st} gc={gc}/>}
            {page==="specialrooms"&&<SpecialRooms S={S} U={U} st={st}/>}
            {page==="assignments"&&<Assigns S={S} U={U} st={st} gc={gc}/>}
            {page==="meetings"&&<Meetings S={S} U={U} st={st} gc={gc}/>}
            {page==="scheduler"&&<Scheduler S={S} U={U} st={st} gc={gc}/>}
            {page==="reports"&&<Reports S={S} st={st} gc={gc} ay={academicYear} sh={schoolHeader}/>}
            {page==="settings"&&<Settings S={S} U={U} st={st} ay={academicYear} setAY={setAcademicYear} sh={schoolHeader} setSH={setSchoolHeader} div={div}/>}
          </>
        }
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

  // ข้อ 5: นับคาบรวมทั้ง assign ปกติ + คาบที่เป็นครูร่วมในตาราง
  const usedPeriods=(tid)=>{
    let u=0;
    S.assigns.filter(a=>a.teacherId===tid).forEach(a=>{u+=a.totalPeriods||0});
    // นับคาบครูร่วม (coTeacherId) จากตารางที่จัดไปแล้ว
    Object.values(S.schedule).forEach(en=>{en?.forEach(e=>{if(e.coTeacherId===tid)u++})});
    return u;
  };

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

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"flex",gap:10,marginBottom:20,flexWrap:"wrap",alignItems:"center"}}>
      <button onClick={()=>{setEditId(null);setForm({name:"",capacity:0,note:""});setModal(true)}} style={BS()}><Icon name="plus" size={16}/>เพิ่มห้องพิเศษ</button>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(260px,1fr))",gap:16}}>
      {S.specialRooms.map(r=>{const sc=subCount(r.id);return<div key={r.id} style={{background:"#fff",borderRadius:14,padding:18,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",borderLeft:"4px solid #7C3AED"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div>
            <h4 style={{fontSize:15,fontWeight:700}}>{r.name}</h4>
            {r.note&&<div style={{fontSize:12,color:"#6B7280",marginTop:2}}>{r.note}</div>}
          </div>
          <div style={{display:"flex",gap:6}}>
            <button onClick={()=>openEdit(r)} style={{background:"none",border:"none",cursor:"pointer",color:"#2563EB"}}><Icon name="edit" size={14}/></button>
            <button onClick={()=>{if(sc>0){st("มีวิชาใช้ห้องนี้อยู่ "+sc+" วิชา ลบไม่ได้","error");return}U.setSpecialRooms(p=>p.filter(x=>x.id!==r.id));st("ลบแล้ว","warning")}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444"}}><Icon name="trash" size={14}/></button>
          </div>
        </div>
        <div style={{display:"flex",gap:8,marginTop:10,flexWrap:"wrap"}}>
          {r.capacity>0&&<span style={{background:"#EDE9FE",color:"#5B21B6",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>ความจุ {r.capacity} คน</span>}
          <span style={{background:"#F3F4F6",color:"#374151",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>{sc} วิชาใช้ห้องนี้</span>
        </div>
      </div>})}
      {!S.specialRooms.length&&<div style={{padding:40,textAlign:"center",color:"#9CA3AF",gridColumn:"1/-1"}}>ยังไม่มีห้องพิเศษ — เพิ่มได้เลย เช่น ห้องคอมพิวเตอร์ ห้องแลบ ห้องประกอบอาหาร</div>}
    </div>
    <Modal open={modal} onClose={()=>{setModal(false);setEditId(null)}} title={editId?"แก้ไขห้องพิเศษ":"เพิ่มห้องพิเศษ"}>
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <div><label style={LS}>ชื่อห้อง</label><input style={IS} value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))} placeholder="เช่น ห้องคอมพิวเตอร์ 1, ห้องแลบวิทย์"/></div>
        <div><label style={LS}>ความจุ (คน) — ไม่บังคับ</label><input type="number" min="0" style={IS} value={form.capacity} onChange={e=>setForm(p=>({...p,capacity:parseInt(e.target.value)||0}))}/></div>
        <div><label style={LS}>หมายเหตุ</label><input style={IS} value={form.note} onChange={e=>setForm(p=>({...p,note:e.target.value}))} placeholder="รายละเอียดเพิ่มเติม"/></div>
        <button onClick={save} style={BS()}>{editId?"บันทึก":"เพิ่มห้องพิเศษ"}</button>
      </div>
    </Modal>
  </div>;
}

/* ===== SUBJECTS ===== */
function Subjects({S,U,st,gc}){
  const [modal,setModal]=useState(false);
  const [editId,setEditId]=useState(null);
  const BLANK={code:"",name:"",credits:1,periodsPerWeek:1,departmentId:"",levelId:"",specialRoomId:"",consecutiveAllowed:0};
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
    setForm({code:s.code||"",name:s.name||"",credits:s.credits||1,periodsPerWeek:s.periodsPerWeek||1,
      departmentId:s.departmentId||"",levelId:s.levelId||"",
      specialRoomId:s.specialRoomId||"",consecutiveAllowed:s.consecutiveAllowed||0});
    setModal(true);
  };

  const handleFile=async(e)=>{const f=e.target.files?.[0];if(!f)return;
    let rows;
    if(f.name.endsWith('.csv')){const txt=await f.text();rows=parseCSV(txt)}
    else{rows=await readExcelFile(f)}
    if(!rows||!rows.length){st("ไม่พบข้อมูล","error");return}
    const ns=rows.map(r=>{const dept=S.depts.find(d=>d.name===String(r["กลุ่มสาระ"]||"").trim());const lv=S.levels.find(l=>l.name===String(r["ระดับชั้น"]||"").trim());
      return{id:gid(),code:String(r["รหัสวิชา"]||"").trim(),name:String(r["ชื่อวิชา"]||"").trim(),credits:parseFloat(r["หน่วยกิต"])||1,periodsPerWeek:parseInt(r["คาบ/สัปดาห์"])||1,departmentId:dept?.id||"",levelId:lv?.id||"",specialRoomId:"",consecutiveAllowed:0}
    }).filter(s=>s.name);
    U.setSubjects(p=>[...p,...ns]);st(`นำเข้า ${ns.length} วิชา`);e.target.value=""};

  const exportS=()=>{exportExcel(["รหัสวิชา","ชื่อวิชา","หน่วยกิต","คาบ/สัปดาห์","กลุ่มสาระ","ระดับชั้น"],S.subjects.map(s=>[s.code,s.name,s.credits,s.periodsPerWeek,S.depts.find(d=>d.id===s.departmentId)?.name||"",S.levels.find(l=>l.id===s.levelId)?.name||""]),"รายวิชา_ดาราวิทยาลัย.xlsx","วิชา");st("Export สำเร็จ")};
  const downloadTemplate=()=>{exportExcel(["รหัสวิชา","ชื่อวิชา","หน่วยกิต","คาบ/สัปดาห์","กลุ่มสาระ","ระดับชั้น"],[["ว33201","ฟิสิกส์ 3",1.5,3,"วิทยาศาสตร์","ม.6"]],"Template_วิชา.xlsx","Template");st("ดาวน์โหลด Template")};

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
    return<div style={{background:"#fff",borderRadius:12,overflow:"hidden",boxShadow:"0 1px 3px rgba(0,0,0,0.06)",borderLeft:"3px solid "+c.bg}}>
      <div style={{padding:"12px 14px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:10,color:"#9CA3AF",fontWeight:600}}>{sub.code}</div>
            <h4 style={{fontSize:14,fontWeight:700,marginTop:1,wordBreak:"break-word"}}>{sub.name}</h4>
          </div>
          <div style={{display:"flex",gap:4,flexShrink:0,marginLeft:8}}>
            <button onClick={()=>openEdit(sub)} style={{background:"none",border:"none",cursor:"pointer",color:"#2563EB"}}><Icon name="edit" size={13}/></button>
            <button onClick={()=>{U.setSubjects(p=>p.filter(x=>x.id!==sub.id));st("ลบแล้ว","warning")}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444"}}><Icon name="trash" size={13}/></button>
          </div>
        </div>
        <div style={{display:"flex",gap:4,marginTop:8,flexWrap:"wrap"}}>
          <span style={{background:"#F3F4F6",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>{sub.credits} หน่วยกิต</span>
          <span style={{background:"#F3F4F6",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>{sub.periodsPerWeek} คาบ/สป.</span>
          {sr&&<span style={{background:"#EDE9FE",color:"#5B21B6",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>📍{sr.name}</span>}
          {sub.consecutiveAllowed>0&&<span style={{background:"#FEF3C7",color:"#92400E",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>⚡{sub.consecutiveAllowed}คาบติด</span>}
          {sub.consecutiveAllowed===-1&&<span style={{background:"#EFF6FF",color:"#1E40AF",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>🔀NP</span>}
          {sub.consecutiveAllowed===-2&&<span style={{background:"#FDF4FF",color:"#6B21A8",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>🏛️เศรษฐ-วิศวะ</span>}
        </div>
      </div>
    </div>;
  };

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
      <button onClick={()=>{setEditId(null);setForm(BLANK);setModal(true)}} style={BS()}><Icon name="plus" size={16}/>เพิ่มวิชา</button>
      <button onClick={()=>fileRef.current?.click()} style={BS("#2563EB")}><Icon name="upload" size={16}/>Import Excel</button>
      <button onClick={downloadTemplate} style={BO("#2563EB")}><Icon name="file" size={16}/>Template</button>
      <button onClick={exportS} style={BO("#059669")}><Icon name="download" size={16}/>Export Excel</button>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleFile}/>
    </div>
    {/* Filters */}
    <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
      <select style={{...IS,maxWidth:140}} value={filterLv} onChange={e=>setFilterLv(e.target.value)}>
        <option value="">ทุกระดับ</option>{S.levels.map(l=><option key={l.id} value={l.id}>{l.name}</option>)}
      </select>
      <select style={{...IS,maxWidth:180}} value={filterDept} onChange={e=>setFilterDept(e.target.value)}>
        <option value="">ทุกกลุ่มสาระ</option>{S.depts.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}
      </select>
      <div style={{position:"relative",flex:"1 1 180px",maxWidth:280}}>
        <input style={{...IS,paddingLeft:34}} value={search} onChange={e=>setSearch(e.target.value)} placeholder="ค้นหาวิชา..."/>
        <div style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:"#9CA3AF"}}><Icon name="search" size={14}/></div>
      </div>
      <span style={{fontSize:12,color:"#9CA3AF"}}>{filtered.length} วิชา</span>
    </div>

    {/* Grouped display */}
    {groups.map(({lv,deptGroups})=><div key={lv.id} style={{marginBottom:24}}>
      <div style={{background:"linear-gradient(135deg,#991B1B,#DC2626)",borderRadius:10,padding:"10px 16px",marginBottom:12,display:"flex",alignItems:"center",gap:8}}>
        <span style={{color:"#fff",fontSize:16,fontWeight:700}}>{lv.name}</span>
        <span style={{background:"rgba(255,255,255,0.2)",color:"#fff",fontSize:11,padding:"2px 8px",borderRadius:20}}>{filtered.filter(s=>s.levelId===lv.id).length} วิชา</span>
      </div>
      {deptGroups.map(({dept,subs})=><div key={dept?.id||"none"} style={{marginBottom:16}}>
        <div style={{fontSize:12,fontWeight:700,color:"#6B7280",marginBottom:8,paddingLeft:4,display:"flex",alignItems:"center",gap:6}}>
          {dept&&<span style={{width:10,height:10,borderRadius:"50%",background:gc(dept.id).bg,display:"inline-block"}}/>}
          {dept?.name||"ไม่ระบุกลุ่มสาระ"}
          <span style={{fontWeight:400}}>({subs.length})</span>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(250px,1fr))",gap:10}}>
          {subs.map(sub=><SubCard key={sub.id} sub={sub}/>)}
        </div>
      </div>)}
    </div>)}
    {noLevel.length>0&&<div style={{marginBottom:24}}>
      <div style={{background:"#F3F4F6",borderRadius:10,padding:"10px 16px",marginBottom:12}}><span style={{fontSize:14,fontWeight:700,color:"#6B7280"}}>ไม่ระบุระดับชั้น ({noLevel.length} วิชา)</span></div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(250px,1fr))",gap:10}}>
        {noLevel.map(sub=><SubCard key={sub.id} sub={sub}/>)}
      </div>
    </div>}
    {!filtered.length&&<div style={{padding:40,textAlign:"center",color:"#9CA3AF"}}>ยังไม่มีวิชา</div>}

    <Modal open={modal} onClose={()=>{setModal(false);setEditId(null)}} title={editId?"แก้ไขวิชา":"เพิ่มวิชา"}>
      <div style={{display:"flex",flexDirection:"column",gap:14}}>
        <div><label style={LS}>รหัสวิชา</label><input style={IS} value={form.code} onChange={e=>setForm(p=>({...p,code:e.target.value}))} placeholder="ว33202"/></div>
        <div><label style={LS}>ชื่อวิชา</label><input style={IS} value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))} placeholder="ฟิสิกส์ 4"/></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          <div><label style={LS}>หน่วยกิต</label><input type="number" min="0.5" step="0.5" style={IS} value={form.credits} onChange={e=>setForm(p=>({...p,credits:parseFloat(e.target.value)||0}))}/></div>
          <div><label style={LS}>คาบ/สัปดาห์</label><input type="number" min="1" style={IS} value={form.periodsPerWeek} onChange={e=>setForm(p=>({...p,periodsPerWeek:parseInt(e.target.value)||1}))}/></div>
        </div>
        <div><label style={LS}>ระดับชั้น</label><select style={IS} value={form.levelId} onChange={e=>setForm(p=>({...p,levelId:e.target.value}))}><option value="">--</option>{S.levels.map(l=><option key={l.id} value={l.id}>{l.name}</option>)}</select></div>
        <div><label style={LS}>กลุ่มสาระ</label><select style={IS} value={form.departmentId} onChange={e=>setForm(p=>({...p,departmentId:e.target.value}))}><option value="">--</option>{S.depts.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}</select></div>
        <div><label style={LS}>ห้องพิเศษ (ถ้าต้องใช้) — ตรวจ conflict ข้ามทุกห้อง</label>
          <select style={IS} value={form.specialRoomId} onChange={e=>setForm(p=>({...p,specialRoomId:e.target.value}))}>
            <option value="">-- ไม่ใช้ห้องพิเศษ --</option>
            {S.specialRooms.map(r=><option key={r.id} value={r.id}>{r.name}</option>)}
          </select>
        </div>
        <div><label style={LS}>คาบติดต่อกัน / คาบพิเศษ</label>
          <select style={IS} value={form.consecutiveAllowed} onChange={e=>setForm(p=>({...p,consecutiveAllowed:parseInt(e.target.value)||0}))}>
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

  // ข้อ 2: filter วิชาเฉพาะกลุ่มสาระของครูที่เลือก
  const teacherDeptSubs=teacher?S.subjects.filter(s=>s.departmentId===teacher.departmentId):S.subjects;

  // ข้อ 3: when subject selected, show rooms of that level only
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
        <div><label style={LS}>วิชา (รหัส — ชื่อ) — เฉพาะกลุ่มสาระ{teacher?(" "+S.depts.find(d=>d.id===teacher.departmentId)?.name):""}</label><select style={IS} value={form.subjectId} onChange={e=>{setForm(p=>({...p,subjectId:e.target.value,roomIds:[]}))}}><option value="">--</option>{teacherDeptSubs.map(s=><option key={s.id} value={s.id}>{s.code} — {s.name} ({S.levels.find(l=>l.id===s.levelId)?.name})</option>)}</select></div>
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

/* ===== EMPTY STATE HELPER ===== */
function EmptyState({icon,title}){
  return <div style={{background:"#fff",borderRadius:14,padding:60,textAlign:"center"}}>
    <div style={{fontSize:48,marginBottom:16}}>{icon}</div>
    <h3 style={{fontSize:18,fontWeight:700,color:"#374151"}}>{title}</h3>
  </div>;
}

/* ===== SCHEDULER ENTRY CARD (top-level เพื่อกัน React recreate) ===== */
function SchedulerEntryCard({entry,cellKey,lk,cellCount,selT,mode,S,U,gc,setDrag,setCoM}){
  const sub=S.subjects.find(s=>s.id===entry.subjectId);
  const dept=S.depts.find(d=>d.id===sub?.departmentId);
  const c=dept?gc(dept.id):{bg:"#6B7280",lt:"#F3F4F6",tx:"#374151",bd:"#D1D5DB"};
  const et=S.teachers.find(t=>t.id===entry.teacherId);
  const ct=entry.coTeacherId?S.teachers.find(t=>t.id===entry.coTeacherId):null;
  const isOwn=entry.teacherId===selT||entry.coTeacherId===selT;
  const dimmed=mode==="teacher"&&!!selT&&!isOwn;
  const compact=cellCount>1;

  const removeEntry=()=>U.setSchedule(prev=>({...prev,[cellKey]:(prev[cellKey]||[]).filter(e=>e.id!==entry.id)}));
  const lockEntry=()=>U.setLocks(prev=>({...prev,[cellKey]:true}));
  const unlockEntry=()=>U.setLocks(prev=>({...prev,[cellKey]:false}));

  return (
    <div
      draggable={!lk&&!dimmed}
      onDragStart={e=>{if(dimmed){e.preventDefault();return;}e.stopPropagation();const parts=cellKey.split('_');const fromRoomId=parts.slice(0,parts.length-2).join('_');setDrag({fromKey:cellKey,fromRoomId,entry});}}
      onDragEnd={()=>setDrag(null)}
      style={{
        background:dimmed?"#F9FAFB":c.lt,
        border:"2px solid "+(dimmed?"#E5E7EB":c.bd),
        borderRadius:6,
        padding:compact?"3px 5px":"5px 7px",
        marginBottom:2,
        fontSize:11,
        position:"relative",
        cursor:lk||dimmed?"default":"grab",
        opacity:dimmed?0.45:1,
        transition:"opacity 0.15s",
        userSelect:"none",
      }}
    >
      {compact
        ?<div style={{fontWeight:700,color:dimmed?"#9CA3AF":c.tx,fontSize:10,lineHeight:1.3,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
            {sub?.name||sub?.code}
          </div>
        :<>
            <div style={{fontWeight:700,color:dimmed?"#9CA3AF":c.tx,fontSize:11}}>{sub?.code}</div>
            <div style={{fontWeight:600,color:dimmed?"#9CA3AF":c.tx,fontSize:10}}>{sub?.name}</div>
            <div style={{color:dimmed?"#9CA3AF":c.tx,opacity:0.7,fontSize:10}}>
              {et?.firstName}{ct?" + "+ct.firstName:""}
            </div>
          </>
      }
      {/* action buttons */}
      {!lk&&!compact&&(
        <div style={{display:"flex",gap:3,marginTop:3}}>
          <button onClick={removeEntry} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",padding:0,lineHeight:1}}><Icon name="x" size={10}/></button>
          <button onClick={()=>setCoM({key:cellKey,entryId:entry.id})} style={{background:"none",border:"none",cursor:"pointer",color:"#2563EB",padding:0,lineHeight:1}}><Icon name="users" size={10}/></button>
          <button onClick={lockEntry} style={{background:"none",border:"none",cursor:"pointer",color:"#059669",padding:0,lineHeight:1}}><Icon name="lock" size={10}/></button>
        </div>
      )}
      {!lk&&compact&&(
        <button onClick={removeEntry} style={{position:"absolute",top:1,right:1,background:"none",border:"none",cursor:"pointer",color:"#EF4444",padding:0,lineHeight:1}}><Icon name="x" size={9}/></button>
      )}
      {lk&&(
        <div style={{position:"absolute",top:2,right:4}}>
          <button onClick={unlockEntry} style={{background:"none",border:"none",cursor:"pointer",color:"#059669",padding:0,lineHeight:1}}><Icon name="unlock" size={10}/></button>
        </div>
      )}
    </div>
  );
}

/* ===== SCHEDULER ===== */
function Scheduler({S,U,st,gc}){
  const [mode,setMode]=useState("teacher");
  const [selDept,setSelDept]=useState("");
  const [selT,setSelT]=useState("");
  const [selRoom,setSelRoom]=useState("");
  const [drag,setDrag]=useState(null);
  const dragRef=useRef(null);  // ref สำหรับอ่านใน handleDrop กัน stale/race condition
  const setDragBoth=(v)=>{setDrag(v);dragRef.current=v;};
  const [coM,setCoM]=useState(null);   // {key, entryId} — modal บนการ์ดที่วางแล้ว
  const [coS,setCoS]=useState("");
  const [coDept,setCoDept]=useState("");
  const [cardCoM,setCardCoM]=useState(null); // assignId — modal บน sidebar card
  const [cardCoS,setCardCoS]=useState("");
  const [cardCoDept,setCardCoDept]=useState("");
  const [cardCoMap,setCardCoMap]=useState({}); // {assignId: teacherId}

  const teacher  = S.teachers.find(t=>t.id===selT);
  const asgns    = S.assigns.filter(a=>a.teacherId===selT);
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
  // tRooms: ห้องของครูที่เลือก เรียงตาม sortedRooms (ม.4→ม.5→ม.6, เลขห้องน้อย→มาก)
  const tRoomsSet = new Set(asgns.flatMap(a=>a.roomIds));
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
    S.meetings.filter(m=>m.departmentId===t.departmentId)
      .forEach(m=>m.periods.forEach(p=>b.push({day:m.day,period:p,reason:"ประชุม"})));
    return b;
  },[S.teachers,S.meetings]);

  const isBlk=(tid,day,p)=>blocked(tid).some(b=>b.day===day&&b.period===p);
  const sk=(rid,day,p)=>rid+"_"+day+"_"+p;

  const teacherBusy=(tid,day,period,excludeKey,newSubjectId=null)=>{
    for(const [k,en] of Object.entries(S.schedule)){
      if(k===excludeKey)continue;
      if(!k.endsWith("_"+day+"_"+period))continue;
      if(en?.some(e=>{
        if(e.teacherId!==tid&&e.coTeacherId!==tid)return false;
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
    // NP (-1): อนุญาตคนละห้อง แต่ห้ามซ้ำห้องเดิมในวันเดิว (max 1/ห้อง/วัน)
    if(allowed===-1){
      let c=0;
      for(const [k,en] of Object.entries(S.schedule)){
        if(k===excludeKey)continue;
        const pts=k.split("_");
        if(pts[0]!==roomId||pts[1]!==day)continue;
        en?.forEach(e=>{if(e.subjectId===subjectId)c++;});
      }
      return c>=1;
    }
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
        if(e.teacherId===tid||e.coTeacherId===tid){
          const sub=S.subjects.find(s=>s.id===e.subjectId);
          const ca=sub?.consecutiveAllowed||0;
          if(ca===-1||ca===-2){
            // NP/-2: deduplicate ด้วย subjectId_day_period (ไม่นับซ้ำข้ามห้อง)
            const npKey=e.subjectId+"_"+pts[1]+"_"+pts[2];
            if(!seen.has(npKey)){seen.add(npKey);c++;}
          } else {
            c++;
          }
        }
      });
    });
    return c;
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
      const room=S.rooms.find(r=>r.id===rid);
      if(room&&sub&&room.levelId!==sub.levelId){st("ระดับชั้นไม่ตรงกัน!","error");return;}
      if(specialRoomBusy(entry.subjectId,day,p,drag.fromKey)){
        const sr=S.specialRooms.find(r=>r.id===sub?.specialRoomId);
        st("ห้องพิเศษ '"+(sr?.name||"")+"' ถูกใช้อยู่","error");return;
      }
      // ตรวจ teacher conflict เฉพาะ teacher-mode
      if(selT){
        if(isBlk(entry.teacherId,day,p)){st("ครูถูกล็อคคาบนี้","error");return;}
        if(teacherBusy(entry.teacherId,day,p,drag.fromKey)){st("ครูคนนี้สอนคาบนี้อยู่แล้ว","error");return;}
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
    const room=S.rooms.find(r=>r.id===rid);
    // ห้ามวางในห้องที่ไม่ได้อยู่ใน assignment
    const asgn=S.assigns.find(a=>a.id===drag.assignmentId);
    if(!asgn?.roomIds?.includes(rid)){st("ห้องนี้ไม่ได้รับมอบหมายวิชานี้!","error");setDragBoth(null);return;}
    if(isBlk(drag.teacherId,day,p)){st("ครูถูกล็อคคาบนี้","error");return;}
    if(teacherBusy(drag.teacherId,day,p,null,drag.subjectId)){st("ครูคนนี้สอนคาบนี้อยู่แล้ว (ห้องอื่น)","error");return;}
    if(specialRoomBusy(drag.subjectId,day,p,null)){
      const sr=S.specialRooms.find(r=>r.id===sub?.specialRoomId);
      st("ห้องพิเศษ '"+(sr?.name||"")+"' ถูกใช้อยู่แล้วในคาบนี้","error");return;
    }
    if(room&&sub&&room.levelId!==sub.levelId){st("ระดับชั้นไม่ตรงกัน!","error");return;}
    if(sameSubjectSameDay(drag.subjectId,rid,day,null)){st("วิชานี้มีในวัน"+day+"แล้ว (ห้ามซ้ำ/วัน)","error");return;}
    const placed=countSubjectInRoom(drag.assignmentId,rid);
    const limit=getPerRoomLimit(drag.assignmentId);
    if(placed>=limit){st("ห้องนี้ลงครบ "+limit+" คาบแล้ว","error");return;}
    const coTid=cardCoMap[drag.assignmentId]||null;
    U.setSchedule(prev=>({
      ...prev,
      [key]:[...(prev[key]||[]),{id:gid(),teacherId:drag.teacherId,subjectId:drag.subjectId,assignmentId:drag.assignmentId,coTeacherId:coTid}]
    }));
    setDragBoth(null);
  };

  /* ── co-teacher dept+teacher selector ── */
  const CoTeacherSelect=({coSVal,setCoSFn,coDeptVal,setCoDeptFn,excludeId})=>(
    <div style={{display:"flex",flexDirection:"column",gap:8}}>
      <select style={IS} value={coDeptVal} onChange={e=>{setCoDeptFn(e.target.value);setCoSFn("");}}>
        <option value="">-- เลือกกลุ่มสาระก่อน --</option>
        {S.depts.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}
      </select>
      {coDeptVal&&(
        <select style={IS} value={coSVal} onChange={e=>setCoSFn(e.target.value)}>
          <option value="">-- เลือกครู --</option>
          {S.teachers.filter(t=>t.departmentId===coDeptVal&&t.id!==excludeId).map(t=>{
            const rem=(t.totalPeriods||0)-teacherScheduledTotal(t.id);
            return <option key={t.id} value={t.id}>{t.prefix}{t.firstName} {t.lastName} — เหลือ {rem} คาบ</option>;
          })}
        </select>
      )}
    </div>
  );

  /* ── render timetable table ── */
  const renderTable=(roomIds)=>(
    <div style={{flex:1,overflowX:"auto"}}>
      {roomIds.map(rid=>{
        const rm=S.rooms.find(r=>r.id===rid);
        return (
          <div key={rid} style={{marginBottom:28}}>
            <div style={{marginBottom:8}}>
              <span style={{background:"#DC2626",color:"#fff",padding:"4px 14px",borderRadius:8,fontSize:12,fontWeight:700}}>{rm?.name}</span>
            </div>
            <div style={{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,0.08)",overflow:"hidden"}}>
              <table style={{width:"100%",borderCollapse:"collapse",tableLayout:"fixed",minWidth:700}}>
                <thead>
                  <tr style={{borderBottom:"2px solid #DC2626"}}>
                    <th style={{padding:"10px 8px",background:"#FEF2F2",fontWeight:700,color:"#991B1B",width:60,textAlign:"left",fontSize:13}}>วัน</th>
                    {PERIODS.map(p=>(
                      <th key={p.id} style={{padding:"6px 2px",background:"#FEF2F2",textAlign:"center",borderLeft:"1px solid #FECACA"}}>
                        <div style={{fontSize:11,color:"#991B1B",fontWeight:700}}>คาบ {p.id}</div>
                        <div style={{fontSize:9,color:"#9CA3AF",fontWeight:400}}>{p.time}</div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {DAYS.map((day,di)=>(
                    <tr key={day} style={{background:di%2===0?"#fff":"#FAFAFA"}}>
                      <td style={{padding:"6px 6px",fontWeight:700,fontSize:12,color:"#374151",borderRight:"2px solid #FECACA",borderBottom:"1px solid #F3F4F6"}}>{day}</td>
                      {PERIODS.map(p=>{
                        const key=sk(rid,day,p.id);
                        const en=S.schedule[key]||[];
                        const lk=!!S.locks[key];
                        const bl=mode==="teacher"&&!!selT&&isBlk(selT,day,p.id);
                        return (
                          <td key={p.id}
                            className="dz"
                            onDragOver={e=>{const d=dragRef.current;if(!d){e.currentTarget.classList.remove("over");return;}// ลากจากการ์ด: ตรวจ fromRoomId; ลากจาก sidebar: ตรวจ assignment roomIds
if(d.fromRoomId&&d.fromRoomId!==rid){e.currentTarget.classList.remove("over");return;}
if(d.assignmentId){const a=S.assigns.find(x=>x.id===d.assignmentId);if(!a?.roomIds?.includes(rid)){e.currentTarget.classList.remove("over");return;}}
e.preventDefault();e.currentTarget.classList.add("over");}}
                            onDragLeave={e=>e.currentTarget.classList.remove("over")}
                            onDrop={e=>{e.preventDefault();e.currentTarget.classList.remove("over");handleDrop(rid,day,p.id);}}
                            style={{padding:3,verticalAlign:"top",minHeight:68,borderLeft:"1px solid #F0F0F0",borderBottom:"1px solid #F0F0F0",background:bl?"#FEF9C3":lk?"#F0FDF4":"transparent"}}
                          >
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
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        );
      })}
    </div>
  );

  /* ── render ── */
  return (
    <div style={{animation:"fadeIn 0.3s"}}>

      {/* Mode + selector bar */}
      <div style={{display:"flex",gap:8,marginBottom:14,alignItems:"center",flexWrap:"wrap"}}>
        <div style={{display:"flex",borderRadius:10,overflow:"hidden",border:"1.5px solid #DC2626"}}>
          <button onClick={()=>{setMode("teacher");setSelRoom("");}} style={{padding:"8px 18px",background:mode==="teacher"?"#DC2626":"#fff",color:mode==="teacher"?"#fff":"#DC2626",border:"none",fontWeight:700,fontSize:13,cursor:"pointer",transition:"background 0.15s"}}>จัดรายครู</button>
          <button onClick={()=>{setMode("room");setSelT("");setSelDept("");}} style={{padding:"8px 18px",background:mode==="room"?"#DC2626":"#fff",color:mode==="room"?"#fff":"#DC2626",border:"none",fontWeight:700,fontSize:13,cursor:"pointer",transition:"background 0.15s"}}>จัดรายห้อง</button>
        </div>

        {mode==="teacher"&&<>
          <select style={{...IS,maxWidth:200}} value={selDept} onChange={e=>{setSelDept(e.target.value);setSelT("");}}>
            <option value="">-- ทุกกลุ่มสาระ --</option>
            {S.depts.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}
          </select>
          <select style={{...IS,maxWidth:280}} value={selT} onChange={e=>setSelT(e.target.value)}>
            <option value="">-- เลือกครู --</option>
            {fTeachers.map(t=>{
              const rem=(t.totalPeriods||0)-teacherScheduledTotal(t.id);
              return <option key={t.id} value={t.id}>{t.prefix}{t.firstName} {t.lastName} (เหลือ {rem})</option>;
            })}
          </select>
        </>}

        {mode==="room"&&(
          <select style={{...IS,maxWidth:300}} value={selRoom} onChange={e=>setSelRoom(e.target.value)}>
            <option value="">-- เลือกห้องเรียน --</option>
            {sortedRooms.map(r=>{
              const lv=S.levels.find(l=>l.id===r.levelId);
              return <option key={r.id} value={r.id}>{lv?.name} — {r.name}</option>;
            })}
          </select>
        )}
      </div>

      {/* Teacher summary bar */}
      {mode==="teacher"&&teacher&&(
        <div style={{background:"#fff",borderRadius:10,padding:"10px 16px",marginBottom:12,display:"flex",gap:12,alignItems:"center",flexWrap:"wrap",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
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
        ?<div style={{display:"flex",gap:14}}>
            {/* Sidebar */}
            <div style={{width:270,flexShrink:0,position:"sticky",top:0,alignSelf:"flex-start",maxHeight:"calc(100vh - 200px)",overflowY:"auto"}}>
              <div style={{fontSize:13,fontWeight:700,color:"#374151",marginBottom:10}}>วิชา — ลากวาง</div>
              {asgns.map(a=>{
                const sub=S.subjects.find(s=>s.id===a.subjectId);
                const dept=S.depts.find(d=>d.id===sub?.departmentId);
                const c=dept?gc(dept.id):{bg:"#6B7280",lt:"#F3F4F6",tx:"#374151",bd:"#D1D5DB"};
                const u=aUsed(a.id);
                const rem=a.totalPeriods-u;
                const coCid=cardCoMap[a.id];
                const coTeacher=coCid?S.teachers.find(t=>t.id===coCid):null;
                return (
                  <div key={a.id} style={{background:c.lt,border:"2px solid "+c.bd,borderRadius:12,padding:12,opacity:rem<=0?0.4:1,marginBottom:10}}>
                    <div
                      className="drag-card"
                      draggable={rem>0}
                      onDragStart={()=>setDragBoth({teacherId:selT,subjectId:a.subjectId,assignmentId:a.id})}
                      onDragEnd={()=>setDragBoth(null)}
                      style={{cursor:rem>0?"grab":"default"}}
                    >
                      <div style={{fontSize:13,fontWeight:700,color:c.tx}}>{sub?.code} — {sub?.name}</div>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:8}}>
                        <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                          {a.roomIds.map(rid=>(
                            <span key={rid} style={{background:"rgba(0,0,0,0.1)",padding:"2px 8px",borderRadius:10,fontSize:11,fontWeight:600}}>{S.rooms.find(r=>r.id===rid)?.name}</span>
                          ))}
                        </div>
                        <span style={{background:rem>0?c.bg:"#9CA3AF",color:"#fff",padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700}}>{rem}/{a.totalPeriods}</span>
                      </div>
                    </div>
                    {/* ข้อ 4: เพิ่มครูร่วมบน sidebar */}
                    <div style={{marginTop:8,paddingTop:8,borderTop:"1px solid rgba(0,0,0,0.07)"}}>
                      {coTeacher
                        ?<div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                            <span style={{fontSize:11,color:c.tx}}>ร่วม: {coTeacher.prefix}{coTeacher.firstName}</span>
                            <button onClick={()=>setCardCoMap(p=>({...p,[a.id]:null}))} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",padding:0,fontSize:11}}>✕</button>
                          </div>
                        :<button onClick={()=>setCardCoM(a.id)} style={{fontSize:11,color:c.tx,background:"rgba(0,0,0,0.06)",border:"none",borderRadius:6,padding:"3px 8px",cursor:"pointer",width:"100%",textAlign:"left"}}>+ เพิ่มครูสอนร่วม</button>
                      }
                    </div>
                  </div>
                );
              })}
            </div>
            {renderTable(tRooms)}
          </div>
        :<EmptyState icon="📋" title="เลือกครูเพื่อจัดตาราง"/>
      )}

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
          <button
            onClick={()=>{
              if(!coS||!coM)return;
              const pts=coM.key.split("_");const cDay=pts[1];const cPer=parseInt(pts[2]);
              if(teacherBusy(coS,cDay,cPer,null)){st("ครูท่านนี้สอนคาบนี้อยู่แล้ว","error");return;}
              U.setSchedule(prev=>({...prev,[coM.key]:(prev[coM.key]||[]).map(e=>e.id===coM.entryId?{...e,coTeacherId:coS}:e)}));
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
          <button
            onClick={()=>{
              if(!cardCoS)return;
              setCardCoMap(p=>({...p,[cardCoM]:cardCoS}));
              setCardCoM(null);setCardCoS("");setCardCoDept("");st("กำหนดครูร่วมสำเร็จ");
            }}
            style={BS()}>ยืนยัน</button>
        </div>
      </Modal>
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
        if(e.teacherId!==tid && e.coTeacherId!==tid) return;
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
        if(isBlocked&&rooms.length===0){ cellTxt='X'; extra='background:#ddd;'; }
        else if(rooms.length>0){ cellTxt=rooms.join('<br/>'); extra='background:'+rowBg+';font-weight:700;'; }
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
  // ตารางเรียนระดับชั้น: ห้องเป็นแถว × วัน/คาบเป็นคอลัมน์ (ขาวดำ)
  const subtitle = "ภาคเรียนที่ "+(ay?.semester||"1")+"/"+(ay?.year||"2568")+" "+(sh?.name||"โรงเรียนดาราวิทยาลัย");
  const logoHtml = sh?.logo ? '<img src="'+sh.logo+'" style="width:34px;height:34px;border-radius:50%;object-fit:cover;flex-shrink:0"/>' : '';
  const lvName = filterLevelId ? (S.levels.find(l=>l.id===filterLevelId)?.name||'') : 'ทุกระดับ';
  const title = "ตารางเรียน "+lvName+" ปีการศึกษา "+(ay?.year||"2568");

  // เรียงห้องตามชั้น+เลขห้อง
  const sortKey=(r)=>{ const lv=S.levels.find(l=>l.id===r.levelId)?.name||""; const lvN=parseInt((lv.match(/(\d+)/)||[0,99])[1]); const rmN=parseInt((r.name.match(/(\d+)$/)||[0,0])[1]); return lvN*10000+rmN; };
  let rooms = [...S.rooms].sort((a,b)=>sortKey(a)-sortKey(b));
  if(filterLevelId) rooms=rooms.filter(r=>r.levelId===filterLevelId);
  if(!rooms.length) return '<html><body>ไม่มีห้องเรียนในระดับนี้</body></html>';

  const P=PERIODS.length; const totalCols=DAYS.length*P;
  let headRow1='<th rowspan="2" style="width:52px;border:1px solid #000;background:#e8e8e8;font-size:8px;font-weight:700;padding:2px;text-align:center">ห้อง</th>';
  DAYS.forEach((day,di)=>{
    const br=di<DAYS.length-1?'border-right:2.5px solid #000;':'';
    headRow1+='<th colspan="'+P+'" style="border:1px solid #000;'+br+'background:#333;color:#fff;font-size:8px;font-weight:700;padding:2px 1px;text-align:center">'+day+'</th>';
  });
  let headRow2='';
  DAYS.forEach((_,di)=>{ PERIODS.forEach((p,pi)=>{
    const br=(pi===P-1&&di<DAYS.length-1)?'border-right:2.5px solid #000;':'';
    headRow2+='<th style="border:1px solid #999;'+br+'background:#e8e8e8;font-size:7px;font-weight:700;padding:1px;text-align:center;width:'+(550/totalCols).toFixed(1)+'px">'+p.id+'</th>';
  }); });

  let bodyHTML='';
  // จัดกลุ่มตามระดับชั้น
  const levelIds=[...new Set(rooms.map(r=>r.levelId))];
  levelIds.forEach(lvId=>{
    const lvRooms=rooms.filter(r=>r.levelId===lvId);
    const lvNameStr=S.levels.find(l=>l.id===lvId)?.name||'';
    bodyHTML+='<tr><td colspan="'+(totalCols+1)+'" style="background:#555;color:#fff;font-size:8px;font-weight:700;padding:2px 5px;border:1px solid #000">'+lvNameStr+'</td></tr>';
    lvRooms.forEach((rm,ri)=>{
      const rowBg=ri%2===0?'#fff':'#f5f5f5';
      let row='<tr>';
      row+='<td style="background:'+rowBg+';font-size:8px;padding:2px;border:1px solid #000;font-weight:700;text-align:center;vertical-align:middle">'+rm.name+'</td>';
      DAYS.forEach((_,di)=>{ PERIODS.forEach((p,pi)=>{
        const key=rm.id+"_"+DAYS[di]+"_"+p.id;
        const en=S.schedule[key]||[];
        const br=(pi===P-1&&di<DAYS.length-1)?'border-right:2.5px solid #000;':'';
        let cellTxt=''; let extra='background:'+rowBg+';';
        if(en.length>0){
          cellTxt=en.map(e=>{
            const sub=S.subjects.find(s=>s.id===e.subjectId);
            const t=S.teachers.find(x=>x.id===e.teacherId);
            return (sub?.code||sub?.name||'')+(t?'<br/>'+(t.prefix||'')+(t.firstName||''):'');
          }).join(' / ');
          extra='background:'+rowBg+';font-weight:700;';
        }
        row+='<td style="border:1px solid #ccc;'+br+extra+'font-size:7px;padding:1px 2px;text-align:center;vertical-align:middle;line-height:1.2">'+cellTxt+'</td>';
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
        if(!en.length) return [];
        const isDouble=en.length>1;
        const e=en[sheetIdx]||en[0];
        const sub=S.subjects.find(s=>s.id===e.subjectId);
        const t2=S.teachers.find(x=>x.id===e.teacherId);
        // room2 ใช้เก็บชื่อห้อง, double flag ส่งผ่าน sub prefix
        return[{sub:sub?.name||"",room:(t2?.prefix||"")+(t2?.firstName||""),room2:room.name,double:isDouble}];
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

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"flex",gap:10,marginBottom:24,flexWrap:"wrap"}}>
      <button onClick={exportAllRooms} style={BS("#2563EB")}><Icon name="download" size={16}/>ตารางทุกห้อง (.xlsx)</button>
      <button onClick={exportAllTeachers} style={BS("#7C3AED")}><Icon name="download" size={16}/>ตารางสอนทุกคน (.xlsx)</button>
      <button onClick={exportStatus} style={BS("#059669")}><Icon name="download" size={16}/>รายงานสถานะ (.xlsx)</button>
      <div style={{width:"100%",height:0,borderTop:"1px solid #E5E7EB",margin:"4px 0"}}/>
      <button onClick={printAllTeachersPDF} style={BS("#DC2626")}><Icon name="file" size={16}/>พิมพ์ตารางสอนทุกคน (PDF)</button>
      <button onClick={printAllRoomsPDF} style={BS("#DB2777")}><Icon name="file" size={16}/>พิมพ์ตารางเรียนทุกห้อง (PDF)</button>
      <div style={{width:"100%",height:0,borderTop:"1px solid #E5E7EB",margin:"4px 0"}}/>
      <button onClick={printMasterByDept} style={BS("#374151")}><Icon name="file" size={16}/>📋 ตารางสอนครูรวมกลุ่มสาระ (PDF)</button>
      <div style={{display:"flex",gap:8,alignItems:"center"}}>
        <select style={{...IS,maxWidth:200}} value={masterLevel} onChange={e=>setMasterLevel(e.target.value)}>
          <option value="">-- เลือกระดับชั้น --</option>
          {S.levels.map(l=><option key={l.id} value={l.id}>{l.name}</option>)}
        </select>
        <button onClick={printMasterByLevel} style={BS("#374151")}><Icon name="file" size={16}/>📋 ตารางเรียนรวมระดับชั้น (PDF)</button>
      </div>
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
function Settings({S,U,st,ay,setAY,sh,setSH,div}){
  const logoRef=useRef(null);
  const resetAll=()=>{
    if(!confirm("⚠️ คุณแน่ใจหรือไม่ว่าต้องการลบข้อมูลทั้งหมด?\nข้อมูลที่จัดตารางไว้จะหายทั้งหมด!"))return;
    if(!confirm("ยืนยันอีกครั้ง — ลบข้อมูลทั้งหมดและเริ่มต้นใหม่?"))return;
    U.setLevels((div?.defaultLevels||["ระดับ 1","ระดับ 2","ระดับ 3"]).map(n=>({id:gid(),name:n})));
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
      const isDouble = entries.some(function(e){ return e.double; });
      const inner = entries.map(function(e) {
        let h = '<div class="ent"><div class="ent-sub">' + e.sub + '</div><div class="ent-room">' + e.room + '</div>';
        if (e.room2) h += '<div class="ent-room2">' + e.room2 + '</div>';
        h += '</div>';
        return h;
      }).join("");
      return '<td class="slot' + (isDouble ? ' slot-hi' : '') + '">' + inner + '</td>';
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
    'td.slot-hi{background:#FEF9C3}' +
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
        const isDouble = entries.some(function(e){ return e.double; });
        const inner = entries.map(function(e) {
          let h = '<div class="ent"><div class="ent-sub">' + e.sub + '</div><div class="ent-room">' + e.room + '</div>';
          if (e.room2) h += '<div class="ent-room2">' + e.room2 + '</div>';
          h += '</div>';
          return h;
        }).join("");
        return '<td class="slot' + (isDouble ? ' slot-hi' : '') + '">' + inner + '</td>';
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
    'td.slot-hi{background:#FEF9C3}' +
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
