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

// ===== FIRESTORE TIMETABLE HELPERS (Realtime) =====
const DATA_FIELDS = ["levels","plans","depts","teachers","subjects","rooms","specialRooms","assigns","meetings","schedule","locks"];
// ตรวจ environment: localhost = dev, github.io = production
const IS_DEV = typeof window!=="undefined" && (window.location.hostname==="localhost"||window.location.hostname==="127.0.0.1");
const FS_COLLECTION = IS_DEV ? "timetable_dev" : "timetable";

// Save ข้อมูลทั้งหมดไป Firestore (merge เพื่อไม่ทับ _init)
const fsSaveTimetable = async (divId, data) => {
  const {db} = getFB(); if(!db) return;
  const payload = {};
  DATA_FIELDS.forEach(f => { if(data[f] !== undefined) payload[f] = data[f]; });
  if(data.schoolHeader) payload.schoolHeader = data.schoolHeader;
  if(data.academicYear) payload.academicYear = data.academicYear;
  // ใช้ setDoc ไม่ merge เพื่อให้ schedule ถูก replace ทั้งก้อน (กัน entries เก่าค้าง)
  await setDoc(doc(db,FS_COLLECTION,divId), payload);
};

// Subscribe realtime — returns unsubscribe function
const fsSubscribeTimetable = (divId, onData) => {
  const {db} = getFB(); if(!db) return ()=>{};
  return onSnapshot(doc(db,FS_COLLECTION,divId), (snap) => {
    // ถ้า document ไม่มี → ส่ง {} เพื่อให้ระบบ init state ว่างได้ (ไม่ค้าง syncing)
    onData(snap.exists() ? snap.data() : {});
  }, (err) => { console.warn("Firestore subscribe error:", err); onData({}); });
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
        <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",marginBottom:20}}>
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
        <div style={{background:"#fff",borderRadius:12,padding:16,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",marginBottom:16}}>
          <div style={{position:"relative"}}>
            <input style={{...IS,paddingLeft:36}} value={search} onChange={e=>setSearch(e.target.value)} placeholder="ค้นหาชื่อหรืออีเมล..."/>
            <div style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:"#9CA3AF"}}><Icon name="search" size={14}/></div>
          </div>
        </div>

        {loading&&<div style={{textAlign:"center",padding:40,color:"#6B7280"}}>กำลังโหลด...</div>}

        {/* ── ตารางผู้ใช้ ── */}
        <div style={{background:"#fff",borderRadius:12,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",overflow:"hidden"}}>
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
// ชื่อวิชาย่อ: ใช้ shortName ถ้ามี ไม่งั้นใช้ name เต็ม
const subDisplayName = (sub) => sub?.shortName||sub?.name||"";


// ===== Design tokens (Dara red scheme) =====
const CRED="#B91C1C";      // แดงดารา หลัก
const CBGW="#FFFFFF";       // white card
const IS={width:"100%",padding:"10px 14px",border:"1.5px solid #E5E7EB",borderRadius:12,fontSize:14,outline:"none",fontFamily:"inherit",boxSizing:"border-box",background:"#fff",color:"#1A1A1A"};
const BS=(c=CRED)=>({padding:"10px 20px",background:c,color:"#fff",border:"none",borderRadius:12,fontSize:14,fontWeight:600,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6,fontFamily:"inherit",letterSpacing:"0.01em"});
const BO=(c=CRED)=>({padding:"10px 20px",background:"transparent",color:c,border:`2px solid ${c}`,borderRadius:12,fontSize:14,fontWeight:600,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6,fontFamily:"inherit"});
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
        <input
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
        <div
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
const saveLS = (key, data) => { try { localStorage.setItem(`dara_${key}`, JSON.stringify(data)); } catch(e) {} };
const loadLS = (key, fb) => { try { const d = localStorage.getItem(`dara_${key}`); return d ? JSON.parse(d) : fb; } catch(e) { return fb; } };

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



// ===== TOAST NOTIFICATION =====
const Toast=({message,type="success",onClose})=>{useEffect(()=>{const t=setTimeout(onClose,3000);return()=>clearTimeout(t)},[onClose]);return<div style={{position:"fixed",top:24,right:24,zIndex:9999,background:type==="error"?"#DC2626":type==="warning"?"#D97706":"#059669",color:"#fff",padding:"14px 24px",borderRadius:12,fontSize:14,fontWeight:600,boxShadow:"0 10px 30px rgba(0,0,0,0.2)",display:"flex",alignItems:"center",gap:8,animation:"slideIn 0.3s ease"}}><Icon name={type==="error"?"alert":"check"} size={16}/>{message}</div>};

const DIVISIONS=[
  {id:"p1",name:"ประถมศึกษาตอนต้น",short:"ประถมต้น",defaultLevels:["ป.1","ป.2","ป.3"]},
  {id:"p2",name:"ประถมศึกษาตอนปลาย",short:"ประถมปลาย",defaultLevels:["ป.4","ป.5","ป.6"]},
  {id:"m1",name:"มัธยมศึกษาตอนต้น",short:"มัธยมต้น",defaultLevels:["ม.1","ม.2","ม.3"]},
  {id:"m2",name:"มัธยมศึกษาตอนปลาย",short:"มัธยมปลาย",defaultLevels:["ม.4","ม.5","ม.6"]},
];


// ===== MAIN APP COMPONENT =====
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

        // ตรวจ pre-key เสมอ ถ้า perms ยังไม่มีสิทธิ์ใดเลย (รวมกรณี user เคย login แล้วแต่ admin เพิ่มสิทธิ์ทีหลังผ่าน pre-key)
        const hasAnyAccess=perms&&Object.values(perms.divisions||{}).some(Boolean);
        if((!perms||!hasAnyAccess)&&db){
          const emailKey=makePreKey(u.email);
          const preSnap=await getDoc(doc(db,"permissions",emailKey));
          if(preSnap.exists()&&!preSnap.data().merged){
            const preData=preSnap.data();
            const divs=preData.divisions||{p1:false,p2:false,m1:false,m2:false};
            await setDoc(doc(db,"permissions",u.uid),{
              displayName:u.displayName||"",email:u.email,divisions:divs,preAdded:false,
            },{merge:true});
            await setDoc(doc(db,"permissions",emailKey),{merged:true},{merge:true});
            perms={...perms,divisions:divs};
          }
        }

        if(!perms){
          const emptyDivs={p1:false,p2:false,m1:false,m2:false};
          await fsSetPermissions(u.uid,{displayName:u.displayName||"",email:u.email,divisions:emptyDivs});
          setUserPerms({divisions:emptyDivs});
        } else {
          // merge เฉพาะ displayName/email ไม่ทับ divisions ที่ admin ตั้งไว้
          await setDoc(doc(db,"permissions",u.uid),{displayName:u.displayName||"",email:u.email},{merge:true});
          // re-fetch เพื่อให้ได้ค่าล่าสุดจาก Firestore (กัน race condition)
          const freshPerms=await fsGetPermissions(u.uid);
          setUserPerms(freshPerms||perms);
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
  useEffect(()=>{
    saveLS("schoolHeader",schoolHeader);
    if(fsReadyRef.current) syncToFirestore();
  },[schoolHeader]);
  // บันทึก division ที่เลือกไว้
  useEffect(()=>{ localStorage.setItem("dara_division",divId); },[divId]);

  const stateRef=useRef({});
  useEffect(()=>{stateRef.current={levels,plans,depts,teachers,subjects,rooms,specialRooms,assigns,meetings,schedule,locks}},[levels,plans,depts,teachers,subjects,rooms,specialRooms,assigns,meetings,schedule,locks]);
  // sync schoolHeader และ academicYear ผ่าน GAS ด้วย เพื่อให้ทุกเครื่องเห็นโลโก้และปีการศึกษาเดียวกัน
  const shRef=useRef({});
  useEffect(()=>{shRef.current={schoolHeader,academicYear};},[schoolHeader,academicYear]);

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
    fsReadyRef.current=false;
    setPage("dashboard");
    st("เปลี่ยนเป็น "+d.name);
  };

  // Auto-switch ไปยัง division แรกที่มีสิทธิ์ ถ้า divId ปัจจุบันไม่มีสิทธิ์
  useEffect(()=>{
    if(!firebaseConfigured||!userPerms)return;
    const currentOk=userPerms?.divisions?.[divId]===true;
    if(!currentOk){
      const firstAllowed=DIVISIONS.find(d=>userPerms?.divisions?.[d.id]===true);
      if(firstAllowed&&firstAllowed.id!==divId){
        switchDivision(firstAllowed.id);
      }
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[userPerms]);

  // ===== FIRESTORE REALTIME SYNC =====
  const saveTimer=useRef(null);
  const fsReadyRef=useRef(false); // กัน loop: onSnapshot trigger → setState → save → onSnapshot
  const isSavingRef=useRef(false); // กัน onSnapshot overwrite ขณะ save

  // debounced save ไป Firestore (500ms หลังจากมีการเปลี่ยนแปลง)
  const syncToFirestore=useCallback((immediate=false)=>{
    const {db}=getFB(); if(!db) return;
    clearTimeout(saveTimer.current);
    const delay=immediate?0:500;
    saveTimer.current=setTimeout(async()=>{
      isSavingRef.current=true;
      setSyncing(true);
      try{
        await fsSaveTimetable(divId,{...stateRef.current,schoolHeader:shRef.current?.schoolHeader,academicYear:shRef.current?.academicYear});
      }catch(e){console.warn("Firestore save error:",e);}
      setSyncing(false);
      setTimeout(()=>{isSavingRef.current=false;},300); // รอ onSnapshot ผ่านไปก่อน
    },delay);
  },[divId]);

  // Subscribe realtime onSnapshot เมื่อ login และเมื่อ switch division
  useEffect(()=>{
    const {db}=getFB(); if(!db||!authUser) return;
    fsReadyRef.current=false;
    setSyncing(true);
    const unsub=fsSubscribeTimetable(divId,(d)=>{
      if(!fsReadyRef.current){
        // โหลดครั้งแรก: Firestore เป็น source of truth — ล้าง localStorage cache เก่าก่อน
        const keepKeys=["dara_academicYear","dara_schoolHeader","dara_division"];
        Object.keys(localStorage)
          .filter(k=>k.startsWith("dara_"+divId)&&!keepKeys.includes(k))
          .forEach(k=>localStorage.removeItem(k));
        // set state จาก Firestore (ถ้า Firestore ว่าง ก็ว่างจริงๆ ไม่เอา localStorage)
        setLevels(d.levels?.length?d.levels:DIVISIONS.find(x=>x.id===divId)?.defaultLevels.map(n=>({id:gid(),name:n}))||[]);
        setPlans(d.plans||[]);
        setDepts(d.depts||[]);
        setTeachers(d.teachers||[]);
        setSubjects(d.subjects||[]);
        setRooms(d.rooms||[]);
        setSpecialRooms(d.specialRooms||[]);
        setAssigns(d.assigns||[]);
        setMeetings(d.meetings||[]);
        setSchedule(d.schedule||{});
        setLocks(d.locks||{});
        if(d.schoolHeader?.name)   setSchoolHeader(sh=>({...sh,...d.schoolHeader}));
        if(d.academicYear?.year)   setAcademicYear(ay=>({...ay,...d.academicYear}));
        fsReadyRef.current=true;
        setSyncing(false);
        setGasReady(true);
      } else {
        // Realtime update จากเครื่องอื่น — skip ถ้ากำลัง save อยู่ (กัน overwrite)
        if(isSavingRef.current) return;
        if(d.levels)       setLevels(d.levels);
        if(d.plans)        setPlans(d.plans);
        if(d.depts)        setDepts(d.depts);
        if(d.teachers)     setTeachers(d.teachers);
        if(d.subjects)     setSubjects(d.subjects);
        if(d.rooms)        setRooms(d.rooms);
        if(d.specialRooms) setSpecialRooms(d.specialRooms);
        if(d.assigns)      setAssigns(d.assigns);
        if(d.meetings)     setMeetings(d.meetings);
        if(d.schedule)     setSchedule(d.schedule);
        if(d.locks)        setLocks(d.locks);
        if(d.schoolHeader?.name) setSchoolHeader(sh=>({...sh,...d.schoolHeader}));
        if(d.academicYear?.year) setAcademicYear(ay=>({...ay,...d.academicYear}));
      }
    });
    // timeout fallback: ถ้า 8 วินาทียัง sync ไม่เสร็จ ให้ถือว่าเสร็จ (document อาจไม่มี)
    const fallback=setTimeout(()=>{
      if(!fsReadyRef.current){
        fsReadyRef.current=true;
        setSyncing(false);
        setGasReady(true);
      }
    },8000);
    return ()=>{ unsub(); clearTimeout(fallback); fsReadyRef.current=false; };
  },[divId, authUser]);

  // Auto-save ไป localStorage (cache offline) + Firestore เมื่อข้อมูลเปลี่ยน
  useEffect(()=>{ saveLS(divId+"_levels",levels);       if(fsReadyRef.current) syncToFirestore(); },[levels]);
  useEffect(()=>{ saveLS(divId+"_plans",plans);         if(fsReadyRef.current) syncToFirestore(); },[plans]);
  useEffect(()=>{ saveLS(divId+"_depts",depts);         if(fsReadyRef.current) syncToFirestore(); },[depts]);
  useEffect(()=>{ saveLS(divId+"_teachers",teachers);   if(fsReadyRef.current) syncToFirestore(); },[teachers]);
  useEffect(()=>{ saveLS(divId+"_subjects",subjects);   if(fsReadyRef.current) syncToFirestore(); },[subjects]);
  useEffect(()=>{ saveLS(divId+"_rooms",rooms);         if(fsReadyRef.current) syncToFirestore(); },[rooms]);
  useEffect(()=>{ saveLS(divId+"_specialRooms",specialRooms); if(fsReadyRef.current) syncToFirestore(); },[specialRooms]);
  useEffect(()=>{ saveLS(divId+"_assigns",assigns);     if(fsReadyRef.current) syncToFirestore(); },[assigns]);
  useEffect(()=>{ saveLS(divId+"_meetings",meetings);   if(fsReadyRef.current) syncToFirestore(); },[meetings]);
  useEffect(()=>{ saveLS(divId+"_schedule",schedule);   if(fsReadyRef.current) syncToFirestore(); },[schedule]);
  useEffect(()=>{ saveLS(divId+"_locks",locks);         if(fsReadyRef.current) syncToFirestore(); },[locks]);

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

  return <div style={{display:"flex",height:"100vh",fontFamily:"'Sarabun','Noto Sans Thai',sans-serif",background:"linear-gradient(145deg,#EEF2FF 0%,#F8F9FF 40%,#FFF5F5 100%)",overflow:"hidden"}}>
    <style>{`@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700;800&display=swap');*{box-sizing:border-box;margin:0;padding:0}::-webkit-scrollbar{width:5px}::-webkit-scrollbar-thumb{background:#D4C5BA;border-radius:4px}::-webkit-scrollbar-track{background:transparent}@keyframes slideIn{from{transform:translateX(100px);opacity:0}to{transform:translateX(0);opacity:1}}@keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}.ni:hover{background:rgba(255,255,255,0.12)!important;border-radius:10px}.ni.a{background:rgba(255,255,255,0.15)!important;border-radius:10px}input:focus,select:focus{border-color:#991B1B!important;box-shadow:0 0 0 3px rgba(153,27,27,0.12)!important}input,select{transition:border-color 0.15s,box-shadow 0.15s}.drag-card{cursor:grab;user-select:none}.drag-card:active{cursor:grabbing}.dz{transition:background 0.15s,outline 0.15s}.dz.over{background:#FEE2E2!important;outline:2px dashed #DC2626}button:hover{opacity:0.88;transition:opacity 0.15s}select{appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%236B7280' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 12px center;padding-right:36px!important}.div-sel{appearance:none!important;background:rgba(0,0,0,0.2)!important;background-image:none!important;border:1px solid rgba(255,255,255,0.25)!important;border-radius:10px!important;color:#fff!important;font-size:13px!important;font-weight:600!important;font-family:inherit!important;padding:8px 32px 8px 12px!important;width:100%!important;cursor:pointer!important;outline:none!important;transition:border-color 0.15s}.div-sel:focus{box-shadow:0 0 0 2px rgba(255,255,255,0.2)!important;border-color:rgba(255,255,255,0.5)!important}.div-sel option{background:#991B1B;color:#fff}`}</style>

    <div style={{width:side?240:0,background:"linear-gradient(180deg,#B91C1C 0%,#991B1B 100%)",transition:"width 0.3s",overflow:"hidden",flexShrink:0,display:"flex",flexDirection:"column",boxShadow:"2px 0 12px rgba(185,28,28,0.2)"}}>
      <div style={{padding:"20px 16px",borderBottom:"1px solid rgba(255,255,255,0.1)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          {schoolHeader.logo
            ?<img src={schoolHeader.logo} alt="logo" style={{width:38,height:38,borderRadius:10,objectFit:"cover",flexShrink:0}}/>
            :<div style={{width:38,height:38,borderRadius:10,background:"rgba(255,255,255,0.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,fontWeight:800,color:"#fff",flexShrink:0}}>ด</div>
          }
          <div><div style={{color:"#fff",fontSize:14,fontWeight:700}}>{schoolHeader.name||"ดาราวิทยาลัย"}</div><div style={{color:"rgba(255,255,255,0.6)",fontSize:10}}>ระบบจัดตารางสอน v3</div></div>
        </div>
      </div>
      {/* Division selector — dropdown */}
      <div style={{padding:"10px 12px",borderBottom:"1px solid rgba(255,255,255,0.1)"}}>
        <div style={{fontSize:10,color:"rgba(255,255,255,0.5)",marginBottom:5,paddingLeft:2,fontWeight:600}}>ระดับการศึกษา</div>
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
        {nav.map(n=><div key={n.id} className={`ni ${page===n.id?"a":""}`} onClick={()=>setPage(n.id)} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 12px",borderRadius:10,cursor:"pointer",color:page===n.id?"#fff":"rgba(255,255,255,0.7)",fontSize:13,fontWeight:page===n.id?700:400,marginBottom:2,transition:"all 0.15s",background:page===n.id?"rgba(255,255,255,0.15)":"transparent"}}><Icon name={n.icon} size={16}/>{n.label}</div>)}
      </nav>
      <div style={{padding:"12px 16px",borderTop:"1px solid #F3F4F6"}}>
        {firebaseConfigured&&authUser&&<div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,padding:"8px 10px",background:"#F9FAFB",borderRadius:10}}>
          <div style={{width:32,height:32,borderRadius:"50%",background:"linear-gradient(135deg,#B91C1C,#991B1B)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,color:"#fff",flexShrink:0}}>
            {(authUser.displayName||authUser.email||"U")[0].toUpperCase()}
          </div>
          <div style={{flex:1,minWidth:0}}>
            <div style={{color:"#111",fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{authUser.displayName||authUser.email}</div>
            <div style={{color:"#9CA3AF",fontSize:10,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{authUser.email}</div>
          </div>
        </div>}
        <div style={{display:"flex",gap:6,marginBottom:6}}>
          {firebaseConfigured&&<button onClick={handleLogout} style={{flex:1,padding:"6px 0",background:"#F9FAFB",border:"1px solid #FECACA",borderRadius:8,color:CRED,fontSize:11,fontWeight:600,cursor:"pointer"}}>ออกจากระบบ</button>}
          {firebaseConfigured&&<button onClick={()=>setShowAdmin(true)} style={{flex:1,padding:"6px 0",background:"#F9FAFB",border:"1px solid #E5E7EB",borderRadius:8,color:"#374151",fontSize:11,fontWeight:600,cursor:"pointer"}}>🔐 Admin</button>}
        </div>
        <div style={{color:"#D1D5DB",fontSize:10,textAlign:"center"}}>พัฒนาโดย พนิต เกิดมงคล</div>
      </div>
    </div>

    <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
      <header style={{height:60,background:"rgba(255,255,255,0.85)",backdropFilter:"blur(12px)",borderBottom:"1px solid rgba(240,240,240,0.8)",display:"flex",alignItems:"center",padding:"0 20px",gap:12,flexShrink:0,boxShadow:"0 1px 8px rgba(0,0,0,0.05)"}}>
        <button onClick={()=>setSide(!side)} style={{background:"none",border:"none",cursor:"pointer",color:"#9CA3AF",padding:4,borderRadius:8,display:"flex"}}><Icon name="menu" size={20}/></button>
        <div style={{display:"flex",alignItems:"center",gap:6,fontSize:13,color:"#9CA3AF"}}>
          <span style={{cursor:"pointer",color:"#9CA3AF"}} onClick={()=>setPage("dashboard")}>🏠</span>
          <span>/</span>
          <span style={{color:"#111",fontWeight:600}}>{nav.find(n=>n.id===page)?.label}</span>
        </div>
        <span style={{fontSize:10,background:"#F9FAFB",color:CRED,padding:"2px 10px",borderRadius:20,fontWeight:700,border:"1px solid #FECACA"}}>{div.short}</span>
        <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:10}}>
          {syncing
            ?<span style={{fontSize:11,color:"#D97706",background:"#FFFBEB",padding:"3px 10px",borderRadius:20,border:"1px solid #FDE68A",fontWeight:600}}>⏳ กำลัง sync...</span>
            :<span style={{fontSize:11,color:"#059669",background:"#F0FDF4",padding:"3px 10px",borderRadius:20,border:"1px solid #BBF7D0",fontWeight:600}}>● sync แล้ว</span>
          }
          {firebaseConfigured&&authUser?.photoURL&&(
            <img src={authUser.photoURL} alt="avatar" style={{width:32,height:32,borderRadius:"50%",objectFit:"cover",border:"2px solid #E5E7EB"}}/>
          )}
        </div>
      </header>
      <main style={{flex:1,overflow:"auto",padding:"20px 24px",background:"#F3F4F6"}}>
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
            {page==="homeroom"&&<HomeroomSettings S={S} U={U} st={st}/>}
            {page==="plans"&&<Plans S={S} U={U} st={st}/>}
            {page==="departments"&&<Depts S={S} U={U} st={st} gc={gc}/>}
            {page==="teachers"&&<Teachers S={S} U={U} st={st} gc={gc}/>}
            {page==="subjects"&&<Subjects S={S} U={U} st={st} gc={gc}/>}
            {page==="specialrooms"&&<SpecialRooms S={S} U={U} st={st}/>}
            {page==="assignments"&&<Assigns S={S} U={U} st={st} gc={gc}/>}
            {page==="meetings"&&<Meetings S={S} U={U} st={st} gc={gc}/>}
            {page==="scheduler"&&<Scheduler S={S} U={U} st={st} gc={gc} isSavingRef={isSavingRef} fsReadyRef={fsReadyRef} fsSave={(s)=>fsSaveTimetable(divId,{...stateRef.current,schedule:s})}/>}
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
      {stats.map((s,i)=><div key={i} style={{background:"#fff",borderRadius:14,padding:20,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}><div style={{fontSize:28,fontWeight:800}}>{s.v}</div><div style={{fontSize:13,color:"#6B7280",marginTop:2}}>{s.l}</div><div style={{height:4,background:s.c,borderRadius:2,marginTop:12,width:"40%"}}/></div>)}
    </div>
    <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
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
      {S.levels.map(lv=><div key={lv.id} style={{background:"#fff",borderRadius:14,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",overflow:"hidden"}}>
        <div style={{background:"linear-gradient(135deg,#991B1B,#DC2626)",padding:"16px 20px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <h3 style={{color:"#fff",fontSize:18,fontWeight:700}}>{lv.name}</h3>
          <div style={{display:"flex",gap:6}}>
            <button onClick={()=>editLv(lv)} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:6,padding:6,color:"#fff",cursor:"pointer"}}><Icon name="edit" size={14}/></button>
            <button onClick={()=>{U.setLevels(p=>p.filter(l=>l.id!==lv.id));st("ลบแล้ว","warning")}} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:6,padding:6,color:"#fff",cursor:"pointer"}}><Icon name="trash" size={14}/></button>
          </div>
        </div>
        <div style={{padding:16}}>
          {/* วันเข้าหอประชุม */}
          <div style={{marginBottom:10,padding:"8px 12px",background:"#F0F9FF",borderRadius:8,border:"1px solid #BAE6FD"}}>
            <div style={{fontSize:11,fontWeight:700,color:"#0369A1",marginBottom:6}}>🏛️ วันเข้าหอประชุม (Assembly 08.00-08.30)</div>
            <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
              {[{val:"",label:"ไม่มี"},...DAYS.map(d=>({val:d,label:d}))].map(opt=>(
                <button key={opt.val}
                  onClick={()=>U.setLevels(p=>p.map(l=>l.id===lv.id?{...l,assemblyDay:opt.val}:l))}
                  style={{padding:"3px 10px",borderRadius:20,border:`1.5px solid ${(lv.assemblyDay||"")===(opt.val)?"#0369A1":"#D1D5DB"}`,background:(lv.assemblyDay||"")===(opt.val)?"#0369A1":"#fff",color:(lv.assemblyDay||"")===(opt.val)?"#fff":"#374151",fontSize:11,fontWeight:(lv.assemblyDay||"")===(opt.val)?700:400,cursor:"pointer"}}>
                  {opt.label}
                </button>
              ))}
            </div>
          </div>
          <div style={{fontSize:12,fontWeight:600,color:"#9CA3AF",marginBottom:6}}>ห้องเรียน:</div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {S.rooms.filter(r=>r.levelId===lv.id).map(rm=>{
              const plan=S.plans.find(p=>p.id===rm.planId);
              return<span key={rm.id} style={{background:"#DBEAFE",color:"#1E40AF",fontSize:12,padding:"4px 12px",borderRadius:20,fontWeight:600,display:"inline-flex",alignItems:"center",gap:4}}>
                {rm.name}{plan?" ("+plan.name+")":""}
                <button onClick={()=>{const n=prompt("แก้ไขชื่อห้อง:",rm.name);if(n){U.setRooms(p=>p.map(r=>r.id===rm.id?{...r,name:n}:r));st("แก้ไขสำเร็จ")}}} style={{background:"none",border:"none",cursor:"pointer",color:"#1E40AF",padding:0}}><Icon name="edit" size={10}/></button>
                <button onClick={()=>U.setRooms(p=>p.filter(r=>r.id!==rm.id))} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",padding:0}}><Icon name="x" size={10}/></button>
              </span>;
            })}
            {!S.rooms.filter(r=>r.levelId===lv.id).length&&<span style={{fontSize:12,color:"#9CA3AF"}}>ยังไม่มี</span>}
          </div>
        </div>
      </div>)}
    </div>
    <Modal open={rm} onClose={()=>setRm(false)} title="เพิ่มห้องเรียน">
      <div style={{display:"flex",flexDirection:"column",gap:16}}>
        <div><label style={LS}>ระดับชั้น</label><select style={IS} value={rf.levelId} onChange={e=>setRf(p=>({...p,levelId:e.target.value}))}><option value="">--</option>{S.levels.map(l=><option key={l.id} value={l.id}>{l.name}</option>)}</select></div>
        <div><label style={LS}>แผนการเรียน (ถ้ามี)</label><select style={IS} value={rf.planId} onChange={e=>setRf(p=>({...p,planId:e.target.value}))}><option value="">--</option>{S.plans.filter(p=>!p.levelIds?.length||p.levelIds.includes(rf.levelId)).map(p=>{const subs=p.subPlans?.length?" — "+p.subPlans.join(", "):"";return<option key={p.id} value={p.id}>{p.name}{subs}</option>})}</select></div>
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
      {S.plans.map(plan=><div key={plan.id} style={{background:"#fff",borderRadius:14,padding:20,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
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
      {S.depts.map(d=>{const c=gc(d.id);return<div key={d.id} style={{background:"#fff",borderRadius:14,overflow:"hidden",boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
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

/* ===== TEACHERS ===== */
function Teachers({S,U,st,gc}){
  const [modal,setModal]=useState(false);
  const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({prefix:"",firstName:"",lastName:"",teacherCode:"",departmentId:"",specialRoles:[],totalPeriods:0});
  const resetForm=()=>setForm({prefix:"",firstName:"",lastName:"",teacherCode:"",departmentId:"",specialRoles:[],totalPeriods:0});
  const [search,setSearch]=useState("");
  const fileRef=useRef(null);

  const save=()=>{
    if(!form.firstName||!form.departmentId){st("กรุณากรอกให้ครบ","error");return}
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

  const filtered=S.teachers.filter(t=>`${t.prefix}${t.firstName} ${t.lastName}`.includes(search)||S.depts.find(d=>d.id===t.departmentId)?.name?.includes(search));

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"flex",gap:10,marginBottom:20,flexWrap:"wrap",alignItems:"center"}}>
      <button onClick={()=>{setEditId(null);resetForm();setModal(true)}} style={BS()}><Icon name="plus" size={16}/>เพิ่มครู</button>
      <button onClick={()=>fileRef.current?.click()} style={BS("#2563EB")}><Icon name="upload" size={16}/>Import Excel</button>
      <button onClick={downloadTemplate} style={BO("#2563EB")}><Icon name="file" size={16}/>Template</button>
      <button onClick={exportT} style={BO("#059669")}><Icon name="download" size={16}/>Export Excel</button>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleFile}/>
      <div style={{position:"relative",flex:"1 1 200px",maxWidth:350}}><input style={{...IS,paddingLeft:38}} value={search} onChange={e=>setSearch(e.target.value)} placeholder="ค้นหาครู..."/><div style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:"#9CA3AF"}}><Icon name="search" size={16}/></div></div>
    </div>

    <div style={{background:"#fff",borderRadius:14,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",overflow:"auto"}}>
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
        <div><label style={LS}>รหัสครู (Username)</label><input style={IS} value={form.teacherCode||""} onChange={e=>setForm(p=>({...p,teacherCode:e.target.value}))} placeholder="เช่น T001, prachya@dara.ac.th"/></div>
        <div><label style={LS}>กลุ่มสาระ</label><SearchSelect value={form.departmentId} onChange={v=>setForm(p=>({...p,departmentId:v}))} options={[{value:"",label:"--"},...S.depts.map(d=>({value:d.id,label:d.name}))]} placeholder="-- เลือกกลุ่มสาระ --"/></div>
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
      {S.specialRooms.map(r=>{const sc=subCount(r.id);return<div key={r.id} style={{background:"#fff",borderRadius:14,padding:18,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",borderLeft:"4px solid #7C3AED"}}>
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
    return<div style={{background:"#fff",borderRadius:12,overflow:"hidden",boxShadow:"0 2px 12px rgba(0,0,0,0.06)",borderLeft:"3px solid "+c.bg}}>
      <div style={{padding:"12px 14px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:10,color:"#9CA3AF",fontWeight:600}}>{sub.code}</div>
            <h4 style={{fontSize:14,fontWeight:700,marginTop:1,wordBreak:"break-word"}}>{sub.name}</h4>
            {sub.shortName&&<div style={{fontSize:11,color:"#6B7280",marginTop:1}}>ชื่อย่อ: <strong>{sub.shortName}</strong></div>}
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
          {sub.allDepts&&<span style={{background:"#FEF9C3",color:"#92400E",padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:700}}>🏫 ทุกกลุ่มสาระ</span>}
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
        <div><label style={LS}>ชื่อวิชาเต็ม</label><input style={IS} value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))} placeholder="ฟิสิกส์ 4"/></div>
        <div><label style={LS}>ชื่อย่อ <span style={{fontWeight:400,color:"#9CA3AF"}}>(แสดงบนการ์ดและตารางพิมพ์)</span></label><input style={IS} value={form.shortName||""} onChange={e=>setForm(p=>({...p,shortName:e.target.value}))} placeholder="ฟิสิกส์"/></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          <div><label style={LS}>หน่วยกิต</label><input type="number" min="0.5" step="0.5" style={IS} value={form.credits} onChange={e=>setForm(p=>({...p,credits:parseFloat(e.target.value)||0}))}/></div>
          <div><label style={LS}>คาบ/สัปดาห์</label><input type="number" min="1" style={IS} value={form.periodsPerWeek} onChange={e=>setForm(p=>({...p,periodsPerWeek:parseInt(e.target.value)||1}))}/></div>
        </div>
        <div><label style={LS}>ระดับชั้น</label><SearchSelect value={form.levelId} onChange={v=>setForm(p=>({...p,levelId:v}))} options={[{value:"",label:"--"},...S.levels.map(l=>({value:l.id,label:l.name}))]} placeholder="-- เลือกระดับชั้น --"/></div>
        <div><label style={LS}>กลุ่มสาระ</label><SearchSelect value={form.departmentId} onChange={v=>setForm(p=>({...p,departmentId:v}))} options={[{value:"",label:"--"},...S.depts.map(d=>({value:d.id,label:d.name}))]} placeholder="-- เลือกกลุ่มสาระ --"/></div>
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

        {/* allDepts flag */}
        <label style={{display:"flex",alignItems:"flex-start",gap:12,padding:"12px 14px",borderRadius:12,border:`2px solid ${form.allDepts?"#D97706":"#E5E7EB"}`,background:form.allDepts?"#FFFBEB":"#F9FAFB",cursor:"pointer"}}>
          <input type="checkbox" checked={!!form.allDepts} onChange={e=>setForm(p=>({...p,allDepts:e.target.checked}))} style={{marginTop:2,accentColor:"#D97706",flexShrink:0}}/>
          <div>
            <div style={{fontSize:13,fontWeight:700,color:form.allDepts?"#92400E":"#374151"}}>🏫 วิชาที่ทุกกลุ่มสาระสอนร่วมกัน</div>
            <div style={{fontSize:11,color:"#6B7280",marginTop:2}}>เช่น กิจกรรมพัฒนาผู้เรียน, ลูกเสือ — ครูต่างสาระสามารถ assign วิชานี้ได้ และระบบจะตรวจการชนของครูทุกคนที่สอนวิชานี้</div>
          </div>
        </label>

        <button onClick={save} style={BS()}>{editId?"บันทึก":"เพิ่มวิชา"}</button>
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
    <div style={{background:"#fff",borderRadius:14,padding:20,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",marginBottom:20}}>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:16}}>
        <span style={{fontSize:20}}>🔒</span>
        <h3 style={{fontSize:15,fontWeight:700,margin:0}}>คาบล็อกส่วนตัว</h3>
        <span style={{fontSize:12,color:"#6B7280"}}>— {teacher.prefix}{teacher.firstName} {teacher.lastName}</span>
      </div>
      <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end",marginBottom:16,padding:"14px 16px",background:"#FFF7ED",borderRadius:12,border:"1px solid #FED7AA"}}>
        <div style={{flex:"1 1 130px"}}>
          <label style={LS}>วัน</label>
          <select style={IS} value={plDay} onChange={e=>setPlDay(e.target.value)}>
            <option value="">-- เลือกวัน --</option>
            {DAYS.map(d=><option key={d}>{d}</option>)}
          </select>
        </div>
        <div style={{flex:"2 1 300px"}}>
          <label style={LS}>คาบ (เลือกได้หลายคาบ)</label>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {PERIODS.map(p=>(
              <button key={p.id}
                onClick={()=>setPlPeriods(prev=>prev.includes(p.id)?prev.filter(x=>x!==p.id):[...prev,p.id])}
                style={{width:44,height:44,borderRadius:8,border:`2px solid ${plPeriods.includes(p.id)?"#DC2626":"#D1D5DB"}`,background:plPeriods.includes(p.id)?"#DC2626":"#fff",color:plPeriods.includes(p.id)?"#fff":"#374151",fontSize:14,fontWeight:700,cursor:"pointer"}}>
                {p.id}
              </button>
            ))}
          </div>
        </div>
        <div style={{flex:"1 1 160px"}}>
          <label style={LS}>เหตุผล (ไม่บังคับ)</label>
          <input style={IS} value={plReason} onChange={e=>setPlReason(e.target.value)} placeholder="ติดธุระ, อบรม ฯ" onKeyDown={e=>e.key==="Enter"&&addLock()}/>
        </div>
        <button onClick={addLock} style={{...BS("#C2410C"),flexShrink:0}}>+ เพิ่มล็อก</button>
      </div>
      {personalLocks.length===0
        ?<div style={{textAlign:"center",color:"#9CA3AF",fontSize:13,padding:"12px 0"}}>ยังไม่มีคาบล็อกส่วนตัว</div>
        :<div style={{display:"flex",flexDirection:"column",gap:8}}>
          {[...personalLocks].sort((a,b)=>DAYS.indexOf(a.day)-DAYS.indexOf(b.day)).map(pl=>(
            <div key={pl.id} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 14px",background:"#FFF7ED",borderRadius:10,border:"1px solid #FED7AA"}}>
              <span style={{fontSize:16}}>🔒</span>
              <div style={{flex:1}}>
                <span style={{fontWeight:700,color:"#C2410C",fontSize:13}}>วัน{pl.day}</span>
                <span style={{color:"#6B7280",fontSize:12,marginLeft:8}}>คาบ {(pl.periods||[]).join(", ")}</span>
                {pl.reason&&<span style={{marginLeft:8,fontSize:11,background:"#FFEDD5",color:"#9A3412",padding:"1px 8px",borderRadius:20,fontWeight:600}}>{pl.reason}</span>}
              </div>
              <button onClick={()=>removeLock(pl.id)} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",padding:4}}><Icon name="trash" size={14}/></button>
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
      alert(`${ns.length>0?`นำเข้าสำเร็จ ${ns.length} รายการ\n\n`:""}ข้าม ${failLog.length} รายการ:\n${lines}${extra}\n\n💡 วิธีแก้: กด Export ก่อน แล้วใช้ไฟล์นั้นเป็นแม่แบบ`);
    } else if(ns.length){
      st(`นำเข้า ${ns.length} รายการ`);
    }
    e.target.value="";
  };

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"flex",gap:12,marginBottom:24,alignItems:"center",flexWrap:"wrap"}}>
      <SearchSelect value={selDept} onChange={v=>{setSelDept(v);setSel("")}} options={[{value:"",label:"-- เลือกกลุ่มสาระก่อน --"},...S.depts.map(d=>({value:d.id,label:d.name}))]} placeholder="-- เลือกกลุ่มสาระก่อน --" style={{maxWidth:280}}/>
      {selDept&&<SearchSelect value={sel} onChange={v=>setSel(v)}
        options={[{value:"",label:"-- เลือกครู --"},
          // แสดงครูกลุ่มสาระนั้นก่อน จากนั้นครูสาระอื่น
          ...S.teachers.filter(t=>t.departmentId===selDept).map(t=>({value:t.id,label:`${t.prefix}${t.firstName} ${t.lastName}`})),
          ...(S.teachers.filter(t=>t.departmentId!==selDept).length>0
            ? [{value:"__sep__",label:"──── ครูกลุ่มสาระอื่น ────",disabled:true},...S.teachers.filter(t=>t.departmentId!==selDept).map(t=>{const d=S.depts.find(x=>x.id===t.departmentId);return{value:t.id,label:`${t.prefix}${t.firstName} ${t.lastName}${d?" ["+d.name+"]":""}`};})]
            : [])
        ]}
        placeholder="-- เลือกครู --" style={{maxWidth:380}}/> }
      {sel&&<button onClick={()=>{setForm({subjectId:"",roomIds:[],totalPeriods:0});setBasket([]);setModalDeptFilter(teacher?.departmentId||"");setModal(true)}} style={BS()}><Icon name="plus" size={16}/>เพิ่มวิชา</button>}
      <div style={{marginLeft:"auto",display:"flex",gap:8}}>
        <button onClick={exportAssigns} style={BO("#059669")}><Icon name="download" size={16}/>Export ทั้งหมด</button>
        <button onClick={()=>fileRefA.current?.click()} style={BO("#2563EB")}><Icon name="upload" size={16}/>Import</button>
        <input ref={fileRefA} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={importAssigns}/>
      </div>
    </div>
    {teacher&&<div>
      <div style={{background:"#fff",borderRadius:14,padding:20,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",marginBottom:20,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
        <div><h3 style={{fontSize:18,fontWeight:700}}>{teacher.prefix}{teacher.firstName} {teacher.lastName}</h3><div style={{fontSize:13,color:"#6B7280",marginTop:4}}>{S.depts.find(d=>d.id===teacher.departmentId)?.name}</div></div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <div style={{background:"#DBEAFE",color:"#1E40AF",padding:"8px 16px",borderRadius:10,fontWeight:700,fontSize:13}}>📋 ได้รับ: {teacherQuota}</div>
          <div style={{background:"#FEF3C7",color:"#92400E",padding:"8px 16px",borderRadius:10,fontWeight:700,fontSize:13}}>📝 มอบหมาย: {totalAssigned}</div>
          <div style={{background:"#D1FAE5",color:"#065F46",padding:"8px 16px",borderRadius:10,fontWeight:700,fontSize:13}}>✅ ลงตารางแล้ว: {totalScheduled}</div>
          <div style={{background:notScheduled>0?"#FEE2E2":"#F3F4F6",color:notScheduled>0?"#991B1B":"#6B7280",padding:"8px 16px",borderRadius:10,fontWeight:700,fontSize:13}}>
            {notScheduled>0?"⚠️ ยังไม่ลง: "+notScheduled:"✓ ลงครบแล้ว"}
          </div>
          <div style={{background:remaining>=0?"#EFF6FF":"#FEE2E2",color:remaining>=0?"#1D4ED8":"#991B1B",padding:"8px 16px",borderRadius:10,fontWeight:700,fontSize:13}}>เหลือ: {remaining}</div>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:16}}>
        {asgns.map(a=>{const sub=S.subjects.find(s=>s.id===a.subjectId);const dept=S.depts.find(d=>d.id===sub?.departmentId);const c=dept?gc(dept.id):{bg:"#6B7280",lt:"#F3F4F6",tx:"#374151"};const ca=sub?.consecutiveAllowed||0;return<div key={a.id} style={{background:"#fff",borderRadius:14,borderLeft:`4px solid ${c.bg}`,padding:16,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
          {(()=>{
            const aScheduled=(()=>{
              const seen=new Set();let cnt=0;
              Object.entries(S.schedule).forEach(([k,en])=>{
                const pts=k.split("_");
                en?.forEach(e=>{
                  if(e.assignmentId!==a.id)return;
                  const sub2=S.subjects.find(s=>s.id===e.subjectId);
                  const ca2=sub2?.consecutiveAllowed||0;
                  if(ca2===-1||ca2===-2){const npk=e.subjectId+"_"+pts[pts.length-2]+"_"+pts[pts.length-1];if(!seen.has(npk)){seen.add(npk);cnt++;}}
                  else cnt++;
                });
              });
              return cnt;
            })();
            // NP/-2: มอบหมายที่แสดงควรเป็น periodsPerWeek ไม่ใช่ totalPeriods (ที่อาจ × ห้อง)
            const aAssigned=(ca===-1||ca===-2)?(sub?.periodsPerWeek||a.totalPeriods):a.totalPeriods;
            const aPending=aAssigned-aScheduled;
            return <div style={{display:"flex",justifyContent:"space-between"}}><div>
            <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
              <h4 style={{fontSize:15,fontWeight:700}}>{sub?.code} — {subDisplayName(sub)}</h4>
              {ca===-1&&<span style={{fontSize:9,background:"#EFF6FF",color:"#1E40AF",padding:"1px 6px",borderRadius:8,fontWeight:700}}>🔀NP</span>}
              {ca===-2&&<span style={{fontSize:9,background:"#FDF4FF",color:"#6B21A8",padding:"1px 6px",borderRadius:8,fontWeight:700}}>🏛️เศรษฐ-วิศวะ</span>}
              {ca>0&&<span style={{fontSize:9,background:"#FEF3C7",color:"#92400E",padding:"1px 6px",borderRadius:8,fontWeight:700}}>⚡{ca}ติด</span>}
            </div>
            <div style={{display:"flex",gap:6,marginTop:6,flexWrap:"wrap"}}>
              <span style={{fontSize:11,color:"#6B7280"}}>มอบหมาย {aAssigned} คาบ</span>
              <span style={{fontSize:11,background:"#D1FAE5",color:"#065F46",padding:"1px 8px",borderRadius:20,fontWeight:600}}>✅ ลงแล้ว {aScheduled}</span>
              {aPending>0&&<span style={{fontSize:11,background:"#FEE2E2",color:"#991B1B",padding:"1px 8px",borderRadius:20,fontWeight:600}}>⚠️ ยังไม่ลง {aPending}</span>}
            </div>
          </div>
            <div style={{display:"flex",gap:6}}>
              <button onClick={()=>{const n=prompt("แก้ไขจำนวนคาบ:",a.totalPeriods);if(n!==null){U.setAssigns(p=>p.map(x=>x.id===a.id?{...x,totalPeriods:parseInt(n)||1}:x));st("แก้ไขสำเร็จ")}}} style={{background:"none",border:"none",cursor:"pointer",color:"#2563EB"}}><Icon name="edit" size={14}/></button>
              <button onClick={()=>{
                if(!window.confirm("ลบวิชานี้?\n\n⚠️ คาบที่ลงตารางไว้จะถูกลบออกด้วย"))return;
                // ลบ assignment
                U.setAssigns(p=>p.filter(x=>x.id!==a.id));
                // ลบ schedule entries ที่ผูกกับ assignment นี้ด้วย
                U.setSchedule(prev=>{
                  const next={};
                  Object.entries(prev).forEach(([k,en])=>{
                    const filtered=(en||[]).filter(e=>e.assignmentId!==a.id);
                    if(filtered.length) next[k]=filtered;
                  });
                  return next;
                });
                st("ลบแล้ว (รวมคาบในตาราง)","warning");
              }} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444"}}><Icon name="trash" size={14}/></button>
            </div>
          </div>;})()}
          <div style={{display:"flex",gap:6,marginTop:10,flexWrap:"wrap"}}>{a.roomIds.map(rid=><span key={rid} style={{background:"#DBEAFE",color:"#1E40AF",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>{S.rooms.find(r=>r.id===rid)?.name}</span>)}</div>
        </div>})}
        {coAsgnsA.length>0&&<>
          <div style={{gridColumn:"1/-1",fontSize:12,fontWeight:700,color:"#7C3AED",marginTop:4,marginBottom:-8}}>👥 วิชาที่เป็นครูร่วม</div>
          {coAsgnsA.map(a=>{const sub=S.subjects.find(s=>s.id===a.subjectId);const dept=S.depts.find(d=>d.id===sub?.departmentId);const c=dept?gc(dept.id):{bg:"#7C3AED",lt:"#F5F3FF",tx:"#5B21B6"};const ca=sub?.consecutiveAllowed||0;const mainT=S.teachers.find(t=>t.id===a.teacherId);return<div key={a.id} style={{background:"#F5F3FF",borderRadius:14,borderLeft:"4px solid #7C3AED",padding:16,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
            <div style={{fontSize:10,color:"#7C3AED",fontWeight:700,marginBottom:6}}>👥 ครูร่วม (ของ {mainT?.prefix}{mainT?.firstName} {mainT?.lastName})</div>
            <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
              <h4 style={{fontSize:15,fontWeight:700,color:"#5B21B6"}}>{sub?.code} — {subDisplayName(sub)}</h4>
              {ca===-1&&<span style={{fontSize:9,background:"#EFF6FF",color:"#1E40AF",padding:"1px 6px",borderRadius:8,fontWeight:700}}>🔀NP</span>}
              {ca===-2&&<span style={{fontSize:9,background:"#FDF4FF",color:"#6B21A8",padding:"1px 6px",borderRadius:8,fontWeight:700}}>🏛️เศรษฐ-วิศวะ</span>}
              {ca>0&&<span style={{fontSize:9,background:"#FEF3C7",color:"#92400E",padding:"1px 6px",borderRadius:8,fontWeight:700}}>⚡{ca}ติด</span>}
            </div>
            <div style={{fontSize:12,color:"#6B7280",marginTop:4}}>{a.totalPeriods} คาบ/สัปดาห์</div>
            <div style={{display:"flex",gap:6,marginTop:10,flexWrap:"wrap"}}>{a.roomIds.map(rid=><span key={rid} style={{background:"#EDE9FE",color:"#5B21B6",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>{S.rooms.find(r=>r.id===rid)?.name}</span>)}</div>
          </div>})}
        </>}
      </div>
    </div>}
    {teacher&&<PersonalLockPanel teacher={teacher} U={U} st={st} sel={sel}/>}
    <Modal open={modal} onClose={()=>{setModal(false);setBasket([]);}} title={`มอบหมายวิชา — ${teacher?.prefix||""}${teacher?.firstName||""}`}>
      <div style={{display:"flex",flexDirection:"column",gap:14}}>

        {/* ── ตะกร้าวิชาที่เพิ่มแล้ว ── */}
        {basket.length>0&&(
          <div style={{background:"#F0FDF4",border:"1.5px solid #BBF7D0",borderRadius:12,padding:"10px 14px"}}>
            <div style={{fontSize:12,fontWeight:700,color:"#065F46",marginBottom:8}}>
              🛒 วิชาที่รอบันทึก ({basket.length} รายการ)
            </div>
            {basket.map((b,bi)=>{
              const bs=S.subjects.find(s=>s.id===b.subjectId);
              return(
                <div key={bi} style={{display:"flex",alignItems:"center",gap:8,marginBottom:4,background:"#fff",borderRadius:8,padding:"5px 10px",border:"1px solid #D1FAE5"}}>
                  <div style={{flex:1,fontSize:12}}>
                    <span style={{fontWeight:700,color:"#065F46"}}>{bs?.code}</span>
                    <span style={{color:"#374151",marginLeft:6}}>{bs?.name}</span>
                    <span style={{color:"#9CA3AF",marginLeft:6,fontSize:11}}>
                      {b.roomIds.map(rid=>S.rooms.find(r=>r.id===rid)?.name).join(", ")}
                      {b.totalPeriods>0?` · ${b.totalPeriods} คาบ`:""}
                    </span>
                  </div>
                  <button onClick={()=>setBasket(p=>p.filter((_,i)=>i!==bi))}
                    style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",fontSize:14,padding:0,flexShrink:0}}>✕</button>
                </div>
              );
            })}
          </div>
        )}

        {/* ── ฟอร์มเพิ่มวิชาใหม่ ── */}
        <div style={{background:"#F9FAFB",borderRadius:12,padding:"14px 16px",border:"1px solid #E5E7EB"}}>
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
              ? <button onClick={()=>{setModalDeptFilter("");setForm(p=>({...p,subjectId:"",roomIds:[]}));}}
                  style={{fontSize:10,padding:"2px 10px",borderRadius:20,border:"1.5px solid #7C3AED",background:"#F5F3FF",color:"#5B21B6",cursor:"pointer",fontWeight:600}}>
                  📚 สาระอื่น
                </button>
              : <button onClick={()=>{setModalDeptFilter(teacher?.departmentId||"");setForm(p=>({...p,subjectId:"",roomIds:[]}));}}
                  style={{fontSize:10,padding:"2px 10px",borderRadius:20,border:"1.5px solid #DC2626",background:"#FEF2F2",color:"#991B1B",cursor:"pointer",fontWeight:600}}>
                  ⭐ สาระหลัก
                </button>
            }
          </div>

          {/* pills กลุ่มสาระอื่น */}
          {modalDeptFilter!==teacher?.departmentId&&(
            <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:8}}>
              {[{id:"",name:"ทั้งหมด"},...S.depts.filter(d=>d.id!==teacher?.departmentId)].map(d=>(
                <button key={d.id}
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
                  <button key={rm.id}
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
              <input type="number" min="0" style={{...IS,width:90}} value={form.totalPeriods}
                onChange={e=>setForm(p=>({...p,totalPeriods:parseInt(e.target.value)||0}))}/>
            </div>
          )}

          {/* ปุ่ม + เพิ่มใส่ตะกร้า */}
          <button
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
          <button onClick={()=>{setModal(false);setBasket([]);}} style={{...BO(),flex:1}}>ยกเลิก</button>
          <button
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

  return <div style={{animation:"fadeIn 0.3s"}}>
    {/* filter level */}
    <div style={{display:"flex",gap:8,marginBottom:20,flexWrap:"wrap",alignItems:"center"}}>
      <span style={{fontSize:13,fontWeight:600,color:"#374151"}}>แสดงระดับชั้น:</span>
      {[{id:"",name:"ทั้งหมด"},...S.levels].map(lv=>(
        <button key={lv.id}
          onClick={()=>setFilterLevel(lv.id)}
          style={{padding:"5px 14px",borderRadius:20,border:`2px solid ${filterLevel===lv.id?"#DC2626":"#E5E7EB"}`,background:filterLevel===lv.id?"#DC2626":"#fff",color:filterLevel===lv.id?"#fff":"#374151",fontSize:12,fontWeight:filterLevel===lv.id?700:400,cursor:"pointer"}}>
          {lv.name}
        </button>
      ))}
    </div>

    {/* ตาราง */}
    <div style={{background:"#fff",borderRadius:14,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",overflow:"hidden"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
        <thead>
          <tr style={{background:"#F9FAFB"}}>
            <th style={{padding:"12px 16px",textAlign:"left",fontWeight:700,color:"#374151",borderBottom:"2px solid #E5E7EB",width:100}}>ระดับชั้น</th>
            <th style={{padding:"12px 16px",textAlign:"left",fontWeight:700,color:"#374151",borderBottom:"2px solid #E5E7EB",width:120}}>ห้อง</th>
            <th style={{padding:"12px 16px",textAlign:"left",fontWeight:700,color:"#374151",borderBottom:"2px solid #E5E7EB"}}>ครูประจำชั้นหลัก 1</th>
            <th style={{padding:"12px 16px",textAlign:"left",fontWeight:700,color:"#374151",borderBottom:"2px solid #E5E7EB"}}>ครูประจำชั้นหลัก 2</th>
            <th style={{padding:"12px 16px",textAlign:"left",fontWeight:700,color:"#374151",borderBottom:"2px solid #E5E7EB"}}>ครูประจำชั้นร่วม</th>
            <th style={{padding:"12px 8px",textAlign:"center",fontWeight:700,color:"#374151",borderBottom:"2px solid #E5E7EB",width:80}}></th>
          </tr>
        </thead>
        <tbody>
          {sorted.map((rm,i)=>{
            const lv=S.levels.find(l=>l.id===rm.levelId);
            const isEdit=editId===rm.id;
            return(
              <tr key={rm.id} style={{borderBottom:"1px solid #F3F4F6",background:isEdit?"#FFF7ED":i%2===0?"#fff":"#FAFAFA"}}>
                <td style={{padding:"10px 16px",fontWeight:600,color:"#6B7280",fontSize:12}}>{lv?.name||""}</td>
                <td style={{padding:"10px 16px",fontWeight:700,color:"#1E40AF"}}>{rm.name}</td>
                {isEdit?(
                  <>
                    <td style={{padding:"6px 10px"}}>
                      <SearchSelect value={form.homeroom1} onChange={v=>setForm(p=>({...p,homeroom1:v}))} options={teacherOptions} placeholder="-- เลือกครู --"/>
                    </td>
                    <td style={{padding:"6px 10px"}}>
                      <SearchSelect value={form.homeroom2} onChange={v=>setForm(p=>({...p,homeroom2:v}))} options={teacherOptions} placeholder="-- เลือกครู --"/>
                    </td>
                    <td style={{padding:"6px 10px"}}>
                      <SearchSelect value={form.homeroomCo} onChange={v=>setForm(p=>({...p,homeroomCo:v}))} options={teacherOptions} placeholder="-- เลือกครู --"/>
                    </td>
                    <td style={{padding:"6px 8px",textAlign:"center"}}>
                      <div style={{display:"flex",gap:4,justifyContent:"center"}}>
                        <button onClick={save} style={{...BS(),fontSize:11,padding:"4px 12px"}}>บันทึก</button>
                        <button onClick={()=>setEditId(null)} style={{...BO(),fontSize:11,padding:"4px 10px"}}>ยกเลิก</button>
                      </div>
                    </td>
                  </>
                ):(
                  <>
                    <td style={{padding:"10px 16px",color:rm.homeroom1?"#111":"#9CA3AF",fontSize:12}}>{rm.homeroom1||"—"}</td>
                    <td style={{padding:"10px 16px",color:rm.homeroom2?"#111":"#9CA3AF",fontSize:12}}>{rm.homeroom2||"—"}</td>
                    <td style={{padding:"10px 16px",color:rm.homeroomCo?"#111":"#9CA3AF",fontSize:12}}>{rm.homeroomCo||"—"}</td>
                    <td style={{padding:"10px 8px",textAlign:"center"}}>
                      <button onClick={()=>openEdit(rm)} style={{...BO("#2563EB"),fontSize:11,padding:"4px 12px"}}><Icon name="edit" size={12}/>แก้ไข</button>
                    </td>
                  </>
                )}
              </tr>
            );
          })}
          {!sorted.length&&<tr><td colSpan={6} style={{padding:30,textAlign:"center",color:"#9CA3AF"}}>ยังไม่มีห้องเรียน</td></tr>}
        </tbody>
      </table>
    </div>

    {/* ปุ่มรีเซ็ตทั้งหมด */}
    <div style={{marginTop:16}}>
      <button onClick={()=>{
        if(!window.confirm("รีเซ็ตครูประจำชั้นทุกห้อง?"))return;
        U.setRooms(p=>p.map(r=>({...r,homeroom1:"",homeroom2:"",homeroomCo:""})));
        st("รีเซ็ตแล้ว","warning");
      }} style={{...BO("#DC2626"),fontSize:12}}>
        🔄 รีเซ็ตครูประจำชั้นทุกห้อง
      </button>
    </div>
  </div>;
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
      <button style={TAB_STYLE(tab==="dept")} onClick={()=>setTab("dept")}>🔒 คาบล็อคกลุ่มสาระ (เดิม)</button>
      <button style={TAB_STYLE(tab==="custom")} onClick={()=>setTab("custom")}>📅 คาบล็อคแผนก (หลายวัน)</button>
    </div>

    {/* ── Tab 1: คาบล็อคกลุ่มสาระ เดิม ── */}
    {tab==="dept"&&<>
      <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",marginBottom:24,maxWidth:600}}>
        <h3 style={{fontSize:16,fontWeight:700,marginBottom:16}}>เพิ่มคาบล็อคกลุ่มสาระ</h3>
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          <div><label style={LS}>กลุ่มสาระ</label>
            <SearchSelect value={deptForm.departmentId} onChange={v=>setDeptForm(p=>({...p,departmentId:v}))} options={[{value:"",label:"--"},...S.depts.map(d=>({value:d.id,label:d.name}))]} placeholder="-- เลือกกลุ่มสาระ --"/>
          </div>
          <div><label style={LS}>วัน</label>
            <select style={IS} value={deptForm.day} onChange={e=>setDeptForm(p=>({...p,day:e.target.value}))}>
              <option value="">--</option>{DAYS.map(d=><option key={d}>{d}</option>)}
            </select>
          </div>
          <div><label style={LS}>คาบ</label>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              {PERIODS.map(p=><button key={p.id}
                onClick={()=>setDeptForm(prev=>({...prev,periods:prev.periods.includes(p.id)?prev.periods.filter(x=>x!==p.id):[...prev.periods,p.id]}))}
                style={{width:48,height:48,borderRadius:10,border:`2px solid ${deptForm.periods.includes(p.id)?"#DC2626":"#D1D5DB"}`,background:deptForm.periods.includes(p.id)?"#DC2626":"#fff",color:deptForm.periods.includes(p.id)?"#fff":"#374151",fontSize:16,fontWeight:700,cursor:"pointer"}}>
                {p.id}
              </button>)}
            </div>
          </div>
          <button onClick={saveDept} style={BS()}>เพิ่มคาบล็อค</button>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:16}}>
        {deptMeetings.map(m=>{
          const dept=S.depts.find(d=>d.id===m.departmentId);
          const c=dept?gc(dept.id):{bg:"#6B7280"};
          return<div key={m.id} style={{background:"#fff",borderRadius:14,borderLeft:`4px solid ${c.bg}`,padding:16,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
            <div style={{display:"flex",justifyContent:"space-between"}}>
              <div>
                <h4 style={{fontSize:15,fontWeight:700}}>{dept?.name}</h4>
                <div style={{fontSize:13,color:"#6B7280",marginTop:4}}>วัน{m.day} — คาบ {(m.periods||[]).slice().sort().join(", ")}</div>
              </div>
              <button onClick={()=>{U.setMeetings(p=>p.filter(x=>x.id!==m.id));st("ลบแล้ว","warning")}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444"}}><Icon name="trash" size={14}/></button>
            </div>
          </div>;
        })}
      </div>
    </>}

    {/* ── Tab 2: คาบล็อคแผนก หลายวันหลายคาบ ── */}
    {tab==="custom"&&<>
      <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",marginBottom:24}}>
        <h3 style={{fontSize:16,fontWeight:700,marginBottom:16}}>เพิ่มคาบล็อคแผนก</h3>
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          {/* ชื่อ */}
          <div>
            <label style={LS}>ชื่อคาบล็อค</label>
            <input style={{...IS,maxWidth:400}} value={cusForm.name} onChange={e=>setCusForm(p=>({...p,name:e.target.value}))} placeholder="เช่น ประชุมวิชาการ, อบรม, สอบกลางภาค"/>
          </div>

          {/* ตาราง grid วัน × คาบ เลือกได้หลายช่อง */}
          <div>
            <label style={LS}>เลือกวัน × คาบ (คลิกเพื่อเลือก/ยกเลิก)</label>
            <div style={{overflowX:"auto"}}>
              <table style={{borderCollapse:"collapse",minWidth:500}}>
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
                <button onClick={()=>setCusForm(p=>({...p,slots:[]}))} style={{marginLeft:8,fontSize:11,color:"#EF4444",background:"none",border:"none",cursor:"pointer"}}>ล้างทั้งหมด</button>
              </div>
            )}
          </div>
          <button onClick={saveCustom} style={BS()}>เพิ่มคาบล็อค</button>
        </div>
      </div>

      {/* รายการ custom locks */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(320px,1fr))",gap:16}}>
        {customMeetings.map(m=>{
          const slotsByDay=DAYS.map(day=>{
            const ps=(m.slots||[]).filter(s=>s.day===day).map(s=>s.period).sort((a,b)=>a-b);
            return ps.length?{day,periods:ps}:null;
          }).filter(Boolean);
          return<div key={m.id} style={{background:"#fff",borderRadius:14,borderLeft:"4px solid #DC2626",padding:16,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
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
              <button onClick={()=>{U.setMeetings(p=>p.filter(x=>x.id!==m.id));st("ลบแล้ว","warning")}} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",flexShrink:0}}><Icon name="trash" size={14}/></button>
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
  return <div style={{background:"#fff",borderRadius:14,padding:60,textAlign:"center"}}>
    <div style={{fontSize:48,marginBottom:16}}>{icon}</div>
    <h3 style={{fontSize:18,fontWeight:700,color:"#374151"}}>{title}</h3>
  </div>;
}

/* ===== SCHEDULER ENTRY CARD (top-level เพื่อกัน React recreate) ===== */
function SchedulerEntryCard({entry,cellKey,lk,cellCount,selT,mode,S,U,gc,setDrag,setCoM}){
  const [showActions,setShowActions]=useState(false);
  const sub=S.subjects.find(s=>s.id===entry.subjectId);
  const dept=S.depts.find(d=>d.id===sub?.departmentId);
  const c=dept?gc(dept.id):{bg:"#6B7280",lt:"#F3F4F6",tx:"#374151",bd:"#D1D5DB"};
  const et=S.teachers.find(t=>t.id===entry.teacherId);
  const coIds=entry.coTeacherIds?.length?entry.coTeacherIds:(entry.coTeacherId?[entry.coTeacherId]:[]);
  const coTeachers=coIds.map(id=>S.teachers.find(t=>t.id===id)).filter(Boolean);
  const isOwn=entry.teacherId===selT||coIds.includes(selT);
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
      onMouseEnter={()=>setShowActions(true)}
      onMouseLeave={()=>setShowActions(false)}
      style={{
        background:dimmed?"#F9FAFB":c.lt,
        border:"2px solid "+(dimmed?"#E5E7EB":c.bd),
        borderRadius:8,
        padding:compact?"3px 22px 3px 5px":"6px 8px",
        marginBottom:2,
        fontSize:11,
        position:"relative",
        cursor:lk||dimmed?"default":"grab",
        opacity:dimmed?0.4:1,
        transition:"opacity 0.2s,box-shadow 0.15s",
        userSelect:"none",
        boxShadow:dimmed?"none":"0 1px 3px rgba(0,0,0,0.08)",
      }}
    >
      {compact
        ?<>
            <div style={{fontWeight:700,color:dimmed?"#9CA3AF":c.tx,fontSize:10,lineHeight:1.3,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
              {subDisplayName(sub)||sub?.code}
            </div>
            {/* ชื่อครู ใน compact */}
            {et&&<div style={{fontSize:9,color:dimmed?"#9CA3AF":c.tx,opacity:0.75,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{et.firstName}</div>}
            {/* action buttons สำหรับ compact — แสดงเมื่อ hover */}
            {!lk&&(
              <div style={{position:"absolute",top:1,right:1,display:"flex",gap:1,opacity:showActions?1:0,transition:"opacity 0.15s"}}>
                <button onMouseDown={e=>{e.stopPropagation();e.preventDefault();setCoM({key:cellKey,entryId:entry.id});}} style={{background:"rgba(255,255,255,0.9)",border:"none",cursor:"pointer",color:"#2563EB",padding:"1px 2px",lineHeight:1,borderRadius:3}}><Icon name="users" size={9}/></button>
                <button onMouseDown={e=>{e.stopPropagation();e.preventDefault();removeEntry();}} style={{background:"rgba(255,255,255,0.9)",border:"none",cursor:"pointer",color:"#EF4444",padding:"1px 2px",lineHeight:1,borderRadius:3}}><Icon name="x" size={9}/></button>
                <button onMouseDown={e=>{e.stopPropagation();e.preventDefault();lockEntry();}} style={{background:"rgba(255,255,255,0.9)",border:"none",cursor:"pointer",color:"#059669",padding:"1px 2px",lineHeight:1,borderRadius:3}}><Icon name="lock" size={9}/></button>
              </div>
            )}
          </>
        :<>
            <div style={{fontWeight:700,color:dimmed?"#9CA3AF":c.tx,fontSize:11}}>{sub?.code}</div>
            <div style={{fontWeight:600,color:dimmed?"#9CA3AF":c.tx,fontSize:10}}>{subDisplayName(sub)}</div>
            <div style={{color:dimmed?"#9CA3AF":c.tx,opacity:0.7,fontSize:10}}>
              {et?.firstName}{coTeachers.length>0?" + "+coTeachers.map(t=>t.firstName).join(", "):""}
            </div>
          </>
      }
      {/* action buttons สำหรับ non-compact */}
      {!lk&&!compact&&(
        <div style={{display:"flex",gap:3,marginTop:3}}>
          <button onClick={removeEntry} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",padding:0,lineHeight:1}}><Icon name="x" size={10}/></button>
          <button onClick={()=>setCoM({key:cellKey,entryId:entry.id})} style={{background:"none",border:"none",cursor:"pointer",color:"#2563EB",padding:0,lineHeight:1}}><Icon name="users" size={10}/></button>
          <button onClick={lockEntry} style={{background:"none",border:"none",cursor:"pointer",color:"#059669",padding:0,lineHeight:1}}><Icon name="lock" size={10}/></button>
        </div>
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
function Scheduler({S,U,st,gc,isSavingRef,fsReadyRef,fsSave}){
  const [mode,setMode]=useState("teacher");
  const [selDept,setSelDept]=useState("");
  const [selT,setSelT]=useState("");
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
          U.setSchedule(bestResult.schedule);
          setAutoResult({
            placed: bestResult.placed,
            skipped: bestResult.skipped,
            details: bestResult.details,
            runs: opts.runs,
          });
          setAutoRunning(false);
          setAutoProgress(null);
          st(`Auto จัด (${opts.runs} รอบ): วาง ${bestResult.placed} คาบ, ข้าม ${bestResult.skipped} คาบ`, "success");
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
      const room=S.rooms.find(r=>r.id===rid);
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
    const room=S.rooms.find(r=>r.id===rid);
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
    if(room&&sub&&room.levelId!==sub.levelId){
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
          <div key={rid} style={{marginBottom:28}}>
            <div style={{marginBottom:8,display:"flex",alignItems:"center",gap:8}}>
              <span style={{background:lc.head,color:"#fff",padding:"5px 16px",borderRadius:10,fontSize:12,fontWeight:700,letterSpacing:"0.02em",boxShadow:`0 2px 6px ${lc.head}55`}}>{rm?.name}</span>
              {rmPlan&&<span style={{background:lc.bg,color:lc.head,border:`1.5px solid ${lc.border}`,padding:"4px 12px",borderRadius:20,fontSize:11,fontWeight:700}}>{rmPlan.name}</span>}
              {rmLevel&&<span style={{color:"#9CA3AF",fontSize:11}}>{rmLevel.name}</span>}
            </div>
            <div style={{background:"#fff",borderRadius:14,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",overflow:"hidden",border:`1px solid ${lc.border}`}}>
              <table style={{width:"100%",borderCollapse:"collapse",tableLayout:"fixed",minWidth:700}}>
                <thead>
                  <tr style={{borderBottom:`2px solid ${lc.head}`}}>
                    <th style={{padding:"10px 10px",background:lc.head,color:"#fff",width:62,textAlign:"left",fontSize:13,fontWeight:700,letterSpacing:"0.02em"}}>วัน</th>
                    {PERIODS.map(p=>(
                      <th key={p.id} style={{padding:"6px 2px",background:lc.head,textAlign:"center",borderLeft:"1px solid rgba(255,255,255,0.2)"}}>
                        <div style={{fontSize:11,color:"#fff",fontWeight:700}}>คาบ {p.id}</div>
                        <div style={{fontSize:9,color:"rgba(255,255,255,0.7)",fontWeight:400}}>{p.time}</div>
                      </th>
                    ))}
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
                        return (
                          <td key={p.id}
                            className="dz"
                            onDragOver={e=>{const d=dragRef.current;if(!d){e.currentTarget.classList.remove("over");return;}
if(d.fromRoomId&&d.fromRoomId!==rid){e.currentTarget.classList.remove("over");return;}
if(d.assignmentId){const a=S.assigns.find(x=>x.id===d.assignmentId);const sCa=S.subjects.find(s=>s.id===d.subjectId)?.consecutiveAllowed||0;const ok=a?.roomIds?.includes(rid)||(sCa===-2&&S.assigns.some(x=>x.subjectId===d.subjectId&&x.roomIds?.includes(rid)));if(!ok){e.currentTarget.classList.remove("over");return;}}
e.preventDefault();e.currentTarget.classList.add("over");}}
                            onDragLeave={e=>e.currentTarget.classList.remove("over")}
                            onDrop={e=>{e.preventDefault();e.currentTarget.classList.remove("over");handleDrop(rid,day,p.id);}}
                            style={{padding:3,verticalAlign:"top",minHeight:68,borderLeft:`1px solid ${lc.border}`,borderBottom:`1px solid ${lc.border}`,background:bl?"#FEF9C3":lk?"#F0FDF4":"inherit"}}
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
                  );})}
                </tbody>
              </table>
            </div>
          </div>
        );
      })}
    </div>
  );

  /* ── ตารางสรุปสัปดาห์ครู (sticky bottom) ── */
  const renderTeacherWeeklySummary=()=>{
    if(!selT||mode!=="teacher") return null;
    const teacher=S.teachers.find(t=>t.id===selT);
    const totalUsed=teacherScheduledTotal(selT);
    const quota=teacher?.totalPeriods||0;
    return (
      <div style={{position:"fixed",bottom:16,left:"calc(240px + 16px)",zIndex:200,fontFamily:"'Sarabun','Noto Sans Thai',sans-serif"}}>
        {/* Compact pill — แสดงตลอด */}
        <div
          onClick={()=>setShowWeekly(v=>!v)}
          style={{display:"inline-flex",alignItems:"center",gap:8,padding:"7px 14px",background:"rgba(30,58,95,0.95)",backdropFilter:"blur(8px)",borderRadius:showWeekly?"12px 12px 0 0":12,cursor:"pointer",userSelect:"none",boxShadow:"0 4px 16px rgba(0,0,0,0.25)",width:230,boxSizing:"border-box"}}
        >
          <span style={{fontSize:12,fontWeight:700,color:"#fff",whiteSpace:"nowrap"}}>📋 {teacher?.prefix}{teacher?.firstName} {teacher?.lastName}</span>
          <span style={{fontSize:12,background:totalUsed>=quota?"#D1FAE5":"#FEF3C7",color:totalUsed>=quota?"#065F46":"#92400E",padding:"2px 10px",borderRadius:20,fontWeight:800,flexShrink:0}}>
            {totalUsed}/{quota} คาบ {totalUsed>=quota?"✓":""}
          </span>
          <span style={{fontSize:11,color:"rgba(255,255,255,0.75)"}}>{showWeekly?"▼ ซ่อน":"▲ แสดง"}</span>
        </div>
        {/* ตารางสรุป — expand ขึ้นข้างบน */}
        {showWeekly&&(
          <div style={{maxHeight:"50vh",overflowY:"auto",background:"#fff",borderRadius:"0 12px 0 0",boxShadow:"0 -4px 20px rgba(0,0,0,0.18)",border:"1px solid #BFDBFE",borderBottom:"none",width:"calc(100vw - 296px)",maxWidth:900,position:"absolute",bottom:"100%",left:0}}>
            <div style={{overflow:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",tableLayout:"fixed",minWidth:680}}>
                <thead>
                  <tr>
                    <th style={{padding:"7px 10px",background:"#1E3A5F",color:"#fff",width:72,textAlign:"left",fontSize:12,fontWeight:700,position:"sticky",top:0,zIndex:2}}>วัน</th>
                    {PERIODS.map(p=>(
                      <th key={p.id} style={{padding:"5px 3px",background:"#1E3A5F",textAlign:"center",borderLeft:"1px solid rgba(255,255,255,0.15)",position:"sticky",top:0,zIndex:2}}>
                        <div style={{fontSize:12,color:"#fff",fontWeight:700}}>คาบ {p.id}</div>
                        <div style={{fontSize:9,color:"rgba(255,255,255,0.65)"}}>{p.time}</div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {DAYS.map((day,di)=>(
                    <tr key={day} style={{background:di%2===0?"#FFFFFF":"#F0F7FF",borderBottom:"1px solid #E0EEFF"}}>
                      <td style={{padding:"7px 10px",fontWeight:700,fontSize:12,color:"#1E3A5F",borderRight:"2px solid #BFDBFE",background:"#EFF6FF"}}>{day}</td>
                  {PERIODS.map(p=>{
                    const blk=isBlk(selT,day,p.id);
                    // หาทุกห้องที่ครูสอนในคาบนี้
                    const roomsThisPeriod=[];
                    Object.entries(S.schedule).forEach(([k,en])=>{
                      if(!k.endsWith("_"+day+"_"+p.id)) return;
                      en?.forEach(e=>{
                        const coIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
                        if(e.teacherId===selT||coIds.includes(selT)){
                          const pts=k.split("_");
                          const rmId=pts.slice(0,pts.length-2).join("_");
                          const rm=S.rooms.find(r=>r.id===rmId);
                          const sub=S.subjects.find(s=>s.id===e.subjectId);
                          if(!roomsThisPeriod.find(x=>x.rmId===rmId))
                            roomsThisPeriod.push({rmId,rmName:rm?.name||"?",subName:subDisplayName(sub)||"?"});
                        }
                      });
                    });
                    return (
                      <td key={p.id} style={{textAlign:"center",padding:"5px 3px",borderLeft:"1px solid #F0F0F0",verticalAlign:"middle",minHeight:48}}>
                        {blk
                          ? <div style={{background:"#FEF9C3",color:"#92400E",fontSize:10,borderRadius:6,padding:"3px 5px",fontWeight:700}}>
                              🔒{S.meetings.some(m=>m.day===day&&m.periods?.includes(p.id)&&m.departmentId===teacher?.departmentId)?"ประชุม":blocked(selT).find(b=>b.day===day&&b.period===p.id)?.reason||"ล็อค"}
                            </div>
                          : roomsThisPeriod.length>0
                            ? roomsThisPeriod.map((r,i)=>(
                                <div key={i} style={{background:"#FEF2F2",border:"1px solid #FECACA",borderRadius:6,padding:"4px 6px",marginBottom:i<roomsThisPeriod.length-1?2:0}}>
                                  <div style={{fontSize:11,fontWeight:800,color:"#1E40AF"}}>{r.rmName}</div>
                                  <div style={{fontSize:10,color:"#374151",fontWeight:600}}>{r.subName}</div>
                                </div>
                              ))
                            : <span style={{color:"#D1D5DB",fontSize:12}}>—</span>
                        }
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
            </div>
          </div>
        )}
      </div>
    );
  };

  /* ── render ── */
  return (
    <div style={{animation:"fadeIn 0.3s"}}>

      {/* Mode + selector bar */}
      <div style={{display:"flex",gap:8,marginBottom:14,alignItems:"center",flexWrap:"wrap"}}>
        <div style={{display:"flex",borderRadius:10,overflow:"hidden",border:"1.5px solid "+CRED,boxShadow:"0 2px 8px rgba(185,28,28,0.15)"}}>
          <button onClick={()=>{setMode("teacher");setSelRoom("");}} style={{padding:"8px 20px",background:mode==="teacher"?CRED:"#fff",color:mode==="teacher"?"#fff":CRED,border:"none",fontWeight:700,fontSize:13,cursor:"pointer",transition:"background 0.15s"}}>จัดรายครู</button>
          <button onClick={()=>{setMode("room");setSelT("");setSelDept("");}} style={{padding:"8px 20px",background:mode==="room"?CRED:"#fff",color:mode==="room"?"#fff":CRED,border:"none",fontWeight:700,fontSize:13,cursor:"pointer",transition:"background 0.15s"}}>จัดรายห้อง</button>
        </div>

        {mode==="teacher"&&<>
          <SearchSelect value={selDept} onChange={v=>{setSelDept(v);setSelT("");}} options={[{value:"",label:"-- ทุกกลุ่มสาระ --"},...S.depts.map(d=>({value:d.id,label:d.name}))]} placeholder="-- ทุกกลุ่มสาระ --" style={{maxWidth:200}}/>
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
        {/* Auto Schedule + ล้างคาบกำพร้า */}
        <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center",flexShrink:0}}>
          <button onClick={async()=>{
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
            if(!window.confirm(`พบ ${removed} คาบกำพร้า\nลบออกทั้งหมดไหม?`))return;

            if(isSavingRef) isSavingRef.current=true;
            if(fsReadyRef)  fsReadyRef.current=false;
            U.setSchedule(next);
            try{
              if(fsSave) await fsSave(next);
              st(`ลบ ${removed} คาบกำพร้าแล้ว ✅`,"warning");
            }catch(e){
              st("ลบ local แล้ว แต่ save cloud ล้มเหลว","error");
            }finally{
              // unlock หลัง save เสร็จแน่ๆ + รอ onSnapshot ผ่านไป 1 รอบ
              setTimeout(()=>{
                if(fsReadyRef)  fsReadyRef.current=true;
                if(isSavingRef) isSavingRef.current=false;
              },1500);
            }
          }} style={{...BO("#DC2626"),fontSize:12,padding:"7px 12px",whiteSpace:"nowrap",flexShrink:0}}>
            🧹 ล้างคาบกำพร้า
          </button>
          <button onClick={runAutoSchedule} disabled={autoRunning}
            style={{...BS("#059669"),opacity:autoRunning?0.6:1,position:"relative",minWidth:160}}>
            {autoRunning
              ? <span style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{display:"inline-block",width:14,height:14,border:"2px solid rgba(255,255,255,0.4)",borderTopColor:"#fff",borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/>
                  รอบ {autoProgress?.run||0}/{autoProgress?.total||10}...
                </span>
              : "⚡ Auto จัดตาราง"
            }
          </button>
        </div>
      </div>

      {/* Auto result panel */}
      {autoResult&&(
        <div style={{background:autoResult.skipped===0?"#F0FDF4":"#FFFBEB",border:`1.5px solid ${autoResult.skipped===0?"#86EFAC":"#FDE68A"}`,borderRadius:12,padding:"12px 16px",marginBottom:12,display:"flex",gap:16,alignItems:"flex-start",flexWrap:"wrap"}}>
          <div style={{fontSize:13,fontWeight:700,color:autoResult.skipped===0?"#065F46":"#92400E",display:"flex",gap:12,flexWrap:"wrap",alignItems:"center"}}>
            {autoResult.skipped===0?"✅":"⚠️"}
            <span>จัดด้วย <strong>{autoResult.runs} รอบ</strong> — วาง <strong>{autoResult.placed}</strong> คาบ</span>
            {autoResult.skipped>0&&<span style={{color:"#DC2626"}}>| ข้ามไม่ได้ <strong>{autoResult.skipped}</strong> คาบ</span>}
          </div>
          {autoResult.details.length>0&&(
            <div style={{fontSize:11,color:"#92400E",flex:1}}>
              ❌ ไม่สามารถจัดได้: {autoResult.details.slice(0,5).join(", ")}{autoResult.details.length>5?` และอีก ${autoResult.details.length-5} รายการ`:""}
            </div>
          )}
          <button onClick={()=>setAutoResult(null)} style={{background:"none",border:"none",cursor:"pointer",color:"#9CA3AF",fontSize:16}}>✕</button>
        </div>
      )}

      {/* Teacher summary bar */}
      {mode==="teacher"&&teacher&&(
        <div style={{background:CBGW,borderRadius:14,padding:"12px 18px",marginBottom:12,display:"flex",gap:12,alignItems:"center",flexWrap:"wrap",boxShadow:"0 2px 12px rgba(0,0,0,0.06)",border:"1px solid #F0F0F0"}}>
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
        ?<div style={{display:"flex",flexDirection:"column",gap:0}}><div style={{display:"flex",gap:14}}>
            {/* Sidebar */}
            <div style={{width:200,flexShrink:0,position:"sticky",top:0,alignSelf:"flex-start",maxHeight:"calc(100vh - 200px)",overflowY:"auto"}}>
              <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:8}}>วิชา — ลากวาง</div>
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
                const coTeachers2=coIds2.map(id=>S.teachers.find(t=>t.id===id)).filter(Boolean);
                const buns=bundleMap[a.id]||[];
                return (
                  <div key={a.id} style={{background:rem<=0?"#F3F4F6":lc.bg,border:`1.5px solid ${rem<=0?"#D1D5DB":lc.border}`,borderRadius:12,padding:"10px 12px",marginBottom:10,boxShadow:rem<=0?"none":`0 2px 8px ${lc.head}22`,transition:"all 0.2s",position:"relative"}}>
                    {/* ปุ่ม ⚙️ settings มุมขวาบน */}
                    <button
                      onClick={()=>setShowGearId(showGearId===a.id?null:a.id)}
                      title="ครูร่วม / วิชาคู่"
                      style={{position:"absolute",top:6,right:6,background:"rgba(0,0,0,0.07)",border:"none",borderRadius:6,width:22,height:22,cursor:"pointer",fontSize:12,display:"flex",alignItems:"center",justifyContent:"center",color:rem<=0?"#9CA3AF":lc.tx}}>⚙</button>

                    {coAsgnsIds.has(a.id)&&<div style={{fontSize:9,color:"#7C3AED",fontWeight:700,marginBottom:3}}>👥 ครูร่วม ({S.teachers.find(t=>t.id===a.teacherId)?.firstName||""})</div>}

                    <div
                      className="drag-card"
                      draggable={rem>0&&!coAsgnsIds.has(a.id)}
                      onDragStart={()=>setDragBoth({teacherId:selT,subjectId:a.subjectId,assignmentId:a.id})}
                      onDragEnd={()=>setDragBoth(null)}
                      style={{cursor:rem>0&&!coAsgnsIds.has(a.id)?"grab":"default",paddingRight:20}}
                    >
                      {/* ชื่อวิชา ตัวใหญ่ชัดเจน */}
                      <div style={{fontSize:13,fontWeight:800,color:rem<=0?"#9CA3AF":lc.tx,lineHeight:1.4,marginBottom:2,textDecoration:rem<=0?"line-through":"none"}}>
                        {subDisplayName(sub)||sub?.code}
                      </div>
                      <div style={{fontSize:10,color:rem<=0?"#9CA3AF":lc.head,fontWeight:700,marginBottom:4}}>{sub?.code}</div>

                      {/* badges */}
                      <div style={{display:"flex",gap:3,flexWrap:"wrap",marginBottom:5}}>
                        {sub?.consecutiveAllowed===-1&&<span style={{fontSize:8,background:"#EFF6FF",color:"#1E40AF",padding:"1px 5px",borderRadius:6,fontWeight:700}}>NP</span>}
                        {sub?.consecutiveAllowed===-2&&<span style={{fontSize:8,background:"#FDF4FF",color:"#6B21A8",padding:"1px 5px",borderRadius:6,fontWeight:700}}>เศรษฐ-วิศวะ</span>}
                        {sub?.consecutiveAllowed>0&&<span style={{fontSize:8,background:"#FEF3C7",color:"#92400E",padding:"1px 5px",borderRadius:6,fontWeight:700}}>⚡{sub.consecutiveAllowed}ติด</span>}
                        {(()=>{const sr=S.specialRooms.find(r=>r.id===sub?.specialRoomId);return sr?<span style={{fontSize:8,background:"#EDE9FE",color:"#5B21B6",padding:"1px 5px",borderRadius:6,fontWeight:700}}>📍{sr.name}</span>:null;})()}
                      </div>

                      {/* ห้องเรียน + คาบคงเหลือ */}
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                        <div style={{display:"flex",gap:3,flexWrap:"wrap"}}>
                          {a.roomIds.map(rid=>(
                            <span key={rid} style={{background:lc.head,color:"#fff",padding:"2px 7px",borderRadius:8,fontSize:10,fontWeight:700}}>{S.rooms.find(r=>r.id===rid)?.name}</span>
                          ))}
                        </div>
                        <span style={{background:rem>0?lc.head:"#9CA3AF",color:"#fff",padding:"3px 9px",borderRadius:20,fontSize:11,fontWeight:800,flexShrink:0}}>{rem}/{totalForCard}</span>
                      </div>

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
                        {coTeachers2.map((ct2)=>(
                          <div key={ct2.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3}}>
                            <span style={{fontSize:10,color:lc.tx}}>{ct2.firstName} {ct2.lastName}</span>
                            <button onClick={()=>setCardCoMap(p=>({...p,[a.id]:coIds2.filter(id=>id!==ct2.id)}))} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",padding:0,fontSize:12}}>✕</button>
                          </div>
                        ))}
                        {coTeachers2.length<4&&(
                          <button onClick={()=>{ setShowGearId(null); setCardCoM(a.id); }} style={{fontSize:10,color:lc.head,background:"rgba(0,0,0,0.06)",border:`1px solid ${lc.border}`,borderRadius:6,padding:"3px 8px",cursor:"pointer",width:"100%",textAlign:"left",marginBottom:6}}>
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
                            <button onClick={()=>setBundleMap(p=>({...p,[a.id]:buns.filter((_,i)=>i!==bi)}))} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",padding:0,fontSize:10}}>✕</button>
                          </div>;
                        })}
                        <button onClick={()=>{ setShowGearId(null); setShowBundleM(a.id); setBundleSelSub(""); setBundleSelTeacher(""); }} style={{fontSize:10,color:"#059669",background:"rgba(5,150,105,0.08)",border:"1px solid #BBF7D0",borderRadius:6,padding:"3px 8px",cursor:"pointer",width:"100%",textAlign:"left"}}>+ เพิ่มวิชาคู่</button>
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
          <button
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
          <button
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
                  <button onClick={()=>setBundleMap(p=>({...p,[showBundleM]:(p[showBundleM]||[]).filter((_,i)=>i!==bi)}))} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",fontSize:16}}>✕</button>
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
              <button
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
          <div style={{background:"#fff",borderRadius:20,boxShadow:"0 30px 60px rgba(0,0,0,0.25)",width:"min(520px,94%)",maxHeight:"90vh",display:"flex",flexDirection:"column",overflow:"hidden",fontFamily:"'Sarabun','Noto Sans Thai',sans-serif"}}>
            {/* Header */}
            <div style={{background:"linear-gradient(135deg,#991B1B,#B91C1C)",padding:"20px 24px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <div>
                <div style={{color:"#fff",fontSize:17,fontWeight:700}}>⚡ Auto จัดตารางสอน</div>
                <div style={{color:"rgba(255,255,255,0.7)",fontSize:12,marginTop:2}}>เลือกเงื่อนไขก่อนกด "เริ่มจัด"</div>
              </div>
              <button onClick={()=>setShowAutoModal(false)} style={{background:"rgba(255,255,255,0.15)",border:"none",borderRadius:8,padding:"6px 10px",cursor:"pointer",color:"#fff",fontSize:16}}>✕</button>
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
                      <input type="radio" name="autoMode" value={o.val} checked={autoOpts.mode===o.val} onChange={()=>setAutoOpts(p=>({...p,mode:o.val}))} style={{marginTop:2,accentColor:o.safe?"#059669":"#DC2626",flexShrink:0}}/>
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
                      <input type="checkbox" checked={!!autoOpts[o.key]} onChange={e=>setAutoOpts(p=>({...p,[o.key]:e.target.checked}))} style={{marginTop:2,accentColor:"#2563EB",flexShrink:0}}/>
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
                      <input type="checkbox" checked={!!autoOpts[o.key]} onChange={e=>setAutoOpts(p=>({...p,[o.key]:e.target.checked}))} style={{marginTop:2,accentColor:"#7C3AED",flexShrink:0}}/>
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
                      <select value={autoOpts.maxConsecTeacher} onChange={e=>setAutoOpts(p=>({...p,maxConsecTeacher:parseInt(e.target.value)}))} style={{padding:"6px 28px 6px 10px",border:"1.5px solid #D97706",borderRadius:8,fontSize:13,fontWeight:700,color:"#92400E",background:"#fff",cursor:"pointer",outline:"none",fontFamily:"inherit"}}>
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
                    <button key={o.val} onClick={()=>setAutoOpts(p=>({...p,runs:o.val}))} style={{flex:"1 1 80px",padding:"10px 8px",borderRadius:12,border:`2px solid ${autoOpts.runs===o.val?"#059669":"#E5E7EB"}`,background:autoOpts.runs===o.val?"#F0FDF4":"#F9FAFB",cursor:"pointer",fontFamily:"inherit"}}>
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
              <button onClick={()=>setShowAutoModal(false)} style={BO()}>ยกเลิก</button>
              <button
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
            const teacherNames=[t,...coTs].filter(Boolean).map(x=>x.firstName||'').join('+');
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

/* ===== REPORTS ===== */
function Reports({S,U,st,gc,ay,sh}){
  const fileRefSched=useRef(null);
  const [selTeacherPDF,setSelTeacherPDF]=useState("");
  const [selRoomPDF,setSelRoomPDF]=useState("");
  const [selTeacherXL,setSelTeacherXL]=useState("");
  const [selRoomXL,setSelRoomXL]=useState("");
  const [showNewRoomPDF,setShowNewRoomPDF]=useState(false);
  const [newRoomPDFOpts,setNewRoomPDFOpts]=useState({selectedRooms:[],layout:"2portrait"});
  const [showNewTeacherPDF,setShowNewTeacherPDF]=useState(false);
  const [selectedTeachersPDF,setSelectedTeachersPDF]=useState([]);
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
    const w=window.open('','_blank');
    if(!w){st("Browser บล็อก popup","error");return;}
    const pages=[];
    for(let i=0;i<list.length;i+=2) pages.push(list.slice(i,i+2));
    const pagesHTML=pages.map((pair,pi)=>`
      <div style="page-break-after:${pi<pages.length-1?"always":"avoid"};padding:6mm 8mm;box-sizing:border-box;">
        ${pair.map(t=>buildTeacherTableHTML(t,S,ay,sh)).join(`<div style="border-top:1px dashed #ccc;margin:8px 0;"></div>`)}
      </div>`).join("");
    const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/>
      <style>
        @page{size:A4 portrait;margin:0}
        body{font-family:'TH SarabunNew','Sarabun','Arial',sans-serif;margin:0;padding:0;}
        td,th{word-wrap:break-word;overflow:hidden;line-height:1.2;}
        @media print{body{margin:0;}}
      </style></head><body>${pagesHTML}</body></html>`;
    w.document.write(html);
    w.document.close();
    setTimeout(()=>w.print(),700);
  };

  // Export ตารางห้องเรียน ตาม format import_Schedule.xlsx
  const exportScheduleJSON=()=>{
    const data={version:1,exportedAt:new Date().toISOString(),schedule:S.schedule,locks:S.locks,assigns:S.assigns,teachers:S.teachers,subjects:S.subjects,rooms:S.rooms,levels:S.levels,plans:S.plans,depts:S.depts,meetings:S.meetings,specialRooms:S.specialRooms};
    const blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});
    const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`backup_timetable_${new Date().toISOString().slice(0,10)}.json`;a.click();
    st("Backup สำเร็จ");
  };

  const exportRoomScheduleXLSX=async(rooms)=>{
    const roomList=rooms||S.rooms;
    if(!roomList.length){st("ไม่มีห้องเรียน","error");return;}

    // โหลด SheetJS ถ้ายังไม่มี
    let XLib=window.XLSX;
    if(!XLib){
      await new Promise((res,rej)=>{
        const s=document.createElement("script");
        s.src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
        s.onload=res; s.onerror=rej;
        document.head.appendChild(s);
      });
      XLib=window.XLSX;
    }

    const DAYS_TH=["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
    const PERIOD_TIMES=[
      {id:1,start:"08.30",end:"09.20"},{id:2,start:"09.20",end:"10.10"},
      {id:3,start:"10.25",end:"11.15"},{id:4,start:"11.15",end:"12.05"},
      {id:5,start:"13.00",end:"13.50"},{id:6,start:"13.50",end:"14.40"},
      {id:7,start:"14.50",end:"15.40"},
    ];

    const wb=XLib.utils.book_new();
    roomList.forEach(room=>{
      const rows=[["วัน","รหัสวิชา","เริ่มเวลา","หมดเวลา","รหัสผู้ใช้งาน(ครูผู้สอน)"]];
      DAYS_TH.forEach(day=>{
        PERIOD_TIMES.forEach(p=>{
          const entries=S.schedule[room.id+"_"+day+"_"+p.id]||[];
          if(!entries.length){
            rows.push([day,"",p.start,p.end,""]);
          } else {
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
      XLib.utils.book_append_sheet(wb,ws,room.name.slice(0,31));
    });
    XLib.writeFile(wb,`ตารางสอนห้องเรียน_${ay?.year||"2568"}.xlsx`);
    st(`Export ${roomList.length} ห้อง สำเร็จ`);
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
      if(!window.confirm(`Restore ตารางสอน?\n\nไฟล์: ${f.name}\nบันทึกเมื่อ: ${data.exportedAt||"ไม่ทราบ"}\n\n⚠️ ข้อมูลตารางสอนปัจจุบันจะถูกทับ`))return;

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


  const exportRoomXL=(rm)=>{const h=["วัน",...PERIODS.map(p=>`คาบ${p.id}(${p.time})`)];const d=DAYS.map(day=>[day,...PERIODS.map(p=>{const en=S.schedule[`${rm.id}_${day}_${p.id}`]||[];return en.map(e=>{const sub=S.subjects.find(s=>s.id===e.subjectId);const t=S.teachers.find(x=>x.id===e.teacherId);return`${sub?.code||""} ${subDisplayName(sub)||""} (${t?.firstName||""})`}).join(" / ")})]);exportExcel(h,d,`ตารางเรียน_${rm.name}.xlsx`,rm.name);st(`Export ${rm.name}`)};

  const exportTeacherXL=(t)=>{const h=["วัน",...PERIODS.map(p=>`คาบ${p.id}(${p.time})`)];const d=DAYS.map(day=>[day,...PERIODS.map(p=>{let parts=[];Object.entries(S.schedule).forEach(([k,en])=>{if(!k.endsWith(`_${day}_${p.id}`))return;en?.forEach(e=>{const xCoIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);if(e.teacherId===t.id||xCoIds.includes(t.id)){const sub=S.subjects.find(s=>s.id===e.subjectId);const rid=k.split("_")[0];const rm=S.rooms.find(r=>r.id===rid);parts.push(`${sub?.code||""} ${sub?.name||""} (${rm?.name||""})`)}})});return parts.join(" / ")})]);exportExcel(h,d,`ตารางสอน_${t.prefix}${t.firstName}.xlsx`,"ตารางสอน");st(`Export ${t.firstName}`)};

  const exportAllRooms=()=>{exportExcelMulti(S.rooms.map(rm=>({name:rm.name,headers:["วัน",...PERIODS.map(p=>`คาบ${p.id}(${p.time})`)],rows:DAYS.map(day=>[day,...PERIODS.map(p=>{const en=S.schedule[`${rm.id}_${day}_${p.id}`]||[];return en.map(e=>{const sub=S.subjects.find(s=>s.id===e.subjectId);const t=S.teachers.find(x=>x.id===e.teacherId);return`${sub?.code||""} ${subDisplayName(sub)||""} (${t?.firstName||""})`}).join(" / ")})])})),"ตารางเรียนทุกห้อง.xlsx");st("Export ทุกห้อง")};

  const exportAllTeachers=()=>{exportExcelMulti(S.teachers.map(t=>({name:`${t.firstName} ${t.lastName}`,headers:["วัน",...PERIODS.map(p=>`คาบ${p.id}(${p.time})`)],rows:DAYS.map(day=>[day,...PERIODS.map(p=>{let parts=[];Object.entries(S.schedule).forEach(([k,en])=>{if(!k.endsWith(`_${day}_${p.id}`))return;en?.forEach(e=>{if(e.teacherId===t.id||e.coTeacherId===t.id){const sub=S.subjects.find(s=>s.id===e.subjectId);const rid=k.split("_")[0];const rm=S.rooms.find(r=>r.id===rid);parts.push(`${sub?.code||""} ${sub?.name||""} (${rm?.name||""})`)}})});return parts.join(" / ")})])})),"ตารางสอนทุกคน.xlsx");st("Export ทุกคน")};

  const exportStatus=()=>{
    const sheets=[{name:"ห้องเรียน",headers:["ห้อง","จัดแล้ว","ทั้งหมด","%"],rows:roomSt.map(r=>[r.room.name,r.filled,r.total,`${r.pct}%`])},{name:"ครู",headers:["ชื่อ","คาบได้รับ","จัดแล้ว","เหลือ","สถานะ"],rows:teacherSt.filter(t=>t.tot>0).map(t=>[`${t.teacher.prefix}${t.teacher.firstName} ${t.teacher.lastName}`,t.tot,t.used,t.rem,t.rem===0?"ครบ":"เหลือ "+t.rem])}];
    exportExcelMulti(sheets,"รายงานสถานะ.xlsx");st("Export สำเร็จ");
  };

  // PDF print for teacher
  // PDF: ตารางสอนครู — แสดง วิชา + ห้อง (ไม่มีครูร่วม)
  const printTeacherPDF=(t)=>{
    const w=window.open('','_blank');
    // helper เรียงห้องจากน้อยไปมาก
    const sortParts=(parts)=>parts.sort((a,b)=>{
      const numA=parseInt((a.room.match(/(\d+)$/)||[0,9999])[1]);
      const numB=parseInt((b.room.match(/(\d+)$/)||[0,9999])[1]);
      if(numA!==numB) return numA-numB;
      return a.room.localeCompare(b.room,"th");
    });
    let html=pdfPage(
      "ตารางสอน "+(t.prefix||"")+(t.firstName||"")+" "+(t.lastName||""),
      "ภาคเรียนที่ "+(ay?.semester||"1")+"/"+(ay?.year||"2568")+" "+(sh?.name||"โรงเรียนดาราวิทยาลัย"),
      DAYS.map(day=>({day,cells:PERIODS.map(p=>{
        let parts=[];
        Object.entries(S.schedule).forEach(([k,en])=>{
          const keySuffix="_"+day+"_"+p.id;
          if(!k.endsWith(keySuffix))return;
          en?.forEach(e=>{
            const pCoIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
            if(e.teacherId===t.id||pCoIds.includes(t.id)){
              const sub=S.subjects.find(s=>s.id===e.subjectId);
              const rid=k.split("_")[0];
              const rm=S.rooms.find(r=>r.id===rid);
              parts.push({sub:(sub?.shortName||sub?.name||""),room:rm?.name||"",room2:""});
            }
          });
        });
        return sortParts(parts);
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
        return[{sub:(sub?.shortName||sub?.name||""),room:(t2?.prefix||"")+(t2?.firstName||""),room2:room.name,double:isDouble}];
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


  // PDF แบบใหม่ — 2 ห้อง/หน้า A4 แนวตั้ง, 3 แถว/คาบ, คาบพักแนวตั้ง, auto homeroom

  // สร้าง HTML ตารางเรียนแบบเดียวกับ Excel ต้นแบบ
  const buildRoomTableHTML=(room,opts={})=>{
    const lvl=S.levels.find(l=>l.id===room.levelId);
    const asmDay=lvl?.assemblyDay||"";
    const h1=room.homeroom1||""; const h2=room.homeroom2||""; const hco=room.homeroomCo||"";
    const yr=ay?.year||"2568";
    const logoImg=sh?.logo?`<img src="${sh.logo}" style="height:40px;vertical-align:middle;margin-right:8px;"/>` :"";
    const title=opts.title||("ตารางเรียน "+room.name);

    const getCells=(day,pid)=>{
      const key=room.id+"_"+day+"_"+pid;
      return(S.schedule[key]||[]).map(e=>{
        const sub=S.subjects.find(s=>s.id===e.subjectId);
        const t=S.teachers.find(t=>t.id===e.teacherId);
        const cos=(e.coTeacherIds||[]).map(id=>S.teachers.find(x=>x.id===id)).filter(Boolean);
        return{th:sub?.name||sub?.code||"",en:sub?.shortName||"",tch:[t,...cos].filter(Boolean).map(x=>"อ."+x.firstName).join(", ")};
      });
    };

    // colgroup — % based
    const colgroup=`<colgroup>
      <col style="width:5%;"><col style="width:2.5%;">
      <col style="width:12%;"><col style="width:12%;">
      <col style="width:2.5%;">
      <col style="width:12%;"><col style="width:12%;">
      <col style="width:2.5%;">
      <col style="width:12%;"><col style="width:12%;">
      <col style="width:2.5%;">
      <col style="width:12%;">
    </colgroup>`;

    // vert cell: ข้อความแนวตั้ง ใช้กับทุก break column
    const vert=(txt,bg="#fffde7",fw="normal",fs="7pt")=>
      `<div style="writing-mode:vertical-rl;transform:rotate(180deg);white-space:nowrap;font-size:${fs};font-weight:${fw};padding:2px 0;text-align:center;">${txt}</div>`;

    const HDR=[
      {label:"คาบ 1",time:"08.30 - 09.20"},
      {label:"คาบ 2",time:"09.20 - 10.10"},
      {label:"คาบ 3",time:"10.25 - 11.15"},
      {label:"คาบ 4",time:"11.15 - 12.05"},
      {label:"คาบ 5",time:"13.00 - 13.50"},
      {label:"คาบ 6",time:"13.50 - 14.40"},
      {label:"คาบ 7",time:"14.50 - 15.40"},
    ];

    // Header — แถว 1: ชื่อคาบ, แถว 2: เวลา
    // col: วัน | 08.00-08.30 | คาบ1 | คาบ2 | 10.10-10.25 | คาบ3 | คาบ4 | 12.05-13.00 | คาบ5 | คาบ6 | 14.40-14.50 | คาบ7
    const BRK=["08.00-08.30","10.10-10.25","12.05-13.00","14.40-14.50"];
    const BRKV=["08.00-08.30","Morning Break","Lunch Time","Afternoon Break"];

    // Header row1 — ชื่อคาบ (break columns มี rowspan=2)
    const h1row=`<tr style="background:#f0f0f0;height:22px;max-height:22px;">
      <th rowspan="2" style="border:1px solid #666;padding:0;position:relative;vertical-align:middle;font-size:7pt;height:38px;max-height:38px;overflow:hidden;">
        <div style="position:absolute;top:0;left:0;width:100%;height:100%;">
          <svg style="position:absolute;top:0;left:0;width:100%;height:100%;" preserveAspectRatio="none">
            <line x1="0" y1="0" x2="100%" y2="100%" stroke="#888" stroke-width="0.8"/>
          </svg>
          <div style="position:absolute;top:2px;right:2px;font-size:5.5pt;color:#555;">เวลา</div>
          <div style="position:absolute;bottom:2px;left:2px;font-size:5.5pt;color:#555;">วัน</div>
        </div>
      </th>
      <th rowspan="2" style="border:1px solid #666;background:#fffde7;padding:0;height:38px;max-height:38px;overflow:hidden;">${vert(BRK[0],"#fffde7","normal","6.5pt")}</th>
      <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[0].label}</th>
      <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[1].label}</th>
      <th rowspan="2" style="border:1px solid #666;background:#fffde7;padding:0;height:38px;max-height:38px;overflow:hidden;">${vert(BRK[1],"#fffde7","normal","6.5pt")}</th>
      <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[2].label}</th>
      <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[3].label}</th>
      <th rowspan="2" style="border:1px solid #666;background:#fffde7;padding:0;height:38px;max-height:38px;overflow:hidden;">${vert(BRK[2],"#fffde7","normal","6.5pt")}</th>
      <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[4].label}</th>
      <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[5].label}</th>
      <th rowspan="2" style="border:1px solid #666;background:#fffde7;padding:0;height:38px;max-height:38px;overflow:hidden;">${vert(BRK[3],"#fffde7","normal","6.5pt")}</th>
      <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[6].label}</th>
    </tr>
    <tr style="background:#f0f0f0;height:16px;max-height:16px;">
      ${[0,1,2,3,4,5,6].map(i=>`<td style="border:1px solid #888;font-size:6.5pt;text-align:center;padding:1px;height:16px;">${HDR[i].time}</td>`).join("")}
    </tr>`;

    const DAYS_TH=["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
    let body="";

    DAYS_TH.forEach((day,di)=>{
      const isAsm=asmDay===day;
      // Homeroom: คำเดียว "Homeroom" หรือ 2 คำ "หอประชุม Assembly"
      const hmTxt=isAsm?"หอประชุม Assembly":"Homeroom";
      const hmBg=isAsm?"#e8f5e9":"#fafff7";
      const D=[1,2,3,4,5,6,7].map(pid=>getCells(day,pid));

      const cell=(arr,type)=>{
        const v=arr.map(c=>c[type]).filter(Boolean).join("<br>");
        const s=type==="th"
          ?"font-size:8.5pt;font-weight:bold;"
          :type==="en"
          ?"font-size:7.5pt;color:#444;"
          :"font-size:7.5pt;color:#1a237e;";
        return`<td style="border:1px solid #ddd;border-top:none;border-bottom:none;text-align:center;vertical-align:middle;padding:1px;${s}">${v}</td>`;
      };
      // แถวแรกของวัน — border-top ชัด
      const cellTop=(arr,type)=>cell(arr,type).replace("border-top:none;","border-top:1px solid #888;");
      // แถวสุดท้ายของวัน — border-bottom ชัด
      const cellBot=(arr,type)=>cell(arr,type).replace("border-bottom:none;","border-bottom:1px solid #888;");

      const BKcell=(rows,vtext,bg="#fffde7")=>
        `<td rowspan="${rows}" style="border:1px solid #888;background:${bg};padding:0;vertical-align:middle;">${vert(vtext,bg,"normal","6.5pt")}</td>`;

      body+=`
        <tr style="height:20px;">
          <td rowspan="3" style="border:1px solid #888;text-align:center;font-weight:bold;font-size:8.5pt;vertical-align:middle;background:#f5f5f5;">${day}</td>
          <td rowspan="3" style="border:1px solid #888;background:${hmBg};padding:0;vertical-align:middle;">${vert(hmTxt,hmBg,"bold","7pt")}</td>
          ${cellTop(D[0],"th")}${cellTop(D[1],"th")}
          ${BKcell(3,"พักน้อย  15  นาที")}
          ${cellTop(D[2],"th")}${cellTop(D[3],"th")}
          ${BKcell(3,"พักกลางวัน  55  นาที")}
          ${cellTop(D[4],"th")}${cellTop(D[5],"th")}
          ${BKcell(3,"พักน้อย  10  นาที")}
          ${cellTop(D[6],"th")}
        </tr>
        <tr style="height:17px;">
          ${cell(D[0],"en")}${cell(D[1],"en")}
          ${cell(D[2],"en")}${cell(D[3],"en")}
          ${cell(D[4],"en")}${cell(D[5],"en")}
          ${cell(D[6],"en")}
        </tr>
        <tr style="height:17px;">
          ${cellBot(D[0],"tch")}${cellBot(D[1],"tch")}
          ${cellBot(D[2],"tch")}${cellBot(D[3],"tch")}
          ${cellBot(D[4],"tch")}${cellBot(D[5],"tch")}
          ${cellBot(D[6],"tch")}
        </tr>`;
    });

    const footer=(h1||h2||hco)?`
      <div style="margin-top:6px;font-size:10pt;font-family:'TH SarabunNew','Sarabun',sans-serif;text-align:right;line-height:2;">
        ${h1||h2?`<div><b>ครูประจำชั้นหลัก</b>&emsp;&emsp;${h1}${h2?"&emsp;&emsp;&emsp;&emsp;&emsp;"+h2:""}</div>`:""}
        ${hco?`<div><b>ครูประจำชั้นร่วม</b>&emsp;&emsp;${hco}</div>`:""}
      </div>`:"";

    return`
      <div style="text-align:center;margin-bottom:5px;font-family:'TH SarabunNew','Sarabun',sans-serif;">
        ${logoImg}<b style="font-size:13pt;">${title}&emsp;&emsp;ปีการศึกษา ${yr}</b>
      </div>
      <table style="width:100%;border-collapse:collapse;table-layout:fixed;">
        ${colgroup}
        <thead>${h1row}</thead>
        <tbody>${body}</tbody>
      </table>
      ${footer}`;
  };

  const printRoomPDFNew=(rooms,opts={})=>{
    const roomList=Array.isArray(rooms)?rooms:[rooms];
    if(!roomList.length){st("ไม่มีห้องที่เลือก","error");return;}
    const w=window.open('','_blank');
    if(!w){st("Browser บล็อก popup","error");return;}

    const layout=opts.layout||"2portrait"; // "2portrait" | "1landscape"

    let pagesHTML="";
    if(layout==="1landscape"){
      // 1 ห้อง / 1 หน้า A4 แนวนอน
      pagesHTML=roomList.map((rm,pi)=>`
        <div style="page-break-after:${pi<roomList.length-1?"always":"avoid"};padding:8mm 10mm;box-sizing:border-box;">
          ${buildRoomTableHTML(rm,{})}
        </div>`).join("");
    } else {
      // 2 ห้อง / 1 หน้า A4 แนวตั้ง (default)
      const pages=[];
      for(let i=0;i<roomList.length;i+=2) pages.push(roomList.slice(i,i+2));
      pagesHTML=pages.map((pair,pi)=>`
        <div style="page-break-after:${pi<pages.length-1?"always":"avoid"};padding:6mm 8mm;box-sizing:border-box;">
          ${pair.map(rm=>buildRoomTableHTML(rm,{})).join(`<div style="border-top:1px dashed #ccc;margin:6px 0;"></div>`)}
        </div>`).join("");
    }

    const pageSize=layout==="1landscape"?"A4 landscape":"A4 portrait";
    const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/>
    <style>
      @page{size:${pageSize};margin:0}
      body{font-family:'TH SarabunNew','Sarabun','Arial',sans-serif;margin:0;padding:0;}
      td,th{word-wrap:break-word;overflow:hidden;line-height:1.15;}
      @media print{body{margin:0;}}
    </style></head><body>${pagesHTML}</body></html>`;

    w.document.write(html);
    w.document.close();
    setTimeout(()=>w.print(),700);
  };
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
            const pCoIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
            if(e.teacherId===t.id||pCoIds.includes(t.id)){
              const sub=S.subjects.find(s=>s.id===e.subjectId);
              const rid=k.split("_")[0];
              const rm=S.rooms.find(r=>r.id===rid);
              parts.push({sub:(sub?.shortName||sub?.name||""),room:rm?.name||"",room2:""});
            }
          });
        });
        // เรียงห้องจากน้อยไปมาก (เช่น ม.6/1, ม.6/5, ม.6/6)
        parts.sort((a,b)=>{
          const numA=parseInt((a.room.match(/(\d+)$/)||[0,9999])[1]);
          const numB=parseInt((b.room.match(/(\d+)$/)||[0,9999])[1]);
          if(numA!==numB) return numA-numB;
          return a.room.localeCompare(b.room,"th");
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

  return <div style={{animation:"fadeIn 0.3s",display:"flex",flexDirection:"column",gap:20}}>

    {/* ── PRINT CENTER ── */}
    <div style={{background:"#fff",borderRadius:16,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.07)"}}>
      <h3 style={{fontSize:17,fontWeight:800,marginBottom:20,display:"flex",alignItems:"center",gap:8}}>🖨️ Print / Export Center</h3>

      {/* ─ Section 1: PDF ─ */}
      <div style={{marginBottom:20}}>
        <div style={{fontSize:13,fontWeight:700,color:"#DC2626",marginBottom:12,borderBottom:"2px solid #FEE2E2",paddingBottom:6}}>📄 PDF — พิมพ์ตาราง</div>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>

          {/* ตารางสอนครู */}
          <div style={{background:"#FFF5F5",borderRadius:12,padding:"12px 16px",border:"1px solid #FECDD3"}}>
            <div style={{fontSize:13,fontWeight:700,color:"#991B1B",marginBottom:8}}>👨‍🏫 ตารางสอนครู</div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              <button onClick={printAllTeachersPDF} style={{...BS("#DC2626"),fontSize:12,padding:"7px 16px"}}>พิมพ์ทุกคน (แบบเดิม)</button>
              <button onClick={printMasterByDept} style={{...BS("#991B1B"),fontSize:12,padding:"7px 16px"}}>รวมกลุ่มสาระ</button>
              <button onClick={()=>{setSelectedTeachersPDF([]);setShowNewTeacherPDF(true);}} style={{...BS("#7C3AED"),fontSize:12,padding:"7px 16px"}}>🆕 พิมพ์แบบใหม่ (2คน/หน้า)</button>
            </div>
            <div style={{marginTop:8,display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:12,color:"#6B7280"}}>รายคน:</span>
              <div style={{flex:"1 1 200px",maxWidth:280}}>
                <SearchSelect value={selTeacherPDF} onChange={v=>setSelTeacherPDF(v)}
                  options={[{value:"",label:"-- เลือกครู --"},...S.teachers.map(t=>({value:t.id,label:`${t.prefix}${t.firstName} ${t.lastName}`}))]}
                  placeholder="-- เลือกครู --"/>
              </div>
              <button onClick={()=>{const t=S.teachers.find(x=>x.id===selTeacherPDF);if(t)printTeacherPDF(t);else st("เลือกครูก่อน","error");}}
                style={{...BS("#DC2626"),fontSize:12,padding:"7px 14px"}}>🖨️ พิมพ์</button>
            </div>
          </div>

          {/* ตารางเรียนห้อง */}
          <div style={{background:"#FFF5F5",borderRadius:12,padding:"12px 16px",border:"1px solid #FECDD3"}}>
            <div style={{fontSize:13,fontWeight:700,color:"#991B1B",marginBottom:8}}>🏫 ตารางเรียนห้อง</div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              <button onClick={printAllRoomsPDF} style={{...BS("#DB2777"),fontSize:12,padding:"7px 16px"}}>พิมพ์ทุกห้อง (แบบเดิม)</button>
              <button onClick={()=>{setNewRoomPDFOpts({selectedRooms:[],layout:"2portrait"});setShowNewRoomPDF(true);}} style={{...BS("#7C3AED"),fontSize:12,padding:"7px 16px"}}>🆕 PDF แบบใหม่</button>
              <button onClick={()=>exportRoomScheduleXLSX()} style={{...BS("#059669"),fontSize:12,padding:"7px 16px"}}><Icon name="download" size={13}/>📊 Excel ตารางห้อง (import format)</button>
            </div>
            <div style={{marginTop:8,display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:12,color:"#6B7280"}}>รายระดับ:</span>
              <div style={{flex:"0 1 160px"}}>
                <select style={{...IS,fontSize:12}} value={masterLevel} onChange={e=>setMasterLevel(e.target.value)}>
                  <option value="">-- ระดับชั้น --</option>
                  {S.levels.map(l=><option key={l.id} value={l.id}>{l.name}</option>)}
                </select>
              </div>
              <button onClick={printMasterByLevel} style={{...BS("#7C3AED"),fontSize:12,padding:"7px 14px"}}>🖨️ พิมพ์</button>
            </div>
            <div style={{marginTop:8,display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:12,color:"#6B7280"}}>รายห้อง (แบบเดิม):</span>
              <div style={{flex:"1 1 200px",maxWidth:280}}>
                <SearchSelect value={selRoomPDF} onChange={v=>setSelRoomPDF(v)}
                  options={[{value:"",label:"-- เลือกห้อง --"},...S.rooms.map(r=>({value:r.id,label:r.name}))]}
                  placeholder="-- เลือกห้อง --"/>
              </div>
              <button onClick={()=>{const r=S.rooms.find(x=>x.id===selRoomPDF);if(r)printRoomPDF(r);else st("เลือกห้องก่อน","error");}}
                style={{...BS("#DB2777"),fontSize:12,padding:"7px 14px"}}>🖨️ พิมพ์</button>
            </div>
          </div>

        </div>
      </div>

      {/* Modal: พิมพ์แบบใหม่ */}
      {showNewRoomPDF&&(
        <div style={{position:"fixed",inset:0,zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.5)"}}>
          <div style={{background:"#fff",borderRadius:16,boxShadow:"0 20px 60px rgba(0,0,0,0.3)",width:"min(560px,94%)",maxHeight:"90vh",overflowY:"auto",padding:24,fontFamily:"inherit"}}>
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
                    return<button key={opt.val} onClick={()=>setNewRoomPDFOpts(p=>({...p,layout:opt.val}))}
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
                    return<button key={r.id}
                      onClick={()=>setNewRoomPDFOpts(p=>({...p,selectedRooms:sel?p.selectedRooms.filter(id=>id!==r.id):[...p.selectedRooms,r.id]}))}
                      style={{padding:"4px 12px",borderRadius:20,border:`2px solid ${sel?"#7C3AED":"#E5E7EB"}`,background:sel?"#7C3AED":"#fff",color:sel?"#fff":"#374151",fontSize:12,fontWeight:sel?700:400,cursor:"pointer"}}>
                      {r.name}
                    </button>;
                  })}
                </div>
                <div style={{display:"flex",gap:6,marginTop:6}}>
                  <button onClick={()=>setNewRoomPDFOpts(p=>({...p,selectedRooms:S.rooms.map(r=>r.id)}))} style={{fontSize:11,color:"#7C3AED",background:"none",border:"1px solid #E5E7EB",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>เลือกทั้งหมด</button>
                  <button onClick={()=>setNewRoomPDFOpts(p=>({...p,selectedRooms:[]}))} style={{fontSize:11,color:"#6B7280",background:"none",border:"1px solid #E5E7EB",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>ล้าง</button>
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
              <button onClick={()=>setShowNewRoomPDF(false)} style={{...BO(),flex:1}}>ยกเลิก</button>
              <button
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
              <button
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
                  ?`${newRoomPDFOpts.selectedRooms?.length||0} หน้า`
                  :`${Math.ceil((newRoomPDFOpts.selectedRooms?.length||0)/2)} หน้า`})
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: พิมพ์ตารางสอนครูแบบใหม่ */}
      {showNewTeacherPDF&&(
        <div style={{position:"fixed",inset:0,zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.5)"}}>
          <div style={{background:"#fff",borderRadius:16,boxShadow:"0 20px 60px rgba(0,0,0,0.3)",width:"min(560px,94%)",maxHeight:"90vh",overflowY:"auto",padding:24,fontFamily:"inherit"}}>
            <div style={{fontSize:16,fontWeight:800,marginBottom:4}}>🆕 พิมพ์ตารางสอนครูแบบใหม่</div>
            <div style={{fontSize:11,color:"#6B7280",marginBottom:16}}>A4 แนวตั้ง — 2 คนต่อหน้า · แสดงวิชา+ห้อง+ชื่ออังกฤษ</div>
            <div>
              <label style={LS}>เลือกครู (กดหลายคนได้)</label>
              <div style={{display:"flex",gap:6,flexWrap:"wrap",maxHeight:220,overflowY:"auto",padding:6,border:"1px solid #E5E7EB",borderRadius:8}}>
                {[...S.teachers].sort((a,b)=>{
                  const da=S.depts.find(d=>d.id===a.departmentId)?.name||"";
                  const db=S.depts.find(d=>d.id===b.departmentId)?.name||"";
                  if(da!==db)return da.localeCompare(db,"th");
                  return a.firstName.localeCompare(b.firstName,"th");
                }).map(t=>{
                  const sel=selectedTeachersPDF.includes(t.id);
                  const dept=S.depts.find(d=>d.id===t.departmentId)?.name||"";
                  return<button key={t.id}
                    onClick={()=>setSelectedTeachersPDF(p=>sel?p.filter(id=>id!==t.id):[...p,t.id])}
                    style={{padding:"4px 12px",borderRadius:20,border:`2px solid ${sel?"#7C3AED":"#E5E7EB"}`,background:sel?"#7C3AED":"#fff",color:sel?"#fff":"#374151",fontSize:12,fontWeight:sel?700:400,cursor:"pointer"}}>
                    {t.prefix}{t.firstName} {t.lastName}
                    {dept&&<span style={{fontSize:10,opacity:0.7,marginLeft:4}}>[{dept}]</span>}
                  </button>;
                })}
              </div>
              <div style={{display:"flex",gap:6,marginTop:6}}>
                <button onClick={()=>setSelectedTeachersPDF(S.teachers.map(t=>t.id))} style={{fontSize:11,color:"#7C3AED",background:"none",border:"1px solid #E5E7EB",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>เลือกทั้งหมด</button>
                <button onClick={()=>setSelectedTeachersPDF([])} style={{fontSize:11,color:"#6B7280",background:"none",border:"1px solid #E5E7EB",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>ล้าง</button>
                <span style={{fontSize:11,color:"#6B7280",alignSelf:"center"}}>เลือก {selectedTeachersPDF.length} คน → {Math.ceil(selectedTeachersPDF.length/2)} หน้า</span>
              </div>
            </div>
            <div style={{display:"flex",gap:10,marginTop:20}}>
              <button onClick={()=>setShowNewTeacherPDF(false)} style={{...BO(),flex:1}}>ยกเลิก</button>
              <button disabled={!selectedTeachersPDF.length}
                onClick={()=>{
                  const teachers=S.teachers.filter(t=>selectedTeachersPDF.includes(t.id));
                  const w=window.open('','_blank');
                  if(!w){st("Browser บล็อก popup","error");return;}
                  const saved=w.setTimeout;w.setTimeout=()=>{};
                  printTeacherPDFNew(teachers);
                  setTimeout(()=>{if(w.setTimeout)w.setTimeout=saved;},100);
                }}
                style={{...BO("#7C3AED"),flex:1,opacity:selectedTeachersPDF.length?1:0.4,fontSize:12}}>
                👁️ ดูตัวอย่าง
              </button>
              <button disabled={!selectedTeachersPDF.length}
                onClick={()=>{
                  printTeacherPDFNew(S.teachers.filter(t=>selectedTeachersPDF.includes(t.id)));
                  setShowNewTeacherPDF(false);
                }}
                style={{...BS("#7C3AED"),flex:2,opacity:selectedTeachersPDF.length?1:0.4}}>
                🖨️ พิมพ์ ({Math.ceil(selectedTeachersPDF.length/2)} หน้า)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ─ Section 2: Excel ─ */}
      <div style={{marginBottom:20}}>
        <div style={{fontSize:13,fontWeight:700,color:"#2563EB",marginBottom:12,borderBottom:"2px solid #BFDBFE",paddingBottom:6}}>📊 Excel — ดาวน์โหลด</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <button onClick={exportAllRooms}    style={{...BS("#2563EB"),fontSize:12,padding:"7px 14px"}}><Icon name="download" size={13}/>ตารางทุกห้อง</button>
          <button onClick={exportAllTeachers} style={{...BS("#7C3AED"),fontSize:12,padding:"7px 14px"}}><Icon name="download" size={13}/>ตารางสอนทุกคน</button>
          <button onClick={exportStatus}      style={{...BS("#059669"),fontSize:12,padding:"7px 14px"}}><Icon name="download" size={13}/>รายงานสถานะ</button>
        </div>
        <div style={{marginTop:8,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <div style={{flex:"1 1 200px",maxWidth:260}}>
            <SearchSelect value={selTeacherXL} onChange={v=>setSelTeacherXL(v)}
              options={[{value:"",label:"-- ครูรายคน (Excel) --"},...S.teachers.map(t=>({value:t.id,label:`${t.prefix}${t.firstName} ${t.lastName}`}))]}
              placeholder="-- ครูรายคน (Excel) --"/>
          </div>
          <button onClick={()=>{const t=S.teachers.find(x=>x.id===selTeacherXL);if(t)exportTeacherXL(t);else st("เลือกครูก่อน","error");}}
            style={{...BS("#7C3AED"),fontSize:12,padding:"7px 14px"}}><Icon name="download" size={13}/>ดาวน์โหลด</button>
          <div style={{flex:"1 1 180px",maxWidth:220}}>
            <SearchSelect value={selRoomXL} onChange={v=>setSelRoomXL(v)}
              options={[{value:"",label:"-- ห้องรายคน (Excel) --"},...S.rooms.map(r=>({value:r.id,label:r.name}))]}
              placeholder="-- ห้องรายคน (Excel) --"/>
          </div>
          <button onClick={()=>{const r=S.rooms.find(x=>x.id===selRoomXL);if(r)exportRoomXL(r);else st("เลือกห้องก่อน","error");}}
            style={{...BS("#2563EB"),fontSize:12,padding:"7px 14px"}}><Icon name="download" size={13}/>ดาวน์โหลด</button>
        </div>
      </div>

      {/* ─ Section 3: Backup/Restore ─ */}
      <div>
        <div style={{fontSize:13,fontWeight:700,color:"#0891B2",marginBottom:12,borderBottom:"2px solid #BAE6FD",paddingBottom:6}}>💾 Backup / Restore</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <button onClick={exportScheduleJSON} style={{...BS("#0891B2"),fontSize:12,padding:"7px 14px"}}><Icon name="download" size={13}/>💾 Backup (.json)</button>
          <button onClick={()=>fileRefSched.current?.click()} style={{...BO("#0891B2"),fontSize:12,padding:"7px 14px",display:"flex",alignItems:"center",gap:6}}><Icon name="upload" size={13}/>📥 Restore (.json)</button>
          <input ref={fileRefSched} type="file" accept=".json" style={{display:"none"}} onChange={importScheduleJSON}/>
          <span style={{fontSize:11,color:"#9CA3AF"}}>— รองรับทั้ง backup บางส่วน (ตารางสอน) และ full backup (ทุกข้อมูล)</span>
        </div>
      </div>
    </div>

    {/* ── สถานะห้องเรียน ── */}
    <h3 style={{fontSize:18,fontWeight:700}}>สถานะห้องเรียน</h3>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:12}}>
      {roomSt.map(({room,filled,total,pct})=><div key={room.id} style={{padding:14,borderRadius:10,background:pct===100?"#F0FDF4":pct>0?"#FFFBEB":"#FEF2F2",border:`1px solid ${pct===100?"#BBF7D0":pct>0?"#FDE68A":"#FECACA"}`,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
        <div style={{display:"flex",justifyContent:"space-between"}}><span style={{fontWeight:700,fontSize:13}}>{room.name}</span><span style={{fontSize:12,fontWeight:700,color:pct===100?"#059669":pct>0?"#D97706":"#DC2626"}}>{pct}%</span></div>
        <div style={{height:5,background:"rgba(0,0,0,0.08)",borderRadius:3,marginTop:6,overflow:"hidden"}}><div style={{width:`${pct}%`,height:"100%",background:pct===100?"#059669":pct>0?"#D97706":"#DC2626",borderRadius:3}}/></div>
        <div style={{display:"flex",gap:5,marginTop:8}}>
          <button onClick={()=>exportRoomXL(room)} style={{background:"none",border:"1.5px solid #2563EB",borderRadius:6,padding:"2px 8px",color:"#2563EB",fontSize:10,fontWeight:600,cursor:"pointer"}}>Excel</button>
          <button onClick={()=>printRoomPDF(room)} style={{background:"none",border:"1.5px solid #DC2626",borderRadius:6,padding:"2px 8px",color:"#DC2626",fontSize:10,fontWeight:600,cursor:"pointer"}}>PDF</button>
        </div>
      </div>)}
      {!roomSt.length&&<div style={{padding:20,color:"#9CA3AF"}}>ยังไม่มีห้องเรียน</div>}
    </div>

    {/* ── สถานะครู ── */}
    <h3 style={{fontSize:18,fontWeight:700,marginTop:4}}>สถานะครู</h3>
    <div style={{background:"#fff",borderRadius:14,padding:20,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
        <thead><tr style={{background:"#F9FAFB"}}>{["ชื่อ","ได้รับ","จัดแล้ว","เหลือ","สถานะ","Export"].map(h=><th key={h} style={{padding:"8px 12px",textAlign:"left",fontWeight:600,color:"#6B7280",fontSize:12}}>{h}</th>)}</tr></thead>
        <tbody>
          {teacherSt.filter(t=>t.tot>0).map(ts=><tr key={ts.teacher.id} style={{borderTop:"1px solid #F3F4F6"}}>
            <td style={{padding:"8px 12px",fontWeight:600,fontSize:12}}>{ts.teacher.prefix}{ts.teacher.firstName} {ts.teacher.lastName}</td>
            <td style={{padding:"8px 12px",fontSize:12}}>{ts.tot}</td>
            <td style={{padding:"8px 12px",fontSize:12}}>{ts.used}</td>
            <td style={{padding:"8px 12px",fontWeight:700,color:ts.rem>0?"#D97706":"#059669",fontSize:12}}>{ts.rem}</td>
            <td style={{padding:"8px 12px"}}>{ts.rem===0?<span style={{background:"#D1FAE5",color:"#065F46",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:700}}>ครบ</span>:<span style={{background:"#FEF3C7",color:"#92400E",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:700}}>เหลือ {ts.rem}</span>}</td>
            <td style={{padding:"8px 12px"}}><div style={{display:"flex",gap:4}}>
              <button onClick={()=>exportTeacherXL(ts.teacher)} style={{background:"none",border:"1.5px solid #2563EB",borderRadius:5,padding:"2px 8px",color:"#2563EB",fontSize:10,fontWeight:600,cursor:"pointer"}}>Excel</button>
              <button onClick={()=>printTeacherPDF(ts.teacher)} style={{background:"none",border:"1.5px solid #DC2626",borderRadius:5,padding:"2px 8px",color:"#DC2626",fontSize:10,fontWeight:600,cursor:"pointer"}}>PDF</button>
            </div></td>
          </tr>)}
          {!teacherSt.filter(t=>t.tot>0).length&&<tr><td colSpan={6} style={{padding:24,textAlign:"center",color:"#9CA3AF"}}>ยังไม่มีครูที่กำหนดคาบ</td></tr>}
        </tbody>
      </table>
    </div>
  </div>;
}
/* ===== SETTINGS */
function Settings({S,U,st,ay,setAY,sh,setSH,div}){
  const logoRef=useRef(null);
  // helper: ล้าง localStorage dara_ keys และ force sync ไป GAS
  const clearLocalAndSync=async(newState)=>{
    // ล้าง localStorage ทุก key ของ division นี้
    Object.keys(localStorage)
      .filter(k=>k.startsWith("dara_"+div?.id)||k==="dara_division")
      .forEach(k=>localStorage.removeItem(k));
    // force sync ข้อมูลใหม่ไป GAS ทันที
    const {db}=getFB();
    if(db){ setSyncing(true); try{ await fsSaveTimetable(div?.id||"m2",newState); }catch(e){} setSyncing(false); }
  };

  const resetAll=async()=>{
    if(!confirm("⚠️ คุณแน่ใจหรือไม่ว่าต้องการลบข้อมูลทั้งหมด?\nข้อมูลที่จัดตารางไว้จะหายทั้งหมด!"))return;
    if(!confirm("ยืนยันอีกครั้ง — ลบข้อมูลทั้งหมดและเริ่มต้นใหม่?"))return;
    const newLevels=(div?.defaultLevels||["ระดับ 1","ระดับ 2","ระดับ 3"]).map(n=>({id:gid(),name:n}));
    const emptyState={levels:newLevels,plans:[],depts:[],teachers:[],subjects:[],rooms:[],specialRooms:[],assigns:[],meetings:[],schedule:{},locks:{}};
    U.setLevels(newLevels);
    U.setPlans([]);U.setDepts([]);U.setTeachers([]);U.setSubjects([]);
    U.setRooms([]);U.setSpecialRooms([]);U.setAssigns([]);U.setMeetings([]);U.setSchedule({});U.setLocks({});
    await clearLocalAndSync(emptyState);
    st("รีเซ็ทข้อมูลทั้งหมดแล้ว และ sync แล้ว","warning");
  };
  const resetScheduleOnly=async()=>{
    if(!confirm("ลบเฉพาะข้อมูลตารางสอน (ข้อมูลครู/วิชา/ห้องยังอยู่)?"))return;
    U.setSchedule({});U.setLocks({});
    // ล้าง schedule ใน localStorage และ sync
    ["schedule","locks"].forEach(k=>localStorage.removeItem("dara_"+div?.id+"_"+k));
    const {db:db2}=getFB();
    if(db2){ setSyncing(true); try{ await fsSaveTimetable(div?.id||"m2",{...stateRef.current,schedule:{},locks:{}}); }catch(e){} setSyncing(false); }
    st("ล้างตารางสอนแล้ว และ sync แล้ว","warning");
  };
  const handleLogo=(e)=>{const f=e.target.files?.[0];if(!f)return;const reader=new FileReader();reader.onload=ev=>{setSH(p=>({...p,logo:ev.target.result}));st("อัพโหลดโลโก้สำเร็จ")};reader.readAsDataURL(f);e.target.value=""};

  return <div style={{animation:"fadeIn 0.3s"}}>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(400px,1fr))",gap:24}}>
      {/* Academic Year */}
      <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
        <h3 style={{fontSize:16,fontWeight:700,marginBottom:20}}>ปีการศึกษา</h3>
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          <div><label style={LS}>ปีการศึกษา (พ.ศ.)</label><input style={IS} value={ay.year} onChange={e=>{
            setAY(p=>({...p,year:e.target.value}));
          }} placeholder="2568"/></div>
          <div><label style={LS}>ภาคเรียนที่</label><select style={IS} value={ay.semester} onChange={e=>setAY(p=>({...p,semester:e.target.value}))}><option value="1">1</option><option value="2">2</option></select></div>
          <button onClick={()=>{
            if(!window.confirm(`เปลี่ยนปีการศึกษา → รีเซ็ตครูประจำชั้นทุกห้องด้วยไหม?\n(กด OK = รีเซ็ต, Cancel = ไม่รีเซ็ต)`))return;
            U.setRooms(p=>p.map(r=>({...r,homeroom1:"",homeroom2:"",homeroomCo:""})));
            st("รีเซ็ตครูประจำชั้นทุกห้องแล้ว","warning");
          }} style={{...BO("#D97706"),fontSize:12}}>🔄 รีเซ็ตครูประจำชั้นทุกห้อง (เมื่อเปลี่ยนปี)</button>
        </div>
      </div>

      {/* School Header + Logo */}
      <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
        <h3 style={{fontSize:16,fontWeight:700,marginBottom:20}}>หัวเอกสาร (สำหรับ PDF)</h3>
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          <div><label style={LS}>ชื่อโรงเรียน</label><input style={IS} value={sh.name} onChange={e=>setSH(p=>({...p,name:e.target.value}))} placeholder="โรงเรียนดาราวิทยาลัย"/></div>
          <div>
            <label style={LS}>โลโก้โรงเรียน (จะแสดงในตาราง PDF)</label>
            <div style={{display:"flex",alignItems:"center",gap:14,marginTop:8}}>
              {sh.logo
                ?<img src={sh.logo} alt="logo" style={{width:56,height:56,borderRadius:"50%",objectFit:"cover",border:"2px solid #E5E7EB"}} onError={e=>{e.target.style.display='none'}}/>
                :<div style={{width:56,height:56,borderRadius:"50%",background:"#F3F4F6",border:"2px dashed #D1D5DB",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:"#9CA3AF"}}>LOGO</div>
              }
              <div style={{flex:1,display:"flex",flexDirection:"column",gap:8}}>
                <input
                  style={{...IS,fontSize:12}}
                  value={sh.logo||""}
                  onChange={e=>setSH(p=>({...p,logo:e.target.value}))}
                  placeholder="วาง URL รูปภาพ เช่น https://drive.google.com/uc?id=..."
                />
                <div style={{display:"flex",gap:6}}>
                  <button onClick={()=>logoRef.current?.click()} style={{...BO("#2563EB"),fontSize:12,padding:"6px 12px"}}><Icon name="upload" size={13}/>Upload ไฟล์ (เครื่องนี้เท่านั้น)</button>
                  {sh.logo&&<button onClick={()=>{setSH(p=>({...p,logo:""}));st("ลบโลโก้แล้ว","warning")}} style={{...BO("#DC2626"),fontSize:12,padding:"6px 12px"}}><Icon name="trash" size={13}/>ลบ</button>}
                </div>
              </div>
              <input ref={logoRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleLogo}/>
            </div>
            <div style={{padding:"8px 12px",background:"#EFF6FF",borderRadius:8,marginTop:8,fontSize:12,color:"#1E40AF"}}>
              💡 <strong>แนะนำ:</strong> อัพโลโก้ขึ้น Google Drive → คลิกขวา → "Get link" → เปลี่ยน <code>drive.google.com/file/d/ID/view</code> เป็น <code>drive.google.com/uc?id=ID</code> แล้ววาง URL ด้านบน — ทุกเครื่องจะเห็นโลโก้เดียวกัน
            </div>
          </div>
        </div>
      </div>

      {/* Reset */}
      <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
        <h3 style={{fontSize:16,fontWeight:700,marginBottom:20,color:"#DC2626"}}>รีเซ็ทข้อมูล</h3>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <button onClick={resetScheduleOnly} style={BO("#D97706")}><Icon name="trash" size={16}/>ล้างเฉพาะตารางสอน</button>
          <button onClick={()=>{
            // เว้น academicYear และ schoolHeader ไว้ ลบแค่ข้อมูลหลัก
            const keepKeys=["dara_academicYear","dara_schoolHeader","dara_division"];
            Object.keys(localStorage)
              .filter(k=>k.startsWith("dara_")&&!keepKeys.includes(k))
              .forEach(k=>localStorage.removeItem(k));
            st("ล้าง Cache แล้ว — กำลัง reload...","warning");
            setTimeout(()=>window.location.reload(),1000);
          }} style={BO("#6B7280")}><Icon name="x" size={16}/>ล้าง Cache (แก้ข้อมูลไม่ตรง)</button>
          <div style={{fontSize:12,color:"#6B7280"}}>ลบข้อมูลตารางสอนที่จัดไว้ แต่ข้อมูลครู วิชา ห้อง ยังอยู่</div>
          <div style={{borderTop:"1px solid #E5E7EB",paddingTop:12,marginTop:4}}/>
          <button onClick={resetAll} style={BS("#DC2626")}><Icon name="trash" size={16}/>รีเซ็ทข้อมูลทั้งหมด</button>
          <div style={{fontSize:12,color:"#DC2626"}}>⚠️ ลบข้อมูลทุกอย่าง — ไม่สามารถกู้คืนได้</div>
        </div>
      </div>

      {/* Summary */}
      <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.06)"}}>
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

// Merge entries ที่วิชาเดียวกันในคาบเดียว → แสดงชื่อวิชาแค่ครั้งเดียว เรียงห้องลงมา
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
    return { sub: g.sub, roomHtml: roomHtml, room2: '', double: g.double, roomCount: g.rooms.length };
  });
}

function buildTeacherTableHTML(teacher, S, ay, sh) {
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
    <col style="width:5%;"><col style="width:2.5%;">
    <col style="width:12%;"><col style="width:12%;">
    <col style="width:2.5%;">
    <col style="width:12%;"><col style="width:12%;">
    <col style="width:2.5%;">
    <col style="width:12%;"><col style="width:12%;">
    <col style="width:2.5%;">
    <col style="width:12%;">
  </colgroup>`;

  const vert=(txt,bg="#fffde7",fw="normal",fs="6.5pt")=>
    `<div style="writing-mode:vertical-rl;transform:rotate(180deg);white-space:nowrap;font-size:${fs};font-weight:${fw};padding:2px 0;text-align:center;">${txt}</div>`;

  const HDR=[
    {label:"คาบ 1",time:"08.30 - 09.20"},
    {label:"คาบ 2",time:"09.20 - 10.10"},
    {label:"คาบ 3",time:"10.25 - 11.15"},
    {label:"คาบ 4",time:"11.15 - 12.05"},
    {label:"คาบ 5",time:"13.00 - 13.50"},
    {label:"คาบ 6",time:"13.50 - 14.40"},
    {label:"คาบ 7",time:"14.50 - 15.40"},
  ];
  const BRK=["08.00-08.30","10.10-10.25","12.05-13.00","14.40-14.50"];

  const hdrRow=`<tr style="background:#f0f0f0;height:22px;max-height:22px;">
    <th rowspan="2" style="border:1px solid #666;padding:0;position:relative;vertical-align:middle;font-size:7pt;height:38px;max-height:38px;overflow:hidden;">
      <div style="position:absolute;top:0;left:0;width:100%;height:100%;">
        <svg style="position:absolute;top:0;left:0;width:100%;height:100%;" preserveAspectRatio="none">
          <line x1="0" y1="0" x2="100%" y2="100%" stroke="#888" stroke-width="0.8"/>
        </svg>
        <div style="position:absolute;top:2px;right:2px;font-size:5.5pt;color:#555;">เวลา</div>
        <div style="position:absolute;bottom:2px;left:2px;font-size:5.5pt;color:#555;">วัน</div>
      </div>
    </th>
    <th rowspan="2" style="border:1px solid #666;background:#fffde7;padding:0;height:38px;max-height:38px;overflow:hidden;">${vert(BRK[0])}</th>
    <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[0].label}</th>
    <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[1].label}</th>
    <th rowspan="2" style="border:1px solid #666;background:#fffde7;padding:0;height:38px;max-height:38px;overflow:hidden;">${vert(BRK[1])}</th>
    <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[2].label}</th>
    <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[3].label}</th>
    <th rowspan="2" style="border:1px solid #666;background:#fffde7;padding:0;height:38px;max-height:38px;overflow:hidden;">${vert(BRK[2])}</th>
    <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[4].label}</th>
    <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[5].label}</th>
    <th rowspan="2" style="border:1px solid #666;background:#fffde7;padding:0;height:38px;max-height:38px;overflow:hidden;">${vert(BRK[3])}</th>
    <th style="border:1px solid #666;font-size:8pt;font-weight:bold;text-align:center;padding:1px;height:22px;">${HDR[6].label}</th>
  </tr>
  <tr style="background:#f0f0f0;height:16px;max-height:16px;">
    ${HDR.map(h=>`<td style="border:1px solid #888;font-size:6.5pt;text-align:center;padding:1px;height:16px;">${h.time}</td>`).join("")}
  </tr>`;

  // หา assemblyDay จากห้องที่ครูสอน (ใช้ level ของห้องแรก)
  const teacherRooms=[...new Set(
    Object.keys(S.schedule).flatMap(k=>
      (S.schedule[k]||[]).filter(e=>e.teacherId===teacher.id||(e.coTeacherIds||[]).includes(teacher.id))
        .map(()=>k.split("_")[0])
    )
  )].map(id=>S.rooms.find(r=>r.id===id)).filter(Boolean);

  const DAYS_TH=["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
  let body="";

  DAYS_TH.forEach((day,di)=>{
    const D=[1,2,3,4,5,6,7].map(pid=>getCells(day,pid));
    const bgRow=di%2===0?"":"background:#fafafa;";

    // หา assemblyDay ของห้องที่ครูสอนในวันนี้ (ถ้ามี)
    const roomsThisDay=teacherRooms.filter(rm=>{
      const lvl=S.levels.find(l=>l.id===rm.levelId);
      return lvl?.assemblyDay===day;
    });
    const isAsm=roomsThisDay.length>0;
    const hmTxt=isAsm?"หอประชุม Assembly":"Homeroom";
    const hmBg=isAsm?"#e8f5e9":"#fafff7";

    // teacher PDF: แต่ละ cell มี 2 แถว — ชื่อวิชา / ห้อง+ชื่ออังกฤษ
    const cell1=(arr)=>{ // แถวบน: ชื่อวิชาไทย
      const v=arr.map(c=>c.th).filter(Boolean).join("<br>");
      return`<td style="border:1px solid #ddd;border-bottom:none;text-align:center;vertical-align:middle;padding:1px;font-size:8.5pt;font-weight:bold;">${v}</td>`;
    };
    const cell2=(arr)=>{ // แถวล่าง: ห้อง + ชื่ออังกฤษ
      const v=arr.map(c=>`${c.room}${c.en?" — "+c.en:""}`).filter(Boolean).join("<br>");
      return`<td style="border:1px solid #ddd;border-top:none;text-align:center;vertical-align:middle;padding:1px;font-size:7.5pt;color:#1a237e;">${v}</td>`;
    };
    const BKcell=(rows,txt)=>
      `<td rowspan="${rows}" style="border:1px solid #888;background:#fffde7;padding:0;vertical-align:middle;">${vert(txt)}</td>`;

    body+=`
      <tr style="height:22px;${bgRow}">
        <td rowspan="2" style="border:1px solid #888;text-align:center;font-weight:bold;font-size:8.5pt;vertical-align:middle;background:#f5f5f5;">${day}</td>
        <td rowspan="2" style="border:1px solid #888;background:${hmBg};padding:0;vertical-align:middle;">${vert(hmTxt,hmBg,"bold","7pt")}</td>
        ${cell1(D[0])}${cell1(D[1])}
        ${BKcell(2,"พักน้อย  15  นาที")}
        ${cell1(D[2])}${cell1(D[3])}
        ${BKcell(2,"พักกลางวัน  55  นาที")}
        ${cell1(D[4])}${cell1(D[5])}
        ${BKcell(2,"พักน้อย  10  นาที")}
        ${cell1(D[6])}
      </tr>
      <tr style="height:18px;${bgRow}">
        ${cell2(D[0])}${cell2(D[1])}
        ${cell2(D[2])}${cell2(D[3])}
        ${cell2(D[4])}${cell2(D[5])}
        ${cell2(D[6])}
      </tr>`;
  });

  return`
    <div style="text-align:center;margin-bottom:5px;font-family:'TH SarabunNew','Sarabun',sans-serif;">
      ${logoImg}<b style="font-size:13pt;">${title}&emsp;&emsp;ปีการศึกษา ${yr}</b>
      ${dept?`<div style="font-size:9pt;color:#555;">${dept}</div>`:""}
    </div>
    <table style="width:100%;border-collapse:collapse;table-layout:fixed;font-family:'TH SarabunNew','Sarabun',sans-serif;">
      ${colgroup}
      <thead>${hdrRow}</thead>
      <tbody>${body}</tbody>
    </table>`;
}

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
    const dayCells = r.cells.map(function(rawEntries) {
      if (!rawEntries || !rawEntries.length) return '<td class="slot"></td>';
      const entries = groupEntries(rawEntries);
      const isDouble = entries.some(function(e){ return e.double || e.roomCount > 1; });
      const inner = entries.map(function(e) {
        let h = '<div class="ent"><div class="ent-sub">' + e.sub + '</div>' + e.roomHtml + '</div>';
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
      const dayCells = r.cells.map(function(rawEntries) {
        if (!rawEntries || !rawEntries.length) return '<td class="slot"></td>';
        const entries = groupEntries(rawEntries);
        const isDouble = entries.some(function(e){ return e.double || e.roomCount > 1; });
        const inner = entries.map(function(e) {
          let h = '<div class="ent"><div class="ent-sub">' + e.sub + '</div>' + e.roomHtml + '</div>';
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
