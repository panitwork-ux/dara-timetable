import { useState, useCallback, useEffect, useRef, useMemo } from "react";
import * as XLSX from 'xlsx';
import { initializeApp } from "firebase/app";
import { getAuth, GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged } from "firebase/auth";
import { getFirestore, initializeFirestore, persistentLocalCache, persistentMultipleTabManager, doc, getDoc, setDoc, collection, getDocs, onSnapshot } from "firebase/firestore";

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
    // autoDetectLongPolling แก้ปัญหา WebChannel 400 error บน GitHub Pages
    _db=initializeFirestore(_fbApp,{
      experimentalAutoDetectLongPolling:true,
      localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }),
    });
    // persistentLocalCache ใน initializeFirestore แทน enableIndexedDbPersistence (deprecated)
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
  const [addPerms,setAddPerms]=useState({p1:false,p2:false,m1:false,m2:false,canEdit:false,isTeacher:false});
  const [addLoading,setAddLoading]=useState(false);
  // Bulk import
  const [bulkMode,setBulkMode]=useState(false);
  const [bulkText,setBulkText]=useState("");
  const [bulkPerms,setBulkPerms]=useState({p1:false,p2:false,m1:false,m2:false,canEdit:false,isTeacher:false});
  const [bulkLoading,setBulkLoading]=useState(false);
  const [bulkResult,setBulkResult]=useState(null);

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
    const emails=addEmail.split('\n').map(e=>e.trim().toLowerCase()).filter(e=>e&&e.includes('@'));
    if(!emails.length){showToast("กรุณากรอกอีเมลอย่างน้อย 1 รายการ","error");return;}
    setAddLoading(true);
    const {db}=getFB();
    if(!db){showToast("Firebase ไม่พร้อม","error");setAddLoading(false);return;}
    let ok=0;
    for(const email of emails){
      const existing=users.find(u=>u.email===email&&!u.preAdded);
      const uid=existing?existing.uid:makePreKey(email);
      await setDoc(doc(db,"permissions",uid),{
        email,
        displayName:existing?.displayName||"",
        divisions:addPerms,
        preAdded:!existing,
        merged:false,
      },{merge:true});
      ok++;
    }
    showToast(`บันทึกสำเร็จ ${ok} อีเมล`);
    setAddEmail("");
    setAddPerms({p1:false,p2:false,m1:false,m2:false,canEdit:false,isTeacher:false});
    setAddLoading(false);
    loadUsers();
  };

  // ── Bulk Import: วิเคราะห์ข้อความ CSV/TSV/บรรทัดธรรมดา แล้ว import ทีเดียว ──
  // รองรับรูปแบบ: email,p1,p2,m1,m2,canEdit,isTeacher (0/1/true/false/ใช่/ไม่ใช่)
  // ถ้าไม่มี column สิทธิ์ → ใช้ bulkPerms ที่เลือกไว้
  const parseBulkText=(raw)=>{
    const lines=raw.split('\n').map(l=>l.trim()).filter(l=>l&&!l.startsWith('#'));
    if(!lines.length) return [];
    // ตรวจว่าบรรทัดแรกเป็น header หรือเปล่า
    const hasHeader=lines[0].toLowerCase().includes('email')||lines[0].toLowerCase().includes('อีเมล');
    const dataLines=hasHeader?lines.slice(1):lines;
    const toBool=v=>{
      if(!v&&v!==0) return null;
      const s=String(v).trim().toLowerCase();
      return s==='1'||s==='true'||s==='yes'||s==='ใช่'||s==='✓'||s==='x';
    };
    return dataLines.map(line=>{
      // รองรับ tab, comma, semicolon, pipe
      const parts=line.split(/[\t,;|]/).map(p=>p.trim().replace(/^["']|["']$/g,''));
      const email=parts[0]?.toLowerCase();
      if(!email||!email.includes('@')) return null;
      const hasPerms=parts.length>1;
      return {
        email,
        divisions: hasPerms ? {
          p1:  toBool(parts[1])??bulkPerms.p1,
          p2:  toBool(parts[2])??bulkPerms.p2,
          m1:  toBool(parts[3])??bulkPerms.m1,
          m2:  toBool(parts[4])??bulkPerms.m2,
          canEdit:    toBool(parts[5])??bulkPerms.canEdit,
          isTeacher:  toBool(parts[6])??bulkPerms.isTeacher,
        } : {...bulkPerms},
        valid: true,
      };
    }).filter(Boolean);
  };

  const bulkPreview=useMemo(()=>parseBulkText(bulkText),[bulkText,bulkPerms]);

  const handleBulkImport=async()=>{
    if(!bulkPreview.length){showToast("ไม่พบอีเมลที่ถูกต้อง","error");return;}
    setBulkLoading(true);
    const {db}=getFB();
    if(!db){showToast("Firebase ไม่พร้อม","error");setBulkLoading(false);return;}
    let ok=0,skip=0;
    const results=[];
    for(const row of bulkPreview){
      try{
        const existing=users.find(u=>u.email===row.email&&!u.preAdded);
        const uid=existing?existing.uid:makePreKey(row.email);
        // ตรวจ canEdit กับ isTeacher ห้ามเป็น true พร้อมกัน
        const divs={...row.divisions};
        if(divs.canEdit) divs.isTeacher=false;
        if(divs.isTeacher) divs.canEdit=false;
        await setDoc(doc(db,"permissions",uid),{
          email:row.email,
          displayName:existing?.displayName||"",
          divisions:divs,
          preAdded:!existing,
          merged:false,
        },{merge:true});
        results.push({email:row.email,status:"ok"});
        ok++;
      }catch(e){
        results.push({email:row.email,status:"error",msg:e.message});
        skip++;
      }
    }
    setBulkResult({ok,skip,results});
    setBulkLoading(false);
    if(ok>0){
      showToast(`นำเข้าสำเร็จ ${ok} อีเมล${skip>0?" (ผิดพลาด "+skip+")":""}`);
      loadUsers();
    } else {
      showToast("นำเข้าไม่สำเร็จ","error");
    }
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
          <p style={{fontSize:12,color:"#6B7280",marginBottom:16}}>เพิ่มอีเมลพร้อมสิทธิ์ได้เลย — วางหลายอีเมลพร้อมกันได้ (แต่ละบรรทัด) เมื่อผู้ใช้ login ครั้งแรกระบบจะจำสิทธิ์ที่ตั้งไว้</p>
          <div style={{display:"flex",gap:14,flexWrap:"wrap"}}>
            <div style={{flex:"1 1 260px"}}>
              <label style={LS}>อีเมล <span style={{fontWeight:400,color:"#9CA3AF"}}>(1 บรรทัด = 1 อีเมล)</span></label>
              <textarea
                style={{...IS,height:88,resize:"vertical",fontFamily:"monospace",fontSize:12}}
                value={addEmail}
                onChange={e=>setAddEmail(e.target.value)}
                placeholder={"teacher1@web1.dara.ac.th\nteacher2@web1.dara.ac.th\nteacher3@web1.dara.ac.th"}
              />
              <div style={{fontSize:11,color:"#6B7280",marginTop:2}}>
                {addEmail.split('\n').filter(l=>l.trim()&&l.includes('@')).length} อีเมลที่ถูกต้อง
              </div>
            </div>
            <div style={{flex:"2 1 300px",display:"flex",flexDirection:"column",gap:10}}>
              <div>
                <label style={LS}>ระดับที่เข้าได้</label>
                <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                  {Object.entries(divNames).map(([k,name])=>(
                    <label key={k} style={{display:"flex",alignItems:"center",gap:5,cursor:"pointer",padding:"6px 10px",borderRadius:8,border:`2px solid ${addPerms[k]?"#DC2626":"#D1D5DB"}`,background:addPerms[k]?"#FEE2E2":"#F9FAFB",userSelect:"none"}}>
                      <input type="checkbox" checked={!!addPerms[k]} onChange={e=>setAddPerms(p=>({...p,[k]:e.target.checked}))} style={{width:14,height:14,accentColor:"#DC2626"}}/>
                      <span style={{fontSize:12,fontWeight:addPerms[k]?700:400,color:addPerms[k]?"#991B1B":"#374151"}}>{name}</span>
                    </label>
                  ))}
                </div>
              </div>
              <div>
                <label style={LS}>สิทธิ์การใช้งาน</label>
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                  <label style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer",padding:"7px 14px",borderRadius:8,border:`2px solid ${addPerms.canEdit?"#7C3AED":"#D1D5DB"}`,background:addPerms.canEdit?"#EDE9FE":"#F9FAFB",userSelect:"none"}}>
                    <input type="checkbox" checked={!!addPerms.canEdit} onChange={e=>setAddPerms(p=>({...p,canEdit:e.target.checked,isTeacher:e.target.checked?false:p.isTeacher}))} style={{width:14,height:14,accentColor:"#7C3AED"}}/>
                    <span style={{fontSize:13,fontWeight:addPerms.canEdit?700:400,color:addPerms.canEdit?"#5B21B6":"#374151"}}>✏️ แก้ไขตารางได้</span>
                  </label>
                  <label style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer",padding:"7px 14px",borderRadius:8,border:`2px solid ${addPerms.isTeacher?"#0891B2":"#D1D5DB"}`,background:addPerms.isTeacher?"#ECFEFF":"#F9FAFB",userSelect:"none"}}>
                    <input type="checkbox" checked={!!addPerms.isTeacher} onChange={e=>setAddPerms(p=>({...p,isTeacher:e.target.checked,canEdit:e.target.checked?false:p.canEdit}))} style={{width:14,height:14,accentColor:"#0891B2"}}/>
                    <span style={{fontSize:13,fontWeight:addPerms.isTeacher?700:400,color:addPerms.isTeacher?"#0E7490":"#374151"}}>🔄 แลกคาบอย่างเดียว</span>
                  </label>
                </div>
                <div style={{fontSize:11,color:"#6B7280",marginTop:4}}>
                  {addPerms.canEdit?"→ เข้าได้ทุกเมนู แก้ตารางได้":addPerms.isTeacher?"→ เข้าได้แค่เมนูแลกคาบ":"→ เข้าดูตารางได้ตามระดับที่เลือก"}
                </div>
              </div>
              <button
                onClick={handleAddEmail}
                disabled={addLoading}
                style={{...BS(),alignSelf:"flex-start",opacity:addLoading?0.6:1}}
              >
                {addLoading?"กำลังบันทึก...":"บันทึกสิทธิ์"}
              </button>
            </div>
          </div>
        </div>

        {/* ── Bulk Import ── */}
        <div style={{background:"#fff",borderRadius:14,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.06)",marginBottom:20}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:4}}>
            <h2 style={{fontSize:15,fontWeight:700}}>📥 Bulk Import จาก CSV / วางข้อความ</h2>
            <button
              onClick={()=>{setBulkMode(v=>!v);setBulkResult(null);}}
              style={{background:"none",border:"1px solid #D1D5DB",borderRadius:8,padding:"4px 14px",fontSize:12,cursor:"pointer",color:bulkMode?"#DC2626":"#374151",fontWeight:600}}
            >{bulkMode?"▲ ซ่อน":"▼ เปิด"}</button>
          </div>
          {!bulkMode&&<p style={{fontSize:12,color:"#6B7280",margin:0}}>นำเข้าหลายอีเมลพร้อมสิทธิ์จากไฟล์ CSV หรือวางข้อความ — รองรับรูปแบบ <code style={{background:"#F3F4F6",padding:"1px 5px",borderRadius:4}}>email,p1,p2,m1,m2,canEdit,isTeacher</code></p>}

          {bulkMode&&(
            <div style={{marginTop:12}}>
              <p style={{fontSize:12,color:"#6B7280",marginBottom:12}}>
                รองรับรูปแบบ: <code style={{background:"#F3F4F6",padding:"1px 5px",borderRadius:4}}>อีเมล,p1,p2,m1,m2,canEdit,isTeacher</code> (0/1) คั่นด้วย comma, tab, semicolon, หรือ pipe<br/>
                ถ้าไม่ระบุ column สิทธิ์ → ใช้ค่าจาก "สิทธิ์ default" ด้านล่าง
              </p>

              {/* Default permissions สำหรับ row ที่ไม่มี column สิทธิ์ */}
              <div style={{background:"#F9FAFB",borderRadius:10,padding:"12px 16px",marginBottom:12,border:"1px solid #E5E7EB"}}>
                <div style={{fontSize:12,fontWeight:600,color:"#374151",marginBottom:8}}>สิทธิ์ default (ใช้เมื่อ CSV ไม่มี column สิทธิ์)</div>
                <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                  {Object.entries(divNames).map(([k,name])=>(
                    <label key={k} style={{display:"flex",alignItems:"center",gap:5,cursor:"pointer",padding:"5px 10px",borderRadius:8,border:`2px solid ${bulkPerms[k]?"#DC2626":"#D1D5DB"}`,background:bulkPerms[k]?"#FEE2E2":"#fff",userSelect:"none",fontSize:12}}>
                      <input type="checkbox" checked={!!bulkPerms[k]} onChange={e=>setBulkPerms(p=>({...p,[k]:e.target.checked}))} style={{width:13,height:13,accentColor:"#DC2626"}}/>
                      <span style={{fontWeight:bulkPerms[k]?700:400,color:bulkPerms[k]?"#991B1B":"#374151"}}>{name}</span>
                    </label>
                  ))}
                  <label style={{display:"flex",alignItems:"center",gap:5,cursor:"pointer",padding:"5px 10px",borderRadius:8,border:`2px solid ${bulkPerms.canEdit?"#7C3AED":"#D1D5DB"}`,background:bulkPerms.canEdit?"#EDE9FE":"#fff",userSelect:"none",fontSize:12}}>
                    <input type="checkbox" checked={!!bulkPerms.canEdit} onChange={e=>setBulkPerms(p=>({...p,canEdit:e.target.checked,isTeacher:e.target.checked?false:p.isTeacher}))} style={{width:13,height:13,accentColor:"#7C3AED"}}/>
                    <span style={{fontWeight:bulkPerms.canEdit?700:400,color:bulkPerms.canEdit?"#5B21B6":"#374151"}}>✏️ แก้ตาราง</span>
                  </label>
                  <label style={{display:"flex",alignItems:"center",gap:5,cursor:"pointer",padding:"5px 10px",borderRadius:8,border:`2px solid ${bulkPerms.isTeacher?"#0891B2":"#D1D5DB"}`,background:bulkPerms.isTeacher?"#ECFEFF":"#fff",userSelect:"none",fontSize:12}}>
                    <input type="checkbox" checked={!!bulkPerms.isTeacher} onChange={e=>setBulkPerms(p=>({...p,isTeacher:e.target.checked,canEdit:e.target.checked?false:p.canEdit}))} style={{width:13,height:13,accentColor:"#0891B2"}}/>
                    <span style={{fontWeight:bulkPerms.isTeacher?700:400,color:bulkPerms.isTeacher?"#0E7490":"#374151"}}>🔄 แลกคาบ</span>
                  </label>
                </div>
              </div>

              {/* Text area */}
              <label style={LS}>วางข้อความ CSV / อีเมลรายบรรทัด</label>
              <textarea
                style={{...IS,height:140,resize:"vertical",fontFamily:"monospace",fontSize:12,marginBottom:6}}
                value={bulkText}
                onChange={e=>{setBulkText(e.target.value);setBulkResult(null);}}
                placeholder={"# ตัวอย่าง (มี header หรือไม่มีก็ได้)\nemail,p1,p2,m1,m2,canEdit,isTeacher\nteacher1@web1.dara.ac.th,0,0,1,1,1,0\nteacher2@web1.dara.ac.th,1,1,0,0,0,0\n\n# หรือวางอีเมลอย่างเดียว (ใช้สิทธิ์ default)\nteacher3@web1.dara.ac.th\nteacher4@web1.dara.ac.th"}
              />

              {/* Preview */}
              {bulkText.trim()&&(
                <div style={{marginBottom:12}}>
                  <div style={{fontSize:12,fontWeight:600,color:"#374151",marginBottom:6}}>
                    ตัวอย่างก่อน import — พบ <span style={{color:"#059669",fontWeight:700}}>{bulkPreview.length}</span> อีเมลที่ถูกต้อง
                  </div>
                  <div style={{maxHeight:160,overflowY:"auto",border:"1px solid #E5E7EB",borderRadius:8,fontSize:11}}>
                    <table style={{width:"100%",borderCollapse:"collapse"}}>
                      <thead>
                        <tr style={{background:"#F9FAFB",position:"sticky",top:0}}>
                          {["อีเมล","ประถมต้น","ประถมปลาย","มัธยมต้น","มัธยมปลาย","แก้ตาราง","แลกคาบ"].map(h=>(
                            <th key={h} style={{padding:"5px 8px",textAlign:"left",fontWeight:600,color:"#6B7280",borderBottom:"1px solid #E5E7EB",whiteSpace:"nowrap"}}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {bulkPreview.slice(0,50).map((row,i)=>(
                          <tr key={i} style={{borderBottom:"1px solid #F3F4F6",background:i%2===0?"#fff":"#FAFAFA"}}>
                            <td style={{padding:"4px 8px",fontFamily:"monospace",color:"#111",maxWidth:220,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{row.email}</td>
                            {["p1","p2","m1","m2","canEdit","isTeacher"].map(k=>(
                              <td key={k} style={{padding:"4px 8px",textAlign:"center",color:row.divisions[k]?"#059669":"#D1D5DB",fontSize:14}}>
                                {row.divisions[k]?"✓":"–"}
                              </td>
                            ))}
                          </tr>
                        ))}
                        {bulkPreview.length>50&&(
                          <tr><td colSpan={7} style={{padding:"6px 8px",color:"#9CA3AF",fontSize:11,textAlign:"center"}}>...และอีก {bulkPreview.length-50} รายการ</td></tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* Result */}
              {bulkResult&&(
                <div style={{marginBottom:12,padding:"10px 14px",borderRadius:8,background:bulkResult.skip>0&&bulkResult.ok===0?"#FEE2E2":bulkResult.skip>0?"#FEF3C7":"#D1FAE5",border:`1px solid ${bulkResult.skip>0&&bulkResult.ok===0?"#FECACA":bulkResult.skip>0?"#FDE68A":"#A7F3D0"}`}}>
                  <div style={{fontWeight:700,fontSize:13,marginBottom:4}}>
                    {bulkResult.ok>0&&<span style={{color:"#065F46"}}>✅ สำเร็จ {bulkResult.ok} รายการ </span>}
                    {bulkResult.skip>0&&<span style={{color:"#92400E"}}>⚠️ ผิดพลาด {bulkResult.skip} รายการ</span>}
                  </div>
                  {bulkResult.results.filter(r=>r.status==="error").map((r,i)=>(
                    <div key={i} style={{fontSize:11,color:"#DC2626",fontFamily:"monospace"}}>✕ {r.email}: {r.msg}</div>
                  ))}
                </div>
              )}

              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <button
                  onClick={handleBulkImport}
                  disabled={bulkLoading||!bulkPreview.length}
                  style={{...BS("#059669"),opacity:(bulkLoading||!bulkPreview.length)?0.5:1}}
                >
                  {bulkLoading?"กำลัง import...":"📥 Import "+bulkPreview.length+" อีเมล"}
                </button>
                {bulkText&&<button
                  onClick={()=>{setBulkText("");setBulkResult(null);}}
                  style={{background:"none",border:"1px solid #D1D5DB",borderRadius:8,padding:"8px 14px",fontSize:13,cursor:"pointer",color:"#6B7280"}}
                >ล้าง</button>}
              </div>
            </div>
          )}
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
                      {Object.entries(u.divisions||{}).filter(([k,v])=>v&&divNames[k]).map(([k])=>(
                        <span key={k} style={{background:"#FEE2E2",color:"#991B1B",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>{divNames[k]}</span>
                      ))}
                      {u.divisions?.canEdit&&<span style={{background:"#EDE9FE",color:"#5B21B6",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>✏️ แก้ตาราง</span>}
                      {u.divisions?.isTeacher&&<span style={{background:"#ECFEFF",color:"#0E7490",padding:"2px 10px",borderRadius:20,fontSize:11,fontWeight:600}}>🔄 แลกคาบ</span>}
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
                <div style={{borderTop:"1px solid #E5E7EB",paddingTop:8,marginTop:2}}>
                  <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>สิทธิ์พิเศษ</label>
                  <label style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer",padding:"10px 14px",borderRadius:10,background:editPerms.canEdit?"#EDE9FE":"#F9FAFB",border:`1.5px solid ${editPerms.canEdit?"#7C3AED":"#E5E7EB"}`,marginBottom:6}}>
                    <input type="checkbox" checked={!!editPerms.canEdit} onChange={e=>setEditPerms(p=>({...p,canEdit:e.target.checked,isTeacher:e.target.checked?false:p.isTeacher}))} style={{width:16,height:16,accentColor:"#7C3AED"}}/>
                    <div><div style={{fontWeight:700,color:editPerms.canEdit?"#5B21B6":"#374151",fontSize:14}}>✏️ แก้ไขตารางได้</div><div style={{fontSize:11,color:"#6B7280"}}>เข้าได้ทุกเมนู สามารถแก้ตารางสอนได้</div></div>
                  </label>
                  <label style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer",padding:"10px 14px",borderRadius:10,background:editPerms.isTeacher?"#ECFEFF":"#F9FAFB",border:`1.5px solid ${editPerms.isTeacher?"#0891B2":"#E5E7EB"}`}}>
                    <input type="checkbox" checked={!!editPerms.isTeacher} onChange={e=>setEditPerms(p=>({...p,isTeacher:e.target.checked,canEdit:e.target.checked?false:p.canEdit}))} style={{width:16,height:16,accentColor:"#0891B2"}}/>
                    <div><div style={{fontWeight:700,color:editPerms.isTeacher?"#0E7490":"#374151",fontSize:14}}>🔄 แลกคาบอย่างเดียว</div><div style={{fontSize:11,color:"#6B7280"}}>เข้าได้เฉพาะเมนูแลกคาบ / สอนแทน</div></div>
                  </label>
                </div>
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
        <button onClick={handlePrint} style={{background:"#B91C1C",color:"#fff",border:"none",borderRadius:8,padding:"8px 20px",fontSize:14,fontWeight:700,cursor:"pointer"}}>🖨️ พิมพ์</button>
        <button onClick={onClose} style={{background:"#374151",color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",fontSize:14,cursor:"pointer"}}>✕ ปิด</button>
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
  const [userPerms,setUserPerms]=useState(()=>{
    // โหลดจาก localStorage cache ก่อน เผื่อ Firestore offline
    try{const c=localStorage.getItem("dara_perms_cache");return c?JSON.parse(c):null;}catch{return null;}
  });
  const [showAdmin,setShowAdmin]=useState(false);

  // helper: fsGetPermissions พร้อม timeout 8 วินาที
  const fsGetPermsWithTimeout=async(uid)=>{
    return Promise.race([
      fsGetPermissions(uid),
      new Promise((_,rej)=>setTimeout(()=>rej(new Error("timeout")),8000))
    ]).catch(()=>null);
  };

  // refresh permissions จาก Firestore (เรียกได้ทุกเวลา)
  const refreshPerms=async(u)=>{
    const user=u||authUser;
    if(!user)return;
    const perms=await fsGetPermsWithTimeout(user.uid);
    if(perms){
      setUserPerms(perms);
      try{localStorage.setItem("dara_perms_cache",JSON.stringify(perms));}catch{}
    }
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

        // โหลด permissions ครั้งแรก (พร้อม timeout)
        let perms=await fsGetPermsWithTimeout(u.uid);

        // ตรวจ pre-key เสมอ ถ้า perms ยังไม่มีสิทธิ์ใดเลย
        const hasAnyAccess=perms&&Object.values(perms.divisions||{}).some(v=>Boolean(v)&&v!=='false');
        if((!perms||!hasAnyAccess)&&db){
          try{
            const emailKey=makePreKey(u.email);
            const preSnap=await Promise.race([getDoc(doc(db,"permissions",emailKey)),new Promise((_,r)=>setTimeout(()=>r(new Error("t")),5000))]).catch(()=>null);
            if(preSnap?.exists()&&!preSnap.data().merged){
              const preData=preSnap.data();
              const divs=preData.divisions||{p1:false,p2:false,m1:false,m2:false};
              await Promise.race([setDoc(doc(db,"permissions",u.uid),{displayName:u.displayName||"",email:u.email,divisions:divs,preAdded:false},{merge:true}),new Promise((_,r)=>setTimeout(()=>r(),5000))]).catch(()=>{});
              await Promise.race([setDoc(doc(db,"permissions",emailKey),{merged:true},{merge:true}),new Promise((_,r)=>setTimeout(()=>r(),5000))]).catch(()=>{});
              perms={...perms,divisions:divs};
            }
          }catch{}
        }

        if(!perms){
          // Firestore offline — ใช้ cache จาก localStorage
          const cached=localStorage.getItem("dara_perms_cache");
          if(cached){
            try{
              const cachedPerms=JSON.parse(cached);
              setUserPerms(cachedPerms);
            }catch{setUserPerms({divisions:{p1:false,p2:false,m1:false,m2:false}});}
          } else {
            // สร้าง empty perms (จะ update เมื่อ online)
            const emptyDivs={p1:false,p2:false,m1:false,m2:false};
            try{await Promise.race([fsSetPermissions(u.uid,{displayName:u.displayName||"",email:u.email,divisions:emptyDivs}),new Promise((_,r)=>setTimeout(()=>r(),5000))]).catch(()=>{});}catch{}
            setUserPerms({divisions:emptyDivs});
          }
        } else {
          // merge displayName/email ไม่ทับ divisions
          try{await Promise.race([setDoc(doc(db,"permissions",u.uid),{displayName:u.displayName||"",email:u.email},{merge:true}),new Promise((_,r)=>setTimeout(()=>r(),5000))]).catch(()=>{});}catch{}
          // re-fetch เพื่อให้ได้ค่าล่าสุด
          const freshPerms=await fsGetPermsWithTimeout(u.uid);
          const finalPerms=freshPerms||perms;
          setUserPerms(finalPerms);
          try{localStorage.setItem("dara_perms_cache",JSON.stringify(finalPerms));}catch{}
        }

        // Real-time listener — permissions อัปเดตทันทีเมื่อ admin แก้ไข
        if(db){
          unsubPerms=onSnapshot(doc(db,"permissions",u.uid),(snap)=>{
            if(snap.exists()){
              const data=snap.data();
              setUserPerms(data);
              try{localStorage.setItem("dara_perms_cache",JSON.stringify(data));}catch{}
            }
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
  // eslint-disable-next-line react-hooks/exhaustive-deps
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
      setTimeout(()=>{isSavingRef.current=false;},700); // รอ onSnapshot ผ่านไปก่อน (700ms กัน network ช้า)
    },delay);
  },[divId]);

  // Subscribe realtime onSnapshot เมื่อ login และเมื่อ switch division
  useEffect(()=>{
    const {db}=getFB(); if(!db||!authUser) return;
    fsReadyRef.current=false;
    setSyncing(true);
    const unsub=fsSubscribeTimetable(divId,(d)=>{
      if(!fsReadyRef.current){
        // ตรวจว่า Firestore ส่งข้อมูลจริงมาหรือเปล่า
        // ถ้าว่างทุก field → อาจเป็น offline หรือ document ว่างจริง
        const hasRealData=(d.teachers?.length||0)+(d.subjects?.length||0)+(d.rooms?.length||0)+(d.levels?.length||0)>0;
        const hasLocalData=Object.keys(localStorage).some(k=>k.startsWith("dara_"+divId+"_teachers")||k.startsWith("dara_"+divId+"_rooms"));

        if(!hasRealData&&hasLocalData){
          // Firestore ส่งว่างมา แต่ localStorage มีข้อมูล → ใช้ localStorage แทน (offline guard)
          fsReadyRef.current=true;
          setSyncing(false);
          setGasReady(true);
          return;
        }

        if(hasRealData){
          // Firestore มีข้อมูลจริง → ล้าง localStorage cache เก่าแล้วใช้ Firestore
          const keepKeys=["dara_academicYear","dara_schoolHeader","dara_division"];
          Object.keys(localStorage)
            .filter(k=>k.startsWith("dara_"+divId)&&!keepKeys.includes(k))
            .forEach(k=>localStorage.removeItem(k));
        }
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
    {id:"swap",icon:"layers",label:"แลกคาบ / สอนแทน"},
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

  // Filter division selector ตาม permissions
  const availDivs=firebaseConfigured
    ?DIVISIONS.filter(d=>Boolean(userPerms?.divisions?.[d.id]))
    :DIVISIONS;

  const divHasAccess=!firebaseConfigured||Boolean(userPerms?.divisions?.[divId]);

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
        {nav.map(n=>{
          // ตรวจสิทธิ์ครู (isTeacher) — ให้เข้าได้แค่หน้า swap
          const isTeacherOnly = firebaseConfigured && userPerms?.divisions?.isTeacher === true && !userPerms?.divisions?.canEdit;
          const isLocked = isTeacherOnly && n.id !== "swap";
          return (
            <div
              key={n.id}
              className={`ni ${page===n.id?"a":""}`}
              onClick={()=>{ if(!isLocked) setPage(n.id); }}
              title={isLocked?"คุณมีสิทธิ์เฉพาะหน้าแลกคาบ / สอนแทน":""}
              style={{
                display:"flex",alignItems:"center",gap:10,padding:"9px 12px",borderRadius:10,
                cursor:isLocked?"not-allowed":"pointer",
                color: isLocked ? "rgba(255,255,255,0.25)" : (page===n.id?"#fff":"rgba(255,255,255,0.7)"),
                fontSize:13,fontWeight:page===n.id?700:400,marginBottom:2,transition:"all 0.15s",
                background:page===n.id?"rgba(255,255,255,0.15)":"transparent",
                opacity: isLocked ? 0.4 : 1,
                userSelect:"none",
              }}
            >
              <Icon name={n.icon} size={16}/>
              {n.label}
              {isLocked && <span style={{marginLeft:"auto",fontSize:10,opacity:0.6}}>🔒</span>}
            </div>
          );
        })}
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
        {/* ── Guard 0: กำลังโหลดสิทธิ์จาก Firebase ── */}
        {firebaseConfigured&&authUser&&userPerms===null
          ?<div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:"100%",gap:12}}>
              <div style={{width:36,height:36,border:"3px solid #E5E7EB",borderTopColor:"#B91C1C",borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/>
              <p style={{color:"#9CA3AF",fontSize:13}}>กำลังโหลดสิทธิ์การเข้าใช้งาน...</p>
              <button onClick={()=>window.location.reload()} style={{marginTop:8,padding:"8px 20px",background:"#B91C1C",color:"#fff",border:"none",borderRadius:8,fontSize:13,cursor:"pointer"}}>🔄 Reload</button>
            </div>
        /* ── Guard 1: ไม่มีสิทธิ์ระดับชั้นนี้ ── */
        :firebaseConfigured&&!divHasAccess
          ?<div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:"100%",gap:16}}>
              <div style={{fontSize:48}}>🔒</div>
              <h2 style={{fontSize:20,fontWeight:700,color:"#374151"}}>ไม่มีสิทธิ์เข้าระดับนี้</h2>
              <p style={{color:"#6B7280",fontSize:14}}>กรุณาติดต่อผู้ดูแลระบบเพื่อขอสิทธิ์ {div.name}</p>
              <button onClick={()=>{try{localStorage.removeItem("dara_perms_cache");}catch{}refreshPerms&&refreshPerms();window.location.reload();}} style={{marginTop:8,padding:"10px 24px",background:"#B91C1C",color:"#fff",border:"none",borderRadius:10,fontSize:14,fontWeight:700,cursor:"pointer"}}>🔄 รีเฟรชสิทธิ์</button>
              <p style={{color:"#9CA3AF",fontSize:12}}>ถ้าได้รับสิทธิ์แล้วแต่ยังเข้าไม่ได้ ให้กดปุ่มนี้</p>
            </div>
          // ── Guard 2: ครู (isTeacher only) ห้ามเข้าหน้าอื่นนอกจาก swap ──
          :firebaseConfigured&&userPerms?.divisions?.isTeacher===true&&!(userPerms?.divisions?.canEdit)&&page!=="swap"
          ?<div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:"100%",gap:16}}>
              <div style={{fontSize:48}}>🔒</div>
              <h2 style={{fontSize:20,fontWeight:700,color:"#374151"}}>ไม่มีสิทธิ์เข้าหน้านี้</h2>
              <p style={{color:"#6B7280",fontSize:14}}>คุณมีสิทธิ์เฉพาะหน้า <strong>แลกคาบ / สอนแทน</strong> เท่านั้น</p>
              <button onClick={()=>setPage("swap")} style={{padding:"10px 24px",background:"#DC2626",color:"#fff",border:"none",borderRadius:10,fontSize:14,fontWeight:700,cursor:"pointer"}}>→ ไปหน้าแลกคาบ</button>
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
            {page==="swap"&&<SwapPage S={S} st={st} ay={academicYear} sh={schoolHeader}/>}
            {page==="reports"&&<Reports S={S} U={U} st={st} gc={gc} ay={academicYear} sh={schoolHeader}/>}
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
  const addLv=()=>{
    const n=prompt("ชื่อระดับชั้น:");
    if(n){U.setLevels(p=>[...p,{id:gid(),name:n,divisionId:guessDivision(n)}]);st("เพิ่มสำเร็จ")}
  };
  const editLv=(lv)=>{
    const n=prompt("แก้ไขชื่อระดับชั้น:",lv.name);
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
          {/* ระดับการศึกษา (division) */}
          <div style={{marginBottom:10,padding:"8px 12px",background:"#FFF7ED",borderRadius:8,border:"1px solid #FED7AA"}}>
            <div style={{fontSize:11,fontWeight:700,color:"#92400E",marginBottom:6}}>🏫 ระดับการศึกษา (กำหนดเวลาคาบ 6-7)</div>
            <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
              {DIVISIONS.map(div=>{
                const cur=lv.divisionId||guessDivision(lv.name);
                const active=cur===div.id;
                return(
                  <button key={div.id}
                    onClick={()=>U.setLevels(p=>p.map(l=>l.id===lv.id?{...l,divisionId:div.id}:l))}
                    style={{padding:"3px 10px",borderRadius:20,border:`1.5px solid ${active?"#92400E":"#D1D5DB"}`,background:active?"#92400E":"#fff",color:active?"#fff":"#374151",fontSize:11,fontWeight:active?700:400,cursor:"pointer"}}>
                    {div.short}
                  </button>
                );
              })}
            </div>
            <div style={{fontSize:10,color:"#92400E",marginTop:4}}>
              {(()=>{const d=DIVISIONS.find(d=>d.id===(lv.divisionId||guessDivision(lv.name)));
                return d?.id==="p1"?"⏰ คาบ 6 = 13.50-14.40 | พัก | คาบ 7 = 14.50-15.40"
                  :"⏰ พักหลังคาบ 5 | คาบ 6 = 14.00-14.50 | คาบ 7 = 14.50-15.40";
              })()}
            </div>
          </div>
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
              <button onClick={()=>{setEditAssign(a);setEditForm({roomIds:[...(a.roomIds||[])],totalPeriods:a.totalPeriods});}} style={{background:"none",border:"none",cursor:"pointer",color:"#2563EB",padding:2}}><Icon name="edit" size={14}/></button>
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
    {editAssign&&(()=>{
      const eSub=S.subjects.find(s=>s.id===editAssign.subjectId);
      const eRooms=eSub?.levelId?S.rooms.filter(r=>r.levelId===eSub.levelId):S.rooms;
      const autoTP=(eSub?.periodsPerWeek||1)*Math.max(editForm.roomIds.length,1);
      return(
        <Modal open={!!editAssign} onClose={()=>setEditAssign(null)} title={"✏️ แก้ไข — "+(eSub?.code||"")+" "+(eSub?.name||"")}>
          <div style={{display:"flex",flexDirection:"column",gap:16}}>
            <div style={{background:"#F9FAFB",borderRadius:10,padding:"10px 14px"}}>
              <div style={{fontSize:14,fontWeight:700}}>{eSub?.code} — {eSub?.name}</div>
            </div>
            <div>
              <label style={LS}>ห้องเรียน <span style={{fontSize:11,color:"#9CA3AF"}}>(กดเลือก/ยกเลิก)</span></label>
              <div style={{display:"flex",flexWrap:"wrap",gap:6,maxHeight:180,overflowY:"auto"}}>
                {eRooms.map(rm=>{
                  const on=editForm.roomIds.includes(rm.id);
                  return <button key={rm.id} onClick={()=>setEditForm(p=>({...p,roomIds:on?p.roomIds.filter(r=>r!==rm.id):[...p.roomIds,rm.id]}))}
                    style={{padding:"5px 14px",borderRadius:20,border:"2px solid "+(on?"#DC2626":"#D1D5DB"),background:on?"#FEE2E2":"#fff",color:on?"#991B1B":"#374151",fontSize:12,fontWeight:on?700:400,cursor:"pointer"}}>{on?"✓ ":""}{rm.name}</button>;
                })}
              </div>
              <div style={{display:"flex",gap:6,marginTop:8}}>
                <button onClick={()=>setEditForm(p=>({...p,roomIds:eRooms.map(r=>r.id)}))} style={{fontSize:11,color:"#DC2626",background:"none",border:"1px solid #FECACA",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>เลือกทั้งหมด</button>
                <button onClick={()=>setEditForm(p=>({...p,roomIds:[]}))} style={{fontSize:11,color:"#6B7280",background:"none",border:"1px solid #E5E7EB",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>ล้าง</button>
                <span style={{fontSize:11,color:"#6B7280"}}>เลือก {editForm.roomIds.length} ห้อง</span>
              </div>
            </div>
            <div>
              <label style={LS}>จำนวนคาบ/สัปดาห์</label>
              <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                <input type="number" min="0" style={{...IS,width:100}} value={editForm.totalPeriods} onChange={e=>setEditForm(p=>({...p,totalPeriods:parseInt(e.target.value)||0}))}/>
                {eSub?.periodsPerWeek&&editForm.roomIds.length>0&&<button onClick={()=>setEditForm(p=>({...p,totalPeriods:autoTP}))} style={{fontSize:11,background:"#EFF6FF",color:"#1D4ED8",border:"1px solid #BFDBFE",borderRadius:8,padding:"4px 12px",cursor:"pointer"}}>อัตโนมัติ: {eSub.periodsPerWeek}×{editForm.roomIds.length}={autoTP}</button>}
              </div>
            </div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setEditAssign(null)} style={{...BO(),flex:1}}>ยกเลิก</button>
              <button disabled={editForm.roomIds.length===0} onClick={()=>{
                const finalTP=editForm.totalPeriods||autoTP||1;
                U.setAssigns(p=>p.map(x=>x.id===editAssign.id?{...x,roomIds:editForm.roomIds,totalPeriods:finalTP}:x));
                setEditAssign(null);st("แก้ไขสำเร็จ ✓");
              }} style={{...BS(),flex:2,opacity:editForm.roomIds.length===0?0.4:1}}>💾 บันทึก</button>
            </div>
          </div>
        </Modal>
      );
    })()}
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
            {/* ชื่อครู + ครูร่วม ใน compact */}
            {et&&<div style={{fontSize:9,color:dimmed?"#9CA3AF":c.tx,opacity:0.75,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
              {et.firstName}{coTeachers.length>0&&<span style={{color:"#7C3AED",fontWeight:700}}>{" +"+coTeachers.map(t=>t.firstName).join(",")}</span>}
            </div>}
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
                            className="dz"
                            onDragOver={e=>{const d=dragRef.current;if(!d){e.currentTarget.classList.remove("over");return;}
if(d.fromRoomId&&d.fromRoomId!==rid){e.currentTarget.classList.remove("over");return;}
if(d.assignmentId){const a=S.assigns.find(x=>x.id===d.assignmentId);const sCa=S.subjects.find(s=>s.id===d.subjectId)?.consecutiveAllowed||0;const ok=a?.roomIds?.includes(rid)||(sCa===-2&&S.assigns.some(x=>x.subjectId===d.subjectId&&x.roomIds?.includes(rid)));if(!ok){e.currentTarget.classList.remove("over");return;}}
e.preventDefault();e.currentTarget.classList.add("over");}}
                            onDragLeave={e=>e.currentTarget.classList.remove("over")}
                            onDrop={e=>{e.preventDefault();e.currentTarget.classList.remove("over");handleDrop(rid,day,p.id);}}
                            style={{padding:3,verticalAlign:"top",minHeight:68,borderLeft:`1px solid ${lc.border}`,borderBottom:`1px solid ${lc.border}`,background:customLock?"#FFF3E0":bl?"#FEF9C3":lk?"#F0FDF4":"inherit"}}
                          >
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
                    {(()=>{const tDiv=selT?getDivisionForTeacher(selT,S):"m2";const pListW=getPeriodCfg(tDiv).periods;return pListW.map(p=>(
                      <th key={p.id} style={{padding:"5px 3px",background:"#1E3A5F",textAlign:"center",borderLeft:"1px solid rgba(255,255,255,0.15)",position:"sticky",top:0,zIndex:2}}>
                        <div style={{fontSize:12,color:"#fff",fontWeight:700}}>คาบ {p.id}</div>
                        <div style={{fontSize:9,color:"rgba(255,255,255,0.65)"}}>{p.time}</div>
                      </th>
                    ));})()}
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
                        {(() => {
                          const customLock=(S.meetings||[]).find(m=>m.type==="custom"&&(m.slots||[]).some(s=>s.day===day&&s.period===p.id));
                          if(customLock) return (
                            <div style={{background:"#FFF3E0",color:"#E65100",fontSize:10,borderRadius:6,padding:"3px 5px",fontWeight:700}}>
                              🏫 {customLock.name}
                            </div>
                          );
                          if(blk) return (
                            <div style={{background:"#FEF9C3",color:"#92400E",fontSize:10,borderRadius:6,padding:"3px 5px",fontWeight:700}}>
                              🔒{S.meetings.some(m=>m.day===day&&m.periods?.includes(p.id)&&m.departmentId===teacher?.departmentId)?"ประชุม":blocked(selT).find(b=>b.day===day&&b.period===p.id)?.reason||"ล็อค"}
                            </div>
                          );
                          if(roomsThisPeriod.length>0) return roomsThisPeriod.map((r,i)=>(
                            <div key={i} style={{background:"#FEF2F2",border:"1px solid #FECACA",borderRadius:6,padding:"4px 6px",marginBottom:i<roomsThisPeriod.length-1?2:0}}>
                              <div style={{fontSize:11,fontWeight:800,color:"#1E40AF"}}>{r.rmName}</div>
                              <div style={{fontSize:10,color:"#374151",fontWeight:600}}>{r.subName}</div>
                            </div>
                          ));
                          return <span style={{color:"#D1D5DB",fontSize:12}}>—</span>;
                        })()}
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
                        {coTeachers2.map((ct2)=>{
                          const isFromSchedule=coIdsFromSchedule.has(ct2.id);
                          return<div key={ct2.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3}}>
                            <span style={{fontSize:10,color:lc.tx}}>
                              {ct2.firstName} {ct2.lastName}
                              {isFromSchedule&&<span style={{fontSize:8,color:"#059669",marginLeft:3}}>📅ในตาราง</span>}
                            </span>
                            <button onClick={()=>{
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
  const DAYS_SW=["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
  // PERIODS_SW คำนวณตาม division ของครูที่เลือก (ใช้ default ก่อน, update เมื่อเลือกครู)
  const REASON_OPTS=["ติดธุระ","ลาป่วย","ลากิจ","ไปราชการ","ไปอบรม","อื่นๆ"];
  const DAY_IDX={จันทร์:1,อังคาร:2,พุธ:3,พฤหัสบดี:4,ศุกร์:5};
  const [teacherA,setTeacherA]=useState("");
  const [absentDateFrom,setAbsentDateFrom]=useState("");
  const [absentDateTo,setAbsentDateTo]=useState("");
  const [reason,setReason]=useState("ติดธุระ");
  const [reasonOther,setReasonOther]=useState("");
  const [absentSlots,setAbsentSlots]=useState([]);
  const [searched,setSearched]=useState(false);
  const [results,setResults]=useState([]);
  const [selected,setSelected]=useState({});
  const fmtDate=(d)=>{if(!d)return"___________";const[y,m,d2]=d.split("-");return d2+"/"+m+"/"+(parseInt(y)+543);};
  const getDayRange=(from,to)=>{
    if(!from)return[];const end=to||from;const result=[];const cur=new Date(from);const endD=new Date(end);
    while(cur<=endD){const dow=cur.getDay();if(dow>=1&&dow<=5)result.push({dateStr:cur.toISOString().split("T")[0],dayName:["อาทิตย์","จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์","เสาร์"][dow]});cur.setDate(cur.getDate()+1);}
    return result;
  };
  const absentRange=getDayRange(absentDateFrom,absentDateTo);
  const absentDayNames=new Set(absentRange.map(r=>r.dayName));
  // PERIODS_SW ตาม division ของครู A — อัปเดตเมื่อเลือกครู
  const tADivSW=teacherA?getDivisionForTeacher(teacherA,S):"m2";
  const PERIODS_SW=getPeriodCfg(tADivSW).periods.map(p=>({...p,time:p.time.replace(/-/g,"–")}));
  const getEntries=(tid,day,pid)=>{
    const out=[];
    Object.entries(S.schedule).forEach(([k,en])=>{
      if(!en?.length)return;const pts=k.split("_");
      if(pts[pts.length-2]!==day||parseInt(pts[pts.length-1])!==pid)return;
      en.forEach(e=>{
        const coIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);
        if(e.teacherId!==tid&&!coIds.includes(tid))return;
        const sub=S.subjects.find(s=>s.id===e.subjectId);const rid=pts.slice(0,-2).join("_");const rm=S.rooms.find(r=>r.id===rid);
        out.push({subId:e.subjectId,subName:sub?.shortName||sub?.name||sub?.code||"—",subFullName:sub?.name||sub?.code||"—",roomId:rid,roomName:rm?.name||"—"});
      });
    });
    return out;
  };
  const isFree=(tid,day,pid)=>{
    const t=S.teachers.find(x=>x.id===tid);
    if(!t)return false;

    // 1) คาบล็อคส่วนตัวครู (personalLocks)
    if((t.personalLocks||[]).some(pl=>pl.day===day&&(pl.periods||[]).includes(pid)))return false;

    // 2) หน้าที่พิเศษ (specialRoles → ฝ่ายวิชาการ/วินัย)
    const roleBlocked=(t.specialRoles||[]).some(rid=>{
      const role=SROLES.find(r=>r.id===rid);
      return (role?.blocked||[]).some(bl=>bl.day===day&&(bl.periods||[]).includes(pid));
    });
    if(roleBlocked)return false;

    // 3) คาบล็อคกลุ่มสาระ (type="dept" หรือไม่มี type)
    const deptBlocked=(S.meetings||[]).some(m=>{
      if(m.type&&m.type!=="dept")return false; // ข้าม custom/homeroom
      if(m.isAssembly||m.isHomeroom)return false;
      if(m.departmentId!==t.departmentId)return false;
      return m.day===day&&(m.periods||[]).includes(pid);
    });
    if(deptBlocked)return false;

    // 4) คาบล็อคทั้งโรงเรียน (type="custom" — ล็อคทุกคน)
    const customBlocked=(S.meetings||[]).some(m=>{
      if(m.type!=="custom")return false;
      return (m.slots||[]).some(sl=>sl.day===day&&sl.period===pid);
    });
    if(customBlocked)return false;

    // 5) meeting ส่วนตัวครู (teacherId ตรงกัน)
    if((S.meetings||[]).some(m=>m.teacherId===tid&&m.day===day&&(m.periods||[]).includes(pid)))return false;

    // 6) คาบที่ล็อคไว้ใน schedule
    const hasLockedSlot=Object.entries(S.schedule).some(([k,en])=>{
      if(!en?.length)return false;
      const pts=k.split("_");
      if(pts[pts.length-2]!==day||parseInt(pts[pts.length-1])!==pid)return false;
      if(!S.locks[k])return false;
      return en.some(e=>{const coIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);return e.teacherId===tid||coIds.includes(tid);});
    });
    if(hasLockedSlot)return false;

    // 7) มีคาบสอนอยู่แล้ว
    return Object.entries(S.schedule).every(([k,en])=>{
      if(!en?.length)return true;
      const pts=k.split("_");
      if(pts[pts.length-2]!==day||parseInt(pts[pts.length-1])!==pid)return true;
      return en.every(e=>{const coIds=e.coTeacherIds?.length?e.coTeacherIds:(e.coTeacherId?[e.coTeacherId]:[]);return e.teacherId!==tid&&!coIds.includes(tid);});
    });
  };
  const toggleSlot=(day,pid)=>{setAbsentSlots(prev=>{const has=prev.some(s=>s.day===day&&s.period===pid);return has?prev.filter(s=>!(s.day===day&&s.period===pid)):[...prev,{day,period:pid}];});setSearched(false);};
  const calcReturnDates=(returnDayName,anchorDate)=>{
    if(!anchorDate||!returnDayName)return[];
    const base=new Date(anchorDate);const baseIdx=base.getDay();const targetIdx=DAY_IDX[returnDayName]??1;
    const minDate=new Date(anchorDate);minDate.setDate(minDate.getDate()-14);
    const dates=[];
    for(let w=-2;w<=4;w++){const d=new Date(base);d.setDate(base.getDate()+(targetIdx-baseIdx)+w*7);if(d>=minDate)dates.push(d.toISOString().split("T")[0]);}
    return dates;
  };
  const doSearch=()=>{
    if(!teacherA){st("เลือกครู A ก่อน","error");return;}
    if(!absentSlots.length){st("เลือกคาบที่ครู A ไม่อยู่ก่อน","error");return;}
    const res=[];
    absentSlots.forEach(({day,period:pid})=>{
      const entriA=getEntries(teacherA,day,pid);if(!entriA.length)return;
      entriA.forEach(({subName,subFullName,roomId,roomName})=>{
        const candidates=S.teachers.filter(t=>{
          if(t.id===teacherA)return false;
          if(!S.assigns.some(a=>a.teacherId===t.id&&(a.roomIds||[]).includes(roomId)))return false;
          return isFree(t.id,day,pid);
        }).map(t=>{
          const returnSlots=[];
          DAYS_SW.forEach(rd=>{PERIODS_SW.forEach(rp=>{
            const bEntries=getEntries(t.id,rd,rp.id).filter(e=>e.roomId===roomId);
            if(!bEntries.length)return;
            if(!isFree(teacherA,rd,rp.id))return;
            if(!S.assigns.some(a=>a.teacherId===teacherA&&(a.roomIds||[]).includes(roomId)))return;
            const subB=bEntries[0];
            calcReturnDates(rd,absentDateFrom).forEach(calcDate=>{
              if(rd===day&&rp.id===pid&&calcDate===absentDateFrom)return;
              returnSlots.push({day:rd,period:rp.id,time:rp.time,calcDate,subBName:subB.subFullName||subB.subName,subBRoom:subB.roomName});
            });
          });});
          const seen=new Set();
          return{teacher:t,returnSlots:returnSlots.filter(s=>{const k=s.day+"_"+s.period+"_"+s.calcDate;if(seen.has(k))return false;seen.add(k);return true;})};
        }).filter(c=>c.returnSlots.length>0);
        res.push({day,period:pid,time:PERIODS_SW.find(p=>p.id===pid)?.time,subName,subFullName,roomId,roomName,candidates});
      });
    });
    setResults(res);setSelected({});setSearched(true);
    if(!res.length)st("ไม่พบครูที่สอนแทนได้","warning");
  };
  // ── สร้าง HTML ฟอร์มแลกคาบ (A4 แนวนอน) ──
  const buildSwapHtml=()=>{
    const filledKeys=results.map(r=>r.day+"_"+r.period+"_"+r.roomId).filter(k=>selected[k]);
    if(!filledKeys.length)return null;
    const tA=S.teachers.find(t=>t.id===teacherA);
    const school=sh?.name||"โรงเรียนดาราวิทยาลัย";
    const yr=ay?.year||"2568";const sem=ay?.semester||"1";
    const finalReason=reason==="อื่นๆ"?(reasonOther||"อื่นๆ"):reason;
    const logo=sh?.logo?'<img src="'+sh.logo+'" style="height:50px;vertical-align:middle;margin-right:10px;"/>':"";
    const deptA=S.depts.find(d=>d.id===tA?.departmentId);
    const absentRangeStr=absentDateTo&&absentDateTo!==absentDateFrom?fmtDate(absentDateFrom)+" — "+fmtDate(absentDateTo):fmtDate(absentDateFrom);
    const rows=filledKeys.map(k=>{const r=results.find(r=>r.day+"_"+r.period+"_"+r.roomId===k);const sel=selected[k];const tB=S.teachers.find(t=>t.id===sel.subTeacherId);
      // หาวันที่จริงของ slot นี้จาก absentRange (match ตาม dayName)
      const slotDate=absentRange.find(ar=>ar.dayName===r.day)?.dateStr||absentDateFrom;
      return{r,sel,tB,slotDate};});
    const tAName=(tA?.prefix||"")+(tA?.firstName||"")+" "+(tA?.lastName||"");
    const tableRows=rows.map(({r,sel,tB,slotDate},i)=>
      '<tr><td style="text-align:center">'+(i+1)+'</td>'+
      '<td style="text-align:center">'+r.day+'<br/><b>'+fmtDate(slotDate)+'</b><br/>คาบ '+r.period+'<br/><span style="font-size:10pt;color:#555;">('+r.time+')</span></td>'+
      '<td>'+(sel.subBName||r.subFullName||r.subName)+'<br/><span style="font-size:10pt;color:#555;">ห้อง '+(sel.subBRoom||r.roomName)+'</span></td>'+
      '<td><b>'+(tB?.prefix||"")+(tB?.firstName||"")+" "+(tB?.lastName||"")+'</b></td>'+
      '<td style="text-align:center">'+sel.subDay+'<br/><b>'+fmtDate(sel.calcDate||"")+'</b><br/>คาบ '+sel.subPeriod+'<br/><span style="font-size:10pt;color:#555;">('+( PERIODS_SW.find(p=>p.id===sel.subPeriod)?.time||"")+')</span></td>'+
      '<td>'+(r.subFullName||r.subName)+'<br/><span style="font-size:10pt;color:#555;">ห้อง '+r.roomName+'</span></td>'+
      '<td></td></tr>'
    ).join("");
    return '<!DOCTYPE html><html><head><meta charset="utf-8"/>'+
      '<style>@page{size:A4 landscape;margin:10mm 12mm}*{box-sizing:border-box}'+
      "body{font-family:'TH SarabunNew','Sarabun',sans-serif;font-size:13pt;color:#000;margin:0}"+
      '.hdr{display:flex;align-items:center;justify-content:center;gap:10px;margin-bottom:4px}'+
      '.hdr h1{font-size:17pt;font-weight:700;margin:0}'+
      '.info{display:grid;grid-template-columns:repeat(4,1fr);gap:2px 12px;margin:8px 0;font-size:12pt}'+
      '.info .lbl{font-weight:700}'+
      'table{width:100%;border-collapse:collapse;font-size:12pt;margin:6px 0}'+
      'thead tr{background:#B91C1C;color:#fff}'+
      'thead th{padding:6px 8px;font-weight:700;text-align:center;border:1px solid #8B0000}'+
      'tbody tr:nth-child(even){background:#FFF5F5}'+
      'td{padding:5px 8px;border:1px solid #D1D5DB;vertical-align:middle}'+
      '.sigs{display:flex;justify-content:space-around;margin-top:14px}'+
      '.sig{flex:1;text-align:center}'+
      '.sig-line{display:block;width:85%;margin:0 auto 4px;border-bottom:1px solid #000}'+
      '@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}'+
      '</style></head><body>'+
      '<div class="hdr">'+logo+'<div><h1>แบบฟอร์มขอแลกเปลี่ยนคาบสอน / สอนแทน</h1>'+
      '<div style="text-align:center;font-size:11pt;color:#444">'+school+' | ภาคเรียนที่ '+sem+'/'+yr+'</div></div></div>'+
      '<div class="info">'+
        '<div><span class="lbl">ครูผู้ขอแลก: </span>'+tAName+'</div>'+
        '<div><span class="lbl">กลุ่มสาระ: </span>'+(deptA?.name||"—")+'</div>'+
        '<div><span class="lbl">วันที่ไม่อยู่: </span>'+absentRangeStr+'</div>'+
        '<div><span class="lbl">เหตุผล: </span>'+finalReason+'</div>'+
      '</div>'+
      '<table><thead><tr>'+
        '<th style="width:3%">#</th>'+
        '<th style="width:13%">คาบที่ขอ</th>'+
        '<th style="width:18%">วิชา/ห้อง (ที่ครูสอนแทน)</th>'+
        '<th style="width:15%">ครูสอนแทน</th>'+
        '<th style="width:13%">คาบที่ '+tAName+' สอนคืน</th>'+
        '<th style="width:18%">วิชา/ห้อง (ที่ '+tAName+' สอนคืน)</th>'+
        '<th style="width:20%">หมายเหตุ</th>'+
      '</tr></thead><tbody>'+tableRows+'</tbody></table>'+
      '<div class="sigs">'+
        '<div class="sig"><span class="sig-line"></span><div>'+tAName+'</div><div style="font-size:10pt;color:#555">ผู้ขอแลก วันที่ ___________</div></div>'+
        '<div class="sig"><span class="sig-line"></span><div>(............................)</div><div style="font-size:10pt;color:#555">หัวหน้ากลุ่มสาระ'+(deptA?.name?"<br/>"+deptA.name:"")+'</div></div>'+
      '</div></body></html>';
  };
  const printForm=()=>{
    const html=buildSwapHtml();
    if(!html){st("เลือกครูสอนแทนอย่างน้อย 1 คาบก่อน","error");return;}
    const w=window.open("","_blank");
    if(!w){st("Browser บล็อก popup","error");return;}
    w.document.write(html);w.document.close();setTimeout(()=>w.print(),500);
    st("กำลังเปิดหน้า print...");
  };
  const downloadSwapPDF=()=>{
    const html=buildSwapHtml();
    if(!html){st("เลือกครูสอนแทนอย่างน้อย 1 คาบก่อน","error");return;}
    const tA=S.teachers.find(t=>t.id===teacherA);
    const tAName=(tA?.prefix||"")+(tA?.firstName||"")+" "+(tA?.lastName||"");
    const blob=new Blob([html],{type:"text/html;charset=utf-8"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;
    a.download="แลกคาบ_"+tAName.trim()+"_"+fmtDate(absentDateFrom)+".html";
    a.click();URL.revokeObjectURL(url);
    st("ดาวน์โหลดไฟล์แล้ว — เปิดไฟล์แล้วสั่ง Print → Save as PDF");
  };
  return(
    <div style={{animation:"fadeIn 0.3s",display:"flex",flexDirection:"column",gap:12,maxWidth:680,margin:"0 auto",padding:"0 0 24px"}}>
      <style>{`@media(max-width:520px){.swap-period-grid{grid-template-columns:repeat(4,1fr)!important}}`}</style>

      {/* ── ขั้นที่ 1 ── */}
      <div style={{background:"#fff",borderRadius:16,border:"0.5px solid #E5E7EB",padding:"20px 20px 16px"}}>

        {/* step header */}
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16}}>
          <div style={{width:28,height:28,borderRadius:"50%",background:"#B91C1C",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:600,flexShrink:0}}>1</div>
          <div>
            <div style={{fontSize:15,fontWeight:600,color:"#111"}}>ครูที่ขอแลก และคาบที่ไม่อยู่</div>
            <div style={{fontSize:11,color:"#9CA3AF",marginTop:1}}>เลือกครู → เหตุผล → วันที่ → คาบ</div>
          </div>
        </div>

        {/* ครู */}
        <div style={{marginBottom:14}}>
          <div style={{fontSize:12,color:"#6B7280",marginBottom:5}}>ครูผู้ขอแลก</div>
          <SearchSelect value={teacherA} onChange={v=>{setTeacherA(v);setAbsentSlots([]);setSearched(false);}}
            options={[{value:"",label:"-- เลือกครู --"},...S.teachers.map(t=>({value:t.id,label:t.prefix+t.firstName+" "+t.lastName}))]}
            placeholder="-- เลือกครู --"/>
        </div>

        {/* เหตุผล */}
        <div style={{marginBottom:14}}>
          <div style={{fontSize:12,color:"#6B7280",marginBottom:5}}>เหตุผล</div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {REASON_OPTS.map(r=>(
              <button key={r} onClick={()=>setReason(r)} style={{
                padding:"6px 13px",borderRadius:20,cursor:"pointer",fontFamily:"inherit",
                border:"1.5px solid "+(reason===r?"#B91C1C":"#E5E7EB"),
                background:reason===r?"#FEF2F2":"#fff",
                color:reason===r?"#991B1B":"#6B7280",
                fontSize:13,fontWeight:reason===r?600:400,
                transition:"all 0.12s"
              }}>{r}</button>
            ))}
          </div>
          {reason==="อื่นๆ"&&<input style={{...IS,marginTop:8}} value={reasonOther} onChange={e=>setReasonOther(e.target.value)} placeholder="ระบุเหตุผล..."/>}
        </div>

        {/* วันที่ */}
        <div style={{marginBottom:10}}>
          <div style={{fontSize:12,color:"#6B7280",marginBottom:5}}>วันที่ไม่อยู่</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            <div>
              <div style={{fontSize:11,color:"#9CA3AF",marginBottom:3}}>เริ่มต้น</div>
              <input type="date" style={{...IS,fontSize:14}} value={absentDateFrom}
                onChange={e=>{setAbsentDateFrom(e.target.value);setAbsentSlots([]);setSearched(false);}}/>
            </div>
            <div>
              <div style={{fontSize:11,color:"#9CA3AF",marginBottom:3}}>สิ้นสุด (ถ้ามากกว่า 1 วัน)</div>
              <input type="date" style={{...IS,fontSize:14}} value={absentDateTo}
                min={absentDateFrom||undefined} onChange={e=>{setAbsentDateTo(e.target.value);setAbsentSlots([]);setSearched(false);}}/>
            </div>
          </div>
          {absentRange.length>0&&(
            <div style={{display:"flex",gap:6,flexWrap:"wrap",marginTop:8}}>
              {absentRange.map(r=>(
                <span key={r.dateStr} style={{display:"inline-flex",alignItems:"center",gap:4,background:"#FEF2F2",color:"#991B1B",padding:"3px 10px",borderRadius:20,fontSize:12,fontWeight:500}}>
                  📌 {r.dayName} {fmtDate(r.dateStr)}
                </span>
              ))}
            </div>
          )}
        </div>

        {/* ยังไม่เลือกครู */}
        {!teacherA&&(
          <div style={{borderRadius:10,padding:"18px 16px",textAlign:"center",color:"#9CA3AF",fontSize:13,border:"1px dashed #E5E7EB",marginTop:4,background:"#FAFAFA"}}>
            เลือกครูก่อน เพื่อแสดงตารางคาบสอน
          </div>
        )}

        {/* ตารางคาบ */}
        {teacherA&&(
          <div style={{marginTop:4}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8,borderTop:"0.5px solid #F3F4F6",paddingTop:12}}>
              <span style={{fontSize:13,fontWeight:500,color:"#374151"}}>เลือกคาบที่ไม่อยู่</span>
              <span style={{fontSize:11,color:"#9CA3AF"}}>แตะ/คลิกที่คาบที่มีวิชา</span>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:6}}>
              {DAYS_SW.map(day=>{
                const inRange=absentRange.length===0||absentDayNames.has(day);
                const daySlots=PERIODS_SW.map(p=>({p,ents:getEntries(teacherA,day,p.id),picked:absentSlots.some(s=>s.day===day&&s.period===p.id)}));
                const dayHasClass=daySlots.some(s=>s.ents.length>0);
                const dayPickedCount=daySlots.filter(s=>s.picked).length;
                const showFull=inRange||absentRange.length===0;

                return (
                  <div key={day} style={{
                    borderRadius:10,
                    border:"0.5px solid "+(dayPickedCount>0?"#FECACA":showFull&&dayHasClass?"#FED7AA":"#E5E7EB"),
                    background:dayPickedCount>0?"#FFF5F5":showFull&&dayHasClass?"#FFFBEB":"#FAFAFA",
                    overflow:"hidden",
                    opacity:!showFull&&absentRange.length>0?0.4:1,
                  }}>
                    {/* วัน header */}
                    <div style={{
                      padding:"8px 12px",display:"flex",alignItems:"center",justifyContent:"space-between",
                      borderBottom:showFull&&dayHasClass?"0.5px solid "+(dayPickedCount>0?"#FECACA":"#FED7AA"):"none"
                    }}>
                      <div style={{display:"flex",alignItems:"center",gap:7}}>
                        <span style={{fontSize:13,fontWeight:600,color:dayPickedCount>0?"#991B1B":showFull&&dayHasClass?"#92400E":"#9CA3AF"}}>
                          {day}
                        </span>
                        {!dayHasClass&&<span style={{fontSize:11,color:"#C4B5A5"}}>ไม่มีคาบสอน</span>}
                        {!showFull&&absentRange.length>0&&<span style={{fontSize:11,color:"#C4B5A5"}}>ไม่ใช่วันที่เลือก</span>}
                      </div>
                      {dayPickedCount>0&&(
                        <span style={{background:"#FEE2E2",color:"#991B1B",padding:"2px 9px",borderRadius:20,fontSize:11,fontWeight:500}}>
                          ✓ {dayPickedCount} คาบ
                        </span>
                      )}
                    </div>

                    {/* คาบ grid — 7 cols บนคอม, 4 cols บนมือถือ */}
                    {showFull&&dayHasClass&&(
                      <div className="swap-period-grid" style={{
                        display:"grid",
                        gridTemplateColumns:"repeat(7,1fr)",
                        gap:5,padding:"8px 10px 10px"
                      }}>
                        {daySlots.map(({p,ents,picked})=>{
                          const hasClass=ents.length>0;
                          const canClick=hasClass&&(inRange||absentRange.length===0);
                          return (
                            <button key={p.id} onClick={()=>canClick&&toggleSlot(day,p.id)}
                              disabled={!canClick}
                              style={{
                                borderRadius:8,padding:"8px 4px",textAlign:"center",
                                cursor:canClick?"pointer":"default",fontFamily:"inherit",
                                border:picked?"1.5px solid #DC2626":hasClass&&canClick?"0.5px solid #FED7AA":"0.5px solid #E5E7EB",
                                background:picked?"#FEE2E2":hasClass&&canClick?"#FFFBEB":"#fff",
                                opacity:hasClass?1:0.35,
                                transition:"all 0.12s",
                                minWidth:0,
                              }}>
                              <div style={{fontSize:12,fontWeight:600,color:"#374151"}}>คาบ {p.id}</div>
                              <div style={{fontSize:9,color:"#9CA3AF",margin:"1px 0 2px",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{p.time}</div>
                              {ents.map((e,i)=>(
                                <div key={i}>
                                  <div style={{fontSize:11,fontWeight:600,color:picked?"#991B1B":canClick?"#1E40AF":"#C4B5A5",lineHeight:1.2,overflow:"hidden",display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical"}}>
                                    {e.subName}
                                  </div>
                                  <div style={{fontSize:10,color:"#9CA3AF"}}>{e.roomName}</div>
                                </div>
                              ))}
                              {picked&&<div style={{fontSize:9,color:"#DC2626",fontWeight:600,marginTop:2}}>✕ ขอแลก</div>}
                            </button>
                          );
                        })}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            <button onClick={doSearch} style={{
              width:"100%",marginTop:12,padding:"12px",borderRadius:10,border:"none",
              background:absentSlots.length>0?"#B91C1C":"#9CA3AF",
              color:"#fff",fontSize:14,fontWeight:600,cursor:absentSlots.length>0?"pointer":"default",
              fontFamily:"inherit",transition:"background 0.15s",
              boxShadow:absentSlots.length>0?"0 2px 8px rgba(185,28,28,0.25)":"none"
            }}>
              {absentSlots.length>0?`🔍 ค้นหาครูสอนแทน (${absentSlots.length} คาบ)`:"เลือกคาบที่ไม่อยู่ก่อน"}
            </button>
          </div>
        )}
      </div>

      {/* ── ขั้นที่ 2 ── */}
      {searched&&(
        <div style={{background:"#fff",borderRadius:16,border:"0.5px solid #E5E7EB",padding:"20px 20px 16px"}}>

          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
            <div style={{width:28,height:28,borderRadius:"50%",background:"#059669",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:600,flexShrink:0}}>2</div>
            <div>
              <div style={{fontSize:15,fontWeight:600,color:"#111"}}>เลือกครูสอนแทน</div>
              <div style={{fontSize:11,color:"#9CA3AF",marginTop:1}}>เลือกครูและคาบที่สอนคืน</div>
            </div>
          </div>

          {results.length===0
            ?<div style={{textAlign:"center",padding:"24px 0",color:"#9CA3AF",fontSize:13}}>ไม่พบครูที่สอนแทนได้</div>
            :results.map(r=>{
              const key=r.day+"_"+r.period+"_"+r.roomId;
              const sel=selected[key];
              return(
                <div key={key} style={{
                  marginBottom:10,
                  border:"0.5px solid "+(sel?"#6EE7B7":"#E5E7EB"),
                  borderRadius:12,overflow:"hidden"
                }}>
                  {/* คาบ header */}
                  <div style={{
                    background:sel?"#F0FDF4":"#FFF5F5",
                    padding:"9px 14px",
                    borderBottom:"0.5px solid "+(sel?"#A7F3D0":"#FECACA"),
                    display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"
                  }}>
                    <span style={{
                      background:sel?"#D1FAE5":"#FEE2E2",
                      color:sel?"#065F46":"#991B1B",
                      padding:"3px 10px",borderRadius:20,fontSize:12,fontWeight:600
                    }}>{r.day} คาบ {r.period}</span>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:13,fontWeight:600,color:"#111"}}>{r.subName}</div>
                      <div style={{fontSize:11,color:"#9CA3AF"}}>ห้อง {r.roomName}</div>
                    </div>
                    {sel&&<span style={{
                      background:"#D1FAE5",color:"#065F46",
                      padding:"2px 9px",borderRadius:20,fontSize:11,fontWeight:600
                    }}>✅ เลือกแล้ว</span>}
                  </div>

                  {r.candidates.length===0
                    ?<div style={{padding:"12px 14px",color:"#9CA3AF",fontSize:12}}>ไม่มีครูว่างในเงื่อนไข</div>
                    :<div style={{padding:"10px 12px",display:"flex",flexDirection:"column",gap:8}}>
                      {r.candidates.map(({teacher:tB,returnSlots})=>(
                        <div key={tB.id} style={{
                          borderRadius:10,padding:"10px 12px",
                          border:"0.5px solid "+(sel?.subTeacherId===tB.id?"#6EE7B7":"#E5E7EB"),
                          background:sel?.subTeacherId===tB.id?"#F0FDF4":"#FAFAFA"
                        }}>
                          {/* Teacher header + load bar */}
                          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6}}>
                            <div style={{fontSize:13,fontWeight:600,color:"#1E3A5F"}}>
                              {tB.prefix}{tB.firstName} {tB.lastName}
                            </div>
                            {(()=>{
                              let totalB=0;
                              DAYS_SW.forEach(d=>PERIODS_SW.forEach(p=>{if(getEntries(tB.id,d,p.id).length>0)totalB++;}));
                              const cap=tB.totalPeriods||0;
                              const pct=cap>0?Math.round(totalB/cap*100):0;
                              const col=pct>=90?"#DC2626":pct>=70?"#D97706":"#059669";
                              return(
                                <div style={{display:"flex",alignItems:"center",gap:5}}>
                                  <span style={{fontSize:10,fontWeight:700,color:col}}>{totalB}/{cap||"?"}</span>
                                  <div style={{width:36,height:5,background:"#E5E7EB",borderRadius:3,overflow:"hidden"}}>
                                    <div style={{width:`${Math.min(pct,100)}%`,height:"100%",background:col,borderRadius:3,transition:"width 0.3s"}}/>
                                  </div>
                                </div>
                              );
                            })()}
                          </div>

                          {/* ── Mini Timetable ── */}
                          <div style={{overflowX:"auto",marginBottom:8,borderRadius:6,border:"1px solid #E5E7EB"}}>
                            <table style={{borderCollapse:"collapse",width:"100%",fontSize:9}}>
                              <thead>
                                <tr>
                                  <th style={{padding:"2px 5px",background:"#7F1D1D",color:"#fff",fontSize:8,width:30,textAlign:"center"}}>วัน╲คาบ</th>
                                  {PERIODS_SW.map(p=>(
                                    <th key={p.id} style={{padding:"2px 2px",background:"#B91C1C",color:"#fff",fontSize:8,textAlign:"center",minWidth:24}}>
                                      {p.id}
                                    </th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {DAYS_SW.map((d,di)=>(
                                  <tr key={d} style={{background:di%2===1?"#FFF9F9":"#fff"}}>
                                    <td style={{padding:"2px 4px",fontSize:8,fontWeight:600,color:"#374151",whiteSpace:"nowrap",borderRight:"1px solid #E5E7EB",background:"#FEF2F2",textAlign:"center"}}>{d.slice(0,3)}</td>
                                    {PERIODS_SW.map(p=>{
                                      const ents=getEntries(tB.id,d,p.id);
                                      const isSubSlot=d===r.day&&p.id===r.period;
                                      const isRetSlot=sel?.subTeacherId===tB.id&&d===sel?.subDay&&p.id===sel?.subPeriod;
                                      const locked=!isFree(tB.id,d,p.id)&&!ents.length;
                                      return(
                                        <td key={p.id} style={{
                                          padding:"1px",textAlign:"center",border:"1px solid #F3F4F6",
                                          background:isSubSlot?"#FDE68A":isRetSlot?"#A7F3D0":ents.length?"#BFDBFE":locked?"#F1F5F9":"#fff",
                                          minWidth:24,height:22,
                                        }} title={ents[0]?.subName||""}>
                                          {isSubSlot&&<span style={{color:"#B45309",fontWeight:800,fontSize:10}}>★</span>}
                                          {isRetSlot&&!isSubSlot&&<span style={{color:"#065F46",fontWeight:800,fontSize:10}}>✓</span>}
                                          {!isSubSlot&&!isRetSlot&&ents.length>0&&<span style={{color:"#1E40AF",fontWeight:700,fontSize:9}}>{ents[0].subName?.slice(0,2)||"■"}</span>}
                                          {!isSubSlot&&!isRetSlot&&locked&&<span style={{color:"#CBD5E1",fontSize:8}}>🔒</span>}
                                        </td>
                                      );
                                    })}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                            <div style={{display:"flex",gap:8,padding:"3px 5px",flexWrap:"wrap",background:"#FAFAFA",borderTop:"1px solid #F3F4F6"}}>
                              {[["#FDE68A","★ คาบสอนแทน"],["#BFDBFE","■ สอนอยู่"],["#A7F3D0","✓ คาบสอนคืน"],["#F1F5F9","🔒 ล็อค"]].map(([bg,lbl])=>(
                                <div key={lbl} style={{display:"flex",alignItems:"center",gap:2,fontSize:8,color:"#6B7280"}}>
                                  <div style={{width:8,height:8,background:bg,border:"1px solid #E5E7EB",borderRadius:1,flexShrink:0}}/>
                                  {lbl}
                                </div>
                              ))}
                            </div>
                          </div>

                          {/* return slots — ยุบตาม day+period แล้วเลือกวันที่ */}
                          <div style={{fontSize:11,fontWeight:600,color:"#374151",marginBottom:5}}>เลือกคาบสอนคืน:</div>
                          {(()=>{
                            const groups={};
                            returnSlots.forEach(rs=>{
                              const gk=rs.day+"_"+rs.period;
                              if(!groups[gk]) groups[gk]={day:rs.day,period:rs.period,time:rs.time,subBName:rs.subBName,subBRoom:rs.subBRoom,dates:[]};
                              groups[gk].dates.push(rs.calcDate);
                            });
                            const groupList=Object.values(groups);
                            const selGk=sel?.subTeacherId===tB.id?(sel.subDay+"_"+sel.subPeriod):null;
                            return(
                              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                                {groupList.map((g,gi)=>{
                                  const isGrpAct=selGk===g.day+"_"+g.period;
                                  return(
                                    <div key={gi} style={{borderRadius:8,border:`1.5px solid ${isGrpAct?"#059669":"#E5E7EB"}`,background:isGrpAct?"#F0FDF4":"#fff",overflow:"hidden"}}>
                                      <button onClick={()=>{
                                        if(isGrpAct){setSelected(p=>{const n={...p};delete n[key];return n;});}
                                        else{setSelected(p=>({...p,[key]:{subTeacherId:tB.id,subDay:g.day,subPeriod:g.period,calcDate:g.dates[0],subBName:g.subBName,subBRoom:g.subBRoom}}));}
                                      }} style={{width:"100%",padding:"8px 10px",border:"none",background:"none",cursor:"pointer",textAlign:"left",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                                        <div>
                                          <span style={{fontSize:12,fontWeight:700,color:isGrpAct?"#065F46":"#374151"}}>{isGrpAct?"✓ ":""}{g.day} คาบ {g.period}</span>
                                          <span style={{fontSize:10,color:"#9CA3AF",marginLeft:6}}>{g.time}</span>
                                          {g.subBName&&<div style={{fontSize:10,color:"#1E40AF",marginTop:1}}>📚 {g.subBName}</div>}
                                        </div>
                                        <div style={{display:"flex",alignItems:"center",gap:4}}>
                                          <span style={{fontSize:10,color:"#6B7280",background:"#F3F4F6",padding:"1px 7px",borderRadius:10}}>{g.dates.length} วัน</span>
                                          <span style={{fontSize:11,color:"#9CA3AF"}}>{isGrpAct?"▲":"▼"}</span>
                                        </div>
                                      </button>
                                      {isGrpAct&&(
                                        <div style={{padding:"4px 10px 10px",borderTop:"1px solid #E5E7EB"}}>
                                          <div style={{fontSize:10,color:"#6B7280",marginBottom:5}}>เลือกวันที่สอนคืน:</div>
                                          <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
                                            {g.dates.map((dt,di)=>{
                                              const isDateAct=sel?.calcDate===dt;
                                              return(
                                                <button key={di} onClick={()=>setSelected(p=>({...p,[key]:{subTeacherId:tB.id,subDay:g.day,subPeriod:g.period,calcDate:dt,subBName:g.subBName,subBRoom:g.subBRoom}}))}
                                                  style={{padding:"4px 10px",borderRadius:6,fontSize:11,fontWeight:isDateAct?700:400,border:`1.5px solid ${isDateAct?"#059669":"#D1D5DB"}`,background:isDateAct?"#D1FAE5":"#fff",color:isDateAct?"#065F46":"#374151",cursor:"pointer"}}>
                                                  {isDateAct?"✓ ":""}{fmtDate(dt)}
                                                </button>
                                              );
                                            })}
                                          </div>
                                        </div>
                                      )}
                                    </div>
                                  );
                                })}
                              </div>
                            );
                          })()}
                        </div>
                      ))}
                    </div>
                  }
                </div>
              );
            })
          }

          {/* ปุ่ม */}
          {results.length>0&&(
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginTop:6}}>
              <button onClick={printForm} disabled={!Object.keys(selected).length}
                style={{
                  padding:"12px",borderRadius:10,border:"none",fontFamily:"inherit",fontSize:13,fontWeight:600,
                  background:Object.keys(selected).length?"#059669":"#9CA3AF",
                  color:"#fff",cursor:Object.keys(selected).length?"pointer":"default",
                  transition:"background 0.12s"
                }}>
                🖨️ พิมพ์ฟอร์ม ({Object.keys(selected).length} คาบ)
              </button>
              <button onClick={downloadSwapPDF} disabled={!Object.keys(selected).length}
                style={{
                  padding:"12px",borderRadius:10,border:"none",fontFamily:"inherit",fontSize:13,fontWeight:600,
                  background:Object.keys(selected).length?"#2563EB":"#9CA3AF",
                  color:"#fff",cursor:Object.keys(selected).length?"pointer":"default",
                  transition:"background 0.12s"
                }}>
                📥 โหลด PDF ({Object.keys(selected).length} คาบ)
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
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
    <div style={{background:"#fff",borderRadius:16,padding:24,boxShadow:"0 2px 12px rgba(0,0,0,0.07)"}}>
      <div style={{display:"flex",gap:8,marginBottom:20,flexWrap:"wrap",alignItems:"center"}}>
        {[["print","🖨️ Print Center"],["hours","⏱ ชั่วโมงสอน"],["conflicts","⚠️ ตรวจ Conflict"]].map(([v,l])=>(
          <button key={v} onClick={()=>setReportTab(v)} style={{padding:"8px 18px",borderRadius:10,fontWeight:700,fontSize:13,border:`2px solid ${reportTab===v?"#B91C1C":"#E5E7EB"}`,background:reportTab===v?"#B91C1C":"#fff",color:reportTab===v?"#fff":"#374151",cursor:"pointer"}}>{l}</button>
        ))}
        <button onClick={()=>setShowPrintDesigner(true)} style={{marginLeft:"auto",padding:"8px 18px",borderRadius:10,fontWeight:700,fontSize:13,border:"2px solid #7C3AED",background:"#F5F3FF",color:"#7C3AED",cursor:"pointer"}}>🎨 Print Designer</button>
        <button onClick={()=>setShowPrintSettings(true)} style={{fontSize:12,padding:"5px 14px",borderRadius:20,border:"2px solid #B91C1C",background:"#FEF2F2",color:"#B91C1C",cursor:"pointer",fontWeight:600}}>⚙️ ตั้งค่าการพิมพ์</button>
      </div>
      <PrintSettingsPanel open={showPrintSettings} onClose={()=>setShowPrintSettings(false)} onApply={(s)=>{setPrintSettings(s);st("บันทึกการตั้งค่าแล้ว ✓");}} />
      <PrintPreviewModal data={printPreview} onClose={()=>setPrintPreview(null)} ps={printSettings}/>

      {reportTab==="hours" && <TeacherHoursSummary S={S} />}
      {reportTab==="conflicts" && <ConflictPanel S={S} />}
      <div style={{display:reportTab==="print"?"block":"none"}}>

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
              <button onClick={()=>{setSelectedTeachersPDF([]);setTeacherSearchQ("");setShowNewTeacherPDF(true);}} style={{...BS("#7C3AED"),fontSize:12,padding:"7px 16px"}}>🆕 พิมพ์แบบใหม่ (2คน/หน้า)</button>
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
              <button onClick={()=>{setExcelSelectedRooms([]);setShowExcelModal(true);}} style={{...BS("#059669"),fontSize:12,padding:"7px 16px"}}><Icon name="download" size={13}/>📊 Excel ตารางห้อง</button>
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
          <div style={{background:"#fff",borderRadius:16,boxShadow:"0 20px 60px rgba(0,0,0,0.3)",width:"min(520px,94%)",maxHeight:"90vh",overflowY:"auto",padding:24,fontFamily:"inherit"}}>
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
                  return<button key={r.id}
                    onClick={()=>setExcelSelectedRooms(p=>sel?p.filter(id=>id!==r.id):[...p,r.id])}
                    style={{padding:"4px 12px",borderRadius:20,border:`2px solid ${sel?"#059669":"#E5E7EB"}`,background:sel?"#059669":"#fff",color:sel?"#fff":"#374151",fontSize:12,fontWeight:sel?700:400,cursor:"pointer"}}>
                    {r.name}
                  </button>;
                })}
              </div>
              <div style={{display:"flex",gap:6,marginTop:6,alignItems:"center"}}>
                <button onClick={()=>setExcelSelectedRooms(S.rooms.map(r=>r.id))} style={{fontSize:11,color:"#059669",background:"none",border:"1px solid #D1FAE5",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>เลือกทั้งหมด</button>
                <button onClick={()=>setExcelSelectedRooms([])} style={{fontSize:11,color:"#6B7280",background:"none",border:"1px solid #E5E7EB",borderRadius:6,padding:"2px 10px",cursor:"pointer"}}>ล้าง</button>
                <span style={{fontSize:11,color:"#6B7280"}}>เลือกแล้ว {excelSelectedRooms.length} ห้อง → {excelSelectedRooms.length} sheets</span>
              </div>
            </div>

            <div style={{display:"flex",gap:10,marginTop:20}}>
              <button onClick={()=>setShowExcelModal(false)} style={{...BO(),flex:1}}>ยกเลิก</button>
              <button
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
          <div style={{background:"#fff",borderRadius:16,boxShadow:"0 20px 60px rgba(0,0,0,0.3)",width:"min(560px,94%)",maxHeight:"90vh",overflowY:"auto",padding:24,fontFamily:"inherit"}}>
            <div style={{fontSize:16,fontWeight:800,marginBottom:4}}>🆕 พิมพ์ตารางสอนครูแบบใหม่</div>
            <div style={{fontSize:11,color:"#6B7280",marginBottom:16}}>A4 แนวตั้ง — 2 คนต่อหน้า · แสดงวิชา+ห้อง+ชื่ออังกฤษ</div>
            <div>
              <label style={LS}>เลือกครู (กดหลายคนได้)</label>
              <input
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
            <div style={{display:"flex",gap:10,marginTop:20,flexWrap:"wrap"}}>
              <button onClick={()=>setShowNewTeacherPDF(false)} style={{...BO(),flex:1,minWidth:80}}>ยกเลิก</button>
              <button
                onClick={()=>{
                  if(!selectedTeachersPDF.length){st("เลือกครูก่อน","error");return;}
                  printTeacherPDFNew(S.teachers.filter(t=>selectedTeachersPDF.includes(t.id)));
                }}
                style={{...BO("#7C3AED"),flex:1,minWidth:80,opacity:selectedTeachersPDF.length?1:0.4,fontSize:12}}>
                👁️ ดูตัวอย่าง
              </button>
              <button
                onClick={()=>{
                  if(!selectedTeachersPDF.length){st("เลือกครูก่อน","error");return;}
                  const list2=S.teachers.filter(t=>selectedTeachersPDF.includes(t.id));
                  setShowNewTeacherPDF(false);
                  setTimeout(()=>setPrintPreview({html:buildF2Html(list2,S,ay,sh,printSettings)}),50);
                }}
                style={{...BS("#7C3AED"),flex:2,minWidth:120,opacity:selectedTeachersPDF.length?1:0.4}}>
                🖨️ แบบ 2 — 2คน/หน้า ({Math.ceil(selectedTeachersPDF.length/2)} หน้า)
              </button>
              <button
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
        <select value={filterDept} onChange={e=>setFilterDept(e.target.value)} style={{padding:"7px 10px",borderRadius:8,border:"1px solid #D1D5DB",fontSize:13,flex:1,minWidth:160}}>
          <option value="">กลุ่มสาระทั้งหมด</option>
          {S.depts.map(d=><option key={d.id} value={d.id}>{d.name}</option>)}
        </select>
        <div style={{display:"flex",gap:6}}>
          {[["name","ชื่อ"],["used","คาบสอน"],["rem","คงเหลือ"]].map(([v,l])=>(
            <button key={v} onClick={()=>setSortBy(v)} style={{padding:"6px 12px",borderRadius:8,fontSize:12,fontWeight:600,border:`2px solid ${sortBy===v?"#B91C1C":"#E5E7EB"}`,background:sortBy===v?"#FEE2E2":"#fff",color:sortBy===v?"#B91C1C":"#6B7280",cursor:"pointer"}}>เรียง{l}</button>
          ))}
        </div>
      </div>

      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
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
            <div key={i} style={{background:"#fff",border:"1px solid #FCA5A5",borderRadius:10,padding:"10px 14px",display:"flex",gap:12,alignItems:"flex-start"}}>
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

/* ===== PRINT DESIGNER ===== */

// ===== Field options =====
const PD_FIELD_OPTIONS = [
  { value:"subject_name",  label:"ชื่อวิชา (เต็ม)" },
  { value:"subject_short", label:"ชื่อวิชา (ย่อ)" },
  { value:"subject_code",  label:"รหัสวิชา" },
  { value:"teacher_name",  label:"ชื่อครู" },
  { value:"teacher_fname", label:"ชื่อครู (ชื่อต้น)" },
  { value:"teacher_code",  label:"รหัสครู" },
  { value:"room_name",     label:"ห้องเรียน" },
  { value:"period_time",   label:"เวลา" },
  { value:"period_num",    label:"คาบที่" },
  { value:"custom_text",   label:"ข้อความกำหนดเอง" },
  { value:"empty",         label:"(ว่าง)" },
];

// ===== Column types =====
// type: "period" | "break" | "homeroom" | "assembly" | "custom"
const PD_DEFAULT_COLUMNS = [
  { id:"c_day",   type:"day",      label:"วัน",        width:55,  show:true },
  { id:"c_hr",    type:"homeroom", label:"Homeroom",   width:22,  show:false, vertical:true, bg:"#FFF9E6", textColor:"#92400E" },
  { id:"c_asm",   type:"assembly", label:"Assembly",   width:22,  show:false, vertical:true, bg:"#FFF9E6", textColor:"#92400E" },
  { id:"c_p1",    type:"period",   label:"คาบ 1",      width:100, show:true, periodId:1, timeLabel:"08.30-09.20" },
  { id:"c_p2",    type:"period",   label:"คาบ 2",      width:100, show:true, periodId:2, timeLabel:"09.20-10.10" },
  { id:"c_brk1",  type:"break",    label:"พัก",        width:22,  show:false, timeLabel:"10.10-10.25", vertical:true, bg:"#F3F4F6" },
  { id:"c_p3",    type:"period",   label:"คาบ 3",      width:100, show:true, periodId:3, timeLabel:"10.25-11.15" },
  { id:"c_p4",    type:"period",   label:"คาบ 4",      width:100, show:true, periodId:4, timeLabel:"11.15-12.05" },
  { id:"c_brk2",  type:"break",    label:"พัก",        width:22,  show:false, timeLabel:"12.05-13.00", vertical:true, bg:"#F3F4F6" },
  { id:"c_p5",    type:"period",   label:"คาบ 5",      width:100, show:true, periodId:5, timeLabel:"13.00-13.50" },
  { id:"c_brk3",  type:"break",    label:"พัก",        width:22,  show:false, timeLabel:"13.50-14.00", vertical:true, bg:"#F3F4F6" },
  { id:"c_p6",    type:"period",   label:"คาบ 6",      width:100, show:true, periodId:6, timeLabel:"14.00-14.50" },
  { id:"c_p7",    type:"period",   label:"คาบ 7",      width:100, show:true, periodId:7, timeLabel:"14.50-15.40" },
];

const PD_DEFAULT_LAYOUT = {
  // cell content rows
  cellRows: [
    { field:"subject_short", fontSize:13, bold:true,  align:"center", color:"#1F2937", customText:"" },
    { field:"teacher_name",  fontSize:11, bold:false, align:"center", color:"#4B5563", customText:"" },
    { field:"room_name",     fontSize:10, bold:false, align:"center", color:"#6B7280", customText:"" },
  ],
  // columns config
  columns: PD_DEFAULT_COLUMNS.map(c => ({...c})),
  // header style
  headerBg:      "#B91C1C",
  headerText:    "#fff",
  rowAltBg:      "#FFF5F5",
  showAltRow:    true,
  showBorder:    true,
  borderColor:   "#E5E7EB",
  fontFamily:    "TH SarabunNew",
  fontSize:      100,
  rowHeight:     100,
  showPeriodNum:  true,
  showPeriodTime: true,
  paperSize:      "A4",
  orientation:    "landscape",
  marginMm:       10,
  // header block
  showLogo:       true,
  logoSize:       48,
  logoPosition:   "left",   // left | center | right
  titleText:      "",
  titleFontSize:  16,
  subtitleText:   "",
  subtitleFontSize:12,
  showYear:       true,
  headerLayout:   "logo-left", // logo-left | logo-center | no-logo
  // footer
  showFooter:     true,
  footerLeft:     "ลงชื่อ _________________________ รองฝ่ายวิชาการ",
  footerRight:    "ลงชื่อ _________________________ ผู้อำนวยการ",
};

function loadPDLayouts() {
  try { const s = localStorage.getItem("dara_pdLayouts2"); return s ? JSON.parse(s) : {}; } catch { return {}; }
}
function savePDLayouts(layouts) {
  try { localStorage.setItem("dara_pdLayouts2", JSON.stringify(layouts)); } catch {}
}

function renderPDCell(entry, field, S, customText) {
  if (!entry) return "";
  const sub = S.subjects.find(s => s.id === entry.subjectId);
  const tch = S.teachers.find(t => t.id === entry.teacherId);
  switch (field) {
    case "subject_name":  return sub?.name || "";
    case "subject_short": return sub?.shortName || sub?.name || "";
    case "subject_code":  return sub?.code || "";
    case "teacher_name":  return tch ? `${tch.firstName||""} ${tch.lastName||""}`.trim() : "";
    case "teacher_fname": return tch?.firstName || "";
    case "teacher_code":  return tch?.teacherCode || "";
    case "room_name": {
      const room = S.rooms.find(r => r.id === entry.roomId) || S.specialRooms?.find(r => r.id === entry.specialRoomId);
      return room?.name || "";
    }
    case "period_time": {
      const p = PERIODS.find(p => String(p.id) === String(entry.periodId));
      return p?.time || "";
    }
    case "period_num": return entry.periodId ? `คาบ ${entry.periodId}` : "";
    case "custom_text": return customText || "";
    case "empty": return "";
    default: return "";
  }
}

// ===== Build full print HTML =====
function buildPDPrintHTML(layout, S, ay, sh, targetType, targetId) {
  const DAYS_TH = ["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
  const fScale = layout.fontSize / 100;
  const rScale = layout.rowHeight / 100;
  const hBg = layout.headerBg;
  const hTxt = layout.headerText;
  const bd = layout.showBorder ? `1px solid ${layout.borderColor||"#ddd"}` : "1px solid transparent";
  const ff = layout.fontFamily;
  const cols = layout.columns.filter(c => c.show);

  // ข้อมูลตาราง
  let targetName = "";
  let rowData = []; // [{label, homeroomText, assemblyText, cells:{periodId→entry}}]

  if (targetType === "room") {
    const room = S.rooms.find(r => r.id === targetId);
    if (!room) return "<p>ไม่พบห้องเรียน</p>";
    targetName = `ตารางเรียน${room.name}`;
    DAYS_TH.forEach(day => {
      const cells = {};
      PERIODS.forEach(p => {
        const k = `${room.id}_${day}_${p.id}`;
        cells[p.id] = (S.schedule[k]||[])[0] || null;
      });
      rowData.push({ label: day, cells, homeroomText:"", assemblyText:"" });
    });
  } else {
    const teacher = S.teachers.find(t => t.id === targetId);
    if (!teacher) return "<p>ไม่พบครู</p>";
    targetName = `ตารางสอน${teacher.firstName||""} ${teacher.lastName||""}`;
    DAYS_TH.forEach(day => {
      const cells = {};
      PERIODS.forEach(p => {
        let found = null;
        Object.entries(S.schedule||{}).forEach(([k, ens]) => {
          const pts = k.split("_");
          if (pts[pts.length-2] === day && parseInt(pts[pts.length-1]) === p.id) {
            ens?.forEach(e => {
              if ([e.teacherId,...(e.coTeacherIds||[])].includes(teacher.id))
                found = {...e, roomId: pts[0]};
            });
          }
        });
        cells[p.id] = found;
      });
      rowData.push({ label: day, cells, homeroomText:"", assemblyText:"" });
    });
  }

  const yr = ay?.year || (sh?.year) || "2569";
  const schoolName = sh?.schoolName || "โรงเรียนดาราวิทยาลัย";
  const titleText = layout.titleText || targetName;
  const subtitleText = layout.subtitleText || `ภาคเรียนที่ 1/${yr} ${schoolName}`;
  const logoData = sh?.logo || "";

  // Header HTML
  let headerHtml = "";
  if (layout.headerLayout === "logo-left" && logoData) {
    headerHtml = `<div class="hdr-wrap hdr-left">
      <img src="${logoData}" class="hdr-logo" style="width:${layout.logoSize||48}px;height:${layout.logoSize||48}px">
      <div class="hdr-text">
        <div class="hdr-title" style="font-size:${(layout.titleFontSize||16)*fScale}px">${titleText}</div>
        <div class="hdr-sub" style="font-size:${(layout.subtitleFontSize||12)*fScale}px">${subtitleText}</div>
      </div>
    </div>`;
  } else if (layout.headerLayout === "logo-center" && logoData) {
    headerHtml = `<div class="hdr-wrap hdr-center">
      <img src="${logoData}" class="hdr-logo" style="width:${layout.logoSize||48}px;height:${layout.logoSize||48}px">
      <div class="hdr-title" style="font-size:${(layout.titleFontSize||16)*fScale}px">${titleText}</div>
      <div class="hdr-sub" style="font-size:${(layout.subtitleFontSize||12)*fScale}px">${subtitleText}</div>
    </div>`;
  } else {
    headerHtml = `<div class="hdr-wrap hdr-center">
      <div class="hdr-title" style="font-size:${(layout.titleFontSize||16)*fScale}px">${titleText}</div>
      <div class="hdr-sub" style="font-size:${(layout.subtitleFontSize||12)*fScale}px">${subtitleText}</div>
    </div>`;
  }

  // Column headers
  const colHeaders = cols.map(col => {
    if (col.type === "day") {
      return `<th class="th-day" style="width:${col.width}px;background:${hBg};color:${hTxt};border:${bd}">
        <div>วัน</div><div style="font-size:${9*fScale}px;font-weight:400;opacity:0.8">\\เวลา</div>
      </th>`;
    }
    if (col.type === "break" || col.type === "homeroom" || col.type === "assembly") {
      return `<th style="width:${col.width}px;background:${col.bg||hBg};color:${col.textColor||hTxt};border:${bd};writing-mode:vertical-rl;text-orientation:mixed;font-size:${9*fScale}px;padding:4px 2px;white-space:nowrap">
        ${col.label || ""}${col.timeLabel ? `<br><span style="font-weight:400;font-size:${8*fScale}px">${col.timeLabel}</span>` : ""}
      </th>`;
    }
    if (col.type === "period") {
      return `<th style="width:${col.width}px;background:${hBg};color:${hTxt};border:${bd};text-align:center;padding:${5*rScale}px 4px">
        ${layout.showPeriodNum ? `<div style="font-size:${11*fScale}px;font-weight:700">${col.label}</div>` : ""}
        ${layout.showPeriodTime ? `<div style="font-size:${9*fScale}px;font-weight:400;opacity:0.85">${col.timeLabel||""}</div>` : ""}
      </th>`;
    }
    return `<th style="width:${col.width}px;background:${hBg};color:${hTxt};border:${bd}">${col.label||""}</th>`;
  }).join("");

  // Data rows
  const dataRows = rowData.map((row, ri) => {
    const altBg = layout.showAltRow && ri%2===1 ? layout.rowAltBg : "#fff";
    const cells = cols.map(col => {
      if (col.type === "day") {
        return `<td class="td-day" style="background:${hBg};color:${hTxt};border:${bd};font-size:${12*fScale}px;font-weight:700;text-align:center;padding:${6*rScale}px 4px">${row.label}</td>`;
      }
      if (col.type === "homeroom") {
        return `<td style="background:${col.bg||"#FFF9E6"};border:${bd};writing-mode:vertical-rl;text-orientation:mixed;font-size:${9*fScale}px;color:${col.textColor||"#92400E"};text-align:center;padding:2px">${row.homeroomText||col.label||"Homeroom"}</td>`;
      }
      if (col.type === "assembly") {
        return `<td style="background:${col.bg||"#FFF9E6"};border:${bd};writing-mode:vertical-rl;text-orientation:mixed;font-size:${9*fScale}px;color:${col.textColor||"#92400E"};text-align:center;padding:2px">${row.assemblyText||col.label||"Assembly"}</td>`;
      }
      if (col.type === "break") {
        return `<td style="background:${col.bg||"#F3F4F6"};border:${bd};writing-mode:vertical-rl;text-orientation:mixed;font-size:${8*fScale}px;color:#6B7280;text-align:center;padding:2px">${col.timeLabel||"พัก"}</td>`;
      }
      if (col.type === "period") {
        const entry = row.cells[col.periodId];
        if (!entry) return `<td style="background:${altBg};border:${bd};min-height:${40*rScale}px"></td>`;
        const lines = layout.cellRows.map(rowCfg => {
          if (rowCfg.field === "empty") return "";
          const val = renderPDCell(entry, rowCfg.field, S, rowCfg.customText);
          if (!val) return "";
          return `<div style="font-size:${rowCfg.fontSize*fScale}px;font-weight:${rowCfg.bold?"700":"400"};text-align:${rowCfg.align};color:${rowCfg.color};line-height:1.3">${val}</div>`;
        }).filter(Boolean).join("");
        return `<td style="background:${altBg};border:${bd};padding:${4*rScale}px 3px;vertical-align:middle;text-align:center">${lines}</td>`;
      }
      return `<td style="background:${altBg};border:${bd}"></td>`;
    }).join("");
    return `<tr>${cells}</tr>`;
  }).join("");

  // Footer
  const footerHtml = layout.showFooter ? `
    <div class="footer">
      <span>${layout.footerLeft||""}</span>
      <span>${layout.footerRight||""}</span>
    </div>` : "";

  // colgroup
  const colgroup = cols.map(col => `<col style="width:${col.width}px">`).join("");

  return `<!DOCTYPE html><html><head>
    <meta charset="UTF-8">
    <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
      @page { size: ${layout.paperSize||"A4"} ${layout.orientation||"landscape"}; margin: ${layout.marginMm||10}mm; }
      body { font-family: '${ff}','Sarabun',sans-serif; margin:0; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
      table { border-collapse:collapse; width:100%; }
      .hdr-wrap { margin-bottom:10px; }
      .hdr-left { display:flex; align-items:center; gap:12px; }
      .hdr-center { display:flex; flex-direction:column; align-items:center; text-align:center; gap:2px; }
      .hdr-logo { object-fit:contain; }
      .hdr-title { font-weight:700; }
      .hdr-sub { color:#555; }
      .td-day { white-space:nowrap; }
      .footer { display:flex; justify-content:space-between; margin-top:10px; font-size:${10*fScale}px; padding:0 8px; }
      @media print { body { -webkit-print-color-adjust:exact; } }
    </style>
  </head><body>
    ${headerHtml}
    <table>
      <colgroup>${colgroup}</colgroup>
      <thead><tr>${colHeaders}</tr></thead>
      <tbody>${dataRows}</tbody>
    </table>
    ${footerHtml}
  </body></html>`;
}

// ===== Preview Table Component =====
function PreviewTable({ layout, S, targetType, targetId }) {
  const DAYS_TH = ["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์"];
  const fScale = (layout.fontSize||100) / 100;
  const rScale = (layout.rowHeight||100) / 100;
  const hBg = layout.headerBg || "#B91C1C";
  const hTxt = layout.headerText || "#fff";
  const bd = layout.showBorder ? `1px solid ${layout.borderColor||"#E5E7EB"}` : "none";
  const cols = (layout.columns || PD_DEFAULT_COLUMNS).filter(c => c.show);

  let rowData = [];
  if (targetType === "room") {
    const room = S.rooms.find(r => r.id === targetId);
    if (!room) return null;
    DAYS_TH.forEach(day => {
      const cells = {};
      PERIODS.forEach(p => { const k=`${room.id}_${day}_${p.id}`; cells[p.id]=(S.schedule[k]||[])[0]||null; });
      rowData.push({ label:day, cells });
    });
  } else {
    const teacher = S.teachers.find(t => t.id === targetId);
    if (!teacher) return null;
    DAYS_TH.forEach(day => {
      const cells = {};
      PERIODS.forEach(p => {
        let found = null;
        Object.entries(S.schedule||{}).forEach(([k,ens]) => {
          const pts=k.split("_");
          if(pts[pts.length-2]===day&&parseInt(pts[pts.length-1])===p.id)
            ens?.forEach(e=>{ if([e.teacherId,...(e.coTeacherIds||[])].includes(teacher.id)) found={...e,roomId:pts[0]}; });
        });
        cells[p.id] = found;
      });
      rowData.push({ label:day, cells });
    });
  }

  return (
    <div style={{fontFamily:layout.fontFamily,overflowX:"auto"}}>
      <table style={{borderCollapse:"collapse",width:"100%"}}>
        <thead>
          <tr>
            {cols.map(col => {
              if (col.type==="day") return <th key={col.id} style={{background:hBg,color:hTxt,border:bd,width:col.width,padding:`${5*rScale}px 4px`,fontSize:11*fScale,textAlign:"center"}}>วัน</th>;
              if (col.type==="break"||col.type==="homeroom"||col.type==="assembly") return (
                <th key={col.id} style={{background:col.bg||hBg,color:col.textColor||hTxt,border:bd,width:col.width,writingMode:"vertical-rl",fontSize:8*fScale,padding:"4px 2px",textAlign:"center"}}>
                  {col.label}{col.timeLabel&&<span style={{fontSize:7*fScale,display:"block",opacity:0.8}}>{col.timeLabel}</span>}
                </th>
              );
              if (col.type==="period") return (
                <th key={col.id} style={{background:hBg,color:hTxt,border:bd,width:col.width,padding:`${4*rScale}px 4px`,textAlign:"center"}}>
                  {layout.showPeriodNum&&<div style={{fontSize:10*fScale,fontWeight:700}}>{col.label}</div>}
                  {layout.showPeriodTime&&<div style={{fontSize:8*fScale,opacity:0.85}}>{col.timeLabel}</div>}
                </th>
              );
              return <th key={col.id} style={{background:hBg,color:hTxt,border:bd,width:col.width}}>{col.label}</th>;
            })}
          </tr>
        </thead>
        <tbody>
          {rowData.map((row,ri) => {
            const altBg = layout.showAltRow&&ri%2===1 ? layout.rowAltBg : "#fff";
            return (
              <tr key={row.label}>
                {cols.map(col => {
                  if (col.type==="day") return <td key={col.id} style={{background:hBg,color:hTxt,border:bd,textAlign:"center",fontWeight:700,fontSize:11*fScale,padding:`${5*rScale}px 4px`,whiteSpace:"nowrap"}}>{row.label}</td>;
                  if (col.type==="homeroom"||col.type==="assembly") return <td key={col.id} style={{background:col.bg||"#FFF9E6",border:bd,writingMode:"vertical-rl",fontSize:8*fScale,color:col.textColor||"#92400E",textAlign:"center",padding:2}}>{col.label}</td>;
                  if (col.type==="break") return <td key={col.id} style={{background:col.bg||"#F3F4F6",border:bd,writingMode:"vertical-rl",fontSize:8*fScale,color:"#6B7280",textAlign:"center",padding:2}}>{col.timeLabel||"พัก"}</td>;
                  if (col.type==="period") {
                    const entry = row.cells[col.periodId];
                    if (!entry) return <td key={col.id} style={{background:altBg,border:bd,minHeight:35*rScale}}></td>;
                    return (
                      <td key={col.id} style={{background:altBg,border:bd,padding:`${3*rScale}px 3px`,verticalAlign:"middle",textAlign:"center"}}>
                        {layout.cellRows.map((rowCfg,rci) => {
                          if(rowCfg.field==="empty") return null;
                          const val = renderPDCell(entry, rowCfg.field, S, rowCfg.customText);
                          if(!val) return null;
                          return <div key={rci} style={{fontSize:rowCfg.fontSize*fScale,fontWeight:rowCfg.bold?"700":"400",textAlign:rowCfg.align,color:rowCfg.color,lineHeight:1.3}}>{val}</div>;
                        })}
                      </td>
                    );
                  }
                  return <td key={col.id} style={{background:altBg,border:bd}}></td>;
                })}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// ===== Main PrintDesignerModal =====
function PrintDesignerModal({ open, onClose, S, ay, sh }) {
  const [layouts, setLayouts] = useState(loadPDLayouts);
  const [activeName, setActiveName] = useState("");
  const [layout, setLayout] = useState({ ...PD_DEFAULT_LAYOUT, columns: PD_DEFAULT_COLUMNS.map(c=>({...c})) });
  const [previewTarget, setPreviewTarget] = useState({ type:"room", id:"" });
  const [newName, setNewName] = useState("");
  const [showSaveAs, setShowSaveAs] = useState(false);
  const [tab, setTab] = useState("cell"); // cell | columns | style | header

  useEffect(() => {
    if (!open) return;
    const names = Object.keys(loadPDLayouts());
    if (names.length && !activeName) {
      const saved = loadPDLayouts();
      setActiveName(names[0]);
      setLayout({ ...PD_DEFAULT_LAYOUT, columns: PD_DEFAULT_COLUMNS.map(c=>({...c})), ...saved[names[0]] });
    }
  }, [open]);

  if (!open) return null;

  const upd = (key, val) => setLayout(p => ({ ...p, [key]: val }));
  const updCol = (id, key, val) => setLayout(p => ({
    ...p,
    columns: p.columns.map(c => c.id === id ? { ...c, [key]: val } : c)
  }));
  const updCellRow = (i, key, val) => {
    const rows = [...layout.cellRows];
    rows[i] = { ...rows[i], [key]: val };
    upd("cellRows", rows);
  };
  const moveCellRow = (i, dir) => {
    const rows = [...layout.cellRows];
    const j = i + dir;
    if (j < 0 || j >= rows.length) return;
    [rows[i], rows[j]] = [rows[j], rows[i]];
    upd("cellRows", rows);
  };
  const addCellRow = () => upd("cellRows", [...layout.cellRows, { field:"empty", fontSize:11, bold:false, align:"center", color:"#6B7280", customText:"" }]);
  const removeCellRow = (i) => upd("cellRows", layout.cellRows.filter((_,ri) => ri !== i));

  const addCustomCol = () => {
    const newCol = { id:`c_custom_${Date.now()}`, type:"custom", label:"กำหนดเอง", width:80, show:true, bg:"#fff", textColor:"#333" };
    upd("columns", [...layout.columns, newCol]);
  };
  const addBreakCol = () => {
    const newCol = { id:`c_brk_${Date.now()}`, type:"break", label:"พัก", width:22, show:true, timeLabel:"", vertical:true, bg:"#F3F4F6" };
    upd("columns", [...layout.columns, newCol]);
  };
  const removeCol = (id) => upd("columns", layout.columns.filter(c => c.id !== id));
  const moveCol = (id, dir) => {
    const cols = [...layout.columns];
    const i = cols.findIndex(c => c.id === id);
    const j = i + dir;
    if (j < 0 || j >= cols.length) return;
    [cols[i], cols[j]] = [cols[j], cols[i]];
    upd("columns", cols);
  };

  const saveLayout = (name) => {
    const nl = { ...loadPDLayouts(), [name]: layout };
    setLayouts(nl); savePDLayouts(nl); setActiveName(name);
  };
  const deleteLayout = (name) => {
    const nl = { ...loadPDLayouts() }; delete nl[name];
    setLayouts(nl); savePDLayouts(nl);
    const rem = Object.keys(nl);
    if (rem.length) { setActiveName(rem[0]); setLayout({ ...PD_DEFAULT_LAYOUT, columns: PD_DEFAULT_COLUMNS.map(c=>({...c})), ...nl[rem[0]] }); }
    else { setActiveName(""); setLayout({ ...PD_DEFAULT_LAYOUT, columns: PD_DEFAULT_COLUMNS.map(c=>({...c})) }); }
  };

  const doPrint = () => {
    if (!previewTarget.id) return;
    const html = buildPDPrintHTML(layout, S, ay, sh, previewTarget.type, previewTarget.id);
    const w = window.open("", "_blank");
    w.document.write(html); w.document.close();
    setTimeout(() => w.print(), 600);
  };

  // Bulk print — พิมพ์ทุกห้องหรือทุกครูด้วย layout เดียวกัน
  const doPrintAll = (type) => {
    const targets = type === "room"
      ? S.rooms
      : S.teachers.filter(t => t.totalPeriods > 0);
    if (!targets.length) return;

    // สร้าง HTML ทุกหน้ารวมกัน โดยใส่ page-break ระหว่างกัน
    const pages = targets.map(target => {
      const id = target.id;
      // เอา body ของแต่ละหน้า
      const full = buildPDPrintHTML(layout, S, ay, sh, type, id);
      // แกะเอาแค่ส่วนใน <body>...</body>
      const bodyMatch = full.match(/<body>([\s\S]*?)<\/body>/);
      return bodyMatch ? bodyMatch[1] : full;
    });

    // ใช้ header/style จากหน้าแรก
    const first = buildPDPrintHTML(layout, S, ay, sh, type, targets[0].id);
    const headMatch = first.match(/([\s\S]*?<body>)/);
    const head = headMatch ? headMatch[1] : '<html><body>';
    const closeMatch = first.match(/<\/body>[\s\S]*$/);
    const close = closeMatch ? closeMatch[0] : '</body></html>';

    // เพิ่ม page-break-after ระหว่างแต่ละหน้า
    const combined = pages.map((p, i) =>
      i < pages.length - 1
        ? `<div style="page-break-after:always">${p}</div>`
        : `<div>${p}</div>`
    ).join('');

    const w = window.open("", "_blank");
    w.document.write(head + combined + close);
    w.document.close();
    setTimeout(() => w.print(), 800);
  };

  const IS2 = { width:"100%", padding:"6px 10px", border:"1.5px solid #E5E7EB", borderRadius:8, fontSize:12, outline:"none", fontFamily:"inherit", background:"#fff" };
  const LS = { fontSize:12, fontWeight:600, display:"block", marginBottom:4, color:"#374151" };
  const TABS = [["cell","📋 เนื้อหาช่อง"],["columns","↔️ คอลัมน์"],["style","🎨 สไตล์"],["header","🏷 หัว/ท้าย"]];

  return (
    <div style={{position:"fixed",inset:0,zIndex:4000,background:"rgba(0,0,0,0.65)",display:"flex",alignItems:"stretch",justifyContent:"flex-end"}}>
      {/* ===== LEFT: Preview ===== */}
      <div style={{flex:1,overflow:"auto",padding:20,display:"flex",flexDirection:"column",gap:12}}>
        <div style={{fontWeight:700,fontSize:15,color:"#fff"}}>👁 ตัวอย่าง</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center",background:"rgba(255,255,255,0.1)",borderRadius:10,padding:10}}>
          <select value={previewTarget.type} onChange={e=>setPreviewTarget(p=>({...p,type:e.target.value,id:""}))} style={{...IS2,width:"auto",background:"rgba(255,255,255,0.95)"}}>
            <option value="room">ห้องเรียน</option>
            <option value="teacher">ครู</option>
          </select>
          <select value={previewTarget.id} onChange={e=>setPreviewTarget(p=>({...p,id:e.target.value}))} style={{...IS2,flex:1,minWidth:150,background:"rgba(255,255,255,0.95)"}}>
            <option value="">-- เลือก --</option>
            {previewTarget.type==="room"
              ? S.rooms.map(r=><option key={r.id} value={r.id}>{r.name}</option>)
              : S.teachers.map(t=><option key={t.id} value={t.id}>{t.firstName} {t.lastName}</option>)}
          </select>
          <button onClick={doPrint} disabled={!previewTarget.id} style={{padding:"8px 20px",background:previewTarget.id?"#B91C1C":"#9CA3AF",color:"#fff",border:"none",borderRadius:8,fontWeight:700,cursor:previewTarget.id?"pointer":"default",fontSize:13,whiteSpace:"nowrap"}}>🖨 พิมพ์</button>
        </div>
        {/* Bulk print buttons */}
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>doPrintAll("teacher")} style={{flex:1,padding:"8px 12px",background:"#1D4ED8",color:"#fff",border:"none",borderRadius:8,fontWeight:700,fontSize:12,cursor:"pointer",whiteSpace:"nowrap"}}>
            🖨 พิมพ์ครูทุกคน ({S.teachers.filter(t=>t.totalPeriods>0).length} คน)
          </button>
          <button onClick={()=>doPrintAll("room")} style={{flex:1,padding:"8px 12px",background:"#059669",color:"#fff",border:"none",borderRadius:8,fontWeight:700,fontSize:12,cursor:"pointer",whiteSpace:"nowrap"}}>
            🖨 พิมพ์ทุกห้อง ({S.rooms.length} ห้อง)
          </button>
        </div>
        {previewTarget.id
          ? <div style={{flex:1,background:"#fff",borderRadius:12,overflow:"auto",padding:14}}>
              {/* header preview */}
              {(layout.showLogo||layout.titleText) && (
                <div style={{marginBottom:8,display:"flex",flexDirection:layout.headerLayout==="logo-left"?"row":"column",alignItems:layout.headerLayout==="logo-left"?"center":"center",gap:8,textAlign:layout.headerLayout==="logo-center"?"center":"left"}}>
                  {layout.showLogo && sh?.logo && <img src={sh.logo} style={{width:layout.logoSize||48,height:layout.logoSize||48,objectFit:"contain"}} alt="logo"/>}
                  <div>
                    <div style={{fontSize:layout.titleFontSize||16,fontWeight:700,fontFamily:layout.fontFamily}}>{layout.titleText || `ตาราง${previewTarget.type==="room"?"เรียน":"สอน"}`}</div>
                    <div style={{fontSize:layout.subtitleFontSize||12,color:"#555",fontFamily:layout.fontFamily}}>{layout.subtitleText || `${sh?.schoolName||"โรงเรียนดาราวิทยาลัย"}`}</div>
                  </div>
                </div>
              )}
              <PreviewTable layout={layout} S={S} targetType={previewTarget.type} targetId={previewTarget.id} />
              {layout.showFooter && (
                <div style={{display:"flex",justifyContent:"space-between",marginTop:8,fontSize:11,fontFamily:layout.fontFamily,padding:"0 4px"}}>
                  <span>{layout.footerLeft}</span><span>{layout.footerRight}</span>
                </div>
              )}
            </div>
          : <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",color:"rgba(255,255,255,0.4)",fontSize:14}}>เลือกห้องหรือครูเพื่อดูตัวอย่าง</div>
        }
      </div>

      {/* ===== RIGHT: Designer Panel ===== */}
      <div style={{width:"min(440px,48vw)",background:"#fff",display:"flex",flexDirection:"column",boxShadow:"-8px 0 40px rgba(0,0,0,0.25)"}}>
        {/* Title bar */}
        <div style={{padding:"14px 18px",background:"#B91C1C",display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
          <div style={{color:"#fff",fontWeight:700,fontSize:15}}>🎨 Print Designer</div>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#fff",fontSize:22,cursor:"pointer",lineHeight:1}}>✕</button>
        </div>

        {/* Saved layouts */}
        <div style={{padding:"10px 14px",borderBottom:"1px solid #F0F0F0",background:"#FAFAFA",flexShrink:0}}>
          <div style={{fontSize:12,fontWeight:600,marginBottom:6,color:"#374151"}}>💾 รูปแบบที่บันทึก</div>
          <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:6}}>
            {Object.keys(layouts).map(n => (
              <button key={n} onClick={()=>{setActiveName(n);const s=loadPDLayouts();setLayout({...PD_DEFAULT_LAYOUT,columns:PD_DEFAULT_COLUMNS.map(c=>({...c})),...s[n]});}}
                style={{padding:"3px 8px",borderRadius:6,fontSize:11,fontWeight:600,border:`2px solid ${activeName===n?"#B91C1C":"#D1D5DB"}`,background:activeName===n?"#FEE2E2":"#fff",color:activeName===n?"#B91C1C":"#374151",cursor:"pointer",display:"flex",alignItems:"center",gap:3}}>
                {n}
                <span onClick={ev=>{ev.stopPropagation();if(window.confirm(`ลบ "${n}"?`))deleteLayout(n);}} style={{color:"#9CA3AF",fontSize:10}}>✕</span>
              </button>
            ))}
            {!Object.keys(layouts).length && <span style={{fontSize:11,color:"#9CA3AF"}}>ยังไม่มีรูปแบบ</span>}
          </div>
          {showSaveAs ? (
            <div style={{display:"flex",gap:6}}>
              <input value={newName} onChange={e=>setNewName(e.target.value)} placeholder="ชื่อ เช่น ม.1-3, ม.4-6" style={{...IS2,flex:1}}
                onKeyDown={e=>{if(e.key==="Enter"&&newName.trim()){saveLayout(newName.trim());setNewName("");setShowSaveAs(false);}}} />
              <button onClick={()=>{if(newName.trim()){saveLayout(newName.trim());setNewName("");setShowSaveAs(false);}}} style={{padding:"5px 12px",background:"#B91C1C",color:"#fff",border:"none",borderRadius:7,fontSize:12,fontWeight:600,cursor:"pointer"}}>บันทึก</button>
              <button onClick={()=>setShowSaveAs(false)} style={{padding:"5px 8px",border:"1px solid #E5E7EB",borderRadius:7,background:"none",cursor:"pointer",fontSize:12}}>ยกเลิก</button>
            </div>
          ) : (
            <div style={{display:"flex",gap:6}}>
              {activeName && <button onClick={()=>saveLayout(activeName)} style={{flex:1,padding:"5px",background:"#1D4ED8",color:"#fff",border:"none",borderRadius:7,fontSize:11,fontWeight:600,cursor:"pointer"}}>💾 ทับ "{activeName}"</button>}
              <button onClick={()=>{setShowSaveAs(true);setNewName("");}} style={{flex:1,padding:"5px",background:"#059669",color:"#fff",border:"none",borderRadius:7,fontSize:11,fontWeight:600,cursor:"pointer"}}>+ บันทึกใหม่</button>
            </div>
          )}
        </div>

        {/* Tabs */}
        <div style={{display:"flex",borderBottom:"2px solid #F0F0F0",flexShrink:0}}>
          {TABS.map(([v,l]) => (
            <button key={v} onClick={()=>setTab(v)} style={{flex:1,padding:"8px 4px",fontSize:11,fontWeight:tab===v?700:400,border:"none",borderBottom:tab===v?"2px solid #B91C1C":"2px solid transparent",background:"none",color:tab===v?"#B91C1C":"#6B7280",cursor:"pointer",marginBottom:-2}}>
              {l}
            </button>
          ))}
        </div>

        {/* Tab content */}
        <div style={{flex:1,overflowY:"auto",padding:"14px 16px"}}>

          {/* ===== TAB: CELL ===== */}
          {tab==="cell" && (
            <div>
              <div style={{fontSize:12,fontWeight:700,color:"#1F2937",marginBottom:10}}>เนื้อหาในแต่ละช่อง (แต่ละบรรทัด)</div>
              <div style={{display:"flex",flexDirection:"column",gap:8}}>
                {layout.cellRows.map((row, i) => (
                  <div key={i} style={{background:"#F9FAFB",borderRadius:10,padding:"10px 12px",border:"1.5px solid #E5E7EB"}}>
                    <div style={{display:"flex",gap:5,alignItems:"center",marginBottom:6}}>
                      <span style={{fontSize:11,color:"#9CA3AF",fontWeight:600,minWidth:18}}>#{i+1}</span>
                      <select value={row.field} onChange={e=>updCellRow(i,"field",e.target.value)} style={{...IS2,flex:1}}>
                        {PD_FIELD_OPTIONS.map(opt=><option key={opt.value} value={opt.value}>{opt.label}</option>)}
                      </select>
                      <button onClick={()=>moveCellRow(i,-1)} disabled={i===0} style={{padding:"3px 6px",borderRadius:5,border:"1px solid #D1D5DB",background:"#fff",cursor:"pointer",opacity:i===0?0.3:1,fontSize:12}}>↑</button>
                      <button onClick={()=>moveCellRow(i,1)} disabled={i===layout.cellRows.length-1} style={{padding:"3px 6px",borderRadius:5,border:"1px solid #D1D5DB",background:"#fff",cursor:"pointer",opacity:i===layout.cellRows.length-1?0.3:1,fontSize:12}}>↓</button>
                      <button onClick={()=>removeCellRow(i)} style={{padding:"3px 6px",borderRadius:5,border:"1px solid #FCA5A5",background:"#FEF2F2",color:"#B91C1C",cursor:"pointer",fontSize:12}}>✕</button>
                    </div>
                    {row.field==="custom_text" && (
                      <input value={row.customText||""} onChange={e=>updCellRow(i,"customText",e.target.value)} placeholder="ข้อความ..." style={{...IS2,marginBottom:6}}/>
                    )}
                    <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                      <label style={{fontSize:11,display:"flex",alignItems:"center",gap:4}}>
                        ขนาด: <input type="number" min={7} max={22} value={row.fontSize} onChange={e=>updCellRow(i,"fontSize",parseInt(e.target.value)||11)} style={{width:44,padding:"2px 5px",border:"1px solid #D1D5DB",borderRadius:5,fontSize:11}}/>
                      </label>
                      <label style={{fontSize:11,display:"flex",alignItems:"center",gap:4,cursor:"pointer"}}>
                        <input type="checkbox" checked={row.bold} onChange={e=>updCellRow(i,"bold",e.target.checked)} style={{accentColor:"#B91C1C"}}/>ตัวหนา
                      </label>
                      <select value={row.align} onChange={e=>updCellRow(i,"align",e.target.value)} style={{padding:"2px 6px",border:"1px solid #D1D5DB",borderRadius:5,fontSize:11}}>
                        <option value="left">ซ้าย</option><option value="center">กลาง</option><option value="right">ขวา</option>
                      </select>
                      <label style={{fontSize:11,display:"flex",alignItems:"center",gap:4}}>
                        สี:<input type="color" value={row.color} onChange={e=>updCellRow(i,"color",e.target.value)} style={{width:26,height:22,border:"none",cursor:"pointer"}}/>
                      </label>
                    </div>
                  </div>
                ))}
                <button onClick={addCellRow} style={{padding:"8px",borderRadius:10,border:"2px dashed #D1D5DB",background:"none",color:"#6B7280",fontSize:12,cursor:"pointer",fontWeight:600}}>+ เพิ่มบรรทัด</button>
              </div>
            </div>
          )}

          {/* ===== TAB: COLUMNS ===== */}
          {tab==="columns" && (
            <div>
              <div style={{fontSize:12,fontWeight:700,color:"#1F2937",marginBottom:4}}>คอลัมน์ในตาราง (ลากปรับลำดับ / เปิด-ปิด)</div>
              <div style={{fontSize:11,color:"#9CA3AF",marginBottom:10}}>เพิ่มคอลัมน์พัก, Homeroom, Assembly ได้ตามต้องการ</div>
              <div style={{display:"flex",gap:6,marginBottom:12}}>
                <button onClick={addBreakCol} style={{flex:1,padding:"7px",background:"#F3F4F6",border:"1.5px dashed #9CA3AF",borderRadius:8,fontSize:11,fontWeight:600,cursor:"pointer",color:"#4B5563"}}>+ พักระหว่างคาบ</button>
                <button onClick={addCustomCol} style={{flex:1,padding:"7px",background:"#EFF6FF",border:"1.5px dashed #93C5FD",borderRadius:8,fontSize:11,fontWeight:600,cursor:"pointer",color:"#1D4ED8"}}>+ คอลัมน์กำหนดเอง</button>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {layout.columns.map((col, i) => (
                  <div key={col.id} style={{background: col.show?"#F9FAFB":"#F9FAFB",borderRadius:8,padding:"8px 10px",border:`1.5px solid ${col.show?"#E5E7EB":"#F3F4F6"}`,opacity:col.show?1:0.5}}>
                    <div style={{display:"flex",gap:6,alignItems:"center"}}>
                      {/* toggle */}
                      <button onClick={()=>updCol(col.id,"show",!col.show)} style={{padding:"3px 8px",borderRadius:6,border:"none",background:col.show?"#D1FAE5":"#F3F4F6",color:col.show?"#065F46":"#9CA3AF",fontSize:11,fontWeight:700,cursor:"pointer",minWidth:36}}>{col.show?"✓":"−"}</button>
                      {/* type badge */}
                      <span style={{fontSize:10,padding:"2px 6px",borderRadius:4,background:col.type==="period"?"#DBEAFE":col.type==="break"?"#F3F4F6":col.type==="homeroom"||col.type==="assembly"?"#FEF9C3":"#F0FDF4",color:col.type==="period"?"#1D4ED8":col.type==="break"?"#6B7280":"#92400E",fontWeight:600}}>
                        {col.type==="day"?"วัน":col.type==="period"?"คาบ":col.type==="break"?"พัก":col.type==="homeroom"?"HR":col.type==="assembly"?"Asm":"กำหนด"}
                      </span>
                      {/* label edit */}
                      <input value={col.label||""} onChange={e=>updCol(col.id,"label",e.target.value)} style={{...IS2,flex:1,padding:"3px 7px",fontSize:11}}/>
                      {/* width */}
                      <input type="number" min={15} max={250} value={col.width} onChange={e=>updCol(col.id,"width",parseInt(e.target.value)||60)} style={{width:50,padding:"2px 5px",border:"1px solid #D1D5DB",borderRadius:5,fontSize:11}}/>
                      <span style={{fontSize:10,color:"#9CA3AF"}}>px</span>
                      {/* move */}
                      <button onClick={()=>moveCol(col.id,-1)} disabled={i===0} style={{padding:"2px 5px",borderRadius:5,border:"1px solid #D1D5DB",background:"#fff",cursor:"pointer",opacity:i===0?0.3:1,fontSize:11}}>↑</button>
                      <button onClick={()=>moveCol(col.id,1)} disabled={i===layout.columns.length-1} style={{padding:"2px 5px",borderRadius:5,border:"1px solid #D1D5DB",background:"#fff",cursor:"pointer",opacity:i===layout.columns.length-1?0.3:1,fontSize:11}}>↓</button>
                      {/* delete (ไม่ลบ day/period แกนหลัก) */}
                      {(col.type==="break"||col.type==="custom"||col.type==="homeroom"||col.type==="assembly") && (
                        <button onClick={()=>removeCol(col.id)} style={{padding:"2px 5px",borderRadius:5,border:"1px solid #FCA5A5",background:"#FEF2F2",color:"#B91C1C",cursor:"pointer",fontSize:11}}>✕</button>
                      )}
                    </div>
                    {/* extra options for break/homeroom/assembly */}
                    {(col.type==="break"||col.type==="homeroom"||col.type==="assembly"||col.type==="custom") && (
                      <div style={{display:"flex",gap:8,marginTop:6,flexWrap:"wrap",alignItems:"center"}}>
                        {(col.type==="break") && (
                          <label style={{fontSize:11,display:"flex",alignItems:"center",gap:4}}>
                            เวลา: <input value={col.timeLabel||""} onChange={e=>updCol(col.id,"timeLabel",e.target.value)} placeholder="10.10-10.25" style={{padding:"2px 6px",border:"1px solid #D1D5DB",borderRadius:5,fontSize:11,width:90}}/>
                          </label>
                        )}
                        <label style={{fontSize:11,display:"flex",alignItems:"center",gap:4}}>
                          สีพื้น:<input type="color" value={col.bg||"#F3F4F6"} onChange={e=>updCol(col.id,"bg",e.target.value)} style={{width:24,height:20,border:"none",cursor:"pointer"}}/>
                        </label>
                        <label style={{fontSize:11,display:"flex",alignItems:"center",gap:4}}>
                          สีตัวอักษร:<input type="color" value={col.textColor||"#6B7280"} onChange={e=>updCol(col.id,"textColor",e.target.value)} style={{width:24,height:20,border:"none",cursor:"pointer"}}/>
                        </label>
                      </div>
                    )}
                    {col.type==="period" && (
                      <div style={{display:"flex",gap:8,marginTop:5,flexWrap:"wrap",alignItems:"center"}}>
                        <label style={{fontSize:11,display:"flex",alignItems:"center",gap:4}}>
                          คาบ ID: <input type="number" min={1} max={20} value={col.periodId||""} onChange={e=>updCol(col.id,"periodId",parseInt(e.target.value))} style={{width:44,padding:"2px 5px",border:"1px solid #D1D5DB",borderRadius:5,fontSize:11}}/>
                        </label>
                        <label style={{fontSize:11,display:"flex",alignItems:"center",gap:4}}>
                          เวลา: <input value={col.timeLabel||""} onChange={e=>updCol(col.id,"timeLabel",e.target.value)} placeholder="08.30-09.20" style={{padding:"2px 6px",border:"1px solid #D1D5DB",borderRadius:5,fontSize:11,width:100}}/>
                        </label>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ===== TAB: STYLE ===== */}
          {tab==="style" && (
            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              {/* Colors */}
              <div>
                <div style={{fontSize:12,fontWeight:700,color:"#1F2937",marginBottom:8}}>🎨 สี</div>
                <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:8}}>
                  <label style={{flex:1,fontSize:12,minWidth:100}}>สีหัวตาราง
                    <div style={{display:"flex",gap:5,marginTop:4}}>
                      <input type="color" value={layout.headerBg} onChange={e=>upd("headerBg",e.target.value)} style={{width:32,height:26,border:"none",cursor:"pointer"}}/>
                      <input type="color" value={layout.headerText} onChange={e=>upd("headerText",e.target.value)} style={{width:32,height:26,border:"none",cursor:"pointer"}}/>
                      <span style={{fontSize:10,color:"#9CA3AF",alignSelf:"center"}}>พื้น / ตัวอักษร</span>
                    </div>
                  </label>
                  <label style={{flex:1,fontSize:12,minWidth:100}}>สีแถวคี่
                    <div style={{marginTop:4}}><input type="color" value={layout.rowAltBg} onChange={e=>upd("rowAltBg",e.target.value)} style={{width:32,height:26,border:"none",cursor:"pointer"}}/></div>
                  </label>
                  <label style={{flex:1,fontSize:12,minWidth:100}}>สีเส้นขอบ
                    <div style={{marginTop:4}}><input type="color" value={layout.borderColor||"#E5E7EB"} onChange={e=>upd("borderColor",e.target.value)} style={{width:32,height:26,border:"none",cursor:"pointer"}}/></div>
                  </label>
                </div>
                <div style={{display:"flex",gap:14,flexWrap:"wrap"}}>
                  {[["showAltRow","สลับสีแถว"],["showBorder","เส้นขอบ"],["showPeriodNum","แสดงเลขคาบ"],["showPeriodTime","แสดงเวลาในหัว"]].map(([k,l])=>(
                    <label key={k} style={{display:"flex",alignItems:"center",gap:5,fontSize:12,cursor:"pointer"}}>
                      <input type="checkbox" checked={!!layout[k]} onChange={e=>upd(k,e.target.checked)} style={{accentColor:"#B91C1C"}}/>{l}
                    </label>
                  ))}
                </div>
              </div>
              {/* Font */}
              <div>
                <div style={{fontSize:12,fontWeight:700,color:"#1F2937",marginBottom:6}}>🔤 ฟอนต์และขนาด</div>
                <select value={layout.fontFamily} onChange={e=>upd("fontFamily",e.target.value)} style={{...IS2,marginBottom:8}}>
                  {["Sarabun","TH SarabunNew","Arial","Tahoma","Kanit","Prompt"].map(f=><option key={f} value={f}>{f}</option>)}
                </select>
                <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:6}}>
                  <span style={{fontSize:11,minWidth:90}}>ขนาดตัวอักษร {layout.fontSize}%</span>
                  <input type="range" min={60} max={150} value={layout.fontSize} onChange={e=>upd("fontSize",+e.target.value)} style={{flex:1,accentColor:"#B91C1C"}}/>
                </div>
                <div style={{display:"flex",gap:8,alignItems:"center"}}>
                  <span style={{fontSize:11,minWidth:90}}>ความสูงแถว {layout.rowHeight}%</span>
                  <input type="range" min={60} max={180} value={layout.rowHeight} onChange={e=>upd("rowHeight",+e.target.value)} style={{flex:1,accentColor:"#B91C1C"}}/>
                </div>
              </div>
              {/* Paper */}
              <div>
                <div style={{fontSize:12,fontWeight:700,color:"#1F2937",marginBottom:6}}>📄 กระดาษ</div>
                <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:6}}>
                  {[["A4","A4"],["A3","A3"],["A5","A5"]].map(([v,l])=>(
                    <button key={v} onClick={()=>upd("paperSize",v)} style={{flex:1,padding:"6px",border:`2px solid ${layout.paperSize===v?"#B91C1C":"#E5E7EB"}`,borderRadius:7,background:layout.paperSize===v?"#FEE2E2":"#fff",color:layout.paperSize===v?"#B91C1C":"#374151",fontWeight:600,cursor:"pointer",fontSize:12}}>{l}</button>
                  ))}
                  {[["landscape","แนวนอน"],["portrait","แนวตั้ง"]].map(([v,l])=>(
                    <button key={v} onClick={()=>upd("orientation",v)} style={{flex:1,padding:"6px",border:`2px solid ${layout.orientation===v?"#B91C1C":"#E5E7EB"}`,borderRadius:7,background:layout.orientation===v?"#FEE2E2":"#fff",color:layout.orientation===v?"#B91C1C":"#374151",fontWeight:600,cursor:"pointer",fontSize:12}}>{l}</button>
                  ))}
                </div>
                <div style={{display:"flex",gap:8,alignItems:"center"}}>
                  <span style={{fontSize:11,minWidth:80}}>ขอบ {layout.marginMm||10}mm</span>
                  <input type="range" min={3} max={20} value={layout.marginMm||10} onChange={e=>upd("marginMm",+e.target.value)} style={{flex:1,accentColor:"#B91C1C"}}/>
                </div>
              </div>
              {/* Reset */}
              <button onClick={()=>upd("columns",PD_DEFAULT_COLUMNS.map(c=>({...c})))||setLayout({...PD_DEFAULT_LAYOUT,columns:PD_DEFAULT_COLUMNS.map(c=>({...c}))})} style={{padding:"9px",border:"2px solid #E5E7EB",borderRadius:8,background:"#F9FAFB",color:"#6B7280",fontWeight:600,cursor:"pointer",fontSize:12}}>↩ รีเซ็ตทั้งหมด</button>
            </div>
          )}

          {/* ===== TAB: HEADER ===== */}
          {tab==="header" && (
            <div style={{display:"flex",flexDirection:"column",gap:12}}>
              {/* Logo */}
              <div>
                <div style={{fontSize:12,fontWeight:700,color:"#1F2937",marginBottom:6}}>🖼 โลโก้</div>
                <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:6}}>
                  {sh?.logo
                    ? <img src={sh.logo} style={{width:40,height:40,objectFit:"contain",border:"1px solid #E5E7EB",borderRadius:6}} alt="logo"/>
                    : <div style={{width:40,height:40,background:"#F3F4F6",borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,color:"#9CA3AF",textAlign:"center"}}>ไม่มี<br/>โลโก้</div>
                  }
                  <div style={{fontSize:11,color:"#6B7280",flex:1}}>โลโก้มาจากหน้า ตั้งค่า → โลโก้โรงเรียน<br/>{sh?.logo?"✅ มีโลโก้แล้ว":"❌ ยังไม่มีโลโก้"}</div>
                </div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                  <label style={{display:"flex",alignItems:"center",gap:5,cursor:"pointer",fontSize:12}}>
                    <input type="checkbox" checked={!!layout.showLogo} onChange={e=>upd("showLogo",e.target.checked)} style={{accentColor:"#B91C1C"}}/>แสดงโลโก้
                  </label>
                  <label style={{fontSize:11,display:"flex",alignItems:"center",gap:4}}>
                    ขนาด: <input type="number" min={24} max={80} value={layout.logoSize||48} onChange={e=>upd("logoSize",+e.target.value)} style={{width:48,padding:"2px 5px",border:"1px solid #D1D5DB",borderRadius:5,fontSize:11}}/>px
                  </label>
                </div>
              </div>
              {/* Header layout */}
              <div>
                <label style={{...LS}}>รูปแบบหัว</label>
                <div style={{display:"flex",gap:6}}>
                  {[["logo-left","โลโก้ซ้าย"],["logo-center","โลโก้กลาง"],["no-logo","ไม่มีโลโก้"]].map(([v,l])=>(
                    <button key={v} onClick={()=>upd("headerLayout",v)} style={{flex:1,padding:"6px",border:`2px solid ${layout.headerLayout===v?"#B91C1C":"#E5E7EB"}`,borderRadius:7,background:layout.headerLayout===v?"#FEE2E2":"#fff",color:layout.headerLayout===v?"#B91C1C":"#374151",fontWeight:600,cursor:"pointer",fontSize:11}}>{l}</button>
                  ))}
                </div>
              </div>
              {/* Title */}
              <div>
                <label style={{...LS}}>ชื่อตาราง (ว่าง = ชื่อห้อง/ครูอัตโนมัติ)</label>
                <input value={layout.titleText||""} onChange={e=>upd("titleText",e.target.value)} placeholder="ตารางเรียน ม.4/3" style={IS2}/>
                <div style={{display:"flex",gap:8,alignItems:"center",marginTop:6}}>
                  <span style={{fontSize:11,minWidth:60}}>ขนาด {layout.titleFontSize||16}px</span>
                  <input type="range" min={10} max={28} value={layout.titleFontSize||16} onChange={e=>upd("titleFontSize",+e.target.value)} style={{flex:1,accentColor:"#B91C1C"}}/>
                </div>
              </div>
              {/* Subtitle */}
              <div>
                <label style={{...LS}}>คำบรรยาย (ว่าง = โรงเรียน/ปีการศึกษาอัตโนมัติ)</label>
                <input value={layout.subtitleText||""} onChange={e=>upd("subtitleText",e.target.value)} placeholder="ภาคเรียนที่ 1/2569 โรงเรียนดาราวิทยาลัย" style={IS2}/>
                <div style={{display:"flex",gap:8,alignItems:"center",marginTop:6}}>
                  <span style={{fontSize:11,minWidth:60}}>ขนาด {layout.subtitleFontSize||12}px</span>
                  <input type="range" min={9} max={20} value={layout.subtitleFontSize||12} onChange={e=>upd("subtitleFontSize",+e.target.value)} style={{flex:1,accentColor:"#B91C1C"}}/>
                </div>
              </div>
              {/* Footer */}
              <div>
                <label style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer",fontSize:12,marginBottom:8}}>
                  <input type="checkbox" checked={!!layout.showFooter} onChange={e=>upd("showFooter",e.target.checked)} style={{accentColor:"#B91C1C"}}/>แสดงส่วนลงชื่อท้ายตาราง
                </label>
                {layout.showFooter && <>
                  <label style={{...LS}}>ซ้าย</label>
                  <input value={layout.footerLeft||""} onChange={e=>upd("footerLeft",e.target.value)} style={{...IS2,marginBottom:8}}/>
                  <label style={{...LS}}>ขวา</label>
                  <input value={layout.footerRight||""} onChange={e=>upd("footerRight",e.target.value)} style={IS2}/>
                </>}
              </div>
            </div>
          )}

        </div>{/* end tab content */}
      </div>{/* end right panel */}
    </div>
  );
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
    const code=prompt("🔐 การลบข้อมูลทั้งหมดต้องใช้รหัสผ่าน\n\n⚠️ ถ้าไม่ทราบรหัส ให้ถาม อ.พนิต เกิดมงคล");
    if(code===null)return;
    if(code!=="100625"){alert("❌ รหัสไม่ถูกต้อง\n\nหากต้องการดำเนินการ กรุณาติดต่อ อ.พนิต เกิดมงคล");return;}
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
    const code=prompt("🔐 การลบตารางสอนต้องใช้รหัสผ่าน\n\n⚠️ ถ้าไม่ทราบรหัส ให้ถาม อ.พนิต เกิดมงคล");
    if(code===null)return;
    if(code!=="100625"){alert("❌ รหัสไม่ถูกต้อง\n\nหากต้องการดำเนินการ กรุณาติดต่อ อ.พนิต เกิดมงคล");return;}
    if(!confirm("ลบเฉพาะข้อมูลตารางสอน (ข้อมูลครู/วิชา/ห้องยังอยู่)?"))return;
    U.setSchedule({});U.setLocks({});
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
const loadPrintSettings=()=>{try{const s=localStorage.getItem("dara_printSettings");return s?{...DEFAULT_PRINT_SETTINGS,...JSON.parse(s)}:DEFAULT_PRINT_SETTINGS;}catch{return DEFAULT_PRINT_SETTINGS;}};
const savePrintSettings=(s)=>{try{localStorage.setItem("dara_printSettings",JSON.stringify(s));}catch{}};
function PrintSettingsPanel({open,onClose,onApply}){
  const [s,setS]=useState(loadPrintSettings);
  if(!open)return null;
  const u=(k,v)=>setS(p=>({...p,[k]:v}));
  return(
    <div style={{position:"fixed",inset:0,zIndex:3000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.5)"}}>
      <div style={{background:"#fff",borderRadius:16,padding:28,width:"min(540px,95vw)",maxHeight:"90vh",overflowY:"auto",boxShadow:"0 20px 60px rgba(0,0,0,0.3)",fontFamily:"'Sarabun',sans-serif"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:18}}>
          <h2 style={{fontSize:18,fontWeight:700}}>⚙️ ตั้งค่าการพิมพ์</h2>
          <button onClick={onClose} style={{background:"none",border:"none",fontSize:20,cursor:"pointer",color:"#6B7280"}}>✕</button>
        </div>
        <div style={{marginBottom:16}}>
          <label style={{fontSize:13,fontWeight:600,display:"block",marginBottom:8}}>🔤 ฟอนต์</label>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {PRINT_FONTS.map(f=><button key={f} onClick={()=>u("fontFamily",f)} style={{padding:"6px 14px",borderRadius:8,border:"2px solid "+(s.fontFamily===f?"#B91C1C":"#D1D5DB"),background:s.fontFamily===f?"#FEE2E2":"#fff",fontFamily:f,fontSize:13,cursor:"pointer",fontWeight:s.fontFamily===f?700:400}}>{f}</button>)}
          </div>
        </div>
        <div style={{marginBottom:16}}>
          <label style={{fontSize:13,fontWeight:600,display:"block",marginBottom:8}}>📏 ขนาดตัวอักษร: <b style={{color:"#B91C1C"}}>{s.fontSize}%</b></label>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <button onClick={()=>u("fontSize",Math.max(60,s.fontSize-10))} style={{width:32,height:32,borderRadius:8,border:"1px solid #D1D5DB",background:"#F9FAFB",fontSize:16,cursor:"pointer"}}>−</button>
            <input type="range" min={60} max={160} value={s.fontSize} onChange={e=>u("fontSize",parseInt(e.target.value))} style={{flex:1,accentColor:"#B91C1C"}}/>
            <button onClick={()=>u("fontSize",Math.min(160,s.fontSize+10))} style={{width:32,height:32,borderRadius:8,border:"1px solid #D1D5DB",background:"#F9FAFB",fontSize:16,cursor:"pointer"}}>+</button>
            <button onClick={()=>u("fontSize",100)} style={{fontSize:11,color:"#6B7280",background:"none",border:"1px solid #D1D5DB",borderRadius:6,padding:"3px 8px",cursor:"pointer"}}>รีเซ็ต</button>
          </div>
        </div>
        <div style={{marginBottom:16}}>
          <label style={{fontSize:13,fontWeight:600,display:"block",marginBottom:8}}>↕️ ความสูงแถว: <b style={{color:"#B91C1C"}}>{s.rowHeight}%</b></label>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <button onClick={()=>u("rowHeight",Math.max(60,s.rowHeight-10))} style={{width:32,height:32,borderRadius:8,border:"1px solid #D1D5DB",background:"#F9FAFB",fontSize:16,cursor:"pointer"}}>−</button>
            <input type="range" min={60} max={160} value={s.rowHeight} onChange={e=>u("rowHeight",parseInt(e.target.value))} style={{flex:1,accentColor:"#B91C1C"}}/>
            <button onClick={()=>u("rowHeight",Math.min(160,s.rowHeight+10))} style={{width:32,height:32,borderRadius:8,border:"1px solid #D1D5DB",background:"#F9FAFB",fontSize:16,cursor:"pointer"}}>+</button>
            <button onClick={()=>u("rowHeight",100)} style={{fontSize:11,color:"#6B7280",background:"none",border:"1px solid #D1D5DB",borderRadius:6,padding:"3px 8px",cursor:"pointer"}}>รีเซ็ต</button>
          </div>
        </div>
        <div style={{marginBottom:16}}>
          <label style={{fontSize:13,fontWeight:600,display:"block",marginBottom:8}}>🎨 สีหัวตาราง</label>
          <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
            {Object.entries(PRINT_COLORS).map(([name,c])=>(
              <button key={name} onClick={()=>u("color",name)} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:4,padding:"8px 12px",borderRadius:10,border:"2px solid "+(s.color===name?"#B91C1C":"#E5E7EB"),background:s.color===name?"#FEF2F2":"#fff",cursor:"pointer"}}>
                <div style={{width:36,height:20,borderRadius:5,background:c.header,display:"flex",alignItems:"center",justifyContent:"center"}}><span style={{color:c.headerText,fontSize:9,fontWeight:700}}>วัน</span></div>
                <div style={{width:36,height:10,borderRadius:4,background:c.rowAlt,border:"1px solid #eee"}}/>
                <span style={{fontSize:11,fontWeight:s.color===name?700:400,color:s.color===name?"#B91C1C":"#374151"}}>{name}</span>
              </button>
            ))}
          </div>
        </div>
        <div style={{marginBottom:18,display:"flex",gap:20}}>
          <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",fontSize:13}}>
            <input type="checkbox" checked={s.showAltRow} onChange={e=>u("showAltRow",e.target.checked)} style={{width:16,height:16,accentColor:"#B91C1C"}}/>สลับสีแถว
          </label>
          <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",fontSize:13}}>
            <input type="checkbox" checked={s.showBorder} onChange={e=>u("showBorder",e.target.checked)} style={{width:16,height:16,accentColor:"#B91C1C"}}/>เส้นขอบ
          </label>
        </div>
        <div style={{background:"#F9FAFB",borderRadius:10,padding:10,marginBottom:18}}>
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
          <button onClick={()=>setS(DEFAULT_PRINT_SETTINGS)} style={{...BO(),flex:1}}>↩ ค่าเริ่มต้น</button>
          <button onClick={()=>{savePrintSettings(s);onApply(s);onClose();}} style={{...BS(),flex:2}}>💾 บันทึก &amp; ใช้งาน</button>
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
