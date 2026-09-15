import React,{useState,useRef,useEffect} from 'react';
export default function TeacherPicker({teachers,value,onChange}){
 const [open,setOpen]=useState(false),[query,setQuery]=useState('');const input=useRef(null),trigger=useRef(null);
 useEffect(()=>{if(open)input.current?.focus();},[open]);
 const name=t=>`${t.prefix||''}${t.firstName||''} ${t.lastName||''}`.trim(),chosen=teachers.find(t=>t.id===value);
 const list=teachers.filter(t=>(name(t)+' '+(t.code||'')).toLocaleLowerCase().includes(query.trim().toLocaleLowerCase())).sort((a,b)=>name(a).localeCompare(name(b),'th'));
 return <div className="teacher-picker"><button ref={trigger} type="button" aria-expanded={open} onClick={()=>{setOpen(!open);setQuery('')}}>{chosen?name(chosen):'ค้นหาและเลือกครู'} <span>⌕</span></button>{open&&<div className="teacher-options"><input ref={input} aria-label="พิมพ์ค้นหาครูผู้ขอแลก" placeholder="พิมพ์ชื่อ นามสกุล หรือรหัสครู…" value={query} onChange={e=>setQuery(e.target.value)} onKeyDown={e=>{if(e.key==='Escape'){setOpen(false);trigger.current?.focus()}}}/><small>พบ {list.length} คน</small><div className="teacher-results">{list.map(t=><button type="button" key={t.id} aria-pressed={value===t.id} onClick={()=>{onChange(t.id);setOpen(false);trigger.current?.focus()}}>{name(t)}{value===t.id?' ✓':''}</button>)}{!list.length&&<p>ไม่พบชื่อครู ลองพิมพ์ชื่อบางส่วน</p>}</div><button type="button" onClick={()=>{setOpen(false);trigger.current?.focus()}}>ปิดรายชื่อ</button></div>}</div>;
}
