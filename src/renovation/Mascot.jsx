import React, {useEffect,useState} from 'react';
const notes=['ค่อย ๆ จัดทีละคาบก็ได้นะ วันนี้เราทำได้อีกนิดแล้ว','พักสายตาสักครู่ แล้วค่อยกลับมาจัดต่อด้วยกันนะ','ตารางที่ลงตัว เริ่มจากทีละช่องเล็ก ๆ','ไหล่ผ่อนคลาย จิบน้ำสักนิด แล้วไปต่อกัน'];
export default function Mascot({compact=false}) {
 const [paused,setPaused]=useState(()=>{try{return localStorage.getItem('dara_mascot_paused')==='true'}catch{return false}});
 const [note,setNote]=useState(0),[hello,setHello]=useState(false);
 useEffect(()=>{try{localStorage.setItem('dara_mascot_paused',String(paused))}catch{};window.dispatchEvent(new CustomEvent('dara-mascot-motion',{detail:paused}));},[paused]);
 useEffect(()=>{const update=e=>setPaused(e.detail);window.addEventListener('dara-mascot-motion',update);return()=>window.removeEventListener('dara-mascot-motion',update)},[]);
 useEffect(()=>{if(!hello)return;const t=setTimeout(()=>setHello(false),900);return()=>clearTimeout(t)},[hello]);
 return <section className={'dara-companion '+(compact?'companion-compact':'companion-banner')+(paused?' motion-paused':'')} aria-label="น้องดารา เพื่อนพักสายตา">
  <div className="companion-copy"><span className="companion-label">เพื่อนตัวเล็กของวันทำงาน</span><h2>{compact?'น้องดาราอยู่เป็นเพื่อนนะ':'สวัสดี วันนี้มาจัดตารางไปด้วยกันนะ'}</h2><p aria-live="polite">{notes[note]}</p><div className="companion-controls"><button onClick={()=>{setNote((note+1)%notes.length);setHello(true)}}>ส่งกำลังใจให้หน่อย ✧</button><button aria-pressed={paused} onClick={()=>setPaused(!paused)}>{paused?'เล่นการเคลื่อนไหว':'พักการเคลื่อนไหว'}</button></div></div>
  <button className={'mascot-stage '+(hello?'say-hello':'')} aria-label="ทักทายน้องดารา" onClick={()=>{setNote((note+1)%notes.length);setHello(true)}}><span className="mascot-orbit orbit-one"/><span className="mascot-orbit orbit-two"/><img src={(import.meta.env?.BASE_URL||'./')+'mascot-dara.png'} alt="น้องดารา ลูกนกฮูกสีครีมผูกผ้าพันคอแดง ถือสมุดตารางสอน" width="180" height="180"/><span className="mascot-shadow"/></button>
 </section>;
}
