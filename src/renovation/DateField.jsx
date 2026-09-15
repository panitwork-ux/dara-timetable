import React,{useState,useEffect} from 'react';
export default function DateField({value,onChange,label}){
 const parts=value?value.split('-'):[];
 const [day,D]=useState(parts[2]||''),[month,M]=useState(parts[1]||''),[year,Y]=useState(parts[0]?String(+parts[0]+543):String(new Date().getFullYear()+543));
 useEffect(()=>{if(value){const [y,m,d]=value.split('-');D(d);M(m);Y(String(+y+543));}},[value]);
 const update=(d,m,y)=>{D(d);M(m);Y(y);const n=+y-543,iso=String(n).padStart(4,'0')+'-'+m+'-'+d,date=new Date(iso+'T12:00:00Z');onChange(d&&m&&y.length===4&&n>=1900&&n<=2200&&!isNaN(date)&&date.toISOString().slice(0,10)===iso?iso:'');};
 const invalid=day&&month&&year.length===4&&!value;
 return <div className="date-field"><div className="date-parts" role="group" aria-label={label}><select aria-label={label+' วัน'} value={day} onChange={e=>update(e.target.value,month,year)}><option value="">วัน</option>{Array.from({length:31},(_,i)=>String(i+1).padStart(2,'0')).map(d=><option key={d}>{d}</option>)}</select><span>/</span><select aria-label={label+' เดือน'} value={month} onChange={e=>update(day,e.target.value,year)}><option value="">เดือน</option>{Array.from({length:12},(_,i)=>String(i+1).padStart(2,'0')).map(m=><option key={m}>{m}</option>)}</select><span>/</span><input aria-label={label+' ปี พ.ศ.'} inputMode="numeric" maxLength={4} value={year} onChange={e=>update(day,month,e.target.value.replace(/\D/g,''))}/></div><small>วัน / เดือน / ปี พ.ศ.</small>{invalid&&<small role="alert">วันที่ไม่มีอยู่จริง กรุณาตรวจวันและเดือน</small>}</div>;
}
