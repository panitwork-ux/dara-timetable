import React from 'react';
export default function DateField({value,onChange,label,min}){
 const formatted=value?value.split('-').reverse().map((part,i)=>i===2?String(Number(part)+543):part).join('/'):'วัน / เดือน / ปี พ.ศ.';
 return <div className="date-field"><input type="date" aria-label={label} lang="th-TH" value={value} min={min} onChange={e=>onChange(e.target.value)}/><small className="date-caption">{formatted}</small></div>;
}
