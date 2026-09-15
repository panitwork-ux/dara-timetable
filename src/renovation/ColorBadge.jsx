import React from 'react';
import {departmentTone,levelTone} from './colors.mjs';
export default function ColorBadge({item,kind='department'}){const t=kind==='level'?levelTone(item):departmentTone(item);return <span className="entity-badge" style={{color:t.ink,background:t.bg,borderColor:t.ink+'30'}}><i style={{background:t.ink}}/>{item?.name||'ยังไม่ระบุ'}</span>}
