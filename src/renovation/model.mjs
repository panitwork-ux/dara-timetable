export const DAYS = ['จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์'];
export const teacherLoad = (S, id) => new Set(Object.entries(S.schedule || {}).flatMap(([key, entries]) => entries.filter(e => e.teacherId === id || (e.coTeacherIds || []).includes(id) || e.coTeacherId === id).map(() => key.split('_').slice(1).join('_')))).size;
export function progress(S) {
  const placedByAssignment = {};
  Object.values(S.schedule || {}).flat().forEach(e => { if(e.assignmentId) placedByAssignment[e.assignmentId]=(placedByAssignment[e.assignmentId]||0)+1; });
  const total = S.assigns.reduce((n,a)=>n+(Number(a.totalPeriods)||0),0);
  const placed = S.assigns.reduce((n,a)=>n+Math.min(Number(a.totalPeriods)||0,placedByAssignment[a.id]||0),0);
  return { total, placed, remaining:Math.max(0,total-placed), percent:total?Math.round(placed/total*100):0 };
}
export function seedPreview() {
  if (import.meta.env?.VITE_LIVE_FIREBASE !== 'false') return;
  const prefix='dara_preview_';
  if(localStorage.getItem(prefix+'seed_v1')) return;
  const levels=[{id:'lv4',name:'ม.4',divisionId:'m2'},{id:'lv5',name:'ม.5',divisionId:'m2'},{id:'lv6',name:'ม.6',divisionId:'m2'}];
  const depts=['คณิตศาสตร์','วิทยาศาสตร์','ภาษาต่างประเทศ','ภาษาไทย','สังคมศึกษา','ศิลปะ'].map((name,i)=>({id:'d'+i,name}));
  const names=[['กานต์พิชชา','ใจดี'],['ธนกร','แสงทอง'],['ปิยาภรณ์','รักเรียน'],['ณัฐวุฒิ','อินทร์คำ'],['วรัญญา','บุญมา'],['ศุภชัย','ศรีสุข'],['พิมพ์ชนก','วงศ์คำ'],['นภัสสร','แสนดี']];
  const teachers=names.map(([firstName,lastName],i)=>({id:'t'+i,prefix:i%2?'นาย':'นางสาว',firstName,lastName,teacherCode:'T00'+(i+1),departmentId:'d'+(i%6),specialRoles:[],totalPeriods:18}));
  const rooms=levels.flatMap(l=>[1,2].map(n=>({id:l.id+'r'+n,name:l.name+'/'+n,levelId:l.id,planId:'plan1'})));
  const subjects=levels.flatMap((lv,li)=>['คณิตศาสตร์พื้นฐาน','วิทยาศาสตร์','ภาษาอังกฤษ','ภาษาไทย','สังคมศึกษา','ศิลปะ'].map((name,i)=>({id:lv.id+'s'+i,name,shortName:name,code:['ค','ว','อ','ท','ส','ศ'][i]+'3'+(li+1)+'101',departmentId:'d'+i,levelId:lv.id,periodsPerWeek:3,consecutiveAllowed:0,specialRoomId:''})));
  const assigns=rooms.flatMap((room,ri)=>subjects.filter(s=>s.levelId===room.levelId).slice(0,4).map((subject,si)=>({id:'a'+ri+'x'+si,teacherId:'t'+si,subjectId:subject.id,roomIds:[room.id],totalPeriods:3})));
  const schedule={};
  rooms.forEach((r,ri)=>subjects.filter(s=>s.levelId===r.levelId).slice(0,4).forEach((sub,si)=>{ const day=DAYS[ri%5], period=(si+ri)%7+1; if(ri<4) schedule[r.id+'_'+day+'_'+period]=[{id:'e'+ri+'x'+si,assignmentId:'a'+ri+'x'+si,teacherId:'t'+si,subjectId:sub.id,coTeacherIds:[]}]; }));
  const data={levels,depts,teachers,rooms,subjects,assigns,schedule,plans:[{id:'plan1',name:'วิทยาศาสตร์ – คณิตศาสตร์'}],specialRooms:[],meetings:[],locks:{}};
  Object.entries(data).forEach(([key,value])=>localStorage.setItem(prefix+'m2_'+key,JSON.stringify(value)));
  localStorage.setItem(prefix+'academicYear',JSON.stringify({year:'2569',semester:'1'}));
  localStorage.setItem(prefix+'schoolHeader',JSON.stringify({name:'โรงเรียนดาราวิทยาลัย',logo:''}));
  localStorage.setItem(prefix+'division','m2');
  localStorage.setItem(prefix+'seed_v1','1');
}
