export const TONES = [
 {ink:'#94324b',bg:'#faedf1'}, {ink:'#286a8b',bg:'#eaf4fa'},
 {ink:'#326d58',bg:'#eaf5ef'}, {ink:'#856019',bg:'#faf3e3'},
 {ink:'#665399',bg:'#f1edf9'}, {ink:'#936039',bg:'#fbefe5'},
 {ink:'#386e77',bg:'#e9f5f5'}, {ink:'#865781',bg:'#f7edf5'},
];
const hash=s=>Array.from(String(s||'')).reduce((h,c)=>(h*31+c.charCodeAt(0))>>>0,0);
export function departmentTone(dept){
 const names=['คณิต','วิทยาศาสตร์','ต่างประเทศ','ภาษาไทย','สังคม','ศิลปะ','สุขศึกษา','การงาน'];
 const i=names.findIndex(n=>dept?.name?.includes(n));
 return TONES[i<0?hash(dept?.id||dept?.name)%TONES.length:i];
}
export function levelTone(level){const n=String(level?.name||'').match(/\d+/);return TONES[n?(Number(n[0])-1)%TONES.length:hash(level?.id)%TONES.length];}
