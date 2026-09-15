export function readLiveConfig(env = {}) {
  const preview = env.VITE_LIVE_FIREBASE === 'false';
  const firebase = Object.fromEntries(['apiKey','authDomain','projectId','storageBucket','messagingSenderId','appId'].map(key => [key, env['VITE_FIREBASE_'+key.replace(/[A-Z]/g, c=>'_'+c).toUpperCase()] || '']));
  const missing = ['apiKey','authDomain','projectId','appId'].filter(k=>!firebase[k] || firebase[k].includes('YOUR'));
  const collection = env.VITE_FIRESTORE_COLLECTION || 'timetable';
  return {preview,firebase,collection,ready:!preview&&!missing.length&&['timetable','timetable_dev'].includes(collection),missing};
}
export function schoolAccount(user) { return !!user?.emailVerified && /^[^@]+@web1[.]dara[.]ac[.]th$/.test(user.email || ''); }
export function effectivePermissions(data, admin) { return admin ? {...data,divisions:{p1:true,p2:true,m1:true,m2:true,canEdit:true,isTeacher:false}} : data || {divisions:{}}; }
export function canonical(value) { return JSON.stringify(sort(value)); }
function sort(v) { return Array.isArray(v)?v.map(sort):v&&typeof v==='object'?Object.fromEntries(Object.keys(v).sort().map(k=>[k,sort(v[k])])):v; }
