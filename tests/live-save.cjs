const fs=require('node:fs'),assert=require('node:assert/strict'),path=require('node:path'),esbuild=require('esbuild');
const root=path.resolve(__dirname,'..');
(async()=>{
 let source=fs.readFileSync(root+'/src/App.jsx','utf8');
 source=source.replace(/import \{([^}]+)\} from "firebase\/firestore";/,'const {$1}=globalThis.__fsMocks;');
 const start=source.indexOf('const getFB=()=>{'),end=source.indexOf('// Firestore helpers',start);
 source=source.slice(0,start)+'const getFB=()=>({db:{}});\n'+source.slice(end)+'\nexport {fsSaveTimetable};';
 let current={levels:[],schedule:{old:[]}},written=null;
 globalThis.__fsMocks={doc:(_db,name,id)=>({name,id}),runTransaction:async(_db,fn)=>fn({get:async()=>({exists:()=>true,data:()=>current}),set:(ref,payload)=>{written={ref,payload};}})};
 await esbuild.build({stdin:{contents:source,resolveDir:root+'/src',loader:'jsx'},bundle:true,platform:'node',format:'cjs',jsx:'automatic',loader:{'.css':'empty'},define:{'import.meta.env':'{}'},external:['react','react-dom'],outfile:root+'/verify-save.cjs',logLevel:'silent'});
 const {fsSaveTimetable}=require(root+'/verify-save.cjs');
 const {canonical}=await import('../src/renovation/live-config.mjs');
 const expected=canonical(current),next={levels:[],schedule:{next:[]},notADataField:'ignored'};
 await fsSaveTimetable('m2',next,expected);assert.equal(written.ref.name,'timetable');assert.equal(written.ref.id,'m2');assert.deepEqual(written.payload,{levels:[],schedule:{next:[]}});
 current={levels:[],schedule:{someoneElse:[]}};written=null;
 await assert.rejects(()=>fsSaveTimetable('m2',next,expected),/เครื่องอื่น/);assert.equal(written,null,'conflicting remote document must never be replaced');
 console.log('PASS live save: configured path, field selection, whole schedule replacement and concurrent-edit rejection');
})().catch(e=>{console.error(e);process.exitCode=1;});
