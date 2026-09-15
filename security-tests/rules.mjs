import {readFileSync} from 'node:fs';
import {initializeTestEnvironment,assertSucceeds,assertFails} from '@firebase/rules-unit-testing';
import {doc,setDoc,getDoc,getDocs,collection,deleteDoc,updateDoc} from 'firebase/firestore';
const env=await initializeTestEnvironment({projectId:'demo-dara-rules',firestore:{rules:readFileSync(new URL('../firestore.rules',import.meta.url),'utf8')}});
const data={levels:[],plans:[],depts:[],teachers:[],subjects:[],rooms:[],specialRooms:[],assigns:[],meetings:[],schedule:{},locks:{}};
const ctx=(uid,extra={})=>env.authenticatedContext(uid,{email:uid+'@web1.dara.ac.th',email_verified:true,...extra}).firestore();
try{
 await env.withSecurityRulesDisabled(async c=>{for(const [uid,divisions] of Object.entries({teacher:{m2:true,isTeacher:true,canEdit:false},editor:{m2:true,canEdit:true},other:{m1:true}}))await setDoc(doc(c.firestore(),'permissions',uid),{displayName:uid,email:uid+'@web1.dara.ac.th',divisions});for(const name of ['timetable','timetable_dev'])await setDoc(doc(c.firestore(),name,'m2'),data);});
 const teacher=ctx('teacher'),editor=ctx('editor'),admin=ctx('admin',{admin:true});
 for(const db of [env.unauthenticatedContext().firestore(),ctx('outside',{email:'outside@example.com'}),ctx('unverified',{email_verified:false})]){await assertFails(getDoc(doc(db,'timetable','m2')));await assertFails(setDoc(doc(db,'permissions','teacher'),{displayName:'x'}));}
 await assertSucceeds(getDoc(doc(teacher,'permissions','teacher')));
 await assertFails(getDoc(doc(teacher,'permissions','editor')));await assertFails(getDocs(collection(teacher,'permissions')));
 await assertFails(updateDoc(doc(teacher,'permissions','teacher'),{'divisions.canEdit':true}));
 await assertFails(setDoc(doc(ctx('new'),'permissions','new'),{displayName:'new',email:'new@web1.dara.ac.th',divisions:{m2:true}}));
 await assertSucceeds(setDoc(doc(ctx('new'),'permissions','new'),{displayName:'new',email:'new@web1.dara.ac.th'}));
 await assertSucceeds(updateDoc(doc(teacher,'permissions','teacher'),{displayName:'Updated'}));
 for(const name of ['timetable','timetable_dev']){
   await assertSucceeds(getDoc(doc(teacher,name,'m2')));await assertFails(getDoc(doc(teacher,name,'m1')));
   await assertFails(setDoc(doc(teacher,name,'m2'),data));await assertSucceeds(setDoc(doc(editor,name,'m2'),data));
   await assertFails(setDoc(doc(editor,name,'m1'),data));await assertFails(deleteDoc(doc(editor,name,'m2')));
   await assertFails(setDoc(doc(editor,name,'m2'),{...data,schedule:[]}));await assertFails(setDoc(doc(editor,name,'m2'),{...data,unexpected:true}));
   await assertFails(setDoc(doc(admin,name,'m2','nested','x'),data));await assertSucceeds(setDoc(doc(admin,name,'p1'),data));
 }
 await assertSucceeds(getDocs(collection(admin,'permissions')));await assertSucceeds(updateDoc(doc(admin,'permissions','teacher'),{divisions:{m1:true,canEdit:true}}));
 await assertFails(getDoc(doc(teacher,'timetable','m2')));await assertFails(setDoc(doc(editor,'unknown','x'),data));
 console.log('PASS Firestore rules: anonymous, domain, verification, own profile, privilege escalation, roles, divisions, revocation, schema, dev/prod and nested paths');
}finally{await env.cleanup();}
