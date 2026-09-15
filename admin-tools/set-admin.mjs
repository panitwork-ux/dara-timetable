import {initializeApp,applicationDefault} from 'firebase-admin/app';
import {getAuth} from 'firebase-admin/auth';
const [projectId,uid,confirmation]=process.argv.slice(2);
if(!projectId||!uid||confirmation!=='--confirm-admin')throw Error('Usage: node set-admin.mjs PROJECT_ID VERIFIED_UID --confirm-admin');
initializeApp({credential:applicationDefault(),projectId});
const user=await getAuth().getUser(uid);
if(!user.emailVerified||!/^\S+@web1[.]dara[.]ac[.]th$/.test(user.email||''))throw Error('Expected verified school account');
await getAuth().setCustomUserClaims(uid,{...user.customClaims,admin:true});
console.log('Admin claim set for verified UID:',uid,'Sign out and sign in again.');
