import {loadEnv} from 'vite';
import {readLiveConfig} from '../src/renovation/live-config.mjs';
const config=readLiveConfig({...loadEnv('production',process.cwd(),'VITE_'),...process.env});
if(!config.ready)throw Error('Live Firebase configuration missing or invalid. Set VITE_FIREBASE_* and VITE_LIVE_FIREBASE=true.');
console.log('Live configuration present. Project:',config.firebase.projectId,'Collection:',config.collection);
