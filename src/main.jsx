import {seedPreview} from "./renovation/model.mjs";
try { seedPreview(); } catch(e) { console.warn("Preview storage unavailable", e); }
import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
