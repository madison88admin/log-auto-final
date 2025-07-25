//if (process.env.NODE_ENV === 'development') {
  // @ts-ignore
  //import('./mocks/browser').then(({ worker }) => {
  //worker.start();
  //});
//}

import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.tsx'
import './index.css'


ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)