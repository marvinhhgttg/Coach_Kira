import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';
import './index.css';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>
);

// Zweitinstanz: deutlicher Hinweis, damit v2 nie mit dem Live-Cockpit verwechselt wird
if (typeof window !== 'undefined' && /^coachkira-v2\./.test(window.location.hostname)) {
  const b = document.createElement('div');
  b.textContent = 'v2 · Testinstanz – Daten aus dem echten Sheet, automatische Kira-Texte aus';
  b.setAttribute('role', 'status');
  b.style.cssText = 'position:fixed;left:50%;bottom:10px;transform:translateX(-50%);z-index:9999;background:#7c3aed;color:#fff;font:600 12px/1.2 system-ui,sans-serif;padding:6px 12px;border-radius:999px;box-shadow:0 2px 8px rgba(0,0,0,.35);pointer-events:none';
  document.body.appendChild(b);
  document.title = 'v2 · ' + document.title;
}
