// Publishes the Apps Script backend logic to the frontend's public/ folder so
// Vercel serves it at /backend-logic.js. The self-updating Apps Script loader
// fetches that URL and runs it, so backend changes go live on merge with no
// re-paste. Runs automatically before every `vite build`.
import { copyFileSync, mkdirSync } from 'node:fs';

mkdirSync('public', { recursive: true });
copyFileSync('APPS_SCRIPT_Code.gs', 'public/backend-logic.js');
console.log('sync-backend: copied APPS_SCRIPT_Code.gs -> public/backend-logic.js');
