// Bake an exported admin localStorage dump (warz_admin_data.json = {events, codes})
// into events-data.js (the baked-in baseline). Leaves warz_data.json untouched.
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const raw = fs.readFileSync(path.join(root, 'warz_admin_data.json'), 'utf8').replace(/^﻿/, '');
const data = JSON.parse(raw);

const events = Array.isArray(data.events) ? data.events : [];
const codes  = Array.isArray(data.codes)  ? data.codes  : [];

const js =
  'window.WARZ_EVENTS = ' + JSON.stringify(events, null, 2) +
  ';\n\nwindow.WARZ_MASTER_CODES = ' + JSON.stringify(codes, null, 2) + ';\n';
fs.writeFileSync(path.join(root, 'events-data.js'), js, 'utf8');

const winners = events.reduce((s, e) => s + (e.winners ? e.winners.length : 0), 0);
console.log('events:', events.length, '| winners:', winners, '| codes:', codes.length);
console.log('names:', events.map(e => e.name).join(' / '));
