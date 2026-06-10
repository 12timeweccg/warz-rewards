// Convert the data pulled from Netlify Blobs (_published_pull.json) into the
// local baked-in files: events-data.js + warz_data.json
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const pull = JSON.parse(fs.readFileSync(path.join(root, '_published_pull.json'), 'utf8'));

const events = Array.isArray(pull.events) ? pull.events : [];
const codes  = Array.isArray(pull.codes)  ? pull.codes  : [];
const items  = Array.isArray(pull.items)  ? pull.items  : [];

// events-data.js
const js =
  'window.WARZ_EVENTS = ' + JSON.stringify(events, null, 2) +
  ';\n\nwindow.WARZ_MASTER_CODES = ' + JSON.stringify(codes, null, 2) + ';\n';
fs.writeFileSync(path.join(root, 'events-data.js'), js, 'utf8');

// warz_data.json (only overwrite if the pull actually has items)
if (items.length) {
  fs.writeFileSync(path.join(root, 'warz_data.json'),
    JSON.stringify({ items }, null, 2) + '\n', 'utf8');
}

const winners = events.reduce((s, e) => s + (e.winners ? e.winners.length : 0), 0);
console.log(`events-data.js written: ${events.length} events, ${winners} winners, ${codes.length} codes`);
console.log(`warz_data.json written: ${items.length} items`);
