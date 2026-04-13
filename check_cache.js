const d = JSON.parse(require('fs').readFileSync('group_messages_cache.json','utf8'));
const w = d.filter(m => m.pushName);
console.log('Total:', d.length, 'Con pushName:', w.length);
if (w.length > 0) {
  const n = new Set(w.map(m => m.pushName));
  console.log('Nombres:', [...n].join(', '));
} else {
  console.log('Keys primer msg:', Object.keys(d[0] || {}).join(', '));
  console.log('Participant:', d[0]?.key?.participant);
}
