// Presupuesto de peso (npm run size, después de npm run build). La app se abre en celulares
// con datos móviles: lo que se descarga al abrirla debe ser poco. Si algo se pasa, el comando falla.
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { gzipSync } from 'node:zlib';

const KB = 1024;
// Lo que se descarga al abrir la app (sin librerías que se cargan solo al usarlas).
const BUDGET = { js: 40 * KB, css: 10 * KB, html: 8 * KB, total: 55 * KB }; // comprimido (gzip)
const dist = 'dist';
const gz = (f) => gzipSync(readFileSync(f), { level: 9 }).length;
const html = readFileSync(join(dist, 'index.html'), 'utf8');
const startFiles = [...html.matchAll(/(?:src|href)="\/?(assets\/[^"]+)"/g)].map((m) => m[1]);
const sizes = { js: 0, css: 0, html: gz(join(dist, 'index.html')) };
for (const f of startFiles) sizes[f.endsWith('.css') ? 'css' : 'js'] += gz(join(dist, f));
sizes.total = sizes.js + sizes.css + sizes.html;

const row = (k) => `${k.padEnd(6)} ${(sizes[k] / KB).toFixed(1).padStart(6)} KB  (máx. ${BUDGET[k] / KB} KB)${sizes[k] > BUDGET[k] ? '  ← SE PASA' : ''}`;
console.log('Peso al abrir la app (gzip):');
['html', 'js', 'css', 'total'].forEach((k) => console.log('  ' + row(k)));
// Librerías pesadas: deben cargarse solo al usarse, nunca al abrir.
const vendor = [];
const walk = (d) => readdirSync(d).forEach((f) => { const p = join(d, f); statSync(p).isDirectory() ? walk(p) : vendor.push(p); });
walk(join(dist, 'vendor'));
console.log('Se cargan solo al usarse:');
vendor.filter((f) => f.endsWith('.js')).forEach((f) => console.log(`  ${f.slice(dist.length + 1).padEnd(32)} ${(gz(f) / KB).toFixed(0).padStart(5)} KB`));
const eager = vendor.filter((f) => html.includes(f.slice(dist.length + 1)));
if (eager.length) console.log('ERROR: estas librerías se cargan al abrir la app:', eager.join(', '));
const over = Object.keys(BUDGET).filter((k) => sizes[k] > BUDGET[k]);
if (over.length || eager.length) { console.log('\nPresupuesto de peso superado.'); process.exit(1); }
console.log('\nDentro del presupuesto ✓');
