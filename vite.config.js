import { defineConfig } from 'vite';
import { readFileSync, writeFileSync, readdirSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';
import { createHash } from 'node:crypto';

// Al construir, el service worker recibe la lista de todos los archivos de la app (con sus nombres
// con huella) y una versión nueva, para que la app funcione sin internet desde la primera visita
// y se actualice sola cuando se publica un cambio.
function serviceWorkerAssets() {
  let outDir;
  return {
    name: 'sw-assets',
    apply: 'build',
    configResolved(c) { outDir = c.build.outDir; },
    closeBundle() {
      const files = [];
      const walk = (d) => readdirSync(d).forEach((f) => {
        const p = join(d, f);
        if (statSync(p).isDirectory()) walk(p);
        else files.push(relative(outDir, p).split('\\').join('/'));
      });
      walk(outDir);
      const assets = ['./', ...files.filter((f) => f !== 'sw.js' && !f.endsWith('.map') && f !== 'vendor/pdfjs/LICENSE')];
      const hash = createHash('sha1');
      for (const f of files.filter((x) => x !== 'sw.js').sort()) hash.update(f).update(readFileSync(join(outDir, f)));
      const sw = join(outDir, 'sw.js');
      writeFileSync(sw, readFileSync(sw, 'utf8')
        .replace('__VERSION__', hash.digest('hex').slice(0, 10))
        .replace('__ASSETS__', JSON.stringify(assets, null, 2)));
    },
  };
}

export default defineConfig({
  plugins: [serviceWorkerAssets()],
  server: { port: 5173 },
  build: { outDir: 'dist', emptyOutDir: true },
});
