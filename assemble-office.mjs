import fs from 'node:fs';
import path from 'node:path';
for (const [source, destination] of Object.entries(JSON.parse(fs.readFileSync('office-files.json', 'utf8')))) {
  fs.mkdirSync(path.dirname(destination), {recursive: true});
  fs.renameSync(source, destination);
}
await import('./restore-build.mjs');
