import fs from 'node:fs';
import { createHash } from 'node:crypto';
for (const item of JSON.parse(fs.readFileSync('build-parts.json','utf8'))) {
 const bytes=Buffer.concat(item.parts.map(p=>fs.readFileSync(p)));
 if(createHash('sha256').update(bytes).digest('hex')!==item.sha256)throw Error('Build part checksum failed: '+item.target);
 fs.writeFileSync(item.target,bytes);
 for(const part of item.parts)fs.unlinkSync(part);
 console.log('Restored '+item.target);
}
