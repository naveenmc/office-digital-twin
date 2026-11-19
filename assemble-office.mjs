import fs from 'node:fs';
import path from 'node:path';
for (const [source, destination] of Object.entries(JSON.parse(fs.readFileSync('office-files-next.json', 'utf8')))) {
  fs.mkdirSync(path.dirname(destination), {recursive: true});
  fs.renameSync(source, destination);
}
const [data] = JSON.parse(fs.readFileSync('build-parts-next.json', 'utf8'));
const web = 'Builds/WebGL';
const index = path.join(web, 'index.html');
const parts = data.parts.map(file => file.slice(web.length + 1));
const bootstrap = `async function loadOfficeData() {
  const parts = ${JSON.stringify(parts)};
  const buffers = await Promise.all(parts.map(async (url) => {
    const response = await fetch(url);
    if (!response.ok) throw new Error('Could not load ' + url);
    return response.arrayBuffer();
  }));
  const zipped = new Blob(buffers);
  const hash = new Uint8Array(await crypto.subtle.digest('SHA-256', await zipped.arrayBuffer()));
  const hex = Array.from(hash, byte => byte.toString(16).padStart(2, '0')).join('');
  if (hex !== ${JSON.stringify(data.sha256)}) throw new Error('Office data download failed its integrity check. Reload to retry.');
  const unpacked = await new Response(zipped.stream().pipeThrough(new DecompressionStream('gzip'))).blob();
  return URL.createObjectURL(unpacked);
}
`;
let html = fs.readFileSync(index, 'utf8');
if (!html.includes('createUnityInstance(canvas, config,') || !html.includes('script.onload = () => {')) throw new Error('Unexpected Unity template');
html = html.replace('      script.onload = () => {', bootstrap + '      script.onload = async () => {\n        try { config.dataUrl = await loadOfficeData(); } catch (error) { unityShowBanner(error.message, "error"); return; }');
fs.writeFileSync(index, html);
