import http from 'node:http';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { pathToFileURL } from 'node:url';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const directory = path.dirname(fileURLToPath(import.meta.url));
const dataDirectory = process.env.OFFICE_DATA_DIR || directory;
fs.mkdirSync(dataDirectory, { recursive: true, mode: 0o700 });
const dataFile = path.join(dataDirectory, 'office-data.json');
const buildRoot = path.resolve(directory, '../Builds/WebGL');
const host = process.env.OFFICE_HOST || '127.0.0.1';
const port = Number(process.env.OFFICE_PORT || process.env.PORT || 8787);
const allowedOrigin = process.env.OFFICE_ALLOWED_ORIGIN || '';
const sessions = new Map();
const identities = new Map();
const messages = [];
const voiceSignals = new Map();
const cameraFrames = new Map();
let pings = [];
let presentation = null;
const run = promisify(execFile);
const mediaRoot = path.join(dataDirectory, 'presentation-media');
let uploading = false;
const mediaTypes = { png:'image/png', jpg:'image/jpeg', jpeg:'image/jpeg', mp4:'video/mp4', webm:'video/webm' };
const docs = new Set(['pdf','ppt','pptx']);
async function uploadPresentation(req, token, fileName) {
  const player = authenticate(token);
  const name = text(fileName, 150);
  const ext = name && path.extname(name).slice(1).toLowerCase();
  if (!ext || (!Object.hasOwn(mediaTypes, ext) && !docs.has(ext))) fail('Choose PDF, PPT, PPTX, PNG, JPG, MP4 or WebM.');
  if (uploading) fail('A presentation is already being prepared.', 409);
  uploading = true;
  const id = randomUUID();
  const folder = path.join(mediaRoot, id);
  try {
    fs.mkdirSync(folder, { recursive:true, mode:0o700 });
    const source = path.join(folder, 'source.' + ext);
    let bytes = 0;
    const chunks = [];
    for await (const chunk of req) {
      bytes += chunk.length;
      if (bytes > 24 * 1024 * 1024) fail('Presentation exceeds 24 MB.', 413);
      chunks.push(chunk);
    }
    if (!bytes) fail('The file is empty.');
    fs.writeFileSync(source, Buffer.concat(chunks), { mode:0o600 });
    const pages = [];
    if (docs.has(ext)) {
      let pdf = source;
      if (ext !== 'pdf') {
        try {
          await run(process.env.OFFICE_SOFFICE || 'soffice', ['-env:UserInstallation=' + pathToFileURL(path.join(folder,'profile')).href,
            '--headless','--convert-to','pdf','--outdir',folder,source], { timeout:45000 });
          pdf = path.join(folder,'source.pdf');
        } catch { fail('PowerPoint conversion failed. Try exporting your slides as PDF.', 422); }
      }
      try {
        await run(process.env.OFFICE_PDFTOPPM || 'pdftoppm', ['-f','1','-l','30','-scale-to','1200','-png',pdf,path.join(folder,'page')], { timeout:45000 });
      } catch { fail('Could not render this PDF. Check the file and converter.', 422); }
      for (const file of fs.readdirSync(folder).filter(f => /^page-\d+\.png$/.test(f))
        .sort((a,b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]))) pages.push('/media/' + id + '/' + file);
      if (!pages.length) fail('The document contains no readable pages.', 422);
    } else {
      fs.renameSync(source, path.join(folder, 'media.' + ext));
      pages.push('/media/' + id + '/media.' + ext);
    }
    const previousId = presentation?.mediaId;
    presentation = { title:name, name:player.name, owner:player.id, mediaId:id,
      kind:ext === 'mp4' || ext === 'webm' ? 'video' : 'image', pages, page:0, url:pages[0] };
    if (previousId) fs.rmSync(path.join(mediaRoot,previousId), { recursive:true, force:true });
    return { ok:true };
  } finally {
    uploading = false;
    if (presentation?.mediaId !== id) fs.rmSync(folder, { recursive:true, force:true });
  }
}
let bookings = [];
if (fs.existsSync(dataFile)) {
  const saved = JSON.parse(fs.readFileSync(dataFile, 'utf8'));
  if (!Array.isArray(saved.bookings)) throw new Error('Invalid booking storage. Restore a valid office-data.json.');
  bookings = saved.bookings;
  for (const person of saved.identities || []) identities.set(person.token, person);
}
const now = () => Math.floor(Date.now() / 1000);
const text = (value, limit) => typeof value === 'string' && value.trim().length > 0 && value.trim().length <= limit && !/[\u0000-\u001f\u007f]/.test(value) ? value.trim() : null;
const fail = (message, status = 400) => { const error = new Error(message); error.status = status; throw error; };
const persist = () => {
  const temporary = dataFile + '.tmp';
  fs.writeFileSync(temporary, JSON.stringify({ bookings, identities: [...identities.values()] }, null, 2), { mode: 0o600 });
  fs.renameSync(temporary, dataFile);
};
const equipment = new Map([
  ['office-lights', true], ['office-auto-detection', true],
  ['office-door-1', false], ['office-door-2', false], ['office-door-3', false], ['office-door-4', false], ['office-door-5', false],
  ['office-display-1', true], ['island-glass-walls', true], ['room-1-glass-walls', true],
]);
const fixtureModes = new Map();
const chairs = new Map();
const reactionIcons = { smile: '😊', heart: '❤️', laugh: '😂', wow: '😮' };
const snapshot = (viewerId) => {
  for (const [key, player] of sessions) if (player.seen < now() - 60) sessions.delete(key);
  const activeIds = new Set([...sessions.values()].map(player => player.id));
  for (const chair of chairs.values()) if (!activeIds.has(chair.occupant)) chair.occupant = '';
  pings = pings.filter(p => p.expires > now());
  return {
    bookings: bookings.filter(b => b.end > now()),
    participants: [...sessions.values()].map(({ token, ...person }) => person),
    messages: messages.filter(message => !message.recipient || message.sender === viewerId || message.recipient === viewerId), pings, presentation,
    chairs: [...chairs.values()],
    fixtures: Array.from(fixtureModes, ([id, mode]) => ({ id, mode })),
    equipment: Array.from(equipment, ([id, enabled]) => ({ id, enabled }))
  };
};
function authenticate(token) {
  if (typeof token !== 'string' || token.length > 100) fail('Join a session first.', 401);
  const session = sessions.get(token);
  if (!session || session.seen < now() - 60) fail('Session expired. Join again.', 401);
  session.seen = now();
  return session;
}
// A single Node process serializes the conflict check and booking write without awaits.
function transact(body) {
  if (body.action === 'join') {
    const name = text(body.name, 24);
    if (!name) fail('Enter a display name (1–24 characters).');
    snapshot();
    if (sessions.size >= 32) fail('This prototype is limited to 32 simultaneous users.', 429);
    const known = typeof body.token === 'string' ? identities.get(body.token) : null;
    const token = known ? known.token : randomUUID() + randomUUID();
    const player = { token, id: known ? known.id : randomUUID(), name, x: 0, y: 0, z: 0, yaw: 0, seen: now() };
    identities.set(token, { token, id: player.id, name });
    persist();
    sessions.set(token, player);
    return { ok: true, token, id: player.id, state: snapshot(player.id) };
  }
  const player = authenticate(body.token);
  if (body.action === 'voice-camera') {
    if (typeof body.enabled !== 'boolean') fail('Invalid camera state.');
    player.camera = body.enabled;
    if (!body.enabled) cameraFrames.delete(player.id);
    return { ok: true };
  } else if (body.action === 'camera-frames') {
    const time = Date.now();
    for (const [id, frame] of cameraFrames) if (time - frame.time > 1500) cameraFrames.delete(id);
    if (body.frame !== undefined) {
      if (!player.camera || typeof body.frame !== 'string' || body.frame.length > 16384 ||
          !/^\/9j\/[A-Za-z0-9+/]+={0,2}$/.test(body.frame)) fail('Invalid camera frame.');
      const previous = cameraFrames.get(player.id);
      if (!previous || time - previous.time >= 150) cameraFrames.set(player.id, { time, jpeg: body.frame });
    }
    if (!Array.isArray(body.from) || body.from.length > 31) fail('Invalid camera subscriptions.');
    const requested = new Set(body.from);
    const frames = [];
    for (const peer of sessions.values()) {
      if (peer.id === player.id || !peer.camera || peer.seen < now() - 60 || !requested.has(peer.id)) continue;
      const frame = cameraFrames.get(peer.id);
      if (frame) frames.push({ id: peer.id, time: frame.time, jpeg: frame.jpeg });
    }
    return { ok: true, frames };
  } else if (body.action === 'voice-poll') {
    const state = snapshot();
    const signals = (voiceSignals.get(player.id) || []).filter(s => s.time > now() - 30);
    voiceSignals.delete(player.id);
    return { ok: true, peers: state.participants.map(p => ({ id: p.id, name: p.name, camera: p.camera === true })), signals };
  } else if (body.action === 'voice-signal') {
    const target = [...sessions.values()].find(p => p.id === body.to && p.seen >= now() - 60);
    if (!target || target.id === player.id) fail('Voice recipient is not connected.', 404);
    const data = body.signal;
    if (!data || !['offer', 'answer', 'candidate'].includes(data.type) || JSON.stringify(data).length > 24000) fail('Invalid voice signal.');
    if ((data.type === 'offer' || data.type === 'answer') && typeof data.sdp !== 'string') fail('Missing voice description.');
    if (data.type === 'candidate' && (!data.candidate || typeof data.candidate.candidate !== 'string')) fail('Missing voice candidate.');
    const queue = (voiceSignals.get(target.id) || []).filter(s => s.time > now() - 30);
    if (queue.length >= 64) fail('Voice signal queue is full.', 429);
    queue.push({ from: player.id, signal: data, time: now() }); voiceSignals.set(target.id, queue);
    return { ok: true };
  } else if (body.action === 'equipment') {
    if (!equipment.has(body.equipmentId) || typeof body.enabled !== 'boolean') fail('Invalid equipment state.');
    equipment.set(body.equipmentId, body.enabled);
  } else if (body.action === 'chair') {
    const id = body.chairId;
    const operation = body.mode;
    if (!Number.isInteger(id) || id < 1 || id > 1024 ||
        !['sit', 'stand', 'slide-in', 'slide-out', 'left', 'right', 'move'].includes(operation)) fail('Invalid chair control.');
    snapshot();
    const chair = chairs.get(id) || { id, occupant: '', slide: 0, swivel: 0 };
    if (chair.occupant && chair.occupant !== player.id) fail('That seat is occupied.', 409);
    if (operation === 'sit') {
      for (const other of chairs.values()) if (other.occupant === player.id) other.occupant = '';
      chair.occupant = player.id;
    } else if (operation === 'stand') {
      chair.occupant = '';
    } else if (operation === 'slide-in' || operation === 'slide-out') {
      chair.slide = Math.max(0, Math.min(2, chair.slide + (operation === 'slide-out' ? 1 : -1)));
    } else if (operation === 'move') {
      if (chair.occupant) fail('Stand before moving the chair.', 409);
      if (!Number.isFinite(body.x) || !Number.isFinite(body.z) || Math.hypot(body.x, body.z) > 45)
        fail('Invalid chair position.');
      chair.x = body.x;
      chair.z = body.z;
    } else {
      chair.swivel = (chair.swivel + (operation === 'right' ? 1 : -1) + 24) % 24;
    }
    chairs.set(id, chair);
  } else if (body.action === 'fixture') {
    if (!Number.isInteger(body.fixtureId) || body.fixtureId < 1 || body.fixtureId > 2048 ||
        !['auto', 'on', 'off'].includes(body.mode)) fail('Invalid ceiling light control.');
    if (body.mode === 'auto') fixtureModes.delete(body.fixtureId);
    else fixtureModes.set(body.fixtureId, body.mode);
  } else if (body.action === 'sync') {
    for (const axis of ['x', 'y', 'z', 'yaw']) {
      if (!Number.isFinite(body[axis]) || Math.abs(body[axis]) > 10000) fail('Invalid player position.');
    }
    for (const axis of ['x', 'y', 'z', 'yaw']) player[axis] = body[axis];
  } else if (body.action === 'book') {
    const title = text(body.title, 80);
    if (!title || body.roomId !== 'meeting-01') fail('Choose Meeting 01 and enter a title (1–80 characters).');
    if (!Number.isInteger(body.start) || !Number.isInteger(body.end) || body.start < now() - 60 || body.start > now() + 7 * 86400 || body.end - body.start < 15 * 60 || body.end - body.start > 120 * 60) fail('Bookings must last 15–120 minutes and begin within the next seven days.');
    if (bookings.some(b => b.roomId === body.roomId && b.start < body.end && b.end > body.start)) fail('That time overlaps another reservation.', 409);
    if (bookings.filter(b => b.owner === player.id && b.end > now()).length >= 8) fail('Limit of eight active reservations per session reached.', 429);
    const prior = bookings;
    bookings = [...bookings.filter(b => b.end > now()), { id: randomUUID(), roomId: body.roomId, title, owner: player.id, start: body.start, end: body.end }];
    try { persist(); } catch (error) { bookings = prior; throw error; }
  } else if (body.action === 'cancel') {
    const booking = bookings.find(b => b.id === body.bookingId);
    if (!booking) fail('Reservation no longer exists.', 404);
    if (booking.owner !== player.id) fail('Only the booking owner can cancel this reservation.', 403);
    const prior = bookings;
    bookings = bookings.filter(b => b !== booking);
    try { persist(); } catch (error) { bookings = prior; throw error; }
  } else if (body.action === 'present') {
    const title = text(body.title, 80);
    const url = text(body.text, 500);
    let parsed;
    try { parsed = new URL(url); } catch { fail('Enter a valid HTTPS presentation link.'); }
    if (!title || parsed.protocol !== 'https:' || parsed.username || parsed.password) fail('Enter a title and an HTTPS link without embedded credentials.');
    if (presentation?.mediaId) fs.rmSync(path.join(mediaRoot,presentation.mediaId), { recursive:true, force:true });
    presentation = { title, url: parsed.href, name: player.name };
  } else if (body.action === 'presentation-page') {
    if (presentation?.owner !== player.id) fail('Only the presenter can change the page.', 403);
    if (!Number.isInteger(body.page) || body.page < 0 || body.page >= presentation.pages.length) fail('Invalid page.');
    presentation.page = body.page;
    presentation.url = presentation.pages[body.page];
  } else if (body.action === 'ping') {
    const label = text(body.text, 100);
    if (!label || ['x', 'y', 'z'].some(axis => !Number.isFinite(body[axis]) || Math.abs(body[axis]) > 10000)) fail('Invalid location marker.');
    pings.push({ id: randomUUID(), name: player.name, label, x: body.x, y: body.y, z: body.z, expires: now() + 10 });
    pings = pings.slice(-12);
  } else if (body.action === 'chat' || body.action === 'react') {
    const reaction = body.action === 'react' && typeof body.text === 'string' && Object.hasOwn(reactionIcons, body.text) ? body.text : '';
    const message = body.action === 'chat' ? text(body.text, 240) : reactionIcons[reaction];
    if (!message) fail('Enter a message (1–240 characters).');
    const last = messages.findLast(m => m.sender === player.id);
    if (last && last.time > now() - 1) fail('Wait a moment before sending another message.', 429);
    const recipient = body.to || '';
    if (recipient && (typeof recipient !== 'string' || ![...sessions.values()].some(p => p.id === recipient && p.seen >= now() - 60)))
      fail('That player is offline.', 404);
    messages.push({ id: randomUUID(), sender: player.id, recipient, reaction, name: player.name, text: message, time: now() });
    if (messages.length > 30) messages.shift();
  } else fail('Unknown action.');
  return { ok: true, state: snapshot(player.id) };
}
const mime = { '.html': 'text/html; charset=utf-8', '.js': 'application/javascript', '.wasm': 'application/wasm', '.data': 'application/octet-stream', '.json': 'application/json', '.css': 'text/css', '.png': 'image/png', '.ico': 'image/x-icon' };
const server = http.createServer(async (req, res) => {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  const origin = req.headers.origin;
  const sameOrigin = origin === `http://${req.headers.host}` || origin === `https://${req.headers.host}`;
  const corsAllowed = !origin || sameOrigin || (allowedOrigin && origin === allowedOrigin);
  if (!corsAllowed) { res.writeHead(403); res.end('Origin not allowed.'); return; }
  if (origin) { res.setHeader('Access-Control-Allow-Origin', origin); res.setHeader('Vary', 'Origin'); }
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return; }
  if (req.url === '/health' && (req.method === 'GET' || req.method === 'HEAD')) { res.writeHead(200, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }); res.end(req.method === 'HEAD' ? '' : JSON.stringify({ ok: true, service: 'office-multiplayer' })); return; }
  const address = new URL(req.url, 'http://localhost');
  if (address.pathname === '/api/presentation' && req.method === 'POST') {
    res.setHeader('Content-Type','application/json');
    try {
      const token = String(req.headers.authorization || '').replace(/^Bearer /,'');
      const result = await uploadPresentation(req, token, address.searchParams.get('name'));
      res.writeHead(200); res.end(JSON.stringify(result));
    } catch (error) {
      res.writeHead(error.status || 500); res.end(JSON.stringify({ ok:false, error:error.status ? error.message : 'Upload failed.' }));
    }
    return;
  }
  const media = /^\/media\/([a-f0-9-]{36})\/(page-\d+\.png|media\.(png|jpe?g|mp4|webm))$/.exec(address.pathname);
  if (media && ['GET','HEAD'].includes(req.method)) {
    try {
      if (media[1] !== presentation?.mediaId) { res.writeHead(404); res.end(); return; }
      const file = path.join(mediaRoot,media[1],media[2]);
      const length = fs.statSync(file).size;
      const range = /^bytes=(\d+)-(\d*)$/.exec(req.headers.range || '');
      const start = range ? Number(range[1]) : 0;
      const end = range?.[2] ? Math.min(Number(range[2]),length-1) : length-1;
      if (start > end) { res.writeHead(416,{'Content-Range':'bytes */'+length}); res.end(); return; }
      const headers = {'Content-Type':mediaTypes[path.extname(file).slice(1)],'Content-Length':end-start+1,'Accept-Ranges':'bytes','Cache-Control':'no-store'};
      if (range) headers['Content-Range'] = `bytes ${start}-${end}/${length}`;
      res.writeHead(range ? 206 : 200, headers);
      if (req.method === 'HEAD') res.end(); else fs.createReadStream(file,{start,end}).pipe(res);
    } catch { if (!res.headersSent) res.writeHead(404); res.end(); }
    return;
  }
  if (req.url === '/api/office' && req.method === 'POST') {
    res.setHeader('Content-Type', 'application/json'); res.setHeader('Cache-Control', 'no-store');
    try {
      let bytes = 0;
      const chunks = [];
      for await (const chunk of req) {
        bytes += chunk.length;
        if (bytes > 32768) fail('Request too large.', 413);
        chunks.push(chunk);
      }
      let body;
      try { body = JSON.parse(Buffer.concat(chunks).toString('utf8')); } catch { fail('Invalid JSON.'); }
      if (!body || typeof body !== 'object' || Array.isArray(body)) fail('Invalid request.');
      const result = transact(body);
      res.writeHead(200); res.end(JSON.stringify(result));
    } catch (error) {
      // transact must run before writing headers so conflict and auth failures keep their status.
      if (!res.headersSent) res.writeHead(error.status || 500);
      res.end(JSON.stringify({ ok: false, error: error.status ? error.message : 'Unable to save the update. Please retry.' }));
    }
    return;
  }
  if (req.method !== 'GET' && req.method !== 'HEAD') { res.writeHead(405); res.end(); return; }
  try {
    let pathname = decodeURIComponent(new URL(req.url, 'http://localhost').pathname);
    if (pathname === '/') pathname = '/index.html';
    const file = path.resolve(buildRoot, '.' + pathname);
    if (!file.startsWith(buildRoot + path.sep) || !fs.statSync(file).isFile() || !fs.realpathSync(file).startsWith(buildRoot + path.sep)) throw new Error('not found');
    const compressed = /\.(gz|br)$/.test(file);
    const plain = compressed ? file.replace(/\.(gz|br)$/, '') : file;
    res.setHeader('Content-Type', mime[path.extname(plain)] || 'application/octet-stream');
    if (compressed) res.setHeader('Content-Encoding', file.endsWith('.br') ? 'br' : 'gzip');
    const metadata = fs.statSync(file);
    const etag = `"${metadata.size.toString(16)}-${Math.trunc(metadata.mtimeMs).toString(16)}"`;
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('ETag', etag);
    if (req.headers['if-none-match'] === etag) { res.writeHead(304); res.end(); return; }
    res.writeHead(200);
    if (req.method === 'HEAD') res.end(); else fs.createReadStream(file).pipe(res);
  } catch {
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Build not found. In Unity use Office Digital Twin > Build WebGL, then reopen this address.');
  }
});
server.requestTimeout = 10000;
server.headersTimeout = 10000;
server.listen(port, host, () => console.log(`Office server: http://${host}:${port}`));
