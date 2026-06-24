/**
 * NexusTwin — Industry 6.0 Backend Server
 * Express REST API + WebSocket live telemetry + Excel export (SheetJS)
 */

const express = require('express');
const cors    = require('cors');
const http    = require('http');
const WebSocket = require('ws');
const XLSX    = require('xlsx');
const { v4: uuidv4 } = require('uuid');
const path    = require('path');
const fs      = require('fs');

const app    = express();
const server = http.createServer(app);
const wss    = new WebSocket.Server({ server });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

/* ══════════════════════════════════════════
   IN-MEMORY DATA STORE
══════════════════════════════════════════ */
let ASSETS = [
  { id:'CNC-07', name:'CNC Mill #7',     type:'CNC Machine',      rul:6,  health:42, status:'crit', zone:'A3', mfr:'Fanuc',      installedDate:'2020-03-15' },
  { id:'ROB-03', name:'Robot Arm 3',     type:'Robotic Arm',      rul:14, health:68, status:'warn', zone:'A3', mfr:'KUKA',       installedDate:'2021-06-10' },
  { id:'CNV-01', name:'Conveyor A1',     type:'Conveyor System',  rul:45, health:88, status:'ok',   zone:'B1', mfr:'Bosch',      installedDate:'2019-11-20' },
  { id:'PMP-09', name:'Hydraulic Pump',  type:'Hydraulic Pump',   rul:22, health:74, status:'warn', zone:'B2', mfr:'Parker',     installedDate:'2021-02-08' },
  { id:'CNC-12', name:'CNC Mill #12',    type:'CNC Machine',      rul:58, health:91, status:'ok',   zone:'A3', mfr:'Fanuc',      installedDate:'2022-04-01' },
  { id:'MOT-05', name:'Drive Motor 5',   type:'Electric Motor',   rul:71, health:95, status:'ok',   zone:'C4', mfr:'Siemens',    installedDate:'2022-09-14' },
  { id:'CMP-02', name:'Compressor 2',    type:'Compressor',       rul:33, health:82, status:'ok',   zone:'B2', mfr:'Atlas Copco',installedDate:'2020-07-30' },
  { id:'ROB-07', name:'Robot Arm 7',     type:'Robotic Arm',      rul:64, health:93, status:'ok',   zone:'C4', mfr:'ABB',        installedDate:'2022-01-22' },
  { id:'HEX-01', name:'Heat Exchanger',  type:'Heat Exchanger',   rul:41, health:86, status:'ok',   zone:'D1', mfr:'Alfa Laval', installedDate:'2020-12-05' },
];

let ALERTS = [
  { id:'ALT-001', type:'crit', msg:'CNC-07 bearing inner-race wear — RUL 6 days',      asset:'CNC-07', ack:false, ts: new Date().toISOString() },
  { id:'ALT-002', type:'warn', msg:'ROB-03 joint torque variance exceeding 18%',        asset:'ROB-03', ack:false, ts: new Date(Date.now()-660000).toISOString() },
  { id:'ALT-003', type:'warn', msg:'PMP-09 suction pressure dropping — cavitation risk',asset:'PMP-09', ack:false, ts: new Date(Date.now()-2040000).toISOString() },
  { id:'ALT-004', type:'info', msg:'CMP-02 maintenance window confirmed',               asset:'CMP-02', ack:true,  ts: new Date(Date.now()-3600000).toISOString() },
  { id:'ALT-005', type:'info', msg:'AI model retrained — accuracy +1.3%',               asset:'System', ack:true,  ts: new Date(Date.now()-10800000).toISOString() },
];

let WORK_ORDERS = [
  { id:'WO-2214', asset:'CNC-07', title:'Replace bearing pack B-7209',         status:'open',    priority:'Critical', tech:'RK', due:'2026-06-21', created:'2026-06-16', progress:0  },
  { id:'WO-2215', asset:'PMP-09', title:'Inspect hydraulic inlet filter',       status:'open',    priority:'High',     tech:'SP', due:'2026-06-26', created:'2026-06-16', progress:0  },
  { id:'WO-2210', asset:'ROB-03', title:'Joint 4 gear assembly inspection',     status:'inprog',  priority:'High',     tech:'AK', due:'2026-06-28', created:'2026-06-14', progress:45 },
  { id:'WO-2211', asset:'CMP-02', title:'Compressor valve V2 replacement',      status:'inprog',  priority:'Medium',   tech:'TM', due:'2026-07-02', created:'2026-06-15', progress:20 },
  { id:'WO-2204', asset:'CNV-01', title:'Belt tension adjustment A1',           status:'done',    priority:'Low',      tech:'SP', due:'2026-06-14', created:'2026-06-10', progress:100},
  { id:'WO-2206', asset:'MOT-05', title:'Motor 5 thermal paste reapplication',  status:'done',    priority:'Low',      tech:'RK', due:'2026-06-12', created:'2026-06-09', progress:100},
];

let MAINT_PLAN = [
  { date:'2026-06-21', asset:'CNC-07', task:'Replace bearing pack B-7209',  type:'Critical', dur:'4h',   tech:'RK', parts:'B-7209 × 2, Grease 500g' },
  { date:'2026-06-28', asset:'ROB-03', task:'Inspect & replace gear J4',    type:'High',     dur:'3h',   tech:'AK', parts:'Gear set GS-44, Torque sensor' },
  { date:'2026-06-28', asset:'PMP-09', task:'Clean inlet filter',           type:'High',     dur:'1.5h', tech:'SP', parts:'Filter element F-09' },
  { date:'2026-07-02', asset:'CMP-02', task:'Replace valve assembly V2',    type:'Medium',   dur:'2h',   tech:'TM', parts:'Valve V2-C, O-rings' },
  { date:'2026-07-05', asset:'CNV-01', task:'Belt tension adjustment',      type:'Medium',   dur:'1h',   tech:'SP', parts:'Tension gauge, Lubricant' },
  { date:'2026-07-10', asset:'MOT-05', task:'Thermal compound reapplication',type:'Low',     dur:'0.5h', tech:'RK', parts:'Thermal paste 100g' },
];

// Rolling telemetry buffer — last 500 readings per asset
const TELEMETRY_BUFFER = {};
ASSETS.forEach(a => { TELEMETRY_BUFFER[a.id] = []; });

/* ══════════════════════════════════════════
   SENSOR SIMULATION ENGINE
══════════════════════════════════════════ */
const SENSOR_BASES = {
  'CNC Machine':     { vibration:1.4, temperature:68, acoustic:74, current:14.2, flow:8.4 },
  'Robotic Arm':     { vibration:0.9, temperature:55, acoustic:62, current:11.8, flow:0   },
  'Conveyor System': { vibration:0.6, temperature:42, acoustic:55, current:8.6,  flow:12.0},
  'Hydraulic Pump':  { vibration:1.1, temperature:72, acoustic:68, current:16.4, flow:22.5},
  'Electric Motor':  { vibration:0.7, temperature:61, acoustic:58, current:18.2, flow:0   },
  'Compressor':      { vibration:1.2, temperature:65, acoustic:71, current:15.0, flow:18.0},
  'Heat Exchanger':  { vibration:0.4, temperature:85, acoustic:48, current:6.2,  flow:35.0},
};

function noisify(base, pct, stress = 1) {
  return +(base + (Math.random() - 0.5) * base * pct * stress).toFixed(3);
}

function generateTelemetry(asset) {
  const bases = SENSOR_BASES[asset.type] || SENSOR_BASES['CNC Machine'];
  const stress = asset.status === 'crit' ? 2.2 : asset.status === 'warn' ? 1.5 : 1.0;
  const ts = new Date().toISOString();
  const row = {
    timestamp: ts,
    assetId:   asset.id,
    assetName: asset.name,
    status:    asset.status,
    health:    asset.health,
    rul:       asset.rul,
    vibration:   bases.flow  ? noisify(bases.vibration,  0.12, stress) : null,
    temperature: noisify(bases.temperature, 0.04, stress),
    acoustic:    noisify(bases.acoustic,    0.05, stress),
    current:     bases.current ? noisify(bases.current, 0.06, stress) : null,
    flow:        bases.flow    ? noisify(bases.flow,    0.04, stress) : null,
    oee:         +(90 + Math.random() * 3).toFixed(2),
  };

  // Slowly degrade health/rul for stressed assets
  if (Math.random() > 0.97) {
    asset.health = Math.max(5, +(asset.health - 0.05).toFixed(2));
    asset.rul    = Math.max(1, +(asset.rul - 0.01).toFixed(2));
    row.health   = asset.health;
    row.rul      = asset.rul;
    if (asset.health < 40)  asset.status = 'crit';
    else if (asset.health < 72) asset.status = 'warn';
    else asset.status = 'ok';
  }

  TELEMETRY_BUFFER[asset.id].push(row);
  if (TELEMETRY_BUFFER[asset.id].length > 500) TELEMETRY_BUFFER[asset.id].shift();
  return row;
}

/* ══════════════════════════════════════════
   WEBSOCKET — live telemetry broadcast
══════════════════════════════════════════ */
const clients = new Set();

wss.on('connection', ws => {
  clients.add(ws);
  // send current asset snapshot immediately on connect
  ws.send(JSON.stringify({ type: 'snapshot', assets: ASSETS, alerts: ALERTS }));
  ws.on('close', () => clients.delete(ws));
});

function broadcast(msg) {
  const str = JSON.stringify(msg);
  clients.forEach(ws => { if (ws.readyState === WebSocket.OPEN) ws.send(str); });
}

// Tick every 800ms — generate telemetry for all assets and broadcast
setInterval(() => {
  const readings = ASSETS.map(a => generateTelemetry(a));
  broadcast({ type: 'telemetry', data: readings, ts: new Date().toISOString() });
}, 800);

// Every 15s broadcast an alert if any asset drops below threshold
setInterval(() => {
  ASSETS.forEach(a => {
    if (a.status === 'crit' && !ALERTS.find(al => al.asset === a.id && !al.ack && al.type === 'crit')) {
      const al = { id: 'ALT-' + Date.now(), type:'crit', msg:`${a.id} health critical — ${a.health.toFixed(0)}% — immediate action required`, asset:a.id, ack:false, ts:new Date().toISOString() };
      ALERTS.unshift(al);
      if (ALERTS.length > 50) ALERTS.pop();
      broadcast({ type:'alert', alert: al });
    }
  });
}, 15000);

/* ══════════════════════════════════════════
   REST API — ASSETS
══════════════════════════════════════════ */
app.get('/api/assets', (req, res) => res.json(ASSETS));

app.post('/api/assets', (req, res) => {
  const { id, name, type, zone, mfr, installedDate } = req.body;
  if (!id || !name || !type) return res.status(400).json({ error: 'id, name, type required' });
  if (ASSETS.find(a => a.id === id)) return res.status(409).json({ error: 'Asset ID already exists' });
  const asset = { id, name, type, zone: zone||'A3', mfr: mfr||'Unknown', installedDate: installedDate||new Date().toISOString().slice(0,10), rul:90, health:100, status:'ok' };
  ASSETS.push(asset);
  TELEMETRY_BUFFER[id] = [];
  broadcast({ type:'asset_added', asset });
  res.status(201).json(asset);
});

app.patch('/api/assets/:id', (req, res) => {
  const a = ASSETS.find(x => x.id === req.params.id);
  if (!a) return res.status(404).json({ error: 'Not found' });
  Object.assign(a, req.body);
  broadcast({ type:'asset_updated', asset: a });
  res.json(a);
});

app.delete('/api/assets/:id', (req, res) => {
  const idx = ASSETS.findIndex(x => x.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Not found' });
  ASSETS.splice(idx, 1);
  delete TELEMETRY_BUFFER[req.params.id];
  broadcast({ type:'asset_removed', assetId: req.params.id });
  res.json({ ok: true });
});

/* ══════════════════════════════════════════
   REST API — TELEMETRY
══════════════════════════════════════════ */
app.get('/api/telemetry', (req, res) => {
  // Return last N readings for all assets (or specific asset)
  const { assetId, limit = 100 } = req.query;
  if (assetId) {
    const buf = TELEMETRY_BUFFER[assetId] || [];
    return res.json(buf.slice(-+limit));
  }
  // All assets — return last 50 per asset
  const result = {};
  Object.keys(TELEMETRY_BUFFER).forEach(id => { result[id] = TELEMETRY_BUFFER[id].slice(-50); });
  res.json(result);
});

/* ══════════════════════════════════════════
   REST API — ALERTS
══════════════════════════════════════════ */
app.get('/api/alerts', (req, res) => res.json(ALERTS));

app.patch('/api/alerts/:id', (req, res) => {
  const al = ALERTS.find(x => x.id === req.params.id);
  if (!al) return res.status(404).json({ error: 'Not found' });
  Object.assign(al, req.body);
  broadcast({ type:'alert_updated', alert: al });
  res.json(al);
});

app.post('/api/alerts/ack-all', (req, res) => {
  ALERTS.forEach(a => a.ack = true);
  broadcast({ type:'alerts_acked_all' });
  res.json({ ok: true });
});

/* ══════════════════════════════════════════
   REST API — WORK ORDERS
══════════════════════════════════════════ */
app.get('/api/workorders', (req, res) => res.json(WORK_ORDERS));

app.post('/api/workorders', (req, res) => {
  const wo = { id: 'WO-' + Date.now(), progress:0, created: new Date().toISOString().slice(0,10), ...req.body };
  WORK_ORDERS.unshift(wo);
  broadcast({ type:'wo_added', wo });
  res.status(201).json(wo);
});

app.patch('/api/workorders/:id', (req, res) => {
  const wo = WORK_ORDERS.find(x => x.id === req.params.id);
  if (!wo) return res.status(404).json({ error: 'Not found' });
  Object.assign(wo, req.body);
  broadcast({ type:'wo_updated', wo });
  res.json(wo);
});

app.delete('/api/workorders/:id', (req, res) => {
  const idx = WORK_ORDERS.findIndex(x => x.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Not found' });
  WORK_ORDERS.splice(idx, 1);
  res.json({ ok: true });
});

/* ══════════════════════════════════════════
   REST API — MAINTENANCE PLAN
══════════════════════════════════════════ */
app.get('/api/maintplan', (req, res) => res.json(MAINT_PLAN));

app.post('/api/maintplan', (req, res) => {
  MAINT_PLAN.push(req.body);
  MAINT_PLAN.sort((a,b) => a.date.localeCompare(b.date));
  res.status(201).json(req.body);
});

/* ══════════════════════════════════════════
   EXCEL EXPORT ENGINE
   Supports: assets, telemetry, alerts, workorders, maintplan, full
══════════════════════════════════════════ */
function buildAssetSheet() {
  return ASSETS.map(a => ({
    'Asset ID':       a.id,
    'Name':           a.name,
    'Type':           a.type,
    'Zone':           a.zone,
    'Manufacturer':   a.mfr,
    'Install Date':   a.installedDate,
    'Health (%)':     a.health,
    'RUL (days)':     a.rul,
    'Status':         a.status,
  }));
}

function buildTelemetrySheet(assetId) {
  const buf = assetId ? (TELEMETRY_BUFFER[assetId]||[]) : Object.values(TELEMETRY_BUFFER).flat();
  return buf.slice(-500).sort((a,b) => a.timestamp.localeCompare(b.timestamp)).map(r => ({
    'Timestamp':         r.timestamp,
    'Asset ID':          r.assetId,
    'Asset Name':        r.assetName,
    'Status':            r.status,
    'Health (%)':        r.health,
    'RUL (days)':        r.rul,
    'Vibration (g)':     r.vibration ?? 'N/A',
    'Temperature (°C)':  r.temperature,
    'Acoustic (dB)':     r.acoustic,
    'Current (A)':       r.current ?? 'N/A',
    'Flow (L/m)':        r.flow ?? 'N/A',
    'OEE (%)':           r.oee,
  }));
}

function buildAlertsSheet() {
  return ALERTS.map(a => ({
    'Alert ID':   a.id,
    'Type':       a.type,
    'Asset':      a.asset,
    'Message':    a.msg,
    'Timestamp':  a.ts,
    'Acknowledged': a.ack ? 'Yes' : 'No',
  }));
}

function buildWOSheet() {
  return WORK_ORDERS.map(w => ({
    'WO#':         w.id,
    'Asset':       w.asset,
    'Task':        w.title,
    'Status':      w.status,
    'Priority':    w.priority,
    'Technician':  w.tech,
    'Due Date':    w.due,
    'Created':     w.created,
    'Progress (%)':w.progress,
  }));
}

function buildMaintSheet() {
  return MAINT_PLAN.map(m => ({
    'Date':       m.date,
    'Asset':      m.asset,
    'Task':       m.task,
    'Type':       m.type,
    'Duration':   m.dur,
    'Technician': m.tech,
    'Parts':      m.parts,
  }));
}

function styleSheet(ws, headerColor = '0C1A2E') {
  // Set column widths
  const cols = Object.keys(ws).filter(k => k !== '!ref' && k !== '!cols' && k !== '!merges' && k[1] === '1');
  ws['!cols'] = cols.map(() => ({ wch: 22 }));
  return ws;
}

function buildWorkbook(sheets) {
  const wb = XLSX.utils.book_new();
  wb.Props = {
    Title:   'NexusTwin Export',
    Subject: 'Industry 6.0 Predictive Maintenance',
    Author:  'NexusTwin System',
    CreatedDate: new Date(),
  };

  sheets.forEach(({ name, data }) => {
    if (!data || data.length === 0) {
      data = [{ 'No data available': 'No records found' }];
    }
    const ws = XLSX.utils.json_to_sheet(data);
    styleSheet(ws);
    XLSX.utils.book_append_sheet(wb, ws, name);
  });
  return wb;
}

// Export endpoint
app.get('/api/export', (req, res) => {
  const { type = 'full', assetId, fmt = 'xlsx' } = req.query;

  let sheets = [];

  if (type === 'assets' || type === 'full') {
    sheets.push({ name: 'Assets', data: buildAssetSheet() });
  }
  if (type === 'telemetry' || type === 'full') {
    sheets.push({ name: 'Live Telemetry', data: buildTelemetrySheet(assetId) });
  }
  if (type === 'alerts' || type === 'full') {
    sheets.push({ name: 'Alerts', data: buildAlertsSheet() });
  }
  if (type === 'workorders' || type === 'full') {
    sheets.push({ name: 'Work Orders', data: buildWOSheet() });
  }
  if (type === 'maintplan' || type === 'full') {
    sheets.push({ name: 'Maintenance Plan', data: buildMaintSheet() });
  }
  if (type === 'rul') {
    const rulData = ASSETS.map(a => ({
      'Asset ID':    a.id,
      'Name':        a.name,
      'Type':        a.type,
      'RUL (days)':  a.rul,
      'Health (%)':  a.health,
      'Status':      a.status,
      'Predicted Fail Date': new Date(Date.now() + a.rul*86400000).toISOString().slice(0,10),
      'Confidence (%)': Math.round(80 + Math.random()*15),
      'Model':       'LSTM + TFT',
    }));
    sheets.push({ name: 'RUL Forecasts', data: rulData });
  }

  const wb  = buildWorkbook(sheets);
  const now = new Date().toISOString().replace(/[:.]/g,'-').slice(0,19);
  const filename = `NexusTwin_${type}_${now}.xlsx`;

  const buf = XLSX.write(wb, { bookType:'xlsx', type:'buffer' });
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Length', buf.length);
  res.send(buf);
});

/* ══════════════════════════════════════════
   ANALYTICS ENDPOINT
══════════════════════════════════════════ */
app.get('/api/analytics', (req, res) => {
  const total  = ASSETS.length;
  const crit   = ASSETS.filter(a=>a.status==='crit').length;
  const warn   = ASSETS.filter(a=>a.status==='warn').length;
  const ok     = ASSETS.filter(a=>a.status==='ok').length;
  const avgH   = +(ASSETS.reduce((s,a)=>s+a.health,0)/total).toFixed(1);
  const avgRUL = +(ASSETS.reduce((s,a)=>s+a.rul,0)/total).toFixed(1);
  const activeAlerts = ALERTS.filter(a=>!a.ack).length;
  const openWO = WORK_ORDERS.filter(w=>w.status==='open').length;
  res.json({ total, crit, warn, ok, avgHealth:avgH, avgRUL, activeAlerts, openWO, oee: +(90+Math.random()*3).toFixed(2), ts: new Date().toISOString() });
});

/* ══════════════════════════════════════════
   START
══════════════════════════════════════════ */
const PORT = process.env.PORT || 3001;
server.listen(PORT, () => {
  console.log(`\n✅  NexusTwin backend running on http://localhost:${PORT}`);
  console.log(`🔌  WebSocket live feed:  ws://localhost:${PORT}`);
  console.log(`📊  Export endpoint:      http://localhost:${PORT}/api/export?type=full\n`);
  console.log(`  Endpoints:`);
  console.log(`  GET  /api/assets            — list assets`);
  console.log(`  POST /api/assets            — add asset`);
  console.log(`  GET  /api/telemetry         — all telemetry buffers`);
  console.log(`  GET  /api/alerts            — all alerts`);
  console.log(`  GET  /api/workorders        — work orders`);
  console.log(`  GET  /api/maintplan         — maintenance plan`);
  console.log(`  GET  /api/analytics         — KPI summary`);
  console.log(`  GET  /api/export?type=full  — download Excel (.xlsx)\n`);
});