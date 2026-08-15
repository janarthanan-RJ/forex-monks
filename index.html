<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>FOREX MONKS</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root{
    --ink:#0A1420;
    --surface:#111D2B;
    --surface2:#182739;
    --border:#233247;
    --text:#E7EDF3;
    --muted:#7C8FA3;
    --gain:#34D399;
    --loss:#F4685E;
    --gold:#E7A93D;
  }
  *{box-sizing:border-box;}
  body{
    margin:0;
    background:var(--ink);
    color:var(--text);
    font-family:'Inter',sans-serif;
    min-height:100vh;
  }
  .wrap{max-width:1180px;margin:0 auto;padding:32px 24px 80px;}
  header{display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:6px;}
  h1{
    font-family:'Space Grotesk',sans-serif;
    font-weight:700;
    font-size:28px;
    margin:0;
    letter-spacing:-0.02em;
  }
  .sub{color:var(--muted);font-size:13px;margin-top:4px;}
  /* ticker tape - signature element */
  .ticker{
    margin:22px 0 28px;
    overflow:hidden;
    border-top:1px solid var(--border);
    border-bottom:1px solid var(--border);
    background:var(--surface);
    white-space:nowrap;
    position:relative;
  }
  .ticker-track{
    display:inline-flex;
    animation:scroll 30s linear infinite;
    padding:10px 0;
  }
  .ticker.paused .ticker-track{animation-play-state:paused;}
  @keyframes scroll{
    from{transform:translateX(0);}
    to{transform:translateX(-50%);}
  }
  .tick{
    font-family:'IBM Plex Mono',monospace;
    font-size:13px;
    padding:0 22px;
    border-right:1px solid var(--border);
    display:inline-flex;
    gap:8px;
    align-items:center;
  }
  .tick .sym{color:var(--text);font-weight:500;}
  .tick .pl.up{color:var(--gain);}
  .tick .pl.down{color:var(--loss);}
  .empty-ticker{padding:14px 22px;color:var(--muted);font-size:13px;font-family:'IBM Plex Mono',monospace;}

  .stats{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:28px;}
  .stat{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px 18px;}
  .stat .label{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px;}
  .stat .value{font-family:'IBM Plex Mono',monospace;font-size:22px;font-weight:500;}
  .stat .value.up{color:var(--gain);}
  .stat .value.down{color:var(--loss);}

  .actions{display:flex;gap:10px;margin-bottom:18px;flex-wrap:wrap;}
  button{
    font-family:'Inter',sans-serif;
    font-size:13px;
    font-weight:500;
    border-radius:8px;
    padding:9px 16px;
    cursor:pointer;
    border:1px solid var(--border);
    background:var(--surface2);
    color:var(--text);
    transition:background .15s, border-color .15s;
  }
  button:hover{border-color:#354a63;}
  button.primary{background:var(--gold);border-color:var(--gold);color:#241a05;font-weight:600;}
  button.primary:hover{background:#f0b64f;}
  button.ghost{background:transparent;}
  input[type=file]{display:none;}

  table{width:100%;border-collapse:collapse;background:var(--surface);border:1px solid var(--border);border-radius:10px;overflow:hidden;}
  thead th{
    text-align:left;
    font-size:11px;
    text-transform:uppercase;
    letter-spacing:.06em;
    color:var(--muted);
    padding:12px 14px;
    border-bottom:1px solid var(--border);
    background:var(--surface2);
  }
  tbody td{
    padding:11px 14px;
    font-size:13px;
    border-bottom:1px solid var(--border);
    font-family:'IBM Plex Mono',monospace;
  }
  tbody td.cause,tbody td.notes{font-family:'Inter',sans-serif;}
  tbody tr:last-child td{border-bottom:none;}
  tbody tr:hover{background:rgba(255,255,255,0.02);}
  .pl-cell.up{color:var(--gain);}
  .pl-cell.down{color:var(--loss);}
  .side{font-size:11px;padding:2px 8px;border-radius:5px;font-family:'Inter',sans-serif;font-weight:600;}
  .side.long{background:rgba(52,211,153,.12);color:var(--gain);}
  .side.short{background:rgba(244,104,94,.12);color:var(--loss);}
  .row-actions button{padding:4px 9px;font-size:12px;}
  .empty-state{padding:60px 20px;text-align:center;color:var(--muted);}
  .empty-state .big{font-size:15px;color:var(--text);margin-bottom:6px;}

  /* modal */
  .overlay{position:fixed;inset:0;background:rgba(4,9,15,.6);display:none;align-items:center;justify-content:center;z-index:20;}
  .overlay.open{display:flex;}
  .modal{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:26px;width:520px;max-width:92vw;max-height:88vh;overflow-y:auto;}
  .modal h2{font-family:'Space Grotesk',sans-serif;font-size:18px;margin:0 0 18px;}
  .field{margin-bottom:14px;}
  .field label{display:block;font-size:12px;color:var(--muted);margin-bottom:6px;}
  .field input,.field select,.field textarea{
    width:100%;background:var(--ink);border:1px solid var(--border);border-radius:7px;
    color:var(--text);padding:9px 11px;font-family:'Inter',sans-serif;font-size:13px;
  }
  .field textarea{resize:vertical;min-height:56px;}
  .grid2{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
  .modal-actions{display:flex;justify-content:flex-end;gap:10px;margin-top:6px;}
  .err{color:var(--loss);font-size:12px;margin-top:6px;display:none;}
  .err.show{display:block;}

  .tabs{display:flex;gap:4px;margin-bottom:18px;border-bottom:1px solid var(--border);}
  .tab-btn{
    background:transparent;border:none;border-radius:0;
    padding:10px 16px;font-size:13px;color:var(--muted);
    border-bottom:2px solid transparent;margin-bottom:-1px;
  }
  .tab-btn:hover{color:var(--text);border-color:var(--border);}
  .tab-btn.active{color:var(--gold);border-bottom-color:var(--gold);font-weight:600;}

  .value.up{color:var(--gain);}
  .value.down{color:var(--loss);}

  .motivation{
    background:var(--surface);border:1px solid var(--border);border-left:3px solid var(--gold);
    border-radius:8px;padding:14px 18px;font-size:14px;color:var(--text);
    font-family:'Space Grotesk',sans-serif;font-weight:500;
  }

  .panel-card{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px 20px;}
  .panel-card-title{font-family:'Space Grotesk',sans-serif;font-size:14px;font-weight:600;}
  .fee-note{font-size:13px;color:var(--muted);line-height:1.6;margin-top:8px;}
  .fee-note strong{color:var(--text);}

  .toast{
    position:fixed;bottom:24px;left:50%;transform:translateX(-50%) translateY(20px);
    background:var(--surface2);border:1px solid var(--border);border-radius:8px;
    padding:12px 20px;font-size:13px;color:var(--text);
    opacity:0;pointer-events:none;transition:opacity .2s, transform .2s;z-index:30;
  }
  .toast.show{opacity:1;transform:translateX(-50%) translateY(0);}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div>
      <h1>FOREX MONKS</h1>
      <div class="sub" id="subline">No trades logged yet</div>
    </div>
  </header>

  <div class="ticker" id="ticker"><div class="ticker-track" id="tickerTrack"></div></div>

  <div class="stats">
    <div class="stat"><div class="label">Total P/L</div><div class="value" id="statPL">$0.00</div></div>
    <div class="stat"><div class="label">Win rate</div><div class="value" id="statWin">—</div></div>
    <div class="stat"><div class="label">Trades logged</div><div class="value" id="statCount">0</div></div>
    <div class="stat"><div class="label">Profit factor</div><div class="value" id="statPF">—</div></div>
  </div>

  <div class="tabs">
    <button class="tab-btn active" data-tab="trades">Trades</button>
    <button class="tab-btn" data-tab="performance">Performance overview</button>
    <button class="tab-btn" data-tab="commission">Commission</button>
  </div>

  <div class="tab-panel" id="tab-trades">
    <div class="actions">
      <button class="primary" id="addBtn">+ New trade</button>
      <button id="importBtn">Import Excel</button>
      <input type="file" id="fileInput" accept=".xlsx,.xls">
      <button id="exportBtn">Export Excel</button>
      <button class="ghost" id="clearBtn">Clear all</button>
    </div>

    <table>
      <thead>
        <tr>
          <th>EM</th><th>PAIR</th><th>L/S</th><th>ENTRY</th><th>EXIT</th><th>LOTS</th><th>PIPS</th><th>$</th><th>REASON</th><th></th>
        </tr>
      </thead>
      <tbody id="tbody"></tbody>
    </table>
    <div class="empty-state" id="emptyState" style="display:none;">
      <div class="big">No trades yet</div>
      <div>Log your first trade or import an existing Excel sheet to get started.</div>
    </div>
  </div>

  <div class="tab-panel" id="tab-performance" style="display:none;">
    <div class="motivation" id="motivationBanner">Log a trade to see your performance overview.</div>

    <div class="stats" style="margin-top:18px;">
      <div class="stat"><div class="label">Winning trades</div><div class="value up" id="statWins">0</div></div>
      <div class="stat"><div class="label">Losing trades</div><div class="value down" id="statLosses">0</div></div>
      <div class="stat"><div class="label">Best trade</div><div class="value up" id="statBest">—</div></div>
      <div class="stat"><div class="label">Worst trade</div><div class="value down" id="statWorst">—</div></div>
    </div>
  </div>

  <div class="tab-panel" id="tab-commission" style="display:none;">
    <div class="panel-card">
      <div class="panel-card-title">Commission structure (Goat Funded Trader)</div>
      <div class="fee-note">Forex pairs and gold: <strong>lot size × $5.00</strong> round-turn (MT5 splits this into $2.50 to open and $2.50 to close). Crypto, indices and commodities trade commission-free on the raw spread model — you only pay the spread.</div>
    </div>

    <div class="stats" style="margin-top:18px;">
      <div class="stat"><div class="label">Total commission</div><div class="value down" id="statCommission">$0.00</div></div>
      <div class="stat"><div class="label">Net P/L after fees</div><div class="value" id="statNetPL">$0.00</div></div>
    </div>

    <table style="margin-top:18px;">
      <thead>
        <tr><th>Pair</th><th>Type</th><th>Trades</th><th>Total lots</th><th>Commission</th></tr>
      </thead>
      <tbody id="commissionBody"></tbody>
    </table>
    <div class="empty-state" id="commissionEmpty" style="display:none;">
      <div class="big">No trades yet</div>
      <div>Commission breakdown appears once you log trades.</div>
    </div>
  </div>
</div>

<div class="overlay" id="overlay">
  <div class="modal">
    <h2 id="modalTitle">New trade</h2>
    <div class="grid2">
      <div class="field"><label>EM</label><input type="text" id="fDate" placeholder="e.g. 12 Aug, London AM, Asia session"></div>
      <div class="field"><label>Symbol</label><input type="text" id="fSymbol" placeholder="AAPL"></div>
    </div>
    <div class="grid2">
      <div class="field"><label>Side</label>
        <select id="fSide"><option value="Long">Long</option><option value="Short">Short</option></select>
      </div>
      <div class="field"><label>Quantity</label><input type="number" id="fQty" step="any" placeholder="100"></div>
    </div>
    <div class="grid2">
      <div class="field"><label>Entry price</label><input type="number" id="fEntry" step="any" placeholder="150.00"></div>
      <div class="field"><label>Exit price</label><input type="number" id="fExit" step="any" placeholder="155.00"></div>
    </div>
    <div class="field"><label>Cause / reason</label>
      <input type="text" id="fCause" list="causeSuggestions" placeholder="e.g. OB, FVG + BOS, CHoCH, news trade...">
      <datalist id="causeSuggestions">
        <option value="OB"></option><option value="FVG + BOS"></option><option value="CHoCH"></option>
        <option value="News trade"></option><option value="Breakout"></option><option value="Reversal"></option>
        <option value="FOMO"></option><option value="Revenge trade"></option>
      </datalist>
    </div>
    <div class="field"><label>Notes</label><textarea id="fNotes" placeholder="Any extra context..."></textarea></div>
    <div class="err" id="formErr">Fill in pair, lots, entry and exit price.</div>
    <div class="modal-actions">
      <button class="ghost" id="cancelBtn">Cancel</button>
      <button class="primary" id="saveBtn">Save trade</button>
    </div>
  </div>
</div>

<div class="overlay" id="confirmOverlay">
  <div class="modal" style="width:380px;">
    <h2 id="confirmTitle">Are you sure?</h2>
    <p id="confirmMessage" style="font-size:13px;color:var(--muted);margin:0 0 20px;"></p>
    <div class="modal-actions">
      <button class="ghost" id="confirmCancel">Cancel</button>
      <button class="primary" id="confirmOk" style="background:var(--loss);border-color:var(--loss);color:#fff;">Delete</button>
    </div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
let trades = [];
let editingId = null;
const STORAGE_KEY = 'trades';

function uid(){ return Date.now().toString(36) + Math.random().toString(36).slice(2,7); }

let toastTimer = null;
function showToast(msg){
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(()=> el.classList.remove('show'), 3200);
}

function customConfirm(message, title){
  return new Promise((resolve)=>{
    document.getElementById('confirmTitle').textContent = title || 'Are you sure?';
    document.getElementById('confirmMessage').textContent = message;
    const overlay = document.getElementById('confirmOverlay');
    overlay.classList.add('open');
    const okBtn = document.getElementById('confirmOk');
    const cancelBtn = document.getElementById('confirmCancel');
    function cleanup(result){
      overlay.classList.remove('open');
      okBtn.removeEventListener('click', onOk);
      cancelBtn.removeEventListener('click', onCancel);
      overlay.removeEventListener('click', onOverlay);
      resolve(result);
    }
    function onOk(){ cleanup(true); }
    function onCancel(){ cleanup(false); }
    function onOverlay(e){ if(e.target === overlay) cleanup(false); }
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', onCancel);
    overlay.addEventListener('click', onOverlay);
  });
}

// Classifies an instrument and returns its pip size + contract size,
// following standard industry conventions:
//   Forex majors (EUR/USD etc): pip = 0.0001, contract = 100,000 units, quote = USD -> $10/pip/lot fixed
//   Forex JPY pairs: pip = 0.01, contract = 100,000 units
//   Forex crosses where USD is NOT the quote currency (USD/CAD, EUR/JPY): pip value is price-dependent
//   Gold (XAU/USD): pip = 0.10, contract = 100 oz -> $10/pip/lot fixed
//   Crypto (BTC/USD etc): pip = 1.00 (whole dollar), contract = 1 unit -> $1/pip/lot fixed
function classifyInstrument(symbolRaw){
  const s = (symbolRaw||'').toUpperCase().replace(/[^A-Z]/g,'');
  if(s.includes('XAU') || s.includes('GOLD')) return {type:'metal', pipSize:0.10, contract:100};
  const cryptoTickers = ['BTC','ETH','LTC','XRP','SOL','DOGE','ADA','BNB','AVAX','DOT'];
  if(cryptoTickers.some(c=>s.includes(c))) return {type:'crypto', pipSize:1.00, contract:1};
  if(s.length>=6){
    const quote = s.slice(-3);
    if(quote === 'JPY') return {type:'fx', pipSize:0.01, contract:100000, quote};
    return {type:'fx', pipSize:0.0001, contract:100000, quote};
  }
  return {type:'fx', pipSize:0.0001, contract:100000, quote:'USD'};
}

function pipSizeFor(symbol){
  return classifyInstrument(symbol).pipSize;
}

// Dollar value of 1 pip at 1.00 standard lot for this trade.
// Fixed $10/lot for gold and USD-quoted forex majors; fixed $1/lot for crypto;
// price-dependent for crosses where the quote currency isn't USD (uses the trade's exit price as the reference rate).
function pipValuePerLot(t){
  const cfg = classifyInstrument(t.symbol);
  if(cfg.type === 'metal' || cfg.type === 'crypto') return cfg.contract * cfg.pipSize;
  if(cfg.quote === 'USD') return cfg.contract * cfg.pipSize;
  const refPrice = t.exit || t.entry || 1;
  return (cfg.contract * cfg.pipSize) / refPrice;
}

function calcPips(t){
  const dir = t.side === 'Short' ? -1 : 1;
  const priceDiff = (t.exit - t.entry) * dir;
  return priceDiff / pipSizeFor(t.symbol);
}

function calcPL(t){
  return calcPips(t) * t.qty * pipValuePerLot(t);
}

// Goat Funded Trader fee structure: forex pairs & gold are charged a flat
// $5.00 per lot round-turn (MT5 splits it $2.50 open / $2.50 close).
// Crypto, indices and commodities trade commission-free (raw spread model).
function calcCommission(t){
  const cfg = classifyInstrument(t.symbol);
  if(cfg.type === 'fx' || cfg.type === 'metal') return t.qty * 5.00;
  return 0;
}

function fmtMoney(n){
  const sign = n < 0 ? '-' : '';
  return sign + '$' + Math.abs(n).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
}

async function loadTrades(){
  try{
    const res = await window.storage.get('trades', false);
    trades = res ? JSON.parse(res.value) : [];
  }catch(e){ trades = []; }
  render();
}

async function saveTrades(){
  try{ await window.storage.set('trades', JSON.stringify(trades), false); }
  catch(e){ console.error('Storage error', e); }
}

function render(){
  const tbody = document.getElementById('tbody');
  const empty = document.getElementById('emptyState');
  tbody.innerHTML = '';
  if(trades.length === 0){
    empty.style.display = 'block';
    document.querySelector('table').style.display = 'none';
  } else {
    empty.style.display = 'none';
    document.querySelector('table').style.display = 'table';
  }
  const sorted = [...trades].sort((a,b)=> (b.createdAt||0) - (a.createdAt||0));
  for(const t of sorted){
    const pl = calcPL(t);
    const pips = calcPips(t);
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${t.date || ''}</td>
      <td>${escapeHtml(t.symbol)}</td>
      <td><span class="side ${t.side==='Short'?'short':'long'}">${t.side}</span></td>
      <td>${Number(t.entry).toFixed(t.symbol && pipSizeFor(t.symbol)>=0.1 ? 2 : 5)}</td>
      <td>${Number(t.exit).toFixed(t.symbol && pipSizeFor(t.symbol)>=0.1 ? 2 : 5)}</td>
      <td>${t.qty}</td>
      <td class="pl-cell ${pips>=0?'up':'down'}">${pips>=0?'+':''}${pips.toFixed(1)}</td>
      <td class="pl-cell ${pl>=0?'up':'down'}">${fmtMoney(pl)}</td>
      <td class="cause">${escapeHtml(t.cause||'')}</td>
      <td class="row-actions">
        <button data-edit="${t.id}">Edit</button>
        <button data-del="${t.id}">Delete</button>
      </td>`;
    tbody.appendChild(tr);
  }
  tbody.querySelectorAll('[data-edit]').forEach(b=>b.addEventListener('click', ()=>openModal(b.dataset.edit)));
  tbody.querySelectorAll('[data-del]').forEach(b=>b.addEventListener('click', ()=>deleteTrade(b.dataset.del)));

  renderStats();
  renderTicker();
  renderPerformance();
  renderCommission();
}

function escapeHtml(s){
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function renderStats(){
  const sub = document.getElementById('subline');
  sub.textContent = trades.length ? `${trades.length} trade${trades.length===1?'':'s'} logged` : 'No trades logged yet';

  const pls = trades.map(calcPL);
  const total = pls.reduce((a,b)=>a+b,0);
  const wins = pls.filter(p=>p>0);
  const losses = pls.filter(p=>p<0);
  const winRate = trades.length ? (wins.length/trades.length*100) : null;
  const grossWin = wins.reduce((a,b)=>a+b,0);
  const grossLoss = Math.abs(losses.reduce((a,b)=>a+b,0));
  const pf = grossLoss > 0 ? (grossWin/grossLoss) : (grossWin>0 ? Infinity : null);

  const plEl = document.getElementById('statPL');
  plEl.textContent = fmtMoney(total);
  plEl.className = 'value ' + (total>=0?'up':'down');

  document.getElementById('statWin').textContent = winRate===null ? '—' : winRate.toFixed(0)+'%';
  document.getElementById('statCount').textContent = trades.length;
  document.getElementById('statPF').textContent = pf===null ? '—' : (pf===Infinity ? '∞' : pf.toFixed(2));
}

function renderTicker(){
  const track = document.getElementById('tickerTrack');
  const ticker = document.getElementById('ticker');
  if(trades.length === 0){
    track.innerHTML = '<div class="empty-ticker">Log a trade to see it appear here</div>';
    ticker.classList.add('paused');
    return;
  }
  ticker.classList.remove('paused');
  const recent = [...trades].sort((a,b)=> (b.createdAt||0) - (a.createdAt||0)).slice(0,15);
  const build = () => recent.map(t=>{
    const pl = calcPL(t);
    const pips = calcPips(t);
    return `<span class="tick"><span class="sym">${escapeHtml(t.symbol)}</span><span class="pl ${pl>=0?'up':'down'}">${pips>=0?'+':''}${pips.toFixed(1)}p / ${fmtMoney(pl)}</span></span>`;
  }).join('');
  track.innerHTML = build() + build();
}

const WIN_QUOTES = [
  "Green day. Discipline is compounding.",
  "Consistency beats intensity — nice work staying in your edge.",
  "That's the process working. Keep the risk tight and let it run.",
  "You're proving the system works. Protect this streak with good risk management."
];
const LOSS_QUOTES = [
  "Red days happen to every trader — the edge shows up over a sample, not one trade.",
  "Down today, but the plan didn't break. Review the reason and refine the setup.",
  "Losses are tuition. What did this one teach you?",
  "One bad stretch doesn't undo good process. Stay disciplined, size stays the same."
];
const NEUTRAL_QUOTES = [
  "Log your first trade to start tracking your edge.",
  "Flat so far. Patience is a position too."
];

function renderPerformance(){
  const pls = trades.map(t => calcPL(t) - calcCommission(t));
  const wins = pls.filter(p=>p>0);
  const losses = pls.filter(p=>p<0);
  document.getElementById('statWins').textContent = wins.length;
  document.getElementById('statLosses').textContent = losses.length;

  let bestEl = document.getElementById('statBest');
  let worstEl = document.getElementById('statWorst');
  if(trades.length === 0){
    bestEl.textContent = '—';
    worstEl.textContent = '—';
  } else {
    let best = trades[0], worst = trades[0];
    for(const t of trades){
      if(calcPL(t) > calcPL(best)) best = t;
      if(calcPL(t) < calcPL(worst)) worst = t;
    }
    bestEl.textContent = `${best.symbol} ${fmtMoney(calcPL(best))}`;
    worstEl.textContent = `${worst.symbol} ${fmtMoney(calcPL(worst))}`;
  }

  const totalNet = pls.reduce((a,b)=>a+b,0);
  const banner = document.getElementById('motivationBanner');
  let pool = NEUTRAL_QUOTES;
  if(trades.length > 0) pool = totalNet >= 0 ? WIN_QUOTES : LOSS_QUOTES;
  banner.textContent = pool[Math.floor(Math.random()*pool.length)];
}

function renderCommission(){
  const tbody = document.getElementById('commissionBody');
  const empty = document.getElementById('commissionEmpty');
  if(trades.length === 0){
    tbody.innerHTML = '';
    empty.style.display = 'block';
    document.querySelectorAll('#tab-commission table')[0].style.display = 'none';
  } else {
    empty.style.display = 'none';
    document.querySelectorAll('#tab-commission table')[0].style.display = 'table';
  }

  const bySymbol = {};
  let totalCommission = 0, totalPL = 0;
  for(const t of trades){
    const c = calcCommission(t);
    totalCommission += c;
    totalPL += calcPL(t);
    const cfg = classifyInstrument(t.symbol);
    if(!bySymbol[t.symbol]) bySymbol[t.symbol] = {trades:0, lots:0, commission:0, type:cfg.type};
    bySymbol[t.symbol].trades += 1;
    bySymbol[t.symbol].lots += t.qty;
    bySymbol[t.symbol].commission += c;
  }

  const typeLabel = {fx:'Forex', metal:'Gold/metal', crypto:'Crypto'};
  tbody.innerHTML = Object.keys(bySymbol).sort().map(sym=>{
    const row = bySymbol[sym];
    return `<tr>
      <td style="font-family:'Inter',sans-serif;">${escapeHtml(sym)}</td>
      <td style="font-family:'Inter',sans-serif;color:var(--muted);">${typeLabel[row.type]||'Other'}</td>
      <td>${row.trades}</td>
      <td>${row.lots.toFixed(2)}</td>
      <td class="pl-cell down">${fmtMoney(row.commission)}</td>
    </tr>`;
  }).join('');

  document.getElementById('statCommission').textContent = fmtMoney(totalCommission);
  const netEl = document.getElementById('statNetPL');
  const net = totalPL - totalCommission;
  netEl.textContent = fmtMoney(net);
  netEl.className = 'value ' + (net>=0?'up':'down');
}

document.querySelectorAll('.tab-btn').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p=>p.style.display='none');
    btn.classList.add('active');
    document.getElementById('tab-'+btn.dataset.tab).style.display = 'block';
  });
});


function openModal(id){
  editingId = id || null;
  document.getElementById('formErr').classList.remove('show');
  document.getElementById('modalTitle').textContent = id ? 'Edit trade' : 'New trade';
  if(id){
    const t = trades.find(x=>x.id===id);
    document.getElementById('fDate').value = t.date || '';
    document.getElementById('fSymbol').value = t.symbol || '';
    document.getElementById('fSide').value = t.side || 'Long';
    document.getElementById('fQty').value = t.qty;
    document.getElementById('fEntry').value = t.entry;
    document.getElementById('fExit').value = t.exit;
    document.getElementById('fCause').value = t.cause || '';
    document.getElementById('fNotes').value = t.notes || '';
  } else {
    document.getElementById('fDate').value = '';
    document.getElementById('fSymbol').value = '';
    document.getElementById('fSide').value = 'Long';
    document.getElementById('fQty').value = '';
    document.getElementById('fEntry').value = '';
    document.getElementById('fExit').value = '';
    document.getElementById('fCause').value = '';
    document.getElementById('fNotes').value = '';
  }
  document.getElementById('overlay').classList.add('open');
}

function closeModal(){
  document.getElementById('overlay').classList.remove('open');
  editingId = null;
}

function saveTrade(){
  const date = document.getElementById('fDate').value.trim();
  const symbol = document.getElementById('fSymbol').value.trim().toUpperCase();
  const side = document.getElementById('fSide').value;
  const qty = parseFloat(document.getElementById('fQty').value);
  const entry = parseFloat(document.getElementById('fEntry').value);
  const exit = parseFloat(document.getElementById('fExit').value);
  const cause = document.getElementById('fCause').value.trim();
  const notes = document.getElementById('fNotes').value.trim();

  if(!symbol || isNaN(qty) || isNaN(entry) || isNaN(exit)){
    document.getElementById('formErr').classList.add('show');
    return;
  }

  if(editingId){
    const t = trades.find(x=>x.id===editingId);
    Object.assign(t, {date,symbol,side,qty,entry,exit,cause,notes});
  } else {
    trades.push({id:uid(),date,symbol,side,qty,entry,exit,cause,notes,createdAt:Date.now()});
  }
  saveTrades();
  closeModal();
  render();
}

function deleteTrade(id){
  trades = trades.filter(t=>t.id!==id);
  saveTrades();
  render();
}

function exportExcel(){
  const rows = trades.map(t=>({
    EM:t.date, PAIR:t.symbol, 'L/S':t.side, ENTRY:t.entry, EXIT:t.exit,
    LOTS:t.qty, PIPS:Number(calcPips(t).toFixed(1)), '$':Number(calcPL(t).toFixed(2)),
    REASON:t.cause
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Trades');
  XLSX.writeFile(wb, 'forex_monks.xlsx');
}

function importExcel(file){
  const reader = new FileReader();
  reader.onload = (e) => {
    try{
      const wb = XLSX.read(e.target.result, {type:'array'});
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, {defval:''});
      let added = 0;
      for(const r of rows){
        const symbol = String(r.PAIR || r.Pair || r.Symbol || r.symbol || '').trim().toUpperCase();
        const entry = parseFloat(r.ENTRY ?? r.Entry ?? r.entry);
        const exit = parseFloat(r.EXIT ?? r.Exit ?? r.exit);
        const qty = parseFloat(r.LOTS ?? r.Lots ?? r.Quantity ?? r.Qty ?? r.qty);
        if(!symbol || isNaN(entry) || isNaN(exit) || isNaN(qty)) continue;
        let date = r.EM ?? r.Date ?? r.date ?? '';
        date = (date instanceof Date) ? date.toISOString().slice(0,10) : String(date);
        const sideRaw = String(r['L/S'] || r.Side || r.side || 'Long').trim().toUpperCase();
        const side = (sideRaw === 'S' || sideRaw === 'SHORT') ? 'Short' : 'Long';
        trades.push({
          id:uid(), date, symbol, side, qty, entry, exit,
          cause: r.REASON || r.Reason || r.Cause || r.cause || '',
          notes: r.Notes || r.notes || '',
          createdAt: Date.now() - added
        });
        added++;
      }
      saveTrades();
      render();
      showToast(added + ' trade(s) imported.');
    }catch(err){
      showToast('Could not read that file — make sure it is a valid Excel file.');
    }
  };
  reader.readAsArrayBuffer(file);
}

document.getElementById('addBtn').addEventListener('click', ()=>openModal(null));
document.getElementById('cancelBtn').addEventListener('click', closeModal);
document.getElementById('saveBtn').addEventListener('click', saveTrade);
document.getElementById('overlay').addEventListener('click', (e)=>{ if(e.target.id==='overlay') closeModal(); });
document.getElementById('exportBtn').addEventListener('click', exportExcel);
document.getElementById('importBtn').addEventListener('click', ()=>document.getElementById('fileInput').click());
document.getElementById('fileInput').addEventListener('change', (e)=>{
  if(e.target.files[0]) importExcel(e.target.files[0]);
  e.target.value = '';
});
document.getElementById('clearBtn').addEventListener('click', async ()=>{
  const ok = await customConfirm('Delete all logged trades? This cannot be undone.', 'Clear all trades');
  if(ok){
    trades = [];
    saveTrades();
    render();
    showToast('All trades cleared.');
  }
});

loadTrades();
</script>
</body>
</html>
