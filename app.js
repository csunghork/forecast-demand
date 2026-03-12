'use strict';

const DEFAULT_MODELS = [
  {id:'naive', name:'Naive', color:'#95a5a6'},
  {id:'snaive', name:'Seasonal Naive', color:'#7f8c8d'},
  {id:'sma', name:'Simple Moving Average', color:'#16a085'},
  {id:'drift', name:'Drift', color:'#f39c12'},
  {id:'adida', name:'ADIDA', color:'#9b59b6'},
  {id:'imapa', name:'IMAPA', color:'#8e44ad'},
  {id:'ets', name:'ETS Auto', color:'#2e86de'},
  {id:'hw', name:'Holt-Winters', color:'#54a0ff'},
  {id:'des', name:'Double Exp. Smoothing', color:'#2980b9'},
  {id:'ses', name:'Single Exp. Smoothing', color:'#a55eea'},
  {id:'lr', name:'Linear Regression', color:'#fd9644'},
  {id:'wma', name:'Weighted Moving Average', color:'#2bcbba'},
  {id:'seasonal', name:'Seasonal Decomposition', color:'#26de81'},
  {id:'arima', name:'ARIMA Auto', color:'#fc5c65'},
  {id:'sarima', name:'SARIMA Auto', color:'#fed330'},
  {id:'croston', name:'Croston / SBA / TSB', color:'#ff6b81'},
  {id:'theta', name:'Theta', color:'#eccc68'},
  {id:'prophet', name:'Prophet', color:'#1abc9c'},
  {id:'xgb', name:'XGBoost', color:'#d35400'}
];

let models = [...DEFAULT_MODELS];
let canonicalRows = [];
let results = [];
let selectedSku = null;
let lastResponse = null;
let sortKey = 'rank_score';
let sortDir = 1;

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const num = (v, d=1) => v == null || !isFinite(v) ? '—' : Number(v).toFixed(d);
const mean = arr => arr.length ? arr.reduce((a,b)=>a+b,0)/arr.length : 0;
const sum = arr => arr.reduce((a,b)=>a+b,0);

async function init(){
  $('appTitle').textContent = window.APP_CONFIG?.APP_TITLE || 'Forecast Demand';
  $('apiUrl').textContent = window.APP_CONFIG?.API_BASE_URL || 'API not configured';
  renderModelChecklist();
  bindUpload();
  try {
    const res = await fetch(`${window.APP_CONFIG.API_BASE_URL}/api/models`);
    if(res.ok){
      const data = await res.json();
      if(Array.isArray(data.models) && data.models.length){
        models = data.models.map(m => ({id:m.id,name:m.name,color:m.color || '#54a0ff'}));
        renderModelChecklist();
      }
    }
  } catch (e) {
    console.warn('Could not load model list from backend.', e);
  }
}

function bindUpload(){
  $('fileIn').addEventListener('change', e => {
    const file = e.target.files?.[0];
    if(file) processFile(file);
  });
  const dz = $('dropZone');
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
  dz.addEventListener('drop', e => {
    e.preventDefault();
    dz.classList.remove('drag');
    const file = e.dataTransfer.files?.[0];
    if(file) processFile(file);
  });
}

function normalizeHeader(s){
  return String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}
function findHeader(keys, candidates){
  const nk = keys.map(k => ({raw:k, norm:normalizeHeader(k)}));
  for(const c of candidates){
    const cn = normalizeHeader(c);
    const exact = nk.find(x => x.norm === cn);
    if(exact) return exact.raw;
    const partial = nk.find(x => x.norm.includes(cn) || cn.includes(x.norm));
    if(partial) return partial.raw;
  }
  return '';
}
function parseQty(v){
  if(v == null || v === '') return 0;
  if(typeof v === 'number') return isFinite(v) ? v : 0;
  const n = parseFloat(String(v).replace(/,/g,'').trim());
  return isFinite(n) ? n : 0;
}
function parsePostingDate(v){
  if(v == null || v === '') return null;
  if(v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if(typeof v === 'number' && isFinite(v)){
    const dc = XLSX.SSF.parse_date_code(v);
    if(dc) return new Date(dc.y, dc.m - 1, dc.d);
  }
  const s = String(v).trim();
  if(!s) return null;
  const iso = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if(iso) return new Date(+iso[1], +iso[2]-1, +iso[3]);
  const dmy = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
  if(dmy){
    let y = +dmy[3];
    if(y < 100) y += y < 50 ? 2000 : 1900;
    return new Date(y, +dmy[2]-1, +dmy[1]);
  }
  const parsed = new Date(s);
  return isNaN(parsed) ? null : new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}
function weekInMonth4(dt){
  const day = dt.getDate();
  if(day <= 7) return 1;
  if(day <= 14) return 2;
  if(day <= 21) return 3;
  return 4;
}

function setStatus(msg, type='idle'){
  $('statusMsg').textContent = msg;
  $('status').className = `status ${type}`;
}

function renderModelChecklist(){
  $('modelChecklist').innerHTML = models.map(m => `
    <label class="checkItem">
      <input type="checkbox" checked value="${esc(m.id)}" />
      <span class="dot" style="background:${m.color}"></span>
      <span>${esc(m.name)}</span>
    </label>
  `).join('');
}

function getSelectedModels(){
  return [...document.querySelectorAll('#modelChecklist input:checked')].map(el => el.value);
}

function processFile(file){
  setStatus('Reading file…', 'run');
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, {type:'binary', cellDates:true});
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, {defval:null});
      if(!json.length) throw new Error('The file is empty.');

      const keys = Object.keys(json[0]);
      const descKey = findHeader(keys, ['description','desc','productname','itemname','itemdescription']);
      const codeKey = findHeader(keys, ['itemcode','item code','sku','code','item']);
      const postKey = findHeader(keys, ['postingdate','posting date','documentdate','date','invoice date']);
      const yearKey = findHeader(keys, ['year','yr']);
      const monthKey = findHeader(keys, ['month','mon','mth']);
      const weekKey = findHeader(keys, ['week','wk']);
      const qtyKey = findHeader(keys, ['sumofquantity','sum of quantity','quantity','qty','sales','demand','sumqty','sum_qty']);

      if(!codeKey || !qtyKey || !(postKey || (yearKey && monthKey))){
        throw new Error('Need ItemCode + Sum of Quantity + PostingDate or Year/Month.');
      }

      const bucketMap = new Map();
      let accepted = 0;
      let skipped = 0;

      json.forEach(row => {
        const code = String(row[codeKey] ?? '').trim();
        if(!code){ skipped++; return; }
        const desc = String((descKey ? row[descKey] : code) ?? code).trim() || code;
        const qty = parseQty(row[qtyKey]);
        let year, month, week;

        if(postKey){
          const dt = parsePostingDate(row[postKey]);
          if(!dt){ skipped++; return; }
          year = dt.getFullYear();
          month = dt.getMonth() + 1;
          week = weekInMonth4(dt);
        } else {
          year = parseInt(row[yearKey], 10);
          month = parseInt(row[monthKey], 10);
          week = weekKey ? parseInt(row[weekKey], 10) : 1;
          if(!isFinite(year) || !isFinite(month)){ skipped++; return; }
          if(!isFinite(week) || week < 1) week = 1;
          if(week > 4) week = 4;
        }

        const key = `${code}|${year}|${month}|${week}`;
        if(!bucketMap.has(key)) bucketMap.set(key, {code, desc, year, month, week, qty:0});
        bucketMap.get(key).qty += qty;
        accepted++;
      });

      canonicalRows = [...bucketMap.values()]
        .sort((a,b) => a.code.localeCompare(b.code) || a.year - b.year || a.month - b.month || a.week - b.week);

      $('fileInfo').textContent = `${file.name} · ${accepted} rows · ${new Set(canonicalRows.map(r=>r.code)).size} SKUs · ${skipped} skipped`;
      $('loadedCount').textContent = `${canonicalRows.length} canonical period rows ready`;
      $('runBtn').disabled = false;
      $('exportBtn').disabled = true;
      setStatus('File loaded. Ready to forecast.', 'ok');
    } catch (err) {
      console.error(err);
      setStatus(err.message || 'Could not read file.', 'err');
      alert(err.message || 'Could not read file.');
    }
  };
  reader.readAsBinaryString(file);
}

async function runForecast(){
  if(!canonicalRows.length){
    alert('Load a file first.');
    return;
  }
  const api = window.APP_CONFIG?.API_BASE_URL;
  if(!api || api.includes('YOUR-BACKEND-SERVICE')){
    alert('Edit frontend/config.js and set your backend URL first.');
    return;
  }

  const payload = {
    records: canonicalRows,
    settings: {
      horizon_unit: $('horizonUnit').value,
      horizon_n: parseInt($('horizonN').value, 10) || 8,
      lookback_months: parseInt($('lookbackN').value, 10) || 12,
      validation_steps: parseInt($('valN').value, 10) || 4,
      seasonal_period: parseInt($('seasonN').value, 10) || 4,
      models: getSelectedModels(),
      exclude_current_month: $('excludeCurrent').checked,
      timezone: 'Asia/Kuala_Lumpur'
    }
  };

  $('runBtn').disabled = true;
  setStatus('Running backend forecast models…', 'run');
  try {
    const res = await fetch(`${api}/api/forecast`, {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify(payload)
    });
    const data = await res.json();
    if(!res.ok) throw new Error(data.detail || 'Forecast failed.');
    lastResponse = data;
    results = Array.isArray(data.results) ? data.results : [];
    results.sort((a,b) => (a.rank_score ?? 999999) - (b.rank_score ?? 999999));
    selectedSku = results[0]?.code || null;
    renderAll();
    $('exportBtn').disabled = results.length === 0;
    setStatus(`Forecast complete — ${results.length} SKUs processed.`, 'ok');
  } catch (err) {
    console.error(err);
    setStatus(err.message || 'Forecast failed.', 'err');
    alert(err.message || 'Forecast failed.');
  } finally {
    $('runBtn').disabled = false;
  }
}

function renderAll(){
  renderSummary();
  renderSkuList();
  renderOverviewTable();
  renderCompetitionTable();
  renderSkuSelector();
  renderChart();
}

function renderSummary(){
  const s = lastResponse?.summary || {};
  $('kSku').textContent = s.total_skus ?? '—';
  $('kMape').textContent = s.avg_mape == null ? '—' : `${s.avg_mape.toFixed(1)}%`;
  $('kNmae').textContent = s.avg_nmae == null ? '—' : `${s.avg_nmae.toFixed(1)}%`;
  $('kRank').textContent = s.avg_rank_score == null ? '—' : s.avg_rank_score.toFixed(1);
  $('kUp').textContent = s.trending_up ?? 0;
  $('kDn').textContent = s.trending_down ?? 0;
  const total = s.horizon_total ?? 0;
  $('kTotal').textContent = total > 999999 ? `${(total/1e6).toFixed(1)}M` : total > 999 ? `${(total/1000).toFixed(1)}K` : total.toFixed(0);
}

function renderSkuList(){
  const q = $('skuSearch').value.toLowerCase();
  const filtered = results.filter(r => !q || r.code.toLowerCase().includes(q) || String(r.desc).toLowerCase().includes(q));
  $('skuList').innerHTML = filtered.map(r => `
    <button class="skuItem ${selectedSku === r.code ? 'active' : ''}" onclick="pickSku(decodeURIComponent('${encodeURIComponent(r.code)}'))">
      <div class="skuCode">${esc(r.code)}</div>
      <div class="skuDesc">${esc(r.desc)}</div>
      <div class="skuMeta">
        <span>${esc(r.model_name || '—')}</span>
        <span>${r.mape == null ? '—' : r.mape.toFixed(1)+'%'}</span>
      </div>
    </button>
  `).join('') || '<div class="empty">No SKUs yet</div>';
}

function renderSkuSelector(){
  $('skuSelect').innerHTML = '<option value="">Select SKU</option>' + results.map(r => `<option value="${esc(r.code)}" ${selectedSku===r.code?'selected':''}>${esc(r.code)} — ${esc(r.desc.slice(0,28))}</option>`).join('');
}

function pickSku(code){
  selectedSku = code;
  renderSkuList();
  renderSkuSelector();
  renderChart();
}

function getSelectedResult(){
  return results.find(r => r.code === selectedSku) || results[0] || null;
}

function sortResults(rows){
  const copy = [...rows];
  copy.sort((a,b) => {
    let av = a[sortKey], bv = b[sortKey];
    if(sortKey === 'forecast_total'){
      av = sum(a.forecast || []);
      bv = sum(b.forecast || []);
    }
    if(sortKey === 'next4'){
      av = sum((a.forecast || []).slice(0,4));
      bv = sum((b.forecast || []).slice(0,4));
    }
    if(sortKey === 'model_name'){
      av = a.model_name || '';
      bv = b.model_name || '';
    }
    if(sortKey === 'desc'){
      av = a.desc || '';
      bv = b.desc || '';
    }
    if(typeof av === 'string') return sortDir * av.localeCompare(bv);
    av = av == null ? 999999 : av;
    bv = bv == null ? 999999 : bv;
    return sortDir * (av - bv);
  });
  return copy;
}

function setSort(key){
  if(sortKey === key) sortDir *= -1;
  else { sortKey = key; sortDir = 1; }
  renderOverviewTable();
}

function renderOverviewTable(){
  const rows = sortResults(results);
  $('overviewBody').innerHTML = rows.map(r => {
    const trend = r.trend === 'up' ? '<span class="up">↑ UP</span>' : r.trend === 'down' ? '<span class="down">↓ DOWN</span>' : '<span class="flat">→ FLAT</span>';
    return `
      <tr onclick="pickSku(decodeURIComponent('${encodeURIComponent(r.code)}'))" class="${selectedSku === r.code ? 'selected' : ''}">
        <td class="mono strong">${esc(r.code)}</td>
        <td>${esc(r.desc)}</td>
        <td>${esc(r.model_name || '—')}</td>
        <td class="mono">${num(r.rank_score,1)}</td>
        <td class="mono">${num(r.mape,1)}${r.mape==null?'':'%'}</td>
        <td class="mono">${num(r.nmae,1)}${r.nmae==null?'':'%'}</td>
        <td class="mono">${num(r.avg_per_period,1)}</td>
        <td class="mono">${num(r.total_hist,1)}</td>
        <td>${trend}</td>
        <td class="mono">${num(r.outlier_count,0)}</td>
        <td class="mono strong">${num(sum((r.forecast||[]).slice(0,4)),1)}</td>
        <td class="mono strong">${num(sum(r.forecast||[]),1)}</td>
      </tr>
    `;
  }).join('');
}

function renderCompetitionTable(){
  const active = getSelectedModels();
  const head = ['Item Code','Description','Winner','Rank','MAPE%','NMAE%'].concat(active.map(id => `${models.find(m=>m.id===id)?.name || id} Rank`));
  $('competitionHead').innerHTML = '<tr>' + head.map(h => `<th>${esc(h)}</th>`).join('') + '</tr>';
  $('competitionBody').innerHTML = results.map(r => `
    <tr onclick="pickSku(decodeURIComponent('${encodeURIComponent(r.code)}'))">
      <td class="mono strong">${esc(r.code)}</td>
      <td>${esc(r.desc)}</td>
      <td>${esc(r.model_name || '—')}</td>
      <td class="mono">${num(r.rank_score,1)}</td>
      <td class="mono">${num(r.mape,1)}${r.mape==null?'':'%'}</td>
      <td class="mono">${num(r.nmae,1)}${r.nmae==null?'':'%'}</td>
      ${active.map(id => `<td class="mono">${num(r.model_scores?.[id]?.rank_score,1)}</td>`).join('')}
    </tr>
  `).join('');
}

function renderChart(){
  const r = getSelectedResult();
  const wrap = $('chartWrap');
  if(!r){
    wrap.innerHTML = '<div class="empty">Run a forecast and select a SKU.</div>';
    $('skuInfo').innerHTML = '';
    $('modelCompetition').innerHTML = '';
    return;
  }

  wrap.innerHTML = '<canvas id="chartCanvas"></canvas><div id="chartTip" class="chartTip hidden"></div>';
  $('skuInfo').innerHTML = `
    <div class="infoGrid">
      <div><span>Code</span><b>${esc(r.code)}</b></div>
      <div><span>Winner</span><b>${esc(r.model_name || '—')}</b></div>
      <div><span>Periods</span><b>${r.series?.length || 0}</b></div>
      <div><span>Train Steps</span><b>${r.train_n ?? 0}</b></div>
      <div><span>Val Steps</span><b>${r.validation_steps ?? 0}</b></div>
      <div><span>MAPE</span><b>${num(r.mape,2)}${r.mape==null?'':'%'}</b></div>
      <div><span>NMAE</span><b>${num(r.nmae,2)}${r.nmae==null?'':'%'}</b></div>
      <div><span>Outliers</span><b>${r.outlier_count ?? 0}</b></div>
      <div><span>Zero Rate</span><b>${r.pattern?.zero_rate != null ? (r.pattern.zero_rate*100).toFixed(1)+'%' : '—'}</b></div>
      <div><span>Pattern</span><b>${r.pattern?.sparse ? 'SPARSE' : r.pattern?.intermittent ? 'INTERMITTENT' : 'REGULAR'}</b></div>
    </div>
  `;

  const scores = Object.entries(r.model_scores || {})
    .map(([id, v]) => ({id, name:v.name || id, rank:v.rank_score, mape:v.mape}))
    .sort((a,b) => (a.rank ?? 999999) - (b.rank ?? 999999));
  $('modelCompetition').innerHTML = scores.map(s => `
    <div class="scoreRow ${r.winner===s.id?'winner':''}">
      <span>${esc(s.name)}</span>
      <span class="mono">Rank ${num(s.rank,1)} · MAPE ${num(s.mape,1)}${s.mape==null?'':'%'}</span>
    </div>
  `).join('');

  const canvas = $('chartCanvas');
  const tip = $('chartTip');
  const dpr = window.devicePixelRatio || 1;
  const W = wrap.clientWidth;
  const H = Math.max(340, wrap.clientHeight);
  canvas.width = W * dpr;
  canvas.height = H * dpr;
  canvas.style.width = W + 'px';
  canvas.style.height = H + 'px';
  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);

  const hist = r.series || [];
  const histLabels = r.labels || [];
  const valPred = r.model_scores?.[r.winner]?.predictions || [];
  const fc = r.forecast || [];
  const totalPts = hist.length + fc.length;
  const pad = {t:20,r:16,b:50,l:56};
  const cw = W - pad.l - pad.r;
  const ch = H - pad.t - pad.b;
  const all = [...hist, ...valPred, ...fc].filter(v => v != null && isFinite(v));
  const maxV = Math.max(1, ...all) * 1.15;
  const toX = (i) => pad.l + (i / Math.max(1, totalPts - 1)) * cw;
  const toY = (v) => pad.t + ch - (v / maxV) * ch;
  const trainN = r.train_n || Math.max(0, hist.length - (r.validation_steps || 0));

  const drawLine = (pts, color, dash=[], width=2) => {
    const valid = pts.filter(([,y]) => y != null && isFinite(y));
    if(valid.length < 2) return;
    ctx.beginPath();
    ctx.moveTo(valid[0][0], valid[0][1]);
    valid.slice(1).forEach(([x,y]) => ctx.lineTo(x,y));
    ctx.strokeStyle = color;
    ctx.lineWidth = width;
    ctx.setLineDash(dash);
    ctx.stroke();
    ctx.setLineDash([]);
  };
  const drawDot = (x,y,color) => {
    ctx.beginPath();
    ctx.arc(x,y,3.5,0,Math.PI*2);
    ctx.fillStyle = color;
    ctx.fill();
  };

  function paint(hoverIndex=null, mx=0, my=0){
    ctx.clearRect(0,0,W,H);
    for(let g=0; g<=5; g++){
      const y = pad.t + (ch/5) * g;
      ctx.strokeStyle = '#2b3340';
      ctx.beginPath(); ctx.moveTo(pad.l, y); ctx.lineTo(pad.l+cw, y); ctx.stroke();
      ctx.fillStyle = '#7e8aa0';
      ctx.font = '10px Inter, sans-serif';
      ctx.textAlign = 'right';
      ctx.fillText(((maxV)*(1-g/5)).toFixed(0), pad.l-8, y+3);
    }

    if(trainN > 0 && trainN < hist.length){
      const x = toX(trainN-1);
      ctx.strokeStyle = '#ffb84d';
      ctx.setLineDash([4,4]);
      ctx.beginPath(); ctx.moveTo(x,pad.t); ctx.lineTo(x,pad.t+ch); ctx.stroke();
      ctx.setLineDash([]);
    }
    const histPts = hist.map((v,i)=>[toX(i), toY(v)]);
    const valStart = Math.max(0, trainN-1);
    drawLine(histPts.slice(0, trainN), '#2fd3a2', [], 2.2);
    drawLine(histPts.slice(valStart), '#ffb84d', [], 2.2);

    if(valPred.length && trainN > 0){
      const vp = [[toX(trainN-1), toY(hist[trainN-1])], ...valPred.map((v,i)=>[toX(trainN+i), toY(v)])];
      drawLine(vp, '#60a5fa', [], 2.4);
    }
    if(fc.length && hist.length){
      const fp = [[toX(hist.length-1), toY(hist[hist.length-1])], ...fc.map((v,i)=>[toX(hist.length+i), toY(v)])];
      drawLine(fp, '#a78bfa', [6,4], 2.4);
      fc.forEach((v,i) => drawDot(toX(hist.length+i), toY(v), '#a78bfa'));
    }

    const labelStep = Math.max(1, Math.ceil(totalPts / 10));
    for(let i=0; i<totalPts; i+=labelStep){
      const lbl = i < histLabels.length ? histLabels[i] : `F+${i - histLabels.length + 1}`;
      ctx.fillStyle = '#8c98ad';
      ctx.font = '10px Inter, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(lbl, toX(i), H - 18);
    }

    ctx.fillStyle = '#e7edf8';
    ctx.font = '600 13px Inter, sans-serif';
    ctx.textAlign = 'left';
    ctx.fillText(`${r.code} · ${r.desc}`, pad.l, 14);

    if(hoverIndex != null){
      const x = toX(hoverIndex);
      ctx.strokeStyle = 'rgba(255,255,255,0.2)';
      ctx.setLineDash([3,3]);
      ctx.beginPath(); ctx.moveTo(x,pad.t); ctx.lineTo(x,pad.t+ch); ctx.stroke();
      ctx.setLineDash([]);

      const rows = [];
      let yRef = my;
      if(hoverIndex < hist.length){
        rows.push(['Actual', hist[hoverIndex]]);
        yRef = toY(hist[hoverIndex]);
        drawDot(x, yRef, hoverIndex < trainN ? '#2fd3a2' : '#ffb84d');
      }
      if(hoverIndex >= trainN && hoverIndex < hist.length && valPred[hoverIndex-trainN] != null){
        rows.push(['Validation Pred', valPred[hoverIndex-trainN]]);
      }
      if(hoverIndex >= hist.length && fc[hoverIndex-hist.length] != null){
        rows.push(['Forecast', fc[hoverIndex-hist.length]]);
        yRef = toY(fc[hoverIndex-hist.length]);
      }
      tip.classList.remove('hidden');
      tip.innerHTML = `<div class="tipTitle">${esc(hoverIndex < hist.length ? histLabels[hoverIndex] : `F+${hoverIndex-hist.length+1}`)}</div>` + rows.map(rw => `<div class="tipRow"><span>${esc(rw[0])}</span><b>${num(rw[1],2)}</b></div>`).join('');
      tip.style.left = `${Math.min(W-190, Math.max(12, mx + 12))}px`;
      tip.style.top = `${Math.min(H-100, Math.max(12, yRef - 18))}px`;
    } else {
      tip.classList.add('hidden');
    }
  }

  paint();
  canvas.onmousemove = (e) => {
    const rect = canvas.getBoundingClientRect();
    const mx = e.clientX - rect.left;
    const my = e.clientY - rect.top;
    if(mx < pad.l || mx > pad.l+cw || my < pad.t || my > pad.t+ch){ paint(); return; }
    const idx = Math.max(0, Math.min(totalPts-1, Math.round(((mx-pad.l)/cw) * Math.max(1,totalPts-1))));
    paint(idx, mx, my);
  };
  canvas.onmouseleave = () => paint();
}

function exportExcel(){
  if(!results.length){
    alert('Nothing to export.');
    return;
  }
  const wb = XLSX.utils.book_new();
  const h = parseInt($('horizonN').value, 10) || 8;
  const unit = $('horizonUnit').value === 'week' ? 'Wk' : 'Mo';

  const summaryRows = [
    ['Item Code','Description','Winner Model','Rank Score','MAPE%','NMAE%','Avg / Period','Total Hist','Trend','Outliers', ...Array.from({length:h}, (_,i)=>`${unit}${i+1}`), 'Next 4','Forecast Total'],
    ...results.map(r => [
      r.code, r.desc, r.model_name, r.rank_score, r.mape, r.nmae, r.avg_per_period, r.total_hist, r.trend, r.outlier_count,
      ...(r.forecast || []), sum((r.forecast||[]).slice(0,4)), sum(r.forecast||[])
    ])
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryRows), 'Forecast Summary');

  const active = getSelectedModels();
  const compRows = [
    ['Item Code','Description','Winner','Winner Rank','Winner MAPE','Winner NMAE', ...active.map(id => `${models.find(m=>m.id===id)?.name || id} Rank`)],
    ...results.map(r => [
      r.code, r.desc, r.model_name, r.rank_score, r.mape, r.nmae,
      ...active.map(id => r.model_scores?.[id]?.rank_score ?? null)
    ])
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(compRows), 'Model Competition');

  const valRows = [['Item Code','Description','Validation Step','Actual','Winner Prediction']];
  results.forEach(r => {
    const pred = r.val_predictions?.[r.winner] || [];
    (r.val_actual || []).forEach((a,i) => valRows.push([r.code, r.desc, i+1, a, pred[i] ?? null]));
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(valRows), 'Validation Detail');

  const runRows = [
    ['Generated', new Date().toLocaleString()],
    ['API Base URL', window.APP_CONFIG?.API_BASE_URL || ''],
    ['Horizon Unit', $('horizonUnit').value],
    ['Horizon Length', $('horizonN').value],
    ['Lookback Months', $('lookbackN').value],
    ['Validation Steps', $('valN').value],
    ['Season Period', $('seasonN').value],
    ['Selected Models', getSelectedModels().join(', ')]
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(runRows), 'Run Info');

  const ts = new Date();
  const stamp = `${ts.getFullYear()}${String(ts.getMonth()+1).padStart(2,'0')}${String(ts.getDate()).padStart(2,'0')}_${String(ts.getHours()).padStart(2,'0')}${String(ts.getMinutes()).padStart(2,'0')}`;
  XLSX.writeFile(wb, `ForecastDemand_${stamp}.xlsx`);
}

window.runForecast = runForecast;
window.exportExcel = exportExcel;
window.pickSku = pickSku;
window.setSort = setSort;
window.renderSkuList = renderSkuList;
window.addEventListener('resize', () => { if(results.length) renderChart(); });
window.addEventListener('DOMContentLoaded', init);

window.renderChart = renderChart;
