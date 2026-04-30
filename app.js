
const BLUE = '#1a3a5c';
const BLUE_MID = '#2a5f8c';
const BLUE_LIGHT = '#4a8ab5';
const BLUE_PALE = '#85b7d9';
const BLUE_FAINT = '#c5dced';
const SLATE = '#3a4a5c';
const STEEL = '#5a7a8c';
const MUTED = '#8ca5b5';
const WARN = '#c84b2f';
const NEUTRAL = '#b4b2a9';
const PALETTE = [BLUE, BLUE_MID, BLUE_LIGHT, SLATE, STEEL, BLUE_PALE, MUTED, '#2a4a6c','#1a4a6c','#3a6a8c'];
const gridC = 'rgba(15,14,12,0.07)', tickC = '#7a7870';
const baseChart = {
  responsive: true, maintainAspectRatio: false,
  plugins: { legend: { display: false }, tooltip: { backgroundColor: '#0f0e0c', titleColor: '#f9f6f0', bodyColor: '#c8c5bb', padding: 10, cornerRadius: 4 } },
  scales: {
    x: { grid: { color: gridC }, ticks: { color: tickC, font: { size: 11, family: "'Segoe UI'" } } },
    y: { grid: { color: gridC }, ticks: { color: tickC, font: { size: 11, family: "'Segoe UI'" } } }
  }
};

// Country → Region classifier
const c2r = c => {
  const v = (c||'').toLowerCase().trim();
  if (v === 'india') return 'India';
  if (['singapore','malaysia','indonesia','philippines','thailand','vietnam','china','japan','south korea','australia','new zealand','hong kong','taiwan','bangladesh','sri lanka','myanmar','cambodia'].includes(v)) return 'APAC';
  if (['united states','usa','us','canada','mexico'].includes(v)) return 'North America';
  if (['brazil','argentina','colombia','chile','peru','venezuela','ecuador','bolivia','paraguay','uruguay','costa rica','panama'].includes(v)) return 'Latin America';
  return 'EMEA';
};

let charts = {};
const dc = id => { if (charts[id]) { charts[id].destroy(); delete charts[id]; } };

// Drag & drop
const dz = document.getElementById('dropZone');
const fi = document.getElementById('fileInput');
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag-over'); if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]); });
fi.addEventListener('change', e => { if (e.target.files[0]) processFile(e.target.files[0]); });

function processFile(file) {
  document.getElementById('upload-screen').style.display = 'none';
  document.getElementById('loading').style.display = 'flex';
  const isXlsx = /\.(xlsx|xls)$/i.test(file.name);
  const r = new FileReader();
  if (isXlsx) {
    r.onload = e => setTimeout(() => {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      // Convert to array of arrays
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      // Find the real header row (first row with 5+ non-empty cells containing known keywords)
      const HEADER_KEYWORDS = ['employee','gender','country','company','status','hire','salary','email','org','cost','segment','product','birth'];
      let headerIdx = 0;
      for (let i = 0; i < Math.min(10, aoa.length); i++) {
        const cells = aoa[i].map(c => String(c||'').toLowerCase());
        const nonEmpty = cells.filter(c=>c).length;
        const matches = HEADER_KEYWORDS.filter(k => cells.some(c=>c.includes(k))).length;
        if (nonEmpty >= 5 && matches >= 3) { headerIdx = i; break; }
      }
      const headers = aoa[headerIdx].map(h => String(h||'').trim().toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,''));
      const rows = [];
      for (let i = headerIdx + 1; i < aoa.length; i++) {
        if (aoa[i].every(c => c === '' || c == null)) continue;
        const row = {};
        headers.forEach((h, j) => {
          const val = aoa[i][j];
          // Format dates from SheetJS as YYYY-MM-DD
          if (val instanceof Date) {
            row[h] = val.toISOString().slice(0,10);
          } else {
            row[h] = String(val == null ? '' : val).trim();
          }
        });
        rows.push(row);
      }
      renderReport(processRows(rows, headers, file.name));
    }, 800);
    r.readAsArrayBuffer(file);
  } else {
    r.onload = e => setTimeout(() => renderReport(parseHC(e.target.result, file.name)), 800);
    r.readAsText(file, 'latin-1');
  }
}

// Proper CSV parser — handles quoted fields with commas inside
function parseCSVLine(line, delim) {
  if (delim === '\t') return line.split('\t').map(v => v.trim());
  const result = [];
  let cur = '', inQuote = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuote && line[i+1] === '"') { cur += '"'; i++; }
      else inQuote = !inQuote;
    } else if (ch === ',' && !inQuote) {
      result.push(cur.trim()); cur = '';
    } else {
      cur += ch;
    }
  }
  result.push(cur.trim());
  return result;
}

function parseHC(text, filename) {
  const lines = text.trim().split(/\r?\n/);
  const delim = lines[0].includes('\t') ? '\t' : ',';
  const HEADER_KEYWORDS = ['employee','gender','country','company','status','hire','salary','email','org','cost','segment','product','birth'];
  let headerLineIdx = 0;
  for (let i = 0; i < Math.min(10, lines.length); i++) {
    const cells = parseCSVLine(lines[i], delim).map(c => c.toLowerCase());
    const nonEmpty = cells.filter(c => c).length;
    const matches = HEADER_KEYWORDS.filter(k => cells.some(c => c.includes(k))).length;
    if (nonEmpty >= 5 && matches >= 3) { headerLineIdx = i; break; }
  }
  const headers = parseCSVLine(lines[headerLineIdx], delim).map(h => h.toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,''));
  const rows = [];
  for (let i = headerLineIdx + 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const vals = parseCSVLine(lines[i], delim);
    const row = {};
    headers.forEach((h, j) => row[h] = (vals[j] || '').trim());
    rows.push(row);
  }
  return processRows(rows, headers, filename);
}

function processRows(rows, headers, filename) {
  const find = cs => cs.find(c => headers.includes(c)) || null;

  const colCompany  = find(['company','legal_entity','entity','employer']);
  const colCountry  = find(['country','location']);
  const colOrg2     = find(['org_level_2_cost_center','org_level_2','cost_center','org_level2','l2','costcenter']);
  const colOrg3     = find(['org_level_3_ops_segment','org_level_3','ops_segment','org_level3','l3','sub_department','department']);
  const colOrg4     = find(['org_level_4_product','org_level_4','org_level4','l4','product','product_line']);
  const colGender   = find(['gender','sex']);
  const colHireDate = find(['original_hire_date','hire_date','start_date','date_of_joining']);
  const colLastHire = find(['last_hire_date','rehire_date']);
  const colDOB      = find(['date_of_birth','dob','birth_date']);
  const colStatus   = find(['employment_status','emp_status','employee_status','worker_status']);
  const colTermType  = find(['termination_type','term_type','exit_type']);
  const colTermReason= find(['termination_reason','term_reason','exit_reason','reason_for_leaving']);
  const colTermDate  = find(['termination_date','term_date','exit_date','last_day']);

  const ACTIVE_VALUES = new Set(['active','a','1','yes','leave of absence','loa','leave','on leave','on loa']);
  const colEmpType = find(['employee_type','emp_type','worker_type','employment_type']);
  const VALID_EMP_TYPES = new Set(['regular','peo']);

  const today = new Date();
  rows.forEach(r => {
    const statusVal = (colStatus && r[colStatus] || '').toLowerCase().trim();
    const empTypeVal = (colEmpType && r[colEmpType] || '').toLowerCase().trim();
    const validEmpType = !colEmpType || VALID_EMP_TYPES.has(empTypeVal);
    r._isActive = (!colStatus || ACTIVE_VALUES.has(statusVal)) && validEmpType;
    r._isEligible = validEmpType; // flag for termed rows too
    r._termType   = (colTermType   && r[colTermType])   || null;
    r._termReason = (colTermReason && r[colTermReason]) || null;
    r._termDate   = (colTermDate   && r[colTermDate])   || null;
    r._company = (colCompany && r[colCompany]) || 'Unknown';
    r._country  = (colCountry && r[colCountry]) || 'Unknown';
    r._region   = c2r(r._country);
    r._org2     = (colOrg2 && r[colOrg2]) || 'Unknown';
    r._org3     = (colOrg3 && r[colOrg3]) || 'Unknown';
    r._org4     = (colOrg4 && r[colOrg4]) || 'Unknown';
    r._gender   = (colGender && r[colGender]) || null;
    r._tenureYears = null;
    const rd = (colHireDate && r[colHireDate]) || (colLastHire && r[colLastHire]);
    if (rd) { const d = new Date(rd); if (!isNaN(d)) r._tenureYears = (today - d) / (1000*60*60*24*365.25); }
  });

  const activeRows = rows.filter(r => r._isActive && r._org3 !== 'Packaging');
  const excludedCount = rows.length - activeRows.length;
  // Tag _company on allRows too so KPI company count works across all entities
  return { rows: activeRows, allRows: rows, excludedCount, filename, colOrg3, colOrg4, colCompany, colStatus, colTermType, colTermReason, colTermDate };
}

const countBy = (rows, key) => { const m = {}; rows.forEach(r => { const v = r[key]||'Unknown'; m[v]=(m[v]||0)+1; }); return m; };
const avgBy = (rows, key, val) => { const s={}, c={}; rows.forEach(r => { const k=r[key]||'Unknown'; if (r[val]!=null){s[k]=(s[k]||0)+r[val];c[k]=(c[k]||0)+1;} }); const o={}; Object.keys(s).forEach(k=>o[k]=s[k]/c[k]); return o; };
const sorted = (obj, desc=true) => Object.entries(obj).sort((a,b) => desc ? b[1]-a[1] : a[1]-b[1]);

function estAttrition(activeRows, allRows) {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);
  const exits = allRows.filter(r => {
    if (!r._isActive && (r._termType||'').toLowerCase() === 'voluntary') {
      const d = new Date(r._termDate);
      return !isNaN(d) && d >= cutoff;
    }
    return false;
  }).length;
  const total = activeRows.length + exits;
  return total > 0 ? parseFloat((exits/total*100).toFixed(2)) : null;
}

function renderReport(data) {
  try {
  const { rows, excludedCount, filename, colOrg3, colOrg4, colStatus } = data;
  const total = rows.length;
  const withTenure = rows.filter(r => r._tenureYears != null);
  const avgTenure = withTenure.length ? withTenure.reduce((s,r)=>s+r._tenureYears,0)/withTenure.length : null;
  const attrPct = estAttrition(rows, data.allRows);
  const companyMap = countBy(rows,'_company');
  const companies = Object.keys(companyMap).length;
  const topCompany = sorted(companyMap)[0];

  const termedNote = colStatus && excludedCount > 0 ? `${excludedCount.toLocaleString()} non-active excluded` : '';
  document.getElementById('trendNote').textContent = `${total.toLocaleString()} employees`;
  const ytdStartCo = new Date(new Date().getFullYear(), 0, 1);
  const last30Rows = data.allRows.filter(r => {
    if (r._isActive) return false;
    const d = new Date(r._termDate);
    return !isNaN(d) && d >= ytdStartCo && (r._termType||'').toLowerCase() !== 'transfer';
  });
  const last30Companies = new Set(last30Rows.map(r => r._company));
  const allCompanies = Object.keys(countBy(data.allRows,'_company'));
  // Use full active rows (not Packaging-filtered) for company presence check
  const allActiveRows = data.allRows.filter(r => r._isActive);
  const activeCompanies = new Set(allActiveRows.map(r=>r._company));
  const inactiveCompanies = allCompanies.filter(c => !activeCompanies.has(c) && !last30Companies.has(c));
  document.getElementById('kpiStrip').innerHTML = [
    { label: 'Active entities', value: activeCompanies.size },
    
    { label: 'Active headcount', value: total.toLocaleString() },
    { label: 'Voluntary attrition (YTD)', value: attrPct != null ? attrPct.toFixed(2) + '%' : '—' },
  ].map(k => `<div class="kpi"><div class="kpi-label">${k.label}</div><div class="kpi-value">${k.value}</div></div>`).join('');

  // narrative called after attrition data is built below

  // ── Gender donut — same style as Vol/Invol chart ──
  const gMap = countBy(rows,'_gender');
  const gEntries = sorted(gMap);
  const gColors = [BLUE, BLUE_LIGHT, BLUE_PALE, SLATE, STEEL];
  dc('genderChart');
  if (window._genderResizeObs) { window._genderResizeObs.disconnect(); window._genderResizeObs = null; }
  charts['genderChart'] = new Chart(document.getElementById('genderChart'), {
    type: 'doughnut',
    data: {
      labels: gEntries.map(e=>e[0]),
      datasets: [{ data: gEntries.map(e=>e[1]), backgroundColor: gColors, borderWidth: 3, borderColor: '#fff' }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '58%',
      plugins: {
        legend: { display: true, position: 'bottom', labels: { color: tickC, font: { size: 12, family:"'Segoe UI'" }, padding: 14, boxWidth: 12, boxHeight: 12 } },
        tooltip: { backgroundColor:'#0f0e0c', titleColor:'#f9f6f0', bodyColor:'#c8c5bb',
          callbacks: { label: ctx => ` ${ctx.label}: ${ctx.parsed} (${Math.round(ctx.parsed/total*100)}%)` } }
      },
      animation: { onComplete: function() {
        const ctx2 = this.ctx;
        const W2 = this.chartArea;
        const cx2 = (W2.left+W2.right)/2, cy2 = (W2.top+W2.bottom)/2;
        ctx2.save();
        ctx2.fillStyle = '#0f0e0c'; ctx2.font = '500 24px DM Sans, sans-serif';
        ctx2.textAlign = 'center'; ctx2.textBaseline = 'middle';
        ctx2.fillText(total.toLocaleString(), cx2, cy2 - 10);
        ctx2.fillStyle = '#7a7870'; ctx2.font = '400 11px DM Sans, sans-serif';
        ctx2.fillText('employees', cx2, cy2 + 10);
        ctx2.restore();
        // Pct labels on each slice
        this.data.datasets.forEach((ds, di) => {
          this.getDatasetMeta(di).data.forEach((arc, j) => {
            const val = ds.data[j];
            if (!val) return;
            const pct = Math.round(val/total*100);
            if (pct < 5) return;
            const pos = arc.tooltipPosition();
            ctx2.save();
            ctx2.fillStyle = '#fff';
            ctx2.font = '500 13px DM Sans, sans-serif';
            ctx2.textAlign = 'center'; ctx2.textBaseline = 'middle';
            ctx2.fillText(pct+'%', pos.x, pos.y);
            ctx2.restore();
          });
        });
      }}
    }
  });
  // ── HC Trend — monthly new hires last 12 months, last 6 months highlighted ──
  const hbm = {};
  const twelveMonthsAgo = new Date(); twelveMonthsAgo.setFullYear(twelveMonthsAgo.getFullYear()-1);
  const sixMonthsAgo = new Date(); sixMonthsAgo.setMonth(sixMonthsAgo.getMonth()-6);
  rows.filter(r=>r._tenureYears!=null).forEach(r => {
    const d = new Date(Date.now() - r._tenureYears*365.25*24*60*60*1000);
    if (d < twelveMonthsAgo) return; // last 12 months only
    const ym = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    hbm[ym] = (hbm[ym]||0)+1;
  });
  const sMonths = Object.keys(hbm).sort();
  const fmtMo = m => { const [y,mo]=m.split('-'); return `${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+mo-1]} ${y.slice(2)}`; };
  // Last 6 months get darker blue, earlier months get lighter
  const barBgColors = sMonths.map(m => {
    const d = new Date(m+'-01');
    return d >= sixMonthsAgo ? BLUE : BLUE_PALE;
  });
  dc('hcTrendChart');
  charts['hcTrendChart'] = new Chart(document.getElementById('hcTrendChart'), {
    type: 'bar',
    data: { labels: sMonths.map(fmtMo),
            datasets: [{ label:'New hires', data: sMonths.map(m=>hbm[m]||0), backgroundColor: barBgColors, borderRadius: 4 }] },
    options: { ...baseChart,
      plugins: { ...baseChart.plugins,
        legend: { display: true, position: 'top', labels: { color: tickC, font: { size: 11 }, padding: 12, boxWidth: 10, boxHeight: 10,
          generateLabels: () => [
            { text: 'Last 6 months', fillStyle: BLUE, strokeStyle: BLUE },
            { text: 'Earlier', fillStyle: BLUE_PALE, strokeStyle: BLUE_PALE }
          ]
        }}
      },
      scales: { ...baseChart.scales, x: { ...baseChart.scales.x, ticks: { ...baseChart.scales.x.ticks, maxTicksLimit: 12, autoSkip: false, maxRotation: 45 } } }
    }
  });

  // ── Cost Center chart (Org Level 2 actual values) ──
  const ccMap = countBy(rows,'_org2');
  const ccEntries = sorted(ccMap).slice(0,15);
  dc('costCenterChart');
  charts['costCenterChart'] = new Chart(document.getElementById('costCenterChart'), {
    type: 'bar',
    data: { labels: ccEntries.map(e=>e[0]), datasets: [{ label:'Headcount', data: ccEntries.map(e=>e[1]), backgroundColor: BLUE, borderRadius:3 }] },
    options: { ...baseChart, indexAxis:'y', scales: { x:{...baseChart.scales.x}, y:{...baseChart.scales.y, ticks:{color:tickC, font:{size:11, family:"'Segoe UI'"}}} } }
  });

  // ── Ops Segment chart (Org Level 3) ──
  const opsMap = countBy(rows,'_org3');
  const opsEntries = sorted(opsMap).filter(e=>e[0]!=='Unknown').slice(0,15);
  dc('opsSegChart');
  if (opsEntries.length) {
    charts['opsSegChart'] = new Chart(document.getElementById('opsSegChart'), {
      type: 'bar',
      data: { labels: opsEntries.map(e=>e[0]), datasets: [{ label:'Headcount', data: opsEntries.map(e=>e[1]), backgroundColor: BLUE_MID, borderRadius:3 }] },
      options: { ...baseChart, indexAxis:'y', scales: { x:{...baseChart.scales.x}, y:{...baseChart.scales.y, ticks:{color:tickC, font:{size:11, family:"'Segoe UI'"}}} } }
    });
  } else {
    document.querySelector('#tab-ops .chart-wrap').innerHTML = '<p style="color:var(--ink-3);font-size:13px;padding:1rem 0;">Org Level 3 column not detected in your file.</p>';
  }

  // ── Region chart (countries classified) ──
  const regMap = countBy(rows,'_region');
  const regEntries = sorted(regMap);
  dc('regionChart');
  charts['regionChart'] = new Chart(document.getElementById('regionChart'), {
    type: 'bar',
    data: { labels: regEntries.map(e=>e[0]), datasets: [{ label:'Headcount', data: regEntries.map(e=>e[1]), backgroundColor: BLUE_LIGHT, borderRadius:3 }] },
    options: { ...baseChart, indexAxis:'y', scales: { x:{...baseChart.scales.x}, y:{...baseChart.scales.y, ticks:{color:tickC, font:{size:11, family:"'Segoe UI'"}}} } }
  });

  // ── Attrition — YTD, excl transfers, excl Packaging ──
  const termedRows = data.allRows.filter(r => !r._isActive && r._org3 !== 'Packaging' && r._isEligible);
  const ytdStart = new Date(new Date().getFullYear(), 0, 1);
  const termedLast12 = termedRows.filter(r => {
    if (!r._termDate) return false;
    const d = new Date(r._termDate);
    return !isNaN(d) && d >= ytdStart && (r._termType||'').toLowerCase() !== 'transfer';
  });
  const volRows    = termedLast12.filter(r => (r._termType||'').toLowerCase() === 'voluntary');
  const involRows  = termedLast12.filter(r => (r._termType||'').toLowerCase() === 'involuntary');
  const volCount   = volRows.length;
  const involCount = involRows.length;
  const totalTermed = volCount + involCount;

  document.getElementById('attrNote').textContent = `${totalTermed} exits (YTD Jan–${new Date().toLocaleString('default',{month:'short'})}) · ${volCount} voluntary · ${involCount} involuntary`;

  // ── AI narrative — called here so all attrition vars are in scope ──
  buildNarrative(rows, total, attrPct, volCount, involCount, termedLast12, termedRows, data.allRows);

  const groupBy = (rows, key) => {
    const m = {};
    rows.forEach(r => { const v = r[key]||'Unknown'; m[v]=(m[v]||0)+1; });
    return Object.entries(m).filter(([k])=>k!=='Unknown').sort((a,b)=>b[1]-a[1]);
  };
  const mkSepBar = (id, entries, isVol) => {
    dc(id);
    if (!entries.length) { const el=document.getElementById(id); if(el) el.parentElement.innerHTML='<p style="font-size:12px;color:var(--ink-3);font-style:italic;padding:8px 0">No exits in this category</p>'; return; }
    const vC=[BLUE,BLUE_MID,BLUE_LIGHT,SLATE,STEEL];
    const iC=[WARN,'#e07055','#e89a85','#f0bfb2'];
    charts[id] = new Chart(document.getElementById(id), {
      type:'bar',
      data:{labels:entries.map(e=>e[0]),datasets:[{data:entries.map(e=>e[1]),backgroundColor:entries.map((_,i)=>(isVol?vC:iC)[i%(isVol?vC:iC).length]),borderRadius:4,barThickness:18}]},
      options:{...baseChart,indexAxis:'y',scales:{x:{...baseChart.scales.x,ticks:{...baseChart.scales.x.ticks,stepSize:1}},y:{...baseChart.scales.y,ticks:{color:tickC,font:{size:11}}}}}
    });
  };

  // ── Voluntary vs Involuntary donut ──
  dc('attrSplitChart');
  charts['attrSplitChart'] = new Chart(document.getElementById('attrSplitChart'), {
    type: 'doughnut',
    data: {
      labels: ['Involuntary','Voluntary'],
      datasets: [{ data: [involCount, volCount], backgroundColor: [WARN, BLUE], borderWidth: 3, borderColor: '#fff' }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '58%',
      plugins: {
        legend: { display: true, position: 'bottom', labels: { color: tickC, font: { size: 12, family:"'Segoe UI'" }, padding: 14, boxWidth: 12, boxHeight: 12 } },
        tooltip: { backgroundColor:'#0f0e0c', titleColor:'#f9f6f0', bodyColor:'#c8c5bb',
          callbacks: { label: ctx => ` ${ctx.label}: ${ctx.parsed} (${Math.round(ctx.parsed/totalTermed*100)}%)` } }
      },
      animation: { onComplete: function() {
        const ctx2 = this.ctx;
        const W2 = this.chartArea;
        const cx2 = (W2.left+W2.right)/2, cy2 = (W2.top+W2.bottom)/2;
        ctx2.save();
        ctx2.fillStyle = '#0f0e0c'; ctx2.font = '500 24px DM Sans, sans-serif';
        ctx2.textAlign = 'center'; ctx2.textBaseline = 'middle';
        ctx2.fillText(totalTermed, cx2, cy2 - 10);
        ctx2.fillStyle = '#7a7870'; ctx2.font = '400 11px DM Sans, sans-serif';
        ctx2.fillText('exits', cx2, cy2 + 10);
        ctx2.restore();
        this.data.datasets.forEach((ds, di) => {
          this.getDatasetMeta(di).data.forEach((arc, j) => {
            const val = ds.data[j];
            if (!val) return;
            const pct = Math.round(val/totalTermed*100);
            if (pct < 5) return;
            const pos = arc.tooltipPosition();
            ctx2.save();
            ctx2.fillStyle = '#fff';
            ctx2.font = '500 13px DM Sans, sans-serif';
            ctx2.textAlign = 'center'; ctx2.textBaseline = 'middle';
            ctx2.fillText(pct+'%', pos.x, pos.y);
            ctx2.restore();
          });
        });
      }}
    }
  });

  // ── Vol/Invol donut ──
  dc('attrSplitChart');
  charts['attrSplitChart'] = new Chart(document.getElementById('attrSplitChart'), {
    type:'doughnut', data:{labels:['Involuntary','Voluntary'],datasets:[{data:[involCount,volCount],backgroundColor:[WARN,BLUE],borderWidth:3,borderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'58%',
      plugins:{legend:{display:true,position:'bottom',labels:{color:tickC,font:{size:12},padding:14,boxWidth:12,boxHeight:12}},
        tooltip:{backgroundColor:'#0f0e0c',titleColor:'#f9f6f0',bodyColor:'#c8c5bb',callbacks:{label:ctx=>` ${ctx.label}: ${ctx.parsed} (${totalTermed?Math.round(ctx.parsed/totalTermed*100):0}%)`}}},
      animation:{onComplete:function(){
        const c=this.ctx,w=this.chartArea,cx=(w.left+w.right)/2,cy=(w.top+w.bottom)/2;
        c.save();c.fillStyle='#0f0e0c';c.font='500 24px sans-serif';c.textAlign='center';c.textBaseline='middle';
        c.fillText(totalTermed,cx,cy-10);c.fillStyle='#7a7870';c.font='400 11px sans-serif';c.fillText('exits',cx,cy+10);c.restore();
        this.data.datasets.forEach((ds,di)=>{this.getDatasetMeta(di).data.forEach((arc,j)=>{
          const val=ds.data[j];if(!val||!totalTermed)return;const pct=Math.round(val/totalTermed*100);if(pct<5)return;
          const pos=arc.tooltipPosition();c.save();c.fillStyle='#fff';c.font='500 13px sans-serif';c.textAlign='center';c.textBaseline='middle';c.fillText(pct+'%',pos.x,pos.y);c.restore();
        });});
      }}}
  });

  // ── Regional attrition bars ──
  const regTermMap2 = {};
  volRows.forEach(r => { regTermMap2[r._region]=(regTermMap2[r._region]||0)+1; });
  const regAttrArr = regEntries.map(([reg,activeCount]) => ({
    reg, activeCount, vc: regTermMap2[reg]||0,
    pct: (activeCount+(regTermMap2[reg]||0))>0 ? Math.round((regTermMap2[reg]||0)/(activeCount+(regTermMap2[reg]||0))*100) : 0
  })).sort((a,b)=>b.pct-a.pct);
  const maxRegPct = Math.max(...regAttrArr.map(e=>e.pct),1);
  document.getElementById('regionAttrBars').innerHTML = regAttrArr.map((e,i) =>
    `<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
      <div style="font-size:12px;color:var(--ink-3);width:130px;flex-shrink:0;text-align:right">${e.reg}</div>
      <div style="flex:1;background:#eee;border-radius:3px;height:20px;overflow:hidden">
        <div style="width:${e.pct>0?Math.max(Math.round(e.pct/maxRegPct*100),3):2}%;height:100%;background:${[BLUE,BLUE_MID,BLUE_LIGHT,SLATE,STEEL,BLUE_PALE][i%6]};border-radius:3px;display:flex;align-items:center;padding-left:8px">
          <span style="font-size:11px;font-weight:500;color:#fff">${e.pct>0?e.pct+'%':''}</span>
        </div>
      </div>
      <div style="font-size:11px;color:var(--ink-3);width:50px;text-align:right">${e.vc} exits</div>
    </div>`
  ).join('') + `<p style="font-size:11px;color:var(--ink-3);margin-top:8px">Vol exits ÷ (active + vol exits) · last 30 days</p>`;

  // ── Month on month chart — all time, excl transfers & Packaging ──
  const momMap = {};
  termedRows.filter(r => (r._termType||'').toLowerCase() !== 'transfer').forEach(r => {
    if (!r._termDate) return;
    const d = new Date(r._termDate);
    if (isNaN(d)) return;
    const ym = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    if (!momMap[ym]) momMap[ym] = { Voluntary:0, Involuntary:0 };
    const t = r._termType || 'Unknown';
    if (momMap[ym][t] !== undefined) momMap[ym][t]++;
  });
  const momMonths = Object.keys(momMap).sort().slice(-24);
  const fmtM = m => { const [y,mo]=m.split('-'); return `${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+mo-1]} ${y.slice(2)}`; };
  dc('momChart');
  charts['momChart'] = new Chart(document.getElementById('momChart'), {
    type: 'bar',
    data: {
      labels: momMonths.map(fmtM),
      datasets: [
        { label:'Involuntary', data: momMonths.map(m=>(momMap[m]||{}).Involuntary||0), backgroundColor: WARN, borderRadius:3, stack:'s' },
        { label:'Voluntary',   data: momMonths.map(m=>(momMap[m]||{}).Voluntary||0),   backgroundColor: BLUE, borderRadius:3, stack:'s' },
      ]
    },
    options: { ...baseChart, scales: { x:{...baseChart.scales.x, stacked:true, ticks:{...baseChart.scales.x.ticks, maxTicksLimit:12, autoSkip:true, maxRotation:45}}, y:{...baseChart.scales.y, stacked:true} } }
  });

  // ── Exit breakdowns — cost center ──
  mkSepBar('ccVolChart',   groupBy(volRows,   '_org2'), true);
  mkSepBar('ccInvolChart', groupBy(involRows, '_org2'), false);

  // ── Exit breakdowns — ops segment ──
  mkSepBar('opsVolChart',   groupBy(volRows,   '_org3'), true);
  mkSepBar('opsInvolChart', groupBy(involRows, '_org3'), false);

  // ── Exit breakdowns — product ──
  mkSepBar('prodVolChart',   groupBy(volRows,   '_org4'), true);
  mkSepBar('prodInvolChart', groupBy(involRows, '_org4'), false);

  // ── Termination reasons — separate ──
  mkSepBar('reasonVolChart',   groupBy(volRows,   '_termReason'), true);
  mkSepBar('reasonInvolChart', groupBy(involRows, '_termReason'), false);

  // ── Attrition by ops segment — tabbed — last 12 months ──
  const opsAttrData = {};
  const cutoff12months = new Date();
  cutoff12months.setFullYear(cutoff12months.getFullYear() - 1);
  const termed12m = termedRows.filter(r => {
    if (!r._termDate) return false;
    const d = new Date(r._termDate);
    return !isNaN(d) && d >= cutoff12months && (r._termType||'').toLowerCase() !== 'transfer';
  });
  const vol12m   = termed12m.filter(r => (r._termType||'').toLowerCase() === 'voluntary');
  const invol12m = termed12m.filter(r => (r._termType||'').toLowerCase() === 'involuntary');

  const opsSegments = [...new Set(rows.map(r => r._org3).filter(s => s && s !== 'Unknown'))].sort();

  opsSegments.forEach(seg => {
    const segActive  = rows.filter(r => r._org3 === seg);
    const segVol     = vol12m.filter(r => r._org3 === seg);
    const segInvol   = invol12m.filter(r => r._org3 === seg);

    const byCC = {};
    segActive.forEach(r => { if (!byCC[r._org2]) byCC[r._org2]={active:0,vol:0,invol:0}; byCC[r._org2].active++; });
    segVol.forEach(r => { if (!byCC[r._org2]) byCC[r._org2]={active:0,vol:0,invol:0}; byCC[r._org2].vol++; });
    segInvol.forEach(r => { if (!byCC[r._org2]) byCC[r._org2]={active:0,vol:0,invol:0}; byCC[r._org2].invol++; });

    const byProd = {};
    segActive.forEach(r => { if (!byProd[r._org4]) byProd[r._org4]={active:0,vol:0,invol:0}; byProd[r._org4].active++; });
    segVol.forEach(r => { if (!byProd[r._org4]) byProd[r._org4]={active:0,vol:0,invol:0}; byProd[r._org4].vol++; });
    segInvol.forEach(r => { if (!byProd[r._org4]) byProd[r._org4]={active:0,vol:0,invol:0}; byProd[r._org4].invol++; });

    opsAttrData[seg] = { cc: byCC, prod: byProd, active: segActive.length, vol: segVol.length, invol: segInvol.length };
  });

  function renderOpsAttrChart(seg) {
    const d = opsAttrData[seg];
    if (!d) return;
    document.getElementById('opsAttrCCTitle').textContent = `${seg} — by cost center`;
    document.getElementById('opsAttrProdTitle').textContent = `${seg} — by product`;

    const buildAttrPct = (map, canvasId) => {
      dc(canvasId);
      // Only show rows with at least one exit (vol or invol), sorted by vol attrition % desc
      const labels = Object.keys(map).filter(k => k !== 'Unknown' && map[k].active > 0 && (map[k].vol > 0 || map[k].invol > 0))
        .sort((a,b) => {
          const pctA = map[a].vol / (map[a].active + map[a].vol) * 100;
          const pctB = map[b].vol / (map[b].active + map[b].vol) * 100;
          return pctB - pctA;
        });
      if (!labels.length) return;
      const volPct   = labels.map(l => { const r=map[l]; return r.active+r.vol > 0 ? parseFloat((r.vol/(r.active+r.vol)*100).toFixed(2)) : 0; });
      const involPct = labels.map(l => { const r=map[l]; return r.active+r.invol > 0 ? parseFloat((r.invol/(r.active+r.invol)*100).toFixed(2)) : 0; });
      const h = Math.max(labels.length * 44 + 80, 180);
      document.getElementById(canvasId).parentElement.style.height = h + 'px';
      charts[canvasId] = new Chart(document.getElementById(canvasId), {
        type: 'bar',
        data: { labels, datasets: [
          { label: 'Voluntary attrition %', data: volPct, backgroundColor: BLUE, borderRadius: 3 },
          { label: 'Involuntary attrition %', data: involPct, backgroundColor: WARN, borderRadius: 3 },
        ]},
        options: { ...baseChart, indexAxis: 'y',
          plugins: { ...baseChart.plugins,
            legend: { display: true, position: 'top', labels: { color: tickC, font: { size: 11 }, padding: 12, boxWidth: 10, boxHeight: 10 } },
            tooltip: { backgroundColor:'#0f0e0c', titleColor:'#f9f6f0', bodyColor:'#c8c5bb',
              callbacks: { label: ctx => {
                const lbl = ctx.dataset.label;
                const key = labels[ctx.dataIndex];
                const r = map[key];
                const exits = lbl.includes('Vol') ? r.vol : r.invol;
                return ` ${lbl}: ${ctx.parsed.x.toFixed(2)}% (${exits} exits YTD / ${r.active} active)`;
              }}
            }
          },
          scales: {
            x: { ...baseChart.scales.x, ticks: { ...baseChart.scales.x.ticks, callback: v => v + '%' }, max: Math.max(...volPct, ...involPct, 5) + 2 },
            y: { ...baseChart.scales.y, ticks: { color: tickC, font: { size: 10 } } }
          }
        }
      });
    };

    buildAttrPct(d.cc,   'opsAttrCCChart');
    buildAttrPct(d.prod, 'opsAttrProdChart');
    document.querySelectorAll('#opsAttrTabs .tab').forEach(t => t.classList.toggle('active', t.dataset.seg === seg));
  }

  // Build tabs
  const tabRow = document.getElementById('opsAttrTabs');
  tabRow.innerHTML = '';
  opsSegments.forEach((seg, i) => {
    const btn = document.createElement('button');
    btn.className = 'tab' + (i === 0 ? ' active' : '');
    btn.textContent = seg;
    btn.dataset.seg = seg;
    btn.onclick = () => renderOpsAttrChart(seg);
    tabRow.appendChild(btn);
  });
  if (opsSegments.length) renderOpsAttrChart(opsSegments[0]);



  // ── Tenure chart ──
  const bands = {'< 1 yr':0,'1–2 yr':0,'2–4 yr':0,'4–7 yr':0,'7–10 yr':0,'10+ yr':0};
  withTenure.forEach(r => {
    const t = r._tenureYears;
    if (t<1) bands['< 1 yr']++; else if (t<2) bands['1–2 yr']++;
    else if (t<4) bands['2–4 yr']++; else if (t<7) bands['4–7 yr']++;
    else if (t<10) bands['7–10 yr']++; else bands['10+ yr']++;
  });
  dc('tenureChart');
  charts['tenureChart'] = new Chart(document.getElementById('tenureChart'), {
    type: 'bar',
    data: { labels: Object.keys(bands), datasets: [{ label:'Employees', data: Object.values(bands), backgroundColor: Object.keys(bands).map((_,i)=>PALETTE[i%PALETTE.length]), borderRadius:4 }] },
    options: { ...baseChart }
  });

  // ── Build term maps for summary tables ──
  const ccTermMap = {};
  termedLast12.forEach(r => {
    const cc = r._org2||'Unknown';
    if (!ccTermMap[cc]) ccTermMap[cc] = {Voluntary:0, Involuntary:0};
    const t = (r._termType||'').toLowerCase();
    if (t === 'voluntary') ccTermMap[cc].Voluntary++;
    else if (t === 'involuntary') ccTermMap[cc].Involuntary++;
  });
  const prodTermMap = {};
  termedLast12.forEach(r => {
    const p = r._org4||'Unknown';
    if (!prodTermMap[p]) prodTermMap[p] = {Voluntary:0, Involuntary:0, _total:0};
    const t = (r._termType||'').toLowerCase();
    if (t === 'voluntary') { prodTermMap[p].Voluntary++; prodTermMap[p]._total++; }
    else if (t === 'involuntary') { prodTermMap[p].Involuntary++; prodTermMap[p]._total++; }
  });

  // ── Cost Center table (with vol/invol/transfer) ──
  const avgTCC = avgBy(rows,'_org2','_tenureYears');
  document.getElementById('ccTableBody').innerHTML = ccEntries.map(([cc, count]) => {
    const pct = Math.round(count/total*100);
    const tenure = avgTCC[cc] ? avgTCC[cc].toFixed(1) : '—';
    const fPct = Math.round(rows.filter(r=>r._org2===cc && r._gender && r._gender.toLowerCase().startsWith('f')).length / count * 100);
    const vol = ccTermMap[cc] ? ccTermMap[cc].Voluntary : 0;
    const invol = ccTermMap[cc] ? ccTermMap[cc].Involuntary : 0;
    return `<tr><td>${cc}</td><td>${count.toLocaleString()}</td><td><span class="badge badge-neu">${pct}%</span></td><td>${vol}</td><td>${invol}</td><td>${tenure}</td><td>${isNaN(fPct)?'—':fPct+'%'}</td></tr>`;
  }).join('');

  // ── Ops Segment table ──
  const avgTOps = avgBy(rows,'_org3','_tenureYears');
  document.getElementById('opsTableBody').innerHTML = opsEntries.length
    ? opsEntries.map(([seg, count]) => {
        const pct = Math.round(count/total*100);
        const tenure = avgTOps[seg] ? avgTOps[seg].toFixed(1) : '—';
        return `<tr><td>${seg}</td><td>${count.toLocaleString()}</td><td><span class="badge badge-neu">${pct}%</span></td><td>${tenure}</td></tr>`;
      }).join('')
    : '<tr><td colspan="4" style="color:var(--ink-3);font-style:italic;">Org Level 3 not found in file</td></tr>';

  // ── Product table (with vol/invol/transfer) ──
  const prodMap = countBy(rows,'_org4');
  const prodEntries = sorted(prodMap).filter(e=>e[0]!=='Unknown').slice(0,20);
  const avgTProd = avgBy(rows,'_org4','_tenureYears');
  document.getElementById('productTableBody').innerHTML = prodEntries.length
    ? prodEntries.map(([prod, count]) => {
        const pct = Math.round(count/total*100);
        const tenure = avgTProd[prod] ? avgTProd[prod].toFixed(1) : '—';
        const vol = prodTermMap[prod] ? prodTermMap[prod].Voluntary : 0;
        const invol = prodTermMap[prod] ? prodTermMap[prod].Involuntary : 0;
        return `<tr><td>${prod}</td><td>${count.toLocaleString()}</td><td><span class="badge badge-neu">${pct}%</span></td><td>${vol}</td><td>${invol}</td><td>${tenure}</td></tr>`;
      }).join('')
    : '<tr><td colspan="6" style="color:var(--ink-3);font-style:italic;">Org Level 4 (Product) not found in file</td></tr>';

  document.getElementById('loading').style.display = 'none';
  document.getElementById('report-screen').style.display = 'block';
  } catch(err) {
    document.getElementById('loading').style.display = 'none';
    document.getElementById('upload-screen').style.display = 'flex';
    alert('Report error: ' + err.message + '\n\nLine: ' + (err.stack ? err.stack.split('\n')[1] : 'unknown'));
  }
}

function buildNarrative(rows, total, attrPct, volCount, involCount, termedYTD, termedRows, allRows) {
  const insights = [];
  const flags = [];

  // ── Core metrics ──
  const regMap   = countBy(rows,'_region');
  const ccMap    = countBy(rows,'_org2');
  const opsMap   = countBy(rows,'_org3');
  const withT    = rows.filter(r=>r._tenureYears!=null);
  const avgTen   = withT.length ? (withT.reduce((s,r)=>s+r._tenureYears,0)/withT.length) : 0;
  const newJ     = withT.filter(r=>r._tenureYears<1).length;
  const newJPct  = total > 0 ? (newJ/total*100) : 0;
  const longT    = withT.filter(r=>r._tenureYears>=5).length;
  const longTPct = total > 0 ? (longT/total*100) : 0;
  const gMap     = countBy(rows,'_gender');
  const gList    = sorted(gMap);

  // ── YTD attrition by segment ──
  const volByCC = {}, involByCC = {}, volByOps = {}, involByOps = {};
  termedYTD.filter(r=>(r._termType||'').toLowerCase()==='voluntary').forEach(r=>{
    volByCC[r._org2]=(volByCC[r._org2]||0)+1;
    volByOps[r._org3]=(volByOps[r._org3]||0)+1;
  });
  termedYTD.filter(r=>(r._termType||'').toLowerCase()==='involuntary').forEach(r=>{
    involByCC[r._org2]=(involByCC[r._org2]||0)+1;
    involByOps[r._org3]=(involByOps[r._org3]||0)+1;
  });

  // ── Compute attrition rates per CC ──
  const ccAttrRates = Object.entries(ccMap).map(([cc, active]) => {
    const vol = volByCC[cc]||0;
    const invol = involByCC[cc]||0;
    const volPct = (active+vol) > 0 ? (vol/(active+vol)*100) : 0;
    const involPct = (active+invol) > 0 ? (invol/(active+invol)*100) : 0;
    return { cc, active, vol, invol, volPct, involPct };
  }).filter(d => d.active >= 5); // ignore tiny teams

  const highVolCC = ccAttrRates.filter(d => d.volPct >= 3).sort((a,b)=>b.volPct-a.volPct);
  const highInvolCC = ccAttrRates.filter(d => d.involPct >= 3).sort((a,b)=>b.involPct-a.involPct);

  // ── Ops segment rates ──
  const opsAttrRates = Object.entries(opsMap).map(([seg, active]) => {
    const vol = volByOps[seg]||0;
    const invol = involByOps[seg]||0;
    const volPct = (active+vol) > 0 ? (vol/(active+vol)*100) : 0;
    const involPct = (active+invol) > 0 ? (invol/(active+invol)*100) : 0;
    return { seg, active, vol, invol, volPct, involPct };
  });
  const highVolOps = [...opsAttrRates].sort((a,b)=>b.volPct-a.volPct)[0];
  const highInvolOps = [...opsAttrRates].sort((a,b)=>b.involPct-a.involPct)[0];

  // ── MoM trend — last 3 months vs prior 3 ──
  const momMap = {};
  termedRows.filter(r=>(r._termType||'').toLowerCase()==='voluntary').forEach(r=>{
    if(!r._termDate) return;
    const d = new Date(r._termDate); if(isNaN(d)) return;
    const ym = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    momMap[ym] = (momMap[ym]||0)+1;
  });
  const allMoms = Object.keys(momMap).sort();
  const last3 = allMoms.slice(-3).reduce((s,m)=>s+(momMap[m]||0),0);
  const prior3 = allMoms.slice(-6,-3).reduce((s,m)=>s+(momMap[m]||0),0);
  const trendDir = last3 > prior3*1.3 ? 'accelerating' : last3 < prior3*0.7 ? 'decelerating' : 'stable';

  // ── Build insights ──
  // Opening: workforce size + composition
  const topReg = sorted(regMap)[0];
  const topCC = sorted(ccMap)[0];
  insights.push(`The workforce stands at <strong>${total.toLocaleString()} active employees</strong> (Regular & PEO, excl. Packaging) with <strong>${attrPct !== null ? attrPct.toFixed(2)+'%' : '—'} voluntary attrition YTD</strong> — ${volCount} voluntary and ${involCount} involuntary exits since January 1.`);

  // Insight 1: Attrition concentration
  if (highVolCC.length > 0) {
    const top = highVolCC[0];
    if (top.volPct >= 5) {
      flags.push(`<strong>${top.cc}</strong> has the highest voluntary attrition at <strong>${top.volPct.toFixed(1)}%</strong> YTD (${top.vol} exits from ${top.active} active) — disproportionate given it represents ${Math.round(top.active/total*100)}% of headcount.`);
    }
  }

  // Insight 2: Involuntary spike
  if (highInvolCC.length > 0) {
    const top = highInvolCC[0];
    if (top.involPct >= 3) {
      flags.push(`Involuntary exits are concentrated in <strong>${top.cc}</strong> (${top.invol} exits, ${top.involPct.toFixed(1)}% rate) — worth reviewing if this reflects a structural reduction or a performance management pattern.`);
    }
  }

  // Insight 3: Tenure risk
  if (newJPct >= 20) {
    flags.push(`<strong>${Math.round(newJPct)}% of employees have under 1 year of tenure</strong> — a significant cohort still in the critical engagement window. If voluntary attrition continues YTD, this group is the most at-risk.`);
  } else if (longTPct >= 35) {
    flags.push(`<strong>${Math.round(longTPct)}% of the workforce has 5+ years of tenure</strong> — a retention signal, but also a succession risk if this cohort approaches retirement age.`);
  }

  // Insight 4: Ops segment with highest vol attrition
  if (highVolOps && highVolOps.volPct >= 3) {
    flags.push(`Within ops segments, <strong>${highVolOps.seg}</strong> shows the highest voluntary attrition at <strong>${highVolOps.volPct.toFixed(1)}%</strong> (${highVolOps.vol} exits YTD). ${highInvolOps && highInvolOps.seg !== highVolOps.seg && highInvolOps.involPct >= 2 ? `Separately, <strong>${highInvolOps.seg}</strong> leads on involuntary exits at ${highInvolOps.involPct.toFixed(1)}%.` : ''}`);
  }

  // Insight 5: Trend direction
  if (prior3 > 0) {
    if (trendDir === 'accelerating') {
      flags.push(`Voluntary exits are <strong>accelerating</strong> — the last 3 months recorded ${last3} exits vs ${prior3} in the prior 3 months, a ${Math.round((last3-prior3)/prior3*100)}% increase. This warrants urgent attention.`);
    } else if (trendDir === 'decelerating') {
      flags.push(`Voluntary exits are <strong>decelerating</strong> — ${last3} exits in the last 3 months vs ${prior3} prior, suggesting retention measures may be taking effect.`);
    }
  }

  // Insight 6: Gender imbalance
  if (gList.length >= 2) {
    const mPct = Math.round((gList.find(g=>g[0].toLowerCase()==='male')?.[1]||0)/total*100);
    const fPct = Math.round((gList.find(g=>g[0].toLowerCase()==='female')?.[1]||0)/total*100);
    if (mPct > 0 && fPct > 0 && Math.abs(mPct-fPct) > 40) {
      flags.push(`Gender composition is <strong>${mPct}% Male / ${fPct}% Female</strong> — a significant imbalance that may warrant review in hiring and promotion pipelines.`);
    }
  }

  // Pick best 2-3 flags as the actual summary
  const selected = flags.slice(0, 3);
  const fullText = [insights[0], ...selected].join(' ');
  document.getElementById('aiText').innerHTML = fullText;
}
function switchTab(tab) {
  ['cc','ops','region'].forEach(t => { document.getElementById('tab-'+t).style.display = t===tab?'':'none'; });
  document.querySelectorAll('.tab').forEach((el,i) => { el.classList.toggle('active', ['cc','ops','region'][i]===tab); });
}

function exportReport() {
  // Capture current rendered report body HTML
  const reportHeader = document.querySelector('.report-header').outerHTML;
  const reportBody = document.querySelector('.report-body').outerHTML;
  const date = new Date().toLocaleDateString('en-GB',{day:'numeric',month:'long',year:'numeric'});

  const scriptClose = '<' + '/script>';
  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>HC Analytics Report · ${date}</title>

<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js">` + `</` + `script>
<style>
:root{--ink:#0f0e0c;--ink-2:#3a3830;--ink-3:#7a7870;--paper:#f9f6f0;--surface:#ffffff;--accent:#1a3a5c;--accent-2:#c84b2f;--border:rgba(15,14,12,0.10);--border-strong:rgba(15,14,12,0.18);--radius:4px;--radius-lg:8px}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:var(--paper);color:var(--ink);-webkit-font-smoothing:antialiased}
.report-header{background:linear-gradient(135deg,#0a2240 0%,#1565c0 55%,#1e88e5 100%);color:#fff;padding:2.5rem 3rem 2rem}
.rh-top{display:flex;align-items:flex-start;justify-content:space-between;flex-wrap:wrap;gap:1rem;margin-bottom:1.5rem}
.rh-title{font-family:Georgia,'Times New Roman',serif;font-size:clamp(1.6rem,3.5vw,2.6rem);line-height:1.15}
.rh-title em{font-style:italic;color:#90caf9}
.kpi-strip{display:flex;flex-wrap:wrap;gap:0;border-top:1px solid rgba(255,255,255,.2);padding-top:1.5rem}
.kpi{padding:0 2.5rem 0 0;min-width:140px}
.kpi-label{font-size:11px;opacity:.7;text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px}
.kpi-value{font-size:2.2rem;font-weight:300;line-height:1}
.report-body{max-width:1100px;margin:0 auto;padding:2.5rem 2rem 4rem}
.ai-narrative{border-left:3px solid var(--accent-2);padding:1.25rem 1.5rem;margin-bottom:2.5rem;background:var(--surface);border-radius:0 var(--radius-lg) var(--radius-lg) 0}
.ai-label{font-size:11px;font-weight:500;letter-spacing:.1em;text-transform:uppercase;color:var(--accent-2);margin-bottom:8px;display:flex;align-items:center;gap:6px}
.ai-pulse{width:7px;height:7px;border-radius:50%;background:var(--accent-2)}
.ai-text{font-size:15px;line-height:1.75;color:var(--ink-2)}
.ai-text strong{color:var(--ink);font-weight:500}
.section-head{display:flex;align-items:baseline;justify-content:space-between;margin:2rem 0 1rem;border-bottom:1px solid var(--border);padding-bottom:8px}
.section-title{font-family:Georgia,'Times New Roman',serif;font-size:18px;color:var(--ink)}
.section-note{font-size:12px;color:var(--ink-3)}
.chart-grid-2{display:grid;grid-template-columns:repeat(auto-fit,minmax(360px,1fr));gap:1.5rem;margin-bottom:1.5rem}
.chart-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-lg);padding:1.25rem 1.5rem 1.5rem}
.cc-title{font-size:14px;font-weight:500;color:var(--ink);margin-bottom:2px}
.cc-sub{font-size:12px;color:var(--ink-3);margin-bottom:14px}
.chart-wrap{position:relative;width:100%}
.legend{display:flex;flex-wrap:wrap;gap:14px;margin-bottom:10px}
.legend-item{display:flex;align-items:center;gap:5px;font-size:12px;color:var(--ink-3)}
.legend-swatch{width:10px;height:10px;border-radius:2px;flex-shrink:0}
.tab-row{display:flex;gap:4px;margin-bottom:1.25rem;flex-wrap:wrap}
.tab{font-size:13px;padding:6px 16px;border-radius:99px;border:1px solid var(--border-strong);background:transparent;color:var(--ink-3);cursor:pointer}
.tab.active{background:var(--accent);color:#fff;border-color:var(--accent)}
.hbar-row{display:flex;align-items:center;gap:10px;margin-bottom:10px}
.hbar-label{font-size:13px;color:var(--ink-3);width:140px;flex-shrink:0;text-align:right}
.hbar-track{flex:1;background:#eee;border-radius:3px;height:22px;overflow:hidden}
.hbar-fill{height:100%;border-radius:3px;display:flex;align-items:center;padding-left:10px}
.hbar-fill span{font-size:12px;font-weight:500;color:#fff}
.hbar-pct{font-size:12px;color:var(--ink-3);width:110px;text-align:right;flex-shrink:0}
.full-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-lg);padding:1.25rem 1.5rem 1.75rem;margin-bottom:1.5rem}
.data-table{width:100%;border-collapse:collapse;font-size:13px;margin-top:8px}
.data-table th{text-align:left;font-size:11px;font-weight:500;text-transform:uppercase;letter-spacing:.07em;color:var(--ink-3);padding:8px 10px;border-bottom:1px solid var(--border)}
.data-table td{padding:9px 10px;border-bottom:1px solid var(--border);color:var(--ink-2)}
.data-table tr:last-child td{border-bottom:none}
.badge{display:inline-block;font-size:11px;padding:2px 8px;border-radius:99px;font-weight:500}
.badge-neu{background:#ebebeb;color:#555}
.report-footer{text-align:center;padding:2rem;font-size:12px;color:var(--ink-3);border-top:1px solid var(--border)}
@media(max-width:640px){.report-header{padding:1.5rem}.report-body{padding:1.5rem 1rem 3rem}.chart-grid-2{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="report-header">
  <div class="rh-top">
    <div><div class="rh-title">Headcount &amp; <em>attrition</em></div></div>
    <div style="font-size:13px;opacity:.7">${date}</div>
  </div>
  <div class="kpi-strip">${document.getElementById('kpiStrip').innerHTML}</div>
</div>
<div class="report-body">
${document.querySelector('.report-body').innerHTML}
</div>
<div class="report-footer">Generated ${date} · HC Analytics Report · Data processed locally — no data was stored or transmitted.</div>
<` + `script>
function switchTab(tab){['cc','ops','region'].forEach(t=>{const el=document.getElementById('tab-'+t);if(el)el.style.display=t===tab?'':'none';});document.querySelectorAll('.tab').forEach((el,i)=>{el.classList.toggle('active',['cc','ops','region'][i]===tab);});}
const BLUE='#1a3a5c',BLUE_MID='#2a5f8c',BLUE_LIGHT='#4a8ab5',BLUE_PALE='#85b7d9',SLATE='#3a4a5c',STEEL='#5a7a8c',MUTED='#8ca5b5',WARN='#c84b2f';
const PALETTE=[BLUE,BLUE_MID,BLUE_LIGHT,SLATE,STEEL,BLUE_PALE,MUTED,'#2a4a6c','#1a4a6c','#3a6a8c'];
const gridC='rgba(15,14,12,0.07)',tickC='#7a7870';
${getChartRebuildScript()}
${scriptClose}
</body>
</html>`;

  const blob = new Blob([html], {type:'text/html'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `HC-Analytics-Report-${date.replace(/\s/g,'-')}.html`;
  a.click();
  URL.revokeObjectURL(url);
}

function getChartRebuildScript() {
  // Serialize current chart data so the exported HTML can rebuild them
  const lines = [];
  Object.entries(charts).forEach(([id, chart]) => {
    if (!chart || !chart.data) return;
    const type = chart.config.type;
    const labels = JSON.stringify(chart.data.labels || []);
    const datasets = JSON.stringify(chart.data.datasets.map(ds => ({
      label: ds.label,
      data: ds.data,
      backgroundColor: ds.backgroundColor,
      borderColor: ds.borderColor,
      borderWidth: ds.borderWidth,
      borderRadius: ds.borderRadius,
      fill: ds.fill,
      tension: ds.tension,
      pointRadius: ds.pointRadius,
      stack: ds.stack,
    })));
    const opts = JSON.stringify(chart.options);
    lines.push(`(function(){const el=document.getElementById(${JSON.stringify(id)});if(!el)return;try{new Chart(el,{type:${JSON.stringify(type)},data:{labels:${labels},datasets:${datasets}},options:${opts}});}catch(e){}})();`);
  });
  return lines.join('\n');
}

function resetReport() {
  document.getElementById('upload-screen').style.display = 'flex';
  document.getElementById('fileInput').value = '';
  Object.values(charts).forEach(c=>c.destroy()); charts = {};
  if (window._genderResizeObs) { window._genderResizeObs.disconnect(); window._genderResizeObs = null; }
}

function loadSampleData() {
  const countries = ['India','India','India','United States','United States','United Kingdom','Germany','Singapore','Australia','UAE','France','Canada'];
  const costCenters = ['Engineering','Operations','Finance','Human Resources','Sales','Legal & Compliance'];
  const opsSegs = ['Core Platform','Customer Success','Risk & Control','Financial Planning','Revenue Growth','People Operations'];
  const products = ['Product Alpha','Product Beta','Product Gamma','Product Delta','Product Epsilon'];
  const genders = ['Male','Female','Female','Male','Non-binary'];
  const now = new Date();
  let tsv = 'Employee_ID\tGender\tCountry\tCompany\tOrg_Level_2\tOrg_Level_3\tOrg_Level_4\tOriginal_Hire_Date\tDate_of_Birth\tEmployment_Status\n';
  for (let i = 1; i <= 600; i++) {
    const hya = Math.random()*12;
    const hd = new Date(now - hya*365.25*24*60*60*1000);
    const ay = 22+Math.random()*38;
    const dob = new Date(now - ay*365.25*24*60*60*1000);
    const fmt = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    const status = Math.random() < 0.12 ? 'Terminated' : 'Active';
    tsv += `EMP${String(i).padStart(4,'0')}\t${genders[i%genders.length]}\t${countries[i%countries.length]}\tAcme Corp\t${costCenters[i%costCenters.length]}\t${opsSegs[i%opsSegs.length]}\t${products[i%products.length]}\t${fmt(hd)}\t${fmt(dob)}\t${status}\n`;
  }
  document.getElementById('upload-screen').style.display = 'none';
  document.getElementById('loading').style.display = 'flex';
  setTimeout(() => renderReport(parseHC(tsv, 'sample-data.tsv')), 800);
}
