async function safeLoad() {
  try {
    const res = await fetch('data.xlsx', { cache: 'no-store' });
    if (!res.ok) throw new Error('data.xlsx not ready');

    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws);
    if (!data.length) throw new Error('empty data');

    const latest20 = data.slice(-20);
    window._latest20 = latest20;

    renderFreq(latest20);
    document.getElementById('status').innerText = '✅ 資料載入成功';
  } catch (e) {
    document.getElementById('status').innerText = '📡 資料初始化中，請稍後重新整理';
    console.warn(e);
  }
}

function renderFreq(rows) {
  const freq = {};
  rows.forEach(r => {
    [r.N1, r.N2, r.N3, r.N4, r.N5, r.N6, r['特別號']]
      .forEach(n => freq[n] = (freq[n] || 0) + 1);
  });

  const tbody = document.getElementById('freq');
  tbody.innerHTML = '';

  Object.entries(freq).sort((a,b)=>b[1]-a[1])
    .forEach(([n,c])=>{
      tbody.innerHTML += `<tr><td>${n}</td><td>${c}</td></tr>`;
    });
}

function downloadExcel() {
  if (!window._latest20) {
    alert('資料未準備好');
    return;
  }

  const wb = XLSX.utils.book_new();
  const freq = {};

  window._latest20.forEach(r => {
    [r.N1, r.N2, r.N3, r.N4, r.N5, r.N6, r['特別號']]
      .forEach(n => freq[n] = (freq[n] || 0) + 1);
  });

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['號碼','出現次數'],
      ...Object.entries(freq).sort((a,b)=>b[1]-a[1])
    ]),
    '號碼統計（含特別號）'
  );

  XLSX.writeFile(wb, 'marksix_last20_with_special.xlsx');
}

async function loadLastUpdateTime() {
  try {
    const res = await fetch('last_update.txt', { cache: 'no-store' });
    if (!res.ok) throw new Error();
    const txt = await res.text();
    document.getElementById('last-update').innerText = txt;
  } catch {
    document.getElementById('last-update').innerText = '⚠️ 尚未更新資料';
  }
}

safeLoad();
loadLastUpdateTime();