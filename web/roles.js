/* Adds logins and roles to the app page without changing app/index.html.
   This only shapes the screen. server.py is what actually enforces the rules. */
(function () {
  const me = window.ALBROST_USER;
  const isOwner = me.role === 'owner';
  document.documentElement.classList.add('role-' + me.role);

  // An expired login goes back to the sign-in page.
  const realFetch = window.fetch.bind(window);
  window.fetch = async (...args) => {
    const response = await realFetch(...args);
    if (response.status === 401) location.replace('/login');
    return response;
  };

  // Name and Log out in the header.
  const header = document.querySelector('header');
  const who = document.createElement('span');
  who.className = 'who';
  who.textContent = 'Signed in: ' + me.name;
  const logout = document.createElement('button');
  logout.className = 'hbtn';
  logout.textContent = 'Log out';
  logout.onclick = async () => {
    await realFetch('/api/logout', {method: 'POST'});
    location.replace('/login');
  };
  header.append(who, logout);

  if (isOwner) setUpOwner(); else setUpManager();

  function setUpManager() {
    window.renderRecords = () => {};
    const ownerOnly = () => alertBox('Not allowed', 'Only the owner can change or delete saved entries.');
    window.editSale = window.editSupp = window.deleteSale = window.deleteSupp = ownerOnly;
    const realGo = window.go;
    window.go = tab => realGo(tab === 'records' ? 'dash' : tab);
    document.getElementById('tab-records').remove();
    document.querySelector('nav button[data-tab="records"]').remove();

    // Same as the app's saveWorkbook, plus a clear message when the server says no.
    window.saveWorkbook = async () => {
      wb.calcProperties = wb.calcProperties || {};
      wb.calcProperties.fullCalcOnLoad = true;
      const body = await wb.xlsx.writeBuffer();
      const response = await api('/api/save', {method: 'POST', headers: {'Content-Type': 'application/octet-stream'}, body});
      if (response.ok) return true;
      if (response.status === 403) {
        await alertBox('Not saved', 'This day or entry is already saved, and only the owner can change it. Please enter a new day or a new entry.');
      } else if (response.status === 423) {
        await alertBox('Excel file is open', 'The Excel file is open on the shop computer. Please ask for it to be closed, then press Save again.<br><br>Your typed entry is still here.');
      } else {
        await alertBox('Could not save', 'Something went wrong while saving. Please try again.');
      }
      return false;
    };
  }

  function setUpOwner() {
    // Logs tab.
    const button = document.createElement('button');
    button.dataset.tab = 'logs';
    button.innerHTML = '<span class="ic">🧾</span>Logs';
    button.onclick = () => window.go('logs');
    document.getElementById('nav').append(button);

    const section = document.createElement('section');
    section.className = 'tab';
    section.id = 'tab-logs';
    section.innerHTML = '<div class="card"><h2>Logs</h2>' +
      '<p class="hint">Who signed in and every entry added, edited or deleted, newest first.</p>' +
      '<div class="scroll"><table id="logsTable"></table></div></div>';
    document.querySelector('main').append(section);

    const realGo = window.go;
    window.go = tab => { realGo(tab); if (tab === 'logs') renderLogs(); };

    // "Last updated" column in All Records.
    const realRender = window.renderRecords;
    window.renderRecords = () => { realRender(); addTimestamps(); };
    const realSave = window.saveWorkbook;
    window.saveWorkbook = async () => {
      const saved = await realSave();
      if (document.getElementById('tab-logs').classList.contains('on')) renderLogs();
      return saved;
    };
  }

  const when = iso => new Date(iso).toLocaleString(undefined,
    {day: 'numeric', month: 'short', year: 'numeric', hour: 'numeric', minute: '2-digit'});

  async function renderLogs() {
    const table = document.getElementById('logsTable');
    const response = await fetch('/api/logs', {cache: 'no-store'});
    if (!response.ok) { table.innerHTML = '<tr><td class="empty">Could not load the logs.</td></tr>'; return; }
    const entries = await response.json();
    table.innerHTML = entries.length
      ? '<thead><tr><th>When</th><th>Who</th><th>What</th><th>Details</th></tr></thead><tbody>' +
        entries.map(e => `<tr><td class="when">${esc(when(e.at))}</td><td><b>${esc(e.user)}</b></td>` +
          `<td>${esc(e.action)}</td><td class="details">${esc(e.details)}</td></tr>`).join('') + '</tbody>'
      : '<tr><td class="empty">Nothing logged yet.</td></tr>';
  }

  function stampCell(stamp) {
    const cell = document.createElement('td');
    cell.className = 'when stamp';
    cell.innerHTML = stamp
      ? `<b>${esc(when(stamp.at))}</b>${esc(stamp.by)}`
      : 'Original data';
    return cell;
  }

  function addColumn(table, stamps) {
    table.querySelectorAll('.stamp').forEach(cell => cell.remove());
    const headRow = table.querySelector('thead tr');
    if (!headRow) return;
    const heading = document.createElement('th');
    heading.className = 'stamp';
    heading.textContent = 'Last updated';
    headRow.insertBefore(heading, headRow.lastElementChild);
    table.querySelectorAll('tbody tr').forEach((row, i) => {
      if (stamps[i] !== undefined) row.insertBefore(stampCell(stamps[i]), row.lastElementChild);
    });
  }

  async function addTimestamps() {
    const response = await fetch('/api/times', {cache: 'no-store'});
    if (!response.ok) return;
    const times = await response.json();
    // Same order as the app's renderRecords.
    const days = D.daily.filter(d => d.has).sort((a, b) => b.date.localeCompare(a.date));
    addColumn(document.getElementById('recSales'), days.map(d => times.sales[d.date] || null));
    const entries = D.supp.slice().reverse();
    addColumn(document.getElementById('recSupp'), entries.map(s =>
      times.supp[[s.date || '', s.supplier, s.item, s.buy, s.paid].join('|')] || null));
  }
})();
