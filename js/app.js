/* AutoLog: maintenance records for every car you own.
   Everything is stored in this browser's localStorage. No server. */
(() => {
  'use strict';

  const STORE_KEY = 'autolog.v1';
  const OIL_INTERVAL_MONTHS = 6;
  const RENOTIFY_DAYS = 7;
  const XLSX_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
  const TESSERACT_URL = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';

  /* ---------- Car fields (drives the form and the spreadsheet backup) ---------- */
  const CAR_GROUPS = [
    { title: 'The car', fields: [
      { key: 'year', label: 'Year', ph: '1969', max: 4, mode: 'numeric' },
      { key: 'make', label: 'Make', ph: 'Ford', required: true },
      { key: 'model', label: 'Model', ph: 'Mustang', required: true },
      { key: 'trim', label: 'Trim level', ph: 'Mach 1' },
      { key: 'plate', label: 'License plate number', ph: 'MACH 69' },
      { key: 'engineNumber', label: 'Engine number', ph: 'Stamped on the engine block' },
      { key: 'vin', label: 'VIN', ph: 'Type it, or scan it from a photo', vin: true },
    ] },
    { title: 'Tires', fields: [
      { key: 'tireType', label: 'Tire type', ph: 'BFGoodrich Radial T/A' },
      { key: 'tireSize', label: 'Tire size', ph: 'P235/60R15' },
      { key: 'tirePressure', label: 'Recommended tire pressure', ph: '32 psi front, 30 psi rear', span: true },
    ] },
    { title: 'Oil', fields: [
      { key: 'oilType', label: 'Oil type', ph: '10W-30 high zinc' },
      { key: 'oilQuantity', label: 'Oil quantity', ph: '5 quarts with filter' },
      { key: 'oilFilterBrand', label: 'Oil filter brand', ph: 'Motorcraft' },
      { key: 'oilFilterPart', label: 'Oil filter part number', ph: 'FL-1A' },
      { key: 'crushRingBrand', label: 'Oil pan crush ring brand', ph: 'Dorman' },
      { key: 'crushRingPart', label: 'Oil pan crush ring part number', ph: '65274' },
    ] },
    { title: 'Filters, wipers and lights', fields: [
      { key: 'engineAirFilter', label: 'Engine air filter number', ph: 'FA-41' },
      { key: 'cabinAirFilter', label: 'Cabin air filter number', ph: 'None on this car' },
      { key: 'wiperBrand', label: 'Wiper blades brand', ph: 'Trico' },
      { key: 'wiperPart', label: 'Wiper blades number', ph: '16 inch, 33-160' },
      { key: 'lightsBrand', label: 'Lights brand', ph: 'Sylvania' },
      { key: 'lightsPart', label: 'Lights part number', ph: 'H6024' },
    ] },
  ];
  const CAR_FIELDS = CAR_GROUPS.flatMap((g) => g.fields);

  const STORES = [
    { name: 'AutoZone', note: 'Parts, batteries and free code reading.', url: 'https://www.autozone.com/' },
    { name: "O'Reilly Auto Parts", note: 'Parts, fluids and loaner tools.', url: 'https://www.oreillyauto.com/' },
    { name: 'Advance Auto Parts', note: 'Parts with same-day store pickup.', url: 'https://shop.advanceautoparts.com/' },
    { name: 'NAPA Auto Parts', note: 'Parts counter favored by repair shops.', url: 'https://www.napaonline.com/' },
    { name: 'RockAuto', note: 'Online only. Wide catalog for older cars.', url: 'https://www.rockauto.com/', online: true },
    { name: 'CJ Pony Parts', note: 'Online only. Mustang and classic Ford parts.', url: 'https://www.cjponyparts.com/', online: true },
  ];

  /* ---------- Helpers ---------- */
  const $ = (id) => document.getElementById(id);
  const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const esc = (v) => String(v ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const money = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });
  const miles = new Intl.NumberFormat('en-US');
  const pad = (n) => String(n).padStart(2, '0');
  const toISO = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const todayISO = () => toISO(new Date());
  const fromISO = (s) => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
  const niceDate = (s) => fromISO(s).toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
  const addMonths = (d, n) => {
    const r = new Date(d.getFullYear(), d.getMonth() + n, 1);
    const last = new Date(r.getFullYear(), r.getMonth() + 1, 0).getDate();
    r.setDate(Math.min(d.getDate(), last));
    return r;
  };
  const dayDiff = (a, b) => Math.round((a - b) / 86400000);
  const looksLikeOilChange = (s) => /\boil\b.*\b(change|service|filter)|\blube\b/i.test(s);
  const carName = (c) => [c.year, c.make, c.model].filter(Boolean).join(' ') || 'Unnamed car';

  let toastTimer;
  function toast(msg) {
    const t = $('toast');
    t.textContent = msg;
    t.classList.add('is-on');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('is-on'), 4200);
  }

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.onload = resolve;
      s.onerror = () => { s.remove(); reject(new Error('Could not load ' + src)); };
      document.head.appendChild(s);
    });
  }

  /* ---------- Storage ---------- */
  let db = { users: [], activeUserId: null, cars: [], entries: [], prefs: {} };
  try {
    const raw = JSON.parse(localStorage.getItem(STORE_KEY));
    if (raw && Array.isArray(raw.users)) db = Object.assign(db, raw);
  } catch { /* start fresh */ }

  function save() {
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify(db));
      return true;
    } catch {
      toast('This browser is out of storage space. Remove a VIN photo, then try again.');
      return false;
    }
  }

  const user = () => db.users.find((u) => u.id === db.activeUserId) || null;
  const prefs = () => {
    const u = user();
    if (!u) return {};
    return (db.prefs[u.id] ||= { activeCarId: null, zip: '', notified: {} });
  };
  const myCars = () => db.cars.filter((c) => c.userId === db.activeUserId);
  const activeCar = () => {
    const cars = myCars();
    return cars.find((c) => c.id === prefs().activeCarId) || cars[0] || null;
  };
  const carEntries = (carId) => db.entries.filter((e) => e.carId === carId);

  /* ---------- View state ---------- */
  const view = { sortKey: 'date', sortDir: 'desc', editingId: null, oilTouched: false };
  let pendingVinPhoto = null; // photo chosen in the car form, before saving

  /* ---------- Rendering ---------- */
  function renderAll() {
    renderHeader();
    renderDetails();
    renderReminder();
    renderLog();
    renderStores();
  }

  function renderHeader() {
    const u = user();
    $('profileName').textContent = u ? u.name : 'Profile';
    const cars = myCars();
    const sel = $('carSelect');
    sel.hidden = cars.length === 0;
    const car = activeCar();
    sel.innerHTML = cars.map((c) => `<option value="${esc(c.id)}"${car && c.id === car.id ? ' selected' : ''}>${esc(carName(c))}${c.plate ? ' (' + esc(c.plate) + ')' : ''}</option>`).join('');
  }

  const val = (v) => (v ? `<dd>${esc(v)}</dd>` : '<dd class="unset">Not set</dd>');
  const pair = (a, b) => [a, b].filter(Boolean).join(' ');

  function renderDetails() {
    const car = activeCar();
    const box = $('carDetails');
    $('detailsActions').hidden = !car;
    if (!car) {
      box.innerHTML = `<p class="empty">No cars yet. <button type="button" class="btn btn-primary" id="emptyAddCar">Add your first car</button></p>`;
      $('emptyAddCar').addEventListener('click', () => openCarDialog(null));
      return;
    }
    const groups = [
      ['Tires', [['Type', car.tireType], ['Size', car.tireSize], ['Recommended pressure', car.tirePressure]]],
      ['Oil', [['Type', car.oilType], ['Quantity', car.oilQuantity], ['Oil filter', pair(car.oilFilterBrand, car.oilFilterPart)], ['Oil pan crush ring', pair(car.crushRingBrand, car.crushRingPart)]]],
      ['Filters, wipers and lights', [['Engine air filter', car.engineAirFilter], ['Cabin air filter', car.cabinAirFilter], ['Wiper blades', pair(car.wiperBrand, car.wiperPart)], ['Lights', pair(car.lightsBrand, car.lightsPart)]]],
    ];
    box.innerHTML = `
      <div class="plate-card">
        <div class="car-ident">
          <div>
            <p class="car-name">${esc(carName(car))}</p>
            ${car.trim ? `<p class="car-trim">${esc(car.trim)}</p>` : ''}
          </div>
          ${car.plate ? `<div class="license"><small>License plate</small><strong>${esc(car.plate)}</strong></div>` : ''}
          <dl class="ident-facts">
            <div>
              <dt>VIN</dt>
              <dd class="vin-row${car.vin ? '' : ' unset'}">
                ${car.vin ? `<span class="vin-value">${esc(car.vin)}</span>` : 'Not set'}
                ${car.vinPhoto ? `<button type="button" class="vin-thumb" id="vinThumb" aria-label="View VIN photo"><img src="${esc(car.vinPhoto)}" alt=""></button>` : ''}
              </dd>
            </div>
            <div><dt>Engine number</dt>${val(car.engineNumber)}</div>
          </dl>
        </div>
        <div class="spec-groups">
          ${groups.map(([title, rows]) => `
            <div class="spec-group">
              <h3>${title}</h3>
              <dl>${rows.map(([k, v]) => `<div><dt>${k}</dt>${val(v)}</div>`).join('')}</dl>
            </div>`).join('')}
        </div>
      </div>`;
    const thumb = $('vinThumb');
    if (thumb) thumb.addEventListener('click', () => { $('photoFull').src = car.vinPhoto; $('photoDialog').showModal(); });
  }

  /* ---------- Oil change reminder ---------- */
  function oilStatus(car) {
    const oil = carEntries(car.id).filter((e) => e.oilChange).map((e) => e.date).sort();
    if (!oil.length) return null;
    const last = fromISO(oil[oil.length - 1]);
    const due = addMonths(last, OIL_INTERVAL_MONTHS);
    const today = fromISO(todayISO());
    const daysLeft = dayDiff(due, today);
    const span = Math.max(1, dayDiff(due, last));
    return { last, due, daysLeft, fraction: Math.min(1, Math.max(0, dayDiff(today, last) / span)) };
  }

  function renderReminder() {
    const car = activeCar();
    $('reminder').hidden = !car;
    if (!car) return;
    const st = oilStatus(car);
    const line = $('reminderLine'), gauge = $('reminderGauge'), note = $('reminderNote');
    const isDue = !!st && st.daysLeft <= 0;
    line.classList.toggle('is-due', isDue);
    gauge.classList.toggle('is-due', isDue);
    if (!st) {
      line.textContent = 'No oil change logged yet.';
      gauge.firstElementChild.style.width = '0%';
      gauge.setAttribute('aria-label', 'No oil change logged');
      note.textContent = `Log an oil change for the ${carName(car)} to start the ${OIL_INTERVAL_MONTHS}-month countdown.`;
    } else {
      const dueTxt = niceDate(toISO(st.due));
      if (st.daysLeft > 0) line.textContent = `Next oil change due ${dueTxt}, in ${st.daysLeft} day${st.daysLeft === 1 ? '' : 's'}.`;
      else if (st.daysLeft === 0) line.textContent = 'Oil change due today.';
      else line.textContent = `Oil change overdue by ${-st.daysLeft} day${st.daysLeft === -1 ? '' : 's'}.`;
      gauge.firstElementChild.style.width = Math.round(st.fraction * 100) + '%';
      gauge.setAttribute('aria-label', `${Math.round(st.fraction * 100)}% of the ${OIL_INTERVAL_MONTHS}-month interval used`);
      note.textContent = `Last oil change for the ${carName(car)}: ${niceDate(toISO(st.last))}.`;
    }
    $('calendarBtn').disabled = false;
    const nb = $('notifyBtn');
    if (!('Notification' in window)) { nb.hidden = true; }
    else if (Notification.permission === 'granted') { nb.textContent = 'Notifications are on'; nb.classList.add('is-on'); nb.disabled = true; nb.style.cursor = 'default'; }
    else if (Notification.permission === 'denied') { nb.textContent = 'Notifications blocked in browser settings'; nb.disabled = true; nb.style.cursor = 'default'; }
  }

  let swReg = null;
  if ('serviceWorker' in navigator && location.protocol === 'https:' || location.hostname === 'localhost') {
    navigator.serviceWorker?.register('./sw.js').then((r) => { swReg = r; }).catch(() => {});
  }

  async function notify(title, body, tag) {
    const opts = { body, tag, icon: 'assets/favicon.svg' };
    try {
      const reg = swReg || (navigator.serviceWorker && await navigator.serviceWorker.getRegistration());
      if (reg && reg.showNotification) return await reg.showNotification(title, opts);
      new Notification(title, opts);
    } catch { /* notification support varies; the on-page reminder still shows */ }
  }

  function checkReminders() {
    if (!user() || !('Notification' in window) || Notification.permission !== 'granted') return;
    const p = prefs();
    p.notified ||= {};
    const today = fromISO(todayISO());
    let changed = false;
    for (const car of myCars()) {
      const st = oilStatus(car);
      if (!st || st.daysLeft > 0) continue;
      const last = p.notified[car.id];
      if (last && dayDiff(today, fromISO(last)) < RENOTIFY_DAYS) continue;
      const body = st.daysLeft === 0
        ? `It has been ${OIL_INTERVAL_MONTHS} months since the last oil change.`
        : `The oil change was due ${niceDate(toISO(st.due))}.`;
      notify(`Oil change due: ${carName(car)}`, body, 'oil-' + car.id);
      p.notified[car.id] = todayISO();
      changed = true;
    }
    if (changed) save();
  }

  function downloadCalendar() {
    const car = activeCar();
    if (!car) return;
    const st = oilStatus(car);
    const start = st ? st.due : addMonths(new Date(), OIL_INTERVAL_MONTHS);
    const ymd = (d) => `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
    const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 1);
    const text = (s) => String(s).replace(/([\\;,])/g, '\\$1').replace(/\n/g, '\\n');
    const now = new Date();
    const stamp = `${now.getUTCFullYear()}${pad(now.getUTCMonth() + 1)}${pad(now.getUTCDate())}T${pad(now.getUTCHours())}${pad(now.getUTCMinutes())}${pad(now.getUTCSeconds())}Z`;
    const oilBits = [car.oilType, car.oilQuantity].filter(Boolean).join(', ');
    const filter = pair(car.oilFilterBrand, car.oilFilterPart);
    const desc = ['Time for an oil change.', oilBits && `Oil: ${oilBits}.`, filter && `Filter: ${filter}.`, 'Log it in AutoLog when it is done.'].filter(Boolean).join(' ');
    const ics = [
      'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//AutoLog//Oil change reminder//EN', 'CALSCALE:GREGORIAN',
      'BEGIN:VEVENT',
      `UID:oil-${car.id}@autolog`,
      `DTSTAMP:${stamp}`,
      `DTSTART;VALUE=DATE:${ymd(start)}`,
      `DTEND;VALUE=DATE:${ymd(end)}`,
      `RRULE:FREQ=MONTHLY;INTERVAL=${OIL_INTERVAL_MONTHS}`,
      `SUMMARY:${text('Oil change due: ' + carName(car))}`,
      `DESCRIPTION:${text(desc)}`,
      'BEGIN:VALARM', 'ACTION:DISPLAY', `DESCRIPTION:${text('Oil change due: ' + carName(car))}`, 'TRIGGER:PT9H', 'END:VALARM',
      'END:VEVENT', 'END:VCALENDAR',
    ].join('\r\n');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([ics], { type: 'text/calendar' }));
    a.download = `oil-change-${carName(car).toLowerCase().replace(/[^a-z0-9]+/g, '-')}.ics`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    toast(`Calendar file downloaded. Open it to add the reminder, first due ${niceDate(toISO(start))}.`);
  }

  /* ---------- Maintenance log ---------- */
  const chrono = (a, b) => a.date.localeCompare(b.date) || (a.mileage ?? 0) - (b.mileage ?? 0) || (a.created ?? 0) - (b.created ?? 0);

  function currentFilters() {
    return {
      q: $('fSearch').value.trim().toLowerCase(),
      loc: $('fLoc').value,
      from: $('fFrom').value,
      to: $('fTo').value,
      oil: $('fOil').checked,
    };
  }

  function renderOdometer(total) {
    const [whole, cents] = total.toFixed(2).split('.');
    const digits = whole.padStart(5, '0');
    $('odometer').innerHTML = `<span class="odo-unit">$</span>${[...digits].map((d) => `<span>${d}</span>`).join('')}<span class="odo-unit">.</span>${[...cents].map((d) => `<span class="odo-cents">${d}</span>`).join('')}`;
    $('odometer').setAttribute('aria-label', money.format(total));
  }

  function renderLog() {
    const car = activeCar();
    $('log').hidden = !car;
    if (!car) return;
    const all = carEntries(car.id);

    // Location filter options and suggestions
    const locs = [...new Set(all.map((e) => e.location).filter(Boolean))].sort((a, b) => a.localeCompare(b));
    const locSel = $('fLoc');
    const keep = locs.includes(locSel.value) ? locSel.value : '';
    locSel.innerHTML = '<option value="">All locations</option>' + locs.map((l) => `<option${l === keep ? ' selected' : ''}>${esc(l)}</option>`).join('');
    $('locList').innerHTML = locs.map((l) => `<option value="${esc(l)}"></option>`).join('');

    const f = currentFilters();
    const filtered = all.filter((e) =>
      (!f.q || (e.description + ' ' + (e.location || '')).toLowerCase().includes(f.q)) &&
      (!f.loc || e.location === f.loc) &&
      (!f.from || e.date >= f.from) &&
      (!f.to || e.date <= f.to) &&
      (!f.oil || e.oilChange));
    const isFiltered = !!(f.q || f.loc || f.from || f.to || f.oil);
    $('clearFiltersBtn').hidden = !isFiltered;

    // Running total, always in date order
    const running = new Map();
    let sum = 0;
    [...filtered].sort(chrono).forEach((e) => { sum += e.cost || 0; running.set(e.id, sum); });

    const dir = view.sortDir === 'asc' ? 1 : -1;
    const k = view.sortKey;
    const rows = [...filtered].sort((a, b) => {
      let c;
      if (k === 'date') c = chrono(a, b);
      else if (k === 'cost' || k === 'mileage') c = (a[k] ?? -1) - (b[k] ?? -1);
      else c = String(a[k] || '').localeCompare(String(b[k] || ''), undefined, { sensitivity: 'base' });
      return (c || chrono(a, b)) * dir;
    });

    const total = all.reduce((s, e) => s + (e.cost || 0), 0);
    renderOdometer(total);
    $('odoLabel').textContent = `Total spent on the ${carName(car)}`;

    $('logSummary').textContent = !all.length ? ''
      : isFiltered ? `Showing ${filtered.length} of ${all.length} entries, ${money.format(sum)} of ${money.format(total)}.`
      : `${all.length} entr${all.length === 1 ? 'y' : 'ies'}.`;

    $('logBody').innerHTML = rows.map((e) => `
      <tr>
        <td class="date">${esc(niceDate(e.date))}</td>
        <td class="desc">${esc(e.description)}${e.oilChange ? '<span class="oil-tag">Oil change</span>' : ''}</td>
        <td data-label="Location" class="${e.location ? '' : 'is-blank'}">${esc(e.location || '')}</td>
        <td data-label="Mileage" class="num ${e.mileage == null ? 'is-blank' : ''}">${e.mileage == null ? '' : miles.format(e.mileage)}</td>
        <td data-label="Cost" class="num">${money.format(e.cost || 0)}</td>
        <td data-label="Running total" class="num running">${money.format(running.get(e.id))}</td>
        <td><div class="row-actions">
          <button type="button" data-edit="${esc(e.id)}" aria-label="Edit ${esc(e.description)}, ${esc(niceDate(e.date))}">Edit</button>
          <button type="button" data-del="${esc(e.id)}" aria-label="Delete ${esc(e.description)}, ${esc(niceDate(e.date))}">Delete</button>
        </div></td>
      </tr>`).join('');

    $('logTable').hidden = rows.length === 0;
    const empty = $('logEmpty');
    empty.hidden = rows.length > 0;
    empty.textContent = all.length
      ? 'No entries match these filters.'
      : `Nothing logged for the ${carName(car)} yet. Add the first job with the form above.`;

    document.querySelectorAll('#logTable th').forEach((th) => {
      const b = th.querySelector('[data-sort]');
      if (!b) return;
      if (b.dataset.sort === k) th.setAttribute('aria-sort', view.sortDir === 'asc' ? 'ascending' : 'descending');
      else th.removeAttribute('aria-sort');
    });
    const sortVal = `${k}:${view.sortDir}`;
    const fs = $('fSort');
    if ([...fs.options].some((o) => o.value === sortVal)) fs.value = sortVal;
  }

  function resetEntryForm() {
    view.editingId = null;
    view.oilTouched = false;
    $('entryForm').reset();
    $('eDate').value = todayISO();
    $('entryForm').classList.remove('is-editing');
    $('entryFormTitle').textContent = 'Log completed work';
    $('saveEntryBtn').textContent = 'Add to log';
    $('cancelEditBtn').hidden = true;
    $('entryError').hidden = true;
    ['eDate', 'eDesc', 'eMiles', 'eCost'].forEach((id) => $(id).removeAttribute('aria-invalid'));
  }

  function submitEntry(ev) {
    ev.preventDefault();
    const car = activeCar();
    if (!car) return;
    const date = $('eDate').value;
    const description = $('eDesc').value.trim();
    const milesRaw = $('eMiles').value.trim();
    const costRaw = $('eCost').value.trim();
    const mileage = milesRaw === '' ? null : Number(milesRaw);
    const cost = costRaw === '' ? 0 : Number(costRaw);

    const problems = [];
    const flag = (id, bad, msg) => { $(id).toggleAttribute('aria-invalid', false); if (bad) { $(id).setAttribute('aria-invalid', 'true'); problems.push(msg); } };
    flag('eDate', !/^\d{4}-\d{2}-\d{2}$/.test(date), 'Enter the date the work was completed.');
    flag('eDesc', !description, 'Describe the repair or maintenance.');
    flag('eMiles', mileage !== null && (!Number.isFinite(mileage) || mileage < 0), 'Mileage must be a number, 0 or higher.');
    flag('eCost', !Number.isFinite(cost) || cost < 0, 'Cost must be a number, 0 or higher.');
    const err = $('entryError');
    if (problems.length) {
      err.textContent = problems.join(' ');
      err.hidden = false;
      document.querySelector('#entryForm [aria-invalid="true"]').focus();
      return;
    }
    err.hidden = true;

    const data = { date, description, location: $('eLoc').value.trim(), mileage: mileage === null ? null : Math.round(mileage), cost: Math.round(cost * 100) / 100, oilChange: $('eOil').checked };
    const editing = view.editingId && db.entries.find((e) => e.id === view.editingId);
    if (editing) Object.assign(editing, data);
    else db.entries.push({ id: uid(), carId: car.id, created: Date.now(), ...data });
    if (data.oilChange) delete prefs().notified?.[car.id];
    if (!save()) return;
    resetEntryForm();
    renderLog();
    renderReminder();
    toast(editing ? 'Changes saved.' : 'Added to log.');
  }

  function startEdit(id) {
    const e = db.entries.find((x) => x.id === id);
    if (!e) return;
    view.editingId = id;
    view.oilTouched = true;
    $('eDate').value = e.date;
    $('eDesc').value = e.description;
    $('eLoc').value = e.location || '';
    $('eMiles').value = e.mileage ?? '';
    $('eCost').value = e.cost ?? '';
    $('eOil').checked = !!e.oilChange;
    $('entryForm').classList.add('is-editing');
    $('entryFormTitle').textContent = 'Edit log entry';
    $('saveEntryBtn').textContent = 'Save changes';
    $('cancelEditBtn').hidden = false;
    $('entryForm').scrollIntoView({ block: 'center' });
    $('eDesc').focus({ preventScroll: true });
  }

  function deleteEntry(id) {
    const e = db.entries.find((x) => x.id === id);
    if (!e || !confirm(`Delete "${e.description}" from ${niceDate(e.date)}?`)) return;
    db.entries = db.entries.filter((x) => x.id !== id);
    if (view.editingId === id) resetEntryForm();
    save();
    renderLog();
    renderReminder();
    toast('Entry deleted.');
  }

  /* ---------- Parts stores ---------- */
  function renderStores() {
    const where = (prefs().zip || '').trim();
    const near = where ? `near ${where}` : 'near me';
    const maps = (q) => 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(`${q} ${near}`);
    if ($('zipInput').value !== where && document.activeElement !== $('zipInput')) $('zipInput').value = where;
    $('allStoresLink').href = maps('auto parts store');
    $('storeList').innerHTML = STORES.map((s) => `
      <li>
        <div><h3>${esc(s.name)}</h3><p>${esc(s.note)}</p></div>
        <div class="store-links">
          ${s.online ? '' : `<a href="${esc(maps(s.name))}" target="_blank" rel="noopener">Find the nearest store<span class="visually-hidden">: ${esc(s.name)}</span></a>`}
          <a href="${esc(s.url)}" target="_blank" rel="noopener">${s.online ? 'Shop online' : 'Website'}<span class="visually-hidden">: ${esc(s.name)}</span></a>
        </div>
      </li>`).join('');
  }

  /* ---------- Car dialog ---------- */
  function buildCarForm() {
    $('carFormBody').innerHTML = CAR_GROUPS.map((g) => `
      <fieldset>
        <legend>${g.title}</legend>
        <div class="form-grid">
          ${g.fields.map((f) => f.vin ? `
            <div class="field span-2">
              <label for="c_vin">VIN</label>
              <input type="text" id="c_vin" maxlength="20" autocapitalize="characters" autocomplete="off" spellcheck="false" placeholder="${esc(f.ph)}">
              <p class="hint" id="vinHint" aria-live="polite"></p>
              <div class="vin-tools">
                <button type="button" class="btn btn-outline" id="vinScanBtn">Scan VIN from a photo</button>
                <input type="file" id="vinFile" accept="image/*" capture="environment" hidden>
                <div class="vin-preview" id="vinPreview" hidden>
                  <img id="vinPreviewImg" alt="VIN photo">
                  <button type="button" class="btn btn-quiet" id="vinPhotoRemove">Remove photo</button>
                </div>
              </div>
            </div>` : `
            <div class="field${f.span ? ' span-2' : ''}">
              <label for="c_${f.key}">${f.label}${f.required ? '' : ''}</label>
              <input type="text" id="c_${f.key}" maxlength="${f.max || 80}" placeholder="${esc(f.ph)}"${f.mode ? ` inputmode="${f.mode}"` : ''}${f.required ? ' required' : ''}>
            </div>`).join('')}
        </div>
      </fieldset>`).join('');

    $('c_vin').addEventListener('input', () => {
      const el = $('c_vin');
      const clean = el.value.toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (el.value !== clean) el.value = clean;
      showVinHint();
    });
    $('vinScanBtn').addEventListener('click', () => $('vinFile').click());
    $('vinFile').addEventListener('change', scanVin);
    $('vinPhotoRemove').addEventListener('click', () => { pendingVinPhoto = null; showVinPhoto(); });
  }

  function vinCheckDigitOk(vin) {
    const map = { A: 1, B: 2, C: 3, D: 4, E: 5, F: 6, G: 7, H: 8, J: 1, K: 2, L: 3, M: 4, N: 5, P: 7, R: 9, S: 2, T: 3, U: 4, V: 5, W: 6, X: 7, Y: 8, Z: 9 };
    const w = [8, 7, 6, 5, 4, 3, 2, 10, 0, 9, 8, 7, 6, 5, 4, 3, 2];
    let sum = 0;
    for (let i = 0; i < 17; i++) {
      const ch = vin[i];
      const v = /\d/.test(ch) ? Number(ch) : map[ch];
      if (v === undefined) return false;
      sum += v * w[i];
    }
    const r = sum % 11;
    return vin[8] === (r === 10 ? 'X' : String(r));
  }

  function showVinHint(msg, cls) {
    const hint = $('vinHint');
    hint.className = 'hint';
    if (msg) { hint.textContent = msg; if (cls) hint.classList.add(cls); return; }
    const v = $('c_vin').value;
    if (!v) hint.textContent = '';
    else if (v.length === 17) {
      if (/[IOQ]/.test(v)) { hint.textContent = 'A 17-character VIN never contains I, O or Q. Check for 1 or 0 instead.'; hint.classList.add('is-warn'); }
      else if (vinCheckDigitOk(v)) { hint.textContent = '17 characters and the check digit matches.'; hint.classList.add('is-good'); }
      else hint.textContent = '17 characters. The check digit does not match, so compare it with the car once more (some non-US VINs have no check digit).';
    } else if (v.length < 17) hint.textContent = `${v.length} characters. Cars built before 1981 have shorter VINs; later cars have 17.`;
    else { hint.textContent = `${v.length} characters. A VIN has 17 at most.`; hint.classList.add('is-warn'); }
  }

  function showVinPhoto() {
    $('vinPreview').hidden = !pendingVinPhoto;
    if (pendingVinPhoto) $('vinPreviewImg').src = pendingVinPhoto;
  }

  function shrinkImage(file, maxSide = 900) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      const url = URL.createObjectURL(file);
      img.onload = () => {
        const k = Math.min(1, maxSide / Math.max(img.width, img.height));
        const c = document.createElement('canvas');
        c.width = Math.round(img.width * k);
        c.height = Math.round(img.height * k);
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        URL.revokeObjectURL(url);
        resolve(c.toDataURL('image/jpeg', 0.72));
      };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('unreadable image')); };
      img.src = url;
    });
  }

  function pickVin(text) {
    const tokens = text.toUpperCase().split(/[\s|:;,]+/).map((t) => t.replace(/[^A-Z0-9]/g, '')).filter((t) => t.length >= 9 && t.length <= 17 && /\d/.test(t) && /[A-Z]/.test(t));
    if (!tokens.length) return '';
    const fix = (t) => (t.length === 17 ? t.replace(/[OQ]/g, '0').replace(/I/g, '1') : t);
    const fixed = tokens.map(fix);
    return fixed.find((t) => t.length === 17 && vinCheckDigitOk(t)) || fixed.find((t) => t.length === 17) || fixed.sort((a, b) => b.length - a.length)[0];
  }

  async function scanVin(ev) {
    const file = ev.target.files[0];
    ev.target.value = '';
    if (!file) return;
    const btn = $('vinScanBtn');
    btn.disabled = true;
    try {
      pendingVinPhoto = await shrinkImage(file);
      showVinPhoto();
    } catch {
      showVinHint('That file could not be opened as a photo. Try a JPEG or PNG.', 'is-warn');
      btn.disabled = false;
      return;
    }
    showVinHint('Reading the VIN from your photo. This can take a few seconds.');
    try {
      if (!window.Tesseract) await loadScript(TESSERACT_URL);
      const worker = await window.Tesseract.createWorker('eng');
      await worker.setParameters({ tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ' });
      const { data } = await worker.recognize(file);
      await worker.terminate();
      const found = pickVin(data.text || '');
      if (found) {
        $('c_vin').value = found;
        showVinHint();
        $('vinHint').textContent = 'Read from photo: check it against the car. ' + $('vinHint').textContent;
        $('c_vin').focus();
      } else {
        showVinHint('No VIN could be read from that photo. The photo is attached, so type the VIN in, or retake it closer and in better light.', 'is-warn');
      }
    } catch {
      showVinHint('The photo reader could not load (it needs an internet connection). The photo is attached, so type the VIN in.', 'is-warn');
    }
    btn.disabled = false;
  }

  let editingCarId = null;
  function openCarDialog(car) {
    if (!user()) { openProfiles(); return; }
    editingCarId = car ? car.id : null;
    $('carDialogTitle').textContent = car ? 'Edit car details' : 'Add a car';
    $('saveCarBtn').textContent = car ? 'Save changes' : 'Save car';
    CAR_FIELDS.forEach((f) => { $('c_' + f.key).value = car ? car[f.key] || '' : ''; $('c_' + f.key).removeAttribute('aria-invalid'); });
    pendingVinPhoto = car ? car.vinPhoto || null : null;
    showVinPhoto();
    showVinHint();
    $('carError').hidden = true;
    $('carDialog').showModal();
    carSnapshot = carFormState();
  }

  /* Unsaved-changes guard for the car form */
  let carSnapshot = '';
  const carFormState = () => JSON.stringify([CAR_FIELDS.map((f) => $('c_' + f.key).value.trim()), pendingVinPhoto || '']);

  function requestCloseCarDialog() {
    if (carFormState() === carSnapshot) { $('carDialog').close(); return; }
    const d = $('unsavedDialog');
    if (!d.open) d.showModal();
    $('unsavedNo').focus();
  }

  function submitCar(ev) {
    ev.preventDefault();
    const data = {};
    CAR_FIELDS.forEach((f) => { data[f.key] = $('c_' + f.key).value.trim(); });
    const missing = ['make', 'model'].filter((k) => !data[k]);
    ['make', 'model'].forEach((k) => $('c_' + k).toggleAttribute('aria-invalid', false));
    if (missing.length) {
      missing.forEach((k) => $('c_' + k).setAttribute('aria-invalid', 'true'));
      $('carError').textContent = 'Enter the make and model so you can tell your cars apart.';
      $('carError').hidden = false;
      $('c_' + missing[0]).focus();
      return;
    }
    data.vinPhoto = pendingVinPhoto || '';
    const before = JSON.stringify(db);
    let car = db.cars.find((c) => c.id === editingCarId);
    if (car) Object.assign(car, data);
    else { car = { id: uid(), userId: db.activeUserId, ...data }; db.cars.push(car); }
    prefs().activeCarId = car.id;
    if (!save()) { db = JSON.parse(before); return; }
    $('carDialog').close();
    resetEntryForm();
    renderAll();
    toast(editingCarId ? 'Changes saved.' : `${carName(car)} added.`);
  }

  function deleteCar() {
    const car = activeCar();
    if (!car) return;
    const n = carEntries(car.id).length;
    if (!confirm(`Delete the ${carName(car)} and its ${n} log entr${n === 1 ? 'y' : 'ies'}? Download a backup first if you might want them back.`)) return;
    db.cars = db.cars.filter((c) => c.id !== car.id);
    db.entries = db.entries.filter((e) => e.carId !== car.id);
    prefs().activeCarId = null;
    save();
    resetEntryForm();
    renderAll();
    toast('Car deleted.');
  }

  /* ---------- Profiles ---------- */
  function openProfiles() {
    renderProfiles();
    const d = $('profileDialog');
    if (!d.open) d.showModal();
    if (!db.users.length) $('newProfileName').focus();
  }

  function renderProfiles() {
    const first = db.users.length === 0;
    $('profileDialogTitle').textContent = first ? 'Welcome to AutoLog' : 'Profiles';
    $('profileIntro').textContent = first
      ? 'Start by creating your profile. Your cars and log are kept under it, in this browser on this device.'
      : 'Each person gets their own cars and log. Profiles are kept in this browser, so anyone using this device can open them.';
    $('profileClose').hidden = first;
    $('profileList').hidden = first;
    $('profileList').innerHTML = db.users.map((u) => {
      const n = db.cars.filter((c) => c.userId === u.id).length;
      const active = u.id === db.activeUserId;
      return `<li>
        <span class="who">${esc(u.name)}<small>${n} car${n === 1 ? '' : 's'}${active ? ', in use now' : ''}</small></span>
        <span class="acts">
          ${active ? '' : `<button type="button" class="btn btn-outline" data-switch="${esc(u.id)}">Switch to ${esc(u.name)}</button>`}
          <button type="button" class="btn btn-quiet" data-delprofile="${esc(u.id)}" aria-label="Delete profile ${esc(u.name)}">Delete</button>
        </span>
      </li>`;
    }).join('');
  }

  function switchUser(id) {
    db.activeUserId = id;
    save();
    resetEntryForm();
    clearFilters(false);
    renderAll();
    checkReminders();
  }

  function submitProfile(ev) {
    ev.preventDefault();
    const name = $('newProfileName').value.trim();
    const err = $('profileError');
    if (!name) { err.textContent = 'Enter a name for the profile.'; err.hidden = false; $('newProfileName').focus(); return; }
    if (db.users.some((u) => u.name.toLowerCase() === name.toLowerCase())) { err.textContent = `There is already a profile called ${name}.`; err.hidden = false; return; }
    err.hidden = true;
    const u = { id: uid(), name };
    db.users.push(u);
    $('newProfileName').value = '';
    switchUser(u.id);
    $('profileDialog').close();
    toast(`Profile created. Welcome, ${name}.`);
  }

  function deleteProfile(id) {
    const u = db.users.find((x) => x.id === id);
    if (!u) return;
    const cars = db.cars.filter((c) => c.userId === id);
    if (!confirm(`Delete ${u.name}'s profile, including ${cars.length} car${cars.length === 1 ? '' : 's'} and all of their log entries? This cannot be undone.`)) return;
    const ids = new Set(cars.map((c) => c.id));
    db.entries = db.entries.filter((e) => !ids.has(e.carId));
    db.cars = db.cars.filter((c) => c.userId !== id);
    db.users = db.users.filter((x) => x.id !== id);
    delete db.prefs[id];
    if (db.activeUserId === id) db.activeUserId = db.users[0]?.id || null;
    save();
    resetEntryForm();
    renderAll();
    renderProfiles();
  }

  /* ---------- Backup and restore ---------- */
  const LOG_COLS = ['Car ID', 'Car', 'Date', 'Work done', 'Location', 'Mileage', 'Cost', 'Oil change', 'Running total'];

  async function ensureXlsx() {
    if (!window.XLSX) await loadScript(XLSX_URL);
    return window.XLSX;
  }

  async function downloadBackup() {
    const u = user();
    if (!u) return openProfiles();
    const btn = $('backupBtn');
    btn.disabled = true;
    try {
      const XLSX = await ensureXlsx();
      const cars = myCars();
      const carRows = cars.map((c) => {
        const row = { 'Car ID': c.id };
        CAR_FIELDS.forEach((f) => { row[f.label] = c[f.key] || ''; });
        return row;
      });
      const logRows = [];
      cars.forEach((c) => {
        let sum = 0;
        carEntries(c.id).sort(chrono).forEach((e) => {
          sum += e.cost || 0;
          logRows.push({ 'Car ID': c.id, 'Car': carName(c), 'Date': e.date, 'Work done': e.description, 'Location': e.location || '', 'Mileage': e.mileage ?? '', 'Cost': e.cost || 0, 'Oil change': e.oilChange ? 'Yes' : 'No', 'Running total': Math.round(sum * 100) / 100 });
        });
      });
      const wb = XLSX.utils.book_new();
      const ws1 = XLSX.utils.json_to_sheet(carRows, { header: ['Car ID', ...CAR_FIELDS.map((f) => f.label)] });
      ws1['!cols'] = [{ wch: 16 }, ...CAR_FIELDS.map((f) => ({ wch: Math.max(14, f.label.length + 2) }))];
      const ws2 = XLSX.utils.json_to_sheet(logRows, { header: LOG_COLS });
      ws2['!cols'] = [{ wch: 16 }, { wch: 26 }, { wch: 12 }, { wch: 44 }, { wch: 26 }, { wch: 10 }, { wch: 10 }, { wch: 11 }, { wch: 14 }];
      const ws3 = XLSX.utils.aoa_to_sheet([
        ['AutoLog backup'], ['Profile', u.name], ['Created', todayISO()],
        [], ['To restore: open AutoLog, choose "Restore from backup" and pick this file.'],
        ['You can edit rows here first. Keep the sheet names and column headings as they are.'],
      ]);
      ws3['!cols'] = [{ wch: 16 }, { wch: 60 }];
      XLSX.utils.book_append_sheet(wb, ws1, 'Cars');
      XLSX.utils.book_append_sheet(wb, ws2, 'Maintenance log');
      XLSX.utils.book_append_sheet(wb, ws3, 'About');
      XLSX.writeFile(wb, `autolog-backup-${todayISO()}.xlsx`);
      toast(`Backup downloaded: ${cars.length} car${cars.length === 1 ? '' : 's'}, ${logRows.length} log entr${logRows.length === 1 ? 'y' : 'ies'}.`);
    } catch {
      toast('The backup could not be created. Check your internet connection and try again.');
    }
    btn.disabled = false;
  }

  const num = (v) => { const n = parseFloat(String(v ?? '').replace(/[^0-9.\-]/g, '')); return Number.isFinite(n) ? n : null; };
  function readDate(v) {
    const s = String(v ?? '').trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (m) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${pad(m[1])}-${pad(m[2])}`;
    const d = new Date(s);
    return s && !isNaN(d) ? toISO(d) : null;
  }

  async function restoreBackup(ev) {
    const file = ev.target.files[0];
    ev.target.value = '';
    if (!file) return;
    const u = user();
    if (!u) return openProfiles();
    try {
      const XLSX = await ensureXlsx();
      const wb = XLSX.read(await file.arrayBuffer());
      const s1 = wb.Sheets['Cars'], s2 = wb.Sheets['Maintenance log'];
      if (!s1 || !s2) { toast('That file is not an AutoLog backup. It needs sheets named "Cars" and "Maintenance log".'); return; }
      const opt = { raw: false, defval: '', dateNF: 'yyyy-mm-dd' };
      const carRows = XLSX.utils.sheet_to_json(s1, opt);
      const logRows = XLSX.utils.sheet_to_json(s2, opt);

      const oldPhotos = new Map(myCars().map((c) => [c.id, c.vinPhoto]));
      const idMap = new Map(), nameMap = new Map();
      const cars = carRows.map((r) => {
        const car = { id: uid(), userId: u.id, vinPhoto: oldPhotos.get(String(r['Car ID']).trim()) || '' };
        CAR_FIELDS.forEach((f) => { car[f.key] = String(r[f.label] ?? '').trim().slice(0, 80); });
        if (r['Car ID']) idMap.set(String(r['Car ID']).trim(), car);
        nameMap.set(carName(car).toLowerCase(), car);
        return car;
      }).filter((c) => c.make || c.model);

      let skipped = 0;
      const entries = [];
      logRows.forEach((r, i) => {
        const date = readDate(r['Date']);
        const description = String(r['Work done'] ?? '').trim().slice(0, 200);
        if (!date || !description) { skipped++; return; }
        let car = idMap.get(String(r['Car ID']).trim()) || nameMap.get(String(r['Car']).trim().toLowerCase());
        if (!car || !cars.includes(car)) {
          const label = String(r['Car']).trim() || 'Imported car';
          car = { id: uid(), userId: u.id, vinPhoto: '', make: label, model: '' };
          CAR_FIELDS.forEach((f) => { car[f.key] ??= ''; });
          cars.push(car);
          nameMap.set(label.toLowerCase(), car);
        }
        const mileage = num(r['Mileage']);
        entries.push({ id: uid(), carId: car.id, created: i, date, description, location: String(r['Location'] ?? '').trim().slice(0, 120), mileage: mileage === null ? null : Math.max(0, Math.round(mileage)), cost: Math.max(0, Math.round((num(r['Cost']) || 0) * 100) / 100), oilChange: /^(y|yes|true|1)$/i.test(String(r['Oil change']).trim()) });
      });

      if (!cars.length) { toast('No cars were found in that backup, so nothing was changed.'); return; }
      const have = myCars().length;
      if (have && !confirm(`Restore ${cars.length} car${cars.length === 1 ? '' : 's'} and ${entries.length} log entr${entries.length === 1 ? 'y' : 'ies'} from this backup? This replaces the ${have} car${have === 1 ? '' : 's'} currently in ${u.name}'s profile.`)) return;

      const before = JSON.stringify(db);
      const mine = new Set(myCars().map((c) => c.id));
      db.entries = db.entries.filter((e) => !mine.has(e.carId)).concat(entries);
      db.cars = db.cars.filter((c) => c.userId !== u.id).concat(cars);
      Object.assign(prefs(), { activeCarId: cars[0].id, notified: {} });
      if (!save()) { db = JSON.parse(before); return; }
      resetEntryForm();
      clearFilters(false);
      renderAll();
      toast(`Restored ${cars.length} car${cars.length === 1 ? '' : 's'} and ${entries.length} log entr${entries.length === 1 ? 'y' : 'ies'}.${skipped ? ` Skipped ${skipped} row${skipped === 1 ? '' : 's'} with no date or description.` : ''}`);
    } catch {
      toast('That file could not be read. Choose an AutoLog backup in .xlsx format.');
    }
  }

  /* ---------- Events ---------- */
  function clearFilters(render = true) {
    ['fSearch', 'fLoc', 'fFrom', 'fTo'].forEach((id) => { $(id).value = ''; });
    $('fOil').checked = false;
    if (render) renderLog();
  }

  function init() {
    buildCarForm();

    $('carSelect').addEventListener('change', (e) => {
      prefs().activeCarId = e.target.value;
      save();
      resetEntryForm();
      clearFilters(false);
      renderAll();
    });
    $('addCarBtn').addEventListener('click', () => openCarDialog(null));
    $('editCarBtn').addEventListener('click', () => openCarDialog(activeCar()));
    $('deleteCarBtn').addEventListener('click', deleteCar);
    $('carForm').addEventListener('submit', submitCar);

    $('profileBtn').addEventListener('click', openProfiles);
    $('profileForm').addEventListener('submit', submitProfile);
    $('profileList').addEventListener('click', (e) => {
      const sw = e.target.closest('[data-switch]'), del = e.target.closest('[data-delprofile]');
      if (sw) { switchUser(sw.dataset.switch); $('profileDialog').close(); toast(`Switched to ${user().name}.`); }
      if (del) { deleteProfile(del.dataset.delprofile); if (!db.users.length) openProfiles(); }
    });
    // A profile is required before anything can be saved
    $('profileDialog').addEventListener('cancel', (e) => { if (!db.users.length) e.preventDefault(); });

    // Closing the car form goes through the unsaved-changes check
    const closeDialog = (d) => (d.id === 'carDialog' ? requestCloseCarDialog() : d.close());
    document.querySelectorAll('dialog [data-close]').forEach((b) => b.addEventListener('click', () => closeDialog(b.closest('dialog'))));
    document.querySelectorAll('dialog').forEach((d) => d.addEventListener('click', (e) => {
      if (e.target === d && !(d.id === 'profileDialog' && !db.users.length)) closeDialog(d);
    }));
    $('carDialog').addEventListener('cancel', (e) => { e.preventDefault(); requestCloseCarDialog(); }); // Escape key
    $('unsavedYes').addEventListener('click', () => { $('unsavedDialog').close(); $('carDialog').close(); });
    $('unsavedNo').addEventListener('click', () => $('unsavedDialog').close());

    $('entryForm').addEventListener('submit', submitEntry);
    $('cancelEditBtn').addEventListener('click', resetEntryForm);
    $('eOil').addEventListener('change', () => { view.oilTouched = true; });
    $('eDesc').addEventListener('input', () => { if (!view.oilTouched) $('eOil').checked = looksLikeOilChange($('eDesc').value); });

    ['fSearch', 'fLoc', 'fFrom', 'fTo', 'fOil'].forEach((id) => $(id).addEventListener('input', renderLog));
    $('fSort').addEventListener('change', (e) => { [view.sortKey, view.sortDir] = e.target.value.split(':'); renderLog(); });
    $('clearFiltersBtn').addEventListener('click', () => clearFilters());
    $('logTable').addEventListener('click', (e) => {
      const s = e.target.closest('[data-sort]'), ed = e.target.closest('[data-edit]'), del = e.target.closest('[data-del]');
      if (s) {
        const key = s.dataset.sort;
        if (view.sortKey === key) view.sortDir = view.sortDir === 'asc' ? 'desc' : 'asc';
        else { view.sortKey = key; view.sortDir = key === 'description' || key === 'location' ? 'asc' : 'desc'; }
        renderLog();
      }
      if (ed) startEdit(ed.dataset.edit);
      if (del) deleteEntry(del.dataset.del);
    });

    $('notifyBtn').addEventListener('click', async () => {
      const result = await Notification.requestPermission();
      renderReminder();
      if (result === 'granted') { toast('Notifications are on. AutoLog will alert you when an oil change is due.'); checkReminders(); }
    });
    $('calendarBtn').addEventListener('click', downloadCalendar);

    $('zipInput').addEventListener('input', (e) => { if (!user()) return; prefs().zip = e.target.value.trim(); save(); renderStores(); });

    $('backupBtn').addEventListener('click', downloadBackup);
    $('restoreBtn').addEventListener('click', () => (user() ? $('restoreFile').click() : openProfiles()));
    $('restoreFile').addEventListener('change', restoreBackup);

    resetEntryForm();
    renderAll();
    if (!user()) openProfiles();
    checkReminders();
    document.addEventListener('visibilitychange', () => { if (!document.hidden) { renderReminder(); checkReminders(); } });
  }

  init();
})();
