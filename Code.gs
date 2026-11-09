const SHEETS = {
  Users: ['id', 'created_at', 'email', 'password_hash', 'name', 'role'],
  Pelanggan: ['id', 'created_at', 'nama', 'kontak'],
  SPK: ['id', 'created_at', 'no_spk', 'tanggal_spk', 'pelanggan_id', 'nama_project', 'status', 'nama_pelanggan', 'lokasi_proyek'],
  ItemPanel: ['id', 'created_at', 'spk_id', 'jenis_panel', 'spesifikasi', 'jumlah_unit']
};
const SPK_STATUS = ['Baru', 'Proses', 'Selesai', 'Dikirim'];
const DEFAULT_ADMIN = {
  email: 'admin@panelworks.local',
  password: 'Panel123!',
  name: 'Administrator Panelworks'
};

function doGet(e) {
  setupSpreadsheet();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Panelworks — SPK Pembuatan Panel')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Object.keys(SHEETS).forEach(function(sheetName) {
      let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
      const headers = SHEETS[sheetName];
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(headers);
      } else {
        const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
        headers.forEach(function(header, index) {
          if (current[index] !== header) sheet.getRange(1, index + 1).setValue(header);
        });
      }
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#0d6257')
        .setFontColor('#ffffff')
        .setVerticalAlignment('middle');
      sheet.setFrozenRows(1);
      sheet.setRowHeight(1, 34);
      sheet.autoResizeColumns(1, headers.length);
    });
    ensureDefaultAdmin_();
    return { success: true, message: 'Database siap digunakan.' };
  } catch (error) {
    return { success: false, message: error.message || 'Gagal menyiapkan database.' };
  }
}

function loginAdmin(email, password) {
  try {
    ensureDatabase_();
    const cleanEmail = String(email || '').trim().toLowerCase();
    const cleanPassword = String(password || '');
    if (!cleanEmail || !cleanPassword) return fail_('Email dan password wajib diisi.');
    const rows = getSheet_('Users').getDataRange().getValues();
    const headers = rows.shift() || [];
    const ix = indexMap_(headers);
    const user = rows.find(function(row) {
      return String(row[ix.email] || '').trim().toLowerCase() === cleanEmail;
    });
    if (!user || String(user[ix.password_hash] || '') !== hashPassword_(cleanPassword)) {
      return fail_('Email atau password tidak sesuai.');
    }
    const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
    CacheService.getScriptCache().put('panelworks_session_' + token, JSON.stringify({
      email: cleanEmail,
      name: String(user[ix.name] || 'Administrator'),
      role: String(user[ix.role] || 'Admin')
    }), 21600);
    return {
      success: true,
      message: 'Login berhasil.',
      data: { email: cleanEmail, name: String(user[ix.name] || 'Administrator'), role: String(user[ix.role] || 'Admin'), token: token }
    };
  } catch (error) {
    return fail_(error.message || 'Login gagal.');
  }
}

function getDashboard(token) {
  try {
    requireAuth_(token);
    ensureDatabase_();
    const spkRows = getObjects_('SPK');
    const items = getObjects_('ItemPanel');
    const statusCounts = { Baru: 0, Proses: 0, Selesai: 0, Dikirim: 0 };
    spkRows.forEach(function(row) {
      const status = String(row.status || 'Baru');
      if (Object.prototype.hasOwnProperty.call(statusCounts, status)) statusCounts[status]++;
    });
    const totalPanel = items.reduce(function(sum, row) {
      return sum + (Number(row.jumlah_unit) || 0);
    }, 0);
    return ok_({
      totalSPK: spkRows.length,
      totalPanel: totalPanel,
      statusCounts: statusCounts
    });
  } catch (error) {
    return fail_(error.message || 'Gagal memuat dashboard.');
  }
}

function getPelanggan(token) {
  try {
    requireAuth_(token);
    ensureDatabase_();
    const result = getObjects_('Pelanggan').map(function(row) {
      return { id: String(row.id), nama: String(row.nama || ''), kontak: String(row.kontak || '') };
    });
    result.sort(function(a, b) { return a.nama.localeCompare(b.nama, 'id'); });
    return ok_(result);
  } catch (error) {
    return fail_(error.message || 'Gagal memuat pelanggan.');
  }
}

function createPelanggan(nama, kontak, token) {
  try {
    requireAuth_(token);
    ensureDatabase_();
    const cleanName = String(nama || '').trim();
    if (!cleanName) return fail_('Nama pelanggan wajib diisi.');
    const existing = getObjects_('Pelanggan').some(function(row) {
      return String(row.nama || '').trim().toLowerCase() === cleanName.toLowerCase();
    });
    if (existing) return fail_('Nama pelanggan tersebut sudah terdaftar.');
    const id = Utilities.getUuid();
    getSheet_('Pelanggan').appendRow([id, new Date(), cleanName, String(kontak || '').trim()]);
    return { success: true, message: 'Pelanggan berhasil ditambahkan.', data: { id: id } };
  } catch (error) {
    return fail_(error.message || 'Gagal menambahkan pelanggan.');
  }
}

function updatePelanggan(id, nama, kontak, token) {
  try {
    requireAuth_(token);
    ensureDatabase_();
    const cleanName = String(nama || '').trim();
    if (!String(id || '').trim() || !cleanName) return fail_('ID dan nama pelanggan wajib diisi.');
    const rows = getObjects_('Pelanggan');
    const duplicate = rows.some(function(row) {
      return String(row.id) !== String(id) && String(row.nama || '').trim().toLowerCase() === cleanName.toLowerCase();
    });
    if (duplicate) return fail_('Nama pelanggan tersebut sudah digunakan.');
    const found = findRow_('Pelanggan', id);
    if (!found) return fail_('Pelanggan tidak ditemukan.');
    const sheet = getSheet_('Pelanggan');
    sheet.getRange(found.row, 3, 1, 2).setValues([[cleanName, String(kontak || '').trim()]]);
    return { success: true, message: 'Data pelanggan berhasil diperbarui.' };
  } catch (error) {
    return fail_(error.message || 'Gagal memperbarui pelanggan.');
  }
}

function deletePelanggan(id, token) {
  try {
    requireAuth_(token);
    ensureDatabase_();
    if (getObjects_('SPK').some(function(row) { return String(row.pelanggan_id) === String(id); })) {
      return fail_('Pelanggan sudah digunakan pada SPK dan tidak dapat dihapus.');
    }
    const found = findRow_('Pelanggan', id);
    if (!found) return fail_('Pelanggan tidak ditemukan.');
    getSheet_('Pelanggan').deleteRow(found.row);
    return { success: true, message: 'Pelanggan berhasil dihapus.' };
  } catch (error) {
    return fail_(error.message || 'Gagal menghapus pelanggan.');
  }
}

function createSPK(data, token) {
  const lock = LockService.getScriptLock();
  try {
    requireAuth_(token);
    ensureDatabase_();
    lock.waitLock(20000);
    const payload = validateSPK_(data, false);
    const id = Utilities.getUuid();
    const noSpk = nextSPKNumber_(payload.tanggal_spk);
    const created = new Date();
    const spkSheet = getSheet_('SPK');
    spkSheet.appendRow([id, created, noSpk, payload.tanggal_spk, '', payload.nama_project, 'Baru', payload.nama_pelanggan, payload.lokasi_proyek]);
    const itemSheet = getSheet_('ItemPanel');
    payload.items.forEach(function(item) {
      itemSheet.appendRow([Utilities.getUuid(), created, id, item.jenis_panel, item.spesifikasi, item.jumlah_unit]);
    });
    return { success: true, message: 'SPK berhasil dibuat.', data: { spkId: id, no_spk: noSpk } };
  } catch (error) {
    return fail_(error.message || 'Gagal membuat SPK.');
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function getSPKList(filter, token) {
  try {
    requireAuth_(token);
    ensureDatabase_();
    const f = filter || {};
    const customers = {};
    getObjects_('Pelanggan').forEach(function(c) { customers[String(c.id)] = String(c.nama || ''); });
    const itemsBySpk = {};
    getObjects_('ItemPanel').forEach(function(item) {
      const key = String(item.spk_id);
      if (!itemsBySpk[key]) itemsBySpk[key] = { count: 0, units: 0 };
      itemsBySpk[key].count++;
      itemsBySpk[key].units += Number(item.jumlah_unit) || 0;
    });
    let result = getObjects_('SPK').map(function(row) {
      const stats = itemsBySpk[String(row.id)] || { count: 0, units: 0 };
      return {
        id: String(row.id),
        no_spk: String(row.no_spk || ''),
        tanggal_spk: dateString_(row.tanggal_spk),
        pelanggan_id: String(row.pelanggan_id || ''),
        nama_pelanggan: String(row.nama_pelanggan || customers[String(row.pelanggan_id)] || ''),
        nama_project: String(row.nama_project || ''),
        lokasi_proyek: String(row.lokasi_proyek || ''),
        status: String(row.status || 'Baru'),
        total_item: stats.count,
        total_unit: stats.units
      };
    });
    if (f.status) result = result.filter(function(row) { return row.status === String(f.status); });
    if (f.nama_pelanggan) result = result.filter(function(row) { return row.nama_pelanggan.toLowerCase().indexOf(String(f.nama_pelanggan).trim().toLowerCase()) !== -1; });
    if (f.tanggal_awal) result = result.filter(function(row) { return row.tanggal_spk >= String(f.tanggal_awal); });
    if (f.tanggal_akhir) result = result.filter(function(row) { return row.tanggal_spk <= String(f.tanggal_akhir); });
    result.sort(function(a, b) {
      return b.tanggal_spk.localeCompare(a.tanggal_spk) || b.no_spk.localeCompare(a.no_spk);
    });
    return ok_(result);
  } catch (error) {
    return fail_(error.message || 'Gagal memuat daftar SPK.');
  }
}

function getSPKDetail(spkId, token) {
  try {
    requireAuth_(token);
    ensureDatabase_();
    const found = findRow_('SPK', spkId);
    if (!found) return fail_('SPK tidak ditemukan.');
    const h = found.data;
    const customer = findRow_('Pelanggan', h.pelanggan_id);
    const items = getObjects_('ItemPanel').filter(function(item) {
      return String(item.spk_id) === String(spkId);
    }).map(function(item) {
      return {
        id: String(item.id),
        jenis_panel: String(item.jenis_panel || ''),
        spesifikasi: String(item.spesifikasi || ''),
        jumlah_unit: Number(item.jumlah_unit) || 0
      };
    });
    return ok_({
      header: {
        id: String(h.id),
        no_spk: String(h.no_spk || ''),
        tanggal_spk: dateString_(h.tanggal_spk),
        pelanggan_id: String(h.pelanggan_id || ''),
        nama_pelanggan: String(h.nama_pelanggan || (customer ? customer.data.nama : '') || ''),
        nama_project: String(h.nama_project || ''),
        lokasi_proyek: String(h.lokasi_proyek || ''),
        status: String(h.status || 'Baru')
      },
      items: items
    });
  } catch (error) {
    return fail_(error.message || 'Gagal memuat detail SPK.');
  }
}

function updateSPK(spkId, data, token) {
  const lock = LockService.getScriptLock();
  try {
    requireAuth_(token);
    ensureDatabase_();
    lock.waitLock(20000);
    const found = findRow_('SPK', spkId);
    if (!found) return fail_('SPK tidak ditemukan.');
    const payload = validateSPK_(data, true);
    const sheet = getSheet_('SPK');
    sheet.getRange(found.row, 4, 1, 3).setValues([[payload.tanggal_spk, '', payload.nama_project]]);
    sheet.getRange(found.row, 8).setValue(payload.nama_pelanggan);
    sheet.getRange(found.row, 9).setValue(payload.lokasi_proyek);
    replaceItems_(spkId, payload.items);
    return { success: true, message: 'SPK berhasil diperbarui.' };
  } catch (error) {
    return fail_(error.message || 'Gagal memperbarui SPK.');
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function updateStatusSPK(spkId, statusBaru, token) {
  try {
    requireAuth_(token);
    ensureDatabase_();
    const status = String(statusBaru || '').trim();
    if (SPK_STATUS.indexOf(status) === -1) return fail_('Status SPK tidak valid.');
    const found = findRow_('SPK', spkId);
    if (!found) return fail_('SPK tidak ditemukan.');
    getSheet_('SPK').getRange(found.row, 7).setValue(status);
    return { success: true, message: 'Status SPK diperbarui menjadi ' + status + '.' };
  } catch (error) {
    return fail_(error.message || 'Gagal memperbarui status SPK.');
  }
}

function deleteSPK(spkId, token) {
  const lock = LockService.getScriptLock();
  try {
    requireAuth_(token);
    ensureDatabase_();
    lock.waitLock(20000);
    const found = findRow_('SPK', spkId);
    if (!found) return fail_('SPK tidak ditemukan.');
    deleteRowsByValue_('ItemPanel', 'spk_id', spkId);
    getSheet_('SPK').deleteRow(found.row);
    return { success: true, message: 'SPK dan item panel berhasil dihapus.' };
  } catch (error) {
    return fail_(error.message || 'Gagal menghapus SPK.');
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function generateNoSPK(token) {
  try {
    requireAuth_(token);
    ensureDatabase_();
    return ok_({ no_spk: nextSPKNumber_(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')) });
  } catch (error) {
    return fail_(error.message || 'Gagal membuat nomor SPK.');
  }
}

function ensureDatabase_() {
  const result = setupSpreadsheet();
  if (!result.success) throw new Error(result.message);
}

function ensureDefaultAdmin_() {
  const sheet = getSheet_('Users');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift() || [];
  const ix = indexMap_(headers);
  const exists = data.some(function(row) {
    return String(row[ix.email] || '').trim().toLowerCase() === DEFAULT_ADMIN.email;
  });
  if (!exists) {
    sheet.appendRow([
      Utilities.getUuid(),
      new Date(),
      DEFAULT_ADMIN.email,
      hashPassword_(DEFAULT_ADMIN.password),
      DEFAULT_ADMIN.name,
      'Admin'
    ]);
  }
}

function requireAuth_(token) {
  const value = String(token || '').trim();
  if (!value) throw new Error('Sesi tidak ditemukan. Silakan login kembali.');
  const session = CacheService.getScriptCache().get('panelworks_session_' + value);
  if (!session) throw new Error('Sesi login telah berakhir. Silakan login kembali.');
  const parsed = JSON.parse(session);
  if (parsed.role !== 'Admin') throw new Error('Akses hanya tersedia untuk Admin.');
  return parsed;
}

function hashPassword_(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password), Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    const value = (b + 256) % 256;
    return ('0' + value.toString(16)).slice(-2);
  }).join('');
}

function validateSPK_(data, isUpdate) {
  const payload = data || {};
  const date = String(payload.tanggal_spk || '').trim();
  const customerName = String(payload.nama_pelanggan || '').trim();
  const project = String(payload.nama_project || '').trim();
  const location = String(payload.lokasi_proyek || '').trim();
  const items = Array.isArray(payload.items) ? payload.items : [];
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal SPK wajib diisi dengan format yang benar.');
  if (!customerName) throw new Error('Nama pelanggan wajib diisi.');
  if (!project) throw new Error('Nama project wajib diisi.');
  if (!items.length) throw new Error('SPK harus memiliki minimal satu item panel.');
  const normalized = items.map(function(item) {
    const kind = String(item.jenis_panel || '').trim();
    const spec = String(item.spesifikasi || '').trim();
    const quantity = Number(item.jumlah_unit);
    if (!kind) throw new Error('Jenis panel pada setiap item wajib diisi.');
    if (!spec) throw new Error('Spesifikasi / keterangan pada setiap item wajib diisi.');
    if (!Number.isInteger(quantity) || quantity < 1) throw new Error('Jumlah unit harus berupa bilangan bulat minimal 1.');
    return { id: isUpdate ? String(item.id || '') : '', jenis_panel: kind, spesifikasi: spec, jumlah_unit: quantity };
  });
  return { tanggal_spk: date, nama_pelanggan: customerName, nama_project: project, lokasi_proyek: location, items: normalized };
}

function replaceItems_(spkId, items) {
  const sheet = getSheet_('ItemPanel');
  const existing = getObjects_('ItemPanel').filter(function(row) {
    return String(row.spk_id) === String(spkId);
  });
  const existingById = {};
  existing.forEach(function(item) { existingById[String(item.id)] = item; });
  const keepIds = [];
  const created = new Date();
  items.forEach(function(item) {
    const id = item.id && existingById[item.id] ? item.id : Utilities.getUuid();
    keepIds.push(id);
    if (existingById[id]) {
      const found = findRow_('ItemPanel', id);
      sheet.getRange(found.row, 4, 1, 3).setValues([[item.jenis_panel, item.spesifikasi, item.jumlah_unit]]);
    } else {
      sheet.appendRow([id, created, spkId, item.jenis_panel, item.spesifikasi, item.jumlah_unit]);
    }
  });
  const remove = existing.filter(function(item) { return keepIds.indexOf(String(item.id)) === -1; });
  remove.map(function(item) { return findRow_('ItemPanel', item.id); })
    .filter(Boolean)
    .sort(function(a, b) { return b.row - a.row; })
    .forEach(function(found) { sheet.deleteRow(found.row); });
}

function nextSPKNumber_(dateValue) {
  const date = String(dateValue || '').slice(0, 10);
  const compact = date.replace(/-/g, '');
  if (!/^\d{8}$/.test(compact)) throw new Error('Tanggal tidak valid untuk nomor SPK.');
  const prefix = 'SPK/' + compact.slice(0, 6) + '/';
  const numbers = getObjects_('SPK').map(function(row) {
    const no = String(row.no_spk || '');
    if (no.indexOf(prefix) !== 0) return 0;
    const parsed = Number(no.slice(prefix.length));
    return Number.isFinite(parsed) ? parsed : 0;
  });
  const next = Math.max.apply(null, [0].concat(numbers)) + 1;
  return prefix + String(next).padStart(3, '0');
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getObjects_(sheetName) {
  const values = getSheet_(sheetName).getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values.shift();
  return values.filter(function(row) {
    return row.some(function(cell) { return cell !== '' && cell !== null; });
  }).map(function(row) {
    const object = {};
    headers.forEach(function(header, index) {
      object[String(header)] = row[index];
    });
    return object;
  });
}

function findRow_(sheetName, id) {
  const sheet = getSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  const headers = values[0];
  const idIndex = headers.indexOf('id');
  if (idIndex < 0) return null;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIndex]) === String(id)) {
      const data = {};
      headers.forEach(function(header, j) { data[String(header)] = values[i][j]; });
      return { row: i + 1, data: data };
    }
  }
  return null;
}

function deleteRowsByValue_(sheetName, columnName, value) {
  const sheet = getSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  const columnIndex = values[0].indexOf(columnName);
  if (columnIndex < 0) return;
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][columnIndex]) === String(value)) rows.push(i + 1);
  }
  rows.sort(function(a, b) { return b - a; }).forEach(function(row) { sheet.deleteRow(row); });
}

function indexMap_(headers) {
  const map = {};
  headers.forEach(function(header, index) { map[String(header)] = index; });
  return map;
}

function dateString_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const text = String(value || '');
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);
  return text;
}

function ok_(data) {
  return { success: true, data: data };
}

function fail_(message) {
  return { success: false, message: String(message || 'Terjadi kesalahan.') };
}