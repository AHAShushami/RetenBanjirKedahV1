/* ============================================
   RETEN BANJIR KKM — Main Application Logic
   Handles: Tab navigation, form logic, dynamic rows,
   auto-calculations, Google Sheets, Word export
   ============================================ */

// ============================================
// CONFIGURATION
// ============================================
const CONFIG = {
    // Replace with your actual Google Apps Script Web App URL
    SHEETS_URL: '',
    NEGERI: 'KEDAH',
    DAERAH_LIST: [
        'BALING', 'BANDAR BAHARU', 'KOTA SETAR', 'KUALA MUDA',
        'KUBANG PASU', 'KULIM', 'LANGKAWI', 'PADANG TERAP',
        'PENDANG', 'POKOK SENA', 'SIK', 'YAN'
    ]
};

// Cumulative data cache (loaded from Google Sheets)
let cumulativeCache = {};

// ============================================
// INITIALIZATION
// ============================================
document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    initSharedHeader();
    initAutoCalculations();
    initDynamicRows();
    setDefaultDate();

    // Hide loading
    setTimeout(() => {
        document.getElementById('loading-overlay').classList.add('hidden');
    }, 600);
});

function setDefaultDate() {
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('hdr-tarikh').value = today;
}

// ============================================
// TAB NAVIGATION
// ============================================
function initTabs() {
    const tabs = document.querySelectorAll('.tab');
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const moduleId = tab.dataset.module;
            switchModule(moduleId);
        });
    });
}

function switchModule(moduleId) {
    // Update tabs
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelector(`[data-module="${moduleId}"]`).classList.add('active');

    // Update panels
    document.querySelectorAll('.module-panel').forEach(p => p.classList.remove('active'));
    document.getElementById(`panel-${moduleId}`).classList.add('active');
}

// ============================================
// SHARED HEADER
// ============================================
function initSharedHeader() {
    const daerahSelect = document.getElementById('hdr-daerah');
    daerahSelect.addEventListener('change', () => {
        const daerah = daerahSelect.value;
        // Update Borang 1 daerah display
        const b1Display = document.getElementById('b1-daerah-display');
        if (b1Display) b1Display.textContent = daerah || '—';

        // Load cumulative data for selected daerah
        if (daerah) {
            loadCumulativeData(daerah);
        }
    });
}

function getHeaderData() {
    return {
        negeri: document.getElementById('hdr-negeri').value,
        tarikh: document.getElementById('hdr-tarikh').value,
        daerah: document.getElementById('hdr-daerah').value,
        nama: document.getElementById('hdr-nama').value,
        jawatan: document.getElementById('hdr-jawatan').value,
        nama_semak: document.getElementById('hdr-nama-semak').value
    };
}

function validateHeader() {
    const hdr = getHeaderData();
    if (!hdr.tarikh) { showToast('Sila masukkan tarikh laporan', 'error'); return false; }
    if (!hdr.daerah) { showToast('Sila pilih daerah', 'error'); return false; }
    if (!hdr.nama) { showToast('Sila masukkan nama pelapor', 'error'); return false; }
    return true;
}

// ============================================
// AUTO-CALCULATIONS
// ============================================
function initAutoCalculations() {
    // Borang 1: Pasukan totals
    ['b1-kes-pasukan', 'b1-per-pasukan'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', calcB1Totals);
    });
    ['b1-kes-ahli', 'b1-per-ahli'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', calcB1Totals);
    });

    // Borang 3: Pasukan totals
    ['b3-kes-bil', 'b3-per-bil', 'b3-mhpss-bil', 'b3-mobil-bil', 'b3-statik-bil'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', calcB3Totals);
    });
    ['b3-kes-ahli', 'b3-per-ahli', 'b3-mhpss-ahli', 'b3-mobil-ahli', 'b3-statik-ahli'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', calcB3Totals);
    });

    // Borang 7: Disease totals
    const b7Fields = ['b7-ari-bil', 'b7-age-bil', 'b7-conjunctivitis-bil', 'b7-hfmd-bil',
        'b7-typhoid-bil', 'b7-leptospirosis-bil', 'b7-chickenpox-bil', 'b7-skin-bil',
        'b7-other-bil', 'b7-fever-bil'];
    b7Fields.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', calcB7Totals);
    });

    // Borang 8: AI/BI calculation
    ['b8-positif-bil', 'b8-diperiksa-bil', 'b8-bekas-positif-bil', 'b8-bekas-diperiksa-bil'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', calcB8AIBI);
    });

    // Borang 10: Percentages
    ['b10-tikus-bil', 'b10-lipas-bil', 'b10-diperiksa-bil'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', calcB10Pct);
    });

    // Borang 12: Hospital totals
    ['b12-lelaki-bil', 'b12-perempuan-bil', 'b12-kanak-bil'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', calcB12Totals);
    });

    // Borang 15: Risk score total
    for (let i = 1; i <= 8; i++) {
        const el = document.getElementById(`b15-s${i}`);
        if (el) el.addEventListener('input', calcB15Total);
    }

    // Borang 20: AI/BI calculation (pasca)
    ['b20-positif-bil', 'b20-diperiksa-bil', 'b20-bekas-positif-bil', 'b20-bekas-diperiksa-bil'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', calcB20AIBI);
    });
}

function calcB1Totals() {
    const kesPasukan = num('b1-kes-pasukan');
    const perPasukan = num('b1-per-pasukan');
    const kesAhli = num('b1-kes-ahli');
    const perAhli = num('b1-per-ahli');
    setText('b1-jum-pasukan', kesPasukan + perPasukan);
    setText('b1-jum-ahli', kesAhli + perAhli);
}

function calcB3Totals() {
    const bilFields = ['b3-kes-bil', 'b3-per-bil', 'b3-mhpss-bil', 'b3-mobil-bil', 'b3-statik-bil'];
    const ahliFields = ['b3-kes-ahli', 'b3-per-ahli', 'b3-mhpss-ahli', 'b3-mobil-ahli', 'b3-statik-ahli'];
    setText('b3-jum-bil', sumFields(bilFields));
    setText('b3-jum-ahli', sumFields(ahliFields));
}

function calcB7Totals() {
    const fields = ['b7-ari-bil', 'b7-age-bil', 'b7-conjunctivitis-bil', 'b7-hfmd-bil',
        'b7-typhoid-bil', 'b7-leptospirosis-bil', 'b7-chickenpox-bil', 'b7-skin-bil',
        'b7-other-bil', 'b7-fever-bil'];
    setText('b7-total-bil', sumFields(fields));
}

function calcB8AIBI() {
    const premisPositif = num('b8-positif-bil');
    const premisDiperiksa = num('b8-diperiksa-bil');
    const bekasPositif = num('b8-bekas-positif-bil');
    const bekasDiperiksa = num('b8-bekas-diperiksa-bil');

    const ai = premisDiperiksa > 0 ? (premisPositif / premisDiperiksa * 100) : 0;
    const bi = premisDiperiksa > 0 ? (bekasPositif / premisDiperiksa) : 0;

    setText('b8-ai', ai.toFixed(2) + '%');
    setText('b8-bi', bi.toFixed(2));

    // Show/hide alert
    const alertEl = document.getElementById('b8-alert');
    if (ai > 1 || bi > 5) {
        alertEl.style.display = 'block';
    } else {
        alertEl.style.display = 'none';
    }
}

function calcB10Pct() {
    const diperiksa = num('b10-diperiksa-bil');
    const tikus = num('b10-tikus-bil');
    const lipas = num('b10-lipas-bil');

    const pctTikus = diperiksa > 0 ? (tikus / diperiksa * 100) : 0;
    const pctLipas = diperiksa > 0 ? (lipas / diperiksa * 100) : 0;

    setText('b10-pct-tikus', pctTikus.toFixed(2) + '%');
    setText('b10-pct-lipas', pctLipas.toFixed(2) + '%');
}

function calcB12Totals() {
    const fields = ['b12-lelaki-bil', 'b12-perempuan-bil', 'b12-kanak-bil'];
    setText('b12-total-bil', sumFields(fields));
}

function calcB15Total() {
    let total = 0;
    for (let i = 1; i <= 8; i++) {
        total += num(`b15-s${i}`);
    }
    setText('b15-total-skor', total);
}

function calcB20AIBI() {
    const premisPositif = num('b20-positif-bil');
    const premisDiperiksa = num('b20-diperiksa-bil');
    const bekasPositif = num('b20-bekas-positif-bil');
    const bekasDiperiksa = num('b20-bekas-diperiksa-bil');

    const ai = premisDiperiksa > 0 ? (premisPositif / premisDiperiksa * 100) : 0;
    const bi = premisDiperiksa > 0 ? (bekasPositif / premisDiperiksa) : 0;

    setText('b20-ai', ai.toFixed(2) + '%');
    setText('b20-bi', bi.toFixed(2));

    const alertEl = document.getElementById('b20-alert');
    if (ai > 1 || bi > 5) {
        alertEl.style.display = 'block';
    } else {
        alertEl.style.display = 'none';
    }
}

// ============================================
// DYNAMIC ROW MANAGEMENT
// ============================================
function initDynamicRows() {
    // Add initial rows
    addPPSRow();
    addFasilitiRow();
    addAnggotaRow();
    addMobilisasiRow();
    addKerosakanRow();
}

let ppsRowCount = 0;
function addPPSRow() {
    ppsRowCount++;
    const tbody = document.getElementById('b2-pps-tbody');
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td>${ppsRowCount}</td>
        <td><input type="text" class="cell-input" style="min-width:140px;text-align:left" placeholder="Nama PPS"></td>
        <td><input type="number" class="cell-input b2-kapasiti" min="0" oninput="calcB2Totals()"></td>
        <td><input type="number" class="cell-input b2-pengendali" min="0" oninput="calcB2Totals()"></td>
        <td><input type="number" class="cell-input b2-pemeriksaan" min="0" oninput="calcB2Totals()"></td>
        <td><button class="btn-danger-sm" onclick="removeRow(this)">✕</button></td>
    `;
    tbody.appendChild(tr);
}

function calcB2Totals() {
    setText('b2-total-kapasiti', sumClass('b2-kapasiti'));
    setText('b2-total-pengendali', sumClass('b2-pengendali'));
    setText('b2-total-pemeriksaan', sumClass('b2-pemeriksaan'));
}

let fasilitiRowCount = 0;
function addFasilitiRow() {
    fasilitiRowCount++;
    const tbody = document.getElementById('b4-fasiliti-tbody');
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td>${fasilitiRowCount}</td>
        <td><input type="text" class="cell-input" style="min-width:140px;text-align:left" placeholder="Nama fasiliti"></td>
        <td><select class="cell-select"><option value="">—</option><option value="1">1</option><option value="2">2</option><option value="3">3</option><option value="4">4</option><option value="5">5</option><option value="6">6</option><option value="7">7</option></select></td>
        <td><input type="text" class="cell-input" style="min-width:100px;text-align:left" placeholder="Lokasi"></td>
        <td><input type="date" class="cell-input"></td>
        <td><input type="date" class="cell-input"></td>
        <td><button class="btn-danger-sm" onclick="removeRow(this)">✕</button></td>
    `;
    tbody.appendChild(tr);
}

function addAduanRow() {
    const tbody = document.getElementById('b14-aduan-tbody');
    const count = tbody.children.length + 1;
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td>${count}</td>
        <td><input type="text" class="cell-input" style="min-width:100px" placeholder="No. Aduan"></td>
        <td><select class="cell-select"><option value="">—</option><option value="Orang Awam">Orang Awam</option><option value="Lain-lain">Lain-lain</option></select></td>
        <td><select class="cell-select"><option value="">—</option><option value="Pencemaran Makanan">Pencemaran Makanan</option><option value="Premis">Premis</option><option value="Pengendali">Pengendali</option><option value="Lain-lain">Lain-lain</option></select></td>
        <td><select class="cell-select"><option value="">—</option><option value="Pemeriksaan">Pemeriksaan</option><option value="Amaran">Amaran</option><option value="Pensampelan">Pensampelan</option><option value="Penutupan">Penutupan</option><option value="Lain-lain">Lain-lain</option></select></td>
        <td><input type="text" class="cell-input" style="min-width:100px;text-align:left" placeholder="Catatan"></td>
        <td><button class="btn-danger-sm" onclick="removeRow(this)">✕</button></td>
    `;
    tbody.appendChild(tr);
}

let anggotaRowCount = 0;
function addAnggotaRow() {
    anggotaRowCount++;
    const tbody = document.getElementById('b16-anggota-tbody');
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td>${anggotaRowCount}</td>
        <td><input type="text" class="cell-input" style="min-width:140px;text-align:left" placeholder="Nama"></td>
        <td><input type="text" class="cell-input" style="min-width:100px;text-align:left" placeholder="Jawatan"></td>
        <td><select class="cell-select"><option value="">—</option><option value="Ibu Pejabat">Ibu Pejabat</option><option value="JKN">JKN</option><option value="PKD">PKD</option><option value="Pergigian">Pergigian</option><option value="Hospital">Hospital</option><option value="Institusi">Institusi</option></select></td>
        <td><input type="text" class="cell-input" style="min-width:120px;text-align:left" placeholder="Tempat bertugas"></td>
        <td><select class="cell-select"><option value="">—</option><option value="1">1 - Rumah banjir</option><option value="2">2 - Jalan terputus</option><option value="3">3 - Lain-lain</option></select></td>
        <td><button class="btn-danger-sm" onclick="removeRow(this)">✕</button></td>
    `;
    tbody.appendChild(tr);
}

function addMobilisasiRow() {
    const tbody = document.getElementById('b19-mobilisasi-tbody');
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td><input type="text" class="cell-input" style="min-width:120px;text-align:left" placeholder="Lokasi"></td>
        <td><input type="number" class="cell-input b19-farmasi" min="0" oninput="calcB19Totals()"></td>
        <td><input type="number" class="cell-input b19-perubatan" min="0" oninput="calcB19Totals()"></td>
        <td><input type="number" class="cell-input b19-ppp" min="0" oninput="calcB19Totals()"></td>
        <td><input type="number" class="cell-input b19-jururawat" min="0" oninput="calcB19Totals()"></td>
        <td><input type="number" class="cell-input b19-ppkp" min="0" oninput="calcB19Totals()"></td>
        <td><input type="number" class="cell-input b19-pemandu" min="0" oninput="calcB19Totals()"></td>
        <td class="cell-auto b19-jumlah-row">0</td>
        <td><button class="btn-danger-sm" onclick="removeRow(this); calcB19Totals()">✕</button></td>
    `;
    tbody.appendChild(tr);
}

function calcB19Totals() {
    // Per-row totals
    document.querySelectorAll('#b19-mobilisasi-tbody tr').forEach(tr => {
        const cats = ['b19-farmasi', 'b19-perubatan', 'b19-ppp', 'b19-jururawat', 'b19-ppkp', 'b19-pemandu'];
        let rowTotal = 0;
        cats.forEach(cls => {
            const inp = tr.querySelector(`.${cls}`);
            if (inp) rowTotal += (parseInt(inp.value) || 0);
        });
        const totalCell = tr.querySelector('.b19-jumlah-row');
        if (totalCell) totalCell.textContent = rowTotal;
    });

    // Column totals
    const cats = ['farmasi', 'perubatan', 'ppp', 'jururawat', 'ppkp', 'pemandu'];
    let grandTotal = 0;
    cats.forEach(cat => {
        const total = sumClass(`b19-${cat}`);
        setText(`b19-total-${cat}`, total);
        grandTotal += total;
    });
    setText('b19-total-jumlah', grandTotal);
}

function addKerosakanRow() {
    const tbody = document.getElementById('b21-kerosakan-tbody');
    const count = tbody.children.length + 1;
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td>${count}</td>
        <td><input type="text" class="cell-input" style="min-width:140px;text-align:left" placeholder="Nama fasiliti"></td>
        <td><input type="text" class="cell-input" style="min-width:120px;text-align:left" placeholder="Jenis kerosakan"></td>
        <td><input type="number" class="cell-input b21-rm" min="0" step="0.01" oninput="calcB21Total()"></td>
        <td><button class="btn-danger-sm" onclick="removeRow(this); calcB21Total()">✕</button></td>
    `;
    tbody.appendChild(tr);
}

function calcB21Total() {
    const total = sumClass('b21-rm');
    setText('b21-total-rm', 'RM ' + total.toLocaleString('en-MY', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
}

function addDynamicRow(prefix) {
    const container = document.getElementById(`${prefix}-rows`);
    const index = container.children.length;

    if (prefix === 'b1-bilik-gerakan') {
        const div = document.createElement('div');
        div.className = 'dynamic-row';
        div.dataset.index = index;
        div.innerHTML = `
            <div class="form-grid cols-2" style="align-items:end">
                <div class="form-group">
                    <label>Bilik Gerakan Banjir</label>
                    <input type="text" class="form-control b1-bilik-nama" placeholder="Nama bilik gerakan">
                </div>
                <div class="form-group" style="display:flex;gap:8px;align-items:end">
                    <div style="flex:1">
                        <label>No Telefon / Faks / H.P</label>
                        <input type="text" class="form-control b1-bilik-tel" placeholder="04-XXXXXXX">
                    </div>
                    <button class="btn-danger-sm" onclick="this.closest('.dynamic-row').remove()" style="margin-bottom:4px">✕</button>
                </div>
            </div>
        `;
        container.appendChild(div);
    } else if (prefix === 'b1-fasiliti-risiko') {
        const div = document.createElement('div');
        div.className = 'dynamic-row';
        div.dataset.index = index;
        div.innerHTML = `
            <div class="form-group" style="display:flex;gap:8px;align-items:end">
                <div style="flex:1">
                    <label>Nama Fasiliti</label>
                    <input type="text" class="form-control b1-fasiliti-nama" placeholder="Nama fasiliti kesihatan">
                </div>
                <button class="btn-danger-sm" onclick="this.closest('.dynamic-row').remove()" style="margin-bottom:4px">✕</button>
            </div>
        `;
        container.appendChild(div);
    }
}

function removeRow(btn) {
    btn.closest('tr').remove();
}

// ============================================
// GOOGLE SHEETS INTEGRATION
// ============================================
async function loadCumulativeData(daerah) {
    if (!CONFIG.SHEETS_URL) {
        console.log('Google Sheets URL not configured — cumulative data will show 0');
        return;
    }

    try {
        setConnectionStatus('loading', 'Memuat data...');
        const response = await fetch(`${CONFIG.SHEETS_URL}?action=getCumulative&daerah=${encodeURIComponent(daerah)}&negeri=${CONFIG.NEGERI}`);
        const data = await response.json();

        if (data.status === 'success') {
            cumulativeCache = data.cumulative || {};
            applyCumulativeData();
            setConnectionStatus('success', 'Data dimuat');
            showToast('Data kumulatif berjaya dimuat', 'success');
        } else {
            throw new Error(data.message || 'Failed to load');
        }
    } catch (err) {
        console.error('Failed to load cumulative data:', err);
        setConnectionStatus('error', 'Gagal memuat');
    }
}

function applyCumulativeData() {
    // Apply cumulative values to all kum-field cells
    Object.entries(cumulativeCache).forEach(([fieldId, value]) => {
        const el = document.getElementById(fieldId);
        if (el) {
            el.textContent = value;
        }
    });
}

async function submitModule(moduleId) {
    if (!validateHeader()) return;

    const hdr = getHeaderData();
    const moduleData = collectModuleData(moduleId);

    const payload = {
        action: 'submit',
        module: moduleId,
        header: hdr,
        data: moduleData,
        timestamp: new Date().toISOString()
    };

    if (!CONFIG.SHEETS_URL) {
        showToast('Google Sheets URL belum dikonfigurasi. Data disimpan secara lokal.', 'info');
        console.log('Module data to submit:', payload);
        // Save to localStorage as fallback
        saveToLocal(moduleId, payload);
        return;
    }

    try {
        setConnectionStatus('loading', 'Menghantar...');
        const response = await fetch(CONFIG.SHEETS_URL, {
            method: 'POST',
            body: JSON.stringify(payload),
            headers: { 'Content-Type': 'text/plain' }
        });
        const result = await response.json();

        if (result.status === 'success') {
            showToast(`Data ${moduleId} berjaya dihantar!`, 'success');
            setConnectionStatus('success', 'Sedia');

            // Reload cumulative after submission
            loadCumulativeData(hdr.daerah);
        } else {
            throw new Error(result.message || 'Submission failed');
        }
    } catch (err) {
        console.error('Submit failed:', err);
        showToast('Gagal menghantar data. Sila cuba lagi.', 'error');
        setConnectionStatus('error', 'Gagal');
        saveToLocal(moduleId, payload);
    }
}

function saveToLocal(moduleId, payload) {
    const key = `reten-banjir-${moduleId}-${payload.header.daerah}-${payload.header.tarikh}`;
    localStorage.setItem(key, JSON.stringify(payload));
}

function collectModuleData(moduleId) {
    const data = {};
    const panel = document.getElementById(`panel-${moduleId}`);
    if (!panel) return data;

    // Collect all input values
    panel.querySelectorAll('input, select, textarea').forEach(el => {
        if (el.id) {
            data[el.id] = el.value;
        }
    });

    // Collect dynamic table rows
    const tables = {
        'pps': collectTableData('b2-pps-tbody'),
        'vektor': null,
        'anggota': {
            anggota: collectTableData('b16-anggota-tbody'),
            mobilisasi: collectTableData('b19-mobilisasi-tbody')
        },
        'pasca': {
            kerosakan: collectTableData('b21-kerosakan-tbody')
        },
        'makanan': {
            aduan: collectTableData('b14-aduan-tbody')
        }
    };

    if (tables[moduleId]) {
        data._tables = tables[moduleId];
    }

    return data;
}

function collectTableData(tbodyId) {
    const tbody = document.getElementById(tbodyId);
    if (!tbody) return [];
    const rows = [];
    tbody.querySelectorAll('tr').forEach(tr => {
        const row = [];
        tr.querySelectorAll('input, select').forEach(el => {
            row.push(el.value);
        });
        rows.push(row);
    });
    return rows;
}

// ============================================
// STATE VIEW
// ============================================
async function refreshStateView() {
    const container = document.getElementById('state-view-content');

    if (!CONFIG.SHEETS_URL) {
        container.innerHTML = generateLocalStateView();
        showToast('Paparan negeri menggunakan data lokal', 'info');
        return;
    }

    try {
        container.innerHTML = '<p class="text-center text-muted py-4">Memuat data...</p>';
        const response = await fetch(`${CONFIG.SHEETS_URL}?action=getStateView&negeri=${CONFIG.NEGERI}`);
        const data = await response.json();

        if (data.status === 'success') {
            container.innerHTML = renderStateView(data.stateData);
            showToast('Data peringkat negeri berjaya dimuat', 'success');
        } else {
            throw new Error(data.message);
        }
    } catch (err) {
        console.error('State view failed:', err);
        container.innerHTML = '<p class="text-center" style="color:var(--accent-red)">Gagal memuat data peringkat negeri. Pastikan Google Sheets URL telah dikonfigurasi.</p>';
    }
}

function generateLocalStateView() {
    return `
        <div class="info-box">
            <strong>Nota:</strong> Google Sheets belum dikonfigurasi. Paparan negeri memerlukan sambungan ke Google Sheets untuk mengumpul data dari semua daerah.
            <br><br>
            Sila konfigurasi URL Google Apps Script di dalam <code>app.js</code> → <code>CONFIG.SHEETS_URL</code>
        </div>
    `;
}

function renderStateView(stateData) {
    // This will be populated from actual Google Sheets data
    let html = '<p class="text-muted">Data dari Google Sheets akan dipaparkan di sini.</p>';
    return html;
}

function exportStateView() {
    showToast('Eksport paparan negeri ke Word...', 'info');
    // Would generate a comprehensive Word document with all district data
}

// ============================================
// WORD EXPORT
// ============================================
async function exportModule(moduleId) {
    if (!validateHeader()) return;

    const hdr = getHeaderData();
    showToast(`Menjana dokumen Word untuk modul ${moduleId}...`, 'info');

    try {
        const blob = await generateWordDocument(moduleId, hdr);
        const filename = `Reten_Banjir_${moduleId}_${hdr.daerah}_${hdr.tarikh}.docx`;
        saveAs(blob, filename);
        showToast(`Dokumen ${filename} berjaya dimuat turun!`, 'success');
    } catch (err) {
        console.error('Export failed:', err);
        showToast('Gagal menjana dokumen Word.', 'error');
    }
}

async function generateWordDocument(moduleId, hdr) {
    const { Document, Packer, Paragraph, Table, TableRow, TableCell,
        WidthType, TextRun, HeadingLevel, AlignmentType, BorderStyle } = docx;

    const tableBorder = {
        top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
        left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
        right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
    };

    const sections = [];

    // Build header paragraphs
    const headerParagraphs = [
        new Paragraph({
            children: [new TextRun({ text: 'KPAS/BANJIR', bold: true, size: 24 })],
            alignment: AlignmentType.CENTER,
        }),
        new Paragraph({
            children: [
                new TextRun({ text: `NEGERI: ${hdr.negeri}`, size: 20 }),
                new TextRun({ text: `     TARIKH: ${formatDate(hdr.tarikh)}`, size: 20 }),
            ],
        }),
        new Paragraph({
            children: [new TextRun({ text: `DAERAH: ${hdr.daerah}`, size: 20 })],
        }),
        new Paragraph({ text: '' }),
    ];

    // Build footer paragraphs
    const footerParagraphs = [
        new Paragraph({ text: '' }),
        new Paragraph({
            children: [
                new TextRun({ text: 'Laporan disediakan oleh:', size: 18 }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({ text: `Nama: ${hdr.nama}`, size: 18 }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({ text: `Jawatan: ${hdr.jawatan}`, size: 18 }),
            ],
        }),
        new Paragraph({ text: '' }),
        new Paragraph({
            children: [
                new TextRun({ text: 'Laporan disemak dan disahkan oleh:', size: 18 }),
            ],
        }),
        new Paragraph({
            children: [
                new TextRun({ text: `Nama: ${hdr.nama_semak}`, size: 18 }),
            ],
        }),
    ];

    // Module-specific content
    const moduleContent = buildModuleContent(moduleId, hdr, { Document, Paragraph, Table, TableRow, TableCell, WidthType, TextRun, HeadingLevel, AlignmentType, BorderStyle, tableBorder });

    const doc = new Document({
        sections: [{
            properties: {
                page: {
                    margin: { top: 720, right: 720, bottom: 720, left: 720 },
                },
            },
            children: [
                ...headerParagraphs,
                ...moduleContent,
                ...footerParagraphs,
            ],
        }],
    });

    return await Packer.toBlob(doc);
}

function buildModuleContent(moduleId, hdr, docxLib) {
    const { Paragraph, Table, TableRow, TableCell, WidthType, TextRun, HeadingLevel, AlignmentType, BorderStyle, tableBorder } = docxLib;

    const content = [];

    // Helper to create a simple table
    function makeTable(headers, rows) {
        const headerRow = new TableRow({
            children: headers.map(h => new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 16 })], alignment: AlignmentType.CENTER })],
                borders: tableBorder,
                width: { size: 100 / headers.length, type: WidthType.PERCENTAGE },
            })),
        });

        const dataRows = rows.map(row => new TableRow({
            children: row.map(cell => new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: String(cell), size: 16 })], alignment: AlignmentType.CENTER })],
                borders: tableBorder,
            })),
        }));

        return new Table({
            rows: [headerRow, ...dataRows],
            width: { size: 100, type: WidthType.PERCENTAGE },
        });
    }

    // Helper to read form values for a module section
    function getVal(id) {
        const el = document.getElementById(id);
        return el ? (el.value || el.textContent || '0') : '0';
    }

    // Build content based on module
    switch (moduleId) {
        case 'pra-bencana':
            content.push(new Paragraph({
                children: [new TextRun({ text: 'BORANG 1 — MAKLUMAT PERSIAPAN PENGURUSAN BANJIR', bold: true, size: 22 })],
                heading: HeadingLevel.HEADING_2,
            }));
            content.push(new Paragraph({ text: '' }));
            content.push(makeTable(
                ['Perkara', 'Maklumat'],
                [
                    ['Tarikh Mesyuarat Pusingan 1', getVal('b1-mesyuarat-p1')],
                    ['Tarikh Mesyuarat Pusingan 2', getVal('b1-mesyuarat-p2')],
                    ['Jumlah Daerah Berisiko', getVal('b1-daerah-risiko')],
                    ['Jumlah Daerah Dalam Negeri', getVal('b1-daerah-negeri')],
                ]
            ));
            content.push(new Paragraph({ text: '' }));
            content.push(makeTable(
                ['Pasukan Kesihatan (Bil)', 'Ahli', 'Pasukan Perubatan (Bil)', 'Ahli', 'Jumlah Pasukan', 'Jumlah Ahli'],
                [[getVal('b1-kes-pasukan'), getVal('b1-kes-ahli'), getVal('b1-per-pasukan'), getVal('b1-per-ahli'), getVal('b1-jum-pasukan'), getVal('b1-jum-ahli')]]
            ));
            break;

        case 'rawatan':
            content.push(new Paragraph({
                children: [new TextRun({ text: 'BORANG 6 — MANGSA BANJIR DIPERIKSA', bold: true, size: 22 })],
            }));
            content.push(new Paragraph({ text: '' }));
            content.push(makeTable(
                ['Perkara', 'Bil', 'Kumulatif'],
                [
                    ['Mangsa Diperiksa', getVal('b6-diperiksa-bil'), getVal('b6-diperiksa-kum')],
                    ['Penyakit Berjangkit', getVal('b6-berjangkit-bil'), getVal('b6-berjangkit-kum')],
                    ['Penyakit Tak Berjangkit', getVal('b6-takberjangkit-bil'), getVal('b6-takberjangkit-kum')],
                    ['Kecederaan', getVal('b6-kecederaan-bil'), getVal('b6-kecederaan-kum')],
                ]
            ));
            content.push(new Paragraph({ text: '' }));
            content.push(new Paragraph({
                children: [new TextRun({ text: 'BORANG 7 — PENYAKIT BERJANGKIT', bold: true, size: 22 })],
            }));
            content.push(new Paragraph({ text: '' }));
            content.push(makeTable(
                ['Jenis Penyakit', 'Bil', 'Kumulatif'],
                [
                    ['ARI / URTI', getVal('b7-ari-bil'), getVal('b7-ari-kum')],
                    ['AGE', getVal('b7-age-bil'), getVal('b7-age-kum')],
                    ['Conjunctivitis', getVal('b7-conjunctivitis-bil'), getVal('b7-conjunctivitis-kum')],
                    ['HFMD', getVal('b7-hfmd-bil'), getVal('b7-hfmd-kum')],
                    ['Typhoid', getVal('b7-typhoid-bil'), getVal('b7-typhoid-kum')],
                    ['Leptospirosis', getVal('b7-leptospirosis-bil'), getVal('b7-leptospirosis-kum')],
                    ['Chicken Pox', getVal('b7-chickenpox-bil'), getVal('b7-chickenpox-kum')],
                    ['Skin Infection', getVal('b7-skin-bil'), getVal('b7-skin-kum')],
                    ['Other Notifiable', getVal('b7-other-bil'), getVal('b7-other-kum')],
                    ['Fever', getVal('b7-fever-bil'), getVal('b7-fever-kum')],
                    ['JUMLAH', getVal('b7-total-bil'), getVal('b7-total-kum')],
                ]
            ));
            break;

        default:
            // Generic export: collect all visible data from the module panel
            content.push(new Paragraph({
                children: [new TextRun({ text: `LAPORAN MODUL: ${moduleId.toUpperCase()}`, bold: true, size: 22 })],
            }));
            content.push(new Paragraph({ text: '' }));

            const panel = document.getElementById(`panel-${moduleId}`);
            if (panel) {
                // Find all data tables and convert them
                panel.querySelectorAll('.data-table').forEach((table, idx) => {
                    const title = table.closest('.form-section')?.querySelector('.section-title');
                    if (title) {
                        content.push(new Paragraph({
                            children: [new TextRun({ text: title.textContent, bold: true, size: 18 })],
                        }));
                        content.push(new Paragraph({ text: '' }));
                    }

                    // Extract table data
                    const headers = [];
                    table.querySelectorAll('thead th').forEach(th => headers.push(th.textContent.trim()));

                    const rows = [];
                    table.querySelectorAll('tbody tr, tfoot tr').forEach(tr => {
                        const row = [];
                        tr.querySelectorAll('td').forEach(td => {
                            const input = td.querySelector('input, select');
                            row.push(input ? input.value : td.textContent.trim());
                        });
                        rows.push(row);
                    });

                    if (headers.length > 0) {
                        content.push(makeTable(headers.length > 0 ? headers : ['Data'], rows));
                        content.push(new Paragraph({ text: '' }));
                    }
                });
            }
            break;
    }

    return content;
}

// ============================================
// UTILITY FUNCTIONS
// ============================================
function num(id) {
    const el = document.getElementById(id);
    return el ? (parseInt(el.value) || 0) : 0;
}

function setText(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
}

function sumFields(fieldIds) {
    return fieldIds.reduce((sum, id) => sum + num(id), 0);
}

function sumClass(className) {
    let total = 0;
    document.querySelectorAll(`.${className}`).forEach(el => {
        total += (parseFloat(el.value) || 0);
    });
    return total;
}

function formatDate(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    return d.toLocaleDateString('ms-MY', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;

    const icons = { success: '✓', error: '✕', info: 'ℹ' };
    toast.innerHTML = `<span>${icons[type] || 'ℹ'}</span><span>${message}</span>`;

    container.appendChild(toast);
    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateX(100%)';
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

function setConnectionStatus(state, text) {
    const badge = document.getElementById('connection-status');
    const dot = badge.querySelector('.status-dot');
    const label = badge.querySelector('.status-text');

    label.textContent = text;

    badge.style.borderColor = '';
    badge.style.background = '';
    dot.style.background = '';

    switch (state) {
        case 'loading':
            badge.style.borderColor = 'rgba(245, 158, 11, 0.3)';
            badge.style.background = 'rgba(245, 158, 11, 0.1)';
            badge.style.color = '#f59e0b';
            dot.style.background = '#f59e0b';
            break;
        case 'success':
            badge.style.borderColor = 'rgba(16, 185, 129, 0.2)';
            badge.style.background = 'rgba(16, 185, 129, 0.1)';
            badge.style.color = '#10b981';
            dot.style.background = '#10b981';
            break;
        case 'error':
            badge.style.borderColor = 'rgba(239, 68, 68, 0.3)';
            badge.style.background = 'rgba(239, 68, 68, 0.1)';
            badge.style.color = '#ef4444';
            dot.style.background = '#ef4444';
            break;
    }
}

// Export all button
document.getElementById('btn-export-all')?.addEventListener('click', () => {
    if (!validateHeader()) return;
    const modules = ['pra-bencana', 'pps', 'pendidikan', 'rawatan', 'vektor', 'makanan', 'mental', 'anggota', 'pasca'];
    modules.forEach((mod, idx) => {
        setTimeout(() => exportModule(mod), idx * 500);
    });
});
