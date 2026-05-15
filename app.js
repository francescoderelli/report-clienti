/* ============================================
   app.js - Report Clienti
   Versione rifatta da zero - 2 file
   - Tabella Clienti
   - Sum_of / Sottoprodotti
   - Avanzamento Clienti
   - Analisi TC
   - Nomi TC puliti: m.cognome@acrobaticagroup.com -> COGNOME
   - Nessuna larghezza automatica: larghezze fisse pulite
============================================ */

let pyodide = null;

const fileTab = document.getElementById("fileTabella");
const fileSum = document.getElementById("fileSum");
const btnRun = document.getElementById("btnRun");

const errModal = document.getElementById("errModal");
const errOk = document.getElementById("errOk");
const errText = document.getElementById("errText");

const okModal = document.getElementById("okModal");
const okOk = document.getElementById("okOk");
const okText = document.getElementById("okText");

const overlay = document.getElementById("loadingOverlay");
const statusTab = document.getElementById("statusTab");
const statusSum = document.getElementById("statusSum");

const state = {
  tabOk: false,
  sumOk: false,
  tabBytes: null,
  sumBytes: null,
};

function showOverlay() { if (overlay) overlay.classList.remove("hidden"); }
function hideOverlay() { if (overlay) overlay.classList.add("hidden"); }
function showModal(modal) { if (modal) modal.classList.remove("hidden"); }
function hideModal(modal) { if (modal) modal.classList.add("hidden"); }
function showError(message) {
  if (errText) errText.textContent = message || "Si è verificato un errore.";
  showModal(errModal);
}
function updateRunEnabled() {
  if (btnRun) btnRun.disabled = !(state.tabOk && state.sumOk);
}
function showFileStatus(kind, text = "Verifica in corso...") {
  const el = kind === "tab" ? statusTab : statusSum;
  if (!el) return;
  const textEl = el.querySelector(".fileStatusText");
  if (textEl) textEl.textContent = text;
  el.classList.remove("hidden");
}
function hideFileStatus(kind) {
  const el = kind === "tab" ? statusTab : statusSum;
  if (el) el.classList.add("hidden");
}
async function readAsUint8Array(file) {
  const buf = await file.arrayBuffer();
  return new Uint8Array(buf);
}

async function init() {
  updateRunEnabled();
  showOverlay();

  if (errOk) errOk.addEventListener("click", () => hideModal(errModal));
  if (okOk) okOk.addEventListener("click", () => hideModal(okModal));
  if (fileTab) fileTab.addEventListener("change", () => onFileSelected("tab"));
  if (fileSum) fileSum.addEventListener("change", () => onFileSelected("sum"));
  if (btnRun) btnRun.addEventListener("click", runReport);

  pyodide = await loadPyodide({ indexURL: "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/" });
  try {
    pyodide.setStdout({ batched: () => {} });
    pyodide.setStderr({ batched: () => {} });
  } catch (_) {}

  await pyodide.loadPackage(["pandas", "micropip"]);
  await pyodide.runPythonAsync(`
import sys, io
sys.stdout = io.StringIO()
sys.stderr = io.StringIO()
import micropip
await micropip.install(["openpyxl", "python-dateutil"])
`);

  hideOverlay();
}

init().catch((e) => {
  console.error(e);
  hideOverlay();
  showError("Errore inizializzazione: " + (e?.message || e));
});

async function onFileSelected(kind) {
  if (!pyodide) return;

  const input = kind === "tab" ? fileTab : fileSum;
  if (!input || !input.files || input.files.length !== 1) return;

  if (kind === "tab") {
    state.tabOk = false;
    state.tabBytes = null;
  } else {
    state.sumOk = false;
    state.sumBytes = null;
  }
  updateRunEnabled();
  showFileStatus(kind, "Verifica in corso...");

  try {
    const bytes = await readAsUint8Array(input.files[0]);
    pyodide.globals.set("FILE_BYTES", bytes);

    const res = await pyodide.runPythonAsync(`
import io, re
import pandas as pd

def norm_col(x):
    return re.sub(r"[^A-Z0-9]", "", str(x).strip().upper())

def norm_cols(cols):
    return [norm_col(c) for c in cols]

def score(cols):
    cset = set(cols)
    tab_s = 0
    sum_s = 0
    for m in ["IDSOGGETTO", "TIPO", "CLIENTE"]:
        if m in cset:
            tab_s += 3
    for b in ["RESPONSABILE", "REFERENTE", "STATUS", "STATO", "PREVENTIVATO", "DELIBERATO", "FATTURATO", "INCASSATO"]:
        if any(b in c for c in cols):
            tab_s += 1
    for m in ["ANNO", "MESE", "CODICESOGGETTO", "NOMESOGGETTO"]:
        if m in cset:
            sum_s += 3
    for b in ["CLASSEATTIV", "CODICEPRATICA", "LORDO", "NUMERO"]:
        if any(b in c for c in cols):
            sum_s += 1
    return tab_s, sum_s

def classify(tab_s, sum_s):
    if tab_s >= 6 and tab_s > sum_s + 1:
        return "tabella"
    if sum_s >= 6 and sum_s > tab_s + 1:
        return "sumof"
    return "unknown"

xlsx = bytes(FILE_BYTES)
best_kind = "unknown"
best_score = -1

try:
    df1 = pd.read_excel(io.BytesIO(xlsx), sheet_name=0)
    cols1 = norm_cols(df1.columns)
    t1, s1 = score(cols1)
    k1 = classify(t1, s1)
    sc1 = max(t1, s1)
    if sc1 > best_score:
        best_score = sc1
        best_kind = k1
except Exception:
    pass

try:
    df0 = pd.read_excel(io.BytesIO(xlsx), sheet_name=0, header=None)
    for i in range(min(100, len(df0))):
        row = norm_cols(df0.iloc[i].tolist())
        is_tab = ("IDSOGGETTO" in row and ("TIPO" in row or "CLIENTE" in row))
        is_sum = ("ANNO" in row and "MESE" in row and "CODICESOGGETTO" in row)
        if is_tab or is_sum:
            df2 = pd.read_excel(io.BytesIO(xlsx), sheet_name=0, header=i)
            cols2 = norm_cols(df2.columns)
            t2, s2 = score(cols2)
            k2 = classify(t2, s2)
            sc2 = max(t2, s2)
            if sc2 > best_score:
                best_score = sc2
                best_kind = k2
            break
except Exception:
    pass

best_kind
`);

    const kindFound = typeof res === "string" ? res : res.toJs();
    const expected = kind === "tab" ? "tabella" : "sumof";

    if (kindFound !== expected) {
      if (kind === "tab") {
        fileTab.value = "";
        state.tabOk = false;
        state.tabBytes = null;
      } else {
        fileSum.value = "";
        state.sumOk = false;
        state.sumBytes = null;
      }
      updateRunEnabled();
      hideFileStatus(kind);
      showError("File non riconosciuto o struttura non valida.");
      return;
    }

    if (kind === "tab") {
      state.tabOk = true;
      state.tabBytes = bytes;
    } else {
      state.sumOk = true;
      state.sumBytes = bytes;
    }

    updateRunEnabled();
    hideFileStatus(kind);
    if (okText) okText.textContent = kind === "tab" ? "Tabella Clienti verificata." : "Sum_of verificato.";
    showModal(okModal);

  } catch (e) {
    console.error(e);
    if (kind === "tab") {
      state.tabOk = false;
      state.tabBytes = null;
    } else {
      state.sumOk = false;
      state.sumBytes = null;
    }
    updateRunEnabled();
    hideFileStatus(kind);
    showError("Errore verifica file: " + (e?.message || e));
  }
}

async function runReport() {
  if (!(state.tabOk && state.sumOk)) return;

  try {
    showOverlay();
    pyodide.globals.set("TAB_BYTES", state.tabBytes);
    pyodide.globals.set("SUM_BYTES", state.sumBytes);

    const PY_REPORT = String.raw`
import io
import re
import numpy as np
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# =========================
# HELPERS
# =========================
def norm_id(x):
    if pd.isna(x):
        return ""
    s = str(x).strip().replace("\u00A0", "")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"\.0$", "", s)
    if s.lower() in ("nan", "none", "null"):
        return ""
    return s

def norm_key(x):
    return norm_id(x)

def parse_amount(x):
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip().replace("€", "").replace("\u00A0", "").replace(" ", "")
    if s == "" or s.lower() in ("nan", "none", "null"):
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def parse_count(x):
    if pd.isna(x):
        return 1
    try:
        return max(int(float(str(x).replace(",", "."))), 0)
    except Exception:
        return 1

def norm_col_name(x):
    return re.sub(r"[^A-Z0-9]", "", str(x).strip().upper())

def pick_col(df, keys, fallback_idx=None):
    cols = list(df.columns)
    upmap = {norm_col_name(c): c for c in cols}
    for k in keys:
        ku = norm_col_name(k)
        if ku in upmap:
            return upmap[ku]
    for c in cols:
        cu = norm_col_name(c)
        for k in keys:
            if norm_col_name(k) in cu:
                return c
    if fallback_idx is not None and df.shape[1] > fallback_idx:
        return df.columns[fallback_idx]
    return None

def find_header_row(xlsx_bytes, expected_type):
    df0 = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0, header=None)
    for i in range(min(100, len(df0))):
        row = [norm_col_name(c) for c in df0.iloc[i].tolist()]
        if expected_type == "tabella" and "IDSOGGETTO" in row and ("TIPO" in row or "CLIENTE" in row):
            return i
        if expected_type == "sumof" and "ANNO" in row and "MESE" in row and "CODICESOGGETTO" in row:
            return i
    return 0

def read_excel_robust(xlsx_bytes, expected_type):
    try:
        df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0)
        cols = [norm_col_name(c) for c in df.columns]
        if expected_type == "tabella" and "IDSOGGETTO" in cols:
            return df
        if expected_type == "sumof" and "CODICESOGGETTO" in cols and "ANNO" in cols:
            return df
    except Exception:
        pass
    return pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0, header=find_header_row(xlsx_bytes, expected_type))

def month_to_int(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    try:
        v = int(float(s))
        if 1 <= v <= 12:
            return v
    except Exception:
        pass
    mesi = {
        "gen":1, "gennaio":1, "feb":2, "febbraio":2, "mar":3, "marzo":3,
        "apr":4, "aprile":4, "mag":5, "maggio":5, "giu":6, "giugno":6,
        "lug":7, "luglio":7, "ago":8, "agosto":8, "set":9, "sett":9,
        "settembre":9, "ott":10, "ottobre":10, "nov":11, "novembre":11,
        "dic":12, "dicembre":12
    }
    low = s.lower()
    for k, v in mesi.items():
        if k in low:
            return v
    m = re.search(r"\b(\d{1,2})\b", s)
    if m:
        v = int(m.group(1))
        if 1 <= v <= 12:
            return v
    return np.nan

priority_map = {"07": 7, "06": 6, "04": 5, "03": 4, "05": 3, "01": 2, "02": 1}

def activity_priority(a):
    if pd.isna(a):
        return 0
    s = str(a).strip()
    m = re.match(r"^\s*(\d{2})", s)
    if m:
        return priority_map.get(m.group(1), 0)
    low = s.lower()
    if "deliber" in low:
        return 7
    if "preventiv" in low:
        return 6
    if "richiest" in low:
        return 5
    if "incontr" in low:
        return 4
    if "sopralluog" in low:
        return 3
    if "telefon" in low:
        return 2
    if "appunt" in low:
        return 1
    return 0

def activity_to_rank(v):
    if pd.isna(v):
        return 0
    s = str(v).strip()
    if s == "":
        return 0
    m = re.match(r"^\s*(\d{2})", s)
    if m:
        return {"01":1, "02":2, "05":3, "03":4, "04":5, "06":6, "07":7}.get(m.group(1), 0)
    low = s.lower()
    if "appunt" in low:
        return 1
    if "telefon" in low:
        return 2
    if "sopralluog" in low:
        return 3
    if "incontr" in low:
        return 4
    if "richiest" in low:
        return 5
    if "preventiv" in low:
        return 6
    if "deliber" in low:
        return 7
    return 0

def famiglia_stadio(v):
    r = activity_to_rank(v)
    if r in (1, 2, 4):
        return "Debole"
    if r == 3:
        return "Intermedio"
    if r in (5, 6):
        return "Forte"
    if r == 7:
        return "Convertito"
    return ""

def trend_mensile(v_old, v_m2, v_m1, v_cur):
    old_r = activity_to_rank(v_old)
    m2_r = activity_to_rank(v_m2)
    m1_r = activity_to_rank(v_m1)
    cur_r = activity_to_rank(v_cur)
    if cur_r == 7:
        return "Deliberato"
    if old_r == 0 and m2_r == 0 and m1_r == 0 and cur_r == 0:
        return "Nessuna attività"
    if cur_r == 0 and m1_r == 0 and (old_r > 0 or m2_r > 0):
        return "Dormiente"
    if m1_r == 0 and cur_r > 0:
        return "Riparte"
    if m1_r > 0 and cur_r == 0:
        return "Fermo"
    if cur_r > m1_r:
        return "Avanza"
    if cur_r == m1_r and cur_r > 0:
        return "Stabile"
    if cur_r < m1_r:
        return "Arretra"
    return "Da verificare"

def mesi_senza_miglioramento(ranks):
    vals = [int(x) for x in ranks if x is not None]
    if not vals:
        return 0
    if max(vals) >= 7:
        return 0
    best = 0
    months = 0
    for r in vals:
        if r > best:
            best = r
            months = 0
        else:
            months += 1
    return months

def da_riassegnare(last_rank, mesi_no_improve, trend, old_val, m2_val, m1_val, cur_val, anomalia):
    if last_rank == 7:
        return "No"
    fam = famiglia_stadio(cur_val)
    old_r = activity_to_rank(old_val)
    m2_r = activity_to_rank(m2_val)
    m1_r = activity_to_rank(m1_val)
    cur_r = activity_to_rank(cur_val)
    if old_r == 0 and m2_r == 0 and m1_r == 0 and cur_r == 0:
        return "Si"
    if fam == "Debole":
        if trend in ("Fermo", "Arretra", "Nessuna attività"):
            return "Si"
        if trend == "Stabile" and mesi_no_improve >= 2:
            return "Si"
        if mesi_no_improve >= 3:
            return "Si"
        return "No"
    if fam == "Intermedio":
        if trend in ("Arretra", "Nessuna attività"):
            return "Si"
        if trend == "Fermo" and mesi_no_improve >= 2:
            return "Si"
        if mesi_no_improve >= 3:
            return "Si"
        return "No"
    if fam == "Forte":
        if trend == "Arretra":
            return "Si"
        if trend == "Fermo" and mesi_no_improve >= 3:
            return "Si"
        if anomalia == "Si" and last_rank >= 5:
            return "Si"
        return "No"
    return "No"

def da_attenzionare(last_rank, mesi_no_improve, trend, anomalia, cur_val):
    if last_rank == 7:
        return "No"
    fam = famiglia_stadio(cur_val)
    if fam == "Debole":
        if trend in ("Avanza", "Riparte") and mesi_no_improve >= 1:
            return "Si"
        if trend == "Stabile" and mesi_no_improve >= 1:
            return "Si"
        if anomalia == "Si":
            return "Si"
        return "No"
    if fam == "Intermedio":
        if trend in ("Avanza", "Riparte") and mesi_no_improve >= 1:
            return "Si"
        if trend == "Stabile" and mesi_no_improve >= 2:
            return "Si"
        if trend == "Fermo":
            return "Si"
        if anomalia == "Si":
            return "Si"
        return "No"
    if fam == "Forte":
        if trend in ("Avanza", "Riparte") and mesi_no_improve >= 2:
            return "Si"
        if trend == "Stabile" and mesi_no_improve >= 2:
            return "Si"
        if trend == "Fermo":
            return "Si"
        if anomalia == "Si":
            return "Si"
        return "No"
    return "No"

def azione_consigliata(last_rank, trend, dr, da_att, anomalia):
    if anomalia == "Si":
        return "Controllo manuale"
    if dr == "Si":
        return "Valutare riassegnazione"
    if da_att == "Si":
        return "Attenzionare"
    if last_rank == 7:
        return "Chiudere e consolidare"
    if trend == "Avanza":
        return "Monitorare"
    if trend == "Riparte":
        return "Sostenere avanzamento"
    if trend == "Dormiente":
        return "Riattivare"
    if trend == "Stabile" and last_rank <= 2:
        return "Sollecitare avanzamento"
    if trend == "Stabile":
        return "Verificare evoluzione"
    if trend == "Fermo" and last_rank >= 6:
        return "Recupero commerciale"
    if trend == "Fermo":
        return "Verificare blocco"
    if trend == "Arretra":
        return "Controllo manuale"
    if trend == "Nessuna attività":
        return "Attivare contatto"
    return "Da verificare"

def stadio_riferimento(r):
    for col in [
        "Migliore attività nel periodo",
        "Ultima attività nel periodo",
        "Ultima attività",
        "Ultima attività mese precedente",
        "Ultima attività 2 mesi precedenti",
        "Ultima attività oltre 2 mesi precedenti",
    ]:
        v = r.get(col)
        if pd.notna(v) and str(v).strip() != "":
            return v
    return ""

def stato_relazione(r):
    trend = str(r.get("Trend_Mensile") or "").strip()
    fam = str(r.get("Famiglia_Stadio") or "").strip()
    last_rank = int(r.get("Ultimo_Rank", 0) or 0)
    prev_storico = float(r.get("PREVENTIVATO_STORICO_EUR", 0) or 0)
    del_storico = float(r.get("DELIBERATO_STORICO_EUR", 0) or 0)
    prev_periodo = float(r.get("PREVENTIVATO_PERIODO_EUR", 0) or 0)
    del_periodo = float(r.get("DELIBERATO_PERIODO_EUR", 0) or 0)
    n_prev = int(r.get("N_PREVENTIVI_PERIODO", 0) or 0)
    n_del = int(r.get("N_DELIBERE_PERIODO", 0) or 0)
    storico_valore = prev_storico + del_storico
    periodo_valore = prev_periodo + del_periodo
    periodo_attivo = periodo_valore > 0 or n_prev > 0 or n_del > 0
    if storico_valore > 0 and not periodo_attivo and trend in ("Dormiente", "Nessuna attività", "Fermo", "Da verificare"):
        return "Cliente da riattivare"
    if n_del >= 2:
        return "Fidelizzato"
    if n_del == 1 or del_periodo > 0:
        return "Convertito"
    if last_rank == 7:
        return "Cliente da riattivare"
    if trend == "Dormiente":
        return "Dormiente"
    if prev_periodo > 0 or n_prev > 0 or fam == "Forte":
        return "Caldo"
    if fam in ("Debole", "Intermedio") and trend in ("Avanza", "Riparte", "Stabile"):
        return "In sviluppo"
    if trend == "Nessuna attività" and storico_valore <= 0:
        return "Nuovo/Freddo"
    if trend in ("Fermo", "Arretra"):
        return "Critico"
    return "Da verificare"

def detect_anomalia(r):
    old_r = activity_to_rank(r.get("Ultima attività oltre 2 mesi precedenti"))
    m2_r = activity_to_rank(r.get("Ultima attività 2 mesi precedenti"))
    m1_r = activity_to_rank(r.get("Ultima attività mese precedente"))
    cur_r = activity_to_rank(r.get("Ultima attività"))
    if cur_r > 0 and m1_r > 0 and cur_r < m1_r:
        return "Si"
    if old_r == 7 and cur_r > 0 and cur_r < 7:
        return "Si"
    if m2_r == 7 and (m1_r > 0 and m1_r < 7):
        return "Si"
    return "No"

def status_bucket(x):
    if pd.isna(x):
        return "ALTRO"
    s = str(x).strip().lower().replace("_", " ").replace("-", " ")
    s = " ".join(s.split())
    if "potenziale" in s and "richiest" in s:
        return "POTENZIALI CON RICHIESTA"
    if "potenziale" in s:
        return "POTENZIALI"
    if "semi" in s and "attiv" in s:
        return "SEMI ATTIVO"
    if "inattiv" in s:
        return "INATTIVO"
    if "pers" in s:
        return "PERSI"
    if "attiv" in s:
        return "ATTIVO"
    return "ALTRO"

def safe_div(num, den):
    try:
        den = float(den)
        if den == 0:
            return 0
        return float(num) / den
    except Exception:
        return 0

def format_tc_name(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s == "" or s.lower() in ("nan", "none", "null"):
        return ""
    if "@" in s:
        s = s.split("@")[0]
    if "." in s:
        s = s.split(".")[-1]
    s = s.replace("_", " ").replace("-", " ")
    parts = [p for p in s.split() if p]
    if parts:
        s = parts[-1]
    return s.strip().upper()

# =========================
# LETTURA FILE
# =========================
tab = read_excel_robust(bytes(TAB_BYTES), "tabella")
su = read_excel_robust(bytes(SUM_BYTES), "sumof")

# =========================
# TAB CLIENTI
# =========================
c_id = pick_col(tab, ["ID_SOGGETTO"], fallback_idx=8)
c_tipo = pick_col(tab, ["TIPO"], fallback_idx=15)
c_status = pick_col(tab, ["STATUS", "STATO"], fallback_idx=16)
c_cli = pick_col(tab, ["CLIENTE"], fallback_idx=9)
c_ref = pick_col(tab, ["RESPONSABILE", "REFERENTE"], fallback_idx=7)
c_ca = pick_col(tab, ["CONDOMINI IN ALBERT"], fallback_idx=20)
c_cam = pick_col(tab, ["CONDOMINI AMMINISTRATI"], fallback_idx=21)
c_prev = pick_col(tab, ["PREVENTIVATO"], fallback_idx=22)
c_del = pick_col(tab, ["DELIBERATO"], fallback_idx=23)
c_fat = pick_col(tab, ["FATTURATO"], fallback_idx=24)
c_inc = pick_col(tab, ["INCASSATO"], fallback_idx=25)

clients = pd.DataFrame({
    "ID_Soggetto": tab[c_id].apply(norm_id),
    "Tipo": tab[c_tipo] if c_tipo is not None else np.nan,
    "Status_Cliente": tab[c_status] if c_status is not None else np.nan,
    "Cliente_Tabella": tab[c_cli] if c_cli is not None else np.nan,
    "Referente_Commerciale": tab[c_ref] if c_ref is not None else np.nan,
    "Condomini_in_Albert": tab[c_ca] if c_ca is not None else np.nan,
    "Condomini_Amministrati": tab[c_cam] if c_cam is not None else np.nan,
    "PREVENTIVATO_STORICO_EUR": tab[c_prev].apply(parse_amount) if c_prev is not None else 0,
    "DELIBERATO_STORICO_EUR": tab[c_del].apply(parse_amount) if c_del is not None else 0,
    "FATTURATO_EUR": tab[c_fat].apply(parse_amount) if c_fat is not None else 0,
    "INCASSATO_EUR": tab[c_inc].apply(parse_amount) if c_inc is not None else 0,
})
clients["ID_Soggetto"] = clients["ID_Soggetto"].astype(str).str.strip()

# =========================
# SUM_OF
# =========================
s_anno = pick_col(su, ["ANNO"], fallback_idx=0)
s_mese = pick_col(su, ["MESE"], fallback_idx=1)
s_data = pick_col(su, ["DATA"], fallback_idx=3)
s_att = pick_col(su, ["CLASSE ATTIVITÀ", "CLASSE ATTIVITA", "ATTIVITA", "ATTIVITÀ"], fallback_idx=4)
s_chi = pick_col(su, ["RESPONSABILE", "CHI"], fallback_idx=6)
s_cod = pick_col(su, ["CODICESOGGETTO", "CODICE SOGGETTO"], fallback_idx=8)
s_nome = pick_col(su, ["NOMESOGGETTO", "NOME SOGGETTO"], fallback_idx=9)
s_pratica = pick_col(su, ["CODICEPRATICA", "CODICE PRATICA"], fallback_idx=10)
s_numero = pick_col(su, ["NUMERO"], fallback_idx=12)
s_importo = pick_col(su, ["LORDO€", "LORDO", "IMPORTO", "SUM_OF_IMPORTO"], fallback_idx=13)

sumdf = pd.DataFrame({
    "Anno": su[s_anno],
    "Mese": su[s_mese],
    "Data": su[s_data] if s_data is not None else np.nan,
    "Attivita": su[s_att],
    "Chi": su[s_chi],
    "ID_Soggetto": su[s_cod].apply(norm_id),
    "Nome_Soggetto_Sum": su[s_nome],
    "CodicePratica": su[s_pratica].apply(norm_key) if s_pratica is not None else "",
    "Numero": su[s_numero].apply(parse_count) if s_numero is not None else 1,
    "Importo_EUR": su[s_importo].apply(parse_amount) if s_importo is not None else 0,
})
sumdf["ID_Soggetto"] = sumdf["ID_Soggetto"].astype(str).str.strip()
sumdf["Anno"] = pd.to_numeric(sumdf["Anno"], errors="coerce").astype("Int64")
sumdf["Mese_num"] = sumdf["Mese"].apply(month_to_int).astype("Int64")
sumdf["Data_dt"] = pd.to_datetime(sumdf["Data"], errors="coerce")
sumdf["Prio"] = sumdf["Attivita"].apply(activity_priority).astype(int)
sumdf["Periodo"] = (sumdf["Anno"] * 100 + sumdf["Mese_num"]).astype("Int64")
sumdf = sumdf[(sumdf["ID_Soggetto"] != "")].dropna(subset=["Periodo"]).copy()
sumdf["_row"] = np.arange(len(sumdf))
sumdf["_DataSort"] = sumdf["Data_dt"].fillna(pd.Timestamp("1900-01-01"))

preventivi_periodo = (
    sumdf[sumdf["Prio"] == 6]
    .groupby("ID_Soggetto", as_index=False)
    .agg(PREVENTIVATO_PERIODO_EUR=("Importo_EUR", "sum"), N_PREVENTIVI_PERIODO=("Numero", "sum"))
)

delibere_raw = sumdf[sumdf["Prio"] == 7].copy()
if len(delibere_raw):
    delibere_raw["_pratica_key"] = delibere_raw["CodicePratica"].apply(norm_key)
    delibere_raw.loc[delibere_raw["_pratica_key"] == "", "_pratica_key"] = delibere_raw.loc[
        delibere_raw["_pratica_key"] == "",
        "_row"
    ].apply(lambda x: f"NOCP_{x}")
    delibere_latest = (
        delibere_raw.sort_values(["ID_Soggetto", "_pratica_key", "Data_dt", "Periodo", "_row"])
        .groupby(["ID_Soggetto", "_pratica_key"], as_index=False)
        .tail(1)
    )
    delibere_periodo = (
        delibere_latest.groupby("ID_Soggetto", as_index=False)
        .agg(DELIBERATO_PERIODO_EUR=("Importo_EUR", "sum"), N_DELIBERE_PERIODO=("_pratica_key", "nunique"))
    )
else:
    delibere_latest = pd.DataFrame(columns=list(sumdf.columns) + ["_pratica_key"])
    delibere_periodo = pd.DataFrame(columns=["ID_Soggetto", "DELIBERATO_PERIODO_EUR", "N_DELIBERE_PERIODO"])

# =========================
# ULTIME ATTIVITÀ
# =========================
best_in_month = (
    sumdf.sort_values(["ID_Soggetto", "Periodo", "Prio", "_row"])
    .groupby(["ID_Soggetto", "Periodo"], as_index=False)
    .tail(1)
)
best_last = (
    best_in_month.sort_values(["ID_Soggetto", "Periodo", "Prio", "_row"])
    .groupby("ID_Soggetto", as_index=False)
    .tail(1)
)
last_act = best_last[["ID_Soggetto", "Anno", "Mese_num", "Attivita", "Chi"]].copy()
last_act.rename(columns={
    "Anno": "Anno_Ultima_Attivita",
    "Mese_num": "Mese_Ultima_Attivita",
    "Attivita": "Ultima_Attivita",
    "Chi": "Ultima_Attivita_Fatta_Da"
}, inplace=True)
last_period_act = (
    sumdf.sort_values(["ID_Soggetto", "_DataSort", "Periodo", "_row"])
    .groupby("ID_Soggetto", as_index=False)
    .tail(1)[["ID_Soggetto", "Attivita", "Chi"]]
    .copy()
)
last_period_act.rename(columns={"Attivita": "Ultima attività nel periodo", "Chi": "Ultima attività fatta da"}, inplace=True)
best_period_stage = (
    sumdf.sort_values(["ID_Soggetto", "Prio", "_DataSort", "Periodo", "_row"])
    .groupby("ID_Soggetto", as_index=False)
    .tail(1)[["ID_Soggetto", "Attivita"]]
    .copy()
)
best_period_stage.rename(columns={"Attivita": "Migliore attività nel periodo"}, inplace=True)
ultima_delibera_periodo = sumdf[sumdf["Prio"] == 7].copy()
if len(ultima_delibera_periodo):
    ultima_delibera_periodo = (
        ultima_delibera_periodo.sort_values(["ID_Soggetto", "_DataSort", "Periodo", "_row"])
        .groupby("ID_Soggetto", as_index=False)
        .tail(1)[["ID_Soggetto", "Chi", "CodicePratica", "Data_dt"]]
        .copy()
    )
    ultima_delibera_periodo.rename(columns={
        "Chi": "Ultima delibera fatta da",
        "CodicePratica": "Codice pratica ultima delibera",
        "Data_dt": "Data ultima delibera"
    }, inplace=True)
else:
    ultima_delibera_periodo = pd.DataFrame(columns=[
        "ID_Soggetto", "Ultima delibera fatta da", "Codice pratica ultima delibera", "Data ultima delibera"
    ])

name_map = (
    sumdf[["ID_Soggetto", "Nome_Soggetto_Sum"]]
    .dropna(subset=["Nome_Soggetto_Sum"])
    .drop_duplicates(subset=["ID_Soggetto"], keep="last")
)
corrispondenza = (
    clients[["ID_Soggetto", "Cliente_Tabella"]]
    .merge(name_map, on="ID_Soggetto", how="left")
    .sort_values("ID_Soggetto")
)
final = clients.merge(last_act, on="ID_Soggetto", how="left").merge(name_map, on="ID_Soggetto", how="left")
final["Cliente"] = final["Nome_Soggetto_Sum"].fillna(final["Cliente_Tabella"]).fillna(final["ID_Soggetto"])

# =========================
# AVANZAMENTO CLIENTI
# =========================
admin_mask = final["Tipo"].astype(str).str.strip().str.lower().eq("amministratore")
admins_final = final[admin_mask].copy()
admins_base = admins_final[[
    "ID_Soggetto", "Cliente", "Referente_Commerciale", "PREVENTIVATO_STORICO_EUR", "DELIBERATO_STORICO_EUR"
]].copy()
admins_base["ID_Soggetto"] = admins_base["ID_Soggetto"].astype(str).str.strip()
admins_base = admins_base.merge(preventivi_periodo, on="ID_Soggetto", how="left")
admins_base = admins_base.merge(delibere_periodo, on="ID_Soggetto", how="left")
for c in ["PREVENTIVATO_PERIODO_EUR", "DELIBERATO_PERIODO_EUR"]:
    admins_base[c] = admins_base[c].fillna(0)
for c in ["N_PREVENTIVI_PERIODO", "N_DELIBERE_PERIODO"]:
    admins_base[c] = admins_base[c].fillna(0).astype(int)

max_period = int(best_in_month["Periodo"].max()) if len(best_in_month) else None
if max_period is not None:
    y = max_period // 100
    m = max_period % 100
    cur_date = pd.Timestamp(year=y, month=m, day=1)
    prev1 = int((cur_date - pd.DateOffset(months=1)).strftime("%Y%m"))
    prev2 = int((cur_date - pd.DateOffset(months=2)).strftime("%Y%m"))
else:
    prev1 = None
    prev2 = None

def mese_anno_da_periodo(periodo):
    mesi_it = {
        1:"gennaio", 2:"febbraio", 3:"marzo", 4:"aprile", 5:"maggio", 6:"giugno",
        7:"luglio", 8:"agosto", 9:"settembre", 10:"ottobre", 11:"novembre", 12:"dicembre"
    }
    try:
        p = int(periodo)
        return f"{mesi_it.get(p % 100, str(p % 100))} {p // 100}"
    except Exception:
        return "periodo"

def mese_anno_breve_da_periodo(periodo):
    try:
        p = int(periodo)
        return f"{p % 100:02d}/{p // 100}"
    except Exception:
        return "periodo"

label_old = f"Ultima attività prima di {mese_anno_da_periodo(prev2)}" if prev2 is not None else "Ultima attività prima del periodo"
label_m2 = f"Ultima attività {mese_anno_da_periodo(prev2)}" if prev2 is not None else "Ultima attività mese -2"
label_m1 = f"Ultima attività {mese_anno_da_periodo(prev1)}" if prev1 is not None else "Ultima attività mese precedente"
label_cur = f"Attività {mese_anno_da_periodo(max_period)}" if max_period is not None else "Attività ultimo mese"
min_period = int(sumdf["Periodo"].min()) if len(sumdf) else None
max_period_report = int(sumdf["Periodo"].max()) if len(sumdf) else None
periodo_label = (
    f"da {mese_anno_breve_da_periodo(min_period)} a {mese_anno_breve_da_periodo(max_period_report)}"
    if min_period is not None and max_period_report is not None
    else "del periodo analizzato"
)

admin_months = best_in_month.merge(admins_base[["ID_Soggetto"]], on="ID_Soggetto", how="inner").copy()
cur_df = admin_months[admin_months["Periodo"] == max_period][["ID_Soggetto", "Attivita", "Chi", "Prio"]].copy() if max_period is not None else pd.DataFrame(columns=["ID_Soggetto", "Attivita", "Chi", "Prio"])
cur_df.rename(columns={"Attivita":"Ultima attività", "Chi":"Attività ultimo mese fatta da", "Prio":"Prio_cur"}, inplace=True)
m1_df = admin_months[admin_months["Periodo"] == prev1][["ID_Soggetto", "Attivita", "Prio"]].copy() if prev1 is not None else pd.DataFrame(columns=["ID_Soggetto", "Attivita", "Prio"])
m1_df.rename(columns={"Attivita":"Ultima attività mese precedente", "Prio":"Prio_m1"}, inplace=True)
m2_df = admin_months[admin_months["Periodo"] == prev2][["ID_Soggetto", "Attivita", "Prio"]].copy() if prev2 is not None else pd.DataFrame(columns=["ID_Soggetto", "Attivita", "Prio"])
m2_df.rename(columns={"Attivita":"Ultima attività 2 mesi precedenti", "Prio":"Prio_m2"}, inplace=True)
old_df = admin_months[admin_months["Periodo"] < prev2].copy() if prev2 is not None else pd.DataFrame(columns=admin_months.columns)
if len(old_df):
    old_best = (
        old_df.sort_values(["ID_Soggetto", "Periodo", "Prio", "_row"])
        .groupby("ID_Soggetto", as_index=False)
        .tail(1)[["ID_Soggetto", "Attivita", "Prio"]]
        .copy()
    )
else:
    old_best = pd.DataFrame(columns=["ID_Soggetto", "Attivita", "Prio"])
old_best.rename(columns={"Attivita":"Ultima attività oltre 2 mesi precedenti", "Prio":"Prio_old"}, inplace=True)
history_periods = sorted(admin_months["Periodo"].dropna().astype(int).unique().tolist())
hist_wide = pd.DataFrame({"ID_Soggetto": admins_base["ID_Soggetto"].unique()})
for p in history_periods:
    tmp = admin_months[admin_months["Periodo"] == p][["ID_Soggetto", "Prio"]].copy()
    tmp.rename(columns={"Prio": f"P_{p}"}, inplace=True)
    hist_wide = hist_wide.merge(tmp, on="ID_Soggetto", how="left")

avanzamento_clienti = admins_base.merge(old_best, on="ID_Soggetto", how="left")
avanzamento_clienti = avanzamento_clienti.merge(m2_df, on="ID_Soggetto", how="left")
avanzamento_clienti = avanzamento_clienti.merge(m1_df, on="ID_Soggetto", how="left")
avanzamento_clienti = avanzamento_clienti.merge(cur_df, on="ID_Soggetto", how="left")
avanzamento_clienti = avanzamento_clienti.merge(last_period_act, on="ID_Soggetto", how="left")
avanzamento_clienti = avanzamento_clienti.merge(ultima_delibera_periodo, on="ID_Soggetto", how="left")
avanzamento_clienti = avanzamento_clienti.merge(best_period_stage, on="ID_Soggetto", how="left")
avanzamento_clienti = avanzamento_clienti.merge(hist_wide, on="ID_Soggetto", how="left")
hist_cols = [c for c in avanzamento_clienti.columns if c.startswith("P_")]

def calc_mesi_no_improve_row(r):
    vals = []
    for c in hist_cols:
        v = r.get(c)
        vals.append(0 if pd.isna(v) else int(v))
    return mesi_senza_miglioramento(vals)

avanzamento_clienti["Trend_Mensile"] = avanzamento_clienti.apply(
    lambda r: trend_mensile(
        r.get("Ultima attività oltre 2 mesi precedenti"),
        r.get("Ultima attività 2 mesi precedenti"),
        r.get("Ultima attività mese precedente"),
        r.get("Ultima attività")
    ),
    axis=1
)
avanzamento_clienti["Stadio_Riferimento"] = avanzamento_clienti.apply(stadio_riferimento, axis=1)
avanzamento_clienti["Ultimo_Rank"] = avanzamento_clienti["Stadio_Riferimento"].apply(activity_to_rank)
avanzamento_clienti["Famiglia_Stadio"] = avanzamento_clienti["Stadio_Riferimento"].apply(famiglia_stadio)
avanzamento_clienti["Mesi_senza_miglioramento"] = avanzamento_clienti.apply(calc_mesi_no_improve_row, axis=1)
avanzamento_clienti["Anomalia"] = avanzamento_clienti.apply(detect_anomalia, axis=1)
avanzamento_clienti["Da_Riassegnare"] = avanzamento_clienti.apply(
    lambda r: da_riassegnare(
        int(r.get("Ultimo_Rank", 0) or 0),
        int(r.get("Mesi_senza_miglioramento", 0) or 0),
        r.get("Trend_Mensile"),
        r.get("Ultima attività oltre 2 mesi precedenti"),
        r.get("Ultima attività 2 mesi precedenti"),
        r.get("Ultima attività mese precedente"),
        r.get("Stadio_Riferimento"),
        r.get("Anomalia"),
    ),
    axis=1
)
avanzamento_clienti["Da_Attenzionare"] = avanzamento_clienti.apply(
    lambda r: da_attenzionare(
        int(r.get("Ultimo_Rank", 0) or 0),
        int(r.get("Mesi_senza_miglioramento", 0) or 0),
        r.get("Trend_Mensile"),
        r.get("Anomalia"),
        r.get("Stadio_Riferimento"),
    ),
    axis=1
)
avanzamento_clienti["Azione_Consigliata"] = avanzamento_clienti.apply(
    lambda r: azione_consigliata(
        int(r.get("Ultimo_Rank", 0) or 0),
        r.get("Trend_Mensile"),
        r.get("Da_Riassegnare"),
        r.get("Da_Attenzionare"),
        r.get("Anomalia")
    ),
    axis=1
)
avanzamento_clienti["Stato_Relazione"] = avanzamento_clienti.apply(stato_relazione, axis=1)

# Pulizia nomi in Avanzamento_Clienti.
for _col_nome in [
    "Referente_Commerciale",
    "Attività ultimo mese fatta da",
    "Ultima attività fatta da",
    "Ultima delibera fatta da",
]:
    if _col_nome in avanzamento_clienti.columns:
        avanzamento_clienti[_col_nome] = avanzamento_clienti[_col_nome].apply(format_tc_name)

mask_riattivare = avanzamento_clienti["Stato_Relazione"].eq("Cliente da riattivare")
avanzamento_clienti.loc[mask_riattivare, "Da_Riassegnare"] = "No"
avanzamento_clienti.loc[mask_riattivare, "Da_Attenzionare"] = "Si"
avanzamento_clienti.loc[mask_riattivare, "Azione_Consigliata"] = "Riattivare cliente storico"
mask_dormiente = avanzamento_clienti["Stato_Relazione"].eq("Dormiente") & avanzamento_clienti["Da_Riassegnare"].ne("Si")
avanzamento_clienti.loc[mask_dormiente, "Da_Attenzionare"] = "Si"
avanzamento_clienti.loc[mask_dormiente, "Azione_Consigliata"] = "Riattivare"
mask_freddo = avanzamento_clienti["Stato_Relazione"].eq("Nuovo/Freddo")
avanzamento_clienti.loc[mask_freddo, "Da_Riassegnare"] = "No"
avanzamento_clienti.loc[mask_freddo, "Da_Attenzionare"] = "No"
avanzamento_clienti.loc[mask_freddo, "Azione_Consigliata"] = "Attivare prospect"
mask_valore_relazione = avanzamento_clienti["Stato_Relazione"].isin(["Convertito", "Fidelizzato"])
avanzamento_clienti.loc[mask_valore_relazione, "Da_Riassegnare"] = "No"
avanzamento_clienti.loc[
    mask_valore_relazione & avanzamento_clienti["Da_Attenzionare"].eq("Si"),
    "Azione_Consigliata"
] = "Presidiare relazione"
avanzamento_clienti.loc[avanzamento_clienti["Stato_Relazione"].eq("Fidelizzato"), "Azione_Consigliata"] = "Presidiare relazione"
avanzamento_clienti.loc[
    avanzamento_clienti["Stato_Relazione"].eq("Convertito")
    & avanzamento_clienti["Azione_Consigliata"].isin(["Da verificare", "Attenzionare", "Valutare riassegnazione"]),
    "Azione_Consigliata"
] = "Consolidare relazione"

adv_cols = [
    "Cliente", "Referente_Commerciale",
    "PREVENTIVATO_STORICO_EUR", "DELIBERATO_STORICO_EUR",
    "PREVENTIVATO_PERIODO_EUR", "DELIBERATO_PERIODO_EUR",
    "N_PREVENTIVI_PERIODO", "N_DELIBERE_PERIODO",
    "Ultima attività oltre 2 mesi precedenti",
    "Ultima attività 2 mesi precedenti",
    "Ultima attività mese precedente",
    "Ultima attività",
    "Attività ultimo mese fatta da",
    "Ultima attività nel periodo",
    "Ultima attività fatta da",
    "Ultima delibera fatta da",
    "Codice pratica ultima delibera",
    "Data ultima delibera",
    "Migliore attività nel periodo",
    "Stadio_Riferimento", "Famiglia_Stadio", "Trend_Mensile", "Stato_Relazione",
    "Mesi_senza_miglioramento",
    "Da_Riassegnare", "Da_Attenzionare",
    "Azione_Consigliata", "Anomalia",
]
for col in adv_cols:
    if col not in avanzamento_clienti.columns:
        avanzamento_clienti[col] = np.nan
avanzamento_clienti = avanzamento_clienti[adv_cols].copy()

# =========================
# ANALISI TC
# =========================
clients_tc = clients.copy()
clients_tc["TC"] = clients_tc["Referente_Commerciale"].apply(format_tc_name)
clients_tc = clients_tc[clients_tc["TC"] != ""].copy()
clients_tc["Is_Admin"] = clients_tc["Tipo"].astype(str).str.strip().str.lower().eq("amministratore")
clients_tc["Status_Bucket"] = clients_tc["Status_Cliente"].apply(status_bucket)

sumdf_tc = sumdf.merge(clients[["ID_Soggetto", "Tipo"]], on="ID_Soggetto", how="left")
sumdf_tc["TC"] = sumdf_tc["Chi"].apply(format_tc_name)
sumdf_tc = sumdf_tc[sumdf_tc["TC"] != ""].copy()
sumdf_tc["Is_Admin"] = sumdf_tc["Tipo"].astype(str).str.strip().str.lower().eq("amministratore")

prev_all = (
    sumdf_tc[sumdf_tc["Prio"] == 6]
    .groupby("TC", as_index=False)
    .agg(PREVENTIVI=("Numero", "sum"))
)
prev_admin = (
    sumdf_tc[(sumdf_tc["Prio"] == 6) & (sumdf_tc["Is_Admin"])]
    .groupby("TC", as_index=False)
    .agg(PREVENTIVI_AMM=("Numero", "sum"))
)

delib_tc = delibere_latest.merge(clients[["ID_Soggetto", "Tipo"]], on="ID_Soggetto", how="left") if len(delibere_latest) else pd.DataFrame(columns=list(delibere_latest.columns) + ["Tipo"])
if len(delib_tc):
    delib_tc["TC"] = delib_tc["Chi"].apply(format_tc_name)
    delib_tc = delib_tc[delib_tc["TC"] != ""].copy()
    delib_tc["Is_Admin"] = delib_tc["Tipo"].astype(str).str.strip().str.lower().eq("amministratore")
else:
    delib_tc["TC"] = []
    delib_tc["Is_Admin"] = []

delib_all = (
    delib_tc.groupby("TC", as_index=False)
    .agg(DELIBERE=("_pratica_key", "nunique"), VENDUTO=("Importo_EUR", "sum"))
) if len(delib_tc) else pd.DataFrame(columns=["TC", "DELIBERE", "VENDUTO"])
delib_admin = (
    delib_tc[delib_tc["Is_Admin"]]
    .groupby("TC", as_index=False)
    .agg(DELIBERE_AMM=("_pratica_key", "nunique"), VENDUTO_AMM=("Importo_EUR", "sum"))
) if len(delib_tc) else pd.DataFrame(columns=["TC", "DELIBERE_AMM", "VENDUTO_AMM"])

status_counts = (
    clients_tc[clients_tc["Is_Admin"]]
    .pivot_table(index="TC", columns="Status_Bucket", values="ID_Soggetto", aggfunc="nunique", fill_value=0)
    .reset_index()
)
for col in ["ATTIVO", "SEMI ATTIVO", "INATTIVO", "PERSI", "POTENZIALI", "POTENZIALI CON RICHIESTA"]:
    if col not in status_counts.columns:
        status_counts[col] = 0

avanzamento_tc = avanzamento_clienti.copy()
avanzamento_tc["TC"] = avanzamento_tc["Referente_Commerciale"].apply(format_tc_name)
clienti_con_delibera = (
    avanzamento_tc[avanzamento_tc["N_DELIBERE_PERIODO"] > 0]
    .groupby("TC", as_index=False)
    .agg(CLIENTI_CON_DELIBERA=("Cliente", "nunique"))
)
fidelizzati_tc = (
    avanzamento_tc[avanzamento_tc["Stato_Relazione"] == "Fidelizzato"]
    .groupby("TC", as_index=False)
    .agg(FIDELIZZATI=("Cliente", "nunique"))
)

tc_list = pd.DataFrame({
    "TC": sorted(set(list(sumdf_tc["TC"].dropna().astype(str)) + list(clients_tc["TC"].dropna().astype(str))))
})
analisi_tc = tc_list.copy()
for df_merge in [prev_all, prev_admin, delib_all, delib_admin, clienti_con_delibera, fidelizzati_tc, status_counts]:
    analisi_tc = analisi_tc.merge(df_merge, on="TC", how="left")

num_cols_fill = [
    "PREVENTIVI", "PREVENTIVI_AMM",
    "DELIBERE", "DELIBERE_AMM",
    "VENDUTO", "VENDUTO_AMM",
    "CLIENTI_CON_DELIBERA", "FIDELIZZATI",
    "ATTIVO", "SEMI ATTIVO", "INATTIVO", "PERSI", "POTENZIALI", "POTENZIALI CON RICHIESTA"
]
for c in num_cols_fill:
    if c not in analisi_tc.columns:
        analisi_tc[c] = 0
    analisi_tc[c] = pd.to_numeric(analisi_tc[c], errors="coerce").fillna(0)

analisi_tc["% PREV. AMM."] = analisi_tc.apply(lambda r: safe_div(r["PREVENTIVI_AMM"], r["PREVENTIVI"]), axis=1)
analisi_tc["% DEL. AMM."] = analisi_tc.apply(lambda r: safe_div(r["DELIBERE_AMM"], r["DELIBERE"]), axis=1)
analisi_tc["% VENDUTO AMM."] = analisi_tc.apply(lambda r: safe_div(r["VENDUTO_AMM"], r["VENDUTO"]), axis=1)
analisi_tc["% CHIUSURA"] = analisi_tc.apply(lambda r: safe_div(r["DELIBERE"], r["PREVENTIVI"]), axis=1)
analisi_tc["% CHIUSURA AMM."] = analisi_tc.apply(lambda r: safe_div(r["DELIBERE_AMM"], r["PREVENTIVI_AMM"]), axis=1)
analisi_tc["DELIBERA MEDIA"] = analisi_tc.apply(lambda r: safe_div(r["VENDUTO"], r["DELIBERE"]), axis=1)
analisi_tc["DELIBERA MEDIA AMM."] = analisi_tc.apply(lambda r: safe_div(r["VENDUTO_AMM"], r["DELIBERE_AMM"]), axis=1)

analisi_tc = analisi_tc[[
    "TC",
    "PREVENTIVI",
    "DELIBERE",
    "PREVENTIVI_AMM",
    "% PREV. AMM.",
    "DELIBERE_AMM",
    "% DEL. AMM.",
    "VENDUTO",
    "VENDUTO_AMM",
    "% VENDUTO AMM.",
    "% CHIUSURA",
    "% CHIUSURA AMM.",
    "DELIBERA MEDIA",
    "DELIBERA MEDIA AMM.",
    "CLIENTI_CON_DELIBERA",
    "FIDELIZZATI",
    "ATTIVO",
    "SEMI ATTIVO",
    "INATTIVO",
    "PERSI",
    "POTENZIALI",
    "POTENZIALI CON RICHIESTA",
]].copy()
metric_cols_tc = [c for c in analisi_tc.columns if c != "TC" and not c.startswith("%")]
if len(metric_cols_tc):
    analisi_tc = analisi_tc[analisi_tc[metric_cols_tc].sum(axis=1) != 0].copy()

tot = {c: 0 for c in analisi_tc.columns}
tot["TC"] = "TOTALE"
for c in num_cols_fill:
    if c in analisi_tc.columns:
        tot[c] = analisi_tc[c].sum()
tot["% PREV. AMM."] = safe_div(tot["PREVENTIVI_AMM"], tot["PREVENTIVI"])
tot["% DEL. AMM."] = safe_div(tot["DELIBERE_AMM"], tot["DELIBERE"])
tot["% VENDUTO AMM."] = safe_div(tot["VENDUTO_AMM"], tot["VENDUTO"])
tot["% CHIUSURA"] = safe_div(tot["DELIBERE"], tot["PREVENTIVI"])
tot["% CHIUSURA AMM."] = safe_div(tot["DELIBERE_AMM"], tot["PREVENTIVI_AMM"])
tot["DELIBERA MEDIA"] = safe_div(tot["VENDUTO"], tot["DELIBERE"])
tot["DELIBERA MEDIA AMM."] = safe_div(tot["VENDUTO_AMM"], tot["DELIBERE_AMM"])
analisi_tc = pd.concat([analisi_tc.sort_values("VENDUTO", ascending=False), pd.DataFrame([tot])], ignore_index=True)

# =========================
# SINTESI
# =========================
status_sort_map = {
    "Deliberato": 1,
    "Avanza": 2,
    "Riparte": 3,
    "Stabile": 4,
    "Fermo": 5,
    "Dormiente": 6,
    "Arretra": 7,
    "Nessuna attività": 8,
    "Da verificare": 9,
}
avanzamento_clienti["_trend_sort"] = avanzamento_clienti["Trend_Mensile"].map(status_sort_map).fillna(99)
avanzamento_clienti["_mesi_sort"] = pd.to_numeric(avanzamento_clienti["Mesi_senza_miglioramento"], errors="coerce").fillna(-1)
avanzamento_clienti = (
    avanzamento_clienti
    .sort_values(["_trend_sort", "_mesi_sort", "Cliente"], ascending=[True, False, True])
    .drop(columns=["_trend_sort", "_mesi_sort"])
)
sintesi_avanzamento = (
    avanzamento_clienti
    .groupby(["Migliore attività nel periodo", "Trend_Mensile"], dropna=False)
    .size()
    .reset_index(name="N_Clienti")
    .sort_values(["Migliore attività nel periodo", "Trend_Mensile", "N_Clienti"], ascending=[True, True, False])
)
sintesi_stato = (
    avanzamento_clienti
    .groupby(["Trend_Mensile"], dropna=False)
    .size()
    .reset_index(name="N_Clienti")
    .sort_values("N_Clienti", ascending=False)
)
sintesi_per_referente = (
    avanzamento_clienti
    .groupby(["Referente_Commerciale", "Trend_Mensile"], dropna=False)
    .size()
    .reset_index(name="N_Clienti")
    .pivot_table(index="Referente_Commerciale", columns="Trend_Mensile", values="N_Clienti", aggfunc="sum", fill_value=0)
    .reset_index()
)
da_riassegnare_df = avanzamento_clienti[avanzamento_clienti["Da_Riassegnare"] == "Si"].copy()
da_attenzionare_df = avanzamento_clienti[avanzamento_clienti["Da_Attenzionare"] == "Si"].copy()
anomalie_df = avanzamento_clienti[avanzamento_clienti["Anomalia"] == "Si"].copy()

regole_ra = pd.DataFrame([
    ["SCOPO DEL FILE", "Supportare il monitoraggio commerciale degli amministratori e dei TC."],
    ["PERIODO ANALIZZATO", periodo_label],
    ["FOGLIO ANALISI_TC", "Riepiloga per TC preventivi, delibere, venduto, percentuali sugli amministratori, chiusura e stato portafoglio."],
    ["NOME TC", "I nomi dei TC vengono riportati come cognome maiuscolo. Esempio: m.guaglianone@acrobaticagroup.com diventa GUAGLIANONE."],
    ["AVANZAMENTO_CLIENTI", "Contiene una riga per ogni amministratore con valore storico, valore periodo, attività recenti e azione consigliata."],
    ["DA_RIASSEGNARE", "Indica casi in cui valutare riassegnazione commerciale."],
    ["DA_ATTENZIONARE", "Indica casi da monitorare senza riassegnazione immediata."],
], columns=["Voce", "Spiegazione"])

output_cols = [
    "Cliente", "Referente_Commerciale", "Condomini_in_Albert", "Condomini_Amministrati",
    "Anno_Ultima_Attivita", "Mese_Ultima_Attivita", "Ultima_Attivita", "Ultima_Attivita_Fatta_Da",
    "PREVENTIVATO_STORICO_EUR", "DELIBERATO_STORICO_EUR", "FATTURATO_EUR", "INCASSATO_EUR"
]
header_overrides = {
    "PREVENTIVATO_STORICO_EUR": "Preventivato Storico €",
    "DELIBERATO_STORICO_EUR": "Deliberato Storico €",
    "FATTURATO_EUR": "Fatturato €",
    "INCASSATO_EUR": "Incassato €"
}

# =========================
# SCRITTURA EXCEL
# =========================
out = io.BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    analisi_tc.to_excel(writer, sheet_name="Analisi_TC", index=False)
    riepilogo = (
        final.assign(Tipo=final["Tipo"].fillna("Senza_Tipo"))
        .groupby("Tipo", dropna=False)
        .size()
        .reset_index(name="N_clienti")
        .sort_values("N_clienti", ascending=False)
    )
    riepilogo.to_excel(writer, sheet_name="Riepilogo", index=False)
    corrispondenza.to_excel(writer, sheet_name="Corrispondenza", index=False)
    avanzamento_clienti.to_excel(writer, sheet_name="Avanzamento_Clienti", index=False)
    sintesi_avanzamento.to_excel(writer, sheet_name="Sintesi_Avanzamento", index=False)
    sintesi_stato.to_excel(writer, sheet_name="Sintesi_Stato", index=False)
    sintesi_per_referente.to_excel(writer, sheet_name="Sintesi_Per_Referente", index=False)
    da_riassegnare_df.to_excel(writer, sheet_name="Da_Riassegnare", index=False)
    da_attenzionare_df.to_excel(writer, sheet_name="Da_Attenzionare", index=False)
    anomalie_df.to_excel(writer, sheet_name="Anomalie", index=False)

    used = {
        "00_Regole_RA", "Analisi_TC",
        "Riepilogo", "Corrispondenza", "Avanzamento_Clienti",
        "Sintesi_Avanzamento", "Sintesi_Stato", "Sintesi_Per_Referente",
        "Da_Riassegnare", "Da_Attenzionare", "Anomalie"
    }
    for tipo, df_t in final.groupby(final["Tipo"].fillna("Senza_Tipo"), dropna=False):
        sheet = re.sub(r'[:\\/\?\*\[\]]', '-', str(tipo).strip() if str(tipo).strip() else "Senza_Tipo")[:31]
        base = sheet
        k = 1
        while sheet in used:
            k += 1
            suf = f"_{k}"
            sheet = (base[:31-len(suf)] + suf)[:31]
        used.add(sheet)
        df_t.copy()[output_cols].to_excel(writer, sheet_name=sheet, index=False)

    wb = writer.book
    if "Analisi_TC" in wb.sheetnames:
        wb.active = wb.sheetnames.index("Analisi_TC")

    visible_sheets = {"Analisi_TC", "Avanzamento_Clienti", "Sintesi_Per_Referente", "Da_Riassegnare"}
    for ws in wb.worksheets:
        if ws.title not in visible_sheets:
            ws.sheet_state = "hidden"

    def format_avanzamento_like_sheet(sheet_name):
        if sheet_name not in wb.sheetnames:
            return
        ws = wb[sheet_name]
        header = [c.value for c in ws[1]]
        rename_map = {
            "Ultima attività oltre 2 mesi precedenti": label_old,
            "Ultima attività 2 mesi precedenti": label_m2,
            "Ultima attività mese precedente": label_m1,
            "Ultima attività": label_cur,
            "Ultima attività nel periodo": f"Ultima attività {periodo_label}",
            "Migliore attività nel periodo": f"Migliore attività {periodo_label}",
            "PREVENTIVATO_STORICO_EUR": "Preventivato Storico €",
            "DELIBERATO_STORICO_EUR": "Deliberato Storico €",
            "PREVENTIVATO_PERIODO_EUR": f"Preventivato € {periodo_label}",
            "DELIBERATO_PERIODO_EUR": f"Deliberato € {periodo_label}",
            "N_PREVENTIVI_PERIODO": f"N° Preventivi {periodo_label}",
            "N_DELIBERE_PERIODO": f"N° Delibere {periodo_label}",
        }
        for raw, pretty in rename_map.items():
            if raw in header:
                idx = header.index(raw) + 1
                ws.cell(row=1, column=idx).value = pretty
                if "EUR" in raw:
                    for r in range(2, ws.max_row + 1):
                        ws.cell(r, idx).number_format = u'€ #,##0.00'
        header = [c.value for c in ws[1]]
        if "Data ultima delibera" in header:
            idx = header.index("Data ultima delibera") + 1
            for r in range(2, ws.max_row + 1):
                ws.cell(r, idx).number_format = "dd/mm/yyyy"
        hidden_headers = {"Stadio_Riferimento", "Famiglia_Stadio", "Trend_Mensile", "Stato_Relazione", "Anomalia"}
        for c_idx, h in enumerate(header, start=1):
            if h in hidden_headers:
                ws.column_dimensions[get_column_letter(c_idx)].hidden = True

    for sname in ["Avanzamento_Clienti", "Da_Riassegnare", "Da_Attenzionare", "Anomalie"]:
        format_avanzamento_like_sheet(sname)

    if "Analisi_TC" in wb.sheetnames:
        ws = wb["Analisi_TC"]
        ws.freeze_panes = "A2"
        rename_map_tc = {
            "PREVENTIVI_AMM": "PREVENTIVI AMM.",
            "DELIBERE_AMM": "DELIBERE AMM.",
            "VENDUTO_AMM": "VENDUTO AMM.",
            "CLIENTI_CON_DELIBERA": "CLIENTI CON DELIBERA",
        }
        for cell in ws[1]:
            if cell.value in rename_map_tc:
                cell.value = rename_map_tc[cell.value]
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        header = [c.value for c in ws[1]]
        euro_headers = {"VENDUTO", "VENDUTO AMM.", "DELIBERA MEDIA", "DELIBERA MEDIA AMM."}
        pct_headers = {"% PREV. AMM.", "% DEL. AMM.", "% VENDUTO AMM.", "% CHIUSURA", "% CHIUSURA AMM."}
        for c_idx, h in enumerate(header, start=1):
            if h in euro_headers:
                for r in range(2, ws.max_row + 1):
                    ws.cell(r, c_idx).number_format = u'€ #,##0.00'
            if h in pct_headers:
                for r in range(2, ws.max_row + 1):
                    ws.cell(r, c_idx).number_format = '0.00%'
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(r, 1).value or "").strip().upper() == "TOTALE":
                for c in range(1, ws.max_column + 1):
                    ws.cell(r, c).font = Font(bold=True)
                    ws.cell(r, c).fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

    if "Avanzamento_Clienti" in wb.sheetnames:
        ws = wb["Avanzamento_Clienti"]
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    RED = PatternFill(fill_type="solid", fgColor="FFC7CE")
    YELLOW = PatternFill(fill_type="solid", fgColor="FFF2CC")
    if "Da_Riassegnare" in wb.sheetnames:
        ws = wb["Da_Riassegnare"]
        for r in range(2, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = RED
    if "Da_Attenzionare" in wb.sheetnames:
        ws = wb["Da_Attenzionare"]
        for r in range(2, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = YELLOW
    if "Anomalie" in wb.sheetnames:
        ws = wb["Anomalie"]
        for r in range(2, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = RED

    if "00_Regole_RA" in wb.sheetnames:
        ws = wb["00_Regole_RA"]
        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 120
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    euro_format = u'€ #,##0.00'
    euro_cols = [9, 10, 11, 12]
    type_sheets = [
        s for s in wb.sheetnames
        if s not in (
            "00_Regole_RA", "Analisi_TC", "Riepilogo", "Corrispondenza", "Avanzamento_Clienti",
            "Sintesi_Avanzamento", "Sintesi_Stato", "Sintesi_Per_Referente",
            "Da_Riassegnare", "Da_Attenzionare", "Anomalie"
        )
    ]
    for sname in type_sheets:
        ws = wb[sname]
        for col_idx in euro_cols:
            cur = ws.cell(row=1, column=col_idx).value
            if cur in header_overrides:
                ws.cell(row=1, column=col_idx).value = header_overrides[cur]
        for r in range(2, ws.max_row + 1):
            for c in euro_cols:
                ws.cell(row=r, column=c).number_format = euro_format

    # Bordi: sottili sulle celle interne, spessi sulle intestazioni.
    thin_side = Side(style="thin", color="D9D9D9")
    thick_side = Side(style="medium", color="000000")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    header_border = Border(left=thick_side, right=thick_side, top=thick_side, bottom=thick_side)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.border = thin_border
        for cell in ws[1]:
            cell.border = header_border
            cell.font = Font(bold=True)

    # Larghezze fisse: nessuna larghezza automatica.
    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        if ws.title == "Analisi_TC":
            fixed_widths = {
                "A": 18, "B": 12, "C": 12, "D": 14, "E": 12,
                "F": 14, "G": 12, "H": 15, "I": 15, "J": 12,
                "K": 12, "L": 14, "M": 15, "N": 17,
                "O": 18, "P": 12, "Q": 12, "R": 12,
                "S": 12, "T": 12, "U": 12, "V": 20,
            }
        elif ws.title == "Avanzamento_Clienti":
            fixed_widths = {
                "A": 34, "B": 20, "C": 16, "D": 16,
                "E": 18, "F": 18, "G": 16, "H": 16,
                "I": 22, "J": 22, "K": 22, "L": 22,
                "M": 20, "N": 22, "O": 20, "P": 20,
                "Q": 20, "R": 14, "S": 22, "T": 16,
                "U": 16, "V": 16, "W": 16, "X": 16,
                "Y": 16, "Z": 16, "AA": 24, "AB": 12,
            }
        elif ws.title not in ("00_Regole_RA",):
            fixed_widths = {
                "A": 34, "B": 20, "C": 16, "D": 16,
                "E": 14, "F": 14, "G": 22, "H": 20,
                "I": 16, "J": 16, "K": 16, "L": 16,
            }
        else:
            fixed_widths = {}

        for col_letter, width in fixed_widths.items():
            ws.column_dimensions[col_letter].width = width

out.seek(0)
OUT_BYTES = out.read()
`;

    await pyodide.runPythonAsync(PY_REPORT);
    const outBytes = pyodide.globals.get("OUT_BYTES");
    const u8 = new Uint8Array(outBytes.toJs());
    const blob = new Blob([u8], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    saveAs(blob, "Report_Tipo_Clienti.xlsx");
    hideOverlay();

  } catch (e) {
    console.error("Errore generazione report completo:", e);
    hideOverlay();
    showError("Errore generazione report: " + (e?.message || e));
  }
}
