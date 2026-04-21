/* ============================================
   app.js - Report Clienti
   Versione completa con:
   - verifica singolo file
   - lettura robusta header Excel
   - report standard
   - avanzamento clienti
   - sintesi avanzamento
   - nessun log visibile
============================================ */

let pyodide = null;

const fileTab = document.getElementById("fileTabella");
const fileSum = document.getElementById("fileSum");
const btnRun  = document.getElementById("btnRun");

const errModal = document.getElementById("errModal");
const errOk    = document.getElementById("errOk");
const okModal  = document.getElementById("okModal");
const okOk     = document.getElementById("okOk");
const okText   = document.getElementById("okText");

const overlay  = document.getElementById("loadingOverlay");

const state = {
  tabOk: false,
  sumOk: false,
  tabBytes: null,
  sumBytes: null,
};

function showOverlay() {
  if (overlay) overlay.classList.remove("hidden");
}

function hideOverlay() {
  if (overlay) overlay.classList.add("hidden");
}

function showModal(modal) {
  if (!modal) return;
  modal.classList.remove("hidden");
}

function hideModal(modal) {
  if (!modal) return;
  modal.classList.add("hidden");
}

function updateRunEnabled() {
  if (!btnRun) return;
  btnRun.disabled = !(state.tabOk && state.sumOk);
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

  pyodide = await loadPyodide({
    indexURL: "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/"
  });

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
  showModal(errModal);
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

  try {
    const bytes = await readAsUint8Array(input.files[0]);

    pyodide.globals.set("FILE_BYTES", bytes);
    pyodide.globals.set("EXPECTED_KIND", kind === "tab" ? "tabella" : "sumof");

    const res = await pyodide.runPythonAsync(`
import io
import pandas as pd

def norm_cols(cols):
    return [str(c).strip().upper() for c in cols]

def score(cols):
    tab_must = {"ID_SOGGETTO", "TIPO", "CLIENTE"}
    tab_bonus = [
        "RESPONSABILE", "RESPONSABILEAREA",
        "CONDOMINI IN ALBERT", "CONDOMINI AMMINISTRATI",
        "PREVENTIVATO", "DELIBERATO", "FATTURATO", "INCASSATO"
    ]
    sum_must = {"ANNO", "MESE", "CODICESOGGETTO", "NOMESOGGETTO"}
    sum_bonus = ["CLASSE ATTIV"]

    cset = set(cols)
    tab_s = 0
    sum_s = 0

    for m in tab_must:
        if m in cset:
            tab_s += 3
    for b in tab_bonus:
        if any(b in c for c in cols):
            tab_s += 1

    for m in sum_must:
        if m in cset:
            sum_s += 3
    for b in sum_bonus:
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

# Tentativo 1: header standard
try:
    df1 = pd.read_excel(io.BytesIO(xlsx), sheet_name=0)
    cols1 = norm_cols(df1.columns)
    t1, s1 = score(cols1)
    k1 = classify(t1, s1)
    sc1 = max(t1, s1)
    if sc1 > best_score:
        best_score = sc1
        best_kind = k1
except:
    pass

# Tentativo 2: cerca la riga delle intestazioni vere
try:
    df0 = pd.read_excel(io.BytesIO(xlsx), sheet_name=0, header=None)

    for i in range(min(80, len(df0))):
        row = df0.iloc[i].astype(str).str.upper().str.strip().tolist()

        is_tab = ("ID_SOGGETTO" in row and ("TIPO" in row or "CLIENTE" in row))
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
except:
    pass

best_kind
`);

    const kindFound = typeof res === "string" ? res : res.toJs();

    let ok = false;
    if (kind === "tab") ok = (kindFound === "tabella");
    if (kind === "sum") ok = (kindFound === "sumof");

    if (!ok) {
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
      showModal(errModal);
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

    if (okText) {
      okText.textContent = kind === "tab"
        ? "Tabella Clienti verificata."
        : "Sum_of verificato.";
    }
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
    showModal(errModal);
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
from datetime import date
from dateutil.relativedelta import relativedelta
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill

def norm_id(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    s = s.replace("\u00A0", "")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"\.0$", "", s)
    return s

def sanitize_sheet_name(name: str) -> str:
    name = "Senza_Tipo" if name is None or str(name).strip() == "" or str(name).lower() == "nan" else str(name).strip()
    name = re.sub(r'[:\\\\/\\?\\*\\[\\]]', '-', name)
    return name[:31]

def month_to_int(x):
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    try:
        v = int(float(s))
        if 1 <= v <= 12:
            return v
    except:
        pass

    m = re.search(r"\\b(\\d{1,2})\\b", s)
    if m:
        v = int(m.group(1))
        if 1 <= v <= 12:
            return v

    mesi = {
        "gen":1, "gennaio":1,
        "feb":2, "febbraio":2,
        "mar":3, "marzo":3,
        "apr":4, "aprile":4,
        "mag":5, "maggio":5,
        "giu":6, "giugno":6,
        "lug":7, "luglio":7,
        "ago":8, "agosto":8,
        "set":9, "sett":9, "settembre":9,
        "ott":10, "ottobre":10,
        "nov":11, "novembre":11,
        "dic":12, "dicembre":12
    }
    low = s.lower()
    for k, v in mesi.items():
        if k in low:
            return v

    return np.nan

priority_map = {"07":7, "06":6, "04":5, "03":4, "05":3, "01":2, "02":1}

def activity_priority(a):
    if pd.isna(a):
        return 0
    s = str(a).strip()
    m = re.match(r"^\\s*(\\d{2})", s)
    if m:
        return priority_map.get(m.group(1), 0)
    m2 = re.search(r"\\b(0[1-7])\\b", s)
    if m2:
        return priority_map.get(m2.group(1), 0)
    return 0

def period_to_year_month(period):
    period = int(period)
    anno = period // 100
    mese = period % 100
    return anno, mese

def months_diff(period_from, period_to):
    y1, m1 = period_to_year_month(period_from)
    y2, m2 = period_to_year_month(period_to)
    return (y2 - y1) * 12 + (m2 - m1)

def find_header_row(xlsx_bytes, expected_type):
    df0 = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0, header=None)
    for i in range(min(80, len(df0))):
        row = df0.iloc[i].astype(str).str.upper().str.strip().tolist()
        if expected_type == "tabella":
            if "ID_SOGGETTO" in row and ("TIPO" in row or "CLIENTE" in row):
                return i
        elif expected_type == "sumof":
            if "ANNO" in row and "MESE" in row and "CODICESOGGETTO" in row:
                return i
    return 0

def read_excel_robust(xlsx_bytes, expected_type):
    try:
        df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0)
        cols = [str(c).strip().upper() for c in df.columns]
        if expected_type == "tabella" and "ID_SOGGETTO" in cols:
            return df
        if expected_type == "sumof" and "CODICESOGGETTO" in cols and "ANNO" in cols:
            return df
    except:
        pass

    header_row = find_header_row(xlsx_bytes, expected_type)
    return pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0, header=header_row)

def pick_col(df, keys, fallback_idx=None):
    cols = list(df.columns)
    upmap = {str(c).strip().upper(): c for c in cols}

    for k in keys:
        ku = k.upper()
        if ku in upmap:
            return upmap[ku]

    for c in cols:
        cu = str(c).upper()
        for k in keys:
            if k.upper() in cu:
                return c

    if fallback_idx is not None and df.shape[1] > fallback_idx:
        return df.columns[fallback_idx]

    return None

# =========================================================
# LETTURA FILE
# =========================================================

tab = read_excel_robust(bytes(TAB_BYTES), "tabella")
su  = read_excel_robust(bytes(SUM_BYTES), "sumof")

# =========================================================
# TAB CLIENTI
# =========================================================

c_id   = pick_col(tab, ["ID_SOGGETTO"], fallback_idx=8)
c_tipo = pick_col(tab, ["TIPO"], fallback_idx=15)
c_cli  = pick_col(tab, ["CLIENTE"], fallback_idx=9)
c_ref  = pick_col(tab, ["RESPONSABILE", "REFERENTE"], fallback_idx=7)

c_ca   = pick_col(tab, ["CONDOMINI IN ALBERT"], fallback_idx=20)
c_cam  = pick_col(tab, ["CONDOMINI AMMINISTRATI"], fallback_idx=21)
c_prev = pick_col(tab, ["PREVENTIVATO"], fallback_idx=22)
c_del  = pick_col(tab, ["DELIBERATO"], fallback_idx=23)
c_fat  = pick_col(tab, ["FATTURATO"], fallback_idx=24)
c_inc  = pick_col(tab, ["INCASSATO"], fallback_idx=25)

clients = pd.DataFrame({
    "ID_Soggetto": tab[c_id].apply(norm_id),
    "Tipo": tab[c_tipo] if c_tipo else np.nan,
    "Cliente_Tabella": tab[c_cli] if c_cli else np.nan,
    "Referente_Commerciale": tab[c_ref] if c_ref else np.nan,
    "Condomini_in_Albert": tab[c_ca] if c_ca else np.nan,
    "Condomini_Amministrati": tab[c_cam] if c_cam else np.nan,
    "PREVENTIVATO_EUR": tab[c_prev] if c_prev else np.nan,
    "DELIBERATO_EUR": tab[c_del] if c_del else np.nan,
    "FATTURATO_EUR": tab[c_fat] if c_fat else np.nan,
    "INCASSATO_EUR": tab[c_inc] if c_inc else np.nan,
})

# =========================================================
# SUM OF
# =========================================================

s_anno = pick_col(su, ["ANNO"], fallback_idx=0)
s_mese = pick_col(su, ["MESE"], fallback_idx=1)
s_att  = pick_col(su, ["CLASSE ATTIVITÀ", "CLASSE ATTIVITA", "ATTIVITA", "ATTIVITÀ"], fallback_idx=2)
s_chi  = pick_col(su, ["RESPONSABILE", "CHI"], fallback_idx=4)
s_cod  = pick_col(su, ["CODICESOGGETTO", "CODICE SOGGETTO"], fallback_idx=6)
s_nome = pick_col(su, ["NOMESOGGETTO", "NOME SOGGETTO"], fallback_idx=7)

sumdf = pd.DataFrame({
    "Anno": su[s_anno],
    "Mese": su[s_mese],
    "Attivita": su[s_att],
    "Chi": su[s_chi],
    "ID_Soggetto": su[s_cod].apply(norm_id),
    "Nome_Soggetto_Sum": su[s_nome],
})

sumdf["Anno"] = pd.to_numeric(sumdf["Anno"], errors="coerce").astype("Int64")
sumdf["Mese_num"] = sumdf["Mese"].apply(month_to_int).astype("Int64")
sumdf["Prio"] = sumdf["Attivita"].apply(activity_priority).astype(int)
sumdf["Periodo"] = (sumdf["Anno"] * 100 + sumdf["Mese_num"]).astype("Int64")
sumdf = sumdf[(sumdf["ID_Soggetto"] != "")].dropna(subset=["Periodo"]).copy()
sumdf["_row"] = np.arange(len(sumdf))

# =========================================================
# ULTIMA ATTIVITA'
# =========================================================

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

# =========================================================
# AVANZAMENTO CLIENTI - VERSIONE ROBUSTA
# =========================================================

def stage_label_from_code(code):
    mapping = {
        "01": "Telefonata",
        "02": "Appuntamento",
        "03": "Incontro",
        "04": "Richiesta",
        "05": "Sopralluogo",
        "06": "Preventivo",
        "07": "Delibera",
    }
    return mapping.get(code, code)

def extract_stage_code_safe(a):
    if pd.isna(a):
        return ""
    s = str(a).strip().upper()
    m = re.match(r"^\\s*(\\d{2})", s)
    if m:
        return m.group(1)
    m2 = re.search(r"\\b(0[1-7])\\b", s)
    if m2:
        return m2.group(1)
    return ""

today_period = date.today().year * 100 + date.today().month

adv_base = sumdf.copy()
adv_base["Stage_Code2"] = adv_base["Attivita"].apply(extract_stage_code_safe)
adv_base["Stage_Order2"] = adv_base["Stage_Code2"].map({
    "01": 1,
    "02": 2,
    "03": 3,
    "04": 4,
    "05": 5,
    "06": 6,
    "07": 7,
}).fillna(0).astype(int)

# fallback: se non trova Stage_Order2 usa Prio, che già funziona nel report standard
adv_base["Stage_Order_Final"] = np.where(
    adv_base["Stage_Order2"] > 0,
    adv_base["Stage_Order2"],
    adv_base["Prio"]
).astype(int)

adv_base["Stage_Code_Final"] = adv_base["Stage_Code2"]

prio_to_code = {
    2: "01",
    1: "02",
    4: "03",
    5: "04",
    3: "05",
    6: "06",
    7: "07",
}
mask_missing_code = adv_base["Stage_Code_Final"].eq("") & adv_base["Stage_Order_Final"].gt(0)
adv_base.loc[mask_missing_code, "Stage_Code_Final"] = adv_base.loc[mask_missing_code, "Prio"].map(prio_to_code).fillna("")

adv_base["Stage_Name_Final"] = adv_base["Stage_Code_Final"].apply(stage_label_from_code)

adv_base = adv_base[adv_base["Stage_Order_Final"] > 0].copy()
adv_base = adv_base.sort_values(["ID_Soggetto", "Periodo", "Stage_Order_Final", "_row"]).copy()

records = []

for client_id, g in adv_base.groupby("ID_Soggetto"):
    g = g.sort_values(["Periodo", "Stage_Order_Final", "_row"]).copy()

    max_stage_so_far = 0
    current_stage_code = ""
    current_stage_name = ""
    first_period_current_stage = None

    for _, row in g.iterrows():
        row_stage = int(row["Stage_Order_Final"])
        row_code = row["Stage_Code_Final"]
        row_name = row["Stage_Name_Final"]
        row_period = int(row["Periodo"])

        # avanzamento reale solo se supera il massimo raggiunto
        if row_stage > max_stage_so_far:
            max_stage_so_far = row_stage
            current_stage_code = row_code
            current_stage_name = row_name
            first_period_current_stage = row_period

    if first_period_current_stage is None:
        continue

    last_row = g.sort_values(["Periodo", "_row"]).iloc[-1]
    last_period_seen = int(last_row["Periodo"])
    last_actor = last_row["Chi"] if "Chi" in g.columns else np.nan

    mesi_fermo = months_diff(first_period_current_stage, today_period)

    if current_stage_code == "07":
        stato_avanzamento = "Deliberato"
        da_riassegnare = "No"
    elif current_stage_code in ("01", "02", "03", "04", "05") and mesi_fermo >= 1:
        stato_avanzamento = "Da riassegnare"
        da_riassegnare = "Si"
    elif current_stage_code == "06" and mesi_fermo >= 2:
        stato_avanzamento = "Da riassegnare"
        da_riassegnare = "Si"
    elif mesi_fermo == 0:
        stato_avanzamento = "Avanza"
        da_riassegnare = "No"
    else:
        stato_avanzamento = "Fermo"
        da_riassegnare = "No"

    anno_stage, mese_stage = period_to_year_month(first_period_current_stage)
    anno_last, mese_last = period_to_year_month(last_period_seen)

    records.append({
        "ID_Soggetto": client_id,
        "Codice_Stadio_Attuale": current_stage_code,
        "Stadio_Attuale": current_stage_name,
        "Primo_Anno_Stadio_Attuale": anno_stage,
        "Primo_Mese_Stadio_Attuale": mese_stage,
        "Ultimo_Anno_Rilevato": anno_last,
        "Ultimo_Mese_Rilevato": mese_last,
        "Mesi_Fermo_Nello_Stadio": mesi_fermo,
        "Stato_Avanzamento": stato_avanzamento,
        "Da_Riassegnare": da_riassegnare,
        "Ultima_Attivita_Fatta_Da": last_actor,
    })

adv_df = pd.DataFrame(records)

if adv_df.empty:
    avanzamento_clienti = final[["ID_Soggetto", "Cliente", "Referente_Commerciale"]].copy()
    avanzamento_clienti["Codice_Stadio_Attuale"] = np.nan
    avanzamento_clienti["Stadio_Attuale"] = np.nan
    avanzamento_clienti["Primo_Anno_Stadio_Attuale"] = np.nan
    avanzamento_clienti["Primo_Mese_Stadio_Attuale"] = np.nan
    avanzamento_clienti["Ultimo_Anno_Rilevato"] = np.nan
    avanzamento_clienti["Ultimo_Mese_Rilevato"] = np.nan
    avanzamento_clienti["Mesi_Fermo_Nello_Stadio"] = np.nan
    avanzamento_clienti["Stato_Avanzamento"] = "Nessuna attività"
    avanzamento_clienti["Da_Riassegnare"] = "No"
    avanzamento_clienti["Ultima_Attivita_Fatta_Da"] = np.nan
else:
    avanzamento_clienti = final[["ID_Soggetto", "Cliente", "Referente_Commerciale"]].merge(
        adv_df,
        on="ID_Soggetto",
        how="left"
    )
    avanzamento_clienti["Stato_Avanzamento"] = avanzamento_clienti["Stato_Avanzamento"].fillna("Nessuna attività")
    avanzamento_clienti["Da_Riassegnare"] = avanzamento_clienti["Da_Riassegnare"].fillna("No")

adv_cols = [
    "Cliente",
    "Referente_Commerciale",
    "Codice_Stadio_Attuale",
    "Stadio_Attuale",
    "Primo_Anno_Stadio_Attuale",
    "Primo_Mese_Stadio_Attuale",
    "Ultimo_Anno_Rilevato",
    "Ultimo_Mese_Rilevato",
    "Mesi_Fermo_Nello_Stadio",
    "Stato_Avanzamento",
    "Da_Riassegnare",
    "Ultima_Attivita_Fatta_Da",
]

avanzamento_clienti = avanzamento_clienti[adv_cols].copy()

stage_sort_map = {
    "Telefonata": 1,
    "Appuntamento": 2,
    "Incontro": 3,
    "Richiesta": 4,
    "Sopralluogo": 5,
    "Preventivo": 6,
    "Delibera": 7
}
status_sort_map = {
    "Da riassegnare": 1,
    "Fermo": 2,
    "Avanza": 3,
    "Deliberato": 4,
    "Nessuna attività": 5
}

avanzamento_clienti["_stage_sort"] = avanzamento_clienti["Stadio_Attuale"].map(stage_sort_map).fillna(99)
avanzamento_clienti["_status_sort"] = avanzamento_clienti["Stato_Avanzamento"].map(status_sort_map).fillna(99)
avanzamento_clienti["_mesi_sort"] = pd.to_numeric(avanzamento_clienti["Mesi_Fermo_Nello_Stadio"], errors="coerce").fillna(-1)

avanzamento_clienti = (
    avanzamento_clienti
    .sort_values(["_status_sort", "_stage_sort", "_mesi_sort", "Cliente"], ascending=[True, True, False, True])
    .drop(columns=["_stage_sort", "_status_sort", "_mesi_sort"])
)

sintesi_avanzamento = (
    avanzamento_clienti
    .groupby(["Stadio_Attuale", "Stato_Avanzamento"], dropna=False)
    .size()
    .reset_index(name="N_Clienti")
    .sort_values(["Stadio_Attuale", "Stato_Avanzamento", "N_Clienti"], ascending=[True, True, False])
)

sintesi_stato = (
    avanzamento_clienti
    .groupby(["Stato_Avanzamento"], dropna=False)
    .size()
    .reset_index(name="N_Clienti")
    .sort_values("N_Clienti", ascending=False)
)

# =========================================================
# OUTPUT STANDARD
# =========================================================

output_cols = [
    "Cliente",
    "Referente_Commerciale",
    "Condomini_in_Albert",
    "Condomini_Amministrati",
    "Anno_Ultima_Attivita",
    "Mese_Ultima_Attivita",
    "Ultima_Attivita",
    "Ultima_Attivita_Fatta_Da",
    "PREVENTIVATO_EUR",
    "DELIBERATO_EUR",
    "FATTURATO_EUR",
    "INCASSATO_EUR"
]

header_overrides = {
    "PREVENTIVATO_EUR": "Preventivato €",
    "DELIBERATO_EUR": "Deliberato €",
    "FATTURATO_EUR": "Fatturato €",
    "INCASSATO_EUR": "Incassato €",
}

out = io.BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
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

    used = {
        "Riepilogo",
        "Corrispondenza",
        "Avanzamento_Clienti",
        "Sintesi_Avanzamento",
        "Sintesi_Stato"
    }

    for tipo, df_t in final.groupby(final["Tipo"].fillna("Senza_Tipo"), dropna=False):
        sheet = sanitize_sheet_name(tipo)
        base = sheet
        k = 1
        while sheet in used:
            k += 1
            suf = f"_{k}"
            sheet = (base[:31-len(suf)] + suf)[:31]
        used.add(sheet)
        df_t.copy()[output_cols].to_excel(writer, sheet_name=sheet, index=False)

    wb = writer.book

    euro_format = u'€ #,##0.00'
    euro_cols = [9, 10, 11, 12]

    type_sheets = [
        s for s in wb.sheetnames
        if s not in ("Riepilogo", "Corrispondenza", "Avanzamento_Clienti", "Sintesi_Avanzamento", "Sintesi_Stato")
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

    GREEN = PatternFill(fill_type="solid", fgColor="C6EFCE")
    RED   = PatternFill(fill_type="solid", fgColor="FFC7CE")

    cutoff = date.today() - relativedelta(months=2)
    cutoff_period = cutoff.year * 100 + cutoff.month

    admin_sheet = None
    for s in wb.sheetnames:
        if s not in ("Riepilogo", "Corrispondenza", "Avanzamento_Clienti", "Sintesi_Avanzamento", "Sintesi_Stato") and "amministr" in s.lower():
            admin_sheet = s
            break

    if admin_sheet:
        ws = wb[admin_sheet]
        header = [c.value for c in ws[1]]
        col_anno = header.index("Anno_Ultima_Attivita") + 1
        col_mese = header.index("Mese_Ultima_Attivita") + 1
        max_col = ws.max_column

        for r in range(2, ws.max_row + 1):
            anno = ws.cell(r, col_anno).value
            mese = ws.cell(r, col_mese).value

            if anno is None or mese is None or str(anno).strip() == "" or str(mese).strip() == "":
                fill = RED
            else:
                try:
                    period = int(anno) * 100 + int(mese)
                    fill = RED if period < cutoff_period else GREEN
                except:
                    fill = RED

            for c in range(1, max_col + 1):
                ws.cell(r, c).fill = fill

        wb.active = wb.sheetnames.index(admin_sheet)

    for ws in wb.worksheets:
        for col_idx, col_cells in enumerate(ws.columns, start=1):
            max_len = 0
            for cell in list(col_cells)[:2000]:
                if cell.value is None:
                    continue
                max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 45)

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
    console.error(e);
    hideOverlay();
    showModal(errModal);
  }
}
