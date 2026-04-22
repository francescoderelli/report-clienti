/* ============================================
   app.js - Report Clienti
   Versione evoluta:
   - verifica singolo file
   - errore reale in modal
   - report standard
   - avanzamento amministratori evoluto
   - sintesi per referente
   - fogli Da_Riassegnare e Anomalie
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

function showError(message) {
  if (errText) errText.textContent = message || "Si è verificato un errore.";
  showModal(errModal);
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

  try {
    const bytes = await readAsUint8Array(input.files[0]);
    pyodide.globals.set("FILE_BYTES", bytes);

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
    if (kind === "tab") ok = kindFound === "tabella";
    if (kind === "sum") ok = kindFound === "sumof";

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
from datetime import date
from dateutil.relativedelta import relativedelta
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill

def norm_id(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    s = s.replace("\\u00A0", "")
    s = re.sub(r"\\s+", "", s)
    s = re.sub(r"\\.0$", "", s)
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
        "gen":1, "gennaio":1, "feb":2, "febbraio":2,
        "mar":3, "marzo":3, "apr":4, "aprile":4,
        "mag":5, "maggio":5, "giu":6, "giugno":6,
        "lug":7, "luglio":7, "ago":8, "agosto":8,
        "set":9, "sett":9, "settembre":9,
        "ott":10, "ottobre":10, "nov":11, "novembre":11,
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

def period_to_year_month(period):
    period = int(period)
    return period // 100, period % 100

def rank_to_label(rank):
    mp = {
        0: "",
        1: "Appuntamento",
        2: "Telefonata",
        3: "Sopralluogo",
        4: "Incontro",
        5: "Richiesta",
        6: "Preventivo",
        7: "Delibera",
    }
    return mp.get(int(rank), "")

def activity_to_rank(v):
    if pd.isna(v):
        return 0
    s = str(v).strip()
    if s == "":
        return 0
    m = re.match(r"^\\s*(\\d{2})", s)
    if m:
        code = m.group(1)
        code_map = {
            "01": 1,
            "02": 2,
            "03": 4,
            "04": 5,
            "05": 3,
            "06": 6,
            "07": 7,
        }
        return code_map.get(code, 0)
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

def trend_mensile(v_old, v_m2, v_m1, v_cur):
    old_r = activity_to_rank(v_old)
    m2_r  = activity_to_rank(v_m2)
    m1_r  = activity_to_rank(v_m1)
    cur_r = activity_to_rank(v_cur)

    if cur_r == 7:
        return "Deliberato"
    if old_r == 0 and m2_r == 0 and m1_r == 0 and cur_r == 0:
        return "Nessuna attività"
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
    best = 0
    months = 0
    for r in ranks:
        if r > best:
            best = r
            months = 0
        else:
            months += 1
    return months

def da_riassegnare(last_rank, mesi_no_improve, trend, old_val, m2_val, m1_val, cur_val):
    # Delibera mai da riassegnare
    if last_rank == 7:
        return "No"

    old_r = activity_to_rank(old_val)
    m2_r  = activity_to_rank(m2_val)
    m1_r  = activity_to_rank(m1_val)
    cur_r = activity_to_rank(cur_val)

    # Nessuna attività assoluta -> da riassegnare
    if old_r == 0 and m2_r == 0 and m1_r == 0 and cur_r == 0:
        return "Si"

    # Nessuna attività recente ma storico presente -> da riassegnare
    if trend == "Nessuna attività":
        return "Si"

    # Arretra -> da riassegnare
    if trend == "Arretra":
        return "Si"

    # Fermo -> più aggressivo
    if trend == "Fermo":
        return "Si"

    # Regole per stadio
    if last_rank in (1, 2) and mesi_no_improve >= 1:
        return "Si"
    if last_rank in (3, 4, 5) and mesi_no_improve >= 1:
        return "Si"
    if last_rank == 6 and mesi_no_improve >= 2:
        return "Si"

    return "No"

def esito_manageriale(last_rank, trend, mesi_no_improve):
    if last_rank == 7:
        return "Chiuso"
    if trend == "Avanza" and last_rank >= 5:
        return "Caldo"
    if trend == "Avanza":
        return "Positivo"
    if trend == "Riparte":
        return "Da monitorare"
    if trend == "Stabile" and last_rank <= 2:
        return "Debole"
    if trend == "Stabile" and last_rank >= 5:
        return "In lavorazione"
    if trend == "Fermo" and last_rank >= 6:
        return "Critico"
    if trend == "Fermo":
        return "Bloccato"
    if trend == "Arretra":
        return "Anomalo"
    if mesi_no_improve >= 2 and last_rank > 0:
        return "Da recuperare"
    if trend == "Nessuna attività":
        return "Freddo"
    return "Da verificare"

def azione_consigliata(last_rank, trend, dr, anomalia):
    if anomalia == "Si":
        return "Controllo manuale"
    if dr == "Si":
        return "Valutare riassegnazione"
    if last_rank == 7:
        return "Chiudere e consolidare"
    if trend == "Avanza":
        return "Monitorare"
    if trend == "Riparte":
        return "Sostenere avanzamento"
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

tab = read_excel_robust(bytes(TAB_BYTES), "tabella")
su  = read_excel_robust(bytes(SUM_BYTES), "sumof")

# Tabella clienti
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
    "Tipo": tab[c_tipo] if c_tipo is not None else np.nan,
    "Cliente_Tabella": tab[c_cli] if c_cli is not None else np.nan,
    "Referente_Commerciale": tab[c_ref] if c_ref is not None else np.nan,
    "Condomini_in_Albert": tab[c_ca] if c_ca is not None else np.nan,
    "Condomini_Amministrati": tab[c_cam] if c_cam is not None else np.nan,
    "PREVENTIVATO_EUR": tab[c_prev] if c_prev is not None else np.nan,
    "DELIBERATO_EUR": tab[c_del] if c_del is not None else np.nan,
    "FATTURATO_EUR": tab[c_fat] if c_fat is not None else np.nan,
    "INCASSATO_EUR": tab[c_inc] if c_inc is not None else np.nan,
})
clients["ID_Soggetto"] = clients["ID_Soggetto"].astype(str).str.strip()

# Sum_of
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
sumdf["ID_Soggetto"] = sumdf["ID_Soggetto"].astype(str).str.strip()
sumdf["Anno"] = pd.to_numeric(sumdf["Anno"], errors="coerce").astype("Int64")
sumdf["Mese_num"] = sumdf["Mese"].apply(month_to_int).astype("Int64")
sumdf["Prio"] = sumdf["Attivita"].apply(activity_priority).astype(int)
sumdf["Periodo"] = (sumdf["Anno"] * 100 + sumdf["Mese_num"]).astype("Int64")
sumdf = sumdf[(sumdf["ID_Soggetto"] != "")].dropna(subset=["Periodo"]).copy()
sumdf["_row"] = np.arange(len(sumdf))

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

# Solo amministratori
admin_mask = final["Tipo"].astype(str).str.strip().str.lower().eq("amministratore")
admins_final = final[admin_mask].copy()
admins_base = admins_final[["ID_Soggetto", "Cliente", "Referente_Commerciale"]].copy()
admins_base["ID_Soggetto"] = admins_base["ID_Soggetto"].astype(str).str.strip()

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

admin_months = best_in_month.merge(
    admins_base[["ID_Soggetto"]],
    on="ID_Soggetto",
    how="inner"
).copy()

cur_df = admin_months[admin_months["Periodo"] == max_period][["ID_Soggetto", "Attivita", "Chi", "Prio"]].copy() if max_period is not None else pd.DataFrame(columns=["ID_Soggetto","Attivita","Chi","Prio"])
cur_df.rename(columns={"Attivita":"Ultima attività", "Chi":"Ultima_Attivita_Fatta_Da", "Prio":"Prio_cur"}, inplace=True)

m1_df = admin_months[admin_months["Periodo"] == prev1][["ID_Soggetto", "Attivita", "Prio"]].copy() if prev1 is not None else pd.DataFrame(columns=["ID_Soggetto","Attivita","Prio"])
m1_df.rename(columns={"Attivita":"Ultima attività mese precedente", "Prio":"Prio_m1"}, inplace=True)

m2_df = admin_months[admin_months["Periodo"] == prev2][["ID_Soggetto", "Attivita", "Prio"]].copy() if prev2 is not None else pd.DataFrame(columns=["ID_Soggetto","Attivita","Prio"])
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
    old_best = pd.DataFrame(columns=["ID_Soggetto","Attivita","Prio"])
old_best.rename(columns={"Attivita":"Ultima attività oltre 2 mesi precedenti", "Prio":"Prio_old"}, inplace=True)

# storico sintetico per mesi senza miglioramento
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

avanzamento_clienti["Ultimo_Rank"] = avanzamento_clienti["Ultima attività"].apply(activity_to_rank)
avanzamento_clienti["Mesi_senza_miglioramento"] = avanzamento_clienti.apply(calc_mesi_no_improve_row, axis=1)

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

avanzamento_clienti["Anomalia"] = avanzamento_clienti.apply(detect_anomalia, axis=1)
avanzamento_clienti["Da_Riassegnare"] = avanzamento_clienti.apply(
    lambda r: da_riassegnare(int(r.get("Ultimo_Rank", 0) or 0), int(r.get("Mesi_senza_miglioramento", 0) or 0)),
    axis=1
)
avanzamento_clienti["Esito_Manageriale"] = avanzamento_clienti.apply(
    lambda r: esito_manageriale(
        int(r.get("Ultimo_Rank", 0) or 0),
        r.get("Trend_Mensile"),
        int(r.get("Mesi_senza_miglioramento", 0) or 0)
    ),
    axis=1
)
avanzamento_clienti["Azione_Consigliata"] = avanzamento_clienti.apply(
    lambda r: azione_consigliata(
        int(r.get("Ultimo_Rank", 0) or 0),
        r.get("Trend_Mensile"),
        r.get("Da_Riassegnare"),
        r.get("Anomalia")
    ),
    axis=1
)

adv_cols = [
    "Cliente",
    "Referente_Commerciale",
    "Ultima attività oltre 2 mesi precedenti",
    "Ultima attività 2 mesi precedenti",
    "Ultima attività mese precedente",
    "Ultima attività",
    "Trend_Mensile",
    "Mesi_senza_miglioramento",
    "Da_Riassegnare",
    "Esito_Manageriale",
    "Azione_Consigliata",
    "Anomalia",
    "Ultima_Attivita_Fatta_Da",
]

for col in adv_cols:
    if col not in avanzamento_clienti.columns:
        avanzamento_clienti[col] = np.nan

avanzamento_clienti = avanzamento_clienti[adv_cols].copy()

status_sort_map = {
    "Deliberato": 1,
    "Avanza": 2,
    "Riparte": 3,
    "Stabile": 4,
    "Fermo": 5,
    "Arretra": 6,
    "Nessuna attività": 7,
    "Da verificare": 8,
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
    .groupby(["Ultima attività", "Trend_Mensile"], dropna=False)
    .size()
    .reset_index(name="N_Clienti")
    .sort_values(["Ultima attività", "Trend_Mensile", "N_Clienti"], ascending=[True, True, False])
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
anomalie_df = avanzamento_clienti[avanzamento_clienti["Anomalia"] == "Si"].copy()

# Output standard
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
    sintesi_per_referente.to_excel(writer, sheet_name="Sintesi_Per_Referente", index=False)
    da_riassegnare_df.to_excel(writer, sheet_name="Da_Riassegnare", index=False)
    anomalie_df.to_excel(writer, sheet_name="Anomalie", index=False)

    used = {
        "Riepilogo", "Corrispondenza", "Avanzamento_Clienti",
        "Sintesi_Avanzamento", "Sintesi_Stato", "Sintesi_Per_Referente",
        "Da_Riassegnare", "Anomalie"
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
        if s not in (
            "Riepilogo", "Corrispondenza", "Avanzamento_Clienti",
            "Sintesi_Avanzamento", "Sintesi_Stato", "Sintesi_Per_Referente",
            "Da_Riassegnare", "Anomalie"
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

    GREEN = PatternFill(fill_type="solid", fgColor="C6EFCE")
    RED   = PatternFill(fill_type="solid", fgColor="FFC7CE")
    YELLOW = PatternFill(fill_type="solid", fgColor="FFF2CC")

    if "Avanzamento_Clienti" in wb.sheetnames:
        ws = wb["Avanzamento_Clienti"]
        header = [c.value for c in ws[1]]
        try:
            col_trend = header.index("Trend_Mensile") + 1
            max_col = ws.max_column
            for r in range(2, ws.max_row + 1):
                trend = str(ws.cell(r, col_trend).value or "").strip()
                fill = None
                if trend in ("Avanza", "Riparte", "Deliberato"):
                    fill = GREEN
                elif trend in ("Fermo", "Arretra"):
                    fill = RED
                elif trend in ("Stabile", "Da verificare"):
                    fill = YELLOW
                if fill:
                    for c in range(1, max_col + 1):
                        ws.cell(r, c).fill = fill
        except:
            pass

    for special_sheet in ["Da_Riassegnare", "Anomalie"]:
        if special_sheet in wb.sheetnames:
            ws = wb[special_sheet]
            for r in range(2, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    ws.cell(r, c).fill = RED

    admin_sheet = None
    for s in wb.sheetnames:
        if s not in (
            "Riepilogo", "Corrispondenza", "Avanzamento_Clienti",
            "Sintesi_Avanzamento", "Sintesi_Stato", "Sintesi_Per_Referente",
            "Da_Riassegnare", "Anomalie"
        ) and "amministr" in s.lower():
            admin_sheet = s
            break

    if admin_sheet:
        ws = wb[admin_sheet]
        header = [c.value for c in ws[1]]
        try:
            col_anno = header.index("Anno_Ultima_Attivita") + 1
            col_mese = header.index("Mese_Ultima_Attivita") + 1
            max_col = ws.max_column
            cutoff = date.today() - relativedelta(months=2)
            cutoff_period = cutoff.year * 100 + cutoff.month

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
        except:
            pass

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
    console.error("Errore generazione report completo:", e);
    hideOverlay();
    showError("Errore generazione report: " + (e?.message || e));
  }
}
