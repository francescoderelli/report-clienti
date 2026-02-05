// app.js (v1.2) - GitHub Pages + Pyodide

let pyodide = null;

const logEl = document.getElementById("log");
const fileTab = document.getElementById("fileTabella");
const fileSum = document.getElementById("fileSum");
const btnVerify = document.getElementById("btnVerify");
const btnRun = document.getElementById("btnRun");

function log(msg) {
  logEl.textContent += msg + "\n";
  logEl.scrollTop = logEl.scrollHeight;
}
function clearLog() { logEl.textContent = ""; }

function bothSelected() {
  return fileTab.files.length === 1 && fileSum.files.length === 1;
}

async function readAsUint8Array(file) {
  const buf = await file.arrayBuffer();
  return new Uint8Array(buf);
}

// -----------------------
// INIT
// -----------------------
async function init() {
  clearLog();
  btnVerify.disabled = true;
  btnRun.disabled = true;

  try {
    log("Carico Pyodide...");
    pyodide = await loadPyodide();
    log("Pyodide pronto.");

    log("Scarico pandas...");
    await pyodide.loadPackage(["pandas"]);
    log("pandas OK.");

    log("Carico micropip...");
    await pyodide.loadPackage("micropip"); // <-- FIX fondamentale
    log("micropip OK.");

    log("Installo openpyxl e python-dateutil (può richiedere un po')...");
    await pyodide.runPythonAsync(`
import micropip
await micropip.install(["openpyxl","python-dateutil"])
`);
    log("Pacchetti OK.");

    // abilita bottoni quando entrambi i file sono scelti
    const onChange = () => {
      btnVerify.disabled = !bothSelected();
      btnRun.disabled = true;
    };
    fileTab.addEventListener("change", onChange);
    fileSum.addEventListener("change", onChange);

    btnVerify.addEventListener("click", verifyFiles);
    btnRun.addEventListener("click", runReport);

    log("Seleziona i 2 file e clicca 'Verifica file'.");
  } catch (e) {
    log("ERRORE init:");
    log(String(e));
    console.error(e);
  }
}

init();

// -----------------------
// VERIFY
// -----------------------
async function verifyFiles() {
  clearLog();
  btnRun.disabled = true;

  try {
    log("Verifica file...");

    const tabBytes = await readAsUint8Array(fileTab.files[0]);
    const sumBytes = await readAsUint8Array(fileSum.files[0]);

    pyodide.globals.set("TAB_BYTES", tabBytes);
    pyodide.globals.set("SUM_BYTES", sumBytes);

    const res = await pyodide.runPythonAsync(`
import io
import pandas as pd

def col_count(xlsx_bytes):
    df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=0)
    return df.shape[1]

tab_cols = col_count(bytes(TAB_BYTES))
sum_cols = col_count(bytes(SUM_BYTES))

ok_tab = tab_cols >= 26   # fino a Z
ok_sum = sum_cols >= 8    # fino a H

(tab_cols, sum_cols, ok_tab, ok_sum)
`);
    const [tabCols, sumCols, okTab, okSum] = res.toJs();

    log(\`Tabella Clienti: colonne = \${tabCols} (serve >= 26 fino a Z) -> \${okTab ? "OK" : "NON OK"}\`);
    log(\`Sum_of: colonne = \${sumCols} (serve >= 8 fino a H) -> \${okSum ? "OK" : "NON OK"}\`);

    if (okTab && okSum) {
      log("Verifica superata. Puoi generare l’output.");
      btnRun.disabled = false;
    } else {
      log("Verifica fallita: carica i file corretti.");
      btnRun.disabled = true;
    }
  } catch (e) {
    log("ERRORE verifica:");
    log(String(e));
    console.error(e);
  }
}

// -----------------------
// RUN REPORT (v1.2)
// -----------------------
async function runReport() {
  clearLog();
  btnRun.disabled = true;

  try {
    log("Genero output (v1.2)...");

    const tabBytes = await readAsUint8Array(fileTab.files[0]);
    const sumBytes = await readAsUint8Array(fileSum.files[0]);

    pyodide.globals.set("TAB_BYTES", tabBytes);
    pyodide.globals.set("SUM_BYTES", sumBytes);

    await pyodide.runPythonAsync(`
import io, re
import numpy as np
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill

def excel_col_letter_to_index(letter: str) -> int:
    letter = letter.upper().strip()
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1

def sanitize_sheet_name(name: str) -> str:
    name = "Senza_Tipo" if name is None or str(name).strip() == "" or str(name).lower() == "nan" else str(name).strip()
    name = re.sub(r'[:\\\\/\\?\\*\\[\\]]', '-', name)
    return name[:31]

def month_to_int(x):
    if pd.isna(x): return np.nan
    s = str(x).strip()
    try:
        v = int(float(s))
        if 1 <= v <= 12: return v
    except:
        pass
    m = re.search(r"\\b(\\d{1,2})\\b", s)
    if m:
        v = int(m.group(1))
        if 1 <= v <= 12: return v
    mesi = {"gen":1,"gennaio":1,"feb":2,"febbraio":2,"mar":3,"marzo":3,"apr":4,"aprile":4,"mag":5,"maggio":5,
            "giu":6,"giugno":6,"lug":7,"luglio":7,"ago":8,"agosto":8,"set":9,"sett":9,"settembre":9,
            "ott":10,"ottobre":10,"nov":11,"novembre":11,"dic":12,"dicembre":12}
    low = s.lower()
    for k,v in mesi.items():
        if k in low: return v
    return np.nan

priority_map = {"07":7,"06":6,"04":5,"03":4,"05":3,"01":2,"02":1}
def activity_priority(a):
    if pd.isna(a): return 0
    s = str(a).strip()
    m = re.match(r"^\\s*(\\d{2})", s)
    if m: return priority_map.get(m.group(1), 0)
    m2 = re.search(r"\\b(0[1-7])\\b", s)
    if m2: return priority_map.get(m2.group(1), 0)
    return 0

# --- leggi bytes
tab = pd.read_excel(io.BytesIO(bytes(TAB_BYTES)))
su  = pd.read_excel(io.BytesIO(bytes(SUM_BYTES)))

# --- Tabella Clienti: H I J P U V W X Y Z
idx_H = excel_col_letter_to_index("H")
idx_I = excel_col_letter_to_index("I")
idx_J = excel_col_letter_to_index("J")
idx_P = excel_col_letter_to_index("P")
idx_U = excel_col_letter_to_index("U")
idx_V = excel_col_letter_to_index("V")
idx_W = excel_col_letter_to_index("W")
idx_X = excel_col_letter_to_index("X")
idx_Y = excel_col_letter_to_index("Y")
idx_Z = excel_col_letter_to_index("Z")

clients = pd.DataFrame({
    "ID_Soggetto": tab.iloc[:, idx_I].astype(str).str.strip(),
    "Tipo": tab.iloc[:, idx_P],
    "Cliente_Tabella": tab.iloc[:, idx_J],
    "Referente_Commerciale": tab.iloc[:, idx_H],
    "Condomini_in_Albert": tab.iloc[:, idx_U],
    "Condomini_Amministrati": tab.iloc[:, idx_V],
    "PREVENTIVATO_EUR": tab.iloc[:, idx_W],
    "DELIBERATO_EUR": tab.iloc[:, idx_X],
    "FATTURATO_EUR": tab.iloc[:, idx_Y],
    "INCASSATO_EUR": tab.iloc[:, idx_Z],
})

# --- Sum_of: A B C E G H
idx_A = excel_col_letter_to_index("A")
idx_B = excel_col_letter_to_index("B")
idx_C = excel_col_letter_to_index("C")
idx_E = excel_col_letter_to_index("E")
idx_G = excel_col_letter_to_index("G")
idx_Hs = excel_col_letter_to_index("H")

sumdf = pd.DataFrame({
    "Anno": su.iloc[:, idx_A],
    "Mese": su.iloc[:, idx_B],
    "Attivita": su.iloc[:, idx_C],
    "Chi": su.iloc[:, idx_E],
    "ID_Soggetto": su.iloc[:, idx_G].astype(str).str.strip(),
    "Nome_Soggetto_Sum": su.iloc[:, idx_Hs],
})

sumdf["Anno"] = pd.to_numeric(sumdf["Anno"], errors="coerce").astype("Int64")
sumdf["Mese_num"] = sumdf["Mese"].apply(month_to_int).astype("Int64")
sumdf["Prio"] = sumdf["Attivita"].apply(activity_priority).astype(int)
sumdf["Periodo"] = (sumdf["Anno"] * 100 + sumdf["Mese_num"]).astype("Int64")
sumdf = sumdf.dropna(subset=["ID_Soggetto", "Periodo"])
sumdf["_row"] = np.arange(len(sumdf))

best_in_month = (sumdf.sort_values(["ID_Soggetto","Periodo","Prio","_row"])
                      .groupby(["ID_Soggetto","Periodo"], as_index=False)
                      .tail(1))
best_last = (best_in_month.sort_values(["ID_Soggetto","Periodo","Prio","_row"])
                        .groupby("ID_Soggetto", as_index=False)
                        .tail(1))

last_act = best_last[["ID]()]()
