// app.js - app1.0 + verifica automatica immediata per contenuto (no nome) + modal "File errato"

let pyodide = null;

const logEl = document.getElementById("log");
const fileTab = document.getElementById("fileTabella");
const fileSum = document.getElementById("fileSum");
const btnRun = document.getElementById("btnRun");

// Modal errore custom
const errModal = document.getElementById("errModal");
const errOk = document.getElementById("errOk");

function log(msg){
  logEl.textContent += msg + "\n";
  logEl.scrollTop = logEl.scrollHeight;
}
function clearLog(){ logEl.textContent = ""; }

function bothSelected(){
  return fileTab.files.length === 1 && fileSum.files.length === 1;
}

function showFileErrato(){
  errModal.classList.remove("hidden");
  errOk.focus();
}
errOk.addEventListener("click", () => {
  errModal.classList.add("hidden");
});

// lettura file
async function readAsUint8Array(file){
  const buf = await file.arrayBuffer();
  return new Uint8Array(buf);
}

// -----------------------
// PY: analisi contenuto (NON nome)
// ritorna: kind ("tabella"|"sum_of"|"unknown"), ncols, score_tab, score_sum
// -----------------------
const PY_ANALYZE = String.raw`
import io, re
import pandas as pd
import numpy as np

def excel_col_letter_to_index(letter: str) -> int:
    letter = letter.upper().strip()
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1

df = pd.read_excel(io.BytesIO(bytes(ONE_FILE_BYTES)), sheet_name=0)
ncols = int(df.shape[1])

def safe_col(i):
    if i < 0 or i >= df.shape[1]:
        return pd.Series([], dtype="object")
    return df.iloc[:, i]

# segnali sum_of
colA = safe_col(0)
colB = safe_col(1)
colC = safe_col(2).astype(str).str.lower()

has_activity_codes = colC.str.contains(r"\\b0[1-7]\\b").any()

try:
    a_num = pd.to_numeric(colA, errors="coerce")
    has_year = ((a_num >= 2000) & (a_num <= 2100)).any()
except:
    has_year = False

b_str = colB.astype(str).str.lower()
has_month = (
    b_str.str.contains(r"\\b(1[0-2]|[1-9])\\b").any()
    or b_str.str.contains(r"gen|feb|mar|apr|mag|giu|lug|ago|set|ott|nov|dic").any()
)

# segnali tabella
idx_P = excel_col_letter_to_index("P")
colP = safe_col(idx_P).astype(str).str.lower()
has_tipo_like = colP.str.contains(r"amministr|condomin|ente|privat|azienda|tipo|cliente").any()

idx_W = excel_col_letter_to_index("W")
idx_Z = excel_col_letter_to_index("Z")
has_money_like = False
if ncols > idx_Z:
    block = df.iloc[:, idx_W:idx_Z+1]
    try:
        block_num = block.apply(pd.to_numeric, errors="coerce")
        has_money_like = (block_num.fillna(0) > 0).any().any()
    except:
        has_money_like = False

score_sum = 0
score_tab = 0

if has_year: score_sum += 2
if has_month: score_sum += 1
if has_activity_codes: score_sum += 2
if ncols >= 8: score_sum += 1

if ncols >= 26: score_tab += 2
if has_tipo_like: score_tab += 2
if has_money_like: score_tab += 2

if score_tab > score_sum:
    kind = "tabella"
elif score_sum > score_tab:
    kind = "sum_of"
else:
    kind = "unknown"

(kind, ncols, score_tab, score_sum)
`;

// -----------------------
// PY: report completo -> OUT_BYTES
// -----------------------
const PY_REPORT = String.raw`
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

tab = pd.read_excel(io.BytesIO(bytes(TAB_BYTES)))
su  = pd.read_excel(io.BytesIO(bytes(SUM_BYTES)))

# Tabella: H I J P U V W X Y Z
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

# Sum_of: A B C E G H
idx_A  = excel_col_letter_to_index("A")
idx_B  = excel_col_letter_to_index("B")
idx_C  = excel_col_letter_to_index("C")
idx_E  = excel_col_letter_to_index("E")
idx_G  = excel_col_letter_to_index("G")
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

last_act = best_last[["ID_Soggetto","Anno","Mese_num","Attivita","Chi"]].copy()
last_act.rename(columns={
    "Anno":"Anno_Ultima_Attivita",
    "Mese_num":"Mese_Ultima_Attivita",
    "Attivita":"Ultima_Attivita",
    "Chi":"Ultima_Attivita_Fatta_Da"
}, inplace=True)

name_map = (sumdf[["ID_Soggetto","Nome_Soggetto_Sum"]]
            .dropna(subset=["Nome_Soggetto_Sum"])
            .drop_duplicates(subset=["ID_Soggetto"], keep="last"))

corrispondenza = (clients[["ID_Soggetto","Cliente_Tabella"]]
                  .merge(name_map, on="ID_Soggetto", how="left")
                  .sort_values("ID_Soggetto"))

final = clients.merge(last_act, on="ID_Soggetto", how="left").merge(name_map, on="ID_Soggetto", how="left")
final["Cliente"] = final["Nome_Soggetto_Sum"].fillna(final["Cliente_Tabella"]).fillna(final["ID_Soggetto"])

output_cols = [
    "Cliente","Referente_Commerciale","Condomini_in_Albert","Condomini_Amministrati",
    "Anno_Ultima_Attivita","Mese_Ultima_Attivita","Ultima_Attivita","Ultima_Attivita_Fatta_Da",
    "PREVENTIVATO_EUR","DELIBERATO_EUR","FATTURATO_EUR","INCASSATO_EUR"
]

EURO = chr(8364)
header_overrides = {
    "PREVENTIVATO_EUR": "Preventivato " + EURO,
    "DELIBERATO_EUR":   "Deliberato "   + EURO,
    "FATTURATO_EUR":    "Fatturato "    + EURO,
    "INCASSATO_EUR":    "Incassato "    + EURO,
}

out = io.BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    riepilogo = (final.assign(Tipo=final["Tipo"].fillna("Senza_Tipo"))
                      .groupby("Tipo", dropna=False).size()
                      .reset_index(name="N_clienti")
                      .sort_values("N_clienti", ascending=False))
    riepilogo.to_excel(writer, sheet_name="Riepilogo", index=False)
    corrispondenza.to_excel(writer, sheet_name="Corrispondenza", index=False)

    used = {"Riepilogo","Corrispondenza"}
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

    euro_format = EURO + ' #,##0.00'
    euro_cols = [9,10,11,12]
    type_sheets = [s for s in wb.sheetnames if s not in ("Riepilogo","Corrispondenza")]
    for sname in type_sheets:
        ws = wb[sname]
        for col_idx in euro_cols:
            cur = ws.cell(row=1, column=col_idx).value
            if cur in header_overrides:
                ws.cell(row=1, column=col_idx).value = header_overrides[cur]
        for r in range(2, ws.max_row+1):
            for c in euro_cols:
                ws.cell(row=r, column=c).number_format = euro_format

    GREEN = PatternFill(fill_type="solid", fgColor="C6EFCE")
    RED   = PatternFill(fill_type="solid", fgColor="FFC7CE")
    cutoff = date.today() - relativedelta(months=2)
    cutoff_period = cutoff.year*100 + cutoff.month

    admin_sheet = None
    for s in wb.sheetnames:
        if s not in ("Riepilogo","Corrispondenza") and "amministr" in s.lower():
            admin_sheet = s
            break

    if admin_sheet:
        ws = wb[admin_sheet]
        header = [c.value for c in ws[1]]
        col_anno = header.index("Anno_Ultima_Attivita")+1
        col_mese = header.index("Mese_Ultima_Attivita")+1
        max_col = ws.max_column
        for r in range(2, ws.max_row+1):
            anno = ws.cell(r, col_anno).value
            mese = ws.cell(r, col_mese).value
            if anno is None or mese is None or str(anno).strip()=="" or str(mese).strip()=="":
                fill = RED
            else:
                try:
                    period = int(anno)*100 + int(mese)
                    fill = RED if period < cutoff_period else GREEN
                except:
                    fill = RED
            for c in range(1, max_col+1):
                ws.cell(r,c).fill = fill
        wb.active = wb.sheetnames.index(admin_sheet)

    for ws in wb.worksheets:
        for col_idx, col_cells in enumerate(ws.columns, start=1):
            max_len = 0
            for cell in list(col_cells)[:2000]:
                if cell.value is None:
                    continue
                max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len+2, 45)

out.seek(0)
OUT_BYTES = out.read()
`;

// -----------------------
// INIT
// -----------------------
async function init(){
  clearLog();
  btnRun.disabled = true;

  try{
    log("Carico Pyodide...");
    pyodide = await loadPyodide({
      indexURL: "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/"
    });
    log("Pyodide pronto.");

    log("Scarico pandas...");
    await pyodide.loadPackage(["pandas"]);
    log("pandas OK.");

    log("Carico micropip...");
    await pyodide.loadPackage("micropip");
    log("micropip OK.");

    log("Installo openpyxl e python-dateutil...");
    await pyodide.runPythonAsync(`
import micropip
await micropip.install(["openpyxl","python-dateutil"])
`);
    clearLog();

    fileTab.addEventListener("change", () => onFileChanged("tabella"));
    fileSum.addEventListener("change", () => onFileChanged("sum_of"));

    btnRun.addEventListener("click", runReport);

  }catch(e){
    clearLog();
    log("ERRORE init:");
    log(String(e));
    console.error(e);
  }
}
init();

// -----------------------
// CONTENT-BASED CHECK
// -----------------------
async function analyzeFile(file){
  const bytes = await readAsUint8Array(file);
  pyodide.globals.set("ONE_FILE_BYTES", bytes);
  const res = await pyodide.runPythonAsync(PY_ANALYZE);
  const [kind, ncols] = res.toJs();
  return { kind, ncols };
}

async function onFileChanged(expectedKind){
  btnRun.disabled = true;
  clearLog();

  const file = (expectedKind === "tabella") ? fileTab.files[0] : fileSum.files[0];
  if(!file) return;

  try{
    const info = await analyzeFile(file);

    const minColsOk = (expectedKind === "tabella") ? (info.ncols >= 26) : (info.ncols >= 8);
    const kindOk = (info.kind === expectedKind); // match forte, niente unknown

    if(!minColsOk || !kindOk){
      showFileErrato();
      return;
    }

    if(bothSelected()){
      const tabInfo = await analyzeFile(fileTab.files[0]);
      const sumInfo = await analyzeFile(fileSum.files[0]);

      const okTab = (tabInfo.kind === "tabella" && tabInfo.ncols >= 26);
      const okSum = (sumInfo.kind === "sum_of" && sumInfo.ncols >= 8);

      if(!okTab || !okSum){
        showFileErrato();
        btnRun.disabled = true;
        return;
      }

      btnRun.disabled = false;
    }
  }catch(e){
    console.error(e);
    showFileErrato();
  }
}

// -----------------------
// RUN REPORT
// -----------------------
async function runReport(){
  clearLog();
  btnRun.disabled = true;

  try{
    const tabBytes = await readAsUint8Array(fileTab.files[0]);
    const sumBytes = await readAsUint8Array(fileSum.files[0]);

    pyodide.globals.set("TAB_BYTES", tabBytes);
    pyodide.globals.set("SUM_BYTES", sumBytes);

    await pyodide.runPythonAsync(PY_REPORT);

    const outProxy = pyodide.globals.get("OUT_BYTES");
    const outBytes = outProxy.toJs({ create_proxies: false });
    outProxy.destroy();

    const blob = new Blob([outBytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    saveAs(blob, "Report_Tipo_Clienti.xlsx");

  }catch(e){
    console.error(e);
    alert("Errore generazione file");
  }finally{
    btnRun.disabled = false;
  }
}
