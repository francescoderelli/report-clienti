/* ============================================
   app.js - Report Clienti (app1.0 stabile)
   MODIFICHE:
   - Niente log visibile (logEl rimosso)
   - Niente pillStatus (rimosso)
   - Overlay pulito durante init
   - stdout/stderr Pyodide silenziati
   - Modal "File OK" anche quando carichi un singolo file
   - Run abilitato solo quando entrambi i file sono OK
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

function showOverlay(){ if (overlay) overlay.classList.remove("hidden"); }
function hideOverlay(){ if (overlay) overlay.classList.add("hidden"); }

function showModal(modal){
  if (!modal) return;
  modal.classList.remove("hidden");
}
function hideModal(modal){
  if (!modal) return;
  modal.classList.add("hidden");
}

function updateRunEnabled(){
  btnRun.disabled = !(state.tabOk && state.sumOk);
}

async function readAsUint8Array(file){
  const buf = await file.arrayBuffer();
  return new Uint8Array(buf);
}

async function init(){
  btnRun.disabled = true;
  showOverlay();

  if (errOk) errOk.addEventListener("click", () => hideModal(errModal));
  if (okOk)  okOk.addEventListener("click",  () => hideModal(okModal));

  fileTab.addEventListener("change", () => onFileSelected("tab"));
  fileSum.addEventListener("change", () => onFileSelected("sum"));
  btnRun.addEventListener("click", runReport);

  pyodide = await loadPyodide({
    indexURL: "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/"
  });

  // ✅ silenzia completamente stdout/stderr di pyodide (niente “tool loading”)
  try {
    pyodide.setStdout({ batched: (s) => {} });
    pyodide.setStderr({ batched: (s) => {} });
  } catch (_) {
    // fallback: se setStdout non esiste, lo faremo in Python con redirect
  }

  await pyodide.loadPackage(["pandas", "micropip"]);

  // ✅ install silenziato anche lato Python (ulteriore sicurezza)
  await pyodide.runPythonAsync(`
import sys, io
sys.stdout = io.StringIO()
sys.stderr = io.StringIO()

import micropip
await micropip.install(["openpyxl","python-dateutil"])
`);

  hideOverlay();
}
init().catch(e => {
  console.error(e);
  hideOverlay();
  showModal(errModal);
});

async function onFileSelected(kind){
  if (!pyodide) return;

  // reset del file specifico
  if (kind === "tab") { state.tabOk = false; state.tabBytes = null; }
  else { state.sumOk = false; state.sumBytes = null; }
  updateRunEnabled();

  const input = (kind === "tab") ? fileTab : fileSum;
  if (!input.files || input.files.length !== 1) return;

  const bytes = await readAsUint8Array(input.files[0]);
  pyodide.globals.set("FILE_BYTES", bytes);

  const res = await pyodide.runPythonAsync(`
import io
import pandas as pd

def norm_cols(cols):
    return [str(c).strip().upper() for c in cols]

def score(cols):
    tab_must = {"ID_SOGGETTO","TIPO","CLIENTE"}
    tab_bonus = ["RESPONSABILE","RESPONSABILEAREA","CONDOMINI IN ALBERT","CONDOMINI AMMINISTRATI",
                 "PREVENTIVATO","DELIBERATO","FATTURATO","INCASSATO"]
    sum_must = {"ANNO","MESE","CODICESOGGETTO","NOMESOGGETTO"}
    sum_bonus = ["CLASSE ATTIV"]

    cset = set(cols)
    tab_s = 0
    sum_s = 0

    for m in tab_must:
        if m in cset: tab_s += 3
    for b in tab_bonus:
        if any(b in c for c in cols): tab_s += 1

    for m in sum_must:
        if m in cset: sum_s += 3
    for b in sum_bonus:
        if any(b in c for c in cols): sum_s += 1

    return tab_s, sum_s

def classify(tab_s, sum_s):
    if tab_s >= 6 and tab_s > sum_s + 1: return "tabella"
    if sum_s >= 6 and sum_s > tab_s + 1: return "sumof"
    return "unknown"

xlsx = bytes(FILE_BYTES)

best_kind = "unknown"
best_score = -1
best_ncols = 0

# try 1: header normale
try:
    df1 = pd.read_excel(io.BytesIO(xlsx), sheet_name=0)
    cols1 = norm_cols(df1.columns)
    t1, s1 = score(cols1)
    sc1 = t1 + s1
    k1 = classify(t1, s1)
    if sc1 > best_score:
        best_score = sc1
        best_kind = k1
        best_ncols = df1.shape[1]
except Exception:
    pass

# try 2: header cercato
df0 = pd.read_excel(io.BytesIO(xlsx), sheet_name=0, header=None)

header_row = None
for i in range(min(80, len(df0))):
    row = df0.iloc[i].astype(str).str.upper().str.strip().tolist()
    if "ID_SOGGETTO" in row and ("TIPO" in row or "CLIENTE" in row):
        header_row = i
        break
    if "ANNO" in row and "MESE" in row and "CODICESOGGETTO" in row:
        header_row = i
        break

if header_row is not None:
    df2 = df0.copy()
    df2.columns = df2.iloc[header_row].astype(str)
    df2 = df2.iloc[header_row+1:].reset_index(drop=True)
    cols2 = norm_cols(df2.columns)
    t2, s2 = score(cols2)
    sc2 = t2 + s2
    k2 = classify(t2, s2)
    if sc2 > best_score:
        best_score = sc2
        best_kind = k2
        best_ncols = df2.shape[1]

(best_kind, best_ncols)
`);

  const [kindFound] = res.toJs();

  let ok = false;
  if (kind === "tab") ok = (kindFound === "tabella");
  if (kind === "sum") ok = (kindFound === "sumof");

  if (!ok){
    if (kind === "tab") { fileTab.value = ""; state.tabOk = false; state.tabBytes = null; }
    else { fileSum.value = ""; state.sumOk = false; state.sumBytes = null; }
    updateRunEnabled();
    showModal(errModal);
    return;
  }

  // salva bytes + stato
  if (kind === "tab") { state.tabOk = true; state.tabBytes = bytes; }
  else { state.sumOk = true; state.sumBytes = bytes; }

  updateRunEnabled();

  // ✅ Modal OK anche per singolo file
  if (okText){
    okText.textContent = (kind === "tab")
      ? "Tabella Clienti verificata."
      : "Sum_of verificato.";
  }
  showModal(okModal);
}

async function runReport(){
  if (!(state.tabOk && state.sumOk)) return;

  pyodide.globals.set("TAB_BYTES", state.tabBytes);
  pyodide.globals.set("SUM_BYTES", state.sumBytes);

  const PY_REPORT = String.raw`
import io, re
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
    s = s.replace("\u00A0","")
    s = re.sub(r"\s+","", s)
    s = re.sub(r"\.0$","", s)
    return s

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
    except: pass
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

# Tabella Clienti
c_id   = pick_col(tab, ["ID_SOGGETTO"], fallback_idx=8)
c_tipo = pick_col(tab, ["TIPO"], fallback_idx=15)
c_cli  = pick_col(tab, ["CLIENTE"], fallback_idx=9)
c_ref  = pick_col(tab, ["RESPONSABILE","REFERENTE"], fallback_idx=7)

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

# Sum_of
s_anno = pick_col(su, ["ANNO"], fallback_idx=0)
s_mese = pick_col(su, ["MESE"], fallback_idx=1)
s_att  = pick_col(su, ["CLASSE ATTIVITÀ","CLASSE ATTIVITA","ATTIVITA","ATTIVITÀ"], fallback_idx=2)
s_chi  = pick_col(su, ["RESPONSABILE","CHI"], fallback_idx=4)
s_cod  = pick_col(su, ["CODICESOGGETTO","CODICE SOGGETTO"], fallback_idx=6)
s_nome = pick_col(su, ["NOMESOGGETTO","NOME SOGGETTO"], fallback_idx=7)

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
sumdf = sumdf[(sumdf["ID_Soggetto"]!="")].dropna(subset=["Periodo"]).copy()
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

output_cols = ["Cliente","Referente_Commerciale","Condomini_in_Albert","Condomini_Amministrati",
               "Anno_Ultima_Attivita","Mese_Ultima_Attivita","Ultima_Attivita","Ultima_Attivita_Fatta_Da",
               "PREVENTIVATO_EUR","DELIBERATO_EUR","FATTURATO_EUR","INCASSATO_EUR"]

header_overrides = {
    "PREVENTIVATO_EUR":"Preventivato €",
    "DELIBERATO_EUR":"Deliberato €",
    "FATTURATO_EUR":"Fatturato €",
    "INCASSATO_EUR":"Incassato €",
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

    euro_format = u'€ #,##0.00'
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

  try {
    await pyodide.runPythonAsync(PY_REPORT);

    const outBytes = pyodide.globals.get("OUT_BYTES");
    const u8 = new Uint8Array(outBytes.toJs());
    const blob = new Blob([u8], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    saveAs(blob, "Report_Tipo_Clienti.xlsx");
  } catch (e) {
    console.error(e);
    showModal(errModal);
  }
}
