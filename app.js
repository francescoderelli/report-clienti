/* ============================================
   app.js - Report Clienti (app1.0)
   - Pyodide + pandas + openpyxl + python-dateutil
   - Verifica automatica su singolo file appena caricato
   - Modal "File errato" / "File OK" (senza alert browser)
   - Genera report Excel (fix ID float: 3007.0 -> 3007)
   - Output non corrotto (Uint8Array)
============================================ */

let pyodide = null;

const fileTab = document.getElementById("fileTabella");
const fileSum = document.getElementById("fileSum");
const btnRun  = document.getElementById("btnRun");
const logEl   = document.getElementById("log");

// Modal errore (già nel tuo index)
const errModal = document.getElementById("errModal");
const errOk    = document.getElementById("errOk");

// Modal OK (DEVI averlo in index: vedi nota in fondo)
const okModal  = document.getElementById("okModal");
const okOk     = document.getElementById("okOk");

// Stato file
const state = {
  tabOk: false,
  sumOk: false,
  tabBytes: null,
  sumBytes: null,
};

function log(msg){
  if (!logEl) return;
  logEl.textContent += msg + "\n";
  logEl.scrollTop = logEl.scrollHeight;
}
function clearLog(){
  if (!logEl) return;
  logEl.textContent = "";
}

function showErrModal(){
  if (!errModal) { alert("File errato"); return; }
  errModal.classList.remove("hidden");
}
function hideErrModal(){
  if (!errModal) return;
  errModal.classList.add("hidden");
}
function showOkModal(){
  if (!okModal) return;
  okModal.classList.remove("hidden");
}
function hideOkModal(){
  if (!okModal) return;
  okModal.classList.add("hidden");
}

async function readAsUint8Array(file){
  const buf = await file.arrayBuffer();
  return new Uint8Array(buf);
}

function updateRunEnabled(){
  btnRun.disabled = !(state.tabOk && state.sumOk);
}

async function init(){
  clearLog();
  btnRun.disabled = true;

  // Modal buttons
  if (errOk) errOk.addEventListener("click", hideErrModal);
  if (okOk)  okOk.addEventListener("click", hideOkModal);

  // Hook file change
  fileTab.addEventListener("change", () => onFileSelected("tab"));
  fileSum.addEventListener("change", () => onFileSelected("sum"));

  btnRun.addEventListener("click", runReport);

  log("Carico Pyodide...");
  pyodide = await loadPyodide({
    indexURL: "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/"
  });
  log("Pyodide pronto.");

  // pandas da pacchetto pyodide
  log("Carico pandas...");
  await pyodide.loadPackage(["pandas", "micropip"]);
  log("Installo openpyxl e python-dateutil...");
  await pyodide.runPythonAsync(`
import micropip
await micropip.install(["openpyxl","python-dateutil"])
`);
  log("Pacchetti OK. Seleziona i file.");
}
init().catch(e => {
  console.error(e);
  log("Errore inizializzazione.");
});

async function onFileSelected(kind){
  // kind: "tab" o "sum"
  if (!pyodide) return;

  // reset stato di quel file
  if (kind === "tab") {
    state.tabOk = false;
    state.tabBytes = null;
  } else {
    state.sumOk = false;
    state.sumBytes = null;
  }
  updateRunEnabled();

  const input = (kind === "tab") ? fileTab : fileSum;
  if (!input.files || input.files.length !== 1) return;

  clearLog();
  log("Verifica file...");

  const bytes = await readAsUint8Array(input.files[0]);

  // Verifica contenuto: colonne minime + score su header
  // - Tabella: deve avere ID_SOGGETTO e Tipo e Cliente + colonne economiche
  // - Sum_of : deve avere Anno/Mese + CodiceSoggetto + NomeSoggetto + Classe Attività
  pyodide.globals.set("FILE_BYTES", bytes);

  const res = await pyodide.runPythonAsync(`
import io
import pandas as pd

df = pd.read_excel(io.BytesIO(bytes(FILE_BYTES)), sheet_name=0)

cols = [str(c).strip().upper() for c in df.columns]

def score_tabella(cols):
    s = 0
    must = ["ID_SOGGETTO","TIPO","CLIENTE"]
    for m in must:
        if any(m == c for c in cols): s += 2
    # bonus colonne tipiche
    bonus = ["RESPONSABILE","CONDOMINI IN ALBERT","CONDOMINI AMMINISTRATI","PREVENTIVATO","DELIBERATO","FATTURATO","INCASSATO"]
    for b in bonus:
        if any(b in c for c in cols): s += 1
    return s

def score_sumof(cols):
    s = 0
    must = ["ANNO","MESE","CODICESOGGETTO","NOMESOGGETTO","CLASSE ATTIVITÀ","CLASSE ATTIVITA"]
    # anno/mese
    if any("ANNO" == c for c in cols): s += 2
    if any("MESE" == c for c in cols): s += 2
    # codice/nome
    if any("CODICESOGGETTO" == c or "CODICE SOGGETTO" == c for c in cols): s += 2
    if any("NOMESOGGETTO" == c or "NOME SOGGETTO" == c for c in cols): s += 2
    # classe attività
    if any("CLASSE ATTIV" in c for c in cols): s += 2
    return s

tab_s = score_tabella(cols)
sum_s = score_sumof(cols)

is_tab = tab_s >= 5 and sum_s < 6
is_sum = sum_s >= 6 and tab_s < 6

(df.shape[1], tab_s, sum_s, is_tab, is_sum)
`);

  const [nCols, tabScore, sumScore, isTab, isSum] = res.toJs();

  // Assegna stato in base al kind atteso e al contenuto
  let ok = false;
  if (kind === "tab") ok = !!isTab;
  if (kind === "sum") ok = !!isSum;

  if (!ok){
    // file errato: modal e reset input
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
    showErrModal();
    return;
  }

  // ok: salva bytes e flag
  if (kind === "tab") {
    state.tabOk = true;
    state.tabBytes = bytes;
  } else {
    state.sumOk = true;
    state.sumBytes = bytes;
  }

  updateRunEnabled();

  // Se entrambi OK -> modal ok
  if (state.tabOk && state.sumOk){
    showOkModal();
  }
}

async function runReport(){
  if (!(state.tabOk && state.sumOk)) return;
  clearLog();
  log("Genero report...");

  pyodide.globals.set("TAB_BYTES", state.tabBytes);
  pyodide.globals.set("SUM_BYTES", state.sumBytes);

  // PY_REPORT: logica report (robusta con norm_id)
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

# --- leggi bytes
tab = pd.read_excel(io.BytesIO(bytes(TAB_BYTES)))
su  = pd.read_excel(io.BytesIO(bytes(SUM_BYTES)))

# --- Tabella Clienti: usa NOMI se presenti, altrimenti fallback su lettere standard
tab_cols = {str(c).strip().upper(): c for c in tab.columns}

def col_any(tab, keys, fallback_idx=None):
    for k in keys:
        ku = k.upper()
        if ku in tab_cols:
            return tab_cols[ku]
    # contiene
    for c in tab.columns:
        cu = str(c).upper()
        for k in keys:
            if k.upper() in cu:
                return c
    # fallback su indice
    if fallback_idx is not None and tab.shape[1] > fallback_idx:
        return tab.columns[fallback_idx]
    return None

c_id   = col_any(tab, ["ID_SOGGETTO"], fallback_idx=8)     # spesso è I nei vecchi export
c_tipo = col_any(tab, ["TIPO"], fallback_idx=15)
c_cli  = col_any(tab, ["CLIENTE"], fallback_idx=9)
c_ref  = col_any(tab, ["RESPONSABILE", "REFERENTE"], fallback_idx=7)

c_ca   = col_any(tab, ["CONDOMINI IN ALBERT"], fallback_idx=20)
c_cam  = col_any(tab, ["CONDOMINI AMMINISTRATI"], fallback_idx=21)
c_prev = col_any(tab, ["PREVENTIVATO"], fallback_idx=22)
c_del  = col_any(tab, ["DELIBERATO"], fallback_idx=23)
c_fat  = col_any(tab, ["FATTURATO"], fallback_idx=24)
c_inc  = col_any(tab, ["INCASSATO"], fallback_idx=25)

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

# --- Sum_of: nomi colonne
su_cols = {str(c).strip().upper(): c for c in su.columns}

def su_col_any(keys, fallback_idx=None):
    for k in keys:
        ku = k.upper()
        if ku in su_cols:
            return su_cols[ku]
    for c in su.columns:
        cu = str(c).upper()
        for k in keys:
            if k.upper() in cu:
                return c
    if fallback_idx is not None and su.shape[1] > fallback_idx:
        return su.columns[fallback_idx]
    return None

s_anno = su_col_any(["ANNO"], fallback_idx=0)
s_mese = su_col_any(["MESE"], fallback_idx=1)
s_att  = su_col_any(["CLASSE ATTIVITÀ","CLASSE ATTIVITA","ATTIVITA","ATTIVITÀ"], fallback_idx=2)
s_chi  = su_col_any(["RESPONSABILE","CHI"], fallback_idx=4)
s_cod  = su_col_any(["CODICESOGGETTO","CODICE SOGGETTO"], fallback_idx=6)
s_nome = su_col_any(["NOMESOGGETTO","NOME SOGGETTO"], fallback_idx=7)

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

# --- Scrivi Excel in memoria
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

    # Formato € su I-L e header
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

    # Foglio amministratore: match robusto + colori + attivo
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

    # Auto-width
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

    // Prendi bytes e scarica (no file corrotto)
    const outBytes = pyodide.globals.get("OUT_BYTES");
    const u8 = new Uint8Array(outBytes.toJs()); // fondamentale
    const blob = new Blob([u8], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    saveAs(blob, "Report_Tipo_Clienti.xlsx");
    log("Report creato: Report_Tipo_Clienti.xlsx");
  } catch (e) {
    console.error(e);
    log("Errore durante la generazione.");
    showErrModal();
  }
}
