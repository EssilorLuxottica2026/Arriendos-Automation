"""
Diagnostico v3: enfocado en FE1173.
Reproduce el merge del PRORATEO que hace el pipeline y muestra EN QUE paso
se pierde el vw_percent para esa factura.
"""
from pathlib import Path
import sys
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
SRC_DIR = BASE_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from lease_accounting.pipeline.processor import LeaseAccountingPipeline

PRORATEO_PATH = BASE_DIR / "data" / "input" / "PRORATEO.xlsx"
pipe = LeaseAccountingPipeline(output_dir=BASE_DIR / "data" / "output")

# Llaves de FE1173 segun diagnostico v2 (DESPUES del fix del .0)
VENDOR_RAW = 208501.0
STORE_RAW = "T007"
CECO_RAW = 6563

print("=" * 80)
print("DIAGNOSTICO FE1173")
print("=" * 80)

# Fix nuevo deberia dar '208501'
vk = pipe._text_key(pipe._clean_numeric_code(VENDOR_RAW))
sk = pipe._text_key(STORE_RAW)
ck = pipe._ceco_key(CECO_RAW)
print(f"\nFactura FE1173: vendor_key={vk!r}, store_key={sk!r}, ceco_key={ck!r}")

# Cargar PRORATEO con el loader real
prorateo = pipe._load_prorateo(PRORATEO_PATH)
# Normalizar como lo hace el pipeline
if "vendor_code" in prorateo.columns:
    prorateo["vendor_code"] = prorateo["vendor_code"].map(pipe._clean_numeric_code)
if "vendor" not in prorateo.columns:
    prorateo["vendor"] = None
prorateo["vendor_key"] = prorateo.get("vendor_code", pd.Series(dtype=object)).fillna(prorateo["vendor"]).map(pipe._text_key)
prorateo["store_key"] = prorateo["store"].map(pipe._text_key)
prorateo["ceco_key"] = prorateo["ceco"].map(pipe._ceco_key)
prorateo["vw_percent"] = pd.to_numeric(prorateo["vw_percent"], errors="coerce")
prorateo["vw_percent"] = prorateo["vw_percent"].where(prorateo["vw_percent"] <= 1, prorateo["vw_percent"] / 100.0)

print(f"\nPRORATEO cargado: {len(prorateo)} filas, columnas={list(prorateo.columns)}")

# Buscar filas que matcheen cada key
print("\n[1] Filas del PRORATEO que matchean cada llave de FE1173\n")

match_v = prorateo[prorateo["vendor_key"] == vk]
print(f"-- Match por vendor_key='{vk}': {len(match_v)} filas")
for _, row in match_v.iterrows():
    print(f"   vendor={row.get('vendor')!r}  vendor_code={row.get('vendor_code')!r}  "
          f"store={row.get('store')!r}  ceco={row.get('ceco')!r}  "
          f"vw_percent={row.get('vw_percent')!r}")

match_s = prorateo[prorateo["store_key"] == sk]
print(f"\n-- Match por store_key='{sk}': {len(match_s)} filas")
for _, row in match_s.iterrows():
    print(f"   vendor={row.get('vendor')!r}  vendor_code={row.get('vendor_code')!r}  "
          f"store={row.get('store')!r}  ceco={row.get('ceco')!r}  "
          f"vw_percent={row.get('vw_percent')!r}")

match_c = prorateo[prorateo["ceco_key"] == ck]
print(f"\n-- Match por ceco_key='{ck}': {len(match_c)} filas")
for _, row in match_c.iterrows():
    print(f"   vendor={row.get('vendor')!r}  vendor_code={row.get('vendor_code')!r}  "
          f"store={row.get('store')!r}  ceco={row.get('ceco')!r}  "
          f"vw_percent={row.get('vw_percent')!r}")

# Simular el drop_duplicates del pipeline
print("\n[2] Que sobrevive al drop_duplicates (pipeline logic)\n")

prorateo_vendor = prorateo.drop_duplicates(subset=["vendor_key"])[["vendor_key", "vw_percent"]]
prorateo_store = prorateo.drop_duplicates(subset=["store_key"])[["store_key", "vw_percent"]]
prorateo_ceco = prorateo.drop_duplicates(subset=["ceco_key"])[["ceco_key", "vw_percent"]]

v_row = prorateo_vendor[prorateo_vendor["vendor_key"] == vk]
s_row = prorateo_store[prorateo_store["store_key"] == sk]
c_row = prorateo_ceco[prorateo_ceco["ceco_key"] == ck]

print(f"vw_percent_vendor (via '{vk}'): {v_row['vw_percent'].tolist() if not v_row.empty else 'NO MATCH'}")
print(f"vw_percent_store  (via '{sk}'): {s_row['vw_percent'].tolist() if not s_row.empty else 'NO MATCH'}")
print(f"vw_percent_ceco   (via '{ck}'): {c_row['vw_percent'].tolist() if not c_row.empty else 'NO MATCH'}")

print("\n[3] Resultado final (combine_first en cascada):")
candidates = []
if not v_row.empty:
    candidates.append(("vendor", v_row["vw_percent"].iloc[0]))
if not s_row.empty:
    candidates.append(("store", s_row["vw_percent"].iloc[0]))
if not c_row.empty:
    candidates.append(("ceco", c_row["vw_percent"].iloc[0]))
for src, val in candidates:
    print(f"   via {src}: {val}")
first_valid = next((val for _, val in candidates if pd.notna(val)), None)
print(f"\n==> vw_percent final para FE1173: {first_valid}")
print("    (si sale None aun habiendo matches, significa que la fila elegida por")
print("    drop_duplicates tiene vw_percent=NaN y la cascada no sobrepasa a otra fila buena)")
