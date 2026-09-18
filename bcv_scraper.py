"""
bcv_scraper.py
--------------
Descarga el Excel trimestral del BCV (tipo de cambio de referencia SMC),
localiza la hoja correspondiente a una fecha determinada (nombrada como
"DDMMYYYY", ej. "31032026") y extrae de esa hoja TODAS las monedas listadas:
para cada fila, toma la etiqueta de la columna B (ej. USD, EUR, CNY, TRY, RUB...)
junto con su tasa en la columna G.

Patrón de nombre de archivo:
    https://www.bcv.org.ve/sites/default/files/EstadisticasGeneral/2_1_2{letra}{aa}_smc.xls

    - "2_1_2"  -> fijo
    - {letra}  -> trimestre: a = Q1 (ene-mar), b = Q2 (abr-jun), c = Q3 (jul-sep), d = Q4 (oct-dic)
    - {aa}     -> últimos 2 dígitos del año (25 = 2025, 26 = 2026, ...)
    - "_smc.xls" -> fijo

Lógica de ejecución (pensada para correr 1 vez al día vía cron/GitHub Actions):
    1. Lee bcv_rates.csv (si existe) y determina la última fecha ya capturada
       (formato largo: una fila por moneda, así que se agrupa por columna "fecha").
    2. Si no hay historial -> solo intenta capturar el día de AYER.
       Si hay historial -> hace backfill desde (última_fecha + 1) hasta AYER,
       por si el scraper estuvo caído varios días (cruza de trimestre incluido).
    3. Para cada fecha objetivo, arma la URL del trimestre correspondiente,
       descarga (con caché en memoria para no re-descargar el mismo trimestre),
       busca la hoja "DDMMYYYY" y extrae TODAS las filas con etiqueta en B y
       tasa numérica en G (USD, EUR y cualquier otra moneda que el BCV liste).
    4. Si la hoja de una fecha no existe (fin de semana / feriado / aún no
       publicada), simplemente se omite esa fecha sin marcar error.
    5. Escribe/agrega filas nuevas a bcv_rates.csv -> si una fecha ya está
       capturada (ya tiene filas en el CSV), se omite por completo esa fecha
       para no duplicar.

Uso:
    python bcv_scraper.py            -> corre normal
    python bcv_scraper.py --debug    -> además imprime el contenido crudo de
                                         la hoja encontrada, para verificar
                                         manualmente qué filas trae antes de
                                         confiar en la extracción automática
    python bcv_scraper.py --date 31-03-2026
                                      -> fuerza una fecha específica en vez de "ayer"
                                         (útil para pruebas o para rellenar un hueco puntual)
"""

import argparse
import csv
import json
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import requests
import urllib3

# El servidor de bcv.org.ve tiene un certificado SSL mal configurado
# (le falta el certificado intermedio). Por eso, para ESTE sitio en
# particular, desactivamos la verificación estricta y silenciamos el
# warning que eso genera. No afecta la conexión a ningún otro sitio.
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:
    import xlrd  # los archivos del BCV son .xls (formato viejo de Excel)
except ImportError:
    xlrd = None


BASE_URL = "https://www.bcv.org.ve/sites/default/files/EstadisticasGeneral/2_1_2{letra}{aa}_smc.xls"

QUARTER_LETTER = {1: "a", 2: "b", 3: "c", 4: "d"}  # trimestre -> letra

OUTPUT_CSV = "bcv_rates.csv"
OUTPUT_JSON = "rates.json"

# Columnas B y G (1-based: A=1, B=2, ... G=7)
LABEL_COLUMN_INDEX = 2   # columna B: etiqueta de la moneda
RATE_COLUMN_INDEX = 7    # columna G: tasa

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}


# --------------------------------------------------------------------------
# Utilidades de fecha / URL
# --------------------------------------------------------------------------

def quarter_file_url(d: date) -> str:
    quarter = (d.month - 1) // 3 + 1
    letra = QUARTER_LETTER[quarter]
    aa = f"{d.year % 100:02d}"
    return BASE_URL.format(letra=letra, aa=aa)


def sheet_name_for_date(d: date) -> str:
    """Formato esperado en el Excel: DDMMYYYY, ej. 31032026"""
    return d.strftime("%d%m%Y")


# --------------------------------------------------------------------------
# Descarga y lectura del Excel
# --------------------------------------------------------------------------

def download_file(url: str) -> bytes:
    resp = requests.get(url, headers=HEADERS, timeout=30, verify=False)
    resp.raise_for_status()
    return resp.content


def load_sheets(content: bytes):
    """Devuelve {nombre_hoja: filas} para un .xls"""
    if xlrd is None:
        raise RuntimeError("Instala xlrd: pip install xlrd==1.2.0")
    wb = xlrd.open_workbook(file_contents=content)
    return {sheet.name.strip(): [sheet.row_values(r) for r in range(sheet.nrows)] for sheet in wb.sheets()}


def parse_number(val):
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, str):
        cleaned = val.replace(".", "").replace(",", ".").strip()
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def extract_all_currencies(rows, debug=False, sheet_name=""):
    """
    Recorre todas las filas de la hoja. Para cada fila cuya columna B trae
    una etiqueta de texto no vacía y cuya columna G trae un número válido,
    la toma como una moneda capturada (USD, EUR, CNY, TRY, RUB, o cualquier
    otra que el BCV liste ese trimestre).
    """
    if debug:
        print(f"\n--- DEBUG hoja: {sheet_name} ---")
        for i, row in enumerate(rows[:25]):
            print(i, row)

    results = []
    for row in rows:
        if len(row) < max(LABEL_COLUMN_INDEX, RATE_COLUMN_INDEX):
            continue

        raw_label = row[LABEL_COLUMN_INDEX - 1]
        label = str(raw_label).strip().upper() if raw_label is not None else ""
        rate = parse_number(row[RATE_COLUMN_INDEX - 1])

        # Se descarta si no hay etiqueta de texto real o si G no es un número
        if not label or rate is None:
            continue
        # Evita capturar filas de encabezado tipo "MONEDA" / "TASA" / "TIPO DE CAMBIO"
        if label in {"MONEDA", "CODIGO", "CÓDIGO", "TASA", "TIPO DE CAMBIO", "G"}:
            continue

        results.append((label, rate))

    return results


# --------------------------------------------------------------------------
# CSV histórico / control de duplicados (formato largo: fecha, moneda, tasa)
# --------------------------------------------------------------------------

def load_existing_dates(csv_path: Path):
    """Devuelve el set de fechas (YYYY-MM-DD) que ya tienen al menos una fila."""
    if not csv_path.exists():
        return set()
    with csv_path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        return {row["fecha"] for row in reader}


def append_rows(csv_path: Path, rows):
    write_header = not csv_path.exists()
    with csv_path.open("a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["fecha", "moneda", "tasa", "fuente"])
        if write_header:
            writer.writeheader()
        writer.writerows(rows)


def rebuild_json(csv_path: Path, json_path: Path):
    """
    Lee TODO el CSV histórico y arma un JSON agrupado por fecha, con esta forma:

    {
      "actualizado": "2026-09-18T13:05:00+00:00",
      "fechas": {
        "2026-09-17": {"USD": 199.68, "EUR": 233.45},
        "2026-09-16": {"USD": 198.90, "EUR": 231.10}
      }
    }

    Se reconstruye completo cada vez (no incrementalmente) para que el CSV
    siga siendo la única "fuente de verdad" y el JSON nunca quede desfasado.
    """
    fechas = {}
    if csv_path.exists():
        with csv_path.open("r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                fecha = row["fecha"]
                moneda = row["moneda"]
                tasa = row["tasa"]
                try:
                    tasa = float(tasa)
                except (TypeError, ValueError):
                    pass
                fechas.setdefault(fecha, {})[moneda] = tasa

    payload = {
        "actualizado": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "fechas": dict(sorted(fechas.items(), reverse=True)),  # más reciente primero
    }

    with json_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def last_captured_date(csv_path: Path):
    existing = load_existing_dates(csv_path)
    if not existing:
        return None
    return max(datetime.strptime(d, "%Y-%m-%d").date() for d in existing)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def dates_to_process(forced_date: date = None):
    if forced_date:
        return [forced_date]

    yesterday = date.today() - timedelta(days=1)
    last_date = last_captured_date(Path(OUTPUT_CSV))

    if last_date is None:
        # Primera corrida: solo trae el día de ayer, no todo el histórico
        return [yesterday]

    if last_date >= yesterday:
        return []  # ya está al día, nada que hacer

    # Backfill: desde el día siguiente a la última fecha capturada hasta ayer
    result = []
    d = last_date + timedelta(days=1)
    while d <= yesterday:
        result.append(d)
        d += timedelta(days=1)
    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--date", type=str, help="Formato DD-MM-YYYY, fuerza una fecha específica")
    args = parser.parse_args()

    forced_date = None
    if args.date:
        forced_date = datetime.strptime(args.date, "%d-%m-%Y").date()

    targets = dates_to_process(forced_date)
    if not targets:
        print("Ya está al día, no hay fechas nuevas que capturar.")
        return

    existing_dates = load_existing_dates(Path(OUTPUT_CSV))
    file_cache = {}  # url -> {sheet_name: rows}
    new_rows = []

    for d in targets:
        fecha_iso = d.strftime("%Y-%m-%d")
        if fecha_iso in existing_dates:
            print(f"{fecha_iso}: ya estaba en el CSV, se omite (sin duplicar).")
            continue

        url = quarter_file_url(d)
        sheet_key = sheet_name_for_date(d)

        if url not in file_cache:
            print(f"Descargando trimestre: {url}")
            try:
                content = download_file(url)
                file_cache[url] = load_sheets(content)
            except requests.RequestException as e:
                print(f"  ERROR descargando {url}: {e}")
                file_cache[url] = {}
            except Exception as e:
                print(f"  ERROR leyendo {url}: {e}")
                file_cache[url] = {}

        sheets = file_cache[url]
        if sheet_key not in sheets:
            print(f"{fecha_iso}: hoja '{sheet_key}' no encontrada en {url} "
                  f"(puede ser fin de semana, feriado, o aún no publicada). Se omite.")
            continue

        currencies = extract_all_currencies(sheets[sheet_key], debug=args.debug, sheet_name=sheet_key)
        if not currencies:
            print(f"{fecha_iso}: no se encontraron monedas válidas en la hoja '{sheet_key}'.")
            continue

        resumen = ", ".join(f"{moneda}={tasa}" for moneda, tasa in currencies)
        print(f"{fecha_iso}: {resumen}")

        for moneda, tasa in currencies:
            new_rows.append({"fecha": fecha_iso, "moneda": moneda, "tasa": tasa, "fuente": url})

    if new_rows:
        append_rows(Path(OUTPUT_CSV), new_rows)
        print(f"\n{len(new_rows)} fila(s) nueva(s) agregada(s) a {OUTPUT_CSV}")
    else:
        print("\nNo se agregó ninguna fila nueva.")

    # Reconstruye rates.json siempre, aunque no haya filas nuevas hoy,
    # para que el campo "actualizado" refleje la última vez que corrió el scraper.
    rebuild_json(Path(OUTPUT_CSV), Path(OUTPUT_JSON))
    print(f"{OUTPUT_JSON} actualizado.")


if __name__ == "__main__":
    main()
