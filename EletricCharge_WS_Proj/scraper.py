import time
import json
import logging
import requests
import schedule
from datetime import datetime
from excel_manager import ExcelManager

SEARCH_CITY = "Blumenau"
INTERVAL = 15
EXCEL_FILE = "postos_carregamento.xlsx"

API_URL = "https://api.tupinambaenergia.com.br/stationsShortVersion"
HEADERS = {
    "Accept": "*/*",
    "Origin": "https://eletropostos-tupi.web.app",
    "Referer": "https://eletropostos-tupi.web.app/",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/146.0.0.0 Safari/537.36",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(message)s",
    handlers=[
        logging.FileHandler("scraper.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)


def fetch():
    params = {
        "plugTypes": '["Tipo 2","CCS 2","CHAdeMO"]',
        "fast": "false",
        "searchText": SEARCH_CITY,
    }
    try:
        r = requests.get(API_URL, headers=HEADERS, params=params, timeout=20)
        if r.status_code not in (200, 304):
            logging.error(f"HTTP {r.status_code}")
            return []
        data = r.json()
        logging.info(f"{len(data)} postos recebidos")
        return data
    except Exception as e:
        logging.error(f"erro na requisicao: {e}")
        return []


def parse(raw):
    result = []
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for item in raw:
        if not isinstance(item, dict):
            continue

        connectors = []
        for c in (item.get("connectors") or item.get("plugs") or []):
            if not isinstance(c, dict):
                continue
            power = c.get("power") or c.get("maxPower") or "N/A"
            connectors.append({
                "tipo": str(c.get("type") or c.get("standard") or "N/A").upper(),
                "potencia": f"{power} kW" if power != "N/A" else "N/A",
                "status": get_status(str(c.get("status") or c.get("availability") or "")),
            })

        if not connectors:
            power = item.get("power") or "N/A"
            connectors = [{
                "tipo": "N/A",
                "potencia": f"{power} kW" if power != "N/A" else "N/A",
                "status": get_status(str(item.get("status") or "")),
            }]

        result.append({
            "nome": item.get("name") or item.get("title") or "N/A",
            "endereco": item.get("address") or item.get("street") or "",
            "cidade": item.get("city") or item.get("cidade") or "",
            "estado": item.get("state") or item.get("uf") or "",
            "lat": str(item.get("lat") or item.get("latitude") or ""),
            "lng": str(item.get("lng") or item.get("longitude") or ""),
            "conectores": connectors,
            "coleta_timestamp": ts,
        })

    return result


def get_status(text):
    t = text.lower()
    if any(w in t for w in ["disponível", "livre", "available", "free", "idle", "true"]):
        return "Disponível"
    if any(w in t for w in ["ocupado", "busy", "charging", "em uso", "in_use", "false"]):
        return "Ocupado"
    if any(w in t for w in ["offline", "inativo", "error", "erro", "falha"]):
        return "Offline"
    return "Desconhecido"


def run():
    logging.info(f"coletando... ({SEARCH_CITY})")
    raw = fetch()
    if not raw:
        # salva resposta bruta pra debug se vier vazia
        with open("debug.json", "w", encoding="utf-8") as f:
            json.dump(raw, f, ensure_ascii=False, indent=2)
        return

    stations = parse(raw)
    em = ExcelManager(EXCEL_FILE)
    em.append_records(stations)
    em.update_summary()
    logging.info(f"salvo: {len(stations)} postos -> {EXCEL_FILE}")


if __name__ == "__main__":
    run()
    schedule.every(INTERVAL).minutes.do(run)
    while True:
        schedule.run_pending()
        time.sleep(30)