"""
demo_data.py — Gera dados de demonstração no Excel
para validar o layout antes de rodar o scraper real.
Execute: python demo_data.py
"""

import random
from datetime import datetime, timedelta
from excel_manager import ExcelManager

DEMO_STATIONS = [
    {
        "nome": "Posto Eletroposto Shopping Neumarkt",
        "endereco": "Rua Cel. Marcos Konder, 100",
        "cidade": "Blumenau",
        "estado": "SC",
        "lat": -26.9135,
        "lng": -49.0630,
    },
    {
        "nome": "Charge&Go - Posto Vila Nova",
        "endereco": "Av. Brasil, 2500",
        "cidade": "Blumenau",
        "estado": "SC",
        "lat": -26.9250,
        "lng": -49.0712,
    },
    {
        "nome": "EV Point - Park Shopping",
        "endereco": "Rua Hermann Hering, 440",
        "cidade": "Blumenau",
        "estado": "SC",
        "lat": -26.8980,
        "lng": -49.0755,
    },
    {
        "nome": "Eletroposto Brusque Centro",
        "endereco": "Av. Cel. Marcos Rovaris, 1000",
        "cidade": "Brusque",
        "estado": "SC",
        "lat": -27.0967,
        "lng": -48.9159,
    },
    {
        "nome": "Charge&Go - Indaial Norte",
        "endereco": "Rua Joinville, 300",
        "cidade": "Indaial",
        "estado": "SC",
        "lat": -26.8983,
        "lng": -49.2322,
    },
    {
        "nome": "EV Fast - Gaspar",
        "endereco": "Av. Santa Terezinha, 500",
        "cidade": "Gaspar",
        "estado": "SC",
        "lat": -26.9311,
        "lng": -48.9597,
    },
]

CONNECTORS = [
    {"tipo": "CCS2", "potencia": "50 kW"},
    {"tipo": "CCS2", "potencia": "150 kW"},
    {"tipo": "TYPE2", "potencia": "22 kW"},
    {"tipo": "CHADEMO", "potencia": "50 kW"},
    {"tipo": "CCS2", "potencia": "30 kW"},
]

STATUSES = ["Disponível", "Ocupado", "Ocupado", "Disponível", "Offline"]

def generate_demo():
    em = ExcelManager("postos_carregamento.xlsx")

    # Simula 8 horas de coleta (a cada 15 min = 32 coletas)
    base_time = datetime.now() - timedelta(hours=8)

    print("Gerando dados de demonstração...")
    for i in range(32):
        timestamp = (base_time + timedelta(minutes=15 * i)).strftime("%Y-%m-%d %H:%M:%S")
        records = []
        for station in DEMO_STATIONS:
            n_connectors = random.randint(1, 3)
            connectors = []
            for _ in range(n_connectors):
                conn = random.choice(CONNECTORS).copy()
                # Simula maior utilização no horário comercial
                hour = (base_time + timedelta(minutes=15 * i)).hour
                if 8 <= hour <= 19:
                    status = random.choices(
                        ["Ocupado", "Disponível", "Offline"],
                        weights=[55, 40, 5]
                    )[0]
                else:
                    status = random.choices(
                        ["Ocupado", "Disponível", "Offline"],
                        weights=[20, 75, 5]
                    )[0]
                conn["status"] = status
                connectors.append(conn)

            rec = station.copy()
            rec["conectores"] = connectors
            rec["coleta_timestamp"] = timestamp
            records.append(rec)

        em.append_records(records)
        print(f"  Coleta {i+1}/32 — {timestamp}")

    em.update_summary()
    print("\n✓ Excel gerado com sucesso: postos_carregamento.xlsx")
    print("  Abas: Registros | Postos | Análise")

if __name__ == "__main__":
    generate_demo()
