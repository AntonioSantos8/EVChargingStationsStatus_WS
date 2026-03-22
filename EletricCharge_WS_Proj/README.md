```markdown
# 🔌 EV Scraper — Postos de Carregamento Elétrico

Coleta automática de dados de postos de carregamento elétrico via TupiMob,
com análise de utilização e rentabilidade em Excel.

---

## 📁 Arquivos do projeto

| Arquivo | Descrição |
|---|---|
| `scraper.py` | Script principal de coleta (roda a cada 15 min) |
| `excel_manager.py` | Gerencia criação e atualização do Excel |
| `demo_data.py` | Gera dados fictícios para testar o layout |
| `requirements.txt` | Dependências Python |
| `postos_carregamento.xlsx` | Arquivo de saída (criado automaticamente) |

---

## ⚙️ Instalação

### 1. Pré-requisitos
- Python 3.10+
- Google Chrome instalado
- ChromeDriver compatível com sua versão do Chrome

### 2. Instalar dependências
```bash
pip install -r requirements.txt
```

### 3. ChromeDriver (automático via webdriver-manager)
O script usa `webdriver-manager` para baixar o ChromeDriver automaticamente.
Se preferir manual: baixe em https://chromedriver.chromium.org e ajuste o
caminho na função `build_driver()` em `scraper.py`.

---

## 🚀 Como usar

### Testar com dados fictícios (sem Chrome)
```bash
python demo_data.py
```
Gera o arquivo `postos_carregamento.xlsx` com 32 coletas simuladas.

### Rodar o scraper real
```bash
python scraper.py
```
- Faz a primeira coleta imediatamente
- Repete a cada 15 minutos automaticamente
- Pressione `Ctrl+C` para encerrar

---

## ⚙️ Configurações em `scraper.py`

```python
TUPIMOB_URL = "https://www.tupimob.com.br/mapa"  # URL do mapa
COLLECTION_INTERVAL_MINUTES = 15                 # intervalo de coleta
EXCEL_FILE = "postos_carregamento.xlsx"          # nome do arquivo

CENTER_LAT  = -26.9194   # latitude do centro (ex: Blumenau-SC)
CENTER_LNG  = -49.0661   # longitude do centro
RADIUS_KM   = 50         # raio de busca em km
```

---

## 📊 Estrutura do Excel

### Aba `Registros`
Cada linha = 1 leitura de 1 conector em 1 posto:

| Coluna | Descrição |
|---|---|
| Timestamp Coleta | Data/hora da coleta |
| Nome do Posto | Nome no mapa |
| Endereço | Rua/Avenida |
| Cidade | Cidade |
| Estado | UF |
| Latitude / Longitude | Coordenadas GPS |
| Tipo Conector | CCS2, CHAdeMO, Type2... |
| Potência (kW) | Ex: 50 kW |
| Status Conector | Disponível / Ocupado / Offline |

### Aba `Postos`
Resumo agregado por posto com % de utilização.

### Aba `Análise`
Ranking dos postos por utilização (top 20).

---

## 🔧 Ajuste dos seletores CSS

O TupiMob pode atualizar seu HTML. Se o scraper não encontrar dados:

1. Abra o site no Chrome → F12 (DevTools)
2. Clique em um posto no mapa
3. Inspecione o popup/sidebar que aparece
4. Copie os seletores CSS corretos
5. Atualize a função `_parse_popup()` em `scraper.py`

---

## 📈 Análise de Rentabilidade

Com os dados coletados, é possível calcular:

- **Taxa de utilização por posto** = leituras "Ocupado" / total de leituras
- **Horário de pico** = filtrar por hora na aba Registros
- **Postos mais rentáveis** = cruzar utilização × potência × preço por kWh
- **Rentabilidade estimada** = utilização × potência × horas × tarifa

---

## 📝 Logs

O script gera `scraper.log` com histórico completo das coletas.
