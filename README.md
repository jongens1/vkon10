# Dashboard provoz

Streamlit dashboard nad uloženou kontingenčnou tabuľkou v `datax.xlsx`, hárok `DATA`.

## Spustenie

Vyžaduje Python 3.10 alebo novší. V priečinku repozitára:

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

V bočnom paneli vyberte dátumy a agregáciu hodina/deň. Grafy porovnávajú vzniky a zaradenie do fronty pickovania pre jobline aj kusy. Tabuľka obsahuje rovnaké agregované dáta; prázdne časové intervaly majú nulu.

## Význam dát a parser

- Vzniky sa filtrujú podľa dátumu/hodiny vytvorenia; picky podľa dátumu/hodiny fronty. Ide o dve samostatné časové osi, nie o kohortu vytvorenú vo vybranom období.
- Picky znamenajú zaradenie do fronty, nie dokončenie pickovania.
- Pick dátum `01.01.1900` (a starší) sa nezahrnie do grafov; samostatné KPI ho filtruje podľa času vzniku.
- Parser vyhľadáva názvy hlavičiek bez ohľadu na diakritiku, rozpozná viacriadkové hlavičky a doplní opakované dátumy/hodiny. Nepoužíva pevné čísla riadkov ani stĺpcov.
- Celkové súčty sa nezapočítavajú druhýkrát. Ak sú dostupné riadkové súčty, parser kontroluje ich zhodu so súčtom detailov.
- Neplatné hodnoty alebo nejednoznačné mapovanie blokujú grafy a zobrazia diagnostiku. Expandér „Diagnostika dát“ obsahuje mapovanie, počet riadkov a ukážky normalizovaných dát.
- Aplikácia číta uložené hodnoty; neobnovuje OLAP pripojenie. Pre nové dáta aktualizujte Excel a uložte ho. Cache sa obnoví pri zmene veľkosti alebo času úpravy súboru.

Závislosti: Streamlit, pandas, openpyxl, Plotly.
