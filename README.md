# Compilatore PowerPoint DEIA — app Streamlit

App in cui carichi un **PowerPoint** (template con segnaposto) e i **risultati
della survey** (Excel export Qualtrics); l'app sostituisce i segnaposto e ti fa
scaricare il PowerPoint compilato. La libreria DEIA (`deia_mapping_GM.xlsx`)
non si carica più da frontend: è letta da un file fisso sul server.

## Avvio

```bash
pip install -r requirements.txt
streamlit run app.py
```

Si apre nel browser. Carica i due file e premi **Scarica il PowerPoint compilato**.

## Come scrivere i segnaposto nel PowerPoint

Nel testo delle slide (anche dentro le tabelle) usa le doppie graffe:

| Segnaposto | Diventa |
|---|---|
| `{{strategia.risultato}}` | `2 - ROTTA` |
| `{{strategia.commento}}` | commento dell'area al livello scelto |
| `{{strategia_pianificazione.attivita}}` | attività del sottogruppo al livello scelto |

Convenzione: `{{<id>.risultato}}`, `{{<id>.commento}}` (aree), `{{<id>.attivita}}` (sottogruppi).
Gli `id` sono nel foglio **Libreria** dell'Excel.

## L'Excel di mapping (`deia_mapping.xlsx`)

- **Mapping** — `segnaposto | testo`: è ciò che l'app usa di default.
- **Assessment** — `id | tipo | nome | livello`: cambia il numero di livello (1-4).
- **Libreria** — tutti e 4 i livelli per ogni area/sottogruppo (riferimento).
- **Scala** — i nomi dei 4 livelli (modificabili).

Due modi per cambiare i contenuti:
1. **Veloce** — modifica la colonna `testo` nel foglio *Mapping*.
2. **Dinamico** — cambia il `livello` nel foglio *Assessment* e nell'app spunta
   *"Rigenera il mapping dai livelli"*: l'app ricalcola i testi da *Assessment* + *Libreria*.

## File inclusi

| File | Cosa è |
|---|---|
| `app.py` | l'app Streamlit |
| `pptx_filler.py` | logica di compilazione (riusabile e testabile) |
| `requirements.txt` | dipendenze |
| `template_esempio.pptx` | PowerPoint di prova con i segnaposto (una slide per area) |
| `deia_mapping.xlsx` | Excel di mapping già pronto |
| `deia_framework.json` | libreria completa dei contenuti (fonte dati) |
| `build_*.py` | script con cui sono stati generati json / excel / template |

## Note

- La scala usa una metafora nautica: **1 Deriva · 2 Rotta · 3 Navigazione · 4 Approdo**.
  I nomi dei livelli 3 e 4 sono un'ipotesi: cambiali nel foglio *Scala* (o in
  `deia_framework.json` → `meta.scala`) e si aggiornano ovunque.
- I testi dei livelli non presenti nel documento originale sono generati nello
  stesso stile, con progressione di maturità coerente. Vanno riletti/adattati.
- La sostituzione gestisce segnaposto spezzati su più *run*, tabelle e forme
  raggruppate. Se un segnaposto resta nel file, l'app lo segnala.
