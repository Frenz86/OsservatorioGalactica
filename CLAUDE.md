# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Cos'è

App Streamlit (in italiano) che compila un PowerPoint di report per un assessment di maturità DEIA (Diversità, Equità, Inclusione, Accessibilità) a partire dai risultati di una survey Qualtrics. L'utente carica: template PPTX con segnaposto `{{...}}`, l'Excel export della survey, e la "libreria" DEIA (Excel con testi/pesi); l'app calcola i punteggi, genera il mapping segnaposto→testo e produce il PowerPoint compilato + un Excel con i calcoli.

## Comandi

```bash
pip install -r requirements.txt
streamlit run app.py        # avvia l'app
python verify.py            # test end-to-end della pipeline (non pytest, script con asserzioni proprie)
```

Non esiste una suite pytest né un linter configurato: `verify.py` è l'unico test — va eseguito per intero (non ci sono singoli test isolabili) e stampa `OK`/`FAIL` riga per riga.

Gli script `build_*.py` rigenerano i file dati/template da `deia_framework.json` (vedi sotto) e vanno eseguiti dalla root:
```bash
python build_framework.py       # deia_framework.json <- editando lo script stesso
python build_mapping_excel.py   # deia_mapping.xlsx <- deia_framework.json
python build_template_pptx.py   # template_esempio.pptx <- deia_framework.json
```

## Architettura

**`app.py`** è solo l'interfaccia Streamlit (login hardcoded `admin/admin` in `CREDENZIALI`, upload di template PPTX + risultati survey, diagnostica segnaposto mancanti, download). La libreria DEIA (`deia_mapping_GM.xlsx`) **non** si carica da frontend: è letta da un percorso fisso (`MAPPING_PATH`, accanto ad `app.py`) — per cambiarla in produzione bisogna sostituire quel file sul server. Tutta la logica sta in **`pptx_filler.py`**, scritto per essere testabile senza Streamlit.

### Flusso dati in `pptx_filler.py`

`mapping_from_survey(input_xlsx, mapping_xlsx)` è il punto d'ingresso usato da `app.py`: rileva automaticamente il formato della survey caricata e delega a una delle due pipeline:

- **`mapping_from_raw`** — survey Qualtrics *raw* (colonne domanda `CQ/PQ/SQ/MQ*`, nessuno score precalcolato). Ricalcola gli score pesati (`Σ risposta×peso / Σ peso`) usando il foglio **Mappatura** di `mapping_xlsx` (domanda → categoria/sottogruppo + peso PMI/peso GRANDI). Il tipo impresa (PMI/GRANDI) è derivato dalla colonna `A3` (fascia dipendenti). `PQ19` è esclusa dal calcolo (peso nullo, vedi `_PESO_ZERO_COD`).
- **`mapping_from_qualtrics`** — survey con colonne `Score <area>` / `Score <area> - <sottogruppo>` già calcolate; aggrega più rispondenti per media.

Entrambe convertono lo score medio in livello 1-4 tramite soglie fisse (`_score_to_level`: ≤1.50→1, ≤2.50→2, ≤3.50→3, altrimenti 4) e producono lo stesso `mapping` dict `{segnaposto: testo}` + un `details` dict (medie per area/sottogruppo/panoramica, usato per l'export Excel via `build_calc_workbook`).

Un terzo percorso, **`load_mapping`**, legge un mapping statico da Excel (fogli `Mapping` diretto, oppure `Assessment`+`Libreria` per ricostruirlo dai livelli) — è la logica più vecchia, usata solo da `verify.py`, non dal flusso app corrente basato su survey.

### Convenzione segnaposto

Nel PPTX, dentro doppie graffe: `{{<id>.risultato}}` → `"<livello> - <NOME>"`; `{{<id>.commento}}` (solo aree) e `{{<id>.attivita}}` (solo sottogruppi) → testo del livello. Gli `id` sono quelli del foglio **Libreria**. Le 4 aree (`AREA_ORDER` in `pptx_filler.py`: strategia, cultura, processi, mercato) si riconoscono perché l'`id` non contiene `_`; i sottogruppi sì (es. `strategia_pianificazione`). `{{panoramica.*}}` è calcolato aggregando le 4 aree. `{{radar.aree}}` è un marcatore speciale: non testo ma sostituito con un'immagine radar generata al volo con matplotlib (`_install_dynamic_images`/`_radar_image`), quindi va gestito *prima* della sostituzione testuale in `fill_pptx`.

La sostituzione nel PPTX (`_replace_in_text_frame`) ricompone segnaposto spezzati su più "run" all'interno dello stesso paragrafo (limite di python-pptx quando il testo viene editato manualmente in PowerPoint), scrive tutto il testo risolto nel primo run e svuota gli altri, e attraversa ricorsivamente anche le forme raggruppate (`_iter_shapes`) e le tabelle.

### Generazione dei file sorgente (`build_*.py`)

`deia_framework.json` è la fonte dati unica (single source of truth): contiene la scala a 4 livelli (`meta.scala`, nomi editabili in un solo posto) e, per ognuna delle 4 aree e dei loro 3 sottogruppi ciascuna, i testi per tutti e 4 i livelli di maturità. `build_framework.py` lo genera (i testi sono hardcoded nello script) e valida la completezza (nessun testo mancante). Da questo JSON, `build_mapping_excel.py` genera `deia_mapping.xlsx` (fogli Mapping/Assessment/Libreria/Scala) e `build_template_pptx.py` genera `template_esempio.pptx` (una slide per area coi segnaposto). Modificare i testi/la scala si fa editando `build_framework.py` e rieseguendo la catena degli script, **non** editando direttamente il JSON o l'xlsx generato.

Nota: in produzione l'app usa `deia_mapping_GM.xlsx`, un file di libreria/mappatura più recente e reale (con foglio `Mappatura` per i pesi PMI/GRANDI) distinto dal `deia_mapping.xlsx` "di esempio" generato dagli script `build_*`.

### File di dati non versionabili come codice

`_dati_iniziali/` contiene i file Excel/PowerPoint originali forniti dal cliente (fonte del contenuto DEIA); `qualtrix_output.xlsx` e `template_esempio_compilato.pptx` sono output di esempio/test, non input da modificare a mano.
