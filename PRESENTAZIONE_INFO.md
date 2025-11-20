# Presentazione Progetto Room Occupancy

## File Generato
**Nome file:** `Presentazione_Occupazione_Stanza.pptx`

## Descrizione
Presentazione PowerPoint dettagliata che spiega il lavoro svolto nel progetto di predizione dell'occupazione della stanza utilizzando Machine Learning con algoritmo Random Forest.

## Caratteristiche
- **Numero di slide:** 28 (requisito minimo: 20) ✓
- **Formato:** Microsoft PowerPoint 2007+ (.pptx)
- **Dimensione:** 1.9 MB
- **Lingua:** Italiano
- **Immagini incluse:** 7 grafici/visualizzazioni

## Contenuto della Presentazione

### Struttura (28 slide)

1. **Slide 1:** Titolo principale - "Predizione dell'Occupazione della Stanza"
2. **Slide 2:** Indice della presentazione
3. **Slide 3:** Introduzione al progetto
4. **Slide 4:** Obiettivi specifici del progetto
5. **Slide 5:** Dataset utilizzati
6. **Slide 6:** Caratteristiche del dataset
7. **Slide 7:** Distribuzione dell'occupazione per giorno (con immagine)
8. **Slide 8:** Heatmap campioni per ora (con immagine)
9. **Slide 9:** Andamento temporale delle features (con immagine)
10. **Slide 10:** Box plots distribuzione features (con immagine)
11. **Slide 11:** Analisi dei box plots
12. **Slide 12:** Scelta dell'algoritmo Random Forest
13. **Slide 13:** Metodologia di ottimizzazione
14. **Slide 14:** Frequenze parametri migliori (con immagine)
15. **Slide 15:** Parametri ottimali selezionati
16. **Slide 16:** Training del modello
17. **Slide 17:** Metriche di valutazione utilizzate
18. **Slide 18:** Risultati sul testing set
19. **Slide 19:** Validazione su dataset separato
20. **Slide 20:** Importanza delle features (con immagine)
21. **Slide 21:** Interpretazione feature importance
22. **Slide 22:** Esempio albero della Random Forest (con immagine)
23. **Slide 23:** Vantaggi della soluzione implementata
24. **Slide 24:** Applicazioni pratiche
25. **Slide 25:** Limitazioni e possibili miglioramenti
26. **Slide 26:** Stack tecnologico
27. **Slide 27:** Conclusioni
28. **Slide 28:** Ringraziamenti e Q&A

## Immagini Incluse
Tutte le visualizzazioni disponibili nella cartella `plots/` sono state integrate nella presentazione:

- `campioni_con_presenza_per_giorno.png`
- `heatmap_campioni_per_ora.png`
- `andamento_temporale_features.png`
- `box_plots.png`
- `frequenze_parametri_migliori_50_grid_search.png`
- `feature_importances_random_forest.png`
- `tree_random_forest.png`

## Come Generare la Presentazione

Se si desidera rigenerare la presentazione:

```bash
python3 create_presentation.py
```

## Requisiti
- Python 3.x
- python-pptx library

Per installare le dipendenze:
```bash
pip install python-pptx
```

## Note
La presentazione è stata creata automaticamente usando il modulo `python-pptx` di Python, integrando tutte le informazioni dai notebook Jupyter (`classification.ipynb` e `visualization.ipynb`) e le visualizzazioni generate durante l'analisi.

## Utilizzo
Il file `.pptx` può essere aperto con:
- Microsoft PowerPoint
- LibreOffice Impress
- Google Slides
- Apple Keynote
- Qualsiasi altro software compatibile con il formato PowerPoint

---
**Data di creazione:** 2025-11-20
**Autore:** Generato automaticamente dal progetto roomoccupation
