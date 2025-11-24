# 🌳 Random Forest per rilevare la presenza in una stanza

## 🎯 Obbiettivo
L'obiettivo di questo progetto è usare un algoritmo di **machine learning** (Random Forest) per prevedere la presenza di persone in una stanza, utilizzando i **dati provenienti da dei sensori ambientali**.

## 🛠️ Applicazioni
Alcune possibili applicazioni di questo lavoro possono essere:
- Ottimizzazione del **consumo energetico** negli edifici
- Sistemi di **sicurezza**
- **Organizzazione** e gestione degli spazi

## 🗂️ Organizzazione del progetto
### 1. Visualizzazione
All'interno del notebook [visualization.ipynb](https://github.com/gabrielee20/roomoccupation/blob/main/visualization.ipynb) viene visualizzato il dataset, contenente le informazioni provenienti dai sensori ambientali.

Dalle visualizzazioni abbiamo cercato di capire quali features potevano essere importanti per la classificazione e se esistevano dei “pattern” che si ripetevano.

### 2. Classificazione
All'interno del notebook [classification.ipynb](https://github.com/gabrielee20/roomoccupation/blob/main/classification.ipynb) viene svolta la classificazione, in particolare all'interno si possono trovare i passaggi svolti per:
- Fare il **subsampling** per ottenere training test e testing set
- Selezionare gli **iperparametri**
- Fare il **fitting** del modello
- Svolgere la **valutazione**

[Link al dataset utilizzato](https://archive-beta.ics.uci.edu/dataset/357/occupancy+detection)
[Link alla presentazione del lavoro svolto](https://docs.google.com/presentation/d/1hAzFPP0qTSyECLbZrUDlQdgpPHK8fywlUraZr2Qwywo/edit?usp=sharing)

