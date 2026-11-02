# 🏫 Gestione Sostituzioni Docenti (v3.3)

> **Applicazione desktop professionale per l'automazione, l'ottimizzazione e l'assegnazione trasparente ed equa delle supplenze scolastiche giornaliere.**  
> *Sviluppata specificamente per Istituti Scolastici Superiori, Istituti Alberghieri (IPSEOA) e Secondari di Secondo Grado.*

[![Python](https://img.shields.io/badge/Python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-blue.svg)](https://www.python.org/)
[![GUI](https://img.shields.io/badge/GUI-Tkinter-brightgreen.svg)]()
[![Database](https://img.shields.io/badge/Database-SQLite3-lightgrey.svg)]()
[![Platform](https://img.shields.io/badge/Platform-macOS%20%7C%20Windows-success.svg)]()

---

## 📖 Indice dei Contenuti
1. [Panoramica e Obiettivi](#-panoramica-e-obiettivi)
2. [Algoritmo di Assegnazione e Priorità](#-algoritmo-di-assegnazione-e-priorità)
3. [I 6 Riquadri Dati e Modelli Excel Ufficiali](#-i-6-riquadri-dati-e-modelli-excel-ufficiali)
4. [Materia Alternativa e Religione Cattolica](#-materia-alternativa-e-religione-cattolica)
5. [Procedura Operativa Giornaliera](#-procedura-operativa-giornaliera)
6. [Funzionalità Avanzate dell'Interfaccia](#-funzionalità-avanzate-dellinterfaccia)
7. [Esportazione e Stampe Ufficiali](#-esportazione-e-stampe-ufficiali)
8. [Architettura del Progetto](#-architettura-del-progetto)
9. [Installazione e Avvio](#-installazione-e-avvio)

---

## 🏢 Panoramica e Obiettivi

Il software **Gestione Sostituzioni Docenti** automatizza il complesso processo quotidiano di copertura delle classi scoperte a causa di assenze, permessi brevi e uscite didattiche.

L'applicazione incrocia istantaneamente gli orari curricolari, gli orari personali dei docenti, le attività di potenziamento, le ore a disposizione, i docenti di sostegno e le disponibilità a pagamento per proporre la soluzione ottimale ed economica:

* ⏱️ **Risparmio dell'80% del tempo** nell'elaborazione quotidiana del prospetto supplenze.
* ⚖️ **Equità e trasparenza** nelle assegnazioni grazie all'algoritmo *Round-Robin* che bilancia i carichi.
* 🚫 **Zero sovrapposizioni o doppie assegnazioni**: controllo rigoroso degli impegni orari.
* 💰 **Abbattimento dei costi**: priorità assoluta alle risorse interne a costo zero prima del ricorso alle ore eccedenti a pagamento.
* 📝 **Esportazione pronta all'uso** per registro, bacheca e foglio firme giornaliero.

---

## ⚙️ Algoritmo di Assegnazione e Priorità

Per ciascuna ora scoperta individuata dall'incrocio tra orario generale e personale, il motore applica una rigorosa gerarchia di criteri:

| Priorità | Criterio di Copertura | Colore Tabella | Descrizione Funzionale |
| :---: | :--- | :---: | :--- |
| **1** | **Compresenza** *(Riq. 1)* | 🟩 Verde Acqua | Docente curricolare già presente nella stessa classe e ora (teoria/laboratorio, non impegnato in altro gruppo). Copertura interna immediata a costo zero. |
| **2** | **Compresenza Sostegno** *(Riq. 6)* | 🟪 Lilla / Viola | Docente di sostegno già assegnato alla stessa classe e presente a quell'ora (utilizzato prioritariamente a tutela della continuità didattica della classe). |
| **3** | **Docente Liberato da Classe Assente** | 🟨 Giallo Tenue | Docente (curricolare o sostegno) la cui classe è fuori sede per gite, visite guidate, PCTO o stage. Se più docenti sono liberati, il prospetto li elenca tutti. |
| **4** | **Potenziamento (Round-Robin)** *(Riq. 4)* | 🟧 Arancione | Docente con ora di potenziamento. L'algoritmo bilancia equamente le supplenze tra i docenti disponibili nella giornata. |
| **5** | **Ore a Pagamento** *(Riq. 5)* | 🟩 Verde Chiaro | Docente disponibile a svolgere ore eccedenti a pagamento. Utilizzato come risorsa di salvaguardia prima dello scoperto. |
| **6** | **Classe Scoperta** | 🟥 Rosso Chiaro | Nessun sostituto disponibile. La riga evidenzia la classe scoperta per intervento della presidenza/collaboratori. |

---

## 📋 I 6 Riquadri Dati e Modelli Excel Ufficiali

L'applicazione include template Excel `.xlsx` precompilati, formattati e con guide operative integrate per tutti i 6 riquadri di caricamento dati:

| Riquadro | Denominazione | Colore UI | File Modello Excel | Descrizione Contenuto |
| :---: | :--- | :---: | :--- | :--- |
| **1** | **Orario Classi / Generale** | `#00695C` (Teal) | [`Modello_1_Orario_Classi.xlsx`](./Modello_1_Orario_Classi.xlsx) | Orario curricolare di tutte le classi e docenti (blocchi classe, ore 08:00..15:00, righe `Compr`). |
| **2** | **Orario Personale** | `#0277BD` (Blu) | [`Modello_2_Orario_Personale.xlsx`](./Modello_2_Orario_Personale.xlsx) | Orario individuale per cattedra (`DOCENTE [Nome]`), con classi e ore non didattiche (`D`, `R`, `POT`). |
| **3** | **Materia Alternativa** | `#6A1B9A` (Viola) | [`Modello_3_Materia_Alternativa.xlsx`](./Modello_3_Materia_Alternativa.xlsx) | Quadro orario dei docenti di attività alternativa all'IRC (coppie ORA/GIORNO con `CLASSE DOCENTE`). |
| **4** | **Potenziamento** | `#EF6C00` (Arancio) | [`Modello_4_Potenziamento.xlsx`](./Modello_4_Potenziamento.xlsx) | Docenti in potenziamento per fascia oraria e giorno, gestiti con rotazione Round-Robin. |
| **5** | **A Pagamento** | `#2E7D32` (Verde) | [`Modello_5_A_Pagamento.xlsx`](./Modello_5_A_Pagamento.xlsx) | Docenti disponibili per ore eccedenti retribuite (ore 08:00..15:00, cognomi separati da spazio). |
| **6** | **Docenti Sostegno** | `#D35400` (Terracotta) | [`Modello_6_Docenti_Sostegno.xlsx`](./Modello_6_Docenti_Sostegno.xlsx) | Orario settimanale dei docenti di sostegno assegnati alle classi, con riconoscimento automatico delle classi. |

> Ogni cartella di lavoro contiene 3 fogli:
> 1. **Foglio Operativo (`1_*`, `2_*`, ...)**: il template attivo letto direttamente dal programma.
> 2. **`Esempio_Compilato`**: orario realistico completo di riferimento con casi reali.
> 3. **`Istruzioni_e_Regole`**: prontuario dettagliato con regole di scrittura, gestione omonimie e codici classi.

---

## ✝️ Materia Alternativa e Religione Cattolica

L'applicazione gestisce in modo conforme alle norme scolastiche i 3 scenari dell'Insegnamento della Religione Cattolica (IRC) e delle Attività Alternative:

1. **Manca solo Religione (Alternativa presente)**: il programma cerca un sostituto per gli studenti avvalentesi e annota: `ALTERNATIVA: [Docente] (Non avvalentesi)`.
2. **Mancano sia Religione che Alternativa**: la classe viene accorpata e viene assegnato **un solo sostituto** per l'intero gruppo classe, evitando spreco di risorse.
3. **Manca solo Alternativa (Religione presente)**: il docente di religione fa lezione agli avvalentesi, mentre viene individuato un sostituto dedicato alla vigilanza dei soli studenti non avvalentesi.

---

## 📋 Procedura Operativa Giornaliera

1. **Avvio**: All'apertura l'applicazione ripristina istantaneamente l'orario completo dal database locale SQLite.
2. **Selezione Giorno**: Scegliere il giorno della settimana dal menu a tendina (es. *MARTEDÌ*).
3. **Inserimento Docenti Assenti**: Digitare o selezionare i docenti assenti dalla tendina a comparsa (Box Rosso).
4. **Permessi Brevi**: Inserire i docenti con permesso a ore (es. `ROSSI 2 3`) o usare i selettori rapidi da 1ª a 8ª ora (Box Verde Acqua).
5. **Classi Assenti**: Selezionare le classi fuori sede per gite/stage per liberare i rispettivi docenti (Box Arancione).
6. **Elaborazione**: Cliccare sul pulsante verde **PROCEDI**: il motore genera istantaneamente il quadro sostituzioni.
7. **Controllo & Modifiche**: Fare clic destro su qualsiasi riga per personalizzare un sostituto o marcare l'ora come coperta manualmente.
8. **Esportazione**: Cliccare su **Word**, **Excel** o **PDF** per generare i documenti ufficiali.

---

## 🖥️ Funzionalità Avanzate dell'Interfaccia

* **Tendine a comparsa integrate (Dropdown)**: suggeriscono i nominativi corretti conformi all'anagrafica, eliminando alla radice errori di digitazione e conflitti di omonimia.
* **Classi Sdoppiate (Laboratori Cucina e Sala)**: riconosce e unisce automaticamente i gruppi sdoppiati di laboratorio (es. `1AB` e `1ABSALA`).
* **Menu Contestuale con Tasto Destro**:
  * *Segna come Coperta Manualmente* (sfondo verde salvia).
  * *Modifica Sostituto*: consente di forzare o correggere manualmente l'assegnazione.
  * *Rimuovi Riga* dal prospetto.
  * *Copia negli appunti*.
* **Editor Dati Integrato con Anteprima**: consente di visualizzare, correggere e salvare i dati dei 6 file direttamente dall'interfaccia senza riaprire Excel.

---

## 📄 Esportazione e Stampe Ufficiali

* **Microsoft Word (`.docx`)**:
  * Modalità *A Colori*: fedele all'interfaccia grafica per consultazione digitale.
  * Modalità *Bianco e Nero*: ottimizzata per la stampa economica con colonna firme docenti.
* **Stampa / PDF**: impaginazione formattata con intestazione istituzionale, data, giorno e prospetto ordinato.
* **Microsoft Excel (`.xlsx`)**: foglio di calcolo completo per archiviazione e statistiche d'istituto.

---

## 📁 Architettura del Progetto

```text
├── main.py                             # Entry point dell'applicazione
├── requirements.txt                    # Dipendenze Python
├── alberghiero.png / alberghiero.ico   # Icone istituzionali
├── manuale_utente.html                 # Manuale utente HTML consultabile dall'app
├── Manuale_Utente_Gestione_Sostituzioni.docx # Manuale utente in formato Word
├── avvia_gestione.command              # Script di avvio rapido per macOS
├── build_dmg.sh                        # Script di packaging macOS (DMG / PyInstaller)
├── build_windows.bat                   # Script di packaging per Windows (EXE)
│
├── gestione_sostituzioni/              # Pacchetto applicativo principale
│   ├── app.py                          # Interfaccia grafica utente Tkinter
│   ├── config.py                       # Configurazione centralizzata, colori e costanti
│   ├── data_manager.py                 # Parser CSV/Excel dei 6 riquadri
│   ├── database.py                     # Gestore persistenza SQLite3 locale
│   ├── engine.py                       # Motore algoritmico gerarchico di sostituzione
│   ├── editor.py                       # Finestra editor e anteprima dati
│   ├── exporters.py                    # Esportatori DOCX, XLSX e PDF
│   ├── utils.py                        # Normalizzazione classi, parsing ore e slug docenti
│   └── logging_config.py               # Configurazione avanzata del logging
│
├── Modello_1_Orario_Classi.xlsx        # Modello Excel Riquadro 1 (Orario Classi)
├── Modello_2_Orario_Personale.xlsx     # Modello Excel Riquadro 2 (Orario Personale)
├── Modello_3_Materia_Alternativa.xlsx  # Modello Excel Riquadro 3 (Materia Alternativa)
├── Modello_4_Potenziamento.xlsx        # Modello Excel Riquadro 4 (Potenziamento)
├── Modello_5_A_Pagamento.xlsx          # Modello Excel Riquadro 5 (Ore a Pagamento)
├── Modello_6_Docenti_Sostegno.xlsx     # Modello Excel Riquadro 6 (Docenti Sostegno)
│
└── tests/                              # Suite di test automatizzati
    ├── test_engine.py                  # Test dell'algoritmo di sostituzione
    ├── test_data_manager.py            # Test dei parser CSV ed Excel
    ├── test_database.py                # Test persistenza SQLite
    └── test_utils.py                   # Test utilità e normalizzazione
```

---

## 🚀 Installazione e Avvio

### Prerequisiti
* Python 3.10 o superiore.
* Sistema operativo: macOS o Windows.

### 1. Clonazione del repository
```bash
git clone https://github.com/TUO-USERNAME/gestione-sostituzioni-docenti.git
cd gestione-sostituzioni-docenti
```

### 2. Creazione ambiente virtuale (consigliato)
```bash
python3 -m venv venv
source venv/bin/activate       # Su macOS / Linux
# venv\Scripts\activate        # Su Windows
```

### 3. Installazione delle dipendenze
```bash
pip install -r requirements.txt
```

### 4. Avvio dell'applicazione
```bash
python3 main.py
```
*(Su macOS è possibile fare doppio clic su `avvia_gestione.command`)*

---

## 📜 Licenza
Distribuito per uso amministrativo scolastico. Tutti i diritti riservati.
