# MAUVE++ Audit & Certification Tool - Pubblica Amministrazione

Questo repository ospita il **MAUVE++ Certificatore PA**, un'applicazione sviluppata in Streamlit progettata per semplificare, strutturare e automatizzare il processo di audit di accessibilità per i siti web della Pubblica Amministrazione italiana, in conformità con le direttive **AgID** e le linee guida **WCAG 2.1 (Livello A e AA)**.

**Il tool non si sostituisce al motore di validazione originale**, ma si posiziona come strato analitico, peritale e di business intelligence a valle delle esportazioni ufficiali **MAUVE++** rilasciate dal **CNR**.

---

## 📌 Cosa fa il Tool

Il tool automatizza l'analisi quantitativa e qualitativa dei report di accessibilità, traducendo la complessità dei dati grezzi del file di log in informazioni strutturate e immediatamente leggibili:
1. **Parsa e Normalizza** i dati nativi in formato standard W3C EARL (Evaluation and Report Language).
2. **Calcola Metriche Chiave** di accessibilità e completezza, sia basate sulle singole *Tecniche* WCAG, sia sui *Criteri di Successo* complessivi richiesti dagli obblighi di legge.
3. **Genera Grafici Statistici Interattivi** (istogrammi a barre orizzontali) per mostrare l'esatta ripartizione dei dati suddivisi per i 4 Principi WCAG (Percepibile, Utilizzabile, Comprensibile, Robusto).
4. **Archivia lo Storico degli Audit** in un database locale SQLite per monitorare i progressi di conformità dell'Ente nel tempo e confrontare i diversi rilasci.
5. **Esporta Report Professionali Multiformato** (Fogli Excel XLSX formattati con tabelle separate, tracciati CSV grezzi e Relazioni PDF istituzionali).

---

## ⚙️ Come lo fa: Logica di Matching ed Esempi di Parsing JSON

Il motore dell'applicazione esamina ciclicamente l'array `@graph` del file **JSON-LD** importato, isolando esclusivamente i nodi aventi proprietà `"earl:outcome"`. Di seguito viene illustrato il flusso esatto di matching e decodifica dei dati estratti dal file sorgente.

### 1. Flusso di Estrazione dell'Asserzione Nativa
Quando il validatore MAUVE++ esporta una violazione o un avviso, genera all'interno del file una struttura dati standardizzata di questo tipo:

```json
{
  "@id": "mauve:assertion_35",
  "@type": "earl:Assertion",
  "earl:subject": {
    "@id": "mauve-earl-reporthttps___www.comune.ente.it_sito_argomenti_modulistica"
  },
  "earl:test": {
    "@id": "[http://www.w3.org/WAI/GL/WCAG20/techniques/G18](http://www.w3.org/WAI/GL/WCAG20/techniques/G18)"
  },
  "earl:result": {
    "@type": "earl:TestResult",
    "earl:outcome": {
      "@id": "[http://www.w3.org/ns/earl#failed](http://www.w3.org/ns/earl#failed)"
    }
  }
}

```

### 2. Algoritmo di Matching e Normalizzazione Dati

Il tool intercetta il nodo JSON-LD grezzo ed esegue in tempo reale un processo di decodifica, normalizzazione e arricchimento semantico strutturato secondo la seguente pipeline logica:

| Fase del Pipeline | Chiave JSON Sorgente | Processo Computazionale Applicato | Risultato Finale Ottenuto |
| --- | --- | --- | --- |
| **1. Sanificazione Perimetro** | `earl:subject -> @id` | Applicazione di espressioni regolari del tipo `re.sub(r'^mauve-earl-reporthttps___', 'https://', ...)` e sostituzione degli underscore con slash per ricostruire l'alberatura nativa dell'URL. | `https://www.comune.ente.it/sito/argomenti/modulistica` |
| **2. Estrazione Regola** | `earl:test -> @id` | Isolamento del token finale tramite ancoraggio all'ultimo carattere `/` o `#` dell'URI per identificare la tecnica atomica W3C. | Stringa univoca: `G18` |
| **3. Arricchimento Semantico** | *Dizionario Interno* | Cross-referencing e associazione automatica della tecnica estratta con il paniere di mappatura delle specifiche ufficiali AgID. | Mappatura completa: Criterio `1.4.3`, Livello `AA`, Principio `Percepibile`. |
| **4. Classificazione Esito** | `earl:result -> earl:outcome -> @id` | Intercettazione dell'ancora finale dell'URI (`#failed` o `#cannotTell`) per smistare il dato nei contatori matematici globali. | Classificato come `Failed` (+1 al computo degli errori e ricalcolo indici statistici). |

---

## 📐 Formula di Computazione e Modello di Ponderazione dei Pesi CNR (Rettificato)
Per garantire la perfetta conformità scientifica con i modelli di calcolo nativi di MAUVE++ ed evitare metriche lineari incomplete, il tool implementa la media ponderata (pesata) ufficiale stabilita dal CNR.

Ciascuna asserzione viene associata a un coefficiente di peso basato sul livello di appartenenza del controllo, in modo che la severità dell'impatto sia inversamente proporzionale alla flessibilità del livello analizzato:

Level A (Severità Massima - Blocca l'accessibilità di base): Peso = 3

Level AA (Standard Richiesto per la Pubblica Amministrazione): Peso = 2

Level AAA (Ottimizzazione Avanzata e di Dettaglio): Peso = 1

Calcolo dell'Accessibility Score Ponderato
L'indice percentuale di accessibilità (calcolato separatamente per le singole Tecniche e per i Criteri di Successo complessivi) divide la sommatoria dei successi moltiplicati per il rispettivo peso per la sommatoria dei test totali eseguiti moltiplicati per il medesimo peso. Vengono sistematicamente esclusi dal paniere i test non applicabili (Not Applicable):

Accessibility Score % = [ Somma(Peso_Livello x Successi) / Somma(Peso_Livello x Test_Totali) ] x 100

Calcolo dell'Evaluation Completeness
L'indice misura l'estensione e lo stato di avanzamento dell'audit automatico calcolando la percentuale di regole che hanno già ottenuto un esito definitivo certo (Passed o Failed) rispetto alla totalità dei test riscontrati incluse le situazioni di ambiguità semantica rappresentate dai Warning (cannotTell):

Completeness % = [ (Passed + Failed) / (Passed + Failed + Warning) ] x 100

## 🎯 Perché lo fa: Confutazione dei Falsi Positivi ed Esonero di Responsabilità

> ### ⚠️ Dichiarazione Peritale d'Infrastruttura
> 
> 
> Si attesta formalmente che i sistemi di scansione automatizzata possono incorrere in anomalie temporanee di tokenizzazione della struttura DOM (DOM-Tree), spesso causate dal caricamento asincrono di script o fogli di stile esterni durante la sessione di scraping.
> Questo tool è stato ingegnerizzato specificamente per **intercettare e neutralizzare l'impatto di tali falsi positivi**, proteggendo l'operato tecnico e la reputazione dei fornitori tecnologici della Pubblica Amministrazione.

L'algoritmo del tool analizza e risolve questa criticità strutturale attraverso una logica di isolamento matematico:

* **Isolamento dell'Asimmetria:** Se un errore sistematico legato ad elementi comuni (come l'Header o il Footer) viene rilevato da MAUVE++ esclusivamente su una singola pagina e convalidato come pienamente conforme sulle restanti risorse esaminate (nonostante il codice del template sia identico), il sistema identifica l'incoerenza analitica dello strumento del CNR.
* **Generazione del Focus Ispettivo (Capitolo 1.1):** Il tool ordina la lista delle pagine in modo rigidamente decrescente per numero di errori. Isola l'URL anomalo e inserisce automaticamente nella relazione PDF una nota tecnica formale che attribuisce il picco ad un'incoerenza temporanea del parser automatico e non a un difetto dell'Ente, suggerendo l'apertura di un ticket di manutenzione presso i laboratori del CNR.
* **Normalizzazione Cromatica AgID:** Il tool ricollega i macro-indicatori di accessibilità e completezza alle fasce di voto reali dell'Ente. La palette cromatica del PDF e della dashboard si adatta dinamicamente passando al Verde (maggiore o uguale al 90%), al Giallo/Arancio (maggiore o uguale al 75%) o al Rosso (minore del 75%), escludendo colorazioni fisse che potrebbero trarre in inganno l'RTD.

---

## 🛠️ Come funziona il Tool (Guida all'Uso Rapido)

### 📋 Flusso Operativo Sequenziale

```
[1. Configurazione Ente] ➔ [2. Upload Drag-and-Drop] ➔ [3. Elaborazione e SQLite] ➔ [4. Analisi e Download]

```

#### 🛠️ Step 1: Configurazione ed Input dell'Ente

Nel pannello laterale sinistro (**Sidebar**), inserisci il nome dell'Amministrazione sotto esame (Soggetto Erogatore PA) e imposta la data di generazione dell'audit. I campi di testo dispongono di un sistema di persistenza dello stato per evitare la perdita dei dati durante l'elaborazione.

#### 📂 Step 2: Caricamento Massivo dei Log

Trascina e rilascia l'intero set di file **JSON-LD** esportati da MAUVE++ all'interno del modulo di upload dedicato. Il sistema supporta il caricamento simultaneo di decine di file relativi all'intero campionamento delle pagine del sito.

#### 💾 Step 3: Elaborazione e Persistenza Strutturata

Fai clic sul pulsante **"💾 Elabora e Salva in Database"**. L'applicazione eseguirà istantaneamente la pipeline di parsing, estrarrà le metriche reali, popolerà lo storico all'interno del database locale **SQLite** e aggiornerà i grafici statistici interattivi.

#### 📊 Step 4: Analisi dei Dati ed Esportazione Multiformato

Esamina i risultati aggregati direttamente all'interno della sezione *Dashboard Audit Analizzati*. In corrispondenza della riga dell'audit appena salvato nella tabella storica, fai clic sulle icone di download per esportare:

* 📄 **File CSV:** Tracciato grezzo e lineare delle violazioni per l'importazione in sistemi terzi.
* 📊 **File Excel (XLSX):** Cartella di lavoro formattata professionalmente con gridlines abilitate e fogli separati (*Sommario, Registro Violazioni, Pagine Campionate*).
* 📕 **Relazione PDF:** Documento ufficiale impaginato in formato letter con grafici ad alta risoluzione integrati, indice di conformità reale e nota ispettiva sui falsi positivi inclusa.

---

## 💻 Installazione e Avvio

Assicurati di avere installato Python (versione >= 3.8). Installa le librerie richieste tramite pip:

```bash
pip install streamlit pandas matplotlib openpyxl reportlab
```

Avvia l'applicazione localmente sul tuo browser:

```bash
streamlit run app.py


```
