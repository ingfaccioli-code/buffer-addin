# Buffer Calendario – Outlook Add-in

Conversione della macro VBA `BlockOffTime` per il nuovo Outlook.

## Cosa fa

Quando sei su un appuntamento del calendario, l'add-in crea automaticamente due appuntamenti "buffer" (es. *Tempo di viaggio*):
- uno **prima** dell'evento
- uno **dopo** l'evento

Entrambi vengono impostati come **Fuori sede** e con **categoria gialla**, esattamente come nella macro originale.

---

## File inclusi

| File | Descrizione |
|------|-------------|
| `manifest.xml` | Definisce l'add-in per Outlook |
| `taskpane.html` | Interfaccia utente + logica JavaScript |

---

## Come installare (due metodi)

### Metodo A – Caricamento manuale (test personale, senza server)

Questo metodo richiede che i file siano ospitati su un server HTTPS.
Puoi usare **GitHub Pages** (gratuito):

1. Crea un repository GitHub e carica `taskpane.html`
2. Attiva GitHub Pages nelle impostazioni del repository
3. Modifica `manifest.xml` → sostituisci `https://TUODOMINIO/buffer-addin/taskpane.html` con l'URL GitHub Pages (es. `https://tuonome.github.io/buffer-addin/taskpane.html`)
4. In Outlook (web o desktop nuovo):
   - Vai su **App** (icona a puzzle in basso a sinistra)
   - Clicca **Gestisci app** → **Aggiungi app personalizzata**
   - Carica il file `manifest.xml`

### Metodo B – Distribuzione aziendale (via amministratore Microsoft 365)

Se usi Outlook in un'organizzazione:

1. L'amministratore M365 va su **Interfaccia di amministrazione Microsoft 365**
2. Sezione **Impostazioni** → **App integrate** → **Carica app personalizzata**
3. Carica `manifest.xml`
4. L'add-in sarà disponibile per tutti gli utenti

---

## Come usare

1. Apri un **appuntamento del calendario** (in visualizzazione lettura)
2. Clicca sull'icona **Buffer Calendario** nella barra degli strumenti
3. Si apre il pannello laterale con le informazioni dell'evento
4. Modifica il nome del buffer e la durata (default: 60 minuti)
5. Clicca **Crea buffer prima e dopo**

---

## Requisiti

- Account **Microsoft 365** o **Exchange** (necessario per EWS)
- Nuovo Outlook per Windows / Outlook Web / Outlook Mac
- Autorizzazione **ReadWriteMailbox**

---

## Note tecniche

La macro originale usava `Application.CreateItem` e `AppointmentItem` tramite VBA.
Questo add-in usa le **Exchange Web Services (EWS)** tramite `Office.context.mailbox.makeEwsRequestAsync`,
che è l'equivalente moderno supportato nel nuovo Outlook.

La categoria "Categoria giallo" deve già esistere nel tuo account Outlook affinché venga applicata correttamente.
