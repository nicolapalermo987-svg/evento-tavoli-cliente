# Distribuzione Versione A (ruolo fisso)

Questa versione blocca il ruolo a livello di pacchetto distribuito.

## File usati

- `role-config.js` -> file letto dall'app in produzione
- `role-config.cliente.js` -> preset Cliente
- `role-config.staff.js` -> preset Staff
- `role-config.struttura.js` -> preset Struttura

## Come preparare un pacchetto

1. Scegli il ruolo (`cliente`, `staff`, `struttura`).
2. Copia il preset corrispondente dentro `role-config.js`.
3. Distribuisci il software.

Esempio:

- per la versione Cliente, copia il contenuto di `role-config.cliente.js` in `role-config.js`.

## Automazione pacchetti (consigliato)

Puoi creare automaticamente i 3 pacchetti già pronti:

- `dist-cliente`
- `dist-staff`
- `dist-struttura`

Esegui uno di questi file dalla cartella progetto:

- `build-distribuzioni.bat` (doppio click su Windows)
- `build-distribuzioni.ps1` (PowerShell)

## Nota

Con questa strategia l'utente finale non ha il selettore ruolo nell'interfaccia.
