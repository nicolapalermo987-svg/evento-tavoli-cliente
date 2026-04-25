# Pubblicazione GitHub Pages con 3 link

Questa guida crea 3 URL separati:

- `/cliente/`
- `/staff/`
- `/struttura/`

## 1) Genera la cartella pronta

Esegui:

- `build-github-pages-3links.ps1`

Otterrai la cartella `web-links/`.

## 2) Pubblica `web-links/` su GitHub Pages

Metodo semplice:

1. Crea un repository dedicato (es. `evento-tavoli-web`).
2. Carica dentro **solo** i file della cartella `web-links/`.
3. In GitHub: `Settings` -> `Pages`.
4. Sorgente: branch `main`, cartella `/ (root)`.
5. Salva.

## 3) Link finali da inviare

Se il repo è `https://<utente>.github.io/evento-tavoli-web/`, i link saranno:

- `https://<utente>.github.io/evento-tavoli-web/cliente/`
- `https://<utente>.github.io/evento-tavoli-web/staff/`
- `https://<utente>.github.io/evento-tavoli-web/struttura/`

## Nota importante

Lo storage locale è separato per ruolo (chiavi diverse), quindi aprire un link non sporca i dati degli altri ruoli nello stesso browser.
