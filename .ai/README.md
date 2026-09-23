# Outlook VBA utilities

Sbírka samostatných maker pro Microsoft Outlook na Windows. Nástroje pomáhají diagnostikovat vybrané zprávy, analyzovat velikost schránky, označovat velké položky a hledat nebo automaticky odstraňovat duplicitní e-maily.

Kód je určen pro vestavěný Outlook VBA projekt, nikoli pro samostatné spuštění z příkazové řádky. Některé operace mění obsah schránky: přesouvají zprávy do Odstraněných položek, zapisují kategorii nebo vytvářejí lokální archiv `.msg` souborů.

## Přehled nástrojů

| Soubor | Hlavní vstup | Účel |
| --- | --- | --- |
| `outlook__DiagnostikaVybranych.vba` | `DiagnostikaVybranych` | Vypíše metadata označených zpráv a pokusí se je vložit do schránky. |
| `outlook__EmailCleaner.vba` | `RemoveInboxDuplicates` | Najde duplicity v primárním Inboxu, zobrazí náhled, po potvrzení je přesune do Odstraněných položek a výsledek zapíše do sdíleného logu. |
| `outlook__EmailCleanerAuto.vba` | `CheckAndRemoveDuplicateSilent` | Obsluhuje automatickou kontrolu nově příchozích zpráv, archivaci, mazání a logování. |
| `outlook__ThisOutlookSession_DuplicateWatcher.vba` | `Application_Startup`, `InitDuplicateWatcher` | Napojí automatickou kontrolu na událost `InboxItems.ItemAdd`. |
| `outlook__InboxSizeTools.vba` | `ShowMailboxReport`, `MarkLargeItems` | Analyzuje velikosti složek a označí 50 největších položek kategorií `! Velky mail`. |
| `outlook__UsefullTools.vba` | `CopySelectedMailToChosenFolder` | Zkopíruje označené e-maily do uživatelem zvolené složky. |

Podrobnější mapa je v [onboarding.md](onboarding.md) a vazby mezi komponentami v [architecture.md](architecture.md).

## Požadavky

- klasický Microsoft Outlook pro Windows s dostupným VBA editorem;
- povolené spouštění maker podle firemních bezpečnostních pravidel;
- přístup k primární Exchange schránce pro nástroje, které ji vyhledávají přes `olPrimaryExchangeMailbox`;
- UserForm `frmProgress` pro `EmailCleaner` a `InboxSizeTools`.
- Modul `outlook__EmailCleanerAuto.vba` pro sdílené logování výsledků ručního čištění.

Formulář je exportován jako `frmProgress.frm` s binárním doprovodem `frmProgress.frx`. Očekávané prvky jsou:

- `lblStatus` – stavový text;
- `lblCounter` – počitadlo;
- `txtResults` – víceřádkový výstup;
- `btnClose` – tlačítko pro zavření formuláře.

## Instalace

1. V Outlooku otevři VBA editor pomocí `Alt+F11`.
2. Obsah požadovaných `outlook__*.vba` souborů vlož do standardních modulů VBA projektu.
3. Kód z `outlook__ThisOutlookSession_DuplicateWatcher.vba` přidej do již existujícího modulu `ThisOutlookSession`; nevytvářej pro něj standardní modul.
4. Pokud používáš `EmailCleaner` nebo `InboxSizeTools`, vytvoř či importuj `frmProgress` s prvky uvedenými výše.
5. Před použitím automatického watcheru uprav v `outlook__EmailCleanerAuto.vba` konstanty `LOG_FILE_PATH` a `ARCHIVE_FOLDER_PATH` pro svůj počítač.
6. Pro automatický watcher restartuj Outlook nebo ručně spusť `InitDuplicateWatcher`.

## Spuštění a ověření

Ve VBA editoru lze ručně spustitelné procedury vyvolat pomocí `Alt+F8`. Projekt nemá automatizované testy ani build skript. Změny proto ověřuj nejprve na testovacích zprávách nebo v neprodukčním profilu:

1. ve VBA editoru spusť `Debug > Compile VBAProject`;
2. otestuj příslušné makro na malé, známé sadě zpráv;
3. u mazacích operací zkontroluj Odstraněné položky, lokální archiv a log;
4. u watcheru ověř v Immediate okně (`Ctrl+G`) zprávu o aktivaci a zkontroluj logovací soubor.

## Bezpečnostní upozornění

- `RemoveInboxDuplicates` přesouvá nalezené položky až po potvrzení uživatelem.
- Automatický watcher pracuje bez dialogu. Novou duplicitu se pokusí uložit jako `.msg` a následně volá `Delete`.
- Současná implementace automatického watcheru pokračuje v odstranění zprávy i v případě, že `SaveAs` selže; chyba archivace se pouze zapíše do logu.
- `PurgeOldArchivedEmails` trvale maže lokální `.msg` soubory starší než `ARCHIVE_RETENTION_DAYS`.
- `VERBOSE_LOGGING` je aktuálně nastaveno na `True`, takže loguje každé spuštění kontroly nové položky.
