# Onboarding

## Co je v repozitáři

Projekt obsahuje šest textových exportů či kopií Outlook VBA kódu. Nemá vlastní sestavení, správce závislostí ani automatizované testy. Kód se instaluje do VBA projektu klasického Outlooku a pracuje přímo s Outlook Object Model.

Repozitář je aktuálně rozdělen podle jednotlivých funkcí, nikoli podle sdílených knihoven:

```text
outlook__DiagnostikaVybranych.vba
outlook__EmailCleaner.vba
outlook__EmailCleanerAuto.vba
outlook__InboxSizeTools.vba
outlook__ThisOutlookSession_DuplicateWatcher.vba
outlook__UsefullTools.vba
```

Adresář `backup/` je ignorovaný Gitem a obsahuje lokální ZIP zálohy. Není zdrojem pravdy pro další vývoj.

## Kde začít číst

- Pro nejjednodušší příklad práce s výběrem zpráv začni v `outlook__DiagnostikaVybranych.vba`.
- Pro potvrzované dávkové mazání pokračuj do `outlook__EmailCleaner.vba`.
- Automatický tok sleduj od `Application_Startup` v `outlook__ThisOutlookSession_DuplicateWatcher.vba` do `CheckAndRemoveDuplicateSilent` a `ArchiveAndDelete` v `outlook__EmailCleanerAuto.vba`.
- Analýza celé primární schránky a rekurzivní průchod složkami jsou v `outlook__InboxSizeTools.vba`.
- Přehled komponent a toků je v [architecture.md](architecture.md).

## Veřejné procedury

| Procedura | Chování a vedlejší efekty |
| --- | --- |
| `DiagnostikaVybranych` | Čte metadata vybraných položek, zapisuje do Immediate okna a schránky. |
| `RemoveInboxDuplicates` | Prochází primární Inbox a po náhledu a potvrzení volá `Delete` na duplicitách. |
| `ShowMailboxReport` | Rekurzivně čte položky primární schránky a zobrazí souhrn ve `frmProgress`. |
| `MarkLargeItems` | Prochází primární schránku, nastaví kategorii 50 největším položkám a uloží je. |
| `CopySelectedMailToChosenFolder` | Vytvoří kopie vybraných zpráv a přesune kopie do zvolené složky. |
| `InitDuplicateWatcher` | Připojí obsluhu `ItemAdd`, zaloguje start a spustí úklid lokálního archivu. |
| `PurgeOldArchivedEmails` | Trvale odstraní staré `.msg` soubory z lokálního archivu. |
| `OpenDuplicateLog` | Otevře nakonfigurovaný log v Poznámkovém bloku. |

`CheckAndRemoveDuplicateSilent` a `LogWatcherStart` jsou veřejné kvůli volání mezi moduly, ale běžně je spouští watcher.

## Vývojový postup

1. Uprav zdrojový `.vba` soubor v repozitáři.
2. Přenes změnu do odpovídajícího modulu ve VBA projektu Outlooku. U `ThisOutlookSession` změnu sluč s existujícím obsahem.
3. Spusť `Debug > Compile VBAProject`.
4. Otestuj pouze dotčený tok na kontrolovaných zprávách.
5. U operací se zápisem ověř skutečné vedlejší efekty: Odstraněné položky, kategorii, vytvořené `.msg` soubory a log.
6. Přenes případné opravy zpět do repozitáře, aby se zdroj a Outlook projekt nerozešly.

## Konfigurace a závislosti

- `DUPLICATE_WINDOW_MINUTES` řídí ruční detekci duplicit; automatická varianta používá samostatnou konstantu `AUTO_DUPLICATE_WINDOW_MINUTES`.
- `LOG_FILE_PATH`, `ARCHIVE_FOLDER_PATH` a `ARCHIVE_RETENTION_DAYS` jsou lokální konfigurace automatického watcheru.
- `VERBOSE_LOGGING=True` vytváří podrobný provozní log.
- `frmProgress.frm` a `frmProgress.frx` tvoří exportovaný UserForm a musí zůstat pohromadě. Formulář poskytuje prvky `lblStatus`, `lblCounter`, `txtResults` a `btnClose` používané VBA moduly.
- Diagnostika používá late-bound `MSForms.DataObject`; pokud není dostupný, výstup zůstane alespoň v Immediate okně a zobrazí se v dialogu.

## Rizika a otevřené otázky

- Automatická archivace nekontroluje úspěšnost `SaveAs` před odstraněním zprávy. Před produkčním nasazením je vhodné rozhodnout, zda má selhání archivace mazání zablokovat.
- Cesty pro log a archiv jsou v kódu svázány s konkrétním uživatelským profilem a před použitím na jiném počítači se musí změnit.
- Není doloženo, zda se soubory do Outlooku importují přímo, nebo se jejich obsah ručně kopíruje. Komentáře v kódu popisují ruční vložení.
- Neexistují automatizované testy pro hraniční případy detekce duplicit ani pro chyby Outlook Object Modelu.

## Doporučený první úkol

Uprav automatický watcher tak, aby neodstranil zprávu po neúspěšné archivaci, a ověř toto chování na kontrolované duplicitní zprávě. Exportovaný pár `frmProgress.frm` a `frmProgress.frx` už odstraňuje dřívější mezeru v reprodukovatelnosti UI.
