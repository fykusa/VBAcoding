# Architektura

## Kontext

Všechny komponenty běží uvnitř jednoho Outlook VBA projektu a používají globální objekt `Application`. Není zde samostatná aplikační vrstva ani persistence řízená repozitářem. Stav se nachází v Outlook schránce, lokálních souborech automatického watcheru a v životním cyklu běžící instance Outlooku.

## Komponenty

### Diagnostika výběru

`outlook__DiagnostikaVybranych.vba` pracuje s `Application.ActiveExplorer.Selection`. Pro každou označenou položku čte identifikátor, odesílatele, předmět a čas přijetí. Výsledek posílá do Immediate okna a přes `MSForms.DataObject` do systémové schránky.

### Ruční čištění duplicit

`outlook__EmailCleaner.vba` vyhledá první store typu `olPrimaryExchangeMailbox` a vezme jeho Inbox. Položky seskupuje podle normalizované kombinace odesílatele a předmětu; časy porovnává v okně `DUPLICATE_WINDOW_MINUTES`. Identifikátory nalezených duplicit drží v paměti, zobrazí náhled ve `frmProgress` a teprve po potvrzení načte položky přes `GetItemFromID` a zavolá `Delete`.

Tok:

```text
primární Inbox
  -> průchod položkami
  -> klíč sender + subject
  -> porovnání ReceivedTime
  -> náhled ve frmProgress
  -> potvrzení uživatele
  -> Deleted Items
```

### Automatický watcher duplicit

Watcher je rozdělen mezi dva soubory:

- `outlook__ThisOutlookSession_DuplicateWatcher.vba` vlastní `WithEvents InboxItems`, reaguje na start Outlooku a událost `ItemAdd`;
- `outlook__EmailCleanerAuto.vba` provádí detekci, lokální archivaci, logování, odstranění a retenci archivu.

Při inicializaci se watcher připojí k primárnímu Inboxu, zapíše start do logu a odstraní archivované `.msg` soubory starší než nastavená retence. Při nové položce prochází nejvýše 300 nejnovějších položek seřazených podle `ReceivedTime`. Pokud najde jinou zprávu se stejným odesílatelem a předmětem v časovém okně, zpracuje nově příchozí zprávu jako duplicitu.

```text
Application_Startup / ruční InitDuplicateWatcher
  -> InboxItems.ItemAdd
  -> CheckAndRemoveDuplicateSilent
  -> hledání shody v primárním Inboxu
  -> SaveAs lokální .msg
  -> zápis do logu
  -> Delete nové položky
```

Archivace a odstranění nejsou transakční: chyba `SaveAs` se zaloguje, ale následné `Delete` není touto chybou zablokováno.

### Analýza velikosti schránky

`outlook__InboxSizeTools.vba` obsahuje dva toky nad primární schránkou:

- `ShowMailboxReport` rekurzivně projde složky, sečte velikosti položek, seřadí složky a zobrazí až 30 neprázdných výsledků;
- `MarkLargeItems` shromáždí položky a jejich velikosti, seřadí indexy a až 50 největším nastaví kategorii `! Velky mail`.

Oba toky používají `frmProgress` a během dlouhých operací volají `DoEvents`.

### Kopírování vybraných zpráv

`outlook__UsefullTools.vba` nechá uživatele vybrat cílovou Outlook složku. Pro každý označený `MailItem` vytvoří kopii a přesune tuto kopii do cíle; originál zůstává na místě.

## Sdílené hranice a vazby

- Výběr primární schránky je implementován samostatně ve třech modulech a je založen na `olPrimaryExchangeMailbox`.
- Ruční a automatická detekce používají stejný koncept duplicity, ale samostatné konstanty a samostatnou implementaci.
- `frmProgress` je sdílená UI závislost ručního čističe a nástrojů velikosti schránky.
- Automatický watcher má přímou vazbu na lokální filesystem přes archiv a log.
- Chybové stavy jsou většinou řešeny dialogem, logem nebo `On Error Resume Next`; projekt nemá centrální error-handling vrstvu.

## Záměrné hranice

- Nástroje analyzující celou schránku používají pouze primární Exchange mailbox a vynechávají sdílené stores.
- Automatický watcher zachovává starší zprávu a odstraňuje nově příchozí shodu.
- `ThisOutlookSession` je hostitelský event modul a nelze jej zaměnit za standardní modul.
- Lokální ZIP soubory v `backup/` nejsou součástí verzované architektury.
