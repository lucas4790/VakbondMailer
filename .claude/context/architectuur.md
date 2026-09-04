# Waar welke code staat

Twee projecten in `VakbondMailer.sln`, beide `net10.0-windows`:

- `src/VakbondMailer` — de WPF-app
- `src/VakbondMailer.Tests` — xunit.v3, 84 tests

## De hoofdlijn

**Alle logica die de moeite van het testen waard is, staat in `Services/` en kent het scherm niet.**
De code-behind eromheen leest velden uit, roept een service aan en zet het antwoord terug in beeld.
Dat is bewust zo gesplitst: een WPF-venster is niet te testen vanuit deze omgeving, een service wel.

## Het scherm

`MainWindow` is één klasse, verdeeld over vier bestanden — één per stap van het scherm, zodat je
niet door 600 regels scrolt om de verzendlus te vinden:

| Bestand | Waarover |
|---|---|
| `MainWindow.xaml` | de hele indeling: drie kaarten, elk in twee kolommen |
| `MainWindow.xaml.cs` | velden, constructor, `SizeToScreen()`, de stappenbalk bovenin |
| `MainWindow.Ledenlijst.cs` | stap 1: bestand kiezen, kolommen, selectie |
| `MainWindow.Bericht.cs` | stap 2: sjablonen, velden invoegen, planning, live voorbeeld |
| `MainWindow.Versturen.cs` | stap 3: accounts, testmail, verzendlus, rapport |

`App.xaml` bevat het stijlwoordenboek en zet `ThemeMode="Light"` (de ingebouwde Fluent-theme van
WPF). `App.xaml.cs` pint de taal op nl-NL, zet de accentkleur op FNV-blauw, en schakelt
software-rendering in als `VAKBONDMAILER_SOFTWARE_RENDERING=1` (voor screenshots).

## Services

| Service | Verantwoordelijkheid |
|---|---|
| `RecipientImportService` | CSV/Excel inlezen, e-mailkolom raden, ongeldige en dubbele adressen melden |
| `RecipientSelection` | de lijst zoals die in beeld staat plus wie aangevinkt is; rij-index = ontvanger-index |
| `TemplateRenderer` | `{{velden}}` invullen, en onbekende velden opsporen |
| `TemplateStorageService` / `TemplateLibraryService` | één sjabloon lezen/schrijven / een map met sjablonen tonen |
| `PlanningFields` | `{{Maand}}`, `{{MaandJaar}}`, `{{Datumopties}}` opbouwen uit de gekozen maand en datums |
| `SendSettings` | de regels: alleen `@fnv.nl`, pauze tussen mails, venster voor dubbel-versturen |
| `IMailSender` / `OutlookMailService` | het versturen zelf; de interface bestaat zodat de verzendlus getest kan worden |
| `BulkMailSender` | de verzendlus: één mail per ontvanger, pauze, voortgang, wat mislukte |
| `SendLog` | het logboek dat tijdens het versturen meeloopt |
| `SendReportService` | het `verzendrapport_*.csv` |
| `SendHistoryService` | gehashte verzendgeschiedenis, voor de waarschuwing bij dubbel versturen |
| `AppSettingsService` | kleine instellingen in AppData (de sjablonenmap) |
| `SimpleHtmlFormatter` | `**vet**` en kale links omzetten naar minimale HTML-mail |

## Outlook

`OutlookMailService` praat **late-bound** met Outlook: `Type.GetTypeFromProgID("Outlook.Application")`
plus `dynamic`, zodat er geen Interop-assembly of geïnstalleerde SDK nodig is. Een al draaiende
Outlook wordt overgenomen via een eigen P/Invoke naar `ole32!CLSIDFromProgID` en
`oleaut32!GetActiveObject` — `Marshal.GetActiveObject` bestaat niet op moderne .NET.

Gebruikte Outlook-API: `Session.Accounts` (voor de accountkeuze), `MailItem.SendUsingAccount`,
`Subject`, `Body`/`HTMLBody`, `Attachments.Add`, `Send()`.

## Modellen

`Recipient` (adres plus alle kolommen van die rij) en `SendResult` (per ontvanger: gelukt, en zo
niet, waarom).
