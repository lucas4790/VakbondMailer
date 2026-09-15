# Waarom het zo gebouwd is

Beslissingen met hun reden, zodat ze niet per ongeluk teruggedraaid worden.

## Een desktop-app, geen webapp

Het eerste plan was een gehoste webpagina met Microsoft 365. Dat sneuvelde op één zin van de
opdrachtgever: er komt geen app-registratie. Zonder app-registratie geen Graph API, geen OAuth,
geen hosting die namens iemand mag mailen. Een desktop-app die de al ingelogde Outlook aanstuurt,
heeft niets van dat alles nodig — de gebruiker is al ingelogd, de app leent alleen die sessie.

Bijvangst: er staan geen wachtwoorden in de app, er gaat niets naar een server, en er is niets te
beheren.

## Klassieke Outlook, en dat is een echte beperking

COM-automatisering werkt alleen met de **klassieke** Outlook-desktopapp. De "nieuwe Outlook" uit
de Microsoft Store ondersteunt het niet. Dat staat in `LEESMIJ.txt` en in de release-notities,
want het is de meest waarschijnlijke reden dat de app bij iemand niet werkt.

## Alleen `@fnv.nl` versturen

Op verzoek: mail namens de vakbond mag niet per ongeluk vanaf een privéadres vertrekken. Het
domein staat als constante in `SendSettings` en wordt vóór elke verzending gecontroleerd, ook bij
de testmail. Staat er geen `@fnv.nl`-account in Outlook, dan blokkeert de app het versturen in
plaats van stilletjes het standaardaccount te pakken.

## .NET 10 in plaats van .NET 8

.NET 8 loopt **10 november 2026** uit ondersteuning. .NET 10 is de huidige LTS, ondersteund tot
november 2028. De upgrade bracht meteen de oplossing voor het ontwerpprobleem mee (zie hierna).

## De ingebouwde Fluent-theme, geen externe UI-bibliotheek

De app zag er half af uit: eigen stijlen voor knoppen en kaarten, maar de ComboBoxen en de
kalender stonden nog in Aero-chrome uit Windows 7. Overwogen was WPF-UI (externe bibliotheek).
Sinds .NET 9 heeft WPF echter een **eigen** Fluent-theme, in .NET 10 verder aangevuld:
`ThemeMode="Light"` op `Application`, en de standaardbesturingselementen zien er meteen uit als
Windows 11. Dat is minder afhankelijkheid en minder onderhoud dan een bibliotheek erbij.

De accentkleur wordt in `App.xaml.cs` overschreven met FNV-blauw, want anders volgt de theme de
Windows-accentkleur van de computer waar de app toevallig op draait.

## FNV-kleuren, geverifieerd

Blauw `#009CDE` en groen `#7FBA24` komen uit het logo op fnv.nl zelf, niet uit een schatting.
`GoodBrush` en `BadBrush` zijn aparte kleuren voor "gelukt"/"mislukt", zodat status niet met het
accent verward wordt.

## Adressen in de geschiedenis worden gehasht

De waarschuwing "deze mailing ging kort geleden al naar deze mensen" heeft alleen herkenning
nodig, geen leesbare adressen. SHA-256 dus: de app laat geen ledenlijst achter op schijf.

## Geen persoonsnamen in de repo

De repo staat publiek. De naam van de gebruiker is op verzoek uit code én geschiedenis verwijderd
(`git filter-branch`), en het privé-mailadres in de commits is vervangen door het
GitHub-noreply-adres. Introduceer geen nieuwe namen.

## De verzendlus is losgetrokken van het scherm

`BulkMailSender` krijgt een `IMailSender` en een lijst ontvangers, en weet niets van WPF. Daardoor
is het stuk waar het echt mis kan gaan — half verstuurd, netwerk weg, iemand dubbel — te testen
zonder dat er mail de deur uitgaat.

## Rij-index = ontvanger-index

Sorteren in de tabel staat **uit**. De koppeling tussen de rij die je aanklikt en de ontvanger
waarvan je een voorbeeld ziet, loopt via de index. Daar is eerder een fout in geslopen: na
sorteren werd de verkeerde persoon getoond en getest. Zet sorteren niet terug aan zonder die
koppeling eerst op identiteit te leggen in plaats van op volgorde.

## Zelfstandige .exe

`Maak-Uitgave.ps1` publiceert `--self-contained` als single file, zodat de gebruiker niets hoeft
te installeren. De prijs is een zip van ~60 MB; dat is het waard.
