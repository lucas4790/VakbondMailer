# Vakbond Mailer

Kleine Windows-app om een gepersonaliseerde standaardmail te versturen naar iedereen in een
CSV- of Excel-lijst, via je eigen (al ingelogde) klassieke Outlook. Geen Azure/app-registration,
geen wachtwoorden in de app, geen hosting — alles blijft op je eigen laptop.

## Vereisten

- Windows, met de **klassieke Outlook desktop-app** geïnstalleerd en ingelogd
  (niet de nieuwe "New Outlook"-toggle — die ondersteunt geen automatisering)
- .NET 10 SDK om te bouwen: https://dotnet.microsoft.com/download

## Bouwen en starten (vanaf Windows, PowerShell)

```powershell
cd VakbondMailer
dotnet build
dotnet run --project src\VakbondMailer
```

## Tests draaien

```powershell
dotnet test
```

## Gebruik

1. **Bestand kiezen**: kies een `.csv` of `.xlsx` met in elk geval een kolom met e-mailadressen.
   De app probeert de juiste kolom automatisch te herkennen; corrigeer zo nodig.
2. **Standaardmail opstellen**: typ onderwerp en tekst, klik op een veld-chip om een
   `{{Kolomnaam}}`-placeholder in te voegen (bv. `{{Voornaam}}`). Onderaan zie je een live
   voorbeeld op basis van de eerste rij in de lijst.
   - **Sjablonenmap**: kies één keer een map (bv. `templates-voorbeeld` in dit project, of een
     eigen map) met standaardmails over verschillende onderwerpen. Ze verschijnen in de
     dropdown ernaast; kiezen vult onderwerp en tekst meteen in. "Opslaan als nieuw sjabloon..."
     zet de huidige mail als nieuw bestand in die map, zodat hij de volgende keer ook in de lijst
     staat. De gekozen map wordt onthouden voor de volgende keer dat je de app opent.
3. **Stuur testmail naar mezelf**: verstuurt de gepersonaliseerde mail (op basis van de eerste
   rij) naar je eigen adres, zodat je het resultaat kunt checken voordat je verder gaat.
4. **Verstuur naar iedereen**: na bevestiging worden de mails één voor één verstuurd via Outlook.
   Na afloop staat er een `verzendrapport_*.csv` naast je brondbestand met per ontvanger de status.

## Uitgeven

Een nieuwe versie uitbrengen gaat via een tag; GitHub Actions bouwt en publiceert dan de release:

```powershell
git tag v1.0.2
git push origin v1.0.2
```

Zelf een uitgave-map maken (zelfstandige .exe plus handleiding, sjablonen en voorbeeldlijst):

```powershell
.\scripts\Maak-Uitgave.ps1
```

Let op: de .exe werkt niet los — alle bestanden uit die map moeten bij elkaar blijven.

## Scripts

| Script | Waarvoor |
|---|---|
| `scripts\Maak-Uitgave.ps1` | Bouwt de uitgave-map, optioneel meteen als zip (`-ZipPath`). Wordt ook door de workflow gebruikt. |
| `scripts\Maak-Screenshot.ps1` | Legt het venster vast, om een wijziging visueel te controleren. Start de app met `VAKBONDMAILER_SOFTWARE_RENDERING=1` als een gewone schermopname zwart blijft (machine zonder actieve schermsessie). |

Het app-icoon `src\VakbondMailer\app.ico` is eenmalig gegenereerd en staat in de repo; wil je het
vervangen, zet er dan gewoon een ander `.ico`-bestand neer.
