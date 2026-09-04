# Vakbond Mailer

Kleine Windows-app om een gepersonaliseerde standaardmail te versturen naar iedereen in een
CSV- of Excel-lijst, via je eigen (al ingelogde) klassieke Outlook. Geen Azure/app-registration,
geen wachtwoorden in de app, geen hosting — alles blijft op je eigen laptop.

## Vereisten

- Windows, met de **klassieke Outlook desktop-app** geïnstalleerd en ingelogd
  (niet de nieuwe "New Outlook"-toggle — die ondersteunt geen automatisering)
- .NET 8 SDK om te bouwen: https://dotnet.microsoft.com/download

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

## Publiceren als los .exe-bestand (voor iemand zonder .NET geïnstalleerd)

```powershell
dotnet publish src\VakbondMailer -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

De .exe staat daarna in `src\VakbondMailer\bin\Release\net8.0-windows\win-x64\publish\`.
