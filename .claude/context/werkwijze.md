# Werkwijze: screenshots, CI en uitgeven

## Vanuit WSL werken

De repo staat op de Windows-schijf, onder `source/repos/VakbondMailer` in het gebruikersprofiel,
en wordt gebouwd met de Windows-SDK; in WSL zelf staat geen `dotnet`. Zie `CLAUDE.md` voor de
commando's.
PowerShell-scripts draaien via
`/mnt/c/WINDOWS/System32/WindowsPowerShell/v1.0/powershell.exe -File <windowspad>`.

## Een wijziging in het scherm controleren

Het venster is niet te bedienen vanuit deze omgeving — een dropdown uitklappen of een datum
aanklikken kan niet. Wat wél kan is een stilstaande opname:

1. app starten met `VAKBONDMAILER_SOFTWARE_RENDERING=1`
2. `scripts/Maak-Screenshot.ps1` draaien
3. de PNG bekijken, en bij een refactor die niets mocht veranderen de SHA-256 vergelijken met de
   opname van ervoor

Wees eerlijk over die grens: "zo ziet het er stilstaand uit" is iets anders dan "het werkt".

## Tests

84 tests, allemaal over `Services/`. Ze raken geen UI en geen Outlook. Een wijziging in de
verzendregels, het inlezen of het invullen van velden hoort een test te krijgen; een wijziging in
XAML niet.

## CI

Drie workflows, in het Nederlands benoemd in de Actions-tab:

| Workflow | Trigger | Doet |
|---|---|---|
| Bouwen en uitgeven | push, PR, handmatig, tag | bouwt, test, maakt de zip; publiceert bij een `v*`-tag de release |
| Code scannen op kwetsbaarheden | push, PR, wekelijks, handmatig | CodeQL op C# (manual build mode) |
| Afhankelijkheden controleren | alleen PR | dependency-review op wat een PR toevoegt |

Verder: Dependabot voor de GitHub Actions, en secret scanning met push protection aan.

## Een versie uitbrengen

De release-notities staan in `.github/RELEASE_NOTES.md` — werk daar het blok "Nieuw in x.y.z" bij
vóór het taggen, want de workflow gebruikt dat bestand ongewijzigd.

```bash
git tag -a v1.2.0 -m "VakbondMailer v1.2.0"
git push origin v1.2.0
```

De workflow bouwt, test, zipt en maakt de release aan met de zip eraan. De deellink voor de
gebruiker blijft altijd `https://github.com/lucas4790/VakbondMailer/releases/latest`.

Zelf een uitgave-map maken (zonder release): `scripts/Maak-Uitgave.ps1`, eventueel met `-ZipPath`.
