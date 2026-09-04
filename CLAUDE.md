# VakbondMailer — context voor Claude

Windows-desktopapp (WPF) waarmee een FNV-medewerker een gepersonaliseerde standaardmail stuurt
naar een lijst docenten, om gastlessen in te plannen. Verstuurd via de al ingelogde klassieke
Outlook op de eigen laptop.

## Harde regels

1. **Geen persoonsnamen in de repo.** Niet in code, commits, docs of commentaar. De naam van de
   gebruiker is eerder bewust uit de geschiedenis gehaald; schrijf "de gebruiker" of "de
   FNV-medewerker". Ook geen privé-mailadressen — de repo staat publiek.
2. **Verzenden mag alleen vanaf `@fnv.nl`.** Zit vast in `SendSettings.RequiredEmailDomain` en
   wordt afgedwongen vóór elke verzending, ook bij de testmail. Niet versoepelen.
3. **Geen Azure-app-registratie, geen opgeslagen wachtwoorden, geen hosting.** De gebruiker kan
   die niet aanmaken; dat is de reden dat dit Outlook-COM is en geen Graph API.
4. **Alles is Nederlands**: knoppen, meldingen, commentaar, commits en documentatie.
5. **Outlook-COM alleen op de UI-thread.** COM is hier STA-gebonden; geen `Task.Run` om een
   Outlook-aanroep heen.

## Bouwen, testen, draaien (vanuit WSL)

Er staat geen `dotnet` in WSL; gebruik de Windows-SDK direct:

```bash
cd /mnt/c/Users/<gebruiker>/source/repos/VakbondMailer   # de repo staat op de Windows-schijf
"/mnt/c/Program Files/dotnet/dotnet.exe" build VakbondMailer.sln
"/mnt/c/Program Files/dotnet/dotnet.exe" test src/VakbondMailer.Tests   # 84 tests
"/mnt/c/Program Files/dotnet/dotnet.exe" run --project src/VakbondMailer
```

De app zelf kan alleen echt iets doen op een Windows-sessie met Outlook open. Verzenden is dus
niet te testen vanuit deze omgeving — zeg dat eerlijk in plaats van het te suggereren.

## Verder lezen

- [`.claude/context/project.md`](.claude/context/project.md) — wat de app doet en voor wie
- [`.claude/context/architectuur.md`](.claude/context/architectuur.md) — waar welke code staat
- [`.claude/context/beslissingen.md`](.claude/context/beslissingen.md) — waarom het zo gebouwd is
- [`.claude/context/valkuilen.md`](.claude/context/valkuilen.md) — fouten die al een keer gemaakt zijn
- [`.claude/context/werkwijze.md`](.claude/context/werkwijze.md) — screenshots, CI en uitgeven
