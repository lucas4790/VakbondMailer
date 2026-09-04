Verstuurt een gepersonaliseerde standaardmail naar iedereen in een CSV- of Excel-lijst, via de klassieke Outlook op je eigen computer.

## Nieuw in 1.1.0

- Nieuw uiterlijk: de knoppen, keuzelijsten en de kalender volgen nu de Windows 11-stijl, in de FNV-kleuren
- Het venster past zich aan je scherm aan, en de inhoud staat in twee kolommen naast elkaar in plaats van onder elkaar
- De stappenbalk bovenin houdt bij hoe ver je bent: per stap een vinkje en de stand van zaken
- Lege vakken leggen nu uit wat er hoort te komen, in plaats van leeg te blijven
- Draait op .NET 10, ondersteund tot november 2028 (installeren hoeft nog steeds niet)

## Installeren

1. Download de zip hieronder.
2. Pak hem uit naar een eigen map — **alle bestanden moeten bij elkaar blijven**, de .exe werkt niet los.
3. Dubbelklik `VakbondMailer.exe`. Er hoeft niets geïnstalleerd te worden.
4. Lees `LEESMIJ.txt` voor de stap-voor-stap uitleg.

## Vereisten

- Windows
- De **klassieke** Outlook-desktopapp, geopend en ingelogd (de nieuwe Outlook uit de Microsoft Store ondersteunt geen automatisering)
- Verzenden kan alleen vanaf een `@fnv.nl`-adres

## Wat zit erin

- Ledenlijst inlezen uit CSV of Excel, met controle op ongeldige en dubbele adressen
- Per ontvanger aanvinken wie de mail krijgt
- Sjablonen per onderwerp uit een eigen map
- Gastles inplannen: kies een maand en datums, die vullen `{{Maand}}`, `{{MaandJaar}}` en `{{Datumopties}}`
- Testmail vooraf, voortgang met stopknop, en "mislukte opnieuw" na afloop
- Verzendrapport (.csv) per verzending, en een waarschuwing bij dubbel versturen
