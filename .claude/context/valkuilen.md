# Valkuilen die al een keer geraakt zijn

Kost je anders opnieuw een half uur.

## `Marshal.GetActiveObject` bestaat niet op moderne .NET

Die zat in .NET Framework, niet in .NET (Core) en hoger. Een draaiende Outlook overnemen gaat via
een eigen P/Invoke naar `ole32!CLSIDFromProgID` plus `oleaut32!GetActiveObject`. Zet daar
`[DefaultDllImportSearchPaths(DllImportSearchPath.System32)]` op — zonder dat klaagt de analyzer
terecht over DLL-hijacking (CA5392).

## `InvariantGlobalization=true` laat de app crashen bij het opstarten

Leek onschuldig voor een kleinere .exe, maar WPF-binding valt er direct over:
*"Cannot find non-neutral culture related to 'en-us'"*. Niet terugzetten.

## Een eigen `<Style TargetType="ComboBox">` sloopt de Fluent-theme

Een impliciete stijl zónder `BasedOn` **vervangt** de template van de theme, waarna de oude
Aero-chrome terugkomt — ook als de stijl alleen een lettergrootte zet. Gebruik altijd een keyed
stijl met `BasedOn="{StaticResource {x:Type ComboBox}}"`. Dit was precies de oorzaak van "het ziet
er half af uit".

## `dotnet test` en xunit.v3 op de .NET 10 SDK

De oude VSTest-route werkt niet meer. De oplossing staat in `global.json`:

```json
"test": { "runner": "Microsoft.Testing.Platform" }
```

De property `TestingPlatformDotnetTestSupport` in het testproject is **niet** genoeg; dat is een
keer geprobeerd en de CI bleef rood.

## Het testproject moet hetzelfde TFM hebben

`net10.0` kan niet naar `net10.0-windows` verwijzen (NU1201). Beide projecten staan daarom op
`net10.0-windows`.

## Verzendrapport zonder BOM leest Excel verkeerd

`SendReportService` schrijft UTF-8 **met** BOM. Zonder BOM maakt Excel er `Ren�` van.

## Screenshots blijven zwart

Op een machine zonder actieve schermsessie levert een opname van een GPU-gerenderd WPF-venster een
leeg beeld. Start de app daarom met `VAKBONDMAILER_SOFTWARE_RENDERING=1` en gebruik
`scripts/Maak-Screenshot.ps1` (die gebruikt `PrintWindow` met `PW_RENDERFULLCONTENT`). Verwacht
kleine verschillen in de bovenste ~30 pixels tussen twee opnames: dat is de titelbalk die actief
of inactief is, geen wijziging in de app.

## Impliciete usings dekken `System.IO` niet

In een WPF-project zit `System.IO` niet in de impliciete usings. Dat kostte twee compileerfouten in
services die met bestanden werken.

## `dependency-review.yml` is niet handmatig te starten

Die workflow heeft alleen een `pull_request`-trigger, en dat is zo bedoeld: hij vergelijkt de
afhankelijkheden die een PR *toevoegt*. Zonder PR-diff heeft hij niets te doen. Meld dat eerlijk in
plaats van te zeggen dat alle workflows gedraaid zijn.
