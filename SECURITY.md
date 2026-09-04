# Beveiliging

## Een probleem melden

Zie je een beveiligingsprobleem in deze app? Meld het via
[Security > Report a vulnerability](https://github.com/lucas4790/VakbondMailer/security/advisories/new)
op deze repository. Zo blijft het bericht privé totdat het verholpen is.

Maak er alsjeblieft **geen openbaar issue** van, want daarmee staat het probleem meteen op straat.

## Hoe deze app met gegevens omgaat

Goed om te weten bij het beoordelen van een melding:

- De app draait volledig lokaal op de computer van de gebruiker. Er is geen server, geen
  account en geen internetverbinding nodig.
- Mail wordt verstuurd via de Outlook die al op die computer staat en waar de gebruiker al
  op is ingelogd. De app slaat dus **geen wachtwoorden of tokens** op en heeft geen
  Azure-app-registratie nodig.
- De ingelezen ledenlijst blijft in het werkgeheugen; die wordt nergens door de app
  weggeschreven. Het verzendrapport dat de gebruiker zelf krijgt is de enige uitzondering en
  komt naast het eigen bronbestand te staan.
- Van de verzendgeschiedenis (voor de waarschuwing bij dubbel versturen) worden alleen
  SHA-256-vingerafdrukken van e-mailadressen bewaard, niet de adressen zelf. Die staan in
  `%AppData%\VakbondMailer\` en worden na 90 dagen opgeruimd.
- Er kan alleen verstuurd worden vanaf een adres op het toegestane domein; dat wordt vlak
  vóór het versturen nog een keer gecontroleerd.

## Ondersteunde versies

De nieuwste release krijgt updates. Oudere versies niet.
