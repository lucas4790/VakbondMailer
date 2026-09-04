# Wat de app doet en voor wie

## De gebruiker

Eén persoon: een medewerker van de FNV-vakbond. Geen ontwikkelaar, geen beheerrechten op haar
werklaptop, geen mogelijkheid om iets in Azure aan te maken. Ze krijgt de app als een map met een
zip erin, pakt die uit en dubbelklikt de `.exe` — installeren hoort niet nodig te zijn.

De opdrachtgever is haar partner (de eigenaar van deze repo), die de app laat bouwen en de
uitgaven publiceert.

## Het werk dat de app vervangt

Ze moet veel docenten mailen om **gastlessen** in te plannen: dezelfde uitleg, per persoon
aangepast, en een voorstel voor data in een bepaalde maand. Handmatig is dat tientallen keren
knippen en plakken in Outlook, met de bijbehorende fouten (verkeerde naam, iemand dubbel, iemand
vergeten).

## De stroom door het scherm

Het venster is één pagina met drie stappen, van boven naar beneden:

1. **Ledenlijst** — een `.csv` of `.xlsx` inlezen. De app raadt de e-mailkolom, toont de lijst in
   een tabel, en zet per rij een vinkje "Versturen". Ongeldige en dubbele adressen worden gemeld.
2. **Bericht** — onderwerp en tekst, met `{{Kolomnaam}}`-velden die per ontvanger ingevuld worden.
   Sjablonen komen uit een zelfgekozen map (`.json`-bestanden) en zijn daar ook in op te slaan.
   Rechts staat de **gastlesplanning**: kies een maand en klik datums aan (weekenden en dagen in
   het verleden zijn geblokkeerd), en daaronder een live voorbeeld met een echte ontvanger erin.
3. **Versturen** — kies het Outlook-account (moet `@fnv.nl` zijn), stuur eerst een testmail naar
   jezelf, en verstuur dan naar de aangevinkte ontvangers. Met pauze tussen de mails, een
   stopknop, een logboek per ontvanger, en achteraf "mislukte opnieuw".

## Velden in een sjabloon

- Alles wat als kolom in de ledenlijst staat: `{{Voornaam}}`, `{{School}}`, …
- Drie velden die de app zelf invult vanuit de planning: `{{Maand}}` ("oktober"),
  `{{MaandJaar}}` ("oktober 2026") en `{{Datumopties}}` (de aangeklikte datums als leesbare zin).

Een `{{veld}}` dat nergens vandaan komt, wordt als waarschuwing onder het tekstvak getoond —
liever dat dan een mail met `{{Voornam}}` erin de deur uit.

## Wat de app achterlaat op schijf

- `verzendrapport_*.csv` naast het bronbestand, per verzending, met per ontvanger de status
  (UTF-8 **met** BOM, anders maakt Excel er `Ren�` van)
- een verzendgeschiedenis in AppData met **gehashte** adressen — genoeg om te waarschuwen "deze
  mailing ging 3 dagen geleden al naar 12 van deze mensen", zonder een ledenlijst achter te laten
- de laatst gekozen sjablonenmap, in AppData

Verder niets: geen wachtwoorden, geen adressen, geen mailinhoud.
