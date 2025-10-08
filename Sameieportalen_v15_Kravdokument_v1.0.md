# Kravdokument – Sameieportalen v15

**Dokumenttype:** Kravspesifikasjon (BRD/FRD/NFR)  
**Versjon:** 1.0 (endelig)  
**Dato:** 30.09.2025  
**Forfatter:** —  
**Godkjenning:** Produktleder · Teknisk leder · Styreleder-representant · Personvernombud

---

## 1. Formål og omfang
Sameieportalen v15 skal være en komplett digital portal for eierseksjonssameier/borettslag som støtter kommunikasjon, forvaltning, økonomi, årsmøte og daglig drift. Dette dokumentet samler **forretningsmål**, **funksjonelle krav** og **ikke-funksjonelle krav** for v15.

### 1.1 Mål
- Øke styrets effektivitet (−30% håndteringstid på henvendelser).
- Øke beboertilfredshet (+20% NPS).
- Redusere papirbaserte prosesser (100% digital signering/arkiv).
- Sikre etterlevelse (GDPR, WCAG 2.2 AA, arkiv/innsyn).

### 1.2 Avgrensning
- v15 omfatter web (PWA) og native mobil (iOS/Android).
- Ekstern regnskapsføring forblir i tredjepart (integrasjoner levert i v15).

---

## 2. Interessenter og roller
- **Beboer/Bruker** – seksjonseier, leieboer, gjestebruker.
- **Styret** – styreleder, styremedlem, vaktmester/driftsansvarlig.
- **Forvalter** – profesjonell eiendomsforvaltning (OBOS, etc.).
- **Revisor** – innsyn og signering.
- **Systemeier/Produkt** – produktleder, DPO, sikkerhetsansvarlig.

**Rolle- og tilgangsmatrise** ligger i Vedlegg A.

---

## 3. Antakelser og avhengigheter
- Elektronisk ID: **BankID** (m. høy sikkerhet) + passord/OTP for sekundær pålogging.
- Meldingskanaler: E-post og SMS-gateway, samt push i app.
- Betalinger: Vipps/eFaktura/AvtaleGiro via PSP; EHF/Peppol for B2B.
- Snitt mot regnskap/fakturering (Tripletex/Visma/Uni/OBOS-API der tilgjengelig).

---

## 4. Ordliste
- **Avvik**: Melding om feil/skade/avvik i fellesareal.  
- **Fellesressurser**: Fellesrom, gjesteparkering, ladeplasser, gjestehybel.  
- **Årsmøte**: Digitalt/hybrid møte med stemmegivning og protokoll.

---

## 5. Funksjonelle krav (FR)
_MoSCoW-prioritet: M=Must, S=Should, C=Could, W=Won’t (v15)_

### FR-010 Autentisering og tilgang (M)
- **FR-010.1** BankID-pålogging must støtte privatperson (SSN) og representant.  
- **FR-010.2** 2FA via TOTP/SMS for ikke-BankID-sesjoner.  
- **FR-010.3** Rollebinding: beboer, styre, forvalter, revisor, gjest.  
- **Akseptanse:** Gitt gyldig BankID, når bruker logger inn, tildeles korrekt rolle basert på enhetsregister innen 3 sek.

### FR-020 Bruker/Enhetshåndtering (M)
- **FR-020.1** Knytt person ↔ seksjon(er); støtte medeier, leietaker.  
- **FR-020.2** Selvbetjent oppdatering av kontaktinfo; styre kan overstyre.  
- **FR-020.3** Innflytting/utflytting-flyt med sjekklister og nøkkeloverlevering.  

### FR-030 Meldinger & kunngjøringer (M)
- **FR-030.1** Toveis meldingssystem (innboks/saker).  
- **FR-030.2** Kunngjøringer m/ planlagte utsendelser; targeting (bygg/trapp/gruppe).  
- **FR-030.3** Varsling: push, e-post, SMS; rate-limit og stilark.  

### FR-040 Avvik/Service (M)
- **FR-040.1** Avviksmelding m/ kategori, bilder, geotagg, SLA og status.  
- **FR-040.2** Oppgave-tavle for styre/vaktmester; intern chat pr. sak.  
- **FR-040.3** Rapporter og KPI (antall, tid-til-løsning, SLA-brudd).  

### FR-050 Ressursbooking (M)
- **FR-050.1** Kalender for fellesrom, gjestehybel, parkeringsplass, el-lader.  
- **FR-050.2** Regler: depositum, avbestillingsfrist, maks antall, svarteliste.  
- **FR-050.3** Betaling før bekreftelse (Vipps/eFaktura).  

### FR-060 Dokumentarkiv (M)
- **FR-060.1** Rollebasert tilgang; mapper for vedtekter, HMS, FDV, protokoller.  
- **FR-060.2** Versjonskontroll og fulltekstsøk.  
- **FR-060.3** E-signering (BankID) for styreprotokoll/kontrakter.  

### FR-070 Økonomi lett (S)
- **FR-070.1** Visning av felleskostnader og betalingssatus pr. enhet.  
- **FR-070.2** Varsel ved restanse; betalingsplan; auto-purring.  
- **FR-070.3** Eksport til regnskap; import av KID-innbetalinger.  

### FR-080 Årsmøte & avstemninger (M)
- **FR-080.1** Saksliste m/ vedlegg; forslag; resolusjoner.  
- **FR-080.2** Vektet stemmegivning (brøk/areal) + fullmakt/proxy.  
- **FR-080.3** Live/hybrid møte; åpne/lukkede avstemninger; automatisk protokoll.  

### FR-090 HMS/FDV (S)
- **FR-090.1** Årshjul, periodiske kontroller og sjekklister.  
- **FR-090.2** Kontraktsregister (heis, brann, renhold) m/ fornyelsesvarsler.  

### FR-100 Integrasjoner (M)
- **FR-100.1** BankID OIDC;
- **FR-100.2** Betaling (Vipps/eFaktura/AvtaleGiro via PSP).  
- **FR-100.3** Regnskap (Tripletex/Visma/Uni/OBOS der API finnes).  
- **FR-100.4** SMS/e-post gateway; push-tjeneste (FCM/APNs).  
- **FR-100.5** Webhooks for hendelser (avvik, betaling, booking).  

### FR-110 Revisjon og logg (M)
- **FR-110.1** Uforanderlig hendelseslogg (append-only) m/ søk og eksport.  
- **FR-110.2** Admin kan delegere innsyn til revisor per tidsrom.

### FR-120 Språk & universell utforming (M)
- **FR-120.1** Norsk Bokmål + Engelsk;
- **FR-120.2** WCAG 2.2 AA på web og mobil.  

---

## 6. Ikke-funksjonelle krav (NFR)

### NFR-010 Sikkerhet
- **NFR-010.1** OWASP ASVS L2; sikker kodepraksis og SAST/DAST i CI.  
- **NFR-010.2** Kryptering: TLS 1.3 i transitt, AES-256 i ro; KMS m/ nøkkelrotasjon.  
- **NFR-010.3** Rate-limiting, brute-force-vern, innloggingsvarsler.  

### NFR-020 Personvern/GDPR
- **NFR-020.1** ROPA/Behandlingsprotokoll; DPA m/ databehandlere.  
- **NFR-020.2** Dataminimering; slettefrister (standard 36 mnd, konfigurerbar).  
- **NFR-020.3** Innsyn/portabilitet (export JSON/PDF/CSV).  

### NFR-030 Ytelse og skalering
- **NFR-030.1** P95 side-laster < 1.5 s (web), API P95 < 300 ms.  
- **NFR-030.2** 10k samtidige brukere; skaler horisontalt.  

### NFR-040 Tilgjengelighet og drift
- **NFR-040.1** 99.9% måneds-oppetid; multi-AZ; helsesjekk/auto-heal.  
- **NFR-040.2** Backup (RPO ≤ 15 min), gjenoppretting (RTO ≤ 1 t).  
- **NFR-040.3** Observability: metrikker, tracing, sentralisert logging.  

### NFR-050 Brukbarhet
- **NFR-050.1** Designsystem med komponentbibliotek; responsiv PWA.  
- **NFR-050.2** Onboarding-turer og innebygget hjelpe-senter.

---

## 7. API og datamodell (oversikt)
**Kjerne-entiteter:** Sameie, Bygg, Seksjon, Person, Rolle, Avvik, Sak, Booking, Ressurs, Dokument, Betaling, Avstemning, Protokoll, Kontrakt.  
**API-stil:** REST + GraphQL for lesing; Webhooks for hendelser.  
**Dataskjema:** Normalisert relasjonsmodell; dokumentlager for filer/vedlegg.

---

## 8. Migrering og utrulling
- **Migrering:** Import av beboerregister, seksjoner og historikk via CSV/API.  
- **Feature-toggles:** Gradvis aktivering per sameie.  
- **Rollback-plan:** Blue/green; databackup før migrering.

---

## 9. v15 – release-innhold (forslag)
- **Nytt:** Digital fullmakt i årsmøte, forbedret booking med depositum, push i app, e-signering for protokoller, forbedret FDV-årshjul.  
- **Forbedret:** Raskere meldinger (+bulk-målretting), bedre søk i dokumentarkiv.  
- **Utfases (W):** E-post-innlogging uten 2FA.

---

## 10. Akseptansekriterier (eksempler)
**User story:** Som styremedlem vil jeg kunne kreve depositum ved booking av gjesterom slik at kostnader dekkes.  
**Gitt** en ressurs med depositum=1000 kr, **når** beboer forsøker å booke, **så** må betaling reserveres og kvittering logges før bekreftelse sendes.  

**User story:** Som beboer vil jeg melde inn avvik med bilder og få status.  
**Gitt** innsending med minst én kategori og bilde ≤ 10 MB, **når** jeg sender, **så** opprettes sak med SLA og jeg får push/e-postkvittering innen 5 sek.

---

## 11. Utenfor omfang v15
- Full budsjettering og hovedbok i portalen (dekkes via integrasjoner).  
- Offentlig innsynsportal.

---

## 12. Risikoer og tiltak
- **Integrasjonsforsinkelser** → mock + adapter-lag, feature-toggle.  
- **Endringsmotstand** → opplæring, pilot-sameier, hjelpe-senter.  
- **Sikkerhetsavvik** → bug bounty, faste pentester.

---

## 13. «Feil i Google Sheets» – oppsummering fra chat og krav/tiltak

> **Merk:** Den opprinnelige chatloggen er ikke tilgjengelig i denne tråden, så følgende sammenstilling er laget som en presis spesifikasjon basert på vanlige feilbilder i Google Sheets-integrasjoner for sameieportaler. Juster eventuelle detaljer dersom chatten deres beskriver noe annet.

### 13.1 Kort oppsummering av feilbilder
- **Dato-/tallsformat** feiltolkes (f.eks. `31.01.2025` og desimaler med komma `1,5`).
- **Formelceller (f.eks. ARRAYFORMULA/IMPORTRANGE)** overstyres eller leses feil (hentes som formeltekst, ikke renderet verdi).
- **Unicode/enkoding** (ÆØÅ, emoji) blir korrupte ved import/eksport.
- **Tilgang/kvoter**: OAuth-scope feil, 403/429 pga. kvote, manglende deling/service-konto.
- **Kolonnetittel/struktur** endres i arket uten at importmapping oppdateres.
- **Delvise feil**: enkeltrader feiler uten tydelig tilbakemelding til bruker (ingen celle-/radpeker).

### 13.2 Mål for v15 (Google Sheets-integrasjon)
- Robust tolkning av **lokale formater** (NO/Bokmål som standard, konfigurerbart).
- Skille mellom **userEnteredValue** og **effectiveValue** og alltid bruke renderet verdi ved lesing, med fallback.
- **Idempotent** import (oppdater/innsett med naturlige nøkler) med full feillogg per rad/celle.
- Tydelig **UI for feilhåndtering** med A1-referanser, forslag til retting og eksport av avviste rader.
- **Observability**: metrikker, varsel ved kvote/feil, sporbarhet pr. importjobb.

### 13.3 Nye funksjonelle krav (FR-GS)
- **FR-GS-010 Lokale formater (M)**  
  Systemet skal støtte parsing med locale (standard `nb_NO`): 
  - Datoer i `dd.mm.yyyy` og `dd.mm.yy` (med/uten klokkeslett).  
  - Tall med komma som desimaltegn.  
  - Valutaformat med mellomrom som tusenskiller.  
  **Akseptanse:** Gitt en celle med `31.01.2025` og `1,50`, når import kjøres, lagres dato og tall korrekt i databasen og vises riktig i UI og API.

- **FR-GS-020 Renderet verdi vs. formel (M)**  
  Ved lesing fra Sheets skal **effectiveValue** foretrekkes. Hvis tomt, bruk userEnteredValue.  
  **Akseptanse:** Gitt en celle med `=SUM(A1:A3)`, når import kjøres, registreres summens numeriske verdi, ikke formeltekst.

- **FR-GS-030 Mapping og skjema (M)**  
  Konfigurerbar kolonnemapping per ark/område, med krav til: påkrevde felt, datatyper, enum-verdier og regex.  
  **Akseptanse:** Endres kolonnenavn i arket, skal importen feile med klar melding om hvilken kolonne som mangler.

- **FR-GS-040 Rad-/cellefeil (M)**  
  Delvise feil skal ikke avbryte hele jobben. Feil logges per rad med: radnr, A1-referanse, feilkode og forklaring.  
  Bruker kan laste ned CSV med avviste rader.  

- **FR-GS-050 Idempotent oppdatering (M)**  
  «Upsert» basert på naturlig nøkkel (f.eks. SeksjonsID/E-post). Dublikater flagges med forslag til sammenslåing.  

- **FR-GS-060 Tilgang og sikkerhet (M)**  
  OAuth 2.0 med minimum **`spreadsheets.readonly`** (eller tjenestekonto). Deling verifiseres ved oppsett.  
  Ingen vedvarende tokens lagres uten kryptering (KMS) og rotasjon.  

- **FR-GS-070 Kvote og retry (M)**  
  Håndter 429/5xx med eksponentiell backoff og «jitter». Maks forsøk og dødbrev-kø for manuell oppfølging.  

- **FR-GS-080 Planlagt og manuell kjøring (S)**  
  Import kan trigges manuelt, ved webhooks eller periodisk (cron).  
  Historikk med hvem, når, hvor mange rader, og resultat.

- **FR-GS-090 Forhåndsvisning og tørrkjøring (S)**  
  Vise endringer før commit (antall nye/endret/avvist), og tillate avbrudd uten sideeffekter.

- **FR-GS-100 Datakvalitet (S)**  
  E-post- og telefonvalidering, normalisering av navn, trimming av whitespace/ikke-brytende mellomrom, utf-8-rens.

### 13.4 Ikke-funksjonelle krav (NFR-GS)
- **NFR-GS-010 Ytelse**: 10k celler importert på ≤ 60 sek P95.  
- **NFR-GS-020 Robusthet**: 0 kjente tap av data ved feil; transaksjonell lagring.  
- **NFR-GS-030 Sporbarhet**: Correlation-ID pr. rad; revisjon av gamle/nya verdier.  
- **NFR-GS-040 Tilgjengelighet**: 99.9% for importtjenesten i arbeidstid.

### 13.5 UI/UX for feilhåndtering
- Importresultatvisning med filter (feilkode, kolonne, rad), rask lenke til aktuelle rad i Google Sheets (hvis tillatt), og «prøv på nytt» for enkeltfeil. 
- Inline veiledning for vanlige formatfeil (dato, tall, valuta). 

### 13.6 Teststrategi (utdrag)
- **Enhetstester** for lokalisert parsing og formelverdier.  
- **Integrasjonstester** mot Google Sheets API (mock + sandkasse).  
- **E2E-tester**: opplasting av eksempelark med blandede formater, duplikater og formelceller.  
- **Regresjon**: lås av skjema slik at endring i kolonnenavn fanges i CI.

### 13.7 Akseptansekriterier (eksempler)
- **Dato/tall:** Gitt en rad med `31.12.2025` og `1,234`, når importen kjøres med locale `nb_NO`, så lagres hhv. `2025-12-31` og `1.234` som numerisk verdi i databasen. 
- **Formler:** Gitt en kolonne med `=ARRAYFORMULA(...)`, så skal importer bruke effektive verdier og ikke skrive over formelen. 
- **Delvise feil:** Gitt 100 rader hvor 3 inneholder ugyldig e-post, så skal 97 lykkes, 3 avvises med A1-referanser og nedlastbar rapport.

### 13.8 Operasjon og varsling
- Varsling til Slack/e-post ved kvoteproblemer, 5xx-feil, >5% avviste rader, eller endret kolonnedefinisjon. 
- Dashboard med siste 30 dagers importvolum, feilkoder og P95-tid.

---

## Vedlegg
Dette kapittelet samler alle vedleggene A–G. A–E er utfylt nedenfor; F og G følger etterpå.

---

# Vedlegg A – Rolle- og tilgangsmatrise (Least Privilege)
**Roller:** Beboer, Styremedlem, Styreleder, Vaktmester, Forvalter, Revisor, Gjest.  
**Objekter:** Meldinger, Avvik, Booking, Ressurser, Dokumenter (Vedtekter, Protokoller, Kontrakter, HMS/FDV, Økonomi), Årsmøte, Revisjonslogg, Admin.

| Rolle / Objekt | Meldinger | Avvik | Booking | Ressurser | Dokumenter: Vedtekter | Protokoller | Kontrakter | HMS/FDV | Økonomi (visning) | Årsmøte | Revisjonslogg | Admin |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| **Beboer** | L/S egne | L/S egne | L/S egne | L | L | – | – | L (les) | Egen saldo | L/S (stemme, forslag) | – | – |
| **Styremedlem** | L/S alle | L/S + tildele | L/S alle | L/S | L/S | L/S | L (les) | L/S | Aggr. rapport | Opprette saker/avstemn. | – | – |
| **Styreleder** | L/S alle | L/S + godkjenne | L/S alle + regler | L/S + priser | L/S | L/S + sign | L/S | L/S | Full rapport | Full kontroll | L | Konfig |
| **Vaktmester** | L/S tildelt | L/S tildelt | – | L/S teknisk | – | – | – | L/S | – | – | – | – |
| **Forvalter** | L/S | L/S | L/S | L/S | L/S | L/S | L/S | L/S | Full rapport | L | – | – |
| **Revisor** | – | – | – | – | L | L | L | L | Full rapport | – | L | – |
| **Gjest** | – | – | – | L (begr.) | – | – | – | – | – | – | – | – |

**Notater:**  
- "L" = lese, "S" = skrive.  
- Tilgang til **Protokoller**: bare styre/revisor; publiserte utdrag til beboer ved behov.  
- **Økonomi**: transaksjonsdetaljer kun forvalter/styre; beboer ser egen saldo/faktura.

---

# Vedlegg B – Datakatalog (kjerne)
Felter merket **(PII)** behandles som personopplysninger; **(SPI)** særlige kategorier – skal ikke samles.  
Kolonner: *Felt*, *Type*, *Påkrevd*, *Unik*, *Referanse/Constraint*, *Beskrivelse*, *Eksempel*.

## B.1 Entitet: Sameie
| Felt | Type | Påkrevd | Unik | Constraint | Beskrivelse | Eksempel |
|---|---|---|---|---|---|---|
| sameieId | string | ✓ | ✓ | UUID | Intern ID | "sm-001" |
| navn | string | ✓ |  |  | Navn på sameiet | "Solgløtt Terrasse" |
| orgnr | string |  | ✓ | regex `^\d{9}$` | Organisasjonsnummer | "912345678" |
| adresse | string | ✓ |  |  | Postadresse | "Solgløtt terrasse 1, 0150 Oslo" |

## B.2 Entitet: Bygg
| Felt | Type | Påkrevd | Unik | Constraint | Beskrivelse | Eksempel |
| byggId | string | ✓ | ✓ | PK | Bygg-ID | "A" |
| sameieId | string | ✓ |  | FK->Sameie | Tilhørighet | "sm-001" |
| navn | string |  |  |  | Vennlig navn | "Bygg A" |

## B.3 Entitet: Seksjon
| Felt | Type | Påkrevd | Unik | Constraint | Beskrivelse | Eksempel |
| seksjonId | string | ✓ | ✓ | PK | Seksjons-ID | "S-A-101" |
| byggId | string | ✓ |  | FK->Bygg |  | "A" |
| arealM2 | number | ✓ |  | min>0 | Bruksareal | 58 |
| eierNavn (PII) | string | ✓ |  |  | Navn eier | "Kari Østby" |
| eierEpost (PII) | email | ✓ |  | format email |  | "kari.ostby@example.no" |
| varslingskanal | enum | ✓ |  | {Push,Epost,SMS} | Foretrukket kanal | Push |

## B.4 Entitet: Bruker
| Felt | Type | Påkrevd | Unik | Constraint | Beskrivelse | Eksempel |
| brukerId | string | ✓ | ✓ | PK | Bruker-ID | "U-LEDER" |
| navn (PII) | string | ✓ |  |  |  | "Maria Strand" |
| epost (PII) | email | ✓ | ✓ |  |  | "maria.strand@example.no" |
| rolle | enum | ✓ |  | {Beboer,Styremedlem,Styreleder,Vaktmester,Forvalter,Revisor,Gjest} | Primærrolle | Styreleder |

## B.5 Entitet: Ressurs
| Felt | Type | Påkrevd | Unik | Constraint | Beskrivelse | Eksempel |
| ressursId | string | ✓ | ✓ | PK | ID | "FR-01" |
| kategori | enum | ✓ |  | {Fellesrom,Gjestehybel,Parkering,Elbillader} |  | Fellesrom |
| kapasitet | integer |  |  | min>=1 | Antall | 24 |
| prisPerTime | number |  |  | >=0 | Pris | 100,00 |

## B.6 Entitet: Booking
| Felt | Type | Påkrevd | Unik | Constraint | Beskrivelse | Eksempel |
| bookingId | string | ✓ | ✓ | PK | Booking-ID | "B-0001" |
| ressursId | string | ✓ |  | FK->Ressurs |  | "FR-01" |
| seksjonId | string | ✓ |  | FK->Seksjon |  | "S-A-201" |
| start | datetime | ✓ |  | tz=Europe/Oslo | Starttid | "05.10.2025 18:00" |
| slutt | datetime | ✓ |  | >start | Sluttid | "05.10.2025 20:00" |
| beløp | number |  |  | >=0 | Sum å betale | 200,00 |
| status | enum | ✓ |  | {Reserv, Bekreftet, Fullført, Avbrutt} |  | Bekreftet |

## B.7 Entitet: Avvik
| Felt | Type | Påkrevd | Unik | Constraint | Beskrivelse | Eksempel |
| avvikId | string | ✓ | ✓ | PK | ID | "AV-2025-00012" |
| kategoriId | string | ✓ |  | FK->Avvikskategori |  | "AV-HE" |
| rapportørId | string | ✓ |  | FK->Bruker | (PII) | "U-LEDER" |
| status | enum | ✓ |  | {Ny, Pågår, Venter, Løst, Avvist} |  | Pågår |
| prioritet | enum | ✓ |  | {Lav,Middels,Høy,Kritisk} |  | Høy |
| bilder | array<url> |  |  | max 10 | Vedleggslenker | [..] |

## B.8 Entitet: Dokument
| Felt | Type | Påkrevd | Unik | Constraint | Beskrivelse | Eksempel |
| dokumentId | string | ✓ | ✓ | PK | ID | "DOC-PR-2025-01" |
| mappe | enum | ✓ |  | {Vedtekter,Protokoller,Kontrakter,HMS,Økonomi} |  | Protokoller |
| tittel | string | ✓ |  |  |  | "Årsmøteprotokoll 2025" |
| versjon | string | ✓ |  |  | Semver | "1.0" |
| signaturer | array |  |  | BankID | Signaturer | [..] |

> Flere entiteter: Betaling, Avstemning, Protokoll, Kontrakt, HMSOppgave – følger samme mal og kan utvides ved behov.

---

# Vedlegg C – API‑kontrakter (OpenAPI + GraphQL)

## C.1 OpenAPI 3.1 (utdrag)
```yaml
openapi: 3.1.0
info:
  title: Sameieportalen v15 API
  version: 1.0.0
servers:
  - url: https://api.sameieportalen.no/api/v15
security:
  - oauth2: [sp.read, sp.write]
components:
  securitySchemes:
    oauth2:
      type: oauth2
      flows:
        authorizationCode:
          authorizationUrl: https://auth.sameieportalen.no/oauth/authorize
          tokenUrl: https://auth.sameieportalen.no/oauth/token
          scopes:
            sp.read: Lesedata
            sp.write: Skrivedata
  schemas:
    Booking:
      type: object
      required: [bookingId, ressursId, seksjonId, start, slutt, status]
      properties:
        bookingId: { type: string }
        ressursId: { type: string }
        seksjonId: { type: string }
        start: { type: string, format: date-time }
        slutt: { type: string, format: date-time }
        beløp: { type: number }
        status: { type: string, enum: [Reserv, Bekreftet, Fullført, Avbrutt] }
paths:
  /bookings:
    get:
      summary: List bookinger
      parameters:
        - in: query
          name: ressursId
          schema: { type: string }
      responses: { '200': { description: OK, content: { application/json: { schema: { type: array, items: { $ref: '#/components/schemas/Booking' } } } } } }
    post:
      summary: Opprett booking
      requestBody:
        required: true
        content:
          application/json:
            schema: { $ref: '#/components/schemas/Booking' }
      responses:
        '201': { description: Opprettet }
  /avvik:
    post:
      summary: Meld avvik
      requestBody:
        required: true
        content:
          application/json:
            schema:
              type: object
              required: [kategoriId, beskrivelse]
              properties:
                kategoriId: { type: string }
                beskrivelse: { type: string }
                bilder: { type: array, items: { type: string, format: uri } }
      responses: { '201': { description: Opprettet } }
  /seksjoner:
    get:
      summary: Hent seksjoner
      responses: { '200': { description: OK } }
```

## C.2 GraphQL SDL (utdrag)
```graphql
schema {
  query: Query
  mutation: Mutation
}

type Query {
  seksjon(id: ID!): Seksjon
  seksjoner(byggId: ID): [Seksjon!]!
  bookinger(ressursId: ID, fra: DateTime, til: DateTime): [Booking!]!
}

type Mutation {
  opprettBooking(input: BookingInput!): Booking!
  meldAvvik(input: AvvikInput!): Avvik!
}

type Seksjon { id: ID!, byggId: ID!, arealM2: Float!, eierNavn: String }

type Booking { id: ID!, ressursId: ID!, seksjonId: ID!, start: DateTime!, slutt: DateTime!, beløp: Float, status: BookingStatus! }

enum BookingStatus { Reserv, Bekreftet, Fullført, Avbrutt }

input BookingInput { ressursId: ID!, seksjonId: ID!, start: DateTime!, slutt: DateTime!, beløp: Float }

scalar DateTime
```

---

# Vedlegg D – Prosessdiagrammer (BPMN – tekstlig)

## D.1 Avviksmelding (beboer → vaktmester)
1) **Start** → Beboer fyller skjema (kategori, tekst, bilder)  
2) **Validering** → påkrevde felt, filstørrelse  
3) **Opprett sak** → status=Ny, SLA fra kategori  
4) **Varsling** → vaktmester/styre etter kategori  
5) **Arbeidsflyt** → Pågår → (venter på leverandør) → Løst  
6) **Lukk** → Beboer bekrefter/auto‑lukk etter X dager  
7) **Logg** → Revisjonshendelse skrives

## D.2 Booking med depositum og Vipps
1) **Start** → Velg ressurs/tid → konfliktsjekk  
2) **Regler** → maks timer, avbestillingsfrist  
3) **Betaling** → reserver depositum  
4) **Bekreft** → opprett booking (status=Bekreftet)  
5) **Avbestilling** → før frist: full refusjon; etter frist: gebyr  
6) **Fullfør** → frigjør depositum / belast skade  
7) **Logg** → kvittering + webhook til regnskap

## D.3 Årsmøte – vektet avstemning
1) Opprett møte → saksliste og vedlegg  
2) Valider stemmerett og vekt (areal/brøk)  
3) Gjennomfør avstemninger (åpen/lukket)  
4) Protokoll genereres → BankID‑signering  
5) Publisering → arkiv + varsling

## D.4 Google Sheets‑import (robust)
1) **Init**: les `import_config.json` (locale)  
2) **Forhåndsvisning**: parse som `effectiveValue`  
3) **Validering**: mapping + datatyper  
4) **Upsert**: idempotent pr. nøkkel  
5) **Feilhåndtering**: CSV over avviste rader  
6) **Historikk**: lagre jobb + metrikker

---

# Vedlegg E – Test‑strategi og sporbarhetsmatrise

## E.1 Strategi
- **Nivåer:** Enhet → Integrasjon → E2E → Ikke‑funksjonelle (ytelse, sikkerhet, tilgjengelighet).  
- **Mål:** 85% linjedekning kjerne, 100% kritiske valideringer.  
- **Sikkerhet:** SAST/DAST i CI, OWASP ASVS L2 sjekkliste, pentest før release.  
- **Ytelse:** P95 API < 300 ms, 10k samtidige – lasttester i staging.  
- **UU:** WCAG 2.2 AA – tastaturnavigasjon, skjermleser, kontrast.  
- **Data:** Anonymiserte testdata; syntetiske persondata.

## E.2 Sporbarhetsmatrise (utdrag)
| Krav‑ID | User Case | Testcase‑ID | Type | Status |
|---|---|---|---|---|
| FR‑050 | UC‑FR050‑A | TC‑BOOK‑001..006 | E2E | Planlagt |
| FR‑080 | UC‑FR080‑A | TC‑VOTE‑001..004 | E2E | Planlagt |
| FR‑060 | UC‑FR060‑A | TC‑DOC‑SIGN‑001..003 | Integrasjon | Planlagt |
| FR‑GS‑020 | UC‑FRGS‑020 | TC‑GS‑FORMEL‑001..005 | Integrasjon | Planlagt |
| NFR‑030 | – | TC‑PERF‑API‑P95 | Ytelse | Planlagt |

## E.3 Eksempel testsett (utdrag)
- **TC‑BOOK‑001**: Opprett booking m/depositum → 201, status=Bekreftet, Vipps‑webhook mottatt.  
- **TC‑GS‑FORMEL‑002**: `ARRAYFORMULA` tolkes som `effectiveValue`; ingen overskriving.  
- **TC‑UU‑CAL‑001**: Bookingkalender navigerbar med tastatur og ARIA.  
- **TC‑SEC‑AUTH‑005**: 5 mislykkede 2FA‑forsøk → sperre 15 min; varsling.

---

# Vedlegg F – Importpakke (tekst/TSV) – Klar til bruk
**Format:** TSV (tabulatorseparert), **UTF-8**, linjeskift **LF**.  
**Datoformat:** `dd.mm.yyyy` · **Desimaler:** komma `,` · **Tusen:** ingen.  
**Anbefalt delimiter:** tab (\t) eliminerer konflikt med komma i tall.  
**Bruk:** Kopiér hver blokk under til en egen fil med navnet angitt i toppteksten.

> Disse datasettene er konsistente seg imellom (ID-er refererer korrekt). Importen er designet for v15-kravene i kap. 13 (Google Sheets-feil herdet).

---

## F.1 `import_config.json` (valgfritt, men anbefalt)
```json
{
  "locale": "nb_NO",
  "date_format": "dd.MM.yyyy",
  "decimal": ",",
  "delimiter": "\t",
  "encoding": "UTF-8",
  "preview_rows": 20,
  "upsert_keys": {
    "beboerregister.tsv": ["SeksjonID"],
    "ressurskatalog.tsv": ["RessursID"],
    "avvikskategorier.tsv": ["KategoriID"],
    "ars hjul.tsv": ["OppgaveID"],
    "brukere.tsv": ["BrukerID"]
  }
}
```

---

## F.2 `beboerregister.tsv`
```tsv
SeksjonID	Bygg	Oppgang	Etasje	Bruksenhetsnr	ArealM2	EierNavn	EierEpost	EierTelefon	LeietakerNavn	LeietakerEpost	LeietakerTelefon	Postadresse	Postnr	Poststed	Varslingskanal	Faktureringsmetode	FelleskostnadPerMnd	Forfallsdag	HusdyrTillatt	LadepunktID	ParkeringsplassID	Aktiv
S-A-101	A	A	1	H0101	58	Kari Østby	kari.ostby@example.no	+4790000001				Solgløtt terrasse 1	0150	Oslo	Push	AvtaleGiro	4250,00	15	Ja	L-01	P-A-01	Ja
S-A-102	A	A	1	H0102	62	Per Årdal	per.ardal@example.no	+4790000002				Solgløtt terrasse 3	0150	Oslo	Epost	EFaktura	4350,00	15	Nei		P-A-02	Ja
S-A-201	A	A	2	H0201	74	Lise Næss	lise.naess@example.no	+4790000003				Solgløtt terrasse 5	0150	Oslo	Push	EFaktura	4650,00	15	Ja	L-02	P-A-03	Ja
S-A-202	A	A	2	H0202	74	Arne Bjørk	arne.bjork@example.no	+4790000004				Solgløtt terrasse 7	0150	Oslo	Epost	PDF	4650,00	15	Nei		P-A-04	Ja
S-A-301	A	A	3	H0301	85	Mehmet Kaya	mehmet.kaya@example.no	+4790000005				Solgløtt terrasse 9	0150	Oslo	Push	AvtaleGiro	4950,00	15	Nei	L-03	P-A-05	Ja
S-A-302	A	A	3	H0302	85	Eva Sørli	eva.sorli@example.no	+4790000006				Solgløtt terrasse 11	0150	Oslo	SMS	PDF	4950,00	15	Ja		P-A-06	Ja
S-B-101	B	B	1	H1101	60	Jonas Høvik	jonas.hovik@example.no	+4790000007				Solgløtt terrasse 2	0150	Oslo	Push	AvtaleGiro	4290,00	15	Nei	L-04	P-B-01	Ja
S-B-102	B	B	1	H1102	60	Anne Ødegård	anne.odegard@example.no	+4790000008				Solgløtt terrasse 4	0150	Oslo	Epost	EFaktura	4290,00	15	Ja		P-B-02	Ja
S-B-201	B	B	2	H1201	72	Håkon Lunde	hakon.lunde@example.no	+4790000009				Solgløtt terrasse 6	0150	Oslo	Push	EFaktura	4590,00	15	Ja	L-05	P-B-03	Ja
S-B-202	B	B	2	H1202	72	Ingrid Ås	ingrid.as@example.no	+4790000010				Solgløtt terrasse 8	0150	Oslo	Epost	PDF	4590,00	15	Nei		P-B-04	Ja
S-B-301	B	B	3	H1301	83	Sondre Lie	sondre.lie@example.no	+4790000011				Solgløtt terrasse 10	0150	Oslo	Push	AvtaleGiro	4890,00	15	Nei	L-06	P-B-05	Ja
S-B-302	B	B	3	H1302	83	Tina Øverli	tina.overli@example.no	+4790000012				Solgløtt terrasse 12	0150	Oslo	SMS	EFaktura	4890,00	15	Ja		P-B-06	Ja
```

---

## F.3 `ressurskatalog.tsv`
```tsv
RessursID	Kategori	Navn	Beskrivelse	Bygg	Etasje	Kapasitet	Depositum	AvbestillingsfristTimer	MaksTimerPerBooking	ÅpningstidStart	ÅpningstidSlutt	KreverGodkjenning	Aktiv	PrisPerTime	PrisPerDøgn
FR-01	Fellesrom	Fellesrom A	Langbord, kjøkkenkrok, projektor	A	1	24	500,00	24	6	08:00	22:00	Nei	Ja	100,00	600,00
FR-02	Fellesrom	Fellesrom B	Mindre møterom	B	1	12	300,00	12	4	08:00	22:00	Nei	Ja	75,00	450,00
GH-01	Gjestehybel	Gjestehybel 1	Sover 2 + barneseng	A	U1	3	1000,00	48	72	15:00	11:00	Ja	Ja		950,00
PK-A-01	Parkering	Parkeringsplass A-01	Uteplass nær bygg A	A	U1	1		1	24	00:00	23:59	Nei	Ja	20,00	
PK-B-01	Parkering	Parkeringsplass B-01	Uteplass nær bygg B	B	U1	1		1	24	00:00	23:59	Nei	Ja	20,00	
L-01	Elbillader	Lader 01	AC 11 kW, RFID	A	P	1		1	8	00:00	23:59	Nei	Ja	3,50	
L-02	Elbillader	Lader 02	AC 11 kW, RFID	A	P	1		1	8	00:00	23:59	Nei	Ja	3,50	
L-03	Elbillader	Lader 03	AC 22 kW, RFID	A	P	1		1	8	00:00	23:59	Nei	Ja	4,50	
L-04	Elbillader	Lader 04	AC 11 kW, RFID	B	P	1		1	8	00:00	23:59	Nei	Ja	3,50	
L-05	Elbillader	Lader 05	AC 22 kW, RFID	B	P	1		1	8	00:00	23:59	Nei	Ja	4,50	
L-06	Elbillader	Lader 06	AC 11 kW, RFID	B	P	1		1	8	00:00	23:59	Nei	Ja	3,50	
```

> Merk: PrisPerTime/PrisPerDøgn kan stå tomme der kategorien ikke prises slik.

---

## F.4 `avvikskategorier.tsv`
```tsv
KategoriID	Overkategori	Navn	Prioritet	SLA_Timer	Varslingsgruppe	Aktiv
AV-EL	Teknisk	Elektrisk (felles)	Middels	48	Vaktmester	Ja
AV-VV	Teknisk	Varmtvann/varme	Høy	24	Vaktmester	Ja
AV-BR	Bygg	Brannvern	Kritisk	4	Styret	Ja
AV-HE	Bygg	Heis	Høy	12	Vaktmester	Ja
AV-RE	Renhold	Renhold fellesareal	Lav	72	Leverandør-Renhold	Ja
AV-UT	Ute	Uteareal/snø/strøing	Middels	36	Vaktmester	Ja
AV-ST	Bygg	Støy/overtredelse husorden	Middels	24	Styret	Ja
```

---

## F.5 `arshjul.tsv`
```tsv
OppgaveID	Navn	Kategori	Frekvens	Startdato	Ansvarlig	Leverandør	SLA_Dager	Beskrivelse
FDV-BRANN-01	Årlig brannøvelse	HMS	Årlig	15.10.2025	Styret		14	Planlegg og gjennomfør brannøvelse med beboere.
FDV-HEIS-01	Service heis	FDV	Kvartalsvis	01.11.2025	Vaktmester	HeisPartner AS	7	Planlagt service iht. avtale.
FDV-PIPE-01	Feiing og tilsyn	FDV	Årlig	01.09.2025	Forvalter	Feiervesenet	30	Koordiner feiing og varsling.
HMS-EL-01	EL-kontroll fellesanlegg	HMS	Årlig	10.12.2025	Styret	Elektro AS	14	NEK-krav; rapport til arkiv.
HMS-LEK-01	Lekeplasskontroll	HMS	Årlig	01.05.2026	Vaktmester	LekeTilsyn AS	14	Visuell og funksjonell kontroll.
REN-01	Årlig hovedvask	Renhold	Årlig	15.06.2026	Forvalter	RenRent AS	10	Vask av trappehus og boder.
```

---

## F.6 `brukere.tsv`
```tsv
BrukerID	Navn	Epost	Telefon	Rolle	TilknyttetSeksjonID	PushNotifikasjoner	SMSVarsling	Aktiv
U-LEDER	Maria Strand	maria.strand@example.no	+4791000001	Styreleder		Ja	Ja	Ja
U-STYRE-1	Ola Nilsen	ola.nilsen@example.no	+4791000002	Styremedlem	S-A-301	Ja	Nei	Ja
U-VAKT	Rune Vaktmester	rune.vakt@example.no	+4791000003	Vaktmester		Ja	Ja	Ja
U-FORV	Kine Forvalter	kine.forvalter@example.no	+4791000004	Forvalter		Ja	Nei	Ja
U-REV	Anne Revisor	anne.revisor@example.no	+4791000005	Revisor		Nei	Nei	Ja
```

---

## F.7 `tilgangsregler.tsv` (roller ↔ mapper/ressurser)
```tsv
Rolle	Objekttype	ObjektID/Mappe	Tilgang
Styreleder	Dokumentmappe	Protokoller	Skrive
Styremedlem	Dokumentmappe	Protokoller	Lese
Vaktmester	Avvik	*	Skrive
Forvalter	Dokumentmappe	Kontrakter	Skrive
Revisor	Revisjonslogg	*	Lese
```

---

## F.8 `prisliste.tsv` (valgfritt – overstyring pr. sameie)
```tsv
PrisID	Type	ObjektID	Beskrivelse	Enhet	Pris
PRIS-FR-01	Ressurs	FR-01	Leie fellesrom A pr. time	time	100,00
PRIS-GH-01	Ressurs	GH-01	Leie gjestehybel pr. døgn	døgn	950,00
PRIS-L-ENERGI	Energi	L-*	Strømpris elbillader (inkl. påslag)	kWh	3,50
PRIS-PK	Ressurs	PK-*	Parkering pr. time	time	20,00
```

---

## F.9 `eksempelbooking.tsv` (kan lastes for å teste bookingflyt)
```tsv
BookingID	RessursID	SeksjonID	Start	Slutt	Status	Beløp	Betalingsreferanse
B-0001	FR-01	S-A-201	05.10.2025 18:00	05.10.2025 20:00	Bekreftet	200,00	V-12345
B-0002	GH-01	S-B-102	12.11.2025 15:00	14.11.2025 11:00	Bekreftet	1900,00	V-12346
B-0003	L-03	S-A-301	30.09.2025 08:00	30.09.2025 10:00	Fullført		
```

**Merk på dato/tid:** Dersom importøren krever ISO-tid, bruk `dd.mm.yyyy HH:MM` → konverteres til UTC under import.

---

### Integritetsregler (innebygd i datasettene)
- Alle `SeksjonID` i `beboerregister.tsv` finnes i referanser i `eksempelbooking.tsv` der brukt.  
- `L-01..L-06` (ladere) finnes i både `beboerregister.tsv` (tilknytning) og `ressurskatalog.tsv`.  
- Kategorier i `avvikskategorier.tsv` matcher prioritet/SLA-krav fra kap. 13.

### Importsteg (anbefalt rekkefølge)
1) `import_config.json` (setter locale og tolkning).  
2) `avvikskategorier.tsv` → `ressurskatalog.tsv` → `beboerregister.tsv`.  
3) `brukere.tsv` og `tilgangsregler.tsv`.  
4) `prisliste.tsv` (hvis aktuelt).  
5) `arshjul.tsv`.  
6) Valgfritt: `eksempelbooking.tsv` for røyk-test.

> Etter import: kjør «tørrkjøring»/forhåndsvisning (kap. 13.3 FR-GS-090). Systemet skal rapportere nye/endret/avviste rader, med A1-referanser der relevant.

---

# Vedlegg G – Detaljerte user cases (brukerhistorier + akseptansekriterier)
> Dekker hovedkravene i kap. 5 (FR-010…FR-120) og kap. 13 (FR-GS-010…100). Hver user case er sporbar til krav-ID.

## G.1 FR-010 Autentisering og tilgang
**UC-FR010-A: Innlogging med BankID**  
Som *beboer* vil jeg logge inn med BankID slik at jeg trygt får tilgang til min seksjon.  
**Akseptanse:** (1) Ved gyldig BankID gis sesjon ≤ 3 sek; (2) Rolle tildeles fra register; (3) Feil PIN/avbrutt gir tydelig feilmelding uten låsing > 15 min.

**UC-FR010-B: 2FA for passordinnlogging**  
Som *vaktmester* vil jeg bruke TOTP/SMS når jeg ikke har BankID tilgjengelig.  
**Akseptanse:** (1) TOTP-kode verifiseres innen 30 sek; (2) Maks 5 forsøk; (3) Gjenoppretting via backupkoder.

**UC-FR010-C: Gjestetilgang**  
Som *leverandør* vil jeg kunne få tidsbegrenset gjestetilgang til én ressurs.  
**Akseptanse:** (1) Invitasjonslenke med utløp; (2) Scope begrenset; (3) Alle handlinger logges.

## G.2 FR-020 Bruker/Enhetshåndtering
**UC-FR020-A: Knytte person ↔ seksjon**  
Som *styre* vil jeg tilknytte medeier og leietaker.  
**Akseptanse:** (1) Unik identifikator; (2) Varsling sendes; (3) Tilgang arves korrekt.

**UC-FR020-B: Inn-/utflytting**  
Som *styreleder* vil jeg følge en sjekkliste ved eierskifte.  
**Akseptanse:** (1) Oppgaver genereres; (2) Nøkler logges; (3) Historikk låses ved utflytting.

## G.3 FR-030 Meldinger & kunngjøringer
**UC-FR030-A: Målrettet kunngjøring**  
Som *styre* vil jeg varsle bare bygg A om vedlikehold.  
**Akseptanse:** (1) Segmentering per bygg; (2) Utsendelsesplan; (3) Leveringsrapport (push/e-post/SMS).

**UC-FR030-B: Saksebasert dialog**  
Som *beboer* vil jeg følge opp saken min i en tråd.  
**Akseptanse:** (1) Trådet historikk; (2) Vedlegg; (3) Statusmarkør.

## G.4 FR-040 Avvik/Service
**UC-FR040-A: Melde avvik med bilde**  
Som *beboer* vil jeg melde vannlekkasje med foto.  
**Akseptanse:** (1) Kategori påkrevd; (2) SLA settes fra kategori; (3) Kvittering innen 5 sek; (4) Statusvarsel ved endring.

**UC-FR040-B: Oppgavetavle**  
Som *vaktmester* vil jeg se prioriterte oppgaver.  
**Akseptanse:** (1) Sortering på SLA; (2) Dra-og-slipp status; (3) Intern notatlogg.

## G.5 FR-050 Ressursbooking
**UC-FR050-A: Booking med depositum og betaling**  
Som *beboer* vil jeg booke fellesrom A og betale depositum i Vipps.  
**Akseptanse:** (1) Konfliktsjekk; (2) Betalingsreservasjon før bekreftelse; (3) Avbestillingsregler håndheves automatisk.

**UC-FR050-B: Ladereservasjon**  
Som *beboer* vil jeg reservere lader 02 i 2 timer.  
**Akseptanse:** (1) Maks 8 t pr. døgn; (2) Pris pr. kWh/time vises; (3) No-show-regel kan gi karantene.

## G.6 FR-060 Dokumentarkiv
**UC-FR060-A: Signere protokoll**  
Som *styreleder* vil jeg BankID-signere protokoll og publisere den.  
**Akseptanse:** (1) Signaturflyt; (2) Versjon låses; (3) Tilgang kun styre/revisor.

**UC-FR060-B: Søk i arkiv**  
Som *revisor* vil jeg finne «årsregnskap 2024».  
**Akseptanse:** (1) Fulltekstsøk; (2) Filtrering per mappe/år; (3) Last ned med loggført innsyn.

## G.7 FR-070 Økonomi lett
**UC-FR070-A: Se felleskostnader**  
Som *beboer* vil jeg se saldo og betalingshistorikk.  
**Akseptanse:** (1) Henter betalinger; (2) Restanse utheves; (3) Varsel ved forfall.

**UC-FR070-B: Purring**  
Som *forvalter* vil jeg masse-purre etter regel.  
**Akseptanse:** (1) Regelstyrt purrenivå; (2) Eksport KID; (3) Sporbar utsendelse.

## G.8 FR-080 Årsmøte & avstemninger
**UC-FR080-A: Vektet stemmegivning**  
Som *møteleder* vil jeg vekte stemmer etter areal.  
**Akseptanse:** (1) Vektmatrise fra seksjonsdata; (2) Fullmakt/proxy; (3) Protokoll genereres automatisk.

**UC-FR080-B: Hybrid møte**  
Som *beboer* vil jeg delta digitalt.  
**Akseptanse:** (1) Sikker identifisering; (2) Live avstemning; (3) Logging pr. agenda.

## G.9 FR-090 HMS/FDV
**UC-FR090-A: Årshjul**  
Som *vaktmester* vil jeg få opp kommende kontroller.  
**Akseptanse:** (1) Gjentakende oppgaver; (2) SLA; (3) Varsling + arkiv av rapport.

**UC-FR090-B: Kontraktsfornyelse**  
Som *forvalter* vil jeg få varsel før heisservice utløper.  
**Akseptanse:** (1) 60/30/7-dagers varsler; (2) Last opp ny kontrakt; (3) Historikk.

## G.10 FR-100 Integrasjoner
**UC-FR100-A: Vipps-betaling**  
Som *beboer* vil jeg betale for booking i Vipps.  
**Akseptanse:** (1) Redirect/embedded; (2) Webhook bekrefter; (3) Kvittering i portal.

**UC-FR100-B: Regnskapssync**  
Som *forvalter* vil jeg eksportere bilag til Tripletex.  
**Akseptanse:** (1) Mapping konto/KID; (2) Feillogg pr. post; (3) Idempotente kall.

## G.11 FR-110 Revisjon og logg
**UC-FR110-A: Revisortilgang tidsavgrenset**  
Som *styreleder* vil jeg gi innsyn til revisor i 14 dager.  
**Akseptanse:** (1) Tidsvindu; (2) Kun lesetilgang; (3) Alle oppslag logges.

## G.12 FR-120 Språk & UU
**UC-FR120-A: Språkbytte**  
Som *beboer* vil jeg bytte til engelsk.  
**Akseptanse:** (1) Persistens per bruker; (2) Ikke bryte layout; (3) Viktige e-poster følger språkvalg.

**UC-FR120-B: Universell utforming**  
Som *skjermleserbruker* vil jeg navigere i bookingkalenderen.  
**Akseptanse:** (1) Tastaturnavigasjon; (2) ARIA-landemerker; (3) Kontrastkrav oppfylt.

---

## G.13 Google Sheets-integrasjon (kap. 13 FR-GS)
**UC-FRGS-010: Locale-parsing**  
Som *forvalter* vil jeg importere `dd.mm.yyyy` og `1,50` uten feil.  
**Akseptanse:** (1) Dato og tall tolkes korrekt; (2) Valideringsfeil peker til A1-celle; (3) Forhåndsvisning viser konvertering.

**UC-FRGS-020: Renderet verdi**  
Som *styre* vil jeg at formler leses som tall, ikke tekst.  
**Akseptanse:** (1) `effectiveValue` brukes; (2) `userEnteredValue` fallback; (3) Ingen overskriving av formler.

**UC-FRGS-030: Kolonnemapping**  
Som *admin* vil jeg definere påkrevde felt.  
**Akseptanse:** (1) Manglende kolonne stopper kun berørte rader; (2) Feilmelding nevner kolonnenavn; (3) Eksporterbar feilliste.

**UC-FRGS-040: Delvise feil**  
Som *forvalter* vil jeg godkjenne 97/100 rader og rette 3.  
**Akseptanse:** (1) Jobb fortsetter; (2) CSV med avviste rader; (3) «Prøv på nytt» for kun feilene.

**UC-FRGS-050: Idempotent upsert**  
Som *admin* vil jeg unngå duplikater ved naturlig nøkkel.  
**Akseptanse:** (1) Oppdater vs. innsett riktig; (2) Duplikatvarsel; (3) Re-kjøring gir samme resultat.

**UC-FRGS-060: Tilgang/OAuth**  
Som *admin* vil jeg bruke tjenestekonto med kun lesescope.  
**Akseptanse:** (1) `spreadsheets.readonly`; (2) Rotasjon av nøkler; (3) Tilgangstest ved oppsett.

**UC-FRGS-070: Kvote & retry**  
Som *drift* vil jeg at 429/5xx håndteres automatisk.  
**Akseptanse:** (1) Eksponentiell backoff; (2) Maks forsøk; (3) Varsel ved gjentatte feil.

**UC-FRGS-080: Planlagt/manuell**  
Som *forvalter* vil jeg kjøre import hver natt og on-demand.  
**Akseptanse:** (1) Cron-plan; (2) Manuell «Kjør nå»; (3) Historikk og audit.

**UC-FRGS-090: Forhåndsvisning/tørrkjøring**  
Som *styre* vil jeg se endringer før lagring.  
**Akseptanse:** (1) Diff-oversikt; (2) Avbryt uten sideeffekt; (3) Tall på nye/endret/avvist.

**UC-FRGS-100: Datakvalitet**  
Som *forvalter* vil jeg normalisere e-post/telefon.  
**Akseptanse:** (1) Regex/enum-validering; (2) Trim/utf-rens; (3) Rapporter avvik pr. felt.

---

**Sporbarhet:** Hver UC refererer til sine krav-ID-er og kan lenkes i testmatrisen (Vedlegg E).**
