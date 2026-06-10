# Oversetter-bord

Et enkelt arbeidsbord for å oversette en bok fra **engelsk til norsk**. Du laster
opp den engelske originalen som fil og ser den til venstre; den norske
oversettelsen skriver eller limer du inn til høyre, avsnitt for avsnitt.

## Slik bruker du den

1. **Last opp engelsk original** (Word `.docx` eller tekst `.txt`/`.md`) – den
   havner på venstre side.
2. **Skriv eller lim inn norsk** på høyre side – f.eks. fra DeepL eller ChatGPT.
   Du kan også laste opp en hel norsk oversettelse som fil, så legges den på rad.
3. **Klikk et ord** for å se det tilsvarende avsnittet lyse opp på den andre
   siden (begge veier). Klikk et ord på hver side for å **koble** dem sammen.
4. **Usikker på et ord?** Klikk det og «Merk usikker». Samlede ord kan kopieres
   til ChatGPT/Claude/Gemini, eller slås opp direkte i appen med din egen nøkkel.

Alt lagres automatisk i din egen nettleser. Under «≡ Fil» kan du ta en
arbeidsfil (sikkerhetskopi/deling), hente den fram igjen, og eksportere den
ferdige norske teksten.

## Bruker og prosjekter i skyen (valgfritt)

Logg inn med **👤-knappen** (krever invitasjonskode første gang) for å lagre
arbeidet som **prosjekter i skyen**: da ser du de siste prosjektene dine, kan
fortsette på en annen maskin, og endringer skylagres automatisk mens du jobber.
Hver bruker ser kun sine egne prosjekter. Uten innlogging virker alt som før,
lagret lokalt i nettleseren.

## Innebygd DeepL (valgfritt)

Innlogget kan du oversette et helt kapittel med **DeepL** fra
«🤖 Forbedre med AI»-vinduet. Legg inn din egen DeepL-nøkkel under ⚙︎
(gratis: deepl.com → «DeepL API» → Free-plan, 500 000 tegn/mnd). Nøkkelen
lagres bare i din nettleser; appens tjener videresender bare forespørselen
(DeepL tillater ikke direktekall fra nettlesere).

## AI-funksjoner (valgfritt – krever egen nøkkel)

Du kan koble til **Claude (Anthropic)**, **ChatGPT (OpenAI)** eller **Gemini
(Google)** ved å legge inn din egen API-nøkkel under ⚙︎. Da kan appen:

- **🔍 Slå opp ord** – sender avsnittet ordet står i (engelsk + din norske tekst).
- **Koble ord automatisk** – når du klikker et ord, sendes det ene avsnittet
  (engelsk + norsk) for å finne hvilke ord som hører sammen. Kan slås av under ⚙︎.
- **🤖 Forbedre med AI / fordel kapittel** – sender **hele kapittelet** (engelsk
  + norsk) for å pusse på eller fordele oversettelsen.

Du velger selv om/når du bruker disse. Uten nøkkel kan du i stedet kopiere
teksten og lime inn i en vanlig AI-chat.

> **Hva sendes hvor:** Bare det du selv utløser med en AI-knapp sendes – til den
> AI-tjenesten du har valgt, med din egen nøkkel. Resten av arbeidet skjer lokalt.
>
> **Nøkler:** API-nøkler lagres i nettleserens lagring (localStorage) på denne
> maskinen. Det er greit for et personlig verktøy; del ikke maskinen/nettleseren
> med andre du ikke stoler på mens nøklene ligger der.
>
> **DeepL** kan ikke kobles direkte fra en nettleser-app (DeepL blokkerer det).

## Personvern og innhold

Appen er bare et verktøy – den inneholder **ingen bok**. Teksten du laster opp
blir liggende kun i din egen nettleser. Den sendes ingen steder med mindre du
selv trykker på en AI-knapp (se over). Last ikke opp opphavsrettsbeskyttet
materiale til en offentlig kopi du deler videre.

---

*Teknisk: ren HTML/CSS/JS uten avhengigheter eller byggesteg. Word-filer (.docx)
pakkes ut og leses direkte i nettleseren – vanlig brødtekst og overskrifter leses
godt, men tabeller, tekstbokser og fotnoter kan falle ut. Lim heller inn teksten
hvis et dokument er spesielt komplisert.*
