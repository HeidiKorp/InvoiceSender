# Arvete Saatja

Töölauaprogramm korterite arvete saatmiseks e-postiga. Programm tükeldab koondarve faili, seob arved klientidega korteri numbri järgi, salvestab iga korteri kohta ühe PDF-i ja koostab Outlookis mustandid.

In English: [README.md](README.md)

## Mida programm teeb

1. Vali arvete tüüp: **Kommunaalarved** või **Küttearved**. Tüüp määrab peamiselt meili teema ja sisu malli.
2. Vali arvete fail ja klientide fail.
   - Kommunaalarved: PDF
   - Küttearved: Excel (`.xls`, `.xlsx`)
   - Kliendid: Excel (`.xls`, `.xlsx`, `.xlsm`)
3. Arved tükeldakse ja seotakse klientidega **korteri numbri järgi**. Edenemisriba näitab, et töö käib. Kommunaalarvete lehti ei nimetata ümber, et numbrid oleksid järjest. Ühe- või kahekohaline lünk arvete numbriseerias on ainult hoiatus (`Arvete järjestuses puuduvad numbrid`); see ei tähenda, et päris korteril arve puuduks. Tühi lugem, tagurpidi number või keset jada hüpe kuvatakse kui `Kontrolli üle arved`. Korterid, millel arve puudub, või arved, millel klienti ei ole, jäetakse kõrvale ja kuvatakse viga.
4. Sobitatud arved salvestatakse kausta `arved/<aadress>/<periood>/` arvete faili kõrvale (iga korteri kohta `{korter}.pdf`). Kui sama korter loetakse ühe käivituse jooksul mitmelt kommunaalarve lehelt, jääb esimene leht failina `{korter}.pdf` ja seda saab endiselt meilida; järgmised salvestatakse kui `{korter}-kontrolli.pdf`, seejärel `{korter}-kontrolli-2.pdf` jne, koos teatega, et need failinimed tuleb üle kontrollida. Sama perioodi uuesti töötlemine võib need failid asendada. Kommunaalarve leht, millelt korteri numbrit ei loeta, salvestatakse samasse kausta failina `nimetu-1.pdf`, seejärel `nimetu-2.pdf` ja nii edasi. Nimeta fail ümber korteri numbriks (näiteks `57.pdf`). See korter kuvatakse endiselt real `Puuduvad arved korteritele`.
5. Avane meilimalli aken. Vaikimisi teema on `{address} arve {period} {year}`, näiteks `Õismäe tee 48 arve august 2026`. `{period}` peab olema eesti kuu nimi ja `{year}` aasta 2001–2999; kui ühe arve lehelt neid ei loeta, võetakse andmed teistelt sama faili arvetelt. Teemat ja sisu saab muuta ning malli `.cfg` failina laadida või salvestada.
6. Outlookis (peab olema paigaldatud ja sisse logitud) luuakse mustandid kategooriaga `ArveteSaatja`, vastava(te) kliendi meiliaadressi(de)ga ja korteri PDF manuses. Kui ühel kliendireal on mitu aadressi, saab igaüks oma mustandi. Outlookis peaks avanema kaust **Mustandid**.
7. **Saada mustandid** saadab ainult selle kategooria mustandid.

**Avalehele** tühjendab arvete tüübi ja mõlemad failid ning viib tagasi esimesele kuvale, et saaks töödelda järgmist komplekti. Kui arvete lugemine või mustandite koostamine veel käib, küsib programm enne katkestamist.

Vea korral naaseb aken eelmisesse kasutatavasse olekusse (failivalikud jäävad alles; katkestamine taastab ooteriba). Ootamatud vead kirjutatakse faili `utils/error.log` (või `.exe` kõrvale, kui programm on pakendatud), koos kellaaja, tegevuse ja veateatega.

## Nõuded

- Windows
- Classic Outlook, sisse logitud
- Tesseract OCR eesti keele andmetega (`est`) — vajalik kommunaal-PDF-ide jaoks
- Excel — vajalik küttearvete jaoks
- Python 3 ja projekti teegid (`pandas`, `openpyxl` `.xlsx`/`.xlsm` klientide jaoks, `xlrd` `.xls` jaoks, `pytesseract`, `pypdf`, PyMuPDF, `pywin32`, `ttkbootstrap`)

## Kasutamine

```
python -m run_app
```

või:

```
python run_app.py
```

Kui programm on zip-failina, paki see sobivasse kausta lahti ja ava `.exe`. Kaustas `_internal` on programmi tööks vajalikud failid; neid tavaliselt muutma ei pea.

Exe koostamine:

```
rm -rf build dist *.spec
pyinstaller --onedir --noconsole --name ArveteSaatja --paths . --add-data "tesseract:tesseract" --add-data "config.cfg:." run_app.py
```

## Testid

Paigalda `pytest` ja käivita projektikaustast:

```
python -m pytest test
```

### Sammud programmis

1. Ava programm.
2. Vali arvete tüüp — **Kommunaalarved** või **Küttearved**.
3. Vali arvete fail (PDF kommunaali, Excel kütte puhul) ja klientide fail. Valitud faili tee kuvatakse nupu kõrval.
4. Vajuta **Koosta meilid**. Töötlemist saab peatada nupuga **Katkesta**. **Avalehele** (peaaknas ja meiliaknas) tühjendab vormi; kui töö veel käib, küsib programm enne katkestamist kinnitust.
5. Iga korteri arve salvestatakse kausta `arvete_faili_kaust/arved/aadress/periood`. Kui kommunaalarve lehelt korteri numbrit ei loeta, salvestatakse see samasse kausta nimega `nimetu-1.pdf`. Nimeta fail ise ümber korteri numbriks. Korter jääb teatesse `Puuduvad arved korteritele`.
6. Soovi korral saab selle perioodi kausta kustutada nupuga **Kustuta arvekaust** (kustutatakse ainult selle perioodi kaust, mitte kogu maja kaust).
7. Kontrolli meili teemat ja sisu. Kõikidele korteritele läheb sama teema ja sisu. Kui kõik sobib, vajuta **Salvesta** — programm avab Outlooki ja koostab mustandid.
   - Teemat ja sisu saab muuta aknas ja salvestada.
   - Mallid saab salvestada `.cfg` failina nupuga **Salvesta mall** ja hiljem laadida nupuga **Laadi mall**.
8. Kui mõnele arvele ei leitud klienti või mõnele kliendile ei leitud arvet, kuvatakse sellest teade.
9. Outlookis on mustandid kategooriaga `ArveteSaatja`. Kontrolli need üle. Kui kõik on korras, vajuta programmis **Saada mustandid**. Saadetakse ainult selle kategooria mustandid.

## Klientide fail

Nõutavad veerud: `klient_mail`, `korter`, `yhistu`, `maj_nr`.

- `korter` võib olla number ja üks täht (`44`, `44 a`, `44A`). Täht salvestatakse kujul `44A`.
- `klient_mail` on kohustuslik; ühel real võib olla mitu aadressi, eraldatud `,` või `;` märgiga

Sobitamine käib ainult korteri numbri järgi, mitte meiliaadressi ega tänava järgi.

## Meilimallid

Vaikimisi mallid on failis `config.cfg`. Kohatäited: `{address}`, `{period}`, `{year}`, `{apartment}`, `{ky_name}`. Teema täidetakse arvetest, millel on kehtiv aadress, eesti kuu nimi ja aasta (2001–2999); vigane esimene leht ei jäta pealkirja sõna `periood`, kui teisel arvel on `august`.

Mõlema arvete tüübi teema on:

```
SUBJECT={address} arve {period} {year}
```

Looklevate sulgude vahel on programmis täidetavad väljad — nende nimesid ise ära muuda. Ülejäänud teksti võib vabalt muuta. Lubatud on näiteks:

```
Uus arve {year} {period}
```

## Gmail-konto seadistamine Outlookis

- Paigalda classic Outlook
    - https://support.microsoft.com/en-us/office/install-or-reinstall-classic-outlook-on-a-windows-pc-5c94902b-31a5-4274-abb0-b07f4661edf5
- Ava **Control Panel** ja otsi "Mail (Microsoft Outlook)"
- See avab viisardi, kus saab kontosid hallata ja lisada.
- Konto lisamisel vali käsitsi lisamine (mitte Microsoft 365 konto).
- Kui Gmailil on kaheastmeline kinnitus, loo Outlooki jaoks rakenduse parool:
    - Ava Gmail
    - Ava Google Account Security (profiilist)
    - Jaotisest "Signing in to Google" -> App Passwords (või otsi seda)
    - Loo parool *Mail / Outlook* jaoks
    - Salvesta parool, eemalda tühikud ja kleebi see Outlooki parooliväljale
    - Konto tüüp: IMAP
    - Sisenev server: `imap.gmail.com`
    - Väljuv server: `smtp.gmail.com`
- Kasutajanimeks pane täielik Gmaili aadress, näiteks `nimi@gmail.com`
- All paremal klõpsa **More settings...**
- Kaardil **Outgoing Server**:
    - Märgi "My outgoing server (SMTP) requires authentication"
    - Vali "Use same settings as my incoming mail server"
- Kaardil **Advanced**:
    - Sisenev server (IMAP): **993**, krüpteering **SSL/TLS**
    - Väljuv server (SMTP): **587**, krüpteering **STARTTLS**
- Kui *More settings...* vahele jätta, proovib Outlook porti 25 ilma krüpteeringuta ja Gmail keeldub veaga `530 5.7.0 Authentication Required`

## Tõrgete korral

Vealogi salvestatakse faili `error.log` (programmi juurkausta või `.exe` kõrvale). Küsimuste korral saada see logi arendajale.

Küttearvete tüüpiline viga: Excel on juba avatud enne töötlemist. Sulge Excel tegumihaldurist (otsi "Excel" → peata protsess) ja proovi uuesti.
