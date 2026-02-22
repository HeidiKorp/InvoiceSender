============================================================
ArveteSaatja
Kasutusjuhend
Versioon 2.3.6
============================================================

1. Sissejuhatus
------------------------------------------------------------
ArveteSaatja eesmärk on hõlbustada arvete raamatupidaja tööd
igakuiste kommunaalide ja küttearvete saatmisel.


2. Eeldused
------------------------------------------------------------
Programmi kasutamiseks on kaks eeldust:

1. On olemas arvete fail, mis koosneb mitmest arvest - 
üks leht korteri kohta. Kommunaalide arved on salvestatud 
PDF-formaadis ning küttearved XLS-formaadis.

2. On olemas XLS-fail klientide meiliaadressitega. Iga 
korteri kohta peab olema dokumendis üks rida koos 
meiliaadressiga. Ühes korteril võib olla rohkem kui üks
meiliaadress.


3. Installeerimine
------------------------------------------------------------
1. Programm on pakitud zip-faili, mis tuleb omale sobivas 
kaustas lahti pakkida (parem klikk -> paki lahti)

2. Kaustas _internal asuvad programmi töötamiseks vajalikud
failid. Üldjuhul ei pea nendega midagi tegema. 


4. Kasutamine
------------------------------------------------------------
    1. Avage EXE fail
    2. Valige arve tüüp - kas "Kommunaalarved" või "Küttearved"
    3. Vastavast nupust valige 
        - Arvete fail (PDF kommunaali puhul, XLS küttearve puhul)
        - Klientide fail (XLS fail)
    Valitud failide tee kuvatakse nupust paremal.
    4. Vajutage "Koosta meilid"
        - Failide töötlemine algab. 
        - Failide lugemise protsessi on võimalik peatada nupust 
            "Katkesta"
        - Kui protsess on katkestatud, on võimalik vajadusel 
        valida uued failid.
    5. Kui failid on töödeldud, salvestatakse korterite arved
    uude kausta, mille teekond on:

    - arvete_faili_kaust/aadress/periood

    6. Kui arved on salvestatud, on võimalik see kaust kustutada.
    Selleks klikkige "Kustuta arvekaust". See kustutab ainult
    antud perioodi kausta, mitte tervet maja kausta.

    7. Kui arved on loetud ja salvestatud uude kausta, on kasutajal
    võimalik üle vaadata meili teema ja sisu.
    - Kõikidele korteritele saadetakse meil sama teema ja sisuga.
    - Vaikimisi sisaldab meili teema perioodi ja aastat ning sisu
    kas korteriühistu nime või küttearvete puhul aadressi.

    Kui kõik sobib, siis vajutage "Salvesta" ning programm avab 
    Outlook programmi ning genereerib meilide mustandid.

    Juhul kui kasutaja soovib meili teemat või sisu muuta:
    - Saab ta seda teha lihtsalt avanenud aknas ja salvestada
    - Muuta tekst ning salvestada see konfiguratsioonifailina,
    et seda ka tulevikus kasutada. See salvestatakse .cfg
    failina vabalt valitud kausta "Salvesta mall" nupuga
    ning varem salvestatud konfiguratsiooni saab otsida 
    nupuga "Laadi mall".

    8. Juhul kui mõnele arvele ei leitud vastavat klienti
    või mõnele kliendile ei leitud vastavat arvet, antakse 
    kasutajale sellest teada.

    9. Viimaks avab programm Outlook programmi.
    Kõik genereeritud meilid salvestatakse esialgu mustandite
    kausta ning kategooriaks märgitakse "ArveteSaatja".
    Kasutaja saab vajadusel meilid üle kontrollida ja 
    vajadusel parandusi teha. 

    Kui kõik on korras, siis vajutage ArveteSaatja programmist
    nupule "Saada mustandid" ning arved lähevad teele.
    Pidage silmas, et programm saadab ainult need mustandid, 
    mille kategooriaks on märgitud "ArveteSaatja". Teisi 
    mustandeid ei saadeta automaatselt.


5. Tõrgete korral
------------------------------------------------------------
Tekkinud vead salvestatakse faili "error.log", mis 
salvestatakse ArveteSaatja juurkausta. Kui nende logidega
tekib küsimusi, palun pöörduda arendaja poole.

Üks tüüpiline viga, mis võib tekkida küttearvete töötlemisega
on see, kui Exceli programm on juba lahti enne arvete 
töötlemist. Sel juhul palun peatada Exceli töö:
- Tegumihaldur -> Otsida "Excel" -> parem klikk -> peata protsess

Seejärel alustada küttearvete töötlemist uuesti.


6. Lisa 
------------------------------------------------------------

_internal/config.cfg on konfiguratsioonifail, kus on 
kirjeldatud vaikimisi meili teema ja sisu. Kui on soovi
seda muuta, siis saab seda teha nii:
- Loogeliste sulgude vahel on parameeter, millele antakse
väärtus programmi sees. Seda palun ise mitte muuta! Ülejäänud
teksti muutmisel probleemi pole.

Näidis:
Arve {period} {year}

Siin näites on "Arve" tavaline tekst, mida saab vabalt muuta.
"period" ja "year" on aga muutujad, mille järjekorda saab 
muuta, kuid sisu mitte. Ehk see on lubatud:

Uus arve {year} {period}


