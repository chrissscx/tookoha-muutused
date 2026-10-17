#!/usr/bin/env python3
"""
Võrdleb Taavi toimikute XML eksporti Oksjonikeskuse töötamise registri vastusega
ja väljastab ainult need toimikud, kus võlgniku töökoht on muutunud.

Kasutus:
    python tookoha_muutused.py toimikud.xml oksjonikeskus.csv [valjund.csv]
"""
import csv
import re
import sys
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta

# Töötamine loetakse hiljutiseks, kui see kehtib või lõppes viimase N päeva jooksul.
# Vajalik lühikeste lepingute pärast (nt GoWorkaBit), kus iga tööpäev on eraldi kanne.
HILJUTINE_PAEVI = 30

# Tühi olek tähendab enamasti, et töötamisi ei leitud (rida jäetakse vahele, kuna tööandja puudub).
# Kui tööandja on olemas, aga olek tühi, otsustatakse kuupäevade järgi.
# Peatatud = leping kehtib, töö on peatatud (nt lapsehoolduspuhkus), tööandja sama.
TEADAOLEVAD_OLEKUD = {"", "kehtiv", "tulevikus", "lõpetatud", "peatatud"}

# Õiguslikud vormid, mis nime võrdlemisel ära visatakse
VORMID = r"\b(oü|as|mtü|sa|tü|uü|fie|osaühing|aktsiaselts|mittetulundusühing|sihtasutus)\b"


def norm_toimik(nr):
    """'100/2001/     1' ja '100/2001/1' -> '100/2001/1'"""
    osad = [p.strip().lstrip("0") or "0" for p in nr.split("/")]
    return "/".join(osad)


def norm_nimi(nimi):
    s = nimi.lower()
    s = re.sub(VORMID, " ", s)
    s = re.sub(r"[^\w]", " ", s)
    return " ".join(s.split())


def kood_tekstist(tekst):
    """Registrikood (8 numbrit) või eraisikust tööandja isikukood (11 numbrit)."""
    m = re.search(r"\b(\d{8}|\d{11})\b", tekst or "")
    return m.group(0) if m else ""


def loe_taavi(xml_fail):
    """toimiku nr -> {toimik, isikukood, nimi, kood}. Loeb ainult TOIMIKUD ridu."""
    tulemus = {}
    for _, el in ET.iterparse(xml_fail, events=("end",)):
        if el.tag != "dbf_fail":
            continue
        if el.get("nimi") == "TOIMIKUD":
            rida = el.find("rida")
            if rida is not None:
                v = lambda tag: (rida.findtext(tag) or "").strip()
                tookoht = v("VTOOKOHT")
                tulemus[norm_toimik(v("NUMBER"))] = {
                    "toimik": re.sub(r"\s", "", v("NUMBER")),
                    "isikukood": v("VOLGNIK"),
                    "nimi": tookoht,
                    "kood": kood_tekstist(tookoht),
                }
        el.clear()  # suur fail, ära hoia mälus
    return tulemus


def parse_kp(s):
    """'02.10.2026' või '2026-10-02T01:30:04.000+03:00' -> datetime"""
    s = (s or "").strip()[:19]
    for fmt in ("%d.%m.%Y", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def loe_csv(csv_fail):
    for enc in ("utf-8-sig", "cp1257"):
        try:
            with open(csv_fail, encoding=enc, newline="") as f:
                sisu = f.read()
            break
        except UnicodeDecodeError:
            continue
    murre = csv.Sniffer().sniff(sisu[:5000], delimiters=";\t,")
    for r in csv.DictReader(sisu.splitlines(), dialect=murre):
        yield {k.strip(): (v or "").strip() for k, v in r.items() if k}


def viimased_kanded(read):
    """Sama TÖR ID kohta tuleb mitu rida (Tulevikus -> Kehtiv -> Lõpetatud).
    Jätab igast kandest alles viimati muudetud rea."""
    kanded = {}
    for r in read:
        kp = parse_kp(r.get("Viimase muutmise kuupäev")) or datetime.min
        voti = (norm_toimik(r["Toimik"]), r["TÖR ID"])
        if voti not in kanded or kp > kanded[voti][0]:
            kanded[voti] = (kp, r)
    return [r for _, r in kanded.values()]


def hiljutised_tooandjad(read):
    """toimiku nr -> hiljutised tööandjad (üks kord iga tööandja kohta), viimati töötanud ees."""
    tana = datetime.combine(date.today(), datetime.min.time())
    piir = tana - timedelta(days=HILJUTINE_PAEVI)
    tundmatud = set()
    tulemus = {}

    for r in viimased_kanded(read):
        kood = r.get("Tööandja registrikood või ID", "")
        if not kood:
            continue  # töötamisi ei leitud (tühi rida, ainult isikukood ja toimik)
        if kood == r.get("Isikukood", ""):
            continue  # tööandja on võlgnik ise (FIE), töötasu arestimise akti pole kellele saata
        olek = r.get("Kande olek", "").lower()
        if olek not in TEADAOLEVAD_OLEKUD:
            tundmatud.add(olek)
            continue
        algus = parse_kp(r.get("Töötamise alguse aeg"))
        lopp = parse_kp(r.get("Töötamise lõpp"))
        if algus and algus > tana:
            continue  # pole veel alanud
        if lopp and lopp < piir:
            continue  # lõppes ammu
        kehtiv = lopp is None or lopp >= tana
        viimati = tana if kehtiv else lopp

        andjad = tulemus.setdefault(norm_toimik(r["Toimik"]), {})
        if kood not in andjad or viimati > andjad[kood]["viimati"]:
            andjad[kood] = {
                "kood": kood,
                "nimi": r.get("Tööandja /Tegeliku tööandja nimetus", ""),
                "isikukood": r.get("Isikukood", ""),
                "algus": algus or datetime.min,
                "viimati": viimati,
                "kehtiv": kehtiv,
                "peatatud": olek == "peatatud",
            }

    if tundmatud:
        print(f"Hoiatus: vahele jäetud tundmatu kande olekuga read: {tundmatud}", file=sys.stderr)
    # viimati töötanud ees, võrdsuse korral hiljem alustanud ees
    return {t: sorted(a.values(), key=lambda x: (x["viimati"], x["algus"]), reverse=True)
            for t, a in tulemus.items()}


def sama_tookoht(vana, uus):
    if vana["kood"]:
        return vana["kood"] == uus["kood"]
    return bool(vana["nimi"]) and norm_nimi(vana["nimi"]) == norm_nimi(uus["nimi"])


def main(xml_fail, csv_fail, valjund):
    taavi = loe_taavi(xml_fail)
    uued = hiljutised_tooandjad(loe_csv(csv_fail))

    muutunud = []
    for toimik, andjad in uued.items():
        vana = taavi.get(toimik)
        if vana is None:
            print(f"Hoiatus: toimikut {toimik} pole Taavi ekspordis", file=sys.stderr)
            continue
        # Kui on kehtiv töötamine, võrdle ainult kehtivatega. Hiljuti lõppenud
        # arvestatakse ainult siis, kui kehtivat pole (lühiajalised lepingud).
        vordle = [a for a in andjad if a["kehtiv"]] or andjad
        if any(sama_tookoht(vana, a) for a in vordle):
            continue  # Taavis olev tööandja on endiselt aktuaalne
        uus = vordle[0]
        muutunud.append([
            vana["toimik"],
            uus["isikukood"] or vana["isikukood"],
            vana["nimi"],
            vana["kood"],
            uus["nimi"],
            uus["kood"],
            ("peatatud" if uus["peatatud"] else "kehtiv") if uus["kehtiv"]
            else uus["viimati"].strftime("%d.%m.%Y"),
        ])

    with open(valjund, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Toimik", "Isikukood", "Vana töökoht", "Vana reg nr",
                    "Uus töökoht", "Uus reg nr", "Viimati töötas"])
        w.writerows(muutunud)
    print(f"{len(muutunud)} muutunud töökohta -> {valjund}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "muutunud_tookohad.csv")
