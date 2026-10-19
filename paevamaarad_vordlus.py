#!/usr/bin/env python3
"""
Võrdleb kahte päevamäärade tabelit (nt eilset ja tänast) ja väljastab kõik read, mis erinevad.
Veerus "Muutus" on kirjas, kus rida on: "Ainult vanas", "Ainult uues" või "Muutunud"
(sama isik, tüüp ja perioodi algus mõlemas failis, aga muud väärtused erinevad).

Kasutus:
    python paevamaarad_vordlus.py vana.csv uus.csv [valjund.csv]

Failid võivad olla CSV või Exceli failid (.xlsx vajab: pip install openpyxl).

Vajab samas kaustas faili tookoha_muutused.py (sealt tuleb tabeli lugemine).
"""
import csv
import re
import sys
import unicodedata
from datetime import datetime
from decimal import Decimal

from tookoha_muutused import loe_csv, parse_kp

VEERUD = ["Isikukood", "Tüüp", "Perioodi algus", "Perioodi lõpp",
          "Otsus Paevamäär #1", "Otsus Paevamäär #1 (sentides)",
          "Otsus Paevamäär #2", "Otsus Paevamäär #2 (sentides)"]

# Kui need klapivad, aga midagi muud on teisiti, on see sama rida, mis muutus.
# Muidu on vana rida "Ainult vanas" ja uus "Ainult uues".
VOTI = ["Isikukood", "Tüüp", "Perioodi algus"]

JARJEKORD = ["Ainult vanas", "Muutunud", "Ainult uues"]

NUMBER = re.compile(r"-?\d[\d ]*([.,]\d+)?")


def veeru_voti(nimi):
    """'Otsus Päevamäär #2  (sentides)' ja 'otsus paevamaar #2 (sentides)' -> sama võti"""
    s = unicodedata.normalize("NFKD", nimi).encode("ascii", "ignore").decode()
    return " ".join(s.lower().split())


def loe_tabel(fail):
    """Tagastab (päis, read, tundmatud veerud).
    read: iga rida list VEERUD järjekorras. päis: need veerud nii, nagu nad failis kirjas on."""
    read = list(loe_csv(fail, voti="Isikukood"))
    failis = {veeru_voti(k): k for k in (read[0] if read else {})}
    meie = {veeru_voti(v) for v in VEERUD}
    puudu = [v for v in VEERUD if veeru_voti(v) not in failis]
    if read and puudu:
        raise ValueError(f"Failis pole veerge: {', '.join(puudu)}")
    pais = [failis.get(veeru_voti(v), v) for v in VEERUD]
    tundmatud = [k for v, k in failis.items() if v not in meie]
    return pais, [[r[failis[veeru_voti(v)]] for v in VEERUD] for r in read], tundmatud


def norm(v):
    """Võrdlemiseks: '1.10.2026' == '2026-10-01', '12,30' == '12.3' (CSV vs Excel)."""
    v = " ".join(v.split())
    kp = parse_kp(v)
    if kp:
        return kp.strftime("%d.%m.%Y")
    if NUMBER.fullmatch(v):
        return format(Decimal(v.replace(" ", "").replace(",", ".")).normalize(), "f")
    return v


def mis_muutus(vana, uus):
    muutunud = [i for i, (a, b) in enumerate(zip(vana, uus)) if norm(a) != norm(b)]
    nimed = {VEERUD[i] for i in muutunud}
    return "; ".join(
        f"{VEERUD[i]}: {vana[i] or 'tühi'} → {uus[i] or 'tühi'}" for i in muutunud
        # sentides veerg on sama info, kui euro veerg juba muutus
        if not (VEERUD[i].endswith(" (sentides)") and VEERUD[i][:-11] in nimed))


def main(vana_fail, uus_fail, valjund):
    _, vanad, _ = loe_tabel(vana_fail)
    _, uued, _ = loe_tabel(uus_fail)

    def ruhmad(read):
        tulemus = {}
        for r in read:
            tulemus.setdefault(tuple(map(norm, r)), []).append(r)
        return tulemus

    # Täpselt samad read on mõlemas ja jäävad välja. Kui sama rida on ühes failis
    # mitu korda, loeb ka arv (2 vanas, 3 uues -> 1 rida "Ainult uues").
    # Ka tühjad read (ainult isikukood, päevamäärasid ei leitud) on tavalised read.
    vanad_r, uued_r = ruhmad(vanad), ruhmad(uued)
    lisandunud = [r for k, rr in uued_r.items() for r in rr[len(vanad_r.get(k, [])):]]
    kadunud = [r for k, rr in vanad_r.items() for r in rr[len(uued_r.get(k, [])):]]

    voti = lambda r: tuple(norm(r[VEERUD.index(v)]) for v in VOTI)
    ootel = {}
    for r in kadunud:
        ootel.setdefault(voti(r), []).append(r)

    tulemus = []
    for r in lisandunud:
        paar = ootel.get(voti(r))
        if paar:
            tulemus.append(("Muutunud", r, mis_muutus(paar.pop(0), r)))
        else:
            tulemus.append(("Ainult uues", r, ""))
    tulemus += [("Ainult vanas", r, "") for rr in ootel.values() for r in rr]
    # Sama isiku read üksteise all, et vana ja uut oleks lihtne kõrvutada
    tulemus.sort(key=lambda t: (t[1][0], parse_kp(t[1][2]) or datetime.min, JARJEKORD.index(t[0])))

    with open(valjund, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Muutus"] + VEERUD + ["Mis muutus"])
        w.writerows([liik] + r + [mis] for liik, r, mis in tulemus)
    loend = ", ".join(f"{liik.lower()} {sum(t[0] == liik for t in tulemus)}" for liik in JARJEKORD)
    print(f"{len(tulemus)} erinevat rida ({loend}) -> {valjund}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "paevamaarad_muutused.csv")