#!/usr/bin/env python3
"""
Võrdleb kahte Oksjonikeskuse töötamise registri vastust (nt eilset ja tänast)
ja väljastab toimikud, kus võlgniku töötamistes on vahepeal midagi muutunud.
Väljund on samas vormis nagu tookoha_muutused.py oma: vana ja uus töökoht
kõrvuti (ka siis, kui tööandja on sama), lisaks võlgniku nimi ja lõpus muutuse liik.

Kasutus:
    python oksjonikeskus_vordlus.py vana.csv uus.csv [valjund.csv]

Failid võivad olla ka Exceli failid (.xlsx vajab: pip install openpyxl).

Vajab samas kaustas faili tookoha_muutused.py (sealt tulevad CSV lugemine ja abifunktsioonid).
"""
import csv
import re
import sys

from tookoha_muutused import hiljutised_tooandjad, loe_csv, norm_toimik, viimased_kanded

# Väljad, mille muutumist jälgitakse. Viimase muutmise kuupäeva ei võrrelda,
# sest see võib muutuda ka siis, kui sisuliselt midagi ei muutunud.
JALGITAVAD = ["Kande olek", "Töötamise alguse aeg", "Töötamise lõpp",
              "Töötamise peatamise algus", "Töötamise peatamise lõpp"]

# Väljundis enne tulevad muutused, mille peale on tõenäoliselt vaja midagi teha
JARJEKORD = ["Uus töötamine", "Töötamine lõppes", "Töötamine peatati", "Peatamine lõppes",
             "Töötamine algas", "Olek muutus", "Lõpukuupäev muutus", "Andmed muutusid",
             "Kadus vastusest"]


def loe_kanded(read):
    """Tagastab (toimikud, kanded).
    toimikud: kõik failis olevad toimikud, ka need, kus töötamisi ei leitud.
    kanded: (toimik, TÖR ID) -> viimati muudetud rida."""
    toimikud = {norm_toimik(r["Toimik"]) for r in read if r.get("Toimik")}
    tootamised = [r for r in read
                  if r.get("TÖR ID") and r.get("Tööandja registrikood või ID")
                  # tööandja on võlgnik ise (FIE), töötasu arestimise akti pole kellele saata
                  and r["Tööandja registrikood või ID"] != r.get("Isikukood")]
    kanded = {(norm_toimik(r["Toimik"]), r["TÖR ID"]): r for r in viimased_kanded(tootamised)}
    return toimikud, kanded


def muutuse_liik(vana, uus):
    if vana is None:
        return "Uus töötamine"
    if uus is None:
        return "Kadus vastusest"
    vo = vana.get("Kande olek", "").lower()
    uo = uus.get("Kande olek", "").lower()
    if vo != uo:
        if uo == "lõpetatud":
            return "Töötamine lõppes"
        if uo == "peatatud":
            return "Töötamine peatati"
        if vo == "peatatud" and uo == "kehtiv":
            return "Peatamine lõppes"
        if vo == "tulevikus" and uo == "kehtiv":
            return "Töötamine algas"
        return "Olek muutus"
    if vana.get("Töötamise lõpp", "") != uus.get("Töötamise lõpp", ""):
        return "Lõpukuupäev muutus"
    if any(vana.get(v, "") != uus.get(v, "") for v in JALGITAVAD):
        return "Andmed muutusid"
    return None


def aktuaalsed(andjad):
    """Sama loogika nagu tookoha_muutused.py-s: kehtivad tööandjad, nende puudumisel hiljuti lõppenud."""
    return [a for a in andjad if a["kehtiv"]] or andjad


def viimati(andjad):
    if not andjad:
        return ""
    a = andjad[0]
    if a["kehtiv"]:
        return "peatatud" if a["peatatud"] else "kehtiv"
    return a["viimati"].strftime("%d.%m.%Y")


def toimiku_jarjestus(toimik):
    """'999/2000/240' tuleb enne '999/2000/1035' (numbrite, mitte teksti järgi)."""
    return [int(o) if o.isdigit() else 0 for o in toimik.split("/")]


def lyhike_loend(toimikud, n=5):
    jarjestatud = sorted(toimikud, key=toimiku_jarjestus)
    tekst = ", ".join(jarjestatud[:n])
    return tekst + (f" ja veel {len(jarjestatud) - n}" if len(jarjestatud) > n else "")


def main(vana_fail, uus_fail, valjund):
    vanad_read = list(loe_csv(vana_fail))
    uued_read = list(loe_csv(uus_fail))
    vanad_toimikud, vanad = loe_kanded(vanad_read)
    uued_toimikud, uued = loe_kanded(uued_read)

    # Võrreldakse ainult toimikuid, mis on mõlemas failis. Muidu näiks iga
    # toimik, mida eile ei päritud, täis "uusi" töötamisi.
    yhised = vanad_toimikud & uued_toimikud
    ainult_uues = uued_toimikud - vanad_toimikud
    ainult_vanas = vanad_toimikud - uued_toimikud
    if ainult_uues:
        print(f"Hoiatus: {len(ainult_uues)} toimikut on ainult uues failis ja neid ei võrreldud "
              f"({lyhike_loend(ainult_uues)})", file=sys.stderr)
    if ainult_vanas:
        print(f"Hoiatus: {len(ainult_vanas)} toimikut on ainult vanas failis ja neid ei võrreldud "
              f"({lyhike_loend(ainult_vanas)})", file=sys.stderr)

    # toimik -> {"toimik", "isikukood", "liigid"}
    muutunud = {}
    for voti in vanad.keys() | uued.keys():
        if voti[0] not in yhised:
            continue
        liik = muutuse_liik(vanad.get(voti), uued.get(voti))
        if liik is None:
            continue
        r = uued.get(voti) or vanad[voti]
        m = muutunud.setdefault(voti[0], {
            "toimik": re.sub(r"\s", "", r["Toimik"]),
            "isikukood": r.get("Isikukood", ""),
            "nimi": f'{r.get("Eesnimi", "")} {r.get("Perekonnanimi", "")}'.strip(),
            "liigid": set(),
        })
        m["liigid"].add(liik)

    vanad_andjad = hiljutised_tooandjad(vanad_read)
    uued_andjad = hiljutised_tooandjad(uued_read)

    tulemus = []
    for toimik, m in muutunud.items():
        vana = aktuaalsed(vanad_andjad.get(toimik, []))
        uus = aktuaalsed(uued_andjad.get(toimik, []))
        liigid = sorted(m["liigid"], key=JARJEKORD.index)
        # Kui töötamisi on mitu, on kõik tööandjad samas lahtris " / " vahega
        tulemus.append((JARJEKORD.index(liigid[0]), toimiku_jarjestus(m["toimik"]), [
            m["toimik"],
            m["isikukood"],
            m["nimi"],
            " / ".join(a["nimi"] for a in vana),
            " / ".join(a["kood"] for a in vana),
            " / ".join(a["nimi"] for a in uus),
            " / ".join(a["kood"] for a in uus),
            viimati(uus),
            ", ".join(liigid),
        ]))
    tulemus.sort(key=lambda t: t[:2])

    with open(valjund, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Toimik", "Isikukood", "Nimi", "Vana töökoht", "Vana reg nr",
                    "Uus töökoht", "Uus reg nr", "Viimati töötas", "Muutus"])
        w.writerows(t[2] for t in tulemus)
    print(f"{len(tulemus)} muutunud toimikut -> {valjund}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "oksjonikeskus_muutused.csv")