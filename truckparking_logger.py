#!/usr/bin/env python3
"""
Truckparking logger Maasvlakte Plaza (Noord en Zuid)

Haalt de actuele bezetting op uit NDW open data en voegt per run
een regel per locatie toe aan een CSV-bestand. Bedoeld om elk
uur te draaien (GitHub Actions, Taakplanner of cron).

Bron   https://opendata.ndw.nu/Truckparking_Parking_Status.xml
Formaat DATEX II v3, ParkingStatusPublication

Berekent de bezetting per deelgroep plaatsen, omdat de totaalvelden
van NDW onbruikbaar zijn (zie toelichting bij CAPACITEIT). De ruwe
NDW-totalen worden ter vergelijking ook gelogd.

Alleen standaard Python 3.9+ nodig, geen extra pakketten.

Gebruik
    python truckparking_logger.py                 # schrijft naar ./data
    python truckparking_logger.py --map D:/logs   # andere map
    python truckparking_logger.py --bestand test.xml   # testen met lokaal bestand
"""

import argparse
import csv
import sys
import time
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from pathlib import Path

# ---------------------------------------------------------------- instellingen

BRON_URL = "https://opendata.ndw.nu/Truckparking_Parking_Status.xml"

LOCATIES = {
    "NL-12_411": "Maasvlakte Plaza Noord (Hormuzstraat)",
    "NL-12_418": "Maasvlakte Plaza Zuid (Luzonstraat 6)",
}

LOGBESTAND = "truckparking_log.csv"
GROEPENBESTAND = "truckparking_groepen.csv"

# Capaciteit per deelgroep (groupIndex in de NDW-feed).
#
# Toelichting. In de NDW-feed is het veld "bezet" vrijwel altijd 0 of 1, ook
# als de parking vol staat. Het totaalpercentage per locatie is een ONGEWOGEN
# gemiddelde van de groepspercentages, inclusief lege groepen, en daarmee ook
# onbruikbaar. Wat wel consistent is, is per groep het aantal vrije plaatsen
# samen met het bezettingspercentage. Daaruit volgt per groep een vaste
# capaciteit. Bezet = capaciteit min vrij.
#
# LET OP. Deze waarden zijn afgeleid uit één NDW-momentopname (11-09-2026) en
# daarom voorlopig. Groep 2 bij Noord heeft een bandbreedte van circa 406-410.
# Na enkele weken loggen worden ze herijkt op basis van de kolom
# capaciteit_afgeleid in het groepenbestand. Betekenis van de groepnummers is
# niet gedocumenteerd, de omschrijvingen zijn een hypothese.
CAPACITEIT = {
    "NL-12_411": {            # Noord, Hormuzstraat
        "2": 408,             # hoofdterrein vrachtwagens (hypothese)
        "3": 8,               # vermoedelijk exceptioneel transport (OOG, 8 plaatsen)
        "4": 77,              # onbekend, mogelijk LZV
        "5": 0,               # levert geen plaatsen
        "1000": 88,           # vermoedelijk gevaarlijke stoffen (havenbedrijf noemt 87)
    },
    "NL-12_418": {            # Zuid, Luzonstraat 6
        "2": 164,             # vrachtwagens (NDW-tabel noemt 162)
        "4": 36,              # zwaar transport (NDW-tabel noemt 36)
        "1000": 0,
    },
}
GRENS_CAPACITEIT_AFWIJKING = 3   # afwijking afgeleide vs ingestelde capaciteit waarboven we markeren
FOUTBESTAND = "foutlog.txt"

POGINGEN = 3            # aantal downloadpogingen
WACHTTIJD_SEC = 20      # wachten tussen pogingen
TIMEOUT_SEC = 30

GRENS_OUDERDOM_MIN = 30     # status ouder dan dit aantal minuten wordt gemarkeerd

NS = {"p": "http://datex2.eu/schema/3/parking"}

KOLOMMEN = [
    "datum_lokaal", "tijd_lokaal", "uur_lokaal", "log_tijd_utc",
    "locatie_id", "locatie_naam",
    # berekend, dit zijn de bruikbare cijfers
    "bezet_berekend", "vrij_berekend", "capaciteit", "bezetting_berekend_pct",
    # ruwe NDW-totalen, alleen ter controle
    "ndw_vrij", "ndw_bezet", "ndw_bezetting_pct",
    "site_status", "status_tijd_lokaal", "ouderdom_status_min",
    "controle",
]

GROEP_KOLOMMEN = [
    "datum_lokaal", "tijd_lokaal", "uur_lokaal",
    "locatie_id", "groep_index",
    "ndw_vrij", "ndw_bezet", "ndw_bezetting_pct",
    "capaciteit_ingesteld", "capaciteit_afgeleid", "bezet_berekend",
]

# ---------------------------------------------------------------- tijdzone

try:
    from zoneinfo import ZoneInfo
    TZ_LOKAAL = ZoneInfo("Europe/Amsterdam")
except Exception:
    # Op sommige Windows-installaties ontbreekt de tijdzonedatabase.
    # Dan valt het script terug op de tijdzone van de computer zelf.
    TZ_LOKAAL = None


def naar_lokaal(dt_utc):
    if TZ_LOKAAL is not None:
        return dt_utc.astimezone(TZ_LOKAAL)
    return dt_utc.astimezone()


def afronden_uur(dt):
    """Rondt af op het dichtstbijzijnde hele uur, zodat een vertraagde run
    (bijv. 07:09) toch als 07:00 in de reeks komt."""
    basis = dt.replace(minute=0, second=0, microsecond=0)
    if dt.minute >= 30:
        basis += timedelta(hours=1)
    return basis


def lees_utc(tekst):
    """Leest NDW-tijden zoals 2026-09-11T10:27:56.986493453Z (nanoseconden)."""
    if not tekst:
        return None
    t = tekst.strip().replace("Z", "+00:00")
    if "." in t:
        hoofd, rest = t.split(".", 1)
        fractie, _, zone = rest.partition("+")
        t = f"{hoofd}.{fractie[:6]}+{zone}" if zone else f"{hoofd}.{fractie[:6]}"
    try:
        dt = datetime.fromisoformat(t)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)

# ---------------------------------------------------------------- hulpfuncties


def nl_getal(waarde, decimalen=1):
    """Getal met decimale komma, zodat Nederlandse Excel het goed leest."""
    if waarde is None:
        return ""
    if isinstance(waarde, int):
        return str(waarde)
    return f"{waarde:.{decimalen}f}".replace(".", ",")


def lees_int(element, pad):
    el = element.find(pad, NS)
    if el is None or el.text is None:
        return None
    try:
        return int(float(el.text))
    except ValueError:
        return None


def lees_float(element, pad):
    el = element.find(pad, NS)
    if el is None or el.text is None:
        return None
    try:
        return float(el.text)
    except ValueError:
        return None


def schrijf_fout(map_, bericht):
    nu = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    with open(map_ / FOUTBESTAND, "a", encoding="utf-8") as f:
        f.write(f"{nu}  {bericht}\n")


def voeg_toe(pad, kolommen, rijen):
    nieuw = not pad.exists()
    with open(pad, "a", newline="", encoding="utf-8-sig" if nieuw else "utf-8") as f:
        w = csv.DictWriter(f, fieldnames=kolommen, delimiter=";")
        if nieuw:
            w.writeheader()
        w.writerows(rijen)

# ---------------------------------------------------------------- ophalen


def haal_xml(map_):
    laatste_fout = None
    for poging in range(1, POGINGEN + 1):
        try:
            req = urllib.request.Request(
                BRON_URL, headers={"User-Agent": "truckparking-logger/1.0"}
            )
            with urllib.request.urlopen(req, timeout=TIMEOUT_SEC) as r:
                return r.read()
        except Exception as e:  # netwerk, timeout, http-fout
            laatste_fout = e
            schrijf_fout(map_, f"Poging {poging}/{POGINGEN} mislukt: {e}")
            if poging < POGINGEN:
                time.sleep(WACHTTIJD_SEC)
    raise RuntimeError(f"Ophalen mislukt na {POGINGEN} pogingen: {laatste_fout}")

# ---------------------------------------------------------------- verwerken


def verwerk(xml_bytes, log_utc):
    root = ET.fromstring(xml_bytes)
    log_lok = naar_lokaal(log_utc)
    uur = afronden_uur(log_lok)

    basis = {
        "datum_lokaal": log_lok.strftime("%Y-%m-%d"),
        "tijd_lokaal": log_lok.strftime("%H:%M"),
        "uur_lokaal": uur.strftime("%Y-%m-%d %H:%M"),
        "log_tijd_utc": log_utc.strftime("%Y-%m-%d %H:%M:%S"),
    }

    rijen, groeprijen, gevonden = [], [], set()

    for rec in root.findall("p:parkingRecordStatus", NS):
        ref = rec.find("p:parkingRecordReference", NS)
        loc_id = ref.get("id") if ref is not None else None
        if loc_id not in LOCATIES:
            continue
        gevonden.add(loc_id)

        occ = rec.find("p:parkingOccupancy", NS)
        ndw_vrij = lees_int(occ, "p:parkingNumberOfVacantSpaces") if occ is not None else None
        ndw_bezet = lees_int(occ, "p:parkingNumberOfOccupiedSpaces") if occ is not None else None
        ndw_pct = lees_float(occ, "p:parkingOccupancy") if occ is not None else None

        status_el = rec.find("p:parkingSiteStatus", NS)
        site_status = status_el.text if status_el is not None else ""
        status_utc = lees_utc(
            rec.findtext("p:parkingStatusOriginTime", default="", namespaces=NS)
        )
        ouderdom = (log_utc - status_utc).total_seconds() / 60 if status_utc else None

        controle = []
        caps = CAPACITEIT.get(loc_id, {})
        tot_cap, tot_bezet, tot_vrij, groepen_gezien = 0, 0, 0, set()

        for g in rec.findall("p:groupOfParkingSpacesStatus", NS):
            binnen = g.find("p:groupOfParkingSpacesStatus", NS)
            if binnen is None:
                continue
            gi = g.get("groupIndex", "")
            groepen_gezien.add(gi)
            g_vrij = lees_int(binnen, "p:parkingNumberOfVacantSpaces")
            g_bezet_ndw = lees_int(binnen, "p:parkingNumberOfOccupiedSpaces")
            g_pct = lees_float(binnen, "p:parkingOccupancy")

            # capaciteit afleiden uit vrij en percentage, alleen zinvol als niet (bijna) vol
            afgeleid = None
            if g_vrij is not None and g_pct is not None and 0 <= g_pct < 95 and g_vrij > 0:
                afgeleid = g_vrij / (1 - g_pct / 100)

            cap = caps.get(gi)
            g_bezet = None
            if cap is None:
                controle.append(f"onbekende groep {gi}")
            elif g_vrij is not None:
                if g_vrij > cap:
                    controle.append(f"groep {gi} vrij > capaciteit")
                g_bezet = max(cap - g_vrij, 0)
                tot_cap += cap
                tot_bezet += g_bezet
                tot_vrij += min(g_vrij, cap)
                if afgeleid is not None and abs(afgeleid - cap) > GRENS_CAPACITEIT_AFWIJKING:
                    controle.append(f"groep {gi} capaciteit wijkt af ({afgeleid:.0f} i.p.v. {cap})")

            groeprijen.append({
                "datum_lokaal": basis["datum_lokaal"],
                "tijd_lokaal": basis["tijd_lokaal"],
                "uur_lokaal": basis["uur_lokaal"],
                "locatie_id": loc_id,
                "groep_index": gi,
                "ndw_vrij": nl_getal(g_vrij),
                "ndw_bezet": nl_getal(g_bezet_ndw),
                "ndw_bezetting_pct": nl_getal(g_pct),
                "capaciteit_ingesteld": nl_getal(cap),
                "capaciteit_afgeleid": nl_getal(afgeleid),
                "bezet_berekend": nl_getal(g_bezet),
            })

        ontbrekend = [gi for gi, c in caps.items() if c > 0 and gi not in groepen_gezien]
        if ontbrekend:
            controle.append("groep ontbreekt " + ",".join(ontbrekend))
        if any(v is not None and v < 0 for v in (ndw_vrij, ndw_bezet)):
            controle.append("negatieve waarde in NDW-totaal")
        if ouderdom is not None and ouderdom > GRENS_OUDERDOM_MIN:
            controle.append(f"status ouder dan {GRENS_OUDERDOM_MIN} min")

        pct_ber = tot_bezet / tot_cap * 100 if tot_cap else None

        rijen.append({
            **basis,
            "locatie_id": loc_id,
            "locatie_naam": LOCATIES[loc_id],
            "bezet_berekend": nl_getal(tot_bezet) if tot_cap else "",
            "vrij_berekend": nl_getal(tot_vrij) if tot_cap else "",
            "capaciteit": nl_getal(tot_cap) if tot_cap else "",
            "bezetting_berekend_pct": nl_getal(pct_ber),
            "ndw_vrij": nl_getal(ndw_vrij),
            "ndw_bezet": nl_getal(ndw_bezet),
            "ndw_bezetting_pct": nl_getal(ndw_pct),
            "site_status": site_status,
            "status_tijd_lokaal": naar_lokaal(status_utc).strftime("%Y-%m-%d %H:%M:%S") if status_utc else "",
            "ouderdom_status_min": nl_getal(ouderdom),
            "controle": " | ".join(controle),
        })

    # locatie niet in de feed, toch een regel zodat het gat zichtbaar is
    for loc_id in LOCATIES:
        if loc_id not in gevonden:
            rijen.append({**basis, "locatie_id": loc_id, "locatie_naam": LOCATIES[loc_id],
                          **{k: "" for k in KOLOMMEN if k not in basis and k not in ("locatie_id", "locatie_naam")},
                          "controle": "locatie ontbreekt in NDW-feed"})
    return rijen, groeprijen


def lege_rijen(log_utc, reden):
    log_lok = naar_lokaal(log_utc)
    basis = {
        "datum_lokaal": log_lok.strftime("%Y-%m-%d"),
        "tijd_lokaal": log_lok.strftime("%H:%M"),
        "uur_lokaal": afronden_uur(log_lok).strftime("%Y-%m-%d %H:%M"),
        "log_tijd_utc": log_utc.strftime("%Y-%m-%d %H:%M:%S"),
    }
    rijen = []
    for loc_id, naam in LOCATIES.items():
        rij = {k: "" for k in KOLOMMEN}
        rij.update(basis, locatie_id=loc_id, locatie_naam=naam, controle=reden)
        rijen.append(rij)
    return rijen

# ---------------------------------------------------------------- hoofdprogramma


def main():
    ap = argparse.ArgumentParser(description="Log bezetting Maasvlakte Plaza Noord en Zuid")
    ap.add_argument("--map", default="data", help="map voor de CSV-bestanden (standaard ./data)")
    ap.add_argument("--bestand", help="lokaal XML-bestand gebruiken in plaats van downloaden (testen)")
    args = ap.parse_args()

    map_ = Path(args.map)
    map_.mkdir(parents=True, exist_ok=True)
    log_utc = datetime.now(timezone.utc)

    try:
        xml_bytes = Path(args.bestand).read_bytes() if args.bestand else haal_xml(map_)
        rijen, groeprijen = verwerk(xml_bytes, log_utc)
    except Exception as e:
        schrijf_fout(map_, f"Run mislukt: {e}")
        voeg_toe(map_ / LOGBESTAND, KOLOMMEN, lege_rijen(log_utc, f"ophalen of verwerken mislukt"))
        print(f"FOUT: {e}", file=sys.stderr)
        return 1

    voeg_toe(map_ / LOGBESTAND, KOLOMMEN, rijen)
    if groeprijen:
        voeg_toe(map_ / GROEPENBESTAND, GROEP_KOLOMMEN, groeprijen)

    for r in rijen:
        print(f"{r['tijd_lokaal']}  {r['locatie_naam']:<40} bezet {r['bezet_berekend']:>4} "
              f"van {r['capaciteit']:>4}  ({r['bezetting_berekend_pct']:>5}%)   "
              f"NDW zegt bezet {r['ndw_bezet']:>3}, {r['ndw_bezetting_pct']}%  {r['controle']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
