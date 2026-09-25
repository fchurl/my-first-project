"""
Extrait tous les vins de la selection "Foire aux vins" 2026 d'iDealwine,
avec le prix de livraison Royaume-Uni pour chaque vin, et exporte le
resultat dans une feuille Excel.

Source des donnees : le JSON __NEXT_DATA__ embarque dans le HTML de la
page de listing (index Algolia "prod_lots_price_desc", champ
"unitPriceByCountry.GB" fourni par iDealwine lui-meme pour chaque lot).
Aucune API privee n'est appelee directement : on recupere les memes
pages HTML publiques qu'un navigateur, page par page, via le pattern
de pagination /page-N-selection-FOIRE_AUX_VINS decouvert sur le site.

Notes sur les prix :
- "unitPriceByCountry.GB" est le prix par bouteille pour une livraison
  au Royaume-Uni, exprime en EUROS (pas en GBP), hors TVA francaise
  (le Royaume-Uni est en zone "EXPORT" chez iDealwine depuis le Brexit).
- Une estimation en GBP est ajoutee a partir du taux de change publie
  par iDealwine (endpoint public /api/v2/shop/exchange-rates). C'est
  une estimation, pas le prix reel qui serait facture au paiement.
- Les frais de livraison eux-memes ne sont pas inclus (ils dependent
  du volume total de la commande) et ne sont pas exposes par lot.
"""

import json
import re
import time
from dataclasses import dataclass, fields
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

BASE_URL = "https://www.idealwine.com"
LISTING_SLUG = "selection-FOIRE_AUX_VINS"
ORDER_BY = "price-desc"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    ),
    "Accept-Language": "fr-FR,fr;q=0.9",
}
NEXT_DATA_RE = re.compile(
    r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', re.S
)
OUTPUT_XLSX = Path(__file__).with_name("idealwine_foire_aux_vins_2026.xlsx")


def get_gbp_rate(session: requests.Session) -> float:
    """Taux de change EUR->GBP publie par iDealwine (estimation d'affichage)."""
    r = session.get(f"{BASE_URL}/api/v2/shop/exchange-rates", headers=HEADERS, timeout=20)
    r.raise_for_status()
    for entry in r.json():
        if entry["sourceCurrency"]["code"] == "EUR" and entry["targetCurrency"]["code"] == "GBP":
            return float(entry["ratio"])
    raise RuntimeError("Taux EUR->GBP introuvable dans /api/v2/shop/exchange-rates")


def fetch_page_hits(session: requests.Session, page_number: int) -> tuple[list[dict], int]:
    """Recupere une page de la liste (1-indexee) et renvoie (hits, nb_total_hits)."""
    if page_number == 1:
        path = f"/fr/acheter-du-vin/{LISTING_SLUG}"
    else:
        path = f"/fr/acheter-du-vin/page-{page_number}-{LISTING_SLUG}"
    url = f"{BASE_URL}{path}"
    r = session.get(url, params={"order-by": ORDER_BY}, headers=HEADERS, timeout=30)
    r.raise_for_status()
    m = NEXT_DATA_RE.search(r.text)
    if not m:
        raise RuntimeError(f"__NEXT_DATA__ introuvable sur {url}")
    data = json.loads(m.group(1))
    result = data["props"]["pageProps"]["serverState"]["initialResults"]["prod_lots"]["results"][0]
    return result["hits"], result["nbHits"]


@dataclass
class Wine:
    code: str
    nom: str
    domaine: str
    appellation: str
    region: str
    sous_region: str
    millesime: str
    couleur: str
    format_bouteille: str
    quantite_disponible: int
    prix_fr_ttc_eur: float
    prix_uk_ht_eur: float
    prix_uk_ht_gbp_estime: float
    remise_lot_pct: float
    seuil_remise_lot: str
    prix_uk_avec_remise_eur: float
    note_parker: str
    note_rvf: str
    note_wine_spectator: str
    note_figaro: str
    url_produit: str


COLUMN_LABELS = {
    "code": "Code produit",
    "nom": "Nom du vin",
    "domaine": "Domaine",
    "appellation": "Appellation",
    "region": "Region",
    "sous_region": "Sous-region",
    "millesime": "Millesime",
    "couleur": "Couleur",
    "format_bouteille": "Format",
    "quantite_disponible": "Qte disponible",
    "prix_fr_ttc_eur": "Prix France TTC (EUR)",
    "prix_uk_ht_eur": "Prix livraison UK / bouteille HT (EUR)",
    "prix_uk_ht_gbp_estime": "Prix livraison UK / bouteille HT estime (GBP)",
    "remise_lot_pct": "Remise lot (%)",
    "seuil_remise_lot": "Seuil remise lot",
    "prix_uk_avec_remise_eur": "Prix UK avec remise lot (EUR)",
    "note_parker": "Note Parker",
    "note_rvf": "Note RVF",
    "note_wine_spectator": "Note Wine Spectator",
    "note_figaro": "Note Figaro",
    "url_produit": "Lien iDealwine",
}

# Colonnes additionnelles a completer manuellement (comparaison Wine-Searcher)
MANUAL_COLUMNS = [
    "Meilleur prix UK - Wine-Searcher (GBP)",
    "Marchand (Wine-Searcher)",
    "Lien Wine-Searcher",
    "Ecart iDealwine vs Wine-Searcher (%)",
    "Date de releve",
]


def cents_to_units(value) -> float:
    return round(value / 100.0, 2) if value is not None else None


def parse_hit(hit: dict, gbp_rate: float) -> Wine:
    code = str(hit["code"])
    unit_prices = hit.get("unitPriceByCountry") or {}
    price_fr = cents_to_units(unit_prices.get("FR", hit.get("currentPrice")))
    price_uk_eur = cents_to_units(unit_prices.get("GB"))
    price_uk_gbp = round(price_uk_eur * gbp_rate, 2) if price_uk_eur is not None else None

    has_discount = bool(hit.get("hasLotDiscount"))
    discount_pct = hit.get("lotDiscountPercentage") if has_discount else None
    threshold = hit.get("unitsPerLotOfDiscount") if has_discount else None
    price_uk_with_discount = (
        round(price_uk_eur * (1 - discount_pct / 100.0), 2)
        if (has_discount and price_uk_eur is not None and discount_pct)
        else None
    )

    return Wine(
        code=code,
        nom=hit.get("name", ""),
        domaine=hit.get("domainName", ""),
        appellation=hit.get("subtitleName") or hit.get("appellation", ""),
        region=hit.get("region", ""),
        sous_region=hit.get("subregion", ""),
        millesime=hit.get("vintage") or "N.M.",
        couleur=(hit.get("translatedColors") or {}).get("fr", hit.get("color", "")),
        format_bouteille=hit.get("bottleSize", ""),
        quantite_disponible=hit.get("availableQuantity") or 0,
        prix_fr_ttc_eur=price_fr,
        prix_uk_ht_eur=price_uk_eur,
        prix_uk_ht_gbp_estime=price_uk_gbp,
        remise_lot_pct=discount_pct or "",
        seuil_remise_lot=(f"a partir de {threshold} bouteilles" if threshold else ""),
        prix_uk_avec_remise_eur=price_uk_with_discount or "",
        note_parker=hit.get("noteRobertParker") or "",
        note_rvf=hit.get("noteRvf") or "",
        note_wine_spectator=hit.get("noteWineSpectator") or "",
        note_figaro=hit.get("noteFigaro") or "",
        url_produit=f"{BASE_URL}/fr/acheter-vin/{code}",
    )


def scrape_all_wines() -> list[Wine]:
    session = requests.Session()
    gbp_rate = get_gbp_rate(session)
    print(f"Taux de change iDealwine EUR->GBP : {gbp_rate}")

    wines: list[Wine] = []
    seen_codes: set[str] = set()
    page = 1
    total_expected = None

    while True:
        hits, nb_hits = fetch_page_hits(session, page)
        total_expected = total_expected or nb_hits
        if not hits:
            break
        for hit in hits:
            wine = parse_hit(hit, gbp_rate)
            if wine.code in seen_codes:
                continue
            seen_codes.add(wine.code)
            wines.append(wine)
        print(f"Page {page}: {len(hits)} vins recuperes (total cumule {len(wines)}/{total_expected})")
        if len(wines) >= total_expected:
            break
        page += 1
        time.sleep(1)  # pas de martelage du serveur

    return wines


def write_excel(wines: list[Wine], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Foire aux vins 2026"

    field_names = [f.name for f in fields(Wine)]
    headers = [COLUMN_LABELS[name] for name in field_names] + MANUAL_COLUMNS
    ws.append(headers)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="7A1F2B", end_color="7A1F2B", fill_type="solid")
    for col_idx, _ in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"

    # tri par prix UK decroissant (coeur de la demande : comparaison UK)
    wines_sorted = sorted(
        wines, key=lambda w: (w.prix_uk_ht_eur or 0), reverse=True
    )

    eur_cols = {"prix_fr_ttc_eur", "prix_uk_ht_eur", "prix_uk_avec_remise_eur"}
    gbp_cols = {"prix_uk_ht_gbp_estime"}

    for wine in wines_sorted:
        row = []
        for name in field_names:
            value = getattr(wine, name)
            row.append(value)
        row += [""] * len(MANUAL_COLUMNS)
        ws.append(row)

    last_row = ws.max_row
    for col_idx, name in enumerate(field_names, start=1):
        if name in eur_cols:
            for r in range(2, last_row + 1):
                ws.cell(row=r, column=col_idx).number_format = '#,##0.00 "EUR"'
        elif name in gbp_cols:
            for r in range(2, last_row + 1):
                ws.cell(row=r, column=col_idx).number_format = '#,##0.00 "GBP"'
        if name == "url_produit":
            for r in range(2, last_row + 1):
                cell = ws.cell(row=r, column=col_idx)
                if cell.value:
                    cell.hyperlink = cell.value
                    cell.font = Font(color="0563C1", underline="single")

    widths = {
        "code": 12, "nom": 42, "domaine": 26, "appellation": 22, "region": 16,
        "sous_region": 16, "millesime": 10, "couleur": 10, "format_bouteille": 12,
        "quantite_disponible": 10, "prix_fr_ttc_eur": 16, "prix_uk_ht_eur": 20,
        "prix_uk_ht_gbp_estime": 20, "remise_lot_pct": 12, "seuil_remise_lot": 20,
        "prix_uk_avec_remise_eur": 18, "note_parker": 10, "note_rvf": 10,
        "note_wine_spectator": 10, "note_figaro": 10, "url_produit": 45,
    }
    for col_idx, name in enumerate(field_names, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = widths.get(name, 14)
    for offset, _ in enumerate(MANUAL_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(len(field_names) + offset)].width = 26

    ws.auto_filter.ref = ws.dimensions

    wb.save(path)


def main() -> None:
    wines = scrape_all_wines()
    print(f"Total vins extraits : {len(wines)}")
    write_excel(wines, OUTPUT_XLSX)
    print(f"Fichier Excel ecrit : {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
