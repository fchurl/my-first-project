# -*- coding: utf-8 -*-
"""
Corrige la comparaison iDealwine / Wine-Searcher : separe explicitement
les prix "hors droits & taxe UK" (In Bond) des prix "inc. 20% sales tax"
(droits d'accise UK + TVA 20% deja payes), et ne compare jamais l'un a
l'autre sans ajustement.

Base de comparaison :
- iDealwine "unitPriceByCountry.GB" = prix HT (pas de TVA francaise, le
  Royaume-Uni etant hors zone UE / EXPORT). C'est un prix HORS droits
  d'accise UK et HORS TVA UK -> comparable directement a un prix
  Wine-Searcher "In Bond".
- Pour comparer a un prix Wine-Searcher "inc. 20% sales tax" (retail,
  droits payes), on calcule une estimation du prix iDealwine "livre UK" :
      prix_HT_GBP  ->  + droits d'accise UK (fonction de l'ABV et du
      volume, taux en vigueur au 1er fevrier 2026 : £30.62 / litre
      d'alcool pur pour les vins 8.5%-22% ABV, memes taux vin tranquille
      et effervescent)  ->  x 1.20 (TVA UK)
  Cette estimation reste approximative (elle n'inclut pas la marge d'un
  distributeur UK ni les frais de transport UK locaux), ce qui est
  precise dans la colonne Remarque.

Les degres d'alcool et volumes exacts ont ete recuperes directement sur
les fiches produit iDealwine (champ alcoolProportion / customVolume).
"""

from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

XLSX = Path(__file__).with_name("idealwine_foire_aux_vins_2026.xlsx")
GBP_RATE = 0.8938  # EUR -> GBP, taux publie par iDealwine (07/09/2026)
UK_DUTY_PER_LPA = 30.62  # GBP / litre d'alcool pur, vin 8.5-22% ABV, depuis le 01/02/2026
UK_VAT = 0.20
TODAY = date.today().isoformat()

HEADER_FILL = PatternFill(start_color="7A1F2B", end_color="7A1F2B", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
LINK_FONT = Font(color="0563C1", underline="single")


def duty_per_bottle(abv_pct: float, volume_ml: int) -> float:
    return round(abv_pct / 100 * (volume_ml / 1000) * UK_DUTY_PER_LPA, 2)


# code produit iDealwine -> caracteristiques + releves Wine-Searcher
DATA = {
    "2690621": dict(  # Arnaud Lambert Saumur Breze Clos de Midi 2024
        abv=13.0, volume_ml=750,
        ib_price=None, ib_merchant=None,
        retail_price=17.40, retail_merchant="Vinatis UK",
        url="https://www.wine-searcher.com/find/arnaud+lambert+breeze+clos+de+midi+saumur+blanc+anjou+loire+france/2024/uk",
        remark="Correspondance exacte (millesime 2024). Aucune offre In Bond trouvee.",
    ),
    "2701936": dict(  # Thibault Liger-Belair Bourgogne Les Deux Terres 2022
        abv=13.0, volume_ml=750,
        ib_price=17.00, ib_merchant="Cru World Wine UK (millesime 2018, pas 2022)",
        retail_price=25.30, retail_merchant="Sip Wines Shop (millesime 2023, pas 2022)",
        url="https://www.wine-searcher.com/find/thibault+liger+bel+air+les+deux+terre+bourgogne+rouge+burgundy+france/1/uk",
        remark="Millesime 2022 introuvable sur Wine-Searcher UK (ni In Bond ni retail) : les deux prix ci-contre sont d'autres millesimes, comparaison indicative seulement.",
    ),
    "2673620": dict(  # Domaine Goisot Bourgogne Cotes d'Auxerre Blanc 2023
        abv=14.0, volume_ml=750,
        ib_price=None, ib_merchant=None,
        retail_price=13.40, retail_merchant="Des Journees Entieres dans les Vignes (France, PAS UK)",
        url="https://www.wine-searcher.com/find/guilhem+jean+hugues+goisot+blanc+bourgogne+cote+d+auxerre+burgundy+france/1/uk",
        remark="Aucun vendeur UK (In Bond ou retail) sur Wine-Searcher. Le prix indique est celui d'un caviste FRANCAIS (TVA francaise 20% incluse, pas TVA UK) -> non comparable a un prix livre UK.",
    ),
    "2721080": dict(  # Clos de la Roilette Fleurie 2024
        abv=13.0, volume_ml=750,
        ib_price=12.02, ib_merchant="Goedhuis Waddesdon (In Bond)",
        retail_price=17.94, retail_merchant="Vinatis UK",
        url="https://www.wine-searcher.com/find/dom+coudert+clos+de+la+roilette+fleurie+beaujolais+france/1/uk",
        remark="Correspondance exacte (millesime 2024) sur les deux bases.",
    ),
    "2451560": dict(  # Georges Descombes Morgon 2023
        abv=14.0, volume_ml=750,
        ib_price=None, ib_merchant=None,
        retail_price=25.50, retail_merchant="Parched Wine",
        url="https://www.wine-searcher.com/find/george+descombes+morgon+beaujolais+france/1/uk",
        remark="Correspondance exacte (millesime 2023). Aucune offre In Bond trouvee.",
    ),
    "2595743": dict(  # Francois Chidaine Les Bournais 2024
        abv=13.0, volume_ml=750,
        ib_price=None, ib_merchant=None,
        retail_price=35.00, retail_merchant="Millesima UK (prix/bouteille, caisse de 6)",
        url="https://www.wine-searcher.com/find/francois+chidaine+les+bournais+mont+louis+sur+loire+touraine+france/1/uk",
        remark="Correspondance exacte (millesime 2024). Aucune offre In Bond trouvee.",
    ),
    "2592975": dict(  # Jean-Paul & Benoit Droin Chablis 2024
        abv=13.0, volume_ml=750,
        ib_price=18.33, ib_merchant="Avint-Garde Fine Wine (In Bond)",
        retail_price=34.96, retail_merchant="Uncorked Ltd",
        url="https://www.wine-searcher.com/find/jean+paul+benoit+droin+chablis+burgundy+france/1/uk",
        remark="Correspondance exacte (millesime 2024) sur les deux bases.",
    ),
    "B2110049-101488": dict(  # Domaine de Villaine Bouzeron 2022
        abv=13.0, volume_ml=750,
        ib_price=24.00, ib_merchant="Justerini & Brooks (In Bond)",
        retail_price=29.83, retail_merchant="Millesima UK (caisse de 6)",
        url="https://www.wine-searcher.com/find/dom+a+p+de+villaine+bouzeron+cote+chalonnaise+burgundy+france/1/uk",
        remark="Correspondance exacte (millesime 2022) sur les deux bases.",
    ),
    "2502373": dict(  # Sylvain Pataille Bourgogne 2023
        abv=13.0, volume_ml=750,
        ib_price=20.00, ib_merchant="BBX Berry Bros. & Rudd (In Bond)",
        retail_price=25.68, retail_merchant="Vinatis UK",
        url="https://www.wine-searcher.com/find/dom+sylvain+pataille+bourgogne+rouge+burgundy+france/1/uk",
        remark="Correspondance exacte (millesime 2023) sur les deux bases.",
    ),
    "B2027279-3475": dict(  # La Taille aux Loups Triple Zero (NV)
        abv=13.0, volume_ml=750,
        ib_price=17.00, ib_merchant="Bordeaux Index (In Bond)",
        retail_price=29.90, retail_merchant="Cambridge Wine Merchants",
        url="https://www.wine-searcher.com/find/jacky+blot+dom+de+la+taille+aux+loup+mont+louis+triple+zero+petillant+mousseux+sur+loire+touraine+france/1/uk",
        remark="Vin non millesime (NV) : comparaison directe valide sur les deux bases.",
    ),
    "2607633": dict(  # La Taille aux Loups Clos de Mosny 2024
        abv=12.0, volume_ml=750,
        ib_price=22.00, ib_merchant="Lay & Wheeler (In Bond, millesime 2023, pas 2024)",
        retail_price=33.80, retail_merchant="Millesima UK",
        url="https://www.wine-searcher.com/find/jacky+blot+dom+de+la+taille+aux+loup+clos+mosny+mont+louis+sur+loire+touraine+france/1/uk",
        remark="Retail = millesime 2024 exact. Aucune offre In Bond pour 2024 (celle indiquee est 2023).",
    ),
    "2689778": dict(  # Tenuta delle Terre Nere Etna Rosso 2024
        abv=13.0, volume_ml=750,
        ib_price=16.75, ib_merchant="Justerini & Brooks (In Bond)",
        retail_price=22.50, retail_merchant="VINVM LTD",
        url="https://www.wine-searcher.com/find/tenute+terre+nere+rosso+etna+sicilia+italy/1/uk",
        remark="Correspondance exacte (millesime 2024) sur les deux bases.",
    ),
    "2630946": dict(  # R. Lopez de Heredia Vina Tondonia Reserva 2013
        abv=13.0, volume_ml=750,
        ib_price=26.67, ib_merchant="Cuchet & Co Wines (In Bond)",
        retail_price=None, retail_merchant=None,
        url="https://www.wine-searcher.com/find/r+lopez+de+heredia+vina+tondonia+rsrv+doca+rioja+alta+spain/1/uk",
        remark="Correspondance exacte (millesime 2013) pour l'In Bond. Aucun prix retail (inc. taxe) UK capture pour ce millesime (275 offres au total, catalogue trop long pour etre parcouru en integralite sans abonnement PRO).",
    ),
    "2436392": dict(  # Crystallum Peter Max Pinot Noir 2023
        abv=14.0, volume_ml=750,
        ib_price=24.50, ib_merchant="Cru World Wine UK (In Bond)",
        retail_price=None, retail_merchant=None,
        url="https://www.wine-searcher.com/find/crystallum+peter+max+pinot+noir+western+cape+south+africa/1/uk",
        remark="Correspondance exacte (millesime 2023) pour l'In Bond. Aucun prix retail (inc. taxe) capture pour ce millesime (seuls 2024/2025 ont des offres retail).",
    ),
    "2726548": dict(  # Jamet Cote-Rotie Fructus Voluptas 2024
        abv=13.0, volume_ml=750,
        ib_price=49.17, ib_merchant="Vinum Fine Wines London (In Bond, millesime 2022, pas 2024)",
        retail_price=None, retail_merchant=None,
        url="https://www.wine-searcher.com/find/dom+jamet+fructus+voluptas+cote+rotie+rhone+france/1/uk",
        remark="Millesime 2024 introuvable (In Bond ou retail) ; comparaison sur 2022 In Bond uniquement, a relativiser fortement (millesime different + base HT).",
    ),
    "B2110074-108981": dict(  # De Sousa Grand Cru Extra-Brut Reserve (NV)
        abv=13.0, volume_ml=750,
        ib_price=35.00, ib_merchant="Atlas Fine Wines (In Bond)",
        retail_price=45.33, retail_merchant="Millesima UK (caisse de 6)",
        url="https://www.wine-searcher.com/find/de+sousa+rsrv+grand+cru+blancdeblanc+extra+brut+champagne+france/1/uk",
        remark="Vin non millesime (NV) : comparaison directe valide sur les deux bases.",
    ),
    "B2110074-56614": dict(  # R. Pouillon Grande Vallee Extra-Brut (NM)
        abv=12.0, volume_ml=750,
        ib_price=None, ib_merchant=None,
        retail_price=None, retail_merchant=None,
        url="https://www.wine-searcher.com/find/r+pouillon+grand+valley+extra+brut+champagne+france/1/uk",
        remark="Aucun vendeur UK independant (In Bond ou retail) : seules les offres iDealwine lui-meme apparaissent sur Wine-Searcher. Comparaison impossible.",
    ),
    "2634748": dict(  # Arretxea Irouleguy Hegoxuri 2023
        abv=14.0, volume_ml=750,
        ib_price=31.56, ib_merchant="offre 'By Request' In Bond (millesime 2023)",
        retail_price=44.90, retail_merchant="Sip Wines Shop",
        url="https://www.wine-searcher.com/find/dom+arretxea+cuvee+hegoxuri+irouleguy+south+west+france/1/uk",
        remark="Correspondance exacte (millesime 2023) sur les deux bases.",
    ),
    "2724860": dict(  # Arretxea Irouleguy Tradition 2024
        abv=13.0, volume_ml=750,
        ib_price=None, ib_merchant=None,
        retail_price=None, retail_merchant=None,
        url="https://www.wine-searcher.com/find/dom+arretxea+irouleguy+south+west+france/1/uk",
        remark="Aucun vendeur UK independant, millesime 2024 absent : seule l'offre iDealwine lui-meme (2022, magnum) apparait. Comparaison impossible.",
    ),
    "2482590": dict(  # Fumey-Chatelain Arbois Chardonnay 2023
        abv=13.0, volume_ml=750,
        ib_price=None, ib_merchant=None,
        retail_price=34.75, retail_merchant="Parched Wine",
        url="https://www.wine-searcher.com/find/fumet+chatelain+chard+arbois+jura+france/1/uk",
        remark="Correspondance exacte (millesime 2023). Aucune offre In Bond trouvee.",
    ),
    "2706622": dict(  # Fumey-Chatelain Vin Jaune Arbois 2018
        abv=13.0, volume_ml=750,
        ib_price=None, ib_merchant=None,
        retail_price=None, retail_merchant=None,
        url="https://www.wine-searcher.com/find/fumet+chatelain+jaune+arbois+jura+france/1/uk",
        remark="Millesime 2018 : seule l'offre iDealwine lui-meme apparait sur Wine-Searcher UK (£133.02). Le Vin Jaune AOC est traditionnellement mis en bouteille de 62cl (clavelin) et non 75cl : le champ volume d'iDealwine ne precise pas ce point, donc les prix par bouteille ci-contre ne sont pas necessairement normalises au meme volume -- a verifier manuellement avant toute decision d'achat.",
    ),
    "2607790": dict(  # Domaine de Belliviere Les Premices 2023
        abv=13.0, volume_ml=750,
        ib_price=None, ib_merchant=None,
        retail_price=29.85, retail_merchant="Sip Wines Shop",
        url="https://www.wine-searcher.com/find/eric+nicolas+dom+de+belliviere+premices+coteaux+du+loire+vendomois+france/1/uk",
        remark="Correspondance exacte (millesime 2023, etiquete Jasnieres sur Wine-Searcher). Aucune offre In Bond trouvee.",
    ),
    "2473457": dict(  # Marcel Deiss Complantation 2023
        abv=14.0, volume_ml=750,
        ib_price=None, ib_merchant=None,
        retail_price=21.36, retail_merchant="Vinatis UK (millesime 2024, pas 2023)",
        url="https://www.wine-searcher.com/find/dom+marcel+deiss+blanc+complantation+alsace+france/1/uk",
        remark="Millesime 2023 absent sur Wine-Searcher UK (seul 2024 en stock) ; comparaison a relativiser. Aucune offre In Bond trouvee.",
    ),
    "2632798": dict(  # Markus Molitor Riesling Haus Klosterberg White Capsule 2024
        abv=12.0, volume_ml=750,
        ib_price=None, ib_merchant=None,
        retail_price=17.54, retail_merchant="Vinatis UK",
        url="https://www.wine-searcher.com/find/markus+molitor+haus+kloster+berg+white+capsule+riesling+trocken+mosel+germany/1/uk",
        remark="Correspondance exacte (millesime 2024). Aucune offre In Bond trouvee.",
    ),
    "B2111346-16915": dict(  # Tokaji Aszu 5 Puttonyos Disznoko 2014
        abv=12.0, volume_ml=500,
        ib_price=None, ib_merchant=None,
        retail_price=37.48, retail_merchant="Vinatis UK",
        url="https://www.wine-searcher.com/find/disznoko+aszu+five+puttonyos+tokay+hungary/1/uk",
        remark="Correspondance exacte (millesime 2014, flacon 50cl des deux cotes). Aucune offre In Bond trouvee.",
    ),
}

NEW_COLUMNS = [
    "WS - Prix In Bond HT (GBP)",
    "WS - Vendeur (In Bond)",
    "Ecart HT vs In Bond (%)",
    "WS - Prix retail inc. taxe UK (GBP)",
    "WS - Vendeur (retail)",
    "iDealwine estime livre UK (GBP) [HT+droits+TVA20%]",
    "Ecart livre vs retail inc. taxe (%)",
    "Lien Wine-Searcher",
    "Date de releve",
    "Remarque Wine-Searcher",
]

OBSOLETE_COLUMNS = [
    "Meilleur prix UK - Wine-Searcher (GBP)",
    "Marchand (Wine-Searcher)",
    "Lien Wine-Searcher",
    "Ecart iDealwine vs Wine-Searcher (%)",
    "Date de releve",
    "Remarque Wine-Searcher",
]


def main():
    wb = load_workbook(XLSX)
    ws = wb.active

    header = [c.value for c in ws[1]]
    col = {name: idx + 1 for idx, name in enumerate(header)}
    price_uk_col = col["Prix livraison UK / bouteille HT (EUR)"]

    # supprime les anciennes colonnes obsoletes (droite -> gauche pour ne pas decaler les index restants)
    obsolete_idx = sorted({col[name] for name in OBSOLETE_COLUMNS if name in col}, reverse=True)
    for idx in obsolete_idx:
        ws.delete_cols(idx)

    # ajoute les nouvelles colonnes a la fin
    start_col = ws.max_column + 1
    for offset, name in enumerate(NEW_COLUMNS):
        c = start_col + offset
        cell = ws.cell(row=1, column=c, value=name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        width = 60 if "Remarque" in name or "Lien" in name else 26
        ws.column_dimensions[get_column_letter(c)].width = width

    new_col = {name: start_col + i for i, name in enumerate(NEW_COLUMNS)}

    # recharge le mapping complet des colonnes (les anciennes ont ete supprimees)
    header = [c.value for c in ws[1]]
    col = {name: idx + 1 for idx, name in enumerate(header)}
    price_uk_col = col["Prix livraison UK / bouteille HT (EUR)"]
    code_col = col["Code produit"]

    matched = 0
    for row in range(2, ws.max_row + 1):
        code = str(ws.cell(row=row, column=code_col).value)
        if code not in DATA:
            continue
        info = DATA[code]
        matched += 1

        idealwine_eur = ws.cell(row=row, column=price_uk_col).value
        idealwine_gbp = round(idealwine_eur * GBP_RATE, 2) if idealwine_eur else None
        duty = duty_per_bottle(info["abv"], info["volume_ml"])
        landed_estimate = round((idealwine_gbp + duty) * (1 + UK_VAT), 2) if idealwine_gbp else None

        if info["ib_price"] is not None:
            ws.cell(row=row, column=new_col["WS - Prix In Bond HT (GBP)"], value=info["ib_price"]).number_format = '#,##0.00 "GBP"'
            ws.cell(row=row, column=new_col["WS - Vendeur (In Bond)"], value=info["ib_merchant"])
            if idealwine_gbp:
                ecart_ht = round((info["ib_price"] - idealwine_gbp) / idealwine_gbp, 4)
                c = ws.cell(row=row, column=new_col["Ecart HT vs In Bond (%)"], value=ecart_ht)
                c.number_format = '+0.0%;-0.0%'

        if info["retail_price"] is not None:
            ws.cell(row=row, column=new_col["WS - Prix retail inc. taxe UK (GBP)"], value=info["retail_price"]).number_format = '#,##0.00 "GBP"'
            ws.cell(row=row, column=new_col["WS - Vendeur (retail)"], value=info["retail_merchant"])
            if landed_estimate:
                ws.cell(row=row, column=new_col["iDealwine estime livre UK (GBP) [HT+droits+TVA20%]"], value=landed_estimate).number_format = '#,##0.00 "GBP"'
                ecart_ret = round((info["retail_price"] - landed_estimate) / landed_estimate, 4)
                c = ws.cell(row=row, column=new_col["Ecart livre vs retail inc. taxe (%)"], value=ecart_ret)
                c.number_format = '+0.0%;-0.0%'

        link_cell = ws.cell(row=row, column=new_col["Lien Wine-Searcher"], value=info["url"])
        link_cell.hyperlink = info["url"]
        link_cell.font = LINK_FONT
        ws.cell(row=row, column=new_col["Date de releve"], value=TODAY)
        ws.cell(row=row, column=new_col["Remarque Wine-Searcher"], value=info["remark"])

    ws.auto_filter.ref = ws.dimensions
    wb.save(XLSX)
    print(f"{matched}/25 vins mis a jour dans {XLSX.name}")


if __name__ == "__main__":
    main()
