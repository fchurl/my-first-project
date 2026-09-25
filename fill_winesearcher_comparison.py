# -*- coding: utf-8 -*-
"""
Complete la feuille idealwine_foire_aux_vins_2026.xlsx avec les relevés
manuels effectués sur wine-searcher.com (navigateur, session assistée)
pour les 25 vins selectionnes par l'utilisateur.

Colonnes completees (deja presentes, vides, en fin de tableau) :
  - Meilleur prix UK - Wine-Searcher (GBP)
  - Marchand (Wine-Searcher)
  - Lien Wine-Searcher
  - Ecart iDealwine vs Wine-Searcher (%)
  - Date de releve
Une colonne "Remarque Wine-Searcher" est ajoutee pour les mises en garde
(millesime absent, prix "In Bond" hors droits/TVA UK, pas de vendeur UK
independant, format de bouteille different, etc.).
"""

from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

XLSX = Path(__file__).with_name("idealwine_foire_aux_vins_2026.xlsx")
GBP_RATE = 0.89393  # EUR -> GBP, taux publie par iDealwine (cf script d'extraction)
TODAY = date.today().isoformat()

# code produit iDealwine -> donnees releves Wine-Searcher
# ws_price_gbp : le meilleur prix retail UK "pret a expedier" quand disponible,
#                sinon le meilleur prix trouve (avec remarque explicite)
DATA = {
    "2690621": dict(  # Arnaud Lambert Saumur Breze Clos de Midi 2024
        ws_price=17.40, merchant="Vinatis UK",
        url="https://www.wine-searcher.com/find/arnaud+lambert+breeze+clos+de+midi+saumur+blanc+anjou+loire+france/2024/uk",
        remark="Correspondance exacte (millesime 2024).",
    ),
    "2701936": dict(  # Thibault Liger-Belair Bourgogne Les Deux Terres 2022
        ws_price=26.90, merchant="Millesima UK (2021, la plus proche dispo)",
        url="https://www.wine-searcher.com/find/thibault+liger+bel+air+les+deux+terre+bourgogne+rouge+burgundy+france/1/uk",
        remark="Millesime 2022 absent sur Wine-Searcher UK ; comparaison sur 2021 (Millesima, £26.90) / 2023 (£25.30).",
    ),
    "2673620": dict(  # Domaine Goisot Bourgogne Cotes d'Auxerre Blanc 2023
        ws_price=13.40, merchant="Des Journees Entieres dans les Vignes (France, pas UK)",
        url="https://www.wine-searcher.com/find/guilhem+jean+hugues+goisot+blanc+bourgogne+cote+d+auxerre+burgundy+france/1/uk",
        remark="Aucun vendeur UK ; recherche elargie a l'Europe. Prix France hors UK indicatif seulement.",
    ),
    "2721080": dict(  # Clos de la Roilette Fleurie 2024
        ws_price=17.94, merchant="Vinatis UK",
        url="https://www.wine-searcher.com/find/dom+coudert+clos+de+la+roilette+fleurie+beaujolais+france/1/uk",
        remark="Correspondance exacte (millesime 2024). Option In Bond Goedhuis £12.02 HT droits/taxes UK.",
    ),
    "2451560": dict(  # Georges Descombes Morgon 2023
        ws_price=25.50, merchant="Parched Wine",
        url="https://www.wine-searcher.com/find/george+descombes+morgon+beaujolais+france/1/uk",
        remark="Correspondance exacte (millesime 2023).",
    ),
    "2595743": dict(  # Francois Chidaine Les Bournais 2024
        ws_price=35.00, merchant="Millesima UK (prix/bouteille, caisse de 6)",
        url="https://www.wine-searcher.com/find/francois+chidaine+les+bournais+mont+louis+sur+loire+touraine+france/1/uk",
        remark="Correspondance exacte (millesime 2024).",
    ),
    "2592975": dict(  # Jean-Paul & Benoit Droin Chablis 2024
        ws_price=34.96, merchant="Hedonism Wines",
        url="https://www.wine-searcher.com/find/jean+paul+benoit+droin+chablis+burgundy+france/1/uk",
        remark="Correspondance exacte (millesime 2024).",
    ),
    "B2110049-101488": dict(  # Domaine de Villaine Bouzeron 2022
        ws_price=29.83, merchant="Millesima UK",
        url="https://www.wine-searcher.com/find/dom+a+p+de+villaine+bouzeron+cote+chalonnaise+burgundy+france/1/uk",
        remark="Correspondance exacte (millesime 2022). Option In Bond Justerini&Brooks £24.00 HT droits/taxes UK.",
    ),
    "2502373": dict(  # Sylvain Pataille Bourgogne 2023
        ws_price=25.68, merchant="Vinatis UK",
        url="https://www.wine-searcher.com/find/dom+sylvain+pataille+bourgogne+rouge+burgundy+france/1/uk",
        remark="Correspondance exacte (millesime 2023).",
    ),
    "B2027279-3475": dict(  # La Taille aux Loups Triple Zero (NM)
        ws_price=29.90, merchant="Cambridge Wine Merchants",
        url="https://www.wine-searcher.com/find/jacky+blot+dom+de+la+taille+aux+loup+mont+louis+triple+zero+petillant+mousseux+sur+loire+touraine+france/1/uk",
        remark="Vin non millesime (NV), comparaison directe valide. Option In Bond des £17-25 HT droits/taxes UK.",
    ),
    "2607633": dict(  # La Taille aux Loups Clos de Mosny 2024
        ws_price=33.80, merchant="Millesima UK",
        url="https://www.wine-searcher.com/find/jacky+blot+dom+de+la+taille+aux+loup+clos+mosny+mont+louis+sur+loire+touraine+france/1/uk",
        remark="Correspondance exacte (millesime 2024).",
    ),
    "2689778": dict(  # Tenuta delle Terre Nere Etna Rosso 2024
        ws_price=22.50, merchant="VINVM LTD",
        url="https://www.wine-searcher.com/find/tenute+terre+nere+rosso+etna+sicilia+italy/1/uk",
        remark="Correspondance exacte (millesime 2024).",
    ),
    "2630946": dict(  # R. Lopez de Heredia Vina Tondonia Reserva 2013
        ws_price=27.00, merchant="Cuchet & Co Wines (In Bond)",
        url="https://www.wine-searcher.com/find/r+lopez+de+heredia+vina+tondonia+rsrv+doca+rioja+alta+spain/1/uk",
        remark="Prix In Bond, HORS droits de douane et TVA UK (~+30-40% une fois dedouane) : le cout reel livre est probablement proche voire superieur au prix iDealwine.",
    ),
    "2436392": dict(  # Crystallum Peter Max Pinot Noir 2023
        ws_price=24.50, merchant="Cru World Wine UK (In Bond)",
        url="https://www.wine-searcher.com/find/crystallum+peter+max+pinot+noir+western+cape+south+africa/1/uk",
        remark="Seule option 2023 = In Bond, HORS droits de douane et TVA UK (~+30-40%) : cout reel livre probablement superieur au prix affiche.",
    ),
    "2726548": dict(  # Jamet Cote-Rotie Fructus Voluptas 2024
        ws_price=49.17, merchant="Ideal Wine Company (2022 In Bond, la plus proche dispo)",
        url="https://www.wine-searcher.com/find/dom+jamet+fructus+voluptas+cote+rotie+rhone+france/1/uk",
        remark="Millesime 2024 absent sur Wine-Searcher UK ; comparaison sur 2022 In Bond (HT droits/taxes UK), ecart donc a relativiser.",
    ),
    "B2110074-108981": dict(  # De Sousa Grand Cru Extra-Brut Reserve (NM)
        ws_price=45.33, merchant="Millesima UK",
        url="https://www.wine-searcher.com/find/de+sousa+rsrv+grand+cru+blancdeblanc+extra+brut+champagne+france/1/uk",
        remark="Vin non millesime (NV). Option In Bond Atlas Fine Wines £35.00 HT droits/taxes UK.",
    ),
    "B2110074-56614": dict(  # R. Pouillon Grande Vallee Extra-Brut (NM)
        ws_price=None, merchant="Aucun vendeur UK independant",
        url="https://www.wine-searcher.com/find/r+pouillon+grand+valley+extra+brut+champagne+france/1/uk",
        remark="Seules offres UK listees = iDealwine lui-meme (£40.40/bouteille). Pas de comparaison independante possible.",
    ),
    "2634748": dict(  # Arretxea Irouleguy Hegoxuri 2023
        ws_price=45.00, merchant="Hedonism Wines",
        url="https://www.wine-searcher.com/find/dom+arretxea+cuvee+hegoxuri+irouleguy+south+west+france/1/uk",
        remark="Correspondance exacte (millesime 2023).",
    ),
    "2724860": dict(  # Arretxea Irouleguy Tradition 2024
        ws_price=None, merchant="Aucun vendeur UK independant",
        url="https://www.wine-searcher.com/find/dom+arretxea+irouleguy+south+west+france/1/uk",
        remark="Millesime 2024 absent ; seule offre UK = iDealwine lui-meme (2022, magnum). Pas de comparaison independante possible.",
    ),
    "2482590": dict(  # Fumey-Chatelain Arbois Chardonnay 2023
        ws_price=34.75, merchant="Parched Wine",
        url="https://www.wine-searcher.com/find/fumet+chatelain+chard+arbois+jura+france/1/uk",
        remark="Correspondance exacte (millesime 2023).",
    ),
    "2706622": dict(  # Fumey-Chatelain Vin Jaune Arbois 2018
        ws_price=None, merchant="Aucun vendeur UK independant pour 2018",
        url="https://www.wine-searcher.com/find/fumet+chatelain+jaune+arbois+jura+france/1/uk",
        remark="2018 = uniquement iDealwine lui-meme (£133.02, format non precise). Millesime 2017 chez Sager and Wine : £81/62cl soit ~£98/750ml-equivalent.",
    ),
    "2607790": dict(  # Domaine de Belliviere Les Premices 2023
        ws_price=29.85, merchant="Sip Wines Shop",
        url="https://www.wine-searcher.com/find/eric+nicolas+dom+de+belliviere+premices+coteaux+du+loire+vendomois+france/1/uk",
        remark="Correspondance exacte (millesime 2023) ; etiquete Jasnieres (AOC voisine) sur Wine-Searcher, meme cuvee 'Premices'.",
    ),
    "2473457": dict(  # Marcel Deiss Complantation 2023
        ws_price=21.36, merchant="Vinatis UK (2024, la plus proche dispo)",
        url="https://www.wine-searcher.com/find/dom+marcel+deiss+blanc+complantation+alsace+france/1/uk",
        remark="Millesime 2023 absent sur Wine-Searcher UK (seul 2024 en stock) ; ecart a relativiser.",
    ),
    "2632798": dict(  # Markus Molitor Riesling Haus Klosterberg White Capsule 2024
        ws_price=17.54, merchant="Vinatis UK",
        url="https://www.wine-searcher.com/find/markus+molitor+haus+kloster+berg+white+capsule+riesling+trocken+mosel+germany/1/uk",
        remark="Correspondance exacte (millesime 2024).",
    ),
    "B2111346-16915": dict(  # Tokaji Aszu 5 Puttonyos Disznoko 2014
        ws_price=37.48, merchant="Vinatis UK",
        url="https://www.wine-searcher.com/find/disznoko+aszu+five+puttonyos+tokay+hungary/1/uk",
        remark="Correspondance exacte (millesime 2014). Prix WS pour flacon 50cl (format traditionnel Tokaji) ; a verifier que le format iDealwine est identique.",
    ),
}


def main():
    wb = load_workbook(XLSX)
    ws = wb.active

    header = [c.value for c in ws[1]]
    col = {name: idx + 1 for idx, name in enumerate(header)}

    # ajoute une colonne Remarque si absente
    if "Remarque Wine-Searcher" not in col:
        new_col_idx = ws.max_column + 1
        ws.cell(row=1, column=new_col_idx, value="Remarque Wine-Searcher")
        ws.cell(row=1, column=new_col_idx).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=1, column=new_col_idx).fill = PatternFill(
            start_color="7A1F2B", end_color="7A1F2B", fill_type="solid"
        )
        ws.column_dimensions[ws.cell(row=1, column=new_col_idx).column_letter].width = 60
        col["Remarque Wine-Searcher"] = new_col_idx

    code_col = col["Code produit"]
    price_uk_col = col["Prix livraison UK / bouteille HT (EUR)"]
    ws_price_col = col["Meilleur prix UK - Wine-Searcher (GBP)"]
    merchant_col = col["Marchand (Wine-Searcher)"]
    link_col = col["Lien Wine-Searcher"]
    ecart_col = col["Ecart iDealwine vs Wine-Searcher (%)"]
    date_col = col["Date de releve"]
    remark_col = col["Remarque Wine-Searcher"]

    matched = 0
    for row in range(2, ws.max_row + 1):
        code = str(ws.cell(row=row, column=code_col).value)
        if code not in DATA:
            continue
        info = DATA[code]
        matched += 1

        idealwine_eur = ws.cell(row=row, column=price_uk_col).value
        idealwine_gbp = round(idealwine_eur * GBP_RATE, 2) if idealwine_eur else None

        ws.cell(row=row, column=merchant_col, value=info["merchant"])
        ws.cell(row=row, column=link_col, value=info["url"])
        ws.cell(row=row, column=link_col).hyperlink = info["url"]
        ws.cell(row=row, column=link_col).font = Font(color="0563C1", underline="single")
        ws.cell(row=row, column=date_col, value=TODAY)
        ws.cell(row=row, column=remark_col, value=info["remark"])

        if info["ws_price"] is not None:
            ws.cell(row=row, column=ws_price_col, value=info["ws_price"])
            ws.cell(row=row, column=ws_price_col).number_format = '#,##0.00 "GBP"'
            if idealwine_gbp:
                ecart = round((info["ws_price"] - idealwine_gbp) / idealwine_gbp * 100, 1)
                ws.cell(row=row, column=ecart_col, value=ecart)
                ws.cell(row=row, column=ecart_col).number_format = '+0.0%;-0.0%'
                ws.cell(row=row, column=ecart_col, value=ecart / 100)

    wb.save(XLSX)
    print(f"{matched}/25 vins mis a jour dans {XLSX.name}")


if __name__ == "__main__":
    main()
