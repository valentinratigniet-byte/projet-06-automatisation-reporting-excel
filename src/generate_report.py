"""
Génère un rapport de ventes Excel mis en forme, automatiquement, depuis la base
PostgreSQL (projet 07). Remplace un reporting mensuel fait à la main.

Sortie : output/rapport_ventes_<AAAA-MM-JJ>.xlsx (synthèse + détails + graphiques).
Usage : python src/generate_report.py
"""
import os
from datetime import date
from pathlib import Path

import psycopg2
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DSN = os.environ.get("DATABASE_URL",
                     "postgresql://portfolio:portfolio@127.0.0.1:5433/ecommerce")
OUT = Path(__file__).resolve().parent.parent / "output"

# --- Charte Petrol & Ambre ---
PETROL = "137A8B"; PETROL_D = "0C5563"; AMBRE = "E4A93C"; LIGHT = "EAF1F2"
EURO = '#,##0" €"'
thin = Side(style="thin", color="DCE2E8")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD_FILL = PatternFill("solid", fgColor=PETROL)
HEAD_FONT = Font(bold=True, color="FFFFFF")


def q(cur, sql):
    cur.execute(sql)
    return cur.fetchall()


def header_row(ws, row, headers, start=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start + i, value=h)
        c.fill = HEAD_FILL; c.font = HEAD_FONT; c.border = BORDER
        c.alignment = Alignment(horizontal="center")


def main() -> None:
    OUT.mkdir(exist_ok=True)
    conn = psycopg2.connect(DSN); cur = conn.cursor()

    # --- Extraction + agrégations (le travail manuel, en SQL) ---
    kpis = q(cur, """
        SELECT round(sum(oi.quantity*oi.unit_price),2), count(DISTINCT o.id),
               round(sum(oi.quantity*oi.unit_price)/count(DISTINCT o.id),2)
        FROM orders o JOIN order_item oi ON oi.order_id=o.id""")[0]
    mensuel = q(cur, """
        SELECT to_char(date_trunc('month',o.order_date),'YYYY-MM') AS mois,
               round(sum(oi.quantity*oi.unit_price),0) AS ca,
               count(DISTINCT o.id) AS commandes
        FROM orders o JOIN order_item oi ON oi.order_id=o.id
        GROUP BY 1 ORDER BY 1""")
    top = q(cur, """
        SELECT pr.name, round(sum(oi.quantity*oi.unit_price),0) AS ca, sum(oi.quantity) AS unites
        FROM order_item oi JOIN product pr ON pr.id=oi.product_id
        GROUP BY 1 ORDER BY 2 DESC LIMIT 10""")
    categorie = q(cur, """
        SELECT c.name, round(sum(oi.quantity*oi.unit_price),0) AS ca
        FROM order_item oi JOIN product pr ON pr.id=oi.product_id
        JOIN category c ON c.id=pr.category_id
        GROUP BY 1 ORDER BY 2 DESC""")
    cur.close(); conn.close()

    ca_total, nb_cmd, panier = float(kpis[0]), int(kpis[1]), float(kpis[2])
    # Croissance sur les deux derniers mois COMPLETS : le mois en cours, inachevé, ferait chuter le chiffre.
    complets = [m for m in mensuel if m[0] < date.today().strftime("%Y-%m")]
    mom = None
    if len(complets) >= 2 and complets[-2][1]:
        mom = 100.0 * float(complets[-1][1] - complets[-2][1]) / float(complets[-2][1])

    wb = Workbook()

    # ============ Feuille 1 : Synthèse ============
    ws = wb.active; ws.title = "Synthèse"; ws.sheet_view.showGridLines = False
    ws["B2"] = "Rapport mensuel des ventes"
    ws["B2"].font = Font(bold=True, size=18, color=PETROL_D)
    ws["B3"] = f"Généré automatiquement le {date.today().isoformat()}"
    ws["B3"].font = Font(italic=True, color="5E6B7A")

    cards = [("Chiffre d'affaires", ca_total, EURO),
             ("Commandes", nb_cmd, '#,##0'),
             ("Panier moyen", panier, '#,##0.00" €"'),
             ("Croissance dernier mois complet", (mom / 100 if mom is not None else None), '0.0%')]
    for i, (label, val, fmt) in enumerate(cards):
        col = 2 + i * 2
        lc = ws.cell(row=5, column=col, value=label)
        lc.font = Font(bold=True, color="5E6B7A", size=9)
        vc = ws.cell(row=6, column=col, value=val)
        vc.font = Font(bold=True, size=16, color=PETROL); vc.number_format = fmt
        ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)
        ws.merge_cells(start_row=6, start_column=col, end_row=6, end_column=col + 1)
    for col in range(2, 10):
        ws.column_dimensions[get_column_letter(col)].width = 15

    # ============ Feuille 2 : CA mensuel + graphique ============
    ws2 = wb.create_sheet("CA mensuel"); ws2.sheet_view.showGridLines = False
    header_row(ws2, 1, ["Mois", "Chiffre d'affaires", "Commandes"])
    for r, (mois, ca, cmd) in enumerate(mensuel, start=2):
        ws2.cell(row=r, column=1, value=mois).border = BORDER
        c = ws2.cell(row=r, column=2, value=float(ca)); c.number_format = EURO; c.border = BORDER
        ws2.cell(row=r, column=3, value=int(cmd)).border = BORDER
    ws2.column_dimensions["A"].width = 12; ws2.column_dimensions["B"].width = 18; ws2.column_dimensions["C"].width = 12
    chart = LineChart(); chart.title = "CA par mois"; chart.height = 8; chart.width = 18
    data = Reference(ws2, min_col=2, min_row=1, max_row=len(mensuel) + 1)
    cats = Reference(ws2, min_col=1, min_row=2, max_row=len(mensuel) + 1)
    chart.add_data(data, titles_from_data=True); chart.set_categories(cats)
    chart.series[0].graphicalProperties.line.solidFill = PETROL
    ws2.add_chart(chart, "E2")

    # ============ Feuille 3 : Top produits + barres de données ============
    ws3 = wb.create_sheet("Top produits"); ws3.sheet_view.showGridLines = False
    header_row(ws3, 1, ["Produit", "Chiffre d'affaires", "Unités"])
    for r, (name, ca, unites) in enumerate(top, start=2):
        ws3.cell(row=r, column=1, value=name).border = BORDER
        c = ws3.cell(row=r, column=2, value=float(ca)); c.number_format = EURO; c.border = BORDER
        ws3.cell(row=r, column=3, value=int(unites)).border = BORDER
    ws3.column_dimensions["A"].width = 34; ws3.column_dimensions["B"].width = 18; ws3.column_dimensions["C"].width = 10
    ws3.conditional_formatting.add(
        f"B2:B{len(top)+1}",
        DataBarRule(start_type="min", end_type="max", color=AMBRE))

    # ============ Feuille 4 : Par catégorie + histogramme ============
    ws4 = wb.create_sheet("Par catégorie"); ws4.sheet_view.showGridLines = False
    header_row(ws4, 1, ["Catégorie", "Chiffre d'affaires"])
    for r, (name, ca) in enumerate(categorie, start=2):
        ws4.cell(row=r, column=1, value=name).border = BORDER
        c = ws4.cell(row=r, column=2, value=float(ca)); c.number_format = EURO; c.border = BORDER
    ws4.column_dimensions["A"].width = 18; ws4.column_dimensions["B"].width = 18
    bar = BarChart(); bar.title = "CA par catégorie"; bar.height = 8; bar.width = 16; bar.legend = None
    data = Reference(ws4, min_col=2, min_row=1, max_row=len(categorie) + 1)
    cats = Reference(ws4, min_col=1, min_row=2, max_row=len(categorie) + 1)
    bar.add_data(data, titles_from_data=True); bar.set_categories(cats)
    bar.series[0].graphicalProperties.solidFill = PETROL
    ws4.add_chart(bar, "D2")

    path = OUT / f"rapport_ventes_{date.today().isoformat()}.xlsx"
    wb.save(path)
    print(f"Rapport généré : {path.name}")
    print(f"  CA {ca_total:,.0f} € · {nb_cmd:,} commandes · panier {panier:,.2f} €"
          + (f" · MoM {mom:+.1f}%" if mom is not None else ""))


if __name__ == "__main__":
    main()
