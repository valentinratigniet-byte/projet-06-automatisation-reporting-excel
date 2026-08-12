"""
Mini-app Streamlit (MVP) — déclencher la génération du rapport et le télécharger,
sans ligne de commande. Lancer : streamlit run src/app.py
"""
import glob
import os
import subprocess
import sys
from pathlib import Path

import psycopg2
import streamlit as st

ROOT = Path(__file__).resolve().parent.parent
DSN = os.environ.get("DATABASE_URL",
                     "postgresql://portfolio:portfolio@127.0.0.1:5433/ecommerce")

st.set_page_config(page_title="Rapport ventes automatisé", page_icon="📊")
st.title("📊 Rapport de ventes automatisé")
st.caption("Remplace le reporting Excel manuel : un clic → fichier mis en forme.")

# Aperçu des KPIs (lecture directe de la base)
try:
    con = psycopg2.connect(DSN); cur = con.cursor()
    cur.execute("""SELECT round(sum(oi.quantity*oi.unit_price),0), count(DISTINCT o.id)
                   FROM orders o JOIN order_item oi ON oi.order_id=o.id""")
    ca, cmd = cur.fetchone(); con.close()
    c1, c2 = st.columns(2)
    c1.metric("Chiffre d'affaires", f"{ca:,.0f} €".replace(",", " "))
    c2.metric("Commandes", f"{cmd:,}".replace(",", " "))
except Exception as e:
    st.warning(f"Base indisponible : {e}")

if st.button("🚀 Générer le rapport Excel", type="primary"):
    r = subprocess.run([sys.executable, str(ROOT / "src" / "generate_report.py")],
                       capture_output=True, text=True)
    if r.returncode == 0:
        st.success(r.stdout.strip())
        files = sorted(glob.glob(str(ROOT / "output" / "*.xlsx")))
        if files:
            with open(files[-1], "rb") as f:
                st.download_button("⬇️ Télécharger le rapport", f,
                                   file_name=os.path.basename(files[-1]))
    else:
        st.error(r.stderr.strip())
