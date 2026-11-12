import sys
from pathlib import Path

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from src.estilo import PALETA, aplicar_estilo  # noqa: E402

FIG = Path(__file__).resolve().parents[1] / "reports" / "figuras"


def caja(ax, x, y, w, h, titulo, sub, color, texto="white", ft=10.5):
    ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle="round,pad=0.02,rounding_size=0.12",
                                fc=color, ec="none", zorder=3))
    ax.text(x, y + 0.17, titulo, ha="center", va="center", color=texto, fontsize=ft,
            fontweight="bold", zorder=4)
    ax.text(x, y - 0.2, sub, ha="center", va="center", color=texto, fontsize=8, zorder=4, linespacing=1.3)


def flecha(ax, p1, p2, texto=None, doble=False, color=PALETA["oscuro"], off=(0, 0.12)):
    ax.add_patch(FancyArrowPatch(p1, p2, arrowstyle="<|-|>" if doble else "-|>", mutation_scale=14,
                                 lw=1.6, color=color, zorder=2))
    if texto:
        ax.text((p1[0] + p2[0]) / 2 + off[0], (p1[1] + p2[1]) / 2 + off[1], texto, ha="center",
                va="bottom", fontsize=8, color=color, style="italic", linespacing=1.1)


def arquitectura():
    fig, ax = plt.subplots(figsize=(15, 6.4))
    ax.set_xlim(0, 15)
    ax.set_ylim(0, 6.2)
    ax.axis("off")
    for y0, txt, c in [(3.45, "ENTRENAMIENTO (offline · CRISP-DM: preparación, modelado y evaluación)", "#E6EEF1"),
                       (0.35, "PRODUCCIÓN (online · CRISP-DM: despliegue)", "#FBEBDD")]:
        ax.add_patch(FancyBboxPatch((0.25, y0), 14.5, 2.35, boxstyle="round,pad=0,rounding_size=0.18",
                                    fc=c, ec="none", zorder=0))
        ax.text(0.5, y0 + 2.12, txt, fontsize=8.5, fontweight="bold", color=PALETA["oscuro"], va="center")

    W, H, yt, yb = 2.75, 1.15, 4.5, 1.4
    xs = [1.85, 5.6, 9.45, 13.1]
    caja(ax, xs[0], yt, W, H, "MySQL (XAMPP)", "BD mineria_datos\ntabla players (50 149 filas)", PALETA["oscuro"])
    caja(ax, xs[1], yt, W, H, "connection2.py", "MySQLConnector\nextract_data(query)", PALETA["primario"])
    caja(ax, xs[2], yt, W + 0.35, H, "exportar_modelo.py",
         "imputación por mediana agrupada\nOne-Hot · Random Forest optimizado", PALETA["primario"])
    caja(ax, xs[3], yt, W, H, "modelo_scouting_rf.joblib", "modelo + columnas_X\n+ medianas + umbral Q3",
         PALETA["acento"], ft=9.5)
    flecha(ax, (xs[0] + W / 2, yt), (xs[1] - W / 2, yt), "consulta\nSQL")
    flecha(ax, (xs[1] + W / 2, yt), (xs[2] - W / 2 - 0.18, yt), "Data\nFrame")
    flecha(ax, (xs[2] + W / 2 + 0.18, yt), (xs[3] - W / 2, yt), "joblib\n.dump")

    xb = [13.1, 8.0, 3.9, 1.2]
    caja(ax, xb[0], yb, W, H, "main_api.py", "FastAPI + Pydantic\nPOST /predict · GET / · /info", PALETA["elite"])
    caja(ax, xb[2], yb, W + 0.2, H, "index.html", "formulario + fetch()\nbarra de progreso + historial", PALETA["elite"])
    ax.add_patch(plt.Circle((xb[3], yb), 0.55, fc=PALETA["oscuro"], zorder=3))
    ax.text(xb[3], yb + 0.08, "Scout", ha="center", va="center", color="white", fontsize=9.5, fontweight="bold", zorder=4)
    ax.text(xb[3], yb - 0.17, "usuario", ha="center", va="center", color="white", fontsize=7.5, zorder=4)
    flecha(ax, (xs[3], yt - H / 2), (xb[0], yb + H / 2))
    ax.text(xs[3] + 0.12, (yt + yb) / 2, "joblib.load\n(una vez, lifespan)", fontsize=8, style="italic",
            color=PALETA["oscuro"], va="center")
    flecha(ax, (xb[2] + (W + 0.2) / 2, yb + 0.22), (xb[0] - W / 2, yb + 0.22), "request JSON (HTTP POST /predict)")
    flecha(ax, (xb[0] - W / 2, yb - 0.22), (xb[2] + (W + 0.2) / 2, yb - 0.22),
           "response JSON (predicción + % de probabilidad)", off=(0, -0.42))
    flecha(ax, (xb[3] + 0.55, yb), (xb[2] - (W + 0.2) / 2, yb), doble=True)
    ax.set_title("Arquitectura de despliegue y flujo de datos del ecosistema de scouting",
                 fontsize=13, fontweight="bold", pad=8)
    fig.savefig(FIG / "08_arquitectura_despliegue.png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def crisp_dm():
    fases = [
        ("1. Comprensión\ndel negocio", "Scouting: identificar\nel Top 25 % de valor"),
        ("2. Comprensión\nde los datos", "MySQL + connection2.py\nEDA (figuras 01-02)"),
        ("3. Preparación\nde los datos", "inferencia.py: mediana\nagrupada + One-Hot"),
        ("4. Modelado", "experimentos_modelado.py\nÁrboles, RF, GridSearchCV"),
        ("5. Evaluación", "Train vs Test, ROC-AUC,\nimportancia Gini"),
        ("6. Despliegue", "exportar_modelo.py · joblib\nmain_api.py · index.html"),
    ]
    colores = [PALETA["oscuro"], PALETA["primario"], PALETA["primario"], PALETA["acento"], PALETA["acento"], PALETA["elite"]]
    fig, ax = plt.subplots(figsize=(10, 10))
    ax.set_xlim(-5.4, 5.4)
    ax.set_ylim(-5.6, 5.4)
    ax.set_aspect("equal")
    ax.axis("off")
    R, r = 3.75, 1.3
    pts = [np.array([R * np.cos(np.radians(90 - i * 60)), R * np.sin(np.radians(90 - i * 60))]) for i in range(6)]
    for i in range(6):
        a, b = pts[i], pts[(i + 1) % 6]
        u = (b - a) / np.linalg.norm(b - a)
        ax.add_patch(FancyArrowPatch(a + u * (r + 0.1), b - u * (r + 0.1), connectionstyle="arc3,rad=-0.25",
                                     arrowstyle="-|>", mutation_scale=20, lw=2.4, color="#7D8A97", zorder=5))
    a, b = pts[4], pts[0]
    u = (b - a) / np.linalg.norm(b - a)
    ax.add_patch(FancyArrowPatch(a + u * (r + 0.1), b - u * (r + 0.1), connectionstyle="arc3,rad=-0.12",
                                 arrowstyle="-|>", mutation_scale=14, lw=1.4, ls="--", color="#9AA5B1", zorder=1))
    ax.add_patch(plt.Circle((0, 0), 1.2, fc=PALETA["claro"], ec=PALETA["primario"], lw=2, zorder=2))
    ax.text(0, 0.22, "DATOS", ha="center", va="center", fontsize=14, fontweight="bold", color=PALETA["primario"], zorder=3)
    ax.text(0, -0.3, "Transfermarkt\n(players · MySQL)", ha="center", va="center", fontsize=8.5,
            color=PALETA["oscuro"], zorder=3)
    for (x, y), (t, s), c in zip(pts, fases, colores):
        ax.add_patch(plt.Circle((x, y), r, fc=c, ec="white", lw=3, zorder=3))
        ax.text(x, y + 0.3, t, ha="center", va="center", color="white", fontsize=10, fontweight="bold",
                zorder=4, linespacing=1.15)
        ax.text(x, y - 0.5, s, ha="center", va="center", color="white", fontsize=7.4, zorder=4, linespacing=1.25)
    ax.set_title("Ciclo CRISP-DM aplicado al proyecto de scouting deportivo", fontsize=14,
                 fontweight="bold", pad=4)
    ax.text(0, -5.45, "Flechas sólidas: secuencia del ciclo · Flecha punteada: retroalimentación "
            "(la evaluación puede reabrir la comprensión del negocio)", ha="center", fontsize=8.5, color="#5B6B79")
    fig.savefig(FIG / "09_crisp_dm.png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    aplicar_estilo()
    arquitectura()
    crisp_dm()
    print("Diagramas generados en", FIG)
