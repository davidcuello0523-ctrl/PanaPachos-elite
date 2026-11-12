import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns

PALETA = {
    "primario": "#0F4C5C",
    "elite": "#E36414",
    "acento": "#2A9D8F",
    "neutro": "#8D99AE",
    "oscuro": "#1D2D3A",
    "claro": "#EEF2F5",
}
COLORES_MODELOS = [PALETA["neutro"], PALETA["acento"], PALETA["primario"], PALETA["elite"]]

ETIQUETAS_POSICION = {
    "Attack": "Delantero",
    "Defender": "Defensa",
    "Goalkeeper": "Portero",
    "Midfield": "Mediocampista",
    "Missing": "Sin posición registrada",
}


def etiqueta_variable(col: str) -> str:
    if col.startswith("posicion_"):
        pos = col.removeprefix("posicion_")
        return f"Posición: {ETIQUETAS_POSICION.get(pos, pos)}"
    return {"edad": "Edad (años)", "estatura_cm": "Estatura (cm)"}.get(col, col)


def aplicar_estilo():
    sns.set_theme(style="whitegrid", font_scale=1.0)
    plt.rcParams.update({
        "figure.dpi": 150, "savefig.dpi": 150, "savefig.bbox": "tight",
        "axes.titleweight": "bold", "axes.titlesize": 12, "axes.titlepad": 12,
        "axes.edgecolor": "#B0B8C1", "grid.color": "#E3E8EC",
        "font.family": "DejaVu Sans",
    })


def guardar(fig, ruta):
    sns.despine(fig=fig)
    fig.savefig(ruta, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
