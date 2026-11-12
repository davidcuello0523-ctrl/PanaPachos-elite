import json
import time
import warnings

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from sklearn.dummy import DummyClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.inspection import permutation_importance
from sklearn.metrics import (accuracy_score, classification_report, confusion_matrix,
                             f1_score, precision_score, recall_score, roc_auc_score, roc_curve)
from sklearn.model_selection import GridSearchCV, StratifiedKFold, train_test_split
from sklearn.tree import DecisionTreeClassifier

from src.database.connection2 import MySQLConnector
from src.estilo import (COLORES_MODELOS, ETIQUETAS_POSICION, PALETA, aplicar_estilo,
                        etiqueta_variable, guardar)
from src.preprocessing.inferencia import calcular_medianas, codificar, imputar, limpiar_tipos

warnings.filterwarnings("ignore", message=".*sklearn.utils.parallel.delayed.*")

FIG = "reports/figuras"
RES = "reports/resultados"
SEED = 42

QUERY = """
    SELECT
        p.player_id,
        p.name AS jugador,
        p.position AS posicion,
        p.height_in_cm AS estatura_cm,
        TIMESTAMPDIFF(YEAR, p.date_of_birth, CURDATE()) AS edad,
        p.market_value_in_eur AS valor_mercado
    FROM players p
    WHERE p.market_value_in_eur IS NOT NULL
      AND p.date_of_birth IS NOT NULL;
"""


def miles(n):
    return f"{n:,.0f}".replace(",", " ")


def titulo(texto):
    print("\n" + "=" * 70 + f"\n  {texto}\n" + "=" * 70)


def metricas_modelo(nombre, modelo, X_tr, y_tr, X_te, y_te):
    pred_tr, pred_te = modelo.predict(X_tr), modelo.predict(X_te)
    proba_te = modelo.predict_proba(X_te)[:, 1]
    m = {
        "modelo": nombre,
        "accuracy_train": accuracy_score(y_tr, pred_tr),
        "accuracy_test": accuracy_score(y_te, pred_te),
        "precision_elite": precision_score(y_te, pred_te, zero_division=0),
        "recall_elite": recall_score(y_te, pred_te),
        "f1_elite": f1_score(y_te, pred_te),
        "auc_test": roc_auc_score(y_te, proba_te),
        "matriz_confusion": confusion_matrix(y_te, pred_te).tolist(),
    }
    m["brecha"] = m["accuracy_train"] - m["accuracy_test"]
    print(f"{nombre:<38} Train={m['accuracy_train']:.4f}  Test={m['accuracy_test']:.4f}  "
          f"Brecha={m['brecha']:+.4f}  P={m['precision_elite']:.3f}  R={m['recall_elite']:.3f}  "
          f"F1={m['f1_elite']:.3f}  AUC={m['auc_test']:.3f}")
    return m, proba_te


def main():
    aplicar_estilo()
    metricas = {}

    titulo("1. INGESTA DESDE MySQL Y EDA MÍNIMO")
    df = limpiar_tipos(MySQLConnector().extract_data(QUERY))
    df["valor_mercado"] = pd.to_numeric(df["valor_mercado"], errors="coerce")
    print(f"Registros extraídos: {len(df):,}")
    nulos_antes = df[["posicion", "estatura_cm", "edad", "valor_mercado"]].isna().sum()
    print("Nulos por columna (antes de imputar):\n" + nulos_antes.to_string())
    dist_pos = df["posicion"].value_counts()
    print("Distribución de posiciones:\n" + dist_pos.to_string())

    umbral = df["valor_mercado"].quantile(0.75)
    df["es_elite"] = (df["valor_mercado"] >= umbral).astype(int)
    prop = df["es_elite"].value_counts(normalize=True).sort_index()
    print(f"Umbral Q3 (P75) de valor de mercado: € {umbral:,.0f}")
    print(f"Proporción de clases: Estándar={prop[0]:.2%}  Élite={prop[1]:.2%}")

    medianas_pos = df.groupby("posicion")["estatura_cm"].median()
    metricas["datos"] = {
        "registros": int(len(df)),
        "nulos_antes": {k: int(v) for k, v in nulos_antes.items()},
        "distribucion_posiciones": {k: int(v) for k, v in dist_pos.items()},
        "umbral_q3_eur": float(umbral),
        "n_elite": int(df["es_elite"].sum()),
        "n_estandar": int((df["es_elite"] == 0).sum()),
        "proporcion_elite": float(prop[1]),
        "mediana_estatura_por_posicion_total": {k: float(v) for k, v in medianas_pos.items()},
        "estatura_media": float(df["estatura_cm"].mean()),
        "estaturas_implausibles_menor_140": int((df["estatura_cm"] < 140).sum()),
    }
    print(f"Estaturas implausibles (< 140 cm): {metricas['datos']['estaturas_implausibles_menor_140']}")

    rangos = pd.cut(df["edad"], [0, 20, 23, 26, 29, 32, 35, 100],
                    labels=["≤20", "21-23", "24-26", "27-29", "30-32", "33-35", "≥36"])
    tasa_pos = df.groupby("posicion")["es_elite"].agg(["mean", "size"])
    tasa_edad = df.groupby(rangos, observed=True)["es_elite"].agg(["mean", "size"])
    print("Tasa de Élite por posición:\n" + tasa_pos.round(4).to_string())
    print("Tasa de Élite por rango de edad:\n" + tasa_edad.round(4).to_string())
    metricas["datos"]["tasa_elite_posicion"] = {k: {"tasa": float(r["mean"]), "n": int(r["size"])}
                                                for k, r in tasa_pos.iterrows()}
    metricas["datos"]["tasa_elite_edad"] = {str(k): {"tasa": float(r["mean"]), "n": int(r["size"])}
                                            for k, r in tasa_edad.iterrows()}
    metricas["datos"]["edad_mediana"] = float(df["edad"].median())
    metricas["datos"]["valor_mediano_eur"] = float(df["valor_mercado"].median())
    metricas["datos"]["combinaciones_unicas_X"] = int(
        df[["edad", "estatura_cm", "posicion"]].drop_duplicates().shape[0])

    fig, ax = plt.subplots(figsize=(6.5, 4.2))
    conteo = df["es_elite"].value_counts().sort_index()
    barras = ax.bar(["Estándar (0)", "Élite (1)"], conteo.values,
                    color=[PALETA["primario"], PALETA["elite"]], width=0.55)
    for b, n in zip(barras, conteo.values):
        ax.text(b.get_x() + b.get_width() / 2, n + 400, f"{miles(n)}\n({n / len(df):.1%})",
                ha="center", va="bottom", fontsize=10, fontweight="bold")
    ax.set_ylim(0, conteo.max() * 1.22)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: miles(v)))
    ax.set_ylabel("Número de jugadores")
    ax.set_title(f"Distribución del target es_elite (umbral Q3 = € {miles(umbral)})")
    guardar(fig, f"{FIG}/01_distribucion_clases.png")

    orden = medianas_pos.sort_values().index.tolist()
    fig, ax = plt.subplots(figsize=(9, 4.8))
    sns.boxplot(data=df, x="posicion", y="estatura_cm", order=orden, ax=ax, width=0.5,
                color=PALETA["claro"], linecolor=PALETA["primario"], fliersize=1.5,
                flierprops={"alpha": 0.3})
    ax.axhline(df["estatura_cm"].median(), ls="--", color=PALETA["elite"], lw=1.3,
               label=f"Mediana global = {df['estatura_cm'].median():.0f} cm")
    for i, pos in enumerate(orden):
        ax.text(i + 0.29, medianas_pos[pos], f"{medianas_pos[pos]:.0f} cm", ha="left",
                va="center", fontsize=9, fontweight="bold", color=PALETA["primario"])
    ax.set_xticks(range(len(orden)), [ETIQUETAS_POSICION.get(p, p).replace(" ", "\n", 1)
                                      if p == "Missing" else ETIQUETAS_POSICION.get(p, p) for p in orden])
    ax.set_ylim(150, 212)
    ax.set_xlabel("Posición")
    ax.set_ylabel("Estatura (cm)")
    ax.set_title("Estatura por posición: la mediana difiere entre grupos")
    ax.legend(loc="upper left", frameon=False)
    guardar(fig, f"{FIG}/02_estatura_por_posicion.png")

    titulo("2-3. SPLIT ESTRATIFICADO Y PREPROCESAMIENTO SIN FUGA DE DATOS")
    df_tr, df_te = train_test_split(df, test_size=0.2, stratify=df["es_elite"], random_state=SEED)
    med = calcular_medianas(df_tr)
    df_tr = imputar(df_tr, **med)
    df_te = imputar(df_te, **med)
    nulos_despues = pd.concat([df_tr, df_te])[["estatura_cm", "edad"]].isna().sum()
    print("Medianas de estatura por posición (train):",
          {k: round(v, 1) for k, v in med["medianas_estatura"].items()})
    print(f"Mediana global estatura={med['mediana_global_estatura']}  edad={med['mediana_global_edad']}")
    print("Nulos después de imputar:\n" + nulos_despues.to_string())

    X_tr = codificar(df_tr)
    X_te = codificar(df_te, X_tr.columns.tolist())
    y_tr, y_te = df_tr["es_elite"], df_te["es_elite"]
    print("Columnas X:", X_tr.columns.tolist())
    print(f"Train: {len(X_tr):,} filas | Élite={y_tr.mean():.4%}")
    print(f"Test : {len(X_te):,} filas | Élite={y_te.mean():.4%}")
    metricas["datos"].update({
        "nulos_despues": {k: int(v) for k, v in nulos_despues.items()},
        "medianas_estatura_train": med["medianas_estatura"],
        "mediana_global_estatura_train": med["mediana_global_estatura"],
        "mediana_global_edad_train": med["mediana_global_edad"],
        "columnas_X": X_tr.columns.tolist(),
        "n_train": int(len(X_tr)), "n_test": int(len(X_te)),
        "proporcion_elite_train": float(y_tr.mean()), "proporcion_elite_test": float(y_te.mean()),
    })

    titulo("4. PARADOJA DE LA EXACTITUD (DummyClassifier most_frequent)")
    dummy = DummyClassifier(strategy="most_frequent").fit(X_tr, y_tr)
    pred_d = dummy.predict(X_te)
    metricas["dummy"] = {
        "accuracy_test": accuracy_score(y_te, pred_d),
        "recall_elite": recall_score(y_te, pred_d),
        "precision_elite": precision_score(y_te, pred_d, zero_division=0),
        "f1_elite": f1_score(y_te, pred_d, zero_division=0),
        "matriz_confusion": confusion_matrix(y_te, pred_d).tolist(),
    }
    print(f"Accuracy = {metricas['dummy']['accuracy_test']:.4f}  pero Recall Élite = "
          f"{metricas['dummy']['recall_elite']:.1f}: nunca detecta a un solo jugador Élite.")

    titulo("5. DIAGNÓSTICO DE OVERFITTING")
    arboles = {
        "Árbol sin restricciones": DecisionTreeClassifier(random_state=SEED),
        "Árbol podado": DecisionTreeClassifier(max_depth=4, min_samples_leaf=20, random_state=SEED),
        "Árbol podado + balanced": DecisionTreeClassifier(max_depth=4, min_samples_leaf=20,
                                                          class_weight="balanced", random_state=SEED),
    }
    filas, probas = [], {}
    for nombre, modelo in arboles.items():
        modelo.fit(X_tr, y_tr)
        m, probas[nombre] = metricas_modelo(nombre, modelo, X_tr, y_tr, X_te, y_te)
        m["profundidad"] = int(modelo.get_depth())
        m["hojas"] = int(modelo.get_n_leaves())
        filas.append(m)
    print(f"Profundidad árbol libre = {filas[0]['profundidad']} ({filas[0]['hojas']} hojas) "
          f"vs. podado = {filas[1]['profundidad']} ({filas[1]['hojas']} hojas)")

    titulo("6. RANDOM FOREST + GridSearchCV (StratifiedKFold k=5, scoring=roc_auc)")
    malla = {
        "n_estimators": [100, 200, 300],
        "max_depth": [4, 6, 8, None],
        "min_samples_leaf": [1, 5, 10],
        "max_features": ["sqrt", None],
    }
    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=SEED)
    grid = GridSearchCV(
        RandomForestClassifier(class_weight="balanced", random_state=SEED, n_jobs=-1),
        malla, cv=cv, scoring={"roc_auc": "roc_auc", "f1": "f1"}, refit="roc_auc",
        n_jobs=1, verbose=1,
    )
    t0 = time.time()
    grid.fit(X_tr, y_tr)
    duracion = time.time() - t0
    n_comb = int(np.prod([len(v) for v in malla.values()]))
    i_best = grid.best_index_
    print(f"{n_comb} combinaciones × 5 folds = {n_comb * 5} ajustes en {duracion / 60:.1f} min")
    print("Mejores parámetros:", grid.best_params_)
    print(f"Mejor ROC-AUC CV = {grid.best_score_:.4f}  | F1 CV de esa combinación = "
          f"{grid.cv_results_['mean_test_f1'][i_best]:.4f}")

    res = pd.DataFrame(grid.cv_results_)
    top10 = (res.sort_values("rank_test_roc_auc")
             [["rank_test_roc_auc", "param_n_estimators", "param_max_depth", "param_min_samples_leaf",
               "param_max_features", "mean_test_roc_auc", "std_test_roc_auc", "mean_test_f1",
               "mean_fit_time"]].head(10))
    top10.to_csv(f"{RES}/gridsearch_resultados.csv", index=False, encoding="utf-8-sig")
    print(top10.to_string(index=False))

    rf = grid.best_estimator_
    m_rf, probas["Random Forest optimizado"] = metricas_modelo(
        "Random Forest optimizado", rf, X_tr, y_tr, X_te, y_te)
    filas.append(m_rf)
    print("\nClassification report – RF optimizado (Test):")
    reporte = classification_report(y_te, rf.predict(X_te), target_names=["Estándar", "Élite"], digits=4)
    print(reporte)

    mejores = {k: (None if v is None else v) for k, v in grid.best_params_.items()}
    json.dump({"mejores_parametros": mejores, "mejor_auc_cv": grid.best_score_,
               "f1_cv": float(grid.cv_results_["mean_test_f1"][i_best])},
              open(f"{RES}/mejores_parametros.json", "w", encoding="utf-8"), indent=2, ensure_ascii=False)
    metricas["gridsearch"] = {
        "malla": {k: [str(x) for x in v] for k, v in malla.items()},
        "n_combinaciones": n_comb, "folds": 5, "ajustes": n_comb * 5,
        "duracion_min": duracion / 60,
        "mejores_parametros": mejores,
        "mejor_auc_cv": float(grid.best_score_),
        "std_auc_cv": float(res.loc[i_best, "std_test_roc_auc"]),
        "f1_cv": float(grid.cv_results_["mean_test_f1"][i_best]),
        "classification_report": classification_report(
            y_te, rf.predict(X_te), target_names=["Estándar", "Élite"], output_dict=True),
    }

    tabla = pd.DataFrame(filas)[["modelo", "accuracy_train", "accuracy_test", "brecha",
                                 "precision_elite", "recall_elite", "f1_elite", "auc_test"]]
    tabla.to_csv(f"{RES}/tabla_overfitting.csv", index=False, encoding="utf-8-sig", float_format="%.4f")
    metricas["modelos"] = {f["modelo"]: f for f in filas}

    fig, ax = plt.subplots(figsize=(9, 4.8))
    x = np.arange(len(tabla))
    w = 0.36
    b1 = ax.bar(x - w / 2, tabla["accuracy_train"], w, label="Train", color=PALETA["primario"])
    b2 = ax.bar(x + w / 2, tabla["accuracy_test"], w, label="Test", color=PALETA["elite"])
    for bs in (b1, b2):
        for b in bs:
            ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.01, f"{b.get_height():.3f}",
                    ha="center", va="bottom", fontsize=8.5)
    for i, br in enumerate(tabla["brecha"]):
        txt = "≈ 0.000" if abs(br) < 0.0005 else f"{br:+.3f}"
        ax.text(i, 0.86, f"brecha {txt}", ha="center", fontsize=8.5,
                color=PALETA["oscuro"], fontweight="bold")
    ax.set_xticks(x, [m.replace(" + ", "\n+ ").replace(" optimizado", "\noptimizado") for m in tabla["modelo"]])
    ax.set_ylim(0, 1.0)
    ax.set_ylabel("Accuracy")
    ax.set_title("Diagnóstico de overfitting: Accuracy Train vs. Test")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, 1.0), frameon=False, ncol=2)
    guardar(fig, f"{FIG}/03_overfitting_train_vs_test.png")

    fig, axes = plt.subplots(1, 3, figsize=(14, 4.4), gridspec_kw={"wspace": 0.4})
    for ax, nombre in zip(axes, ["Árbol podado", "Árbol podado + balanced", "Random Forest optimizado"]):
        cm = np.array(metricas["modelos"][nombre]["matriz_confusion"])
        anot = np.array([[f"{miles(v)}\n({v / cm[i].sum():.1%})" for v in fila] for i, fila in enumerate(cm)])
        sns.heatmap(cm / cm.sum(axis=1, keepdims=True), annot=anot, fmt="", cbar=False, ax=ax,
                    cmap=sns.light_palette(PALETA["primario"], as_cmap=True), vmin=0, vmax=1,
                    linewidths=1, linecolor="white", annot_kws={"fontsize": 10})
        ax.set_xticklabels(["Estándar", "Élite"])
        ax.set_yticklabels(["Estándar", "Élite"], rotation=0)
        ax.set_xlabel("Predicción")
        ax.set_ylabel("Real")
        r = metricas["modelos"][nombre]["recall_elite"]
        ax.set_title(f"{nombre}\nRecall Élite = {r:.3f}", fontsize=11)
    fig.suptitle("Matrices de confusión en Test: efecto de class_weight='balanced'",
                 fontweight="bold", y=1.12)
    guardar(fig, f"{FIG}/04_matrices_confusion.png")

    titulo("7. CURVAS ROC COMPARATIVAS")
    fig, ax = plt.subplots(figsize=(6.6, 6))
    for (nombre, color) in zip(["Árbol sin restricciones", "Árbol podado", "Random Forest optimizado"],
                               [PALETA["neutro"], PALETA["acento"], PALETA["elite"]]):
        fpr, tpr, _ = roc_curve(y_te, probas[nombre])
        auc = metricas["modelos"][nombre]["auc_test"]
        ax.plot(fpr, tpr, color=color, lw=2.2, label=f"{nombre} (AUC = {auc:.3f})")
        print(f"AUC {nombre}: {auc:.4f}")
    ax.plot([0, 1], [0, 1], ls="--", color="#9AA5B1", lw=1.2, label="Azar (AUC = 0.500)")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1.01)
    ax.set_xlabel("Tasa de Falsos Positivos (FPR)")
    ax.set_ylabel("Tasa de Verdaderos Positivos (TPR / Recall)")
    ax.set_title("Curvas ROC en Test: Árbol único vs. Random Forest")
    ax.legend(loc="lower right", frameon=True, fontsize=9)
    guardar(fig, f"{FIG}/05_curvas_roc.png")

    titulo("8. IMPORTANCIA DE VARIABLES")
    gini = pd.Series(rf.feature_importances_, index=X_tr.columns).sort_values()
    perm = permutation_importance(rf, X_te, y_te, scoring="roc_auc", n_repeats=10,
                                  random_state=SEED, n_jobs=-1)
    perm_s = pd.Series(perm.importances_mean, index=X_tr.columns)
    perm_std = pd.Series(perm.importances_std, index=X_tr.columns)
    print("Gini:\n" + gini.sort_values(ascending=False).round(4).to_string())
    print("Permutación (caída de ROC-AUC):\n" + perm_s.sort_values(ascending=False).round(4).to_string())
    top3_gini = gini.sort_values(ascending=False).index[:3].tolist()
    top3_perm = perm_s.sort_values(ascending=False).index[:3].tolist()
    coincide = set(top3_gini) == set(top3_perm)
    print(f"Top 3 Gini: {top3_gini} | Top 3 permutación: {top3_perm}")
    print(f"¿Mismas 3 variables en el top? {'Sí' if coincide else 'No'} "
          f"(orden idéntico: {'Sí' if top3_gini == top3_perm else 'No'})")

    for serie, err, archivo, xlabel, ttl in [
        (gini, None, "06_feature_importance.png", "Importancia Gini (reducción media de impureza)",
         "Importancia de variables (Gini) – Random Forest optimizado"),
        (perm_s.sort_values(), perm_std, "07_permutation_importance.png",
         "Caída media del ROC-AUC al permutar la variable (Test)",
         "Importancia por permutación – Random Forest optimizado"),
    ]:
        fig, ax = plt.subplots(figsize=(8, 4.2))
        colores = [PALETA["elite"] if i >= len(serie) - 2 else PALETA["primario"] for i in range(len(serie))]
        ax.barh([etiqueta_variable(c) for c in serie.index], serie.values, color=colores,
                xerr=None if err is None else err[serie.index].values, ecolor=PALETA["oscuro"], capsize=3)
        margen = np.zeros(len(serie)) if err is None else err[serie.index].values
        for i, v in enumerate(serie.values):
            ax.text(v + margen[i] + serie.max() * 0.02, i, f"{v:.3f}", va="center", fontsize=9)
        ax.set_xlim(0, serie.max() * 1.2)
        ax.set_xlabel(xlabel)
        ax.set_title(ttl)
        ax.grid(axis="y", visible=False)
        guardar(fig, f"{FIG}/{archivo}")

    metricas["importancia"] = {
        "gini": gini.sort_values(ascending=False).to_dict(),
        "permutacion": perm_s.sort_values(ascending=False).to_dict(),
        "permutacion_std": perm_std.to_dict(),
        "top3_gini": top3_gini, "top3_permutacion": top3_perm,
        "top3_coincide": coincide, "top3_mismo_orden": top3_gini == top3_perm,
    }

    with open(f"{RES}/metricas.json", "w", encoding="utf-8") as f:
        json.dump(metricas, f, indent=2, ensure_ascii=False, default=float)
    titulo("EXPERIMENTOS COMPLETADOS")
    print(tabla.round(4).to_string(index=False))
    print(f"\nMétricas -> {RES}/metricas.json | Figuras -> {FIG}/")


if __name__ == "__main__":
    main()
