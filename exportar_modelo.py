import os
import json
from datetime import datetime

import pandas as pd
import numpy as np
import joblib
import sklearn
from sklearn.ensemble import RandomForestClassifier
from src.database.connection2 import MySQLConnector
from src.preprocessing.inferencia import limpiar_tipos, calcular_medianas, imputar, codificar

RUTA_PARAMETROS = 'reports/resultados/mejores_parametros.json'
RUTA_METRICAS = 'reports/resultados/metricas.json'


def cargar_hiperparametros():
    if os.path.exists(RUTA_PARAMETROS):
        with open(RUTA_PARAMETROS, encoding='utf-8') as f:
            params = json.load(f)['mejores_parametros']
        print(f"Hiperparámetros tomados de GridSearchCV ({RUTA_PARAMETROS}): {params}")
        return params
    print("No se encontró mejores_parametros.json: se usan los hiperparámetros del docente.")
    return {'n_estimators': 100, 'max_depth': 4}


def main():
    print("=================================================================")
    print("  FASE 1: ENTRENAMIENTO Y PERSISTENCIA DEL MODELO (DEPLOYMENT)")
    print("=================================================================\n")
    db = MySQLConnector()
    query = """
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
    df = db.extract_data(query)
    print(f"Registros extraídos de MySQL: {len(df):,}")

    df = limpiar_tipos(df)
    df['valor_mercado'] = pd.to_numeric(df['valor_mercado'], errors='coerce')
    medianas = calcular_medianas(df)
    df = imputar(df, **medianas)
    umbral_elite = df['valor_mercado'].quantile(0.75)
    df['es_elite'] = (df['valor_mercado'] >= umbral_elite).astype(int)
    X = codificar(df)
    y = df['es_elite']
    print(f"Umbral Élite (Q3): € {umbral_elite:,.0f} | Élite = {y.mean():.2%}")

    hiperparametros = cargar_hiperparametros()
    rf_model = RandomForestClassifier(
        **hiperparametros, class_weight='balanced', random_state=42, n_jobs=-1
    )
    rf_model.fit(X, y)

    metricas_test = {}
    if os.path.exists(RUTA_METRICAS):
        with open(RUTA_METRICAS, encoding='utf-8') as f:
            m = json.load(f)['modelos']['Random Forest optimizado']
        metricas_test = {k: m[k] for k in ('accuracy_train', 'accuracy_test', 'precision_elite',
                                          'recall_elite', 'f1_elite', 'auc_test')}

    paquete_despliegue = {
        'modelo': rf_model,
        'columnas_X': X.columns.tolist(),
        'medianas_estatura': medianas['medianas_estatura'],
        'mediana_global_estatura': medianas['mediana_global_estatura'],
        'mediana_global_edad': medianas['mediana_global_edad'],
        'umbral_elite_eur': float(umbral_elite),
        'posiciones_validas': sorted(df['posicion'].unique().tolist()),
        'hiperparametros': {**hiperparametros, 'class_weight': 'balanced', 'random_state': 42},
        'metricas_test': metricas_test,
        'fecha_entrenamiento': datetime.now().isoformat(timespec='seconds'),
        'version_sklearn': sklearn.__version__,
        'n_registros': int(len(df)),
    }
    os.makedirs('models', exist_ok=True)
    ruta_modelo = 'models/modelo_scouting_rf.joblib'
    joblib.dump(paquete_despliegue, ruta_modelo)

    assert os.path.exists(ruta_modelo)
    recargado = joblib.load(ruta_modelo)
    assert np.array_equal(recargado['modelo'].predict(X.head(50)), rf_model.predict(X.head(50)))
    print(f"Columnas del modelo: {paquete_despliegue['columnas_X']}")
    print(f"Paquete de modelo exportado exitosamente en: '{ruta_modelo}' "
          f"({os.path.getsize(ruta_modelo) / 1024:,.0f} KB) y verificado con joblib.load")


if __name__ == "__main__":
    main()
