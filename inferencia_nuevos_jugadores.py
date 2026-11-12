import joblib
import pandas as pd
import numpy as np
from src.preprocessing.inferencia import preparar_para_modelo


def preprocesar_y_predecir(nuevos_jugadores_df, ruta_paquete='models/modelo_scouting_rf.joblib'):
    paquete = joblib.load(ruta_paquete)
    modelo = paquete['modelo']
    X_nuevo, df = preparar_para_modelo(nuevos_jugadores_df, paquete)
    predicciones = modelo.predict(X_nuevo)
    probabilidades = modelo.predict_proba(X_nuevo)[:, 1]
    df['Prediccion_Clase'] = np.where(predicciones == 1, 'Élite', 'Estándar')
    df['Probabilidad_Elite'] = np.round(probabilidades * 100, 2)
    return df


def main():
    print("=================================================================")
    print("  FASE 2: SIMULACIÓN DE INFERENCIA EN PRODUCCIÓN (NUEVOS DATOS)")
    print("=================================================================\n")
    nuevos_candidatos = pd.DataFrame([
        {'jugador': 'Candidato A', 'posicion': 'Defender', 'estatura_cm': 188, 'edad': 22},
        {'jugador': 'Candidato B', 'posicion': 'Goalkeeper', 'estatura_cm': None, 'edad': 29},
        {'jugador': 'Candidato C', 'posicion': 'Midfield', 'estatura_cm': 172, 'edad': 19}
    ])
    print("Entrada de datos crudos:")
    print(nuevos_candidatos)
    print("\nProcesando inferencia con el modelo serializado...")
    resultados = preprocesar_y_predecir(nuevos_candidatos)
    print(resultados[['jugador', 'posicion', 'edad', 'estatura_cm', 'estatura_imputada',
                      'Prediccion_Clase', 'Probabilidad_Elite']].to_string(index=False))
    assert resultados['estatura_cm'].notna().all() and resultados.loc[1, 'estatura_imputada']


if __name__ == "__main__":
    main()
