# PanaPacho's Elite

Modelo que predice si un jugador es **Élite** (top 25 % de valor de mercado en Transfermarkt) usando su posición, edad y estatura. Random Forest + API con FastAPI + una página web para probarlo.

Trabajo final de Electiva II (Minería de Datos) – David Cuello.

```
MySQL → exportar_modelo.py → models/modelo_scouting_rf.joblib → main_api.py (POST /predict) ⇄ index.html
```

## Archivos

| Archivo | Qué hace |
|---|---|
| `src/database/connection2.py` | Saca los datos de MySQL a un DataFrame |
| `src/preprocessing/inferencia.py` | Imputación por mediana de la posición + One-Hot |
| `experimentos_modelado.py` | EDA, overfitting, GridSearchCV, ROC, importancia de variables |
| `exportar_modelo.py` | Entrena el modelo final y lo guarda con joblib |
| `inferencia_nuevos_jugadores.py` | Prueba el modelo con los candidatos A, B y C |
| `main_api.py` | API (`GET /`, `GET /info`, `POST /predict`, `POST /predict/lote`) |
| `index.html` | Página para usar la API |
| `tests/test_api.py` | Pruebas de la API |
| `scripts/` | Diagramas, capturas e informe |
| `reports/` | Figuras, capturas y métricas |

## Cómo correrlo

Necesitas Python 3.11+ y, para entrenar, MySQL (XAMPP) con la base `mineria_datos` y la tabla `players` del dataset de Transfermarkt.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
```

```powershell
python experimentos_modelado.py
python exportar_modelo.py
python inferencia_nuevos_jugadores.py
python -m pytest tests -v
uvicorn main_api:app --reload --port 8000
start index.html
```

El `.joblib` ya está en el repo, así que para probar basta con levantar la API y abrir `index.html` (no hace falta MySQL). La documentación queda en http://127.0.0.1:8000/docs.

## Ejemplo

```bash
curl -X POST http://127.0.0.1:8000/predict \
  -H "Content-Type: application/json" \
  -d '{"jugador": "Candidato A", "posicion": "Defender", "edad": 22, "estatura_cm": 188}'
```

- `posicion`: `Attack`, `Defender`, `Goalkeeper`, `Midfield` o `Missing`
- `edad`: 15 a 45
- `estatura_cm`: 150 a 215, o `null` (se rellena con la mediana de la posición)

Si los datos no son válidos responde 422.

```json
{
  "jugador": "Candidato B",
  "prediccion": "Estándar",
  "clase": 0,
  "probabilidad_elite": 33.88,
  "probabilidad_estandar": 66.12,
  "estatura_utilizada": 189.0,
  "estatura_imputada": true,
  "umbral_elite_eur": 700000.0
}
```

## Resultados (test)

| Modelo | Acc. Train | Acc. Test | Recall Élite | AUC |
|---|---|---|---|---|
| Árbol sin restricciones | 0.7575 | 0.7157 | 0.1318 | 0.6523 |
| Árbol podado | 0.7330 | 0.7331 | 0.0000 | 0.6785 |
| Random Forest (`class_weight='balanced'`) | 0.6197 | 0.6054 | 0.7648 | 0.6937 |

El árbol podado tiene más accuracy pero no detecta ningún Élite, por eso se eligió el Random Forest. Todo el detalle está en `reports/resultados/metricas.json` y en el informe PDF.
