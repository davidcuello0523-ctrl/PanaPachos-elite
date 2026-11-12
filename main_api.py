import logging
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Literal

import joblib
import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, ConfigDict, Field

from src.preprocessing.inferencia import preparar_para_modelo

RUTA_MODELO = Path(__file__).parent / "models" / "modelo_scouting_rf.joblib"
log = logging.getLogger("uvicorn.error")
estado = {"paquete": None, "error": None}


@asynccontextmanager
async def lifespan(app: FastAPI):
    if RUTA_MODELO.exists():
        estado["paquete"] = joblib.load(RUTA_MODELO)
        log.info("Modelo cargado desde %s", RUTA_MODELO)
    else:
        estado["error"] = (f"No se encontró el modelo en '{RUTA_MODELO}'. "
                           "Ejecute primero: python exportar_modelo.py")
        log.error(estado["error"])
    yield
    estado["paquete"] = None


app = FastAPI(
    title="API Scouting Élite",
    description="Predice si un jugador es Élite (top 25 % de valor de mercado) con un Random Forest entrenado con datos de Transfermarkt.",
    version="1.0.0",
    lifespan=lifespan,
)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

Posicion = Literal["Attack", "Defender", "Goalkeeper", "Midfield", "Missing"]


class JugadorEntrada(BaseModel):
    jugador: str = Field("Jugador sin nombre", max_length=100, description="Nombre del jugador")
    posicion: Posicion = Field(..., description="Posición en el campo (valores del dataset Transfermarkt)")
    edad: int = Field(..., ge=15, le=45, description="Edad en años (15–45)")
    estatura_cm: float | None = Field(
        None, ge=150, le=215,
        description="Estatura en cm (150–215). Si es null se imputa con la mediana de su posición")

    model_config = ConfigDict(json_schema_extra={"examples": [
        {"jugador": "Candidato A", "posicion": "Defender", "edad": 22, "estatura_cm": 188}
    ]})


class PrediccionSalida(BaseModel):
    jugador: str
    prediccion: Literal["Élite", "Estándar"]
    clase: Literal[0, 1]
    probabilidad_elite: float = Field(..., description="Probabilidad de ser Élite (%)")
    probabilidad_estandar: float = Field(..., description="Probabilidad de ser Estándar (%)")
    estatura_utilizada: float
    estatura_imputada: bool
    umbral_elite_eur: float


def obtener_paquete() -> dict:
    if estado["paquete"] is None:
        raise HTTPException(status_code=503, detail=estado["error"] or "Modelo no disponible")
    return estado["paquete"]


def predecir(jugadores: list[JugadorEntrada]) -> list[PrediccionSalida]:
    paquete = obtener_paquete()
    X, df = preparar_para_modelo(pd.DataFrame([j.model_dump() for j in jugadores]), paquete)
    probas = paquete["modelo"].predict_proba(X)[:, 1]
    clases = paquete["modelo"].predict(X)
    return [
        PrediccionSalida(
            jugador=fila.jugador,
            prediccion="Élite" if c == 1 else "Estándar",
            clase=int(c),
            probabilidad_elite=round(float(p) * 100, 2),
            probabilidad_estandar=round((1 - float(p)) * 100, 2),
            estatura_utilizada=round(float(fila.estatura_cm), 1),
            estatura_imputada=bool(fila.estatura_imputada),
            umbral_elite_eur=float(paquete["umbral_elite_eur"]),
        )
        for fila, p, c in zip(df.itertuples(), probas, clases)
    ]


@app.get("/", tags=["Estado"], summary="Health check de la API y del modelo")
def health():
    paquete = estado["paquete"]
    return {
        "estado": "en línea",
        "modelo_cargado": paquete is not None,
        "modelo": type(paquete["modelo"]).__name__ if paquete else None,
        "error": estado["error"],
        "documentacion": "/docs",
    }


@app.get("/info", tags=["Estado"], summary="Metadatos del modelo desplegado")
def info():
    p = obtener_paquete()
    return {
        "algoritmo": type(p["modelo"]).__name__,
        "hiperparametros": p.get("hiperparametros"),
        "variables": p["columnas_X"],
        "posiciones_validas": p.get("posiciones_validas"),
        "umbral_elite_eur": p["umbral_elite_eur"],
        "medianas_estatura": p["medianas_estatura"],
        "metricas_test": p.get("metricas_test"),
        "fecha_entrenamiento": p.get("fecha_entrenamiento"),
        "version_sklearn": p.get("version_sklearn"),
        "n_registros": p.get("n_registros"),
    }


@app.post("/predict", response_model=PrediccionSalida, tags=["Predicción"],
          summary="Clasifica un jugador como Élite o Estándar")
def predict(jugador: JugadorEntrada):
    return predecir([jugador])[0]


@app.post("/predict/lote", response_model=list[PrediccionSalida], tags=["Predicción"],
          summary="Clasifica una lista de jugadores")
def predict_lote(jugadores: list[JugadorEntrada]):
    if not jugadores:
        raise HTTPException(status_code=422, detail="La lista de jugadores está vacía")
    if len(jugadores) > 1000:
        raise HTTPException(status_code=422, detail="Máximo 1000 jugadores por lote")
    return predecir(jugadores)
