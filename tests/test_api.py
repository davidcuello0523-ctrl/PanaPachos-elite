import sys
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from main_api import app  # noqa: E402


@pytest.fixture(scope="module")
def client():
    with TestClient(app) as c:
        yield c


def test_health(client):
    r = client.get("/")
    assert r.status_code == 200
    assert r.json()["modelo_cargado"] is True


def test_info(client):
    r = client.get("/info")
    assert r.status_code == 200
    assert "posicion_Goalkeeper" in r.json()["variables"]


def test_prediccion_normal(client):
    r = client.post("/predict", json={"jugador": "Candidato A", "posicion": "Defender",
                                      "edad": 22, "estatura_cm": 188})
    assert r.status_code == 200
    d = r.json()
    assert d["prediccion"] in ("Élite", "Estándar")
    assert d["clase"] == (1 if d["prediccion"] == "Élite" else 0)
    assert 0 <= d["probabilidad_elite"] <= 100
    assert d["probabilidad_elite"] + d["probabilidad_estandar"] == pytest.approx(100, abs=0.02)
    assert d["estatura_imputada"] is False and d["estatura_utilizada"] == 188


def test_estatura_nula_se_imputa(client):
    r = client.post("/predict", json={"jugador": "Candidato B", "posicion": "Goalkeeper",
                                      "edad": 29, "estatura_cm": None})
    assert r.status_code == 200
    d = r.json()
    assert d["estatura_imputada"] is True
    assert d["estatura_utilizada"] > 185


def test_posicion_invalida(client):
    r = client.post("/predict", json={"posicion": "Delantero centro", "edad": 25, "estatura_cm": 180})
    assert r.status_code == 422


@pytest.mark.parametrize("edad", [10, 50])
def test_edad_fuera_de_rango(client, edad):
    r = client.post("/predict", json={"posicion": "Attack", "edad": edad, "estatura_cm": 180})
    assert r.status_code == 422


def test_lote(client):
    r = client.post("/predict/lote", json=[
        {"posicion": "Attack", "edad": 24, "estatura_cm": 182},
        {"posicion": "Midfield", "edad": 19, "estatura_cm": None},
    ])
    assert r.status_code == 200 and len(r.json()) == 2
    assert r.json()[0]["jugador"] == "Jugador sin nombre"
