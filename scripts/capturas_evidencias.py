import json
import os
import subprocess
import sys
import time
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import requests
from playwright.sync_api import sync_playwright

RAIZ = Path(__file__).resolve().parents[1]
CAP = RAIZ / "reports" / "capturas"
API = "http://127.0.0.1:8000"
LOG = CAP / "uvicorn_log.txt"

EJEMPLOS = ["Candidato A", "Candidato B", "Candidato C", "Mateo Rincón", "Julián Ortega"]


def esperar_api(timeout=40):
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            if requests.get(API + "/", timeout=1).json().get("modelo_cargado"):
                return
        except requests.RequestException:
            pass
        time.sleep(0.5)
    raise RuntimeError("La API no respondió a tiempo")


def log_a_imagen():
    lineas = LOG.read_text(encoding="utf-8", errors="replace").splitlines()
    lineas = [l for l in lineas if l.strip()]
    if len(lineas) > 26:
        lineas = lineas[:5] + ["      ..."] + lineas[-20:]
    fig = plt.figure(figsize=(12, 0.27 * len(lineas) + 0.9))
    fig.patch.set_facecolor("#0C1821")
    fig.text(0.015, 0.985, "PS C:\\...\\Electiva II> uvicorn main_api:app --port 8000", color="#8BE9FD",
             family="monospace", fontsize=9, va="top")
    for i, l in enumerate(lineas):
        color = "#50FA7B" if "200 OK" in l else "#FFB86C" if "422" in l else "#E6EDF3"
        fig.text(0.015, 0.985 - (i + 1.4) / (len(lineas) + 2.2), l[:150], color=color,
                 family="monospace", fontsize=8.6, va="top")
    fig.savefig(CAP / "api_consola_uvicorn.png", dpi=150, facecolor=fig.get_facecolor())
    plt.close(fig)


def main():
    CAP.mkdir(parents=True, exist_ok=True)
    with open(LOG, "w", encoding="utf-8") as flog:
        servidor = subprocess.Popen(
            [sys.executable, "-m", "uvicorn", "main_api:app", "--port", "8000"],
            cwd=RAIZ, stdout=flog, stderr=subprocess.STDOUT, env={**os.environ, "PYTHONIOENCODING": "utf-8"})
        try:
            esperar_api()
            ejemplo = {"jugador": "Candidato B", "posicion": "Goalkeeper", "edad": 29, "estatura_cm": None}
            (CAP / "ejemplo_predict.json").write_text(json.dumps(
                {"request": ejemplo, "response": requests.post(API + "/predict", json=ejemplo).json()},
                indent=2, ensure_ascii=False), encoding="utf-8")
            with sync_playwright() as p:
                nav = p.chromium.launch()
                ctx = nav.new_context(viewport={"width": 1280, "height": 900}, device_scale_factor=1.5, locale="es-CO")
                pg = ctx.new_page()

                pg.goto(API + "/docs")
                pg.wait_for_selector(".opblock-summary")
                pg.wait_for_timeout(800)
                pg.screenshot(path=CAP / "api_swagger_docs.png", full_page=True)

                pg.set_viewport_size({"width": 1280, "height": 5000})

                def ejecutar(op_id, archivo, cuerpo=None):
                    op = pg.locator(f"#{op_id}")
                    op.locator(".opblock-summary").click()
                    op.locator("button.try-out__btn").click()
                    if cuerpo:
                        caja = op.locator("textarea.body-param__text")
                        caja.fill(cuerpo)
                        caja.evaluate("el => { el.style.height = '150px'; el.style.minHeight = '150px'; }")
                    op.locator("button.execute").click()
                    op.locator(".live-responses-table .response-col_status").first.wait_for()
                    pg.wait_for_timeout(700)
                    ini = op.bounding_box()
                    fin = op.locator(".live-responses-table").bounding_box()
                    pg.screenshot(path=CAP / archivo, clip={
                        "x": ini["x"], "y": ini["y"], "width": ini["width"],
                        "height": fin["y"] + fin["height"] - ini["y"] + 12})
                    op.locator(".opblock-summary").click()

                ejecutar("operations-Predicción-predict_predict_post", "api_predict_respuesta.png",
                         '{\n  "jugador": "Mateo Rincón",\n  "posicion": "Attack",\n  "edad": 23,\n  "estatura_cm": 184\n}')
                ejecutar("operations-Estado-info_info_get", "api_health.png")
                pg.set_viewport_size({"width": 1280, "height": 900})

                web = ctx.new_page()
                web.goto((RAIZ / "index.html").as_uri())
                web.wait_for_selector(".estado-api.online")
                web.screenshot(path=CAP / "web_formulario.png", full_page=True)

                def evaluar(nombre):
                    web.locator(".chip", has_text=nombre).click()
                    web.locator("#btnEnviar").click()
                    web.wait_for_function(f"document.getElementById('rNombre').textContent === {nombre!r}")
                    web.wait_for_timeout(1300)
                    return web.locator("#rBadge").text_content()

                resultados = {n: evaluar(n) for n in EJEMPLOS}
                print("Resultados web:", resultados)
                elite = next((n for n in reversed(EJEMPLOS) if resultados[n] == "Élite"), None)
                estandar = next((n for n in reversed(EJEMPLOS) if resultados[n] == "Estándar"), None)
                web.screenshot(path=CAP / "web_historial.png", full_page=True)
                for nombre, archivo in [(elite, "web_resultado_elite.png"), (estandar, "web_resultado_estandar.png")]:
                    if nombre is None:
                        print("AVISO: ningún ejemplo produjo la categoría para", archivo)
                        continue
                    web.reload()
                    web.wait_for_selector(".estado-api.online")
                    evaluar(nombre)
                    a = web.locator("main > section").nth(0).bounding_box()
                    b = web.locator("main > section").nth(1).bounding_box()
                    web.screenshot(path=CAP / archivo, full_page=True, clip={
                        "x": a["x"] - 8, "y": a["y"] - 8, "width": b["x"] + b["width"] - a["x"] + 16,
                        "height": max(a["height"], b["height"]) + 16})

                web.evaluate("""() => fetch('http://127.0.0.1:8000/predict', {method:'POST',
                    headers:{'Content-Type':'application/json'}, body: JSON.stringify({posicion:'Attack', edad:50})})""")
                nav.close()
            time.sleep(0.5)
        finally:
            servidor.terminate()
            servidor.wait(timeout=10)
    log_a_imagen()
    print("Capturas guardadas en", CAP)


if __name__ == "__main__":
    main()
