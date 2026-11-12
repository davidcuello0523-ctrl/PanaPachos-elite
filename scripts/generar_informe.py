import json
import re
import sys
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
from src.estilo import ETIQUETAS_POSICION, etiqueta_variable  # noqa: E402

FIG, CAP, RES = RAIZ / "reports/figuras", RAIZ / "reports/capturas", RAIZ / "reports/resultados"
SALIDA = RAIZ / "Informe_Final_Scouting_David_Cuello.docx"
M = json.loads((RES / "metricas.json").read_text(encoding="utf-8"))
D, MOD, GS, IMP = M["datos"], M["modelos"], M["gridsearch"], M["importancia"]
PRIMARIO = RGBColor(0x1A, 0x1A, 0x1A)
FUENTE = "Calibri"

TITULO = ("Ecosistema Analítico End-to-End para Scouting Deportivo: Clasificación de Jugadores Élite "
          "con Random Forest y Despliegue mediante API REST")


def miles(n):
    return f"{n:,.0f}".replace(",", " ")


def pct(x, d=2):
    return f"{x * 100:.{d}f} %"


def f4(x):
    return f"{x:.4f}"


def set_cell_bg(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)


def campo(parrafo, instr):
    r = parrafo.add_run()
    for tipo, texto in (("begin", None), (None, instr), ("separate", None), (None, "1"), ("end", None)):
        if tipo:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tipo)
            r._r.append(el)
        elif texto == instr:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = instr
            r._r.append(it)
        else:
            t = OxmlElement("w:t")
            t.text = texto
            r._r.append(t)


def fuente_run(run, nombre=FUENTE):
    run.font.name = nombre
    run._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), nombre)
    run._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:cs"), nombre)


TOKENS = re.compile(r"(\*\*.+?\*\*|`.+?`|\*[^*]+?\*|_\{.+?\})")


def texto_rico(par, texto, size=None, negrita=False, cursiva=False):
    for trozo in TOKENS.split(texto):
        if not trozo:
            continue
        if trozo.startswith("**"):
            texto_rico(par, trozo[2:-2], size, True, cursiva)
            continue
        if trozo.startswith("*"):
            texto_rico(par, trozo[1:-1], size, negrita, True)
            continue
        if trozo.startswith("`"):
            r = par.add_run(trozo[1:-1])
            fuente_run(r, "Consolas")
            r.font.size = Pt(10.5 if size is None else size - 0.5)
            r.font.color.rgb = RGBColor(0x26, 0x26, 0x26)
        else:
            sub = trozo.startswith("_{")
            r = par.add_run(trozo[2:-1] if sub else trozo)
            r.font.subscript = sub
            if size:
                r.font.size = Pt(size)
        r.bold = negrita or None
        r.italic = cursiva or None
    return par


class Informe:
    def __init__(self):
        self.doc = Document()
        self.nfig = 0
        self.ntab = 0
        self._estilos()

    def _estilos(self):
        doc = self.doc
        sec = doc.sections[0]
        sec.page_height, sec.page_width = Cm(27.94), Cm(21.59)
        for lado in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
            setattr(sec, lado, Cm(2.5))
        normal = doc.styles["Normal"]
        normal.font.name = FUENTE
        normal.font.size = Pt(12)
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), FUENTE)
        pf = normal.paragraph_format
        pf.line_spacing = 1.5
        pf.space_after = Pt(6)
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for nombre, tam, antes in (("Heading 1", 14, 18), ("Heading 2", 12.5, 12), ("Heading 3", 12, 10)):
            st = doc.styles[nombre]
            st.font.name = FUENTE
            st.font.size = Pt(tam)
            st.font.bold = True
            st.font.italic = False
            st.font.color.rgb = PRIMARIO
            rf = st.element.rPr.rFonts
            for att in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                rf.set(qn(att), FUENTE)
            st.paragraph_format.space_before = Pt(antes)
            st.paragraph_format.space_after = Pt(6)
            st.paragraph_format.line_spacing = 1.15
            st.paragraph_format.keep_with_next = True
            st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        lb = doc.styles["List Bullet"]
        lb.font.name = FUENTE
        lb.font.size = Pt(12)
        lb.paragraph_format.line_spacing = 1.5
        lb.paragraph_format.space_after = Pt(3)
        lb.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    def h1(self, t, nueva_pagina=True):
        h = self.doc.add_heading(t, level=1)
        h.paragraph_format.page_break_before = nueva_pagina
        return h

    def h2(self, t):
        return self.doc.add_heading(t, level=2)

    def h3(self, t):
        return self.doc.add_heading(t, level=3)

    def p(self, t, **kw):
        return texto_rico(self.doc.add_paragraph(), t, **kw)

    def bullets(self, items):
        for it in items:
            texto_rico(self.doc.add_paragraph(style="List Bullet"), it)

    def salto(self):
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    def formula(self, expr):
        par = self.doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_before = Pt(4)
        par.paragraph_format.space_after = Pt(8)
        for trozo in re.split(r"(_\{.+?\}|\^\{.+?\})", expr):
            if not trozo:
                continue
            r = par.add_run(trozo[2:-1] if trozo[:2] in ("_{", "^{") else trozo)
            fuente_run(r, "Cambria Math")
            r.italic = True
            r.font.size = Pt(12)
            r.font.subscript = trozo.startswith("_{")
            r.font.superscript = trozo.startswith("^{")

    def codigo(self, texto, tam=9):
        tabla = self.doc.add_table(rows=1, cols=1)
        tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
        celda = tabla.rows[0].cells[0]
        set_cell_bg(celda, "F2F2F2")
        par = celda.paragraphs[0]
        par.alignment = WD_ALIGN_PARAGRAPH.LEFT
        par.paragraph_format.line_spacing = 1.0
        par.paragraph_format.space_after = Pt(0)
        for i, linea in enumerate(texto.strip("\n").split("\n")):
            if i:
                par.add_run().add_break()
            r = par.add_run(linea)
            fuente_run(r, "Consolas")
            r.font.size = Pt(tam)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(0)

    def _rotulo(self, tipo, n, titulo):
        a = self.doc.add_paragraph()
        a.alignment = WD_ALIGN_PARAGRAPH.LEFT
        a.paragraph_format.space_before = Pt(10)
        a.paragraph_format.space_after = Pt(0)
        a.paragraph_format.keep_with_next = True
        a.paragraph_format.line_spacing = 1.15
        a.add_run(f"{tipo} {n}").bold = True
        b = self.doc.add_paragraph()
        b.alignment = WD_ALIGN_PARAGRAPH.LEFT
        b.paragraph_format.keep_with_next = True
        b.paragraph_format.space_after = Pt(4)
        b.paragraph_format.line_spacing = 1.15
        b.add_run(titulo).italic = True

    def _nota(self, nota):
        par = self.doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.LEFT
        par.paragraph_format.line_spacing = 1.0
        par.paragraph_format.space_after = Pt(10)
        r = par.add_run("Nota. ")
        r.italic = True
        r.font.size = Pt(10)
        texto_rico(par, nota, size=10)

    def figura(self, ruta, titulo, nota="Elaboración propia.", ancho=15.5, alto_max=16.5):
        self.nfig += 1
        self._rotulo("Figura", self.nfig, titulo)
        w, h = Image.open(ruta).size
        ancho = min(ancho, alto_max * w / h)
        par = self.doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.keep_with_next = True
        par.paragraph_format.line_spacing = 1.0
        par.add_run().add_picture(str(ruta), width=Cm(ancho))
        self._nota(nota)
        return self.nfig

    def tabla(self, encabezados, filas, titulo, nota="Elaboración propia.", anchos=None, tam=10):
        self.ntab += 1
        self._rotulo("Tabla", self.ntab, titulo)
        t = self.doc.add_table(rows=1, cols=len(encabezados))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, h in enumerate(encabezados):
            c = t.rows[0].cells[i]
            set_cell_bg(c, "404040")
            par = c.paragraphs[0]
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = par.add_run(h)
            r.bold = True
            r.font.size = Pt(tam)
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        for k, fila in enumerate(filas):
            celdas = t.add_row().cells
            for i, v in enumerate(fila):
                par = celdas[i].paragraphs[0]
                par.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 or len(str(v)) > 30 else WD_ALIGN_PARAGRAPH.CENTER
                texto_rico(par, str(v), size=tam)
                if k % 2:
                    set_cell_bg(celdas[i], "F2F2F2")
        for fila in t.rows:
            for c in fila.cells:
                for par in c.paragraphs:
                    par.paragraph_format.line_spacing = 1.0
                    par.paragraph_format.space_after = Pt(2)
                    par.paragraph_format.space_before = Pt(2)
            trPr = fila._tr.get_or_add_trPr()
            cs = OxmlElement("w:cantSplit")
            trPr.append(cs)
        if anchos:
            for fila in t.rows:
                for c, w in zip(fila.cells, anchos):
                    c.width = Cm(w)
        self._nota(nota)
        return self.ntab

    def guardar(self):
        self.doc.save(SALIDA)


def portada(inf: Informe):
    doc = inf.doc
    sec = doc.sections[0]
    sec.different_first_page_header_footer = True

    def linea(texto, tam=12, negrita=False, antes=0, despues=0, color=None, cursiva=False):
        par = doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_before = Pt(antes)
        par.paragraph_format.space_after = Pt(despues)
        par.paragraph_format.line_spacing = 1.15
        r = par.add_run(texto)
        r.font.size = Pt(tam)
        r.bold = negrita
        r.italic = cursiva
        if color:
            r.font.color.rgb = color
        return par

    logo = doc.add_paragraph()
    logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    logo.add_run().add_picture(str(RAIZ / "reports/logo_ucc.png"), width=Cm(3.6))
    linea("UNIVERSIDAD COOPERATIVA DE COLOMBIA", 15, True, antes=6)
    linea("Sede Montería", 12.5)
    linea("Facultad de Ingeniería", 12.5, antes=10)
    linea("Programa de Ingeniería de Sistemas", 12.5)
    linea("Electiva II – Minería de Datos", 12.5, antes=10, despues=40)
    linea(TITULO, 17, True, color=PRIMARIO, despues=12)
    linea("Informe ejecutivo del proyecto final – Momento 3", 12.5, cursiva=True, despues=60)
    linea("Presentado por:", 12, antes=0)
    linea("David Cuello", 13, True, despues=22)
    linea("Docente:", 12)
    linea("Magíster Royer Mejía Villalba", 13, True, despues=70)
    linea("Montería, Córdoba", 12)
    linea("Noviembre de 2026", 12)

    pie = sec.footer.paragraphs[0]
    pie.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    campo(pie, "PAGE")
    for r in pie.runs:
        r.font.size = Pt(10)


def tabla_contenido(inf: Informe):
    par = inf.doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.LEFT
    par.paragraph_format.page_break_before = True
    r = par.add_run("Tabla de contenido")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = PRIMARIO
    campo(inf.doc.add_paragraph(), 'TOC \\o "1-2" \\h \\z \\u')


def introduccion(inf: Informe):
    inf.h1("1. Introducción y objetivos")
    inf.p("En el fútbol profesional, el departamento de *scouting* (ojeo) debe filtrar miles de perfiles "
          "para encontrar a los pocos jugadores que justifican una inversión. Cada visoría presencial cuesta "
          "viajes, tiempo de analistas y oportunidades perdidas, y un fichaje equivocado compromete presupuestos "
          "de varios millones de euros. Los clubes con recursos limitados necesitan, por tanto, herramientas que "
          "prioricen candidatos de forma objetiva antes de destinar recursos humanos a observarlos.")
    inf.p(f"Este proyecto construye un **ecosistema analítico de extremo a extremo** sobre el conjunto de datos "
          f"de Transfermarkt almacenado en MySQL (base `mineria_datos`, tabla `players`, {miles(50149)} "
          f"jugadores). El sistema extrae los datos con un conector propio, los depura, entrena y optimiza un "
          f"clasificador *Random Forest* que estima si un jugador pertenece al **25 % más valioso del mercado "
          f"(clase Élite, valor ≥ € {miles(D['umbral_q3_eur'])})**, serializa el modelo con Joblib, lo expone "
          f"como una **API REST** (FastAPI) y lo pone en manos del usuario final mediante un **aplicativo web** "
          f"interactivo. Así, el modelo deja de vivir en un *notebook* y se convierte en un servicio que cualquier "
          f"scout puede consultar.")
    inf.h2("1.1. Objetivo general")
    inf.p("Desarrollar y desplegar un sistema de clasificación de jugadores de fútbol en las categorías Élite y "
          "Estándar a partir de variables físicas y demográficas (edad, estatura y posición), siguiendo la "
          "metodología CRISP-DM, desde la extracción de datos relacionales en MySQL hasta su consumo mediante una "
          "API REST y una interfaz web.")
    inf.h2("1.2. Objetivos específicos")
    inf.bullets([
        "Extraer los datos de jugadores desde MySQL con el conector `MySQLConnector` (`connection2.py`) y tratar "
        "los valores faltantes mediante imputación por mediana agrupada por posición.",
        "Construir un target binario por percentil (Top 25 %) y preparar las variables con One-Hot Encoding y "
        "partición estratificada.",
        "Diagnosticar el sobreajuste comparando un árbol de decisión sin restricciones contra uno con poda previa, "
        "y evidenciar la paradoja de la exactitud en un problema desbalanceado.",
        "Corregir el desbalance con `class_weight='balanced'` y optimizar un `RandomForestClassifier` con "
        "`GridSearchCV` y validación cruzada K-Fold estratificada.",
        "Evaluar los modelos con curvas ROC y AUC, e interpretar la importancia de variables (Gini y permutación).",
        "Serializar el paquete del modelo con Joblib, desplegarlo en una API REST con FastAPI y Pydantic, y "
        "consumirlo desde un aplicativo web con formulario, `fetch` y barra de progreso.",
    ])
    inf.h2("1.3. Justificación bajo la metodología CRISP-DM")
    inf.p("CRISP-DM (*Cross-Industry Standard Process for Data Mining*; Chapman et al., 2000) es el estándar de facto "
          "para proyectos de minería de datos: organiza el trabajo en seis fases iterativas que conectan la "
          "necesidad del negocio con un modelo en producción. Se eligió porque el proyecto no termina en una "
          "métrica, sino en un servicio desplegado, y CRISP-DM es la única de las metodologías vistas en el curso "
          "que incluye explícitamente la fase de **Despliegue**. La Tabla 1 y la Figura 1 resumen cómo se cubrió cada fase.")
    inf.tabla(
        ["Fase CRISP-DM", "Aplicación en el proyecto", "Artefacto"],
        [
            ["1. Comprensión del negocio", "Problema: priorizar candidatos con alto valor de mercado para "
             "reducir costos de visoría. Éxito: detectar la mayoría de jugadores Élite (recall alto) con buena "
             "capacidad de discriminación (AUC).", "Este informe, §1"],
            ["2. Comprensión de los datos", f"Consulta SQL a `players`: {miles(D['registros'])} registros válidos, "
             f"{miles(D['nulos_antes']['estatura_cm'])} estaturas nulas, 5 posiciones, umbral Q3 de "
             f"€ {miles(D['umbral_q3_eur'])}.", "`connection2.py`, figuras 2 y 3"],
            ["3. Preparación de los datos", "Imputación por mediana agrupada, target binario P75, One-Hot "
             "Encoding y split estratificado 80/20 sin fuga de datos.", "`src/preprocessing/inferencia.py`"],
            ["4. Modelado", "Árbol libre, árbol podado, árbol balanceado y Random Forest optimizado con "
             f"GridSearchCV ({GS['n_combinaciones']} combinaciones × 5 folds).", "`experimentos_modelado.py`"],
            ["5. Evaluación", "Accuracy Train/Test, Precision, Recall, F1, matrices de confusión, ROC-AUC e "
             "importancia de variables.", "`reports/resultados/`"],
            ["6. Despliegue", "Paquete Joblib, API REST con FastAPI y aplicativo web.",
             "`exportar_modelo.py`, `main_api.py`, `index.html`"],
        ],
        "Correspondencia entre las fases de CRISP-DM y el trabajo realizado", anchos=[3.8, 8.7, 4.0])
    inf.figura(FIG / "09_crisp_dm.png", "Ciclo CRISP-DM aplicado al proyecto y scripts que cubren cada fase",
               "Elaboración propia con base en Chapman et al. (2000).", ancho=12.5)


CONCEPTOS = [
    ("Fase 1 – Ingesta, preprocesamiento e ingeniería de características", [
        ("Imputación por mediana agrupada",
         "Técnica de tratamiento de valores faltantes que reemplaza cada nulo por la mediana del subgrupo al que "
         "pertenece el registro (aquí, la posición), en lugar de un único valor global.",
         "x̂_{i} = mediana{ x_{j} : posición_{j} = posición_{i} }",
         "La mediana es robusta frente a valores atípicos (a diferencia de la media) y el agrupamiento respeta la "
         "estructura real de los datos: un portero no mide lo mismo que un mediocampista (Figura 3). Imputar con la "
         "mediana global sesgaría a los porteros hacia abajo y a los mediocampistas hacia arriba.",
         "En `src/preprocessing/inferencia.py`, `calcular_medianas()` aprende "
         "`df.groupby('posicion')['estatura_cm'].median()` y `imputar()` aplica "
         "`fillna(df['posicion'].map(medianas_estatura))`, con la mediana global como respaldo; la edad se imputa "
         "con la mediana global. Las medianas se guardan en el paquete `.joblib` (`medianas_estatura`, "
         "`mediana_global_estatura`, `mediana_global_edad`)."),
        ("Target binario por percentil (Top 25 %)",
         "Transformación de una variable continua (valor de mercado) en una etiqueta categórica binaria mediante un "
         "umbral estadístico: el tercer cuartil (Q3 o percentil 75).",
         "y_{i} = 1 si valor_{i} ≥ P_{75}(valor);   y_{i} = 0 en otro caso",
         "Convierte un problema de regresión con una distribución muy asimétrica (pocos jugadores valen decenas de "
         "millones) en una decisión de negocio clara: ¿vale la pena seguir a este jugador? El percentil hace que el "
         "umbral sea relativo al mercado y no una cifra arbitraria.",
         f"`umbral_elite = df['valor_mercado'].quantile(0.75)` (€ {miles(D['umbral_q3_eur'])}) y "
         "`df['es_elite'] = (df['valor_mercado'] >= umbral_elite).astype(int)` en `exportar_modelo.py` y "
         "`experimentos_modelado.py`. El umbral se guarda como `umbral_elite_eur`."),
        ("One-Hot Encoding (get_dummies)",
         "Conversión de una variable categórica nominal en un conjunto de columnas binarias (variables *dummy*), "
         "una por categoría, que valen 1 si el registro pertenece a esa categoría y 0 en otro caso.",
         "posición ∈ {A, D, G, M, X}  →  (x_{D}, x_{G}, x_{M}, x_{X}) ∈ {0,1}^{4}",
         "Los algoritmos de scikit-learn requieren entradas numéricas. Asignar números arbitrarios (Delantero = 1, "
         "Defensa = 2...) inventaría un orden inexistente. Con `drop_first=True` se elimina una categoría de "
         "referencia (Delantero) para evitar redundancia (colinealidad perfecta o *trampa de las dummies*).",
         "`pd.get_dummies(df[['posicion']], drop_first=True, dtype=int)` genera "
         f"{', '.join('`' + c + '`' for c in D['columnas_X'][2:])}. En inferencia se usa `get_dummies` sin "
         "`drop_first` seguido de `reindex(columns=columnas_X, fill_value=0)` para obtener exactamente las mismas "
         "columnas del entrenamiento (ver §2.4.2)."),
        ("Muestreo estratificado (stratify)",
         "Técnica de partición Train/Test que conserva en cada subconjunto la misma proporción de clases que "
         "existe en la población original.",
         "P(y=1 | Train) ≈ P(y=1 | Test) ≈ P(y=1)",
         "Con clases desbalanceadas, una partición aleatoria simple podría dejar al conjunto de prueba con "
         "muy pocos jugadores Élite y producir métricas inestables o engañosas.",
         f"`train_test_split(X, y, test_size=0.2, stratify=y, random_state=42)` en `experimentos_modelado.py`. "
         f"Resultado: Élite = {pct(D['proporcion_elite_train'], 3)} en Train y {pct(D['proporcion_elite_test'], 3)} "
         f"en Test. `StratifiedKFold` aplica el mismo principio dentro de la validación cruzada."),
    ]),
    ("Fase 2 – Modelado, diagnóstico y optimización", [
        ("Sobreajuste (overfitting)",
         "Fenómeno en el que un modelo memoriza el ruido y los detalles particulares del conjunto de "
         "entrenamiento, de modo que su desempeño en datos nuevos es sensiblemente inferior.",
         "Brecha = Accuracy_{Train} − Accuracy_{Test}",
         "Un modelo sobreajustado da una falsa sensación de calidad y falla en producción. La brecha Train − Test "
         "es su síntoma principal.",
         f"`DecisionTreeClassifier(random_state=42)` sin restricciones creció hasta profundidad "
         f"{MOD['Árbol sin restricciones']['profundidad']} con {miles(MOD['Árbol sin restricciones']['hojas'])} hojas "
         f"y mostró una brecha de {MOD['Árbol sin restricciones']['brecha']:+.4f} (§3.3)."),
        ("Poda previa (pre-pruning)",
         "Estrategia de regularización que detiene el crecimiento del árbol fijando límites estructurales antes de "
         "entrenarlo: profundidad máxima (`max_depth`) y número mínimo de muestras por hoja (`min_samples_leaf`).",
         None,
         "Impide que el árbol cree reglas para casos aislados; cada decisión debe estar respaldada por un número "
         "mínimo de jugadores, lo que mejora la generalización y la interpretabilidad.",
         f"`DecisionTreeClassifier(max_depth=4, min_samples_leaf=20)`: el árbol pasó de "
         f"{miles(MOD['Árbol sin restricciones']['hojas'])} a {MOD['Árbol podado']['hojas']} hojas y la brecha cayó a "
         f"prácticamente cero. En el Random Forest se exploraron `max_depth` y `min_samples_leaf` dentro de la malla."),
        ("Paradoja de la exactitud (accuracy paradox)",
         "Situación propia de datos desbalanceados en la que un modelo obtiene una exactitud alta simplemente "
         "prediciendo siempre la clase mayoritaria, sin ninguna capacidad real de detectar la clase minoritaria.",
         "Accuracy = (TP + TN) / (TP + TN + FP + FN)",
         "Demuestra que la exactitud no es una métrica adecuada cuando la clase de interés es minoritaria; deben "
         "usarse Recall, Precision, F1 y AUC.",
         f"`DummyClassifier(strategy='most_frequent')` obtuvo accuracy = {f4(M['dummy']['accuracy_test'])} con "
         f"Recall Élite = {M['dummy']['recall_elite']:.1f}. El árbol podado sin balanceo cayó en la misma trampa (§3.2)."),
        ("Penalización por pesos de costo (class_weight='balanced')",
         "Técnica algorítmica que modifica la función de pérdida para que equivocarse con la clase minoritaria "
         "cueste más que equivocarse con la mayoritaria. Cada clase c recibe un peso inversamente proporcional a "
         "su frecuencia:",
         "w_{c} = n / (k · n_{c})",
         f"Donde n es el número de muestras, k el número de clases y n_{{c}} las muestras de la clase c. Con "
         f"{pct(D['proporcion_elite_train'], 1)} de Élite, w_{{Élite}} ≈ "
         f"{1 / (2 * D['proporcion_elite_train']):.2f} y w_{{Estándar}} ≈ {1 / (2 * (1 - D['proporcion_elite_train'])):.2f}: "
         "un error sobre un jugador Élite pesa casi tres veces más. No crea datos sintéticos (a diferencia de SMOTE).",
         "`class_weight='balanced'` en el árbol podado balanceado y en `RandomForestClassifier` de "
         "`experimentos_modelado.py` y `exportar_modelo.py`."),
        ("Random Forest y ensamble (Bagging, Bootstrap y Feature Subspacing)",
         "Algoritmo de aprendizaje por ensamble (Breiman, 2001) que entrena muchos árboles de decisión "
         "independientes y combina sus votos. Cada árbol se entrena con una muestra *bootstrap* (muestreo con "
         "reemplazo del conjunto de entrenamiento) y, en cada división, solo considera un subconjunto aleatorio de "
         "variables (*feature subspacing*, parámetro `max_features`).",
         "p̂(Élite | x) = (1/B) · Σ_{b=1..B} p̂_{b}(Élite | x)",
         "*Bagging* (bootstrap aggregating) reduce la varianza: los errores de árboles poco correlacionados se "
         "compensan al promediar. El subespacio de variables descorrelaciona aún más los árboles. El resultado es "
         "un modelo más estable y con mejor generalización que un árbol único.",
         f"`RandomForestClassifier(class_weight='balanced', random_state=42, n_jobs=-1)` con los mejores "
         f"hiperparámetros: `n_estimators={GS['mejores_parametros']['n_estimators']}`, "
         f"`max_depth={GS['mejores_parametros']['max_depth']}`, "
         f"`min_samples_leaf={GS['mejores_parametros']['min_samples_leaf']}`, "
         f"`max_features='{GS['mejores_parametros']['max_features']}'`."),
        ("Búsqueda combinatoria con GridSearchCV y K-Fold Cross-Validation",
         "Exploración exhaustiva de todas las combinaciones de una malla de hiperparámetros, evaluando cada una "
         "con validación cruzada K-Fold: el entrenamiento se divide en K partes, se entrena con K − 1 y se valida "
         "con la restante, rotando K veces y promediando el resultado.",
         "Score_{CV} = (1/K) · Σ_{k=1..K} AUC_{k}",
         "Evita elegir hiperparámetros por ensayo y error o ajustándolos al conjunto de prueba (lo que sería otra "
         "forma de fuga de información). El promedio de K evaluaciones es una estimación más fiable que una sola "
         "partición.",
         f"`GridSearchCV(..., cv=StratifiedKFold(n_splits=5, shuffle=True, random_state=42), "
         f"scoring={{'roc_auc', 'f1'}}, refit='roc_auc')` sobre {GS['n_combinaciones']} combinaciones "
         f"({GS['ajustes']} entrenamientos, {GS['duracion_min']:.1f} minutos)."),
    ]),
    ("Fase 3 – Evaluación avanzada y despliegue", [
        ("Curva ROC y métrica AUC",
         "La curva ROC (*Receiver Operating Characteristic*) grafica la Tasa de Verdaderos Positivos frente a la "
         "Tasa de Falsos Positivos para todos los umbrales de decisión posibles. El AUC (*Area Under the Curve*) "
         "resume la curva en un número entre 0 y 1 (Fawcett, 2006).",
         "TPR = TP / (TP + FN)        FPR = FP / (FP + TN)",
         "A diferencia de la exactitud, no depende de un umbral fijo ni del desbalance. Su lectura probabilística "
         "es intuitiva: el AUC es la probabilidad de que el modelo asigne una puntuación más alta a un jugador Élite "
         "elegido al azar que a uno Estándar elegido al azar. AUC = 0.5 equivale al azar (diagonal) y AUC = 1.0 a "
         "una clasificación perfecta.",
         "`roc_curve(y_test, predict_proba(X_test)[:, 1])` y `roc_auc_score` en `experimentos_modelado.py`; "
         "también es la métrica `scoring='roc_auc'` que optimiza el GridSearchCV (Figura 6)."),
        ("Importancia de variables (Gini importance)",
         "Medida de la contribución de cada variable a la reducción de impureza de Gini, sumada sobre todas las "
         "divisiones en que interviene y promediada entre los árboles del bosque. La impureza de Gini de un nodo "
         "con proporciones de clase p_{i} es:",
         "Gini = 1 − Σ_{i} p_{i}^{2}",
         "Permite explicar el modelo al negocio: qué características impulsan la predicción. Como la importancia "
         "Gini tiende a favorecer variables continuas con muchos puntos de corte (Strobl et al., 2007), se "
         "contrastó con la importancia por permutación, que mide cuánto cae el AUC en Test al desordenar cada variable.",
         "`rf.feature_importances_` (Figura 7) y `permutation_importance(rf, X_test, y_test, scoring='roc_auc', "
         "n_repeats=10)` (Figura 8)."),
        ("Serialización con Joblib (Joblib vs. Pickle)",
         "Proceso de convertir objetos de Python que viven en memoria RAM en un archivo binario en disco para "
         "reutilizarlos después en otro proceso o servidor. Pickle es el módulo nativo de Python para objetos "
         "generales; Joblib está optimizado para objetos de scikit-learn que contienen grandes arreglos de NumPy, "
         "por lo que es más rápido y eficiente en espacio para modelos como un Random Forest.",
         None,
         "Sin serialización, el modelo, las reglas de sus árboles y las medianas de imputación se pierden al cerrar "
         "el proceso de Python; habría que reentrenar cada vez. Serializar separa el entrenamiento (costoso, offline) "
         "de la inferencia (rápida, online).",
         "`joblib.dump(paquete_despliegue, 'models/modelo_scouting_rf.joblib')` en `exportar_modelo.py` y "
         "`joblib.load(...)` una sola vez en el `lifespan` de `main_api.py`."),
        ("API REST y arquitectura cliente-servidor",
         "Una API REST (*Representational State Transfer*; Fielding, 2000) es un servicio web que expone recursos "
         "mediante URL y verbos HTTP (GET, POST...) e intercambia representaciones en JSON. En la arquitectura "
         "cliente-servidor, el servidor (la API) concentra el modelo y la lógica, y uno o varios clientes (el "
         "navegador, otra aplicación) le envían peticiones.",
         None,
         "Desacopla el modelo de quien lo usa: el aplicativo web, una app móvil o un sistema del club pueden "
         "consumir la misma predicción sin conocer Python ni scikit-learn. Pydantic valida el contrato de datos en "
         "la frontera del sistema.",
         "`main_api.py` con FastAPI: `POST /predict` recibe un `JugadorEntrada` en JSON y devuelve un "
         "`PrediccionSalida` con la clase y el porcentaje de probabilidad; `index.html` lo consume con `fetch`."),
    ]),
]


def marco_teorico(inf: Informe):
    inf.h1("2. Marco teórico y glosario")
    inf.p("Esta sección define con rigor los catorce conceptos exigidos por la guía de la actividad. Para cada uno se "
          "presentan su definición, su formulación cuando aplica, la razón por la que se usa y **cómo se aplicó "
          "concretamente en el código** del proyecto. Al final se incluyen los conceptos de despliegue y monitoreo "
          "de la última clase.")
    n = 0
    for i, (fase, conceptos) in enumerate(CONCEPTOS, 1):
        inf.h2(f"2.{i}. {fase}")
        for nombre, definicion, formula, porque, codigo in conceptos:
            n += 1
            inf.h3(f"2.{i}.{conceptos.index((nombre, definicion, formula, porque, codigo)) + 1}. {nombre}")
            inf.p(f"**Definición.** {definicion}")
            if formula:
                inf.formula(formula)
            inf.p(f"**¿Por qué se usa?** {porque}")
            inf.p(f"**Aplicación en el proyecto.** {codigo}")
    assert n == 14, n
    inf.h2("2.4. Conceptos complementarios de despliegue y monitoreo")
    inf.h3("2.4.1. Despliegue en CRISP-DM")
    inf.p("Un modelo no genera valor mientras viva solo en un *notebook*: en memoria RAM se pierde todo (pesos, "
          "reglas de los árboles, medianas de imputación) al cerrar el proceso. El despliegue es la fase que convierte "
          "el modelo en un servicio estable al que acceden los usuarios del negocio.")
    inf.h3("2.4.2. La «regla de oro»: guardar los transformadores junto al modelo")
    inf.p("Los datos nuevos deben recibir **exactamente el mismo preprocesamiento** que los datos de entrenamiento. "
          "Por eso el paquete `.joblib` no guarda solo el modelo, sino también `columnas_X`, `medianas_estatura`, "
          "`mediana_global_estatura`, `mediana_global_edad` y `umbral_elite_eur`. En este proyecto la regla se llevó "
          "también al código: la API, el script de inferencia, el entrenamiento y los experimentos llaman a las "
          "mismas funciones de `src/preprocessing/inferencia.py`.")
    inf.p("Al aplicar esta regla se detectó un defecto sutil en el script de inferencia de referencia: llamar a "
          "`get_dummies(drop_first=True)` sobre un lote nuevo elimina la primera posición *presente en ese lote*, no "
          "la categoría base del entrenamiento. En el lote de ejemplo (Defensa, Portero, Mediocampista) el Candidato A "
          "(Defensa) quedaba con todas sus columnas en cero, es decir, codificado como Delantero; y una petición "
          "individual de un portero perdía su única columna. La versión del proyecto codifica sin `drop_first` y luego "
          "alinea con `reindex(columns=columnas_X, fill_value=0)`, lo que garantiza columnas idénticas; el módulo "
          "incluye una autoverificación de este caso.")
    inf.h3("2.4.3. Data Drift y Concept Drift")
    inf.p("**Data Drift** (deriva de datos) ocurre cuando cambia con el tiempo la distribución de las variables de "
          "entrada X; por ejemplo, la inflación del mercado de fichajes desplaza todos los valores hacia arriba y el "
          "umbral Q3 calculado hoy deja de representar el Top 25 % de mañana. **Concept Drift** (deriva de concepto) "
          "ocurre cuando cambia la relación entre X y la variable objetivo Y; por ejemplo, cambios reglamentarios o "
          "tácticos que vuelven más o menos valiosas ciertas características físicas (Gama et al., 2014). Ambos "
          "degradan silenciosamente un modelo desplegado, por lo que exigen monitoreo y reentrenamiento periódico (§5.4).")


def resultados(inf: Informe):
    inf.h1("3. Análisis de resultados del código y experimentos")
    inf.h2("3.1. Descripción de los datos y del preprocesamiento")
    inf.p(f"La consulta SQL oficial filtra los jugadores sin valor de mercado o sin fecha de nacimiento, de modo que "
          f"de los {miles(50149)} registros de `players` quedaron **{miles(D['registros'])}** jugadores. La edad se "
          f"calcula en MySQL con `TIMESTAMPDIFF(YEAR, date_of_birth, CURDATE())`. La Tabla 2 resume el estado de los "
          f"datos antes y después del preprocesamiento.")
    tp = D["distribucion_posiciones"]
    inf.tabla(
        ["Indicador", "Valor"],
        [
            ["Registros extraídos de MySQL", miles(D["registros"])],
            ["Estaturas nulas antes de imputar", f"{miles(D['nulos_antes']['estatura_cm'])} "
             f"({D['nulos_antes']['estatura_cm'] / D['registros']:.2%})"],
            ["Edades nulas antes de imputar", miles(D["nulos_antes"]["edad"])],
            ["Nulos después de imputar (estatura / edad)", f"{D['nulos_despues']['estatura_cm']} / {D['nulos_despues']['edad']}"],
            ["Distribución de posiciones", "; ".join(f"{ETIQUETAS_POSICION[k]}: {miles(v)}" for k, v in tp.items())],
            ["Umbral Élite (Q3 del valor de mercado)", f"€ {miles(D['umbral_q3_eur'])}"],
            ["Valor de mercado mediano", f"€ {miles(D['valor_mediano_eur'])}"],
            ["Jugadores Élite / Estándar", f"{miles(D['n_elite'])} ({pct(D['proporcion_elite'])}) / "
             f"{miles(D['n_estandar'])} ({pct(1 - D['proporcion_elite'])})"],
            ["Partición estratificada Train / Test", f"{miles(D['n_train'])} / {miles(D['n_test'])} "
             f"(Élite: {pct(D['proporcion_elite_train'], 3)} / {pct(D['proporcion_elite_test'], 3)})"],
            ["Variables del modelo (X)", ", ".join(f"`{c}`" for c in D["columnas_X"])],
        ],
        "Resumen de los datos y del preprocesamiento", anchos=[6.5, 10], tam=10)
    inf.p(f"La proporción de Élite ({pct(D['proporcion_elite'])}) es ligeramente superior al 25 % teórico porque "
          f"muchos jugadores comparten exactamente el valor del umbral (€ {miles(D['umbral_q3_eur'])}, una cifra "
          f"redonda típica de Transfermarkt) y la regla `>=` los incluye a todos. El problema es, en todo caso, "
          f"claramente **desbalanceado** (aproximadamente 73/27), como muestra la Figura 2.")
    inf.figura(FIG / "01_distribucion_clases.png", "Distribución de la variable objetivo es_elite", ancho=12)
    med = D["mediana_estatura_por_posicion_total"]
    inf.p(f"La Figura 3 justifica la **imputación por mediana agrupada**: la mediana de estatura de los porteros "
          f"({med['Goalkeeper']:.0f} cm) supera en {med['Goalkeeper'] - med['Midfield']:.0f} cm a la de los "
          f"mediocampistas ({med['Midfield']:.0f} cm), y la de los defensas es {med['Defender']:.0f} cm. Imputar con la "
          f"mediana global ({D['mediana_global_estatura_train']:.0f} cm) habría asignado a un portero sin dato una "
          f"estatura de jugador de campo. La edad, sin nulos tras el filtro SQL, se imputa por seguridad con la "
          f"mediana global ({D['mediana_global_edad_train']:.0f} años) para soportar datos nuevos incompletos.")
    inf.figura(FIG / "02_estatura_por_posicion.png", "Distribución de la estatura por posición",
               f"Elaboración propia. El eje se limita a 150-212 cm; el dataset contiene "
               f"{D['estaturas_implausibles_menor_140']} estaturas físicamente imposibles (< 140 cm), errores de "
               f"captura que no se alteraron para conservar el preprocesamiento oficial.", ancho=14)
    inf.p("**Prevención de fuga de datos.** En los experimentos, las medianas se calcularon **solo con el conjunto "
          "de entrenamiento** y se aplicaron después al de prueba; si se calcularan con todo el dataset, información "
          "del test se filtraría al entrenamiento y las métricas serían optimistas. En el modelo final de producción "
          "(`exportar_modelo.py`) sí se usan todos los datos, porque ya no hay un conjunto de prueba que proteger.")

    inf.h3("3.1.1. Protocolo experimental")
    inf.p(f"Todos los modelos se entrenaron con el mismo conjunto de entrenamiento ({miles(D['n_train'])} jugadores) "
          f"y se evaluaron con el mismo conjunto de prueba ({miles(D['n_test'])} jugadores) que el modelo nunca vio, "
          f"con semilla fija (`random_state=42`) para garantizar la reproducibilidad. Las métricas de la clase Élite "
          f"(Precision, Recall, F1) se reportan porque es la clase de interés para el negocio. El script completo es "
          f"`experimentos_modelado.py` y su salida queda en `reports/resultados/`.")
    inf.h2("3.2. Paradoja de la exactitud")
    d = M["dummy"]
    pod = MOD["Árbol podado"]
    inf.p(f"Un clasificador trivial que **siempre** responde «Estándar» (`DummyClassifier(strategy='most_frequent')`) "
          f"obtiene una exactitud de **{pct(d['accuracy_test'])}** en Test, aunque su Recall para la clase Élite es "
          f"**{d['recall_elite']:.1f}**: no identifica ni a uno solo de los {miles(sum(d['matriz_confusion'][1]))} "
          f"jugadores Élite del conjunto de prueba. Cualquier modelo debe superar esta referencia en las métricas de la "
          f"clase minoritaria, no en exactitud.")
    inf.p(f"El experimento reveló que el **árbol podado sin balanceo cae exactamente en esta paradoja**: su accuracy "
          f"en Test ({pct(pod['accuracy_test'])}) es la más alta de todos los modelos y su brecha es nula, pero su Recall "
          f"Élite es {pod['recall_elite']:.1f}. Con solo cuatro niveles de profundidad, en ninguna de sus "
          f"{pod['hojas']} hojas los jugadores Élite llegan a ser mayoría, así que todas las hojas votan «Estándar». "
          f"Juzgado solo por la exactitud parecería el mejor modelo; en realidad es inútil para el scouting.")

    inf.h2("3.3. Diagnóstico de overfitting: árbol no restringido vs. árbol podado")
    orden = ["Árbol sin restricciones", "Árbol podado", "Árbol podado + balanced", "Random Forest optimizado"]
    inf.tabla(
        ["Modelo", "Acc. Train", "Acc. Test", "Brecha", "Precision Élite", "Recall Élite", "F1 Élite", "AUC"],
        [[n, f4(MOD[n]["accuracy_train"]), f4(MOD[n]["accuracy_test"]),
          f"{MOD[n]['brecha']:+.4f}", f4(MOD[n]["precision_elite"]), f4(MOD[n]["recall_elite"]),
          f4(MOD[n]["f1_elite"]), f4(MOD[n]["auc_test"])] for n in orden] +
        [["Dummy (most_frequent)", "–", f4(d["accuracy_test"]), "–", f4(d["precision_elite"]),
          f4(d["recall_elite"]), f4(d["f1_elite"]), "0.5000"]],
        "Comparación Train vs. Test y métricas de la clase Élite en el conjunto de prueba",
        "Elaboración propia a partir de `reports/resultados/tabla_overfitting.csv`. Brecha = Accuracy Train − "
        "Accuracy Test. Precision, Recall y F1 se calculan para la clase Élite (1).",
        anchos=[4.4, 1.6, 1.6, 1.6, 1.9, 1.7, 1.6, 1.6], tam=9)
    libre = MOD["Árbol sin restricciones"]
    inf.figura(FIG / "03_overfitting_train_vs_test.png", "Accuracy en entrenamiento y prueba por modelo", ancho=14.5)
    inf.p(f"**Árbol sin restricciones.** Creció hasta {libre['profundidad']} niveles y "
          f"{miles(libre['hojas'])} hojas, y su exactitud cae de {pct(libre['accuracy_train'])} en Train a "
          f"{pct(libre['accuracy_test'])} en Test: una brecha de **{libre['brecha'] * 100:.2f} puntos porcentuales**, "
          f"la mayor de todos los modelos. Es la firma del sobreajuste: el árbol creó reglas para combinaciones "
          f"muy específicas de edad y estatura que no se repiten en datos nuevos; además, su AUC en Test "
          f"({f4(libre['auc_test'])}) es el más bajo de los modelos entrenados.")
    inf.p(f"Que la exactitud de entrenamiento del árbol libre sea solo {pct(libre['accuracy_train'])} —y no cercana al "
          f"100 % como suele ocurrir— tiene una explicación estructural: con tres variables, los "
          f"{miles(D['registros'])} jugadores se reparten en apenas {miles(D['combinaciones_unicas_X'])} combinaciones "
          f"distintas de edad, estatura y posición. Muchos jugadores idénticos para el modelo tienen etiquetas "
          f"distintas, y ningún árbol puede separarlos: es un límite de información de los datos, no del algoritmo.")
    inf.p(f"**Árbol podado.** La poda previa (`max_depth=4`, `min_samples_leaf=20`) eliminó el sobreajuste "
          f"(brecha {pod['brecha']:+.4f}) y elevó el AUC a {f4(pod['auc_test'])}, lo que confirma que generaliza mejor "
          f"al ordenar a los jugadores por probabilidad. Sin embargo, como se vio en §3.2, su umbral de decisión por "
          f"defecto (50 %) lo lleva a predecir siempre Estándar. La poda resuelve la varianza, pero no el desbalance.")
    bal = MOD["Árbol podado + balanced"]
    rf = MOD["Random Forest optimizado"]
    inf.p(f"**Efecto de `class_weight='balanced'`.** Al penalizar más los errores sobre la clase Élite, el mismo árbol "
          f"podado pasa de Recall {pod['recall_elite']:.3f} a **{bal['recall_elite']:.3f}** y su F1 sube de 0 a "
          f"{bal['f1_elite']:.3f}. La exactitud baja a {pct(bal['accuracy_test'])}, un costo esperado y aceptable: el "
          f"modelo ahora acepta falsos positivos (jugadores Estándar señalados como Élite) a cambio de no dejar escapar a "
          f"los Élite. La Figura 5 lo muestra en las matrices de confusión.")
    inf.figura(FIG / "04_matrices_confusion.png",
               "Matrices de confusión en Test: árbol podado, árbol podado balanceado y Random Forest", ancho=16)
    cm = rf["matriz_confusion"]
    inf.p(f"En términos de negocio, de los {miles(cm[1][0] + cm[1][1])} jugadores Élite del conjunto de prueba, el "
          f"Random Forest detecta {miles(cm[1][1])} ({pct(rf['recall_elite'], 1)}) y deja escapar {miles(cm[1][0])}. "
          f"A cambio, marca como candidatos a {miles(cm[0][1])} jugadores Estándar, por lo que su precisión es "
          f"{pct(rf['precision_elite'], 1)}: aproximadamente 4 de cada 10 jugadores que el sistema recomienda observar "
          f"son realmente Élite, frente a {pct(D['proporcion_elite'], 1)} si se eligieran al azar.")

    inf.h2("3.4. Evaluación del ensamble y curva ROC-AUC")
    mp = GS["mejores_parametros"]
    malla = GS["malla"]
    inf.p(f"El Random Forest se optimizó con `GridSearchCV` sobre una malla de {GS['n_combinaciones']} combinaciones "
          f"(`n_estimators` {malla['n_estimators']}, `max_depth` {malla['max_depth']}, `min_samples_leaf` "
          f"{malla['min_samples_leaf']}, `max_features` {malla['max_features']}) con validación cruzada "
          f"estratificada de 5 particiones ({GS['ajustes']} entrenamientos en {GS['duracion_min']:.1f} minutos), "
          f"optimizando ROC-AUC y registrando también F1. No fue necesario reducir la malla.".replace("'", ""))
    top = pd.read_csv(RES / "gridsearch_resultados.csv")
    inf.tabla(
        ["Rango", "n_estimators", "max_depth", "min_samples_leaf", "max_features", "AUC CV (media ± desv.)", "F1 CV"],
        [[int(r.rank_test_roc_auc), int(r.param_n_estimators),
          "None" if pd.isna(r.param_max_depth) else int(r.param_max_depth), int(r.param_min_samples_leaf),
          "None" if pd.isna(r.param_max_features) else r.param_max_features,
          f"{r.mean_test_roc_auc:.4f} ± {r.std_test_roc_auc:.4f}", f"{r.mean_test_f1:.4f}"]
         for r in top.head(5).itertuples()],
        "Las cinco mejores combinaciones de hiperparámetros del GridSearchCV",
        "Elaboración propia a partir de `reports/resultados/gridsearch_resultados.csv` (top 10 completo).",
        anchos=[1.4, 2.2, 2.0, 2.6, 2.3, 3.6, 1.7], tam=9)
    inf.p(f"La mejor configuración fue **`n_estimators={mp['n_estimators']}`, `max_depth={mp['max_depth']}`, "
          f"`min_samples_leaf={mp['min_samples_leaf']}`, `max_features='{mp['max_features']}'`**, con un ROC-AUC de "
          f"validación cruzada de **{f4(GS['mejor_auc_cv'])} ± {f4(GS['std_auc_cv'])}** y F1 de {f4(GS['f1_cv'])}. "
          f"Dos observaciones: (1) las diez mejores combinaciones están separadas por menos de una milésima de AUC, "
          f"es decir, el modelo es poco sensible a los hiperparámetros dentro de un rango razonable; (2) todas usan "
          f"`max_features='sqrt'`, lo que confirma el valor del *feature subspacing* para descorrelacionar los árboles. "
          f"El AUC obtenido en Test ({f4(rf['auc_test'])}) es coherente con el de validación cruzada, señal de que la "
          f"búsqueda no sobreajustó los hiperparámetros.")
    inf.figura(FIG / "05_curvas_roc.png", "Curvas ROC comparativas en Test: árbol único vs. Random Forest", ancho=11.5)
    inf.p(f"**Interpretación del AUC.** El Random Forest optimizado alcanza **AUC = {f4(rf['auc_test'])}**, frente a "
          f"{f4(libre['auc_test'])} del árbol sin restricciones y {f4(pod['auc_test'])} del árbol podado. Esto significa "
          f"que, si se toma al azar un jugador Élite y uno Estándar, el modelo asigna mayor probabilidad de Élite al "
          f"primero en el {pct(rf['auc_test'], 1)} de los casos (frente al 50 % del azar). La curva naranja domina a "
          f"las demás en casi todo el rango de FPR, en especial en la zona de 0.2 a 0.5, que es la región operativa de "
          f"un filtro de scouting. La mejora del ensamble sobre el árbol único "
          f"(+{(rf['auc_test'] - libre['auc_test']):.4f} frente al libre y +{(rf['auc_test'] - pod['auc_test']):.4f} "
          f"frente al podado) es consistente con la teoría del *bagging*: promediar {mp['n_estimators']} árboles "
          f"reduce la varianza sin aumentar el sesgo.")
    inf.p("Con honestidad, un AUC cercano a 0.70 indica una capacidad de discriminación **moderada**: útil como primer "
          "filtro para priorizar candidatos, pero insuficiente para decidir un fichaje por sí sola. El techo no lo "
          "impone el algoritmo, sino la información disponible: tres variables físicas y demográficas no pueden "
          "capturar el rendimiento deportivo, la liga o el club, que son los principales determinantes del valor de "
          "mercado. Este resultado delimita el alcance del sistema (§5).")

    inf.h2("3.5. Importancia de variables: ¿qué variables pesan más y por qué?")
    g, pm = IMP["gini"], IMP["permutacion"]
    inf.figura(FIG / "06_feature_importance.png", "Importancia de variables (Gini) del Random Forest optimizado",
               ancho=14.5)
    inf.p(f"La Figura 7 muestra una jerarquía muy marcada: **la edad concentra {pct(g['edad'], 1)} de la importancia "
          f"Gini**, seguida por la estatura ({pct(g['estatura_cm'], 1)}) y la condición de portero "
          f"({pct(g['posicion_Goalkeeper'], 1)}); las demás posiciones aportan cerca de 1 % cada una. La importancia por "
          f"permutación (Figura 8) confirma las mismas tres variables en el top, aunque intercambia el segundo y el "
          f"tercer lugar: permutar la edad reduce el AUC en {pm['edad']:.3f}, mientras que la posición de portero "
          f"({pm['posicion_Goalkeeper']:.3f}) y la estatura ({pm['estatura_cm']:.3f}) pesan prácticamente lo mismo.")
    inf.figura(FIG / "07_permutation_importance.png",
               "Importancia por permutación (caída del ROC-AUC en Test) del Random Forest optimizado",
               "Elaboración propia. Barras de error: desviación estándar en 10 repeticiones.", ancho=14.5)
    te, tp_ = D["tasa_elite_edad"], D["tasa_elite_posicion"]
    pico = max(te, key=lambda k: te[k]["tasa"])
    inf.tabla(
        ["Rango de edad (años)", "Jugadores", "Tasa de Élite"],
        [[k, miles(v["n"]), pct(v["tasa"], 1)] for k, v in te.items()],
        "Proporción de jugadores Élite por rango de edad",
        "Elaboración propia a partir de `metricas.json`. Edad calculada a la fecha de ejecución (noviembre de 2026).",
        anchos=[5, 4, 4])
    mayores = te["≥36"]
    inf.p(f"**¿Por qué la edad domina?** La Tabla 5 muestra la relación que el modelo aprendió: la tasa de Élite se "
          f"mantiene entre {pct(min(te[k]['tasa'] for k in ('≤20', '21-23', '24-26', '27-29')), 1)} y "
          f"{pct(te[pico]['tasa'], 1)} hasta los 29 años (máximo en {pico} años) y a partir de ahí se desploma: "
          f"{pct(te['30-32']['tasa'], 1)} entre 30 y 32 años, {pct(te['33-35']['tasa'], 1)} entre 33 y 35 y solo "
          f"{pct(mayores['tasa'], 1)} desde los 36. Es la conocida **curva de valor por edad** del mercado de fichajes: "
          f"los clubes pagan por los jóvenes su potencial y su valor de reventa, el precio se sostiene mientras el "
          f"jugador está en su plenitud física y cae a partir de la treintena, cuando se descuentan los pocos años de "
          f"carrera restantes. La relación no es lineal (una meseta seguida de una caída), justo el tipo de patrón que "
          f"un árbol captura con un par de cortes y que explica la gran reducción de impureza que aporta la edad.")
    inf.p(f"Hay además un efecto del propio dataset que amplifica el peso de la edad: **{miles(mayores['n'])} jugadores "
          f"({mayores['n'] / D['registros']:.1%} del total) tienen 36 años o más**, una proporción muy superior a la de "
          f"futbolistas activos de esa edad. La tabla `players` conserva jugadores de temporadas antiguas cuyo último "
          f"valor de mercado quedó congelado al retirarse o dejar las ligas cubiertas; como la edad se calcula a la "
          f"fecha actual, aparecen como veteranos con valores bajos. Parte del poder predictivo de la edad consiste, "
          f"por tanto, en reconocer jugadores inactivos, algo que debería corregirse filtrando por jugadores activos (§5.4).")
    inf.p(f"**¿Por qué importa la posición de portero y no las demás?** La tasa de Élite de los porteros es "
          f"{pct(tp_['Goalkeeper']['tasa'], 1)}, frente a {pct(tp_['Attack']['tasa'], 1)} de los delanteros, "
          f"{pct(tp_['Midfield']['tasa'], 1)} de los mediocampistas y {pct(tp_['Defender']['tasa'], 1)} de los defensas. "
          f"El mercado cotiza menos a los porteros: cada equipo alinea uno solo, su influencia en goles y asistencias "
          f"—las estadísticas que más mueven el precio— es indirecta, y los traspasos millonarios de porteros son "
          f"excepcionales; en cambio, delanteros y mediocampistas ofensivos concentran las mayores cotizaciones. Las "
          f"tres posiciones de campo tienen tasas parecidas entre sí, por eso sus columnas *dummy* aportan poco: la "
          f"única distinción posicional realmente informativa es «portero frente a jugador de campo».")
    inf.p("**¿Y la estatura?** Su importancia Gini (segunda) es mayor que la de permutación (tercera, empatada con "
          "portero). Hay dos razones. Primera, un sesgo conocido: la importancia Gini favorece a las variables "
          "continuas, que ofrecen muchos puntos de corte posibles, frente a las binarias, que ofrecen uno solo "
          "(Strobl et al., 2007); por eso la permutación, que mide el efecto real sobre el AUC en datos no vistos, es la "
          "referencia más fiable. Segunda, la estatura contiene información en parte redundante con la posición (los "
          "porteros son más altos), de modo que el bosque puede usarla como sustituto de la posición. En el fútbol la "
          "estatura es un requisito en ciertos roles (porteros y defensas centrales), pero no un generador de valor por "
          "sí misma: la técnica, la velocidad y el rendimiento —que no están en el modelo— pesan mucho más.")
    inf.p("**Limitaciones de un modelo con tres variables.** El valor de mercado depende sobre todo de variables "
          "ausentes: goles, asistencias y minutos jugados, la liga y el club donde compite el jugador, su nacionalidad "
          "y selección, la duración de su contrato y su historial de lesiones. Por eso el modelo debe entenderse como un "
          "**filtro demográfico inicial**, no como un estimador del talento.")


def despliegue(inf: Informe):
    inf.h1("4. Arquitectura de despliegue")
    inf.p("La Figura 9 muestra el flujo completo de datos. En la etapa de **entrenamiento** (offline), "
          "`connection2.py` ejecuta la consulta SQL en MySQL (XAMPP) y entrega un DataFrame; `exportar_modelo.py` "
          "imputa, codifica, entrena el Random Forest con los mejores hiperparámetros y serializa un paquete con el "
          "modelo y sus transformadores. En la etapa de **producción** (online), `main_api.py` carga ese paquete una "
          "sola vez al iniciar y atiende peticiones HTTP; `index.html` envía los datos del jugador con `fetch` y "
          "muestra la respuesta al scout.")
    inf.figura(FIG / "08_arquitectura_despliegue.png", "Arquitectura de despliegue y flujo de datos", ancho=16)
    inf.h2("4.1. Exportación del modelo")
    inf.p(f"El modelo final se reentrena con el 100 % de los datos ({miles(D['registros'])} jugadores). Es correcto "
          f"porque los hiperparámetros ya se eligieron con validación cruzada sobre el conjunto de entrenamiento y el "
          f"desempeño ya se estimó en el conjunto de prueba; usar todos los registros solo le da al modelo de "
          f"producción más información. El paquete conserva las claves del script de referencia (`modelo`, "
          f"`columnas_X`, `medianas_estatura`, `mediana_global_estatura`, `mediana_global_edad`, `umbral_elite_eur`) y "
          f"agrega metadatos: `posiciones_validas`, `hiperparametros`, `metricas_test`, `fecha_entrenamiento`, "
          f"`version_sklearn` y `n_registros`. El script verifica al final que el archivo recargado con `joblib.load` "
          f"produce las mismas predicciones.")
    inf.h2("4.2. Contrato JSON de la API")
    inf.p("La API expone cuatro *endpoints*: `GET /` (estado de la API y del modelo), `GET /info` (metadatos), "
          "**`POST /predict`** (un jugador) y `POST /predict/lote` (lista de jugadores). Pydantic valida la entrada: la "
          "posición debe ser uno de los valores del entrenamiento (`Attack`, `Defender`, `Goalkeeper`, `Midfield`, "
          "`Missing`), la edad un entero entre 15 y 45 y la estatura un número entre 150 y 215 o `null`; cualquier "
          "violación devuelve un error **422** antes de llegar al modelo. Ejemplo de petición y respuesta reales:")
    ej = json.loads((CAP / "ejemplo_predict.json").read_text(encoding="utf-8"))
    inf.codigo("POST http://127.0.0.1:8000/predict\nContent-Type: application/json\n\n" +
               json.dumps(ej["request"], indent=2, ensure_ascii=False))
    inf.codigo("HTTP/1.1 200 OK\n\n" + json.dumps(ej["response"], indent=2, ensure_ascii=False))
    inf.p("Si la estatura llega como `null`, la API la imputa con la mediana de la posición guardada en el paquete y "
          "lo informa con `estatura_imputada: true`. El conjunto de pruebas automatizadas `tests/test_api.py` "
          "(pytest + `TestClient`) verifica el health check, una predicción normal, la imputación de estatura nula, "
          "el rechazo de una posición inválida y de edades fuera de rango (422), y el endpoint por lotes.")
    inf.h2("4.3. API en ejecución")
    inf.figura(CAP / "api_swagger_docs.png", "Documentación interactiva Swagger UI generada por FastAPI (/docs)",
               "Captura propia de la API en ejecución (`uvicorn main_api:app --port 8000`).", ancho=14, alto_max=15)
    inf.figura(CAP / "api_predict_respuesta.png",
               "Ejecución de POST /predict desde Swagger con un jugador nuevo y respuesta 200",
               "Captura propia.", ancho=15, alto_max=20)
    inf.figura(CAP / "api_health.png", "Respuesta del endpoint GET /info con los metadatos del modelo desplegado",
               "Captura propia.", ancho=15, alto_max=19)
    inf.figura(CAP / "api_consola_uvicorn.png",
               "Registro de la consola de uvicorn con las peticiones atendidas durante las pruebas",
               "Captura propia del log del servidor (respuestas 200 en verde; validación 422 en naranja).", ancho=16)
    inf.h2("4.4. Aplicativo web respondiendo a jugadores nuevos")
    inf.p("`index.html` es un único archivo con HTML, CSS y JavaScript sin dependencias externas, de modo que funciona "
          "sin conexión a internet. Al cargar consulta `GET /` y muestra si la API está en línea; el formulario "
          "traduce las posiciones al español, permite marcar «Estatura desconocida» (envía `null`), valida los rangos en "
          "el cliente y ofrece botones de ejemplo rápido con los candidatos A, B y C del docente y dos jugadores "
          "adicionales. El resultado se muestra con una insignia Élite/Estándar y una barra de progreso animada cuyo "
          "color cambia con la probabilidad; cada evaluación se agrega al historial de la sesión. Los errores 422 se "
          "traducen a mensajes comprensibles y, si la API está apagada, se indica cómo iniciarla.")
    inf.figura(CAP / "web_formulario.png", "Aplicativo web al iniciar, con la API en línea",
               "Captura propia (Chromium).", ancho=15, alto_max=14)
    inf.figura(CAP / "web_resultado_elite.png", "Predicción de un jugador nuevo clasificado como Élite",
               "Captura propia.", ancho=15, alto_max=14)
    inf.figura(CAP / "web_resultado_estandar.png", "Predicción de un jugador nuevo clasificado como Estándar",
               "Captura propia.", ancho=15, alto_max=14)
    inf.figura(CAP / "web_historial.png", "Historial de la sesión con cinco jugadores nuevos evaluados",
               "Captura propia. El asterisco junto a la estatura indica que fue imputada.", ancho=15, alto_max=17)


def conclusiones(inf: Informe):
    rf = MOD["Random Forest optimizado"]
    libre = MOD["Árbol sin restricciones"]
    inf.h1("5. Conclusiones y recomendaciones de negocio")
    inf.h2("5.1. Conclusiones")
    inf.bullets([
        f"Se construyó y desplegó un ecosistema analítico completo: MySQL → Python → modelo serializado → API REST → "
        f"aplicativo web, cubriendo las seis fases de CRISP-DM. El sistema responde en tiempo real a jugadores nuevos "
        f"y está respaldado por pruebas automatizadas.",
        f"El árbol de decisión sin restricciones evidenció sobreajuste (brecha de {libre['brecha'] * 100:.2f} puntos y "
        f"el AUC más bajo, {libre['auc_test']:.3f}); la poda previa eliminó la brecha y mejoró el AUC.",
        f"La exactitud resultó una métrica engañosa: el modelo trivial y el árbol podado sin balanceo superan el 73 % "
        f"de accuracy sin detectar a un solo jugador Élite. `class_weight='balanced'` fue decisivo para que los "
        f"modelos fueran útiles (Recall Élite de 0 a {rf['recall_elite']:.3f} en el Random Forest).",
        f"El Random Forest optimizado con GridSearchCV y K-Fold estratificado es el mejor modelo (AUC = "
        f"{rf['auc_test']:.4f}, F1 Élite = {rf['f1_elite']:.4f}, brecha de solo {rf['brecha'] * 100:.2f} puntos), lo que "
        f"confirma la ventaja del ensamble sobre el árbol único.",
        "La edad es, con diferencia, la variable más influyente, coherente con la curva de valor por edad del mercado; "
        "ser portero reduce la probabilidad de ser Élite y la estatura aporta poco por sí sola.",
    ])
    inf.h2("5.2. Cómo apoya el sistema la toma de decisiones en scouting")
    inf.bullets([
        f"**Filtro inicial de candidatos.** Ante una lista larga de perfiles, el sistema descarta rápidamente a los de "
        f"baja probabilidad y conserva al {pct(rf['recall_elite'], 0)} de los jugadores Élite, con una tasa de acierto "
        f"entre los recomendados de {pct(rf['precision_elite'], 0)}, muy superior a la tasa base del "
        f"{pct(D['proporcion_elite'], 0)}.",
        "**Priorización de visorías.** La probabilidad (y no solo la etiqueta) permite ordenar a los candidatos y "
        "asignar primero los viajes y las horas de análisis a quienes tienen mayor potencial de mercado.",
        "**Ahorro de costos y decisiones trazables.** Cada recomendación queda registrada con datos y criterios "
        "explícitos, lo que reduce el sesgo subjetivo y facilita justificar las decisiones ante la dirección deportiva.",
        "**Integración.** Al exponerse como API REST, el mismo modelo puede consumirse desde el aplicativo web, una "
        "hoja de cálculo, una aplicación móvil o el sistema de gestión deportiva del club, y el endpoint por lotes "
        "permite evaluar listas completas de una sola vez.",
    ])
    inf.h2("5.3. Limitaciones")
    inf.bullets([
        "**Pocas variables.** Edad, estatura y posición explican solo una parte del valor de mercado; el AUC moderado "
        "refleja ese techo de información.",
        "**Sesgo del valor de mercado como target.** El valor de Transfermarkt es una estimación colectiva que mezcla "
        "talento con factores comerciales (liga, club, nacionalidad, contrato). El modelo predice valor de mercado, no "
        "calidad deportiva.",
        "**Temporalidad e inflación.** Las valoraciones de distintas épocas no son comparables y la edad se calcula a la "
        "fecha actual, también para jugadores con valoraciones antiguas.",
        "**Calidad de datos.** Existen estaturas físicamente imposibles y una categoría de posición «Missing» que "
        "deberían depurarse en la fuente.",
    ])
    inf.h2("5.4. Recomendaciones y trabajo futuro")
    inf.bullets([
        "**Más variables de rendimiento**: integrar las tablas `appearances`, `games`, `clubs` y `competitions` del "
        "mismo dataset (goles, asistencias, minutos, nivel de la liga y del club) y el pie hábil del jugador.",
        "**Filtrar jugadores activos y valoraciones recientes** (tabla `player_valuations`) para que la edad y el valor "
        "correspondan al mismo momento.",
        "**Reentrenamiento periódico** (por ejemplo, después de cada ventana de fichajes) recalculando el umbral Q3, y "
        "**monitoreo de Data Drift y Concept Drift**: comparar la distribución de las entradas y el desempeño del "
        "modelo frente a las valoraciones reales para detectar degradación.",
        "**Ajuste del umbral de decisión** según el costo de negocio: un club que prefiera no perder talento puede bajar "
        "el umbral (más recall); uno con poco presupuesto de visoría puede subirlo (más precisión).",
        "**Seguridad de la API** antes de exponerla fuera de la red local: autenticación (API key u OAuth2), CORS "
        "restringido a los dominios del club, HTTPS, límites de peticiones y registro de auditoría.",
        "**Explicabilidad individual** (por ejemplo, valores SHAP) para mostrar al scout por qué el modelo calificó así "
        "a cada jugador, y empaquetado con Docker para un despliegue reproducible.",
    ])


REFERENCIAS = [
    "Breiman, L. (2001). Random forests. *Machine Learning, 45*(1), 5–32. https://doi.org/10.1023/A:1010933404324",
    "Breiman, L., Friedman, J. H., Olshen, R. A., & Stone, C. J. (1984). *Classification and regression trees*. "
    "Wadsworth.",
    "Cariboo, D. (2026). *Football data from Transfermarkt* [Conjunto de datos]. Kaggle. "
    "https://www.kaggle.com/datasets/davidcariboo/player-scores",
    "Chapman, P., Clinton, J., Kerber, R., Khabaza, T., Reinartz, T., Shearer, C., & Wirth, R. (2000). *CRISP-DM 1.0: "
    "Step-by-step data mining guide*. SPSS Inc.",
    "Fawcett, T. (2006). An introduction to ROC analysis. *Pattern Recognition Letters, 27*(8), 861–874. "
    "https://doi.org/10.1016/j.patrec.2005.10.010",
    "Fielding, R. T. (2000). *Architectural styles and the design of network-based software architectures* [Tesis "
    "doctoral, University of California, Irvine].",
    "Gama, J., Žliobaitė, I., Bifet, A., Pechenizkiy, M., & Bouchachia, A. (2014). A survey on concept drift "
    "adaptation. *ACM Computing Surveys, 46*(4), Artículo 44. https://doi.org/10.1145/2523813",
    "Joblib developers. (2026). *Joblib: Running Python functions as pipeline jobs* [Documentación de software]. "
    "https://joblib.readthedocs.io/",
    "Kohavi, R. (1995). A study of cross-validation and bootstrap for accuracy estimation and model selection. En "
    "*Proceedings of the 14th International Joint Conference on Artificial Intelligence* (Vol. 2, pp. 1137–1143). "
    "Morgan Kaufmann.",
    "Mejía Villalba, R. (2026). *Actividad final y guía de la última clase: despliegue de modelos* [Material de clase]. "
    "Electiva II – Minería de Datos, Universidad Cooperativa de Colombia, sede Montería.",
    "Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., Grisel, O., Blondel, M., Prettenhofer, P., "
    "Weiss, R., Dubourg, V., Vanderplas, J., Passos, A., Cournapeau, D., Brucher, M., Perrot, M., & Duchesnay, É. "
    "(2011). Scikit-learn: Machine learning in Python. *Journal of Machine Learning Research, 12*, 2825–2830.",
    "Pydantic. (2026). *Pydantic documentation* (Versión 2) [Documentación de software]. https://docs.pydantic.dev/",
    "Ramírez, S. (2026). *FastAPI* [Documentación de software]. https://fastapi.tiangolo.com/",
    "scikit-learn developers. (2026). *scikit-learn user guide* (Versión 1.9) [Documentación de software]. "
    "https://scikit-learn.org/stable/user_guide.html",
    "Strobl, C., Boulesteix, A.-L., Zeileis, A., & Hothorn, T. (2007). Bias in random forest variable importance "
    "measures: Illustrations, sources and a solution. *BMC Bioinformatics, 8*, Artículo 25. "
    "https://doi.org/10.1186/1471-2105-8-25",
]


def referencias(inf: Informe):
    inf.h1("6. Referencias")
    for ref in REFERENCIAS:
        par = texto_rico(inf.doc.add_paragraph(), ref)
        par.alignment = WD_ALIGN_PARAGRAPH.LEFT
        par.paragraph_format.left_indent = Cm(1.27)
        par.paragraph_format.first_line_indent = Cm(-1.27)


def anexos(inf: Informe):
    inf.h1("7. Anexos")
    inf.h2("Anexo A. Estructura de carpetas del proyecto")
    inf.codigo("""Electiva II/
├── src/
│   ├── database/connection2.py       Conector MySQL (mysql-connector-python + .env)
│   ├── preprocessing/inferencia.py   Preprocesamiento compartido ("regla de oro")
│   └── estilo.py                     Paleta y estilo de gráficas
├── experimentos_modelado.py          EDA, overfitting, GridSearchCV, ROC, importancias
├── exportar_modelo.py                Entrenamiento final y serialización (Joblib)
├── inferencia_nuevos_jugadores.py    Simulación de inferencia (Candidatos A, B, C)
├── main_api.py                       API REST (FastAPI + Pydantic)
├── index.html                        Aplicativo web cliente (HTML/CSS/JS)
├── models/modelo_scouting_rf.joblib  Paquete del modelo exportado
├── tests/test_api.py                 Pruebas automatizadas de la API (pytest)
├── scripts/                          Diagramas, capturas e informe
├── reports/
│   ├── figuras/                      Figuras 01-09
│   ├── capturas/                     Evidencias de la API y del aplicativo web
│   └── resultados/                   metricas.json, tablas CSV, mejores_parametros.json
├── requirements.txt · README.md · .env.example
└── Informe_Final_Scouting_David_Cuello.pdf""", tam=8.5)
    inf.h2("Anexo B. Instrucciones de ejecución")
    inf.p("Requisitos: Python 3.11 o superior, MySQL/MariaDB (XAMPP) con la base `mineria_datos` cargada y un "
          "archivo `.env` con las credenciales (ver `.env.example`). En PowerShell, desde la carpeta del proyecto:")
    inf.codigo("""python -m venv .venv
.venv\\Scripts\\Activate.ps1
pip install -r requirements.txt

python experimentos_modelado.py          # 1. Experimentos y evidencias
python exportar_modelo.py                # 2. Exporta el modelo .joblib
python inferencia_nuevos_jugadores.py    # 3. (Opcional) Prueba de inferencia
python -m pytest tests -v                # 4. (Opcional) Pruebas de la API
uvicorn main_api:app --reload --port 8000   # 5. API en http://127.0.0.1:8000/docs
start index.html                         # 6. Abre el aplicativo web""", tam=8.5)
    inf.p("Ejemplo de consumo de la API desde PowerShell:")
    inf.codigo("""$body = @{ jugador="Candidato A"; posicion="Defender"; edad=22; estatura_cm=188 } |
         ConvertTo-Json
Invoke-RestMethod -Uri http://127.0.0.1:8000/predict -Method Post `
                  -Body $body -ContentType "application/json\"""", tam=8.5)


def main():
    inf = Informe()
    portada(inf)
    tabla_contenido(inf)
    introduccion(inf)
    marco_teorico(inf)
    resultados(inf)
    despliegue(inf)
    conclusiones(inf)
    referencias(inf)
    anexos(inf)
    inf.guardar()
    print(f"DOCX generado: {SALIDA} ({inf.nfig} figuras, {inf.ntab} tablas)")


if __name__ == "__main__":
    main()
