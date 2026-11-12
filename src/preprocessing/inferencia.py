import pandas as pd

COLUMNAS_NUMERICAS = ["edad", "estatura_cm"]


def limpiar_tipos(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["estatura_cm"] = pd.to_numeric(df["estatura_cm"], errors="coerce")
    df["edad"] = pd.to_numeric(df["edad"], errors="coerce")
    df["posicion"] = df["posicion"].fillna("Sin Posicion")
    return df


def calcular_medianas(df: pd.DataFrame) -> dict:
    return {
        "medianas_estatura": df.groupby("posicion")["estatura_cm"].median().to_dict(),
        "mediana_global_estatura": float(df["estatura_cm"].median()),
        "mediana_global_edad": float(df["edad"].median()),
    }


def imputar(df: pd.DataFrame, medianas_estatura: dict, mediana_global_estatura: float,
            mediana_global_edad: float) -> pd.DataFrame:
    df = df.copy()
    df["estatura_imputada"] = df["estatura_cm"].isna()
    df["estatura_cm"] = df["estatura_cm"].fillna(df["posicion"].map(medianas_estatura))
    df["estatura_cm"] = df["estatura_cm"].fillna(mediana_global_estatura)
    df["edad"] = df["edad"].fillna(mediana_global_edad)
    return df


def codificar(df: pd.DataFrame, columnas_X: list | None = None) -> pd.DataFrame:
    if columnas_X is None:
        dummies = pd.get_dummies(df[["posicion"]], drop_first=True, dtype=int)
        return pd.concat([df[COLUMNAS_NUMERICAS], dummies], axis=1)
    dummies = pd.get_dummies(df[["posicion"]], dtype=int)
    X = pd.concat([df[COLUMNAS_NUMERICAS], dummies], axis=1)
    return X.reindex(columns=columnas_X, fill_value=0)


def preparar_para_modelo(df_crudo: pd.DataFrame, paquete: dict) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = imputar(limpiar_tipos(df_crudo), paquete["medianas_estatura"],
                 paquete["mediana_global_estatura"], paquete["mediana_global_edad"])
    return codificar(df, paquete["columnas_X"]), df


if __name__ == "__main__":
    paq = {"medianas_estatura": {"Goalkeeper": 190.0, "Attack": 181.0},
           "mediana_global_estatura": 182.0, "mediana_global_edad": 26.0,
           "columnas_X": ["edad", "estatura_cm", "posicion_Defender", "posicion_Goalkeeper"]}
    X, d = preparar_para_modelo(pd.DataFrame([{"posicion": "Goalkeeper", "edad": 29, "estatura_cm": None}]), paq)
    assert X.loc[0, "posicion_Goalkeeper"] == 1 and X.loc[0, "posicion_Defender"] == 0
    assert X.loc[0, "estatura_cm"] == 190.0 and bool(d.loc[0, "estatura_imputada"])
    X, _ = preparar_para_modelo(pd.DataFrame([{"posicion": "Attack", "edad": None, "estatura_cm": 175}]), paq)
    assert X.loc[0, ["posicion_Defender", "posicion_Goalkeeper"]].sum() == 0 and X.loc[0, "edad"] == 26.0
    print("OK: preprocesamiento de inferencia verificado")
