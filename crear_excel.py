import pandas as pd


def generar_excel() -> None:
    datos = {
        "dni": [10526358, 15121313, 15161414]
    }

    df = pd.DataFrame(datos)
    df.to_excel("dnis.xlsx", index=False, engine="openpyxl")


if __name__ == "__main__":
    generar_excel()
    print("Archivo dnis.xlsx generado correctamente.")
