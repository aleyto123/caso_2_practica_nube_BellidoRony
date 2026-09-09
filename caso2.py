import re
from typing import Any
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://consultaelectoral.onpe.gob.pe/inicio"


def extraer_por_label(soup: BeautifulSoup, labels: list[str]) -> str:
    """Busca en el HTML un valor asociado a etiquetas como nombre, región, distrito, etc."""
    text = "\n".join(
        " ".join(part.strip().split())
        for part in soup.stripped_strings
    )

    for label in labels:
        pattern = rf"{label}\s*[:\-]?\s*([^\n\r]+)"
        match = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
        if match:
            value = re.sub(r"\s+", " ", match.group(1)).strip(" :;-\t\r\n")
            if value:
                return value

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        for row in rows:
            cells = [cell.get_text(" ", strip=True) for cell in row.find_all(["th", "td"])]
            if len(cells) >= 2:
                for label in labels:
                    if cells[0].lower().startswith(label.lower()) or label.lower() in cells[0].lower():
                        return cells[1]

    return "No disponible"


def normalizar_texto(value: Any) -> str:
    if value is None:
        return "No disponible"
    texto = str(value).strip()
    return re.sub(r"\s+", " ", texto) if texto else "No disponible"


def detectar_miembro_mesa(soup: BeautifulSoup) -> str:
    text = " ".join(soup.stripped_strings).lower()

    if re.search(r"miembro\s+de\s+mesa|miembro\s+mesa|mesa\s+electoral", text):
        return "Sí"
    if re.search(r"no\s+es\s+miembro|no\s+miembro|no\s+participa", text):
        return "No"

    for keyword in ["miembro", "mesa"]:
        if keyword in text:
            return "Sí"
    return "No"


def consultar_dni_onpe(dni: str) -> dict[str, Any]:
    """Consulta información del DNI en la web de la ONPE y devuelve un diccionario normalizado."""
    resultado = {
        "dni": dni,
        "miembro_de_mesa": "No",
        "nombres": "No disponible",
        "ubicacion": "No disponible",
        "direccion_local": "No disponible",
        "error": None,
    }

    try:
        session = requests.Session()
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
            "Accept-Language": "es-PE,es;q=0.9",
        })

        response = session.get(BASE_URL, timeout=20)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")

        form = soup.find("form")
        action = "/inicio"
        method = "get"
        payload = {}

        if form is not None:
            action = form.get("action") or "/inicio"
            method = (form.get("method") or "get").lower()
            for input_tag in form.find_all("input"):
                name = input_tag.get("name")
                if name:
                    payload[name] = input_tag.get("value", "")

        payload["dni"] = dni
        payload["numero"] = dni
        payload["documento"] = dni
        payload["tipoDocumento"] = "DNI"

        if form is not None and method == "post":
            final_response = session.post(urljoin(BASE_URL, action), data=payload, timeout=20)
        else:
            final_response = session.get(urljoin(BASE_URL, action), params=payload, timeout=20)

        final_response.raise_for_status()
        result_soup = BeautifulSoup(final_response.text, "html.parser")

        resultado["miembro_de_mesa"] = detectar_miembro_mesa(result_soup)
        resultado["nombres"] = normalizar_texto(
            extraer_por_label(result_soup, ["nombres", "nombre", "apellidos", "apellido"])
        )

        region = extraer_por_label(result_soup, ["región", "region"])
        provincia = extraer_por_label(result_soup, ["provincia", "prov"])
        distrito = extraer_por_label(result_soup, ["distrito", "dist"])
        resultado["ubicacion"] = f"{region} / {provincia} / {distrito}"

        resultado["direccion_local"] = normalizar_texto(
            extraer_por_label(
                result_soup,
                [
                    "local de votación",
                    "local",
                    "dirección del local",
                    "direccion",
                    "dirección",
                    "domicilio",
                    "lugar de votación",
                ],
            )
        )

        if resultado["nombres"] == "No disponible" and resultado["ubicacion"] == "No disponible / No disponible / No disponible":
            raise ValueError("No se encontraron datos relevantes en la respuesta de la ONPE.")

    except requests.RequestException as exc:
        resultado["error"] = f"Error de red al consultar ONPE: {exc}"
    except Exception as exc:
        resultado["error"] = f"Error al procesar la consulta: {exc}"

    return resultado


def leer_dnis_excel(ruta_excel: str) -> list[str]:
    """Lee los DNIs desde dnis.xlsx y devuelve una lista de valores."""
    df = pd.read_excel(ruta_excel)
    if "dni" not in df.columns:
        raise ValueError("La columna 'dni' no existe en el archivo Excel.")
    return [str(valor).strip() for valor in df["dni"].tolist() if str(valor).strip()]


def main() -> None:
    ruta_excel = "dnis.xlsx"

    try:
        dnis = leer_dnis_excel(ruta_excel)
    except FileNotFoundError:
        print(f"No se encontró el archivo {ruta_excel}. Asegúrate de crearlo antes de ejecutar este script.")
        return
    except Exception as exc:
        print(f"Error al leer {ruta_excel}: {exc}")
        return

    resultados = [consultar_dni_onpe(dni) for dni in dnis]

    df_resultado = pd.DataFrame(
        [
            {
                "dni": item["dni"],
                "miembro_de_mesa": item["miembro_de_mesa"],
                "nombres": item["nombres"],
                "ubicación": item["ubicacion"],
                "dirección": item["direccion_local"],
            }
            for item in resultados
        ]
    )

    df_resultado.to_excel("resultado_consulta.xlsx", index=False, engine="openpyxl")
    print(f"Se guardaron {len(df_resultado)} registros en resultado_consulta.xlsx")

    for item in resultados:
        print("=" * 80)
        print(f"DNI: {item['dni']}")
        print(f"Miembro de mesa: {item['miembro_de_mesa']}")
        print(f"Nombres: {item['nombres']}")
        print(f"Ubicación: {item['ubicacion']}")
        print(f"Dirección del local: {item['direccion_local']}")
        if item.get("error"):
            print(f"Error: {item['error']}")


if __name__ == "__main__":
    main()
