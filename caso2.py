import re
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

from flask import Flask, render_template_string, request, send_file
import pandas as pd
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright
import requests
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_URL = "https://consultaelectoral.onpe.gob.pe/inicio"
OUTPUT_FILE = Path("resultado_consulta.xlsx")
app = Flask(__name__)

HTML = """
<!doctype html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Consulta electoral ONPE</title>
    <style>
        :root { color-scheme: light; font-family: system-ui, sans-serif; }
        body { margin: 0; background: #eef2f7; color: #172033; }
        main { max-width: 920px; margin: 2rem auto; padding: 0 1rem; }
        section { background: white; border-radius: 12px; padding: 1.5rem; margin-bottom: 1rem;
                  box-shadow: 0 4px 18px #17203318; }
        h1 { margin-top: 0; }
        h2 { font-size: 1.1rem; margin-top: 0; }
        form { display: flex; flex-wrap: wrap; gap: .75rem; align-items: center; }
        input[type=text], input[type=file] { flex: 1 1 240px; padding: .7rem; border: 1px solid #b9c3d0; border-radius: 7px; }
        button, .download { border: 0; border-radius: 7px; padding: .75rem 1rem; background: #1261a0; color: white;
                            font-weight: 700; text-decoration: none; cursor: pointer; }
        button:hover, .download:hover { background: #0d4d80; }
        .notice { padding: .8rem; border-radius: 7px; background: #fff4d6; }
        .error { background: #ffe1e1; color: #8b1e1e; }
        .success { background: #e2f6e9; color: #155c2b; }
        .table-wrap { overflow-x: auto; }
        table { border-collapse: collapse; width: 100%; }
        th, td { text-align: left; padding: .65rem; border-bottom: 1px solid #e1e6ec; }
        th { background: #f4f7fa; }
        td { min-width: 120px; vertical-align: top; white-space: pre-wrap; }
        .muted { color: #5e6b7d; }
    </style>
</head>
<body>
<main>
    <section>
        <h1>Consulta electoral por DNI</h1>
        <p>Consulta un documento individual o procesa un Excel con una columna llamada <strong>dni</strong>.</p>
        <form method="post" action="/consultar">
            <input name="dni" type="text" inputmode="numeric" placeholder="DNI individual (8 dígitos)">
            <button type="submit">Consultar DNI</button>
        </form>
    </section>
    <section>
        <h2>Consulta masiva</h2>
        <form method="post" action="/consultar" enctype="multipart/form-data">
            <input name="archivo" type="file" accept=".xlsx" required>
            <button type="submit">Procesar Excel</button>
        </form>
    </section>
    {% if mensaje %}<p class="notice {{ 'error' if error else 'success' }}">{{ mensaje }}
        {% if enlace_oficial %}<br><a href="{{ enlace_oficial }}" target="_blank" rel="noopener">Abrir consulta oficial de ONPE</a>{% endif %}
    </p>{% endif %}
    {% if resultados %}
    <section>
        <h2>Resultados ({{ resultados|length }})</h2>
        <p><a class="download" href="/descargar">Descargar resultado_consulta.xlsx</a></p>
        <div class="table-wrap"><table>
            <thead><tr><th>DNI</th><th>Estado</th><th>Nombres y apellidos</th><th>Región / Provincia / Distrito</th><th>Local de votación</th><th>Dirección</th><th>Referencia</th><th>Mesa</th><th>Orden</th><th>Observación</th></tr></thead>
            <tbody>{% for item in resultados %}
                <tr><td>{{ item.dni }}</td><td>{{ item.miembro_de_mesa }}</td><td>{{ item.nombres }}</td>
                    <td>{{ item.ubicacion }}</td><td>{{ item.local_votacion }}</td><td>{{ item.direccion_local }}</td>
                    <td>{{ item.referencia }}</td><td>{{ item.mesa }}</td><td>{{ item.orden }}</td><td>{{ item.error or 'Consulta correcta' }}</td></tr>
            {% endfor %}</tbody>
        </table></div>
    </section>
    {% endif %}
</main>
</body>
</html>
"""


def normalizar_texto(value: Any) -> str:
    if value is None:
        return "No disponible"
    texto = str(value).strip()
    return re.sub(r"\s+", " ", texto) if texto else "No disponible"


def consultar_dni_onpe(dni: str) -> dict[str, Any]:
    """Reproduce en Chromium el flujo visible de consulta de ONPE."""
    resultado = {
        "dni": dni,
        "miembro_de_mesa": "No disponible",
        "nombres": "No disponible",
        "ubicacion": "No disponible",
        "local_votacion": "No disponible",
        "direccion_local": "No disponible",
        "referencia": "No disponible",
        "mesa": "No disponible",
        "orden": "No disponible",
        "error": None,
    }

    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=False)
            page = browser.new_page(locale="es-PE")
            page.goto(BASE_URL, wait_until="domcontentloaded", timeout=30_000)
            page.get_by_role("textbox", name="Número de DNI").fill(dni)
            page.get_by_role("button", name="Consultar").click()
            page.wait_for_url("**/main/local-de-votacion", timeout=30_000)
            page.wait_for_timeout(500)
            lineas = [line.strip() for line in page.locator("body").inner_text().splitlines() if line.strip()]
            browser.close()

        def siguiente(texto: str) -> str:
            indice = lineas.index(texto)
            return lineas[indice + 1] if indice + 1 < len(lineas) else "No disponible"

        resultado["miembro_de_mesa"] = next(
            (linea for linea in lineas if "MIEMBRO DE MESA" in linea.upper()), "No disponible"
        )
        resultado["nombres"] = siguiente("Nombres y Apellidos")
        resultado["ubicacion"] = siguiente("Región / Provincia / Distrito")
        resultado["mesa"] = siguiente("N° de Mesa:")
        resultado["orden"] = siguiente("N° de Orden:")

        local_index = lineas.index("Tu local de votación")
        datos_locales = lineas[local_index + 1:]
        datos_locales = [linea for linea in datos_locales if linea not in {"ver", "Mapa"}]
        resultado["local_votacion"] = datos_locales[0] if datos_locales else "No disponible"
        resultado["direccion_local"] = datos_locales[1] if len(datos_locales) > 1 else "No disponible"
        referencia = next(
            (linea for linea in datos_locales[2:] if linea.lower().startswith("referencia:")),
            "No disponible",
        )
        resultado["referencia"] = referencia.removeprefix("Referencia:").strip()

    except PlaywrightTimeoutError:
        resultado["error"] = "ONPE no completó la consulta en el tiempo esperado."
    except Exception as exc:
        resultado["error"] = f"Error al procesar la consulta: {exc}"

    return resultado


def leer_dnis_excel(ruta_excel: str | BytesIO) -> list[str]:
    """Lee los DNIs desde un Excel local o subido y devuelve una lista de valores."""
    df = pd.read_excel(ruta_excel) if isinstance(ruta_excel, str) else pd.read_excel(ruta_excel)
    if "dni" not in df.columns:
        raise ValueError("La columna 'dni' no existe en el archivo Excel.")
    return [str(valor).strip() for valor in df["dni"].tolist() if str(valor).strip()]


def crear_excel_resultados(resultados: list[dict[str, Any]]) -> None:
    """Guarda un Excel legible, con columnas completas y formato para revisión."""
    columnas = [
        "DNI", "Estado / cargo", "Nombres y apellidos", "Región / Provincia / Distrito",
        "Local de votación", "Dirección", "Referencia", "N° de mesa", "N° de orden", "Observación",
    ]
    df_resultado = pd.DataFrame(
        [
            {
                columnas[0]: item["dni"],
                columnas[1]: item["miembro_de_mesa"],
                columnas[2]: item["nombres"],
                columnas[3]: item["ubicacion"],
                columnas[4]: item["local_votacion"],
                columnas[5]: item["direccion_local"],
                columnas[6]: item["referencia"],
                columnas[7]: item["mesa"],
                columnas[8]: item["orden"],
                columnas[9]: item.get("error") or "Consulta correcta",
            }
            for item in resultados
        ]
    )
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df_resultado.to_excel(writer, index=False, sheet_name="Resultados")
        worksheet = writer.sheets["Resultados"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.row_dimensions[1].height = 32
        header_fill = PatternFill("solid", fgColor="0B4F71")
        for cell in worksheet[1]:
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        for index, column in enumerate(df_resultado.columns, start=1):
            longest = max(len(str(column)), *(len(str(value)) for value in df_resultado.iloc[:, index - 1]))
            worksheet.column_dimensions[get_column_letter(index)].width = min(max(longest + 2, 14), 42)


@app.route("/", methods=["GET"])
def inicio():
    return render_template_string(HTML)


@app.route("/consultar", methods=["POST"])
def consultar():
    archivo = request.files.get("archivo")
    dni = request.form.get("dni", "").strip()

    try:
        if archivo and archivo.filename:
            if not archivo.filename.lower().endswith(".xlsx"):
                raise ValueError("El archivo debe tener extensión .xlsx.")
            dnis = leer_dnis_excel(BytesIO(archivo.read()))
        elif dni:
            dnis = [dni]
        else:
            raise ValueError("Indica un DNI o selecciona un archivo Excel.")

        if not dnis:
            raise ValueError("El archivo no contiene DNIs para consultar.")
        resultados = [consultar_dni_onpe(valor) for valor in dnis]
        correctos = [item for item in resultados if not item.get("error")]
        errores = len(resultados) - len(correctos)
        if not correctos:
            return render_template_string(
                HTML,
                mensaje=(
                    "ONPE no permitió ninguna consulta automática. "
                    "El sitio está protegido por AWS WAF; realiza la consulta desde su página oficial."
                ),
                error=True,
                enlace_oficial=BASE_URL,
            ), 502
        crear_excel_resultados(resultados)
        mensaje = f"Se procesaron {len(correctos)} DNI(s) correctamente."
        if errores:
            mensaje += f" {errores} consulta(s) fueron bloqueadas o devolvieron error."
        return render_template_string(
            HTML, resultados=resultados, mensaje=mensaje
        )
    except Exception as exc:
        return render_template_string(HTML, mensaje=str(exc), error=True), 400


@app.route("/descargar", methods=["GET"])
def descargar():
    if not OUTPUT_FILE.exists():
        return render_template_string(HTML, mensaje="Todavía no hay un resultado para descargar.", error=True), 404
    return send_file(OUTPUT_FILE, as_attachment=True, download_name=OUTPUT_FILE.name)


def main() -> None:
    app.run(host="0.0.0.0", port=5000)


if __name__ == "__main__":
    main()
