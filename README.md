# Caso 2: Consulta Electoral ONPE/RENIEC

Este proyecto genera un archivo Excel con DNIs de prueba, consulta información electoral por DNI usando la web de la ONPE y guarda resultados en un archivo Excel final.

## 1. Crear el archivo de entrada dnis.xlsx

Primero se genera el archivo Excel base con los DNIs de prueba:

```bash
python crear_excel.py
```

Esto crea el archivo `dnis.xlsx` con la columna `dni` y los siguientes valores:

- 10526358
- 15121313
- 15161414

## 2. Ejecutar la aplicación web localmente

Instala las dependencias del proyecto:

```bash
pip install -r requirements.txt
```

Luego ejecuta el servidor Flask:

```bash
python caso2.py
```

Abre `http://localhost:5000`. La aplicación permite:

- consultar un DNI individual desde el campo de texto
- subir un archivo `.xlsx` con una columna `dni` para una consulta masiva
- ver los resultados y descargar `resultado_consulta.xlsx`

La aplicación consulta la información en la web de la ONPE y extrae datos como `miembro_de_mesa`, `nombres`, `ubicación` y `dirección`.

## 3. Construir y ejecutar el contenedor con Docker

Se incluye un Dockerfile multietapa llamado `Dockerfile.caso2.multistage`.

### Construir la imagen

```bash
docker build -f Dockerfile.caso2.multistage -t caso2-onpe-playwright .
```

### Ejecutar el contenedor en Windows PowerShell

Para que el contenedor pueda leer y escribir los archivos del proyecto local, monta el directorio actual dentro de `/app`:

```powershell
docker run --rm -p 5001:5000 -v "${PWD}:/app" caso2-onpe-playwright
```

### Ejecutar el contenedor en Windows CMD

```cmd
docker run --rm -p 5001:5000 -v "%cd%:/app" caso2-onpe-playwright
```

Después abre `http://localhost:5001`. El contenedor ejecuta Flask con Chromium y una pantalla virtual para reproducir el flujo real de ONPE:

```bash
/app/start.sh
```

## 4. Salidas esperadas

Al ejecutar el flujo, se generan los siguientes archivos:

- `dnis.xlsx`: archivo de entrada opcional con los DNIs de prueba
- `resultado_consulta.xlsx`: resultados extraídos y procesados en un DataFrame; también está disponible desde el botón de descarga

## Notas

- La aplicación reproduce con Playwright el flujo oficial: abre ONPE, escribe el DNI, pulsa `Consultar` y extrae la pantalla de resultados.
- El contenedor usa Chromium y Xvfb, por lo que puede superar el reto que bloqueaba las peticiones directas de `requests`.
- El Excel tiene encabezados con estilo, filtros, congelado de la primera fila, ajuste de texto y anchos limitados para que las referencias largas no oculten el resto de columnas.
