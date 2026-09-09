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

## 2. Ejecutar el scraper localmente

Instala las dependencias del proyecto:

```bash
pip install -r requirements.txt
```

Luego ejecuta el script principal:

```bash
python caso2.py
```

El script:

- lee los DNIs desde `dnis.xlsx`
- consulta la información en la web de la ONPE
- extrae datos como `miembro_de_mesa`, `nombres`, `ubicación` y `dirección`
- guarda el resultado final en `resultado_consulta.xlsx`

## 3. Construir y ejecutar el contenedor con Docker

Se incluye un Dockerfile multietapa llamado `Dockerfile.caso2.multistage`.

### Construir la imagen

```bash
docker build -f Dockerfile.caso2.multistage -t caso2-onpe .
```

### Ejecutar el contenedor

Para que el contenedor pueda acceder a `dnis.xlsx` del proyecto local, monta el directorio actual dentro del contenedor:

```bash
docker run --rm -v "${PWD}:/app" caso2-onpe
```

En Windows PowerShell, también puede usarse:

```powershell
docker run --rm -v "${PWD}:/app" caso2-onpe
```

Esto ejecuta el comando por defecto del contenedor:

```bash
python caso2.py
```

## 4. Salidas esperadas

Al ejecutar el flujo, se generan los siguientes archivos:

- `dnis.xlsx`: archivo base con los DNIs de prueba
- `resultado_consulta.xlsx`: resultados extraídos y procesados en un DataFrame

## Notas

- El scraper maneja errores de consulta y evita que la ejecución falle completamente en caso de que la ONPE responda con un estado inesperado.
- Si el sitio web cambia su estructura HTML, puede requerirse ajustar las reglas de extracción en `caso2.py`.
