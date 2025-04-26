# ING ES Excel to QIF Converter 🏦➡️🧾

## Descripción

Este script de Python convierte los archivos de movimientos de cuenta descargados en formato Excel (`.xls` o `.xlsx`) desde la web de **ING España (ING BANK NV, Sucursal en España)** al formato **QIF (Quicken Interchange Format)**.

## Motivación

ING España permite descargar los movimientos de cuenta en formato Excel, pero muchas aplicaciones populares de finanzas personales como [KMyMoney](https://kmymoney.org/), [GnuCash](https://www.gnucash.org/), [HomeBank](https://www.gethomebank.org), o versiones antiguas de Quicken, funcionan mejor o únicamente con archivos QIF.

Este script automatiza el proceso de conversión, extrayendo la información relevante del Excel de ING Direct y formateándola correctamente en un archivo QIF, ahorrando tiempo y esfuerzo manual. Está especialmente optimizado para la estructura y los patrones de descripción encontrados comúnmente en los extractos de ING España.

## ✨ Características principales

*   **Lee formato Excel ING:** Procesa archivos `.xls` y `.xlsx` descargados de ING.
*   **Conversión a QIF:** Genera un archivo QIF estándar (`!Type:Bank`) listo para importar.
*   **Extracción inteligente de beneficiario (Payee):**
    *   Identifica y elimina prefijos comunes ("Pago en ", "Bizum recibido de ", "Transferencia...", etc.).
    *   Intenta extraer nombres de comercios o entidades que suelen estar en MAYÚSCULAS.
    *   Si no encuentra un patrón en mayúsculas, utiliza el resto de la descripción como beneficiario (útil para nombres propios o descripciones complejas como "24 ( VEINTE Y CUATRO) ALICANTE ES").
*   **Mapeo de Categorías:** Combina las columnas `CATEGORÍA` y `SUBCATEGORÍA` del Excel en el campo Categoría (`L`) del QIF, usando dos puntos (`:`) como separador jerárquico (ej: `LAlimentación:Supermercados y alimentación`).
*   **Memo detallado:**
    *   Utiliza la columna `COMENTARIO` del Excel como parte del Memo (`M`) del QIF.
    *   Identifica el tipo de transacción por el prefijo (Pago, Bizum, Transferencia, Devolución) y lo añade al Memo como `Tipo: [Keyword]` (ej. `MTipo: Bizum`).
*   **Manejo de formatos españoles:** Parsea correctamente importes con coma decimal y fechas en formato `DD/MM/YYYY`.
*   **Validación de datos:**
    *   Comprueba que las columnas esenciales estén presentes.
    *   Valida que las fechas sean válidas y estén en un rango razonable.
    *   Valida que los importes sean numéricos, omitiendo filas con datos inválidos.
*   **Codificación Flexible:** Permite elegir la codificación del archivo QIF de salida (`utf-8` por defecto, recomendado para compatibilidad con acentos).
*   **Modo verboso:** Incluye una opción `-v` para mostrar información detallada del procesamiento y depuración.
*   **Modular:** El código está estructurado en funciones para facilitar su lectura y mantenimiento.

## ⚙️ Requisitos e instalación

1.  **Python:** Necesitas Python 3.6 o superior.
2.  **Bibliotecas:** Instala las dependencias necesarias usando pip:
    ```bash
    pip install pandas xlrd openpyxl
    ```
    *   `pandas`: Para leer archivos Excel.
    *   `xlrd`: Necesario para leer archivos `.xls` antiguos.
    *   `openpyxl`: Necesario para leer archivos `.xlsx` modernos.

## 🚀 Uso

El script se ejecuta desde la línea de comandos:

```bash
python ingxls2qif.py [opciones] <archivo_excel_entrada>
```

**Argumentos:**

*   `archivo_excel_entrada`: Ruta obligatoria a tu archivo Excel (`.xls` o `.xlsx`) descargado de ING.

**Opciones:**

*   `-o ARCHIVO_SALIDA`, `--output ARCHIVO_SALIDA`: Especifica la ruta y nombre del archivo QIF de salida. Por defecto, se crea un archivo con el mismo nombre que el de entrada pero con extensión `.qif`.
*   `--encoding CODIFICACION`: Especifica la codificación del archivo QIF de salida. Opciones: `utf-8` (recomendado y por defecto), `cp1252`, `iso-8859-1`.
*   `-v`, `--verbose`: Activa el modo detallado, mostrando mensajes de depuración durante el procesamiento.
*   `-h`, `--help`: Muestra la ayuda con todos los argumentos y opciones.

**Ejemplos:**

*   **Conversión básica (salida por defecto `movimientos.qif`):**
    ```bash
    python ingxls2qif.py movimientos.xls
    ```
*   **Especificando archivo de salida:**
    ```bash
    python ingxls2qif.py mis_movimientos.xls -o extracto_enero_2025.qif
    ```
*   **Activando modo detallado:**
    ```bash
    python ingxls2qif.py extracto_banco.xls -v
    ```
*   **Usando codificación diferente (menos común):**
    ```bash
    python ingxls2qif.py extracto_banco.xlsx --encoding cp1252
    ```

## 📄 Formato del archivo Excel de entrada (esperado)

El script está diseñado para funcionar con la estructura típica de los archivos Excel descargados desde la web de ING España. Espera encontrar:

1.  Algunas filas iniciales con metadatos (Número de cuenta, Titular, Fecha exportación). El script intenta leer esta información pero no es crítica para la conversión.
2.  **Una fila de cabecera EXACTA** con los siguientes nombres de columna (el script la busca en las primeras 15 filas):
    ```
    F. VALOR, CATEGORÍA, SUBCATEGORÍA, DESCRIPCIÓN, COMENTARIO, IMAGEN, IMPORTE (€), SALDO (€)
    ```
3.  Las filas de datos de transacciones debajo de la cabecera.

**¡Importante!** Si ING cambia significativamente la estructura o los nombres de las columnas en sus exportaciones futuras, el script podría necesitar ajustes.

## 🧾 Formato del archivo QIF de salida

El script genera un archivo QIF estándar (`!Type:Bank`) que debería ser compatible con la mayoría de software que soporta este formato. Los campos se mapean de la siguiente manera:

*   `D`: Fecha (Formato `MM/DD/YYYY`)
*   `T`: Importe (con punto decimal)
*   `P`: Beneficiario/Pagador (Extraído de la descripción)
*   `L`: Categoría (Formato `Categoría:Subcategoría` del Excel)
*   `M`: Memo/Nota (Contiene el `COMENTARIO` del Excel y/o `Tipo: [Keyword]`)
*   `^`: Separador de transacción

## 🔧 Configuración y personalización

Actualmente, la lógica principal (patrones de prefijo, regex de beneficiario, nombres de columna esperados) está definida dentro del script Python.

*   **Nombres de Columna:** Si ING cambia los nombres de columna, puedes intentar ajustar el diccionario `COL_MAP` al principio del script.
*   **Prefijos:** Los patrones de prefijo se definen en la variable `PREFIX_PATTERN`. Puedes añadir o modificar patrones Regex aquí si encuentras nuevos tipos de transacción recurrentes.
*   **Lógica de Extracción:** La función `extract_payee_and_keyword` contiene la lógica para determinar el beneficiario.

Para personalizaciones más avanzadas, sería necesario modificar el código Python.

## ⚠️ Troubleshooting y problemas conocidos

*   **Error "Cabecera no encontrada":** Asegúrate de que tu archivo Excel contiene la fila de cabecera exacta mencionada arriba y que está dentro de las primeras 15 filas. Verifica que no haya filas completamente vacías antes de la cabecera que puedan confundir a `pandas`.
*   **Error "Faltan columnas requeridas":** Verifica que las columnas `F. VALOR`, `DESCRIPCIÓN`, e `IMPORTE (€)` existen en tu archivo Excel después de la fila de cabecera.
*   **Caracteres Raros/Incorrectos (Acentos):** Si ves símbolos extraños en lugar de acentos o 'ñ' en el archivo QIF importado, asegúrate de que estás usando la codificación correcta. Prueba generando el archivo con la opción por defecto (`--encoding utf-8`). Si sigues teniendo problemas, podrías probar con `cp1252` o `iso-8859-1`, aunque `utf-8` es lo más recomendable. Se ha comprobado que hay casos en que el error de codificación se arrastra de los propios datos proporcionados por el banco.
*   **Errores de Lectura de Excel:** Asegúrate de tener instaladas las bibliotecas `pandas`, `xlrd` y `openpyxl` (`pip install pandas xlrd openpyxl`). Si el archivo está protegido o corrupto, pandas no podrá leerlo.
*   **Beneficiario Incorrecto:** Si el beneficiario extraído no es el esperado, revisa la descripción original y la lógica en `extract_payee_and_keyword`. Puedes usar el modo `-v` para ver cómo se procesa cada descripción.

## 🔮 Posibles mejoras futuras

*   **Archivo de configuración externo:** Mover los patrones de prefijo, mapeo de columnas y otras configuraciones a un archivo externo (JSON, YAML) para facilitar la personalización sin editar el script.
*   **Reglas de mapeo avanzadas:** Implementar un sistema de reglas (quizás en el archivo de configuración) para mapear beneficiarios o descripciones específicas a categorías o beneficiarios QIF deseados por el usuario.
*   **Interfaz gráfica (GUI):** Crear una interfaz simple para seleccionar archivos y opciones sin usar la línea de comandos.

## 🤝 Contribuciones

¡Las contribuciones son bienvenidas! Si encuentras errores, tienes sugerencias de mejora o quieres añadir nuevas funcionalidades, por favor, abre un "Issue" o envía un "Pull Request" en GitHub.

## 📜 Licencia

Este proyecto se distribuye bajo la Licencia MIT.

```text
MIT License

Copyright (c) [Año] [Tu Nombre o Nombre del Repositorio]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```
