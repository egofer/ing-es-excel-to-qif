# ING Excel to QIF Converter 🏦➡️🧾

## Descripción

Este script de Python convierte los archivos de movimientos de cuenta descargados en formato Excel (`.xls` o `.xlsx`) desde la web de **ING España (ING BANK NV, Sucursal en España)** al formato **QIF (Quicken Interchange Format)**. El script extrae los detalles de la transacción y coloca el texto descriptivo principal (comercio, persona, etc.) en el campo Memo del QIF, dejando vacío el campo Beneficiario.

## Motivación

ING España permite descargar los movimientos de cuenta en formato Excel, pero muchas aplicaciones populares de finanzas personales como [HomeBank](https://www.gethomebank.org), [KMyMoney](https://kmymoney.org/), [GnuCash](https://www.gnucash.org/) (con plugin QIF), o versiones antiguas de Quicken, funcionan mejor o únicamente con archivos QIF.

Este script automatiza el proceso de conversión, extrayendo la información relevante del Excel de ING y formateándola en un archivo QIF listo para importar, con el texto descriptivo clave en el campo Memo para facilitar la identificación y categorización posterior.

## ✨ Características principales

*   **Lee formato Excel ING:** Procesa archivos `.xls` y `.xlsx` descargados de ING.
*   **Conversión a QIF:** Genera un archivo QIF estándar (`!Type:Bank`) listo para importar.
*   **Extracción de texto descriptivo (para Memo):**
    *   Identifica y elimina prefijos comunes ("Pago en ", "Bizum recibido de ", "Transferencia...", etc.) de la descripción.
    *   Intenta extraer nombres de comercios o entidades que suelen estar en MAYÚSCULAS del texto restante.
    *   Si no encuentra un patrón en mayúsculas, utiliza el *resto de la descripción* (tras quitar el prefijo) como texto principal.
    *   Este texto extraído se coloca en el campo **Memo (`M`)** del archivo QIF.
*   **Beneficiario QIF Vacío:** El campo Beneficiario (`P`) del QIF se deja **intencionadamente vacío**.
*   **Mapeo de categorías:** Combina las columnas `CATEGORÍA` y `SUBCATEGORÍA` del Excel en el campo Categoría (`L`) del QIF, usando dos puntos (`:`) como separador jerárquico (ej: `LAlimentación:Supermercados y alimentación`).
*   **Manejo de formatos españoles:** Parsea correctamente importes con coma decimal y fechas en formato `DD/MM/YYYY`.
*   **Validación de Datos:**
    *   Comprueba que las columnas esenciales estén presentes.
    *   Valida que las fechas sean válidas y estén en un rango razonable.
    *   Valida que los importes sean numéricos, omitiendo filas con datos inválidos.
*   **Codificación flexible:** Permite elegir la codificación del archivo QIF de salida (`utf-8` por defecto, recomendado para compatibilidad con acentos).
*   **Modo Verbose:** Incluye una opción `-v` para mostrar información detallada del procesamiento y depuración.
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
python xls_to_qif.py [opciones] <archivo_excel_entrada>
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
    python xls_to_qif.py movimientos.xlsx
    ```
*   **Especificando archivo de salida:**
    ```bash
    python xls_to_qif.py mis_movimientos.xls -o extracto_enero_2025.qif
    ```
*   **Activando modo detallado:**
    ```bash
    python xls_to_qif.py extracto_banco.xlsx -v
    ```

## 📄 Formato del archivo Excel de entrada (esperado)

El script está diseñado para funcionar con la estructura típica de los archivos Excel descargados desde la web de ING España. Espera encontrar:

1.  Algunas filas iniciales con metadatos.
2.  **Una fila de cabecera EXACTA** con los siguientes nombres de columna (buscada en las primeras 15 filas):
    ```
    F. VALOR, CATEGORÍA, SUBCATEGORÍA, DESCRIPCIÓN, COMENTARIO, IMAGEN, IMPORTE (€), SALDO (€)
    ```
3.  Las filas de datos de transacciones debajo de la cabecera.

**¡Importante!** Si ING cambia la estructura o los nombres de columna, el script podría necesitar ajustes.

## 🧾 Formato del archivo QIF de salida

El script genera un archivo QIF estándar (`!Type:Bank`). Los campos se mapean de la siguiente manera:

*   `D`: Fecha (Formato `MM/DD/YYYY`)
*   `T`: Importe (con punto decimal)
*   `P`: **(VACÍO)** - Este campo se deja en blanco intencionadamente.
*   `L`: Categoría (Formato `Categoría:Subcategoría` del Excel)
*   `M`: Memo/Nota (Contiene el texto descriptivo extraído de la descripción del Excel: comercio, persona, etc.)
*   `^`: Separador de transacción

*(Nota: El comentario original de la columna `COMENTARIO` del Excel no se incluye en el QIF resultante).*

## 🔧 Configuración y personalización

Actualmente, la lógica principal (patrones de prefijo, regex de beneficiario, nombres de columna) está definida dentro del script.

*   **Nombres de Columna:** Puedes intentar ajustar `COL_MAP` si ING cambia los nombres.
*   **Prefijos:** Los patrones se definen en `PREFIX_PATTERN`. Se usan solo para *limpiar* la descripción antes de extraer el texto para el Memo.
*   **Lógica de Extracción:** La función `extract_memo_text` contiene la lógica para determinar el texto que va al campo Memo.

Para personalizaciones más avanzadas, sería necesario modificar el código.

## ⚠️ Troubleshooting y problemas conocidos

*   **Error "Cabecera no encontrada" / "Faltan columnas":** Verifica la estructura de tu archivo Excel y los nombres de columna contra los esperados.
*   **Caracteres Raros/Incorrectos (Acentos):** Usa `--encoding utf-8` (opción por defecto).
*   **Errores de Lectura Excel:** Asegúrate de tener `pandas`, `xlrd`, `openpyxl` instalados.
*   **Memo (`M`) Inesperado:** Usa el modo `-v` para ver cómo se extrae el texto descriptivo de la descripción original y se asigna al campo Memo. Recuerda que el Beneficiario (`P`) estará vacío.

## 🔮 Posibles mejoras futuras

*   **Archivo de configuración externo:** Para patrones de prefijo, mapeo de columnas.
*   **Reglas de Mapeo Avanzadas:** Para asignar Categorías (`L`) o incluso un Beneficiario (`P`) basado en reglas definidas por el usuario sobre el Memo (`M`).
*   **Interfaz Gráfica (GUI).**

## 🤝 Contribuciones

¡Las contribuciones son bienvenidas! Abre un "Issue" o envía un "Pull Request" en GitHub.
