# Convertidor de ING Excel a QIF 🏦➡️🧾

## Descripción

Este script de Python convierte los archivos de movimientos de cuenta descargados en formato Excel (`.xls` o `.xlsx`) desde la web de **ING España (ING BANK NV, Sucursal en España)** al formato **QIF (Quicken Interchange Format)**. El script extrae los detalles de la transacción y coloca el texto descriptivo principal (comercio, persona, etc.) en el campo "Memo" del QIF, dejando vacío el campo "Beneficiario".

## Motivación

ING España permite descargar los movimientos de cuenta en formato Excel, pero muchas aplicaciones populares de finanzas personales como [HomeBank](https://www.gethomebank.org), [KMyMoney](https://kmymoney.org/), [GnuCash](https://www.gnucash.org/) (con plugin QIF), o versiones antiguas de Quicken, funcionan mejor o únicamente con archivos QIF.

Este script automatiza el proceso de conversión, extrayendo la información relevante del Excel de ING y formateándola en un archivo QIF listo para importar, con el texto descriptivo clave en el campo "Memo" para facilitar la identificación y categorización posterior.

## ✨ Características principales

* **Lee formato Excel de ING:** procesa archivos `.xls` y `.xlsx` descargados de ING.
* **Conversión a QIF:** genera un archivo QIF estándar (`!Type:Bank`) listo para importar.
* **Extracción de texto descriptivo (para Memo):**
* Identifica y elimina prefijos comunes ("Pago en ", "Bizum recibido de ", "Transferencia...", etc.) de la descripción.
* Intenta extraer nombres de comercios o entidades que suelen estar en MAYÚSCULAS del texto restante.
* Si no encuentra un patrón en mayúsculas, utiliza el *resto de la descripción* (tras quitar el prefijo) como texto principal.
* Este texto extraído se coloca en el campo **Memo (`M`)** del archivo QIF.


* **Beneficiario QIF vacío:** el campo Beneficiario (`P`) del QIF se deja **intencionadamente vacío**.
* **Mapeo de categorías:** combina las columnas `CATEGORÍA` y `SUBCATEGORÍA` del Excel en el campo Categoría (`L`) del QIF, usando dos puntos (`:`) como separador jerárquico (ej: `LAlimentación:Supermercados y alimentación`).
* **Manejo de formatos españoles:** parsea correctamente importes con coma decimal y fechas en formato `DD/MM/YYYY`.
* **Validación de datos:**
* Comprueba que las columnas esenciales estén presentes.
* Valida que las fechas sean válidas y estén en un rango razonable.
* Valida que los importes sean numéricos, omitiendo filas con datos inválidos.


* **Codificación flexible:** permite elegir la codificación del archivo QIF de salida (`utf-8` por defecto, recomendado para compatibilidad con acentos).
* **Modo detallado (verbose):** incluye una opción `-v` para mostrar información detallada del procesamiento y depuración.
* **Modular:** el código está estructurado en funciones para facilitar su lectura y mantenimiento.

## ⚙️ Requisitos e instalación

1. **Python:** necesitas Python 3.6 o superior.
2. **Bibliotecas:** instala las dependencias necesarias usando pip:
```bash
pip install pandas xlrd openpyxl

```


* `pandas`: para leer archivos Excel.
* `xlrd`: necesario para leer archivos `.xls` antiguos.
* `openpyxl`: necesario para leer archivos `.xlsx` modernos.



## 🚀 Uso

El script se ejecuta desde la línea de comandos:

```bash
python ing2qif.py [opciones] <archivo_excel_entrada>

```

**Argumentos:**

* `archivo_excel_entrada`: ruta obligatoria a tu archivo Excel (`.xls` o `.xlsx`) descargado de ING.

**Opciones:**

* `-o ARCHIVO_SALIDA`, `--output ARCHIVO_SALIDA`: especifica la ruta y nombre del archivo QIF de salida. Por defecto, se crea un archivo con el mismo nombre que el de entrada pero con extensión `.qif`.
* `--encoding CODIFICACION`: especifica la codificación del archivo QIF de salida. Opciones: `utf-8` (recomendado y por defecto), `cp1252`, `iso-8859-1`.
* `-v`, `--verbose`: activa el modo detallado, mostrando mensajes de depuración durante el procesamiento.
* `-h`, `--help`: muestra la ayuda con todos los argumentos y opciones.

**Ejemplos:**

* **Conversión básica (salida por defecto `movimientos.qif`):**
```bash
python ing2qif.py movimientos.xlsx

```


* **Especificando archivo de salida:**
```bash
python ing2qif.py mis_movimientos.xls -o extracto_enero_2025.qif

```


* **Activando modo detallado:**
```bash
python ing2qif.py extracto_banco.xlsx -v

```



## 📄 Formato del archivo Excel de entrada (esperado)

El script está diseñado para funcionar con la estructura típica de los archivos Excel descargados desde la web de ING España. Espera encontrar:

1. Algunas filas iniciales con metadatos.
2. **Una fila de cabecera exacta** con los siguientes nombres de columna (buscada en las primeras 15 filas):
```
F. VALOR, CATEGORÍA, SUBCATEGORÍA, DESCRIPCIÓN, COMENTARIO, IMAGEN, IMPORTE (€), SALDO (€)

```


3. Las filas de datos de transacciones debajo de la cabecera.

**¡Importante!** Si ING cambia la estructura o los nombres de columna, el script podría necesitar ajustes.

## 🧾 Formato del archivo QIF de salida

El script genera un archivo QIF estándar (`!Type:Bank`). Los campos se mapean de la siguiente manera:

* `D`: fecha (formato `MM/DD/YYYY`).
* `T`: importe (con punto decimal).
* `P`: **(VACÍO)** - este campo se deja en blanco intencionadamente.
* `L`: categoría (formato `Categoría:Subcategoría` del Excel).
* `M`: memo/nota (contiene el texto descriptivo extraído de la descripción del Excel: comercio, persona, etc.).
* `^`: separador de transacción.

*(Nota: el comentario original de la columna `COMENTARIO` del Excel no se incluye en el QIF resultante).*

## 🔧 Configuración y personalización

Actualmente, la lógica principal (patrones de prefijo, regex de beneficiario, nombres de columna) está definida dentro del script.

* **Nombres de columna:** puedes intentar ajustar `COL_MAP` si ING cambia los nombres.
* **Prefijos:** los patrones se definen en `PREFIX_PATTERN`. Se usan solo para *limpiar* la descripción antes de extraer el texto para el campo Memo.
* **Lógica de extracción:** la función `extract_memo_text` contiene la lógica para determinar el texto que va al campo Memo.

Para personalizaciones más avanzadas, sería necesario modificar el código.

## ⚠️ Resolución de problemas conocidos

* **Error "Cabecera no encontrada" o "Faltan columnas":** verifica la estructura de tu archivo Excel y los nombres de columna contra los esperados.
* **Caracteres raros o incorrectos (acentos):** usa `--encoding utf-8` (opción por defecto).
* **Errores de lectura de Excel:** asegúrate de tener `pandas`, `xlrd` y `openpyxl` instalados.
* **Memo (`M`) inesperado:** usa el modo `-v` para ver cómo se extrae el texto descriptivo de la descripción original y se asigna al campo Memo. Recuerda que el Beneficiario (`P`) estará vacío.

## 🔮 Posibles mejoras futuras

* **Archivo de configuración externo:** para patrones de prefijo y mapeo de columnas.
* **Reglas de mapeo avanzadas:** para asignar categorías (`L`) o incluso un beneficiario (`P`) basado en reglas definidas por el usuario sobre el campo Memo (`M`).
* **Interfaz gráfica de usuario (GUI).**

## 🤝 Contribuciones

¡Las contribuciones son bienvenidas! Abre una incidencia (issue) o envía una solicitud de cambios (pull request) en GitHub.
