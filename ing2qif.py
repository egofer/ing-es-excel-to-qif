# -*- coding: utf-8 -*-
import datetime
import re
import argparse
from decimal import Decimal, InvalidOperation
import pandas as pd
import sys

# --- Constantes ---
REASONABLE_START_DATE = datetime.datetime(1990, 1, 1)
REASONABLE_END_DATE = datetime.datetime.now() + datetime.timedelta(days=5*365)
EXPECTED_HEADER = ['F. VALOR', 'CATEGORÍA', 'SUBCATEGORÍA',
                   'DESCRIPCIÓN', 'COMENTARIO', 'IMAGEN', 'IMPORTE (€)', 'SALDO (€)']
COL_MAP = {
    'date': 'F. VALOR', 'category': 'CATEGORÍA', 'subcategory': 'SUBCATEGORÍA',
    'description': 'DESCRIPCIÓN', 'comment': 'COMENTARIO', 'amount': 'IMPORTE (€)'
}
REQUIRED_COLS_INTERNAL = ['date', 'description', 'amount']

# --- Compilación de Regex ---
PREFIX_PATTERN = re.compile(
    r"^(?:(Pago)\s+en\s+|(Bizum)\s+(?:recibido(?:\s+de)?|enviado(?:\s+a)?)\s+|(Transferencia)\s+(?:recibida(?:\s+de)?|internacional\s+emitida\s+[A-Z]\d+)\s+|(Devolución)\s+Tarjeta\s+)", re.VERBOSE | re.IGNORECASE)
# ALL_CAPS_PATTERN ahora solo intenta hacer match, la lógica decide si usarlo
ALL_CAPS_PATTERN = re.compile(
    r"^([A-ZÁÉÍÓÚÑ0-9.*\/&-]+(?=\s|$)(?:\s+(?=[A-ZÁÉÍÓÚÑ0-9.*\/&-]+(?:\s|$))[A-ZÁÉÍÓÚÑ0-9.*\/&-]+)*)", re.VERBOSE)

# --- Funciones ---


def parse_arguments():
    """Parsea los argumentos de la línea de comandos."""
    parser = argparse.ArgumentParser(
        description="Convierte extracto bancario Excel a QIF.",
        epilog="Ejemplo: python xls_to_qif.py extracto.xlsx -o salida.qif -v"
    )
    parser.add_argument("excel_file", help="Ruta al archivo Excel.")
    parser.add_argument(
        "-o", "--output", help="Ruta QIF salida (defecto: nombre.qif).")
    parser.add_argument("--encoding", default="utf-8", choices=[
                        "utf-8", "cp1252", "iso-8859-1"], help="Codificación salida QIF (defecto: utf-8).")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="Activar mensajes detallados.")
    return parser.parse_args()


def parse_spanish_decimal(decimal_val, row_num, verbose=False):
    """Convierte valor a Decimal, manejando formato español. Devuelve None si falla."""
    if pd.isna(decimal_val):
        if verbose:
            print(f"  [DEBUG] Fila {row_num}: Importe NaN/Vacío.")
        return None
    decimal_str = str(decimal_val)
    cleaned_str = decimal_str.replace(' ', '').replace('€', '')
    if ',' in cleaned_str:
        cleaned_str = cleaned_str.replace('.', '').replace(',', '.')
    try:
        return Decimal(cleaned_str)
    except InvalidOperation:
        if cleaned_str.lower() == 'nan':
            if verbose:
                print(f"  [DEBUG] Fila {row_num}: Importe NaN.")
            return None
        print(f"  AVISO: Fila {row_num}: No se pudo convertir importe '{
              decimal_str}' (limpio: '{cleaned_str}'). Omitiendo.")
        return None


def find_header_and_metadata(excel_filepath, expected_header, verbose=False):
    """Lee inicio del Excel para encontrar índice de cabecera y metadatos."""
    if verbose:
        print("Buscando cabecera y metadatos...")
    header_row_index = -1
    account_info = {}
    try:
        df_pre = pd.read_excel(excel_filepath, header=None,
                               keep_default_na=False, nrows=15)
    except Exception as e:
        print(f"Error Fatal leyendo inicio Excel: {e}")
        return -1, {}

    header_found_flag = False
    for idx, row_values in enumerate(df_pre.values.tolist()):
        row_str = [str(v).strip() for v in row_values]
        if not header_found_flag and row_str[:len(expected_header)] == expected_header:
            header_found_flag = True
            header_row_index = idx
            if verbose:
                print(f"Cabecera detectada en índice {
                      header_row_index} (Fila Excel {header_row_index + 1}).")
        if len(row_str) > 3:  # Extraer metadatos (simplificado)
            if "Número de cuenta:" in row_str[2]:
                account_info['account_number'] = row_str[3]
            elif "Titular:" in row_str[2]:
                account_info['holder_name'] = row_str[3]
            elif "Fecha exportación:" in row_str[2]:
                account_info['export_date_str'] = row_str[3]

    if not header_found_flag:
        print("Error Fatal: Cabecera no encontrada.")
        return -1, account_info
    if verbose:
        print(f"Metadatos: {account_info}")
    return header_row_index, account_info


def read_excel_data(excel_filepath, header_row_index, verbose=False):
    """Lee los datos principales del Excel usando el índice de cabecera."""
    if verbose:
        print(f"Leyendo datos con cabecera en índice {header_row_index}...")
    try:
        df_data = pd.read_excel(
            excel_filepath, header=header_row_index, keep_default_na=False)
        df_data.columns = df_data.columns.map(
            lambda x: x.strip() if isinstance(x, str) else x)
        if verbose:
            print(f"Columnas leídas: {df_data.columns.tolist()}")
        return df_data
    except Exception as e:
        print(f"Error Fatal leyendo datos Excel: {e}")
        return None


def extract_payee_and_keyword(description, prefix_pattern, all_caps_pattern, verbose=False):
    """Extrae beneficiario y keyword. FORZA FALLBACK si 'Todo Mayus' no consume todo."""
    payee = None
    keyword = None
    remaining_text = description

    prefix_match = prefix_pattern.match(description)
    if prefix_match:
        remaining_text = description[prefix_match.end():].strip()
        keyword = next(
            (g for g in prefix_match.groups() if g is not None), None)
        if verbose:
            print(f"  [DEBUG] Prefijo: '{prefix_match.group(0)}'.")
        if keyword:
            keyword = keyword.capitalize()
        if verbose and keyword:
            print(f"  [DEBUG]   -> Keyword: '{keyword}'")
        if verbose:
            print(f"  [DEBUG] Restante: '{remaining_text}'")
    elif verbose:
        print("  [DEBUG] Prefijo NO Detectado.")

    # Intentar extraer "Todo Mayúsculas"
    if remaining_text:
        name_match_caps = all_caps_pattern.match(remaining_text)
        if name_match_caps:
            matched_caps_text = name_match_caps.group(1).strip()
            # --- NUEVA COMPROBACIÓN ---
            # ¿El match consumió todo el texto restante?
            if name_match_caps.end() == len(remaining_text):
                # Sí, era un patrón "Todo Mayúsculas" genuino
                payee = matched_caps_text
                if verbose:
                    print(f"  [DEBUG] Match Caps (completo) ÉXITO: '{payee}'")
            else:
                # No, coincidió solo parcialmente (ej. "24"). Forzar fallback.
                if verbose:
                    print(f"  [DEBUG] Match Caps PARCIAL ('{
                          matched_caps_text}'). Forzando fallback.")
                payee = remaining_text  # Usar todo el texto restante
                if verbose:
                    print(
                        f"  [DEBUG] Fallback (por match parcial) asignado: '{payee}'")
        else:
            # El patrón "Todo Mayúsculas" no coincidió ni al principio
            if verbose:
                print("  [DEBUG] Match Caps FALLÓ completamente.")
            payee = remaining_text  # Fallback
            if verbose:
                print(f"  [DEBUG] Fallback (por no match) asignado: '{payee}'")
    elif verbose:
        print("  [DEBUG] Texto restante VACÍO.")

    # Limpieza final
    if payee:
        payee = re.sub(r'\s{2,}', ' ', payee).strip()
        if not payee:
            payee = None
    else:
        payee = None

    return payee, keyword


def process_transaction_row(row, original_excel_row, col_map, prefix_pattern, all_caps_pattern, verbose=False):
    """Procesa una fila del DataFrame y devuelve datos para QIF o None."""
    try:
        # --- Validación y Obtención Datos Críticos ---
        raw_date_val = row.get(col_map['date'])
        if pd.isna(raw_date_val) or str(raw_date_val).strip() == '':
            return None  # Omitir fila
        if isinstance(raw_date_val, datetime.datetime):
            tx_date = raw_date_val
        else:
            try:
                date_str = str(raw_date_val).split()[0]
                possible_formats = ["%d/%m/%Y",
                                    "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"]
                tx_date = next((datetime.datetime.strptime(date_str, fmt)
                               for fmt in possible_formats if datetime.datetime.strptime(date_str, fmt)), None)
                if tx_date is None:
                    raise ValueError("Formato fecha no reconocido")
            except (ValueError, TypeError) as e:
                print(f"  OMITIENDO fila {original_excel_row}: Fecha inválida '{
                      raw_date_val}'. {e}")
                return None

        if not (REASONABLE_START_DATE <= tx_date <= REASONABLE_END_DATE):
            print(f"  AVISO: Fila {original_excel_row}: Fecha '{
                  tx_date.strftime('%d/%m/%Y')}' fuera rango.")

        raw_amount_val = row.get(col_map['amount'])
        amount = parse_spanish_decimal(
            raw_amount_val, original_excel_row, verbose)
        if amount is None:
            return None

        # --- Obtener Otros Campos ---
        category_csv = str(row.get(col_map['category'], '')).strip()
        subcategory_csv = str(row.get(col_map['subcategory'], '')).strip()
        description = str(row.get(col_map['description'], '')).strip()
        comment_csv = str(row.get(col_map['comment'], '')).strip()

        # --- Extracción Beneficiario y Keyword ---
        payee_for_qif, tag_keyword = extract_payee_and_keyword(
            description, prefix_pattern, all_caps_pattern, verbose)
        if verbose:
            print(f"  [FINAL] Beneficiario (P): '{payee_for_qif}'")

        # --- Construcción Categoría y Memo ---
        category_parts = [part for part in [
            category_csv, subcategory_csv] if part]
        category_for_qif = ":".join(category_parts)
        if verbose:
            print(f"  [FINAL] Categoría (L): '{category_for_qif}'")
        memo_items = [item for item in [comment_csv, f"Tipo: {
            tag_keyword}" if tag_keyword else None] if item]
        qif_memo = " // ".join(memo_items)
        if verbose:
            print(f"  [FINAL] Memo (M): '{qif_memo}'")

        # --- Devolver datos ---
        return {'date': tx_date, 'amount': amount, 'payee': payee_for_qif, 'category': category_for_qif, 'memo': qif_memo}

    except Exception as e:
        print(
            f"*** Error INESPERADO procesando fila Excel {original_excel_row}: {e} ***")
        try:
            print(f"    Datos fila: {row.to_dict()}")
        except:
            pass
        return None


def generate_qif_file(transactions, qif_filepath, output_encoding, verbose=False):
    """Genera el archivo QIF a partir de la lista de transacciones procesadas."""
    print(f"\nGenerando QIF: {qif_filepath} (Codificación: {output_encoding})")
    write_errors_mode = 'replace' if output_encoding.lower() != 'utf-8' else 'strict'
    try:
        with open(qif_filepath, mode='w', encoding=output_encoding, errors=write_errors_mode) as outfile:
            outfile.write("!Type:Bank\n")
            for tx in transactions:
                outfile.write(f"D{tx['date'].strftime('%m/%d/%Y')}\n")
                outfile.write(f"T{tx['amount']:.2f}\n")
                if tx['payee']:
                    outfile.write(f"P{tx['payee']}\n")
                if tx['category']:
                    outfile.write(f"L{tx['category']}\n")
                if tx['memo']:
                    outfile.write(f"M{tx['memo']}\n")
                outfile.write("^\n")
        print(f"Archivo QIF creado: {qif_filepath}")
        # ... (Recordatorios / Avisos) ...
        return True
    except Exception as e:
        print(f"Error Fatal escribiendo QIF: {e}")
        return False


# --- Función Principal ---
def main():
    """Función principal que orquesta la conversión."""
    print("--- Recordatorio: Requiere 'pandas' y 'xlrd'/'openpyxl' ---")
    args = parse_arguments()
    # ... (Determinar output_filename igual que antes) ...
    if args.output:
        output_filename = args.output
    else:
        base_name = args.excel_file
        if base_name.lower().endswith(('.xls', '.xlsx')):
            output_filename = re.sub(
                r'\.[Xx][Ll][Ss][Xx]?$', '', base_name) + ".qif"
        else:
            output_filename = base_name + ".qif"

    header_idx, account_info = find_header_and_metadata(
        args.excel_file, EXPECTED_HEADER, args.verbose)
    if header_idx == -1:
        sys.exit(1)

    df_data = read_excel_data(args.excel_file, header_idx, args.verbose)
    if df_data is None:
        sys.exit(1)

    # ... (Validación de columnas requeridas igual que antes) ...
    current_columns = df_data.columns.tolist()
    missing_req_cols = [col_name for col_key, col_name in COL_MAP.items(
    ) if col_key in REQUIRED_COLS_INTERNAL and col_name not in current_columns]
    if missing_req_cols:
        print(f"Error Fatal: Faltan columnas: {missing_req_cols}")
        sys.exit(1)

    print("\nProcesando transacciones...")
    processed_transactions = []
    skipped_count = 0
    for idx, row in df_data.iterrows():
        original_excel_row = header_idx + 2 + idx
        processed_data = process_transaction_row(
            row, original_excel_row, COL_MAP, PREFIX_PATTERN, ALL_CAPS_PATTERN, args.verbose
        )
        if processed_data:
            processed_transactions.append(processed_data)
        else:
            skipped_count += 1

    print("-" * 30)
    print(f"Procesamiento completado.")
    processed_count = len(processed_transactions)
    print(f"  Transacciones procesadas: {processed_count}")
    print(f"  Filas omitidas: {skipped_count}")
    if not processed_transactions:
        print("\nError Fatal: No se procesaron transacciones.")
        sys.exit(1)
    print("-" * 30)

    processed_transactions.sort(key=lambda x: x['date'])
    success = generate_qif_file(
        processed_transactions, output_filename, args.encoding, args.verbose)

    print(
        f"\n--- Ejecución Finalizada {'con Éxito' if success else 'con Errores'} ---")
    if not success:
        sys.exit(1)


# --- Punto de Entrada ---
if __name__ == "__main__":
    main()
