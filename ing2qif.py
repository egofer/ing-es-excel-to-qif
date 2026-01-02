#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Converts bank statements downloaded as Excel files (.xls/.xlsx) from
ING Spain (ING BANK NV, Sucursal en España) into the QIF
(Quicken Interchange Format).

This script parses the specific structure of ING's Excel export,
extracts transaction details (date, amount, description, category, comment),
attempts to intelligently determine the Payee (beneficiary) from the
description (using 'all caps' detection or fallback to remaining text),
maps CATEGORÍA and SUBCATEGORÍA to the QIF category field (L),
and includes the original comment plus a transaction type keyword (e.g., 'Tipo: Bizum')
extracted from common description prefixes in the QIF memo field (M).

Includes data validation (date range, valid amount) and a verbose mode
for debugging.

Requires pandas, xlrd, and openpyxl. Install with:
pip install pandas xlrd openpyxl
"""

__author__ = "https://github.com/egofer"
__version__ = "0.3.1"
__status__ = "Development"
__date__ = "2025-12-02"
__license__ = "MIT"

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
                   'DESCRIPCIÓN', 'COMENTARIO', 'IMPORTE (€)', 'SALDO (€)']

COL_MAP = {
    'date': 'F. VALOR', 'category': 'CATEGORÍA', 'subcategory': 'SUBCATEGORÍA',
    'description': 'DESCRIPCIÓN', 'comment': 'COMENTARIO', 'amount': 'IMPORTE (€)'
}
REQUIRED_COLS_INTERNAL = ['date', 'description', 'amount']

# --- Compilación de Regex ---
PREFIX_PATTERN = re.compile(
    r"^(?:Pago\s+en\s+|Bizum\s+(?:recibido(?:\s+de)?|enviado(?:\s+a)?)\s+|Transferencia\s+(?:recibida(?:\s+de)?|internacional\s+emitida\s+[A-Z]\d+)\s+|Devolución\s+Tarjeta\s+)", re.VERBOSE | re.IGNORECASE)
ALL_CAPS_PATTERN = re.compile(
    r"^([A-ZÁÉÍÓÚÑ0-9.*\/&-]+(?=\s|$)(?:\s+(?=[A-ZÁÉÍÓÚÑ0-9.*\/&-]+(?:\s|$))[A-ZÁÉÍÓÚÑ0-9.*\/&-]+)*)", re.VERBOSE)

# --- Funciones ---


def parse_arguments():
    """Parsea los argumentos de la línea de comandos."""
    parser = argparse.ArgumentParser(
        description="Convierte extracto Excel ING a QIF")
    parser.add_argument("excel_file", help="Ruta al archivo Excel.")
    parser.add_argument(
        "-o", "--output", help="Ruta QIF salida (defecto: nombre.qif).")
    parser.add_argument("--encoding", default="utf-8",
                        choices=["utf-8", "cp1252", "iso-8859-1"], help="Codificación salida.")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="Activar mensajes detallados.")
    return parser.parse_args()


def parse_spanish_decimal(decimal_val, row_num, verbose=False):
    """Convierte valor a Decimal, manejando formato español."""
    if pd.isna(decimal_val):
        return None
    decimal_str = str(decimal_val)
    # Limpieza básica
    cleaned_str = decimal_str.replace(' ', '').replace('€', '')
    # Si hay coma, asumimos formato europeo (1.000,00) -> (1000.00)
    if ',' in cleaned_str:
        cleaned_str = cleaned_str.replace('.', '').replace(',', '.')
    try:
        return Decimal(cleaned_str)
    except InvalidOperation:
        if verbose:
            print(f"  [DEBUG] Fila {
                  row_num}: No se pudo convertir importe '{decimal_str}'.")
        return None


def find_header_and_metadata(excel_filepath, expected_header, verbose=False):
    """Lee inicio del Excel para encontrar índice de cabecera."""
    if verbose:
        print("Buscando cabecera y metadatos...")
    header_row_index = -1
    account_info = {}
    try:
        # Leemos las primeras filas sin cabecera para buscar la estructura
        df_pre = pd.read_excel(excel_filepath, header=None,
                               keep_default_na=False, nrows=20)
    except Exception as e:
        print(f"Error Fatal leyendo inicio Excel: {e}")
        return -1, {}

    header_found_flag = False
    for idx, row_values in enumerate(df_pre.values.tolist()):
        # Convertimos todo a string y quitamos espacios para comparar
        row_str = [str(v).strip() for v in row_values]

        # Comparamos solo las columnas necesarias (hasta la longitud del esperado)
        current_signature = row_str[:len(expected_header)]

        if verbose and not header_found_flag:
            pass

        if not header_found_flag and current_signature == expected_header:
            header_found_flag = True
            header_row_index = idx
            if verbose:
                print(f"Cabecera detectada en índice {
                      header_row_index} (Fila Excel {header_row_index + 1}).")

    if not header_found_flag:
        print("Error Fatal: Cabecera no encontrada.")
        print(f"  Se esperaba: {expected_header}")
        return -1, account_info
    return header_row_index, account_info


def read_excel_data(excel_filepath, header_row_index, verbose=False):
    """Lee los datos principales del Excel."""
    try:
        # keep_default_na=False hace que las celdas vacías sean strings vacíos ''
        df_data = pd.read_excel(
            excel_filepath, header=header_row_index, keep_default_na=False)
        df_data.columns = df_data.columns.map(
            lambda x: x.strip() if isinstance(x, str) else x)
        return df_data
    except Exception as e:
        print(f"Error Fatal leyendo datos Excel: {e}")
        return None


def extract_memo_text(description, prefix_pattern, all_caps_pattern, verbose=False):
    """Extrae texto para el Memo."""
    memo_text = None
    remaining_text = description
    prefix_match = prefix_pattern.match(description)
    if prefix_match:
        remaining_text = description[prefix_match.end():].strip()

    if remaining_text:
        name_match_caps = all_caps_pattern.match(remaining_text)
        if name_match_caps:
            if name_match_caps.end() == len(remaining_text):
                memo_text = name_match_caps.group(1).strip()
            else:
                memo_text = remaining_text
        else:
            memo_text = remaining_text

    if memo_text:
        memo_text = re.sub(r'\s{2,}', ' ', memo_text).strip()
        if not memo_text:
            memo_text = None
    return memo_text


def get_excel_date(raw_val):
    """
    Convierte número de serie Excel (float/int) a datetime.
    Excel base date: 30-dic-1899.
    """
    try:
        # Si ya es datetime, devolver
        if isinstance(raw_val, datetime.datetime):
            return raw_val

        # Si es un número (ej. 45992.0), convertir desde fecha base Excel
        if isinstance(raw_val, (int, float)):
            return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=raw_val)

        # Si es string, intentar parsear
        date_str = str(raw_val).split()[0].strip()
        possible_formats = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"]
        for fmt in possible_formats:
            try:
                return datetime.datetime.strptime(date_str, fmt)
            except ValueError:
                continue
    except Exception:
        pass
    return None


def process_transaction_row(row, original_excel_row, col_map, prefix_pattern, all_caps_pattern, verbose=False):
    """Procesa una fila."""
    try:
        # --- Validación Fecha ---
        raw_date_val = row.get(col_map['date'])
        if pd.isna(raw_date_val) or str(raw_date_val).strip() == '':
            return None

        tx_date = get_excel_date(raw_date_val)

        if tx_date is None:
            if verbose:
                print(f"  OMITIENDO fila {
                      original_excel_row}: Fecha inválida '{raw_date_val}'.")
            return None

        if not (REASONABLE_START_DATE <= tx_date <= REASONABLE_END_DATE):
            if verbose:
                print(f"  AVISO: Fila {original_excel_row}: Fecha '{
                      tx_date}' fuera rango razonable.")

        # --- Importe ---
        raw_amount_val = row.get(col_map['amount'])
        amount = parse_spanish_decimal(
            raw_amount_val, original_excel_row, verbose)
        if amount is None:
            return None

        # --- Otros Campos ---
        category_csv = str(row.get(col_map['category'], '')).strip()
        subcategory_csv = str(row.get(col_map['subcategory'], '')).strip()
        description = str(row.get(col_map['description'], '')).strip()

        # --- Lógica Memo ---
        memo_text = extract_memo_text(
            description, prefix_pattern, all_caps_pattern, verbose)

        # --- Construcción Categoría ---
        category_parts = [part for part in [
            category_csv, subcategory_csv] if part]
        category_for_qif = ":".join(category_parts)

        return {
            'date': tx_date,
            'amount': amount,
            'payee': None,
            'category': category_for_qif,
            'memo': memo_text
        }

    except Exception as e:
        print(f"*** Error INESPERADO fila {original_excel_row}: {e} ***")
        return None


def generate_qif_file(transactions, qif_filepath, output_encoding, verbose=False):
    """Escribe el archivo QIF."""
    print(f"\nGenerando QIF: {qif_filepath}")
    try:
        with open(qif_filepath, mode='w', encoding=output_encoding, errors='replace') as outfile:
            outfile.write("!Type:Bank\n")
            for tx in transactions:
                outfile.write(f"D{tx['date'].strftime('%m/%d/%Y')}\n")
                outfile.write(f"T{tx['amount']:.2f}\n")
                if tx['category']:
                    outfile.write(f"L{tx['category']}\n")
                if tx['memo']:
                    outfile.write(f"M{tx['memo']}\n")
                outfile.write("^\n")
        return True
    except Exception as e:
        print(f"Error Fatal escribiendo QIF: {e}")
        return False

# --- Main ---


def main():
    args = parse_arguments()

    # Determinar nombre salida
    if args.output:
        output_filename = args.output
    else:
        base_name = args.excel_file
        output_filename = re.sub(
            r'\.[Xx][Ll][Ss][Xx]?$', '', base_name, flags=re.IGNORECASE) + ".qif"

    header_idx, _ = find_header_and_metadata(
        args.excel_file, EXPECTED_HEADER, args.verbose)
    if header_idx == -1:
        sys.exit(1)

    df_data = read_excel_data(args.excel_file, header_idx, args.verbose)
    if df_data is None:
        sys.exit(1)

    print("\nProcesando transacciones...")
    processed_transactions = []
    skipped_count = 0

    for idx, row in df_data.iterrows():
        processed_data = process_transaction_row(
            row, header_idx + 2 + idx, COL_MAP, PREFIX_PATTERN, ALL_CAPS_PATTERN, args.verbose
        )
        if processed_data:
            processed_transactions.append(processed_data)
        else:
            skipped_count += 1

    print(f"  Transacciones procesadas: {len(processed_transactions)}")
    print(f"  Filas omitidas: {skipped_count}")

    if not processed_transactions:
        print("Error: No se generaron transacciones válidas.")
        sys.exit(1)

    processed_transactions.sort(key=lambda x: x['date'])
    success = generate_qif_file(
        processed_transactions, output_filename, args.encoding, args.verbose)

    if success:
        print("\n--- ¡Conversión Exitosa! ---")
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
