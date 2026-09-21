import pandas as pd
from pathlib import Path
import logging

LOGGER = logging.getLogger(__name__)

def match_totals_in_macro(extracted_data: list, macro_path: Path | None) -> list:
    """
    Toma la lista de facturas (con sus totales) y busca ese total en la columna D
    de la hoja 'TODOS'. Si lo encuentra, sube celda por celda tomando los valores
    numéricos de la columna D y sus textos en la columna C, hasta que encuentre
    una celda no numérica o vacía.
    """
    if not macro_path or not Path(macro_path).exists():
        return extracted_data

    try:
        # Columna C es el índice 2, Columna D es el índice 3, Columna E es el índice 4
        df = pd.read_excel(macro_path, sheet_name="TODOS", header=None, engine="openpyxl", usecols=[2, 3, 4])
    except Exception as exc:
        LOGGER.warning(f"Error loading macro for matching totals: {exc}")
        return extracted_data

    col_c = df.iloc[:, 0].tolist()
    col_d = df.iloc[:, 1].tolist()
    col_e = df.iloc[:, 2].tolist()
    
    def is_numeric(val):
        if pd.isna(val) or val is None:
            return False
        if isinstance(val, (int, float)):
            return True
        try:
            float(str(val).replace(',', ''))
            return True
        except ValueError:
            return False

    used_indices = set()
    
    for item in extracted_data:
        total = item[0]
        if total is None:
            continue
            
        try:
            target_total = float(total)
        except ValueError:
            continue
            
        # Buscamos el total en la columna D
        found = False
        for i, val_d in enumerate(col_d):
            if i in used_indices:
                continue
                
            if is_numeric(val_d) and abs(float(str(val_d).replace(',', '')) - target_total) <= 1.0:
                # ¡Encontramos el total!
                matched_items = []
                j = i - 1
                block_indices = {i}
                
                # Subir por la columna mientras haya datos numéricos
                while j >= 0:
                    val_up = col_d[j]
                    if is_numeric(val_up):
                        text_left = col_c[j]
                        text_left = "" if pd.isna(text_left) else str(text_left).strip()
                        
                        tipo_right = col_e[j]
                        tipo_right = "" if pd.isna(tipo_right) else str(tipo_right).strip()
                        
                        numeric_val = float(str(val_up).replace(',', ''))
                        matched_items.append({
                            "texto": text_left, 
                            "numero": numeric_val,
                            "tipo": tipo_right
                        })
                        
                        block_indices.add(j)
                        j -= 1
                    else:
                        break # Paramos si la celda no es numérica o está vacía
                
                # Guardar en la lista interna de la factura
                if not matched_items:
                    item[1] = [{"texto": "Total encontrado en el archivo, pero no hay desglose numérico en las celdas superiores.", "numero": 0}]
                else:
                    item[1] = matched_items
                
                used_indices.update(block_indices)
                found = True
                break
                
        if not found:
            item[1] = [{"texto": "No se encontraron datos que coincidan con el total de esta factura en el archivo macro.", "numero": 0}]
                
    return extracted_data

import difflib

def get_macro_concept_mapping(macro_path) -> dict:
    if not macro_path or not __import__('pathlib').Path(macro_path).exists():
        return {}
    try:
        df = pd.read_excel(macro_path, sheet_name='TODOS', header=None, engine='openpyxl', usecols=[1, 2, 3])
    except Exception:
        return {}
    mapping = {}
    col_b = df.iloc[:, 0].tolist()
    col_c = df.iloc[:, 1].tolist()
    col_d = df.iloc[:, 2].tolist()
    for b_val, c_val, d_val in zip(col_b, col_c, col_d):
        if pd.isna(d_val) or d_val is None:
            continue
        try:
            float(str(d_val).replace(',', ''))
        except ValueError:
            continue
        account = str(b_val).strip()
        concept = str(c_val).strip()
        if not account.isdigit() or concept == 'nan' or not concept:
            continue
        if account not in mapping:
            mapping[account] = set()
        mapping[account].add(concept)
    return {k: list(v) for k, v in mapping.items()}

def predict_concept_account(description: str, macro_mapping: dict) -> str | None:
    if not description or not macro_mapping:
        return None
    concept_to_account = {}
    for account, concepts in macro_mapping.items():
        for concept in concepts:
            concept_to_account[concept.upper()] = account
    desc_upper = description.upper()
    if desc_upper in concept_to_account:
        return concept_to_account[desc_upper]
    matches = difflib.get_close_matches(desc_upper, concept_to_account.keys(), n=1, cutoff=0.7)
    if matches:
        return concept_to_account[matches[0]]
    return None
