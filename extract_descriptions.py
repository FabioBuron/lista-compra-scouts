# -*- coding: utf-8 -*-
"""
Extract title, description/development, and full table rows from the 6 docx files and generate actividades.js
"""
import os
import docx
import json
import re

BASE = os.path.dirname(os.path.abspath(__file__))
DOCS = {
    "CASTORES": "PROGRAMACIÓN CASTORES KOLKALTILLA.docx",
    "LOBATOS": "PROGRAMACION LOBATOS 2026.docx",
    "EXPLORADORES": "PROGRAMACIÓN EXPLORADORES.docx",
    "PIONEROS": "PROGRAMACIÓN PIONEROS 26.docx",
    "RUTAS PANDORA": "Kolkatilla 26 rutas Pandora.docx",
    "RUTAS ICARO": "CAMPA 2026 RUTAS ICARO.docx",
}

def clean_text(text):
    if not text:
        return ""
    # Reemplazar múltiples espacios horizontales por uno solo, conservando saltos de línea
    text = re.sub(r'[ \t\r\f\v]+', ' ', text).strip()
    return text

def normalize_key(text):
    if not text:
        return ""
    # A minúsculas, eliminar acentos, espacios y caracteres no alfanuméricos
    text = text.lower().strip()
    # Eliminar acentos básicos
    text = text.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    text = text.replace("ñ", "n")
    # Dejar solo caracteres alfanuméricos
    text = re.sub(r'[^a-z0-9]', '', text)
    return text

data_map = {}

for unit, doc_name in DOCS.items():
    path = os.path.join(BASE, doc_name)
    unit_key = unit.lower().replace(" ", "")
    data_map[unit_key] = {}
    
    if not os.path.exists(path):
        print(f"File not found: {doc_name}")
        continue
        
    print(f"Parsing {doc_name} for unit {unit}...")
    doc = docx.Document(path)
    
    for t_idx, table in enumerate(doc.tables):
        title = ""
        description = ""
        table_rows = []
        
        # Primero buscamos el título y descripción
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if not cells:
                continue
            
            # Buscar si alguna celda es la cabecera de título o descripción
            for idx, cell_val in enumerate(cells):
                cell_clean = cell_val.upper().strip()
                
                # Coincidencia para el título
                if cell_clean in ["TÍTULO", "TITULO", "TÍTULO DE LA ACTIVIDAD", "TITULO DE LA ACTIVIDAD"]:
                    if idx + 1 < len(cells):
                        title = clean_text(cells[idx + 1])
                    break
                
                # Coincidencia para la descripción / desarrollo
                if cell_clean in ["DESCRIPCIÓN", "DESCRIPCION", "DESARROLLO", "DESCRIPCIÓN Y DESARROLLO", "DESCRIPCION Y DESARROLLO"]:
                    if idx + 1 < len(cells):
                        description = cells[idx + 1].strip() # Conservar saltos de línea para el modal
                    break
                    
        # Si encontramos título, procedemos a recopilar la tabla completa
        if title:
            # Si la descripción está vacía, buscar observaciones
            if not description:
                for row in table.rows:
                    cells = [cell.text.strip() for cell in row.cells]
                    for idx, cell_val in enumerate(cells):
                        if cell_val.upper().strip() == "OBSERVACIONES":
                            if idx + 1 < len(cells):
                                description = cells[idx + 1].strip()
                            break
            
            # Extraer todas las filas de la tabla
            for row in table.rows:
                row_cells = [clean_text(cell.text) for cell in row.cells]
                # Eliminar duplicados consecutivos por celdas combinadas
                unique_cells = []
                for c in row_cells:
                    if not unique_cells or unique_cells[-1] != c:
                        unique_cells.append(c)
                if any(unique_cells):
                    table_rows.append(unique_cells)
                    
            title_norm = normalize_key(title)
            data_map[unit_key][title_norm] = {
                "original_title": title,
                "description": description or "No hay descripción disponible para esta actividad.",
                "rows": table_rows
            }

# Escribir el archivo javascript
js_content = f"// Archivo generado automáticamente con todas las descripciones y tablas de las actividades\n"
js_content += f"const ACTIVIDADES_COMPLETAS = {json.dumps(data_map, indent=2, ensure_ascii=False)};\n"

output_js_path = os.path.join(BASE, "actividades.js")
with open(output_js_path, "w", encoding="utf-8") as f:
    f.write(js_content)

print(f"\nDone. Written to {output_js_path}")
