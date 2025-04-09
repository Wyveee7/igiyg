import streamlit as st
import ezdxf
import pandas as pd
import tempfile
import os
from openpyxl import Workbook

def extract_entities(dxf_path):
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()
    texts = []
    lines = []
    for e in msp:
        if e.dxftype() == "TEXT":
            x, y = round(e.dxf.insert.x, 2), round(e.dxf.insert.y, 2)
            texts.append((x, y, e.dxf.text.strip()))
        elif e.dxftype() == "LINE":
            x1, y1 = round(e.dxf.start.x, 2), round(e.dxf.start.y, 2)
            x2, y2 = round(e.dxf.end.x, 2), round(e.dxf.end.y, 2)
            lines.append(((x1, y1), (x2, y2)))
    return texts, lines

def group_lines_to_cells(lines):
    horiz = sorted([l for l in lines if round(l[0][1], 2) == round(l[1][1], 2)], key=lambda l: l[0][1])
    vert = sorted([l for l in lines if round(l[0][0], 2) == round(l[1][0], 2)], key=lambda l: l[0][0])
    x_vals = sorted(set([round(p[0], 2) for l in vert for p in l]))
    y_vals = sorted(set([round(p[1], 2) for l in horiz for p in l]), reverse=True)
    cells = []
    for yi in range(len(y_vals) - 1):
        for xi in range(len(x_vals) - 1):
            x0, x1 = x_vals[xi], x_vals[xi+1]
            y0, y1 = y_vals[yi+1], y_vals[yi]
            cells.append(((x0, y0, x1, y1), []))
    return cells

def assign_texts_to_cells(texts, cells):
    for x, y, text in texts:
        for (x0, y0, x1, y1), contents in cells:
            if x0 <= x <= x1 and y0 <= y <= y1:
                contents.append(text)
                break
    return cells

def build_tables_from_cells(cells):
    from collections import defaultdict
    rows_dict = defaultdict(lambda: defaultdict(str))
    for (x0, y0, x1, y1), contents in cells:
        center_y = (y0 + y1) / 2
        center_x = (x0 + x1) / 2
        rows_dict[center_y][center_x] = "
".join(contents)
    rows = []
    for y in sorted(rows_dict.keys(), reverse=True):
        row = []
        for x in sorted(rows_dict[y].keys()):
            row.append(rows_dict[y][x])
        rows.append(row)
    return [rows] if rows else []

# App
st.set_page_config(page_title="Extrair Tabela do DXF", layout="centered")
st.title("📐 Extrator de Tabelas DXF para Excel")

uploaded_file = st.file_uploader("Faça upload do arquivo DXF", type=["dxf"])

if uploaded_file:
    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, uploaded_file.name)
        with open(path, "wb") as f: f.write(uploaded_file.read())

        st.info("⏳ Processando o arquivo...")
        texts, lines = extract_entities(path)
        cells = group_lines_to_cells(lines)
        cells_filled = assign_texts_to_cells(texts, cells)
        tables = build_tables_from_cells(cells_filled)

        if tables:
            output_path = os.path.join(tmpdir, "tabelas_extraidas.xlsx")
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                for i, table in enumerate(tables):
                    pd.DataFrame(table).to_excel(writer, index=False, header=False, sheet_name=f"Tabela_{i+1}")

            st.success(f"✅ {len(tables)} tabela(s) extraída(s)!")
            with open(output_path, "rb") as f:
                st.download_button("📥 Baixar Excel", data=f, file_name="tabelas_extraidas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("Nenhuma tabela reconhecida no arquivo.")