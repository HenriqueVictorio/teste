# Streamlit (leve) – FMC: atualizar "Input Nov" -> "Input Dez" na aba 'Base'
# Otimizações:
# 1) Evita ler a planilha inteira: usa bounds reais (calculate_dimension) para limitar linhas/colunas.
# 2) Procura o header "Key Figure" só nas primeiras linhas úteis (padrão: 5).
# 3) Opera apenas nas colunas onde o header foi encontrado, a partir da linha do header.
# 4) Evita copiar bytes desnecessariamente: passa o arquivo diretamente para o openpyxl.
# 5) Interface Streamlit mínima (sem CSS pesado).

from io import BytesIO
from typing import Dict, Tuple

import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import range_boundaries
from openpyxl.utils.exceptions import InvalidFileException

# ---------------- UI ----------------
st.set_page_config(page_title="FMC Automação (Leve)", page_icon="🌿", layout="centered")

st.title("FMC | Portal de Automação de Relatórios (Offline/Leve)")
st.caption("Envie o .xlsx. O app procura a aba **Base**, acha as colunas **Key Figure** "
           "e substitui `Input Nov` por `Input Dez`, preservando o restante do arquivo.")

uploaded_file = st.file_uploader("Envie o arquivo Excel (.xlsx)", type=["xlsx"])

# ---------------- Núcleo de desempenho ----------------
def _used_bounds(ws: Worksheet) -> Tuple[int, int, int, int]:
    """Retorna os limites reais de uso (min_col, min_row, max_col, max_row)."""
    # Ex.: 'A1:D57' reflete a área com dados, ignorando formatação "fantasma"
    min_c, min_r, max_c, max_r = range_boundaries(ws.calculate_dimension())
    return min_c, min_r, max_c, max_r

def _find_key_figure_positions(ws: Worksheet, header_rows: int = 5) -> Dict[int, int]:
    """
    Retorna {coluna: linha_do_header} para cada coluna cujo header (em até 'header_rows' linhas do topo usado)
    seja exatamente 'Key Figure' (case-insensitive).
    """
    positions: Dict[int, int] = {}
    min_c, min_r, max_c, max_r = _used_bounds(ws)
    max_header_row = min(min_r + header_rows - 1, max_r)
    target = "key figure"

    for r in range(min_r, max_header_row + 1):
        for c in range(min_c, max_c + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            v = cell.value
            if isinstance(v, str) and v.strip().lower() == target:
                # guarda a primeira ocorrência por coluna (mais alta)
                positions.setdefault(c, r)
    return positions

def _replace_values(ws: Worksheet, col_to_header_row: Dict[int, int]) -> int:
    """Para cada coluna alvo, substitui 'Input Nov' por 'Input Dez' abaixo do cabeçalho. Retorna contagem."""
    if not col_to_header_row:
        return 0

    min_c, min_r, max_c, max_r = _used_bounds(ws)
    hits = 0
    for c, header_row in col_to_header_row.items():
        start = header_row + 1
        if start > max_r:
            continue
        for r in range(start, max_r + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            v = cell.value
            if isinstance(v, str) and "Input Nov" in v:
                cell.value = v.replace("Input Nov", "Input Dez")
                hits += 1
    return hits

def _process_workbook(file_like) -> Tuple[BytesIO, Dict[str, int]]:
    """Carrega, processa e devolve stream do arquivo atualizado + sumário por aba."""
    # Carrega sem copiar bytes para memória; mantém formatação (write_only=False)
    wb = load_workbook(file_like, data_only=False, read_only=False, keep_vba=False)

    # Localiza 'Base' (case-insensitive)
    base_name = None
    for name in wb.sheetnames:
        if isinstance(name, str) and name.strip().lower() == "base":
            base_name = name
            break

    summary: Dict[str, int] = {}
    if base_name is not None:
        ws = wb[base_name]
        positions = _find_key_figure_positions(ws, header_rows=5)
        hits = _replace_values(ws, positions)
        if hits:
            summary[ws.title] = hits

    # Salva em memória
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out, summary

# ---------------- Execução ----------------
if uploaded_file is not None:
    try:
        uploaded_file.seek(0)  # garante início do stream
        result_stream, report = _process_workbook(uploaded_file)
    except InvalidFileException:
        st.error("Arquivo inválido. Envie um .xlsx compatível.")
        st.stop()
    except Exception as e:
        st.error(f"Ocorreu um erro inesperado: {e}")
        st.stop()

    if report:
        total = sum(report.values())
        st.success(f"Concluído! {total} célula(s) atualizada(s) na aba {list(report.keys())[0]!r}.")
    else:
        st.warning("Nada para alterar: não encontrei a aba 'Base' ou colunas 'Key Figure' com 'Input Nov'.")

    st.download_button(
        "⬇️ Baixar Excel Atualizado",
        data=result_stream.getvalue(),
        file_name="fmc_key_figure_atualizado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )