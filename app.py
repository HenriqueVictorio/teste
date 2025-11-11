"""Streamlit app for FMC to update Excel Key Figure column."""

from __future__ import annotations

from io import BytesIO
from typing import Dict, Iterable, List, Optional, Tuple

import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.exceptions import InvalidFileException

st.set_page_config(
    page_title="FMC Automação",
    page_icon="🌿",
    layout="centered",
)

_CUSTOM_CSS = """
<style>
:root {
    --fmc-green: #007a4d;
    --fmc-green-light: #ebf5f1;
    --fmc-dark: #1d1d1b;
}

[data-testid="stAppViewContainer"] {
    background: linear-gradient(180deg, #ffffff 0%, var(--fmc-green-light) 100%);
}

[data-testid="stHeader"] {
    background: transparent;
}

.block-container {
    padding-top: 3rem;
    max-width: 880px;
}

.stButton button {
    background-color: var(--fmc-green);
    border: none;
    color: white;
    font-weight: 600;
}

.stButton button:hover {
    background-color: #005f3a;
}

.stDownloadButton button {
    background-color: white;
    border: 2px solid var(--fmc-green);
    color: var(--fmc-green);
    font-weight: 600;
}

.stDownloadButton button:hover {
    background-color: var(--fmc-green);
    color: white;
}
</style>
"""

st.markdown(_CUSTOM_CSS, unsafe_allow_html=True)

st.title("FMC | Portal de Automação de Relatórios")
st.caption(
    "Atualize arquivos de planejamento em segundos: basta enviar o Excel e baixar a versão com 'Input Dez'."
)

st.divider()

with st.expander("ℹ️ Como funciona", expanded=False):
    st.markdown(
        "- O aplicativo localiza todas as colunas **Key Figure** em cada planilha.\n"
        "- Todo valor de célula contendo `Input Nov` é substituído por `Input Dez`.\n"
        "- A formatação original e demais conteúdos do arquivo são preservados."
    )

uploaded_file = st.file_uploader(
    "Envie o arquivo Excel da FMC", type=["xlsx"], help="Somente arquivos .xlsx são aceitos."
)


def _find_key_figure_headers(sheet: Worksheet) -> List[Tuple[int, int]]:
    """Return all (row, column) coordinates where the header equals 'Key Figure'."""
    headers: List[Tuple[int, int]] = []
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, max_col=sheet.max_column):
        for cell in row:
            if isinstance(cell, MergedCell) or cell.row is None or cell.column is None:
                continue
            value = cell.value
            if isinstance(value, str) and value.strip().lower() == "key figure":
                headers.append((cell.row, cell.column))
    return headers


def _replace_input_values(sheet: Worksheet, headers: Iterable[Tuple[int, int]]) -> int:
    """Replace 'Input Nov' with 'Input Dez' below each header and return hit count."""
    replacements = 0
    for header_row, column in headers:
        for row_idx in range(header_row + 1, sheet.max_row + 1):
            cell = sheet.cell(row=row_idx, column=column)
            if isinstance(cell, MergedCell):
                continue
            value = cell.value
            if isinstance(value, str) and "Input Nov" in value:
                cell.value = value.replace("Input Nov", "Input Dez")
                replacements += 1
    return replacements


def _process_workbook(xlsx_bytes: bytes) -> Tuple[BytesIO, Dict[str, int]]:
    """Load workbook, apply replacements, and return stream plus per-sheet summary."""
    workbook = load_workbook(BytesIO(xlsx_bytes), data_only=False)
    summary: Dict[str, int] = {}

    for sheet in workbook.worksheets:
        headers = _find_key_figure_headers(sheet)
        if not headers:
            continue
        hits = _replace_input_values(sheet, headers)
        if hits:
            summary[sheet.title] = hits

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output, summary


result_stream: Optional[BytesIO] = None
report: Dict[str, int] = {}

if uploaded_file:
    file_bytes = uploaded_file.getvalue()
    try:
        result_stream, report = _process_workbook(file_bytes)
    except InvalidFileException:
        st.error("Arquivo inválido. Verifique se está enviando um .xlsx compatível.")
        st.stop()
    except Exception as exc:  # pragma: no cover - defensive for unexpected cases
        st.error(f"Ocorreu um erro inesperado: {exc}")
        st.stop()

    if report:
        total_replacements = sum(report.values())
        st.success(
            f"Processamento concluído! {total_replacements} célula(s) atualizada(s) em "
            f"{len(report)} planilha(s)."
        )
        with st.container():
            st.markdown("**Detalhes por aba:**")
            for sheet_name, count in report.items():
                st.write(f"• {sheet_name}: {count} substituição(ões)")
    else:
        st.warning(
            "Nenhuma ocorrência de 'Input Nov' foi encontrada na(s) coluna(s) Key Figure."
        )

    if result_stream is not None:
        st.download_button(
            "Baixar Excel Atualizado",
            data=result_stream.getvalue(),
            file_name="fmc_key_figure_atualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.divider()

