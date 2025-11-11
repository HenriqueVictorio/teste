import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

st.set_page_config(page_title="FMC Automação", page_icon="🌿")

st.title("🌿 FMC | Atualizador de 'Input Nov' para 'Input Dez'")
st.caption("Envie o Excel e baixe a versão atualizada em segundos!")

arquivo = st.file_uploader("📂 Envie o arquivo Excel (.xlsx)", type=["xlsx"])

if arquivo:
    try:
        wb = load_workbook(arquivo)
    except Exception:
        st.error("Erro ao abrir o arquivo. Verifique se é um .xlsx válido.")
        st.stop()

    # Encontrar aba 'Base' (case-insensitive)
    aba = None
    for nome in wb.sheetnames:
        if nome.strip().lower() == "base":
            aba = wb[nome]
            break

    if not aba:
        st.warning("A aba 'Base' não foi encontrada. Nenhuma alteração feita.")
    else:
        # Achar colunas "Key Figure"
        cabecalhos = []
        for linha in aba.iter_rows(min_row=1, max_row=aba.max_row, max_col=aba.max_column):
            for cel in linha:
                if not isinstance(cel, MergedCell) and isinstance(cel.value, str):
                    if cel.value.strip().lower() == "key figure":
                        cabecalhos.append((cel.row, cel.column))

        alteracoes = 0
        for (linha_ini, col) in cabecalhos:
            for i in range(linha_ini + 1, aba.max_row + 1):
                cel = aba.cell(row=i, column=col)
                if not isinstance(cel, MergedCell) and isinstance(cel.value, str):
                    if "Input Nov" in cel.value:
                        cel.value = cel.value.replace("Input Nov", "Input Dez")
                        alteracoes += 1

        if alteracoes:
            st.success(f"✅ {alteracoes} célula(s) atualizada(s) com sucesso!")
            buffer = BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            st.download_button(
                label="⬇️ Baixar arquivo atualizado",
                data=buffer,
                file_name="fmc_atualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Nenhuma ocorrência de 'Input Nov' foi encontrada nas colunas 'Key Figure'.")
