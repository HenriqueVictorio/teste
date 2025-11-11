import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import time

# ---------------------- CONFIGURAÇÃO VISUAL ----------------------
st.set_page_config(page_title="FMC Automação", page_icon="🌿", layout="centered")

st.markdown("""
<style>
:root {
    --fmc-green: #007a4d;
    --fmc-green-light: #ebf5f1;
}
[data-testid="stAppViewContainer"] {
    background: linear-gradient(180deg, #ffffff 0%, var(--fmc-green-light) 100%);
}
[data-testid="stHeader"] {background: transparent;}
.block-container {padding-top: 3rem; max-width: 900px;}
.stButton button {
    background-color: var(--fmc-green); color: white; font-weight: 600; border: none;
}
.stButton button:hover {background-color: #005f3a;}
.stDownloadButton button {
    background-color: white; border: 2px solid var(--fmc-green); color: var(--fmc-green);
    font-weight: 600;
}
.stDownloadButton button:hover {background-color: var(--fmc-green); color: white;}
</style>
""", unsafe_allow_html=True)

st.title("🌿 FMC | Portal de Automação de Relatórios")
st.caption("Mantém estrutura original do Excel e registra tempos por etapa.")

st.divider()

uploaded_file = st.file_uploader("📂 Envie o arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    global_start = time.time()
    timings = {}

    try:
        # 🕒 Leitura do Excel (mantendo estrutura)
        t0 = time.time()
        wb = load_workbook(uploaded_file)
        if "Base" not in wb.sheetnames:
            st.error("❌ Não foi encontrada aba chamada 'Base'. Verifique o nome da aba.")
            st.stop()
        ws = wb["Base"]
        timings["Leitura do Excel"] = time.time() - t0

        # Extrair cabeçalhos
        headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        df = pd.DataFrame(ws.iter_rows(min_row=2, values_only=True), columns=headers)

        total_changes = 0
        changed_rows = set()

        # 🔁 Substituições “Input Nov” → “Input Dez”
        t1 = time.time()
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, str) and "Input Nov" in cell.value:
                    cell.value = cell.value.replace("Input Nov", "Input Dez")
                    total_changes += 1
        timings["Substituição Input Nov/Dez"] = time.time() - t1

        # 🔄 Alterações em Brand e Regional
        t2 = time.time()
        col_brand = headers.index("Brand") + 1 if "Brand" in headers else None
        col_regional = headers.index("Regional") + 1 if "Regional" in headers else None

        if col_brand:
            for i, cell in enumerate(ws.iter_cols(min_col=col_brand, max_col=col_brand, min_row=2)[0], start=2):
                if cell.value == "ALTACOR":
                    cell.value = "B1"
                    changed_rows.add(i)
                    total_changes += 1
                elif cell.value == "AMETISTA":
                    cell.value = "B2"
                    changed_rows.add(i)
                    total_changes += 1

        if col_regional:
            for i, cell in enumerate(ws.iter_cols(min_col=col_regional, max_col=col_regional, min_row=2)[0], start=2):
                if cell.value == "Cana Cerrado":
                    cell.value = "B3"
                    changed_rows.add(i)
                    total_changes += 1
        timings["Alterações Brand/Regional"] = time.time() - t2

        # 🔢 Atualização das colunas “25-Feb” e “25-Jan” nas linhas alteradas
        t3 = time.time()
        for col_name in ["25-Feb", "25-Jan"]:
            if col_name in headers:
                col_idx = headers.index(col_name) + 1
                for row_num in changed_rows:
                    ws.cell(row=row_num, column=col_idx).value = 10
                    total_changes += 1
        timings["Atualização 25-Feb/25-Jan"] = time.time() - t3

        # 💾 Salvando mantendo tudo original
        t4 = time.time()
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        timings["Gravação Excel"] = time.time() - t4

        total_time = time.time() - global_start

        # ---------------------- RESULTADOS ----------------------
        st.success(f"✅ Automação concluída em {total_time:.2f} segundos, total de {total_changes} alterações.")
        st.write("**Detalhamento de tempo por etapa:**")
        for etapa, t in timings.items():
            st.write(f"• {etapa}: {t:.2f} s")

        st.divider()
        st.download_button(
            "📥 Baixar Excel Atualizado (estrutura preservada)",
            data=output.getvalue(),
            file_name="fmc_atualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.caption("✅ Estrutura, fórmulas e formatação 100% preservadas.")

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
