import streamlit as st
import pandas as pd
from io import BytesIO

# ---------------------- CONFIGURAÇÃO BÁSICA ----------------------
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
st.caption("Atualize planilhas em segundos — substituições automáticas de colunas e valores.")

st.divider()

# ---------------------- UPLOAD ----------------------
uploaded_file = st.file_uploader("📂 Envie o arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Lê apenas a aba "Base" (case-insensitive)
        excel = pd.ExcelFile(uploaded_file)
        sheet_name = next((s for s in excel.sheet_names if s.strip().lower() == "base"), None)
        if not sheet_name:
            st.error("❌ Não foi encontrada aba chamada 'Base'. Verifique o nome da aba.")
            st.stop()

        df = pd.read_excel(excel, sheet_name=sheet_name)

        # ---------------------- MODIFICAÇÕES ----------------------
        total_changes = 0

        # 1️⃣ Substituir "Input Nov" → "Input Dez" em todo o DataFrame
        df = df.map(lambda x: x.replace("Input Nov", "Input Dez") if isinstance(x, str) else x)

        # 2️⃣ Coluna Brand → trocar valores específicos
        if "Brand" in df.columns:
            df["Brand"] = df["Brand"].replace({"ALTACOR": "B1", "AMETISTA": "B2"})
            total_changes += df["Brand"].isin(["B1", "B2"]).sum()

        # 3️⃣ Coluna Regional → trocar valores específicos
        if "Regional" in df.columns:
            df["Regional"] = df["Regional"].replace({"Cana Cerrado": "B3"})
            total_changes += df["Regional"].isin(["B3"]).sum()

        # 4️⃣ Colunas 25-Feb e 25-Jan → todos os valores = 10
        for col in ["25-Feb", "25-Jan"]:
            if col in df.columns:
                df[col] = 10
                total_changes += len(df)

        # ---------------------- RESULTADO ----------------------
        st.success(f"✅ Processamento concluído! {total_changes} modificações realizadas.")

        st.write("**Prévia dos dados atualizados:**")
        st.dataframe(df.head(10))

        # Salvar em memória
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Base")
        output.seek(0)

        # Botão de download
        st.download_button(
            "📥 Baixar Excel Atualizado",
            data=output.getvalue(),
            file_name="fmc_atualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
