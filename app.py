import streamlit as st
import pandas as pd
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
st.caption("Atualize planilhas em segundos — substituições automáticas otimizadas.")

st.divider()

# ---------------------- UPLOAD ----------------------
uploaded_file = st.file_uploader("📂 Envie o arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    start_time = time.time()  # Início da contagem de tempo

    try:
        # Ler aba "Base"
        excel = pd.ExcelFile(uploaded_file)
        sheet_name = next((s for s in excel.sheet_names if s.strip().lower() == "base"), None)
        if not sheet_name:
            st.error("❌ Não foi encontrada aba chamada 'Base'. Verifique o nome da aba.")
            st.stop()

        df = pd.read_excel(excel, sheet_name=sheet_name)
        total_changes = 0

        # 1️⃣ Substituir "Input Nov" → "Input Dez" em todo o DataFrame (strings apenas)
        df = df.applymap(lambda x: x.replace("Input Nov", "Input Dez") if isinstance(x, str) else x)

        # Guardar índice de linhas modificadas
        changed_rows = set()

        # 2️⃣ Coluna Brand → trocar valores específicos
        if "Brand" in df.columns:
            mask_brand = df["Brand"].isin(["ALTACOR", "AMETISTA"])
            df.loc[df["Brand"] == "ALTACOR", "Brand"] = "B1"
            df.loc[df["Brand"] == "AMETISTA", "Brand"] = "B2"
            changed_rows.update(df.index[mask_brand].tolist())
            total_changes += int(mask_brand.sum())

        # 3️⃣ Coluna Regional → trocar valores específicos
        if "Regional" in df.columns:
            mask_regional = df["Regional"] == "Cana Cerrado"
            df.loc[mask_regional, "Regional"] = "B3"
            changed_rows.update(df.index[mask_regional].tolist())
            total_changes += int(mask_regional.sum())

        # 4️⃣ Colunas 25-Feb e 25-Jan → valor 10 apenas nas linhas alteradas
        if changed_rows:
            for col in ["25-Feb", "25-Jan"]:
                if col in df.columns:
                    df.loc[list(changed_rows), col] = 10
                    total_changes += len(changed_rows)

        # ---------------------- FINALIZAÇÃO ----------------------
        elapsed = time.time() - start_time

        st.success(f"✅ Processamento concluído! {total_changes} alterações em {elapsed:.2f} segundos.")
        st.write("**Prévia dos dados atualizados (primeiras 10 linhas):**")
        st.dataframe(df.head(10))

        # ---------------------- SALVAR E PREPARAR DOWNLOAD ----------------------
        output = BytesIO()
        saved_as_excel = False

        # Tentar salvar como .xlsx usando openpyxl (mais comum)
        try:
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Base")
            output.seek(0)
            st.download_button(
                "📥 Baixar Excel Atualizado (.xlsx)",
                data=output.getvalue(),
                file_name="fmc_atualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            saved_as_excel = True
        except Exception as e_openpyxl:
            # Se falhar, tentar sem especificar engine (pandas escolherá se puder)
            try:
                output = BytesIO()
                with pd.ExcelWriter(output) as writer:
                    df.to_excel(writer, index=False, sheet_name="Base")
                output.seek(0)
                st.download_button(
                    "📥 Baixar Excel Atualizado (.xlsx)",
                    data=output.getvalue(),
                    file_name="fmc_atualizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                saved_as_excel = True
            except Exception as e_all:
                # Fallback: gerar CSV para download e informar como instalar libs
                csv_bytes = df.to_csv(index=False).encode("utf-8")
                st.warning("Não foi possível criar um arquivo .xlsx neste ambiente. Fornecendo .csv como alternativa.")
                st.download_button(
                    "📥 Baixar CSV Atualizado (.csv)",
                    data=csv_bytes,
                    file_name="fmc_atualizado.csv",
                    mime="text/csv",
                )
                st.info(
                    "Se quiser habilitar o download em .xlsx no futuro, instale no ambiente um dos motores:\n"
                    "`pip install openpyxl`  (recomendado)  ou  `pip install xlsxwriter`"
                )

        # Exibir log simples por coluna (opcional)
        log_msgs = []
        if "Brand" in df.columns:
            log_msgs.append(f"Brand: {int(df['Brand'].isin(['B1','B2']).sum())} mudanças")
        if "Regional" in df.columns:
            log_msgs.append(f"Regional: {int(df['Regional'].isin(['B3']).sum())} mudanças")
        for col in ["25-Feb", "25-Jan"]:
            if col in df.columns and changed_rows:
                log_msgs.append(f"{col}: {len(changed_rows)} mudanças (linhas alteradas)")

        if log_msgs:
            st.write("**Resumo por coluna:**")
            st.write(" • " + " | ".join(log_msgs))

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
