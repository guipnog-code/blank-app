import os
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF

EXCEL_FILE = "Cadastros_Servidores.xlsx"

def salvar_no_excel(dados):
    df_novo = pd.DataFrame([dados])
    if os.path.exists(EXCEL_FILE):
        df_existente = pd.read_excel(EXCEL_FILE)
        df_final = pd.concat([df_existente, df_novo], ignore_index=True)
    else:
        df_final = df_novo
    df_final.to_excel(EXCEL_FILE, index=False)

def preencher_documentos_oficiais(dados):
    caminho_procuracao = "template_procuracao.pdf"
    caminho_termo = "template_termo.pdf"
    
    pdf_procuracao_bytes = None
    pdf_termo_bytes = None

    if os.path.exists(caminho_procuracao):
        doc_proc = fitz.open(caminho_procuracao)
        pag_proc = doc_proc[0]
        pag_proc.insert_text((92, 184), dados['Nome'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((83, 200), dados['CPF'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((223, 200), dados['RG'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((369, 200), dados['Cargo'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((94, 215), dados['Órgão'], fontsize=8, color=(0,0,0))
        pag_proc.insert_text((284, 215), dados['Data de Ingresso'], fontsize=9, color=(0,0,0))
        pdf_procuracao_bytes = doc_proc.tobytes()
        doc_proc.close()

    if os.path.exists(caminho_termo):
        doc_termo = fitz.open(caminho_termo)
        pag_termo = doc_termo[0]
        pag_termo.insert_text((100, 150), dados['Nome'], fontsize=9, color=(0,0,0))
        pag_termo.insert_text((100, 170), dados['CPF'], fontsize=9, color=(0,0,0))
        pdf_termo_bytes = doc_termo.tobytes()
        doc_termo.close()

    return pdf_procuracao_bytes, pdf_termo_bytes

st.title("📋 Cadastro e Preenchimento de Documentos")
st.write("Preencha os dados abaixo para cadastrar e gerar os documentos em PDF.")

# Inicializa o estado da sessão para manter os botões visíveis
if "pdf_proc" not in st.session_state:
    st.session_state.pdf_proc = None
if "pdf_termo" not in st.session_state:
    st.session_state.pdf_termo = None

with st.form("form_cadastro"):
    st.subheader("Dados Profissionais")
    matricula = st.text_input("Matrícula (SIAPE)")
    cargo = st.text_input("Cargo")
    orgao = st.text_input("Órgão")
    ingresso = st.text_input("Data de Ingresso")

    st.subheader("Dados Pessoais")
    nome = st.text_input("Nome Completo")
    cpf = st.text_input("CPF")
    email = st.text_input("E-mail")
    rg = st.text_input("RG")
    telefone = st.text_input("Telefone")
    estado_civil = st.text_input("Estado Civil")
    cep = st.text_input("CEP")
    endereco = st.text_input("Endereço")
    municipio = st.text_input("Município")
    estado = st.text_input("Estado (UF)")

    submitted = st.form_submit_button("Salvar e Gerar Documentos")

if submitted:
    dados_usuario = {
        "Matrícula": matricula,
        "Cargo": cargo,
        "Órgão": orgao,
        "Data de Ingresso": ingresso,
        "Nome": nome,
        "CPF": cpf,
        "E-mail": email,
        "RG": rg,
        "Telefone": telefone,
        "Estado Civil": estado_civil,
        "CEP": cep,
        "Endereço": endereco,
        "Município": municipio,
        "Estado": estado
    }
    
    salvar_no_excel(dados_usuario)
    st.success("Dados salvos com sucesso no sistema!")

    # Guarda os arquivos na sessão para não sumirem após o clique
    st.session_state.pdf_proc, st.session_state.pdf_termo = preencher_documentos_oficiais(dados_usuario)

# Exibe os botões se os arquivos já estiverem gerados na sessão
if st.session_state.pdf_proc or st.session_state.pdf_termo:
    st.markdown("### 📥 Baixar Documentos Gerados")
    
    col1, col2 = st.columns(2)
    
    if st.session_state.pdf_proc:
        with col1:
            st.download_button(
                label="📄 Baixar Procuração",
                data=st.session_state.pdf_proc,
                file_name="Procuracao_Preenchida.pdf",
                mime="application/pdf"
            )
            
    if st.session_state.pdf_termo:
        with col2:
            st.download_button(
                label="📄 Baixar Termo",
                data=st.session_state.pdf_termo,
                file_name="Termo_Preenchido.pdf",
                mime="application/pdf"
            )