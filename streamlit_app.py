import streamlit as st
import datetime
import pandas as pd
import os
import fitz  # PyMuPDF

PASTA_ATUAL = os.path.dirname(os.path.abspath(__file__))

def salvar_no_excel(dados):
    arquivo_excel = os.path.join(PASTA_ATUAL, "Cadastros_Servidores.xlsx")
    df_novo = pd.DataFrame([dados])
    
    if os.path.exists(arquivo_excel):
        df_existente = pd.read_excel(arquivo_excel)
        df_final = pd.concat([df_existente, df_novo], ignore_index=True)
    else:
        df_final = df_novo
        
    df_final.to_excel(arquivo_excel, index=False)

def preencher_documentos_oficiais(dados, pasta_destino):
    nome_arquivo = dados['nome'].replace(" ", "_")
    caminho_termo = os.path.join(PASTA_ATUAL, "template_termo.pdf")
    caminho_procuracao = os.path.join(PASTA_ATUAL, "template_procuracao.pdf")

    # --- 1. PREENCHENDO O TERMO ---
    if os.path.exists(caminho_termo):
        doc_termo = fitz.open(caminho_termo)
        pag_termo = doc_termo[0]
        
        pag_termo.insert_text((125, 145), dados['nome'], fontsize=9, color=(0,0,0))
        pag_termo.insert_text((82, 170), dados['cpf'], fontsize=9, color=(0,0,0))
        pag_termo.insert_text((265, 170), dados['matricula'], fontsize=9, color=(0,0,0))
        pag_termo.insert_text((377, 169), dados['cargo'], fontsize=9, color=(0,0,0))

        doc_termo.save(os.path.join(pasta_destino, f"Termo_{nome_arquivo}.pdf"))
        doc_termo.close()

    # --- 2. PREENCHENDO A PROCURAÇÃO COM AS COORDENADAS E FONTES AJUSTADAS ---
    if os.path.exists(caminho_procuracao):
        doc_proc = fitz.open(caminho_procuracao)
        pag_proc = doc_proc[0]
        
        pag_proc.insert_text((92, 184), dados['nome'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((83, 200), dados['cpf'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((223, 200), dados['rg'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((369, 200), dados['cargo'], fontsize=9, color=(0,0,0))
        
        # Órgão corrigido para (94, 215) com fonte 8
        pag_proc.insert_text((94, 215), dados['orgao'], fontsize=8, color=(0,0,0))
        # Ingresso corrigido para (284, 215)
        pag_proc.insert_text((284, 215), dados['data_ingresso'], fontsize=9, color=(0,0,0))
        # Estado civil corrigido para (430, 215) com fonte 8
        pag_proc.insert_text((430, 215), dados['estado_civil'], fontsize=8, color=(0,0,0))
        
        pag_proc.insert_text((109, 230), dados['telefone'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((253, 230), dados['email'], fontsize=9, color=(0,0,0))
        
        pag_proc.insert_text((115, 245), dados['endereco'], fontsize=9, color=(0,0,0))
        
        pag_proc.insert_text((99, 260), dados['municipio'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((394, 260), dados['estado'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((457, 260), dados['cep'], fontsize=9, color=(0,0,0))

        doc_proc.save(os.path.join(pasta_destino, f"Procuracao_{nome_arquivo}.pdf"))
        doc_proc.close()

# ---------------------------------------------------------
# INTERFACE DO STREAMLIT
# ---------------------------------------------------------
st.set_page_config(page_title="Sistema de Cadastro", page_icon="📋", layout="centered")

st.title("📋 Cadastro e Preenchimento de Documentos")
st.markdown("Preencha os dados abaixo para gerar o Termo e a Procuração nas coordenadas configuradas.")

with st.form("formulario_cadastro", clear_on_submit=False):
    st.subheader("Dados Profissionais")
    col1, col2 = st.columns(2)
    with col1:
        matricula = st.text_input("Matrícula (SIAPE)")
        cargo = st.text_input("Cargo")
    with col2:
        orgao = st.text_input("Órgão", value="POLÍCIA RODOVIÁRIA FEDERAL")
        data_ingresso = st.date_input("Data de Ingresso", min_value=datetime.date(1950, 1, 1), max_value=datetime.date.today())

    st.divider()

    st.subheader("Dados Pessoais")
    nome = st.text_input("Nome Completo")
    
    col3, col4 = st.columns(2)
    with col3:
        cpf = st.text_input("CPF")
        email = st.text_input("E-mail")
    with col4:
        rg = st.text_input("RG")
        telefone = st.text_input("Telefone")
        
    estado_civil = st.selectbox("Estado Civil", ["Selecione...", "SOLTEIRO(A)", "CASADO(A)", "DIVORCIADO(A)", "VIÚVO(A)"])

    st.divider()

    st.subheader("Endereço")
    cep = st.text_input("CEP")
    endereco = st.text_input("Endereço")
    
    col5, col6 = st.columns(2)
    with col5:
        municipio = st.text_input("Município")
    with col6:
        estado = st.text_input("Estado (UF)")

    btn_salvar = st.form_submit_button("Salvar e Gerar Documentos")

if btn_salvar:
    if nome == "" or cpf == "":
        st.error("⚠️ Por favor, preencha pelo menos o Nome e o CPF!")
    else:
        dados_usuario = {
            "nome": nome, "cpf": cpf, "rg": rg, "estado_civil": estado_civil,
            "orgao": orgao, "cargo": cargo, "matricula": matricula, 
            "data_ingresso": data_ingresso.strftime("%d/%m/%Y"),
            "endereco": endereco, "municipio": municipio, "estado": estado,
            "cep": cep, "telefone": telefone, "email": email,
            "data_cadastro": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        }
        
        pasta_pdfs = os.path.join(PASTA_ATUAL, "Documentos_Gerados")
        os.makedirs(pasta_pdfs, exist_ok=True)
        
        salvar_no_excel(dados_usuario)
        preencher_documentos_oficiais(dados_usuario, pasta_pdfs)
        
        st.success(f"✅ Sucesso! Os documentos de {nome} foram gerados na pasta Documentos_Gerados.")