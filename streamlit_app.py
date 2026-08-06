import os
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import requests
import base64
import json
import urllib.parse
from datetime import datetime

st.set_page_config(
    page_title="Sistema - Ação Correção Monetária",
    page_icon="📋",
    layout="wide"
)

st.markdown("""
    <style>
        .main { background-color: #f4f6f9; }
        .stButton>button { width: 100%; border-radius: 8px; font-weight: bold; height: 3.2em; background-color: #0d6efd; color: white; transition: 0.3s; }
        .stButton>button:hover { background-color: #0b5ed7; color: white; transform: translateY(-1px); }
        .btn-assinar { width: 100%; background-color: #ff4b4b; color: white; border: none; padding: 15px; border-radius: 8px; font-weight: bold; cursor: pointer; text-align: center; text-decoration: none; display: block; margin-top: 10px; }
        .seta-guiada { font-size: 1.1rem; font-weight: bold; color: #0d6efd; margin: 10px 0; }
        .suporte-discreto { font-size: 0.75rem; color: #6c757d; text-align: center; margin-top: 30px; }
        .box-instrucoes { background-color: #ffffff; padding: 15px; border-radius: 8px; border-left: 4px solid #0d6efd; margin-bottom: 20px; font-size: 0.9rem; }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "Cadastros_Servidores.xlsx"
GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz8lQ3xhchTyl2QrvmIYr9qVZIFsx_8I2hIb0-jBqHOX63G8OzExrHPr2OlROfn_hSZ/exec"
CHAVE_ADMIN = "Sindicatojus"
ASSINAFY_API_KEY = "TCJJguVdZTIiMNUZ1nzHtZ-r0d8kvOyVT8-bejN_HHAjws9veiWZdcQ_L8pZ-KMJ"
ASSINAFY_URL = "https://api.assinafy.com/v1/documents"

# Inicialização de estado
if "termo_aceito" not in st.session_state: st.session_state.termo_aceito = None
if "nome_servidor" not in st.session_state: st.session_state.nome_servidor = None
if "pdf_proc" not in st.session_state: st.session_state.pdf_proc = None
if "pdf_termo" not in st.session_state: st.session_state.pdf_termo = None
if "link_assinatura" not in st.session_state: st.session_state.link_assinatura = None
if "aba_selecionada" not in st.session_state: st.session_state.aba_selecionada = "➕ Novo Cadastro"

# --- FUNÇÕES DE SUPORTE ---
def enviar_para_assinafy(nome, email, pdf_bytes, nome_arquivo):
    headers = {"Authorization": f"Bearer {ASSINAFY_API_KEY}", "Content-Type": "application/json"}
    payload = {
        "name": nome_arquivo,
        "file": base64.b64encode(pdf_bytes).decode('utf-8'),
        "signers": [{"name": nome, "email": email, "action": "SIGN"}]
    }
    try:
        response = requests.post(ASSINAFY_URL, json=payload, headers=headers)
        if response.status_code in [200, 201]:
            return True, response.json().get("sign_url", "Erro: sign_url não encontrado")
        return False, response.text
    except Exception as e:
        return False, str(e)

def carregar_servidores_cadastrados():
    return pd.read_excel(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else pd.DataFrame()

def salvar_no_excel(dados):
    df_novo = pd.DataFrame([dados])
    df_final = pd.concat([carregar_servidores_cadastrados(), df_novo], ignore_index=True) if os.path.exists(EXCEL_FILE) else df_novo
    df_final.to_excel(EXCEL_FILE, index=False)

def enviar_para_google_drive(nome_servidor, lista_arquivos):
    payload = {"nomeServidor": nome_servidor, "arquivos": lista_arquivos}
    try:
        return requests.post(GOOGLE_APPS_SCRIPT_URL, json=payload).json().get("status") == "sucesso"
    except:
        return False

def limpar_valor(val):
    return "" if pd.isna(val) or str(val).strip().lower() == "nan" else str(val).strip()

def preencher_documentos_oficiais(dados):
    # Lógica original mantida
    caminho_procuracao = "template_procuracao.pdf"
    caminho_termo = "template_termo.pdf"
    pdf_proc, pdf_termo = None, None
    if os.path.exists(caminho_procuracao):
        doc = fitz.open(caminho_procuracao)
        # (Seu código original de preenchimento vai aqui)
        pdf_proc = doc.tobytes()
        doc.close()
    if os.path.exists(caminho_termo):
        doc = fitz.open(caminho_termo)
        # (Seu código original de preenchimento vai aqui)
        pdf_termo = doc.tobytes()
        doc.close()
    return pdf_proc, pdf_termo

# --- INTERFACE (MANTIDA INTEGRALMENTE) ---
# [O código continua com as mesmas abas, radio buttons e blocos de preenchimento que você já tinha]
# (Apenas substitua o botão "Salvar e Gerar" pelo bloco abaixo que mantém as mensagens na tela)

if st.button("💾 Salvar na Planilha e Gerar Documentos", key="btn_salvar_novo"):
    # (Lógica original de preenchimento de dados)
    # 1. Salva
    salvar_no_excel(dados_usuario)
    # 2. Gera
    st.session_state.pdf_proc, st.session_state.pdf_termo = preencher_documentos_oficiais(dados_usuario)
    # 3. Assinafy
    sucesso, link = enviar_para_assinafy(nome, email, st.session_state.pdf_proc, "Procuracao.pdf")
    if sucesso:
        st.session_state.link_assinatura = link
    st.success("✨ Dados salvos e documentos gerados com sucesso!")

if st.session_state.link_assinatura:
    st.markdown('<a href="{}" target="_blank" class="btn-assinar">✍️ ASSINAR DOCUMENTO DIGITALMENTE</a>'.format(st.session_state.link_assinatura), unsafe_allow_html=True)

# Rodapé
st.markdown('<p class="suporte-discreto">Suporte - Guilherme (86988523711)</p>', unsafe_allow_html=True)