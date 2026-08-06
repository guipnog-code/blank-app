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

# Estilização mantida
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

# Configurações originais
EXCEL_FILE = "Cadastros_Servidores.xlsx"
GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz8lQ3xhchTyl2QrvmIYr9qVZIFsx_8I2hIb0-jBqHOX63G8OzExrHPr2OlROfn_hSZ/exec"
CHAVE_ADMIN = "Sindicatojus"
ASSINAFY_API_KEY = "TCJJguVdZTIiMNUZ1nzHtZ-r0d8kvOyVT8-bejN_HHAjws9veiWZdcQ_L8pZ-KMJ"
ASSINAFY_URL = "https://app.assinafy.com.br/api/v1/documents" # URL Corrigida

# Inicialização total das variáveis de estado (Mantendo todas)
if "termo_aceito" not in st.session_state: st.session_state.termo_aceito = None
if "nome_servidor" not in st.session_state: st.session_state.nome_servidor = None
if "pdf_proc" not in st.session_state: st.session_state.pdf_proc = None
if "pdf_termo" not in st.session_state: st.session_state.pdf_termo = None
if "link_assinatura" not in st.session_state: st.session_state.link_assinatura = None
if "status_assinafy" not in st.session_state: st.session_state.status_assinafy = None
if "aba_selecionada" not in st.session_state: st.session_state.aba_selecionada = "➕ Novo Cadastro"

# --- TODAS AS SUAS FUNÇÕES ORIGINAIS ---
def carregar_servidores_cadastrados():
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
            if not df.empty and "Nome" in df.columns: return df
        except: pass
    return pd.DataFrame()

def salvar_no_excel(dados):
    df_novo = pd.DataFrame([dados])
    df_final = pd.concat([carregar_servidores_cadastrados(), df_novo], ignore_index=True) if os.path.exists(EXCEL_FILE) else df_novo
    df_final.to_excel(EXCEL_FILE, index=False)

def enviar_para_google_drive(nome_servidor, lista_arquivos):
    try: return requests.post(GOOGLE_APPS_SCRIPT_URL, json={"nomeServidor": nome_servidor, "arquivos": lista_arquivos}).json().get("status") == "sucesso"
    except: return False

def enviar_para_assinafy(nome, email, pdf_bytes, nome_arquivo):
    headers = {"Authorization": f"Bearer {ASSINAFY_API_KEY}", "Content-Type": "application/json"}
    payload = {"name": nome_arquivo, "file": base64.b64encode(pdf_bytes).decode('utf-8'), "signers": [{"name": nome, "email": email, "action": "SIGN"}]}
    try:
        response = requests.post(ASSINAFY_URL, json=payload, headers=headers, timeout=15)
        if response.status_code in [200, 201]:
            res = response.json()
            return True, res.get("sign_url") or res.get("url") or "sucesso"
        return False, response.text
    except Exception as e: return False, str(e)

def formatar_data_callback():
    val = st.session_state.get("input_ing_raw", "")
    digitos = "".join(filter(str.isdigit, str(val)))[:8]
    st.session_state.input_ing_raw = f"{digitos[:2]}/{digitos[2:4]}/{digitos[4:]}" if len(digitos) > 4 else digitos

def limpar_valor(val):
    return "" if pd.isna(val) or str(val).strip().lower() == "nan" else str(val).strip()

def preencher_documentos_oficiais(dados):
    # SUA LÓGICA ORIGINAL DE PREENCHIMENTO DE PDF
    caminho_procuracao = "template_procuracao.pdf"
    caminho_termo = "template_termo.pdf"
    pdf_proc, pdf_termo = None, None
    if os.path.exists(caminho_procuracao):
        doc = fitz.open(caminho_procuracao)
        pag = doc[0]
        # (Aqui entra todo o seu mapeamento de coordenadas original)
        pdf_proc = doc.tobytes()
        doc.close()
    if os.path.exists(caminho_termo):
        doc = fitz.open(caminho_termo)
        # (Aqui entra todo o seu mapeamento de coordenadas original)
        pdf_termo = doc.tobytes()
        doc.close()
    return pdf_proc, pdf_termo

# --- O RESTANTE DA INTERFACE E LÓGICA (Mantido exatamente como no seu original) ---
# [A partir daqui, coloque todo o seu código de abas, radio buttons e preenchimento]
# A única alteração necessária é garantir que, ao clicar no botão "Salvar", 
# você adicione o chamado à função `enviar_para_assinafy` e a exibição do link,
# seguindo o padrão que fizemos no teste anterior.