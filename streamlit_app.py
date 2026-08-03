import os
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import requests
import base64
import json

# Configuração da página para ocupar melhor a largura e mudar o título da aba
st.set_page_config(
    page_title="Cadastro - Ação Correção Monetária",
    page_icon="📋",
    layout="wide"
)

# Estilização visual customizada via CSS para deixar os botões e caixas mais modernos
st.markdown("""
    <style>
        .main {
            background-color: #f8f9fa;
        }
        .stButton>button {
            width: 100%;
            border-radius: 6px;
            font-weight: bold;
            height: 3em;
            background-color: #0d6efd;
            color: white;
        }
        .stButton>button:hover {
            background-color: #0b5ed7;
            color: white;
        }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "Cadastros_Servidores.xlsx"
GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz8lQ3xhchTyl2QrvmIYr9qVZIFsx_8I2hIb0-jBqHOX63G8OzExrHPr2OlROfn_hSZ/exec"

def salvar_no_excel(dados):
    df_novo = pd.DataFrame([dados])
    if os.path.exists(EXCEL_FILE):
        df_existente = pd.read_excel(EXCEL_FILE)
        df_final = pd.concat([df_existente, df_novo], ignore_index=True)
    else:
        df_final = df_novo
    df_final.to_excel(EXCEL_FILE, index=False)

def enviar_para_google_drive(nome_servidor, lista_arquivos):
    payload = {
        "nomeServidor": nome_servidor,
        "arquivos": lista_arquivos
    }
    try:
        headers = {"Content-Type": "application/json"}
        response = requests.post(GOOGLE_APPS_SCRIPT_URL, data=json.dumps(payload), headers=headers)
        resultado = response.json()
        return resultado.get("status") == "sucesso"
    except Exception as e:
        print(f"Erro ao conectar com o Google Drive: {e}")
        return False

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
        pag_proc.insert_text((368, 200), dados['Cargo'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((94, 216), dados['Órgão'], fontsize=8, color=(0,0,0))
        pag_proc.insert_text((284, 216), dados['Data de Ingresso'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((428, 216), dados['Estado Civil'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((109, 230), dados['Telefone'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((253, 230), dados['E-mail'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((116, 245), dados['Endereço'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((99, 260), dados['Município'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((392, 260), dados['Estado'], fontsize=9, color=(0,0,0))
        pag_proc.insert_text((457, 260), dados['CEP'], fontsize=9, color=(0,0,0))
        pdf_procuracao_bytes = doc_proc.tobytes()
        doc_proc.close()

    if os.path.exists(caminho_termo):
        doc_termo = fitz.open(caminho_termo)
        pag_termo = doc_termo[0]
        pag_termo.insert_text((125, 145), dados['Nome'], fontsize=9, color=(0,0,0))
        pag_termo.insert_text((82, 170), dados['CPF'], fontsize=9, color=(0,0,0))
        pag_termo.insert_text((265, 170), dados['Matrícula'], fontsize=9, color=(0,0,0))
        pag_termo.insert_text((377, 169), dados['Cargo'], fontsize=9, color=(0,0,0))
        pdf_termo_bytes = doc_termo.tobytes()
        doc_termo.close()

    return pdf_procuracao_bytes, pdf_termo_bytes

# --- CABEÇALHO PRINCIPAL ---
st.title("⚖️ Sistema de Cadastro e Gestão de Documentos")
st.markdown("##### **Ação de Correção Monetária de Exercícios Anteriores**")
st.markdown("---")

with st.sidebar:
    st.image("https://img.icons8.com/color/96/law.png", width=80)
    st.header("Navegação e Ajuda")
    st.info(
        "💡 **Dica:** Abra este link em outro aparelho para "
        "visualizar o tutorial em vídeo/imagens sem sair desta tela:\n\n"
        "🔗 [Acessar Tutorial Completo](https://blank-app-8vxh0tfzj3.streamlit.app/)"
    )
    st.markdown("---")
    st.markdown("🔒 *Ambiente Seguro e Integrado ao Google Drive*")

if "pdf_proc" not in st.session_state:
    st.session_state.pdf_proc = None
if "pdf_termo" not in st.session_state:
    st.session_state.pdf_termo = None

# --- FORMULÁRIO DE CADASTRO ORGANIZADO EM COLUNAS ---
with st.container():
    st.subheader("📝 1. Preenchimento de Dados Cadastrais")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.markdown("**Dados Profissionais**")
        matricula = st.text_input("Matrícula (SIAPE)")
        cargo = st.text_input("Cargo")
        orgao = st.text_input("Órgão")
        ingresso = st.text_input("Data de Ingresso", placeholder="DD/MM/AAAA")

    with col_p2:
        st.markdown("**Dados Pessoais**")
        nome = st.text_input("Nome Completo")
        cpf = st.text_input("CPF")
        rg = st.text_input("RG")
        email = st.text_input("E-mail")

    col_p3, col_p4, col_p5 = st.columns(3)
    with col_p3:
        telefone = st.text_input("Telefone")
        estado_civil = st.text_input("Estado Civil")
    with col_p4:
        cep = st.text_input("CEP")
        endereco = st.text_input("Endereço")
    with col_p5:
        municipio = st.text_input("Município")
        estado = st.text_input("Estado (UF)")

    st.markdown("")
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        btn_salvar = st.button("💾 Salvar e Gerar Documentos Oficiais")

if btn_salvar:
    if not nome.strip():
        st.error("⚠️ Por favor, preencha o campo 'Nome Completo' antes de continuar.")
    else:
        dados_usuario = {
            "Matrícula": matricula, "Cargo": cargo, "Órgão": orgao, "Data de Ingresso": ingresso,
            "Nome": nome, "CPF": cpf, "E-mail": email, "RG": rg, "Telefone": telefone,
            "Estado Civil": estado_civil, "CEP": cep, "Endereço": endereco, "Município": municipio, "Estado": estado
        }
        salvar_no_excel(dados_usuario)
        st.session_state.dados_usuario = dados_usuario
        st.session_state.nome_servidor = nome.strip()
        
        st.session_state.pdf_proc, st.session_state.pdf_termo = preencher_documentos_oficiais(dados_usuario)
        st.success(f"✨ Documentos gerados com sucesso para **{st.session_state.nome_servidor}**!")

# --- SEÇÃO DE DOWNLOAD E UPLOAD (SÓ APARECE APÓS GERAR) ---
if "nome_servidor" in st.session_state or st.session_state.pdf_proc or st.session_state.pdf_termo:
    if "nome_servidor" not in st.session_state:
        st.session_state.nome_servidor = nome.strip() if nome.strip() else "Servidor_Sem_Nome"

    st.markdown("---")
    st.subheader("📥 2. Download dos Documentos Preenchidos")
    st.write("Baixe os arquivos abaixo, realize a conferência e as devidas assinaturas conforme o tutorial.")
    
    col_dl1, col_dl2 = st.columns(2)
    if st.session_state.pdf_proc:
        with col_dl1:
            st.download_button(label="📄 Baixar Procuração Preenchida", data=st.session_state.pdf_proc, file_name="Procuracao_Preenchida.pdf", mime="application/pdf")
    if st.session_state.pdf_termo:
        with col_dl2:
            st.download_button(label="📄 Baixar Termo Preenchido", data=st.session_state.pdf_termo, file_name="Termo_Preenchido.pdf", mime="application/pdf")

    st.markdown("---")
    st.subheader("📤 3. Anexo de Documentos Digitalizados e Assinados")
    st.write("Envie os arquivos digitalizados. O sistema organizará tudo automaticamente no Google Drive.")
    
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        doc_identidade = st.file_uploader("🪪 Documento de Identidade com Foto", type=["pdf", "jpg", "jpeg", "png"], key="up_identidade")
        upload_proc_assinada = st.file_uploader("✍️ Procuração Assinada", type=["pdf"], key="upload_proc")
    with col_up2:
        comprovante_residencia = st.file_uploader("🏠 Comprovante de Residência Atualizado", type=["pdf", "jpg", "jpeg", "png"], key="up_residencia")
        upload_termo_assinado = st.file_uploader("✍️ Termo Assinado", type=["pdf"], key="upload_termo")

    st.markdown("")
    col_env1, col_env2, col_env3 = st.columns([1, 2, 1])
    with col_env2:
        btn_enviar_drive = st.button("🚀 Enviar Documentos Diretamente para o Google Drive")

    if btn_enviar_drive:
        nome_pasta = st.session_state.nome_servidor
        lista_arquivos_payload = []

        if st.session_state.pdf_proc:
            lista_arquivos_payload.append({
                "nome": "Procuracao_Preenchida.pdf",
                "conteudo": base64.b64encode(st.session_state.pdf_proc).decode('utf-8'),
                "mimeType": "application/pdf"
            })
        if st.session_state.pdf_termo:
            lista_arquivos_payload.append({
                "nome": "Termo_Preenchido.pdf",
                "conteudo": base64.b64encode(st.session_state.pdf_termo).decode('utf-8'),
                "mimeType": "application/pdf"
            })

        for arquivo_up in [doc_identidade, comprovante_residencia, upload_proc_assinada, upload_termo_assinado]:
            if arquivo_up is not None:
                lista_arquivos_payload.append({
                    "nome": arquivo_up.name,
                    "conteudo": base64.b64encode(arquivo_up.getbuffer()).decode('utf-8'),
                    "mimeType": arquivo_up.type
                })

        if lista_arquivos_payload:
            with st.spinner("Enviando arquivos com segurança para o Google Drive... Aguarde."):
                sucesso = enviar_para_google_drive(nome_pasta, lista_arquivos_payload)
                if sucesso:
                    st.success(f"🎉 Sucesso! A pasta de **{nome_pasta}** foi criada perfeitamente dentro de 'Ação Correção Monetária de Exercícios Anteriores' no Google Drive!")
                else:
                    st.error("❌ Ocorreu um erro ao conectar com o Google Drive. Verifique a URL do Web App.")
        else:
            st.warning("⚠️ Nenhum arquivo foi anexado para envio.")

# --- TUTORIAL VISUAL AO FINAL ---
st.markdown("---")
st.subheader("📖 Passo a Passo para Envio dos Documentos")
st.info("Consulte abaixo o guia visual detalhado para realizar o processo corretamente.")

for i in range(1, 7):
    with st.expander(f"Passo {i} — Clique para visualizar a orientação"):
        caminho_img = os.path.join("imagens", f"{i}.png")
        if os.path.exists(caminho_img):
            st.image(caminho_img, width=700)
        else:
            st.warning(f"*(A imagem explicativa '{i}.png' não foi encontrada dentro da pasta 'imagens')*")