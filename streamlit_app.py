import os
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import requests
import base64
import json

st.set_page_config(
    page_title="Sistema - Ação Correção Monetária",
    page_icon="📋",
    layout="wide"
)

st.markdown("""
    <style>
        .main { background-color: #f8f9fa; }
        .stButton>button {
            width: 100%;
            border-radius: 6px;
            font-weight: bold;
            height: 3em;
            background-color: #0d6efd;
            color: white;
        }
        .stButton>button:hover { background-color: #0b5ed7; color: white; }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "Cadastros_Servidores.xlsx"
GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz8lQ3xhchTyl2QrvmIYr9qVZIFsx_8I2hIb0-jBqHOX63G8OzExrHPr2OlROfn_hSZ/exec"

# ==========================================
# CHAVE DE ACESSO DEFINIDA PELO ADMINISTRADOR
# ==========================================
CHAVE_ADMIN = "Sindicatojus"

def verificar_admin():
    """Valida se a chave de acesso digitada na barra lateral está correta"""
    st.sidebar.markdown("### 🔐 Acesso Restrito")
    st.sidebar.markdown("*(Exclusivo para preenchimento rápido)*")
    
    # Campo de senha para esconder o texto digitado
    chave_input = st.sidebar.text_input("Digite a chave de acesso:", type="password")
    
    if not chave_input:
        st.sidebar.info("💡 Insira a chave para desbloquear a aba de Servidores Cadastrados.")
        return False
    
    if chave_input.strip() == CHAVE_ADMIN:
        st.sidebar.success("✅ Acesso Liberado!")
        return True
    else:
        st.sidebar.error("❌ Chave incorreta.")
        return False

def carregar_servidores_cadastrados():
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
            if not df.empty and "Nome" in df.columns:
                return df
        except Exception:
            pass
    return pd.DataFrame()

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
        print(f"Erro ao conectar: {e}")
        return False

def preencher_documentos_oficiais(dados):
    caminho_procuracao = "template_procuracao.pdf"
    caminho_termo = "template_termo.pdf"
    
    pdf_procuracao_bytes = None
    pdf_termo_bytes = None

    if os.path.exists(caminho_procuracao):
        doc_proc = fitz.open(caminho_procuracao)
        pag_proc = doc_proc[0]
        pag_proc.insert_text((92, 184), str(dados.get('Nome', '')), fontsize=9, color=(0,0,0))
        pag_proc.insert_text((83, 200), str(dados.get('CPF', '')), fontsize=9, color=(0,0,0))
        pag_proc.insert_text((223, 200), str(dados.get('RG', '')), fontsize=9, color=(0,0,0))
        pag_proc.insert_text((368, 200), str(dados.get('Cargo', '')), fontsize=9, color=(0,0,0))
        pag_proc.insert_text((94, 216), str(dados.get('Órgão', '')), fontsize=8, color=(0,0,0))
        pag_proc.insert_text((284, 216), str(dados.get('Data de Ingresso', '')), fontsize=9, color=(0,0,0))
        pag_proc.insert_text((428, 216), str(dados.get('Estado Civil', '')), fontsize=9, color=(0,0,0))
        pag_proc.insert_text((109, 230), str(dados.get('Telefone', '')), fontsize=9, color=(0,0,0))
        pag_proc.insert_text((253, 230), str(dados.get('E-mail', '')), fontsize=9, color=(0,0,0))
        pag_proc.insert_text((116, 245), str(dados.get('Endereço', '')), fontsize=9, color=(0,0,0))
        pag_proc.insert_text((99, 260), str(dados.get('Município', '')), fontsize=9, color=(0,0,0))
        pag_proc.insert_text((392, 260), str(dados.get('Estado', '')), fontsize=9, color=(0,0,0))
        pag_proc.insert_text((457, 260), str(dados.get('CEP', '')), fontsize=9, color=(0,0,0))
        pdf_procuracao_bytes = doc_proc.tobytes()
        doc_proc.close()

    if os.path.exists(caminho_termo):
        doc_termo = fitz.open(caminho_termo)
        pag_termo = doc_termo[0]
        pag_termo.insert_text((125, 145), str(dados.get('Nome', '')), fontsize=9, color=(0,0,0))
        pag_termo.insert_text((82, 170), str(dados.get('CPF', '')), fontsize=9, color=(0,0,0))
        pag_termo.insert_text((265, 170), str(dados.get('Matrícula', '')), fontsize=9, color=(0,0,0))
        pag_termo.insert_text((377, 169), str(dados.get('Cargo', '')), fontsize=9, color=(0,0,0))
        pdf_termo_bytes = doc_termo.tobytes()
        doc_termo.close()

    return pdf_procuracao_bytes, pdf_termo_bytes

st.title("⚖️ Sistema de Cadastro e Gestão de Documentos")
st.markdown("##### **Ação de Correção Monetária de Exercícios Anteriores**")
st.markdown("---")

# Barra lateral com opção de esconder e chave de acesso restrito
with st.sidebar:
    st.header("💡 Ajuda e Navegação")
    st.info(
        "🔗 [Acessar Tutorial](https://blank-app-8vxh0tfzj3.streamlit.app/)"
    )
    st.markdown("---")
    usuario_autorizado = verificar_admin()

# ABAS DO SITE
aba_novo, aba_salvos = st.tabs(["➕ Novo Cadastro", "📂 Servidores Já Cadastrados (Preenchimento Rápido)"])

with aba_salvos:
    st.subheader("🔍 Selecionar Servidor da Planilha")
    
    if usuario_autorizado:
        df_servidores = carregar_servidores_cadastrados()
        
        if df_servidores.empty:
            st.info("ℹ️ Nenhum servidor cadastrado na planilha até o momento.")
        else:
            lista_nomes = df_servidores["Nome"].dropna().unique().tolist()
            servidor_selecionado = st.selectbox("Escolha o Servidor:", ["-- Selecione --"] + lista_nomes)
            
            if servidor_selecionado != "-- Selecione --":
                dados_linha = df_servidores[df_servidores["Nome"] == servidor_selecionado].iloc[0].to_dict()
                st.write("📋 **Dados carregados da planilha:**")
                st.json(dados_linha)
                
                if st.button("📄 Gerar Documentos para este Servidor"):
                    st.session_state.dados_usuario = dados_linha
                    st.session_state.nome_servidor = str(dados_linha.get("Nome", "")).strip()
                    st.session_state.pdf_proc, st.session_state.pdf_termo = preencher_documentos_oficiais(dados_linha)
                    st.success(f"Documentos gerados com sucesso para {st.session_state.nome_servidor}!")
    else:
        st.warning("🔒 **Conteúdo Restrito.** Insira a chave de acesso correta na barra lateral (ícone de seta no canto superior esquerdo) para visualizar o menu de preenchimento rápido.")

with aba_novo:
    st.subheader("📝 Preenchimento de Dados Cadastrais (Livre para Uso)")

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
        btn_salvar = st.button("💾 Salvar na Planilha e Gerar Documentos")

    if btn_salvar:
        if not nome.strip():
            st.error("⚠️ Por favor, preencha o campo 'Nome Completo'.")
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
            st.success(f"✨ Dados salvos na planilha e documentos gerados para **{st.session_state.nome_servidor}**!")

if "nome_servidor" in st.session_state or "pdf_proc" in st.session_state:
    st.markdown("---")
    st.subheader(f"⚙️ Gestão de Arquivos para: {st.session_state.get('nome_servidor', '')}")
    
    col_dl1, col_dl2 = st.columns(2)
    if st.session_state.get("pdf_proc"):
        with col_dl1:
            st.download_button(label="📄 Baixar Procuração Preenchida", data=st.session_state.pdf_proc, file_name="Procuracao_Preenchida.pdf", mime="application/pdf")
    if st.session_state.get("pdf_termo"):
        with col_dl2:
            st.download_button(label="📄 Baixar Termo Preenchido", data=st.session_state.pdf_termo, file_name="Termo_Preenchido.pdf", mime="application/pdf")

    st.markdown("---")
    st.subheader("📤 Envio de Documentos para o Google Drive")
    
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
        btn_enviar_drive = st.button("🚀 Enviar Arquivos para o Google Drive")

    if btn_enviar_drive:
        nome_pasta = st.session_state.get("nome_servidor", "Servidor")
        lista_arquivos_payload = []

        if st.session_state.get("pdf_proc"):
            lista_arquivos_payload.append({
                "nome": "Procuracao_Preenchida.pdf",
                "conteudo": base64.b64encode(st.session_state.pdf_proc).decode('utf-8'),
                "mimeType": "application/pdf"
            })
        if st.session_state.get("pdf_termo"):
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
            with st.spinner("Enviando arquivos para o Google Drive..."):
                sucesso = enviar_para_google_drive(nome_pasta, lista_arquivos_payload)
                if sucesso:
                    st.success(f"🎉 Sucesso! A pasta de **{nome_pasta}** foi atualizada no Google Drive.")
                else:
                    st.error("❌ Erro ao enviar para o Google Drive.")
        else:
            st.warning("⚠️ Nenhum arquivo anexado.")

# --- SEÇÃO DE TUTORIAL ---
st.markdown("---")
st.subheader("📖 Tutorial para Envio dos Documentos")
st.info(
    "Consulte abaixo o passo a passo ilustrado para realizar o preenchimento, "
    "as devidas assinaturas e o envio correto dos documentos."
)

for i in range(1, 7):
    with st.expander(f"Passo {i} — Clique para visualizar a orientação"):
        caminho_img = os.path.join("imagens", f"{i}.png")
        if os.path.exists(caminho_img):
            st.image(caminho_img, width=700)
        else:
            st.warning(f"*(A imagem explicativa '{i}.png' não foi encontrada dentro da pasta 'imagens')*")
        st.markdown("---")