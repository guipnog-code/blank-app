import os
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import requests
import base64
import json

EXCEL_FILE = "Cadastros_Servidores.xlsx"

# Sua URL do Google Apps Script integrada
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

    def preencher_form_pdf(caminho_template, campos_valores):
        if not os.path.exists(caminho_template):
            return None
        doc = fitz.open(caminho_template)
        for pagina in doc:
            for widget in pagina.widgets():
                nome_campo = widget.field_name
                if nome_campo in campos_valores:
                    widget.field_value = str(campos_valores[nome_campo])
                    widget.update()
        bytes_pdf = doc.tobytes()
        doc.close()
        return bytes_pdf

    mapeamento_procuracao = {
        'Nome': dados['Nome'], 'CPF': dados['CPF'], 'RG': dados['RG'],
        'Cargo': dados['Cargo'], 'Órgão': dados['Órgão'], 'Data de Ingresso': dados['Data de Ingresso'],
        'Estado Civil': dados['Estado Civil'], 'Telefone': dados['Telefone'], 'E-mail': dados['E-mail'],
        'Endereço': dados['Endereço'], 'Município': dados['Município'], 'Estado': dados['Estado'], 'CEP': dados['CEP']
    }

    mapeamento_termo = {
        'Nome': dados['Nome'], 'CPF': dados['CPF'], 'Matrícula': dados['Matrícula'], 'Cargo': dados['Cargo']
    }

    pdf_procuracao_bytes = preencher_form_pdf(caminho_procuracao, mapeamento_procuracao)
    pdf_termo_bytes = preencher_form_pdf(caminho_termo, mapeamento_termo)

    return pdf_procuracao_bytes, pdf_termo_bytes

# --- INTERFACE DO STREAMLIT ---
st.title("📋 Cadastro e Preenchimento de Documentos")
st.info("ℹ️ **Esses dados serão direcionados a uma planilha e salvos no Google Drive.**")
st.write("Preencha os dados abaixo para cadastrar e gerar os documentos em PDF mantendo os campos editáveis.")

with st.sidebar:
    st.header("💡 Dica Importante")
    st.info(
        "O processo será muito mais fácil caso abra este link em outro aparelho "
        "para que possa visualizar o tutorial sem precisar mudar de tela:\n\n"
        "🔗 [Acessar Tutorial](https://blank-app-8vxh0tfzj3.streamlit.app/)"
    )

if "pdf_proc" not in st.session_state:
    st.session_state.pdf_proc = None
if "pdf_termo" not in st.session_state:
    st.session_state.pdf_termo = None

# Campos do Formulário
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

if st.button("Salvar e Gerar Documentos"):
    if not nome.strip():
        st.error("Por favor, preencha o campo 'Nome Completo'.")
    else:
        dados_usuario = {
            "Matrícula": matricula, "Cargo": cargo, "Órgão": orgao, "Data de Ingresso": ingresso,
            "Nome": nome, "CPF": cpf, "E-mail": email, "RG": rg, "Telefone": telefone,
            "Estado Civil": estado_civil, "CEP": cep, "Endereço": endereco, "Município": municipio, "Estado": estado
        }
        salvar_no_excel(dados_usuario)
        st.session_state.dados_usuario = dados_usuario
        st.session_state.nome_servidor = nome.strip()
        
        # Gera os PDFs preservando a editabilidade dos campos
        st.session_state.pdf_proc, st.session_state.pdf_termo = preencher_documentos_oficiais(dados_usuario)
        st.success(f"Dados salvos e documentos gerados para '{st.session_state.nome_servidor}' com sucesso!")

if "nome_servidor" in st.session_state or st.session_state.pdf_proc or st.session_state.pdf_termo:
    if "nome_servidor" not in st.session_state:
        st.session_state.nome_servidor = nome.strip() if nome.strip() else "Servidor_Sem_Nome"

    st.markdown("---")
    st.markdown("### 📥 Baixar Documentos Gerados (Editáveis)")
    col1, col2 = st.columns(2)
    
    if st.session_state.pdf_proc:
        with col1:
            st.download_button(label="📄 Baixar Procuração", data=st.session_state.pdf_proc, file_name="Procuracao_Preenchida.pdf", mime="application/pdf")
    if st.session_state.pdf_termo:
        with col2:
            st.download_button(label="📄 Baixar Termo", data=st.session_state.pdf_termo, file_name="Termo_Preenchido.pdf", mime="application/pdf")

    st.markdown("---")
    st.markdown("### 📤 Anexo de Documentos e Envio para o Google Drive")
    
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        doc_identidade = st.file_uploader("Documento de Identidade com Foto", type=["pdf", "jpg", "jpeg", "png"], key="up_identidade")
    with col_up2:
        comprovante_residencia = st.file_uploader("Comprovante de Residência Atualizado", type=["pdf", "jpg", "jpeg", "png"], key="up_residencia")

    col_up3, col_up4 = st.columns(2)
    with col_up3:
        upload_proc_assinada = st.file_uploader("Enviar Procuração Assinada", type=["pdf"], key="upload_proc")
    with col_up4:
        upload_termo_assinado = st.file_uploader("Enviar Termo Assinado", type=["pdf"], key="upload_termo")

    if st.button("🚀 Enviar Documentos Diretamente para o Google Drive"):
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
            with st.spinner("Enviando arquivos para o Google Drive..."):
                sucesso = enviar_para_google_drive(nome_pasta, lista_arquivos_payload)
                if sucesso:
                    st.success(f"🎉 Sucesso! A pasta '{nome_pasta}' foi criada dentro de 'Ação Correção Monetária de Exercícios Anteriores' no Google Drive.")
                else:
                    st.error("Houve um erro ao enviar os arquivos para o Google Drive.")
        else:
            st.warning("Nenhum arquivo para enviar.")

    st.markdown("---")
    st.markdown("### Tutorial para envio dos documentos (Nesse site)")
    st.info(
        "Como enviar os documentos, o termo e a procuração devem estar assinados "
        "(tutorial está no link da aba lateral do site), o restante dos documentos "
        "devem estar legíveis e digitalizados, também visível na aba de tutorial."
    )
    
    for i in range(1, 7):
        st.subheader(f"Passo {i}")
        caminho_img = os.path.join("imagens", f"{i}.png")
        if os.path.exists(caminho_img):
            st.image(caminho_img, width=700)
        else:
            st.warning(f"*(Imagem '{i}.png' não encontrada dentro da pasta 'imagens')*")
        st.markdown("---")