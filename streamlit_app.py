import os
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import requests

EXCEL_FILE = "Cadastros_Servidores.xlsx"

def salvar_no_excel(dados):
    df_novo = pd.DataFrame([dados])
    if os.path.exists(EXCEL_FILE):
        df_existente = pd.read_excel(EXCEL_FILE)
        df_final = pd.concat([df_existente, df_novo], ignore_index=True)
    else:
        df_final = df_novo
    df_final.to_excel(EXCEL_FILE, index=False)

def enviar_para_google_forms(dados, files_data_form1=None, files_data_form2=None):
    # Vínculo com o Google Forms cortado temporariamente para testes
    pass

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

st.title("📋 Cadastro e Preenchimento de Documentos")
st.info("ℹ️ **Esses dados serão direcionados a uma planilha, para preenchimento de dados.**")
st.write("Preencha os dados abaixo para cadastrar e gerar os documentos em PDF.")

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
        "Matrícula": matricula, "Cargo": cargo, "Órgão": orgao, "Data de Ingresso": ingresso,
        "Nome": nome, "CPF": cpf, "E-mail": email, "RG": rg, "Telefone": telefone,
        "Estado Civil": estado_civil, "CEP": cep, "Endereço": endereco, "Município": municipio, "Estado": estado
    }
    salvar_no_excel(dados_usuario)
    st.session_state.dados_usuario = dados_usuario
    st.success("Dados salvos com sucesso!")
    st.session_state.pdf_proc, st.session_state.pdf_termo = preencher_documentos_oficiais(dados_usuario)

if st.session_state.pdf_proc or st.session_state.pdf_termo:
    st.markdown("---")
    st.markdown("### 📥 Baixar Documentos Gerados")
    col1, col2 = st.columns(2)
    
    if st.session_state.pdf_proc:
        with col1:
            st.download_button(label="📄 Baixar Procuração", data=st.session_state.pdf_proc, file_name="Procuracao_Preenchida.pdf", mime="application/pdf")
    if st.session_state.pdf_termo:
        with col2:
            st.download_button(label="📄 Baixar Termo", data=st.session_state.pdf_termo, file_name="Termo_Preenchido.pdf", mime="application/pdf")

    st.markdown("---")
    st.markdown("### 📤 Anexo de Documentos e Envio")
    
    # Widgets de upload fora do formulário para reter os arquivos no clique do botão de envio
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

    if st.button("🚀 Enviar Dados e Documentos para os Formulários"):
        if "dados_usuario" in st.session_state:
            dados_envio = st.session_state.dados_usuario
            nome_servidor = dados_envio["Nome"].strip()
            
            if not nome_servidor:
                nome_servidor = "Servidor_Sem_Nome"
            
            # Cria a estrutura de pastas 'Documentos Upload/Nome_Do_Servidor'
            pasta_principal = "Documentos Upload"
            os.makedirs(pasta_principal, exist_ok=True)
            
            pasta_servidor = os.path.join(pasta_principal, nome_servidor)
            os.makedirs(pasta_servidor, exist_ok=True)

            arquivos_salvos = 0

            # Salvando Identidade
            if doc_identidade is not None:
                caminho_arquivo = os.path.join(pasta_servidor, doc_identidade.name)
                with open(caminho_arquivo, "wb") as f:
                    f.write(doc_identidade.getbuffer())
                arquivos_salvos += 1

            # Salvando Comprovante de Residência
            if comprovante_residencia is not None:
                caminho_arquivo = os.path.join(pasta_servidor, comprovante_residencia.name)
                with open(caminho_arquivo, "wb") as f:
                    f.write(comprovante_residencia.getbuffer())
                arquivos_salvos += 1

            # Salvando Procuração Assinada
            if upload_proc_assinada is not None:
                caminho_arquivo = os.path.join(pasta_servidor, upload_proc_assinada.name)
                with open(caminho_arquivo, "wb") as f:
                    f.write(upload_proc_assinada.getbuffer())
                arquivos_salvos += 1

            # Salvando Termo Assinado
            if upload_termo_assinado is not None:
                caminho_arquivo = os.path.join(pasta_servidor, upload_termo_assinado.name)
                with open(caminho_arquivo, "wb") as f:
                    f.write(upload_termo_assinado.getbuffer())
                arquivos_salvos += 1

            enviar_para_google_forms(dados_envio)

            st.success(f"Sucesso! Pasta criada e {arquivos_salvos} arquivo(s) salvo(s) em: 'Documentos Upload/{nome_servidor}'")
        else:
            st.warning("Por favor, preencha e salve os dados cadastrais primeiro.")

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