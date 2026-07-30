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
    # 1. Termo de Consentimento (URL 1)
    url_form_1 = "https://docs.google.com/forms/d/e/1FAIpQLSfwwmAw9jqwWv2KTEWXQFMXaz36mECCCuVdYsxlLg48KkrsMQ/formResponse"
    payload_1 = {
        "entry.463599518": dados['Nome'],
        "entry.1304511106": dados['Matrícula']
    }

    # 2. Ação Geral - Correção Monetária (URL 2)
    # Incluindo o campo de e-mail e o parâmetro de aceite do Google Forms ("Enviar por e-mail")
    url_form_2 = "https://docs.google.com/forms/d/e/1FAIpQLScFHB1lA_2cTeg-ANSa0TK3I4LwwMTa6T9cMnxQiWmbBD6XOw/formResponse"
    payload_2 = {
        "entry.336229460": dados['Nome'],
        "entry.1167987372": dados['Matrícula'],
        "entry.918241761": "", 
        "entry.304080830": dados['CPF'],
        "entry.346470482": dados['RG'] + " / " + dados['Órgão'],
        "entry.1131685604": dados['Endereço'],
        "entry.713012878": dados['Município'],
        "entry.1662466686": dados['Estado'],
        "entry.1919228175": dados['CEP'],
        "entry.1147940036": dados['Telefone'],
        "entry.737384383": dados['E-mail'],
        "emailReceipt": "true"  # Equivalente a marcar a caixa de registro de e-mail
    }

    headers = {"User-Agent": "Mozilla/5.0"}
    
    try:
        # Envio para o Termo de Consentimento
        requests.post(url_form_1, data=payload_1, headers=headers)

        # Envio para a Ação Geral (com os documentos pessoais se houver)
        if files_data_form2:
            requests.post(url_form_2, data=payload_2, files=files_data_form2, headers=headers)
        else:
            requests.post(url_form_2, data=payload_2, headers=headers)
            
    except Exception as e:
        print(f"Erro ao enviar para os formulários: {e}")

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
st.write("Preencha os dados abaixo para cadastrar e gerar os documentos em PDF.")

# Configuração da Barra Lateral (Sidebar) com o link oficial integrado
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

    st.subheader("Documentos Adicionais")
    doc_identidade = st.file_uploader("Documento de Identidade com Foto", type=["pdf", "jpg", "jpeg", "png"], key="up_identidade")
    comprovante_residencia = st.file_uploader("Comprovante de Residência Atualizado", type=["pdf", "jpg", "jpeg", "png"], key="up_residencia")

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

    # Preparando arquivos para o formulário principal (Ação Geral)
    arquivos_envio_form2 = {}
    if doc_identidade is not None:
        arquivos_envio_form2["entry.1823246775"] = (doc_identidade.name, doc_identidade.getvalue(), doc_identidade.type)
    if comprovante_residencia is not None:
        arquivos_envio_form2["entry.97304745"] = (comprovante_residencia.name, comprovante_residencia.getvalue(), comprovante_residencia.type)

    enviar_para_google_forms(dados_usuario, files_data_form2=arquivos_envio_form2 if arquivos_envio_form2 else None)
    
    st.success("Dados salvos com sucesso e enviados para os formulários!")

    st.session_state.pdf_proc, st.session_state.pdf_termo = preencher_documentos_oficiais(dados_usuario)

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

    # Botões de Upload dos documentos assinados
    st.markdown("### 📤 Enviar Documentos Assinados Digitalmente")
    
    col_up1, col_up2 = st.columns(2)
    
    with col_up1:
        upload_proc_assinada = st.file_uploader("Enviar Procuração Assinada", type=["pdf"], key="upload_proc")
        if upload_proc_assinada is not None:
            st.success("Procuração assinada enviada com sucesso!")
            
    with col_up2:
        upload_termo_assinado = st.file_uploader("Enviar Termo Assinado", type=["pdf"], key="upload_termo")
        if upload_termo_assinado is not None:
            st.success("Termo assinado enviado com sucesso!")
            # Enviando o termo assinado direto para o Google Forms do Termo de Consentimento
            try:
                dados_aux = {"Nome": nome, "Matrícula": matricula}
                arquivo_termo_envio = {
                    "entry.1145226915": (upload_termo_assinado.name, upload_termo_assinado.getvalue(), upload_termo_assinado.type)
                }
                enviar_para_google_forms(dados_aux, files_data_form1=arquivo_termo_envio)
                st.success("Termo assinado enviado automaticamente para o Google Forms do Termo!")
            except Exception as e:
                st.error(f"Erro ao enviar termo assinado: {e}")

    st.markdown("---")
    
    st.markdown("<h2 style='margin-bottom: 0px;'>📖 Tutorial: Como Assinar Digitalmente e Enviar</h2>", unsafe_allow_html=True)
    st.markdown("<p style='margin-top: 5px; margin-bottom: 10px;'>Siga o passo a passo ilustrado abaixo:</p>", unsafe_allow_html=True)

    audio_path = "Audio_tutorial.mp4"
    if not os.path.exists(audio_path) and os.path.exists("WhatsApp Audio 2026-07-29 at 12.18.46.mp4"):
        audio_path = "WhatsApp Audio 2026-07-29 at 12.18.46.mp4"

    if os.path.exists(audio_path):
        st.video(audio_path)
        st.caption("🎧 *Dica: Você pode reproduzir o guia em áudio/vídeo e acompanhar o passo a passo.*")
    else:
        st.info("💡 *(Arquivo de áudio/vídeo do tutorial não encontrado na raiz do projeto)*")

    passos = [
        ("Passo 1: Baixar os documentos gerados", "passo_01"),
        ("Passo 2: Abrir o aplicativo Gov.br", "passo_02"),
        ("Passo 3: Localizar o serviço", "passo_03"),
        ("Passo 4: Selecionar 'Assinar documentos digitalmente'", "passo_04"),
        ("Passo 5: Clicar em 'Escolher arquivo'", "passo_05"),
        ("Passo 6: Selecionar os PDFs recentes", "passo_06"),
        ("Passo 7: Visualizar o documento carregado", "passo_07"),
        ("Passo 8: Arrastar o quadrado para a área de assinatura", "passo_08"),
        ("Passo 9: Confirmar a assinatura", "passo_09"),
        ("Passo 10: Opção de carregar outro documento", "passo_10"),
        ("Passo 11: Iniciar o processo de assinar ambos", "passo_11"),
        ("Passo 12: Autorização via notificação", "passo_12"),
        ("Passo 13: Digitar o código recebido", "passo_13"),
        ("Passo 14: Clicar em Autorizar", "passo_14"),
        ("Passo 15: Concluir etapa de assinatura", "passo_15"),
        ("Passo 16: Baixar arquivos assinados", "passo_16"),
        ("Passo 17: Menu de opções do navegador", "passo_17"),
        ("Passo 18: Abrir no navegador Chrome", "passo_18"),
        ("Passo 19: Retornar às opções", "passo_19"),
        ("Passo 20: Acessar pasta de Transferências", "passo_20"),
        ("Passo 21: Localizar os arquivos assinados", "passo_21"),
        ("Passo 22: Compartilhar os documentos via WhatsApp", "passo_22"),
    ]

    for titulo, nome_base in passos:
        st.subheader(titulo)
        
        arquivo_encontrado = None
        for pasta in ["Imagens", "imagens"]:
            for ext in [".jpeg", ".jpg", ".JPEG", ".JPG"]:
                caminho_teste = os.path.join(pasta, nome_base + ext)
                if os.path.exists(caminho_teste):
                    arquivo_encontrado = caminho_teste
                    break
            if arquivo_encontrado:
                break
                
        if arquivo_encontrado:
            st.image(arquivo_encontrado, width=400)
        else:
            st.info(f"*(Imagem '{nome_base}.jpeg' não encontrada)*")
        st.markdown("---")