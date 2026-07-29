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

    # Tutorial Passo a Passo - CORRIGIDO
    st.markdown("---")
    st.markdown("## 📖 Tutorial: Como Assinar Digitalmente e Enviar")
    st.write("Siga o passo a passo ilustrado abaixo:")

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
        for ext in [".jpeg", ".jpg", ".JPEG", ".JPG"]:
            caminho_teste = os.path.join("imagens", nome_base + ext)
            if os.path.exists(caminho_teste):
                arquivo_encontrado = caminho_teste
                break
                
        if arquivo_encontrado:
            st.image(arquivo_encontrado, use_container_width=True)
        else:
            st.info(f"*(Imagem não encontrada na pasta 'imagens' para este passo)*")
        st.markdown("---")