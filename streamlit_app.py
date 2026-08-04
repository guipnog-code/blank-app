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
CHAVE_ADMIN = "Sindicatojus"

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

def formatar_data_callback():
    val = st.session_state.get("input_ing_raw", "")
    digitos = "".join(filter(str.isdigit, str(val)))[:8]
    formatado = ""
    if len(digitos) > 4:
        formatado = f"{digitos[:2]}/{digitos[2:4]}/{digitos[4:]}"
    elif len(digitos) > 2:
        formatado = f"{digitos[:2]}/{digitos[2:]}"
    else:
        formatado = digitos
    st.session_state.input_ing_raw = formatado

def limpar_valor(val):
    if val is None or pd.isna(val) or str(val).strip().lower() == "nan":
        return ""
    return str(val).strip()

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

if "nome_servidor" not in st.session_state:
    st.session_state.nome_servidor = None
if "pdf_proc" not in st.session_state:
    st.session_state.pdf_proc = None
if "pdf_termo" not in st.session_state:
    st.session_state.pdf_termo = None

# Cabeçalho Principal com Métricas Rápidas
col_cab1, col_cab2 = st.columns([3, 1])
with col_cab1:
    st.title("⚖️ Sistema de Cadastro e Gestão de Documentos")
    st.markdown("##### **Ação de Correção Monetária de Exercícios Anteriores**")

with col_cab2:
    df_geral = carregar_servidores_cadastrados()
    total_cadastrados = len(df_geral) if not df_geral.empty else 0
    st.metric(label="📊 Total Cadastrados", value=total_cadastrados)

st.markdown("---")

# Barra Lateral Otimizada
with st.sidebar:
    st.header("💡 Ajuda e Navegação")
    st.info("🔗 [Acessar Tutorial Externo](https://blank-app-8vxh0tfzj3.streamlit.app/)")
    st.markdown("---")
    
    usuario_autorizado = False
    with st.expander("🔐 Área do Administrador"):
        st.markdown("*(Exclusivo para preenchimento rápido)*")
        chave_input = st.text_input("Chave de acesso:", type="password", key="input_chave_admin")
        if chave_input.strip() == CHAVE_ADMIN:
            st.success("✅ Liberado!")
            usuario_autorizado = True
        elif chave_input:
            st.error("❌ Incorreta.")

# Criação das abas nativas originais do Streamlit
aba_novo, aba_salvos, aba_tutorial = st.tabs(["➕ Novo Cadastro", "📂 Servidores Já Cadastrados", "📖 Tutorial"])

with aba_salvos:
    st.subheader("🔍 Pesquisar e Selecionar Servidor da Planilha")
    
    if usuario_autorizado:
        df_servidores = carregar_servidores_cadastrados()
        
        if df_servidores.empty:
            st.info("ℹ️ Nenhum servidor cadastrado na planilha até o momento.")
        else:
            opcoes_servidores = []
            mapa_linhas = {}
            
            for idx, row in df_servidores.iterrows():
                nome_servidor = str(row.get("Nome", "Sem Nome"))
                cpf_servidor = str(row.get("CPF", "Sem CPF"))
                rotulo = f"[{idx}] {nome_servidor} - CPF: {cpf_servidor}"
                opcoes_servidores.append(rotulo)
                mapa_linhas[rotulo] = row.to_dict()

            selecao = st.selectbox("Pesquise digitando o nome ou CPF:", ["-- Selecione --"] + opcoes_servidores, key="selectbox_servidor")
            
            if selecao != "-- Selecione --":
                dados_linha = mapa_linhas[selecao]
                
                with st.container(border=True):
                    st.write("📋 **Dados carregados da linha selecionada:**")
                    st.json(dados_linha)
                
                col_b1, col_b2 = st.columns(2)
                
                with col_b1:
                    with st.container(border=True):
                        st.markdown("##### **Gerar Documentos**")
                        if st.button("📄 Gerar Documentos para este Servidor", key="btn_gerar_salvo"):
                            st.session_state.dados_usuario = dados_linha
                            st.session_state.nome_servidor = str(dados_linha.get("Nome", "")).strip()
                            st.session_state.pdf_proc, st.session_state.pdf_termo = preencher_documentos_oficiais(dados_linha)
                            st.success(f"Documentos gerados com sucesso para {st.session_state.nome_servidor}!")

                with col_b2:
                    with st.container(border=True):
                        st.markdown("##### **Links Preenchidos**")
                        
                        base_form_geral = "https://docs.google.com/forms/d/e/1FAIpQLSfwwmAw9jqwWv2KTEWXQFMXaz36mECCCuVdYsxlLg48KkrsMQ/viewform"
                        params_geral = {
                            "entry.463599518": limpar_valor(dados_linha.get("Nome", "")),
                            "entry.1304511106": limpar_valor(dados_linha.get("Matrícula", "")),
                            "emailAddress": limpar_valor(dados_linha.get("E-mail", ""))
                        }
                        url_geral = f"{base_form_geral}?usp=pp_url&{urllib.parse.urlencode(params_geral)}"
                        st.markdown(
                            f"""<a href="{url_geral}" target="_blank">
                                <button style="width:100%; border-radius:6px; font-weight:bold; height:2.5em; background-color:#198754; color:white; border:none; cursor:pointer; margin-bottom: 8px;">
                                    📝 Termo de Consentimento
                                </button>
                            </a>""",
                            unsafe_allow_html=True
                        )

                        base_form_url = "https://docs.google.com/forms/d/e/1FAIpQLScFHB1lA_2cTeg-ANSa0TK3I4LwwMTa6T9cMnxQiWmbBD6XOw/viewform"
                        str_data = limpar_valor(dados_linha.get("Data de Ingresso", ""))
                        ano, mes, dia = "", "", ""
                        try:
                            dt_obj = datetime.strptime(str_data, "%d/%m/%Y")
                            ano, mes, dia = str(dt_obj.year), str(dt_obj.month), str(dt_obj.day)
                        except Exception:
                            pass

                        params = {
                            "entry.336229460": limpar_valor(dados_linha.get("Nome", "")),
                            "entry.1167987372": limpar_valor(dados_linha.get("Matrícula", "")),
                            "entry.918241761_year": ano,
                            "entry.918241761_month": mes,
                            "entry.918241761_day": dia,
                            "entry.304080830": limpar_valor(dados_linha.get("CPF", "")),
                            "entry.346470482": limpar_valor(dados_linha.get("RG", "")),
                            "entry.1131685604": limpar_valor(dados_linha.get("Endereço", "")),
                            "entry.713012878": limpar_valor(dados_linha.get("Município", "")),
                            "entry.1662466686": limpar_valor(dados_linha.get("Estado", "")),
                            "entry.1919228175": limpar_valor(dados_linha.get("CEP", "")),
                            "entry.1147940036": limpar_valor(dados_linha.get("Telefone", "")),
                            "entry.737384383": limpar_valor(dados_linha.get("E-mail", ""))
                        }
                        url_preenchida = f"{base_form_url}?usp=pp_url&{urllib.parse.urlencode(params)}"
                        st.markdown(
                            f"""<a href="{url_preenchida}" target="_blank">
                                <button style="width:100%; border-radius:6px; font-weight:bold; height:2.5em; background-color:#0d6efd; color:white; border:none; cursor:pointer;">
                                    📝 Forms Correção Monetária
                                </button>
                            </a>""",
                            unsafe_allow_html=True
                        )
    else:
        st.warning("🔒 **Conteúdo Restrito.** Abra a **Área do Administrador** na barra lateral e insira a chave de acesso correta.")

with aba_novo:
    col_sub_1, col_sub_2 = st.columns([3, 1])
    with col_sub_1:
        st.subheader("📝 Preenchimento de Dados Cadastrais (Livre para Uso)")
    with col_sub_2:
        # Script JS integrado que simula o clique na aba de tutorial nativa do Streamlit
        st.markdown("""
            <button onclick="
                const tabs = window.parent.document.querySelectorAll('button[data-baseweb=\\'tab\\']');
                if (tabs.length >= 3) { tabs[2].click(); }
            " style="width:100%; border-radius:6px; font-weight:bold; height:3em; background-color:#0d6efd; color:white; border:none; cursor:pointer;">
                💡 Ver Tutorial de Ajuda
            </button>
        """, unsafe_allow_html=True)

    with st.container(border=True):
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.markdown("**Dados Profissionais**")
            matricula = st.text_input("Matrícula (SIAPE)", key="input_mat")
            cargo = st.text_input("Cargo", key="input_cargo")
            orgao = st.text_input("Órgão", key="input_orgao")
            ingresso = st.text_input("Data de Ingresso", placeholder="DD/MM/AAAA", key="input_ing_raw", max_chars=10, on_change=formatar_data_callback)

        with col_p2:
            st.markdown("**Dados Pessoais**")
            nome = st.text_input("Nome Completo", key="input_nome")
            cpf = st.text_input("CPF", key="input_cpf")
            rg = st.text_input("RG - Órgão de Expedição", key="input_rg")
            email = st.text_input("E-mail", key="input_email")

        col_p3, col_p4, col_p5 = st.columns(3)
        with col_p3:
            telefone = st.text_input("Telefone", key="input_tel")
            estado_civil = st.text_input("Estado Civil", key="input_ec")
        with col_p4:
            cep = st.text_input("CEP", key="input_cep")
            endereco = st.text_input("Endereço", key="input_end")
        with col_p5:
            municipio = st.text_input("Município", key="input_mun")
            estado = st.text_input("Estado (UF)", key="input_uf")

    st.markdown("")
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        btn_salvar = st.button("💾 Salvar na Planilha e Gerar Documentos", key="btn_salvar_novo")

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

    if st.session_state.get("nome_servidor") or st.session_state.get("pdf_proc"):
        st.markdown("---")
        with st.container(border=True):
            st.subheader(f"⚙️ Gestão de Arquivos para: {st.session_state.get('nome_servidor', '')}")
            
            col_dl1, col_dl2 = st.columns(2)
            if st.session_state.get("pdf_proc"):
                with col_dl1:
                    st.download_button(label="📄 Baixar Procuração Preenchida", data=st.session_state.pdf_proc, file_name="Procuracao_Preenchida.pdf", mime="application/pdf", key="dl_proc")
            if st.session_state.get("pdf_termo"):
                with col_dl2:
                    st.download_button(label="📄 Baixar Termo Preenchido", data=st.session_state.pdf_termo, file_name="Termo_Preenchido.pdf", mime="application/pdf", key="dl_termo")

        st.markdown("---")
        with st.container(border=True):
            st.subheader("📤 Envio de Documentos para o Google Drive")
            
            col_up1, col_up2 = st.columns(2)
            with col_up1:
                doc_identidade = st.file_uploader("🪪 Documento de Identidade com Foto", type=["pdf", "jpg", "jpeg", "png"], key="up_identidade")
                upload_proc_assinada = st.file_uploader("✍️ Procuração Assinada", type=["pdf"], key="upload_proc")
            with col_up2:
                comprovante_residencia = st.file_uploader("🏠 Comprovante de Residência Atualizado", type=["pdf", "jpg", "jpeg", "png"], key="up_residencia")
                upload_termo_assinado = st.file_uploader("✍️ Termo Assinado", type=["pdf"], key="upload_termo")

            col_env1, col_env2, col_env3 = st.columns([1, 2, 1])
            with col_env2:
                btn_enviar_drive = st.button("🚀 Enviar Arquivos para o Google Drive", key="btn_enviar_drive")

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

with aba_tutorial:
    st.subheader("📖 Tutorial de Utilização do Sistema")
    st.info("Selecione abaixo o dispositivo que você está utilizando para visualizar o tutorial correspondente:")

    tipo_tutorial = st.radio(
        "Escolha a plataforma:",
        ["💻 Tutorial Computador", "📱 Tutorial Celular"],
        horizontal=True,
        label_visibility="collapsed"
    )

    st.markdown("---")

    if tipo_tutorial == "💻 Tutorial Computador":
        st.markdown("### 🖥️ Passo a Passo para Computador")
        st.write("Assista ao vídeo explicativo ou siga o guia detalhado para realizar o processo pelo computador:")
        
        # Carrega o vídeo da pasta especificada se ele existir
        caminho_video_pc = os.path.join("imagens", "Tutorial Computador", "video.mp4") # Altere o nome do arquivo caso seja diferente de video.mp4
        if os.path.exists(caminho_video_pc):
            st.video(caminho_video_pc)
        else:
            # Procura por qualquer arquivo .mp4 na pasta caso o nome exato varie
            pasta_pc = os.path.join("imagens", "Tutorial Computador")
            if os.path.exists(pasta_pc):
                videos_encontrados = [f for f in os.listdir(pasta_pc) if f.endswith(('.mp4', '.mov', '.avi'))]
                if videos_encontrados:
                    st.video(os.path.join(pasta_pc, videos_encontrados[0]))
                else:
                    st.warning("⚠️ Nenhum arquivo de vídeo encontrado na pasta 'imagens/Tutorial Computador'.")
            else:
                st.warning("⚠️ A pasta 'imagens/Tutorial Computador' não foi encontrada.")

        for i in range(1, 4):
            with st.expander(f"Passo {i} (Computador) — Clique para visualizar a orientação"):
                caminho_img = os.path.join("imagens", f"pc_{i}.png")
                if os.path.exists(caminho_img):
                    st.image(caminho_img, width=700)
                else:
                    st.warning(f"*(A imagem explicativa 'pc_{i}.png' não foi encontrada na pasta 'imagens')*")

    else:
        st.markdown("### 📱 Passo a Passo para Celular")
        st.write("Assista ao vídeo explicativo ou siga o guia detalhado para realizar o processo pelo celular:")
        
        # Carrega o vídeo da pasta de celular se existir
        caminho_video_cel = os.path.join("imagens", "Tutorial Celular", "video.mp4")
        if os.path.exists(caminho_video_cel):
            st.video(caminho_video_cel)
        else:
            pasta_cel = os.path.join("imagens", "Tutorial Celular")
            if os.path.exists(pasta_cel):
                videos_encontrados = [f for f in os.listdir(pasta_cel) if f.endswith(('.mp4', '.mov', '.avi'))]
                if videos_encontrados:
                    st.video(os.path.join(pasta_cel, videos_encontrados[0]))
                else:
                    st.warning("⚠️ Nenhum arquivo de vídeo encontrado na pasta 'imagens/Tutorial Celular'.")
            else:
                st.warning("⚠️ A pasta 'imagens/Tutorial Celular' não foi encontrada.")

        for i in range(1, 4):
            with st.expander(f"Passo {i} (Celular) — Clique para visualizar a orientação"):
                caminho_img = os.path.join("imagens", f"cel_{i}.png")
                if os.path.exists(caminho_img):
                    st.image(caminho_img, width=700)
                else:
                    st.warning(f"*(A imagem explicativa 'cel_{i}.png' não foi encontrada na pasta 'imagens')*")