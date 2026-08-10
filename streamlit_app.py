import os
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import requests
import base64
import json
import urllib.parse
import zipfile
import io
from datetime import datetime
import time
import gspread
from google.oauth2 import service_account

st.set_page_config(
    page_title="Sistema - Ação Correção Monetária",
    page_icon="📋",
    layout="wide"
)

st.markdown("""
    <style>
        .main { background-color: #f4f6f9; }
        .stButton>button {
            width: 100%;
            border-radius: 10px;
            font-weight: bold;
            height: 3.5em;
            font-size: 16px;
            background-color: #3b82f6;
            color: white;
            transition: 0.3s;
            border: none;
        }
        .stButton>button:hover { background-color: #2563eb; color: white; transform: translateY(-1px); }
        .btn-assinar { 
            width: 100%; 
            background-color: #495057; 
            color: white; 
            border: none; 
            padding: 10px 15px; 
            border-radius: 8px; 
            font-size: 15px; 
            font-weight: bold; 
            cursor: pointer; 
            text-align: center; 
            text-decoration: none; 
            display: block; 
            margin-top: 10px; 
            transition: 0.3s;
        }
        .btn-assinar:hover { background-color: #343a40; color: white; }
        .btn-whatsapp { width: 100%; background-color: #25d366; color: white; border: none; padding: 12px; border-radius: 10px; font-size: 16px; font-weight: bold; cursor: pointer; text-align: center; text-decoration: none; display: block; margin-bottom: 20px; }
        .seta-guiada { font-size: 1.1rem; font-weight: bold; color: #3b82f6; margin: 10px 0; }
        .suporte-discreto { font-size: 0.75rem; color: #6c757d; text-align: center; margin-top: 30px; }
        .box-instrucoes { background-color: #ffffff; padding: 15px; border-radius: 8px; border-left: 4px solid #3b82f6; margin-bottom: 20px; font-size: 0.9rem; }
        .box-passo { background-color: #ffffff; padding: 20px; border-radius: 12px; margin-bottom: 20px; border-left: 6px solid #3b82f6; box-shadow: 2px 2px 8px rgba(0,0,0,0.05); }
    </style>
""", unsafe_allow_html=True)

GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbz8lQ3xhchTyl2QrvmIYr9qVZIFsx_8I2hIb0-jBqHOX63G8OzExrHPr2OlROfn_hSZ/exec"
CHAVE_ADMIN = "Sindicatojus"
ASSINAFY_API_KEY = "TCJJguVdZTIiMNUZ1nzHtZ-r0d8kvOyVT8-bejN_HHAjws9veiWZdcQ_L8pZ-KMJ"

# --- FUNÇÕES DE CONEXÃO COM O GOOGLE SHEETS ---
def obter_credenciais():
    info = dict(st.secrets["google_sheets"])
    if "private_key" in info and isinstance(info["private_key"], str):
        pk = info["private_key"]
        if "\\n" in pk:
            pk = pk.replace("\\n", "\n")
        info["private_key"] = pk.strip()
    return service_account.Credentials.from_service_account_info(info)

def carregar_servidores_cadastrados():
    try:
        credenciais = obter_credenciais()
        cliente = gspread.authorize(credenciais)
        id_planilha = "1OPl-0WAFUTQt6Nd1VZBTDgXr5SzjvLHNirpwgjf9kXc"
        planilha = cliente.open_by_key(id_planilha)
        aba = planilha.get_worksheet(0)
        return pd.DataFrame(aba.get_all_records())
    except Exception as e:
        st.error(f"Erro na conexão moderna: {e}")
        return pd.DataFrame()

def consultar_status_matricula_api(matricula_buscada):
    try:
        credenciais = obter_credenciais()
        cliente = gspread.authorize(credenciais)
        planilha = cliente.open_by_key("1OPl-0WAFUTQt6Nd1VZBTDgXr5SzjvLHNirpwgjf9kXc")
        aba_pesquisa = planilha.worksheet("Ferramenta de Pesquisa")
        aba_pesquisa.update("B2", [[str(matricula_buscada)]])
        time.sleep(1)
        res = aba_pesquisa.row_values(6)[1:7]
        return ("Enviou", res) if any(str(v).strip() != "" for v in res) else ("Não enviou", [])
    except Exception as e: 
        return (f"Erro na consulta: {e}", [])

# Controle de estado
if "termo_aceito" not in st.session_state:
    st.session_state.termo_aceito = None
if "matricula_verificada" not in st.session_state:
    st.session_state.matricula_verificada = False
if "nome_servidor" not in st.session_state:
    st.session_state.nome_servidor = None
if "pdf_proc" not in st.session_state:
    st.session_state.pdf_proc = None
if "pdf_termo" not in st.session_state:
    st.session_state.pdf_termo = None
if "link_assinatura" not in st.session_state:
    st.session_state.link_assinatura = None
if "status_assinafy" not in st.session_state:
    st.session_state.status_assinafy = None

# --- QUADRO DE CONSENTIMENTO INICIAL ---
if st.session_state.termo_aceito is None:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col_centro_1, col_centro_2, col_centro_3 = st.columns([1, 2, 1])
    with col_centro_2:
        with st.container(border=True):
            st.markdown("### 📋 Termo de Consentimento e Privacidade")
            st.markdown("Esse site tem o objetivo de coletar informações para o ajuizamento da ação de correção monetária de exercícios anteriores.")
            st.markdown("---")
            st.markdown("🔒 **Compartilhamento de dados com o Sinprfpi.**")
            st.markdown("")
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                if st.button("✅ Aceito", key="btn_aceito"):
                    st.session_state.termo_aceito = True
                    st.rerun()
            with col_b2:
                if st.button("❌ Não aceito", key="btn_nao_aceito"):
                    st.session_state.termo_aceito = False
                    st.rerun()
    st.stop()

elif st.session_state.termo_aceito is False:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col_b_1, col_b_2, col_b_3 = st.columns([1, 2, 1])
    with col_b_2:
        st.error("🚫 **Acesso Bloqueado.** \n\nVocê recusou os termos de compartilhamento de dados.")
    st.stop()

# --- VERIFICAÇÃO DE MATRÍCULA LOGO APÓS ACEITAR OS TERMOS ---
if not st.session_state.matricula_verificada:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col_v1, col_v2, col_v3 = st.columns([1, 2, 1])
    with col_v2:
        with st.container(border=True):
            st.markdown("### 🔍 Verificação de Cadastro")
            st.markdown("Informe sua **Matrícula (SIAPE)** para verificar se sua documentação já foi enviada.")
            st.markdown("---")
            mat_check = st.text_input("Digite sua Matrícula:", key="input_mat_verificacao")
            if st.button("Verificar e Continuar", key="btn_verificar_matricula"):
                if not mat_check.strip():
                    st.warning("⚠️ Por favor, informe a matrícula.")
                else:
                    with st.spinner("Consultando status no sistema..."):
                        status_envio, dados_envio = consultar_status_matricula_api(mat_check.strip())
                        if status_envio == "Enviou":
                            st.error("🚫 **Acesso Bloqueado: Documentação Já Enviada!**\n\nIdentificamos que você já enviou os documentos necessários.")
                            st.stop()
                        elif status_envio == "Não enviou":
                            st.session_state.matricula_verificada = True
                            st.session_state["input_mat"] = mat_check.strip()
                            st.success("✅ Verificação concluída! Liberando formulário...")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error(f"❌ Ocorreu um erro na consulta: {status_envio}")
    st.stop()

# --- DEMAIS FUNÇÕES DO SISTEMA ---
def salvar_no_excel(dados):
    try:
        credenciais = obter_credenciais()
        cliente = gspread.authorize(credenciais)
        id_planilha = "1OPl-0WAFUTQt6Nd1VZBTDgXr5SzjvLHNirpwgjf9kXc"
        planilha = cliente.open_by_key(id_planilha)
        aba = planilha.get_worksheet(0)
        
        nova_linha = [
            dados.get('Local Preenchimento Município', ''),
            dados.get('Local Preenchimento Estado', ''),
            dados.get('Matrícula', ''),
            dados.get('Cargo', ''),
            dados.get('Órgão', ''),
            dados.get('Data de Nascimento', ''),
            dados.get('Data de Ingresso', ''),
            dados.get('Nome', ''),
            dados.get('CPF', ''),
            dados.get('E-mail', ''),
            dados.get('RG', ''),
            dados.get('Telefone', ''),
            dados.get('Estado Civil', ''),
            dados.get('CEP', ''),
            dados.get('Endereço', ''),
            dados.get('Município', ''),
            dados.get('Estado', '')
        ]
        
        aba.append_row(nova_linha)
        return True
    except Exception as e:
        print(f"Erro ao salvar na planilha: {e}")
        return False

def enviar_para_google_drive(nome_servidor, lista_arquivos):
    payload = {"nomeServidor": nome_servidor, "arquivos": lista_arquivos}
    try:
        headers = {"Content-Type": "application/json"}
        response = requests.post(GOOGLE_APPS_SCRIPT_URL, data=json.dumps(payload), headers=headers)
        resultado = response.json()
        return resultado.get("status") == "sucesso"
    except Exception as e:
        print(f"Erro ao conectar: {e}")
        return False

def enviar_para_assinafy(nome, email, pdf_bytes, nome_arquivo):
    if not pdf_bytes:
        return False, "PDF não foi gerado."
    return False, "ASSINATURA_MANUAL_NECESSARIA"

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

def formatar_nasc_callback():
    val = st.session_state.get("input_nasc_raw", "")
    digitos = "".join(filter(str.isdigit, str(val)))[:8]
    formatado = ""
    if len(digitos) > 4:
        formatado = f"{digitos[:2]}/{digitos[2:4]}/{digitos[4:]}"
    elif len(digitos) > 2:
        formatado = f"{digitos[:2]}/{digitos[2:]}"
    else:
        formatado = digitos
    st.session_state.input_nasc_raw = formatado

def limpar_valor(val):
    if val is None or pd.isna(val) or str(val).strip().lower() == "nan":
        return ""
    return str(val).strip()

def obter_data_por_extenso():
    meses = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
        5: "maio", 6: "junho", 7: "julho", 8: "agosto",
        9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }
    agora = datetime.now()
    dia = agora.strftime("%d")
    mes = meses[agora.month]
    ano = agora.strftime("%Y")
    return f"{dia} de {mes} de {ano}"

def preencher_documentos_oficiais(dados):
    caminho_procuracao = "template_procuracao.pdf"
    caminho_termo = "template_termo.pdf"
    pdf_procuracao_bytes = None
    pdf_termo_bytes = None
    mun = limpar_valor(dados.get('Local Preenchimento Município', ''))
    est = limpar_valor(dados.get('Local Preenchimento Estado', ''))
    data_extenso = obter_data_por_extenso()
    local_data_str = f"{mun}/{est}, {data_extenso}" if mun and est else f", {data_extenso}"

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
        pag_proc.insert_text((90, 715), local_data_str, fontsize=9, color=(0,0,0))
        pdf_procuracao_bytes = doc_proc.tobytes()
        doc_proc.close()

    if os.path.exists(caminho_termo):
        doc_termo = fitz.open(caminho_termo)
        pag_termo = doc_termo[0]
        pag_termo.insert_text((125, 145), str(dados.get('Nome', '')), fontsize=9, color=(0,0,0))
        pag_termo.insert_text((82, 170), str(dados.get('CPF', '')), fontsize=9, color=(0,0,0))
        pag_termo.insert_text((265, 170), str(dados.get('Matrícula', '')), fontsize=9, color=(0,0,0))
        pag_termo.insert_text((377, 169), str(dados.get('Cargo', '')), fontsize=9, color=(0,0,0))
        pag_termo.insert_text((90, 310), local_data_str, fontsize=9, color=(0,0,0))
        pdf_termo_bytes = doc_termo.tobytes()
        doc_termo.close()

    return pdf_procuracao_bytes, pdf_termo_bytes

if "aba_selecionada" not in st.session_state:
    st.session_state.aba_selecionada = "➕ Novo Cadastro"

col_cab1, col_cab2 = st.columns([3, 1])
with col_cab1:
    st.title("⚖️ Sistema de Cadastro e Gestão de Documentos")
    st.markdown("##### **Ação de Correção Monetária de Exercícios Anteriores**")
with col_cab2:
    df_geral = carregar_servidores_cadastrados()
    total_cadastrados = len(df_geral) if not df_geral.empty else 0
    st.metric(label="📊 Total Cadastrados", value=total_cadastrados)

st.markdown("---")

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
    st.markdown('<p class="suporte-discreto">Suporte - Guilherme (86988523711)</p>', unsafe_allow_html=True)

abas_disponiveis = ["➕ Novo Cadastro", "📂 Servidores Já Cadastrados", "📖 Tutorial"]
st.session_state.aba_selecionada = st.radio(
    "Navegação:",
    abas_disponiveis,
    index=abas_disponiveis.index(st.session_state.aba_selecionada) if st.session_state.aba_selecionada in abas_disponiveis else 0,
    horizontal=True,
    label_visibility="collapsed"
)
st.markdown("---")

# --- ABA: SERVIDORES JÁ CADASTRADOS (ÁREA DO ADMINISTRADOR) ---
if st.session_state.aba_selecionada == "📂 Servidores Já Cadastrados":
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

                        if st.session_state.get("pdf_proc") and st.session_state.get("pdf_termo"):
                            nome_serv_limpo = str(dados_linha.get("Nome", "Servidor")).replace(" ", "_")
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                                zip_file.writestr(f"Procuracao_{nome_serv_limpo}.pdf", st.session_state.pdf_proc)
                                zip_file.writestr(f"Termo_{nome_serv_limpo}.pdf", st.session_state.pdf_termo)
                            zip_buffer.seek(0)
                            st.markdown("<br>", unsafe_allow_html=True)
                            st.download_button(
                                label="📥 Baixar Procuração e Termo Juntos (ZIP)",
                                data=zip_buffer,
                                file_name=f"Documentos_{nome_serv_limpo}.zip",
                                mime="application/zip",
                                key=f"btn_zip_{nome_serv_limpo}"
                            )

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
                        st.markdown(f"""<a href="{url_geral}" target="_blank"><button style="width:100%; border-radius:6px; font-weight:bold; height:2.5em; background-color:#198754; color:white; border:none; cursor:pointer; margin-bottom: 8px;">📝 Termo de Consentimento</button></a>""", unsafe_allow_html=True)

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
                        st.markdown(f"""<a href="{url_preenchida}" target="_blank"><button style="width:100%; border-radius:6px; font-weight:bold; height:2.5em; background-color:#3b82f6; color:white; border:none; cursor:pointer;">📝 Forms Correção Monetária</button></a>""", unsafe_allow_html=True)
    else:
        st.warning("🔒 **Conteúdo Restrito.** Abra a **Área do Administrador** na barra lateral e insira a chave de acesso correta.")

# --- ABA: NOVO CADASTRO ---
elif st.session_state.aba_selecionada == "➕ Novo Cadastro":
    col_sub_1, col_sub_2 = st.columns([3, 1])
    with col_sub_1:
        st.subheader("📝 Preenchimento de Dados Cadastrais")
    with col_sub_2:
        if st.button("💡 Ver Tutorial de Ajuda", key="btn_ir_tutorial"):
            st.session_state.aba_selecionada = "📖 Tutorial"
            st.rerun()

    st.markdown(f'''<a href="https://wa.me/5586988523711" target="_blank"><button class="btn-whatsapp">📱 PRECISA DE AJUDA? Fale conosco no WhatsApp</button></a>''', unsafe_allow_html=True)

    st.markdown("""
        <div class="box-instrucoes">
            <b>📌 Informações Importantes:</b><br>
            • <b>Onde vão os dados:</b> Seus dados são salvos com segurança na planilha e enviados para a pasta oficial do sistema no Google Drive.<br>
            • <b>Como assinar:</b> Após gerar os documentos, você pode assiná-los via painel digital do Assinafy ou baixar/enviar os arquivos assinados.
        </div>
    """, unsafe_allow_html=True)

    if st.session_state.get("pdf_proc") is None:
        with st.container(border=True):
            st.markdown('<div class="box-passo"><b>📍 Passo 1: Preencha os campos abaixo corretamente:</b></div>', unsafe_allow_html=True)

            val_1 = st.session_state.get("input_local_mun", "")
            val_2 = st.session_state.get("input_local_uf", "")
            val_3 = st.session_state.get("input_mat", "")
            val_4 = st.session_state.get("input_nome", "")
            val_5 = st.session_state.get("input_cargo", "")
            val_6 = st.session_state.get("input_cpf", "")
            val_7 = st.session_state.get("input_orgao", "")
            val_8 = st.session_state.get("input_rg", "")
            val_9 = st.session_state.get("input_nasc_raw", "")
            val_10 = st.session_state.get("input_ing_raw", "")
            val_11 = st.session_state.get("input_email", "")
            val_12 = st.session_state.get("input_tel", "")
            val_13 = st.session_state.get("input_cep", "")
            val_14 = st.session_state.get("input_mun", "")
            val_15 = st.session_state.get("input_ec", "")
            val_16 = st.session_state.get("input_end", "")
            val_17 = st.session_state.get("input_uf", "")

            def s(n):
                p = 1
                for x in [val_1, val_2, val_3, val_4, val_5, val_6, val_7, val_8, val_9, val_10, val_11, val_12, val_13, val_14, val_15, val_16, val_17]:
                    if not str(x).strip():
                        break
                    p += 1
                return "➡️ " if n == p else ""

            st.markdown(f"**📍 Local de Preenchimento**")
            col_loc1, col_loc2 = st.columns(2)
            with col_loc1:
                local_municipio = st.text_input(f"{s(1)}1. Município que está", key="input_local_mun")
            with col_loc2:
                local_estado = st.text_input(f"{s(2)}2. Estado", key="input_local_uf")

            st.markdown("---")
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                st.markdown("**💼 Dados Profissionais**")
                matricula = st.text_input(f"{s(3)}3. Matrícula (SIAPE)", key="input_mat")
                cargo = st.text_input(f"{s(5)}5. Cargo", key="input_cargo")
                orgao = st.text_input(f"{s(7)}7. Órgão", key="input_orgao")
                nascimento = st.text_input(f"{s(9)}9. Data de Nascimento", placeholder="DD/MM/AAAA", key="input_nasc_raw", max_chars=10, on_change=formatar_nasc_callback)
                ingresso = st.text_input(f"{s(10)}10. Data de Ingresso", placeholder="DD/MM/AAAA", key="input_ing_raw", max_chars=10, on_change=formatar_data_callback)

            with col_p2:
                st.markdown("**👤 Dados Pessoais**")
                nome = st.text_input(f"{s(4)}4. Nome completo", key="input_nome")
                cpf = st.text_input(f"{s(6)}6. CPF", key="input_cpf")
                rg = st.text_input(f"{s(8)}8. RG - Órgão de Expedição", key="input_rg")
                email = st.text_input(f"{s(11)}11. E-mail", key="input_email")

            col_p3, col_p4, col_p5 = st.columns(3)
            with col_p3:
                telefone = st.text_input(f"{s(12)}12. Telefone", key="input_tel")
                estado_civil = st.text_input(f"{s(15)}15. Estado Civil", key="input_ec")
            with col_p4:
                cep = st.text_input(f"{s(13)}13. CEP", key="input_cep")
                endereco = st.text_input(f"{s(16)}16. Endereço", key="input_end")
            with col_p5:
                municipio = st.text_input(f"{s(14)}14. Município", key="input_mun")
                estado = st.text_input(f"{s(17)}17. Estado (UF)", key="input_uf")

        todos_preenchidos = all(str(x).strip() for x in [val_1, val_2, val_3, val_4, val_5, val_6, val_7, val_8, val_9, val_10, val_11, val_12, val_13, val_14, val_15, val_16, val_17])

        if todos_preenchidos:
            st.markdown('<p class="seta-guiada">➡️ Tudo preenchido! Clique no botão abaixo para salvar:</p>', unsafe_allow_html=True)

        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            btn_salvar = st.button("💾 Salvar na Planilha e Gerar Documentos", key="btn_salvar_novo")

        if btn_salvar:
            if not nome.strip():
                st.error("⚠️ Por favor, preencha o campo 'Nome completo'.")
            elif not email.strip():
                st.error("⚠️ Por favor, preencha o campo 'E-mail'.")
            else:
                dados_usuario = {
                    "Local Preenchimento Município": local_municipio,
                    "Local Preenchimento Estado": local_estado,
                    "Matrícula": matricula, "Cargo": cargo, "Órgão": orgao, 
                    "Data de Nascimento": nascimento, "Data de Ingresso": ingresso,
                    "Nome": nome, "CPF": cpf, "E-mail": email, "RG": rg, "Telefone": telefone,
                    "Estado Civil": estado_civil, "CEP": cep, "Endereço": endereco, "Município": municipio, "Estado": estado
                }
                salvar_no_excel(dados_usuario)
                st.session_state.dados_usuario = dados_usuario
                st.session_state.nome_servidor = nome.strip()
                st.session_state.pdf_proc, st.session_state.pdf_termo = preencher_documentos_oficiais(dados_usuario)
                st.success(f"✨ Dados salvos com sucesso para **{st.session_state.nome_servidor}**!")
                st.rerun()

    if st.session_state.get("pdf_proc") is not None:
        st.markdown("---")
        with st.container(border=True):
            st.markdown('<div class="box-passo" style="border-left-color: #198754;"><b>✅ Passo 2: Baixe os documentos gerados</b></div>', unsafe_allow_html=True)
            st.subheader(f"⚙️ Arquivos prontos para: {st.session_state.get('nome_servidor', '')}")
            sufixo_chave = str(st.session_state.get('nome_servidor', 'usuario')).replace(" ", "_")
            col_dl1, col_dl2 = st.columns(2)
            if st.session_state.get("pdf_proc"):
                with col_dl1:
                    st.download_button(label="📄 Baixar Procuração", data=st.session_state.pdf_proc, file_name="Procuracao_Preenchida.pdf", mime="application/pdf", key=f"dl_proc_{sufixo_chave}")
            if st.session_state.get("pdf_termo"):
                with col_dl2:
                    st.download_button(label="📄 Baixar Termo", data=st.session_state.pdf_termo, file_name="Termo_Preenchido.pdf", mime="application/pdf", key=f"dl_termo_{sufixo_chave}")

            st.markdown("---")
            st.markdown('<div class="box-passo"><b>✍️ Passo 3: Assinatura Digital</b></div>', unsafe_allow_html=True)
            st.markdown('<a href="https://app.assinafy.com.br/dashboard" target="_blank" class="btn-assinar">✍️ ABRIR PAINEL DO ASSINAFY PARA ASSINAR</a>', unsafe_allow_html=True)
            st.info("ℹ️ **Nota:** Clique acima para abrir o painel do Assinafy, faça o upload dos arquivos baixados e realize a assinatura.")

        st.markdown("---")
        with st.container(border=True):
            st.markdown('<div class="box-passo"><b>📤 Passo 4: Envio de Documentos para o Google Drive</b></div>', unsafe_allow_html=True)
            col_up1, col_up2 = st.columns(2)
            with col_up1:
                doc_identidade = st.file_uploader("🪪 1. Documento de Identidade com foto", type=["pdf", "jpg", "jpeg", "png"], key="up_identidade")
                upload_proc_assinada = st.file_uploader("✍️ 3. Procuração assinada", type=["pdf"], key="upload_proc")
            with col_up2:
                comprovante_residencia = st.file_uploader("🏠 2. Comprovante de residência atualizado", type=["pdf", "jpg", "jpeg", "png"], key="up_residencia")
                upload_termo_assinado = st.file_uploader("✍️ 4. Termo assinado", type=["pdf"], key="upload_termo")

            st.markdown("<br>", unsafe_allow_html=True)
            col_env1, col_env2, col_env3 = st.columns([1, 2, 1])
            with col_env2:
                btn_enviar_drive = st.button("🚀 Enviar arquivos para o Google Drive", key="btn_enviar_drive")

            if btn_enviar_drive:
                nome_pasta = st.session_state.get("nome_servidor", "Servidor")
                lista_arquivos_payload = []
                if st.session_state.get("pdf_proc"):
                    lista_arquivos_payload.append({"nome": "Procuracao_Preenchida.pdf", "conteudo": base64.b64encode(st.session_state.pdf_proc).decode('utf-8'), "mimeType": "application/pdf"})
                if st.session_state.get("pdf_termo"):
                    lista_arquivos_payload.append({"nome": "Termo_Preenchido.pdf", "conteudo": base64.b64encode(st.session_state.pdf_termo).decode('utf-8'), "mimeType": "application/pdf"})
                for arquivo_up in [doc_identidade, comprovante_residencia, upload_proc_assinada, upload_termo_assinado]:
                    if arquivo_up is not None:
                        lista_arquivos_payload.append({"nome": arquivo_up.name, "conteudo": base64.b64encode(arquivo_up.getbuffer()).decode('utf-8'), "mimeType": arquivo_up.type})
                if lista_arquivos_payload:
                    with st.spinner("Enviando arquivos para o Google Drive..."):
                        if enviar_para_google_drive(nome_pasta, lista_arquivos_payload):
                            st.success(f"🎉 Sucesso! A pasta de **{nome_pasta}** foi atualizada.")
                            st.session_state.pdf_proc = None
                            st.session_state.pdf_termo = None
                            st.session_state.link_assinatura = None
                            st.session_state.status_assinafy = None
                            st.session_state.nome_servidor = None
                            st.rerun()
                        else:
                            st.error("❌ Erro ao enviar para o Google Drive.")
                else:
                    st.warning("⚠️ Nenhum arquivo anexado.")

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 Fazer Novo Cadastro / Limpar Tela"):
            st.session_state.pdf_proc = None
            st.session_state.pdf_termo = None
            st.rerun()

# --- ABA: TUTORIAL ---
elif st.session_state.aba_selecionada == "📖 Tutorial":
    col_tut_1, col_tut_2 = st.columns([3, 1])
    with col_tut_1:
        st.subheader("📖 Tutorial de Utilização do Sistema")
    with col_tut_2:
        if st.button("⬅️ Voltar ao Cadastro", key="btn_voltar_cadastro"):
            st.session_state.aba_selecionada = "➕ Novo Cadastro"
            st.rerun()

    st.info("Selecione abaixo o dispositivo que você está utilizando para visualizar o tutorial correspondente:")
    tipo_tutorial = st.radio("Escolha a plataforma:", ["💻 Tutorial Computador", "📱 Tutorial Celular"], horizontal=True, label_visibility="collapsed")
    st.markdown("---")

    if tipo_tutorial == "💻 Tutorial Computador":
        st.markdown("### 🖥️ Passo a Passo para Computador")
        caminho_video_pc = os.path.join("imagens", "Tutorial Computador", "video.mp4")
        if os.path.exists(caminho_video_pc):
            st.video(caminho_video_pc)
        else:
            pasta_pc = os.path.join("imagens", "Tutorial Computador")
            if os.path.exists(pasta_pc):
                videos_encontrados = [f for f in os.listdir(pasta_pc) if f.endswith(('.mp4', '.mov', '.avi'))]
                if videos_encontrados:
                    st.video(os.path.join(pasta_pc, videos_encontrados[0]))
                else:
                    st.warning("⚠️ Nenhum arquivo de vídeo encontrado.")
            else:
                st.warning("⚠️ A pasta não foi encontrada.")
        for i in range(1, 4):
            with st.expander(f"Passo {i} (Computador) — Clique para visualizar a orientação"):
                caminho_img = os.path.join("imagens", f"pc_{i}.png")
                if os.path.exists(caminho_img): st.image(caminho_img, width=700)
                caminho_anim_pc = os.path.join("imagens", f"anim_pc_{i}.gif")
                if os.path.exists(caminho_anim_pc): st.image(caminho_anim_pc, caption=f"Movimento Indicativo - Passo {i}", width=400)
    else:
        st.markdown("### 📱 Passo a Passo para Celular")
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
                    st.warning("⚠️ Nenhum arquivo de vídeo encontrado.")
            else:
                st.warning("⚠️ A pasta não foi encontrada.")
        for i in range(1, 4):
            with st.expander(f"Passo {i} (Celular) — Clique para visualizar a orientação"):
                caminho_img = os.path.join("imagens", f"cel_{i}.png")
                if os.path.exists(caminho_img): st.image(caminho_img, width=700)
                caminho_anim_cel = os.path.join("imagens", f"anim_cel_{i}.gif")
                if os.path.exists(caminho_anim_cel): st.image(caminho_anim_cel, caption=f"Movimento Indicativo - Passo {i}", width=300)