# --- Adicione isto dentro da sua função btn_salvar, após o envio para o Assinafy ---

    if btn_salvar:
        # ... (seu código de salvamento na planilha) ...
        
        # Gera os PDFs
        st.session_state.pdf_proc, st.session_state.pdf_termo = preencher_documentos_oficiais(dados_usuario)
        
        # Envia para o Assinafy e captura o link
        sucesso, resposta = enviar_para_assinafy(nome, email, st.session_state.pdf_proc, "Procuracao.pdf")
        
        if sucesso:
            # O Assinafy geralmente retorna um JSON com uma url de assinatura
            st.session_state.link_assinatura = resposta.get("url") # Verifique na documentação do Assinafy o nome exato desse campo
            st.success("Documentos enviados para assinatura!")
        else:
            st.error("Erro ao conectar com Assinafy.")
            
        st.rerun()

# --- No bloco de Gestão de Arquivos (onde você baixa os arquivos), adicione o botão ---

    if tem_documentos:
        # ... (seus botões de download) ...
        
        # Botão de Assinatura Digital
        if "link_assinatura" in st.session_state and st.session_state.link_assinatura:
            st.markdown(f"""
                <a href="{st.session_state.link_assinatura}" target="_blank">
                    <button style="width:100%; background-color:#ff4b4b; color:white; border:none; padding:10px; border-radius:5px; font-weight:bold;">
                        ✍️ CLIQUE AQUI PARA ASSINAR DIGITALMENTE
                    </button>
                </a>
            """, unsafe_allow_html=True)