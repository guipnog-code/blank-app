def preencher_documentos_oficiais(dados):
    caminho_procuracao = "template_procuracao.pdf"
    caminho_termo = "template_termo.pdf"
    
    pdf_procuracao_bytes = None
    pdf_termo_bytes = None

    # Função auxiliar para preencher campos de formulário (AcroFields) mantendo editabilidade
    def preencher_form_pdf(caminho_template, campos_valores):
        if not os.path.exists(caminho_template):
            return None
        
        doc = fitz.open(caminho_template)
        for pagina in doc:
            # Procura por campos de formulário interativos na página
            for widget in pagina.widgets():
                nome_campo = widget.field_name
                if nome_campo in campos_valores:
                    widget.field_value = str(campos_valores[nome_campo])
                    widget.update()
                    
        bytes_pdf = doc.tobytes()
        doc.close()
        return bytes_pdf

    # Mapeamento para a Procuração (substitua 'Nome_do_Campo_No_PDF' pelos nomes reais dos campos do seu template)
    mapeamento_procuracao = {
        'Nome': dados['Nome'],
        'CPF': dados['CPF'],
        'RG': dados['RG'],
        'Cargo': dados['Cargo'],
        'Órgão': dados['Órgão'],
        'Data de Ingresso': dados['Data de Ingresso'],
        'Estado Civil': dados['Estado Civil'],
        'Telefone': dados['Telefone'],
        'E-mail': dados['E-mail'],
        'Endereço': dados['Endereço'],
        'Município': dados['Município'],
        'Estado': dados['Estado'],
        'CEP': dados['CEP']
    }

    # Mapeamento para o Termo
    mapeamento_termo = {
        'Nome': dados['Nome'],
        'CPF': dados['CPF'],
        'Matrícula': dados['Matrícula'],
        'Cargo': dados['Cargo']
    }

    pdf_procuracao_bytes = preencher_form_pdf(caminho_procuracao, mapeamento_procuracao)
    pdf_termo_bytes = preencher_form_pdf(caminho_termo, mapeamento_termo)

    return pdf_procuracao_bytes, pdf_termo_bytes