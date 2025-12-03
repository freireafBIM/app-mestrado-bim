def processar_ifc(caminho_arquivo, id_projeto_input):
    ifc_file = ifcopenshell.open(caminho_arquivo)
    pilares = ifc_file.by_type('IfcColumn')
    dados = []
    
    progresso = st.progress(0)
    total = len(pilares)
    
    for i, pilar in enumerate(pilares):
        progresso.progress((i + 1) / total)
        
        guid = pilar.GlobalId
        nome = pilar.Name if pilar.Name else "S/N"
        
        # Pavimento
        pavimento = "Térreo"
        if pilar.ContainedInStructure:
            pavimento = pilar.ContainedInStructure[0].RelatingStructure.Name
        
        sufixo_pav = limpar_string(pavimento)
        sufixo_nome = limpar_string(nome)

        # --- NOVA CHAVE PRIMÁRIA "A PROVA DE FALHAS" ---
        # Incluímos o NOME do pilar para diferenciar visualmente P1 de P3
        id_unico_pilar = f"{sufixo_nome}-{guid}-{sufixo_pav}-{id_projeto_input}"

        # Geometria
        secao = extrair_secao_robusta(pilar) # Certifique-se que esta função existe no seu código

        armadura = extrair_texto_armadura(pilar)

        dados.append({
            'ID_Unico': id_unico_pilar,   
            'Projeto_Ref': id_projeto_input, 
            'Nome': nome, 
            'Secao': secao,
            'Armadura': armadura, 
            'Pavimento': pavimento,
            'Status': 'A CONFERIR', 
            'Data_Conferencia': '', 
            'Responsavel': ''
        })
    
    dados.sort(key=lambda x: (x['Pavimento'], x['Nome']))
    return dados
