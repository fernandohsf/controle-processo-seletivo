def ajustar_planilha(app, arquivo_id):
    planilha = app.cliente_gspread.open_by_key(arquivo_id)
    sheet = planilha.sheet1

    cell_updates = []

    def inserir_coluna_e_atualizar(coluna, nome, formula):
        sheet.insert_cols([[None]], coluna)
        sheet.update_cell(1, coluna, nome)
    
        for row in range(2, sheet.row_count + 1):
            cell_updates.append({
                'range': f'{chr(64 + coluna)}{row}',
                'values': [[formula.format(row=row)]]
            })

    # COLUNAS A, T e U
    inserir_coluna_e_atualizar(1, 'Número da Inscrição', """=SE('Respostas ao formulário 1'!$B:$B="", "", "PSSI"&AD{row})""")
    inserir_coluna_e_atualizar(20, 'CPF (Conferência)', r"""=TEXTO('Respostas ao formulário 1'!H:H;"000\.000\.000\-00")""")
    inserir_coluna_e_atualizar(21, 'Quantidade de Inscrições neste PSSI', """=CONT.SE(T:T;T{row})""")
    inserir_coluna_e_atualizar(22, 'Notas', "")
    inserir_coluna_e_atualizar(23, 'Analista RH', "")
    inserir_coluna_e_atualizar(24, 'Status da inscrição', "")
    inserir_coluna_e_atualizar(25, 'Motivo de Inaptidão/Eliminação', "")
    inserir_coluna_e_atualizar(26, 'Controle de e-mail', "")
    inserir_coluna_e_atualizar(27, 'Classificação do Candidato (Coordenador)', "")
    inserir_coluna_e_atualizar(28, 'Ranking', """=SEERRO(ORDEM(B$1:B$903;B$1:B$903;VERDADEIRO);"")""")
    inserir_coluna_e_atualizar(29, 'ID Inscrição', """=SE(B{row}="";"";(ESQUERDA('Respostas ao formulário 1'!$E$2;3)&(DIREITA('Respostas ao formulário 1'!$E${row};4)&AB{row})))""")
    inserir_coluna_e_atualizar(30, 'Formatação ID', """=TEXTO(AC{row};"000000000")""")
    
    # MENU SUSPENSO
    def menu_suspenso(lista, letra_coluna):
        menu = {
            'condition': {
                'type': 'ONE_OF_LIST',
                'values': [{'userEnteredValue': item} for item in lista]
            },
            'inputMessage': 'Selecione uma opção',
            'showInputMessage': True,
            'strict': True,
            'ignoreBlanks': True
        }
        sheet.add_validation(f'{letra_coluna}2:{letra_coluna}{sheet.row_count}', menu)

    def aplicar_formatação_condicional(lista_itens, coluna, cores):
        regras = []
        
        for index, item in enumerate(lista_itens):
            cor = cores[index]
            
            regra = {
                "addConditionalFormatRule": {
                    "rule": {
                        "condition": {
                            "type": "TEXT_EQ",
                            "values": [{"userEnteredValue": item}]
                        },
                        "format": {
                            "backgroundColor": {
                                "red": cor[0] / 255,
                                "green": cor[1] / 255,
                                "blue": cor[2] / 255
                            }
                        },
                        "ranges": [{
                            "sheetId": sheet.id,
                            "startRowIndex": 1,  # Começa na segunda linha
                            "startColumnIndex": coluna - 1,  # Coluna (ajuste de 0 index)
                            "endColumnIndex": coluna
                        }]
                    },
                    "index": index
                }
            }
            regras.append(regra)
        
        sheet.batch_update({"requests": regras})

    analistas = ['Elizangela Ribeiro da Silva', 'Fernanda Neris Barroso', 'Wanessa Aparecida Nunes de Matos']
    cores = [(250, 250, 250)]
    menu_suspenso(analistas, 'W')  # Coluna 23 -> 'W'
    aplicar_formatação_condicional(analistas, 23, cores)

    status_inscricao = ['Em análise', 'Apto', 'Inapto', 'Substituído por inscrição mais recente', 'Classificado para Vaga', 'Cadastro de Reserva', 'Desistente']
    cores = [(191, 225, 246), (212, 237, 188), (255, 207, 201), (117, 56, 0), (10, 83, 168), (230, 207, 242), (177, 2, 2)]
    menu_suspenso(status_inscricao, 'X')  # Coluna 24 -> 'X'
    aplicar_formatação_condicional(status_inscricao, 24, cores)

    controle_email = ['Enviar', 'Enviado', 'Não se aplica']
    cores = [(255, 207, 201), (191, 225, 246), (250, 250, 250)]
    menu_suspenso(controle_email, 'Z')  # Coluna 26 -> 'Z'
    aplicar_formatação_condicional(controle_email, 26, cores)

    sheet.format('V1', {"backgroundColor": {"red": 52/255, "green": 168/255, "blue": 83/255}})  # verde
    sheet.format('W1', {"backgroundColor": {"red": 52/255, "green": 168/255, "blue": 83/255}})  # verde
    sheet.format('X1', {"backgroundColor": {"red": 52/255, "green": 168/255, "blue": 83/255}})  # verde
    sheet.format('Y1', {"backgroundColor": {"red": 52/255, "green": 168/255, "blue": 83/255}})  # verde
    sheet.format('Z1', {"backgroundColor": {"red": 52/255, "green": 168/255, "blue": 83/255}})  # verde
    sheet.format('AA1', {"backgroundColor": {"red": 52/255, "green": 168/255, "blue": 83/255}})  # verde

    sheet.batch_update(cell_updates)