import os
from utils import arquivos_selecionados, verificar_selecao

def ajustar_planilha(app):
    if not verificar_selecao(app):
        return
    
    arquivos = arquivos_selecionados(app)
    
    for arquivo_id in arquivos:
        planilha = app.cliente_gspread.open_by_key(arquivo_id)
        sheet = planilha.sheet1

        def inserir_coluna_e_atualizar(coluna, nome, formula_template):
            cell_updates = []
            sheet.insert_cols([[None]], coluna)
            sheet.update_cell(1, coluna, nome)

            coluna_letra = numero_para_letra_coluna(coluna)

            for row in range(2, sheet.row_count + 1):
                formula = formula_template.format(row=row)
                cell_updates.append({
                    'range': f'{coluna_letra}{row}',
                    'values': [[formula]]
                })

            sheet.batch_update(cell_updates)

        # COLUNAS A, T e U
        inserir_coluna_e_atualizar(1, 'Número da Inscrição', "=SE('Respostas ao formulário 1'!$B:$B=\"\", \"\", \"PSSI\"&AD{row})")
        inserir_coluna_e_atualizar(20, 'CPF (Conferência)', r"""=TEXTO('Respostas ao formulário 1'!H:H;"000\.000\.000\-00")""")
        inserir_coluna_e_atualizar(21, 'Quantidade de Inscrições neste PSSI', "=CONT.SE(T:T;T{row})")
        inserir_coluna_e_atualizar(22, 'Notas', "")
        inserir_coluna_e_atualizar(23, 'Analista RH', "")
        inserir_coluna_e_atualizar(24, 'Status da inscrição', "")
        inserir_coluna_e_atualizar(25, 'Motivo de Inaptidão/Eliminação', "")
        inserir_coluna_e_atualizar(26, 'Controle de e-mail', "")
        inserir_coluna_e_atualizar(27, 'Classificação do Candidato (Coordenador)', "")
        inserir_coluna_e_atualizar(28, 'Ranking', """=SEERRO(ORDEM(B$1:B$903;B$1:B$903;VERDADEIRO);"")""")
        inserir_coluna_e_atualizar(29, 'ID Inscrição', "=SE(B{row}=\"\";\"\";(ESQUERDA('Respostas ao formulário 1'!$E$2;3)&(DIREITA('Respostas ao formulário 1'!$E${row};4)&AB{row})))")
        inserir_coluna_e_atualizar(30, 'Formatação ID', "=TEXTO(AC{row};\"000000000\")")
        
        def menu_suspenso(lista, letra_coluna):
            rule = {
                "requests": [
                    {
                        "setDataValidation": {
                            "range": {
                                "sheetId": sheet.id,
                                "startRowIndex": 1,
                                "endRowIndex": sheet.row_count,
                                "startColumnIndex": ord(letra_coluna) - 65,
                                "endColumnIndex": ord(letra_coluna) - 64
                            },
                            "rule": {
                                "condition": {
                                    "type": "ONE_OF_LIST",
                                    "values": [{"userEnteredValue": item} for item in lista]
                                },
                                "inputMessage": "Selecione uma opção",
                                "showCustomUi": True,
                                "strict": True
                            }
                        }
                    }
                ]
            }
            sheet.spreadsheet.batch_update(rule)

        def aplicar_formatação_condicional(lista_itens, coluna, cores):
            requests = []
            
            for item, cor in zip(lista_itens, cores):
                requests.append({
                    "addConditionalFormatRule": {
                        "rule": {
                            "ranges": [{
                                "sheetId": sheet.id,
                                "startRowIndex": 1,
                                "startColumnIndex": coluna - 1, # Ajuste para índice 0
                                "endColumnIndex": coluna
                            }],
                            "booleanRule": {
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
                                }
                            }
                        },
                        "index": 0
                    }
                })

            sheet.spreadsheet.batch_update({"requests": requests})

        analistas = ['Elizangela Ribeiro da Silva', 'Fernanda Neris Barroso', 'Wanessa Aparecida Nunes de Matos']
        cores = [(250, 250, 250), (250, 250, 250), (250, 250, 250)]
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

        sheet.spreadsheet.batch_update({
            "requests": [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 26,  # Coluna 27 (AA)
                            "endIndex": 30     # Coluna 31 (AE), mas o range exclui a última
                        },
                        "properties": {"hiddenByUser": True},
                        "fields": "hiddenByUser"
                    }
                }
            ]
        })

        corrigir_formulas(sheet)

def numero_para_letra_coluna(n):
    resultado = ""
    while n > 0:
        n -= 1
        resultado = chr(65 + (n % 26)) + resultado
        n //= 26
    return resultado

def corrigir_formulas(sheet):
    all_cells = sheet.get_values()  # Obtém todas as células da planilha

    updates = []
    for row_idx, row in enumerate(all_cells, start=1):
        for col_idx, cell in enumerate(row, start=1):
            if isinstance(cell, str) and cell.startswith("'="):  # Verifica se a célula contém um erro
                updates.append({
                    "range": f"{numero_para_letra_coluna(col_idx)}{row_idx}",
                    "values": [[cell.replace("'=", "=")]]
                })
    
    if updates:
        sheet.batch_update(updates)
        
def abrir_planilha(app):
    arquivos = arquivos_selecionados(app)
    if not verificar_selecao(app):
        return
    for arquivo_id in arquivos:
        url = f"https://docs.google.com/spreadsheets/d/{arquivo_id}"
        os.system(f'start "" "{url}"')

def baixar_planilha(app):
    arquivos = arquivos_selecionados(app)
    if not verificar_selecao(app):
        return
    for arquivo_id in arquivos:
        url = f"https://docs.google.com/spreadsheets/d/{arquivo_id}/export?format=xlsx"
        os.system(f'start "" "{url}"')