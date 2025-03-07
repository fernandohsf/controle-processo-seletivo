import os, threading, time
from utils import arquivos_selecionados, criar_barra_progresso, verificar_selecao

def obter_cabecalhos(sheet):
    cabecalhos = sheet.row_values(1)
    return {nome_coluna: idx+1 for idx, nome_coluna in enumerate(cabecalhos)}

def inserir_colunas(sheet, colunas_info, cabecalhos):
    for coluna, nome, formula_template in colunas_info:
        if nome in cabecalhos:
            continue

        sheet.insert_cols([[None]], coluna)
        sheet.update_cell(1, coluna, nome)
        # Atualiza o dicionário para refletir a nova coluna
        cabecalhos[nome] = coluna

def atualizar_formulas(sheet, colunas_info):
    for coluna, nome, formula_template in colunas_info:
        if not formula_template.strip():
            continue

        coluna_letra = numero_para_letra_coluna(coluna)
        cell_updates = []

        for row in range(2, sheet.row_count + 1):
            formula = formula_template.format(row=row)
            cell_updates.append({
                'range': f'{coluna_letra}{row}',
                'values': [[formula]]
            })

        sheet.spreadsheet.values_batch_update({
            "valueInputOption": "USER_ENTERED",
            "data": cell_updates
        })

def executar_ajuste_planilha(app):    
    arquivos = arquivos_selecionados(app)

    app.progress["maximum"] = len(arquivos)+1
    
    colunas_info = [
        (1,  'Número da Inscrição', "=SE(B{row}=\"\"; \"\"; \"PSSI\"&AD{row})"),
        (20, 'CPF (Conferência)', r"""=TEXTO(H:H;"000\.000\.000\-00")"""),
        (21, 'Quantidade de Inscrições neste PSSI', "=CONT.SE(T:T;T{row})"),
        (22, 'Notas', ""),
        (23, 'Analista RH', ""),
        (24, 'Status da inscrição', ""),
        (25, 'Motivo de Inaptidão/Eliminação', ""),
        (26, 'Controle de e-mail', ""),
        (27, 'Classificação do Candidato (Coordenador)', ""),
        (28, 'Ranking', "=SEERRO(ORDEM(B$1:B$1000;B$1:B$1000;VERDADEIRO);\"\")"),
        (29, 'ID Inscrição', "=SE(B{row}=\"\";\"\";(ESQUERDA(E{row};3)&(DIREITA(E{row};4)&AB{row})))"),
        (30, 'Formatação ID', "=TEXTO(AC{row};\"000000000\")")
    ]

    for i, arquivo_id in enumerate(arquivos):
        app.progress["value"] = i + 1  # Atualiza o progresso
        app.canvas.update_idletasks()
        planilha = app.cliente_gspread.open_by_key(arquivo_id)
        sheet = planilha.sheet1

        cabecalhos = {nome: idx+1 for idx, nome in enumerate(sheet.row_values(1))}

        inserir_colunas(sheet, colunas_info, cabecalhos)

        atualizar_formulas(sheet, colunas_info)
        
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

        # OCULTAR COLUNAS
        sheet.spreadsheet.batch_update({
            "requests": [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheet.id,
                            "dimension": "COLUMNS",
                            "startIndex": 27,  # Coluna 27 (AA)
                            "endIndex": 30     # Coluna 31 (AE), mas o range exclui a última
                        },
                        "properties": {"hiddenByUser": True},
                        "fields": "hiddenByUser"
                    }
                }
            ]
        })
    app.progress["value"] = len(arquivos) +1 # Completa a barra
    app.label_barra_progresso.config(text="Concluído!")
    time.sleep(1)
    app.montar_lista_de_arquivos()
    app.botao_enviar_emails.config(state="normal")
    app.texto_pesquisa.config(state="normal")

def numero_para_letra_coluna(n):
    resultado = ""
    while n > 0:
        n -= 1
        resultado = chr(65 + (n % 26)) + resultado
        n //= 26
    return resultado
        
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

def ajustar_planilha(app):
    if not verificar_selecao(app):
        return
    
    app.botao_enviar_emails.config(state="disabled")
    app.texto_pesquisa.config(state="disabled")
    criar_barra_progresso(app)

    thread = threading.Thread(target=executar_ajuste_planilha, args=(app,))
    thread.start()