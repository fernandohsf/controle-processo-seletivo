def mensagens(status, linha, col_map):
    if status == "apto":
        mensagem = f"""
Prezado(a) {linha[col_map[5]]},

Temos a satisfação de confirmar a homologação de sua candidatura, registrada sob o ID {linha[col_map[1]]}.

Em conformidade com a Lei Geral de Proteção de Dados (LGPD), informamos que, em todas as divulgações relacionadas a esse edital, seu nome será substituído pelo seu ID de inscrição. É fundamental que você guarde essa informação com segurança e não a compartilhe com ninguém.

Para garantir sua continuidade neste processo, lembre-se de acompanhar publicações estipuladas no edital, caixa do e-mail cadastrado e manter o telefone informado bom funcionamento.

Somos muito gratos por sua participação em nosso processo seletivo.

Atenciosamente.
Equipe FAPEC - Processo Seletivo
"""
        return mensagem
    elif status == "inapto":
        mensagem = f"""
Prezado(a) {linha[col_map[5]]},

Informamos que sua inscrição não foi homologada, pelo descumprimento do(s) seguinte(s) itens obrigatórios:

{linha[col_map[19]]}

Se você puder cumprir com os itens supracitados, anexando os documentos comprobatórios numa nova inscrição, ficaremos felizes em reavaliar sua solicitação.

Se o período de inscrições já foi encerrado, acompanhe novas oportunidades no nosso portal www.fundacaofapec.org.br e lembre-se de ler o edital com atenção.

Atenciosamente,
Equipe FAPEC - Processos Seletivos."""
        return mensagem
    else:
        return