def mensagens(status, linha, col_map):
    if status == "apto":
        mensagem = f"""
            <p>Prezado(a) <b>{linha[col_map[7]]}</b>,</p>

            <p>Informamos que sua inscrição foi recebida e você está participando do Pocesso Seletivo Simplificado {linha[col_map[5]]} referente à vaga de {linha[col_map[6]]}.<br>
            Salientamos que o número de inscrição associado ao seu processo é <b>{linha[col_map[1]]}</b>.<br>
            Em conformidade com as diretrizes estabelecidas pela Lei Geral de Proteção de Dados (LGPD), todos os dados dos candidatos são mantidos sob sigilo, sendo cada participante identificado exclusivamente por seu número de inscrição.<br>
            Acompanhe o andamento e o resultado da seleção pelo nosso portal www.fundacaofapec.org.br.</p>

            <p>Atenciosamente,<br>
            {linha[col_map[23]]}<br>
            <b>Equipe FAPEC - Processos Seletivos Interno FAPEC</b><br>
            Departamento Pessoal.</p>
            """
        return mensagem
    
    elif status == "inapto":
        mensagem = f"""
            <p>Prezado(a) <b>{linha[col_map[7]]}</b>,</p>

            <p>Informamos que sua inscrição para o Processo Seletivo Simplificado {linha[col_map[5]]}, referente à vaga de {linha[col_map[6]]} não foi aprovada por não atendimento e consequente descumprimento do(s) seguinte(s) itens obrigatórios do Edital:<br>
            {linha[col_map[25]]}.<br>
            Você poderá concorrer em novas oportunidades, desde que cumpra todas as condições estabelecidas pelo instrumento de concorrência para a vaga pleiteada, num novo processo.<br>
            Para este processo de seleção, seu CPF já foi analisado e, portanto, novas inscrições serão automaticamente invalidadas, conforme o Edital.<br>
            Acompanhe novas oportunidades de trabalho que são regularmente publicadas em nosso portal www.fundacaofapec.org.br.</p>

            <p>Atenciosamente,<br>
            {linha[col_map[23]]}<br>
            <b>Equipe FAPEC - Processos Seletivos Interno FAPEC</b><br>
            Departamento Pessoal.</p>
            """
        return mensagem
    
    elif status == "classificado para vaga":
        mensagem = f"""
            <p>Prezado(a) <b>{linha[col_map[7]]}</b>,</p>

            <p>Informamos que você foi selecionado para a vaga de {linha[col_map[6]]}, no Processo Seletivo Simplificado {linha[col_map[5]]}.<br>
            Para efetivar sua contratação, você tem o prazo improrrogável de 2 (dois) dias úteis para encaminhar ao e-mail dp@fapec.org a seguinte documentação, em formato PDF, com digitalização clara e legível:<br>
            a.Certidão de Nascimento ou Casamento;<br>
            b.RG;<br>
            c.CPF;<br>
            d.CNH;<br>
            e.Carteira de Trabalho Digital;<br>
            f.PIS/PASEP/NIT;<br>
            g.Título de Eleitor;<br>
            h.Carteira de reservista;<br>
            i.Comprovante de residência;<br>
            j.Comprovante de conta bancária;<br>
            k.Exame admissional;<br>
            m.Currículo atualizado;<br>
            n.Comprovante dos requisitos de escolaridade exigido no Edital (diploma graduação, especialização, MBA, mestrado, cursos técnicos).</p>

            <p>No campo "Assunto" do e-mail, favor escrever: {linha[col_map[5]]} - {linha[col_map[1]]}. Confira com atenção a relação de documentos no ato de envio.<br>
            Caso não seja encaminhado dentro do prazo, com dados ausentes ou ilegíveis, você será considerado desistente, portanto, Desclassificado da seleção e o próximo aprovado na lista será convocado.<br>
            Reiteramos a importância de observar o prazo e o envio da documentação.</p>

            <p>Atenciosamente,<br>
            {linha[col_map[23]]}<br>
            <b>Equipe FAPEC - Processos Seletivos Interno FAPEC</b><br>
            Departamento Pessoal.</p> 
            """
        return mensagem
    
    elif status == "cadastro de reserva":
        mensagem = f"""
            <p>Prezado(a) <b>{linha[col_map[7]]}</b>,</p>

            <p>Informamos que a sua classificação ficou fora do número de vagas abertas no momento para o Processo Seletivo Simplificado {linha[col_map[5]]}, para a vaga de {linha[col_map[6]]}, assim durante o prazo de 3 meses você ficará no cadastro reserva.<br>
            Isso significa que se houver vacância neste processo seletivo, dentro do período de 3 meses, você poderá vir a ser chamado. Ressaltamos que, em respeito à LGPD e visando preservar os dados dos concorrentes desta seleção, cada candidato será identificado por meio de seu número de inscrição.<br>
            O Edital de classificação encontra-se disponibilizado no seguinte link: <a href="https://fundacaofapec.org.br/concurso/id/202/?processo-seletivo-interno.html">Processos Seletivos Fundação FAPEC</a>.</p>

            <p>Atenciosamente,<br>
            {linha[col_map[23]]}<br>
            <b>Equipe FAPEC - Processos Seletivos Interno FAPEC</b><br>
            Departamento Pessoal.</p>
            """
        return mensagem
    else:
        return