#!/usr/bin/env python3
"""Gera os documentos de prestação de contas da ONG Santidade Integral em Cristo.

  1. Prestacao_de_Contas_Mensal.docx/.pdf  — relatório mensal + balancete (impressão, 2 páginas)
  2. Politica_de_Boas_Praticas.docx/.pdf   — política de governança, transparência e fidelidade
  3. Declaracao_de_Conformidade.docx/.pdf  — atestado periódico das boas práticas aplicadas

Uso:  python3 gerar_prestacao_contas.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from timbrado import *  # noqa: E402,F401

HERE = Path(__file__).resolve().parent
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MONEY = "R$"


# ----------------------------------------------------------------- helpers locais
def money_table(doc, headers, rows, widths_cm, size=8.5, total_label=None, total_cols=None):
    """Tabela com linhas rotuladas e colunas de valor pré-preenchidas com 'R$'.

    rows: lista de (rótulo, [células...]) — células "R$" viram cinza; None vira vazio.
    """
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    borders(t)
    set_cell_margins(t, top=25, bottom=25)
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        shade(c, GREEN_HEX)
        p = c.paragraphs[0]
        spacing(p, after=0)
        run(p, h, bold=True, size=size, color=WHITE)
    for label, cells in rows:
        r = t.add_row().cells
        for c in r:
            spacing(c.paragraphs[0], after=0)
        p = r[0].paragraphs[0]
        bold = isinstance(label, tuple)
        run(p, label[0] if bold else label, bold=bold, size=size)
        if bold:
            for c in r:
                shade(c, LIGHT_HEX)
        for i, val in enumerate(cells, start=1):
            p = r[i].paragraphs[0]
            spacing(p, after=0)
            if val == MONEY:
                run(p, MONEY, size=size, color=GREY)
            elif val:
                run(p, val, size=size)
    for row in t.rows[1:]:
        tr_pr = row._tr.get_or_add_trPr()
        h = OxmlElement("w:trHeight")
        h.set(qn("w:val"), "245")
        h.set(qn("w:hRule"), "atLeast")
        tr_pr.append(h)
    fix_widths(t, widths_cm)
    gap = doc.add_paragraph()
    spacing(gap, after=0)
    gap.runs or run(gap, " ", size=4)
    return t


def check_table(doc, items, widths_cm=(13.0, 1.45, 1.45, 1.5), size=8.5, headers=("Prática", "Sim", "Não", "N/A")):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    borders(t)
    set_cell_margins(t, top=20, bottom=20)
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        shade(c, GREEN_HEX)
        p = c.paragraphs[0]
        spacing(p, after=0)
        if i > 0:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run(p, h, bold=True, size=size, color=WHITE)
    for it in items:
        r = t.add_row().cells
        for c in r:
            spacing(c.paragraphs[0], after=0)
        if isinstance(it, tuple):  # (grupo,) → linha de seção
            for c in r:
                shade(c, LIGHT_HEX)
            p = r[0].paragraphs[0]
            spacing(p, after=0)
            run(p, it[0], bold=True, size=size, color=GREEN)
            continue
        p = r[0].paragraphs[0]
        spacing(p, after=0)
        run(p, it, size=size)
        for i in range(1, len(headers)):
            p = r[i].paragraphs[0]
            spacing(p, after=0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run(p, BOX, size=size + 1)
    fix_widths(t, list(widths_cm))
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return t


def small_title(doc, number, title, before=4):
    p = section_title(doc, number, title)
    spacing(p, before=before, after=1)
    for r in p.runs:
        r.font.size = Pt(10)
    return p


def verse_box(doc, verse, ref, size=9):
    t = doc.add_table(rows=1, cols=1)
    t.autofit = False
    fix_widths(t, [17.4])
    c = t.rows[0].cells[0]
    shade(c, LIGHT_HEX)
    set_cell_margins(t, top=60, bottom=60, left=140, right=140)
    p = c.paragraphs[0]
    spacing(p, after=0)
    run(p, f"“{verse}”  ", italic=True, size=size, color=GREEN)
    run(p, f"— {ref}", bold=True, size=size, color=GREEN)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def finish(doc, name):
    out = HERE / f"{name}.docx"
    doc.save(out)
    print(f"OK -> {out}")
    print(f"OK -> {to_pdf(out)}")
    return out


# ================================================================= 1. RELATÓRIO MENSAL
def build_mensal():
    doc = new_document(
        header_meta="Competência (mês/ano): ______________    Relatório nº: ________    Elaborado em: ____/____/______",
        footer_label="Prestação de Contas Mensal  •  Modelo conforme ITG 2002 (R1) / CFC  •  v1.0",
        top_margin_cm=3.7, bottom_margin_cm=1.3,
    )
    title_block(doc, "PRESTAÇÃO DE CONTAS MENSAL",
                "Relatório de Execução Físico-Financeira e Balancete do Período",
                "Procuramos o que é honesto, não só diante do Senhor, mas também diante dos homens.",
                "2 Coríntios 8:21", after=2)

    # 1 — identificação
    small_title(doc, "1", "Identificação", before=2)
    field_line(doc, ("Período de competência", 1), ("Tesoureiro(a) / resp. financeiro", 1.6), ("CPF", 0.8), size=9)
    field_line(doc, ("Conta bancária exclusiva (banco / ag. / conta)", 1.6), ("Saldo bancário em ____/____ (R$)", 0.9), size=9)

    # 2 — indicadores sociais
    small_title(doc, "2", "Execução física — atendimento do mês")
    money_table(doc,
        ["Indicador", "Meta", "Realizado", "Indicador", "Meta", "Realizado"],
        [("Famílias ativas no cadastro", ["", "", "Cestas básicas entregues (un.)", "", ""]),
         ("Famílias atendidas no mês", ["", "", "Famílias com medicamentos", "", ""]),
         ("Novos cadastros (anamnese)", ["", "", "Itens de medicamento entregues", "", ""]),
         ("Visitas domiciliares realizadas", ["", "", "Famílias em plano de geração de renda", "", ""]),
         ("Pessoas beneficiadas (total)", ["", "", "Desligamentos por autonomia", "", ""])],
        [4.6, 1.3, 1.6, 5.6, 1.3, 1.6], size=8)

    # 3 — receitas
    small_title(doc, "3", "Receitas do mês (ingressos financeiros)")
    money_table(doc,
        ["Fonte de recurso", "Restrição *", "Previsto", "Realizado", "Comprovante (nº / tipo)"],
        [("Doações de pessoas físicas", ["", MONEY, MONEY, ""]),
         ("Doações de pessoas jurídicas", ["", MONEY, MONEY, ""]),
         ("Ofertas e campanhas de igrejas parceiras", ["", MONEY, MONEY, ""]),
         ("Eventos, bazares e ações de arrecadação", ["", MONEY, MONEY, ""]),
         ("Parcerias públicas (termo de fomento / colaboração)", ["", MONEY, MONEY, ""]),
         ("Rendimentos de aplicação financeira", ["", MONEY, MONEY, ""]),
         ("Outras receitas", ["", MONEY, MONEY, ""]),
         (("TOTAL DE RECEITAS",), [None, MONEY, MONEY, ""])],
        [7.2, 1.8, 2.0, 2.0, 4.4], size=8)
    note(doc, "* Restrição: SR = sem restrição (livre)  •  CR = com restrição (vinculada a projeto/doador — ITG 2002 R1).", size=7)

    # 4 — despesas
    small_title(doc, "4", "Despesas do mês (por natureza)")
    money_table(doc,
        ["Natureza da despesa", "Qtd.", "Previsto", "Realizado", "Nº NF / recibo"],
        [(("ATIVIDADE-FIM (finalística)",), [None, None, None, ""]),
         ("Cestas básicas — gêneros alimentícios", ["", MONEY, MONEY, ""]),
         ("Medicamentos e insumos de saúde", ["", MONEY, MONEY, ""]),
         ("Projetos de geração de renda (insumos, ferramentas, cursos)", ["", MONEY, MONEY, ""]),
         ("Transporte e logística das entregas", ["", MONEY, MONEY, ""]),
         (("ATIVIDADE-MEIO (administrativa)",), [None, None, None, ""]),
         ("Material de expediente, impressão e comunicação", ["", MONEY, MONEY, ""]),
         ("Tarifas bancárias e taxas", ["", MONEY, MONEY, ""]),
         ("Outras despesas", ["", MONEY, MONEY, ""]),
         (("TOTAL DE DESPESAS",), [None, MONEY, MONEY, ""])],
        [8.4, 1.2, 2.0, 2.0, 3.8], size=8)

    # 5 — doações em espécie
    small_title(doc, "5", "Doações em bens (não financeiras) — movimentação de estoque")
    money_table(doc,
        ["Item", "Unid.", "Estoque inicial", "Recebido", "Distribuído", "Estoque final", "Valor estimado **"],
        [("Cestas básicas montadas", ["un.", "", "", "", "", MONEY]),
         ("Alimentos avulsos", ["kg", "", "", "", "", MONEY]),
         ("Medicamentos", ["itens", "", "", "", "", MONEY]),
         ("Outros (vestuário, higiene, materiais)", ["", "", "", "", "", MONEY])],
        [5.4, 1.2, 2.0, 1.8, 2.0, 2.0, 3.0], size=8)
    note(doc, "** Bens doados valorizados pelo valor justo (preço médio de mercado) — ITG 2002 (R1), item 27. "
              "Medicamentos: registrar lote e validade na ficha de estoque.", size=7)

    # ---------------------------------------------------------- página 2
    doc.add_page_break()
    small_title(doc, "6", "Balancete do período (regime de caixa)", before=0)
    t = doc.add_table(rows=1, cols=2)
    t.autofit = False
    fix_widths(t, [8.7, 8.7])
    left, right = t.rows[0].cells
    for cell, title in ((left, "APURAÇÃO DO RESULTADO"), (right, "CONCILIAÇÃO BANCÁRIA E DE CAIXA")):
        p = cell.paragraphs[0]
        spacing(p, after=1)
        run(p, title, bold=True, size=9, color=GREEN)
    def inner(cell, items):
        it = cell.add_table(rows=0, cols=2)
        it.autofit = False
        borders(it)
        set_cell_margins(it, top=25, bottom=25)
        for label, strong in items:
            r = it.add_row().cells
            if strong:
                shade(r[0], LIGHT_HEX); shade(r[1], LIGHT_HEX)
            p = r[0].paragraphs[0]; spacing(p, after=0); run(p, label, bold=strong, size=8.5)
            p = r[1].paragraphs[0]; spacing(p, after=0); run(p, MONEY, size=8.5, color=GREY)
        fix_widths(it, [6.1, 2.3])
    inner(left, [("Saldo inicial (caixa + banco)", False), ("(+) Total de receitas (item 3)", False),
                 ("(−) Total de despesas (item 4)", False), ("(=) Superávit / déficit do mês", True),
                 ("(=) SALDO FINAL", True)])
    inner(right, [("Saldo do extrato bancário no último dia", False), ("(+) Caixa físico (fundo fixo)", False),
                  ("(−) Pagamentos emitidos não compensados", False), ("(+) Depósitos em trânsito", False),
                  ("(=) SALDO CONCILIADO (= saldo final)", True)])
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    money_table(doc,
        ["Recursos com restrição (projetos / termos / doador específico)", "Saldo anterior", "Entradas", "Saídas", "Saldo final"],
        [("", [MONEY, MONEY, MONEY, MONEY]), ("", [MONEY, MONEY, MONEY, MONEY]),
         ("Fundo de reserva / contingência", [MONEY, MONEY, MONEY, MONEY])],
        [8.2, 2.3, 2.3, 2.3, 2.3], size=8)

    small_title(doc, "7", "Verificação de conformidade (boas práticas aplicadas no período)")
    check_table(doc, [
        "Todos os lançamentos possuem nota fiscal ou recibo em nome da ONG, arquivados em ordem cronológica",
        "Extrato bancário do período conciliado e anexado a este relatório",
        "Pagamentos autorizados por duas pessoas (princípio dos quatro olhos: tesoureiro + presidente)",
        "Compras acima do limite definido na política realizadas com no mínimo três cotações",
        "Nenhum dirigente, conselheiro ou voluntário remunerado, beneficiado ou com conflito de interesse não declarado",
        "Recursos com restrição aplicados exclusivamente no plano de trabalho pactuado",
        "Lista nominal de famílias atendidas no mês arquivada com sigilo (LGPD) e disponível ao Conselho Fiscal",
        "Estoque de cestas e medicamentos conferido fisicamente; medicamentos entregues somente com prescrição",
        "Resultado do mês comunicado aos doadores / igreja parceira / assembleia",
    ], size=8)

    small_title(doc, "8", "Ocorrências, ressalvas e providências")
    lines_box(doc, 5)

    small_title(doc, "9", "Pareceres e assinaturas")
    check_line(doc, "Conselho Fiscal", ["Parecer favorável", "Favorável com ressalvas", "Desfavorável (justificar no item 8)"], size=8.5)
    signature_row(doc, ["Tesoureiro(a) — elaboração", "Presidente — aprovação", "Conselho Fiscal — parecer"], before=16)
    p = doc.add_paragraph()
    spacing(p, before=6, after=0)
    run(p, "Local e data: ", bold=True, size=8.5)
    run(p, "_" * 60, size=8.5, color=GREY)
    note(doc, "Anexar: extrato bancário, comprovantes numerados, lista de beneficiários (acesso restrito) e, "
              "se houver recurso público, o relatório exigido pelo termo (Lei 13.019/2014).", size=7.5)
    finish(doc, "Prestacao_de_Contas_Mensal")


# ================================================================= 2. POLÍTICA DE BOAS PRÁTICAS
def build_politica():
    doc = new_document(
        header_meta="Documento normativo interno    •    Aprovado em: ____/____/______    •    Revisão: ______",
        footer_label="Política de Boas Práticas de Governança e Prestação de Contas  •  v1.0",
    )
    title_block(doc, "POLÍTICA DE BOAS PRÁTICAS",
                "Governança, Transparência e Prestação de Contas — Honra, Dignidade e Fidelidade",
                "Quem é fiel no pouco, também é fiel no muito.", "Lucas 16:10")

    section_title(doc, "1", "Finalidade e alcance")
    body(doc, "Esta política estabelece os princípios e as práticas que regem a captação, a guarda, a aplicação e a "
              "prestação de contas dos recursos — financeiros e em bens — confiados à Santidade Integral em Cristo "
              "para a assistência e a promoção de famílias em situação de vulnerabilidade. Aplica-se a dirigentes, "
              "conselheiros, tesoureiros, voluntários e parceiros que, em qualquer medida, manuseiem recursos ou "
              "atendam beneficiários em nome da organização.")
    body(doc, "Os recursos administrados não pertencem à diretoria: são bens de Deus, entregues por doadores em "
              "confiança, destinados aos pobres. Por isso a organização se compromete a tratá-los com mordomia, "
              "transparência e rigor superiores ao que a lei exige.")

    section_title(doc, "2", "Fundamento bíblico — os três pilares")
    bullet(doc, "Agir corretamente diante de Deus e dos homens. A administração deve ser inatacável, prestando contas "
                "de forma aberta para que ninguém possa censurar o serviço prestado (2 Coríntios 8:20-21; Romanos 12:17; "
                "Provérbios 3:9).", lead="HONRA —")
    bullet(doc, "O pobre é imagem de Deus e irmão, não objeto de caridade. Servi-lo é servir ao próprio Cristo, com "
                "respeito, discrição e sem humilhação (Mateus 25:40; Tiago 1:27; Provérbios 14:31; Deuteronômio 15:7-8).",
           lead="DIGNIDADE —")
    bullet(doc, "O que se requer dos despenseiros é que sejam achados fiéis. Balanças justas, registros exatos e "
                "cuidado com o pouco para que seja confiado o muito (1 Coríntios 4:2; Lucas 16:10-12; Provérbios 11:1; "
                "Atos 6:1-4).", lead="FIDELIDADE —")

    section_title(doc, "3", "Princípios de governança")
    bullet(doc, "os recursos são aplicados exclusivamente na finalidade estatutária, com economicidade e sem desvio.", lead="Mordomia:")
    bullet(doc, "toda movimentação é registrada, comprovada e divulgada a quem confiou os recursos.", lead="Transparência:")
    bullet(doc, "quem autoriza não é quem paga; quem paga não é quem confere. Nenhuma pessoa controla sozinha o ciclo completo.",
           lead="Segregação de funções:")
    bullet(doc, "decisões sobre beneficiários, compras e parcerias seguem critérios objetivos, nunca preferência pessoal, "
                "parentesco ou troca de favores.", lead="Impessoalidade:")
    bullet(doc, "observância do Estatuto, da ITG 2002 (R1) do Conselho Federal de Contabilidade (entidades sem fins "
                "lucrativos), da Lei 13.019/2014 (MROSC) quando houver recurso público, da LGPD e da legislação "
                "tributária aplicável.", lead="Legalidade:")

    section_title(doc, "4", "Práticas de gestão financeira")
    bullet(doc, "Conta bancária exclusiva em nome da organização. É vedado transitar recursos da ONG em contas pessoais.")
    bullet(doc, "Princípio dos quatro olhos: todo pagamento é autorizado pelo tesoureiro e por um segundo dirigente.")
    bullet(doc, "Toda despesa possui nota fiscal ou recibo em nome da ONG, numerado e arquivado por, no mínimo, cinco anos.")
    bullet(doc, "Compras acima de R$ ________ exigem três cotações registradas; a escolha é justificada por preço e qualidade.")
    bullet(doc, "Caixa físico limitado a R$ ________ (fundo fixo), com reposição mediante prestação de contas.")
    bullet(doc, "Recursos com restrição (vinculados a projeto ou doador) são controlados separadamente e aplicados "
                "somente no plano de trabalho pactuado.")
    bullet(doc, "Dirigentes e conselheiros não recebem remuneração, distribuição de resultados, empréstimos ou qualquer vantagem.")
    bullet(doc, "Manutenção de fundo de reserva equivalente a, pelo menos, um mês de despesas finalísticas.")
    bullet(doc, "Doações em bens são registradas em ficha de estoque com entrada, saída, saldo e valor justo estimado.")

    section_title(doc, "5", "Práticas no atendimento às famílias")
    bullet(doc, "Inclusão mediante Anamnese Social da Família e critérios objetivos de vulnerabilidade, documentados e revisados periodicamente.")
    bullet(doc, "Sigilo absoluto sobre dados e situação das famílias (LGPD). Nenhum beneficiário é exposto em fotos, "
                "listas públicas ou redes sociais sem consentimento livre e por escrito.")
    bullet(doc, "A entrega é feita com respeito, em local e horário dignos, sem filas vexatórias, discursos "
                "constrangedores ou condicionamento da ajuda à adesão religiosa.")
    bullet(doc, "Cada família tem um Plano de Desenvolvimento Familiar com metas de autonomia; a assistência é ponte, "
                "não destino, e a saída por autonomia é celebrada.")
    bullet(doc, "Medicamentos são entregues somente mediante prescrição médica válida, com conferência de lote e validade, "
                "e orientação sobre uso; medicamentos vencidos são descartados conforme norma sanitária.")

    section_title(doc, "6", "Transparência e prestação de contas")
    bullet(doc, "elaboração do relatório de execução físico-financeira com balancete conciliado, assinado pelo "
                "tesoureiro, aprovado pelo presidente e com parecer do Conselho Fiscal.", lead="Mensal:")
    bullet(doc, "demonstrações contábeis (balanço patrimonial, demonstração do resultado do período, notas "
                "explicativas) apresentadas à assembleia geral, com parecer do Conselho Fiscal, até 30 de abril.", lead="Anual:")
    bullet(doc, "resumo das receitas, despesas e resultados sociais enviado aos doadores e igrejas parceiras e "
                "disponível a qualquer interessado que o solicite.", lead="Aos doadores:")
    bullet(doc, "manutenção de canal direto (telefone / e-mail) para dúvidas, sugestões e denúncias, com garantia "
                "de não retaliação.", lead="Canal aberto:")

    section_title(doc, "7", "Conflito de interesses e conduta")
    bullet(doc, "Dirigentes, conselheiros e voluntários assinam anualmente a Declaração de Conformidade e Boas Práticas "
                "e informam qualquer vínculo com fornecedores, parceiros ou beneficiários.")
    bullet(doc, "Quem tiver interesse pessoal em uma decisão se abstém de participar dela, e o fato é registrado em ata.")
    bullet(doc, "É vedado aceitar presentes, comissões ou vantagens de fornecedores e parceiros em razão da função.")
    bullet(doc, "O nome, a logomarca e os recursos da organização não são usados para fins particulares, políticos ou comerciais.")

    section_title(doc, "8", "Vigência, revisão e aprovação")
    body(doc, "Esta política entra em vigor na data de sua aprovação pela diretoria, é revisada ao menos a cada dois anos "
              "e integra os anexos obrigatórios da prestação de contas anual. Casos omissos são resolvidos pela diretoria, "
              "ouvido o Conselho Fiscal, sempre à luz dos princípios do item 2.")
    signature_row(doc, ["Presidente", "Tesoureiro(a)", "Conselho Fiscal"], before=20)
    p = doc.add_paragraph()
    spacing(p, before=8, after=0)
    run(p, "Local e data: ", bold=True, size=9)
    run(p, "_" * 60, size=9, color=GREY)
    finish(doc, "Politica_de_Boas_Praticas")


# ================================================================= 3. DECLARAÇÃO DE CONFORMIDADE
def build_declaracao():
    doc = new_document(
        header_meta="Anexo à prestação de contas    •    Período de referência: ____________________",
        footer_label="Declaração de Conformidade e Boas Práticas  •  v1.0",
    )
    title_block(doc, "DECLARAÇÃO DE CONFORMIDADE E BOAS PRÁTICAS",
                "Atestado de aplicação da Política de Boas Práticas — Honra, Dignidade e Fidelidade",
                "O que se requer dos despenseiros é que cada um seja encontrado fiel.", "1 Coríntios 4:2", after=6)

    body(doc, "Nós, abaixo assinados, dirigentes e responsáveis pela gestão da Santidade Integral em Cristo, "
              "declaramos, sob compromisso de honra diante de Deus e dos homens, que no período de referência "
              "indicado acima a organização observou as práticas a seguir, conforme a Política de Boas Práticas "
              "de Governança e Prestação de Contas em vigor:")

    check_table(doc, [
        ("HONRA — agir corretamente diante de Deus e dos homens",),
        "As contas do período foram apresentadas com balancete conciliado, comprovantes arquivados e parecer do Conselho Fiscal",
        "Os resultados financeiros e sociais foram comunicados aos doadores, igrejas parceiras e assembleia",
        "Nenhuma informação foi omitida, alterada ou apresentada de forma a induzir doadores ou parceiros a erro",
        "O nome e a logomarca da organização foram usados apenas para os fins estatutários",
        ("DIGNIDADE — o pobre como imagem de Deus e irmão",),
        "As famílias foram incluídas por critérios objetivos (anamnese social), sem favorecimento pessoal ou parentesco",
        "As entregas ocorreram com respeito e discrição, sem exposição, filas vexatórias ou condicionamento religioso",
        "Os dados das famílias foram mantidos em sigilo (LGPD); nenhuma imagem foi divulgada sem consentimento escrito",
        "Cada família acompanhada possui plano de desenvolvimento com metas de autonomia revisadas no período",
        "Medicamentos foram entregues somente com prescrição válida, com conferência de lote e validade",
        ("FIDELIDADE — despenseiros achados fiéis",),
        "Todos os recursos transitaram pela conta bancária exclusiva da organização; nada passou por contas pessoais",
        "Todo pagamento foi autorizado por duas pessoas e está lastreado em nota fiscal ou recibo em nome da ONG",
        "As compras acima do limite da política tiveram três cotações registradas",
        "Os recursos com restrição foram aplicados exclusivamente no plano de trabalho pactuado",
        "Nenhum dirigente, conselheiro ou voluntário foi remunerado, beneficiado ou teve conflito de interesse não declarado",
        "O estoque de cestas e medicamentos foi conferido fisicamente e confere com as fichas de controle",
    ], size=8.5)

    p = doc.add_paragraph()
    spacing(p, before=4, after=2)
    run(p, "Ressalvas, exceções e providências adotadas ", bold=True, size=9)
    run(p, "(obrigatório para cada item marcado “Não”):", size=9, color=GREY)
    lines_box(doc, 3)

    body(doc, "Declaramos ainda estar cientes de que a inexatidão desta declaração sujeita os signatários às "
              "responsabilidades estatutárias, civis e penais cabíveis, e de que o Conselho Fiscal e qualquer doador "
              "poderão verificar, a qualquer tempo, os documentos que a fundamentam.", size=9)
    signature_row(doc, ["Presidente", "Tesoureiro(a)", "Conselho Fiscal"], before=18)
    p = doc.add_paragraph()
    spacing(p, before=8, after=0)
    run(p, "Local e data: ", bold=True, size=9)
    run(p, "_" * 60, size=9, color=GREY)
    finish(doc, "Declaracao_de_Conformidade")


if __name__ == "__main__":
    build_mensal()
    build_politica()
    build_declaracao()
