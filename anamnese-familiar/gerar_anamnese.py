#!/usr/bin/env python3
"""Gera a Anamnese Social da Família (DOCX) da ONG Santidade Integral em Cristo.

Uso:  python3 gerar_anamnese.py
Saída: Anamnese_Social_Familiar_Santidade_Integral.docx (mesma pasta)
PDF:   soffice --headless --convert-to pdf <docx>
"""
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = Path(__file__).resolve().parent
LOGO = HERE.parent / "assets" / "logo.png"
OUT = HERE / "Anamnese_Social_Familiar_Santidade_Integral.docx"

GREEN = RGBColor(0x2D, 0x51, 0x30)
GREEN_HEX = "2D5130"
LIGHT_HEX = "E8EFE8"
GREY = RGBColor(0x59, 0x59, 0x59)
FONT = "Calibri"
BOX = "☐"  # ☐


# ----------------------------------------------------------------- helpers
def shade(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def set_cell_margins(table, top=40, bottom=40, left=80, right=80):
    tbl_pr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tbl_pr.append(mar)


def borders(table, color="A6B8A6", size=4):
    tbl_pr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        b.append(el)
    tbl_pr.append(b)


def para_bottom_border(paragraph, color=GREEN_HEX, size=12):
    p_pr = paragraph._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    bdr.append(bottom)
    p_pr.append(bdr)


def run(p, text, bold=False, italic=False, size=9.5, color=None, font=FONT):
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    r.font.name = font
    r._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    if color is not None:
        r.font.color.rgb = color
    return r


def spacing(p, before=0, after=2, line=1.0):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


def section_title(doc, number, title):
    p = doc.add_paragraph()
    spacing(p, before=8, after=3)
    p.paragraph_format.keep_with_next = True
    para_bottom_border(p)
    run(p, f"{number}  ", bold=True, size=11, color=GREEN)
    run(p, title.upper(), bold=True, size=11, color=GREEN)
    return p


LINE_CAP = 104  # capacidade aproximada da linha em caracteres (Calibri 9.5pt, 17.4cm)


def field_line(doc, *items, size=9.5):
    """items: ("Rótulo", peso) → rótulo + linha de preenchimento proporcional; ou texto puro."""
    p = doc.add_paragraph()
    spacing(p, after=3)
    labels = [it for it in items if isinstance(it, tuple)]
    fixed = sum(len(it[0]) + 2 for it in labels) + sum(len(it) for it in items if not isinstance(it, tuple))
    fixed += 3 * (len(items) - 1)
    free = max(LINE_CAP - fixed, 6 * len(labels))
    total_w = sum(it[1] for it in labels) or 1
    for i, it in enumerate(items):
        if isinstance(it, tuple):
            label, w = it
            run(p, f"{label}: ", bold=True, size=size)
            run(p, "_" * max(4, int(free * w / total_w)), size=size, color=GREY)
        else:
            run(p, it, size=size)
        if i < len(items) - 1:
            run(p, "   ", size=size)
    return p


def check_line(doc, label, options, size=9, note=None):
    p = doc.add_paragraph()
    spacing(p, after=3)
    run(p, f"{label}: ", bold=True, size=size)
    for o in options:
        run(p, f"{BOX} {o}  ", size=size)
    if note:
        run(p, note, size=size, color=GREY)
    return p


def table(doc, headers, rows, widths_cm, header_fill=GREEN_HEX, size=8.5, blank_rows=0):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    borders(t)
    set_cell_margins(t)
    for i, (h, w) in enumerate(zip(headers, widths_cm)):
        c = t.rows[0].cells[i]
        c.width = Cm(w)
        shade(c, header_fill)
        p = c.paragraphs[0]
        spacing(p, after=0)
        run(p, h, bold=True, size=size, color=RGBColor(0xFF, 0xFF, 0xFF))
    all_rows = list(rows) + [[""] * len(headers)] * blank_rows
    for r in all_rows:
        cells = t.add_row().cells
        for i, (val, w) in enumerate(zip(r, widths_cm)):
            cells[i].width = Cm(w)
            p = cells[i].paragraphs[0]
            spacing(p, after=0)
            if isinstance(val, tuple):  # (text, bold)
                run(p, val[0], bold=val[1], size=size)
            else:
                run(p, val, size=size)
    # altura mínima para linhas em branco
    for row in t.rows[1:]:
        tr_pr = row._tr.get_or_add_trPr()
        h = OxmlElement("w:trHeight")
        h.set(qn("w:val"), "300")
        h.set(qn("w:hRule"), "atLeast")
        tr_pr.append(h)
    fix_widths(t, widths_cm)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def note(doc, text, size=8.5):
    p = doc.add_paragraph()
    spacing(p, after=3)
    run(p, text, italic=True, size=size, color=GREY)
    return p


def lines_box(doc, n_lines=3, label=None):
    if label:
        p = doc.add_paragraph()
        spacing(p, after=2)
        run(p, label, bold=True, size=9.5)
    for _ in range(n_lines):
        p = doc.add_paragraph()
        spacing(p, before=9, after=0, line=1.0)
        para_bottom_border(p, color="9A9A9A", size=4)
        run(p, " ", size=9)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def fix_widths(t, widths_cm):
    """LibreOffice só respeita larguras se o grid e todas as células forem definidos."""
    tbl_pr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tbl_pr.append(lay)
    for col, w in zip(t.columns, widths_cm):
        col.width = Cm(w)
    for row in t.rows:
        for c, w in zip(row.cells, widths_cm):
            c.width = Cm(w)


def cant_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    tr_pr.append(el)


# ----------------------------------------------------------------- documento
def build():
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.8)
    sec.top_margin = Cm(3.9)
    sec.bottom_margin = Cm(1.6)
    sec.header_distance = Cm(0.7)
    sec.footer_distance = Cm(0.6)

    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(9.5)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    # ---- papel timbrado: logo no canto superior esquerdo + nome da ONG
    header = sec.header
    header.is_linked_to_previous = False
    ht = header.add_table(rows=1, cols=2, width=Cm(17.4))
    ht.autofit = False
    fix_widths(ht, [3.1, 14.3])
    tr_pr = ht.rows[0]._tr.get_or_add_trPr()
    h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), str(int(2.6 * 567)))
    h.set(qn("w:hRule"), "atLeast")
    tr_pr.append(h)
    hp = ht.rows[0].cells[0].paragraphs[0]
    spacing(hp, after=0)
    hp.add_run().add_picture(str(LOGO), height=Cm(2.4))
    c1 = ht.rows[0].cells[1]
    p = c1.paragraphs[0]
    spacing(p, before=4, after=0)
    run(p, "SANTIDADE INTEGRAL EM CRISTO", bold=True, size=11, color=GREEN)
    p = c1.add_paragraph()
    spacing(p, after=0)
    run(p, "Ação Social  •  Assistência e Promoção Familiar", size=8.5, color=GREY)
    p = c1.add_paragraph()
    spacing(p, after=0)
    run(p, "Prontuário nº: ______________    Unidade / Frente de trabalho: ______________________", size=8.5, color=GREY)
    hl = header.add_paragraph()
    spacing(hl, after=0)
    para_bottom_border(hl, size=8)

    # ---- rodapé
    footer = sec.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    spacing(fp, after=0)
    run(fp, "Documento de uso interno — informações protegidas pela LGPD (Lei 13.709/2018). "
            "Anamnese Social Familiar  •  v1.0  •  Página ", size=7.5, color=GREY)
    # campo PAGE
    r = fp.add_run()
    r.font.size = Pt(7.5)
    r.font.color.rgb = GREY
    for tag, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if tag:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tag)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = txt
        r._r.append(el)

    # ---- título
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    spacing(p, after=0)
    run(p, "ANAMNESE SOCIAL DA FAMÍLIA", bold=True, size=16, color=GREEN)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    spacing(p, after=0)
    run(p, "Cadastro, Diagnóstico Socioeconômico e Plano de Desenvolvimento Familiar", size=10, color=GREY)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    spacing(p, before=4, after=8)
    run(p, "“Eu sou o pão da vida; aquele que vem a mim não terá fome, "
           "e quem crê em mim nunca terá sede.”  ", italic=True, size=10, color=GREEN)
    run(p, "João 6:35", bold=True, size=10, color=GREEN)

    # ================================================== 1. ATENDIMENTO
    section_title(doc, "1", "Dados do atendimento")
    field_line(doc, ("Data da anamnese", 1), ("Voluntário(a) responsável", 3), ("Contato", 1.3))
    check_line(doc, "Forma de acesso", ["Busca ativa", "Demanda espontânea", "Encaminhamento (CRAS/UBS/Igreja): ______________"])
    check_line(doc, "Assistência em curso", ["Cesta básica", "Medicamentos", "Ambos", "Nenhuma (novo cadastro)"])
    field_line(doc, ("Desde (mês/ano)", 1), ("Periodicidade da entrega", 2))
    check_line(doc, "Tipo de atendimento", ["Inclusão (1º cadastro)", "Atualização cadastral", "Reavaliação do plano"])

    # ================================================== 2. RESPONSÁVEL FAMILIAR
    section_title(doc, "2", "Identificação do(a) Responsável Familiar (RF)")
    field_line(doc, ("Nome completo", 1))
    field_line(doc, ("Data de nascimento", 1), ("CPF", 1.3), ("RG / Órgão emissor", 1.4), ("NIS (CadÚnico)", 1.1))
    check_line(doc, "Sexo", ["Feminino", "Masculino"], note="   ")
    p = doc.paragraphs[-1]
    run(p, "Estado civil: ", bold=True)
    for o in ("Solteiro(a)", "Casado(a)/União estável", "Separado(a)", "Viúvo(a)"):
        run(p, f"{BOX} {o}  ")
    check_line(doc, "Raça/cor (autodeclarada)", ["Branca", "Preta", "Parda", "Amarela", "Indígena"])
    check_line(doc, "Escolaridade", ["Não alfabetizado(a)", "Fund. incompleto", "Fund. completo",
                                     "Médio incompleto", "Médio completo", "Superior"])
    field_line(doc, ("Telefone / WhatsApp", 1), ("Telefone de recado (nome)", 1.6))
    field_line(doc, ("Endereço", 3), ("Nº", 0.3), ("Bairro/Comunidade", 1.2))
    field_line(doc, ("Município/UF", 1.2), ("CEP", 0.7), ("Ponto de referência", 2))
    check_line(doc, "Zona", ["Urbana", "Rural"], note="      ")
    p = doc.paragraphs[-1]
    run(p, "Tempo de residência no local: ", bold=True, size=9)
    run(p, "____________", color=GREY, size=9)
    run(p, "      CRAS de referência: ", bold=True, size=9)
    run(p, "______________________", color=GREY, size=9)

    # ================================================== 3. COMPOSIÇÃO FAMILIAR
    section_title(doc, "3", "Composição familiar (todos os moradores do domicílio)")
    table(doc,
          ["Nome", "Parentesco c/ RF", "Idade", "Estuda?", "Escolaridade", "Ocupação / fonte de renda",
           "Renda mensal (R$)", "Condição de saúde / deficiência"],
          [], [4.0, 2.2, 1.1, 1.3, 2.0, 3.0, 1.7, 2.1], blank_rows=7)
    field_line(doc, ("Total no domicílio", 1), ("Crianças 0–6", 1), ("7–17 anos", 1),
               ("Idosos 60+", 1), ("Gestantes/nutrizes", 1), ("PcD", 1))
    check_line(doc, "Criança/adolescente fora da escola", ["Não", "Sim — quem/motivo: ________________________"])

    # ================================================== 4. HABITAÇÃO
    section_title(doc, "4", "Condições de habitação e acesso a serviços")
    check_line(doc, "Situação do imóvel", ["Próprio quitado", "Próprio financiado", "Alugado (R$ ______)", "Cedido", "Ocupação", "Outro"])
    check_line(doc, "Tipo / material", ["Alvenaria", "Madeira", "Taipa/pau a pique", "Material reaproveitado", "Misto"])
    field_line(doc, ("Cômodos", 1), ("Dormitórios", 1), ("Pessoas por dormitório", 1),
               f"Banheiro:  {BOX} Interno  {BOX} Externo  {BOX} Não possui")
    check_line(doc, "Água", ["Rede pública", "Poço/nascente", "Cisterna", "Carro-pipa", "Sem acesso regular"])
    check_line(doc, "Esgoto", ["Rede", "Fossa séptica", "Fossa rudimentar", "Céu aberto"], note="    ")
    p = doc.paragraphs[-1]
    run(p, "Energia: ", bold=True)
    for o in ("Regular", "Irregular", "Sem energia"):
        run(p, f"{BOX} {o}  ")
    check_line(doc, "Coleta de lixo", ["Regular", "Queima/enterra", "Descarte irregular"], note="      ")
    p = doc.paragraphs[-1]
    run(p, "Área de risco (enchente/deslizamento): ", bold=True, size=9)
    run(p, f"{BOX} Não  {BOX} Sim", size=9)
    check_line(doc, "Bens e equipamentos", ["Fogão", "Geladeira", "Botijão de gás", "Máq. de lavar", "Celular c/ internet",
                                           "TV", "Veículo/moto"])

    # ================================================== 5. SAÚDE
    section_title(doc, "5", "Saúde e uso de medicamentos")
    check_line(doc, "Família cadastrada na UBS/ESF", ["Sim", "Não"], note="   ")
    p = doc.paragraphs[-1]
    run(p, "Agente Comunitário de Saúde: ", bold=True)
    run(p, "______________________________", color=GREY)
    check_line(doc, "Condições presentes no domicílio",
               ["Hipertensão", "Diabetes", "Doença respiratória", "Transtorno mental", "Deficiência",
                "Gestação", "Acamado(a)", "Uso abusivo de álcool/drogas", "Outra: ________"])
    table(doc,
          ["Medicamento (nome e dosagem)", "Quem utiliza", "Posologia", "Receita médica?", "Forma de obtenção *", "Custo mensal (R$)"],
          [], [5.0, 3.0, 2.4, 2.0, 3.0, 2.0], blank_rows=5)
    note(doc, "* Forma de obtenção: SUS / Farmácia Popular / Compra própria / Doação da ONG / Sem acesso (interrupção do tratamento).")

    # ================================================== 6. SEGURANÇA ALIMENTAR
    section_title(doc, "6", "Segurança alimentar e nutricional")
    note(doc, "Perguntas de triagem adaptadas da Escala Brasileira de Insegurança Alimentar (EBIA) — referência: últimos 3 meses.")
    table(doc,
          ["Nos últimos 3 meses…", "Sim", "Não"],
          [["Os moradores tiveram preocupação de que os alimentos acabassem antes de poderem comprar mais?", BOX, BOX],
           ["Os alimentos acabaram antes que houvesse dinheiro para comprar mais?", BOX, BOX],
           ["Algum adulto diminuiu a quantidade de alimentos ou pulou refeições por falta de dinheiro?", BOX, BOX],
           ["Algum adulto deixou de fazer uma refeição o dia inteiro por falta de dinheiro?", BOX, BOX],
           ["Alguma criança/adolescente comeu menos ou deixou de comer por falta de dinheiro?", BOX, BOX]],
          [14.9, 1.25, 1.25])
    field_line(doc, ("Refeições por dia", 0.4), ("Principal fonte dos alimentos", 2.2))
    check_line(doc, "Produção para consumo próprio (horta, criação de animais)", ["Não", "Sim — o quê: ______________________"])
    note(doc, "Classificação (soma de “Sim”): 0 = segurança alimentar • 1–2 = IA leve • 3–4 = IA moderada • 5 = IA grave.")

    # ================================================== 7. TRABALHO E RENDA
    section_title(doc, "7", "Trabalho, renda e orçamento familiar")
    check_line(doc, "Ocupação do RF", ["Desempregado", "Informal/bicos", "Autônomo", "CLT", "Produtor rural",
                                       "Aposentado/pensionista", "Do lar"])
    field_line(doc, ("Tempo sem trabalho formal", 1), ("Última ocupação", 2.5))
    t = doc.add_table(rows=1, cols=2)
    t.autofit = False
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fix_widths(t, [8.7, 8.7])
    cant_split(t.rows[0])
    left, right = t.rows[0].cells
    # Receitas
    p = left.paragraphs[0]
    spacing(p, after=1)
    run(p, "RECEITAS MENSAIS", bold=True, size=9, color=GREEN)
    receitas = [("Renda do trabalho (soma dos membros)", ""), ("Bolsa Família", ""), ("BPC/LOAS", ""),
                ("Aposentadoria / pensão", ""), ("Seguro-desemprego / auxílio", ""), ("Pensão alimentícia", ""),
                ("Ajuda de terceiros / igreja", ""), ("Outras", ""), (("TOTAL DE RECEITAS", True), "")]
    despesas = [("Aluguel / prestação", ""), ("Água e energia", ""), ("Gás", ""), ("Alimentação", ""),
                ("Medicamentos / saúde", ""), ("Transporte", ""), ("Educação / material escolar", ""),
                ("Dívidas / parcelas", ""), (("TOTAL DE DESPESAS", True), "")]
    for cell, items, title in ((left, receitas, None), (right, despesas, "DESPESAS MENSAIS")):
        if title:
            p = cell.paragraphs[0]
            spacing(p, after=1)
            run(p, title, bold=True, size=9, color=GREEN)
        inner = cell.add_table(rows=0, cols=2)
        inner.autofit = False
        borders(inner)
        inner_widths = [6.0, 2.2]
        set_cell_margins(inner, top=20, bottom=20)
        for label, _ in items:
            r = inner.add_row().cells
            r[0].width, r[1].width = Cm(6.2), Cm(2.3)
            bold = isinstance(label, tuple)
            txt = label[0] if bold else label
            if bold:
                shade(r[0], LIGHT_HEX)
                shade(r[1], LIGHT_HEX)
            p = r[0].paragraphs[0]
            spacing(p, after=0)
            run(p, txt, bold=bold, size=8.5)
            p = r[1].paragraphs[0]
            spacing(p, after=0)
            run(p, "R$", size=8.5, color=GREY)
        fix_widths(inner, inner_widths)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    field_line(doc, ("Renda per capita R$", 1), ("Saldo mensal R$", 1), ("Dívidas acumuladas R$", 1))
    note(doc, "Renda per capita = total de receitas ÷ nº de pessoas no domicílio.  Saldo mensal = receitas − despesas.")
    check_line(doc, "Faixa de renda per capita", ["Até R$ 218 (extrema pobreza)", "R$ 218 a R$ 706 (até ½ salário mínimo)", "Acima de ½ SM"])
    check_line(doc, "Cadastro Único", ["Não cadastrado", "Cadastrado e atualizado", "Desatualizado (> 24 meses)"])

    # ================================================== 8. VÍNCULOS E REDE
    section_title(doc, "8", "Vínculos familiares, comunitários e rede de proteção")
    check_line(doc, "Acompanhamento em serviços", ["CRAS/PAIF", "CREAS/PAEFI", "CAPS", "Conselho Tutelar", "Nenhum"])
    check_line(doc, "Participação comunitária", ["Igreja/comunidade de fé", "Associação de moradores", "Grupo produtivo/cooperativa", "Nenhuma"])
    check_line(doc, "Situações de risco", ["Violência doméstica", "Trabalho infantil", "Negligência", "Isolamento",
                                           "Conflito familiar", "Nenhuma"])
    lines_box(doc, 3, "Rede de apoio (parentes, vizinhos, lideranças) e observações sobre a dinâmica familiar:")

    # ================================================== 9. POTENCIALIDADES
    section_title(doc, "9", "Potencialidades, ativos e aspirações da família")
    note(doc, "Levantamento dos ativos humanos, sociais, físicos e produtivos — base para o projeto de geração de renda.")
    table(doc,
          ["Membro", "Habilidades / ofícios", "Cursos / qualificação", "Experiência profissional (anos / área)", "Disponibilidade"],
          [], [3.0, 5.4, 3.0, 3.4, 2.6], blank_rows=4)
    note(doc, "Exemplos de ofícios: costura, pedreiro, eletricista, cozinha/salgados, cabeleireiro, manicure, roça, criação de animais, motorista, cuidador.")
    check_line(doc, "Ativos produtivos", ["Quintal (____ m²)", "Terra p/ cultivo (___ ha)", "Animais: ______",
                                          "Ferramentas", "Máq. costura", "Veículo"])
    check_line(doc, "Documentação e crédito", ["CNH", "CNPJ/MEI ativo", "DAP/CAF", "Conta bancária",
                                               "Nome negativado", "Sem documentos"])
    check_line(doc, "Interesse da família", ["Emprego formal", "Empreendimento próprio", "Produção agropecuária",
                                             "Qualificação", "Retomar estudos"])
    lines_box(doc, 3, "O que a família sonha para os próximos 12 meses? Quais obstáculos ela identifica?")

    # ================================================== 10. DIAGNÓSTICO
    section_title(doc, "10", "Diagnóstico social (síntese do voluntário)")
    check_line(doc, "Classificação da vulnerabilidade", ["Emergencial (risco imediato)", "Transitória (perda de renda)",
                                                          "Crônica (intergeracional)"])
    check_line(doc, "Prioridade de intervenção", ["Alta", "Média", "Baixa"])
    check_line(doc, "Encaminhamentos", ["CadÚnico", "CRAS", "UBS", "Escola", "Defensoria", "SINE/emprego", "EMATER/Sind. rural"])
    lines_box(doc, 6, "Principais vulnerabilidades, riscos e potencialidades observadas:")

    # ================================================== 11. PLANO
    doc.add_page_break()
    section_title(doc, "11", "Plano de Desenvolvimento Familiar (pactuado com a família)")
    check_line(doc, "Eixo do projeto", ["Trabalho/empregabilidade", "Empreendedorismo", "Agropecuário", "Qualificação", "Misto"])
    field_line(doc, ("Descrição do projeto de geração de renda", 1))
    lines_box(doc, 1)
    field_line(doc, ("Investimento inicial (R$)", 1), f"Fonte:  {BOX} ONG  {BOX} Família  {BOX} Crédito  {BOX} Parceiro",
               ("Renda mensal projetada (R$)", 1))
    table(doc,
          ["Meta / objetivo", "Ações concretas", "Responsável", "Prazo", "Indicador de sucesso"],
          [], [4.2, 5.4, 2.4, 1.8, 3.6], blank_rows=5)
    field_line(doc, ("Redução gradual da cesta a partir de", 1), ("Meta de autonomia (data)", 1),
               ("Reavaliação a cada", 0.4), "meses")

    # ================================================== 12. ACOMPANHAMENTO
    section_title(doc, "12", "Registro de acompanhamento (evolução)")
    table(doc,
          ["Data", "Tipo *", "Situação encontrada / evolução das metas", "Próxima ação", "Voluntário"],
          [], [1.7, 2.2, 8.1, 3.2, 2.2], blank_rows=6)
    note(doc, "* Tipo: V = visita domiciliar • C = contato telefônico • E = entrega (cesta/medicamento) • A = atendimento na sede.")

    # ================================================== 13. CONSENTIMENTO
    section_title(doc, "13", "Termo de consentimento e assinaturas")
    p = doc.add_paragraph()
    spacing(p, after=4)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run(p, "Declaro que as informações prestadas são verdadeiras e autorizo a ONG Santidade Integral em Cristo a registrar, "
           "guardar e utilizar estes dados exclusivamente para fins de assistência social, acompanhamento familiar e "
           "encaminhamento à rede de proteção, nos termos da Lei Geral de Proteção de Dados (Lei nº 13.709/2018). "
           "Estou ciente de que posso solicitar acesso, correção ou exclusão dos meus dados a qualquer momento.", size=9)
    field_line(doc, ("Local e data", 1))
    t = doc.add_table(rows=1, cols=2)
    t.autofit = False
    fix_widths(t, [8.7, 8.7])
    for i, label in enumerate(("Responsável Familiar", "Voluntário(a) responsável pela anamnese")):
        c = t.rows[0].cells[i]
        c.width = Cm(8.7)
        p = c.paragraphs[0]
        spacing(p, before=18, after=0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run(p, "_" * 42, color=GREY)
        p = c.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        spacing(p, after=0)
        run(p, label, size=8.5, color=GREY)

    doc.save(OUT)
    print(f"OK -> {OUT}")


if __name__ == "__main__":
    build()
